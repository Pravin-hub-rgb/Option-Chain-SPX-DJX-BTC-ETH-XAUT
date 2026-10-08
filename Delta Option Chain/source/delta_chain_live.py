"""
Delta Exchange BTC Option Chain — standalone reference implementation.

Surgical companion to latency_harness.py: this one writes a live chain into a
normal desktop Excel workbook (.xlsx) via xlwings, driven by the WebSocket feed
that the Delta website itself uses.

    python delta_chain_live.py --book Chain.xlsx --minutes 10

Design notes that matter (all verified against the India endpoints):

  * Endpoint filter is `expiry_date` in DD-MM-YYYY. There is also an `expiry`
    param, but the API SILENTLY IGNORES it and returns every expiry with
    HTTP 200, so asserting `data["success"]` is not enough — we verify each
    returned symbol carries the expected DDMMYY suffix.
  * `underlying_asset_symbols` is plural. The singular form is ignored and
    returns every asset's contracts.
  * The "BTC price" on Delta's option-chain page is `spot_price` carried in the
    option tickers. It is NOT the perpetual's `close` (that is the perp's own
    last trade, ~30-45 USD away) and not `mark_price`.
  * Freshness is measured from WebSocket receive time and from actual
    spot_price movement, NOT from each option's server `timestamp` — illiquid
    strikes carry stale timestamps while the index is moving.
  * Server `timestamp` is in MICROSECONDS.
"""
import argparse
import asyncio
import json
import os
import statistics
import time
from datetime import datetime, timezone

import requests
import websockets

try:
    import xlwings as xw
except ImportError:
    raise SystemExit("pip install xlwings  (needs Microsoft Excel on Windows)")

REST_BASE = "https://api.india.delta.exchange"
WS_URL = "wss://socket.india.delta.exchange"
ASSET = "BTC"

# Dedicated index channel. Delta publishes the BTC index here roughly every
# 0.26 s, versus ~4.9 s when it rides along inside v2/ticker option frames.
# Subscribe to BOTH: index for the header, v2/ticker for the chain rows.
# NOTE: XAUT's index symbol is .DEXAUTUSD, not .DEXXAUTUSD — read these from
# GET /v2/indices, do not derive them from the underlying ticker name.
INDEX_SYMBOLS = {"BTC": ".DEXBTUSD", "ETH": ".DEETHUSD", "XAUT": ".DEXAUTUSD"}
STALE_AFTER_S = 3.0
WRITE_THROTTLE_S = 0.8

# sheet layout
ROW0 = 10          # first data row
HEADER_ROWS = {
    "A1": "Delta BTC Option Chain (live)",
    "A2": "Index (spot_price)",
    "B2": None,     # filled at runtime
    # Expiry is TEXT on purpose. Excel coerces a dd/mm/yyyy string into a
    # datetime, which then renders as a 1900 serial number.
    "D2": "Expiry (DD-MM-YYYY)",
    "E2": None,
    "G2": "Last update (local)",
    "H2": None,
    "J2": "Server time",
    "K2": None,
    "M2": "Age (s)",
    "N2": None,
    "O2": "Freshness",
    "P2": None,
    "R2": "Last index change",
    "S2": None,
    "T2": "Secs since change",
    "U2": None,
    "V2": "Index feed",
    "W2": None,
}
COLUMNS = [
    "Call Delta", "Call Bid Qty", "Call Bid", "Call Mark", "Call Ask",
    "Call Ask Qty", "Call OI",
    "Strike",
    "Put OI", "Put Bid Qty", "Put Bid", "Put Mark", "Put Ask",
    "Put Ask Qty", "Put Delta",
]


def iso_dmy(dmy):
    dd, mm, yyyy = dmy.split("-")
    return f"{yyyy}-{mm}-{dd}"


def dmy_suffix(dmy):
    """'08-10-2026' -> '081026'"""
    dd, mm, yyyy = dmy.split("-")
    return f"{dd}{mm}{yyyy[2:4]}"


def pick_expiry():
    p = {
        "contract_types": "call_options,put_options",
        "underlying_asset_symbols": ASSET,
    }
    prods = requests.get(
        f"{REST_BASE}/v2/products", params=p, timeout=20
    ).json().get("result") or []
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    dates = sorted({str(x.get("settlement_time") or "")[:10] for x in prods})
    future = [d for d in dates if d >= today]
    if not future:
        raise SystemExit("no live expiries")
    yyyy, mm, dd = future[0].split("-")
    return f"{dd}-{mm}-{yyyy}"


def fetch_chain(expiry_dmy):
    """REST snapshot, filtered + verified."""
    url = f"{REST_BASE}/v2/tickers"
    params = {
        "contract_types": "call_options,put_options",
        "underlying_asset_symbols": ASSET,
        "expiry_date": expiry_dmy,
    }
    t0 = time.time()
    r = requests.get(url, params=params, timeout=15)
    t1 = time.time()
    data = r.json()
    rows = data.get("result") or []

    # Guard: success is not enough. If the filter is ignored we get every
    # expiry with HTTP 200, so verify the suffix ourselves.
    want = dmy_suffix(expiry_dmy)
    rows = [x for x in rows if str(x.get("symbol", "")).endswith(want)]
    if not rows:
        return {}, {}

    sts = max((float(x.get("timestamp") or 0) for x in rows), default=0.0) or None
    spot = rows[0].get("spot_price")
    meta = {
        "rtt_ms": (t1 - t0) * 1000,
        "data_age_s": (t1 - sts / 1e6) if sts else None,
        "spot": spot,
    }
    return rows, meta


def build_chain(rows):
    """Group by strike -> (call, put) pairs, ascending."""
    by = {}
    for t in rows:
        try:
            strike = int(float(t.get("strike_price")))
        except (TypeError, ValueError):
            continue
        g = t.get("greeks") or {}
        q = t.get("quotes") or {}
        entry = {
            "mark": float(t.get("mark_price") or 0),
            "bid": float(q.get("best_bid") or 0),
            "ask": float(q.get("best_ask") or 0),
            "bid_qty": float(q.get("bid_size") or 0),
            "ask_qty": float(q.get("ask_size") or 0),
            "oi": float(t.get("oi_value_usd") or 0),
            "delta": float(g.get("delta") or 0),
        }
        by.setdefault(strike, {})["call" if t.get("contract_type") == "call_options" else "put"] = entry
    return by


class Writer:
    """Batched Excel writer. One range write per block, only on change."""

    def __init__(self, book_path, sheet_name="Chain"):
        self.path = book_path
        self.sheet_name = sheet_name
        self.app = None
        self.book = None
        self.sheet = None
        self._last_payload = None
        self._last_write = 0.0
        self.write_ms = []

    def save(self):
        if self.book is not None:
            self.book.save()

    def open(self):
        if not os.path.exists(self.path):
            xw.Book().save(self.path)
        # Open THE FILE, not app.books[0] — a freshly created xw.Book() lands
        # in the same Excel instance as a new unsaved workbook, so app.books[0]
        # can be that instead of our file and every write is silently lost.
        self.book = xw.Book(self.path)
        self.app = self.book.app
        self.sheet = self.book.sheets[0]
        self.sheet.name = self.sheet_name
        for addr, val in HEADER_ROWS.items():
            if val is not None:
                self.sheet.range(addr).value = val
        for i, name in enumerate(COLUMNS):
            self.sheet.range(ROW0 - 2, i + 1).value = name
        self.sheet.range(ROW0 - 2, 1).api.Font.Bold = True
        return self

    def write_header(self, spot, expiry_dmy, age_s, server_hms,
                     last_change=None, secs_since=None, feed="-"):
        stale = age_s is not None and age_s > STALE_AFTER_S
        # prefix with apostrophe to force TEXT so Excel cannot coerce these
        # into datetime/times and render them as serial numbers
        self.sheet.range("B2").value = f"'{spot}" if spot is not None else "'-"
        self.sheet.range("E2").value = f"'{expiry_dmy}"
        self.sheet.range("H2").value = f"'{datetime.now().strftime('%H:%M:%S')}"
        self.sheet.range("K2").value = f"'{server_hms}" if server_hms else "'-"
        self.sheet.range("N2").value = f"'{round(age_s, 2)}" if age_s is not None else "'-"
        self.sheet.range("P2").value = "'STALE" if stale else "'LIVE"
        self.sheet.range("P2").api.Font.Color = 0x0000FF if stale else 0x008000
        # Makes the publish cadence visible: the client can see the number only
        # moves when the exchange actually publishes a new index value.
        self.sheet.range("S2").value = (
            f"'{datetime.fromtimestamp(last_change).strftime('%H:%M:%S')}"
            if last_change else "'-"
        )
        self.sheet.range("U2").value = (
            f"'{round(secs_since, 2)}" if secs_since is not None else "'-"
        )
        self.sheet.range("W2").value = f"'{feed}"

    def write_chain(self, by_strike):
        strikes = sorted(by_strike)
        if not strikes:
            return
        rows = []
        for s in strikes:
            c = by_strike[s].get("call") or {}
            p = by_strike[s].get("put") or {}
            rows.append([
                c.get("delta", 0), c.get("bid_qty", 0), c.get("bid", 0),
                c.get("mark", 0), c.get("ask", 0), c.get("ask_qty", 0),
                c.get("oi", 0),
                s,
                p.get("oi", 0), p.get("bid_qty", 0), p.get("bid", 0),
                p.get("mark", 0), p.get("ask", 0), p.get("ask_qty", 0),
                p.get("delta", 0),
            ])
        payload = tuple(tuple(r) for r in rows)
        now = time.time()
        if payload == self._last_payload:
            return
        if now - self._last_write < WRITE_THROTTLE_S:
            return
        t0 = time.time()
        # ONE batched write for the whole block
        self.sheet.range(ROW0, 1).value = rows
        self._last_write = t0
        self._last_payload = payload
        self.write_ms.append((time.time() - t0) * 1000)


async def run(args):
    expiry = args.expiry if args.expiry != "auto" else pick_expiry()
    expiry_iso = iso_dmy(expiry)
    print(f"expiry: {expiry} ({expiry_iso})")

    rows, meta = fetch_chain(expiry)
    if not rows:
        raise SystemExit(f"no contracts for {expiry}")
    symbols = sorted({str(x["symbol"]) for x in rows})
    print(f"rest snapshot: {len(rows)} rows, spot={meta['spot']}, "
          f"rtt={meta['rtt_ms']:.0f}ms, age={meta['data_age_s']:.2f}s")

    w = Writer(args.book).open()
    by = build_chain(rows)
    w.write_chain(by)
    w.write_header(meta["spot"], expiry, meta["data_age_s"], None)
    print(f"excel ready: {args.book} (sheet {args.sheet})")

    latest = {}
    spot = {"v": None, "changed": 0.0, "server": None, "feed": "-"}
    ages = []
    change_gaps = []
    idx_changes = []
    last_idx_value = None
    last_idx_change = 0.0
    end = time.time() + args.minutes * 60

    async with websockets.connect(WS_URL) as ws:
        # Two channels: the dedicated index feed for the header (fast, ~0.26s)
        # and v2/ticker for the chain rows (index inside it updates ~4.9s).
        await ws.send(json.dumps({
            "type": "subscribe",
            "payload": {"channels": [
                {"name": "spot_price", "symbols": [INDEX_SYMBOLS[ASSET]]},
                {"name": "v2/ticker", "symbols": symbols},
            ]},
        }))
        print(f"subscribed: spot_price[{INDEX_SYMBOLS[ASSET]}] + v2/ticker[{len(symbols)}]")
        print("streaming...")
        while time.time() < end:
            try:
                msg = await asyncio.wait_for(ws.recv(), timeout=2.0)
            except asyncio.TimeoutError:
                # idle tick: refresh header so STALE / cadence stay live
                ssc = time.time() - last_idx_change if last_idx_change else None
                w.write_header(spot["v"], expiry, ssc, spot["server"],
                               last_idx_change, ssc, spot["feed"])
                continue
            try:
                d = json.loads(msg)
            except ValueError:
                continue
            recv = time.time()
            mtype = d.get("type")

            if mtype == "spot_price":
                # Dedicated index channel — primary header source.
                p = d.get("price")
                if p is not None:
                    if p != last_idx_value:
                        # gap between CONSECUTIVE VALUE changes = real cadence
                        if last_idx_change:
                            idx_changes.append(recv - last_idx_change)
                        last_idx_value = p
                        last_idx_change = recv
                    spot["v"] = float(p)
                    spot["feed"] = f"spot_price {INDEX_SYMBOLS[ASSET]}"
                ssc = recv - last_idx_change if last_idx_change else None
                w.write_header(spot["v"], expiry, ssc, spot["server"],
                               last_idx_change, ssc, spot["feed"])
                continue

            if mtype != "v2/ticker":
                continue

            sym = d.get("symbol")
            if sym:
                latest[sym] = d
            # Fallback only: if the index channel dies, option spot_price keeps
            # the header alive (slower, ~4.9s).
            if spot["feed"] == "-":
                sp = d.get("spot_price")
                if sp is not None and float(sp) != spot["v"]:
                    spot["v"] = float(sp)
                    spot["changed"] = recv
                    spot["feed"] = "v2/ticker (fallback)"
            sts = d.get("timestamp")
            if sts and spot["feed"] == "-":
                spot["server"] = datetime.fromtimestamp(
                    float(sts) / 1e6).strftime("%H:%M:%S")
            by = build_chain(list(latest.values()))
            w.write_chain(by)
            ssc = recv - spot["changed"] if spot["changed"] else None
            if ssc is not None and ssc < 1e9:
                ages.append(ssc)
            w.write_header(spot["v"], expiry, ssc, spot["server"],
                           spot["changed"], ssc, spot["feed"])

    if idx_changes:
        c = sorted(idx_changes)
        print(f"\nINDEX CHANGE CADENCE ({spot['feed']}):")
        print(f"  changes={len(c)}  median={c[len(c) // 2]:.3f}s  "
              f"p95={c[int(len(c) * 0.95)]:.3f}s  max={max(c):.3f}s")
    if ages:
        a = sorted(ages)
        print(f"\nheader age: samples={len(a)} avg={statistics.mean(a):.3f}s "
              f"p95={a[int(len(a) * 0.95)]:.3f}s max={max(a):.3f}s")
    if w.write_ms:
        print(f"excel write ms: avg={statistics.mean(w.write_ms):.1f} "
              f"max={max(w.write_ms):.1f} n={len(w.write_ms)}")
    w.save()
    print(f"saved: {w.path} / sheet {w.sheet_name}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--book", default="DeltaChain.xlsx")
    ap.add_argument("--sheet", default="Chain")
    ap.add_argument("--expiry", default="auto")
    ap.add_argument("--minutes", type=float, default=10.0)
    args = ap.parse_args()
    asyncio.run(run(args))


if __name__ == "__main__":
    main()