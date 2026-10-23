"""
Latency harness for the Delta BTC option chain.

Measures how fresh the index price is under REST polling vs WebSocket, and
compares the app's index (option-ticker spot_price) against an INDEPENDENT
WebSocket client at the same instant.

Usage:
    python latency_harness.py --minutes 3
    python latency_harness.py --minutes 10 --expiry auto --compare

Writes latency_log.csv with one row per sample.
"""
import argparse
import asyncio
import csv
import json
import os
import statistics
import sys
import time
from datetime import datetime, timezone

import requests
import websockets

REST_BASE = "https://api.india.delta.exchange"
WS_URL = "wss://socket.india.delta.exchange"
ASSET = "BTC"
OUT_CSV = "latency_log.csv"

COLUMNS = [
    "mode",              # rest | ws | ws_independent
    "local_time",        # ISO local time of the observation
    "local_epoch",       # float seconds, for precise maths
    "server_ts",         # server timestamp (from payload, microseconds)
    "server_ts_iso",     # converted from microseconds
    "spot_price",
    "rtt_ms",            # REST only: request round trip
    "recv_delay_ms",     # ws only: server_ts -> local receive
    "data_age_s",        # local receive -> server_ts
    "source",            # which endpoint produced it
]


def log(rows, path=OUT_CSV):
    exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        if not exists:
            w.writeheader()
        for r in rows:
            w.writerow(r)


def now_iso(epoch=None):
    return datetime.fromtimestamp(epoch if epoch is not None else time.time()).isoformat(
        timespec="milliseconds"
    )


def _server_ts(row):
    ts = row.get("timestamp")
    try:
        return float(ts)
    except (TypeError, ValueError):
        return None


def rest_sample(expiry_dmy):
    """One REST snapshot with latency measurement."""
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
    if not rows:
        return None
    # Guard: the filter can be silently ignored, so verify the suffix.
    # expiry_dmy is DD-MM-YYYY -> symbol suffix is '081026' (DDMMYY), not DDMMYYYY
    dd, mm, yy = expiry_dmy.split("-")
    suffix = f"{dd}{mm}{yy[2:4]}"
    filtered = [x for x in rows if str(x.get("symbol", "")).endswith(suffix)]
    if rows and not filtered:
        print(f"  [guard] {len(rows)} rows returned but none end with {suffix}; "
              f"expiry filter was ignored, discarding")
        return None
    rows = filtered
    if not rows:
        return None
    sts = max((_server_ts(x) or 0) for x in rows)
    sts = sts or None
    spot = rows[0].get("spot_price")
    return {
        "mode": "rest",
        "local_time": now_iso(t1),
        "local_epoch": f"{t1:.6f}",
        "server_ts": sts if sts else "",
        "server_ts_iso": now_iso(sts / 1e6) if sts else "",
        "spot_price": spot if spot is not None else "",
        "rtt_ms": f"{(t1 - t0) * 1000:.1f}",
        "recv_delay_ms": "",
        "data_age_s": f"{t1 - sts / 1e6:.3f}" if sts else "",
        "source": "rest/v2/tickers?expiry_date",
    }


async def ws_samples(duration_s, symbols, tag, on_tick=None):
    """Subscribe and record every push for duration_s seconds."""
    out = []
    connected = False
    async with websockets.connect(WS_URL) as ws:
        await ws.send(json.dumps({
            "type": "subscribe",
            "payload": {"channels": [{"name": "v2/ticker", "symbols": symbols}]},
        }))
        connected = True
        deadline = time.time() + duration_s
        while time.time() < deadline:
            try:
                msg = await asyncio.wait_for(ws.recv(), timeout=max(0.5, deadline - time.time()))
            except asyncio.TimeoutError:
                break
            recv = time.time()
            try:
                d = json.loads(msg)
            except ValueError:
                continue
            if d.get("type") != "v2/ticker":
                continue
            spot = d.get("spot_price")
            if spot is None:
                continue
            sts = _server_ts(d)
            row = {
                "mode": tag,
                "local_time": now_iso(recv),
                "local_epoch": f"{recv:.6f}",
                "server_ts": sts if sts else "",
                "server_ts_iso": now_iso(sts / 1e6) if sts else "",
                "spot_price": spot,
                "rtt_ms": "",
                "recv_delay_ms": f"{(recv - sts / 1e6) * 1000:.1f}" if sts else "",
                "data_age_s": f"{recv - sts / 1e6:.3f}" if sts else "",
                "source": "ws/v2/ticker " + str(d.get("symbol")),
            }
            out.append(row)
            if on_tick:
                on_tick(row)
    return out, connected


def pick_expiry(arg):
    if arg and arg != "auto":
        return arg
    url = f"{REST_BASE}/v2/products"
    p = {"contract_types": "call_options,put_options", "underlying_asset_symbols": ASSET}
    prods = requests.get(url, params=p, timeout=20).json().get("result") or []
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    dates = sorted({str(x.get("settlement_time") or "")[:10] for x in prods})
    future = [d for d in dates if d >= today]
    if not future:
        raise SystemExit("no live expiries found")
    # settlement_time is ISO (2026-10-08); the tickers endpoint wants DD-MM-YYYY.
    yyyy, mm, dd = future[0].split("-")
    return f"{dd}-{mm}-{yyyy}"


def _compare_streams(a_rows, b_rows, window=1.0):
    """Pair spot_price observations from two concurrent WS clients that are
    within `window` seconds of each other, and return the signed differences."""
    a = [(float(r["local_epoch"]), float(r["spot_price"]))
         for r in a_rows if r.get("spot_price") not in ("", None)]
    b = [(float(r["local_epoch"]), float(r["spot_price"]))
         for r in b_rows if r.get("spot_price") not in ("", None)]
    if not a or not b:
        return []
    a.sort()
    b.sort()
    dif = []
    j = 0
    for ta, sa in a:
        # advance b to the nearest sample at or after ta
        while j < len(b) and b[j][0] < ta - window:
            j += 1
        best = None
        for k in range(max(0, j - 3), min(len(b), j + 4)):
            tb, sb = b[k]
            if abs(tb - ta) <= window and (best is None or abs(tb - ta) < abs(best[0] - ta)):
                best = (tb, sb)
        if best is not None:
            dif.append(sb - sa)
    return dif


def summary(rows, label):
    ages = [float(r["data_age_s"]) for r in rows if r.get("data_age_s") not in ("", None)]
    rtts = [float(r["rtt_ms"]) for r in rows if r.get("rtt_ms") not in ("", None)]
    print(f"\n--- {label} ---")
    print(f"  samples        : {len(rows)}")
    if ages:
        ages_sorted = sorted(ages)
        p95 = ages_sorted[int(len(ages_sorted) * 0.95)] if len(ages_sorted) > 1 else ages_sorted[0]
        print(f"  data_age_s avg : {statistics.mean(ages):.3f}")
        print(f"  data_age_s min : {min(ages):.3f}")
        print(f"  data_age_s p95 : {p95:.3f}")
        print(f"  data_age_s max : {max(ages):.3f}")
        under2 = sum(1 for a in ages if a < 2.0) / len(ages) * 100
        print(f"  age < 2s       : {under2:.1f}%")
        under1 = sum(1 for a in ages if a < 1.0) / len(ages) * 100
        print(f"  age < 1s       : {under1:.1f}%")
    if rtts:
        print(f"  rtt_ms avg     : {statistics.mean(rtts):.1f}")
    return ages


async def run(args):
    expiry = pick_expiry(args.expiry)
    print(f"expiry under test: {expiry}")
    symbols = []
    url = f"{REST_BASE}/v2/tickers"
    p = {
        "contract_types": "call_options,put_options",
        "underlying_asset_symbols": ASSET,
        "expiry_date": expiry,
    }
    tick = requests.get(url, params=p, timeout=20).json().get("result") or []
    symbols = sorted({str(x["symbol"]) for x in tick if x.get("symbol")})
    print(f"subscribing to {len(symbols)} option symbols")

    if os.path.exists(args.out):
        os.remove(args.out)

    # Phase 1: REST polling
    rest_rows = []
    t_end = time.time() + args.rest_seconds
    while time.time() < t_end:
        try:
            s = rest_sample(expiry)
            if s:
                rest_rows.append(s)
                log([s], args.out)
        except Exception as e:
            print("rest error:", e)
        time.sleep(args.rest_interval)
    summary(rest_rows, "REST (expiry_date snapshot)")

    # Phase 2: WebSocket
    ws_rows = []
    t_end = time.time() + args.ws_seconds
    while time.time() < t_end:
        rows, _ = await ws_samples(
            min(20, max(1, t_end - time.time())), symbols, "ws"
        )
        ws_rows.extend(rows)
        for r in rows:
            log([r], args.out)
    summary(ws_rows, "WebSocket (v2/ticker)")

    # Phase 3: independent WS client running CONCURRENTLY with the app stream,
    # so both observe the same index at the same instants and can be compared.
    if args.compare:
        ind_rows = []
        shared = []

        async def pump(tag, dur, syms, bucket):
            t_end = time.time() + dur
            while time.time() < t_end:
                rows, _ = await ws_samples(
                    min(10, max(1, t_end - time.time())), syms, tag
                )
                bucket.extend(rows)
                for r in rows:
                    log([r], args.out)

        await asyncio.gather(
            pump("ws", args.compare_seconds, symbols, ws_rows),
            pump("ws_independent", args.compare_seconds, symbols[:12], ind_rows),
        )
        summary(ws_rows, "WebSocket (app stream, concurrent)")
        summary(ind_rows, "WebSocket (independent client, concurrent)")

        dif = _compare_streams(ws_rows, ind_rows)
        if dif:
            print(f"\n--- app ws vs independent ws, paired within 1s ---")
            print(f"  paired samples : {len(dif)}")
            print(f"  diff avg       : {statistics.mean(dif):+.4f}")
            print(f"  diff abs max   : {max(abs(x) for x in dif):.4f}")
            print(f"  diff abs mean  : {statistics.mean([abs(x) for x in dif]):.4f}")
            consistent = all(x > 0.05 for x in dif) or all(x < -0.05 for x in dif)
            print(f"  consistent sign: {consistent}")
            print(f"  VERDICT        : {'FIELD-MAPPING BUG' if consistent else 'OK - no fixed offset'}")
        else:
            print("\nstill no paired samples; streams did not overlap")

    print(f"\nlog written to {args.out}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--minutes", type=float, default=3.0)
    ap.add_argument("--expiry", default="auto")
    ap.add_argument("--rest-seconds", type=float, default=None)
    ap.add_argument("--ws-seconds", type=float, default=None)
    ap.add_argument("--compare-seconds", type=float, default=None)
    ap.add_argument("--rest-interval", type=float, default=2.0)
    ap.add_argument("--compare", action="store_true")
    ap.add_argument("--out", default=OUT_CSV)
    args = ap.parse_args()
    total = args.minutes * 60
    if args.rest_seconds is None:
        args.rest_seconds = total * 0.25
    if args.ws_seconds is None:
        args.ws_seconds = total * 0.45
    if args.compare_seconds is None:
        args.compare_seconds = total * 0.30
    print(f"plan: rest {args.rest_seconds:.0f}s, ws {args.ws_seconds:.0f}s, "
          f"independent {args.compare_seconds:.0f}s")
    asyncio.run(run(args))


if __name__ == "__main__":
    main()