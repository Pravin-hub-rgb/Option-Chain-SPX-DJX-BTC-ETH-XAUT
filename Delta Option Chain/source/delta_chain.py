"""Delta exchange feed, chain construction and terminal display.

REST products/tickers fetches, the option-chain build (build_option_chain),
the WebSocket plumbing (ws_listener/index_listener/process_frame), the OI and
spot fallback pollers, expiry bookkeeping and the shared formatting helpers
(fmt_oi/safe_float/n_blocks) live here so main.py only wires the Excel render
and tool lifecycle. State stays in delta_state.py.
"""


import asyncio
import json
import threading
import time
from datetime import datetime, timezone

import pandas as pd
import requests
import websockets

from delta_diagnostics import (
    _flat,
    feed_log,
)
from delta_spot import (
    note_ws_recv,
    set_spot,
    spot_age,
)
from delta_state import (
    EXPIRY_REFRESH_SECONDS,
    HTTP_TIMEOUT,
    INDEX_ASSET,
    INDEX_CHANNEL,
    INDEX_SYMBOLS,
    INDEX_WS_URL,
    NET_FAIL_LIMIT,
    OI_POLL_SECONDS,
    REST_BASE,
    SPOT_POLL_SECONDS,
    SPOT_TICKER_SYMBOLS,
    STALE_AFTER_S,
    STALE_FLAG_SECONDS,
    WS_URL,
    _net,
    _parity_spot_cache,
    _PUT_INTRINSIC_WARNED,
    _PUT_INTRINSIC_WARN_INTERVAL,
    active_expiries,
    active_expiries_dmy,
    expiry_symbols,
    feed_state,
    live_data,
    spot_cmp,
    spot_source,
    symbol_expiry,
)


def net_note(ok):
    """Har network round-trip ke baad call karo."""
    if ok:
        if _net["fails"] >= NET_FAIL_LIMIT:
            print("Connection restored", flush=True)
        _net["fails"] = 0
        _net["since"] = None
        return
    _net["fails"] += 1
    if _net["fails"] == NET_FAIL_LIMIT:
        _net["since"] = time.time()
        print("Connection lost - retrying in background", flush=True)


_console_warn_last = {}


def _quiet_console_error(site, detail):
    """Rate-limit one site's console error line; keep every detail in feed.log."""
    now = time.monotonic()
    if now - _console_warn_last.get(site, 0.0) > 120.0:
        print(f"{site}: delta.exchange is slow right now - retrying "
              "(full detail in feed.log)", flush=True)
        _console_warn_last[site] = now
    feed_log("network_retry", site=site, detail=str(detail)[:160])


def n_blocks(asset):
    """Every sheet shows two expiry blocks — near expiry and the next one.

    XAUT already only had 2 blocks (it rarely lists a third expiry with
    contracts); BTC and ETH now match, like the US tool from this build.
    The render loop clears and hides each asset's third expiry columns
    (W:AE) on its next cycle — one-time, then they are never touched.
    """
    return 2


def safe_float(val):
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def fmt_oi(val):
    """Coerce OI to a plain float for Excel (so VLOOKUP/XLOOKUP works).

    Handles raw numbers, old Rs/INR text, and legacy "$1.50M"/"$2.00K" strings
    from a previous build so the one-off bulk cleanup passes.
    """
    if isinstance(val, str):
        s = val.replace("₹", "").replace("Rs", "").replace("INR", "").replace("$", "").replace(",", "").strip()
        s_lower = s.lower().rstrip()
        if s_lower.endswith("m"):
            try:
                return float(s_lower[:-1]) * 1_000_000
            except ValueError:
                return 0.0
        if s_lower.endswith("k"):
            try:
                return float(s_lower[:-1]) * 1_000
            except ValueError:
                return 0.0
        try:
            val = float(s)
        except ValueError:
            return 0.0
    if val is None or val != val:
        return 0.0
    return float(val)


def iso_to_dmy(iso):
    yy, mm, dd = iso.split("-")
    return f"{dd}-{mm}-{yy}"


def dmy_to_iso(dmy):
    dd, mm, yy = dmy.split("-")
    return f"{yy}-{mm}-{dd}"


_thread_local = threading.local()


def _session():
    # keep-alive per thread: ~2x faster than a fresh connection each call
    s = getattr(_thread_local, "session", None)
    if s is None:
        s = requests.Session()
        _thread_local.session = s
    return s


def set_feed_state(asset, state, detail="", retry_in=0.0):
    st = feed_state.setdefault(asset, {})
    if st.get("state") == state and st.get("detail") == detail:
        # The same state is being reported again. Writing it to the log every
        # time would put a line in the log every poll for as long as the
        # failure lasts.
        st["retry_in"] = retry_in
        return
    st["state"] = state
    st["detail"] = detail
    st["retry_in"] = retry_in
    feed_log("state", asset=asset, state=state, detail=detail,
             retry_in=round(retry_in, 1))


def note_feed_good(asset, label=""):
    st = feed_state.setdefault(asset, {})
    st["state"] = "OK"
    st["detail"] = ""
    st["retry_in"] = 0.0
    st["last_good_ts"] = time.time()
    if label:
        st["last_good_label"] = label


def feed_header_note(asset):
    """Text for the sheet header, or "" when there is nothing to say."""
    st = feed_state.get(asset) or {}
    state = st.get("state", "OK")
    if state == "OK":
        age = time.time() - (st.get("last_good_ts") or 0)
        if st.get("last_good_ts") and age > STALE_FLAG_SECONDS:
            return f"STALE: last good {st.get('last_good_label', '')}".strip()
        return ""
    last = st.get("last_good_label", "never")
    return f"{state}: {st.get('detail', '')} (last good {last})".strip()


def source_delay_note(asset):
    """Everything the header should say about how fresh the data is.

    Two things are worth saying: that the feed is degraded (a REST call
    failed, the socket dropped), and that the index quote has gone quiet.
    Both are silent otherwise, and a trader looking at a frozen spot has no
    way of telling a quiet market from a broken feed.

    Returns "" in the normal case, so the header text does not change and
    the cell is not rewritten on every cycle.
    """
    parts = []
    feed = feed_header_note(asset)
    if feed:
        parts.append(feed)
    try:
        age = spot_age(asset)
    except Exception:
        age = float("inf")
    if age != float("inf") and age > STALE_AFTER_S:
        parts.append(f"SPOT STALE ({int(age)}s)")
        if _parity_spot_cache.get(asset) is not None:
            parts.append(f"PARITY SPOT active")
    return " | ".join(parts)


def _get_json(url, params):
    """GET with keep-alive session, 429 backoff and one socket-reset retry.

    Every failure is now logged with enough detail to act on: endpoint, params,
    HTTP status, exception type, the head of the body and how long it took.
    A 429 honours Retry-After instead of guessing.
    """
    last = None
    started = time.perf_counter()
    for attempt in range(3):
        try:
            resp = _session().get(url, params=params, timeout=HTTP_TIMEOUT)
            if resp.status_code == 429:
                retry_after = resp.headers.get("Retry-After")
                try:
                    wait = float(retry_after)
                except (TypeError, ValueError):
                    wait = 1.5 * (attempt + 1)
                wait = min(max(wait, 0.5), 30.0)
                feed_log("rest_429", url=url.split("/v2/")[-1],
                         params=_flat(params), attempt=attempt + 1,
                         retry_after=retry_after, wait_s=wait)
                last = requests.HTTPError("429 Too Many Requests")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp.json()
        except (requests.ConnectionError, requests.exceptions.ChunkedEncodingError) as e:
            last = e
            feed_log("rest_conn", url=url.split("/v2/")[-1], params=_flat(params),
                     attempt=attempt + 1, error=type(e).__name__, detail=str(e)[:180])
            time.sleep(0.4)
        except Exception as e:
            last = e
            body = ""
            status = None
            r = locals().get("resp")
            if r is not None:
                status = getattr(r, "status_code", None)
                try:
                    body = r.text[:180]
                except Exception:
                    body = ""
            feed_log("rest_error", url=url.split("/v2/")[-1], params=_flat(params),
                     attempt=attempt + 1, status=status,
                     error=type(e).__name__, detail=str(e)[:180], body=body,
                     elapsed_ms=round((time.perf_counter() - started) * 1000, 1))
            break
    feed_log("rest_failed", url=url.split("/v2/")[-1], params=_flat(params),
             error=type(last).__name__ if last else "unknown",
             elapsed_ms=round((time.perf_counter() - started) * 1000, 1))
    raise last


def fetch_products(asset):
    """Products for one underlying.

    NOTE: the products endpoint filters on 'underlying_asset_symbols' (plural).
    The singular 'underlying_asset_symbol' is silently ignored and returns every
    asset's contracts, which previously leaked ETH/XAUT symbols into BTC's expiries.
    """
    url = f"{REST_BASE}/v2/products"
    params = {"contract_types": "call_options,put_options", "underlying_asset_symbols": asset}
    out = _get_json(url, params)["result"]
    # Belt-and-braces: never trust the filter alone.
    clean = []
    for p in out:
        ua = p.get("underlying_asset") or {}
        sym = ua.get("symbol") if isinstance(ua, dict) else p.get("underlying_asset_symbol")
        if sym in (None, asset):
            clean.append(p)
    return clean


def fetch_tickers(asset, expiry_dmy):
    """Fetch the option chain for ONE expiry.

    Delta filters on 'expiry_date' and it must be DD-MM-YYYY (sending ISO gives
    HTTP 400). Note there is also an 'expiry' param, but the API silently IGNORES
    it and returns every expiry at once, so it must not be used here.
    """
    url = f"{REST_BASE}/v2/tickers"
    params = {
        "contract_types": "call_options,put_options",
        "underlying_asset_symbols": asset,
        "expiry_date": expiry_dmy,
    }
    rows = _get_json(url, params)["result"]

    # Guard: 'success' is NOT sufficient. If the filter is silently ignored the
    # API still returns success plus every expiry for the asset. Verify the rows
    # really belong to the requested expiry and underlying.
    want_suffix = _dmy_to_suffix(expiry_dmy)
    ok = []
    for t in rows:
        sym = str(t.get("symbol") or "")
        if not sym:
            continue
        if f"-{asset}-" not in sym:
            continue
        if want_suffix and not sym.endswith(want_suffix):
            continue
        ok.append(t)
    if rows and not ok:
        print(
            f"[{asset}] expiry filter returned {len(rows)} rows but NONE match "
            f"{expiry_dmy} (suffix {want_suffix}); discarding",
            flush=True,
        )
        return []
    return ok


def _dmy_to_suffix(dmy):
    """'08-10-2026' -> '081026' (the symbol suffix Delta uses)."""
    parts = str(dmy).split("-")
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        dd, mm, yy = parts
        return f"{dd}{mm}{yy[2:]}"
    return ""


def _today_settlement(products, today_str):
    """When today's expiry actually settles, as an aware UTC datetime.

    Delta settles BTC and ETH options at 12:00 UTC and XAUT at 16:00 UTC, and
    it changes them without warning. Reading the time off the products is
    safer than hard-coding an hour: an expiry that disappears from the sheet
    before it has really expired is a trader's contract vanishing in the
    middle of a session, with three hours of trading left on it.
    """
    stamps = []
    for p in products:
        st = p.get("settlement_time") or ""
        if len(st) < 20 or st[:10] != today_str or not st.endswith("Z"):
            continue
        try:
            stamps.append(
                datetime.strptime(st[:19], "%Y-%m-%dT%H:%M:%S").replace(
                    tzinfo=timezone.utc
                )
            )
        except ValueError:
            continue
    return min(stamps) if stamps else None


def available_expiry_dates(asset, products, now=None):
    now = now or datetime.now(timezone.utc)
    today_str = now.strftime("%Y-%m-%d")

    dates = sorted(set(
        p["settlement_time"][:10] for p in products
        if p["settlement_time"][:10] >= today_str
    ))

    # Today's expiry stays in the list until it has actually settled.
    settle = _today_settlement(products, today_str)
    if settle is not None and now >= settle and dates and dates[0] == today_str:
        dates = dates[1:]

    return dates


def load_expiry_tickers(asset, expiry_iso):
    expiry_dmy = iso_to_dmy(expiry_iso)
    tickers = fetch_tickers(asset, expiry_dmy)
    return apply_tickers(asset, tickers, expiry_iso)


def apply_tickers(asset, tickers, expiry_iso):
    """Merge ticker rows into live_data; record symbol->expiry so new strikes render."""
    n = 0
    for t in tickers:
        symbol = t.get("symbol", "")
        if not symbol:
            continue
        ct = t.get("contract_type", "")
        if ct not in ("call_options", "put_options"):
            continue
        # A single malformed row must never abort the whole batch: this runs on
        # the OI poll loop, so an unhandled exception here would silently kill
        # updates for that asset for the rest of the session.
        strike = safe_float(t.get("strike_price"))
        if strike <= 0:
            continue
        sp = t.get("spot_price")
        if sp is not None:
            spot_cmp[asset] = safe_float(sp)
        ltp = safe_float(t.get("close"))
        if ltp == 0.0:
            ltp = safe_float(t.get("mark_price"))
        quotes = t.get("quotes") or {}
        entry = {
            "strike": int(strike),
            "type": "call" if ct == "call_options" else "put",
            "ltp": ltp,
            "bid": safe_float(quotes.get("best_bid")),
            "ask": safe_float(quotes.get("best_ask")),
            "oi": safe_float(t.get("oi_value_usd")),
        }
        if symbol not in live_data[asset]:
            live_data[asset][symbol] = entry
        else:
            live_data[asset][symbol].update(entry)
        if expiry_iso:
            symbol_expiry[asset][symbol] = expiry_iso
            if expiry_iso not in expiry_symbols[asset]:
                expiry_symbols[asset][expiry_iso] = []
            if symbol not in expiry_symbols[asset][expiry_iso]:
                expiry_symbols[asset][expiry_iso].append(symbol)
        n += 1
    return n


def probe_expiry_count(asset, expiry_iso):
    """Contracts for one expiry.

    Returns (count, state). 'OK' means a real count. 'ERROR' means the request
    failed and the count is meaningless -- these must never be confused, which
    is exactly how a transient failure used to look like 'no contracts'.
    """
    try:
        tickers = fetch_tickers(asset, iso_to_dmy(expiry_iso))
        n = len(tickers)
        # HTTP 200 is not proof of good data: check the symbols really belong to
        # this expiry before trusting the count.
        _yy, _mm, _dd = expiry_iso.split("-")
        suffix = f"{_dd}{_mm}{_yy[2:]}"          # symbols carry DDMMYY
        ok_suffix = sum(1 for t in tickers if suffix in str(t.get("symbol", "")))
        matched = (ok_suffix == n) if n else True
        feed_log("probe_ok", asset=asset, expiry=expiry_iso, rows=n,
                 suffix_ok=ok_suffix, suffix_checked=matched)
        if n == 0:
            feed_log("probe_no_contracts", asset=asset, expiry=expiry_iso, rows=0)
            return 0, "NO_CONTRACTS"
        if not matched:
            feed_log("probe_suffix_mismatch", asset=asset, expiry=expiry_iso,
                     rows=n, suffix_ok=ok_suffix, suffix=suffix, level="warn")
        return n, "OK"
    except Exception as e:
        feed_log("probe_error", asset=asset, expiry=expiry_iso,
                 error=type(e).__name__, detail=str(e)[:180])
        return 0, "ERROR"


async def oi_poller(asset):
    # stagger assets so their request bursts never align (rate-limit friendly)
    await asyncio.sleep({"BTC": 0.0, "ETH": 0.7, "XAUT": 1.4}.get(asset, 0.0))
    while True:
        try:
            dmies = [d for d in active_expiries_dmy[asset] if d]
            if dmies:
                results = await asyncio.gather(
                    *[asyncio.to_thread(fetch_tickers, asset, d) for d in dmies],
                    return_exceptions=True,
                )
                for dmy, res in zip(dmies, results):
                    if isinstance(res, BaseException):
                        _quiet_console_error(
                            f"OI poller error ({asset} {dmy})", res
                        )
                        net_note(False)
                        set_feed_state(asset, "ERROR", str(res)[:80],
                                       retry_in=OI_POLL_SECONDS)
                        continue
                    apply_tickers(asset, res, dmy_to_iso(dmy))
                    net_note(True)
                    # This is the feed's heartbeat, and it is what the sheet
                    # header reads. The state used to be cleared only when an
                    # expiry rolled, so a perfectly healthy feed sat at
                    # LOADING for the whole session and the header said so on
                    # every row.
                    note_feed_good(asset, dmy_to_iso(dmy))
        except Exception as e:
            _quiet_console_error(f"OI poller error ({asset})", e)
            net_note(False)
            set_feed_state(asset, "ERROR", type(e).__name__,
                           retry_in=OI_POLL_SECONDS)
        await asyncio.sleep(OI_POLL_SECONDS)


async def spot_poller():
    """Fallback index price. Primary source is the WebSocket option tickers
    (see process_frame); this only fills the gap while WS is down."""
    assets = ["BTC", "ETH", "XAUT"]
    i = 0
    while True:
        asset = assets[i % len(assets)]
        i += 1
        try:
            # Only consult the perp if the live index feeds have not produced a price.
            if spot_source.get(asset) in ("ws", "index-public"):
                await asyncio.sleep(SPOT_POLL_SECONDS)
                continue

            def fetch(a=asset):
                return _get_json(f"{REST_BASE}/v2/tickers/{SPOT_TICKER_SYMBOLS[a]}", None)

            data = await asyncio.to_thread(fetch)
            res = data.get("result") or {}
            set_spot(asset, res.get("spot_price"), "rest-perp")
        except Exception as e:
            _quiet_console_error(f"Spot poller error ({asset})", e)
        await asyncio.sleep(SPOT_POLL_SECONDS)


def current_ws_symbols():
    syms = ["BTCUSDT", "ETHUSDT", "XAUTUSD"]
    for asset in ("BTC", "ETH", "XAUT"):
        for expiry in active_expiries[asset]:
            syms.extend(expiry_symbols[asset].get(expiry, []))
    # dedupe preserve order
    seen = set()
    out = []
    for s in syms:
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


async def ws_listener():
    # Option chain rows — legacy host, because v2/ticker only exists there.
    retry = 1
    while True:
        try:
            async with websockets.connect(WS_URL) as ws:
                retry = 1
                await ws.send(json.dumps(
                    {"type": "subscribe", "payload": {"channels": [
                        {"name": "v2/ticker", "symbols": current_ws_symbols()},
                    ]}}
                ))
                while True:
                    try:
                        message = await asyncio.wait_for(ws.recv(), timeout=30)
                    except asyncio.TimeoutError:
                        # keep subscription fresh for rolled/pinned expiries
                        await ws.send(json.dumps(
                            {"type": "subscribe", "payload": {"channels": [
                                {"name": "v2/ticker",
                                 "symbols": current_ws_symbols()},
                            ]}}
                        ))
                        continue
                    data = json.loads(message)
                    if data.get("type") == "v2/ticker":
                        process_frame(data)
                        net_note(True)
        except Exception:
            net_note(False)
            await asyncio.sleep(retry)
            retry = min(retry * 2, 60)


async def index_listener():
    """Header index price from the documented PUBLIC socket.

    The index feed is the one thing that has an official public home, and it
    arrives far faster (~0.25 s) than the copy embedded in v2/ticker option
    frames (~4.9 s). Frames here use compact keys: {"p","sy","ts","type"}.
    """
    retry = 1
    while True:
        try:
            async with websockets.connect(INDEX_WS_URL) as ws:
                retry = 1
                payload = [{"name": INDEX_CHANNEL,
                            "symbols": [INDEX_SYMBOLS[a] for a in ("BTC", "ETH", "XAUT")]}]
                await ws.send(json.dumps(
                    {"type": "subscribe", "payload": {"channels": payload}}
                ))
                while True:
                    try:
                        message = await asyncio.wait_for(ws.recv(), timeout=30)
                    except asyncio.TimeoutError:
                        continue
                    try:
                        data = json.loads(message)
                    except ValueError:
                        continue
                    if data.get("type") != INDEX_CHANNEL:
                        continue
                    # public socket uses "sy"/"p"; legacy uses "symbol"/"price"
                    sym = data.get("sy") or data.get("symbol")
                    px = data.get("p") if data.get("p") is not None else data.get("price")
                    asset = INDEX_ASSET.get(sym)
                    if asset and px is not None:
                        note_ws_recv(asset)
                        set_spot(asset, px, "index-public")
                        net_note(True)
        except Exception as e:
            _quiet_console_error("Index listener error", e)
            await asyncio.sleep(retry)
            retry = min(retry * 2, 60)


def process_frame(data):
    symbol = data.get("symbol", "")
    if not symbol:
        return
    if symbol in ("BTCUSDT",) or symbol.startswith(("C-BTC-", "P-BTC-")):
        asset = "BTC"
    elif symbol in ("ETHUSDT",) or symbol.startswith(("C-ETH-", "P-ETH-")):
        asset = "ETH"
    elif symbol in ("XAUTUSD",) or symbol.startswith(("C-XAUT-", "P-XAUT-")):
        asset = "XAUT"
    else:
        return
    note_ws_recv(asset)
    if symbol in ("BTCUSDT", "ETHUSDT", "XAUTUSD"):
        # Perp/USDT feed. Not the website's index — used only as a fallback, and
        # only when no option ticker has spoken yet.
        if spot_source.get(asset) not in ("ws", "index-public"):
            set_spot(asset, data.get("spot_price") or data.get("close"), "ws-perp")
        return
    ct = data.get("contract_type", "")
    if ct not in ("call_options", "put_options"):
        return
    # PRIMARY index source: the option ticker's spot_price. This is exactly what
    # https://www.delta.exchange/app/options_chain/ displays in its header.
    set_spot(asset, data.get("spot_price"), "ws")
    strike = safe_float(data.get("strike_price"))
    if strike <= 0:
        return
    ltp = safe_float(data.get("close"))
    if ltp == 0.0:
        ltp = safe_float(data.get("mark_price"))
    quotes = data.get("quotes") or {}
    bid = safe_float(quotes.get("best_bid"))
    ask = safe_float(quotes.get("best_ask"))
    if symbol not in live_data[asset]:
        live_data[asset][symbol] = {
            "strike": int(strike),
            "type": "call" if ct == "call_options" else "put",
            "ltp": ltp, "bid": bid, "ask": ask, "oi": 0.0,
        }
    else:
        live_data[asset][symbol]["ltp"] = ltp
        live_data[asset][symbol]["bid"] = bid
        live_data[asset][symbol]["ask"] = ask


def fmt_expiry(iso_str):
    try:
        parts = iso_str[:10].split("-")
        months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        return f"{int(parts[2])}-{months[int(parts[1])-1]}-{parts[0]}"
    except Exception:
        return iso_str[:10]


def print_chain(asset, chains_data):
    for bi, (chain, pcr, expiry) in enumerate(chains_data):
        sp = spot_cmp[asset]
        if chain is None or chain.empty or sp is None:
            continue
        sentiment = "BULLISH" if pcr > 1.0 else "BEARISH"
        print(f"{asset} Block {bi+1} | {fmt_expiry(expiry)} | Spot: {sp:.2f} | PCR: {pcr:.2f} ({sentiment}) | {len(chain)} strikes")
        call_cols = ["Call LTP", "Call Bid", "Call Ask", "Call OI"]
        put_cols = ["Put OI", "Put Bid", "Put Ask", "Put LTP"]
        header = "  ".join(f"{c:>10}" for c in call_cols)
        header += "  " + "Strike".rjust(8)
        header += "  " + "  ".join(f"{c:>10}" for c in put_cols)
        print(header)
        print("-" * len(header))
        for _, row in chain.iterrows():
            parts_c = []
            for c in call_cols:
                v = row.get(c, 0)
                parts_c.append(str(v).rjust(10) if c == "Call OI" else f"{float(v):>10.2f}")
            parts_p = []
            for c in put_cols:
                v = row.get(c, 0)
                parts_p.append(str(v).rjust(10) if c == "Put OI" else f"{float(v):>10.2f}")
            strike_str = f"{int(row['strike']):>8}"
            print(f"{'  '.join(parts_c)}  {strike_str}  {'  '.join(parts_p)}")
        print()


def _put_intrinsic_rows(asset, spot, frame):
    """Observer: note puts quoted below their intrinsic value; never change them.

    The exchange occasionally publishes a put near zero while the live spot
    sits far above its strike (79200 put at 0.4 against BTC ~81815). The
    trader sees the exchange's numbers faithfully - this only notes the
    anomaly in feed.log so support can tell a data blip from a real price.
    """
    if spot is None or spot <= 0 or frame is None or getattr(frame, "empty", True):
        return
    if "Put LTP" not in frame.columns:
        return
    now = time.time()
    for key in [k for k in _PUT_INTRINSIC_WARNED if k[0] == asset]:
        if now - _PUT_INTRINSIC_WARNED[key] < _PUT_INTRINSIC_WARN_INTERVAL:
            return
    strikes = frame["strike"].apply(safe_float)
    ltps = frame["Put LTP"].apply(safe_float)
    intrinsic = float(spot) - strikes
    below = (strikes > 0) & (ltps > 0) & (intrinsic > ltps)
    if not below.any():
        return
    row = frame.loc[below.idxmax()]
    strike, ltp = safe_float(row.get("strike")), safe_float(row.get("Put LTP"))
    _PUT_INTRINSIC_WARNED[(asset, str(row.get("strike")))] = now
    feed_log(
        "put_intrinsic_warning",
        asset=asset,
        strike=f"{strike:g}",
        spot=f"{float(spot):.2f}",
        ltp=str(row.get("Put LTP")),
        intrinsic=f"{float(spot) - strike:.2f}",
        note="exchange quote looks wrong; shown as-is",
    )


def build_option_chain(asset, expiry_iso, range_val=50):
    sp = spot_cmp[asset]
    ld = live_data[asset]
    se = symbol_expiry[asset]
    if sp is None:
        return None, 0.0

    filtered = {}
    for sym, data in ld.items():
        if se.get(sym) == expiry_iso:
            filtered[sym] = data
    if not filtered:
        return None, 0.0

    all_calls = [e for e in filtered.values() if e["type"] == "call"]
    all_puts = [e for e in filtered.values() if e["type"] == "put"]
    if not all_calls or not all_puts:
        return None, 0.0

    strikes = sorted(set(e["strike"] for e in filtered.values()))
    atm_strike = min(strikes, key=lambda s: abs(s - sp))
    atm_idx = strikes.index(atm_strike)
    start = max(0, atm_idx - range_val)
    end = min(len(strikes), start + range_val * 2 + 1)
    if end - start < range_val * 2 + 1:
        start = max(0, end - range_val * 2 - 1)
    target_strikes = set(strikes[start:end])

    df_calls = pd.DataFrame([e for e in all_calls if e["strike"] in target_strikes])\
        .rename(columns={"oi": "Call OI", "ltp": "Call LTP", "bid": "Call Bid", "ask": "Call Ask"})\
        .drop(columns=["type"], errors="ignore")
    df_puts = pd.DataFrame([e for e in all_puts if e["strike"] in target_strikes])\
        .rename(columns={"oi": "Put OI", "ltp": "Put LTP", "bid": "Put Bid", "ask": "Put Ask"})\
        .drop(columns=["type"], errors="ignore")

    merged = pd.merge(df_calls, df_puts, on="strike", how="outer")
    if merged.empty:
        return None, 0.0
    merged = merged.sort_values("strike", ascending=True).reset_index(drop=True)
    merged = merged.fillna(0.0)
    # PCR totals from RAW values (before $ formatting strips the M/K multipliers)
    total_call_oi = sum(safe_float(v) for v in merged["Call OI"])
    total_put_oi = sum(safe_float(v) for v in merged["Put OI"])
    merged["Call OI"] = merged["Call OI"].apply(fmt_oi)
    merged["Put OI"] = merged["Put OI"].apply(fmt_oi)
    pcr = total_put_oi / total_call_oi if total_call_oi > 0 else 0.0
    _put_intrinsic_rows(asset, sp, merged)
    cols = ["Call LTP", "Call Bid", "Call Ask", "Call OI", "strike", "Put OI", "Put Bid", "Put Ask", "Put LTP"]
    available = [c for c in cols if c in merged.columns]
    return merged[available], pcr
