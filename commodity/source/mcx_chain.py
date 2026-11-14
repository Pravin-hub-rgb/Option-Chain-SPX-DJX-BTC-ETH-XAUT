"""Option chain builder for the MCX Commodity Option Chain tool.

Discovers MCX option instrument keys from the Upstox instrument master
(`assets.upstox.com/.../MCX.json.gz`), groups calls and puts by underlying and
expiry, fetches bid/ask/last/OI via `/v2/market-quote/quotes`, and builds the
9-column block rows the Excel writer renders.

Mirrors delta_chain.py's interface so mcx_main consumes it the same way.
"""

import gzip
import io
import json
import math
import os
import re
import threading
import time
from datetime import datetime, timezone

import requests

import mcx_state
from mcx_diagnostics import feed_log
from mcx_state import (
    BLOCK_COLS,
    EXPIRY_REFRESH_SECONDS,
    HTTP_TIMEOUT,
    MCX_SEGMENT,
    MCX_SYMBOLS,
    NET_FAIL_LIMIT,
    ROW_EXTENT_DEFAULT,
    SPOT_POLL_SECONDS,
    STALE_AFTER_S,
    UPSTOX_BASE_URL,
    UPSTOX_INSTRUMENTS_URL,
    _net,
    _pending_date_loads,
    active_expiries,
    active_expiries_dmy,
    available_dates,
    expiry_symbols,
    feed_state,
    live_data,
    symbol_expiry,
    user_pinned,
)

_instrument_lock = threading.Lock()
_instrument_master = {}      # instrument_key -> contract dict
_instrument_by_symbol = {}   # symbol -> list of contract dicts (CE/PE options)
_instruments_loaded = False
_instruments_error = None

# symbol -> expiry_dmy -> list of instrument keys (CE then PE)
_expiry_keys_cache = {}

# symbol -> expiry_dmy -> built chain rows (list of 9-col row lists)
_chain_cache = {}
_chain_cache_ts = {}

_net_fail = {sym: 0 for sym in MCX_SYMBOLS}


def _now_ms():
    return int(datetime.now(tz=timezone.utc).timestamp() * 1000)


def _ms_to_dmy(ms):
    if not ms:
        return ""
    try:
        return datetime.fromtimestamp(ms / 1000, tz=timezone.utc).strftime("%d-%m-%Y")
    except (ValueError, OSError, OverflowError):
        return ""


def _dmy_to_iso(dmy):
    """DD-MM-YYYY -> YYYY-MM-DD."""
    try:
        dt = datetime.strptime(dmy, "%d-%m-%Y")
        return dt.strftime("%Y-%m-%d")
    except ValueError:
        return ""


def iso_to_dmy(iso):
    """YYYY-MM-DD -> DD-MM-YYYY (public wrapper)."""
    return _iso_to_dmy(iso)


def _iso_to_dmy(iso):
    """YYYY-MM-DD -> DD-MM-YYYY."""
    if not iso or len(iso) < 10:
        return ""
    return f"{iso[8:10]}-{iso[5:7]}-{iso[0:4]}"


def fetch_instruments(token):
    """Download and parse the MCX instrument master. Cached per process."""
    global _instrument_master, _instrument_by_symbol, _instruments_loaded, _instruments_error
    with _instrument_lock:
        if _instruments_loaded and _instrument_master:
            return _instrument_master
        if not UPSTOX_INSTRUMENTS_URL:
            _instruments_error = "no upstox_instruments_url configured"
            return {}
        try:
            headers = {"Authorization": f"Bearer {token}", "Accept": "application/gzip"}
            r = requests.get(UPSTOX_INSTRUMENTS_URL, headers=headers, timeout=60)
            if r.status_code != 200:
                _instruments_error = f"instruments HTTP {r.status_code}"
                return {}
            raw = gzip.decompress(r.content)
            instruments = json.loads(raw)
        except Exception as e:
            _instruments_error = f"instruments fetch failed: {e}"
            return {}

        master = {}
        by_symbol = {sym: [] for sym in MCX_SYMBOLS}
        for it in instruments:
            if not isinstance(it, dict):
                continue
            seg = str(it.get("segment", "")).upper()
            itype = str(it.get("instrument_type", "")).upper()
            if seg != MCX_SEGMENT or itype not in ("CE", "PE"):
                continue
            name = str(it.get("name", "")).upper()
            sym = str(it.get("trading_symbol", "")).upper()
            key = it.get("instrument_key")
            expiry_ms = it.get("expiry")
            strike = it.get("strike")
            if strike is None:
                sym_text = str(it.get("trading_symbol", ""))
                m = re.search(r"(\d+)\s+(CE|PE)\s", sym_text)
                if m:
                    try:
                        strike = float(m.group(1))
                    except (TypeError, ValueError):
                        strike = None
            if not key or not expiry_ms:
                continue
            expiry_ms = int(expiry_ms)
            contract = {
                "instrument_key": key,
                "trading_symbol": it.get("trading_symbol"),
                "name": name,
                "expiry_ms": expiry_ms,
                "expiry_dmy": _ms_to_dmy(expiry_ms),
                "expiry_iso": datetime.fromtimestamp(expiry_ms / 1000, tz=timezone.utc).strftime("%Y-%m-%d"),
                "strike": float(strike) if strike is not None else None,
                "type": itype,
            }
            master[key] = contract
            for u in MCX_SYMBOLS:
                if u in name or u in sym:
                    by_symbol[u].append(contract)
                    break

        _instrument_master = master
        _instrument_by_symbol = by_symbol
        _instruments_loaded = True
        _instruments_error = None
        feed_log("mcx_instruments_loaded", contracts=len(master),
                           per_symbol={u: len(v) for u, v in by_symbol.items()})
        return master


def instruments_error():
    return _instruments_error


def available_expiry_dates(symbol, products=None):
    """Sorted unique future expiries (DD-MM-YYYY) for a symbol."""
    _ensure_symbol_keys(symbol)
    expiries = set()
    for contract in mcx_state.live_data.get(symbol, {}).values():
        pass  # placeholder; real source is _instrument_by_symbol
    for contract in _instrument_by_symbol.get(symbol, []):
        expiries.add(contract["expiry_dmy"])
    now = datetime.now(tz=timezone.utc).date()
    future = []
    for dmy in expiries:
        try:
            dt = datetime.strptime(dmy, "%d-%m-%Y").date()
        except ValueError:
            continue
        if dt >= now:
            future.append(dmy)
    return sorted(future)


def _ensure_symbol_keys(symbol):
    """Build expiry->instrument_keys for a symbol from the master."""
    if symbol in _expiry_keys_cache:
        return
    keys_by_expiry = {}
    for contract in _instrument_by_symbol.get(symbol, []):
        keys_by_expiry.setdefault(contract["expiry_dmy"], []).append(contract)
    _expiry_keys_cache[symbol] = keys_by_expiry


def get_option_keys(symbol, expiry_dmy):
    """Instrument keys (CE then PE, strike ascending) for an expiry."""
    _ensure_symbol_keys(symbol)
    contracts = _expiry_keys_cache.get(symbol, {}).get(expiry_dmy, [])
    ces = sorted([c for c in contracts if c["type"] == "CE" and c.get("strike") is not None],
                 key=lambda c: c["strike"])
    pes = sorted([c for c in contracts if c["type"] == "PE" and c.get("strike") is not None],
                 key=lambda c: c["strike"])
    return ces, pes


def fetch_quotes(instrument_keys, token):
    """Fetch quotes for a batch of instrument keys. Returns dict key->quote."""
    if not instrument_keys:
        return {}
    keys_param = ",".join(instrument_keys)
    try:
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        r = requests.get(
            f"{UPSTOX_BASE_URL}/market-quote/quotes",
            headers=headers,
            params={"instrument_key": keys_param},
            timeout=HTTP_TIMEOUT,
        )
        if r.status_code != 200:
            return {}
        data = r.json().get("data", {})
        out = {}
        for k, v in data.items():
            norm_key = k.replace(":", "|")
            out[norm_key] = v
        return out
    except Exception:
        return {}


def _quote_price(q, field):
    if not isinstance(q, dict):
        return None
    ohlc = q.get("ohlc") or {}
    last = q.get("last_price")
    if field == "last":
        return last
    if field == "bid":
        depth = q.get("depth") or {}
        buy = depth.get("buy") or []
        if buy:
            return buy[0].get("price")
        return ohlc.get("close")
    if field == "ask":
        depth = q.get("depth") or {}
        sell = depth.get("sell") or []
        if sell:
            return sell[0].get("price")
        return ohlc.get("close")
    return None


def _quote_oi(q):
    if not isinstance(q, dict):
        return None
    return q.get("oi")


def build_chain(symbol, expiry_dmy, token, spot=None, range_val=9999):
    """Build 9-column block rows for a symbol/expiry.

    Returns (rows, pcr, n_contracts, fresh) where rows is a list of
    [call_ltp, call_bid, call_ask, call_oi, strike, put_oi, put_bid, put_ask,
    put_ltp]. Strikes are centered on the ATM strike within range_val when a
    spot price is supplied. `fresh` is True when at least one price/OI came
    back non-zero from Upstox.
    """
    ces, pes = get_option_keys(symbol, expiry_dmy)
    if not ces and not pes:
        return [], 0.0, 0, False

    keys = [c["instrument_key"] for c in ces] + [p["instrument_key"] for p in pes]
    quotes = fetch_quotes(keys, token)
    if not quotes:
        feed_log("mcx_quotes_empty", symbol=symbol, expiry=expiry_dmy,
                           keys=len(keys))
        return [], 0.0, 0, False

    ce_by_strike = {c["strike"]: c for c in ces}
    pe_by_strike = {p["strike"]: p for p in pes}
    all_strikes = sorted(set(ce_by_strike.keys()) | set(pe_by_strike.keys()))

    if spot:
        try:
            sp = float(spot)
        except (TypeError, ValueError):
            sp = None
    else:
        sp = None

    if sp is not None and all_strikes:
        atm_strike = min(all_strikes, key=lambda s: abs(s - sp))
        atm_idx = all_strikes.index(atm_strike)
        start = max(0, atm_idx - int(range_val))
        end = min(len(all_strikes), start + int(range_val) * 2 + 1)
        if end - start < int(range_val) * 2 + 1:
            start = max(0, end - int(range_val) * 2 - 1)
        target = set(all_strikes[start:end])
        all_strikes = [s for s in all_strikes if s in target]

    rows = []
    total_call_oi = 0.0
    total_put_oi = 0.0
    n = 0
    any_price = False
    for strike in all_strikes:
        row = [None] * 9
        row[4] = strike  # strike column
        ce = ce_by_strike.get(strike)
        pe = pe_by_strike.get(strike)
        if ce:
            q = quotes.get(ce["instrument_key"], {})
            row[0] = _quote_price(q, "last")   # call ltp
            row[1] = _quote_price(q, "bid")    # call bid
            row[2] = _quote_price(q, "ask")    # call ask
            row[3] = _quote_oi(q)              # call oi
            for f in (row[0], row[1], row[2], row[3]):
                try:
                    if float(f or 0) != 0:
                        any_price = True
                except (TypeError, ValueError):
                    pass
            try:
                total_call_oi += float(row[3] or 0)
            except (TypeError, ValueError):
                pass
            n += 1
        if pe:
            q = quotes.get(pe["instrument_key"], {})
            row[5] = _quote_oi(q)              # put oi
            row[6] = _quote_price(q, "bid")    # put bid
            row[7] = _quote_price(q, "ask")    # put ask
            row[8] = _quote_price(q, "last")   # put ltp
            for f in (row[5], row[6], row[7], row[8]):
                try:
                    if float(f or 0) != 0:
                        any_price = True
                except (TypeError, ValueError):
                    pass
            try:
                total_put_oi += float(row[5] or 0)
            except (TypeError, ValueError):
                pass
            n += 1
        rows.append(row)

    pcr = total_put_oi / total_call_oi if total_call_oi > 0 else 0.0
    fresh = any_price

    # Cache last known good chain so the sheet still shows something marked
    # STALE when the market is closed and Upstox returns empty quotes.
    if fresh:
        mcx_state.last_good_chain.setdefault(symbol, {})[expiry_dmy] = {
            "rows": [list(r) for r in rows],
            "pcr": pcr,
            "ts": time.time(),
        }

    feed_log("mcx_chain_built", symbol=symbol, expiry=expiry_dmy,
                   strikes=len(rows), contracts=n, pcr=round(pcr, 4),
                   fresh=fresh)
    return rows, pcr, n, fresh


def get_last_good_chain(symbol, expiry_dmy):
    """Return cached last-good (rows, pcr) for a symbol/expiry, or (None, 0.0)."""
    entry = mcx_state.last_good_chain.get(symbol, {}).get(expiry_dmy)
    if not entry:
        return None, 0.0
    return entry["rows"], entry["pcr"]


def last_good_age(symbol, expiry_dmy):
    """Seconds since the last good chain was cached, or inf."""
    entry = mcx_state.last_good_chain.get(symbol, {}).get(expiry_dmy)
    if not entry:
        return float("inf")
    return time.time() - entry["ts"]


def apply_tickers(symbol, tickers, expiry_dmy):
    """Store fetched chain rows into live_data. Returns contract count."""
    live_data[symbol] = {str(row[4]): {"CE": {}, "PE": {}} for row in tickers}
    for row in tickers:
        strike = str(row[4])
        entry = live_data[symbol].setdefault(strike, {"CE": {}, "PE": {}})
        entry["CE"] = {
            "last": row[0], "bid": row[1], "ask": row[2], "oi": row[3],
        }
        entry["PE"] = {
            "last": row[8], "bid": row[6], "ask": row[7], "oi": row[5],
        }
    active_expiries[symbol] = [expiry_dmy] + [e for e in active_expiries.get(symbol, []) if e != expiry_dmy]
    active_expiries_dmy[symbol] = [_iso_to_dmy(_dmy_to_iso(expiry_dmy))] if expiry_dmy else []
    return len(tickers)


def load_expiry_tickers(symbol, expiry_dmy, token):
    """Fetch and apply the full chain for a symbol/expiry. Returns count."""
    rows, pcr, n, fresh = build_chain(symbol, expiry_dmy, token)
    if rows:
        apply_tickers(symbol, rows, expiry_dmy)
        note_feed_good(symbol, expiry_dmy)
    return n


def fetch_products(symbol, token=None):
    """Product list for a symbol (used by expiry roller to refresh expiries).

    Returns list of {symbol, settlement_time} dicts from the instrument master.
    """
    fetch_instruments(token) if token else None
    products = []
    now = _now_ms()
    seen = set()
    for contract in _instrument_by_symbol.get(symbol, []):
        try:
            exp_ms = int(contract["expiry_ms"])
        except (TypeError, ValueError):
            continue
        if exp_ms <= now:
            continue
        iso = contract["expiry_iso"]
        if iso in seen:
            continue
        seen.add(iso)
        products.append({
            "symbol": contract["instrument_key"],
            "settlement_time": iso,
        })
    return products


def probe_expiry_count(symbol, expiry_dmy):
    """Probe an expiry: returns (contract_count, state) where state is OK/ERROR.

    Callers check `count > 0` to know if an expiry has tradable contracts.
    """
    try:
        ces, pes = get_option_keys(symbol, expiry_dmy)
        count = len(ces) + len(pes)
        if count == 0:
            return 0, "ERROR"
        keys = [c["instrument_key"] for c in ces[:1]] + [p["instrument_key"] for p in pes[:1]]
        quotes = fetch_quotes(keys, None)  # token None -> no network; just structure check
        return count, "OK"
    except Exception:
        return 0, "ERROR"


def note_feed_good(symbol, expiry_dmy):
    fs = mcx_state.feed_state.setdefault(symbol, {"state": "OK", "detail": "", "retry_in": 0.0,
                                                   "last_good_ts": 0.0, "last_good_label": "never"})
    fs["state"] = "OK"
    fs["detail"] = expiry_dmy or ""
    fs["last_good_ts"] = time.time()
    fs["last_good_label"] = expiry_dmy or "never"
    fs["retry_in"] = 0.0
    mcx_state._net["fails"] = 0
    mcx_state._net["since"] = None


def set_feed_state(symbol, state, detail="", retry_in=0.0):
    mcx_state.feed_state[symbol] = {
        "state": state, "detail": detail, "retry_in": retry_in,
        "last_good_ts": mcx_state.feed_state.get(symbol, {}).get("last_good_ts", 0.0),
        "last_good_label": mcx_state.feed_state.get(symbol, {}).get("last_good_label", "never"),
    }


def fmt_oi(v):
    """Normalize open interest to a float."""
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def n_blocks(symbol):
    """Number of blocks to show for a symbol (3, or fewer if thin)."""
    try:
        expiries = available_expiry_dates(symbol)
    except Exception:
        expiries = []
    return min(3, max(1, len(expiries))) if expiries else 3


def spot_poller(symbol, token):
    """Poll the reference (futures LTP) price for a symbol. Returns price or None."""
    from mcx_spot import fetch_spot
    price, source = fetch_spot(symbol, token)
    if price is not None:
        set_spot(symbol, price, source)
    return price


def set_spot(symbol, price, source):
    from mcx_spot import set_spot as _set_spot
    _set_spot(symbol, price, source)


def oi_poller(symbol, token):
    """Refresh OI for the active expiry of a symbol."""
    expiries = active_expiries.get(symbol) or available_expiry_dates(symbol)
    if not expiries:
        return 0
    expiry = expiries[0]
    rows, pcr, n, fresh = build_chain(symbol, expiry, token)
    if rows:
        apply_tickers(symbol, rows, expiry)
    return n


def print_chain(symbol, expiry_dmy, max_rows=10):
    """Debug: print the first few rows of a built chain."""
    rows, pcr, n, fresh = build_chain(symbol, expiry_dmy, None)
    for row in rows[:max_rows]:
        print(row)


def _console_warn_last(msg, exc):
    print(f"[WARN] {msg}: {exc}", flush=True)


def _quiet_console_error(msg, exc):
    pass


def source_delay_note():
    return "MCX Upstox REST (market closed on weekends -> empty quotes)"
