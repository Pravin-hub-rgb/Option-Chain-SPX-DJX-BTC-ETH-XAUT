"""Spot / reference price for the MCX Commodity Option Chain tool.

MCX options are priced off the underlying futures contract. The reference
("spot") is fetched from Upstox `/market-quote/ltp` for the nearest futures
instrument key (discovered from the MCX instrument master) and polled every
SPOT_POLL_SECONDS.

Mirrors delta_spot.py's interface so mcx_main / mcx_chain consume it the same
way."""

import math
import time
from datetime import datetime

import mcx_state
from mcx_state import SPOT_POLL_SECONDS, STALE_AFTER_S, SPOT_TICKER_SYMBOLS

import requests


def set_spot(symbol, price, source):
    """Record the latest spot price for a symbol."""
    mcx_state.spot_cmp[symbol] = {
        "price": float(price),
        "ts": time.time(),
        "source": source,
    }
    mcx_state.spot_source[symbol] = source
    mcx_state.last_spot_change[symbol] = time.time()


def spot_age(symbol):
    """Seconds since the spot price was last updated, or None if never set."""
    entry = mcx_state.spot_cmp.get(symbol)
    if not entry or entry.get("ts") is None:
        return None
    return time.time() - entry["ts"]


def spot_is_stale(symbol):
    """True when the spot price is older than STALE_AFTER_S or never set."""
    age = spot_age(symbol)
    if age is None:
        return True
    return age > STALE_AFTER_S


def note_ws_recv(symbol):
    """Record a WebSocket receive time for a symbol (keep-alive bookkeeping)."""
    mcx_state.last_ws_recv[symbol] = time.time()


def _mid(row):
    bid = row.get("bid") or 0.0
    ask = row.get("ask") or 0.0
    if bid > 0 and ask > 0:
        return (bid + ask) / 2.0
    return row.get("last") or row.get("ltp") or 0.0


def _spread_ok(row, max_spread_abs, max_spread_pct):
    bid = row.get("bid") or 0.0
    ask = row.get("ask") or 0.0
    if bid <= 0 or ask <= 0 or ask < bid:
        return False
    mid_price = (bid + ask) / 2.0
    if mid_price <= 0:
        return False
    return (ask - bid) <= max(max_spread_abs, mid_price * max_spread_pct)


def _years_to_expiry(expiry_iso):
    """Days to expiry as a year fraction, same shape as the US tool."""
    if not expiry_iso:
        return None
    try:
        y, m, d = expiry_iso[:10].split("-")
        exp = datetime(int(y), int(m), int(d))
    except Exception:
        return None
    today = datetime.utcnow().date()
    return max(0.0, (exp - today).days / 365.0)


def _parity_spot(symbol, expiry_iso, spot_hint):
    """Put-call-parity-implied spot for a symbol, full Delta parity.

    Parity:  C - P = S*exp(-qT) - K*exp(-rT)
    =>  S = K*exp(-(r-q)T) + (C-P)*exp(qT)

    Takes the median of the N strikes nearest the current spot estimate,
    rejecting crossed and wide-spread quotes. A deadband smooths the cached
    value."""
    if not expiry_iso:
        return None
    ld = mcx_state.live_data.get(symbol, {})
    if not ld:
        return None
    T = _years_to_expiry(expiry_iso)
    if T is None or T <= 0:
        return None
    r = mcx_state.PARITY_RATE
    q = mcx_state.PARITY_DIVIDEND_Q
    forward_discount = math.exp(-(r - q) * T)
    call_put_adj = math.exp(q * T)

    calls = {}
    puts = {}
    for sym_key, data in ld.items():
        se = mcx_state.symbol_expiry.get(symbol, {}).get(sym_key)
        if se != expiry_iso:
            continue
        strike = data.get("strike")
        if not strike or strike <= 0:
            continue
        if data.get("type") == "CE":
            calls[strike] = data
        elif data.get("type") == "PE":
            puts[strike] = data

    common = set(calls) & set(puts)
    if not common:
        return None

    def spread_ok(row):
        return _spread_ok(row, mcx_state.PARITY_MAX_SPREAD, mcx_state.PARITY_MAX_SPREAD_PCT)

    def parity_for_strike(strike):
        c = calls[strike]
        p = puts[strike]
        if not spread_ok(c) or not spread_ok(p):
            return None
        call_mid = _mid(c)
        put_mid = _mid(p)
        if call_mid <= 0 or put_mid <= 0:
            return None
        return strike * forward_discount + (call_mid - put_mid) * call_put_adj

    hint = spot_hint if spot_hint else 0.0
    ranked = sorted(common, key=lambda s: abs(s - hint))
    take = ranked[:mcx_state.PARITY_USE_N_STRIKES]
    estimates = [est for s in take if (est := parity_for_strike(s)) is not None]
    if not estimates:
        return None
    estimates.sort()
    n = len(estimates)
    if n % 2 == 1:
        parity_val = estimates[n // 2]
    else:
        parity_val = (estimates[n // 2 - 1] + estimates[n // 2]) / 2.0

    cached = mcx_state._parity_spot_cache.get(symbol)
    if cached is not None and abs(parity_val - cached) < mcx_state.PARITY_DEADBAND:
        return cached
    mcx_state._parity_spot_cache[symbol] = parity_val
    return parity_val


def parity_spot_or_none(symbol):
    """Best parity spot for the active expiry, or None."""
    expiries = mcx_state.active_expiries.get(symbol) or []
    expiry = expiries[0] if expiries else None
    if not expiry:
        return None
    hint = mcx_state.spot_cmp.get(symbol) or 0.0
    return _parity_spot(symbol, expiry, hint)


def fetch_spot(symbol, token):
    """Fetch the reference (futures LTP) price for a symbol from Upstox.

    Returns (price, source) or (None, None) on failure. Requires
    mcx_state.SPOT_FUTURES_KEY[symbol] to be populated from the instrument
    master first."""
    key = mcx_state.SPOT_FUTURES_KEY.get(symbol)
    if not key:
        return None, None
    try:
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        r = requests.get(
            f"{mcx_state.UPSTOX_BASE_URL}/market-quote/ltp",
            headers=headers,
            params={"instrument_key": key},
            timeout=mcx_state.HTTP_TIMEOUT,
        )
        if r.status_code != 200:
            return None, None
        data = r.json().get("data", {})
        # Upstox returns the key both as "MCX_FO|12345" and "MCX_FO:12345"
        entry = data.get(key) or data.get(key.replace("|", ":"))
        if not entry:
            return None, None
        price = entry.get("last_price")
        if price is None:
            return None, None
        return float(price), "upstox_ltp"
    except Exception:
        return None, None
