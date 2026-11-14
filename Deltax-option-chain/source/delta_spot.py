"""Shared spot/freshness state and parity spot helpers for the Delta tool.

The index-price freshness state (note_ws_recv/set_spot/spot_age) and the
put-call parity spot estimator live here so both the feed plumbing and the
Excel render path in main.py use one shared module. State itself stays in
delta_state.py — this module only reads/mutates its dicts.
"""

import math
import time
from datetime import date

from delta_state import (
    PARITY_DEADBAND,
    PARITY_DIVIDEND_Q,
    PARITY_MAX_SPREAD,
    PARITY_MAX_SPREAD_PCT,
    PARITY_RATE,
    PARITY_USE_N_STRIKES,
    STALE_AFTER_S,
    _parity_spot_cache,
    active_expiries,
    last_spot_change,
    last_ws_recv,
    live_data,
    spot_cmp,
    spot_source,
    symbol_expiry,
)


# Freshness tracking. Deliberately keyed on WebSocket RECIEVE time and on
# spot_price changes, NOT on per-option 'timestamp' — illiquid strikes can carry
# an old server timestamp while the index is moving, which would read as false
# staleness.


def note_ws_recv(asset):
    """Called for every WebSocket frame touching this asset."""
    last_ws_recv[asset] = time.time()


def set_spot(asset, raw, source):
    """Record the index price and when it last actually moved."""
    if raw in (None, ""):
        return
    try:
        val = float(raw)
    except (TypeError, ValueError):
        return
    if spot_cmp.get(asset) != val:
        last_spot_change[asset] = time.time()
    spot_cmp[asset] = val
    spot_source[asset] = source


def spot_age(asset):
    """Seconds since the last WebSocket frame / last actual index move."""
    recv = last_ws_recv.get(asset) or 0.0
    moved = last_spot_change.get(asset) or 0.0
    ref = max(recv, moved) if (recv or moved) else 0.0
    if not ref:
        return float("inf")
    return time.time() - ref


def spot_is_stale(asset):
    return spot_age(asset) > STALE_AFTER_S


def _years_to_expiry(expiry_iso):
    """Days to expiry as a year fraction, same shape as the US tool."""
    if not expiry_iso:
        return None
    try:
        y, m, d = expiry_iso[:10].split("-")
        exp = date(int(y), int(m), int(d))
    except Exception:
        return None
    today = date.today()
    return max(0.0, (exp - today).days / 365.0)


def _parity_spot(asset, expiry_iso, spot_hint):
    """Index spot from put-call parity on our own live chain.

    Parity:  C - P = S*exp(-qT) - K*exp(-rT)
    =>  S = K*exp(-(r-q)T) + (C-P)*exp(qT)

    For crypto options q=0 (no dividend), r=PARITY_RATE. The median of the N
    strikes nearest the current spot estimate is taken, rejecting crossed and
    wide-spread quotes. A deadband smooths the cached value.
    """
    if not expiry_iso:
        return None
    ld = live_data.get(asset, {})
    if not ld:
        return None
    T = _years_to_expiry(expiry_iso)
    if T is None or T <= 0:
        return None
    forward_discount = math.exp(-(PARITY_RATE - PARITY_DIVIDEND_Q) * T)
    call_put_adj = math.exp(PARITY_DIVIDEND_Q * T)

    calls = {}
    puts = {}
    for sym, data in ld.items():
        se = symbol_expiry.get(asset, {}).get(sym)
        if se != expiry_iso:
            continue
        strike = data.get("strike")
        if not strike or strike <= 0:
            continue
        if data.get("type") == "call":
            calls[strike] = data
        elif data.get("type") == "put":
            puts[strike] = data

    common = set(calls) & set(puts)
    if not common:
        return None

    def mid(row):
        bid = row.get("bid") or 0.0
        ask = row.get("ask") or 0.0
        if bid > 0 and ask > 0:
            return (bid + ask) / 2.0
        return row.get("ltp") or 0.0

    def spread_ok(row):
        bid = row.get("bid") or 0.0
        ask = row.get("ask") or 0.0
        if bid <= 0 or ask <= 0 or ask < bid:
            return False
        mid_price = (bid + ask) / 2.0
        if mid_price <= 0:
            return False
        return (ask - bid) <= max(PARITY_MAX_SPREAD, mid_price * PARITY_MAX_SPREAD_PCT)

    def parity_for_strike(strike):
        c = calls[strike]
        p = puts[strike]
        if not spread_ok(c) or not spread_ok(p):
            return None
        call_mid = mid(c)
        put_mid = mid(p)
        if call_mid <= 0 or put_mid <= 0:
            return None
        return strike * forward_discount + (call_mid - put_mid) * call_put_adj

    hint = spot_hint if spot_hint else 0.0
    ranked = sorted(common, key=lambda s: abs(s - hint))
    take = ranked[:PARITY_USE_N_STRIKES]
    estimates = [est for s in take if (est := parity_for_strike(s)) is not None]
    if not estimates:
        return None
    estimates.sort()
    n = len(estimates)
    if n % 2 == 1:
        parity_val = estimates[n // 2]
    else:
        parity_val = (estimates[n // 2 - 1] + estimates[n // 2]) / 2.0

    cached = _parity_spot_cache.get(asset)
    if cached is not None and abs(parity_val - cached) < PARITY_DEADBAND:
        return cached
    _parity_spot_cache[asset] = parity_val
    return parity_val


def parity_spot_or_none(asset):
    """Best parity spot for the active expiry, or None."""
    expiry = active_expiries[asset][0] if active_expiries[asset] else None
    if not expiry:
        return None
    hint = spot_cmp.get(asset) or 0.0
    return _parity_spot(asset, expiry, hint)
