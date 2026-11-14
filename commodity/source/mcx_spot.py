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


def _years_to_expiry(expiry_iso):
    """Fractional years from now to an ISO expiry date."""
    try:
        expiry = datetime.strptime(expiry_iso[:10], "%Y-%m-%d")
    except (ValueError, TypeError):
        return 0.0
    delta = expiry - datetime.utcnow()
    return max(0.0, delta.total_seconds() / (365.25 * 86400))


def _parity_spot(symbol, strikes):
    """Put-call-parity-implied spot for a symbol, or None when not enough data.

    Reads live_data[symbol][strike] for CE/PE marks at the same strike and
    backs out the reference price. Best-effort: returns None when the chain
    has not populated yet."""
    chain = mcx_state.live_data.get(symbol, {})
    if not chain:
        return None
    prices = []
    for strike, entry in chain.items():
        ce = entry.get("CE") or {}
        pe = entry.get("PE") or {}
        ce_mark = ce.get("mark") or ce.get("last")
        pe_mark = pe.get("mark") or pe.get("last")
        if ce_mark is None or pe_mark is None:
            continue
        try:
            strike_f = float(strike)
        except (TypeError, ValueError):
            continue
        # parity: C - P = S - K*exp(-r*t)  =>  S = C - P + K*exp(-r*t)
        t = _years_to_expiry(mcx_state.active_expiries.get(symbol, [None])[0] or "")
        r = mcx_state.PARITY_RATE
        discount = strike_f * math.exp(-r * t)
        prices.append((float(ce_mark) - float(pe_mark)) + discount)
    if not prices:
        return None
    return sum(prices) / len(prices)


def parity_spot_or_none(symbol):
    """Public accessor for the parity-implied spot."""
    strikes = list((mcx_state.live_data.get(symbol) or {}).keys())
    return _parity_spot(symbol, strikes)


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
