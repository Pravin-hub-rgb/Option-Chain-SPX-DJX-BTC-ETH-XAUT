"""Test-harness BTC option chain fetcher for the US tool.

Mirrors the Delta Exchange (India endpoint) REST fetch used by the Delta tool,
but emits rows in the shape US's us_chain pipeline expects (spot, rows,
expiries, source). BTC trades 24/7, so this gives a live test vehicle for the
US recalc/strikethrough fixes while the US equity market is closed.

Test harness only - delete this module (and the "btc" role wiring) after testing.
"""

import threading
import time
from datetime import datetime, timezone

import requests

REST_BASE = "https://api.delta.exchange"
HTTP_TIMEOUT = 15.0

_thread_local = threading.local()


def _session():
    s = getattr(_thread_local, "session", None)
    if s is None:
        s = requests.Session()
        _thread_local.session = s
    return s


def _get_json(url, params):
    resp = _session().get(url, params=params, timeout=HTTP_TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def _safe_float(val, default=0.0):
    if val is None:
        return default
    try:
        f = float(val)
    except (TypeError, ValueError):
        return default
    if f != f:  # NaN
        return default
    return f


def fetch_products(asset):
    url = f"{REST_BASE}/v2/products"
    params = {"contract_types": "call_options,put_options", "underlying_asset_symbols": asset}
    out = _get_json(url, params).get("result") or []
    clean = []
    for p in out:
        ua = p.get("underlying_asset") or {}
        sym = ua.get("symbol") if isinstance(ua, dict) else p.get("underlying_asset_symbol")
        if sym in (None, asset):
            clean.append(p)
    return clean


def available_expiry_dates(asset, products, now=None):
    now = now or datetime.now(timezone.utc)
    today_str = now.strftime("%Y-%m-%d")
    dates = sorted(set(
        p["settlement_time"][:10] for p in products
        if p.get("settlement_time") and p["settlement_time"][:10] >= today_str
    ))
    return dates


def iso_to_dmy(iso):
    if not iso:
        return ""
    try:
        y, m, d = iso[:10].split("-")
        return f"{d}-{m}-{y}"
    except Exception:
        return ""


def fetch_tickers(asset, expiry_dmy):
    url = f"{REST_BASE}/v2/tickers"
    params = {
        "contract_types": "call_options,put_options",
        "underlying_asset_symbols": asset,
        "expiry_date": expiry_dmy,
    }
    rows = _get_json(url, params).get("result") or []
    want_suffix = ""
    parts = str(expiry_dmy).split("-")
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        dd, mm, yy = parts
        want_suffix = f"{dd}{mm}{yy[2:]}"
    ok = []
    for t in rows:
        sym = str(t.get("symbol") or "")
        if not sym or f"-{asset}-" not in sym:
            continue
        if want_suffix and not sym.endswith(want_suffix):
            continue
        ok.append(t)
    if rows and not ok:
        print(f"[{asset}] expiry filter returned {len(rows)} rows but NONE match {expiry_dmy}; discarding", flush=True)
        return []
    return ok


def fetch_btc_payload(sym):
    """Return (spot, rows, expiries, source) in US pipeline shape."""
    asset = "BTC"
    products = fetch_products(asset)
    dates = available_expiry_dates(asset, products)
    to_fetch = dates[:4]

    rows = []
    spot = None
    for expiry_iso in to_fetch:
        expiry_dmy = iso_to_dmy(expiry_iso)
        try:
            tickers = fetch_tickers(asset, expiry_dmy)
        except Exception as e:
            print(f"[BTC] ticker fetch failed for {expiry_iso}: {e}", flush=True)
            continue
        for t in tickers:
            t_symbol = t.get("symbol", "")
            if not t_symbol:
                continue
            ct = t.get("contract_type", "")
            if ct not in ("call_options", "put_options"):
                continue
            strike = _safe_float(t.get("strike_price"))
            if strike <= 0:
                continue
            sp = t.get("spot_price")
            if sp is not None and spot is None:
                spot = _safe_float(sp)
            ltp = _safe_float(t.get("close"))
            if ltp == 0.0:
                ltp = _safe_float(t.get("mark_price"))
            quotes = t.get("quotes") or {}
            rows.append({
                "contract": t_symbol,
                "type": "call" if ct == "call_options" else "put",
                "strike": int(strike),
                "expiry": expiry_iso,
                "root": asset,
                "bid": _safe_float(quotes.get("best_bid")),
                "ask": _safe_float(quotes.get("best_ask")),
                "oi": _safe_float(t.get("oi_value_usd")),
                "ltp": ltp,
            })

    if spot is None or spot == 0.0:
        try:
            data = _get_json(f"{REST_BASE}/v2/tickers/BTCUSD", None)
            res = data.get("result") or {}
            spot = _safe_float(res.get("spot_price") or res.get("close"))
        except Exception as e:
            print(f"[BTC] spot fetch failed: {e}", flush=True)
            spot = 0.0

    return (spot if spot else 0.0, rows, dates, "delta")
