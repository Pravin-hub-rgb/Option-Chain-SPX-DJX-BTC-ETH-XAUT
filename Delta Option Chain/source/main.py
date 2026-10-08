"""
Delta Exchange Option Chain Tool (India endpoint)
Dual-asset: BTC + ETH simultaneously.
3-block layout with editable dates and dynamic strike range.
OI via REST poll (config oi_poll_seconds), spot via fast single-ticker poll
(config spot_poll_seconds), prices also via WebSocket, Excel update every 100ms.
"""

import asyncio
import bisect
import calendar
import hashlib
import hmac
import json
import os
import sys
import platform
import re
import socket
import threading
import time
from datetime import datetime, timezone, timedelta, date

_orig_getaddrinfo = socket.getaddrinfo

def _ipv4_getaddrinfo(*args, **kwargs):
    results = _orig_getaddrinfo(*args, **kwargs)
    v4 = [r for r in results if r[0] == socket.AF_INET]
    return v4 if v4 else results

socket.getaddrinfo = _ipv4_getaddrinfo

import pandas as pd
import requests
import websockets

import config

BASE_DIR = config.base_dir()

REST_BASE = "https://api.india.delta.exchange"

# Two sockets on purpose.
#
# INDEX_WS_URL is the documented PUBLIC host (changelog 17.04.26 migrated
# spot_price et al here). We use it for the header index price only — that feed
# is identical in shape and pushes every ~0.25 s.
#
# WS_URL is the legacy host, kept for the option chain rows because the public
# host has no `v2/ticker` channel: it was renamed to `ticker` there and uses a
# compact array schema (d[].q / d[].oi / d[].g) that is not documented. Decoding
# it is a separate job; until then the legacy channel keeps working for public
# data. Legacy public channels are slated for removal 31 Jul 2026.
WS_URL = "wss://socket.india.delta.exchange"
INDEX_WS_URL = "wss://public-socket.india.delta.exchange"

live_data = {"BTC": {}, "ETH": {}, "XAUT": {}}
spot_cmp = {"BTC": None, "ETH": None, "XAUT": None}
symbol_expiry = {"BTC": {}, "ETH": {}, "XAUT": {}}
expiry_symbols = {"BTC": {}, "ETH": {}, "XAUT": {}}
active_expiries = {"BTC": [], "ETH": [], "XAUT": []}
active_expiries_dmy = {"BTC": [], "ETH": [], "XAUT": []}
user_pinned = {"BTC": [False, False, False], "ETH": [False, False, False], "XAUT": [False, False, False]}
cfg = config.load_config()

HTTP_TIMEOUT = float(cfg.get("http_timeout_seconds", 15))
OI_POLL_SECONDS = float(cfg.get("oi_poll_seconds", 5))
EXPIRY_REFRESH_SECONDS = float(cfg.get("expiry_refresh_seconds", 30))
SPOT_POLL_SECONDS = float(cfg.get("spot_poll_seconds", 0.35))

# --- connection health indicator ------------------------------------------
# Internet ghatne par sheet freeze ho jata hai; user ko pata chalna chahiye
# ki data stale hai. 3 consecutive failures ke baad header me red tag dikhta
# hai, aur connection wapas aate hi apne aap hat jaata hai.
NET_FAIL_LIMIT = 3
_net = {"fails": 0, "since": None}
_net_drawn = {}   # (asset, block) -> last text written to that tag cell


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


def net_badge():
    """Header tag text ('' = sab theek, kuch nahi likhna)."""
    if _net["fails"] < NET_FAIL_LIMIT:
        return ""
    secs = int(time.time() - (_net["since"] or time.time()))
    m, s = divmod(secs, 60)
    span = f"{m}m {s:02d}s" if m else f"{s}s"
    return f"OFFLINE {span} - reconnecting..."


def net_paint(sheet, asset, bi, col_l):
    """Tag cell (row 1, block ka last column) likho — sirf jab badle."""
    tag = net_badge()
    prev = _net_drawn.get((asset, bi), None)
    if tag == prev:
        return
    try:
        cell = sheet.range((1, col_l + 8))
        cell.value = tag if tag else None
        if tag:
            cell.font.bold = True
            cell.font.color = (192, 0, 0)
    except Exception:
        return
    _net_drawn[(asset, bi)] = tag

# Product key: scrambled = (DDMMYYYY ^ mask8) ^ vmask(variant); all digits vary
# key = scrambled*10000 + variant*100 + SHA256 check (100 variants/date, sab valid)
_SK = [0x5A, 0xA7, 0x13, 0x66, 0xF0, 0x2B, 0x9C, 0x41]
_SE = [30, 255, 50, 81, 129, 123, 234, 115, 121, 245, 126, 95, 188, 83,
       168, 1, 13, 221, 43, 66, 164, 69, 169]

LICENSE_PATH = os.path.join(BASE_DIR, "license.json")
_LICENSE_SECRET = bytes(b ^ _SK[i % 8] for i, b in enumerate(_SE))


def _key_mask():
    return int(hashlib.sha256(_LICENSE_SECRET + b"|mask").hexdigest()[:16], 16) % 100000000


def _vmask(variant):
    h = hashlib.sha256(f"{variant}|vm|".encode() + _LICENSE_SECRET).hexdigest()
    return int(h[:16], 16) % 100000000


def _key_check(token, variant):
    h = hashlib.sha256(f"{token}|{variant}|".encode() + _LICENSE_SECRET).hexdigest()
    return int(h[:4], 16) % 100


_pending_date_loads = []
_date_req_sig = {}
_book_sheet_cache = {}
_last_rows = {}
_block_hidden = {}
available_dates = {"BTC": [], "ETH": [], "XAUT": []}
_dropdown_sig = {}
_usd_fmt_done = {}

MONTHS = {"JAN":"01","FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06",
          "JUL":"07","AUG":"08","SEP":"09","OCT":"10","NOV":"11","DEC":"12"}
MONTHS_REV = {v:k for k,v in MONTHS.items()}

BLOCK_COLS = [
    {"start": 1,  "atm": "A", "call_start": "A", "call_end": "D", "put_start": "F", "put_end": "I"},
    {"start": 12, "atm": "L", "call_start": "L", "call_end": "O", "put_start": "Q", "put_end": "T"},
    {"start": 23, "atm": "W", "call_start": "W", "call_end": "Z", "put_start": "AB", "put_end": "AE"},
]


def n_blocks(asset):
    """XAUT rarely has more than 2 expiries with contracts — show only 2 blocks."""
    return 2 if asset == "XAUT" else 3


def safe_float(val):
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def fmt_oi(val):
    # USD only — strip any ₹/INR leftovers and coerce to float
    if isinstance(val, str):
        s = val.replace("₹", "").replace("Rs", "").replace("INR", "").replace("$", "").replace(",", "").strip()
        try:
            val = float(s)
        except ValueError:
            return "$0"
    if val is None or val != val:
        return "$0"
    if val >= 1_000_000:
        return f"${val / 1_000_000:.2f}M"
    elif val >= 1_000:
        return f"${val / 1_000:.2f}K"
    else:
        return f"${val:.0f}"


def get_column_letter(col):
    """Convert 1-indexed column number to Excel letter (A, B, ..., Z, AA, AB, ...)"""
    result = ""
    while col > 0:
        col -= 1
        result = chr(65 + col % 26) + result
        col //= 26
    return result


def show_error(msg):
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Option Chain Tool", 0x10)
    else:
        print(msg, flush=True)


def show_info(msg):
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Option Chain Tool", 0x40)
    else:
        print(msg, flush=True)


def month_end(d_year, d_month):
    return date(d_year, d_month, calendar.monthrange(d_year, d_month)[1])


def current_month_end(today=None):
    today = today or date.today()
    return month_end(today.year, today.month)


def next_month_end(today=None):
    today = today or date.today()
    if today.month == 12:
        return month_end(today.year + 1, 1)
    return month_end(today.year, today.month + 1)


def key_to_date(key):
    """Decode hashed product key -> valid-until date, or None if invalid."""
    try:
        key = int(str(key).strip())
    except (TypeError, ValueError):
        return None
    if key <= 0:
        return None
    token_p, rest = divmod(key, 10000)
    variant, chk = divmod(rest, 100)
    token = token_p ^ _vmask(variant)
    if token <= 0 or not hmac.compare_digest(
        str(chk).zfill(2), str(_key_check(token, variant)).zfill(2)
    ):
        return None  # SHA256 check digits do not match (variant counted in)
    raw = token ^ _key_mask()
    s = f"{raw:08d}"
    if len(s) != 8:
        return None
    try:
        dd, mm, yyyy = int(s[0:2]), int(s[2:4]), int(s[4:8])
        if yyyy < 2020 or yyyy > 2100:
            return None
        return date(yyyy, mm, dd)  # calendar-safe: 30 Feb -> ValueError
    except ValueError:
        return None


def key_accepted(valid_until, today=None):
    """Key valid if its date is a future cycle end, max 45 days out."""
    today = today or date.today()
    return valid_until is not None and today < valid_until <= today + timedelta(days=45)


def _license_sig(data):
    canon = "|".join(str(data.get(k, "")) for k in ("valid_through", "issued", "key_month"))
    return hmac.new(_LICENSE_SECRET, canon.encode(), hashlib.sha256).hexdigest()


def _verify_sig(data):
    sig = data.get("sig")
    if not isinstance(sig, str):
        return False
    return hmac.compare_digest(sig, _license_sig(data))


def license_is_valid(today=None):
    today = today or date.today()
    try:
        with open(LICENSE_PATH) as f:
            data = json.load(f)
        if not _verify_sig(data):
            return False, None
        vt = date.fromisoformat(str(data.get("valid_through", "")))
        try:
            ls = date.fromisoformat(str(data.get("last_seen", "")))
        except ValueError:
            ls = None
        if ls is not None and ls > today:
            return False, None  # system clock rolled back -> reject
        if vt >= today:
            if ls != today:
                data["last_seen"] = today.isoformat()
                with open(LICENSE_PATH, "w") as f:
                    json.dump(data, f, indent=2)
            return True, vt
        return False, vt
    except (OSError, ValueError, KeyError, TypeError):
        # no license yet -> NOT pre-activated, ask for key on first run
        return False, None


def save_license(valid_until, today=None):
    today = today or date.today()
    payload = {
        "valid_through": valid_until.isoformat(),
        "issued": today.isoformat(),
        "key_month": f"{valid_until.year:04d}-{valid_until.month:02d}",
        "last_seen": today.isoformat(),
    }
    payload["sig"] = _license_sig(payload)
    with open(LICENSE_PATH, "w") as f:
        json.dump(payload, f, indent=2)


def ask_product_key():
    """Wide product-key popup. Returns stripped str, '' or None if cancelled."""
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        result = {"key": None}

        win = tk.Toplevel(root)
        win.title("Product Key")
        win.geometry("480x200")
        win.resizable(False, False)
        win.attributes("-topmost", True)

        tk.Label(
            win, text="License expired.",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", padx=24, pady=(22, 4))
        tk.Label(
            win, text="Enter product key (from owner):",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=24, pady=(0, 8))

        entry = tk.Entry(
            win, font=("Consolas", 13), width=32, justify="center",
            relief="solid", bd=1,
        )
        entry.pack(fill="x", padx=24)
        entry.focus_set()

        def on_ok():
            result["key"] = entry.get()
            win.destroy()

        def on_cancel():
            result["key"] = None
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=18)
        tk.Button(btns, text="OK", width=14, font=("Segoe UI", 10), command=on_ok).pack(side="left", padx=10)
        tk.Button(btns, text="Cancel", width=14, font=("Segoe UI", 10), command=on_cancel).pack(side="left", padx=10)

        win.bind("<Return>", lambda e: on_ok())
        win.bind("<Escape>", lambda e: on_cancel())
        win.protocol("WM_DELETE_WINDOW", on_cancel)

        # center on screen
        win.update_idletasks()
        x = (win.winfo_screenwidth() - 480) // 2
        y = (win.winfo_screenheight() - 200) // 2
        win.geometry(f"+{x}+{y}")

        root.wait_window(win)
        try:
            root.destroy()
        except Exception:
            pass
        k = result["key"]
        return k.strip() if isinstance(k, str) else k
    except Exception as e:
        print(f"Key dialog failed: {e}", flush=True)
        return None


def ensure_license():
    """Return True if app may run; False -> user must fix / we exit."""
    ok, vt = license_is_valid()
    if ok:
        return True

    # expired — require product key; invalid key -> ask again
    while True:
        key = ask_product_key()
        if key is None or str(key).strip() == "":
            show_error(
                "Product key required.\n\n"
                "Contact the owner for the key.\n"
                "App will close."
            )
            return False
        me = key_to_date(key)
        if key_accepted(me):
            save_license(me, date.today())
            show_info(
                f"License activated through {me.strftime('%d/%m/%Y')}.\n"
                "Continuing..."
            )
            return True
        show_error("Invalid product key. Please try again.")
        # loop back -> dialog reopens for another attempt


def open_excel_files():
    if platform.system() != "Windows":
        return
    for asset in ("BTC", "ETH", "XAUT"):
        path = os.path.join(BASE_DIR, f"{asset.lower()}_chain.xlsx")
        if not os.path.exists(path):
            show_error(f"File not found:\n{path}\n\nKeep this file next to OptionChain.exe and run again.")
            continue
        try:
            os.remove(path + ":Zone.Identifier")
        except OSError:
            pass
        try:
            os.startfile(path)
        except OSError as e:
            show_error(f"Could not open {asset.lower()}_chain.xlsx:\n{e}")


def iso_to_dmy(iso):
    yy, mm, dd = iso.split("-")
    return f"{dd}-{mm}-{yy}"


def dmy_to_iso(dmy):
    dd, mm, yy = dmy.split("-")
    return f"{yy}-{mm}-{dd}"


def parse_user_date(s):
    # Excel often coerces cell text into a datetime object before we see it
    if isinstance(s, datetime):
        return s.strftime("%Y-%m-%d")
    s = str(s).strip().upper()
    # primary: DD/MM/YYYY (matches US tool)
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        dd, mm, yyyy = m.groups()
        mi, di = int(mm), int(dd)
        if 1 <= mi <= 12 and 1 <= di <= 31:
            return f"{yyyy}-{mi:02d}-{di:02d}"
        return None
    # legacy: 26SEP26
    m = re.match(r"^(\d{2})([A-Z]{3})(\d{2})$", s)
    if m:
        dd, mmm, yy = m.groups()
        mm = MONTHS.get(mmm)
        if not mm:
            return None
        return f"20{yy}-{mm}-{dd}"
    # Excel datetime string e.g. "2026-09-26 00:00:00" / "26-09-2026"
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$", s)
    if m:
        return s[:10]
    m = re.match(r"^(\d{2})-(\d{2})-(\d{4})$", s)
    if m:
        dd, mm, yy = m.groups()
        return f"{yy}-{mm}-{dd}"
    return None


def iso_to_user(iso):
    # DD/MM/YYYY (matches US tool)
    if not iso or len(iso) < 10:
        return ""
    yyyy, mm, dd = iso[:10].split("-")
    return f"{dd}/{mm}/{yyyy}"


_thread_local = threading.local()


def _session():
    # keep-alive per thread: ~2x faster than a fresh connection each call
    s = getattr(_thread_local, "session", None)
    if s is None:
        s = requests.Session()
        _thread_local.session = s
    return s


def _get_json(url, params):
    """GET with keep-alive session, 429 backoff and one socket-reset retry."""
    last = None
    for attempt in range(3):
        try:
            resp = _session().get(url, params=params, timeout=HTTP_TIMEOUT)
            if resp.status_code == 429:
                last = requests.HTTPError("429 Too Many Requests")
                time.sleep(1.5 * (attempt + 1))
                continue
            resp.raise_for_status()
            return resp.json()
        except (requests.ConnectionError, requests.exceptions.ChunkedEncodingError) as e:
            last = e
            time.sleep(0.4)
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


def available_expiry_dates(asset, products):
    now = datetime.now(timezone.utc)
    today_str = now.strftime("%Y-%m-%d")

    dates = sorted(set(
        p["settlement_time"][:10] for p in products
        if p["settlement_time"][:10] >= today_str
    ))

    skip_hour = 16 if asset == "XAUT" else 8
    if now.hour >= skip_hour and dates and dates[0] == today_str:
        dates = dates[1:]

    return dates


def get_default_expiries(asset, products):
    return available_expiry_dates(asset, products)[:3]


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
        sp = t.get("spot_price")
        if sp is not None:
            spot_cmp[asset] = safe_float(sp)
        ltp = safe_float(t.get("close"))
        if ltp == 0.0:
            ltp = safe_float(t.get("mark_price"))
        entry = {
            "strike": int(t.get("strike_price", 0)),
            "type": "call" if ct == "call_options" else "put",
            "ltp": ltp,
            "bid": safe_float(t.get("quotes", {}).get("best_bid")),
            "ask": safe_float(t.get("quotes", {}).get("best_ask")),
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
    try:
        return len(fetch_tickers(asset, iso_to_dmy(expiry_iso)))
    except Exception:
        return 0


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
                        print(f"OI poller error ({asset} {dmy}): {res}", flush=True)
                        net_note(False)
                        continue
                    apply_tickers(asset, res, dmy_to_iso(dmy))
                    net_note(True)
        except Exception as e:
            print(f"OI poller error ({asset}): {e}", flush=True)
            net_note(False)
        await asyncio.sleep(OI_POLL_SECONDS)


# Perp tickers, used ONLY as a fallback for the index price. On these the
# 'close' field is the perpetual's own last trade (~30-45 USD from the index),
# so the website's number must never be sourced from here when a live option
# ticker is available.
SPOT_TICKER_SYMBOLS = {"BTC": "BTCUSD", "ETH": "ETHUSD", "XAUT": "XAUTUSD"}

# Dedicated index channel. Delta pushes the index here every ~0.25 s, versus
# ~4.9 s when it only rides along inside v2/ticker option frames (it refreshes
# only when an option prints). Symbols come from GET /v2/indices and are NOT
# derivable from the ticker name — XAUT's is .DEXAUTUSD, not .DEXXAUTUSD.
INDEX_SYMBOLS = {"BTC": ".DEXBTUSD", "ETH": ".DEETHUSD", "XAUT": ".DEXAUTUSD"}
INDEX_CHANNEL = "spot_price"
INDEX_ASSET = {v: k for k, v in INDEX_SYMBOLS.items()}

# Freshness tracking. Deliberately keyed on WebSocket RECIEVE time and on
# spot_price changes, NOT on per-option 'timestamp' — illiquid strikes can carry
# an old server timestamp while the index is moving, which would read as false
# staleness.
last_ws_recv = {"BTC": 0.0, "ETH": 0.0, "XAUT": 0.0}
last_spot_change = {"BTC": 0.0, "ETH": 0.0, "XAUT": 0.0}
spot_source = {"BTC": None, "ETH": None, "XAUT": None}
STALE_AFTER_S = 3.0


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
            print(f"Spot poller error ({asset}): {e}", flush=True)
        await asyncio.sleep(SPOT_POLL_SECONDS)


async def expiry_roller(asset):
    first = True
    while True:
        if first:
            first = False
        else:
            await asyncio.sleep(EXPIRY_REFRESH_SECONDS)
        try:
            products = await asyncio.to_thread(fetch_products, asset)
            # keep symbol->expiry / expiry->symbols fresh so new strikes render
            new_smap = {}
            new_esym = {}
            for p in products:
                sym = p.get("symbol", "")
                st = (p.get("settlement_time") or "")[:10]
                if not sym or not st:
                    continue
                new_smap[sym] = st
                new_esym.setdefault(st, []).append(sym)
            symbol_expiry[asset].update(new_smap)
            expiry_symbols[asset].update(new_esym)

            all_dates = available_expiry_dates(asset, products)
            available_dates[asset] = all_dates  # dropdown source (all future expiries)
            probe_dates = all_dates[:6]
            counts = await asyncio.gather(
                *[asyncio.to_thread(probe_expiry_count, asset, d) for d in probe_dates]
            )
            with_data = [d for d, c in zip(probe_dates, counts) if c > 0]

            pinned_vals = [active_expiries[asset][i] for i in range(n_blocks(asset)) if user_pinned[asset][i]]
            # only auto-assign expiries that actually have contracts;
            # if fewer than slots (e.g. XAUT often has 2), leave extra slots empty
            desired = [d for d in with_data if d not in pinned_vals]

            changed_slots = []
            slot = 0
            for i in range(n_blocks(asset)):
                if user_pinned[asset][i]:
                    continue
                cur = active_expiries[asset][i] if i < len(active_expiries[asset]) else None
                tgt = desired[slot] if slot < len(desired) else None
                slot += 1
                if tgt == cur:
                    continue
                if tgt:
                    await asyncio.to_thread(load_expiry_tickers, asset, tgt)
                changed_slots.append((i, cur, tgt))
                active_expiries[asset][i] = tgt
                active_expiries_dmy[asset][i] = iso_to_dmy(tgt) if tgt else None

            if changed_slots:
                try:
                    import xlwings as xw
                    book = find_open_book(xw, os.path.join(BASE_DIR, f"{asset.lower()}_chain.xlsx"))
                    if book:
                        sheet = book.sheets[0]
                        for i, old, new in changed_slots:
                            cell = ["D3", "O3", "Z3"][i]
                            _set_date_cell(sheet.range(cell), iso_to_user(new) if new else "")
                except Exception as e:
                    print(f"Roller cell write error ({asset}): {e}", flush=True)
                print(f"[{asset}] expiry rolled: " + ", ".join(f"{o or '--'} -> {n or '--'}" for _, o, n in changed_slots), flush=True)
        except Exception as e:
            print(f"Expiry roller error ({asset}): {e}", flush=True)


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
            print(f"Index listener error: {e}", flush=True)
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
    strike_str = data.get("strike_price")
    if strike_str is None:
        return
    ltp = safe_float(data.get("close"))
    if ltp == 0.0:
        ltp = safe_float(data.get("mark_price"))
    bid = safe_float(data.get("quotes", {}).get("best_bid"))
    ask = safe_float(data.get("quotes", {}).get("best_ask"))
    if symbol not in live_data[asset]:
        live_data[asset][symbol] = {
            "strike": int(strike_str),
            "type": "call" if ct == "call_options" else "put",
            "ltp": ltp, "bid": bid, "ask": ask, "oi": 0.0,
        }
    else:
        live_data[asset][symbol]["ltp"] = ltp
        live_data[asset][symbol]["bid"] = bid
        live_data[asset][symbol]["ask"] = ask


# ---------- Terminal Display (Linux testing only) ----------

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
    cols = ["Call LTP", "Call Bid", "Call Ask", "Call OI", "strike", "Put OI", "Put Bid", "Put Ask", "Put LTP"]
    available = [c for c in cols if c in merged.columns]
    return merged[available], pcr


def update_user_dates(asset, book, sheet):
    """Read date cells only — HTTP loads are queued to date_loader (never blocks event loop)."""
    for i, cell in enumerate(["D3", "O3", "Z3"]):
        if i >= n_blocks(asset):
            # block not shown (e.g. XAUT 3rd) — keep its date cell empty
            rng = sheet.range(cell)
            if rng.value not in (None, ""):
                _set_date_cell(rng, "")
            user_pinned[asset][i] = False
            _date_req_sig.pop((asset, i), None)
            continue
        rng = sheet.range(cell)
        val = rng.value
        if val is None or (isinstance(val, str) and val.strip() == ""):
            _set_date_cell(rng, iso_to_user(active_expiries[asset][i]))
            user_pinned[asset][i] = False
            _date_req_sig.pop((asset, i), None)
            continue
        parsed = parse_user_date(val)
        if not parsed:
            continue
        # Excel datetime object or odd text → normalize back to DDMMMYY text
        if not (isinstance(val, str) and val.strip().upper() == iso_to_user(parsed)):
            _set_date_cell(rng, iso_to_user(parsed))
        if parsed == active_expiries[asset][i]:
            _date_req_sig.pop((asset, i), None)
            continue
        # queue once per distinct requested date (no retry storm on empty/error)
        if _date_req_sig.get((asset, i)) == parsed:
            continue
        _date_req_sig[(asset, i)] = parsed
        _pending_date_loads.append((asset, i, parsed))


def _set_date_cell(rng, text):
    """Write date as TEXT so Excel cannot coerce it back to a serial number."""
    try:
        rng.api.NumberFormat = "@"
    except Exception:
        pass
    rng.value = text


async def date_loader():
    while True:
        if not _pending_date_loads:
            await asyncio.sleep(0.2)
            continue
        asset, i, parsed = _pending_date_loads.pop(0)
        if i >= n_blocks(asset):
            continue
        try:
            tickers = await asyncio.to_thread(fetch_tickers, asset, iso_to_dmy(parsed))
            if not tickers:
                print(f"[{asset}] no contracts for {parsed}", flush=True)
                continue
            n = apply_tickers(asset, tickers, parsed)
            active_expiries[asset][i] = parsed
            active_expiries_dmy[asset][i] = iso_to_dmy(parsed)
            user_pinned[asset][i] = True
            print(f"[{asset}] block {i + 1} pinned -> {parsed} ({n} contracts)", flush=True)
        except Exception as e:
            # transient failure — allow requeue after short backoff (no 10/s storm)
            _date_req_sig.pop((asset, i), None)
            print(f"Date fetch error ({asset} {parsed}): {e}", flush=True)
            await asyncio.sleep(3)


excel_fail_since = {"BTC": None, "ETH": None, "XAUT": None}
excel_connected = {"BTC": False, "ETH": False, "XAUT": False}
excel_miss_counts = {"BTC": 0, "ETH": 0, "XAUT": 0}
excel_first_attempt = {"BTC": None, "ETH": None, "XAUT": None}


def excel_apps_count(xw):
    try:
        return len(list(xw.apps))
    except Exception:
        return -1


def find_open_book(xw, fname):
    target = os.path.abspath(fname).lower()
    try:
        apps = list(xw.apps)
    except Exception:
        return None
    for app in apps:
        try:
            books = list(app.books)
        except Exception:
            continue
        for book in books:
            try:
                if os.path.abspath(book.fullname).lower() == target:
                    return book
            except Exception:
                continue
    return None


def force_usd_block_formats(sheet, block, asset, bi, n_rows):
    """Prices = 0.00, OI = text with $ only (kills ₹ number format).

    Applied once per (asset, block, row-count) — after that the OI cells are
    text-formatted, so Excel can no longer parse "$503" into ₹503.
    """
    if n_rows <= 0:
        return
    key = (asset, bi, n_rows)
    if _usd_fmt_done.get((asset, bi)) == key:
        return
    c0 = block["start"]
    end = 4 + n_rows
    for off in (0, 1, 2, 6, 7, 8):
        col = get_column_letter(c0 + off)
        try:
            sheet.range(f"{col}5:{col}{end}").api.NumberFormat = "0.00"
        except Exception:
            pass
    try:
        sheet.range(f"{get_column_letter(c0 + 4)}5:{get_column_letter(c0 + 4)}{end}").api.NumberFormat = "0.0#"
    except Exception:
        pass
    for off in (3, 5):
        col = get_column_letter(c0 + off)
        rng = sheet.range(f"{col}5:{col}{end}")
        try:
            rng.api.NumberFormat = "@"
        except Exception:
            pass
        # rewrite numeric/₹ leftovers as $ text (format now text — no new conversions)
        for r in range(5, end + 1):
            cell = sheet.range((r, c0 + off))
            v = cell.value
            if v is None or v == "":
                continue
            if not isinstance(v, str):
                cell.value = fmt_oi(safe_float(v))
            elif "₹" in v or "Rs" in v or "INR" in v:
                cell.value = fmt_oi(v)
    _usd_fmt_done[(asset, bi)] = key


def setup_expiry_dropdown(sheet, asset):
    """Hidden helper column + in-cell dropdown on the expiry cells (like US tool)."""
    dates = available_dates[asset]
    if not dates:
        return
    try:
        sheet.range("AH1:AH500").clear()
        sheet.range("AH1:AH500").api.NumberFormat = "@"  # text BEFORE write — else "02/10/2026" -> Feb 10 date
        sheet.range(f"AH1:AH{len(dates)}").value = [[iso_to_user(d)] for d in dates]
        sheet.range((1, 34)).api.EntireColumn.Hidden = True
    except Exception as e:
        print(f"Dropdown helper write ({asset}): {e}", flush=True)
        return
    formula = f"=$AH$1:$AH${len(dates)}"
    for cell in ["D3", "O3", "Z3"][:n_blocks(asset)]:
        rng = sheet.range(cell)
        try:
            rng.api.NumberFormat = "@"  # text — stops Excel date conversion
        except Exception:
            pass
        try:
            rng.api.Validation.Delete()
        except Exception:
            pass
        try:
            rng.api.Validation.Add(Type=3, Formula1=formula)
            rng.api.Validation.ShowError = False  # dropdown = hint, typing still allowed
        except Exception as e:
            print(f"Dropdown error ({asset} {cell}): {e}", flush=True)


def write_to_excel(asset):
    try:
        import xlwings as xw
    except ImportError:
        show_error("Required component 'xlwings' is missing. Please use the provided OptionChain.exe.")
        os._exit(1)
    fname = os.path.join(BASE_DIR, f"{asset.lower()}_chain.xlsx")
    try:
        book = _book_sheet_cache.get(asset)
        if book is not None:
            try:
                _ = book.name  # COM liveness check
            except Exception:
                book = None
                _book_sheet_cache.pop(asset, None)
        if book is None:
            book = find_open_book(xw, fname)
            if book is not None:
                _book_sheet_cache[asset] = book
        if book is None:
            apps_count = excel_apps_count(xw)
            if apps_count > 0:
                return
            if excel_connected[asset]:
                excel_miss_counts[asset] += 1
                if excel_miss_counts[asset] >= 10:
                    os._exit(0)
                return
            if excel_first_attempt[asset] is None:
                excel_first_attempt[asset] = time.time()
            elif time.time() - excel_first_attempt[asset] > 90:
                show_error(
                    "Could not connect to Microsoft Excel.\n\n"
                    "Please run OptionChain.exe again."
                )
                os._exit(1)
            return
        excel_miss_counts[asset] = 0
        excel_connected[asset] = True
        sheet = book.sheets[0]

        v = sheet.range("A2").value
        if isinstance(v, str):
            try:
                v = float(v.strip())
            except ValueError:
                v = None
        user_provided = isinstance(v, (int, float)) and v >= 1

        # B2 hint — re-write only if missing (user can't permanently lose it)
        try:
            b2 = sheet.range("B2")
            if b2.value != "Total strikes — edit A2":
                b2.value = "Total strikes — edit A2"
                b2.font.bold = True
                b2.font.color = (89, 89, 89)
        except Exception:
            pass

        update_user_dates(asset, book, sheet)

        # expiry dropdown (list of DD/MM/YYYY dates) — refresh when list changes
        _sig = tuple(available_dates[asset])
        if _sig and _dropdown_sig.get(asset) != _sig:
            setup_expiry_dropdown(sheet, asset)
            _dropdown_sig[asset] = _sig

        sp = spot_cmp[asset] or 0

        if not user_provided:
            range_val = 9999  # A2 khali → exchange ke paas jitne strikes, sab (koi cap nahi)
        else:
            range_val = max(int(v) // 2, 1)  # user limit — A2 kabhi overwrite nahi hota

        nb = n_blocks(asset)
        for bi, block in enumerate(BLOCK_COLS):
            if bi >= nb:
                # block not shown for this asset — clear once, leave hidden
                col_l = block["start"]
                if _last_rows.get((asset, bi), 0) or _block_hidden.get((asset, bi)) is None:
                    try:
                        cs = get_column_letter(block["start"])
                        ce = get_column_letter(block["start"] + 8)
                        sheet.range(f"{cs}1:{ce}1000").clear()
                        sheet.range(f"{cs}:{ce}").api.Hidden = True
                    except Exception:
                        pass
                    _last_rows[(asset, bi)] = 0
                    _block_hidden[(asset, bi)] = True
                continue
            if _block_hidden.get((asset, bi)):
                try:
                    cs = get_column_letter(block["start"])
                    ce = get_column_letter(block["start"] + 8)
                    sheet.range(f"{cs}:{ce}").api.Hidden = False
                except Exception:
                    pass
                _block_hidden[(asset, bi)] = False
            expiry = active_expiries[asset][bi] if bi < len(active_expiries[asset]) else None
            if not expiry:
                col_l = block["start"]
                sheet.range((1, col_l)).value = f"{asset} Spot: --"
                sheet.range((1, col_l + 4)).value = "PCR: --"
                # clear any stale rows from a previously assigned expiry
                if _last_rows.get((asset, bi), 0):
                    try:
                        cs = get_column_letter(block["start"])
                        ce = get_column_letter(block["start"] + 8)
                        sheet.range(f"{cs}5:{ce}1000").clear()
                    except Exception:
                        pass
                    _last_rows[(asset, bi)] = 0
                continue

            chain, pcr = build_option_chain(asset, expiry, range_val)
            col_l = block["start"]
            cnt = len(chain) if chain is not None else 0
            sheet.range((1, col_l)).value = f"{asset} Spot: {sp:.2f} | {cnt} strikes"
            sheet.range((1, col_l + 4)).value = f"PCR: {pcr:.2f}"
            net_paint(sheet, asset, bi, col_l)

            if chain is not None and not chain.empty:
                sheet.range((5, block["start"])).options(index=False, header=False).value = chain

                strikes = chain["strike"].tolist()
                n = len(strikes)
                # shrink → clear stale rows below (avoids ghost data on expiry roll)
                prev_n = _last_rows.get((asset, bi), 0)
                if cnt < prev_n:
                    try:
                        cs = get_column_letter(block["start"])
                        ce = get_column_letter(block["start"] + 8)
                        sheet.range(f"{cs}{5 + cnt}:{ce}{4 + prev_n}").clear()
                    except Exception:
                        pass
                _last_rows[(asset, bi)] = cnt

                # force USD-only formats (kills ₹ currency rendering)
                force_usd_block_formats(sheet, block, asset, bi, cnt)

                # batched side colors: strikes sorted → <sp is prefix, >sp is suffix
                k = bisect.bisect_left(strikes, sp)
                if k > 0:
                    sheet.range(f"{block['call_start']}5:{block['call_end']}{4 + k}").color = (198, 224, 180)
                m = bisect.bisect_right(strikes, sp)
                if m < n:
                    sheet.range(f"{block['put_start']}{5 + m}:{block['put_end']}{4 + n}").color = (255, 200, 200)

                atm_strike = min(strikes, key=lambda s: abs(s - sp))
                atm_row = strikes.index(atm_strike) + 5
                atm_range = f"{block['atm']}{atm_row}:{block['put_end']}{atm_row}"
                sheet.range(atm_range).color = (255, 255, 0)
            else:
                clear_start = get_column_letter(block["start"])
                clear_end = get_column_letter(block["start"] + 8)
                sheet.range(f"{clear_start}5:{clear_end}1000").clear()
                _last_rows[(asset, bi)] = 0

        excel_fail_since[asset] = None

    except Exception as e:
        print(f"Excel error ({asset}): {e}", flush=True)
        now = time.time()
        if excel_fail_since[asset] is None:
            excel_fail_since[asset] = now
        elif now - excel_fail_since[asset] > 30:
            show_error(
                "Lost connection to Microsoft Excel.\n\n"
                "Please run OptionChain.exe again.\n\n"
                "Details: " + str(e)
            )
            os._exit(1)


async def display_loop():
    refresh = cfg.get("refresh_interval_seconds", 0.1)
    is_linux = platform.system() != "Windows"
    license_tick = 0
    while True:
        await asyncio.sleep(refresh)
        # re-check license roughly every ~60s while running
        license_tick += 1
        if license_tick >= max(1, int(60 / max(float(refresh), 0.05))):
            license_tick = 0
            ok, _vt = license_is_valid()
            if not ok:
                if not ensure_license():
                    os._exit(1)
        if is_linux:
            chains_data = {}
            for asset in ("BTC", "ETH", "XAUT"):
                blocks = []
                for bi, block in enumerate(BLOCK_COLS):
                    if bi >= n_blocks(asset):
                        continue
                    expiry = active_expiries[asset][bi] if bi < len(active_expiries[asset]) else None
                    if not expiry:
                        continue
                    chain, pcr = build_option_chain(asset, expiry, 999)
                    blocks.append((chain, pcr, expiry))
                chains_data[asset] = blocks

            os.system("clear")
            for asset in ("BTC", "ETH", "XAUT"):
                print_chain(asset, chains_data[asset])
            print(f"Refreshing every {refresh}s | Press Ctrl+C to stop", flush=True)
        else:
            for asset in ("BTC", "ETH", "XAUT"):
                write_to_excel(asset)


def init_asset(asset):
    products = fetch_products(asset)

    for p in products:
        sym = p["symbol"]
        st = p.get("settlement_time", "")[:10]
        symbol_expiry[asset][sym] = st
        if st not in expiry_symbols[asset]:
            expiry_symbols[asset][st] = []
        expiry_symbols[asset][st].append(sym)

    active_expiries[asset] = get_default_expiries(asset, products)
    active_expiries_dmy[asset] = [iso_to_dmy(e) for e in active_expiries[asset]]

    for expiry_iso in active_expiries[asset]:
        load_expiry_tickers(asset, expiry_iso)


async def main_async():
    open_excel_files()
    try:
        init_asset("BTC")
        init_asset("ETH")
        init_asset("XAUT")
    except Exception as e:
        show_error(
            "Could not reach Delta Exchange API.\n\n"
            "Please check your internet connection and run OptionChain.exe again.\n\n"
            "Details: " + str(e)
        )
        return
    await asyncio.gather(
        ws_listener(),
        index_listener(),
        spot_poller(),
        oi_poller("BTC"),
        oi_poller("ETH"),
        oi_poller("XAUT"),
        expiry_roller("BTC"),
        expiry_roller("ETH"),
        expiry_roller("XAUT"),
        date_loader(),
        display_loop(),
    )


def kill_previous_instances():
    if platform.system() != "Windows":
        return
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.windll.kernel32
    TH32CS_SNAPPROCESS = 0x2
    my_pid = os.getpid()
    parent_pid = None

    class PROCESSENTRY32(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(wintypes.ULONG)),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", ctypes.c_char * 260),
        ]

    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap == -1 or snap is None:
        return
    try:
        entry = PROCESSENTRY32()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
        matches = []
        if kernel32.Process32First(snap, ctypes.byref(entry)):
            while True:
                name = entry.szExeFile.decode(errors="ignore").lower()
                if name == "optionchain.exe":
                    matches.append((entry.th32ProcessID, entry.th32ParentProcessID))
                if not kernel32.Process32Next(snap, ctypes.byref(entry)):
                    break
    finally:
        kernel32.CloseHandle(snap)

    for pid, ppid in matches:
        if pid == my_pid:
            parent_pid = ppid
            break

    for pid, ppid in matches:
        if pid != my_pid and pid != parent_pid:
            h = kernel32.OpenProcess(0x0001, False, pid)
            if h:
                kernel32.TerminateProcess(h, 1)
                kernel32.CloseHandle(h)


def main():
    kill_previous_instances()
    if not ensure_license():
        os._exit(1)
    try:
        asyncio.run(main_async())
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
