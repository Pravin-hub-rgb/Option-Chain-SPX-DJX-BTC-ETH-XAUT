"""
US Stock / S&P 500 Option Chain Tool
Primary: BigClawd free brokerage chain (near real-time, no key).
Fallback: CBOE delayed quotes (~15 min) when primary is down.
Three Excel files auto-open (Delta-style):
  spx_option_chain.xlsx       -> locked to SPX (S&P500 index)
  djx_option_chain.xlsx       -> locked to DJX (Dow index)
  us_stock_option_chain.xlsx  -> editable K1 (default AAPL)
Shared cache by symbol; one process drives all books.
"""

import asyncio
import calendar
import hashlib
import hmac
import math
import os
import platform
import re
import shutil
import subprocess
import time
from datetime import date, datetime, timedelta

import pandas as pd
import requests

import config
from us_diagnostics import (
    append_app_exception,
    append_connection_event,
    excel_process_details,
    append_performance_summary,
    write_startup_diagnostics,
)
from us_performance import PerformanceMetrics, itm_color_ranges

BASE_DIR = config.base_dir()
cfg = config.load_config()

POLL_GAP = float(cfg.get("poll_gap_seconds", 2))
ROUND_PAUSE = float(cfg.get("round_pause_seconds", 2))
EXPIRY_REFRESH = float(cfg.get("expiry_refresh_seconds", 60))
REFRESH = float(cfg.get("refresh_interval_seconds", 0.1))
SPX_SYMBOL = str(cfg.get("spx_symbol", "SPX")).upper()
DJX_SYMBOL = str(cfg.get("djx_symbol", "DJX")).upper()
DEFAULT_SYMBOL = str(cfg.get("default_symbol", "AAPL")).upper()

SOURCE = str(cfg.get("data_source", "auto")).lower()  # auto | bigclawd | cboe

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

# --- connection health indicator ------------------------------------------
# Market/API ghatne par sheet freeze ho jata hai; user ko pata chalna chahiye
# ki data stale hai. NET_FAIL_LIMIT consecutive failures ke baad header me red
# tag dikhta hai, aur connection wapas aate hi apne aap hat jaata hai.
NET_FAIL_LIMIT = 2
_net = {"fails": 0, "since": None}
_net_drawn = {}   # (role, block) -> last text written to that tag cell


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


def net_paint(sheet, role, bi, col_l):
    """Tag cell (row 1, block ka last column) likho — sirf jab badle."""
    tag = net_badge()
    prev = _net_drawn.get((role, bi), None)
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
    _net_drawn[(role, bi)] = tag

BIGCLAWD_CHAIN = "https://bigclawd.com/api/v1/trading/options/chain/{}"
BIGCLAWD_QUOTE = "https://bigclawd.com/api/v1/trading/quote/{}"
CBOE_URL = "https://cdn.cboe.com/api/global/delayed_quotes/options/{}.json"
YAHOO_CHART = "https://query1.finance.yahoo.com/v8/finance/chart/{}?interval=1d&range=1d"
HTTP_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "application/json",
}

# US index option roots — CBOE delayed needs "_" prefix; BigClawd partial support
INDEX_SYMBOLS = {
    "SPX", "SPXW", "XSP", "DJX", "DJXW", "NDX", "NDXP",
    "VIX", "VIXW", "RUT", "RUTW", "OEX", "XEO",
}
# BigClawd brokerage feed does not serve these index chains (503)
_BIGCLAWD_NO_INDEX = {"DJX", "DJXW", "XSP", "SPXW", "OEX", "XEO"}


def clean_symbol(sym):
    """User-facing ticker -> API root (strip ^, map . class shares)."""
    s = str(sym or "").strip().upper().replace(".", "-")
    if s.startswith("^"):
        s = s[1:]
    return s


def is_index_symbol(sym):
    return clean_symbol(sym) in INDEX_SYMBOLS


def base_root(sym):
    """DJXW -> DJX, SPXW -> SPX. Standard (non-weekly) root of an index family.

    CBOE index payloads interleave the standard chain (DJX/SPX, the "mega")
    with the weekly chain (DJXW/SPXW, the "mini") on the same strikes. Users
    want the standard contract, so every comparison is made against this root.
    """
    s = clean_symbol(sym)
    return s[:-1] if (len(s) > 2 and s.endswith("W") and s[:-1] in INDEX_SYMBOLS) else s


# Expiry -> set of roots seen for it in the latest payload, so the expiry
# picker can tell a standard expiry (has the base root) from a weekly-only one.
expiry_roots = {}
# base root -> expiry -> count of contracts that actually report OI > 0
oi_roots = {}


def _note_expiry_roots(sym, rows):
    """Record which roots exist per expiry, and whether any real OI showed up."""
    base = base_root(sym)
    seen = expiry_roots.setdefault(base, {})
    oi = oi_roots.setdefault(base, {})
    live = set()
    for r in rows:
        exp = r.get("expiry")
        root = r.get("root") or base
        if not exp:
            continue
        live.add(exp)
        roots = seen.setdefault(exp, set())
        roots.add(root)
        if safe_float(r.get("oi")) > 0:
            oi[exp] = oi.get(exp, 0) + 1
    # drop expiries no longer in the payload (e.g. after a source switch)
    for exp in list(seen):
        if exp not in live:
            del seen[exp]
            oi.pop(exp, None)


def standard_expiries(sym, all_expiries):
    """Rank expiries: standard contract first, then ones that actually have OI.

    An index like DJX lists weekly chains (DJXW only) every few days plus
    monthly/quarterly DJX. Sorting purely by date makes the default slots show
    the mini weekly chain, which is what users reported seeing.

    Within the standard group, prefer expiries carrying real open interest:
    CBOE sometimes lists a standard root for a near expiry but publishes OI=0
    for every strike, which leaves the default slot full of zeros. Those are
    still selectable, just not in the first slots. Weekly expiries follow.
    """
    base = base_root(sym)
    roots = expiry_roots.get(base) or {}
    std = [d for d in all_expiries if base in (roots.get(d) or ())]
    weekly = [d for d in all_expiries if d not in set(std)]

    def has_oi(exp):
        return bool(oi_roots.get(base, {}).get(exp))

    std.sort(key=lambda d: (0 if has_oi(d) else 1,))
    return std + weekly


def cboe_symbol(sym):
    """CBOE CDN path: indices use underscore prefix (_SPX, _DJX)."""
    s = clean_symbol(sym)
    if is_index_symbol(s):
        return f"_{s}"
    return s


def bigclawd_strike_count(sym):
    # Index chains are huge / flaky at strike_count=100 — use 50 (API default)
    return 50 if is_index_symbol(sym) else 100

# sticky preferred source per symbol after first success; "auto" tries BigClawd then CBOE
_source_choice = {}

ROLE_FILES = {
    "spx": "spx_option_chain.xlsx",
    "djx": "djx_option_chain.xlsx",
    "stock": "us_stock_option_chain.xlsx",
}
ROLE_ORDER = ("spx", "djx", "stock")
# roles whose K1 is locked (not user-editable)
LOCKED_ROLES = frozenset({"spx", "djx"})

# active ticker per Excel file/role (spx/djx fixed; stock from K1)
role_symbol = {"spx": SPX_SYMBOL, "djx": DJX_SYMBOL, "stock": DEFAULT_SYMBOL}
SYMBOL_CELL = "K1"

live_data = {}
spot_cmp = {}
active_expiries = {}
user_pinned = {}
all_known_expiries = {}


def active_symbols():
    """Unique symbols currently shown across all files (shared cache)."""
    out = []
    for role in ROLE_ORDER:
        s = role_symbol[role]
        if s and s not in out:
            out.append(s)
    return out


def ensure_sym_state(sym):
    if sym not in live_data:
        live_data[sym] = {}
        spot_cmp[sym] = None
        active_expiries[sym] = [None, None, None]
        user_pinned[sym] = [False, False, False]
        all_known_expiries[sym] = []

MONTHS = {"JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06",
          "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"}
MONTHS_REV = {v: k for k, v in MONTHS.items()}

BLOCK_COLS = [
    {"start": 1,  "atm": "A", "call_start": "A", "call_end": "D", "put_start": "F", "put_end": "I"},
    {"start": 12, "atm": "L", "call_start": "L", "call_end": "O", "put_start": "Q", "put_end": "T"},
    {"start": 23, "atm": "W", "call_start": "W", "call_end": "Z", "put_start": "AB", "put_end": "AE"},
]

CHAIN_COLS = ["Call LTP", "Call Bid", "Call Ask", "Call OI",
              "strike", "Put OI", "Put Bid", "Put Ask", "Put LTP"]

OPT_RE = re.compile(r"^([A-Z]+)(\d{6})([CP])(\d{8})$")

_rate_lock = asyncio.Lock()
_last_source_call = 0.0
_pending_date_loads = []
_dropdown_sig = {}
_refetch_now = False


def safe_float(val, default=0.0):
    if val is None:
        return default
    try:
        f = float(val)
    except (ValueError, TypeError):
        return default
    if math.isnan(f) or math.isinf(f):
        return default
    return f


def fmt_oi(val):
    # Force USD only — strip any ₹/INR and coerce to float
    if isinstance(val, str):
        s = val.replace("₹", "").replace("Rs", "").replace("INR", "").replace("$", "").replace(",", "").strip()
        try:
            val = float(s)
        except ValueError:
            return "$0"
    try:
        val = float(val)
    except (TypeError, ValueError):
        return "$0"
    if val >= 1_000_000:
        return f"${val / 1_000_000:.2f}M"
    if val >= 1_000:
        return f"${val / 1_000:.2f}K"
    return f"${val:.0f}"


def get_column_letter(col):
    result = ""
    while col > 0:
        col -= 1
        result = chr(65 + col % 26) + result
        col //= 26
    return result


def show_error(msg):
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "US Stock Option Chain", 0x10)
    else:
        print(msg, flush=True)


def show_info(msg):
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "US Stock Option Chain", 0x40)
    else:
        print(msg, flush=True)


def record_excel_connection(role, status, detail=""):
    if _excel_connection_status.get(role) == status:
        return
    _excel_connection_status[role] = status
    try:
        append_connection_event(role, status, detail)
    except OSError as e:
        print(f"Could not log Excel connection state ({role}): {e}", flush=True)


def log_excel_exception(role, exc, context="Excel/COM update"):
    now = time.monotonic()
    signature = (type(exc).__name__, str(exc))
    key = (role, context)
    previous = _last_excel_exception.get(key)
    if previous and previous[0] == signature and now - previous[1] < 60:
        return
    _last_excel_exception[key] = (signature, now)
    try:
        append_app_exception(f"{context} for {role}", exc)
    except OSError as log_error:
        print(f"Could not log {context} traceback ({role}): {log_error}", flush=True)


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
        import json
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
    import json
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


def excel_path(role):
    return os.path.join(BASE_DIR, ROLE_FILES[role])


def default_k1(role):
    return role_symbol.get(role) or DEFAULT_SYMBOL


def is_locked_role(role):
    return role in LOCKED_ROLES


_template_ready = set()


def _unmerge_all(ws):
    """Drop every merged range on the sheet.

    The layout is a plain grid (three side-by-side blocks, headers on row 4),
    so a merged range is never intentional -- it only turns up when the user
    merges cells by hand while tidying the sheet. Leaving one in place makes
    openpyxl hand back read-only MergedCell objects for everything inside it,
    and writing to those killed the app at startup. Unmerging is always safe
    because the labels we care about are rewritten right after.
    """
    if not ws.merged_cells.ranges:
        return False
    for rng in list(ws.merged_cells.ranges):
        try:
            ws.unmerge_cells(str(rng))
        except Exception:
            pass
    return True


def _safe_set(ws, row, col, val):
    """Write to a cell, skipping read-only merged leftovers instead of crashing."""
    try:
        cell = ws.cell(row=row, column=col)
        if type(cell).__name__ == "MergedCell":
            return False
        cell.value = val
        return True
    except (AttributeError, ValueError, TypeError):
        return False


def _write_template(ws, role):
    """(Re)write every label the tool owns. Returns True if anything changed."""
    changed = False
    headers = ["Call LTP", "Call Bid", "Call Ask", "Call OI", "Strike",
               "Put OI", "Put Bid", "Put Ask", "Put LTP"]
    for block in BLOCK_COLS:
        for i, h in enumerate(headers):
            r, c = 4, block["start"] + i
            cell = ws.cell(row=r, column=c)
            if getattr(cell, "value", None) == h:
                continue
            if _safe_set(ws, r, c, h):
                changed = True
    j1 = "Symbol (locked):" if is_locked_role(role) else "Symbol (change):"
    if _safe_set(ws, 1, 10, j1):
        changed = True
    if is_locked_role(role) or not normalize_symbol(ws.cell(row=1, column=11).value):
        if _safe_set(ws, 1, 11, default_k1(role)):
            changed = True
    if _safe_set(ws, 2, 2, "± strikes — edit A2"):
        changed = True
    for col, label in ((1, "Expiry 1"), (3, "(edit D3)"), (12, "Expiry 2"),
                       (14, "(edit O3)"), (23, "Expiry 3"), (25, "(edit Z3)")):
        if _safe_set(ws, 3, col, label):
            changed = True
    for row in (2, 3):
        cell = ws.cell(row=row, column=10)
        if getattr(cell, "value", None) and _safe_set(ws, row, 10, None):
            changed = True
    return changed


def _rebuild_fresh(path):
    """Last resort: swap a broken template for a clean one.

    The damaged file is kept as .bak so nothing the user did is silently
    thrown away, then an empty sheet is created for the caller to fill in.
    Read-only attribute is cleared first -- Excel's "read-only recommended" or
    a file copied off a locked share would otherwise make the replace fail.
    """
    try:
        if os.path.exists(path):
            try:
                os.chmod(path, 0o666)  # drop read-only flag
            except Exception:
                pass
            try:
                shutil.copy2(path, path + ".bak")
            except Exception:
                pass
            try:
                os.remove(path)
            except Exception as e:
                print(f"Could not replace {os.path.basename(path)}: {e}", flush=True)
    except Exception:
        pass
    from openpyxl import Workbook, load_workbook
    fresh = path
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Chain"
        wb.save(path)
        return load_workbook(path)
    except Exception as e:
        # folder itself is read-only -> write alongside with a fresh name
        print(f"Cannot write {os.path.basename(path)}: {e} - using alternate name", flush=True)
        fresh = path.replace(".xlsx", "_new.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Chain"
        wb.save(fresh)
        return load_workbook(fresh)


def ensure_excel_file(role):
    path = excel_path(role)
    if role in _template_ready and os.path.exists(path):
        return path
    created = False
    if not os.path.exists(path):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Chain"
        wb.save(path)
        created = True
    from openpyxl import load_workbook
    wb = None
    try:
        wb = load_workbook(path)
        ws = wb.active
    except Exception as e:
        # unreadable / corrupt workbook (truncated zip, half-written on a crash)
        print(f"Template unreadable ({role}: {e}) - rebuilding sheet", flush=True)
        wb = _rebuild_fresh(path)
        ws = wb.active
    need_save = created
    try:
        need_save = _unmerge_all(ws) or need_save
        need_save = _write_template(ws, role) or need_save
    except Exception as e:
        # anything unexpected in a file the user edited -> start clean
        print(f"Template repair failed ({role}: {e}) - rebuilding sheet", flush=True)
        wb = _rebuild_fresh(path)
        ws = wb.active
        try:
            _write_template(ws, role)
        except Exception as e2:
            print(f"Fresh template still failing ({role}: {e2})", flush=True)
        need_save = True
    if need_save:
        try:
            wb.save(path)
        except Exception as e:
            print(f"Template save failed ({role}: {e}) - rebuilding sheet", flush=True)
            wb = _rebuild_fresh(path)
            ws = wb.active
            try:
                _write_template(ws, role)
                wb.save(path)
            except Exception as e2:
                # last chance: leave the rebuilt sheet on disk as-is
                print(f"Could not save labels for {role}: {e2}", flush=True)
    _template_ready.add(role)
    return path


def open_excel_files():
    if platform.system() != "Windows":
        return
    for role in ROLE_ORDER:
        path = ensure_excel_file(role)
        try:
            os.remove(path + ":Zone.Identifier")
        except OSError:
            pass
        try:
            os.startfile(path)
        except OSError as e:
            show_error(f"Could not open {ROLE_FILES[role]}:\n{e}")
            try:
                append_app_exception(f"opening {role} workbook", e)
            except OSError as log_error:
                print(f"Could not log workbook open failure ({role}): {log_error}", flush=True)
            try:
                append_app_exception(f"opening {role} workbook", e)
            except OSError as log_error:
                print(f"Could not log workbook open failure ({role}): {log_error}", flush=True)


def normalize_symbol(raw):
    if raw is None:
        return ""
    s = str(raw).strip().upper().replace(".", "-")
    # strip hint text if user pasted label into K1
    for junk in ("(CHANGE)", "(CHANGEABLE)", "CHANGE:", "TICKER", "SYMBOL"):
        if junk in s:
            s = s.split(junk)[0].strip()
    s = s.strip(" :->")
    if s.startswith("^"):
        s = s[1:].strip()
    if not s or s in ("SYMBOL", "TICKER", "K1", "EDIT"):
        return ""
    return s


def read_symbol_cell(sheet, role):
    locked = default_k1(role) if is_locked_role(role) else None
    if locked is not None:
        # locked — force fixed ticker (user edits overwritten)
        if sheet.range(SYMBOL_CELL).value != locked:
            sheet.range(SYMBOL_CELL).value = locked
        return locked
    raw = sheet.range(SYMBOL_CELL).value
    sym = normalize_symbol(raw)
    if not sym:
        sheet.range(SYMBOL_CELL).value = role_symbol["stock"] or DEFAULT_SYMBOL
        return role_symbol["stock"] or DEFAULT_SYMBOL
    if sym != str(raw).strip():
        sheet.range(SYMBOL_CELL).value = sym
    return sym


def iso_to_user(iso):
    # DD/MM/YYYY
    if not iso or len(iso) < 10:
        return ""
    yyyy, mm, dd = iso[:10].split("-")
    return f"{dd}/{mm}/{yyyy}"


def parse_user_date(s):
    s = str(s).strip()
    # primary: DD/MM/YYYY
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        dd, mm, yyyy = m.groups()
        mi, di = int(mm), int(dd)
        if 1 <= mi <= 12 and 1 <= di <= 31:
            return f"{yyyy}-{mi:02d}-{di:02d}"
        return None
    # secondary: DD-MMM-YY / DDMMMYY (old)
    s2 = s.upper()
    m = re.match(r"^(\d{1,2})[- ]?([A-Z]{3})[- ]?(\d{2,4})$", s2)
    if m:
        dd, mmm, yy = m.groups()
        mm = MONTHS.get(mmm)
        if not mm:
            return None
        if len(yy) == 2:
            yy = "20" + yy
        return f"{yy}-{mm}-{int(dd):02d}"
    # ISO YYYY-MM-DD
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        return s
    return None


def yymmdd_to_iso(yymmdd):
    return f"20{yymmdd[0:2]}-{yymmdd[2:4]}-{yymmdd[4:6]}"


async def throttled(fn, *args, **kwargs):
    global _last_source_call
    async with _rate_lock:
        now = time.monotonic()
        wait = _last_source_call + POLL_GAP - now
        if wait > 0:
            await asyncio.sleep(wait)
        _last_source_call = time.monotonic()
        for attempt in range(3):
            try:
                return await asyncio.to_thread(fn, *args, **kwargs)
            except Exception as e:
                if attempt == 2:
                    raise
                backoff = 15 * (attempt + 1)
                print(f"Fetch error, retry in {backoff}s: {e}", flush=True)
                await asyncio.sleep(backoff)
                _last_source_call = time.monotonic()


def _price(ltp, bid, ask):
    ltp = safe_float(ltp)
    bid = safe_float(bid)
    ask = safe_float(ask)
    if ltp > 0:
        return ltp
    if bid > 0 and ask > 0:
        return (bid + ask) / 2
    # CBOE no-quote often sends bid=0, ask=20 sentinel — never mark lone ask as LTP
    if bid > 0:
        return bid
    return 0.0


def _is_no_quote(bid, ask):
    """CBOE illiquid sentinel: bid=0, ask=20 (exact)."""
    return safe_float(bid) == 0.0 and safe_float(ask) == 20.0


def _quote_quality(bid, ask, oi, ltp):
    """Higher is better — used when same strike has DJX vs DJXW (or SPX vs SPXW)."""
    b, a = safe_float(bid), safe_float(ask)
    real_book = 1 if (b > 0 and a > 0 and a >= b and not _is_no_quote(b, a)) else 0
    return (real_book, 1 if safe_float(ltp) > 0 else 0, safe_float(oi))


# Yahoo ^SPX/^RUT/^VIX/^DJX quotes are ~15 min delayed; GSPC/NDX/DJI/SPY are live.
YAHOO_SPOT_CANDIDATES = {
    "SPX": ("^GSPC", "^SPX"),
    "SPXW": ("^GSPC", "^SPX"),
    "DJX": ("^DJX",),
    "DJXW": ("^DJX",),
    "RUT": ("^RUT",),
    "NDX": ("^NDX",),
    "VIX": ("^VIX",),
}


def yahoo_spot(sym):
    """Freshest Yahoo spot — prefer live aliases over delayed ^SPX-style feeds."""
    root = clean_symbol(sym)
    if is_index_symbol(root):
        candidates = list(YAHOO_SPOT_CANDIDATES.get(root, (f"^{root}",)))
    else:
        candidates = [root]
    now = time.time()
    best_price = None
    best_age = None
    for cand in candidates:
        try:
            resp = requests.get(
                YAHOO_CHART.format(cand),
                headers=HTTP_HEADERS,
                timeout=10,
            )
            if resp.status_code != 200:
                continue
            result = ((resp.json().get("chart") or {}).get("result") or [None])[0]
            if not result:
                continue
            meta = result.get("meta") or {}
            price = safe_float(meta.get("regularMarketPrice")) or safe_float(
                meta.get("chartPreviousClose")
            )
            if price <= 0:
                continue
            ts = safe_float(meta.get("regularMarketTime")) or now
            age = abs(now - ts)
            if best_price is None or age < best_age:
                best_price, best_age = price, age
            if age <= 10:
                return price
        except Exception:
            continue
    return best_price


def _fresher_spot(sym, cboe_or_chain_spot):
    """Prefer Yahoo spot for indexes when available (CBOE current_price can be stale)."""
    if not is_index_symbol(sym):
        return cboe_or_chain_spot
    y = yahoo_spot(sym)
    if y:
        return y
    return cboe_or_chain_spot


def _fetch_cboe(sym):
    cboe_sym = cboe_symbol(sym)
    url = CBOE_URL.format(cboe_sym)
    # index payloads can be multi-MB (_SPX ~13MB)
    timeout = 60 if is_index_symbol(sym) else 20
    resp = requests.get(url, headers=HTTP_HEADERS, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()["data"]
    rows = []
    expiries = set()
    for o in data.get("options", []):
        m = OPT_RE.match(str(o.get("option", "")))
        if not m:
            continue
        _root, yymmdd, cp, strike_raw = m.groups()
        expiry = yymmdd_to_iso(yymmdd)
        expiries.add(expiry)
        strike = int(strike_raw) / 1000.0
        if strike <= 0:
            continue
        bid = safe_float(o.get("bid"))
        ask = safe_float(o.get("ask"))
        if _is_no_quote(bid, ask):
            bid = ask = 0.0
        rows.append({
            "contract": str(o.get("option", "")),
            "root": _root,
            "strike": strike,
            "type": "call" if cp == "C" else "put",
            "ltp": _price(o.get("last_trade_price"), bid, ask),
            "bid": bid,
            "ask": ask,
            "oi": safe_float(o.get("open_interest")),
            "volume": safe_float(o.get("volume")),
            "iv": safe_float(o.get("iv")),
            "delta": safe_float(o.get("delta")),
            "expiry": expiry,
        })
    # CBOE packs DJX+DJXW / SPX+SPXW on same strike — keep one per (expiry,type,strike)
    want_root = clean_symbol(sym)
    # Standard (non-weekly) root always wins over the weekly one — DJX > DJXW,
    # SPX > SPXW. Weekly chains are often illiquid yet can show a *tighter*
    # stale book, so scoring by quote quality first would wrongly promote the
    # mini contract. Root match is the dominant key; quality only breaks ties
    # when the same root appears twice.
    by_key = {}
    for r in rows:
        k = (r["expiry"], r["type"], r["strike"])
        prev = by_key.get(k)
        if prev is None:
            by_key[k] = r
            continue
        prev_score = (
            1 if prev.get("root") == want_root else 0,
            *_quote_quality(prev["bid"], prev["ask"], prev["oi"], prev["ltp"]),
        )
        new_score = (
            1 if r.get("root") == want_root else 0,
            *_quote_quality(r["bid"], r["ask"], r["oi"], r["ltp"]),
        )
        if new_score[:1] != prev_score[:1]:
            # different root (standard vs weekly) -> standard wins outright
            if new_score[0] > prev_score[0]:
                by_key[k] = r
        elif new_score > prev_score:
            by_key[k] = r
    rows = list(by_key.values())
    _note_expiry_roots(sym, rows)
    spot = _fresher_spot(sym, safe_float(data.get("current_price")) or None)
    return spot, rows, sorted(expiries), "cboe"


def _fetch_bigclawd(sym):
    root = clean_symbol(sym)
    is_idx = is_index_symbol(root)
    resp = requests.get(
        BIGCLAWD_CHAIN.format(root),
        headers=HTTP_HEADERS,
        params={"strike_count": bigclawd_strike_count(root)},
        timeout=45 if is_idx else 25,
    )
    resp.raise_for_status()
    payload = resp.json()
    expiries = set()
    rows = []
    for block in payload.get("expirations") or []:
        expiry = str(block.get("date") or "")
        if not expiry:
            continue
        expiries.add(expiry)
        for side_key, side in (("calls", "call"), ("puts", "put")):
            for o in block.get(side_key) or []:
                strike = safe_float(o.get("strike"))
                if strike <= 0:
                    continue
                bid = safe_float(o.get("bid"))
                ask = safe_float(o.get("ask"))
                if _is_no_quote(bid, ask):
                    bid = ask = 0.0
                g = o.get("greeks") or {}
                contract = f"{root}{expiry[2:4]}{expiry[5:7]}{expiry[8:10]}"
                contract += ("C" if side == "call" else "P")
                contract += f"{int(round(strike * 1000)):08d}"
                rows.append({
                    "contract": contract,
                    "strike": strike,
                    "type": side,
                    "ltp": _price(o.get("last"), bid, ask),
                    "bid": bid,
                    "ask": ask,
                    "oi": safe_float(o.get("open_interest")),
                    "volume": safe_float(o.get("volume")),
                    "iv": safe_float(o.get("implied_volatility")),
                    "delta": safe_float(g.get("delta")),
                    "expiry": expiry,
                    "root": root,
                })

    _note_expiry_roots(sym, rows)

    spot = None
    try:
        q = requests.get(BIGCLAWD_QUOTE.format(root), headers=HTTP_HEADERS, timeout=10)
        if q.status_code == 200:
            spot = safe_float(q.json().get("price")) or None
    except Exception:
        pass
    if spot is None and rows:
        # fall back: no single spot field on chain; leave None and keep prior
        spot = spot_cmp.get(sym)
    spot = _fresher_spot(sym, spot)

    return spot, rows, sorted(expiries), "bigclawd"


def fetch_symbol_payload(sym):
    """Return (spot, rows, expiries, source_name). Honors config data_source."""
    root = clean_symbol(sym)
    forced = SOURCE if SOURCE in ("bigclawd", "cboe") else None
    if forced:
        order = [forced]
    else:
        preferred = _source_choice.get(sym)
        # index roots BigClawd cannot serve -> CBOE first (skip doomed 503)
        if root in _BIGCLAWD_NO_INDEX:
            base = ["cboe", "bigclawd"]
        elif is_index_symbol(root):
            base = ["bigclawd", "cboe"]
        else:
            base = ["bigclawd", "cboe"]
        if preferred:
            order = [preferred] + base
        else:
            order = base
        # de-dupe while keeping order
        seen = set()
        order = [x for x in order if not (x in seen or seen.add(x))]

    errors = []
    for name in order:
        try:
            if name == "bigclawd":
                result = _fetch_bigclawd(sym)
            else:
                result = _fetch_cboe(sym)
            if _source_choice.get(sym) != name:
                _source_choice[sym] = name
                print(f"[{sym}] data source: {name}", flush=True)
            return result
        except Exception as e:
            errors.append(f"{name}: {e}")
            if forced:
                raise
            print(f"[{sym}] {name} failed, trying next: {e}", flush=True)
    raise RuntimeError("; ".join(errors) or "no data source")


async def load_symbol(sym):
    ensure_sym_state(sym)
    spot, rows, expiries, _src = await throttled(fetch_symbol_payload, sym)
    if spot:
        spot_cmp[sym] = spot

    all_known_expiries[sym] = expiries
    if active_expiries[sym] == [None, None, None]:
        dates = available_expiries(expiries, sym)
        active_expiries[sym] = (dates[:3] + [None, None, None])[:3]

    # full replace — drops stale foreign-root / expired keys after source switch
    fresh = {}
    for row in rows:
        key = row["contract"] or f"{sym}-{row['type']}-{row['strike']}-{row['expiry']}"
        fresh[key] = row
    live_data[sym] = fresh

    return {"options": rows, "n": len(rows)}


def available_expiries(all_expiries, sym=None):
    """Expiries to offer, standard (non-weekly) contracts first.

    Date order alone makes the default slots land on the weekly mini chain for
    indices, so ordering is re-ranked by whether the expiry actually carries
    the base root (DJX/SPX). Weekly expiries stay selectable, just not first.
    """
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    dates = sorted(d for d in all_expiries if d >= today)
    if now.hour >= 16 and dates and dates[0] == today:
        dates = dates[1:]
    if sym is None:
        return dates
    return standard_expiries(sym, dates)


async def init_symbol(sym):
    data = await load_symbol(sym)
    print(f"[{sym}] loaded {data.get('n', 0)} contracts | spot {spot_cmp[sym]}", flush=True)
    print(f"[{sym}] expiries: {', '.join(active_expiries[sym])}", flush=True)


async def data_poller():
    global _refetch_now
    while True:
        try:
            for sym in active_symbols():
                await load_symbol(sym)
            syms = active_symbols()
            spots = ", ".join(f"{s}={spot_cmp.get(s)}" for s in syms)
            print(f"[poll] {spots} rows={sum(len(live_data.get(s) or {}) for s in syms)}", flush=True)
            net_note(True)
            _refetch_now = False
        except Exception as e:
            print(f"Poller error ({','.join(active_symbols())}): {e}", flush=True)
            net_note(False)
        if _refetch_now:
            await asyncio.sleep(POLL_GAP)
        else:
            await asyncio.sleep(ROUND_PAUSE)


async def expiry_roller():
    while True:
        await asyncio.sleep(EXPIRY_REFRESH)
        for sym in active_symbols():
            try:
                ensure_sym_state(sym)
                dates = available_expiries(all_known_expiries[sym], sym)
                if not dates:
                    continue
                pinned = [active_expiries[sym][i] for i in range(3) if user_pinned[sym][i]]
                desired = [d for d in dates if d not in pinned][:3]
                di = 0
                changed = []
                for i in range(3):
                    if user_pinned[sym][i]:
                        continue
                    tgt = desired[di] if di < len(desired) else None
                    di += 1
                    cur = active_expiries[sym][i]
                    if tgt == cur:
                        continue
                    active_expiries[sym][i] = tgt
                    changed.append((i, tgt))
                    print(f"[{sym}] slot {i + 1}: {cur or '--'} -> {tgt or '--'}", flush=True)
                if changed:
                    # Rolling an expiry swaps the strike ladder, so the cached
                    # payload/shape must go or the new expiry's first write and
                    # recolour would be skipped as "unchanged".
                    _block_payload.clear()
                    _block_shape.clear()
                    # write back to every open book currently showing this symbol
                    try:
                        import xlwings as xw
                        for role in ROLE_ORDER:
                            if role_symbol[role] != sym:
                                continue
                            book = find_open_book(xw, excel_path(role))
                            if not book:
                                continue
                            sheet = book.sheets[0]
                            for i, tgt in changed:
                                cell = sheet.range(["D3", "O3", "Z3"][i])
                                try:
                                    cell.api.NumberFormat = "@"
                                except Exception:
                                    pass
                                cell.value = iso_to_user(tgt) if tgt else ""
                                cell.color = (255, 255, 153)
                                cell.font.bold = True
                            write_hint_labels(sheet, role)
                    except Exception as e:
                        print(f"Expiry write-back error ({sym}): {e}", flush=True)
            except Exception as e:
                print(f"Expiry roller error ({sym}): {e}", flush=True)


async def process_pending_dates():
    global _pending_date_loads
    while True:
        if _pending_date_loads:
            sym, i, parsed = _pending_date_loads.pop(0)
            try:
                if parsed not in all_known_expiries[sym]:
                    print(f"[{sym}] no contracts for {parsed}", flush=True)
                    continue
                active_expiries[sym][i] = parsed
                user_pinned[sym][i] = True
                # A different expiry means a different strike ladder, so drop the
                # cached payload/shape or the first write and the first recolour
                # of the new expiry would be suppressed.
                _block_payload.clear()
                _block_shape.clear()
                print(f"[{sym}] pinned block {i + 1} -> {parsed}", flush=True)
            except Exception as e:
                print(f"Date pin error ({sym} {parsed}): {e}", flush=True)
        else:
            await asyncio.sleep(0.5)


def build_option_chain(sym, expiry, range_val=50):
    sp = spot_cmp[sym]
    if sp is None or not expiry:
        return None, 0.0
    filtered = [d for d in live_data[sym].values() if d["expiry"] == expiry]
    if not filtered:
        return None, 0.0
    all_calls = [e for e in filtered if e["type"] == "call"]
    all_puts = [e for e in filtered if e["type"] == "put"]
    if not all_calls or not all_puts:
        return None, 0.0

    strikes = sorted(set(e["strike"] for e in filtered))
    atm_strike = min(strikes, key=lambda s: abs(s - sp))
    atm_idx = strikes.index(atm_strike)
    # Some index chains span a very wide strike range with a dense listing at
    # the money (DJX: ~42 strikes from 325..625 around a ~513 spot). Rendering
    # the whole span buries the liquid ATM region under far-OTM strikes whose OI
    # is legitimately 0 on every public feed, so the user only sees zeros.
    # Key off the span relative to spot, not the strike count: RUT lists 50
    # strikes across only ~9% of spot and its OI is fine, DJX spans ~60%.
    window = range_val
    if is_index_symbol(sym) and sp > 0:
        span_pct = (strikes[-1] - strikes[0]) / float(sp)
        if span_pct > 0.30 and len(strikes) < window * 2 + 1:
            window = max(5, min(15, (len(strikes) - 1) // 2))
    start = max(0, atm_idx - window)
    end = min(len(strikes), start + window * 2 + 1)
    if end - start < window * 2 + 1:
        start = max(0, end - window * 2 - 1)
    target = set(strikes[start:end])

    want = base_root(sym)

    def _best_by_strike(rows):
        """Pick one contract per strike: standard root first, then quote quality.

        Same strike can exist as both DJX and DJXW (or SPX/SPXW). Quality alone
        would sometimes pick the weekly because its book looks tighter, so root
        match dominates and quality only orders rows of the same root.
        """
        picked = {}
        for e in rows:
            s = e["strike"]
            if s not in target:
                continue
            prev = picked.get(s)
            if prev is None:
                picked[s] = e
                continue
            is_std_e = 1 if e.get("root") == want else 0
            is_std_p = 1 if prev.get("root") == want else 0
            if is_std_e != is_std_p:
                if is_std_e > is_std_p:
                    picked[s] = e
            elif _quote_quality(e["bid"], e["ask"], e["oi"], e["ltp"]) > _quote_quality(
                prev["bid"], prev["ask"], prev["oi"], prev["ltp"]
            ):
                picked[s] = e
        return picked

    best_calls = _best_by_strike(all_calls)
    best_puts = _best_by_strike(all_puts)
    call_rows = [{
        "strike": e["strike"],
        "Call LTP": e["ltp"],
        "Call Bid": e["bid"],
        "Call Ask": e["ask"],
        "Call OI": e["oi"],
    } for e in best_calls.values()]
    put_rows = [{
        "strike": e["strike"],
        "Put OI": e["oi"],
        "Put Bid": e["bid"],
        "Put Ask": e["ask"],
        "Put LTP": e["ltp"],
    } for e in best_puts.values()]

    df_calls = pd.DataFrame(call_rows)
    df_puts = pd.DataFrame(put_rows)
    if df_calls.empty or df_puts.empty:
        return None, 0.0

    merged = pd.merge(df_calls, df_puts, on="strike", how="outer")
    if merged.empty:
        return None, 0.0
    merged = merged.sort_values("strike").reset_index(drop=True).fillna(0.0)
    merged["Call OI"] = merged["Call OI"].apply(fmt_oi)
    merged["Put OI"] = merged["Put OI"].apply(fmt_oi)

    def oi_total(series):
        total = 0.0
        for v in series:
            s = str(v).replace("$", "")
            mult = 1.0
            if s.endswith("M"):
                mult, s = 1_000_000, s[:-1]
            elif s.endswith("K"):
                mult, s = 1_000, s[:-1]
            try:
                total += float(s) * mult
            except ValueError:
                pass
        return total

    tc, tp = oi_total(merged["Call OI"]), oi_total(merged["Put OI"])
    pcr = tp / tc if tc > 0 else 0.0
    available = [c for c in CHAIN_COLS if c in merged.columns]
    return merged[available], pcr


excel_connected = {r: False for r in ROLE_ORDER}
excel_miss_counts = {r: 0 for r in ROLE_ORDER}
excel_first_attempt = {r: None for r in ROLE_ORDER}


def excel_apps_count(xw):
    try:
        return len(list(xw.apps))
    except Exception as e:
        log_excel_exception("EXCEL", e, "enumerating Excel applications")
        return -1


def excel_process_running():
    """True if any EXCEL.EXE is alive, checked without touching COM.

    If Excel has crashed or been killed, the cached COM server is a dead
    reference and the next `xw.apps` / `app.books` call raises "RPC server is
    unavailable" and can take the whole process down with it. Checking the
    process table first means we simply stop writing and wait for the customer
    to bring Excel back, instead of dying silently.
    """
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        # Unknown -> assume it is alive so behaviour is unchanged.
        return True
    return "EXCEL.EXE" in (result.stdout or "").upper()


def find_open_book(xw, fname):
    target = os.path.abspath(fname).lower()
    role = next(
        (candidate for candidate, file_name in ROLE_FILES.items() if file_name.lower() in target),
        "unknown",
    )
    if not excel_process_running():
        # Do not touch COM at all while Excel is gone.
        return None
    try:
        apps = list(xw.apps)
    except Exception as e:
        log_excel_exception(role, e, "enumerating Excel applications")
        return None
    for app in apps:
        try:
            books = list(app.books)
        except Exception as e:
            log_excel_exception(role, e, "enumerating open workbooks")
            continue
        for book in books:
            try:
                if os.path.abspath(book.fullname).lower() == target:
                    return book
            except Exception as e:
                log_excel_exception(role, e, "reading workbook path")
                continue
    return None


def update_expiry_dropdowns(sym, sheet):
    dates = available_expiries(all_known_expiries[sym], sym)
    if not dates:
        return
    options = [iso_to_user(d) for d in dates]
    joined = ",".join(options)
    if _dropdown_sig.get(sym) == joined:
        return
    _dropdown_sig[sym] = joined
    # helper list in far-right column AH (avoids Excel 255-char inline limit)
    helper_col = 34  # AH
    helper_end = max(len(options), 40)
    try:
        sheet.range((1, helper_col), (helper_end, helper_col)).clear()
    except Exception as e:
        log_excel_exception(sym, e, "clearing expiry dropdown helper column")
    # xlwings Sheet has no .cell() — write helper column as one range
    try:
        col_letter = get_column_letter(helper_col)
        sheet.range(f"{col_letter}1:{col_letter}{len(options)}").value = [
            [s] for s in options
        ]
    except Exception as e:
        print(f"dropdown helper write: {e}", flush=True)
        return
    try:
        sheet.range((1, helper_col)).api.EntireColumn.Hidden = True
    except Exception as e:
        log_excel_exception(sym, e, "hiding expiry dropdown helper column")
    formula = f"=$AH$1:$AH${len(options)}"
    for cell in ("D3", "O3", "Z3"):
        try:
            sheet.range(cell).api.Validation.Delete()
        except Exception as e:
            log_excel_exception(sym, e, f"clearing expiry validation {cell}")
        try:
            sheet.range(cell).api.Validation.Add(
                Type=3,
                Formula1=formula,
            )
        except Exception as e:
            print(f"Dropdown error ({sym} {cell}): {e}", flush=True)


def update_user_dates(sym, sheet):
    for i, cell in enumerate(["D3", "O3", "Z3"]):
        rng = sheet.range(cell)
        try:
            rng.api.NumberFormat = "@"  # text — stops Excel date conversion / #####
        except Exception:
            pass
        val = rng.value
        # Excel may have converted text into a datetime — force back to DD/MM/YYYY text
        if isinstance(val, datetime):
            rng.value = val.strftime("%d/%m/%Y")
            continue
        if val is None or (isinstance(val, str) and val.strip() == ""):
            rng.value = iso_to_user(active_expiries[sym][i])
            user_pinned[sym][i] = False
            continue
        user_str = str(val).strip()
        user_up = user_str.upper()
        if user_up in ("EXPIRY", "DD/MM/YYYY", "DDMMMYY", "EDIT"):
            rng.value = iso_to_user(active_expiries[sym][i])
            user_pinned[sym][i] = False
            continue
        parsed = parse_user_date(user_str)
        if parsed and parsed != active_expiries[sym][i]:
            # normalize display to DD/MM/YYYY
            if not re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", user_str):
                try:
                    rng.value = iso_to_user(parsed)
                except Exception:
                    pass
            _pending_date_loads.append((sym, i, parsed))
        elif parsed and re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", user_str) is None:
            # same expiry but odd format — normalize
            try:
                rng.value = iso_to_user(parsed)
            except Exception:
                pass


_sheet_fmt_done = set()
_usd_fmt_done = set()
# Per (role, block) payload signature. The block range write is the single most
# expensive call in the render loop, and between ticks most rows are unchanged,
# so an unchanged chain is skipped entirely rather than rewritten.
_block_payload = {}
# Per (role, block) shape signature, so ATM/ITM recolouring and number
# formatting only run when the strike ladder actually changes.
_block_shape = {}
_performance_metrics = PerformanceMetrics()
_last_performance_flush = time.monotonic()
_performance_log_error_shown = False
_excel_connection_status = {}
_last_excel_exception = {}


def forget_workbook_formatting(path):
    normalized_path = os.path.abspath(path).lower()
    _sheet_fmt_done.discard(normalized_path)
    stale_keys = {
        key for key in _usd_fmt_done
        if isinstance(key, tuple) and key[0] == normalized_path
    }
    _usd_fmt_done.difference_update(stale_keys)


# Labels written every run (cheap) so user always sees changeable hints
# J1 is role-dependent — written inside write_hint_labels
HINT_LABELS = (
    ("B2", "± strikes — edit A2"),
    ("A3", "Expiry 1"),
    ("C3", "(edit D3)"),
    ("L3", "Expiry 2"),
    ("N3", "(edit O3)"),
    ("W3", "Expiry 3"),
    ("Y3", "(edit Z3)"),
)

_HINT_CLEAR = ("J2", "J3")


def write_hint_labels(sheet, role="stock"):
    j1_text = "Symbol (locked):" if is_locked_role(role) else "Symbol (change):"
    try:
        r = sheet.range("J1")
        r.value = j1_text
        r.font.bold = True
        r.font.color = (89, 89, 89)
    except Exception as e:
        log_excel_exception(role, e, "writing symbol label")
    for cell, text in HINT_LABELS:
        try:
            r = sheet.range(cell)
            r.value = text
            r.font.bold = True
            r.font.color = (89, 89, 89)
        except Exception as e:
            log_excel_exception(role, e, f"writing worksheet hint {cell}")
    for cell in _HINT_CLEAR:
        try:
            sheet.range(cell).value = None
        except Exception as e:
            log_excel_exception(role, e, f"clearing owned hint cell {cell}")
    # keep Symbol label from squishing (cheap — every run)
    try:
        sheet.range("J:J").api.ColumnWidth = 22
    except Exception as e:
        log_excel_exception(role, e, "setting symbol column width")


def format_sheet(sheet, role="stock"):
    """Bold headers, widths, input highlights ? heavy part once; hints every run."""
    # NOTE: the hint labels used to be written here on every cycle, before the
    # one-shot gate below. They are static text that depends only on whether the
    # role is locked, so re-writing 10 cells plus font settings every tick cost
    # ~107ms per sheet and dominated the render loop. They are written once in
    # the gated block at the bottom instead.

    key = None
    try:
        key = sheet.book.fullname.lower()
    except Exception:
        key = id(sheet)
    if key in _sheet_fmt_done:
        return

    HEADER_FILL = (31, 78, 121)       # dark blue
    INPUT_FILL = (255, 255, 153)      # light yellow = editable
    LABEL_FILL = (217, 217, 217)      # gray

    # column widths: prices/OI/strike ~11, expiry wider for DD/MM/YYYY
    widths = {}
    for block in BLOCK_COLS:
        for off in range(9):
            widths[block["start"] + off] = 11
    widths[3] = 11    # C = (edit D3) hint
    widths[4] = 13    # D = Expiry 1 input DD/MM/YYYY
    widths[9] = 22    # J Symbol label — full text
    widths[10] = 12   # K symbol value
    widths[14] = 11   # N = (edit O3)
    widths[15] = 13   # O Expiry 2
    widths[25] = 11   # Y = (edit Z3)
    widths[26] = 13   # Z Expiry 3
    widths[34] = 10   # AH helper
    try:
        for c, w in widths.items():
            sheet.columns(get_column_letter(c)).width = w
    except Exception as e:
        try:
            for c, w in widths.items():
                sheet.range(f"{get_column_letter(c)}:{get_column_letter(c)}").api.ColumnWidth = w
        except Exception as fallback_error:
            log_excel_exception(role, e, "setting worksheet column widths")
            log_excel_exception(role, fallback_error, "setting fallback column widths")

    try:
        # Row 1 spot / PCR
        for block in BLOCK_COLS:
            sheet.range((1, block["start"]), (1, block["start"] + 4)).font.bold = True
            sheet.range((1, block["start"]), (1, block["start"] + 4)).font.size = 12
        # Row 4 headers
        for block in BLOCK_COLS:
            hdr = sheet.range((4, block["start"]), (4, block["start"] + 8))
            hdr.font.bold = True
            hdr.font.color = (255, 255, 255)
            hdr.color = HEADER_FILL
        # Labels + hints
        for cell, _t in HINT_LABELS:
            sheet.range(cell).font.bold = True
        for cell in ("B2", "A3", "L3", "W3", "C3", "N3", "Y3", "J1"):
            sheet.range(cell).color = LABEL_FILL
        # Input cells highlighted (do NOT wipe user values)
        # Locked files (SPX/DJX): K1 gray + forced — not an input
        input_cells = ["A2", "D3", "O3", "Z3"]
        if role == "stock":
            input_cells.append("K1")
        for cell in input_cells:
            r = sheet.range(cell)
            r.color = INPUT_FILL
            r.font.bold = True
        if is_locked_role(role):
            k1 = sheet.range("K1")
            k1.color = LABEL_FILL
            k1.font.bold = True
            k1.value = default_k1(role)
        elif not normalize_symbol(sheet.range("K1").value):
            sheet.range("K1").value = DEFAULT_SYMBOL
        for cell in ("D3", "O3", "Z3"):
            try:
                sheet.range(cell).api.NumberFormat = "@"
            except Exception as e:
                log_excel_exception(role, e, f"setting expiry cell format {cell}")
        write_hint_labels(sheet, role)
    except Exception as e:
        print(f"format_sheet: {e}", flush=True)
        log_excel_exception(role, e, "formatting worksheet")
    else:
        _sheet_fmt_done.add(key)


def force_usd_block_formats(sheet, block, n_rows):
    """Prices = 0.00, OI = text with $ only (kills ₹ number format)."""
    if n_rows <= 0:
        return
    c0 = block["start"]
    end = 4 + n_rows
    try:
        sheet_key = os.path.abspath(sheet.book.fullname).lower()
    except Exception:
        sheet_key = str(id(sheet))
    format_key = (sheet_key, c0, n_rows)
    if format_key in _usd_fmt_done:
        return
    try:
        for off in (0, 1, 2, 6, 7, 8):
            col = get_column_letter(c0 + off)
            sheet.range(f"{col}5:{col}{end}").api.NumberFormat = "0.00"
        sheet.range(f"{get_column_letter(c0 + 4)}5:{get_column_letter(c0 + 4)}{end}").api.NumberFormat = "0.0#"
        for off in (3, 5):
            col = get_column_letter(c0 + off)
            rng = sheet.range(f"{col}5:{col}{end}")
            rng.api.NumberFormat = "@"
            # rewrite numeric/₹ leftovers as $ text
            for r in range(5, end + 1):
                cell = sheet.range((r, c0 + off))
                v = cell.value
                if v is None or v == "":
                    continue
                if isinstance(v, (int, float)) or (isinstance(v, str) and ("₹" in v or v.startswith("$") is False and any(ch.isdigit() for ch in v))):
                    if isinstance(v, (int, float)):
                        cell.value = fmt_oi(v)
                    elif "₹" in str(v) or "Rs" in str(v):
                        cell.value = fmt_oi(v)
        _usd_fmt_done.add(format_key)
    except Exception as e:
        print(f"force_usd formats: {e}", flush=True)
        log_excel_exception("EXCEL", e, "formatting option-chain cells")


def _write_to_excel(role):
    try:
        import xlwings as xw
    except ImportError:
        show_error("Required component 'xlwings' is missing.")
        os._exit(1)
    path = ensure_excel_file(role)
    try:
        book = find_open_book(xw, path)
        if book is None:
            forget_workbook_formatting(path)
            apps_count = excel_apps_count(xw)
            status = (
                "excel_running_workbook_not_found"
                if apps_count > 0
                else "excel_not_found"
            )
            record_excel_connection(role, status, f"expected_workbook={path}")
            if apps_count > 0:
                return
            if excel_connected[role]:
                # Excel was connected and is now gone (closed by the customer,
                # crashed, or restarting after an add-in prompt). Keep waiting
                # for it to come back instead of quitting.
                #
                # This used to be `if excel_miss_counts[role] >= 10:
                # os._exit(0)`, which silently killed the EXE about 10 seconds
                # after Excel went away, so nothing rewrote the sheets even
                # after the customer reopened Excel. The miss counter is now
                # tracked for diagnostics only.
                excel_miss_counts[role] += 1
                return
            if excel_first_attempt[role] is None:
                excel_first_attempt[role] = time.time()
            elif time.time() - excel_first_attempt[role] > 90:
                show_error(
                    "Could not connect to Microsoft Excel.\n\n"
                    "Please run OptionChain.exe again."
                )
                os._exit(1)
            return
        excel_miss_counts[role] = 0
        excel_connected[role] = True
        try:
            excel_info = (
                f"; excel_version={book.app.api.Version}"
                f"; excel_operating_system={book.app.api.OperatingSystem}"
            )
            try:
                excel_details = excel_process_details(book.app.api.Hwnd)
                excel_info += (
                    f"; excel_process_bitness={excel_details['bitness']}"
                    f"; excel_process_is_administrator="
                    f"{excel_details['is_administrator']}"
                )
            except (OSError, TypeError, ValueError) as e:
                excel_info += (
                    f"; excel_process_is_administrator=unavailable:"
                    f"{type(e).__name__}:{e}"
                )
        except Exception as e:
            excel_info = f"; excel_details_unavailable={type(e).__name__}: {e}"
        record_excel_connection(role, "connected", f"workbook={path}{excel_info}")
        sheet = book.sheets[0]
        formatting_started = time.perf_counter()
        try:
            format_sheet(sheet, role)
        finally:
            _performance_metrics.record(
                role, "worksheet_setup", time.perf_counter() - formatting_started
            )

        global _refetch_now
        new_sym = read_symbol_cell(sheet, role)
        if new_sym and new_sym != role_symbol[role]:
            print(f"[{role}] symbol {role_symbol[role]} -> {new_sym}", flush=True)
            role_symbol[role] = new_sym
            _pending_date_loads.clear()
            _dropdown_sig.pop(new_sym, None)
            # New symbol means different rows and a different ladder, so the
            # payload and shape signatures from the old one must not suppress
            # the first write or the first recolour.
            _block_payload.clear()
            _block_shape.clear()
            _refetch_now = True
            ensure_sym_state(new_sym)
            for cell in ("D3", "O3", "Z3"):
                sheet.range(cell).value = ""
            for block in BLOCK_COLS:
                clear_start = get_column_letter(block["start"])
                clear_end = get_column_letter(block["start"] + 8)
                sheet.range(f"{clear_start}5:{clear_end}2000").clear()

        sym = role_symbol[role]
        ensure_sym_state(sym)

        v = sheet.range("A2").value
        if isinstance(v, str) and v.strip() and not v.strip().replace(".", "", 1).isdigit():
            sheet.range("A2").value = None
            v = None
        user_provided = isinstance(v, (int, float)) and v >= 1
        if not user_provided and isinstance(v, str):
            try:
                fv = float(v)
                if fv >= 1:
                    v = fv
                    user_provided = True
            except ValueError:
                pass

        try:
            update_user_dates(sym, sheet)
        except Exception as e:
            print(f"update_user_dates ({sym}): {e}", flush=True)
            log_excel_exception(role, e, "reading expiry input cells")
        try:
            update_expiry_dropdowns(sym, sheet)
        except Exception as e:
            _dropdown_sig.pop(sym, None)
            print(f"update_expiry_dropdowns ({sym}): {e}", flush=True)
            log_excel_exception(role, e, "updating expiry dropdowns")

        sp = spot_cmp[sym] or 0
        MAX_A2 = 100
        if not user_provided:
            range_val = MAX_A2  # empty A2 → all (API returns ≤100 around ATM)
        else:
            desired = int(v)
            if desired > MAX_A2:
                desired = MAX_A2
                sheet.range("A2").value = desired
            range_val = max(desired, 1)

        for bi, block in enumerate(BLOCK_COLS):
            expiry = active_expiries[sym][bi]
            col_l = block["start"]
            if not expiry:
                sheet.range((1, col_l)).value = f"{sym} Spot: --"
                sheet.range((1, col_l + 4)).value = "PCR: --"
                continue

            chain_started = time.perf_counter()
            chain = None
            try:
                chain, pcr = build_option_chain(sym, expiry, range_val)
            finally:
                _performance_metrics.record(
                    role,
                    "chain_build",
                    time.perf_counter() - chain_started,
                    rows=len(chain) if chain is not None else 0,
                )
            cnt = len(chain) if chain is not None else 0
            sheet.range((1, col_l)).value = f"{sym} Spot: {sp:.2f} | {cnt} strikes"
            sheet.range((1, col_l + 4)).value = f"PCR: {pcr:.2f}"
            net_paint(sheet, role, bi, col_l)

            if chain is not None and not chain.empty:
                # set OI as text BEFORE write so Excel cannot coerce $223 -> INR number
                n_rows = len(chain)
                c0 = block["start"]
                for off in (3, 5):
                    try:
                        sheet.range((5, c0 + off), (4 + n_rows, c0 + off)).api.NumberFormat = "@"
                    except Exception as e:
                        log_excel_exception(
                            role, e, f"formatting open-interest column {c0 + off}"
                        )

                write_started = time.perf_counter()
                try:
                    # Only write when the contents actually differ. Any data
                    # change alters the signature and forces the write, so live
                    # updates still land; unchanged blocks cost nothing.
                    payload_sig = hash(tuple(map(tuple, chain.values.tolist())))
                    if payload_sig != _block_payload.get((role, bi)):
                        _block_payload[(role, bi)] = payload_sig
                        sheet.range((5, block["start"])).options(index=False, header=False).value = chain
                    else:
                        _performance_metrics.record(role, "excel_range_skipped", 0.0, rows=n_rows)
                finally:
                    _performance_metrics.record(
                        role,
                        "excel_range_write",
                        time.perf_counter() - write_started,
                        rows=n_rows,
                    )

                # clear stale rows below (fixes A2 shrink leaving old data)
                if n_rows < 200:
                    stale_a = f"{get_column_letter(c0)}{5 + n_rows}:{get_column_letter(c0 + 8)}2000"
                    try:
                        sheet.range(stale_a).clear()
                    except Exception as e:
                        log_excel_exception(role, e, f"clearing stale rows in block {bi + 1}")

                formatting_started = time.perf_counter()
                try:
                    # ATM row and ITM ranges depend only on which strikes are
                    # present, not on their prices. Skip when the ladder is
                    # unchanged; force_usd_block_formats is already self-gating.
                    strikes = chain["strike"].tolist()
                    shape = (n_rows, strikes[0] if strikes else 0,
                             strikes[-1] if strikes else 0)
                    if _block_shape.get((role, bi)) != shape:
                        _block_shape[(role, bi)] = shape

                        force_usd_block_formats(sheet, block, n_rows)

                        atm_strike = min(strikes, key=lambda s: abs(s - sp))
                        atm_row = strikes.index(atm_strike) + 5

                        atm_range = f"{block['atm']}{atm_row}:{block['put_end']}{atm_row}"
                        sheet.range(atm_range).color = (255, 255, 0)
                        sheet.range(atm_range).font.bold = True

                        for cell_range, color in itm_color_ranges(block, strikes, sp):
                            sheet.range(cell_range).color = color
                finally:
                    _performance_metrics.record(
                        role,
                        "excel_formatting",
                        time.perf_counter() - formatting_started,
                        rows=n_rows,
                    )
            else:
                clear_start = get_column_letter(block["start"])
                clear_end = get_column_letter(block["start"] + 8)
                sheet.range(f"{clear_start}5:{clear_end}2000").clear()
    except Exception as e:
        print(f"Excel error ({role} / {role_symbol.get(role)}): {e}", flush=True)
        record_excel_connection(role, "excel_com_error", f"{type(e).__name__}: {e}")
        log_excel_exception(role, e)


def write_to_excel(role):
    started = time.perf_counter()
    try:
        _write_to_excel(role)
    finally:
        _performance_metrics.record(
            role, "excel_pass", time.perf_counter() - started
        )


def print_chain(sym):
    sp = spot_cmp[sym]
    for bi, expiry in enumerate(active_expiries[sym]):
        if not expiry:
            continue
        chain, pcr = build_option_chain(sym, expiry, 999)
        if chain is None or chain.empty:
            continue
        sentiment = "BULLISH" if pcr > 1.0 else "BEARISH"
        print(f"{sym} Block {bi + 1} | {expiry} | Spot: {sp:.2f} | PCR: {pcr:.2f} ({sentiment}) | {len(chain)} strikes")
        header = "  ".join(f"{c:>10}" for c in CHAIN_COLS)
        print(header)
        print("-" * len(header))
        for _, row in chain.iterrows():
            parts = []
            for c in CHAIN_COLS:
                v = row.get(c, 0)
                if c in ("Call OI", "Put OI", "strike"):
                    parts.append(str(v).rjust(10))
                else:
                    parts.append(f"{float(v):>10.2f}")
            print("  ".join(parts))
        print()


async def display_loop():
    is_windows = platform.system() == "Windows"
    license_tick = 0
    global _excel_gone_logged
    _excel_gone_logged = False
    while True:
        render_started = time.perf_counter()
        # re-check license roughly every ~60s while running
        license_tick += 1
        if license_tick >= max(1, int(60 / max(REFRESH, 0.05))):
            license_tick = 0
            ok, _vt = license_is_valid()
            if not ok:
                if not ensure_license():
                    os._exit(1)
        if is_windows:
            # A dead/crashed Excel must never stop the chain. Each role is
            # isolated, and while Excel is gone we skip the COM pass entirely
            # and wait for the customer to reopen it. Previously the RPC failure
            # propagated and killed the EXE, after which nothing wrote at all.
            excel_alive = excel_process_running()
            for role in ROLE_ORDER:
                if not excel_alive:
                    break
                try:
                    write_to_excel(role)
                except Exception as e:
                    log_excel_exception(role, e, "render pass")
            if not excel_alive and not _excel_gone_logged:
                _excel_gone_logged = True
                print("Excel is not running - pausing writes until it is reopened.", flush=True)
                _last_excel_exception.clear()
            elif excel_alive and _excel_gone_logged:
                # Excel came back. Everything cached about the sheets is now
                # wrong: the workbooks are freshly opened and their cells are
                # empty, so the payload/shape/format caches must be dropped or
                # the app would treat the blank sheets as "already written" and
                # leave them blank.
                _excel_gone_logged = False
                _block_payload.clear()
                _block_shape.clear()
                _sheet_fmt_done.clear()
                _usd_fmt_done.clear()
                _dropdown_sig.clear()
                print("Excel reopened - rewriting all sheets from scratch.", flush=True)
            render_seconds = time.perf_counter() - render_started
            _performance_metrics.record("ALL", "render_cycle", render_seconds)
            if render_seconds > REFRESH:
                _performance_metrics.record(
                    "ALL", "refresh_overrun", render_seconds - REFRESH
                )
            await flush_performance_metrics()
        else:
            os.system("clear")
            for sym in active_symbols():
                print_chain(sym)
            print(f"Symbols {','.join(active_symbols())} | poll {POLL_GAP}s | Ctrl+C stop", flush=True)

        # Sleep AFTER the work, for at least as long as the work took.
        #
        # The customer builds their strategy in the same Excel instance, so if
        # writing the chains leaves no idle gap, Excel is busy whenever they try
        # to click a cell and the workbook feels stuck. Measured with a heavy
        # 800-row formula workbook open, one pass took ~880ms; against a 1.0s
        # refresh that is ~88% duty cycle. Backing off to an equal amount of
        # idle time keeps the workbook usable and lets the refresh stretch
        # automatically when their workbook is heavy, while a light workbook
        # still updates at the configured interval.
        await asyncio.sleep(max(float(REFRESH), time.perf_counter() - render_started))


async def flush_performance_metrics():
    global _last_performance_flush, _performance_log_error_shown
    now = time.monotonic()
    if now - _last_performance_flush < 30:
        return
    lines = _performance_metrics.summary_lines(now)
    try:
        await asyncio.to_thread(append_performance_summary, lines)
    except OSError as e:
        if not _performance_log_error_shown:
            show_error(
                "Could not append performance diagnostics to startup_diag.log.\n\n"
                "Check that your user profile is writable and has free space.\n\n"
                f"Details: {e}"
            )
            _performance_log_error_shown = True
    else:
        _performance_metrics.reset(now)
        _last_performance_flush = now


async def main_async():
    for sym in active_symbols():
        ensure_sym_state(sym)
    open_excel_files()
    errs = []
    for sym in active_symbols():
        try:
            await init_symbol(sym)
        except Exception as e:
            errs.append(f"{sym}: {e}")
    if errs and len(errs) >= len(active_symbols()):
        show_error(
            "Could not reach option data sources (BigClawd / CBOE).\n\n"
            "Check internet connection and run again.\n\n"
            + "\n".join(errs)
        )
        return
    await asyncio.gather(
        data_poller(),
        expiry_roller(),
        process_pending_dates(),
        display_loop(),
    )


def main():
    try:
        write_startup_diagnostics()
    except OSError as e:
        show_error(f"Could not write startup diagnostics:\n{e}")
    try:
        if not ensure_license():
            os._exit(1)
        asyncio.run(main_async())
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
