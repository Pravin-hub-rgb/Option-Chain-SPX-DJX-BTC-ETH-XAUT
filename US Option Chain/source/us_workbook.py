"""Workbook file layer: symbols-in-sheets helpers, templates, open/ensure.

Everything here talks about the workbook FILES on disk (create, repair,
rebuild from scratch, open in Excel) plus the small text helpers that decide
what goes into a sheet cell. It does not touch Excel's COM file handles:
those live later in the render/instance modules.

Moved verbatim from main.py. main.py imports the names it still needs back,
so call sites keep working. open_excel_files resolves find_open_book from
main with a late import to stay import-safe.
"""

import math
import os
import platform
import shutil

from us_diagnostics import append_app_exception
from us_licensing import show_error
from us_state import (
    _auto_open_done,
    _template_ready,
    BASE_DIR,
    DEFAULT_SYMBOL,
    LOCKED_ROLES,
    role_symbol,
    ROLE_FILES,
    ROLE_ORDER,
)

# US index option roots — CBOE delayed needs "_" prefix; BigClawd partial support
INDEX_SYMBOLS = {
    "SPX", "SPXW", "XSP", "DJX", "DJXW", "NDX", "NDXP",
    "VIX", "VIXW", "RUT", "RUTW", "OEX", "XEO",
}

BLOCK_COLS = [
    {"start": 1,  "atm": "A", "call_start": "A", "call_end": "D", "put_start": "F", "put_end": "I"},
    {"start": 12, "atm": "L", "call_start": "L", "call_end": "O", "put_start": "Q", "put_end": "T"},
    {"start": 23, "atm": "W", "call_start": "W", "call_end": "Z", "put_start": "AB", "put_end": "AE"},
]


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


def get_column_letter(col):
    result = ""
    while col > 0:
        col -= 1
        result = chr(65 + col % 26) + result
        col //= 26
    return result


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


def excel_path(role):
    return os.path.join(BASE_DIR, ROLE_FILES[role])


def default_k1(role):
    return role_symbol.get(role) or DEFAULT_SYMBOL


def is_locked_role(role):
    return role in LOCKED_ROLES


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
    import xlwings as xw
    for role in ROLE_ORDER:
        path = excel_path(role)
        try:
            os.remove(path + ":Zone.Identifier")
        except OSError:
            pass
        # If the workbook is already open in Excel, don't open it again.
        # Opening a second copy with the same name throws "same name" errors,
        # and openpyxl cannot load a file Excel has locked anyway.
        try:
            import main as _main
            book = _main.find_open_book(xw, path)
            if book is not None:
                # Already open - the customer opened it themselves, so the
                # tool must never open it again behind their back.
                _auto_open_done[role] = True
                continue
        except Exception:
            pass
        # Workbook not open — ensure the file exists, then open it.
        try:
            ensure_excel_file(role)
        except Exception as e:
            print(f"ensure_excel_file failed ({role}): {e}", flush=True)
        try:
            os.startfile(path)
            # We have now opened it. If the customer closes it later, it stays
            # closed.
            _auto_open_done[role] = True
        except OSError as e:
            show_error(f"Could not open {ROLE_FILES[role]}:\n{e}")
            try:
                append_app_exception(f"opening {role} workbook", e)
            except OSError as log_error:
                print(f"Could not log workbook open failure ({role}): {log_error}", flush=True)
