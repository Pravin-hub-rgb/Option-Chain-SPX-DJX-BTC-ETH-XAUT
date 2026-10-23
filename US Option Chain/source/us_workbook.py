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
import threading
import platform
import shutil

from us_diagnostics import append_app_exception, feed_log
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

# How many expiry blocks a sheet shows. The customer's strategy sheet reads
# blocks 1+2 only (near and next expiry), so block 3 is retired: the render
# loop clears it once per session and never touches those columns again,
# which cuts a third of the per-cycle Excel traffic. The layout constants
# above keep their third entry so the columns are still here if the customer
# ever asks for a third expiry back.
BLOCKS_PER_SHEET = 2

# Per-block expiry input cells, one per SHOWN block (D3 = block 1, O3 = block 2).
EXPIRY_CELLS = ("D3", "O3", "Z3")[:BLOCKS_PER_SHEET]


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
    for block in BLOCK_COLS[:BLOCKS_PER_SHEET]:
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
    for si, block in enumerate(BLOCK_COLS[:BLOCKS_PER_SHEET], 1):
        if _safe_set(ws, 3, block["start"], f"Expiry {si}"):
            changed = True
        if _safe_set(ws, 3, block["start"] + 2, f"(edit {EXPIRY_CELLS[si - 1]})"):
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


def short_path(path):
    """The Windows 8.3 short name for a path, when there is one.

    Excel's Workbooks.Open stops working somewhere past 128 characters and
    never returns. A package unzipped under a user's Downloads folder is
    over that easily, and the symptom is the worst kind: the workbook never
    opens, the tool says "Excel is not running" while Excel is sitting right
    there, and the sheet stays empty. That is what "it does not work" looks
    like on the customer's machine.

    The short name is the same file, so it opens in a fraction of a second.
    """
    try:
        import ctypes

        buffer = ctypes.create_unicode_buffer(1024)
        if ctypes.windll.kernel32.GetShortPathNameW(path, buffer, 1024):
            return buffer.value or path
    except Exception:
        pass
    return path


AUTO_OPEN_TIMEOUT_SECONDS = 45.0


def open_excel_files():
    """Open the three workbooks in ONE Excel, and get on with the data.

    Two things were wrong with the old way of doing this.

    It went through the Windows shell, which starts a SEPARATE Excel for
    every workbook it is given. Three workbooks meant three Excel processes,
    two of them quit, and the one that survived was never registered where
    the tool can reach it. The tool then said "Excel is not running" while
    Excel was sitting right there, and wrote nothing.

    And it opened them through that same shell, so a package under a long
    path never opened at all. Excel's Workbooks.Open gives up somewhere past
    128 characters and never returns.

    So: one Excel, opened through the API the tool writes with, using the
    8.3 short name for a long path. Opening Excel can take a while on a
    cold or busy machine, and the chain does not need Excel to be up first,
    so the open runs on its own thread and the render loop picks the
    workbooks up as soon as they are there.

    The open pass is guarded now. Workbooks.Open aimed at an Excel this
    process does not own - the Delta tool's instance, or one whose
    controlling process died earlier - can block forever, because a busy
    Excel answers automation by simply not answering. A dead-quiet daemon
    thread and a blank sheet was the exact customer symptom, so a watchdog
    times the pass out, logs it, and retries the missing workbooks in a
    fresh Excel instance of our own.
    """
    if platform.system() != "Windows":
        return
    finished = threading.Event()
    threading.Thread(target=_open_excel_files, args=(finished, False), daemon=True).start()

    def _watchdog():
        if finished.wait(AUTO_OPEN_TIMEOUT_SECONDS):
            return
        feed_log(
            "excel_auto_open_timeout",
            note=(f"chain workbooks still not open after {int(AUTO_OPEN_TIMEOUT_SECONDS)}s; "
                  "retrying once in a separate Excel instance"),
        )
        retry_done = threading.Event()
        threading.Thread(target=_open_excel_files, args=(retry_done, True), daemon=True).start()

        def _retry_watchdog():
            if not retry_done.wait(AUTO_OPEN_TIMEOUT_SECONDS):
                feed_log(
                    "excel_auto_open_retry_timeout",
                    note="the retry Excel instance also did not finish; waiting for the customer to open the workbooks",
                )
        threading.Thread(target=_retry_watchdog, daemon=True).start()

    threading.Thread(target=_watchdog, daemon=True).start()


def _open_excel_files(finished=None, own_instance=False):
    try:
        import pythoncom

        pythoncom.CoInitialize()
    except Exception:
        pythoncom = None
    app = None
    started_excel = False
    try:
        import xlwings as xw

        try:
            apps = list(xw.apps)
        except Exception:
            apps = []
        try:
            if own_instance:
                # The first pass may still be wedged inside Workbooks.Open on
                # some busy or orphaned Excel; the retry always creates a
                # fresh instance of our own so the books land somewhere we
                # can definitely reach.
                started_excel = True
                app = xw.App(visible=True)
            else:
                started_excel = not apps
                app = apps[0] if apps else xw.App(visible=True)
        except Exception as e:
            print(f"Could not start Excel: {e}", flush=True)
            if finished is not None:
                finished.set()
            return
        feed_log("excel_auto_open_started", apps=len(apps), own_instance=own_instance)
        for role in ROLE_ORDER:
            path = excel_path(role)
            try:
                ensure_excel_file(role)
            except Exception as e:
                print(f"ensure_excel_file failed ({role}): {e}", flush=True)
            try:
                # Lazy import: us_excel imports this module, so importing it
                # at the top would be a circle.
                from us_excel import find_open_book

                if find_open_book(xw, path) is not None:
                    # Already open - the customer opened it themselves, so the
                    # tool must never open it again behind their back.
                    _auto_open_done[role] = True
                    feed_log("excel_auto_open_done", role=role, via="already_open")
                    continue
            except Exception:
                pass
            target = path if len(path) <= 120 else short_path(path)
            try:
                app.api.Workbooks.Open(target)
                _auto_open_done[role] = True
                feed_log("excel_auto_open_done", role=role, via="api")
            except Exception as e:
                print(f"Could not open {ROLE_FILES[role]}: {e}", flush=True)
                feed_log("excel_auto_open_failed", role=role, error=str(e))
                show_error(
                    f"Could not open {ROLE_FILES[role]}:\n\n{e}\n\n"
                    "If the package was unzipped into a long path, move it "
                    "to a short one - C:\\OptionChain - and run the tool again."
                )
        if finished is not None:
            finished.set()
    except Exception as e:
        print(f"Could not open the workbooks: {type(e).__name__}: {e}", flush=True)
        if finished is not None:
            finished.set()
    finally:
        if started_excel:
            _close_empty_book1(app)
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _close_empty_book1(app):
    """Close the empty Book1 Excel always starts with.

    Only when this process started Excel, and only when the book is empty:
    a customer who has their own Book1 open with work in it must never have
    it closed underneath them.
    """
    try:
        for book in list(app.books):
            if book.name.strip().lower() != "book1":
                continue
            try:
                if book.sheets[0].range("A1").value is None:
                    book.close()
            except Exception:
                continue
    except Exception:
        pass
