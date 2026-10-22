"""
US Stock / S&P 500 Option Chain Tool
Primary: Databento OPRA live when an API key is configured.
Fallback: BigClawd free brokerage chain then CBOE delayed quotes.
Three Excel files auto-open (Delta-style):
  spx_option_chain.xlsx       -> locked to SPX (S&P500 index)
  djx_option_chain.xlsx       -> locked to DJX (Dow index)
  us_stock_option_chain.xlsx  -> editable K1 (default AAPL)
Shared cache by symbol; one process drives all books.
"""

import asyncio
import atexit
import os
import platform
import re
import shutil
import subprocess
import threading
import time
from collections import defaultdict
from datetime import date, datetime

import requests

import config
import calc_mode_guard
import databento_feed
from us_diagnostics import (
    append_app_exception,
    append_connection_event,
    append_runtime_warning,
    excel_process_details,
    append_performance_summary,
    build_id,
    write_startup_diagnostics,
    feed_log,
)
from us_performance import PerformanceMetrics, color_signature, itm_color_ranges
from product_identity import identify_product_directory
from excel_helpers import (
    find_book_and_app_in_apps,
    find_book_by_name_in_apps,
    prefer_expiry_root,
    refresh_delay,
    reopen_due,
    should_restore_calculation,
    workbook_name_is_open_in_apps,
    workbook_name_open_anywhere_in_apps,
)
from us_time import eastern_now, filter_expiry_dates

import us_state
from us_licensing import ensure_license, license_is_valid, save_license, show_error, show_info
from us_workbook import (
    _rebuild_fresh, _safe_set, _unmerge_all, _write_template, base_root,
    clean_symbol, default_k1, ensure_excel_file, excel_path, get_column_letter,
    is_index_symbol, is_locked_role, normalize_symbol, open_excel_files,
    BLOCK_COLS, INDEX_SYMBOLS,
)
from us_spot import (
    PARITY_SANITY_PCT, PARITY_TOLERANCE_PCT, PARITY_TOLERANCE_POINTS,
    SPOT_POLL_SECONDS, SPOT_STALE_SECONDS,
    _databento_spot, _parity_log_at, _parity_spot, _parity_spot_cache,
    _spot_divergence, _yahoo_spot_fetch, note_spot, refresh_tencent_spot_cache,
    source_delay_note, spot_source_for, tencent_spot, yahoo_spot,
)
from us_chain import (
    _header_at, _set_header_at, _short_build, active_symbols, build_option_chain,
    available_expiries, data_poller, ensure_sym_state, expiry_roller,
    fmt_oi, init_symbol, iso_to_user, parse_user_date, process_pending_dates,
    read_symbol_cell, spot_poller,
    CHAIN_COLS,
)
from us_state import (
    _FIRST_SIGHT_BOUND, _UNPROTECTED_WRITE_INTERVAL, _app_cache,
    _auto_open_done, _book_sheet_cache, _bulk_will_recalc, _calc_manual_failed,
    _calc_manual_notice_shown, _calc_notice_shown, _calc_session,
    _databento_feed, _dropdown_sig, _excel_connection_status,
    _excel_liveness_cache, _feed_access, _last_bulk_restore,
    _last_excel_exception, _last_performance_flush, _last_recalc_ms,
    _last_source_call, _last_symbol_call, _last_rows, _manual_since,
    _net, _net_drawn, _pending_date_loads, _performance_log_error_shown,
    _performance_metrics, _rate_lock, _screen_blocked_until,
    _screen_failures, _SCREEN_BACKOFF_SECONDS, _SCREEN_FAILURE_LIMIT,
    _sheet_altered, _sheet_cache, _speed_log_at, _spot_cache,
    _spot_lock, _template_ready, active_expiries,
    all_known_expiries, BASE_DIR, cfg, DEFAULT_SYMBOL, DJX_SYMBOL,
    EXPIRY_REFRESH, live_data, LOCKED_ROLES, MANUAL_WATCHDOG_SECONDS,
    net_badge, net_note,
    NET_FAIL_LIMIT, oi_roots, expiry_roots, POLL_GAP, PRODUCT_EXE, PRODUCT_ID,
    REFRESH, role_symbol, ROLE_FILES, ROLE_ORDER, ROUND_PAUSE,
    SOURCE_MIN_INTERVAL, SPX_SYMBOL, spot_cmp, SYMBOL_CELL,
    SYMBOL_POLL_SECONDS, user_pinned,
)

# --- connection health indicator ------------------------------------------
# Market/API ghatne par sheet freeze ho jata hai; user ko pata chalna chahiye
# ki data stale hai. NET_FAIL_LIMIT consecutive failures ke baad header me red
# tag dikhta hai, aur connection wapas aate hi apne aap hat jaata hai.


def net_paint(sheet, role, bi, col_l):
    """Tag cell (row 1, block ka last column) likho â€” sirf jab badle."""
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


    """True when the sheet still has the layout the tool expects.

    The tool writes to fixed columns. If the customer inserts or deletes a
    column, those positions no longer line up and the tool would write over
    their own data. A blank cell is fine - the tool is about to fill it. Only
    a cell holding something that is not our header means the sheet was
    restructured.
    """
    try:
        a4 = sheet.range((4, 1)).value
    except Exception:
        return True
    if a4 is None or str(a4).strip() == "":
        return True
    return str(a4).strip() == "Call LTP"


excel_connected = {r: False for r in ROLE_ORDER}
excel_miss_counts = {r: 0 for r in ROLE_ORDER}
excel_first_attempt = {r: None for r in ROLE_ORDER}
_excel_startup_notice_shown = {r: False for r in ROLE_ORDER}
_reopen_after = {r: 0.0 for r in ROLE_ORDER}
# Set once the tool has asked Excel to open a workbook. After that it never
# asks again: a customer who closes the workbook must not watch it pop back
# up. Re-armed only if the file itself disappears, so a deleted workbook can
# still be recreated and opened.


_last_window_status = {}
def sheet_structure_ok(sheet, role):
    """True when the sheet still has the layout the tool expects.

    The tool writes to fixed columns. If the customer inserts or deletes a
    column, those positions no longer line up and the tool would write over
    their own data. A blank cell is fine - the tool is about to fill it. Only
    a cell holding something that is not our header means the sheet was
    restructured.
    """
    try:
        a4 = sheet.range((4, 1)).value
    except Exception:
        return True
    if a4 is None or str(a4).strip() == "":
        return True
    return str(a4).strip() == "Call LTP"


def _contiguous_runs(indices):
    runs = []
    start = prev = None
    for i in indices:
        if start is None:
            start = prev = i
        elif i == prev + 1:
            prev = i
        else:
            runs.append((start, prev))
            start = prev = i
    if start is not None:
        runs.append((start, prev))
    return runs


def show_status(book, role, status, detail=""):
    """Put the tool's health where the user can see it without asking for anything.

    The customer is not technical and does not read logs. Support works by
    screen-sharing, so the one thing guaranteed to be on screen is this
    workbook's own title bar. It carries the build and a plain-language status,
    which is enough to tell a working tool from a broken one and to name the
    build, without the customer having to do anything at all.

    Only written when the text actually changes, so it costs nothing per cycle.

    The clock in the title is deliberately coarse. Rewriting the window caption
    makes Excel repaint the title bar, and on a large workbook that repaint
    lands on top of whatever the customer is doing - scrolling included - and
    shows up as a stutter. A five second clock is still useful and does not
    repaint the window on every refresh.
    """
    now = time.monotonic()
    key = id(book)
    last = _last_window_status.get(key)
    if last is not None:
        last_text, last_at = last
        if last_text.split("  ")[0] == f"{role.upper()}  |  Option Chain {_short_build()}  |  {status}":
            # Same role/build/status: only refresh the clock every few seconds.
            if now - last_at < 5.0:
                return
    stamp = time.strftime("%H:%M:%S")
    title = f"{role.upper()}  |  Option Chain {_short_build()}  |  {status}  {stamp}"
    # Show how long the customer's own workbook takes to recalculate after we
    # touch it. This is the number behind "the strategy feels slow", so it
    # belongs where support can see it during a screen-share.
    if _last_recalc_ms >= 50.0:
        title += f"  |  recalc {_last_recalc_ms:.0f}ms"
    if last is not None and last[0] == title:
        return
    _last_window_status[key] = (title, now)
    try:
        book.api.Windows(1).Caption = title
    except Exception as e:
        log_excel_exception(role, e, "setting window status")


def excel_apps_count(xw):
    try:
        return len(list(xw.apps))
    except Exception as e:
        log_excel_exception("EXCEL", e, "enumerating Excel applications")
        return -1


def excel_process_running(max_age=1.0):
    """True if any EXCEL.EXE is alive, checked without touching COM.

    If Excel has crashed or been killed, the cached COM server is a dead
    reference and the next `xw.apps` / `app.books` call raises "RPC server is
    unavailable" and can take the whole process down with it. Checking the
    process table first means we simply stop writing and wait for the customer
    to bring Excel back, instead of dying silently.

    The answer is cached briefly. Spawning tasklist per workbook per cycle
    added a large fraction of a second to every render pass, which is a lot to
    pay for a check that only needs to be approximately current.
    """
    global _excel_liveness_cache
    now = time.monotonic()
    if _excel_liveness_cache is not None:
        stamp, running = _excel_liveness_cache
        if now - stamp <= max_age:
            return running
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
        _excel_liveness_cache = (now, True)
        return True
    running = "EXCEL.EXE" in (result.stdout or "").upper()
    _excel_liveness_cache = (now, running)
    return running


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
    book, _app = find_book_and_app_in_apps(
        apps,
        target,
        lambda exc, context: log_excel_exception(role, exc, context),
    )
    if book is not None:
        return book
    # The customer may already have this workbook open from a different folder
    # (an old install, a copy they made, the source tree). Excel will not open
    # a second workbook with the same name, so opening our own copy fails and
    # leaves the tool unable to write anything. Use the copy that is already
    # open â€” it is the same product, and it is the one on their screen.
    book, _app = find_book_by_name_in_apps(
        apps,
        fname,
        lambda exc, context: log_excel_exception(role, exc, context),
    )
    if book is not None:
        try:
            record_excel_connection(
                role,
                "attached_to_open_copy_elsewhere",
                f"using={book.fullname} expected={os.path.abspath(fname)}",
            )
            feed_log(
                "workbook_attached_elsewhere",
                role=role,
                using=book.fullname,
                expected=os.path.abspath(fname),
            )
        except Exception:
            pass
    return book


def any_book_named(xw, fname):
    """True if the exact workbook path is open in any reachable Excel app.

    If an Excel instance cannot be inspected, defer reopening for this cycle
    rather than risk opening a duplicate workbook.
    """
    try:
        apps = list(xw.apps)
    except Exception as e:
        log_excel_exception("EXCEL", e, "enumerating Excel applications before reopen")
        return True
    return workbook_name_is_open_in_apps(
        apps,
        fname,
        lambda exc, context: log_excel_exception("EXCEL", exc, context),
    )


def update_expiry_dropdowns(sym, sheet):
    # Date order: the chain the customer wants is almost always the next
    # expiry out, and their own broker lists them nearest first.
    dates = available_expiries(all_known_expiries[sym], sym, by_date=True)
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
    # xlwings Sheet has no .cell() â€” write helper column as one range
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
            # text â€” stops Excel date conversion / #####
            # Setting the format dirties the cell, so only do it once per
            # sheet instead of on every refresh.
            fmt_key = (id(sheet), cell)
            if fmt_key not in _user_fmt_done:
                rng.api.NumberFormat = "@"
                _user_fmt_done.add(fmt_key)
        except Exception:
            pass
        val = rng.value
        # Excel may have converted text into a datetime â€” force back to DD/MM/YYYY text
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
            # same expiry but odd format â€” normalize
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
# Last written values per (role, block), so only the rows that actually moved
# are written back. Writing the whole block dirties every cell in it, and
# Excel repaints the dirty region when the write window closes - that is what
# made scrolling stutter.
_block_values = {}
# Cells whose number format has already been set. Setting a number format
# dirties the cell, and doing that on every refresh is wasted work.
_oi_fmt_done = set()
_user_fmt_done = set()
# Last applied ITM/ATM colouring per (role, block). Used to skip repainting
# when the picture has not changed, so a live price tick does not dirty every
# cell of the block and make Excel recalculate the customer's own formulas.
_colour_state = {}
# Rows this block used on its previous cycle, so a shrink can clear exactly
# what was there instead of wiping a fixed block of rows every tick.


def _screen_pause_active():
    """True while the screen-updating attempt is paused after failures."""
    return _screen_blocked_until > time.monotonic()


def _note_screen_ok():
    global _screen_failures, _screen_blocked_until
    _screen_failures = 0
    _screen_blocked_until = 0.0


def _note_screen_failed(exc):
    global _screen_failures, _screen_blocked_until
    _screen_failures += 1
    if _screen_failures >= _SCREEN_FAILURE_LIMIT:
        _screen_blocked_until = time.monotonic() + _SCREEN_BACKOFF_SECONDS
        feed_log(
            "screen_updating_paused",
            failures=_screen_failures,
            seconds=_SCREEN_BACKOFF_SECONDS,
            error=type(exc).__name__,
        )


def forget_workbook_formatting(path):
    normalized_path = os.path.abspath(path).lower()
    _sheet_fmt_done.discard(normalized_path)
    stale_keys = {
        key for key in _usd_fmt_done
        if isinstance(key, tuple) and key[0] == normalized_path
    }
    _usd_fmt_done.difference_update(stale_keys)


# Labels written every run (cheap) so user always sees changeable hints
# J1 is role-dependent â€” written inside write_hint_labels
HINT_LABELS = (
    ("B2", "Â± strikes â€” edit A2"),
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
    # keep Symbol label from squishing (cheap â€” every run)
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
    widths[9] = 22    # J Symbol label â€” full text
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
        # Locked files (SPX/DJX): K1 gray + forced â€” not an input
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
        # Stock role: K1 is user-editable â€” never write its value here.
        # The initial value is set by _write_template when the file is created.
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
    """Prices = 0.00, OI = text with $ only (kills â‚¹ number format).

    The OI columns are already written as "$123" text by build_option_chain, so
    the only thing needed here is the text number format that stops Excel
    coercing them back into a number.

    The previous version also walked the OI column one cell at a time to
    convert leftovers. That was 1,206 COM round trips per block â€” about twelve
    seconds for a 600-row block â€” and it ran again every time the row count
    changed. It is now one bulk read and one bulk write, once per sheet, and
    the per-cycle cost is a single range write.
    """
    if n_rows <= 0:
        return
    c0 = block["start"]
    end = 4 + n_rows
    try:
        sheet_key = os.path.abspath(sheet.book.fullname).lower()
    except Exception:
        sheet_key = str(id(sheet))
    fmt_key = (sheet_key, c0, "numfmt")
    if fmt_key not in _usd_fmt_done:
        try:
            for off in (0, 1, 2, 6, 7, 8):
                col = get_column_letter(c0 + off)
                sheet.range(f"{col}5:{col}{end}").api.NumberFormat = "0.00"
            strike_col = get_column_letter(c0 + 4)
            sheet.range(f"{strike_col}5:{strike_col}{end}").api.NumberFormat = "0.0#"
            for off in (3, 5):
                col = get_column_letter(c0 + off)
                sheet.range(f"{col}5:{col}{end}").api.NumberFormat = "@"
            _usd_fmt_done.add(fmt_key)
        except Exception as e:
            print(f"force_usd formats: {e}", flush=True)
            log_excel_exception("EXCEL", e, "formatting option-chain cells")
            return
        # One-off bulk cleanup of values left as numbers or â‚¹ by an older
        # build. One read and one write for the whole sheet, not per cell.
        for off in (3, 5):
            col = get_column_letter(c0 + off)
            try:
                rng = sheet.range(f"{col}5:{col}{end}")
                values = rng.value
            except Exception as e:
                log_excel_exception("EXCEL", e, f"reading open-interest column {c0 + off}")
                continue
            if not values:
                continue
            changed = False
            for i, v in enumerate(values):
                if isinstance(v, (int, float)):
                    values[i] = fmt_oi(v)
                    changed = True
                elif isinstance(v, str) and ("â‚¹" in v or "Rs" in v):
                    values[i] = fmt_oi(v)
                    changed = True
            if changed:
                try:
                    rng.value = values
                except Exception as e:
                    log_excel_exception("EXCEL", e, f"writing open-interest column {c0 + off}")


_workbook_missing = {}  # path -> [first_seen_monotonic, already_notified]


def notice_workbook_missing(fname):
    """Tell the user once when Excel is open but this workbook is not.

    Without this the tool simply waits, and since it runs without a console the
    customer sees nothing at all and assumes it has stopped working.

    Tracked per workbook: all three are checked every cycle, so a global timer
    would be reset by whichever books are still connected and the missing one
    would never be reported.
    """
    now = time.monotonic()
    entry = _workbook_missing.get(fname)
    if entry is None:
        _workbook_missing[fname] = [now, False]
        return
    # Excel may still be opening the workbooks it was just asked to open, so
    # give it a moment before telling the user anything is wrong.
    if entry[1] or now - entry[0] < 20:
        return
    entry[1] = True
    try:
        show_error(
            "Microsoft Excel is open, but this workbook is not:\n\n"
            f"{os.path.basename(fname)}\n\n"
            "Expected in:\n"
            f"{os.path.dirname(fname)}\n\n"
            "Open it from there and the tool will start writing again.\n\n"
            "If a copy of the same file is open in another folder, close that "
            "one first - Excel will not open two workbooks with the same name."
        )
    except Exception as e:
        print(f"Could not show workbook-missing notice: {e}", flush=True)


def clear_workbook_missing_notice(fname):
    _workbook_missing.pop(fname, None)


def _write_to_excel(role):
    try:
        import xlwings as xw
    except ImportError:
        show_error("Required component 'xlwings' is missing.")
        os._exit(1)
    path = ensure_excel_file(role)
    try:
        # Cached COM book/sheet objects â€” avoids enumerating all apps + all
        # books on every cycle, which dominated the render loop. Dropped on
        # staleness so a closed/renamed workbook recovers.
        book = _book_sheet_cache.get(role)
        cached_sheet = _sheet_cache.get(role)
        if book is not None:
            try:
                _ = book.name
                if cached_sheet is not None:
                    _ = cached_sheet.name
            except Exception as e:
                log_excel_exception(role, e, "cached COM object went stale")
                record_excel_connection(role, "cached_book_unresponsive", str(e))
                book = None
                cached_sheet = None
                _book_sheet_cache.pop(role, None)
                _sheet_cache.pop(role, None)
        if book is None:
            book = find_open_book(xw, path)
            if book is not None:
                _book_sheet_cache[role] = book
                _sheet_cache.pop(role, None)
        if book is None:
            forget_workbook_formatting(path)
            # If the file itself has gone, allow one more open: it was
            # recreated from the template, so the customer still needs to see
            # it. If the file is still there, they closed it on purpose and
            # we leave it closed.
            if not os.path.exists(path):
                _auto_open_done[role] = False
            apps_count = excel_apps_count(xw)
            status = (
                "excel_running_workbook_not_found"
                if apps_count > 0
                else "excel_not_found"
            )
            record_excel_connection(role, status, f"expected_workbook={path}")
            message = (
                f"OPEN {os.path.basename(path)}"
                if apps_count > 0
                else "EXCEL IS CLOSED - reopen it"
            )
            if apps_count > 0:
                now = time.monotonic()
                # Never ask Excel to open a second workbook with a name it
                # already has open. Excel answers that with a modal "cannot
                # open two workbooks with the same name" dialog, and the tool
                # used to put that dialog up every twenty seconds for hours.
                # The open copy is used instead (see find_open_book).
                same_name_open = False
                try:
                    same_name_open = workbook_name_open_anywhere_in_apps(
                        list(xw.apps),
                        path,
                        lambda exc, context: log_excel_exception(
                            role, exc, context
                        ),
                    )
                except Exception as e:
                    log_excel_exception(role, e, "checking for a same-named workbook")
                if same_name_open:
                    record_excel_connection(
                        role,
                        "same_name_open_elsewhere",
                        f"expected_workbook={path}",
                    )
                if (
                    reopen_due(now, _reopen_after[role])
                    and not any_book_named(xw, path)
                    and not same_name_open
                    and not _auto_open_done.get(role)
                ):
                    _reopen_after[role] = now
                    # Open it once, then leave it alone.
                    #
                    # Reopening on a timer meant that a customer who closed
                    # the workbook - or closed Excel - watched it pop back up
                    # every twenty seconds, and each attempt could put Excel's
                    # "cannot open two workbooks with the same name" dialog
                    # on screen. One attempt, then we wait for the customer.
                    _auto_open_done[role] = True
                    try:
                        os.startfile(path)
                        record_excel_connection(
                            role, "workbook_reopen_requested", f"path={path}"
                        )
                    except OSError as e:
                        log_excel_exception(role, e, "reopening missing workbook")
                        try:
                            append_app_exception(f"reopening {role} workbook", e)
                        except OSError as log_error:
                            print(
                                f"Could not log workbook reopen failure ({role}): {log_error}",
                                flush=True,
                            )
                if excel_connected[role]:
                    notice_workbook_missing(path)
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
                if not _excel_startup_notice_shown.get(role):
                    _excel_startup_notice_shown[role] = True
                    show_error(
                        "Microsoft Excel is taking a long time to appear.\n\n"
                        "The tool is still running and will start as soon as\n"
                        "Excel is open with the chain workbooks.\n\n"
                        "Open Excel and the three workbook files, then leave\n"
                        "them open. No need to run anything else."
                    )
            return
        excel_miss_counts[role] = 0
        _reopen_after[role] = 0.0
        clear_workbook_missing_notice(path)
        excel_connected[role] = True
        _excel_startup_notice_shown[role] = False
        show_status(book, role, "LIVE")
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
        sheet = cached_sheet
        if sheet is None:
            sheet = book.sheets[0]
            _sheet_cache[role] = sheet

        # The tool writes to fixed columns. If the customer has inserted or
        # deleted a column, those positions no longer line up and the tool
        # would write over their own data. Detect it and stop rather than
        # corrupting the sheet.
        if not sheet_structure_ok(sheet, role):
            if not _sheet_altered.get(role):
                _sheet_altered[role] = True
                record_excel_connection(role, "sheet_layout_altered", f"path={path}")
                try:
                    show_error(
                        f"The layout of {os.path.basename(path)} has changed, so "
                        "the tool has stopped writing to it.\n\n"
                        "A column was inserted or deleted, which moves the "
                        "chain out of the cells the tool uses, and writing "
                        "there would overwrite your own data.\n\n"
                        "Undo the column change, or close this workbook "
                        "without saving so the tool can rebuild it."
                    )
                except Exception:
                    pass
            return

        formatting_started = time.perf_counter()
        try:
            format_sheet(sheet, role)
        finally:
            _performance_metrics.record(
                role, "worksheet_setup", time.perf_counter() - formatting_started
            )

        global _cycle_wrote
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
            _colour_state.clear()
            us_state._refetch_now = True
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
            range_val = MAX_A2  # empty A2 â†’ all (API returns â‰¤100 around ATM)
        else:
            desired = int(v)
            if desired > MAX_A2:
                desired = MAX_A2
                sheet.range("A2").value = desired
            range_val = max(desired, 1)

        # Read the whole header row once.
        #
        # Every COM call into Excel costs about 33ms on this machine, no
        # matter how small the range is, so asking for one cell at a time is
        # what made the sheet feel slow: six reads per sheet, every cycle.
        # One read of A1:AE1 covers all three blocks.
        try:
            header_values = list(sheet.range((1, 1), (1, 31)).value or [])
        except Exception as e:
            log_excel_exception(role, e, "reading the header row")
            header_values = []

        # Blocks of this sheet whose data changed and still has to be written.
        pending_writes = []

        for bi, block in enumerate(BLOCK_COLS):
            expiry = active_expiries[sym][bi]
            col_l = block["start"]
            if not expiry:
                # Only rewrite header cells when the text actually changed.
                # Writing an identical string still dirties the cell and forces
                # every dependent formula to recalculate â€” that cost is what
                # makes a strategy workbook feel stuck.
                head_text = f"{sym} Spot: --"
                pcr_text = "PCR: --"
                try:
                    if _header_at(header_values, col_l) != head_text:
                        sheet.range((1, col_l)).value = head_text
                        _set_header_at(header_values, col_l, head_text)
                        _cycle_wrote = True
                    if _header_at(header_values, col_l + 4) != pcr_text:
                        sheet.range((1, col_l + 4)).value = pcr_text
                        _set_header_at(header_values, col_l + 4, pcr_text)
                        _cycle_wrote = True
                except Exception as e:
                    log_excel_exception(role, e, "writing empty-block header")
                # Clear stale rows only when this block was previously non-empty,
                # or on first sighting. Clearing to a fixed 2000 rows every cycle
                # dirties thousands of cells for no visible difference.
                prev_n = _last_rows.get((role, bi))
                if prev_n is None or prev_n != 0:
                    first_sight = _FIRST_SIGHT_BOUND if prev_n is None else 0
                    last_row = max(prev_n or 0, first_sight)
                    clear_start = get_column_letter(col_l)
                    clear_end = get_column_letter(col_l + 8)
                    try:
                        sheet.range(f"{clear_start}5:{clear_end}{4 + last_row}").clear()
                    except Exception as e:
                        log_excel_exception(role, e, f"clearing empty block {bi + 1}")
                _last_rows[(role, bi)] = 0
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
            _head = f"{sym} Spot: {sp:.2f} | {cnt} strikes"
            # Never let a delay go unannounced. A silent drop to the CBOE
            # fallback means the trader is reading 15-minute-old prices, and
            # that is exactly what "the data is stale" means to them.
            _note = source_delay_note(sym)
            if _note:
                _head = f"{_head} | {_note}"
            pcr_text = f"PCR: {pcr:.2f}"
            try:
                if _header_at(header_values, col_l) != _head:
                    sheet.range((1, col_l)).value = _head
                    _set_header_at(header_values, col_l, _head)
                    _cycle_wrote = True
                if _header_at(header_values, col_l + 4) != pcr_text:
                    sheet.range((1, col_l + 4)).value = pcr_text
                    _set_header_at(header_values, col_l + 4, pcr_text)
                    _cycle_wrote = True
            except Exception as e:
                log_excel_exception(role, e, "writing chain header")
            net_paint(sheet, role, bi, col_l)

            if chain is not None and not chain.empty:
                # set OI as text BEFORE write so Excel cannot coerce $223 -> INR number
                n_rows = len(chain)
                c0 = block["start"]
                # Setting the number format dirties the cells, so do it once
                # per block rather than on every refresh.
                fmt_key = (role, bi, n_rows)
                if fmt_key not in _oi_fmt_done:
                    for off in (3, 5):
                        try:
                            sheet.range((5, c0 + off), (4 + n_rows, c0 + off)).api.NumberFormat = "@"
                        except Exception as e:
                            log_excel_exception(
                                role, e, f"formatting open-interest column {c0 + off}"
                            )
                    _oi_fmt_done.add(fmt_key)

                write_started = time.perf_counter()
                try:
                    # Only write when the contents actually differ. Any data
                    # change alters the signature and forces the write, so live
                    # updates still land; unchanged blocks cost nothing.
                    values = chain.values.tolist()
                    payload_sig = hash(tuple(map(tuple, values)))
                    if payload_sig != _block_payload.get((role, bi)):
                        prev = _block_values.get((role, bi))
                        changed = None
                        if prev is not None and len(prev) == len(values):
                            changed = [
                                i for i, (a, b) in enumerate(zip(prev, values)) if a != b
                            ]
                        runs = _contiguous_runs(changed) if changed else []
                        _block_values[(role, bi)] = values
                        _block_payload[(role, bi)] = payload_sig
                        _cycle_wrote = True
                        # The Excel write is deferred to the end of the block
                        # loop, so all of this sheet's blocks go out together.
                        pending_writes.append((bi, block, chain, n_rows, runs))
                    else:
                        _performance_metrics.record(role, "excel_range_skipped", 0.0, rows=n_rows)
                finally:
                    _performance_metrics.record(
                        role,
                        "excel_range_write",
                        time.perf_counter() - write_started,
                        rows=n_rows,
                    )

                # Clear whatever sits below the live table.
                #
                # Only clears rows that can hold leftovers: what this block used
                # previously, plus on first sighting a generous sweep to catch
                # anything an earlier run left behind.
                prev_n = _last_rows.get((role, bi))
                if prev_n is None or n_rows < prev_n:
                    first_sight = _FIRST_SIGHT_BOUND if prev_n is None else 0
                    last_row = max(n_rows, prev_n or 0, first_sight)
                    stale_a = (
                        f"{get_column_letter(c0)}{5 + n_rows}"
                        f":{get_column_letter(c0 + 8)}{4 + last_row}"
                    )
                    try:
                        sheet.range(stale_a).clear()
                    except Exception as e:
                        log_excel_exception(role, e, f"clearing stale rows in block {bi + 1}")
                _last_rows[(role, bi)] = n_rows

                formatting_started = time.perf_counter()
                try:
                    # force_usd_block_formats is self-gating, so it only does
                    # work when it must.
                    force_usd_block_formats(sheet, block, n_rows)

                    strikes = chain["strike"].tolist()
                    cs = get_column_letter(c0)
                    ce = get_column_letter(c0 + 8)

                    # Recolour only when the colouring actually changes.
                    #
                    # The colours are derived from spot: the call block below
                    # spot is ITM green, the put block above it is ITM red, and
                    # the ATM row is yellow. Those boundaries only move when
                    # spot crosses a strike, which is rare. The previous code
                    # cleared and repainted the whole block on every cycle,
                    # which dirties every cell of a 600-row block and makes
                    # Excel recalculate the customer's own formulas for no
                    # visible change.
                    colour_sig = color_signature(strikes, sp, len(strikes))
                    if _colour_state.get((role, bi)) != colour_sig:
                        sheet.range(f"{cs}5:{ce}{4 + n_rows}").color = None
                        if strikes:
                            for cell_range, color in itm_color_ranges(block, strikes, sp):
                                sheet.range(cell_range).color = color

                            atm_strike = min(strikes, key=lambda s: abs(s - sp))
                            atm_row = strikes.index(atm_strike) + 5

                            atm_range = f"{block['atm']}{atm_row}:{block['put_end']}{atm_row}"
                            sheet.range(atm_range).color = (255, 255, 0)
                            sheet.range(atm_range).font.bold = True
                        _colour_state[(role, bi)] = colour_sig
                finally:
                    _performance_metrics.record(
                        role,
                        "excel_formatting",
                        time.perf_counter() - formatting_started,
                        rows=n_rows,
                    )
            else:
                # An empty block only needs clearing when it stops being empty,
                # or on first sighting. Clearing 1000 rows every cycle was both
                # wasteful and needlessly expensive for Excel to do.
                prev_n = _last_rows.get((role, bi))
                if prev_n is None or prev_n != 0:
                    first_sight = _FIRST_SIGHT_BOUND if prev_n is None else 0
                    last_row = max(prev_n or 0, first_sight)
                    clear_start = get_column_letter(col_l)
                    clear_end = get_column_letter(col_l + 8)
                    try:
                        sheet.range(f"{clear_start}5:{clear_end}{4 + last_row}").clear()
                    except Exception as e:
                        log_excel_exception(role, e, f"clearing empty chain block {bi + 1}")
                _last_rows[(role, bi)] = 0

        # Push this sheet's changed blocks out in as few COM calls as possible.
        #
        # Every call into Excel costs roughly 33ms on this machine whatever
        # the range size, so three blocks written separately cost three calls
        # and about 100ms of pure waiting. The blocks sit side by side on the
        # sheet, so they can go out as one range with the spacer columns
        # between them left blank.
        if pending_writes:
            _flush_block_writes(sheet, role, pending_writes)
    except Exception as e:
        print(f"Excel error ({role} / {role_symbol.get(role)}): {e}", flush=True)
        record_excel_connection(role, "excel_com_error", f"{type(e).__name__}: {e}")
        log_excel_exception(role, e)


def _flush_block_writes(sheet, role, pending):
    """Write the changed blocks of one sheet.

    How the block is written depends on whether we own Excel's calculation
    mode.

    - When we have Excel in manual calculation, nothing recalculates until we
      ask it to, so the cost is the COM calls. Write the whole block in one
      call.
    - When we could not get manual calculation - Excel was busy, or a dialog
      was up - every cell written can trigger a recalculation of the
      customer's own workbook. That is the case that made their sheet crawl, so
      write as few cells as possible: only the rows that moved, one call per
      contiguous run.

    Either way each block is written on its own, into its own columns, so a
    block that did not change is never touched.
    """
    if not pending:
        return
    started = time.perf_counter()
    protected = _calc_session.get("owned") and not _calc_session.get("failed")
    try:
        for bi, block, chain, n_rows, runs in pending:
            try:
                if protected:
                    if len(runs) == 1:
                        start, end = runs[0]
                        sheet.range(
                            (5 + start, block["start"]),
                            (4 + end, block["start"] + 8),
                        ).value = chain.iloc[start : end + 1].values
                        _performance_metrics.record(
                            role, "excel_rows_written", 0.0, rows=end - start + 1
                        )
                    else:
                        sheet.range((5, block["start"])).options(
                            index=False, header=False
                        ).value = chain
                else:
                    # Unprotected window: every cell can trigger a recalc of
                    # the customer's workbook, so touch as few cells as we can.
                    if runs and len(runs) <= 8:
                        for start, end in runs:
                            sheet.range(
                                (5 + start, block["start"]),
                                (4 + end, block["start"] + 8),
                            ).value = chain.iloc[start : end + 1].values
                        _performance_metrics.record(
                            role, "excel_rows_written", 0.0,
                            rows=sum(end - start + 1 for start, end in runs),
                        )
                    else:
                        sheet.range((5, block["start"])).options(
                            index=False, header=False
                        ).value = chain
            except Exception as e:
                log_excel_exception(role, e, f"writing chain block {bi + 1}")
    finally:
        _performance_metrics.record(
            role,
            "excel_block_flush",
            time.perf_counter() - started,
            rows=sum(item[3] for item in pending),
        )


def write_to_excel(role):
    started = time.perf_counter()
    try:
        _write_to_excel(role)
    finally:
        _performance_metrics.record(
            role, "excel_pass", time.perf_counter() - started
        )


def _time_recalc(wrote):
    """Recalculate once inside the manual window and time it.

    This is the real cost the customer feels: how long their own strategy
    workbook takes to recalc when we change a cell it depends on. Plain
    Application.Calculate, not CalculateUntilAsyncQueriesDone (that is for
    asynchronous queries such as Power Query). Only done when we actually
    wrote something and Excel was on auto/semiauto â€” recalculating for a
    customer who chose manual would undo their choice.
    """
    global _last_recalc_ms
    if not wrote:
        _performance_metrics.record("ALL", "excel_recalc_skip_nowrite", 0.0)
        return 0.0
    if not _bulk_will_recalc:
        _performance_metrics.record("ALL", "excel_recalc_skip_mode", 0.0)
        return 0.0
    app = _app_cache
    if app is None:
        _performance_metrics.record("ALL", "excel_recalc_skip_noapp", 0.0)
        return 0.0
    try:
        started = time.perf_counter()
        app.api.Calculate()
        elapsed = (time.perf_counter() - started) * 1000.0
    except Exception as e:
        log_excel_exception("ALL", e, "forcing workbook recalculation")
        return 0.0
    _last_recalc_ms = elapsed
    _performance_metrics.record("ALL", "excel_recalc", elapsed / 1000.0)
    return elapsed


def render_cycle():
    """One full pass over every workbook, inside a single bulk-mode window.

    Screen updating and calculation are suspended once per cycle rather than
    once per book, so a cycle costs three writes with repaint and recalculation
    suspended, then one restore. Returns nothing; timing is recorded in
    _performance_metrics.
    """
    started = time.perf_counter()
    global _cycle_wrote
    _cycle_wrote = False

    # If Excel would not switch to manual calculation, every cell we write
    # makes it recalculate. Writing anyway at full speed is what made the
    # sheet feel stuck, so slow down to a rate the customer's workbook can
    # absorb, and tell them once what to clear.
    if _calc_manual_failed:
        now = time.monotonic()
        if now - _last_unprotected_write < _UNPROTECTED_WRITE_INTERVAL:
            _performance_metrics.record("ALL", "render_cycle_skipped_busy", 0.0)
            return
        _last_unprotected_write = now
        _performance_metrics.record("ALL", "render_cycle_unprotected", 0.0)
        if not _calc_manual_notice_shown:
            _calc_manual_notice_shown = True
            try:
                show_error(
                    "Excel would not switch to manual calculation, so the "
                    "workbook is recalculating on every update and will feel "
                    "slow.\n\n"
                    "Close any Excel dialog that is open, click outside the "
                    "cell you are editing, and close any third-party Excel "
                    "add-in. The tool will speed back up on its own."
                )
            except Exception:
                pass

    restore = _begin_excel_bulk_mode(bulk_app())
    try:
        for role in ROLE_ORDER:
            if not excel_process_running():
                break
            try:
                write_to_excel(role)
            except Exception as e:
                log_excel_exception(role, e, "render pass")
        _time_recalc(_cycle_wrote)
    finally:
        if restore is not None:
            try:
                restore()
                # If we already recalculated inside the window, restoring
                # automatic should find nothing left to do. A second pass means
                # the cost is being paid twice.
                try:
                    if _app_cache is not None and int(_app_cache.api.CalculationState) != 0:
                        _performance_metrics.record("ALL", "excel_recalc_second_pass", 1.0)
                    else:
                        _performance_metrics.record("ALL", "excel_recalc_second_pass", 0.0)
                except Exception:
                    pass
            except Exception as e:
                log_excel_exception("ALL", e, "leaving bulk mode")
        elapsed = time.perf_counter() - started
        _performance_metrics.record("ALL", "render_cycle", elapsed)
        # Never leave the trader in manual calculation between write windows.
        try:
            ensure_calculation_not_stranded()
        except Exception:
            pass
        _log_speed_summary(elapsed)


def _log_speed_summary(cycle_ms):
    """Write the whole pipeline's timings into the log every half minute.

    Every future complaint should come with numbers, so one line carries the
    whole picture: how long a cycle took, how long the customer's own workbook
    took to recalculate, how fresh the index quote is, and how many contracts
    are being tracked.
    """
    global _speed_log_at
    now = time.monotonic()
    if now - _speed_log_at < 30.0:
        return
    _speed_log_at = now
    feed = _databento_feed
    quotes = 0
    if feed is not None:
        try:
            with feed._lock:
                quotes = len(feed._quotes)
        except Exception:
            quotes = 0
    quote_age = None
    quote_feed = None
    try:
        quote_feed, quote_age = spot_source_for("SPX")
    except Exception:
        pass
    feed_log(
        "speed",
        cycle_ms=round(cycle_ms * 1000.0, 1),
        recalc_ms=round(_last_recalc_ms, 1) if _last_recalc_ms else 0,
        quotes=quotes,
        spot_quote_age_s=(round(quote_age, 1) if quote_age is not None else None),
        spot_feed=quote_feed,
    )


def bulk_app():
    """Resolve the Excel App COM object to use for bulk mode, caching it."""
    global _app_cache
    try:
        import xlwings as _xw
    except ImportError:
        return None
    if _app_cache is not None:
        try:
            _ = _app_cache.version
            return _app_cache
        except Exception:
            _app_cache = None
    try:
        for role in ROLE_ORDER:
            book = find_open_book(_xw, excel_path(role))
            if book is not None:
                _app_cache = book.app
                return _app_cache
    except Exception as e:
        log_excel_exception("ALL", e, "resolving bulk-mode app")
    return None


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


# xlwings' .calculation getter hands back a stale value while Excel is busy, so a
# "did the restore work?" check built on it can pass falsely and leave the trader
# stuck on manual. Set AND verify through the raw API with these ids.
_CALC_IDS = {"automatic": -4105, "semiautomatic": 2, "manual": -4135}


def _raw_calc_mode(app):
    return int(app.api.Calculation)


def _set_calc_mode(app, name):
    app.api.Calculation = _CALC_IDS[name]
    return _raw_calc_mode(app) == _CALC_IDS[name]


def ensure_calculation_not_stranded(app=None):
    """Safety net: never leave the trader in manual calculation.

    If Excel is still on manual well after any write window should have closed,
    put it back. Otherwise one failed restore makes every formula the trader
    types silently fail to compute.
    """
    global _manual_since
    if app is None:
        app = _app_cache
    if app is None:
        return
    try:
        mode = _raw_calc_mode(app)
    except Exception:
        return
    now = time.time()
    if mode != -4135:
        _manual_since = 0.0
        return
    if _manual_since == 0.0:
        _manual_since = now
        return
    if now - _manual_since < MANUAL_WATCHDOG_SECONDS:
        return
    try:
        if _set_calc_mode(app, "automatic"):
            _manual_since = 0.0
    except Exception:
        pass


def _begin_excel_bulk_mode(app=None):
    """Suspend repaint + recalculation for the duration of one render cycle.

    Every Excel write otherwise repaints the workbook and triggers a full
    recalculation. That is cheap in an empty sheet but expensive when the user
    has a strategy workbook open in the same Excel instance, which is where
    refresh times blow up and where the "it feels slow" complaint comes from.

    Returns a restore callable, or None if nothing could be changed. Failures
    are expected and non-fatal: both settings raise if the user is mid-edit or
    a dialog is up, and there may be no workbook to talk to yet.
    """
    state = {"app": None, "screen": None, "calc": None}
    global _app_cache, _last_bulk_restore
    try:
        if app is None:
            app = _app_cache
        if app is None:
            import xlwings as xw
            for role in ROLE_ORDER:
                book = find_open_book(xw, excel_path(role))
                if book is not None:
                    app = book.app
                    break
            if app is None:
                return None
        _app_cache = app
        try:
            # xlwings.App has no .name; .version is the cheap liveness probe
            # that goes through to Excel and raises if the instance is gone.
            _ = app.version
        except Exception as e:
            log_excel_exception("ALL", e, "cached Excel app went stale")
            _app_cache = None
            return None
        state["app"] = app
        try:
            if _screen_pause_active():
                # Setting it keeps failing and Excel is busy. Stop trying for a
                # while rather than making four failing calls every cycle.
                state["screen"] = None
            elif app.screen_updating:
                state["screen"] = True
                app.screen_updating = False
                _note_screen_ok()
            else:
                # Already off: either the customer turned it off, or the other
                # OptionChain tool has it off for its own write window. Leave it
                # and do not record it, or we would "restore" the other tool's
                # temporary setting and leave the customer's Excel hidden.
                state["screen"] = None
        except Exception as e:
            _note_screen_failed(e)
            log_excel_exception("ALL", e, "disabling screen_updating")
            state["screen"] = None
        try:
            # Read the mode through the raw API on purpose. xlwings' own getter
            # is calculation_i2s[self.xl.Calculation], and when Excel is busy or
            # showing a modal it hands back an HRESULT instead of a valid
            # xlCalculation value, so that lookup raises KeyError.
            try:
                raw = int(app.api.Calculation)
            except Exception:
                raw = None
            current = {
                -4105: "automatic",   # xlCalculationAutomatic
                2: "semiautomatic",   # xlCalculationSemiautomatic
                -4135: "manual",      # xlCalculationManual
            }.get(raw)
            if current == "manual":
                # Already manual: either the customer wants that, or the other
                # OptionChain tool has it off for its own write window. Either way,
                # leave it alone and do not record it for restore -- recording
                # it would "restore" the other tool's temporary setting and
                # leave the customer's Excel stuck on manual calculation.
                state["calc"] = None
                _calc_session["owned"] = False
            else:
                state["calc"] = current or "automatic"
                # Written BEFORE Excel is changed. If this process dies inside
                # the window, the next launch finds the record and puts Excel
                # back, instead of leaving their formulas silently dead.
                calc_mode_guard.acquire(state["calc"], PRODUCT_ID)
                state["calc_guarded"] = True
                # If the workbook was already on automatic, Excel will
                # recalculate by itself once we restore. Semiautomatic does
                # the same for ordinary formulas.
                global _bulk_will_recalc
                _bulk_will_recalc = state["calc"] in ("automatic", "semiautomatic")
                if _calc_session["failed"]:
                    # We already know Excel will not let us change the mode.
                    # Stop asking: every failed attempt costs a COM call and
                    # makes Excel busier, which is the loop that made the
                    # customer's workbook crawl.
                    state["calc"] = None
                    state["calc_guarded"] = False
                else:
                    # Take the calculation mode once, for the whole session.
                    #
                    # It used to be set at the start of every cycle and restored
                    # at the end of every cycle. On the customer's machine
                    # that switch failed almost every time, so the write window
                    # was never protected and every cell we touched triggered a
                    # full recalculation of their own workbook. Leaving Excel
                    # in manual for the session and recalculating it ourselves
                    # once per cycle is both cheaper and far less likely to
                    # collide with the other OptionChain tool.
                    try:
                        app.calculation = "manual"
                        _calc_session["owned"] = True
                        _calc_session["mode"] = state["calc"]
                        _calc_session["failed"] = False
                        _calc_session["app"] = app
                    except Exception as exc:
                        _calc_session["failed"] = True
                        _calc_session["owned"] = False
                        feed_log(
                            "calc_manual_unavailable",
                            error=type(exc).__name__,
                            detail=str(exc)[:120],
                        )
                        if not _calc_notice_shown:
                            _calc_notice_shown = True
                            try:
                                show_error(
                                    "Excel will not switch to manual calculation, "
                                    "so the workbook recalculates on every "
                                    "update and can feel slow.\n\n"
                                    "This is usually one of three things:\n"
                                    "  - a cell is in edit mode - press Escape\n"
                                    "  - an Excel dialog is open - close it\n"
                                    "  - an Excel add-in is busy - turn it off\n\n"
                                    "The tool carries on and recovers on its own, "
                                    "but Excel will be working harder than it "
                                    "needs to until that is cleared."
                                )
                            except Exception:
                                pass
                        state["calc"] = None
                        state["calc_guarded"] = False
        except Exception as e:
            log_excel_exception("ALL", e, "switching to manual calculation")
            state["calc"] = None
    except Exception as e:
        log_excel_exception("ALL", e, "entering bulk mode")
        return None

    def restore():
        if state.get("restored"):
            return
        state["restored"] = True
        app = state.get("app")
        if app is None:
            return
        try:
            _ = app.version
        except Exception:
            return
        # The calculation mode is deliberately NOT restored here. We take it
        # once for the whole session and release it on exit, so the two
        # OptionChain tools no longer fight over it several times a second.
        # _emergency_restore() puts it back when the tool stops.
        if state.get("screen") is not None:
            try:
                if not app.screen_updating:
                    app.screen_updating = state["screen"]
            except Exception as e:
                log_excel_exception("ALL", e, "restoring screen_updating")

    _last_bulk_restore = restore
    return restore


async def display_loop():
    is_windows = platform.system() == "Windows"
    license_tick = 0
    global _excel_gone_logged, _app_cache
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
            if excel_alive:
                if _excel_gone_logged:
                    # Excel came back. Everything cached about the sheets is now
                    # wrong: the workbooks are freshly opened and their cells are
                    # empty, so the payload/shape/format caches must be dropped
                    # or the app would treat the blank sheets as "already
                    # written" and leave them blank.
                    _excel_gone_logged = False
                    _block_payload.clear()
                    _colour_state.clear()
                    _sheet_fmt_done.clear()
                    _usd_fmt_done.clear()
                    _dropdown_sig.clear()
                    _app_cache = None
                    for role in ROLE_ORDER:
                        _reopen_after[role] = 0.0
                    print("Excel reopened - rewriting all sheets from scratch.", flush=True)
                    _last_excel_exception.clear()

                render_cycle()

            elif not _excel_gone_logged:
                _excel_gone_logged = True
                print("Excel is not running - pausing writes until it is reopened.", flush=True)
                _last_excel_exception.clear()
            render_seconds = time.perf_counter() - render_started
            _performance_metrics.record("ALL", "render_cycle", render_seconds)
            # Never leave the trader in manual calculation between write windows.
            try:
                ensure_calculation_not_stranded()
            except Exception:
                pass
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
        await asyncio.sleep(
            refresh_delay(
                REFRESH,
                REFRESH,
                time.perf_counter() - render_started,
            )
        )


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
    # After opening the workbooks, not before: at process start xw.apps can be
    # empty even though Excel is already running, and recovery then gave up and
    # discarded the record while leaving Excel stuck on manual.
    recover_calc_mode()
    errs = []
    for sym in active_symbols():
        try:
            await init_symbol(sym)
        except Exception as e:
            errs.append(f"{sym}: {e}")
    if errs and len(errs) >= len(active_symbols()):
        show_error(
            "Could not reach option data sources (Databento / BigClawd / CBOE).\n\n"
            "Check internet connection and run again.\n\n"
            + "\n".join(errs)
        )
        return
    await asyncio.gather(
        data_poller(),
        spot_poller(),
        expiry_roller(),
        process_pending_dates(),
        display_loop(),
    )


def _instance_guard_log(message):
    try:
        path = os.path.join(BASE_DIR, "instance_guard.log")
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except OSError:
        pass


def _same_folder(first, second):
    """True when two paths name the same directory."""
    try:
        return os.path.normcase(os.path.abspath(first)) == os.path.normcase(
            os.path.abspath(second)
        )
    except Exception:
        return False


def _lock_path():
    return os.path.join(BASE_DIR, "instance.lock")


def _read_lock_pid():
    """PID recorded in this folder's lock file, or None."""
    try:
        with open(_lock_path(), encoding="ascii") as fh:
            return int(fh.read().strip().split()[0])
    except (OSError, ValueError, IndexError):
        return None


def _write_lock_pid(pid):
    try:
        with open(_lock_path(), "w", encoding="ascii") as fh:
            fh.write(str(int(pid)))
    except OSError:
        pass


def _pid_alive(pid):
    """True only if the process is genuinely still running.

    OpenProcess succeeds on a process that has been terminated but whose
    object has not been released yet, so opening it is not proof of life. The
    exit code is: STILL_ACTIVE (259) means it is running, anything else means
    it has finished.
    """
    if not pid or pid <= 0:
        return False
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            SYNCHRONIZE = 0x00100000
            STILL_ACTIVE = 259
            handle = ctypes.windll.kernel32.OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, False, int(pid)
            )
            if not handle:
                return False
            try:
                code = wintypes.DWORD()
                if not ctypes.windll.kernel32.GetExitCodeProcess(
                    handle, ctypes.byref(code)
                ):
                    return False
                return code.value == STILL_ACTIVE
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        except Exception:
            return False
    try:
        os.kill(int(pid), 0)
    except OSError:
        return False
    except Exception:
        return True
    return True


def _terminate_pid(pid):
    """Terminate a process. True when it is gone afterwards."""
    if os.name == "nt":
        try:
            import ctypes
            PROCESS_TERMINATE = 0x0001
            SYNCHRONIZE = 0x00100000
            handle = ctypes.windll.kernel32.OpenProcess(
                PROCESS_TERMINATE | SYNCHRONIZE, False, int(pid)
            )
            if not handle:
                return not _pid_alive(pid)
            try:
                ctypes.windll.kernel32.TerminateProcess(handle, 1)
                # Wait for it to actually go, rather than guessing.
                ctypes.windll.kernel32.WaitForSingleObject(handle, 5000)
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        except Exception:
            return not _pid_alive(pid)
    else:
        try:
            import signal
            os.kill(int(pid), signal.SIGTERM)
        except Exception:
            return not _pid_alive(pid)
    for _ in range(30):
        if not _pid_alive(pid):
            return True
        time.sleep(0.1)
    return not _pid_alive(pid)


def _take_over_locked_instance():
    """Stop an older instance of this tool running from this same folder.

    The lock file lives in this folder, so a PID found in it is by definition
    an instance of this tool in this folder. That is what lets a second
    launch take over even when it cannot read the other process's executable
    path - which is what happens when the two run at different privilege
    levels or when security software blocks the query.

    Returns True when another instance was found and stopped, False when there
    was none, and raises RuntimeError when one was found but could not be
    stopped, because two writers on the same workbooks is worse than one
    writer refusing to start.
    """
    pid = _read_lock_pid()
    if not pid or pid == os.getpid():
        return False
    if not _pid_alive(pid):
        _instance_guard_log(f"  lock pid={pid} is gone, taking over")
        return False
    _instance_guard_log(f"  lock pid={pid} is alive in this folder, stopping it")
    if _terminate_pid(pid):
        _instance_guard_log(f"  lock pid={pid} stopped; this launch takes over")
        return True
    raise RuntimeError(
        f"Another copy of {PRODUCT_EXE} (pid {pid}) is still running from "
        f"this folder and could not be stopped."
    )


def _process_image_path(kernel32, pid):
    """Full path of a running process, or None if it cannot be read."""
    import ctypes
    from ctypes import wintypes

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return None
    try:
        size = wintypes.DWORD(32768)
        buf = ctypes.create_unicode_buffer(size.value)
        if not kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return None
        return buf.value
    finally:
        kernel32.CloseHandle(h)


def kill_previous_instances():
    """Stop an older copy of *this* tool so a double-click takes over cleanly.

    The executable is named for the product - us_option_chain.exe - so the
    process name alone says which tool it is. That removes the need to guess
    the product from the folder, which is what went wrong when both packages
    were extracted into the same place.

    The lock file in this folder is checked first. It proves that a process
    running from this folder is an instance of this tool, so it can be stopped
    without having to read its executable path - which is exactly what fails
    when the two run at different privilege levels or when security software
    blocks the query.

    Never let the process scan break startup: any failure there is logged and
    ignored. A duplicate that cannot be stopped is a hard error, because two
    writers on the same workbooks is worse than one writer refusing to start.
    """
    if platform.system() != "Windows":
        return
    my_pid = os.getpid()
    my_dir = os.path.normcase(os.path.abspath(BASE_DIR))
    _instance_guard_log(
        f"pid={my_pid} product={PRODUCT_ID} exe={PRODUCT_EXE} dir={my_dir}"
    )

    # A previous instance of this tool in this folder, stopped without needing
    # to read its path.
    try:
        _take_over_locked_instance()
    except RuntimeError as e:
        _instance_guard_log(f"  duplicate could not be stopped: {e}")
        raise
    except Exception as e:
        _instance_guard_log(f"  lock check failed: {type(e).__name__}: {e}")

    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32
        TH32CS_SNAPPROCESS = 0x2

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.POINTER(wintypes.ULONG)),
                ("th32ModuleID", wintypes.DWORD),
                ("th32Threads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", ctypes.c_char * 260),
            ]

        snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if snap == -1 or snap is None:
            _instance_guard_log(f"snapshot failed; my_dir={my_dir}")
            return
        try:
            entry = PROCESSENTRY32()
            entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
            matches = []
            if kernel32.Process32First(snap, ctypes.byref(entry)):
                while True:
                    # The executable is named for the product, so the process
                    # name alone identifies it. The other product's tool has a
                    # different name and is never a candidate.
                    if entry.szExeFile.decode(errors="ignore").lower() == PRODUCT_EXE:
                        matches.append(int(entry.th32ProcessID))
                    if not kernel32.Process32Next(snap, ctypes.byref(entry)):
                        break
        finally:
            kernel32.CloseHandle(snap)

        others = [pid for pid in matches if pid != my_pid]
        _instance_guard_log(
            f"pid={my_pid} product={PRODUCT_ID} dir={my_dir} "
            f"matched={matches} others={others}"
        )

        for pid in others:
            path = _process_image_path(kernel32, pid)
            if not path:
                # Cannot read the path, so cannot tell which folder it is in.
                # The lock file already covers the same-folder case, so leave
                # this one alone rather than killing another product's tool.
                _instance_guard_log(f"  pid={pid} path unreadable, left alone")
                continue
            other_dir = os.path.normcase(os.path.dirname(os.path.abspath(path)))
            if not _same_folder(other_dir, my_dir):
                _instance_guard_log(
                    f"  pid={pid} at {other_dir}, different folder, left alone"
                )
                continue
            if _terminate_pid(pid):
                _instance_guard_log(
                    f"  pid={pid} at {other_dir}, stopped; this launch takes over"
                )
    except Exception as e:  # never block startup on this
        _instance_guard_log(f"guard error: {type(e).__name__}: {e}")
    finally:
        # Record ourselves as the instance holding this folder, so the next
        # launch can take over without reading process paths.
        _write_lock_pid(my_pid)

def recover_calc_mode():
    """Put Excel back on automatic if a previous run died mid-write-window.

    A crash, a kill from Task Manager or a power cut while calculation was
    suspended would otherwise leave the customer's formulas silently stopped.
    The record is consulted first; Excel's current mode is only read if there
    is something to undo.
    """
    should, mode, reason = calc_mode_guard.pending_recovery()
    if not should:
        return False
    if reason:
        print(f"  calculation-mode record found: {reason}", flush=True)
    try:
        import xlwings as xw

        app = None
        for _ in range(10):
            for role in ROLE_ORDER:
                book = find_open_book(xw, excel_path(role))
                if book is not None:
                    app = book.app
                    break
            if app is not None:
                break
            time.sleep(0.5)
        if app is None:
            print("  calc recovery: no chain workbook is reachable", flush=True)
            return False          # keep the record so a later launch can retry
        if int(app.api.Calculation) != calc_mode_guard.XL_MANUAL:
            # Excel is not on manual, so there is nothing to undo. Clear the
            # record rather than touching a mode the customer may have set.
            calc_mode_guard.clear_record()
            return False
        app.calculation = mode or "automatic"
        calc_mode_guard.clear_record()
        print(f"  restored Excel calculation to {mode or 'automatic'}", flush=True)
        return True
    except Exception as e:
        print(f"  could not restore calculation mode: {e}", flush=True)
        return False


def _emergency_restore():
    """Put Excel back the way we found it, on a clean exit.

    The calculation mode is held for the whole session, so it has to be
    released here. Only if this tool set it in the first place - if it was
    already on manual calculation when we started, that belongs to the
    customer or to the other OptionChain tool and is left alone.
    """
    restore = _last_bulk_restore
    if restore is not None:
        try:
            restore()
        except Exception:
            pass
    if not _calc_session.get("owned"):
        return
    app = _calc_session.get("app")
    mode = _calc_session.get("mode")
    if app is None or not mode:
        return
    try:
        _ = app.version
    except Exception:
        return
    try:
        for _attempt in range(3):
            try:
                if _set_calc_mode(app, mode):
                    break
            except Exception:
                pass
            time.sleep(0.5)
        calc_mode_guard.release(PRODUCT_ID)
    except Exception:
        pass
    finally:
        _calc_session["owned"] = False


atexit.register(_emergency_restore)


def _start_databento():
    """Resolve the per-PC key and start the live feed.

    The Databento key is mandatory: if it cannot be resolved (env / key file /
    first-run dialog all empty or cancelled) the app exits with an error box.
    """
    global _databento_feed
    from importlib.util import find_spec

    if find_spec("databento") is None:
        show_error(
            "Databento is required but the 'databento' package is not installed.\n\n"
            "Reinstall the application or contact the owner."
        )
        os._exit(1)
    try:
        key = databento_feed.ensure_api_key()
    except Exception as e:
        feed_log("databento_key_error", error=type(e).__name__, detail=str(e)[:180])
        try:
            append_app_exception("resolving Databento API key", e)
        except OSError:
            pass
        show_error(
            "Could not resolve the Databento API key.\n\n"
            f"{type(e).__name__}: {e}\n\n"
            "Contact the owner for a valid key."
        )
        os._exit(1)
    if not key:
        feed_log("databento_no_key", note="dialog cancelled or blank")
        show_error(
            "Databento API key required.\n\n"
            "Enter the Databento API key to continue.\n"
            "The key is stored only on this PC, never in the app folder.\n"
            "App will close."
        )
        os._exit(1)
    _databento_feed = databento_feed.DatabentoFeed(key, event_log=feed_log)
    us_state._databento_feed = _databento_feed
    feed_log("databento_key_loaded", masked=databento_feed.mask_key(key))


def _stop_databento():
    feed = _databento_feed
    if feed is not None:
        try:
            feed.stop()
        except Exception:
            pass


atexit.register(_stop_databento)


def main():
    try:
        write_startup_diagnostics()
    except OSError as e:
        show_error(f"Could not write startup diagnostics:\n{e}")
    try:
        # Take over from any older instance BEFORE anything that can block.
        #
        # The activation dialog and the API-key dialog both block, and the
        # guard used to run after them. Two launches therefore both sat on a
        # dialog, neither ever reached the guard, and both went on to write the
        # same three workbooks. Two writers = two Databento sessions and a
        # sheet that shows a mix of two snapshots.
        try:
            kill_previous_instances()
        except RuntimeError as e:
            show_error(
                f"{e}\n\n"
                f"Two copies of {PRODUCT_EXE} writing the same workbooks "
                "will corrupt what you see, so this one will not start.\n\n"
                "Close the other copy - it is in the taskbar - and run "
                f"{PRODUCT_EXE} again."
            )
            os._exit(1)
        # One line that says what is running, from where, and whether this
        # instance owns Excel's calculation mode. The customer never has to
        # run a diagnostic command to find that out.
        feed_log(
            "startup",
            pid=os.getpid(),
            product=PRODUCT_ID,
            exe=PRODUCT_EXE,
            folder=BASE_DIR,
            build=_short_build(),
        )
        if not ensure_license():
            os._exit(1)
        _start_databento()
        asyncio.run(main_async())
    except KeyboardInterrupt:
        pass
    finally:
        _emergency_restore()


if __name__ == "__main__":
    main()
