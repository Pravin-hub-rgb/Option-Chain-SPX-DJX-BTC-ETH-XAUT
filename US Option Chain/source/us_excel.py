"""Excel COM layer of the US option chain.

Extracted from main.py as part of the module split: every write to the
workbooks, the bulk-mode window, the screen/calculation guards and the
render cycle live here. Shared state comes from us_state; chain building
from us_chain; workbook files and text helpers from us_workbook.
"""

import os
import platform
import re
import subprocess
import threading
import time
from collections import defaultdict
from datetime import date, datetime, timedelta

import us_state
from us_state import (
    MANUAL_WATCHDOG_SECONDS, POLL_GAP, PRODUCT_ID, REFRESH,
    ROLE_FILES, ROLE_ORDER, _CALC_IDS, _FIRST_SIGHT_BOUND,
    _SCREEN_BACKOFF_SECONDS, _SCREEN_FAILURE_LIMIT, _UNPROTECTED_WRITE_INTERVAL,
    _auto_open_done, _book_sheet_cache, _calc_failed_at, _calc_manual_failed,
    _calc_notice_shown, _calc_session,
    _dropdown_sig, _excel_connection_status, _last_excel_exception,
    _last_rows, _net_drawn, _pending_date_loads, _performance_metrics,
    _sheet_altered, _sheet_cache, active_expiries, all_known_expiries,
    net_badge, role_symbol, spot_cmp, user_pinned,
)
from us_chain import (
    _header_at, _set_header_at, _short_build, active_symbols,
    available_expiries, build_option_chain, ensure_sym_state, fmt_oi,
    cell_value_from_pre, iso_to_user, parse_user_date, print_chain,
    read_symbol_cell,
)
from us_spot import source_delay_note, spot_source_for
from us_workbook import (
    BLOCKS_PER_SHEET, BLOCK_COLS, EXPIRY_CELLS, default_k1, ensure_excel_file,
    excel_path, get_column_letter, is_locked_role, open_excel_files,
)
from excel_helpers import (
    long_form, refresh_delay,
)
from us_diagnostics import (
    append_app_exception, append_connection_event, excel_process_details, feed_log,
)
from us_performance import color_signature, flush_performance_metrics, itm_color_ranges
from us_licensing import ensure_license, license_is_valid, show_error
import asyncio
import calc_mode_guard
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
# Refreshed each cycle during which Excel could not answer its workbook list
# (dialog up, instance being torn down). Within the TTL the attach state is
# "unknown", not "missing".
_UNKNOWN_TTL_SECONDS = 10.0
_unknown_logged_at = {}


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


def pad_block_rows(values, extent):
    """Pad one block's rows with blank rows up to the fixed extent.

    The customer's strategy formulas carry literal row ranges. A chain that
    changes its live row count with every expiry would otherwise walk rows out
    from under those ranges; blank padding keeps the block's occupied range
    (5 .. 4+extent) constant so the same range always covers the live data.
    """
    if extent <= len(values):
        return values
    blank = [None] * (len(values[0]) if values else 9)
    return values + [list(blank) for _ in range(extent - len(values))]


# Blocks whose deep-clear check has already run in this session. A workbook
# written by an older build carries stray chain rows well below the live
# table (that build wrote as many rows as the feed delivered); the routine
# first-sight sweep only reaches _FIRST_SIGHT_BOUND rows, so once per block
# per run we look below the fixed extent and clear to the template depth
# only when something is actually there.
_deep_swept = set()


# Sheets whose unused expiry blocks were swept in this session. One sheet
# per workbook per run: anything older than a session ends up cleared again
# on the next one.
_retired_sheets = set()


def _retire_unused_blocks(sheet, role):
    """Wipe the expiry blocks a sheet does not show any more.

    The sheet shows BLOCKS_PER_SHEET expiry blocks. Every block beyond them
    (columns W..AE for a two-block sheet) is swept once per sheet object —
    labels, headers and any chain data an older build left there — and its
    expiry dropdown is removed, so the customer cannot pin an expiry into a
    block that never renders. Re-swept if Excel restarts mid-session with a
    fresh sheet holding old disk contents. The layout constants stay, so
    turning a block back on later needs only BLOCKS_PER_SHEET.
    """
    if (role, id(sheet)) in _retired_sheets:
        return
    _retired_sheets.add((role, id(sheet)))
    for block in BLOCK_COLS[BLOCKS_PER_SHEET:]:
        c0 = block["start"]
        try:
            sheet.range((1, c0), (2000, c0 + 8)).clear()
        except Exception as e:
            log_excel_exception(role, e, "clearing unused expiry block")
    unused = ("D3", "O3", "Z3")[BLOCKS_PER_SHEET:]
    for cell in unused:
        try:
            sheet.range(cell).api.Validation.Delete()
        except Exception as e:
            log_excel_exception(role, e, f"clearing expired input validation {cell}")
        try:
            sheet.range(cell).value = None
        except Exception as e:
            log_excel_exception(role, e, f"clearing expired input cell {cell}")
    feed_log("expiry_block_retired", role=role, shown=BLOCKS_PER_SHEET,
             cleared=[block["start"] for block in BLOCK_COLS[BLOCKS_PER_SHEET:]])


def _deep_clear_below_block(sheet, role, bi, c0, n_rows):
    """Clear junk an older build left below the fixed extent, if any.

    One read decides it; the clear only happens when the region is not
    already empty, so the normal case costs one cheap COM read per block
    per run.
    """
    if (role, bi) in _deep_swept:
        return
    _deep_swept.add((role, bi))
    deep_start = 5 + n_rows
    try:
        deep = sheet.range((deep_start, c0), (2000, c0 + 8)).value
    except Exception:
        return
    if not deep:
        return
    rows = deep if isinstance(deep, list) else [[deep]]
    if not any(any(cell is not None for cell in row) for row in rows):
        return
    dcs = get_column_letter(c0)
    dce = get_column_letter(c0 + 8)
    try:
        sheet.range(f"{dcs}{deep_start}:{dce}2000").clear()
        feed_log("chain_deep_clear", role=role, block=bi + 1, from_row=deep_start)
    except Exception as e:
        log_excel_exception(role, e, f"deep clearing stale rows in block {bi + 1}")


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
    if us_state._last_recalc_ms >= 50.0:
        title += f"  |  recalc {us_state._last_recalc_ms:.0f}ms"
    if last is not None and last[0] == title:
        return
    _last_window_status[key] = (title, now)
    try:
        book.api.Windows(1).Caption = title
    except Exception as e:
        log_excel_exception(role, e, "setting window status")


def _owner_alive():
    """Cheap liveness probe on the tool's own Excel instance."""
    app = us_state._owner_app
    if app is None:
        return False
    try:
        v = app.version
        return True
    except Exception as e:
        try:
            feed_log("owner_alive_failed", error=type(e).__name__, detail=str(e)[:180], hwnd=us_state._owner_hwnd)
        except Exception:
            pass
        return False


def excel_process_running(max_age=2.0):
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
    now = time.monotonic()
    if us_state._excel_liveness_cache is not None:
        stamp, running = us_state._excel_liveness_cache
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
        us_state._excel_liveness_cache = (now, True)
        return True
    running = "EXCEL.EXE" in (result.stdout or "").upper()
    us_state._excel_liveness_cache = (now, running)
    return running


def find_open_book(xw, fname):
    """Find the workbook inside the tool's PRIVATE Excel instance.

    The chain books live in the instance the tool started itself (the owner,
    see us_workbook._ensure_owner_app), so scanning every Excel on the
    machine is not only wasted work now - it was the loophole that let the
    two chain tools open workbooks in each other's Excel and fight over
    calculation mode. One chain tool, one Excel instance, no cross-search.

    When our instance is there but does not answer its workbook list (a
    dialog is up, the instance is being torn down), the failure is recorded
    as "unknown" and None returned: the caller must treat that as
    unanswerable this cycle, never as a missing workbook.
    """
    target = os.path.abspath(fname)
    role = next(
        (candidate for candidate, file_name in ROLE_FILES.items() if file_name.lower() in target.lower()),
        "unknown",
    )
    app = us_state._owner_app
    if app is None:
        return None
    try:
        books = list(app.books)
    except Exception as e:
        log_excel_exception(role, e, "enumerating the tool's own Excel workbooks")
        us_state._enum_unknown[role] = time.monotonic()
        return None
    try:
        target_form = long_form(target).lower()
    except Exception:
        target_form = target.lower()
    for book in books:
        try:
            candidate = long_form(os.path.abspath(book.fullname)).lower()
        except Exception as e:
            log_excel_exception(role, e, "reading open workbook path")
            continue
        if candidate == target_form:
            return book
    # DEBUG: log mismatch
    try:
        feed_log(
            "find_open_book_nomatch",
            role=role,
            target=target_form,
            books=";".join(long_form(os.path.abspath(b.fullname)).lower() for b in books),
        )
    except Exception:
        pass
    return None


_CLOSE_ALL_EXIT_SECONDS = 75.0


def maybe_exit_after_close_all(excel_alive=None):
    """Exit like every application does: when all its windows are gone.

    All three chain workbooks confirmed missing for a while means the
    customer does not want the chain right now. The old behaviour kept the
    exe alive and it then reopened workbooks behind the customer's back
    while they were running the other chain tool - a zombie that kept the
    popups alive too. The customer simply reruns the exe when they want the
    chain again.
    """
    if excel_alive is False:
        # No live Excel to write to at all: the private instance closes with
        # its last workbook, so this is independent confirmation.
        now = time.monotonic()
        for role in ROLE_ORDER:
            us_state._missing_since[role] = now
        try:
            feed_log("maybe_exit_excel_dead", hwnd=us_state._owner_hwnd, missing=list(us_state._missing_since.keys()))
        except Exception:
            pass
    if len(us_state._missing_since) == len(ROLE_ORDER):
        now = time.monotonic()
        if us_state._all_gone_since is None:
            us_state._all_gone_since = now
        elif now - us_state._all_gone_since >= _CLOSE_ALL_EXIT_SECONDS:
            feed_log(
                "tool_exit_close_all",
                note="every chain workbook is closed; the tool exits like any other app",
            )
            try:
                calc_mode_guard.release(PRODUCT_ID)
            except Exception:
                pass
            _heal_keyboard()
            os._exit(0)
    elif us_state._all_gone_since is not None:
        us_state._all_gone_since = None
        us_state._missing_since.clear()


_OPEN_RETRY_GRACE_SECONDS = 45.0

_EXCEL_GONE_WATCHDOG_EXIT_SECONDS = 75.0


def _excel_gone_tick(dead_since, alive, now=None):
    """One watchdog beat, kept pure so tests can drive every 75s outcome."""
    if alive:
        return None, False
    now = time.monotonic() if now is None else now
    if dead_since is None:
        return now, False
    return dead_since, now - dead_since >= _EXCEL_GONE_WATCHDOG_EXIT_SECONDS


def _excel_gone_watchdog(interval=5.0):
    """Exit the tool even when the asyncio loop is wedged on dead Excel.

    The display loop handles a closed Excel politely: books go missing, the
    close-all timer runs, exit <=75s. That path is cooperative - it only
    works while the loop still turns. A wedged COM retry can freeze the
    whole loop instead (feed silent while the process lingers). This thread
    never touches COM: it asks user32 whether our private Excel window is
    still a window. A destroyed hwnd means the instance was shot; after the
    same 75s grace the tool exits with the very same close-all receipt.
    """
    dead_since = None
    while True:
        time.sleep(interval)
        hwnd = us_state._owner_hwnd
        if not hwnd:
            dead_since = None
            continue
        try:
            import ctypes

            alive = bool(ctypes.windll.user32.IsWindow(int(hwnd)))
        except Exception:
            continue
        dead_since, exit_now = _excel_gone_tick(dead_since, alive)
        if exit_now:
            feed_log(
                "tool_exit_close_all",
                note="every chain workbook is closed; the tool exits like any other app",
            )
            try:
                calc_mode_guard.release(PRODUCT_ID)
            except Exception:
                pass
            _heal_keyboard()
            os._exit(0)


_excel_gone_watchdog_started = False


def start_excel_gone_watchdog():
    """Start the independent dead-Excel net exactly once."""
    global _excel_gone_watchdog_started
    if _excel_gone_watchdog_started:
        return
    _excel_gone_watchdog_started = True
    threading.Thread(
        target=_excel_gone_watchdog, name="excel-gone-watchdog", daemon=True
    ).start()


def _pump_com():
    """Keep the display thread's COM apartment alive between cycles.

    Excel proxies created by DispatchEx need their message queue pumped or
    the marshaled connection dies underneath us: the very next cycle then
    answers 'Object is not connected to server' for an Excel that is still
    running, every attach fails, and the close-all timer ends the tool with
    Excel right there on the screen. One cheap pump per display cycle keeps
    the proxy healthy for the whole session.
    """
    try:
        import pythoncom

        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def retry_open_pass_after_grace():
    """Retry a wedged startup open pass once, on the display thread.

    The pass runs inline at startup; if Excel answered every call it takes
    under a second and every role is latched. When the pass wedged (Excel
    busy coming up, a dead-quiet COM answer), nothing is latched yet - so
    after the grace period we run the pass another single time and then
    leave the books alone, exactly as the customer's never-reopen rule
    requires.
    """
    if us_state._open_pass_retried or us_state._open_pass_at is None:
        return
    now = time.monotonic()
    if now - us_state._open_pass_at < _OPEN_RETRY_GRACE_SECONDS:
        return
    us_state._open_pass_retried = True
    if all(us_state._auto_open_done.get(role) for role in ROLE_ORDER):
        return
    feed_log(
        "excel_open_retried",
        note=(f"nothing latched {int(_OPEN_RETRY_GRACE_SECONDS)}s after the startup "
              "open pass; retrying it once"),
    )
    open_excel_files()


def update_expiry_dropdowns(sym, sheet):
    # Date order: the chain the customer wants is almost always the next
    # expiry out, and their own broker lists them nearest first.
    dates = available_expiries(all_known_expiries[sym], sym, by_date=True)
    if not dates:
        return
    options = [iso_to_user(d) for d in dates]
    # Include today's date in the signature so the dropdown is forced to
    # re-sync when the calendar rolls over, and is sheet-aware so a symbol
    # used by multiple roles doesn't share a stale validation list.
    from us_time import eastern_now
    now, _ = eastern_now()
    today = now.strftime("%Y-%m-%d")
    sig_key = (sym, id(sheet))
    joined = f"{today}|{','.join(options)}"
    if _dropdown_sig.get(sig_key) == joined:
        return
    _dropdown_sig[sig_key] = joined
    # helper list in far-right column AH (avoids Excel 255-char inline limit)
    helper_col = 34  # AH
    helper_end = 100
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
    # Add one trailing blank cell to the formula to bust Excel's validation
    # cache: a stale list can persist even after the helper column changes.
    formula = f"=$AH$1:$AH${len(options) + 1}"
    for cell in EXPIRY_CELLS:
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


def _looks_like_excel_date_serial(val):
    """Excel serial 40909 = 1/Jan/2012, 49746 = 1/Jan/2024. Dates near now."""
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    return 40909 <= val <= 60000


def _excel_serial_to_date(serial):
    base = date(1899, 12, 30)
    return base + timedelta(days=int(serial))


def _purge_past_date_cell(sym, i, cell, rng, user_str):
    """Replace a past date in a cell with the first future expiry.
    Called before the _user_dates_seen check so old past dates are not
    adopted and frozen."""
    from us_time import eastern_now
    now, _ = eastern_now()
    today = now.strftime("%Y-%m-%d")
    parsed = parse_user_date(user_str)
    if not parsed or parsed >= today:
        return
    dates = available_expiries(all_known_expiries[sym], sym, by_date=True)
    replacement = dates[0] if dates else None
    if not replacement:
        return
    try:
        rng.value = iso_to_user(replacement)
        rng.color = (255, 255, 153)
    except Exception:
        pass
    active_expiries[sym][i] = replacement
    user_pinned[sym][i] = False
    us_state._user_dates_seen.pop((sym, i), None)
    print(f"[{sym}] purge_past: {parsed} -> {replacement} in {cell}", flush=True)
    return


def update_user_dates(sym, sheet, pre=None):
    for i, cell in enumerate(EXPIRY_CELLS):
        rng = sheet.range(cell)
        try:
            # text — stops Excel date conversion / #####
            # Set the text format only when it is not already "@": a set
            # dirties the cell and an identical format every cycle is pure
            # waste. Excel resets the format when the user picks from the
            # dropdown; the next pass then sees General and re-applies once.
            if rng.api.NumberFormat != "@":
                rng.api.NumberFormat = "@"
        except Exception:
            pass
        val = cell_value_from_pre(pre, cell) if pre is not None else rng.value
        # Excel may have converted text into a datetime or a serial number
        if isinstance(val, datetime):
            try:
                rng.value = val.strftime("%d/%m/%Y")
            except Exception:
                pass
            continue
        # Excel serial number (e.g. 46123 = 23/10/2026). Convert back to text.
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            if _looks_like_excel_date_serial(val):
                try:
                    d = _excel_serial_to_date(val)
                    rng.value = d.strftime("%d/%m/%Y")
                except Exception:
                    pass
                continue
        if val is None or (isinstance(val, str) and val.strip() == ""):
            from us_time import eastern_now
            now, _ = eastern_now()
            today = now.strftime("%Y-%m-%d")
            ae = active_expiries[sym][i]
            if ae and ae < today:
                dates = available_expiries(all_known_expiries[sym], sym, by_date=True)
                ae = dates[0] if dates else None
                active_expiries[sym][i] = ae
                user_pinned[sym][i] = False
            rng.value = iso_to_user(ae) if ae else ""
            user_pinned[sym][i] = False
            continue
        user_str = str(val).strip()
        user_up = user_str.upper()
        if user_up in ("EXPIRY", "DD/MM/YYYY", "DDMMMYY", "EDIT"):
            from us_time import eastern_now
            now, _ = eastern_now()
            today = now.strftime("%Y-%m-%d")
            ae = active_expiries[sym][i]
            if ae and ae < today:
                dates = available_expiries(all_known_expiries[sym], sym, by_date=True)
                ae = dates[0] if dates else None
                active_expiries[sym][i] = ae
            rng.value = iso_to_user(ae) if ae else ""
            user_pinned[sym][i] = False
            continue
        # A value the tool itself wrote is not a user edit.
        #
        # The workbook keeps whatever expiry was in the cell from the previous
        # session. When the tool writes a new expiry there and then reads the
        # cell back, the old value looked like the user had chosen it, and the
        # slot was pinned to it for good. A pinned slot is never changed, so
        # the three blocks sat on a monthly expiry and a daily expiry and a
        # weekly one, and no amount of waiting would move them.
        written = us_state._tool_written_dates.get((sym, i))
        if written is not None and user_str == str(written).strip():
            continue
        # Purge past dates BEFORE the _user_dates_seen check, so a leftover
        # past date from the previous session is not adopted as the user's
        # choice and frozen forever.
        _purge_past_date_cell(sym, i, cell, rng, user_str)
        # Only a cell that CHANGED since the tool last looked at it was edited
        # by somebody. The first look just adopts whatever is there, so a
        # leftover from the previous session is never mistaken for a choice.
        seen_key = (sym, i)
        if seen_key not in us_state._user_dates_seen:
            us_state._user_dates_seen[seen_key] = user_str
            continue
        changed_by_user = us_state._user_dates_seen.get(seen_key) != user_str
        us_state._user_dates_seen[seen_key] = user_str
        parsed = parse_user_date(user_str)
        if parsed:
            from us_time import eastern_now
            now, _ = eastern_now()
            today = now.strftime("%Y-%m-%d")
            if parsed < today:
                # Replace past dates with the first future expiry, not just flag.
                dates = available_expiries(all_known_expiries[sym], sym, by_date=True)
                replacement = dates[0] if dates else None
                if replacement:
                    try:
                        rng.value = iso_to_user(replacement)
                        rng.color = (255, 255, 153)
                    except Exception:
                        pass
                    active_expiries[sym][i] = replacement
                    user_pinned[sym][i] = False
                    us_state._user_dates_seen.pop((sym, i), None)
                    print(f"[{sym}] past expiry {parsed} replaced with {replacement}", flush=True)
                continue
        if parsed and parsed != active_expiries[sym][i] and changed_by_user:
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
# Last written values per (role, block), so only the rows that actually moved
# are written back. Writing the whole block dirties every cell in it, and
# Excel repaints the dirty region when the write window closes - that is what
# made scrolling stutter.
_block_values = {}
# Cells whose number format has already been set. Setting a number format
# dirties the cell, and doing that on every refresh is wasted work.
_user_fmt_done = set()
# Last applied ITM/ATM colouring per (role, block). Used to skip repainting
# when the picture has not changed, so a live price tick does not dirty every
# cell of the block and make Excel recalculate the customer's own formulas.
_colour_state = {}
# Rows this block used on its previous cycle, so a shrink can clear exactly
# what was there instead of wiping a fixed block of rows every tick.


def _screen_pause_active():
    """True while the screen-updating attempt is paused after failures."""
    return us_state._screen_blocked_until > time.monotonic()


def _note_screen_ok():
    us_state._screen_failures = 0
    us_state._screen_blocked_until = 0.0


def _note_screen_failed(exc):
    us_state._screen_failures += 1
    if us_state._screen_failures >= _SCREEN_FAILURE_LIMIT:
        us_state._screen_blocked_until = time.monotonic() + _SCREEN_BACKOFF_SECONDS
        feed_log(
            "screen_updating_paused",
            failures=us_state._screen_failures,
            seconds=_SCREEN_BACKOFF_SECONDS,
            error=type(exc).__name__,
        )


def _screen_update_watchdog_tick():
    """Force ScreenUpdating back on if this instance is ever seen with it off.

    The tool itself no longer toggles ScreenUpdating (flipping it around the
    write window repainted the whole sheet on every release and froze edits
    that started mid-window). What remains possible is a screen left off by
    an older build still running while Excel reopens, or a third-party
    add-in. This tick, run once per display cycle while no write window is
    open, is the net that guarantees paint against that class. The
    screen-failure backoff still applies so it never hammers a busy
    instance.
    """
    app = us_state._app_cache or us_state._owner_app
    if app is None:
        return
    if us_state._bulk_window_open or _screen_pause_active():
        return
    try:
        busy = app.screen_updating
    except Exception:
        return
    if busy:
        us_state._screen_restore_owed = False
        return
    try:
        app.screen_updating = True
    except Exception as e:
        log_excel_exception("ALL", e, "screen watchdog force-on")
        return
    us_state._screen_restore_owed = False
    now = time.monotonic()
    if now - us_state._screen_watchdog_logged_at >= 60.0:
        us_state._screen_watchdog_logged_at = now
        feed_log(
            "screen_updating_forced_on",
            note="watchdog restored paint after a failed in-window restore",
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
# J1 is role-dependent — written inside write_hint_labels
HINT_LABELS = (
    ("B2", "± strikes — edit A2"),
    ("A3", "Expiry 1"),
    ("C3", "(edit D3)"),
    ("L3", "Expiry 2"),
    ("N3", "(edit O3)"),
)

# Blocks beyond BLOCKS_PER_SHEET are retired: labels, headers and data in
# their columns are swept once per session (see _retire_unused_blocks), so
# only Expiry 1/2 hints exist here.
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
        for block in BLOCK_COLS[:BLOCKS_PER_SHEET]:
            sheet.range((1, block["start"]), (1, block["start"] + 4)).font.bold = True
            sheet.range((1, block["start"]), (1, block["start"] + 4)).font.size = 12
        # Row 4 headers
        for block in BLOCK_COLS[:BLOCKS_PER_SHEET]:
            hdr = sheet.range((4, block["start"]), (4, block["start"] + 8))
            hdr.font.bold = True
            hdr.font.color = (255, 255, 255)
            hdr.color = HEADER_FILL
        # Labels + hints
        for cell, _t in HINT_LABELS:
            sheet.range(cell).font.bold = True
        for cell in ("B2", "A3", "L3", "C3", "N3", "J1"):
            sheet.range(cell).color = LABEL_FILL
        # Input cells highlighted (do NOT wipe user values)
        # Locked files (SPX/DJX): K1 gray + forced — not an input
        input_cells = ["A2", *EXPIRY_CELLS]
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
        # Stock role: K1 is user-editable — never write its value here.
        # The initial value is set by _write_template when the file is created.
        for cell in EXPIRY_CELLS:
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
    """Prices = 0.00, Strike = 0.0, OI = numeric with $ display.

    All columns are written as numbers by build_option_chain (fmt_oi returns
    float), so every column gets a numeric NumberFormat. OI uses "$#,##0" so it
    displays as a dollar amount but stays numeric — this lets the customer
    write XLOOKUP formulas that add OI values without hitting text+number
    #VALUE! errors.

    The one-off bulk cleanup below fixes up any cells left as numbers or
    text with ₹ markers by an older build. One read and one write for the
    whole column, once per sheet.
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
                sheet.range(f"{col}5:{col}{end}").api.NumberFormat = "$#,##0"
            _usd_fmt_done.add(fmt_key)
        except Exception as e:
            print(f"force_usd formats: {e}", flush=True)
            log_excel_exception("EXCEL", e, "formatting option-chain cells")
            return
        # One-off bulk cleanup of values left as numbers or ₹ by an older
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
                elif isinstance(v, str):
                    s = v.strip()
                    if s and (
                        s.startswith("$")
                        or "₹" in v
                        or "Rs" in v
                        or "INR" in v
                        or "," in s
                        or "M" in s
                        or "K" in s
                    # Plain numerals in a text cell ("223") are leftovers
                    # from the old text-format pass - the OI columns hold
                    # no words, so any string left there can be turned
                    # into its number.
                        or s.lstrip("+-").replace(".", "").isdigit()
                    ):
                        values[i] = fmt_oi(v)
                        changed = True
            if changed:
                try:
                    rng.value = values
                except Exception as e:
                    log_excel_exception("EXCEL", e, f"writing open-interest column {c0 + off}")


def _find_chain_sheet(book):
    """Return the sheet this product writes to, by NAME not by index.

    ``book.sheets[0]`` is the sheet Excel reports as first — which is the
    customer's active/new sheet when they insert one, not necessarily the
    "Chain" sheet this tool owns. Writing there overwrites the customer's
    data with the header band (e.g. "Expiry 1" scattered across A1:N3 of a
    brand-new sheet). Always look the owned sheet up by name.
    """
    try:
        return book.sheets["Chain"]
    except Exception:
        pass
    for sh in book.sheets:
        try:
            if sh.name == "Chain":
                return sh
        except Exception:
            continue
    return book.sheets[0]


def _write_to_excel(role):
    try:
        import xlwings as xw
    except ImportError:
        show_error("Required component 'xlwings' is missing.")
        _heal_keyboard()
        os._exit(1)
    path = ensure_excel_file(role)
    try:
        # Cached COM book/sheet objects — avoids enumerating all apps + all
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
            record_excel_connection(
                role,
                "own_excel_running_workbook_not_found" if _owner_alive() else "excel_not_found",
                f"expected_workbook={path}",
            )
            now = time.monotonic()
            if now - us_state._enum_unknown.get(role, 0.0) < _UNKNOWN_TTL_SECONDS:
                # Excel did not answer its workbook list this cycle (a dialog
                # is up or the instance is being torn down). The truth is
                # "unknown", never "missing": the session is kept and this
                # cycle skipped. This is what ends the workbook-missing modal
                # that used to refire every couple of minutes - a busy Excel
                # is no longer an error state.
                if excel_connected[role]:
                    logged_at = _unknown_logged_at.get(role, 0.0)
                    if now - logged_at > _UNKNOWN_TTL_SECONDS:
                        _unknown_logged_at[role] = now
                        feed_log(
                            "excel_attach_unknown",
                            role=role,
                            note="Excel did not answer this cycle; keeping the session",
                        )
                return
            # A confirmed absence: the customer closed this workbook, and it
            # stays closed. No reopen, no modal - just a log line and a kept
            # hole until the workbook is open again or the tool exits.
            if excel_connected[role]:
                excel_connected[role] = False
                _sheet_cache.pop(role, None)
                feed_log(
                    "workbook_closed_by_customer",
                    role=role,
                    note="the chain workbook was closed; it stays closed",
                )
            us_state._missing_since[role] = now
            return
        excel_miss_counts[role] = 0
        us_state._enum_unknown.pop(role, None)
        _unknown_logged_at.pop(role, None)
        us_state._missing_since.pop(role, None)
        excel_connected[role] = True
        show_status(book, role, "LIVE")
        if _excel_connection_status.get(role) != "connected":
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
            sheet = _find_chain_sheet(book)
            _sheet_cache[role] = sheet

        # Retire the third expiry block the moment the sheet is reachable —
        # not only at the first data write — so the customer never stares at
        # the previous session's Block 3 through the cold-start window.
        _retire_unused_blocks(sheet, role)

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

        try:
            pre = sheet.range((1, 1), (4, 31)).value
        except Exception as e:
            log_excel_exception(role, e, "reading the header block")
            pre = None
        if pre is not None and (not pre or not isinstance(pre[0], (list, tuple))):
            pre = None
        header_values = list(pre[0]) if pre else []

        new_sym = read_symbol_cell(sheet, role, pre=pre)
        if new_sym and new_sym != role_symbol[role]:
            print(f"[{role}] symbol {role_symbol[role]} -> {new_sym}", flush=True)
            role_symbol[role] = new_sym
            _pending_date_loads.clear()
            _dropdown_sig.clear()
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

        v = cell_value_from_pre(pre, "A2") if pre is not None else sheet.range("A2").value
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
            update_user_dates(sym, sheet, pre=pre)
        except Exception as e:
            print(f"update_user_dates ({sym}): {e}", flush=True)
            log_excel_exception(role, e, "reading expiry input cells")
        try:
            update_expiry_dropdowns(sym, sheet)
        except Exception as e:
            _dropdown_sig.clear()
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

        # Blocks of this sheet whose data changed and still has to be written.
        pending_writes = []

        _retire_unused_blocks(sheet, role)

        for bi, block in enumerate(BLOCK_COLS[:BLOCKS_PER_SHEET]):
            expiry = active_expiries[sym][bi]
            col_l = block["start"]
            # The cell above the block names the expiry that block shows, so
            # it has to follow the block.
            #
            # It used to be written only when the expiry ROLLED, which is
            # once a minute, and never at all when a slot was first assigned.
            # So the blocks could show dailies while the cells above them
            # still said older months - and the customer had no way to tell
            # which chain they were looking at.
            if expiry and not user_pinned[sym][bi]:
                try:
                    text = iso_to_user(expiry)
                    if us_state._expiry_cell_written.get((sym, bi)) != text:
                        cell_ref = EXPIRY_CELLS[bi]
                        rng = sheet.range(cell_ref)
                        rng.api.NumberFormat = "@"
                        rng.value = text
                        us_state._expiry_cell_written[(sym, bi)] = text
                        us_state._tool_written_dates[(sym, bi)] = text
                        us_state._user_dates_seen[(sym, bi)] = text
                except Exception as e:
                    log_excel_exception(role, e, "writing expiry date cell")
            if not expiry:
                # Only rewrite header cells when the text actually changed.
                # Writing an identical string still dirties the cell and forces
                # every dependent formula to recalculate — that cost is what
                # makes a strategy workbook feel stuck.
                head_text = f"{sym} Spot: --"
                pcr_text = "PCR: --"
                try:
                    if _header_at(header_values, col_l) != head_text:
                        sheet.range((1, col_l)).value = head_text
                        _set_header_at(header_values, col_l, head_text)
                        us_state._cycle_wrote = True
                    if _header_at(header_values, col_l + 4) != pcr_text:
                        sheet.range((1, col_l + 4)).value = pcr_text
                        _set_header_at(header_values, col_l + 4, pcr_text)
                        us_state._cycle_wrote = True
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
            # Never let a delay go unannounced. A silent drop to a delayed
            # feed means the trader is reading old prices, and that is
            # exactly what "the data is stale" means to them.
            _note = source_delay_note(sym)
            if _note:
                _head = f"{_head} | {_note}"
            pcr_text = f"PCR: {pcr:.2f}"
            try:
                if _header_at(header_values, col_l) != _head:
                    sheet.range((1, col_l)).value = _head
                    _set_header_at(header_values, col_l, _head)
                    us_state._cycle_wrote = True
                if _header_at(header_values, col_l + 4) != pcr_text:
                    sheet.range((1, col_l + 4)).value = pcr_text
                    _set_header_at(header_values, col_l + 4, pcr_text)
                    us_state._cycle_wrote = True
            except Exception as e:
                log_excel_exception(role, e, "writing chain header")
            net_paint(sheet, role, bi, col_l)

            if chain is not None and not chain.empty:
                # OI columns stay numeric: build_option_chain writes floats
                # and force_usd_block_formats formats them "$#,##0", so the
                # customer's XLOOKUP arithmetic sees plain numbers. The old
                # write path re-armed them to text format ("@") whenever the
                # block grew a row, which fought that numeric format and
                # nudged text into their formulas - retired for good.
                c0 = block["start"]

                write_started = time.perf_counter()
                try:
                    # Only write when the contents actually differ. Any data
                    # change alters the signature and forces the write, so live
                    # updates still land; unchanged blocks cost nothing.
                    values = chain.values.tolist()
                    # The block always occupies the same rows. Live strikes
                    # fill from the top; rows past the live count stay blank,
                    # so a customer range written for a 37-strike expiry
                    # (P5:P41) keeps covering the whole chain when the next
                    # expiry runs to 60 strikes.
                    # The extent never shrinks: the highest rows this session
                    # has seen for the block stay the floor, so a shorter
                    # expiry cannot yank rows out from under the customer's
                    # offset ranges.
                    seen = us_state._row_extent_seen.get((role, bi), 0)
                    extent = max(us_state.ROW_EXTENT_DEFAULT, seen, len(values))
                    if extent != seen:
                        if seen:
                            feed_log(
                                "chain_extent_changed",
                                role=role,
                                block=bi + 1,
                                from_rows=seen,
                                to_rows=extent,
                            )
                        us_state._row_extent_seen[(role, bi)] = extent
                    values = pad_block_rows(values, extent)
                    # Sizes downstream (stale clear, formats, colour paints)
                    # all key off the extent so the geometry stays constant.
                    n_rows = extent
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
                        us_state._cycle_wrote = True
                        # The Excel write is deferred to the end of the block
                        # loop, so all of this sheet's blocks go out together.
                        pending_writes.append((bi, block, values, n_rows, runs, prev))
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
                # Only clears rows that can hold leftovers: what this block
                # used previously (below the fixed extent that only ever
                # grows, so only possible on the first cycle this block is
                # seen), plus on first sighting a generous sweep to catch
                # anything an earlier run left behind. A live shrink is
                # handled by the padded payload diff instead: the rows the
                # block let go of are written back blank as blanks.
                prev_n = _last_rows.get((role, bi))
                if prev_n is None or n_rows < prev_n:
                    first_sight = _FIRST_SIGHT_BOUND if prev_n is None else 0
                    last_row = max(n_rows, prev_n or 0, first_sight)
                    if last_row > n_rows:
                        # A reversed range (rows n..n-1) resolves to its
                        # top-left and clips the wrong row, so only clear
                        # when there is something below the live rows.
                        stale_a = (
                            f"{get_column_letter(c0)}{5 + n_rows}"
                            f":{get_column_letter(c0 + 8)}{4 + last_row}"
                        )
                        try:
                            sheet.range(stale_a).clear()
                        except Exception as e:
                            log_excel_exception(role, e, f"clearing stale rows in block {bi + 1}")
                _last_rows[(role, bi)] = n_rows
                _deep_clear_below_block(sheet, role, bi, c0, n_rows)

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
      customer's own workbook. That is the case that made their sheet crawl.
      Scattered runs are merged into one span-write whose unchanged rows
      carry the previous values, so the dirty set stays small in Excel's
      eyes while the number of COM calls stays at one.

    Either way each block is written on its own, into its own columns, so a
    block that did not change is never touched.

    Pending entries carry the block's padded rows (values padded with blank
    rows up to the fixed extent, never the raw DataFrame), so even a
    whole-block write keeps the geometry the customer's literal ranges rely
    on.
    """
    if not pending:
        return
    started = time.perf_counter()
    protected = _calc_session.get("owned") and not _calc_session.get("failed")
    try:
        for bi, block, row_block, n_rows, runs, prev in pending:
            try:
                if protected:
                    if len(runs) == 1:
                        start, end = runs[0]
                        sheet.range((5 + start, block["start"])).value = [
                            list(r) for r in row_block[start : end + 1]
                        ]
                        _performance_metrics.record(
                            role, "excel_rows_written", 0.0, rows=end - start + 1
                        )
                    else:
                        sheet.range((5, block["start"])).options(
                            index=False, header=False
                        ).value = [list(r) for r in row_block]
                else:
                    # Unprotected window: every cell can trigger a recalc of
                    # the customer's workbook, so the write window is one
                    # merged span instead of one call per contiguous run.
                    if runs and len(runs) <= 8:
                        lo = runs[0][0]
                        hi = runs[-1][1]
                        merged_rows = []
                        for i in range(lo, hi + 1):
                            in_run = any(start <= i <= end for start, end in runs)
                            if in_run:
                                merged_rows.append(list(row_block[i]))
                            elif prev is not None and i < len(prev):
                                merged_rows.append(list(prev[i]))
                            else:
                                merged_rows.append(list(row_block[i]))
                        sheet.range((5 + lo, block["start"])).value = merged_rows
                        _performance_metrics.record(
                            role, "excel_rows_written", 0.0, rows=hi - lo + 1
                        )
                    else:
                        sheet.range((5, block["start"])).options(
                            index=False, header=False
                        ).value = [list(r) for r in row_block]
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


def _update_recalc_backoff(elapsed_ms):
    """Feed the real cost of Application.Calculate into the adaptive refresh.

    The CalculationState probe (_probe_calc_state_after_restore) only ever reads
    0 in practice - it runs after a synchronous Calculate that already finished
    - so it never stretched the loop. Use the measured recalc duration instead:
    a workbook heavy enough to cross the threshold earns idle time proportional
    to that cost; a light workbook stays at the configured refresh (no behaviour
    change for a light load).
    """
    secs = float(elapsed_ms) / 1000.0
    if secs <= us_state.RECALC_BACKOFF_THRESHOLD_SECONDS:
        return
    alpha = us_state.RECALC_BACKOFF_ALPHA
    us_state._recalc_backoff_seconds = (
        alpha * secs + (1.0 - alpha) * us_state._recalc_backoff_seconds
    )


def _time_recalc(wrote):
    """Recalculate once inside the manual window and time it.

    This is the real cost the customer feels: how long their own strategy
    workbook takes to recalc when we change a cell it depends on. Plain
    Application.Calculate, not CalculateUntilAsyncQueriesDone (that is for
    asynchronous queries such as Power Query). Only done when we actually
    wrote something and Excel was on auto/semiauto — recalculating for a
    customer who chose manual would undo their choice. A cycle skipped for
    customer input owes its recalculation, so _recalc_owed carries it here:
    the handwritten spreadsheet the customer was working in gets its fresh
    values on the very next cycle we can actually run.
    """
    global _recalc_input_skip_logged_at, _recalc_strategy_skip_logged_at
    owed = us_state._recalc_owed
    us_state._recalc_owed = False
    if not wrote and not owed:
        _performance_metrics.record("ALL", "excel_recalc_skip_nowrite", 0.0)
        return 0.0
    if not us_state._bulk_will_recalc:
        _performance_metrics.record("ALL", "excel_recalc_skip_mode", 0.0)
        return 0.0
    app = us_state._app_cache
    if app is None:
        _performance_metrics.record("ALL", "excel_recalc_skip_noapp", 0.0)
        if owed:
            # Still owed - Excel comes back and the debt is paid then.
            us_state._recalc_owed = True
        return 0.0
    # Last-instant guard: the trader may have started typing after the
    # pre-window check but before this Calculate. Force-calculating now is
    # exactly what froze the edit frame, so carry the debt instead.
    if _input_is_recent() or _excel_in_edit_mode(app):
        us_state._recalc_owed = True
        now = time.monotonic()
        if now - _recalc_input_skip_logged_at > 30.0:
            _recalc_input_skip_logged_at = now
            feed_log(
                "excel_recalc_skipped_user_input",
                note="customer's hands are on Excel; recalculation carried",
            )
        _performance_metrics.record("ALL", "excel_recalc_skip_user_input", 0.0)
        return 0.0
    # Strategy-safe (Option C): chain data keeps updating every cycle; the heavy
    # part is the full-workbook Calculate. Run it live ONLY when the last
    # measured Calculate cost is known and below STRATEGY_RECALC_MAX_MS; otherwise
    # carry the debt (Option B behaviour). This gives live P&L on cheap workbooks
    # while self-protecting on heavy ones.
    try:
        _on_strategy_sheet = (_active_sheet_is_chain(app) is False)
    except Exception:
        _on_strategy_sheet = False
    if _on_strategy_sheet:
        _last_ms = us_state._last_recalc_ms
        _threshold = us_state.STRATEGY_RECALC_MAX_MS
        if not (_last_ms > 0.0 and _last_ms < _threshold):
            now = time.monotonic()
            if now - _recalc_strategy_skip_logged_at > 30.0:
                _recalc_strategy_skip_logged_at = now
                feed_log(
                    "excel_recalc_skipped_strategy_sheet",
                    note="active sheet is not a chain sheet; recalculation carried",
                )
            us_state._recalc_owed = True
            _performance_metrics.record("ALL", "excel_recalc_skipped_strategy_sheet", 0.0)
            return 0.0
    try:
        started = time.perf_counter()
        app.api.Calculate()
        elapsed = (time.perf_counter() - started) * 1000.0
    except Exception as e:
        log_excel_exception("ALL", e, "forcing workbook recalculation")
        if owed:
            # A busy Excel must not silently eat a recalculation the skip
            # promised; retry on the next free cycle.
            us_state._recalc_owed = True
        return 0.0
    us_state._last_recalc_ms = elapsed
    _update_recalc_backoff(elapsed)
    _performance_metrics.record("ALL", "excel_recalc", elapsed / 1000.0)
    return elapsed


# Any keystroke or click inside this window counts as "hands on the keys".
# Config-driven so a slower typist can be covered without a code change.
_USER_INPUT_GRACE_SECONDS = us_state.INPUT_GRACE_SECONDS

_recalc_input_skip_logged_at = 0.0
_recalc_strategy_skip_logged_at = 0.0
_watchdog_input_skip_logged_at = 0.0
_session_release_input_skip_logged_at = 0.0


def _last_user_input_age():
    """Seconds since the last input event of any kind, or None."""
    import ctypes

    class _LASTINPUTINFO(ctypes.Structure):
        _fields_ = (
            ("cbSize", ctypes.c_uint),
            ("dwTime", ctypes.c_ulong),
        )

    info = _LASTINPUTINFO()
    info.cbSize = ctypes.sizeof(info)
    try:
        if not ctypes.windll.user32.GetLastInputInfo(ctypes.byref(info)):
            return None
        tick = ctypes.windll.kernel32.GetTickCount()
        return max(0.0, (tick - int(info.dwTime)) / 1000.0)
    except Exception:
        return None


def _input_is_recent():
    """True while any keyboard/mouse input happened inside the grace window."""
    age = _last_user_input_age()
    return age is not None and age < _USER_INPUT_GRACE_SECONDS


def _foreground_pid():
    """Process id owning the current foreground window, or None."""
    import ctypes
    from ctypes import wintypes

    try:
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return None
        pid = wintypes.DWORD(0)
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value or None
    except Exception:
        return None


def _owner_window_pid():
    """Process id of our private Excel instance, or None."""
    import ctypes
    from ctypes import wintypes

    hwnd = us_state._owner_hwnd
    try:
        if hwnd:
            pid = wintypes.DWORD(0)
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value:
                return pid.value
    except Exception:
        pass
    app = us_state._owner_app
    if app is None:
        return None
    try:
        return int(app.pid)
    except Exception:
        return None


def _user_is_excel_interacting():
    """True while the customer's hands are on OUR Excel this very moment.

    Any keystroke or click within the last ~0.7 s counts as active input,
    but only counts against the chain when the foreground window belongs to
    our own Excel instance: typing an email or working in some other
    workbook never pauses the chain. A yielded cycle skips the write
    window and the forced recalculation alike, and the missed
    recalculation is carried over so their formulas stay live.
    """
    if platform.system() != "Windows":
        return False
    try:
        age = _last_user_input_age()
        if age is None or age > _USER_INPUT_GRACE_SECONDS:
            return False
        pid = _foreground_pid()
        owner = _owner_window_pid()
        if pid is None or owner is None:
            return False
        return pid == owner
    except Exception:
        return False


def _excel_focus_is_edit_control(app):
    """True when the cell/formula-bar edit control owns Excel's focus.

    Application.Ready stays True during in-cell editing, so it cannot tell
    "trader is typing a formula" from "Excel is idle". The Win32 focus class
    can: the grid is EXCEL7, while the live edit control is EXCEL6.
    """
    if platform.system() != "Windows":
        return False
    import ctypes
    from ctypes import wintypes

    try:
        hwnd = int(app.api.Hwnd)
    except Exception:
        return False
    if not hwnd:
        return False
    try:
        user32 = ctypes.windll.user32
        user32.GetWindowThreadProcessId.argtypes = [
            wintypes.HWND, ctypes.POINTER(wintypes.DWORD)
        ]
        user32.GetWindowThreadProcessId.restype = wintypes.DWORD
        user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.c_void_p]
        user32.GetGUIThreadInfo.restype = wintypes.BOOL
        user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

        class _GUITHREADINFO(ctypes.Structure):
            _fields_ = (
                ("cbSize", wintypes.DWORD),
                ("flags", wintypes.DWORD),
                ("hwndActive", wintypes.HWND),
                ("hwndFocus", wintypes.HWND),
                ("hwndCapture", wintypes.HWND),
                ("hwndMenuOwner", wintypes.HWND),
                ("hwndMoveSize", wintypes.HWND),
                ("hwndCaret", wintypes.HWND),
                ("rcCaret", wintypes.RECT),
            )

        pid = wintypes.DWORD()
        tid = user32.GetWindowThreadProcessId(
            wintypes.HWND(hwnd), ctypes.byref(pid)
        )
        if not tid:
            return False
        info = _GUITHREADINFO()
        info.cbSize = ctypes.sizeof(info)
        if not user32.GetGUIThreadInfo(tid, ctypes.byref(info)):
            return False
        focus = info.hwndFocus
        if not focus:
            return False
        buf = ctypes.create_unicode_buffer(64)
        user32.GetClassNameW(focus, buf, 64)
        cls = buf.value.upper()
        return cls.startswith("EXCEL6") or cls == "EXCEL<"
    except Exception:
        return False


def _excel_in_edit_mode(app):
    """Detect whether the customer is actively editing Excel.

    Returns True when Excel is in a state where a write cycle would deadlock:
    - CutCopyMode is non-zero (drag-copy in progress)
    - the live EXCEL6 edit control owns focus (in-cell/formula-bar edit)
    - Application.Ready is False (a modal state)

    If the COM call itself fails or raises, treat as edit mode and skip.
    """
    if app is None:
        return False
    try:
        if app.api.CutCopyMode:
            return True
    except Exception:
        return True
    if _excel_focus_is_edit_control(app):
        return True
    try:
        if not app.api.Ready:
            return True
    except Exception:
        return True
    return False


def render_cycle():
    """One full pass over every workbook, inside a single bulk-mode window.

    Screen updating and calculation are suspended once per cycle rather than
    once per book, so a cycle costs three writes with repaint and recalculation
    suspended, then one restore. Returns nothing; timing is recorded in
    _performance_metrics.
    """
    started = time.perf_counter()
    us_state._cycle_wrote = False

    # A rejected Databento key cannot fix itself, so no chain can be
    # downloaded and the sheets will stay empty until the customer puts a
    # working key in. Say so once, plainly: a blank sheet whose header says
    # LOADING forever is indistinguishable from a broken app, and the
    # customer has no way to guess that the key is the problem.
    feed = us_state._databento_feed
    if feed is not None and not us_state._databento_auth_notice_shown:
        try:
            reason = feed.auth_error()
        except Exception:
            reason = None
        if reason:
            us_state._databento_auth_notice_shown = True
            try:
                show_error(
                    "Databento rejected your API key, so no option chains can "
                    "be downloaded and the sheets will stay empty.\n\n"
                    f"Details: {reason}\n\n"
                    "Enter a working key when the tool asks for it, then run "
                    "the tool again."
                )
            except Exception:
                pass

    # If Excel would not switch to manual calculation, every cell we write
    # makes it recalculate. Writing anyway at full speed is what made the
    # sheet feel stuck, so slow down to a rate the customer's workbook can
    # absorb, and tell them once what to clear.
    #
    # A workbook close is the one refusal that must never reach them: closing
    # briefly puts Excel's save prompt up and every switch fails for those
    # seconds. The marks are cleared the moment the chains are gone.
    if us_state._calc_manual_failed:
        _clear_refused_calc_if_chain_gone()
    if us_state._calc_manual_failed:
        # Don't show the slow-down popup when Excel is gone — the customer
        # closed it, not a busy cell or add-in.
        # If Excel is closing or already gone, the refusal is a teardown
        # artifact, not a real busy cell or add-in. Clear the marks so the
        # slow-down popup never fires on close, and skip this cycle.
        try:
            import ctypes
            _hwnd = us_state._owner_hwnd
            _closing = False
            if _hwnd and not ctypes.windll.user32.IsWindow(int(_hwnd)):
                _closing = True
            if len(getattr(us_state, "_missing_since", {})) == len(us_state.ROLE_ORDER):
                _closing = True
            if getattr(us_state, "_all_gone_since", None) is not None:
                _closing = True
            try:
                if not excel_process_running():
                    _closing = True
            except Exception:
                pass
            if _closing:
                _clear_refused_calc_marks()
        except Exception:
            pass
        if not us_state._calc_manual_failed:
            _performance_metrics.record("ALL", "render_cycle_skipped_noexcel", 0.0)
            return

        owner_app = us_state._owner_app
        if owner_app is None:
            # No owner reference at all — Excel is gone.
            _clear_refused_calc_marks()
        else:
            try:
                _ = owner_app.version  # COM liveness probe
            except Exception:
                # COM is dead but process may still be winding down.
                _clear_refused_calc_marks()
            else:
                # Owner COM is alive — check process table too.
                try:
                    if not excel_process_running():
                        _clear_refused_calc_marks()
                except Exception:
                    pass
        if not us_state._calc_manual_failed:
            _performance_metrics.record("ALL", "render_cycle_skipped_noexcel", 0.0)
            return
        now = time.monotonic()
        if now - us_state._last_unprotected_write < _UNPROTECTED_WRITE_INTERVAL:
            _performance_metrics.record("ALL", "render_cycle_skipped_busy", 0.0)
            return
        us_state._last_unprotected_write = now
        _performance_metrics.record("ALL", "render_cycle_unprotected", 0.0)
        if (
            not us_state._calc_manual_notice_shown
            and now - us_state._calc_failed_since >= _CALC_NOTICE_GRACE_SECONDS
        ):
            us_state._calc_manual_notice_shown = True
            feed_log(
                "calc_manual_unavailable_notice",
                note="Excel refused manual calculation; running in auto-recalc (slower) until Excel accepts the switch",
            )

    # Input yield: never open a write window while the customer's hands are
    # on our Excel. Their keystrokes land on a repaint-numbed frame - the
    # formula they type shows nothing, backspace appears to do nothing -
    # and the window's own COM calls race their typing. Skip the cycle as
    # a whole; the forced recalculation it owed is carried to the next
    # one, so strategy formulas never sit stale behind the skip.
    if _user_is_excel_interacting():
        now_mono_skip = time.monotonic()
        if now_mono_skip - us_state._user_input_skip_log_at > 30.0:
            us_state._user_input_skip_log_at = now_mono_skip
            feed_log(
                "render_cycle_skipped_user_input",
                note="customer is typing or selecting in our Excel",
            )
        us_state._recalc_owed = True
        _performance_metrics.record("ALL", "render_cycle_skipped_user_input", 0.0)
        return

    # Detect Excel edit mode: the customer may be dragging a formula, filling
    # cells, or editing a value while the tool runs. Writing to the Chain sheet
    # at this moment deadlocks Excel's COM apartment (Application.Ready goes
    # False, CutCopyMode is active). Skip the write cycle entirely and poll
    # again next tick - there is no timeout that forces a window through while
    # the edit is still live: opening one freezes the frame the formula is
    # being typed into, and every write inside it is refused anyway.
    app = us_state._app_cache
    if app is not None and _excel_in_edit_mode(app):
        now_mono_edit = time.monotonic()
        if now_mono_edit - us_state._user_edit_skip_log_at > 30.0:
            us_state._user_edit_skip_log_at = now_mono_edit
            feed_log(
                "render_cycle_skipped_user_edit",
                note="Excel is in edit mode; the write window stays shut",
            )
        _performance_metrics.record("ALL", "render_cycle_skipped_user_edit", 0.0)
        return

    # Final pre-window guard: the input/edit checks above take a few COM
    # round-trips, and the trader can start typing during them. Opening the
    # window now is the freeze we are fixing, so yield once more.
    if _user_is_excel_interacting():
        us_state._recalc_owed = True
        _performance_metrics.record(
            "ALL", "render_cycle_skipped_user_input_late", 0.0
        )
        return

    # Strategy-safe mode (data-freshness variant): chain data keeps updating
    # every cycle even while parked on a non-Chain sheet. Writing is cheap in
    # manual calc mode, so only the heavy Calculate and calc-mode flip are
    # deferred (see _time_recalc and the watchdog). Owed recalc carried.
    if app is not None and _active_sheet_is_chain(app) is False:
        now_mono_strat = time.monotonic()
        if now_mono_strat - us_state._strategy_sheet_skip_log_at > 30.0:
            us_state._strategy_sheet_skip_log_at = now_mono_strat
            feed_log(
                "render_cycle_skipped_strategy_sheet",
                note="active sheet is not a chain sheet; writes continue, recalc deferred",
            )
        us_state._recalc_owed = True
        _performance_metrics.record(
            "ALL", "render_cycle_skipped_strategy_sheet", 0.0
        )


    aborted = False
    restore = _begin_excel_bulk_mode(bulk_app())
    try:
        for role in ROLE_ORDER:
            if not excel_process_running():
                break
            if _user_is_excel_interacting():
                # The customer's hands reached our Excel while this window was
                # open. Drop what is left of the window and let the finally
                # below hand the screen back; their edit runs on a live frame.
                aborted = True
                us_state._recalc_owed = True
                _performance_metrics.record(
                    "ALL", "render_cycle_aborted_user_input", 0.0
                )
                break
            try:
                write_to_excel(role)
            except Exception as e:
                log_excel_exception(role, e, "render pass")
        if not aborted:
            _time_recalc(us_state._cycle_wrote)
    finally:
        if aborted:
            _release_manual_on_user_abort()
        if restore is not None:
            try:
                restore()
                # If we already recalculated inside the window, restoring
                # automatic should find nothing left to do. A second pass means
                # the cost is being paid twice.
                if us_state._app_cache is not None:
                    _probe_calc_state_after_restore()
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


def _probe_calc_state_after_restore():
    now = time.monotonic()
    cached = us_state._calc_state_probe
    if cached is not None and now - cached[0] < 30.0:
        _performance_metrics.record("ALL", "excel_recalc_second_pass", cached[1])
        return
    try:
        state = int(us_state._app_cache.api.CalculationState)
    except Exception:
        return
    hit = 1.0 if state != 0 else 0.0
    us_state._calc_state_probe = (now, hit)
    _performance_metrics.record("ALL", "excel_recalc_second_pass", hit)
    if hit:
        _note_recalc_second_pass(now)


def _note_recalc_second_pass(now):
    """Record that Excel still owed a recalc after the write window.

    The customer's own formula workbook is heavier than the configured refresh
    assumed. Stretch the loop's sleep so Excel gets idle time, and let it decay
    back once Excel comes out of every write window clean.
    """
    widened = min(us_state.WIDEN_MAX, us_state._refresh_widen * us_state.WIDEN_STEP)
    if widened != us_state._refresh_widen:
        us_state._refresh_widen = widened
        feed_log("refresh_widened", widen=round(widened, 2))
    us_state._refresh_widen_hold_until = now + us_state.WIDEN_DECAY_SECONDS


def _log_speed_summary(cycle_ms):
    """Write the whole pipeline's timings into the log every half minute.

    Every future complaint should come with numbers, so one line carries the
    whole picture: how long a cycle took, how long the customer's own workbook
    took to recalculate, how fresh the index quote is, and how many contracts
    are being tracked.
    """
    now = time.monotonic()
    if now - us_state._speed_log_at < 30.0:
        return
    us_state._speed_log_at = now
    feed = us_state._databento_feed
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
        recalc_ms=round(us_state._last_recalc_ms, 1) if us_state._last_recalc_ms else 0,
        quotes=quotes,
        spot_quote_age_s=(round(quote_age, 1) if quote_age is not None else None),
        spot_feed=quote_feed,
    )



# xlwings' .calculation getter hands back a stale value while Excel is busy, so a
# "did the restore work?" check built on it can pass falsely and leave the trader
# stuck on manual. Set AND verify through the raw API with these ids.

def _raw_calc_mode(app):
    return int(app.api.Calculation)


_DEAD_COM_CODES = {
    0x800A01A8,  # OLE error 0x800A01A8 - the app object is gone
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE - the Excel process is gone
    0x800706B5,  # RPC_S_CALL_FAILED
    0x80010108,  # RPC_E_DISCONNECTED
}


def _com_dead(exc):
    """True when the failure says the Excel instance itself is gone.

    A switch refused while a modal is up is a recoverable refusal and keeps
    the usual handling. When Excel has closed, recording a refusal only
    nuclear-armes the slow-down notice a minute later.
    """
    codes = []
    hr = getattr(exc, "hresult", None)
    if hr is not None:
        codes.append(hr)
    args = getattr(exc, "args", ()) or ()
    for item in args:
        if isinstance(item, int):
            codes.append(item)
        winerr = getattr(item, "winerror", None)
        if winerr is not None:
            codes.append(winerr)
        scode = getattr(item, "scode", None)
        if scode is not None:
            codes.append(scode)
    seen = [c & 0xFFFFFFFF for c in codes if isinstance(c, int)]
    return any(code in _DEAD_COM_CODES for code in seen)


def _release_dead_excel_state():
    """After Excel is gone: drop the calc session, marks and stale handles."""
    _calc_session["owned"] = False
    _calc_session["failed"] = False
    _calc_session["mode"] = None
    _calc_session["app"] = None
    try:
        calc_mode_guard.release(PRODUCT_ID)
    except Exception:
        pass
    us_state._app_cache = None
    _clear_refused_calc_marks()


def _excel_alive_after(exc):
    """Ground truth behind any switch failure: is Excel still running?

    Some teardown moments throw 0x800A03EC ("Unable to set the Calculation
    property") - the same HRESULT a real dialog-up refusal raises. Only the
    process check separates a closing Excel from one that is actually there
    refusing. Errors probing for the process default to "alive".
    """
    try:
        return excel_process_running()
    except Exception:
        return True


def _set_calc_mode(app, name):
    app.api.Calculation = _CALC_IDS[name]
    return _raw_calc_mode(app) == _CALC_IDS[name]


_calc_watchdog_skip_logged_at = 0.0


def _other_chain_tool_holds():
    """True when the current manual window belongs to a running sibling tool.

    The watchdog must never rescue a mode a live tool owns: the sibling chain
    product deliberately holds manual for its own write windows, and flipping
    to automatic underneath it is the fight that made Excel flap whenever
    both tools ran side by side.
    """
    global _calc_watchdog_skip_logged_at
    try:
        holds = calc_mode_guard.other_holder_alive(PRODUCT_ID)
    except Exception as e:
        feed_log(
            "calc_watchdog_holdercheck_error",
            error=type(e).__name__,
            detail=str(e)[:150],
        )
        return False
    if not holds:
        return False
    now = time.time()
    if now - _calc_watchdog_skip_logged_at >= 60.0:
        _calc_watchdog_skip_logged_at = now
        feed_log(
            "calc_watchdog_skip_holder",
            note="another chain tool holds manual calculation; watchdog left it alone",
        )
    return True


_last_observed_calc_raw = None
_calc_mode_flap_count = 0
_calc_flap_summary_after = 0.0
_CALC_FLAP_SUMMARY_SECONDS = 30.0
_CALC_ID_NAMES = {-4105: "automatic", 2: "semiautomatic", -4135: "manual"}


def _observe_raw_calc(raw):
    """Count a transition of Excel's calculation mode between cycles.

    raw comes from the read _begin_excel_bulk_mode already does every
    cycle, so this adds no COM round-trips. An unreadable value (None)
    resets the baseline so a single glitch is not mistaken for a flip.
    """
    global _last_observed_calc_raw, _calc_mode_flap_count, _calc_flap_summary_after
    if raw is None:
        _last_observed_calc_raw = None
        return
    previous = _last_observed_calc_raw
    _last_observed_calc_raw = raw
    if previous is None or previous == raw:
        return
    _calc_mode_flap_count += 1
    now = time.time()
    if now >= _calc_flap_summary_after:
        _calc_flap_summary_after = now + _CALC_FLAP_SUMMARY_SECONDS
        feed_log(
            "calc_mode_flap_summary",
            flap_total=_calc_mode_flap_count,
            last_from=_CALC_ID_NAMES.get(previous, str(previous)),
            last_to=_CALC_ID_NAMES.get(raw, str(raw)),
            note="cumulative calculation-mode transitions seen between cycles",
        )


def _open_book_names(app):
    """Lower-case names of the workbooks open in `app`, or None when Excel
    refuses to enumerate (a modal is up, or the instance is being torn down).
    """
    try:
        return {os.path.basename(book.name).lower() for book in app.books}
    except Exception:
        return None


def _chain_books_missing(open_names):
    """True when `open_names` is readable and holds no chain workbook.

    Reads-may-fail is treated the same as "cannot tell": an empty or None
    enumeration does NOT count as closed, because Excel always keeps at least
    one workbook open while it is running normally.
    """
    if not open_names:
        return False
    return not any(name in open_names for name in ROLE_FILES.values())


_US_CHAIN_SHEET_NAMES = frozenset({"CHAIN"})


def _active_sheet_is_chain(app):
    """True if Excel's active sheet is this tool's Chain sheet.

    Returns None when the active sheet cannot be read (COM failure or no app),
    which callers treat as "cannot tell" -> do NOT pause, so a COM hiccup
    never stalls the whole feed. Only an explicit non-Chain sheet name pauses.
    """
    if app is None:
        return None
    try:
        name = app.api.ActiveSheet.Name
    except Exception:
        return None
    if name is None:
        return None
    try:
        return str(name).strip().upper() in _US_CHAIN_SHEET_NAMES
    except Exception:
        return None




_SAVE_PROMPT_CLASSES = ("NUIDialog", "#32770")
_SAVE_PROMPT_POLL_SECONDS = 0.25
_SAVE_PROMPT_ATTEMPT_LIMIT = 8


def _prompt_belongs_to_chain(hwnd):
    """True when the prompt sits over a chain workbook window.

    The chain sheets are rewritten by every cycle, so their unsaved contents
    are never worth keeping - the tool answers those prompts itself instead
    of letting them pile up in the background. The customer's other workbooks
    get no auto-answer: what is unsaved in those is theirs to save or drop,
    and one wrong click on their side would be data loss. Excel captions live
    on the main window the prompt is owned by, so climb to the owner and
    recognise the chain wording there.
    """
    import ctypes

    user32 = ctypes.windll.user32
    owner = user32.GetAncestor(hwnd, 3)  # GA_ROOTOWNER
    if not owner:
        return False
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(owner, buf, 512)
    # File-backed captions join words with an underscore
    # ("us_stock_option_chain - Excel") while the tool's rendered frames
    # space them ("SPX | Option Chain ..."); normalise so one phrase covers
    # both shapes.
    caption = buf.value.replace("_", " ").replace("-", " ").lower()
    return "option chain" in caption or "options chain" in caption


# ----------------------------------------------------------------- keyboard.
# The save-prompt answerer used to synthesise its Alt+N press with
# keybd_event and four time.sleep calls around it. A process killed between
# the Alt-DOWN and the Alt-UP - close-all exits, a crash, antivirus - leaves
# Alt stuck down for the whole desktop: typed letters turn into menu
# accelerators, menus pop open by themselves, ESC becomes Alt+ESC (windows
# minimise) and the keyboard looks broken until a reboot. The customer saw
# this on their own machine and it forced a restart. Everything that keeps
# that from happening again lives here.

_VK_MENU = 0x12
_VK_N = 0x4E
_KEYUP = 2

# No human holds a modifier down for ten seconds; a synthetic one stuck down
# for ten seconds is a leak we are entitled to repair.
_STUCK_MODIFIER_SECONDS = 10.0

# Nonempty while our chord is in flight, so exit paths can wait for it.
_alt_chord_active = threading.Event()
_alt_modifiers_held_since = {}


def _prompt_is_foreground(hwnd):
    """True when the prompt window already has the keyboard focus.

    A modal Excel prompt pops up as the foreground window, so the usual case
    is a direct match. When anything else is foreground we do NOT steal the
    focus with SetForegroundWindow: stealing is exactly what pushed our
    Alt+N chord into whatever the customer was typing.
    """
    import ctypes

    try:
        user32 = ctypes.windll.user32
    except Exception:
        return False
    current = user32.GetForegroundWindow()
    for _ in range(8):
        if not current:
            return False
        if current == hwnd:
            return True
        current = user32.GetAncestor(current, 3)  # GA_ROOTOWNER
    return False


def _send_alt_n(hwnd):
    """Press the Don't Save accelerator (Alt+N) on a prompt window.

    Returns True when the chord fired. Three rules keep the stuck-Alt
    incident from ever coming back:
    1. the prompt must already be foreground (no SetForegroundWindow, no
       focus theft in the middle of the customer's typing);
    2. Alt-UP goes out last and on EVERY path - the finally block fires
       even when the chord itself raises - and with the down-leg only a
       couple of instructions above the release, an interrupted process can
       strand Alt for microseconds, not the old quarter of a second;
    3. the chord is one flight: _alt_chord_active lets exit paths wait for
       it before os._exit, and _watch_stuck_modifiers_tick heals whatever
       slips through anyway.
    """
    import ctypes

    if not _prompt_is_foreground(hwnd):
        return False
    try:
        user32 = ctypes.windll.user32
    except Exception:
        return False
    _alt_chord_active.set()
    try:
        user32.keybd_event(_VK_MENU, 0, 0, 0)
        user32.keybd_event(_VK_N, 0, 0, 0)
        user32.keybd_event(_VK_N, 0, _KEYUP, 0)
        return True
    except Exception as e:
        log_excel_exception("ALL", e, "alt+n chord")
        return False
    finally:
        try:
            # LAST leg, unconditional: this key-up is the modifier ever
            # getting stuck again.
            user32.keybd_event(_VK_MENU, 0, _KEYUP, 0)
        except Exception:
            pass
        _alt_chord_active.clear()


def _heal_keyboard():
    """Release every modifier the OS may hold keyboard-down.

    A previous crashed run can leave Alt or Win stuck across the whole
    desktop. Spurious key-UP events are harmless when the keys are not down,
    so healing at startup and right before every hard exit costs nothing and
    undoes the sticky state the customer once needed a reboot for.
    """
    import ctypes

    try:
        user32 = ctypes.windll.user32
    except Exception:
        return
    for vk in (0x12, 0xA4, 0x5B, 0x5C):  # menu, Lmenu, Lwin, Rwin
        try:
            user32.keybd_event(vk, 0, _KEYUP, 0)
        except Exception:
            pass


def _watch_stuck_modifiers_tick():
    """Release Alt/Win held down way too long with no chord of ours going.

    Last-resort self-repair. GetAsyncKeyState cannot tell a physical hold
    from a stuck synthetic one, so we only fire on something held beyond
    what a human ever keeps pressed, and only while no chord is active.
    """
    import ctypes

    if _alt_chord_active.is_set():
        _alt_modifiers_held_since.clear()
        return
    if platform.system() != "Windows":
        return
    try:
        user32 = ctypes.windll.user32
    except Exception:
        return
    now = time.monotonic()
    for vk in (0x12, 0x5B, 0x5C):  # menu, Lwin, Rwin
        try:
            down = bool(user32.GetAsyncKeyState(vk) & 0x8000)
        except Exception:
            down = False
        key = (vk,)
        if not down:
            _alt_modifiers_held_since.pop(key, None)
            continue
        since = _alt_modifiers_held_since.setdefault(key, now)
        if now - since < _STUCK_MODIFIER_SECONDS:
            continue
        try:
            user32.keybd_event(vk, 0, _KEYUP, 0)
        except Exception:
            pass
        _alt_modifiers_held_since.pop(key, None)
        feed_log(
            "stuck_modifier_released",
            vk=hex(vk),
            seconds=_STUCK_MODIFIER_SECONDS,
        )


def _answer_save_prompt(hwnd):
    """Press the prompt's Don't Save button; the way it was answered.

    Classic prompts are plain dialogs owning standard Button children, so a
    BM_CLICK on the Don't Save button answers them directly - no keyboard,
    no focus change. The modern NetUI prompt has no button windows at all -
    captions empty, everything inside one NetUIHWND - so for those the
    Alt+N accelerator is pressed while the dialog itself is already the
    foreground window. Returns "button", "alt_n" or "" (nothing could be
    sent this poll; the next poll retries within the attempt limit).
    """
    import ctypes

    user32 = ctypes.windll.user32
    button = 0
    while True:
        button = user32.FindWindowExW(hwnd, button, "Button", None)
        if not button:
            break
        text = ctypes.create_unicode_buffer(128)
        user32.GetWindowTextW(button, text, 128)
        label = text.value.replace("&", "").lower()
        if "don" in label and "save" in label:
            user32.SendMessageW(button, 0x00F5, 0, 0)  # BM_CLICK
            return "button"
    if _send_alt_n(hwnd):
        return "alt_n"
    return ""


def _chain_save_prompt_once(attempts):
    """One scan: find Excel save prompts over chain windows and answer them.

    `attempts` remembers how many times each prompt has been poked so a
    stuck dialog is never hammered forever; the limit just keeps a broken
    dialog from cycling clicks until the workbook is closed one way or the
    other.
    """
    if _input_is_recent():
        # Never send Alt+N (or click a button) while the trader is typing:
        # a save prompt that is already foreground can be mid-edit too.
        return
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    callback = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    cls = ctypes.create_unicode_buffer(64)
    hits = []

    def _scan(hwnd, _lparam):
        user32.GetClassNameW(hwnd, cls, 64)
        if cls.value in _SAVE_PROMPT_CLASSES and user32.IsWindowVisible(hwnd):
            hits.append(hwnd)
        return True

    proc = callback(_scan)
    user32.EnumWindows(proc, 0)
    for hwnd in hits:
        if not user32.IsWindow(hwnd):
            attempts.pop(hwnd, None)
            continue
        if not _prompt_belongs_to_chain(hwnd):
            continue
        count = attempts.get(hwnd, 0)
        if count >= _SAVE_PROMPT_ATTEMPT_LIMIT:
            continue
        attempts[hwnd] = count + 1
        way = _answer_save_prompt(hwnd)
        feed_log(
            "chain_save_prompt_answered",
            attempt=count + 1,
            way=way or "unavailable",
        )


def _chain_save_prompt_watchdog():
    attempts = {}
    while True:
        try:
            _chain_save_prompt_once(attempts)
        except Exception:
            pass
        time.sleep(_SAVE_PROMPT_POLL_SECONDS)


_prompt_watchdog_started = False


def start_chain_save_prompt_watchdog():
    """Start the save-prompt answerer, once per process.

    It runs on its own thread on purpose: the prompt freezes every COM call
    on the main thread, so a poll from inside display_loop would sit blocked
    on exactly the dialog it is trying to answer. Win32-only work (window
    enumeration and keystrokes) needs no COM and stays safe off-thread.
    """
    global _prompt_watchdog_started
    if _prompt_watchdog_started or platform.system() != "Windows":
        return
    _prompt_watchdog_started = True
    threading.Thread(
        target=_chain_save_prompt_watchdog,
        name="chain-save-prompt-answerer",
        daemon=True,
    ).start()


def _release_session_manual(app):
    """Hand the manual window back when the customer closed the chains.

    While we own manual calculation, their own workbook only recalculates
    when we ask it to. Once no chain workbook is left, nothing of ours is
    being written and hanging on to it would silently freeze their formulas
    - so release it, in both Excel and the crash record.
    """
    mode = _calc_session.get("mode")
    try:
        if mode:
            _set_calc_mode(app, mode)
    except Exception as e:
        log_excel_exception("ALL", e, "releasing session manual calculation")
    try:
        calc_mode_guard.release(PRODUCT_ID)
    except Exception as e:
        log_excel_exception("ALL", e, "releasing calculation guard on close")
    _calc_session["owned"] = False
    _calc_session["failed"] = False
    _calc_session["mode"] = None
    _calc_session["app"] = None
    us_state._manual_since = 0.0


def _release_manual_on_user_abort():
    """Hand calculation back the instant the trader reaches for the keyboard.

    US normally holds manual for the whole session, but an aborted write window
    is different: the customer is mid-edit and a forced Calculate would freeze
    what they are typing. Restore the mode we found, but never do it while the
    sibling chain tool owns the guard - that is the cross-tool invariant.
    """
    if not _calc_session.get("owned"):
        return False
    try:
        sibling_holds = calc_mode_guard.other_holder_alive(PRODUCT_ID)
    except Exception as e:
        # Cannot prove the sibling is not holding it: leaving manual alone is
        # safer than yanking it out from under a live Delta write window.
        feed_log(
            "calc_abort_holdercheck_error",
            error=type(e).__name__,
            detail=str(e)[:150],
        )
        return False
    if sibling_holds:
        feed_log(
            "calc_abort_restore_skipped_holder",
            note="sibling chain tool owns manual calculation; left it alone",
        )
        return False
    app = _calc_session.get("app") or us_state._app_cache
    mode = _calc_session.get("mode") or "automatic"
    ok = False
    if app is not None:
        try:
            ok = _set_calc_mode(app, mode)
        except Exception as e:
            log_excel_exception(
                "ALL", e, "restoring calculation mode after user abort"
            )
    if not ok:
        # Excel refuses mode changes while a cell is being edited, which is
        # exactly when this path runs. Keep the session and guard so the mode
        # we found is still owned and the next free cycle can recalculate;
        # clearing them here would release the guard while Excel stays manual.
        if app is None:
            _calc_session["owned"] = False
            _calc_session["mode"] = None
            _calc_session["app"] = None
            us_state._manual_since = 0.0
            try:
                calc_mode_guard.release(PRODUCT_ID)
            except Exception as e:
                log_excel_exception(
                    "ALL", e, "releasing calculation guard after user abort"
                )
        else:
            us_state._manual_since = time.time()
        feed_log("calc_abort_restore_deferred", mode=mode)
        return False
    _calc_session["owned"] = False
    _calc_session["mode"] = None
    _calc_session["app"] = None
    us_state._manual_since = 0.0
    try:
        calc_mode_guard.release(PRODUCT_ID)
    except Exception as e:
        log_excel_exception(
            "ALL", e, "releasing calculation guard after user abort"
        )
    feed_log("calc_abort_restored", mode=mode)
    return True


def _clear_refused_calc_marks():
    """Drop the refused-manual marks the moment the chains are gone.

    The save prompt Excel shows while the workbook closes refuses COM calls
    for a few seconds, and the refusal marks left behind would fire the
    slow-down notice long after the close had finished. Cheap when calm: no
    state, no COM.
    """
    us_state._calc_manual_failed = False
    us_state._calc_failed_since = 0.0
    us_state._calc_failed_at = 0.0
    _calc_session["failed"] = False
    us_state._bulk_will_recalc = False


def _clear_refused_calc_if_chain_gone():
    """Mark-clearing before the render slow path.

    Runs only while a refusal is still latched, so a calm cycle costs nothing.
    Reads the books of the cached Excel instance; when no chain workbook is
    left the close has finished and the marks must not survive it.
    """
    if not us_state._calc_manual_failed:
        return
    app = us_state._app_cache
    if app is None:
        # No cached Excel at all: the instance is gone, and the chains cannot
        # be open. A dead-handle refusal must not survive the close.
        _clear_refused_calc_marks()
        return
    open_names = _open_book_names(app)
    if open_names is None:
        # The instance refuses to enumerate (teardown or a modal). With a live
        # Excel that is "cannot tell", so only treat it as closed when there
        # is genuinely nothing running: a taskbar Close-all leaves neither
        # process nor cache, and the marks still have to go.
        try:
            gone = not excel_process_running()
        except Exception:
            gone = False
        if gone:
            _clear_refused_calc_marks()
            return
        # Excel is alive but books can't be enumerated — typically a save
        # prompt on workbook close. If this state has persisted for more than
        # _CANNOT_TELL_TIMEOUT seconds, assume the workbook closed and clear
        # the marks so the notice never fires.
        if (us_state._calc_failed_since
                and time.monotonic() - us_state._calc_failed_since >= _CANNOT_TELL_TIMEOUT):
            _clear_refused_calc_marks()
        return
    if not _chain_books_missing(open_names):
        return
    _clear_refused_calc_marks()

_CALC_NOTICE_GRACE_SECONDS = 60.0
_CANNOT_TELL_TIMEOUT = 15.0
_NO_BOOK_SKIP_LOG_SECONDS = 60.0


def ensure_calculation_not_stranded(app=None):
    """Safety net: never leave the trader in manual calculation.

    If Excel is still on manual well after any write window should have closed,
    put it back. Otherwise one failed restore makes every formula the trader
    types silently fail to compute.
    """
    global _watchdog_input_skip_logged_at
    if _input_is_recent():
        now = time.monotonic()
        if now - _watchdog_input_skip_logged_at > 30.0:
            _watchdog_input_skip_logged_at = now
            feed_log(
                "watchdog_skipped_recent_input",
                note="customer is actively using Excel; calc mode left alone",
            )
        return
    if app is None:
        app = us_state._app_cache
    if app is None:
        app = bulk_app()
    if app is None:
        return
    # Strategy-safe parity: while the trader is parked on a non-Chain sheet,
    # never flip calc mode back to automatic behind their back.
    if _active_sheet_is_chain(app) is False:
        now = time.monotonic()
        if now - _watchdog_input_skip_logged_at > 30.0:
            _watchdog_input_skip_logged_at = now
            feed_log(
                "watchdog_skipped_strategy_sheet",
                note="active sheet is not a chain sheet; calc mode left alone",
            )
        return


    if _calc_session.get("owned") and not _calc_session.get("failed"):
        return
    try:
        mode = _raw_calc_mode(app)
    except Exception:
        return
    now = time.time()
    if mode != -4135:
        us_state._manual_since = 0.0
        return
    if us_state._manual_since == 0.0:
        us_state._manual_since = now
        return
    if now - us_state._manual_since < MANUAL_WATCHDOG_SECONDS:
        return
    if _other_chain_tool_holds():
        return
    try:
        if _set_calc_mode(app, "automatic"):
            us_state._manual_since = 0.0
    except Exception:
        pass


def _begin_excel_bulk_mode(app=None):
    """Take the manual-calculation session for one render cycle.

    ScreenUpdating is deliberately never read or written here. Toggling it
    around the write window made Excel repaint the entire sheet on every
    release - the sheet visibly flashed, several times a second at the 0.2 s
    cadence - and any click into a cell that landed while the screen was off
    left the rest of that edit frozen behind a stale frame, so the formula the
    customer was typing never appeared. With the screen left on, Excel
    repaints only the ranges we actually write, which is a handful of column
    strips per cycle, and their typing, dragging and editing behave normally
    between cycles. The costly half is the recalculation, and that is the part
    this window still owns: calculation is switched to manual ONCE for the
    whole session below, and one Calculate() pays for it per cycle.

    Every Excel write otherwise repaints the workbook and triggers a full
    recalculation. That is cheap in an empty sheet but expensive when the user
    has a strategy workbook open in the same Excel instance, which is where
    refresh times blow up and where the "it feels slow" complaint comes from.

    Returns a restore callable, or None if nothing could be changed. Failures
    are expected and non-fatal: switching calculation raises if the user is
    mid-edit or a dialog is up, and there may be no workbook to talk to yet.
    """
    # The watchdog needs to know when NOT to interfere; the window is
    # marked open only when a restore callable actually comes back.
    global _session_release_input_skip_logged_at
    us_state._bulk_window_open = False
    state = {"app": None, "calc": None}
    try:
        if app is None:
            app = us_state._app_cache
        if app is None:
            app = us_state._owner_app
        if app is None:
            return None
        us_state._app_cache = app
        try:
            # xlwings.App has no .name; .version is the cheap liveness probe
            # that goes through to Excel and raises if the instance is gone.
            _ = app.version
        except Exception as e:
            log_excel_exception("ALL", e, "cached Excel app went stale")
            # A dead instance ends the session as well: marks, guard record and
            # the stale handle all go together, or the refusal latch would
            # survive the close and pop its notice a minute later.
            _release_dead_excel_state()
            return None
        state["app"] = app
        # Closed-workbook gate. When the customer is closing the chains, Excel
        # is either gone or busy answering its save prompt, and both states
        # refuse COM calls. There is nothing of ours to write, so the honest
        # response is to skip the write window - which also hands manual
        # calculation back the moment the chains are no longer open, instead
        # of holding it over whatever workbook the customer has left.
        open_names = _open_book_names(app)
        if _chain_books_missing(open_names):
            if _input_is_recent():
                # The trader is typing in whatever workbook is left. Switching
                # calculation out from under them is the same freeze we are
                # fixing, so leave the session held for one more quiet cycle.
                now = time.monotonic()
                if now - _session_release_input_skip_logged_at > 30.0:
                    _session_release_input_skip_logged_at = now
                    feed_log(
                        "session_release_skipped_user_input",
                        note="chain books gone but customer is typing; calc left alone",
                    )
                return None
            _release_session_manual(app)
            _clear_refused_calc_marks()
            now = time.monotonic()
            if now - us_state._no_book_skip_logged_at > _NO_BOOK_SKIP_LOG_SECONDS:
                us_state._no_book_skip_logged_at = now
                feed_log(
                    "bulk_mode_skipped_no_chain_workbook",
                    note="no chain workbook is open; skipping the write window",
                )
            return None
        # Deliberate no-op where the ScreenUpdating toggle used to live. The
        # screen is now left exactly as Excel itself keeps it: flipping it
        # around the write window repainted the whole sheet on every release
        # (the flicker count) and froze any edit that started mid-window.
        # The first COM call below tells us instantly if Excel can talk.
        try:
            # Read the mode through the raw API on purpose. xlwings' own getter
            # is calculation_i2s[self.xl.Calculation], and when Excel is busy or
            # showing a modal it hands back an HRESULT instead of a valid
            # xlCalculation value, so that lookup raises KeyError.
            try:
                raw = int(app.api.Calculation)
            except Exception:
                raw = None
            _observe_raw_calc(raw)
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
                if not _calc_session.get("owned"):
                    _calc_session["owned"] = False
            else:
                # Coexistence guard: if the sibling Option Chain tool already
                # acquired manual calculation, do not try to take it too:
                # calc_mode_guard.acquire() is a file append, not a lock, so
                # both tools can pass it at once and then fight over the
                # actual Excel switch -- the first error dialog the customer
                # sees on joint startup. Skip the switch entirely and let the
                # sibling own it for this cycle.
                if calc_mode_guard.other_holder_alive(PRODUCT_ID):
                    state["calc"] = None
                else:
                    state["calc"] = current or "automatic"
                    # Written BEFORE Excel is changed. If this process dies inside
                    # the window, the next launch finds the record and puts Excel
                    # back, instead of leaving their formulas silently dead.
                    calc_mode_guard.acquire(state["calc"], PRODUCT_ID)
                    state["calc_guarded"] = True
                if state["calc"]:
                    # If the workbook was already on automatic, Excel will
                    # recalculate by itself once we restore. Semiautomatic does
                    # the same for ordinary formulas.
                    us_state._bulk_will_recalc = state["calc"] in ("automatic", "semiautomatic")
                    if _calc_session["failed"]:
                        # We already know Excel will not let us change the mode.
                        # Stop asking: every failed attempt costs a COM call and
                        # makes Excel busier, which is the loop that made the
                        # customer's workbook crawl. But after a cooldown, the user
                        # may have cleared the cause (pressing Escape, closing a
                        # dialog), so we retry once to recover the manual window.
                        now = time.monotonic()
                        if now - us_state._calc_failed_at < us_state.CALC_RETRY_COOLDOWN_SECONDS:
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
                            was_failed = us_state._calc_manual_failed
                            app.calculation = "manual"
                            _calc_session["owned"] = True
                            _calc_session["mode"] = state["calc"]
                            _calc_session["failed"] = False
                            _calc_session["app"] = app
                            us_state._calc_manual_failed = False
                            us_state._calc_failed_since = 0.0
                            if was_failed:
                                feed_log("calc_manual_restored",
                                         detail="recovered manual calculation after retry")
                        except Exception as exc:
                            state["calc"] = None
                            state["calc_guarded"] = False
                            if _com_dead(exc) or not _excel_alive_after(exc):
                                # Excel died mid-window (the customer closed
                                # it). A dead handle is not a refusal: latching
                                # one would pop the slow-down notice about a
                                # minute later. Drop the state and let the
                                # reopen path bring Excel back.
                                _release_dead_excel_state()
                                feed_log(
                                    "calc_switch_excel_gone",
                                    detail=type(exc).__name__,
                                )
                            else:
                                _calc_session["failed"] = True
                                _calc_session["owned"] = False
                                if not us_state._calc_failed_since:
                                    us_state._calc_failed_since = time.monotonic()
                                us_state._calc_failed_at = time.monotonic()
                                feed_log(
                                    "calc_manual_unavailable",
                                    error=type(exc).__name__,
                                    detail=str(exc)[:120],
                                )
                                us_state._calc_manual_failed = True
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
        us_state._bulk_window_open = False
        app = state.get("app")
        if app is None:
            return
        # No paint action here either: the screen was never taken, so it is
        # already exactly where Excel and the customer want it. Every release
        # used to force a full-sheet repaint here - that was the flicker.
        # The calculation mode is deliberately NOT restored here. We take it
        # once for the whole session and release it on exit, so the two
        # OptionChain tools no longer fight over it several times a second.
        # _emergency_restore() puts it back when the tool stops.

    us_state._last_bulk_restore = restore
    us_state._bulk_window_open = True
    return restore


def bulk_app():
    """Resolve the Excel App COM object to use for bulk mode, caching it."""
    if us_state._app_cache is not None:
        try:
            _ = us_state._app_cache.version
            return us_state._app_cache
        except Exception:
            us_state._app_cache = None
    app = us_state._owner_app
    if app is None:
        return None
    try:
        _ = app.version
        us_state._app_cache = app
        return app
    except Exception:
        us_state._owner_app = None
        us_state._owner_hwnd = None
        return None


async def display_loop():
    is_windows = platform.system() == "Windows"
    license_tick = 0
    us_state._excel_gone_logged = False
    if is_windows:
        # The chain save prompt would otherwise sit unanswered in the
        # background blocking COM until the customer notices it.
        start_chain_save_prompt_watchdog()
    while True:
        render_started = time.perf_counter()
        # re-check license roughly every ~60s while running
        license_tick += 1
        if license_tick >= max(1, int(60 / max(REFRESH, 0.05))):
            license_tick = 0
            ok, _vt = license_is_valid()
            if not ok:
                if not ensure_license():
                    _heal_keyboard()
                    os._exit(1)
        if is_windows:
            _pump_com()
            # A modifier stuck from a crash - or from anything else on the
            # machine - is healed here, once per cycle, without ever
            # touching a key a human could legitimately still be holding.
            _watch_stuck_modifiers_tick()
            # The chain lives in its private Excel instance, so the write pass
            # runs while that instance is alive - whichever other Excel
            # windows exist on the machine are none of this tool's business.
            # Keep the cached hwnd current: Excel can recreate its main window
            # when workbooks open/close, and a stale handle makes the independent
            # watchdog (and IsWindow checks) believe the instance is dead.
            if us_state._owner_app is not None:
                try:
                    _h = us_state._owner_app.api.Hwnd
                    if _h:
                        us_state._owner_hwnd = _h
                except Exception:
                    pass
            excel_alive = _owner_alive() or us_state._owner_app is None
            retry_open_pass_after_grace()
            if excel_alive:
                if us_state._excel_gone_logged:
                    # Excel came back. Everything cached about the sheets is now
                    # wrong: the workbooks are freshly opened and their cells are
                    # empty, so the payload/shape/format caches must be dropped
                    # or the app would treat the blank sheets as "already
                    # written" and leave them blank.
                    us_state._excel_gone_logged = False
                    _block_payload.clear()
                    _colour_state.clear()
                    _sheet_fmt_done.clear()
                    _usd_fmt_done.clear()
                    _dropdown_sig.clear()
                    us_state._app_cache = None
                    _book_sheet_cache.clear()
                    _sheet_cache.clear()
                    us_state._missing_since.clear()
                    print("Excel reopened - rewriting all sheets from scratch.", flush=True)
                    _last_excel_exception.clear()

                # The window is closed at this point, so anything the last
                # restore could not finish gets finished here.
                _screen_update_watchdog_tick()
                render_cycle()

            elif not us_state._excel_gone_logged:
                us_state._excel_gone_logged = True
                print(
                    "The chain Excel was closed. The tool stops on its own "
                    "shortly - run it again when you want the chain back.",
                    flush=True,
                )
                _last_excel_exception.clear()
                # Excel's address space is gone, so any manual window we held
                # is gone with it. Drop the session and the claim so the next
                # cycle starts clean instead of talking to a dead app, and the
                # crash record cannot outlive the instance it belonged to.
                _calc_session["owned"] = False
                _calc_session["failed"] = False
                _calc_session["mode"] = None
                _calc_session["app"] = None
                us_state._calc_manual_failed = False
                us_state._calc_failed_since = 0.0
                us_state._calc_failed_at = 0.0
                us_state._bulk_will_recalc = False
                us_state._app_cache = None
                us_state._owner_app = None
                us_state._owner_hwnd = None
                _book_sheet_cache.clear()
                _sheet_cache.clear()
                try:
                    calc_mode_guard.release(PRODUCT_ID)
                except Exception:
                    pass
            maybe_exit_after_close_all(excel_alive=excel_alive)
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
        # writing the chains leaves no idle gap, Excel is busy whenever they
        # try to click a cell and the workbook feels stuck. Measured with a
        # heavy 800-row formula workbook open, one pass took ~880ms; against a
        # 1.0s refresh that is ~88% duty cycle. Backing off to an equal amount
        # of idle time keeps the workbook usable and lets the refresh stretch
        # automatically when their workbook is heavy, while a light workbook
        # still updates at the configured interval.
        #
        # The stretch is adaptive: a CalculationState probe that still shows
        # work owed after the write window (see _note_recalc_second_pass)
        # multiplies the configured interval, and the stretch decays back to
        # 1.0 after a quiet spell so a fast workbook is never left crawling.
        now_mono = time.monotonic()
        if (
            now_mono > us_state._refresh_widen_hold_until
            and us_state._refresh_widen > 1.0
        ):
            us_state._refresh_widen = max(1.0, us_state._refresh_widen / 2.0)
            us_state._refresh_widen_hold_until = now_mono + us_state.WIDEN_DECAY_SECONDS
            if us_state._refresh_widen > 1.0:
                feed_log("refresh_widen_decayed", widen=round(us_state._refresh_widen, 2))
        # Relax the recalc-driven backoff between cycles so a workbook that went
        # quiet shrinks back toward the configured refresh instead of staying
        # throttled.
        if us_state._recalc_backoff_seconds > 0.0:
            us_state._recalc_backoff_seconds *= us_state.RECALC_BACKOFF_DECAY
            if us_state._recalc_backoff_seconds < us_state.RECALC_BACKOFF_THRESHOLD_SECONDS:
                us_state._recalc_backoff_seconds = 0.0
        await asyncio.sleep(
            min(
                us_state.MAX_REFRESH_SECONDS,
                max(
                    refresh_delay(
                        REFRESH * us_state._refresh_widen,
                        REFRESH,
                        time.perf_counter() - render_started,
                    ),
                    us_state.RECALC_BACKOFF_FACTOR * us_state._recalc_backoff_seconds,
                ),
            )
        )





