"""
MCX Commodity Option Chain Tool
MCX options (CRUDEOIL, GOLD, SILVER) via Upstox REST market-quote.
Three Excel files auto-open:
  crudeoil_chain.xlsx  -> locked to CRUDEOIL
  gold_chain.xlsx      -> locked to GOLD
  silver_chain.xlsx    -> locked to SILVER
Shared cache by symbol; one process drives all books.
"""

import asyncio
import atexit
import os
import platform
import re
import socket
import sys
import time
from datetime import datetime

from mcx_diagnostics import (
    append_app_exception,
    append_connection_event,
    append_performance_summary,
    build_id,
    feed_log,
    write_startup_diagnostics,
)

_orig_getaddrinfo = socket.getaddrinfo


def _ipv4_getaddrinfo(*args, **kwargs):
    results = _orig_getaddrinfo(*args, **kwargs)
    v4 = [r for r in results if r[0] == socket.AF_INET]
    return v4 if v4 else results


socket.getaddrinfo = _ipv4_getaddrinfo

import mcx_chain
import mcx_spot
import mcx_state
from mcx_state import (
    BASE_DIR, BLOCK_COLS, EXPIRY_REFRESH_SECONDS, HTTP_TIMEOUT, MCX_SYMBOLS,
    NET_FAIL_LIMIT, OI_POLL_SECONDS, PRODUCT_EXE, PRODUCT_ID, ROW_EXTENT_DEFAULT,
    SPOT_POLL_SECONDS, STALE_AFTER_S, _auto_open_done, _block_hidden,
    _block_payload, _block_values, _book_sheet_cache, _calc_state_probe,
    _colour_state, _cycle_wrote, _date_req_sig, _dropdown_sig,
    _excel_connection_status, _excel_liveness_cache, _last_excel_exception,
    _last_rows, _net, _net_drawn, _pending_date_loads, _performance_metrics,
    _recalc_owed, _refresh_widen, _screen_restore_owed, _sheet_altered_notice,
    _sheet_cache, _usd_fmt_done, active_expiries, active_expiries_dmy,
    available_dates, cfg, excel_connected, excel_fail_notice_shown,
    excel_fail_since, excel_miss_counts, expiry_symbols, feed_state,
    last_spot_change, last_ws_recv, live_data, spot_cmp, spot_source,
    symbol_expiry, user_pinned,
)
from mcx_spot import (
    parity_spot_or_none,
    set_spot,
    spot_age,
    spot_is_stale,
)
from mcx_chain import (
    apply_tickers,
    available_expiry_dates,
    build_chain,
    fetch_instruments,
    fetch_products,
    fetch_quotes,
    fmt_oi,
    get_last_good_chain,
    iso_to_dmy,
    last_good_age,
    load_expiry_tickers,
    n_blocks,
    note_feed_good,
    oi_poller,
    probe_expiry_count,
    set_feed_state,
    spot_poller,
)

from mcx_instance import kill_previous_instances

from dotenv import load_dotenv

import calc_mode_guard


def net_badge():
    """Header tag text ('' = sab theek, kuch nahi likhna)."""
    if _net["fails"] < NET_FAIL_LIMIT:
        return ""
    secs = int(time.time() - (_net["since"] or time.time()))
    m, s = divmod(secs, 60)
    span = f"{m}m {s:02d}s" if m else f"{s}s"
    return f"OFFLINE {span} - reconnecting..."


def net_paint(sheet, symbol, bi, col_l):
    """Tag cell (row 1, block ka last column) likho — sirf jab badle."""
    tag = net_badge()
    prev = _net_drawn.get((symbol, bi), None)
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
    _net_drawn[(symbol, bi)] = tag


from mcx_instance import kill_previous_instances

from dotenv import load_dotenv

import calc_mode_guard

AUTO_OPEN_TIMEOUT_SECONDS = 30.0


def show_error(msg):
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "MCX Option Chain", 0x10)
    else:
        print(msg, flush=True)


def open_excel_files():
    """Open the three workbooks in ONE Excel, and get on with the data."""
    if platform.system() != "Windows":
        return
    mcx_state._open_pass_at = time.monotonic()
    _open_excel_files()


def _pump_com():
    """Keep the display thread's COM apartment alive between cycles."""
    try:
        import pythoncom
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def _ensure_owner_app(xw):
    """Return the tool's private Excel instance, creating it when needed."""
    app = mcx_state._owner_app
    if app is not None:
        try:
            _ = app.version
            return app
        except Exception:
            mcx_state._owner_app = None
            mcx_state._owner_hwnd = None
            mcx_state._app_cache = None
    try:
        app = xw.App(visible=True)
    except Exception as e:
        print(f"Could not start Excel: {e}", flush=True)
        return None
    mcx_state._owner_app = app
    mcx_state._app_cache = app
    try:
        mcx_state._owner_hwnd = app.api.Hwnd
    except Exception:
        mcx_state._owner_hwnd = None
    try:
        feed_log(
            "excel_open_own_instance",
            note="starting a private Excel instance for this tool only",
        )
    except Exception:
        pass
    return app


def _open_excel_files(finished=None, own_instance=False):
    pythoncom = None
    app = None
    started_excel = False
    try:
        import xlwings as xw

        app = _ensure_owner_app(xw)
        if app is None:
            show_error(
                "Microsoft Excel could not be started.\n\n"
                "Please make sure Excel is installed and works, then run "
                "the chain again."
            )
            if finished is not None:
                finished.set()
            return
        started_excel = True
        feed_log("excel_auto_open_started", own_instance=True)
        for symbol in MCX_SYMBOLS:
            path = os.path.join(BASE_DIR, f"{symbol.lower()}_chain.xlsx")
            if not os.path.exists(path):
                record_excel_connection(symbol, "workbook_file_missing", f"path={path}")
                show_error(f"File not found:\n{path}\n\nKeep this file next to MCXOptionChain.exe and run again.")
                continue
            try:
                os.remove(path + ":Zone.Identifier")
            except OSError:
                pass
            try:
                if find_open_book(xw, path) is not None:
                    _auto_open_done[symbol] = True
                    feed_log("excel_auto_open_done", symbol=symbol, via="already_open")
                    continue
            except Exception:
                pass
            target = path if len(path) <= 120 else short_path(path)
            try:
                app.api.Workbooks.Open(target)
                _auto_open_done[symbol] = True
                feed_log("excel_auto_open_done", symbol=symbol, via="api")
            except Exception as e:
                print(f"Could not open {symbol.lower()}_chain.xlsx: {e}", flush=True)
                feed_log("excel_auto_open_failed", symbol=symbol, error=str(e))
                show_error(
                    f"Could not open {symbol.lower()}_chain.xlsx:\n\n{e}\n\n"
                    "If the package was unzipped into a long path, move it to "
                    "a short one - C:\\OptionChain - and run the tool again."
                )
                try:
                    append_app_exception(f"opening {symbol} workbook", e)
                except OSError as log_error:
                    print(f"Could not log workbook open failure ({symbol}): {log_error}", flush=True)
        if finished is not None:
            finished.set()
    except Exception as e:
        print(f"Could not open the workbooks: {type(e).__name__}: {e}", flush=True)
        if finished is not None:
            finished.set()
    finally:
        try:
            if app is not None:
                _new = app.api.Hwnd
                if _new:
                    mcx_state._owner_hwnd = _new
        except Exception:
            pass
        if started_excel:
            _close_empty_book1(app)


def _close_empty_book1(app):
    """Close the empty Book1 Excel always starts with."""
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


_OPEN_RETRY_GRACE_SECONDS = 45.0


def retry_open_pass_after_grace():
    """Retry a wedged startup open pass once, on the display thread."""
    if mcx_state._open_pass_retried or mcx_state._open_pass_at is None:
        return
    now = time.monotonic()
    if now - mcx_state._open_pass_at < _OPEN_RETRY_GRACE_SECONDS:
        return
    mcx_state._open_pass_retried = True
    if all(mcx_state._auto_open_done.get(symbol) for symbol in MCX_SYMBOLS):
        return
    feed_log(
        "excel_open_retried",
        note=(f"nothing latched {int(_OPEN_RETRY_GRACE_SECONDS)}s after the startup "
              "open pass; retrying it once"),
    )
    open_excel_files()


def parse_user_date(s):
    """Excel often coerces cell text into a datetime object before we see it."""
    if isinstance(s, datetime):
        return s.strftime("%Y-%m-%d")
    s = str(s).strip().upper()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        dd, mm, yyyy = m.groups()
        mi, di = int(mm), int(dd)
        if 1 <= mi <= 12 and 1 <= di <= 31:
            return f"{yyyy}-{mi:02d}-{di:02d}"
        return None
    m = re.match(r"^(\d{2})([A-Z]{3})(\d{2})$", s)
    if m:
        dd, mmm, yy = m.groups()
        mm = mcx_state.MONTHS.get(mmm)
        if not mm:
            return None
        return f"20{yy}-{mm}-{dd}"
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


# ---------------------------------------------------------------------------
# Feed observability.
# ---------------------------------------------------------------------------
from mcx_diagnostics import (
    _flat,
    feed_log,
)

# asset -> {state, detail, retry_in, last_good_ts, last_good_label}


def get_default_expiries(symbol, products):
    """The first expiries the slots start with — always n_blocks wide."""
    dates = available_expiry_dates(symbol, products)[:n_blocks(symbol)]
    while len(dates) < n_blocks(symbol):
        dates.append(None)
    return dates


async def expiry_roller(symbol):
    first = True
    while True:
        if first:
            first = False
        else:
            await asyncio.sleep(EXPIRY_REFRESH_SECONDS)
        try:
            products = await asyncio.to_thread(fetch_products, symbol)
            new_smap = {}
            new_esym = {}
            for p in products:
                sym_key = p.get("symbol", "")
                st = (p.get("settlement_time") or "")[:10]
                if not sym_key or not st:
                    continue
                new_smap[sym_key] = st
                new_esym.setdefault(st, []).append(sym_key)
            symbol_expiry[symbol].update(new_smap)
            expiry_symbols[symbol].update(new_esym)

            all_dates = available_expiry_dates(symbol, products)
            available_dates[symbol] = all_dates
            probe_dates = all_dates[:6]
            counts = await asyncio.gather(
                *[asyncio.to_thread(probe_expiry_count, symbol, d) for d in probe_dates]
            )
            states = [st for _, st in counts]
            nums = [n for n, _ in counts]
            with_data = [d for d, n in zip(probe_dates, nums) if n > 0]

            feed_log("roller_decision", symbol=symbol, probe_dates=probe_dates,
                     nums=nums, states=states, with_data=with_data,
                     active_before=list(active_expiries[symbol]),
                     n_blocks=n_blocks(symbol))

            if not with_data and "ERROR" in states:
                feed_log("roller_all_probes_failed", symbol=symbol,
                         expiry_count=len(probe_dates), states=states)
                await asyncio.sleep(2.0 + __import__("random").uniform(0, 1.5))
                counts = await asyncio.gather(
                    *[asyncio.to_thread(probe_expiry_count, symbol, d)
                      for d in probe_dates]
                )
                states = [st for _, st in counts]
                nums = [n for n, _ in counts]
                with_data = [d for d, n in zip(probe_dates, nums) if n > 0]
                feed_log("roller_retry_result", symbol=symbol,
                         recovered=bool(with_data), rows=nums)

            if "ERROR" in states:
                feed_log("roller_probe_incomplete", symbol=symbol,
                         states=states, active=list(active_expiries[symbol]),
                         note="keeping the current expiry assignment")
                continue

            pinned_vals = [active_expiries[symbol][i] for i in range(n_blocks(symbol)) if user_pinned[symbol][i]]
            desired = [d for d in with_data if d not in pinned_vals]

            changed_slots = []
            slot = 0
            for i in range(n_blocks(symbol)):
                if user_pinned[symbol][i]:
                    continue
                cur = active_expiries[symbol][i] if i < len(active_expiries[symbol]) else None
                tgt = desired[slot] if slot < len(desired) else None
                slot += 1
                if tgt == cur:
                    continue
                if not with_data and cur:
                    set_feed_state(symbol, "ERROR",
                                   "no expiry data from Upstox", retry_in=0.0)
                    continue
                if tgt:
                    await asyncio.to_thread(load_expiry_tickers, symbol, tgt, __import__("os").environ.get("UPSTOX_API_TOKEN", ""))
                    note_feed_good(symbol, tgt)
                while len(active_expiries[symbol]) <= i:
                    active_expiries[symbol].append(None)
                while len(active_expiries_dmy[symbol]) <= i:
                    active_expiries_dmy[symbol].append(None)
                changed_slots.append((i, cur, tgt))
                active_expiries[symbol][i] = tgt
                active_expiries_dmy[symbol][i] = iso_to_dmy(tgt) if tgt else None

            feed_log("roller_result", symbol=symbol,
                     active_after=list(active_expiries[symbol]),
                     changed=len(changed_slots))

            if changed_slots:
                try:
                    import xlwings as xw
                    book = find_open_book(xw, os.path.join(BASE_DIR, f"{symbol.lower()}_chain.xlsx"))
                    if book:
                        sheet = book.sheets[0]
                        for i, old, new in changed_slots:
                            cell = ["D3", "O3", "Z3"][i]
                            _set_date_cell(sheet.range(cell), iso_to_user(new) if new else "")
                except Exception as e:
                    print(f"Roller cell write error ({symbol}): {e}", flush=True)
                print(f"[{symbol}] expiry rolled: " + ", ".join(f"{o or '--'} -> {n or '--'}" for _, o, n in changed_slots), flush=True)
        except Exception as e:
            mcx_chain._console_warn_last(f"Expiry roller error ({symbol})", e)


# ---------- Terminal Display (Linux testing only) ----------


def update_user_dates(symbol, book, sheet, pre=None):
    """Read date cells only — HTTP loads are queued to date_loader (never blocks event loop)."""
    expiries = active_expiries.get(symbol) or []
    for i, col3 in enumerate((4, 15, 26)):
        if i >= n_blocks(symbol):
            rng = sheet.range((3, col3))
            if cell_value(pre, sheet, 3, col3) not in (None, ""):
                _set_date_cell(rng, "")
            user_pinned[symbol][i] = False
            _date_req_sig.pop((symbol, i), None)
            continue
        rng = sheet.range((3, col3))
        val = cell_value(pre, sheet, 3, col3)
        if val is None or (isinstance(val, str) and val.strip() == ""):
            cur = expiries[i] if i < len(expiries) else None
            _set_date_cell(rng, iso_to_user(cur) if cur else "")
            user_pinned[symbol][i] = False
            _date_req_sig.pop((symbol, i), None)
            continue
        parsed = parse_user_date(val)
        if not parsed:
            continue
        cur = expiries[i] if i < len(expiries) else None
        if not (isinstance(val, str) and val.strip().upper() == iso_to_user(parsed)):
            _set_date_cell(rng, iso_to_user(parsed))
        if parsed == cur:
            _date_req_sig.pop((symbol, i), None)
            continue
        if _date_req_sig.get((symbol, i)) == parsed:
            continue
        _date_req_sig[(symbol, i)] = parsed
        _pending_date_loads.append((symbol, i, parsed))


def setup_expiry_dropdown(sheet, symbol):
    """Hidden helper column + in-cell dropdown on the expiry cells (like Delta/US)."""
    dates = available_dates.get(symbol) or []
    if not dates:
        return
    try:
        sheet.range("AH1:AH500").clear()
        sheet.range("AH1:AH500").api.NumberFormat = "@"
        sheet.range(f"AH1:AH{len(dates)}").value = [[iso_to_user(d)] for d in dates]
        sheet.range((1, 34)).api.EntireColumn.Hidden = True
    except Exception as e:
        log_excel_exception(symbol, e, "writing expiry dropdown source")
        return
    formula = f"=$AH$1:$AH${len(dates)}"
    for cell in ["D3", "O3", "Z3"][:n_blocks(symbol)]:
        rng = sheet.range(cell)
        try:
            rng.api.NumberFormat = "@"
        except Exception as e:
            log_excel_exception(symbol, e, f"formatting expiry input {cell}")
        try:
            rng.api.Validation.Delete()
        except Exception:
            pass
        try:
            rng.api.Validation.Add(Type=3, Formula1=formula)
            rng.api.Validation.ShowError = False  # dropdown = hint, typing still allowed
        except Exception as e:
            log_excel_exception(symbol, e, f"applying expiry dropdown to {cell}")


def _set_date_cell(rng, text):
    """Write date as TEXT so Excel cannot coerce it back to a serial number."""
    try:
        rng.api.NumberFormat = "@"
    except Exception as e:
        log_excel_exception("EXCEL", e, "expiry date number format")
    rng.value = text


async def date_loader():
    while True:
        if not _pending_date_loads:
            await asyncio.sleep(0.2)
            continue
        symbol, i, parsed = _pending_date_loads.pop(0)
        if i >= n_blocks(symbol):
            continue
        try:
            token = os.environ.get("UPSTOX_API_TOKEN", "")
            n = await asyncio.to_thread(load_expiry_tickers, symbol, parsed, token)
            active_expiries[symbol][i] = parsed
            active_expiries_dmy[symbol][i] = iso_to_dmy(parsed)
            user_pinned[symbol][i] = True
            print(f"[{symbol}] block {i + 1} pinned -> {parsed} ({n} contracts)", flush=True)
        except Exception as e:
            _date_req_sig.pop((symbol, i), None)
            mcx_chain._quiet_console_error(f"Date fetch error ({symbol} {parsed})", e)
            await asyncio.sleep(3)


_UNKNOWN_TTL_SECONDS = 10.0
_unknown_logged_at = {}
_CLOSE_ALL_EXIT_SECONDS = 75.0


def _owner_alive():
    """Cheap liveness probe on the tool's own Excel instance."""
    app = mcx_state._owner_app
    if app is None:
        return False
    try:
        _ = app.version
        return True
    except Exception:
        return False


def maybe_exit_after_close_all(excel_alive=None):
    """Exit like every application does: when all its windows are gone."""
    if excel_alive is False:
        now = time.monotonic()
        for symbol in MCX_SYMBOLS:
            mcx_state._missing_since[symbol] = now
    if len(mcx_state._missing_since) == 3:
        now = time.monotonic()
        if mcx_state._all_gone_since is None:
            mcx_state._all_gone_since = now
        elif now - mcx_state._all_gone_since >= _CLOSE_ALL_EXIT_SECONDS:
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
    elif mcx_state._all_gone_since is not None:
        mcx_state._all_gone_since = None
        mcx_state._missing_since.clear()


def excel_process_running(max_age=2.0):
    """True if any EXCEL.EXE is alive, checked without touching COM."""
    now = time.monotonic()
    if mcx_state._excel_liveness_cache is not None:
        stamp, running = mcx_state._excel_liveness_cache
        if now - stamp <= max_age:
            return running
    try:
        import subprocess
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        mcx_state._excel_liveness_cache = (now, True)
        return True
    running = "EXCEL.EXE" in (result.stdout or "").upper()
    mcx_state._excel_liveness_cache = (now, running)
    return running


def _long_form(path):
    """The Windows long form of a path, or the path itself."""
    try:
        import ctypes
        buffer = ctypes.create_unicode_buffer(1024)
        if ctypes.windll.kernel32.GetLongPathNameW(path, buffer, 1024):
            return buffer.value or path
    except Exception:
        pass
    return path


def find_open_book(xw, fname):
    """Find the workbook inside the tool's PRIVATE Excel instance."""
    target = os.path.abspath(fname)
    target_long = _long_form(target).lower()
    target_name = os.path.basename(target).lower()
    symbol = target_name.split("_chain", 1)[0].upper()
    if not excel_process_running():
        return None
    app = mcx_state._owner_app
    if app is None:
        return None
    try:
        books = list(app.books)
    except Exception as e:
        log_excel_exception(symbol, e, "enumerating the tool's own Excel workbooks")
        mcx_state._enum_unknown[symbol] = time.monotonic()
        return None
    for book in books:
        try:
            fullname = book.fullname
            base = os.path.basename(fullname).lower()
        except Exception as e:
            log_excel_exception(symbol, e, "reading open workbook path")
            continue
        if base == target_name and _long_form(fullname).lower() == target_long:
            return book
    return None


def _contiguous_runs(indices):
    """Merge changed-row indices into contiguous (start, end) runs."""
    if not indices:
        return []
    runs = []
    start = previous = indices[0]
    for index in indices[1:]:
        if index == previous + 1:
            previous = index
        else:
            runs.append((start, previous))
            start = previous = index
    runs.append((start, previous))
    return runs


def pad_block_rows(values, extent):
    """Pad one block's rows with blank rows up to the fixed extent."""
    if extent <= len(values):
        return values
    blank = [None] * (len(values[0]) if values else 9)
    return values + [list(blank) for _ in range(extent - len(values))]


_deep_swept = set()


def _deep_clear_below_block(sheet, symbol, bi, c0, n_rows):
    """Clear junk an older build left below the fixed extent, if any."""
    if (symbol, bi) in _deep_swept:
        return
    _deep_swept.add((symbol, bi))
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
        feed_log("chain_deep_clear", symbol=symbol, block=bi + 1, from_row=deep_start)
    except Exception as e:
        log_excel_exception(symbol, e, f"deep clearing stale rows in block {bi + 1}")


def get_column_letter(col):
    """Convert 1-indexed column number to Excel letter (A, B, ..., Z, AA, AB, ...)"""
    result = ""
    while col > 0:
        col -= 1
        result = chr(65 + col % 26) + result
        col //= 26
    return result


def cell_value(pre, sheet, row, col):
    """Read one cell from the merged 4x31 pre-read, falling back to COM."""
    if pre is not None:
        try:
            block = pre[row - 1][col - 1]
        except Exception:
            return None
        return block if block is not None else None
    if sheet is None:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def force_usd_block_formats(sheet, block, symbol, bi, n_rows):
    """Prices = 0.00, OI = numeric (kills any Rs/INR number format)."""
    if n_rows <= 0:
        return
    c0 = block["start"]
    end = 4 + n_rows
    try:
        sheet_key = os.path.abspath(sheet.book.fullname).lower()
    except Exception:
        sheet_key = str(id(sheet))
    fmt_key = (sheet_key, c0, "numfmt")
    if fmt_key in _usd_fmt_done:
        return
    try:
        for off in (0, 1, 2, 6, 7, 8):
            col = get_column_letter(c0 + off)
            sheet.range(f"{col}5:{col}{end}").api.NumberFormat = "0.00"
        strike_col = get_column_letter(c0 + 4)
        sheet.range(f"{strike_col}5:{strike_col}{end}").api.NumberFormat = "0.0#"
        for off in (3, 5):
            col = get_column_letter(c0 + off)
            sheet.range(f"{col}5:{col}{end}").api.NumberFormat = "#,##0"
        _usd_fmt_done.add(fmt_key)
    except Exception as e:
        log_excel_exception(symbol, e, "formatting option-chain cells")
        return
    for off in (3, 5):
        col = get_column_letter(c0 + off)
        try:
            rng = sheet.range(f"{col}5:{col}{end}")
            values = rng.value
        except Exception as e:
            log_excel_exception(symbol, e, f"reading open-interest column {c0 + off}")
            continue
        if not values:
            continue
        changed = False
        for i, v in enumerate(values):
            if isinstance(v, str) and ("Rs" in v or "INR" in v or "$" in v or "," in v or v.endswith("M") or v.endswith("K") or v.endswith("m") or v.endswith("k")):
                values[i] = fmt_oi(v)
                changed = True
            elif isinstance(v, str):
                try:
                    values[i] = float(v)
                    changed = True
                except (ValueError, TypeError):
                    pass
        if changed:
            try:
                rng.value = values
            except Exception as e:
                log_excel_exception(symbol, e, f"writing open-interest column {c0 + off}")


def _flush_block_writes(sheet, symbol, pending):
    """Write everything the block loop queued, in as few calls as possible."""
    if not pending:
        return
    pending = sorted(pending, key=lambda item: item[0])
    protected = mcx_state._bulk_owning_manual
    payloads = []
    for bi, block, rows_block, n_rows, runs, prev in pending:
        if protected and len(runs) == 1:
            start, end = runs[0]
            rows = [list(row) for row in rows_block[start : end + 1]]
            top = 5 + start
        elif (not protected) and runs and len(runs) <= 8:
            lo, hi = runs[0][0], runs[-1][1]
            rows = []
            for i in range(lo, hi + 1):
                if any(s <= i <= e for s, e in runs):
                    rows.append(list(rows_block[i]))
                elif prev is not None and i < len(prev):
                    rows.append(list(prev[i]))
                else:
                    rows.append(list(rows_block[i]))
            top = 5 + lo
        else:
            rows = [list(row) for row in rows_block]
            top = 5
        payloads.append({
            "bi": bi,
            "start": block["start"],
            "end": block["start"] + 8,
            "rows": rows,
            "top": top,
            "values": rows_block,
        })
    consecutive = all(
        payloads[i]["bi"] + 1 == payloads[i + 1]["bi"]
        for i in range(len(payloads) - 1)
    )
    first_c = min(p["start"] for p in payloads)
    last_c = max(p["end"] for p in payloads)
    min_row = min(p["top"] - 5 for p in payloads)
    max_row = max(p["top"] - 5 + len(p["rows"]) - 1 for p in payloads)
    individual_cells = sum(len(p["rows"]) * 9 for p in payloads)
    merged_cells = (max_row - min_row + 1) * (last_c - first_c + 1)
    merge = (
        consecutive
        and len(payloads) >= 2
        and merged_cells <= 4 * individual_cells
    )
    try:
        if merge:
            grid = []
            for r in range(min_row, max_row + 1):
                row = [None] * (last_c - first_c + 1)
                for p in payloads:
                    c_off = p["start"] - first_c
                    row_top = p["top"] - 5
                    depth = len(p["rows"])
                    n_rows = len(p["values"])
                    if row_top <= r < row_top + depth:
                        row[c_off : c_off + 9] = p["rows"][r - row_top]
                    elif r < n_rows:
                        row[c_off : c_off + 9] = list(p["values"][r])
                grid.append(row)
            sheet.range((5 + min_row, first_c)).value = grid
            mcx_state._cycle_wrote = True
        else:
            for bi, block, rows_block, n_rows, runs, prev in pending:
                if protected:
                    if len(runs) == 1:
                        start, end = runs[0]
                        sheet.range((5 + start, block["start"])).value = [
                            list(r) for r in rows_block[start : end + 1]
                        ]
                    else:
                        sheet.range((5, block["start"])).options(
                            index=False, header=False
                        ).value = [list(r) for r in rows_block]
                elif runs and len(runs) <= 8:
                    lo = runs[0][0]
                    hi = runs[-1][1]
                    merged_rows = []
                    for i in range(lo, hi + 1):
                        in_run = any(s <= i <= e for s, e in runs)
                        if in_run:
                            merged_rows.append(list(rows_block[i]))
                        elif prev is not None and i < len(prev):
                            merged_rows.append(list(prev[i]))
                        else:
                            merged_rows.append(list(rows_block[i]))
                    sheet.range((5 + lo, block["start"])).value = merged_rows
                else:
                    sheet.range((5, block["start"])).options(
                        index=False, header=False
                    ).value = [list(r) for r in rows_block]
                mcx_state._cycle_wrote = True
    except Exception as e:
        log_excel_exception(symbol, e, "flushing block writes")


def record_excel_connection(symbol, status, detail=""):
    if mcx_state._excel_connection_status.get(symbol) == status:
        return
    mcx_state._excel_connection_status[symbol] = status
    try:
        append_connection_event(symbol, status, detail)
    except OSError as e:
        print(f"Could not log Excel connection state ({symbol}): {e}", flush=True)


def log_excel_exception(symbol, exc, context="Excel/COM update"):
    now = time.monotonic()
    signature = (type(exc).__name__, str(exc))
    key = (symbol, context)
    previous = mcx_state._last_excel_exception.get(key)
    if previous and previous[0] == signature and now - previous[1] < 60:
        return
    mcx_state._last_excel_exception[key] = (signature, now)
    try:
        append_app_exception(f"{context} for {symbol}", exc)
    except OSError as log_error:
        print(f"Could not log {context} traceback ({symbol}): {log_error}", flush=True)


def sheet_structure_ok(sheet):
    """True when the sheet still has the layout the tool expects."""
    try:
        a4 = sheet.range((4, 1)).value
    except Exception:
        return True
    if a4 is None or str(a4).strip() == "":
        return True
    return str(a4).strip() == "Call LTP"


def short_path(path):
    """The Windows 8.3 short name for a path, when there is one."""
    try:
        import ctypes
        buffer = ctypes.create_unicode_buffer(1024)
        if ctypes.windll.kernel32.GetShortPathNameW(path, buffer, 1024):
            return buffer.value or path
    except Exception:
        pass
    return path


def _heal_keyboard():
    try:
        import ctypes
        ctypes.windll.user32.keybd_event(0x12, 0, 0x0002, 0)
        ctypes.windll.user32.keybd_event(0x12, 0, 0x0002 | 0x0001, 0)
    except Exception:
        pass


_CALC_IDS = {"automatic": -4105, "semiautomatic": 2, "manual": -4135}


def render_cycle(symbol, book, sheet, token):
    """One render pass for a symbol: read sheet, build chain, write to Excel."""
    try:
        pre = sheet.range((1, 1), (4, 31)).options(transpose=False).value
    except Exception:
        pre = None

    update_user_dates(symbol, book, sheet, pre)

    # Expiry dropdown (refreshes when the available-expiry list changes)
    _sig = tuple(available_dates.get(symbol) or [])
    if _sig and _dropdown_sig.get(symbol) != _sig:
        setup_expiry_dropdown(sheet, symbol)
        _dropdown_sig[symbol] = _sig

    # Row 2 hint — re-write only if missing
    try:
        b2 = sheet.range("B2")
        if cell_value(pre, sheet, 2, 2) != "Total strikes — edit A2":
            b2.value = "Total strikes — edit A2"
            b2.font.bold = True
            b2.font.color = (89, 89, 89)
    except Exception as e:
        log_excel_exception(symbol, e, "updating worksheet instruction")

    # User spot limit from A2 (empty = show all strikes)
    v = cell_value(pre, sheet, 2, 1)
    if isinstance(v, str):
        try:
            v = float(v.strip())
        except ValueError:
            v = None
    user_provided = isinstance(v, (int, float)) and v >= 1
    range_val = 9999 if not user_provided else max(int(v) // 2, 1)

    # Spot: live price, else parity-implied, else stale
    sp = spot_cmp.get(symbol)
    if (not sp or spot_is_stale(symbol)):
        parity = parity_spot_or_none(symbol)
        if parity:
            sp = parity

    expiries = active_expiries.get(symbol) or []
    pending = []
    for bi, block in enumerate(BLOCK_COLS):
        if bi >= n_blocks(symbol):
            continue
        expiry = expiries[bi] if bi < len(expiries) else None
        if not expiry:
            continue
        rows, pcr, n, fresh = build_chain(symbol, expiry, token, spot=sp, range_val=range_val)
        stale = False
        if not fresh:
            # Market closed / feed empty: show last known good data marked STALE.
            # On the very first run there is nothing cached yet, so fall back to
            # the zero-price strike ladder just built so the trader still sees
            # the structure instead of a blank sheet.
            lrows, lpcr = get_last_good_chain(symbol, expiry)
            if lrows:
                rows, pcr, n = lrows, lpcr, len(lrows)
                stale = True
            elif not rows:
                continue  # no contracts at all for this expiry
            else:
                stale = True  # fresh quotes all zero (market closed)

        col_l = block["start"]
        spot_disp = f"{sp:.2f}" if sp is not None else "--"
        spot_text = f"{symbol} Spot: {spot_disp} | {n} strikes"
        # Stale-data announcement: silent normally so the header cell is not
        # rewritten every cycle; a trader reading frozen data needs to know.
        parts = []
        feed = net_badge()
        if feed:
            parts.append(feed)
        if spot_is_stale(symbol):
            try:
                age = spot_age(symbol)
            except Exception:
                age = float("inf")
            if age is not None and age != float("inf") and age > STALE_AFTER_S:
                parts.append(f"SPOT STALE ({int(age)}s)")
        if stale:
            try:
                age = last_good_age(symbol, expiry)
            except Exception:
                age = float("inf")
            if age != float("inf"):
                parts.append(f"DATA STALE ({int(age)}s)")
            else:
                parts.append("DATA STALE")
        if parts:
            spot_text = f"{spot_text} | {' | '.join(parts)}"
        pcr_text = f"PCR: {pcr:.2f}"

        try:
            if cell_value(pre, sheet, 1, col_l) != spot_text:
                sheet.range((1, col_l)).value = spot_text
                mcx_state._cycle_wrote = True
            if cell_value(pre, sheet, 1, col_l + 4) != pcr_text:
                sheet.range((1, col_l + 4)).value = pcr_text
                mcx_state._cycle_wrote = True
        except Exception as e:
            log_excel_exception(symbol, e, "writing chain header")

        extent = ROW_EXTENT_DEFAULT
        padded = pad_block_rows(rows, extent)
        pending.append((bi, block, padded, extent, [(0, len(rows) - 1)], None))
        net_paint(sheet, symbol, bi, col_l)

    if pending:
        _flush_block_writes(sheet, symbol, pending)
        force_usd_block_formats(sheet, BLOCK_COLS[0], symbol, 0, ROW_EXTENT_DEFAULT)

    return bool(pending)


async def main_async(token):
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        asyncio.get_running_loop().set_exception_handler(
            lambda loop, context: None
        )
    except Exception:
        pass
    if os.name == "nt":
        _heal_keyboard()
    for symbol in MCX_SYMBOLS:
        mcx_state.excel_connected[symbol] = False
    try:
        open_excel_files()
    except Exception as e:
        print(f"open_excel_files failed: {type(e).__name__}: {e}", flush=True)
        append_app_exception("opening Excel workbooks", e)
    import xlwings as xw
    books = {}
    for symbol in MCX_SYMBOLS:
        path = os.path.join(BASE_DIR, f"{symbol.lower()}_chain.xlsx")
        book = find_open_book(xw, path)
        books[symbol] = book
    # Discovery of instruments up front
    try:
        fetch_instruments(token)
    except Exception as e:
        print(f"instrument discovery failed: {e}", flush=True)
    await asyncio.gather(
        *[expiry_roller(symbol) for symbol in MCX_SYMBOLS],
        *[date_loader() for symbol in MCX_SYMBOLS],
        display_loop(token, books),
    )


def _run_doctor():
    from doctor import run_doctor
    report_path = os.path.join(BASE_DIR, "doctor-report.txt")
    try:
        report, ok = run_doctor(BASE_DIR, out=report_path)
    except Exception as e:
        report = f"the self-test itself failed: {type(e).__name__}: {e}\n"
        ok = False
        try:
            with open(report_path, "w", encoding="utf-8") as handle:
                handle.write(report)
        except OSError:
            pass
    try:
        os.startfile(report_path)
    except OSError:
        pass
    sys.exit(0 if ok else 1)


async def display_loop(token, books):
    """Main render loop: cycle through symbols, write to Excel, sleep adaptively."""
    last_spot_poll = {symbol: 0.0 for symbol in MCX_SYMBOLS}
    last_oi_poll = {symbol: 0.0 for symbol in MCX_SYMBOLS}
    while True:
        cycle_start = time.monotonic()
        wrote_anything = False
        for symbol in MCX_SYMBOLS:
            book = books.get(symbol)
            if book is None:
                continue
            try:
                sheet = book.sheets[0]
            except Exception:
                continue
            try:
                if sheet_structure_ok(sheet):
                    if render_cycle(symbol, book, sheet, token):
                        wrote_anything = True
                        mcx_state.excel_connected[symbol] = True
                    else:
                        mcx_state.excel_miss_counts[symbol] += 1
                else:
                    mcx_state.excel_miss_counts[symbol] += 1
            except Exception as e:
                log_excel_exception(symbol, e, "render cycle")
                mcx_state.excel_miss_counts[symbol] += 1
        # Spot + OI background pollers (throttled)
        now = time.monotonic()
        for symbol in MCX_SYMBOLS:
            if now - last_spot_poll[symbol] >= SPOT_POLL_SECONDS:
                try:
                    await asyncio.to_thread(spot_poller, symbol, token)
                except Exception:
                    pass
                last_spot_poll[symbol] = now
            if now - last_oi_poll[symbol] >= OI_POLL_SECONDS:
                try:
                    await asyncio.to_thread(oi_poller, symbol, token)
                except Exception:
                    pass
                last_oi_poll[symbol] = now
        mcx_state._cycle_wrote = False
        # Adaptive sleep
        elapsed = time.monotonic() - cycle_start
        sleep_for = max(0.1, 0.5 - elapsed)
        await asyncio.sleep(sleep_for)


def main():
    if any(arg in ("--doctor", "--selftest", "/doctor") for arg in sys.argv[1:]):
        _run_doctor()
        return
    try:
        write_startup_diagnostics()
    except OSError as e:
        show_error(f"Could not write startup diagnostics:\n{e}")
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
        _heal_keyboard()
        os._exit(1)
    feed_log(
        "startup",
        pid=os.getpid(),
        product=PRODUCT_ID,
        exe=PRODUCT_EXE,
        folder=BASE_DIR,
        build=build_id(),
    )
    load_dotenv()
    token = os.environ.get("UPSTOX_API_TOKEN", "")
    if not token:
        show_error(
            "Upstox API token required.\n\n"
            "Set UPSTOX_API_TOKEN in .env next to this executable.\n"
            "App will close."
        )
        _heal_keyboard()
        os._exit(1)
    asyncio.run(main_async(token))


if __name__ == "__main__":
    main()
