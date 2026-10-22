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
from us_excel import (
    display_loop, find_open_book, _set_calc_mode,
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
