"""
Delta Exchange Option Chain Tool (India endpoint)
Dual-asset: BTC + ETH simultaneously.
3-block layout with editable dates and dynamic strike range.
OI via REST poll (config oi_poll_seconds), spot via fast single-ticker poll
(config spot_poll_seconds), prices also via WebSocket, Excel update every 100ms.
"""

import asyncio
import atexit
import bisect
import calendar
import json
import math
import os
import random
import platform
import re
import socket
import subprocess
import sys
import threading
import time
from datetime import datetime, timezone, date

from delta_diagnostics import (
    append_app_exception,
    append_connection_event,
    append_performance_summary,
    build_id,
    write_startup_diagnostics,
)

_orig_getaddrinfo = socket.getaddrinfo

def _ipv4_getaddrinfo(*args, **kwargs):
    results = _orig_getaddrinfo(*args, **kwargs)
    v4 = [r for r in results if r[0] == socket.AF_INET]
    return v4 if v4 else results

socket.getaddrinfo = _ipv4_getaddrinfo

import pandas as pd
import requests
import websockets

import calc_mode_guard
import delta_state

from delta_state import (
    BASE_DIR,
    BLOCK_COLS,
    EXPIRY_REFRESH_SECONDS,
    HTTP_TIMEOUT,
    INDEX_ASSET,
    INDEX_CHANNEL,
    INDEX_SYMBOLS,
    INDEX_WS_URL,
    MANUAL_WATCHDOG_SECONDS,
    MONTHS,
    NET_FAIL_LIMIT,
    OI_POLL_SECONDS,
    PRODUCT_EXE,
    PRODUCT_ID,
    REST_BASE,
    SPOT_POLL_SECONDS,
    SPOT_TICKER_SYMBOLS,
    STALE_AFTER_S,
    STALE_FLAG_SECONDS,
    WS_URL,
    _CALC_CHURN_SUMMARY_SECONDS,
    _CALC_IDS,
    _CALC_ID_NAMES,
    _FIRST_SIGHT_BOUND,
    _auto_open_done,
    _block_hidden,
    _block_payload,
    _block_values,
    _book_sheet_cache,
    _colour_state,
    _date_req_sig,
    _dropdown_sig,
    _excel_connection_status,
    _last_excel_exception,
    _last_rows,
    _last_window_status,
    _net,
    _net_drawn,
    _pending_date_loads,
    _performance_metrics,
    _sheet_altered_notice,
    _sheet_cache,
    _usd_fmt_done,
    active_expiries,
    active_expiries_dmy,
    available_dates,
    cfg,
    excel_connected,
    excel_fail_notice_shown,
    excel_fail_since,
    excel_miss_counts,
    expiry_symbols,
    feed_state,
    last_spot_change,
    last_ws_recv,
     live_data,
     spot_cmp,
     spot_source,
     symbol_expiry,
     user_pinned,
     PARITY_RATE,
     PARITY_DIVIDEND_Q,
     PARITY_USE_N_STRIKES,
     PARITY_DEADBAND,
     PARITY_MAX_SPREAD,
     PARITY_MAX_SPREAD_PCT,
     _parity_spot_cache,
     _spot_divergence,
     _CALC_CHURN_SUMMARY_SECONDS,
)

# Identifies this tool inside the shared calculation-mode record, so Delta and
# the US tool running together in one Excel instance do not fight over it.

# The executable is named for the product, so the process name alone says which
# tool it is. The customer can see it in Task Manager, and the instance guard
# never has to guess the product from the folder.

# The most recent bulk-mode restore callable, so an exit path that never
# reaches render_cycle's finally can still put Excel back.

# Trader controls read from the sheet: pause and minimum interval.

# Set when a cycle actually put something into Excel, and when the workbook's
# own recalculation was timed. Nothing here is a guess: both come from the
# write path and from a real Excel.Calculate() call.
# When the last "speed" line was written to the feed log.


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



# --- connection health indicator ------------------------------------------
# Internet ghatne par sheet freeze ho jata hai; user ko pata chalna chahiye
# ki data stale hai. 3 consecutive failures ke baad header me red tag dikhta
# hai, aur connection wapas aate hi apne aap hat jaata hai.


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


# --- quiet console retries ----------------------------------------------------
# The per-poll error prints were a wall of identical HTTPS stacks during every
# brief exchange hiccup - a customer sees ten red lines in two minutes and
# panics, when the poller was already retrying and recovering on its own. The
# console now gets ONE calm line per site per 2-minute window; every occurrence
# still lands in feed.log with its full detail for support.

_console_warn_last = {}


def _quiet_console_error(site, detail):
    """Rate-limit one site's console error line; keep every detail in feed.log."""
    now = time.monotonic()
    if now - _console_warn_last.get(site, 0.0) > 120.0:
        print(f"{site}: delta.exchange is slow right now - retrying "
              "(full detail in feed.log)", flush=True)
        _console_warn_last[site] = now
    feed_log("network_retry", site=site, detail=str(detail)[:160])


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
from delta_licensing import (
    ensure_license,
    license_is_valid,
    show_error,
)









# Resolved (book, sheet) pairs. Accessing book.sheets[0] costs a COM round-trip
# and a late-binding Dispatch every single call, which dominated the render
# loop. Cached, and dropped on any COM error so a closed workbook or a renamed
# sheet recovers instead of failing silently.
# Excel App COM object, reused across cycles for the bulk-mode calls.
# How far below the live table to sweep the first time we see a block in this
# run. _last_rows is in-memory, so a workbook left dirty by an earlier run is
# otherwise never cleaned up. Generous enough for any strike count the customer
# can ask for, and paid for once per block per run rather than every cycle.
# Per (asset, block) payload signature, so an unchanged chain is not rewritten
# to Excel at all.
# Per (asset, block) previous rows, so a change inside the block can be
# written as one span that carries the unchanged rows instead of a
# whole-block repaint or one call per scattered row.
# Per (asset, block) signature of the ITM/ATM picture. Repainting is four
# Excel calls per block, so it only happens when the picture would change.




def n_blocks(asset):
    """Every sheet shows two expiry blocks — near expiry and the next one.

    XAUT already only had 2 blocks (it rarely lists a third expiry with
    contracts); BTC and ETH now match, like the US tool from this build.
    The render loop clears and hides each asset's third expiry columns
    (W:AE) on its next cycle — one-time, then they are never touched.
    """
    return 2


_retired_blocks = set()


def _retire_unused_blocks(sheet, asset):
    """Clear + hide blocks beyond n_blocks the moment a sheet is reachable.

    This used to wait for the render loop's first payload-driven block
    sweep, which lands a few minutes into a cold start — long enough for
    the customer to see the previous session's third expiry still holding
    data. Connect time is what the customer calls "the tool started".
    """
    for bi, block in enumerate(BLOCK_COLS):
        if bi >= n_blocks(asset):
            key = (asset, bi, id(sheet))
            if key in _retired_blocks:
                continue
            try:
                cs = get_column_letter(block["start"])
                ce = get_column_letter(block["start"] + 8)
                sheet.range(f"{cs}1:{ce}1000").clear()
                sheet.range(f"{cs}:{ce}").api.EntireColumn.Hidden = True
                _last_rows[(asset, bi)] = 0
                _block_hidden[(asset, bi)] = True
                _retired_blocks.add(key)
                feed_log(
                    "expiry_block_retired",
                    asset=asset,
                    shown=n_blocks(asset),
                    cleared=[block["start"]],
                )
            except Exception as e:
                log_excel_exception(asset, e, f"retiring unused block {bi + 1}")


def safe_float(val):
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def fmt_oi(val):
    """Coerce OI to a plain float for Excel (so VLOOKUP/XLOOKUP works).

    Handles raw numbers, old Rs/INR text, and legacy "$1.50M"/"$2.00K" strings
    from a previous build so the one-off bulk cleanup passes.
    """
    if isinstance(val, str):
        s = val.replace("₹", "").replace("Rs", "").replace("INR", "").replace("$", "").replace(",", "").strip()
        s_lower = s.lower().rstrip()
        if s_lower.endswith("m"):
            try:
                return float(s_lower[:-1]) * 1_000_000
            except ValueError:
                return 0.0
        if s_lower.endswith("k"):
            try:
                return float(s_lower[:-1]) * 1_000
            except ValueError:
                return 0.0
        try:
            val = float(s)
        except ValueError:
            return 0.0
    if val is None or val != val:
        return 0.0
    return float(val)


def get_column_letter(col):
    """Convert 1-indexed column number to Excel letter (A, B, ..., Z, AA, AB, ...)"""
    result = ""
    while col > 0:
        col -= 1
        result = chr(65 + col % 26) + result
        col //= 26
    return result






def record_excel_connection(asset, status, detail=""):
    if _excel_connection_status.get(asset) == status:
        return
    _excel_connection_status[asset] = status
    try:
        append_connection_event(asset, status, detail)
    except OSError as e:
        print(f"Could not log Excel connection state ({asset}): {e}", flush=True)


def log_excel_exception(asset, exc, context="Excel/COM update"):
    now = time.monotonic()
    signature = (type(exc).__name__, str(exc))
    key = (asset, context)
    previous = _last_excel_exception.get(key)
    if previous and previous[0] == signature and now - previous[1] < 60:
        return
    _last_excel_exception[key] = (signature, now)
    try:
        append_app_exception(f"{context} for {asset}", exc)
    except OSError as log_error:
        print(f"Could not log {context} traceback ({asset}): {log_error}", flush=True)


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
















        # loop back -> dialog reopens for another attempt


def sheet_structure_ok(sheet):
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


AUTO_OPEN_TIMEOUT_SECONDS = 30.0


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
    so the open runs on the display thread (inline) and the render loop picks
    the workbooks up as soon as they are there.

    The open pass is also guarded now. Workbooks.Open aimed at an Excel this
    process does not own - the other Option Chain tool's instance, or one
    whose controlling process died earlier - can block forever, because a
    busy Excel answers automation by simply not answering. A wedged pass is
    retried once from the display loop after AUTO_OPEN_TIMEOUT_SECONDS (see
    retry_open_pass_after_grace below).

    Inline is not just a simplification - it is correctness. Creating the
    private instance on a one-shot thread while the display loop read that
    App object made every probe from the display thread answer "The
    application called an interface that was marshalled for a different
    thread": the books never attached and the close-all timer ended the
    tool with Excel right there on the screen. One thread owns COM here,
    and it is the thread that writes.
    """
    if platform.system() != "Windows":
        return
    delta_state._open_pass_at = time.monotonic()
    _open_excel_files()


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


def _ensure_owner_app(xw):
    """Return the tool's private Excel instance, creating it when needed.

    Every chain workbook opens in this instance and nothing else attaches to
    any other Excel. That is what makes the US and the Delta chains
    independent on one machine: separate instances cannot contend over
    calculation mode, cannot hand workbooks to each other, and a busy dialog
    in one cannot block the other.
    """
    app = delta_state._owner_app
    if app is not None:
        try:
            _ = app.version
            return app
        except Exception:
            delta_state._owner_app = None
            delta_state._owner_hwnd = None
            delta_state._app_cache = None
    try:
        app = xw.App(visible=True)
    except Exception as e:
        print(f"Could not start Excel: {e}", flush=True)
        return None
    delta_state._owner_app = app
    delta_state._app_cache = app
    try:
        delta_state._owner_hwnd = app.api.Hwnd
    except Exception:
        delta_state._owner_hwnd = None
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

        # Nothing to choose between first pass and retry: both use (or
        # create) the private instance. The retry simply catches a pass that
        # died before the books made it into Excel.
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
        for asset in ("BTC", "ETH", "XAUT"):
            path = os.path.join(delta_state.BASE_DIR, f"{asset.lower()}_chain.xlsx")
            if not os.path.exists(path):
                record_excel_connection(asset, "workbook_file_missing", f"path={path}")
                show_error(f"File not found:\n{path}\n\nKeep this file next to OptionChain.exe and run again.")
                continue
            try:
                os.remove(path + ":Zone.Identifier")
            except OSError:
                pass
            # Already open? The customer opened it themselves, so the tool must
            # never open it again behind their back, and the file handle is
            # already theirs.
            try:
                if find_open_book(xw, path) is not None:
                    _auto_open_done[asset] = True
                    feed_log("excel_auto_open_done", asset=asset, via="already_open")
                    continue
            except Exception:
                pass
            target = path if len(path) <= 120 else short_path(path)
            try:
                app.api.Workbooks.Open(target)
                # We have now opened it. If the customer closes it later, it
                # stays closed.
                _auto_open_done[asset] = True
                feed_log("excel_auto_open_done", asset=asset, via="api")
            except Exception as e:
                print(f"Could not open {asset.lower()}_chain.xlsx: {e}", flush=True)
                feed_log("excel_auto_open_failed", asset=asset, error=str(e))
                show_error(
                    f"Could not open {asset.lower()}_chain.xlsx:\n\n{e}\n\n"
                    "If the package was unzipped into a long path, move it to "
                    "a short one - C:\\OptionChain - and run the tool again."
                )
                try:
                    append_app_exception(f"opening {asset} workbook", e)
                except OSError as log_error:
                    print(f"Could not log workbook open failure ({asset}): {log_error}", flush=True)
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
                    delta_state._owner_hwnd = _new
        except Exception:
            pass
        if started_excel:
            _close_empty_book1(app)


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


_OPEN_RETRY_GRACE_SECONDS = 45.0


def retry_open_pass_after_grace():
    """Retry a wedged startup open pass once, on the display thread.

    The pass runs inline at startup. If Excel answered everything it takes
    under two seconds and all three books latch; only a wedged pass leaves
    nothing latched, and that is worth exactly one retry after the grace
    period - never more, because a workbook the customer closed on purpose
    must stay closed.
    """
    if delta_state._open_pass_retried or delta_state._open_pass_at is None:
        return
    now = time.monotonic()
    if now - delta_state._open_pass_at < _OPEN_RETRY_GRACE_SECONDS:
        return
    delta_state._open_pass_retried = True
    if all(delta_state._auto_open_done.get(asset) for asset in ("BTC", "ETH", "XAUT")):
        return
    feed_log(
        "excel_open_retried",
        note=(f"nothing latched {int(_OPEN_RETRY_GRACE_SECONDS)}s after the startup "
              "open pass; retrying it once"),
    )
    open_excel_files()


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


# ---------------------------------------------------------------------------
# Feed observability.
#
# The tool used to fail silently: a REST error became "0 contracts", a block
# lost its expiry, the header showed "--" and the trader was never told why.
# Every request and every probe now lands in a size-capped log, and the state
# is surfaced in the sheet header.
# ---------------------------------------------------------------------------
# The log directory honours DELTA_OPTION_CHAIN_LOG_DIR so tests and support
# runs never write into the real customer logs.
from delta_diagnostics import (
    _flat,
    feed_log,
)

# asset -> {state, detail, retry_in, last_good_ts, last_good_label}






def set_feed_state(asset, state, detail="", retry_in=0.0):
    st = feed_state.setdefault(asset, {})
    if st.get("state") == state and st.get("detail") == detail:
        # The same state is being reported again. Writing it to the log every
        # time would put a line in the log every poll for as long as the
        # failure lasts.
        st["retry_in"] = retry_in
        return
    st["state"] = state
    st["detail"] = detail
    st["retry_in"] = retry_in
    feed_log("state", asset=asset, state=state, detail=detail,
             retry_in=round(retry_in, 1))


def note_feed_good(asset, label=""):
    st = feed_state.setdefault(asset, {})
    st["state"] = "OK"
    st["detail"] = ""
    st["retry_in"] = 0.0
    st["last_good_ts"] = time.time()
    if label:
        st["last_good_label"] = label


def feed_header_note(asset):
    """Text for the sheet header, or "" when there is nothing to say."""
    st = feed_state.get(asset) or {}
    state = st.get("state", "OK")
    if state == "OK":
        age = time.time() - (st.get("last_good_ts") or 0)
        if st.get("last_good_ts") and age > STALE_FLAG_SECONDS:
            return f"STALE: last good {st.get('last_good_label', '')}".strip()
        return ""
    last = st.get("last_good_label", "never")
    return f"{state}: {st.get('detail', '')} (last good {last})".strip()


def source_delay_note(asset):
    """Everything the header should say about how fresh the data is.

    Two things are worth saying: that the feed is degraded (a REST call
    failed, the socket dropped), and that the index quote has gone quiet.
    Both are silent otherwise, and a trader looking at a frozen spot has no
    way of telling a quiet market from a broken feed.

    Returns "" in the normal case, so the header text does not change and
    the cell is not rewritten on every cycle.
    """
    parts = []
    feed = feed_header_note(asset)
    if feed:
        parts.append(feed)
    try:
        age = spot_age(asset)
    except Exception:
        age = float("inf")
    if age != float("inf") and age > STALE_AFTER_S:
        parts.append(f"SPOT STALE ({int(age)}s)")
        if _parity_spot_cache.get(asset) is not None:
            parts.append(f"PARITY SPOT active")
    return " | ".join(parts)


def _get_json(url, params):
    """GET with keep-alive session, 429 backoff and one socket-reset retry.

    Every failure is now logged with enough detail to act on: endpoint, params,
    HTTP status, exception type, the head of the body and how long it took.
    A 429 honours Retry-After instead of guessing.
    """
    last = None
    started = time.perf_counter()
    for attempt in range(3):
        try:
            resp = _session().get(url, params=params, timeout=HTTP_TIMEOUT)
            if resp.status_code == 429:
                retry_after = resp.headers.get("Retry-After")
                try:
                    wait = float(retry_after)
                except (TypeError, ValueError):
                    wait = 1.5 * (attempt + 1)
                wait = min(max(wait, 0.5), 30.0)
                feed_log("rest_429", url=url.split("/v2/")[-1],
                         params=_flat(params), attempt=attempt + 1,
                         retry_after=retry_after, wait_s=wait)
                last = requests.HTTPError("429 Too Many Requests")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp.json()
        except (requests.ConnectionError, requests.exceptions.ChunkedEncodingError) as e:
            last = e
            feed_log("rest_conn", url=url.split("/v2/")[-1], params=_flat(params),
                     attempt=attempt + 1, error=type(e).__name__, detail=str(e)[:180])
            time.sleep(0.4)
        except Exception as e:
            last = e
            body = ""
            status = None
            r = locals().get("resp")
            if r is not None:
                status = getattr(r, "status_code", None)
                try:
                    body = r.text[:180]
                except Exception:
                    body = ""
            feed_log("rest_error", url=url.split("/v2/")[-1], params=_flat(params),
                     attempt=attempt + 1, status=status,
                     error=type(e).__name__, detail=str(e)[:180], body=body,
                     elapsed_ms=round((time.perf_counter() - started) * 1000, 1))
            break
    feed_log("rest_failed", url=url.split("/v2/")[-1], params=_flat(params),
             error=type(last).__name__ if last else "unknown",
             elapsed_ms=round((time.perf_counter() - started) * 1000, 1))
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


def _today_settlement(products, today_str):
    """When today's expiry actually settles, as an aware UTC datetime.

    Delta settles BTC and ETH options at 12:00 UTC and XAUT at 16:00 UTC, and
    it changes them without warning. Reading the time off the products is
    safer than hard-coding an hour: an expiry that disappears from the sheet
    before it has really expired is a trader's contract vanishing in the
    middle of a session, with three hours of trading left on it.
    """
    stamps = []
    for p in products:
        st = p.get("settlement_time") or ""
        if len(st) < 20 or st[:10] != today_str or not st.endswith("Z"):
            continue
        try:
            stamps.append(
                datetime.strptime(st[:19], "%Y-%m-%dT%H:%M:%S").replace(
                    tzinfo=timezone.utc
                )
            )
        except ValueError:
            continue
    return min(stamps) if stamps else None


def available_expiry_dates(asset, products, now=None):
    now = now or datetime.now(timezone.utc)
    today_str = now.strftime("%Y-%m-%d")

    dates = sorted(set(
        p["settlement_time"][:10] for p in products
        if p["settlement_time"][:10] >= today_str
    ))

    # Today's expiry stays in the list until it has actually settled.
    settle = _today_settlement(products, today_str)
    if settle is not None and now >= settle and dates and dates[0] == today_str:
        dates = dates[1:]

    return dates


def get_default_expiries(asset, products):
    """The first expiries the slots start with — always n_blocks wide.

    Every slot list must carry exactly n_blocks entries or the roller dies on
    its first write ("list assignment index out of range"): the loop guards
    the reads but the writes at the end of each blocked pass are plain
    indexing. Slots without a real expiry stay None until the roller fills
    them, and the writer renders "Spot: --" for the empty block instead of
    crashing.
    """
    dates = available_expiry_dates(asset, products)[:n_blocks(asset)]
    while len(dates) < n_blocks(asset):
        dates.append(None)
    return dates


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
        # A single malformed row must never abort the whole batch: this runs on
        # the OI poll loop, so an unhandled exception here would silently kill
        # updates for that asset for the rest of the session.
        strike = safe_float(t.get("strike_price"))
        if strike <= 0:
            continue
        sp = t.get("spot_price")
        if sp is not None:
            spot_cmp[asset] = safe_float(sp)
        ltp = safe_float(t.get("close"))
        if ltp == 0.0:
            ltp = safe_float(t.get("mark_price"))
        quotes = t.get("quotes") or {}
        entry = {
            "strike": int(strike),
            "type": "call" if ct == "call_options" else "put",
            "ltp": ltp,
            "bid": safe_float(quotes.get("best_bid")),
            "ask": safe_float(quotes.get("best_ask")),
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
    """Contracts for one expiry.

    Returns (count, state). 'OK' means a real count. 'ERROR' means the request
    failed and the count is meaningless -- these must never be confused, which
    is exactly how a transient failure used to look like 'no contracts'.
    """
    try:
        tickers = fetch_tickers(asset, iso_to_dmy(expiry_iso))
        n = len(tickers)
        # HTTP 200 is not proof of good data: check the symbols really belong to
        # this expiry before trusting the count.
        _yy, _mm, _dd = expiry_iso.split("-")
        suffix = f"{_dd}{_mm}{_yy[2:]}"          # symbols carry DDMMYY
        ok_suffix = sum(1 for t in tickers if suffix in str(t.get("symbol", "")))
        matched = (ok_suffix == n) if n else True
        feed_log("probe_ok", asset=asset, expiry=expiry_iso, rows=n,
                 suffix_ok=ok_suffix, suffix_checked=matched)
        if n == 0:
            feed_log("probe_no_contracts", asset=asset, expiry=expiry_iso, rows=0)
            return 0, "NO_CONTRACTS"
        if not matched:
            feed_log("probe_suffix_mismatch", asset=asset, expiry=expiry_iso,
                     rows=n, suffix_ok=ok_suffix, suffix=suffix, level="warn")
        return n, "OK"
    except Exception as e:
        feed_log("probe_error", asset=asset, expiry=expiry_iso,
                 error=type(e).__name__, detail=str(e)[:180])
        return 0, "ERROR"


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
                        _quiet_console_error(
                            f"OI poller error ({asset} {dmy})", res
                        )
                        net_note(False)
                        set_feed_state(asset, "ERROR", str(res)[:80],
                                       retry_in=OI_POLL_SECONDS)
                        continue
                    apply_tickers(asset, res, dmy_to_iso(dmy))
                    net_note(True)
                    # This is the feed's heartbeat, and it is what the sheet
                    # header reads. The state used to be cleared only when an
                    # expiry rolled, so a perfectly healthy feed sat at
                    # LOADING for the whole session and the header said so on
                    # every row.
                    note_feed_good(asset, dmy_to_iso(dmy))
        except Exception as e:
            _quiet_console_error(f"OI poller error ({asset})", e)
            net_note(False)
            set_feed_state(asset, "ERROR", type(e).__name__,
                           retry_in=OI_POLL_SECONDS)
        await asyncio.sleep(OI_POLL_SECONDS)


# Perp tickers, used ONLY as a fallback for the index price. On these the
# 'close' field is the perpetual's own last trade (~30-45 USD from the index),
# so the website's number must never be sourced from here when a live option
# ticker is available.

# Dedicated index channel. Delta pushes the index here every ~0.25 s, versus
# ~4.9 s when it only rides along inside v2/ticker option frames (it refreshes
# only when an option prints). Symbols come from GET /v2/indices and are NOT
# derivable from the ticker name — XAUT's is .DEXAUTUSD, not .DEXXAUTUSD.

# Freshness tracking. Deliberately keyed on WebSocket RECIEVE time and on
# spot_price changes, NOT on per-option 'timestamp' — illiquid strikes can carry
# an old server timestamp while the index is moving, which would read as false
# staleness.


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


def _years_to_expiry(expiry_iso):
    """Days to expiry as a year fraction, same shape as the US tool."""
    if not expiry_iso:
        return None
    try:
        y, m, d = expiry_iso[:10].split("-")
        exp = date(int(y), int(m), int(d))
    except Exception:
        return None
    today = date.today()
    return max(0.0, (exp - today).days / 365.0)


def _parity_spot(asset, expiry_iso, spot_hint):
    """Index spot from put-call parity on our own live chain.

    Parity:  C - P = S*exp(-qT) - K*exp(-rT)
    =>  S = K*exp(-(r-q)T) + (C-P)*exp(qT)

    For crypto options q=0 (no dividend), r=PARITY_RATE. The median of the N
    strikes nearest the current spot estimate is taken, rejecting crossed and
    wide-spread quotes. A deadband smooths the cached value.
    """
    if not expiry_iso:
        return None
    ld = live_data.get(asset, {})
    if not ld:
        return None
    T = _years_to_expiry(expiry_iso)
    if T is None or T <= 0:
        return None
    forward_discount = math.exp(-(PARITY_RATE - PARITY_DIVIDEND_Q) * T)
    call_put_adj = math.exp(PARITY_DIVIDEND_Q * T)

    calls = {}
    puts = {}
    for sym, data in ld.items():
        se = symbol_expiry.get(asset, {}).get(sym)
        if se != expiry_iso:
            continue
        strike = data.get("strike")
        if not strike or strike <= 0:
            continue
        if data.get("type") == "call":
            calls[strike] = data
        elif data.get("type") == "put":
            puts[strike] = data

    common = set(calls) & set(puts)
    if not common:
        return None

    def mid(row):
        bid = row.get("bid") or 0.0
        ask = row.get("ask") or 0.0
        if bid > 0 and ask > 0:
            return (bid + ask) / 2.0
        return row.get("ltp") or 0.0

    def spread_ok(row):
        bid = row.get("bid") or 0.0
        ask = row.get("ask") or 0.0
        if bid <= 0 or ask <= 0 or ask < bid:
            return False
        mid_price = (bid + ask) / 2.0
        if mid_price <= 0:
            return False
        return (ask - bid) <= max(PARITY_MAX_SPREAD, mid_price * PARITY_MAX_SPREAD_PCT)

    def parity_for_strike(strike):
        c = calls[strike]
        p = puts[strike]
        if not spread_ok(c) or not spread_ok(p):
            return None
        call_mid = mid(c)
        put_mid = mid(p)
        if call_mid <= 0 or put_mid <= 0:
            return None
        return strike * forward_discount + (call_mid - put_mid) * call_put_adj

    hint = spot_hint if spot_hint else 0.0
    ranked = sorted(common, key=lambda s: abs(s - hint))
    take = ranked[:PARITY_USE_N_STRIKES]
    estimates = [est for s in take if (est := parity_for_strike(s)) is not None]
    if not estimates:
        return None
    estimates.sort()
    n = len(estimates)
    if n % 2 == 1:
        parity_val = estimates[n // 2]
    else:
        parity_val = (estimates[n // 2 - 1] + estimates[n // 2]) / 2.0

    cached = _parity_spot_cache.get(asset)
    if cached is not None and abs(parity_val - cached) < PARITY_DEADBAND:
        return cached
    _parity_spot_cache[asset] = parity_val
    return parity_val


def parity_spot_or_none(asset):
    """Best parity spot for the active expiry, or None."""
    expiry = active_expiries[asset][0] if active_expiries[asset] else None
    if not expiry:
        return None
    hint = spot_cmp.get(asset) or 0.0
    return _parity_spot(asset, expiry, hint)


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
            _quiet_console_error(f"Spot poller error ({asset})", e)
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
            states = [st for _, st in counts]
            nums = [n for n, _ in counts]
            with_data = [d for d, n in zip(probe_dates, nums) if n > 0]

            # Retry once with jittered backoff when every probe failed. A single
            # bad round must not be allowed to strip the sheet.
            feed_log("roller_decision", asset=asset, probe_dates=probe_dates,
                     nums=nums, states=states, with_data=with_data,
                     active_before=list(active_expiries[asset]),
                     n_blocks=n_blocks(asset))

            if not with_data and "ERROR" in states:
                feed_log("roller_all_probes_failed", asset=asset,
                         expiry_count=len(probe_dates), states=states)
                await asyncio.sleep(2.0 + random.uniform(0, 1.5))
                counts = await asyncio.gather(
                    *[asyncio.to_thread(probe_expiry_count, asset, d)
                      for d in probe_dates]
                )
                states = [st for _, st in counts]
                nums = [n for n, _ in counts]
                with_data = [d for d, n in zip(probe_dates, nums) if n > 0]
                feed_log("roller_retry_result", asset=asset,
                         recovered=bool(with_data), rows=nums)

            # A probe that failed must not shift the slots.
            #
            # The slots are filled from the list of expiries that have data,
            # so one missing expiry slides every slot along. The ladder then
            # swaps to a different expiry - and swaps back on the next pass,
            # when the probe works. The trader watches the chain change for
            # no reason, and the expiry they had pinned moves out from under
            # them.
            #
            # So if any probe failed, keep what is on screen and try again in
            # a few seconds.
            if "ERROR" in states:
                feed_log("roller_probe_incomplete", asset=asset,
                         states=states, active=list(active_expiries[asset]),
                         note="keeping the current expiry assignment")
                continue

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
                if not with_data and cur:
                    # Every probe failed. Keep the expiry we already have: the
                    # trader keeps the last good chain on screen (flagged stale)
                    # instead of watching the sheet blank out for a timeout.
                    set_feed_state(asset, "ERROR",
                                   "no expiry data from Delta", retry_in=0.0)
                    continue
                if tgt:
                    await asyncio.to_thread(load_expiry_tickers, asset, tgt)
                    note_feed_good(asset, tgt)
                # The slot lists must actually be this long before indexed
                # writes: assets whose defaults were created narrower (XAUT
                # ships with one real expiry) would otherwise blow up every
                # roller pass with "list assignment index out of range".
                while len(active_expiries[asset]) <= i:
                    active_expiries[asset].append(None)
                while len(active_expiries_dmy[asset]) <= i:
                    active_expiries_dmy[asset].append(None)
                changed_slots.append((i, cur, tgt))
                active_expiries[asset][i] = tgt
                active_expiries_dmy[asset][i] = iso_to_dmy(tgt) if tgt else None

            feed_log("roller_result", asset=asset,
                     active_after=list(active_expiries[asset]),
                     changed=len(changed_slots))

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
            _quiet_console_error(f"Expiry roller error ({asset})", e)


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
            _quiet_console_error("Index listener error", e)
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
    strike = safe_float(data.get("strike_price"))
    if strike <= 0:
        return
    ltp = safe_float(data.get("close"))
    if ltp == 0.0:
        ltp = safe_float(data.get("mark_price"))
    quotes = data.get("quotes") or {}
    bid = safe_float(quotes.get("best_bid"))
    ask = safe_float(quotes.get("best_ask"))
    if symbol not in live_data[asset]:
        live_data[asset][symbol] = {
            "strike": int(strike),
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


def update_user_dates(asset, book, sheet, pre=None):
    """Read date cells only — HTTP loads are queued to date_loader (never blocks event loop)."""
    for i, col3 in enumerate((4, 15, 26)):
        if i >= n_blocks(asset):
            # block not shown (e.g. XAUT 3rd) — keep its date cell empty
            rng = sheet.range((3, col3))
            if cell_value(pre, sheet, 3, col3) not in (None, ""):
                _set_date_cell(rng, "")
            user_pinned[asset][i] = False
            _date_req_sig.pop((asset, i), None)
            continue
        rng = sheet.range((3, col3))
        val = cell_value(pre, sheet, 3, col3)
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
    except Exception as e:
        log_excel_exception("EXCEL", e, "expiry date number format")
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
            _quiet_console_error(f"Date fetch error ({asset} {parsed})", e)
            await asyncio.sleep(3)


# Set once the tool has asked Excel to open a workbook. After that it never
# asks again: a customer who closes the workbook must not watch it pop back
# up. Re-armed only if the file itself disappears, so a deleted workbook can
# still be recreated and opened.
# Set when the sheet's layout no longer matches what the tool expects, so the
# customer is told once rather than on every cycle.


_UNKNOWN_TTL_SECONDS = 10.0
_unknown_logged_at = {}
_CLOSE_ALL_EXIT_SECONDS = 75.0


def _owner_alive():
    """Cheap liveness probe on the tool's own Excel instance."""
    app = delta_state._owner_app
    if app is None:
        return False
    try:
        _ = app.version
        return True
    except Exception:
        return False


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
        for asset in ("BTC", "ETH", "XAUT"):
            delta_state._missing_since[asset] = now
    if len(delta_state._missing_since) == 3:
        now = time.monotonic()
        if delta_state._all_gone_since is None:
            delta_state._all_gone_since = now
        elif now - delta_state._all_gone_since >= _CLOSE_ALL_EXIT_SECONDS:
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
    elif delta_state._all_gone_since is not None:
        delta_state._all_gone_since = None
        delta_state._missing_since.clear()



# True only while THIS window holds Excel in manual (we set it and will
# put it back). The write path can go row-granularity when we own the
# mode; when somebody else owns the sheet everything degrades to one
# merged span so the customer's workbook recalc is hit exactly once.

# (monotonic_stamp, second-pass flag) for the post-restore
# CalculationState probe; the flag only needs one COM read per half
# minute.


def excel_process_running(max_age=2.0):
    """True if any EXCEL.EXE is alive, checked without touching COM.

    When Excel crashes or is killed, the cached COM server is a dead reference.
    The next `xw.apps` / `app.books` call then raises "RPC server is
    unavailable" and takes the whole process down, after which nothing writes
    at all. Checking the process table first lets us simply wait for the
    customer to reopen Excel instead of dying silently.

    The answer is cached briefly. Spawning tasklist per workbook per cycle
    added roughly a quarter of a second to every render pass, which is a lot to
    pay for a check that only needs to be approximately current.
    """
    now = time.monotonic()
    if delta_state._excel_liveness_cache is not None:
        stamp, running = delta_state._excel_liveness_cache
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
        delta_state._excel_liveness_cache = (now, True)
        return True
    running = "EXCEL.EXE" in (result.stdout or "").upper()
    delta_state._excel_liveness_cache = (now, running)
    return running


def _long_form(path):
    """The Windows long form of a path, or the path itself.

    Workbooks.Open was handed the 8.3 short name (long paths break Excel's
    open), and Excel then reports THAT form as the book's fullname:
    'C:\\Users\\Pravin\\DOWNLO~1\\...\\btc_chain.xlsx'. Comparing that against
    the tool's own long abspath never matched, find_open_book returned None
    every cycle, and the writer sat on rows=0 with three books wide open.
    """
    try:
        import ctypes

        buffer = ctypes.create_unicode_buffer(1024)
        if ctypes.windll.kernel32.GetLongPathNameW(path, buffer, 1024):
            return buffer.value or path
    except Exception:
        pass
    return path


def find_open_book(xw, fname):
    """Find the workbook inside the tool's PRIVATE Excel instance.

    The chain books live in the instance the tool started itself (the owner,
    see _ensure_owner_app), so scanning every Excel on the machine is not
    only wasted work now - it was the loophole that let the two chain tools
    open workbooks in each other's Excel and fight over calculation mode.
    One chain tool, one Excel instance, no cross-searching.

    When our instance is there but does not answer its workbook list (a
    dialog is up, the instance is being torn down), the failure is recorded
    as "unknown" and None returned: the caller must treat that as
    unanswerable this cycle, never as a missing workbook.
    """
    target = os.path.abspath(fname)
    target_long = _long_form(target).lower()
    target_name = os.path.basename(target).lower()
    asset = target_name.split("_chain", 1)[0].upper()
    if not excel_process_running():
        # Do not touch COM at all while Excel is gone.
        return None
    app = delta_state._owner_app
    if app is None:
        return None
    try:
        books = list(app.books)
    except Exception as e:
        log_excel_exception(asset, e, "enumerating the tool's own Excel workbooks")
        delta_state._enum_unknown[asset] = time.monotonic()
        return None
    for book in books:
        try:
            fullname = book.fullname
            base = os.path.basename(fullname).lower()
        except Exception as e:
            log_excel_exception(asset, e, "reading open workbook path")
            continue
        if base == target_name and _long_form(fullname).lower() == target_long:
            # Exact file, whatever form the path was opened in.
            return book
    return None


def _contiguous_runs(indices):
    """Merge changed-row indices into contiguous (start, end) runs.

    Writing once per contiguous run used to cost a COM call per run; each
    round-trip costs about 33ms, so a dozen scattered moves froze the
    sheet for roughly a third of a second. Runs give the writer one
    addressable span instead.
    """
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


def _deep_clear_below_block(sheet, asset, bi, c0, n_rows):
    """Clear junk an older build left below the fixed extent, if any.

    One read decides it; the clear only happens when the region is not
    already empty, so the normal case costs one cheap COM read per block
    per run.
    """
    if (asset, bi) in _deep_swept:
        return
    _deep_swept.add((asset, bi))
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
        feed_log("chain_deep_clear", asset=asset, block=bi + 1, from_row=deep_start)
    except Exception as e:
        log_excel_exception(asset, e, f"deep clearing stale rows in block {bi + 1}")


def cell_value(pre, sheet, row, col):
    """Read one cell from the merged 4x31 pre-read, falling back to COM.

    The sheet's status strip (A2/B2, the three date cells D3/O3/Z3 and
    the block header texts) fits inside rows 1-4, columns 1-31, so one
    block read per sheet replaces about a dozen individual cell reads
    every cycle. Individual reads only happen if the merged one failed.
    """
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


def force_usd_block_formats(sheet, block, asset, bi, n_rows):
    """Prices = 0.00, OI = numeric with $ display (kills the Rs number format).

    Formats land once per sheet, and the OI cleanup is one bulk read plus one
    bulk write whenever the sheet is first touched. The previous version
    walked the OI column one cell at a time - 1,206 COM round trips per
    600-row block - and re-ran every time the row count changed, which is
    where the 340-712 ms excel_formatting spikes came from.
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
        log_excel_exception(asset, e, "formatting option-chain cells")
        return
    # One-off bulk cleanup of values left as text or Rs-currency by an older
    # build.  One read and one write for the whole sheet, not per cell.
    for off in (3, 5):
        col = get_column_letter(c0 + off)
        try:
            rng = sheet.range(f"{col}5:{col}{end}")
            values = rng.value
        except Exception as e:
            log_excel_exception(asset, e, f"reading open-interest column {c0 + off}")
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
                log_excel_exception(asset, e, f"writing open-interest column {c0 + off}")


def _flush_block_writes(sheet, asset, pending):
    """Write everything the block loop queued, in as few calls as possible.

    Per block, protected runs write just the changed run and unprotected runs
    merge into one span whose unchanged rows carry their previous values.
    When every changed block is also consecutive, the blocks go out as ONE
    range: spacer columns become blank rows, gaps inside a block keep their
    current values, and nothing outside the changed blocks is touched.

    Two rules keep the combined write safe, and both were learned the hard way
    when the US tool tried the same thing and reverted it:

    - The grid carries the block's CURRENT content, never its previous content.
      Carrying the previous content repopulates rows that had just been
      cleared because the block shrank, and the sheet then shows stale strikes
      below the live table.
    - A block that did not change is never in the grid at all. Only the changed
      blocks take part, so an unchanged block's columns are never written.

    The combined rectangle is only used when it is not much bigger than the
    rows it replaces. A rectangle that spans two blocks whose changed rows sit
    far apart would dirty thousands of cells the customer never asked us to
    touch, and every dirty cell is recalculated when Excel is put back into
    automatic mode.

    Pending entries carry the block's padded rows (values padded with blank
    rows up to the fixed extent, never a raw DataFrame), so every write keeps
    the geometry the customer's literal ranges rely on.
    """
    if not pending:
        return
    pending = sorted(pending, key=lambda item: item[0])
    protected = delta_state._bulk_owning_manual
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
    flush_started = time.perf_counter()
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
                        # Current content, not previous content. A block that
                        # shrank must not have its old tail written back.
                        row[c_off : c_off + 9] = list(p["values"][r])
                grid.append(row)
            sheet.range((5 + min_row, first_c)).value = grid
            delta_state._cycle_wrote = True
        else:
            for bi, block, rows_block, n_rows, runs, prev in pending:
                try:
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
                                merged_rows.append(list(values[i]))
                            elif prev is not None and i < len(prev):
                                merged_rows.append(list(prev[i]))
                            else:
                                merged_rows.append(list(values[i]))
                        sheet.range((5 + lo, block["start"])).value = merged_rows
                    else:
                        sheet.range((5, block["start"])).options(
                            index=False, header=False
                        ).value = [list(r) for r in rows_block]
                    delta_state._cycle_wrote = True
                except Exception as e:
                    log_excel_exception(asset, e, f"writing chain rows block {bi + 1}")
    except Exception as e:
        log_excel_exception(asset, e, f"flushing chain blocks {asset}")
    finally:
        _performance_metrics.record(
            asset,
            "excel_range_write",
            time.perf_counter() - flush_started,
            rows=sum(len(p["rows"]) for p in payloads),
        )


def invalidate_com_cache(asset=None):
    """Drop cached Excel COM objects so the next write re-resolves them.

    Called on any COM error and whenever we notice Excel went away. Without
    this, a stale handle keeps failing forever and presents as the app having
    "stopped writing" while the sheet sits unchanged.
    """
    if asset is None:
        _book_sheet_cache.clear()
        _sheet_cache.clear()
        delta_state._app_cache = None
    else:
        _book_sheet_cache.pop(asset, None)
        _sheet_cache.pop(asset, None)
    # Per-shape bookkeeping is also suspect after a reconnect: the sheet may
    # have been reopened from a different copy.
    _block_payload.clear()
    _usd_fmt_done.clear()
    _dropdown_sig.clear()
    _block_hidden.clear()
    _colour_state.clear()
    _last_rows.clear()


def excel_alive():
    """True when the tool's private Excel instance is reachable."""
    return _owner_alive()


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
        log_excel_exception(asset, e, "writing expiry dropdown source")
        return
    formula = f"=$AH$1:$AH${len(dates)}"
    for cell in ["D3", "O3", "Z3"][:n_blocks(asset)]:
        rng = sheet.range(cell)
        try:
            rng.api.NumberFormat = "@"  # text — stops Excel date conversion
        except Exception as e:
            log_excel_exception(asset, e, f"formatting expiry input {cell}")
        try:
            rng.api.Validation.Delete()
        except Exception:
            # Excel raises when the cell has no existing validation.
            pass
        try:
            rng.api.Validation.Add(Type=3, Formula1=formula)
            rng.api.Validation.ShowError = False  # dropdown = hint, typing still allowed
        except Exception as e:
            print(f"Dropdown error ({asset} {cell}): {e}", flush=True)
            log_excel_exception(asset, e, f"applying expiry dropdown to {cell}")




def _short_build():
    """Just the sha, so the status line stays short enough for a title bar."""
    if delta_state._SHORT_BUILD is None:
        raw = str(build_id())
        for token in raw.split():
            if token.startswith("build_sha="):
                delta_state._SHORT_BUILD = token.split("=", 1)[1]
                break
        else:
            delta_state._SHORT_BUILD = "unstamped"
    return delta_state._SHORT_BUILD


def show_status(book, asset, status, detail=""):
    """Put the tool's health where the user can see it without asking for anything.

    The customer is not technical and does not read logs. Support works by
    screen-sharing, so the one thing guaranteed to be on screen is this
    workbook's own title bar. It carries the build and a plain-language status,
    which is enough to tell a working tool from a broken one and to name the
    build, without the customer having to do anything at all.

    Only written when the text actually changes, so it costs nothing per cycle.
    """
    stamp = time.strftime("%H:%M:%S")
    title = f"{asset}  |  Option Chain {_short_build()}  |  {status}  {stamp}"
    key = id(book)
    if _last_window_status.get(key) == title:
        return
    _last_window_status[key] = title
    try:
        book.api.Windows(1).Caption = title
    except Exception as e:
        log_excel_exception(asset, e, "setting window status")


def excel_environment_detail(book):
    """Excel version and bitness, for diagnosing machine-specific failures.

    A 32-bit Excel driving 64-bit Python is the classic reason this works on one
    Windows machine and not another, so it is recorded on every connect.
    """
    try:
        api = book.app.api
        version = getattr(api, "Version", "unknown")
        os_name = getattr(api, "OperatingSystem", "unknown")
        try:
            import ctypes

            bitness = "64-bit" if ctypes.sizeof(ctypes.c_void_p) == 8 else "32-bit"
        except Exception:
            bitness = "unknown"
        return (
            f"excel_version={version}; excel_operating_system={os_name}; "
            f"excel_process_bitness={bitness}"
        )
    except Exception as e:
        return f"excel_detail_unavailable={type(e).__name__}"




def _find_chain_sheet(book, asset):
    """Return the sheet this product writes to, by NAME not by index.

    Delta's chain sheets are named after the asset (BTC/ETH/XAUT), not "Chain".
    ``book.sheets[0]`` is the sheet Excel reports as first — which can be a
    customer's new sheet if they inserted one, not the asset's owned sheet.
    Writing there overwrites the customer's data with the header band.

    Fall back to sheets[0] only if the named sheet is missing, then let the
    structure check (``sheet_structure_ok``) decide whether the write is safe.
    """
    try:
        return book.sheets[asset]
    except Exception:
        pass
    for sh in book.sheets:
        try:
            if sh.name == asset:
                return sh
        except Exception:
            continue
    return book.sheets[0]


def _write_to_excel(asset):
    try:
        import xlwings as xw
    except ImportError:
        show_error("Required component 'xlwings' is missing. Please use the provided OptionChain.exe.")
        _heal_keyboard()
        os._exit(1)
    fname = os.path.join(BASE_DIR, f"{asset.lower()}_chain.xlsx")
    try:
        book = _book_sheet_cache.get(asset)
        cached_sheet = _sheet_cache.get(asset)
        if book is not None:
            # Liveness probe. Cheap (one COM read) and it is what catches a
            # workbook the user closed underneath us, a renamed sheet or an
            # Excel restart, instead of failing on the next real write.
            try:
                _ = book.name
                if cached_sheet is not None:
                    _ = cached_sheet.name
            except Exception as e:
                log_excel_exception(asset, e, "cached COM object went stale")
                record_excel_connection(asset, "cached_book_unresponsive", str(e))
                book = None
                cached_sheet = None
                _book_sheet_cache.pop(asset, None)
                _sheet_cache.pop(asset, None)
        if book is None:
            book = find_open_book(xw, fname)
            if book is not None:
                _book_sheet_cache[asset] = book
                _sheet_cache.pop(asset, None)
        if book is None:
            record_excel_connection(
                asset,
                "own_excel_running_workbook_not_found"
                if _owner_alive() else "excel_not_found",
                f"expected_workbook={fname}",
            )
            now = time.monotonic()
            if now - delta_state._enum_unknown.get(asset, 0.0) < _UNKNOWN_TTL_SECONDS:
                # Excel did not answer its workbook list this cycle (a dialog
                # is up or the instance is being torn down). The truth is
                # "unknown", never "missing": the session is kept and this
                # cycle skipped. This is what ends the workbook-missing modal
                # that used to refire every couple of minutes - a busy Excel
                # is no longer an error state.
                if excel_connected[asset]:
                    logged_at = _unknown_logged_at.get(asset, 0.0)
                    if now - logged_at > _UNKNOWN_TTL_SECONDS:
                        _unknown_logged_at[asset] = now
                        feed_log(
                            "excel_attach_unknown",
                            asset=asset,
                            note="Excel did not answer this cycle; keeping the session",
                        )
                return
            # A confirmed absence: the customer closed this workbook, and it
            # stays closed. No reopen, no modal - just a log line and a kept
            # hole until the workbook is open again or the tool exits.
            if excel_connected[asset]:
                excel_connected[asset] = False
                _sheet_cache.pop(asset, None)
                feed_log(
                    "workbook_closed_by_customer",
                    asset=asset,
                    note="the chain workbook was closed; it stays closed",
                )
            delta_state._missing_since[asset] = now
            return
        excel_miss_counts[asset] = 0
        delta_state._enum_unknown.pop(asset, None)
        _unknown_logged_at.pop(asset, None)
        delta_state._missing_since.pop(asset, None)
        excel_connected[asset] = True
        show_status(book, asset, "LIVE")
        if _excel_connection_status.get(asset) != "connected":
            record_excel_connection(
                asset, "connected", f"workbook={fname}; {excel_environment_detail(book)}"
            )
        else:
            record_excel_connection(asset, "connected", "")
        sheet = cached_sheet
        if sheet is None:
            sheet = _find_chain_sheet(book, asset)
            _sheet_cache[asset] = sheet

        # Retire the third expiry block the moment the sheet is reachable —
        # not only at the first data write — so the customer never stares at
        # the previous session's Block 3 through the cold-start window.
        _retire_unused_blocks(sheet, asset)

        # The tool writes to fixed columns. If the customer inserts or deletes
        # a column, those positions no longer line up and the tool would write
        # over their own data. A blank cell is fine - the tool is about to
        # fill it. Only a cell holding something that is not our header means
        # the sheet was restructured.
        if not sheet_structure_ok(sheet):
            if not _sheet_altered_notice.get(asset):
                _sheet_altered_notice[asset] = True
                show_error(
                    f"The layout of {os.path.basename(fname)} has changed, so "
                    "the tool will not write to it.\n\n"
                    "The columns it fills are fixed. Undo the change - or "
                    "close the file and reopen the original workbook - and the "
                    "tool will carry on."
                )
            return
        _sheet_altered_notice[asset] = False

        # One merged read for the whole status strip; any per-cell read
        # below falls back through cell_value() only if this failed.
        try:
            pre_rows = sheet.range((1, 1), (4, 31)).value
            pre = pre_rows if isinstance(pre_rows, list) and isinstance(pre_rows[0], list) else None
        except Exception as e:
            log_excel_exception(asset, e, "merged status strip read")
            pre = None

        v = cell_value(pre, sheet, 2, 1)
        if isinstance(v, str):
            try:
                v = float(v.strip())
            except ValueError:
                v = None
        user_provided = isinstance(v, (int, float)) and v >= 1

        # B2 hint — re-write only if missing (user can't permanently lose it)
        try:
            b2 = sheet.range("B2")
            if cell_value(pre, sheet, 2, 2) != "Total strikes — edit A2":
                b2.value = "Total strikes — edit A2"
                b2.font.bold = True
                b2.font.color = (89, 89, 89)
        except Exception as e:
            log_excel_exception(asset, e, "updating worksheet instruction")

        update_user_dates(asset, book, sheet, pre)

        # expiry dropdown (list of DD/MM/YYYY dates) — refresh when list changes
        _sig = tuple(available_dates[asset])
        if _sig and _dropdown_sig.get(asset) != _sig:
            setup_expiry_dropdown(sheet, asset)
            _dropdown_sig[asset] = _sig

        sp = spot_cmp[asset] or 0
        if not sp or spot_is_stale(asset):
            parity = parity_spot_or_none(asset)
            if parity:
                sp = parity

        if not user_provided:
            # A2 empty = show every strike the exchange sends, whatever the
            # count: measured 20-105 rows per expiry in the live logs, so
            # capping by default would only ever trim real strikes off the
            # biggest lists. A2 stays the way to narrow it on purpose.
            range_val = 9999
        else:
            range_val = max(int(v) // 2, 1)  # user limit — A2 kabhi overwrite nahi hota

        nb = n_blocks(asset)
        # Blocks are queued here and written once, after the loop. Writing each
        # block as soon as it is built costs one COM call per block; the flush
        # can usually do the whole sheet in a single call.
        pending_writes = []
        for bi, block in enumerate(BLOCK_COLS):
            if bi >= nb:
                # block not shown for this asset — clear once, leave hidden
                col_l = block["start"]
                if _last_rows.get((asset, bi), 0) or _block_hidden.get((asset, bi)) is None:
                    try:
                        cs = get_column_letter(block["start"])
                        ce = get_column_letter(block["start"] + 8)
                        sheet.range(f"{cs}1:{ce}1000").clear()
                        sheet.range(f"{cs}:{ce}").api.EntireColumn.Hidden = True
                    except Exception as e:
                        log_excel_exception(asset, e, f"clearing unused block {bi + 1}")
                    _last_rows[(asset, bi)] = 0
                    _block_hidden[(asset, bi)] = True
                continue
            if _block_hidden.get((asset, bi)):
                try:
                    cs = get_column_letter(block["start"])
                    ce = get_column_letter(block["start"] + 8)
                    sheet.range(f"{cs}:{ce}").api.EntireColumn.Hidden = False
                except Exception as e:
                    log_excel_exception(asset, e, f"showing active block {bi + 1}")
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
                    except Exception as e:
                        log_excel_exception(asset, e, f"clearing empty block {bi + 1}")
                    _last_rows[(asset, bi)] = 0
                continue

            build_started = time.perf_counter()
            chain = None
            try:
                chain, pcr = build_option_chain(asset, expiry, range_val)
            finally:
                _performance_metrics.record(
                    asset,
                    "chain_build",
                    time.perf_counter() - build_started,
                    rows=len(chain) if chain is not None else 0,
                )
            col_l = block["start"]
            cnt = len(chain) if chain is not None else 0
            # Only rewrite the header when the text actually changed. Writing an
            # identical string still dirties the cell and still makes Excel
            # recalculate everything that depends on it, so writing it every
            # cycle costs the customer a recalculation for nothing.
            spot_text = f"{asset} Spot: {sp:.2f} | {cnt} strikes"
            # Never let a delay go unannounced. A silent drop means the trader
            # is reading old data, and that is exactly what "the data is
            # stale" means to them. The note is empty in the normal case, so
            # the header text does not change and the cell is not rewritten.
            _note = source_delay_note(asset)
            if _note:
                spot_text = f"{spot_text} | {_note}"
            pcr_text = f"PCR: {pcr:.2f}"
            try:
                if cell_value(pre, sheet, 1, col_l) != spot_text:
                    sheet.range((1, col_l)).value = spot_text
                    delta_state._cycle_wrote = True
                if cell_value(pre, sheet, 1, col_l + 4) != pcr_text:
                    sheet.range((1, col_l + 4)).value = pcr_text
                    delta_state._cycle_wrote = True
            except Exception as e:
                log_excel_exception(asset, e, "writing chain header")
            net_paint(sheet, asset, bi, col_l)

            if chain is not None and not chain.empty:
                write_started = time.perf_counter()
                try:
                    # Only write when the payload actually differs. The block
                    # range write is the single most expensive call in the whole
                    # render loop, and between ticks most rows are unchanged,
                    # so a cheap Python-side signature check avoids a full Excel
                    # rewrite for no visible difference.
                    values = chain.values.tolist()
                    # The block always occupies the same rows. Live strikes
                    # fill from the top; rows past the live count stay blank,
                    # so a customer range written for a 37-strike expiry
                    # (P5:P41) keeps covering the whole chain when the next
                    # expiry runs longer.
                    # The extent never shrinks: the highest rows this session
                    # has seen for the block stay the floor, so a shorter
                    # expiry cannot yank rows out from under the customer's
                    # offset ranges.
                    seen = delta_state._row_extent_seen.get((asset, bi), 0)
                    extent = max(delta_state.ROW_EXTENT_DEFAULT, seen, len(values))
                    if extent != seen:
                        if seen:
                            feed_log(
                                "chain_extent_changed",
                                asset=asset,
                                block=bi + 1,
                                from_rows=seen,
                                to_rows=extent,
                            )
                        delta_state._row_extent_seen[(asset, bi)] = extent
                    values = pad_block_rows(values, extent)
                    # Sizes downstream (stale clear, formulas ranges, colour)
                    # all key off the extent so the geometry stays constant.
                    n_rows = extent
                    payload_sig = hash(tuple(map(tuple, values)))
                    if payload_sig != _block_payload.get((asset, bi)):
                        prev = _block_values.get((asset, bi))
                        changed = None
                        if prev is not None and len(prev) == len(values):
                            changed = [i for i in range(len(values)) if list(prev[i]) != list(values[i])]
                        runs = _contiguous_runs(changed) if changed else []
                        pending_writes.append(
                            (bi, block, values, n_rows, runs, prev)
                        )
                        _block_payload[(asset, bi)] = payload_sig
                        _block_values[(asset, bi)] = values
                    else:
                        _performance_metrics.record(asset, "excel_range_skipped", 0.0, rows=n_rows)
                finally:
                    _performance_metrics.record(
                        asset,
                        "excel_block_prep",
                        time.perf_counter() - write_started,
                        rows=cnt,
                    )

                strikes = chain["strike"].tolist()
                n = len(strikes)
                cs = get_column_letter(block["start"])
                ce = get_column_letter(block["start"] + 8)

                # Clear whatever sits below the live table. Two cases matter:
                #   - the block is below its recorded extent (only possible on
                #     the first cycle this block is seen, since the extent only
                #     ever grows), and
                #   - this is the first time we have seen this block, which
                #     happens on every tool restart because _last_rows is
                #     in-memory only. Without the second case a table left over
                #     from an earlier run survived forever, because prev_n
                #     was 0 and "n_rows < prev_n" was never true.
                prev_n = _last_rows.get((asset, bi))
                if prev_n is None:
                    first_sight = _FIRST_SIGHT_BOUND
                else:
                    first_sight = 0
                if prev_n is None or n_rows < prev_n:
                    try:
                        last_row = max(n_rows, prev_n or 0, first_sight)
                        if last_row < n_rows:
                            last_row = n_rows
                        if last_row > n_rows:
                            # Our own block columns only. The customer may keep
                            # their own formulas in these rows further right.
                            # A reversed range (rows n..n-1) resolves to its
                            # top-left and clips the wrong row, so only clear
                            # when there is something below the live rows.
                            sheet.range(f"{cs}{5 + n_rows}:{ce}{4 + last_row}").clear()
                    except Exception as e:
                        log_excel_exception(asset, e, f"clearing stale rows in block {bi + 1}")
                _last_rows[(asset, bi)] = n_rows
                _deep_clear_below_block(sheet, asset, bi, block["start"], n_rows)

                formatting_started = time.perf_counter()
                try:
                    # force USD-only formats (kills currency rendering). This one
                    # is self-gating, so it only does work when it must.
                    force_usd_block_formats(sheet, block, asset, bi, cnt)

                    # Recolour the whole block, but only when the picture
                    # would actually change.
                    #
                    # The old gate was a "shape" signature of (row count,
                    # first strike, last strike). That was wrong: the ATM row
                    # and the ITM/OTM split all derive from spot, and spot
                    # was not in the signature. So when price moved without
                    # the strike window changing size, nothing was repainted
                    # and stale yellow and red rows were left behind.
                    #
                    # The signature below carries everything the colouring
                    # depends on: the row count, where spot splits the calls
                    # from the puts, and which row is at the money. An
                    # unchanged picture is not repainted, which is four Excel
                    # calls per block saved on every cycle where the price
                    # has not crossed a strike.
                    k = bisect.bisect_left(strikes, sp)
                    m = bisect.bisect_right(strikes, sp)
                    atm_index = (
                        min(range(len(strikes)), key=lambda i: abs(strikes[i] - sp))
                        if strikes
                        else -1
                    )
                    colour_sig = (n, k, m, atm_index)
                    if _colour_state.get((asset, bi)) != colour_sig:
                        _colour_state[(asset, bi)] = colour_sig
                        # Clearing first is what makes it correct: setting
                        # .color only paints, so a row that leaves the ITM
                        # region would otherwise keep its old colour forever.
                        sheet.range(f"{cs}5:{ce}{4 + n}").color = None
                        if k > 0:
                            sheet.range(f"{block['call_start']}5:{block['call_end']}{4 + k}").color = (198, 224, 180)
                        if m < n:
                            sheet.range(f"{block['put_start']}{5 + m}:{block['put_end']}{4 + n}").color = (255, 200, 200)
                        if n:
                            atm_row = atm_index + 5
                            atm_range = f"{block['atm']}{atm_row}:{block['put_end']}{atm_row}"
                            sheet.range(atm_range).color = (255, 255, 0)
                finally:
                    _performance_metrics.record(
                        asset,
                        "excel_formatting",
                        time.perf_counter() - formatting_started,
                        rows=cnt,
                    )
            else:
                # An empty block only needs clearing when it stops being empty,
                # or on the first sighting. Clearing 1000 rows every cycle was
                # both wasteful and needlessly expensive for Excel to do.
                prev_n = _last_rows.get((asset, bi))
                if prev_n is None or prev_n != 0:
                    clear_start = get_column_letter(block["start"])
                    clear_end = get_column_letter(block["start"] + 8)
                    last_row = max(prev_n or 0, _FIRST_SIGHT_BOUND)
                    try:
                        sheet.range(f"{clear_start}5:{clear_end}{4 + last_row}").clear()
                    except Exception as e:
                        log_excel_exception(asset, e, f"clearing empty chain block {bi + 1}")
                _last_rows[(asset, bi)] = 0

        # One flush for the whole sheet. Everything the block loop queued goes
        # out now, in as few Excel calls as the layout allows.
        _flush_block_writes(sheet, asset, pending_writes)

        excel_fail_since[asset] = None
        excel_fail_notice_shown[asset] = False

    except Exception as e:
        print(f"Excel error ({asset}): {e}", flush=True)
        record_excel_connection(asset, "excel_com_error", f"{type(e).__name__}: {e}")
        log_excel_exception(asset, e)
        # Any COM failure may mean the workbook was closed, renamed or Excel was
        # restarted. Drop every cached object so the next cycle re-resolves from
        # scratch rather than writing into a dead COM handle forever.
        invalidate_com_cache(asset)
        now = time.time()
        if excel_fail_since[asset] is None:
            excel_fail_since[asset] = now
        elif now - excel_fail_since[asset] > 30 and not excel_fail_notice_shown[asset]:
            excel_fail_notice_shown[asset] = True
            show_error(
                "Lost connection to Microsoft Excel.\n\n"
                "The tool is still running and will keep trying on its own.\n"
                "To resume now: close Excel, reopen the chain workbooks,\n"
                "and leave Excel open for a few seconds.\n\n"
                "Details: " + str(e)
            )



# xlwings' .calculation getter hands back a stale/garbage value while Excel is
# busy, so a "did the restore work?" check built on it can pass falsely and leave
# the trader stuck on manual. Set AND verify through the raw API with these ids.

# Calculation-mode churn tracking: any observed transition between Excel's
# raw calculation id from one render cycle to the next is counted, whoever
# caused it (this tool, the sibling chain product, or the customer). On a
# healthy session the counter stays flat; the moment the two tools start
# fighting over the mode it climbs every few seconds, which the summary log
# line makes visible without any profiling.



def _observe_raw_calc(raw):
    """Count a transition of Excel's calculation mode between cycles.

    raw comes from the read `_begin_excel_bulk_mode` already does every
    cycle, so this adds no COM round-trips. An unreadable value (None)
    resets the baseline so a single glitch is not mistaken for a flip.
    """
    if raw is None:
        delta_state._last_observed_calc_raw = None
        return
    previous = delta_state._last_observed_calc_raw
    delta_state._last_observed_calc_raw = raw
    if previous is None or previous == raw:
        return
    delta_state._calc_churn_count += 1
    now = time.time()
    if now >= delta_state._calc_churn_summary_after:
        delta_state._calc_churn_summary_after = now + _CALC_CHURN_SUMMARY_SECONDS
        feed_log(
            "calc_switch_summary",
            churn_total=delta_state._calc_churn_count,
            last_from=_CALC_ID_NAMES.get(previous, str(previous)),
            last_to=_CALC_ID_NAMES.get(raw, str(raw)),
            note="cumulative calculation-mode transitions seen between cycles",
        )


def _other_chain_tool_holds():
    """True when the current manual window belongs to a still-running tool.

    The 4s watchdog must never rescue a mode a live tool owns: the sibling
    chain product deliberately holds manual between its sheets, and flipping
    to automatic underneath it made Excel flap every few seconds whenever
    both tools ran together.
    """
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
    if now - delta_state._calc_watchdog_skip_logged_at >= 60.0:
        delta_state._calc_watchdog_skip_logged_at = now
        feed_log(
            "calc_watchdog_skip_holder",
            note="another chain tool holds manual calculation; watchdog left it alone",
        )
    return True


def _raw_calc_mode(app):
    return int(app.api.Calculation)


def _set_calc_mode(app, name):
    app.api.Calculation = _CALC_IDS[name]
    return _raw_calc_mode(app) == _CALC_IDS[name]


def ensure_calculation_not_stranded(app=None):
    """Safety net: a trader must never be left in manual calculation.

    If Excel is still on manual well after any write window should have closed,
    put it back. Without this, one failed restore leaves every formula the
    trader types silently uncalculated, with no visible cause.
    """
    if app is None:
        app = delta_state._app_cache
    if app is None:
        app = delta_state._owner_app
    if app is None:
        return
    try:
        mode = _raw_calc_mode(app)
    except Exception:
        return
    now = time.time()
    if mode != -4135:                       # not manual
        delta_state._manual_since = 0.0
        return
    if delta_state._manual_since == 0.0:
        delta_state._manual_since = now
        return
    if now - delta_state._manual_since < MANUAL_WATCHDOG_SECONDS:
        return
    if _other_chain_tool_holds():
        return
    try:
        if _set_calc_mode(app, "automatic"):
            delta_state._manual_since = 0.0
            feed_log("calc_watchdog_restored",
                     note="Excel was stranded on manual calculation")
    except Exception as e:
        feed_log("calc_watchdog_error", error=type(e).__name__, detail=str(e)[:150])


# The chain workbooks this tool writes. Closing any of them is the one
# refusal that must never become a customer-facing dialog.
_CHAIN_BOOK_NAMES = frozenset(
    f"{asset.lower()}_chain.xlsx" for asset in ("BTC", "ETH", "XAUT")
)

_CALC_NOTICE_GRACE_SECONDS = 60.0
_DELTA_CANNOT_TELL_TIMEOUT = 15.0


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
    return not any(name in open_names for name in _CHAIN_BOOK_NAMES)


_SAVE_PROMPT_CLASSES = ("NUIDialog", "#32770")
_SAVE_PROMPT_POLL_SECONDS = 0.25
_SAVE_PROMPT_ATTEMPT_LIMIT = 8

# The chain frames are named after their books, so a prompt over one of them
# is recognised two ways: the rendered title wording ("BTC | Options Chain")
# and the file-backed caption ("btc_chain - Excel") once underscores and
# dashes are read as spaces.
_DELTA_CHAIN_CAPTION_TOKENS = (
    "option chain",
    "options chain",
    "btc chain",
    "eth chain",
    "xaut chain",
)


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
    caption = buf.value.replace("_", " ").replace("-", " ").lower()
    return any(token in caption for token in _DELTA_CHAIN_CAPTION_TOKENS)


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


def _screen_update_watchdog_tick():
    """Finish a failed screen restore once Excel starts answering again.

    The restore inside the window retries three times and then hands the
    job to this tick ("Excel was busy" is a moment, not a state). It runs
    once per display cycle while no bulk window is open, and it only ever
    touches our own private instance's paint state.
    """
    app = delta_state._app_cache or delta_state._owner_app
    if app is None:
        return
    if delta_state._bulk_window_open:
        return
    try:
        busy = app.screen_updating
    except Exception:
        return
    if busy:
        delta_state._screen_restore_owed = False
        return
    try:
        app.screen_updating = True
    except Exception as e:
        log_excel_exception("ALL", e, "screen watchdog force-on")
        return
    delta_state._screen_restore_owed = False
    now = time.monotonic()
    if now - delta_state._screen_watchdog_logged_at >= 60.0:
        delta_state._screen_watchdog_logged_at = now
        feed_log(
            "screen_updating_forced_on",
            note="watchdog restored paint after a failed in-window restore",
        )


_DEAD_COM_CODES = {
    0x800A01A8,  # OLE error 0x800A01A8 - the app object is gone
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE - the Excel process is gone
    0x800706B5,  # RPC_S_CALL_FAILED
    0x80010108,  # RPC_E_DISCONNECTED
}


def _com_dead(exc):
    """True when the failure says the Excel instance itself is gone.

    A switch refused while a modal is up stays a recoverable refusal. Excel
    having closed records nothing: the marks would otherwise arm the
    slow-down notice a minute after the customer closed the workbook.
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


def _release_dead_excel_state():
    """After Excel is gone: drop the session, marks and the stale handle."""
    delta_state._bulk_owning_manual = False
    delta_state._calc_manual_failed = False
    delta_state._calc_failed_since = 0.0
    try:
        calc_mode_guard.release(PRODUCT_ID)
    except Exception:
        pass
    delta_state._app_cache = None


def _begin_excel_bulk_mode():
    """Take the manual-calculation session for one render cycle.

    ScreenUpdating is deliberately never read or written here. Toggling it
    around the write window made Excel repaint the entire sheet on every
    release - the sheet visibly flashed - and any click into a cell that
    landed while the screen was off left the rest of that edit frozen behind
    a stale frame, so the formula the customer was typing never appeared.
    With the screen left on, Excel repaints only the ranges we actually
    write, and their typing, dragging and editing behave normally between
    cycles. The costly half is the recalculation, and the calculation mode
    below is what owns it.

    Every Excel write otherwise triggers a full recalculation. That is cheap
    in an empty sheet but expensive when the user has a strategy workbook
    open, which is where refresh times blow up.

    Returns a restore callable, or None if nothing could be changed. Failures
    are expected and non-fatal: the calculation switch raises if the user is
    mid-edit, and there is no workbook to talk to yet.
    """
    # The watchdog needs to know when NOT to interfere; the window is
    # marked open only when a restore callable actually comes back.
    delta_state._bulk_window_open = False
    state = {"app": None, "calc": None, "calc_guarded": False}
    try:
        app = delta_state._app_cache
        if app is None:
            app = delta_state._owner_app
        if app is None:
            return None
        try:
            # xlwings.App has no .name; .version is the cheap liveness probe
            # that goes through to Excel and raises if the instance is gone.
            _ = app.version
        except Exception as e:
            log_excel_exception("ALL", e, "cached Excel app went stale")
            # A dead instance ends the whole session: marks and the stale
            # handle go together, or the refused-manual latch would survive
            # the close and pop its notice a minute later.
            _release_dead_excel_state()
            return None
        state["app"] = app
        # Closed-workbook gate. When the customer is closing the chains, Excel
        # is either gone or busy answering its save prompt, and both states
        # refuse COM calls. There is nothing of ours to write, so the honest
        # response is to skip the window - which is also why the refused
        # manual switch that used to pop a dialog on every workbook close no
        # longer can: the switch is never attempted while the chains are gone.
        open_names = _open_book_names(app)
        if _chain_books_missing(open_names):
            delta_state._bulk_owning_manual = False
            delta_state._calc_manual_failed = False
            delta_state._calc_failed_since = 0.0
            delta_state._app_cache = app
            now = time.monotonic()
            if now - delta_state._no_book_skip_logged_at >= 60.0:
                delta_state._no_book_skip_logged_at = now
                feed_log(
                    "bulk_mode_skipped_no_chain_workbook",
                    note="no chain workbook is open; skipping the write window",
                )
            return None
        # "Cannot tell" state: open_names is None (COM refused), but Excel
        # process is still alive. This happens during save-prompt teardown on
        # workbook close. If this state persists beyond the timeout, the
        # workbook is gone and any refusal marks must be cleared so the
        # notice never fires.
        if open_names is None and delta_state._calc_manual_failed:
            try:
                gone = not excel_process_running()
            except Exception:
                gone = False
            if gone or (
                delta_state._calc_failed_since
                and time.monotonic() - delta_state._calc_failed_since >= _DELTA_CANNOT_TELL_TIMEOUT
            ):
                delta_state._bulk_owning_manual = False
                delta_state._calc_manual_failed = False
                delta_state._calc_failed_since = 0.0
                delta_state._calc_notice_shown = False
        # Deliberate no-op where the ScreenUpdating toggle used to live. The
        # screen is now left exactly as Excel itself keeps it: flipping it
        # around the write window repainted the whole sheet on every release
        # (the flicker count) and froze any edit that started mid-window - the
        # customer typed a formula and saw nothing until they pressed Enter.
        # The first COM call below tells us instantly if Excel can talk.
        try:
            # xlwings exposes calculation as a lowercase NAME
            # ("automatic" / "semiautomatic" / "manual"); the setter only
            # accepts those names.
            #
            # Read it through the raw API on purpose. xlwings' own getter does
            # calculation_i2s[self.xl.Calculation], and when Excel is busy or
            # showing a modal it hands back an HRESULT instead of a valid
            # xlCalculation value, so the lookup raises KeyError.
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
                # Option Chain tool set it for its own write window. Either way,
                # leave it alone and do not record it for restore -- recording it
                # would "restore" the other tool's temporary setting and leave
                # the customer's Excel stuck on manual calculation.
                state["calc"] = None
            else:
                # Session-level calc retry cooldown: if the last attempt to
                # switch to manual failed, skip until the cooldown expires so
                # we do not hammer COM and make the customer's workbook crawl.
                if delta_state._calc_manual_failed:
                    now = time.monotonic()
                    if now - delta_state._calc_failed_at < delta_state.CALC_RETRY_COOLDOWN_SECONDS:
                        state["calc"] = None
                    else:
                        # Cooldown expired - retry once to recover.
                        delta_state._calc_manual_failed = False
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
                if state["calc"]:
                    calc_mode_guard.acquire(state["calc"], PRODUCT_ID)
                    state["calc_guarded"] = True
                    # If the workbook was already on automatic, Excel will
                    # recalculate by itself once we restore. Semiautomatic does
                    # the same for ordinary formulas.
                    delta_state._bulk_will_recalc = state["calc"] in ("automatic", "semiautomatic")
                    try:
                        app.calculation = "manual"
                        delta_state._bulk_owning_manual = True
                        delta_state._calc_failed_since = 0.0
                    except Exception as exc:
                        was_guarded = state.get("calc_guarded", False)
                        state["calc"] = None
                        state["calc_guarded"] = False
                        if _com_dead(exc) or not _excel_alive_after(exc):
                            # Excel died mid-window (the customer closed it).
                            # Not a refusal: drop the session so the reopen
                            # path can bring Excel back without the notice.
                            _release_dead_excel_state()
                            feed_log(
                                "calc_switch_excel_gone",
                                detail=type(exc).__name__,
                            )
                        else:
                            was_guarded = state.get("calc_guarded", False)
                            if not delta_state._calc_failed_since:
                                delta_state._calc_failed_since = time.monotonic()
                            delta_state._calc_failed_at = time.monotonic()
                            delta_state._calc_manual_failed = True
                            feed_log(
                                "calc_manual_unavailable",
                                error=type(exc).__name__,
                                detail=str(exc)[:120],
                            )
                            if was_guarded:
                                try:
                                    calc_mode_guard.release(PRODUCT_ID)
                                except Exception:
                                    pass
                            # A workbook close refuses the switch for as long as
                            # its save prompt is up, and that is not a fault a
                            # customer needs explained to their face. Only a
                            # refusal that has held longer than the grace period
                            # - and the gate above only lets us attempt the
                            # switch while the chains are open - earns the
                            # dialog. Excel being gone never does: the marks it
                            # would read were just cleared.
                            if (
                                not delta_state._calc_notice_shown
                                and time.monotonic() - delta_state._calc_failed_since
                                >= _CALC_NOTICE_GRACE_SECONDS
                            ):
                                delta_state._calc_notice_shown = True
                                show_error(
                                    "Excel will not switch to manual calculation,\n"
                                    "so the workbook recalculates on every update\n"
                                    "and will feel slow.\n\n"
                                    "This is usually one of three things:\n"
                                    "  - a cell is in edit mode - press Escape\n"
                                    "  - an Excel dialog is open - close it\n"
                                    "  - an Excel add-in is busy - turn it off\n\n"
                                    "The tool will speed back up on its own once\n"
                                    "Excel accepts the mode switch again.",
                                )
        except Exception as e:
            log_excel_exception("ALL", e, "switching to manual calculation")
            state["calc"] = None
    except Exception as e:
        log_excel_exception("ALL", e, "entering bulk mode")
        return None

    def restore():
        delta_state._bulk_owning_manual = False
        delta_state._bulk_window_open = False
        app = state.get("app")
        if app is None:
            return
        if state.get("calc") is not None:
            # Verify the restore actually took. A single setter call can fail
            # while Excel is busy, and leaving the trader in manual calculation
            # means every formula they type silently refuses to compute.
            want = state["calc"]
            ok = False
            for attempt in range(3):
                try:
                    ok = _set_calc_mode(app, want)
                    if ok:
                        break
                    feed_log("calc_restore_retry", attempt=attempt + 1,
                             want=want, got=_raw_calc_mode(app))
                except Exception as e:
                    log_excel_exception("ALL", e, "restoring calculation mode")
                    feed_log("calc_restore_error", attempt=attempt + 1,
                             error=type(e).__name__, detail=str(e)[:150])
                time.sleep(0.5)
            if not ok:
                try:
                    feed_log("calc_restore_FAILED", want=want, got=_raw_calc_mode(app),
                             note="trader left in manual calculation")
                except Exception:
                    pass
        if state.get("calc_guarded"):
            # Only drop our claim once Excel has actually been put back.
            try:
                calc_mode_guard.release(PRODUCT_ID)
            except Exception as e:
                log_excel_exception("ALL", e, "releasing calculation guard")
        # No paint action here either: the screen was never taken, so it is
        # already exactly where Excel and the customer want it. Every release
        # used to force a full-sheet repaint here - that was the flicker - and
        # an edit that began mid-window froze behind the stale frame.

        # The module-level slot, not a local of this closure. Assigned from
        # inside a nested function without going through the state module, it
        # would create a local and the exit path would never find it, so
        # Excel could be left in manual calculation after a crash.
        delta_state._last_bulk_restore = restore

    delta_state._bulk_window_open = True
    return restore


def write_to_excel(asset):
    started = time.perf_counter()
    try:
        _write_to_excel(asset)
    finally:
        _performance_metrics.record(
            asset, "excel_pass", time.perf_counter() - started
        )


def _time_recalc(wrote):
    """Recalculate once, inside the manual window, and time it.

    This is the number nobody had: how long the customer's own workbook takes
    to recalculate when we change a cell it depends on. Every timing we had
    before stopped at the edge of the suspended-calculation window, so the very
    cost that makes their machine feel stuck was never measured.

    Plain Application.Calculate, not CalculateUntilAsyncQueriesDone: that one
    is for asynchronous queries such as Power Query and data connections, not
    for ordinary formulas.

    Only done when we actually wrote something, and only when the workbook was
    already set to recalculate by itself. Recalculating for someone who chose
    manual would be doing something they asked us not to. A cycle skipped for
    customer input owes its recalculation, so _recalc_owed carries it here.
    """
    owed = delta_state._recalc_owed
    delta_state._recalc_owed = False
    if not wrote and not owed:
        _performance_metrics.record("ALL", "excel_recalc_skip_nowrite", 0.0)
        return 0.0
    if not delta_state._bulk_will_recalc:
        _performance_metrics.record("ALL", "excel_recalc_skip_mode", 0.0)
        return 0.0
    app = delta_state._app_cache
    if app is None:
        _performance_metrics.record("ALL", "excel_recalc_skip_noapp", 0.0)
        if owed:
            # Still owed - Excel comes back and the debt is paid then.
            delta_state._recalc_owed = True
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
            delta_state._recalc_owed = True
        return 0.0
    delta_state._last_recalc_ms = elapsed
    _performance_metrics.record("ALL", "excel_recalc", elapsed / 1000.0)
    return elapsed


def _log_speed_summary(cycle_ms):
    """Write the whole pipeline's timings into the log every half minute.

    Every future complaint should come with numbers, so one line carries the
    whole picture: how long a cycle took, how long the customer's own workbook
    took to recalculate, how fresh each spot quote is, and how many strikes
    are being tracked.
    """
    now = time.monotonic()
    if now - delta_state._speed_log_at < 30.0:
        return
    delta_state._speed_log_at = now
    ages = {}
    for asset in ("BTC", "ETH", "XAUT"):
        try:
            age = spot_age(asset)
        except Exception:
            age = None
        ages[asset] = round(age, 1) if age is not None else None
    try:
        strikes = sum(_last_rows.values())
    except Exception:
        strikes = 0
    feed_log(
        "speed",
        cycle_ms=round(cycle_ms * 1000.0, 1),
        recalc_ms=round(delta_state._last_recalc_ms, 1) if delta_state._last_recalc_ms else 0,
        strikes=strikes,
        spot_age_btc_s=ages.get("BTC"),
        spot_age_eth_s=ages.get("ETH"),
        spot_age_xaut_s=ages.get("XAUT"),
    )


def _probe_calc_state_after_restore():
    """One CalculationState COM read per half minute, instead of per cycle.

    The probe only asks whether Excel still owed a recalculation after the
    restore, which is a slow-changing signal; reading it every cycle added
    a round-trip per pass for no extra insight.
    """
    if delta_state._app_cache is None:
        return
    now = time.monotonic()
    cached = delta_state._calc_state_probe
    if cached is not None and now - cached[0] < 30.0:
        _performance_metrics.record("ALL", "excel_recalc_second_pass", cached[1])
        return
    try:
        state = int(delta_state._app_cache.api.CalculationState)
    except Exception:
        return
    hit = 1.0 if state != 0 else 0.0
    delta_state._calc_state_probe = (now, hit)
    _performance_metrics.record("ALL", "excel_recalc_second_pass", hit)
    if hit:
        _note_recalc_second_pass(now)


def _note_recalc_second_pass(now):
    """Record that Excel still owed a recalc after the write window.

    The customer's own formula workbook is heavier than the configured refresh
    assumed. Stretch the loop's sleep so Excel gets idle time, and let it decay
    back once Excel comes out of every write window clean.
    """
    widened = min(
        delta_state.WIDEN_MAX, delta_state._refresh_widen * delta_state.WIDEN_STEP
    )
    if widened != delta_state._refresh_widen:
        delta_state._refresh_widen = widened
        feed_log("refresh_widened", widen=round(widened, 2))
    delta_state._refresh_widen_hold_until = now + delta_state.WIDEN_DECAY_SECONDS


# Any keystroke or click inside this window counts as "hands on the keys".
_USER_INPUT_GRACE_SECONDS = 0.7


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

    hwnd = delta_state._owner_hwnd
    try:
        if hwnd:
            pid = wintypes.DWORD(0)
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value:
                return pid.value
    except Exception:
        pass
    app = delta_state._owner_app
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


def render_cycle():
    """One full pass over every workbook, inside a single bulk-mode window.

    Screen updating and calculation mode are suspended once per cycle rather
    than once per asset, so a cycle costs three writes with repaint and
    recalculation suspended, then one restore.
    """
    started = time.perf_counter()
    # TEST SCAFFOLD - not part of the product. Set OC_TEST_PAUSE=1 to keep the
    # process alive and attached to Excel but skip every write, so a test can
    # separate "the tool is running" from "the tool is writing". No cell, no UI.
    if os.environ.get("OC_TEST_PAUSE") == "1":
        _performance_metrics.record("ALL", "render_cycle", time.perf_counter() - started)
        return



    delta_state._cycle_wrote = False

    # Input yield: never open a write window while the customer's hands are
    # on our Excel. Their keystrokes land on a repaint-numbed frame - the
    # formula they type shows nothing, backspace appears to do nothing -
    # and the window's own COM calls race their typing. Skip the cycle as
    # a whole; the forced recalculation it owed is carried to the next
    # one, so their strategy formulas never sit stale behind the skip.
    if _user_is_excel_interacting():
        now_mono_skip = time.monotonic()
        if now_mono_skip - delta_state._user_input_skip_log_at > 30.0:
            delta_state._user_input_skip_log_at = now_mono_skip
            feed_log(
                "render_cycle_skipped_user_input",
                note="customer is typing or selecting in our Excel",
            )
        delta_state._recalc_owed = True
        _performance_metrics.record("ALL", "render_cycle_skipped_user_input", 0.0)
        return

    # Detect Excel edit mode: skip the write cycle when the customer is
    # dragging a formula, filling cells, or editing a value. Writing to the
    # chain sheet at this moment deadlocks Excel's COM apartment. There is
    # no timeout that forces a window through while the edit is still live:
    # opening one freezes the frame the formula is being typed into, and
    # every write inside it is refused anyway.
    app = delta_state._app_cache
    if app is not None:
        edit_detected = False
        try:
            if app.api.CutCopyMode or not app.api.Ready:
                edit_detected = True
        except Exception:
            edit_detected = True
        if edit_detected:
            now_mono_edit = time.monotonic()
            if now_mono_edit - delta_state._user_edit_skip_log_at > 30.0:
                delta_state._user_edit_skip_log_at = now_mono_edit
                feed_log(
                    "render_cycle_skipped_user_edit",
                    note="Excel is in edit mode; the write window stays shut",
                )
            _performance_metrics.record("ALL", "render_cycle_skipped_user_edit", 0.0)
            return

    restore = _begin_excel_bulk_mode()
    try:
        for asset in ("BTC", "ETH", "XAUT"):
            if _user_is_excel_interacting():
                # The customer's hands reached our Excel while this window was
                # open. Drop what is left of the window and let the finally
                # below hand the screen back; their edit runs on a live frame.
                delta_state._recalc_owed = True
                _performance_metrics.record(
                    "ALL", "render_cycle_aborted_user_input", 0.0
                )
                break
            write_to_excel(asset)
        _time_recalc(delta_state._cycle_wrote)
    finally:
        if restore is not None:
            try:
                restore()
                # If we already recalculated inside the window, restoring
                # automatic should find nothing left to do. If Excel is still
                # calculating or pending, restoring started a second,
                # duplicated pass and the cost is being paid twice.
                try:
                    _probe_calc_state_after_restore()
                except Exception:
                    pass
            except Exception as e:
                log_excel_exception("ALL", e, "leaving bulk mode")
        elapsed = time.perf_counter() - started
        _performance_metrics.record("ALL", "render_cycle", elapsed)
        try:
            _log_speed_summary(elapsed)
        except Exception:
            pass
        # Never leave the trader in manual calculation between write windows.
        # (kept: the watchdog stays exactly once per cycle, in render_cycle)
        try:
            ensure_calculation_not_stranded()
        except Exception:
            pass


async def flush_performance_metrics():
    now = time.monotonic()
    if now - delta_state._last_performance_flush < 30:
        return
    lines = _performance_metrics.summary_lines(now)
    try:
        await asyncio.to_thread(append_performance_summary, lines)
    except OSError as e:
        if not delta_state._performance_log_error_shown:
            show_error(
                "Could not append performance diagnostics to startup_diag.log.\n\n"
                "Please check that your user profile has free space and is writable.\n\n"
                f"Details: {e}"
            )
            delta_state._performance_log_error_shown = True
    else:
        _performance_metrics.reset(now)
        delta_state._last_performance_flush = now


async def display_loop():
    refresh = cfg.get("refresh_interval_seconds", 0.5)
    is_linux = platform.system() != "Windows"
    license_tick = 0
    delta_state._excel_gone_logged = False
    if not is_linux:
        # The chain save prompt would otherwise sit unanswered in the
        # background blocking COM until the customer notices it.
        start_chain_save_prompt_watchdog()
    while True:
        render_started = time.perf_counter()
        # re-check license roughly every ~60s while running
        license_tick += 1
        if license_tick >= max(1, int(60 / max(float(refresh), 0.05))):
            license_tick = 0
            ok, _vt = license_is_valid()
            if not ok:
                if not ensure_license():
                    _heal_keyboard()
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
            _pump_com()
            # A modifier stuck from a crash - or from anything else on the
            # machine - is healed here, once per cycle, without ever
            # touching a key a human could legitimately still be holding.
            _watch_stuck_modifiers_tick()
            # The chain lives in its private Excel instance, so the write pass
            # runs while that instance is alive - whichever other Excel
            # windows exist on the machine are none of this tool's business.
            # Keep the cached hwnd current so _owner_window_pid (and any handle-
            # based check) does not act on a stale handle after workbooks open.
            if delta_state._owner_app is not None:
                try:
                    _h = delta_state._owner_app.api.Hwnd
                    if _h:
                        delta_state._owner_hwnd = _h
                except Exception:
                    pass
            excel_alive = _owner_alive() or delta_state._owner_app is None
            retry_open_pass_after_grace()
            if excel_alive:
                if delta_state._excel_gone_logged:
                    # Excel came back with freshly opened, empty workbooks, so
                    # every cached "already written" claim is now wrong. Drop
                    # them or the app would leave the sheets blank.
                    delta_state._excel_gone_logged = False
                    _block_payload.clear()
                    _usd_fmt_done.clear()
                    _dropdown_sig.clear()
                    _colour_state.clear()
                    _book_sheet_cache.clear()
                    _sheet_cache.clear()
                    _last_excel_exception.clear()
                    print("Excel reopened - rewriting all sheets from scratch.", flush=True)
                # The window is closed at this point, so anything the last
                # restore could not finish gets finished here.
                _screen_update_watchdog_tick()
                render_cycle()
            elif not delta_state._excel_gone_logged:
                delta_state._excel_gone_logged = True
                print(
                    "The chain Excel was closed. The tool stops on its own "
                    "shortly - run it again when you want the chain back.",
                    flush=True,
                )
                _last_excel_exception.clear()
                # Excel's address space is gone, so any refused-switch spell
                # belongs to an instance that no longer exists. Clear it so a
                # fresh Excel starts calm instead of carrying a stale refusal.
                delta_state._calc_manual_failed = False
                delta_state._calc_failed_since = 0.0
                delta_state._app_cache = None
                delta_state._owner_app = None
                delta_state._owner_hwnd = None
                _book_sheet_cache.clear()
                _sheet_cache.clear()
            maybe_exit_after_close_all(excel_alive=excel_alive)
            render_seconds = time.perf_counter() - render_started
            if render_seconds > float(refresh):
                _performance_metrics.record(
                    "ALL", "refresh_overrun", render_seconds - float(refresh)
                )
            await flush_performance_metrics()

        # Sleep AFTER the work, for at least as long as the work took, so a
        # heavy workbook open alongside the chain always leaves Excel idle for
        # an equal stretch and stays usable for building a strategy.
        #
        # The stretch is adaptive: a CalculationState probe that still shows
        # work owed after the write window (see _note_recalc_second_pass)
        # multiplies the configured interval, and the stretch decays back to
        # 1.0 after a quiet spell so a fast workbook is never left crawling.
        now_mono = time.monotonic()
        if (
            now_mono > delta_state._refresh_widen_hold_until
            and delta_state._refresh_widen > 1.0
        ):
            delta_state._refresh_widen = max(1.0, delta_state._refresh_widen / 2.0)
            delta_state._refresh_widen_hold_until = (
                now_mono + delta_state.WIDEN_DECAY_SECONDS
            )
            if delta_state._refresh_widen > 1.0:
                feed_log(
                    "refresh_widen_decayed", widen=round(delta_state._refresh_widen, 2)
                )
        await asyncio.sleep(
            max(
                float(refresh) * delta_state._refresh_widen,
                time.perf_counter() - render_started,
            )
        )


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


def _is_proactor_write_noise(context):
    """One known-harmless asyncio shape: the proactor write race."""
    exc = context.get("exception")
    handle = context.get("handle")
    return bool(isinstance(exc, AssertionError) and "_loop_writing" in repr(handle))


def _drop_proactor_write_noise(loop, context):
    """Swallow one known-harmless asyncio traceback, log everything else.

    _ProactorBaseWritePipeTransport._loop_writing can race with a transport
    that is finishing its previous write and raise
    `assert f is self._write_fut` inside asyncio itself. Python 3.10 prints
    the traceback on most launches of this tool; nothing in the product
    fails because of it (the write queue keeps flowing, and the tool is
    left running for hours in a single session). Upstream class of bug, not
    ours. Dropping just that one pattern keeps the customer's console
    honest: a traceback they see is a traceback worth calling support
    about.
    """
    if _is_proactor_write_noise(context):
        return
    loop.default_exception_handler(context)


_orig_call_exception_handler = asyncio.BaseEventLoop.call_exception_handler


def _quietly_call_exception_handler(self, context):
    """The drain lives on the dispatcher, not on one loop's handler.

    Any loop this process creates - the main asyncio loop, or a helper
    thread's own loop - gets the same net at the dispatch point instead of
    relying on per-loop installation timing. Everything else still flows.
    """
    if _is_proactor_write_noise(context):
        return
    _orig_call_exception_handler(self, context)


asyncio.BaseEventLoop.call_exception_handler = _quietly_call_exception_handler


async def main_async():
    # The asyncio loop thread makes its own COM calls (find_open_book on every
    # cycle, sheet probes, writes). pywin32 needs an apartment for those: the
    # opener thread ran CoInitialize for itself, but the writer thread never
    # did, and lazy app-handle resolution inside the loop then answered with
    # "OLE error 0x800a01a8" for a perfectly alive Excel - the writer saw no
    # books to write to even though the opener had just opened them.
    try:
        import pythoncom

        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        asyncio.get_running_loop().set_exception_handler(_drop_proactor_write_noise)
    except Exception:
        pass
    # First, undo whatever a previous crashed run left behind: a stuck
    # synthetic modifier is what once made the whole desktop act like
    # Alt was held down. Harmless when nothing is stuck.
    if os.name == "nt":
        _heal_keyboard()
    open_excel_files()
    # After opening the workbooks, not before. At process start xw.apps can be
    # empty even though the customer's Excel is already running, and the
    # recovery was then giving up and discarding the record while leaving Excel
    # stuck on manual calculation.
    recover_calc_mode()
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


from delta_instance import (
    kill_previous_instances,
)

















def recover_calc_mode():
    """Put Excel back on automatic if a previous run died mid-write-window.

    A crash, a kill from Task Manager or a power cut while calculation was
    suspended would otherwise leave the customer's formulas silently stopped.
    The record is consulted first; Excel's current mode is only read if there
    is something to undo.
    """
    should, mode, reason = calc_mode_guard.pending_recovery()
    if not should:
        if reason:
            print(f"  calc recovery not needed: {reason}", flush=True)
        return False
    print(f"  calculation-mode record found: {reason}", flush=True)
    try:
        # Recovery belongs to our private instance only. The old code scanned
        # the running object table and switched calculation on apps[0] -
        # whoever that happened to be, including the customer's own Excel.
        app = delta_state._owner_app
        if app is None:
            print("  calc recovery: no private Excel instance yet", flush=True)
            return False          # keep the record so a later launch can retry
        raw = int(app.api.Calculation)
        if raw != calc_mode_guard.XL_MANUAL:
            print(f"  calc recovery: Excel is not on manual (raw={raw})", flush=True)
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
    """Last-ditch restore on a clean exit, in case a cycle never finished."""
    restore = delta_state._last_bulk_restore
    if restore is None:
        return
    try:
        restore()
    except Exception:
        pass


atexit.register(_emergency_restore)


def _run_doctor():
    """Check the install and open the report for the customer.

    The executable is built without a console, so the report goes to a file
    and that file is opened. The customer sees what failed in plain English
    and support gets one file to read.
    """
    from doctor import run_doctor

    report_path = os.path.join(delta_state.BASE_DIR, "doctor-report.txt")
    try:
        report, ok = run_doctor(delta_state.BASE_DIR, out=report_path)
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


def main():
    # "delta_option_chain.exe --doctor" checks the install and writes a
    # report. It is the first thing to run on a machine where the tool is not
    # behaving, and it needs no Python on the customer's machine.
    if any(arg in ("--doctor", "--selftest", "/doctor") for arg in sys.argv[1:]):
        _run_doctor()
        return
    try:
        write_startup_diagnostics()
    except OSError as e:
        show_error(f"Could not write startup diagnostics:\n{e}")
    # Take over from any older instance BEFORE anything that can block.
    #
    # The activation dialog blocks, and the guard used to run after it. Two
    # launches therefore both sat on a dialog, neither ever reached the guard,
    # and both went on to write the same three workbooks. Two writers = two
    # chain sessions and a sheet that shows a mix of two snapshots.
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
    if not ensure_license():
        _heal_keyboard()
        os._exit(1)
    try:
        asyncio.run(main_async())
    except KeyboardInterrupt:
        pass
    finally:
        _emergency_restore()

if __name__ == "__main__":
    main()
