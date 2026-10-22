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
import time

import databento_feed

from us_diagnostics import (
    append_app_exception,
    write_startup_diagnostics,
    feed_log,
)
from us_licensing import ensure_license, show_error
from us_workbook import INDEX_SYMBOLS, open_excel_files
from us_instance import (
    kill_previous_instances, recover_calc_mode, _emergency_restore,
)
from us_chain import (
    active_symbols, available_expiries, data_poller, ensure_sym_state,
    expiry_roller, init_symbol, process_pending_dates, spot_poller,
    _short_build,
)

import us_state
from us_state import (
    BASE_DIR, PRODUCT_EXE, PRODUCT_ID, _databento_feed, expiry_roots,
)

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
