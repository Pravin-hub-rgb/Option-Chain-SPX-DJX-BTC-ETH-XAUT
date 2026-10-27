"""Shared module-level state for the US Option Chain tool.

Everything here is one object shared across the whole program: caches keyed
by symbol or role, Excel book handles, locks, and the small scalars the
render loop flips every cycle. main.py and the split modules all import the
same objects from here, so there is exactly one copy of each piece of state.

The definitions are moved verbatim from main.py; only the import wiring
changed. Runtime names that functions re-bind through `global` live in
main.py until their owning module is extracted, which keeps the meshes
honest one line at a time.
"""

import asyncio
import threading
import time

import config
from us_performance import PerformanceMetrics

BASE_DIR = config.base_dir()

# Identifies this tool inside the shared calculation-mode record, so US and
# the Delta tool running together in one Excel instance do not fight over it.
PRODUCT_ID = "us"

# The executable is named for the product, so the process name alone says which
# tool it is. The customer can see it in Task Manager, and the instance guard
# never has to guess the product from the folder - which is what went wrong
# when both packages were extracted into the same place.
PRODUCT_EXE = "us_option_chain.exe"

# The most recent bulk-mode restore callable, so an exit path that never
# reaches the render loop's finally can still put Excel back.
_last_bulk_restore = None

cfg = config.load_config()

POLL_GAP = float(cfg.get("poll_gap_seconds", 2))
ROUND_PAUSE = float(cfg.get("round_pause_seconds", 2))
EXPIRY_REFRESH = float(cfg.get("expiry_refresh_seconds", 60))
REFRESH = float(cfg.get("refresh_interval_seconds", 0.5))
# A write window stays shut while the trader's hands are on our Excel. 0.7 s
# was shorter than a normal pause between words, so a window could open in the
# middle of an edit; 1.5 s is long enough to cover real typing and still
# resumes as soon as the keys go quiet.
INPUT_GRACE_SECONDS = float(cfg.get("input_grace_seconds", 1.5))
# How often each symbol's chain is rebuilt. The old loop rebuilt all three in
# one pass and slept between passes, so every symbol was about six seconds
# behind. Each symbol now has its own period.
SYMBOL_POLL_SECONDS = float(cfg.get("symbol_poll_seconds", 0.5))
# Floor on how often the same symbol may be rebuilt. Small, because the data
# is already in memory from the live feed.
SOURCE_MIN_INTERVAL = float(cfg.get("source_min_interval_seconds", 0.25))
SPX_SYMBOL = str(cfg.get("spx_symbol", "SPX")).upper()
DJX_SYMBOL = str(cfg.get("djx_symbol", "DJX")).upper()
DEFAULT_SYMBOL = str(cfg.get("default_symbol", "AAPL")).upper()

# --- connection health indicator ------------------------------------------
# Market/API ghatne par sheet freeze ho jata hai; user ko pata chalna chahiye
# ki data stale hai. NET_FAIL_LIMIT consecutive failures ke baad header me red
# tag dikhta hai, aur connection wapas aate hi apne aap hat jaata hai.
NET_FAIL_LIMIT = 2
_CALC_IDS = {"automatic": -4105, "semiautomatic": 2, "manual": -4135}
_net = {"fails": 0, "since": None}
_net_drawn = {}   # (role, block) -> last text written to that tag cell

# Expiry -> set of roots seen for it in the latest payload, so the expiry
# picker can tell a standard expiry (has the base root) from a weekly-only one.
expiry_roots = {}
# base root -> expiry -> count of contracts that actually report OI > 0
oi_roots = {}

# Configured once in main(), after the owner/customer key has been resolved.
_databento_feed = None
ROLE_FILES = {
    "spx": "spx_option_chain.xlsx",
    "djx": "djx_option_chain.xlsx",
    "stock": "us_stock_option_chain.xlsx",
    "btc": "btc_option_chain.xlsx",
}
ROLE_ORDER = ("spx", "djx", "stock", "btc")
# roles whose K1 is locked (not user-editable)
LOCKED_ROLES = frozenset({"spx", "djx", "btc"})

# active ticker per Excel file/role (spx/djx fixed; stock from K1)
role_symbol = {"spx": SPX_SYMBOL, "djx": DJX_SYMBOL, "stock": DEFAULT_SYMBOL, "btc": "BTC"}
SYMBOL_CELL = "K1"

live_data = {}
spot_cmp = {}
active_expiries = {}
user_pinned = {}
all_known_expiries = {}

# The expiry date the tool itself last wrote into each (symbol, slot) cell,
# and the last value it saw there. The difference between the two is what
# tells a customer edit from a leftover.
#
# Without this the tool read its own write-back and took it for a user edit.
# That pinned the slots to whatever the workbook happened to hold from the
# previous session - a monthly expiry - and the expiry roller then left them
# alone for good, because a pinned slot is never changed. The customer saw
# November, December and January side by side and no way out of it.
_tool_written_dates = {}

# (symbol, slot) -> the last expiry date the tool saw in that cell. A cell
# that has not changed since the last look was not edited by anybody.
_user_dates_seen = {}

# (symbol, slot) -> the expiry the tool last wrote into the date cell. The
# cell has to follow the block under it, and this is how the tool knows
# whether it already has.
_expiry_cell_written = {}

_rate_lock = asyncio.Lock()
# Serialises access to the shared Databento feed so two symbols never mutate
# it at the same time. It does not add any delay of its own.
_feed_access = asyncio.Semaphore(1)
_last_source_call = 0.0
# Last rebuild time per symbol, so the same symbol is not rebuilt twice in a
# row while another symbol is still waiting its turn.
_last_symbol_call = {}
_pending_date_loads = []
_dropdown_sig = {}
_refetch_now = False

_template_ready = set()

_spot_cache = {}          # root -> (price, age_seconds, source, monotonic, quote_ts)
_spot_lock = threading.Lock()

_auto_open_done = {r: False for r in ROLE_ORDER}
# Set when a workbook's layout no longer matches what the tool writes to.
_sheet_altered = {r: False for r in ROLE_ORDER}

# The Excel instance this tool owns. Nothing else in the module attaches to
# any other Excel: workbook finding, calculation mode and the save-prompt
# answerer all work solely inside it.
_owner_app = None
_owner_hwnd = None

# Workbook-attach bookkeeping. A failed COM enumeration means "unknown, try
# again next cycle", never "the workbook is missing"; `_missing_since`
# refreshes each cycle a workbook is confirmed absent and drives the
# close-all exit.
_enum_unknown = {}
_missing_since = {}
_all_gone_since = None

# The startup open pass, run inline on the display thread. `_open_pass_at`
# timestamps the attempt so the display loop can retry one wedged pass once
# after a grace period; `_open_pass_retried` limits it to that one retry.
_open_pass_at = None
_open_pass_retried = False

_SHORT_BUILD = None

_timezone_warning_logged = False
_excel_gone_logged = False

_excel_liveness_cache = None  # (monotonic_timestamp, running_bool)
_calc_state_probe = None

_last_rows = {}
# How far below the live table to sweep the first time we see a block in this
# run, to catch anything an earlier run left behind. Paid once per block.
_FIRST_SIGHT_BOUND = 300
# Every chain block is written to a fixed number of rows (ROW_EXTENT_DEFAULT,
# grown only when a chain really is bigger), padded with blanks below the live
# strike count. The customer's strategy sheets carry formulas with literal
# ranges - XLOOKUP(E6, BTC!P5:P33, BTC!R5:R33, 0) - and a chain whose live row
# count moves with every expiry silently walks rows out from under those
# ranges: one leg of their lookup then returns the 0 default and their book
# shows values that are wrong with no error anywhere. A fixed geometry keeps
# every row their formula can see inside the block forever.
ROW_EXTENT_DEFAULT = 128
_row_extent_seen = {}
# Adaptive refresh. When Excel still owed a recalculation pass the moment we
# handed the mode back (excel_recalc_second_pass), the customer's own formula
# workbook is heavier than the configured refresh assumed, so the render loop
# stretches its sleep to give Excel idle time instead of starving it. The
# widening decays back to 1.0 after a quiet spell.
_refresh_widen = 1.0
_refresh_widen_hold_until = 0.0
WIDEN_STEP = 1.5
WIDEN_MAX = 2.0
WIDEN_DECAY_SECONDS = 120.0
_performance_metrics = PerformanceMetrics()
_last_performance_flush = time.monotonic()
_performance_log_error_shown = False
_excel_connection_status = {}
_last_excel_exception = {}
_cycle_wrote = False
_bulk_will_recalc = False
# Who owns Excel's calculation mode, for the whole session.
#
# "owned" is true only when this tool changed the mode itself. If Excel was
# already on manual calculation - because the customer wants that, or because
# the other OptionChain tool owns it - it is left exactly as it was found.
_calc_session = {
    "owned": False,
    "failed": False,
    "mode": None,
    "app": None,
}
# The customer is told once, not on every cycle.
_calc_notice_shown = False
# Screen updating is set and restored around every write. When Excel is busy
# that switch fails, and four failing calls a cycle make Excel busier still.
# After a few failures the attempt is paused for a while.
_screen_failures = 0
_SCREEN_FAILURE_LIMIT = 3
_SCREEN_BACKOFF_SECONDS = 60.0
# ScreenUpdating self-heal net. The tool no longer toggles ScreenUpdating: the
# write/save window era turned one failed restore into a frozen frame, and the
# toggle itself makes Excel repaint the entire sheet on every release, which
# reads on a fast render cadence as constant flicker. The only writer left in
# the product is _screen_update_watchdog_tick (force ON, never OFF). These
# fields keep its backoff and the owed latch honest.
_screen_restore_owed = False
_bulk_owning_manual = False
_bulk_window_open = False
_screen_watchdog_logged_at = 0.0
# A render cycle skipped because the customer's hands were on our Excel
# owes its forced recalculation to the next writable cycle.
_recalc_owed = False
# Rate-limit the skip receipts so support sees one line, not one per cycle.
_user_input_skip_log_at = 0.0
_user_edit_skip_log_at = 0.0
_strategy_sheet_skip_log_at = 0.0

_last_recalc_ms = 0.0

# Strategy-safe recalc (Option C): when parked on a non-chain sheet, run a live
# Application.Calculate() only if the last measured full-workbook Calculate cost
# was below this threshold (ms). Cheap workbooks get live P&L; heavy ones fall
# back to the deferred (Option B) behaviour. Tunable; 0/unknown measurement
# always skips.
STRATEGY_RECALC_MAX_MS = float(cfg.get("strategy_recalc_max_ms", 200.0))

# Real Application.Calculate cost, used to drive the adaptive refresh. The
# CalculationState probe (_calc_state_probe above) only ever reads 0 in
# practice - it runs after a synchronous Calculate that already finished - so it
# never stretched the loop. The measured recalc duration is the signal instead.
_recalc_backoff_seconds = 0.0
RECALC_BACKOFF_FACTOR = float(cfg.get("recalc_backoff_factor", 5.0))
RECALC_BACKOFF_THRESHOLD_SECONDS = float(cfg.get("recalc_backoff_threshold_seconds", 0.025))
RECALC_BACKOFF_ALPHA = 0.5
RECALC_BACKOFF_DECAY = 0.9
MAX_REFRESH_SECONDS = float(cfg.get("max_refresh_seconds", 5.0))
# Cached (book, sheet) COM objects per role — book.sheets[0] is a COM round-trip
# every call, and it dominated the render loop. Cached and dropped on staleness.
_book_sheet_cache = {}
_sheet_cache = {}

_speed_log_at = 0.0

_app_cache = None

# Set when Excel refused to switch to manual calculation. The write window
# is then unprotected, so every cell we touch makes Excel recalculate. When
# that happens the render loop slows itself down instead of fighting Excel,
# and the customer is told once what to clear.
_calc_manual_failed = False
_calc_manual_notice_shown = False
_calc_failed_at = 0.0
# Monotonic timestamp of the FIRST refusal of the current stuck spell.
# _calc_failed_at is refreshed by every cooldown retry, so it cannot tell a
# customer who glanced at a save prompt from one stuck for a minute. The
# notice is only shown once a refusal has held for the grace period, so a
# workbook close - where the save prompt refuses one or two switches for a
# few seconds - never warns at all.
_calc_failed_since = 0.0
CALC_RETRY_COOLDOWN_SECONDS = 10.0
# Set when the customer has been told that Excel would not switch to manual
# calculation, so the message appears once and not on every cycle.
_calc_notice_shown = False
# When the write window is unprotected, write at most this often.
_UNPROTECTED_WRITE_INTERVAL = 3.0
_last_unprotected_write = 0.0
# Rate limiter for the "no chain workbook open" skip log, so the chains-closed
# state does not append a line to startup_diag.log on every render cycle.
_no_book_skip_logged_at = 0.0

# Set once the customer has been told that Databento rejected their key, so
# the message appears a single time and not on every cycle.
_databento_auth_notice_shown = False

_manual_since = 0.0
MANUAL_WATCHDOG_SECONDS = 4.0


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
    return ""
