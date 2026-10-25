"""Shared state for the Delta Option Chain tool.

Every piece of state the product has lives here, so there is exactly one copy
of each. The modules import the objects they need from here and mutate them
in place; the few places that REBIND a name rather than mutate it write to
delta_state.<name> instead, so the change is seen by everyone.

This mirrors the US tool's us_state.py exactly, so the two products have the
same shape and a fix to one reads the same in the other.
"""

import os
import time

import config
from delta_performance import PerformanceMetrics

BASE_DIR = config.base_dir()

PRODUCT_ID = "delta"

PRODUCT_EXE = "delta_option_chain.exe"

_last_bulk_restore = None

# True while Excel has gone away and the tool is waiting for it to come back.
_excel_gone_logged = False

_cycle_wrote = False

_bulk_will_recalc = False

_last_recalc_ms = 0.0

# ScreenUpdating self-heal net. The tool no longer toggles ScreenUpdating: the
# old window era turned one failed restore into a frozen frame, and the toggle
# itself makes Excel repaint the entire sheet on every release, which reads on
# a fast render cadence as constant flicker. The only writer left in the
# product is _screen_update_watchdog_tick (force ON, never OFF). These fields
# keep its backoff and the owed latch honest.
_screen_restore_owed = False
_bulk_window_open = False
_screen_watchdog_logged_at = 0.0
# A render cycle skipped because the customer's hands were on our Excel
# owes its forced recalculation to the next writable cycle.
_recalc_owed = False

# Rate-limit the skip receipts so support sees one line, not one per cycle.
_user_input_skip_log_at = 0.0
_user_edit_skip_log_at = 0.0

_speed_log_at = 0.0

REST_BASE = "https://api.india.delta.exchange"

WS_URL = "wss://socket.india.delta.exchange"

INDEX_WS_URL = "wss://public-socket.india.delta.exchange"

live_data = {"BTC": {}, "ETH": {}, "XAUT": {}}

spot_cmp = {"BTC": None, "ETH": None, "XAUT": None}

symbol_expiry = {"BTC": {}, "ETH": {}, "XAUT": {}}

expiry_symbols = {"BTC": {}, "ETH": {}, "XAUT": {}}

active_expiries = {"BTC": [], "ETH": [], "XAUT": []}

active_expiries_dmy = {"BTC": [], "ETH": [], "XAUT": []}

user_pinned = {"BTC": [False, False, False], "ETH": [False, False, False], "XAUT": [False, False, False]}

cfg = config.load_config()

HTTP_TIMEOUT = float(cfg.get("http_timeout_seconds", 15))

OI_POLL_SECONDS = float(cfg.get("oi_poll_seconds", 5))

EXPIRY_REFRESH_SECONDS = float(cfg.get("expiry_refresh_seconds", 30))

SPOT_POLL_SECONDS = float(cfg.get("spot_poll_seconds", 0.35))

NET_FAIL_LIMIT = 3

_net = {"fails": 0, "since": None}

_net_drawn = {}

_pending_date_loads = []

_date_req_sig = {}

_book_sheet_cache = {}

_sheet_cache = {}

_app_cache = None

_last_rows = {}

_FIRST_SIGHT_BOUND = 300

# Every chain block is written to a fixed number of rows (ROW_EXTENT_DEFAULT,
# grown only when a chain really is bigger), padded with blanks below the live
# strike count. The customer's strategy sheets carry formulas with literal
# ranges - XLOOKUP(E6, BTC!P5:P33, BTC!R5:R33, 0) - and a chain whose live row
# count moves with every expiry silently walks rows out from under those
# ranges: one leg of their lookup then returns the 0 default and their book
# shows values that are wrong with no error anywhere. A fixed geometry keeps
# every row their formula can see inside the block forever.
ROW_EXTENT_DEFAULT = 256

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

# Real Application.Calculate cost, used to drive the adaptive refresh. The
# CalculationState probe only ever reads 0 in practice, so the measured recalc
# duration drives the backoff instead of the no-op probe.
_recalc_backoff_seconds = 0.0
RECALC_BACKOFF_FACTOR = float(cfg.get("recalc_backoff_factor", 5.0))
RECALC_BACKOFF_THRESHOLD_SECONDS = float(cfg.get("recalc_backoff_threshold_seconds", 0.025))
RECALC_BACKOFF_ALPHA = 0.5
RECALC_BACKOFF_DECAY = 0.9
MAX_REFRESH_SECONDS = float(cfg.get("max_refresh_seconds", 5.0))

_block_payload = {}

_block_values = {}

_block_hidden = {}

_colour_state = {}

available_dates = {"BTC": [], "ETH": [], "XAUT": []}

_dropdown_sig = {}

_usd_fmt_done = set()

_performance_metrics = PerformanceMetrics()

_last_performance_flush = time.monotonic()

_performance_log_error_shown = False

_excel_connection_status = {}

_last_excel_exception = {}

MONTHS = {"JAN":"01","FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06",
          "JUL":"07","AUG":"08","SEP":"09","OCT":"10","NOV":"11","DEC":"12"}

MONTHS_REV = {v:k for k,v in MONTHS.items()}

BLOCK_COLS = [
    {"start": 1,  "atm": "A", "call_start": "A", "call_end": "D", "put_start": "F", "put_end": "I"},
    {"start": 12, "atm": "L", "call_start": "L", "call_end": "O", "put_start": "Q", "put_end": "T"},
    {"start": 23, "atm": "W", "call_start": "W", "call_end": "Z", "put_start": "AB", "put_end": "AE"},
]

STALE_FLAG_SECONDS = 10.0

feed_state = {"BTC": {"state": "LOADING", "detail": "", "retry_in": 0.0,
                      "last_good_ts": 0.0, "last_good_label": "never"},
              "ETH": {"state": "LOADING", "detail": "", "retry_in": 0.0,
                      "last_good_ts": 0.0, "last_good_label": "never"},
              "XAUT": {"state": "LOADING", "detail": "", "retry_in": 0.0,
                       "last_good_ts": 0.0, "last_good_label": "never"}}

SPOT_TICKER_SYMBOLS = {"BTC": "BTCUSD", "ETH": "ETHUSD", "XAUT": "XAUTUSD"}

INDEX_SYMBOLS = {"BTC": ".DEXBTUSD", "ETH": ".DEETHUSD", "XAUT": ".DEXAUTUSD"}

INDEX_CHANNEL = "spot_price"

INDEX_ASSET = {v: k for k, v in INDEX_SYMBOLS.items()}

last_ws_recv = {"BTC": 0.0, "ETH": 0.0, "XAUT": 0.0}

last_spot_change = {"BTC": 0.0, "ETH": 0.0, "XAUT": 0.0}

spot_source = {"BTC": None, "ETH": None, "XAUT": None}

STALE_AFTER_S = 3.0

excel_fail_since = {"BTC": None, "ETH": None, "XAUT": None}

excel_fail_notice_shown = {"BTC": False, "ETH": False, "XAUT": False}

excel_connected = {"BTC": False, "ETH": False, "XAUT": False}

excel_miss_counts = {"BTC": 0, "ETH": 0, "XAUT": 0}

_auto_open_done = {"BTC": False, "ETH": False, "XAUT": False}

_sheet_altered_notice = {"BTC": False, "ETH": False, "XAUT": False}

_excel_liveness_cache = None

_bulk_owning_manual = False

# The Excel instance this tool owns. Book finding, the calc session and the
# miss bookkeeping work only inside it; no other Excel on the machine is
# ever touched, so the two chain tools cannot interfere with each other.
_owner_app = None
_owner_hwnd = None

# Workbook-attach bookkeeping: a failed COM enumeration is "unknown" for a
# short TTL, never "missing"; `_missing_since` refreshes every cycle a
# workbook is confirmed absent and drives the close-all exit.
_enum_unknown = {}
_missing_since = {}
_all_gone_since = None

# The startup open pass, run inline on the display thread. `_open_pass_at`
# timestamps the attempt so the display loop can retry one wedged pass once
# after a grace period; `_open_pass_retried` limits it to that one retry.
_open_pass_at = None
_open_pass_retried = False

_calc_state_probe = None

_last_window_status = {}

_SHORT_BUILD = None

_manual_since = 0.0

MANUAL_WATCHDOG_SECONDS = 4.0

_CALC_IDS = {"automatic": -4105, "semiautomatic": 2, "manual": -4135}

_CALC_ID_NAMES = {-4105: "automatic", 2: "semiautomatic", -4135: "manual"}

_last_observed_calc_raw = None

_calc_churn_count = 0

_calc_churn_summary_after = 0.0

_CALC_CHURN_SUMMARY_SECONDS = 60.0

_calc_watchdog_skip_logged_at = 0.0

# Parity spot (crypto options: q=0, r small positive)
PARITY_RATE = float(cfg.get("parity_rate", 0.045))
PARITY_DIVIDEND_Q = float(cfg.get("parity_dividend_q", 0.0))
PARITY_USE_N_STRIKES = int(cfg.get("parity_use_n_strikes", 5))
PARITY_DEADBAND = float(cfg.get("parity_deadband", 0.05))
PARITY_TOLERANCE_POINTS = float(cfg.get("parity_tolerance_points", 3.0))
PARITY_TOLERANCE_PCT = float(cfg.get("parity_tolerance_pct", 0.002))
PARITY_MAX_SPREAD = float(cfg.get("parity_max_spread", 2.0))
PARITY_MAX_SPREAD_PCT = float(cfg.get("parity_max_spread_pct", 0.05))
PARITY_SANITY_PCT = float(cfg.get("parity_sanity_pct", 0.05))

_parity_spot_cache = {}
_spot_divergence = {}
_parity_log_at = {}

# Session-level calc retry cooldown: once Excel refuses to switch to manual
# calculation in a cycle, we stop trying until the cooldown expires so we do not
# hammer COM and make the customer's workbook crawl.
_calc_failed_at = 0.0
CALC_RETRY_COOLDOWN_SECONDS = 10.0
_calc_manual_failed = False
_calc_notice_shown = False
# Monotonic timestamp of the FIRST refusal of the current stuck spell.
# _calc_failed_at is refreshed by every cooldown retry, so it cannot tell a
# customer who glanced at a save prompt from one stuck for a minute. The
# notice is only shown once a refusal has held for the grace period, so a
# workbook close - where the save prompt refuses one or two switches for a
# few seconds - never warns at all.
_calc_failed_since = 0.0
# Rate limiter for the "no chain workbook open" skip log, so the chains-closed
# state does not append a line to the diagnostics log on every render cycle.
_no_book_skip_logged_at = 0.0

