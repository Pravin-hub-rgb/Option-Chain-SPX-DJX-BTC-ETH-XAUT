"""Shared state for the MCX Commodity Option Chain tool.

Every piece of state the product has lives here, so there is exactly one copy
of each. The modules import the objects they need from here and mutate them
in place; the few places that REBIND a name rather than mutate it write to
mcx_state.<name> instead, so the change is seen by everyone.

Mirrors delta_state.py / us_state.py so the three products have the same
shape.
"""

import os
import time

import config
from mcx_performance import PerformanceMetrics

BASE_DIR = config.base_dir()

PRODUCT_ID = "mcx"

PRODUCT_EXE = "mcx_option_chain.exe"

_last_bulk_restore = None

# True while Excel has gone away and the tool is waiting for it to come back.
_excel_gone_logged = False

_cycle_wrote = False

_bulk_will_recalc = False

_last_recalc_ms = 0.0

# ScreenUpdating self-heal net.
_screen_restore_owed = False
_screen_watchdog_logged_at = 0.0
_recalc_owed = False

# Rate-limit the skip receipts so support sees one line, not one per cycle.
_user_input_skip_log_at = 0.0
_user_edit_skip_log_at = 0.0

_speed_log_at = 0.0

# Upstox data source.
UPSTOX_BASE_URL = str(config.load_config().get("upstox_base_url", "https://api.upstox.com/v2"))
UPSTOX_INSTRUMENTS_URL = str(config.load_config().get("upstox_instruments_url", ""))
MCX_SEGMENT = str(config.load_config().get("mcx_segment", "MCX_FO"))

MCX_SYMBOLS = [s.strip().upper() for s in str(config.load_config().get("mcx_symbols", "CRUDEOIL,GOLD,SILVER")).split(",") if s.strip()]

live_data = {sym: {} for sym in MCX_SYMBOLS}

# Last known good chain per (symbol, expiry_dmy): {rows, ts}. Kept so the
# sheet still shows something marked STALE when the market is closed and
# Upstox returns empty quotes.
last_good_chain = {sym: {} for sym in MCX_SYMBOLS}

spot_cmp = {sym: None for sym in MCX_SYMBOLS}

symbol_expiry = {sym: {} for sym in MCX_SYMBOLS}

expiry_symbols = {sym: {} for sym in MCX_SYMBOLS}

active_expiries = {sym: [] for sym in MCX_SYMBOLS}

active_expiries_dmy = {sym: [] for sym in MCX_SYMBOLS}

user_pinned = {sym: [False, False, False] for sym in MCX_SYMBOLS}

cfg = config.load_config()

HTTP_TIMEOUT = float(cfg.get("http_timeout_seconds", 15))

OI_POLL_SECONDS = float(cfg.get("oi_poll_seconds", 5))

EXPIRY_REFRESH_SECONDS = float(cfg.get("expiry_refresh_seconds", 30))

SPOT_POLL_SECONDS = float(cfg.get("spot_poll_seconds", 0.35))

STRATEGY_RECALC_MAX_MS = float(cfg.get("strategy_recalc_max_ms", 200.0))

NET_FAIL_LIMIT = 3

_net = {"fails": 0, "since": None}

_net_drawn = {}

_pending_date_loads = []

_date_req_sig = {}

_book_sheet_cache = {}

_app_cache = None

_last_rows = {}

_FIRST_SIGHT_BOUND = 300

ROW_EXTENT_DEFAULT = int(cfg.get("row_extent_default", 256))

_row_extent_seen = {}

_refresh_widen = 1.0

_refresh_widen_hold_until = 0.0

WIDEN_STEP = 1.5

WIDEN_MAX = 2.0

WIDEN_DECAY_SECONDS = 120.0

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

available_dates = {sym: [] for sym in MCX_SYMBOLS}

_dropdown_sig = {}

_usd_fmt_done = set()

_performance_metrics = PerformanceMetrics()

_last_performance_flush = time.monotonic()

_performance_log_error_shown = False

_excel_connection_status = {}

_last_excel_exception = {}

MONTHS = {"JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06",
          "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"}

BLOCK_COLS = [
    {"start": 1,  "atm": "A", "call_start": "A", "call_end": "D", "put_start": "F", "put_end": "I"},
    {"start": 12, "atm": "L", "call_start": "L", "call_end": "O", "put_start": "Q", "put_end": "T"},
    {"start": 23, "atm": "W", "call_start": "W", "call_end": "Z", "put_start": "AB", "put_end": "AE"},
]

STALE_FLAG_SECONDS = 10.0

feed_state = {sym: {"state": "LOADING", "detail": "", "retry_in": 0.0,
                    "last_good_ts": 0.0, "last_good_label": "never"} for sym in MCX_SYMBOLS}

# Upstox instrument key per underlying symbol (populated from instrument master,
# e.g. CRUDEOIL -> "MCX_FO|590045"). Fallback spot ticker symbols.
SPOT_TICKER_SYMBOLS = {sym: "" for sym in MCX_SYMBOLS}

# MCX spot is not a separate index symbol; we use the nearest futures contract
# as the spot reference. Populated at runtime.
SPOT_FUTURES_KEY = {sym: "" for sym in MCX_SYMBOLS}

last_ws_recv = {sym: 0.0 for sym in MCX_SYMBOLS}

last_spot_change = {sym: 0.0 for sym in MCX_SYMBOLS}

spot_source = {sym: None for sym in MCX_SYMBOLS}

STALE_AFTER_S = 3.0

excel_fail_since = {sym: None for sym in MCX_SYMBOLS}

excel_fail_notice_shown = {sym: False for sym in MCX_SYMBOLS}

excel_connected = {sym: False for sym in MCX_SYMBOLS}

excel_miss_counts = {sym: 0 for sym in MCX_SYMBOLS}

_auto_open_done = {sym: False for sym in MCX_SYMBOLS}

_sheet_altered_notice = {sym: False for sym in MCX_SYMBOLS}

_excel_liveness_cache = None

_bulk_owning_manual = False

# The Excel instance this tool owns.
_owner_app = None
_owner_hwnd = None

_enum_unknown = {}
_missing_since = {}
_all_gone_since = None

_open_pass_at = None
_open_pass_retried = False

_calc_state_probe = None

MANUAL_WATCHDOG_SECONDS = 4.0

_CALC_IDS = {"automatic": -4105, "semiautomatic": 2, "manual": -4135}

_CALC_ID_NAMES = {-4105: "automatic", 2: "semiautomatic", -4135: "manual"}

_last_observed_calc_raw = None

_calc_churn_count = 0

_calc_churn_summary_after = 0.0

_CALC_CHURN_SUMMARY_SECONDS = 60.0

_calc_watchdog_skip_logged_at = 0.0

# Parity spot (commodity options: q=0, r small positive)
PARITY_RATE = float(cfg.get("parity_rate", 0.045))
PARITY_DIVIDEND_Q = float(cfg.get("parity_dividend_q", 0.0))
PARITY_USE_N_STRIKES = int(cfg.get("parity_use_n_strikes", 5))
PARITY_DEADBAND = float(cfg.get("parity_deadband", 0.05))
PARITY_MAX_SPREAD = float(cfg.get("parity_max_spread", 2.0))
PARITY_MAX_SPREAD_PCT = float(cfg.get("parity_max_spread_pct", 0.05))

_parity_spot_cache = {}

_PUT_INTRINSIC_WARNED = {}
_PUT_INTRINSIC_WARN_INTERVAL = 300.0

_calc_failed_at = 0.0
CALC_RETRY_COOLDOWN_SECONDS = 10.0
_calc_manual_failed = False
_calc_notice_shown = False
_calc_failed_since = 0.0
_no_book_skip_logged_at = 0.0
