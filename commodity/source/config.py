import json
import os
import sys


def base_dir():
    if getattr(sys, "frozen", False):
        # PyInstaller: sys.executable is the real exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    path = os.path.join(base_dir(), "config.json")
    defaults = {
        "refresh_interval_seconds": 0.5,
        "input_grace_seconds": 1.5,
        "oi_poll_seconds": 2,
        "expiry_refresh_seconds": 30,
        "spot_poll_seconds": 0.35,
        "http_timeout_seconds": 15,
        "strategy_recalc_max_ms": 200.0,
        "upstox_base_url": "https://api.upstox.com/v2",
        "upstox_instruments_url": "https://assets.upstox.com/market-quote/instruments/exchange/MCX.json.gz",
        "mcx_symbols": "CRUDEOIL,GOLD,SILVER",
        "mcx_segment": "MCX_FO",
        "row_extent_default": 256,
        "recalc_backoff_factor": 5.0,
        "recalc_backoff_threshold_seconds": 0.025,
        "max_refresh_seconds": 5.0,
    }
    try:
        with open(path) as f:
            cfg = json.load(f)
    except (OSError, ValueError):
        return defaults
    for k, v in defaults.items():
        cfg.setdefault(k, v)
    return cfg
