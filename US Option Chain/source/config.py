import json
import os
import sys


def base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    path = os.path.join(base_dir(), "config.json")
    defaults = {
        "refresh_interval_seconds": 0.1,
        "poll_gap_seconds": 2,
        "round_pause_seconds": 2,
        "expiry_refresh_seconds": 60,
        "spx_symbol": "SPX",
        "djx_symbol": "DJX",
        "default_symbol": "AAPL",
        "data_source": "auto",
    }
    try:
        with open(path) as f:
            cfg = json.load(f)
    except (OSError, ValueError):
        return defaults
    for k, v in defaults.items():
        cfg.setdefault(k, v)
    return cfg
