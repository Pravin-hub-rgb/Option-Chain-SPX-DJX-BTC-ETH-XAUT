import json
import os
import sys


def base_dir():
    if getattr(sys, "frozen", False):
        # PyInstaller: sys.executable is the real exe
        return os.path.dirname(sys.executable)
    if "__compiled__" in globals():
        # Nuitka onefile: __file__/sys.executable point into the temp
        # extraction dir; the real exe location stays in sys.argv[0]
        return os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    path = os.path.join(base_dir(), "config.json")
    try:
        with open(path) as f:
            return json.load(f)
    except (OSError, ValueError):
        return {"refresh_interval_seconds": 0.1}
