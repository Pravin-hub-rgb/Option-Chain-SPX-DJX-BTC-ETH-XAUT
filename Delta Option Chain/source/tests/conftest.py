import os
import sys
import tempfile

_SOURCE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _SOURCE_DIR not in sys.path:
    sys.path.insert(0, _SOURCE_DIR)

_LOG_ROOT = tempfile.mkdtemp(prefix="delta_oc_test_logs_")
os.environ.setdefault("DELTA_OPTION_CHAIN_LOG_DIR", _LOG_ROOT)
