"""Test isolation for every pytest run, not just the full suite.

Discovery sorts test files alphabetically, so test_00_isolation.py runs
first in a whole-suite run - but a targeted run like `pytest test_us_lag.py`
never imports it, and without the override the product modules then write
their test noise straight into the customer's real feed logs (this happened
on 23-10-2026: calc_* test events landed in the live feed.log). conftest.py
is imported by pytest before any test file, so the override is always set.
"""

import os
import sys
import tempfile

_SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
_TEST_LOG_ROOT = tempfile.mkdtemp(prefix="us_oc_test_logs_")
os.environ["US_OPTION_CHAIN_LOG_DIR"] = _TEST_LOG_ROOT

if _SOURCE_DIR not in sys.path:
    sys.path.insert(0, _SOURCE_DIR)
