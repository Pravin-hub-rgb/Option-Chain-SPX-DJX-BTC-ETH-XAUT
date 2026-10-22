"""Test isolation: point this product's logs at a throwaway directory.

Discovery imports test files alphabetically, so this lowest-named file runs
first and must set the log-directory override BEFORE any product module is
imported, or the suite writes straight into the customer's real feed logs.
"""

import os
import tempfile

_LOG_ROOT = tempfile.mkdtemp(prefix="oc_test_logs_")
os.environ["DELTA_OPTION_CHAIN_LOG_DIR"] = _LOG_ROOT

import sys

_SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
if _SOURCE_DIR not in sys.path:
    sys.path.insert(0, _SOURCE_DIR)

import unittest


class IsolationTests(unittest.TestCase):
    def test_delta_log_dir_override_is_honoured(self):
        import main
        self.assertEqual(_LOG_ROOT, main.FEED_LOG_DIR)

    def test_delta_log_dir_is_not_the_customer_directory(self):
        import main
        self.assertNotEqual(
            os.path.join(os.environ.get("LOCALAPPDATA", "."), "DeltaOptionChain"),
            main.FEED_LOG_DIR,
        )


if __name__ == "__main__":
    unittest.main()
