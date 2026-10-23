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
        import delta_diagnostics
        self.assertEqual(_LOG_ROOT, delta_diagnostics.FEED_LOG_DIR)

    def test_delta_log_dir_is_not_the_customer_directory(self):
        import delta_diagnostics
        self.assertNotEqual(
            os.path.join(os.environ.get("LOCALAPPDATA", "."), "DeltaOptionChain"),
            delta_diagnostics.FEED_LOG_DIR,
        )

    def test_delta_app_exception_log_honours_the_override(self):
        """app.log/startup diag writes used to bypass the override and leak
        test noise into the customer's real app.log."""
        import delta_diagnostics
        self.assertEqual(
            _LOG_ROOT,
            delta_diagnostics._log_directory(),
        )
        exc = RuntimeError("isolation probe")
        exc.__traceback__ = None
        delta_diagnostics.append_app_exception("isolation probe", exc)
        self.assertTrue(os.path.exists(os.path.join(_LOG_ROOT, "app.log")))


if __name__ == "__main__":
    unittest.main()
