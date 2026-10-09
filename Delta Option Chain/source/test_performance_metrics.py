import os
import tempfile
import unittest
from unittest.mock import patch

from performance_metrics import PerformanceMetrics
from startup_diagnostics import (
    append_app_exception,
    append_connection_event,
    append_performance_summary,
)


class PerformanceMetricsTests(unittest.TestCase):
    def test_summary_aggregates_timings_and_row_counts(self):
        metrics = PerformanceMetrics()
        metrics.record("BTC", "chain_build", 0.01, rows=12)
        metrics.record("BTC", "chain_build", 0.03, rows=8)

        summary = metrics.summary_lines(now=metrics._started + 30)

        self.assertIn("performance.BTC.chain_build: count=2, avg_ms=20.00, max_ms=30.00, rows=20", summary)
        self.assertIn(
            "performance.interval_seconds=30.0; scopes are assets or ALL", summary
        )

    def test_reset_starts_a_fresh_aggregation_interval(self):
        metrics = PerformanceMetrics()
        metrics.record("ALL", "refresh_overrun", 0.02)
        reset_time = metrics._started + 30
        metrics.reset(reset_time)

        summary = metrics.summary_lines(now=reset_time + 30)

        self.assertEqual(
            summary,
            ["performance.interval_seconds=30.0; scopes are assets or ALL"],
        )

    def test_summary_is_appended_to_user_diagnostics_log(self):
        with tempfile.TemporaryDirectory() as temp:
            with patch.dict(os.environ, {"LOCALAPPDATA": temp}):
                path = append_performance_summary(
                    ["performance.ALL.render_cycle: count=1, avg_ms=20.00"]
                )

            with open(path, encoding="utf-8") as log_file:
                contents = log_file.read()

        self.assertIn("--- Delta Option Chain performance ---", contents)
        self.assertIn("performance.ALL.render_cycle", contents)

    def test_large_diagnostic_log_is_rotated_before_append(self):
        with tempfile.TemporaryDirectory() as temp:
            log_dir = os.path.join(temp, "DeltaOptionChain")
            os.makedirs(log_dir)
            path = os.path.join(log_dir, "startup_diag.log")
            with open(path, "w", encoding="utf-8") as log_file:
                log_file.write("old log\n" + ("x" * (1024 * 1024)))

            with patch.dict(os.environ, {"LOCALAPPDATA": temp}):
                append_performance_summary(["performance.interval_seconds=30.0"])

            with open(path + ".1", encoding="utf-8") as rotated:
                self.assertTrue(rotated.read().startswith("old log"))
            with open(path, encoding="utf-8") as current:
                self.assertIn("performance.interval_seconds=30.0", current.read())

    def test_connection_transitions_are_recorded_without_workbook_values(self):
        with tempfile.TemporaryDirectory() as temp:
            with patch.dict(os.environ, {"LOCALAPPDATA": temp}):
                append_connection_event("BTC", "connected", "workbook=C:\\Delta\\btc_chain.xlsx")
                path = os.path.join(
                    temp, "DeltaOptionChain", "startup_diag.log"
                )
                with open(path, encoding="utf-8") as log_file:
                    contents = log_file.read()

        self.assertIn("asset=BTC", contents)
        self.assertIn("status=connected", contents)
        self.assertNotIn("price=", contents)
        self.assertNotIn("strategy", contents)

    def test_unexpected_exception_is_written_with_traceback(self):
        with tempfile.TemporaryDirectory() as temp:
            with patch.dict(os.environ, {"LOCALAPPDATA": temp}):
                try:
                    raise RuntimeError("Excel COM failure")
                except RuntimeError as exc:
                    path = append_app_exception("Excel update for BTC", exc)

            with open(path, encoding="utf-8") as log_file:
                contents = log_file.read()

        self.assertIn("context=Excel update for BTC", contents)
        self.assertIn("RuntimeError: Excel COM failure", contents)
        self.assertIn("Traceback (most recent call last)", contents)


if __name__ == "__main__":
    unittest.main()
