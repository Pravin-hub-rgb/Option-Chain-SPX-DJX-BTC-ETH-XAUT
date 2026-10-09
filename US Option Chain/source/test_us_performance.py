import os
import tempfile
import unittest
from unittest.mock import patch

from us_diagnostics import append_app_exception, append_connection_event
from us_performance import PerformanceMetrics, itm_color_ranges


class ItmColorRangeTests(unittest.TestCase):
    def setUp(self):
        self.block = {
            "call_start": "A",
            "call_end": "D",
            "put_start": "F",
            "put_end": "I",
        }

    def test_ranges_match_per_row_rules_around_spot(self):
        self.assertEqual(
            itm_color_ranges(self.block, [90, 95, 100, 105, 110], 100),
            [
                ("A5:D6", (198, 224, 180)),
                ("F8:I9", (255, 200, 200)),
            ],
        )

    def test_no_itm_ranges_when_spot_is_outside_chain(self):
        self.assertEqual(
            itm_color_ranges(self.block, [90, 95, 100], 120),
            [("A5:D7", (198, 224, 180))],
        )
        self.assertEqual(
            itm_color_ranges(self.block, [90, 95, 100], 80),
            [("F5:I7", (255, 200, 200))],
        )

    def test_empty_chain_has_no_color_ranges(self):
        self.assertEqual(itm_color_ranges(self.block, [], 100), [])

    def test_batched_ranges_match_legacy_per_row_colors(self):
        for strikes in ([90, 95, 100, 105, 110], [80, 90, 100], [100, 110, 120]):
            for spot in (85, 95, 100, 107, 125):
                legacy = {}
                atm_strike = min(strikes, key=lambda value: abs(value - spot))
                atm_row = strikes.index(atm_strike) + 5
                for row in range(atm_row, atm_row + 1):
                    legacy[(row, "call")] = (255, 255, 0)
                    legacy[(row, "put")] = (255, 255, 0)
                for index, strike in enumerate(strikes):
                    row = index + 5
                    if strike < spot:
                        legacy[(row, "call")] = (198, 224, 180)
                    if strike > spot:
                        legacy[(row, "put")] = (255, 200, 200)

                batched = {}
                batched[(atm_row, "call")] = (255, 255, 0)
                batched[(atm_row, "put")] = (255, 255, 0)
                for cell_range, color in itm_color_ranges(self.block, list(strikes), spot):
                    start_row = int("".join(c for c in cell_range.split(":")[0] if c.isdigit()))
                    end_row = int("".join(c for c in cell_range.split(":")[1] if c.isdigit()))
                    side = "call" if cell_range.startswith(self.block["call_start"]) else "put"
                    for row in range(start_row, end_row + 1):
                        batched[(row, side)] = color
                self.assertEqual(batched, legacy)


class UsPerformanceMetricTests(unittest.TestCase):
    def test_aggregates_and_resets(self):
        metrics = PerformanceMetrics()
        metrics.record("spx", "excel_write", 0.01, rows=10)
        metrics.record("spx", "excel_write", 0.03, rows=20)
        summary = metrics.summary_lines(now=metrics._started + 30)
        self.assertIn(
            "performance.spx.excel_write: count=2, avg_ms=20.00, "
            "max_ms=30.00, rows=30",
            summary,
        )
        metrics.reset(now=metrics._started + 30)
        self.assertEqual(len(metrics.summary_lines(now=metrics._started + 30)), 1)

    def test_connection_and_traceback_logs_are_written(self):
        with tempfile.TemporaryDirectory() as temp:
            with patch.dict(os.environ, {"LOCALAPPDATA": temp}):
                connection = append_connection_event(
                    "spx", "connected", r"workbook=C:\test\spx.xlsx"
                )
                try:
                    raise RuntimeError("COM test failure")
                except RuntimeError as exc:
                    errors = append_app_exception("Excel update", exc)

            with open(connection, encoding="utf-8") as source:
                connection_text = source.read()
            with open(errors, encoding="utf-8") as source:
                error_text = source.read()

        self.assertIn("status=connected", connection_text)
        self.assertIn("RuntimeError: COM test failure", error_text)
        self.assertIn("Traceback (most recent call last)", error_text)


if __name__ == "__main__":
    unittest.main()
