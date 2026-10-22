import os
import tempfile
import unittest
from datetime import datetime

from excel_helpers import (
    find_book_and_app_in_apps,
    prefer_expiry_root,
    refresh_delay,
    reopen_due,
    should_restore_calculation,
    workbook_name_is_open_in_apps,
)
from us_time import eastern_now, filter_expiry_dates


class FakeBook:
    def __init__(self, fullname):
        self.fullname = fullname


class FakeApp:
    def __init__(self, books):
        self.books = books


class DjxRootSelectionTests(unittest.TestCase):
    def test_standard_root_does_not_gap_fill_with_zero_oi_weekly_rows(self):
        rows = [
            {"expiry": "2026-10-16", "root": "DJX", "strike": 510, "oi": 120},
            {"expiry": "2026-10-16", "root": "DJXW", "strike": 505, "oi": 0},
            {"expiry": "2026-10-16", "root": "DJXW", "strike": 510, "oi": 0},
            {"expiry": "2026-10-16", "root": "DJXW", "strike": 515, "oi": 0},
            {"expiry": "2026-10-23", "root": "DJXW", "strike": 510, "oi": 2},
        ]
        self.assertEqual(
            prefer_expiry_root(rows, "2026-10-16", "DJX"),
            [rows[0]],
        )

    def test_weekly_only_expiry_keeps_weekly_rows(self):
        rows = [
            {"expiry": "2026-10-19", "root": "DJXW", "strike": 510},
            {"expiry": "2026-10-19", "root": "DJXW", "strike": 515},
        ]
        self.assertEqual(prefer_expiry_root(rows, "2026-10-19", "DJX"), rows)

    def test_bigclawd_rows_with_base_root_are_unchanged(self):
        rows = [
            {"expiry": "2026-10-16", "root": "SPX", "strike": 6000},
            {"expiry": "2026-10-16", "root": "SPX", "strike": 6005},
        ]
        self.assertEqual(prefer_expiry_root(rows, "2026-10-16", "SPX"), rows)


class ExpiryTimezoneTests(unittest.TestCase):
    def test_eastern_cutoff_excludes_today_at_sixteen_only(self):
        expiries = ["2026-10-12", "2026-10-16"]
        before_cutoff = datetime(2026, 10, 12, 15, 59)
        after_cutoff = datetime(2026, 10, 12, 16, 0)
        self.assertEqual(
            filter_expiry_dates(expiries, before_cutoff, True),
            expiries,
        )
        self.assertEqual(
            filter_expiry_dates(expiries, after_cutoff, True),
            ["2026-10-16"],
        )

    def test_timezone_fallback_keeps_today_selectable(self):
        expiries = ["2026-10-12", "2026-10-16"]
        local_now = datetime(2026, 10, 12, 20, 0)
        self.assertEqual(
            filter_expiry_dates(expiries, local_now, False),
            expiries,
        )

    def test_timezone_failure_is_reported_and_returns_explicit_fallback(self):
        warnings = []
        local_now = datetime(2026, 10, 12, 20, 0)

        def missing_zone():
            raise ImportError("zone data missing")

        now, available = eastern_now(
            timezone_factory=missing_zone,
            now_factory=lambda *args: local_now,
            warning=warnings.append,
        )

        self.assertEqual(now, local_now)
        self.assertFalse(available)
        self.assertEqual(len(warnings), 1)


class WorkbookRecoveryTests(unittest.TestCase):
    def test_reopen_retries_no_more_than_every_twenty_seconds(self):
        self.assertTrue(reopen_due(100.0, 0.0))
        self.assertFalse(reopen_due(119.9, 100.0))
        self.assertTrue(reopen_due(120.0, 100.0))

    def test_workbook_is_found_across_multiple_excel_instances(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "djx_option_chain.xlsx")
            app1 = FakeApp([])
            app2 = FakeApp([FakeBook(path)])
            self.assertTrue(workbook_name_is_open_in_apps([app1, app2], path))
            book, owner = find_book_and_app_in_apps([app1, app2], path)
            self.assertIs(owner, app2)
            self.assertEqual(book.fullname, path)

    def test_uninspectable_excel_instance_prevents_duplicate_open(self):
        class BrokenApp:
            @property
            def books(self):
                raise RuntimeError("COM unavailable")

        errors = []
        result = workbook_name_is_open_in_apps(
            [BrokenApp()],
            "C:\\chains\\djx_option_chain.xlsx",
            lambda exc, context: errors.append((str(exc), context)),
        )
        self.assertTrue(result)
        self.assertEqual(errors[0][0], "COM unavailable")


class RefreshControlTests(unittest.TestCase):
    def test_refresh_uses_floor_control_and_render_cost_without_ratchet(self):
        self.assertEqual(refresh_delay(1.0, 3.0, 0.2), 3.0)
        self.assertEqual(refresh_delay(1.0, 0.5, 2.0), 2.0)
        self.assertEqual(refresh_delay(1.0, 0.5, 0.2), 1.0)

    def test_restore_only_when_calculation_remains_manual(self):
        self.assertTrue(should_restore_calculation(-4135, -4135))
        self.assertFalse(should_restore_calculation(-4105, -4135))


if __name__ == "__main__":
    unittest.main()
