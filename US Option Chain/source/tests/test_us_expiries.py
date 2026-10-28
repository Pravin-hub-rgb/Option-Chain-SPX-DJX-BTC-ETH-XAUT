"""Regression tests for the expiry list.

The customer reported that only the far-away monthly expiries were available.
The cause was not the data: the definition request for the weekly parent was
being answered with "504 The remote gateway timed out", so the weekly chain
never arrived and the weekly expiries never reached the drop-down.

The definition feed repeats an instrument's record every day it is live, so a
multi-day window multiplies the payload for no extra information. Asking for
one calendar day per request keeps every request under the gateway's limit.
"""
import unittest
import os
import sys
from datetime import date, datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

import databento_feed
import main
import us_excel


class DefinitionChunkTests(unittest.TestCase):
    def setUp(self):
        self.start = datetime(2026, 10, 19, 0, 0, 0, tzinfo=timezone.utc)
        self.end = datetime(2026, 10, 21, 19, 40, 0, tzinfo=timezone.utc)

    def test_splits_into_one_day_pieces(self):
        chunks = databento_feed._definition_chunks(self.start, self.end)
        self.assertEqual(len(chunks), 3)
        for chunk_start, chunk_end in chunks:
            self.assertLessEqual((chunk_end - chunk_start).total_seconds(), 86400)

    def test_covers_the_whole_range(self):
        chunks = databento_feed._definition_chunks(self.start, self.end)
        self.assertEqual(chunks[0][0], self.start)
        self.assertEqual(chunks[-1][1], self.end)
        for i in range(1, len(chunks)):
            self.assertEqual(chunks[i][0], chunks[i - 1][1])

    def test_a_single_day_is_one_chunk(self):
        start = datetime(2026, 10, 21, 0, 0, 0, tzinfo=timezone.utc)
        end = datetime(2026, 10, 21, 19, 40, 0, tzinfo=timezone.utc)
        chunks = databento_feed._definition_chunks(start, end)
        self.assertEqual(len(chunks), 1)

    def test_lookback_is_small_enough_to_survive_the_gateway(self):
        """A two day window in one request is what came back as a 504."""
        # Each chunk is at most one day, so the payload per request is the
        # same size that was measured working.
        chunks = databento_feed._definition_chunks(self.start, self.end)
        for chunk_start, chunk_end in chunks:
            self.assertLessEqual((chunk_end - chunk_start).days, 1)


class DropdownOrderTests(unittest.TestCase):
    """The drop-down is ordered by date so the near expiries are at the top."""

    def setUp(self):
        # Today is filtered out by the expiry rules, so use later dates.
        # Relative to date.today() so the suite never ages out of green.
        today = date.today()
        self.dates = [
            (today + timedelta(days=29)).isoformat(),
            (today + timedelta(days=1)).isoformat(),
            (today + timedelta(days=49)).isoformat(),
            (today + timedelta(days=5)).isoformat(),
            (today + timedelta(days=17)).isoformat(),
            (today + timedelta(days=3)).isoformat(),
        ]

    def test_by_date_is_nearest_first(self):
        ordered = main.available_expiries(self.dates, None, by_date=True)
        self.assertEqual(ordered, sorted(self.dates))
        self.assertEqual(ordered[0], (date.today() + timedelta(days=1)).isoformat())

    def test_default_order_is_nearest_first_too(self):
        """The three slots sit side by side, so they show three consecutive
        expiries: the dailies and weeklies next to each other, the way a
        trader watches them."""
        nearest = self.dates[1]
        main.expiry_roots["SPX"] = {
            self.dates[0]: {"SPX"},
            self.dates[2]: {"SPX"},
            nearest: {"SPXW"},
        }
        try:
            ordered = main.available_expiries(self.dates, "SPX")
            self.assertEqual(ordered, sorted(self.dates))
            self.assertEqual(ordered[0], nearest)
            self.assertEqual(ordered[1], self.dates[5])
            self.assertEqual(ordered[2], self.dates[3])
        finally:
            main.expiry_roots.pop("SPX", None)

    def test_dropdown_uses_date_order(self):
        import inspect
        src = inspect.getsource(us_excel.update_expiry_dropdowns)
        self.assertIn("by_date=True", src)


if __name__ == "__main__":
    unittest.main()
