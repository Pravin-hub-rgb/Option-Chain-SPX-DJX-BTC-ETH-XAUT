"""Regression tests for the market-closed path.

Outside trading hours the exchange stops sending quotes for the contracts the
tool subscribes to. The chain must still render (strikes and open interest
come from Databento and are correct) instead of coming up blank, and the
status line must say so plainly.

A previous version raised OverflowError here, because the record age is
infinity when no record has ever been received, and int(infinity) is not a
number. That took down every render pass and left the sheet empty.
"""
import math
import threading
import unittest

import databento_feed


class _StubFeed(databento_feed.DatabentoFeed):
    """DatabentoFeed without touching the network."""

    def __init__(self, started=True, connected=True, age=0.0):
        self._started_flag = started
        self._connected_flag = connected
        self._age = age
        self._lock = threading.Lock()
        self._quotes = {}
        self.band = 120
        self._defs = {"SPX": [{"r": "SPX", "root": "SPX", "e": "2026-10-23", "k": 7700.0, "t": "call", "oi": 5}]}

    def is_started(self):
        return self._started_flag

    def connected(self):
        return self._connected_flag

    def last_record_age(self):
        return self._age


class StatusTextTests(unittest.TestCase):
    def test_infinite_age_does_not_raise(self):
        feed = _StubFeed(age=float("inf"))
        text = feed.status_text()
        self.assertIsInstance(text, str)
        # Outside trading hours it says there are no live quotes; inside them
        # it says the feed is stale. Either way it must not blow up on an
        # infinite age, which is what a never-received record reports.
        self.assertTrue(
            "NO LIVE QUOTES" in text or "DATABENTO STALE" in text, text
        )
        self.assertIn("no data yet", text)

    def test_large_finite_age_is_formatted(self):
        feed = _StubFeed(age=4321.0)
        text = feed.status_text()
        self.assertIn("4321s", text)

    def test_fresh_feed_is_live(self):
        feed = _StubFeed(age=0.5)
        self.assertEqual(feed.status_text(), "LIVE")

    def test_not_started_reports_connecting(self):
        feed = _StubFeed(started=False, age=float("inf"))
        self.assertIn("DATABENTO", feed.status_text())


class ChainStillRendersTests(unittest.TestCase):
    def test_rows_survive_without_any_quotes(self):
        """A contract with no quote still yields a row (strike + OI intact)."""
        feed = _StubFeed(age=float("inf"))
        rows = feed._build_rows(
            "SPX",
            [
                {"r": "SPX   261023C07700000", "root": "SPX", "e": "2026-10-23",
                 "k": 7700.0, "t": "call", "oi": 42},
                {"r": "SPX   261023P07700000", "root": "SPX", "e": "2026-10-23",
                 "k": 7700.0, "t": "put", "oi": 7},
            ],
            7700.0,
            ["2026-10-23"],
        )
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["strike"], 7700.0)
        self.assertEqual(rows[0]["oi"], 42)
        self.assertEqual(rows[0]["bid"], 0.0)
        self.assertEqual(rows[1]["oi"], 7)


class MathGuardTests(unittest.TestCase):
    def test_isfinite_helper_behaviour(self):
        self.assertTrue(math.isfinite(0.0))
        self.assertFalse(math.isfinite(float("inf")))


if __name__ == "__main__":
    unittest.main()
