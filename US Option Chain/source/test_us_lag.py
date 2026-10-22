"""Regression tests for the lag fixes.

Three separate things were making the sheet feel stuck:

1. force_usd_block_formats walked the open-interest column one cell at a
   time - 1206 COM round trips per block, about twelve seconds for a 600-row
   block. The chain already writes OI as "$123" text, so the per-cell walk was
   pure waste.
2. The ITM/ATM colouring cleared and repainted every cell of the block on
   every tick, dirtying thousands of cells and making Excel recalculate the
   customer's own formulas for no visible change.
3. The poll loop was one sequential pass over all three symbols behind a global
   two-second rate gate, so every symbol was about six seconds behind.

Also covered: a workbook with the same file name open from another folder must
be used rather than fought with, because Excel refuses to open two workbooks
with the same name.
"""
import os
import sys
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import excel_helpers
import us_performance


class FakeBook:
    def __init__(self, fullname):
        self.fullname = fullname


class FakeApp:
    def __init__(self, books):
        self.books = books


class SameNameWorkbookTests(unittest.TestCase):
    """A workbook with the same name open elsewhere must be found, not fought."""

    def test_finds_same_name_from_another_folder(self):
        target = r"C:\tool\spx_option_chain.xlsx"
        other = r"C:\somewhere\else\spx_option_chain.xlsx"
        apps = [FakeApp([FakeBook(other)])]
        book, app = excel_helpers.find_book_by_name_in_apps(apps, target)
        self.assertIsNotNone(book)
        self.assertEqual(book.fullname, other)

    def test_exact_path_still_wins(self):
        target = r"C:\tool\spx_option_chain.xlsx"
        apps = [FakeApp([FakeBook(target)])]
        book, _app = excel_helpers.find_book_and_app_in_apps(apps, target)
        self.assertIsNotNone(book)
        self.assertEqual(book.fullname, target)

    def test_different_name_is_not_matched(self):
        target = r"C:\tool\spx_option_chain.xlsx"
        apps = [FakeApp([FakeBook(r"C:\tool\djx_option_chain.xlsx")])]
        book, _app = excel_helpers.find_book_by_name_in_apps(apps, target)
        self.assertIsNone(book)

    def test_name_open_anywhere(self):
        target = r"C:\tool\spx_option_chain.xlsx"
        other = r"C:\other\spx_option_chain.xlsx"
        self.assertTrue(
            excel_helpers.workbook_name_open_anywhere_in_apps(
                [FakeApp([FakeBook(other)])], target
            )
        )
        self.assertFalse(
            excel_helpers.workbook_name_open_anywhere_in_apps(
                [FakeApp([FakeBook(r"C:\other\djx_option_chain.xlsx")])], target
            )
        )

    def test_unreadable_app_is_treated_as_open(self):
        """If Excel cannot be inspected, do not open a duplicate."""
        class BrokenApp:
            @property
            def books(self):
                raise RuntimeError("cannot enumerate")

        self.assertTrue(
            excel_helpers.workbook_name_open_anywhere_in_apps(
                [BrokenApp()], r"C:\tool\spx_option_chain.xlsx"
            )
        )


class ColorSignatureTests(unittest.TestCase):
    """The colouring must be repainted only when it actually changes."""

    BLOCK = {"start": 1, "atm": "A", "call_start": "A", "call_end": "D",
             "put_start": "F", "put_end": "I"}

    def _strikes(self, n=100, lo=7000.0):
        return [lo + 5.0 * i for i in range(n)]

    def test_same_spot_same_signature(self):
        strikes = self._strikes()
        a = us_performance.color_signature(strikes, 7100.0, len(strikes))
        b = us_performance.color_signature(strikes, 7100.0, len(strikes))
        self.assertEqual(a, b)

    def test_spot_crossing_a_strike_changes_signature(self):
        strikes = self._strikes()
        a = us_performance.color_signature(strikes, 7100.0, len(strikes))
        b = us_performance.color_signature(strikes, 7200.0, len(strikes))
        self.assertNotEqual(a, b)

    def test_spot_moving_inside_a_strike_keeps_signature(self):
        """A tick that does not cross a strike must not repaint the block.

        Strikes are 7000, 7005, 7010 ... 7100 and 7105 sit either side of
        7100.5, so 7100.4 and 7100.6 both keep the same ATM strike and the
        same ITM boundaries, and the block must not be repainted.
        """
        strikes = self._strikes()
        a = us_performance.color_signature(strikes, 7100.4, len(strikes))
        b = us_performance.color_signature(strikes, 7100.6, len(strikes))
        self.assertEqual(a, b)

    def test_spot_crossing_the_midpoint_moves_the_atm_row(self):
        """Crossing the midpoint between two strikes must repaint the block.

        The midpoint between 7100 and 7105 is 7102.5.
        """
        strikes = self._strikes()
        a = us_performance.color_signature(strikes, 7102.4, len(strikes))
        b = us_performance.color_signature(strikes, 7102.6, len(strikes))
        self.assertNotEqual(a, b)

    def test_row_count_change_changes_signature(self):
        strikes = self._strikes()
        a = us_performance.color_signature(strikes, 7100.0, len(strikes))
        b = us_performance.color_signature(strikes, 7100.0, len(strikes) + 1)
        self.assertNotEqual(a, b)

    def test_empty_strikes(self):
        sig = us_performance.color_signature([], 7100.0, 0)
        self.assertEqual(sig, (0, 0, 0, 0))


class SpotCacheTests(unittest.TestCase):
    """The spot must be served from a cache the background task keeps warm."""

    def setUp(self):
        import us_spot
        self.main = us_spot

    def test_tencent_parse_handles_batched_body(self):
        # The real Tencent body puts the feed timestamp at field 30.
        def line(key, price, stamp):
            fields = ["1", "2", "3", price, "7710.0"] + ["x"] * 25 + [stamp]
            return f'v_{key}="' + "~".join(fields) + '";'

        body = (
            line("usINX", "7713.90", "2026-10-21 10:13:40") + "\n"
            + line("usDJI", "52484.18", "2026-10-21 10:13:41")
        )
        parsed = self.main._tencent_parse(body)
        self.assertIn("usINX", parsed)
        self.assertIn("usDJI", parsed)
        self.assertAlmostEqual(parsed["usINX"][0], 7713.90, places=2)
        self.assertEqual(parsed["usDJI"][1], "2026-10-21 10:13:41")

    def test_parse_skips_malformed_lines(self):
        parsed = self.main._tencent_parse('garbage\nv_x="1~2~3~";\n')
        self.assertEqual(parsed, {})

    def test_dji_scale_is_applied(self):
        """Tencent quotes the full DJIA; the DJX index is 1/100 of that."""
        self.assertEqual(self.main.TENCENT_SCALE.get("usDJI"), 0.01)

    def test_cache_serves_without_network(self):
        with self.main._spot_lock:
            self.main._spot_cache["SPX"] = (
                7000.0, 0.0, "yahoo", time.monotonic(), time.time()
            )
        try:
            self.assertAlmostEqual(self.main.spot_for_symbol("SPX"), 7000.0, places=2)
            self.assertAlmostEqual(self.main.tencent_spot("SPX"), 7000.0, places=2)
        finally:
            with self.main._spot_lock:
                self.main._spot_cache.pop("SPX", None)

    def test_stale_cached_quote_is_still_served(self):
        """An old quote is the last known level - better than zero.

        It is flagged in the header as QUOTE Ns OLD, so the trader can see it
        is not live, rather than seeing a blank or a zero.
        """
        with self.main._spot_lock:
            self.main._spot_cache["SPX"] = (
                7000.0, 9999.0, "tencent", time.monotonic(), time.time() - 9999
            )
        try:
            self.assertAlmostEqual(
                self.main.spot_for_symbol("SPX"), 7000.0, places=2
            )
        finally:
            with self.main._spot_lock:
                self.main._spot_cache.pop("SPX", None)

    def test_expired_cache_entry_is_rejected(self):
        """Even a fresh quote is dropped once the cache entry itself ages out."""
        with self.main._spot_lock:
            self.main._spot_cache["SPX"] = (7000.0, 0.0, "yahoo", 0.0, time.time())
        try:
            self.assertIsNone(self.main.spot_for_symbol("SPX"))
        finally:
            with self.main._spot_lock:
                self.main._spot_cache.pop("SPX", None)

    def test_newest_quote_wins(self):
        """The quote with the latest timestamp is the one kept.

        The age a feed reports is how far behind the wall clock it is, so a
        smaller age is a more recent quote. Keeping the smallest age is
        keeping the newest quote.
        """
        with self.main._spot_lock:
            self.main._spot_cache.clear()
        try:
            self.main._store_spot("SPX", 100.0, 30.0, "tencent")
            self.main._store_spot("SPX", 200.0, 2.0, "yahoo")
            self.assertAlmostEqual(self.main.spot_for_symbol("SPX"), 200.0, places=2)
            # An older quote must not replace the newer one.
            self.main._store_spot("SPX", 300.0, 60.0, "tencent")
            self.assertAlmostEqual(self.main.spot_for_symbol("SPX"), 200.0, places=2)
        finally:
            with self.main._spot_lock:
                self.main._spot_cache.clear()

    def test_very_old_quote_never_replaces_a_newer_one(self):
        """A fifteen minute old feed must not drag the index level back.

        Yahoo's ^SPX, ^DJX, ^RUT and ^VIX feeds are stale. When the live
        feed fails and the stale one answers, letting it replace the live
        quote would make the spot jump backwards, which is the "spot goes back
        and forth" complaint.
        """
        with self.main._spot_lock:
            self.main._spot_cache.clear()
        try:
            self.main._store_spot("SPX", 7100.0, 1.0, "yahoo")
            self.main._store_spot("SPX", 7000.0, 900.0, "yahoo")
            self.assertAlmostEqual(
                self.main.spot_for_symbol("SPX"), 7100.0, places=2
            )
        finally:
            with self.main._spot_lock:
                self.main._spot_cache.clear()


class PollCadenceTests(unittest.TestCase):
    """The poll loop must not serialise the symbols behind a global gate."""

    def setUp(self):
        import us_chain
        self.main = us_chain

    def test_source_min_interval_is_small(self):
        self.assertLessEqual(self.main.SOURCE_MIN_INTERVAL, 0.5)

    def test_symbol_poll_period_is_one_second_or_less(self):
        self.assertLessEqual(self.main.SYMBOL_POLL_SECONDS, 1.0)

    def test_throttle_is_per_symbol_not_global(self):
        """Two different symbols must not queue behind one another."""
        import inspect
        src = inspect.getsource(self.main.throttled)
        self.assertIn("_last_symbol_call", src)
        self.assertNotIn("POLL_GAP - now", src)


class UsdFormatTests(unittest.TestCase):
    """The per-cell OI walk must not come back."""

    def test_no_per_cell_oi_loop(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel.force_usd_block_formats)
        self.assertNotIn("for r in range(5, end + 1)", src)
        self.assertIn("rng.value", src)


class ContiguousRunsTests(unittest.TestCase):
    """Only the rows that moved may be written back."""

    def setUp(self):
        import us_excel
        self.main = us_excel

    def test_single_row(self):
        self.assertEqual(self.main._contiguous_runs([4]), [(4, 4)])

    def test_contiguous_run(self):
        self.assertEqual(self.main._contiguous_runs([1, 2, 3]), [(1, 3)])

    def test_separate_runs(self):
        self.assertEqual(
            self.main._contiguous_runs([0, 1, 5, 9, 10]),
            [(0, 1), (5, 5), (9, 10)],
        )

    def test_empty(self):
        self.assertEqual(self.main._contiguous_runs([]), [])


class TencentStalenessTests(unittest.TestCase):
    """A stale index quote must not silently drive the ATM row."""

    def setUp(self):
        import us_spot
        self.main = us_spot

    def test_recent_stamp_is_fresh(self):
        from us_time import eastern_now
        eastern, _ok = eastern_now()
        stamp = eastern.strftime("%Y-%m-%d %H:%M:%S")
        age = self.main._tencent_feed_age_seconds(stamp)
        self.assertIsNotNone(age)
        self.assertLess(age, 30)

    def test_yesterday_stamp_is_old(self):
        from datetime import timedelta
        from us_time import eastern_now
        eastern, _ok = eastern_now()
        stamp = (eastern - timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
        age = self.main._tencent_feed_age_seconds(stamp)
        self.assertIsNotNone(age)
        self.assertGreater(age, 3600)

    def test_unparsable_stamp_is_unknown(self):
        self.assertIsNone(self.main._tencent_feed_age_seconds("not a stamp"))
        self.assertIsNone(self.main._tencent_feed_age_seconds(""))

    def test_stale_threshold_is_sane(self):
        self.assertGreater(self.main.SPOT_STALE_SECONDS, 30)
        self.assertLess(self.main.SPOT_STALE_SECONDS, 600)


class TitleBarThrottleTests(unittest.TestCase):
    """Rewriting the window caption repaints the window; do it sparingly."""

    def setUp(self):
        import us_excel
        self.main = us_excel

    def test_window_status_stores_a_tuple(self):
        """The stored value is (text, timestamp), not bare text."""
        import inspect
        src = inspect.getsource(self.main.show_status)
        self.assertIn("_last_window_status[key] = (title, now)", src)

    def test_throttle_window_is_used(self):
        import inspect
        src = inspect.getsource(self.main.show_status)
        self.assertIn("now - last_at < 5.0", src)


class PerCycleWriteGatingTests(unittest.TestCase):
    """Nothing may dirty cells on every refresh unless it has to."""

    def test_oi_number_format_is_gated(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn("_oi_fmt_done", src)

    def test_user_date_format_is_gated(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel.update_user_dates)
        self.assertIn("_user_fmt_done", src)

    def test_row_level_diff_is_used(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn("_block_values", src)
        self.assertIn("_contiguous_runs", src)

    def test_scattered_changes_use_one_com_call(self):
        """Writing each changed row separately is far slower than one call.

        Measured on this machine: 12 scattered rows take 1.2ms in one call and
        436ms in twelve, because every call into Excel costs about 33ms no
        matter how few cells it carries. That is what freezes Excel while the
        customer scrolls, so scattered changes must fall back to one write.
        """
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._flush_block_writes)
        self.assertNotIn("for start, end in _contiguous_runs(changed):", src)
        self.assertIn("len(runs) == 1", src)

    def test_blocks_are_not_batched_into_one_wide_write(self):
        """The blocks must not be written as one wide range.

        It was cheaper in COM calls, but it painted over the blocks that had
        not changed and left them empty.
        """
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._flush_block_writes)
        self.assertNotIn("grid = [[None] * width", src)
        self.assertNotIn("(4 + max_rows, width)", src)

    def test_header_row_is_read_once_per_sheet(self):
        """Six one-cell header reads per sheet became one row read."""
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn("sheet.range((1, 1), (4, 31)).value", src)
        # No per-cell header comparison left.
        self.assertNotIn("sheet.range((1, col_l)).value !=", src)
        self.assertNotIn("sheet.range((1, col_l + 4)).value !=", src)


class AutoOpenLatchTests(unittest.TestCase):
    """A workbook the customer closes must not pop back up.

    The tool used to call os.startfile on a timer whenever it could not find
    its workbook. A customer who closed the workbook - or closed Excel - watched
    it reappear every twenty seconds, and each attempt could put Excel's
    "cannot open two workbooks with the same name" dialog on screen.
    """

    def setUp(self):
        import us_excel
        self.main = us_excel

    def test_latch_exists_per_role(self):
        self.assertEqual(
            sorted(self.main._auto_open_done.keys()),
            sorted(self.main.ROLE_ORDER),
        )

    def test_reopen_is_gated_on_the_latch(self):
        import inspect
        src = inspect.getsource(self.main._write_to_excel)
        self.assertIn("not _auto_open_done.get(role)", src)

    def test_startup_open_sets_the_latch(self):
        import inspect
        import us_workbook
        src = inspect.getsource(us_workbook.open_excel_files)
        self.assertIn("_auto_open_done[role] = True", src)

    def test_only_two_places_launch_excel(self):
        """Nothing else may open a workbook behind the customer's back."""
        import inspect
        import main
        import us_workbook
        import us_excel
        self.assertEqual(
            inspect.getsource(main).count("os.startfile(")
            + inspect.getsource(us_workbook).count("os.startfile(")
            + inspect.getsource(us_excel).count("os.startfile("),
            2,
        )

    def test_latch_is_rearmed_only_when_the_file_is_gone(self):
        import inspect
        src = inspect.getsource(self.main._write_to_excel)
        self.assertIn("if not os.path.exists(path):", src)
        self.assertIn("_auto_open_done[role] = False", src)


class SheetStructureGuardTests(unittest.TestCase):
    """A restructured sheet must stop the tool, not be written over."""

    class FakeSheet:
        def __init__(self, a4):
            self._a4 = a4

        def range(self, *args):
            class R:
                def __init__(self, v):
                    self.value = v
            return R(self._a4)

    def setUp(self):
        import us_excel
        self.main = us_excel

    def test_intact_sheet_is_ok(self):
        sheet = self.FakeSheet("Call LTP")
        self.assertTrue(self.main.sheet_structure_ok(sheet, "spx"))

    def test_blank_sheet_is_ok(self):
        sheet = self.FakeSheet(None)
        self.assertTrue(self.main.sheet_structure_ok(sheet, "spx"))

    def test_inserted_column_is_detected(self):
        sheet = self.FakeSheet("My Notes")
        self.assertFalse(self.main.sheet_structure_ok(sheet, "spx"))

    def test_unreadable_sheet_is_allowed(self):
        class Broken:
            def range(self, *a):
                raise RuntimeError("COM dead")
        self.assertTrue(self.main.sheet_structure_ok(Broken(), "spx"))

    def test_guard_is_in_the_write_path(self):
        import inspect
        src = inspect.getsource(self.main._write_to_excel)
        self.assertIn("sheet_structure_ok", src)


class ParitySpotTests(unittest.TestCase):
    """Put-call parity from our own chain: a real-time index level that needs
    no external feed, so the spot never freezes when a free endpoint dies."""

    def setUp(self):
        import us_spot
        self.main = us_spot

    def _rows(self, spot, call_bid, call_ask, put_bid, put_ask, strike=7750.0,
              expiry=None):
        return [
            {"expiry": expiry or self.main.date.today().isoformat(),
             "strike": strike, "type": "call",
             "bid": call_bid, "ask": call_ask, "ltp": 0.0, "oi": 100},
            {"expiry": expiry or self.main.date.today().isoformat(),
             "strike": strike, "type": "put",
             "bid": put_bid, "ask": put_ask, "ltp": 0.0, "oi": 100},
        ]

    def test_parity_reproduces_the_index_level(self):
        # C - P = S - K  =>  S = K + (C - P). Same-day expiry, no discount,
        # and both mids equal, so the index level is the strike itself.
        rows = self._rows(7750.0, 30.0, 32.0, 30.0, 32.0, strike=7750.0)
        parity = self.main._parity_spot(rows, 7750.0)
        self.assertIsNotNone(parity)
        self.assertAlmostEqual(parity, 7750.0, delta=1.5)

    def test_parity_needs_both_sides(self):
        rows = [{"expiry": "2026-12-18", "strike": 7750.0, "type": "call",
                 "bid": 10.0, "ask": 12.0, "ltp": 0.0, "oi": 100}]
        self.assertIsNone(self.main._parity_spot(rows, 7750.0))

    def test_parity_ignores_a_missing_quote(self):
        rows = self._rows(7750.0, 0.0, 0.0, 30.0, 32.0)
        self.assertIsNone(self.main._parity_spot(rows, 7750.0))

    def test_parity_of_no_rows(self):
        self.assertIsNone(self.main._parity_spot([], 7750.0))

    def test_parity_uses_only_the_front_expiry(self):
        today = self.main.date.today().isoformat()
        rows = self._rows(7750.0, 30.0, 32.0, 30.0, 32.0, expiry=today)
        # A much later expiry with a completely different price. If it were
        # used the answer would be far away from 7750.
        rows += self._rows(7750.0, 900.0, 1000.0, 1.0, 2.0,
                           expiry="2027-01-16")
        parity = self.main._parity_spot(rows, 7750.0)
        self.assertAlmostEqual(parity, 7750.0, delta=1.5)

    def test_header_stays_clean(self):
        """The header shows LIVE / STALE / MISMATCH, nothing else.

        A clutter of numbers in the header reads as noise to a trader, so the
        quote age lives in the log instead.
        """
        import inspect
        import us_excel
        src = inspect.getsource(us_excel.source_delay_note)
        self.assertNotIn("quote {", src)
        self.assertIn("STALE", src)
        self.assertIn("SPOT MISMATCH", src)

    def test_thin_book_is_not_used_for_parity(self):
        """Near expiry the quotes go thin. A wide spread is not an index level."""
        today = self.main.date.today().isoformat()
        # call 30/90 - a 60 point wide book
        rows = self._rows(7750.0, 30.0, 90.0, 30.0, 32.0, expiry=today)
        self.assertIsNone(self.main._parity_spot(rows, 7750.0))

    def test_ill_formed_book_is_rejected(self):
        """A book with the bid above the ask is nonsense."""
        today = self.main.date.today().isoformat()
        rows = self._rows(7750.0, 32.0, 30.0, 30.0, 32.0, expiry=today)
        self.assertIsNone(self.main._parity_spot(rows, 7750.0))

    def test_market_closed_missing_quotes_gives_no_parity(self):
        """With no quotes there is no parity, so no garbage spot can appear."""
        rows = self._rows(7750.0, 0.0, 0.0, 0.0, 0.0)
        self.assertIsNone(self.main._parity_spot(rows, 7750.0))

    def test_sanity_limit_rejects_a_wildly_wrong_parity(self):
        """A parity nowhere near the external quote is thrown away."""
        import inspect
        import us_chain
        src = inspect.getsource(us_chain.load_symbol)
        self.assertIn("spot_parity_rejected", src)
        self.assertIn("PARITY_SANITY_PCT", src)


class BlockWriteSafetyTests(unittest.TestCase):
    """A block that did not change must never be written over.

    Writing all three blocks of a sheet as one wide range was cheaper in COM
    calls, but the combined grid carried blanks for the blocks that had not
    changed, so an unchanged block came back empty and the sheet looked like
    it was blanking out and then filling in again.
    """

    def setUp(self):
        import us_excel
        self.main = us_excel

    def test_blocks_are_written_separately(self):
        import inspect
        src = inspect.getsource(self.main._flush_block_writes)
        # No combined wide range.
        self.assertNotIn("sheet.range((5, 1), (4 + max_rows, width))", src)
        self.assertNotIn("excel_blocks_batched", src)
        # Each block is written on its own.
        self.assertIn("for bi, block, chain, n_rows, runs, prev in pending:", src)
        self.assertIn("block[\"start\"]", src)

    def test_a_failed_block_write_does_not_stop_the_others(self):
        import inspect
        src = inspect.getsource(self.main._flush_block_writes)
        self.assertIn("except Exception as e:", src)
        self.assertIn("writing chain block", src)


class LiveApiContractTests(unittest.TestCase):
    """The live API prescribes how a session must be kept healthy.

    From the Databento live API documentation:
      - "When reconnecting to the gateway, clients should resubscribe to all
        desired symbols." A reconnect restores the TCP session but not the
        subscription, so the feed goes silent after any network blip unless
        the symbols are asked for again.
      - "if a client does not receive any messages from the gateway for the
        duration of one heartbeat interval plus ten seconds, the session can
        be considered hung. Clients are instructed to disconnect from the
        gateway and reconnect upon detecting hung connections."
    """

    def setUp(self):
        import databento_feed
        self.feed_module = databento_feed

    def test_reconnect_resubscribes(self):
        import inspect
        src = inspect.getsource(self.feed_module.DatabentoFeed._on_reconnect)
        self.assertIn("threading.Thread", src)
        self.assertIn("_resubscribe_after_reconnect", src)
        worker = inspect.getsource(
            self.feed_module.DatabentoFeed._resubscribe_after_reconnect
        )
        self.assertIn("self._subscribe(live, target)", worker)
        self.assertIn("self._sub_lock", worker)

    def test_hung_session_is_dropped(self):
        import inspect
        src = inspect.getsource(self.feed_module.DatabentoFeed._watch_hung_connection)
        self.assertIn("live.terminate()", src)
        self.assertIn("last_message_age", src)

    def test_hung_threshold_is_heartbeat_plus_ten(self):
        self.assertEqual(
            self.feed_module.HUNG_AFTER_SECONDS,
            self.feed_module.HEARTBEAT_INTERVAL_S + 10,
        )

    def test_heartbeat_counts_as_liveness(self):
        """A heartbeat proves the session is alive even with no data."""
        import inspect
        src = inspect.getsource(self.feed_module.DatabentoFeed._on_record)
        self.assertIn("_last_message", src)

    def test_reconnect_policy_is_enabled(self):
        import inspect
        src = inspect.getsource(self.feed_module.DatabentoFeed._open_session)
        self.assertIn('reconnect_policy="reconnect"', src)

    def test_ts_out_is_requested(self):
        """ts_out gives the gateway send time, so latency is measurable."""
        import inspect
        src = inspect.getsource(self.feed_module.DatabentoFeed._open_session)
        self.assertIn("ts_out=True", src)

    def test_subscriptions_are_chunked_not_per_symbol(self):
        """The gateway throttles subscription requests, so they are batched."""
        import inspect
        src = inspect.getsource(self.feed_module.DatabentoFeed._subscribe)
        self.assertIn("symbols=chunk", src)
        self.assertIn("for start in range(0, len(missing), SUBSCRIBE_CHUNK)", src)


class SubscriptionMarginTests(unittest.TestCase):
    """Subscribe wider than the sheet shows.

    The gateway throttles subscription requests to ten a second, and each
    request carries a chunk of symbols. If the subscribed band were exactly
    the displayed band, every strike the index moved would ask for another
    symbol and the request rate would creep towards the limit.
    """

    def setUp(self):
        import databento_feed
        self.feed_module = databento_feed

    def test_margin_is_wider_than_the_display_band(self):
        self.assertGreater(self.feed_module.SUBSCRIBE_MARGIN, 0)

    def test_target_symbols_uses_the_margin(self):
        import inspect
        src = inspect.getsource(self.feed_module.DatabentoFeed._target_symbols)
        self.assertIn("SUBSCRIBE_MARGIN", src)

    def test_speed_summary_is_logged(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._log_speed_summary)
        self.assertIn('"speed"', src)
        self.assertIn("cycle_ms", src)
        self.assertIn("recalc_ms", src)
        self.assertIn("spot_quote_age_s", src)



class CalcFlapGuardTests(unittest.TestCase):
    """Session-owned manual calculation must not flap.

    The watchdog used to fight our own write window, and the already-manual
    branch used to drop the ownership flag every cycle, so the mode flipped
    automatic / manual every couple of seconds and Excel recalculated on
    nearly every visible move.
    """

    def setUp(self):
        import us_state

        self.session = dict(us_state._calc_session)
        self.app_cache = us_state._app_cache
        self.manual_since = us_state._manual_since

    def tearDown(self):
        import us_state

        us_state._calc_session.clear()
        us_state._calc_session.update(self.session)
        us_state._app_cache = self.app_cache
        us_state._manual_since = self.manual_since

    def test_watchdog_steps_aside_while_we_own_manual(self):
        import us_state
        import us_excel

        class DeadApp:
            @property
            def version(self):
                raise AssertionError("watchdog must not touch Excel")

        app = DeadApp()
        us_state._calc_session["owned"] = True
        us_state._calc_session["failed"] = False
        us_excel.ensure_calculation_not_stranded(app)
        self.assertEqual(us_state._manual_since, 0.0)

    def test_watchdog_fires_outside_our_session(self):
        import us_state
        import us_excel
        import time

        class FlippableApp:
            def __init__(self):
                self.mode = -4135

            @property
            def version(self):
                return "16.0"

            @property
            def api(self):
                probe = self
                class Api:
                    @property
                    def Calculation(self):
                        return probe.mode

                    @Calculation.setter
                    def Calculation(self, value):
                        probe.mode = value
                return Api()

            @property
            def _calc_value(self):
                return self.mode

            @property
            def Calculation(self):
                return self.mode

        app = FlippableApp()
        us_state._calc_session.pop("owned", None)
        us_state._manual_since = time.time() - (us_state.MANUAL_WATCHDOG_SECONDS + 5)
        us_excel.ensure_calculation_not_stranded(app)
        self.assertEqual(app.mode, -4105)
        self.assertEqual(us_state._manual_since, 0.0)

    def test_already_manual_branch_keeps_the_ownership(self):
        import us_state
        import us_excel

        class ManualApp:
            def __init__(self):
                self._screen = True

            @property
            def version(self):
                return "16.0"

            @property
            def screen_updating(self):
                return self._screen

            @screen_updating.setter
            def screen_updating(self, value):
                self._screen = value

            @property
            def api(self):
                probe = self
                class Api:
                    @property
                    def Calculation(self):
                        return -4135
                return Api()

        app = ManualApp()
        us_state._calc_session["owned"] = True
        us_state._calc_session["failed"] = False
        us_excel._begin_excel_bulk_mode(app)
        self.assertTrue(us_state._calc_session.get("owned"))

    def test_already_manual_branch_disowns_a_foreign_manual(self):
        import us_state
        import us_excel

        class ManualApp:
            def __init__(self):
                self._screen = True

            @property
            def version(self):
                return "16.0"

            @property
            def screen_updating(self):
                return self._screen

            @screen_updating.setter
            def screen_updating(self, value):
                self._screen = value

            @property
            def api(self):
                probe = self
                class Api:
                    @property
                    def Calculation(self):
                        return -4135
                return Api()

        app = ManualApp()
        us_state._calc_session["owned"] = False
        us_excel._begin_excel_bulk_mode(app)
        self.assertFalse(us_state._calc_session.get("owned"))



class ComTaxTests(unittest.TestCase):
    """Per-cycle COM chatter that does not change what the trader sees.

    Each Excel get/set is a round trip of tens of milliseconds, so the
    connection detail read, the calculation-state probe, and the tasklist
    check must all be asked far less often than once per cycle.
    """

    def setUp(self):
        import us_excel
        self.main = us_excel

    def test_merged_block_read_replaces_per_cell_reads(self):
        import inspect

        src = inspect.getsource(self.main._write_to_excel)
        self.assertIn("sheet.range((1, 1), (4, 31)).value", src)
        self.assertNotIn('v = sheet.range("A2").value', src)

    def test_connection_detail_is_read_only_on_change(self):
        import inspect

        src = inspect.getsource(self.main._write_to_excel)
        self.assertIn("_excel_connection_status.get(role) != \"connected\"", src)

    def test_tasklist_cache_is_two_s(self):
        import inspect

        sig = inspect.signature(self.main.excel_process_running)
        self.assertEqual(sig.parameters["max_age"].default, 2.0)

    def test_watchdog_runs_once_per_cycle(self):
        import inspect

        self.assertIn(
            "ensure_calculation_not_stranded",
            inspect.getsource(self.main.render_cycle),
        )
        self.assertNotIn(
            "ensure_calculation_not_stranded",
            inspect.getsource(self.main.display_loop),
        )

    def test_calc_state_probe_is_gated_to_thirty_s(self):
        import inspect

        src = inspect.getsource(self.main._probe_calc_state_after_restore)
        self.assertIn("30.0", src)
        self.assertIn("us_state._calc_state_probe", src)



class UnprotectedMergeTests(unittest.TestCase):
    """The unprotected window writes one merged span, not many small ones."""


    def setUp(self):
        import us_excel
        self.main = us_excel

    def test_scattered_runs_become_one_span_write(self):
        import inspect

        src = inspect.getsource(self.main._flush_block_writes)
        self.assertIn("merged_rows", src)
        self.assertNotIn("for start, end in runs:", src)
        self.assertIn("excel_rows_written", src)

    def test_unchanged_rows_carry_the_previous_values(self):
        import inspect

        src = inspect.getsource(self.main._flush_block_writes)
        self.assertIn("list(prev[i])", src)



class FeedPollerLagTests(unittest.TestCase):
    """Feed calls in the hot path must not stall the pollers."""


    def setUp(self):
        import databento_feed
        self.feed_module = databento_feed

    def test_universe_wait_budget_is_five_seconds(self):
        import inspect

        src = inspect.getsource(self.feed_module.DatabentoFeed.ensure_universe)
        self.assertIn("time.monotonic() + 5.0", src)
        self.assertNotIn("time.monotonic() + 30.0", src)

    def test_subscribe_log_only_when_symbols_added(self):
        import inspect

        src = inspect.getsource(self.feed_module.DatabentoFeed._subscribe)
        self.assertIn("if not missing:", src)

    def test_snapshot_prunes_quotes_to_target(self):
        import inspect

        src = inspect.getsource(self.feed_module.DatabentoFeed.snapshot)
        self.assertIn("not in self._target", src)

    def test_on_record_takes_the_lock_once_per_cbbo_record(self):
        import inspect

        src = inspect.getsource(self.feed_module.DatabentoFeed._on_record)
        self.assertEqual(src.count("with self._lock:"), 3)

    def test_chain_poller_stops_blocking_after_first_sync(self):
        import inspect
        import us_chain

        src = inspect.getsource(us_chain._fetch_databento)
        self.assertIn("wait=sym not in _synced_symbols", src)
        self.assertIn("_synced_symbols.add(sym)", src)

    def test_reconnect_resubscribe_runs_off_the_callback_thread(self):
        import inspect

        src = inspect.getsource(self.feed_module.DatabentoFeed._on_reconnect)
        self.assertIn("threading.Thread", src)



class CoexistenceGuardTests(unittest.TestCase):
    """The watchdog must not rescue manual while the sibling tool still holds it."""


    def setUp(self):
        import us_excel
        import us_state

        self.main = us_excel
        self.session = dict(us_state._calc_session)
        self.app_cache = us_state._app_cache
        self.manual_since = us_state._manual_since

    def tearDown(self):
        import us_state

        us_state._calc_session.clear()
        us_state._calc_session.update(self.session)
        us_state._app_cache = self.app_cache
        us_state._manual_since = self.manual_since

    def _flippable_app(self):
        class FlippableApp:
            def __init__(self):
                self.mode = -4135

            @property
            def api(self):
                probe = self

                class Api:
                    @property
                    def Calculation(self):
                        return probe.mode

                    @Calculation.setter
                    def Calculation(self, value):
                        probe.mode = value

                return Api()

        return FlippableApp()

    def test_watchdog_restores_when_no_sibling_holds(self):
        import time
        import us_state
        from unittest import mock

        app = self._flippable_app()
        us_state._calc_session.pop("owned", None)
        us_state._manual_since = time.time() - (us_state.MANUAL_WATCHDOG_SECONDS + 5)
        with mock.patch.object(
            self.main.calc_mode_guard, "other_holder_alive", return_value=False
        ):
            self.main.ensure_calculation_not_stranded(app)
        self.assertEqual(app.mode, -4105)
        self.assertEqual(us_state._manual_since, 0.0)

    def test_watchdog_leaves_manual_while_sibling_holds(self):
        import time
        import us_state
        from unittest import mock

        app = self._flippable_app()
        us_state._calc_session.pop("owned", None)
        manual = time.time() - (us_state.MANUAL_WATCHDOG_SECONDS + 5)
        us_state._manual_since = manual
        with mock.patch.object(
            self.main.calc_mode_guard, "other_holder_alive", return_value=True
        ):
            self.main.ensure_calculation_not_stranded(app)
        self.assertEqual(app.mode, -4135)
        self.assertEqual(us_state._manual_since, manual)

    def test_holder_check_failure_fails_open(self):
        import time
        import us_state
        from unittest import mock

        app = self._flippable_app()
        us_state._calc_session.pop("owned", None)
        us_state._manual_since = time.time() - (us_state.MANUAL_WATCHDOG_SECONDS + 5)
        with mock.patch.object(
            self.main.calc_mode_guard,
            "other_holder_alive",
            side_effect=RuntimeError("boom"),
        ):
            self.main.ensure_calculation_not_stranded(app)
        self.assertEqual(app.mode, -4105)

    def test_skip_log_is_throttled(self):
        import inspect

        src = inspect.getsource(self.main._other_chain_tool_holds)
        self.assertIn("calc_watchdog_skip_holder", src)
        self.assertIn("60.0", src)



class FlapCounterTests(unittest.TestCase):
    """Observed calculation-mode flips are counted without extra COM calls."""


    _NAMES = (
        "_last_observed_calc_raw",
        "_calc_mode_flap_count",
        "_calc_flap_summary_after",
    )

    def setUp(self):
        import us_excel

        self.main = us_excel
        self.saved = {n: getattr(us_excel, n) for n in self._NAMES}

    def tearDown(self):
        for name, value in self.saved.items():
            setattr(self.main, name, value)

    def test_transition_counts_one_flap(self):
        self.main._last_observed_calc_raw = -4105
        before = self.main._calc_mode_flap_count
        self.main._observe_raw_calc(-4135)
        self.assertEqual(self.main._calc_mode_flap_count, before + 1)

    def test_steady_mode_counts_nothing(self):
        self.main._last_observed_calc_raw = -4135
        before = self.main._calc_mode_flap_count
        self.main._observe_raw_calc(-4135)
        self.assertEqual(self.main._calc_mode_flap_count, before)

    def test_none_resets_the_baseline(self):
        self.main._last_observed_calc_raw = -4105
        before = self.main._calc_mode_flap_count
        self.main._observe_raw_calc(None)
        self.assertIsNone(self.main._last_observed_calc_raw)
        self.main._observe_raw_calc(-4135)
        self.assertEqual(self.main._calc_mode_flap_count, before)

    def test_counter_reads_the_existing_probe(self):
        import inspect

        self.assertIn(
            "_observe_raw_calc(raw)",
            inspect.getsource(self.main._begin_excel_bulk_mode),
        )

    def test_summary_is_throttled_to_thirty_seconds(self):
        import inspect

        self.assertEqual(self.main._CALC_FLAP_SUMMARY_SECONDS, 30.0)
        src = inspect.getsource(self.main._observe_raw_calc)
        self.assertIn("calc_mode_flap_summary", src)


if __name__ == "__main__":
    unittest.main()
