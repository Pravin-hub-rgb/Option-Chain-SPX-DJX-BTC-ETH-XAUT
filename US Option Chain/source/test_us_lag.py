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
import pywintypes
from unittest.mock import MagicMock, patch


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

    def test_oi_columns_are_never_text_again(self):
        """OI must stay numeric: the old write path re-armed "@" (text format)
        on the OI columns whenever a block grew a row, fighting the numeric
        format force_usd_block_formats applies and putting text into the
        customer's XLOOKUP formulas. The "@" on the date cells D3/O3/Z3 is a
        separate, deliberate format (dates are text) and must stay."""
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertNotIn(
            "sheet.range((5, c0 + off), (4 + n_rows, c0 + off)).api.NumberFormat",
            src,
        )
        self.assertNotIn("_oi_fmt_done", src)
        self.assertIn("force_usd_block_formats", src)

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
    "cannot open two workbooks with the same name" dialog on screen. The
    private-instance contract replaced all of that: one Excel instance owned
    by the tool, books opened in it once at startup, and a deliberately
    closed workbook stays closed.
    """

    def setUp(self):
        import us_excel
        self.main = us_excel

    def test_latch_exists_per_role(self):
        self.assertEqual(
            sorted(self.main.us_state._auto_open_done.keys()),
            sorted(self.main.ROLE_ORDER),
        )

    def test_reopen_is_gated_on_the_latch(self):
        import inspect
        import us_workbook
        src = inspect.getsource(us_workbook._open_excel_files)
        self.assertIn("_owner_book(app, path) is not None", src)

    def test_startup_open_sets_the_latch(self):
        import inspect
        import us_workbook
        src = inspect.getsource(us_workbook._open_excel_files)
        self.assertIn("_auto_open_done[role] = True", src)

    def test_only_two_places_launch_excel(self):
        """No os.startfile workbook reopen may survive anywhere.

        main.py keeps a single os.startfile for opening the self-test report
        file, never a workbook.
        """
        import inspect
        import main
        import us_workbook
        import us_excel
        self.assertEqual(
            inspect.getsource(us_workbook).count("os.startfile(")
            + inspect.getsource(us_excel).count("os.startfile("),
            0,
        )
        main_src = inspect.getsource(main)
        self.assertEqual(main_src.count("os.startfile("), 1)
        self.assertIn("report_path", main_src)

    def test_chain_value_writes_anchor_at_row_five(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._flush_block_writes)
        self.assertIn("sheet.range((5 + start, block[\"start\"])).value = [", src)
        self.assertIn("sheet.range((5 + lo, block[\"start\"])).value = merged_rows", src)
        self.assertNotIn("(4 + end, block[\"start\"] + 8)", src)
        self.assertNotIn("(4 + hi, block[\"start\"] + 8)", src)

    def test_stale_clear_only_when_below_live_rows(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn("if last_row > n_rows:", src)

    def test_latch_rearm_paths_exist(self):
        import inspect
        src = inspect.getsource(self.main._write_to_excel)
        # The rearm is GONE: a confirmed absence keeps the workbook closed.
        self.assertNotIn("excel_auto_open_rearmed", src)
        # A busy Excel is recorded as unknown, never as missing.
        self.assertIn("excel_attach_unknown", src)
        self.assertIn("workbook_closed_by_customer", src)
        self.assertNotIn("os.startfile(", src)
        self.assertNotIn("Microsoft Excel is open", src)

    def test_startup_open_runs_inline_on_the_display_thread(self):
        """The open pass must not live on its own thread.

        An opener thread created the private Excel instance while the
        display loop read that App object from another COM apartment:
        every probe from the display thread then answered "marshalled for
        a different thread", the books never attached and the close-all
        timer ended the tool with Excel on the screen. Inline on the
        display thread is the contract now; a wedged pass gets exactly
        one retry from the same thread after a grace period.
        """
        import inspect
        import us_workbook
        src = inspect.getsource(us_workbook.open_excel_files)
        self.assertNotIn("threading.Thread", src)
        self.assertNotIn("finished.wait(", src)
        self.assertIn("_open_excel_files()", src)
        self.assertIn("_open_pass_at", src)
        starter = inspect.getsource(us_workbook._open_excel_files)
        self.assertIn("excel_auto_open_started", starter)
        self.assertIn("excel_auto_open_failed", starter)
        self.assertIn("_ensure_owner_app", starter)
        self.assertNotIn("apps[0]", starter)
        self.assertNotIn("list(xw.apps)", starter)
        import us_excel
        retry_src = inspect.getsource(us_excel.retry_open_pass_after_grace)
        self.assertIn("_OPEN_RETRY_GRACE_SECONDS", retry_src)
        self.assertIn("excel_open_retried", retry_src)
        self.assertIn("open_excel_files()", retry_src)
        loop_src = inspect.getsource(us_excel.display_loop)
        self.assertIn("retry_open_pass_after_grace()", loop_src)

    def test_close_all_exits_the_tool(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel.maybe_exit_after_close_all)
        self.assertIn("tool_exit_close_all", src)
        self.assertIn("os._exit(0)", src)

    def test_display_thread_owns_a_com_apartment(self):
        """main_async must out-live the open pass's CoInit/CoUninit pair.

        Without its own apartment on the display thread, the open pass's
        CoUninitialize tears down the only apartment in the thread and
        ejects the Excel proxies created inside it; the writer then saw
        "Object is not connected to server" for Excel that was running,
        nothing ever attached, and the close-all timer killed the tool.
        """
        import inspect
        import main
        src = inspect.getsource(main.main_async)
        self.assertIn("pythoncom.CoInitialize()", src)
        self.assertLess(
            src.index("pythoncom.CoInitialize()"),
            src.index("open_excel_files()"),
        )

    def test_display_loop_pumps_the_com_apartment(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel.display_loop)
        self.assertIn("_pump_com()", src)

    def test_exact_book_match_normalizes_short_form_paths(self):
        import inspect
        import us_excel
        # The private-instance finder keeps the normalisation.
        src = inspect.getsource(us_excel.find_open_book)
        self.assertIn("long_form(os.path.abspath(book.fullname))", src)
        self.assertIn("long_form(target)", src)


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
        self.assertIn(
            "for bi, block, row_block, n_rows, runs, prev in pending:", src
        )
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
        from unittest.mock import patch

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
        # The sibling delta tool may genuinely be running on this machine while
        # the tests run - the watchdog must stay environment-independent here.
        with patch.object(
            us_excel.calc_mode_guard, "other_holder_alive", return_value=False
        ):
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




class CalcModeCoexistenceStartupTests(unittest.TestCase):
    """Startup race: second tool must skip calc switch when sibling holds guard."""

    def test_other_holder_alive_check_in_bulk_mode(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._begin_excel_bulk_mode)
        self.assertIn("other_holder_alive", src)

    def test_skips_switch_when_sibling_holds_guard(self):
        from unittest.mock import patch, MagicMock
        import us_excel
        import us_state
        app = MagicMock()
        app.api.Calculation = -4105  # automatic
        us_state._calc_session["failed"] = False
        us_state._app_cache = app
        with patch.object(us_excel.calc_mode_guard, 'other_holder_alive', return_value=True):
            with patch.object(us_excel.calc_mode_guard, 'acquire') as mock_acquire:
                restore = us_excel._begin_excel_bulk_mode(app)
                self.assertIsNotNone(restore)
                self.assertFalse(mock_acquire.called,
                                 'acquire should not be called when sibling holds guard')


class ChainExtentStabilityTests(unittest.TestCase):
    """The customer's strategy formulas carry literal ranges.

    The complaint: their own sheets (VLOOKUP/XLOOKUP against the chain
    columns) showed values that were wrong and lagging. Root cause one: a
    chain whose live row count changes with every expiry walks rows out from
    under those literal ranges, so one leg of the lookup silently becomes the
    if_not_found default. The write path therefore pads every block to a
    fixed extent; root cause two (the 0.1s default refresh starving Excel)
    and root cause three (the OI text-format re-arm fighting the numeric
    format) are fixed alongside.
    """

    def test_blocks_write_a_fixed_extent(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn("ROW_EXTENT_DEFAULT", src)
        self.assertIn("chain_extent_changed", src)
        self.assertIn("n_rows = extent", src)

    def test_flush_writes_padded_rows_not_the_dataframe(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._flush_block_writes)
        # The padded rows are the payload; a raw DataFrame write would push
        # an unfixed row count at the customer's ranges again.
        self.assertNotIn("chain.iloc", src)
        self.assertIn("row_block", src)

    def test_write_path_pads_before_diffing(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn("pad_block_rows(values, extent)", src)
        self.assertIn(
            "pending_writes.append((bi, block, values, n_rows, runs, prev))", src
        )

    def test_pad_block_rows_pads_with_blank_rows(self):
        import us_excel
        self.assertEqual(
            us_excel.pad_block_rows([[1, 2], [3, 4]], 4),
            [[1, 2], [3, 4], [None, None], [None, None]],
        )
        self.assertEqual(
            us_excel.pad_block_rows([], 2), [[None] * 9, [None] * 9]
        )
        self.assertEqual(us_excel.pad_block_rows([[1, 2]], 1), [[1, 2]])

    def test_deep_clear_sweeps_legacy_junk_once_per_block(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._deep_clear_below_block)
        self.assertIn("_deep_swept", src)
        self.assertIn("chain_deep_clear", src)

    def test_extent_is_not_shrunk_after_growing(self):
        """A stuck-at-max keeps every literal customer range valid forever."""
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn(
            "extent = max(us_state.ROW_EXTENT_DEFAULT, seen, len(values))", src
        )
        # No branch ever lowers the seen extent: the only write to it comes
        # from the grow path, which only ever records the max.
        self.assertNotIn("_row_extent_seen[(role, bi)] = len(values)", src)


class RefreshWidenTests(unittest.TestCase):
    """When Excel still owes a recalc after our window the loop stretches.

    The probe already logs the signal (excel_recalc_second_pass); the widen
    turns that signal into idle time for the customer's own workbook, and
    decays back once Excel comes out of every write window clean.
    """

    def setUp(self):
        import us_state
        self._saved = (
            us_state._refresh_widen,
            us_state._refresh_widen_hold_until,
            us_state._calc_state_probe,
            us_state._app_cache,
        )

    def tearDown(self):
        import us_state
        us_state._refresh_widen = self._saved[0]
        us_state._refresh_widen_hold_until = self._saved[1]
        us_state._calc_state_probe = self._saved[2]
        us_state._app_cache = self._saved[3]

    def test_second_pass_widens_refresh(self):
        import us_excel
        import us_state

        class _App:
            class api:
                CalculationState = 2

        us_state._refresh_widen = 1.0
        us_state._calc_state_probe = None
        us_state._app_cache = _App()
        us_excel._probe_calc_state_after_restore()
        self.assertEqual(us_state._refresh_widen, 1.5)
        self.assertGreater(us_state._refresh_widen_hold_until, 0.0)

    def test_widen_caps_at_four_and_holds_through_the_window(self):
        import us_excel
        import us_state

        us_state._refresh_widen = 3.0
        held_at = time.monotonic() + 10.0
        us_state._refresh_widen_hold_until = held_at
        us_excel._note_recalc_second_pass(time.monotonic())
        # 3.0 * 1.5 = 4.5 caps at WIDEN_MAX.
        self.assertEqual(us_state._refresh_widen, 4.0)
        # The hold extends past the previous one, so the display loop cannot
        # decay the stretch in the middle of the spell.
        self.assertGreater(us_state._refresh_widen_hold_until, held_at)

    def test_display_loop_uses_and_decays_widen(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel.display_loop)
        self.assertIn("REFRESH * us_state._refresh_widen", src)
        self.assertIn("refresh_widen_decayed", src)
        self.assertIn("_refresh_widen_hold_until", src)


class ExpiryLayoutTests(unittest.TestCase):
    """US sheets show two expiry blocks; the third is retired once per session.

    The customer's strategy sheet reads blocks 1+2 only, so block 3 was a
    third of the per-cycle Excel write/recalc traffic for columns nobody
    looks at. The blocks retire the delta way: cleared once per session,
    then left alone until the next session.
    """

    def test_third_block_wipes_on_connect_not_on_first_write(self):
        """The sweep must ride the per-cycle connect step, not the first
        data write - a cold start takes minutes and the customer stares at
        the old Block 3 the whole time otherwise."""
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn("_retire_unused_blocks(sheet, role)", src)
        self.assertLess(
            src.index("_retire_unused_blocks(sheet, role)"),
            src.index("if not sheet_structure_ok(sheet, role)"),
        )

    def test_two_blocks_are_shown(self):
        import us_workbook
        self.assertEqual(us_workbook.BLOCKS_PER_SHEET, 2)
        # Block 3 remains in the layout constants (start col 23) so it can
        # be switched back on, but no active block starts past it.
        self.assertEqual(len(us_workbook.BLOCK_COLS), 3)
        self.assertEqual(
            tuple(b["start"] for b in us_workbook.BLOCK_COLS[:us_workbook.BLOCKS_PER_SHEET]),
            (1, 12),
        )

    def test_expiry_cells_match_blocks(self):
        import us_workbook
        self.assertEqual(us_workbook.EXPIRY_CELLS, ("D3", "O3"))

    def test_writer_loops_only_active_blocks(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel._write_to_excel)
        self.assertIn("enumerate(BLOCK_COLS[:BLOCKS_PER_SHEET])", src)
        self.assertIn("_retire_unused_blocks", src)
        self.assertIn("EXPIRY_CELLS[bi]", src)

    def test_template_writes_only_two_expiry_labels(self):
        import inspect
        import us_workbook
        src = inspect.getsource(us_workbook._write_template)
        self.assertNotIn("(edit Z3)", src)
        self.assertIn("Expiry {si}", src)
        self.assertIn("f\"Expiry {si}\"", src)

    def test_retire_sweeps_unused_columns_once_per_session(self):
        import us_excel
        import us_workbook

        class R:
            def __init__(self, sheet, ref):
                self._sheet = sheet
                self._ref = ref
            def clear(self):
                self._sheet.cleared.append(self._ref)
            def delete(self):
                pass

        class Sheet:
            def __init__(self):
                self.cleared = []
            def range(self, *args):
                return R(self, args)

        sheet = Sheet()
        saved = us_excel._retired_sheets
        us_excel._retired_sheets = set()
        try:
            us_excel._retire_unused_blocks(sheet, "stock")
            # Once per sheet object: a second call for the same sheet is a no-op.
            us_excel._retire_unused_blocks(sheet, "stock")
        finally:
            us_excel._retired_sheets = saved
        # Block 3 spans (1,23)..(2000,31) and is swept exactly once.
        self.assertEqual(sheet.cleared, [((1, 23), (2000, 31))])


if __name__ == "__main__":
    unittest.main()

class WatchdogReconnectTests(unittest.TestCase):
    """ensure_calculation_not_stranded should reconnect when _app_cache is None."""

    def test_reconnect_in_us_watchdog(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel.ensure_calculation_not_stranded)
        self.assertIn('bulk_app', src)

class CalcShutdownSuppressionTests(unittest.TestCase):
    """Closing the workbook must not put the calc-mode dialog on screen.

    Answering Excel's save prompt refuses every COM call for a few seconds
    while EXCEL.EXE is still alive, so checking the process cannot see it.
    The notice is only allowed once a refusal has held for the grace period
    WITH the chain workbooks still open, and the refusal marks are cleared
    the moment the chains are gone - which is also when a held manual window
    is handed back.
    """

    def setUp(self):
        import us_state

        self._saved = (
            us_state._calc_manual_failed,
            us_state._calc_failed_since,
            us_state._calc_failed_at,
            us_state._app_cache,
            dict(us_state._calc_session),
        )

    def tearDown(self):
        import us_state

        (
            us_state._calc_manual_failed,
            us_state._calc_failed_since,
            us_state._calc_failed_at,
            us_state._app_cache,
        ) = self._saved[:4]
        us_state._calc_session.clear()
        us_state._calc_session.update(self._saved[4])

    def test_no_immediate_dialog_on_a_refused_switch(self):
        import inspect
        import us_excel

        src = inspect.getsource(us_excel._begin_excel_bulk_mode)
        self.assertNotIn("show_error", src)

    def test_the_notice_waits_for_the_grace_period(self):
        import inspect
        import us_excel

        src = inspect.getsource(us_excel.render_cycle)
        self.assertIn("_CALC_NOTICE_GRACE_SECONDS", src)
        self.assertIn("us_state._calc_failed_since", src)

    def test_marks_are_cleared_when_the_chains_are_gone(self):
        from unittest.mock import MagicMock
        import us_state
        import us_excel

        us_state._calc_manual_failed = True
        us_state._calc_failed_since = 12345.0
        book = MagicMock()
        book.name = "StrategyBook.xlsx"
        app = MagicMock()
        app.books = [book]
        us_state._app_cache = app
        us_excel._clear_refused_calc_if_chain_gone()
        self.assertFalse(us_state._calc_manual_failed)
        self.assertEqual(us_state._calc_failed_since, 0.0)

    def test_marks_survive_while_the_chains_are_open(self):
        from unittest.mock import MagicMock
        import us_state
        import us_excel

        us_state._calc_manual_failed = True
        us_state._calc_failed_since = 12345.0
        book = MagicMock()
        book.name = "spx_option_chain.xlsx"
        app = MagicMock()
        app.books = [book]
        us_state._app_cache = app
        us_excel._clear_refused_calc_if_chain_gone()
        self.assertTrue(us_state._calc_manual_failed)
        self.assertEqual(us_state._calc_failed_since, 12345.0)

    def test_closing_the_chains_releases_a_held_manual_window(self):
        from unittest.mock import MagicMock, patch
        import us_state
        import us_excel

        book = MagicMock()
        book.name = "Book1.xlsx"
        app = MagicMock()
        app.books = [book]
        us_state._app_cache = app
        us_state._calc_session.update(
            {"owned": True, "failed": False, "mode": "automatic", "app": app}
        )
        with patch.object(us_excel, "_set_calc_mode") as set_mode:
            with patch.object(us_excel.calc_mode_guard, "release") as release:
                result = us_excel._begin_excel_bulk_mode(app)
        self.assertIsNone(result)
        self.assertTrue(set_mode.called)
        self.assertTrue(release.called)
        self.assertFalse(us_state._calc_session.get("owned"))

    def test_the_chains_open_keeping_the_window(self):
        from unittest.mock import MagicMock, patch
        import us_state
        import us_excel

        book = MagicMock()
        book.name = "us_stock_option_chain.xlsx"
        app = MagicMock()
        app.books = [book]
        app.api.Calculation = -4105
        us_state._app_cache = app
        with patch.object(
            us_excel.calc_mode_guard, "other_holder_alive", return_value=False
        ):
            with patch.object(us_excel.calc_mode_guard, "acquire") as acquire:
                with patch.object(us_excel.calc_mode_guard, "release") as release:
                    result = us_excel._begin_excel_bulk_mode(app)
        self.assertIsNotNone(result)
        self.assertFalse(release.called)
        self.assertTrue(acquire.called)


class DeadExcelRefusalTests(unittest.TestCase):
    """Excel closing while we own the manual window is not a refusal.

    A taskbar "Close all windows" leaves every handle dead: the next bending
    of the manual switch used to latch a refusal and pop the slow-down notice
    a minute later. Dead COM means Excel is gone - drop the session and the
    marks, and never show that notice about it.
    """

    DEAD = pywintypes.com_error(-2146827864, "OLE error 0x800a01a8", None, None)

    def setUp(self):
        import us_excel
        import us_state

        self.us_excel = us_excel
        self.us_state = us_state
        self._guard_probe = patch.object(
            us_excel.calc_mode_guard, "other_holder_alive", return_value=False
        )
        self._guard_probe.start()
        self._release_probe = patch.object(us_excel.calc_mode_guard, "release")
        self._release_probe.start()
        self._saved = (
            us_state._calc_manual_failed,
            us_state._calc_failed_since,
            us_state._calc_failed_at,
            us_state._app_cache,
            dict(us_excel._calc_session),
        )
        us_state._calc_manual_failed = False
        us_state._calc_failed_since = 0.0
        us_state._calc_failed_at = 0.0

    def tearDown(self):
        us_excel = self.us_excel
        us_state = self.us_state
        self._guard_probe.stop()
        self._release_probe.stop()
        (
            us_state._calc_manual_failed,
            us_state._calc_failed_since,
            us_state._calc_failed_at,
            us_state._app_cache,
        ) = self._saved[:4]
        us_excel._calc_session.clear()
        us_excel._calc_session.update(self._saved[4])

    def _dead_app(self):
        dead_err = self.DEAD

        class DeadApp:
            @property
            def version(self):
                raise dead_err

        return DeadApp()

    def test_app_cache_none_clears_refusal_marks(self):
        us_excel = self.us_excel
        us_state = self.us_state

        us_state._calc_manual_failed = True
        us_state._app_cache = None
        us_excel._clear_refused_calc_if_chain_gone()
        self.assertFalse(us_state._calc_manual_failed)

    def test_dead_enumeration_with_no_process_clears_marks(self):
        us_excel = self.us_excel
        us_state = self.us_state

        dead_err = self.DEAD

        class LockedApp:
            @property
            def books(self):
                raise dead_err

        with patch.object(us_excel, "excel_process_running", lambda: False):
            us_state._calc_manual_failed = True
            us_state._app_cache = LockedApp()
            us_excel._clear_refused_calc_if_chain_gone()
            self.assertFalse(us_state._calc_manual_failed)

        with patch.object(us_excel, "excel_process_running", lambda: True):
            us_state._calc_manual_failed = True
            us_state._app_cache = LockedApp()
            us_excel._clear_refused_calc_if_chain_gone()
            # A live Excel with an unreadable book list might still open the
            # chains - that stays "cannot tell".
            self.assertTrue(us_state._calc_manual_failed)

    def test_dead_version_probe_ends_the_session(self):
        us_excel = self.us_excel
        us_state = self.us_state

        us_state._calc_manual_failed = True
        us_state._app_cache = self._dead_app()
        result = us_excel._begin_excel_bulk_mode(us_state._app_cache)
        self.assertIsNone(result)
        self.assertIsNone(us_state._app_cache)
        self.assertFalse(us_state._calc_manual_failed)

    def test_dead_switch_is_not_latched_but_living_refusal_is(self):
        us_excel = self.us_excel
        us_state = self.us_state

        class ChainApp:
            def __init__(self, err):
                self.err = err

            @property
            def version(self):
                return "16.0"

            @property
            def books(self):
                return [type("B", (), {"name": "us_stock_option_chain.xlsx"})()]

            @property
            def screen_updating(self):
                return False

            @property
            def api(self):
                return type("Api", (), {"Calculation": 2})()

            calculation = property(
                lambda self: "manual",
                lambda self, v: self._raise(),
            )

            def _raise(self):
                raise self.err

        dead = ChainApp(self.DEAD)
        us_state._app_cache = dead
        us_excel._begin_excel_bulk_mode(dead)
        self.assertFalse(us_state._calc_manual_failed)
        self.assertIsNone(us_state._app_cache)

        alive_refusal = ChainApp(RuntimeError("nope, dialog is up"))
        us_state._app_cache = alive_refusal
        with patch.object(us_excel, "excel_process_running", lambda: True):
            us_excel._begin_excel_bulk_mode(alive_refusal)
        self.assertTrue(us_state._calc_manual_failed)


    def test_teardown_error_with_no_process_is_a_close(self):
        # The 13:58 shape: Excel throws "Unable to set the Calculation
        # property" (a real-refusal HRESULT) while the process is dying from
        # a taskbar close-all. The process check is the ground truth.
        us_excel = self.us_excel
        us_state = self.us_state

        teardown_err = pywintypes.com_error(
            -2147352567, "Exception occurred.", None, None
        )

        class TeardownApp:
            def __init__(self):
                self.err = teardown_err

            @property
            def version(self):
                return "16.0"

            @property
            def books(self):
                return [type("B", (), {"name": "us_stock_option_chain.xlsx"})()]

            @property
            def screen_updating(self):
                return False

            @property
            def api(self):
                return type("Api", (), {"Calculation": 2})()

            calculation = property(
                lambda self: "manual",
                lambda self, v: self._raise(),
            )

            def _raise(self):
                raise self.err

        app = TeardownApp()
        us_state._app_cache = app
        with patch.object(us_excel, "excel_process_running", lambda: False):
            us_excel._begin_excel_bulk_mode(app)
        self.assertFalse(us_state._calc_manual_failed)
        self.assertIsNone(us_state._app_cache)

        alive = TeardownApp()
        us_state._app_cache = alive
        with patch.object(us_excel, "excel_process_running", lambda: True):
            us_excel._begin_excel_bulk_mode(alive)
        # Same error while Excel is genuinely there stays a refusal.
        self.assertTrue(us_state._calc_manual_failed)


class ChainSavePromptAnswererTests(unittest.TestCase):
    """The tool answers Excel's Save / Don't Save prompt over chain books.

    A customer closing the workbook would otherwise leave that prompt in the
    background: it blocks every COM call until it is answered, which is also
    how the refused manual-calculation switches start. Nothing unsaved in a
    chain workbook is worth keeping - every cell is rewritten next cycle -
    so Don't Save is safe there, and only there.
    """

    def test_the_answerer_runs_off_the_render_thread(self):
        import inspect
        import us_excel

        src = inspect.getsource(us_excel.display_loop)
        self.assertIn("start_chain_save_prompt_watchdog()", src)
        starter = inspect.getsource(us_excel.start_chain_save_prompt_watchdog)
        self.assertIn("_prompt_watchdog_started", starter)
        self.assertIn('platform.system() != "Windows"', starter)
        self.assertIn("threading.Thread(", starter)
        self.assertIn("daemon=True", starter)

    def test_only_chain_windows_are_answered(self):
        from unittest.mock import patch

        import us_excel

        fake = _FakeUser32("SPX | Option Chain None | LIVE 09:08:13 - Excel")
        with patch.dict("sys.modules"):  # keep the live import well-behaved
            with patch("ctypes.windll") as windll:
                windll.user32 = fake.user32
                self.assertTrue(
                    us_excel._prompt_belongs_to_chain(0x1234),
                    "a prompt over the rendered chain frame must be answered",
                )

        fake = _FakeUser32("us_stock_option_chain - Excel")
        with patch("ctypes.windll") as windll:
            windll.user32 = fake.user32
            # file captions join words with underscores
            self.assertTrue(
                us_excel._prompt_belongs_to_chain(0x1234),
                "the file-backed caption must be recognised too",
            )

        fake = _FakeUser32("Research Budget 2026 - Excel")
        with patch("ctypes.windll") as windll:
            windll.user32 = fake.user32
            self.assertFalse(
                us_excel._prompt_belongs_to_chain(0x1234),
                "the customer's own workbook keeps its prompt untouched",
            )

    def test_the_modern_prompt_is_answered_without_fighting_window_shapes(self):
        import inspect
        import us_excel

        src = inspect.getsource(us_excel._answer_save_prompt)
        self.assertIn('FindWindowExW(hwnd, button, "Button", None)', src)
        self.assertIn("0x00F5", src)
        self.assertIn("_send_alt_n(hwnd)", src)
        accel = inspect.getsource(us_excel._send_alt_n)
        self.assertIn("SetForegroundWindow(hwnd)", accel)
        self.assertIn("0x12", accel)
        self.assertIn("0x4E", accel)

    def test_a_stuck_prompt_is_never_hammered_forever(self):
        import inspect
        import us_excel

        self.assertEqual(us_excel._SAVE_PROMPT_CLASSES, ("NUIDialog", "#32770"))
        src = inspect.getsource(us_excel._chain_save_prompt_once)
        self.assertIn("_SAVE_PROMPT_ATTEMPT_LIMIT", src)
        self.assertIn("attempts[hwnd] = count + 1", src)
        self.assertIn("feed_log(", src)
        looper = inspect.getsource(us_excel._chain_save_prompt_watchdog)
        self.assertIn("time.sleep(", looper)
        self.assertIn("except Exception", looper)

    def test_the_question_only_travels_to_chain_owned_dialogs(self):
        import inspect
        import us_excel

        self.assertEqual(us_excel._SAVE_PROMPT_ATTEMPT_LIMIT, 8)
        self.assertGreater(us_excel._SAVE_PROMPT_POLL_SECONDS, 0.0)
        src = inspect.getsource(us_excel._prompt_belongs_to_chain)
        self.assertIn("GetAncestor", src)
        # underscore captions like us_stock_option_chain must match too
        self.assertIn('.replace("_", " ")', src)
        self.assertIn('"option chain" in caption', src)


class _FakeUser32:
    """A stand-in for user32 so _prompt_belongs_to_chain can run without a
    real window hierarchy on the test machine.

    GetAncestor(3) returns the owner handle; GetWindowTextW fills the caller's
    buffer with the configured caption, matching the real signature closely
    enough for the one window the function looks at.
    """

    def __init__(self, caption):
        import ctypes as _ct

        self.caption = caption
        self._buf_type = _ct.create_unicode_buffer

        class _U32:
            GetAncestor = staticmethod(lambda hwnd, flag: 0x4321 if flag == 3 else 0)

            def GetWindowTextW(self, hwnd, buf, size):  # noqa: N802
                buf.value = caption
                return len(caption)

            def IsWindow(self, hwnd):
                return True

            def IsWindowVisible(self, hwnd):
                return True

        self.user32 = _U32()
