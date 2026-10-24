import inspect
import math
import os
import time
import unittest
import pywintypes
from unittest.mock import patch

import main as delta_main
import delta_state


class DeltaComTaxTests(unittest.TestCase):
    """The C2 COM-tax fixes ported into the Delta tool."""


    def test_environment_detail_only_on_connection_change(self):
        src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn('record_excel_connection(asset, "connected", "")', src)
        self.assertIn('_excel_connection_status.get(asset) != "connected"', src)
        self.assertIn("excel_environment_detail(book)", src)

    def test_status_strip_is_one_merged_read(self):
        src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn("sheet.range((1, 1), (4, 31)).value", src)
        self.assertNotIn('sheet.range("A2").value', src)

    def test_date_cells_read_through_the_merge(self):
        src = inspect.getsource(delta_main.update_user_dates)
        self.assertIn("enumerate((4, 15, 26))", src)
        self.assertNotIn('"D3", "O3", "Z3"', src)

    def test_tasklist_cache_is_two_seconds(self):
        src = inspect.getsource(delta_main.excel_process_running)
        self.assertIn("max_age=2.0", src)

    def test_calc_state_probe_is_gated_to_thirty_seconds(self):
        src = inspect.getsource(delta_main._probe_calc_state_after_restore)
        self.assertIn("< 30.0", src)
        rc = inspect.getsource(delta_main.render_cycle)
        self.assertNotIn("CalculationState", rc)
        self.assertIn("_probe_calc_state_after_restore()", rc)

    def test_watchdog_runs_once_per_cycle(self):
        src = inspect.getsource(delta_main.render_cycle)
        self.assertEqual(src.count("ensure_calculation_not_stranded("), 1)

    def test_merged_read_failure_falls_back_to_per_cell_reads(self):
        class FakeSheet:
            def range(self, spec):
                return FakeRange()

        class FakeRange:
            @property
            def value(self):
                return "sheet-read"

        self.assertEqual(
            delta_main.cell_value(None, FakeSheet(), 2, 1),
            "sheet-read",
        )
        self.assertIsNone(delta_main.cell_value(None, None, 2, 1))
        self.assertIsNone(delta_main.cell_value([[1]], None, 2, 1))
        self.assertEqual(delta_main.cell_value([[5, 7]], None, 1, 2), 7)


class DeltaSpanWriteTests(unittest.TestCase):
    """The C3 row-run write behaviour ported into the Delta tool."""


    def test_contiguous_runs_merge_adjacent_indices(self):
        self.assertEqual(
            delta_main._contiguous_runs([0, 3, 4, 5, 10]),
            [(0, 0), (3, 5), (10, 10)],
        )
        self.assertEqual(delta_main._contiguous_runs([]), [])

    def test_scattered_changes_write_one_merged_span(self):
        src = inspect.getsource(delta_main._flush_block_writes)
        self.assertIn("merged_rows", src)
        self.assertNotIn("for start, end in runs:", src)

    def test_unchanged_rows_carry_previous_values(self):
        src = inspect.getsource(delta_main._flush_block_writes)
        self.assertIn("merged_rows.append(list(prev[i]))", src)

    def test_ownership_flag_gates_row_granularity(self):
        src = inspect.getsource(delta_main._begin_excel_bulk_mode)
        self.assertIn("delta_state._bulk_owning_manual = True", src)
        self.assertIn("delta_state._bulk_owning_manual = False", src)
        self.assertIn("protected = delta_state._bulk_owning_manual",
                      inspect.getsource(delta_main._flush_block_writes))


class DeltaFormattingTests(unittest.TestCase):
    """A1: the OI formatting cost was 340-712 ms of per-cell COM walking."""

    def test_no_per_cell_walk_left_in_formats(self):
        src = inspect.getsource(delta_main.force_usd_block_formats)
        self.assertNotIn("for r in range(", src)
        self.assertNotIn("cell.value", src)
        self.assertIn("values = rng.value", src)
        self.assertIn("rng.value = values", src)

    def test_format_gate_ignores_row_count(self):
        src = inspect.getsource(delta_main.force_usd_block_formats)
        self.assertIn('(sheet_key, c0, "numfmt")', src)
        self.assertNotIn("(asset, bi, n_rows)", src)


class _FakeRange:
    def __init__(self, recorder, top, left, bottom, right):
        self._recorder = recorder
        self.top = top
        self.left = left
        self.bottom = bottom
        self.right = right

    @property
    def value(self):
        return None

    @value.setter
    def value(self, data):
        self._recorder.append((self.top, self.left, self.bottom, self.right, data))


class _FakeSheet:
    """Records every range write so a test can inspect the exact grid."""

    def __init__(self):
        self.writes = []

    def range(self, spec, other=None):
        if other is None:
            row, col = spec
            return _FakeRange(self.writes, row, col, row, col)
        return _FakeRange(self.writes, spec[0], spec[1], other[0], other[1])

    @property
    def only_write(self):
        assert len(self.writes) == 1, "expected one COM call, got %d" % len(self.writes)
        return self.writes[0]


class DeadExcelRefusalTests(unittest.TestCase):
    """Excel closing under us must not arm the manual-switch notice.

    A taskbar Close-all leaves every handle dead; the next manual switch used
    to latch a refusal and pop the slow-down dialog a minute later. Dead COM
    means Excel is gone - drop the session and never show that dialog.
    """

    DEAD = pywintypes.com_error(-2146827864, "OLE error 0x800a01a8", None, None)

    def setUp(self):
        delta_main.delta_state._calc_manual_failed = True
        delta_main.delta_state._calc_failed_since = 0.0
        self._release_probe = patch.object(
            delta_main.calc_mode_guard, "release"
        )
        self._release_probe.start()
        self._holder_probe = patch.object(
            delta_main.calc_mode_guard, "other_holder_alive", return_value=False
        )
        self._holder_probe.start()
        self._acquire_probe = patch.object(delta_main.calc_mode_guard, "acquire")
        self._acquire_probe.start()

    def tearDown(self):
        self._release_probe.stop()
        self._holder_probe.stop()
        self._acquire_probe.stop()
        delta_main.delta_state._calc_manual_failed = False
        delta_main.delta_state._calc_failed_since = 0.0
        delta_main.delta_state._app_cache = None
        delta_main.delta_state._bulk_owning_manual = False

    def test_com_dead_classifier_separates_dead_from_busy(self):
        self.assertTrue(delta_main._com_dead(self.DEAD))
        self.assertFalse(
            delta_main._com_dead(
                pywintypes.com_error(-2147352567, "modal up", None, None)
            )
        )
        self.assertFalse(delta_main._com_dead(RuntimeError("dialog is up")))

    def test_dead_version_probe_drops_marks_and_cache(self):
        dead_err = self.DEAD

        class DeadApp:
            @property
            def version(self):
                raise dead_err

        ds = delta_main.delta_state
        ds._app_cache = DeadApp()
        result = delta_main._begin_excel_bulk_mode()
        self.assertIsNone(result)
        self.assertIsNone(ds._app_cache)
        self.assertFalse(ds._calc_manual_failed)
        self.assertFalse(ds._bulk_owning_manual)


    def test_teardown_error_with_no_process_drops_marks(self):
        import pywintypes as _pw

        teardown_err = _pw.com_error(-2147352567, "Exception occurred.", None, None)

        class TeardownApp:
            def __init__(self):
                self.err = teardown_err

            @property
            def version(self):
                return "16.0"

            @property
            def books(self):
                return [type("B", (), {"name": "btc_chain.xlsx"})()]

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

        ds = delta_main.delta_state
        ds._calc_manual_failed = False
        ds._app_cache = TeardownApp()
        with patch.object(delta_main, "excel_process_running", lambda: False):
            delta_main._begin_excel_bulk_mode()
        self.assertFalse(ds._calc_manual_failed)
        self.assertIsNone(ds._app_cache)

        ds._app_cache = TeardownApp()
        with patch.object(delta_main, "excel_process_running", lambda: True):
            delta_main._begin_excel_bulk_mode()
        self.assertTrue(ds._calc_manual_failed)


class DeltaMergedBlockWriteTests(unittest.TestCase):
    """The combined block write, checked on the grid it actually produces.

    The US tool tried writing all three blocks as one wide range and reverted
    it: the combined grid carried blanks for the blocks that had not changed,
    so an unchanged block came back empty and the sheet looked like it was
    blanking out and then filling in again. These tests pin the two rules that
    make the combined write safe.
    """

    def setUp(self):
        import delta_state
        import pandas as pd
        import main as delta_main
        from unittest.mock import patch

        self.delta_main = delta_main
        self.delta_state = delta_state
        self.pd = pd
        self.patch = patch

    def _pending(self, block_index, values, prev, runs):
        """Build one queued block the way _write_to_excel queues it.

        The pending tuple carries the padded rows (here the test data stands
        in for them) - no DataFrame - just like the writer now queues.
        """
        block = self.delta_main.BLOCK_COLS[block_index]
        return (block_index, block, values, len(values), runs, prev)

    def _flush(self, sheet, pending):
        with self.patch.object(self.delta_state, "_bulk_owning_manual", True):
            self.delta_main._flush_block_writes(sheet, "BTC", pending)

    def test_consecutive_blocks_go_out_as_one_call(self):
        values = [[1, 2, 3, 4, 5, 6, 7, 8, 9] for _ in range(4)]
        prev = [list(row) for row in values]
        pending = [
            self._pending(0, values, prev, [(0, 0)]),
            self._pending(1, values, prev, [(0, 0)]),
        ]
        sheet = _FakeSheet()
        self._flush(sheet, pending)
        top, left, bottom, right, grid = sheet.only_write
        # Top-left anchor: the payload carries the full merge width (the two
        # corner form was removed because a degenerate run turned 5..4 into a
        # row-4 stomp).
        self.assertEqual((top, left), (5, 1))
        self.assertEqual(len(grid[0]), self.delta_main.BLOCK_COLS[1]["start"] + 8)
        self.assertEqual(len(grid), 1)
        # Both blocks carry their current row, and the spacer columns between
        # them are blank rather than an old value.
        self.assertEqual(grid[0][0:9], values[0])
        # Blocks sit in columns 1-9 and 12-20, so columns 10 and 11 are the
        # spacer and must be blank rather than an old value.
        self.assertEqual(grid[0][9:11], [None, None])
        self.assertEqual(grid[0][11:20], values[0])

    def test_an_unchanged_block_is_never_written(self):
        values = [[1, 2, 3, 4, 5, 6, 7, 8, 9] for _ in range(4)]
        prev = [list(row) for row in values]
        # The writer only queues blocks that changed; the flush must then keep
        # block 2's columns out of the write entirely.
        pending = [self._pending(0, values, prev, [(0, 0)])]
        sheet = _FakeSheet()
        self._flush(sheet, pending)
        top, left, bottom, right, grid = sheet.only_write
        self.assertEqual((top, left), (5, 1))
        self.assertEqual(len(grid[0]), 9)

    def test_a_block_that_shrank_does_not_get_its_old_tail_back(self):
        # Block 0 went from five rows to two. Its previous rows 1..4 were
        # cleared a moment ago and must not be written back.
        stale = [["stale"] * 9 for _ in range(5)]
        fresh = [[0, 1, 2, 3, 4, 5, 6, 7, 8], [9, 9, 9, 9, 9, 9, 9, 9, 9]]
        other = [[1, 2, 3, 4, 5, 6, 7, 8, 9] for _ in range(6)]
        pending = [
            self._pending(0, fresh, stale, [(0, 0)]),
            self._pending(1, other, [list(r) for r in other], [(1, 2)]),
        ]
        sheet = _FakeSheet()
        self._flush(sheet, pending)
        top, left, bottom, right, grid = sheet.only_write
        flat = [cell for row in grid for cell in row]
        self.assertNotIn("stale", flat)
        # Row index 2 of block 0 is past its current row count, so it is blank
        # even though the block used to have data there.
        self.assertEqual(grid[2][0:9], [None] * 9)
        self.assertEqual(grid[1][0:9], fresh[1])

    def test_rows_inside_the_grid_carry_current_content(self):
        values = [[r] * 9 for r in range(10)]
        prev = [["old"] * 9 for _ in range(10)]
        other = [[r] * 9 for r in range(8)]
        pending = [
            self._pending(0, values, prev, [(0, 1)]),
            self._pending(1, other, [list(r) for r in other], [(5, 6)]),
        ]
        sheet = _FakeSheet()
        self._flush(sheet, pending)
        top, left, bottom, right, grid = sheet.only_write
        # Block 0 row 2 was not part of the changed run, so it carries the
        # block's current value, not the previous one.
        self.assertEqual(grid[2][0], 2)
        self.assertNotIn("old", [cell for row in grid for cell in row])

    def test_blocks_that_are_not_consecutive_are_written_separately(self):
        values = [[1, 2, 3, 4, 5, 6, 7, 8, 9] for _ in range(4)]
        prev = [list(row) for row in values]
        pending = [
            self._pending(0, values, prev, [(0, 0)]),
            self._pending(2, values, prev, [(0, 0)]),
        ]
        sheet = _FakeSheet()
        self._flush(sheet, pending)
        self.assertEqual(len(sheet.writes), 2)
        lefts = sorted(w[1] for w in sheet.writes)
        self.assertEqual(lefts, [1, 23])

    def test_a_huge_rectangle_is_not_used(self):
        values = [[1, 2, 3, 4, 5, 6, 7, 8, 9] for _ in range(60)]
        prev = [list(row) for row in values]
        # The two changed runs sit far apart, so one rectangle would dirty
        # hundreds of cells the customer never asked us to touch.
        pending = [
            self._pending(0, values, prev, [(0, 1)]),
            self._pending(1, values, prev, [(50, 51)]),
        ]
        sheet = _FakeSheet()
        self._flush(sheet, pending)
        self.assertEqual(len(sheet.writes), 2)


class DeltaExpiryRollerTests(unittest.TestCase):
    """A failed probe must not shift the slots.

    The slots are filled from the list of expiries that have data, so one
    missing expiry slides every slot along. The ladder then swaps to a
    different expiry - and swaps back on the next pass, when the probe
    works. The trader watches the chain change for no reason.
    """

    def setUp(self):
        import main as delta_main
        self.delta_main = delta_main

    def test_a_failed_probe_keeps_the_current_assignment(self):
        import inspect

        src = inspect.getsource(self.delta_main.expiry_roller)
        self.assertIn('if "ERROR" in states:', src)
        self.assertIn("roller_probe_incomplete", src)
        # The skip must come before the slots are reassigned.
        self.assertLess(
            src.index('if "ERROR" in states:'),
            src.index("changed_slots = []"),
        )

    def test_the_retry_recomputes_the_states(self):
        import inspect

        src = inspect.getsource(self.delta_main.expiry_roller)
        # After the retry the states have to be read again, or a probe that
        # failed the first time still looks failed after it succeeded.
        self.assertIn('states = [st for _, st in counts]', src)


class DeltaDelayNoteTests(unittest.TestCase):
    """What the header says when the data is not fresh.

    A trader looking at a frozen spot has no way of telling a quiet market
    from a broken feed. The note has to say so - and it has to be empty in
    the normal case, so the header text does not change and the cell is not
    rewritten on every cycle.
    """

    def setUp(self):
        import time

        import main as delta_main
        self.delta_main = delta_main
        self.time = time

    def test_a_healthy_feed_says_nothing(self):
        self.delta_main.feed_state["BTC"] = {
            "state": "OK",
            "detail": "",
            "retry_in": 0.0,
            "last_good_ts": time.time(),
        }
        self.delta_main.last_ws_recv["BTC"] = time.time()
        self.delta_main.last_spot_change["BTC"] = time.time()
        self.assertEqual(self.delta_main.source_delay_note("BTC"), "")

    def test_a_degraded_feed_is_announced(self):
        self.delta_main.feed_state["BTC"] = {
            "state": "ERROR",
            "detail": "http 503",
            "retry_in": 5.0,
            "last_good_ts": time.time() - 60,
        }
        note = self.delta_main.source_delay_note("BTC")
        self.assertIn("ERROR", note)

    def test_a_quiet_spot_is_announced(self):
        self.delta_main.feed_state["BTC"] = {
            "state": "OK",
            "detail": "",
            "retry_in": 0.0,
            "last_good_ts": time.time(),
        }
        self.delta_main.last_ws_recv["BTC"] = time.time() - 30
        self.delta_main.last_spot_change["BTC"] = time.time() - 30
        note = self.delta_main.source_delay_note("BTC")
        self.assertIn("SPOT STALE", note)

    def test_a_never_seen_feed_does_not_crash(self):
        self.delta_main.feed_state.pop("BTC", None)
        self.delta_main.last_ws_recv["BTC"] = 0.0
        self.delta_main.last_spot_change["BTC"] = 0.0
        self.assertIsInstance(self.delta_main.source_delay_note("BTC"), str)


class DeltaSheetStructureGuardTests(unittest.TestCase):
    """A restructured sheet must stop the tool, not be written over.

    The tool writes to fixed columns. If the customer inserts a column, those
    positions no longer line up and the tool would write over their own data.
    """

    class FakeSheet:
        def __init__(self, a4):
            self._a4 = a4

        def range(self, *args):
            class R:
                def __init__(self, v):
                    self.value = v

            return R(self._a4)

    def setUp(self):
        import main as delta_main
        self.delta_main = delta_main

    def test_intact_sheet_is_ok(self):
        sheet = self.FakeSheet("Call LTP")
        self.assertTrue(self.delta_main.sheet_structure_ok(sheet))

    def test_blank_sheet_is_ok(self):
        sheet = self.FakeSheet(None)
        self.assertTrue(self.delta_main.sheet_structure_ok(sheet))

    def test_inserted_column_is_detected(self):
        sheet = self.FakeSheet("My Notes")
        self.assertFalse(self.delta_main.sheet_structure_ok(sheet))

    def test_unreadable_sheet_is_allowed(self):
        class Broken:
            def range(self, *args):
                raise RuntimeError("RPC server is unavailable")

        self.assertTrue(self.delta_main.sheet_structure_ok(Broken()))


class DeltaExpirySettlementTests(unittest.TestCase):
    """Today's expiry must stay in the list until it has really expired.

    Delta settles BTC and ETH options at 12:00 UTC and XAUT at 16:00 UTC. The
    old code dropped today's expiry at 08:00 UTC for BTC and ETH, so for four
    hours of every trading day the contract that was about to expire was
    missing from the sheet.
    """

    def setUp(self):
        import main as delta_main
        self.delta_main = delta_main

    def _products(self, settlements):
        return [{"settlement_time": s} for s in settlements]

    def _at(self, hour, minute=0):
        from datetime import datetime, timezone
        return datetime(2026, 10, 22, hour, minute, tzinfo=timezone.utc)

    def test_todays_expiry_survives_until_it_settles(self):
        products = self._products([
            "2026-10-22T12:00:00Z",
            "2026-10-23T12:00:00Z",
        ])
        dates = self.delta_main.available_expiry_dates(
            "BTC", products, now=self._at(8, 30)
        )
        self.assertEqual(dates, ["2026-10-22", "2026-10-23"])

    def test_todays_expiry_is_dropped_once_it_has_settled(self):
        products = self._products([
            "2026-10-22T12:00:00Z",
            "2026-10-23T12:00:00Z",
        ])
        dates = self.delta_main.available_expiry_dates(
            "BTC", products, now=self._at(12, 1)
        )
        self.assertEqual(dates, ["2026-10-23"])

    def test_xaut_settles_later_than_btc(self):
        products = self._products([
            "2026-10-22T16:00:00Z",
            "2026-10-23T16:00:00Z",
        ])
        dates = self.delta_main.available_expiry_dates(
            "XAUT", products, now=self._at(13, 0)
        )
        self.assertEqual(dates, ["2026-10-22", "2026-10-23"])

    def test_an_expired_expiry_is_never_offered(self):
        products = self._products([
            "2026-10-21T12:00:00Z",
            "2026-10-22T12:00:00Z",
        ])
        dates = self.delta_main.available_expiry_dates(
            "BTC", products, now=self._at(8, 30)
        )
        self.assertEqual(dates, ["2026-10-22"])

    def test_the_settlement_hour_comes_from_the_products(self):
        # A product that settles at a different hour must be honoured, not
        # overridden by a hard-coded number.
        products = self._products([
            "2026-10-22T09:45:00Z",
            "2026-10-23T09:45:00Z",
        ])
        self.assertEqual(
            self.delta_main.available_expiry_dates(
                "BTC", products, now=self._at(9, 50)
            ),
            ["2026-10-23"],
        )


class DeltaOiNumericsTests(unittest.TestCase):
    """OI columns must be numeric so client VLOOKUP/XLOOKUP formulas work."""

    def test_fmt_oi_returns_float_not_string(self):
        self.assertIsInstance(delta_main.fmt_oi(1500000), float)
        self.assertEqual(delta_main.fmt_oi(1500000), 1500000.0)

    def test_fmt_oi_handles_legacy_money_strings(self):
        # Old builds wrote "$1.50M", "$2.00K", "$100", "Rs 500"
        self.assertEqual(delta_main.fmt_oi("$1.50M"), 1500000.0)
        self.assertEqual(delta_main.fmt_oi("$2.00K"), 2000.0)
        self.assertEqual(delta_main.fmt_oi("$100"), 100.0)
        self.assertEqual(delta_main.fmt_oi("Rs 500"), 500.0)
        self.assertEqual(delta_main.fmt_oi("1,234"), 1234.0)

    def test_fmt_oi_handles_none_and_nan(self):
        self.assertEqual(delta_main.fmt_oi(None), 0.0)
        self.assertEqual(delta_main.fmt_oi(float("nan")), 0.0)

    def test_oI_columns_use_numeric_format_not_text(self):
        src = inspect.getsource(delta_main.force_usd_block_formats)
        # OI columns (offsets 3, 5) must use numeric format, not "@"
        self.assertIn('NumberFormat = "#,##0"', src)
        self.assertNotIn('NumberFormat = "@"',
                         src.split('for off in (3, 5)')[1].split('_usd_fmt_done')[0])


class DeltaParitySpotTests(unittest.TestCase):
    """Parity spot fills the gap when the external spot feed goes stale."""

    def test_parity_spot_returns_correct_value(self):
        """S = K*exp(-(r-q)T) + (C-P)*exp(qT).  With q=0: S = K*exp(-rT) + (C-P)."""
        expiry_iso = "2027-12-31"
        T = delta_main._years_to_expiry(expiry_iso)
        self.assertIsNotNone(T)
        self.assertGreater(T, 0)

        # Set up a call/put pair at strike 50000
        delta_state.live_data["BTC"].clear()
        delta_state.live_data["BTC"]["C-BTC-50000"] = {
            "strike": 50000, "type": "call", "bid": 1000, "ask": 1010, "ltp": 1005
        }
        delta_state.live_data["BTC"]["P-BTC-50000"] = {
            "strike": 50000, "type": "put", "bid": 800, "ask": 810, "ltp": 805
        }
        delta_state.symbol_expiry["BTC"] = {"C-BTC-50000": expiry_iso, "P-BTC-50000": expiry_iso}
        delta_state.active_expiries["BTC"] = [expiry_iso]

        r = delta_state.PARITY_RATE
        expected = 50000 * math.exp(-r * T) + 200.0

        result = delta_main.parity_spot_or_none("BTC")
        self.assertIsNotNone(result)
        self.assertAlmostEqual(result, expected, places=0)

    def test_parity_spot_uses_median_of_n_strikes(self):
        """Multiple strikes: median, not mean."""
        expiry_iso = "2027-12-31"
        delta_state.live_data["ETH"].clear()
        delta_state.live_data["ETH"].update({
            f"C-ETH-{s}": {"strike": s, "type": "call", "bid": 100, "ask": 102, "ltp": 101}
            for s in [10000, 11000, 12000, 13000, 14000]
        })
        delta_state.live_data["ETH"].update({
            f"P-ETH-{s}": {"strike": s, "type": "put", "bid": 50, "ask": 52, "ltp": 51}
            for s in [10000, 11000, 12000, 13000, 14000]
        })
        for s in [10000, 11000, 12000, 13000, 14000]:
            delta_state.symbol_expiry["ETH"][f"C-ETH-{s}"] = expiry_iso
            delta_state.symbol_expiry["ETH"][f"P-ETH-{s}"] = expiry_iso
        delta_state.active_expiries["ETH"] = [expiry_iso]

        result = delta_main.parity_spot_or_none("ETH")
        self.assertIsNotNone(result)
        # With symmetric spreads, all parity estimates are similar
        self.assertGreater(result, 0)

    def test_parity_deadband_smooths_cache(self):
        """Changes smaller than deadband return the cached value."""
        expiry_iso = "2027-12-31"
        delta_state.live_data["BTC"].clear()
        # Two strikes, both at parity
        for s in [50000, 51000]:
            delta_state.live_data["BTC"][f"C-BTC-{s}"] = {
                "strike": s, "type": "call", "bid": 1000, "ask": 1002, "ltp": 1001
            }
            delta_state.live_data["BTC"][f"P-BTC-{s}"] = {
                "strike": s, "type": "put", "bid": 800, "ask": 802, "ltp": 801
            }
            delta_state.symbol_expiry["BTC"][f"C-BTC-{s}"] = expiry_iso
            delta_state.symbol_expiry["BTC"][f"P-BTC-{s}"] = expiry_iso
        delta_state.active_expiries["BTC"] = [expiry_iso]

        first = delta_main.parity_spot_or_none("BTC")
        self.assertIsNotNone(first)
        # Now change one quote slightly — the new parity will differ from
        # cached, but if within deadband the cache wins.
        delta_state.live_data["BTC"]["C-BTC-50000"]["bid"] = 1000.01  # tiny bump
        second = delta_main.parity_spot_or_none("BTC")
        self.assertEqual(second, first, "Deadband should suppress small changes")

    def test_parity_returns_none_without_contracts(self):
        delta_state.live_data["XAUT"].clear()
        self.assertIsNone(delta_main.parity_spot_or_none("XAUT"))


class DeltaCalcRetryCooldownTests(unittest.TestCase):
    """Session-level cooldown prevents COM hammering when Excel won't switch modes."""

    def test_cooldown_constants_exist(self):
        self.assertTrue(hasattr(delta_state, "CALC_RETRY_COOLDOWN_SECONDS"))
        self.assertTrue(hasattr(delta_state, "_calc_failed_at"))
        self.assertTrue(hasattr(delta_state, "_calc_manual_failed"))

    def test_cooldown_logic_in_bulk_mode(self):
        src = inspect.getsource(delta_main._begin_excel_bulk_mode)
        self.assertIn("CALC_RETRY_COOLDOWN_SECONDS", src)
        self.assertIn("_calc_manual_failed", src)
        self.assertIn("calc_manual_unavailable", src)




class DeltaCalcModeCoexistenceTests(unittest.TestCase):
    """When both tools start simultaneously, the second to acquire the guard
    must skip the Excel mode switch entirely instead of failing and showing
    an error dialog."""


    def test_other_holder_alive_check_is_in_bulk_mode(self):
        src = inspect.getsource(delta_main._begin_excel_bulk_mode)
        self.assertIn("other_holder_alive", src)

    def test_other_holder_alive_guards_before_acquire(self):
        """When other_holder_alive returns True, state[calc] is None
        and calc_mode_guard.acquire must not be called."""
        from unittest.mock import patch, MagicMock
        app = MagicMock()
        app.api.Calculation = -4105  # automatic
        app.screen_updating = True
        delta_state._app_cache = app
        delta_state._calc_manual_failed = False
        with patch.object(delta_main.calc_mode_guard, 'other_holder_alive', return_value=True):
            with patch.object(delta_main.calc_mode_guard, 'acquire') as mock_acquire:
                restore = delta_main._begin_excel_bulk_mode()
                self.assertIsNotNone(restore)
                self.assertFalse(mock_acquire.called,
                                 'acquire should not be called when another tool holds the guard')

    def test_no_other_holder_proceeds_normally(self):
        """When no other holder is alive, the guard is acquired normally."""
        from unittest.mock import patch, MagicMock
        app = MagicMock()
        app.api.Calculation = -4105  # automatic
        app.screen_updating = True
        delta_state._app_cache = app
        delta_state._calc_manual_failed = False
        with patch.object(delta_main.calc_mode_guard, 'other_holder_alive', return_value=False):
            with patch.object(delta_main.calc_mode_guard, 'acquire') as mock_acquire:
                restore = delta_main._begin_excel_bulk_mode()
                self.assertIsNotNone(restore)
                self.assertTrue(mock_acquire.called,
                                'acquire should be called when no other tool holds the guard')


class ChainsCloseSuppressionTests(unittest.TestCase):
    """Closing the chain workbooks must never put the calc dialog on screen.

    Excel's save prompt refuses every COM call for a few seconds while
    EXCEL.EXE is still alive, and that refusal used to fire the dialog on
    the first cycle. The write window is now skipped once the chains are
    gone, and the dialog only fires after a refusal has held for the grace
    period while a chain workbook is still open.
    """

    def setUp(self):
        self._saved = (
            delta_state._calc_manual_failed,
            delta_state._calc_failed_since,
            delta_state._calc_notice_shown,
            delta_state._calc_failed_at,
            delta_state._app_cache,
            delta_state._bulk_owning_manual,
        )

    def tearDown(self):
        restore = self._saved
        (
            delta_state._calc_manual_failed,
            delta_state._calc_failed_since,
            delta_state._calc_notice_shown,
            delta_state._calc_failed_at,
            delta_state._app_cache,
            delta_state._bulk_owning_manual,
        ) = restore

    def _app_with_books(self, names, refusing=True):
        from unittest.mock import MagicMock

        books = []
        for name in names:
            book = MagicMock()
            book.name = name
            books.append(book)

        class _Api:
            Calculation = -4105

        class App:
            version = "16.0"

            def __init__(inner):
                inner.api = _Api()
                inner._screen = True
                inner.books = books

            @property
            def screen_updating(inner):
                return inner._screen

            @screen_updating.setter
            def screen_updating(inner, value):
                inner._screen = value

            @property
            def calculation(inner):
                return "automatic"

            @calculation.setter
            def calculation(inner, value):
                if value == "manual" and refusing:
                    raise RuntimeError(
                        "Unable to set the Calculation property of the Application class"
                    )

        return App()

    def _run_cycle(self, app, failed_for=None):
        import time
        import unittest.mock as mock

        delta_state._app_cache = app
        delta_state._calc_failed_at = time.monotonic() - 60.0
        delta_state._bulk_owning_manual = False
        if failed_for is None:
            delta_state._calc_manual_failed = False
            delta_state._calc_failed_since = 0.0
        else:
            delta_state._calc_manual_failed = True
            delta_state._calc_failed_since = time.monotonic() - failed_for
        with mock.patch.object(
            delta_main.calc_mode_guard, "other_holder_alive", return_value=False
        ), mock.patch.object(
            delta_main.calc_mode_guard, "acquire"
        ), mock.patch.object(
            delta_main.calc_mode_guard, "release"
        ), mock.patch.object(
            delta_main, "excel_process_running", return_value=True
        ), mock.patch.object(
            delta_main, "show_error"
        ) as show_error:
            restore = delta_main._begin_excel_bulk_mode()
        return restore, show_error

    def test_closing_the_chains_clears_the_refusal_without_dialog(self):
        app = self._app_with_books(["StrategyBook.xlsx"])
        restore, show_error = self._run_cycle(app, failed_for=600.0)
        self.assertIsNone(restore)
        show_error.assert_not_called()
        self.assertFalse(delta_state._calc_manual_failed)
        self.assertEqual(delta_state._calc_failed_since, 0.0)

    def test_a_short_blip_runs_silent(self):
        app = self._app_with_books(["btc_chain.xlsx"])
        restore, show_error = self._run_cycle(app, failed_for=30.0)
        self.assertIsNotNone(restore)
        show_error.assert_not_called()
        self.assertFalse(delta_state._calc_notice_shown)
        self.assertTrue(delta_state._calc_manual_failed)

    def test_a_stuck_refusal_earns_the_dialog_first_try(self):
        app = self._app_with_books(["btc_chain.xlsx", "eth_chain.xlsx"])
        restore, show_error = self._run_cycle(app, failed_for=90.0)
        self.assertIsNotNone(restore)
        self.assertEqual(show_error.call_count, 1)
        self.assertTrue(delta_state._calc_notice_shown)

    def test_an_open_chain_keeps_the_window_quiet(self):
        app = self._app_with_books(["eth_chain.xlsx"], refusing=False)
        restore, show_error = self._run_cycle(app)
        self.assertIsNotNone(restore)
        show_error.assert_not_called()

    def test_close_skips_the_guard_entirely(self):
        from unittest.mock import patch

        app = self._app_with_books(["StrategyBook.xlsx"])
        with patch.object(delta_main.calc_mode_guard, "acquire") as acquire:
            delta_state._app_cache = app
            result = delta_main._begin_excel_bulk_mode()
        self.assertIsNone(result)
        self.assertFalse(acquire.called)


class ChainSavePromptAnswererTests(unittest.TestCase):
    """The tool answers Excel's Save / Don't Save prompt over chain books.

    A customer closing btc/eth/xaut_chain would otherwise leave that prompt
    sitting in the background: it blocks every COM call until it is answered,
    which is also how the refused manual-calculation switches start. Nothing
    unsaved in a chain workbook is worth keeping - every cell is rewritten
    next cycle - so Don't Save is safe there, and only there.
    """

    def test_the_answerer_runs_off_the_render_thread(self):
        import inspect
        import main as delta_main

        src = inspect.getsource(delta_main.display_loop)
        self.assertIn("start_chain_save_prompt_watchdog()", src)
        starter = inspect.getsource(delta_main.start_chain_save_prompt_watchdog)
        self.assertIn("_prompt_watchdog_started", starter)
        self.assertIn('platform.system() != "Windows"', starter)
        self.assertIn("threading.Thread(", starter)
        self.assertIn("daemon=True", starter)

    def test_only_chain_windows_are_answered(self):
        import unittest.mock as mock

        import main as delta_main

        fake = _FakeUser32("BTC | Options Chain | LIVE - Excel")
        with mock.patch("ctypes.windll") as windll:
            windll.user32 = fake.user32
            self.assertTrue(
                delta_main._prompt_belongs_to_chain(0x1234),
                "a prompt over the rendered chain frame must be answered",
            )

        fake = _FakeUser32("btc_chain - Excel")
        with mock.patch("ctypes.windll") as windll:
            windll.user32 = fake.user32
            # file captions join words with underscores
            self.assertTrue(
                delta_main._prompt_belongs_to_chain(0x1234),
                "the file-backed caption must be recognised too",
            )

        fake = _FakeUser32("Research Budget 2026 - Excel")
        with mock.patch("ctypes.windll") as windll:
            windll.user32 = fake.user32
            self.assertFalse(
                delta_main._prompt_belongs_to_chain(0x1234),
                "the customer's own workbook keeps its prompt untouched",
            )

    def test_the_modern_prompt_is_answered_without_fighting_window_shapes(self):
        import inspect
        import main as delta_main

        src = inspect.getsource(delta_main._answer_save_prompt)
        self.assertIn('FindWindowExW(hwnd, button, "Button", None)', src)
        self.assertIn("0x00F5", src)
        self.assertIn("_send_alt_n(hwnd)", src)
        accel = inspect.getsource(delta_main._send_alt_n)
        # The stuck-Alt incident rules: no focus stealing from a background
        # prompt, the Alt-UP fires on every path, and the flight flag lets
        # the exit paths wait for the chord before os._exit.
        self.assertIn("_prompt_is_foreground(hwnd)", accel)
        self.assertNotIn("SetForegroundWindow(hwnd)", accel)
        self.assertEqual(delta_main._VK_MENU, 0x12)
        self.assertEqual(delta_main._VK_N, 0x4E)
        self.assertIn("keybd_event(_VK_MENU, 0, 0, 0)", accel)
        self.assertIn("keybd_event(_VK_N, 0, 0, 0)", accel)
        self.assertIn("finally:", accel)
        self.assertIn("keybd_event(_VK_MENU, 0, _KEYUP, 0)", accel)
        self.assertIn("_alt_chord_active", accel)

    def test_a_stuck_prompt_is_never_hammered_forever(self):
        import inspect
        import main as delta_main

        self.assertEqual(delta_main._SAVE_PROMPT_CLASSES, ("NUIDialog", "#32770"))
        src = inspect.getsource(delta_main._chain_save_prompt_once)
        self.assertIn("_SAVE_PROMPT_ATTEMPT_LIMIT", src)
        self.assertIn("attempts[hwnd] = count + 1", src)
        self.assertIn("feed_log(", src)
        looper = inspect.getsource(delta_main._chain_save_prompt_watchdog)
        self.assertIn("time.sleep(", looper)
        self.assertIn("except Exception", looper)

    def test_the_question_only_travels_to_chain_owned_dialogs(self):
        import inspect
        import main as delta_main

        self.assertEqual(delta_main._SAVE_PROMPT_ATTEMPT_LIMIT, 8)
        self.assertGreater(delta_main._SAVE_PROMPT_POLL_SECONDS, 0.0)
        src = inspect.getsource(delta_main._prompt_belongs_to_chain)
        self.assertIn("GetAncestor", src)
        # underscore captions like btc_chain must match too
        self.assertIn('.replace("_", " ")', src)
        self.assertIn("for token in _DELTA_CHAIN_CAPTION_TOKENS", src)


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


class KeyboardModifierArmorTests(unittest.TestCase):
    """Nothing may leave a modifier stuck system-wide - ever again.

    A customer's whole PC acting like Alt was held down came from the old
    _send_alt_n: Alt down, four sleeps, Alt up, and an exit in between kept
    the modifier pressed across every desktop app (typed letters became
    ribbon accelerators, ESC became Alt+ESC and minimised windows, menus
    popped open). The pins below hold the rules: no chord to a background
    prompt, the Alt-UP on every code path, healing at startup and before
    every hard exit, and a ten-second stuck-modifier repair.
    """

    def test_a_background_prompt_gets_no_chord(self):
        from unittest.mock import patch
        import main as delta_main

        class FakeUser32:
            def __init__(self):
                self.fg = 0x777
                self.calls = []

            def GetForegroundWindow(self):
                return self.fg

            def GetAncestor(self, hwnd, flag):
                return 0

            def keybd_event(self, vk, scan, flags, extra):
                self.calls.append((vk, flags))

        fake = FakeUser32()
        with patch("ctypes.windll") as windll:
            windll.user32 = fake
            self.assertFalse(delta_main._send_alt_n(0x200))
            self.assertEqual(
                fake.calls, [], "no synthetic keys into a background prompt"
            )

    def test_the_chord_releases_alt_on_every_path(self):
        from unittest.mock import patch
        import main as delta_main

        class FakeUser32:
            def __init__(self, raise_on_n=False):
                self.fg = 0x200
                self.calls = []
                self.raise_on_n = raise_on_n

            def GetForegroundWindow(self):
                return self.fg

            def GetAncestor(self, hwnd, flag):
                return 0

            def keybd_event(self, vk, scan, flags, extra):
                if self.raise_on_n and vk == delta_main._VK_N and flags == 0:
                    raise OSError("synthetic failure")
                self.calls.append((vk, flags))

        fake = FakeUser32()
        with patch("ctypes.windll") as windll:
            windll.user32 = fake
            self.assertTrue(delta_main._send_alt_n(0x200))
        # The chord ends with an Alt release, always.
        self.assertEqual(fake.calls[-1], (delta_main._VK_MENU, delta_main._KEYUP))

        fake = FakeUser32(raise_on_n=True)
        with patch("ctypes.windll") as windll:
            windll.user32 = fake
            self.assertFalse(delta_main._send_alt_n(0x200))
        # Even a failing chord ends with the Alt-UP the desktop needs.
        self.assertEqual(fake.calls[-1], (delta_main._VK_MENU, delta_main._KEYUP))

    def test_heal_releases_menu_and_win_keys(self):
        from unittest.mock import patch
        import main as delta_main

        class FakeUser32:
            def __init__(self):
                self.calls = []

            def keybd_event(self, vk, scan, flags, extra):
                self.calls.append((vk, flags))

        fake = FakeUser32()
        with patch("ctypes.windll") as windll:
            windll.user32 = fake
            delta_main._heal_keyboard()
        self.assertIn((0x12, 2), fake.calls)
        self.assertIn((0x5B, 2), fake.calls)
        self.assertIn((0x5C, 2), fake.calls)

    def test_startup_heals_and_hard_exits_heal(self):
        import inspect
        import main as delta_main

        self.assertIn("_heal_keyboard()", inspect.getsource(delta_main.main_async))
        maybe_src = inspect.getsource(delta_main.maybe_exit_after_close_all)
        self.assertIn("_heal_keyboard()", maybe_src)
        self.assertIn("os._exit(0)", maybe_src)
        # The heal precedes the exit, so a mid-chord kill cannot strand Alt.
        self.assertLess(
            maybe_src.index("_heal_keyboard()"), maybe_src.index("os._exit(0)")
        )
        wl_src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn("_heal_keyboard()", wl_src)
        dl_src = inspect.getsource(delta_main.display_loop)
        self.assertIn("_heal_keyboard()", dl_src)
        self.assertIn("_watch_stuck_modifiers_tick()", dl_src)
        self.assertIn("_heal_keyboard()", inspect.getsource(delta_main.main))

    def test_a_ten_second_stuck_alt_is_released(self):
        from unittest.mock import patch
        import time
        import main as delta_main

        class FakeUser32:
            def __init__(self):
                self.calls = []

            def GetAsyncKeyState(self, vk):
                return 0x8000

            def keybd_event(self, vk, scan, flags, extra):
                self.calls.append((vk, flags))

        fake = FakeUser32()
        delta_main._alt_modifiers_held_since.clear()
        delta_main._alt_chord_active.clear()
        try:
            with patch("ctypes.windll") as windll:
                windll.user32 = fake
                # First sight only records - no release inside the grace.
                delta_main._watch_stuck_modifiers_tick()
                self.assertEqual(fake.calls, [])
                # ...ten seconds of the same key later: repaired.
                delta_main._alt_modifiers_held_since[(0x12,)] = (
                    time.monotonic() - 11.0
                )
                delta_main._watch_stuck_modifiers_tick()
                self.assertIn((0x12, 2), fake.calls)
                self.assertNotIn((0x12,), delta_main._alt_modifiers_held_since)
        finally:
            delta_main._alt_modifiers_held_since.clear()
            delta_main._alt_chord_active.clear()


class ScreenUpdatingSelfHealTests(unittest.TestCase):
    """The tool itself must never touch ScreenUpdating - and never stick off.

    Toggling the screen around every write window made Excel repaint the
    whole sheet on each release: at a 0.2 s cadence that reads as constant
    flicker, and a click into a cell inside a window froze the rest of the
    edit on a stale frame (typing a formula showed nothing). The toggle-era
    pins are gone; what remains: (1) opening and closing the calculation
    session never reads or writes the screen, (2) the watchdog is the only
    net that pushes paint back on, and only from outside a write window,
    (3) a screen left off by anything else - an old build, or an add-in -
    gets fixed by that net.
    """

    def _screen_touch_counter_app(self, screen=False):
        class App:
            def __init__(self):
                self.screen = screen
                self.touches = []

            @property
            def version(self):
                return "16.0"

            @property
            def books(self):
                return [type("B", (), {"name": "btc_chain.xlsx"})()]

            @property
            def screen_updating(self):
                self.touches.append("read")
                return self.screen

            @screen_updating.setter
            def screen_updating(self, value):
                self.touches.append("write")
                self.screen = value

            @property
            def api(self):
                # MANUAL (-4135): the calculation session keeps the customer's
                # own choice, which keeps this stub off the guard entirely.
                return type("Api", (), {"Calculation": -4135})()

        return App()

    def test_opening_and_closing_the_session_never_touches_the_screen(self):
        app = self._screen_touch_counter_app(screen=True)
        delta_state._app_cache = app
        delta_state._bulk_window_open = False
        try:
            restore = delta_main._begin_excel_bulk_mode()
            self.assertIsNotNone(restore)
            self.assertTrue(delta_state._bulk_window_open)
            restore()
            self.assertEqual(
                app.touches,
                [],
                "ScreenUpdating was read or written; that is the flicker bug",
            )
            self.assertTrue(app.screen)
            self.assertFalse(delta_state._bulk_window_open)
        finally:
            delta_state._bulk_window_open = False
            delta_state._app_cache = None

    def test_a_screen_left_off_is_untouched_here_and_netted_by_the_watchdog(self):
        app = self._screen_touch_counter_app(screen=False)
        delta_state._app_cache = app
        restore = delta_main._begin_excel_bulk_mode()
        restore()
        self.assertFalse(app.screen, "the write session has no business painting")
        # Now the net, outside any window.
        delta_state._screen_restore_owed = True
        delta_state._screen_watchdog_logged_at = 0.0
        delta_state._bulk_window_open = False
        try:
            delta_main._screen_update_watchdog_tick()
            self.assertTrue(app.screen)
        finally:
            delta_state._screen_restore_owed = False
            delta_state._app_cache = None

    def test_the_watchdog_finishes_a_failed_restore(self):
        import delta_state
        import main as delta_main

        class StickyApp:
            def __init__(self):
                self.screen = False
                self.attempts = []

            @property
            def version(self):
                return "16.0"

            @property
            def screen_updating(self):
                return self.screen

            @screen_updating.setter
            def screen_updating(self, value):
                self.screen = value
                self.attempts.append(value)

        app = StickyApp()
        delta_state._app_cache = app
        delta_state._owner_app = app
        delta_state._screen_restore_owed = True
        delta_state._bulk_window_open = False
        delta_state._screen_watchdog_logged_at = 0.0
        try:
            delta_main._screen_update_watchdog_tick()
            self.assertTrue(app.screen)
            self.assertFalse(delta_state._screen_restore_owed)
        finally:
            delta_state._screen_restore_owed = False

    def test_the_watchdog_stays_out_of_the_write_window(self):
        import delta_state
        import main as delta_main

        class App:
            @property
            def screen_updating(self):
                raise AssertionError("watchdog probed an open write window")

        app = App()
        delta_state._app_cache = app
        delta_state._bulk_window_open = True
        try:
            delta_main._screen_update_watchdog_tick()
        finally:
            delta_state._bulk_window_open = False


class RenderYieldTests(unittest.TestCase):
    """The chain pauses only while the customer's hands are on OUR Excel.

    Typing a formula into a repaint-numbed frame is what produced "I typed
    and nothing is showing". Full speed stays the rule; a cycle is skipped
    only for fresh input in our own instance, and what that skip owes the
    recalculation is carried, never dropped.
    """

    def test_our_excel_in_front_with_fresh_input_yields(self):
        from unittest.mock import patch
        import main as delta_main

        with patch.object(delta_main, "_last_user_input_age", return_value=0.1), \
             patch.object(delta_main, "_foreground_pid", return_value=4242), \
             patch.object(delta_main, "_owner_window_pid", return_value=4242):
            self.assertTrue(delta_main._user_is_excel_interacting())

    def test_another_app_in_front_never_yields(self):
        from unittest.mock import patch
        import main as delta_main

        with patch.object(delta_main, "_last_user_input_age", return_value=0.1), \
             patch.object(delta_main, "_foreground_pid", return_value=4242), \
             patch.object(delta_main, "_owner_window_pid", return_value=9999):
            self.assertFalse(delta_main._user_is_excel_interacting())

    def test_older_input_and_unresolvable_windows_do_not_yield(self):
        from unittest.mock import patch
        import main as delta_main

        with patch.object(delta_main, "_last_user_input_age", return_value=2.0), \
             patch.object(delta_main, "_foreground_pid", return_value=4242), \
             patch.object(delta_main, "_owner_window_pid", return_value=4242):
            self.assertFalse(delta_main._user_is_excel_interacting())

        with patch.object(delta_main, "_last_user_input_age", return_value=None), \
             patch.object(delta_main, "_foreground_pid", return_value=None), \
             patch.object(delta_main, "_owner_window_pid", return_value=None):
            self.assertFalse(delta_main._user_is_excel_interacting())

    def test_render_cycle_yields_before_the_write_window(self):
        import inspect
        import main as delta_main

        src = inspect.getsource(delta_main.render_cycle)
        self.assertIn("_user_is_excel_interacting()", src)
        self.assertLess(
            src.index("_user_is_excel_interacting()"),
            src.index("_begin_excel_bulk_mode()"),
        )
        self.assertIn("render_cycle_skipped_user_input", src)

    def test_a_yielded_cycle_carries_the_recalc(self):
        import delta_state
        import main as delta_main

        delta_state._recalc_owed = False
        delta_state._bulk_will_recalc = True
        delta_state._app_cache = None
        # Nothing written, nothing owed: plain skip.
        delta_main._time_recalc(False)
        self.assertFalse(delta_state._recalc_owed)
        # Owed but Excel absent: the debt survives and is paid later.
        delta_state._recalc_owed = True
        delta_main._time_recalc(False)
        self.assertTrue(delta_state._recalc_owed)

    def test_an_owed_recalc_survives_a_busy_excel(self):
        import delta_state
        import main as delta_main

        class BusyRecalcApp:
            @property
            def api(self):
                return self

            Calculation = -4105

            def Calculate(self):
                raise RuntimeError("Excel is busy mid-recalc")

        app = BusyRecalcApp()
        delta_state._app_cache = app
        delta_state._bulk_will_recalc = True
        try:
            with patch.object(delta_main, "log_excel_exception"):
                # Nothing written this cycle - but a debt was carried in
                # from a yielded one. A busy Excel must not eat it.
                delta_state._recalc_owed = True
                delta_main._time_recalc(False)
                self.assertTrue(
                    delta_state._recalc_owed,
                    "the skip promised this recalculation; retry next cycle",
                )
                # A written cycle owns its own recalc: if it fails, that
                # failure is one failed write, not an unpaid debt.
                delta_state._recalc_owed = False
                delta_main._time_recalc(True)
                self.assertFalse(delta_state._recalc_owed)
        finally:
            delta_state._recalc_owed = False
            delta_state._bulk_will_recalc = False
            delta_state._app_cache = None


class StartupConsoleNoiseTests(unittest.TestCase):
    """Python 3.10 prints a bogus asyncio traceback on most launches.

    _ProactorBaseWritePipeTransport._loop_writing races a finished write
    and raises `assert f is self._write_fut` inside asyncio itself; the
    customer sees a red console traceback under a chain that took a day of
    work. Startup wires a loop exception handler that drains exactly that
    one pattern and hands every other exception to the default handler
    untouched.
    """

    def test_main_wires_the_quiet_handler(self):
        import inspect
        import delta_state
        import main as delta_main

        src = inspect.getsource(delta_main.main_async)
        self.assertIn("set_exception_handler(_drop_proactor_write_noise)", src)

    def test_only_the_write_race_is_drained(self):
        import delta_state
        import main as delta_main

        class _Loop:
            def __init__(self):
                self.default_calls = []

            def default_exception_handler(self, context):
                self.default_calls.append(context)

        loop = _Loop()
        delta_main._drop_proactor_write_noise(
            loop,
            {
                "exception": AssertionError("assert f is self._write_fut"),
                "handle": "<Handle _ProactorBaseWritePipeTransport._loop_writing()>",
            },
        )
        self.assertEqual(loop.default_calls, [])

        real_problem = {
            "exception": OSError("pipe failed for real"),
            "handle": "<Handle _ProactorBaseWritePipeTransport._loop_writing()>",
        }
        delta_main._drop_proactor_write_noise(loop, real_problem)
        self.assertEqual(loop.default_calls, [real_problem])

        other_assert = {
            "exception": AssertionError("isolated assert"),
            "handle": "<Handle somewhere.else>",
        }
        delta_main._drop_proactor_write_noise(loop, other_assert)
        self.assertEqual(loop.default_calls, [real_problem, other_assert])


class ChainExtentStabilityTests(unittest.TestCase):
    """The customer's strategy formulas carry literal row ranges.

    The complaint: their own sheets (VLOOKUP/XLOOKUP against the chain
    columns) showed values that were wrong and lagging. Root cause one: a
    chain whose live row count changes with every expiry walks rows out from
    under those literal ranges, so one leg of the lookup silently becomes the
    if_not_found default. The write path therefore pads every block to a
    fixed extent, and the refresh widens adaptively when the customer's own
    workbook is heavy.
    """

    def test_blocks_write_a_fixed_extent(self):
        import inspect
        src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn("ROW_EXTENT_DEFAULT", src)
        self.assertIn("chain_extent_changed", src)
        self.assertIn("n_rows = extent", src)

    def test_flush_writes_padded_rows_not_the_dataframe(self):
        import inspect
        src = inspect.getsource(delta_main._flush_block_writes)
        self.assertNotIn("chain.iloc", src)
        self.assertIn("rows_block", src)

    def test_write_path_pads_before_diffing(self):
        import inspect
        src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn("pad_block_rows(values, extent)", src)
        self.assertIn("pending_writes.append(", src)

    def test_pad_block_rows_pads_with_blank_rows(self):
        self.assertEqual(
            delta_main.pad_block_rows([[1, 2], [3, 4]], 4),
            [[1, 2], [3, 4], [None, None], [None, None]],
        )
        self.assertEqual(
            delta_main.pad_block_rows([], 2), [[None] * 9, [None] * 9]
        )
        self.assertEqual(delta_main.pad_block_rows([[1, 2]], 1), [[1, 2]])

    def test_extent_is_not_shrunk_after_growing(self):
        """A stuck-at-max keeps every literal customer range valid forever."""
        import inspect
        src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn(
            "extent = max(delta_state.ROW_EXTENT_DEFAULT, seen, len(values))", src
        )
        self.assertNotIn("_row_extent_seen[(asset, bi)] = len(values)", src)

    def test_deep_clear_sweeps_legacy_junk_once_per_block(self):
        import inspect
        src = inspect.getsource(delta_main._deep_clear_below_block)
        self.assertIn("_deep_swept", src)
        self.assertIn("chain_deep_clear", src)

    def test_every_asset_shows_two_blocks(self):
        """BTC and ETH join XAUT on two expiry blocks per sheet.

        The customer asked for two expiries per sheet everywhere (their
        strategy books read blocks 1+2 only); the third expiry columns get
        cleared and hidden by the existing bi >= nb machinery.
        """
        for asset in ("BTC", "ETH", "XAUT"):
            self.assertEqual(delta_main.n_blocks(asset), 2)
        self.assertEqual(delta_main.BLOCK_COLS[2]["start"], 23)

    def test_third_block_wipes_on_connect_not_on_first_write(self):
        """The sweep must ride the per-cycle connect step, not the first
        data write - a cold start takes minutes and the customer stares at
        the old Block 3 the whole time otherwise."""
        import inspect
        src = inspect.getsource(delta_main._retire_unused_blocks)
        self.assertIn("expiry_block_retired", src)
        self.assertIn("_block_hidden[(asset, bi)] = True", src)
        self.assertLess(
            src.index("get_column_letter(block[\"start\"])"),
            src.index("api.EntireColumn.Hidden"),
        )
        writer = inspect.getsource(delta_main._write_to_excel)
        self.assertIn("_retire_unused_blocks(sheet, asset)", writer)
        self.assertLess(
            writer.index("_retire_unused_blocks(sheet, asset)"),
            writer.index("if not sheet_structure_ok(sheet)"),
        )


class DeltaRefreshWidenTests(unittest.TestCase):
    """When Excel still owes a recalc after our window the loop stretches."""

    def setUp(self):
        self._saved = (
            delta_state._refresh_widen,
            delta_state._refresh_widen_hold_until,
            delta_state._calc_state_probe,
            delta_state._app_cache,
        )

    def tearDown(self):
        delta_state._refresh_widen = self._saved[0]
        delta_state._refresh_widen_hold_until = self._saved[1]
        delta_state._calc_state_probe = self._saved[2]
        delta_state._app_cache = self._saved[3]

    def test_second_pass_widens_refresh(self):
        class _App:
            class api:
                CalculationState = 2

        delta_state._refresh_widen = 1.0
        delta_state._calc_state_probe = None
        delta_state._app_cache = _App()
        delta_main._probe_calc_state_after_restore()
        self.assertEqual(delta_state._refresh_widen, 1.5)
        self.assertGreater(delta_state._refresh_widen_hold_until, 0.0)

    def test_widen_caps_at_four_and_holds_through_the_window(self):
        delta_state._refresh_widen = 3.0
        held_at = time.monotonic() + 10.0
        delta_state._refresh_widen_hold_until = held_at
        delta_main._note_recalc_second_pass(time.monotonic())
        # 3.0 * 1.5 = 4.5 caps at WIDEN_MAX.
        self.assertEqual(delta_state._refresh_widen, 4.0)
        self.assertGreater(delta_state._refresh_widen_hold_until, held_at)

    def test_display_loop_uses_and_decays_widen(self):
        import inspect
        src = inspect.getsource(delta_main.display_loop)
        self.assertIn("float(refresh) * delta_state._refresh_widen", src)
        self.assertIn("refresh_widen_decayed", src)
        self.assertIn("_refresh_widen_hold_until", src)


class AutoOpenRecoveryTests(unittest.TestCase):
    """A wedged Workbooks.Open must time out, log, and retry once.

    The startup open pass used to be one silent daemon thread: Workbooks.Open
    aimed at an Excel another tool owned (or whose controller died earlier)
    can block forever, the pass never reported anything, and the writer sat
    on "excel_running_workbook_not_found" for the whole session - the
    customer sees a blank sheet that never fills. Every pass now runs inside
    (or creates) the tool's private instance; the retry simply re-runs the
    open once.
    """

    def test_open_pass_is_watchdog_gated(self):
        """The pass runs inline on the display thread, with one inline retry.

        An opener thread created the private Excel instance while the
        display loop read that App object from another COM apartment;
        every probe answered "marshalled for a different thread" and no
        workbook ever attached. Inline on the display thread is the
        contract now; a wedged pass is retried exactly once from the
        same thread after a grace period.
        """
        src = inspect.getsource(delta_main.open_excel_files)
        self.assertNotIn("threading.Thread", src)
        self.assertNotIn("finished.wait(", src)
        self.assertIn("_open_excel_files()", src)
        self.assertIn("_open_pass_at", src)
        retry_src = inspect.getsource(delta_main.retry_open_pass_after_grace)
        self.assertIn("_OPEN_RETRY_GRACE_SECONDS", retry_src)
        self.assertIn("excel_open_retried", retry_src)
        self.assertIn("open_excel_files()", retry_src)
        loop_src = inspect.getsource(delta_main.display_loop)
        self.assertIn("retry_open_pass_after_grace()", loop_src)

    def test_open_pass_reports_each_book(self):
        opener = inspect.getsource(delta_main._open_excel_files)
        self.assertIn("excel_auto_open_started", opener)
        self.assertIn('excel_auto_open_done", asset=asset', opener)
        self.assertIn("excel_auto_open_failed", opener)
        self.assertIn("finished.set()", opener)
        self.assertIn("own_instance", opener)

    def test_writer_never_reopens_a_closed_workbook(self):
        writer = inspect.getsource(delta_main._write_to_excel)
        self.assertNotIn("excel_auto_open_rearmed", writer)
        self.assertNotIn("delta_state._auto_open_rearm_started", writer)
        self.assertNotIn("open_excel_files()", writer)
        self.assertNotIn("os.startfile(", writer)
        self.assertNotIn("notice_workbook_missing", writer)
        self.assertNotIn("Microsoft Excel is open", writer)
        # A busy Excel is "unknown", never "missing".
        self.assertIn("excel_attach_unknown", writer)
        self.assertIn("workbook_closed_by_customer", writer)

    def test_open_pass_never_scans_foreign_excel(self):
        opener = inspect.getsource(delta_main._open_excel_files)
        self.assertNotIn("apps = list(xw.apps)", opener)
        self.assertNotIn("apps[0]", opener)
        self.assertIn("_ensure_owner_app", opener)

    def test_loop_thread_initializes_com_apartment(self):
        src = inspect.getsource(delta_main.main_async)
        self.assertIn("pythoncom.CoInitialize()", src)
        self.assertLess(
            src.index("pythoncom.CoInitialize()"),
            src.index("open_excel_files()"),
        )

    def test_default_expiries_are_two_wide(self):
        src = inspect.getsource(delta_main.get_default_expiries)
        self.assertIn("n_blocks(asset)", src)
        self.assertIn("dates.append(None)", src)

    def test_roller_pads_slot_lists_before_writes(self):
        src = inspect.getsource(delta_main.expiry_roller)
        self.assertIn("while len(active_expiries[asset]) <= i:", src)
        self.assertLess(
            src.index("while len(active_expiries[asset]) <= i:"),
            src.index("active_expiries[asset][i] = tgt"),
        )

    def test_chain_value_writes_anchor_at_row_five(self):
        src = inspect.getsource(delta_main._flush_block_writes)
        self.assertIn("sheet.range((5 + min_row, first_c)).value = grid", src)
        self.assertIn("sheet.range((5 + start, block[\"start\"])).value = [", src)
        self.assertIn("sheet.range((5 + lo, block[\"start\"])).value = merged_rows", src)
        # Reversed two-corner ranges (row 5..4) resolve to their top-left
        # and stomp the header row - they must not come back.
        self.assertNotIn("(4 + max_row, last_c)", src)
        self.assertNotIn("(4 + end, block[\"start\"] + 8)", src)
        self.assertNotIn("(4 + hi, block[\"start\"] + 8)", src)

    def test_stale_clear_only_when_below_live_rows(self):
        src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn("if last_row > n_rows:", src)

    def test_find_open_book_matches_short_form_fullnames(self):
        src = inspect.getsource(delta_main.find_open_book)
        self.assertIn("_long_form", src)
        self.assertIn("GetLongPathNameW", inspect.getsource(delta_main._long_form))
        # Only the private instance is searched, and the match normalises
        # the workbook's fullname through the Windows long-form API.
        self.assertIn("_long_form(fullname).lower() == target_long", src)
        self.assertNotIn("list(xw.apps)", src)
        self.assertNotIn("same_name", src)

    def test_long_form_round_trips_short_paths(self):
        import ctypes

        path = os.path.abspath(__file__)
        buffer = ctypes.create_unicode_buffer(1024)
        short = path
        if ctypes.windll.kernel32.GetShortPathNameW(path, buffer, 1024):
            short = buffer.value or path
        self.assertTrue(short and short.upper() != path.upper())
        self.assertEqual(
            os.path.abspath(delta_main._long_form(short)).lower(),
            path.lower(),
        )


if __name__ == "__main__":
    unittest.main()
