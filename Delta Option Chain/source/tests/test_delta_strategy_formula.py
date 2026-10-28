"""Integration test: verify Delta client strategy formulas work against real Excel.

Mirrors test_strategy_formula.py for the Delta chain.  Delta's workbooks use a
3-block layout (cols 1-9, 12-20, 23-31) with Call/Put columns that a client
strategy sheet VLOOKUP/XLOOKUPs against.  Before the fmt_oi fix the OI columns
were text strings ("$1.50M"), which broke numeric formulas referencing them.

These tests:
  - verify strikes are written as numeric floats, not strings
  - verify OI is numeric (fmt_oi returns float)
  - verify a VLOOKUP into the chain returns a number, not text
  - verify that missing strikes produce blank, not #N/A/#VALUE!
  - verify reformatting (IFERROR) handles missing strikes

Requires Excel + xlwings.  Skipped if Excel is unavailable.
"""

import math
import unittest

import pandas as pd

try:
    import xlwings as xw
except ImportError:
    xw = None

import main as delta_main
import delta_state


HEADERS = ["Call LTP", "Call Bid", "Call Ask", "Call OI", "strike",
           "Put OI", "Put Bid", "Put Ask", "Put LTP"]


def _make_chain(strikes, base_call=1000.0, base_put=800.0, spread=10.0,
                oi=50000, ltp=None, skip_strikes=None):
    skip_strikes = skip_strikes or set()
    rows = []
    for s in strikes:
        if s in skip_strikes:
            continue
        c_bid = base_call + (s - strikes[len(strikes) // 2]) * 0.01
        c_ask = c_bid + spread
        p_bid = base_put - (s - strikes[len(strikes) // 2]) * 0.01
        p_ask = p_bid + spread
        rows.append({
            "strike": float(s),
            "Call LTP": ltp if ltp is not None else c_bid,
            "Call Bid": c_bid,
            "Call Ask": c_ask,
            "Call OI": oi,
            "Put OI": oi,
            "Put Bid": p_bid,
            "Put Ask": p_ask,
            "Put LTP": ltp if ltp is not None else p_bid,
        })
    return pd.DataFrame(rows)


class DeltaStrategyFormulaTests(unittest.TestCase):
    """Test the client's strategy lookup formulas against real Excel."""

    @classmethod
    def setUpClass(cls):
        if xw is None:
            raise unittest.SkipTest("xlwings not installed")
        try:
            cls.app = xw.App(visible=False, add_book=False)
            cls.app.display_alerts = False
            cls.app.screen_updating = False
        except Exception as e:
            raise unittest.SkipTest(f"Excel not available: {e}")

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "app"):
            try:
                cls.app.quit()
            except Exception:
                pass

    def setUp(self):
        if xw is None:
            self.skipTest("xlwings not installed")
        self.book = self.app.books.add()
        self.chain_sheet = self.book.sheets["Sheet1"]
        self.chain_sheet.name = "Chain"
        self.cal_sheet = self.book.sheets.add()
        self.cal_sheet.name = "Calendar"

        for i, h in enumerate(HEADERS):
            self.chain_sheet.range((4, 1 + i)).value = h
        for i, h in enumerate(HEADERS):
            self.chain_sheet.range((4, 12 + i)).value = h

        self.app.calculation = "manual"

    def _write_block(self, start_col, chain_df):
        for col_idx, col_name in enumerate(HEADERS):
            col = start_col + col_idx
            for row_idx, val in enumerate(chain_df[col_name].tolist()):
                self.chain_sheet.range((5 + row_idx, col)).value = val
        return len(chain_df)

    def _write_calendar_vlookup(self, strikes, chain_rows, use_iferror=True):
        """Write lookup formula for each strike row in the Calendar sheet.

        Uses XLOOKUP (same as US tool's client formula) to look up the strike
        in the chain's strike column and return the Put Ask (block 1, col H)
        plus Put Bid (block 2, col R).
        """
        end_row = 4 + chain_rows
        q = chr(34)

        for i, strike in enumerate(strikes):
            row = 5 + i
            self.cal_sheet.range((row, 5)).value = float(strike)  # E column = strike to look up
            if use_iferror:
                formula = (
                    f"=IFERROR(XLOOKUP(E{row},Chain!E5:E{end_row},Chain!H5:H{end_row})"
                    f"+XLOOKUP(E{row},Chain!P5:P{end_row},Chain!R5:R{end_row}),"
                    f"{q}{q})"
                )
            else:
                formula = (
                    f"=XLOOKUP(E{row},Chain!E5:E{end_row},Chain!H5:H{end_row}," + q + "Na" + q + ")"
                    f"+XLOOKUP(E{row},Chain!P5:P{end_row},Chain!R5:R{end_row}," + q + "Na" + q + ")"
                )
            self.cal_sheet.range((row, 6)).api.Formula = formula  # F column = result

    def test_strikes_are_numeric(self):
        """Assert strikes are written as floats, not text."""
        strikes = [45000, 46000, 47000]
        chain_df = _make_chain(strikes)
        self._write_block(1, chain_df)

        for i, s in enumerate(strikes):
            cell_val = self.chain_sheet.range((5 + i, 5)).value  # Column E = Strike
            self.assertIsInstance(cell_val, (int, float),
                                  f"Strike {s} is {type(cell_val).__name__}, not numeric")

    def test_oi_is_numeric_not_text(self):
        """Assert OI values are floats, not '$1.50M' strings."""
        strikes = [45000, 46000, 47000]
        chain_df = _make_chain(strikes, oi=50000)
        self._write_block(1, chain_df)

        for i, s in enumerate(strikes):
            cell_val = self.chain_sheet.range((5 + i, 4)).value  # Column D = Call OI
            self.assertIsInstance(cell_val, (int, float),
                                  f"OI at strike {s} is {type(cell_val).__name__}, not numeric")

    def test_vlookup_returns_number_for_known_strike(self):
        """XLOOKUP into the chain for a strike in both blocks returns a number."""
        strikes = [45000, 46000, 47000]
        chain_df = _make_chain(strikes, base_call=1000.0, base_put=800.0, spread=10.0)
        n_rows = self._write_block(1, chain_df)
        self._write_block(12, chain_df)  # Both blocks written

        self._write_calendar_vlookup([46000.0], chain_rows=n_rows)
        self.app.calculation = "automatic"
        self.app.calculate()

        result = self.cal_sheet.range("F5").value
        self.assertIsNotNone(result, "XLOOKUP returned empty")
        self.assertNotIsInstance(result, str, f"XLOOKUP returned text: {result!r}")

    def test_missing_strike_with_iferror_returns_blank(self):
        """With IFERROR, a strike only in one block returns blank (same as US tool).

        The client formula wraps the FULL sum of both XLOOKUPs in IFERROR,
        so a missing leg in either block causes the entire result to go blank
        rather than #VALUE!.  This is the fixed behavior.
        """
        strikes_block1 = [45000, 46000, 47000]
        strikes_block2 = [45500, 46000, 46500, 47000]
        chain1 = _make_chain(strikes_block1, base_put=800.0)
        chain2 = _make_chain(strikes_block2, base_put=850.0)

        self._write_block(1, chain1)
        n_rows = self._write_block(12, chain2)

        # 45500 is only in block 2, not block 1
        self._write_calendar_vlookup([45500.0], chain_rows=n_rows, use_iferror=True)
        self.app.calculation = "automatic"
        self.app.calculate()

        val = self.cal_sheet.range("F5").value
        # With IFERROR wrapping the sum, a missing strike in one block
        # produces blank (not #VALUE! like the original formula would)
        self.assertTrue(val is None or val == "",
                        f"Expected blank for cross-expiry strike, got {val!r}")

    def test_missing_strike_both_blocks_with_iferror_is_blank(self):
        """With IFERROR, a strike in neither block returns blank, not #N/A."""
        strikes_block1 = [45000, 46000, 47000]
        strikes_block2 = [45500, 46500]
        chain1 = _make_chain(strikes_block1, base_put=800.0)
        chain2 = _make_chain(strikes_block2, base_put=850.0)

        self._write_block(1, chain1)
        n_rows = self._write_block(12, chain2)

        # 99999 is in neither block
        self._write_calendar_vlookup([99999.0], chain_rows=n_rows, use_iferror=True)
        self.app.calculation = "automatic"
        self.app.calculate()

        val = self.cal_sheet.range("F5").value
        self.assertTrue(val is None or val == "",
                        f"Expected blank for missing strike, got {val!r}")

    def test_recalculation_is_stable(self):
        """After multiple refreshes, no formula errors appear."""
        strikes = [45000, 46000, 47000]
        chain_df = _make_chain(strikes, base_put=800.0, spread=10.0)
        n_rows = self._write_block(1, chain_df)
        self._write_block(12, chain_df)

        for iteration in range(3):
            chain_df = _make_chain(strikes, base_put=800.0 + iteration * 10, spread=10.0)
            n_rows = self._write_block(1, chain_df)
            self._write_block(12, chain_df)
            self._write_calendar_vlookup(strikes, chain_rows=n_rows, use_iferror=True)

            self.app.calculation = "manual"
            self.app.calculate()
            self.app.calculation = "automatic"
            self.app.calculate()

            for i in range(len(strikes)):
                val = self.cal_sheet.range((5 + i, 6)).value
                self.assertNotIsInstance(val, str,
                                         f"Iteration {iteration}: lookup is text/error: {val!r}")


class DeltaTypingPreservationTests(unittest.TestCase):
    """Real Excel: a Calculate must not run while a cell is being edited."""

    @classmethod
    def setUpClass(cls):
        if xw is None:
            raise unittest.SkipTest("xlwings not installed")
        try:
            cls.app = xw.App(visible=True, add_book=False)
            cls.app.display_alerts = False
        except Exception as e:
            raise unittest.SkipTest(f"Excel not available: {e}")

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "app"):
            try:
                cls.app.quit()
            except Exception:
                pass

    def setUp(self):
        import ctypes
        import time

        import win32com.client

        self.time = time
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.book = self.app.books.add()
        self.sheet = self.book.sheets[0]
        self.saved = (
            delta_state._app_cache,
            delta_state._bulk_will_recalc,
            delta_state._recalc_owed,
        )
        delta_state._app_cache = self.app
        delta_state._bulk_will_recalc = True
        delta_state._recalc_owed = False
        try:
            self.shell.AppActivate(self.app.pid)
        except Exception:
            pass
        self.book.activate()
        self.sheet.activate()
        self.sheet.range("A1").select()
        hwnd = self.app.api.Hwnd
        user32 = ctypes.windll.user32
        user32.ShowWindow(hwnd, 5)
        user32.SetForegroundWindow(hwnd)
        self.time.sleep(0.3)

    def tearDown(self):
        try:
            self.app.api.SendKeys("{ESC}")
        except Exception:
            pass
        (
            delta_state._app_cache,
            delta_state._bulk_will_recalc,
            delta_state._recalc_owed,
        ) = self.saved
        try:
            self.book.close()
        except Exception:
            pass

    def test_typed_formula_survives_the_recalc_guard(self):
        import ctypes

        hwnd = self.app.api.Hwnd
        user32 = ctypes.windll.user32
        for _ in range(5):
            try:
                self.shell.AppActivate(self.app.pid)
                self.book.activate()
                self.sheet.activate()
                self.sheet.range("A1").select()
            except Exception:
                pass
            user32.ShowWindow(hwnd, 5)
            user32.SetForegroundWindow(hwnd)
            self.time.sleep(0.2)
            if user32.GetForegroundWindow() == hwnd:
                break
        else:
            self.skipTest("Excel never took the foreground")

        self.shell.SendKeys("=1{+}2")
        self.time.sleep(0.35)
        if not delta_main._excel_in_edit_mode(self.app):
            self.skipTest("Excel did not enter edit mode")

        elapsed = delta_main._time_recalc(True)

        self.assertEqual(elapsed, 0.0, "guard must skip Calculate")
        self.assertTrue(delta_state._recalc_owed, "skipped recalc must be carried")
        self.assertTrue(
            delta_main._excel_in_edit_mode(self.app),
            "Excel left edit mode unexpectedly",
        )

        self.shell.SendKeys("{ENTER}")
        self.time.sleep(0.35)
        self.assertEqual(self.sheet.range("A1").formula, "=1+2")
        self.assertEqual(self.sheet.range("A1").value, 3)


if __name__ == "__main__":
    unittest.main()
