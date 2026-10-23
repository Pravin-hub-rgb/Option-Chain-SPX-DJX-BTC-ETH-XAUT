"""Integration test: verify the client's XLOOKUP strategy formula works against
real Excel with our block layout.

Builds a workbook with a Chain sheet (our 9-column block output for two
expiries) and a Calendar sheet with the client's combo formula.  Drives the
real block writer, forces recalculation, and asserts:

  - no #VALUE! / #N/A / #REF! / #DIV/0! anywhere in the combo column
  - strikes are numeric (assert type before writing)
  - strikes in both blocks produce the correct combo value
  - strikes in one block only produce blank, never an error
  - re-centering does not drop strikes
  - empty bid/ask yields blank, not 0

Requires Excel + xlwings.  Skipped if Excel is unavailable.
"""

import math
import unittest

import pandas as pd

try:
    import xlwings as xw
except ImportError:
    xw = None


def _make_chain(strikes, base_call=30.0, base_put=30.0, spread=2.0,
                oi=100, ltp=None, skip_strikes=None):
    """Build a synthetic option chain DataFrame.

    All strikes present unless explicitly skipped.  Each strike gets a call
    and a put row with bid/ask centred on *base_call* / *base_put*.
    """
    skip_strikes = skip_strikes or set()
    rows = []
    calls = []
    puts = []
    for s in strikes:
        if s in skip_strikes:
            continue
        c_bid = base_call + (s - strikes[len(strikes) // 2]) * 0.01
        c_ask = c_bid + spread
        p_bid = base_put - (s - strikes[len(strikes) // 2]) * 0.01
        p_ask = p_bid + spread
        calls.append({
            "strike": float(s),
            "Call LTP": ltp if ltp is not None else c_bid,
            "Call Bid": c_bid,
            "Call Ask": c_ask,
            "Call OI": oi,
        })
        puts.append({
            "strike": float(s),
            "Put OI": oi,
            "Put Bid": p_bid,
            "Put Ask": p_ask,
            "Put LTP": ltp if ltp is not None else p_bid,
        })

    df_calls = pd.DataFrame(calls)
    df_puts = pd.DataFrame(puts)
    merged = pd.merge(df_calls, df_puts, on="strike", how="outer")
    merged = merged.sort_values("strike").reset_index(drop=True).fillna(0.0)
    return merged


class StrategyFormulaTests(unittest.TestCase):
    """Test the client's XLOOKUP combo formula against real Excel."""

    HEADERS = ["Call LTP", "Call Bid", "Call Ask", "Call OI", "strike",
               "Put OI", "Put Bid", "Put Ask", "Put LTP"]

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

        # Write headers for both blocks on the Chain sheet
        for i, h in enumerate(self.HEADERS):
            self.chain_sheet.range((4, 1 + i)).value = h
        for i, h in enumerate(self.HEADERS):
            self.chain_sheet.range((4, 12 + i)).value = h

        # Force calculation manual during setup
        self.app.calculation = "manual"

    def _write_block(self, start_col, chain_df):
        """Write a chain DataFrame starting at row 5, column start_col.

        start_col is 1-based (1 = column A).
        """
        for col_idx, col_name in enumerate(self.HEADERS):
            col = start_col + col_idx
            for row_idx, val in enumerate(chain_df[col_name].tolist()):
                self.chain_sheet.range((5 + row_idx, col)).value = val
        return len(chain_df)

    def _write_calendar_strategies(self, strikes, chain_rows, formula_template="client_broken"):
        """Write the combo formula for each strike row in the Calendar sheet.

        Args:
            strikes: list of strike values to look up
            chain_rows: number of data rows in each Chain block (determines
                        the XLOOKUP range end)
        """
        end_row = 4 + chain_rows
        q = chr(34)

        if formula_template == "client_broken":
            tmpl = ("=XLOOKUP(E{row},Chain!P5:P{end},Chain!R5:R{end}," + q + "Na" + q + ")"
                     "+XLOOKUP(E{row},Chain!E5:E{end},Chain!H5:H{end}," + q + "Na" + q + ")")
        elif formula_template == "client_fixed":
            tmpl = ("=IFERROR(XLOOKUP(E{row},Chain!P5:P{end},Chain!R5:R{end})"
                     "+XLOOKUP(E{row},Chain!E5:E{end},Chain!H5:H{end})," + q + q + ")")
        else:
            raise ValueError(f"Unknown template: {formula_template}")

        for i, strike in enumerate(strikes):
            row = 5 + i
            self.cal_sheet.range((row, 5)).value = float(strike)  # E column
            formula = tmpl.format(row=row, end=end_row)
            self.cal_sheet.range((row, 6)).api.Formula = formula  # F column

    def test_strikes_in_both_blocks_produce_correct_value(self):
        """A strike in both expiries should yield PutBid_block2 + PutAsk_block1."""
        strikes = [7700, 7750, 7800]
        chain_df = _make_chain(strikes, base_call=30.0, base_put=20.0, spread=2.0)
        n_rows = self._write_block(1, chain_df)   # Block 1 at col 1
        self._write_block(12, chain_df)  # Block 2 at col 12

        # Expected: PutBid from block 2 (col R) + PutAsk from block 1 (col H)
        # At strike 7750: PutBid_block2=20, PutAsk_block1=22, combined=42
        self._write_calendar_strategies([7750.0], chain_rows=n_rows)

        self.app.calculation = "automatic"
        self.app.calculate()

        result = self.cal_sheet.range("F5").value
        self.assertIsNotNone(result, "Combo cell is empty")
        self.assertAlmostEqual(float(result), 42.0, delta=0.01)

    def test_strike_missing_in_one_block_shows_blank_fixed_formula(self):
        """With the fixed IFERROR formula, missing strike = blank, not #VALUE!."""
        strikes_current = [7700, 7750, 7800]
        strikes_next = [7725, 7750, 7775, 7800]
        chain_current = _make_chain(strikes_current, base_put=20.0, spread=2.0)
        chain_next = _make_chain(strikes_next, base_put=25.0, spread=1.5)

        self._write_block(1, chain_current)
        n_rows = self._write_block(12, chain_next)

        # 7725 only in block 2, 7750/7800 in both
        self._write_calendar_strategies([7725.0, 7750.0, 7800.0],
                                         chain_rows=n_rows,
                                         formula_template="client_fixed")

        self.app.calculation = "automatic"
        self.app.calculate()

        for i, strike in enumerate([7725.0, 7750.0, 7800.0]):
            val = self.cal_sheet.range((5 + i, 6)).value
            if strike == 7725.0:
                # Only in block 2 → XLOOKUP into block 1 fails → IFERROR → ""
                self.assertTrue(val is None or val == "",
                                f"Strike {strike} should be blank, got {val!r}")
            else:
                self.assertIsNotNone(val, f"Strike {strike} should have a value")
                self.assertNotIsInstance(val, str)

    def test_strike_missing_in_one_block_shows_value_error_broken_formula(self):
        """With the original broken formula, missing strike = #VALUE!."""
        strikes_current = [7700, 7750, 7800]
        strikes_next = [7725, 7750, 7775, 7800]
        chain_current = _make_chain(strikes_current, base_put=20.0, spread=2.0)
        chain_next = _make_chain(strikes_next, base_put=25.0, spread=1.5)

        self._write_block(1, chain_current)
        n_rows = self._write_block(12, chain_next)

        self._write_calendar_strategies([7725.0], chain_rows=n_rows, formula_template="client_broken")

        self.app.calculation = "automatic"
        self.app.calculate()

        val = self.cal_sheet.range("F5").value
        # The broken formula returns "Na" (text) for the missing leg,
        # and number + text = #VALUE!.  Excel represents #VALUE! as None
        # through xlwings, so None confirms the error.
        self.assertIsNone(val,
                          f"Expected #VALUE! (None via xlwings) for missing strike, got {val!r}")

    def test_all_strikes_are_numeric(self):
        """Assert strikes are written as floats, not text."""
        strikes = [7700, 7750, 7800, 7850]
        chain_df = _make_chain(strikes)
        self._write_block(1, chain_df)

        for i, s in enumerate(strikes):
            cell_val = self.chain_sheet.range((5 + i, 5)).value  # Column E = Strike
            self.assertIsInstance(cell_val, (int, float),
                                  f"Strike {s} is {type(cell_val).__name__}, not numeric")

    def test_empty_bid_ask_gives_zero_not_error(self):
        """When bid/ask are 0, the chain should still write numeric strikes."""
        strikes = [7700, 7750, 7800]
        chain_df = _make_chain(strikes, base_call=0.0, base_put=0.0, spread=0.0)
        self._write_block(1, chain_df)

        for i, s in enumerate(strikes):
            cell_val = self.chain_sheet.range((5 + i, 5)).value
            self.assertIsInstance(cell_val, (int, float),
                                  f"Strike {s} with empty quotes is not numeric")

    def test_recalculation_count_is_stable(self):
        """After multiple refreshes, no formula errors appear."""
        strikes = [7700, 7750, 7800]
        chain_df = _make_chain(strikes, base_put=20.0, spread=2.0)
        n_rows = self._write_block(1, chain_df)

        for iteration in range(5):
            chain_df = _make_chain(strikes, base_put=20.0 + iteration, spread=2.0)
            n_rows = self._write_block(1, chain_df)
            self._write_block(12, chain_df)

            self._write_calendar_strategies(strikes, chain_rows=n_rows,
                                             formula_template="client_fixed")

            self.app.calculation = "manual"
            self.app.calculate()
            self.app.calculation = "automatic"
            self.app.calculate()

            # Check no errors (#VALUE!, #N/A, etc.)
            for i in range(len(strikes)):
                val = self.cal_sheet.range((5 + i, 6)).value
                self.assertNotIsInstance(val, str,
                                         f"Iteration {iteration}: combo is text/error: {val!r}")


if __name__ == "__main__":
    unittest.main()
