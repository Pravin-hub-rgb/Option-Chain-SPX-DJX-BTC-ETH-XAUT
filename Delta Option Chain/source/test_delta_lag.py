import inspect
import unittest

import main as delta_main


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
        src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn("merged_rows", src)
        self.assertNotIn("for start, end in runs:", src)

    def test_unchanged_rows_carry_previous_values(self):
        src = inspect.getsource(delta_main._write_to_excel)
        self.assertIn("merged_rows.append(list(prev[i]))", src)

    def test_ownership_flag_gates_row_granularity(self):
        src = inspect.getsource(delta_main._begin_excel_bulk_mode)
        self.assertIn("_bulk_owning_manual = True", src)
        self.assertIn("_bulk_owning_manual = False", src)
        self.assertIn("protected = _bulk_owning_manual",
                      inspect.getsource(delta_main._write_to_excel))


if __name__ == "__main__":
    unittest.main()
