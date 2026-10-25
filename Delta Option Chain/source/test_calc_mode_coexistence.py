import json
import os
import tempfile
import time
import unittest
from unittest import mock

import calc_mode_guard
import delta_state
import main as delta_main


class TempStateTestBase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.mkdtemp()
        self._old_appdata = os.environ.get("LOCALAPPDATA")
        os.environ["LOCALAPPDATA"] = self._tmp
        self.record_path = os.path.join(self._tmp, "OptionChain", calc_mode_guard.RECORD_NAME)

    def tearDown(self):
        if self._old_appdata is None:
            os.environ.pop("LOCALAPPDATA", None)
        else:
            os.environ["LOCALAPPDATA"] = self._old_appdata

    def _write_record(self, holders):
        os.makedirs(os.path.dirname(self.record_path), exist_ok=True)
        data = {
            "original_mode": "automatic",
            "created_at": time.time(),
            "seen_at": time.time(),
            "holders": holders,
        }
        with open(self.record_path, "w", encoding="utf-8") as handle:
            json.dump(data, handle)


class OtherHolderAliveTests(TempStateTestBase):
    def test_true_when_a_foreign_live_process_holds_the_record(self):
        self._write_record({"delta": os.getpid()})
        self.assertTrue(calc_mode_guard.other_holder_alive("us"))

    def test_false_when_only_our_own_claim_is_listed(self):
        self._write_record({"us": os.getpid()})
        self.assertFalse(calc_mode_guard.other_holder_alive("us"))

    def test_false_when_there_is_no_record(self):
        self.assertFalse(calc_mode_guard.other_holder_alive("us"))

    def test_false_when_the_foreign_pid_is_not_running(self):
        self._write_record({"delta": "not-a-pid"})
        self.assertFalse(calc_mode_guard.other_holder_alive("us"))


class FakeCalcApi:
    def __init__(self, mode):
        self.Calculation = mode


class FakeApp:
    def __init__(self, mode=-4135):
        self.api = FakeCalcApi(mode)


class WatchdogCoexistenceTests(TempStateTestBase):
    # Phase-1 gate: the watchdog skips any restore while a human is typing.
    # These tests drive the rescue paths, so pin the input timer to "idle"
    # instead of letting physical typing on the dev machine trip the skip:
    # the real _last_user_input_age reads system-wide GetLastInputInfo, which
    # is whatever the developer typing happens to do during the run.
    def _idle_input(self):
        both = (
            mock.patch.object(delta_main, "_input_is_recent", return_value=False),
            mock.patch.object(delta_main, "_last_user_input_age", return_value=999.0),
        )
        for p in both:
            p.start()
            self.addCleanup(p.stop)

    def test_watchdog_leaves_manual_alone_while_the_sibling_runs(self):
        self._idle_input()
        self._write_record({"us": os.getpid()})
        app = FakeApp(mode=-4135)
        delta_state._manual_since = time.time() - delta_main.MANUAL_WATCHDOG_SECONDS - 1
        with mock.patch.object(delta_main, "feed_log"):
            delta_main.ensure_calculation_not_stranded(app=app)
        self.assertEqual(int(app.api.Calculation), -4135)
        delta_state._manual_since = 0.0

    def test_watchdog_still_rescues_when_no_tool_holds_the_record(self):
        self._idle_input()
        app = FakeApp(mode=-4135)
        delta_state._manual_since = time.time() - delta_main.MANUAL_WATCHDOG_SECONDS - 1
        with mock.patch.object(delta_main, "feed_log"):
            delta_main.ensure_calculation_not_stranded(app=app)
        self.assertEqual(int(app.api.Calculation), -4105)
        delta_state._manual_since = 0.0
    def test_watchdog_leaves_manual_alone_on_non_chain_sheet(self):
        self._idle_input()
        app = FakeApp(mode=-4135)
        delta_state._manual_since = time.time() - delta_main.MANUAL_WATCHDOG_SECONDS - 1
        delta_main._watchdog_input_skip_logged_at = 0.0


        with mock.patch.object(delta_main, "_active_sheet_is_chain", return_value=False), \
             mock.patch.object(delta_main, "feed_log") as log:
            delta_main.ensure_calculation_not_stranded(app=app)
        self.assertEqual(int(app.api.Calculation), -4135)
        events = [c for c in log.call_args_list
                  if c.args and c.args[0] == "watchdog_skipped_strategy_sheet"]
        self.assertEqual(len(events), 1)
        delta_state._manual_since = 0.0




class ChurnObservationTests(unittest.TestCase):
    def setUp(self):
        self._old_last = delta_state._last_observed_calc_raw
        self._old_count = delta_state._calc_churn_count
        delta_state._last_observed_calc_raw = None
        delta_state._calc_churn_count = 0

    def tearDown(self):
        delta_state._last_observed_calc_raw = self._old_last
        delta_state._calc_churn_count = self._old_count

    def test_transitions_are_counted_and_summary_logged_once_per_window(self):
        with mock.patch.object(delta_main, "feed_log") as log:
            delta_main._observe_raw_calc(-4105)
            delta_main._observe_raw_calc(-4135)
            delta_main._observe_raw_calc(-4105)
            self.assertGreaterEqual(delta_state._calc_churn_count, 2)
            logged = [c for c in log.call_args_list if c.args[0] == "calc_switch_summary"]
            self.assertEqual(len(logged), 1)

    def test_glitch_does_not_count_as_a_transition(self):
        with mock.patch.object(delta_main, "feed_log"):
            delta_main._observe_raw_calc(-4105)
            delta_main._observe_raw_calc(None)
            delta_main._observe_raw_calc(-4105)
            self.assertEqual(delta_state._calc_churn_count, 0)



class WatchdogReconnectTests(unittest.TestCase):
    """ensure_calculation_not_stranded runs against the tool's own Excel only."""

    def test_reconnect_in_watchdog_source(self):
        import inspect
        import main as delta_main
        src = inspect.getsource(delta_main.ensure_calculation_not_stranded)
        self.assertIn('delta_state._app_cache', src)
        self.assertIn('delta_state._owner_app', src)
        self.assertNotIn('xw.apps', src)

    def test_watchdog_reports_stranded_excel_without_modal(self):
        import types
        import main as delta_main
        old_cache = delta_state._app_cache
        old_owner = delta_state._owner_app
        old_manual = delta_state._manual_since
        try:
            app = types.SimpleNamespace()
            delta_state._app_cache = app
            delta_state._owner_app = app
            delta_state._manual_since = 1e9
            with mock.patch.object(delta_main, "_input_is_recent", return_value=False), \
                 mock.patch.object(delta_main, "_last_user_input_age", return_value=999.0), \
                 mock.patch.object(delta_main, "_raw_calc_mode", return_value=-4135), \
                 mock.patch.object(delta_main, "_other_chain_tool_holds", return_value=False), \
                 mock.patch.object(delta_main, "_set_calc_mode", return_value=True) as mode:
                delta_main.ensure_calculation_not_stranded(app)
            mode.assert_called_once_with(app, "automatic")
            self.assertEqual(delta_state._manual_since, 0.0)
        finally:
            delta_state._app_cache = old_cache
            delta_state._owner_app = old_owner
            delta_state._manual_since = old_manual


if __name__ == "__main__":
    unittest.main()
