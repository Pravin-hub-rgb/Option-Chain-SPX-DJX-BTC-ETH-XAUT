import json
import os
import tempfile
import time
import unittest
from unittest import mock

import calc_mode_guard
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
    def test_watchdog_leaves_manual_alone_while_the_sibling_runs(self):
        self._write_record({"us": os.getpid()})
        app = FakeApp(mode=-4135)
        delta_main._manual_since = time.time() - delta_main.MANUAL_WATCHDOG_SECONDS - 1
        with mock.patch.object(delta_main, "feed_log"):
            delta_main.ensure_calculation_not_stranded(app=app)
        self.assertEqual(int(app.api.Calculation), -4135)
        delta_main._manual_since = 0.0

    def test_watchdog_still_rescues_when_no_tool_holds_the_record(self):
        app = FakeApp(mode=-4135)
        delta_main._manual_since = time.time() - delta_main.MANUAL_WATCHDOG_SECONDS - 1
        with mock.patch.object(delta_main, "feed_log"):
            delta_main.ensure_calculation_not_stranded(app=app)
        self.assertEqual(int(app.api.Calculation), -4105)
        delta_main._manual_since = 0.0


class ChurnObservationTests(unittest.TestCase):
    def setUp(self):
        self._old_last = delta_main._last_observed_calc_raw
        self._old_count = delta_main._calc_churn_count
        delta_main._last_observed_calc_raw = None
        delta_main._calc_churn_count = 0

    def tearDown(self):
        delta_main._last_observed_calc_raw = self._old_last
        delta_main._calc_churn_count = self._old_count

    def test_transitions_are_counted_and_summary_logged_once_per_window(self):
        with mock.patch.object(delta_main, "feed_log") as log:
            delta_main._observe_raw_calc(-4105)
            delta_main._observe_raw_calc(-4135)
            delta_main._observe_raw_calc(-4105)
            self.assertGreaterEqual(delta_main._calc_churn_count, 2)
            logged = [c for c in log.call_args_list if c.args[0] == "calc_switch_summary"]
            self.assertEqual(len(logged), 1)

    def test_glitch_does_not_count_as_a_transition(self):
        with mock.patch.object(delta_main, "feed_log"):
            delta_main._observe_raw_calc(-4105)
            delta_main._observe_raw_calc(None)
            delta_main._observe_raw_calc(-4105)
            self.assertEqual(delta_main._calc_churn_count, 0)


if __name__ == "__main__":
    unittest.main()
