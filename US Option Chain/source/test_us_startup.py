import unittest
from unittest.mock import patch

import main
import us_instance


class StartupLicensingTests(unittest.TestCase):
    """Startup order.

    The instance guard must run before anything that can block, because both
    the activation dialog and the API-key dialog block. The guard used to run
    after them, so two launches both sat on a dialog, neither ever reached the
    guard, and both went on to write the same three workbooks. Two writers mean
    two Databento sessions and a sheet showing a mix of two snapshots - the
    spot appears to crawl and the data stops updating.

    No test in here may put a real message box on the developer's screen, so
    show_error is patched in every test that could reach it.
    """

    @patch.object(main, "write_startup_diagnostics")
    @patch.object(main, "ensure_license", return_value=False)
    @patch.object(main, "kill_previous_instances")
    @patch.object(main, "show_error")
    @patch.object(main.os, "_exit", side_effect=SystemExit)
    @patch.object(main.asyncio, "run")
    def test_missing_license_stops_older_instance_first(
        self,
        run_async,
        exit_process,
        show_error,
        stop_previous,
        ensure_license,
        write_diagnostics,
    ):
        with self.assertRaises(SystemExit):
            main.main()

        stop_previous.assert_called_once_with()
        ensure_license.assert_called_once_with()
        exit_process.assert_called_once_with(1)
        run_async.assert_not_called()

    @patch.object(main, "write_startup_diagnostics")
    @patch.object(main, "ensure_license", return_value=True)
    @patch.object(main, "kill_previous_instances")
    @patch.object(main, "_start_databento")
    @patch.object(main, "show_error")
    @patch.object(main.os, "_exit", side_effect=SystemExit)
    @patch.object(main.asyncio, "run", side_effect=SystemExit)
    def test_guard_runs_before_the_databento_key_dialog(
        self,
        run_async,
        exit_process,
        show_error,
        start_databento,
        stop_previous,
        ensure_license,
        write_diagnostics,
    ):
        """The API-key dialog blocks too, so the guard must come first too."""
        with self.assertRaises(SystemExit):
            main.main()

        stop_previous.assert_called_once_with()
        start_databento.assert_called_once_with()

    @patch.object(main, "write_startup_diagnostics")
    @patch.object(main, "ensure_license", return_value=True)
    @patch.object(main, "kill_previous_instances")
    @patch.object(main, "_start_databento")
    @patch.object(main, "show_error")
    @patch.object(main.os, "_exit", side_effect=SystemExit)
    @patch.object(main.asyncio, "run")
    def test_duplicate_that_cannot_be_stopped_refuses_to_start(
        self,
        run_async,
        exit_process,
        show_error,
        start_databento,
        stop_previous,
        ensure_license,
        write_diagnostics,
    ):
        """Two writers on the same workbooks is worse than one writer saying no."""
        stop_previous.side_effect = RuntimeError(
            "Another copy of OptionChain (pid 123) is still running"
        )
        with self.assertRaises(SystemExit):
            main.main()

        exit_process.assert_called_once_with(1)
        show_error.assert_called_once()
        start_databento.assert_not_called()
        run_async.assert_not_called()


class ProductNamingTests(unittest.TestCase):
    """The executable name says which product it is."""

    def test_product_exe_is_named_for_the_product(self):
        self.assertEqual(main.PRODUCT_EXE, "us_option_chain.exe")

    def test_guard_matches_only_our_own_executable(self):
        """The other product's tool must never be a candidate."""
        import inspect
        src = inspect.getsource(us_instance.kill_previous_instances)
        self.assertIn("PRODUCT_EXE", src)
        self.assertNotIn('"optionchain.exe"', src)


class LockFileTests(unittest.TestCase):
    """A same-folder duplicate is found without reading process paths."""

    def test_lock_file_is_used(self):
        import inspect
        src = inspect.getsource(us_instance.kill_previous_instances)
        self.assertIn("_take_over_locked_instance", src)

    def test_lock_path_is_in_the_product_folder(self):
        import inspect
        src = inspect.getsource(us_instance._lock_path)
        self.assertIn("BASE_DIR", src)

    def test_duplicate_that_cannot_be_stopped_raises(self):
        import inspect
        src = inspect.getsource(us_instance._take_over_locked_instance)
        self.assertIn("RuntimeError", src)


class DatabentoKeyGateTests(unittest.TestCase):
    """A key saved on disk can die at any time; the app must not open on it."""

    def _gate(self, ensure, validate, calls):
        return main._ensure_working_databento_key(
            _ensure=ensure,
            _validate=validate,
            _drop=lambda: calls.append("drop"),
            _error=lambda msg: calls.append(("error", msg)),
            _exit=lambda code: calls.append(("exit", code)),
        )

    @staticmethod
    def _errors(calls):
        return [msg for call in calls if isinstance(call, tuple)
                and call[0] == "error" for msg in call[1:]]

    def test_dead_saved_key_is_wiped_and_asked_again(self):
        keys = iter(["db-expired", "db-new"])
        attempts = []

        def ensure():
            key = next(keys)
            attempts.append(key)
            return key

        def validate(key):
            return (False, "Databento rejected this key.", True) if key.endswith(
                "expired"
            ) else (True, "", False)

        calls = []
        self.assertEqual("db-new", self._gate(ensure, validate, calls))
        self.assertEqual(["drop"], [c for c in calls if c == "drop"])
        self.assertEqual(1, len(self._errors(calls)))

    def test_dead_key_twice_exits_instead_of_looping_dialogs(self):
        calls = []

        def ensure():
            return "db-expired-from-env"

        with self.assertRaises(SystemExit):
            self._gate(ensure, lambda key: (False, "rejected", True), calls)
        self.assertTrue(
            any(("exit", 1) == c for c in calls),
            calls,
        )

    def test_cancelled_dialog_closes_the_app(self):
        calls = []

        def ensure():
            return None

        with self.assertRaises(SystemExit):
            self._gate(ensure, lambda key: (True, "", False), calls)
        self.assertTrue(any(("exit", 1) == c for c in calls))

    def test_first_error_message_reaches_the_dialog(self):
        calls = []

        def ensure():
            return "db-k"

        with self.assertRaises(SystemExit):
            self._gate(ensure, lambda key: (False, "Databento rejected this key.", True), calls)
        self.assertTrue(
            any("Databento rejected" in m for m in self._errors(calls))
        )

    def test_an_unreachable_databento_keeps_the_key(self):
        """A key that could not be checked is not a key that was refused.

        The connection dropping used to delete the customer's working key and
        then demand a new one, so a flaky internet connection cost them a
        trip to the Databento website.
        """
        calls = []

        def ensure():
            return "db-good"

        with self.assertRaises(SystemExit):
            self._gate(
                ensure,
                lambda key: (False, "Cannot reach Databento (ConnectTimeout).", False),
                calls,
            )
        # The key must NOT have been dropped.
        self.assertEqual([], [c for c in calls if c == "drop"])
        self.assertTrue(any(("exit", 1) == c for c in calls), calls)
        self.assertTrue(
            any("kept" in m for m in self._errors(calls)), self._errors(calls)
        )


if __name__ == "__main__":
    unittest.main()
