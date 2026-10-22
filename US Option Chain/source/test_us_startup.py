import unittest
from unittest.mock import patch

import main


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
        src = inspect.getsource(main.kill_previous_instances)
        self.assertIn("PRODUCT_EXE", src)
        self.assertNotIn('"optionchain.exe"', src)


class LockFileTests(unittest.TestCase):
    """A same-folder duplicate is found without reading process paths."""

    def test_lock_file_is_used(self):
        import inspect
        src = inspect.getsource(main.kill_previous_instances)
        self.assertIn("_take_over_locked_instance", src)

    def test_lock_path_is_in_the_product_folder(self):
        import inspect
        src = inspect.getsource(main._lock_path)
        self.assertIn("BASE_DIR", src)

    def test_duplicate_that_cannot_be_stopped_raises(self):
        import inspect
        src = inspect.getsource(main._take_over_locked_instance)
        self.assertIn("RuntimeError", src)


if __name__ == "__main__":
    unittest.main()
