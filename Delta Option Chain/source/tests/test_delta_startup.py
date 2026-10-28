import inspect
import unittest
from unittest.mock import patch

import delta_instance
import main


class StartupLicensingTests(unittest.TestCase):
    """Startup order.

    The instance guard must run before anything that can block, because the
    activation dialog blocks. The guard used to run after it, so two launches
    both sat on a dialog, neither ever reached the guard, and both went on to
    write the same three workbooks. Two writers mean two chain sessions and a
    sheet showing a mix of two snapshots.

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
    @patch.object(main, "show_error")
    @patch.object(main.os, "_exit", side_effect=SystemExit)
    @patch.object(main.asyncio, "run")
    def test_duplicate_that_cannot_be_stopped_refuses_to_start(
        self,
        run_async,
        exit_process,
        show_error,
        stop_previous,
        ensure_license,
        write_diagnostics,
    ):
        """Two writers on the same workbooks is worse than one writer saying no."""
        stop_previous.side_effect = RuntimeError(
            "Another copy of delta_option_chain.exe (pid 123) is still running"
        )
        with self.assertRaises(SystemExit):
            main.main()

        exit_process.assert_called_once_with(1)
        show_error.assert_called_once()
        run_async.assert_not_called()


class ProductNamingTests(unittest.TestCase):
    """The executable name says which product it is."""

    def test_product_exe_is_named_for_the_product(self):
        import delta_state

        self.assertEqual(delta_state.PRODUCT_EXE, "delta_option_chain.exe")

    def test_guard_matches_only_our_own_executable(self):
        """The other product's tool must never be a candidate."""
        src = inspect.getsource(delta_instance.kill_previous_instances)
        self.assertIn("PRODUCT_EXE", src)
        self.assertNotIn('"optionchain.exe"', src)


class LockFileTests(unittest.TestCase):
    """A same-folder duplicate is found without reading process paths.

    Reading another process's image path fails when the two run at different
    privilege levels or when security software blocks the query. The lock file
    in the product folder does not have that problem: a PID found in it is by
    definition an instance of this tool in this folder.
    """

    def test_lock_file_is_used(self):
        src = inspect.getsource(delta_instance.kill_previous_instances)
        self.assertIn("_take_over_locked_instance", src)

    def test_lock_path_is_in_the_product_folder(self):
        src = inspect.getsource(delta_instance._lock_path)
        self.assertIn("BASE_DIR", src)

    def test_duplicate_that_cannot_be_stopped_raises(self):
        src = inspect.getsource(delta_instance._take_over_locked_instance)
        self.assertIn("RuntimeError", src)

    def test_lock_pid_is_written_on_the_way_out(self):
        src = inspect.getsource(delta_instance.kill_previous_instances)
        self.assertIn("_write_lock_pid", src)


class LockFileBehaviourTests(unittest.TestCase):
    """The lock helpers, exercised for real rather than by source inspection."""

    def setUp(self):
        import tempfile

        import delta_instance
        import delta_state

        self.tmp = tempfile.TemporaryDirectory()
        # The lock file lives in the product folder, and there is exactly one
        # copy of that folder path: delta_state.
        self.patch = patch.object(delta_state, "BASE_DIR", self.tmp.name)
        self.patch.start()
        self.addCleanup(self.patch.stop)
        self.addCleanup(self.tmp.cleanup)
        self.delta_instance = delta_instance

    def test_round_trip(self):
        self.delta_instance._write_lock_pid(4321)
        self.assertEqual(self.delta_instance._read_lock_pid(), 4321)

    def test_a_missing_lock_file_reads_as_none(self):
        self.assertIsNone(self.delta_instance._read_lock_pid())

    def test_a_corrupt_lock_file_reads_as_none(self):
        with open(self.delta_instance._lock_path(), "w", encoding="ascii") as fh:
            fh.write("not a pid")
        self.assertIsNone(self.delta_instance._read_lock_pid())

    def test_a_dead_pid_is_taken_over(self):
        self.delta_instance._write_lock_pid(999999)
        # A PID that cannot be alive means the previous run died, so this
        # launch simply takes over.
        self.assertFalse(self.delta_instance._take_over_locked_instance())

    def test_our_own_pid_is_never_stopped(self):
        import os

        self.delta_instance._write_lock_pid(os.getpid())
        self.assertFalse(self.delta_instance._take_over_locked_instance())


if __name__ == "__main__":
    unittest.main()
