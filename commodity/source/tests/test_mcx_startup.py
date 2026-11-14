import inspect
import unittest
from unittest.mock import patch

import mcx_instance
import main


class StartupOrderTests(unittest.TestCase):
    """Startup order.

    The instance guard must run before anything that can block. Two writers on
    the same workbooks is worse than one writer refusing to start.
    """

    @patch.object(main, "write_startup_diagnostics")
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
        write_diagnostics,
    ):
        """Two writers on the same workbooks is worse than one writer saying no."""
        stop_previous.side_effect = RuntimeError(
            "Another copy of mcx_option_chain.exe (pid 123) is still running"
        )
        with self.assertRaises(SystemExit):
            main.main()

        exit_process.assert_called_once_with(1)
        show_error.assert_called_once()
        run_async.assert_not_called()

    @patch.object(main, "write_startup_diagnostics")
    @patch.object(main, "kill_previous_instances")
    @patch.object(main, "show_error")
    @patch.object(main.os, "_exit", side_effect=SystemExit)
    @patch.object(main.asyncio, "run")
    def test_instance_guard_runs_before_event_loop(
        self,
        run_async,
        exit_process,
        show_error,
        stop_previous,
        write_diagnostics,
    ):
        """The instance guard runs before the async event loop starts."""
        main.main()

        stop_previous.assert_called_once_with()
        run_async.assert_called_once()


class ProductNamingTests(unittest.TestCase):
    """The executable name says which product it is."""

    def test_product_exe_is_named_for_the_product(self):
        import mcx_state

        self.assertEqual(mcx_state.PRODUCT_EXE, "mcx_option_chain.exe")

    def test_guard_matches_only_our_own_executable(self):
        """The other product's tool must never be a candidate."""
        src = inspect.getsource(mcx_instance.kill_previous_instances)
        self.assertIn("PRODUCT_EXE", src)
        self.assertNotIn('"optionchain.exe"', src)


class LockFileTests(unittest.TestCase):
    """A same-folder duplicate is found without reading process paths."""

    def test_lock_file_is_used(self):
        src = inspect.getsource(mcx_instance.kill_previous_instances)
        self.assertIn("_take_over_locked_instance", src)

    def test_lock_path_is_in_the_product_folder(self):
        src = inspect.getsource(mcx_instance._lock_path)
        self.assertIn("BASE_DIR", src)

    def test_duplicate_that_cannot_be_stopped_raises(self):
        src = inspect.getsource(mcx_instance._take_over_locked_instance)
        self.assertIn("RuntimeError", src)

    def test_lock_pid_is_written_on_the_way_out(self):
        src = inspect.getsource(mcx_instance.kill_previous_instances)
        self.assertIn("_write_lock_pid", src)


class LockFileBehaviourTests(unittest.TestCase):
    """The lock helpers, exercised for real rather than by source inspection."""

    def setUp(self):
        import tempfile

        import mcx_instance
        import mcx_state

        self.tmp = tempfile.TemporaryDirectory()
        self.patch = patch.object(mcx_state, "BASE_DIR", self.tmp.name)
        self.patch.start()
        self.addCleanup(self.patch.stop)
        self.addCleanup(self.tmp.cleanup)
        self.mcx_instance = mcx_instance

    def test_round_trip(self):
        self.mcx_instance._write_lock_pid(4321)
        self.assertEqual(self.mcx_instance._read_lock_pid(), 4321)

    def test_a_missing_lock_file_reads_as_none(self):
        self.assertIsNone(self.mcx_instance._read_lock_pid())

    def test_a_corrupt_lock_file_reads_as_none(self):
        with open(self.mcx_instance._lock_path(), "w", encoding="ascii") as fh:
            fh.write("not a pid")
        self.assertIsNone(self.mcx_instance._read_lock_pid())

    def test_a_dead_pid_is_taken_over(self):
        self.mcx_instance._write_lock_pid(999999)
        self.assertFalse(self.mcx_instance._take_over_locked_instance())

    def test_our_own_pid_is_never_stopped(self):
        import os

        self.mcx_instance._write_lock_pid(os.getpid())
        self.assertFalse(self.mcx_instance._take_over_locked_instance())


if __name__ == "__main__":
    unittest.main()
