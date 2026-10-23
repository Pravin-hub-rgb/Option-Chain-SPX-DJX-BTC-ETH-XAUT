"""What happens when the Databento key is rejected.

The key worked at 08:17 on 2026-10-22 and was rejected with 401 at 11:07 the
same morning, from the same file on the same machine. The code treated that
like a hiccup: three retries per chunk, for every parent, for every symbol,
and then the whole thing again with a five-to-one-hundred-and-twenty-second
backoff. Seventeen minutes of the log was nothing but 401s, and the customer
saw a blank sheet whose header said LOADING forever.

A rejected key does not fix itself, so the feed latches off, says so once, and
stops asking.
"""

import unittest
from unittest.mock import patch


import databento_feed
import us_chain
import us_excel
import us_state


class _FakeHttpError(Exception):
    """Shaped like the SDK's BentoHttpError, which carries the status."""

    def __init__(self, status):
        super().__init__(f"{status} something went wrong")
        self.http_status = status


class _FakeResponse:
    def __init__(self, status_code):
        self.status_code = status_code


class _FakeRequestsError(Exception):
    """Shaped like requests.HTTPError, which carries the response."""

    def __init__(self, status_code):
        super().__init__(f"{status_code} Client Error")
        self.response = _FakeResponse(status_code)


class AuthDetectionTests(unittest.TestCase):
    def test_a_401_is_an_auth_failure(self):
        self.assertTrue(databento_feed._is_auth_failure(_FakeHttpError(401)))

    def test_a_403_is_an_auth_failure(self):
        self.assertTrue(databento_feed._is_auth_failure(_FakeHttpError(403)))

    def test_a_raw_401_is_an_auth_failure(self):
        self.assertTrue(databento_feed._is_auth_failure(_FakeRequestsError(401)))

    def test_a_server_error_is_not_an_auth_failure(self):
        self.assertFalse(databento_feed._is_auth_failure(_FakeHttpError(500)))

    def test_a_timeout_is_not_an_auth_failure(self):
        self.assertFalse(databento_feed._is_auth_failure(TimeoutError("timed out")))


class AuthLatchTests(unittest.TestCase):
    def setUp(self):
        self.feed = databento_feed.DatabentoFeed(key="db-test-key")
        # Keep the expected failure out of the test output.
        patcher = patch.object(self.feed, "_log")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_a_rejected_key_is_remembered(self):
        self.feed._set_auth_error("SPX", _FakeHttpError(401))
        self.assertTrue(self.feed.auth_error())

    def test_the_reason_is_logged_once(self):
        with patch.object(self.feed, "_log") as log:
            self.feed._set_auth_error("SPX", _FakeHttpError(401))
            self.feed._set_auth_error("SPX", _FakeHttpError(401))
        events = [call.args[0] for call in log.call_args_list]
        self.assertEqual(events.count("databento_auth_failed"), 1)

    def test_status_text_says_the_key_was_rejected(self):
        self.feed._set_auth_error("SPX", _FakeHttpError(401))
        self.assertEqual(self.feed.status_text(), "DATABENTO KEY REJECTED")

    def test_a_latched_feed_never_asks_again(self):
        self.feed._set_auth_error("SPX", _FakeHttpError(401))
        with patch.object(
            self.feed, "_fetch_universe", side_effect=AssertionError("must not fetch")
        ) as fetch:
            self.assertFalse(self.feed.ensure_universe("SPX"))
            fetch.assert_not_called()

    def test_a_latched_feed_still_serves_what_it_already_has(self):
        self.feed._defs["SPX"] = [{"r": "SPX"}]
        self.feed._set_auth_error("SPX", _FakeHttpError(401))
        self.assertTrue(self.feed.ensure_universe("SPX"))

    def test_the_key_is_never_written_into_the_log(self):
        with patch.object(self.feed, "_log") as log:
            self.feed._set_auth_error("SPX", _FakeHttpError(401))
        for call in log.call_args_list:
            for value in call.kwargs.values():
                self.assertNotIn("db-test-key", str(value))


class UnprotectedWriteTests(unittest.TestCase):
    """When Excel will not switch to manual calculation, the write window is
    unprotected: every cell we touch makes Excel recalculate.

    The code that slows the render loop down for that case was reading two
    names it also assigned, which made them local to the function and raised
    UnboundLocalError on the first read. It sat behind a flag that nothing
    ever set, so the slow path never engaged and the customer's workbook
    crawled.
    """

    def setUp(self):
        self._guard_probe = patch.object(
            us_excel.calc_mode_guard, "other_holder_alive", return_value=False
        )
        self._guard_probe.start()
        self._process_probe = patch.object(
            us_excel, "excel_process_running", return_value=True
        )
        self._process_probe.start()
        self._saved = (
            us_state._calc_manual_failed,
            us_state._calc_manual_notice_shown,
            us_state._calc_notice_shown,
            us_state._calc_failed_since,
            us_state._last_unprotected_write,
        )
        us_state._calc_manual_failed = False
        us_state._calc_manual_notice_shown = False
        us_state._calc_notice_shown = False
        us_state._calc_failed_since = 0.0
        us_state._last_unprotected_write = 0.0

    def tearDown(self):
        self._guard_probe.stop()
        self._process_probe.stop()
        (
            us_state._calc_manual_failed,
            us_state._calc_manual_notice_shown,
            us_state._calc_notice_shown,
            us_state._calc_failed_since,
            us_state._last_unprotected_write,
        ) = self._saved

    @staticmethod
    def _refusing_app():
        class RefusingApp:
            def __init__(self):
                self.screen_updating = True
                self.api = _Api()

            @property
            def version(self):
                return "16.0"

            @property
            def calculation(self):
                return -4105

            @calculation.setter
            def calculation(self, value):
                raise RuntimeError(
                    "Unable to set the Calculation property of the Application class"
                )

        class _Api:
            Calculation = -4105

        return RefusingApp()

    def _reset_session(self):
        us_state._calc_session.update(
            {"owned": False, "failed": False, "mode": None, "app": None}
        )

    def test_a_refused_manual_switch_is_recorded(self):
        self._reset_session()
        with patch.object(us_excel, "show_error"):
            us_excel._begin_excel_bulk_mode(self._refusing_app())
        self.assertTrue(us_state._calc_manual_failed)

    def test_the_first_refusal_of_a_spell_is_timestamped(self):
        self._reset_session()
        self.assertEqual(us_state._calc_failed_since, 0.0)
        us_excel._begin_excel_bulk_mode(self._refusing_app())
        self.assertGreater(us_state._calc_failed_since, 0.0)

    def test_a_retry_does_not_move_the_spell_start(self):
        self._reset_session()
        us_excel._begin_excel_bulk_mode(self._refusing_app())
        first = us_state._calc_failed_since
        self.assertGreater(first, 0.0)
        # The next cycle would be inside the retry cooldown, so no switch is
        # attempted and the spell start must not move.
        us_excel._begin_excel_bulk_mode(self._refusing_app())
        self.assertEqual(us_state._calc_failed_since, first)

    def test_the_slow_path_is_reachable(self):
        """The names the slow path reads must be bound in us_state, not as
        locals of render_cycle."""
        import inspect

        src = inspect.getsource(us_excel.render_cycle)
        self.assertIn("us_state._calc_manual_failed", src)
        self.assertIn("us_state._last_unprotected_write", src)
        self.assertIn("us_state._calc_manual_notice_shown", src)


class ExpiryRollerImportTests(unittest.TestCase):
    """The lazy import of us_excel has to come before the endless loop.

    A name assigned anywhere in a function is local to it for the whole
    function, so reading it before the assignment raises UnboundLocalError.
    The import used to sit after "while True", so it never ran, and the first
    expiry roll died before the workbook's date cell was updated.
    """

    def test_the_import_is_before_the_loop(self):
        import ast
        import inspect

        tree = ast.parse(inspect.getsource(us_chain.expiry_roller))
        import_line = None
        while_line = None
        for node in ast.walk(tree):
            if isinstance(node, ast.Import) and import_line is None:
                for alias in node.names:
                    if alias.asname == "_main":
                        import_line = node.lineno
            if isinstance(node, ast.While) and while_line is None:
                while_line = node.lineno
        self.assertIsNotNone(import_line)
        self.assertIsNotNone(while_line)
        self.assertLess(import_line, while_line)

    def test_it_is_readable_by_name(self):
        import us_chain

        self.assertTrue(callable(us_chain.expiry_roller))


if __name__ == "__main__":
    unittest.main()
