import io
import unittest
from contextlib import redirect_stderr
from unittest.mock import Mock, patch

import requests

import PurviewClassificationsReport as report


class RequestRecoveryTests(unittest.TestCase):
    def test_connection_reset_retries_same_read_request(self):
        for method in ("GET", "POST"):
            with self.subTest(method=method):
                response = Mock()
                response.json.return_value = {"entities": [{"guid": "requested-guid"}]}
                session = Mock()
                session.request.side_effect = [
                    requests.ConnectionError("Connection reset"), response,
                ]
                stderr = io.StringIO()
                with patch.object(report, "sleep") as sleep, redirect_stderr(stderr):
                    result = report.request_json(
                        session, method, "https://test.example/entity", {},
                        params=[("guid", "requested-guid")],
                    )
                self.assertEqual(result, response.json.return_value)
                self.assertEqual(session.request.call_count, 2)
                self.assertEqual(
                    session.request.call_args_list[0], session.request.call_args_list[1]
                )
                sleep.assert_called_once_with(1)
                self.assertIn("retrying", stderr.getvalue())

    def test_exhausted_connection_retries_raise_explicit_error(self):
        session = Mock()
        error = requests.ConnectionError("Connection reset")
        session.request.side_effect = error
        with patch.object(report, "sleep") as sleep, redirect_stderr(io.StringIO()):
            with self.assertRaises(report.PurviewApiError) as raised:
                report.request_json(session, "GET", "https://test.example/entity", {})
        self.assertIs(raised.exception.__cause__, error)
        self.assertEqual(session.request.call_count, 3)
        self.assertEqual([call.args[0] for call in sleep.call_args_list], [1, 2])

    def test_http_permission_failures_are_not_retried(self):
        session = Mock()
        response = Mock()
        response.text = "Forbidden"
        response.raise_for_status.side_effect = requests.HTTPError(
            "Forbidden", response=response
        )
        session.request.return_value = response
        with patch.object(report, "sleep") as sleep:
            with self.assertRaisesRegex(report.PurviewApiError, "Forbidden"):
                report.request_json(session, "GET", "https://test.example/entity", {})
        self.assertEqual(session.request.call_count, 1)
        sleep.assert_not_called()

    def test_read_timeout_remains_available_to_batch_splitter(self):
        session = Mock()
        error = requests.ReadTimeout("Timed out")
        session.request.side_effect = error
        with self.assertRaises(report.PurviewApiError) as raised:
            report.request_json(session, "GET", "https://test.example/entity", {})
        self.assertIs(raised.exception.__cause__, error)
        self.assertEqual(session.request.call_count, 1)

    def test_certificate_errors_are_not_retried(self):
        session = Mock()
        session.request.side_effect = requests.exceptions.SSLError("Invalid certificate")
        with patch.object(report, "sleep") as sleep:
            with self.assertRaises(report.PurviewApiError):
                report.request_json(session, "GET", "https://test.example/entity", {})
        self.assertEqual(session.request.call_count, 1)
        sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()
