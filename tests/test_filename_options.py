import io
import sys
import unittest
from contextlib import redirect_stderr
from pathlib import Path
from unittest.mock import patch

import PurviewClassificationsReport as report


class FilenameOptionTests(unittest.TestCase):
    def test_filename_prefix_defaults_to_classifications(self):
        with patch.object(sys, "argv", ["report", "--data-source", "ALL"]):
            args = report.parse_args()
        self.assertEqual(args.filename_prefix, "classifications")

    def test_custom_filename_prefix_builds_output_path(self):
        with patch.object(sys, "argv", [
            "report", "--data-source", "ALL", "--filename-prefix", "FULL",
        ]):
            args = report.parse_args()
        self.assertEqual(
            report.report_output_path("reports", args.filename_prefix),
            Path("reports") / "FULL.xlsx",
        )

    def test_old_filename_option_is_rejected(self):
        with (
            patch.object(sys, "argv", [
                "report", "--data-source", "ALL", "--filename-suffix", "FULL",
            ]),
            redirect_stderr(io.StringIO()),
        ):
            with self.assertRaises(SystemExit) as raised:
                report.parse_args()
        self.assertEqual(raised.exception.code, 2)

    def test_empty_prefix_error_uses_new_option_terminology(self):
        with self.assertRaisesRegex(report.PurviewApiError, "filename prefix"):
            report.report_output_path("reports", " . ")


if __name__ == "__main__":
    unittest.main()
