import unittest
from unittest.mock import Mock, patch

import PurviewClassificationsReport as report


class ModifiedTimeFilterTests(unittest.TestCase):
    def test_catalog_update_time_filter_is_used_for_each_window(self):
        for window, expected_range in (
            ("24h", "LAST_24H"),
            ("7d", "LAST_7D"),
            ("30d", "LAST_30D"),
        ):
            with self.subTest(window=window):
                with patch.object(
                    report, "request_json", return_value={"value": []}
                ) as request:
                    list(report.list_catalog_assets(
                        Mock(), "https://test.purview.azure.com", {}, 100, window
                    ))

                self.assertEqual(request.call_args.kwargs["json"]["filter"], {
                    "updateTime": expected_range,
                })

    def test_update_time_filter_is_preserved_across_pages(self):
        assets = [{"id": "recent-table"}, {"id": "recent-column"}]
        with patch.object(report, "request_json", side_effect=[
            {"value": assets[:1], "continuationToken": "next-page"},
            {"value": assets[1:]},
        ]) as request:
            results = list(report.list_catalog_assets(
                Mock(), "https://test.purview.azure.com", {}, 100, "24h"
            ))

        self.assertEqual(results, assets)
        self.assertEqual(request.call_count, 2)
        for call in request.call_args_list:
            self.assertEqual(call.kwargs["json"]["filter"], {
                "updateTime": "LAST_24H",
            })
        self.assertNotIn(
            "continuationToken", request.call_args_list[0].kwargs["json"]
        )
        self.assertEqual(
            request.call_args_list[1].kwargs["json"]["continuationToken"],
            "next-page",
        )

    def test_no_time_filter_is_sent_without_modified_within(self):
        with patch.object(
            report, "request_json", return_value={"value": [{"id": "table"}]}
        ) as request:
            results = list(report.list_catalog_assets(
                Mock(), "https://test.purview.azure.com", {}, 100
            ))

        self.assertEqual(results, [{"id": "table"}])
        self.assertNotIn("filter", request.call_args.kwargs["json"])


if __name__ == "__main__":
    unittest.main()
