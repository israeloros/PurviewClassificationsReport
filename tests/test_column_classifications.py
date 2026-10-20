import io
import unittest
from contextlib import redirect_stderr
from unittest.mock import Mock, patch

import requests

import PurviewClassificationsReport as report


class ColumnClassificationTests(unittest.TestCase):
    def setUp(self):
        self.qualified_name = "s3://example-bucket/sample.csv"
        self.file = {
            "guid": "file",
            "typeName": "aws_s3_v2_object",
            "attributes": {"name": "sample.csv", "qualifiedName": self.qualified_name},
            "relationshipAttributes": {
                "attachedSchema": [{"guid": "schema", "typeName": "delimited_text_schema"}],
            },
        }
        self.schema = {
            "guid": "schema",
            "typeName": "delimited_text_schema",
            "attributes": {
                "name": "delimited_text_schema",
                "qualifiedName": self.qualified_name + "#delimited_text_schema",
            },
            "relationshipAttributes": {
                "columns": [{"guid": f"column-{index}"} for index in range(4)],
                "relatedDataSets": [{"guid": "unrelated-dataset"}],
            },
        }
        self.columns = [
            {
                "guid": f"column-{index}",
                "typeName": "delimited_text_column",
                "attributes": {
                    "name": f":csv/{index}",
                    "qualifiedName": self.qualified_name + f"#delimited_text_schema//{index}",
                    "logicalType": "string",
                },
                "relationshipAttributes": {
                    "referredSchema": {
                        "guid": "schema", "displayText": "delimited_text_schema",
                    },
                },
                "classifications": [{"typeName": classification}],
            }
            for index, classification in enumerate((
                "MICROSOFT.FINANCIAL.US.ABA_ROUTING_NUMBER",
                "MICROSOFT.PERSONAL.NAME",
                "MICROSOFT.PERSONAL.PHYSICALADDRESS",
                "MICROSOFT.GOVERNMENT.ARGENTINA.DNI_NUMBER",
            ))
        ]
        self.entities = {
            entity["guid"]: entity
            for entity in [self.file, self.schema] + self.columns
        }

    def retrieve(self, **kwargs):
        assets = kwargs.pop("assets", [{"id": "file"}])
        return list(report.list_classified_columns(
            Mock(), "https://test.purview.azure.com", {}, assets, **kwargs
        ))

    def respond(self, session, method, url, headers, params):
        guids = [value for key, value in params if key == "guid"]
        return {"entities": [self.entities[guid] for guid in guids]}

    def test_attached_schema_and_columns_are_fetched_in_batches(self):
        with patch.object(report, "request_json", side_effect=self.respond) as request:
            columns = self.retrieve(batch_size=2)

        self.assertEqual({column["id"] for column in columns}, {
            column["guid"] for column in self.columns
        })
        self.assertTrue(all(column["dataType"] == "string" for column in columns))
        self.assertTrue(all(column["parentGuid"] == "schema" for column in columns))
        requested = []
        for call in request.call_args_list:
            guids = [value for key, value in call.kwargs["params"] if key == "guid"]
            self.assertLessEqual(len(guids), 2)
            self.assertIn(("minExtInfo", "true"), call.kwargs["params"])
            self.assertIn(("ignoreRelationships", "false"), call.kwargs["params"])
            requested.extend(guids)
        self.assertEqual(len(requested), len(set(requested)))
        self.assertEqual(set(requested), set(self.entities))
        index = report.build_source_index([{
            "name": "s3-source",
            "properties": {"serviceUrl": "s3://example-bucket"},
        }])
        self.assertEqual(
            len(report.group_columns_by_source(columns, index)["s3-source"]), 4
        )

    def test_inline_columns_and_cycles_do_not_duplicate_results(self):
        self.columns[0]["relationshipAttributes"]["schema"] = {"guid": "schema"}

        def respond(*args, **kwargs):
            response = self.respond(*args, **kwargs)
            response["referredEntities"] = {
                column["guid"]: column for column in self.columns
            }
            return response

        with patch.object(report, "request_json", side_effect=respond) as request:
            columns = self.retrieve(assets=[{"id": "file"}, {"id": "column-0"}])
        self.assertEqual(len(columns), 4)
        requested = [
            value for call in request.call_args_list
            for key, value in call.kwargs["params"] if key == "guid"
        ]
        self.assertEqual(len(requested), len(set(requested)))

    def test_incomplete_referred_column_is_fetched_before_emitting(self):
        def respond(*args, **kwargs):
            response = self.respond(*args, **kwargs)
            if response["entities"] == [self.file]:
                response["referredEntities"] = {
                    "column-0": {
                        "guid": "column-0",
                        "typeName": "delimited_text_column",
                        "isIncomplete": True,
                        "attributes": {},
                    }
                }
            return response

        with patch.object(report, "request_json", side_effect=respond):
            columns = self.retrieve()
        self.assertEqual(len(columns), 4)
        self.assertEqual(
            next(column for column in columns if column["id"] == "column-0")["classification"],
            self.columns[0]["classifications"],
        )

    def test_legacy_schema_and_column_attribute_references_are_followed(self):
        self.file["relationshipAttributes"] = {}
        self.file["attributes"]["tabular_schema"] = {"guid": "schema"}
        self.schema["attributes"]["columns"] = self.schema["relationshipAttributes"].pop("columns")
        with patch.object(report, "request_json", side_effect=self.respond):
            self.assertEqual(len(self.retrieve()), 4)

    def test_direct_database_columns_remain_supported(self):
        self.columns[0]["typeName"] = "azure_sql_column"
        self.columns[0]["attributes"]["dataType"] = "varchar"
        self.columns[0]["relationshipAttributes"] = {
            "table": {"guid": "table", "displayText": "Customers"}
        }
        with patch.object(report, "request_json", side_effect=self.respond):
            columns = self.retrieve(assets=[{"id": "column-0"}])
        self.assertEqual(len(columns), 1)
        self.assertEqual(columns[0]["dataType"], "varchar")
        self.assertEqual(columns[0]["parentName"], "Customers")

    def test_unrelated_incomplete_entities_are_not_fetched(self):
        def respond(*args, **kwargs):
            response = self.respond(*args, **kwargs)
            response["referredEntities"] = {
                "process": {
                    "guid": "process", "typeName": "Process", "isIncomplete": True,
                }
            }
            return response

        with patch.object(report, "request_json", side_effect=respond):
            self.assertEqual(len(self.retrieve()), 4)

    def test_empty_assets_make_no_requests(self):
        with patch.object(report, "request_json") as request:
            self.assertEqual(self.retrieve(assets=[]), [])
        request.assert_not_called()

    def test_incomplete_tables_and_schemas_still_expand_columns(self):
        self.file["isIncomplete"] = True
        self.schema["isIncomplete"] = True
        stderr = io.StringIO()
        with (
            patch.object(report, "request_json", side_effect=self.respond),
            redirect_stderr(stderr),
        ):
            columns = self.retrieve(batch_size=2)
        self.assertEqual(len(columns), 4)
        self.assertEqual(stderr.getvalue(), "")

    def test_incomplete_columns_retain_available_classifications(self):
        self.columns[0]["isIncomplete"] = True
        stderr = io.StringIO()
        with (
            patch.object(report, "request_json", side_effect=self.respond),
            redirect_stderr(stderr),
        ):
            columns = self.retrieve()
        self.assertEqual(len(columns), 4)
        self.assertEqual(
            next(column for column in columns if column["id"] == "column-0")["classification"],
            self.columns[0]["classifications"],
        )
        self.assertEqual(stderr.getvalue(), "")

    def test_incomplete_entities_without_columns_are_not_retrieval_failures(self):
        self.file["isIncomplete"] = True
        self.file["relationshipAttributes"] = {}
        stderr = io.StringIO()
        with (
            patch.object(report, "request_json", side_effect=self.respond),
            redirect_stderr(stderr),
        ):
            self.assertEqual(self.retrieve(), [])
        self.assertEqual(stderr.getvalue(), "")

    def test_complete_entities_do_not_warn(self):
        stderr = io.StringIO()
        with (
            patch.object(report, "request_json", side_effect=self.respond),
            redirect_stderr(stderr),
        ):
            self.assertEqual(len(self.retrieve()), 4)
        self.assertEqual(stderr.getvalue(), "")

    def test_missing_column_details_raise_explicit_error(self):
        with patch.object(report, "request_json", return_value={"entities": []}):
            with self.assertRaisesRegex(report.PurviewApiError, "requested entity 'file'"):
                self.retrieve()

    def test_request_failures_propagate(self):
        with patch.object(
            report, "request_json", side_effect=report.PurviewApiError("Forbidden")
        ):
            with self.assertRaisesRegex(report.PurviewApiError, "Forbidden"):
                self.retrieve()

    def test_timed_out_batch_is_split_without_losing_columns(self):
        first_request = True

        def respond(*args, **kwargs):
            nonlocal first_request
            if first_request:
                first_request = False
                try:
                    raise requests.ReadTimeout("Batch read timed out")
                except requests.ReadTimeout as exc:
                    raise report.PurviewApiError("Batch failed") from exc
            return self.respond(*args, **kwargs)

        stderr = io.StringIO()
        with (
            patch.object(report, "request_json", side_effect=respond) as request,
            redirect_stderr(stderr),
        ):
            columns = self.retrieve(assets=[{"id": "file"}, {"id": "schema"}])
        self.assertEqual(len(columns), 4)
        self.assertIn("retrying the same entities", stderr.getvalue())
        self.assertTrue(all(
            len([value for key, value in call.kwargs["params"] if key == "guid"]) == 1
            for call in request.call_args_list[1:]
        ))

    def test_single_entity_timeout_remains_an_error(self):
        error = report.PurviewApiError("Single entity timed out")
        error.__cause__ = requests.ReadTimeout("Timeout")
        with patch.object(report, "request_json", side_effect=error):
            with self.assertRaisesRegex(report.PurviewApiError, "Single entity"):
                self.retrieve()


if __name__ == "__main__":
    unittest.main()
