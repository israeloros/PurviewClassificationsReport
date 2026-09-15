import argparse
import os
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from azure.core.exceptions import AzureError
from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


API_VERSION = "2023-09-01"
DEFAULT_ENV_FILE = "purview.env"
DEFAULT_OUTPUT_FILE = "purview-data-source-classifications.xlsx"
REQUEST_TIMEOUT_SECONDS = 60
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SELECTED_FILL = PatternFill("solid", fgColor="D9EAD3")
HEADER_FONT = Font(name="Arial", color="FFFFFF", bold=True)
BODY_FONT = Font(name="Arial", color="000000")


class PurviewApiError(RuntimeError):
    pass


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Create an XLSX report of Microsoft Purview classifications by "
            "registered data source."
        )
    )
    source_selection = parser.add_mutually_exclusive_group(required=True)
    source_selection.add_argument(
        "--data-source",
        help=(
            "Registered Purview data source name to include in the report, or ALL "
            "to include every registered data source."
        ),
    )
    source_selection.add_argument(
        "--list-data-sources",
        action="store_true",
        help="Enumerate registered Purview data sources on screen and exit.",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT_FILE,
        help=f"Output XLSX path (default: {DEFAULT_OUTPUT_FILE}).",
    )
    parser.add_argument(
        "--env-file",
        default=DEFAULT_ENV_FILE,
        help=f"Environment file containing PURVIEW_ACCOUNT_NAME (default: {DEFAULT_ENV_FILE}).",
    )
    parser.add_argument(
        "--qualified-name-prefix",
        help=(
            "Optional qualifiedName prefix used to identify assets for the selected "
            "data source. Use this when the registration does not expose a usable endpoint."
        ),
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=1000,
        choices=range(1, 1001),
        metavar="[1-1000]",
        help="Purview discovery results per request (default: 1000).",
    )
    args = parser.parse_args()
    if args.qualified_name_prefix and (
        not args.data_source or args.data_source.casefold() == "all"
    ):
        parser.error(
            "--qualified-name-prefix requires a specific --data-source value."
        )
    return args


def request_json(session, method, url, headers, **kwargs):
    try:
        response = session.request(
            method,
            url,
            headers=headers,
            timeout=REQUEST_TIMEOUT_SECONDS,
            **kwargs,
        )
        response.raise_for_status()
        return response.json()
    except requests.RequestException as exc:
        details = ""
        if exc.response is not None:
            details = f" Response: {exc.response.text[:1000]}"
        raise PurviewApiError(f"{method} {url} failed: {exc}.{details}") from exc
    except ValueError as exc:
        raise PurviewApiError(f"{method} {url} returned invalid JSON.") from exc


def list_data_sources(session, endpoint, headers):
    url = f"{endpoint}/scan/datasources?api-version={API_VERSION}"
    data_sources = []

    while url:
        response = request_json(session, "GET", url, headers)
        data_sources.extend(response.get("value", []))
        next_link = response.get("nextLink")
        url = urljoin(endpoint, next_link) if next_link else None

    return data_sources


def list_catalog_assets(session, endpoint, headers, page_size):
    url = f"{endpoint}/datamap/api/search/query?api-version={API_VERSION}"
    continuation_token = None

    while True:
        body = {"keywords": None, "limit": page_size}
        if continuation_token:
            body["continuationToken"] = continuation_token

        response = request_json(session, "POST", url, headers, json=body)
        yield from response.get("value", [])

        continuation_token = response.get("continuationToken")
        if not continuation_token:
            break


def normalize_locator(value):
    text = str(value or "").strip().lower().replace("\\", "/")
    if not text:
        return ""

    parsed = urlparse(text if "://" in text else f"//{text}")
    if parsed.netloc:
        text = f"{parsed.netloc}{parsed.path}"

    return text.strip("/")


def iter_nested_values(value, key_path=()):
    if isinstance(value, dict):
        for key, child in value.items():
            yield from iter_nested_values(child, key_path + (str(key),))
    elif isinstance(value, list):
        for child in value:
            yield from iter_nested_values(child, key_path)
    elif isinstance(value, str):
        yield key_path, value


def data_source_locators(data_source):
    locator_keys = {
        "endpoint",
        "host",
        "hostname",
        "qualifiedname",
        "resourceuri",
        "serverendpoint",
        "url",
    }
    locators = set()

    for key_path, value in iter_nested_values(data_source.get("properties", {})):
        if key_path and key_path[-1].lower() in locator_keys:
            locator = normalize_locator(value)
            if locator:
                locators.add(locator)

    properties = data_source.get("properties") or {}
    endpoint = source_property(
        data_source,
        "endpoint",
        "serverEndpoint",
        "resourceUri",
        "url",
    )
    for child_name_key in ("databaseName", "resourceName", "warehouseName"):
        child_name = properties.get(child_name_key)
        if endpoint and isinstance(child_name, str) and child_name.strip():
            locators.add(
                normalize_locator(
                    f"{normalize_locator(endpoint)}/{child_name.strip('/')}"
                )
            )

    return sorted(locators, key=len, reverse=True)


def locator_matches(qualified_name, locator):
    qualified_name = normalize_locator(qualified_name)
    locator = normalize_locator(locator)
    if not qualified_name or not locator:
        return False

    return (
        qualified_name == locator
        or qualified_name.startswith(f"{locator}/")
        or f"/{locator}/" in f"/{qualified_name}/"
    )


def build_source_index(data_sources, selected_name=None, selected_prefix=None):
    source_index = []
    for source in data_sources:
        locators = data_source_locators(source)
        if (
            selected_name
            and source.get("name", "").casefold() == selected_name.casefold()
            and selected_prefix
        ):
            normalized_prefix = normalize_locator(selected_prefix)
            if normalized_prefix:
                locators = [normalized_prefix] + [
                    locator for locator in locators if locator != normalized_prefix
                ]
        source_index.append((source, locators))
    return sorted(
        source_index,
        key=lambda item: (
            bool(selected_name)
            and item[0].get("name", "").casefold() != selected_name.casefold()
        ),
    )


def match_asset_to_source(asset, source_index):
    qualified_name = asset.get("qualifiedName")
    matches = []

    for source, locators in source_index:
        for locator in locators:
            if locator_matches(qualified_name, locator):
                matches.append((len(locator), source))
                break

    if not matches:
        return None

    matches.sort(key=lambda item: item[0], reverse=True)
    return matches[0][1]


def classification_names(asset):
    names = set()
    for classification in asset.get("classification") or []:
        if isinstance(classification, str):
            name = classification
        elif isinstance(classification, dict):
            name = classification.get("typeName") or classification.get("name")
        else:
            name = None
        if name:
            names.add(name)
    return sorted(names, key=str.casefold)


def source_property(source, *keys):
    properties = source.get("properties") or {}
    for key in keys:
        value = properties.get(key)
        if value not in (None, ""):
            if isinstance(value, dict):
                return (
                    value.get("referenceName")
                    or value.get("name")
                    or str(value)
                )
            return value
    return ""


def aggregate_assets(data_sources, assets, source_index):
    assets_by_source = defaultdict(list)
    unmatched_count = 0

    for asset in assets:
        source = match_asset_to_source(asset, source_index)
        if source is None:
            unmatched_count += 1
            continue
        assets_by_source[source.get("name", "")].append(asset)

    summaries = {}
    for source in data_sources:
        source_name = source.get("name", "")
        source_assets = assets_by_source[source_name]
        classification_counts = Counter()
        classified_assets = 0

        for asset in source_assets:
            names = classification_names(asset)
            if names:
                classified_assets += 1
                classification_counts.update(names)
            else:
                classification_counts["(Unclassified)"] += 1

        summaries[source_name] = {
            "matched_assets": len(source_assets),
            "classified_assets": classified_assets,
            "classification_assignments": sum(
                count
                for name, count in classification_counts.items()
                if name != "(Unclassified)"
            ),
            "classification_counts": classification_counts,
        }

    return assets_by_source, summaries, unmatched_count


def add_table(worksheet, name):
    if worksheet.max_row < 2:
        worksheet.append(["No records found"] + [""] * (worksheet.max_column - 1))

    table = Table(
        displayName=name,
        ref=f"A1:{worksheet.cell(worksheet.max_row, worksheet.max_column).coordinate}",
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)


def format_worksheet(worksheet, widths):
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.sheet_view.showGridLines = False
    worksheet.row_dimensions[1].height = 28

    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width


def create_workbook(
    output_path,
    report_sources,
    report_scope,
    source_index,
    assets_by_source,
    summaries,
):
    workbook = Workbook()
    data_sources_sheet = workbook.active
    data_sources_sheet.title = "Data Sources"
    summary_sheet = workbook.create_sheet("Classification Summary")
    assets_sheet = workbook.create_sheet("Assets & Classifications")

    data_sources_sheet.append(
        [
            "Selected",
            "Data Source",
            "Kind",
            "ID",
            "Collection",
            "Location",
            "Resource Group",
            "Subscription ID",
            "Resource Name",
            "Endpoint",
            "Asset Match Locators",
            "Matched Assets",
            "Classified Assets",
            "Classification Assignments",
        ]
    )
    locator_map = {
        source.get("name", ""): locators for source, locators in source_index
    }
    for source in sorted(
        report_sources, key=lambda item: item.get("name", "").casefold()
    ):
        source_name = source.get("name", "")
        summary = summaries[source_name]
        data_sources_sheet.append(
            [
                "Yes",
                source_name,
                source.get("kind", ""),
                source.get("id", ""),
                source_property(source, "collection"),
                source_property(source, "location"),
                source_property(source, "resourceGroup"),
                source_property(source, "subscriptionId"),
                source_property(source, "resourceName"),
                source_property(
                    source,
                    "endpoint",
                    "serverEndpoint",
                    "resourceUri",
                    "url",
                ),
                "\n".join(locator_map[source_name]),
                summary["matched_assets"],
                summary["classified_assets"],
                summary["classification_assignments"],
            ]
        )
        for cell in data_sources_sheet[data_sources_sheet.max_row]:
            cell.fill = SELECTED_FILL

    summary_sheet.append(
        [
            "Data Source",
            "Kind",
            "Classification",
            "Asset Count",
        ]
    )
    source_by_name = {
        source.get("name", ""): source for source in report_sources
    }
    for source_name in sorted(source_by_name, key=str.casefold):
        counts = summaries[source_name]["classification_counts"]
        if not counts:
            summary_sheet.append(
                [
                    source_name,
                    source_by_name[source_name].get("kind", ""),
                    "(No matched assets)",
                    0,
                ]
            )
            continue
        for classification, count in sorted(
            counts.items(), key=lambda item: item[0].casefold()
        ):
            summary_sheet.append(
                [
                    source_name,
                    source_by_name[source_name].get("kind", ""),
                    classification,
                    count,
                ]
            )

    assets_sheet.append(
        [
            "Data Source",
            "Asset Name",
            "Asset GUID",
            "Entity Type",
            "Asset Type",
            "Qualified Name",
            "Classification",
            "Description",
            "Collection ID",
        ]
    )
    for source in sorted(
        report_sources, key=lambda item: item.get("name", "").casefold()
    ):
        source_name = source.get("name", "")
        source_assets = sorted(
            assets_by_source[source_name],
            key=lambda asset: (
                str(asset.get("qualifiedName", "")).casefold(),
                str(asset.get("name", "")).casefold(),
            ),
        )
        for asset in source_assets:
            classifications = classification_names(asset)
            if not classifications:
                continue
            for classification in classifications:
                asset_type = asset.get("assetType") or []
                if isinstance(asset_type, list):
                    asset_type = ", ".join(str(value) for value in asset_type)
                assets_sheet.append(
                    [
                        source_name,
                        asset.get("name", ""),
                        asset.get("id", ""),
                        asset.get("entityType", ""),
                        asset_type,
                        asset.get("qualifiedName", ""),
                        classification,
                        asset.get("description", ""),
                        asset.get("collectionId", ""),
                    ]
                )

    add_table(data_sources_sheet, "DataSourcesTable")
    add_table(summary_sheet, "ClassificationSummaryTable")
    add_table(assets_sheet, "AssetsClassificationsTable")
    format_worksheet(
        data_sources_sheet,
        {
            "A": 10,
            "B": 28,
            "C": 25,
            "D": 42,
            "E": 24,
            "F": 18,
            "G": 22,
            "H": 38,
            "I": 28,
            "J": 45,
            "K": 55,
            "L": 16,
            "M": 18,
            "N": 25,
        },
    )
    format_worksheet(
        summary_sheet,
        {"A": 30, "B": 28, "C": 45, "D": 16},
    )
    format_worksheet(
        assets_sheet,
        {
            "A": 28,
            "B": 35,
            "C": 38,
            "D": 30,
            "E": 30,
            "F": 70,
            "G": 45,
            "H": 55,
            "I": 38,
        },
    )

    generated_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    workbook.properties.title = "Microsoft Purview data source classifications"
    workbook.properties.subject = (
        f"Classification inventory for data source scope {report_scope}"
    )
    workbook.properties.creator = "PurviewDataSourceClassifications.py"
    workbook.properties.description = (
        f"Generated {generated_at}. Assets are mapped to registered data sources "
        "using the longest endpoint or qualified-name locator from registration metadata."
    )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def main():
    args = parse_args()
    load_dotenv(args.env_file)
    account_name = os.getenv("PURVIEW_ACCOUNT_NAME")
    if not account_name:
        print(
            f"PURVIEW_ACCOUNT_NAME was not found in {args.env_file} or the environment.",
            file=sys.stderr,
        )
        return 1

    endpoint = f"https://{account_name}.purview.azure.com"
    try:
        credential = DefaultAzureCredential()
        access_token = credential.get_token("https://purview.azure.net/.default").token
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        with requests.Session() as session:
            print("Retrieving registered data sources...")
            data_sources = list_data_sources(session, endpoint, headers)
            if args.list_data_sources:
                sorted_sources = sorted(
                    data_sources,
                    key=lambda source: source.get("name", "").casefold(),
                )
                if not sorted_sources:
                    print("No registered data sources found.")
                for number, source in enumerate(sorted_sources, start=1):
                    kind = source.get("kind", "")
                    kind_suffix = f" ({kind})" if kind else ""
                    print(f"{number}. {source.get('name', '')}{kind_suffix}")
                return 0

            include_all_sources = args.data_source.casefold() == "all"
            selected_source = None
            if not include_all_sources:
                selected_source = next(
                    (
                        source
                        for source in data_sources
                        if source.get("name", "").casefold()
                        == args.data_source.casefold()
                    ),
                    None,
                )
            if not include_all_sources and selected_source is None:
                available = ", ".join(
                    sorted(
                        (source.get("name", "") for source in data_sources),
                        key=str.casefold,
                    )
                )
                raise PurviewApiError(
                    f"Data source '{args.data_source}' was not found. "
                    f"Available data sources: {available or '(none)'}"
                )

            report_sources = (
                data_sources if include_all_sources else [selected_source]
            )
            source_index = build_source_index(
                data_sources,
                None if include_all_sources else selected_source.get("name", ""),
                args.qualified_name_prefix,
            )
            if selected_source is not None:
                selected_locators = next(
                    locators
                    for source, locators in source_index
                    if source is selected_source
                )
                if not selected_locators:
                    raise PurviewApiError(
                        "The selected registration has no endpoint metadata that can be "
                        "matched to asset qualified names. Rerun with "
                        "--qualified-name-prefix."
                    )

            print("Retrieving catalog assets...")
            assets = list(
                list_catalog_assets(session, endpoint, headers, args.page_size)
            )

        assets_by_source, summaries, unmatched_count = aggregate_assets(
            data_sources,
            assets,
            source_index,
        )
        create_workbook(
            args.output,
            report_sources,
            "ALL" if include_all_sources else selected_source.get("name", ""),
            source_index,
            assets_by_source,
            summaries,
        )

        report_asset_count = sum(
            len(assets_by_source[source.get("name", "")])
            for source in report_sources
        )
        report_scope = (
            "all registered data sources"
            if include_all_sources
            else f"'{selected_source.get('name', '')}'"
        )
        print(
            f"Created {Path(args.output).resolve()} with {report_asset_count} assets "
            f"for {report_scope}."
        )
        if unmatched_count:
            print(
                f"Warning: {unmatched_count} catalog assets could not be mapped to a "
                "registered data source using registration endpoint metadata.",
                file=sys.stderr,
            )
        return 0
    except (AzureError, PurviewApiError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
