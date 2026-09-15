# Purview Classifications Report

`PurviewClassificationsReport.py` generates an Excel (`.xlsx`) report that maps
Microsoft Purview data classifications to registered data sources. It queries
the Purview Data Map (scan and catalog search APIs), matches catalog assets to
registered data source registrations using endpoint/qualified-name metadata,
and produces a workbook with three sheets:

- **Data Sources** — registered sources in scope, with location, resource
  group, subscription, endpoint, match locators, and asset/classification
  counts.
- **Classification Summary** — classification name and asset count per data
  source.
- **Assets & Classifications** — one row per asset/classification pair, with
  qualified name, entity/asset type, description, and collection ID.

## Requirements

- Python 3.9+
- Azure credentials with permission to read the target Purview account (Data
  Reader / Data Source Administrator or equivalent), resolved via
  [`DefaultAzureCredential`](https://learn.microsoft.com/python/api/azure-identity/azure.identity.defaultazurecredential)
  (environment/service principal, managed identity, Azure CLI login, etc.)

Install dependencies:

```powershell
pip install -r requirements.txt
```

## Configuration

The script reads its Purview account name (and, if you use a service
principal, its credentials) from an environment file — `purview.env` by
default, or any file passed via `--env-file`.

1. Copy the sample file and fill in your values:

   ```powershell
   Copy-Item purview-sample.env purview.env
   ```

2. Edit `purview.env`:

   | Variable | Required | Description |
   |---|---|---|
   | `PURVIEW_ACCOUNT_NAME` | Yes | Name of the Purview account (used to build `https://<name>.purview.azure.com`). |
   | `AZURE_TENANT_ID` | If using a service principal | Entra ID tenant ID. |
   | `AZURE_CLIENT_ID` | If using a service principal | App registration (client) ID. |
   | `AZURE_CLIENT_SECRET` | If using a service principal | App registration client secret. |
   | `AZURE_SUBSCRIPTION_ID` | Optional | Subscription ID, if needed for context. |

   If you omit the `AZURE_*` service principal variables, `DefaultAzureCredential`
   falls back to other available credentials in order (e.g. environment
   variables already set in your shell, Managed Identity, `az login` session,
   Visual Studio/VS Code sign-in).

> ⚠️ **Security note:** `purview.env` holds live secrets once configured.
> Keep it out of source control (add it to `.gitignore`), never share it, and
> rotate the client secret immediately if it has ever been committed or
> exposed.

## Usage

Run the script from this folder with the Python interpreter:

```powershell
python PurviewClassificationsReport.py [options]
```

### Options

| Option | Description |
|---|---|
| `--data-source <name>` | Registered data source name to report on, or `ALL` for every registered source. Mutually exclusive with `--list-data-sources`. |
| `--list-data-sources` | List registered data source names on screen and exit. |
| `--output <path>` | Output `.xlsx` path (default: `purview-data-source-classifications.xlsx`). |
| `--env-file <path>` | Environment file to load (default: `purview.env`). |
| `--qualified-name-prefix <prefix>` | Optional `qualifiedName` prefix to match assets when the registration has no usable endpoint metadata. Requires a specific `--data-source` (not `ALL`). |
| `--page-size <1-1000>` | Number of catalog search results per request (default: 1000). |

### Examples

List all registered data sources:

```powershell
python PurviewClassificationsReport.py --list-data-sources
```

Generate a report for a single data source:

```powershell
python PurviewClassificationsReport.py --data-source "SqlServer-Prod"
```

Generate a report for every registered data source, with a custom output path:

```powershell
python PurviewClassificationsReport.py --data-source ALL --output "reports\allsourceclassifications.xlsx"
```

Use a non-default environment file:

```powershell
python PurviewClassificationsReport.py --data-source "SqlServer-Prod" --env-file ".\purview-prod.env"
```

Force asset matching via a qualified-name prefix (useful when a source's
registration metadata has no endpoint that matches catalog asset qualified
names):

```powershell
python PurviewClassificationsReport.py --data-source "DataLake-Raw" --qualified-name-prefix "https://mydatalake.dfs.core.windows.net/raw"
```

Reduce catalog search page size (e.g. to work around throttling):

```powershell
python PurviewClassificationsReport.py --data-source ALL --page-size 200
```

## License

This project is licensed under the [MIT License](LICENSE).

## Notes

- If the script reports assets it couldn't match to a registered data source,
  it prints a warning with the unmatched count; consider `--qualified-name-prefix`
  for that source.
- The generated workbook opens in Excel with filterable tables, frozen header
  rows, and styled headers on each sheet.
