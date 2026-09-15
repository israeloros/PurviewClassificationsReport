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
- An existing Microsoft Purview account
- Azure credentials authorized for the target Purview account, resolved via
  [`DefaultAzureCredential`](https://learn.microsoft.com/python/api/azure-identity/azure.identity.defaultazurecredential)
  (environment/service principal, managed identity, Azure CLI login, etc.)

Install dependencies:

```powershell
pip install -r requirements.txt
```

### Service principal and Purview permissions

A service principal is the recommended authentication method for scheduled,
unattended, or automated report generation. For local interactive use, a
service principal is optional because `DefaultAzureCredential` can also use an
authenticated Azure CLI, Visual Studio, or Visual Studio Code session.

To use a service principal:

1. Create a Microsoft Entra application and service principal, then create a
   client secret. See:
   [API authentication for Microsoft Purview data planes](https://learn.microsoft.com/purview/data-gov-api-rest-data-plane)
   and
   [Create Azure service principals using the Azure CLI](https://learn.microsoft.com/cli/azure/azure-cli-sp-tutorial-1).
2. In the Microsoft Purview governance portal, assign the service principal
   these Data Map roles:
   - **Data Curator** — required to query the catalog data plane.
   - **Data Source Administrator** — required to enumerate registered data
     sources through the scanning data plane.
3. Assign the roles at the root collection to report across the entire Data
   Map, or at the appropriate collection when the report should be restricted
   to that collection and its descendants. A **Collection Admin** must perform
   the role assignments.
4. Set `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET` in
   `purview.env`. `DefaultAzureCredential` automatically uses these values.

For more information, see:

- [Microsoft Purview data-plane API authentication](https://learn.microsoft.com/purview/data-gov-api-rest-data-plane)
- [Azure Identity client library for Python](https://learn.microsoft.com/python/api/overview/azure/identity-readme)
- [Service principal authentication for Microsoft Purview](https://learn.microsoft.com/purview/data-map-service-principal)

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
| `--filename-suffix <suffix>` | Filename suffix placed before the `.xlsx` extension (default: `classifications`). |
| `--output-directory <path>` | Directory where generated files are stored (default: `reports`). |
| `--file-per-data-source` | Generate a separate workbook for each data source in scope and prefix each filename with its data source name. Requires `--data-source`. |
| `--env-file <path>` | Environment file to load (default: `purview.env`). |
| `--qualified-name-prefix <prefix>` | Optional `qualifiedName` prefix to match assets when the registration has no usable endpoint metadata. Requires a specific `--data-source` (not `ALL`). |
| `--page-size <1-1000>` | Number of catalog search results per request (default: 1000). |
| `--modified-within <24h\|7d\|30d>` | Only include catalog assets modified within the previous 24 hours, 7 days, or 30 days. |

### Examples

List all registered data sources:

```powershell
python PurviewClassificationsReport.py --list-data-sources
```

Generate a report for a single data source:

```powershell
python PurviewClassificationsReport.py --data-source "SqlServer-Prod"
```

Generate a combined report for every registered data source. By default, this
creates `reports\classifications.xlsx`:

```powershell
python PurviewClassificationsReport.py --data-source ALL
```

Generate a separate report for every registered data source:

```powershell
python PurviewClassificationsReport.py --data-source ALL --file-per-data-source
```

This creates files such as `reports\SqlServer-Prod-classifications.xlsx`. Data
source characters that are invalid in Windows filenames are replaced with
underscores.

Use a custom output directory and filename suffix:

```powershell
python PurviewClassificationsReport.py --data-source ALL --file-per-data-source --output-directory "exports" --filename-suffix "classification-inventory"
```

This creates files such as
`exports\SqlServer-Prod-classification-inventory.xlsx`.

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

Generate separate reports containing only assets modified in the previous
seven days:

```powershell
python PurviewClassificationsReport.py --data-source ALL --file-per-data-source --modified-within 7d
```

## License

This project is licensed under the [MIT License](LICENSE).

## Notes

- If the script reports assets it couldn't match to a registered data source,
  it prints a warning with the unmatched count; consider `--qualified-name-prefix`
  for that source.
- `--modified-within` filters on the Purview catalog asset's `modifiedTime`.
  It does not represent the time when an individual classification was assigned.
- The generated workbook opens in Excel with filterable tables, frozen header
  rows, and styled headers on each sheet.
