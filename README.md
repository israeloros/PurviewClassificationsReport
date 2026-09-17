# Purview Classifications Report

`PurviewClassificationsReport.py` generates an Excel (`.xlsx`) report that maps
Microsoft Purview data classifications to registered data sources. It queries
the Purview Data Map (scan and catalog search APIs), matches catalog assets to
registered data source registrations using endpoint/qualified-name metadata,
and produces a workbook with four sheets:

- **Data Sources** — registered sources in scope, with location, resource
  group, subscription, endpoint, match locators, and asset/classification
  counts.
- **Classification Summary** — classification name and asset count per data
  source.
- **Assets & Classifications** — one row per asset/classification pair, with
  qualified name, entity/asset type, description, and collection ID.
- **Column Classifications** — one row per classification assigned to a column
  asset, with column name, GUID, entity/data type, parent asset, qualified
  name, description, and collection ID. Column details are expanded through
  the Atlas entity API because discovery search results do not reliably
  include child columns.

## Requirements

- Python 3.9+
- An existing Microsoft Purview account
- Azure credentials authorized for the target Purview account. The script can
  use [`DefaultAzureCredential`](https://learn.microsoft.com/python/api/azure-identity/azure.identity.defaultazurecredential)
  (managed identity, Azure CLI login, Visual Studio/VS Code, and other
  supported credentials) or an explicitly configured service principal.

Install dependencies:

```powershell
pip install -r requirements.txt
```

## Authentication and Purview access

Whichever authentication method you choose, the Microsoft Entra identity used
by the script must have these Microsoft Purview Data Map roles:

- **Data Reader** — permits catalog searches and reading classifications.
- **Data Source Administrator** — permits listing registered data sources
  through the scanning data plane.

In the Microsoft Purview governance portal, open **Data Map > Collections**,
select the collection, and assign both roles to the user, managed identity, or
service principal that will run the script. Assign them at the root collection
to report across the entire Data Map, or at a child collection to restrict
access to that collection and its descendants. A **Collection Admin** must
perform these assignments.

### Option 1: Azure credentials

Use this option for interactive local execution or when the script runs on an
Azure resource with a managed identity. It is the default authentication mode
and uses `DefaultAzureCredential`.

For local execution with Azure CLI credentials:

1. Install the [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli).
2. Open PowerShell and confirm that the Azure CLI is available:

   ```powershell
   az version
   ```

3. Sign in to the Microsoft Entra tenant containing the Purview account.
   Replace `<tenant-id>` with the `AZURE_TENANT_ID` value from `purview.env`:

   ```powershell
   az login --tenant "<tenant-id>"
   ```

   A browser window opens. Sign in with the Azure account that has been
   granted access to Purview. If a browser cannot be opened, use device-code
   authentication instead:

   ```powershell
   az login --tenant "<tenant-id>" --use-device-code
   ```

   Follow the displayed instructions to open
   [https://microsoft.com/devicelogin](https://microsoft.com/devicelogin),
   enter the generated code, and complete sign-in.

4. Select the subscription containing the Purview account. Replace
   `<subscription-id>` with the `AZURE_SUBSCRIPTION_ID` value from
   `purview.env`:

   ```powershell
   az account set --subscription "<subscription-id>"
   ```

5. Verify the active Azure account, tenant, and subscription:

   ```powershell
   az account show --output table
   ```

   Confirm that the displayed tenant and subscription match the values in
   `purview.env`. If the wrong user is signed in, run `az logout`, then repeat
   the login steps.
6. Ask a Purview **Collection Admin** to assign your signed-in user the
   **Data Reader** and **Data Source Administrator** roles described above.
7. For Azure credential authentication, `purview.env` only needs these three
   settings:

   ```dotenv
   PURVIEW_ACCOUNT_NAME="<purview-account-name>"
   AZURE_TENANT_ID="<tenant-id>"
   AZURE_SUBSCRIPTION_ID="<subscription-id>"
   ```

   Do not include `AZURE_CLIENT_ID` or `AZURE_CLIENT_SECRET` when using Azure
   CLI, Visual Studio, Visual Studio Code, Azure PowerShell, or managed
   identity credentials. Those variables are for service-principal
   authentication and may cause Azure Identity to attempt the wrong credential
   type when they are present.
8. Run the report. The `--authentication-mode` option can be omitted because
   `azure-credential` is the default:

   ```powershell
   python PurviewClassificationsReport.py --data-source ALL --authentication-mode azure-credential
   ```

`DefaultAzureCredential` can also use supported Visual Studio, Visual Studio
Code, Azure PowerShell, and managed identity credentials. For managed identity,
enable the identity on the Azure resource and assign that identity the same
Purview roles before running the command above.

### Option 2: Service principal

Use a service principal for scheduled, unattended, or automated report
generation:

1. Create a Microsoft Entra application and service principal, then create a
   client secret. See
   [Create Azure service principals using the Azure CLI](https://learn.microsoft.com/cli/azure/azure-cli-sp-tutorial-1).
2. Ask a Purview **Collection Admin** to assign the service principal the
   **Data Reader** and **Data Source Administrator** roles described above.
3. Configure `purview.env` with the Purview account and service-principal
   credentials:

   ```dotenv
   PURVIEW_ACCOUNT_NAME="<purview-account-name>"
   AZURE_TENANT_ID="<tenant-id>"
   AZURE_CLIENT_ID="<application-client-id>"
   AZURE_CLIENT_SECRET="<client-secret>"
   AZURE_SUBSCRIPTION_ID="<subscription-id>"
   ```

4. Run the report with the service-principal authentication mode:

   ```powershell
   python PurviewClassificationsReport.py --data-source ALL --authentication-mode service-principal
   ```

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

2. Edit `purview.env`. Include only the variables required by the selected
   authentication method:

   | Variable | Required | Description |
   |---|---|---|
   | `PURVIEW_ACCOUNT_NAME` | Yes | Name of the Purview account (used to build `https://<name>.purview.azure.com`). |
   | `AZURE_TENANT_ID` | Yes | Entra ID tenant containing the Purview account. |
   | `AZURE_CLIENT_ID` | If using a service principal | App registration (client) ID. |
   | `AZURE_CLIENT_SECRET` | If using a service principal | App registration client secret. |
   | `AZURE_SUBSCRIPTION_ID` | Yes | Subscription containing the Purview account; used when selecting the Azure CLI subscription. |

   With `--authentication-mode azure-credential`, the script uses
   `DefaultAzureCredential` to select an available Azure identity. The file
   should contain only `PURVIEW_ACCOUNT_NAME`, `AZURE_TENANT_ID`, and
   `AZURE_SUBSCRIPTION_ID`; omit `AZURE_CLIENT_ID` and `AZURE_CLIENT_SECRET`.
   With `--authentication-mode service-principal`, add both
   service-principal variables.

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
| `--authentication-mode <azure-credential\|service-principal>` | Use the Azure Identity default credential chain or the service principal configured in the environment file (default: `azure-credential`). |
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

Generate a report using Azure CLI, developer-tool, or managed identity
credentials:

```powershell
python PurviewClassificationsReport.py --data-source ALL --authentication-mode azure-credential
```

Generate a report using a service principal configured in `purview.env`:

```powershell
python PurviewClassificationsReport.py --data-source ALL --authentication-mode service-principal
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
