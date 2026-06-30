# EU ETS and FuelEU Compliance Dashboard

This project turns the workbook `EU_ETS_FuelEU_Compliance_Calculator_10_1_1.xlsx` into a Railway-ready web dashboard.

## What it includes

- A browser-based dashboard styled from the provided HTML concept
- Editable calculator inputs with spreadsheet-style derived outputs
- Instant browser autosave with background server synchronization
- Cloudflare R2-compatible durable state storage
- Vessel-level and multi-voyage PDF compliance statements
- Bulk `.xlsx` and `.xlsm` imports for Voyage Inputs and matching reference-library sheets
- Optional remote workbook sync from a configured workbook URL
- Editable reference sheets for parameters, fuel factors, fleet, ports, flags, derogations, methodology, and formula guide
- A JavaScript calculation engine that mirrors the workbook logic for:
  - EU ETS scope and allowance exposure
  - FuelEU attained intensity
  - FuelEU compliance balance
  - FuelEU penalty
  - Vessel-level summary rollups

## Run locally

```bash
npm install
npm start
```

Open `http://localhost:3000`.

## Import an Excel workbook

Use **Import Excel** in the application header. The importer recognizes these workbook sheets:

- `Calculator`
- `EU ETS & Fuel EU Calculator`
- `Parameters`
- `Fuel_Reference`
- `Fleet_DB`
- `Port_DB`
- `Flag_States`
- `Derogations`
- `Methodology`
- `Formula_Guide`

The preview lets the user select which detected sheets to apply. Voyage Inputs can replace the current rows or merge/update matching rows. Reference-library sheets are replaced only when they are present and selected. Calculated Excel columns are not imported; the application recalculates all dashboard results from the imported input and library data before saving the updated state.

## Sync a remote workbook automatically

The server can poll a configured workbook URL and apply any changed workbook sheets to the shared dashboard state.

Set these server-side variables:

```text
WORKBOOK_SYNC_SOURCE_URL=<direct workbook download URL or SharePoint share link>
WORKBOOK_SYNC_INTERVAL_MS=300000
WORKBOOK_SYNC_HEADERS_JSON={"Cookie":"<SharePoint or Microsoft 365 auth cookie if required>"}
WORKBOOK_SYNC_AZURE_TENANT_ID=<optional Entra tenant ID for Graph app sync>
WORKBOOK_SYNC_AZURE_CLIENT_ID=<optional Entra app client ID for Graph app sync>
WORKBOOK_SYNC_AZURE_CLIENT_SECRET=<optional Entra app client secret for Graph app sync>
```

Notes:

- `WORKBOOK_SYNC_SOURCE_URL` is required to enable automatic sync.
- `WORKBOOK_SYNC_INTERVAL_MS` is optional. The default is `300000` (5 minutes).
- `WORKBOOK_SYNC_HEADERS_JSON` is optional and lets you pass request headers when the workbook source requires authentication.
- The three `WORKBOOK_SYNC_AZURE_*` variables are optional. When all three are present and `WORKBOOK_SYNC_SOURCE_URL` is a SharePoint share link, the server uses Microsoft Graph application auth and no longer depends on a manually refreshed user session link.
- The sync layer now understands SharePoint workbook pages. If the source returns SharePoint HTML instead of an `.xlsx` or `.xlsm`, the server will extract the signed `FileGetUrl` bootstrap and then download the workbook automatically.
- For protected SharePoint links, Railway still needs a valid authenticated request path. The simplest option is a direct signed `FileGetUrl`. The more durable option is Microsoft Graph application auth against the original SharePoint share link. A browser-cookie bootstrap via `WORKBOOK_SYNC_HEADERS_JSON` still works, but it remains session-dependent.
- If the source redirects to Microsoft login, the workbook sync status will report that access is still required.
- The dashboard recalculates its summaries from the imported calculator and library sheets. The dashboard output sheets themselves are not copied cell-for-cell because they are derived views, but they refresh automatically after each successful sync.

When enabled, the dashboard header shows workbook-sync status, the sync mode in use, and any temporary access-token expiry that was discovered from SharePoint. It also exposes a `Sync Workbook` button for an immediate manual refresh.

## Regenerate the workbook seed

The exporter reads the workbook path from `COMPLIANCE_WORKBOOK_PATH`. If that variable is not set, it falls back to the original local workbook path used during development.

```bash
set COMPLIANCE_WORKBOOK_PATH=C:\path\to\EU_ETS_FuelEU_Compliance_Calculator_10_1_1.xlsx
python scripts/export_compliance_workbook.py
```

## Deploy on Railway

- Root directory: repository root
- Start command: `npm start`
- Health check path: `/api/health`
- Node version: `>=20`

### Durable R2 storage

Add these server-side variables to the Railway service:

```text
R2_ACCOUNT_ID=<Cloudflare account ID>
R2_ACCESS_KEY_ID=<R2 API token access key>
R2_SECRET_ACCESS_KEY=<R2 API token secret>
R2_BUCKET_NAME=<bucket name>
R2_OBJECT_KEY=fuel-ets/dashboard-state.json
```

`R2_OBJECT_KEY` is optional. The application uses the value above by default.

When all four required R2 variables are present, the header displays `Saved to R2`. Without them, the application continues to autosave in the browser and uses a server-file fallback for local testing, but Railway container storage is not durable.

The R2 credentials are used only by the Node server and are never sent to the browser.

### Optional access protection

The public application and shared state API can be protected with browser Basic Authentication:

```text
APP_USERNAME=<shared application username>
APP_PASSWORD=<strong shared password>
```

When both values are set, Railway's health check remains public while every application page and data API requires authentication.

### Persistence API

- `GET /api/storage/status` reports whether durable R2 storage is active.
- `GET /api/state` retrieves the current shared dashboard snapshot.
- `PUT /api/state` stores the current dashboard snapshot.
- `GET /api/workbook-sync/status` reports remote workbook sync configuration and the last sync result.
- `POST /api/workbook-sync/run` triggers an immediate remote workbook sync.
- `POST /api/reports/compliance-statement` generates the vessel statement PDF.
- `POST /api/import/excel` parses a workbook and returns a sheet-by-sheet import preview.
