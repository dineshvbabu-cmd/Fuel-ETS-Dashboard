# EU ETS and FuelEU Compliance Dashboard

This project turns the workbook `EU_ETS_FuelEU_Compliance_Calculator_10_1_1.xlsx` into a Railway-ready web dashboard.

## What it includes

- A browser-based dashboard styled from the provided HTML concept
- Editable calculator inputs with spreadsheet-style derived outputs
- Instant browser autosave with background server synchronization
- Cloudflare R2-compatible durable state storage
- Vessel-level and multi-voyage PDF compliance statements
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
- `POST /api/reports/compliance-statement` generates the vessel statement PDF.
