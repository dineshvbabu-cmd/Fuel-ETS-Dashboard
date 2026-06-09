# EU ETS and FuelEU Compliance Dashboard

This project turns the workbook `EU_ETS_FuelEU_Compliance_Calculator_10_1_1.xlsx` into a Railway-ready web dashboard.

## What it includes

- A browser-based dashboard styled from the provided HTML concept
- Editable calculator inputs with spreadsheet-style derived outputs
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
