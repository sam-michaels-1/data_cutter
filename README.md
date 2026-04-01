# Data Cutter

A client-side web application that converts raw ARR/Revenue Excel data into a professional multi-tab analysis workbook, complete with retention metrics, cohort analysis, and interactive dashboards.

## Overview

Upload a raw Excel file containing customer-level ARR or revenue data, map your columns through a guided wizard, and Data Cutter will generate a formula-driven Excel workbook with retention waterfall analysis, cohort tracking, and top customer breakdowns — all at your chosen time granularity. A live dashboard lets you explore the data interactively before downloading.

Everything runs in the browser — no backend or server required.

## Features

- **Guided import wizard** — 7-step flow to upload data, detect format, and configure output
- **Auto-detection** — Automatically infers date columns, customer IDs, ARR values, scale factors, and data frequency
- **Multi-granularity output** — Generate monthly, quarterly, and/or annual views from a single upload
- **Formula-driven workbook** — Output uses Excel formulas, not static values, so the workbook stays live
- **Retention metrics** — Lost-only, Punitive, and Net Dollar Retention calculated per period
- **Cohort analysis** — Logo and NDR cohort heatmaps grouped by acquisition period
- **Top customers** — Customer ranking, concentration metrics, and status tracking
- **Live dashboard** — Filter by attribute values and explore metrics before downloading

## Project Structure

```
Data Cutter/
└── frontend/                 # React + TypeScript + Vite
    └── src/
        ├── App.tsx           # Router (5 pages)
        ├── api/              # API client and dashboard data fetcher
        ├── engine/           # Excel generation engine (TypeScript)
        │   ├── generator.ts  # Main orchestrator — builds the output workbook
        │   ├── clean_data.ts # Raw and aggregated data tabs
        │   ├── retention.ts  # Retention metric tabs
        │   ├── cohort.ts     # Cohort analysis tabs
        │   ├── top_customers.ts # Top customers tab
        │   ├── formatting.ts # Cell styling and formatting
        │   ├── detect.ts     # Column role auto-detection
        │   ├── config_builder.ts # Wizard config → engine config
        │   ├── compute.ts    # Dashboard metric computation
        │   ├── utils.ts      # Formula builders and helpers
        │   └── types.ts      # Shared type definitions
        ├── pages/            # ImportPage, DashboardPage, CohortPage, CustomersPage, DownloadPage
        ├── components/       # Wizard steps, charts, shared UI
        ├── hooks/            # useWizard, useDashboard
        ├── types/            # TypeScript type definitions
        └── utils/            # Formatting utilities
```

## Prerequisites

- Node.js 18+
- npm 9+

## Getting Started

```bash
cd frontend
npm install
npm run dev
```

The app will be available at `http://localhost:5173`.

## Usage Workflow

1. **Upload** — Drag and drop your raw Excel file (.xlsx or .xlsm)
2. **Input format** — Select raw list or cleaned table format; columns are auto-detected
3. **Frequency** — Confirm or override the detected data frequency (monthly/quarterly)
4. **Data type** — Choose ARR or Revenue
5. **Granularity** — Select output granularities (monthly, quarterly, and/or annual)
6. **Identifiers** — Select and rename customer attributes for filtering
7. **Review & generate** — Confirm settings and generate the workbook
8. **Explore dashboard** — Filter by attributes and explore retention, cohort, and customer metrics
9. **Download** — Export the fully-formatted Excel workbook

## Input Data Format

Data Cutter supports two input formats:

**Raw list** — One row per customer per period:

| Column | Description |
|--------|-------------|
| Date | Period end date (monthly, quarterly, or annual) |
| Customer ID | Unique customer identifier |
| ARR / Revenue | Numeric value (raw, in thousands, or in millions) |
| Attributes (optional) | Categorical fields for filtering (e.g., Region, Segment, Product) |

**Cleaned table** — Customers as rows, date periods as columns:

| Customer ID | Attributes... | Q1 2023 | Q2 2023 | Q3 2023 | ... |
|-------------|---------------|---------|---------|---------|-----|
| Acme Corp   | Enterprise    | 120,000 | 125,000 | 130,000 | ... |

Data Cutter auto-detects the format, column roles, scale factor (1, 1K, or 1M), and data frequency — all of which can be adjusted in the wizard.

## Output Workbook Tabs

| Tab | Contents |
|-----|----------|
| Control | Configuration reference |
| Raw Data | Copy of the uploaded data |
| Clean Data | ARR by customer and period at each granularity |
| Retention | Lost-only, Punitive, and Net Dollar Retention metrics |
| Cohort | Cohort heatmap — logo counts and NDR by acquisition cohort |
| Top Customers | Customer rankings and concentration metrics |

Tabs are generated for each selected granularity (e.g., separate monthly and quarterly retention tabs).

## Tech Stack

| Layer | Technology |
|-------|------------|
| Framework | React 19 + TypeScript |
| Build tool | Vite |
| Excel I/O | ExcelJS |
| Charts | Recharts |
| Styling | Tailwind CSS |
| Routing | React Router |
