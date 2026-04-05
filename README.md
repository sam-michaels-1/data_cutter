# Data Cutter

A client-side web application that converts raw ARR/Revenue Excel data into a professional multi-tab analysis workbook, complete with retention metrics, cohort analysis, distribution histograms, and interactive dashboards.

## Overview

Upload a raw Excel file containing customer-level ARR or revenue data, map your columns through a guided wizard, and Data Cutter will generate a formula-driven Excel workbook with retention waterfall analysis, cohort tracking, and top customer breakdowns — all at your chosen time granularity. A live dashboard with six analysis views lets you explore the data interactively before downloading.

Everything runs in the browser — no backend or server required.

## Features

### Import Wizard
- **7-step guided flow** — Upload, detect format, configure frequency, data type, granularity, identifiers, and review
- **Auto-detection** — Infers date columns, customer IDs, ARR values, scale factors (1/1K/1M), and data frequency
- **Dual input formats** — Supports both raw list (one row per customer per period) and cleaned table (customers as rows, periods as columns)
- **Multi-granularity output** — Generate monthly, quarterly, and/or annual views from a single upload
- **Fiscal year support** — Configure fiscal year end month for proper period alignment

### Dashboard
- **Stats cards** — Total ARR/Revenue, customer count, YoY growth, and three retention metrics (lost-only, punitive, NDR)
- **ARR trend chart** — Bar chart of ARR over time with YoY growth percentages
- **Waterfall chart** — Revenue movement breakdown: BOP, new logo, upsell, downsell, churn, EOP
- **Top customers table** — Ranked by ARR with change %, concentration %, status badges, and sparkline trends
- **Attribute filtering** — Filter all views by any combination of customer attributes

### Histograms & Distributions
- **ARR histogram** — Distribution of ARR across customer brackets
- **Growth histogram** — Distribution of YoY growth rates
- **Mekko charts** — Multi-dimensional stacked column charts with customizable X/Y axes for ARR and customer count
- **Identifier pie charts** — Breakdown by each customer attribute
- **2x2 growth grids** — Segment matrix of YoY growth by selected dimensions
- **Retention grids** — Net retention and loss-only retention by segment

### Cohort Analysis
- **Cohort heatmap** — Color-coded grid of cohort performance over time
- **Four metrics** — ARR, Net Dollar Retention, Logo Retention, Customer Count
- **Granularity switching** — View cohorts at monthly, quarterly, or annual level
- **Attribute filtering** — Slice cohorts by any customer dimension

### Top Customers
- **Ranked table** — Customer name, attributes, cohort, current ARR, change %, and % of total
- **Status tracking** — Growth, Stable, Declining, and New badges
- **Sparkline trends** — Inline trend charts showing historical ARR
- **Configurable top-N** — View top 10, 15, 25, or 50 customers

### Excel Export
- **Formula-driven workbook** — Output uses Excel formulas, not static values
- **Multiple tabs** — Control, Raw Data, Clean Data, Retention, Cohort, and Top Customers tabs per granularity

## Project Structure

```
Data Cutter/
└── frontend/                     # React + TypeScript + Vite
    └── src/
        ├── App.tsx               # Router (6 pages)
        ├── api/                  # Client-side data fetching
        │   ├── client.ts         # File upload, column detection, workbook generation
        │   ├── dashboard.ts      # Dashboard metric computation
        │   └── histograms.ts     # Histogram/distribution computation
        ├── engine/               # Excel generation engine
        │   ├── generator.ts      # Main orchestrator — builds the output workbook
        │   ├── clean_data.ts     # Raw and aggregated data tabs
        │   ├── retention.ts      # Retention metric tabs (lost-only, punitive, NDR)
        │   ├── cohort.ts         # Cohort analysis tabs
        │   ├── top_customers.ts  # Top customers tab
        │   ├── histograms.ts     # Distribution and segment calculations
        │   ├── compute.ts        # Data transformation (aggregation, derived metrics)
        │   ├── detect.ts         # Column role and format auto-detection
        │   ├── config_builder.ts # Wizard config → engine config
        │   ├── formatting.ts     # Cell styling and Excel formatting
        │   ├── utils.ts          # Formula builders, column helpers, layout utilities
        │   └── types.ts          # Shared type definitions
        ├── pages/
        │   ├── ImportPage.tsx     # 7-step import wizard
        │   ├── DashboardPage.tsx  # Overview stats, charts, and top customers
        │   ├── HistogramsPage.tsx # Distributions, Mekko charts, grids, pie charts
        │   ├── CohortPage.tsx    # Cohort heatmap analysis
        │   ├── CustomersPage.tsx # Top customers ranking table
        │   └── DownloadPage.tsx  # Excel workbook export
        ├── components/
        │   ├── steps/            # Wizard steps (Upload, InputFormat, Frequency, DataType, Granularity, Identifiers, Review)
        │   ├── dashboard/        # StatsCards, ARRBarChart, WaterfallChart, TopCustomersTable, CohortHeatmap
        │   ├── histograms/       # ARRHistogram, GrowthHistogram, MekkoChart, IdentifierPieCharts, TwoByTwoGrid, RetentionGrids
        │   ├── Sidebar.tsx       # Navigation (Data → Analysis → Export)
        │   ├── AppShell.tsx      # Layout wrapper
        │   ├── FileUpload.tsx    # Drag-and-drop file upload
        │   ├── ColumnMapper.tsx  # Column mapping UI
        │   ├── StepIndicator.tsx # Wizard progress indicator
        │   ├── MultiSelectDropdown.tsx  # Multi-select attribute filter
        │   ├── AttributeFilterBar.tsx   # Filter bar for analysis pages
        │   └── SessionProvider.tsx      # Session context
        ├── hooks/
        │   ├── useWizard.ts      # Wizard state management (reducer-based)
        │   ├── useDashboard.ts   # Dashboard data loading and filtering
        │   └── useHistogramData.ts # Histogram data loading with axis selection
        ├── types/                # TypeScript type definitions
        └── utils/                # Formatting helpers (currency, percentages, axes)
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
5. **Granularity** — Select output granularities (monthly, quarterly, and/or annual) and fiscal year end
6. **Identifiers** — Select and rename customer attributes for filtering and segmentation
7. **Review & generate** — Confirm settings and generate the workbook
8. **Explore** — Navigate the dashboard, histograms, cohort, and customers pages to analyze your data
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
