---
name: testing-data-cutter
description: How to run the data_cutter frontend and drive its import wizard, granularity toggle, grid axis selectors, and attribute filters for end-to-end UI testing.
---

# Testing data_cutter end-to-end

Frontend-only Vite + React app in `frontend/` — no backend; all data processing happens client-side (ExcelJS + a TS engine in `frontend/src/engine/`).

## Dev server

- Node is not on PATH by default: `export PATH=/home/ubuntu/.nvm/versions/node/v24.19.0/bin:$PATH`
- `cd frontend && npm install` (usually already done — `node_modules/` ships in the repo), then `npm run dev` → http://localhost:5173
- Vite HMR picks up engine edits live; an already-open page reloads automatically.

## Loading data

There is no account or login. On the Import tab (`/import`):
1. Click **"Try with sample data"** (loads `frontend/public/sample-data.xlsx` — 500-customer quarterly ARR cube, columns Dec'21–Sep'25, attributes Geography + Customer Type). The wizard jumps straight to step 7 (Review & Generate).
2. Click **"Generate Data Pack"**, wait for "Generation Complete!", then click **"View Dashboard"**.
3. Navigate via the sidebar (Histograms = `/histograms`). Analysis nav links redirect to `/import` until a session is generated.

## Histograms page controls

- Granularity toggle (Annual/Quarterly) = buttons top-right of the page header — NOT sticky; scroll to top first.
- "Mekko Axes" and "Grid Axes" sections each have native `<select>` dropdowns for X-Axis/Y-Axis (options: Cohort, Geography, Customer Type). Default Y auto-picks the first identifier that isn't X — setting X=Geography yields Y=Cohort.
- The Cohort/Geography/Customer Type filter bar is sticky at the top of the scroll area. The Cohort dropdown is multi-select with "Select All | Deselect All" links.
- Empty grids render "Not enough data" (TwoByTwoGrid early-returns when xLabels or yLabels is empty).

## Expected values for sanity checks (sample data, unfiltered)

- Annual: YoY grid cohort columns FY'21–FY'23 + Total (FY'24 is a new-logo cohort with no prior period and is intentionally excluded); subtitle "FY'23 to FY'24"; Africa row total ≈11.2% growth / 111.2% net retention; grand totals ≈5.7% growth / 105.7% net retention / 97.6% lost-only (net retention matches the Dashboard stat).
- Quarterly: grid cohort columns Q4'23–Q3'24 + Total; subtitle "Q3'24 to Q3'25". The Cohort filter lists the union of grid-eligible and newest cohorts (8 options: Q4'23–Q3'25).
- Selecting only a post-prior-period cohort (e.g. Q3'25) intentionally yields "Not enough data" grids; Mekko/pie charts still render that cohort.

## Console checks

The browser console tool only returns logs produced by the script you pass — install a hook early and read it later:
`window.__pageErrors=[]; const e=console.error; console.error=(...a)=>{__pageErrors.push(a.join(' ')); return e(...a)}` then `__pageErrors` at the end.
