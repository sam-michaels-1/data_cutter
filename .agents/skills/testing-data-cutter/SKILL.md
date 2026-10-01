---
name: testing-data-cutter
description: How to run the data_cutter frontend and drive its import wizard, granularity toggle, grid axis selectors, attribute filters, and generated-workbook verification for end-to-end UI testing.
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
2. Click **"Generate Data Pack"**, wait for "Generation Complete!" (~15s), then click **"Download Excel"** — saves `data-pack-output.xlsx` to `~/Downloads` (~1.1 MB).
3. "View Dashboard" loads the analysis UI. Analysis nav links redirect to `/import` until a session is generated.

## Histograms page controls

- Granularity toggle (Annual/Quarterly) = buttons top-right of the page header — NOT sticky; scroll to top first.
- "Mekko Axes" and "Grid Axes" sections each have native `<select>` dropdowns for X-Axis/Y-Axis (options: Cohort, Geography, Customer Type). Default Y auto-picks the first identifier that isn't X — setting X=Geography yields Y=Cohort.
- The Cohort/Geography/Customer Type filter bar is sticky at the top of the scroll area. The Cohort dropdown is multi-select with "Select All | Deselect All" links.
- Empty grids render "Not enough data" (TwoByTwoGrid early-returns when xLabels or yLabels is empty).

## Verifying the generated workbook (no Excel/LibreOffice on the box)

No spreadsheet GUI or openpyxl is installed. To inspect the downloaded xlsx:

- Use `verify_workbook.mjs` in this skill's directory (needs `npm i exceljs hyperformula` in a scratch dir; npm registry works). It checks structure, labels, fonts, and recalculates the whole book — run `node verify_workbook.mjs` after editing the FILE path.
- exceljs reads structure/labels/formulas; cell values are NOT cached (ExcelJS writes formulas only), so numeric results must be recalculated. The script recalcs via HyperFormula `buildFromSheets` (cells → 2D arrays, `{formula}` → `'='+f`, Dates → Excel serials `(d-Date.UTC(1899,11,30))/86400000`), after patching four formula shapes HF can't evaluate:
  1. `MATCH(TRUE, INDEX(range<>0,0),0)` cohort formulas — HF doesn't know bare `TRUE` and can't INDEX a computed array. Wrapped in IFERROR they silently become `"n.a."`, which makes SUMIFS checks trivially pass for the wrong reason. Fix: precompute the cohort literal (header of the first nonzero ARR column per row) and inject it.
  2. `RANK(...)` in the Customer Rank columns — not a registered HF function. Inject computed literals (`1 + count(col > x)`), or skip if you only need retention checks (nothing references rank).
  3. `_xlfn.`/`_xlws.` prefixes (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`, `_xlfn.UNIQUE`) — strip the prefixes; HF supports the underlying functions.
  4. `INDEX('S'!$C$7:$E$506,0,MATCH($B$5,'S'!$C$6:$E$6,0))` column-pick criteria in Summary tabs — rewrite to the resolved column range `'S'!$X$7:$X$506` (the match key is a literal like "Geography").
- `Clean Annual Data` row-3 year cell also uses the MATCH(TRUE) pattern — inject the year literal (first year whose CQD quarter==4) so annual SUMIFS evaluate.
- Also worth scanning: no cell's `value.formula` should start with '=' — ExcelJS writes `<f>` verbatim, so a leading '=' produces `==ref` cells that parse-error in Excel.

## Expected values for sanity checks (sample data, unfiltered)

- Annual: YoY grid cohort columns FY'21–FY'23 + Total (FY'24 is a new-logo cohort with no prior period and is intentionally excluded); subtitle "FY'23 to FY'24"; Africa row total ≈11.2% growth / 111.2% net retention; grand totals ≈5.7% growth / 105.7% net retention / 97.6% lost-only (net retention matches the Dashboard stat).
- Quarterly: grid cohort columns Q4'23–Q3'24 + Total; subtitle "Q3'24 to Q3'25". The Cohort filter lists the union of grid-eligible and newest cohorts (8 options: Q4'23–Q3'25).
- Selecting only a post-prior-period cohort (e.g. Q3'25) intentionally yields "Not enough data" grids; Mekko/pie charts still render that cohort.
- Generated pack (quarterly sample): 15 sheets — Control, Data Summary, Annual/Quarterly Summary, Annual Top Customer Analysis, Annual/Quarterly Cohort, Annual/Quarterly Retention, Quarterly Retention (QoQ), Clean Annual/Quarterly Data, Clean Quarterly Data (QoQ), Raw Data>>, Customer Cube. All "… Check" rows on Control should evaluate to 0. Retention block pitch is 19 rows on every retention tab; on the QoQ tab the three Section 1 retention rows ARE the annualized variants ("% Annualized Lost-Only/Punitive/Net Retention"; net row bold + indented label) — no plain % Lost-Only/Punitive/Net rows appear there. BoP rows link directly to the prior EoP column once one exists (purple font 7030A0); the first yoyOffset columns keep green SUMIFS (ARR) / COUNTIFS (customers). Monthly input additionally produces 'Clean Monthly Data (MoM)' + 'Monthly Retention (MoM)' tabs (×12 annualization) — absent for quarterly input.
- QoQ annualized spot values (All block): Q3'22 → 100.0% / 96.8% / 111.8%; Q1'22 AnnNet = (BoP + upsell×4)/BoP = 114.6%.

## Console checks

The browser console tool only returns logs produced by the script you pass — install a hook early and read it later:
`window.__pageErrors=[]; const e=console.error; console.error=(...a)=>{__pageErrors.push(a.join(' ')); return e(...a)}` then `__pageErrors` at the end.
