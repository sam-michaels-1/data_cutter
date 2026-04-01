/**
 * Main orchestrator for the Data Pack Excel generator.
 * Port of data-pack-app/engine/generator.py for client-side use with ExcelJS.
 */
import ExcelJS from 'exceljs';
import type { EngineConfig, FilterBlock, CleanTabResult } from './types';
import { getYoyOffset, computeCleanLayout, colLetter } from './utils';
import { generateBaseCleanData, generateAggregatedCleanData } from './clean_data';
import { generateRetentionTab } from './retention';
import { generateCohortTab } from './cohort';
import { generateTopCustomersTab, TOP_N } from './top_customers';
import {
  formatControlTab, formatCleanDataTab, formatRetentionTab,
  formatCohortTab, formatTopCustomersTab, applyFormulaColoring
} from './formatting';

function colNumFromLetter(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.toUpperCase().charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Generate the complete data pack Excel file.
 * Returns an ExcelJS Workbook that can be saved as a buffer.
 */
export async function generateDataPack(
  config: EngineConfig,
  srcWb: ExcelJS.Workbook,
  onProgress?: (msg: string) => void
): Promise<ExcelJS.Workbook> {
  const log = onProgress || (() => {});

  log('Reading raw data...');
  const srcWs = srcWb.getWorksheet(config.raw_data_sheet);
  if (!srcWs) throw new Error(`Sheet "${config.raw_data_sheet}" not found`);

  const { uniqueDates, uniqueCustomers } = extractUniques(srcWs, config);
  log(`Found ${uniqueCustomers.length} unique customers`);
  log(`Found ${uniqueDates.length} unique time periods`);

  // Create output workbook
  const wb = new ExcelJS.Workbook();

  // --- Tab 1: Control ---
  log('Creating Control tab...');
  createControlTab(wb, config);

  // --- Copy raw data ---
  log('Copying raw data...');
  copyRawData(wb, srcWb, config);

  // --- Determine tabs ---
  const granularity = config.time_granularity;
  const outputGrans = config.output_granularities;
  const cleanTabs: Record<string, CleanTabResult> = {};

  // --- Base clean data tab ---
  log(`Generating Clean ${capitalize(granularity)} Data...`);
  const baseResult = generateBaseCleanData(wb, config, uniqueDates, uniqueCustomers);
  cleanTabs[granularity] = baseResult;

  // --- Aggregated tabs ---
  if (granularity === 'monthly') {
    const quarterlyDates = computeQuarterlyDates(uniqueDates, config);
    log(`Generating Clean Quarterly Data (${quarterlyDates.length} quarters)...`);
    const qResult = generateAggregatedCleanData(
      wb, config, baseResult.sheetName, baseResult.layout,
      'quarterly', uniqueCustomers, quarterlyDates);
    cleanTabs['quarterly'] = qResult;

    const annualDates = computeAnnualDates(uniqueDates, config);
    log(`Generating Clean Annual Data (${annualDates.length} years)...`);
    const aResult = generateAggregatedCleanData(
      wb, config, baseResult.sheetName, baseResult.layout,
      'annual', uniqueCustomers, annualDates);
    cleanTabs['annual'] = aResult;

  } else if (granularity === 'quarterly') {
    const annualDates = computeAnnualFromQuarterlyDates(uniqueDates, config);
    log(`Generating Clean Annual Data (${annualDates.length} years)...`);
    const aResult = generateAggregatedCleanData(
      wb, config, baseResult.sheetName, baseResult.layout,
      'annual', uniqueCustomers, annualDates);
    cleanTabs['annual'] = aResult;
  }

  // --- Build filter blocks ---
  const filterBlocks = buildFilterBlocks(config);
  const numAttrs = Object.keys(config.attributes).length;

  // --- Retention tabs ---
  for (const g of getAvailableGranularities(granularity, outputGrans)) {
    if (g in cleanTabs) {
      const { sheetName, layout, firstDataRow, lastDataRow } = cleanTabs[g];
      log(`Generating ${capitalize(g)} Retention...`);
      generateRetentionTab(wb, config, sheetName, layout, firstDataRow, lastDataRow, g, filterBlocks);

      // Format retention tab
      const yoyOffset = getYoyOffset(g);
      const numDerived = layout.num_dates - yoyOffset;
      const filterStartCol = 2;
      const cohortFc = filterStartCol + numAttrs;
      const s1Label = cohortFc + 2;
      const s1Start = s1Label + 1;
      const s1End = s1Start + numDerived - 1;
      const s2Label = s1End + 2;
      const s2Start = s2Label + 1;
      const s2End = s2Start + numDerived - 1;
      const s3Label = s2End + 2;
      const s3Start = s3Label + 1;
      const s3End = s3Start + numDerived - 1;

      const retSheet = `${capitalize(g)} Retention`;
      log(`Formatting ${retSheet}...`);
      const retWs = wb.getWorksheet(retSheet);
      if (retWs) {
        formatRetentionTab(
          retWs, config, filterBlocks, numDerived, numAttrs,
          s1Label, s1Start, s1End, s2Label, s2Start, s2End,
          s3Label, s3Start, s3End, filterStartCol, cohortFc);
      }
    }
  }

  // --- Cohort tabs ---
  for (const g of ['quarterly', 'annual'] as const) {
    if (g in cleanTabs) {
      const { sheetName, layout, firstDataRow, lastDataRow } = cleanTabs[g];
      log(`Generating ${capitalize(g)} Cohort...`);
      generateCohortTab(wb, config, sheetName, layout, firstDataRow, lastDataRow, g, filterBlocks);

      // Format cohort tab
      const numDates = layout.num_dates;
      const numCohorts = numDates;
      const qCol = 2;
      const yColC = 3;
      const filterStart = 4;
      const filterEnd = filterStart + numAttrs - 1;
      const cohortLabelCol = filterEnd + 1;
      const s1Start = cohortLabelCol + 1;
      const s1End = s1Start + numDates - 1;
      const s2Start = s1End + 2;
      const s2End = s2Start + numDates - 1;
      const s3LabelCol = s2End + 2;
      const s3StartValCol = s3LabelCol + 1;
      const s3DataStart = s3StartValCol + 1;
      const s3DataEnd = s3DataStart + numDates - 1;
      const s4LabelCol = s3DataEnd + 2;
      const s4StartValCol = s4LabelCol + 1;
      const s4DataStart = s4StartValCol + 1;
      const s4DataEnd = s4DataStart + numDates - 1;

      const cohSheet = `${capitalize(g)} Cohort`;
      log(`Formatting ${cohSheet}...`);
      const cohWs = wb.getWorksheet(cohSheet);
      if (cohWs) {
        formatCohortTab(
          cohWs, config, filterBlocks,
          numDates, numCohorts, numAttrs,
          qCol, yColC, filterStart, cohortLabelCol,
          s1Start, s1End, s2Start, s2End,
          s3LabelCol, s3StartValCol, s3DataStart, s3DataEnd,
          s4LabelCol, s4StartValCol, s4DataStart, s4DataEnd,
          g);
      }
    }
  }

  // --- Top Customer Analysis ---
  if ('annual' in cleanTabs) {
    const { sheetName, layout, firstDataRow, lastDataRow } = cleanTabs['annual'];
    log('Generating Annual Top Customer Analysis...');
    generateTopCustomersTab(wb, config, sheetName, layout, firstDataRow, lastDataRow);

    // Format
    const numDates = layout.num_dates;
    const rankNumCol = 2;
    const custIdCol = 3;
    const attrStart = 4;
    const cohortCol = attrStart + numAttrs;
    const s1Start = cohortCol + 1;
    const s1End = s1Start + numDates - 1;
    const s2Start = s1End + 2;
    const s2End = s2Start + numDates - 2;
    const s3Start = s2End + 2;
    const s3End = s3Start + numDates - 1;
    const firstCustomerRow = 7;
    const lastCustomerRow = firstCustomerRow + TOP_N - 1;
    const rTopTotal = lastCustomerRow + 1;
    const rOther = rTopTotal + 1;
    const rTotal = rOther + 1;
    const rMemoStart = rTotal + 2;

    log('Formatting Annual Top Customer Analysis...');
    const topWs = wb.getWorksheet('Annual Top Customer Analysis');
    if (topWs) {
      formatTopCustomersTab(
        topWs, config, layout,
        firstCustomerRow, lastCustomerRow,
        rTopTotal, rOther, rTotal, rMemoStart,
        numDates, rankNumCol, custIdCol, attrStart, numAttrs, cohortCol,
        s1Start, s1End, s2Start, s2End, s3Start, s3End);
    }
  }

  // --- Control tab check summary ---
  log('Adding control tab check summary...');
  const checkTabs = addControlChecks(wb, granularity, outputGrans);

  // --- Format Control tab ---
  log('Formatting Control tab...');
  const controlWs = wb.getWorksheet('Control');
  if (controlWs) formatControlTab(controlWs, checkTabs);

  // --- Format Clean Data tabs ---
  for (const g of getAvailableGranularities(granularity, outputGrans)) {
    if (g in cleanTabs) {
      const { sheetName: sn, layout: ly, firstDataRow: fd, lastDataRow: ld } = cleanTabs[g];
      log(`Formatting ${sn}...`);
      const cleanWs = wb.getWorksheet(sn);
      if (cleanWs) formatCleanDataTab(cleanWs, ly, fd, ld, g);
    }
  }

  // --- Reorder sheets ---
  reorderSheets(wb, granularity);

  // --- Apply formula coloring ---
  log('Applying formula color-coding...');
  applyFormulaColoring(wb);

  log('Done!');
  return wb;
}


// --- Helper functions ---

function capitalize(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function extractUniques(ws: ExcelJS.Worksheet, config: EngineConfig): { uniqueDates: Date[]; uniqueCustomers: string[] } {
  const dateColIdx = colNumFromLetter(config.date_col);
  const custColIdx = colNumFromLetter(config.customer_id_col);
  const firstRow = config.raw_data_first_row;

  const dates = new Set<number>();
  const dateMap = new Map<number, Date>();
  const customers = new Set<string>();

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < firstRow) return;
    const dateVal = row.getCell(dateColIdx).value;
    const custVal = row.getCell(custColIdx).value;
    if (dateVal != null) {
      if (dateVal instanceof Date) {
        const t = dateVal.getTime();
        dates.add(t);
        dateMap.set(t, dateVal);
      }
    }
    if (custVal != null) {
      customers.add(String(custVal));
    }
  });

  const sortedDateTimes = [...dates].sort((a, b) => a - b);
  const uniqueDates = sortedDateTimes.map(t => dateMap.get(t)!);
  const uniqueCustomers = [...customers].sort();

  return { uniqueDates, uniqueCustomers };
}

function createControlTab(wb: ExcelJS.Workbook, config: EngineConfig): void {
  const ws = wb.addWorksheet('Control');

  ws.getCell(3, 2).value = 'Raw Data Type:';
  ws.getCell(3, 3).value = (config.data_type || 'arr') === 'arr' ? 'ARR' : 'Revenue';

  ws.getCell(4, 2).value = 'Scale Factor (Divide By):';
  ws.getCell(4, 3).value = config.scale_factor;

  ws.getCell(5, 2).value = 'Raw Data Timing Interval:';
  ws.getCell(5, 3).value = capitalize(config.time_granularity);

  const monthNames: Record<number, string> = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April',
    5: 'May', 6: 'June', 7: 'July', 8: 'August',
    9: 'September', 10: 'October', 11: 'November', 12: 'December'
  };
  ws.getCell(6, 2).value = 'Fiscal Year End:';
  ws.getCell(6, 3).value = monthNames[config.fiscal_year_end_month];

  ws.getCell(7, 2).value = 'Fiscal Year Month #:';
  ws.getCell(7, 3).value = config.fiscal_year_end_month;
}

function copyRawData(wb: ExcelJS.Workbook, srcWb: ExcelJS.Workbook, config: EngineConfig): void {
  const srcWs = srcWb.getWorksheet(config.raw_data_sheet);
  if (!srcWs) return;

  // Separator tab
  const wsSep = wb.addWorksheet('Raw Data>>');
  wsSep.getCell(1, 1).value = 'Raw Data >>';

  // Copy data
  const ws = wb.addWorksheet(config.raw_data_sheet);
  srcWs.eachRow({ includeEmpty: false }, (srcRow, rowNumber) => {
    srcRow.eachCell({ includeEmpty: false }, (srcCell, colNumber) => {
      ws.getCell(rowNumber, colNumber).value = srcCell.value;
    });
  });
}

function computeQuarterlyDates(monthlyDates: Date[], config: EngineConfig): Date[] {
  const fyMonth = config.fiscal_year_end_month;
  return monthlyDates.filter(d => {
    const month = d.getMonth() + 1;
    return (fyMonth - month + 12) % 3 === 0;
  });
}

function computeAnnualDates(monthlyDates: Date[], config: EngineConfig): Date[] {
  const fyMonth = config.fiscal_year_end_month;
  return monthlyDates.filter(d => (d.getMonth() + 1) === fyMonth);
}

function computeAnnualFromQuarterlyDates(quarterlyDates: Date[], config: EngineConfig): Date[] {
  const fyMonth = config.fiscal_year_end_month;
  return quarterlyDates.filter(d => (d.getMonth() + 1) === fyMonth);
}

function buildFilterBlocks(config: EngineConfig): FilterBlock[] {
  const attrNames = Object.keys(config.attributes);

  // Block 1: Total Business
  const totalFilters: Record<string, string> = {};
  for (const name of attrNames) totalFilters[name] = '<>';
  totalFilters['Cohort'] = '<>';

  const blocks: FilterBlock[] = [{ title: 'Total Business', filters: totalFilters }];

  for (const breakout of config.filter_breakouts || []) {
    const blockFilters: Record<string, string> = {};
    for (const name of attrNames) blockFilters[name] = '<>';
    blockFilters['Cohort'] = '<>';
    Object.assign(blockFilters, breakout.filters || {});
    blocks.push({ title: breakout.title || 'Filtered', filters: blockFilters });
  }

  return blocks;
}

function getAvailableGranularities(baseGranularity: string, outputGranularities?: string[]): string[] {
  const allFromBase: Record<string, string[]> = {
    monthly: ['monthly', 'quarterly', 'annual'],
    quarterly: ['quarterly', 'annual'],
    annual: ['annual'],
  };

  const available = allFromBase[baseGranularity] || [baseGranularity];
  if (outputGranularities) {
    return available.filter(g => outputGranularities.includes(g));
  }
  return available;
}

function addControlChecks(wb: ExcelJS.Workbook, granularity: string, outputGrans?: string[]): [string, string][] {
  const ws = wb.getWorksheet('Control');
  if (!ws) return [];

  const checkTabs: [string, string][] = [];
  for (const g of getAvailableGranularities(granularity, outputGrans)) {
    const retName = `${capitalize(g)} Retention`;
    if (wb.getWorksheet(retName)) {
      checkTabs.push([`${capitalize(g)} Retention Check`, retName]);
    }
  }
  for (const g of ['annual', 'quarterly'] as const) {
    const cohName = `${capitalize(g)} Cohort`;
    if (wb.getWorksheet(cohName)) {
      checkTabs.push([`${capitalize(g)} Cohort Check`, cohName]);
    }
  }

  const R_HDR = 10;
  ws.getCell(R_HDR, 2).value = 'Check Summary';
  ws.getCell(R_HDR, 3).value = 'Value (= 0)';

  const checkCellRefs: string[] = [];
  for (let i = 0; i < checkTabs.length; i++) {
    const [label, tabName] = checkTabs[i];
    const r = R_HDR + 1 + i;
    ws.getCell(r, 2).value = label;
    ws.getCell(r, 3).value = { formula: `'${tabName}'!B1` };
    checkCellRefs.push(`C${r}`);
  }

  const rTotal = R_HDR + 1 + checkTabs.length;
  ws.getCell(rTotal, 2).value = 'Total Check (= 0)';
  if (checkCellRefs.length > 0) {
    ws.getCell(rTotal, 3).value = { formula: `SUM(${checkCellRefs.join(',')})` };
  }

  return checkTabs;
}

function reorderSheets(wb: ExcelJS.Workbook, granularity: string): void {
  const desiredOrder = ['Control'];

  if (granularity === 'monthly') {
    desiredOrder.push(
      'Annual Top Customer Analysis',
      'Annual Cohort', 'Quarterly Cohort',
      'Annual Retention', 'Quarterly Retention', 'Monthly Retention',
      'Clean Annual Data', 'Clean Quarterly Data', 'Clean Monthly Data',
    );
  } else if (granularity === 'quarterly') {
    desiredOrder.push(
      'Annual Top Customer Analysis',
      'Annual Cohort', 'Quarterly Cohort',
      'Annual Retention', 'Quarterly Retention',
      'Clean Annual Data', 'Clean Quarterly Data',
    );
  } else {
    desiredOrder.push(
      'Annual Top Customer Analysis',
      'Annual Cohort',
      'Annual Retention',
      'Clean Annual Data',
    );
  }

  desiredOrder.push('Raw Data>>');
  // Add the raw data sheet name
  const sheetNames = wb.worksheets.map(ws => ws.name);
  const rawSheetName = sheetNames[sheetNames.length - 1];
  if (rawSheetName && !desiredOrder.includes(rawSheetName)) {
    desiredOrder.push(rawSheetName);
  }

  // Build final order
  const finalOrder: string[] = [];
  for (const name of desiredOrder) {
    if (wb.getWorksheet(name) && !finalOrder.includes(name)) {
      finalOrder.push(name);
    }
  }
  for (const name of sheetNames) {
    if (!finalOrder.includes(name)) {
      finalOrder.push(name);
    }
  }

  // Reorder by moving sheets to their correct positions
  // ExcelJS doesn't have a direct reorder API, so we use orderNo
  for (let i = 0; i < finalOrder.length; i++) {
    const ws = wb.getWorksheet(finalOrder[i]);
    if (ws) {
      ws.orderNo = i;
    }
  }
}
