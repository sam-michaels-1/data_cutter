/**
 * Translate wizard selections into the engine's config dictionary.
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig } from './types';
import { colNum } from './utils';

/**
 * Auto-detect the raw data frequency by examining date intervals.
 */
function detectRawDataFrequency(wb: Workbook, sheetName: string, dateCol: string, headerRow = 1): string {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) return 'annual';

  const colIdx = colNum(dateCol);
  const dates: Date[] = [];

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber <= headerRow) return;
    const val = row.getCell(colIdx).value;
    if (val instanceof Date) {
      dates.push(val);
    }
  });

  if (dates.length < 2) return 'annual';

  const unique = [...new Set(dates.map(d => d.getTime()))].sort().map(t => new Date(t));
  if (unique.length < 2) return 'annual';

  const gaps: number[] = [];
  const limit = Math.min(10, unique.length - 1);
  for (let i = 0; i < limit; i++) {
    gaps.push((unique[i + 1].getTime() - unique[i].getTime()) / (1000 * 60 * 60 * 24));
  }
  const avgGap = gaps.reduce((a, b) => a + b, 0) / gaps.length;

  if (avgGap < 45) return 'monthly';
  if (avgGap < 120) return 'quarterly';
  return 'annual';
}

/**
 * Read unique values from a column (for auto-generating filter breakouts).
 */
function getUniqueValues(wb: Workbook, sheetName: string, colLetterStr: string, maxUnique = 50, headerRow = 1): string[] {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) return [];

  const colIdx = colNum(colLetterStr);
  const values = new Set<string>();

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber <= headerRow) return;
    if (values.size >= maxUnique) return;
    const val = row.getCell(colIdx).value;
    if (val == null) return;
    const valStr = String(val).trim();
    if (valStr) values.add(valStr);
  });

  return [...values].sort();
}

export interface BuildConfigParams {
  wb: Workbook;
  sheetName: string;
  dataType: string;
  dateCol: string;
  customerIdCol: string;
  arrCol: string;
  attributes: { display_name: string; letter: string }[];
  outputGranularities: string[];
  fiscalYearEndMonth: number;
  rowCount: number;
  scaleFactor: number;
  dataFrequency?: string;
  // Cleaned table fields
  inputFormat?: 'raw' | 'cleaned';
  dateColumns?: string[];
  headerRow?: number;
  dateHeaderRow?: number;
}

/**
 * Build the config dict the engine expects.
 */
export function buildEngineConfig(params: BuildConfigParams): EngineConfig {
  const {
    wb, sheetName, dataType, dateCol, customerIdCol, arrCol,
    attributes, outputGranularities, fiscalYearEndMonth,
    rowCount, scaleFactor, dataFrequency,
    inputFormat, dateColumns, headerRow, dateHeaderRow,
  } = params;

  // Build ordered attributes dict
  const attrs: Record<string, string> = {};
  for (const attr of attributes) {
    attrs[attr.display_name] = attr.letter;
  }

  // Use user-provided frequency override if available, otherwise auto-detect
  const rawFreq = dataFrequency || detectRawDataFrequency(wb, sheetName, dateCol, headerRow);

  // Auto-generate filter breakouts from the first attribute
  const filterBreakouts: { title: string; filters: Record<string, string> }[] = [];
  if (attributes.length > 0) {
    const firstAttr = attributes[0];
    const uniqueVals = getUniqueValues(wb, sheetName, firstAttr.letter, 50, headerRow);
    for (const val of uniqueVals) {
      filterBreakouts.push({
        title: val,
        filters: { [firstAttr.display_name]: val },
      });
    }
  }

  return {
    raw_data_sheet: sheetName,
    raw_data_first_row: (headerRow || 1) + 1,  // data starts one row after headers
    raw_data_last_row: rowCount + (headerRow || 1),  // rowCount data rows after header
    customer_id_col: customerIdCol,
    date_col: dateCol,
    arr_col: arrCol,
    attributes: attrs,
    time_granularity: rawFreq,
    output_granularities: outputGranularities,
    fiscal_year_end_month: fiscalYearEndMonth,
    scale_factor: scaleFactor,
    filter_breakouts: filterBreakouts,
    data_type: dataType,
    input_format: inputFormat || 'raw',
    date_columns: dateColumns,
    date_header_row: dateHeaderRow,
  };
}
