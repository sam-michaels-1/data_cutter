/**
 * Auto-detection service.
 * Reads sheet data and infers column types, scale factor, etc.
 * Port of backend/services/detect.py for client-side use.
 */
import type { Workbook } from 'exceljs';

const SAMPLE_ROWS = 50;

export interface ColumnInfo {
  letter: string;
  header: string;
  sample_values: string[];
}

export interface DetectedMapping {
  date_col: string;
  customer_id_col: string;
  arr_col: string;
  attribute_cols: { header: string; letter: string }[];
}

export interface DetectResult {
  columns: ColumnInfo[];
  detected_mapping: DetectedMapping;
  row_count: number;
  auto_scale_factor: number;
  detected_frequency: string;
}

function colLetterFromNum(n: number): string {
  let result = '';
  let num = n;
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

function colNumFromLetter(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.toUpperCase().charCodeAt(i) - 64);
  }
  return result;
}

function isDate(v: unknown): boolean {
  return v instanceof Date && !isNaN(v.getTime());
}

function isNumeric(v: unknown): boolean {
  if (typeof v === 'boolean') return false;
  if (typeof v === 'number') return true;
  if (typeof v === 'string') {
    const cleaned = v.replace(/,/g, '');
    return cleaned !== '' && !isNaN(Number(cleaned));
  }
  return false;
}

function toNumber(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') return parseFloat(v.replace(/,/g, ''));
  return 0;
}

/**
 * Detect the raw data frequency by examining date intervals.
 */
function detectFrequency(dates: Date[]): string {
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
 * Detect columns from a loaded ExcelJS workbook sheet.
 */
export function detectColumns(wb: Workbook, sheetName: string): DetectResult {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) throw new Error(`Sheet "${sheetName}" not found`);

  // Read headers (row 1)
  const headerRow = ws.getRow(1);
  const headers: { col_num: number; letter: string; header: string }[] = [];
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (cell.value != null) {
      headers.push({
        col_num: colNumber,
        letter: colLetterFromNum(colNumber),
        header: String(cell.value).trim(),
      });
    }
  });

  if (headers.length === 0) {
    throw new Error('No headers found in row 1 of the selected sheet.');
  }

  // Collect sample values per column
  const colSamples: Record<number, unknown[]> = {};
  for (const h of headers) colSamples[h.col_num] = [];

  let totalDataRows = 0;
  const allDates: Date[] = [];

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // skip header
    const firstCellVal = row.getCell(1).value;
    if (firstCellVal == null) return;
    totalDataRows++;

    if (totalDataRows <= SAMPLE_ROWS) {
      for (const h of headers) {
        const val = row.getCell(h.col_num).value;
        colSamples[h.col_num].push(val instanceof Date ? val : (val as unknown));
      }
    }

    // Collect dates for frequency detection
    for (const h of headers) {
      const val = row.getCell(h.col_num).value;
      if (val instanceof Date) {
        allDates.push(val);
      }
    }
  });

  // Classify each column
  const colStats: Record<number, {
    date_ratio: number; num_ratio: number; str_ratio: number;
    unique_count: number; total_sampled: number; max_abs: number;
    type?: string;
  }> = {};

  for (const h of headers) {
    const vals = colSamples[h.col_num];
    const nonNone = vals.filter(v => v != null);
    if (nonNone.length === 0) {
      colStats[h.col_num] = { type: 'empty', date_ratio: 0, num_ratio: 0, str_ratio: 0, unique_count: 0, total_sampled: 0, max_abs: 0 };
      continue;
    }

    const dateCount = nonNone.filter(v => isDate(v)).length;
    const dateRatio = dateCount / nonNone.length;

    const numCount = nonNone.filter(v => isNumeric(v)).length;
    const numRatio = numCount / nonNone.length;

    const strCount = nonNone.filter(v => typeof v === 'string').length;
    const strRatio = strCount / nonNone.length;

    let uniqueCount: number;
    try {
      uniqueCount = new Set(nonNone.map(v => String(v))).size;
    } catch {
      uniqueCount = nonNone.length;
    }

    let maxAbs = 0;
    if (numRatio > 0.5) {
      const nums: number[] = [];
      for (const v of nonNone) {
        try {
          if (typeof v === 'number' && typeof v !== 'boolean') {
            nums.push(Math.abs(v));
          } else if (typeof v === 'string') {
            nums.push(Math.abs(parseFloat(v.replace(/,/g, ''))));
          }
        } catch { /* skip */ }
      }
      maxAbs = nums.length > 0 ? Math.max(...nums) : 0;
    }

    colStats[h.col_num] = { date_ratio: dateRatio, num_ratio: numRatio, str_ratio: strRatio, unique_count: uniqueCount, total_sampled: nonNone.length, max_abs: maxAbs };
  }

  // Detect date column: highest date_ratio > 0.8, leftmost wins
  let dateCol: string | null = null;
  for (const h of headers) {
    const s = colStats[h.col_num];
    if ((s.date_ratio || 0) > 0.8) {
      dateCol = h.letter;
      break;
    }
  }

  // Detect ARR column: numeric, highest max_abs
  let arrCol: string | null = null;
  let bestMaxAbs = -1;
  for (const h of headers) {
    const s = colStats[h.col_num];
    if (h.letter === dateCol) continue;
    if ((s.num_ratio || 0) > 0.5 && (s.max_abs || 0) > bestMaxAbs) {
      bestMaxAbs = s.max_abs || 0;
      arrCol = h.letter;
    }
  }

  // Detect customer ID column: text, high cardinality
  let customerIdCol: string | null = null;
  let bestCardinality = -1;
  for (const h of headers) {
    const s = colStats[h.col_num];
    if (h.letter === dateCol || h.letter === arrCol) continue;
    if ((s.str_ratio || 0) > 0.5) {
      const card = s.unique_count || 0;
      if (card > bestCardinality) {
        bestCardinality = card;
        customerIdCol = h.letter;
      }
    }
  }

  // Detect attribute columns: everything remaining with low cardinality
  const attributeCols: { header: string; letter: string }[] = [];
  for (const h of headers) {
    const s = colStats[h.col_num];
    if (h.letter === dateCol || h.letter === arrCol || h.letter === customerIdCol) continue;
    if (s.type === 'empty') continue;
    const unique = s.unique_count || 0;
    if (unique > 0 && unique < 100) {
      attributeCols.push({ header: h.header, letter: h.letter });
    }
  }

  // Auto scale factor
  let arrTotalEstimate = 0;
  if (arrCol) {
    const arrCn = headers.find(h => h.letter === arrCol)?.col_num;
    if (arrCn) {
      let sampleSum = 0;
      for (const v of colSamples[arrCn] || []) {
        try {
          if (typeof v === 'number' && typeof v !== 'boolean') {
            sampleSum += Math.abs(v);
          } else if (typeof v === 'string') {
            sampleSum += Math.abs(parseFloat(v.replace(/,/g, '')));
          }
        } catch { /* skip */ }
      }
      const sampledN = Math.min(SAMPLE_ROWS, totalDataRows) || 1;
      arrTotalEstimate = (sampleSum / sampledN) * totalDataRows;
    }
  }

  let autoScale = 1;
  if (arrTotalEstimate > 1_000_000_000) autoScale = 1_000_000;
  else if (arrTotalEstimate > 1_000_000) autoScale = 1_000;

  // Detect frequency from date column
  const dateColNum = headers.find(h => h.letter === dateCol)?.col_num;
  const dateValues: Date[] = [];
  if (dateColNum) {
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      const val = row.getCell(dateColNum).value;
      if (val instanceof Date) dateValues.push(val);
    });
  }
  const detectedFrequency = detectFrequency(dateValues);

  // Build column info for the frontend
  const columnInfo: ColumnInfo[] = headers.map(h => {
    const samples = (colSamples[h.col_num] || []).slice(0, 5).filter(v => v != null).map(v => String(v));
    return { letter: h.letter, header: h.header, sample_values: samples };
  });

  return {
    columns: columnInfo,
    detected_mapping: {
      date_col: dateCol || headers[0].letter,
      customer_id_col: customerIdCol || (headers.length > 1 ? headers[1].letter : headers[0].letter),
      arr_col: arrCol || headers[headers.length - 1].letter,
      attribute_cols: attributeCols,
    },
    row_count: totalDataRows,
    auto_scale_factor: autoScale,
    detected_frequency: detectedFrequency,
  };
}
