/**
 * Auto-detection service.
 * Reads sheet data and infers column types, scale factor, and data frequency.
 */
import type { Workbook } from 'exceljs';
import { colLetter } from './utils';

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
  header_row: number;           // which row the headers were found in
}

export interface DetectTableResult {
  columns: ColumnInfo[];
  date_columns: string[];       // column letters with date headers (the ARR data)
  customer_id_col: string;      // column letter for customer name/ID
  attribute_cols: { header: string; letter: string }[];
  row_count: number;
  auto_scale_factor: number;
  detected_frequency: string;
  header_row: number;           // which row the headers were found in
  date_header_row?: number;     // row containing date headers (may differ from header_row for wide tables)
}

function isDate(v: unknown): boolean {
  return v instanceof Date && !isNaN(v.getTime());
}

const MONTH_MAP: Record<string, number> = {
  jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2,
  apr: 3, april: 3, may: 4, jun: 5, june: 5,
  jul: 6, july: 6, aug: 7, august: 7, sep: 8, september: 8,
  oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11,
};

function normalizeYear(y: number): number {
  if (y < 100) return y >= 70 ? 1900 + y : 2000 + y;
  return y;
}

function isValidYear(y: number): boolean {
  return y >= 1990 && y <= 2100;
}

/**
 * Parse a header string into a Date if it looks like a recognizable date format.
 * Handles: Q1 2023, 1Q23, 2023-Q1, Jan-23, Mar 2023, 3/1/2023, bare 4-digit years, etc.
 * Returns null if not a recognizable date.
 */
export function parseHeaderDate(s: string): Date | null {
  if (!s || typeof s !== 'string') return null;
  const t = s.trim();
  if (t.length < 2) return null;

  // Quarter formats: Q1 2023, Q1'23, Q1-2023, Q1-23
  let m = t.match(/^Q([1-4])[\s'/-]*(\d{2,4})$/i);
  if (m) {
    const q = parseInt(m[1]);
    const y = normalizeYear(parseInt(m[2]));
    if (isValidYear(y)) return new Date(y, (q - 1) * 3, 1);
  }

  // Quarter formats: 1Q23, 1Q2023
  m = t.match(/^([1-4])Q[\s'/-]*(\d{2,4})$/i);
  if (m) {
    const q = parseInt(m[1]);
    const y = normalizeYear(parseInt(m[2]));
    if (isValidYear(y)) return new Date(y, (q - 1) * 3, 1);
  }

  // Quarter formats: 2023-Q1, 2023 Q1
  m = t.match(/^(\d{4})[\s-]+Q([1-4])$/i);
  if (m) {
    const y = parseInt(m[1]);
    const q = parseInt(m[2]);
    if (isValidYear(y)) return new Date(y, (q - 1) * 3, 1);
  }

  // Month-year: Jan-23, Jan 2023, January 2023, Jan-2023, Mar'23
  m = t.match(/^([A-Za-z]+)[\s'/-]*(\d{2,4})$/);
  if (m) {
    const mon = MONTH_MAP[m[1].toLowerCase()];
    if (mon !== undefined) {
      const y = normalizeYear(parseInt(m[2]));
      if (isValidYear(y)) return new Date(y, mon, 1);
    }
  }

  // Month-year reversed: 2023-Jan, 2023 March
  m = t.match(/^(\d{4})[\s-]+([A-Za-z]+)$/);
  if (m) {
    const y = parseInt(m[1]);
    const mon = MONTH_MAP[m[2].toLowerCase()];
    if (mon !== undefined && isValidYear(y)) return new Date(y, mon, 1);
  }

  // Numeric date: M/D/YYYY or MM/DD/YYYY
  m = t.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (m) {
    const y = normalizeYear(parseInt(m[3]));
    if (isValidYear(y)) {
      const d = new Date(y, parseInt(m[1]) - 1, parseInt(m[2]));
      if (!isNaN(d.getTime())) return d;
    }
  }

  // ISO-ish date: YYYY-MM-DD
  m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const y = parseInt(m[1]);
    if (isValidYear(y)) {
      const d = new Date(y, parseInt(m[2]) - 1, parseInt(m[3]));
      if (!isNaN(d.getTime())) return d;
    }
  }

  // Bare 4-digit year: 2023
  m = t.match(/^(\d{4})$/);
  if (m) {
    const y = parseInt(m[1]);
    if (isValidYear(y)) return new Date(y, 0, 1);
  }

  // Excel serial number: 5-digit number in valid date range (~1990-2100)
  m = t.match(/^(\d{5})$/);
  if (m) {
    const serial = parseInt(m[1]);
    if (serial >= 30000 && serial <= 80000) {
      const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
      const jsDate = new Date(excelEpoch.getTime() + serial * 86400000);
      if (!isNaN(jsDate.getTime()) && isValidYear(jsDate.getFullYear())) return jsDate;
    }
  }

  // Fallback: try native Date parser but validate year range
  const fallback = new Date(t);
  if (!isNaN(fallback.getTime())) {
    const fy = fallback.getFullYear();
    if (isValidYear(fy)) return fallback;
  }

  return null;
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
 * Scan the first maxScanRows rows to find the best header row.
 * Picks the row with the most populated cells (best header candidate).
 * Then also scans data rows to discover any columns with data not in the header.
 */
function findHeaderRow(
  ws: import('exceljs').Worksheet,
  _minCols = 2,
  maxScanRows = 20
): { headerRowNum: number; headers: { col_num: number; letter: string; header: string; rawValue: unknown }[] } {
  // Phase 1: Find the row with the most populated cells in the scan range
  let bestRow = -1;
  let bestCells: { col_num: number; letter: string; header: string; rawValue: unknown }[] = [];

  for (let r = 1; r <= maxScanRows; r++) {
    const row = ws.getRow(r);
    const cells: { col_num: number; letter: string; header: string; rawValue: unknown }[] = [];
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (cell.value != null) {
        cells.push({
          col_num: colNumber,
          letter: colLetter(colNumber),
          header: String(cell.value).trim(),
          rawValue: cell.value,
        });
      }
    });
    if (cells.length > bestCells.length) {
      bestRow = r;
      bestCells = cells;
    }
  }

  if (bestRow === -1 || bestCells.length < 2) {
    throw new Error(`No header row found in the first ${maxScanRows} rows of the sheet. Expected a row with at least 2 populated cells.`);
  }

  // Phase 2: Scan a few data rows below the header to discover columns that have
  // data but were empty in the header row (e.g., customer name col with no header label)
  const headerColNums = new Set(bestCells.map(c => c.col_num));
  const discoveredCols = new Map<number, string>(); // col_num -> first value seen

  for (let r = bestRow + 1; r <= bestRow + 5 && r <= maxScanRows + 10; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (cell.value != null && !headerColNums.has(colNumber) && !discoveredCols.has(colNumber)) {
        discoveredCols.set(colNumber, String(cell.value).trim());
      }
    });
  }

  // Add discovered columns to the header list with their column letter as a fallback label
  // Also check the row above/below the header for possible labels
  for (const [colNum, _firstVal] of discoveredCols) {
    let label = `Column ${colLetter(colNum)}`;

    // Check the header row itself (might have been skipped if null)
    // Check rows above the header for a label
    for (let r = Math.max(1, bestRow - 2); r <= bestRow; r++) {
      const val = ws.getRow(r).getCell(colNum).value;
      if (val != null && typeof val === 'string' && val.trim().length > 0) {
        label = val.trim();
        break;
      }
    }

    bestCells.push({
      col_num: colNum,
      letter: colLetter(colNum),
      header: label,
      rawValue: label,
    });
  }

  // Sort by column number so columns appear in natural order
  bestCells.sort((a, b) => a.col_num - b.col_num);

  return { headerRowNum: bestRow, headers: bestCells };
}

/**
 * Detect columns from a loaded ExcelJS workbook sheet.
 */
export function detectColumns(wb: Workbook, sheetName: string): DetectResult {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) throw new Error(`Sheet "${sheetName}" not found`);

  // Find header row dynamically
  const { headerRowNum, headers: rawHeaders } = findHeaderRow(ws);
  const headers = rawHeaders.map(h => ({ col_num: h.col_num, letter: h.letter, header: h.header }));

  if (headers.length === 0) {
    throw new Error('No headers found in the selected sheet.');
  }

  // Collect sample values per column
  const colSamples: Record<number, unknown[]> = {};
  for (const h of headers) colSamples[h.col_num] = [];

  let totalDataRows = 0;
  const allDates: Date[] = [];

  const dataStartRow = headerRowNum + 1;
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < dataStartRow) return; // skip header and anything above
    const firstCellVal = row.getCell(headers[0].col_num).value;
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
      if (rowNumber < dataStartRow) return;
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
    header_row: headerRowNum,
  };
}


/**
 * Detect columns from a cleaned/pivoted table where:
 * - Rows = customers
 * - Columns with date headers = ARR data by time period
 * - Other columns = customer ID, attributes
 */
export function detectTableColumns(wb: Workbook, sheetName: string): DetectTableResult {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) throw new Error(`Sheet "${sheetName}" not found`);

  // Find header row dynamically
  const { headerRowNum, headers } = findHeaderRow(ws);

  // Classify headers as date vs non-date
  let dateCols: { col_num: number; letter: string; header: string; date: Date }[] = [];
  let nonDateCols: { col_num: number; letter: string; header: string }[] = [];

  const dataStartRow = headerRowNum + 1;

  for (const h of headers) {
    // Try parsing the header text as a date
    const parsedFromText = parseHeaderDate(h.header);
    if (parsedFromText) {
      dateCols.push({ ...h, date: parsedFromText });
    } else if (h.rawValue instanceof Date && !isNaN(h.rawValue.getTime())) {
      // ExcelJS thinks it's a date but header text doesn't look like one.
      // Validate: if data rows below are mostly strings, it's a text column.
      let strCount = 0;
      let checked = 0;
      for (let r = dataStartRow; r < dataStartRow + 10; r++) {
        const val = ws.getRow(r).getCell(h.col_num).value;
        if (val == null) continue;
        checked++;
        if (typeof val === 'string') strCount++;
      }
      if (checked > 0 && strCount > checked * 0.5) {
        nonDateCols.push(h);
      } else {
        dateCols.push({ ...h, date: h.rawValue as Date });
      }
    } else {
      nonDateCols.push(h);
    }
  }

  // FALLBACK: If very few date columns detected (0-1) in a wide table,
  // findHeaderRow likely picked a data row instead of the actual header row.
  // Scan rows ABOVE the detected header for one containing Date values or
  // Excel serial numbers — that's the real date header row.
  let dateHeaderRow = headerRowNum; // default: same as header row
  if (dateCols.length <= 1 && headers.length > 5) {
    let bestDateRow = -1;
    let bestDateCells: { col_num: number; letter: string; header: string; date: Date }[] = [];

    for (let r = 1; r < headerRowNum; r++) {
      const row = ws.getRow(r);
      const foundDates: { col_num: number; letter: string; header: string; date: Date }[] = [];

      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.value == null) return;
        const letter = colLetter(colNumber);
        const headerText = String(cell.value).trim();

        // Check if cell is a Date object
        if (cell.value instanceof Date && !isNaN(cell.value.getTime())) {
          foundDates.push({ col_num: colNumber, letter, header: headerText, date: cell.value });
          return;
        }

        // Check if cell value is an Excel serial number (numeric, in date range)
        if (typeof cell.value === 'number') {
          const serial = cell.value;
          // Excel serial numbers for 1990-2100: ~32874 to ~73415
          if (serial >= 30000 && serial <= 80000) {
            // Convert Excel serial to JS Date: Excel epoch is Jan 0, 1900
            const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
            const jsDate = new Date(excelEpoch.getTime() + serial * 86400000);
            if (!isNaN(jsDate.getTime()) && isValidYear(jsDate.getFullYear())) {
              foundDates.push({ col_num: colNumber, letter, header: headerText, date: jsDate });
              return;
            }
          }
        }

        // Check if header text parses as a date
        const parsed = parseHeaderDate(headerText);
        if (parsed) {
          foundDates.push({ col_num: colNumber, letter, header: headerText, date: parsed });
        }
      });

      if (foundDates.length > bestDateCells.length) {
        bestDateRow = r;
        bestDateCells = foundDates;
      }
    }

    // If we found a row above with significantly more dates, use it
    if (bestDateCells.length >= 3) {
      dateHeaderRow = bestDateRow;
      // Use the date columns from the discovered row
      const dateColNums = new Set(bestDateCells.map(d => d.col_num));
      dateCols = bestDateCells;
      // Non-date columns are any header columns NOT in the date row
      nonDateCols = headers.filter(h => !dateColNums.has(h.col_num));

      // Try to find better labels for non-date columns from the date header row
      for (let i = 0; i < nonDateCols.length; i++) {
        const cell = ws.getRow(bestDateRow).getCell(nonDateCols[i].col_num);
        if (cell.value != null && typeof cell.value === 'string' && cell.value.trim().length > 0) {
          nonDateCols[i] = { ...nonDateCols[i], header: cell.value.trim() };
        }
      }
    }
  }

  // Sort date columns chronologically
  dateCols.sort((a, b) => a.date.getTime() - b.date.getTime());
  const dateColumnLetters = dateCols.map(d => d.letter);

  // Detect frequency from date headers
  const detectedFrequency = detectFrequency(dateCols.map(d => d.date));

  // Collect sample values for non-date columns to classify them
  const colSamples: Record<number, unknown[]> = {};
  for (const h of nonDateCols) colSamples[h.col_num] = [];

  let totalDataRows = 0;
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < dataStartRow) return;
    // Check first header column for data presence
    const firstCellVal = row.getCell(headers[0].col_num).value;
    if (firstCellVal == null) return;
    totalDataRows++;

    if (totalDataRows <= SAMPLE_ROWS) {
      for (const h of nonDateCols) {
        const val = row.getCell(h.col_num).value;
        colSamples[h.col_num].push(val);
      }
    }
  });

  // Classify non-date columns
  const colStats: Record<number, { str_ratio: number; unique_count: number }> = {};
  for (const h of nonDateCols) {
    const vals = colSamples[h.col_num].filter(v => v != null);
    if (vals.length === 0) {
      colStats[h.col_num] = { str_ratio: 0, unique_count: 0 };
      continue;
    }
    const strCount = vals.filter(v => typeof v === 'string').length;
    const uniqueCount = new Set(vals.map(v => String(v))).size;
    colStats[h.col_num] = { str_ratio: strCount / vals.length, unique_count: uniqueCount };
  }

  // Customer ID: non-date column with highest cardinality (prefer text, but accept any)
  let customerIdCol: string | null = null;
  let bestCardinality = -1;
  // First pass: prefer text columns
  for (const h of nonDateCols) {
    const s = colStats[h.col_num];
    if (s.str_ratio > 0.5 && s.unique_count > bestCardinality) {
      bestCardinality = s.unique_count;
      customerIdCol = h.letter;
    }
  }
  // Second pass: if no text column found, pick highest cardinality of any type
  if (!customerIdCol) {
    bestCardinality = -1;
    for (const h of nonDateCols) {
      const s = colStats[h.col_num];
      if (s.unique_count > bestCardinality) {
        bestCardinality = s.unique_count;
        customerIdCol = h.letter;
      }
    }
  }
  // Final fallback: first non-date column
  if (!customerIdCol && nonDateCols.length > 0) {
    customerIdCol = nonDateCols[0].letter;
  }

  // Attribute columns: ALL remaining non-date, non-customer-ID columns
  // (let the user decide which to include via the Identifiers step)
  const attributeCols: { header: string; letter: string }[] = [];
  for (const h of nonDateCols) {
    if (h.letter === customerIdCol) continue;
    const s = colStats[h.col_num];
    if (s.unique_count > 0) {
      attributeCols.push({ header: h.header, letter: h.letter });
    }
  }

  // Auto scale factor from date column values
  let sampleSum = 0;
  let sampleCount = 0;
  if (dateCols.length > 0) {
    const firstDateColNum = dateCols[0].col_num;
    const lastDateColNum = dateCols[dateCols.length - 1].col_num;
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber < dataStartRow) return;
      if (sampleCount >= SAMPLE_ROWS) return;
      sampleCount++;
      for (let cn = firstDateColNum; cn <= lastDateColNum; cn++) {
        const val = row.getCell(cn).value;
        if (typeof val === 'number') sampleSum += Math.abs(val);
      }
    });
  }
  const arrTotalEstimate = sampleCount > 0
    ? (sampleSum / sampleCount) * totalDataRows
    : 0;

  let autoScale = 1;
  if (arrTotalEstimate > 1_000_000_000) autoScale = 1_000_000;
  else if (arrTotalEstimate > 1_000_000) autoScale = 1_000;

  // Build column info for the frontend
  const columnInfo: ColumnInfo[] = headers.map(h => {
    const samples: string[] = [];
    if (colSamples[h.col_num]) {
      samples.push(...colSamples[h.col_num].slice(0, 5).filter(v => v != null).map(v => String(v)));
    }
    return { letter: h.letter, header: h.header, sample_values: samples };
  });

  return {
    columns: columnInfo,
    date_columns: dateColumnLetters,
    customer_id_col: customerIdCol || headers[0].letter,
    attribute_cols: attributeCols,
    row_count: totalDataRows,
    auto_scale_factor: autoScale,
    detected_frequency: detectedFrequency,
    header_row: dateHeaderRow !== headerRowNum ? dateHeaderRow : headerRowNum,
    date_header_row: dateHeaderRow !== headerRowNum ? dateHeaderRow : undefined,
  };
}
