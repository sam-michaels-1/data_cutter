/**
 * Top Customer Analysis tab generator.
 * Port of data-pack-app/engine/top_customers.py for client-side use with ExcelJS.
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig } from './types';
import type { CleanLayout } from './utils';
import { colLetter } from './utils';

export const TOP_N = 25;
export const CONCENTRATION_TIERS = [3, 5, 10, 25];

export function generateTopCustomersTab(
  wb: Workbook, config: EngineConfig,
  cleanSheetName: string, cleanLayout: CleanLayout,
  firstDataRow: number, lastDataRow: number
): string {
  const sheetName = 'Annual Top Customer Analysis';
  const ws = wb.addWorksheet(sheetName);

  const numDates = cleanLayout.num_dates;
  const numAttrs = cleanLayout.num_attrs;
  const attrNames = Object.keys(config.attributes);

  const rankNumCol = 2;
  const custIdCol = 3;
  const attrStart = 4;
  const attrEnd = attrStart + numAttrs - 1;
  const cohortCol = attrEnd + 1;

  const s1Start = cohortCol + 1;
  const s1End = s1Start + numDates - 1;
  const s2Start = s1End + 2;
  const s2End = s2Start + numDates - 2;
  const s3Start = s2End + 2;
  const cdrFirst = firstDataRow;
  const cdrLast = lastDataRow;
  const cleanRankCol = colLetter(cleanLayout.rank);
  const cleanCustIdCol = colLetter(cleanLayout.cust_id);
  const metricLabel = (config.data_type || 'arr') === 'arr' ? 'ARR' : 'Revenue';

  // Row 3: Units
  ws.getCell(3, 1).value = 'Units';
  ws.getCell(3, rankNumCol).value = { formula: 'Control!$C$4' };
  const unitsCell = `$${colLetter(rankNumCol)}$3`;

  // Row 5: Section headers
  ws.getCell(5, s1Start).value = metricLabel;
  ws.getCell(5, s2Start).value = '% YoY Growth';
  ws.getCell(5, s3Start).value = '% of Total';

  // Row 6: Column headers
  ws.getCell(6, custIdCol).value = 'Customer ID';
  for (let i = 0; i < attrNames.length; i++) {
    ws.getCell(6, attrStart + i).value = attrNames[i];
  }
  ws.getCell(6, cohortCol).value = 'Annual Cohort';

  for (let i = 0; i < numDates; i++) {
    const cleanCol = cleanLayout.arr_start + i;
    ws.getCell(6, s1Start + i).value = { formula: `'${cleanSheetName}'!${colLetter(cleanCol)}6` };
  }
  for (let i = 0; i < numDates - 1; i++) {
    ws.getCell(6, s2Start + i).value = { formula: `${colLetter(s1Start + i + 1)}6` };
  }
  for (let i = 0; i < numDates; i++) {
    ws.getCell(6, s3Start + i).value = { formula: `${colLetter(s1Start + i)}6` };
  }

  // Data rows
  const firstCustomerRow = 7;
  const lastCustomerRow = firstCustomerRow + TOP_N - 1;

  for (let rank = 1; rank <= TOP_N; rank++) {
    const row = firstCustomerRow + rank - 1;

    if (rank === 1) {
      ws.getCell(row, rankNumCol).value = 1;
    } else {
      ws.getCell(row, rankNumCol).value = { formula: `${colLetter(rankNumCol)}${row - 1}+1` };
    }

    const rnl = colLetter(rankNumCol);
    ws.getCell(row, custIdCol).value = {
      formula: `_xlfn.XLOOKUP($${rnl}${row},'${cleanSheetName}'!$${cleanRankCol}$${cdrFirst}:$${cleanRankCol}$${cdrLast},'${cleanSheetName}'!${colLetter(cleanLayout.cust_id)}$${cdrFirst}:${colLetter(cleanLayout.cust_id)}$${cdrLast})`
    };

    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      const cleanAttrCol = cleanLayout.attr_start + attrIdx;
      ws.getCell(row, attrStart + attrIdx).value = {
        formula: `_xlfn.XLOOKUP($${rnl}${row},'${cleanSheetName}'!$${cleanRankCol}$${cdrFirst}:$${cleanRankCol}$${cdrLast},'${cleanSheetName}'!${colLetter(cleanAttrCol)}$${cdrFirst}:${colLetter(cleanAttrCol)}$${cdrLast})`
      };
    }

    ws.getCell(row, cohortCol).value = {
      formula: `_xlfn.XLOOKUP($${rnl}${row},'${cleanSheetName}'!$${cleanRankCol}$${cdrFirst}:$${cleanRankCol}$${cdrLast},'${cleanSheetName}'!${colLetter(cleanLayout.cohort)}$${cdrFirst}:${colLetter(cleanLayout.cohort)}$${cdrLast})`
    };

    // Section 1: ARR
    const cidl = colLetter(custIdCol);
    for (let i = 0; i < numDates; i++) {
      const cleanCol = cleanLayout.arr_start + i;
      ws.getCell(row, s1Start + i).value = {
        formula: `SUMIFS('${cleanSheetName}'!${colLetter(cleanCol)}$${cdrFirst}:${colLetter(cleanCol)}$${cdrLast},'${cleanSheetName}'!$${cleanCustIdCol}$${cdrFirst}:$${cleanCustIdCol}$${cdrLast},$${cidl}${row})/${unitsCell}`
      };
    }

    // Section 2: % YoY Growth
    for (let i = 0; i < numDates - 1; i++) {
      const currCol = colLetter(s1Start + i + 1);
      const prevCol = colLetter(s1Start + i);
      ws.getCell(row, s2Start + i).value = { formula: `IFERROR(${currCol}${row}/${prevCol}${row}-1,"n.a.")` };
    }

    // Section 3: % of Total
    const totalRow = lastCustomerRow + 3;
    for (let i = 0; i < numDates; i++) {
      const arrCol = colLetter(s1Start + i);
      ws.getCell(row, s3Start + i).value = { formula: `${arrCol}${row}/${arrCol}$${totalRow}` };
    }
  }

  // Summary rows
  const rTopTotal = lastCustomerRow + 1;
  const rOther = rTopTotal + 1;
  const rTotal = rOther + 1;

  ws.getCell(rTopTotal, custIdCol).value = `Top ${TOP_N} Customers`;
  for (let i = 0; i < numDates; i++) {
    const arrCl = colLetter(s1Start + i);
    ws.getCell(rTopTotal, s1Start + i).value = { formula: `SUM(${arrCl}${firstCustomerRow}:${arrCl}${lastCustomerRow})` };
    if (i > 0) {
      const curr = colLetter(s1Start + i);
      const prev = colLetter(s1Start + i - 1);
      ws.getCell(rTopTotal, s2Start + i - 1).value = { formula: `IFERROR(${curr}${rTopTotal}/${prev}${rTopTotal}-1,"n.a.")` };
    }
    ws.getCell(rTopTotal, s3Start + i).value = { formula: `${arrCl}${rTopTotal}/${arrCl}$${rTotal}` };
  }

  ws.getCell(rOther, custIdCol).value = '(+) Other Customers';
  for (let i = 0; i < numDates; i++) {
    const arrCl = colLetter(s1Start + i);
    ws.getCell(rOther, s1Start + i).value = { formula: `${arrCl}${rTotal}-${arrCl}${rTopTotal}` };
    if (i > 0) {
      const curr = colLetter(s1Start + i);
      const prev = colLetter(s1Start + i - 1);
      ws.getCell(rOther, s2Start + i - 1).value = { formula: `IFERROR(${curr}${rOther}/${prev}${rOther}-1,"n.a.")` };
    }
    ws.getCell(rOther, s3Start + i).value = { formula: `${arrCl}${rOther}/${arrCl}$${rTotal}` };
  }

  ws.getCell(rTotal, custIdCol).value = `Total ${metricLabel}`;
  for (let i = 0; i < numDates; i++) {
    const cleanCol = cleanLayout.arr_start + i;
    ws.getCell(rTotal, s1Start + i).value = {
      formula: `SUMIFS('${cleanSheetName}'!${colLetter(cleanCol)}$${cdrFirst}:${colLetter(cleanCol)}$${cdrLast},'${cleanSheetName}'!$${cleanCustIdCol}$${cdrFirst}:$${cleanCustIdCol}$${cdrLast},"<>")/${unitsCell}`
    };
    if (i > 0) {
      const curr = colLetter(s1Start + i);
      const prev = colLetter(s1Start + i - 1);
      ws.getCell(rTotal, s2Start + i - 1).value = { formula: `IFERROR(${curr}${rTotal}/${prev}${rTotal}-1,"n.a.")` };
    }
  }

  // Concentration Memo
  const rMemoStart = rTotal + 2;
  ws.getCell(rMemoStart, custIdCol).value = 'Memo:';

  for (let tierIdx = 0; tierIdx < CONCENTRATION_TIERS.length; tierIdx++) {
    const tier = CONCENTRATION_TIERS[tierIdx];
    const row = rMemoStart + 1 + tierIdx;
    ws.getCell(row, rankNumCol).value = tier;
    ws.getCell(row, custIdCol).value = { formula: `"Top "&${colLetter(rankNumCol)}${row}&" Customers Concentration"` };

    const rnl = colLetter(rankNumCol);
    for (let i = 0; i < numDates; i++) {
      const arrCl = colLetter(s1Start + i);
      ws.getCell(row, s1Start + i).value = {
        formula: `SUMIFS(${arrCl}$${firstCustomerRow}:${arrCl}$${lastCustomerRow},$${rnl}$${firstCustomerRow}:$${rnl}$${lastCustomerRow},"<="&$${rnl}${row})`
      };
      if (i > 0) {
        const curr = colLetter(s1Start + i);
        const prev = colLetter(s1Start + i - 1);
        ws.getCell(row, s2Start + i - 1).value = { formula: `IFERROR(${curr}${row}/${prev}${row}-1,"n.a.")` };
      }
      ws.getCell(row, s3Start + i).value = { formula: `${arrCl}${row}/${arrCl}$${rTotal}` };
    }
  }

  return sheetName;
}
