/**
 * Summary tab generator.
 * Builds a per-period metric summary table segmented by a selected
 * customer identifier (attribute or cohort).
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig } from './types';
import type { CleanLayout } from './utils';
import { colLetter } from './utils';

export interface SummarySectionLayout {
  key: 'gross' | 'net' | 'logo' | 'pct_of_total' | 'dollars' | 'customers' | 'per_customer';
  startRow: number;   // -1 when the section has no rows
  numRows: number;
}

/**
 * Compute the row layout for the summary tab sections.
 * Sections start at row 10 with a blank row between them.
 * Retention sections only have rows for periods with a prior-year period.
 */
export function computeSummarySections(numDates: number, numDerived: number): SummarySectionLayout[] {
  const counts = [numDerived, numDerived, numDerived, numDates, numDates, numDates, numDates];
  const keys: SummarySectionLayout['key'][] = ['gross', 'net', 'logo', 'pct_of_total', 'dollars', 'customers', 'per_customer'];
  let r = 10;
  return keys.map((key, i) => {
    const numRows = counts[i];
    if (numRows <= 0) return { key, startRow: -1, numRows: 0 };
    const startRow = r;
    r += numRows + 1;
    return { key, startRow, numRows };
  });
}

export function generateSummaryTab(
  wb: Workbook, config: EngineConfig,
  cleanSheetName: string, cleanLayout: CleanLayout,
  firstDataRow: number, lastDataRow: number,
  granularity: string, segmentIdentifier: string, segmentValues: (string | { formula: string })[]
): string {
  const gran = granularity.charAt(0).toUpperCase() + granularity.slice(1);
  const sheetName = `${gran} Summary`;
  const ws = wb.addWorksheet(sheetName);

  const numDates = cleanLayout.num_dates;
  const yoyOffset = cleanLayout.yoy_offset;
  const numAttrs = cleanLayout.num_attrs;
  const attrNames = Object.keys(config.attributes);
  const metricLabel = (config.data_type || 'arr') === 'arr' ? 'ARR' : 'Revenue';
  const cohortHeader = `${gran} Cohort`;

  const allCol = 4;        // D
  const firstSegCol = 5;   // E
  const lastSegCol = firstSegCol + Math.max(segmentValues.length, 1) - 1;
  const checkCol = lastSegCol + 1;
  const firstSegCl = colLetter(firstSegCol);
  const lastSegCl = colLetter(lastSegCol);
  const checkCl = colLetter(checkCol);

  const sections = computeSummarySections(numDates, cleanLayout.num_derived);
  const sec = (key: SummarySectionLayout['key']) => sections.find(s => s.key === key)!;

  // --- Row 1: Check summary (sum of the check column next to the $ section) ---
  const dollarsSec = sec('dollars');
  ws.getCell(1, 1).value = 'Check Summary';
  ws.getCell(1, 2).value = {
    formula: `SUM(${checkCl}${dollarsSec.startRow}:${checkCl}${dollarsSec.startRow + dollarsSec.numRows - 1})`
  };

  // --- Row 3: Units ---
  ws.getCell(3, 1).value = 'Units';
  ws.getCell(3, 2).value = { formula: 'Control!$C$4' };

  // --- Row 5: Segment identifier selector ---
  ws.getCell(5, 1).value = 'Segment Identifier';
  const idCell = ws.getCell(5, 2);
  idCell.value = segmentIdentifier;
  idCell.dataValidation = {
    type: 'list',
    allowBlank: false,
    formulae: [`"${[...attrNames, cohortHeader].join(',')}"`],
  };

  // --- Row 7: identifier banner ---
  ws.getCell(7, allCol).value = { formula: '$B$5' };

  // --- Row 8: column headers ---
  ws.getCell(8, 2).value = 'Metric';
  ws.getCell(8, 3).value = 'Period';
  ws.getCell(8, allCol).value = 'All';
  for (let i = 0; i < segmentValues.length; i++) {
    ws.getCell(8, firstSegCol + i).value = segmentValues[i];
  }
  ws.getCell(8, checkCol).value = 'Check';

  // Dynamic criteria column on the clean tab selected by B5
  const critStartCol = numAttrs > 0 ? cleanLayout.attr_start : cleanLayout.cohort;
  const critStartCl = colLetter(critStartCol);
  const critEndCl = colLetter(cleanLayout.cohort);
  const CRIT = `INDEX('${cleanSheetName}'!$${critStartCl}$${firstDataRow}:$${critEndCl}$${lastDataRow},0,MATCH($B$5,'${cleanSheetName}'!$${critStartCl}$6:$${critEndCl}$6,0))`;

  function R(colNum: number): string {
    const cl = colLetter(colNum);
    return `'${cleanSheetName}'!$${cl}$${firstDataRow}:$${cl}$${lastDataRow}`;
  }
  function allSum(colNum: number): string {
    return `SUM(${R(colNum)})`;
  }
  function allCount(colNum: number): string {
    return `COUNTIF(${R(colNum)},"<>"&0)`;
  }
  function segSum(colNum: number, L: string): string {
    return `SUMIFS(${R(colNum)},${CRIT},${L}$8)`;
  }
  function segCount(colNum: number, L: string): string {
    return `COUNTIFS(${R(colNum)},"<>"&0,${CRIT},${L}$8)`;
  }

  const sectionLabels: Record<SummarySectionLayout['key'], string> = {
    gross: 'Gross Retention',
    net: 'Net Retention',
    logo: 'Logo Retention',
    pct_of_total: `% of ${metricLabel}`,
    dollars: `$ ${metricLabel}`,
    customers: 'Ending Customers',
    per_customer: `${metricLabel} per Customer`,
  };

  const lastDataCol = checkCol;
  for (const section of sections) {
    if (section.numRows <= 0) continue;
    ws.getCell(section.startRow, 2).value = sectionLabels[section.key];

    for (let rowIdx = 0; rowIdx < section.numRows; rowIdx++) {
      const r = section.startRow + rowIdx;
      // Retention sections index into derived periods; others cover all periods
      const periodIdx = rowIdx + (section.key === 'gross' || section.key === 'net' || section.key === 'logo' ? yoyOffset : 0);
      const arrCl = colLetter(cleanLayout.arr_start + periodIdx);
      ws.getCell(r, 3).value = { formula: `'${cleanSheetName}'!${arrCl}$6` };

      for (let c = allCol; c <= lastSegCol; c++) {
        const L = colLetter(c);
        const isAll = c === allCol;
        const sum = (colNum: number) => isAll ? allSum(colNum) : segSum(colNum, L);
        const cnt = (colNum: number) => isAll ? allCount(colNum) : segCount(colNum, L);
        let formula: string | null = null;

        switch (section.key) {
          case 'gross':
          case 'net': {
            const priorArr = cleanLayout.arr_start + periodIdx - yoyOffset;
            const churn = cleanLayout.churn_start + rowIdx;
            const down = cleanLayout.downsell_start + rowIdx;
            const up = cleanLayout.upsell_start + rowIdx;
            const extra = section.key === 'net' ? `+${sum(up)}` : '';
            formula = `IFERROR((${sum(priorArr)}+${sum(churn)}+${sum(down)}${extra})/${sum(priorArr)},"n.a.")`;
            break;
          }
          case 'logo': {
            const priorArr = cleanLayout.arr_start + periodIdx - yoyOffset;
            const churn = cleanLayout.churn_start + rowIdx;
            formula = `IFERROR((${cnt(priorArr)}-${cnt(churn)})/${cnt(priorArr)},"n.a.")`;
            break;
          }
          case 'pct_of_total': {
            const dRow = dollarsSec.startRow + rowIdx;
            formula = `IFERROR(${L}${dRow}/$D${dRow},"n.a.")`;
            break;
          }
          case 'dollars': {
            formula = `${sum(cleanLayout.arr_start + periodIdx)}/$B$3`;
            break;
          }
          case 'customers': {
            formula = cnt(cleanLayout.arr_start + periodIdx);
            break;
          }
          case 'per_customer': {
            const custSec = sec('customers');
            formula = `IFERROR(${L}${dollarsSec.startRow + rowIdx}/${L}${custSec.startRow + rowIdx},"n.a.")`;
            break;
          }
        }
        if (formula) ws.getCell(r, c).value = { formula };
      }

      // Check column on $ rows: All minus the sum of the segment columns
      if (section.key === 'dollars') {
        const f = segmentValues.length > 0
          ? `D${r}-SUM(${firstSegCl}${r}:${lastSegCl}${r})`
          : '0';
        ws.getCell(r, checkCol).value = { formula: f };
      }
    }
  }

  // Column widths
  ws.getColumn(1).width = 14;
  ws.getColumn(2).width = 24;
  ws.getColumn(3).width = 12;
  for (let c = allCol; c <= lastDataCol; c++) {
    ws.getColumn(c).width = 13;
  }

  // Freeze panes
  ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 8, showGridLines: false }];

  return sheetName;
}
