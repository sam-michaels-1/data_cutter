/**
 * Data Summary tab generator.
 * One block per customer attribute: distinct values with customer counts
 * and share of customers, plus a check row against COUNTA of the column.
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig } from './types';
import type { CleanLayout } from './utils';
import { colLetter } from './utils';

const SPARE_ROWS = 10;
const BLOCK_WIDTH = 4;  // 3 data columns + 1 spacer

export function generateDataSummaryTab(
  wb: Workbook, config: EngineConfig,
  cleanSheetName: string, cleanLayout: CleanLayout,
  firstDataRow: number, lastDataRow: number,
  attrValueCounts: Record<string, number>
): { sheetName: string; valueCells: Record<string, { col: string; firstRow: number; lastRow: number }> } {
  const sheetName = 'Data Summary';
  const ws = wb.addWorksheet(sheetName);

  const attrNames = Object.keys(config.attributes);
  const checkRefs: string[] = [];
  const valueCells: Record<string, { col: string; firstRow: number; lastRow: number }> = {};

  for (let k = 0; k < attrNames.length; k++) {
    const attrName = attrNames[k];
    const c = 2 + k * BLOCK_WIDTH;                 // block start column (B, F, J, ...)
    const valCl = colLetter(c);
    const cntCl = colLetter(c + 1);
    const pctCl = colLetter(c + 2);
    const cleanCl = colLetter(cleanLayout.attr_start + k);

    const distinct = attrValueCounts[attrName] || 0;
    const numRows = distinct + SPARE_ROWS;
    const firstRow = 7;
    const lastRow = firstRow + numRows - 1;
    const totalRow = lastRow + 1;
    const checkRow = totalRow + 1;

    const cleanRange = `'${cleanSheetName}'!$${cleanCl}$${firstDataRow}:$${cleanCl}$${lastDataRow}`;

    // Headers
    ws.getCell(5, c).value = attrName;
    ws.getCell(6, c).value = 'Value';
    ws.getCell(6, c + 1).value = '# Customers';
    ws.getCell(6, c + 2).value = '% of Customers';

    // Value rows (dynamic sorted-unique extraction; per-cell INDEX because
    // ExcelJS cannot write dynamic-array spill metadata)
    for (let r = firstRow; r <= lastRow; r++) {
      const n = r - firstRow + 1;
      ws.getCell(r, c).value = {
        formula: `IFERROR(INDEX(_xlfn._xlws.SORT(_xlfn.UNIQUE(${cleanRange})),${n}),"")`
      };
      ws.getCell(r, c + 1).value = {
        formula: `IF(${valCl}${r}="","",COUNTIFS(${cleanRange},${valCl}${r}))`
      };
      ws.getCell(r, c + 2).value = {
        formula: `IF(${valCl}${r}="","",${cntCl}${r}/${cntCl}${totalRow})`
      };
    }

    // Total row
    ws.getCell(totalRow, c).value = 'Total';
    ws.getCell(totalRow, c + 1).value = { formula: `SUM(${cntCl}${firstRow}:${cntCl}${lastRow})` };
    ws.getCell(totalRow, c + 2).value = { formula: `SUM(${pctCl}${firstRow}:${pctCl}${lastRow})` };

    // Check row: counted customers vs all rows with a value
    ws.getCell(checkRow, c).value = 'Check';
    ws.getCell(checkRow, c + 1).value = {
      formula: `${cntCl}${totalRow}-COUNTA(${cleanRange})`
    };
    checkRefs.push(`${cntCl}${checkRow}`);

    valueCells[attrName] = { col: valCl, firstRow, lastRow: firstRow + distinct - 1 };
  }

  // --- Row 1: Check summary ---
  ws.getCell(1, 1).value = 'Check Summary';
  ws.getCell(1, 2).value = { formula: `SUM(${checkRefs.join(',')})` };

  // Column widths
  for (let k = 0; k < attrNames.length; k++) {
    const c = 2 + k * BLOCK_WIDTH;
    ws.getColumn(c).width = 18;
    ws.getColumn(c + 1).width = 12;
    ws.getColumn(c + 2).width = 14;
  }

  // Freeze panes
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 6, showGridLines: false }];

  return { sheetName, valueCells };
}
