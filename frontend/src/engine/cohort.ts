/**
 * Cohort tab generators.
 * Port of data-pack-app/engine/cohort.py for client-side use with ExcelJS.
 */
import type { Workbook, Worksheet } from 'exceljs';
import type { EngineConfig, FilterBlock } from './types';
import type { CleanLayout } from './utils';
import { colLetter } from './utils';

export function generateCohortTab(
  wb: Workbook, config: EngineConfig,
  cleanSheetName: string, cleanLayout: CleanLayout,
  firstDataRow: number, lastDataRow: number,
  granularity: string, filterBlocks: FilterBlock[]
): string {
  const sheetName = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Cohort`;
  const ws = wb.addWorksheet(sheetName);

  const numDates = cleanLayout.num_dates;
  const numAttrs = cleanLayout.num_attrs;
  const attrNames = Object.keys(config.attributes);
  const numCohorts = numDates;

  // Layout columns
  const qCol = 2;
  const yCol = 3;
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

  // Counter row (row 1)
  ws.getCell(1, s3DataStart).value = 0;
  for (let i = 1; i < numDates; i++) {
    ws.getCell(1, s3DataStart + i).value = { formula: `${colLetter(s3DataStart + i - 1)}1+1` };
  }
  for (let i = 0; i < numDates; i++) {
    ws.getCell(1, s4DataStart + i).value = { formula: `${colLetter(s3DataStart + i)}1` };
  }

  ws.getCell(1, 1).value = 'Check Summary';
  ws.getCell(3, 1).value = 'Units';
  ws.getCell(3, qCol).value = { formula: 'Control!$C$4' };
  const unitsCell = `$${colLetter(qCol)}$3`;

  const periodPrefix = granularity === 'quarterly' ? 'Q' : 'Y';
  const metricLabel = (config.data_type || 'arr') === 'arr' ? 'ARR' : 'Revenue';
  const cdrFirst = firstDataRow;
  const cdrLast = lastDataRow;

  const checkRefs: string[] = [];

  for (let blockIdx = 0; blockIdx < filterBlocks.length; blockIdx++) {
    const block = filterBlocks[blockIdx];
    const blockStart = 6 + blockIdx * (numCohorts + 9);

    writeCohortBlock(
      ws, blockStart, block, config,
      cleanSheetName, cleanLayout, cdrFirst, cdrLast,
      qCol, yCol, filterStart, filterEnd, cohortLabelCol,
      s1Start, s1End, s2Start, s2End,
      s3LabelCol, s3StartValCol, s3DataStart, s3DataEnd,
      s4LabelCol, s4StartValCol, s4DataStart, s4DataEnd,
      numDates, numCohorts, numAttrs, attrNames,
      unitsCell, granularity, periodPrefix, metricLabel
    );

    const checkRow = blockStart + 3 + numCohorts + 4;
    checkRefs.push(`${colLetter(s1Start)}${checkRow}:${colLetter(s1End)}${checkRow}`);
    checkRefs.push(`${colLetter(s2Start)}${checkRow}:${colLetter(s2End)}${checkRow}`);
  }

  if (checkRefs.length > 0) {
    ws.getCell(1, 2).value = { formula: `SUM(${checkRefs.join(',')})` };
  }

  return sheetName;
}

function writeCohortBlock(
  ws: Worksheet, startRow: number, block: FilterBlock, _config: EngineConfig,
  cleanSheet: string, cleanLayout: CleanLayout,
  cdrFirst: number, cdrLast: number,
  qCol: number, yCol: number, filterStart: number, _filterEnd: number, cohortLabelCol: number,
  s1Start: number, s1End: number, s2Start: number, s2End: number,
  s3Label: number, s3StartVal: number, s3DataStart: number, _s3DataEnd: number,
  s4Label: number, s4StartVal: number, s4DataStart: number, _s4DataEnd: number,
  numDates: number, numCohorts: number, numAttrs: number, attrNames: string[],
  unitsCell: string, granularity: string, periodPrefix: string, metricLabel: string
): void {
  const { title, filters } = block;
  const rTitle = startRow;
  const rSectionHeaders = startRow + 2;
  const rHeaders = startRow + 3;

  // Title
  ws.getCell(rTitle, qCol).value = `${title} ${granularity.charAt(0).toUpperCase() + granularity.slice(1)} ${metricLabel} by Cohort`;

  // Section headers
  ws.getCell(rSectionHeaders, s1Start).value = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} ${metricLabel}`;
  ws.getCell(rSectionHeaders, s2Start).value = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Customers`;
  ws.getCell(rSectionHeaders, s3Label).value = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} ${metricLabel} Retention`;
  ws.getCell(rSectionHeaders, s4Label).value = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Logo Retention`;

  // Column headers
  ws.getCell(rHeaders, qCol).value = 'Quarter';
  ws.getCell(rHeaders, yCol).value = 'Year';
  for (let attrIdx = 0; attrIdx < attrNames.length; attrIdx++) {
    ws.getCell(rHeaders, filterStart + attrIdx).value = attrNames[attrIdx];
  }
  ws.getCell(rHeaders, cohortLabelCol).value = 'Cohort';

  // Period headers
  for (let i = 0; i < numDates; i++) {
    const cleanDateCol = cleanLayout.arr_start + i;
    ws.getCell(rHeaders, s1Start + i).value = { formula: `'${cleanSheet}'!${colLetter(cleanDateCol)}$6` };
    ws.getCell(rHeaders, s2Start + i).value = { formula: `${colLetter(s1Start + i)}${rHeaders}` };
  }

  // Retention headers
  ws.getCell(rHeaders, s3Label).value = 'Cohort';
  ws.getCell(rHeaders, s3StartVal).value = 'Starting Size';
  ws.getCell(rHeaders, s4Label).value = 'Cohort';
  ws.getCell(rHeaders, s4StartVal).value = 'Starting Size';

  for (let i = 0; i < numDates; i++) {
    ws.getCell(rHeaders, s3DataStart + i).value = { formula: `"${periodPrefix}"&${colLetter(s3DataStart + i)}$1` };
    ws.getCell(rHeaders, s4DataStart + i).value = { formula: `"${periodPrefix}"&${colLetter(s4DataStart + i)}$1` };
  }

  const firstCohortRow = rHeaders + 1;
  const srcArrStart = colLetter(cleanLayout.arr_start);

  // Cohort data rows
  for (let cohortIdx = 0; cohortIdx < numCohorts; cohortIdx++) {
    const row = firstCohortRow + cohortIdx;

    // Quarter
    if (cohortIdx === 0) {
      ws.getCell(row, qCol).value = { formula: `'${cleanSheet}'!$${srcArrStart}$2` };
    } else {
      const prevQ = colLetter(qCol);
      if (granularity === 'quarterly') {
        ws.getCell(row, qCol).value = { formula: `IF(${prevQ}${row - 1}=4,1,${prevQ}${row - 1}+1)` };
      } else {
        ws.getCell(row, qCol).value = { formula: `${prevQ}${row - 1}` };
      }
    }

    // Year
    if (cohortIdx === 0) {
      ws.getCell(row, yCol).value = { formula: `'${cleanSheet}'!$${srcArrStart}$3` };
    } else {
      const prevY = colLetter(yCol);
      const prevQ = colLetter(qCol);
      if (granularity === 'quarterly') {
        ws.getCell(row, yCol).value = { formula: `IF(${prevQ}${row - 1}=4,${prevY}${row - 1}+1,${prevY}${row - 1})` };
      } else {
        ws.getCell(row, yCol).value = { formula: `${prevY}${row - 1}+1` };
      }
    }

    // Filters
    for (let attrIdx = 0; attrIdx < attrNames.length; attrIdx++) {
      const fc = filterStart + attrIdx;
      if (cohortIdx === 0) {
        ws.getCell(row, fc).value = filters[attrNames[attrIdx]] || '<>';
      } else {
        ws.getCell(row, fc).value = { formula: `${colLetter(fc)}${row - 1}` };
      }
    }

    // Cohort label
    const ql = colLetter(qCol);
    const yl = colLetter(yCol);
    if (granularity === 'quarterly') {
      ws.getCell(row, cohortLabelCol).value = { formula: `"Q"&${ql}${row}&"'"&RIGHT(${yl}${row},2)` };
    } else {
      ws.getCell(row, cohortLabelCol).value = { formula: `"FY"&"'"&RIGHT(${yl}${row},2)` };
    }

    // SUMIFS criteria
    const gl = colLetter(cohortLabelCol);
    const critParts: string[] = [];
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      const cleanCol = cleanLayout.attr_start + attrIdx;
      critParts.push(
        `'${cleanSheet}'!$${colLetter(cleanCol)}$${cdrFirst}:$${colLetter(cleanCol)}$${cdrLast},$${colLetter(filterStart + attrIdx)}${row}`
      );
    }
    critParts.push(
      `'${cleanSheet}'!$${colLetter(cleanLayout.cohort)}$${cdrFirst}:$${colLetter(cleanLayout.cohort)}$${cdrLast},$${gl}${row}`
    );
    const criteriaStr = critParts.join(',');

    // Section 1: ARR
    for (let pi = 0; pi < numDates; pi++) {
      const cleanCol = cleanLayout.arr_start + pi;
      const sumRange = `'${cleanSheet}'!${colLetter(cleanCol)}$${cdrFirst}:${colLetter(cleanCol)}$${cdrLast}`;
      ws.getCell(row, s1Start + pi).value = { formula: `SUMIFS(${sumRange},${criteriaStr})/${unitsCell}` };
    }

    // Section 2: Customer count
    for (let pi = 0; pi < numDates; pi++) {
      const cleanCol = cleanLayout.arr_start + pi;
      const countRange = `'${cleanSheet}'!${colLetter(cleanCol)}$${cdrFirst}:${colLetter(cleanCol)}$${cdrLast}`;
      ws.getCell(row, s2Start + pi).value = { formula: `COUNTIFS(${countRange},"<>"&0,${criteriaStr})` };
    }

    // Section 3: ARR Retention
    const s1s = colLetter(s1Start);
    const s1e = colLetter(s1End);
    ws.getCell(row, s3Label).value = { formula: `${gl}${row}` };
    const s3sv = colLetter(s3StartVal);
    ws.getCell(row, s3StartVal).value = {
      formula: `_xlfn.XLOOKUP(${colLetter(s3Label)}${row},${colLetter(s1Start)}$${rHeaders}:${colLetter(s1End)}$${rHeaders},${s1s}${row}:${s1e}${row})`
    };

    for (let pi = 0; pi < numDates - cohortIdx; pi++) {
      const arrCol = s1Start + cohortIdx + pi;
      ws.getCell(row, s3DataStart + pi).value = { formula: `${colLetter(arrCol)}${row}/$${s3sv}${row}` };
    }

    // Section 4: Logo Retention
    const s2s = colLetter(s2Start);
    const s2e = colLetter(s2End);
    ws.getCell(row, s4Label).value = { formula: `${gl}${row}` };
    const s4sv = colLetter(s4StartVal);
    ws.getCell(row, s4StartVal).value = {
      formula: `_xlfn.XLOOKUP(${colLetter(s4Label)}${row},${colLetter(s2Start)}$${rHeaders}:${colLetter(s2End)}$${rHeaders},${s2s}${row}:${s2e}${row})`
    };

    for (let pi = 0; pi < numDates - cohortIdx; pi++) {
      const countCol = s2Start + cohortIdx + pi;
      ws.getCell(row, s4DataStart + pi).value = { formula: `${colLetter(countCol)}${row}/$${s4sv}${row}` };
    }
  }

  // Summary rows
  const lastCohortRow = firstCohortRow + numCohorts - 1;
  const rTotal = lastCohortRow + 1;
  const rMedian = rTotal + 1;
  const rWeighted = rMedian + 1;
  const rCheck = rWeighted + 1;

  // Filter columns for summary rows
  for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
    const fc = filterStart + attrIdx;
    for (const r of [rTotal, rMedian, rWeighted, rCheck]) {
      ws.getCell(r, fc).value = { formula: `${colLetter(fc)}${r - 1}` };
    }
  }

  // Total row
  for (let pi = 0; pi < numDates; pi++) {
    const dc = s1Start + pi;
    const dcl = colLetter(dc);
    ws.getCell(rTotal, dc).value = { formula: `SUM(${dcl}${firstCohortRow}:${dcl}${lastCohortRow})` };

    const dc2 = s2Start + pi;
    const dc2l = colLetter(dc2);
    ws.getCell(rTotal, dc2).value = { formula: `SUM(${dc2l}${firstCohortRow}:${dc2l}${lastCohortRow})` };
  }

  // Average, Median, Weighted for retention
  for (let pi = 0; pi < numDates; pi++) {
    const dc3 = s3DataStart + pi;
    const dc3l = colLetter(dc3);
    ws.getCell(rTotal, dc3).value = { formula: `AVERAGE(${dc3l}${firstCohortRow}:${dc3l}${lastCohortRow})` };
    ws.getCell(rMedian, dc3).value = { formula: `MEDIAN(${dc3l}${firstCohortRow}:${dc3l}${lastCohortRow})` };

    const s3svl = colLetter(s3StartVal);
    ws.getCell(rWeighted, dc3).value = {
      formula: `SUMPRODUCT((${dc3l}${firstCohortRow}:${dc3l}${lastCohortRow}<>"")*${dc3l}${firstCohortRow}:${dc3l}${lastCohortRow},$${s3svl}${firstCohortRow}:$${s3svl}${lastCohortRow})/SUMPRODUCT((${dc3l}${firstCohortRow}:${dc3l}${lastCohortRow}<>"")*1,$${s3svl}${firstCohortRow}:$${s3svl}${lastCohortRow})`
    };

    const dc4 = s4DataStart + pi;
    const dc4l = colLetter(dc4);
    ws.getCell(rTotal, dc4).value = { formula: `AVERAGE(${dc4l}${firstCohortRow}:${dc4l}${lastCohortRow})` };
    ws.getCell(rMedian, dc4).value = { formula: `MEDIAN(${dc4l}${firstCohortRow}:${dc4l}${lastCohortRow})` };

    const s4svl = colLetter(s4StartVal);
    ws.getCell(rWeighted, dc4).value = {
      formula: `SUMPRODUCT((${dc4l}${firstCohortRow}:${dc4l}${lastCohortRow}<>"")*${dc4l}${firstCohortRow}:${dc4l}${lastCohortRow},$${s4svl}${firstCohortRow}:$${s4svl}${lastCohortRow})/SUMPRODUCT((${dc4l}${firstCohortRow}:${dc4l}${lastCohortRow}<>"")*1,$${s4svl}${firstCohortRow}:$${s4svl}${lastCohortRow})`
    };
  }

  // Check row
  for (let pi = 0; pi < numDates; pi++) {
    const dc = s1Start + pi;
    const dcl = colLetter(dc);
    const cleanCol = cleanLayout.arr_start + pi;
    const critCheck: string[] = [];
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      const cc = cleanLayout.attr_start + attrIdx;
      critCheck.push(
        `'${cleanSheet}'!$${colLetter(cc)}$${cdrFirst}:$${colLetter(cc)}$${cdrLast},$${colLetter(filterStart + attrIdx)}${rCheck}`
      );
    }
    const critStr = critCheck.join(',');
    const sumRange = `'${cleanSheet}'!${colLetter(cleanCol)}$${cdrFirst}:${colLetter(cleanCol)}$${cdrLast}`;
    if (critStr) {
      ws.getCell(rCheck, dc).value = { formula: `${dcl}${rTotal}*${unitsCell}-SUMIFS(${sumRange},${critStr})` };
    } else {
      ws.getCell(rCheck, dc).value = { formula: `${dcl}${rTotal}*${unitsCell}-SUM(${sumRange})` };
    }

    const dc2 = s2Start + pi;
    const dc2l = colLetter(dc2);
    const countRange = `'${cleanSheet}'!${colLetter(cleanCol)}$${cdrFirst}:${colLetter(cleanCol)}$${cdrLast}`;
    if (critStr) {
      ws.getCell(rCheck, dc2).value = { formula: `${dc2l}${rTotal}-COUNTIFS(${countRange},"<>"&0,${critStr})` };
    } else {
      ws.getCell(rCheck, dc2).value = { formula: `${dc2l}${rTotal}-COUNTIF(${countRange},"<>"&0)` };
    }
  }

  // Row labels
  ws.getCell(rTotal, cohortLabelCol).value = 'Total';
  ws.getCell(rTotal, s3Label).value = 'Average';
  ws.getCell(rTotal, s4Label).value = 'Average';
  ws.getCell(rMedian, s3Label).value = 'Median';
  ws.getCell(rMedian, s4Label).value = 'Median';
  ws.getCell(rWeighted, s3Label).value = 'Dollar-Weighted Average';
  ws.getCell(rWeighted, s4Label).value = 'Size-Weighted Average';
  ws.getCell(rCheck, cohortLabelCol).value = 'Check';
}
