/**
 * Retention tab generators.
 * Port of data-pack-app/engine/retention.py for client-side use with ExcelJS.
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig, FilterBlock } from './types';
import type { CleanLayout } from './utils';
import { colLetter, getYoyOffset } from './utils';

const BLOCK_HEIGHT = 19;

export function generateRetentionTab(
  wb: Workbook, config: EngineConfig,
  cleanSheetName: string, cleanLayout: CleanLayout,
  firstDataRow: number, lastDataRow: number,
  granularity: string, filterBlocks: FilterBlock[]
): string {
  const sheetName = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Retention`;
  const ws = wb.addWorksheet(sheetName);

  const numDerived = cleanLayout.num_derived;
  const numAttrs = cleanLayout.num_attrs;
  const attrNames = Object.keys(config.attributes);
  const yoyOffset = cleanLayout.yoy_offset;

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

  // --- Row 1: Check summary ---
  ws.getCell(1, 1).value = 'Check Summary';
  const checkRefs: string[] = [];
  for (let blockIdx = 0; blockIdx < filterBlocks.length; blockIdx++) {
    const blockStart = 5 + blockIdx * BLOCK_HEIGHT;
    const checkRow = blockStart + 11;
    checkRefs.push(`${colLetter(s1Start)}${checkRow}:${colLetter(s1End)}${checkRow}`);
    checkRefs.push(`${colLetter(s2Start)}${checkRow}:${colLetter(s2End)}${checkRow}`);
  }
  ws.getCell(1, 2).value = { formula: `SUM(${checkRefs.join(',')})` };

  // --- Row 3: Units ---
  ws.getCell(3, 1).value = 'Units';
  ws.getCell(3, filterStartCol).value = { formula: 'Control!$C$4' };
  const unitsCell = `$${colLetter(filterStartCol)}$3`;

  const metricLabel = (config.data_type || 'arr') === 'arr' ? 'ARR' : 'Revenue';
  const cdrFirst = firstDataRow;
  const cdrLast = lastDataRow;

  // --- Generate each block ---
  for (let blockIdx = 0; blockIdx < filterBlocks.length; blockIdx++) {
    const block = filterBlocks[blockIdx];
    const blockStart = 5 + blockIdx * BLOCK_HEIGHT;

    writeRetentionBlock(
      ws, blockStart, block, config,
      cleanSheetName, cleanLayout, cdrFirst, cdrLast,
      s1Label, s1Start, s1End, s2Label, s2Start, s2End,
      s3Label, s3Start, s3End, filterStartCol, cohortFc,
      numDerived, numAttrs, attrNames, unitsCell, yoyOffset,
      granularity, metricLabel
    );
  }

  return sheetName;
}

function writeRetentionBlock(
  ws: import('exceljs').Worksheet, start: number, block: FilterBlock,
  config: EngineConfig,
  cleanSheet: string, cleanLayout: CleanLayout,
  cdrFirst: number, cdrLast: number,
  s1Label: number, s1Start: number, s1End: number,
  s2Label: number, s2Start: number, s2End: number,
  s3Label: number, s3Start: number, s3End: number,
  filterStart: number, cohortFc: number,
  numDerived: number, numAttrs: number, attrNames: string[],
  unitsCell: string, yoyOffset: number, granularity: string,
  metricLabel: string
): void {
  const { title, filters } = block;

  const rTitle = start;
  const rSections = start + 1;
  const rHeader = start + 2;
  const rBop = start + 3;
  const rChurn = start + 4;
  const rDownsell = start + 5;
  const rUpsell = start + 6;
  const rRetained = start + 7;
  const rNewLogo = start + 8;
  const rEop = start + 9;
  const rGrowth = start + 10;
  const rCheck = start + 11;
  const rLostRet = start + 13;
  const rPunitRet = start + 14;
  const rNetRet = start + 15;
  const rNlPct = start + 16;
  const rNlGrowth = start + 17;

  // Title
  ws.getCell(rTitle, s1Label).value = `${title} ${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Retention Analysis`;

  // Section sub-headers
  ws.getCell(rSections, s1Label).value = 'Net Retention Analysis';
  ws.getCell(rSections, s2Label).value = 'Customer Retention Analysis';
  ws.getCell(rSections, s3Label).value = `${metricLabel} / Customer`;

  // Column headers
  for (let attrIdx = 0; attrIdx < attrNames.length; attrIdx++) {
    ws.getCell(rHeader, filterStart + attrIdx).value = attrNames[attrIdx];
  }
  ws.getCell(rHeader, cohortFc).value = 'Cohort';

  // Date headers from clean data
  for (let i = 0; i < numDerived; i++) {
    const churnCol = cleanLayout.churn_start + i;
    const churnCl = colLetter(churnCol);
    ws.getCell(rHeader, s1Start + i).value = { formula: `'${cleanSheet}'!${churnCl}$6` };
    ws.getCell(rHeader, s2Start + i).value = { formula: `${colLetter(s1Start + i)}${rHeader}` };
    ws.getCell(rHeader, s3Start + i).value = { formula: `${colLetter(s2Start + i)}${rHeader}` };
  }

  // Filter values
  const allDataRows = [rBop, rChurn, rDownsell, rUpsell, rRetained, rNewLogo, rEop, rGrowth, rCheck, rLostRet, rPunitRet, rNetRet, rNlPct, rNlGrowth];
  for (let attrIdx = 0; attrIdx < attrNames.length; attrIdx++) {
    const fc = filterStart + attrIdx;
    ws.getCell(rBop, fc).value = filters[attrNames[attrIdx]] || '<>';
    for (const r of allDataRows.slice(1)) {
      ws.getCell(r, fc).value = { formula: `${colLetter(fc)}${rBop}` };
    }
  }
  ws.getCell(rBop, cohortFc).value = filters['Cohort'] || '<>';
  for (const r of allDataRows.slice(1)) {
    ws.getCell(r, cohortFc).value = { formula: `${colLetter(cohortFc)}${rBop}` };
  }

  // Helper: build SUMIFS criteria
  function criteria(row: number): string {
    const parts: string[] = [];
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      const cc = cleanLayout.attr_start + attrIdx;
      parts.push(
        `'${cleanSheet}'!$${colLetter(cc)}$${cdrFirst}:$${colLetter(cc)}$${cdrLast},$${colLetter(filterStart + attrIdx)}${row}`
      );
    }
    const cc = cleanLayout.cohort;
    parts.push(
      `'${cleanSheet}'!$${colLetter(cc)}$${cdrFirst}:$${colLetter(cc)}$${cdrLast},$${colLetter(cohortFc)}${row}`
    );
    return parts.join(',');
  }

  function buildSumifs(cleanColNum: number, row: number): string {
    const cl = colLetter(cleanColNum);
    const rng = `'${cleanSheet}'!${cl}$${cdrFirst}:${cl}$${cdrLast}`;
    return `SUMIFS(${rng},${criteria(row)})/${unitsCell}`;
  }

  function buildCountifsNonzero(cleanColNum: number, row: number): string {
    const cl = colLetter(cleanColNum);
    const rng = `'${cleanSheet}'!${cl}$${cdrFirst}:${cl}$${cdrLast}`;
    return `COUNTIFS(${rng},"<>"&0,${criteria(row)})`;
  }

  // ===== SECTION 1: Net Retention =====
  ws.getCell(rBop, s1Label).value = `BoP ${metricLabel}`;
  ws.getCell(rChurn, s1Label).value = '(-) Churn';
  ws.getCell(rDownsell, s1Label).value = '(-) Downsell';
  ws.getCell(rUpsell, s1Label).value = '(+) Upsell / Cross-sell';
  ws.getCell(rRetained, s1Label).value = `Retained ${metricLabel}`;
  ws.getCell(rNewLogo, s1Label).value = '(+) New Logo';
  ws.getCell(rEop, s1Label).value = `EoP ${metricLabel}`;
  ws.getCell(rGrowth, s1Label).value = '% Growth';
  ws.getCell(rCheck, s1Label).value = 'Check';
  ws.getCell(rLostRet, s1Label).value = '% Lost-Only Retention';
  ws.getCell(rPunitRet, s1Label).value = '% Punitive Retention';
  ws.getCell(rNetRet, s1Label).value = '% Net Retention';
  ws.getCell(rNlPct, s1Label).value = '% New Logo % of BoP';
  ws.getCell(rNlGrowth, s1Label).value = '% New Logo Growth';

  for (let i = 0; i < numDerived; i++) {
    const dc = s1Start + i;
    const dcl = colLetter(dc);

    ws.getCell(rBop, dc).value = { formula: buildSumifs(cleanLayout.arr_start + i, rBop) };
    ws.getCell(rChurn, dc).value = { formula: buildSumifs(cleanLayout.churn_start + i, rChurn) };
    ws.getCell(rDownsell, dc).value = { formula: buildSumifs(cleanLayout.downsell_start + i, rDownsell) };
    ws.getCell(rUpsell, dc).value = { formula: buildSumifs(cleanLayout.upsell_start + i, rUpsell) };
    ws.getCell(rRetained, dc).value = { formula: `SUM(${dcl}${rBop}:${dcl}${rUpsell})` };
    ws.getCell(rNewLogo, dc).value = { formula: buildSumifs(cleanLayout.new_biz_start + i, rNewLogo) };
    ws.getCell(rEop, dc).value = { formula: `SUM(${dcl}${rRetained}:${dcl}${rNewLogo})` };
    ws.getCell(rGrowth, dc).value = { formula: `${dcl}${rEop}/${dcl}${rBop}-1` };

    // Check
    const eopCol = cleanLayout.arr_start + yoyOffset + i;
    const eopCl = colLetter(eopCol);
    const eopRng = `'${cleanSheet}'!${eopCl}$${cdrFirst}:${eopCl}$${cdrLast}`;
    ws.getCell(rCheck, dc).value = { formula: `${dcl}${rEop}*${unitsCell}-SUMIFS(${eopRng},${criteria(rCheck)})` };

    ws.getCell(rLostRet, dc).value = { formula: `SUM(${dcl}${rBop}:${dcl}${rChurn})/${dcl}${rBop}` };
    ws.getCell(rPunitRet, dc).value = { formula: `SUM(${dcl}${rBop}:${dcl}${rDownsell})/${dcl}${rBop}` };
    ws.getCell(rNetRet, dc).value = { formula: `SUM(${dcl}${rBop}:${dcl}${rUpsell})/${dcl}${rBop}` };
    ws.getCell(rNlPct, dc).value = { formula: `${dcl}${rNewLogo}/${dcl}${rBop}` };

    if (i >= yoyOffset) {
      const priorCl = colLetter(s1Start + i - yoyOffset);
      ws.getCell(rNlGrowth, dc).value = { formula: `${dcl}${rNewLogo}/${priorCl}${rNewLogo}-1` };
    }
  }

  // ===== SECTION 2: Customer Retention =====
  ws.getCell(rBop, s2Label).value = 'BoP Customers';
  ws.getCell(rChurn, s2Label).value = '(-) Churned Customers';
  ws.getCell(rDownsell, s2Label).value = 'Retained Customers';
  ws.getCell(rUpsell, s2Label).value = '(+) New Logo';
  ws.getCell(rRetained, s2Label).value = 'EoP Customers';
  ws.getCell(rNewLogo, s2Label).value = '% Growth';
  ws.getCell(rCheck, s2Label).value = 'Check';
  ws.getCell(rLostRet, s2Label).value = 'Logo Retention';
  ws.getCell(rPunitRet, s2Label).value = 'New Logo % of BoP';

  for (let i = 0; i < numDerived; i++) {
    const dc = s2Start + i;
    const dcl = colLetter(dc);

    ws.getCell(rBop, dc).value = { formula: buildCountifsNonzero(cleanLayout.arr_start + i, rBop) };

    const churnCol = cleanLayout.churn_start + i;
    const cl = colLetter(churnCol);
    const rng = `'${cleanSheet}'!${cl}$${cdrFirst}:${cl}$${cdrLast}`;
    ws.getCell(rChurn, dc).value = { formula: `-(COUNTIFS(${rng},"<>"&0,${criteria(rChurn)}))` };

    ws.getCell(rDownsell, dc).value = { formula: `SUM(${dcl}${rBop}:${dcl}${rChurn})` };

    ws.getCell(rUpsell, dc).value = { formula: buildCountifsNonzero(cleanLayout.new_biz_start + i, rUpsell) };
    ws.getCell(rRetained, dc).value = { formula: `SUM(${dcl}${rDownsell}:${dcl}${rUpsell})` };
    ws.getCell(rNewLogo, dc).value = { formula: `${dcl}${rRetained}/${dcl}${rBop}-1` };

    const eopCol = cleanLayout.arr_start + yoyOffset + i;
    const eopCl = colLetter(eopCol);
    const eopRng = `'${cleanSheet}'!${eopCl}$${cdrFirst}:${eopCl}$${cdrLast}`;
    ws.getCell(rCheck, dc).value = { formula: `${dcl}${rRetained}-COUNTIFS(${eopRng},"<>"&0,${criteria(rCheck)})` };

    ws.getCell(rLostRet, dc).value = { formula: `${dcl}${rDownsell}/${dcl}${rBop}` };
    ws.getCell(rPunitRet, dc).value = { formula: `${dcl}${rUpsell}/${dcl}${rBop}` };
  }

  // ===== SECTION 3: ARR / Customer =====
  ws.getCell(rBop, s3Label).value = 'BoP Customers';
  ws.getCell(rChurn, s3Label).value = '(-) Churned Customers';
  ws.getCell(rDownsell, s3Label).value = '(+/-) Upsell / Cross-sell';
  ws.getCell(rUpsell, s3Label).value = 'Retained Customers';
  ws.getCell(rRetained, s3Label).value = '(+) New Logo';
  ws.getCell(rNewLogo, s3Label).value = 'EoP Customers';
  ws.getCell(rEop, s3Label).value = `% Growth ${metricLabel}/Cust.`;
  ws.getCell(rLostRet, s3Label).value = 'New Logo vs Churn';
  ws.getCell(rPunitRet, s3Label).value = 'New Logo vs Retained';

  for (let i = 0; i < numDerived; i++) {
    const dc = s3Start + i;
    const dcl = colLetter(dc);
    const s1Dcl = colLetter(s1Start + i);
    const s2Dcl = colLetter(s2Start + i);

    ws.getCell(rBop, dc).value = { formula: `IFERROR(${s1Dcl}${rBop}/${s2Dcl}${rBop},"n.a.")` };
    ws.getCell(rChurn, dc).value = { formula: `IFERROR(-${s1Dcl}${rChurn}/${s2Dcl}${rChurn},"n.a.")` };
    ws.getCell(rDownsell, dc).value = { formula: `IFERROR(${dcl}${rUpsell}-SUM(${dcl}${rBop}:${dcl}${rChurn}),"n.a.")` };
    ws.getCell(rUpsell, dc).value = { formula: `IFERROR(${s1Dcl}${rRetained}/${s2Dcl}${rDownsell},"n.a.")` };
    ws.getCell(rRetained, dc).value = { formula: `IFERROR(${s1Dcl}${rNewLogo}/${s2Dcl}${rUpsell},"n.a.")` };
    ws.getCell(rNewLogo, dc).value = { formula: `IFERROR(${s1Dcl}${rEop}/${s2Dcl}${rRetained},"n.a.")` };
    ws.getCell(rEop, dc).value = { formula: `IFERROR(${dcl}${rNewLogo}/${dcl}${rBop}-1,"n.a.")` };
    ws.getCell(rLostRet, dc).value = { formula: `IFERROR(-${dcl}${rRetained}/${dcl}${rChurn},"n.a.")` };
    ws.getCell(rPunitRet, dc).value = { formula: `IFERROR(${dcl}${rRetained}/${dcl}${rUpsell},"n.a.")` };
  }
}
