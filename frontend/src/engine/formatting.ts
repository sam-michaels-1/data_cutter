/**
 * Formatting module for all Data Pack tabs.
 * Applies fonts, number formats, alignment, borders, and conditional formatting.
 */
import type { Workbook, Worksheet, Style } from 'exceljs';
import type { FilterBlock } from './types';
import { colLetter } from './utils';

// Number formats
const NF_DOLLAR = '* _(* "$"\\ #,##0_);_(* "$"\\ \\(#,##0\\);* \\-_);* @_)';
const NF_DOLLAR_DEC = '* _(* "$"\\ #,##0.0_);_(* "$"\\ \\(#,##0.0\\);* \\-_);* @_)';
const NF_NUMBER = '* #,##0_);* \\(#,##0\\);* \\-_);* @_)';
const NF_NUMBER_DEC = '* #,##0.0_);* \\(#,##0.0\\);* \\-_);* @_)';
const NF_PCT = '* #,##0%_);* \\(#,##0%\\);* \\-_)';
const NF_PCT_DEC = '* #,##0.0%_);* \\(#,##0.0%\\);* \\-_)';
const NF_TIMES = '* #,##0.0\\x_);* \\(#,##0.0\\x\\);* \\-_);* @_)';
const NF_DATE = "mmm\\ \\'yy";

// Colors
const GREEN_COLOR = '00B050';
const PURPLE_COLOR = '7030A0';
const BLUE_COLOR = '0000FF';

const DIRECT_LINK_RE = /^=?\$?[A-Z]+\$?\d+$/i;

function baseFont(bold = false): Partial<Style['font']> {
  return { name: 'Times New Roman', size: 10, bold };
}

function setFont(ws: Worksheet, row: number, col: number, bold = false, color?: string) {
  const cell = ws.getCell(row, col);
  cell.font = { name: 'Times New Roman', size: 10, bold, color: color ? { argb: color } : undefined };
}

function setNumFmt(ws: Worksheet, row: number, col: number, fmt: string) {
  ws.getCell(row, col).numFmt = fmt;
}

function setAlign(ws: Worksheet, row: number, col: number, horizontal: 'left' | 'center' | 'right' | 'centerContinuous') {
  ws.getCell(row, col).alignment = { horizontal };
}

function formatUnitsCell(ws: Worksheet, row: number, labelCol: number, valueCol: number): void {
  setFont(ws, row, labelCol, true);
  const cell = ws.getCell(row, valueCol);
  cell.font = { name: 'Times New Roman', size: 10, italic: true, color: { argb: GREEN_COLOR } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } };
  cell.numFmt = '"$ "#,##0';
  cell.border = {
    top: { style: 'dotted' },
    bottom: { style: 'dotted' },
    left: { style: 'dotted' },
    right: { style: 'dotted' },
  };
}

export function formatControlTab(ws: Worksheet, checkTabs?: [string, string][]): void {
  // Set base font for all populated cells
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      cell.font = baseFont();
    });
  });

  // Bold labels in column B (rows 3-7)
  for (let r = 3; r <= 7; r++) {
    setFont(ws, r, 2, true);
  }

  // Blue font + yellow fill for input cells in column C
  for (let r = 3; r <= 7; r++) {
    const cell = ws.getCell(r, 3);
    cell.font = { name: 'Times New Roman', size: 10, color: { argb: BLUE_COLOR } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } };
    cell.alignment = { horizontal: 'center' };
  }

  // Row 8: Input Format (if present — only for cleaned table inputs)
  if (ws.getCell(8, 2).value) {
    setFont(ws, 8, 2, true);
  }
  if (ws.getCell(8, 3).value) {
    const cell8 = ws.getCell(8, 3);
    cell8.font = { name: 'Times New Roman', size: 10, color: { argb: BLUE_COLOR } };
    cell8.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } };
    cell8.alignment = { horizontal: 'center' };
  }

  // Column widths
  ws.getColumn('A').width = 8.73;
  ws.getColumn('B').width = 25.63;
  ws.getColumn('C').width = 15.63;

  // Check Summary section
  if (checkTabs && checkTabs.length > 0) {
    const R_HDR = 10;
    setFont(ws, R_HDR, 2, true);
    setFont(ws, R_HDR, 3, true);
    setAlign(ws, R_HDR, 3, 'center');

    for (let i = 0; i < checkTabs.length; i++) {
      const r = R_HDR + 1 + i;
      setFont(ws, r, 2, false);
      setNumFmt(ws, r, 3, NF_NUMBER);
      setAlign(ws, r, 3, 'center');
    }

    const rTotal = R_HDR + 1 + checkTabs.length;
    setFont(ws, rTotal, 2, true);
    setFont(ws, rTotal, 3, true);
    setNumFmt(ws, rTotal, 3, NF_NUMBER);
    setAlign(ws, rTotal, 3, 'center');

    // Conditional formatting
    const totalRef = `C${rTotal}`;
    ws.addConditionalFormatting({
      ref: totalRef,
      rules: [
        {
          type: 'cellIs', operator: 'equal', formulae: ['0'], priority: 1,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'C6EFCE' } } },
        },
        {
          type: 'cellIs', operator: 'notEqual' as any, formulae: ['0'], priority: 2,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC7CE' } } },
        },
      ],
    });
  }
}

export function formatCleanDataTab(ws: Worksheet, layout: import('./utils').CleanLayout, firstDataRow: number, lastDataRow: number, _granularity: string): void {
  const maxCol = layout.new_biz_end;

  // Base font
  for (let r = 1; r <= lastDataRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      ws.getCell(r, c).font = baseFont();
    }
  }

  // Row 1: bold
  for (let c = 1; c <= maxCol; c++) setFont(ws, 1, c, true);

  // Row 5: section headers bold
  for (let c = 1; c <= maxCol; c++) {
    if (ws.getCell(5, c).value) setFont(ws, 5, c, true);
  }

  // Row 6: bold, center
  for (let c = 1; c <= maxCol; c++) {
    if (ws.getCell(6, c).value) {
      setFont(ws, 6, c, true);
      setAlign(ws, 6, c, 'center');
    }
  }

  // Date format on row 6 ARR columns
  for (let c = layout.arr_start; c <= layout.arr_end; c++) {
    setNumFmt(ws, 6, c, NF_DATE);
  }

  // ARR data
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    for (let c = layout.arr_start; c <= layout.arr_end; c++) {
      setNumFmt(ws, r, c, NF_NUMBER);
    }
  }

  // Derived sections
  for (const section of ['churn', 'downsell', 'upsell', 'new_biz'] as const) {
    const s = layout[`${section}_start`];
    const e = layout[`${section}_end`];
    for (let r = firstDataRow; r <= lastDataRow; r++) {
      for (let c = s; c <= e; c++) {
        setNumFmt(ws, r, c, NF_NUMBER);
      }
    }
  }

  // Row 1 totals
  for (let c = layout.arr_start; c <= maxCol; c++) {
    setNumFmt(ws, 1, c, NF_NUMBER);
  }

  // Freeze panes
  ws.views = [{ state: 'frozen', xSplit: layout.arr_start - 1, ySplit: firstDataRow - 1, showGridLines: false }];
}

export function formatRetentionTab(
  ws: Worksheet, _config: import('./types').EngineConfig, filterBlocks: FilterBlock[],
  numDerived: number, numAttrs: number,
  s1Label: number, s1Start: number, s1End: number,
  s2Label: number, s2Start: number, s2End: number,
  s3Label: number, s3Start: number, s3End: number,
  filterStart: number, cohortFc: number
): void {
  const maxCol = s3End;
  const maxRow = 5 + filterBlocks.length * 19;

  // Base font
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      ws.getCell(r, c).font = baseFont();
    }
  }

  // Freeze panes
  ws.views = [{ state: 'frozen', xSplit: s1Start - 1, ySplit: 7, showGridLines: false }];

  // Units cell
  formatUnitsCell(ws, 3, 1, filterStart);

  for (let blockIdx = 0; blockIdx < filterBlocks.length; blockIdx++) {
    const start = 5 + blockIdx * 19;

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

    // Title bold
    setFont(ws, rTitle, s1Label, true);

    // Section headers
    for (const col of [s1Label, s2Label, s3Label]) {
      setFont(ws, rSections, col, true);
      setAlign(ws, rSections, col, 'centerContinuous');
    }

    // Column headers
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      setFont(ws, rHeader, filterStart + attrIdx, true);
      setAlign(ws, rHeader, filterStart + attrIdx, 'center');
    }
    setFont(ws, rHeader, cohortFc, true);
    setAlign(ws, rHeader, cohortFc, 'center');

    // Date headers
    for (let i = 0; i < numDerived; i++) {
      for (const sStart of [s1Start, s2Start, s3Start]) {
        setFont(ws, rHeader, sStart + i, true);
        setAlign(ws, rHeader, sStart + i, 'center');
        setNumFmt(ws, rHeader, sStart + i, NF_DATE);
      }
    }

    // Filter values center + blue on first row
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      const fc = filterStart + attrIdx;
      setFont(ws, rBop, fc, true, BLUE_COLOR);
      setAlign(ws, rBop, fc, 'center');
      for (let r = rChurn; r <= rNlGrowth; r++) {
        setAlign(ws, r, fc, 'center');
      }
    }
    setFont(ws, rBop, cohortFc, true, BLUE_COLOR);
    setAlign(ws, rBop, cohortFc, 'center');

    // Section 1 formatting
    for (const r of [rBop, rRetained, rEop, rNetRet]) {
      setFont(ws, r, s1Label, true);
      for (let c = s1Start; c <= s1End; c++) setFont(ws, r, c, true);
    }

    for (let i = 0; i < numDerived; i++) {
      const c = s1Start + i;
      for (const r of [rBop, rRetained, rEop]) setNumFmt(ws, r, c, NF_DOLLAR);
      for (const r of [rChurn, rDownsell, rUpsell, rNewLogo, rCheck]) setNumFmt(ws, r, c, NF_NUMBER);
      setNumFmt(ws, rGrowth, c, NF_PCT_DEC);
      for (const r of [rLostRet, rPunitRet, rNetRet, rNlPct, rNlGrowth]) setNumFmt(ws, r, c, NF_PCT_DEC);
    }

    // Section 2 formatting
    for (const r of [rBop, rDownsell, rRetained]) {
      setFont(ws, r, s2Label, true);
      for (let c = s2Start; c <= s2End; c++) setFont(ws, r, c, true);
    }

    for (let i = 0; i < numDerived; i++) {
      const c = s2Start + i;
      for (const r of [rBop, rChurn, rDownsell, rUpsell, rRetained, rCheck]) setNumFmt(ws, r, c, NF_NUMBER);
      setNumFmt(ws, rNewLogo, c, NF_PCT);
      setNumFmt(ws, rLostRet, c, NF_PCT);
      setNumFmt(ws, rPunitRet, c, NF_PCT);
    }

    // Section 3 formatting
    for (const r of [rBop, rUpsell, rNewLogo]) {
      setFont(ws, r, s3Label, true);
      for (let c = s3Start; c <= s3End; c++) setFont(ws, r, c, true);
    }

    for (let i = 0; i < numDerived; i++) {
      const c = s3Start + i;
      for (const r of [rBop, rUpsell, rNewLogo]) setNumFmt(ws, r, c, NF_DOLLAR_DEC);
      for (const r of [rChurn, rDownsell, rRetained]) setNumFmt(ws, r, c, NF_NUMBER_DEC);
      setNumFmt(ws, rEop, c, NF_PCT);
      setNumFmt(ws, rLostRet, c, NF_TIMES);
      setNumFmt(ws, rPunitRet, c, NF_TIMES);

      for (let r = rBop; r <= rNlGrowth; r++) {
        setAlign(ws, r, c, 'right');
      }
    }

    // Indent sub-item labels in section 1
    for (const r of [rChurn, rDownsell, rUpsell, rNewLogo, rNetRet]) {
      const cell = ws.getCell(r, s1Label);
      cell.alignment = { ...cell.alignment, indent: 1 };
    }

    // Indent sub-item labels in section 2 (Churned, New Logo, % Growth)
    for (const r of [rChurn, rUpsell, rNewLogo]) {
      const cell = ws.getCell(r, s2Label);
      cell.alignment = { ...cell.alignment, indent: 1 };
    }

    // Indent sub-item labels in section 3 (Churned, Upsell/Cross-sell, New Logo, % Growth)
    for (const r of [rChurn, rDownsell, rRetained, rEop]) {
      const cell = ws.getCell(r, s3Label);
      cell.alignment = { ...cell.alignment, indent: 1 };
    }

    // Italicize the 5 percentage rows across all section 1 columns
    for (const r of [rLostRet, rPunitRet, rNetRet, rNlPct, rNlGrowth]) {
      for (let c = 1; c <= maxCol; c++) {
        const cell = ws.getCell(r, c);
        cell.font = { ...cell.font, italic: true };
      }
    }

    // Italicize % Growth row in section 2 (label + data)
    {
      const cell = ws.getCell(rNewLogo, s2Label);
      cell.font = { ...cell.font, italic: true };
      for (let c = s2Start; c <= s2End; c++) {
        const dc = ws.getCell(rNewLogo, c);
        dc.font = { ...dc.font, italic: true };
      }
    }

    // Italicize % Growth ARR/Cust. row in section 3 (label + data)
    {
      const cell = ws.getCell(rEop, s3Label);
      cell.font = { ...cell.font, italic: true };
      for (let c = s3Start; c <= s3End; c++) {
        const dc = ws.getCell(rEop, c);
        dc.font = { ...dc.font, italic: true };
      }
    }
  }
}

export function formatCohortTab(
  ws: Worksheet, _config: import('./types').EngineConfig, filterBlocks: FilterBlock[],
  numDates: number, numCohorts: number, numAttrs: number,
  qCol: number, yCol: number, filterStart: number, cohortLabelCol: number,
  s1Start: number, _s1End: number, s2Start: number, _s2End: number,
  s3Label: number, s3StartVal: number, s3DataStart: number, _s3DataEnd: number,
  s4Label: number, s4StartVal: number, s4DataStart: number, s4DataEnd: number,
  _granularity: string
): void {
  const maxCol = s4DataEnd;
  const s3DataEnd = s3DataStart + numDates - 1;

  // Units cell
  formatUnitsCell(ws, 3, 1, qCol);

  for (let blockIdx = 0; blockIdx < filterBlocks.length; blockIdx++) {
    const blockStart = 6 + blockIdx * (numCohorts + 9);
    const rSectionHeaders = blockStart + 2;
    const rHeaders = blockStart + 3;
    const firstCohortRow = rHeaders + 1;
    const lastCohortRow = firstCohortRow + numCohorts - 1;
    const rTotal = lastCohortRow + 1;
    const rMedian = rTotal + 1;
    const rWeighted = rMedian + 1;
    const rCheck = rWeighted + 1;

    // Title bold
    setFont(ws, blockStart, qCol, true);

    // Section headers
    for (const col of [s1Start, s2Start, s3Label, s4Label]) {
      setFont(ws, rSectionHeaders, col, true);
      setAlign(ws, rSectionHeaders, col, 'centerContinuous');
    }

    // Column headers
    for (const col of [qCol, yCol]) {
      setFont(ws, rHeaders, col, true);
      setAlign(ws, rHeaders, col, 'center');
    }
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      setFont(ws, rHeaders, filterStart + attrIdx, true);
      setAlign(ws, rHeaders, filterStart + attrIdx, 'center');
    }
    setFont(ws, rHeaders, cohortLabelCol, true);
    setAlign(ws, rHeaders, cohortLabelCol, 'center');

    for (let i = 0; i < numDates; i++) {
      setFont(ws, rHeaders, s1Start + i, true);
      setAlign(ws, rHeaders, s1Start + i, 'center');
      setFont(ws, rHeaders, s2Start + i, true);
      setAlign(ws, rHeaders, s2Start + i, 'center');
    }

    setFont(ws, rHeaders, s3StartVal, true);
    setAlign(ws, rHeaders, s3StartVal, 'center');
    setFont(ws, rHeaders, s4StartVal, true);
    setAlign(ws, rHeaders, s4StartVal, 'center');
    for (let i = 0; i < numDates; i++) {
      setFont(ws, rHeaders, s3DataStart + i, true);
      setAlign(ws, rHeaders, s3DataStart + i, 'center');
      setFont(ws, rHeaders, s4DataStart + i, true);
      setAlign(ws, rHeaders, s4DataStart + i, 'center');
    }

    // Blue font on first cohort row filter values
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      setFont(ws, firstCohortRow, filterStart + attrIdx, true, BLUE_COLOR);
    }
    setFont(ws, firstCohortRow, cohortLabelCol, false, BLUE_COLOR);

    // Data rows
    for (let r = firstCohortRow; r <= rCheck; r++) {
      setAlign(ws, r, qCol, 'center');
      setAlign(ws, r, yCol, 'center');
      for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
        setAlign(ws, r, filterStart + attrIdx, 'center');
      }
      setAlign(ws, r, cohortLabelCol, 'center');

      for (let i = 0; i < numDates; i++) {
        setNumFmt(ws, r, s1Start + i, NF_NUMBER);
        setNumFmt(ws, r, s2Start + i, NF_NUMBER);
      }

      setNumFmt(ws, r, s3StartVal, NF_DOLLAR);
      setNumFmt(ws, r, s4StartVal, NF_NUMBER);

      for (let i = 0; i < numDates; i++) {
        setNumFmt(ws, r, s3DataStart + i, NF_PCT);
        setNumFmt(ws, r, s4DataStart + i, NF_PCT);
      }
    }

    // Total row: bold only (NOT italic), $ format for ARR section
    for (let c = 1; c <= maxCol; c++) {
      if (ws.getCell(rTotal, c).value != null) {
        setFont(ws, rTotal, c, true);
      }
    }
    for (let i = 0; i < numDates; i++) {
      setNumFmt(ws, rTotal, s1Start + i, NF_DOLLAR);
    }

    // Average/Median/Weighted: bold + italic across retention data columns
    for (const r of [rTotal, rMedian, rWeighted]) {
      // Italic on retention section labels and data (s3 and s4)
      for (let c = s3Label; c <= s4DataEnd; c++) {
        const cell = ws.getCell(r, c);
        if (cell.value != null) {
          cell.font = { ...baseFont(true), italic: true, color: cell.font?.color };
        }
      }
    }
    // Median and Weighted rows: bold + italic across ALL columns
    for (const r of [rMedian, rWeighted]) {
      for (let c = 1; c <= maxCol; c++) {
        if (ws.getCell(r, c).value != null) {
          const cell = ws.getCell(r, c);
          cell.font = { ...baseFont(true), italic: true, color: cell.font?.color };
        }
      }
    }

    // Bold "Cohort" label in retention sections
    setFont(ws, rHeaders, s3Label, true);
    setAlign(ws, rHeaders, s3Label, 'center');
    setFont(ws, rHeaders, s4Label, true);
    setAlign(ws, rHeaders, s4Label, 'center');

    // Left-align summary labels in retention sections
    for (const r of [rTotal, rMedian, rWeighted]) {
      ws.getCell(r, s3Label).alignment = { horizontal: 'left' };
      ws.getCell(r, s4Label).alignment = { horizontal: 'left' };
    }

    // Borders: underline under section headers and column headers, top border above Total
    const thinBorder = { style: 'thin' as const };
    const s1End = s1Start + numDates - 1;
    const s2End = s2Start + numDates - 1;
    const sectionRanges: [number, number][] = [
      [s1Start, s1End],
      [s2Start, s2End],
      [s3Label, s3DataEnd],
      [s4Label, s4DataEnd],
    ];
    for (const [startC, endC] of sectionRanges) {
      // Bottom border under section headers row
      for (let c = startC; c <= endC; c++) {
        ws.getCell(rSectionHeaders, c).border = { ...ws.getCell(rSectionHeaders, c).border, bottom: thinBorder };
      }
      // Bottom border under column headers row
      for (let c = startC; c <= endC; c++) {
        ws.getCell(rHeaders, c).border = { ...ws.getCell(rHeaders, c).border, bottom: thinBorder };
      }
      // Top border above Total row
      for (let c = startC; c <= endC; c++) {
        ws.getCell(rTotal, c).border = { ...ws.getCell(rTotal, c).border, top: thinBorder };
      }
    }

    // Conditional formatting — ARR Retention: 3-color (red min, white at 1.0, green max)
    const s3TopLeft = `${colLetter(s3DataStart)}${firstCohortRow}`;
    const s3BotRight = `${colLetter(s3DataEnd)}${lastCohortRow}`;
    ws.addConditionalFormatting({
      ref: `${s3TopLeft}:${s3BotRight}`,
      rules: [{
        type: 'colorScale',
        priority: 1,
        cfvo: [
          { type: 'min' },
          { type: 'num', value: 1.0 },
          { type: 'max' },
        ],
        color: [
          { argb: 'FFF8696B' },
          { argb: 'FFFFFFFF' },
          { argb: 'FF63BE7B' },
        ],
      } as any],
    });

    // Conditional formatting — Logo Retention: 2-color (yellow min, green max)
    const s4TopLeft = `${colLetter(s4DataStart)}${firstCohortRow}`;
    const s4BotRight = `${colLetter(s4DataEnd)}${lastCohortRow}`;
    ws.addConditionalFormatting({
      ref: `${s4TopLeft}:${s4BotRight}`,
      rules: [{
        type: 'colorScale',
        priority: 1,
        cfvo: [
          { type: 'min' },
          { type: 'max' },
        ],
        color: [
          { argb: 'FFFFEB84' },
          { argb: 'FF63BE7B' },
        ],
      } as any],
    });
  }

  // Base font pass
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (!cell.font || cell.font.name !== 'Times New Roman') {
        cell.font = { ...baseFont(), ...cell.font, name: 'Times New Roman', size: 10 };
      }
    });
  });
}

export function formatTopCustomersTab(
  ws: Worksheet, _config: import('./types').EngineConfig, _layout: import('./utils').CleanLayout,
  firstCustomerRow: number, lastCustomerRow: number,
  rTopTotal: number, rOther: number, rTotal: number, rMemoStart: number,
  numDates: number,
  rankNumCol: number, custIdCol: number, attrStart: number, numAttrs: number, cohortCol: number,
  s1Start: number, _s1End: number, s2Start: number, _s2End: number, s3Start: number, s3End: number
): void {
  const maxCol = s3End;

  // Base font
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      cell.font = baseFont();
    });
  });

  // Units cell
  formatUnitsCell(ws, 3, 1, rankNumCol);

  // Row 5: Section headers
  setFont(ws, 5, s1Start, true); setAlign(ws, 5, s1Start, 'centerContinuous');
  setFont(ws, 5, s2Start, true); setAlign(ws, 5, s2Start, 'centerContinuous');
  setFont(ws, 5, s3Start, true); setAlign(ws, 5, s3Start, 'centerContinuous');

  // Row 6: Column headers
  for (const c of [custIdCol, ...Array.from({ length: numAttrs }, (_, i) => attrStart + i), cohortCol]) {
    setFont(ws, 6, c, true);
    setAlign(ws, 6, c, 'center');
  }
  for (let i = 0; i < numDates; i++) {
    setFont(ws, 6, s1Start + i, true); setAlign(ws, 6, s1Start + i, 'center');
  }
  for (let i = 0; i < numDates - 1; i++) {
    setFont(ws, 6, s2Start + i, true); setAlign(ws, 6, s2Start + i, 'center');
  }
  for (let i = 0; i < numDates; i++) {
    setFont(ws, 6, s3Start + i, true); setAlign(ws, 6, s3Start + i, 'center');
  }

  // Data rows
  for (let r = firstCustomerRow; r <= lastCustomerRow; r++) {
    setAlign(ws, r, rankNumCol, 'center');
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) setAlign(ws, r, attrStart + attrIdx, 'center');
    setAlign(ws, r, cohortCol, 'center');

    for (let i = 0; i < numDates; i++) setNumFmt(ws, r, s1Start + i, NF_NUMBER);
    for (let i = 0; i < numDates - 1; i++) {
      setNumFmt(ws, r, s2Start + i, NF_PCT);
      setAlign(ws, r, s2Start + i, 'right');
    }
    for (let i = 0; i < numDates; i++) {
      setNumFmt(ws, r, s3Start + i, NF_PCT_DEC);
      setAlign(ws, r, s3Start + i, 'right');
    }
  }

  // Summary rows bold
  for (const r of [rTopTotal, rTotal]) {
    for (let c = 1; c <= maxCol; c++) {
      if (ws.getCell(r, c).value != null) setFont(ws, r, c, true);
    }
    for (let i = 0; i < numDates; i++) setNumFmt(ws, r, s1Start + i, NF_NUMBER);
    for (let i = 0; i < numDates - 1; i++) setNumFmt(ws, r, s2Start + i, NF_PCT);
    for (let i = 0; i < numDates; i++) setNumFmt(ws, r, s3Start + i, NF_PCT_DEC);
  }

  // Other row
  for (let i = 0; i < numDates; i++) setNumFmt(ws, rOther, s1Start + i, NF_NUMBER);
  for (let i = 0; i < numDates - 1; i++) setNumFmt(ws, rOther, s2Start + i, NF_PCT);
  for (let i = 0; i < numDates; i++) setNumFmt(ws, rOther, s3Start + i, NF_PCT_DEC);

  // Memo rows
  for (let tierIdx = 0; tierIdx < 4; tierIdx++) {
    const r = rMemoStart + 1 + tierIdx;
    for (let i = 0; i < numDates; i++) setNumFmt(ws, r, s1Start + i, NF_NUMBER);
    for (let i = 0; i < numDates - 1; i++) setNumFmt(ws, r, s2Start + i, NF_PCT);
    for (let i = 0; i < numDates; i++) setNumFmt(ws, r, s3Start + i, NF_PCT_DEC);
  }
}

/**
 * Apply formula auditing color-coding to all sheets.
 */
export function applyFormulaColoring(wb: Workbook, skipSheets?: string[]): void {
  const skip = new Set(skipSheets || []);

  for (const ws of wb.worksheets) {
    if (skip.has(ws.name)) continue;

    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.value == null) return;

        const oldFont = cell.font || {};
        const fname = oldFont.name || 'Times New Roman';
        const fsize = oldFont.size || 10;
        const bold = oldFont.bold || false;
        const italic = oldFont.italic || false;

        // Check if it's a formula
        const val = cell.value;
        if (typeof val === 'object' && val !== null && 'formula' in val) {
          const formula = (val as { formula: string }).formula;
          if (formula.includes('!')) {
            // Cross-sheet reference → green
            cell.font = { name: fname, size: fsize, bold, italic, color: { argb: GREEN_COLOR } };
          } else if (DIRECT_LINK_RE.test(formula)) {
            // Direct cell link → purple
            cell.font = { name: fname, size: fsize, bold, italic, color: { argb: PURPLE_COLOR } };
          }
        } else if (typeof val === 'number') {
          // Hardcoded numeric → blue
          cell.font = { name: fname, size: fsize, bold, italic, color: { argb: BLUE_COLOR } };
        }
      });
    });
  }
}

/**
 * Remove gridlines from every worksheet.
 */
export function removeGridlines(wb: Workbook): void {
  for (const ws of wb.worksheets) {
    if (ws.views && ws.views.length > 0) {
      ws.views = ws.views.map(v => ({ ...v, showGridLines: false }));
    } else {
      ws.views = [{ showGridLines: false }];
    }
  }
}

/**
 * Apply pastel tab colors by section.
 */
export function applyTabColors(wb: Workbook): void {
  const colorMap: [RegExp | string, string][] = [
    ['Control', '92D050'],
    [/Retention/, '5B9BD5'],
    [/Cohort/, 'FFD966'],
    [/Top Customer/, 'F4B084'],
    [/^Clean/, 'C0C0C0'],
  ];

  for (const ws of wb.worksheets) {
    for (const [pattern, color] of colorMap) {
      const match = typeof pattern === 'string'
        ? ws.name === pattern
        : pattern.test(ws.name);
      if (match) {
        (ws as any).properties = (ws as any).properties || {};
        (ws as any).properties.tabColor = { argb: `FF${color}` };
        break;
      }
    }
  }
}
