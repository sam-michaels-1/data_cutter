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

const THIN_BORDER = { style: 'thin' as const };

/** Center-across-selection across [c1, c2] with an underline spanning the selection. */
function centerAcrossUnderline(ws: Worksheet, row: number, c1: number, c2: number): void {
  for (let c = c1; c <= c2; c++) {
    const cell = ws.getCell(row, c);
    cell.alignment = { ...cell.alignment, horizontal: 'centerContinuous' };
    cell.border = { ...cell.border, bottom: THIN_BORDER };
  }
}

/** Thin underline (bottom border) across [c1, c2]. */
function underlineSpan(ws: Worksheet, row: number, c1: number, c2: number): void {
  for (let c = c1; c <= c2; c++) {
    const cell = ws.getCell(row, c);
    cell.border = { ...cell.border, bottom: THIN_BORDER };
  }
}

/** Thin overline (top border) across [c1, c2] — underlines the numbers feeding a sum row. */
function overlineSpan(ws: Worksheet, row: number, c1: number, c2: number): void {
  for (let c = c1; c <= c2; c++) {
    const cell = ws.getCell(row, c);
    cell.border = { ...cell.border, top: THIN_BORDER };
  }
}

// Ranges where formula color-coding should not repaint the font — e.g. cells
// shaded by color-scale conditional formatting, where colored text is hard to
// read (the summary retention %s keep black text like the cohort tabs).
const blackTextRanges = new WeakMap<Worksheet, [number, number, number, number][]>();

function markBlackTextRange(ws: Worksheet, r1: number, c1: number, r2: number, c2: number): void {
  const ranges = blackTextRanges.get(ws) || [];
  ranges.push([r1, c1, r2, c2]);
  blackTextRanges.set(ws, ranges);
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

  // Row 5: section headers bold, centered across each section, underlined
  for (let c = 1; c <= maxCol; c++) {
    if (ws.getCell(5, c).value) setFont(ws, 5, c, true);
  }
  const sectionSpans: [number, number][] = [
    [layout.cust_id, layout.rank],
    [layout.arr_start, layout.arr_end],
    [layout.churn_start, layout.churn_end],
    [layout.downsell_start, layout.downsell_end],
    [layout.upsell_start, layout.upsell_end],
    [layout.new_biz_start, layout.new_biz_end],
  ];
  for (const [sc, ec] of sectionSpans) {
    centerAcrossUnderline(ws, 5, sc, ec);
  }

  // Row 6: bold, center, underlined
  for (let c = 1; c <= maxCol; c++) {
    if (ws.getCell(6, c).value) {
      setFont(ws, 6, c, true);
      setAlign(ws, 6, c, 'center');
    }
  }
  underlineSpan(ws, 6, layout.cust_id, maxCol);

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

  // Column widths
  ws.getColumn(1).width = 7;
  ws.getColumn(layout.cust_id).width = 24;
  for (let c = layout.attr_start; c <= layout.attr_end; c++) ws.getColumn(c).width = 14;
  ws.getColumn(layout.cohort).width = 15;
  ws.getColumn(layout.rank).width = 15;
  ws.getColumn(layout.label).width = 8;
  for (let c = layout.arr_start; c <= maxCol; c++) ws.getColumn(c).width = 11;

  // Freeze panes
  ws.views = [{ state: 'frozen', xSplit: layout.arr_start - 1, ySplit: firstDataRow - 1, showGridLines: false }];
}

export function formatRetentionTab(
  ws: Worksheet, _config: import('./types').EngineConfig, filterBlocks: FilterBlock[],
  numDerived: number, numAttrs: number,
  s1Label: number, s1Start: number, s1End: number,
  s2Label: number, s2Start: number, s2End: number,
  s3Label: number, s3Start: number, s3End: number,
  filterStart: number, cohortFc: number,
  hasAnn = false
): void {
  const maxCol = s3End;
  const blockHeight = hasAnn ? 22 : 19;
  const maxRow = 5 + filterBlocks.length * blockHeight;

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
    const start = 5 + blockIdx * blockHeight;

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
    const annRows = hasAnn ? [start + 16, start + 17, start + 18] : [];
    const rNlPct = start + (hasAnn ? 19 : 16);
    const rNlGrowth = start + (hasAnn ? 20 : 17);

    // Title bold, centered across the block, underlined
    setFont(ws, rTitle, s1Label, true);
    centerAcrossUnderline(ws, rTitle, s1Label, s3End);

    // Section headers bold, centered across each section, underlined
    for (const [lbl, sEnd] of [[s1Label, s1End], [s2Label, s2End], [s3Label, s3End]] as [number, number][]) {
      setFont(ws, rSections, lbl, true);
      centerAcrossUnderline(ws, rSections, lbl, sEnd);
    }

    // Column headers
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      setFont(ws, rHeader, filterStart + attrIdx, true);
      setAlign(ws, rHeader, filterStart + attrIdx, 'center');
    }
    setFont(ws, rHeader, cohortFc, true);
    setAlign(ws, rHeader, cohortFc, 'center');
    underlineSpan(ws, rHeader, filterStart, cohortFc);
    for (const [lbl, sEnd] of [[s1Label, s1End], [s2Label, s2End], [s3Label, s3End]] as [number, number][]) {
      underlineSpan(ws, rHeader, lbl, sEnd);
    }

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
      for (const r of [rLostRet, rPunitRet, rNetRet, ...annRows, rNlPct, rNlGrowth]) setNumFmt(ws, r, c, NF_PCT_DEC);
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

    // Italicize the percentage rows across all section 1 columns
    for (const r of [rLostRet, rPunitRet, rNetRet, ...annRows, rNlPct, rNlGrowth]) {
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

    // Overlines above the sum rows in each section
    overlineSpan(ws, rRetained, s1Label, s1End);   // Retained = sum of BoP..Upsell
    overlineSpan(ws, rEop, s1Label, s1End);        // EoP = sum of Retained..New Logo
    overlineSpan(ws, rDownsell, s2Label, s2End);   // Retained Customers = sum of BoP..Churned
    overlineSpan(ws, rRetained, s2Label, s2End);   // EoP Customers = sum of Retained..New Logo
    overlineSpan(ws, rUpsell, s3Label, s3End);     // Retained Customers
    overlineSpan(ws, rNewLogo, s3Label, s3End);    // EoP Customers
  }

  // Column widths
  ws.getColumn(1).width = 7;
  for (let c = filterStart; c <= cohortFc; c++) ws.getColumn(c).width = 14;
  ws.getColumn(cohortFc + 1).width = 2.5;
  for (const [lbl, s, e] of [[s1Label, s1Start, s1End], [s2Label, s2Start, s2End], [s3Label, s3Start, s3End]] as [number, number, number][]) {
    ws.getColumn(lbl).width = 26;
    for (let c = s; c <= e; c++) ws.getColumn(c).width = 12;
  }
  ws.getColumn(s1End + 1).width = 2.5;
  ws.getColumn(s2End + 1).width = 2.5;
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
    const s1End = s1Start + numDates - 1;
    const s2End = s2Start + numDates - 1;

    // Title bold, centered across the block, underlined
    setFont(ws, blockStart, qCol, true);
    centerAcrossUnderline(ws, blockStart, qCol, maxCol);

    // Section headers bold, centered across each section, underlined
    const sectionHeaderSpans: [number, number][] = [
      [s1Start, s1End],
      [s2Start, s2End],
      [s3Label, s3DataEnd],
      [s4Label, s4DataEnd],
    ];
    for (const [sc, ec] of sectionHeaderSpans) {
      setFont(ws, rSectionHeaders, sc, true);
      centerAcrossUnderline(ws, rSectionHeaders, sc, ec);
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
    for (const [startC, endC] of sectionHeaderSpans) {
      // Bottom border under column headers row
      for (let c = startC; c <= endC; c++) {
        ws.getCell(rHeaders, c).border = { ...ws.getCell(rHeaders, c).border, bottom: THIN_BORDER };
      }
      // Top border above Total row
      for (let c = startC; c <= endC; c++) {
        ws.getCell(rTotal, c).border = { ...ws.getCell(rTotal, c).border, top: THIN_BORDER };
      }
    }
    // Bottom border under the leading column headers (Quarter/Year/filters/Cohort)
    underlineSpan(ws, rHeaders, qCol, cohortLabelCol);

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

  // Column widths
  ws.getColumn(1).width = 7;
  ws.getColumn(qCol).width = 9;
  ws.getColumn(yCol).width = 8;
  for (let c = filterStart; c <= cohortLabelCol; c++) ws.getColumn(c).width = 13;
  for (let c = s1Start; c <= s1Start + numDates - 1; c++) ws.getColumn(c).width = 12;
  for (let c = s2Start; c <= s2Start + numDates - 1; c++) ws.getColumn(c).width = 12;
  ws.getColumn(s3Label).width = 24;
  ws.getColumn(s3StartVal).width = 13;
  for (let c = s3DataStart; c <= s3DataStart + numDates - 1; c++) ws.getColumn(c).width = 9;
  ws.getColumn(s4Label).width = 24;
  ws.getColumn(s4StartVal).width = 13;
  for (let c = s4DataStart; c <= s4DataEnd; c++) ws.getColumn(c).width = 9;
  for (const c of [s1Start + numDates, s2Start + numDates, s3DataStart + numDates]) {
    ws.getColumn(c).width = 2.5;
  }
}

export function formatSummaryTab(
  ws: Worksheet, lastSegCol: number,
  sections: import('./summary').SummarySectionLayout[],
  headersAreDates = false
): void {
  const allCol = 4;         // D
  const checkCol = lastSegCol + 1;
  const lastRow = Math.max(...sections.map(s => s.startRow + s.numRows - 1), 8);

  const sectionNumFmt: Record<string, string> = {
    gross: NF_PCT,
    net: NF_PCT,
    logo: NF_PCT,
    ann_gross: NF_PCT,
    ann_net: NF_PCT,
    pct_of_total: NF_PCT,
    dollars: NF_DOLLAR,
    customers: NF_NUMBER,
    per_customer: NF_DOLLAR,
  };

  // Base font
  for (let r = 1; r <= lastRow; r++) {
    for (let c = 1; c <= checkCol; c++) {
      ws.getCell(r, c).font = baseFont();
    }
  }

  // Units cell
  formatUnitsCell(ws, 3, 1, 2);

  // Segment identifier input cell (blue on light yellow, like Control inputs)
  const idCell = ws.getCell(5, 2);
  idCell.font = { name: 'Times New Roman', size: 10, color: { argb: BLUE_COLOR } };
  idCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } };
  idCell.alignment = { horizontal: 'center' };
  setFont(ws, 5, 1, true);

  // Row 7 banner: identifier name centered across the segment columns, underlined
  setFont(ws, 7, allCol, true);
  for (let c = allCol; c <= lastSegCol; c++) {
    setAlign(ws, 7, c, 'centerContinuous');
  }
  underlineSpan(ws, 7, allCol, lastSegCol);

  // Row 8 headers: bold, centered, bottom border
  const thinBorder = { style: 'thin' as const };
  for (let c = 2; c <= checkCol; c++) {
    setFont(ws, 8, c, true);
    setAlign(ws, 8, c, 'center');
    ws.getCell(8, c).border = { ...ws.getCell(8, c).border, bottom: thinBorder };
  }
  // Header cells with plain-string values are inputs (blue on light yellow);
  // formula-linked headers are green via applyFormulaColoring and get NF_DATE
  // only when the linked cells hold dates (the cohort fallback)
  for (let c = allCol; c <= lastSegCol; c++) {
    const cell = ws.getCell(8, c);
    const v = cell.value;
    if (v != null && typeof v === 'object' && 'formula' in v) {
      if (headersAreDates) cell.numFmt = NF_DATE;
    } else {
      cell.font = { name: 'Times New Roman', size: 10, bold: true, color: { argb: BLUE_COLOR } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } };
    }
  }

  // Check summary cell + check column
  setNumFmt(ws, 1, 2, NF_NUMBER);

  // Section bodies
  for (const section of sections) {
    if (section.numRows <= 0) continue;
    const fmt = sectionNumFmt[section.key];
    setFont(ws, section.startRow, 2, true);

    for (let r = section.startRow; r < section.startRow + section.numRows; r++) {
      for (let c = allCol; c <= lastSegCol; c++) {
        setNumFmt(ws, r, c, fmt);
        setAlign(ws, r, c, 'right');
      }
      if (section.key === 'dollars') {
        setNumFmt(ws, r, checkCol, NF_NUMBER);
        setAlign(ws, r, checkCol, 'right');
      }
    }

    // 3-color scale (red → white at 1.0 → green) on the retention sections.
    // Text stays black — like the cohort tabs — so formula color-coding skips them.
    if (['gross', 'net', 'logo', 'ann_gross', 'ann_net'].includes(section.key)) {
      markBlackTextRange(ws, section.startRow, allCol, section.startRow + section.numRows - 1, lastSegCol);
      ws.addConditionalFormatting({
        ref: `${colLetter(allCol)}${section.startRow}:${colLetter(lastSegCol)}${section.startRow + section.numRows - 1}`,
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
        }],
      });
    }
  }
}

export function formatDataSummaryTab(
  ws: Worksheet, attrBlocks: { col: number; numRows: number }[]
): void {
  const thinBorder = { style: 'thin' as const };
  const maxCol = attrBlocks.length > 0 ? attrBlocks[attrBlocks.length - 1].col + 2 : 1;
  const maxRow = Math.max(...attrBlocks.map(b => 9 + b.numRows), 6);

  // Base font
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      ws.getCell(r, c).font = baseFont();
    }
  }

  setFont(ws, 1, 1, true);
  setNumFmt(ws, 1, 2, NF_NUMBER);

  for (const { col, numRows } of attrBlocks) {
    const firstRow = 7;
    const lastRow = firstRow + numRows - 1;
    const totalRow = lastRow + 1;
    const checkRow = totalRow + 1;

    // Attribute name + column headers
    setFont(ws, 5, col, true);
    for (let c = col; c <= col + 2; c++) {
      setFont(ws, 6, c, true);
      setAlign(ws, 6, c, 'center');
      ws.getCell(6, c).border = { ...ws.getCell(6, c).border, bottom: thinBorder };
    }

    for (let r = firstRow; r <= lastRow; r++) {
      setNumFmt(ws, r, col + 1, NF_NUMBER);
      setNumFmt(ws, r, col + 2, NF_PCT_DEC);
      setAlign(ws, r, col + 1, 'right');
      setAlign(ws, r, col + 2, 'right');
    }

    // Total row bold with top border; check row italic
    for (let c = col; c <= col + 2; c++) {
      setFont(ws, totalRow, c, true);
      ws.getCell(totalRow, c).border = { ...ws.getCell(totalRow, c).border, top: thinBorder };
    }
    setNumFmt(ws, totalRow, col + 1, NF_NUMBER);
    setNumFmt(ws, totalRow, col + 2, NF_PCT);
    setNumFmt(ws, checkRow, col + 1, NF_NUMBER);
    const checkCell = ws.getCell(checkRow, col);
    checkCell.font = { ...baseFont(), italic: true };
    const checkVal = ws.getCell(checkRow, col + 1);
    checkVal.font = { ...baseFont(), italic: true };
  }
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
  const s1End = s1Start + numDates - 1;
  const s2End = s2Start + numDates - 2;

  // Base font
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      cell.font = baseFont();
    });
  });

  // Units cell
  formatUnitsCell(ws, 3, 1, rankNumCol);

  // Row 5: Section headers bold, centered across each section, underlined
  for (const [sc, ec] of [[s1Start, s1End], [s2Start, s2End], [s3Start, s3End]] as [number, number][]) {
    setFont(ws, 5, sc, true);
    // A section can be empty (a one-period sheet has no YoY columns);
    // then the underline goes on the heading cell itself.
    if (ec >= sc) centerAcrossUnderline(ws, 5, sc, ec);
    else underlineSpan(ws, 5, sc, sc);
  }

  // Row 6: Column headers bold, centered, underlined
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
  underlineSpan(ws, 6, custIdCol, s3End);

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

  // Overlines: underline the customer numbers above each sum row
  overlineSpan(ws, rTopTotal, custIdCol, s3End);   // customers above Top-N total
  overlineSpan(ws, rTotal, custIdCol, s3End);      // Other Customers above Total

  // Other row
  for (let i = 0; i < numDates; i++) setNumFmt(ws, rOther, s1Start + i, NF_NUMBER);
  for (let i = 0; i < numDates - 1; i++) setNumFmt(ws, rOther, s2Start + i, NF_PCT);
  for (let i = 0; i < numDates; i++) setNumFmt(ws, rOther, s3Start + i, NF_PCT_DEC);
  const otherCell = ws.getCell(rOther, custIdCol);
  otherCell.alignment = { ...otherCell.alignment, indent: 1 };

  // Memo section: italic, "Memo:" label also underlined
  for (let r = rMemoStart; r <= rMemoStart + 4; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      if (cell.value != null) cell.font = { ...cell.font, italic: true };
    }
  }
  const memoCell = ws.getCell(rMemoStart, custIdCol);
  memoCell.font = { ...memoCell.font, underline: true };

  // Memo rows
  for (let tierIdx = 0; tierIdx < 4; tierIdx++) {
    const r = rMemoStart + 1 + tierIdx;
    for (let i = 0; i < numDates; i++) setNumFmt(ws, r, s1Start + i, NF_NUMBER);
    for (let i = 0; i < numDates - 1; i++) setNumFmt(ws, r, s2Start + i, NF_PCT);
    for (let i = 0; i < numDates; i++) setNumFmt(ws, r, s3Start + i, NF_PCT_DEC);
  }

  // Column widths
  ws.getColumn(1).width = 7;
  ws.getColumn(rankNumCol).width = 7;
  ws.getColumn(custIdCol).width = 26;
  for (let c = attrStart; c <= cohortCol; c++) ws.getColumn(c).width = 14;
  for (let c = s1Start; c <= s1End; c++) ws.getColumn(c).width = 12;
  for (let c = s2Start; c <= s2End; c++) ws.getColumn(c).width = 10;
  for (let c = s3Start; c <= s3End; c++) ws.getColumn(c).width = 10;
  ws.getColumn(s1End + 1).width = 2.5;
  if (s2End >= s2Start) {
    ws.getColumn(s2End + 1).width = 2.5;
  } else {
    // Empty growth section: '% YoY Growth' holds its column alone
    ws.getColumn(s2Start).width = 14;
  }
}

/**
 * Apply formula auditing color-coding to all sheets.
 */
export function applyFormulaColoring(wb: Workbook, skipSheets?: string[]): void {
  const skip = new Set(skipSheets || []);

  for (const ws of wb.worksheets) {
    if (skip.has(ws.name)) continue;

    const skipRanges = blackTextRanges.get(ws);

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.value == null) return;
        if (skipRanges?.some(([r1, c1, r2, c2]) =>
          rowNumber >= r1 && rowNumber <= r2 && colNumber >= c1 && colNumber <= c2)) return;

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
    [/Data Summary/, 'D9D2E9'],
    [/Summary/, 'A9D18E'],
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

function colNumFromLetter(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.toUpperCase().charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Format the copied raw data sheet (e.g. "Customer Cube") to match the
 * rest of the pack: Times New Roman, an underlined header row, number
 * formats on the data cells, and readable column widths.
 */
export function formatRawDataTab(ws: Worksheet, config: import('./types').EngineConfig): void {
  const headerRow = config.raw_data_first_row - 1;
  const idCol = colNumFromLetter(config.customer_id_col);
  const attrCols = new Set(Object.values(config.attributes || {}).map(l => colNumFromLetter(l)));

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      cell.font = baseFont();
      if (rowNumber === headerRow) {
        cell.font = { ...cell.font, bold: true };
        cell.alignment = { horizontal: 'center' };
        cell.border = { bottom: THIN_BORDER };
      } else if (rowNumber >= config.raw_data_first_row) {
        const v = cell.value;
        if (v instanceof Date) {
          cell.numFmt = 'm/d/yyyy';
          cell.alignment = { horizontal: 'center' };
        } else if (typeof v === 'number') {
          cell.numFmt = NF_NUMBER;
          cell.alignment = { horizontal: 'right' };
        }
      }
    });
  });

  // Column widths: wider for identifiers, snug for data columns
  const usedCols = new Set<number>();
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (_cell, colNumber) => usedCols.add(colNumber));
  });
  ws.getColumn(1).width = 7;
  for (const cn of usedCols) {
    if (cn === 1) continue;
    if (cn === idCol) ws.getColumn(cn).width = 24;
    else if (attrCols.has(cn)) ws.getColumn(cn).width = 14;
    else ws.getColumn(cn).width = 12;
  }
}
