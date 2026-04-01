/**
 * Utility functions for Excel formula generation.
 * Column math, formula builders, and layout computation.
 */

/** Convert 1-based column number to Excel column letter. e.g. 1='A', 27='AA'. */
export function colLetter(n: number): string {
  let result = '';
  let num = n;
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

/** Convert Excel column letter to 1-based number. e.g. 'A'=1, 'AA'=27. */
export function colNum(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}

/** Build a range reference like 'Sheet'!$A$2:$A$23498. */
export function makeRange(
  sheet: string, colLetterStr: string,
  firstRow: number, lastRow: number,
  absCol = true, absRow = true
): string {
  const dc = absCol ? '$' : '';
  const dr = absRow ? '$' : '';
  return `'${sheet}'!${dc}${colLetterStr}${dr}${firstRow}:${dc}${colLetterStr}${dr}${lastRow}`;
}

/** Build a cell reference like 'Sheet'!$A$1. */
export function makeCell(
  sheet: string, colLetterStr: string, row: number,
  absCol = true, absRow = true
): string {
  const dc = absCol ? '$' : '';
  const dr = absRow ? '$' : '';
  return `'${sheet}'!${dc}${colLetterStr}${dr}${row}`;
}

/** Build a SUMIFS formula. criteriaPairs: [[range, criteria], ...]. */
export function sumifs(sumRange: string, ...criteriaPairs: [string, string][]): string {
  const parts = [sumRange];
  for (const [crRange, crValue] of criteriaPairs) {
    parts.push(crRange, crValue);
  }
  return `=SUMIFS(${parts.join(',')})`;
}

/** Build a COUNTIFS formula. criteriaPairs: [[range, criteria], ...]. */
export function countifs(...criteriaPairs: [string, string][]): string {
  const parts: string[] = [];
  for (const [crRange, crValue] of criteriaPairs) {
    parts.push(crRange, crValue);
  }
  return `=COUNTIFS(${parts.join(',')})`;
}

export interface CleanLayout {
  cust_id: number;
  attr_start: number;
  attr_end: number;
  num_attrs: number;
  cohort: number;
  rank: number;
  label: number;
  arr_start: number;
  arr_end: number;
  churn_start: number;
  churn_end: number;
  downsell_start: number;
  downsell_end: number;
  upsell_start: number;
  upsell_end: number;
  new_biz_start: number;
  new_biz_end: number;
  yoy_offset: number;
  num_dates: number;
  num_derived: number;
}

/**
 * Compute column positions for a clean data tab.
 * Returns all column positions (1-based column numbers).
 */
export function computeCleanLayout(numAttrs: number, numDates: number, yoyOffset: number): CleanLayout {
  const numDerived = numDates - yoyOffset;

  const custIdCol = 2;  // B
  const attrStart = 3;  // C
  const attrEnd = attrStart + numAttrs - 1;
  const cohortCol = attrEnd + 1;
  const rankCol = cohortCol + 1;
  const labelCol = rankCol + 1;

  const arrStart = labelCol + 1;
  const arrEnd = arrStart + numDates - 1;

  const churnStart = arrEnd + 2;
  const churnEnd = churnStart + numDerived - 1;

  const downsellStart = churnEnd + 2;
  const downsellEnd = downsellStart + numDerived - 1;

  const upsellStart = downsellEnd + 2;
  const upsellEnd = upsellStart + numDerived - 1;

  const newBizStart = upsellEnd + 2;
  const newBizEnd = newBizStart + numDerived - 1;

  return {
    cust_id: custIdCol,
    attr_start: attrStart,
    attr_end: attrEnd,
    num_attrs: numAttrs,
    cohort: cohortCol,
    rank: rankCol,
    label: labelCol,
    arr_start: arrStart,
    arr_end: arrEnd,
    churn_start: churnStart,
    churn_end: churnEnd,
    downsell_start: downsellStart,
    downsell_end: downsellEnd,
    upsell_start: upsellStart,
    upsell_end: upsellEnd,
    new_biz_start: newBizStart,
    new_biz_end: newBizEnd,
    yoy_offset: yoyOffset,
    num_dates: numDates,
    num_derived: numDerived,
  };
}

/** Return display label for a granularity level. */
export function granularityLabel(granularity: string): string {
  return { monthly: 'Monthly', quarterly: 'Quarterly', annual: 'Annual' }[granularity] || granularity;
}

/** Get the year-over-year offset for a given granularity. */
export function getYoyOffset(granularity: string): number {
  return { monthly: 12, quarterly: 4, annual: 1 }[granularity] || 1;
}

/** Build the period label formula for column headers (row 6). */
export function periodLabelFormula(granularity: string, quarterCell: string, yearCell: string): string | null {
  if (granularity === 'quarterly') {
    return `="Q"&${quarterCell}&"'"&RIGHT(${yearCell},2)`;
  } else if (granularity === 'annual') {
    return `="FY"&"'"&RIGHT(${yearCell},2)`;
  }
  return null;  // monthly uses date values
}
