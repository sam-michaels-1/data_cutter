/**
 * Summary computation engine.
 * Builds the retention/metric summary table (periods x segments) for the
 * Summary page, mirroring the Excel `${Gran} Summary` tab formulas.
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig } from './types';
import type { AttributeOption } from '../types/dashboard';
import {
  readRawData,
  readCleanedData,
  aggregateToGranularity,
  buildArrMatrix,
  computeDerived,
  buildCohortMap,
  buildAttrLookup,
  getYoyOffset,
  getPivotValue,
  computeAttributeOptions,
} from './compute';
import { getIdentifierValue } from './histograms';

export interface SummarySection {
  key: 'gross' | 'net' | 'logo' | 'pct_of_total' | 'dollars' | 'customers' | 'per_customer';
  label: string;
  format: 'pct' | 'currency' | 'count';
  rows: { period: string; values: (number | null)[] }[];
}

export interface SummaryResult {
  sections: SummarySection[];
  columns: string[];               // ['All', ...segmentValues]
  identifier: string;
  identifiers: string[];
  granularity: string;
  available_granularities: string[];
  scale_factor: number;
  data_type: string;
  attribute_options: AttributeOption[];
}

export function computeSummaryData(
  wb: Workbook,
  config: EngineConfig,
  granularity?: string,
  filters?: Record<string, string | string[]>,
  identifier?: string,
): SummaryResult {
  let rawRows = config.input_format === 'cleaned'
    ? readCleanedData(wb, config)
    : readRawData(wb, config);
  const fyMonth = config.fiscal_year_end_month;
  const scaleFactor = config.scale_factor;
  const outputGrans = config.output_granularities;
  const attrNames = Object.keys(config.attributes || {});
  const dataType = config.data_type || 'arr';
  const metricLabel = dataType === 'revenue' ? 'Revenue' : 'ARR';

  const attributeOptions = computeAttributeOptions(rawRows, attrNames);

  // Apply regular attribute filters (not Cohort — applied post-aggregation)
  if (filters) {
    for (const [attrName, attrValue] of Object.entries(filters)) {
      if (attrName === 'Cohort') continue;
      if (Array.isArray(attrValue)) {
        if (attrValue.length > 0) rawRows = rawRows.filter(r => attrValue.includes(String(r[attrName])));
      } else if (attrValue) {
        rawRows = rawRows.filter(r => r[attrName] === attrValue);
      }
    }
  }

  let targetGran = granularity && outputGrans.includes(granularity) ? granularity : null;
  if (!targetGran) {
    const prefOrder = ['annual', 'quarterly', 'monthly'];
    targetGran = prefOrder.find(g => outputGrans.includes(g)) || outputGrans[0];
  }
  const available = ['annual', 'quarterly', 'monthly'].filter(g => outputGrans.includes(g));

  const { records } = aggregateToGranularity(rawRows, targetGran, fyMonth, dataType);
  const { pivot, periods } = buildArrMatrix(records);
  const yoyOffset = getYoyOffset(targetGran);
  const cohortMap = buildCohortMap(pivot, periods);
  const attrLookup = buildAttrLookup(rawRows, attrNames);

  const cohortValuesSet = new Set(cohortMap.values());
  const cohortValues = periods.filter(p => cohortValuesSet.has(p));
  const allAttributeOptions: AttributeOption[] = [
    { name: 'Cohort', values: cohortValues, multiSelect: true },
    ...attributeOptions,
  ];

  // Apply cohort filter (post-aggregation, on the pivot)
  const cohortFilter = filters?.['Cohort'];
  if (cohortFilter && Array.isArray(cohortFilter) && cohortFilter.length > 0 && cohortFilter.length < cohortValues.length) {
    const selectedSet = new Set(cohortFilter);
    for (const [cust] of [...pivot]) {
      const custCohort = cohortMap.get(cust);
      if (!custCohort || !selectedSet.has(custCohort)) {
        pivot.delete(cust);
      }
    }
  }

  const derived = computeDerived(pivot, periods, yoyOffset);
  const sf = scaleFactor;

  // Segment identifier and column groups
  const identifiers = ['Cohort', ...attrNames];
  const effectiveIdentifier = identifier && identifiers.includes(identifier)
    ? identifier
    : (attrNames[0] || 'Cohort');

  const customers = [...pivot.keys()];
  const segOf = new Map<string, string>();
  for (const cust of customers) {
    segOf.set(cust, getIdentifierValue(cust, effectiveIdentifier, cohortMap, attrLookup));
  }
  let segmentValues = [...new Set(segOf.values())];
  if (effectiveIdentifier === 'Cohort') {
    segmentValues = segmentValues.sort((a, b) => periods.indexOf(a) - periods.indexOf(b));
  } else {
    segmentValues = segmentValues.sort((a, b) => a.localeCompare(b));
  }
  const columnGroups: string[][] = [
    customers,
    ...segmentValues.map(v => customers.filter(c => segOf.get(c) === v)),
  ];

  const numPeriods = periods.length;

  // Per-period aggregates for each column (index 0 = All)
  const eop: number[][] = [];      // scaled $
  const eopCount: number[][] = [];
  const bop: number[][] = [];
  const churnSum: number[][] = [];
  const downSum: number[][] = [];
  const upSum: number[][] = [];
  const bopCount: number[][] = [];
  const churnCount: number[][] = [];

  for (const group of columnGroups) {
    const eopCol: number[] = [];
    const eopCntCol: number[] = [];
    const bopCol: number[] = [];
    const churnCol: number[] = [];
    const downCol: number[] = [];
    const upCol: number[] = [];
    const bopCntCol: number[] = [];
    const churnCntCol: number[] = [];

    for (let i = 0; i < numPeriods; i++) {
      const p = periods[i];
      let e = 0, c = 0;
      for (const cust of group) {
        const v = getPivotValue(pivot, cust, p);
        e += v;
        if (v !== 0) c++;
      }
      eopCol.push(e / sf);
      eopCntCol.push(c);

      if (i >= yoyOffset) {
        const prior = periods[i - yoyOffset];
        let b = 0, ch = 0, dn = 0, up = 0, bc = 0, cc = 0;
        for (const cust of group) {
          b += getPivotValue(pivot, cust, prior);
          ch += derived.churn.get(cust)?.get(p) || 0;
          dn += derived.downsell.get(cust)?.get(p) || 0;
          up += derived.upsell.get(cust)?.get(p) || 0;
          if (getPivotValue(pivot, cust, prior) !== 0) bc++;
          if ((derived.churn.get(cust)?.get(p) || 0) !== 0) cc++;
        }
        bopCol.push(b);
        churnCol.push(ch);
        downCol.push(dn);
        upCol.push(up);
        bopCntCol.push(bc);
        churnCntCol.push(cc);
      }
    }
    eop.push(eopCol);
    eopCount.push(eopCntCol);
    bop.push(bopCol);
    churnSum.push(churnCol);
    downSum.push(downCol);
    upSum.push(upCol);
    bopCount.push(bopCntCol);
    churnCount.push(churnCntCol);
  }

  const numCols = columnGroups.length;
  const derivedRows = Math.max(numPeriods - yoyOffset, 0);

  function rowValues(get: (ci: number, i: number) => number | null, count: number, offset: number) {
    const rows: { period: string; values: (number | null)[] }[] = [];
    for (let j = 0; j < count; j++) {
      const i = j + offset;
      const values: (number | null)[] = [];
      for (let ci = 0; ci < numCols; ci++) values.push(get(ci, i));
      rows.push({ period: periods[i], values });
    }
    return rows;
  }

  const dollarsRows = rowValues((ci, i) => eop[ci][i], numPeriods, 0);
  const customersRows = rowValues((ci, i) => eopCount[ci][i], numPeriods, 0);

  const sections: SummarySection[] = [
    {
      key: 'gross',
      label: 'Gross Retention',
      format: 'pct',
      rows: rowValues((ci, i) => {
        const j = i - yoyOffset;
        const b = bop[ci][j];
        return b !== 0 ? (b + churnSum[ci][j] + downSum[ci][j]) / b : null;
      }, derivedRows, yoyOffset),
    },
    {
      key: 'net',
      label: 'Net Retention',
      format: 'pct',
      rows: rowValues((ci, i) => {
        const j = i - yoyOffset;
        const b = bop[ci][j];
        return b !== 0 ? (b + churnSum[ci][j] + downSum[ci][j] + upSum[ci][j]) / b : null;
      }, derivedRows, yoyOffset),
    },
    {
      key: 'logo',
      label: 'Logo Retention',
      format: 'pct',
      rows: rowValues((ci, i) => {
        const j = i - yoyOffset;
        const bc = bopCount[ci][j];
        return bc !== 0 ? (bc - churnCount[ci][j]) / bc : null;
      }, derivedRows, yoyOffset),
    },
    {
      key: 'pct_of_total',
      label: `% of ${metricLabel}`,
      format: 'pct',
      rows: rowValues((ci, i) => {
        const all = eop[0][i];
        return all !== 0 ? eop[ci][i] / all : null;
      }, numPeriods, 0),
    },
    { key: 'dollars', label: `$ ${metricLabel}`, format: 'currency', rows: dollarsRows },
    { key: 'customers', label: 'Ending Customers', format: 'count', rows: customersRows },
    {
      key: 'per_customer',
      label: `${metricLabel} per Customer`,
      format: 'currency',
      rows: rowValues((ci, i) => {
        const c = eopCount[ci][i];
        return c !== 0 ? eop[ci][i] / c : null;
      }, numPeriods, 0),
    },
  ];

  return {
    sections,
    columns: ['All', ...segmentValues],
    identifier: effectiveIdentifier,
    identifiers,
    granularity: targetGran,
    available_granularities: available,
    scale_factor: scaleFactor,
    data_type: dataType,
    attribute_options: allAttributeOptions,
  };
}
