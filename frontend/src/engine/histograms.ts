/**
 * Histogram / distribution computation engine.
 * Builds on the core compute helpers to produce data for the Histograms page.
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig } from './types';
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

/* ─── Result types ─── */

export interface HistogramBucket {
  label: string;
  count: number;
  min: number;
  max: number;
}

export interface MekkoSegment {
  yLabel: string;
  value: number;
  pct: number;
}

export interface MekkoColumn {
  xLabel: string;
  xTotal: number;
  xPct: number;
  stacks: MekkoSegment[];
}

export interface MekkoData {
  columns: MekkoColumn[];
  yLabels: string[];
}

export interface PieSlice {
  label: string;
  value: number;
  pct: number;
}

export interface PieChartData {
  identifierName: string;
  arrSlices: PieSlice[];
  countSlices: PieSlice[];
}

export interface GridCell {
  value: number | null;
  count: number;
  totalARR: number;
}

export interface GridData {
  xLabels: string[];
  yLabels: string[];
  grid: (GridCell | null)[][];   // rows[y][x]
  xTotals: (number | null)[];
  yTotals: (number | null)[];
}

export interface HistogramResult {
  arrHistogram: HistogramBucket[];
  growthHistogram: HistogramBucket[];
  mekkoARR: MekkoData;
  mekkoCount: MekkoData;
  pieCharts: PieChartData[];
  growthGrid: GridData;
  netRetentionGrid: GridData;
  lossRetentionGrid: GridData;
  identifiers: string[];
  available_granularities: string[];
  granularity: string;
  scale_factor: number;
  data_type: string;
  attribute_options: { name: string; values: string[]; multiSelect?: boolean }[];
  latestPeriodLabel: string;
  priorPeriodLabel: string;
}

/* ─── Helpers ─── */

function percentile(sorted: number[], p: number): number {
  if (sorted.length === 0) return 0;
  const idx = (p / 100) * (sorted.length - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return sorted[lo];
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
}

function roundBucket(val: number, direction: 'floor' | 'ceil'): number {
  // Round to nearest "nice" number for bucket boundaries
  const abs = Math.abs(val);
  let step: number;
  if (abs >= 10_000_000) step = 1_000_000;
  else if (abs >= 1_000_000) step = 500_000;
  else if (abs >= 100_000) step = 100_000;
  else if (abs >= 10_000) step = 10_000;
  else if (abs >= 1_000) step = 1_000;
  else step = 100;

  return direction === 'floor'
    ? Math.floor(val / step) * step
    : Math.ceil(val / step) * step;
}

function buildPercentileBuckets(values: number[], count: number): HistogramBucket[] {
  if (values.length === 0) return [];
  const sorted = [...values].sort((a, b) => a - b);
  const p10 = roundBucket(percentile(sorted, 10), 'floor');
  const p90 = roundBucket(percentile(sorted, 90), 'ceil');

  if (p10 >= p90) {
    // Degenerate case: all values similar
    return [{ label: formatBucketLabel(sorted[0], sorted[sorted.length - 1], true, true), count: values.length, min: sorted[0], max: sorted[sorted.length - 1] }];
  }

  const rawInterior = count - 2; // 8 interior buckets
  const step = (p90 - p10) / rawInterior;
  const rawBoundaries: number[] = [p10];
  for (let i = 1; i <= rawInterior; i++) {
    rawBoundaries.push(roundBucket(p10 + step * i, 'ceil'));
  }

  // Deduplicate boundaries after rounding to avoid degenerate buckets like "$30K - $30K"
  const boundaries: number[] = [rawBoundaries[0]];
  for (let i = 1; i < rawBoundaries.length; i++) {
    if (rawBoundaries[i] !== boundaries[boundaries.length - 1]) {
      boundaries.push(rawBoundaries[i]);
    }
  }
  const interior = boundaries.length - 1;

  const buckets: HistogramBucket[] = [];

  // Floor bucket: <= p10
  buckets.push({ label: '', count: 0, min: -Infinity, max: p10 });
  // Interior buckets
  for (let i = 0; i < interior; i++) {
    buckets.push({ label: '', count: 0, min: boundaries[i], max: boundaries[i + 1] });
  }
  // Ceiling bucket: > p90
  buckets.push({ label: '', count: 0, min: p90, max: Infinity });

  // Count values into buckets
  for (const v of values) {
    if (v <= p10) { buckets[0].count++; continue; }
    if (v > p90) { buckets[buckets.length - 1].count++; continue; }
    for (let i = 1; i <= interior; i++) {
      if (v <= boundaries[i]) { buckets[i].count++; break; }
    }
  }

  // Generate labels
  buckets[0].label = formatBucketLabel(-Infinity, p10, true, false);
  for (let i = 1; i <= interior; i++) {
    buckets[i].label = formatBucketLabel(boundaries[i - 1], boundaries[i], false, false);
  }
  buckets[buckets.length - 1].label = formatBucketLabel(p90, Infinity, false, true);

  return buckets;
}

function formatDollar(v: number): string {
  const abs = Math.abs(v);
  if (abs >= 1_000_000_000) return `$${(v / 1_000_000_000).toFixed(1)}B`;
  if (abs >= 1_000_000) return `$${(v / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000) return `$${(v / 1_000).toFixed(0)}K`;
  return `$${v.toFixed(0)}`;
}

function formatPctLabel(v: number): string {
  return `${(v * 100).toFixed(0)}%`;
}

function formatBucketLabel(min: number, max: number, isFloor: boolean, isCeiling: boolean, isPct = false): string {
  const fmt = isPct ? formatPctLabel : formatDollar;
  if (isFloor) return `<= ${fmt(max)}`;
  if (isCeiling) return `> ${fmt(min)}`;
  return `${fmt(min)} - ${fmt(max)}`;
}

function buildPctBuckets(values: number[], count: number): HistogramBucket[] {
  if (values.length === 0) return [];
  const sorted = [...values].sort((a, b) => a - b);
  const p10 = Math.floor(percentile(sorted, 10) * 20) / 20; // round to nearest 5%
  const p90 = Math.ceil(percentile(sorted, 90) * 20) / 20;

  if (p10 >= p90) {
    const label = formatBucketLabel(sorted[0], sorted[sorted.length - 1], true, true, true);
    return [{ label, count: values.length, min: sorted[0], max: sorted[sorted.length - 1] }];
  }

  const interior = count - 2;
  const step = (p90 - p10) / interior;
  const boundaries: number[] = [p10];
  for (let i = 1; i <= interior; i++) {
    boundaries.push(Math.round((p10 + step * i) * 100) / 100);
  }

  const buckets: HistogramBucket[] = [];
  buckets.push({ label: '', count: 0, min: -Infinity, max: p10 });
  for (let i = 0; i < interior; i++) {
    buckets.push({ label: '', count: 0, min: boundaries[i], max: boundaries[i + 1] });
  }
  buckets.push({ label: '', count: 0, min: p90, max: Infinity });

  for (const v of values) {
    if (v <= p10) { buckets[0].count++; continue; }
    if (v > p90) { buckets[buckets.length - 1].count++; continue; }
    for (let i = 1; i <= interior; i++) {
      if (v <= boundaries[i]) { buckets[i].count++; break; }
    }
  }

  buckets[0].label = formatBucketLabel(-Infinity, p10, true, false, true);
  for (let i = 1; i <= interior; i++) {
    buckets[i].label = formatBucketLabel(boundaries[i - 1], boundaries[i], false, false, true);
  }
  buckets[buckets.length - 1].label = formatBucketLabel(p90, Infinity, false, true, true);

  return buckets;
}

function getIdentifierValue(
  custId: string,
  identifier: string,
  cohortMap: Map<string, string>,
  attrLookup: Record<string, Record<string, string>>
): string {
  if (identifier === 'Cohort') return cohortMap.get(custId) || 'Unknown';
  return attrLookup[custId]?.[identifier] || 'Unknown';
}

function buildMekko(
  customers: string[],
  xAxis: string,
  yAxis: string | null,
  valueGetter: (cust: string) => number,
  cohortMap: Map<string, string>,
  attrLookup: Record<string, Record<string, string>>
): MekkoData {
  // Group by X
  const xGroups = new Map<string, Map<string, number>>();
  const allYLabels = new Set<string>();

  for (const cust of customers) {
    const xVal = getIdentifierValue(cust, xAxis, cohortMap, attrLookup);
    const yVal = yAxis ? getIdentifierValue(cust, yAxis, cohortMap, attrLookup) : 'Total';
    allYLabels.add(yVal);

    if (!xGroups.has(xVal)) xGroups.set(xVal, new Map());
    const yMap = xGroups.get(xVal)!;
    yMap.set(yVal, (yMap.get(yVal) || 0) + valueGetter(cust));
  }

  const grandTotal = [...xGroups.values()].reduce((sum, yMap) => {
    for (const v of yMap.values()) sum += v;
    return sum;
  }, 0);

  const yLabels = [...allYLabels].sort();

  const columns: MekkoColumn[] = [...xGroups.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([xLabel, yMap]) => {
      let xTotal = 0;
      for (const v of yMap.values()) xTotal += v;

      const stacks: MekkoSegment[] = yLabels.map(yLabel => {
        const value = yMap.get(yLabel) || 0;
        return { yLabel, value, pct: xTotal > 0 ? value / xTotal : 0 };
      });

      return {
        xLabel,
        xTotal,
        xPct: grandTotal > 0 ? xTotal / grandTotal : 0,
        stacks,
      };
    });

  return { columns, yLabels };
}

/** Ensure all grids share the same xLabels, inserting null columns where needed. */
function unifyGridColumns(grids: GridData[]): void {
  const allLabels = new Set<string>();
  for (const g of grids) for (const l of g.xLabels) allLabels.add(l);
  const unified = [...allLabels].sort();

  for (const g of grids) {
    if (g.xLabels.length === unified.length && g.xLabels.every((l, i) => l === unified[i])) continue;
    const oldIdx = new Map<string, number>();
    g.xLabels.forEach((l, i) => oldIdx.set(l, i));
    g.grid = g.grid.map(row => unified.map(l => { const i = oldIdx.get(l); return i != null ? row[i] : null; }));
    g.xTotals = unified.map(l => { const i = oldIdx.get(l); return i != null ? g.xTotals[i] : null; });
    g.xLabels = unified;
  }
}

/** Build grid using ARR-weighted average of per-customer metric values (used for retention). */
function buildWeightedGrid(
  customers: string[],
  xAxis: string,
  yAxis: string | null,
  metricGetter: (cust: string) => number | null,
  arrGetter: (cust: string) => number,
  cohortMap: Map<string, string>,
  attrLookup: Record<string, Record<string, string>>
): GridData {
  const cells = new Map<string, { values: number[]; arrs: number[] }>();
  const xSet = new Set<string>();
  const ySet = new Set<string>();

  for (const cust of customers) {
    const metric = metricGetter(cust);
    if (metric == null) continue;
    const arr = arrGetter(cust);
    const xVal = getIdentifierValue(cust, xAxis, cohortMap, attrLookup);
    const yVal = yAxis ? getIdentifierValue(cust, yAxis, cohortMap, attrLookup) : 'Total';
    xSet.add(xVal);
    ySet.add(yVal);
    const key = `${yVal}|${xVal}`;
    if (!cells.has(key)) cells.set(key, { values: [], arrs: [] });
    cells.get(key)!.values.push(metric);
    cells.get(key)!.arrs.push(arr);
  }

  const xLabels = [...xSet].sort();
  const yLabels = [...ySet].sort();

  function weightedAvg(vals: number[], weights: number[]): number | null {
    if (vals.length === 0) return null;
    const totalWeight = weights.reduce((a, b) => a + b, 0);
    if (totalWeight === 0) return null;
    let sum = 0;
    for (let i = 0; i < vals.length; i++) sum += vals[i] * weights[i];
    return sum / totalWeight;
  }

  const grid: (GridCell | null)[][] = [];
  for (const yLabel of yLabels) {
    const row: (GridCell | null)[] = [];
    for (const xLabel of xLabels) {
      const cell = cells.get(`${yLabel}|${xLabel}`);
      if (!cell || cell.values.length === 0) {
        row.push(null);
      } else {
        const totalARR = cell.arrs.reduce((a, b) => a + b, 0);
        row.push({ value: weightedAvg(cell.values, cell.arrs), count: cell.values.length, totalARR });
      }
    }
    grid.push(row);
  }

  // Column totals
  const xTotals = xLabels.map((_, xi) => {
    const allVals: number[] = [];
    const allArrs: number[] = [];
    for (const yLabel of yLabels) {
      const cell = cells.get(`${yLabel}|${xLabels[xi]}`);
      if (cell) { allVals.push(...cell.values); allArrs.push(...cell.arrs); }
    }
    return weightedAvg(allVals, allArrs);
  });

  // Row totals
  const yTotals = yLabels.map((yLabel) => {
    const allVals: number[] = [];
    const allArrs: number[] = [];
    for (const xLabel of xLabels) {
      const cell = cells.get(`${yLabel}|${xLabel}`);
      if (cell) { allVals.push(...cell.values); allArrs.push(...cell.arrs); }
    }
    return weightedAvg(allVals, allArrs);
  });

  return { xLabels, yLabels, grid, xTotals, yTotals };
}

/** Build grid using aggregate portfolio growth: sum(curr) / sum(prior) - 1 per cell. Matches dashboard methodology. */
function buildAggregateGrowthGrid(
  allCustomers: string[],
  xAxis: string,
  yAxis: string | null,
  currGetter: (cust: string) => number,
  priorGetter: (cust: string) => number,
  cohortMap: Map<string, string>,
  attrLookup: Record<string, Record<string, string>>
): GridData {
  const cells = new Map<string, { custs: Set<string>; sumCurr: number; sumPrior: number }>();
  const xSet = new Set<string>();
  const ySet = new Set<string>();

  for (const cust of allCustomers) {
    const curr = currGetter(cust);
    const prior = priorGetter(cust);
    const xVal = getIdentifierValue(cust, xAxis, cohortMap, attrLookup);
    const yVal = yAxis ? getIdentifierValue(cust, yAxis, cohortMap, attrLookup) : 'Total';
    xSet.add(xVal);
    ySet.add(yVal);
    const key = `${yVal}|${xVal}`;
    if (!cells.has(key)) cells.set(key, { custs: new Set(), sumCurr: 0, sumPrior: 0 });
    const cell = cells.get(key)!;
    cell.custs.add(cust);
    cell.sumCurr += curr;
    cell.sumPrior += prior;
  }

  const xLabels = [...xSet].sort();
  const yLabels = [...ySet].sort();

  const grid: (GridCell | null)[][] = [];
  for (const yLabel of yLabels) {
    const row: (GridCell | null)[] = [];
    for (const xLabel of xLabels) {
      const cell = cells.get(`${yLabel}|${xLabel}`);
      if (!cell || cell.custs.size === 0) {
        row.push(null);
      } else if (cell.sumPrior === 0) {
        // All new logos — can't compute growth
        row.push({ value: null, count: cell.custs.size, totalARR: cell.sumCurr });
      } else {
        row.push({
          value: cell.sumCurr / cell.sumPrior - 1,
          count: cell.custs.size,
          totalARR: cell.sumCurr,
        });
      }
    }
    grid.push(row);
  }

  // Column totals (aggregate across all y-values for each x)
  const xTotals: (number | null)[] = xLabels.map((_, xi) => {
    let sumCurr = 0, sumPrior = 0;
    for (const yLabel of yLabels) {
      const cell = cells.get(`${yLabel}|${xLabels[xi]}`);
      if (cell) { sumCurr += cell.sumCurr; sumPrior += cell.sumPrior; }
    }
    return sumPrior > 0 ? sumCurr / sumPrior - 1 : null;
  });

  // Row totals (aggregate across all x-values for each y)
  const yTotals: (number | null)[] = yLabels.map((yLabel) => {
    let sumCurr = 0, sumPrior = 0;
    for (const xLabel of xLabels) {
      const cell = cells.get(`${yLabel}|${xLabel}`);
      if (cell) { sumCurr += cell.sumCurr; sumPrior += cell.sumPrior; }
    }
    return sumPrior > 0 ? sumCurr / sumPrior - 1 : null;
  });

  return { xLabels, yLabels, grid, xTotals, yTotals };
}

/* ─── Main computation ─── */

export function computeHistogramData(
  wb: Workbook,
  config: EngineConfig,
  granularity?: string,
  filters?: Record<string, string | string[]>,
  mekkoXAxis?: string,
  mekkoYAxis?: string,
  gridXAxis?: string,
  gridYAxis?: string,
): HistogramResult {
  let rawRows = config.input_format === 'cleaned'
    ? readCleanedData(wb, config)
    : readRawData(wb, config);
  const fyMonth = config.fiscal_year_end_month;
  const scaleFactor = config.scale_factor;
  const outputGrans = config.output_granularities;
  const attrNames = Object.keys(config.attributes || {});
  const dataType = config.data_type || 'arr';

  const attributeOptions = computeAttributeOptions(rawRows, attrNames);

  // Apply regular filters (not Cohort)
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

  // Build cohort attribute option
  const cohortValuesSet = new Set(cohortMap.values());
  const cohortValues = periods.filter(p => cohortValuesSet.has(p));
  const allAttributeOptions = [
    { name: 'Cohort', values: cohortValues, multiSelect: true },
    ...attributeOptions,
  ];

  // Apply cohort filter
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

  const latestPeriod = periods[periods.length - 1] || '';
  const derivedPeriods = periods.slice(yoyOffset);
  const latestDerived = derivedPeriods[derivedPeriods.length - 1] || '';

  // Identifiers available for axes
  const identifiers = ['Cohort', ...attrNames];
  const effectiveMekkoX = mekkoXAxis && identifiers.includes(mekkoXAxis) ? mekkoXAxis : 'Cohort';
  const effectiveMekkoY = mekkoYAxis && identifiers.includes(mekkoYAxis) ? mekkoYAxis : (identifiers.length > 1 ? identifiers.find(i => i !== effectiveMekkoX) || null : null);
  const effectiveGridX = gridXAxis && identifiers.includes(gridXAxis) ? gridXAxis : 'Cohort';
  const effectiveGridY = gridYAxis && identifiers.includes(gridYAxis) ? gridYAxis : (identifiers.length > 1 ? identifiers.find(i => i !== effectiveGridX) || null : null);

  // Collect per-customer data for latest period
  const activeCustomers: string[] = [];
  const custLatestARR = new Map<string, number>();
  for (const [cust] of pivot) {
    const arr = getPivotValue(pivot, cust, latestPeriod);
    if (arr > 0) {
      activeCustomers.push(cust);
      custLatestARR.set(cust, arr);
    }
  }

  // Customers active in the prior period (for retention calculations — includes churned customers)
  const priorPeriodCustomers: string[] = [];
  const custPriorARR = new Map<string, number>();
  if (latestDerived) {
    const priorIdx = periods.indexOf(latestDerived) - yoyOffset;
    if (priorIdx >= 0) {
      const priorPeriod = periods[priorIdx];
      for (const [cust] of pivot) {
        const prior = getPivotValue(pivot, cust, priorPeriod);
        if (prior > 0) {
          priorPeriodCustomers.push(cust);
          custPriorARR.set(cust, prior);
        }
      }
    }
  }

  // Per-customer LTM growth rate (only for customers present in both periods)
  const custGrowthRate = new Map<string, number>();
  if (latestDerived && periods.length > yoyOffset) {
    const priorIdx = periods.indexOf(latestDerived) - yoyOffset;
    const priorPeriod = periods[priorIdx];
    for (const cust of activeCustomers) {
      const curr = getPivotValue(pivot, cust, latestDerived);
      const prior = getPivotValue(pivot, cust, priorPeriod);
      if (prior > 0 && curr > 0) {
        custGrowthRate.set(cust, curr / prior - 1);
      }
    }
  }

  // Per-customer net retention & loss-only retention
  const custNetRetention = new Map<string, number>();
  const custLossRetention = new Map<string, number>();
  if (latestDerived) {
    for (const cust of priorPeriodCustomers) {
      const priorIdx = periods.indexOf(latestDerived) - yoyOffset;
      if (priorIdx < 0) continue;
      const priorPeriod = periods[priorIdx];
      const prior = getPivotValue(pivot, cust, priorPeriod);
      if (prior <= 0) continue;
      const churn = derived.churn.get(cust)?.get(latestDerived) || 0;
      const downsell = derived.downsell.get(cust)?.get(latestDerived) || 0;
      const upsell = derived.upsell.get(cust)?.get(latestDerived) || 0;
      // Net retention for this customer = (prior + churn + downsell + upsell) / prior
      custNetRetention.set(cust, (prior + churn + downsell + upsell) / prior);
      // Loss-only = (prior + churn) / prior
      custLossRetention.set(cust, (prior + churn) / prior);
    }
  }

  // A) ARR Histogram
  const arrValues = activeCustomers.map(c => (custLatestARR.get(c) || 0));
  const arrHistogram = buildPercentileBuckets(arrValues, 10);

  // E) Growth Histogram
  const growthValues = [...custGrowthRate.values()];
  const growthHistogram = buildPctBuckets(growthValues, 10);

  // B) Mekko ARR
  const mekkoARR = buildMekko(
    activeCustomers,
    effectiveMekkoX,
    effectiveMekkoY,
    (cust) => (custLatestARR.get(cust) || 0) / sf,
    cohortMap,
    attrLookup
  );

  // C) Mekko Customer Count
  const mekkoCount = buildMekko(
    activeCustomers,
    effectiveMekkoX,
    effectiveMekkoY,
    () => 1,
    cohortMap,
    attrLookup
  );

  // D) Pie charts per identifier
  const pieCharts: PieChartData[] = identifiers.map(id => {
    const arrByVal = new Map<string, number>();
    const countByVal = new Map<string, number>();
    let totalARR = 0;

    for (const cust of activeCustomers) {
      const val = getIdentifierValue(cust, id, cohortMap, attrLookup);
      const arr = (custLatestARR.get(cust) || 0) / sf;
      arrByVal.set(val, (arrByVal.get(val) || 0) + arr);
      countByVal.set(val, (countByVal.get(val) || 0) + 1);
      totalARR += arr;
    }

    const totalCount = activeCustomers.length;

    const arrSlices: PieSlice[] = [...arrByVal.entries()]
      .sort(([, a], [, b]) => b - a)
      .map(([label, value]) => ({ label, value, pct: totalARR > 0 ? value / totalARR : 0 }));

    const countSlices: PieSlice[] = [...countByVal.entries()]
      .sort(([, a], [, b]) => b - a)
      .map(([label, value]) => ({ label, value, pct: totalCount > 0 ? value / totalCount : 0 }));

    return { identifierName: id, arrSlices, countSlices };
  });

  // F) Growth grid — aggregate portfolio growth matching dashboard methodology
  const priorPeriodForGrowth = (latestDerived && periods.indexOf(latestDerived) >= yoyOffset)
    ? periods[periods.indexOf(latestDerived) - yoyOffset] : '';
  const allRelevantCustomers = [...new Set([...activeCustomers, ...priorPeriodCustomers])];
  const growthGrid = buildAggregateGrowthGrid(
    allRelevantCustomers,
    effectiveGridX,
    effectiveGridY,
    (cust) => getPivotValue(pivot, cust, latestPeriod),
    (cust) => priorPeriodForGrowth ? getPivotValue(pivot, cust, priorPeriodForGrowth) : 0,
    cohortMap,
    attrLookup
  );

  // G) Net retention grid (ARR-weighted)
  const custsWithRetention = priorPeriodCustomers.filter(c => custNetRetention.has(c));
  const netRetentionGrid = buildWeightedGrid(
    custsWithRetention,
    effectiveGridX,
    effectiveGridY,
    (cust) => custNetRetention.get(cust) ?? null,
    (cust) => (custPriorARR.get(cust) || 0) / sf,
    cohortMap,
    attrLookup
  );

  const lossRetentionGrid = buildWeightedGrid(
    custsWithRetention,
    effectiveGridX,
    effectiveGridY,
    (cust) => custLossRetention.get(cust) ?? null,
    (cust) => (custPriorARR.get(cust) || 0) / sf,
    cohortMap,
    attrLookup
  );

  // Unify columns across all three grids so they line up
  unifyGridColumns([growthGrid, netRetentionGrid, lossRetentionGrid]);

  return {
    arrHistogram,
    growthHistogram,
    mekkoARR,
    mekkoCount,
    pieCharts,
    growthGrid,
    netRetentionGrid,
    lossRetentionGrid,
    identifiers,
    available_granularities: available,
    granularity: targetGran,
    scale_factor: scaleFactor,
    data_type: dataType,
    attribute_options: allAttributeOptions,
    latestPeriodLabel: latestDerived || latestPeriod,
    priorPeriodLabel: priorPeriodForGrowth,
  };
}
