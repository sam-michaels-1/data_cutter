/**
 * Dashboard computation engine.
 * Port of backend/services/compute.py - replaces pandas with plain TypeScript.
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig } from './types';

function colNumFromLetter(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.toUpperCase().charCodeAt(i) - 64);
  }
  return result;
}

interface RawRecord {
  date: Date;
  customer_id: string;
  arr: number;
  [key: string]: unknown;
}

interface AggRecord {
  customer_id: string;
  period_label: string;
  period_sort: number;
  arr: number;
}

function readRawData(wb: Workbook, config: EngineConfig): RawRecord[] {
  const ws = wb.getWorksheet(config.raw_data_sheet);
  if (!ws) return [];

  const dateIdx = colNumFromLetter(config.date_col);
  const custIdx = colNumFromLetter(config.customer_id_col);
  const arrIdx = colNumFromLetter(config.arr_col);
  const attrCols: Record<string, number> = {};
  for (const [name, letter] of Object.entries(config.attributes || {})) {
    attrCols[name] = colNumFromLetter(letter);
  }

  const rows: RawRecord[] = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < 2) return;
    const dateVal = row.getCell(dateIdx).value;
    const custVal = row.getCell(custIdx).value;
    const arrVal = row.getCell(arrIdx).value;

    if (dateVal == null || custVal == null) return;
    if (!(dateVal instanceof Date)) return;

    const arr = typeof arrVal === 'number' ? arrVal : (arrVal ? parseFloat(String(arrVal)) || 0 : 0);
    const record: RawRecord = { date: dateVal, customer_id: String(custVal).trim(), arr };

    for (const [name, colIdx] of Object.entries(attrCols)) {
      const val = row.getCell(colIdx).value;
      record[name] = val != null ? String(val).trim() : '';
    }

    rows.push(record);
  });

  return rows;
}

function assignFiscalPeriod(date: Date, fyMonth: number): { year: number; quarter: number } {
  const month = date.getMonth() + 1;
  const fiscalYear = month > fyMonth ? date.getFullYear() + 1 : date.getFullYear();
  const fiscalQuarter = Math.floor(((month - (fyMonth + 1) + 12) % 12) / 3) + 1;
  return { year: fiscalYear, quarter: fiscalQuarter };
}

function aggregateToGranularity(
  rawRows: RawRecord[], granularity: string, fyMonth: number, dataType: string
): { records: AggRecord[]; periodDateMap: Record<string, string> } {
  const records: AggRecord[] = [];
  const periodDateMap: Record<string, string> = {};

  if (granularity === 'annual') {
    // Group by customer_id + fy_year
    const groups = new Map<string, { arr: number; maxDate: Date }>();
    for (const row of rawRows) {
      const { year } = assignFiscalPeriod(row.date, fyMonth);
      const isAtFyEnd = row.date.getMonth() + 1 === fyMonth;
      if (dataType !== 'revenue' && !isAtFyEnd) continue;

      const key = `${row.customer_id}|${year}`;
      const existing = groups.get(key);
      if (existing) {
        existing.arr += row.arr;
        if (row.date > existing.maxDate) existing.maxDate = row.date;
      } else {
        groups.set(key, { arr: row.arr, maxDate: row.date });
      }
    }

    const periodMaxDates = new Map<string, Date>();
    for (const [key, val] of groups) {
      const [custId, yearStr] = key.split('|');
      const year = parseInt(yearStr);
      const label = `FY'${(year % 100).toString().padStart(2, '0')}`;
      records.push({ customer_id: custId, period_label: label, period_sort: year, arr: val.arr });

      const existing = periodMaxDates.get(label);
      if (!existing || val.maxDate > existing) periodMaxDates.set(label, val.maxDate);
    }
    for (const [label, date] of periodMaxDates) {
      periodDateMap[label] = date.toISOString().split('T')[0];
    }

  } else if (granularity === 'quarterly') {
    const groups = new Map<string, { arr: number; maxDate: Date }>();
    for (const row of rawRows) {
      const { year, quarter } = assignFiscalPeriod(row.date, fyMonth);
      const isAtQEnd = (fyMonth - (row.date.getMonth() + 1) + 12) % 3 === 0;
      if (dataType !== 'revenue' && !isAtQEnd) continue;

      const key = `${row.customer_id}|${year}|${quarter}`;
      const existing = groups.get(key);
      if (existing) {
        existing.arr += row.arr;
        if (row.date > existing.maxDate) existing.maxDate = row.date;
      } else {
        groups.set(key, { arr: row.arr, maxDate: row.date });
      }
    }

    const periodMaxDates = new Map<string, Date>();
    for (const [key, val] of groups) {
      const [custId, yearStr, qStr] = key.split('|');
      const year = parseInt(yearStr);
      const quarter = parseInt(qStr);
      const label = `Q${quarter}'${(year % 100).toString().padStart(2, '0')}`;
      const sort = year * 10 + quarter;
      records.push({ customer_id: custId, period_label: label, period_sort: sort, arr: val.arr });

      const existing = periodMaxDates.get(label);
      if (!existing || val.maxDate > existing) periodMaxDates.set(label, val.maxDate);
    }
    for (const [label, date] of periodMaxDates) {
      periodDateMap[label] = date.toISOString().split('T')[0];
    }

  } else {
    // monthly
    const groups = new Map<string, { arr: number; maxDate: Date }>();
    for (const row of rawRows) {
      const dateKey = `${row.date.getFullYear()}-${(row.date.getMonth() + 1).toString().padStart(2, '0')}`;
      const key = `${row.customer_id}|${dateKey}`;
      const existing = groups.get(key);
      if (existing) {
        existing.arr += row.arr;
        if (row.date > existing.maxDate) existing.maxDate = row.date;
      } else {
        groups.set(key, { arr: row.arr, maxDate: row.date });
      }
    }

    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const periodMaxDates = new Map<string, Date>();
    for (const [key, val] of groups) {
      const [custId, dateKey] = key.split('|');
      const [yearStr, monthStr] = dateKey.split('-');
      const year = parseInt(yearStr);
      const month = parseInt(monthStr);
      const label = `${monthNames[month - 1]} '${(year % 100).toString().padStart(2, '0')}`;
      const sort = year * 100 + month;
      records.push({ customer_id: custId, period_label: label, period_sort: sort, arr: val.arr });

      const existing = periodMaxDates.get(label);
      if (!existing || val.maxDate > existing) periodMaxDates.set(label, val.maxDate);
    }
    for (const [label, date] of periodMaxDates) {
      periodDateMap[label] = date.toISOString().split('T')[0];
    }
  }

  return { records, periodDateMap };
}

function buildArrMatrix(records: AggRecord[]): { pivot: Map<string, Map<string, number>>; periods: string[] } {
  // Get sorted periods
  const periodSet = new Map<string, number>();
  for (const r of records) {
    const existing = periodSet.get(r.period_label);
    if (existing == null || r.period_sort < existing) {
      periodSet.set(r.period_label, r.period_sort);
    }
  }
  const periods = [...periodSet.entries()].sort((a, b) => a[1] - b[1]).map(e => e[0]);

  // Build pivot: customer_id -> period_label -> arr
  const pivot = new Map<string, Map<string, number>>();
  for (const r of records) {
    let custMap = pivot.get(r.customer_id);
    if (!custMap) {
      custMap = new Map();
      pivot.set(r.customer_id, custMap);
    }
    custMap.set(r.period_label, (custMap.get(r.period_label) || 0) + r.arr);
  }

  return { pivot, periods };
}

function getYoyOffset(granularity: string): number {
  return { monthly: 12, quarterly: 4, annual: 1 }[granularity] || 1;
}

function getPivotValue(pivot: Map<string, Map<string, number>>, cust: string, period: string): number {
  return pivot.get(cust)?.get(period) || 0;
}

interface DerivedData {
  churn: Map<string, Map<string, number>>;
  downsell: Map<string, Map<string, number>>;
  upsell: Map<string, Map<string, number>>;
  new_biz: Map<string, Map<string, number>>;
}

function computeDerived(pivot: Map<string, Map<string, number>>, periods: string[], yoyOffset: number): DerivedData {
  const result: DerivedData = { churn: new Map(), downsell: new Map(), upsell: new Map(), new_biz: new Map() };
  const numDerived = periods.length - yoyOffset;
  if (numDerived <= 0) return result;

  for (const [cust] of pivot) {
    const churnMap = new Map<string, number>();
    const dsMap = new Map<string, number>();
    const upMap = new Map<string, number>();
    const nbMap = new Map<string, number>();

    for (let i = 0; i < numDerived; i++) {
      const priorPeriod = periods[i];
      const currPeriod = periods[yoyOffset + i];
      const prior = getPivotValue(pivot, cust, priorPeriod);
      const curr = getPivotValue(pivot, cust, currPeriod);

      churnMap.set(currPeriod, (curr === 0 && prior > 0) ? -prior : 0);
      dsMap.set(currPeriod, (curr > 0 && prior > 0 && curr < prior) ? curr - prior : 0);
      upMap.set(currPeriod, (curr > 0 && prior > 0 && curr > prior) ? curr - prior : 0);
      nbMap.set(currPeriod, (curr > 0 && prior === 0) ? curr : 0);
    }

    result.churn.set(cust, churnMap);
    result.downsell.set(cust, dsMap);
    result.upsell.set(cust, upMap);
    result.new_biz.set(cust, nbMap);
  }

  return result;
}

function buildCohortMap(pivot: Map<string, Map<string, number>>, periods: string[]): Map<string, string> {
  const cohortMap = new Map<string, string>();
  for (const [cust] of pivot) {
    for (const p of periods) {
      if (getPivotValue(pivot, cust, p) > 0) {
        cohortMap.set(cust, p);
        break;
      }
    }
  }
  return cohortMap;
}

function sumDerivedPeriod(derived: Map<string, Map<string, number>>, period: string): number {
  let total = 0;
  for (const [, m] of derived) {
    total += m.get(period) || 0;
  }
  return total;
}

function computeAttributeOptions(rawRows: RawRecord[], attrNames: string[]): { name: string; values: string[] }[] {
  const options: { name: string; values: string[] }[] = [];
  for (const name of attrNames) {
    const vals = new Set<string>();
    for (const row of rawRows) {
      const v = row[name];
      if (v && typeof v === 'string' && v !== '') vals.add(v);
    }
    options.push({ name, values: [...vals].sort() });
  }
  return options;
}

function buildAttrLookup(rawRows: RawRecord[], attrNames: string[]): Record<string, Record<string, string>> {
  if (attrNames.length === 0) return {};
  const lookup: Record<string, Record<string, string>> = {};
  for (const row of rawRows) {
    if (lookup[row.customer_id]) continue;
    const attrs: Record<string, string> = {};
    for (const name of attrNames) {
      attrs[name] = row[name] ? String(row[name]) : '';
    }
    lookup[row.customer_id] = attrs;
  }
  return lookup;
}

export interface DashboardResult {
  overview: {
    periods: string[];
    arr_over_time: number[];
    arr_growth_pcts: (number | null)[];
    waterfall: {
      period_label: string;
      bop: number;
      new_logo: number;
      upsell: number;
      downsell: number;
      churn: number;
      eop: number;
    } | null;
    stats: {
      total_arr: number;
      customer_count: number;
      net_retention_pct: number | null;
      yoy_growth_pct: number | null;
      lost_only_retention_pct: number | null;
      punitive_retention_pct: number | null;
    };
    top_customers: {
      rank: number;
      name: string;
      arr: number;
      change_pct: number | null;
      pct_of_total: number;
      trend: number[];
      status: string;
      attributes: Record<string, string>;
      cohort: string;
    }[];
    latest_period_label: string;
    latest_period_date: string;
  };
  cohort: {
    periods: string[];
    cohorts: {
      label: string;
      count: number;
      starting_arr: number;
      arr: (number | null)[];
      customers: (number | null)[];
      ndr: (number | null)[];
      logo_retention: (number | null)[];
    }[];
  };
  granularity: string;
  available_granularities: string[];
  scale_factor: number;
  attribute_options: { name: string; values: string[] }[];
  data_type: string;
}

export function computeDashboard(
  wb: Workbook, config: EngineConfig,
  granularity?: string,
  filters?: Record<string, string>,
  topN = 10
): DashboardResult {
  let rawRows = readRawData(wb, config);
  const fyMonth = config.fiscal_year_end_month;
  const scaleFactor = config.scale_factor;
  const outputGrans = config.output_granularities;
  const attrNames = Object.keys(config.attributes || {});
  const dataType = config.data_type || 'arr';

  const attributeOptions = computeAttributeOptions(rawRows, attrNames);

  // Apply filters
  if (filters) {
    for (const [attrName, attrValue] of Object.entries(filters)) {
      if (attrValue) {
        rawRows = rawRows.filter(r => r[attrName] === attrValue);
      }
    }
  }

  // Determine granularity
  let targetGran = granularity && outputGrans.includes(granularity) ? granularity : null;
  if (!targetGran) {
    const prefOrder = ['annual', 'quarterly', 'monthly'];
    targetGran = prefOrder.find(g => outputGrans.includes(g)) || outputGrans[0];
  }

  const available = ['annual', 'quarterly', 'monthly'].filter(g => outputGrans.includes(g));

  const { records, periodDateMap } = aggregateToGranularity(rawRows, targetGran, fyMonth, dataType);
  const { pivot, periods } = buildArrMatrix(records);
  const yoyOffset = getYoyOffset(targetGran);
  const derived = computeDerived(pivot, periods, yoyOffset);
  const cohortMap = buildCohortMap(pivot, periods);
  const attrLookup = buildAttrLookup(rawRows, attrNames);

  const sf = scaleFactor;

  // ARR over time
  const arrOverTime = periods.map(p => {
    let total = 0;
    for (const [, custMap] of pivot) total += custMap.get(p) || 0;
    return Math.round((total / sf) * 100) / 100;
  });

  // Growth
  const arrGrowthPcts: (number | null)[] = arrOverTime.map((val, i) => {
    if (i < yoyOffset) return null;
    const prior = arrOverTime[i - yoyOffset];
    return prior !== 0 ? Math.round((val / prior - 1) * 10000) / 10000 : null;
  });

  const latestPeriodLabel = periods[periods.length - 1] || '';
  const latestPeriodDate = periodDateMap[latestPeriodLabel] || '';

  const derivedPeriods = periods.slice(yoyOffset);

  let waterfall = null;
  let stats = {
    total_arr: arrOverTime[arrOverTime.length - 1] || 0,
    customer_count: 0,
    net_retention_pct: null as number | null,
    yoy_growth_pct: null as number | null,
    lost_only_retention_pct: null as number | null,
    punitive_retention_pct: null as number | null,
  };

  if (periods.length > 0) {
    let count = 0;
    for (const [, custMap] of pivot) {
      if ((custMap.get(periods[periods.length - 1]) || 0) > 0) count++;
    }
    stats.customer_count = count;
  }

  if (derivedPeriods.length > 0) {
    const latest = derivedPeriods[derivedPeriods.length - 1];
    const priorIdx = periods.indexOf(latest) - yoyOffset;
    const priorPeriod = periods[priorIdx];

    let bop = 0;
    for (const [, custMap] of pivot) bop += custMap.get(priorPeriod) || 0;
    bop /= sf;

    const churnTotal = sumDerivedPeriod(derived.churn, latest) / sf;
    const downsellTotal = sumDerivedPeriod(derived.downsell, latest) / sf;
    const upsellTotal = sumDerivedPeriod(derived.upsell, latest) / sf;
    const newLogoTotal = sumDerivedPeriod(derived.new_biz, latest) / sf;
    const retained = bop + churnTotal + downsellTotal + upsellTotal;
    const eop = retained + newLogoTotal;

    waterfall = {
      period_label: latest,
      bop: Math.round(bop * 100) / 100,
      new_logo: Math.round(newLogoTotal * 100) / 100,
      upsell: Math.round(upsellTotal * 100) / 100,
      downsell: Math.round(downsellTotal * 100) / 100,
      churn: Math.round(churnTotal * 100) / 100,
      eop: Math.round(eop * 100) / 100,
    };

    stats.total_arr = Math.round(eop * 100) / 100;
    stats.net_retention_pct = bop !== 0 ? Math.round((retained / bop) * 10000) / 10000 : null;
    stats.yoy_growth_pct = bop !== 0 ? Math.round((eop / bop - 1) * 10000) / 10000 : null;
    stats.lost_only_retention_pct = bop !== 0 ? Math.round(((bop + churnTotal) / bop) * 10000) / 10000 : null;
    stats.punitive_retention_pct = bop !== 0 ? Math.round(((bop + churnTotal + downsellTotal) / bop) * 10000) / 10000 : null;
  }

  // Top customers
  const latestPeriod = periods[periods.length - 1] || '';
  const custArr: [string, number][] = [];
  for (const [cust, custMap] of pivot) {
    custArr.push([cust, custMap.get(latestPeriod) || 0]);
  }
  custArr.sort((a, b) => b[1] - a[1]);
  const topCusts = custArr.slice(0, topN);
  const totalArrRaw = custArr.reduce((sum, [, v]) => sum + v, 0);

  const topCustomers = topCusts.map(([custId, arrVal], idx) => {
    const trend = periods.map(p => Math.round((getPivotValue(pivot, custId, p) / sf) * 100) / 100);
    let changePct: number | null = null;
    if (periods.length >= 2) {
      const prevArr = getPivotValue(pivot, custId, periods[periods.length - 2]) / sf;
      const currArr = arrVal / sf;
      changePct = prevArr > 0 ? Math.round((currArr / prevArr - 1) * 10000) / 10000 : null;
    }
    let status = 'New';
    if (changePct != null) {
      status = changePct > 0.05 ? 'Growth' : changePct < -0.05 ? 'Declining' : 'Stable';
    }

    return {
      rank: idx + 1,
      name: custId,
      arr: Math.round((arrVal / sf) * 100) / 100,
      change_pct: changePct,
      pct_of_total: totalArrRaw > 0 ? Math.round((arrVal / totalArrRaw) * 10000) / 10000 : 0,
      trend,
      status,
      attributes: attrLookup[custId] || {},
      cohort: cohortMap.get(custId) || '',
    };
  });

  // Cohort
  const uniqueCohorts = periods.filter(p => {
    for (const [, label] of cohortMap) {
      if (label === p) return true;
    }
    return false;
  });

  const cohorts = uniqueCohorts.map(cohortLabel => {
    const custsInCohort: string[] = [];
    for (const [cust, label] of cohortMap) {
      if (label === cohortLabel) custsInCohort.push(cust);
    }

    const cohortIdx = periods.indexOf(cohortLabel);
    const arrValues = periods.map((p, i) => {
      if (i < cohortIdx) return null;
      let total = 0;
      for (const c of custsInCohort) total += getPivotValue(pivot, c, p);
      return Math.round((total / sf) * 100) / 100;
    });

    const custCounts = periods.map((p, i) => {
      if (i < cohortIdx) return null;
      let count = 0;
      for (const c of custsInCohort) if (getPivotValue(pivot, c, p) > 0) count++;
      return count;
    });

    let startingArr = 0;
    for (const c of custsInCohort) startingArr += getPivotValue(pivot, c, cohortLabel);
    let startingCount = 0;
    for (const c of custsInCohort) if (getPivotValue(pivot, c, cohortLabel) > 0) startingCount++;

    const ndr = periods.map((p, i) => {
      if (i < cohortIdx || startingArr === 0) return null;
      let total = 0;
      for (const c of custsInCohort) total += getPivotValue(pivot, c, p);
      return Math.round((total / startingArr) * 10000) / 10000;
    });

    const logoRet = periods.map((p, i) => {
      if (i < cohortIdx || startingCount === 0) return null;
      let count = 0;
      for (const c of custsInCohort) if (getPivotValue(pivot, c, p) > 0) count++;
      return Math.round((count / startingCount) * 10000) / 10000;
    });

    return {
      label: cohortLabel,
      count: startingCount,
      starting_arr: Math.round((startingArr / sf) * 100) / 100,
      arr: arrValues,
      customers: custCounts,
      ndr,
      logo_retention: logoRet,
    };
  });

  return {
    overview: {
      periods,
      arr_over_time: arrOverTime,
      arr_growth_pcts: arrGrowthPcts,
      waterfall,
      stats,
      top_customers: topCustomers,
      latest_period_label: latestPeriodLabel,
      latest_period_date: latestPeriodDate,
    },
    cohort: { periods, cohorts },
    granularity: targetGran,
    available_granularities: available,
    scale_factor: scaleFactor,
    attribute_options: attributeOptions,
    data_type: dataType,
  };
}
