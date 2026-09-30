/**
 * Per-customer deep-dive computation.
 * Built on the same aggregation pipeline as compute.ts — no duplicated
 * aggregation logic (readRawData/readCleanedData, aggregateToGranularity,
 * buildArrMatrix are reused directly).
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig } from './types';
import type {
  AttributeTransition,
  CustomerDetail,
  CustomerEventType,
  CustomerListEntry,
  CustomerMovement,
  CustomerPeriodEntry,
} from '../types/dashboard';
import {
  readRawData,
  readCleanedData,
  aggregateToGranularity,
  buildArrMatrix,
  buildCohortMap,
  buildAttrLookup,
  getPivotValue,
  getYoyOffset,
  periodLabelForDate,
} from './compute';
import { toLocalISODate } from './utils';

function resolveGranularity(config: EngineConfig, granularity?: string): { target: string; available: string[] } {
  const outputGrans = config.output_granularities;
  const available = ['annual', 'quarterly', 'monthly'].filter(g => outputGrans.includes(g));
  let target = granularity && outputGrans.includes(granularity) ? granularity : null;
  if (!target) {
    const prefOrder = ['annual', 'quarterly', 'monthly'];
    target = prefOrder.find(g => outputGrans.includes(g)) || outputGrans[0];
  }
  return { target, available };
}

function loadRows(wb: Workbook, config: EngineConfig) {
  return config.input_format === 'cleaned'
    ? readCleanedData(wb, config)
    : readRawData(wb, config);
}

/** All customers in the workbook, sorted alphabetically, with latest-period ARR. */
export function computeCustomers(wb: Workbook, config: EngineConfig): CustomerListEntry[] {
  const rawRows = loadRows(wb, config);
  const { target } = resolveGranularity(config);
  const { records } = aggregateToGranularity(
    rawRows, target, config.fiscal_year_end_month, config.data_type || 'arr'
  );
  const { pivot, periods } = buildArrMatrix(records);
  const latest = periods[periods.length - 1] || '';
  const sf = config.scale_factor;

  // Names come from the raw rows, not the pivot: for non-revenue data the
  // aggregation drops rows not at a period boundary, which would hide
  // customers whose records all fall mid-period.
  const names = new Set<string>();
  for (const r of rawRows) names.add(r.customer_id);

  const list: CustomerListEntry[] = [...names].map((name) => ({
    name,
    current_arr: Math.round((getPivotValue(pivot, name, latest) / sf) * 100) / 100,
  }));
  list.sort((a, b) => a.name.localeCompare(b.name));
  return list;
}

export function computeCustomerDetail(
  wb: Workbook, config: EngineConfig,
  customerName: string, granularity?: string
): CustomerDetail | null {
  const rawRows = loadRows(wb, config);
  const fyMonth = config.fiscal_year_end_month;
  const dataType = config.data_type || 'arr';
  const sf = config.scale_factor;
  const attrNames = Object.keys(config.attributes || {});
  const resolved = resolveGranularity(config, granularity);
  let target = resolved.target;
  const available = resolved.available;

  let { records, periodDateMap } = aggregateToGranularity(rawRows, target, fyMonth, dataType);
  let { pivot, periods } = buildArrMatrix(records);

  // For non-revenue data, aggregation keeps only period-boundary rows, so a
  // customer with only mid-period records is absent at coarser granularities.
  // Fall back to the finest available granularity before giving up.
  if (!pivot.has(customerName)) {
    const finest = ['monthly', 'quarterly', 'annual'].find(g => available.includes(g));
    if (!finest || finest === target) return null;
    target = finest;
    ({ records, periodDateMap } = aggregateToGranularity(rawRows, target, fyMonth, dataType));
    ({ pivot, periods } = buildArrMatrix(records));
    if (!pivot.has(customerName) || periods.length === 0) return null;
  }
  if (periods.length === 0) return null;

  const yoyOffset = getYoyOffset(target);
  const cohortMap = buildCohortMap(pivot, periods);
  const attrLookup = buildAttrLookup(rawRows, attrNames);
  const custMap = pivot.get(customerName)!;

  // Period-over-period timeline with movement events
  const timeline: CustomerPeriodEntry[] = [];
  const movements: CustomerMovement[] = [];
  let seenNonZero = false;
  let prevRaw: number | null = null;
  for (const p of periods) {
    const raw = custMap.get(p) || 0;
    const arr = Math.round((raw / sf) * 100) / 100;
    let change: number | null = null;
    let changePct: number | null = null;
    let event: CustomerEventType = 'inactive';

    if (prevRaw != null) {
      change = Math.round(((raw - prevRaw) / sf) * 100) / 100;
      changePct = prevRaw !== 0 ? Math.round((raw / prevRaw - 1) * 10000) / 10000 : null;
      if (prevRaw === 0 && raw > 0) event = seenNonZero ? 'reactivation' : 'new';
      else if (prevRaw > 0 && raw === 0) event = 'churn';
      else if (raw > prevRaw) event = 'upsell';
      else if (raw < prevRaw) event = 'downsell';
      else event = 'flat';
    } else if (raw > 0) {
      event = 'new';
    }

    if (raw > 0) seenNonZero = true;
    timeline.push({ period_label: p, arr, change, change_pct: changePct, event });
    if (event !== 'flat' && event !== 'inactive') {
      movements.push({ period_label: p, type: event, amount: change ?? arr });
    }
    prevRaw = raw;
  }

  // Headline stats
  const latestLabel = periods[periods.length - 1];
  const currentRaw = custMap.get(latestLabel) || 0;
  const currentArr = Math.round((currentRaw / sf) * 100) / 100;

  // First non-zero period by raw value — a tiny ARR can round to 0 after
  // scaling, so checking the rounded timeline would skip it.
  const firstIdx = periods.findIndex(p => (custMap.get(p) || 0) > 0);
  const firstRaw = firstIdx >= 0 ? (custMap.get(periods[firstIdx]) || 0) : 0;

  let peakRaw = 0;
  let peakLabel = '';
  for (const p of periods) {
    const v = custMap.get(p) || 0;
    if (v > peakRaw) { peakRaw = v; peakLabel = p; }
  }

  let lifetimeRaw = 0;
  for (const p of periods) lifetimeRaw += custMap.get(p) || 0;

  let totalRawAll = 0;
  for (const [, m] of pivot) totalRawAll += m.get(latestLabel) || 0;

  // Latest non-empty attribute values (raw-format attributes can change over
  // time, so the first row's value may be stale); cleaned data is constant.
  let headerAttributes = attrLookup[customerName] || {};
  if (config.input_format !== 'cleaned' && attrNames.length > 0) {
    const sortedRows = rawRows
      .filter(r => r.customer_id === customerName)
      .sort((a, b) => a.date.getTime() - b.date.getTime());
    headerAttributes = {};
    for (const name of attrNames) {
      headerAttributes[name] = '';
      for (const row of sortedRows) {
        const v = row[name] ? String(row[name]) : '';
        if (v !== '') headerAttributes[name] = v;
      }
    }
  }

  const totalChangeRaw = firstIdx >= 0 ? currentRaw - firstRaw : null;
  const totalChangePct = firstRaw > 0 ? Math.round((currentRaw / firstRaw - 1) * 10000) / 10000 : null;

  let yoyChangePct: number | null = null;
  const yoyPriorIdx = periods.length - 1 - yoyOffset;
  if (yoyPriorIdx >= 0) {
    const priorRaw = custMap.get(periods[yoyPriorIdx]) || 0;
    if (priorRaw > 0) yoyChangePct = Math.round((currentRaw / priorRaw - 1) * 10000) / 10000;
  }

  // CAGR uses real elapsed time between the first non-zero period and the
  // latest period (periodDateMap holds each period's max date), so gaps in
  // the calendar don't compress the span and inflate the rate.
  let cagr: number | null = null;
  if (firstIdx >= 0 && firstRaw > 0 && currentRaw > 0) {
    const firstDate = periodDateMap[periods[firstIdx]];
    const lastDate = periodDateMap[latestLabel];
    if (firstDate && lastDate) {
      const years =
        (new Date(lastDate + 'T00:00:00').getTime() - new Date(firstDate + 'T00:00:00').getTime()) /
        (365.25 * 24 * 60 * 60 * 1000);
      if (years >= 1) {
        cagr = Math.round((Math.pow(currentRaw / firstRaw, 1 / years) - 1) * 10000) / 10000;
      }
    }
  }

  // Status — same rules as the top-customers computation in compute.ts
  let status = 'New';
  if (periods.length >= 2) {
    const prevArr = custMap.get(periods[periods.length - 2]) || 0;
    if (prevArr > 0) {
      const chg = currentRaw / prevArr - 1;
      status = chg > 0.05 ? 'Growth' : chg < -0.05 ? 'Declining' : 'Stable';
    }
  }

  // Attribute transitions (raw-format data only; cleaned rows carry one value per customer)
  const transitions: AttributeTransition[] = [];
  if (config.input_format !== 'cleaned' && attrNames.length > 0) {
    const custRows = rawRows
      .filter(r => r.customer_id === customerName)
      .sort((a, b) => a.date.getTime() - b.date.getTime());
    for (const name of attrNames) {
      let last: string | null = null;
      for (const row of custRows) {
        const v = row[name] ? String(row[name]) : '';
        if (last == null) { last = v; continue; }
        if (v !== last) {
          transitions.push({
            attribute: name,
            from: last,
            to: v,
            date: toLocalISODate(row.date),
            period_label: periodLabelForDate(row.date, target, fyMonth),
          });
          last = v;
        }
      }
    }
  }

  return {
    name: customerName,
    attributes: headerAttributes,
    cohort: cohortMap.get(customerName) || '',
    status,
    current_arr: currentArr,
    first_arr: Math.round((firstRaw / sf) * 100) / 100,
    first_period_label: firstIdx >= 0 ? periods[firstIdx] : '',
    peak_arr: Math.round((peakRaw / sf) * 100) / 100,
    peak_period_label: peakLabel,
    lifetime_total: Math.round((lifetimeRaw / sf) * 100) / 100,
    pct_of_total: totalRawAll > 0 ? Math.round((currentRaw / totalRawAll) * 10000) / 10000 : 0,
    total_change: totalChangeRaw != null ? Math.round((totalChangeRaw / sf) * 100) / 100 : null,
    total_change_pct: totalChangePct,
    yoy_change_pct: yoyChangePct,
    cagr,
    timeline,
    movements,
    attribute_transitions: transitions,
    granularity: target,
    available_granularities: available,
    scale_factor: sf,
    data_type: dataType,
  };
}
