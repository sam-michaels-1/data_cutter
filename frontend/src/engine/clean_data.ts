/**
 * Clean Data tab generators.
 * Port of data-pack-app/engine/clean_data.py for client-side use with ExcelJS.
 */
import type { Workbook } from 'exceljs';
import type { EngineConfig, CleanTabResult } from './types';
import { colLetter, computeCleanLayout, makeRange, sumifs, getYoyOffset, type CleanLayout } from './utils';

/**
 * Generate the base clean data tab at the raw data's native granularity.
 */
export function generateBaseCleanData(
  wb: Workbook, config: EngineConfig,
  uniqueDates: Date[], uniqueCustomers: string[]
): CleanTabResult {
  const granularity = config.time_granularity;
  const sheetName = `Clean ${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Data`;
  const ws = wb.addWorksheet(sheetName);

  const rawSheet = config.raw_data_sheet;
  const rawFirst = config.raw_data_first_row;
  const rawLast = config.raw_data_last_row;
  const rawCustCol = config.customer_id_col;
  const rawDateCol = config.date_col;
  const rawArrCol = config.arr_col;
  const attrs = config.attributes;
  const attrNames = Object.keys(attrs);
  const numAttrs = attrNames.length;
  const numDates = uniqueDates.length;
  const numCustomers = uniqueCustomers.length;
  const yoyOffset = getYoyOffset(granularity);
  const fyMonth = config.fiscal_year_end_month;

  const layout = computeCleanLayout(numAttrs, numDates, yoyOffset);
  const firstDataRow = 7;
  const lastDataRow = firstDataRow + numCustomers - 1;

  // --- Row 1: Fiscal month end + column totals ---
  ws.getCell(1, 1).value = 'Fiscal Month End:';
  ws.getCell(1, 2).value = { formula: 'Control!C7' };

  for (const sectionKey of ['arr', 'churn', 'downsell', 'upsell', 'new_biz'] as const) {
    const start = layout[`${sectionKey}_start`];
    const end = sectionKey === 'arr' ? layout.arr_end : layout[`${sectionKey}_end`];
    for (let c = start; c <= end; c++) {
      const cl = colLetter(c);
      ws.getCell(1, c).value = { formula: `SUM(${cl}${firstDataRow}:${cl}${lastDataRow})` };
    }
  }

  // --- Row 2: Quarter calculation ---
  ws.getCell(2, layout.label).value = 'Quarter';
  const dataType = config.data_type || 'arr';

  if (granularity === 'monthly') {
    for (let i = 0; i < numDates; i++) {
      const c = layout.arr_start + i;
      const cl = colLetter(c);
      if (dataType === 'revenue') {
        ws.getCell(2, c).value = { formula: `INT(MOD(MONTH(${cl}6)-(Control!$C$7+1),12)/3)+1` };
      } else {
        ws.getCell(2, c).value = { formula: `IF(MOD($B$1-MONTH(${cl}6),3)=0,INT(MOD(MONTH(${cl}6)-(Control!$C$7+1),12)/3)+1,0)` };
      }
    }
  } else if (granularity === 'quarterly') {
    const firstDate = uniqueDates[0];
    const month = firstDate.getMonth() + 1;
    const fq = Math.floor(((month - (fyMonth + 1) + 12) % 12) / 3) + 1;
    ws.getCell(2, layout.arr_start).value = fq;
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(2, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,1,${prevCl}2+1)` };
    }
  } else if (granularity === 'annual') {
    ws.getCell(2, layout.arr_start).value = 4;
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(2, layout.arr_start + i).value = { formula: `${prevCl}2` };
    }
  }

  copyHelperRowsToDerived(ws, layout, 2);

  // --- Row 3: Year calculation ---
  ws.getCell(3, layout.label).value = 'Year';

  if (granularity === 'monthly') {
    for (let i = 0; i < numDates; i++) {
      const c = layout.arr_start + i;
      const cl = colLetter(c);
      ws.getCell(3, c).value = { formula: `IF(${cl}2>0,YEAR(${cl}6),0)` };
    }
  } else if (granularity === 'quarterly') {
    const firstDate = uniqueDates[0];
    const firstMonth = firstDate.getMonth() + 1;
    const firstYear = firstMonth > fyMonth ? firstDate.getFullYear() + 1 : firstDate.getFullYear();
    ws.getCell(3, layout.arr_start).value = firstYear;
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(3, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,${prevCl}3+1,${prevCl}3)` };
    }
  } else if (granularity === 'annual') {
    const firstDate = uniqueDates[0];
    ws.getCell(3, layout.arr_start).value = firstDate.getFullYear();
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(3, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,${prevCl}3+1,${prevCl}3)` };
    }
  }

  copyHelperRowsToDerived(ws, layout, 3);

  // --- Row 4: Month row (date values for SUMIFS matching) ---
  if (granularity === 'quarterly' || granularity === 'annual') {
    ws.getCell(4, layout.label).value = 'Month';
    for (let i = 0; i < numDates; i++) {
      ws.getCell(4, layout.arr_start + i).value = uniqueDates[i];
    }
    copyHelperRowsToDerived(ws, layout, 4);
  }

  // --- Row 5: Section headers ---
  ws.getCell(5, layout.attr_start).value = 'Customer Identifying Information';
  const metricLabel = dataType === 'arr' ? 'ARR' : 'Revenue';
  ws.getCell(5, layout.arr_start).value = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} ${metricLabel} by Date`;
  ws.getCell(5, layout.churn_start).value = 'Churn?';
  ws.getCell(5, layout.downsell_start).value = 'Downsell?';
  ws.getCell(5, layout.upsell_start).value = 'Upsell?';
  ws.getCell(5, layout.new_biz_start).value = 'New Business Dollars?';

  // --- Row 6: Column headers ---
  ws.getCell(6, layout.cust_id).value = 'Customer ID';
  for (let i = 0; i < attrNames.length; i++) {
    ws.getCell(6, layout.attr_start + i).value = attrNames[i];
  }
  ws.getCell(6, layout.cohort).value = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Cohort`;
  ws.getCell(6, layout.rank).value = 'Customer Rank #';

  // Date headers in ARR section
  for (let i = 0; i < numDates; i++) {
    const c = layout.arr_start + i;
    const cl = colLetter(c);
    if (granularity === 'quarterly') {
      ws.getCell(6, c).value = { formula: `"Q"&${cl}2&"'"&RIGHT(${cl}3,2)` };
    } else if (granularity === 'annual') {
      ws.getCell(6, c).value = { formula: `"FY"&"'"&RIGHT(${cl}3,2)` };
    } else {
      ws.getCell(6, c).value = uniqueDates[i];
    }
  }

  // Derived section date headers
  for (let i = 0; i < layout.num_derived; i++) {
    const arrRefCol = layout.arr_start + yoyOffset + i;
    const arrRefCl = colLetter(arrRefCol);

    const churnCol = layout.churn_start + i;
    ws.getCell(6, churnCol).value = { formula: `${arrRefCl}6` };

    const downsellCol = layout.downsell_start + i;
    ws.getCell(6, downsellCol).value = { formula: `${colLetter(churnCol)}6` };

    const upsellCol = layout.upsell_start + i;
    ws.getCell(6, upsellCol).value = { formula: `${colLetter(downsellCol)}6` };

    const newBizCol = layout.new_biz_start + i;
    ws.getCell(6, newBizCol).value = { formula: `${colLetter(upsellCol)}6` };
  }

  // --- Customer data rows (7+) ---
  const rawArrRange = makeRange(rawSheet, rawArrCol, rawFirst, rawLast);
  const rawDateRange = makeRange(rawSheet, rawDateCol, rawFirst, rawLast);
  const rawCustRange = makeRange(rawSheet, rawCustCol, rawFirst, rawLast);

  const arrStartCl = colLetter(layout.arr_start);
  const arrEndCl = colLetter(layout.arr_end);
  const dateRefRow = (granularity === 'quarterly' || granularity === 'annual') ? 4 : 6;

  for (let idx = 0; idx < uniqueCustomers.length; idx++) {
    const custId = uniqueCustomers[idx];
    const row = firstDataRow + idx;
    const custIdCl = colLetter(layout.cust_id);

    // Column B: Customer ID
    ws.getCell(row, layout.cust_id).value = custId;

    // Attribute lookups via XLOOKUP
    for (let attrIdx = 0; attrIdx < attrNames.length; attrIdx++) {
      const rawAttrCol = attrs[attrNames[attrIdx]];
      const attrCol = layout.attr_start + attrIdx;
      ws.getCell(row, attrCol).value = {
        formula: `_xlfn.XLOOKUP($${custIdCl}${row},'${rawSheet}'!$${rawCustCol}$${rawFirst}:$${rawCustCol}$${rawLast},'${rawSheet}'!$${rawAttrCol}$${rawFirst}:$${rawAttrCol}$${rawLast})`
      };
    }

    // Cohort (first non-zero ARR period)
    ws.getCell(row, layout.cohort).value = {
      formula: `IFERROR(INDEX($${arrStartCl}$6:$${arrEndCl}$6,MATCH(TRUE,INDEX(${arrStartCl}${row}:${arrEndCl}${row}<>0,0),0)),"n.a.")`
    };

    // Rank
    ws.getCell(row, layout.rank).value = {
      formula: `RANK(${arrEndCl}${row},$${arrEndCl}$${firstDataRow}:$${arrEndCl}$${lastDataRow})`
    };

    // ARR SUMIFS
    for (let i = 0; i < numDates; i++) {
      const c = layout.arr_start + i;
      const cl = colLetter(c);
      ws.getCell(row, c).value = {
        formula: sumifs(
          rawArrRange,
          [rawDateRange, `${cl}$${dateRefRow}`],
          [rawCustRange, `$${custIdCl}${row}`]
        ).slice(1) // remove leading '='
      };
    }

    // Derived sections
    for (let i = 0; i < layout.num_derived; i++) {
      const priorCl = colLetter(layout.arr_start + i);
      const currCl = colLetter(layout.arr_start + yoyOffset + i);
      const priorRef = `${priorCl}${row}`;
      const currRef = `${currCl}${row}`;

      ws.getCell(row, layout.churn_start + i).value = {
        formula: `IF(AND(${currRef}=0,${priorRef}>0),-${priorRef},0)`
      };
      ws.getCell(row, layout.downsell_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}>0,${currRef}<${priorRef}),${currRef}-${priorRef},0)`
      };
      ws.getCell(row, layout.upsell_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}>0,${currRef}>${priorRef}),${currRef}-${priorRef},0)`
      };
      ws.getCell(row, layout.new_biz_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}=0),${currRef},0)`
      };
    }
  }

  return { sheetName, layout, firstDataRow, lastDataRow };
}


/**
 * Generate an aggregated clean data tab (e.g., quarterly from monthly).
 */
export function generateAggregatedCleanData(
  wb: Workbook, config: EngineConfig,
  sourceSheet: string, sourceLayout: CleanLayout,
  targetGranularity: string, uniqueCustomers: string[],
  targetDates: Date[]
): CleanTabResult {
  const sheetName = `Clean ${targetGranularity.charAt(0).toUpperCase() + targetGranularity.slice(1)} Data`;
  const ws = wb.addWorksheet(sheetName);

  const numDates = targetDates.length;
  const numAttrs = sourceLayout.num_attrs;
  const yoyOffset = getYoyOffset(targetGranularity);
  const numCustomers = uniqueCustomers.length;
  const dataType = config.data_type || 'arr';

  const layout = computeCleanLayout(numAttrs, numDates, yoyOffset);
  const firstDataRow = 7;
  const lastDataRow = firstDataRow + numCustomers - 1;

  const srcArrStartCl = colLetter(sourceLayout.arr_start);
  const srcArrEndCl = colLetter(sourceLayout.arr_end);

  // --- Row 1 ---
  ws.getCell(1, 1).value = 'Fiscal Month End:';
  ws.getCell(1, 2).value = { formula: 'Control!C7' };

  for (const sectionKey of ['arr', 'churn', 'downsell', 'upsell', 'new_biz'] as const) {
    const start = layout[`${sectionKey}_start`];
    const end = sectionKey === 'arr' ? layout.arr_end : layout[`${sectionKey}_end`];
    for (let c = start; c <= end; c++) {
      const cl = colLetter(c);
      ws.getCell(1, c).value = { formula: `SUM(${cl}${firstDataRow}:${cl}${lastDataRow})` };
    }
  }

  // --- Row 2: Quarter ---
  ws.getCell(2, layout.label).value = 'Quarter';

  if (targetGranularity === 'quarterly') {
    ws.getCell(2, layout.arr_start).value = {
      formula: `INDEX('${sourceSheet}'!$${srcArrStartCl}2:$${srcArrEndCl}2,MATCH(TRUE,INDEX('${sourceSheet}'!$${srcArrStartCl}2:$${srcArrEndCl}2<>0,),0))`
    };
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(2, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,1,${prevCl}2+1)` };
    }
  } else if (targetGranularity === 'annual') {
    if (dataType === 'revenue') {
      for (let i = 0; i < numDates; i++) {
        ws.getCell(2, layout.arr_start + i).value = '<>';
      }
    } else {
      ws.getCell(2, layout.arr_start).value = 4;
      for (let i = 1; i < numDates; i++) {
        const prevCl = colLetter(layout.arr_start + i - 1);
        ws.getCell(2, layout.arr_start + i).value = { formula: `${prevCl}2` };
      }
    }
  }

  copyHelperRowsToDerived(ws, layout, 2);

  // --- Row 3: Year ---
  ws.getCell(3, layout.label).value = 'Year';

  if (targetGranularity === 'quarterly') {
    ws.getCell(3, layout.arr_start).value = {
      formula: `INDEX('${sourceSheet}'!$${srcArrStartCl}3:$${srcArrEndCl}3,MATCH(TRUE,INDEX('${sourceSheet}'!$${srcArrStartCl}3:$${srcArrEndCl}3<>0,),0))`
    };
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(3, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,${prevCl}3+1,${prevCl}3)` };
    }
  } else if (targetGranularity === 'annual') {
    ws.getCell(3, layout.arr_start).value = {
      formula: `INDEX('${sourceSheet}'!$${srcArrStartCl}3:$${srcArrEndCl}3,MATCH(TRUE,INDEX('${sourceSheet}'!$${srcArrStartCl}3:$${srcArrEndCl}3<>0,),0))`
    };
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      if (dataType === 'revenue') {
        ws.getCell(3, layout.arr_start + i).value = { formula: `${prevCl}3+1` };
      } else {
        ws.getCell(3, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,${prevCl}3+1,${prevCl}3)` };
      }
    }
  }

  copyHelperRowsToDerived(ws, layout, 3);

  // --- Row 5: Section headers ---
  const metricLabel = dataType === 'arr' ? 'ARR' : 'Revenue';
  ws.getCell(5, layout.attr_start).value = 'Customer Identifying Information';
  ws.getCell(5, layout.arr_start).value = `${targetGranularity.charAt(0).toUpperCase() + targetGranularity.slice(1)} ${metricLabel} by Date`;
  ws.getCell(5, layout.churn_start).value = 'Churn?';
  ws.getCell(5, layout.downsell_start).value = 'Downsell?';
  ws.getCell(5, layout.upsell_start).value = 'Upsell?';
  ws.getCell(5, layout.new_biz_start).value = 'New Business Dollars?';

  // --- Row 6: Column headers ---
  ws.getCell(6, layout.cust_id).value = 'Customer ID';
  const attrNames = Object.keys(config.attributes);
  for (let i = 0; i < attrNames.length; i++) {
    ws.getCell(6, layout.attr_start + i).value = attrNames[i];
  }
  ws.getCell(6, layout.cohort).value = `${targetGranularity.charAt(0).toUpperCase() + targetGranularity.slice(1)} Cohort`;
  ws.getCell(6, layout.rank).value = 'Customer Rank #';

  // Date headers
  for (let i = 0; i < numDates; i++) {
    const c = layout.arr_start + i;
    const cl = colLetter(c);
    if (targetGranularity === 'quarterly') {
      ws.getCell(6, c).value = { formula: `"Q"&${cl}2&"'"&RIGHT(${cl}3,2)` };
    } else if (targetGranularity === 'annual') {
      ws.getCell(6, c).value = { formula: `"FY"&"'"&RIGHT(${cl}3,2)` };
    }
  }

  // Derived section date headers
  for (let i = 0; i < layout.num_derived; i++) {
    const arrRefCl = colLetter(layout.arr_start + yoyOffset + i);
    const churnCol = layout.churn_start + i;
    ws.getCell(6, churnCol).value = { formula: `${arrRefCl}6` };
    ws.getCell(6, layout.downsell_start + i).value = { formula: `${colLetter(churnCol)}6` };
    ws.getCell(6, layout.upsell_start + i).value = { formula: `${colLetter(layout.downsell_start + i)}6` };
    ws.getCell(6, layout.new_biz_start + i).value = { formula: `${colLetter(layout.upsell_start + i)}6` };
  }

  // --- Customer data rows ---
  const arrStartCl = colLetter(layout.arr_start);
  const arrEndCl = colLetter(layout.arr_end);

  for (let idx = 0; idx < uniqueCustomers.length; idx++) {
    const row = firstDataRow + idx;
    // Customer ID: reference source
    ws.getCell(row, layout.cust_id).value = {
      formula: `'${sourceSheet}'!${colLetter(sourceLayout.cust_id)}${row}`
    };

    // Attributes: reference source
    for (let attrIdx = 0; attrIdx < numAttrs; attrIdx++) {
      const srcCol = sourceLayout.attr_start + attrIdx;
      ws.getCell(row, layout.attr_start + attrIdx).value = {
        formula: `'${sourceSheet}'!${colLetter(srcCol)}${row}`
      };
    }

    // Cohort
    ws.getCell(row, layout.cohort).value = {
      formula: `IFERROR(INDEX($${arrStartCl}$6:$${arrEndCl}$6,MATCH(TRUE,INDEX(${arrStartCl}${row}:${arrEndCl}${row}<>0,0),0)),"n.a.")`
    };

    // Rank
    ws.getCell(row, layout.rank).value = {
      formula: `RANK(${arrEndCl}${row},$${arrEndCl}$${firstDataRow}:$${arrEndCl}$${lastDataRow})`
    };

    // ARR: SUMIFS across source row matching Quarter & Year
    for (let i = 0; i < numDates; i++) {
      const c = layout.arr_start + i;
      const cl = colLetter(c);
      ws.getCell(row, c).value = {
        formula: `SUMIFS('${sourceSheet}'!$${srcArrStartCl}${row}:$${srcArrEndCl}${row},'${sourceSheet}'!$${srcArrStartCl}$3:$${srcArrEndCl}$3,${cl}$3,'${sourceSheet}'!$${srcArrStartCl}$2:$${srcArrEndCl}$2,${cl}$2)`
      };
    }

    // Derived sections
    for (let i = 0; i < layout.num_derived; i++) {
      const priorCl = colLetter(layout.arr_start + i);
      const currCl = colLetter(layout.arr_start + yoyOffset + i);
      const priorRef = `${priorCl}${row}`;
      const currRef = `${currCl}${row}`;

      ws.getCell(row, layout.churn_start + i).value = {
        formula: `IF(AND(${currRef}=0,${priorRef}>0),-${priorRef},0)`
      };
      ws.getCell(row, layout.downsell_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}>0,${currRef}<${priorRef}),${currRef}-${priorRef},0)`
      };
      ws.getCell(row, layout.upsell_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}>0,${currRef}>${priorRef}),${currRef}-${priorRef},0)`
      };
      ws.getCell(row, layout.new_biz_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}=0),${currRef},0)`
      };
    }
  }

  return { sheetName, layout, firstDataRow, lastDataRow };
}


/**
 * Generate a clean data tab from already-cleaned table data.
 * Instead of SUMIFS on raw data, directly references cells in the source table.
 */
export function generateCleanDataFromTable(
  wb: Workbook, config: EngineConfig,
  uniqueDates: Date[], uniqueCustomers: string[],
  sourceDateColNums: number[]
): CleanTabResult {
  const granularity = config.time_granularity;
  const sheetName = `Clean ${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Data`;
  const ws = wb.addWorksheet(sheetName);

  const srcSheet = config.raw_data_sheet;
  const srcCustCol = config.customer_id_col;
  const attrs = config.attributes;
  const attrNames = Object.keys(attrs);
  const numAttrs = attrNames.length;
  const numDates = uniqueDates.length;
  const numCustomers = uniqueCustomers.length;
  const yoyOffset = getYoyOffset(granularity);
  const dataType = config.data_type || 'arr';

  const layout = computeCleanLayout(numAttrs, numDates, yoyOffset);
  const firstDataRow = 7;
  const lastDataRow = firstDataRow + numCustomers - 1;

  // Source data starts at row 2 (row 1 is headers)
  const srcFirstDataRow = config.raw_data_first_row;

  // --- Row 1: Fiscal month end + column totals ---
  ws.getCell(1, 1).value = 'Fiscal Month End:';
  ws.getCell(1, 2).value = { formula: 'Control!C7' };

  for (const sectionKey of ['arr', 'churn', 'downsell', 'upsell', 'new_biz'] as const) {
    const start = layout[`${sectionKey}_start`];
    const end = sectionKey === 'arr' ? layout.arr_end : layout[`${sectionKey}_end`];
    for (let c = start; c <= end; c++) {
      const cl = colLetter(c);
      ws.getCell(1, c).value = { formula: `SUM(${cl}${firstDataRow}:${cl}${lastDataRow})` };
    }
  }

  // --- Row 2: Quarter calculation ---
  ws.getCell(2, layout.label).value = 'Quarter';

  if (granularity === 'monthly') {
    for (let i = 0; i < numDates; i++) {
      const c = layout.arr_start + i;
      const cl = colLetter(c);
      if (dataType === 'revenue') {
        ws.getCell(2, c).value = { formula: `INT(MOD(MONTH(${cl}6)-(Control!$C$7+1),12)/3)+1` };
      } else {
        ws.getCell(2, c).value = { formula: `IF(MOD($B$1-MONTH(${cl}6),3)=0,INT(MOD(MONTH(${cl}6)-(Control!$C$7+1),12)/3)+1,0)` };
      }
    }
  } else if (granularity === 'quarterly') {
    const firstDate = uniqueDates[0];
    const month = firstDate.getMonth() + 1;
    const fyMonth = config.fiscal_year_end_month;
    const fq = Math.floor(((month - (fyMonth + 1) + 12) % 12) / 3) + 1;
    ws.getCell(2, layout.arr_start).value = fq;
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(2, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,1,${prevCl}2+1)` };
    }
  } else if (granularity === 'annual') {
    ws.getCell(2, layout.arr_start).value = 4;
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(2, layout.arr_start + i).value = { formula: `${prevCl}2` };
    }
  }

  copyHelperRowsToDerived(ws, layout, 2);

  // --- Row 3: Year calculation ---
  ws.getCell(3, layout.label).value = 'Year';

  if (granularity === 'monthly') {
    for (let i = 0; i < numDates; i++) {
      const c = layout.arr_start + i;
      const cl = colLetter(c);
      ws.getCell(3, c).value = { formula: `IF(${cl}2>0,YEAR(${cl}6),0)` };
    }
  } else if (granularity === 'quarterly') {
    const firstDate = uniqueDates[0];
    const fyMonth = config.fiscal_year_end_month;
    const firstMonth = firstDate.getMonth() + 1;
    const firstYear = firstMonth > fyMonth ? firstDate.getFullYear() + 1 : firstDate.getFullYear();
    ws.getCell(3, layout.arr_start).value = firstYear;
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(3, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,${prevCl}3+1,${prevCl}3)` };
    }
  } else if (granularity === 'annual') {
    const firstDate = uniqueDates[0];
    ws.getCell(3, layout.arr_start).value = firstDate.getFullYear();
    for (let i = 1; i < numDates; i++) {
      const prevCl = colLetter(layout.arr_start + i - 1);
      ws.getCell(3, layout.arr_start + i).value = { formula: `IF(${prevCl}2=4,${prevCl}3+1,${prevCl}3)` };
    }
  }

  copyHelperRowsToDerived(ws, layout, 3);

  // --- Row 4: Month row (date values for aggregated tab matching) ---
  if (granularity === 'quarterly' || granularity === 'annual') {
    ws.getCell(4, layout.label).value = 'Month';
    for (let i = 0; i < numDates; i++) {
      ws.getCell(4, layout.arr_start + i).value = uniqueDates[i];
    }
    copyHelperRowsToDerived(ws, layout, 4);
  }

  // --- Row 5: Section headers ---
  ws.getCell(5, layout.attr_start).value = 'Customer Identifying Information';
  const metricLabel = dataType === 'arr' ? 'ARR' : 'Revenue';
  ws.getCell(5, layout.arr_start).value = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} ${metricLabel} by Date`;
  ws.getCell(5, layout.churn_start).value = 'Churn?';
  ws.getCell(5, layout.downsell_start).value = 'Downsell?';
  ws.getCell(5, layout.upsell_start).value = 'Upsell?';
  ws.getCell(5, layout.new_biz_start).value = 'New Business Dollars?';

  // --- Row 6: Column headers ---
  ws.getCell(6, layout.cust_id).value = 'Customer ID';
  for (let i = 0; i < attrNames.length; i++) {
    ws.getCell(6, layout.attr_start + i).value = attrNames[i];
  }
  ws.getCell(6, layout.cohort).value = `${granularity.charAt(0).toUpperCase() + granularity.slice(1)} Cohort`;
  ws.getCell(6, layout.rank).value = 'Customer Rank #';

  // Date headers in ARR section
  for (let i = 0; i < numDates; i++) {
    const c = layout.arr_start + i;
    const cl = colLetter(c);
    if (granularity === 'quarterly') {
      ws.getCell(6, c).value = { formula: `"Q"&${cl}2&"'"&RIGHT(${cl}3,2)` };
    } else if (granularity === 'annual') {
      ws.getCell(6, c).value = { formula: `"FY"&"'"&RIGHT(${cl}3,2)` };
    } else {
      ws.getCell(6, c).value = uniqueDates[i];
    }
  }

  // Derived section date headers
  for (let i = 0; i < layout.num_derived; i++) {
    const arrRefCol = layout.arr_start + yoyOffset + i;
    const arrRefCl = colLetter(arrRefCol);
    const churnCol = layout.churn_start + i;
    ws.getCell(6, churnCol).value = { formula: `${arrRefCl}6` };
    const downsellCol = layout.downsell_start + i;
    ws.getCell(6, downsellCol).value = { formula: `${colLetter(churnCol)}6` };
    const upsellCol = layout.upsell_start + i;
    ws.getCell(6, upsellCol).value = { formula: `${colLetter(downsellCol)}6` };
    const newBizCol = layout.new_biz_start + i;
    ws.getCell(6, newBizCol).value = { formula: `${colLetter(upsellCol)}6` };
  }

  // --- Customer data rows (7+) ---
  const arrStartCl = colLetter(layout.arr_start);
  const arrEndCl = colLetter(layout.arr_end);
  const srcCustColLetter = srcCustCol;

  for (let idx = 0; idx < uniqueCustomers.length; idx++) {
    const row = firstDataRow + idx;
    const srcRow = srcFirstDataRow + idx;

    // Column B: Customer ID — reference source sheet
    ws.getCell(row, layout.cust_id).value = {
      formula: `'${srcSheet}'!${srcCustColLetter}${srcRow}`
    };

    // Attribute lookups — direct reference to source attribute columns
    for (let attrIdx = 0; attrIdx < attrNames.length; attrIdx++) {
      const rawAttrCol = attrs[attrNames[attrIdx]];
      ws.getCell(row, layout.attr_start + attrIdx).value = {
        formula: `'${srcSheet}'!${rawAttrCol}${srcRow}`
      };
    }

    // Cohort (first non-zero ARR period)
    ws.getCell(row, layout.cohort).value = {
      formula: `IFERROR(INDEX($${arrStartCl}$6:$${arrEndCl}$6,MATCH(TRUE,INDEX(${arrStartCl}${row}:${arrEndCl}${row}<>0,0),0)),"n.a.")`
    };

    // Rank
    ws.getCell(row, layout.rank).value = {
      formula: `RANK(${arrEndCl}${row},$${arrEndCl}$${firstDataRow}:$${arrEndCl}$${lastDataRow})`
    };

    // ARR values — direct cell reference to source table
    for (let i = 0; i < numDates; i++) {
      const srcDateColLetter = colLetter(sourceDateColNums[i]);
      ws.getCell(row, layout.arr_start + i).value = {
        formula: `'${srcSheet}'!${srcDateColLetter}${srcRow}`
      };
    }

    // Derived sections (churn, downsell, upsell, new biz) — same formulas as raw path
    for (let i = 0; i < layout.num_derived; i++) {
      const priorCl = colLetter(layout.arr_start + i);
      const currCl = colLetter(layout.arr_start + yoyOffset + i);
      const priorRef = `${priorCl}${row}`;
      const currRef = `${currCl}${row}`;

      ws.getCell(row, layout.churn_start + i).value = {
        formula: `IF(AND(${currRef}=0,${priorRef}>0),-${priorRef},0)`
      };
      ws.getCell(row, layout.downsell_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}>0,${currRef}<${priorRef}),${currRef}-${priorRef},0)`
      };
      ws.getCell(row, layout.upsell_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}>0,${currRef}>${priorRef}),${currRef}-${priorRef},0)`
      };
      ws.getCell(row, layout.new_biz_start + i).value = {
        formula: `IF(AND(${currRef}>0,${priorRef}=0),${currRef},0)`
      };
    }
  }

  return { sheetName, layout, firstDataRow, lastDataRow };
}


/**
 * Copy Quarter/Year helper rows to derived sections.
 */
function copyHelperRowsToDerived(ws: import('exceljs').Worksheet, layout: CleanLayout, helperRow: number): void {
  const yoy = layout.yoy_offset;
  for (let i = 0; i < layout.num_derived; i++) {
    const arrSrcCol = layout.arr_start + yoy + i;
    const churnCol = layout.churn_start + i;
    ws.getCell(helperRow, churnCol).value = { formula: `${colLetter(arrSrcCol)}${helperRow}` };

    const downsellCol = layout.downsell_start + i;
    ws.getCell(helperRow, downsellCol).value = { formula: `${colLetter(churnCol)}${helperRow}` };

    const upsellCol = layout.upsell_start + i;
    ws.getCell(helperRow, upsellCol).value = { formula: `${colLetter(downsellCol)}${helperRow}` };

    const newBizCol = layout.new_biz_start + i;
    ws.getCell(helperRow, newBizCol).value = { formula: `${colLetter(upsellCol)}${helperRow}` };
  }
}
