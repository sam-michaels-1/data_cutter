import assert from 'node:assert/strict';
import { after, before, test } from 'node:test';
import ExcelJS from 'exceljs';
import { createServer } from 'vite';

let server;
let aggregateToGranularity;
let buildArrMatrix;
let computeDashboard;
let computeSummaryData;
let computeHistogramData;
let computeCustomerDetail;
let generateDataPack;
let buildEngineConfig;
let colLetter;

before(async () => {
  server = await createServer({
    server: { middlewareMode: true, watch: null, ws: false },
    appType: 'custom',
  });
  ({ aggregateToGranularity, buildArrMatrix, computeDashboard } = await server.ssrLoadModule('/src/engine/compute.ts'));
  ({ computeSummaryData } = await server.ssrLoadModule('/src/engine/summary_compute.ts'));
  ({ computeHistogramData } = await server.ssrLoadModule('/src/engine/histograms.ts'));
  ({ computeCustomerDetail } = await server.ssrLoadModule('/src/engine/customer_detail.ts'));
  ({ generateDataPack } = await server.ssrLoadModule('/src/engine/generator.ts'));
  ({ buildEngineConfig } = await server.ssrLoadModule('/src/engine/config_builder.ts'));
  ({ colLetter } = await server.ssrLoadModule('/src/engine/utils.ts'));
});

after(async () => {
  await server?.close();
});

function monthlyRows(startYear = 2023, startMonth = 0, endYear = 2026, endMonth = 7) {
  const rows = [];
  for (let month = startYear * 12 + startMonth; month <= endYear * 12 + endMonth; month++) {
    rows.push({
      date: new Date(Math.floor(month / 12), month % 12 + 1, 0),
      customer_id: 'Customer A',
      arr: 100,
      Segment: 'SMB',
    });
  }
  return rows;
}

function cube(rows, { frequency = 'monthly', fyMonth = 12, dataType = 'revenue' } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Customers');
  ws.addRow(['Customer ID', 'Segment', ...rows.map(row => row.date)]);
  ws.addRow(['Customer A', 'SMB', ...rows.map(row => row.arr)]);
  const config = buildEngineConfig({
    wb,
    sheetName: 'Customers',
    dataType,
    dateCol: 'C',
    customerIdCol: 'A',
    arrCol: 'C',
    attributes: [{ display_name: 'Segment', letter: 'B' }],
    outputGranularities: ['annual', 'quarterly', 'monthly'],
    fiscalYearEndMonth: fyMonth,
    rowCount: 1,
    scaleFactor: 1,
    dataFrequency: frequency,
    inputFormat: 'cleaned',
    dateColumns: rows.map((_, i) => colLetter(i + 3)),
    headerRow: 1,
  });
  return { wb, config };
}

test('annual revenue stops at the last December endpoint, not August FY26', () => {
  const result = aggregateToGranularity(monthlyRows(), 'annual', 12, 'revenue');
  const { periods } = buildArrMatrix(result.records);
  assert.deepEqual(periods, ["FY'23", "FY'24", "FY'25"]);
  assert.deepEqual(result.records.map(row => row.arr), [1200, 1200, 1200]);
  assert.deepEqual(result.periodDateMap, {
    "FY'23": '2023-12-31', "FY'24": '2024-12-31', "FY'25": '2025-12-31',
  });
});

test('FY26 becomes available when the source includes December 2026', () => {
  const result = aggregateToGranularity(monthlyRows(2023, 0, 2026, 11), 'annual', 12, 'revenue');
  assert.deepEqual(buildArrMatrix(result.records).periods, ["FY'23", "FY'24", "FY'25", "FY'26"]);
  assert.equal(result.records.at(-1).arr, 1200);
  assert.equal(result.periodDateMap["FY'26"], '2026-12-31');
});

test('completed annual ARR remains a year-end snapshot rather than a sum', () => {
  const result = aggregateToGranularity(monthlyRows(), 'annual', 12, 'arr');
  assert.deepEqual(buildArrMatrix(result.records).periods, ["FY'23", "FY'24", "FY'25"]);
  assert.deepEqual(result.records.map(row => row.arr), [100, 100, 100]);
});

test('a June fiscal year end excludes the year ending in June 2026 until June is supplied', () => {
  const rows = monthlyRows(2023, 6, 2026, 4);
  const result = aggregateToGranularity(rows, 'annual', 6, 'revenue');
  assert.deepEqual(buildArrMatrix(result.records).periods, ["FY'24", "FY'25"]);
  assert.deepEqual(result.records.map(row => row.arr), [1200, 1200]);
  assert.equal(result.periodDateMap["FY'25"], '2025-06-30');
  const completed = aggregateToGranularity(monthlyRows(2023, 6, 2026, 5), 'annual', 6, 'revenue');
  assert.equal(buildArrMatrix(completed.records).periods.at(-1), "FY'26");
});

test('monthly and quarterly revenue still retain the newest data', () => {
  const rows = monthlyRows();
  const monthly = aggregateToGranularity(rows, 'monthly', 12, 'revenue');
  assert.equal(buildArrMatrix(monthly.records).periods.at(-1), "Aug '26");
  assert.equal(monthly.records.at(-1).arr, 100);
  const quarterly = aggregateToGranularity(rows, 'quarterly', 12, 'revenue');
  assert.equal(buildArrMatrix(quarterly.records).periods.at(-1), "Q3'26");
  assert.equal(quarterly.records.at(-1).arr, 200);
});

test('no completed years produces no annual records, including empty input', () => {
  for (const rows of [[], monthlyRows(2026, 0, 2026, 7)]) {
    assert.deepEqual(aggregateToGranularity(rows, 'annual', 12, 'revenue'), {
      records: [], periodDateMap: {},
    });
  }
});

test('completion uses source coverage even for churned customers and filtered segments', () => {
  const unfiltered = monthlyRows(2025, 0, 2025, 11);
  const churned = unfiltered.slice(0, 6).map(row => ({ ...row, customer_id: 'Churned' }));
  const result = aggregateToGranularity(churned, 'annual', 12, 'revenue', [...unfiltered, ...churned]);
  assert.equal(result.records[0].arr, 600);
  assert.equal(result.periodDateMap["FY'25"], '2025-12-31');
});

test('dashboard, summary, histograms and customer timelines use the same completed annual periods', () => {
  const { wb, config } = cube(monthlyRows());
  const expected = ["FY'23", "FY'24", "FY'25"];
  const dashboard = computeDashboard(wb, config, 'annual');
  assert.deepEqual(dashboard.overview.periods, expected);
  assert.deepEqual(dashboard.overview.arr_over_time, [1200, 1200, 1200]);
  assert.equal(dashboard.overview.latest_period_date, '2025-12-31');
  const summary = computeSummaryData(wb, config, 'annual');
  assert.deepEqual(summary.sections.find(section => section.key === 'dollars').rows.map(row => row.period), expected);
  assert.equal(computeHistogramData(wb, config, 'annual').latestPeriodLabel, "FY'25");
  const detail = computeCustomerDetail(wb, config, 'Customer A', 'annual');
  assert.deepEqual(detail.timeline.map(row => row.period_label), expected);
});

test('attribute filters do not turn a completed source year into an incomplete year', () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Revenue');
  ws.addRow(['Date', 'Customer ID', 'Revenue', 'Segment']);
  for (const row of monthlyRows(2024, 0, 2025, 11)) {
    ws.addRow([row.date, 'Customer A', row.arr, 'Enterprise']);
    if (row.date.getMonth() < 6) ws.addRow([row.date, 'Customer B', row.arr, 'SMB']);
  }
  const config = buildEngineConfig({
    wb, sheetName: 'Revenue', dataType: 'revenue',
    dateCol: 'A', customerIdCol: 'B', arrCol: 'C',
    attributes: [{ display_name: 'Segment', letter: 'D' }],
    outputGranularities: ['annual'], fiscalYearEndMonth: 12,
    rowCount: ws.rowCount - 1, scaleFactor: 1, dataFrequency: 'monthly',
  });
  const filters = { Segment: ['SMB'] };
  const dashboard = computeDashboard(wb, config, 'annual', filters);
  assert.deepEqual(dashboard.overview.periods, ["FY'24", "FY'25"]);
  assert.deepEqual(dashboard.overview.arr_over_time, [600, 600]);
  assert.equal(dashboard.overview.latest_period_date, '2025-12-31');
  const summary = computeSummaryData(wb, config, 'annual', filters);
  assert.deepEqual(summary.sections.find(section => section.key === 'dollars').rows.map(row => row.period), ["FY'24", "FY'25"]);
  assert.equal(computeHistogramData(wb, config, 'annual', filters).latestPeriodLabel, "FY'25");
});

test('monthly and quarterly Excel packs already omit annual FY26 without December', async () => {
  for (const dataType of ['revenue', 'arr']) {
    for (const frequency of ['monthly', 'quarterly']) {
      const rows = monthlyRows().filter(row => frequency === 'monthly' || (row.date.getMonth() + 1) % 3 === 0);
      const { wb, config } = cube(rows, { frequency, dataType });
      const progress = [];
      const output = await generateDataPack(config, wb, msg => progress.push(msg));
      assert.ok(progress.includes('Generating Clean Annual Data (3 years)...'));
      const annual = output.getWorksheet('Clean Annual Data');
      const headers = [];
      annual.getRow(6).eachCell(cell => {
        if (cell.value?.formula?.startsWith('"FY"')) headers.push(cell.value.formula);
      });
      assert.equal(headers.length, 3);
      assert.equal(buildArrMatrix(aggregateToGranularity(rows, 'annual', 12, dataType).records).periods.length, headers.length);
      if (frequency === 'monthly') {
        const monthlyDates = [];
        output.getWorksheet('Clean Monthly Data').getRow(6).eachCell(cell => {
          if (cell.value instanceof Date) monthlyDates.push(cell.value);
        });
        assert.equal(monthlyDates.at(-1).getMonth(), 7);
        assert.equal(monthlyDates.at(-1).getFullYear(), 2026);
      }
    }
  }
});
