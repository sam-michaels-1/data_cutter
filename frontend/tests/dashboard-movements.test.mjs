import assert from 'node:assert/strict';
import { after, before, test } from 'node:test';
import ExcelJS from 'exceljs';
import { HyperFormula } from 'hyperformula';
import { createServer } from 'vite';

let server;
let computeDashboard;
let computeDerived;
let readCleanedData;
let readRawData;
let computeSummaryData;
let buildEngineConfig;
let detectTableColumns;
let excelText;
let colLetter;
let generateDataPack;
let generateSummaryTab;
let computeSummarySections;
let computeCleanLayout;

before(async () => {
  server = await createServer({
    server: { middlewareMode: true, watch: null, ws: false },
    appType: 'custom',
  });
  ({ computeDashboard, computeDerived, readCleanedData, readRawData } = await server.ssrLoadModule('/src/engine/compute.ts'));
  ({ computeSummaryData } = await server.ssrLoadModule('/src/engine/summary_compute.ts'));
  ({ buildEngineConfig } = await server.ssrLoadModule('/src/engine/config_builder.ts'));
  ({ detectTableColumns } = await server.ssrLoadModule('/src/engine/detect.ts'));
  ({ excelText, colLetter, computeCleanLayout } = await server.ssrLoadModule('/src/engine/utils.ts'));
  ({ generateDataPack } = await server.ssrLoadModule('/src/engine/generator.ts'));
  ({ generateSummaryTab, computeSummarySections } = await server.ssrLoadModule('/src/engine/summary.ts'));
});

after(async () => {
  await server?.close();
});

function fixture(granularity = 'monthly', numPeriods = granularity === 'monthly' ? 14 : 6, scaleFactor = 1) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Customers');
  const dates = Array.from({ length: numPeriods }, (_, i) => new Date(2025, (i + 1) * (granularity === 'monthly' ? 1 : 3), 0));
  ws.addRow(['Customer ID', 'Segment', ...dates]);
  for (const [name, prior, current, segment] of [
    ['Churn', 100, 0, 'SMB'],
    ['Downsell', 100, 50, 'SMB'],
    ['Upsell', 100, 125, 'Enterprise'],
    ['New', 0, 25, 'Enterprise'],
  ]) {
    ws.addRow([name, segment, ...dates.map((_, i) => i === numPeriods - 1 ? current : prior)]);
  }
  const config = buildEngineConfig({
    wb, sheetName: 'Customers', dataType: 'arr',
    customerIdCol: 'A', dateCol: 'C', arrCol: 'C',
    attributes: [{ display_name: 'Segment', letter: 'B' }],
    outputGranularities: [granularity, 'annual'],
    fiscalYearEndMonth: 12, rowCount: 4, scaleFactor,
    dataFrequency: granularity, inputFormat: 'cleaned',
    dateColumns: dates.map((_, i) => colLetter(i + 3)), headerRow: 1,
  });
  return { wb, ws, config };
}

test('Excel text extraction handles errors, cached formulas, rich text, links and dates', () => {
  assert.equal(excelText({ error: '#N/A' }), '');
  assert.equal(excelText({ formula: 'A1', result: { error: '#VALUE!' } }), '');
  assert.equal(excelText({ formula: 'A1', result: 'SMB' }), 'SMB');
  assert.equal(excelText({ sharedFormula: 'A1', result: 'Enterprise' }), 'Enterprise');
  assert.equal(excelText({ formula: 'A1' }), '');
  assert.equal(excelText({ richText: [{ text: 'VC ' }, { text: 'Backed' }] }), 'VC Backed');
  assert.equal(excelText({ text: 'SMB', hyperlink: 'https://example.com' }), 'SMB');
  assert.equal(excelText(new Date(Date.UTC(2023, 0, 31))), '2023-01-31');
  assert.equal(excelText(0), '0');
  assert.equal(excelText(false), 'false');
});

test('cleaned-table error segments become Unknown without losing customer revenue', async () => {
  const { wb, ws, config } = fixture();
  ws.getCell('B2').value = { error: '#N/A' };
  ws.getCell('B3').value = { formula: 'NA()', result: { error: '#N/A' } };
  ws.getCell('B4').value = { richText: [{ text: 'VC ' }, { text: 'Backed' }] };
  ws.getCell('B5').value = { formula: '"Enterprise"', result: 'Enterprise' };
  // Exercise the same serialization as an uploaded .xlsx, not only in-memory cells.
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(await wb.xlsx.writeBuffer());
  const rows = readCleanedData(loaded, config);
  assert.equal(rows.length, 56);
  assert.deepEqual([...new Set(rows.map(row => row.Segment))].sort(), ['Enterprise', 'Unknown', 'VC Backed']);
  const dashboard = computeDashboard(loaded, config, 'monthly');
  assert.equal(dashboard.overview.stats.total_arr, 200);
  assert.deepEqual(dashboard.attribute_options.find(option => option.name === 'Segment').values, ['Enterprise', 'Unknown', 'VC Backed']);
  const summary = computeSummaryData(loaded, config, 'monthly', undefined, 'Segment');
  assert.deepEqual(summary.columns, ['All', 'Enterprise', 'Unknown', 'VC Backed']);
  const unknown = computeDashboard(loaded, config, 'monthly', { Segment: ['Unknown'] });
  assert.equal(unknown.overview.stats.total_arr, 50);
  const detected = detectTableColumns(loaded, 'Customers');
  assert.ok(!detected.columns.flatMap(col => col.sample_values).includes('[object Object]'));
  const rebuilt = buildEngineConfig({
    wb: loaded, sheetName: 'Customers', dataType: 'arr', dateCol: 'C', arrCol: 'C', customerIdCol: 'A',
    attributes: [{ display_name: 'Segment', letter: 'B' }],
    outputGranularities: ['monthly'], fiscalYearEndMonth: 12, rowCount: 4,
    scaleFactor: 1, dataFrequency: 'monthly', inputFormat: 'cleaned', dateColumns: config.date_columns,
  });
  assert.deepEqual(rebuilt.filter_breakouts.map(block => block.title), ['Enterprise', 'Unknown', 'VC Backed']);
});

test('raw-format attributes use the same object-safe normalization', () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Revenue');
  ws.addRow(['Date', 'Customer ID', 'Revenue', 'Segment']);
  ws.addRow([new Date(2025, 0, 31), 'A', 100, { error: '#N/A' }]);
  ws.addRow([new Date(2025, 0, 31), 'B', 200, { richText: [{ text: 'SMB' }] }]);
  const rows = readRawData(wb, {
    raw_data_sheet: 'Revenue', date_col: 'A', customer_id_col: 'B', arr_col: 'C', attributes: { Segment: 'D' },
  });
  assert.deepEqual(rows.map(row => row.Segment), ['Unknown', 'SMB']);
  assert.equal(rows.reduce((total, row) => total + row.arr, 0), 300);
});

for (const [granularity, factor] of [['monthly', 12], ['quarterly', 4]]) {
  test(`${granularity} annualized waterfall scales latest movements by ${factor}, not BoP`, () => {
    const { wb, config } = fixture(granularity);
    const result = computeDashboard(wb, config, granularity).overview;
    const annualized = result.annualized_waterfall;
    assert.equal(annualized.annualization_factor, factor);
    assert.equal(annualized.bop, 300);
    assert.equal(annualized.churn, -100 * factor);
    assert.equal(annualized.downsell, -50 * factor);
    assert.equal(annualized.upsell, 25 * factor);
    assert.equal(annualized.new_logo, 25 * factor);
    assert.equal(annualized.eop, 300 - 100 * factor);
    assert.equal(result.waterfall.eop, 200);
    assert.equal(result.stats.total_arr, 200);
    assert.equal(result.stats.annualized_punitive_retention_pct, 1 - factor / 2);
    assert.equal((annualized.bop + annualized.churn + annualized.downsell) / annualized.bop, result.stats.annualized_punitive_retention_pct);
    assert.equal(annualized.bop + annualized.churn + annualized.downsell + annualized.upsell + annualized.new_logo, annualized.eop);
    assert.notEqual(annualized.prior_period_label, annualized.period_label);
  });

  test(`${granularity} annualized movements respect segment and cohort filters and scaling`, () => {
    const { wb, config } = fixture(granularity, undefined, 1000);
    for (const filters of [{ Segment: ['SMB'] }, { Cohort: [granularity === 'monthly' ? "Jan '25" : "Q1'25"] }]) {
      const result = computeDashboard(wb, config, granularity, filters).overview;
      const bridge = result.annualized_waterfall;
      assert.ok(bridge);
      assert.equal(bridge.new_logo, 0);
      assert.equal(result.stats.total_arr, filters.Segment ? 0.05 : 0.175);
      assert.equal(bridge.bop, filters.Segment ? 0.2 : 0.3);
      assert.equal(bridge.churn, -100 * factor / 1000);
    }
  });
}

test('annualized mode works with two periods even without a full YoY lookback', () => {
  const { wb, config } = fixture('monthly', 2);
  const result = computeDashboard(wb, config, 'monthly').overview;
  assert.equal(result.waterfall, null);
  assert.equal(result.annualized_waterfall.eop, -900);
});

test('one period and annual granularity do not produce annualized waterfalls', () => {
  const { wb, config } = fixture('monthly', 1);
  assert.equal(computeDashboard(wb, config, 'monthly').overview.annualized_waterfall, null);
  const longer = fixture();
  assert.equal(computeDashboard(longer.wb, longer.config, 'annual').overview.annualized_waterfall, null);
});

test('zero BoP still produces an annualized new-customer bridge without inventing retention', () => {
  const { wb, ws, config } = fixture('monthly', 2);
  for (let r = 2; r <= 5; r++) ws.getCell(r, 3).value = 0;
  const result = computeDashboard(wb, config, 'monthly').overview;
  assert.equal(result.annualized_waterfall.bop, 0);
  assert.equal(result.annualized_waterfall.eop, 2400);
  assert.equal(result.stats.annualized_punitive_retention_pct, null);
});

test('movement buckets reconcile across positive, zero and negative revenue balances', () => {
  for (const prior of [-20, -10, 0, 10, 20]) {
    for (const current of [-20, -10, 0, 10, 20]) {
      const pivot = new Map([['Customer', new Map([['Prior', prior], ['Current', current]])]]);
      const derived = computeDerived(pivot, ['Prior', 'Current'], 1);
      const amounts = Object.values(derived).map(bucket => bucket.get('Customer').get('Current'));
      assert.equal(prior + amounts.reduce((sum, amount) => sum + amount, 0), current);
      assert.ok(derived.downsell.get('Customer').get('Current') <= 0);
      assert.ok(derived.upsell.get('Customer').get('Current') >= 0);
    }
  }
});

test('negative credits and their recovery reconcile both actual and annualized waterfalls', () => {
  const { wb, ws, config } = fixture();
  ws.getCell('A2').value = 'Credit Recovery';
  for (let c = 3; c < ws.columnCount; c++) ws.getCell(2, c).value = -100;
  const result = computeDashboard(wb, config, 'monthly').overview;
  assert.equal(result.stats.total_arr, 200);
  assert.equal(result.waterfall.bop, 100);
  assert.equal(result.waterfall.upsell, 125);
  const bridge = result.annualized_waterfall;
  assert.equal(bridge.bop, 100);
  assert.equal(bridge.upsell, 1500);
  assert.equal(bridge.eop, 1300);
  assert.equal(result.arr_over_time.at(-1), result.waterfall.eop);
});

test('Excel clean tabs use the same signed-movement conditions for raw and cleaned inputs', async () => {
  for (const inputFormat of ['raw', 'cleaned']) {
    let { wb, ws, config } = fixture();
    if (inputFormat === 'raw') {
      const rawWb = new ExcelJS.Workbook();
      const raw = rawWb.addWorksheet('Revenue');
      raw.addRow(['Date', 'Customer ID', 'Revenue', 'Segment']);
      for (let r = 2; r <= ws.rowCount; r++) {
        for (let c = 3; c <= ws.columnCount; c++) {
          raw.addRow([ws.getCell(1, c).value, ws.getCell(r, 1).value, ws.getCell(r, c).value, ws.getCell(r, 2).value]);
        }
      }
      wb = rawWb;
      config = {
        ...config, raw_data_sheet: 'Revenue', raw_data_last_row: raw.rowCount,
        customer_id_col: 'B', date_col: 'A', arr_col: 'C', attributes: { Segment: 'D' },
        input_format: 'raw', date_columns: undefined,
      };
    }
    const output = await generateDataPack(config, wb);
    for (const sheet of output.worksheets.filter(sheet => sheet.name.startsWith('Clean '))) {
      let downsellColumn;
      let upsellColumn;
      sheet.getRow(5).eachCell((cell, column) => {
        if (cell.value === 'Downsell?') downsellColumn = column;
        if (cell.value === 'Upsell?') upsellColumn = column;
      });
      const down = sheet.getCell(7, downsellColumn).value;
      const up = sheet.getCell(7, upsellColumn).value;
      if (down?.formula) assert.match(down.formula, /^IF\(AND\([A-Z]+7<>0,[A-Z]+7<[A-Z]+7\),/);
      if (up?.formula) assert.match(up.formula, /^IF\(AND\([A-Z]+7<>0,[A-Z]+7>[A-Z]+7\),/);
    }
  }
});

test('wide-table exports preserve source rows after invalid IDs and sum repeated customer rows', async () => {
  const { wb, ws, config } = fixture();
  const amounts = Array(14).fill(300);
  ws.insertRows(3, [
    [null, 'SMB', ...amounts],
    [{ error: '#N/A' }, 'SMB', ...amounts],
    [{ formula: 'NA()', result: { error: '#N/A' } }, 'SMB', ...amounts],
  ]);
  ws.addRow(['Downsell', 'SMB', ...Array(14).fill(10)]);
  config.raw_data_last_row = ws.rowCount;
  const output = await generateDataPack(config, wb);
  const clean = output.getWorksheet('Clean Monthly Data');
  const layout = computeCleanLayout(1, 14, 12);
  assert.equal(clean.getCell(7, layout.cust_id).value.formula, "'Customers'!A2");
  assert.equal(clean.getCell(8, layout.cust_id).value.formula, "'Customers'!A6");
  assert.equal(clean.getCell(9, layout.cust_id).value.formula, "'Customers'!A7");
  assert.equal(clean.getCell(10, layout.cust_id).value.formula, "'Customers'!A8");
  assert.equal(clean.getCell(8, layout.arr_start).value.formula, "'Customers'!C6+'Customers'!C9");
  assert.equal(clean.getCell(8, layout.arr_end).value.formula, "'Customers'!P6+'Customers'!P9");
  assert.equal(clean.getCell(9, layout.arr_end).value.formula, "'Customers'!P7");
  assert.match(clean.getCell(8, layout.attr_start).value.formula, /B6/);
  assert.equal(computeDashboard(wb, config, 'monthly').overview.stats.total_arr, 210);
});

for (const granularity of ['monthly', 'quarterly']) {
  test(`${granularity} Excel annualized summaries recalculate to web retention with signed credits`, () => {
    const { wb, ws, config } = fixture(granularity);
    const balances = [
      ['Refund', 10, -5, 'SMB'],
      ['Recovery', -5, 0, 'SMB'],
      ['New Credit', 0, -2, 'SMB'],
      ['Back to Positive', -5, 10, 'Enterprise'],
      ['Downsell', 30, 25, 'Enterprise'],
    ];
    for (let i = 0; i < balances.length; i++) {
      const [name, prior, current, segment] = balances[i];
      ws.getRow(i + 2).values = [name, segment, ...config.date_columns.map((_, p) => p === config.date_columns.length - 1 ? current : prior)];
    }
    config.raw_data_last_row = ws.rowCount;
    const numDates = config.date_columns.length;
    const factor = granularity === 'monthly' ? 12 : 4;
    const layout = computeCleanLayout(1, numDates, factor);
    const cleanName = `Clean ${granularity === 'monthly' ? 'Monthly' : 'Quarterly'} Data`;
    const clean = wb.addWorksheet(cleanName);
    clean.getCell(6, layout.attr_start).value = 'Segment';
    clean.getCell(6, layout.cohort).value = 'Cohort';
    balances.forEach(([name, prior, current, segment], i) => {
      clean.getCell(7 + i, layout.cust_id).value = name;
      clean.getCell(7 + i, layout.attr_start).value = segment;
      for (let p = 0; p < numDates; p++) clean.getCell(7 + i, layout.arr_start + p).value = p === numDates - 1 ? current : prior;
    });
    const control = wb.addWorksheet('Control');
    control.getCell('C4').value = 1;
    const summaryName = generateSummaryTab(wb, config, cleanName, layout, 7, 11, granularity, 'Segment', ['SMB', 'Enterprise']);
    // HyperFormula cannot evaluate INDEX(range,0,MATCH(...)) as a criteria column.
    // Resolve that selector only; preserve the generated retention formulas.
    const selector = `INDEX('${cleanName}'!$C$7:$D$11,0,MATCH($B$5,'${cleanName}'!$C$6:$D$6,0))`;
    const criteria = `'${cleanName}'!$C$7:$C$11`;
    const sheets = {};
    for (const sheet of [clean, control, wb.getWorksheet(summaryName)]) {
      sheets[sheet.name] = Array.from({ length: sheet.rowCount }, (_, r) =>
        Array.from({ length: sheet.columnCount }, (_, c) => {
          const value = sheet.getCell(r + 1, c + 1).value;
          return value?.formula ? `=${value.formula.replaceAll(selector, criteria)}` : value ?? null;
        }));
    }
    const hf = HyperFormula.buildFromSheets(sheets, { licenseKey: 'gpl-v3', useArrayArithmetic: true });
    try {
      const sections = computeSummarySections(numDates, layout.num_derived);
      const web = computeSummaryData(wb, config, granularity, undefined, 'Segment');
      for (const key of ['ann_gross', 'ann_net']) {
        const excelSection = sections.find(section => section.key === key);
        const webRow = web.sections.find(section => section.key === key).rows.at(-1);
        for (const [segment, column] of [['All', 4], ['SMB', 5], ['Enterprise', 6]]) {
          const actual = hf.getCellValue({ sheet: hf.getSheetId(summaryName), row: excelSection.startRow + excelSection.numRows - 2, col: column - 1 });
          const expected = webRow.values[web.columns.indexOf(segment)];
          assert.equal(typeof actual, 'number', `${key} / ${segment}: ${JSON.stringify(actual)}`);
          assert.ok(Math.abs(actual - expected) < 1e-9, `${key} / ${segment}: Excel ${actual} vs web ${expected}`);
        }
      }
    } finally { hf.destroy(); }
  });
}
