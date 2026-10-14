import assert from 'node:assert/strict';
import { after, before, test } from 'node:test';
import ExcelJS from 'exceljs';
import { createServer } from 'vite';

let server;
let detectTableColumns;
let wizardReducer;
let INITIAL_STATE;

before(async () => {
  server = await createServer({
    server: { middlewareMode: true, watch: null, ws: false },
    appType: 'custom',
  });
  ({ detectTableColumns } = await server.ssrLoadModule('/src/engine/detect.ts'));
  ({ wizardReducer, INITIAL_STATE } = await server.ssrLoadModule('/src/hooks/useWizard.ts'));
});

after(async () => {
  await server?.close();
});

function customerCube({ numericIds = true } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Customers');
  ws.addRow([
    'Customer Name',
    null,
    'Segment',
    'Cohort',
    new Date(2023, 0, 31),
    new Date(2023, 1, 28),
  ]);
  for (let i = 0; i < 50; i++) {
    ws.addRow([numericIds ? 1500 + i : `Customer ${i}`, null, i % 2 ? 'SMB' : 'VC Backed', 'Pre-2023', 10 + i, 20 + i]);
  }
  return wb;
}

function detectedState() {
  const detection = detectTableColumns(customerCube(), 'Customers');
  return wizardReducer(INITIAL_STATE, {
    type: 'DETECT_TABLE_SUCCESS',
    columns: detection.columns,
    dateColumns: detection.date_columns,
    customerNameCol: detection.customer_id_col,
    attributes: detection.attribute_cols,
    scaleFactor: detection.auto_scale_factor,
    rowCount: detection.row_count,
    detectedFrequency: detection.detected_frequency,
    headerRow: detection.header_row,
    fiscalLabeled: detection.fiscal_labeled,
  });
}

test('numeric customer IDs outrank low-cardinality text segments', () => {
  const result = detectTableColumns(customerCube(), 'Customers');
  assert.equal(result.customer_id_col, 'A');
  assert.deepEqual(result.attribute_cols, [
    { header: 'Segment', letter: 'C' },
    { header: 'Cohort', letter: 'D' },
  ]);
  assert.equal(result.header_row, 1);
  assert.equal(result.row_count, 50);
  assert.equal(result.detected_frequency, 'monthly');
  assert.deepEqual(result.date_columns, ['E', 'F']);
});

test('text customer IDs continue to be detected', () => {
  const result = detectTableColumns(customerCube({ numericIds: false }), 'Customers');
  assert.equal(result.customer_id_col, 'A');
  assert.deepEqual(result.attribute_cols.map((attr) => attr.letter), ['C', 'D']);
});

test('unique numeric measurements do not displace equally unique customer names', () => {
  const wb = customerCube({ numericIds: false });
  const ws = wb.getWorksheet('Customers');
  ws.getCell('A1').value = 'Measurement';
  ws.getCell('B1').value = 'Customer Name';
  for (let r = 2; r <= 51; r++) {
    ws.getCell(r, 2).value = ws.getCell(r, 1).value;
    ws.getCell(r, 1).value = r * 10;
  }
  assert.equal(detectTableColumns(wb, 'Customers').customer_id_col, 'B');
});

test('correcting a mistaken customer column restores Segment as an identifier', () => {
  let state = detectedState();
  state = wizardReducer(state, { type: 'SET_CUSTOMER_NAME_COL', col: 'C' });
  assert.deepEqual(state.detectedAttributes.map((attr) => attr.letter), ['A', 'D']);
  state = wizardReducer(state, { type: 'SET_CUSTOMER_NAME_COL', col: 'A' });
  assert.deepEqual(state.detectedAttributes, [
    { header: 'Segment', letter: 'C' },
    { header: 'Cohort', letter: 'D' },
  ]);
});

test('remapping removes the new customer column from selections and preserves other renamed identifiers', () => {
  let state = detectedState();
  state = wizardReducer(state, {
    type: 'SET_SELECTED_ATTRIBUTES',
    attrs: [{ display_name: 'Customer Segment', letter: 'C' }, { display_name: 'Original Cohort', letter: 'D' }],
  });
  state = wizardReducer(state, { type: 'SET_CUSTOMER_NAME_COL', col: 'C' });
  assert.deepEqual(state.selectedAttributes, [{ display_name: 'Original Cohort', letter: 'D' }]);
  assert.ok(!state.detectedAttributes.some((attr) => ['B', 'C', 'E', 'F'].includes(attr.letter)));
});

test('remapping preserves detected identifiers populated after the displayed sample rows', () => {
  const wb = customerCube();
  const ws = wb.getWorksheet('Customers');
  for (let r = 2; r <= 6; r++) ws.getCell(r, 4).value = null;
  const result = detectTableColumns(wb, 'Customers');
  assert.deepEqual(result.columns.find((col) => col.letter === 'D').sample_values, Array(5).fill('Pre-2023'));
  let state = { ...detectedState(), columns: result.columns, detectedAttributes: result.attribute_cols };
  state = wizardReducer(state, { type: 'SET_CUSTOMER_NAME_COL', col: 'D' });
  state = wizardReducer(state, { type: 'SET_CUSTOMER_NAME_COL', col: 'A' });
  assert.deepEqual(state.detectedAttributes.map((attr) => attr.letter), ['C', 'D']);
});

test('remapping preserves identifier labels recovered from a separate date-header row', () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Customers');
  ws.addRow(['Customer Name', null, 'Segment', null, new Date(2023, 0, 31), new Date(2023, 1, 28), new Date(2023, 2, 31)]);
  for (let i = 0; i < 50; i++) ws.addRow([1000 + i, null, 'SMB', 'Pre-2023', 500.5, 600.5, 700.5]);
  const result = detectTableColumns(wb, 'Customers');
  assert.equal(result.header_row, 1);
  assert.equal(result.columns.find((col) => col.letter === 'C').header, 'Segment');
  let state = { ...detectedState(), columns: result.columns, dateColumns: result.date_columns, detectedAttributes: result.attribute_cols };
  state = wizardReducer(state, { type: 'SET_CUSTOMER_NAME_COL', col: 'C' });
  state = wizardReducer(state, { type: 'SET_CUSTOMER_NAME_COL', col: 'A' });
  assert.equal(state.detectedAttributes.find((attr) => attr.letter === 'C').header, 'Segment');
});

test('raw-data mapping changes release former roles and exclude current date, customer and revenue roles', () => {
  const columns = ['Customer ID', 'Segment', 'Date', 'Revenue', 'Alternative Date', 'Alternative Revenue'].map((header, i) => ({
    header, letter: String.fromCharCode(65 + i), sample_values: ['value'],
  }));
  const state = wizardReducer({
    ...INITIAL_STATE,
    columns,
    detectedAttributes: [{ header: 'Customer ID', letter: 'A' }],
    selectedAttributes: [{ display_name: 'Customer ID', letter: 'A' }],
    confirmedMapping: { date_col: 'C', customer_id_col: 'B', arr_col: 'D' },
  }, {
    type: 'SET_CONFIRMED_MAPPING',
    mapping: { date_col: 'E', customer_id_col: 'A', arr_col: 'F' },
  });
  assert.deepEqual(state.detectedAttributes.map((attr) => attr.letter), ['B', 'C', 'D']);
  assert.deepEqual(state.selectedAttributes, []);
});

test('switching formats clears identifier selections from the previous detection', () => {
  const state = wizardReducer({
    ...detectedState(),
    selectedAttributes: [{ display_name: 'Segment', letter: 'C' }],
  }, { type: 'SET_INPUT_FORMAT', format: 'raw' });
  assert.deepEqual(state.selectedAttributes, []);
  assert.deepEqual(state.detectedAttributes, []);
});

test('sample workbook keeps its customer ID and attribute mappings', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(new URL('../public/sample-data.xlsx', import.meta.url).pathname);
  const result = detectTableColumns(wb, wb.worksheets[0].name);
  assert.equal(result.columns.find((col) => col.letter === result.customer_id_col).header, 'Customer');
  assert.deepEqual(result.attribute_cols.map((attr) => attr.header), ['Geography', 'Customer Type']);
});
