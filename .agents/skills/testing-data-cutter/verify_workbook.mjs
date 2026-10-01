import ExcelJS from 'exceljs';
import { HyperFormula } from 'hyperformula';
import fs from 'fs';

const FILE = '/home/ubuntu/Downloads/data-pack-output.xlsx';
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(FILE);

const results = [];
const check = (name, pass, detail = '') => {
  results.push({ name, pass, detail });
  console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

// ---------- structural checks (unchanged from before) ----------
const sheetNames = wb.worksheets.map(w => w.name);
check("Sheet 'Clean Quarterly Data (QoQ)' exists", sheetNames.includes('Clean Quarterly Data (QoQ)'));
check("Sheet 'Quarterly Retention (QoQ)' exists", sheetNames.includes('Quarterly Retention (QoQ)'));

const strVal = (ws, r, c) => {
  const v = ws.getCell(r, c).value;
  if (v == null) return null;
  if (typeof v === 'string') return v;
  if (typeof v === 'object' && v.richText) return v.richText.map(t => t.text).join('');
  return null;
};

function findLabelCol(ws, needle = '% Net Retention') {
  for (let c = 1; c <= ws.columnCount; c++)
    for (let r = 1; r <= Math.min(ws.rowCount, 300); r++)
      if (strVal(ws, r, c) === needle) return c;
  return -1;
}
function titleRows(ws) {   // block titles only: contain granularity word, end with 'Retention Analysis'
  const rows = new Set();
  for (let r = 1; r <= ws.rowCount; r++)
    for (let c = 1; c <= ws.columnCount; c++) {
      const v = strVal(ws, r, c);
      if (v && /Retention Analysis$/.test(v) && /Quarterly|Annual|Monthly/.test(v)) rows.add(r);
    }
  return [...rows].sort((a, b) => a - b);
}

const qoqWs = wb.getWorksheet('Quarterly Retention (QoQ)');
const lc = findLabelCol(qoqWs, '% Annualized Net Retention');
check('QoQ tab: Section 1 label column located', lc > 0, `col=${lc}`);

// QoQ tab: only the annualized retention rows (plain % Lost-Only/Punitive/Net removed)
const qoqSeq = [
  '% Annualized Lost-Only Retention', '% Annualized Punitive Retention',
  '% Annualized Net Retention', '% New Logo % of BoP', '% New Logo Growth',
];
let allOk = true, blocksChecked = 0;
for (let r = 1; r <= qoqWs.rowCount; r++) {
  if (strVal(qoqWs, r, lc) === qoqSeq[0]) {
    blocksChecked++;
    qoqSeq.forEach((exp, i) => {
      if (strVal(qoqWs, r + i, lc) !== exp) { allOk = false; }
    });
  }
}
check('QoQ tab: annualized-only retention sequence in every block',
  allOk && blocksChecked > 0, `${blocksChecked} blocks checked`);
let plainCount = 0;
qoqWs.eachRow(row => row.eachCell(cell => {
  if (['% Lost-Only Retention', '% Punitive Retention', '% Net Retention'].includes(cell.value)) plainCount++;
}));
check('QoQ tab: no non-annualized retention rows', plainCount === 0, `${plainCount} found`);
{
  const tr = titleRows(qoqWs);
  const strides = tr.slice(1).map((r, i) => r - tr[i]);
  check('QoQ tab: block pitch is 19 rows', strides.every(s => s === 19), `blocks=${tr.length} strides=${[...new Set(strides)]}`);
}
for (const name of ['Quarterly Retention', 'Annual Retention']) {
  const ws = wb.getWorksheet(name);
  const lc2 = findLabelCol(ws, '% Net Retention');
  const yoySeq = ['% Lost-Only Retention', '% Punitive Retention', '% Net Retention', '% New Logo % of BoP', '% New Logo Growth'];
  let annCount = 0, seqOk = true, seqBlocks = 0;
  ws.eachRow(row => row.eachCell(cell => {
    const v = cell.value;
    const s = typeof v === 'string' ? v : (v && v.richText ? v.richText.map(t => t.text).join('') : '');
    if (/Annualized/.test(s)) annCount++;
  }));
  for (let r = 1; r <= ws.rowCount; r++) {
    if (strVal(ws, r, lc2) === yoySeq[0]) {
      seqBlocks++;
      yoySeq.forEach((exp, i) => { if (strVal(ws, r + i, lc2) !== exp) seqOk = false; });
    }
  }
  check(`YoY tab '${name}': no 'Annualized' labels`, annCount === 0, `${annCount} found`);
  check(`YoY tab '${name}': standard retention sequence in every block`, seqOk && seqBlocks > 0, `${seqBlocks} blocks`);
  const tr = titleRows(ws);
  const strides = tr.slice(1).map((r, i) => r - tr[i]);
  check(`YoY tab '${name}': block pitch is 19 rows`, strides.every(s => s === 19), `blocks=${tr.length} strides=${[...new Set(strides)]}`);
}

// ---------- faithful HyperFormula evaluation ----------
// Precompute literals for cells HF can't evaluate:
//  - formulas containing MATCH(TRUE, ...) : cohort = first header with nonzero row value
//  - formulas containing RANK(            : descending rank of ref cell in ref range
//  - Clean Annual Data year row (H3:K3)  : INDEX/MATCH array -> literal years
const excelSerial = d => (d.getTime() - Date.UTC(1899, 11, 30)) / 86400000;

const colNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
const colName = n => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

// ---------- BoP = prior EoP direct links (purple font) ----------
// BoP[i] links to the EoP column yoyOffset back; the first yoyOffset columns
// keep SUMIFS/COUNTIFS. applyFormulaColoring turns direct links purple (7030A0).
const PURPLE = /7030A0$/i;
const fontArgb = (ws, r, c) => (ws.getCell(r, c).font && ws.getCell(r, c).font.color && ws.getCell(r, c).font.color.argb) || '';
const retSheets = { 'Quarterly Retention': 4, 'Annual Retention': 1, 'Quarterly Retention (QoQ)': 1 };
for (const [name, off] of Object.entries(retSheets)) {
  const ws = wb.getWorksheet(name);
  if (!ws) { check(`${name}: sheet present`, false); continue; }
  const s1lc = findLabelCol(ws, name.includes('QoQ') ? '% Annualized Net Retention' : '% Net Retention');
  const s2lc = findLabelCol(ws, '(-) Churned Customers');
  const tr = titleRows(ws);
  for (const [band, bandLc, eopOff] of [['s1 ARR', s1lc, 6], ['s2 Customers', s2lc, 4]]) {
    let links = 0, purple = 0, first = 0, bad = [];
    for (const t of tr) {
      const rBop = t + 3, rEop = t + 3 + eopOff;
      for (let c = bandLc + 1; ; c++) {
        const v = ws.getCell(rBop, c).value;
        if (v == null) break;
        const i = c - (bandLc + 1);
        const f = v && typeof v === 'object' ? v.formula : null;
        if (i >= off) {
          if (f === `${colName(c - off)}${rEop}`) links++;
          else bad.push(`${colName(c)}${rBop}:${f}`);
          if (PURPLE.test(fontArgb(ws, rBop, c))) purple++;
        } else if (f && /^(SUMIFS|COUNTIFS)/.test(f)) first++;
      }
    }
    check(`${name} ${band}: BoP links to EoP ${off} col(s) back, purple`, links > 0 && bad.length === 0 && purple === links,
      `${links} links (${purple} purple), ${first} head cols${bad.length ? ' BAD: ' + bad.slice(0, 3).join(' | ') : ''}`);
  }
}
{
  // % Annualized Net Retention: bold label + indent + bold data (QoQ tab)
  let r = -1;
  for (let rr = 1; rr <= qoqWs.rowCount; rr++) if (strVal(qoqWs, rr, lc) === '% Annualized Net Retention') { r = rr; break; }
  const labelCell = qoqWs.getCell(r, lc), dataCell = qoqWs.getCell(r, lc + 1);
  check('QoQ tab: % Annualized Net Retention bold + indented',
    !!labelCell.font?.bold && !!dataCell.font?.bold && (labelCell.alignment?.indent || 0) > 0,
    `bold(label)=${labelCell.font?.bold} bold(data)=${dataCell.font?.bold} indent=${labelCell.alignment?.indent}`);
}

// s1 '% Growth' row (EoP ARR % Growth): italic label+data, indented, whole-percent —
// same treatment as the s2 'EoP Customers % Growth' row, on every retention tab.
for (const name of Object.keys(retSheets)) {
  const ws = wb.getWorksheet(name);
  if (!ws) continue;
  const lc2 = findLabelCol(ws, name.includes('QoQ') ? '% Annualized Net Retention' : '% Net Retention');
  const tr = titleRows(ws);
  let checked = 0, bad = [];
  for (const t of tr) {
    const r = t + 10; // rGrowth
    const lbl = ws.getCell(r, lc2);
    if (strVal(ws, r, lc2) !== '% Growth') { bad.push(`${name} block@${t} label=${strVal(ws, r, lc2)}`); continue; }
    checked++;
    if (!lbl.font?.italic) bad.push(`${name} r${r} label not italic`);
    if ((lbl.alignment?.indent || 0) < 1) bad.push(`${name} r${r} label not indented`);
    for (let c = lc2 + 1; c <= lc2 + 12; c++) {
      const dc = ws.getCell(r, c);
      if (dc.value == null) break;
      if (!dc.font?.italic) { bad.push(`${name} ${colName(c)}${r} data not italic`); break; }
      if (dc.numFmt && dc.numFmt.includes('.0')) { bad.push(`${name} ${colName(c)}${r} decimal pct ${dc.numFmt}`); break; }
    }
  }
  check(`${name}: s1 % Growth italic + indented + whole-percent`, checked > 0 && bad.length === 0,
    `${checked} blocks${bad.length ? ' BAD: ' + bad.slice(0, 3).join(' | ') : ''}`);
}

// Rewrite INDEX('<sheet>'!$A$r1:$B$r2,0,MATCH($X$n,'<sheet>'!$C$6:$D$6,0))
// -> '<sheet>'!$<picked col>$r1:$<picked col>$r2  (HF can't INDEX a whole column)
function rewriteColumnIndex(f, curSheetName) {
  const re = /INDEX\('([^']+)'!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+),0,MATCH\((\$?[A-Z]+\$?\d+),'[^']*'!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+),0\)\)/g;
  return f.replace(re, (m, sheet, c1, r1, c2, r2, lookupRef, h1, hRow, h2) => {
    const srcWs = wb.getWorksheet(sheet);
    const curWs = wb.getWorksheet(curSheetName);
    if (!srcWs || !curWs) return m;
    const lv = curWs.getCell(lookupRef.replace(/\$/g, '')).value;  // lookup cell on current sheet
    const lookup = typeof lv === 'object' && lv ? (lv.result ?? lv.formula ?? null) : lv;
    const hs = colNum(h1), he = colNum(h2);
    let k = -1;
    for (let c = hs; c <= he; c++) {
      const hv = srcWs.getCell(+hRow, c).value;
      const hs2 = typeof hv === 'object' && hv ? (hv.richText ? hv.richText.map(t => t.text).join('') : hv.result) : hv;
      if (hs2 === lookup) { k = c - hs + 1; break; }
    }
    if (k < 1) return m;
    const picked = colName(colNum(c1) + k - 1);
    return `'${sheet}'!$${picked}$${r1}:$${picked}$${r2}`;
  });
}

function cellOut(v, curSheetName) {
  if (v == null) return null;
  if (typeof v === 'object' && v.formula != null) {
    let f = v.formula;
    if (!f.startsWith('=') && /^# /.test(f)) return f;           // '# Customers' misread
    f = f.replace(/_xlfn\._xlws\./g, '').replace(/_xlfn\./g, '')
         .replace(/\bTRUE\b(?!\()/g, 'TRUE()').replace(/\bFALSE\b(?!\()/g, 'FALSE()');
    f = rewriteColumnIndex(f, curSheetName);
    return '=' + f;
  }
  if (v instanceof Date) return excelSerial(v);
  if (typeof v === 'object' && v.richText) return v.richText.map(t => t.text).join('');
  if (typeof v === 'object' && v.result !== undefined) return v.result;
  return v;
}

function buildSheets(literals) {
  const sheets = {};
  for (const ws of wb.worksheets) {
    const rows = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const vals = [];
      for (let c = 1; c <= ws.columnCount; c++) {
        const key = `${ws.name}!${r},${c}`;
        vals.push(key in literals ? literals[key] : cellOut(row.getCell(c).value, ws.name));
      }
      rows.push(vals);
    }
    sheets[ws.name] = rows;
  }
  return sheets;
}

// Collect cells needing literals
const needLiteral = [];   // {sheet,row,col,kind,formula}
for (const ws of wb.worksheets) {
  ws.eachRow(row => row.eachCell(cell => {
    const v = cell.value;
    if (v && typeof v === 'object' && typeof v.formula === 'string') {
      if (/MATCH\(TRUE,/.test(v.formula)) needLiteral.push({ sheet: ws.name, row: cell.row, col: cell.col, kind: 'cohort', formula: v.formula });
      else if (/\bRANK\s*\(/.test(v.formula)) needLiteral.push({ sheet: ws.name, row: cell.row, col: cell.col, kind: 'rank', formula: v.formula });
    }
  }));
}
console.log(`cells needing literals: ${needLiteral.length} (cohort=${needLiteral.filter(x=>x.kind==='cohort').length}, rank=${needLiteral.filter(x=>x.kind==='rank').length})`);

// Pass 0: Clean Annual Data year row — compute from Clean Quarterly Data literals.
// CQD rows 2 (quarter) & 3 (year) are literal numbers.
const cqd = wb.getWorksheet('Clean Quarterly Data');
const cad = wb.getWorksheet('Clean Annual Data');
const literals = {};
if (cad) {
  // find CAD year-row cells with MATCH(TRUE,...)
  for (let c = 1; c <= cad.columnCount; c++) {
    const v = cad.getCell(3, c).value;
    if (v && typeof v === 'object' && v.formula && /MATCH\(TRUE/.test(v.formula)) {
      // INDEX('Clean Quarterly Data'!$H3:$W3, MATCH(TRUE, INDEX(quarterRow=4,0),0))
      // => year of first quarterly column whose quarter==4. Other year cells are IF(prevQ=4, prevY+1, prevY).
      // Compute directly: read CQD row2/row3 literal arrays.
      const qrow = [], yrow = [];
      for (let qc = 1; qc <= cqd.columnCount; qc++) { qrow.push(cqd.getCell(2, qc).value); yrow.push(cqd.getCell(3, qc).value); }
      let firstYr = null;
      for (let i = 0; i < qrow.length; i++) if (qrow[i] === 4) { firstYr = yrow[i]; break; }
      // count how many consecutive year cells this column index represents
      // H3=firstYr, I3=firstYr+1, ... (all quarter cells =4)
      const base = c - 8;  // H is col 8
      literals[`${'Clean Annual Data'}!3,${c}`] = firstYr + base;
    }
  }
}

// Pass 1: build with cohort/rank cells blank, evaluate ARR columns
const sheets1 = buildSheets(literals);
const hf1 = HyperFormula.buildFromSheets(sheets1, { licenseKey: 'gpl-v3' });
const val1 = (sheet, r, c) => {
  const v = hf1.getCellValue({ sheet: hf1.getSheetId(sheet), row: r - 1, col: c - 1 });
  if (v && typeof v === 'object' && v.value !== undefined) return v.value;
  return v;
};

// Compute cohort literals: pattern INDEX($X$6:$Y$6,MATCH(TRUE(),INDEX(r...<>0,0),0))
// => header (row6) of first nonzero cell in the row's ARR range
for (const it of needLiteral.filter(x => x.kind === 'cohort')) {
  const m = it.formula.match(/INDEX\(\$?([A-Z]+)\$?6:\$?([A-Z]+)\$?6,MATCH\(TRUE,INDEX\(([A-Z]+)\d+:\$?([A-Z]+)\d+<>0,0\)/);
  if (!m) { console.log('unparsed cohort formula', it.sheet, it.formula); continue; }
  const colNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const h1 = colNum(m[1]), h2 = colNum(m[2]), d1 = colNum(m[3]), d2 = colNum(m[4]);
  let lit = 'n.a.';
  for (let c = d1; c <= d2; c++) {
    const v = val1(it.sheet, it.row, c);
    if (typeof v === 'number' && v !== 0) { lit = val1(it.sheet, 6, c); break; }
  }
  literals[`${it.sheet}!${it.row},${it.col}`] = lit;
}
// Compute rank literals: RANK(ref, range) -> 1 + count(range > ref)
for (const it of needLiteral.filter(x => x.kind === 'rank')) {
  const m = it.formula.match(/RANK\(([A-Z]+)(\d+),\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)/);
  if (!m) continue;
  const colNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const refCol = colNum(m[1]), refRow = +m[2], rCol = colNum(m[3]), r1 = +m[4], r2 = +m[6];
  const x = val1(it.sheet, refRow, refCol);
  if (typeof x !== 'number') { literals[`${it.sheet}!${it.row},${it.col}`] = 0; continue; }
  let cnt = 0;
  for (let r = r1; r <= r2; r++) {
    const v = val1(it.sheet, r, rCol);
    if (typeof v === 'number' && v > x) cnt++;
  }
  literals[`${it.sheet}!${it.row},${it.col}`] = cnt + 1;
}

// Pass 2: full build with literals
console.log('rebuilding with literals...');
const sheets2 = buildSheets(literals);
const hf = HyperFormula.buildFromSheets(sheets2, { licenseKey: 'gpl-v3' });
const val = (sheet, r, c) => {
  const v = hf.getCellValue({ sheet: hf.getSheetId(sheet), row: r - 1, col: c - 1 });
  if (v && typeof v === 'object' && v.value !== undefined) return v.value;
  return v;
};

// how many residual error cells overall?
let errCount = 0;
const errSamples = [];
for (const ws of wb.worksheets) {
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= ws.columnCount; c++) {
      const v = hf.getCellValue({ sheet: hf.getSheetId(ws.name), row: r - 1, col: c - 1 });
      if (v && typeof v === 'object' && v.value !== undefined && typeof v.value === 'string' && v.value.startsWith('#')) {
        errCount++;
        if (errSamples.length < 10) errSamples.push(`${ws.name}!${String.fromCharCode(64 + (c <= 26 ? c : 0)) || c}${r}=${v.value}`);
      }
    }
  }
}
console.log(`residual error cells after patching: ${errCount}`, errSamples.join(' | '));

// ---------- Control checks ----------
const ctrl = wb.getWorksheet('Control');
const ctrlVals = [];
for (let r = 1; r <= ctrl.rowCount; r++) {
  const label = strVal(ctrl, r, 2);
  if (label && /Check/.test(label) && label !== 'Check Summary') {
    ctrlVals.push({ label, val: val('Control', r, 3) });
  }
}
ctrlVals.forEach(x => console.log(`Control '${x.label}' = ${JSON.stringify(x.val)}`));
const qoqCheck = ctrlVals.find(x => x.label === 'Quarterly Retention (QoQ) Check');
check("Control: 'Quarterly Retention (QoQ) Check' exists and = 0",
  qoqCheck && qoqCheck.val === 0, `value=${JSON.stringify(qoqCheck && qoqCheck.val)}`);
for (const l of ['Quarterly Retention Check', 'Annual Retention Check']) {
  const row = ctrlVals.find(x => x.label === l);
  check(`Control: '${l}' = 0`, row && row.val === 0, `value=${JSON.stringify(row && row.val)}`);
}
const total = ctrlVals.find(x => x.label === 'Total Check (= 0)');
check("Control: 'Total Check (= 0)' = 0", total && total.val === 0, `value=${JSON.stringify(total && total.val)}`);
const otherBad = ctrlVals.filter(x => typeof x.val !== 'number' || Math.abs(x.val) > 1e-6);
check('Control: ALL check rows evaluate to 0', otherBad.length === 0,
  otherBad.map(x => `${x.label}=${JSON.stringify(x.val)}`).join('; '));

// ---------- QoQ annualized values + hand check ----------
let rBop = -1, rChurn = -1, rDown = -1, rUp = -1, rAnnL = -1, rAnnP = -1, rAnnN = -1;
for (let r = 1; r <= qoqWs.rowCount; r++) {
  const v = strVal(qoqWs, r, lc);
  if (v === 'BoP ARR' && rBop === -1) rBop = r;
  if (v === '(-) Churn' && rChurn === -1) rChurn = r;
  if (v === '(-) Downsell' && rDown === -1) rDown = r;
  if (v === '(+) Upsell / Cross-sell' && rUp === -1) rUp = r;
  if (v === '% Annualized Lost-Only Retention' && rAnnL === -1) rAnnL = r;
  if (v === '% Annualized Punitive Retention' && rAnnP === -1) rAnnP = r;
  if (v === '% Annualized Net Retention' && rAnnN === -1) rAnnN = r;
}
const dc = lc + 1;
const g = r => val('Quarterly Retention (QoQ)', r, dc);
const bop = g(rBop), ch = g(rChurn), dn = g(rDown), up = g(rUp);
const annNet = g(rAnnN);
const hdr = val('Quarterly Retention (QoQ)', rBop - 1, dc);
console.log(`block1 col${dc} '${hdr}': BoP=${bop} churn=${ch} down=${dn} up=${up} annNet=${annNet}`);
const numeric = [bop, ch, dn, up, annNet].every(x => typeof x === 'number');
check('QoQ annualized cells evaluate to numbers', numeric, `annNet=${annNet}`);
if (numeric) {
  const expected = (bop + (ch + dn + up) * 4) / bop;
  check('Hand-check: AnnNet = (BoP + (churn+down+up)*4)/BoP',
    Math.abs(expected - annNet) < 1e-9, `expected ${expected}, got ${annNet}`);
  check('QoQ annualized value plausible (0-200%)', annNet > -0.5 && annNet < 2, `annNet=${annNet}`);
}

const annSample = [];
for (let i = 0; i < 6; i++) {
  annSample.push({
    hdr: val('Quarterly Retention (QoQ)', rBop - 1, dc + i),
    lost: val('Quarterly Retention (QoQ)', rAnnL, dc + i),
    punit: val('Quarterly Retention (QoQ)', rAnnP, dc + i),
    net: val('Quarterly Retention (QoQ)', rAnnN, dc + i),
  });
}
console.log('Annualized sample:', JSON.stringify(annSample));

// Also verify QoQ retention values are NONTRIVIAL in a breakout block (not all-zero trivial check).
// Find second block's BoP row and check it's nonzero in some column.
{
  const bopRows = [];
  for (let r = 1; r <= qoqWs.rowCount; r++) if (strVal(qoqWs, r, lc) === 'BoP ARR') bopRows.push(r);
  let nonzero = 0, totalCells = 0;
  for (const br of bopRows) {
    for (let c = dc; c < dc + 15; c++) {
      const v = val('Quarterly Retention (QoQ)', br, c);
      if (typeof v === 'number') { totalCells++; if (v !== 0) nonzero++; }
    }
  }
  check('QoQ retention: breakout blocks produce nonzero evaluated values (checks are not trivially 0)',
    nonzero > bopRows.length * 5, `${nonzero}/${totalCells} nonzero BoP cells across ${bopRows.length} blocks`);
}

// ---------- HTML report ----------
const pct = v => typeof v === 'number' ? (v * 100).toFixed(1) + '%' : String(v);
const num = v => typeof v === 'number' ? v.toFixed(4) : String(v);
const html = `<!doctype html><html><head><meta charset="utf-8"><title>QoQ verification</title>
<style>body{font-family:system-ui;margin:32px;background:#f8fafc;color:#111}h1{font-size:20px}
table{border-collapse:collapse;margin:12px 0;background:#fff}td,th{border:1px solid #cbd5e1;padding:4px 10px;font-size:13px;text-align:left}
.pass{color:#15803d;font-weight:600}.fail{color:#b91c1c;font-weight:600}</style></head><body>
<h1>data-pack-output.xlsx &mdash; QoQ retention verification (HyperFormula recalc)</h1>
<h2>Checks</h2><table><tr><th>Check</th><th>Result</th><th>Detail</th></tr>
${results.map(r => `<tr><td>${r.name}</td><td class="${r.pass ? 'pass' : 'fail'}">${r.pass ? 'PASS' : 'FAIL'}</td><td>${String(r.detail).replace(/</g, '&lt;')}</td></tr>`).join('')}
</table>
<h2>Sheet list</h2><p>${sheetNames.join(' &middot; ')}</p>
<h2>Control checks (evaluated)</h2><table><tr><th>Label</th><th>Value</th></tr>
${ctrlVals.map(x => `<tr><td>${x.label}</td><td>${num(x.val)}</td></tr>`).join('')}
</table>
<h2>Quarterly Retention (QoQ) &mdash; block 1 annualized rows (evaluated)</h2>
<table><tr><th>Period</th><th>% Ann. Lost-Only</th><th>% Ann. Punitive</th><th>% Ann. Net</th></tr>
${annSample.map(s => `<tr><td>${s.hdr}</td><td>${pct(s.lost)}</td><td>${pct(s.punit)}</td><td>${pct(s.net)}</td></tr>`).join('')}
</table>
</body></html>`;
fs.writeFileSync('/tmp/xlverify/report.html', html);
console.log('\nWrote /tmp/xlverify/report.html');
const failed = results.filter(r => !r.pass);
console.log(`\n==== ${results.length - failed.length}/${results.length} checks passed ====`);
process.exit(failed.length ? 1 : 0);
