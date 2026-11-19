import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { resolveChartData, parseRangeRef, formatRangeRef, niceTicks } from '../ui/src/editors/sheets/chartData.js';
import { readChartsFromXlsx, writeChartsIntoXlsx, chartXml } from '../ui/src/editors/sheets/io/xlsxCharts.js';

// tiny getValues over an in-memory grid per sheet
const grids = {
  'Sales 2024': [['', 'North', 'South'], ['Jan', 10, 20], ['Feb', 15, 25], ['Mar', 12, 30]],
  Data: [['Fruit', 'Count'], ['Apple', 5], ['Pear', 3], ['Plum', 2]],
};
function getValues(ref) {
  const R = parseRangeRef(ref); if (!R) return null;
  const g = grids[R.sheet]; if (!g) return null;
  const values = [];
  for (let r = R.r1; r <= R.r2; r++) { const row = []; for (let c = R.c1; c <= R.c2; c++) row.push(g[r] ? g[r][c] ?? null : null); values.push(row); }
  return { values, texts: values.map((row) => row.map((v) => (v == null ? '' : String(v)))) };
}

test('range refs parse and format', () => {
  assert.deepEqual(parseRangeRef("'Sales 2024'!$A$1:$C$4"), { sheet: 'Sales 2024', r1: 0, c1: 0, r2: 3, c2: 2 });
  assert.equal(formatRangeRef({ sheet: "Bob's", r1: 1, c1: 1, r2: 3, c2: 1 }), "'Bob''s'!$B$2:$B$4");
  assert.equal(formatRangeRef({ sheet: 'Data', r1: 0, c1: 0, r2: 0, c2: 0 }), 'Data!$A$1');
});

test('resolveChartData: header row + category column', () => {
  const d = resolveChartData({ type: 'column', range: "'Sales 2024'!$A$1:$C$4" }, getValues);
  assert.deepEqual(d.categories, ['Jan', 'Feb', 'Mar']);
  assert.equal(d.series.length, 2);
  assert.equal(d.series[0].name, 'North');
  assert.deepEqual(d.series[1].values, [20, 25, 30]);
  assert.equal(d.series[0].refs.val, "'Sales 2024'!$B$2:$B$4");
  assert.equal(d.series[0].refs.cat, "'Sales 2024'!$A$2:$A$4");
  assert.equal(d.series[0].refs.name, "'Sales 2024'!$B$1");
});

test('resolveChartData: byRows', () => {
  const d = resolveChartData({ type: 'line', range: "'Sales 2024'!$A$1:$C$4", byRows: true }, getValues);
  assert.deepEqual(d.categories, ['North', 'South']);
  assert.deepEqual(d.series.map((s) => s.name), ['Jan', 'Feb', 'Mar']);
  assert.deepEqual(d.series[0].values, [10, 20]);
  assert.equal(d.series[0].refs.val, "'Sales 2024'!$B$2:$C$2");
});

test('resolveChartData: numbers only, scatter', () => {
  grids.N = [[1, 2], [2, 4], [3, 9]];
  const d = resolveChartData({ type: 'column', range: 'N!$A$1:$B$3' }, getValues);
  assert.deepEqual(d.categories, ['1', '2', '3']);
  assert.equal(d.series.length, 2);
  const s = resolveChartData({ type: 'scatter', range: 'N!$A$1:$B$3' }, getValues);
  assert.deepEqual(s.categories, [1, 2, 3]);
  assert.deepEqual(s.series[0].values, [2, 4, 9]);
  assert.equal(resolveChartData({ type: 'pie', range: 'Nope!$A$1' }, getValues).series.length, 0);
});

test('niceTicks', () => {
  const t = niceTicks(-3, 27, 5);
  assert.ok(t.min <= -3 && t.max >= 27);
  assert.ok(t.ticks.includes(0));
});

test('xlsx chart write → read round trip', async () => {
  const wb = new ExcelJS.Workbook();
  const s1 = wb.addWorksheet('Sales 2024');
  grids['Sales 2024'].forEach((row) => s1.addRow(row.map((v) => (v === '' ? null : v))));
  const s2 = wb.addWorksheet('Data');
  grids.Data.forEach((row) => s2.addRow(row));
  const buf = await wb.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buf);
  const col = { id: 'c1', type: 'column', stacked: 'stacked', title: 'Sales & more', legend: 'bottom', xTitle: 'Month', yTitle: 'Units', range: "'Sales 2024'!$A$1:$C$4", byRows: false, series: null,
    _xl: { from: { c: 4, r: 1, dx: 10, dy: 5 }, to: { c: 10, r: 16, dx: 0, dy: 0 } } };
  const pie = { id: 'c2', type: 'pie', title: 'Fruit', legend: 'right', range: 'Data!$A$1:$B$4', series: null, _xl: { from: { c: 3, r: 0, dx: 0, dy: 0 }, to: { c: 8, r: 12, dx: 0, dy: 0 } } };
  const line = { id: 'c3', type: 'scatter', title: '', legend: 'none', range: "'Sales 2024'!$B$2:$C$4", series: null, _xl: { from: { c: 0, r: 20, dx: 0, dy: 0 }, to: { c: 5, r: 30, dx: 0, dy: 0 } } };
  await writeChartsIntoXlsx(zip, [
    { sheetName: 'Sales 2024', charts: [col, line], data: [resolveChartData(col, getValues), resolveChartData(line, getValues)] },
    { sheetName: 'Data', charts: [pie], data: [resolveChartData(pie, getValues)] },
  ]);
  const out = await zip.generateAsync({ type: 'nodebuffer' });
  const ct = await zip.file('[Content_Types].xml').async('string');
  assert.match(ct, /drawingml\.chart\+xml/);
  const back = await readChartsFromXlsx(await JSZip.loadAsync(out));
  const a = back.get('Sales 2024');
  assert.equal(a.length, 2);
  assert.equal(a[0].type, 'column');
  assert.equal(a[0].stacked, 'stacked');
  assert.equal(a[0].title, 'Sales & more');
  assert.equal(a[0].legend, 'bottom');
  assert.equal(a[0].xTitle, 'Month');
  assert.equal(a[0].yTitle, 'Units');
  assert.equal(a[0].series.length, 2);
  assert.equal(a[0].series[1].valRef, "'Sales 2024'!$C$2:$C$4");
  assert.equal(a[0].series[1].catRef, "'Sales 2024'!$A$2:$A$4");
  assert.equal(a[0].series[0].nameRef, "'Sales 2024'!$B$1");
  assert.equal(a[0].series[0].name, 'North');
  assert.match(a[0].series[0].color, /^#[0-9a-f]{6}$/);
  assert.deepEqual(a[0].anchor, { from: { c: 4, r: 1, dx: 10, dy: 5 }, to: { c: 10, r: 16, dx: 0, dy: 0 } });
  assert.equal(a[1].type, 'scatter');
  assert.equal(a[1].legend, 'none');
  const p = back.get('Data');
  assert.equal(p[0].type, 'pie');
  assert.equal(p[0].title, 'Fruit');
  assert.equal(p[0].series[0].valRef, 'Data!$B$2:$B$4');
  // round trip through the editor's resolver again
  const d = resolveChartData(p[0], getValues);
  assert.deepEqual(d.categories, ['Apple', 'Pear', 'Plum']);
  assert.deepEqual(d.series[0].values, [5, 3, 2]);
  // exceljs still loads the file
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(out);
  assert.equal(wb2.worksheets.length, 2);
  assert.equal(wb2.getWorksheet('Data').getCell('A2').value, 'Apple');
});

test('chart xml for every type is well-formed-ish', () => {
  for (const type of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter']) {
    const c = { type, title: 'T <x>', legend: 'right', range: "'Sales 2024'!$A$1:$C$4" };
    const x = chartXml(c, resolveChartData(c, getValues));
    assert.match(x, /<c:chartSpace[\s\S]*<\/c:chartSpace>$/);
    assert.ok(!x.includes('T <x>'));
    // balanced tags
    const opens = (x.match(/<c:ser>/g) || []).length; const closes = (x.match(/<\/c:ser>/g) || []).length;
    assert.equal(opens, closes);
  }
});
