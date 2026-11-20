// Office 1.1 sheets: LAMBDA family, structured references, 3D references, intersection, R1C1, iterative calculation,
// 1904 dates, tables, pivot tables, sparklines, icon sets, pictures/shapes, rich text, sidecar – engine + xlsx round trip.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { Book } from '../ui/src/editors/sheets/model.js';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { isErr } from '../ui/src/editors/sheets/formula/errors.js';
import { toOoxml, fromOoxml, parse } from '../ui/src/editors/sheets/formula/parser.js';
import { toR1C1, fromR1C1 } from '../ui/src/editors/sheets/formula/r1c1.js';
import { readXlsx, writeXlsx } from '../ui/src/editors/sheets/io/xlsx.js';
import { addTable, toggleTotals } from '../ui/src/editors/sheets/tables.js';
import { computePivot, renderPivot, pivotArea } from '../ui/src/editors/sheets/pivot.js';
import { formatValue } from '../ui/src/editors/sheets/format/numfmt.js';

const loc = makeLocale({ lang: 'en' });
const col = (s) => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const at = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); return [+m[2] - 1, col(m[1])]; };
function book(data, sheets = 1) {
  const b = Book.blank({ loc });
  for (let i = 1; i < sheets; i++) b.addSheet(`Sheet${i + 1}`);
  b.run('t', () => { for (const [k, v] of Object.entries(data)) { const [sn, a] = k.includes('!') ? k.split('!') : ['Sheet1', k]; const [r, c] = at(a); b.input(b.sheetByName(sn), r, c, String(v)); } });
  return b;
}
const val = (b, a, sn = 'Sheet1') => { const [r, c] = at(a); const v = b.value(b.sheetByName(sn), r, c); return isErr(v) ? v.code : v; };
const set = (b, a, v, sn = 'Sheet1') => { const [r, c] = at(a); b.run('x', () => b.input(b.sheetByName(sn), r, c, String(v))); };

test('LAMBDA, LET, MAP, REDUCE, SCAN, MAKEARRAY, BYROW, BYCOL, ISOMITTED', () => {
  const b = book({ A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 });
  const f = {
    C1: '=LAMBDA(x,x*2)(5)', C2: '=LET(f,LAMBDA(x,y,x+y),f(2,3))', C3: '=REDUCE(0,A1:A3,LAMBDA(a,v,a+v))', C4: '=SUM(MAP(A1:A3,B1:B3,LAMBDA(a,b,a*b)))',
    C5: '=SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*c)))', C6: '=SUM(BYROW(A1:B3,LAMBDA(r,MAX(r))))', C7: '=INDEX(SCAN(0,A1:A3,LAMBDA(a,v,a+v)),3)',
    C8: '=SUM(BYCOL(A1:B3,LAMBDA(c,SUM(c))))', C9: '=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(4)', C10: '=LAMBDA(x,x)', C11: '=LET(x,2,y,x*3,x+y)',
  };
  for (const [k, v] of Object.entries(f)) set(b, k, v);
  assert.deepEqual(Object.keys(f).map((k) => val(b, k)), [10, 5, 6, 140, 18, 60, 6, 66, 4, '#CALC!', 8]);
  // MAP spills
  set(b, 'E1', '=MAP(A1:A3,LAMBDA(v,v*v))');
  assert.deepEqual([val(b, 'E1'), val(b, 'E2'), val(b, 'E3')], [1, 4, 9]);
  // named LAMBDA (recursive)
  b.setName('FACT2', 'LAMBDA(n,IF(n<=1,1,n*FACT2(n-1)))');
  set(b, 'D1', '=FACT2(5)');
  assert.equal(val(b, 'D1'), 120);
});

test('LAMBDA parameters are written as _xlpm. and read back', () => {
  const f = 'LET(x,1,f,LAMBDA(y,y+x),f(2))';
  const x = toOoxml(f);
  assert.equal(x, '_xlfn.LET(_xlpm.x,1,_xlpm.f,_xlfn.LAMBDA(_xlpm.y,_xlpm.y+_xlpm.x),_xlpm.f(2))');
  assert.equal(fromOoxml(x), f);
  assert.equal(toOoxml('MAP(A1:A3,LAMBDA(v,v*2))'), '_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))');
  assert.equal(toOoxml('SUM(A1#)'), 'SUM(_xlfn.ANCHORARRAY(A1))');
  assert.equal(fromOoxml('SUM(_xlfn.ANCHORARRAY(A1))'), 'SUM(A1#)');
});

test('3D references and the intersection operator', () => {
  const b = book({ 'Sheet1!A1': 1, 'Sheet2!A1': 2, 'Sheet3!A1': 3, 'Sheet1!B1': '=SUM(Sheet1:Sheet3!A1)', 'Sheet1!B2': "=COUNT('Sheet1:Sheet3'!A1:A2)" }, 3);
  assert.equal(val(b, 'B1'), 6); assert.equal(val(b, 'B2'), 3);
  set(b, 'A1', 10, 'Sheet2');
  assert.equal(val(b, 'B1'), 14);
  b.renameSheet(b.sheetByName('Sheet3'), 'Last');
  assert.equal(b.cell(b.sheets[0], 0, 1).f, 'SUM(Sheet1:Last!A1)');
  set(b, 'E1', '=A1:A5 A1:C1'); set(b, 'E2', '=SUM(A1:B5 B1:C5)'); set(b, 'E3', '=A1 B2');
  assert.equal(val(b, 'E1'), 1); assert.equal(val(b, 'E2'), 14 + 3); assert.equal(val(b, 'E3'), '#NULL!');
});

test('tables and structured references (+ totals row, OOXML form)', () => {
  const b = book({ A1: 'Item', B1: 'Qty', C1: 'Price', A2: 'a', B2: 2, C2: 1.5, A3: 'b', B3: 4, C3: 2, A4: 'c', B4: 1, C4: 10 });
  const s = b.sheets[0];
  const t = addTable(b, s, { r1: 0, c1: 0, r2: 3, c2: 2 });
  assert.equal(t.name, 'Table1');
  set(b, 'E1', '=SUM(Table1[Qty])'); set(b, 'E2', '=COUNTA(Table1[#Headers])'); set(b, 'E3', '=ROWS(Table1[#All])'); set(b, 'E4', '=SUM(Table1[[Qty]:[Price]])');
  set(b, 'D2', '=[@Qty]*[@Price]'); set(b, 'D3', '=Table1[@Qty]*Table1[@Price]');
  assert.deepEqual([val(b, 'E1'), val(b, 'E2'), val(b, 'E3'), val(b, 'E4'), val(b, 'D2'), val(b, 'D3')], [7, 3, 4, 20.5, '#REF!', 8]);
  set(b, 'B2', 5);
  assert.equal(val(b, 'E1'), 10);
  toggleTotals(b, s, t.id, true);
  const tb = s.tables[0];
  assert.equal(tb.totals, true); assert.deepEqual(tb.range, { r1: 0, c1: 0, r2: 4, c2: 2 });
  assert.equal(val(b, 'C5'), 13.5); // SUBTOTAL(109,Table1[Price])
  assert.equal(val(b, 'E5'), null);
  set(b, 'F1', '=Table1[[#Totals],[Price]]');
  assert.equal(val(b, 'F1'), 13.5);
  assert.equal(toOoxml('[@Qty]*2', { table: 'Table1' }), 'Table1[[#This Row],[Qty]]*2');
  assert.equal(fromOoxml('Table1[[#This Row],[Qty]]*2'), 'Table1[@Qty]*2');
  assert.equal(toOoxml("Table1[[#Headers],[Unit Price]:[Qty]]"), 'Table1[[#Headers],[Unit Price]:[Qty]]');
  assert.ok(parse('SUM(Table1[[#Data],[Qty]])'));
});

test('R1C1 display and input', () => {
  assert.equal(toR1C1('SUM(A1:B$2)+$C$3', 2, 3), 'SUM(R[-2]C[-3]:R2C[-2])+R3C3');
  assert.equal(fromR1C1('SUM(R[-2]C[-3]:R2C[-2])+R3C3', 2, 3), 'SUM(A1:B$2)+$C$3');
  assert.equal(fromR1C1('RC[-1]*2+ROUND(R[1]C,0)', 4, 4), 'D5*2+ROUND(E6,0)');
  assert.equal(fromR1C1('SUM(C[-1])+SUM(R1:R2)', 0, 1), 'SUM(A:A)+SUM($1:$2)');
});

test('iterative calculation resolves circular references', () => {
  const b = book({ A1: '=B1+1', B1: '=IF(A1>10,10,A1)' });
  assert.equal(val(b, 'A1'), '#CIRC!');
  b.iterate = { on: true, count: 100, delta: 0.001 };
  b.calc.recalc(null);
  assert.equal(val(b, 'A1'), 11); assert.equal(val(b, 'B1'), 10);
  const c = book({ A1: 100, B1: '=A1+C1*0.5', C1: '=B1*0.1' });
  c.iterate = { on: true, count: 200, delta: 1e-9 };
  c.calc.recalc(null);
  assert.ok(Math.abs(val(c, 'B1') - 100 / 0.95) < 1e-6);
});

test('1904 date system shifts dates', async () => {
  const b = Book.blank({ loc });
  b.date1904 = true;
  b.run('d', () => b.input(b.sheets[0], 0, 0, '=DATE(1904,1,2)'));
  assert.equal(val(b, 'A1'), 1);
  assert.equal(formatValue(0, 'yyyy-mm-dd', loc).text, '1904-01-01');
  const bytes = await writeXlsx(b);
  const zip = await JSZip.loadAsync(bytes);
  assert.match(await zip.file('xl/workbook.xml').async('string'), /date1904="1"/);
  const r = await readXlsx(bytes, { loc });
  assert.equal(r.date1904, true);
  assert.equal(formatValue(r.value(r.sheets[0], 0, 0), 'yyyy-mm-dd', loc).text, '1904-01-02');
  const n = Book.blank({ loc }); // a new 1900 book switches back
  assert.equal(n.date1904, false);
  assert.equal(formatValue(1, 'yyyy-mm-dd', loc).text, '1900-01-01');
});

function salesBook() {
  const rows = [['Region', 'Product', 'Sales', 'Qty'], ['East', 'A', 10, 1], ['West', 'A', 20, 2], ['East', 'B', 5, 3], ['West', 'B', 7, 1], ['East', 'A', 3, 2], ['North', 'C', 8, 5]];
  const b = Book.blank({ loc });
  const s = b.sheets[0]; s.name = 'Data';
  b.run('d', () => rows.forEach((row, r) => row.forEach((v, c) => b.input(s, r, c, String(v)))));
  return b;
}
test('pivot table: calculation, layout and cells', () => {
  const b = salesBook();
  const out = b.addSheet('Pivot');
  const pv = { id: 'p1', name: 'PivotTable1', source: 'Data!$A$1:$D$7', at: { r: 2, c: 0 }, rows: ['Region'], cols: ['Product'], values: [{ field: 'Sales', fn: 'sum' }], filters: [] };
  const res = computePivot(b, pv);
  assert.equal(res.error, undefined);
  const grid = res.rows.map((r) => r.map((x) => (x ? x.v : null)));
  assert.deepEqual(grid[0].slice(0, 2), ['Sum of Sales', 'Product']);
  assert.deepEqual(grid[1], ['Region', 'A', 'B', 'C', 'Grand Total']);
  assert.deepEqual(grid[2], ['East', 13, 5, null, 18]);
  assert.deepEqual(grid[4], ['West', 20, 7, null, 27]);
  assert.deepEqual(grid[5], ['Grand Total', 33, 12, 8, 53]);
  b.run('p', () => { b.touch(out, true); const area = renderPivot(b, out, pv, res); pv.range = area; out.pivots.push(pv); });
  assert.equal(b.value(out, 4, 1), 13);
  assert.deepEqual(pivotArea(pv, res), { r1: 2, c1: 0, r2: 7, c2: 4 });
  // filters + several values
  const pv2 = { ...pv, cols: [], values: [{ field: 'Sales', fn: 'sum' }, { field: 'Qty', fn: 'average' }], filters: [{ field: 'Product', exclude: ['C'] }] };
  const r2 = computePivot(b, pv2);
  const g2 = r2.rows.map((r) => r.map((x) => (x ? x.v : null)));
  assert.deepEqual(g2[0], [null, 'Values', null]); // like Excel: the value fields are a column level
  assert.deepEqual(g2[1], ['Region', 'Sum of Sales', 'Average of Qty']);
  assert.deepEqual(g2.slice(2).map((r) => r[0]), ['East', 'West', 'Grand Total']);
  assert.equal(r2.loc.rowOff, 2);
});

test('xlsx: pivot parts are written and read back (refreshOnLoad, records, workbook cache)', async () => {
  const b = salesBook();
  const out = b.addSheet('Pivot');
  const pv = { id: 'p1', name: 'PivotTable1', source: 'Data!$A$1:$D$7', at: { r: 2, c: 0 }, rows: ['Region'], cols: [], values: [{ field: 'Sales', fn: 'sum' }], filters: [{ field: 'Product', exclude: [] }] };
  b.run('p', () => { b.touch(out, true); const res = computePivot(b, pv); pv.range = renderPivot(b, out, pv, res); out.pivots.push(pv); });
  const bytes = await writeXlsx(b);
  const zip = await JSZip.loadAsync(bytes);
  const wb = await zip.file('xl/workbook.xml').async('string');
  assert.match(wb, /<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/);
  const cache = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string');
  assert.match(cache, /refreshOnLoad="1"/); assert.match(cache, /<worksheetSource ref="A1:D7" sheet="Data"\/>/);
  assert.match(cache, /<cacheField name="Region" numFmtId="0"><sharedItems count="3"><s v="East"\/><s v="North"\/><s v="West"\/><\/sharedItems>/);
  const recs = await zip.file('xl/pivotCache/pivotCacheRecords1.xml').async('string');
  assert.match(recs, /count="6"/);
  const tbl = await zip.file('xl/pivotTables/pivotTable1.xml').async('string');
  assert.match(tbl, /<location ref="A5:B9" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"\/>/);
  assert.match(tbl, /<dataField name="Sum of Sales" fld="2" baseField="0" baseItem="0"\/>/);
  const ct = await zip.file('[Content_Types].xml').async('string');
  assert.match(ct, /pivotTable\+xml/); assert.match(ct, /pivotCacheDefinition\+xml/);
  const r = await readXlsx(bytes, { loc });
  const p = r.sheetByName('Pivot').pivots[0];
  assert.deepEqual([p.name, p.rows, p.values[0].field, p.filters[0].field, p.source, p.readOnly], ['PivotTable1', ['Region'], 'Sales', 'Product', 'Data!$A$1:$D$7', false]);
  // unchanged pivot → parts byte-identical on the next save
  const again = await JSZip.loadAsync(await writeXlsx(r));
  assert.equal(await again.file('xl/pivotTables/pivotTable1.xml').async('string'), tbl);
  // deleting it removes the parts and the workbook entry
  r.sheetByName('Pivot').pivots = [];
  const z3 = await JSZip.loadAsync(await writeXlsx(r));
  assert.equal(z3.file('xl/pivotTables/pivotTable1.xml'), null);
  assert.doesNotMatch(await z3.file('xl/workbook.xml').async('string'), /pivotCache/);
});

test('xlsx: tables, structured references, sparklines, icon sets, pictures, shapes, rich text, iteration, sidecar', async () => {
  const b = book({ A1: 'Item', B1: 'Qty', C1: 'Double', A2: 'a', B2: 2, A3: 'b', B3: 4, D1: 1, E1: 3, F1: 2 });
  const s = b.sheets[0];
  addTable(b, s, { r1: 0, c1: 0, r2: 2, c2: 2 });
  set(b, 'C2', '=[@Qty]*2'); set(b, 'C5', '=SUM(Table1[Qty])');
  s.sparklines = [{ id: 'g', type: 'column', color: '#376092', neg: '#D00000', markerColor: '#D00000', highColor: '#D00000', lowColor: '#D00000', high: true, items: [{ ref: 'Sheet1!D1:F1', r: 0, c: 6 }] }];
  s.cf = [{ id: 'i', ranges: [{ r1: 0, c1: 3, r2: 0, c2: 5 }], type: 'icons', set: '3Arrows', stops: [], priority: 1 }];
  s.images = [{ id: 'im', src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', anchor: { r: 6, c: 1, dx: 0, dy: 0, w: 120, h: 80 } }];
  s.shapes = [{ id: 'sh', kind: 'roundRect', anchor: { r: 6, c: 4, dx: 0, dy: 0, w: 150, h: 60 }, fill: '#4472C4', line: '#2F528F', text: 'Hello' }];
  b.run('r', () => b.setCell(s, 10, 0, { v: 'Bold red', rich: [{ t: 'Bold', b: true }, { t: ' red', color: '#FF0000' }] }));
  b.iterate = { on: true, count: 50, delta: 0.01 };
  s.view = { type: 'kanban', kanban: { col: 0 } };
  const bytes = await writeXlsx(b);
  const zip = await JSZip.loadAsync(bytes);
  const sx = await zip.file('xl/worksheets/sheet1.xml').async('string');
  assert.match(sx, /<tableParts count="1"><tablePart r:id="rId\d+"\/><\/tableParts>/);
  assert.match(sx, /<f>Table1\[\[#This Row\],\[Qty\]\]\*2<\/f>/);
  assert.match(sx, /<x14:sparklineGroup type="column"[^>]* high="1"/); assert.match(sx, /<xm:f>Sheet1!D1:F1<\/xm:f><xm:sqref>G1<\/xm:sqref>/);
  assert.match(sx, /<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"\/>/);
  const tx = await zip.file('xl/tables/table1.xml').async('string');
  assert.match(tx, /<table [^>]*id="1" name="Table1" displayName="Table1" ref="A1:C3"/);
  const dr = await zip.file('xl/drawings/drawing1.xml').async('string');
  assert.match(dr, /<xdr:pic>/); assert.match(dr, /<a:prstGeom prst="roundRect">/); assert.match(dr, /<a:t>Hello<\/a:t>/);
  assert.ok(zip.file('xl/media/image1.png'));
  assert.match(await zip.file('xl/sharedStrings.xml').async('string'), /<si><r><rPr><b\/><\/rPr><t xml:space="preserve">Bold<\/t><\/r><r><rPr><color rgb="FFFF0000"\/><\/rPr><t xml:space="preserve"> red<\/t><\/r><\/si>/);
  assert.match(await zip.file('xl/workbook.xml').async('string'), /<calcPr iterate="1" iterateCount="50" iterateDelta="0.01"/);
  assert.match(await zip.file('customXml/item1.xml').async('string'), /"kanban"/);
  // read back
  const r = await readXlsx(bytes, { loc });
  const rs = r.sheets[0];
  assert.equal(rs.tables[0].name, 'Table1'); assert.equal(r.value(rs, 1, 2), 4); assert.equal(r.value(rs, 4, 2), 6);
  assert.equal(rs.cells.get(1 * 16384 + 2).f, 'Table1[@Qty]*2');
  assert.equal(rs.sparklines[0].type, 'column'); assert.equal(rs.sparklines[0].items[0].c, 6);
  assert.equal(rs.cf[0].type, 'icons'); assert.equal(rs.cf[0].set, '3Arrows');
  assert.equal(rs.images.length, 1); assert.equal(rs.shapes[0].text, 'Hello'); assert.equal(rs.shapes[0].kind, 'roundRect');
  assert.deepEqual(rs.cells.get(10 * 16384).rich.map((x) => x.t), ['Bold', ' red']);
  assert.deepEqual(r.iterate, { on: true, count: 50, delta: 0.01 });
  assert.equal(rs.view.type, 'kanban');
  // untouched → drawing/table/ext byte-identical; moved picture patched in place
  const z2 = await JSZip.loadAsync(await writeXlsx(r));
  assert.equal(await z2.file('xl/drawings/drawing1.xml').async('string'), dr);
  assert.equal(await z2.file('xl/tables/table1.xml').async('string'), tx);
  rs.images[0].anchor = { ...rs.images[0].anchor, r: 9, c: 2 };
  rs.shapes = [];
  const z3 = await JSZip.loadAsync(await writeXlsx(r));
  const d3 = await z3.file('xl/drawings/drawing1.xml').async('string');
  assert.match(d3, /<xdr:from><xdr:col>2<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>9<\/xdr:row>/);
  assert.doesNotMatch(d3, /roundRect/);
  // table resize on row insert
  r.insert(rs, 'row', 1, 2);
  assert.deepEqual(rs.tables[0].range, { r1: 0, c1: 0, r2: 4, c2: 2 });
});

test('charts keep data labels and trendlines; compare diff; old codepages', async () => {
  const { chartXml } = await import('../ui/src/editors/sheets/io/xlsxCharts.js');
  const x = chartXml({ type: 'column', title: 'T', labels: true, trend: true, legend: 'none' }, { categories: ['a', 'b', 'c'], series: [{ name: 'S', values: [1, 3, 2], color: '#4472C4' }] });
  assert.match(x, /<c:invertIfNegative val="0"\/><c:dLbls>.*<c:showVal val="1"\/>.*<\/c:dLbls><c:trendline>.*<c:trendlineType val="linear"\/>.*<\/c:trendline><c:cat>/);
  const b = book({ A1: 'x', B1: 1, C1: 2 });
  b.setSheetProps(b.sheets[0], { charts: [{ id: 'c', type: 'line', title: 'L', legend: 'none', range: 'Sheet1!$B$1:$C$1', byRows: true, labels: true, series: null, anchor: { r: 3, c: 0, dx: 0, dy: 0, w: 300, h: 200 } }] });
  const r = await readXlsx(await writeXlsx(b), { loc });
  assert.equal(r.sheets[0].charts[0].labels || r.sheets[0].charts[0].series.some((s) => s.labels), true);
  // compare
  const { diffBooks } = await import('../ui/src/editors/sheets/plus/diff.js');
  if (diffBooks) {
    const a = book({ A1: 1, A2: 'x' }); const c = book({ A1: 2, A3: 'new' });
    const d = diffBooks(a, c).sheets[0].cells.map((q) => `${q.kind}:${q.r},${q.c}`);
    assert.deepEqual(d, ['changed:0,0', 'removed:1,0', 'added:2,0']);
  }
  // codepages (TextDecoder-backed) for .xls/.dbf
  const { cptable } = await import('../ui/src/editors/sheets/io/codepage.js');
  assert.equal(cptable.utils.decode(437, [0x81, 0x9a, 0xe1]), 'üÜß');
  assert.equal(cptable.utils.decode(1251, [0xcf, 0xf0, 0xe8]), 'При');
  assert.equal(cptable.utils.decode(932, [0x93, 0xfa, 0x96, 0x7b]), '日本');
  assert.deepEqual(cptable.utils.encode(850, 'ø'), [0x9b]);
});
