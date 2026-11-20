import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { Book } from '../ui/src/editors/sheets/model.js';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { readOds, writeOds } from '../ui/src/editors/sheets/io/ods.js';
import { dataStyleXml, readDataStyles } from '../ui/src/editors/sheets/io/ods-styles.js';
import { parseXml } from '../ui/src/editors/_shared/xmltree.js';
import { openBytes, saveBytes } from '../ui/src/editors/sheets/io/formats.js';

const loc = makeLocale({ lang: 'en' });
const K = (r, c) => r * 16384 + c;
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function styled() {
  const b = Book.blank({ loc });
  const s = b.sheets[0];
  b.renameSheet(s, 'Report');
  const d = b.addSheet('Data 2');
  b.run('fill', () => {
    b.input(s, 0, 0, 'Title'); b.input(s, 1, 0, '1234.5'); b.input(s, 2, 0, '0.25'); b.input(s, 3, 0, '45000'); b.input(s, 4, 0, '=A2*2');
    b.input(s, 5, 0, '=\'Data 2\'!B2+1'); b.input(d, 1, 1, '41');
    b.input(s, 1, 1, 'Apple'); b.input(s, 2, 1, 'Pear'); b.input(s, 3, 1, 'Plum');
    b.input(s, 1, 2, '10'); b.input(s, 2, 2, '20'); b.input(s, 3, 2, '30');
    b.input(s, 6, 0, 'TRUE');
  });
  const set = (r, c, st) => { const cell = s.cells.get(K(r, c)); s.cells.set(K(r, c), { ...cell, s: b.styles.intern(st) }); };
  set(0, 0, { b: true, i: true, size: 16, font: 'Georgia', color: '#C00000', fill: '#FFF2CC', h: 'center', va: 'center', wrap: true, bd: { t: { s: 'thin', c: '#000000' }, b: { s: 'double', c: '#1F4E79' }, l: { s: 'medium', c: '#FF0000' }, r: { s: 'dashed', c: '#00B050' } } });
  set(1, 0, { nf: '#,##0.00', u: true, st: true });
  set(2, 0, { nf: '0.0%' });
  set(3, 0, { nf: 'dd.mm.yyyy', rot: 45, indent: 2 });
  set(4, 0, { nf: '#,##0.00 [$€];[Red]-#,##0.00 [$€]', unlocked: true });
  s.merges.push({ r1: 0, c1: 0, r2: 0, c2: 2 });
  s.colW.set(0, 140); s.colW.set(3, 30); s.hiddenCols.add(4);
  s.rowH.set(0, 36); s.hiddenRows.add(8);
  s.freeze = { rows: 1, cols: 1 };
  s.zoom = 125; s.gridlines = false; s.color = '#FF0000';
  s.cf.push({ id: 'c1', ranges: [{ r1: 1, c1: 2, r2: 3, c2: 2 }], type: 'cell', op: 'gt', f1: '15', style: { fill: '#FFC7CE', color: '#9C0006' } });
  s.cf.push({ id: 'c2', ranges: [{ r1: 1, c1: 2, r2: 3, c2: 2 }], type: 'scale', colors: ['#F8696B', '#63BE7B'] });
  s.dv.push({ id: 'v1', ranges: [{ r1: 1, c1: 1, r2: 5, c2: 1 }], type: 'list', f1: '"Apple,Pear,Plum"', allowBlank: true, dropdown: true, prompt: { title: 'Fruit', text: 'Pick one' }, error: { style: 'warning', title: 'Oops', text: 'Not a fruit', show: true } });
  s.dv.push({ id: 'v2', ranges: [{ r1: 1, c1: 2, r2: 3, c2: 2 }], type: 'whole', op: 'between', f1: '1', f2: '100', allowBlank: true, error: { style: 'stop', show: true } });
  s.cells.set(K(1, 1), { ...s.cells.get(K(1, 1)), note: 'A fruit\nsecond line', noteAuthor: 'Ann', link: 'https://example.com/' });
  s.cells.set(K(9, 0), { v: 'Bold and plain', rich: [{ t: 'Bold', b: true, color: '#FF0000' }, { t: ' and plain' }] });
  s.images.push({ id: 'i1', src: PNG, anchor: { r: 10, c: 1, dx: 4, dy: 2, w: 60, h: 40 } });
  s.charts.push({ id: 'ch1', type: 'column', stacked: false, title: 'Fruit', legend: 'bottom', xTitle: 'Kind', yTitle: 'Count', byRows: false, range: null, series: [{ name: 'Count', nameRef: null, catRef: 'Report!$B$2:$B$4', valRef: 'Report!$C$2:$C$4', color: '#2563EB', labels: true }], anchor: { r: 12, c: 3, dx: 0, dy: 0, w: 400, h: 240 } });
  b.setName('Fruits', 'Report!$B$2:$B$4');
  s.print.orientation = 'landscape';
  s.filter = { range: { r1: 0, c1: 1, r2: 3, c2: 2 }, cols: {} };
  b.calc.rebuild(); b.calc.recalc(null, { force: true });
  return b;
}

test('data styles: Excel codes ↔ ODF number styles', () => {
  const back = (code) => { const x = dataStyleXml(code, 'N1', loc); const root = parseXml(`<r xmlns:number="n" xmlns:style="s">${x.xml}</r>`); return readDataStyles([root]).get('N1'); };
  assert.equal(back('#,##0.00'), '#,##0.00');
  assert.equal(back('0%'), '0%');
  assert.equal(back('0.00%'), '0.00%');
  assert.equal(back('dd.mm.yyyy'), 'dd\\.mm\\.yyyy');
  assert.equal(back('hh:mm:ss'), 'hh\\:mm\\:ss');
  assert.equal(back('[h]:mm'), '[h]\\:mm');
  assert.equal(back('0.00E+00'), '0.00E+00');
  assert.equal(back('#,##0;[Red]-#,##0'), '#,##0;[Red]-#,##0');
  assert.equal(back('0;-0;"zero"'), '0;-0;"zero"');
  assert.equal(back('@'), '@');
  assert.equal(back('mmmm d, yyyy'), 'mmmm d, yyyy');
  assert.equal(dataStyleXml('General', 'N2', loc).general, true);
  assert.equal(dataStyleXml('dd.mm.yyyy', 'N3', loc).kind, 'date');
  assert.equal(dataStyleXml('h:mm', 'N4', loc).kind, 'time');
  assert.equal(dataStyleXml('0%', 'N5', loc).kind, 'percentage');
});

for (const ext of ['ods', 'fods']) {
  test(`${ext}: styles, sizes, freeze, CF, validation, notes, images and charts survive a round trip`, async () => {
    const b = styled();
    const bytes = await writeOds(b, { ext });
    if (ext === 'ods') {
      assert.equal(new TextDecoder().decode(bytes.subarray(30, 38)), 'mimetype');
      assert.equal(bytes[8], 0);
      const z = await JSZip.loadAsync(bytes);
      const man = await z.file('META-INF/manifest.xml').async('string');
      assert.match(man, /Object 1\/" manifest:version="1.2" manifest:media-type="application\/vnd.oasis.opendocument.chart"/);
      assert.match(man, /Pictures\/hw1/);
      const content = await z.file('content.xml').async('string');
      assert.match(content, /style:map style:condition="cell-content\(\)&gt;15"/);
      assert.match(content, /calcext:color-scale/);
      assert.match(content, /table:condition="of:cell-content-is-in-list\(&quot;Apple&quot;;&quot;Pear&quot;;&quot;Plum&quot;\)"/);
      assert.match(content, /of:=\[\$'Data 2'\.B2\]\+1/);
      const settings = await z.file('settings.xml').async('string');
      assert.match(settings, /VerticalSplitPosition" config:type="int">1</);
    }
    const r = await readOds(bytes, { loc, ext });
    const s = r.sheets[0]; const d = r.sheets[1];
    assert.equal(s.name, 'Report'); assert.equal(d.name, 'Data 2');
    const st = (rr, c) => s.cells.get(K(rr, c)).s || {};
    const t = st(0, 0);
    assert.equal(t.b, true); assert.equal(t.i, true); assert.equal(t.size, 16); assert.equal(t.font, 'Georgia'); assert.equal(t.color, '#C00000'); assert.equal(t.fill, '#FFF2CC');
    assert.equal(t.h, 'center'); assert.equal(t.va, 'center'); assert.equal(t.wrap, true);
    assert.deepEqual(t.bd.t, { s: 'thin', c: '#000000' }); assert.deepEqual(t.bd.b, { s: 'double', c: '#1F4E79' }); assert.deepEqual(t.bd.l, { s: 'medium', c: '#FF0000' }); assert.deepEqual(t.bd.r, { s: 'dashed', c: '#00B050' });
    assert.equal(st(1, 0).nf, '#,##0.00'); assert.equal(st(1, 0).u, true); assert.equal(st(1, 0).st, true);
    assert.equal(st(2, 0).nf, '0.0%');
    assert.equal(st(3, 0).nf, 'dd\\.mm\\.yyyy'); assert.equal(st(3, 0).rot, 45); assert.equal(st(3, 0).indent, 2);
    assert.equal(st(4, 0).nf, '#,##0.00 €;[Red]-#,##0.00 €'); assert.equal(st(4, 0).unlocked, true);
    assert.equal(r.text(s, 1, 0), '1,234.50');
    assert.equal(r.text(s, 3, 0), '15.03.2023');
    assert.equal(r.value(s, 4, 0), 2469);
    assert.equal(r.value(s, 5, 0), 42);
    assert.equal(r.value(s, 6, 0), true);
    assert.deepEqual(s.merges, [{ r1: 0, c1: 0, r2: 0, c2: 2 }]);
    assert.equal(s.colW.get(0), 140); assert.equal(s.colW.get(3), 30); assert.ok(s.hiddenCols.has(4));
    assert.equal(s.rowH.get(0), 36); assert.ok(s.hiddenRows.has(8));
    assert.deepEqual(s.freeze, { rows: 1, cols: 1 }); assert.equal(s.zoom, 125); assert.equal(s.gridlines, false); assert.equal(s.color, '#FF0000');
    const cf = s.cf.find((x) => x.type === 'cell');
    assert.equal(cf.op, 'gt'); assert.equal(cf.f1, '15'); assert.equal(cf.style.fill, '#FFC7CE'); assert.deepEqual(cf.ranges, [{ r1: 1, c1: 2, r2: 3, c2: 2 }]);
    assert.ok(s.cf.find((x) => x.type === 'scale'));
    const list = s.dv.find((x) => x.type === 'list');
    assert.equal(list.f1, '"Apple,Pear,Plum"'); assert.deepEqual(list.ranges, [{ r1: 1, c1: 1, r2: 5, c2: 1 }]); assert.equal(list.prompt.text, 'Pick one'); assert.equal(list.error.style, 'warning');
    const whole = s.dv.find((x) => x.type === 'whole'); assert.equal(whole.op, 'between'); assert.equal(whole.f2, '100');
    const n = s.cells.get(K(1, 1)); assert.equal(n.note, 'A fruit\nsecond line'); assert.equal(n.noteAuthor, 'Ann'); assert.equal(n.link, 'https://example.com/');
    assert.deepEqual(s.cells.get(K(9, 0)).rich, [{ t: 'Bold', b: true, color: '#FF0000' }, { t: ' and plain' }]);
    assert.equal(s.images.length, 1); assert.equal(s.images[0].anchor.r, 10); assert.equal(s.images[0].anchor.c, 1); assert.equal(s.images[0].src, PNG);
    assert.equal(s.charts.length, 1);
    const ch = s.charts[0];
    assert.equal(ch.type, 'column'); assert.equal(ch.title, 'Fruit'); assert.equal(ch.legend, 'bottom'); assert.equal(ch.xTitle, 'Kind'); assert.equal(ch.yTitle, 'Count');
    assert.equal(ch.series[0].valRef, 'Report!$C$2:$C$4'); assert.equal(ch.series[0].catRef, 'Report!$B$2:$B$4'); assert.equal(ch.series[0].color, '#2563eb'); assert.equal(ch.series[0].labels, true);
    assert.equal(ch.anchor.r, 12); assert.equal(ch.anchor.c, 3);
    assert.equal([...r.names.values()].find((x) => x.name === 'Fruits').ref, 'Report!$B$2:$B$4');
    assert.equal(s.print.orientation, 'landscape');
    assert.deepEqual(s.filter.range, { r1: 0, c1: 1, r2: 3, c2: 2 });
  });
}

test('ods patch saving keeps unknown parts and unchanged charts byte-identical', async () => {
  const b = styled();
  const first = await writeOds(b, { ext: 'ods' });
  const z0 = await JSZip.loadAsync(first);
  z0.file('Basic/Standard/Module1.xml', '<module>macro</module>');
  z0.file('Thumbnails/thumbnail.png', new Uint8Array([1, 2, 3]));
  const objOrig = await z0.file('Object 1/content.xml').async('string');
  const withExtra = await z0.generateAsync({ type: 'uint8array' });
  const r = await readOds(withExtra, { loc, ext: 'ods' });
  r.input(r.sheets[0], 20, 0, 'new');
  const out = await writeOds(r, { ext: 'ods' });
  const z = await JSZip.loadAsync(out);
  assert.equal(await z.file('Basic/Standard/Module1.xml').async('string'), '<module>macro</module>');
  assert.deepEqual([...await z.file('Thumbnails/thumbnail.png').async('uint8array')], [1, 2, 3]);
  assert.equal(await z.file('Object 1/content.xml').async('string'), objOrig);
  assert.equal(await z.file('styles.xml').async('string'), await z0.file('styles.xml').async('string'));
  // editing the chart rewrites only its object
  const r2 = await readOds(out, { loc, ext: 'ods' });
  r2.sheets[0].charts[0].title = 'Changed';
  const z2 = await JSZip.loadAsync(await writeOds(r2, { ext: 'ods' }));
  assert.match(await z2.file('Object 1/content.xml').async('string'), /<text:p>Changed<\/text:p>/);
  const r3 = await readOds(await z2.generateAsync({ type: 'uint8array' }), { loc, ext: 'ods' });
  assert.equal(r3.sheets[0].charts[0].title, 'Changed');
  assert.equal(r3.text(r3.sheets[0], 20, 0), 'new');
});

test('formats.js uses the own OpenDocument reader/writer (no engine needed)', async () => {
  const b = styled();
  const { bytes, note } = await saveBytes(b, 'ods', {});
  assert.equal(note, undefined);
  const r = await openBytes(bytes, 'ods', { loc });
  assert.equal(r.sheets[0].cells.get(K(0, 0)).s.b, true);
  assert.deepEqual(r.fidelity, []);
});

test('Excel-made ods: chart in a cell (relative position) is read', async () => {
  const fs = await import('node:fs');
  const b = await readOds(fs.readFileSync(new URL('./fixtures/excel-chart.ods', import.meta.url)), { loc, ext: 'ods' });
  const ch = b.sheets[0].charts[0];
  assert.equal(ch.title, 'Fruit');
  assert.equal(ch.series[0].valRef, 'Tabelle1!$B$2:$B$4');
  assert.equal(ch.series[0].catRef, 'Tabelle1!$A$2:$A$4');
  assert.equal(ch.anchor.c, 3);
  assert.deepEqual(b.props.defaultFont, { font: 'Calibri', size: 11 });
});

test('xls (BIFF8): fonts, fills, borders and alignment are read without the engine', async () => {
  const fs = await import('node:fs');
  const { readSheetJS } = await import('../ui/src/editors/sheets/io/sheetjs.js');
  const b = readSheetJS(fs.readFileSync(new URL('./fixtures/excel-styled.xls', import.meta.url)), { loc, ext: 'xls' });
  const s = b.sheets[0];
  const a1 = s.cells.get(K(0, 0)).s;
  assert.equal(a1.font, 'Georgia'); assert.equal(a1.size, 16); assert.equal(a1.b, true); assert.equal(a1.i, true);
  assert.equal(a1.h, 'center'); assert.equal(a1.va, 'center'); assert.equal(a1.wrap, true);
  assert.equal(a1.bd.t.s, 'thin'); assert.equal(a1.bd.b.s, 'double'); assert.equal(a1.bd.l.s, 'medium');
  assert.equal(a1.fill, '#FFFFCC'); // palette colour of the .xls file
  assert.equal(s.cells.get(K(1, 0)).s.u, true);
  assert.equal(s.cells.get(K(2, 0)).s.rot, 45);
  assert.equal(s.cells.get(K(3, 0)).s.h, 'right'); assert.equal(s.cells.get(K(3, 0)).s.st, true);
  assert.equal(s.cells.get(K(4, 1)).s.fill, '#FF0000');
  assert.deepEqual(b.props.defaultFont, { font: 'Calibri', size: 11 });
});
