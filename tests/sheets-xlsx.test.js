import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Book } from '../ui/src/editors/sheets/model.js';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { readXlsx, writeXlsx } from '../ui/src/editors/sheets/io/xlsx.js';
import { readSheetJS, writeSheetJS } from '../ui/src/editors/sheets/io/sheetjs.js';

const loc = makeLocale({ lang: 'en' });
const K = (r, c) => r * 16384 + c;

function sample() {
  const b = Book.blank({ loc });
  const s = b.sheets[0];
  b.run('fill', () => {
    b.input(s, 0, 0, 'Item'); b.input(s, 0, 1, 'Price');
    b.input(s, 1, 0, 'Apple'); b.input(s, 1, 1, '1.5');
    b.input(s, 2, 0, 'Pear'); b.input(s, 2, 1, '2.25');
    b.input(s, 3, 0, 'Total'); b.input(s, 3, 1, '=SUM(B2:B3)');
    b.input(s, 4, 0, '11/18/2026');
    b.input(s, 5, 0, '=XLOOKUP("Pear",A2:A3,B2:B3)');
    b.input(s, 6, 0, '=SEQUENCE(3)');
  });
  b.format(s, [{ r1: 0, c1: 0, r2: 0, c2: 1 }], { b: true, fill: '#FFEE00', color: '#C00000' });
  b.format(s, [{ r1: 1, c1: 1, r2: 3, c2: 1 }], { nf: '#,##0.00 "€"' });
  b.borders(s, { r1: 0, c1: 0, r2: 3, c2: 1 }, 'all');
  b.merge(s, { r1: 10, c1: 0, r2: 10, c2: 2 }, { center: true });
  b.setSizes(s, 'col', [0], 120);
  b.setSheetProps(s, { freeze: { rows: 1, cols: 0 } });
  b.run('n', () => { b.patchCell(s, 1, 0, { note: 'Fresh from the market' }); b.patchCell(s, 2, 0, { link: 'https://example.com/pear' }); });
  b.setSheetProps(s, { dv: [{ id: 'd', ranges: [{ r1: 1, c1: 2, r2: 5, c2: 2 }], type: 'list', f1: '"Open,Done"', allowBlank: true, error: { style: 'stop', show: true } }] });
  b.setSheetProps(s, { cf: [{ id: 'c', ranges: [{ r1: 1, c1: 1, r2: 2, c2: 1 }], type: 'cell', op: 'gt', f1: '2', style: { fill: '#C6EFCE', color: '#006100' } }] });
  b.setSheetProps(s, { charts: [{ id: 'ch1', type: 'column', title: 'Prices', legend: 'bottom', range: 'Sheet1!$A$1:$B$3', byRows: false, anchor: { r: 1, c: 4, dx: 0, dy: 0, w: 400, h: 240 } }] });
  const s2 = b.addSheet('My Data');
  b.run('x', () => { b.input(s2, 0, 0, "=Sheet1!B4*2"); });
  return b;
}
const getValues = (b) => (ref) => {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(ref);
  if (!m) return null;
  const sh = b.sheetByName((m[1] || m[2]).replace(/''/g, "'"));
  const col = (x) => x.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const r1 = +m[4] - 1; const c1 = col(m[3]); const r2 = m[6] ? +m[6] - 1 : r1; const c2 = m[5] ? col(m[5]) : c1;
  const values = []; const texts = [];
  for (let r = r1; r <= r2; r++) { const a = []; const t = []; for (let c = c1; c <= c2; c++) { a.push(b.value(sh, r, c)); t.push(b.text(sh, r, c)); } values.push(a); texts.push(t); }
  return { values, texts };
};

test('xlsx write → exceljs read', async () => {
  const b = sample();
  const bytes = await writeXlsx(b, { getValues: getValues(b) });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(bytes));
  const ws = wb.getWorksheet('Sheet1');
  assert.equal(ws.getCell('A2').value, 'Apple');
  assert.equal(ws.getCell('B2').value, 1.5);
  assert.equal(ws.getCell('B4').value.formula, 'SUM(B2:B3)');
  assert.equal(ws.getCell('B4').value.result, 3.75);
  assert.equal(ws.getCell('A1').font.bold, true);
  assert.equal(ws.getCell('A1').fill.fgColor.argb, 'FFFFEE00');
  assert.equal(ws.getCell('B2').numFmt, '#,##0.00 "€"');
  assert.equal(ws.getCell('A1').border.top.style, 'thin');
  assert.ok(ws.getCell('B11').isMerged);
  assert.equal(Math.round(ws.getColumn(1).width), 16);
  assert.equal(ws.views[0].state, 'frozen');
  assert.equal(ws.views[0].ySplit, 1);
  assert.equal(ws.getCell('A2').note.texts ? ws.getCell('A2').note.texts.map((t) => t.text).join('') : ws.getCell('A2').note, 'Fresh from the market');
  assert.equal(ws.getCell('A3').value.hyperlink, 'https://example.com/pear');
  assert.equal(ws.getCell('C2').dataValidation.type, 'list');
  assert.equal(wb.getWorksheet('My Data').getCell('A1').value.result, 7.5);
  const zip = await JSZip.loadAsync(bytes);
  assert.ok(zip.file('xl/charts/chart1.xml'), 'chart part written');
  assert.match(await zip.file('xl/worksheets/sheet1.xml').async('string'), /<f t="array" ref="A7:A9">_xlfn.SEQUENCE\(3\)<\/f>/);
});

test('xlsx round trip through our reader', async () => {
  const b = sample();
  const bytes = await writeXlsx(b, { getValues: getValues(b) });
  const r = await readXlsx(bytes, { loc });
  const s = r.sheets[0];
  assert.equal(r.sheets.length, 2);
  assert.equal(r.sheets[1].name, 'My Data');
  assert.equal(r.value(s, 3, 1), 3.75);
  assert.equal(s.cells.get(K(3, 1)).f, 'SUM(B2:B3)');
  assert.equal(r.value(s, 5, 0), 2.25);
  assert.equal(r.value(s, 8, 0), 3, 'spill recomputed');
  assert.equal(s.cells.get(K(0, 0)).s.b, true);
  assert.equal(s.cells.get(K(0, 0)).s.fill, '#FFEE00');
  assert.equal(s.cells.get(K(1, 1)).s.nf, '#,##0.00 "€"');
  assert.equal(r.text(s, 4, 0), '11/18/2026');
  assert.deepEqual(s.merges[0], { r1: 10, c1: 0, r2: 10, c2: 2 });
  assert.equal(s.colW.get(0), 120);
  assert.deepEqual(s.freeze, { rows: 1, cols: 0 });
  assert.equal(s.cells.get(K(1, 0)).note, 'Fresh from the market');
  assert.equal(s.cells.get(K(2, 0)).link, 'https://example.com/pear');
  assert.equal(s.dv[0].type, 'list'); assert.equal(s.dv[0].f1, '"Open,Done"');
  assert.equal(s.cf[0].type, 'cell'); assert.equal(s.cf[0].style.fill, '#C6EFCE');
  assert.equal(s.charts.length, 1); assert.equal(s.charts[0].type, 'column'); assert.equal(s.charts[0].title, 'Prices');
  assert.equal(r.value(r.sheets[1], 0, 0), 7.5);
});

test('lossless patching keeps unknown parts byte-identical (vba, customXml, untouched charts)', async () => {
  const b = sample();
  let bytes = await writeXlsx(b, { ext: 'xlsm', getValues: getValues(b) });
  // add parts we don't understand, like Excel would have them
  const zip = await JSZip.loadAsync(bytes);
  const vba = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4, 5]);
  zip.file('xl/vbaProject.bin', vba);
  let rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  rels = rels.replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
  zip.file('xl/_rels/workbook.xml.rels', rels);
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>');
  zip.file('[Content_Types].xml', ct);
  zip.file('customXml/item1.xml', '<root>keep me</root>');
  const chartBefore = await zip.file('xl/charts/chart1.xml').async('string');
  // an x14 extension in the sheet (e.g. sparklines) must survive
  let sx = await zip.file('xl/worksheets/sheet1.xml').async('string');
  sx = sx.replace('</worksheet>', '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups/></ext></extLst></worksheet>');
  zip.file('xl/worksheets/sheet1.xml', sx);
  bytes = await zip.generateAsync({ type: 'uint8array' });

  const r = await readXlsx(bytes, { loc, ext: 'xlsm' });
  assert.ok(r.fidelity.some((x) => /Macros/.test(x)));
  r.run('edit', () => r.input(r.sheets[0], 1, 1, '10'));
  const out = await writeXlsx(r, { ext: 'xlsm', getValues: getValues(r) });
  const z2 = await JSZip.loadAsync(out);
  assert.deepEqual(await z2.file('xl/vbaProject.bin').async('uint8array'), vba);
  assert.equal(await z2.file('customXml/item1.xml').async('string'), '<root>keep me</root>');
  assert.equal(await z2.file('xl/charts/chart1.xml').async('string'), chartBefore, 'unchanged chart untouched');
  assert.match(await z2.file('[Content_Types].xml').async('string'), /macroEnabled\.main\+xml/);
  assert.match(await z2.file('xl/worksheets/sheet1.xml').async('string'), /x14:sparklineGroups/);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(out));
  assert.equal(wb.getWorksheet('Sheet1').getCell('B4').value.result, 12.25);
});

test('reads a workbook written by exceljs (styles, shared formulas, merges)', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.getCell('A1').value = 'Name'; ws.getCell('A1').font = { bold: true, italic: true };
  ws.getCell('B1').value = 3; ws.getCell('B2').value = 4;
  ws.getCell('C1').value = { formula: 'B1*2', result: 6 };
  ws.getCell('C2').value = { sharedFormula: 'C1', result: 8 };
  ws.fillFormula ? 0 : 0;
  ws.mergeCells('A5:C6');
  ws.getColumn(2).width = 20;
  ws.getRow(3).height = 30;
  ws.getCell('B1').numFmt = '0.00%';
  ws.getCell('D1').value = new Date(Date.UTC(2024, 0, 15));
  const buf = await wb.xlsx.writeBuffer();
  const r = await readXlsx(new Uint8Array(buf), { loc });
  const s = r.sheets[0];
  assert.equal(s.name, 'Data');
  assert.equal(s.cells.get(K(0, 0)).s.b, true);
  assert.equal(s.cells.get(K(0, 0)).s.i, true);
  assert.equal(s.cells.get(K(1, 2)).f, 'B2*2');
  assert.equal(r.value(s, 1, 2), 8);
  assert.equal(r.text(s, 0, 1), '300.00%');
  assert.equal(s.merges.length, 1);
  assert.equal(s.colW.get(1), 145);
  assert.equal(s.rowH.get(2), 40);
  assert.equal(r.value(s, 0, 3), 45306);
});

test('other formats via SheetJS (ods, csv-like, xls) round trip values and formulas', async () => {
  const b = sample();
  for (const ext of ['ods', 'xls', 'fods']) {
    const bytes = writeSheetJS(b, ext);
    const r = readSheetJS(bytes, { loc, ext });
    const s = r.sheets[0];
    assert.equal(r.value(s, 1, 0), 'Apple', ext);
    assert.equal(r.value(s, 3, 1), 3.75, ext);
    if (ext !== 'xls') assert.equal(s.cells.get(K(3, 1)).f, 'SUM(B2:B3)', ext); // SheetJS CE writes BIFF8 without formulas
  }
});
