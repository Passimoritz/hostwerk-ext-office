// Word tables, paragraph and character properties: docx round trip (reader ↔ writer) and the width/style helpers.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { defaultModel, para, table } from '../ui/src/editors/docs/io/pm.js';
import { readDocx, fieldKind } from '../ui/src/editors/docs/io/docx-read.js';
import { writeDocx } from '../ui/src/editors/docs/io/docx-write.js';
import { builtinTableStyles, tableWidthCss, tableLook, tint, borderCss } from '../ui/src/editors/docs/io/tables.js';

const T = (text, attrs) => ({ type: 'text', text, marks: attrs ? [{ type: 'textStyle', attrs }] : undefined });
async function roundTrip(blocks) {
  const m = defaultModel({});
  m.doc = { type: 'doc', content: blocks };
  const bytes = await writeDocx(m, { ext: 'docx' });
  return { model: await readDocx(bytes), zip: await JSZip.loadAsync(bytes) };
}

test('tables: style, look, width, indent, borders, margins, rows and cells survive docx', async () => {
  const t = table([['A', 'B'], ['1', '2'], ['3', '4']], { header: true });
  t.attrs = { styleId: 'GridTable4-Accent1', tstyle: builtinTableStyles()['GridTable4-Accent1'].spec, look: { firstRow: true, lastRow: true, firstColumn: false, lastColumn: false, noHBand: false, noVBand: true }, width: { type: 'pct', w: 5000 }, layout: 'fixed', indent: 360, borders: { top: 'double:4:C00000:0', insideV: 'nil' }, cellMargins: { top: 57, left: 113, bottom: 57, right: 113 } };
  t.content[1].attrs = { height: 600, heightRule: 'exact', cantSplit: true };
  Object.assign(t.content[2].content[0].attrs, { vAlign: 'center', textDir: 'btLr', noWrap: true, borders: { left: 'dashed:8:00B050:0' }, margins: { left: 300 } });
  const { model, zip } = await roundTrip([t, para('x')]);
  const r = model.doc.content.find((n) => n.type === 'table');
  assert.equal(r.attrs.styleId, 'GridTable4-Accent1');
  assert.deepEqual(r.attrs.width, { type: 'pct', w: 5000 });
  assert.equal(r.attrs.layout, 'fixed');
  assert.equal(r.attrs.indent, 360);
  assert.equal(r.attrs.look.lastRow, true);
  assert.equal(r.attrs.borders.top, 'double:4:C00000:0');
  assert.equal(r.attrs.borders.insideV, 'nil');
  assert.deepEqual(r.attrs.cellMargins, { top: 57, left: 113, bottom: 57, right: 113 });
  assert.equal(r.attrs.tstyle.cond.firstRow.fill, '4472C4', 'the style written into styles.xml is read back');
  assert.deepEqual(r.content[1].attrs.height, 600);
  assert.equal(r.content[1].attrs.heightRule, 'exact');
  assert.equal(r.content[1].attrs.cantSplit, true);
  const c = r.content[2].content[0].attrs;
  assert.equal(c.vAlign, 'center'); assert.equal(c.textDir, 'btLr'); assert.equal(c.noWrap, true);
  assert.equal(c.borders.left, 'dashed:8:00B050:0'); assert.equal(c.margins.left, 300);
  assert.equal(r.content[0].content[0].type, 'tableHeader');
  const styles = await zip.file('word/styles.xml').async('string');
  assert.match(styles, /w:styleId="GridTable4-Accent1"/);
  const doc = await zip.file('word/document.xml').async('string');
  // schema order inside tblPr / tcPr
  assert.match(doc, /<w:tblPr><w:tblStyle w:val="GridTable4-Accent1"\/><w:tblW w:w="5000" w:type="pct"\/><w:tblInd[^>]*\/><w:tblBorders>[\s\S]*?<\/w:tblBorders><w:tblLayout w:type="fixed"\/><w:tblCellMar>/);
  assert.match(doc, /<w:tcBorders>[\s\S]*?<\/w:tcBorders><w:noWrap\/><w:tcMar>[\s\S]*?<\/w:tcMar><w:textDirection w:val="btLr"\/><w:vAlign w:val="center"\/>/);
});

test('paragraph and character properties survive docx', async () => {
  const p = { type: 'paragraph', attrs: { keepNext: true, keepLines: true, rightIndent: 36, lineRule: 'exact', lineExact: 18, pborders: { top: 'single:4:auto:4' }, pshading: '#FFF2CC', widow: false, tabs: [{ pos: 300, align: 'right', leader: 'dot' }], dropCap: { type: 'drop', lines: 3 } }, content: [T('Text\t1', { letterSpacing: 2, scale: 150, position: 3, caps: 'small', dstrike: true, hidden: true, lang: 'en-GB' })] };
  const { model } = await roundTrip([p]);
  const a = model.doc.content[0].attrs;
  for (const k of ['keepNext', 'keepLines', 'rightIndent', 'lineRule', 'lineExact', 'pshading', 'widow']) assert.deepEqual(a[k], p.attrs[k], k);
  assert.equal(a.pborders.top, 'single:4:auto:4');
  assert.deepEqual(a.tabs, [{ pos: 300, align: 'right', leader: 'dot' }]);
  assert.deepEqual(a.dropCap, { type: 'drop', lines: 3 });
  const ts = model.doc.content[0].content[0].marks.find((m) => m.type === 'textStyle').attrs;
  for (const k of ['letterSpacing', 'scale', 'position', 'caps', 'dstrike', 'hidden', 'lang']) assert.equal(ts[k], p.content[0].marks[0].attrs[k], k);
});

test('formula fields are recognised', () => {
  assert.equal(fieldKind('=SUM(ABOVE)'), 'formula');
  assert.equal(fieldKind(' = average(left) \\# "0,00"'), 'formula');
  assert.equal(fieldKind('PAGE'), 'page');
});

test('widths: fixed tables keep their grid (may run into the margin), autofit shrinks to the text width, pct takes its share', () => {
  const cols = [100, 200, 300];
  const fixed = tableWidthCss({ layout: 'fixed', width: { type: 'dxa', w: 12500 } }, cols);
  assert.equal(fixed.table.width, `${Math.round((12500 / 20) * (4 / 3))}px`);
  assert.equal(fixed.cols[0], '100px');
  const auto = tableWidthCss({}, cols);
  assert.equal(auto.table.width, 'min(600px, 100%)');
  assert.equal(auto.cols[2], '50%');
  const pct = tableWidthCss({ width: { type: 'pct', w: 2500 } }, cols);
  assert.equal(pct.table.width, '50%');
  assert.equal(tableWidthCss({ indent: 300 }, cols).table['margin-left'], '20px');
});

test('table style look and colours', () => {
  // Word's own tints (lighter 80 % / 40 %), within one step of rounding
  const near = (a, b) => [0, 2, 4].every((k) => Math.abs(parseInt(a.slice(k, k + 2), 16) - parseInt(b.slice(k, k + 2), 16)) <= 2);
  assert.ok(near(tint('4472C4', 0.2), 'D9E2F3'), tint('4472C4', 0.2));
  assert.ok(near(tint('4472C4', 0.6), '8EAADB'), tint('4472C4', 0.6));
  assert.equal(borderCss('nil'), 'none');
  assert.equal(borderCss('single:8:FF0000:0'), '1.3px solid #FF0000');
  const l = tableLook({ tstyle: builtinTableStyles()['GridTable4-Accent1'].spec, look: { firstRow: true, noHBand: false } });
  assert.ok(l.cls.includes('of-fr-bg') && l.cls.includes('of-bh-bg') && l.cls.includes('of-lk-fr'));
  assert.equal(l.vars['--fr-bg'], '#4472C4');
  assert.equal(Object.keys(builtinTableStyles()).length, 41);
});
