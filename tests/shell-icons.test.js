// File type families of the coloured icons (ui/src/lib/fileTypes.js) – node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import { familyOf, iconInfo, FAMILIES, KIND_COLOR } from '../ui/src/lib/fileTypes.js';
import { FORMATS } from '../ui/src/lib/formats.js';

test('app families: documents blue, spreadsheets green, presentations orange, PDF red', () => {
  for (const x of ['a.docx', 'a.DOC', 'a.odt', 'a.rtf', 'b.dotx', 'c.docm']) assert.equal(familyOf(x), 'word', x);
  for (const x of ['a.xlsx', 'a.xls', 'a.ods', 'a.csv', 'a.xlsm']) assert.equal(familyOf(x), 'sheet', x);
  for (const x of ['a.pptx', 'a.ppt', 'a.odp', 'a.ppsx']) assert.equal(familyOf(x), 'slides', x);
  assert.equal(familyOf('scan.PDF'), 'pdf');
  assert.equal(KIND_COLOR.docs, FAMILIES.word.color);
  assert.equal(KIND_COLOR.sheets, FAMILIES.sheet.color);
  assert.equal(KIND_COLOR.slides, FAMILIES.slides.color);
  assert.equal(KIND_COLOR.pdf, FAMILIES.pdf.color);
});

test('other families', () => {
  assert.equal(familyOf('m.eml'), 'mail');
  assert.equal(familyOf('m.msg'), 'mail');
  assert.equal(familyOf('c.vcf'), 'contact');
  assert.equal(familyOf('c.ics'), 'calendar');
  assert.equal(familyOf('b.epub'), 'ebook');
  assert.equal(familyOf('d.vsdx'), 'diagram');
  assert.equal(familyOf('p.png'), 'image');
  assert.equal(familyOf('z.zip'), 'archive');
  assert.equal(familyOf('g.gsheet'), 'google');
  assert.equal(familyOf('k.key'), 'iwork');
  assert.equal(familyOf('n.numbers'), 'iwork');
  assert.equal(familyOf('readme'), 'file');
  assert.equal(familyOf('x.unknownext'), 'file');
  assert.equal(familyOf({ name: 'Ordner', dir: true }), 'folder');
  assert.equal(familyOf('docx'), 'word', 'a bare extension works too');
});

test('Google and iWork files: glyph by what they are, colour of their app', () => {
  const g = iconInfo('Budget.gsheet');
  assert.equal(g.glyph, 'grid');
  assert.equal(g.badge, 'cloud');
  assert.equal(iconInfo('Talk.gslides').glyph, 'screen');
  assert.equal(iconInfo('Text.pages').glyph, 'lines');
  assert.equal(iconInfo('Zahlen.numbers').glyph, 'grid');
  assert.equal(iconInfo('Vortrag.key').glyph, 'screen');
  assert.notEqual(iconInfo('Vortrag.key').color, iconInfo('Zahlen.numbers').color);
});

test('every format Office knows has a coloured family (no grey office file)', () => {
  const grey = Object.keys(FORMATS).filter((ext) => FORMATS[ext].kind !== 'viewer' && familyOf(`x.${ext}`) === 'file');
  assert.deepEqual(grey, []);
  for (const f of Object.values(FAMILIES)) assert.match(f.color, /^#[0-9a-f]{6}$/);
});

test('iconInfo carries the extension for the label under large icons', () => {
  assert.equal(iconInfo('Report 2026.DOCX').ext, 'docx');
  assert.equal(iconInfo({ name: 'x', dir: true }).ext, '');
});
