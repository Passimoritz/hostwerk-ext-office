'use strict';
/* Text, facts and thumbnails from tiny generated office files (built with the own zip writer / CFB builder). */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const X = require('../server/extract');
const fx = require('./server-fixtures');
const { Library } = require('../server/library');
const { Meta } = require('../server/meta');
const { Versions } = require('../server/versions');
const { Trash } = require('../server/trash');
const { Index } = require('../server/search');
const SET = require('../server/settings');

test('docx: paragraphs, metadata', async () => {
  const buf = fx.docx('Quarterly budget\nZweite Zeile äöü');
  assert.strictEqual(await X.text(buf, 'docx'), 'Quarterly budget\nZweite Zeile äöü');
  const i = await X.info(buf, 'docx');
  assert.strictEqual(i.title, 'Test title');
  assert.strictEqual(i.author, 'Anna');
  assert.strictEqual(i.pages, 1);
  assert.strictEqual(i.words, 4);
  assert.strictEqual(i.created, Date.parse('2024-01-02T03:04:05Z'));
  assert.strictEqual(await X.text(buf, 'docm'), 'Quarterly budget\nZweite Zeile äöü');
});

test('xlsx: shared strings, numbers, inline strings, sheet names', async () => {
  const buf = fx.xlsx([['Item', 'Price'], ['Bread', 2.5]]);
  const t = await X.text(buf, 'xlsx');
  assert.strictEqual(t, 'Budget\nItem\tPrice\nBread\t2.5\nInline note');
  assert.strictEqual((await X.info(buf, 'xlsx')).sheets, 1);
});

test('pptx: slides in order with notes', async () => {
  const buf = fx.pptx([{ text: ['Title one', 'Point A'] }, { text: ['Second slide'], notes: 'Speaker note' }]);
  assert.strictEqual(await X.text(buf, 'pptx'), '--- 1 ---\nTitle one\nPoint A\n\n--- 2 ---\nSecond slide\n\nSpeaker note');
  assert.strictEqual((await X.info(buf, 'pptx')).slides, 2);
});

test('odt: content.xml with spaces, meta.xml, thumbnail', async () => {
  const buf = fx.odt(['Heading', 'Some  notes'], { title: 'Notes', author: 'Ben', pages: 2, words: 3 });
  assert.strictEqual(await X.text(buf, 'odt'), 'Heading\nSome  notes');
  const i = await X.info(buf, 'odt');
  assert.deepStrictEqual([i.title, i.author, i.pages, i.words], ['Notes', 'Ben', 2, 3]);
  const th = X.thumb(buf, 'odt');
  assert.strictEqual(th.type, 'image/png');
  assert.ok(X.hasThumb(buf, 'odt'));
  assert.strictEqual(X.thumb(fx.docx('x'), 'docx'), null);
});

test('legacy binary: Word 97 piece table and PowerPoint 97 text atoms', async () => {
  assert.strictEqual(await X.text(fx.doc('Hello Word\rSecond paragraph äö'), 'doc'), 'Hello Word\nSecond paragraph äö');
  assert.strictEqual(await X.text(fx.ppt(['First slide', 'Zweite Folie']), 'ppt'), '--- 1 ---\nFirst slide\n\n--- 2 ---\nZweite Folie');
  assert.strictEqual((await X.info(fx.ppt(['a', 'b', 'c']), 'ppt')).slides, 3);
  // field codes: instruction dropped, result kept
  assert.strictEqual(require('../server/cfb').cleanWord('Page \u0013 PAGE \u00142\u0015 of 3\r'), 'Page 2 of 3');
});

test('rtf, html, txt (BOM, UTF-16), csv', async () => {
  assert.strictEqual(await X.text(Buffer.from('{\\rtf1\\ansi{\\fonttbl{\\f0 Arial;}}\\f0 Hallo \\\'e4\\par Welt\\u8364?}'), 'rtf'), 'Hallo ä\nWelt€');
  assert.strictEqual(await X.text(Buffer.from('<html><head><title>x</title><style>p{}</style></head><body><p>A &amp; B</p><script>alert(1)</script><p>C</p></body></html>'), 'html'), 'A & B\nC');
  assert.strictEqual(await X.text(Buffer.from('﻿äöü', 'utf8'), 'txt'), 'äöü');
  assert.strictEqual(await X.text(Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('Grüße', 'utf16le')]), 'txt'), 'Grüße');
  assert.strictEqual(await X.text(Buffer.from('a;b\n1;2'), 'csv'), 'a;b\n1;2');
  assert.strictEqual(await X.text(Buffer.from('nothing'), 'xyz'), '');
  assert.strictEqual(await X.text(Buffer.from('not a zip'), 'docx'), '');
});

test('snippets around the hit', () => {
  const t = `${'x '.repeat(100)}The budget for 2026 is final. ${'y '.repeat(100)}`;
  const s = X.snippet(t, 'BUDGET', 20);
  assert.ok(s.startsWith('…') && s.endsWith('…') && /budget for 2026/.test(s));
  assert.strictEqual(X.snippet('short text', 'missing'), 'short text');
  assert.strictEqual(X.pdfDate("D:20240131120000+01'00'"), Date.UTC(2024, 0, 31, 11, 0, 0));
});

test('library: create, write with conflict + versions, rename keeps metadata, trash/restore, index search', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'office-lib-'));
  const lib0 = path.join(root, 'office');
  let settings = SET.normalize({}, lib0);
  const meta = new Meta(path.join(root, 'meta.json'));
  const versions = new Versions(path.join(root, 'versions'));
  const trash = new Trash(path.join(root, 'trash'));
  const index = new Index({ file: path.join(root, 'index.json'), base: () => settings.libraryDir, libs: () => ({}) });
  const lib = new Library({ settings: () => settings, files: null, meta, versions, trash, index, sharedKeys: () => new Set() });
  lib.mkdir('', 'Work');
  const e = lib.create('Work', 'Plan.docx', fx.docx('Budget planning'));
  assert.strictEqual(e.path, 'Work/Plan.docx');
  assert.strictEqual(lib.create('Work', 'Plan.docx', fx.docx('x')).name, 'Plan (2).docx');
  const r0 = lib.read({ path: 'Work/Plan.docx' });
  assert.throws(() => lib.write({ path: 'Work/Plan.docx' }, fx.docx('B'), { mtime: r0.mtime - 5000 }), (x) => x.status === 409 && x.extra.mtime === r0.mtime);
  const w = lib.write({ path: 'Work/Plan.docx' }, fx.docx('Budget planning v2'), { mtime: r0.mtime });
  assert.ok(w.ok && w.version);
  assert.strictEqual(versions.list('Work/Plan.docx').length, 1);
  meta.set('Work/Plan.docx', { favorite: true });
  lib.rename('Work', 'Projects');
  assert.ok(meta.get('Projects/Plan.docx').favorite, 'metadata follows the folder rename');
  assert.strictEqual(versions.list('Projects/Plan.docx').length, 1, 'versions follow too');
  assert.throws(() => lib.rename('Projects/Plan.docx', 'Plan (2).docx'), /already exists/);
  const restored = lib.restoreVersion({ path: 'Projects/Plan.docx' }, versions.list('Projects/Plan.docx')[0].id);
  assert.strictEqual(await X.text(fs.readFileSync(path.join(lib0, 'Projects', 'Plan.docx')), 'docx'), 'Budget planning');
  assert.ok(restored && versions.list('Projects/Plan.docx').length === 2);
  const [t] = lib.toTrash(['Projects/Plan.docx']);
  assert.ok(!fs.existsSync(path.join(lib0, 'Projects', 'Plan.docx')));
  assert.strictEqual(meta.get('Projects/Plan.docx'), null);
  lib.restore([t.id]);
  assert.ok(meta.get('Projects/Plan.docx').favorite, 'metadata comes back with the file');
  await index.reindex();
  const hits = index.search('budget');
  assert.deepStrictEqual(hits.map((h) => h.rel).sort(), ['Projects/Plan.docx']);
  const list = Library.sortEntries(lib.listDir('Projects'), 'name');
  assert.deepStrictEqual(list.map((x) => x.name), ['Plan (2).docx', 'Plan.docx']);
  assert.throws(() => lib.loc({ path: '../escape.docx' }));
  assert.throws(() => lib.move(['Projects'], 'Projects'), /into itself/);
  settings = { ...settings };
  index.stop(); meta.flush();
});
