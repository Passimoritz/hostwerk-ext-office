'use strict';
/* Templates (1.3): package conversion document ↔ template, the template store (own folder, imported folders, default). */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const F = require('./server-fixtures');
const T = require('../server/templates');
const { entries, inflate } = require('../server/zippatch');

const part = (buf, name) => { const e = entries(buf).find((x) => x.name === name); return e ? inflate(e).toString('utf8') : null; };

test('docx → dotx → docx switches only the main content type', () => {
  const doc = F.docx('Hallo Vorlage');
  assert.strictEqual(T.packageType(doc), T.MAIN.docx);
  const tpl = T.convertPackage(doc, 'dotx');
  assert.strictEqual(T.packageType(tpl), T.MAIN.dotx);
  assert.match(part(tpl, 'word/document.xml'), /Hallo Vorlage/);
  assert.deepStrictEqual(entries(tpl).map((e) => e.name), entries(doc).map((e) => e.name));
  const back = T.convertPackage(tpl, 'docx', { fresh: true });
  assert.strictEqual(T.packageType(back), T.MAIN.docx);
  // already the target type: the bytes stay as they are
  assert.strictEqual(T.convertPackage(doc, 'docx'), doc);
});

test('xlsx → xltx and pptx → potx', () => {
  assert.strictEqual(T.packageType(T.convertPackage(F.xlsx([['a', 1]]), 'xltx')), T.MAIN.xltx);
  assert.strictEqual(T.packageType(T.convertPackage(F.pptx([{ text: ['T'] }]), 'potx')), T.MAIN.potx);
});

test('odt → ott: mimetype first and stored, manifest root entry follows', () => {
  const odt = F.odt(['Absatz']);
  const ott = T.convertPackage(odt, 'ott');
  const list = entries(ott);
  assert.strictEqual(list[0].name, 'mimetype');
  assert.strictEqual(list[0].method, 0);
  assert.strictEqual(T.packageType(ott), 'application/vnd.oasis.opendocument.text-template');
  const mf = part(ott, 'META-INF/manifest.xml');
  if (mf && /full-path="\/"/.test(mf)) assert.match(mf, /media-type="application\/vnd\.oasis\.opendocument\.text-template"/);
  assert.strictEqual(T.packageType(T.convertPackage(ott, 'odt', { fresh: true })), 'application/vnd.oasis.opendocument.text');
});

test('maps: template ↔ document, macro templates keep macros', () => {
  assert.strictEqual(T.TEMPLATE_OF.docx, 'dotx');
  assert.strictEqual(T.TEMPLATE_OF.docm, 'dotm');
  assert.strictEqual(T.DOC_OF.dotm, 'docm');
  assert.strictEqual(T.DOC_OF.otp, 'odp');
  assert.ok(T.isTemplateExt('xltx') && !T.isTemplateExt('xlsx'));
  assert.ok(T.usable('Brief.dotx', false) && !T.usable('Brief.docx', false) && T.usable('Brief.docx', true) && !T.usable('a.txt', true) && !T.usable('~$Brief.dotx', true));
});

test('store: own folder, imported folder (read-only), default, instantiate, rename', () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'of-tpl-'));
  const lib = path.join(tmp, 'lib'); const imp = path.join(tmp, 'imp', 'Briefe');
  fs.mkdirSync(lib, { recursive: true }); fs.mkdirSync(imp, { recursive: true });
  fs.writeFileSync(path.join(imp, 'Normal.dotx'), T.convertPackage(F.docx('Normal'), 'dotx'));
  fs.writeFileSync(path.join(imp, 'readme.txt'), 'x');
  let settings = { libraryDir: lib, templates: { dir: '', saveTo: 'folder', folders: [path.join(tmp, 'imp')], defaults: { docs: '', sheets: '', slides: '' } } };
  const store = new T.TemplateStore({ settings: () => settings, saveSettings: (p) => { settings = { ...settings, templates: { ...settings.templates, ...p.templates, defaults: { ...settings.templates.defaults, ...((p.templates || {}).defaults || {}) } } }; }, lang: () => 'de' });
  assert.strictEqual(path.basename(store.ownDir()), 'Vorlagen');
  const src = path.join(lib, 'Bericht.xlsx');
  fs.writeFileSync(src, F.xlsx([['x', 2]]));
  const { data, ext } = store.asTemplate(src);
  assert.strictEqual(ext, 'xltx');
  const full = store.addOwn(data, ext, 'Bericht', 'Firma');
  assert.ok(full.startsWith(path.join(lib, 'Vorlagen', 'Firma')));
  const list = store.list();
  const own = list.find((t) => t.source === 'own');
  const imported = list.find((t) => t.source === 'folder');
  assert.ok(own && own.kind === 'sheets' && own.category === 'Firma' && own.path === 'Vorlagen/Firma/Bericht.xltx');
  assert.ok(imported && imported.readOnly && imported.category === 'Briefe' && imported.kind === 'docs');
  assert.strictEqual(list.length, 2); // readme.txt is no template
  const made = store.instantiate(imported.full);
  assert.strictEqual(made.ext, 'docx');
  assert.strictEqual(T.packageType(made.data), T.MAIN.docx);
  settings.templates.defaults.docs = imported.id; store.invalidate();
  assert.strictEqual(store.defaultFor('docs').id, imported.id);
  assert.throws(() => store.rename(imported.id, 'X'), /read-only/);
  const renamed = store.rename(own.id, 'Monatsbericht');
  assert.strictEqual(renamed.name, 'Monatsbericht');
  assert.ok(store.inTemplateFolder(renamed.full) && !store.inTemplateFolder(src));
  assert.throws(() => store.asTemplate(path.join(lib, 'x.csv')), /template form/);
  fs.rmSync(tmp, { recursive: true, force: true });
});
