'use strict';
/* Office 1.1 server modules – pure logic (no Hostwerk needed). */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const crypto = require('crypto');
const { makeZip, readZip } = require('../server/zip');
const { readCfb } = require('../server/cfb');
const { writeCfb } = require('../server/cfbwrite');
const ZP = require('../server/zippatch');
const E = require('../server/encrypt');
const S = require('../server/sheet');
const { cleanFile, cleanJpeg, cleanPng } = require('../server/clean');
const { Components, untar } = require('../server/components');
const C = require('../server/collab');
const { Remote } = require('../server/remote');
const F = require('../server/forms');
const RU = require('../server/rules');
const { DataSources, redact } = require('../server/datasources');
const { DbSheets, parseFilter } = require('../server/dbsheets');
const { nextRun, Thumbs, LibBackup } = require('../server/extras');
const SE = require('../server/semantic');
const FX = require('./server-fixtures');

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'office-v11-'));

/* ---------------- CFB writer ---------------- */
test('CFB writer: mini and big streams, nested storages, readable by the reader', () => {
  const big = crypto.randomBytes(70000); const small = Buffer.from('hello mini stream');
  const files = { EncryptionInfo: small, EncryptedPackage: big, '\u0006DataSpaces/Version': Buffer.alloc(76, 1), '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary': Buffer.alloc(200, 2) };
  const m = readCfb(writeCfb(files));
  for (const [k, v] of Object.entries(files)) assert.ok(m.get(k) && m.get(k).equals(v), k);
});
test('CFB writer: many FAT sectors (> 7 MB stream)', () => {
  const big = crypto.randomBytes(8 * 1048576);
  assert.ok(readCfb(writeCfb({ A: big, B: Buffer.from('x') })).get('A').equals(big));
});

/* ---------------- encryption ---------------- */
test('OOXML agile encryption: round trip, wrong password, already encrypted', () => {
  const z = FX.docx('Geheim\nZweite Zeile');
  const enc = E.encryptOoxml(z, 'pässwört', { spinCount: 1000 });
  assert.ok(E.isEncryptedOoxml(enc));
  const info = readCfb(enc).get('EncryptionInfo');
  assert.deepStrictEqual([...info.subarray(0, 4)], [4, 0, 4, 0]);
  assert.match(info.subarray(8).toString(), /spinCount="1000"[^>]*hashAlgorithm="SHA512"/);
  assert.ok(E.decryptOoxml(enc, 'pässwört').equals(z));
  assert.throws(() => E.decryptOoxml(enc, 'falsch'), (e) => e.code === 'EPASSWORD');
  assert.throws(() => E.encryptOoxml(enc, 'x'), (e) => e.code === 'EALREADY');
  // integrity: a flipped byte in the package is detected
  const m = readCfb(enc); const pkg = Buffer.from(m.get('EncryptedPackage')); pkg[100] ^= 1;
  const files = {}; for (const [k, v] of m) files[k] = v; files.EncryptedPackage = pkg;
  assert.throws(() => E.decryptOoxml(writeCfb(files), 'pässwört'), (e) => e.code === 'EFORMAT');
});
test('ODF encryption: AES-256-CBC + PBKDF2, mimetype first and stored, thumbnail dropped', () => {
  const man = '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/></manifest:manifest>';
  const content = `<office:document-content>${'<text:p>Hallo</text:p>'.repeat(300)}</office:document-content>`;
  const z = makeZip([{ name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', store: true }, { name: 'content.xml', data: content }, { name: 'styles.xml', data: '<s/>' }, { name: 'Thumbnails/thumbnail.png', data: 'png' }, { name: 'META-INF/manifest.xml', data: man }]);
  const enc = E.encryptOdf(z, 'Geheim', { iterations: 1000 });
  const ze = ZP.entries(enc);
  assert.strictEqual(ze[0].name, 'mimetype'); assert.strictEqual(ze[0].method, 0);
  assert.ok(!ze.some((e) => e.name.startsWith('Thumbnails/')));
  const m = readZip(enc).get('META-INF/manifest.xml').read().toString();
  assert.match(m, /full-path="content.xml"[^>]*manifest:size="\d+"><manifest:encryption-data/);
  assert.match(m, /aes256-cbc/); assert.match(m, /iteration-count="1000"/);
  assert.ok(!readZip(enc).get('content.xml').read().toString().includes('Hallo'));
  assert.ok(E.isEncryptedOdf(enc));
  const dec = E.decryptOdf(enc, 'Geheim');
  assert.strictEqual(readZip(dec).get('content.xml').read().toString(), content);
  assert.throws(() => E.decryptOdf(enc, 'nein'), (e) => e.code === 'EPASSWORD');
});

/* ---------------- zip patch ---------------- */
test('zip patch keeps untouched entries byte-identical, replaces, removes and adds', () => {
  const z = makeZip({ 'a.xml': 'A'.repeat(500), 'b.bin': crypto.randomBytes(300), 'c.txt': 'C' });
  const before = ZP.entries(z);
  const out = ZP.patchZip(z, { 'c.txt': 'neu', 'a.xml': null, 'd.txt': 'D' });
  const after = ZP.entries(out);
  assert.deepStrictEqual(after.map((e) => e.name), ['b.bin', 'c.txt', 'd.txt']);
  assert.ok(after[0].raw.equals(before[1].raw));
  assert.strictEqual(readZip(out).get('c.txt').read().toString(), 'neu');
});

/* ---------------- sheets ---------------- */
test('sheet: read xlsx cells (shared, inline, numbers) and append a row – rest unchanged', () => {
  const x = FX.xlsx([['Name', 'Betrag'], ['Miete', 800]]);
  const book = S.readBook(x, 'xlsx');
  assert.strictEqual(book.sheets[0].name, 'Budget');
  assert.strictEqual(S.cellValue(book, 'Budget', 'B2'), 800);
  assert.strictEqual(S.cellValue(book, '', 'A3'), 'Inline note');
  const out = S.appendRow(x, 'xlsx', { sheet: 'Budget', values: ['Strom <&>', 61.5, true, null, '=HYPERLINK("x")'] });
  assert.strictEqual(out.row, 4);
  const b2 = S.readBook(out.data, 'xlsx');
  assert.strictEqual(S.cellValue(b2, 'Budget', 'A4'), 'Strom <&>');
  assert.strictEqual(S.cellValue(b2, 'Budget', 'B4'), 61.5);
  assert.strictEqual(S.cellValue(b2, 'Budget', 'C4'), true);
  assert.strictEqual(S.cellValue(b2, 'Budget', 'E4'), '=HYPERLINK("x")'); // a string, never a formula
  const before = new Map(ZP.entries(x).map((e) => [e.name, e.raw])); const after = ZP.entries(out.data);
  for (const e of after) if (e.name !== 'xl/worksheets/sheet1.xml') assert.ok(e.raw.equals(before.get(e.name)), e.name);
  assert.deepStrictEqual(S.rangeValues(b2, 'Budget', 'B1:B4'), ['Betrag', 800, 61.5]);
  assert.throws(() => S.appendRow(x, 'xlsx', { sheet: 'Nope', values: ['a'] }), /does not exist/);
});
test('sheet: ods read + append keeps the trailing repeated rows; csv append quotes and defuses formulas', () => {
  const content = '<office:document-content xmlns:office="o" xmlns:table="t" xmlns:text="x"><office:body><office:spreadsheet><table:table table:name="Daten"><table:table-row><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row><table:table-row table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="1024"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>';
  const ods = makeZip([{ name: 'mimetype', data: 'application/vnd.oasis.opendocument.spreadsheet', store: true }, { name: 'content.xml', data: content }]);
  const out = S.appendRow(ods, 'ods', { sheet: 'Daten', values: ['B', 3] });
  const c2 = readZip(out.data).get('content.xml').read().toString();
  assert.match(c2, /number-rows-repeated="1048574"/);
  const book = S.readBook(out.data, 'ods');
  assert.strictEqual(S.cellValue(book, 'Daten', 'A2'), 'B'); assert.strictEqual(S.cellValue(book, 'Daten', 'B2'), 3);
  assert.strictEqual(ZP.entries(out.data)[0].name, 'mimetype');
  const csv = Buffer.from('Name;Wert\r\nx;1\r\n');
  const c = S.appendRow(csv, 'csv', { values: ['=1+1', 'a;b', 'zeile\n2'] });
  assert.strictEqual(c.data.toString(), 'Name;Wert\r\nx;1\r\n\'=1+1;"a;b";"zeile\n2"\r\n');
});

/* ---------------- clean ---------------- */
function dirtyDocx() {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  return makeZip({
    '[Content_Types].xml': '<Types><Override PartName="/word/document.xml" ContentType="x"/><Override PartName="/word/comments.xml" ContentType="c"/><Override PartName="/docProps/core.xml" ContentType="core"/><Override PartName="/docProps/custom.xml" ContentType="cu"/><Override PartName="/docProps/app.xml" ContentType="app"/></Types>',
    '_rels/.rels': '<Relationships><Relationship Id="rId1" Type="officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="custom-properties" Target="docProps/custom.xml"/><Relationship Id="rId3" Type="thumbnail" Target="docProps/thumbnail.jpeg"/></Relationships>',
    'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId5" Type="comments" Target="comments.xml"/></Relationships>',
    'word/document.xml': `<w:document ${W}><w:body><w:p w:rsidR="00AB12CD"><w:r><w:t>Behalten</w:t></w:r><w:ins w:id="1" w:author="Anna"><w:r><w:t> neu</w:t></w:r></w:ins><w:del w:id="2" w:author="Anna"><w:r><w:delText>weg</w:delText></w:r></w:del><w:commentRangeStart w:id="0"/><w:r><w:rPr><w:vanish/></w:rPr><w:t>versteckt</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>`,
    'word/comments.xml': `<w:comments ${W}><w:comment w:id="0" w:author="Anna"><w:p><w:r><w:t>Kommentar</w:t></w:r></w:p></w:comment></w:comments>`,
    'word/settings.xml': `<w:settings ${W}><w:attachedTemplate r:id="rId1"/><w:rsids><w:rsidRoot w:val="00AB12CD"/></w:rsids></w:settings>`,
    'docProps/core.xml': '<cp:coreProperties xmlns:cp="c" xmlns:dc="d"><dc:title>Titel</dc:title><dc:creator>Anna Muster</dc:creator><cp:lastModifiedBy>Bernd</cp:lastModifiedBy></cp:coreProperties>',
    'docProps/app.xml': '<Properties><Company>Muster GmbH</Company><Template>C:\\Users\\anna\\Vorlage.dotm</Template><TotalTime>42</TotalTime></Properties>',
    'docProps/custom.xml': '<Properties><property name="Kunde"/></Properties>',
    'docProps/thumbnail.jpeg': Buffer.from([0xff, 0xd8, 0xff, 0xd9]),
  });
}
test('clean docx: author, company, comments, tracked changes accepted, hidden text, rsid, custom props, thumbnail', () => {
  const { data, report } = cleanFile(dirtyDocx(), 'docx');
  const z = readZip(data);
  const doc = z.get('word/document.xml').read().toString();
  assert.match(doc, /Behalten/); assert.match(doc, / neu/);
  assert.doesNotMatch(doc, /weg|versteckt|w:ins|w:del|comment|rsid/);
  assert.ok(!z.has('word/comments.xml') && !z.has('docProps/custom.xml') && !z.has('docProps/thumbnail.jpeg'));
  assert.doesNotMatch(z.get('[Content_Types].xml').read().toString(), /comments|custom/);
  assert.doesNotMatch(z.get('word/_rels/document.xml.rels').read().toString(), /comments/);
  assert.doesNotMatch(z.get('_rels/.rels').read().toString(), /custom|thumbnail/);
  const core = z.get('docProps/core.xml').read().toString();
  assert.match(core, /Titel/); assert.doesNotMatch(core, /Anna|Bernd/);
  const app = z.get('docProps/app.xml').read().toString();
  assert.doesNotMatch(app, /Muster GmbH|anna/); assert.match(app, /<Template>Vorlage.dotm<\/Template>/);
  assert.doesNotMatch(z.get('word/settings.xml').read().toString(), /rsids|attachedTemplate/);
  assert.ok(report.author && report.comments && report.revisions >= 2 && report.hidden === 1 && report.custom && report.thumbnail && report.rsid);
});
test('clean images: JPEG EXIF/XMP and PNG text chunks removed, image data kept', () => {
  const seg = (m, body) => { const b = Buffer.alloc(4); b[0] = 0xff; b[1] = m; b.writeUInt16BE(body.length + 2, 2); return Buffer.concat([b, body]); };
  const jpg = Buffer.concat([Buffer.from([0xff, 0xd8]), seg(0xe0, Buffer.from('JFIF\0')), seg(0xe1, Buffer.from('Exif\0\0GPS-Daten')), Buffer.from([0xff, 0xda, 0, 2, 1, 2, 3, 0xff, 0xd9])]);
  const c = cleanJpeg(jpg);
  assert.ok(!c.includes(Buffer.from('GPS'))); assert.ok(c.includes(Buffer.from('JFIF'))); assert.ok(c.subarray(-5).equals(Buffer.from([1, 2, 3, 0xff, 0xd9])));
  const chunk = (t, d) => { const l = Buffer.alloc(4); l.writeUInt32BE(d.length); return Buffer.concat([l, Buffer.from(t), d, Buffer.alloc(4)]); };
  const png = Buffer.concat([Buffer.from('89504e470d0a1a0a', 'hex'), chunk('IHDR', Buffer.alloc(13)), chunk('tEXt', Buffer.from('Author\0Anna')), chunk('IDAT', Buffer.alloc(5)), chunk('IEND', Buffer.alloc(0))]);
  const p = cleanPng(png);
  assert.ok(!p.includes(Buffer.from('Anna'))); assert.ok(p.includes(Buffer.from('IDAT')));
});
test('clean pdf: Info values blanked in place – same length, xref offsets stay valid', () => {
  const pdf = Buffer.from('%PDF-1.4\n1 0 obj\n<< /Title (Bericht) /Author (Anna Muster) /Producer (Word) >>\nendobj\nxref\n0 1\ntrailer\n<< /Info 1 0 R >>\n%%EOF');
  const { data, report } = cleanFile(pdf, 'pdf');
  assert.strictEqual(data.length, pdf.length);
  assert.doesNotMatch(data.toString('latin1'), /Anna|Word/); assert.match(data.toString('latin1'), /Bericht/);
  assert.ok(report.author);
});
test('clean odt: creator, annotations, tracked changes, thumbnail', () => {
  const z = makeZip([{ name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', store: true },
    { name: 'meta.xml', data: '<office:meta><meta:initial-creator>Anna</meta:initial-creator><dc:creator>Bernd</dc:creator><dc:title>T</dc:title><meta:user-defined meta:name="Kunde">X</meta:user-defined></office:meta>' },
    { name: 'content.xml', data: '<office:text><text:tracked-changes><text:changed-region text:id="c1"><text:deletion><text:p>alt</text:p></text:deletion></text:changed-region></text:tracked-changes><text:p>Hallo<text:change text:change-id="c1"/><office:annotation><dc:creator>Anna</dc:creator><text:p>Notiz</text:p></office:annotation></text:p></office:text>' },
    { name: 'Thumbnails/thumbnail.png', data: 'png' },
    { name: 'META-INF/manifest.xml', data: '<manifest:manifest><manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/></manifest:manifest>' }]);
  const { data } = cleanFile(z, 'odt');
  const r = readZip(data);
  assert.doesNotMatch(r.get('meta.xml').read().toString(), /Anna|Bernd|Kunde/);
  const c = r.get('content.xml').read().toString();
  assert.match(c, /Hallo/); assert.doesNotMatch(c, /alt|Notiz|Anna|text:change/);
  assert.ok(!r.has('Thumbnails/thumbnail.png'));
  assert.strictEqual(ZP.entries(data)[0].name, 'mimetype');
});

/* ---------------- components ---------------- */
function tgz(files) {
  const parts = [];
  for (const [name, data] of Object.entries(files)) {
    const h = Buffer.alloc(512); h.write(name, 0); h.write('0000644\0', 100); h.write('0000000\0', 108); h.write('0000000\0', 116);
    h.write(`${data.length.toString(8).padStart(11, '0')}\0`, 124); h.write('00000000000\0', 136); h.write('        ', 148); h[156] = 48; h.write('ustar\0', 257); h.write('00', 263);
    let sum = 0; for (const b of h) sum += b; h.write(`${sum.toString(8).padStart(6, '0')}\0 `, 148);
    parts.push(h, data, Buffer.alloc((512 - (data.length % 512)) % 512));
  }
  parts.push(Buffer.alloc(1024));
  return zlib.gzipSync(Buffer.concat(parts));
}
test('components: pinned hash checked, tar unpacked, only listed files served; wrong hash installs nothing', async () => {
  const pkg = tgz({ 'package/dist/lib.min.js': Buffer.from('console.log(1)'), 'package/data/a.bcmap': Buffer.from([1, 2, 3]), 'package/other.txt': Buffer.from('x') });
  assert.strictEqual(untar(zlib.gunzipSync(pkg)).get('package/dist/lib.min.js').toString(), 'console.log(1)');
  const file = Buffer.from('traineddata');
  const sha = (b) => crypto.createHash('sha256').update(b).digest('hex');
  const registry = { demo: { label: 'Demo', version: '1', license: 'MIT', sources: [{ url: 'https://example.test/p.tgz', sha256: sha(pkg), size: pkg.length, tar: { 'package/dist/lib.min.js': 'lib.min.js' }, tarDirs: { 'package/data/': 'data/' } }, { url: 'https://example.test/x', sha256: sha(file), size: file.length, file: 'lang/deu.traineddata' }] } };
  const fetchImpl = async (u) => ({ ok: true, status: 200, arrayBuffer: async () => (u.endsWith('.tgz') ? pkg : file) });
  const dir = tmp();
  const c = new Components({ dir, fetchImpl, registry });
  assert.strictEqual(c.view('demo').state, 'missing');
  const v = await c.install('demo');
  assert.strictEqual(v.state, 'ready'); assert.strictEqual(v.files.length, 3); assert.strictEqual(v.fileCount, 3);
  assert.strictEqual(c.file('demo', 'lib.min.js').type, 'text/javascript; charset=utf-8');
  assert.ok(c.file('demo', 'data/a.bcmap')); assert.strictEqual(c.file('demo', 'other.txt'), null); assert.strictEqual(c.file('demo', '../installed.json'), null);
  c.remove('demo');
  const bad = new Components({ dir, registry: { demo: { ...registry.demo, sources: [{ ...registry.demo.sources[0], sha256: '0'.repeat(64) }] } }, fetchImpl });
  await assert.rejects(bad.install('demo'), (e) => e.code === 'EHASH');
  assert.strictEqual(bad.view('demo').state, 'error'); assert.ok(!fs.existsSync(path.join(dir, 'demo')));
  assert.throws(() => c.def('nope'), /Unknown component/);
});

/* ---------------- co-editing relay with real Yjs (dev dependency) ---------------- */
test('collab relay: two Yjs docs converge, late joiner gets the state, awareness removal, leader hand-over', (t) => {
  let Y, sync, aw, enc, dec;
  try { Y = require('yjs'); sync = require('y-protocols/sync'); aw = require('y-protocols/awareness'); enc = require('lib0/encoding'); dec = require('lib0/decoding'); } catch { t.skip('yjs not installed'); return; }
  const relay = new C.Collab();
  const client = (name) => {
    const doc = new Y.Doc(); const a = new aw.Awareness(doc); const inbox = []; const ctl = [];
    const peer = { send: (b) => inbox.push(Buffer.from(b)), close: () => {} };
    const conn = relay.connect('Bericht.docx', peer, { user: name });
    const handle = () => {
      while (inbox.length) {
        const buf = inbox.shift(); const d = dec.createDecoder(new Uint8Array(buf)); const type = dec.readVarUint(d);
        if (type === 0) { const e = enc.createEncoder(); enc.writeVarUint(e, 0); sync.readSyncMessage(d, e, doc, 'remote'); if (enc.length(e) > 1) conn.message(Buffer.from(enc.toUint8Array(e))); } else if (type === 1) aw.applyAwarenessUpdate(a, dec.readVarUint8Array(d), 'remote'); else if (type === 100) ctl.push(JSON.parse(dec.readVarString(d)));
      }
    };
    doc.on('update', (u, origin) => { if (origin === 'remote') return; const e = enc.createEncoder(); enc.writeVarUint(e, 0); sync.writeUpdate(e, u); conn.message(Buffer.from(enc.toUint8Array(e))); });
    a.on('update', ({ added, updated, removed }, origin) => { if (origin === 'remote') return; const e = enc.createEncoder(); enc.writeVarUint(e, 1); enc.writeVarUint8Array(e, aw.encodeAwarenessUpdate(a, [...added, ...updated, ...removed])); conn.message(Buffer.from(enc.toUint8Array(e))); });
    return { doc, a, conn, handle, ctl };
  };
  const A = client('Anna'); A.handle();
  A.doc.getText('t').insert(0, 'Hallo');
  A.a.setLocalStateField('user', { name: 'Anna' });
  const B = client('Bernd');
  for (let i = 0; i < 4; i++) { A.handle(); B.handle(); }
  assert.strictEqual(B.doc.getText('t').toString(), 'Hallo');
  B.doc.getText('t').insert(5, ' Welt');
  for (let i = 0; i < 4; i++) { A.handle(); B.handle(); }
  assert.strictEqual(A.doc.getText('t').toString(), 'Hallo Welt');
  assert.ok(A.ctl[0].leader && !B.ctl[0].leader);
  assert.ok([...B.a.getStates().values()].some((s) => s.user && s.user.name === 'Anna'));
  const late = client('Carla'); for (let i = 0; i < 4; i++) { late.handle(); A.handle(); B.handle(); }
  assert.strictEqual(late.doc.getText('t').toString(), 'Hallo Welt');
  A.conn.close(); B.handle();
  assert.ok(B.ctl.some((m) => m.type === 'leader' && m.leader));
  assert.ok(![...B.a.getStates().values()].some((s) => s.user && s.user.name === 'Anna'));
  B.conn.close(); late.conn.close();
  for (const c of [A, B, late]) c.a.destroy(); // y-protocols awareness keeps an interval
  relay.sweep(Date.now() + 10 * 60000);
  assert.strictEqual(relay.rooms.size, 0);
});
test('collab relay: limits and garbage', () => {
  const relay = new C.Collab({ clients: 1 });
  const closed = [];
  const c1 = relay.connect('x', { send: () => {}, close: (c) => closed.push(c) });
  assert.throws(() => relay.connect('x', { send: () => {}, close: () => {} }), /Too many people/);
  c1.message(Buffer.from([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff]));
  assert.deepStrictEqual(closed, [1003]);
});

/* ---------------- remote control ---------------- */
test('remote control: presenter state reaches the phone, commands reach the presenter, rate limit', () => {
  let now = 1000; const r = new Remote({ now: () => now });
  const s = r.create('Anna');
  const toPresenter = []; const toPhone = [];
  assert.throws(() => r.command(s.token, { cmd: 'next' }), /not connected/);
  const pr = r.attach(s.token, { send: (m) => toPresenter.push(m), close: () => {} });
  const off = r.phone(s.token, { send: (m) => toPhone.push(m), end: () => {} });
  pr.message(JSON.stringify({ type: 'state', slide: 2, total: 9, notes: 'Hallo <b>', timer: { running: true, elapsed: 5000 } }));
  assert.strictEqual(toPhone[toPhone.length - 1].slide, 2); assert.strictEqual(toPhone[toPhone.length - 1].notes, 'Hallo <b>');
  r.command(s.token, { cmd: 'goto', n: '4' }); r.command(s.token, { cmd: 'laser', x: 3, y: -1 });
  assert.deepStrictEqual(toPresenter.filter((m) => m.cmd), [{ type: 'goto', cmd: 'goto', n: 4, value: 4 }, { type: 'laser', cmd: 'laser', on: true, x: 1, y: 0, value: { x: 1, y: 0, on: true } }]);
  pr.message(JSON.stringify({ type: 'notes', value: 'Nur Notiz' })); pr.message(JSON.stringify({ type: 'slide', value: { index: 3, total: 9 } }));
  assert.strictEqual(toPhone[toPhone.length - 1].notes, 'Nur Notiz'); assert.strictEqual(toPhone[toPhone.length - 1].slide, 3);
  r.command(s.token, { type: 'prev' }); assert.strictEqual(toPresenter[toPresenter.length - 1].type, 'prev');
  assert.throws(() => r.command(s.token, { cmd: 'rm -rf' }), /Unknown command/);
  for (let i = 0; i < 17; i++) r.command(s.token, { cmd: 'next' });
  assert.throws(() => r.command(s.token, { cmd: 'next' }), /Too many/);
  off();
  now += 13 * 3600000; assert.strictEqual(r.get(s.token), null);
  assert.strictEqual(r.get('../../etc'), null);
});

/* ---------------- forms ---------------- */
test('forms: definition and answers are validated, spam stamp', () => {
  const f = F.validateForm({ title: 'Anmeldung', columns: [{ label: 'Name', required: true }, { label: 'Alter', type: 'number', min: 0, max: 120 }, { label: 'E-Mail', type: 'email' }, { label: 'Essen', type: 'choice', options: 'Fleisch\nVegan' }, { label: 'AGB', type: 'checkbox' }] });
  assert.strictEqual(f.columns[3].options.length, 2);
  assert.deepStrictEqual(F.checkAnswer(f, { f0: ' Anna ', f1: '41,5', f2: 'a@b.de', f3: 'Vegan', f4: 'on' }), ['Anna', 41.5, 'a@b.de', 'Vegan', true]);
  assert.throws(() => F.checkAnswer(f, { f1: 'x', f2: 'nope', f3: 'Pizza' }), (e) => e.fields.f0 === 'required' && e.fields.f1 === 'number' && e.fields.f2 === 'email' && e.fields.f3 === 'choice');
  assert.throws(() => F.validateForm({ columns: [] }), /1 to 60/);
  const secret = crypto.randomBytes(32); const tok = F.newToken();
  const st = F.stamp(secret, tok, 1000);
  assert.ok(!F.stampOk(secret, tok, st, 1500)); assert.ok(F.stampOk(secret, tok, st, 4000)); assert.ok(!F.stampOk(secret, F.newToken(), st, 4000)); assert.ok(!F.stampOk(secret, tok, 'x.y', 4000));
  const html = F.formPage('de', { ...f, title: '<script>' }, { base: '/api/office/pub/form/x', stamp: st });
  assert.doesNotMatch(html, /<script>/); assert.match(html, /name="website"/);
});

/* ---------------- rules ---------------- */
test('rules: validation, conditions, fire only on the change to "met"', async () => {
  assert.throws(() => RU.validateRules([{ when: { cell: 'Z0', op: '>' } }]), /invalid cell/);
  assert.throws(() => RU.validateRules([{ when: { cell: 'A1', op: 'eval' } }]), /invalid condition/);
  const list = RU.validateRules([{ id: 'r1', when: { sheet: 'Budget', cell: '$B$2', op: '>', value: '1000' }, then: { notify: { channel: 'hostwerk', text: '{cell} = {value}' } } }, { id: 'r2', when: { sheet: 'Budget', range: 'B1:B9', op: '>=', value: '900', agg: 'sum' } }]);
  assert.strictEqual(list[0].when.cell, 'B2');
  let value = 800; let mtime = 1;
  const read = async () => ({ buf: FX.xlsx([['x', 'y'], ['Miete', value]]), ext: 'xlsx', mtime });
  const sent = [];
  const dir = tmp();
  const rules = new RU.Rules({ file: path.join(dir, 'rules.json'), read, notify: async (r, vars) => sent.push([r.id, RU.fill(r.then.notify.text, vars)]) });
  rules.set('b.xlsx', { path: 'b.xlsx' }, list);
  await new Promise((ok) => setTimeout(ok, 20));
  assert.deepStrictEqual(sent, []);
  value = 1200; mtime = 2; await rules.check('b.xlsx');
  assert.deepStrictEqual(sent, [['r1', 'B2 = 1200'], ['r2', '']]);
  mtime = 3; await rules.check('b.xlsx'); assert.strictEqual(sent.length, 2); // still met → silent
  value = 10; mtime = 4; await rules.check('b.xlsx'); value = 5000; mtime = 5; await rules.check('b.xlsx');
  assert.strictEqual(sent.length, 4); // both rules met again
  assert.strictEqual(rules.get('b.xlsx')[0].last.value, 5000);
  rules.rekey('b.xlsx', 'Ordner/b.xlsx'); assert.ok(rules.get('Ordner/b.xlsx').length === 2);
});

/* ---------------- data sources ---------------- */
test('data sources: services table from /overview, players, redaction', async () => {
  const api = {
    '/overview': { web: [{ id: 'w1', name: 'Apache', httpPort: 80, proc: { state: 'running', stats: { cpu: 1.234, memory: 52428800 }, startedAt: 0 } }], db: [{ id: 'd1', name: 'MariaDB', port: 3306, proc: { state: 'stopped' } }], apps: [], minecraft: [{ id: 'mc1', name: 'Survival', maxPlayers: 20, proc: { state: 'running', extra: { players: ['Steve', 'Alex'], playerCount: 2, tps: 19.9 } } }] },
    '/live/uptime': { checks: [{ name: 'Shop', url: 'https://admin:geheim@shop.test/', ok: true, uptime24: 0.995, ms: 120 }] },
  };
  const ds = new DataSources({ get: async (p) => api[p] });
  const s = await ds.query('services');
  assert.deepStrictEqual(s.rows[0].slice(0, 6), ['Apache', 'web', 'running', 80, 1.2, 50]);
  assert.strictEqual((await ds.query('playerCount', {})).value, 2);
  assert.deepStrictEqual((await ds.query('players', { server: 'Survival' })).rows, [['Steve', 'Survival'], ['Alex', 'Survival']]);
  assert.strictEqual((await ds.query('service', { service: 'MariaDB' })).value, 'stopped');
  const u = await ds.query('uptime');
  assert.strictEqual(u.rows[0][1], 'https://admin:••••••@shop.test/'); assert.strictEqual(u.rows[0][3], 99.5);
  await assert.rejects(ds.query('../settings'), /Unknown data source/);
  assert.strictEqual(redact('sk-abcdefghijklmnop'), '••••••');
});

/* ---------------- database tables (node:sqlite as the driver) ---------------- */
test('database sheets: paging/filter through the driver, writes in one transaction with rollback', async (t) => {
  let DatabaseSync; try { ({ DatabaseSync } = require('node:sqlite')); } catch { t.skip('node:sqlite not available'); return; }
  const db = new DatabaseSync(':memory:');
  db.exec('CREATE TABLE kunden (id INTEGER PRIMARY KEY, name TEXT, ort TEXT); INSERT INTO kunden (name, ort) VALUES (\'Anna\', \'Köln\'), (\'Bernd\', \'Bonn\'), (\'Carla\', \'Köln\'); CREATE TABLE log (msg TEXT);');
  const cols = (table) => db.prepare(`PRAGMA table_info("${table.replace(/"/g, '""')}")`).all().map((c) => ({ name: c.name, type: c.type, key: c.pk ? 'PRI' : '' }));
  const driver = {
    db,
    databases: async () => [{ name: 'main' }],
    tables: async () => db.prepare("SELECT name, type FROM sqlite_master WHERE type IN ('table','view')").all().map((r) => ({ name: r.name, view: r.type === 'view' })),
    structure: async (d, table) => ({ columns: cols(table) }),
    tableRows: async (d, table, o) => { const c = cols(table); const f = (o.filters || [])[0]; const w = f ? ` WHERE "${f.col}" = ?` : ''; const p = f ? [f.value] : []; return { columns: c, pk: c.filter((x) => x.key).map((x) => x.name), rows: db.prepare(`SELECT * FROM "${table}"${w} LIMIT ${o.size} OFFSET ${(o.page - 1) * o.size}`).all(...p).map((r) => c.map((x) => r[x.name])), total: db.prepare(`SELECT COUNT(*) n FROM "${table}"${w}`).get(...p).n }; },
  };
  const ctxDb = { connections: () => [{ id: 's1', name: 'Test', type: 'sqlite', kind: 'sql', running: true }], driver: () => ({ d: driver }), guard: () => {} };
  const ds = new DbSheets(ctxDb);
  assert.deepStrictEqual((await ds.tables('s1', 'main')).map((x) => x.name).sort(), ['kunden', 'log']);
  const page = await ds.rows('s1', 'main', 'kunden', { offset: 1, limit: 1 });
  assert.deepStrictEqual(page.rows, [[2, 'Bernd', 'Bonn']]); assert.deepStrictEqual(page.primaryKey, ['id']); assert.strictEqual(page.total, 3);
  const f = await ds.rows('s1', 'main', 'kunden', { filter: JSON.stringify([{ col: 'ort', op: '=', value: 'Köln' }]) });
  assert.strictEqual(f.total, 2);
  assert.ok((await ds.rows('s1', 'main', 'log', {})).readOnly);
  await ds.write('s1', 'main', 'kunden', [{ op: 'update', key: { id: 1 }, values: { ort: 'Bonn"; DROP TABLE kunden; --' } }, { op: 'insert', values: { name: 'Dora', ort: 'Aachen' } }, { op: 'delete', key: { id: 3 } }]);
  assert.deepStrictEqual(db.prepare('SELECT name, ort FROM kunden ORDER BY id').all().map((r) => [r.name, r.ort]), [['Anna', 'Bonn"; DROP TABLE kunden; --'], ['Bernd', 'Bonn'], ['Dora', 'Aachen']]);
  // the second change fails (row gone) → the first one is rolled back
  await assert.rejects(ds.write('s1', 'main', 'kunden', [{ op: 'update', key: { id: 1 }, values: { ort: 'X' } }, { op: 'update', key: { id: 99 }, values: { ort: 'Y' } }]), /did not find its row/);
  assert.strictEqual(db.prepare('SELECT ort FROM kunden WHERE id = 1').get().ort, 'Bonn"; DROP TABLE kunden; --');
  await assert.rejects(ds.write('s1', 'main', 'kunden', [{ op: 'update', key: { id: 1 }, values: { 'ort"=1;--': 'x' } }]), /Unknown column/);
  await assert.rejects(ds.write('s1', 'main', 'log', [{ op: 'insert', values: { msg: 'x' } }]), /no primary key/);
  await assert.rejects(ds.rows('s1', 'main', 'kunden; DROP', {}), /Table not found/);
  assert.throws(() => parseFilter('[{"col":"a","op":"; DROP"}]'), /Invalid filter/);
});

/* ---------------- schedules, thumbs, library backup, semantic ---------------- */
test('report schedules: next run hourly/daily/weekly/monthly', () => {
  const from = new Date(2026, 10, 20, 9, 30).getTime(); // Friday
  assert.strictEqual(new Date(nextRun({ every: 'daily', time: '08:00' }, from)).getDate(), 21);
  assert.strictEqual(new Date(nextRun({ every: 'daily', time: '10:00' }, from)).getHours(), 10);
  const w = new Date(nextRun({ every: 'weekly', weekday: 1, time: '07:15' }, from)); assert.strictEqual(w.getDay(), 1); assert.strictEqual(w.getDate(), 23);
  const m = new Date(nextRun({ every: 'monthly', day: 1, time: '06:00' }, from)); assert.strictEqual(m.getMonth(), 11); assert.strictEqual(m.getDate(), 1);
  assert.strictEqual(new Date(nextRun({ every: 'hourly', time: '00:45' }, from)).getMinutes(), 45);
});
test('thumbnails cache per mtime; library backup and restore keep versions', async () => {
  const dir = tmp();
  const th = new Thumbs(path.join(dir, 't'));
  th.put('a.docx', 10, Buffer.from('png')); assert.strictEqual(String(th.get('a.docx', 10).data), 'png'); assert.strictEqual(th.get('a.docx', 11), null); assert.ok(th.has('a.docx', 10) && !th.has('b.docx', 10)); th.flush();
  const lib = path.join(dir, 'lib'); fs.mkdirSync(path.join(lib, 'Ordner'), { recursive: true });
  fs.writeFileSync(path.join(lib, 'a.txt'), 'eins'); fs.writeFileSync(path.join(lib, 'Ordner', 'b.docx'), 'zwei'); fs.writeFileSync(path.join(lib, '~$lock.docx'), 'junk');
  const written = [];
  const lb = new LibBackup({ dir: path.join(dir, 'bk'), base: () => lib, log: () => {}, write: (rel, data) => { written.push(rel); fs.writeFileSync(path.join(lib, ...rel.split('/')), data); } });
  const c = await lb.create({ force: true });
  assert.strictEqual(c.files, 2);
  fs.writeFileSync(path.join(lib, 'a.txt'), 'geändert'); fs.rmSync(path.join(lib, 'Ordner', 'b.docx'));
  const r = await lb.restore(c.id, { paths: ['a.txt'] });
  assert.deepStrictEqual([r.restored, written], [1, ['a.txt']]); assert.strictEqual(fs.readFileSync(path.join(lib, 'a.txt'), 'utf8'), 'eins');
  await lb.restore(c.id); assert.ok(fs.existsSync(path.join(lib, 'Ordner', 'b.docx')));
  await assert.rejects(lb.restore('../x.zip'), /not found/);
});
test('search by meaning: int8 vectors, cosine, AI search words over the full-text index', async () => {
  const v = [0.1, -0.5, 0.25, 0.9]; const back = SE.dequant(SE.quantize(v));
  assert.ok(SE.cosine(v, back) > 0.999);
  const index = { data: { files: {} }, get: () => null, search: (t) => (t === 'rechnung' ? [{ rel: 'Rechnung 2026.docx', score: 3, snippet: '…' }] : t === 'invoice' ? [{ rel: 'Invoice.xlsx', score: 1, snippet: '' }] : []) };
  const ai = { api: async () => ({ json: async () => ({ models: [{ name: 'llama3' }] }) }), chats: { oneShot: async () => 'Rechnung, Invoice, Beleg' } };
  const s = new SE.Semantic({ file: path.join(tmp(), 's.json'), ai: () => ai, index });
  const r = await s.search('Wo sind meine Belege für Strom?');
  assert.strictEqual(r.mode, 'ai-terms'); assert.deepStrictEqual(r.hits.map((h) => h.rel), ['Rechnung 2026.docx', 'Invoice.xlsx']);
  const plain = await new SE.Semantic({ file: path.join(tmp(), 's.json'), ai: () => null, index }).search('rechnung');
  assert.strictEqual(plain.mode, 'fulltext'); assert.strictEqual(plain.hits.length, 1);
});
