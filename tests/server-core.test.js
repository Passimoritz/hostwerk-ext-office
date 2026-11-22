'use strict';
/* Pure logic of the Office backend: paths, names, settings, zip, versions, trash, shares, ONLYOFFICE, engine, converter plan. */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const P = require('../server/paths');
const SET = require('../server/settings');
const Z = require('../server/zip');
const V = require('../server/versions');
const { Trash, expired: trashExpired } = require('../server/trash');
const S = require('../server/shares');
const W = require('../server/onlyoffice');
const { runArgs, APPS } = require('../server/servers');
const E = require('../server/engine');
const C = require('../server/convert');
const G = require('../server/google');
const { Meta, cleanPatch } = require('../server/meta');
const { snappy } = require('../server/iwork');

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'office-test-'));

test('path safety: normRel', () => {
  assert.strictEqual(P.normRel('/a//b/'), 'a/b');
  assert.strictEqual(P.normRel('a\\b\\c.docx'), 'a/b/c.docx');
  assert.strictEqual(P.normRel(''), '');
  assert.strictEqual(P.normRel('./a/./b'), 'a/b');
  for (const bad of ['../x', 'a/../../x', 'a/..', 'C:\\Windows', '\\\\server\\share', 'a\0b']) assert.throws(() => P.normRel(bad), bad);
  assert.throws(() => P.normRel('a/CON/b'), /reserved/);
  assert.throws(() => P.normRel('a/x?.docx'));
});

test('path safety: checkName and Windows reserved names', () => {
  for (const n of ['CON', 'con.txt', 'NUL', 'com1', 'LPT9.docx', 'aux.tar.gz']) assert.throws(() => P.checkName(n), /reserved/, n);
  for (const n of ['a:b', 'a/b', 'a\\b', 'a*b', 'a"b', 'a<b', 'a|b']) assert.throws(() => P.checkName(n), n);
  assert.throws(() => P.checkName('name.'), /dot or a space/);
  assert.throws(() => P.checkName(''), /enter a name/);
  assert.throws(() => P.checkName('x'.repeat(201)), /too long/);
  assert.strictEqual(P.checkName('  Bericht 2024.docx '), 'Bericht 2024.docx');
  assert.strictEqual(P.checkName('console.docx'), 'console.docx');
  assert.strictEqual(P.safeName('a/b:c?.docx'), 'a-b-c-.docx');
  assert.strictEqual(P.safeName('CON'), '_CON');
});

test('path safety: resolveIn never leaves the base', () => {
  const base = tmp();
  assert.strictEqual(P.resolveIn(base, 'a/b.docx'), path.join(base, 'a', 'b.docx'));
  assert.strictEqual(P.resolveIn(base, ''), path.resolve(base));
  assert.throws(() => P.resolveIn(base, '../outside.txt'));
  assert.ok(P.inside(base, path.join(base, 'x')));
  assert.ok(!P.inside(base, `${base}-other${path.sep}x`));
  // symlink/junction pointing outside
  const outside = tmp();
  try { fs.symlinkSync(outside, path.join(base, 'link'), 'junction'); } catch { return; }
  assert.throws(() => P.resolveIn(base, 'link/file.txt'), /outside/);
});

test('unique names', () => {
  const taken = new Set(['Report.docx', 'Report (2).docx', 'Folder', 'a.b.xlsx']);
  const ex = (n) => taken.has(n);
  assert.strictEqual(P.uniqueName('New.docx', ex), 'New.docx');
  assert.strictEqual(P.uniqueName('Report.docx', ex), 'Report (3).docx');
  assert.strictEqual(P.uniqueName('Report (2).docx', ex), 'Report (3).docx');
  assert.strictEqual(P.uniqueName('Folder', ex, true), 'Folder (2)');
  assert.strictEqual(P.uniqueName('a.b.xlsx', ex), 'a.b (2).xlsx');
  // a counter never pushes a name past the 200 characters checkName allows (the file would be unreachable)
  const long = `${'x'.repeat(195)}.docx`;
  const u = P.uniqueName(long, (n) => n === long);
  assert.ok(u.length <= 200 && u.endsWith(' (2).docx'));
  assert.strictEqual(P.checkName(u), u);
  assert.ok(P.isJunk('~$Report.docx') && P.isJunk('.~lock.x.odt#') && P.isJunk('Thumbs.db') && !P.isJunk('Report.docx'));
});

test('settings: defaults exactly as in the contract', () => {
  const d = SET.defaults('/lib');
  assert.strictEqual(d.libraryDir, '/lib');
  assert.deepStrictEqual(d.general, { autosave: true, autosaveSeconds: 20, newFormat: 'ms', author: '', spellcheck: true, language: 'auto', confirmDelete: true, trashDays: 30, keepVersions: 30, versionDays: 90, view: 'grid', sort: 'modified', openFromFiles: true, recentCount: 20, units: 'cm' });
  assert.deepStrictEqual(d.docs.margins, { top: 25, right: 25, bottom: 20, left: 25 });
  assert.strictEqual(d.converters.timeout, 120);
  assert.strictEqual(d.converters.autoEngine, 'ask');
  assert.deepStrictEqual(d.sharing, { enabled: true, defaultDays: 7, allowDownload: true });
});

test('settings: deep merge and validation', () => {
  const lib = path.resolve(os.tmpdir(), 'lib');
  const cur = SET.normalize({}, lib);
  let r = SET.merge(cur, { general: { autosaveSeconds: 60 }, docs: { margins: { top: 30 } }, unknown: 1 }, lib);
  assert.deepStrictEqual(r.errors, []);
  assert.strictEqual(r.settings.general.autosaveSeconds, 60);
  assert.strictEqual(r.settings.general.autosave, true);
  assert.deepStrictEqual(r.settings.docs.margins, { top: 30, right: 25, bottom: 20, left: 25 });
  assert.ok(!('unknown' in r.settings));
  r = SET.merge(cur, { general: { view: 'cards', trashDays: 0, autosave: 'yes' }, docs: { pageSize: 'Letter' }, integrations: { onlyoffice: { url: 'javascript:alert(1)' } } }, lib);
  assert.deepStrictEqual(r.errors.sort(), ['general.autosave', 'general.trashDays', 'general.view', 'integrations.onlyoffice.url']);
  assert.strictEqual(r.settings.general.view, 'grid');
  assert.strictEqual(r.settings.docs.pageSize, 'Letter');
  r = SET.merge(cur, { integrations: { onlyoffice: { url: 'https://docs.example.com/' } }, libraryDir: 'relative/path' }, lib);
  assert.strictEqual(r.settings.integrations.onlyoffice.url, 'https://docs.example.com');
  assert.deepStrictEqual(r.errors, ['libraryDir']);
  r = SET.merge(cur, { general: { recentCount: '30' }, pdf: { zoom: 125 } }, lib);
  assert.strictEqual(r.settings.general.recentCount, 30);
  assert.strictEqual(r.settings.pdf.zoom, '125');
  // stored garbage falls back to defaults
  assert.strictEqual(SET.normalize({ general: { view: 'x' } }, lib).general.view, 'grid');
  // ONLYOFFICE is the only full editor: a removed full editor as defaultOpen → onlyoffice, its settings key is dropped
  const old = SET.normalize({ integrations: { defaultOpen: 'removed-editor', 'removed-editor': { url: 'http://localhost:9980' }, onlyoffice: { url: 'http://localhost:9981' } } }, lib);
  assert.strictEqual(old.integrations.defaultOpen, 'onlyoffice');
  assert.deepStrictEqual(Object.keys(old.integrations).sort(), ['defaultOpen', 'google', 'hostUrl', 'onlyoffice', 'osOpenWith']);
  assert.strictEqual(old.integrations.osOpenWith, true, '"Open with" is on by default');
  assert.strictEqual(SET.merge(cur, { integrations: { osOpenWith: false } }, lib).settings.integrations.osOpenWith, false);
  for (const v of ['builtin', 'onlyoffice', 'desktop']) assert.strictEqual(SET.normalize({ integrations: { defaultOpen: v } }, lib).integrations.defaultOpen, v);
  assert.deepStrictEqual(SET.ENUMS['integrations.defaultOpen'], ['builtin', 'onlyoffice', 'desktop']);
  assert.deepStrictEqual(SET.merge(cur, { integrations: { defaultOpen: 'removed-editor' } }, lib).errors, ['integrations.defaultOpen'], 'PUT refuses it');
});

// yauzl is Hostwerk's own zip reader (host.lib('yauzl')) – the test uses it from a Hostwerk checkout next to the
// extension (or HOSTWERK_DIR); in CI without Hostwerk that part is skipped with a note (nothing is vendored).
function hostwerkYauzl() {
  const dirs = [process.env.HOSTWERK_DIR, path.join(__dirname, '..', '..', '..', 'hostwerk')].filter(Boolean);
  for (const d of dirs) { try { return require(require.resolve('yauzl', { paths: [d] })); } catch { /* next */ } }
  return null;
}

test('zip writer round trip (own reader and yauzl)', async (t) => {
  const big = Buffer.from('Hallo Welt äöü '.repeat(1000));
  const buf = Z.makeZip([{ name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', store: true }, { name: 'dir/', dir: true }, { name: 'dir/text.txt', data: big }, { name: 'Ümlaut.xml', data: '<a/>' }, { name: 'empty.bin', data: Buffer.alloc(0) }]);
  assert.ok(Z.isZip(buf));
  const z = Z.readZip(buf);
  assert.deepStrictEqual([...z.keys()], ['mimetype', 'dir/', 'dir/text.txt', 'Ümlaut.xml', 'empty.bin']);
  assert.ok(z.get('dir/text.txt').read().equals(big));
  assert.strictEqual(z.get('mimetype').method, 0);
  assert.strictEqual(z.get('dir/text.txt').method, 8);
  assert.strictEqual(z.get('empty.bin').read().length, 0);
  // the first bytes of an ODF must be the stored mimetype
  assert.strictEqual(buf.toString('latin1', 30, 38), 'mimetype');
  const yauzl = hostwerkYauzl();
  if (!yauzl) { t.diagnostic('yauzl cross-check skipped – no Hostwerk checkout found (set HOSTWERK_DIR to include it)'); return; }
  const names = await new Promise((resolve, reject) => {
    yauzl.fromBuffer(buf, { lazyEntries: true }, (err, zf) => {
      if (err) return reject(err);
      const out = {};
      zf.on('entry', (e) => { if (/\/$/.test(e.fileName)) { out[e.fileName] = null; zf.readEntry(); return; } zf.openReadStream(e, (er, st) => { if (er) return reject(er); const parts = []; st.on('data', (c) => parts.push(c)); st.on('end', () => { out[e.fileName] = Buffer.concat(parts); zf.readEntry(); }); }); });
      zf.on('end', () => resolve(out));
      zf.readEntry();
    });
  });
  assert.ok(names['dir/text.txt'].equals(big));
  assert.strictEqual(names['Ümlaut.xml'].toString(), '<a/>');
  assert.ok('dir/' in names);
});

test('zip: streaming writer gives the same bytes, crc32', () => {
  const parts = [];
  const w = new Z.ZipWriter((c) => { parts.push(c); return true; });
  const d = new Date(2024, 0, 2, 3, 4, 6);
  w.add('a.txt', 'abc', { mtime: d }); w.end();
  const mem = new Z.ZipWriter(); mem.add('a.txt', 'abc', { mtime: d });
  assert.ok(Buffer.concat(parts).equals(mem.end()));
  assert.strictEqual(Z.crc32(Buffer.from('123456789')), 0xcbf43926);
  assert.throws(() => { const x = new Z.ZipWriter(); x.add('a', '1'); x.add('a', '2'); }, /duplicate/);
  assert.strictEqual(Z.cleanName('../../etc/passwd'), 'etc/passwd');
  assert.throws(() => Z.readZip(Buffer.from('not a zip at all, really not')), /Not a zip/);
});

test('versions: pruning keeps named versions, drops old and surplus ones', () => {
  const now = Date.UTC(2026, 0, 31);
  const day = 86400000;
  const list = [
    { id: 'a', time: now - 1 * day }, { id: 'b', time: now - 2 * day }, { id: 'c', time: now - 3 * day },
    { id: 'old', time: now - 100 * day }, { id: 'named', time: now - 400 * day, label: 'Final' },
  ];
  assert.deepStrictEqual(V.pruneList(list, { keep: 2, days: 90, now }).sort(), ['c', 'old']);
  assert.deepStrictEqual(V.pruneList(list, { keep: 30, days: 90, now }), ['old']);
  assert.deepStrictEqual(V.pruneList(list, { keep: 0, days: 0, now }).sort(), ['a', 'b', 'c', 'old']);
});

test('versions store: dedupe, gap for automatic versions, labels, rekey', () => {
  const v = new V.Versions(tmp());
  const t0 = Date.now() - 3600000;
  const a = v.add('Docs/a.docx', Buffer.from('one'), { gapMs: 300000, time: t0 });
  assert.ok(a && a.auto);
  assert.strictEqual(v.add('Docs/a.docx', Buffer.from('one'), { gapMs: 300000 }), null, 'same content');
  assert.strictEqual(v.add('Docs/a.docx', Buffer.from('two'), { gapMs: 300000, time: t0 + 1000 }), null, 'too soon after an automatic one');
  const named = v.add('Docs/a.docx', Buffer.from('three'), { label: 'Draft 1' });
  assert.strictEqual(named.label, 'Draft 1');
  assert.strictEqual(v.list('Docs/a.docx').length, 2);
  v.rekey('Docs', 'Archive/Docs');
  assert.strictEqual(v.list('Docs/a.docx').length, 0);
  assert.strictEqual(v.list('Archive/Docs/a.docx').length, 2);
  assert.strictEqual(fs.readFileSync(v.file('Archive/Docs/a.docx', a.id)).toString(), 'one');
  assert.strictEqual(v.setLabel('Archive/Docs/a.docx', a.id, 'Start').label, 'Start');
  assert.strictEqual(v.file('Archive/Docs/a.docx', '../../x'), null);
  v.drop('Archive');
  assert.strictEqual(v.list('Archive/Docs/a.docx').length, 0);
});

test('trash: put, restore with free name, purge old', () => {
  const lib = tmp(); const t = new Trash(tmp());
  fs.mkdirSync(path.join(lib, 'A'));
  fs.writeFileSync(path.join(lib, 'A', 'x.docx'), 'x');
  const e = t.put(path.join(lib, 'A', 'x.docx'), 'A/x.docx');
  assert.ok(!fs.existsSync(path.join(lib, 'A', 'x.docx')));
  fs.writeFileSync(path.join(lib, 'A', 'x.docx'), 'new one');
  const r = t.restore(e.id, (rel) => path.join(lib, ...rel.split('/').filter(Boolean)), (rel) => fs.existsSync(path.join(lib, rel)));
  assert.strictEqual(r.rel, 'A/x (2).docx');
  assert.strictEqual(fs.readFileSync(path.join(lib, 'A', 'x (2).docx'), 'utf8'), 'x');
  const now = Date.now();
  assert.deepStrictEqual(trashExpired([{ id: '1', deletedAt: now - 31 * 86400000 }, { id: '2', deletedAt: now }], 30, now), ['1']);
});

test('metadata: patch validation, rekey follows renames, locks', () => {
  assert.deepStrictEqual(cleanPatch({ tags: ' a , b,a ', favorite: 1 }), { tags: ['a', 'b'], favorite: true });
  assert.throws(() => cleanPatch({ color: 'url(x)' }), /color/);
  const m = new Meta(path.join(tmp(), 'meta.json'));
  m.set('A/x.docx', { favorite: true, tags: ['t'] });
  m.opened({ key: 'A/x.docx', path: 'A/x.docx' });
  m.rekey('A', 'B');
  assert.strictEqual(m.get('A/x.docx'), null);
  assert.ok(m.get('B/x.docx').favorite);
  assert.strictEqual(m.recent()[0].path, 'B/x.docx');
  assert.deepStrictEqual(m.tags(), [{ tag: 't', count: 1 }]);
  const l = m.lock('B/x.docx', 'Anna');
  assert.strictEqual(m.lockOf('B/x.docx').by, 'Anna');
  assert.strictEqual(m.lockOf('B/x.docx', Date.now() + 13 * 3600000), null, 'locks expire');
  assert.ok(l.since);
  m.flush();
});

test('shares: expiry, lookup, passwords', () => {
  const now = Date.UTC(2026, 5, 1);
  assert.strictEqual(S.expiryOf(undefined, 7, now), now + 7 * 86400000);
  assert.strictEqual(S.expiryOf(0, 7, now), null);
  assert.strictEqual(S.expiryOf(3, 7, now), now + 3 * 86400000);
  assert.strictEqual(S.expiryOf(now + 5000, 7, now), now + 5000);
  assert.throws(() => S.expiryOf(now - 5000, 7, now), /future/);
  const token = S.newToken();
  assert.ok(S.TOKEN.test(token));
  const list = [{ token, expires: now + 1000 }, { token: S.newToken(), expires: null }];
  assert.ok(S.find(list, token, now));
  assert.strictEqual(S.find(list, token, now + 2000), null, 'expired');
  assert.strictEqual(S.find(list, 'short', now), null);
  assert.ok(S.find(list, list[1].token, now + 1e12), 'no expiry');
  const h = S.hashPassword('geheim');
  assert.ok(S.checkPassword(h, 'geheim'));
  assert.ok(!S.checkPassword(h, 'Geheim'));
  assert.ok(!S.checkPassword('plain', 'plain'));
  const page = S.page({ lang: 'de', share: { name: '<script>x</script>.docx', allowDownload: false }, base: '/api/office/pub/t', mode: 'text', text: '<b>hi</b>', size: 10 });
  assert.ok(!page.includes('<script>x'), 'names are escaped');
  assert.ok(page.includes('&lt;b&gt;hi'));
});

test('ONLYOFFICE file tokens: sign, verify, expiry, tampering, file id', () => {
  const secret = Buffer.alloc(32, 7);
  const now = 1_800_000_000_000;
  const id = W.fileIdOf(null, 'Docs/a b.docx');
  assert.deepStrictEqual(W.parseFileId(id), { root: null, path: 'Docs/a b.docx' });
  assert.deepStrictEqual(W.parseFileId(W.fileIdOf('htdocs', 'x.odt')), { root: 'htdocs', path: 'x.odt' });
  const t = W.signToken(secret, { k: 'Docs/a b.docx', r: null, p: 'Docs/a b.docx', w: true }, 60000, now);
  assert.strictEqual(W.verifyToken(secret, t, { now: now + 1000, fileId: id }).p, 'Docs/a b.docx');
  assert.strictEqual(W.verifyToken(secret, t, { now: now + 61000, fileId: id }), null, 'expired');
  assert.strictEqual(W.verifyToken(secret, t, { now, fileId: W.fileIdOf(null, 'other.docx') }), null, 'other file');
  assert.strictEqual(W.verifyToken(Buffer.alloc(32, 8), t, { now }), null, 'other secret');
  const [body, sig] = t.split('.');
  const forged = Buffer.from(JSON.stringify({ k: 'x', p: '../../etc', exp: now + 1e9 })).toString('base64url');
  assert.strictEqual(W.verifyToken(secret, `${forged}.${sig}`, { now }), null, 'tampered body');
  assert.strictEqual(W.verifyToken(secret, `${body}.${sig}x`, { now }), null);
  assert.strictEqual(W.verifyToken(secret, '', { now }), null);
});

test('ONLYOFFICE JWT (HS256), document key/type and container', () => {
  const tok = W.jwtSign({ document: { key: 'k' } }, 'secret');
  assert.deepStrictEqual(W.jwtVerify(tok, 'secret'), { document: { key: 'k' } });
  assert.strictEqual(W.jwtVerify(tok, 'other'), null);
  assert.strictEqual(W.jwtVerify(tok, ''), null);
  const none = `${Buffer.from('{"alg":"none"}').toString('base64url')}.${tok.split('.')[1]}.`;
  assert.strictEqual(W.jwtVerify(none, 'secret'), null);
  assert.strictEqual(W.jwtVerify(W.jwtSign({ exp: 1 }, 's'), 's'), null, 'expired');
  assert.match(W.ooKey('a', 123), /^[A-Za-z0-9_-]{20,128}$/);
  assert.notStrictEqual(W.ooKey('a', 123), W.ooKey('a', 124));
  assert.strictEqual(W.ooType('xlsx'), 'cell'); assert.strictEqual(W.ooType('pptx'), 'slide'); assert.strictEqual(W.ooType('odt'), 'word');
  assert.deepStrictEqual(Object.keys(APPS), ['onlyoffice']);
  const args = runArgs('onlyoffice', { secret: 's3' });
  assert.ok(args.includes('JWT_SECRET=s3') && args.includes('127.0.0.1:9981:80') && args[args.length - 1] === APPS.onlyoffice.image);
  assert.throws(() => runArgs('other', { secret: 's' }));
});

test('Office engine: version discovery, assets, checksum, msiexec arguments', () => {
  const html = '<a href="25.8.7/">25.8.7/</a> <a href="26.8.3/">26.8.3/</a> <a href="26.8.10/">x</a> <a href="/libreoffice/">up</a>';
  assert.deepStrictEqual(E.parseVersions(html), ['26.8.10', '26.8.3', '25.8.7']);
  assert.strictEqual(E.cmpVersion('26.8.10', '26.8.9'), 1);
  assert.deepStrictEqual(E.asset('26.8.3', 'win32', 'x64'), { url: 'https://download.documentfoundation.org/libreoffice/stable/26.8.3/win/x86_64/LibreOffice_26.8.3_Win_x86-64.msi', name: 'LibreOffice_26.8.3_Win_x86-64.msi', kind: 'msi' });
  assert.strictEqual(E.asset('26.8.3', 'linux', 'x64').url, 'https://download.documentfoundation.org/libreoffice/stable/26.8.3/deb/x86_64/LibreOffice_26.8.3_Linux_x86-64_deb.tar.gz');
  assert.strictEqual(E.asset('26.8.3', 'darwin', 'arm64').name, 'LibreOffice_26.8.3_MacOS_aarch64.dmg');
  assert.strictEqual(E.asset('1', 'aix', 'x64'), null);
  const sha = '0e6918c273eabbdbbd800ee0dacbbf130a14c13c2e4cec861b06dd59e015bf94';
  assert.strictEqual(E.parseSha256(`<li>MD5: ${'a'.repeat(32)}</li><li><a>SHA-256 Hash</a>:<br><tt>${sha}</tt>`), sha);
  assert.strictEqual(E.parseSha256('nothing'), null);
  assert.deepStrictEqual(E.msiArgs('C:\\d\\x.msi', 'C:\\Users\\A B\\office-engine.new', 'C:\\d\\l.log'), ['/a', '"C:\\d\\x.msi"', '/qn', 'TARGETDIR="C:\\Users\\A B\\office-engine.new"', '/l*v', '"C:\\d\\l.log"']);
  assert.throws(() => E.msiArgs('C:\\x".msi', 'C:\\t', 'C:\\l'), /Invalid path/);
  assert.ok(C.engineFiles('C:\\e').length >= 1);
});

test('converter plan: MS Office for Microsoft formats, LibreOffice for OpenDocument, nothing → null', () => {
  const ms = { word: 'w', excel: 'e', powerpoint: 'p', visio: null };
  const lo = { path: 'soffice' };
  assert.deepStrictEqual(C.plan('docx', 'pdf', { ms, lo: null }), { tool: 'ms', app: 'word' });
  assert.deepStrictEqual(C.plan('xlsx', 'pdf', { ms }), { tool: 'ms', app: 'excel' });
  assert.deepStrictEqual(C.plan('ppt', 'pptx', { ms }), { tool: 'ms', app: 'powerpoint' });
  assert.deepStrictEqual(C.plan('pdf', 'docx', { ms }), { tool: 'ms', app: 'word' });
  assert.deepStrictEqual(C.plan('odt', 'pdf', { ms, lo }), { tool: 'lo' });
  assert.deepStrictEqual(C.plan('docx', 'pdf', { ms, lo }, 'libreoffice'), { tool: 'lo' });
  assert.deepStrictEqual(C.plan('docx', 'pdf', { ms, lo, msEnabled: false }), { tool: 'lo' });
  assert.strictEqual(C.plan('vsdx', 'pdf', { ms }), null, 'no Visio');
  assert.strictEqual(C.plan('docx', 'pdf', { ms: null, lo: null }), null);
  assert.strictEqual(C.loFilter('xlsx', 'csv'), 'csv:Text - txt - csv (StarCalc):44,34,76,1');
  assert.strictEqual(C.loFilter('docx', 'txt'), 'txt:Text (encoded):UTF8');
  assert.strictEqual(C.loInfilter('pdf', 'docx'), 'writer_pdf_import');
  assert.ok(C.targets('docx', { ms, lo: null }).includes('pdf'));
  assert.ok(!C.PS_SCRIPT.includes('${'), 'the PowerShell script takes nothing from string interpolation');
});

test('Google: pointer files, PKCE, auth URL', () => {
  assert.strictEqual(G.parsePointer('{"url": "https://docs.google.com/open?id=1AbCdEfGhIjK", "doc_id": "1AbCdEfGhIjKlMnOp", "email": "a@b.c"}'), '1AbCdEfGhIjKlMnOp');
  assert.strictEqual(G.parsePointer('{"resource_id": "spreadsheet:1ZyXwVuTsRqPoNm"}'), '1ZyXwVuTsRqPoNm');
  assert.strictEqual(G.parsePointer('{"url": "https://docs.google.com/presentation/d/1QwErTyUiOpAs/edit"}'), '1QwErTyUiOpAs');
  assert.strictEqual(G.parsePointer('no json'), null);
  assert.strictEqual(G.parsePointer('{"doc_id": "../../etc"}'), null);
  const p = G.pkce();
  assert.match(p.verifier, /^[A-Za-z0-9_-]{43,128}$/);
  const u = new URL(G.authUrl({ clientId: 'x.apps.googleusercontent.com', redirectUri: 'http://127.0.0.1:7780/api/office/google/callback', challenge: p.challenge, state: 's' }));
  assert.strictEqual(u.searchParams.get('code_challenge_method'), 'S256');
  assert.strictEqual(u.searchParams.get('redirect_uri'), 'http://127.0.0.1:7780/api/office/google/callback');
  assert.strictEqual(G.entryOf({ id: '1', name: 'Plan', mimeType: 'application/vnd.google-apps.spreadsheet' }).ext, 'xlsx');
});

test('snappy (iWork IWA chunks)', () => {
  // "abcabcabcabc": literal "abc" + copy (offset 3, length 9)
  const block = Buffer.from([12, (3 - 1) << 2, 0x61, 0x62, 0x63, ((9 - 4) << 2) | 1, 3]);
  assert.strictEqual(snappy(block).toString(), 'abcabcabcabc');
});
