'use strict';
/*
 * Soft locks and the lock marks of Hostwerk's Files (server/meta.js, server/marks.js): the key a lock is taken under is
 * the key the folder's marks look up – for files of any root and of the library root; requests of one editor in order
 * (a late heartbeat never brings a released lock back, a late unlock never releases the next editor's lock); the own
 * window never sees its own lock; locks expire; the same user always takes a stale lock over.
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const SET = require('../server/settings');
const { Meta } = require('../server/meta');
const { Versions } = require('../server/versions');
const { Trash } = require('../server/trash');
const { Library } = require('../server/library');
const { markKey, folderMarks } = require('../server/marks');

const made = [];
const tmp = () => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'office-locks-')); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });

function filesCore(roots) {
  return {
    resolve(root, rel) {
      const base = roots[root];
      if (!base) throw Object.assign(new Error('unknown root'), { status: 400 });
      const full = path.resolve(base, `.${path.sep}${String(rel || '')}`);
      if (full !== path.resolve(base) && !full.startsWith(path.resolve(base) + path.sep)) throw Object.assign(new Error('outside'), { status: 403 });
      return { base, full };
    },
  };
}
function setup() {
  const dir = tmp();
  const libDir = path.join(dir, 'office');
  const docs = path.join(dir, 'Documents');
  fs.mkdirSync(path.join(docs, 'Projekte - Claude', 'ISR'), { recursive: true });
  fs.mkdirSync(path.join(libDir, 'Work'), { recursive: true });
  fs.writeFileSync(path.join(docs, 'Projekte - Claude', 'ISR', 'Plan v0.3.docx'), 'x');
  fs.writeFileSync(path.join(docs, 'Projekte - Claude', 'ISR', 'notes.txt'), 'x');
  fs.writeFileSync(path.join(libDir, 'Work', 'Offer.docx'), 'x');
  const settings = SET.normalize({}, libDir);
  const meta = new Meta(path.join(dir, 'meta.json'));
  const lib = new Library({ settings: () => settings, files: filesCore({ 'ext:office:library': libDir, documents: docs, home: dir }), meta, versions: new Versions(path.join(dir, 'versions')), trash: new Trash(path.join(dir, 'trash')), index: null, sharedKeys: () => new Set() });
  const marksOf = (root, dirRel, client = null, now) => { const d = lib.loc({ root, path: dirRel }); return folderMarks({ d, list: fs.readdirSync(d.full, { withFileTypes: true }), meta, shared: new Set(), client, now }).names; };
  return { lib, meta, marksOf };
}

test('lock key: a file of a non-library root – the lock of the editor is the one the folder marks find', () => {
  const { lib, meta, marksOf } = setup();
  const f = lib.loc({ root: 'documents', path: 'Projekte - Claude/ISR/Plan v0.3.docx' });
  assert.strictEqual(f.key, '@documents/Projekte - Claude/ISR/Plan v0.3.docx');
  assert.strictEqual(markKey(lib.loc({ root: 'documents', path: 'Projekte - Claude/ISR' }), 'Plan v0.3.docx'), f.key);
  // the same folder asked with backslashes / trailing slash (Windows) gives the same key
  assert.strictEqual(markKey(lib.loc({ root: 'documents', path: 'Projekte - Claude\\ISR\\' }), 'Plan v0.3.docx'), f.key);
  meta.lock(f.key, 'Anna', Date.now(), { client: 'w1', sid: 'e1', seq: 1 });
  assert.deepStrictEqual(Object.keys(marksOf('documents', 'Projekte - Claude/ISR', 'w2')), ['Plan v0.3.docx']);
  assert.strictEqual(marksOf('documents', 'Projekte - Claude/ISR', 'w2')['Plan v0.3.docx'].lock.by, 'Anna');
  assert.ok(marksOf('documents', 'Projekte - Claude/ISR', 'w2')['Plan v0.3.docx'].lock.since > 0, 'since for the tooltip');
  assert.strictEqual(meta.unlock(lib.loc({ root: 'documents', path: 'Projekte - Claude/ISR/Plan v0.3.docx' }).key, { sid: 'e1', seq: 2 }), true);
  assert.deepStrictEqual(marksOf('documents', 'Projekte - Claude/ISR', 'w2'), {}, 'mark gone after the unlock');
});

test('lock key: the library root and a root that contains the library use library paths', () => {
  const { lib, meta, marksOf } = setup();
  const viaRoot = lib.loc({ root: 'ext:office:library', path: 'Work/Offer.docx' });
  const viaHome = lib.loc({ root: 'home', path: 'office/Work/Offer.docx' });
  assert.strictEqual(viaRoot.key, 'Work/Offer.docx');
  assert.strictEqual(viaHome.key, 'Work/Offer.docx');
  meta.lock(viaHome.key, 'Anna', Date.now(), { client: 'w1', sid: 'e1', seq: 1 });
  assert.ok(marksOf('ext:office:library', 'Work', 'w2')['Offer.docx'].lock);
  assert.ok(marksOf(null, 'Work', 'w2')['Offer.docx'].lock);
  meta.unlock(viaRoot.key, { sid: 'e1', seq: 2 });
  assert.deepStrictEqual(marksOf('home', 'office/Work', 'w2'), {});
});

test('own window: Files never shows the lock of its own window (its editor is closed when Files shows)', () => {
  const { lib, meta, marksOf } = setup();
  const f = lib.loc({ root: 'documents', path: 'Projekte - Claude/ISR/Plan v0.3.docx' });
  meta.lock(f.key, 'Anna', Date.now(), { client: 'w1', sid: 'e1', seq: 1 });
  assert.deepStrictEqual(marksOf('documents', 'Projekte - Claude/ISR', 'w1'), {}, 'own window: no mark');
  assert.ok(marksOf('documents', 'Projekte - Claude/ISR', 'w2')['Plan v0.3.docx'].lock, 'another window: mark');
  assert.ok(marksOf('documents', 'Projekte - Claude/ISR', null)['Plan v0.3.docx'].lock, 'older UI without client: mark');
});

test('order: a heartbeat arriving after the unlock never brings the lock back', () => {
  const { meta } = setup();
  const k = '@documents/a.docx';
  meta.lock(k, 'Anna', Date.now(), { client: 'w1', sid: 'e1', seq: 1 });
  // the heartbeat (seq 2) is still on its way when the editor closes (unlock seq 3) – it arrives last
  assert.strictEqual(meta.unlock(k, { sid: 'e1', seq: 3 }), true);
  const late = meta.lock(k, 'Anna', Date.now(), { client: 'w1', sid: 'e1', seq: 2 });
  assert.ok(late.ignored);
  assert.strictEqual(meta.lockOf(k), null);
  // the same editor coming back (bfcache, window shown again) locks with a newer seq
  meta.lock(k, 'Anna', Date.now(), { client: 'w1', sid: 'e1', seq: 4 });
  assert.ok(meta.lockOf(k));
});

test('order: the late unlock of a closed editor never releases the lock of the editor that opened the file again', () => {
  const { meta } = setup();
  const k = '@documents/a.docx';
  meta.lock(k, 'Anna', Date.now(), { client: 'w1', sid: 'old', seq: 1 });
  meta.lock(k, 'Anna', Date.now(), { client: 'w1', sid: 'new', seq: 1 }); // opened again – takes over
  assert.strictEqual(meta.unlock(k, { sid: 'old', seq: 2 }), false);
  assert.strictEqual(meta.lockOf(k).by, 'Anna');
  assert.strictEqual(meta.unlock(k, { sid: 'new', seq: 2 }), true);
  // without sid (older UI, trash): unconditional
  meta.lock(k, 'Anna'); assert.strictEqual(meta.unlock(k), true);
});

test('expiry: an editor that never unlocked leaves the mark for at most 5 minutes; renewals keep since', () => {
  const { lib, meta, marksOf } = setup();
  const f = lib.loc({ root: 'documents', path: 'Projekte - Claude/ISR/Plan v0.3.docx' });
  const t0 = Date.now();
  const a = meta.lock(f.key, 'Anna', t0, { client: 'w1', sid: 'e1', seq: 1 });
  assert.strictEqual(a.renewed, false);
  const b = meta.lock(f.key, 'Anna', t0 + 120000, { client: 'w1', sid: 'e1', seq: 2 });
  assert.strictEqual(b.renewed, true, 'a heartbeat is no change for the marks');
  assert.strictEqual(b.since, t0);
  assert.ok(marksOf('documents', 'Projekte - Claude/ISR', 'w2', t0 + 120000 + 4 * 60000)['Plan v0.3.docx']);
  assert.deepStrictEqual(marksOf('documents', 'Projekte - Claude/ISR', 'w2', t0 + 120000 + 5 * 60000 + 1), {}, 'expired');
  assert.strictEqual(meta.lockOf(f.key), null);
});

test('Entry.lock stays { by, since } (no window ids in entries)', () => {
  const { meta } = setup();
  meta.lock('x.docx', 'Anna', Date.now(), { client: 'w1', sid: 'e1', seq: 1 });
  assert.deepStrictEqual(Object.keys(meta.fields('x.docx').lock).sort(), ['by', 'since']);
});

test('routes: lock/unlock take client/sid/seq, the same user takes a stale lock over, office-lock broadcast on change only', async () => {
  const handlers = {};
  const mk = (m) => (p, ...fns) => { handlers[`${m} ${p}`] = fns[fns.length - 1]; };
  const broadcasts = [];
  const { lib, meta } = setup();
  const host = {
    router: { get: mk('GET'), post: mk('POST'), put: mk('PUT'), delete: mk('DELETE'), use: () => {} },
    publicRouter: { get: () => {}, post: () => {}, put: () => {}, delete: () => {}, use: () => {} },
    lib: () => { const e = () => () => {}; e.raw = () => () => {}; e.urlencoded = () => () => {}; e.json = () => () => {}; e.static = () => () => {}; return e; },
    error: (s, m) => Object.assign(new Error(m), { status: s }), t: (m, v) => String(m).replace(/\{(\w+)\}/g, (x, k) => (v && v[k] != null ? v[k] : x)),
    broadcast: (m) => broadcasts.push(m), log: () => {},
  };
  const o = { library: lib, meta, userName: () => 'Anna', shareList: () => [], settings: () => ({ integrations: {} }) };
  try { require('../server/routes').routes(host, o); } catch { /* routes past the lock ones may need more of o – not needed here */ }
  assert.ok(handlers['POST /office/lock'] && handlers['POST /office/unlock'], 'lock routes registered');
  const call = (p, body) => new Promise((resolve) => { const res = { statusCode: 200, status(c) { this.statusCode = c; return this; }, json(j) { resolve({ status: this.statusCode, body: j }); } }; handlers[p]({ body, query: {} }, res); });
  const f = { root: 'documents', path: 'Projekte - Claude/ISR/Plan v0.3.docx' };
  let r = await call('POST /office/lock', { ...f, client: 'w1', sid: 'e1', seq: 1 });
  assert.strictEqual(r.status, 200);
  assert.deepStrictEqual(broadcasts.map((m) => [m.type, m.locked]), [['office-lock', true]]);
  await call('POST /office/lock', { ...f, client: 'w1', sid: 'e1', seq: 2 });
  assert.strictEqual(broadcasts.length, 1, 'heartbeat: no broadcast');
  // a stale lock of a crashed window of the same user: opening again takes it over silently
  r = await call('POST /office/lock', { ...f, client: 'w9', sid: 'e9', seq: 1 });
  assert.strictEqual(r.status, 200); assert.strictEqual(r.body.lock.other, false);
  // someone else: 409 with by/since
  meta.lock(lib.loc(f).key, 'Ben', Date.now(), { client: 'wB', sid: 'eB', seq: 1 });
  r = await call('POST /office/lock', { ...f, client: 'w1', sid: 'e1', seq: 3 });
  assert.strictEqual(r.status, 409); assert.strictEqual(r.body.lock.by, 'Ben'); assert.ok(r.body.lock.since);
  meta.lock(lib.loc(f).key, 'Anna', Date.now(), { client: 'w1', sid: 'e1', seq: 4 });
  broadcasts.length = 0;
  r = await call('POST /office/unlock', { ...f, client: 'w1', sid: 'e1', seq: 5 });
  assert.strictEqual(r.body.released, true);
  assert.deepStrictEqual(broadcasts.map((m) => [m.type, m.locked, m.root, m.path]), [['office-lock', false, 'documents', f.path]]);
  assert.strictEqual(meta.lockOf(lib.loc(f).key), null);
});
