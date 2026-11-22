'use strict';
/*
 * HTTP routes of the Office extension (contract: ARCHITECTURE.md → Backend API). host.router = /api/office/… with
 * Hostwerk's login; host.publicRouter = /api/office/pub/… (share links) and /api/office/wopi/oo/… (ONLYOFFICE file +
 * callback, token – the path name is kept so links of running Document Servers stay valid).
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { FORMATS, extOf, mimeOf } = require('./formats');
const { Library } = require('./library');
const { ZipWriter } = require('./zip');
const X = require('./extract');
const S = require('./shares');
const W = require('./onlyoffice');
const G = require('./google');
const desktop = require('./desktop');
const { view: versionView } = require('./versions');
const { checkName, normRel } = require('./paths');
const { folderMarks } = require('./marks');

const IMG = /^(png|jpe?g|gif|webp|bmp|avif)$/;
const ACTIVE = /^(html?|xhtml|svg|xml|mht|mhtml)$/; // could run scripts when opened directly → sandboxed
const MAX_READ = 200 * 1048576;
/** Script of the embedding page: load ONLYOFFICE's api.js and start the editor */
const FRAME_JS = `(function () {
  var c = JSON.parse(document.getElementById('cfg').textContent);
  var box = document.getElementById('editor');
  function send(event, message, extra) { var m = { source: 'office-frame', app: c.app, event: event }; if (message) m.message = message; for (var k in (extra || {})) m[k] = extra[k]; parent.postMessage(m, location.origin); }
  function fail(m) { box.textContent = m; box.style.cssText += ';display:flex;align-items:center;justify-content:center;font:14px system-ui;color:#444'; send('error', m); }
  var s = document.createElement('script'); s.src = c.apiUrl;
  s.onerror = function () { fail('The ONLYOFFICE Document Server does not answer.'); };
  s.onload = function () {
    try {
      var cfg = c.config; cfg.width = '100%'; cfg.height = '100%';
      cfg.events = { onAppReady: function () { send('ready'); }, onDocumentStateChange: function (ev) { send(ev && ev.data ? 'modified' : 'saved'); }, onError: function (ev) { send('error', ev && ev.data && (ev.data.errorDescription || String(ev.data.errorCode))); } };
      new window.DocsAPI.DocEditor('editor', cfg);
    } catch (e) { fail(String(e && e.message || e)); }
  };
  document.head.appendChild(s);
})();`;

const contentDisposition = (type, name) => `${type}; filename="${String(name).replace(/[^\x20-\x7e]|["\\]/g, '_')}"; filename*=UTF-8''${encodeURIComponent(name)}`;
const boolQ = (v) => v === '1' || v === 'true' || v === true || v === 1;

function routes(host, o) {
  const { router: r, publicRouter: pub } = host;
  const express = host.lib('express');
  const raw = express.raw({ type: () => true, limit: '200mb' });
  const form = express.urlencoded({ extended: false, limit: '8kb' });
  const lib = o.library;
  const err = (status, text, vars, extra) => host.error(status, text, vars, extra ? { extra } : undefined);
  const q = (req) => ({ root: req.query.root ? String(req.query.root) : null, path: String(req.query.path || '') });
  const b = (req) => ({ root: req.body && req.body.root ? String(req.body.root) : null, path: String((req.body && req.body.path) || '') });

  /* ---------------- status & settings ---------------- */
  /** Is the ONLYOFFICE Document Server reachable? → { ok, version?, error? } (cached 30 s, fresh with ?test=1) */
  const probes = new Map();
  const probe = async (url, fresh) => {
    if (!url) return { ok: false, error: host.t('Not set up') };
    const c = probes.get('onlyoffice');
    if (!fresh && c && c.url === url && Date.now() - c.t < 30000) return c.v;
    let v;
    try {
      const base = url.replace(/\/+$/, '');
      const r0 = await W.fetchUrl(`${base}/healthcheck`, { timeout: 5000, max: 65536 });
      if (r0.status !== 200 || !/true/.test(r0.body.toString('utf8'))) throw new Error(`HTTP ${r0.status}`);
      let version = null;
      try { const r1 = await W.fetchUrl(`${base}/web-apps/apps/api/documents/api.js`, { timeout: 5000, max: 4 * 1048576 }); const m = r1.body.toString('utf8').match(/@version\s+([\d.]+)|DocsAPI\.DocEditor\.version\s*=\s*function\s*\(\)\s*\{\s*return\s*['"]([\d.]+)/); version = m ? m[1] || m[2] : null; } catch { /* only the version is missing */ }
      v = { ok: true, version };
    } catch (e) { v = { ok: false, error: host.t('{name} does not answer: {error}', { name: 'ONLYOFFICE', error: e.message }) }; }
    probes.set('onlyoffice', { url, t: Date.now(), v });
    return v;
  };
  r.get('/office/status', async (req, res) => {
    const c = await o.converters(boolQ(req.query.fresh));
    const s = o.settings();
    const test = boolQ(req.query.test);
    const po = await probe(s.integrations.onlyoffice.url, test);
    res.json({ version: host.version, libraryDir: s.libraryDir, platform: process.platform,
      converters: { libreoffice: c.lo ? { path: c.lo.path, version: c.lo.version, source: c.lo.source } : null, msoffice: c.msEnabled ? c.ms : null },
      engine: o.engine.status(),
      integrations: { onlyoffice: { ...po, url: s.integrations.onlyoffice.url || '', configured: !!s.integrations.onlyoffice.url, jwt: !!s.integrations.onlyoffice.secret }, defaultOpen: s.integrations.defaultOpen, google: { configured: !!s.integrations.google.clientId, connected: !!o.drive.tokens() } },
      desktop: o.hasDesktop(), webWindow: o.canWindow(), libraryRoot: o.libRoot ? o.libRoot() : null });
  });

  /* ---------------- "Open with" / default app (SDK 3.15 host.fileTypes – osreg.js) ---------------- */
  // { supported, platform, code, dev, enabled, registered, stale, kinds: { docs: { state: default|partial|openwith|none } … }, notice }
  r.get('/office/os', async (req, res) => res.json(await o.osreg.state()));
  r.post('/office/os/register', async (req, res) => res.json(await o.osreg.register({ notice: false })));
  r.post('/office/os/unregister', async (req, res) => res.json(await o.osreg.unregister({ byUser: true })));
  r.post('/office/os/defaults', async (req, res) => res.json(await o.osreg.openDefaults()));
  // Windows' "How do you want to open this file?" for one kind – body (raw, optional): an empty file of that type
  r.post('/office/os/openwith', express.raw({ type: () => true, limit: '8mb' }), async (req, res) => res.json(await o.osreg.openWith(String(req.query.kind || ''), Buffer.isBuffer(req.body) ? req.body : null)));
  r.post('/office/os/seen', (req, res) => res.json(o.osreg.seen()));
  // live refresh: { type: 'office' } after every successful change (debounced)
  let bcTimer = null;
  r.use('/office', (req, res, next) => {
    if (req.method !== 'GET' && req.method !== 'HEAD' && !/^\/(un)?lock\/?$/.test(req.path)) res.on('finish', () => { if (res.statusCode < 400 && !bcTimer) { bcTimer = setTimeout(() => { bcTimer = null; try { host.broadcast({ type: 'office' }); } catch { /* egal */ } }, 250); if (bcTimer.unref) bcTimer.unref(); } });
    next();
  });
  r.get('/office/settings', (req, res) => res.json(o.settings()));
  r.put('/office/settings', (req, res) => res.json(o.saveSettings(req.body || {})));

  /* ---------------- listing & search ---------------- */
  r.get('/office/files', (req, res) => {
    const s = o.settings();
    const view = String(req.query.view || 'all');
    const dir = normRel(req.query.dir || '');
    const filt = { kind: req.query.kind || '', family: req.query.family || '', tag: req.query.tag || '', q: req.query.q || '', ext: req.query.ext || '' };
    const sort = String(req.query.sort || (view === 'recent' ? 'opened' : view === 'trash' ? 'deleted' : s.general.sort));
    let entries;
    if (view === 'recent') entries = lib.fromKeys(o.meta.recent(200)).slice(0, s.general.recentCount);
    else if (view === 'favorites') entries = lib.fromKeys(o.meta.favorites());
    else if (view === 'shared') entries = lib.fromKeys([...new Set(o.shareList().filter((x) => !S.expired(x)).map((x) => x.key))]);
    else if (view === 'trash') {
      entries = o.trash.list().map((e) => { const ext = e.dir ? '' : extOf(e.name); const f = FORMATS[ext] || {}; return { id: e.id, path: e.path, name: e.name, dir: !!e.dir, ext, kind: e.dir ? null : f.kind || null, family: e.dir ? null : f.family || null, size: e.size || 0, mtime: Math.floor(e.mtime || e.deletedAt), ctime: Math.floor(e.mtime || e.deletedAt), deletedAt: e.deletedAt, from: e.from, purgeAt: e.deletedAt + s.general.trashDays * 86400000, favorite: false, tags: [], color: '', note: '', lock: null, shared: false, thumb: false }; });
    } else if (view === 'templates') {
      entries = o.tplStore.list().map((t) => ({ id: t.id, template: true, path: t.path != null ? t.path : '', name: t.name, category: t.category, source: t.source, readOnly: t.readOnly, default: t.default, dir: false, ext: t.ext, kind: t.kind, family: t.family, size: t.size, mtime: t.mtime, ctime: t.created, favorite: false, tags: [], color: '', note: '', lock: null, shared: false, thumb: false }));
    } else entries = lib.listDir(dir, { recursive: !!filt.q || !!filt.tag });
    entries = Library.filter(entries, filt);
    if (view !== 'recent' || req.query.sort) entries = Library.sortEntries(entries, sort);
    const counts = { favorites: lib.fromKeys(o.meta.favorites()).length, trash: o.trash.list().length, shared: new Set(o.shareList().filter((x) => !S.expired(x)).map((x) => x.key)).size, templates: o.tplStore.list().length };
    res.json({ dir, view, entries, counts, crumbs: lib.crumbs(dir).map((c, i) => (i === 0 ? { ...c, name: host.t('Library') } : c)) });
  });
  r.get('/office/search', async (req, res) => {
    const text = String(req.query.q || '').trim();
    if (!text) return res.json({ hits: [], indexing: !!o.index.running });
    // ?semantic=1: search by meaning (1.1, semantic.js) – same hit shape plus mode
    if (boolQ(req.query.semantic) && o.semantic) {
      const out = await o.semantic.search(text, { dir: normRel(req.query.dir || ''), limit: 60 });
      const kind = String(req.query.kind || '');
      const hits = []; for (const h of out.hits) { try { const e = lib.entryOf({ path: h.rel }); if (!kind || e.kind === kind) hits.push({ ...e, snippet: h.snippet, score: h.score }); } catch { /* gone */ } }
      return res.json({ hits, indexing: !!o.index.running, mode: out.mode });
    }
    const dir = normRel(req.query.dir || '');
    const kind = String(req.query.kind || '');
    const hits = [];
    for (const h of o.index.search(text, { dir, filter: kind ? (rel) => (FORMATS[extOf(rel)] || {}).kind === kind : null, limit: 200 })) {
      try { hits.push({ ...lib.entryOf({ path: h.rel }), snippet: h.snippet, score: h.score }); } catch { /* gone */ }
    }
    res.json({ hits, indexing: !!o.index.running });
  });

  /* ---------------- file operations ---------------- */
  r.post('/office/folder', (req, res) => res.json(lib.mkdir((req.body || {}).dir || '', (req.body || {}).name)));
  r.post('/office/rename', (req, res) => res.json(lib.rename((req.body || {}).root ? b(req) : (req.body || {}).path, (req.body || {}).name)));
  r.post('/office/move', (req, res) => res.json({ entries: lib.move((req.body || {}).paths, (req.body || {}).to || '') }));
  r.post('/office/copy', (req, res) => res.json({ entries: lib.copy((req.body || {}).paths, (req.body || {}).to || '') }));
  r.post('/office/trash', (req, res) => { const entries = lib.toTrash((req.body || {}).paths); res.json({ ids: entries.map((e) => e.id), entries }); });
  r.post('/office/trash/restore', (req, res) => res.json({ entries: lib.restore((req.body || {}).ids) }));
  r.post('/office/trash/purge', (req, res) => res.json({ ok: true, removed: lib.purge((req.body || {}).ids) }));
  r.put('/office/meta', (req, res) => { if (!b(req).path) throw err(400, 'Invalid path'); const l = lib.loc(b(req)); if (!fs.existsSync(l.full)) throw err(404, 'File not found'); o.meta.set(l.key, req.body || {}); res.json(lib.entry(l)); });
  r.get('/office/tags', (req, res) => res.json({ tags: o.meta.tags() }));
  /**
   * Marks for a folder of Hostwerk's Files (files.marks): GET /office/marks?root=&path= → { names: { 'Report.docx':
   * { favorite, shared, lock: { by } , tags } } } – only office files that have something to show.
   */
  r.get('/office/marks', (req, res) => {
    const root = req.query.root ? String(req.query.root) : null;
    const dir = String(req.query.path || '').replace(/[\\]/g, '/').split('/').filter((x) => x && x !== '.').join('/');
    if (dir.split('/').includes('..')) throw err(403, 'Invalid path');
    const d = lib.loc(root ? { root, path: dir } : { path: dir });
    let list = [];
    try { list = fs.readdirSync(d.full, { withFileTypes: true }); } catch { return res.json({ names: {} }); }
    const shared = new Set(o.shareList().filter((x) => !S.expired(x)).map((x) => x.key));
    // types=1: every office file of the folder (Hostwerk 3.2.x without files.icon shows the file type as a mark);
    // client: the asking window – the locks of its own (closed) editors are not shown
    res.json(folderMarks({ d, list, meta: o.meta, shared, client: req.query.client ? String(req.query.client) : null, types: req.query.types === '1' }));
  });

  /* ---------------- raw ---------------- */
  // .gdoc/.gsheet/.gslides/.gdraw: the bytes are the Google file itself (exported as docx/xlsx/pptx/pdf) – pointer=1 gives the JSON
  const GPTR = /^g(doc|sheet|slides|draw)$/;
  const googleIdOf = (full) => { const id = G.parsePointer(fs.readFileSync(full, 'utf8').slice(0, 20000)); if (!id) throw err(422, 'This Google file does not contain a document id.'); return id; };
  r.get('/office/raw', async (req, res, next) => {
    if (boolQ(req.query.pointer) || !GPTR.test(extOf(String(req.query.path || '')))) return next();
    const f = lib.read(q(req));
    if (!o.drive.tokens()) throw err(409, 'Connect your Google account first (Office → Settings → Integrations).', null, { google: true });
    const g = await o.drive.read(googleIdOf(f.full), req.query.as ? String(req.query.as) : '');
    o.meta.opened({ key: f.key, root: f.root, path: f.rel });
    res.setHeader('Content-Type', g.mime); res.setHeader('Cache-Control', 'no-store');
    res.setHeader('x-mtime', String(f.mtime)); res.setHeader('x-google', '1'); res.setHeader('x-ext', g.ext);
    res.setHeader('Access-Control-Expose-Headers', 'x-mtime, x-google, x-ext');
    if (boolQ(req.query.download)) res.setHeader('Content-Disposition', contentDisposition('attachment', g.name));
    res.end(g.data);
  });
  r.put('/office/raw', raw, async (req, res, next) => {
    if (boolQ(req.query.pointer) || !GPTR.test(extOf(String(req.query.path || '')))) return next();
    const f = lib.read(q(req));
    if (!o.drive.tokens()) throw err(409, 'Connect your Google account first (Office → Settings → Integrations).', null, { google: true });
    const ext = String(req.query.ext || { gdoc: 'docx', gsheet: 'xlsx', gslides: 'pptx', gdraw: 'pdf' }[extOf(f.full)]);
    const g = await o.drive.write(googleIdOf(f.full), Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0), ext);
    res.json({ ok: true, mtime: f.mtime, size: g.size || 0, version: null, google: true });
  });
  r.get('/office/raw', (req, res) => {
    const f = lib.read(q(req));
    if (f.st.size > MAX_READ * 5) throw err(413, 'The file is too large.');
    const name = path.basename(f.full);
    const dl = boolQ(req.query.download);
    res.setHeader('x-mtime', String(f.mtime));
    res.setHeader('Access-Control-Expose-Headers', 'x-mtime');
    res.setHeader('Content-Type', mimeOf(name));
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    if (ACTIVE.test(extOf(name))) res.setHeader('Content-Security-Policy', 'sandbox; default-src \'none\'; img-src data:; style-src \'unsafe-inline\'');
    res.setHeader('Content-Disposition', contentDisposition(dl ? 'attachment' : 'inline', name));
    if (!dl) o.meta.opened({ key: f.key, root: f.root, path: f.rel });
    fs.createReadStream(f.full).on('error', (e) => res.destroy(e)).pipe(res);
  });
  r.put('/office/raw', raw, (req, res) => {
    const data = Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0);
    // a NEW file (also in a Hostwerk root, e.g. "Convert to PDF" next to a file-browser file): the name must be valid
    const loc0 = lib.loc(q(req));
    if (!fs.existsSync(loc0.full)) checkName(path.basename(loc0.full));
    const out = lib.write(q(req), data, { mtime: req.query.mtime, force: boolQ(req.query.force), label: String(req.query.label || '').slice(0, 120) });
    res.json(out);
  });
  r.post('/office/create', raw, (req, res) => {
    const data = Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0);
    // root = a folder of Hostwerk's Files (files.new "New document" in any folder) – else the library
    const dir = req.query.root ? { root: String(req.query.root), dir: String(req.query.dir || '') } : req.query.dir || '';
    res.json(lib.create(dir, String(req.query.name || ''), data, { keep: boolQ(req.query.keep) }));
  });
  r.get('/office/info', async (req, res) => {
    const f = lib.file(q(req));
    const e = lib.entry(f, f.st);
    let facts = {};
    if (f.st.size <= MAX_READ) { try { facts = await X.info(await fs.promises.readFile(f.full), e.ext, o.libs()); } catch (x) { facts = { error: host.t('The file could not be read: {error}', { error: x.message }) }; } }
    const label = (FORMATS[e.ext] || {}).label;
    res.json({ ...e, ...facts, label: label ? host.t(label) : null, edit: (FORMATS[e.ext] || {}).edit || null, versions: o.versions.list(f.key).length, full: f.full });
  });
  // embedded thumbnail (docProps/thumbnail.*, ODF Thumbnails/…) or one the UI rendered (PUT /office/thumb, per mtime); ETag/304
  const thumbMiss = new Map(); // key → mtime without an embedded thumbnail (no need to open the zip again)
  r.get('/office/thumb', async (req, res) => {
    const f = lib.file(q(req));
    const mtime = Math.floor(f.st.mtimeMs);
    const etag = `"t${mtime.toString(36)}-${f.st.size.toString(36)}"`;
    res.setHeader('Cache-Control', 'private, max-age=60, must-revalidate');
    res.setHeader('ETag', etag);
    if (req.headers['if-none-match'] === etag) return res.status(304).end();
    let t = null;
    if (f.st.size <= MAX_READ && thumbMiss.get(f.key) !== mtime) {
      t = X.thumb(await fs.promises.readFile(f.full), extOf(f.full));
      if (!t) { thumbMiss.set(f.key, mtime); if (thumbMiss.size > 5000) thumbMiss.clear(); }
    }
    if (!t && o.thumbs) { const got = o.thumbs.get(f.key, mtime); if (got) t = { ...got, rendered: true }; }
    if (!t) { res.removeHeader('ETag'); if (req.query.probe) return res.status(204).end(); throw err(404, 'No thumbnail'); } // probe=1: “is there one?” without an error in the browser console
    res.setHeader('Content-Type', t.type);
    res.setHeader('X-Content-Type-Options', 'nosniff');
    if (t.rendered) res.setHeader('x-thumb', 'rendered');
    res.end(t.data);
  });

  /* ---------------- versions ---------------- */
  r.get('/office/versions', (req, res) => { const l = lib.loc(q(req)); res.json({ versions: o.versions.list(l.key).map(versionView) }); });
  r.get('/office/versions/raw', (req, res) => {
    const l = lib.loc(q(req));
    const p = o.versions.file(l.key, String(req.query.id || ''));
    if (!p) throw err(404, 'Version not found');
    res.setHeader('Content-Type', mimeOf(l.full));
    res.setHeader('Cache-Control', 'no-store');
    if (boolQ(req.query.download)) { const v = o.versions.get(l.key, String(req.query.id)); const ext = path.extname(l.full); res.setHeader('Content-Disposition', contentDisposition('attachment', `${path.basename(l.full, ext)} (${new Date(v ? v.time : Date.now()).toISOString().slice(0, 16).replace(/[T:]/g, '-')})${ext}`)); }
    fs.createReadStream(p).pipe(res);
  });
  r.post('/office/versions/restore', (req, res) => res.json(lib.restoreVersion(b(req), String((req.body || {}).id || ''))));
  r.put('/office/versions/label', (req, res) => { const l = lib.loc(b(req)); const v = o.versions.setLabel(l.key, String((req.body || {}).id || ''), (req.body || {}).label); if (!v) throw err(404, 'Version not found'); res.json(versionView(v)); });
  r.delete('/office/versions', (req, res) => { const l = lib.loc(q(req)); o.versions.remove(l.key, String(req.query.id || '')); res.json({ ok: true }); });

  /* ---------------- templates ---------------- */
  require('./routesTpl').templateRoutes(host, o, { raw, err, b, q });

  /* ---------------- convert ---------------- */
  r.post('/office/convert', raw, async (req, res) => {
    const data = Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0);
    if (!data.length) throw err(400, 'The file is empty.');
    const from = String(req.query.from || ''); const to = String(req.query.to || '');
    let out;
    try { out = await o.convert(data, from, to); } catch (e) { throw err(e.status || 500, e.raw || e.message, e.vars || { from, to }); }
    res.setHeader('Content-Type', mimeOf(`x.${to}`));
    res.setHeader('x-converter', out.tool);
    res.setHeader('Access-Control-Expose-Headers', 'x-converter');
    res.end(out.data);
  });
  r.get('/office/convert/targets', async (req, res) => res.json({ targets: o.targetsFor(String(req.query.from || '').toLowerCase(), await o.converters()) }));

  /* ---------------- zip ---------------- */
  r.get('/office/zip', async (req, res) => {
    let paths; try { paths = JSON.parse(String(req.query.paths || '[]')); } catch { throw err(400, 'Invalid list of files'); }
    if (!Array.isArray(paths) || !paths.length) throw err(400, 'Invalid list of files');
    const items = paths.slice(0, 1000).map((p) => lib.loc({ path: p })).filter((l) => l.rel && fs.existsSync(l.full));
    if (!items.length) throw err(404, 'File not found');
    const name = items.length === 1 ? `${path.basename(items[0].full).replace(/\.[^.]+$/, '')}.zip` : `Office ${new Date().toISOString().slice(0, 10)}.zip`;
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', contentDisposition('attachment', name));
    const z = new ZipWriter((c) => res.write(c));
    const drain = () => new Promise((ok) => res.once('drain', ok));
    let total = 0;
    const addFile = async (full, inZip) => {
      const st = fs.statSync(full);
      if (st.isDirectory()) {
        z.addDir(inZip, { mtime: st.mtime });
        for (const e of fs.readdirSync(full)) if (!/^~\$|^\.~lock\./.test(e)) await addFile(path.join(full, e), `${inZip}/${e}`);
        return;
      }
      total += st.size;
      if (total > 3.5 * 1024 * 1048576) throw err(413, 'The selection is too large for a zip download (more than 3.5 GB).');
      if (!z.add(inZip, await fs.promises.readFile(full), { mtime: st.mtime })) await drain();
    };
    try {
      for (const l of items) await addFile(l.full, path.basename(l.full));
      z.end(); res.end();
    } catch (e) { if (!res.headersSent) throw e; res.destroy(e); }
  });

  /* ---------------- desktop ---------------- */
  const desk = (fn) => async (req, res) => {
    const l = lib.loc(b(req));
    if (!fs.existsSync(l.full)) throw err(404, 'File not found');
    if (!o.hasDesktop()) throw err(409, 'This Hostwerk runs without a screen – open the file on the computer itself.');
    if (fn === desktop.open) { const no = desktop.refuse(l.full); if (no) throw err(no.status, no.error, null, { hint: host.t('Use “Open” in Office – everything stays inside Hostwerk.') }); }
    await fn(l.full);
    res.json({ ok: true });
  };
  r.post('/office/open-desktop', desk(desktop.open));
  r.post('/office/reveal', desk(desktop.reveal));

  /* ---------------- locks ---------------- */
  // soft lock: { ok, lock: { by, since, other } } – 409 { error, lock } when someone else holds it (force: take it over).
  // body: { root?, path, client?, sid?, seq? } (EditorHost/ui/src/lib/locks.js). The same user (or the same window)
  // always takes a lock over silently – a stale lock never blocks the one who left it.
  const lockBody = (req) => { const x = req.body || {}; return { client: x.client ? String(x.client) : null, sid: x.sid ? String(x.sid) : null, seq: Number(x.seq) }; };
  // the marks in Files of every window follow at once (files.jsx listens) – not on a mere heartbeat
  const lockChanged = (l, locked) => { try { host.broadcast({ type: 'office-lock', key: l.key, root: l.root || null, path: l.rel, locked }); } catch { /* egal */ } };
  r.post('/office/lock', (req, res) => {
    const l = lib.loc(b(req)); const cur = o.meta.lockOf(l.key); const by = String((req.body || {}).by || o.userName()).slice(0, 80);
    const x = lockBody(req);
    if (cur && cur.by !== by && !(x.client && cur.client === x.client) && !(req.body || {}).force) return res.status(409).json({ error: host.t('{name} is editing this file right now.', { name: cur.by }), lock: { by: cur.by, since: cur.since, other: true } });
    const n = o.meta.lock(l.key, by, Date.now(), x);
    if (!n.ignored && !n.renewed) lockChanged(l, true);
    res.json({ ok: true, lock: { by: n.by, since: n.since, other: false, ...(n.ignored ? { ignored: true } : {}) } });
  });
  r.post('/office/unlock', (req, res) => { const l = lib.loc(b(req)); const gone = o.meta.unlock(l.key, lockBody(req)); if (gone) lockChanged(l, false); res.json({ ok: true, released: gone }); });

  /* ---------------- shares ---------------- */
  // url: absolute – remote access (own domain first) if it runs, else the address this request came in on
  const shareUrls = (token, req) => [...new Set([...o.linkBases(), (o.settings().integrations.hostUrl && !/host\.docker\.internal/.test(o.settings().integrations.hostUrl)) ? o.settings().integrations.hostUrl : `${req.protocol}://${req.get('host')}`])].map((x) => `${x.replace(/\/+$/, '')}/api/office/pub/${token}`);
  const shareOut = (s, req) => { const urls = shareUrls(s.token, req); return { ...S.view(s), url: urls[0], urls }; };
  r.get('/office/shares', (req, res) => res.json({ shares: o.shareList().map((s) => shareOut(s, req)) }));
  r.post('/office/shares', (req, res) => {
    const s = o.settings().sharing;
    if (!s.enabled) throw err(409, 'Sharing is switched off in the Office settings.');
    const f = lib.file(b(req));
    const body = req.body || {};
    const list = o.shareList();
    if (list.length >= 500) throw err(409, 'At most 500 share links – remove old ones first.');
    if (body.password && String(body.password).length < 4) throw err(400, 'The password needs at least 4 characters.');
    const sh = { id: crypto.randomBytes(6).toString('hex'), token: S.newToken(), key: f.key, root: f.root, path: f.rel, name: path.basename(f.full), created: Date.now(),
      expires: S.expiryOf(body.expires !== undefined ? body.expires : body.days, s.defaultDays), hash: body.password ? S.hashPassword(body.password) : null,
      allowDownload: body.allowDownload !== undefined ? !!body.allowDownload : s.allowDownload, views: 0 };
    o.saveShares([...list, sh]);
    res.json(shareOut(sh, req));
  });
  r.delete('/office/shares/:id', (req, res) => { o.saveShares(o.shareList().filter((s) => s.id !== req.params.id)); res.json({ ok: true }); });

  /* ---------------- assets ---------------- */
  r.get('/office/assets/:file', (req, res) => {
    const n = String(req.params.file || '');
    if (!/^[a-z0-9_-]+\.(js|css)$/.test(n)) throw err(404, 'Not found');
    const p = path.join(host.dir, 'ui', 'assets', n);
    if (!fs.existsSync(p)) throw err(404, 'Not found');
    res.setHeader('Content-Type', n.endsWith('.js') ? 'text/javascript; charset=utf-8' : 'text/css; charset=utf-8');
    res.setHeader('Cache-Control', req.query.v ? 'public, max-age=31536000, immutable' : 'no-cache');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    fs.createReadStream(p).pipe(res);
  });

  /* ---------------- private Office engine ---------------- */
  /** Engine state for the shell: { state: missing|installing|updating|ready|system|error, version, size, path, progress: { percent, text }, error, … } */
  const engineView = async (check) => {
    const st = o.engine.status();
    const c = await o.converters();
    let state = st.installing ? (st.installed ? 'updating' : 'installing') : st.installed ? 'ready' : st.error ? 'error' : (c.lo && c.lo.source !== 'engine') || (c.ms && c.msEnabled) ? 'system' : 'missing';
    const out = { state, version: st.version, size: st.size, sizeHint: st.sizeHint, path: st.path, progress: { percent: st.progress, text: st.text }, error: st.error, jobId: st.jobId, latest: st.latest, updateAvailable: false, supported: st.supported, autoEngine: o.settings().converters.autoEngine,
      system: { libreoffice: c.lo && c.lo.source !== 'engine' ? { path: c.lo.path, version: c.lo.version } : null, msoffice: c.msEnabled ? c.ms : null } };
    if (check) { try { const l = await o.engine.latest({ fresh: true }); out.latest = l.version; } catch (e) { out.checkError = e.message; } }
    out.updateAvailable = !!(out.version && out.latest && require('./engine').cmpVersion(out.latest, out.version) > 0);
    if (state === 'error' && st.installed) out.state = 'ready';
    return out;
  };
  r.get('/office/engine', async (req, res) => res.json(await engineView(boolQ(req.query.check))));
  r.post('/office/engine/install', async (req, res) => { o.engine.install(); res.json(await engineView(false)); });
  r.post('/office/engine/update', async (req, res) => { o.engine.install({ update: true }); res.json(await engineView(false)); });
  r.delete('/office/engine', async (req, res) => { o.engine.remove(); o.resetConverters(); res.json(await engineView(false)); });

  /* ---------------- full editor (ONLYOFFICE Document Server) ---------------- */
  // app is optional – 'onlyoffice' is the only full editor; anything else is refused
  const appOf = (v) => { const a = String(v || 'onlyoffice'); if (a !== 'onlyoffice') throw err(400, 'Unknown office server'); return a; };
  r.get('/office/integrations/state', async (req, res) => res.json(await o.servers.state()));
  r.post('/office/integrations/setup', (req, res) => { const j = o.servers.setup(appOf((req.body || {}).app)); res.json({ ok: true, jobId: j.id }); });
  r.post('/office/integrations/start', async (req, res) => res.json(await o.servers.action(appOf((req.body || {}).app), 'start')));
  r.post('/office/integrations/stop', async (req, res) => res.json(await o.servers.action(appOf((req.body || {}).app), 'stop')));
  r.delete('/office/integrations/:app', async (req, res) => res.json(await o.servers.remove(appOf(req.params.app))));
  async function makeSession(req, app, where, lang) {
    const f = lib.file(where);
    const s = o.settings().integrations;
    const conf = s[app] || {};
    if (!conf.url) throw err(409, 'This office server is not set up yet – set it up under Office → Settings → Integrations.');
    const base = (s.hostUrl || `${req.protocol}://${req.get('host')}`).replace(/\/+$/, '');
    const id = W.fileIdOf(f.root, f.rel);
    const ext = extOf(f.full);
    const ttl = 10 * 3600000;
    const token = W.signToken(o.secret(), { k: f.key, r: f.root, p: f.rel, w: true, u: o.userName() }, ttl);
    const q2 = `access_token=${encodeURIComponent(token)}`;
    const config = {
      document: { fileType: ext, key: W.ooKey(f.key, f.st.mtimeMs), title: path.basename(f.full), url: `${base}/api/office/wopi/oo/${id}/file?${q2}`, permissions: { edit: true, download: true, print: true, comment: true, review: true } },
      documentType: W.ooType(ext),
      editorConfig: { callbackUrl: `${base}/api/office/wopi/oo/${id}/callback?${q2}`, lang: lang, mode: 'edit', user: { id: 'hostwerk', name: o.userName() }, customization: { autosave: true, forcesave: true, compactHeader: true } },
    };
    if (conf.secret) config.token = W.jwtSign(config, conf.secret);
    return withFrame({ app, apiUrl: `${conf.url.replace(/\/+$/, '')}/web-apps/apps/api/documents/api.js`, config }, conf.url);
  }
  r.post('/office/integrations/session', async (req, res) => res.json(await makeSession(req, appOf((req.body || {}).app), b(req), host.lang())));
  /*
   * Embedding page: Hostwerk's window CSP blocks scripts from other origins, so the shell iframes this page.
   * It carries its OWN CSP that allows exactly the configured Document Server; the script is a separate file (no inline JS).
   */
  function withFrame(payload, serverUrl) {
    const now = Date.now();
    for (const [k, v] of o.frames) if (v.until < now) o.frames.delete(k);
    const sid = crypto.randomBytes(18).toString('base64url');
    o.frames.set(sid, { payload, origin: new URL(serverUrl).origin, until: now + 10 * 60000 });
    return { ...payload, frame: `/api/office/integrations/frame?session=${sid}` };
  }
  r.get('/office/integrations/frame', async (req, res) => {
    // ?session=<id> from POST …/session, or directly ?app=&path=[&root=][&lang=]
    let sid = String(req.query.session || '');
    if (!sid && req.query.app) sid = new URL((await makeSession(req, appOf(req.query.app), q(req), /^(de|en)$/.test(String(req.query.lang)) ? String(req.query.lang) : host.lang())).frame, 'http://x').searchParams.get('session');
    const f = o.frames.get(sid);
    if (!f || f.until < Date.now()) throw err(404, 'This editor session has expired – open the file again.');
    const src = f.origin;
    res.setHeader('Content-Security-Policy', [`default-src 'none'`, `script-src 'self' ${src}`, `style-src 'self' 'unsafe-inline' ${src}`, `img-src 'self' data: blob: ${src}`, `font-src 'self' data: ${src}`, `connect-src 'self' ${src}`, `frame-src ${src}`, `form-action ${src}`, "frame-ancestors 'self'", "base-uri 'none'"].join('; '));
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('Referrer-Policy', 'no-referrer');
    const data = JSON.stringify(f.payload).replace(/</g, '\\u003c');
    res.type('html').send(`<!doctype html><html><head><meta charset="utf-8"><title>Office</title><style>html,body,#editor{margin:0;height:100%;width:100%;overflow:hidden;background:#fff}iframe{border:0;width:100%;height:100%}</style></head><body><div id="editor"></div><script type="application/json" id="cfg">${data}</script><script src="/api/office/integrations/frame.js"></script></body></html>`);
  });
  r.get('/office/integrations/frame.js', (req, res) => {
    res.setHeader('Content-Type', 'text/javascript; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache');
    res.send(FRAME_JS);
  });

  /* ---------------- Google Drive ---------------- */
  const gstate = () => { const st = o.drive.state(); return { ...st, guide: st.guide.map((x) => host.t(x)) }; };
  r.get('/office/google/state', (req, res) => res.json(gstate()));
  r.post('/office/google/connect', async (req, res) => { const r0 = await o.drive.connect(); res.json({ ...gstate(), ...(r0 && r0.url ? { url: r0.url } : {}) }); });
  r.post('/office/google/disconnect', async (req, res) => { await o.drive.disconnect(); res.json(gstate()); });
  r.get('/office/google/files', async (req, res) => res.json(await o.drive.list({ folder: String(req.query.folder || 'root'), q: String(req.query.q || '').slice(0, 200), view: String(req.query.view || ''), pageToken: String(req.query.page || '') })));
  r.get('/office/google/meta', async (req, res) => res.json(await o.drive.meta(String(req.query.id || ''))));
  r.get('/office/google/raw', async (req, res) => {
    const f = await o.drive.read(String(req.query.id || ''), req.query.as ? String(req.query.as) : '');
    res.setHeader('Content-Type', f.mime); res.setHeader('Cache-Control', 'no-store');
    res.setHeader('x-name', encodeURIComponent(f.name)); res.setHeader('x-ext', f.ext); res.setHeader('x-mtime', String(f.meta.mtime || ''));
    res.setHeader('Access-Control-Expose-Headers', 'x-name, x-ext, x-mtime');
    if (boolQ(req.query.download)) res.setHeader('Content-Disposition', contentDisposition('attachment', f.name));
    res.end(f.data);
  });
  r.put('/office/google/raw', raw, async (req, res) => res.json(await o.drive.write(String(req.query.id || ''), Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0), String(req.query.ext || ''))));
  r.post('/office/google/import', async (req, res) => { const body = req.body || {}; const f = await o.drive.read(String(body.id || ''), body.as ? String(body.as) : ''); res.json(lib.create(body.dir || '', f.name, f.data)); });
  r.get('/office/google/resolve', (req, res) => {
    const f = lib.file(q(req));
    if (!/^g(doc|sheet|slides|draw|form|site|map|jam|link)$/.test(extOf(f.full))) throw err(400, 'This is not a Google file.');
    const id = G.parsePointer(fs.readFileSync(f.full, 'utf8').slice(0, 20000));
    if (!id) throw err(422, 'This Google file does not contain a document id.');
    res.json({ id, ext: extOf(f.full), connected: !!o.drive.tokens() });
  });

  /* ---------------- public: share links ---------------- */
  const guard = o.guard;
  const shareFor = (req, res) => {
    const s = S.find(o.shareList(), req.params.token);
    res.setHeader('X-Robots-Tag', 'noindex'); res.setHeader('Referrer-Policy', 'no-referrer'); res.setHeader('X-Content-Type-Options', 'nosniff');
    if (!s) { guard.fail(req.ip); res.status(404).type('html').send(S.gonePage(S.langOf(req))); return null; }
    let l; try { l = lib.file({ root: s.root, path: s.path }); } catch { res.status(404).type('html').send(S.gonePage(S.langOf(req))); return null; }
    const unlocked = !s.hash || (String(req.headers.cookie || '').split(/;\s*/).some((c) => c === `ofs_${s.id}=${S.unlockValue(o.secret(), s)}`));
    return { s, l, unlocked };
  };
  const modeOf = async (s, l) => {
    const ext = extOf(l.full);
    if (ext === 'pdf') return { mode: 'pdf' };
    if (IMG.test(ext)) return { mode: 'image' };
    const kind = (FORMATS[ext] || {}).kind;
    if (kind && kind !== 'viewer' && o.plan(ext, 'pdf', await o.converters())) return { mode: 'pdf', converted: true };
    if (X.canText(ext) && l.st.size <= 30 * 1048576) { try { const t = await X.text(fs.readFileSync(l.full), ext, o.libs()); if (t) return { mode: 'text', text: t.slice(0, 300000) }; } catch { /* fall through */ } }
    return { mode: 'none' };
  };
  pub.get('/office/pub/:token', async (req, res) => {
    if (guard.locked(req.ip)) return res.status(429).type('text').send('Too many attempts – try again later.');
    const x = shareFor(req, res); if (!x) return;
    res.setHeader('Content-Security-Policy', "default-src 'none'; style-src 'unsafe-inline'; img-src 'self'; frame-src 'self'; form-action 'self'; base-uri 'none'");
    const lang = S.langOf(req);
    if (!x.unlocked) return res.type('html').send(S.passwordPage(lang, x.s, false));
    o.countView(x.s.id);
    const m = await modeOf(x.s, x.l);
    res.type('html').send(S.page({ lang, share: x.s, base: req.originalUrl.split('?')[0].replace(/\/+$/, ''), size: x.l.st.size, ...m }));
  });
  pub.post('/office/pub/:token', form, (req, res) => {
    if (guard.locked(req.ip)) return res.status(429).type('text').send('Too many attempts – try again later.');
    const x = shareFor(req, res); if (!x) return;
    if (x.s.hash && !S.checkPassword(x.s.hash, (req.body || {}).password)) { guard.fail(req.ip); return res.status(403).type('html').send(S.passwordPage(S.langOf(req), x.s, true)); }
    const p = req.originalUrl.split('?')[0].replace(/\/+$/, '');
    res.setHeader('Set-Cookie', `ofs_${x.s.id}=${S.unlockValue(o.secret(), x.s)}; Path=${p}; HttpOnly; SameSite=Lax; Max-Age=86400`);
    res.redirect(303, p);
  });
  const sendFile = (res, full, name, type, disp) => { res.setHeader('Content-Type', type); res.setHeader('Content-Disposition', contentDisposition(disp, name)); res.setHeader('Cache-Control', 'private, no-store'); fs.createReadStream(full).pipe(res); };
  pub.get('/office/pub/:token/file', (req, res) => {
    const x = shareFor(req, res); if (!x) return;
    if (!x.unlocked) return res.status(403).type('text').send('Locked');
    const ext = extOf(x.l.full);
    if (!(ext === 'pdf' || IMG.test(ext))) return res.status(404).type('text').send('Not found');
    res.setHeader('X-Frame-Options', 'SAMEORIGIN');
    sendFile(res, x.l.full, x.s.name, mimeOf(x.l.full), 'inline');
  });
  pub.get('/office/pub/:token/pdf', async (req, res) => {
    const x = shareFor(req, res); if (!x) return;
    if (!x.unlocked) return res.status(403).type('text').send('Locked');
    try {
      const file = await o.sharePdf(x.l);
      res.setHeader('X-Frame-Options', 'SAMEORIGIN');
      sendFile(res, file, `${x.s.name.replace(/\.[^.]+$/, '')}.pdf`, 'application/pdf', 'inline');
    } catch { res.status(503).type('text').send('The preview is not available right now.'); }
  });
  pub.get('/office/pub/:token/download', (req, res) => {
    const x = shareFor(req, res); if (!x) return;
    if (!x.unlocked) return res.status(403).type('text').send('Locked');
    if (!x.s.allowDownload) return res.status(403).type('text').send('Download not allowed');
    sendFile(res, x.l.full, x.s.name, 'application/octet-stream', 'attachment');
  });
  // Google sign-in lands here only if the in-app window did not catch the redirect itself
  // Loopback redirect of the Google sign-in: normally the in-app window catches it; this finishes the sign-in otherwise
  pub.get('/office/google/callback', async (req, res) => {
    const page = (t) => res.type('html').send(`<!doctype html><meta charset="utf-8"><title>Google</title><body style="font-family:system-ui;padding:24px;color-scheme:light dark">${S.esc(t)}</body>`);
    if (!o.drive.pending) return page(host.t('You can close this window.'));
    try { await o.drive.finish(`http://127.0.0.1${req.originalUrl}`); try { host.broadcast({ type: 'office' }); } catch { /* egal */ } page(host.t('Google Drive is connected – you can close this window.')); } catch (e) { res.status(400); page(e.vars ? host.t(e.message, e.vars) : host.t(e.message)); }
  });

  /* ---------------- public: ONLYOFFICE (token = signed file access, see onlyoffice.js) ---------------- */
  const tokenAuth = (req) => {
    const t = W.verifyToken(o.secret(), req.query.access_token || String(req.headers.authorization || '').replace(/^Bearer\s+/i, ''), { fileId: req.params.id });
    if (!t) throw Object.assign(new Error('Invalid or expired token'), { status: 401 });
    const l = lib.file({ root: t.r, path: t.p });
    return { t, l };
  };
  pub.get('/office/wopi/oo/:id/file', (req, res) => {
    const { l } = tokenAuth(req);
    res.setHeader('Content-Type', mimeOf(l.full));
    fs.createReadStream(l.full).pipe(res);
  });
  pub.post('/office/wopi/oo/:id/callback', async (req, res) => {
    let a;
    try { a = tokenAuth(req); } catch { return res.json({ error: 1 }); }
    const conf = o.settings().integrations.onlyoffice;
    let body = req.body || {};
    let verified = false;
    if (conf.secret) {
      const tok = body.token || String(req.headers.authorization || '').replace(/^Bearer\s+/i, '');
      const p = W.jwtVerify(tok, conf.secret);
      if (!p) return res.json({ error: 1 });
      body = p.payload && typeof p.payload === 'object' ? p.payload : p.status !== undefined ? p : body;
      verified = true;
    }
    const status = Number(body.status);
    if ((status === 2 || status === 6) && body.url) {
      let u; try { u = new URL(String(body.url)); } catch { return res.json({ error: 1 }); }
      // without a JWT only files from the configured Document Server are taken
      if (!verified) { let cu; try { cu = new URL(conf.url); } catch { return res.json({ error: 1 }); } if (u.host !== cu.host) return res.json({ error: 1 }); }
      try {
        const got = await W.fetchUrl(u.href, { timeout: 120000, max: 200 * 1048576 });
        if (got.status !== 200) return res.json({ error: 1 });
        lib.write({ root: a.l.root, path: a.l.rel }, got.body, { force: true });
      } catch (e) { host.log('onlyoffice save failed', e.message); return res.json({ error: 1 }); }
    }
    res.json({ error: 0 });
  });

  /* ---------------- Office 1.1 (routes11.js) – before the error handler ---------------- */
  if (typeof o.extraRoutes === 'function') o.extraRoutes(host, o);

  /* ---------------- errors (status + translated text + extra fields like { mtime }) ---------------- */
  const onError = (e, req, res, next) => {
    if (res.headersSent || !e || !e.status) return next(e);
    const msg = e.vars ? host.t(e.raw || e.message, e.vars) : host.t(e.message);
    res.status(e.status).json({ error: msg, ...(e.extra || {}) });
  };
  r.use(onError);
  pub.use((e, req, res, next) => { if (res.headersSent || !e || !e.status) return next(e); if (/\/office\/(wopi|pub)\//.test(req.originalUrl)) return res.status(e.status).json({ error: e.message }); return next(e); });
}

module.exports = { routes, contentDisposition, FRAME_JS };
