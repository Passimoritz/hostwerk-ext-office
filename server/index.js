'use strict';
/*
 * Office – server part (native Hostwerk extension, SDK 3.x). Contract: ARCHITECTURE.md.
 * Data (never inside the user's library): <Hostwerk data>/office-data/
 *   settings.json · meta.json (favorites, tags, colors, notes, recent) · shares.json · index.json (full text)
 *   versions/ · trash/ · templates/ · cache/ (share PDFs) · backups/ · google.json (sealed tokens) · secret.key
 * Library default: <Hostwerk data>/office · private Office engine: <Hostwerk data>/office-engine
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const C = require('./convert');
const SET = require('./settings');
const { Meta } = require('./meta');
const { Versions } = require('./versions');
const { Trash } = require('./trash');
const { Library, Templates } = require('./library');
const { Index } = require('./search');
const { Engine } = require('./engine');
const { Servers } = require('./servers');
const { Drive } = require('./google');
const X = require('./extract');
const { readJson, writeJson } = require('./store');
const { ZipWriter, readZip } = require('./zip');
const { routes } = require('./routes');
const { extOf, FORMATS } = require('./formats');
const { OsReg } = require('./osreg');
const V11 = require('./v11');
const { routes11 } = require('./routes11');

let state = null;
// per activation (key: its host): an update activates the new module while the old one is still shutting down – the old
// deactivate() must never stop the new activation (it removed the new "Office documents" root in Files after an update)
const states = new WeakMap();

function activate(host) {
  const { ctx } = host;
  const dataRoot = ctx.paths.data;
  const dataDir = path.join(dataRoot, 'office-data');
  fs.mkdirSync(dataDir, { recursive: true });
  const defaultLibrary = path.join(dataRoot, 'office');
  const file = (n) => path.join(dataDir, n);
  const timers = [];
  const every = (ms, fn) => { const t = setInterval(() => { try { const r = fn(); if (r && r.catch) r.catch(() => {}); } catch { /* next time */ } }, ms); if (t.unref) t.unref(); timers.push(() => clearInterval(t)); };
  const later = (ms, fn) => { const t = setTimeout(() => { try { const r = fn(); if (r && r.catch) r.catch(() => {}); } catch { /* egal */ } }, ms); if (t.unref) t.unref(); timers.push(() => clearTimeout(t)); };
  const core = (n) => { try { return host.core(n); } catch { return null; } };
  const libOpt = (n) => { try { return host.lib(n); } catch { return null; } };

  /* ---------- settings ---------- */
  let settingsCache = null;
  const settings = () => settingsCache || (settingsCache = SET.normalize(readJson(file('settings.json'), {}), defaultLibrary));
  const saveSettings = (patch) => {
    const before = settings();
    const { settings: next, errors } = SET.merge(before, patch, defaultLibrary);
    if (errors.length) throw host.error(400, 'Invalid setting: {keys}', { keys: errors.join(', ') });
    fs.mkdirSync(next.libraryDir, { recursive: true });
    writeJson(file('settings.json'), next);
    settingsCache = next;
    if (next.libraryDir !== before.libraryDir) { watch(); index.soon(500); addRoot(); }
    if (JSON.stringify(next.converters) !== JSON.stringify(before.converters)) resetConverters();
    // switched on: register right away (also when this version was done before); off: remove
    if (next.integrations.osOpenWith !== before.integrations.osOpenWith) (next.integrations.osOpenWith === false ? osreg.unregister({ byUser: false }) : osreg.supported() ? osreg.register({ notice: false }) : Promise.resolve()).catch((e) => { if (e.code !== 'EDEVBUILD') host.log('open with', e.message); });
    return next;
  };

  /* ---------- stores ---------- */
  const meta = new Meta(file('meta.json'));
  const versions = new Versions(file('versions'));
  const trash = new Trash(file('trash'));
  const templates = new Templates(file('templates'));
  const libs = () => ({ unpdf: libOpt('unpdf'), mailparser: libOpt('mailparser') });
  const index = new Index({ file: file('index.json'), base: () => settings().libraryDir, libs, log: host.log });
  let shares = readJson(file('shares.json'), []);
  const shareList = () => shares;
  const saveShares = (list) => { shares = list; writeJson(file('shares.json'), shares); };
  let viewSave = null;
  const countView = (id) => { const s = shares.find((x) => x.id === id); if (!s) return; s.views = (s.views || 0) + 1; s.lastView = Date.now(); if (!viewSave) { viewSave = setTimeout(() => { viewSave = null; try { writeJson(file('shares.json'), shares); } catch { /* next */ } }, 2000); if (viewSave.unref) viewSave.unref(); } };
  const sharedKeys = () => new Set(shares.filter((s) => !(s.expires && s.expires <= Date.now())).map((s) => s.key));
  const rekeyShares = (a, b) => { let ch = false; for (const s of shares) if (s.key === a || s.key.startsWith(`${a}/`)) { s.key = b + s.key.slice(a.length); if (!s.root) s.path = s.key; s.name = path.posix.basename(s.key); ch = true; } if (ch) saveShares(shares); };

  let hooks = null; // Office 1.1 (v11.js): rules, forms, publishing, co-editing follow writes and renames
  const library = new Library({ settings, files: ctx.files || null, meta, versions, trash, index, sharedKeys, dataDir, onChange: () => index.soon(), onRekey: (a, b) => { rekeyShares(a, b); if (hooks) hooks.onRekey(a, b); }, onWrite: (l) => { if (hooks) hooks.onWrite(l); }, thumbHas: (k, m) => !!(hooks && hooks.thumbs && hooks.thumbs.has(k, m)) });

  /* ---------- secret for tokens/cookies ---------- */
  let secretBuf = null;
  const secret = () => {
    if (secretBuf) return secretBuf;
    const f = file('secret.key');
    try { secretBuf = Buffer.from(fs.readFileSync(f, 'utf8').trim(), 'hex'); } catch { /* new */ }
    if (!secretBuf || secretBuf.length < 32) { secretBuf = crypto.randomBytes(32); fs.writeFileSync(f, secretBuf.toString('hex'), { mode: 0o600 }); }
    return secretBuf;
  };

  /* ---------- converters ---------- */
  const engineDir = path.join(dataRoot, 'office-engine');
  let conv = null;
  const resetConverters = () => { conv = null; };
  const converters = async (fresh = false) => {
    if (!fresh && conv && Date.now() - conv.t < 60000) return conv.v;
    const s = settings().converters;
    const [lo, ms] = await Promise.all([C.detectLibreOffice({ override: s.libreoffice, engineDir }).catch(() => null), C.detectMsOffice().catch(() => null)]);
    const v = { lo, ms, msEnabled: s.msoffice !== false };
    conv = { t: Date.now(), v };
    return v;
  };
  const convCtx = () => ({ preferred: settings().converters.preferred, timeoutSec: settings().converters.timeout, profileDir: file('lo-profile'), scriptDir: file('scripts') });
  const convert = (data, from, to) => C.convert(data, from, to, { ...convCtx(), detect: () => converters() });
  const plan = (from, to, avail) => C.plan(from, to, avail, settings().converters.preferred);

  /* ---------- private Office engine ---------- */
  const dl = core('download');
  const jobs = core('jobs');
  const runJob = jobs && jobs.runJob ? jobs.runJob : (title, fn) => { const j = { id: crypto.randomUUID(), state: 'running', progress: () => {}, status: () => {}, log: () => {} }; j.promise = Promise.resolve().then(() => fn(j)).then((r) => { j.state = 'done'; j.result = r; }, (e) => { j.state = 'error'; j.error = e.message; }); return j; };
  const engine = new Engine({
    dir: engineDir, t: host.t, log: host.log, runJob,
    fetchText: (u) => (dl && dl.fetchText ? dl.fetchText(u, { timeout: 30000 }) : fetch(u).then((r) => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.text(); })),
    downloadFile: (u, dest, opts) => { if (!dl || !dl.downloadFile) throw new Error('Downloads are not available'); return dl.downloadFile(u, dest, opts); },
    findSoffice: (dir) => C.engineFiles(dir).find((p) => { try { return fs.statSync(p).isFile(); } catch { return false; } }) || null,
    probe: async (exe) => { if (!exe) throw new Error(host.t('The Office engine was unpacked but LibreOffice was not found in it.')); const r = await C.convert(Buffer.from('Hostwerk Office engine test'), 'txt', 'pdf', { ...convCtx(), preferred: 'libreoffice', timeoutSec: 300, detect: async () => ({ lo: { path: exe }, ms: null }) }); if (r.data.subarray(0, 4).toString() !== '%PDF') throw new Error(host.t('The Office engine does not work on this PC.')); resetConverters(); },
  });
  if (ctx.patchUpdates && typeof ctx.patchUpdates.register === 'function') { try { const off = ctx.patchUpdates.register('office-engine', engine.patchSource()); timers.push(() => { try { off(); } catch { /* egal */ } }); } catch (e) { host.log('patchUpdates', e.message); } }
  // autoEngine 'auto': set up silently when no converter exists (a minute after start, not during Hostwerk's start)
  later(90000, async () => { const s = settings().converters; if (s.autoEngine !== 'auto' || engine.status().installed) return; const c = await converters(true); if (!c.lo && !(c.ms && c.msEnabled)) engine.install(); });

  /* ---------- full editor (ONLYOFFICE) & Google ---------- */
  const servers = new Servers({ docker: () => ctx.docker || null, runJob, t: host.t, settings, save: saveSettings, hostPort: () => ctx.port || 7780, log: host.log });
  const secrets = core('secrets');
  // in-app web window (SDK 3.14: host.webWindow.open, older cores: ctx.app.openWebWindow) – never an external browser
  const stableWindow = () => { try { return typeof host.supports === 'function' && host.supports('webWindow') && host.webWindow && typeof host.webWindow.open === 'function'; } catch { return false; } };
  const canWindow = () => stableWindow() || !!(ctx.app && ctx.app.webWindow && typeof ctx.app.openWebWindow === 'function');
  const drive = new Drive({
    file: file('google.json'), settings, canWindow,
    seal: (t) => (secrets && secrets.seal ? secrets.seal(t) : t), open: (v) => (secrets && secrets.open ? secrets.open(v) : v),
    openWindow: (w) => (stableWindow() ? host.webWindow.open(w) : ctx.app.openWebWindow(w)),
    redirectUri: () => `http://127.0.0.1:${ctx.port || 7780}/api/office/google/callback`,
  });

  /* ---------- share previews (office → PDF, cached per version) ---------- */
  const cacheDir = file('cache');
  const sharePdf = async (l) => {
    fs.mkdirSync(cacheDir, { recursive: true });
    const f = path.join(cacheDir, `${crypto.createHash('sha1').update(`${l.key}\n${Math.floor(l.st.mtimeMs)}\n${l.st.size}`).digest('hex')}.pdf`);
    if (fs.existsSync(f)) return f;
    const r = await convert(await fs.promises.readFile(l.full), extOf(l.full), 'pdf');
    fs.writeFileSync(f, r.data);
    return f;
  };
  // wrong share passwords: lock an address for 10 minutes after 20 failures
  const fails = new Map();
  const guard = { fail: (ip) => { const e = fails.get(ip) || { n: 0, t: Date.now() }; if (Date.now() - e.t > 600000) { e.n = 0; e.t = Date.now(); } e.n++; fails.set(ip, e); }, locked: (ip) => { const e = fails.get(ip); return !!e && e.n >= 20 && Date.now() - e.t < 600000; } };

  const platform = core('platform');
  const hasDesktop = () => (platform && typeof platform.HAS_DESKTOP === 'boolean' ? platform.HAS_DESKTOP : true);
  const userName = () => settings().general.author || (() => { try { return os.userInfo().username; } catch { return 'Hostwerk'; } })();
  const linkBases = () => { try { const st = ctx.remote && ctx.remote.status(); if (!st || !st.running) return []; return [...(st.domain ? [`https://${st.domain}${st.port === 443 ? '' : `:${st.port}`}`] : []), ...(st.urls || []).map((u) => u.replace(/\/$/, ''))]; } catch { return []; } };

  /* ---------- Office inside Hostwerk's Files (SDK 3.15): the library as a root, "Open with" in the system ---------- */
  // root id 'ext:office:library' – every Hostwerk file route works with it; re-added when libraryDir changes
  const LIB_ROOT = 'library';
  function addRoot() {
    if (!(typeof host.supports === 'function' && host.supports('files.roots') && host.files)) return null;
    try { return host.files.addRoot({ id: LIB_ROOT, label: host.t('Office documents'), path: settings().libraryDir, icon: 'file', create: true }); } catch (e) { host.log('files.addRoot', e.message); return null; }
  }
  const osreg = new OsReg({ host, file: file('os.json'), settings, log: host.log });

  const o = { settings, saveSettings, library, meta, versions, trash, templates, index, libs, converters, resetConverters, convert, plan, targetsFor: C.targets, engine, servers, drive, canWindow,
    frames: new Map(), secret, shareList, saveShares, countView, sharePdf, guard, hasDesktop, userName, linkBases, osreg, libRoot: () => `ext:${host.id || 'office'}:${LIB_ROOT}` };
  hooks = V11.setup(host, o, { dataRoot, dataDir, every, later });
  o.extraRoutes = routes11;
  routes(host, o);

  /* ---------- background work ---------- */
  let watcher = null;
  function watch() {
    try { if (watcher) watcher.close(); } catch { /* egal */ }
    watcher = null;
    try { watcher = fs.watch(settings().libraryDir, { recursive: true }, (ev, name) => { if (name && /(^|[\\/])(~\$|\.~lock\.)/.test(String(name))) return; index.soon(4000); }); watcher.on('error', () => {}); } catch { /* recursive watch not supported – periodic reindex below */ }
  }
  fs.mkdirSync(settings().libraryDir, { recursive: true });
  addRoot();
  // an update from 1.0.x: the old module's deactivate() runs after this activation and removes the root once more
  later(3000, () => addRoot());
  // the root's name is plain text in Hostwerk's language – follow a language switch
  let rootLang = typeof host.lang === 'function' ? host.lang() : null;
  every(20000, () => { const l = typeof host.lang === 'function' ? host.lang() : null; if (l && l !== rootLang) { rootLang = l; addRoot(); } });
  // "Open with": once per install/upgrade, after Hostwerk's start settled (never blocks activate)
  later(4000, () => osreg.sync());
  watch();
  later(5000, () => index.reindex());
  every(30 * 60000, () => index.reindex());
  const maintain = () => {
    const g = settings().general;
    for (const e of trash.purgeOld(g.trashDays)) { meta.drop(`#trash/${e.id}`); versions.drop(`#trash/${e.id}`); }
    versions.pruneAll({ keep: g.keepVersions, days: g.versionDays });
    try { for (const f of fs.readdirSync(cacheDir)) { const p = path.join(cacheDir, f); if (Date.now() - fs.statSync(p).mtimeMs > 7 * 86400000) fs.rmSync(p, { force: true }); } } catch { /* no cache */ }
    snapshot();
  };
  later(120000, maintain);
  every(6 * 3600000, maintain);

  /* ---------- Hostwerk integration ---------- */
  // "New" menu (opens the Office page with this action – the shell reads Hostwerk.on('action'))
  // Only when the manifest has no contributes.new (the integrator may declare them there – then they would appear twice)
  const manifestNew = (() => { try { const m = JSON.parse(fs.readFileSync(path.join(host.dir, 'manifest.json'), 'utf8')); return !!(m.contributes && Array.isArray(m.contributes.new) && m.contributes.new.length); } catch { return false; } })();
  if (!manifestNew && typeof host.setNewEntries === 'function' && ctx.extMarket) {
    try {
      host.setNewEntries([
        { id: 'new-docs', title: { de: 'Neues Dokument', en: 'New document' }, description: { de: 'Word / OpenDocument', en: 'Word / OpenDocument' }, icon: 'file-text' },
        { id: 'new-sheets', title: { de: 'Neue Tabelle', en: 'New spreadsheet' }, description: { de: 'Excel / OpenDocument', en: 'Excel / OpenDocument' }, icon: 'sheet' },
        { id: 'new-slides', title: { de: 'Neue Präsentation', en: 'New presentation' }, description: { de: 'PowerPoint / OpenDocument', en: 'PowerPoint / OpenDocument' }, icon: 'presentation' },
      ]);
    } catch (e) { host.log('setNewEntries', e.message); }
  }
  // Quick search: recently opened office files
  // root ids may contain ":" themselves (own:…, web:…) – so the root travels as its own parameter
  const hrefOf = (r) => `#/ext/office?action=open&target=${encodeURIComponent(r.path)}${r.root ? `&root=${encodeURIComponent(r.root)}` : ''}`;
  if (typeof host.searchSource === 'function') {
    host.searchSource(() => meta.recent(30).map((r) => ({ label: path.posix.basename(r.path), hint: r.root ? `${host.t('Office')} · ${r.root}` : `${host.t('Office')} · ${path.posix.dirname(r.path) === '.' ? host.t('Library') : path.posix.dirname(r.path)}`, icon: 'file-text', href: hrefOf(r) })));
  }
  if (typeof host.contribute === 'function') {
    try { host.contribute('search', { run: (match) => { const qs = typeof match === 'string' ? match : (match && match.q) || ''; return qs && qs.length >= 3 ? index.search(qs, { limit: 8 }).map((h) => ({ type: 'office', label: path.posix.basename(h.rel), hint: h.snippet, href: hrefOf({ path: h.rel }) })) : []; } }); } catch (e) { host.log('contribute search', e.message); }
    try { host.contribute('clickguard', { block: [/^\/office\/(engine|integrations\/setup|trash\/purge|google\/connect)/] }); } catch { /* older Hostwerk */ }
  }
  // AI: read-only tools + @-source
  if (typeof host.aiExtend === 'function') {
    const findFile = (p) => { const hit = library.fromKeys([String(p || '')])[0]; if (hit) return hit; const name = String(p || '').toLowerCase(); const r = index.search(name, { limit: 1 })[0]; if (r) return library.entryOf({ path: r.rel }); throw new Error(`Keine Office-Datei „${p}“ gefunden – office_search hilft.`); };
    const readText = async (p) => { const e = findFile(p); const l = library.file({ path: e.path }); const cached = index.get(e.path); const t = cached && cached.mtime === Math.floor(l.st.mtimeMs) && cached.text ? cached.text : await X.text(await fs.promises.readFile(l.full), e.ext, libs()); return { e, t }; };
    host.aiExtend({
      about: { de: 'Office-Dokumente (Bibliothek mit Word, Excel, PowerPoint, OpenDocument, PDF)', en: 'office documents (library with Word, Excel, PowerPoint, OpenDocument, PDF)' },
      vocab: ['Office', 'Word', 'Excel', 'PowerPoint', 'OpenDocument', 'LibreOffice'],
      groups: { office: { de: 'Office', en: 'Office' } },
      tools({ add, S }) {
        const str = (d) => (typeof S === 'function' ? S(d) : { type: 'string', description: d });
        add('office_list', 'office', 'Office-Dateien der Bibliothek auflisten (Ordner dir, optional Name q, view: all|recent|favorites). Liefert Pfad, Art, Größe, Änderungsdatum.', { dir: str('Ordner in der Bibliothek ("" = oben)'), q: str('Teil des Namens'), view: str('all | recent | favorites') }, [], async ({ dir = '', q = '', view = 'all' }) => {
          const list = view === 'recent' ? library.fromKeys(meta.recent(50)) : view === 'favorites' ? library.fromKeys(meta.favorites()) : Library.filter(library.listDir(dir, { recursive: !!q }), { q });
          return { dateien: list.slice(0, 100).map((e) => ({ pfad: e.path, name: e.name, ordner: e.dir, art: e.kind, groesse: e.size, geaendert: new Date(e.mtime).toISOString() })) };
        }, { risk: 'read', describe: () => ({ ask: 'Office-Dateien auflisten', verb: 'Auflisten' }) });
        add('office_search', 'office', 'Volltextsuche in den Office-Dateien (Word, Excel, PowerPoint, OpenDocument, PDF, RTF, Text). Liefert Pfad und Textausschnitt.', { q: str('Suchbegriffe') }, ['q'], async ({ q }) => ({ treffer: index.search(String(q || ''), { limit: 20 }).map((h) => ({ pfad: h.rel, ausschnitt: h.snippet })) }), { risk: 'read', describe: ({ q }) => ({ ask: `In Office-Dateien nach „${q}“ suchen`, verb: 'Suchen' }) });
        add('office_read', 'office', 'Den Text einer Office-Datei lesen (Pfad aus office_list/office_search). Nur lesen, höchstens 20000 Zeichen ab offset.', { path: str('Pfad in der Bibliothek'), offset: { type: 'number', description: 'Startzeichen (Standard 0)' } }, ['path'], async ({ path: p, offset = 0 }) => { const { e, t } = await readText(p); const start = Math.max(0, Number(offset) || 0); return { pfad: e.path, name: e.name, zeichen: t.length, text: t.slice(start, start + 20000), weiter: start + 20000 < t.length ? start + 20000 : null }; }, { risk: 'read', describe: ({ path: p }) => ({ ask: `Office-Datei „${p}“ lesen`, verb: 'Lesen' }) });
        add('office_info', 'office', 'Eckdaten einer Office-Datei: Seiten, Wörter, Folien, Blätter, Autor, Titel, Datum.', { path: str('Pfad in der Bibliothek') }, ['path'], async ({ path: p }) => { const e = findFile(p); const l = library.file({ path: e.path }); return { ...e, ...(await X.info(await fs.promises.readFile(l.full), e.ext, libs())) }; }, { risk: 'read', describe: ({ path: p }) => ({ ask: `Eckdaten von „${p}“ lesen`, verb: 'Lesen' }) });
      },
      context: {
        kind: 'office', label: { de: 'Office-Dateien', en: 'Office files' },
        search: async (qs) => { const s = String(qs || '').trim(); const list = s ? index.search(s, { limit: 20 }).map((h) => h.rel) : meta.recent(20).filter((r) => !r.root).map((r) => r.path); return list.slice(0, 20).map((rel) => ({ id: rel, label: path.posix.basename(rel) })); },
        resolve: async (id) => { const { e, t } = await readText(id); return `${e.name} (${e.path})\n\n${t.slice(0, 7800)}`; },
      },
    });
  }

  /* ---------- backups (metadata snapshots) & transfer ---------- */
  const backupDir = file('backups');
  const META_FILES = ['settings.json', 'meta.json', 'shares.json'];
  function snapshot(force = false) {
    fs.mkdirSync(backupDir, { recursive: true });
    const list = fs.readdirSync(backupDir).filter((f) => /^office-meta-\d+\.zip$/.test(f)).sort();
    const last = list[list.length - 1];
    if (!force && last && Date.now() - Number(last.match(/(\d+)/)[1]) < 20 * 3600000) return null;
    const z = new ZipWriter();
    for (const f of META_FILES) if (fs.existsSync(file(f))) z.add(f, fs.readFileSync(file(f)));
    const td = file('templates');
    for (const f of fs.existsSync(td) ? fs.readdirSync(td) : []) z.add(`templates/${f}`, fs.readFileSync(path.join(td, f)));
    const name = `office-meta-${Date.now()}.zip`;
    fs.writeFileSync(path.join(backupDir, name), z.end());
    for (const old of list.slice(0, Math.max(0, list.length - 13))) fs.rmSync(path.join(backupDir, old), { force: true });
    return name;
  }
  if (host.backups && typeof host.backups.register === 'function') {
    try {
      host.backups.register({
        label: { de: 'Office (Einstellungen, Favoriten, Freigaben, Vorlagen)', en: 'Office (settings, favorites, shares, templates)' },
        list: () => (fs.existsSync(backupDir) ? fs.readdirSync(backupDir) : []).filter((f) => /^office-meta-\d+\.zip$/.test(f)).map((f) => { const t = Number(f.match(/(\d+)/)[1]); return { id: f, label: `Office · ${new Date(t).toLocaleString(host.lang() === 'de' ? 'de-DE' : 'en-US')}`, t, size: fs.statSync(path.join(backupDir, f)).size, auto: true, svc: 'office', svcName: 'Office', href: '#/ext/office' }; }).sort((a, b) => b.t - a.t),
        restore: async (id) => {
          if (!/^office-meta-\d+\.zip$/.test(String(id))) throw host.error(404, 'Backup not found');
          snapshot(true);
          const z = readZip(fs.readFileSync(path.join(backupDir, id)));
          for (const [n, e] of z) { if (e.dir) continue; if (META_FILES.includes(n) || /^templates\/[\w.-]+$/.test(n)) { fs.mkdirSync(path.dirname(file(n)), { recursive: true }); fs.writeFileSync(file(n), e.read()); } }
          settingsCache = null; meta.data = readJson(file('meta.json'), { files: {}, recent: [] }); shares = readJson(file('shares.json'), []);
          return { ok: true };
        },
        remove: async (id) => { if (/^office-meta-\d+\.zip$/.test(String(id))) fs.rmSync(path.join(backupDir, id), { force: true }); },
        // external backup: the documents themselves + metadata
        offsite: async ({ dir }) => {
          meta.flush();
          const out = fs.createWriteStream(path.join(dir, 'office.zip'));
          const z = new ZipWriter((c) => out.write(c));
          const addTree = (base, prefix) => { if (!fs.existsSync(base)) return; for (const e of fs.readdirSync(base, { withFileTypes: true })) { const p = path.join(base, e.name); if (e.isDirectory()) addTree(p, `${prefix}${e.name}/`); else if (e.isFile()) z.add(`${prefix}${e.name}`, fs.readFileSync(p)); } };
          addTree(settings().libraryDir, 'library/');
          for (const f of META_FILES) if (fs.existsSync(file(f))) z.add(`data/${f}`, fs.readFileSync(file(f)));
          addTree(file('templates'), 'data/templates/');
          z.end();
          await new Promise((ok, bad) => { out.on('error', bad); out.end(ok); });
          return [path.join(dir, 'office.zip')];
        },
      });
    } catch (e) { host.log('backups', e.message); }
  }
  if (host.transfer && typeof host.transfer.register === 'function') {
    try {
      host.transfer.register({
        label: { de: 'Office-Dokumente und Einstellungen', en: 'Office documents and settings' },
        export: async ({ dir }) => {
          meta.flush();
          fs.cpSync(settings().libraryDir, path.join(dir, 'library'), { recursive: true });
          fs.mkdirSync(path.join(dir, 'data'), { recursive: true });
          for (const f of [...META_FILES, 'templates', 'versions']) if (fs.existsSync(file(f))) fs.cpSync(file(f), path.join(dir, 'data', f), { recursive: true });
        },
        import: async ({ dir, test }) => {
          if (!fs.existsSync(path.join(dir, 'library'))) throw new Error('No Office data in this package');
          if (test) return { ok: true };
          fs.cpSync(path.join(dir, 'library'), settings().libraryDir, { recursive: true, force: false, errorOnExist: false });
          for (const f of ['meta.json', 'shares.json', 'templates', 'versions']) { const src = path.join(dir, 'data', f); if (fs.existsSync(src)) fs.cpSync(src, file(f), { recursive: true, force: true }); }
          const st = path.join(dir, 'data', 'settings.json');
          if (fs.existsSync(st)) { const imp = readJson(st, {}); delete imp.libraryDir; delete imp.converters; saveSettings(imp); }
          meta.data = readJson(file('meta.json'), { files: {}, recent: [] }); shares = readJson(file('shares.json'), []);
          index.soon(1000);
          return { ok: true };
        },
      });
    } catch (e) { host.log('transfer', e.message); }
  }

  let down = false;
  const shutdown = () => {
    if (down) return; // the shutdown hook and deactivate() both call it – once is enough
    down = true;
    for (const off of timers.splice(0)) { try { off(); } catch { /* egal */ } }
    try { if (watcher) watcher.close(); } catch { /* egal */ }
    watcher = null;
    if (viewSave) { clearTimeout(viewSave); viewSave = null; try { writeJson(file('shares.json'), shares); } catch { /* egal */ } }
    if (hooks) { try { hooks.shutdown(); } catch { /* egal */ } }
    index.stop(); meta.flush(); C.killAll();
    try { if (host.files && typeof host.files.removeRoot === 'function') host.files.removeRoot(LIB_ROOT); } catch { /* Hostwerk removes it anyway */ }
  };
  host.on('shutdown', shutdown);
  state = { shutdown, o };
  states.set(host, state);
  return { ok: true };
}

function deactivate(host) {
  const s = host && states.has(host) ? states.get(host) : state;
  if (!s) return;
  try { s.shutdown(); } catch { /* egal */ }
  if (state === s) state = null;
}

/**
 * Removing the extension (SDK 3.3 hook, never on switch-off or update): take Hostwerk out of "Open with" (Hostwerk does
 * it itself as well) and drop the root in Files. The documents stay where they are.
 */
async function uninstall(host) {
  try { if (typeof host.supports === 'function' && host.supports('fileTypes') && host.fileTypes) await host.fileTypes.unregister('office'); } catch (e) { if (host.log) host.log('open with', e.message); }
  try { if (host.files && typeof host.files.removeRoot === 'function') host.files.removeRoot('library'); } catch { /* gone anyway */ }
}

module.exports = { activate, deactivate, uninstall, FORMATS };
