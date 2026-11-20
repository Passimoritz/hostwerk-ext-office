'use strict';
/*
 * Office 1.1 wiring: builds the 1.1 services, puts them on `o` (routes11.js uses them), registers notifications,
 * the web-page kind (host.sites), the library in the backup center, the websockets and the background timers.
 *   setup(host, o, { dataRoot, dataDir, every, later }) → { onWrite(l), onRekey(a, b), shutdown() }
 */
const fs = require('fs');
const path = require('path');
const { DataSources } = require('./datasources');
const { DbSheets } = require('./dbsheets');
const { Fields } = require('./fields');
const F = require('./forms');
const { Rules, fill } = require('./rules');
const { Notifier } = require('./notify');
const { Collab } = require('./collab');
const { Remote } = require('./remote');
const { Components } = require('./components');
const { Semantic } = require('./semantic');
const { Thumbs, Publish, Reports, LibBackup } = require('./extras');
const S = require('./sheet');
const SH = require('./shares');
const X = require('./extract');
const { extOf } = require('./formats');

function setup(host, o, { dataRoot, dataDir, every, later }) {
  const ctx = host.ctx;
  const lib = o.library;
  const file = (n) => path.join(dataDir, n);
  o.dataFile = file;
  const offs = [];
  const libOpt = (n) => { try { return host.lib(n); } catch { return null; } };
  const coreOpt = (n) => { try { return host.core(n); } catch { return null; } };

  /* ---------- names, urls, rights ---------- */
  o.ownerName = () => { const g = (ctx.settings && ctx.settings.get().general) || {}; const you = coreOpt('you'); let n = null; try { n = you && you.nameFor ? you.nameFor(g, 'extensions') : null; } catch { /* egal */ } return n || o.settings().general.author || null; };
  const session = (req) => { try { return ctx.remote && typeof ctx.remote.isLoopback === 'function' && !ctx.remote.isLoopback(req) ? ctx.remote.session(req) : null; } catch { return null; } };
  o.userOf = (req) => { const s = session(req); return (s && s.user) || o.ownerName() || host.t('Owner'); };
  /** team members (remote access) only get what Hostwerk's API would allow them */
  o.allowed = (req, method, p) => { const s = session(req); if (!s || !ctx.users || typeof ctx.users.allowed !== 'function') return true; try { return !!ctx.users.allowed({ method, path: p, url: `/api${p}`, originalUrl: `/api${p}`, headers: req.headers, ip: req.socket && req.socket.remoteAddress }, s); } catch { return false; } };
  o.pubUrls = (sub, req) => { const s = o.settings().integrations; const own = s.hostUrl && !/host\.docker\.internal/.test(s.hostUrl) ? s.hostUrl : req ? `${req.protocol}://${req.get('host')}` : `http://127.0.0.1:${ctx.port || 7780}`; return [...new Set([...o.linkBases(), own])].map((x) => `${x.replace(/\/+$/, '')}/api/office/pub/${sub}`); };
  o.pubUrl = (sub, req) => o.pubUrls(sub, req)[0];
  o.lanHint = (req) => { const urls = o.pubUrls('x', req); return { reachable: urls.some((u) => !/^https?:\/\/(localhost|127\.|\[::1\])/.test(u)), remote: o.linkBases().length > 0 }; };
  o.expiryOf = (v) => SH.expiryOf(v === undefined ? null : v, 0);
  o.hashPassword = SH.hashPassword; o.checkPassword = SH.checkPassword; o.gonePage = SH.gonePage;

  /* ---------- notifications ---------- */
  const reg = (id, de, en) => { try { host.notifications.register(id, { de, en }, { on: true }); } catch (e) { host.log('notifications', e.message); } };
  reg('rule', 'Office: eine Regel ist erfüllt', 'Office: a rule is met');
  reg('form', 'Office: neue Formular-Antwort', 'Office: new form answer');
  reg('report', 'Office: geplanter Bericht', 'Office: scheduled report');
  o.notifier = new Notifier({ host });

  /* ---------- 1–3: data, databases, fields ---------- */
  o.data = new DataSources({ get: (p) => host.get(p), t: host.t, ctx });
  try { const ev = coreOpt('events'); if (ev && ev.bus) offs.push(o.data.watch(ev.bus, (m) => host.broadcast(m))); } catch (e) { host.log('data watch', e.message); }
  o.dbs = new DbSheets(ctx.db || null, { lib: libOpt('mongodb') });
  const readLoc = async (loc) => { const l = lib.file(loc); if (l.st.size > 100 * 1048576) throw Object.assign(new Error('The file is too large.'), { status: 413 }); return { buf: await fs.promises.readFile(l.full), ext: extOf(l.full), mtime: Math.floor(l.st.mtimeMs), key: l.key }; };
  o.fields = new Fields({ t: host.t, data: o.data, userName: o.ownerName, readFile: readLoc, info: (buf, ext) => X.info(buf, ext, o.libs()) });

  /* ---------- 4: forms ---------- */
  o.forms = new F.Forms({ file: file('forms.json') });
  const busy = (key) => !!o.meta.lockOf(key) || (() => { const r = o.collab.rooms.get(key); return !!(r && r.peers.size); })();
  o.appendNow = (f, values) => {
    const l = lib.file({ root: f.root, path: f.path });
    const out = S.appendRow(fs.readFileSync(l.full), extOf(l.full), { sheet: f.sheet, values });
    lib.write({ root: f.root, path: f.path }, out.data, { mtime: Math.floor(l.st.mtimeMs) });
    return out.row;
  };
  o.formAnswer = async (f, row) => {
    if (busy(f.key)) {
      if ((f.queue || []).length >= 5000) throw new Error('queue full');
      o.forms.update(f.id, { queue: [...(f.queue || []), { t: Date.now(), values: row }] });
      try { host.broadcast({ type: 'office-form', id: f.id, path: f.path, root: f.root || null, pending: f.queue.length }); } catch { /* egal */ }
    } else o.appendNow(f, row);
    o.forms.update(f.id, { responses: (f.responses || 0) + 1, lastResponse: Date.now() });
    if (f.notify) o.notifier.send(f.notify, { event: 'form', title: host.t('New answer: {name}', { name: f.title }), text: f.columns.map((c, i) => `${c.label}: ${S.dateText(row[i + (f.timestamp ? 1 : 0)])}`).join('\n').slice(0, 1500), url: '/ext/office' }).catch((e) => host.log('form notify', e.message));
  };
  const flushForms = () => {
    for (const f of o.forms.list()) {
      if (!(f.queue || []).length || busy(f.key)) continue;
      let done = 0;
      try { for (const q of f.queue) { o.appendNow(f, q.values); done++; } } catch (e) { host.log('form queue', f.id, e.message); }
      o.forms.update(f.id, { queue: f.queue.slice(done) });
    }
  };
  every(30000, flushForms);

  /* ---------- 5: rules ---------- */
  o.rules = new Rules({
    file: file('rules.json'), read: readLoc, log: host.log,
    stat: async (loc) => { const l = lib.file(loc); return { mtime: Math.floor(l.st.mtimeMs) }; },
    notify: (rule, vars) => o.notifier.send(rule.then.notify.channel, { event: 'rule', title: host.t('Office rule: {file}', { file: vars.file }), text: fill(rule.then.notify.text || host.t('{cell} is now {value}.'), vars), kind: 'warn', url: '/ext/office' }),
  });
  every(120000, () => o.rules.checkAll());

  /* ---------- 6–7: co-editing, remote control (websockets) ---------- */
  o.collab = new Collab();
  o.remote = new Remote();
  every(60000, () => { o.collab.sweep(); o.remote.sweep(); });
  const WS = libOpt('ws');
  const wss = WS ? new (WS.WebSocketServer || WS.Server)({ noServer: true, maxPayload: 16 * 1048576, perMessageDeflate: false, clientTracking: false }) : null;
  const refuse = (socket, code, text) => { try { socket.write(`HTTP/1.1 ${code} ${text}\r\nConnection: close\r\n\r\n`); } catch { /* egal */ } socket.destroy(); };
  const alive = (ws) => { let ok = true; ws.on('pong', () => { ok = true; }); const t = setInterval(() => { if (!ok) { ws.terminate(); return; } ok = false; try { ws.ping(); } catch { /* egal */ } }, 30000); if (t.unref) t.unref(); ws.on('close', () => clearInterval(t)); };
  if (wss) {
    host.on('upgrade', (req, socket, head) => {
      let u; try { u = new URL(req.url, 'http://x'); } catch { return false; }
      const m = /^\/(?:ws|api)\/office\/(collab|remote)\/?$/.exec(u.pathname);
      if (!m) return false;
      if (m[1] === 'collab') {
        let l; try { l = lib.file({ root: u.searchParams.get('root') || null, path: u.searchParams.get('path') || '' }); } catch { refuse(socket, 404, 'Not Found'); return true; }
        if (!o.allowed(req, 'PUT', '/office/raw')) { refuse(socket, 403, 'Forbidden'); return true; }
        const user = o.userOf(req);
        wss.handleUpgrade(req, socket, head, (ws) => {
          let conn;
          try { conn = o.collab.connect(l.key, { send: (b) => { if (ws.readyState === 1) ws.send(b); }, close: (c, rsn) => ws.close(c, rsn) }, { user }); } catch (e) { ws.close(1013, String(e.message).slice(0, 100)); return; }
          alive(ws);
          ws.on('message', (data) => conn.message(Buffer.isBuffer(data) ? data : Buffer.concat([].concat(data))));
          ws.on('close', () => conn.close());
          ws.on('error', () => {});
        });
        return true;
      }
      const token = u.searchParams.get('token') || '';
      if (!o.remote.get(token) || !o.allowed(req, 'POST', '/office/remote')) { refuse(socket, 404, 'Not Found'); return true; }
      if (u.searchParams.get('role') === 'phone') {
        wss.handleUpgrade(req, socket, head, (ws) => {
          const sendj = (m0) => { if (ws.readyState === 1) ws.send(JSON.stringify(m0)); };
          let off; try { off = o.remote.phone(token, { send: sendj, end: (m0) => { sendj(m0); ws.close(1000, 'ended'); } }); } catch (e) { ws.close(1013, 'busy'); return; }
          alive(ws);
          ws.on('message', (d) => { let m0; try { m0 = JSON.parse(String(d).slice(0, 4096)); } catch { return; } try { o.remote.command(token, m0); } catch (e) { sendj({ type: 'error', error: host.t(e.message) }); } });
          ws.on('close', () => { if (off) off(); });
          ws.on('error', () => {});
        });
        return true;
      }
      wss.handleUpgrade(req, socket, head, (ws) => {
        const pr = o.remote.attach(token, { send: (m0) => { if (ws.readyState === 1) ws.send(JSON.stringify(m0)); }, close: () => ws.close(4000, 'replaced') });
        if (!pr) { ws.close(4404, 'gone'); return; }
        alive(ws);
        ws.on('message', (d) => { const s = String(d); if (s.length <= 64 * 1024) pr.message(s); });
        ws.on('close', () => pr.close());
        ws.on('error', () => {});
      });
      return true;
    });
  } else host.log('ws library missing – co-editing and the remote control are off');

  /* ---------- 8: components ---------- */
  o.components = new Components({ dir: path.join(dataRoot, 'office-components'), offline: () => { try { return host.net && host.net.offline(); } catch { return false; } }, log: host.log, broadcast: (m) => host.broadcast(m) });

  /* ---------- 9: search by meaning ---------- */
  o.semantic = new Semantic({ file: file('semantic.json'), ai: () => ctx.ai || null, index: o.index, log: host.log });
  later(180000, () => o.semantic.refresh(15));
  every(10 * 60000, () => o.semantic.refresh(25));

  /* ---------- 12: publish (host.sites) ---------- */
  o.publish = new Publish(file('publish'));
  /** fallback page when the UI sends no rendered HTML: the document's text, readable on phones, no scripts */
  o.textPage = async (l) => {
    const text = l.st.size <= 50 * 1048576 ? await X.text(await fs.promises.readFile(l.full), extOf(l.full), o.libs()).catch(() => '') : '';
    const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const title = path.basename(l.full).replace(/\.[^.]+$/, '');
    const body = String(text || '').slice(0, 1500000).split(/\n{2,}|\r?\n(?=--- \d+ ---)/).map((p) => p.trim()).filter(Boolean).map((p) => (/^--- \d+ ---$/.test(p) ? '<hr>' : `<p>${esc(p).replace(/\n/g, '<br>')}</p>`)).join('\n');
    return { 'index.html': `<!doctype html><html lang="${host.lang()}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(title)}</title><link rel="stylesheet" href="style.css"></head><body><main><h1>${esc(title)}</h1>\n${body}</main></body></html>`,
      'style.css': ':root{color-scheme:light dark}body{margin:0;font:17px/1.6 system-ui,-apple-system,"Segoe UI",sans-serif;background:Canvas;color:CanvasText}main{max-width:760px;margin:0 auto;padding:32px 20px 64px}h1{font-size:28px;line-height:1.25}p{margin:0 0 1em;white-space:normal}hr{border:0;border-top:1px solid #8884;margin:2em 0}' };
  };
  if (host.sites && typeof host.sites.register === 'function') {
    try {
      host.sites.register({ id: 'office-doc', title: { de: 'Office-Dokument', en: 'Office document' }, defaults: { pub: '' },
        render: async (opt) => { const files = o.publish.filesOf(opt && opt.pub); if (!files) throw new Error(host.t('This Office page was removed.')); return files; } });
    } catch (e) { host.log('sites', e.message); }
  }

  /* ---------- 13: reports ---------- */
  o.reports = new Reports({
    file: file('reports.json'), readFile: readLoc, convert: (b, f, t) => o.convert(b, f, t), notifier: o.notifier, t: host.t, log: host.log,
    saveCopy: (loc, name, data) => { const l = lib.loc(loc); const relDir = l.rel.includes('/') ? l.rel.slice(0, l.rel.lastIndexOf('/')) : ''; return lib.create(l.root ? { root: l.root, dir: relDir } : relDir, name, data); },
  });
  every(60000, () => o.reports.tick());

  /* ---------- 14: library in the backup center ---------- */
  o.libBackup = new LibBackup({ dir: file('library-backups'), base: () => o.settings().libraryDir, log: host.log, yauzl: libOpt('yauzl'), write: (rel, data) => lib.write({ path: rel }, data, { force: true, label: 'Restored from a library backup' }) });
  if (host.backups && typeof host.backups.register === 'function') {
    try {
      host.backups.register({
        label: { de: 'Office-Bibliothek (Dokumente)', en: 'Office library (documents)' },
        list: () => o.libBackup.list().map((b) => ({ id: b.id, label: `${host.t('Office library')} · ${new Date(b.t).toLocaleString(host.lang() === 'de' ? 'de-DE' : 'en-US')}`, t: b.t, size: b.size, auto: true, svc: 'office-library', svcName: host.t('Office library'), href: '#/ext/office' })),
        restore: async (id) => o.libBackup.restore(id),
        remove: async (id) => o.libBackup.remove(id),
      }, 'library');
    } catch (e) { host.log('backups library', e.message); }
  }
  later(15 * 60000, () => o.libBackup.create());
  every(6 * 3600000, () => o.libBackup.create());

  /* ---------- 15: thumbnails ---------- */
  o.thumbs = new Thumbs(file('cache/thumbs'));

  /* ---------- hooks from the library ---------- */
  const ruleTimers = new Map();
  const onWrite = (l) => {
    if (o.rules.data.files[l.key] && !ruleTimers.has(l.key)) { const t = setTimeout(() => { ruleTimers.delete(l.key); o.rules.check(l.key).catch(() => {}); }, 1500); if (t.unref) t.unref(); ruleTimers.set(l.key, t); }
    for (const p of o.publish.list()) if (p.key === l.key && p.auto && p.mtime !== l.mtime) { o.publish.update(p.id, { stale: true }); try { host.broadcast({ type: 'office-publish', id: p.id, path: p.path, root: p.root || null }); } catch { /* egal */ } }
    o.collab.external(l.key, l.mtime);
  };
  const onRekey = (a, b) => { o.rules.rekey(a, b); o.forms.rekey(a, b); o.publish.rekey(a, b); o.reports.rekey(a, b); o.collab.rename(a, b); };
  const shutdown = () => { try { o.thumbs.flush(); } catch { /* egal */ } for (const off of offs.splice(0)) { try { off(); } catch { /* egal */ } } for (const t of ruleTimers.values()) clearTimeout(t); ruleTimers.clear(); };
  return { onWrite, onRekey, shutdown, thumbs: o.thumbs };
}

module.exports = { setup };
