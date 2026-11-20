'use strict';
/*
 * Office 1.1 server routes (contract: ARCHITECTURE.md → "Office 1.1 – expansion contract" and "Server 1.1").
 * host.router = /api/office/… behind Hostwerk's login · host.publicRouter = /api/office/pub/… (token pages)
 * Websockets: /ws/office/collab and /ws/office/remote (Hostwerk hands only /ws/* upgrades of logged-in users to
 * extensions; /api/office/… is accepted too in case a future core passes it on).
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { extOf, mimeOf } = require('./formats');
const E = require('./encrypt');
const { cleanFile } = require('./clean');
const F = require('./forms');
const S = require('./sheet');
const R = require('./remote');
const { uniqueName, stemExt } = require('./paths');
const { contentDisposition } = require('./routes');
const { readJson, writeJson } = require('./store');

const boolQ = (v) => v === '1' || v === 'true' || v === true || v === 1;
const OOXML = /^(docx|docm|dotx|dotm|xlsx|xlsm|xltx|xltm|pptx|pptm|ppsx|ppsm|potx|potm)$/;
const ODF = /^(odt|ods|odp|odg|ott|ots|otp|otg)$/;

function routes11(host, o) {
  const { router: r, publicRouter: pub } = host;
  const express = host.lib('express');
  const json = express.json({ limit: '4kb' });
  const big = express.json({ limit: '3mb' });
  const form = express.urlencoded({ extended: false, limit: '64kb' });
  const raw = express.raw({ type: () => true, limit: '4mb' });
  const lib = o.library;
  const err = (status, text, vars, extra) => Object.assign(host.error(status, text, vars), extra ? { extra } : {});
  const where = (src) => ({ root: src && src.root ? String(src.root) : null, path: String((src && src.path) || '') });
  const ctx = host.ctx;

  /* ---------------- helpers ---------------- */
  /** a NEW file next to l with "<stem> (<tag>).<ext>" (free name) → Entry */
  const sibling = (l, tag, data) => {
    const dir = path.dirname(l.full); const [stem, ext] = stemExt(path.basename(l.full));
    const name = uniqueName(`${stem} (${tag})${ext}`, (x) => fs.existsSync(path.join(dir, x)));
    const relDir = l.rel.includes('/') ? l.rel.slice(0, l.rel.lastIndexOf('/')) : '';
    return lib.create(l.root ? { root: l.root, dir: relDir } : relDir, name, data, { keep: true });
  };
  const restorePoint = (l, label) => { try { o.versions.add(l.key, fs.readFileSync(l.full), { label, auto: false }); } catch { /* no versions */ } };

  /* ================= 1. live data ================= */
  r.get('/office/data/sources', (req, res) => res.json({ sources: o.data.sources() }));
  r.get('/office/data/query', async (req, res) => {
    const { source, params, ...rest } = req.query;
    let p = rest; if (params) { try { p = { ...rest, ...JSON.parse(String(params)) }; } catch { throw err(400, 'Invalid parameters'); } }
    for (const k of Object.keys(p)) p[k] = String(p[k]).slice(0, 200);
    try { res.json(await o.data.query(String(source || ''), p)); } catch (e) { if (e.status) throw e; throw err(502, 'Hostwerk could not deliver this data: {error}', { error: e.message }); }
  });

  /* ================= 2. database tables ================= */
  r.get('/office/db/servers', (req, res) => res.json({ servers: o.dbs.servers() }));
  r.get('/office/db/databases', async (req, res) => res.json({ databases: await o.dbs.databases(req.query.server) }));
  r.get('/office/db/tables', async (req, res) => res.json({ tables: await o.dbs.tables(req.query.server, req.query.db) }));
  r.get('/office/db/rows', async (req, res) => {
    try { res.json(await o.dbs.rows(req.query.server, req.query.db, String(req.query.table || ''), { offset: req.query.offset, limit: req.query.limit, sort: req.query.sort, dir: req.query.dir, filter: req.query.filter })); } catch (e) { if (e.status) throw e; throw err(502, 'The database answered: {error}', { error: e.message }); }
  });
  r.post('/office/db/rows', express.json({ limit: '5mb' }), async (req, res) => {
    const b = req.body || {};
    try { res.json(await o.dbs.write(b.server, b.db, String(b.table || ''), b.changes)); } catch (e) { if (e.status) throw e; throw err(409, 'Nothing was saved – the database answered: {error}', { error: e.message }); }
  });

  /* ================= 3. live fields ================= */
  r.get('/office/fields', (req, res) => res.json({ fields: o.fields.catalog() }));
  r.get('/office/fields/value', async (req, res) => {
    let p = {}; if (req.query.params) { try { p = JSON.parse(String(req.query.params)); } catch { throw err(400, 'Invalid parameters'); } }
    if (!p || typeof p !== 'object' || Array.isArray(p)) throw err(400, 'Invalid parameters');
    res.json(await o.fields.value(String(req.query.id || ''), { ...p, lang: host.lang() }));
  });

  /* ================= 4. forms ================= */
  const formView = (f, req) => ({ id: f.id, title: f.title, path: f.path, root: f.root || null, name: f.name, sheet: f.sheet, columns: f.columns, timestamp: f.timestamp, description: f.description, thanks: f.thanks, notify: f.notify,
    public: { enabled: f.public.enabled, expires: f.public.expires || null, password: !!f.public.hash }, closed: !!f.closed, created: f.created, responses: f.responses || 0, lastResponse: f.lastResponse || null, pending: (f.queue || []).length,
    url: o.pubUrl(`form/${f.token}`, req), urls: o.pubUrls(`form/${f.token}`, req), open: F.open(f) });
  r.get('/office/forms', (req, res) => {
    let list = o.forms.list();
    if (req.query.path || req.query.root) { const l = lib.loc(where(req.query)); list = list.filter((f) => f.key === l.key); }
    res.json({ forms: list.map((f) => formView(f, req)) });
  });
  r.post('/office/forms', (req, res) => {
    const b = req.body || {};
    const f0 = lib.file(where(b));
    const ext = extOf(f0.full);
    if (!/^(xlsx|xlsm|ods|csv|tsv)$/.test(ext)) throw err(415, 'Forms work with xlsx, ods and csv files.');
    const def = F.validateForm(b);
    if (def.sheet) { const book = S.readBook(fs.readFileSync(f0.full), ext, { maxCells: 1000 }); if (!book.sheets.some((s) => s.name === def.sheet)) throw err(404, 'The sheet {name} does not exist (any more).', { name: def.sheet }); }
    if (def.public.password && def.public.password.length < 4) throw err(400, 'The password needs at least 4 characters.');
    const f = { id: crypto.randomBytes(6).toString('hex'), token: F.newToken(), key: f0.key, root: f0.root, path: f0.rel, name: path.basename(f0.full), ...def,
      public: { enabled: def.public.enabled, expires: o.expiryOf(def.public.expires), hash: def.public.password ? o.hashPassword(def.public.password) : null }, created: Date.now(), responses: 0, queue: [] };
    // header row once, when the sheet is still empty
    try { const book = S.readBook(fs.readFileSync(f0.full), ext, { maxCells: 50 }); const sh = S.sheetOf(book, def.sheet); if (sh && !Object.keys(sh.cells).length && b.header !== false) o.appendNow(f, [...(def.timestamp ? [host.t('Submitted at')] : []), ...def.columns.map((c) => c.label)]); } catch { /* the answers still arrive */ }
    o.forms.add(f);
    res.json(formView(f, req));
  });
  r.put('/office/forms/:id', (req, res) => {
    const f = o.forms.get(String(req.params.id)); if (!f) throw err(404, 'Form not found');
    const b = req.body || {};
    const def = F.validateForm({ ...formView(f, req), ...b, public: { ...f.public, ...(b.public || {}) } });
    const pub0 = b.public || {};
    const patch = { title: def.title, description: def.description, columns: def.columns, thanks: def.thanks, notify: def.notify, timestamp: def.timestamp,
      public: { enabled: def.public.enabled, expires: pub0.expires !== undefined ? o.expiryOf(pub0.expires) : f.public.expires, hash: pub0.password === '' || pub0.password === null ? null : pub0.password ? o.hashPassword(String(pub0.password)) : f.public.hash } };
    if (b.closed !== undefined) patch.closed = !!b.closed;
    res.json(formView(o.forms.update(f.id, patch), req));
  });
  r.post('/office/forms/:id/close', (req, res) => { const f = o.forms.update(String(req.params.id), { closed: true }); res.json(formView(f, req)); });
  r.delete('/office/forms/:id', (req, res) => { if (!o.forms.remove(String(req.params.id))) throw err(404, 'Form not found'); res.json({ ok: true }); });
  r.get('/office/forms/:id/pending', (req, res) => { const f = o.forms.get(String(req.params.id)); if (!f) throw err(404, 'Form not found'); res.json({ rows: (f.queue || []).map((q) => q.values.map((v) => (v && typeof v === 'object' ? S.dateText(v) : v))), sheet: f.sheet }); });
  r.post('/office/forms/:id/pending/ack', (req, res) => { const f = o.forms.get(String(req.params.id)); if (!f) throw err(404, 'Form not found'); const n = Math.max(0, Math.min((f.queue || []).length, Number((req.body || {}).count) || 0)); o.forms.update(f.id, { queue: f.queue.slice(n) }); res.json({ ok: true, pending: f.queue.length }); });

  // public form page
  const formHits = new Map(); // `${ip}\n${id}` → [t…]
  const limited = (ip, f) => { const k = `${ip}\n${f.id}`; const now = Date.now(); const list = (formHits.get(k) || []).filter((t) => now - t < 600000); const all = (formHits.get(f.id) || []).filter((t) => now - t < 3600000); if (list.length >= 10 || all.length >= 300) return true; list.push(now); all.push(now); formHits.set(k, list); formHits.set(f.id, all); if (formHits.size > 5000) formHits.clear(); return false; };
  const pubHead = (res) => { res.setHeader('X-Robots-Tag', 'noindex'); res.setHeader('Referrer-Policy', 'no-referrer'); res.setHeader('X-Content-Type-Options', 'nosniff'); res.setHeader('X-Frame-Options', 'DENY'); res.setHeader('Content-Security-Policy', "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'"); res.setHeader('Cache-Control', 'no-store'); };
  const langOf = (req) => (/^de\b/i.test(String(req.headers['accept-language'] || '')) ? 'de' : 'en');
  const formFor = (req, res) => {
    pubHead(res);
    if (o.guard.locked(req.ip)) { res.status(429).type('text').send('Too many attempts – try again later.'); return null; }
    const f = o.forms.byToken(req.params.token);
    if (!f) { o.guard.fail(req.ip); res.status(404).type('html').send(o.gonePage(langOf(req))); return null; }
    const base = req.originalUrl.split('?')[0].replace(/\/+$/, '').replace(/\/unlock$/, '');
    const unlocked = !f.public.hash || String(req.headers.cookie || '').split(/;\s*/).some((c) => c === `off_${f.id}=${crypto.createHmac('sha256', o.secret()).update(`${f.token}\n${f.public.hash}`).digest('base64url')}`);
    return { f, base, unlocked, lang: langOf(req) };
  };
  pub.get('/office/pub/form/:token', (req, res) => {
    const x = formFor(req, res); if (!x) return;
    if (!F.open(x.f)) return res.type('html').send(F.formPage(x.lang, x.f, { closed: true }));
    if (!x.unlocked) return res.type('html').send(F.formPage(x.lang, x.f, { locked: true, base: x.base }));
    res.type('html').send(F.formPage(x.lang, x.f, { base: x.base, stamp: F.stamp(o.secret(), x.f.token), done: req.query.sent === '1' }));
  });
  pub.post('/office/pub/form/:token/unlock', form, (req, res) => {
    const x = formFor(req, res); if (!x) return;
    const base = x.base;
    if (x.f.public.hash && !o.checkPassword(x.f.public.hash, (req.body || {}).password)) { o.guard.fail(req.ip); return res.status(403).type('html').send(F.formPage(x.lang, x.f, { locked: true, wrongPassword: true, base })); }
    res.setHeader('Set-Cookie', `off_${x.f.id}=${crypto.createHmac('sha256', o.secret()).update(`${x.f.token}\n${x.f.public.hash}`).digest('base64url')}; Path=${base}; HttpOnly; SameSite=Lax; Max-Age=86400`);
    res.redirect(303, base);
  });
  pub.post('/office/pub/form/:token', form, async (req, res) => {
    const x = formFor(req, res); if (!x) return;
    const f = x.f; const body = req.body || {};
    if (!F.open(f)) return res.status(410).type('html').send(F.formPage(x.lang, f, { closed: true }));
    if (!x.unlocked) return res.status(403).type('html').send(F.formPage(x.lang, f, { locked: true, base: x.base }));
    // spam: honeypot filled or sent faster than a human → pretend success, store nothing
    if (body.website || !F.stampOk(o.secret(), f.token, body._t)) return res.redirect(303, `${x.base}?sent=1`);
    if (limited(req.ip, f)) return res.status(429).type('text').send('Too many answers from your address – try again later.');
    let values;
    try { values = F.checkAnswer(f, body); } catch (e) { if (!e.fields) throw e; return res.status(400).type('html').send(F.formPage(x.lang, f, { base: x.base, errors: e.fields, values: body, stamp: F.stamp(o.secret(), f.token) })); }
    // a real date/time cell (local time of this Hostwerk); date fields of the form become real dates too
    const row = [...(f.timestamp ? [S.localStamp()] : []), ...values.map((v, i) => (f.columns[i] && f.columns[i].type === 'date' && /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? { $date: String(v) } : v))];
    try { await o.formAnswer(f, row); } catch (e) { host.log('form answer failed', e.message); return res.status(503).type('text').send('The answer could not be saved right now – please try again later.'); }
    res.redirect(303, `${x.base}?sent=1`);
  });

  /* ================= 5. rules ================= */
  r.get('/office/rules/channels', (req, res) => res.json({ channels: o.notifier.channels() }));
  r.get('/office/rules', (req, res) => { const l = lib.loc(where(req.query)); res.json({ rules: o.rules.get(l.key) }); });
  r.put('/office/rules', express.json({ limit: '256kb' }), async (req, res) => {
    const b = req.body; const src = Array.isArray(b) ? req.query : { ...req.query, ...(b || {}) };
    const f = lib.file(where(src));
    if (!/^(xlsx|xlsm|ods|csv|tsv)$/.test(extOf(f.full))) throw err(415, 'Rules work with xlsx, ods and csv files.');
    const list = Array.isArray(b) ? b : (b && b.rules) || [];
    res.json({ rules: o.rules.set(f.key, { root: f.root, path: f.rel }, list) });
  });
  r.post('/office/rules/test', async (req, res) => {
    const f = lib.file(where(req.body || {}));
    const book = S.readBook(fs.readFileSync(f.full), extOf(f.full));
    const list = require('./rules').validateRules((req.body || {}).rules || o.rules.get(f.key));
    res.json({ results: list.map((x) => ({ id: x.id, ...require('./rules').evaluate(book, x) })) });
  });

  /* ================= 6. co-editing (state; the relay is the websocket) ================= */
  r.get('/office/collab/state', (req, res) => { const l = lib.loc(where(req.query)); const room = o.collab.rooms.get(l.key); res.json({ key: l.key, peers: room ? room.peerList() : [], ws: '/ws/office/collab', yjs: require('./collab').hasYjs }); });

  /* ================= 7. remote control ================= */
  r.post('/office/remote', async (req, res) => {
    const s = o.remote.create(o.ownerName() || '');
    const urls = o.pubUrls(`remote/${s.token}`, req);
    let qr = null;
    try { qr = await host.lib('qrcode').toString(urls[0], { type: 'svg', margin: 1, errorCorrectionLevel: 'M' }); } catch (e) { host.log('qrcode', e.message); }
    res.json({ token: s.token, url: urls[0], urls, qr, ws: `/ws/office/remote?token=${s.token}`, expires: s.expires, lan: o.lanHint(req) });
  });
  r.delete('/office/remote/:token', (req, res) => { o.remote.end(String(req.params.token)); res.json({ ok: true }); });
  const remoteFor = (req, res) => { res.setHeader('X-Robots-Tag', 'noindex'); res.setHeader('Referrer-Policy', 'no-referrer'); res.setHeader('X-Content-Type-Options', 'nosniff'); res.setHeader('Cache-Control', 'no-store'); const s = o.remote.get(req.params.token); if (!s) { o.guard.fail(req.ip); res.status(404).type('html').send(o.gonePage(langOf(req))); return null; } return s; };
  pub.get('/office/pub/remote/:token', (req, res) => {
    if (o.guard.locked(req.ip)) return res.status(429).type('text').send('Too many attempts – try again later.');
    const s = remoteFor(req, res); if (!s) return;
    res.setHeader('Content-Security-Policy', "default-src 'none'; script-src 'self'; style-src 'unsafe-inline'; connect-src 'self'; base-uri 'none'; frame-ancestors 'none'");
    res.type('html').send(R.phonePage(langOf(req), req.originalUrl.split('?')[0].replace(/\/+$/, '')));
  });
  pub.get('/office/pub/remote/:token/app.js', (req, res) => { if (!remoteFor(req, res)) return; res.type('text/javascript; charset=utf-8').send(R.PHONE_JS); });
  pub.get('/office/pub/remote/:token/events', (req, res) => {
    const s = remoteFor(req, res); if (!s) return;
    res.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-store', Connection: 'keep-alive', 'X-Accel-Buffering': 'no' });
    const ph = { send: (m) => res.write(`data: ${JSON.stringify(m)}\n\n`), end: (m) => { res.write(`data: ${JSON.stringify(m)}\n\n`); res.end(); } };
    let off;
    try { off = o.remote.phone(s.token, ph); } catch (e) { res.end(); return; }
    const ping = setInterval(() => res.write(': ping\n\n'), 25000);
    req.on('close', () => { clearInterval(ping); if (off) off(); });
  });
  pub.post('/office/pub/remote/:token/cmd', json, (req, res) => { if (!remoteFor(req, res)) return; try { res.json(o.remote.command(req.params.token, req.body || {})); } catch (e) { res.status(e.status || 400).json({ error: e.message }); } });

  /* ================= 8. components on demand ================= */
  r.get('/office/components', (req, res) => res.json({ components: o.components.list() }));
  r.post('/office/components/install', (req, res) => { const id = String((req.body || {}).id || ''); o.components.def(id); o.components.install(id).catch(() => {}); res.json(o.components.view(id)); });
  r.delete('/office/components/:id', (req, res) => res.json(o.components.remove(String(req.params.id))));
  r.get(/^\/office\/components\/([a-z0-9-]{1,30})\/(.{1,200})$/, (req, res) => {
    const f = o.components.file(req.params[0], decodeURIComponent(req.params[1]));
    if (!f) throw err(404, 'Not found');
    const etag = `"${f.sha256.slice(0, 32)}"`;
    res.setHeader('ETag', etag); res.setHeader('Cache-Control', 'private, max-age=2592000, immutable'); res.setHeader('X-Content-Type-Options', 'nosniff');
    if (req.headers['if-none-match'] === etag) return res.status(304).end();
    res.setHeader('Content-Type', f.type); res.setHeader('Content-Length', String(f.size));
    fs.createReadStream(f.full).on('error', (e) => res.destroy(e)).pipe(res);
  });

  /* ================= 9. search by meaning ================= */
  r.get('/office/search/meaning', async (req, res) => {
    const out = await o.semantic.search(String(req.query.q || ''), { dir: req.query.dir ? String(req.query.dir) : '', limit: Math.min(100, Number(req.query.limit) || 30) });
    const hits = []; for (const h of out.hits) { try { hits.push({ ...lib.entryOf({ path: h.rel }), snippet: h.snippet, score: h.score }); } catch { /* gone */ } }
    res.json({ hits, mode: out.mode, model: out.model || null, terms: out.terms || null, index: o.semantic.state() });
  });

  /* ================= 10. encryption ================= */
  const VAULT_SRC = 'office';
  const vault = async () => { try { const v = ctx.passwordVault || (await host.use('passwordVault')); return v && typeof v.status === 'function' ? v : null; } catch { return null; } };
  const vaultKey = (l) => `file:${crypto.createHash('sha256').update(l.key).digest('hex').slice(0, 40)}`;
  const kindOf = (buf, ext) => (E.isEncryptedOoxml(buf) ? 'ooxml' : ODF.test(ext) && E.isEncryptedOdf(buf) ? 'odf' : null);
  r.get('/office/encrypt/state', async (req, res) => {
    const l = lib.file(where(req.query)); const ext = extOf(l.full);
    const buf = l.st.size <= 200 * 1048576 ? fs.readFileSync(l.full) : Buffer.alloc(0);
    const v = await vault(); let vs = { available: false };
    if (v) { try { const st = await v.status(); vs = { available: !!st.available, locked: !!st.locked, name: st.vault ? st.vault.name : null }; if (st.available && !st.locked) vs.hasEntry = !!(await v.find({ source: VAULT_SRC, key: vaultKey(l) })); } catch (e) { vs = { available: true, locked: e.status === 423, error: e.message }; } }
    res.json({ encrypted: kindOf(buf, ext), supported: OOXML.test(ext) ? 'ooxml' : ODF.test(ext) ? 'odf' : ext === 'pdf' ? 'pdf-not-supported' : null, vault: vs });
  });
  const passwordFrom = async (b, l, { forEncrypt }) => {
    if (b.password) { const p = String(b.password); if (p.length < (forEncrypt ? 4 : 1) || p.length > 255) throw err(400, 'The password needs at least 4 characters.'); return { pw: p }; }
    if (b.vault || b.fromVault) {
      const v = await vault(); if (!v) throw err(409, 'The password vault is not available – type the password instead.');
      if (forEncrypt) { const pw = crypto.randomBytes(18).toString('base64url'); return { pw, generated: true, v }; }
      const got = await v.reveal({ source: VAULT_SRC, key: vaultKey(l) }); if (!got || !got.password) throw err(404, 'The vault has no password for this file.'); return { pw: got.password, fromVault: true };
    }
    throw err(400, 'Please enter a password.');
  };
  const doEncrypt = async (req, res) => {
    const b = req.body || {}; const l = lib.file(where(b)); const ext = extOf(l.full);
    if (ext === 'pdf') throw err(501, 'PDF passwords are not available yet – encrypt the Office file instead, or export the PDF with a password in the desktop app.');
    if (!OOXML.test(ext) && !ODF.test(ext)) throw err(415, 'Only Word, Excel, PowerPoint and OpenDocument files can be encrypted.');
    if (l.st.size > 200 * 1048576) throw err(413, 'The file is too large.');
    const buf = fs.readFileSync(l.full);
    if (kindOf(buf, ext)) throw err(409, 'This file is already encrypted.');
    const p = await passwordFrom(b, l, { forEncrypt: true });
    let out; try { out = OOXML.test(ext) ? E.encryptOoxml(buf, p.pw) : E.encryptOdf(buf, p.pw); } catch (e) { throw err(e.status || 400, e.message); }
    let vaultSaved = false;
    if (p.v || b.saveToVault) {
      const v = p.v || (await vault()); if (!v) throw err(409, 'The password vault is not available – type the password instead.');
      await v.save({ source: VAULT_SRC, key: vaultKey(l), name: `Office: ${path.basename(l.full)}`, username: '', password: p.pw, url: '', notes: host.t('Password of the encrypted Office file {name}', { name: l.root ? `${l.root}:${l.rel}` : l.rel }), folder: 'Office' });
      vaultSaved = true;
    }
    if (b.mode === 'copy') return res.json({ ok: true, entry: sibling(l, host.t('encrypted'), out), vaultSaved, generated: !!p.generated });
    restorePoint(l, 'Before encrypting');
    const w = lib.write(where(b), out, { force: true });
    res.json({ ok: true, ...w, vaultSaved, generated: !!p.generated, versionsWithPlainText: o.versions.list(l.key).length });
  };
  r.post('/office/encrypt', doEncrypt);
  const doDecrypt = async (req, res) => {
    const b = req.body || {}; const l = lib.file(where(b)); const ext = extOf(l.full);
    if (l.st.size > 200 * 1048576) throw err(413, 'The file is too large.');
    const buf = fs.readFileSync(l.full);
    const kind = kindOf(buf, ext);
    if (!kind) throw err(409, 'This file is not encrypted.');
    // soft: the editor asks again on a wrong password – an expected answer, so 200 { ok: false, reason } instead of a
    // 403/404 that every browser logs as a red console error
    const soft = boolQ(b.soft);
    let p;
    try { p = await passwordFrom(b, l, { forEncrypt: false }); } catch (e) { if (soft && e.status === 404) return res.json({ ok: false, reason: 'vault', error: host.t(e.message) }); throw e; }
    let out; try { out = kind === 'ooxml' ? E.decryptOoxml(buf, p.pw) : E.decryptOdf(buf, p.pw); } catch (e) {
      if (e.code === 'EPASSWORD') { o.guard.fail(`dec:${req.ip}`); if (soft) return res.json({ ok: false, reason: 'password', error: host.t('The password is not correct.') }); throw err(403, 'The password is not correct.'); }
      throw err(e.status || 422, e.message);
    }
    if (b.mode === 'inplace') { const w = lib.write(where(b), out, { force: true }); return res.json({ ok: true, ...w }); }
    if (b.mode === 'copy') return res.json({ ok: true, entry: sibling(l, host.t('decrypted'), out) });
    // default: the bytes for the editor (the file stays encrypted)
    res.setHeader('Content-Type', mimeOf(l.full)); res.setHeader('Cache-Control', 'no-store'); res.setHeader('x-mtime', String(Math.floor(l.st.mtimeMs))); res.setHeader('Access-Control-Expose-Headers', 'x-mtime');
    res.end(out);
  };
  r.post('/office/decrypt', doDecrypt);
  // saving an opened encrypted file: the editor sends plain bytes, the server encrypts with the given / vault password
  r.put('/office/encrypt/raw', express.raw({ type: () => true, limit: '200mb' }), async (req, res) => {
    const l = lib.file(where(req.query)); const ext = extOf(l.full);
    const pwh = req.headers['x-office-password'] ? decodeURIComponent(String(req.headers['x-office-password'])) : '';
    const p = await passwordFrom({ password: pwh, fromVault: !pwh }, l, { forEncrypt: false });
    const data = Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0);
    let out; try { out = OOXML.test(ext) ? E.encryptOoxml(data, p.pw) : E.encryptOdf(data, p.pw); } catch (e) { throw err(e.status || 400, e.message); }
    res.json(lib.write(where(req.query), out, { mtime: req.query.mtime, force: boolQ(req.query.force) }));
  });

  /* ================= 11. clean before sharing ================= */
  const CLEAN_ITEMS = [['author', 'Author, company and last editor', 'author'], ['comments', 'Comments', 'comments'], ['revisions', 'Tracked changes (accepted)', 'revisions'], ['hidden', 'Hidden text', 'hidden'], ['custom', 'Custom properties', 'custom'], ['thumbnail', 'Thumbnail', 'thumbnail'], ['rsid', 'Revision ids', 'rsid'], ['exif', 'Image metadata (EXIF)', 'exif'], ['templates', 'Template and file paths', 'paths']];
  const WARN = { macros: 'The file contains macros – they stay.', hiddenSheets: 'Hidden sheets stay – check them.', hiddenSlides: 'Hidden slides stay – check them.', hiddenStyles: 'Text hidden by a style stays – check it.', pdfCompressed: 'This PDF keeps its metadata in compressed parts – it cannot be cleaned here.', pdfPartial: 'Parts of this PDF are compressed – some metadata may stay.', pdfAnnotations: 'PDF comments (annotations) stay.', pdfScripts: 'The PDF contains scripts.' };
  const cleanRun = (b) => {
    const l = lib.file(where(b)); const ext = extOf(l.full);
    if (l.st.size > 200 * 1048576) throw err(413, 'The file is too large.');
    const buf = fs.readFileSync(l.full);
    if (kindOf(buf, ext)) throw err(409, 'Decrypt the file first, then clean it.');
    // options: { author: false, … } or the shell's remove: [ids] (only those)
    const opts = Array.isArray(b.remove) ? Object.fromEntries(CLEAN_ITEMS.map(([id]) => [id, b.remove.includes(id)])) : b.options || {};
    let out; try { out = cleanFile(buf, ext, opts); } catch (e) { throw err(e.status || 422, e.message); }
    return { l, out };
  };
  r.post('/office/clean/preview', (req, res) => {
    const { out } = cleanRun({ ...(req.body || {}), remove: undefined, options: {} });
    const rp = out.report;
    const items = CLEAN_ITEMS.map(([id, label, key]) => ({ id, label: host.t(label), count: typeof rp[key] === 'number' ? rp[key] : rp[key] ? 1 : 0 })).filter((x) => x.count);
    res.json({ items, warnings: out.warnings.map((w) => ({ id: w, label: host.t(WARN[w] || w) })) });
  });
  r.post('/office/clean', async (req, res) => {
    const b = req.body || {};
    const { l, out } = cleanRun(b);
    const report = { ...out.report, warnings: out.warnings };
    if (Array.isArray(b.remove)) { // shell shape → Entry (+ report)
      if (b.copy === false) { restorePoint(l, 'Before cleaning'); lib.write(where(b), out.data, { force: true }); return res.json({ ...lib.entryOf(where(b)), report }); }
      return res.json({ ...sibling(l, host.t('cleaned'), out.data), report });
    }
    if (b.mode === 'inplace') { restorePoint(l, 'Before cleaning'); return res.json({ ok: true, report, ...lib.write(where(b), out.data, { force: true }) }); }
    if (b.mode === 'download') { res.setHeader('Content-Type', mimeOf(l.full)); res.setHeader('Content-Disposition', contentDisposition('attachment', path.basename(l.full))); res.setHeader('x-clean-report', Buffer.from(JSON.stringify(report)).toString('base64')); res.setHeader('Access-Control-Expose-Headers', 'x-clean-report'); return res.end(out.data); }
    res.json({ ok: true, report, entry: sibling(l, host.t('cleaned'), out.data) });
  });

  /* ================= 12. publish as web page ================= */
  const sites = () => { if (!host.sites) throw err(501, 'Publishing needs Hostwerk 3.2 or newer.'); return host.sites; };
  const checkFiles = (files) => {
    if (!files || typeof files !== 'object' || Array.isArray(files)) throw err(400, 'Invalid page files');
    const names = Object.keys(files); let size = 0;
    if (!names.includes('index.html') || names.length > 12) throw err(400, 'A page needs index.html and at most 12 files.');
    for (const n of names) { if (!/^[a-z0-9][a-z0-9_-]{0,60}\.(html|json|css|js|svg|txt)$/i.test(n) || typeof files[n] !== 'string') throw err(400, 'Invalid file name {name}', { name: n }); size += Buffer.byteLength(files[n]); }
    if (size > 2 * 1048576) throw err(413, 'The page is larger than 2 MB – images are probably too big.');
    return files;
  };
  r.get('/office/publish/targets', (req, res) => res.json({ targets: sites().targets() }));
  r.get('/office/publish', (req, res) => {
    let list = o.publish.list();
    if (req.query.path || req.query.root) { const l = lib.loc(where(req.query)); list = list.filter((p) => p.key === l.key); }
    const live = new Map((host.sites ? host.sites.list() : []).map((x) => [x.id, x]));
    const pages = list.map((p) => { const site = live.get(p.siteId) || null; return { ...o.publish.view(p), site, url: site ? site.url : null, target: site ? site.target : null, updatedAt: p.publishedAt }; });
    res.json({ pages, publications: pages });
  });
  r.post('/office/publish', big, async (req, res) => {
    const b = req.body || {}; const l = lib.file(where(b));
    // without files from the UI: a plain readable page from the document's text (the UI can PUT the rendered page later)
    const files = checkFiles(b.files || await o.textPage(l));
    let t = (b.target && typeof b.target === 'object') ? b.target : {};
    if (typeof b.target === 'string') { const hit = sites().targets().find((x) => x.id === b.target); if (!hit) throw err(404, 'Unknown target'); t = { web: hit.web, vhost: hit.vhost }; }
    if (!b.sub && b.path) b.sub = String(b.path).replace(/^\/+|\/+$/g, '');
    const p = o.publish.add({ key: l.key, root: l.root, path: l.rel, name: path.basename(l.full), files, mtime: Math.floor(l.st.mtimeMs), auto: b.auto !== false });
    try {
      const site = await sites().save({ kind: 'office-doc', web: String(t.web || ''), vhost: t.vhost ? String(t.vhost) : null, sub: String(b.sub || ''), options: { pub: p.id }, interval: 86400, title: String(b.title || path.basename(l.full)).slice(0, 120), on: true });
      o.publish.update(p.id, { siteId: site.id });
      res.json({ ...o.publish.view(o.publish.get(p.id)), site, url: site.url || null });
    } catch (e) { o.publish.remove(p.id); throw err(e.status || 400, e.message); }
  });
  r.put('/office/publish/:id', big, async (req, res) => {
    const p = o.publish.get(String(req.params.id)); if (!p) throw err(404, 'Page not found');
    const b = req.body || {};
    const patch = { stale: false };
    if (b.files) { patch.files = checkFiles(b.files); patch.publishedAt = Date.now(); try { patch.mtime = Math.floor(lib.file({ root: p.root, path: p.path }).st.mtimeMs); } catch { /* file gone – keep the page */ } }
    if (b.auto !== undefined) patch.auto = !!b.auto;
    o.publish.update(p.id, patch);
    const site = p.siteId && b.files ? await sites().publish(p.siteId) : null;
    res.json({ ...o.publish.view(o.publish.get(p.id)), site });
  });
  r.delete('/office/publish/:id', (req, res) => { const p = o.publish.get(String(req.params.id)); if (!p) throw err(404, 'Page not found'); try { if (p.siteId && host.sites) host.sites.remove(p.siteId); } catch { /* already gone */ } o.publish.remove(p.id); res.json({ ok: true }); });

  /* ================= 13. scheduled reports ================= */
  r.get('/office/reports', (req, res) => { let list = o.reports.list(); if (req.query.path || req.query.root) { const l = lib.loc(where(req.query)); list = list.filter((x) => x.key === l.key); } res.json({ reports: list, channels: o.notifier.channels() }); });
  const reportBody = (b0) => {
    const b = { ...b0 }; const s = { ...(b.schedule || {}) };
    const ev = { day: 'daily', week: 'weekly', month: 'monthly', hour: 'hourly' };
    if (s.every || b.every) s.every = ev[s.every || b.every] || s.every || b.every;
    if (s.at) s.time = s.at;
    if (b.at && !s.time) s.time = b.at;
    if (Array.isArray(b.recipients) && b.recipients[0] && !b.channel) b.channel = b.recipients[0].channel;
    return { ...b, schedule: { ...b, ...s } };
  };
  r.post('/office/reports', (req, res) => { const b = reportBody(req.body || {}); if (b.id) return res.json(o.reports.update(String(b.id), b)); const l = lib.file(where(b)); res.json(o.reports.add({ ...b, key: l.key, root: l.root, path: l.rel, name: path.basename(l.full) })); });
  r.put('/office/reports/:id', (req, res) => res.json(o.reports.update(String(req.params.id), reportBody(req.body || {}))));
  r.delete('/office/reports/:id', (req, res) => { o.reports.remove(String(req.params.id)); res.json({ ok: true }); });
  r.post('/office/reports/:id/run', async (req, res) => res.json(await o.reports.run(String(req.params.id))));

  /* ================= 14. library backups (backup center) ================= */
  r.get('/office/library-backups', (req, res) => res.json({ backups: o.libBackup.list(), settings: o.libBackup.state() }));
  r.post('/office/library-backups', async (req, res) => res.json(await o.libBackup.create({ force: true })));
  r.post('/office/library-backups/restore', async (req, res) => { const b = req.body || {}; res.json(await o.libBackup.restore(String(b.id || ''), { paths: Array.isArray(b.paths) ? b.paths : null })); });

  /* ---------------- library backups per file (shell): snapshots that contain this file · raw · restore ---------------- */
  const libRel = (src) => { const l = lib.loc(where(src)); if (l.root) throw err(409, 'Library backups only cover the Office library.'); return l; };
  r.get('/office/backups', (req, res) => { const l = libRel(req.query); res.json({ backups: o.libBackup.versionsOf(l.rel) }); });
  r.get('/office/backups/raw', (req, res) => {
    const l = libRel(req.query); const data = o.libBackup.readFile(String(req.query.id || ''), l.rel);
    if (!data) throw err(404, 'Backup not found');
    res.setHeader('Content-Type', mimeOf(l.full)); res.setHeader('Cache-Control', 'no-store');
    if (boolQ(req.query.download)) res.setHeader('Content-Disposition', contentDisposition('attachment', path.basename(l.full)));
    res.end(data);
  });
  r.post('/office/backups/restore', async (req, res) => { const b = req.body || {}; const l = libRel(b); const r0 = await o.libBackup.restore(String(b.id || ''), { paths: [l.rel] }); res.json({ ...r0, ...(fs.existsSync(l.full) ? lib.entryOf({ path: l.rel }) : {}) }); });

  /* ---------------- vault shortcuts (shell): generated password kept in Hostwerk's password vault ---------------- */
  r.get('/office/vault', async (req, res) => {
    const v = await vault();
    if (!v) return res.json({ available: false, unlocked: false });
    try { const st = await v.status(); res.json({ available: !!st.available, unlocked: !!st.available && !st.locked, name: st.vault ? st.vault.name : null }); } catch (e) { res.json({ available: true, unlocked: false, error: e.message }); }
  });
  /** run the encrypt/decrypt handler with the vault as password source and answer with the Entry */
  const viaVault = (fn, extra) => async (req, res) => {
    const b = { ...(req.body || {}), ...((req.body || {}).loc || {}) };
    let result = null; let raw = null;
    const fake = { json: (x) => { result = x; }, setHeader: () => {}, end: (x) => { raw = x; } };
    await fn({ body: { ...b, vault: true, fromVault: true, ...extra }, ip: req.ip, headers: req.headers }, fake);
    res.json({ ...lib.entryOf(where(b)), result: result || (raw ? { ok: true } : null) });
  };
  r.post('/office/vault/encrypt', viaVault((rq, rs) => doEncrypt(rq, rs), {}));
  r.post('/office/vault/decrypt', viaVault((rq, rs) => doDecrypt(rq, rs), { mode: 'inplace' }));

  /* ---------------- saved signatures of the PDF editor (per user, data folder, size-limited) ---------------- */
  const SIG = { count: 8, each: 512 * 1024, total: 2 * 1048576 };
  const sigFile = () => o.dataFile('signatures.json');
  const sigAll = () => readJson(sigFile(), {});
  r.get('/office/signatures', (req, res) => {
    const all = sigAll(); const list = all && typeof all === 'object' && Array.isArray(all[o.userOf(req)]) ? all[o.userOf(req)] : [];
    res.setHeader('Cache-Control', 'no-store');
    res.json({ signatures: list });
  });
  r.put('/office/signatures', big, (req, res) => {
    const b = req.body || {}; const inList = Array.isArray(b) ? b : Array.isArray(b.signatures) ? b.signatures : null;
    if (!inList || inList.length > SIG.count) throw err(400, 'Invalid parameters');
    let total = 0;
    const list = inList.map((x) => {
      const u = String((x && x.url) || '');
      if (!/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/.test(u) || u.length > SIG.each) throw err(400, 'A signature must be a PNG, JPEG or WebP image up to 512 KB.');
      total += u.length;
      const num = (v, max) => { const n = Number(v); return Number.isFinite(n) && n > 0 ? Math.min(max, Math.round(n)) : null; };
      return { url: u, w: num(x.w, 8000), h: num(x.h, 8000), at: num(x.at, 8.64e15) || Date.now() };
    });
    if (total > SIG.total) throw err(400, 'The saved signatures are too large together (max. 2 MB).');
    const all = sigAll(); const key = o.userOf(req);
    const next = all && typeof all === 'object' && !Array.isArray(all) ? all : {};
    if (list.length) next[key] = list; else delete next[key];
    writeJson(sigFile(), next);
    res.json({ signatures: list });
  });

  /* ---------------- QR code (SVG) ---------------- */
  r.get('/office/qr', async (req, res) => {
    const text = String(req.query.text || '');
    if (!text || text.length > 2000) throw err(400, 'Invalid parameters');
    const svg = await host.lib('qrcode').toString(text, { type: 'svg', margin: 1, errorCorrectionLevel: 'M' });
    res.setHeader('Content-Type', 'image/svg+xml'); res.setHeader('Cache-Control', 'private, max-age=3600'); res.setHeader('X-Content-Type-Options', 'nosniff');
    res.send(svg);
  });

  /* ---------------- AI for the editors (Hostwerk's configured provider, never a fixed cloud) ---------------- */
  r.post('/office/ai', express.json({ limit: '1mb' }), async (req, res) => {
    const b = req.body || {}; const ai = ctx.ai;
    if (!ai) throw err(409, 'The AI assistant is not ready.');
    const lang = b.lang === 'en' || b.lang === 'de' ? b.lang : host.lang();
    const text = String(b.text || '').slice(0, 60000);
    if (b.task && ['translate', 'explain', 'summarize', 'describe'].includes(b.task) && ai.quick) { const r0 = await ai.quick.run(b.task, { text, lang, context: String(b.context || '').slice(0, 300) }); return res.json({ text: r0.text }); }
    if (!ai.chats || typeof ai.chats.oneShot !== 'function') throw err(409, 'The AI assistant is not ready.');
    const prompt = String(b.prompt || '').slice(0, 8000);
    if (!prompt && !text) throw err(400, 'Invalid parameters');
    const sys = `${String(b.system || '').slice(0, 4000) || 'You help inside an office suite.'} Answer in ${lang === 'de' ? 'German' : 'English'}.${b.json ? ' Answer with valid JSON only, no code fences.' : ''}`;
    let out;
    try { out = await ai.chats.oneShot([{ role: 'system', content: sys }, { role: 'user', content: text ? `${prompt}\n\n---\n${text}` : prompt }], { timeout: 180000, temperature: b.json ? 0 : 0.3 }); } catch (e) { throw err(502, 'The AI did not answer: {error}', { error: e.message }); }
    let json = null;
    if (b.json) { try { json = JSON.parse(String(out).trim().replace(/^```(?:json)?\s*|\s*```$/g, '')); } catch { json = null; } }
    res.json({ text: String(out || ''), ...(b.json ? { json } : {}) });
  });

  /* ================= 15. thumbnails from the UI ================= */
  r.put('/office/thumb', raw, (req, res) => {
    const l = lib.file(where(req.query));
    const data = Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0);
    const isPng = data.length >= 33 && data.readUInt32BE(0) === 0x89504e47;
    const isWebp = data.length >= 16 && data.toString('latin1', 0, 4) === 'RIFF' && data.toString('latin1', 8, 12) === 'WEBP';
    if (!(isPng || isWebp) || data.length > 2 * 1048576) throw err(400, 'A thumbnail must be a PNG or WebP image up to 2 MB.');
    if (isPng) { const w = data.readUInt32BE(16); const h = data.readUInt32BE(20); if (!w || !h || w > 2048 || h > 2048) throw err(400, 'A thumbnail must be a PNG or WebP image up to 2 MB.'); }
    const mtime = Number(req.query.mtime) || Math.floor(l.st.mtimeMs);
    if (Math.abs(mtime - Math.floor(l.st.mtimeMs)) > 1) throw err(409, 'The file changed – render the thumbnail again.', null, { mtime: Math.floor(l.st.mtimeMs) });
    o.thumbs.put(l.key, Math.floor(l.st.mtimeMs), data, isPng ? 'image/png' : 'image/webp');
    res.json({ ok: true });
  });
}

module.exports = { routes11 };
