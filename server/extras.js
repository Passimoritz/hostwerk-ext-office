'use strict';
/*
 * Smaller 1.1 stores:
 *   Thumbs      – thumbnails rendered by the UI (PNG), cached per file key + mtime (cache/thumbs/)
 *   Publish     – documents published as web pages through host.sites (the UI renders the HTML, we keep and republish it)
 *   Reports     – scheduled reports (PDF or the file's own format) to several recipients each (own timer with persistence;
 *                 Hostwerk's scheduler has no task type for extensions)
 *   LibBackup   – zip snapshots of the whole library for Hostwerk's backup center (restore never deletes, every
 *                 overwritten file keeps its current content as a version)
 *   nextRun(schedule, from) – pure schedule helper
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { readJson, writeJson } = require('./store');
const { ZipWriter, readZip } = require('./zip');
const P = require('./paths');
const { DISCORD } = require('./notify');

const sha1 = (s) => crypto.createHash('sha1').update(String(s)).digest('hex');

/** type: 'image/png' | 'image/webp'; an in-memory index (index.json) answers has() without touching the disk */
class Thumbs {
  constructor(dir, { max = 3000 } = {}) { this.dir = dir; this.max = max; this.idx = readJson(path.join(dir, 'index.json'), {}); this.timer = null; }
  file(key) { return path.join(this.dir, `${sha1(key)}.img`); }
  saveIdx() { if (this.timer) return; this.timer = setTimeout(() => { this.timer = null; try { writeJson(path.join(this.dir, 'index.json'), this.idx); } catch { /* next time */ } }, 1000); if (this.timer.unref) this.timer.unref(); }
  put(key, mtime, data, type = 'image/png') {
    fs.mkdirSync(this.dir, { recursive: true });
    fs.writeFileSync(this.file(key), data);
    this.idx[sha1(key)] = { mtime, type, t: Date.now() };
    this.prune(); this.saveIdx();
  }
  has(key, mtime) { const m = this.idx[sha1(key)]; return !!m && m.mtime === mtime; }
  get(key, mtime) {
    const m = this.idx[sha1(key)]; if (!m || m.mtime !== mtime) return null;
    try { return { data: fs.readFileSync(this.file(key)), type: m.type || 'image/png' }; } catch { return null; }
  }
  prune() {
    const keys = Object.keys(this.idx); if (keys.length <= this.max) return;
    for (const k of keys.sort((a, b) => this.idx[a].t - this.idx[b].t).slice(0, keys.length - this.max)) { delete this.idx[k]; fs.rmSync(path.join(this.dir, `${k}.img`), { force: true }); }
  }
  flush() { if (this.timer) { clearTimeout(this.timer); this.timer = null; try { writeJson(path.join(this.dir, 'index.json'), this.idx); } catch { /* egal */ } } }
}

class Publish {
  constructor(dir) { this.dir = dir; this.file = path.join(dir, 'publish.json'); fs.mkdirSync(dir, { recursive: true }); this.items = readJson(this.file, []); }
  save() { writeJson(this.file, this.items); }
  list() { return this.items; }
  get(id) { return this.items.find((p) => p.id === id) || null; }
  view(p) { const { files, ...rest } = p; return { ...rest, files: files ? Object.keys(files) : this.filesOf(p.id) ? Object.keys(this.filesOf(p.id)) : [] }; }
  filesOf(id) { if (!/^[a-f0-9]{12}$/.test(String(id))) return null; return readJson(path.join(this.dir, `${id}.json`), null); }
  add(p) {
    if (this.items.length >= 100) throw Object.assign(new Error('At most 100 published pages.'), { status: 409 });
    const id = crypto.randomBytes(6).toString('hex');
    writeJson(path.join(this.dir, `${id}.json`), p.files);
    const it = { id, key: p.key, root: p.root || null, path: p.path, name: p.name, mtime: p.mtime, auto: p.auto !== false, stale: false, createdAt: Date.now(), publishedAt: Date.now(), siteId: null };
    this.items.push(it); this.save(); return it;
  }
  update(id, patch) { const p = this.get(id); if (!p) return null; const { files, ...rest } = patch; if (files) writeJson(path.join(this.dir, `${id}.json`), files); Object.assign(p, rest); this.save(); return p; }
  remove(id) { this.items = this.items.filter((p) => p.id !== id); fs.rmSync(path.join(this.dir, `${id}.json`), { force: true }); this.save(); }
  rekey(a, b) { let ch = false; for (const p of this.items) if (p.key === a || p.key.startsWith(`${a}/`)) { p.key = b + p.key.slice(a.length); if (!p.root) p.path = p.key; ch = true; } if (ch) this.save(); }
}

/* ---------------- schedules ---------------- */
const EVERY = ['hourly', 'daily', 'weekly', 'monthly'];
function normSchedule(b = {}) {
  const every = EVERY.includes(b.every) ? b.every : 'weekly';
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(b.time || '08:00'));
  const hh = m ? Math.min(23, Number(m[1])) : 8; const mm = m ? Math.min(59, Number(m[2])) : 0;
  return { every, time: `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`, weekday: Math.min(6, Math.max(0, Number.isFinite(Number(b.weekday)) ? Number(b.weekday) : 1)), day: Math.min(28, Math.max(1, Number(b.day) || 1)) };
}
/** next run strictly after `from` (local time) */
function nextRun(s0, from = Date.now()) {
  const s = normSchedule(s0);
  const [hh, mm] = s.time.split(':').map(Number);
  const d = new Date(from);
  if (s.every === 'hourly') { const n = new Date(d); n.setMinutes(mm, 0, 0); if (n.getTime() <= from) n.setHours(n.getHours() + 1); return n.getTime(); }
  const n = new Date(d); n.setHours(hh, mm, 0, 0);
  if (s.every === 'daily') { if (n.getTime() <= from) n.setDate(n.getDate() + 1); return n.getTime(); }
  if (s.every === 'weekly') { let add = (s.weekday - n.getDay() + 7) % 7; if (add === 0 && n.getTime() <= from) add = 7; n.setDate(n.getDate() + add); return n.getTime(); }
  n.setDate(s.day); if (n.getTime() <= from) { n.setMonth(n.getMonth() + 1); n.setDate(s.day); } return n.getTime();
}

const CHANNELS = ['hostwerk', 'alerts', 'mail', 'discord', 'ntfy'];
const MAIL = /^[^\s@<>,;"]{1,64}@[^\s@<>,;"]{1,255}\.[^\s@<>,;"]{2,}$/;
/**
 * recipients of a report: [{ channel: hostwerk | alerts (every alert channel of Hostwerk) | mail | discord | ntfy, to }]
 * mail: one or more addresses ("a@x.de, b@y.de", ≤ 20) – empty = Hostwerk's alert address · discord: a webhook address –
 * empty = the alert webhook · ntfy: a full address (https://ntfy.sh/topic) or a topic name – empty = the alert address.
 * Throws 400 with the first invalid entry.
 */
function normRecipients(list, fallbackChannel) {
  const src = Array.isArray(list) && list.length ? list : [{ channel: fallbackChannel === 'all' ? 'alerts' : fallbackChannel || 'mail', to: '' }];
  const out = [];
  for (const r0 of src.slice(0, 20)) {
    const channel = r0 && CHANNELS.includes(r0.channel) ? r0.channel : r0 && r0.channel === 'all' ? 'alerts' : null;
    if (!channel) throw Object.assign(new Error('Unknown channel'), { status: 400 });
    let to = String((r0 && r0.to) || '').trim().slice(0, 2000);
    if (channel === 'hostwerk' || channel === 'alerts') to = '';
    if (channel === 'mail' && to) {
      const list2 = to.split(/[\s,;]+/).filter(Boolean);
      const bad = list2.find((x) => !MAIL.test(x));
      if (bad || list2.length > 20) throw Object.assign(new Error('{to} is not a valid e-mail address.'), { status: 400, vars: { to: bad || to } });
      to = list2.join(', ');
    }
    if (channel === 'discord' && to && !DISCORD.test(to)) throw Object.assign(new Error('This is not a Discord webhook address.'), { status: 400 });
    if (channel === 'ntfy' && to && !/^https?:\/\/[^\s/]+\/[\w-]{1,64}$/.test(to) && !/^[\w-]{1,64}$/.test(to)) throw Object.assign(new Error('Enter an ntfy address (https://ntfy.sh/topic) or just the topic.'), { status: 400 });
    if (!out.some((x) => x.channel === channel && x.to === to)) out.push({ channel, to });
  }
  return out;
}
const FORMATS = ['pdf', 'docx', 'odt', 'xlsx', 'ods', 'csv', 'pptx', 'odp', 'same'];
const MIME = { pdf: 'application/pdf', csv: 'text/csv', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', odt: 'application/vnd.oasis.opendocument.text', ods: 'application/vnd.oasis.opendocument.spreadsheet', odp: 'application/vnd.oasis.opendocument.presentation' };

class Reports {
  /** deps: { file, readFile(loc) → { buf, ext }, convert(buf, from, to), notifier, saveCopy(loc, name, data) → Entry, t, log } */
  constructor(deps) { this.d = deps; this.items = readJson(deps.file, []); this.running = new Set(); }
  save() { writeJson(this.d.file, this.items); }
  list() { return this.items.map((it) => this.view(it)); }
  /** the shape the UI edits (recipients always present – older entries had one channel; schedule.at = time) */
  view(it) { return { ...it, recipients: it.recipients || normRecipients(null, it.channel), format: it.format || 'pdf', schedule: { ...it.schedule, at: it.schedule.time } }; }
  norm(b, old = {}) {
    const ch = ['hostwerk', 'discord', 'ntfy', 'mail', 'all'];
    const channel = ch.includes(b.channel) ? b.channel : old.channel || 'mail';
    const recipients = b.recipients !== undefined ? normRecipients(b.recipients, channel) : old.recipients || normRecipients(null, channel);
    const f0 = String(b.format || '').toLowerCase();
    const format = FORMATS.includes(f0) ? f0 : old.format || 'pdf';
    const schedule = normSchedule({ ...(old.schedule || {}), ...(b.schedule || b) });
    return { schedule, channel: recipients[0].channel === 'alerts' ? 'all' : recipients[0].channel, recipients, format, saveCopy: b.saveCopy !== undefined ? !!b.saveCopy : old.saveCopy !== undefined ? old.saveCopy : true, on: b.on !== undefined ? !!b.on : old.on !== false, title: String(b.title != null ? b.title : old.title || '').slice(0, 120) };
  }
  add(b) {
    if (this.items.length >= 100) throw Object.assign(new Error('At most 100 scheduled reports.'), { status: 409 });
    const it = { id: crypto.randomBytes(6).toString('hex'), key: b.key, root: b.root || null, path: b.path, name: b.name, ...this.norm(b), createdAt: Date.now(), lastRun: null, lastError: null };
    it.nextRun = nextRun(it.schedule);
    this.items.push(it); this.save(); return this.view(it);
  }
  update(id, b) { const it = this.items.find((x) => x.id === id); if (!it) throw Object.assign(new Error('Report not found'), { status: 404 }); Object.assign(it, this.norm(b, it)); it.nextRun = nextRun(it.schedule); this.save(); return this.view(it); }
  remove(id) { this.items = this.items.filter((x) => x.id !== id); this.save(); }
  rekey(a, b) { let ch = false; for (const p of this.items) if (p.key === a || p.key.startsWith(`${a}/`)) { p.key = b + p.key.slice(a.length); if (!p.root) p.path = p.key; p.name = p.key.split('/').pop(); ch = true; } if (ch) this.save(); }
  /** the file in the report's format: as it is, or through the converter (PDF by default) */
  async render(it, f) {
    const want = !it.format || it.format === 'same' ? f.ext : it.format;
    if (want === f.ext) return { data: f.buf, ext: f.ext };
    return { data: (await this.d.convert(f.buf, f.ext, want)).data, ext: want };
  }
  async run(id) {
    const it = this.items.find((x) => x.id === id); if (!it) throw Object.assign(new Error('Report not found'), { status: 404 });
    if (this.running.has(id)) return { ok: false, running: true };
    this.running.add(id);
    try {
      const f = await this.d.readFile({ root: it.root, path: it.path });
      const out = await this.render(it, f);
      const date = new Date().toISOString().slice(0, 10);
      const name = `${it.name.replace(/\.[^.]+$/, '')} ${date}.${out.ext}`;
      let saved = null;
      if (it.saveCopy && out.ext === 'pdf') { try { saved = this.d.saveCopy({ root: it.root, path: it.path }, name, out.data); } catch (e) { this.d.log('report copy failed', e.message); } }
      const t = this.d.t;
      const msg = { event: 'report', title: it.title || t('Report: {name}', { name: it.name }), text: t('The scheduled report {name} of {date} is attached.', { name: it.name, date }), kind: 'info', url: '/ext/office', attachment: { name, data: out.data, type: MIME[out.ext] || 'application/octet-stream' } };
      // every recipient on its own – one failing channel does not stop the others
      const results = [];
      for (const r of it.recipients || normRecipients(null, it.channel)) {
        try { await this.d.notifier.send(r.channel === 'alerts' ? 'all' : r.channel, msg, { to: r.to || null }); results.push({ channel: r.channel, to: r.to, ok: true }); } catch (e) { results.push({ channel: r.channel, to: r.to, ok: false, error: String(t(e.message || String(e), e.vars)).slice(0, 200) }); }
      }
      const failed = results.filter((x) => !x.ok);
      it.lastRun = Date.now(); it.lastResults = results;
      it.lastError = failed.length ? failed.map((x) => `${x.to || x.channel}: ${x.error}`).join(' · ').slice(0, 300) : null;
      return { ok: failed.length < results.length, size: out.data.length, saved, results, error: it.lastError };
    } catch (e) { it.lastRun = Date.now(); it.lastError = String(e.message || e).slice(0, 300); this.d.log('report failed', it.name, it.lastError); return { ok: false, error: it.lastError }; } finally { this.running.delete(id); it.nextRun = nextRun(it.schedule); this.save(); }
  }
  /** due reports (missed ones while Hostwerk was off run once) */
  async tick(now = Date.now()) { for (const it of this.items) if (it.on && it.nextRun && it.nextRun <= now) await this.run(it.id); }
}

/* ---------------- library snapshots for the backup center ---------------- */
class LibBackup {
  /** deps: { dir, base(), write(rel, data), log, keep, maxBytes, yauzl } */
  constructor(deps) { this.d = { keep: 7, maxBytes: 2 * 1024 * 1048576, ...deps }; this.busy = null; this.lastSkip = null; }
  files() { try { return fs.readdirSync(this.d.dir).filter((f) => /^library-\d+\.zip$/.test(f)).sort(); } catch { return []; } }
  list() { return this.files().map((f) => { const t = Number(f.match(/(\d+)/)[1]); let size = 0; try { size = fs.statSync(path.join(this.d.dir, f)).size; } catch { /* gone */ } return { id: f, t, size }; }).sort((a, b) => b.t - a.t); }
  state() { return { keep: this.d.keep, maxBytes: this.d.maxBytes, skipped: this.lastSkip }; }
  walk(base) {
    const out = []; let total = 0;
    const go = (dir, rel, depth) => {
      let list = []; try { list = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
      for (const e of list) {
        if (P.isJunk(e.name)) continue;
        const full = path.join(dir, e.name); const r = rel ? `${rel}/${e.name}` : e.name;
        if (e.isDirectory() && depth < 30) go(full, r, depth + 1);
        else if (e.isFile()) { const st = fs.statSync(full); total += st.size; out.push({ full, rel: r, st }); }
      }
    };
    go(base, '', 0);
    return { files: out, total };
  }
  create({ force = false } = {}) {
    if (this.busy) return this.busy;
    this.busy = (async () => {
      const list = this.files(); const last = list[list.length - 1];
      if (!force && last && Date.now() - Number(last.match(/(\d+)/)[1]) < 20 * 3600000) return { skipped: 'recent' };
      const { files, total } = this.walk(this.d.base());
      if (total > this.d.maxBytes) { this.lastSkip = { at: Date.now(), total }; this.d.log('library snapshot skipped – library too large', total); return { skipped: 'too-large', total }; }
      fs.mkdirSync(this.d.dir, { recursive: true });
      const name = `library-${Date.now()}.zip`; const tmp = path.join(this.d.dir, `${name}.tmp`);
      const out = fs.createWriteStream(tmp);
      const z = new ZipWriter((c) => out.write(c));
      for (const f of files) {
        let data; try { data = await fs.promises.readFile(f.full); } catch { continue; }
        // office files are zips already – store them, deflate the rest
        if (!z.add(f.rel, data, { mtime: f.st.mtime, store: /\.(docx|xlsx|pptx|od[tspg]|zip|jpe?g|png|pdf|mp4)$/i.test(f.rel) })) await new Promise((ok) => out.once('drain', ok));
        await new Promise((ok) => setImmediate(ok));
      }
      z.end();
      await new Promise((ok, bad) => { out.on('error', bad); out.end(ok); });
      fs.renameSync(tmp, path.join(this.d.dir, name));
      for (const old of this.files().slice(0, Math.max(0, this.files().length - this.d.keep))) fs.rmSync(path.join(this.d.dir, old), { force: true });
      this.lastSkip = null;
      return { id: name, files: files.length, size: fs.statSync(path.join(this.d.dir, name)).size };
    })().finally(() => { this.busy = null; });
    return this.busy;
  }
  /** restore (all or some paths): files that differ are written through the library (current content → version) */
  async restore(id, { paths = null } = {}) {
    if (!/^library-\d+\.zip$/.test(String(id))) throw Object.assign(new Error('Backup not found'), { status: 404 });
    const file = path.join(this.d.dir, id);
    if (!fs.existsSync(file)) throw Object.assign(new Error('Backup not found'), { status: 404 });
    const want = paths ? new Set(paths.map((p) => P.normRel(p))) : null;
    let restored = 0; let same = 0;
    const one = async (name, data) => {
      let rel; try { rel = P.normRel(name); } catch { return; }
      if (!rel || name.endsWith('/') || (want && !want.has(rel))) return;
      const full = P.resolveIn(this.d.base(), rel);
      try { const cur = fs.readFileSync(full); if (cur.equals(data)) { same++; return; } } catch { /* new */ }
      this.d.write(rel, data); restored++;
      await new Promise((ok) => setImmediate(ok));
    };
    if (this.d.yauzl) {
      await new Promise((ok, bad) => {
        this.d.yauzl.open(file, { lazyEntries: true, decodeStrings: true }, (e, zf) => {
          if (e) return bad(e);
          zf.on('error', bad); zf.on('end', ok);
          zf.on('entry', (en) => {
            if (/\/$/.test(en.fileName) || (want && !want.has(en.fileName))) { zf.readEntry(); return; }
            zf.openReadStream(en, (e2, rs) => { if (e2) return bad(e2); const parts = []; rs.on('data', (c) => parts.push(c)); rs.on('error', bad); rs.on('end', () => { one(en.fileName, Buffer.concat(parts)).then(() => zf.readEntry(), bad); }); });
          });
          zf.readEntry();
        });
      });
    } else {
      if (fs.statSync(file).size > 512 * 1048576) throw Object.assign(new Error('This backup is too large to restore here.'), { status: 413 });
      for (const [n, e] of readZip(fs.readFileSync(file))) if (!e.dir) await one(n, e.read());
    }
    return { ok: true, restored, unchanged: same };
  }
  remove(id) { if (/^library-\d+\.zip$/.test(String(id))) fs.rmSync(path.join(this.d.dir, id), { force: true }); }
  /** central directory of a snapshot without reading the whole zip → Map(name → { size, csize, method, off, crc }) */
  index(id) {
    if (!/^library-\d+\.zip$/.test(String(id))) return new Map();
    const cache = this.idxCache || (this.idxCache = new Map());
    if (cache.has(id)) return cache.get(id);
    const file = path.join(this.d.dir, id); let fd;
    try { fd = fs.openSync(file, 'r'); } catch { return new Map(); }
    try {
      const size = fs.fstatSync(fd).size; const tail = Buffer.alloc(Math.min(size, 65557)); fs.readSync(fd, tail, 0, tail.length, size - tail.length);
      let e = -1; for (let i = tail.length - 22; i >= 0; i--) if (tail.readUInt32LE(i) === 0x06054b50) { e = i; break; }
      if (e < 0) return new Map();
      const cdSize = tail.readUInt32LE(e + 12); const cdOff = tail.readUInt32LE(e + 16);
      if (cdSize > 64 * 1048576) return new Map();
      const cd = Buffer.alloc(cdSize); fs.readSync(fd, cd, 0, cdSize, cdOff);
      const out = new Map();
      for (let p = 0; p + 46 <= cd.length && cd.readUInt32LE(p) === 0x02014b50;) {
        const nl = cd.readUInt16LE(p + 28); const xl = cd.readUInt16LE(p + 30); const cl = cd.readUInt16LE(p + 32);
        const name = cd.toString('utf8', p + 46, p + 46 + nl);
        out.set(name, { method: cd.readUInt16LE(p + 10), crc: cd.readUInt32LE(p + 16), csize: cd.readUInt32LE(p + 20), size: cd.readUInt32LE(p + 24), off: cd.readUInt32LE(p + 42) });
        p += 46 + nl + xl + cl;
      }
      if (cache.size > 20) cache.clear();
      cache.set(id, out);
      return out;
    } finally { fs.closeSync(fd); }
  }
  /** snapshots that contain rel: [{ id, t, size (of the file in it), changed (≠ the previous one) }] */
  versionsOf(rel) {
    const out = []; let prev = null;
    for (const b of this.list().reverse()) { const e = this.index(b.id).get(rel); if (!e) continue; out.push({ id: b.id, time: b.t, size: e.size, label: prev && prev.crc === e.crc ? null : 'changed' }); prev = e; }
    return out.reverse();
  }
  readFile(id, rel) {
    const e = this.index(id).get(rel); if (!e) return null;
    if (e.size > 200 * 1048576) throw Object.assign(new Error('The file is too large.'), { status: 413 });
    const fd = fs.openSync(path.join(this.d.dir, id), 'r');
    try {
      const h = Buffer.alloc(30); fs.readSync(fd, h, 0, 30, e.off);
      if (h.readUInt32LE(0) !== 0x04034b50) throw new Error('Broken backup');
      const data = Buffer.alloc(e.csize); fs.readSync(fd, data, 0, e.csize, e.off + 30 + h.readUInt16LE(26) + h.readUInt16LE(28));
      return e.method === 8 ? require('zlib').inflateRawSync(data, { maxOutputLength: e.size || 1 }) : data;
    } finally { fs.closeSync(fd); }
  }
}

module.exports = { Thumbs, Publish, Reports, LibBackup, nextRun, normSchedule, normRecipients };
