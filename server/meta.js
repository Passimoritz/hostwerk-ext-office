'use strict';
/*
 * Metadata of files (favorite, tags, color, note), the "recently opened" list and soft locks.
 * Stored in <data>/meta.json – keyed by the file key (library path, or "@<root>/<path>" for files from Hostwerk's roots),
 * so renames/moves only change the key (rekey) and the metadata follows the file.
 * Locks live in memory only (they end with Hostwerk) and expire when an editor forgets to unlock.
 */
const { readJson, writeJson } = require('./store');
const { under } = require('./versions');

const COLORS = ['', 'red', 'orange', 'yellow', 'green', 'teal', 'blue', 'purple', 'pink', 'gray'];
// soft locks are renewed every 2 minutes by the open editor – a closed or reloaded window must not keep the file
// "open" for long (it used to be 12 h, so every reload left a stale "open" badge and a warning in the next editor)
const LOCK_MS = 5 * 60000;

/** Validated metadata patch → clean object (throws 400 on nonsense) */
function cleanPatch(p) {
  const out = {};
  const bad = (m) => Object.assign(new Error(m), { status: 400 });
  if (p.favorite !== undefined) out.favorite = !!p.favorite;
  if (p.tags !== undefined) {
    const list = Array.isArray(p.tags) ? p.tags : String(p.tags || '').split(',');
    out.tags = [...new Set(list.map((t) => String(t).trim().replace(/\s+/g, ' ').slice(0, 40)).filter(Boolean))].slice(0, 30);
  }
  if (p.color !== undefined) {
    const c = String(p.color || '');
    if (!COLORS.includes(c) && !/^#[0-9a-f]{6}$/i.test(c)) throw bad('Unknown color');
    out.color = c;
  }
  if (p.note !== undefined) out.note = String(p.note || '').slice(0, 4000);
  return out;
}

class Meta {
  constructor(file) {
    this.file = file;
    this.data = readJson(file, null) || {};
    if (!this.data.files || typeof this.data.files !== 'object') this.data.files = {};
    if (!Array.isArray(this.data.recent)) this.data.recent = [];
    this.locks = new Map();
    this.seqs = new Map(); // editor (sid) → newest lock/unlock request seen { seq, t }
    this.timer = null;
  }
  save(now = false) {
    if (now) { clearTimeout(this.timer); this.timer = null; writeJson(this.file, this.data); return; }
    if (this.timer) return;
    this.timer = setTimeout(() => { this.timer = null; try { writeJson(this.file, this.data); } catch { /* next time */ } }, 300);
    if (this.timer.unref) this.timer.unref();
  }
  flush() { if (this.timer) this.save(true); }
  get(key) { return this.data.files[key] || null; }
  set(key, patch) {
    const cur = { favorite: false, tags: [], color: '', note: '', ...(this.data.files[key] || {}), ...cleanPatch(patch || {}) };
    if (!cur.favorite && !cur.tags.length && !cur.color && !cur.note) delete this.data.files[key];
    else this.data.files[key] = cur;
    this.save();
    return cur;
  }
  /** Entry fields for a key */
  fields(key) {
    const m = this.data.files[key] || {};
    const l = this.lockOf(key);
    return { favorite: !!m.favorite, tags: m.tags || [], color: m.color || '', note: m.note || '', lock: l ? { by: l.by, since: l.since } : null };
  }
  favorites() { return Object.entries(this.data.files).filter(([, m]) => m.favorite).map(([k]) => k); }
  /** All tags with count: [{ tag, count }] */
  tags() {
    const n = new Map();
    for (const m of Object.values(this.data.files)) for (const t of m.tags || []) n.set(t, (n.get(t) || 0) + 1);
    return [...n.entries()].map(([tag, count]) => ({ tag, count })).sort((a, b) => b.count - a.count || a.tag.localeCompare(b.tag));
  }
  rekey(oldKey, newKey) {
    let changed = false;
    for (const k of Object.keys(this.data.files)) {
      if (!under(k, oldKey)) continue;
      this.data.files[newKey + k.slice(oldKey.length)] = this.data.files[k];
      if (newKey + k.slice(oldKey.length) !== k) delete this.data.files[k];
      changed = true;
    }
    for (const r of this.data.recent) if (under(r.key, oldKey)) { r.key = newKey + r.key.slice(oldKey.length); if (!r.root) r.path = r.key; changed = true; }
    for (const [k, l] of [...this.locks]) if (under(k, oldKey)) { this.locks.delete(k); this.locks.set(newKey + k.slice(oldKey.length), l); }
    if (changed) this.save();
  }
  drop(key) {
    for (const k of Object.keys(this.data.files)) if (under(k, key)) delete this.data.files[k];
    this.data.recent = this.data.recent.filter((r) => !under(r.key, key));
    this.save();
  }
  /** Record an open: { key, root, path } */
  opened(loc, now = Date.now()) {
    this.data.recent = [{ key: loc.key, root: loc.root || null, path: loc.path, t: now }, ...this.data.recent.filter((r) => r.key !== loc.key)].slice(0, 200);
    this.save();
  }
  recent(limit = 20) { return this.data.recent.slice(0, limit); }
  forget(key) { this.data.recent = this.data.recent.filter((r) => r.key !== key); this.save(); }

  /*
   * Soft locks: { by, since, until, client, sid }. client = one browser window (Files hides the marks of its own window –
   * an editor in that window is closed when Files shows), sid = one open editor; seq orders the requests of one editor:
   * a heartbeat that arrives after the unlock (several connections, a request still on its way) never brings the lock
   * back, and the late unlock of a closed editor never releases the lock of the editor that opened the file again.
   */
  /** true = an older request of this editor than one already seen (ignore it) */
  staleSeq(sid, seq, now = Date.now()) {
    if (!sid || !Number.isFinite(seq)) return false;
    const s = this.seqs.get(sid);
    if (s && seq <= s.seq) return true;
    this.seqs.set(sid, { seq, t: now });
    if (this.seqs.size > 500) for (const [k, v] of this.seqs) if (now - v.t > 2 * LOCK_MS) this.seqs.delete(k);
    return false;
  }
  /** → { by, since, client, renewed } – renewed: the same editor only renewed it (nothing changed for the marks) */
  lock(key, by, now = Date.now(), o = {}) {
    const who = String(by || 'Hostwerk').slice(0, 80);
    const client = o.client ? String(o.client).slice(0, 64) : null;
    const sid = o.sid ? String(o.sid).slice(0, 64) : null;
    const prev = this.locks.get(key);
    const cur = this.lockOf(key, now);
    if (sid && this.staleSeq(sid, Number(o.seq), now)) return cur ? { ...cur, ignored: true } : { by: who, since: now, client, ignored: true, released: true };
    const l = { by: who, since: cur && cur.by === who ? cur.since : now, until: now + LOCK_MS, client, sid };
    this.locks.set(key, l);
    return { by: l.by, since: l.since, client, renewed: !!(cur && prev && prev.sid === sid && prev.client === client) };
  }
  /** → true when a lock was released (o.sid: only the lock of that editor – another one may hold the file by now) */
  unlock(key, o = {}, now = Date.now()) {
    const sid = o && o.sid ? String(o.sid).slice(0, 64) : null;
    if (sid && this.staleSeq(sid, Number(o.seq), now)) return false;
    const l = this.locks.get(key);
    if (!l) return false;
    if (sid && l.sid && l.sid !== sid && l.until >= now) return false;
    this.locks.delete(key);
    return true;
  }
  lockOf(key, now = Date.now()) {
    const l = this.locks.get(key);
    if (!l) return null;
    if (l.until < now) { this.locks.delete(key); return null; }
    return { by: l.by, since: l.since, until: l.until, client: l.client || null };
  }
}

module.exports = { Meta, cleanPatch, COLORS };
