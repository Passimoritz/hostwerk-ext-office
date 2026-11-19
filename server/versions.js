'use strict';
/*
 * Versions: earlier contents of a file, kept in the data folder (never next to the document).
 *   <dir>/index.json         { keys: { <file key>: <folder id> } }   – file key = library path or "@root/path"
 *   <dir>/<folder id>/list.json [{ id, time, size, label, auto, hash }] + <id>.bin
 * Renames/moves only change the key (rekey), the stored versions stay. Pruning (pruneList) keeps named versions,
 * drops unnamed ones older than versionDays and everything beyond keepVersions.
 * Autosave would create a version every few seconds – automatic versions are therefore at least `gapMs` apart.
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { readJson, writeJson, newId } = require('./store');

const DAY = 86400000;

/** Which versions to drop → [id]. list: [{ id, time, label }]. Named versions are never dropped. */
function pruneList(list, { keep = 30, days = 90, now = Date.now() } = {}) {
  const drop = new Set();
  const unnamed = (list || []).filter((v) => !v.label).sort((a, b) => b.time - a.time);
  unnamed.forEach((v, i) => {
    if (days > 0 && now - v.time > days * DAY) drop.add(v.id);
    else if (i >= Math.max(0, keep)) drop.add(v.id);
  });
  return [...drop];
}

const sha1 = (buf) => crypto.createHash('sha1').update(buf).digest('hex');
/** Prefix match for keys: the key itself or (folders) everything below it */
const under = (key, prefix) => key === prefix || key.startsWith(`${prefix}/`);

class Versions {
  constructor(dir) { this.dir = dir; fs.mkdirSync(dir, { recursive: true }); this.idxFile = path.join(dir, 'index.json'); }
  index() { const i = readJson(this.idxFile, null); return i && i.keys ? i : { keys: {} }; }
  saveIndex(i) { writeJson(this.idxFile, i); }
  folderOf(key, create = false) {
    const i = this.index();
    let f = i.keys[key];
    if (!f && create) { f = newId(); i.keys[key] = f; this.saveIndex(i); }
    return f ? path.join(this.dir, f) : null;
  }
  list(key) {
    const f = this.folderOf(key);
    if (!f) return [];
    return readJson(path.join(f, 'list.json'), []).sort((a, b) => b.time - a.time);
  }
  writeList(key, list) { const f = this.folderOf(key, true); fs.mkdirSync(f, { recursive: true }); writeJson(path.join(f, 'list.json'), list); }
  file(key, id) {
    if (!/^[a-z0-9]{6,40}$/.test(String(id || ''))) return null;
    const f = this.folderOf(key);
    if (!f) return null;
    const p = path.join(f, `${id}.bin`);
    return fs.existsSync(p) ? p : null;
  }
  get(key, id) { return this.list(key).find((v) => v.id === id) || null; }

  /**
   * Keep bytes as a version. opts: { label, auto, gapMs, keep, days, time }
   * → the new version, or null (same content as the newest version, or too soon after the last automatic one)
   */
  add(key, data, opts = {}) {
    const buf = Buffer.isBuffer(data) ? data : Buffer.from(data);
    if (opts.keep === 0 && !opts.label) return null;
    const list = this.list(key);
    const hash = sha1(buf);
    const last = list[0];
    const now = opts.time || Date.now();
    if (last && last.hash === hash) {
      if (opts.label && !last.label) { last.label = String(opts.label).slice(0, 120); last.auto = false; this.writeList(key, list); return last; }
      return opts.label ? last : null;
    }
    if (!opts.label && opts.auto !== false && opts.gapMs && list.some((v) => v.auto && now - v.time < opts.gapMs)) return null;
    const v = { id: newId(), time: now, size: buf.length, label: opts.label ? String(opts.label).slice(0, 120) : '', auto: opts.auto !== false && !opts.label, hash };
    const f = this.folderOf(key, true);
    fs.mkdirSync(f, { recursive: true });
    fs.writeFileSync(path.join(f, `${v.id}.bin`), buf);
    list.unshift(v);
    this.writeList(key, list);
    if (opts.keep != null || opts.days != null) this.prune(key, { keep: opts.keep, days: opts.days, now });
    return v;
  }
  setLabel(key, id, label) {
    const list = this.list(key);
    const v = list.find((x) => x.id === id);
    if (!v) return null;
    v.label = String(label || '').trim().slice(0, 120);
    v.auto = v.auto && !v.label;
    this.writeList(key, list);
    return v;
  }
  remove(key, id) {
    const list = this.list(key);
    const p = this.file(key, id);
    if (p) fs.rmSync(p, { force: true });
    this.writeList(key, list.filter((v) => v.id !== id));
  }
  prune(key, opts) {
    const list = this.list(key);
    const drop = pruneList(list, opts);
    if (!drop.length) return 0;
    const f = this.folderOf(key);
    for (const id of drop) fs.rmSync(path.join(f, `${id}.bin`), { force: true });
    this.writeList(key, list.filter((v) => !drop.includes(v.id)));
    return drop.length;
  }
  pruneAll(opts) {
    let n = 0;
    for (const key of Object.keys(this.index().keys)) { try { n += this.prune(key, opts); } catch { /* next */ } }
    // empty folders and keys without versions
    const i = this.index();
    for (const [key, f] of Object.entries(i.keys)) {
      const list = readJson(path.join(this.dir, f, 'list.json'), []);
      if (!list.length) { fs.rmSync(path.join(this.dir, f), { recursive: true, force: true }); delete i.keys[key]; }
    }
    this.saveIndex(i);
    return n;
  }
  /** Key changed (rename/move/trash): the file itself and – for folders – everything below it */
  rekey(oldKey, newKey) {
    const i = this.index();
    let changed = false;
    for (const k of Object.keys(i.keys)) {
      if (!under(k, oldKey)) continue;
      const nk = newKey + k.slice(oldKey.length);
      if (i.keys[nk] && nk !== k) fs.rmSync(path.join(this.dir, i.keys[nk]), { recursive: true, force: true }); // target had old versions
      i.keys[nk] = i.keys[k];
      if (nk !== k) delete i.keys[k];
      changed = true;
    }
    if (changed) this.saveIndex(i);
  }
  /** Copy (file copy): versions are not copied – a copy starts fresh. Drop all versions of key (and below). */
  drop(key) {
    const i = this.index();
    for (const k of Object.keys(i.keys)) if (under(k, key)) { fs.rmSync(path.join(this.dir, i.keys[k]), { recursive: true, force: true }); delete i.keys[k]; }
    this.saveIndex(i);
  }
  /** Total size in bytes (status/backups) */
  size() {
    let n = 0;
    for (const f of Object.values(this.index().keys)) for (const v of readJson(path.join(this.dir, f, 'list.json'), [])) n += v.size || 0;
    return n;
  }
}

/** Public view of a version (no hash) */
const view = (v) => ({ id: v.id, time: v.time, size: v.size, label: v.label || '', auto: !!v.auto });

module.exports = { Versions, pruneList, view, under };
