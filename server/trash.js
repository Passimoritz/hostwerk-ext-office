'use strict';
/*
 * Trash of the library: deleted files/folders move to <data>/trash/<id>/<name> (same drive as the data folder – a move
 * across drives copies). index: <data>/trash.json [{ id, name, from, deletedAt, dir, size }]. Restoring picks a free name
 * when the original place is taken. Auto-purge after trashDays.
 */
const fs = require('fs');
const path = require('path');
const { readJson, writeJson, newId } = require('./store');
const { uniqueName } = require('./paths');

function moveAny(a, b) {
  try { fs.renameSync(a, b); } catch (e) {
    if (!['EXDEV', 'EPERM'].includes(e.code)) throw e;
    fs.cpSync(a, b, { recursive: true, errorOnExist: true, force: false, preserveTimestamps: true });
    fs.rmSync(a, { recursive: true, force: true });
  }
}
function sizeOf(p) {
  try {
    const st = fs.statSync(p);
    if (!st.isDirectory()) return st.size;
    let n = 0;
    for (const e of fs.readdirSync(p)) n += sizeOf(path.join(p, e));
    return n;
  } catch { return 0; }
}

/** Ids of trash entries older than days → [id] */
function expired(list, days, now = Date.now()) {
  if (!(days > 0)) return [];
  return (list || []).filter((e) => now - e.deletedAt > days * 86400000).map((e) => e.id);
}

class Trash {
  constructor(dir) { this.dir = dir; this.file = path.join(dir, 'trash.json'); fs.mkdirSync(dir, { recursive: true }); }
  list() { return readJson(this.file, []).sort((a, b) => b.deletedAt - a.deletedAt); }
  save(list) { writeJson(this.file, list); }
  /** Move full (inside the library, relative path rel) to the trash → entry */
  put(full, rel, now = Date.now()) {
    const st = fs.statSync(full);
    const id = newId();
    const name = path.basename(full);
    const box = path.join(this.dir, id);
    fs.mkdirSync(box, { recursive: true });
    const size = st.isDirectory() ? sizeOf(full) : st.size;
    moveAny(full, path.join(box, name));
    const e = { id, name, from: rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '', path: rel, deletedAt: now, dir: st.isDirectory(), size, mtime: st.mtimeMs };
    const list = this.list(); list.push(e); this.save(list);
    return e;
  }
  get(id) { return this.list().find((e) => e.id === id) || null; }
  /**
   * Back to its folder (or the library root if that is gone). resolve(relDir) → absolute folder.
   * → { entry, rel } (rel = new relative path)
   */
  restore(id, resolveDir, existsDir) {
    const e = this.get(id);
    if (!e) throw Object.assign(new Error('Not found in the trash'), { status: 404 });
    const from = existsDir(e.from) ? e.from : '';
    const dir = resolveDir(from);
    fs.mkdirSync(dir, { recursive: true });
    const name = uniqueName(e.name, (n) => fs.existsSync(path.join(dir, n)), e.dir);
    moveAny(path.join(this.dir, id, e.name), path.join(dir, name));
    fs.rmSync(path.join(this.dir, id), { recursive: true, force: true });
    this.save(this.list().filter((x) => x.id !== id));
    return { entry: e, rel: from ? `${from}/${name}` : name };
  }
  /** Delete for good (ids missing = everything) → removed entries */
  purge(ids) {
    const list = this.list();
    const drop = ids ? list.filter((e) => ids.includes(e.id)) : list;
    for (const e of drop) fs.rmSync(path.join(this.dir, e.id), { recursive: true, force: true });
    this.save(list.filter((e) => !drop.includes(e)));
    return drop;
  }
  purgeOld(days, now = Date.now()) { const ids = expired(this.list(), days, now); return ids.length ? this.purge(ids) : []; }
  file(id) { const e = this.get(id); return e ? path.join(this.dir, id, e.name) : null; }
}

module.exports = { Trash, expired, moveAny, sizeOf };
