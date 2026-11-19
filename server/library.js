'use strict';
/*
 * The Office library: files below settings.libraryDir (default <Hostwerk data>/office) plus files from Hostwerk's file
 * browser roots (root + path, resolved through Hostwerk's files core). Everything that changes files goes through here:
 * listing, folders, rename/move/copy with free names, trash, raw read/write with conflict check and versions, create.
 * Metadata, versions and the trash live in the data folder and follow renames/moves (rekey).
 */
const fs = require('fs');
const path = require('path');
const P = require('./paths');
const { FORMATS, extOf } = require('./formats');
const { writeFileAtomic, readJson, writeJson, newId } = require('./store');
const { moveAny } = require('./trash');

const fail = P.fail;
const VERSION_GAP = 5 * 60000; // automatic versions at least 5 minutes apart (autosave would create one every few seconds)

class Library {
  /** deps: { settings(), files (Hostwerk files core or null), meta, versions, trash, index, sharedKeys(), dataDir, onChange(rel) } */
  constructor(deps) { this.d = deps; }
  get base() { const b = this.d.settings().libraryDir; fs.mkdirSync(b, { recursive: true }); return b; }

  /* ---------- locations ---------- */
  /** { path } or { root, path } → { key, root, rel, full } */
  loc(q) {
    const root = q && q.root ? String(q.root) : null;
    if (root) {
      const files = this.d.files;
      if (!files || typeof files.resolve !== 'function') throw fail('Hostwerk\'s file roots are not available.', 501);
      const rel = String(q.path || '').replace(/\\/g, '/').split('/').filter((x) => x && x !== '.').join('/');
      if (rel.split('/').includes('..')) throw fail('Invalid path', 403);
      const { full } = files.resolve(root, rel);
      // the library itself as a root of Hostwerk's Files (host.files.addRoot, "Office documents") or any root that
      // contains the library: the same file as a library path – favourites, versions, shares and the index stay one
      const base = this.d.settings().libraryDir;
      if (base && P.inside(base, full)) {
        const lr = path.relative(path.resolve(base), path.resolve(full)).split(path.sep).filter(Boolean).join('/');
        return { key: lr, root: null, rel: lr, full };
      }
      return { key: `@${root}/${rel}`, root, rel, full };
    }
    const rel = P.normRel(q && q.path);
    return { key: rel, root: null, rel, full: P.resolveIn(this.base, rel) };
  }
  /** Existing file location (throws 404) */
  file(q) {
    const l = this.loc(q);
    let st; try { st = fs.statSync(l.full); } catch { throw fail('File not found', 404); }
    if (st.isDirectory()) throw fail('This is a folder, not a file.', 400);
    return { ...l, st };
  }
  /** Folder location: library rel ('' = library) or { root, dir } of Hostwerk's Files */
  dirLoc(rel) {
    const l = rel && typeof rel === 'object' ? this.loc({ root: rel.root, path: rel.dir || rel.path || '' }) : this.loc({ path: rel });
    let st; try { st = fs.statSync(l.full); } catch { if (!l.rel) { fs.mkdirSync(l.full, { recursive: true }); return l; } throw fail('Folder not found', 404); }
    if (!st.isDirectory()) throw fail('This is a file, not a folder.', 400);
    return l;
  }

  /* ---------- entries ---------- */
  entry(l, st) {
    const s = st || fs.statSync(l.full);
    const name = path.basename(l.full);
    const dir = s.isDirectory();
    const ext = dir ? '' : extOf(name);
    const f = FORMATS[ext];
    const idx = !l.root && this.d.index ? this.d.index.get(l.rel) : null;
    return {
      path: l.rel, ...(l.root ? { root: l.root } : {}), name, dir, ext, kind: dir ? null : f ? f.kind : null, family: dir ? null : f ? f.family : null,
      size: dir ? 0 : s.size, mtime: Math.floor(s.mtimeMs), ctime: Math.floor(s.birthtimeMs || s.ctimeMs),
      ...this.d.meta.fields(l.key), shared: this.d.sharedKeys().has(l.key), thumb: !!(idx && idx.thumb && idx.mtime === Math.floor(s.mtimeMs)),
    };
  }
  entryOf(q) { const l = this.loc(q); return this.entry(l); }
  crumbs(rel) {
    const parts = rel ? rel.split('/') : [];
    return [{ name: 'Library', path: '' }, ...parts.map((n, i) => ({ name: n, path: parts.slice(0, i + 1).join('/') }))];
  }
  /** All entries of a folder (optionally recursive for q) */
  listDir(rel, { recursive = false, limit = 5000 } = {}) {
    const out = [];
    const go = (r, depth) => {
      const l = this.dirLoc(r);
      let list = [];
      try { list = fs.readdirSync(l.full, { withFileTypes: true }); } catch { return; }
      for (const e of list) {
        if (P.isJunk(e.name) || out.length >= limit) continue;
        const cr = r ? `${r}/${e.name}` : e.name;
        try { const st = fs.statSync(path.join(l.full, e.name)); out.push(this.entry({ key: cr, root: null, rel: cr, full: path.join(l.full, e.name) }, st)); if (recursive && st.isDirectory() && depth < 20) go(cr, depth + 1); } catch { /* broken link */ }
      }
    };
    go(rel, 0);
    return out;
  }
  /** Entries for keys (favorites, recent, shared) – missing files are skipped */
  fromKeys(items) {
    const out = [];
    for (const it of items) {
      try {
        const q = typeof it === 'string' ? (it.startsWith('@') ? { root: it.slice(1, it.indexOf('/')), path: it.slice(it.indexOf('/') + 1) } : { path: it }) : it;
        const l = this.loc(q);
        const st = fs.statSync(l.full);
        out.push({ ...this.entry(l, st), ...(it && it.t ? { opened: it.t } : {}) });
      } catch { /* gone */ }
    }
    return out;
  }
  static sortEntries(list, sort = 'modified') {
    const desc = String(sort).startsWith('-');
    const key = String(sort).replace(/^-/, '');
    const cmpName = (a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
    const by = {
      name: cmpName, modified: (a, b) => b.mtime - a.mtime, created: (a, b) => b.ctime - a.ctime, size: (a, b) => b.size - a.size,
      kind: (a, b) => String(a.kind || '').localeCompare(String(b.kind || '')) || cmpName(a, b), type: (a, b) => String(a.ext).localeCompare(String(b.ext)) || cmpName(a, b),
      opened: (a, b) => (b.opened || 0) - (a.opened || 0), deleted: (a, b) => (b.deletedAt || 0) - (a.deletedAt || 0),
    }[key] || ((a, b) => b.mtime - a.mtime);
    return [...list].sort((a, b) => (a.dir !== b.dir ? (a.dir ? -1 : 1) : desc ? -by(a, b) : by(a, b)));
  }
  static filter(list, { kind, family, tag, q, ext } = {}) {
    const ql = String(q || '').toLowerCase().trim();
    return list.filter((e) => (!kind || e.dir || e.kind === kind) && (!family || e.dir || e.family === family) && (!ext || e.ext === ext) && (!tag || (e.tags || []).includes(tag)) && (!ql || e.name.toLowerCase().includes(ql)));
  }

  /* ---------- file operations (library only) ---------- */
  mkdir(dir, name) {
    const d = this.dirLoc(P.normRel(dir));
    const n = P.uniqueName(P.checkName(name), (x) => fs.existsSync(path.join(d.full, x)), true);
    fs.mkdirSync(path.join(d.full, n));
    const rel = P.joinRel(d.rel, n);
    this.changed(rel);
    return this.entryOf({ path: rel });
  }
  /** rel0: library path or { root, path } of Hostwerk's Files (metadata, versions and shares follow) */
  rename(rel0, name) {
    const l = this.loc(rel0 && typeof rel0 === 'object' ? rel0 : { path: rel0 });
    if (!l.rel) throw fail('The library itself cannot be renamed.', 400);
    if (!fs.existsSync(l.full)) throw fail('File not found', 404);
    const n = P.checkName(name);
    const dir = path.dirname(l.full);
    const parent = l.rel.includes('/') ? l.rel.slice(0, l.rel.lastIndexOf('/')) : '';
    if (n === path.basename(l.full)) return this.entry(l);
    // same name with other case (Windows): rename in two steps
    const target = path.join(dir, n);
    if (fs.existsSync(target) && target.toLowerCase() !== l.full.toLowerCase()) throw fail('A file or folder with this name already exists.', 409);
    if (target.toLowerCase() === l.full.toLowerCase() && target !== l.full) { const tmp = `${l.full}.~ren${Date.now()}`; fs.renameSync(l.full, tmp); fs.renameSync(tmp, target); } else fs.renameSync(l.full, target);
    const rel = P.joinRel(parent, n);
    if (l.root) { const nk = `@${l.root}/${rel}`; this.rekey(l.key, nk); return this.entryOf({ root: l.root, path: rel }); }
    this.rekey(l.rel, rel);
    this.changed(rel);
    return this.entryOf({ path: rel });
  }
  move(paths, to) {
    const d = this.dirLoc(P.normRel(to));
    const out = [];
    for (const p of [].concat(paths || [])) {
      const l = this.loc({ path: p });
      if (!l.rel || !fs.existsSync(l.full)) continue;
      if (d.rel === l.rel || d.rel.startsWith(`${l.rel}/`)) throw fail('A folder cannot be moved into itself.', 400);
      if (path.dirname(l.full) === d.full) { out.push(this.entry(l)); continue; }
      const isDir = fs.statSync(l.full).isDirectory();
      const n = P.uniqueName(path.basename(l.full), (x) => fs.existsSync(path.join(d.full, x)), isDir);
      moveAny(l.full, path.join(d.full, n));
      const rel = P.joinRel(d.rel, n);
      this.rekey(l.rel, rel);
      out.push(this.entryOf({ path: rel }));
    }
    this.changed(d.rel);
    return out;
  }
  copy(paths, to) {
    const d = this.dirLoc(P.normRel(to));
    const out = [];
    for (const p of [].concat(paths || [])) {
      const l = this.loc({ path: p });
      if (!l.rel || !fs.existsSync(l.full)) continue;
      if (d.rel === l.rel || d.rel.startsWith(`${l.rel}/`)) throw fail('A folder cannot be copied into itself.', 400);
      const isDir = fs.statSync(l.full).isDirectory();
      const n = P.uniqueName(path.basename(l.full), (x) => fs.existsSync(path.join(d.full, x)), isDir);
      fs.cpSync(l.full, path.join(d.full, n), { recursive: true, errorOnExist: true, force: false, preserveTimestamps: false });
      out.push(this.entryOf({ path: P.joinRel(d.rel, n) }));
    }
    this.changed(d.rel);
    return out;
  }
  toTrash(paths) {
    const out = [];
    for (const p of [].concat(paths || [])) {
      const l = this.loc({ path: p });
      if (!l.rel || !fs.existsSync(l.full)) continue;
      this.d.meta.unlock(l.key);
      const e = this.d.trash.put(l.full, l.rel);
      this.rekey(l.rel, `#trash/${e.id}`);
      out.push(e);
    }
    this.changed('');
    return out;
  }
  restore(ids) {
    const out = [];
    for (const id of [].concat(ids || [])) {
      const r = this.d.trash.restore(String(id), (rel) => this.loc({ path: rel }).full, (rel) => { try { return fs.statSync(this.loc({ path: rel }).full).isDirectory(); } catch { return false; } });
      this.rekey(`#trash/${id}`, r.rel);
      out.push(this.entryOf({ path: r.rel }));
    }
    this.changed('');
    return out;
  }
  purge(ids) {
    const gone = this.d.trash.purge(ids ? [].concat(ids).map(String) : undefined);
    for (const e of gone) { this.d.meta.drop(`#trash/${e.id}`); this.d.versions.drop(`#trash/${e.id}`); }
    return gone.length;
  }
  rekey(a, b) { this.d.meta.rekey(a, b); this.d.versions.rekey(a, b); if (this.d.onRekey) this.d.onRekey(a, b); }
  changed(rel) { if (this.d.onChange) this.d.onChange(rel); }

  /* ---------- raw ---------- */
  read(q) { const l = this.file(q); return { ...l, mtime: Math.floor(l.st.mtimeMs) }; }
  /**
   * Save bytes. opts: { mtime, force, label } → { ok, mtime, size, version }
   * 409 { mtime } when the file changed on disk since mtime (unless force).
   */
  write(q, data, { mtime, force = false, label = '' } = {}) {
    const l = this.loc(q);
    const s = this.d.settings().general;
    let st = null; try { st = fs.statSync(l.full); } catch { /* new file */ }
    if (st && st.isDirectory()) throw fail('This is a folder, not a file.', 400);
    if (st && !force && mtime != null && mtime !== '' && Math.abs(Math.floor(st.mtimeMs) - Number(mtime)) > 1) throw Object.assign(fail('The file was changed by someone else in the meantime.', 409), { extra: { mtime: Math.floor(st.mtimeMs) } });
    let version = null;
    if (st && s.keepVersions > 0) {
      try {
        const prev = fs.readFileSync(l.full);
        const v = this.d.versions.add(l.key, prev, { gapMs: VERSION_GAP, time: Math.floor(st.mtimeMs), keep: s.keepVersions, days: s.versionDays });
        if (v) version = v.id;
      } catch (e) { if (e.code !== 'ENOENT') throw e; }
    }
    fs.mkdirSync(path.dirname(l.full), { recursive: true });
    writeFileAtomic(l.full, data);
    if (label) { const v = this.d.versions.add(l.key, data, { label, auto: false }); if (v) version = v.id; }
    const now = fs.statSync(l.full);
    if (!l.root) this.changed(l.rel);
    return { ok: true, mtime: Math.floor(now.mtimeMs), size: now.size, version };
  }
  /** New file with a free name in dir → Entry */
  create(dir, name, data, { keep = false } = {}) {
    const d = this.dirLoc(dir && typeof dir === 'object' ? dir : P.normRel(dir));
    const n0 = keep ? P.checkName(name) : P.safeName(name || 'Untitled');
    const n = P.uniqueName(n0, (x) => fs.existsSync(path.join(d.full, x)));
    const full = path.join(d.full, n);
    fs.writeFileSync(full, data || Buffer.alloc(0), { flag: 'wx' });
    const rel = P.joinRel(d.rel, n);
    if (d.root) return this.entryOf({ root: d.root, path: rel });
    this.changed(rel);
    return this.entryOf({ path: rel });
  }
  /** Restore a version: current content becomes a version first */
  restoreVersion(q, id) {
    const l = this.file(q);
    const src = this.d.versions.file(l.key, id);
    if (!src) throw fail('Version not found', 404);
    const data = fs.readFileSync(src);
    const v = this.d.versions.get(l.key, id);
    this.d.versions.add(l.key, fs.readFileSync(l.full), { label: `Before restoring ${new Date(v ? v.time : Date.now()).toISOString().slice(0, 16).replace('T', ' ')}`, auto: false });
    writeFileAtomic(l.full, data);
    if (!l.root) this.changed(l.rel);
    return this.entry(l);
  }
}

/* ---------- own templates ---------- */
class Templates {
  constructor(dir) { this.dir = dir; this.file = path.join(dir, 'templates.json'); fs.mkdirSync(dir, { recursive: true }); }
  list() { return readJson(this.file, []); }
  add(src, { name, category, ext }) {
    const id = newId();
    const e = extOf(`x.${ext}`);
    if (!e) throw fail('Unknown format', 400);
    fs.copyFileSync(src, path.join(this.dir, `${id}.${e}`));
    const f = FORMATS[e] || {};
    const t = { id, name: P.safeName(String(name || '').slice(0, 120) || 'Template'), category: String(category || 'My templates').slice(0, 60), ext: e, kind: f.kind || null, family: f.family || null, size: fs.statSync(path.join(this.dir, `${id}.${e}`)).size, created: Date.now() };
    writeJson(this.file, [...this.list(), t]);
    return t;
  }
  get(id) { return this.list().find((t) => t.id === id) || null; }
  path(id) { const t = this.get(id); return t ? path.join(this.dir, `${t.id}.${t.ext}`) : null; }
  remove(id) { const t = this.get(id); if (!t) throw fail('Template not found', 404); fs.rmSync(path.join(this.dir, `${t.id}.${t.ext}`), { force: true }); writeJson(this.file, this.list().filter((x) => x.id !== id)); }
}

module.exports = { Library, Templates, VERSION_GAP };
