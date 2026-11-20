'use strict';
/*
 * Full text index of the library: <data>/index.json { files: { <path>: { mtime, size, text, thumb, err? } } }.
 * Incremental – a file is read again only when mtime or size changed. Runs in the background in small steps
 * (setImmediate between files) so Hostwerk stays responsive; started on activate and (debounced) after writes.
 */
const fs = require('fs');
const path = require('path');
const { readJson, writeJson } = require('./store');
const X = require('./extract');
const { isJunk } = require('./paths');
const { extOf } = require('./formats');

const MAX_FILE = 80 * 1048576;
const MAX_FILES = 50000;

/** All files below base → [{ rel, full, mtime, size }] (junk and hidden skipped) */
function walk(base, limit = MAX_FILES) {
  const out = [];
  const go = (dir, rel, depth) => {
    if (out.length >= limit || depth > 30) return;
    let list = [];
    try { list = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of list) {
      if (isJunk(e.name)) continue;
      const full = path.join(dir, e.name); const r = rel ? `${rel}/${e.name}` : e.name;
      if (e.isDirectory()) go(full, r, depth + 1);
      else if (e.isFile()) { try { const st = fs.statSync(full); out.push({ rel: r, full, mtime: Math.floor(st.mtimeMs), size: st.size }); } catch { /* gone */ } }
      if (out.length >= limit) return;
    }
  };
  go(base, '', 0);
  return out;
}

/** Score a document for terms (all must match; name hits count more) → number or 0 */
function score(name, text, terms) {
  const n = String(name).toLowerCase(); const t = String(text || '').toLowerCase();
  let s = 0;
  for (const term of terms) {
    const inName = n.includes(term); const i = t.indexOf(term);
    if (!inName && i < 0) return 0;
    s += (inName ? 10 : 0) + (i >= 0 ? 1 + Math.min(5, t.split(term).length - 1) : 0);
  }
  return s;
}

const XLS = /^(xls|xlt)$/;
class Index {
  /** deps: { file, base: () => libraryDir, libs: () => ({ unpdf, mailparser }), log } */
  constructor(deps) {
    this.d = deps;
    this.data = readJson(deps.file, null) || {};
    if (!this.data.files || this.data.base !== deps.base()) this.data = { base: deps.base(), files: {} };
    this.running = null; this.again = false; this.timer = null; this.stopped = false; this.dirty = false;
  }
  get(rel) { return this.data.files[rel] || null; }
  save() { if (!this.dirty) return; this.dirty = false; try { writeJson(this.d.file, this.data); } catch (e) { this.d.log && this.d.log('index save failed', e.message); } }
  /** Re-read changed files (one at a time) → Promise */
  reindex() {
    if (this.running) { this.again = true; return this.running; }
    this.running = (async () => {
      do {
        this.again = false;
        const base = this.d.base();
        if (this.data.base !== base) this.data = { base, files: {} };
        const list = walk(base);
        const seen = new Set();
        let n = 0;
        for (const f of list) {
          if (this.stopped) return;
          seen.add(f.rel);
          const ext = extOf(f.rel);
          const cur = this.data.files[f.rel];
          // .xls text came with 1.2 (BIFF8 reader) – entries of older indexes are read again once
          if (cur && cur.mtime === f.mtime && cur.size === f.size && !(XLS.test(ext) && cur.v !== 2)) continue;
          const rec = { mtime: f.mtime, size: f.size, text: '', thumb: false, ...(XLS.test(ext) ? { v: 2 } : {}) };
          if (X.canText(ext) && f.size <= MAX_FILE) {
            try {
              const buf = await fs.promises.readFile(f.full);
              rec.text = await X.text(buf, ext, this.d.libs());
              rec.thumb = X.hasThumb(buf, ext);
            } catch (e) { rec.err = String(e.message || e).slice(0, 200); }
          }
          this.data.files[f.rel] = rec; this.dirty = true;
          if (++n % 20 === 0) this.save();
          await new Promise((r) => setImmediate(r));
        }
        for (const k of Object.keys(this.data.files)) if (!seen.has(k)) { delete this.data.files[k]; this.dirty = true; }
        this.save();
      } while (this.again && !this.stopped);
    })().catch((e) => { this.d.log && this.d.log('reindex failed', e.message); }).finally(() => { this.running = null; });
    return this.running;
  }
  /** Debounced reindex after changes */
  soon(ms = 3000) { clearTimeout(this.timer); this.timer = setTimeout(() => this.reindex(), ms); if (this.timer.unref) this.timer.unref(); }
  /** → [{ rel, score, snippet }] best first; dir limits to a folder, filter(rel) for kind etc. */
  search(q, { dir = '', filter = null, limit = 200 } = {}) {
    const terms = String(q || '').toLowerCase().split(/\s+/).filter(Boolean).slice(0, 8);
    if (!terms.length) return [];
    const out = [];
    for (const [rel, rec] of Object.entries(this.data.files)) {
      if (dir && !(rel === dir || rel.startsWith(`${dir}/`))) continue;
      if (filter && !filter(rel)) continue;
      const s = score(path.posix.basename(rel), rec.text, terms);
      if (s > 0) out.push({ rel, score: s, snippet: X.snippet(rec.text, terms.join(' ')) });
    }
    return out.sort((a, b) => b.score - a.score).slice(0, limit);
  }
  stop() { this.stopped = true; clearTimeout(this.timer); this.save(); }
}

module.exports = { Index, walk, score };
