'use strict';
/*
 * Marks of the office files in one folder of Hostwerk's Files (files.marks, GET /office/marks): favourite, shared, open
 * in an editor (soft lock), tags. Keys are built exactly like Library.loc builds them – the lock an editor takes for
 * { root, path } is the lock the folder's marks find (library root and roots containing the library: library paths).
 */
const { FORMATS, extOf } = require('./formats');

/** key of a file <name> in the folder location d ({ root, rel } from Library.loc) – same as Library.loc(file).key */
function markKey(d, name) {
  const rel = d.rel ? `${d.rel}/${name}` : name;
  return d.root ? `@${d.root}/${rel}` : rel;
}

/**
 * d: folder location · list: fs.Dirent[] · meta: Meta · shared: Set of shared keys · client: the asking window –
 * its own locks are not shown (when Files shows in a window, that window's editor is closed: a lock of it is only a
 * request still on its way or one that got lost) · types: also list every office file (Hostwerk 3.2.x type marks)
 * → { names: { name: { favorite?, shared?, lock?: { by, since, until }, tags? } }, types? }
 */
function folderMarks({ d, list, meta, shared, client = null, types = false, now = Date.now() }) {
  const names = {};
  const all = types ? [] : null;
  for (const e of list.slice(0, 5000)) {
    if (!e.isFile() || !FORMATS[extOf(e.name)]) continue;
    const key = markKey(d, e.name);
    const f = meta.fields(key);
    const l = meta.lockOf(key, now);
    const lock = l && !(client && l.client && l.client === client) ? { by: l.by, since: l.since, until: l.until } : null;
    const m = { ...(f.favorite ? { favorite: true } : {}), ...(shared.has(key) ? { shared: true } : {}), ...(lock ? { lock } : {}), ...((f.tags || []).length ? { tags: f.tags } : {}) };
    if (Object.keys(m).length) names[e.name] = m;
    if (all) all.push(e.name);
  }
  return { names, ...(all ? { types: all } : {}) };
}

module.exports = { markKey, folderMarks };
