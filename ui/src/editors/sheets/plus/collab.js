/*
 * Co-editing of a workbook through host.collab (Yjs over Hostwerk's relay): a cell-level Y.Map
 * ("cells": "<sheet>\u0001<row>\u0001<col>" → { v, f, s, note } as plain JSON, null = cleared) plus "sheets" (names in
 * order) and the awareness state { user, sel: { sheet, r, c, ranges } } for the coloured selection rectangles.
 * Only edits made during the session travel; everybody opened the same saved file. Structural edits (rows/columns,
 * sorting) send the whole sheet (bounded) so peers end up identical. The leader (lowest client id) saves the file.
 * No Yjs import here: the provider's Y.Doc comes from the shared lib-yjs; plain values in a Y.Map need no constructors.
 */
import { errOf } from '../formula/errors.js';

const SEP = '\u0001';
const MAX_SHEET_PUSH = 50000;
const key = (sheet, r, c) => `${sheet}${SEP}${r}${SEP}${c}`;
const ser = (cell) => (cell ? { v: cell.v && cell.v.code ? { e: cell.v.code } : cell.v ?? null, f: cell.f || null, s: cell.s ? { ...cell.s, _xf: undefined } : null, note: cell.note || null, link: cell.link || null, rich: cell.rich || null } : null);

export async function startCollab(host, book, { onPeers, onStatus, user } = {}) {
  const provider = await host.collab.connect(user ? { user } : {});
  const doc = provider.doc;
  const cells = doc.getMap('cells');
  const meta = doc.getMap('meta');
  let applying = false;
  const deser = (x) => {
    if (!x) return null;
    const cell = {};
    if (x.v && typeof x.v === 'object' && x.v.e) cell.v = errOf(x.v.e); else if (x.v != null) cell.v = x.v;
    if (x.f) cell.f = x.f;
    if (x.s) { const s = { ...x.s }; delete s._xf; cell.s = book.styles.intern(s); }
    if (x.note) cell.note = x.note;
    if (x.link) cell.link = x.link;
    if (x.rich) cell.rich = x.rich;
    return cell;
  };
  // local → remote: every finished transaction of the book
  const origAfter = book.afterChange.bind(book);
  book.afterChange = (tx) => {
    origAfter(tx);
    if (applying) return;
    doc.transact(() => {
      if (tx.book) meta.set('sheets', book.sheets.map((s) => s.name));
      const full = new Set();
      for (const [sheet, s] of tx.snaps) if (s.cells && sheet.cells.size <= MAX_SHEET_PUSH) full.add(sheet);
      for (const sheet of full) {
        const prefix = `${sheet.name}${SEP}`;
        const keep = new Set();
        for (const [k, cell] of sheet.cells) { const kk = key(sheet.name, Math.floor(k / 16384), k % 16384); keep.add(kk); cells.set(kk, ser(cell)); }
        for (const kk of cells.keys()) if (kk.startsWith(prefix) && !keep.has(kk)) cells.set(kk, null);
        meta.set(`struct${SEP}${sheet.name}`, Date.now());
      }
      for (const [sheet, k, , after] of tx.cells) if (!full.has(sheet)) cells.set(key(sheet.name, Math.floor(k / 16384), k % 16384), ser(after));
    }, 'local');
  };
  // remote → local
  const apply = (keys) => {
    const bySheet = new Map();
    for (const kk of keys) { const [sn, r, c] = kk.split(SEP); const sh = book.sheetByName(sn); if (!sh) continue; if (!bySheet.has(sh)) bySheet.set(sh, []); bySheet.get(sh).push([+r, +c, cells.get(kk)]); }
    if (!bySheet.size) return;
    applying = true;
    try {
      book.begin('Co-editing', { remote: true });
      for (const [sh, list] of bySheet) for (const [r, c, x] of list) book.setCell(sh, r, c, deser(x));
    } finally {
      const n = book.undoStack.length;
      book.end({ remote: true });
      if (book.undoStack.length > n) book.undoStack.pop(); // others' edits are not in my undo history
      applying = false;
    }
  };
  const onCells = (ev, tr) => { if (tr.local || tr.origin === 'local') return; apply([...ev.keysChanged]); };
  cells.observe(onCells);
  const onMeta = (ev, tr) => {
    if (tr.local || tr.origin === 'local') return;
    if (ev.keysChanged.has('sheets')) {
      const names = meta.get('sheets') || [];
      applying = true;
      try { for (const n of names) if (!book.sheetByName(n)) book.addSheet(n); } finally { applying = false; }
    }
  };
  meta.observe(onMeta);
  await Promise.race([provider.synced, new Promise((r) => setTimeout(r, 4000))]);
  if (cells.size) apply([...cells.keys()]); // edits the others made before I joined
  // awareness: selections + leader
  const aw = provider.awareness;
  const peers = () => {
    const out = [];
    aw.getStates().forEach((st, id) => { if (id !== doc.clientID && st && st.user) out.push({ id, ...st.user, sel: st.sel || null }); });
    return out;
  };
  const isLeader = () => { let min = doc.clientID; aw.getStates().forEach((st, id) => { if (st && st.user && id < min) min = id; }); return min === doc.clientID; };
  const onAw = () => onPeers && onPeers(peers(), isLeader());
  aw.on('change', onAw);
  provider.on('status', (s) => onStatus && onStatus(s));
  onAw();
  return {
    provider,
    get leader() { return isLeader(); },
    setSelection(sel) { aw.setLocalStateField('sel', sel); },
    peers,
    stop() {
      cells.unobserve(onCells); meta.unobserve(onMeta); aw.off('change', onAw);
      book.afterChange = origAfter;
      try { host.collab.disconnect(); } catch { provider.destroy(); }
    },
  };
}
