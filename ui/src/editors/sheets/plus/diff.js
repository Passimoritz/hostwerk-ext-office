/* Cell-by-cell difference of two workbooks (pure, used by the Compare view and the tests). */
import { MAXC } from '../formula/refs.js';

/** pure: { sheets: [{ name, status: 'added'|'removed'|'same'|'changed', cells: [{ r, c, kind, a, b }] }] } */
export function diffBooks(A, B) {
  const out = [];
  const names = [...new Set([...A.sheets.map((s) => s.name), ...B.sheets.map((s) => s.name)])];
  for (const name of names) {
    const a = A.sheetByName(name); const b = B.sheetByName(name);
    if (!a || !b) { out.push({ name, status: a ? 'removed' : 'added', cells: [] }); continue; }
    const keys = new Set([...a.cells.keys(), ...b.cells.keys()]);
    const cells = [];
    for (const k of keys) {
      const x = a.cells.get(k); const y = b.cells.get(k);
      const tx = x ? `${x.f ? `=${x.f}` : ''}\u0001${A.text(a, Math.floor(k / MAXC), k % MAXC)}` : '';
      const ty = y ? `${y.f ? `=${y.f}` : ''}\u0001${B.text(b, Math.floor(k / MAXC), k % MAXC)}` : '';
      const ea = !x || (x.v == null && !x.f); const eb = !y || (y.v == null && !y.f);
      if (ea && eb) continue;
      if (tx === ty) continue;
      const r = Math.floor(k / MAXC); const c = k % MAXC;
      cells.push({ r, c, kind: ea ? 'added' : eb ? 'removed' : 'changed', a: ea ? null : { text: A.text(a, r, c), f: x.f || null }, b: eb ? null : { text: B.text(b, r, c), f: y.f || null } });
    }
    cells.sort((p, q) => p.r - q.r || p.c - q.c);
    out.push({ name, status: cells.length ? 'changed' : 'same', cells });
  }
  return { sheets: out };
}
