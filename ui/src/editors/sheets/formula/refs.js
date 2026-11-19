/* A1 address helpers (all indexes zero-based internally). */
export const MAXR = 1048576;
export const MAXC = 16384;

const COLS = [];
export function colName(c) {
  if (c < 702 && COLS[c]) return COLS[c];
  let s = '';
  let n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  if (c < 702) COLS[c] = s;
  return s;
}
export function colIndex(letters) {
  let n = 0;
  const s = letters.toUpperCase();
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}
export const key = (r, c) => r * MAXC + c;
export const rowOf = (k) => Math.floor(k / MAXC);
export const colOf = (k) => k % MAXC;
export const addr = (r, c) => colName(c) + (r + 1);

/** 'B3' → { r: 2, c: 1 } (no sheet, $ ignored) or null */
export function parseA1(s) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(s).trim());
  if (!m) return null;
  const c = colIndex(m[1]); const r = Number(m[2]) - 1;
  if (c >= MAXC || r < 0 || r >= MAXR) return null;
  return { r, c };
}

export const needsQuote = (name) => !/^[A-Za-z_À-￿][\w.À-￿]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^R\d*C\d*$/i.test(name) || /^(TRUE|FALSE)$/i.test(name);
export const quoteSheet = (name) => (needsQuote(name) ? `'${String(name).replace(/'/g, "''")}'` : name);

/** Range text: 'A1', 'A1:B5', 'A:C', '2:4' with absolute flags; sheet prefix when given. */
export function rangeText({ sheet, r1, c1, r2, c2, ar1, ac1, ar2, ac2, kind }, withSheet = true) {
  const pre = withSheet && sheet ? `${quoteSheet(sheet)}!` : '';
  const k = kind || (r1 === 0 && r2 === MAXR - 1 && !(c1 === 0 && c2 === MAXC - 1) ? 'cols' : c1 === 0 && c2 === MAXC - 1 && !(r1 === 0 && r2 === MAXR - 1) ? 'rows' : r1 === r2 && c1 === c2 ? 'cell' : 'range');
  const C = (c, a) => (a ? '$' : '') + colName(c);
  const R = (r, a) => (a ? '$' : '') + (r + 1);
  if (k === 'cols') return `${pre}${C(c1, ac1)}:${C(c2, ac2)}`;
  if (k === 'rows') return `${pre}${R(r1, ar1)}:${R(r2, ar2)}`;
  if (k === 'cell') return `${pre}${C(c1, ac1)}${R(r1, ar1)}`;
  return `${pre}${C(c1, ac1)}${R(r1, ar1)}:${C(c2, ac2)}${R(r2, ar2)}`;
}

/** Plain range text without $ ('A1:B5'), single cell collapses. */
export const rangeA1 = (g) => (g.r1 === g.r2 && g.c1 === g.c2 ? addr(g.r1, g.c1) : `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`);

/** Parse "Sheet 1!A1:B3", "'It''s'!A:A", "A1", "3:5" → { sheet|null, r1, c1, r2, c2 } or null. */
export function parseRange(text) {
  let s = String(text || '').trim();
  let sheet = null;
  const q = /^'((?:[^']|'')+)'!/.exec(s) || /^([^!'"]+)!/.exec(s);
  if (q) { sheet = q[1].replace(/''/g, "'"); s = s.slice(q[0].length); }
  s = s.replace(/\$/g, '');
  let m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(s);
  if (m) {
    const c1 = colIndex(m[1]); const r1 = +m[2] - 1;
    const c2 = m[3] ? colIndex(m[3]) : c1; const r2 = m[4] ? +m[4] - 1 : r1;
    if (c1 >= MAXC || c2 >= MAXC || r1 < 0 || r2 < 0 || r1 >= MAXR || r2 >= MAXR) return null;
    return { sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
  }
  m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(s);
  if (m) { const a = colIndex(m[1]); const b = colIndex(m[2]); if (a >= MAXC || b >= MAXC) return null; return { sheet, r1: 0, r2: MAXR - 1, c1: Math.min(a, b), c2: Math.max(a, b) }; }
  m = /^(\d+):(\d+)$/.exec(s);
  if (m) { const a = +m[1] - 1; const b = +m[2] - 1; if (a < 0 || b < 0) return null; return { sheet, c1: 0, c2: MAXC - 1, r1: Math.min(a, b), r2: Math.max(a, b) }; }
  return null;
}

export const inRange = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;
export const rangesIntersect = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
export const normRange = (g) => ({ ...g, r1: Math.min(g.r1, g.r2), r2: Math.max(g.r1, g.r2), c1: Math.min(g.c1, g.c2), c2: Math.max(g.c1, g.c2) });
