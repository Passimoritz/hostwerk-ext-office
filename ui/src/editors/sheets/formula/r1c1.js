/*
 * R1C1 reference style (setting sheets.r1c1 / View → R1C1): display and input only – the model and files keep A1.
 *   toR1C1('SUM(A1:B$2)', 2, 3)   → 'SUM(R[-2]C[-3]:R2C[-2])'   (cell D3 = row 2, col 3 zero-based)
 *   fromR1C1('SUM(R[-2]C[-3]:R2C[-2])', 2, 3) → 'SUM(A1:B$2)'
 */
import { tokenize, sheetPrefix } from './parser.js';
import { MAXR, MAXC, colName } from './refs.js';

const part = (abs, v, cur) => (abs ? String(v + 1) : v === cur ? '' : `[${v - cur}]`);

/** one reference token (parser ref object) → R1C1 text relative to (r, c) */
export function refToR1C1(ref, r, c) {
  const pre = sheetPrefix(ref);
  const R = (row, a) => `R${part(a, row, r)}`;
  const C = (col, a) => `C${part(a, col, c)}`;
  if (ref.kind === 'cols') return `${pre}${C(ref.c1, ref.ac1)}${ref.c1 === ref.c2 && ref.ac1 === ref.ac2 ? '' : `:${C(ref.c2, ref.ac2)}`}`;
  if (ref.kind === 'rows') return `${pre}${R(ref.r1, ref.ar1)}${ref.r1 === ref.r2 && ref.ar1 === ref.ar2 ? '' : `:${R(ref.r2, ref.ar2)}`}`;
  const a = `${R(ref.r1, ref.ar1)}${C(ref.c1, ref.ac1)}`;
  if (ref.kind === 'cell') return pre + a;
  return `${pre}${a}:${R(ref.r2, ref.ar2)}${C(ref.c2, ref.ac2)}`;
}

/** A1 formula (storage, without '=') → R1C1 text for the cell (r, c) */
export function toR1C1(formula, r, c) {
  return tokenize(formula).map((t) => (t.t === 'ref' ? refToR1C1(t.ref, r, c) : t.text)).join('');
}

const NUM = '(\\[-?\\d+\\]|\\d+)?';
const RC = new RegExp(`^R${NUM}C${NUM}`, 'i');
const RONLY = new RegExp(`^R${NUM}(?::R${NUM})?(?![\\w(\\[])`, 'i');
const CONLY = new RegExp(`^C${NUM}(?::C${NUM})?(?![\\w(\\[])`, 'i');
const val = (x, cur) => { if (x == null || x === '') return { v: cur, a: false }; if (x[0] === '[') return { v: cur + Number(x.slice(1, -1)), a: false }; return { v: Number(x) - 1, a: true }; };
const A = (v, a) => `${a ? '$' : ''}${v}`;
const ok = (v, max) => v >= 0 && v < max;

/** R1C1 text (user input, without '=') → A1 text; strings and names are left alone. Throws on references off the sheet. */
export function fromR1C1(text, r, c) {
  const s = String(text);
  let out = '';
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === '"') { let j = i + 1; while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; } out += s.slice(i, j + 1); i = j + 1; continue; }
    const rest = s.slice(i);
    let m = RC.exec(rest);
    if (m) {
      const after = rest.slice(m[0].length);
      if (!/^[A-Za-z_\d(.]/.test(after)) {
        const a = val(m[1], r); const b = val(m[2], c);
        if (!ok(a.v, MAXR) || !ok(b.v, MAXC)) throw new Error('#REF!');
        out += `${A(colName(b.v), b.a)}${A(a.v + 1, a.a)}`;
        i += m[0].length; continue;
      }
    }
    m = RONLY.exec(rest);
    if (m && /^R\d|^R\[|^R:/i.test(m[0]) && !/^\w/.test(rest.slice(m[0].length))) {
      const a = val(m[1], r); const b = m[0].includes(':') ? val(m[2], r) : a;
      if (!ok(a.v, MAXR) || !ok(b.v, MAXR)) throw new Error('#REF!');
      out += `${A(a.v + 1, a.a)}:${A(b.v + 1, b.a)}`; i += m[0].length; continue;
    }
    m = CONLY.exec(rest);
    if (m && /^C\d|^C\[|^C:/i.test(m[0]) && !/^\w/.test(rest.slice(m[0].length))) {
      const a = val(m[1], c); const b = m[0].includes(':') ? val(m[2], c) : a;
      if (!ok(a.v, MAXC) || !ok(b.v, MAXC)) throw new Error('#REF!');
      out += `${A(colName(a.v), a.a)}:${A(colName(b.v), b.a)}`; i += m[0].length; continue;
    }
    // identifiers (function names, names) are copied whole so "ROUND(" is not read as R…
    const id = /^[A-Za-z_\\À-￿][\w.À-￿]*/.exec(rest);
    if (id) { out += id[0]; i += id[0].length; continue; }
    out += ch; i++;
  }
  return out;
}
