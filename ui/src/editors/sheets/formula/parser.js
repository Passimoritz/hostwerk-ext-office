/*
 * Formula tokenizer + parser (→ AST) + token-level rewriting (copy shift, insert/delete rows/cols, move, sheet
 * rename, locale translation). Formulas are stored WITHOUT the leading '=' in English notation (',' separators,
 * '.' decimals). Input may use ';' separators (always) and ',' decimals when the locale says so.
 */
import { MAXR, MAXC, colIndex, colName, quoteSheet } from './refs.js';
import { errOf, ERROR_CODES } from './errors.js';
import { ERROR_CODES_DE, internalErr, internalFn, localFn, localErr, fnLangOf } from './i18n.js';

// English and German error codes (#N/A, #NV …) – German ones become English tokens
const ERR_RE = new RegExp(`^(${[...ERROR_CODES, ...ERROR_CODES_DE].sort((a, b) => b.length - a.length).map((c) => c.replace(/[.*+?^${}()|[\]\\/!]/g, '\\$&')).join('|')})`, 'i');
const SHEET_Q = /^'((?:[^']|'')+)'!/;
const SHEET_U = /^([A-Za-z_À-￿][\w.À-￿]*)!/;
const SHEET_3D = /^([A-Za-z_À-￿][\w.À-￿]*):([A-Za-z_À-￿][\w.À-￿]*)!/;
const SPECIALS = { '#ALL': 'all', '#DATA': 'data', '#HEADERS': 'headers', '#TOTALS': 'totals', '#THIS ROW': 'thisRow' };
const SPECIAL_TEXT = { all: '#All', data: '#Data', headers: '#Headers', totals: '#Totals', thisRow: '#This Row' };
/** end index (exclusive) of a bracket group starting at s[i] === '[' (column names escape [ ] # ' @ with ') or -1 */
function bracketEnd(s, i) {
  let depth = 0;
  for (let j = i; j < s.length; j++) {
    const ch = s[j];
    if (ch === "'") { j++; continue; }
    if (ch === '[') depth++;
    else if (ch === ']') { depth--; if (!depth) return j + 1; }
  }
  return -1;
}
const unescCol = (x) => x.replace(/'(.)/g, '$1');
export const escCol = (x) => String(x).replace(/(['#[\]@])/g, "'$1");
/** structured reference body ("[Col]", "[@Col]", "[[#Headers],[A]:[B]]", "[#All]", "[]") → { specials, c1, c2, thisRow } */
export function parseStructured(body) {
  const inner = body.slice(1, -1).trim();
  const out = { specials: [], c1: null, c2: null };
  const col = (t) => { const x = t.trim(); return x.startsWith('[') ? unescCol(x.slice(1, -1)) : unescCol(x); };
  const item = (t) => {
    t = t.trim(); if (!t) return;
    if (t.startsWith('@')) { out.specials.push('thisRow'); const rest = t.slice(1).trim(); if (rest) item(rest.startsWith('[') ? rest : `[${rest}]`); return; }
    const m = /^\[(.*)\]$/s.exec(t);
    const plain = m && !/^\[[^\]]*\]\s*:/.test(t) ? m[1] : null;
    if (plain != null && SPECIALS[plain.trim().toUpperCase()]) { out.specials.push(SPECIALS[plain.trim().toUpperCase()]); return; }
    if (SPECIALS[t.toUpperCase()]) { out.specials.push(SPECIALS[t.toUpperCase()]); return; }
    // [A]:[B] or [A]
    let depth = 0; let split = -1;
    for (let j = 0; j < t.length; j++) { const ch = t[j]; if (ch === "'") { j++; continue; } if (ch === '[') depth++; else if (ch === ']') depth--; else if (ch === ':' && !depth) { split = j; break; } }
    if (split >= 0) { out.c1 = col(t.slice(0, split)); out.c2 = col(t.slice(split + 1)); } else { out.c1 = col(t); out.c2 = out.c1; }
  };
  if (!inner) return out;
  if (inner.startsWith('[') || inner.startsWith('@')) {
    // list of bracket items separated by commas (top level)
    let depth = 0; let start = 0;
    for (let j = 0; j <= inner.length; j++) {
      const ch = inner[j];
      if (ch === "'") { j++; continue; }
      if (ch === '[') depth++; else if (ch === ']') depth--;
      else if ((ch === ',' || ch === ';' || j === inner.length) && !depth) { item(inner.slice(start, j)); start = j + 1; }
      if (j === inner.length) break;
    }
  } else if (SPECIALS[inner.toUpperCase()]) out.specials.push(SPECIALS[inner.toUpperCase()]);
  else { out.c1 = unescCol(inner); out.c2 = out.c1; }
  return out;
}
/** structured reference → text; thisRowAt = OOXML form ([#This Row]) instead of @ */
export function structuredText(table, sp, { ooxml = false } = {}) {
  const cols = sp.c1 == null ? '' : sp.c1 === sp.c2 ? `[${escCol(sp.c1)}]` : `[${escCol(sp.c1)}]:[${escCol(sp.c2)}]`;
  const sps = sp.specials.filter((x) => x !== 'thisRow' || ooxml);
  const at = sp.specials.includes('thisRow') && !ooxml;
  let body;
  if (at) body = cols ? (sp.c1 === sp.c2 && !/[\s'#[\]@]/.test(sp.c1) ? `[@${escCol(sp.c1)}]` : `[@${cols}]`) : '[@]';
  else if (!sps.length) body = cols ? (sp.c1 === sp.c2 ? cols : `[${cols}]`) : '[]';
  else if (!cols && sps.length === 1) body = `[${SPECIAL_TEXT[sps[0]]}]`;
  else body = `[${[...sps.map((x) => `[${SPECIAL_TEXT[x]}]`), ...(cols ? [cols] : [])].join(',')}]`;
  return `${table || ''}${body}`;
}
const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;
const COLR = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/;
const ROWR = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\w.(])/;
const IDENT = /^[A-Za-z_\\À-￿][\w.À-￿]*/;
const NUM_EN = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/;
const NUM_DE = /^(\d+,?\d*)([eE][+-]?\d+)?/;

function cellAt(s) {
  const m = CELL.exec(s);
  if (!m) return null;
  const c = colIndex(m[2]); const r = Number(m[4]) - 1;
  if (c >= MAXC || r < 0 || r >= MAXR) return null;
  const next = s[m[0].length];
  if (next && /[\w(.]/.test(next)) return null; // LOG10( or a longer name
  return { len: m[0].length, c, r, ac: !!m[1], ar: !!m[3] };
}

/**
 * Tokenize. opts.decimal ',' → numbers use ',' decimals and ',' is never a separator outside numbers?? (no: ',' after
 * a non-number still separates). ';' always separates arguments (or rows inside {}).
 */
export function tokenize(src, opts = {}) {
  const de = opts.decimal === ',';
  const toks = [];
  let i = 0;
  let brace = 0;
  const s = String(src);
  while (i < s.length) {
    const ch = s[i];
    if (/\s/.test(ch)) { let j = i; while (j < s.length && /\s/.test(s[j])) j++; toks.push({ t: 'ws', text: s.slice(i, j), pos: i }); i = j; continue; }
    if (ch === '"') {
      let j = i + 1; let v = '';
      while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { v += '"'; j += 2; continue; } break; } v += s[j++]; }
      toks.push({ t: 'str', v, text: s.slice(i, j + 1), pos: i }); i = j + 1; continue;
    }
    if (ch === '#') {
      const m = ERR_RE.exec(s.slice(i, i + 20));
      if (m) { toks.push({ t: 'err', v: internalErr(m[1]), text: m[0], pos: i }); i += m[0].length; continue; }
      toks.push({ t: 'hash', text: '#', pos: i }); i++; continue;
    }
    // punctuation never starts a reference or a number – handled below without the reference regexes
    if (PUNCT.has(ch)) {
      const two = s.slice(i, i + 2);
      if (two === '<>' || two === '<=' || two === '>=') { toks.push({ t: 'op', v: two, text: two, pos: i }); i += 2; continue; }
      if ('+-*/^&=<>%'.includes(ch)) { toks.push({ t: 'op', v: ch, text: ch, pos: i }); i++; continue; }
      if (ch === '(') { toks.push({ t: 'lp', text: ch, pos: i }); i++; continue; }
      if (ch === ')') { toks.push({ t: 'rp', text: ch, pos: i }); i++; continue; }
      if (ch === ',') { toks.push({ t: brace ? 'col' : 'sep', text: ch, pos: i }); i++; continue; }
      if (ch === ';') { toks.push({ t: brace ? 'row' : 'sep', text: ch, pos: i }); i++; continue; }
    }
    const rest = s.slice(i, i + 300);
    // references with optional sheet prefix
    let sheet = null; let sheet2 = null; let pl = 0;
    const bang = ch === "'" || rest.indexOf('!') > 0; // a sheet prefix needs a '!'
    let m = bang ? SHEET_Q.exec(rest) : null;
    if (m) { sheet = m[1].replace(/''/g, "'"); pl = m[0].length; const p3 = sheet.indexOf(':'); if (p3 > 0) { sheet2 = sheet.slice(p3 + 1); sheet = sheet.slice(0, p3); } } else if (bang && (m = SHEET_3D.exec(rest))) { sheet = m[1]; sheet2 = m[2]; pl = m[0].length; } else if (bang && (m = SHEET_U.exec(rest))) { sheet = m[1]; pl = m[0].length; }
    const body = pl ? rest.slice(pl) : rest;
    let ref = null;
    const colon = body.indexOf(':') > 0;
    const c0 = body.charCodeAt(0);
    const digit = c0 >= 48 && c0 <= 57;
    let mm = colon && !digit ? COLR.exec(body) : null;
    if (mm && colIndex(mm[2]) < MAXC && colIndex(mm[4]) < MAXC) {
      const a = colIndex(mm[2]); const b = colIndex(mm[4]);
      ref = { len: mm[0].length, kind: 'cols', r1: 0, r2: MAXR - 1, c1: a, c2: b, ar1: false, ar2: false, ac1: !!mm[1], ac2: !!mm[3] };
    } else if (colon && (digit || c0 === 36) && (mm = ROWR.exec(body)) && (sheet || !toks.length || prevAllowsRef(toks))) {
      const a = Number(mm[2]) - 1; const b = Number(mm[4]) - 1;
      if (a >= 0 && b >= 0 && a < MAXR && b < MAXR) ref = { len: mm[0].length, kind: 'rows', c1: 0, c2: MAXC - 1, r1: a, r2: b, ac1: false, ac2: false, ar1: !!mm[1], ar2: !!mm[3] };
    } else if (!digit) {
      const c1 = cellAt(body);
      if (c1) {
        ref = { len: c1.len, kind: 'cell', r1: c1.r, c1: c1.c, r2: c1.r, c2: c1.c, ar1: c1.ar, ac1: c1.ac, ar2: c1.ar, ac2: c1.ac };
        if (body[c1.len] === ':') {
          const c2 = cellAt(body.slice(c1.len + 1));
          if (c2) Object.assign(ref, { len: c1.len + 1 + c2.len, kind: 'range', r2: c2.r, c2: c2.c, ar2: c2.ar, ac2: c2.ac });
        }
      }
    }
    if (ref) {
      const text = rest.slice(0, pl + ref.len);
      delete ref.len;
      toks.push({ t: 'ref', ref: sheet2 ? { sheet, sheet2, ...ref } : { sheet, ...ref }, text, pos: i }); i += text.length; continue;
    }
    if (sheet) { // Sheet!Name (sheet-scoped name)
      const n = IDENT.exec(body);
      if (n) { const text = rest.slice(0, pl + n[0].length); toks.push({ t: 'name', v: n[0], sheet, text, pos: i }); i += text.length; continue; }
    }
    m = (de ? NUM_DE : NUM_EN).exec(rest);
    if (m && !(de && brace > 0 && false)) {
      const txt = m[0];
      const v = Number(txt.replace(',', '.'));
      toks.push({ t: 'num', v, text: txt, pos: i }); i += txt.length; continue;
    }
    m = IDENT.exec(rest);
    if (m) {
      let name = m[0];
      if (s[i + name.length] === '[') { // Table1[Col]
        const e = bracketEnd(s, i + name.length);
        if (e > 0) { const text = s.slice(i, e); toks.push({ t: 'sref', table: name, spec: parseStructured(s.slice(i + name.length, e)), text, pos: i }); i = e; continue; }
      }
      const xp = /^_xlpm\./i.exec(name);
      if (xp) name = name.slice(6);
      const after = s.slice(i + name.length).match(/^\s*\(/);
      if (after) {
        let fn = name.toUpperCase().replace(/^_XLFN\._XLWS\.|^_XLFN\.|^_XLWS\./, '');
        if (fn.startsWith('_XLL.')) fn = fn.slice(5);
        toks.push({ t: 'func', v: fn, text: name, pos: i, ...(xp ? { pm: true } : {}) }); i += m[0].length; continue;
      }
      const up = name.toUpperCase();
      if (!xp && (up === 'TRUE' || up === 'FALSE' || up === 'WAHR' || up === 'FALSCH')) { toks.push({ t: 'bool', v: up === 'TRUE' || up === 'WAHR', text: name, pos: i }); i += name.length; continue; }
      toks.push({ t: 'name', v: name, text: name, pos: i }); i += m[0].length; continue;
    }
    const two = s.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') { toks.push({ t: 'op', v: two, text: two, pos: i }); i += 2; continue; }
    if ('+-*/^&=<>%'.includes(ch)) { toks.push({ t: 'op', v: ch, text: ch, pos: i }); i++; continue; }
    if (ch === '(') { toks.push({ t: 'lp', text: ch, pos: i }); i++; continue; }
    if (ch === ')') { toks.push({ t: 'rp', text: ch, pos: i }); i++; continue; }
    if (ch === '{') { brace++; toks.push({ t: 'lb', text: ch, pos: i }); i++; continue; }
    if (ch === '}') { brace = Math.max(0, brace - 1); toks.push({ t: 'rb', text: ch, pos: i }); i++; continue; }
    if (ch === ',') { toks.push({ t: brace ? 'col' : 'sep', text: ch, pos: i }); i++; continue; }
    if (ch === ';') { toks.push({ t: brace ? 'row' : 'sep', text: ch, pos: i }); i++; continue; }
    if (ch === ':') { toks.push({ t: 'colon', text: ch, pos: i }); i++; continue; }
    if (ch === '@') { toks.push({ t: 'at', text: ch, pos: i }); i++; continue; }
    if (ch === '[') { const e = bracketEnd(s, i); if (e > 0) { const text = s.slice(i, e); toks.push({ t: 'sref', table: null, spec: parseStructured(text), text, pos: i }); i = e; continue; } }
    if (ch === '!' || ch === '[' || ch === ']') { toks.push({ t: 'bad', text: ch, pos: i }); i++; continue; }
    toks.push({ t: 'bad', text: ch, pos: i }); i++;
  }
  return toks;
}
function prevAllowsRef(toks) {
  for (let k = toks.length - 1; k >= 0; k--) {
    const t = toks[k];
    if (t.t === 'ws') continue;
    return t.t === 'op' || t.t === 'lp' || t.t === 'sep' || t.t === 'colon';
  }
  return true;
}

/* ---------------- Parser ---------------- */
const BIN = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export class ParseError extends Error {}

const PUNCT = new Set(['+', '-', '*', '/', '^', '&', '=', '<', '>', '%', '(', ')', ',', ';']);
const ISECT_L = new Set(['ref', 'name', 'rp', 'sref']);
const ISECT_R = new Set(['ref', 'name', 'sref']);
export function parse(src, opts) {
  const all = tokenize(src, opts);
  // a space between two references is the intersection operator
  const toks = [];
  for (let k = 0; k < all.length; k++) {
    const t = all[k];
    if (t.t !== 'ws') { toks.push(t); continue; }
    const prev = toks[toks.length - 1]; const nx = all[k + 1];
    if (prev && nx && ISECT_L.has(prev.t) && ISECT_R.has(nx.t)) toks.push({ t: 'op', v: ' ', text: t.text, pos: t.pos });
  }
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const fail = (msg) => { throw new ParseError(msg || 'Formula error'); };

  function expr(min = 0) {
    let left = prefix();
    for (;;) {
      const t = peek();
      if (!t || t.t !== 'op' || !(t.v in BIN)) break;
      const prec = BIN[t.v];
      if (prec < min) break;
      next();
      const right = expr(t.v === '^' ? prec + 0.5 : prec + 1); // ^ is left-assoc in Excel: 2^3^2 = 64
      left = { t: 'bin', op: t.v, a: left, b: right };
    }
    return left;
  }
  function prefix() {
    const t = peek();
    if (t && t.t === 'op' && (t.v === '-' || t.v === '+')) { next(); const a = prefix(); return t.v === '-' ? { t: 'neg', a } : a; }
    if (t && t.t === 'at') { next(); return prefix(); }
    return postfix(isect());
  }
  function isect() {
    let n = range();
    while (peek() && peek().t === 'op' && peek().v === ' ') { next(); n = { t: 'isect', a: n, b: range() }; }
    return n;
  }
  function postfix(n) {
    while (peek() && peek().t === 'op' && peek().v === '%') { next(); n = { t: 'pct', a: n }; }
    return n;
  }
  function args() { // after '('
    const out = [];
    if (peek() && peek().t === 'rp') { next(); return out; }
    for (;;) {
      const q = peek();
      if (q && (q.t === 'sep' || q.t === 'rp')) out.push({ t: 'missing' });
      else out.push(expr());
      const z = next();
      if (!z) fail('Missing )');
      if (z.t === 'rp') break;
      if (z.t !== 'sep') fail();
    }
    return out;
  }
  function range() {
    let n = primary();
    // LAMBDA(x, x+1)(5) – calling the result of a function / parenthesis
    while ((n.t === 'fn' || n.t === 'paren' || n.t === 'call') && peek() && peek().t === 'lp') { next(); n = { t: 'call', f: n, args: args() }; }
    while (peek() && (peek().t === 'colon' || peek().t === 'hash')) {
      if (next().t === 'hash') { n = { t: 'spill', a: n }; continue; }
      n = { t: 'rangeop', a: n, b: primary() };
    }
    return n;
  }
  function primary() {
    const t = next();
    if (!t) fail('Unexpected end');
    switch (t.t) {
      case 'num': return { t: 'num', v: t.v };
      case 'str': return { t: 'str', v: t.v };
      case 'bool': return { t: 'bool', v: t.v };
      case 'err': return { t: 'err', v: errOf(t.v) };
      case 'ref': return { t: 'ref', ...t.ref };
      case 'name': return { t: 'name', v: t.v, sheet: t.sheet || null };
      case 'sref': return { t: 'sref', table: t.table, ...t.spec };
      case 'func': {
        if (!peek() || peek().t !== 'lp') fail();
        next();
        return { t: 'fn', name: t.v, args: args() };
      }
      case 'lp': {
        const items = [expr()];
        while (peek() && peek().t === 'sep') { next(); items.push(expr()); }
        if (!peek() || next().t !== 'rp') fail('Missing )');
        return items.length > 1 ? { t: 'union', items } : { t: 'paren', a: items[0] };
      }
      case 'lb': {
        const rows = [[]];
        for (;;) {
          const z = next();
          if (!z) fail('Missing }');
          if (z.t === 'rb') break;
          if (z.t === 'col') continue;
          if (z.t === 'row') { rows.push([]); continue; }
          let neg = false;
          let v = z;
          if (z.t === 'op' && z.v === '-') { neg = true; v = next(); }
          if (!v) fail();
          const val = v.t === 'num' ? (neg ? -v.v : v.v) : v.t === 'str' ? v.v : v.t === 'bool' ? v.v : v.t === 'err' ? errOf(v.v) : fail('Bad array');
          rows[rows.length - 1].push(val);
        }
        return { t: 'array', rows };
      }
      default: fail(`Unexpected ${t.text}`);
    }
    return null;
  }
  if (!toks.length) fail('Empty formula');
  const ast = expr();
  if (p < toks.length) fail(`Unexpected ${toks[p].text}`);
  return ast;
}

/* ---------------- Rewriting ---------------- */
export function sheetPrefix(ref) {
  if (!ref.sheet) return '';
  if (ref.sheet2) { const both = `${ref.sheet}:${ref.sheet2}`; return /^[A-Za-z_À-￿][\w.À-￿]*$/.test(ref.sheet) && /^[A-Za-z_À-￿][\w.À-￿]*$/.test(ref.sheet2) && quoteSheet(ref.sheet) === ref.sheet && quoteSheet(ref.sheet2) === ref.sheet2 ? `${both}!` : `'${both.replace(/'/g, "''")}'!`; }
  return `${quoteSheet(ref.sheet)}!`;
}
function refText(ref) {
  const pre = sheetPrefix(ref);
  const C = (c, a) => (a ? '$' : '') + colName(c);
  const R = (r, a) => (a ? '$' : '') + (r + 1);
  if (ref.kind === 'cols') return `${pre}${C(ref.c1, ref.ac1)}:${C(ref.c2, ref.ac2)}`;
  if (ref.kind === 'rows') return `${pre}${R(ref.r1, ref.ar1)}:${R(ref.r2, ref.ar2)}`;
  if (ref.kind === 'cell') return `${pre}${C(ref.c1, ref.ac1)}${R(ref.r1, ref.ar1)}`;
  return `${pre}${C(ref.c1, ref.ac1)}${R(ref.r1, ref.ar1)}:${C(ref.c2, ref.ac2)}${R(ref.r2, ref.ar2)}`;
}
const refErr = (ref) => (ref.sheet ? `${quoteSheet(ref.sheet)}!#REF!` : '#REF!');

/** Apply fn(ref) → ref | null (#REF!) | undefined (unchanged) to every reference; returns the new formula text. */
export function mapRefs(formula, fn, opts) {
  const toks = tokenize(formula, opts);
  let out = '';
  for (const t of toks) {
    if (t.t === 'ref') {
      const r = fn({ ...t.ref });
      out += r === undefined ? t.text : r === null ? refErr(t.ref) : refText(r);
    } else out += t.text;
  }
  return out;
}

/** Copy/fill: shift relative parts by dr/dc. Out of the sheet → #REF!. */
export function shiftFormula(formula, dr, dc) {
  if (!dr && !dc) return formula;
  return mapRefs(formula, (ref) => {
    const n = { ...ref };
    if (ref.kind !== 'cols') { if (!ref.ar1) n.r1 += dr; if (!ref.ar2) n.r2 += dr; }
    if (ref.kind !== 'rows') { if (!ref.ac1) n.c1 += dc; if (!ref.ac2) n.c2 += dc; }
    if (n.r1 < 0 || n.r2 < 0 || n.c1 < 0 || n.c2 < 0 || n.r1 >= MAXR || n.r2 >= MAXR || n.c1 >= MAXC || n.c2 >= MAXC) return null;
    return n;
  });
}

/**
 * Insert (count > 0) or delete (count < 0) rows/columns at index `at` on sheet `target`.
 * `here` = name of the sheet the formula lives on (refs without sheet prefix point there).
 */
export function adjustFormula(formula, { target, here, axis, at, count }) {
  const same = (s) => String(s || here).toLowerCase() === String(target).toLowerCase();
  return mapRefs(formula, (ref) => {
    if (!same(ref.sheet)) return undefined;
    const lo = axis === 'row' ? 'r1' : 'c1';
    const hi = axis === 'row' ? 'r2' : 'c2';
    if ((axis === 'row' && ref.kind === 'cols') || (axis === 'col' && ref.kind === 'rows')) return undefined;
    const n = { ...ref };
    const max = axis === 'row' ? MAXR : MAXC;
    if (count > 0) {
      if (n[lo] >= at) n[lo] += count;
      if (n[hi] >= at) n[hi] += count;
      if (n[hi] >= max) { if (n[lo] >= max) return null; n[hi] = max - 1; }
      return n[lo] === ref[lo] && n[hi] === ref[hi] ? undefined : n;
    }
    const del = -count; const end = at + del - 1;
    const single = ref.kind === 'cell';
    if (single) {
      if (n[lo] >= at && n[lo] <= end) return null;
      if (n[lo] > end) { n[lo] -= del; n[hi] -= del; return n; }
      return undefined;
    }
    let a = n[lo]; let b = n[hi];
    if (a >= at && b <= end) return null; // whole range deleted
    if (a > end) a -= del; else if (a >= at) a = at;
    if (b > end) b -= del; else if (b >= at) b = at - 1;
    if (b < a) return null;
    n[lo] = a; n[hi] = b;
    return a === ref[lo] && b === ref[hi] ? undefined : n;
  });
}

/** Cut/paste or drag-move of `src` (on sheet `srcSheet`) by dr/dc, possibly to another sheet `dstSheet`. */
export function moveFormulaRefs(formula, { here, srcSheet, src, dr, dc, dstSheet }) {
  const same = (s) => String(s || here).toLowerCase() === String(srcSheet).toLowerCase();
  return mapRefs(formula, (ref) => {
    if (!same(ref.sheet)) return undefined;
    if (ref.r1 >= src.r1 && ref.r2 <= src.r2 && ref.c1 >= src.c1 && ref.c2 <= src.c2) {
      const n = { ...ref, r1: ref.r1 + dr, r2: ref.r2 + dr, c1: ref.c1 + dc, c2: ref.c2 + dc };
      if (dstSheet && dstSheet.toLowerCase() !== String(here).toLowerCase()) n.sheet = dstSheet;
      else if (dstSheet && ref.sheet) n.sheet = dstSheet;
      return n;
    }
    return undefined;
  });
}

export function renameSheetInFormula(formula, oldName, newName) {
  const lo = oldName.toLowerCase();
  const toks = tokenize(formula);
  let out = '';
  for (const t of toks) {
    if (t.t === 'ref' && t.ref.sheet && (t.ref.sheet.toLowerCase() === lo || (t.ref.sheet2 && t.ref.sheet2.toLowerCase() === lo))) out += refText({ ...t.ref, sheet: t.ref.sheet.toLowerCase() === lo ? newName : t.ref.sheet, ...(t.ref.sheet2 ? { sheet2: t.ref.sheet2.toLowerCase() === lo ? newName : t.ref.sheet2 } : {}) });
    else if (t.t === 'name' && t.sheet && t.sheet.toLowerCase() === lo) out += `${quoteSheet(newName)}!${t.v}`;
    else out += t.text;
  }
  return out;
}

/** References in a formula with their text positions (for colouring while editing). */
export function refsIn(formula, opts) {
  return tokenize(formula, opts).filter((t) => t.t === 'ref').map((t) => ({ ...t.ref, pos: t.pos, len: t.text.length }));
}

/**
 * Locale translation. toInternal: user text (without '=') in the given locale → English storage.
 * toDisplay: storage → what a German user expects to see/edit (';' separators, ',' decimals, and with loc.fnLang 'de'
 * German function names, WAHR/FALSCH and German error codes – input accepts German and English names).
 */
export function toInternal(text, loc) {
  const de = loc && loc.decimal === ',';
  const fl = fnLangOf(loc);
  const toks = tokenize(text, { decimal: de ? ',' : '.' });
  return toks.map((t) => (t.t === 'sep' ? ',' : t.t === 'row' ? ';' : t.t === 'col' ? ',' : t.t === 'num' ? t.text.replace(',', '.')
    : t.t === 'func' ? internalFn(t.v, fl) : t.t === 'bool' ? (t.v ? 'TRUE' : 'FALSE') : t.t === 'err' ? t.v : t.t === 'sref' ? structuredText(t.table, t.spec) : t.text)).join('');
}
export function toDisplay(formula, loc) {
  const de = !!loc && loc.decimal === ',';
  const fl = fnLangOf(loc);
  if (!de && fl === 'en') return formula;
  return tokenize(formula).map((t) => {
    if (t.t === 'sep') return de ? ';' : t.text;
    if (t.t === 'num') return de ? t.text.replace('.', ',') : t.text;
    if (t.t === 'col') return de ? '.' : t.text;
    if (fl === 'de' && t.t === 'func') return localFn(t.v, fl);
    if (fl === 'de' && t.t === 'bool') return t.v ? 'WAHR' : 'FALSCH';
    if (fl === 'de' && t.t === 'err') return localErr(t.v, fl);
    return t.text;
  }).join('');
}

/** Static precedents of an AST (refs + names + volatile flag) */
const VOLATILE = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'INDIRECT', 'OFFSET', 'CELL', 'INFO']);
export function collectDeps(ast, out = { refs: [], names: [], volatile: false }) {
  if (!ast || typeof ast !== 'object') return out;
  switch (ast.t) {
    case 'ref': out.refs.push(ast); break;
    case 'name': out.names.push(ast); break;
    case 'fn': if (VOLATILE.has(ast.name)) out.volatile = true; for (const a of ast.args) collectDeps(a, out); break;
    case 'sref': (out.srefs || (out.srefs = [])).push(ast); break;
    case 'call': collectDeps(ast.f, out); for (const a of ast.args) collectDeps(a, out); break;
    case 'bin': case 'rangeop': case 'isect': collectDeps(ast.a, out); collectDeps(ast.b, out); break;
    case 'neg': case 'pct': case 'paren': case 'spill': collectDeps(ast.a, out); if (ast.t === 'spill') out.spill = true; break;
    case 'union': for (const a of ast.items) collectDeps(a, out); break;
    default: break;
  }
  return out;
}

/** Functions written with the _xlfn. prefix in OOXML (newer than Excel 2007) */
export const XLFN = new Set(['MAP', 'REDUCE', 'SCAN', 'MAKEARRAY', 'BYROW', 'BYCOL', 'ISOMITTED', 'ANCHORARRAY', 'IMAGE', 'TEXTSPLIT', 'VALUETOTEXT', 'ARRAYTOTEXT', 'IFS', 'IFNA', 'SWITCH', 'MAXIFS', 'MINIFS', 'CONCAT', 'TEXTJOIN', 'XLOOKUP', 'XMATCH', 'UNIQUE', 'SORT', 'SORTBY', 'FILTER', 'SEQUENCE', 'RANDARRAY', 'LET', 'LAMBDA', 'XOR', 'DAYS', 'ISOWEEKNUM', 'NUMBERVALUE', 'AGGREGATE', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'MODE.SNGL', 'MODE.MULT', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'CEILING.MATH', 'FLOOR.MATH', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT', 'VSTACK', 'HSTACK', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND', 'IFERROR_', 'COVARIANCE.P', 'COVARIANCE.S', 'T.TEST', 'T.DIST', 'T.INV', 'CHISQ.TEST', 'F.TEST', 'BINOM.DIST', 'EXPON.DIST', 'GAMMA.DIST', 'POISSON.DIST', 'WEIBULL.DIST', 'LOGNORM.DIST', 'CONFIDENCE.NORM', 'CONFIDENCE.T', 'BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT', 'UNICHAR', 'UNICODE', 'ARABIC', 'BASE', 'DECIMAL', 'COT', 'COTH', 'CSC', 'CSCH', 'SEC', 'SECH', 'ACOT', 'ACOTH', 'COMBINA', 'PERMUTATIONA', 'PDURATION', 'RRI', 'FORMULATEXT', 'ISFORMULA', 'SHEET', 'SHEETS', 'GAMMA', 'GAUSS', 'PHI']);
/** names bound by LET / LAMBDA in an AST (they are written as _xlpm.name in OOXML) */
export function lambdaParams(ast, out = new Set()) {
  if (!ast || typeof ast !== 'object') return out;
  if (ast.t === 'fn') {
    if (ast.name === 'LET') ast.args.forEach((a, i) => { if (i % 2 === 0 && i < ast.args.length - 1 && a.t === 'name') out.add(a.v.toUpperCase()); });
    if (ast.name === 'LAMBDA') ast.args.forEach((a, i) => { if (i < ast.args.length - 1 && a.t === 'name') out.add(a.v.toUpperCase()); });
    for (const a of ast.args) lambdaParams(a, out);
  } else for (const k of ['a', 'b', 'f']) if (ast[k]) lambdaParams(ast[k], out);
  if (ast.items) for (const a of ast.items) lambdaParams(a, out);
  if (ast.t === 'call') for (const a of ast.args) lambdaParams(a, out);
  return out;
}
const XLWS = new Set(['SORT', 'FILTER']);
/** English storage → OOXML formula text (adds _xlfn. / _xlpm. prefixes, A1# → ANCHORARRAY, [@Col] → Table[[#This Row],[Col]]) */
const NEEDS_OOXML = /[#[]|LET|LAMBDA/i;
const FN_CALL = /([A-Za-z_][\w.]*)\s*\(/g;
export function toOoxml(formula, { table } = {}) {
  if (!NEEDS_OOXML.test(formula)) { // fast path: only function prefixes could change
    let need = false; FN_CALL.lastIndex = 0; let q;
    while ((q = FN_CALL.exec(formula))) if (XLFN.has(q[1].toUpperCase())) { need = true; break; }
    if (!need) return formula;
  }
  const toks = tokenize(formula);
  let params = null;
  if (/LET|LAMBDA/i.test(formula)) { try { params = lambdaParams(parse(formula)); } catch { params = null; } }
  const out = [];
  for (const t of toks) {
    if (t.t === 'func') {
      if (params && params.has(t.v)) out.push(`_xlpm.${t.text}`);
      else out.push(XLFN.has(t.v) ? `_xlfn.${XLWS.has(t.v) ? '_xlws.' : ''}${t.v}` : t.v);
    } else if (t.t === 'name' && params && params.has(t.v.toUpperCase()) && !t.sheet) out.push(`_xlpm.${t.text}`);
    else if (t.t === 'hash' && out.length && toks[toks.indexOf(t) - 1] && toks[toks.indexOf(t) - 1].t === 'ref') out.push(`_xlfn.ANCHORARRAY(${out.pop()})`);
    else if (t.t === 'sref') out.push(structuredText(t.table || table || '', t.spec, { ooxml: true }));
    else out.push(t.text);
  }
  return out.join('');
}
/** OOXML formula → storage (prefixes are stripped by the tokenizer) */
export const fromOoxml = (f) => {
  if (!/_xl|\[|ANCHORARRAY/i.test(f || '')) return String(f || ''); // nothing to translate (most formulas)
  const toks = tokenize(String(f || ''));
  const out = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === 'func' && t.v === 'ANCHORARRAY' && toks[i + 1] && toks[i + 1].t === 'lp' && toks[i + 2] && toks[i + 2].t === 'ref' && toks[i + 3] && toks[i + 3].t === 'rp') { out.push(`${toks[i + 2].text}#`); i += 3; continue; }
    out.push(t.t === 'func' ? (t.pm ? t.text : t.v) : t.t === 'sref' ? structuredText(t.table, t.spec) : t.text);
  }
  return out.join('');
};
