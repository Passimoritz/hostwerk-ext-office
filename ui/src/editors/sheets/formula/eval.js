/*
 * AST evaluation. Values: number | string | boolean | null (blank) | FErr | Ref (a range on a sheet) | matrix (any[][]).
 * The context (see calc.js) supplies cell access:
 *   ctx = { sheet, r, c, loc, get(sheet, r, c), sheetByName(name), name(name, sheet) → { ast, sheet } | null,
 *           bounds(sheet) → { rows, cols } (used size, clips whole columns/rows), fn: function table }
 */
import { ERR, isErr, FErr } from './errors.js';
import { MAXR, MAXC } from './refs.js';
import { fmtGeneral } from '../format/numfmt.js';

export class Ref {
  constructor(sheet, r1, c1, r2, c2) { this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.r2 = r2; this.c2 = c2; }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
  get single() { return this.r1 === this.r2 && this.c1 === this.c2; }
}
export const isRef = (v) => v instanceof Ref;
/** LAMBDA value: parameter names, body AST and the LET/LAMBDA names visible where it was made */
export class Lambda {
  constructor(params, body, locals) { this.params = params; this.body = body; this.locals = locals; }
}
export const isLambda = (v) => v instanceof Lambda;
/** marker for a LAMBDA parameter that was left out (ISOMITTED) */
export const OMITTED = Object.freeze({ omitted: true });
/** call a LAMBDA value with already evaluated arguments */
export function callLambda(fn, args, ctx) {
  if (!isLambda(fn)) return ERR.VALUE;
  if (args.length > fn.params.length) return ERR.VALUE;
  if ((ctx.depth = (ctx.depth || 0) + 1) > 200) { ctx.depth--; return ERR.CALC; }
  try {
    const locals = new Map(fn.locals || []);
    fn.params.forEach((p, i) => locals.set(p, i < args.length && args[i] !== undefined ? args[i] : OMITTED));
    return evaluate(fn.body, { ...ctx, locals });
  } finally { ctx.depth--; }
}
/** the cell range of a structured reference (Table1[Col], [@Col], Table1[[#Headers],[A]:[B]]) or an error */
export function structuredRef(n, ctx) {
  const t = ctx.table ? ctx.table(n.table, ctx.sheet, ctx.r, ctx.c) : null;
  if (!t) return ERR.REF;
  const g = t.range;
  const hdr = t.header ? 1 : 0; const tot = t.totals ? 1 : 0;
  const data = { r1: g.r1 + hdr, r2: g.r2 - tot };
  let c1 = g.c1; let c2 = g.c2;
  if (n.c1 != null) {
    const ix = (name) => t.columns.findIndex((x) => String(x).toLowerCase() === String(name).toLowerCase());
    const a = ix(n.c1); const b = ix(n.c2 == null ? n.c1 : n.c2);
    if (a < 0 || b < 0) return ERR.REF;
    c1 = g.c1 + Math.min(a, b); c2 = g.c1 + Math.max(a, b);
  }
  const sp = n.specials || [];
  let r1; let r2;
  if (sp.includes('thisRow')) {
    if (ctx.sheet !== t.sheet || ctx.r < data.r1 || ctx.r > data.r2) return ERR.VALUE;
    r1 = ctx.r; r2 = ctx.r;
  } else if (!sp.length || (sp.length === 1 && sp[0] === 'data')) { r1 = data.r1; r2 = data.r2; } else if (sp.includes('all')) { r1 = g.r1; r2 = g.r2; } else {
    const want = new Set(sp);
    r1 = want.has('headers') ? (hdr ? g.r1 : null) : want.has('data') ? data.r1 : want.has('totals') && tot ? g.r2 : null;
    r2 = want.has('totals') ? (tot ? g.r2 : null) : want.has('data') ? data.r2 : want.has('headers') && hdr ? g.r1 : null;
    if (r1 == null || r2 == null) return ERR.REF;
  }
  if (r2 < r1) return ERR.REF;
  return new Ref(t.sheet, r1, c1, r2, c2);
}
export const isMat = (v) => Array.isArray(v);

/** clip whole columns/rows to the used area so A:A doesn't create a million values */
export function clip(ref, ctx) {
  if (ref.r2 - ref.r1 < 5000 && ref.c2 - ref.c1 < 500) return ref;
  const b = ctx.bounds(ref.sheet);
  const r2 = Math.min(ref.r2, Math.max(ref.r1, b.rows - 1));
  const c2 = Math.min(ref.c2, Math.max(ref.c1, b.cols - 1));
  return r2 === ref.r2 && c2 === ref.c2 ? ref : new Ref(ref.sheet, ref.r1, ref.c1, r2, c2);
}

/** 2D values of a ref (clipped) */
export function refValues(ref, ctx) {
  const g = clip(ref, ctx);
  const out = new Array(g.rows);
  for (let r = 0; r < g.rows; r++) {
    const row = new Array(g.cols);
    for (let c = 0; c < g.cols; c++) row[c] = ctx.get(g.sheet, g.r1 + r, g.c1 + c);
    out[r] = row;
  }
  return out;
}

/** Ref → matrix, matrix stays, scalar → [[v]] */
export function toMatrix(v, ctx) {
  if (isRef(v)) return refValues(v, ctx);
  if (isMat(v)) return v;
  return [[v]];
}

/** Single value: 1-cell ref → its value; bigger ref → #VALUE! (no implicit intersection); 1x1 matrix → value */
export function scalar(v, ctx) {
  if (isRef(v)) return v.single ? ctx.get(v.sheet, v.r1, v.c1) : implicit(v, ctx);
  if (isMat(v)) return v.length && v[0].length ? v[0][0] : ERR.VALUE;
  return v;
}
function implicit(ref, ctx) {
  // implicit intersection (legacy behaviour for functions that need one value)
  if (ref.c1 === ref.c2 && ctx.r >= ref.r1 && ctx.r <= ref.r2 && ref.sheet === ctx.sheet) return ctx.get(ref.sheet, ctx.r, ref.c1);
  if (ref.r1 === ref.r2 && ctx.c >= ref.c1 && ctx.c <= ref.c2 && ref.sheet === ctx.sheet) return ctx.get(ref.sheet, ref.r1, ctx.c);
  return ERR.VALUE;
}

export function toNum(v, loc) {
  if (typeof v === 'number') return v;
  if (v == null || v === '') return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) return v;
  if (typeof v === 'string') {
    const n = textToNumber(v, loc);
    return n == null ? ERR.VALUE : n;
  }
  return ERR.VALUE;
}
/** "12", "1.5e3", "50%", "1,5" (de) → number; null if it isn't one */
export function textToNumber(s, loc) {
  let t = s.trim();
  if (!t) return null;
  let pct = false;
  if (t.endsWith('%')) { pct = true; t = t.slice(0, -1).trim(); }
  if (loc && loc.decimal === ',') t = t.replace(/\./g, '').replace(',', '.'); else t = t.replace(/,(?=\d{3}(\D|$))/g, '');
  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return null;
  const n = Number(t);
  return pct ? n / 100 : n;
}
export function toStr(v, loc) {
  if (typeof v === 'string') return v;
  if (v == null) return '';
  if (typeof v === 'number') return fmtGeneral(v, loc);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isErr(v)) return v;
  return String(v);
}
export function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v == null || v === '') return false;
  if (isErr(v)) return v;
  if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE' || u === 'WAHR') return true; if (u === 'FALSE' || u === 'FALSCH') return false; }
  return ERR.VALUE;
}

/* comparison like Excel: numbers < text < booleans, text case-insensitive, blank = 0 / "" / FALSE */
export function compare(a, b) {
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ta = typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : 2;
  const tb = typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : 2;
  if (ta !== tb) return ta - tb;
  if (ta === 1) { const x = a.toLowerCase(); const y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  if (ta === 2) return (a ? 1 : 0) - (b ? 1 : 0);
  return a < b ? -1 : a > b ? 1 : 0;
}

/* Floating point cleanup like Excel (15 significant digits) */
export const clean = (n) => (typeof n === 'number' && Number.isFinite(n) && n !== 0 && Math.abs(n) < 1e300 ? Number(n.toPrecision(15)) : n);

function binop(op, a, b, loc) {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  switch (op) {
    case '&': { return toStr(a, loc) + toStr(b, loc); }
    case '=': return compare(a, b) === 0;
    case '<>': return compare(a, b) !== 0;
    case '<': return compare(a, b) < 0;
    case '>': return compare(a, b) > 0;
    case '<=': return compare(a, b) <= 0;
    case '>=': return compare(a, b) >= 0;
    default: break;
  }
  const x = toNum(a, loc); if (isErr(x)) return x;
  const y = toNum(b, loc); if (isErr(y)) return y;
  let r;
  switch (op) {
    case '+': r = x + y; break;
    case '-': r = x - y; break;
    case '*': r = x * y; break;
    case '/': if (y === 0) return ERR.DIV0; r = x / y; break;
    case '^': if (x === 0 && y === 0) return ERR.NUM; r = Math.pow(x, y); break;
    default: return ERR.VALUE;
  }
  if (!Number.isFinite(r)) return ERR.NUM;
  return op === '+' || op === '-' ? clean(r) : r;
}

/** elementwise application with broadcasting (1×n against m×1 etc.) */
export function broadcast(a, b, f) {
  const ra = a.length; const ca = a[0].length; const rb = b.length; const cb = b[0].length;
  const R = Math.max(ra, rb); const C = Math.max(ca, cb);
  const out = [];
  for (let i = 0; i < R; i++) {
    const row = [];
    for (let j = 0; j < C; j++) {
      const x = ra === 1 ? (ca === 1 ? a[0][0] : a[0][j]) : ca === 1 ? (a[i] ? a[i][0] : undefined) : a[i] && a[i][j];
      const y = rb === 1 ? (cb === 1 ? b[0][0] : b[0][j]) : cb === 1 ? (b[i] ? b[i][0] : undefined) : b[i] && b[i][j];
      row.push(x === undefined || y === undefined ? ERR.NA : f(x, y));
    }
    out.push(row);
  }
  return out;
}
export const mapMat = (m, f) => m.map((row) => row.map(f));

/** Evaluate; returns Ref for references (callers decide scalar/matrix). */
export function evaluate(n, ctx) {
  switch (n.t) {
    case 'num': return n.v;
    case 'str': return n.v;
    case 'bool': return n.v;
    case 'err': return n.v;
    case 'missing': return null;
    case 'paren': return evaluate(n.a, ctx);
    case 'array': return n.rows.map((r) => r.slice());
    case 'ref': {
      const sh = n.sheet ? ctx.sheetByName(n.sheet) : ctx.sheet;
      if (!sh) return ERR.REF;
      if (n.sheet2) { // 3D reference Sheet1:Sheet3!A1 → union over the sheets in between
        const list = ctx.sheetsBetween ? ctx.sheetsBetween(n.sheet, n.sheet2) : null;
        if (!list || !list.length) return ERR.REF;
        const u = list.map((x) => new Ref(x, Math.min(n.r1, n.r2), Math.min(n.c1, n.c2), Math.max(n.r1, n.r2), Math.max(n.c1, n.c2)));
        if (u.length === 1) return u[0];
        u.union = true; u.three = true;
        return u;
      }
      return new Ref(sh, n.r1 > n.r2 ? n.r2 : n.r1, n.c1 > n.c2 ? n.c2 : n.c1, Math.max(n.r1, n.r2), Math.max(n.c1, n.c2));
    }
    case 'sref': return structuredRef(n, ctx);
    case 'isect': {
      const a = evaluate(n.a, ctx); const b = evaluate(n.b, ctx);
      if (!isRef(a) || !isRef(b)) return isErr(a) ? a : isErr(b) ? b : ERR.VALUE;
      if (a.sheet !== b.sheet) return ERR.VALUE;
      const r1 = Math.max(a.r1, b.r1); const r2 = Math.min(a.r2, b.r2); const c1 = Math.max(a.c1, b.c1); const c2 = Math.min(a.c2, b.c2);
      return r1 > r2 || c1 > c2 ? ERR.NULL : new Ref(a.sheet, r1, c1, r2, c2);
    }
    case 'call': {
      const f = evaluate(n.f, ctx);
      if (isErr(f)) return f;
      return callLambda(f, n.args.map((x) => (x.t === 'missing' ? undefined : evaluate(x, ctx))), ctx);
    }
    case 'name': {
      if (ctx.locals && !n.sheet && ctx.locals.has(n.v.toUpperCase())) { const v = ctx.locals.get(n.v.toUpperCase()); return v === OMITTED ? null : v; }
      const def = ctx.name(n.v, n.sheet ? ctx.sheetByName(n.sheet) : ctx.sheet);
      if (!def) return ERR.NAME;
      if ((ctx.depth = (ctx.depth || 0) + 1) > 50) { ctx.depth--; return ERR.CALC; }
      try { return evaluate(def.ast, def.sheet ? { ...ctx, sheet: def.sheet } : ctx); } finally { ctx.depth--; }
    }
    case 'neg': case 'pct': {
      const v = evaluate(n.a, ctx);
      const f = (x) => { const y = toNum(x, ctx.loc); return isErr(y) ? y : n.t === 'neg' ? -y : y / 100; };
      if (isRef(v) && !v.single) return mapMat(refValues(v, ctx), f);
      if (isMat(v)) return mapMat(v, f);
      return f(scalar(v, ctx));
    }
    case 'bin': {
      const a = evaluate(n.a, ctx);
      const b = evaluate(n.b, ctx);
      const am = (isRef(a) && !a.single) || isMat(a);
      const bm = (isRef(b) && !b.single) || isMat(b);
      if (am || bm) return broadcast(toMatrix(am ? a : scalar(a, ctx), ctx), toMatrix(bm ? b : scalar(b, ctx), ctx), (x, y) => binop(n.op, x, y, ctx.loc));
      return binop(n.op, scalar(a, ctx), scalar(b, ctx), ctx.loc);
    }
    case 'rangeop': {
      const a = evaluate(n.a, ctx); const b = evaluate(n.b, ctx);
      if (!isRef(a) || !isRef(b)) return isErr(a) ? a : isErr(b) ? b : ERR.VALUE;
      if (a.sheet !== b.sheet) return ERR.VALUE;
      return new Ref(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
    }
    case 'spill': {
      const a = evaluate(n.a, ctx);
      if (!isRef(a)) return ERR.REF;
      const sp = ctx.spillOf && ctx.spillOf(a.sheet, a.r1, a.c1);
      if (!sp) return ERR.REF;
      return new Ref(a.sheet, a.r1, a.c1, a.r1 + sp.rows - 1, a.c1 + sp.cols - 1);
    }
    case 'union': {
      const items = n.items.map((x) => evaluate(x, ctx));
      const bad = items.find((x) => !isRef(x));
      if (bad !== undefined) return isErr(bad) ? bad : ERR.VALUE;
      const u = items.slice(); u.union = true; // functions that accept unions flatten it
      return u;
    }
    case 'fn': {
      // a LET/LAMBDA name or a defined name holding a LAMBDA, called like a function
      const local = ctx.locals && ctx.locals.get(n.name);
      if (local !== undefined && isLambda(local)) return callLambda(local, n.args.map((x) => (x.t === 'missing' ? undefined : evaluate(x, ctx))), ctx);
      const def = ctx.fn[n.name];
      if (!def) {
        const nd = ctx.name && ctx.name(n.name, ctx.sheet);
        if (!nd) return ERR.NAME;
        const f = evaluate(nd.ast, nd.sheet ? { ...ctx, sheet: nd.sheet } : ctx);
        if (!isLambda(f)) return isErr(f) ? f : ERR.VALUE;
        return callLambda(f, n.args.map((x) => (x.t === 'missing' ? undefined : evaluate(x, ctx))), ctx);
      }
      if (def.lazy) return def.fn(n.args, ctx);
      const args = n.args.map((x) => evaluate(x, ctx));
      return def.fn(args, ctx, n);
    }
    default: return ERR.VALUE;
  }
}

/** Final result of a cell formula: Ref → value or matrix; 1×1 matrix → value; blank → 0 */
export function finalize(v, ctx) {
  if (isLambda(v) || v === OMITTED) return ERR.CALC;
  if (isRef(v)) v = v.single ? ctx.get(v.sheet, v.r1, v.c1) : refValues(v, ctx);
  if (isMat(v)) {
    if (v.union) return ERR.VALUE;
    if (!v.length || !v[0].length) return ERR.CALC;
    if (v.length === 1 && v[0].length === 1) v = v[0][0];
    else return v.map((row) => row.map((x) => (x == null ? 0 : x)));
  }
  if (v == null) return 0;
  if (typeof v === 'number' && !Number.isFinite(v)) return ERR.NUM;
  if (v instanceof FErr) return v;
  return v;
}

export { MAXR, MAXC };
