/*
 * Function table. Own implementations for everything where Excel's argument semantics matter (aggregates, criteria,
 * lookups, references, dynamic arrays, dates, text formatting); the rest comes from @formulajs/formulajs through a
 * generic wrapper (ranges → arrays, element-wise lifting over arrays, Error/Date results converted).
 * Function names are English here (storage); German names are mapped for input/display in formula/i18n.js.
 */
import * as FJ from '@formulajs/formulajs';
import { ERR, isErr, errOf, ERROR_TYPE } from './errors.js';
import { Ref, isRef, isMat, scalar, toNum, toStr, toBool, compare, refValues, toMatrix, clip, evaluate, clean, textToNumber, Lambda, isLambda, callLambda, OMITTED } from './eval.js';
import { parseRange, colName, quoteSheet, MAXR, MAXC } from './refs.js';
import { formatValue } from '../format/numfmt.js';
import { parseInput } from '../format/parse.js';
import { serialFromParts, partsFromSerial, todaySerial, nowSerial, dateToSerial, daysInMonth } from '../format/dates.js';

const FN = {};
const def = (names, fn, opts = {}) => { for (const n of names.split(' ')) FN[n] = { fn, ...opts }; };
const lazy = (names, fn) => def(names, fn, { lazy: true });

/* ---------------- helpers ---------------- */
const firstErr = (arr) => arr.find(isErr);
/** iterate all values of args; cb(value, direct) – direct = typed argument (not from a range/array) */
function each(args, ctx, cb) {
  for (const a of args) {
    if (isRef(a)) { const g = clip(a, ctx); for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { if (cb(ctx.get(g.sheet, r, c), false, g.sheet, r) === false) return; } } else if (isMat(a)) {
      if (a.union) { if (each(a, ctx, cb) === false) return; continue; }
      for (const row of a) for (const v of row) if (cb(v, false) === false) return;
    } else if (cb(a, true) === false) return;
  }
}
/** numbers of args with Excel aggregate rules; returns number[] or FErr */
function nums(args, ctx, { all = false } = {}) {
  const out = [];
  let err = null;
  each(args, ctx, (v, direct) => {
    if (isErr(v)) { err = v; return false; }
    if (typeof v === 'number') out.push(v);
    else if (direct) {
      if (v == null) { if (all) out.push(0); return undefined; }
      const n = toNum(v, ctx.loc);
      if (isErr(n)) { err = n; return false; }
      out.push(n);
    } else if (all) { if (typeof v === 'boolean') out.push(v ? 1 : 0); else if (typeof v === 'string') out.push(0); }
    return undefined;
  });
  return err || out;
}
const S = (v, ctx) => scalar(v, ctx);
const N = (v, ctx) => { const x = scalar(v, ctx); return toNum(x, ctx.loc); };
const T = (v, ctx) => { const x = scalar(v, ctx); return toStr(x, ctx.loc); };
const B = (v, ctx) => toBool(scalar(v, ctx));
const isMulti = (v) => (isRef(v) && !v.single) || (isMat(v) && !(v.length === 1 && v[0].length === 1));

/** element-wise lifting: when an argument is a range/array where a scalar is expected, map over it */
function lifted(fn, scalarIdx) {
  return (args, ctx, node) => {
    const idx = scalarIdx || args.map((_, i) => i);
    const multi = idx.filter((i) => i < args.length && isMulti(args[i]));
    if (!multi.length) return fn(args.map((a, i) => (idx.includes(i) ? S(a, ctx) : a)), ctx, node);
    const mats = {};
    let R = 1; let C = 1;
    for (const i of multi) { mats[i] = toMatrix(args[i], ctx); R = Math.max(R, mats[i].length); C = Math.max(C, mats[i][0].length); }
    const out = [];
    for (let r = 0; r < R; r++) {
      const row = [];
      for (let c = 0; c < C; c++) {
        const a2 = args.map((a, i) => {
          if (!(i in mats)) return idx.includes(i) ? S(a, ctx) : a;
          const m = mats[i];
          const rr = m.length === 1 ? 0 : r; const cc = m[0].length === 1 ? 0 : c;
          return m[rr] && m[rr][cc] !== undefined ? m[rr][cc] : ERR.NA;
        });
        row.push(fn(a2, ctx, node));
      }
      out.push(row);
    }
    return out;
  };
}
/** scalar function with numeric args */
function numfn(names, f, min = 1) {
  def(names, lifted((args, ctx) => {
    if (args.length < min) return ERR.VALUE;
    const xs = [];
    for (const a of args) { const x = a == null ? null : toNum(a, ctx.loc); if (isErr(x)) return x; xs.push(x); }
    const r = f(...xs);
    return typeof r === 'number' ? (Number.isFinite(r) ? r : ERR.NUM) : r;
  }));
}

/* ---------------- math ---------------- */
const roundHalf = (x, d = 0) => {
  const m = 10 ** d;
  const v = Math.abs(x) * m;
  const r = Math.round(Number(v.toPrecision(15))) / m;
  return x < 0 ? -r : r;
};
numfn('ROUND', (x, d) => roundHalf(x, Math.trunc(d || 0)), 1);
numfn('ROUNDUP', (x, d) => { const m = 10 ** Math.trunc(d || 0); const v = Number((Math.abs(x) * m).toPrecision(15)); return Math.sign(x) * Math.ceil(v) / m; });
numfn('ROUNDDOWN', (x, d) => { const m = 10 ** Math.trunc(d || 0); const v = Number((Math.abs(x) * m).toPrecision(15)); return Math.sign(x) * Math.floor(v) / m; });
numfn('TRUNC', (x, d) => { const m = 10 ** Math.trunc(d || 0); return Math.trunc(Number((x * m).toPrecision(15))) / m; });
numfn('INT', (x) => Math.floor(x));
numfn('MOD', (n, d) => (d === 0 ? ERR.DIV0 : clean(n - d * Math.floor(n / d))), 2);
numfn('ABS', Math.abs);
numfn('SQRT', (x) => (x < 0 ? ERR.NUM : Math.sqrt(x)));
numfn('POWER', (x, y) => (x === 0 && y === 0 ? ERR.NUM : Math.pow(x, y)), 2);
numfn('EXP', Math.exp);
numfn('LN', (x) => (x <= 0 ? ERR.NUM : Math.log(x)));
numfn('LOG10', (x) => (x <= 0 ? ERR.NUM : Math.log10(x)));
numfn('LOG', (x, b) => (x <= 0 || (b != null && b <= 0) ? ERR.NUM : Math.log(x) / Math.log(b == null ? 10 : b)));
numfn('SIGN', Math.sign);
numfn('EVEN', (x) => { const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; });
numfn('ODD', (x) => { let r = Math.ceil(Math.abs(x)); if (r % 2 === 0) r += 1; return x < 0 ? -r : r; });
numfn('MROUND', (x, m) => (m === 0 ? 0 : x * m < 0 ? ERR.NUM : roundHalf(x / m) * m), 2);
numfn('CEILING CEILING.PRECISE ISO.CEILING', (x, s = 1) => (s == null || s === 0 ? 0 : Math.ceil(Number((x / Math.abs(s)).toPrecision(15))) * Math.abs(s)));
numfn('FLOOR FLOOR.PRECISE', (x, s = 1) => (s === 0 ? ERR.DIV0 : Math.floor(Number((x / Math.abs(s)).toPrecision(15))) * Math.abs(s)));
numfn('CEILING.MATH', (x, s = 1, mode = 0) => { s = Math.abs(s || 1); return x < 0 && mode ? -Math.ceil(Number((-x / s).toPrecision(15))) * s : Math.ceil(Number((x / s).toPrecision(15))) * s; });
numfn('FLOOR.MATH', (x, s = 1, mode = 0) => { s = Math.abs(s || 1); return x < 0 && mode ? -Math.floor(Number((-x / s).toPrecision(15))) * s : Math.floor(Number((x / s).toPrecision(15))) * s; });
numfn('PI', () => Math.PI, 0);
def('RAND', () => Math.random());
numfn('RANDBETWEEN', (a, b) => Math.floor(Math.random() * (Math.floor(b) - Math.ceil(a) + 1)) + Math.ceil(a), 2);
numfn('QUOTIENT', (a, b) => (b === 0 ? ERR.DIV0 : Math.trunc(a / b)), 2);

def('SUM', (args, ctx) => { const xs = nums(args, ctx); return isErr(xs) ? xs : clean(xs.reduce((s, x) => s + x, 0)); });
def('PRODUCT', (args, ctx) => { const xs = nums(args, ctx); return isErr(xs) ? xs : xs.length ? xs.reduce((s, x) => s * x, 1) : 0; });
def('SUMSQ', (args, ctx) => { const xs = nums(args, ctx); return isErr(xs) ? xs : xs.reduce((s, x) => s + x * x, 0); });
def('AVERAGE', (args, ctx) => { const xs = nums(args, ctx); return isErr(xs) ? xs : xs.length ? clean(xs.reduce((s, x) => s + x, 0) / xs.length) : ERR.DIV0; });
def('AVERAGEA', (args, ctx) => { const xs = nums(args, ctx, { all: true }); return isErr(xs) ? xs : xs.length ? xs.reduce((s, x) => s + x, 0) / xs.length : ERR.DIV0; });
def('MIN', (args, ctx) => { const xs = nums(args, ctx); return isErr(xs) ? xs : xs.length ? Math.min(...xs) : 0; });
def('MAX', (args, ctx) => { const xs = nums(args, ctx); return isErr(xs) ? xs : xs.length ? Math.max(...xs) : 0; });
def('MINA', (args, ctx) => { const xs = nums(args, ctx, { all: true }); return isErr(xs) ? xs : xs.length ? Math.min(...xs) : 0; });
def('MAXA', (args, ctx) => { const xs = nums(args, ctx, { all: true }); return isErr(xs) ? xs : xs.length ? Math.max(...xs) : 0; });
def('COUNT', (args, ctx) => { let n = 0; each(args, ctx, (v, d) => { if (typeof v === 'number' || (d && (typeof v === 'boolean' || (typeof v === 'string' && textToNumber(v, ctx.loc) != null)))) n++; }); return n; });
def('COUNTA', (args, ctx) => { let n = 0; each(args, ctx, (v, d) => { if (v != null || (d && false)) n++; }); return n; });
def('COUNTBLANK', (args, ctx) => { let n = 0; each(args, ctx, (v) => { if (v == null || v === '') n++; }); return n; });
const stat = (xs, kind) => {
  const n = xs.length;
  const mean = xs.reduce((s, x) => s + x, 0) / n;
  const ss = xs.reduce((s, x) => s + (x - mean) ** 2, 0);
  if (kind === 'vs') return n < 2 ? ERR.DIV0 : ss / (n - 1);
  if (kind === 'vp') return n < 1 ? ERR.DIV0 : ss / n;
  if (kind === 'ss') return n < 2 ? ERR.DIV0 : Math.sqrt(ss / (n - 1));
  return n < 1 ? ERR.DIV0 : Math.sqrt(ss / n);
};
def('STDEV STDEV.S', (a, c) => { const xs = nums(a, c); return isErr(xs) ? xs : stat(xs, 'ss'); });
def('STDEVP STDEV.P', (a, c) => { const xs = nums(a, c); return isErr(xs) ? xs : stat(xs, 'sp'); });
def('VAR VAR.S', (a, c) => { const xs = nums(a, c); return isErr(xs) ? xs : stat(xs, 'vs'); });
def('VARP VAR.P', (a, c) => { const xs = nums(a, c); return isErr(xs) ? xs : stat(xs, 'vp'); });
def('MEDIAN', (a, c) => { const xs = nums(a, c); if (isErr(xs)) return xs; if (!xs.length) return ERR.NUM; xs.sort((x, y) => x - y); const m = xs.length >> 1; return xs.length % 2 ? xs[m] : (xs[m - 1] + xs[m]) / 2; });
def('LARGE', (a, c) => { const xs = nums([a[0]], c); const k = N(a[1], c); if (isErr(xs)) return xs; if (isErr(k)) return k; xs.sort((x, y) => y - x); return k < 1 || k > xs.length ? ERR.NUM : xs[Math.ceil(k) - 1]; });
def('SMALL', (a, c) => { const xs = nums([a[0]], c); const k = N(a[1], c); if (isErr(xs)) return xs; if (isErr(k)) return k; xs.sort((x, y) => x - y); return k < 1 || k > xs.length ? ERR.NUM : xs[Math.ceil(k) - 1]; });
def('RANK RANK.EQ', (a, c) => { const v = N(a[0], c); const xs = nums([a[1]], c); const asc = a[2] != null && N(a[2], c) ? 1 : 0; if (isErr(v)) return v; if (isErr(xs)) return xs; if (!xs.includes(v)) return ERR.NA; return 1 + xs.filter((x) => (asc ? x < v : x > v)).length; });
def('SUMPRODUCT', (args, ctx) => {
  const ms = args.map((a) => toMatrix(a, ctx));
  if (!ms.length) return ERR.VALUE;
  const R = ms[0].length; const C = ms[0][0].length;
  if (ms.some((m) => m.length !== R || m[0].length !== C)) return ERR.VALUE;
  let s = 0;
  for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
    let p = 1;
    for (const m of ms) { const v = m[r][c]; if (isErr(v)) return v; p *= typeof v === 'number' ? v : typeof v === 'boolean' && ms.length === 1 ? 0 : 0; }
    s += p;
  }
  return clean(s);
});

/* ---------------- logical ---------------- */
lazy('IF', (n, ctx) => {
  const c = scalarOrMat(evaluate(n[0], ctx), ctx);
  const pick = (b) => (b ? (n[1] ? evaluate(n[1], ctx) : true) : n.length > 2 ? (n[2].t === 'missing' ? 0 : evaluate(n[2], ctx)) : false);
  if (isMat(c)) { const a = toMatrix(evaluate(n[1] || { t: 'bool', v: true }, ctx), ctx); const b = n[2] ? toMatrix(evaluate(n[2], ctx), ctx) : [[false]]; return c.map((row, i) => row.map((v, j) => { const t = toBool(v); if (isErr(t)) return t; const m = t ? a : b; return (m[i] || m[0])[j] !== undefined ? (m[i] || m[0])[j] : (m[i] || m[0])[0]; })); }
  const b = toBool(c);
  return isErr(b) ? b : pick(b);
});
function scalarOrMat(v, ctx) { return isMulti(v) ? toMatrix(v, ctx) : S(v, ctx); }
lazy('IFS', (n, ctx) => { for (let i = 0; i + 1 < n.length; i += 2) { const b = B(evaluate(n[i], ctx), ctx); if (isErr(b)) return b; if (b) return evaluate(n[i + 1], ctx); } return ERR.NA; });
lazy('IFERROR', (n, ctx) => { const v = evaluate(n[0], ctx); if (isMulti(v)) { const m = toMatrix(v, ctx); return m.map((row) => row.map((x) => (isErr(x) ? S(evaluate(n[1], ctx), ctx) : x))); } const s = S(v, ctx); return isErr(s) ? evaluate(n[1], ctx) : v; });
lazy('IFNA', (n, ctx) => { const v = evaluate(n[0], ctx); const s = isMulti(v) ? null : S(v, ctx); return isErr(s) && s.code === '#N/A' ? evaluate(n[1], ctx) : v; });
def('AND', (args, ctx) => { let r = true; let any = false; let err = null; each(args, ctx, (v, d) => { if (isErr(v)) { err = v; return false; } if (v == null || (typeof v === 'string' && !d)) return undefined; const b = toBool(v); if (isErr(b)) { err = b; return false; } any = true; if (!b) r = false; return undefined; }); return err || (any ? r : ERR.VALUE); });
def('OR', (args, ctx) => { let r = false; let any = false; let err = null; each(args, ctx, (v, d) => { if (isErr(v)) { err = v; return false; } if (v == null || (typeof v === 'string' && !d)) return undefined; const b = toBool(v); if (isErr(b)) { err = b; return false; } any = true; if (b) r = true; return undefined; }); return err || (any ? r : ERR.VALUE); });
def('XOR', (args, ctx) => { let r = false; let err = null; each(args, ctx, (v) => { if (isErr(v)) { err = v; return false; } if (v == null || typeof v === 'string') return undefined; if (toBool(v)) r = !r; return undefined; }); return err || r; });
def('NOT', lifted((a) => { const b = toBool(a[0]); return isErr(b) ? b : !b; }));
def('TRUE', () => true);
def('FALSE', () => false);
lazy('SWITCH', (n, ctx) => { const v = S(evaluate(n[0], ctx), ctx); if (isErr(v)) return v; let i = 1; for (; i + 1 < n.length; i += 2) { const c = S(evaluate(n[i], ctx), ctx); if (compare(v, c) === 0 && typeof v === typeof c) return evaluate(n[i + 1], ctx); } return i < n.length ? evaluate(n[i], ctx) : ERR.NA; });
lazy('CHOOSE', (n, ctx) => { const i = N(evaluate(n[0], ctx), ctx); if (isErr(i)) return i; const k = Math.trunc(i); return k < 1 || k >= n.length ? ERR.VALUE : evaluate(n[k], ctx); });
lazy('LET', (n, ctx) => {
  if (n.length < 3 || n.length % 2 === 0) return ERR.VALUE;
  const locals = new Map(ctx.locals || []);
  const c2 = { ...ctx, locals };
  for (let i = 0; i + 1 < n.length; i += 2) { if (n[i].t !== 'name') return ERR.VALUE; locals.set(n[i].v.toUpperCase(), evaluate(n[i + 1], c2)); }
  return evaluate(n[n.length - 1], c2);
});

/* ---------------- LAMBDA helpers ---------------- */
lazy('LAMBDA', (n, ctx) => {
  if (!n.length) return ERR.VALUE;
  const params = [];
  for (let i = 0; i < n.length - 1; i++) { const p = n[i].t === 'name' ? n[i].v : n[i].t === 'sref' && !n[i].table && n[i].c1 && !n[i].specials.length ? n[i].c1 : null; if (!p) return ERR.VALUE; params.push(p.toUpperCase()); } // [y] = optional
  if (new Set(params).size !== params.length) return ERR.VALUE;
  return new Lambda(params, n[n.length - 1], ctx.locals || null);
});
lazy('ISOMITTED', (n, ctx) => (n[0] && n[0].t === 'name' && ctx.locals ? ctx.locals.get(n[0].v.toUpperCase()) === OMITTED : false));
const lamArg = (v) => (isLambda(v) ? v : null);
const callV = (f, args, ctx) => { const r = callLambda(f, args, ctx); return isRef(r) ? (r.single ? ctx.get(r.sheet, r.r1, r.c1) : refValues(r, ctx)) : r; };
const one = (v, ctx) => { if (isMat(v)) return v.length === 1 && v[0].length === 1 ? v[0][0] : ERR.CALC; return v; };
def('MAP', (a, ctx) => {
  const f = lamArg(a[a.length - 1]); if (!f || a.length < 2) return ERR.VALUE;
  const ms = a.slice(0, -1).map((x) => toMatrix(x, ctx));
  const R = Math.max(...ms.map((m) => m.length)); const C = Math.max(...ms.map((m) => m[0].length));
  const out = [];
  for (let r = 0; r < R; r++) { const row = []; for (let c = 0; c < C; c++) { const args = ms.map((m) => (m.length === 1 && m[0].length === 1 ? m[0][0] : m[r] && m[r][c] !== undefined ? m[r][c] : ERR.NA)); row.push(one(callV(f, args, ctx), ctx)); } out.push(row); }
  return out;
});
def('REDUCE', (a, ctx) => {
  const f = lamArg(a[2]); if (!f) return ERR.VALUE;
  let acc = a[0] === undefined ? null : isRef(a[0]) ? scalar(a[0], ctx) : a[0];
  for (const row of toMatrix(a[1], ctx)) for (const v of row) { acc = callV(f, [acc, v], ctx); }
  return acc;
});
def('SCAN', (a, ctx) => {
  const f = lamArg(a[2]); if (!f) return ERR.VALUE;
  let acc = isRef(a[0]) ? scalar(a[0], ctx) : a[0];
  return toMatrix(a[1], ctx).map((row) => row.map((v) => { acc = one(callV(f, [acc, v], ctx), ctx); return acc; }));
});
def('MAKEARRAY', (a, ctx) => {
  const rows = toNum(scalar(a[0], ctx), ctx.loc); const cols = toNum(scalar(a[1], ctx), ctx.loc); const f = lamArg(a[2]);
  if (isErr(rows)) return rows; if (isErr(cols)) return cols; if (!f) return ERR.VALUE;
  const R = Math.trunc(rows); const C = Math.trunc(cols);
  if (R < 1 || C < 1 || R * C > 1e6) return ERR.VALUE;
  const out = [];
  for (let r = 1; r <= R; r++) { const row = []; for (let c = 1; c <= C; c++) row.push(one(callV(f, [r, c], ctx), ctx)); out.push(row); }
  return out;
});
def('BYROW', (a, ctx) => { const f = lamArg(a[1]); if (!f) return ERR.VALUE; return toMatrix(a[0], ctx).map((row) => [one(callV(f, [[row]], ctx), ctx)]); });
def('BYCOL', (a, ctx) => { const f = lamArg(a[1]); if (!f) return ERR.VALUE; const m = toMatrix(a[0], ctx); return [m[0].map((_, c) => one(callV(f, [m.map((row) => [row[c]])], ctx), ctx))]; });

/* ---------------- criteria (COUNTIF …) ---------------- */
const wildRe = (s) => new RegExp(`^${s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '[\\s\\S]').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?')}$`, 'i');
export function criteria(crit, loc) {
  if (isErr(crit)) return (v) => isErr(v) && v.code === crit.code;
  if (typeof crit === 'number') return (v) => (typeof v === 'number' ? v === crit : typeof v === 'string' && textToNumber(v, loc) === crit);
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit == null) return (v) => v == null || v === '';
  const s = String(crit);
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(s);
  const op = m[1] || '=';
  const rest = m[2];
  if (rest === '') return op === '=' ? (v) => v == null || v === '' : op === '<>' ? (v) => !(v == null || v === '') : () => false;
  let n = textToNumber(rest, loc);
  if (n == null) { const p = parseInput(rest, loc); if (typeof p.v === 'number' && p.fmt) n = p.v; }
  if (n != null) {
    const f = { '=': (x) => x === n, '<>': (x) => x !== n, '<': (x) => x < n, '>': (x) => x > n, '<=': (x) => x <= n, '>=': (x) => x >= n }[op];
    return (v) => (typeof v === 'number' ? f(v) : op === '<>');
  }
  const up = rest.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') { const b = up === 'TRUE'; return op === '<>' ? (v) => v !== b : (v) => v === b; }
  if (/^#/.test(rest)) return (v) => (isErr(v) && v.code === up) === (op !== '<>');
  if (op === '=' || op === '<>') {
    const re = /[*?]/.test(rest) ? wildRe(rest) : null;
    const lo = rest.toLowerCase();
    const eq = (v) => typeof v === 'string' && (re ? re.test(v) : v.toLowerCase() === lo);
    return op === '=' ? eq : (v) => !eq(v);
  }
  return (v) => typeof v === 'string' && { '<': (c) => c < 0, '>': (c) => c > 0, '<=': (c) => c <= 0, '>=': (c) => c >= 0 }[op](compare(v, rest));
}
function ifs(args, ctx, start, cb) {
  const pairs = [];
  for (let i = start; i + 1 < args.length; i += 2) {
    if (!isRef(args[i]) && !isMat(args[i])) return ERR.VALUE;
    pairs.push([toMatrix(isRef(args[i]) ? clip(args[i], ctx) : args[i], ctx), criteria(S(args[i + 1], ctx), ctx.loc)]);
  }
  if (!pairs.length) return ERR.VALUE;
  const R = pairs[0][0].length; const C = pairs[0][0][0].length;
  for (const [m] of pairs) if (m.length !== R || m[0].length !== C) return ERR.VALUE;
  for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) if (pairs.every(([m, f]) => f(m[r][c]))) cb(r, c);
  return null;
}
const sumRangeVal = (ref, ctx, r, c) => (isRef(ref) ? ctx.get(ref.sheet, ref.r1 + r, ref.c1 + c) : isMat(ref) ? (ref[r] || [])[c] : null);
def('COUNTIF', (a, ctx) => { let n = 0; const e = ifs([a[0], a[1]], ctx, 0, () => { n++; }); return e || n; });
def('COUNTIFS', (a, ctx) => { let n = 0; const e = ifs(a, ctx, 0, () => { n++; }); return e || n; });
def('SUMIF', (a, ctx) => { let s = 0; let err = null; const tgt = a[2] && (isRef(a[2]) || isMat(a[2])) ? a[2] : a[0]; const e = ifs([a[0], a[1]], ctx, 0, (r, c) => { const v = sumRangeVal(tgt, ctx, r, c); if (isErr(v)) err = v; else if (typeof v === 'number') s += v; }); return e || err || clean(s); });
def('SUMIFS', (a, ctx) => { let s = 0; let err = null; const e = ifs(a, ctx, 1, (r, c) => { const v = sumRangeVal(a[0], ctx, r, c); if (isErr(v)) err = v; else if (typeof v === 'number') s += v; }); return e || err || clean(s); });
def('AVERAGEIF', (a, ctx) => { let s = 0; let n = 0; const tgt = a[2] && (isRef(a[2]) || isMat(a[2])) ? a[2] : a[0]; const e = ifs([a[0], a[1]], ctx, 0, (r, c) => { const v = sumRangeVal(tgt, ctx, r, c); if (typeof v === 'number') { s += v; n++; } }); return e || (n ? s / n : ERR.DIV0); });
def('AVERAGEIFS', (a, ctx) => { let s = 0; let n = 0; const e = ifs(a, ctx, 1, (r, c) => { const v = sumRangeVal(a[0], ctx, r, c); if (typeof v === 'number') { s += v; n++; } }); return e || (n ? s / n : ERR.DIV0); });
def('MAXIFS', (a, ctx) => { let m = null; const e = ifs(a, ctx, 1, (r, c) => { const v = sumRangeVal(a[0], ctx, r, c); if (typeof v === 'number' && (m == null || v > m)) m = v; }); return e || (m == null ? 0 : m); });
def('MINIFS', (a, ctx) => { let m = null; const e = ifs(a, ctx, 1, (r, c) => { const v = sumRangeVal(a[0], ctx, r, c); if (typeof v === 'number' && (m == null || v < m)) m = v; }); return e || (m == null ? 0 : m); });

/* ---------------- lookup & reference ---------------- */
function vec(v, ctx) { // ranges/arrays → { get(i), n, horizontal }
  const m = toMatrix(isRef(v) ? clip(v, ctx) : v, ctx);
  const horizontal = m.length === 1 && m[0].length > 1;
  const n = horizontal ? m[0].length : m.length;
  return { n, horizontal, m, get: (i) => (horizontal ? m[0][i] : m[i][0]) };
}
const sameType = (a, b) => (typeof a === 'number' && typeof b === 'number') || (typeof a === 'string' && typeof b === 'string') || (typeof a === 'boolean' && typeof b === 'boolean');
function exactIndex(val, get, n, wild, reverse) {
  const re = wild && typeof val === 'string' && /[*?~]/.test(val) ? wildRe(val) : null;
  const lo = typeof val === 'string' ? val.toLowerCase() : null;
  for (let k = 0; k < n; k++) {
    const i = reverse ? n - 1 - k : k;
    const x = get(i);
    if (re ? typeof x === 'string' && re.test(x) : lo != null ? typeof x === 'string' && x.toLowerCase() === lo : sameType(x, val) && compare(x, val) === 0) return i;
  }
  return -1;
}
function approxIndex(val, get, n, desc) { // binary search: largest ≤ val (asc) / smallest ≥ val (desc)
  let lo = 0; let hi = n - 1; let ans = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let x = get(mid);
    let k = mid;
    while (x == null && k < hi) x = get(++k);
    if (x == null || !sameType(x, val)) { const c0 = x == null ? 1 : compare(x, val); if (desc ? c0 < 0 : c0 > 0) hi = mid - 1; else { lo = k + 1; } continue; }
    const c = compare(x, val);
    if (desc ? c >= 0 : c <= 0) { ans = k; lo = k + 1; } else hi = mid - 1;
  }
  return ans;
}
def('VLOOKUP HLOOKUP', (a, ctx, node) => {
  const h = node.name === 'HLOOKUP';
  const val = S(a[0], ctx); if (isErr(val)) return val;
  const tbl = toMatrix(isRef(a[1]) ? clip(a[1], ctx) : a[1], ctx);
  const idx = N(a[2], ctx); if (isErr(idx)) return idx;
  const approx = a.length < 4 || a[3] == null ? true : B(a[3], ctx);
  const n = h ? tbl[0].length : tbl.length;
  const width = h ? tbl.length : tbl[0].length;
  const k = Math.trunc(idx);
  if (k < 1) return ERR.VALUE;
  if (k > width) return ERR.REF;
  const get = (i) => (h ? tbl[0][i] : tbl[i][0]);
  const i = approx ? approxIndex(val, get, n, false) : exactIndex(val, get, n, true);
  if (i < 0) return ERR.NA;
  const v = h ? tbl[k - 1][i] : tbl[i][k - 1];
  return v == null ? 0 : v;
});
def('MATCH', (a, ctx) => {
  const val = S(a[0], ctx); if (isErr(val)) return val;
  const v = vec(a[1], ctx);
  if (v.m.length > 1 && v.m[0].length > 1) return ERR.NA;
  const type = a.length > 2 && a[2] != null ? N(a[2], ctx) : 1;
  const i = type === 0 ? exactIndex(val, v.get, v.n, true) : approxIndex(val, v.get, v.n, type < 0);
  return i < 0 ? ERR.NA : i + 1;
});
function xfind(val, v, mode, search) {
  if (search === 2 || search === -2) { const i = approxIndex(val, v.get, v.n, search < 0); if (i >= 0 && compare(v.get(i), val) === 0) return i; if (mode === 0) return -1; }
  const rev = search === -1;
  let i = exactIndex(val, v.get, v.n, mode === 2, rev);
  if (i >= 0 || mode === 0 || mode === 2) return i;
  let best = -1;
  for (let k = 0; k < v.n; k++) {
    const x = v.get(k);
    if (!sameType(x, val)) continue;
    const c = compare(x, val);
    if (mode === -1 && c < 0 && (best < 0 || compare(x, v.get(best)) > 0)) best = k;
    if (mode === 1 && c > 0 && (best < 0 || compare(x, v.get(best)) < 0)) best = k;
  }
  i = best;
  return i;
}
def('XLOOKUP', (a, ctx) => {
  const val = S(a[0], ctx); if (isErr(val)) return val;
  const lv = vec(a[1], ctx);
  const mode = a[4] != null ? N(a[4], ctx) : 0; const search = a[5] != null ? N(a[5], ctx) : 1;
  const i = xfind(val, lv, mode, search);
  if (i < 0) return a.length > 3 && a[3] != null ? a[3] : ERR.NA;
  const ret = a[2];
  if (isRef(ret)) return lv.horizontal ? new Ref(ret.sheet, ret.r1, ret.c1 + i, ret.r2, ret.c1 + i) : new Ref(ret.sheet, ret.r1 + i, ret.c1, ret.r1 + i, ret.c2);
  const m = toMatrix(ret, ctx);
  return lv.horizontal ? m.map((row) => [row[i]]) : [m[i]];
});
def('XMATCH', (a, ctx) => { const val = S(a[0], ctx); if (isErr(val)) return val; const i = xfind(val, vec(a[1], ctx), a[2] != null ? N(a[2], ctx) : 0, a[3] != null ? N(a[3], ctx) : 1); return i < 0 ? ERR.NA : i + 1; });
def('LOOKUP', (a, ctx) => {
  const val = S(a[0], ctx); if (isErr(val)) return val;
  const lv = vec(a[1], ctx);
  const i = approxIndex(val, lv.get, lv.n, false);
  if (i < 0) return ERR.NA;
  if (a[2]) { const rv = vec(a[2], ctx); return i < rv.n ? rv.get(i) : ERR.NA; }
  return lv.get(i);
});
def('INDEX', (a, ctx) => {
  let src = a[0];
  if (isMat(src) && src.union) { const ar = a[3] != null ? N(a[3], ctx) : 1; src = src[ar - 1]; if (!src) return ERR.REF; }
  let r = a[1] != null ? N(a[1], ctx) : 0; let c = a[2] != null ? N(a[2], ctx) : 0;
  if (isErr(r)) return r; if (isErr(c)) return c;
  r = Math.trunc(r); c = Math.trunc(c);
  const rows = isRef(src) ? src.rows : toMatrix(src, ctx).length;
  const cols = isRef(src) ? src.cols : toMatrix(src, ctx)[0].length;
  if (a.length === 2 && rows === 1) { c = r; r = 1; }
  if (r < 0 || c < 0 || r > rows || c > cols) return ERR.REF;
  if (isRef(src)) {
    const r1 = r ? src.r1 + r - 1 : src.r1; const r2 = r ? r1 : src.r2;
    const c1 = c ? src.c1 + c - 1 : src.c1; const c2 = c ? c1 : src.c2;
    return new Ref(src.sheet, r1, c1, r2, c2);
  }
  const m = toMatrix(src, ctx);
  if (r && c) return m[r - 1][c - 1];
  if (r) return [m[r - 1]];
  return m.map((row) => [row[c - 1]]);
});
def('ROW', (a, ctx) => { if (!a.length || a[0] == null) return ctx.r + 1; if (!isRef(a[0])) return ERR.VALUE; const g = a[0]; if (g.single || g.c1 === g.c2 && g.rows === 1) return g.r1 + 1; const out = []; for (let r = g.r1; r <= clip(g, ctx).r2; r++) out.push([r + 1]); return out.length === 1 ? out[0][0] : out; });
def('COLUMN', (a, ctx) => { if (!a.length || a[0] == null) return ctx.c + 1; if (!isRef(a[0])) return ERR.VALUE; const g = a[0]; if (g.cols === 1) return g.c1 + 1; const row = []; for (let c = g.c1; c <= clip(g, ctx).c2; c++) row.push(c + 1); return [row]; });
def('ROWS', (a, ctx) => (isRef(a[0]) ? a[0].rows : isMat(a[0]) ? a[0].length : 1));
def('COLUMNS', (a, ctx) => (isRef(a[0]) ? a[0].cols : isMat(a[0]) ? a[0][0].length : 1));
def('AREAS', (a) => (isMat(a[0]) && a[0].union ? a[0].length : 1));
def('OFFSET', (a, ctx) => {
  const g = a[0]; if (!isRef(g)) return ERR.VALUE;
  const dr = N(a[1], ctx); const dc = N(a[2], ctx);
  if (isErr(dr)) return dr; if (isErr(dc)) return dc;
  const h = a[3] != null ? N(a[3], ctx) : g.rows; const w = a[4] != null ? N(a[4], ctx) : g.cols;
  if (isErr(h) || isErr(w) || h === 0 || w === 0) return ERR.REF;
  const r1 = g.r1 + Math.trunc(dr); const c1 = g.c1 + Math.trunc(dc);
  const r2 = r1 + Math.trunc(Math.abs(h)) - 1; const c2 = c1 + Math.trunc(Math.abs(w)) - 1;
  if (r1 < 0 || c1 < 0 || r2 >= MAXR || c2 >= MAXC) return ERR.REF;
  return new Ref(g.sheet, r1, c1, r2, c2);
});
def('INDIRECT', (a, ctx) => {
  const s = T(a[0], ctx); if (isErr(s)) return s;
  const a1 = a.length < 2 || a[1] == null ? true : B(a[1], ctx);
  let text = s.trim();
  let sheet = ctx.sheet;
  if (!a1) {
    const m = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(text);
    if (!m) return ERR.REF;
    if (m[1] || m[2]) { sheet = ctx.sheetByName((m[1] || m[2]).replace(/''/g, "'")); if (!sheet) return ERR.REF; }
    return new Ref(sheet, +m[3] - 1, +m[4] - 1, (m[5] ? +m[5] : +m[3]) - 1, (m[6] ? +m[6] : +m[4]) - 1);
  }
  const g = parseRange(text);
  if (!g) {
    const nm = ctx.name(text, ctx.sheet);
    if (nm) { const v = evaluate(nm.ast, nm.sheet ? { ...ctx, sheet: nm.sheet } : ctx); return isRef(v) ? v : ERR.REF; }
    return ERR.REF;
  }
  if (g.sheet) { sheet = ctx.sheetByName(g.sheet); if (!sheet) return ERR.REF; }
  return new Ref(sheet, g.r1, g.c1, g.r2, g.c2);
});
def('ADDRESS', (a, ctx) => {
  const r = N(a[0], ctx); const c = N(a[1], ctx); if (isErr(r)) return r; if (isErr(c)) return c;
  const abs = a[2] != null ? N(a[2], ctx) : 1; const a1 = a[3] == null ? true : B(a[3], ctx);
  const sh = a[4] != null ? T(a[4], ctx) : '';
  const pre = sh ? `${quoteSheet(sh)}!` : '';
  if (!a1) return `${pre}R${abs === 1 || abs === 2 ? r : `[${r}]`}C${abs === 1 || abs === 3 ? c : `[${c}]`}`;
  return `${pre}${abs === 1 || abs === 3 ? '$' : ''}${colName(c - 1)}${abs === 1 || abs === 2 ? '$' : ''}${r}`;
});
def('TRANSPOSE', (a, ctx) => { const m = toMatrix(a[0], ctx); return m[0].map((_, j) => m.map((row) => (row[j] == null ? 0 : row[j]))); });
def('HYPERLINK', (a, ctx) => { const url = T(a[0], ctx); if (ctx.onLink && typeof url === 'string') ctx.onLink(url); return a.length > 1 ? S(a[1], ctx) : url; });
def('ISFORMULA', (a, ctx) => (isRef(a[0]) ? !!ctx.formulaAt(a[0].sheet, a[0].r1, a[0].c1) : ERR.VALUE));
def('FORMULATEXT', (a, ctx) => { if (!isRef(a[0])) return ERR.VALUE; const f = ctx.formulaAt(a[0].sheet, a[0].r1, a[0].c1); return f ? `=${f}` : ERR.NA; });
lazy('ISREF', (n, ctx) => isRef(evaluate(n[0], ctx)));
def('SHEET', (a, ctx) => (a.length && isRef(a[0]) ? ctx.sheetIndex(a[0].sheet) + 1 : a.length ? (ctx.sheetByName(T(a[0], ctx)) ? ctx.sheetIndex(ctx.sheetByName(T(a[0], ctx))) + 1 : ERR.NA) : ctx.sheetIndex(ctx.sheet) + 1));
def('SHEETS', (a, ctx) => ctx.sheetCount());

/* ---------------- SUBTOTAL / AGGREGATE ---------------- */
const AGG = {
  1: (xs) => (xs.length ? xs.reduce((s, x) => s + x, 0) / xs.length : ERR.DIV0), 2: (xs) => xs.length, 3: null,
  4: (xs) => (xs.length ? Math.max(...xs) : 0), 5: (xs) => (xs.length ? Math.min(...xs) : 0), 6: (xs) => xs.reduce((s, x) => s * x, 1),
  7: (xs) => stat(xs, 'ss'), 8: (xs) => stat(xs, 'sp'), 9: (xs) => clean(xs.reduce((s, x) => s + x, 0)), 10: (xs) => stat(xs, 'vs'), 11: (xs) => stat(xs, 'vp'),
  12: (xs) => { if (!xs.length) return ERR.NUM; xs.sort((x, y) => x - y); const m = xs.length >> 1; return xs.length % 2 ? xs[m] : (xs[m - 1] + xs[m]) / 2; },
  13: (xs) => { const cnt = new Map(); let best = null; let bc = 1; for (const x of xs) { const c = (cnt.get(x) || 0) + 1; cnt.set(x, c); if (c > bc) { bc = c; best = x; } } return best == null ? ERR.NA : best; },
};
function subValues(refs, ctx, { hidden, errors, nested }) {
  const xs = []; let count3 = 0; let err = null;
  for (const a of refs) {
    if (!isRef(a)) { const m = toMatrix(a, ctx); for (const row of m) for (const v of row) { if (isErr(v)) { if (!errors) err = err || v; } else if (typeof v === 'number') xs.push(v); if (v != null) count3++; } continue; }
    const g = clip(a, ctx);
    for (let r = g.r1; r <= g.r2; r++) {
      const hid = ctx.rowHidden ? ctx.rowHidden(g.sheet, r) : false;
      if (hid === 'filter' || (hid && hidden)) continue;
      for (let c = g.c1; c <= g.c2; c++) {
        if (nested) { const f = ctx.formulaAt(g.sheet, r, c); if (f && /\b(SUBTOTAL|AGGREGATE)\s*\(/i.test(f)) continue; }
        const v = ctx.get(g.sheet, r, c);
        if (isErr(v)) { if (!errors) err = err || v; continue; }
        if (v != null) count3++;
        if (typeof v === 'number') xs.push(v);
      }
    }
  }
  return { xs, count3, err };
}
def('SUBTOTAL', (a, ctx) => {
  const f = N(a[0], ctx); if (isErr(f)) return f;
  const k = Math.trunc(f) % 100;
  if (!AGG[k] && k !== 3) return ERR.VALUE;
  const { xs, count3, err } = subValues(a.slice(1), ctx, { hidden: f >= 100, errors: false, nested: true });
  if (err) return err;
  return k === 3 ? count3 : AGG[k](xs);
});
def('AGGREGATE', (a, ctx) => {
  const f = N(a[0], ctx); const o = N(a[1], ctx); if (isErr(f)) return f; if (isErr(o)) return o;
  const opt = Math.trunc(o);
  const flags = { nested: opt <= 3, hidden: [1, 3, 5, 7].includes(opt), errors: [2, 3, 6, 7].includes(opt) };
  if (f >= 14 && f <= 19) {
    const { xs, err } = subValues([a[2]], ctx, flags);
    if (err) return err;
    const k = N(a[3], ctx); if (isErr(k)) return k;
    xs.sort((x, y) => x - y);
    if (f === 14) return k < 1 || k > xs.length ? ERR.NUM : xs[xs.length - Math.ceil(k)];
    if (f === 15) return k < 1 || k > xs.length ? ERR.NUM : xs[Math.ceil(k) - 1];
    const p = f === 17 || f === 19 ? k / 4 : k;
    return percentile(xs, p, f === 18 || f === 19);
  }
  const { xs, count3, err } = subValues(a.slice(2), ctx, flags);
  if (err) return err;
  return f === 3 ? count3 : AGG[f] ? AGG[f](xs) : ERR.VALUE;
});
function percentile(xs, p, exc) {
  const n = xs.length;
  if (!n || p < 0 || p > 1) return ERR.NUM;
  const h = exc ? (n + 1) * p - 1 : (n - 1) * p;
  if (h < 0 || h > n - 1) return ERR.NUM;
  const lo = Math.floor(h);
  return xs[lo] + (h - lo) * ((xs[lo + 1] ?? xs[lo]) - xs[lo]);
}
def('PERCENTILE PERCENTILE.INC', (a, c) => { const xs = nums([a[0]], c); const p = N(a[1], c); if (isErr(xs)) return xs; if (isErr(p)) return p; return percentile(xs.sort((x, y) => x - y), p, false); });
def('PERCENTILE.EXC', (a, c) => { const xs = nums([a[0]], c); const p = N(a[1], c); if (isErr(xs)) return xs; if (isErr(p)) return p; return percentile(xs.sort((x, y) => x - y), p, true); });
def('QUARTILE QUARTILE.INC', (a, c) => { const xs = nums([a[0]], c); const q = N(a[1], c); if (isErr(xs)) return xs; if (isErr(q)) return q; return percentile(xs.sort((x, y) => x - y), Math.trunc(q) / 4, false); });

/* ---------------- dynamic arrays ---------------- */
const rowsOf = (m) => m;
def('SEQUENCE', (a, ctx) => {
  const R = a[0] != null ? N(a[0], ctx) : 1; const C = a[1] != null ? N(a[1], ctx) : 1;
  const st = a[2] != null ? N(a[2], ctx) : 1; const sp = a[3] != null ? N(a[3], ctx) : 1;
  for (const x of [R, C, st, sp]) if (isErr(x)) return x;
  if (R < 1 || C < 1) return ERR.CALC;
  if (R * C > 1e6) return ERR.NUM;
  const out = [];
  for (let r = 0; r < R; r++) { const row = []; for (let c = 0; c < C; c++) row.push(clean(st + (r * C + c) * sp)); out.push(row); }
  return out;
});
def('RANDARRAY', (a, ctx) => {
  const R = a[0] != null ? N(a[0], ctx) : 1; const C = a[1] != null ? N(a[1], ctx) : 1;
  const lo = a[2] != null ? N(a[2], ctx) : 0; const hi = a[3] != null ? N(a[3], ctx) : 1; const int = a[4] != null ? B(a[4], ctx) : false;
  const out = [];
  for (let r = 0; r < R; r++) { const row = []; for (let c = 0; c < C; c++) row.push(int ? Math.floor(lo + Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo)); out.push(row); }
  return out;
});
const keyOf = (v) => (typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${v}`);
def('UNIQUE', (a, ctx) => {
  const m = toMatrix(a[0], ctx); const byCol = a[1] != null && B(a[1], ctx) === true; const once = a[2] != null && B(a[2], ctx) === true;
  const items = byCol ? m[0].map((_, j) => m.map((row) => row[j])) : rowsOf(m);
  const cnt = new Map(); const first = new Map();
  items.forEach((it, i) => { const k = it.map(keyOf).join('\u0000'); cnt.set(k, (cnt.get(k) || 0) + 1); if (!first.has(k)) first.set(k, i); });
  const keep = [...first.entries()].filter(([k]) => !once || cnt.get(k) === 1).map(([, i]) => items[i].map((x) => (x == null ? 0 : x)));
  if (!keep.length) return ERR.CALC;
  return byCol ? keep[0].map((_, i) => keep.map((col) => col[i])) : keep;
});
function sortRows(rows, keys) {
  return rows.map((r, i) => [r, i]).sort((x, y) => {
    for (const { get, ord } of keys) { const c = cmpSort(get(x[1]), get(y[1])); if (c) return c * ord; }
    return x[1] - y[1];
  }).map((p) => p[0]);
}
export function cmpSort(a, b) { // blanks last, errors after booleans
  const blankA = a == null || a === ''; const blankB = b == null || b === '';
  if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 1 : -1;
  if (isErr(a) || isErr(b)) return isErr(a) === isErr(b) ? 0 : isErr(a) ? 1 : -1;
  return compare(a, b);
}
def('SORT', (a, ctx) => {
  const m = toMatrix(a[0], ctx);
  const idx = a[1] != null ? N(a[1], ctx) : 1; const ord = a[2] != null ? N(a[2], ctx) : 1; const byCol = a[3] != null && B(a[3], ctx) === true;
  if (isErr(idx) || isErr(ord)) return ERR.VALUE;
  const items = byCol ? m[0].map((_, j) => m.map((row) => row[j])) : m;
  const s = sortRows(items, [{ get: (i) => items[i][idx - 1], ord: ord < 0 ? -1 : 1 }]);
  return byCol ? s[0].map((_, i) => s.map((col) => col[i])) : s;
});
def('SORTBY', (a, ctx) => {
  const m = toMatrix(a[0], ctx);
  const keys = [];
  for (let i = 1; i < a.length; i += 2) { const v = vec(a[i], ctx); const ord = a[i + 1] != null ? N(a[i + 1], ctx) : 1; keys.push({ get: (k) => v.get(k), ord: ord < 0 ? -1 : 1 }); }
  return sortRows(m, keys);
});
def('FILTER', (a, ctx) => {
  const m = toMatrix(a[0], ctx); const inc = toMatrix(a[1], ctx);
  const byRow = inc.length === m.length && inc[0].length === 1;
  const byCol = !byRow && inc.length === 1 && inc[0].length === m[0].length;
  if (!byRow && !byCol) return ERR.VALUE;
  const ok = (v) => { const b = toBool(v); return !isErr(b) && b; };
  const out = byRow ? m.filter((_, i) => ok(inc[i][0])) : m.map((row) => row.filter((_, j) => ok(inc[0][j])));
  if (!out.length || !out[0].length) return a.length > 2 ? S(a[2], ctx) : ERR.CALC;
  return out;
});
def('TOCOL', (a, ctx) => { const m = toMatrix(a[0], ctx); const ign = a[1] != null ? N(a[1], ctx) : 0; const out = []; for (const row of m) for (const v of row) { if ((ign === 1 || ign === 3) && v == null) continue; if ((ign === 2 || ign === 3) && isErr(v)) continue; out.push([v == null ? 0 : v]); } return out.length ? out : ERR.CALC; });
def('TOROW', (a, ctx) => { const m = toMatrix(a[0], ctx); const out = []; for (const row of m) for (const v of row) out.push(v == null ? 0 : v); return [out]; });

/* ---------------- information ---------------- */
const lift1 = (names, f) => def(names, lifted((a, ctx) => f(a[0], ctx, a)));
lift1('ISBLANK', (v) => v == null);
lift1('ISERROR', (v) => isErr(v));
lift1('ISERR', (v) => isErr(v) && v.code !== '#N/A');
lift1('ISNA', (v) => isErr(v) && v.code === '#N/A');
lift1('ISNUMBER', (v) => typeof v === 'number');
lift1('ISTEXT', (v) => typeof v === 'string');
lift1('ISNONTEXT', (v) => typeof v !== 'string');
lift1('ISLOGICAL', (v) => typeof v === 'boolean');
lift1('ISEVEN', (v, ctx) => { const n = toNum(v, ctx.loc); return isErr(n) ? n : Math.trunc(n) % 2 === 0; });
lift1('ISODD', (v, ctx) => { const n = toNum(v, ctx.loc); return isErr(n) ? n : Math.abs(Math.trunc(n) % 2) === 1; });
lift1('N', (v) => (typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : isErr(v) ? v : 0));
lift1('T', (v) => (typeof v === 'string' ? v : isErr(v) ? v : ''));
def('NA', () => ERR.NA);
def('ERROR.TYPE', (a, ctx) => { const v = S(a[0], ctx); return isErr(v) ? ERROR_TYPE[v.code] || ERR.NA : ERR.NA; });
def('TYPE', (a, ctx) => { const v = a[0]; if (isMulti(v)) return 64; const s = S(v, ctx); return typeof s === 'number' || s == null ? 1 : typeof s === 'string' ? 2 : typeof s === 'boolean' ? 4 : isErr(s) ? 16 : 1; });

/* ---------------- text ---------------- */
const lt = (names, f) => def(names, lifted((a, ctx) => { for (const x of a) if (isErr(x)) return x; return f(a, ctx); }));
lt('TEXT', (a, ctx) => { const v = a[0]; const f = toStr(a[1], ctx.loc); return formatValue(typeof v === 'string' ? (textToNumber(v, ctx.loc) ?? v) : v, f, ctx.loc).text; });
lt('LEN', (a, ctx) => toStr(a[0], ctx.loc).length);
lt('LEFT', (a, ctx) => { const n = a.length > 1 && a[1] != null ? toNum(a[1], ctx.loc) : 1; return n < 0 ? ERR.VALUE : toStr(a[0], ctx.loc).slice(0, n); });
lt('RIGHT', (a, ctx) => { const n = a.length > 1 && a[1] != null ? toNum(a[1], ctx.loc) : 1; const s = toStr(a[0], ctx.loc); return n < 0 ? ERR.VALUE : n === 0 ? '' : s.slice(-n); });
lt('MID', (a, ctx) => { const s = toStr(a[0], ctx.loc); const st = toNum(a[1], ctx.loc); const n = toNum(a[2], ctx.loc); if (isErr(st) || isErr(n)) return ERR.VALUE; return st < 1 || n < 0 ? ERR.VALUE : s.substr(st - 1, n); });
lt('UPPER', (a, ctx) => toStr(a[0], ctx.loc).toUpperCase());
lt('LOWER', (a, ctx) => toStr(a[0], ctx.loc).toLowerCase());
lt('PROPER', (a, ctx) => toStr(a[0], ctx.loc).toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (m, p, ch) => p + ch.toUpperCase()));
lt('TRIM', (a, ctx) => toStr(a[0], ctx.loc).replace(/ +/g, ' ').trim());
lt('CLEAN', (a, ctx) => toStr(a[0], ctx.loc).replace(/[\x00-\x1f]/g, ''));
lt('REPT', (a, ctx) => { const n = toNum(a[1], ctx.loc); return n < 0 || n > 32767 ? ERR.VALUE : toStr(a[0], ctx.loc).repeat(Math.trunc(n)); });
lt('EXACT', (a, ctx) => toStr(a[0], ctx.loc) === toStr(a[1], ctx.loc));
lt('SUBSTITUTE', (a, ctx) => {
  const s = toStr(a[0], ctx.loc); const o = toStr(a[1], ctx.loc); const n = toStr(a[2], ctx.loc);
  if (!o) return s;
  if (a.length > 3 && a[3] != null) { const k = toNum(a[3], ctx.loc); let i = -1; for (let j = 0; j < k; j++) { i = s.indexOf(o, i + 1); if (i < 0) return s; } return s.slice(0, i) + n + s.slice(i + o.length); }
  return s.split(o).join(n);
});
lt('REPLACE', (a, ctx) => { const s = toStr(a[0], ctx.loc); const st = toNum(a[1], ctx.loc); const n = toNum(a[2], ctx.loc); return s.slice(0, st - 1) + toStr(a[3], ctx.loc) + s.slice(st - 1 + n); });
lt('FIND', (a, ctx) => { const st = a.length > 2 && a[2] != null ? toNum(a[2], ctx.loc) : 1; const i = toStr(a[1], ctx.loc).indexOf(toStr(a[0], ctx.loc), st - 1); return i < 0 || st < 1 ? ERR.VALUE : i + 1; });
lt('SEARCH', (a, ctx) => { const st = a.length > 2 && a[2] != null ? toNum(a[2], ctx.loc) : 1; const hay = toStr(a[1], ctx.loc); const nd = toStr(a[0], ctx.loc); const re = wildRe(nd).source.slice(1, -1); const m = new RegExp(re, 'i').exec(hay.slice(st - 1)); return m ? m.index + st : ERR.VALUE; });
lt('VALUE', (a, ctx) => { const v = a[0]; if (typeof v === 'number') return v; const p = parseInput(toStr(v, ctx.loc), ctx.loc); return typeof p.v === 'number' ? p.v : ERR.VALUE; });
lt('NUMBERVALUE', (a, ctx) => { const dec = a[1] != null ? toStr(a[1], ctx.loc) : ctx.loc.decimal; const grp = a[2] != null ? toStr(a[2], ctx.loc) : ctx.loc.group; let s = toStr(a[0], ctx.loc).replace(/\s/g, '').split(grp).join(''); s = s.replace(dec, '.'); const pct = (s.match(/%/g) || []).length; s = s.replace(/%/g, ''); const n = Number(s); return s === '' ? 0 : Number.isFinite(n) ? n / 100 ** pct : ERR.VALUE; });
lt('DATEVALUE', (a, ctx) => { const p = parseInput(toStr(a[0], ctx.loc), ctx.loc); return typeof p.v === 'number' && p.fmt ? Math.floor(p.v) : ERR.VALUE; });
lt('TIMEVALUE', (a, ctx) => { const p = parseInput(toStr(a[0], ctx.loc), ctx.loc); return typeof p.v === 'number' && p.fmt ? p.v - Math.floor(p.v) : ERR.VALUE; });
lt('CHAR', (a, ctx) => { const n = toNum(a[0], ctx.loc); return n < 1 || n > 255 ? ERR.VALUE : String.fromCharCode(n); });
lt('CODE', (a, ctx) => { const s = toStr(a[0], ctx.loc); return s ? s.charCodeAt(0) : ERR.VALUE; });
lt('UNICHAR', (a, ctx) => String.fromCodePoint(toNum(a[0], ctx.loc)));
lt('UNICODE', (a, ctx) => { const s = toStr(a[0], ctx.loc); return s ? s.codePointAt(0) : ERR.VALUE; });
lt('FIXED', (a, ctx) => { const d = a[1] != null ? toNum(a[1], ctx.loc) : 2; const no = a[2] != null && toBool(a[2]) === true; return formatValue(toNum(a[0], ctx.loc), `${no ? '0' : '#,##0'}${d > 0 ? `.${'0'.repeat(d)}` : ''}`, ctx.loc).text; });
lt('DOLLAR', (a, ctx) => { const d = a[1] != null ? toNum(a[1], ctx.loc) : 2; return formatValue(toNum(a[0], ctx.loc), ctx.loc.currencyFmt.replace(/0\.00|0,00/, d > 0 ? `0.${'0'.repeat(d)}` : '0'), ctx.loc).text; });
lt('TEXTBEFORE', (a, ctx) => { const s = toStr(a[0], ctx.loc); const d = toStr(a[1], ctx.loc); const n = a[2] != null ? toNum(a[2], ctx.loc) : 1; const i = n > 0 ? nthIndex(s, d, n) : nthLast(s, d, -n); return i < 0 ? (a[5] != null ? a[5] : ERR.NA) : s.slice(0, i); });
lt('TEXTAFTER', (a, ctx) => { const s = toStr(a[0], ctx.loc); const d = toStr(a[1], ctx.loc); const n = a[2] != null ? toNum(a[2], ctx.loc) : 1; const i = n > 0 ? nthIndex(s, d, n) : nthLast(s, d, -n); return i < 0 ? (a[5] != null ? a[5] : ERR.NA) : s.slice(i + d.length); });
function nthIndex(s, d, n) { let i = -1; for (let k = 0; k < n; k++) { i = s.indexOf(d, i + 1); if (i < 0) return -1; } return i; }
function nthLast(s, d, n) { let i = s.length; for (let k = 0; k < n; k++) { i = s.lastIndexOf(d, i - 1); if (i < 0) return -1; } return i; }
def('TEXTSPLIT', (a, ctx) => {
  const s = T(a[0], ctx); const cd = a[1] != null ? T(a[1], ctx) : null; const rd = a[2] != null ? T(a[2], ctx) : null;
  const rows = rd ? s.split(rd) : [s];
  const m = rows.map((r) => (cd ? r.split(cd) : [r]));
  const w = Math.max(...m.map((r) => r.length));
  return m.map((r) => { while (r.length < w) r.push(ERR.NA); return r; });
});
def('CONCAT CONCATENATE', (a, ctx) => { let s = ''; let err = null; each(a, ctx, (v) => { if (isErr(v)) { err = v; return false; } s += toStr(v, ctx.loc); return undefined; }); return err || s; });
def('TEXTJOIN', (a, ctx) => { const d = T(a[0], ctx); const ign = B(a[1], ctx); const parts = []; let err = null; each(a.slice(2), ctx, (v) => { if (isErr(v)) { err = v; return false; } const s = toStr(v, ctx.loc); if (!(ign && s === '')) parts.push(s); return undefined; }); return err || parts.join(d); });

/* ---------------- date & time ---------------- */
const P = (n) => partsFromSerial(n);
lt('DATE', (a, ctx) => { const y = toNum(a[0], ctx.loc); const m = toNum(a[1], ctx.loc); const d = toNum(a[2], ctx.loc); for (const x of [y, m, d]) if (isErr(x)) return x; const s = serialFromParts(Math.trunc(y), Math.trunc(m), Math.trunc(d)); return s < 0 ? ERR.NUM : s; });
lt('TIME', (a, ctx) => { const h = toNum(a[0], ctx.loc); const m = toNum(a[1], ctx.loc); const s = toNum(a[2], ctx.loc); for (const x of [h, m, s]) if (isErr(x)) return x; const t = (Math.trunc(h) * 3600 + Math.trunc(m) * 60 + Math.trunc(s)) / 86400; return t < 0 ? ERR.NUM : t - Math.floor(t); });
const dpart = (names, f) => lt(names, (a, ctx) => { const n = toNum(a[0], ctx.loc); if (isErr(n)) return n; if (n < 0) return ERR.NUM; return f(P(n), n, a, ctx); });
dpart('YEAR', (p) => p.y); dpart('MONTH', (p) => p.m); dpart('DAY', (p) => p.d);
dpart('HOUR', (p) => p.h); dpart('MINUTE', (p) => p.mi); dpart('SECOND', (p) => p.s + (p.ms >= 500 ? 1 : 0));
dpart('WEEKDAY', (p, n, a, ctx) => {
  const t = a[1] != null ? toNum(a[1], ctx.loc) : 1; const d = p.dow; // 0 = Sunday
  if (t === 1 || t === 17) return d + 1;
  if (t === 2 || t === 11) return ((d + 6) % 7) + 1;
  if (t === 3) return (d + 6) % 7;
  if (t >= 12 && t <= 16) return ((d - (t - 10) + 7) % 7) + 1;
  return ERR.NUM;
});
const addMonths = (n, k, eom) => { const p = P(n); let y = p.y; let m = p.m + k; y += Math.floor((m - 1) / 12); m = ((((m - 1) % 12) + 12) % 12) + 1; const dim = daysInMonth(y, m); return serialFromParts(y, m, eom ? dim : Math.min(p.d, dim)); };
lt('EDATE', (a, ctx) => { const n = toNum(a[0], ctx.loc); const k = toNum(a[1], ctx.loc); if (isErr(n)) return n; if (isErr(k)) return k; return addMonths(Math.floor(n), Math.trunc(k), false); });
lt('EOMONTH', (a, ctx) => { const n = toNum(a[0], ctx.loc); const k = toNum(a[1], ctx.loc); if (isErr(n)) return n; if (isErr(k)) return k; return addMonths(Math.floor(n), Math.trunc(k), true); });
def('TODAY', () => todaySerial());
def('NOW', () => nowSerial());
lt('DAYS', (a, ctx) => { const e = toNum(a[0], ctx.loc); const s = toNum(a[1], ctx.loc); if (isErr(e)) return e; if (isErr(s)) return s; return Math.floor(e) - Math.floor(s); });
lt('DATEDIF', (a, ctx) => {
  const s = toNum(a[0], ctx.loc); const e = toNum(a[1], ctx.loc); const u = toStr(a[2], ctx.loc).toUpperCase();
  if (isErr(s) || isErr(e)) return ERR.VALUE;
  if (s > e) return ERR.NUM;
  const A = P(Math.floor(s)); const Z = P(Math.floor(e));
  let months = (Z.y - A.y) * 12 + (Z.m - A.m); if (Z.d < A.d) months--;
  switch (u) {
    case 'D': return Math.floor(e) - Math.floor(s);
    case 'M': return months;
    case 'Y': return Math.floor(months / 12);
    case 'YM': return months % 12;
    case 'MD': { let d = Z.d - A.d; if (d < 0) { const pm = Z.m === 1 ? 12 : Z.m - 1; d += daysInMonth(Z.m === 1 ? Z.y - 1 : Z.y, pm); } return d; }
    case 'YD': { let ys = serialFromParts(Z.y, A.m, A.d); if (ys > Math.floor(e)) ys = serialFromParts(Z.y - 1, A.m, A.d); return Math.floor(e) - ys; }
    default: return ERR.NUM;
  }
});
function holidaysOf(v, ctx) { const set = new Set(); if (v == null) return set; each([v], ctx, (x) => { if (typeof x === 'number') set.add(Math.floor(x)); }); return set; }
const isWeekend = (n, mask) => mask[(P(n).dow + 6) % 7] === '1'; // mask Monday..Sunday
function weekendMask(w, ctx) {
  if (w == null) return '0000011';
  const v = S(w, ctx);
  if (typeof v === 'string' && /^[01]{7}$/.test(v)) return v;
  const k = toNum(v, ctx.loc);
  const map = { 1: '0000011', 2: '1000001', 3: '1100000', 4: '0110000', 5: '0011000', 6: '0001100', 7: '0000110', 11: '0000001', 12: '1000000', 13: '0100000', 14: '0010000', 15: '0001000', 16: '0000100', 17: '0000010' };
  return map[k] || '0000011';
}
function networkdays(s, e, mask, hol) { let n = 0; const st = Math.min(s, e); const en = Math.max(s, e); for (let d = st; d <= en; d++) if (!isWeekend(d, mask) && !hol.has(d)) n++; return s > e ? -n : n; }
def('NETWORKDAYS', (a, ctx) => { const s = N(a[0], ctx); const e = N(a[1], ctx); if (isErr(s)) return s; if (isErr(e)) return e; return networkdays(Math.floor(s), Math.floor(e), '0000011', holidaysOf(a[2], ctx)); });
def('NETWORKDAYS.INTL', (a, ctx) => { const s = N(a[0], ctx); const e = N(a[1], ctx); if (isErr(s)) return s; if (isErr(e)) return e; return networkdays(Math.floor(s), Math.floor(e), weekendMask(a[2], ctx), holidaysOf(a[3], ctx)); });
function workday(s, k, mask, hol) { let d = s; const step = k < 0 ? -1 : 1; let n = Math.abs(k); while (n > 0) { d += step; if (!isWeekend(d, mask) && !hol.has(d)) n--; } return d; }
def('WORKDAY', (a, ctx) => { const s = N(a[0], ctx); const k = N(a[1], ctx); if (isErr(s)) return s; if (isErr(k)) return k; return workday(Math.floor(s), Math.trunc(k), '0000011', holidaysOf(a[2], ctx)); });
def('WORKDAY.INTL', (a, ctx) => { const s = N(a[0], ctx); const k = N(a[1], ctx); if (isErr(s)) return s; if (isErr(k)) return k; return workday(Math.floor(s), Math.trunc(k), weekendMask(a[2], ctx), holidaysOf(a[3], ctx)); });
lt('WEEKNUM', (a, ctx) => { const n = Math.floor(toNum(a[0], ctx.loc)); const t = a[1] != null ? toNum(a[1], ctx.loc) : 1; if (t === 21) return isoWeek(n); const p = P(n); const jan1 = serialFromParts(p.y, 1, 1); const start = t === 2 || t === 11 ? 1 : t >= 12 && t <= 17 ? t - 10 : 0; const off = (P(jan1).dow - start + 7) % 7; return Math.floor((n - jan1 + off) / 7) + 1; });
function isoWeek(n) { const p = P(n); const dow = (p.dow + 6) % 7; const thu = n - dow + 3; const y = P(thu).y; const jan4 = serialFromParts(y, 1, 4); return 1 + Math.round((thu - jan4 - 3 + ((P(jan4).dow + 6) % 7)) / 7); }
lt('ISOWEEKNUM', (a, ctx) => isoWeek(Math.floor(toNum(a[0], ctx.loc))));

/* ---------------- everything else from formulajs ---------------- */
const NUMLIST = new Set(['AVEDEV', 'DEVSQ', 'GEOMEAN', 'HARMEAN', 'KURT', 'SKEW', 'SKEW.P', 'MODE', 'MODE.SNGL', 'MODE.MULT', 'TRIMMEAN', 'GCD', 'LCM', 'MULTINOMIAL', 'STDEVA', 'STDEVPA', 'VARA', 'VARPA', 'PERCENTRANK', 'PERCENTRANK.INC', 'PERCENTRANK.EXC', 'RANK.AVG', 'QUARTILE.EXC', 'IRR', 'FVSCHEDULE']);
const LISTARG = { TRIMMEAN: [0], PERCENTRANK: [0], 'PERCENTRANK.INC': [0], 'PERCENTRANK.EXC': [0], 'RANK.AVG': [1], 'QUARTILE.EXC': [0], IRR: [0], FVSCHEDULE: [1], NPV: 'rest1' };
const ARRAYARG = new Set(['MMULT', 'MINVERSE', 'MDETERM', 'FREQUENCY', 'LINEST', 'LOGEST', 'TREND', 'GROWTH', 'CORREL', 'PEARSON', 'RSQ', 'SLOPE', 'INTERCEPT', 'STEYX', 'COVAR', 'COVARIANCE.P', 'COVARIANCE.S', 'FORECAST', 'FORECAST.LINEAR', 'TTEST', 'T.TEST', 'CHITEST', 'CHISQ.TEST', 'FTEST', 'F.TEST', 'ZTEST', 'Z.TEST', 'PROB', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'XNPV', 'XIRR', 'MIRR', 'VSTACK', 'HSTACK', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'EXPAND', 'SERIESSUM', 'DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DGET', 'DMAX', 'DMIN', 'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DSUM', 'DVAR', 'DVARP', 'DAYS360', 'YEARFRAC']);
function fjResult(r) {
  if (r instanceof Error) return errOf(r.message || '#VALUE!');
  if (r instanceof Date) return dateToSerial(r);
  if (typeof r === 'number') return Number.isFinite(r) ? r : ERR.NUM;
  if (Array.isArray(r)) { const m = Array.isArray(r[0]) ? r : r.map((x) => [x]); return m.map((row) => row.map((x) => { const y = fjResult(x); return isMat(y) ? ERR.VALUE : y; })); }
  if (r === undefined) return ERR.VALUE;
  return r;
}
const fjErr = (v) => (isErr(v) ? Object.assign(new Error(v.code), { message: v.code }) : v);
function onlyNums(v, ctx) { const xs = nums([v], ctx); return isErr(xs) ? xs : xs; }
function wrapFJ(name, f) {
  const call = (args) => { try { return fjResult(f(...args.map(fjErr))); } catch { return ERR.VALUE; } };
  if (NUMLIST.has(name) && !LISTARG[name]) return (args, ctx) => { const xs = nums(args, ctx); return isErr(xs) ? xs : call([xs]); };
  if (LISTARG[name]) {
    const spec = LISTARG[name];
    return (args, ctx) => {
      const a2 = [];
      for (let i = 0; i < args.length; i++) {
        const list = spec === 'rest1' ? i >= 1 : spec.includes(i);
        const v = list ? onlyNums(args[i], ctx) : S(args[i], ctx);
        if (isErr(v)) return v;
        if (spec === 'rest1' && i >= 1) a2.push(...v); else a2.push(v);
      }
      return call(spec === 'rest1' ? [a2[0], a2.slice(1)] : a2);
    };
  }
  if (ARRAYARG.has(name)) return (args, ctx) => call(args.map((a) => (isRef(a) ? (a.single ? ctx.get(a.sheet, a.r1, a.c1) : refValues(a, ctx)) : a)));
  return lifted((args) => { const e = firstErr(args); return e || call(args); });
}
(function addFormulaJs() {
  const walk = (o, p) => {
    for (const k of Object.keys(o)) {
      if (k === 'utils') continue;
      const v = o[k]; const n = p ? `${p}.${k}` : k;
      if (typeof v === 'function') { if (!FN[n]) FN[n] = { fn: wrapFJ(n, v) }; for (const s of Object.keys(v)) if (v[s] && (typeof v[s] === 'function' || typeof v[s] === 'object')) walk({ [s]: v[s] }, n); } else if (v && typeof v === 'object') walk(v, n);
    }
  };
  walk(FJ, '');
}());

export const FUNCTIONS = FN;
export const FUNCTION_NAMES = Object.keys(FN).sort();
