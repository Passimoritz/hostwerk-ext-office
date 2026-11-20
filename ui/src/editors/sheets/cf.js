/*
 * Conditional formatting evaluation for rendering: cfAt(book, sheet) → (r, c) → { fill, color, b, i, u, st, bar, scale } | null.
 * Range statistics (top/bottom, averages, duplicates, colour scales, data bars) are computed once per book version.
 */
import { shiftFormula } from './formula/parser.js';
import { isErr } from './formula/errors.js';
import { compare, toBool } from './formula/eval.js';
import { MAXC } from './formula/refs.js';

const inR = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;

function cellsOf(book, sheet, ranges, cb) {
  for (const g of ranges) {
    const b = book.bounds(sheet);
    const r2 = Math.min(g.r2, b.rows - 1); const c2 = Math.min(g.c2, b.cols - 1);
    for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) cb(book.value(sheet, r, c), r, c);
  }
}
function hex(c) { const m = /^#?([0-9a-f]{6})$/i.exec(c || ''); return m ? [parseInt(m[1].slice(0, 2), 16), parseInt(m[1].slice(2, 4), 16), parseInt(m[1].slice(4), 16)] : [255, 255, 255]; }
const mix = (a, b, t) => `rgb(${a.map((x, i) => Math.round(x + (b[i] - x) * t)).join(',')})`;

function stopValue(stop, nums, min, max) {
  const t = stop && stop.type;
  if (t === 'num') return Number(stop.val);
  if (t === 'percent') return min + ((max - min) * Number(stop.val)) / 100;
  if (t === 'percentile') { const s = nums.slice().sort((a, b) => a - b); const h = (s.length - 1) * (Number(stop.val) / 100); const lo = Math.floor(h); return s.length ? s[lo] + (h - lo) * ((s[lo + 1] ?? s[lo]) - s[lo]) : 0; }
  if (t === 'max') return max;
  return min;
}

function prepare(book, sheet, rule) {
  const p = { rule };
  if (['dup', 'unique', 'top', 'bottom', 'above', 'below', 'scale', 'bar', 'icons'].includes(rule.type)) {
    const nums = []; const counts = new Map();
    cellsOf(book, sheet, rule.ranges, (v) => {
      if (typeof v === 'number') nums.push(v);
      if (v != null && v !== '') { const k = typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${v}`; counts.set(k, (counts.get(k) || 0) + 1); }
    });
    p.counts = counts;
    p.nums = nums;
    p.min = nums.length ? Math.min(...nums) : 0;
    p.max = nums.length ? Math.max(...nums) : 0;
    p.avg = nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : 0;
    if (rule.type === 'top' || rule.type === 'bottom') {
      const s = nums.slice().sort((a, b) => (rule.type === 'top' ? b - a : a - b));
      const n = rule.percent ? Math.max(1, Math.floor((s.length * (rule.rank || 10)) / 100)) : rule.rank || 10;
      p.cut = s.length ? s[Math.min(n, s.length) - 1] : null;
    }
    if (rule.type === 'scale') {
      const stops = rule.stops && rule.stops.length === (rule.colors || []).length ? rule.stops : (rule.colors || []).length === 3 ? [{ type: 'min' }, { type: 'percentile', val: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }];
      p.sv = stops.map((s) => stopValue(s, nums, p.min, p.max));
      p.cols = (rule.colors || ['#F8696B', '#63BE7B']).map(hex);
    }
    if (rule.type === 'icons') {
      const n = +String(rule.set || '3')[0] || 3;
      const st = rule.stops && rule.stops.length === n ? rule.stops : Array.from({ length: n }, (_, i) => ({ type: 'percent', val: (i * 100) / n }));
      p.th = st.map((x) => ({ v: x.type === 'formula' ? Number(x.val) : stopValue(x.type === 'number' ? { type: 'num', val: x.val } : x, nums, p.min, p.max), gte: x.gte !== false }));
      p.n = n;
    }
    if (rule.type === 'bar') { const st = rule.stops || []; p.bmin = st[0] ? stopValue(st[0], nums, p.min, p.max) : Math.min(0, p.min); p.bmax = st[1] ? stopValue(st[1], nums, p.min, p.max) : p.max; }
  }
  return p;
}

function test(book, sheet, p, r, c) {
  const rule = p.rule;
  const v = book.value(sheet, r, c);
  const g = rule.ranges[0];
  const ev = (f) => {
    if (f == null || f === '') return null;
    const sf = shiftFormula(String(f), r - g.r1, c - g.c1);
    return book.calc.evalText(sf, sheet, r, c);
  };
  switch (rule.type) {
    case 'cell': {
      const a = ev(rule.f1); const b = ev(rule.f2);
      if (isErr(v) || isErr(a)) return false;
      const vv = v == null || v === '' ? (typeof a === 'number' ? 0 : '') : v;
      const ca = compare(vv, a);
      switch (rule.op) {
        case 'gt': return ca > 0; case 'lt': return ca < 0; case 'ge': return ca >= 0; case 'le': return ca <= 0; case 'eq': return ca === 0; case 'ne': return ca !== 0;
        case 'between': { const lo = compare(a, b) <= 0 ? a : b; const hi = lo === a ? b : a; return compare(vv, lo) >= 0 && compare(vv, hi) <= 0; }
        case 'notBetween': { const lo = compare(a, b) <= 0 ? a : b; const hi = lo === a ? b : a; return compare(vv, lo) < 0 || compare(vv, hi) > 0; }
        default: return false;
      }
    }
    case 'text': {
      const t = String(v ?? '').toLowerCase(); const q = String(rule.text || '').toLowerCase();
      return rule.op === 'notContains' ? !t.includes(q) : rule.op === 'begins' ? t.startsWith(q) : rule.op === 'ends' ? t.endsWith(q) : rule.op === 'eq' ? t === q : t.includes(q);
    }
    case 'dup': case 'unique': {
      if (v == null || v === '') return false;
      const k = typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${v}`;
      const n = p.counts.get(k) || 0;
      return rule.type === 'dup' ? n > 1 : n === 1;
    }
    case 'top': return typeof v === 'number' && p.cut != null && v >= p.cut;
    case 'bottom': return typeof v === 'number' && p.cut != null && v <= p.cut;
    case 'above': return typeof v === 'number' && v > p.avg;
    case 'below': return typeof v === 'number' && v < p.avg;
    case 'blank': return v == null || (typeof v === 'string' && !v.trim());
    case 'notBlank': return !(v == null || (typeof v === 'string' && !v.trim()));
    case 'error': return isErr(v);
    case 'noError': return !isErr(v);
    case 'formula': { const x = ev(rule.formula); const bb = toBool(Array.isArray(x) ? x[0][0] : x); return bb === true; }
    default: return false;
  }
}

export function cfAt(book, sheet) {
  if (!sheet.cf.length) return null;
  const key = `${book.version}:${sheet.version}`;
  if (sheet._cf && sheet._cf.key === key) return sheet._cf.fn;
  const rules = sheet.cf.filter((x) => x.ranges && x.ranges.length && x.type !== 'raw').slice().sort((a, b) => (a.priority || 0) - (b.priority || 0));
  const preps = new Map();
  const memo = new Map();
  const fn = (r, c) => {
    const k = r * MAXC + c;
    if (memo.has(k)) return memo.get(k);
    let out = null;
    for (const rule of rules) {
      if (!rule.ranges.some((g) => inR(g, r, c))) continue;
      let p = preps.get(rule);
      if (!p) { p = prepare(book, sheet, rule); preps.set(rule, p); }
      if (rule.type === 'scale') {
        const v = book.value(sheet, r, c);
        if (typeof v !== 'number') continue;
        const sv = p.sv; const cols = p.cols;
        let fill;
        if (v <= sv[0]) fill = mix(cols[0], cols[0], 0);
        else if (v >= sv[sv.length - 1]) fill = mix(cols[cols.length - 1], cols[cols.length - 1], 0);
        else { let i = 0; while (i < sv.length - 2 && v > sv[i + 1]) i++; const t = sv[i + 1] === sv[i] ? 0 : (v - sv[i]) / (sv[i + 1] - sv[i]); fill = mix(cols[i], cols[i + 1], t); }
        out = { ...(out || {}), fill: out && out.fill ? out.fill : fill };
        continue;
      }
      if (rule.type === 'icons') {
        const v = book.value(sheet, r, c);
        if (typeof v !== 'number' || (out && out.icon)) continue;
        let i = 0;
        for (let k = 1; k < p.n; k++) { const t = p.th[k]; if (t.gte ? v >= t.v : v > t.v) i = k; }
        if (rule.reverse) i = p.n - 1 - i;
        out = { ...(out || {}), icon: { set: rule.set || '3TrafficLights1', i, n: p.n, hide: rule.showValue === false } };
        continue;
      }
      if (rule.type === 'bar') {
        const v = book.value(sheet, r, c);
        if (typeof v !== 'number') continue;
        const span = p.bmax - p.bmin || 1;
        out = { ...(out || {}), bar: out && out.bar ? out.bar : { pct: Math.max(0, Math.min(1, (v - p.bmin) / span)), color: rule.color || '#638EC6' } };
        continue;
      }
      let ok = false;
      try { ok = test(book, sheet, p, r, c); } catch { ok = false; }
      if (!ok) continue;
      const s = rule.style || {};
      out = { ...s, ...(out || {}) }; // earlier (higher priority) rules win
      if (rule.stop) break;
    }
    memo.set(k, out);
    return out;
  };
  sheet._cf = { key, fn };
  return fn;
}
