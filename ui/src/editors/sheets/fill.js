/*
 * Fill handle: continue a series like Excel (numbers, dates, weekdays, months, quarters, "Item 1" …).
 * Pure – date helpers (partsFromSerial / serialFromParts from format/dates.js) are injected through opts.
 */

const LISTS = [
  ['weekday', ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']],
  ['weekday', ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']],
  ['weekday', ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']],
  ['weekday', ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']],
  ['monthname', ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']],
  ['monthname', ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']],
  ['monthname', ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember']],
  ['monthname', ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez']],
];
const QUARTER_RE = /^(Q|Qtr\s?|Quarter\s?|Quartal\s?)([1-4])$/i;
const TEXTNUM_RE = /^(.*?)(\d+)$/;

const round = (x) => (Number.isFinite(x) ? Number(x.toPrecision(15)) : x);
const caseOf = (s) => (s === s.toUpperCase() && s !== s.toLowerCase() ? 'upper' : s === s.toLowerCase() ? 'lower' : 'title');
const applyCase = (s, c) => (c === 'upper' ? s.toUpperCase() : c === 'lower' ? s.toLowerCase() : s);

/** least-squares line through (i, y_i) → value at index x */
function trend(ys) {
  const n = ys.length;
  if (n === 1) return (x) => ys[0] + x;
  let sx = 0; let sy = 0; let sxx = 0; let sxy = 0;
  ys.forEach((y, i) => { sx += i; sy += y; sxx += i * i; sxy += i * y; });
  const b = (n * sxy - sx * sy) / (n * sxx - sx * sx);
  const a = (sy - b * sx) / n;
  return (x) => round(a + b * x);
}

function listOf(items) {
  if (!items.every((it) => typeof it.v === 'string')) return null;
  for (const [kind, list] of LISTS) {
    const low = list.map((x) => x.toLowerCase());
    const idx = items.map((it) => low.indexOf(it.v.trim().toLowerCase()));
    if (idx.every((i) => i >= 0)) return { kind, list, idx, style: caseOf(items[0].v.trim()) };
  }
  return null;
}

function textNum(items) {
  if (!items.every((it) => typeof it.v === 'string')) return null;
  const q = items.map((it) => QUARTER_RE.exec(it.v.trim()));
  if (q.every(Boolean) && q.every((m) => m[1] === q[0][1])) return { quarter: true, prefix: q[0][1], nums: q.map((m) => Number(m[2])) };
  const m = items.map((it) => TEXTNUM_RE.exec(it.v));
  if (!m.every(Boolean) || !m.every((x) => x[1] === m[0][1])) return null;
  if (m[0][1] === '' ) return null; // pure digit strings are not a text series
  const width = m[0][2].length > 1 && m[0][2][0] === '0' ? m[0][2].length : 0;
  return { prefix: m[0][1], nums: m.map((x) => Number(x[2])), width };
}

function daysInMonth(y, m, ser, parts) {
  // first day of next month minus one day
  const ny = m === 12 ? y + 1 : y; const nm = m === 12 ? 1 : m + 1;
  return parts(ser(ny, nm, 1, 0, 0, 0) - 1).d;
}

function monthStep(items, opts) {
  const { partsFromSerial: parts, serialFromParts: ser } = opts;
  if (!parts || !ser || items.length < 2) return null;
  const p = items.map((it) => parts(it.v));
  const mi = p.map((x) => x.y * 12 + (x.m - 1));
  const step = mi[1] - mi[0];
  if (!step) return null;
  for (let i = 1; i < p.length; i++) {
    if (mi[i] - mi[i - 1] !== step) return null;
    const expected = Math.min(p[0].d, daysInMonth(p[i].y, p[i].m, ser, parts));
    if (p[i].d !== expected) return null;
    if (p[i].h !== p[0].h || p[i].mi !== p[0].mi || p[i].s !== p[0].s) return null;
  }
  return { step, p0: p[0], mi0: mi[0] };
}

const isNum = (it) => typeof it.v === 'number' && Number.isFinite(it.v);

/** 'copy'|'linear'|'date'|'month'|'year'|'weekday'|'monthname'|'textnum'|'cycle' */
export function detectSeriesKind(items, opts = {}) {
  if (!items || !items.length) return 'copy';
  if (opts.copy) return items.length === 1 ? 'copy' : 'cycle';
  if (items.every(isNum)) {
    const dates = items.every((it) => it.date);
    if (dates) {
      if (items.length === 1) return 'date';
      const ms = monthStep(items, opts);
      if (ms) return ms.step % 12 === 0 ? 'year' : 'month';
      return 'date';
    }
    if (items.some((it) => it.date)) return 'cycle';
    return items.length === 1 ? 'copy' : 'linear';
  }
  const l = listOf(items);
  if (l) return l.kind;
  if (textNum(items)) return 'textnum';
  return items.length === 1 ? 'copy' : 'cycle';
}

/** Next `count` values after `items` (in fill direction). */
export function fillSeries(items, count, opts = {}) {
  const out = [];
  if (!items || !items.length || count <= 0) return out;
  const n = items.length;
  const kind = detectSeriesKind(items, opts);
  const cyc = () => { for (let k = 0; k < count; k++) out.push(items[k % n].v); return out; };
  switch (kind) {
    case 'copy': case 'cycle': return cyc();
    case 'linear': case 'date': {
      const f = trend(items.map((it) => it.v));
      for (let k = 0; k < count; k++) out.push(f(n + k));
      return out;
    }
    case 'month': case 'year': {
      const { step, p0, mi0 } = monthStep(items, opts);
      const { serialFromParts: ser, partsFromSerial: parts } = opts;
      for (let k = 0; k < count; k++) {
        const t = mi0 + step * (n + k);
        const y = Math.floor(t / 12); const m = (t - y * 12) + 1;
        const d = Math.min(p0.d, daysInMonth(y, m, ser, parts));
        out.push(ser(y, m, d, p0.h || 0, p0.mi || 0, p0.s || 0));
      }
      return out;
    }
    case 'weekday': case 'monthname': {
      const { list, idx, style } = listOf(items);
      const L = list.length;
      // unwrap indices so Sun→Mon counts as +1
      const un = [idx[0]];
      for (let i = 1; i < n; i++) { let v = idx[i]; while (v < un[i - 1]) v += L; un.push(v); }
      const f = n === 1 ? (x) => un[0] + x : trend(un);
      for (let k = 0; k < count; k++) out.push(applyCase(list[(((Math.round(f(n + k))) % L) + L) % L], style));
      return out;
    }
    case 'textnum': {
      const t = textNum(items);
      const f = n === 1 ? (x) => t.nums[0] + x : trend(t.nums);
      for (let k = 0; k < count; k++) {
        const num = Math.round(f(n + k));
        if (t.quarter) out.push(`${t.prefix}${((((num - 1) % 4) + 4) % 4) + 1}`);
        else out.push(t.prefix + (num < 0 ? `-${String(-num).padStart(t.width, '0')}` : String(num).padStart(t.width, '0')));
      }
      return out;
    }
    default: return cyc();
  }
}
