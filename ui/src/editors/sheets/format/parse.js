/*
 * What a user types into a cell → { v, fmt? } (like Excel: numbers, percent, currency, dates, times, booleans,
 * errors – everything else stays text). Formulas ('=…') are NOT handled here.
 */
import { errOf, ERROR_CODES } from '../formula/errors.js';
import { internalErr } from '../formula/i18n.js';
import { serialFromParts, daysInMonth } from './dates.js';
import { makeLocale, MONTH_NAMES } from './locale.js';
import { makeCurrencyFormat } from './numfmt.js';

const DEFAULT_LOC = makeLocale();
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const reCache = new Map();
function numRe(loc) {
  const k = loc.decimal + loc.group;
  let re = reCache.get(k);
  if (!re) {
    re = new RegExp(`^(\\d{1,3}(?:${esc(loc.group)}\\d{3})+|\\d+)?(?:${esc(loc.decimal)}(\\d*))?(?:[eE]([+-]?\\d+))?$`);
    reCache.set(k, re);
  }
  return re;
}

/** unsigned number in loc's notation → { v, decLen, grouped, sci } | null */
function parseNum(str, loc) {
  const m = numRe(loc).exec(str);
  if (!m || (!m[1] && !m[2])) return null;
  const grouped = !!m[1] && m[1].includes(loc.group);
  const ip = (m[1] || '0').split(loc.group).join('');
  const v = Number(`${ip}.${m[2] || '0'}${m[3] ? `e${m[3]}` : ''}`);
  return Number.isFinite(v) ? { v, decLen: (m[2] || '').length, grouped, sci: !!m[3] } : null;
}

const CUR_PRE = /^([€$£¥])\s*(.*)$/;
const CUR_POST = /^(.*?)\s*([€$£¥]|EUR|USD|CHF|GBP)$/i;

function parseSigned(t, loc, loose) {
  let s = t, neg = false, pct = false, sym = null, after = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
  const sign = () => { if (s[0] === '-' || s[0] === '+') { if (s[0] === '-') neg = !neg; s = s.slice(1).trim(); } };
  sign();
  if (s.endsWith('%')) { pct = true; s = s.slice(0, -1).trim(); }
  let m;
  if ((m = CUR_PRE.exec(s))) { sym = m[1]; s = m[2].trim(); }
  else if ((m = CUR_POST.exec(s)) && m[1]) { sym = m[2]; s = m[1].trim(); after = true; }
  if (sym) sign();
  const n = parseNum(s, loc);
  if (!n) return null;
  let v = neg ? -n.v : n.v;
  if (pct) v = Number((v / 100).toPrecision(15));
  if (loose) return { v };
  let fmt;
  if (pct) fmt = n.decLen ? `0.${'0'.repeat(n.decLen)}%` : '0%';
  else if (sym) fmt = makeCurrencyFormat({ symbol: sym.length === 1 ? sym : sym.toUpperCase(), decimals: n.decLen ? 2 : 0, after });
  else if (n.sci) fmt = '0.00E+00';
  else if (n.grouped) fmt = n.decLen ? '#,##0.00' : '#,##0';
  return fmt ? { v, fmt } : { v };
}

const TIME_RE = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:[.,]\d+)?))?\s*([aApP][mM]?)?$/;
function parseTime(str) {
  const m = TIME_RE.exec(str);
  if (!m) return null;
  let h = +m[1];
  const mi = +m[2], s = m[3] ? +m[3].replace(',', '.') : 0;
  if (mi >= 60 || s >= 60) return null;
  if (m[4]) {
    if (h < 1 || h > 12) return null;
    const pm = m[4][0].toLowerCase() === 'p';
    h = (h % 12) + (pm ? 12 : 0);
  }
  return { frac: (h * 3600 + mi * 60 + s) / 86400, sec: !!m[3], ampm: !!m[4], over: h >= 24 };
}

let monthMap = null;
function monthOf(name) {
  if (!monthMap) {
    monthMap = new Map();
    for (const L of Object.values(MONTH_NAMES)) {
      L.months.forEach((n, i) => monthMap.set(n.toLowerCase(), i + 1));
      L.monthsShort.forEach((n, i) => monthMap.set(n.toLowerCase(), i + 1));
    }
    monthMap.set('mär', 3); monthMap.set('sept', 9); monthMap.set('maerz', 3);
  }
  return monthMap.get(name.toLowerCase()) || 0;
}

function year(y) {
  if (y === undefined || y === '') return new Date().getFullYear();
  if (y.length <= 2) { const n = +y; return n < 30 ? 2000 + n : 1900 + n; }
  if (y.length === 4) return +y;
  return NaN;
}
function mk(y, mo, d) {
  if (!(y >= 1900 && y <= 9999) || !(mo >= 1 && mo <= 12) || !(d >= 1 && d <= daysInMonth(y, mo))) return null;
  return serialFromParts(y, mo, d);
}

const NUM_DATE = /^(\d{1,4})([./-])(\d{1,2})(?:\2(\d{0,4}))?$/;
const NAME_DMY = /^(\d{1,2})\.?[\s-]*([A-Za-zÄäÖöÜü]{3,})\.?(?:[\s-]+(\d{2,4}))?$/;
const NAME_MDY = /^([A-Za-zÄäÖöÜü]{3,})\.?\s+(\d{1,2})(?:,?\s+(\d{4}))?$/;

function parseDatePart(str, loc) {
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(str);
  if (m) return mk(+m[1], +m[2], +m[3]);
  if ((m = NUM_DATE.exec(str))) {
    const [, a, sep, b, c] = m;
    if (a.length > 2 && loc.dateOrder !== 'YMD') return null;
    if (loc.dateOrder === 'DMY') {
      if (c === undefined && sep === '/') return null; // German "3/4" is a fraction
      return mk(year(c), +b, +a);
    }
    if (loc.dateOrder === 'MDY') return mk(year(c), +a, +b);
    return c ? mk(year(a), +b, +c) : mk(year(undefined), +a, +b);
  }
  if ((m = NAME_DMY.exec(str))) { const mo = monthOf(m[2]); return mo ? mk(year(m[3]), mo, +m[1]) : null; }
  if ((m = NAME_MDY.exec(str))) { const mo = monthOf(m[1]); return mo ? mk(year(m[3]), mo, +m[2]) : null; }
  return null;
}

function parseDateTime(t, loc) {
  const tm = parseTime(t);
  if (tm) return { v: tm.frac, fmt: tm.over ? '[h]:mm:ss' : tm.ampm && loc.lang === 'de' ? (tm.sec ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : tm.sec ? loc.timeSec : loc.time };
  const m = /^(.*?)(?:\s+|T)(\d{1,2}:\d{1,2}(?::\d{1,2}(?:[.,]\d+)?)?\s*(?:[aApP][mM]?)?)$/.exec(t);
  if (m) {
    const d = parseDatePart(m[1].trim(), loc);
    const tt = parseTime(m[2]);
    if (d != null && tt && !tt.over) return { v: d + tt.frac, fmt: loc.dateTime };
    return null;
  }
  const d = parseDatePart(t, loc);
  return d != null ? { v: d, fmt: loc.shortDate } : null;
}

const FRAC_RE = /^([-+]?)(?:(\d+)\s+)?(\d+)\/(\d+)$/;
function parseFraction(t) {
  const m = FRAC_RE.exec(t);
  if (!m || +m[4] === 0) return null;
  const v = (m[2] ? +m[2] : 0) + +m[3] / +m[4];
  return { v: m[1] === '-' ? -v : v, fmt: m[4].length > 1 ? '# ??/??' : '# ?/?' };
}

/** Typed text → { v, fmt? } */
export function parseInput(text, loc = DEFAULT_LOC) {
  if (text == null) return { v: null };
  const s = String(text);
  if (s === '') return { v: null };
  if (s[0] === "'") return { v: s.slice(1) };
  const t = s.trim();
  if (!t) return { v: s };
  const up = t.toUpperCase();
  if (up === 'TRUE' || up === loc.trueStr.toUpperCase() || (loc.lang === 'de' && up === 'WAHR')) return { v: true };
  if (up === 'FALSE' || up === loc.falseStr.toUpperCase() || (loc.lang === 'de' && up === 'FALSCH')) return { v: false };
  if (t[0] === '#' && ERROR_CODES.includes(internalErr(up))) return { v: errOf(internalErr(up)) };
  if (t[0] === '=') return { v: s };
  return parseSigned(t, loc, false) || parseDateTime(t, loc) || parseFraction(t) || { v: s };
}

/** A number in loc's notation (also with %, currency, sign, parentheses) or null – for CSV import */
export function parseNumberLoose(str, loc = DEFAULT_LOC) {
  if (str == null) return null;
  const t = String(str).trim();
  if (!t) return null;
  const r = parseSigned(t, loc, true);
  return r ? r.v : null;
}
