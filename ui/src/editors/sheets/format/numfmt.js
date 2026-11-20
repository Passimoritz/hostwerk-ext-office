/*
 * Excel number format engine: formatValue(v, code, loc) → { text, color }.
 * Codes are compiled once (cached) into sections of small tokens; rendering just walks them.
 * Format codes are always the English/file form ('#,##0.00', 'dd.mm.yyyy'); separators in the OUTPUT follow loc.
 */
import { isErr } from '../formula/errors.js';
import { localErr } from '../formula/i18n.js';
import { partsFromSerial, dateToSerial } from './dates.js';
import { makeLocale } from './locale.js';

const DEFAULT_LOC = makeLocale();

export const BUILTIN_FORMATS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)', 7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};

const COLORS = { black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00' };
const PALETTE = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];

/* ---------------- compile ---------------- */

function splitSections(fmt) {
  const out = [];
  let cur = '';
  for (let i = 0; i < fmt.length; i++) {
    const c = fmt[i];
    if (c === '"') { const j = fmt.indexOf('"', i + 1); const e = j < 0 ? fmt.length - 1 : j; cur += fmt.slice(i, e + 1); i = e; continue; }
    if (c === '\\' || c === '_' || c === '*') { cur += fmt.slice(i, i + 2); i++; continue; }
    if (c === '[') { const j = fmt.indexOf(']', i); const e = j < 0 ? fmt.length - 1 : j; cur += fmt.slice(i, e + 1); i = e; continue; }
    if (c === ';') { out.push(cur); cur = ''; continue; }
    cur += c;
  }
  out.push(cur);
  return out;
}

const isDT = (t) => t.t === 'y' || t.t === 'M' || t.t === 'd' || t.t === 'h' || t.t === 's' || t.t === 'ampm' || t.t === 'el';
function litOf(t) {
  switch (t.t) {
    case 'lit': return t.s; case 'pt': return '.'; case 'comma': return ','; case 'slash': return '/'; case 'pct': return '%';
    case 'ph': return t.c === '?' ? ' ' : t.c; case 'digit': return t.s; case 'E': return t.plus ? 'E+' : 'E-'; case 'at': return '@';
    default: return '';
  }
}

function parseSection(str) {
  const toks = [];
  let color = null, cond = null;
  const lit = (s) => { const l = toks[toks.length - 1]; if (l && l.t === 'lit') l.s += s; else toks.push({ t: 'lit', s }); };
  for (let i = 0; i < str.length;) {
    const c = str[i];
    if (c === '"') { const j = str.indexOf('"', i + 1); const e = j < 0 ? str.length : j; lit(str.slice(i + 1, e)); i = e + 1; continue; }
    if (c === '\\') { if (i + 1 < str.length) lit(str[i + 1]); i += 2; continue; }
    if (c === '_') { lit(' '); i += 2; continue; }
    if (c === '*') { i += 2; continue; }
    if (c === '[') {
      const j = str.indexOf(']', i); const e = j < 0 ? str.length : j; const inner = str.slice(i + 1, e); i = e + 1;
      const lo = inner.toLowerCase().trim();
      let m;
      if (COLORS[lo]) color = COLORS[lo];
      else if ((m = /^colou?r\s*(\d+)$/.exec(lo))) color = `#${PALETTE[+m[1] - 1] || '000000'}`;
      else if ((m = /^(<=|>=|<>|<|>|=)\s*(-?\d*\.?\d+(?:e[+-]?\d+)?)$/.exec(lo))) cond = { op: m[1], val: +m[2] };
      else if (lo[0] === '$') { const sym = inner.slice(1).split('-')[0]; if (sym) lit(sym); }
      else if ((m = /^(h+|m+|s+)$/.exec(lo))) toks.push({ t: 'el', u: m[1][0], n: m[1].length });
      continue;
    }
    const lo = c.toLowerCase();
    if (lo === 'g' && str.substr(i, 7).toLowerCase() === 'general') { toks.push({ t: 'gen' }); i += 7; continue; }
    if (lo === 'a') {
      if (str.substr(i, 5).toUpperCase() === 'AM/PM') { toks.push({ t: 'ampm', full: true }); i += 5; continue; }
      if (str.substr(i, 3).toUpperCase() === 'A/P') { toks.push({ t: 'ampm', full: false }); i += 3; continue; }
    }
    if (lo === 'y' || lo === 'm' || lo === 'd' || lo === 'h' || lo === 's') {
      let j = i; while (j < str.length && str[j].toLowerCase() === lo) j++;
      toks.push({ t: lo === 'm' ? 'M' : lo, n: j - i }); i = j; continue;
    }
    if (lo === 'e' && (str[i + 1] === '+' || str[i + 1] === '-')) { toks.push({ t: 'E', plus: str[i + 1] === '+' }); i += 2; continue; }
    if (c === '0' || c === '#' || c === '?') { toks.push({ t: 'ph', c }); i++; continue; }
    if (c === '.') { toks.push({ t: 'pt' }); i++; continue; }
    if (c === ',') { toks.push({ t: 'comma' }); i++; continue; }
    if (c === '%') { toks.push({ t: 'pct' }); i++; continue; }
    if (c === '/') { toks.push({ t: 'slash' }); i++; continue; }
    if (c === '@') { toks.push({ t: 'at' }); i++; continue; }
    if (c >= '1' && c <= '9') { toks.push({ t: 'digit', s: c }); i++; continue; }
    lit(c); i++;
  }
  const sec = { toks, color, cond };
  if (toks.some(isDT)) return dateSection(sec);
  if (toks.some((t) => t.t === 'ph')) return numSection(sec);
  if (toks.some((t) => t.t === 'gen')) { sec.type = 'general'; return sec; }
  if (toks.some((t) => t.t === 'at')) { sec.type = 'text'; return sec; }
  sec.type = toks.length ? 'lit' : 'empty';
  return sec;
}

function dateSection(sec) {
  const T = sec.toks, out = [];
  let ampm = false, subsec = 0;
  for (let i = 0; i < T.length; i++) {
    const t = T[i];
    if (t.t === 'ampm') ampm = true;
    const prev = out[out.length - 1];
    if (t.t === 'pt' && prev && (prev.t === 's' || (prev.t === 'el' && prev.u === 's')) && T[i + 1] && T[i + 1].t === 'ph' && T[i + 1].c === '0') {
      let n = 0; while (T[i + 1 + n] && T[i + 1 + n].t === 'ph' && T[i + 1 + n].c === '0') n++;
      const k = Math.min(n, 3);
      out.push({ t: 'sub', n: k }); subsec = Math.max(subsec, k); i += n; continue;
    }
    if (isDT(t)) out.push(t); else out.push({ t: 'lit', s: litOf(t) });
  }
  const idx = [];
  out.forEach((t, i) => { if (t.t !== 'lit' && t.t !== 'ampm' && t.t !== 'sub') idx.push(i); });
  for (let k = 0; k < idx.length; k++) {
    const t = out[idx[k]];
    if (t.t !== 'M' || t.n > 2) continue;
    const prev = out[idx[k - 1]], next = out[idx[k + 1]];
    if ((prev && (prev.t === 'h' || (prev.t === 'el' && prev.u === 'h'))) || (next && (next.t === 's' || (next.t === 'el' && next.u === 's')))) t.t = 'min';
  }
  sec.type = 'date'; sec.toks = out; sec.ampm = ampm; sec.subsec = subsec;
  sec.hasDate = out.some((t) => t.t === 'y' || t.t === 'd' || t.t === 'M');
  return sec;
}

function numSection(sec) {
  const T = sec.toks;
  let pct = 0;
  for (const t of T) if (t.t === 'pct') pct++;
  sec.pct = pct;
  const eIdx = T.findIndex((t) => t.t === 'E');
  const slashIdx = eIdx >= 0 ? -1 : T.findIndex((t, i) => t.t === 'slash' && T.slice(0, i).some((x) => x.t === 'ph'));
  if (slashIdx >= 0) return fracSection(sec, slashIdx);
  const end = eIdx < 0 ? T.length : eIdx;
  let ptIdx = -1;
  for (let i = 0; i < end; i++) if (T[i].t === 'pt') { ptIdx = i; break; }
  const intEnd = ptIdx < 0 ? end : ptIdx;
  const phAt = (i) => T[i] && T[i].t === 'ph';
  let grouping = false, scale = 0;
  const items = [];
  for (let i = 0; i < T.length; i++) {
    const t = T[i];
    if (t.t === 'comma') {
      let p = i - 1; while (p >= 0 && T[p].t === 'comma') p--;
      let n = i + 1; while (n < T.length && T[n].t === 'comma') n++;
      if (i < intEnd && phAt(p) && phAt(n) && n < intEnd) { grouping = true; continue; }
      if (phAt(p)) { scale++; continue; }
      items.push({ t: 'lit', s: ',' }); continue;
    }
    if (t.t === 'ph') { items.push({ t: 'ph', c: t.c, r: i < intEnd ? 'i' : i < end ? 'd' : 'e' }); continue; }
    if (t.t === 'pt' && i === ptIdx) { items.push({ t: 'pt' }); continue; }
    if (t.t === 'E') { items.push(t); continue; }
    if (t.t === 'lit') { items.push(t); continue; }
    items.push({ t: 'lit', s: litOf(t) });
  }
  sec.type = 'num'; sec.items = items; sec.scale = scale; sec.grouping = grouping; sec.sci = eIdx >= 0;
  sec.intPh = items.filter((t) => t.r === 'i').map((t) => t.c);
  sec.decPh = items.filter((t) => t.r === 'd').map((t) => t.c);
  sec.expPh = items.filter((t) => t.r === 'e').map((t) => t.c);
  sec.plus = eIdx >= 0 && T[eIdx].plus;
  return sec;
}

function fracSection(sec, s) {
  const T = sec.toks;
  let j = s + 1;
  const den = [];
  while (j < T.length && (T[j].t === 'ph' || T[j].t === 'digit')) { den.push(T[j]); j++; }
  let k = s - 1;
  const num = [];
  while (k >= 0 && T[k].t === 'ph') { num.unshift(T[k].c); k--; }
  const items = [];
  for (let i = 0; i <= k; i++) {
    const t = T[i];
    items.push(t.t === 'ph' ? { t: 'ph', c: t.c, r: 'w' } : { t: 'lit', s: litOf(t) });
  }
  items.push({ t: 'frac' });
  for (let i = j; i < T.length; i++) items.push({ t: 'lit', s: litOf(T[i]) });
  sec.type = 'frac'; sec.items = items; sec.numPh = num;
  sec.denPh = den.map((d) => (d.t === 'digit' ? d.s : d.c));
  sec.fixedDen = den.some((d) => d.t === 'digit') ? Number(sec.denPh.join('').replace(/[#?]/g, '0')) || 0 : 0;
  sec.wholePh = items.filter((t) => t.r === 'w').map((t) => t.c);
  return sec;
}

const cache = new Map();
function compile(code) {
  let c = cache.get(code);
  if (c) return c;
  const sections = splitSections(code).map(parseSection);
  c = { sections, hasCond: sections.some((s) => s.cond), general: sections.length === 1 && sections[0].type === 'general' && sections[0].toks.length === 1 };
  if (cache.size > 4000) cache.clear();
  cache.set(code, c);
  return c;
}

function resolveFmt(fmt, loc) {
  if (fmt == null || fmt === '') return 'General';
  if (typeof fmt === 'number') return fmt === 14 ? loc.shortDate : fmt === 22 ? loc.dateTime : BUILTIN_FORMATS[fmt] || 'General';
  const lo = fmt.toLowerCase();
  return lo === 'general' || lo === 'standard' ? 'General' : fmt;
}

/* ---------------- numbers ---------------- */

function roundFixed(x, d) {
  const y = Number(x.toPrecision(15));
  if (y >= 1e21) return y.toLocaleString('en-US', { useGrouping: false, minimumFractionDigits: d, maximumFractionDigits: d });
  const s = String(y);
  if (s.includes('e')) return y.toFixed(d);
  return Number(`${Math.round(Number(`${s}e${d}`))}e-${d}`).toFixed(d);
}
function roundSplit(x, d) {
  const s = roundFixed(x, d);
  const i = s.indexOf('.');
  return i < 0 ? [s, ''] : [s.slice(0, i), s.slice(i + 1)];
}

function fillInt(digits, phs) {
  const out = new Array(phs.length).fill('');
  let di = digits.length - 1;
  for (let k = phs.length - 1; k >= 0; k--) {
    if (di >= 0) { if (k === 0) { out[k] = digits.slice(0, di + 1); di = -1; } else out[k] = digits[di--]; }
    else out[k] = phs[k] === '0' ? '0' : phs[k] === '?' ? ' ' : '';
  }
  return out;
}
function group(str, sep) {
  const m = /^(\s*)(\d*)(.*)$/.exec(str);
  const d = m[2];
  let out = '';
  for (let i = 0; i < d.length; i++) { if (i && (d.length - i) % 3 === 0) out += sep; out += d[i]; }
  return m[1] + out + m[3];
}
function decDigits(dp, phs) {
  const arr = phs.map((c, k) => dp[k] || '0');
  for (let k = arr.length - 1; k >= 0; k--) {
    if (phs[k] === '0' || arr[k] !== '0') break;
    arr[k] = phs[k] === '?' ? ' ' : '';
  }
  return arr;
}

function emit(sec, intDigits, dp, expStr, loc) {
  const intOut = fillInt(intDigits, sec.intPh);
  const intStr = sec.grouping ? group(intOut.join(''), loc.group) : null;
  const decOut = decDigits(dp, sec.decPh);
  let s = '', ii = 0, di = 0;
  let pending = !sec.intPh.length && intDigits;
  for (const it of sec.items) {
    if (it.t === 'lit') s += it.s;
    else if (it.t === 'ph') {
      if (it.r === 'i') { s += sec.grouping ? (ii === 0 ? intStr : '') : intOut[ii]; ii++; }
      else if (it.r === 'd') s += decOut[di++];
    } else if (it.t === 'pt') { if (pending) { s += intDigits; pending = false; } s += loc.decimal; }
    else if (it.t === 'E') s += expStr;
  }
  if (pending) s = intDigits + s;
  return s;
}

function renderNum(sec, a, loc) {
  const x = a * 100 ** sec.pct / 1000 ** sec.scale;
  if (!sec.sci) {
    const [ip, dp] = roundSplit(x, sec.decPh.length);
    return emit(sec, ip === '0' ? '' : ip, dp, '', loc);
  }
  const ni = sec.intPh.length, nd = sec.decPh.length;
  const eng = ni > 1 && sec.intPh.includes('#');
  let exp = 0, ip, dp;
  if (x !== 0) {
    const e = Math.floor(Math.log10(x));
    exp = eng ? Math.floor(e / ni) * ni : e - (Math.max(ni, 1) - 1);
    for (let t = 0; t < 3; t++) {
      [ip, dp] = roundSplit(x / 10 ** exp, nd);
      if (Number(ip) >= 10 ** (eng ? ni : Math.max(ni, 1))) exp += eng ? ni : 1; else break;
    }
  } else [ip, dp] = roundSplit(0, nd);
  const minE = sec.expPh.filter((c) => c === '0').length;
  const expStr = `E${exp < 0 ? '-' : sec.plus ? '+' : ''}${String(Math.abs(exp)).padStart(minE, '0')}`;
  return emit(sec, ip === '0' ? '' : ip, dp, expStr, loc);
}

function bestFrac(x, maxD) {
  if (x === 0) return [0, 1];
  let p0 = 0, q0 = 1, p1 = 1, q1 = 0, v = x;
  for (let i = 0; i < 64; i++) {
    const a = Math.floor(v);
    const p2 = a * p1 + p0, q2 = a * q1 + q0;
    if (q2 > maxD) {
      const k = Math.floor((maxD - q0) / q1);
      const ps = p0 + k * p1, qs = q0 + k * q1;
      if (qs > 0 && Math.abs(x - ps / qs) < Math.abs(x - p1 / q1)) return [ps, qs];
      return [p1, q1];
    }
    p0 = p1; q0 = q1; p1 = p2; q1 = q2;
    const r = v - a;
    if (r < 1e-10) break;
    v = 1 / r;
  }
  return [p1, q1];
}

function renderFrac(sec, a) {
  const x = a * 100 ** sec.pct;
  const hasWhole = sec.wholePh.length > 0;
  let w = hasWhole ? Math.floor(x) : 0;
  const f = x - w;
  let nu, de;
  if (sec.fixedDen) { de = sec.fixedDen; nu = Math.round(f * de); }
  else [nu, de] = bestFrac(f, 10 ** Math.max(1, sec.denPh.length) - 1);
  if (hasWhole && nu === de && nu > 0) { w += 1; nu = 0; }
  const numW = sec.numPh.length, denW = sec.denPh.length;
  let fracStr;
  if (hasWhole && nu === 0) fracStr = ' '.repeat(numW + 1 + denW);
  else {
    let ns = String(nu);
    if (ns.length < numW) ns = ns.padStart(numW, sec.numPh.includes('0') ? '0' : sec.numPh.includes('?') ? ' ' : '');
    let ds = String(de);
    if (!sec.fixedDen && ds.length < denW && sec.denPh.includes('?')) ds = ds.padEnd(denW, ' ');
    fracStr = `${ns}/${ds}`;
  }
  let wholeStr = hasWhole ? fillInt(w === 0 ? '' : String(w), sec.wholePh).join('') : '';
  if (hasWhole && nu === 0 && !wholeStr.trim()) wholeStr = '0';
  const blankWhole = hasWhole && !wholeStr.trim();
  let s = '', seenW = 0, afterWhole = false;
  for (const it of sec.items) {
    if (it.t === 'ph') { if (seenW++ === 0) s += wholeStr; afterWhole = seenW === sec.wholePh.length; }
    else if (it.t === 'frac') { s += fracStr; afterWhole = false; }
    else if (!(blankWhole && afterWhole)) s += it.s;
  }
  return s;
}

/* ---------------- dates ---------------- */
const pad2 = (n) => (n < 10 ? `0${n}` : String(n));

function renderDate(sec, v, loc) {
  if (v < 0 || v > 2958465.999999) return '########';
  const unit = sec.subsec ? 10 ** (3 - sec.subsec) : 1000;
  const total = Math.round(Math.round(v * 86400000) / unit) * unit;
  const p = partsFromSerial(total / 86400000);
  let s = '';
  for (const t of sec.toks) {
    switch (t.t) {
      case 'lit': s += t.s; break;
      case 'y': s += t.n <= 2 ? pad2(p.y % 100) : String(p.y); break;
      case 'M': s += t.n === 1 ? p.m : t.n === 2 ? pad2(p.m) : t.n === 3 ? loc.monthsShort[p.m - 1] : t.n === 5 ? loc.months[p.m - 1][0] : loc.months[p.m - 1]; break;
      case 'min': s += t.n === 1 ? p.mi : pad2(p.mi); break;
      case 'd': s += t.n === 1 ? p.d : t.n === 2 ? pad2(p.d) : t.n === 3 ? loc.daysShort[p.dow] : loc.days[p.dow]; break;
      case 'h': { const h = sec.ampm ? p.h % 12 || 12 : p.h; s += t.n === 1 ? h : pad2(h); break; }
      case 's': s += t.n === 1 ? p.s : pad2(p.s); break;
      case 'sub': s += loc.decimal + String(p.ms).padStart(3, '0').slice(0, t.n); break;
      case 'el': s += String(Math.floor(total / (t.u === 'h' ? 3600000 : t.u === 'm' ? 60000 : 1000))).padStart(t.n, '0'); break;
      case 'ampm': s += t.full ? (p.h < 12 ? 'AM' : 'PM') : p.h < 12 ? 'A' : 'P'; break;
      default: break;
    }
  }
  return s;
}

/* ---------------- General ---------------- */

/** Excel's "General" display (≈ 11 characters, scientific for very large/small values) */
export function fmtGeneral(n, loc = DEFAULT_LOC) {
  if (n === 0) return '0';
  const a = Math.abs(n);
  let s;
  if (a >= 1e11 || a < 1e-4) {
    let e = Math.floor(Math.log10(a));
    let m = roundFixed(a / 10 ** e, 5);
    if (Number(m) >= 10) { e++; m = roundFixed(a / 10 ** e, 5); }
    m = m.replace(/\.?0+$/, '');
    s = `${m}E${e < 0 ? '-' : '+'}${pad2(Math.abs(e))}`;
  } else if (Number.isInteger(a)) s = String(a);
  else {
    const intDigits = a >= 1 ? Math.floor(Math.log10(a)) + 1 : 0;
    s = roundFixed(a, intDigits > 0 ? Math.max(0, 10 - intDigits) : 9);
    if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  }
  if (loc.decimal !== '.') s = s.replace('.', loc.decimal);
  return n < 0 ? `-${s}` : s;
}

/* ---------------- public ---------------- */

function testCond(c, v) {
  switch (c.op) {
    case '<': return v < c.val; case '<=': return v <= c.val; case '>': return v > c.val; case '>=': return v >= c.val;
    case '=': return v === c.val; case '<>': return v !== c.val; default: return false;
  }
}
function pick(c, v) {
  const S = c.sections.length >= 4 ? c.sections.slice(0, 3) : c.sections;
  if (c.hasCond) {
    for (const s of S) {
      if (s.type === 'text') continue;
      if (!s.cond || testCond(s.cond, v)) return [s, !(s.cond && (s.cond.op === '<' || s.cond.op === '<=') && s.cond.val <= 0)];
    }
    return [null, true];
  }
  if (S.length === 1 || S[1].type === 'text') return [S[0], true];
  if (S.length === 2 || S[2].type === 'text') return v >= 0 ? [S[0], true] : [S[1], false];
  return v > 0 ? [S[0], true] : v < 0 ? [S[1], false] : [S[2], true];
}

const plain = (text) => ({ text, color: null });

/** Render a cell value with an Excel format code (or a built-in id) for a locale */
export function formatValue(v, fmt, loc = DEFAULT_LOC) {
  if (v == null || v === '') return plain('');
  if (isErr(v)) return plain(localErr(v.code, loc));
  if (typeof v === 'boolean') return plain(v ? loc.trueStr : loc.falseStr);
  if (v instanceof Date) v = dateToSerial(v);
  const c = compile(resolveFmt(fmt, loc));
  if (typeof v === 'string') {
    const sec = c.sections.length >= 4 ? c.sections[3] : c.sections[0].type === 'text' ? c.sections[0] : null;
    if (!sec) return plain(v);
    let s = '';
    for (const t of sec.toks) s += t.t === 'at' ? v : t.t === 'lit' ? t.s : '';
    return { text: s, color: sec.color };
  }
  if (typeof v !== 'number') return plain(String(v));
  if (!Number.isFinite(v)) return plain('#NUM!');
  if (c.general) return plain(fmtGeneral(v, loc));
  const [sec, minus] = pick(c, v);
  if (!sec) return plain(fmtGeneral(v, loc));
  if (sec.type === 'date') return { text: renderDate(sec, v, loc), color: sec.color };
  const a = Math.abs(v);
  let text;
  switch (sec.type) {
    case 'num': text = renderNum(sec, a, loc); break;
    case 'frac': text = renderFrac(sec, a); break;
    case 'general': text = sec.toks.map((t) => (t.t === 'gen' ? fmtGeneral(a, loc) : t.t === 'lit' ? t.s : '')).join(''); break;
    case 'text': text = fmtGeneral(a, loc); break;
    case 'lit': text = sec.toks.map((t) => t.s || '').join(''); break;
    default: return { text: '', color: sec.color };
  }
  if (minus && v < 0) text = `-${text}`;
  return { text, color: sec.color };
}

export const isDateFormat = (fmt) => compile(resolveFmt(fmt, DEFAULT_LOC)).sections[0].type === 'date';
export const isTimeOnly = (fmt) => { const s = compile(resolveFmt(fmt, DEFAULT_LOC)).sections[0]; return s.type === 'date' && !s.hasDate; };

/** 'general'|'number'|'currency'|'accounting'|'percent'|'scientific'|'fraction'|'date'|'time'|'text'|'custom' */
export function fmtKind(fmt) {
  const code = resolveFmt(fmt, DEFAULT_LOC);
  const c = compile(code);
  const s = c.sections[0];
  if (c.general) return 'general';
  if (s.type === 'date') return s.hasDate ? 'date' : 'time';
  if (s.type === 'text') return 'text';
  if (s.type === 'frac') return 'fraction';
  if (s.type !== 'num') return 'custom';
  if (s.sci) return 'scientific';
  if (s.pct) return 'percent';
  if (code.includes('*') && code.includes('_')) return 'accounting';
  const lits = s.items.filter((t) => t.t === 'lit').map((t) => t.s).join('');
  if (/[€$£¥]|\b[A-Z]{3}\b/.test(lits)) return 'currency';
  if (!lits.trim()) return 'number';
  return 'custom';
}

/** Decimals Excel's General shows for v (base for "Increase/Decrease decimals" on General cells) */
export function decimalsForValue(v) {
  if (typeof v !== 'number' || !Number.isFinite(v)) return 0;
  const m = /\.(\d+)/.exec(fmtGeneral(v, DEFAULT_LOC).split('E')[0]);
  return m ? m[1].length : 0;
}

function shiftDec(s, delta) {
  const act = [];
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === '"') { const j = s.indexOf('"', i + 1); i = j < 0 ? s.length : j; continue; }
    if (c === '\\' || c === '_' || c === '*') { i++; continue; }
    if (c === '[') { const j = s.indexOf(']', i); i = j < 0 ? s.length : j; continue; }
    if (c === '0' || c === '#' || c === '?' || c === '.') act.push({ i, c });
    else if ((c === 'E' || c === 'e') && (s[i + 1] === '+' || s[i + 1] === '-')) { act.push({ i, c: 'E' }); i++; }
  }
  const eK = act.findIndex((a) => a.c === 'E');
  const lim = eK < 0 ? act.length : eK;
  const pK = act.slice(0, lim).findIndex((a) => a.c === '.');
  if (pK >= 0) {
    let last = pK;
    while (last + 1 < lim && act[last + 1].c !== '.') last++;
    if (delta > 0) { const at = act[last].i + 1; return s.slice(0, at) + '0'.repeat(delta) + s.slice(at); }
    const nDec = last - pK;
    const rm = Math.min(nDec, -delta);
    if (rm === nDec) return s.slice(0, act[pK].i) + s.slice(act[last].i + 1);
    return s.slice(0, act[last - rm].i + 1) + s.slice(act[last].i + 1);
  }
  if (delta <= 0) return s;
  let lastInt = -1;
  for (let k = 0; k < lim; k++) if (act[k].c !== '.') lastInt = k;
  if (lastInt < 0) return s;
  const at = act[lastInt].i + 1;
  return `${s.slice(0, at)}.${'0'.repeat(delta)}${s.slice(at)}`;
}

/** Increase (delta > 0) / decrease (delta < 0) the decimals of a format; v = the cell value (needed for General) */
export function changeDecimals(fmt, delta, v) {
  const code = resolveFmt(fmt, DEFAULT_LOC);
  const c = compile(code);
  const s0 = c.sections[0];
  if (c.general || s0.type === 'general') {
    const d = Math.max(0, decimalsForValue(v) + delta);
    return d ? `0.${'0'.repeat(d)}` : '0';
  }
  if (s0.type === 'date' || s0.type === 'text' || s0.type === 'frac') return code;
  return splitSections(code).map((s, i) => (c.sections[i] && c.sections[i].type === 'num' ? shiftDec(s, delta) : s)).join(';');
}

export function makeNumberFormat({ decimals = 2, thousands = false, negRed = false, negParens = false } = {}) {
  const base = (thousands ? '#,##0' : '0') + (decimals > 0 ? `.${'0'.repeat(decimals)}` : '');
  if (negParens) return `${base}_);${negRed ? '[Red]' : ''}(${base})`;
  if (negRed) return `${base};[Red]-${base}`;
  return base;
}

const quoteSym = (sym) => (/^[€$£¥]$/.test(sym) ? sym : `"${String(sym).replace(/"/g, '')}"`);
export function makeCurrencyFormat({ symbol = '€', decimals = 2, after = true, negRed = false } = {}) {
  const num = `#,##0${decimals > 0 ? `.${'0'.repeat(decimals)}` : ''}`;
  const q = quoteSym(symbol);
  const pos = after ? `${num} ${q}` : `${q}${num}`;
  return negRed ? `${pos};[Red]-${pos}` : pos;
}

export function ACCOUNTING(loc = DEFAULT_LOC) {
  if (loc.lang === 'de' || loc.currency === '€') return '_-* #,##0.00 €_-;-* #,##0.00 €_-;_-* "-"?? €_-;_-@_-';
  return '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
}

/** Category presets for a "Format cells" dialog (labels English – translated by the UI) */
export function FORMAT_PRESETS(loc = DEFAULT_LOC) {
  return [
    { id: 'general', label: 'General', code: 'General' },
    { id: 'number', label: 'Number', code: '0.00' },
    { id: 'currency', label: 'Currency', code: loc.currencyFmt },
    { id: 'accounting', label: 'Accounting', code: ACCOUNTING(loc) },
    { id: 'shortDate', label: 'Short date', code: loc.shortDate },
    { id: 'longDate', label: 'Long date', code: loc.longDate },
    { id: 'time', label: 'Time', code: loc.timeSec },
    { id: 'percent', label: 'Percentage', code: '0.00%' },
    { id: 'fraction', label: 'Fraction', code: '# ?/?' },
    { id: 'scientific', label: 'Scientific', code: '0.00E+00' },
    { id: 'text', label: 'Text', code: '@' },
  ];
}

/** Common custom codes for the "Custom" list */
export function CUSTOM_FORMATS(loc = DEFAULT_LOC) {
  return [
    'General', '0', '0.00', '#,##0', '#,##0.00', '#,##0.00;[Red]-#,##0.00', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)',
    loc.currencyFmt, makeCurrencyFormat({ symbol: loc.currency, decimals: 0, after: loc.lang === 'de' }), `${loc.currencyFmt};[Red]-${loc.currencyFmt}`,
    '0%', '0.00%', '0.00E+00', '##0.0E+0', '# ?/?', '# ??/??', '?/4', '?/8',
    loc.shortDate, 'dd.mm.yyyy', 'dd.mm.yy', 'yyyy-mm-dd', 'm/d/yyyy', 'd-mmm-yy', 'd. mmmm yyyy', 'mmm d, yyyy', 'mmmm yyyy', 'dddd', loc.longDate,
    'hh:mm', 'hh:mm:ss', 'h:mm AM/PM', 'h:mm:ss AM/PM', '[h]:mm:ss', 'mm:ss', 'mm:ss.0', `${loc.shortDate} hh:mm`, '@',
  ].filter((x, i, a) => a.indexOf(x) === i);
}

/** compiled sections of a format code (number formats → OpenDocument data styles, io/ods-styles.js) */
export const compileFormat = (fmt, loc = DEFAULT_LOC) => compile(resolveFmt(fmt, loc));
