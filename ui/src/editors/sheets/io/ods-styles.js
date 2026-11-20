/*
 * OpenDocument spreadsheet styles ↔ the sheet model (io/ods.js):
 *  - Excel number format codes ↔ ODF data styles (number:number-style, date-style, time-style, percentage-style,
 *    text-style; sections → style:map like LibreOffice: "value()>=0" → P0, the main style is the last section)
 *  - cell style objects ↔ style:style family="table-cell" (fonts, fills, borders, alignment, rotation, protection)
 *  - units (cm/mm/in/pt → px)
 * String building + the small tree of _shared/xmltree.js – runs in Node tests.
 */
import { compileFormat } from '../format/numfmt.js';
import { kids, kid, text as textOfEl } from '../../_shared/xmltree.js';

export const xesc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '');

/* ---------------- units ---------------- */
export function toPx(s) {
  const m = /^\s*(-?[\d.]+(?:e-?\d+)?)\s*(cm|mm|in|inch|pt|pc|px)?\s*$/i.exec(String(s ?? ''));
  if (!m) return null;
  const v = +m[1];
  switch ((m[2] || 'px').toLowerCase()) {
    case 'cm': return (v * 96) / 2.54; case 'mm': return (v * 96) / 25.4; case 'in': case 'inch': return v * 96;
    case 'pt': return (v * 4) / 3; case 'pc': return v * 16; default: return v;
  }
}
export const toPt = (s) => { const p = toPx(s); return p == null ? null : p * 0.75; };
export const inch = (px) => `${(Math.max(0, px) / 96).toFixed(4)}in`;
export const cm = (px) => `${((px / 96) * 2.54).toFixed(3)}cm`;

/* ---------------- number formats → data styles ---------------- */
const COLOR_NAME = { '#000000': 'Black', '#0000FF': 'Blue', '#00FFFF': 'Cyan', '#00FF00': 'Green', '#FF00FF': 'Magenta', '#FF0000': 'Red', '#FFFFFF': 'White', '#FFFF00': 'Yellow' };
const OPS = { '<=': '&lt;=', '>=': '&gt;=', '<>': '!=', '<': '&lt;', '>': '&gt;', '=': '=' };
const ntext = (s) => (s ? `<number:text>${xesc(s)}</number:text>` : '');

function numberSectionBody(sec) {
  const items = sec.items;
  const firstPh = items.findIndex((t) => t.t === 'ph' || t.t === 'pt');
  let lastPh = -1; items.forEach((t, i) => { if (t.t === 'ph' || t.t === 'pt' || t.t === 'E') lastPh = i; });
  const lit = (a, b) => items.slice(a, b).filter((t) => t.t === 'lit').map((t) => t.s).join('');
  const before = firstPh < 0 ? lit(0, items.length) : lit(0, firstPh);
  const after = firstPh < 0 ? '' : lit(lastPh + 1, items.length);
  // literals between integer digits ("000-000") → embedded text, position = integer digits to the right
  const emb = [];
  if (firstPh >= 0) {
    let right = 0;
    for (let i = lastPh; i >= firstPh; i--) {
      const t = items[i];
      if (t.t === 'ph' && t.r === 'i') right++;
      else if (t.t === 'lit' && items.slice(firstPh, i).some((x) => x.r === 'i') && !items.slice(i).some((x) => x.t === 'pt' || x.r === 'd' || x.t === 'E') ) emb.unshift(`<number:embedded-text number:position="${right}">${xesc(t.s)}</number:embedded-text>`);
    }
  }
  const minInt = sec.intPh.filter((c) => c === '0').length;
  const dec = sec.decPh.length; const minDec = sec.decPh.filter((c) => c === '0').length;
  if (firstPh < 0) return { body: ntext(before), kind: sec.pct ? 'percentage' : 'number' };
  let num;
  if (sec.sci) num = `<number:scientific-number number:decimal-places="${dec}" number:min-integer-digits="${Math.max(minInt, 1)}" number:min-exponent-digits="${Math.max(1, sec.expPh.filter((c) => c === '0').length)}"${sec.grouping ? ' number:grouping="true"' : ''}/>`;
  else num = `<number:number number:decimal-places="${dec}"${minDec !== dec ? ` number:min-decimal-places="${minDec}"` : ''} number:min-integer-digits="${minInt}"${sec.grouping ? ' number:grouping="true"' : ''}${sec.scale ? ` number:display-factor="${1000 ** sec.scale}"` : ''}${emb.length ? `>${emb.join('')}</number:number>` : '/>'}`;
  return { body: ntext(before) + num + ntext(after), kind: sec.pct ? 'percentage' : 'number' };
}
function fracSectionBody(sec) {
  const items = sec.items; const fi = items.findIndex((t) => t.t === 'frac');
  const lit = (a, b) => items.slice(a, b).filter((t) => t.t === 'lit').map((t) => t.s).join('');
  const whole = sec.wholePh.length;
  const firstW = items.findIndex((t) => t.r === 'w');
  const before = lit(0, firstW >= 0 ? firstW : fi);
  const after = lit(fi + 1, items.length);
  const den = sec.fixedDen ? ` number:denominator-value="${sec.fixedDen}"` : '';
  return ntext(before.trimEnd()) + `<number:fraction${whole ? ` number:min-integer-digits="${sec.wholePh.filter((c) => c === '0').length}"` : ''} number:min-numerator-digits="${Math.max(1, sec.numPh.length)}" number:min-denominator-digits="${Math.max(1, sec.denPh.length)}"${den}/>` + ntext(after);
}
function dateSectionBody(sec) {
  let out = ''; let elapsed = false;
  const T = sec.toks;
  for (let i = 0; i < T.length; i++) {
    const t = T[i];
    const long = t.n >= 2 ? ' number:style="long"' : '';
    switch (t.t) {
      case 'y': out += `<number:year${t.n > 2 ? ' number:style="long"' : ''}/>`; break;
      case 'M': out += t.n >= 3 ? `<number:month number:textual="true"${t.n === 4 ? ' number:style="long"' : ''}/>` : `<number:month${long}/>`; break;
      case 'd': out += t.n >= 3 ? `<number:day-of-week${t.n >= 4 ? ' number:style="long"' : ''}/>` : `<number:day${long}/>`; break;
      case 'h': out += `<number:hours${long}/>`; break;
      case 'min': out += `<number:minutes${long}/>`; break;
      case 's': { const sub = T[i + 1] && T[i + 1].t === 'sub' ? T[i + 1].n : 0; out += `<number:seconds${long}${sub ? ` number:decimal-places="${sub}"` : ''}/>`; if (sub) i++; break; }
      case 'el': elapsed = true; out += t.u === 'h' ? `<number:hours${t.n >= 2 ? ' number:style="long"' : ''}/>` : t.u === 'm' ? `<number:minutes${t.n >= 2 ? ' number:style="long"' : ''}/>` : `<number:seconds${t.n >= 2 ? ' number:style="long"' : ''}/>`; break;
      case 'ampm': out += '<number:am-pm/>'; break;
      case 'sub': out += ntext('.') + ntext('0'.repeat(t.n)); break;
      case 'lit': out += ntext(t.s); break;
      default: break;
    }
  }
  return { body: out.replace(/<\/number:text><number:text>/g, ''), elapsed };
}
/** one compiled section → { tag, body, attrs } */
function sectionStyle(sec) {
  const color = sec.color ? `<style:text-properties fo:color="${sec.color.toLowerCase()}"/>` : '';
  if (sec.type === 'date') {
    const d = dateSectionBody(sec);
    return { tag: sec.hasDate ? 'number:date-style' : 'number:time-style', attrs: !sec.hasDate && d.elapsed ? ' number:truncate-on-overflow="false"' : '', body: color + d.body };
  }
  if (sec.type === 'num') { const n = numberSectionBody(sec); return { tag: n.kind === 'percentage' ? 'number:percentage-style' : 'number:number-style', attrs: '', body: color + n.body }; }
  if (sec.type === 'frac') return { tag: 'number:number-style', attrs: '', body: color + fracSectionBody(sec) };
  if (sec.type === 'text') return { tag: 'number:text-style', attrs: '', body: color + sec.toks.map((t) => (t.t === 'at' ? '<number:text-content/>' : t.t === 'lit' ? ntext(t.s) : '')).join('') };
  if (sec.type === 'general') return { tag: 'number:number-style', attrs: '', body: color + sec.toks.map((t) => (t.t === 'gen' ? '<number:number number:min-integer-digits="1"/>' : t.t === 'lit' ? ntext(t.s) : '')).join('') };
  return { tag: 'number:number-style', attrs: '', body: color + sec.toks.map((t) => (t.t === 'lit' ? ntext(t.s) : '')).join('') };
}

/**
 * Excel format code → { xml (all data styles, the main one named `name`), kind ('date'|'time'|'percentage'|'number'|
 * 'text'), general } – general codes need no data style (xml '').
 */
export function dataStyleXml(code, name, loc) {
  let comp;
  try { comp = compileFormat(code, loc); } catch { return { xml: '', kind: 'number', general: true }; }
  if (comp.general) return { xml: '', kind: 'number', general: true };
  const secs = comp.sections.slice();
  // a 4th (text) section has no place in a number style (ODF has no map to text styles in number styles)
  if (secs.length === 4 || (secs.length > 1 && secs[secs.length - 1].type === 'text')) secs.pop();
  const main = sectionStyle(secs[secs.length - 1]);
  let maps = ''; let pre = '';
  const conds = secs.length === 2 ? ['>=0'] : secs.length === 3 ? ['>0', '<0'] : [];
  for (let i = 0; i < secs.length - 1; i++) {
    const s = secs[i];
    const c = s.cond ? `${s.cond.op}${s.cond.val}` : conds[i];
    if (!c) continue;
    const m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(c);
    const pn = `${name}P${i}`;
    const st = sectionStyle(s);
    pre += `<${st.tag} style:name="${pn}" style:volatile="true"${st.attrs}>${st.body}</${st.tag}>`;
    maps += `<style:map style:condition="value()${OPS[m[1]]}${m[2]}" style:apply-style-name="${pn}"/>`;
  }
  const first = secs[0];
  const kind = first.type === 'date' ? (first.hasDate ? 'date' : 'time') : first.type === 'text' ? 'text' : first.type === 'num' && first.pct ? 'percentage' : 'number';
  // a single text section ("@") is the text style itself
  return { xml: `${pre}<${main.tag} style:name="${name}"${main.attrs}>${main.body}${maps}</${main.tag}>`, kind, general: false };
}

/* ---------------- data styles → number format codes ---------------- */
const q = (s) => (s === '' ? '' : /^[ \-–(),:/.$€£¥]+$/.test(s) && !/[0#?]/.test(s) ? s.replace(/[/:]/g, '\\$&').replace(/[.]/g, '\\.') : `"${s.replace(/"/g, '')}"`);
function intPart(minInt, grouping) {
  let z = '0'.repeat(minInt);
  if (!grouping) return z || '#';
  if (z.length < 4) z = '#,###'.slice(0, 5 - z.length) + z;
  else { let o = ''; for (let i = 0; i < z.length; i++) { if (i && (z.length - i) % 3 === 0) o += ','; o += z[i]; } z = o; }
  return z;
}
function numberCode(el, pct) {
  const a = el.a;
  const minInt = a['number:min-integer-digits'] != null ? +a['number:min-integer-digits'] : 1;
  let ip = intPart(minInt, a['number:grouping'] === 'true');
  // embedded text: insert into the integer digits (position = digits to the right)
  const emb = kids(el, 'number:embedded-text').map((e) => ({ p: +e.a['number:position'] || 0, t: textOfEl(e) })).sort((x, y) => y.p - x.p);
  for (const e of emb) {
    let count = 0; let at = ip.length;
    for (let i = ip.length - 1; i >= 0 && count < e.p; i--) { if (/[0#?]/.test(ip[i])) count++; at = i; }
    if (e.p === 0) at = ip.length;
    while (ip.replace(/[^0#?]/g, '').length < e.p) { ip = `#${ip}`; at++; }
    ip = ip.slice(0, at) + q(e.t) + ip.slice(at);
  }
  const dec = a['number:decimal-places'] != null ? +a['number:decimal-places'] : (el.n === 'number:number' && !emb.length ? null : 0);
  const minDec = a['number:min-decimal-places'] != null ? +a['number:min-decimal-places'] : a['loext:min-decimal-places'] != null ? +a['loext:min-decimal-places'] : dec;
  if (el.n === 'number:number' && dec == null && minInt <= 1 && a['number:grouping'] !== 'true') return pct ? '0' : 'General';
  let s = ip + (dec ? `.${'0'.repeat(Math.min(minDec ?? dec, dec))}${'#'.repeat(Math.max(0, dec - (minDec ?? dec)))}` : '');
  const f = +a['number:display-factor'] || 1;
  if (f >= 1000) s += ','.repeat(Math.round(Math.log10(f) / 3));
  if (el.n === 'number:scientific-number') s += `E+${'0'.repeat(Math.max(1, +a['number:min-exponent-digits'] || 2))}`;
  return s;
}
function oneStyleCode(st) {
  let out = ''; let color = '';
  const elapsed = st.a['number:truncate-on-overflow'] === 'false';
  const pct = st.n === 'number:percentage-style';
  for (const e of kids(st)) {
    const a = e.a; const long = a['number:style'] === 'long';
    switch (e.n) {
      case 'style:text-properties': { const c = String(a['fo:color'] || '').toUpperCase(); if (COLOR_NAME[c]) color = `[${COLOR_NAME[c]}]`; break; }
      case 'number:text': { const t = textOfEl(e); out += pct && t.includes('%') ? t.split('%').map(q).join('%') : q(t); break; }
      case 'number:number': case 'number:scientific-number': out += numberCode(e, pct); break;
      case 'number:fraction': {
        const w = a['number:min-integer-digits'];
        const n = Math.max(1, +a['number:min-numerator-digits'] || 1); const d = +a['number:denominator-value'] || 0;
        out += `${w != null ? `${'#'.repeat(+w ? 0 : 1)}${'0'.repeat(+w || 0)} ` : ''}${'?'.repeat(n)}/${d || '?'.repeat(Math.max(1, +a['number:min-denominator-digits'] || 1))}`;
        break;
      }
      case 'number:currency-symbol': { const t = textOfEl(e); out += t ? `[$${t}]` : ''; break; }
      case 'number:year': out += long ? 'yyyy' : 'yy'; break;
      case 'number:month': out += a['number:textual'] === 'true' ? (long ? 'mmmm' : 'mmm') : long ? 'mm' : 'm'; break;
      case 'number:day': out += long ? 'dd' : 'd'; break;
      case 'number:day-of-week': out += long ? 'dddd' : 'ddd'; break;
      case 'number:hours': out += elapsed && !/\[h/.test(out) ? (long ? '[hh]' : '[h]') : long ? 'hh' : 'h'; break;
      case 'number:minutes': out += long ? 'mm' : 'm'; break;
      case 'number:seconds': { out += long ? 'ss' : 's'; const d = +a['number:decimal-places'] || 0; if (d) out += `.${'0'.repeat(d)}`; break; }
      case 'number:am-pm': out += 'AM/PM'; break;
      case 'number:text-content': out += '@'; break;
      case 'number:boolean': out += 'TRUE'; break;
      case 'number:quarter': out += long ? '"Q"' : '"Q"'; break;
      default: break;
    }
  }
  if (pct && !out.includes('%')) out += '%';
  return color + out;
}
/** name → code for every data style below `root` elements (office:styles / office:automatic-styles) */
export function readDataStyles(roots) {
  const els = new Map();
  for (const r of roots) for (const e of kids(r)) if (/^number:(number|currency|percentage|date|time|boolean|text)-style$/.test(e.n) && e.a['style:name']) els.set(e.a['style:name'], e);
  const memo = new Map();
  const code = (name, depth = 0) => {
    if (memo.has(name)) return memo.get(name);
    const st = els.get(name); if (!st || depth > 4) return null;
    if (st.n === 'number:boolean-style') { memo.set(name, '"TRUE";"TRUE";"FALSE"'); return memo.get(name); }
    const main = oneStyleCode(st);
    const maps = kids(st, 'style:map').map((m) => ({ cond: String(m.a['style:condition'] || '').replace(/\s/g, ''), code: code(m.a['style:apply-style-name'], depth + 1) })).filter((m) => m.code != null);
    let out;
    if (!maps.length) out = main;
    else if (maps.length === 1 && maps[0].cond === 'value()>=0') out = `${maps[0].code};${main}`;
    else if (maps.length === 2 && maps[0].cond === 'value()>0' && maps[1].cond === 'value()<0') out = `${maps[0].code};${maps[1].code};${main}`;
    else {
      out = maps.slice(0, 2).map((m) => { const c = /^value\(\)(<=|>=|!=|<|>|=)(-?[\d.]+)$/.exec(m.cond); return `${c ? `[${c[1] === '!=' ? '<>' : c[1]}${c[2]}]` : ''}${m.code}`; }).join(';');
      out += `;${main}`;
    }
    memo.set(name, out || 'General');
    return memo.get(name);
  };
  const out = new Map();
  for (const name of els.keys()) { const c = code(name); if (c && c !== 'General') out.set(name, c); }
  return out;
}

/* ---------------- borders ---------------- */
const BW = { hair: 0.26, thin: 0.74, dotted: 0.74, dashed: 0.74, dashDot: 0.74, dashDotDot: 0.74, slantDashDot: 1.76, medium: 1.76, mediumDashed: 1.76, mediumDashDot: 1.76, mediumDashDotDot: 1.76, thick: 2.49, double: 2.01 };
const BS = { dotted: 'dotted', dashDotDot: 'dotted', mediumDashDotDot: 'dotted', hair: 'solid', dashed: 'dashed', mediumDashed: 'dashed', dashDot: 'dashed', mediumDashDot: 'dashed', slantDashDot: 'dashed', double: 'double' };
export function borderCss(e) { return `${BW[e.s] || 0.74}pt ${BS[e.s] || 'solid'} ${(e.c || '#000000').toLowerCase()}`; }
export function parseBorder(v) {
  if (!v || /^\s*none\s*$/i.test(v) || /hidden/.test(v)) return null;
  const parts = String(v).trim().split(/\s+/);
  let w = 0.74; let style = 'solid'; let color = '#000000';
  for (const p of parts) {
    if (/^#[0-9a-f]{6}$/i.test(p)) color = p.toUpperCase();
    else if (/^[\d.]+(pt|cm|mm|in|px)$/.test(p)) w = toPt(p);
    else if (/^(solid|dotted|dashed|double|dash-dot|dash-dot-dot|fine-dashed|double-thin|groove|ridge|inset|outset)$/.test(p)) style = p;
    else if (p === 'thin') w = 0.74; else if (p === 'medium') w = 1.76; else if (p === 'thick') w = 2.49;
  }
  if (w === 0) return null;
  let s;
  if (style === 'double' || style === 'double-thin') s = 'double';
  else if (style === 'dotted') s = w > 1.2 ? 'mediumDashDotDot' : 'dotted';
  else if (style === 'dashed' || style === 'fine-dashed') s = w > 1.2 ? 'mediumDashed' : 'dashed';
  else if (style === 'dash-dot') s = w > 1.2 ? 'mediumDashDot' : 'dashDot';
  else if (style === 'dash-dot-dot') s = w > 1.2 ? 'mediumDashDotDot' : 'dashDotDot';
  else s = w < 0.4 ? 'hair' : w <= 1.2 ? 'thin' : w <= 2.2 ? 'medium' : 'thick';
  return { s, c: color };
}

/* ---------------- cell style ↔ ODF properties ---------------- */
const HALIGN = { left: 'start', center: 'center', centerContinuous: 'center', right: 'end', justify: 'justify', distributed: 'justify', fill: 'start' };
const quoteFont = (f) => (/\s/.test(f) ? `'${f}'` : f);
/** font-face declaration */
export const fontFace = (f) => `<style:font-face style:name="${xesc(f)}" svg:font-family="${xesc(quoteFont(f))}"${/mono|courier|consol/i.test(f) ? ' style:font-pitch="fixed"' : ''}/>`;

/** text properties (cells, rich-text spans) */
export function textProps(s, def = {}) {
  const tp = [];
  if (s.font && s.font !== def.font) tp.push(`style:font-name="${xesc(s.font)}" style:font-name-asian="${xesc(s.font)}" style:font-name-complex="${xesc(s.font)}"`);
  if (s.size && s.size !== def.size) tp.push(`fo:font-size="${s.size}pt" style:font-size-asian="${s.size}pt" style:font-size-complex="${s.size}pt"`);
  if (s.b) tp.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
  if (s.i) tp.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
  if (s.u) tp.push(`style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"${s.u === 'double' ? ' style:text-underline-type="double"' : ''}`);
  if (s.st) tp.push('style:text-line-through-style="solid" style:text-line-through-type="single"');
  if (s.color) tp.push(`fo:color="${s.color.toLowerCase()}"`);
  const va = s.vert || s.va_;
  if (va === 'superscript') tp.push('style:text-position="super 58%"');
  else if (va === 'subscript') tp.push('style:text-position="sub 58%"');
  return tp.length ? `<style:text-properties ${tp.join(' ')}/>` : '';
}

/** model style → inner XML of a table-cell style (properties only) */
export function cellStyleProps(s, def) {
  const tcp = []; const pp = [];
  if (s.fill) tcp.push(`fo:background-color="${s.fill.toLowerCase()}"`);
  const bd = s.bd || {};
  const sides = ['t', 'b', 'l', 'r'].map((k) => (bd[k] ? borderCss(bd[k]) : 'none'));
  if (bd.t || bd.b || bd.l || bd.r) {
    if (sides.every((x) => x === sides[0])) tcp.push(`fo:border="${sides[0]}"`);
    else tcp.push(`fo:border-top="${sides[0]}" fo:border-bottom="${sides[1]}" fo:border-left="${sides[2]}" fo:border-right="${sides[3]}"`);
  }
  if (bd.d) { if (bd.d.down !== false) tcp.push(`style:diagonal-tl-br="${borderCss(bd.d)}"`); if (bd.d.up) tcp.push(`style:diagonal-bl-tr="${borderCss(bd.d)}"`); }
  if (s.va) tcp.push(`style:vertical-align="${s.va === 'center' || s.va === 'distributed' || s.va === 'justify' ? 'middle' : s.va === 'top' ? 'top' : 'bottom'}"`);
  if (s.wrap) tcp.push('fo:wrap-option="wrap"');
  if (s.shrink) tcp.push('style:shrink-to-fit="true"');
  if (s.rot) { if (s.rot === 255) tcp.push('style:direction="ttb"'); else tcp.push(`style:rotation-angle="${s.rot <= 90 ? s.rot : 360 - (s.rot - 90)}" style:rotation-align="none"`); }
  if (s.unlocked || s.hiddenF) tcp.push(`style:cell-protect="${s.unlocked ? (s.hiddenF ? 'formula-hidden' : 'none') : 'protected formula-hidden'}"`);
  if (s.h && s.h !== 'general') { tcp.push('style:text-align-source="fix"'); if (s.h === 'fill') tcp.push('style:repeat-content="true"'); pp.push(`fo:text-align="${HALIGN[s.h] || 'start'}"`); }
  if (s.indent) pp.push(`fo:margin-left="${(s.indent * 10).toFixed(1)}pt"`);
  return `${tcp.length ? `<style:table-cell-properties ${tcp.join(' ')}/>` : ''}${pp.length ? `<style:paragraph-properties ${pp.join(' ')}/>` : ''}${textProps(s, def)}`;
}

/** properties of one style:style element (no inheritance) → partial model style (+ data style name, maps) */
export function readStyleProps(el, fonts) {
  const s = {};
  const tcp = kid(el, 'style:table-cell-properties'); const pp = kid(el, 'style:paragraph-properties'); const tp = kid(el, 'style:text-properties');
  if (tcp) {
    const a = tcp.a;
    const bg = a['fo:background-color']; if (bg) s.fill = bg === 'transparent' ? null : bg.toUpperCase();
    const bd = {};
    if (a['fo:border'] != null) { const b = parseBorder(a['fo:border']); for (const k of ['t', 'b', 'l', 'r']) bd[k] = b; }
    for (const [attr, k] of [['fo:border-top', 't'], ['fo:border-bottom', 'b'], ['fo:border-left', 'l'], ['fo:border-right', 'r']]) if (a[attr] != null) bd[k] = parseBorder(a[attr]);
    const dd = a['style:diagonal-tl-br'] != null ? parseBorder(a['style:diagonal-tl-br']) : undefined; const du = a['style:diagonal-bl-tr'] != null ? parseBorder(a['style:diagonal-bl-tr']) : undefined;
    if (dd || du) bd.d = { ...(dd || du), down: !!dd, up: !!du }; else if (dd === null || du === null) bd.d = null;
    if (Object.keys(bd).length) s.bd = bd;
    const va = a['style:vertical-align']; if (va) s.va = va === 'middle' ? 'center' : va === 'top' ? 'top' : va === 'automatic' ? null : 'bottom';
    if (a['fo:wrap-option']) s.wrap = a['fo:wrap-option'] === 'wrap';
    if (a['style:shrink-to-fit']) s.shrink = a['style:shrink-to-fit'] === 'true';
    if (a['style:direction'] === 'ttb') s.rot = 255;
    else if (a['style:rotation-angle'] != null) { const r = Math.round(((parseFloat(a['style:rotation-angle']) % 360) + 360) % 360); s.rot = r === 0 ? null : r <= 90 ? r : r >= 270 ? 90 + (360 - r) : null; }
    const cp = a['style:cell-protect'];
    if (cp) { s.unlocked = cp === 'none' || cp === 'formula-hidden'; s.hiddenF = /formula-hidden|hidden-and-protected/.test(cp); }
    if (a['style:repeat-content'] === 'true') s.h = 'fill';
    if (a['style:text-align-source'] === 'value-type') s._autoAlign = true;
  }
  if (pp) {
    const ta = pp.a['fo:text-align'];
    if (ta && s.h !== 'fill' && !s._autoAlign) s.h = ta === 'center' ? 'center' : ta === 'end' || ta === 'right' ? 'right' : ta === 'justify' ? 'justify' : 'left';
    const ml = pp.a['fo:margin-left']; if (ml) { const pt = toPt(ml); if (pt && pt > 1) s.indent = Math.max(1, Math.round(pt / 10)); }
  }
  if (tp) Object.assign(s, readTextProps(tp, fonts));
  return s;
}
export function readTextProps(tp, fonts) {
  const s = {}; const a = tp.a;
  const fn = a['style:font-name'] || a['fo:font-family'];
  if (fn) s.font = (fonts && fonts.get(fn)) || String(fn).replace(/^['"]|['"]$/g, '');
  if (a['fo:font-size'] && /pt$/.test(a['fo:font-size'])) s.size = Math.round(parseFloat(a['fo:font-size']) * 2) / 2;
  if (a['fo:font-weight']) s.b = a['fo:font-weight'] === 'bold' || +a['fo:font-weight'] >= 600;
  if (a['fo:font-style']) s.i = a['fo:font-style'] === 'italic' || a['fo:font-style'] === 'oblique';
  if (a['style:text-underline-style']) s.u = a['style:text-underline-style'] === 'none' ? false : a['style:text-underline-type'] === 'double' ? 'double' : true;
  if (a['style:text-line-through-style']) s.st = a['style:text-line-through-style'] !== 'none';
  if (a['fo:color']) s.color = a['fo:color'].toUpperCase();
  const pos = a['style:text-position'];
  if (pos) s.vert = /^super|^\d/.test(pos) && !/^0%?\s/.test(pos) ? 'superscript' : /^sub|^-/.test(pos) ? 'subscript' : null;
  return s;
}
