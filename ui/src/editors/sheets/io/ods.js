/*
 * OpenDocument spreadsheets (ods / ots / fods) without the Office engine – own reader and writer with full styles:
 * cell styles (fonts, fills, borders, alignment, rotation, protection, number formats as data styles), merges, column
 * widths / row heights / hidden + outline groups, frozen panes, zoom and gridlines (settings.xml), conditional formats
 * (style:map for Excel + calcext:conditional-formats for LibreOffice), data validation, comments, links, rich text,
 * named ranges, autofilter, print ranges + orientation, images, charts (embedded chart objects).
 *
 * Patch saving: a file that came from disk keeps every part the editor does not write (styles.xml, meta.xml,
 * Thumbnails, Basic, Configurations2, embedded objects, unknown frames and spreadsheet elements such as data pilot
 * tables) – content.xml and settings.xml are written new, unchanged chart objects stay byte-identical.
 */
import JSZip from 'jszip';
import { Book, newSheet, uid } from '../model.js';
import { MAXC, MAXR, colName } from '../formula/refs.js';
import { fromOoxml } from '../formula/parser.js';
import { errOf, isErr } from '../formula/errors.js';
import { partsFromSerial, serialFromParts } from '../format/dates.js';
import { toOpenFormula, fromOpenFormula } from './odf-formula.js';
import { parseXml, kids, kid, all, text as textOfEl, serialize } from '../../_shared/xmltree.js';
import { xesc, toPx, inch, dataStyleXml, readDataStyles, cellStyleProps, readStyleProps, readTextProps, textProps, fontFace } from './ods-styles.js';
import { chartParts, chartContentXml, chartInlineXml, parseOdfChart, odfToExcelRange, excelToOdfRange } from './ods-chart.js';
import { resolveChartData } from '../chartData.js';

const K = (r, c) => r * MAXC + c;
const R = (k) => Math.floor(k / MAXC);
const C = (k) => k % MAXC;
export const ODS_MIME = { ods: 'application/vnd.oasis.opendocument.spreadsheet', ots: 'application/vnd.oasis.opendocument.spreadsheet-template', fods: 'application/vnd.oasis.opendocument.spreadsheet' };
const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0', text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0', draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink', dc: 'http://purl.org/dc/elements/1.1/', meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0', number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  presentation: 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0', svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', chart: 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
  dr3d: 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0', form: 'urn:oasis:names:tc:opendocument:xmlns:form:1.0', script: 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
  ooo: 'http://openoffice.org/2004/office', of: 'urn:oasis:names:tc:opendocument:xmlns:of:1.2', config: 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
  tableooo: 'http://openoffice.org/2009/table', calcext: 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0', loext: 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
  field: 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0', xhtml: 'http://www.w3.org/1999/xhtml',
};
const nsAttrs = (extra = {}) => Object.entries({ ...Object.fromEntries(Object.entries(NS).map(([k, v]) => [`xmlns:${k}`, v])), ...extra }).map(([k, v]) => `${k}="${xesc(v)}"`).join(' ');
const pad = (n, w = 2) => String(n).padStart(w, '0');
const odfSheet = (n) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${String(n).replace(/'/g, "''")}'`);
const cellAddr = (sheet, r, c) => `${odfSheet(sheet)}.${colName(c)}${r + 1}`;
const rangeAddr = (sheet, g) => (g.r1 === g.r2 && g.c1 === g.c2 ? cellAddr(sheet, g.r1, g.c1) : `${cellAddr(sheet, g.r1, g.c1)}:${cellAddr(sheet, g.r2, g.c2)}`);
const OF = (f) => toOpenFormula(f).replace(/^of:=/, '');
const MIMEX = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', emf: 'image/x-emf', wmf: 'image/x-wmf', tif: 'image/tiff', tiff: 'image/tiff' };

function b64ToBytes(b64) {
  if (typeof atob === 'function') { const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; }
  return new Uint8Array(Buffer.from(b64, 'base64')); // eslint-disable-line no-undef
}
function bytesToB64(u) {
  if (typeof btoa === 'function') { let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }
  return Buffer.from(u).toString('base64'); // eslint-disable-line no-undef
}
const dataUrl = (src) => { const m = /^data:([^;,]+)(?:;base64)?,(.*)$/s.exec(src || ''); return m ? { mime: m[1], bytes: b64ToBytes(m[2]) } : null; };

/* ---------------- dates ---------------- */
function isoOf(v, d1904) {
  const p = partsFromSerial(v, d1904);
  return `${pad(Math.max(0, p.y), 4)}-${pad(p.m)}-${pad(Math.max(1, p.d))}T${pad(p.h)}:${pad(p.mi)}:${pad(p.s)}${p.ms ? `.${pad(p.ms, 3)}` : ''}`;
}
function durationOf(v) {
  const neg = v < 0; let ms = Math.round(Math.abs(v) * 86400000);
  const h = Math.floor(ms / 3600000); ms -= h * 3600000; const m = Math.floor(ms / 60000); ms -= m * 60000; const s = ms / 1000;
  return `${neg ? '-' : ''}PT${pad(h)}H${pad(m)}M${s < 10 ? '0' : ''}${s.toFixed(ms % 1000 ? 3 : 0)}S`;
}
function serialOfIso(s, d1904) {
  const m = /^(-?\d{4,})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d)(?:\.(\d+))?)?)?/.exec(s || ''); if (!m) return null;
  return serialFromParts(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0), m[7] ? Math.round(+`0.${m[7]}` * 1000) : 0, d1904);
}
function fractionOfDuration(s) {
  const m = /^(-)?P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?)?$/.exec(String(s || '').trim()); if (!m) return null;
  const v = ((+m[2] || 0) * 86400 + (+m[3] || 0) * 3600 + (+m[4] || 0) * 60 + (+m[5] || 0)) / 86400;
  return m[1] ? -v : v;
}

/* ---------------- text paragraphs ---------------- */
function spaces(t) {
  return xesc(t).replace(/\t/g, '<text:tab/>').replace(/ {2,}/g, (m) => ` <text:s${m.length - 1 > 1 ? ` text:c="${m.length - 1}"` : ''}/>`).replace(/^ /, '<text:s/>');
}
const paras = (str) => String(str).split(/\r?\n/).map((l) => `<text:p>${spaces(l)}</text:p>`).join('');

/* ---------------- writer: style registry ---------------- */
class Styles {
  constructor(book, def) {
    this.book = book; this.def = def;
    this.cells = new Map(); this.cellKind = new Map(); this.data = new Map(); this.auto = []; this.dataXml = [];
    this.cols = new Map(); this.rows = new Map(); this.tables = new Map(); this.texts = new Map();
    this.fonts = new Set([def.font]); this.cf = new Map(); this.cfXml = []; this.masters = new Map();
  }
  font(s) { if (s && s.font) this.fonts.add(s.font); }
  dataStyle(nf) {
    if (nf == null || nf === 'General' || nf === 0) return { name: null, kind: 'number' };
    const code = String(nf);
    let x = this.data.get(code);
    if (!x) {
      const name = `N${this.data.size + 100}`;
      const d = dataStyleXml(nf, name, this.book.loc);
      x = d.general ? { name: null, kind: 'number' } : { name, kind: d.kind };
      if (!d.general) this.dataXml.push(d.xml);
      this.data.set(code, x);
    }
    return x;
  }
  /** named style for a conditional format (style:map / calcext need common styles) */
  cfStyle(st) {
    const key = JSON.stringify(st || {});
    let n = this.cf.get(key);
    if (!n) { n = `HWcf${this.cf.size + 1}`; this.cf.set(key, n); this.font(st); const d = st && st.nf ? this.dataStyle(st.nf) : null; this.cfXml.push(`<style:style style:name="${n}" style:family="table-cell" style:parent-style-name="Default"${d && d.name ? ` style:data-style-name="${d.name}"` : ''}>${cellStyleProps(st || {}, {})}</style:style>`); }
    return n;
  }
  cell(s, maps) {
    if (!s && !maps) return null;
    const { _xf, ...st } = s || {}; // eslint-disable-line no-unused-vars
    const key = JSON.stringify(st) + (maps || '');
    let n = this.cells.get(key);
    if (!n) {
      n = `ce${this.cells.size + 1}`; this.cells.set(key, n);
      this.font(st);
      const d = this.dataStyle(st.nf);
      this.cellKind.set(n, d.kind);
      this.auto.push(`<style:style style:name="${n}" style:family="table-cell" style:parent-style-name="Default"${d.name ? ` style:data-style-name="${d.name}"` : ''}>${cellStyleProps(st, this.def)}${maps || ''}</style:style>`);
    }
    return n;
  }
  col(w) {
    const key = Math.round(w);
    let n = this.cols.get(key);
    if (!n) { n = `co${this.cols.size + 1}`; this.cols.set(key, n); this.auto.push(`<style:style style:name="${n}" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="${inch(w)}"/></style:style>`); }
    return n;
  }
  row(h, custom) {
    const key = `${Math.round(h)}|${custom ? 1 : 0}`;
    let n = this.rows.get(key);
    if (!n) { n = `ro${this.rows.size + 1}`; this.rows.set(key, n); this.auto.push(`<style:style style:name="${n}" style:family="table-row"><style:table-row-properties style:row-height="${inch(h)}" fo:break-before="auto" style:use-optimal-row-height="${custom ? 'false' : 'true'}"/></style:style>`); }
    return n;
  }
  table(sh, master) {
    const key = `${sh.hidden ? 0 : 1}|${sh.color || ''}|${master}`;
    let n = this.tables.get(key);
    if (!n) { n = `ta${this.tables.size + 1}`; this.tables.set(key, n); this.auto.push(`<style:style style:name="${n}" style:family="table" style:master-page-name="${xesc(master)}"><style:table-properties table:display="${sh.hidden ? 'false' : 'true'}" style:writing-mode="lr-tb"${sh.color ? ` tableooo:tab-color="${sh.color.toLowerCase()}"` : ''}/></style:style>`); }
    return n;
  }
  text(run) {
    const { t, ...st } = run; // eslint-disable-line no-unused-vars
    const key = JSON.stringify(st);
    let n = this.texts.get(key);
    if (!n) { n = `T${this.texts.size + 1}`; this.texts.set(key, n); this.font(st); this.auto.push(`<style:style style:name="${n}" style:family="text">${textProps({ ...st, vert: st.va }, {})}</style:style>`); }
    return n;
  }
  /** master page per print setup: Default = portrait */
  master(p) {
    const land = p && p.orientation === 'landscape';
    const key = `${land ? 1 : 0}|${p && p.gridlines ? 1 : 0}|${p && p.headings ? 1 : 0}`;
    if (key === '0|0|0' && this.hasDefault !== false) return 'Default';
    let n = this.masters.get(key);
    if (!n) { let i = this.masters.size; do { n = `HWmp${++i}`; } while (this.taken && this.taken.has(n)); this.masters.set(key, n); }
    return n;
  }
}

/* ---------------- writer: conditional formats + validation ---------------- */
const CF_OP = { gt: '>', lt: '<', ge: '>=', le: '<=', eq: '=', ne: '!=' };
function styleMapOf(rule, base) {
  if (rule.type === 'cell') {
    const a = OF(rule.f1 || '0'); const b = OF(rule.f2 || '0');
    const cond = rule.op === 'between' ? `cell-content-is-between(${a},${b})` : rule.op === 'notBetween' ? `cell-content-is-not-between(${a},${b})` : `cell-content()${CF_OP[rule.op] || '='}${a}`;
    return { cond, base };
  }
  if (rule.type === 'formula') return { cond: `is-true-formula(${OF(rule.formula || 'FALSE')})`, base };
  return null;
}
const STOP_T = { min: 'minimum', max: 'maximum', num: 'number', percent: 'percent', percentile: 'percentile', formula: 'formula', autoMin: 'auto-minimum', autoMax: 'auto-maximum' };
function calcextRule(rule, st, base) {
  const ap = rule.style ? ` calcext:apply-style-name="${st.cfStyle(rule.style)}"` : ` calcext:apply-style-name="${st.cfStyle({})}"`;
  const cnd = (v) => `<calcext:condition${ap} calcext:value="${xesc(v)}" calcext:base-cell-address="${xesc(base)}"/>`;
  const qt = (t) => `"${String(t || '').replace(/"/g, '""')}"`;
  const tl = base.replace(/^.*\./, '');
  switch (rule.type) {
    case 'cell': { const a = OF(rule.f1 || '0'); const b = OF(rule.f2 || '0'); return cnd(rule.op === 'between' ? `between(${a},${b})` : rule.op === 'notBetween' ? `not-between(${a},${b})` : `${CF_OP[rule.op] || '='}${a}`); }
    case 'formula': return cnd(`formula-is(${OF(rule.formula || 'FALSE')})`);
    case 'text': return cnd(`${({ contains: 'contains-text', notContains: 'not-contains-text', begins: 'begins-with', ends: 'ends-with' })[rule.op || 'contains'] || 'contains-text'}(${qt(rule.text)})`);
    case 'dup': return cnd('duplicate');
    case 'unique': return cnd('unique');
    case 'top': case 'bottom': return cnd(`${rule.type}-${rule.percent ? 'percent' : 'elements'}(${rule.rank || 10})`);
    case 'above': return cnd('above-average');
    case 'below': return cnd('below-average');
    case 'error': return cnd('is-error');
    case 'blank': return cnd(`formula-is(LEN(TRIM([.${tl}]))=0)`);
    case 'notBlank': return cnd(`formula-is(LEN(TRIM([.${tl}]))>0)`);
    case 'scale': {
      const cols = rule.colors && rule.colors.length ? rule.colors : ['#F8696B', '#FFEB84', '#63BE7B'];
      const stops = rule.stops && rule.stops.length === cols.length ? rule.stops : cols.length === 3 ? [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }];
      return `<calcext:color-scale>${stops.map((s, i) => `<calcext:color-scale-entry calcext:value="${xesc(s.val ?? 0)}" calcext:type="${STOP_T[s.type] || 'number'}" calcext:color="${String(cols[i]).toLowerCase()}"/>`).join('')}</calcext:color-scale>`;
    }
    case 'bar': {
      const stops = rule.stops && rule.stops.length === 2 ? rule.stops : [{ type: 'autoMin' }, { type: 'autoMax' }];
      return `<calcext:data-bar calcext:max-length="100" calcext:negative-color="#ff0000" calcext:positive-color="${String(rule.color || '#638EC6').toLowerCase()}" calcext:axis-color="#000000">${stops.map((s) => `<calcext:formatting-entry calcext:value="${xesc(s.val ?? 0)}" calcext:type="${STOP_T[s.type] || 'number'}"/>`).join('')}</calcext:data-bar>`;
    }
    case 'icons': {
      const n = +String(rule.set || '3TrafficLights1')[0] || 3;
      const stops = rule.stops && rule.stops.length === n ? rule.stops : Array.from({ length: n }, (_, i) => ({ type: 'percent', val: String(Math.round((i * 100) / n)) }));
      return `<calcext:icon-set calcext:icon-set-type="${xesc(rule.set || '3TrafficLights1')}"${rule.showValue === false ? ' calcext:show-value="false"' : ''}>${stops.map((s) => `<calcext:formatting-entry calcext:value="${xesc(s.val ?? 0)}" calcext:type="${STOP_T[s.type] || 'percent'}"/>`).join('')}</calcext:icon-set>`;
    }
    default: return '';
  }
}
const DV_OP = { equal: '=', notEqual: '!=', greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=' };
function dvCondition(dv) {
  const F = (f) => { const s = String(f ?? '').trim(); return s === '' ? '0' : OF(s); };
  const cmp = (fn) => (dv.op === 'between' ? `${fn}-is-between(${F(dv.f1)};${F(dv.f2)})` : dv.op === 'notBetween' ? `${fn}-is-not-between(${F(dv.f1)};${F(dv.f2)})` : `${fn}()${DV_OP[dv.op] || '='}${F(dv.f1)}`);
  switch (dv.type) {
    case 'whole': return `of:cell-content-is-whole-number() and ${cmp('cell-content')}`;
    case 'decimal': return `of:cell-content-is-decimal-number() and ${cmp('cell-content')}`;
    case 'date': return `of:cell-content-is-date() and ${cmp('cell-content')}`;
    case 'time': return `of:cell-content-is-time() and ${cmp('cell-content')}`;
    case 'textLength': return `of:${cmp('cell-content-text-length')}`;
    case 'custom': return `of:is-true-formula(${F(dv.f1)})`;
    case 'list': {
      const f = String(dv.f1 || '').trim();
      const lit = /^"(.*)"$/s.exec(f);
      if (lit) return `of:cell-content-is-in-list(${lit[1].split(',').map((x) => `"${x.trim().replace(/"/g, '""')}"`).join(';')})`;
      return `of:cell-content-is-in-list(${F(f)})`;
    }
    default: return null;
  }
}

/* ---------------- writer ---------------- */
function bounds(sh) {
  let r2 = -1; let c2 = -1;
  const see = (r, c) => { if (r > r2) r2 = r; if (c > c2) c2 = c; };
  for (const k of sh.cells.keys()) see(R(k), C(k));
  for (const k of sh.spill.keys()) see(R(k), C(k));
  for (const m of sh.merges) see(m.r2, m.c2);
  return { r2, c2 };
}
/** clip a (maybe whole-column) range to something writable */
function clip(g, b) {
  const r2 = Math.min(g.r2, Math.max(b.r2, g.r1 + 199), g.r1 + 9999);
  const c2 = Math.min(g.c2, Math.max(b.c2, g.c1 + 25), g.c1 + 255);
  return { r1: g.r1, c1: g.c1, r2: Math.max(g.r1, r2), c2: Math.max(g.c1, c2) };
}

function cellValueXml(book, sh, r, c, v, kind, hasF) {
  if (v == null || v === '') return hasF ? { a: ' office:value-type="string" office:string-value="" calcext:value-type="string"', p: '' } : { a: '', p: '' };
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return { a: ' office:value-type="string" office:string-value="" calcext:value-type="error"', p: '<text:p>#NUM!</text:p>' };
    const txt = book.text(sh, r, c);
    if (kind === 'date') return { a: ` office:value-type="date" office:date-value="${isoOf(v, book.date1904)}" calcext:value-type="date"`, p: paras(txt) };
    if (kind === 'time') return { a: ` office:value-type="time" office:time-value="${durationOf(v)}" calcext:value-type="time"`, p: paras(txt) };
    if (kind === 'percentage') return { a: ` office:value-type="percentage" office:value="${v}" calcext:value-type="percentage"`, p: paras(txt) };
    return { a: ` office:value-type="float" office:value="${v}" calcext:value-type="float"`, p: paras(txt) };
  }
  if (typeof v === 'boolean') return { a: ` office:value-type="boolean" office:boolean-value="${v}" calcext:value-type="boolean"`, p: `<text:p>${v ? 'TRUE' : 'FALSE'}</text:p>` };
  if (isErr(v)) return { a: ' office:value-type="string" office:string-value="" calcext:value-type="error"', p: `<text:p>${xesc(v.code === '#CIRC!' ? '#VALUE!' : v.code)}</text:p>` };
  return { a: ' office:value-type="string" calcext:value-type="string"', p: null };
}
function richParas(st, runs) {
  const out = [[]];
  for (const run of runs) {
    const lines = String(run.t).split(/\r?\n/);
    lines.forEach((l, i) => { if (i) out.push([]); if (l) out[out.length - 1].push({ ...run, t: l }); });
  }
  return out.map((p) => `<text:p>${p.map((run) => { const plain = !Object.keys(run).some((k) => k !== 't' && run[k]); return plain ? spaces(run.t) : `<text:span text:style-name="${st.text(run)}">${spaces(run.t)}</text:span>`; }).join('')}</text:p>`).join('');
}

function frameXml(sh, a, inner, name, z) {
  const colX = (c) => { let x = 0; for (let i = 0; i < c; i++) x += sh.hiddenCols.has(i) ? 0 : sh.colW.get(i) ?? sh.defW; return x; };
  const rowY = (r) => { let y = 0; for (let i = 0; i < r; i++) y += sh.hiddenRows.has(i) ? 0 : sh.rowH.get(i) ?? sh.defH; return y; };
  const x = colX(a.c || 0) + (a.dx || 0); const y = rowY(a.r || 0) + (a.dy || 0);
  return `<draw:frame draw:z-index="${z}" draw:name="${xesc(name)}" svg:width="${inch(a.w || 200)}" svg:height="${inch(a.h || 150)}" svg:x="${inch(x)}" svg:y="${inch(y)}">${inner}</draw:frame>`;
}

function sheetXml(book, sh, si, st, out) {
  const b = bounds(sh);
  const name = sh.name;
  const extra = new Map(); // key → { maps: [], dv: name }
  const ex = (r, c) => { const k = K(r, c); let e = extra.get(k); if (!e) { e = {}; extra.set(k, e); } return e; };
  // conditional formats: style:map for Excel/ODF 1.2 consumers, calcext for LibreOffice
  const calc = [];
  sh.cf.forEach((rule) => {
    if (!rule.ranges || !rule.ranges.length || rule.type === 'raw') return;
    const g0 = rule.ranges[0];
    const base = cellAddr(name, g0.r1, g0.c1);
    const inner = calcextRule(rule, st, base);
    if (inner) calc.push(`<calcext:conditional-format calcext:target-range-address="${xesc(rule.ranges.map((g) => rangeAddr(name, g)).join(' '))}">${inner}</calcext:conditional-format>`);
    const m = styleMapOf(rule, base);
    if (!m) return;
    const mx = `<style:map style:condition="${xesc(m.cond)}" style:apply-style-name="${st.cfStyle(rule.style || {})}" style:base-cell-address="${xesc(base)}"/>`;
    for (const g of rule.ranges) { const q = clip(g, b); for (let r = q.r1; r <= q.r2; r++) for (let c = q.c1; c <= q.c2; c++) { const e = ex(r, c); (e.maps || (e.maps = [])).push(mx); } }
  });
  sh.dv.forEach((dv) => {
    const cond = dvCondition(dv);
    const vn = `val${out.validations.length + 1}`;
    const g0 = dv.ranges && dv.ranges[0]; if (!g0) return;
    const msg = (tag, m, extraA = '') => `<table:${tag} table:title="${xesc(m.title || '')}" table:display="${m.show === false ? 'false' : 'true'}"${extraA}>${m.text ? paras(m.text) : ''}</table:${tag}>`;
    out.validations.push(`<table:content-validation table:name="${vn}"${cond ? ` table:condition="${xesc(cond)}"` : ''} table:allow-empty-cell="${dv.allowBlank === false ? 'false' : 'true'}" table:display-list="${dv.type === 'list' ? (dv.dropdown === false ? 'none' : 'unsorted') : 'unsorted'}" table:base-cell-address="${xesc(cellAddr(name, g0.r1, g0.c1))}">${dv.prompt ? msg('help-message', { ...dv.prompt, show: true }) : ''}${msg('error-message', dv.error || {}, ` table:message-type="${({ stop: 'stop', warning: 'warning', information: 'information' })[(dv.error || {}).style] || 'stop'}"`)}</table:content-validation>`);
    for (const g of dv.ranges) { const q = clip(g, b); for (let r = q.r1; r <= q.r2; r++) for (let c = q.c1; c <= q.c2; c++) ex(r, c).dv = vn; }
  });
  // merges
  const span = new Map(); const covered = new Set();
  for (const m of sh.merges) { span.set(K(m.r1, m.c1), m); for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(K(r, c)); }
  // what to emit per row
  const rows = new Map();
  const add = (k) => { const r = R(k); let s = rows.get(r); if (!s) { s = new Set(); rows.set(r, s); } s.add(C(k)); };
  for (const k of sh.cells.keys()) add(k);
  for (const k of sh.spill.keys()) add(k);
  for (const k of extra.keys()) add(k);
  for (const k of span.keys()) add(k);
  for (const k of covered) add(k);
  let maxR = b.r2; let maxC = b.c2;
  for (const k of extra.keys()) { maxR = Math.max(maxR, R(k)); maxC = Math.max(maxC, C(k)); }
  for (const r of [...sh.rowH.keys(), ...sh.hiddenRows, ...sh.filterHidden, ...sh.outlineRows.keys()]) if (r < 100000) maxR = Math.max(maxR, r);
  for (const c of [...sh.colW.keys(), ...sh.hiddenCols, ...sh.outlineCols.keys()]) if (c < 1024) maxC = Math.max(maxC, c);
  maxC = Math.max(maxC, 0);
  // columns (with outline groups)
  const colUnits = [];
  for (let c = 0; c <= maxC; c++) {
    const w = sh.colW.get(c) ?? sh.defW; const hid = sh.hiddenCols.has(c); const lvl = sh.outlineCols.get(c) || 0;
    const key = `${Math.round(w)}|${hid}|${lvl}`;
    const last = colUnits[colUnits.length - 1];
    if (last && last.key === key) last.n++; else colUnits.push({ key, n: 1, lvl, xml: (n) => `<table:table-column table:style-name="${st.col(w)}"${n > 1 ? ` table:number-columns-repeated="${n}"` : ''}${hid ? ' table:visibility="collapse"' : ''} table:default-cell-style-name="Default"/>` });
  }
  const grouped = (units, tag) => {
    let x = ''; let lvl = 0;
    for (const u of units) { while (lvl < u.lvl) { x += `<table:table-${tag}-group>`; lvl++; } while (lvl > u.lvl) { x += `</table:table-${tag}-group>`; lvl--; } x += u.xml(u.n); }
    while (lvl > 0) { x += `</table:table-${tag}-group>`; lvl--; }
    return x;
  };
  // rows
  const rowUnits = [];
  for (let r = 0; r <= maxR; r++) {
    const custom = sh.rowH.has(r); const h = sh.rowH.get(r) ?? sh.defH;
    const vis = sh.hiddenRows.has(r) ? 'collapse' : sh.filterHidden.has(r) ? 'filter' : '';
    const lvl = sh.outlineRows.get(r) || 0;
    const rs = st.row(h, custom);
    const cols = rows.get(r);
    if (!cols) {
      const key = `${rs}|${vis}|${lvl}`;
      const last = rowUnits[rowUnits.length - 1];
      if (last && last.key === key) { last.n++; continue; }
      rowUnits.push({ key, n: 1, lvl, xml: (n) => `<table:table-row table:style-name="${rs}"${n > 1 ? ` table:number-rows-repeated="${n}"` : ''}${vis ? ` table:visibility="${vis}"` : ''}><table:table-cell table:number-columns-repeated="${maxC + 1}"/></table:table-row>` });
      continue;
    }
    let x = ''; let c0 = 0;
    for (const c of [...cols].sort((p, q) => p - q)) {
      if (c > c0) x += `<table:table-cell${c - c0 > 1 ? ` table:number-columns-repeated="${c - c0}"` : ''}/>`;
      x += cellXml(book, sh, r, c, st, extra.get(K(r, c)), span.get(K(r, c)), covered.has(K(r, c)), out);
      c0 = c + 1;
    }
    if (c0 <= maxC) x += `<table:table-cell${maxC + 1 - c0 > 1 ? ` table:number-columns-repeated="${maxC + 1 - c0}"` : ''}/>`;
    rowUnits.push({ key: null, n: 1, lvl, xml: () => `<table:table-row table:style-name="${rs}"${vis ? ` table:visibility="${vis}"` : ''}>${x}</table:table-row>` });
  }
  // shapes: images, charts, kept frames
  let shapes = ''; let z = 0;
  for (const im of sh.images) {
    const d = dataUrl(im.src);
    let inner;
    if (out.flat) { if (!d) continue; inner = `<draw:image><office:binary-data>${bytesToB64(d.bytes)}</office:binary-data></draw:image>`; }
    else {
      let href = im._ods && im._ods.href && im._ods.src === im.src ? im._ods.href : null;
      if (href) out.keepPictures.add(href);
      else {
        if (!d) continue;
        const ext = (Object.entries(MIMEX).find(([, m]) => m === d.mime) || ['png'])[0];
        href = `Pictures/hw${out.pictures.length + 1}${uid('i').slice(-6)}.${ext}`;
        out.pictures.push({ path: href, bytes: d.bytes, mime: d.mime });
      }
      inner = `<draw:image xlink:href="${xesc(href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`;
    }
    shapes += frameXml(sh, im.anchor, inner, im.name || `Image ${z + 1}`, ++z);
  }
  for (const ch of sh.charts) {
    const a = ch.anchor || { r: 0, c: 0, dx: 0, dy: 0, w: 480, h: 288 };
    const unchanged = ch._ods && ch._ods.sig === JSON.stringify(chartSigOds(ch)) && !out.flat;
    let dir = unchanged ? ch._ods.dir : null;
    let ranges = ch._ods && ch._ods.ranges;
    if (!dir || out.flat) {
      let data;
      try { data = resolveChartData(ch, out.getValues); } catch { data = { categories: [], series: [] }; }
      const parts = chartParts(ch, data, name, { w: a.w || 480, h: a.h || 288 });
      ranges = parts.ranges.join(' ');
      if (out.flat) { shapes += frameXml(sh, a, `<draw:object${ranges ? ` draw:notify-on-update-of-ranges="${xesc(ranges)}"` : ''}>${chartInlineXml(parts)}</draw:object>`, ch.title || `Chart ${z + 1}`, ++z); continue; }
      dir = ch._ods && ch._ods.dir ? ch._ods.dir : out.newObjectDir();
      out.objects.push({ dir, xml: chartContentXml(parts), replaced: !!(ch._ods && ch._ods.dir) });
    } else out.keepObjects.add(dir);
    shapes += frameXml(sh, a, `<draw:object${ranges ? ` draw:notify-on-update-of-ranges="${xesc(ranges)}"` : ''} xlink:href="${xesc(dir)}/" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`, ch.title || `Chart ${z + 1}`, ++z);
  }
  for (const raw of (sh._ods && sh._ods.frames) || []) shapes += raw;
  const master = sh._ods && sh._ods.master && sh._ods.printSig === printSig(sh) ? sh._ods.master : st.master(sh.print);
  const pa = sh.print && sh.print.area ? (() => { const g = /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i.exec(sh.print.area); if (!g) return ''; const ci = (s) => s.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return ` table:print-ranges="${xesc(rangeAddr(name, { r1: +g[2] - 1, c1: ci(g[1]), r2: +g[4] - 1, c2: ci(g[3]) }))}"`; })() : '';
  // sheet-scoped names
  const local = [...book.names.values()].filter((d) => d.scope === sh && !d.raw).map((d) => namedXml(d, name)).join('');
  const prot = sh.protect || sh.protection ? ' table:protected="true"' : '';
  return `<table:table table:name="${xesc(name)}" table:style-name="${st.table(sh, master)}"${pa}${prot}>${shapes ? `<table:shapes>${shapes}</table:shapes>` : ''}${grouped(colUnits, 'column')}${grouped(rowUnits, 'row')}${local ? `<table:named-expressions>${local}</table:named-expressions>` : ''}${calc.length ? `<calcext:conditional-formats>${calc.join('')}</calcext:conditional-formats>` : ''}</table:table>`;
}
const printSig = (sh) => JSON.stringify([sh.print && sh.print.orientation, !!(sh.print && sh.print.gridlines), !!(sh.print && sh.print.headings)]);
export const chartSigOds = (ch) => { const { id, _orig, _sig, _xl, _dirty, _ods, anchor, ...rest } = ch; return rest; }; // eslint-disable-line no-unused-vars

function namedXml(d, sheetName) {
  const ref = String(d.ref || '').trim();
  const of = OF(ref);
  const m = /^\[([^\]]+)\]$/.exec(of);
  if (m) { const addr = m[1].replace(/^\./, `$${odfSheet(sheetName)}.`); return `<table:named-range table:name="${xesc(d.name)}" table:base-cell-address="$${xesc(odfSheet(sheetName))}.$A$1" table:cell-range-address="${xesc(addr)}"/>`; }
  return `<table:named-expression table:name="${xesc(d.name)}" table:base-cell-address="$${xesc(odfSheet(sheetName))}.$A$1" table:expression="of:=${xesc(of)}"/>`;
}

function cellXml(book, sh, r, c, st, e, mg, isCovered, out) {
  const k = K(r, c);
  const cell = sh.cells.get(k); const sp = !cell ? sh.spill.get(k) : null;
  const v = cell ? cell.v : sp ? sp.v : null;
  const sname = st.cell(cell && cell.s, e && e.maps ? e.maps.join('') : null);
  const kind = sname ? st.cellKind.get(sname) : 'number';
  const tag = isCovered ? 'table:covered-table-cell' : 'table:table-cell';
  let a = '';
  if (sname) a += ` table:style-name="${sname}"`;
  if (e && e.dv) a += ` table:content-validation-name="${e.dv}"`;
  if (mg) a += ` table:number-columns-spanned="${mg.c2 - mg.c1 + 1}" table:number-rows-spanned="${mg.r2 - mg.r1 + 1}"`;
  const f = cell && cell.f ? cell.f : null;
  if (f) a += ` table:formula="${xesc(toOpenFormula(f))}"`;
  const val = cellValueXml(book, sh, r, c, v, kind, !!f);
  a += val.a;
  let inner = '';
  if (cell && cell.note) inner += `<office:annotation${cell.noteShown ? ' office:display="true"' : ''}>${cell.noteAuthor ? `<dc:creator>${xesc(cell.noteAuthor)}</dc:creator>` : ''}${paras(cell.note)}</office:annotation>`;
  let p = val.p;
  if (p == null) p = cell && cell.rich && cell.rich.map((x) => x.t).join('') === v ? richParas(st, cell.rich) : paras(String(v));
  const link = cell && (cell.link || cell.link_);
  if (link && p) p = p.replace(/^<text:p>([\s\S]*?)<\/text:p>/, (m0, t) => `<text:p><text:a xlink:href="${xesc(link.startsWith('#') ? `#${link.slice(1).replace(/!/, '.')}` : link)}" xlink:type="simple">${t}</text:a></text:p>`);
  inner += p || '';
  const keep = sh._ods && sh._ods.cellFrames && sh._ods.cellFrames.get(k);
  if (keep) inner += keep.join('');
  return inner ? `<${tag}${a}>${inner}</${tag}>` : `<${tag}${a}/>`;
}

function settingsXml(book, confRaw) {
  const item = (n, t, v) => `<config:config-item config:name="${n}" config:type="${t}">${xesc(v)}</config:config-item>`;
  const tables = book.sheets.map((sh) => {
    const fr = sh.freeze || { rows: 0, cols: 0 }; const fz = fr.rows > 0 || fr.cols > 0;
    const sel = sh.sel || { r: 0, c: 0 };
    return `<config:config-item-map-entry config:name="${xesc(sh.name)}">${item('CursorPositionX', 'int', sel.c || 0)}${item('CursorPositionY', 'int', sel.r || 0)}${item('HorizontalSplitMode', 'short', fr.cols > 0 ? 2 : 0)}${item('VerticalSplitMode', 'short', fr.rows > 0 ? 2 : 0)}${item('HorizontalSplitPosition', 'int', fr.cols || 0)}${item('VerticalSplitPosition', 'int', fr.rows || 0)}${item('ActiveSplitRange', 'short', fz ? (fr.rows > 0 ? 2 : 3) : 2)}${item('PositionLeft', 'int', 0)}${item('PositionRight', 'int', fr.cols || 0)}${item('PositionTop', 'int', 0)}${item('PositionBottom', 'int', fr.rows || 0)}${item('ZoomType', 'short', 0)}${item('ZoomValue', 'int', Math.round(sh.zoom || 100))}${item('PageViewZoomValue', 'int', 60)}${item('ShowGrid', 'boolean', sh.gridlines !== false)}</config:config-item-map-entry>`;
  }).join('');
  const act = book.sheets[book.active] || book.sheets[0];
  const view = `<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>${item('ViewId', 'string', 'view1')}<config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named>${item('ActiveTable', 'string', act ? act.name : '')}${item('ShowGrid', 'boolean', !act || act.gridlines !== false)}${item('HasColumnRowHeaders', 'boolean', !act || act.headings !== false)}${item('HasSheetTabs', 'boolean', true)}</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>`;
  return { view, conf: confRaw || '' };
}

function masterPagesXml(st) {
  const lay = (name, land, grid, head) => `<style:page-layout style:name="${name}"><style:page-layout-properties fo:page-width="${land ? '29.7cm' : '21cm'}" fo:page-height="${land ? '21cm' : '29.7cm'}" style:print-orientation="${land ? 'landscape' : 'portrait'}" fo:margin-top="1.5cm" fo:margin-bottom="1.5cm" fo:margin-left="1.2cm" fo:margin-right="1.2cm" style:print="annotations charts drawings objects zero-values${grid ? ' grid' : ''}${head ? ' headers' : ''}" style:writing-mode="lr-tb"/></style:page-layout>`;
  const layouts = []; const masters = [];
  for (const [key, n] of st.masters) { const [l, g, h] = key.split('|').map((x) => x === '1'); layouts.push(lay(`${n}L`, l, g, h)); masters.push(`<style:master-page style:name="${n}" style:page-layout-name="${n}L"/>`); }
  return { layouts: layouts.join(''), masters: masters.join(''), lay };
}

function stylesDoc(st, def, mp, flatParts) {
  const fonts = [...st.fonts].filter(Boolean).map(fontFace).join('');
  const cellDefault = `<style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="1.25cm"/><style:text-properties style:font-name="${xesc(def.font)}" fo:font-size="${def.size}pt" style:font-name-asian="${xesc(def.font)}" style:font-size-asian="${def.size}pt" style:font-name-complex="${xesc(def.font)}" style:font-size-complex="${def.size}pt" fo:language="${xesc(flatParts.lang || 'en')}" fo:country="${xesc(flatParts.country || 'US')}"/></style:default-style>`;
  const officeStyles = `${cellDefault}<style:style style:name="Default" style:family="table-cell"/>${st.cfXml.join('')}`;
  const auto = `${mp.lay('pm1', false, false, false)}${mp.layouts}`;
  const master = `<style:master-page style:name="Default" style:page-layout-name="pm1"/>${mp.masters}`;
  return { fonts, officeStyles, auto, master };
}

/** inject our CF styles and master pages into a kept styles part (string) */
function patchStyles(xml, st, mp, { cf = true, pages = true, fonts = true } = {}) {
  let x = xml.replace(/<style:style style:name="HWcf\d+"[\s\S]*?<\/style:style>/g, '');
  const ins = (closeTag, add, openTag) => {
    if (!add) return;
    if (x.includes(closeTag)) x = x.replace(closeTag, `${add}${closeTag}`);
    else if (openTag && new RegExp(`<${openTag}\\s*/>`).test(x)) x = x.replace(new RegExp(`<${openTag}\\s*/>`), `<${openTag}>${add}</${openTag}>`);
  };
  if (cf) {
    ins('</office:styles>', st.cfXml.join(''), 'office:styles');
    if (!/<office:styles[\s>]/.test(x) && st.cfXml.length) x = x.replace(/(<office:automatic-styles|<office:master-styles|<office:body)/, `<office:styles>${st.cfXml.join('')}</office:styles>$1`);
  }
  if (pages) { ins('</office:automatic-styles>', mp.layouts, 'office:automatic-styles'); ins('</office:master-styles>', mp.masters, 'office:master-styles'); }
  if (!fonts) return x;
  // fonts the CF styles use
  const need = [...st.fonts].filter((f) => f && !x.includes(`style:name="${xesc(f)}"`)).map(fontFace).join('');
  ins('</office:font-face-decls>', need, 'office:font-face-decls');
  return x;
}

/**
 * book → ods / ots / fods bytes. opts: { ext, getValues (chart data; default: from the book) }
 */
export async function writeOds(book, { ext = 'ods', getValues = null } = {}) {
  const flat = ext === 'fods';
  const P = book._ods && book._ods.flat === flat ? book._ods : null;
  const def = book.props.defaultFont || { font: book.settings.font || 'Calibri', size: book.settings.fontSize || 11 };
  const st = new Styles(book, def);
  if (P && P.masters && !P.masters.has('Default')) st.hasDefault = false;
  if (P && P.masters) st.taken = P.masters;
  const used = new Set(P ? P.files.filter((f) => /^Object \d+\/?$/.test(f.path.split('/')[0] + '/')).map((f) => f.path.split('/')[0]) : []);
  let objN = 0;
  const out = {
    flat, validations: [], pictures: [], objects: [], keepPictures: new Set(), keepObjects: new Set(),
    getValues: getValues || ((ref) => valuesOf(book, ref)),
    newObjectDir() { let n; do { n = `Object ${++objN}`; } while (used.has(n)); used.add(n); return n; },
  };
  const tables = book.sheets.map((sh, i) => sheetXml(book, sh, i, st, out)).join('');
  const globalNames = [...book.names.values()].filter((d) => !d.scope && !d.raw).map((d) => namedXml(d, (book.sheets[0] || { name: 'Sheet1' }).name)).join('');
  const dbr = book.sheets.map((sh, i) => (sh.filter && sh.filter.range ? `<table:database-range table:name="__Anonymous_Sheet_DB__${i}" table:target-range-address="${xesc(rangeAddr(sh.name, sh.filter.range))}" table:display-filter-buttons="true"/>` : '')).join('');
  const calcSet = `<table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true">${book.date1904 ? '<table:null-date table:date-value="1904-01-01"/>' : ''}${book.iterate && book.iterate.on ? `<table:iteration table:status="enable" table:steps="${book.iterate.count || 100}" table:minimum-difference="${book.iterate.delta || 0.001}"/>` : ''}</table:calculation-settings>`;
  const keepBefore = P ? P.before.join('') : ''; const keepAfter = P ? P.after.join('') : '';
  const spreadsheet = `<office:spreadsheet>${keepBefore}${calcSet}${out.validations.length ? `<table:content-validations>${out.validations.join('')}</table:content-validations>` : ''}${tables}${globalNames ? `<table:named-expressions>${globalNames}</table:named-expressions>` : ''}${dbr || (P && P.dbKeep.length) ? `<table:database-ranges>${dbr}${P ? P.dbKeep.join('') : ''}</table:database-ranges>` : ''}${keepAfter}</office:spreadsheet>`;
  const mp = masterPagesXml(st);
  const lang = /^de/.test((book.loc && book.loc.lang) || '') ? { lang: 'de', country: 'DE' } : { lang: 'en', country: 'US' };
  const sd = stylesDoc(st, def, mp, lang);
  const contentFonts = [...st.fonts].filter(Boolean).filter((f) => !(P && P.contentFonts.some((x) => x.includes(`style:name="${xesc(f)}"`)))).map(fontFace).join('') + (P ? P.contentFonts.join('') : '');
  const autoStyles = `${st.dataXml.join('')}${st.auto.join('')}`;
  const sets = settingsXml(book, P && P.confRaw);
  const rootNs = nsAttrs(P ? P.ns : {});
  const meta = `<office:meta><meta:generator>Hostwerk Office</meta:generator><meta:creation-date>${new Date().toISOString().replace(/\.\d+Z$/, '')}</meta:creation-date></office:meta>`;
  if (flat) {
    let head = P ? patchStyles(P.flatStyles || '<office:styles/>', st, mp, { pages: false }) : `<office:font-face-decls>${sd.fonts}</office:font-face-decls><office:styles>${sd.officeStyles}</office:styles>`;
    const missing = [...st.fonts].filter((f) => f && !head.includes(`style:name="${xesc(f)}"`)).map(fontFace).join('');
    if (missing) head = head.includes('</office:font-face-decls>') ? head.replace('</office:font-face-decls>', `${missing}</office:font-face-decls>`) : `<office:font-face-decls>${missing}</office:font-face-decls>${head.replace(/<office:font-face-decls\s*\/>/, '')}`;
    const pageAuto = P ? `${P.flatPageLayouts || ''}${mp.layouts}` : sd.auto;
    const masterPart = P ? patchStyles(P.flatMaster || '<office:master-styles/>', st, { ...mp, layouts: '' }, { cf: false, fonts: false }) : `<office:master-styles>${sd.master}</office:master-styles>`;
    const xml = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document ${rootNs} office:version="1.2" office:mimetype="${ODS_MIME.fods}">${P && P.flatMeta ? P.flatMeta : meta}<office:settings>${sets.view}${sets.conf}</office:settings>${head}<office:automatic-styles>${pageAuto}${autoStyles}</office:automatic-styles>${masterPart}<office:body>${spreadsheet}</office:body></office:document>`;
    return new TextEncoder().encode(xml);
  }
  const content = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${rootNs} office:version="1.2"><office:scripts/><office:font-face-decls>${contentFonts}</office:font-face-decls><office:automatic-styles>${autoStyles}</office:automatic-styles><office:body>${spreadsheet}</office:body></office:document-content>`;
  const stylesXml = P && P.stylesXml != null ? patchStyles(P.stylesXml, st, mp) : `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${nsAttrs()} office:version="1.2"><office:font-face-decls>${sd.fonts}</office:font-face-decls><office:styles>${sd.officeStyles}</office:styles><office:automatic-styles>${sd.auto}</office:automatic-styles><office:master-styles>${sd.master}</office:master-styles></office:document-styles>`;
  const settings = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-settings xmlns:office="${NS.office}" xmlns:xlink="${NS.xlink}" xmlns:config="${NS.config}" xmlns:ooo="${NS.ooo}" office:version="1.2"><office:settings>${sets.view}${sets.conf}</office:settings></office:document-settings>`;
  const metaXml = P && P.metaXml ? P.metaXml : `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta xmlns:office="${NS.office}" xmlns:meta="${NS.meta}" xmlns:dc="${NS.dc}" office:version="1.2">${meta}</office:document-meta>`;
  const mime = ODS_MIME[ext] || ODS_MIME.ods;
  const zip = new JSZip();
  zip.file('mimetype', mime, { compression: 'STORE' });
  const manifest = new Map(); // path → media type
  const enc = new TextEncoder();
  const put = (path, data, type) => { zip.file(path, data); manifest.set(path, type); };
  put('content.xml', enc.encode(content), 'text/xml');
  put('styles.xml', enc.encode(stylesXml), 'text/xml');
  put('meta.xml', enc.encode(metaXml), 'text/xml');
  put('settings.xml', enc.encode(settings), 'text/xml');
  // kept parts of the original package
  const replacedObjects = new Set(out.objects.map((o) => o.dir));
  const droppedObjects = new Set();
  if (P) {
    const ourDirs = new Set(P.chartDirs);
    for (const d of ourDirs) if (!out.keepObjects.has(d) && !replacedObjects.has(d)) droppedObjects.add(d);
    for (const f of P.files) {
      const p = f.path;
      if (['mimetype', 'content.xml', 'styles.xml', 'meta.xml', 'settings.xml', 'META-INF/manifest.xml'].includes(p)) continue;
      const top = p.split('/')[0];
      if (droppedObjects.has(top)) continue;
      if (/^ObjectReplacements\//.test(p) && (droppedObjects.has(p.slice(19)) || replacedObjects.has(p.slice(19)))) continue;
      if (replacedObjects.has(top) && p === `${top}/content.xml`) continue;
      if (f.dir) { manifest.set(p, P.manifest.get(p) || ''); continue; }
      zip.file(p, await f.data());
      manifest.set(p, P.manifest.get(p) || (p.endsWith('.xml') ? 'text/xml' : MIMEX[(p.split('.').pop() || '').toLowerCase()] || ''));
    }
    for (const [p, t] of P.manifest) if (p.endsWith('/') && !droppedObjects.has(p.replace(/\/$/, '')) && p !== '/' && !manifest.has(p) && [...manifest.keys()].some((x) => x.startsWith(p))) manifest.set(p, t);
  }
  for (const pic of out.pictures) put(pic.path, pic.bytes, pic.mime);
  for (const o of out.objects) { manifest.set(`${o.dir}/`, 'application/vnd.oasis.opendocument.chart'); put(`${o.dir}/content.xml`, enc.encode(o.xml), 'text/xml'); if (!manifest.has(`${o.dir}/styles.xml`)) put(`${o.dir}/styles.xml`, enc.encode(CHART_STYLES), 'text/xml'); }
  const entries = [`<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${mime}"/>`];
  for (const [p, t] of manifest) entries.push(`<manifest:file-entry manifest:full-path="${xesc(p)}"${p.endsWith('/') ? ' manifest:version="1.2"' : ''} manifest:media-type="${xesc(t)}"/>`);
  zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">${entries.join('')}</manifest:manifest>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: mime });
}

/* Excel shows an embedded chart only when the object has a styles part */
const CHART_STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="${NS.office}" xmlns:style="${NS.style}" xmlns:draw="${NS.draw}" xmlns:svg="${NS.svg}" office:version="1.2"><office:styles><style:default-style style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#4472c4" draw:stroke="solid" svg:stroke-color="#172c51"/></style:default-style></office:styles></office:document-styles>`;
function valuesOf(book, ref) {
  const m = /^(?:(?:'((?:[^']|'')+)'|([^!'\s]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(String(ref || '').trim());
  if (!m) return null;
  const sheet = m[1] != null ? m[1].replace(/''/g, "'") : m[2];
  const sh = sheet ? book.sheetByName(sheet) : book.sheets[book.active]; if (!sh) return null;
  const ci = (s) => s.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const r1 = +m[4] - 1; const c1 = ci(m[3]); const r2 = m[6] ? +m[6] - 1 : r1; const c2 = m[5] ? ci(m[5]) : c1;
  const values = []; const texts = [];
  for (let r = r1; r <= Math.min(r2, r1 + 5000); r++) { const a = []; const tx = []; for (let c = c1; c <= Math.min(c2, c1 + 200); c++) { a.push(book.value(sh, r, c)); tx.push(book.text(sh, r, c)); } values.push(a); texts.push(tx); }
  return { values, texts };
}

/* ================================================================ reader ================================================================ */
function paraText(p, runs, ctx, base) {
  let s = '';
  const walk = (el, style) => {
    for (const x of el.c || []) {
      if (typeof x === 'string') { s += x; if (runs) runs.push({ t: x, ...style }); continue; }
      switch (x.n) {
        case 'text:s': { const n = +(x.a['text:c'] || 1); const t = ' '.repeat(n); s += t; if (runs) runs.push({ t, ...style }); break; }
        case 'text:tab': s += '\t'; if (runs) runs.push({ t: '\t', ...style }); break;
        case 'text:line-break': s += '\n'; if (runs) runs.push({ t: '\n', ...style }); break;
        case 'text:span': { const ts = x.a['text:style-name'] && ctx.textStyle(x.a['text:style-name']); walk(x, ts ? { ...style, ...ts } : style); break; }
        case 'text:a': if (!ctx.link) ctx.link = x.a['xlink:href']; walk(x, style); break;
        case 'office:annotation': case 'text:note': break;
        default: walk(x, style);
      }
    }
  };
  walk(p, base || {});
  return s;
}

function mergeRuns(runs, v) {
  const out = [];
  for (const q of runs) {
    const { t, ...rest } = q; const l = out[out.length - 1];
    if (l) { const { t: lt, ...lr } = l; if (JSON.stringify(lr) === JSON.stringify(rest)) { l.t = lt + t; continue; } }
    out.push({ t, ...rest });
  }
  if (out.map((q) => q.t).join('') !== v) return undefined;
  return out.some((q) => Object.keys(q).length > 1) ? out : undefined;
}
/** coalesce cell points into rectangles */
function rects(points) {
  const byRow = new Map();
  for (const [r, c] of points) { let a = byRow.get(r); if (!a) { a = []; byRow.set(r, a); } a.push(c); }
  const runs = [];
  for (const [r, cs] of [...byRow].sort((x, y) => x[0] - y[0])) {
    cs.sort((x, y) => x - y);
    let s = cs[0]; let p = cs[0];
    for (let i = 1; i <= cs.length; i++) { if (i < cs.length && cs[i] === p + 1) { p = cs[i]; continue; } runs.push({ r, c1: s, c2: p }); if (i < cs.length) { s = cs[i]; p = cs[i]; } }
  }
  const out = []; const open = new Map();
  for (const run of runs) {
    const key = `${run.c1}:${run.c2}`; const g = open.get(key);
    if (g && g.r2 === run.r - 1) g.r2 = run.r;
    else { const n = { r1: run.r, c1: run.c1, r2: run.r, c2: run.c2 }; out.push(n); open.set(key, n); }
  }
  return out;
}

const CF_OP_R = { '>': 'gt', '<': 'lt', '>=': 'ge', '<=': 'le', '=': 'eq', '!=': 'ne', '<>': 'ne' };
const ofExpr = (s) => { const x = fromOpenFormula(`of:=${String(s).trim()}`); return x != null ? fromOoxml(x) : String(s).trim(); };
/** split "a,b" / "a;b" at the top level */
function args2(s) {
  let depth = 0; let q = false;
  for (let i = 0; i < s.length; i++) { const ch = s[i]; if (ch === '"') q = !q; else if (!q && (ch === '(' || ch === '[' || ch === '{')) depth++; else if (!q && (ch === ')' || ch === ']' || ch === '}')) depth--; else if (!q && !depth && (ch === ',' || ch === ';')) return [s.slice(0, i), s.slice(i + 1)]; }
  return [s, ''];
}
function cfFromCondition(v) {
  let m;
  const s = String(v || '').trim();
  if ((m = /^(?:cell-content-is-)?between\((.*)\)$/s.exec(s))) { const [a, b] = args2(m[1]); return { type: 'cell', op: 'between', f1: ofExpr(a), f2: ofExpr(b) }; }
  if ((m = /^(?:cell-content-is-)?not-between\((.*)\)$/s.exec(s))) { const [a, b] = args2(m[1]); return { type: 'cell', op: 'notBetween', f1: ofExpr(a), f2: ofExpr(b) }; }
  if ((m = /^(?:cell-content\(\))?(>=|<=|!=|<>|>|<|=)(.*)$/s.exec(s))) return { type: 'cell', op: CF_OP_R[m[1]], f1: ofExpr(m[2]) };
  if ((m = /^(?:formula-is|is-true-formula)\((.*)\)$/s.exec(s))) return { type: 'formula', formula: ofExpr(m[1]) };
  if ((m = /^(contains-text|not-contains-text|begins-with|ends-with)\("(.*)"\)$/s.exec(s))) return { type: 'text', op: { 'contains-text': 'contains', 'not-contains-text': 'notContains', 'begins-with': 'begins', 'ends-with': 'ends' }[m[1]], text: m[2].replace(/""/g, '"') };
  if (s === 'duplicate') return { type: 'dup' };
  if (s === 'unique') return { type: 'unique' };
  if ((m = /^(top|bottom)-(elements|percent)\((\d+)\)$/.exec(s))) return { type: m[1], percent: m[2] === 'percent', rank: +m[3] };
  if (s === 'above-average' || s === 'above-equal-average') return { type: 'above' };
  if (s === 'below-average' || s === 'below-equal-average') return { type: 'below' };
  if (s === 'is-error') return { type: 'error' };
  return null;
}
const STOP_R = Object.fromEntries(Object.entries(STOP_T).map(([k, v]) => [v, k]));
function dvFromCondition(cond) {
  let s = String(cond || '').replace(/^(of|ooo|oooc):/, '').trim();
  const out = { type: 'any', op: 'between', f1: '', f2: '' };
  if (!s) return out;
  let m;
  const cmp = (rest, fn) => {
    let q;
    if ((q = new RegExp(`^${fn}-is-between\\((.*)\\)$`, 's').exec(rest))) { const [a, b] = args2(q[1]); out.op = 'between'; out.f1 = ofExpr(a); out.f2 = ofExpr(b); return true; }
    if ((q = new RegExp(`^${fn}-is-not-between\\((.*)\\)$`, 's').exec(rest))) { const [a, b] = args2(q[1]); out.op = 'notBetween'; out.f1 = ofExpr(a); out.f2 = ofExpr(b); return true; }
    if ((q = new RegExp(`^${fn}\\(\\)\\s*(>=|<=|!=|<>|>|<|=)(.*)$`, 's').exec(rest))) { out.op = { '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '=': 'equal', '!=': 'notEqual', '<>': 'notEqual' }[q[1]]; out.f1 = ofExpr(q[2]); return true; }
    return false;
  };
  if ((m = /^cell-content-is-(whole-number|decimal-number|date|time)\(\)\s*and\s*(.*)$/s.exec(s))) { out.type = { 'whole-number': 'whole', 'decimal-number': 'decimal', date: 'date', time: 'time' }[m[1]]; cmp(m[2].trim(), 'cell-content'); return out; }
  if (/^cell-content-text-length/.test(s)) { out.type = 'textLength'; cmp(s, 'cell-content-text-length'); return out; }
  if ((m = /^cell-content-is-in-list\((.*)\)$/s.exec(s))) {
    out.type = 'list';
    const items = []; let cur = ''; let q = false;
    for (const ch of m[1]) { if (ch === '"') { q = !q; cur += ch; } else if (ch === ';' && !q) { items.push(cur); cur = ''; } else cur += ch; }
    items.push(cur);
    if (items.every((x) => /^\s*"(.*)"\s*$/s.test(x))) out.f1 = `"${items.map((x) => /^\s*"(.*)"\s*$/s.exec(x)[1].replace(/""/g, '"')).join(',')}"`;
    else out.f1 = ofExpr(m[1]);
    return out;
  }
  if ((m = /^is-true-formula\((.*)\)$/s.exec(s))) { out.type = 'custom'; out.f1 = ofExpr(m[1]); return out; }
  return out;
}

/** ODF bytes → Book (with styles). opts: { loc, settings, ext } */
export async function readOds(bytes, { loc, settings = {}, ext = 'ods' } = {}) {
  const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const isZip = u[0] === 0x50 && u[1] === 0x4b;
  const dec = new TextDecoder();
  let zip = null; let contentRoot; let stylesRoot; let settingsRoot = null; let rawStyles = null; let rawMeta = null; let rawSettings = null; let flatText = null;
  if (isZip) {
    zip = await JSZip.loadAsync(u);
    const str = async (p) => { const f = zip.file(p); return f ? dec.decode(await f.async('uint8array')) : null; };
    const c = await str('content.xml'); if (!c) throw new Error('No content.xml – not an OpenDocument spreadsheet');
    contentRoot = parseXml(c);
    rawStyles = await str('styles.xml'); stylesRoot = rawStyles ? parseXml(rawStyles) : null;
    rawSettings = await str('settings.xml'); settingsRoot = rawSettings ? parseXml(rawSettings) : null;
    rawMeta = await str('meta.xml');
  } else {
    flatText = dec.decode(u);
    contentRoot = parseXml(flatText); stylesRoot = contentRoot; settingsRoot = contentRoot;
    if (!contentRoot || !/office:document/.test(contentRoot.n)) throw new Error('Not an OpenDocument spreadsheet');
  }
  const book = new Book({ loc, settings });
  const fid = [];
  // fonts
  const fonts = new Map();
  for (const root of [stylesRoot, contentRoot]) for (const ff of all(kid(root, 'office:font-face-decls'), 'style:font-face')) fonts.set(ff.a['style:name'], String(ff.a['svg:font-family'] || ff.a['style:name']).replace(/^['"]|['"]$/g, ''));
  // styles
  const officeStyles = kid(stylesRoot, 'office:styles');
  const autoContent = kid(contentRoot, 'office:automatic-styles');
  const autoStylesPart = stylesRoot !== contentRoot ? kid(stylesRoot, 'office:automatic-styles') : null;
  const dataStyles = readDataStyles([officeStyles, autoStylesPart, autoContent].filter(Boolean));
  const named = new Map(); const auto = new Map(); let defaultEl = null;
  for (const e of kids(officeStyles)) { if (e.n === 'style:style') named.set(`${e.a['style:family']}|${e.a['style:name']}`, e); else if (e.n === 'style:default-style' && e.a['style:family'] === 'table-cell') defaultEl = e; }
  for (const e of kids(autoContent)) if (e.n === 'style:style') auto.set(`${e.a['style:family']}|${e.a['style:name']}`, e);
  const styleEl = (fam, n) => auto.get(`${fam}|${n}`) || named.get(`${fam}|${n}`);
  const defProps = defaultEl ? readStyleProps(defaultEl, fonts) : {};
  const defNamed = named.get('table-cell|Default');
  const dp = { ...defProps, ...(defNamed ? readStyleProps(defNamed, fonts) : {}) };
  const defFont = { font: dp.font || 'Liberation Sans', size: dp.size || 10 };
  book.props.defaultFont = defFont;
  const resolved = new Map();
  const chain = (n, depth = 0) => {
    const el = styleEl('table-cell', n); if (!el || depth > 8) return { props: {}, nf: null };
    const parent = el.a['style:parent-style-name'];
    const up = parent ? chain(parent, depth + 1) : { props: { ...dp }, nf: null };
    const own = readStyleProps(el, fonts);
    const props = { ...up.props };
    for (const [k, v] of Object.entries(own)) { if (k === 'bd' && v) props.bd = { ...(props.bd || {}), ...v }; else props[k] = v; }
    const dsn = el.a['style:data-style-name'];
    return { props, nf: dsn ? dataStyles.get(dsn) || null : up.nf };
  };
  const cellStyle = (n) => {
    if (!n) return null;
    if (resolved.has(n)) return resolved.get(n);
    const { props, nf } = chain(n);
    const s = { ...props };
    delete s._autoAlign;
    if (s.font === defFont.font) delete s.font;
    if (s.size === defFont.size) delete s.size;
    if (s.color === '#000000') delete s.color;
    if (s.bd) { for (const k of Object.keys(s.bd)) if (!s.bd[k]) delete s.bd[k]; if (!Object.keys(s.bd).length) delete s.bd; }
    if (nf && nf !== 'General') s.nf = nf;
    const x = book.styles.intern(s);
    resolved.set(n, x);
    return x;
  };
  const textStyles = new Map();
  const ctxText = { textStyle: (n) => { if (textStyles.has(n)) return textStyles.get(n); const el = styleEl('text', n); const tp = el && kid(el, 'style:text-properties'); const s = tp ? readTextProps(tp, fonts) : null; const r = s ? Object.fromEntries(Object.entries(s).filter(([, v]) => v != null && v !== false).map(([k, v]) => [k === 'vert' ? 'va' : k, v])) : null; textStyles.set(n, r); return r; } };
  const colStyleW = (n) => { const el = n && styleEl('table-column', n); const p = el && kid(el, 'style:table-column-properties'); return p && p.a['style:column-width'] ? toPx(p.a['style:column-width']) : null; };
  const rowStyleH = (n) => { const el = n && styleEl('table-row', n); const p = el && kid(el, 'style:table-row-properties'); if (!p || !p.a['style:row-height']) return null; return { h: toPx(p.a['style:row-height']), custom: p.a['style:use-optimal-row-height'] !== 'true' }; };
  const tableProps = (n) => { const el = n && styleEl('table', n); const p = el && kid(el, 'style:table-properties'); return { display: p ? p.a['table:display'] !== 'false' : true, tab: p && (p.a['tableooo:tab-color'] || p.a['table:tab-color']), master: el && el.a['style:master-page-name'] }; };
  // master pages → print orientation
  const masters = new Map();
  { const ms = kid(stylesRoot, 'office:master-styles'); const pl = new Map(); for (const e of kids(autoStylesPart || kid(stylesRoot, 'office:automatic-styles'), 'style:page-layout')) pl.set(e.a['style:name'], kid(e, 'style:page-layout-properties')); for (const m of kids(ms, 'style:master-page')) { const p = pl.get(m.a['style:page-layout-name']); masters.set(m.a['style:name'], p ? p.a : {}); } }
  // spreadsheet
  const body = kid(contentRoot, 'office:body');
  const ss = kid(body, 'office:spreadsheet');
  if (!ss) throw new Error('Not an OpenDocument spreadsheet');
  const keep = { before: [], after: [], dbKeep: [] };
  let seenTable = false;
  const validations = new Map();
  const globalNames = []; let dbRanges = [];
  for (const e of kids(ss)) {
    if (e.n === 'table:table') { seenTable = true; continue; }
    if (e.n === 'table:calculation-settings') {
      const nd = kid(e, 'table:null-date'); if (nd && /^1904-01-01/.test(nd.a['table:date-value'] || '')) book.date1904 = true;
      const it = kid(e, 'table:iteration'); if (it && it.a['table:status'] === 'enable') book.iterate = { on: true, count: +(it.a['table:steps'] || 100), delta: +(it.a['table:minimum-difference'] || 0.001) };
      continue;
    }
    if (e.n === 'table:content-validations') { for (const v of kids(e, 'table:content-validation')) validations.set(v.a['table:name'], v); continue; }
    if (e.n === 'table:named-expressions') { globalNames.push(e); continue; }
    if (e.n === 'table:database-ranges') { dbRanges = kids(e, 'table:database-range'); continue; }
    (seenTable ? keep.after : keep.before).push(serialize(e));
    if (e.n === 'table:data-pilot-tables') fid.push('Pivot tables (DataPilot) – kept, not editable here');
  }
  // settings (freeze, zoom, grid, active sheet)
  const view = new Map(); let activeName = null;
  if (settingsRoot) {
    for (const set of all(settingsRoot, 'config:config-item-set')) {
      if (set.a['config:name'] !== 'ooo:view-settings') continue;
      for (const mapNamed of all(set, 'config:config-item-map-named')) {
        if (mapNamed.a['config:name'] !== 'Tables') continue;
        for (const ent of kids(mapNamed, 'config:config-item-map-entry')) { const o = {}; for (const it of kids(ent, 'config:config-item')) o[it.a['config:name']] = textOfEl(it); view.set(ent.a['config:name'], o); }
      }
      for (const it of all(set, 'config:config-item')) if (it.a['config:name'] === 'ActiveTable') activeName = textOfEl(it);
    }
  }
  const sheetsByName = new Map();
  const chartDirs = []; const frameJobs = [];
  for (const tEl of kids(ss, 'table:table')) {
    const sh = newSheet(tEl.a['table:name'] || `Sheet${book.sheets.length + 1}`);
    book.sheets.push(sh); sheetsByName.set(sh.name, sh);
    const tp = tableProps(tEl.a['table:style-name']);
    if (!tp.display) sh.hidden = 'hidden';
    if (tp.tab) sh.color = String(tp.tab).toUpperCase();
    sh._ods = { frames: [], cellFrames: new Map(), master: tp.master || null };
    const mpa = tp.master ? masters.get(tp.master) : null;
    if (mpa) { if (mpa['style:print-orientation'] === 'landscape') sh.print.orientation = 'landscape'; const pr = mpa['style:print'] || ''; if (/\bgrid\b/.test(pr)) sh.print.gridlines = true; if (/\bheaders\b/.test(pr)) sh.print.headings = true; }
    sh._ods.printSig = printSig(sh);
    if (tEl.a['table:print-ranges']) { const pr = odfToExcelRange(tEl.a['table:print-ranges']); if (pr) sh.print.area = pr.replace(/^.*!/, ''); }
    if (tEl.a['table:protected'] === 'true') sh.protection = { on: true };
    const colDefault = []; // per column: default cell style name (runs)
    let col = 0;
    const walkCols = (el, lvl) => {
      for (const x of kids(el)) {
        if (x.n === 'table:table-column') {
          const n = Math.min(+(x.a['table:number-columns-repeated'] || 1), MAXC - col);
          const w = colStyleW(x.a['table:style-name']);
          const hid = x.a['table:visibility'] === 'collapse';
          const ds = x.a['table:default-cell-style-name'];
          if (n <= 1024 || (!hid && !lvl)) colDefault.push({ c1: col, c2: col + n - 1, s: ds && ds !== 'Default' ? ds : null });
          if (col < 1024) for (let i = 0; i < Math.min(n, 1024 - col); i++) { if (w != null && Math.abs(w - sh.defW) > 0.5) sh.colW.set(col + i, Math.round(w)); if (hid) sh.hiddenCols.add(col + i); if (lvl) sh.outlineCols.set(col + i, Math.min(7, lvl)); }
          col += n;
        } else if (x.n === 'table:table-column-group') walkCols(x, lvl + 1);
        else if (/^table:table-(header-)?columns$/.test(x.n)) walkCols(x, lvl);
      }
    };
    walkCols(tEl, 0);
    // default column width: the most common width when every column has one
    const colDef = (c) => { for (const d of colDefault) if (c >= d.c1 && c <= d.c2) return d.s; return null; };
    const cfMaps = new Map(); // style name → [points]
    const dvPoints = new Map();
    let row = 0;
    const walkRows = (el, lvl) => {
      for (const x of kids(el)) {
        if (x.n === 'table:table-row') { readRow(x, lvl); continue; }
        if (x.n === 'table:table-row-group') { walkRows(x, lvl + 1); continue; }
        if (/^table:table-(header-)?rows$/.test(x.n)) { walkRows(x, lvl); continue; }
        if (x.n === 'table:shapes') { for (const f of kids(x)) frameJobs.push({ sh, el: f, cell: null }); continue; }
        if (x.n === 'table:named-expressions') { readNames(x, sh); continue; }
        if (x.n === 'calcext:conditional-formats') { readCalcext(x, sh); continue; }
      }
    };
    const readRow = (x, lvl) => {
      const rep = Math.min(+(x.a['table:number-rows-repeated'] || 1), MAXR - row);
      const rs = rowStyleH(x.a['table:style-name']);
      const vis = x.a['table:visibility'];
      const rowDefStyle = x.a['table:default-cell-style-name'];
      const cells = kids(x).filter((c) => c.n === 'table:table-cell' || c.n === 'table:covered-table-cell');
      const hasContent = cells.some((c) => c.c && c.c.length || c.a['office:value-type'] || c.a['table:formula']);
      if (rep <= 100000 || hasContent) {
        const n = Math.min(rep, 100000);
        for (let i = 0; i < n; i++) {
          const r = row + i;
          if (rs && rs.custom && Math.abs(rs.h - sh.defH) > 0.5 && (rep < 5000)) sh.rowH.set(r, Math.round(rs.h));
          if (vis === 'collapse' && rep < 5000) sh.hiddenRows.add(r);
          if (vis === 'filter' && rep < 5000) sh.filterHidden.add(r);
          if (lvl && rep < 5000) sh.outlineRows.set(r, Math.min(7, lvl));
        }
      }
      let c = 0;
      for (const ce of cells) {
        const crep = Math.min(+(ce.a['table:number-columns-repeated'] || 1), MAXC - c);
        const covered = ce.n === 'table:covered-table-cell';
        const sname = ce.a['table:style-name'] || rowDefStyle || colDef(c);
        const vt = ce.a['office:value-type'];
        const hasF = !!ce.a['table:formula'];
        const ps = kids(ce, 'text:p');
        const ann = kid(ce, 'office:annotation');
        const frames = kids(ce).filter((f) => /^draw:/.test(f.n));
        const content = vt || hasF || ps.length || ann || frames.length;
        const cells2 = crep * rep;
        if (!content && (!sname || cells2 > 5000)) {
          if (ce.a['table:content-validation-name'] && cells2 <= 50000) for (let i = 0; i < rep; i++) for (let j = 0; j < crep; j++) pushPt(dvPoints, ce.a['table:content-validation-name'], row + i, c + j);
          if (sname && cells2 <= 50000 && styleHasMaps(sname)) for (let i = 0; i < rep; i++) for (let j = 0; j < crep; j++) pushPt(cfMaps, sname, row + i, c + j);
          c += crep; continue;
        }
        // value
        let v = null;
        const ctx = { ...ctxText, link: null };
        const richRuns = []; let rich = false;
        const txt = ps.map((p, i) => { const runs = []; const t = paraText(p, runs, ctx, null); if (runs.some((q) => Object.keys(q).length > 1)) rich = true; if (i) richRuns.push({ t: '\n' }); richRuns.push(...runs); return t; }).join('\n');
        const cvt = ce.a['calcext:value-type'];
        if (cvt === 'error') v = errOf(txt || '#N/A');
        else if (vt === 'float' || vt === 'percentage' || vt === 'currency') v = +ce.a['office:value'];
        else if (vt === 'date') v = serialOfIso(ce.a['office:date-value'], book.date1904);
        else if (vt === 'time') v = fractionOfDuration(ce.a['office:time-value']);
        else if (vt === 'boolean') v = ce.a['office:boolean-value'] === 'true' || ce.a['office:boolean-value'] === '1';
        else if (vt === 'string') v = ce.a['office:string-value'] != null ? ce.a['office:string-value'] : txt;
        else if (!vt && txt) v = txt;
        if (hasF && typeof v === 'string' && /^#(N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!)$/.test(v) && !cvt) v = errOf(v);
        const s = cellStyle(sname);
        for (let i = 0; i < rep; i++) {
          for (let j = 0; j < crep; j++) {
            const r = row + i; const cc = c + j;
            if (cells2 > 20000 && (i || j) && v == null) continue;
            const cell = {};
            if (v != null && !(covered && v === '')) cell.v = v; else cell.v = null;
            if (hasF) {
              const of = fromOpenFormula(ce.a['table:formula']);
              if (of != null) { cell.f = fromOoxml(of); cell._cached = cell.v; cell._cachedF = cell.f; }
            }
            if (s) cell.s = s;
            if (rich && typeof v === 'string' && !hasF) { const rr = mergeRuns(richRuns, v); if (rr) cell.rich = rr; }
            if (ctx.link) cell.link = ctx.link.startsWith('#') ? `#${ctx.link.slice(1).replace(/\.(?=\$?[A-Z]+\$?\d)/, '!')}` : ctx.link;
            if (ann) { cell.note = kids(ann, 'text:p').map((p) => paraText(p, null, ctxText)).join('\n'); const cr = kid(ann, 'dc:creator'); cell.noteAuthor = cr ? textOfEl(cr) : ''; if (ann.a['office:display'] === 'true') cell.noteShown = true; }
            if (cell.v != null || cell.f || cell.s || cell.note) sh.cells.set(K(r, cc), cell);
            if (!covered) {
              const cs = +(ce.a['table:number-columns-spanned'] || 1); const rsn = +(ce.a['table:number-rows-spanned'] || 1);
              if (cs > 1 || rsn > 1) sh.merges.push({ r1: r, c1: cc, r2: r + rsn - 1, c2: cc + cs - 1 });
            }
            if (ce.a['table:content-validation-name']) pushPt(dvPoints, ce.a['table:content-validation-name'], r, cc);
            if (sname && styleHasMaps(sname)) pushPt(cfMaps, sname, r, cc);
            if (!i && !j) for (const f of frames) frameJobs.push({ sh, el: f, cell: { r, c: cc } });
          }
        }
        c += crep;
      }
      row += rep;
    };
    const readNames = (x, scope) => {
      for (const n of kids(x)) {
        const name = n.a['table:name']; if (!name) continue;
        let ref = null;
        if (n.n === 'table:named-range') ref = odfToExcelRange(String(n.a['table:cell-range-address'] || '').replace(/^\$/, ''));
        else if (n.n === 'table:named-expression') { const of = fromOpenFormula(n.a['table:expression'] || ''); if (of != null) ref = of; }
        if (!ref) continue;
        book.names.set(`${name.toUpperCase()}|${scope ? scope.id : ''}`, { name, ref: fromOoxml(ref), scope: scope || null, comment: '' });
      }
    };
    const readCalcext = (x, shx) => {
      for (const cf of kids(x, 'calcext:conditional-format')) {
        const ranges = String(cf.a['calcext:target-range-address'] || '').split(/\s+/).map((a) => odfToExcelRange(a)).filter(Boolean).map((a) => rangeOf(a)).filter(Boolean);
        if (!ranges.length) continue;
        for (const e of kids(cf)) {
          let rule = null;
          if (e.n === 'calcext:condition') {
            rule = cfFromCondition(e.a['calcext:value']);
            const stn = e.a['calcext:apply-style-name'];
            if (rule && stn) { const el = named.get(`table-cell|${stn}`); if (el) { const p = readStyleProps(el, fonts); const sty = {}; for (const k of ['b', 'i', 'u', 'st', 'color', 'fill', 'bd']) if (p[k]) sty[k] = p[k]; const dsn = el.a['style:data-style-name']; if (dsn && dataStyles.get(dsn)) sty.nf = dataStyles.get(dsn); rule.style = sty; } }
          } else if (e.n === 'calcext:color-scale') {
            const ents = kids(e, 'calcext:color-scale-entry');
            rule = { type: 'scale', colors: ents.map((q) => String(q.a['calcext:color'] || '#000000').toUpperCase()), stops: ents.map((q) => ({ type: STOP_R[q.a['calcext:type']] || 'num', ...(/^(minimum|maximum|auto-minimum|auto-maximum)$/.test(q.a['calcext:type']) ? {} : { val: q.a['calcext:value'] }) })) };
          } else if (e.n === 'calcext:data-bar') {
            const ents = kids(e, 'calcext:formatting-entry');
            rule = { type: 'bar', color: String(e.a['calcext:positive-color'] || '#638EC6').toUpperCase(), stops: ents.map((q) => { const t = STOP_R[q.a['calcext:type']] || 'num'; return { type: t === 'autoMin' ? 'min' : t === 'autoMax' ? 'max' : t, ...(/minimum|maximum/.test(q.a['calcext:type']) ? {} : { val: q.a['calcext:value'] }) }; }) };
          } else if (e.n === 'calcext:icon-set') {
            rule = { type: 'icons', set: e.a['calcext:icon-set-type'] || '3TrafficLights1', stops: kids(e, 'calcext:formatting-entry').map((q) => ({ type: STOP_R[q.a['calcext:type']] || 'percent', val: q.a['calcext:value'] })), showValue: e.a['calcext:show-value'] !== 'false' };
          }
          if (rule) shx.cf.push({ id: uid('cf'), ranges, ...rule });
        }
      }
      shx._odsCalcext = true;
    };
    walkRows(tEl, 0);
    // conditional formats from style:map (files without calcext, e.g. Excel's or ODF 1.2 writers)
    if (!sh._odsCalcext) for (const [sname, pts] of cfMaps) {
      const el = styleEl('table-cell', sname);
      for (const m of kids(el, 'style:map')) {
        const cond = String(m.a['style:condition'] || '');
        const rule = /^is-true-formula\(/.test(cond) ? cfFromCondition(cond) : cfFromCondition(cond.replace(/^cell-content-is-(not-)?between/, (q, n) => (n ? 'not-between' : 'between')));
        if (!rule) continue;
        const target = named.get(`table-cell|${m.a['style:apply-style-name']}`);
        if (target) { const p = readStyleProps(target, fonts); const sty = {}; for (const k of ['b', 'i', 'u', 'st', 'color', 'fill', 'bd']) if (p[k]) sty[k] = p[k]; rule.style = sty; }
        sh.cf.push({ id: uid('cf'), ranges: rects(pts), ...rule });
      }
    }
    // data validation
    for (const [vn, pts] of dvPoints) {
      const v = validations.get(vn); if (!v) continue;
      const d = dvFromCondition(v.a['table:condition']);
      const help = kid(v, 'table:help-message'); const err = kid(v, 'table:error-message');
      sh.dv.push({
        id: uid('dv'), ranges: rects(pts), ...d, allowBlank: v.a['table:allow-empty-cell'] !== 'false', dropdown: v.a['table:display-list'] !== 'none',
        prompt: help && help.a['table:display'] !== 'false' ? { title: help.a['table:title'] || '', text: kids(help, 'text:p').map(textOfEl).join('\n') } : null,
        error: { style: err ? ({ stop: 'stop', warning: 'warning', information: 'information' })[err.a['table:message-type']] || 'stop' : 'stop', title: err ? err.a['table:title'] || '' : '', text: err ? kids(err, 'text:p').map(textOfEl).join('\n') : '', show: err ? err.a['table:display'] !== 'false' : true },
      });
    }
    // view
    const vw = view.get(sh.name);
    if (vw) {
      const rows = vw.VerticalSplitMode === '2' ? +vw.VerticalSplitPosition || 0 : 0; const cols = vw.HorizontalSplitMode === '2' ? +vw.HorizontalSplitPosition || 0 : 0;
      if (rows || cols) sh.freeze = { rows, cols };
      if (vw.ZoomValue && +vw.ZoomValue !== 100) sh.zoom = +vw.ZoomValue;
      if (vw.ShowGrid === 'false') sh.gridlines = false;
      if (vw.CursorPositionX != null) sh.sel = { r: +vw.CursorPositionY || 0, c: +vw.CursorPositionX || 0 };
    }
    if (settings.gridlines === false) sh.gridlines = false;
    function styleHasMaps(n) { const el = styleEl('table-cell', n); return !!(el && kid(el, 'style:map')); }
  }
  function pushPt(map, key, r, c) { let a = map.get(key); if (!a) { a = []; map.set(key, a); } a.push([r, c]); }
  function rangeOf(ref) { const m = /^(?:(?:'(?:[^']|'')+'|[^!]+)!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(ref); if (!m) return null; const ci = (s) => s.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; const r1 = +m[2] - 1; const c1 = ci(m[1]); return { r1, c1, r2: m[4] ? +m[4] - 1 : r1, c2: m[3] ? ci(m[3]) : c1 }; }
  if (!book.sheets.length) book.sheets.push(newSheet('Sheet1'));
  // global names, autofilters
  for (const ne of globalNames) for (const n of kids(ne)) {
    const name = n.a['table:name']; if (!name) continue;
    let ref = null;
    if (n.n === 'table:named-range') ref = odfToExcelRange(String(n.a['table:cell-range-address'] || '').replace(/^\$/, ''));
    else if (n.n === 'table:named-expression') ref = fromOpenFormula(n.a['table:expression'] || '');
    if (ref) book.names.set(`${name.toUpperCase()}|`, { name, ref: fromOoxml(ref), scope: null, comment: '' });
  }
  for (const d of dbRanges) {
    const ref = odfToExcelRange(String(d.a['table:target-range-address'] || '').replace(/^\$/, ''));
    const m = ref && /^(?:'((?:[^']|'')+)'|([^!]+))!/.exec(ref);
    const sh = m ? sheetsByName.get((m[1] || m[2]).replace(/''/g, "'")) : null;
    if (d.a['table:display-filter-buttons'] === 'true' && sh && !sh.filter) { const g = rangeOf(ref); if (g) { sh.filter = { range: g, cols: {} }; continue; } }
    keep.dbKeep.push(serialize(d));
  }
  // frames: images, charts, everything else kept as it is
  const cw = (sh, c) => (sh.hiddenCols.has(c) ? 0 : sh.colW.get(c) ?? sh.defW);
  const rh = (sh, r) => (sh.hiddenRows.has(r) ? 0 : sh.rowH.get(r) ?? sh.defH);
  const colX = (sh, c) => { let x = 0; for (let i = 0; i < c; i++) x += cw(sh, i); return x; };
  const rowY = (sh, r) => { let y = 0; for (let i = 0; i < r; i++) y += rh(sh, i); return y; };
  const anchorAt = (sh, x, y, w, h, cell) => {
    if (cell) { const cx = colX(sh, cell.c); const cy = rowY(sh, cell.r); if (x < cx - 1 || y < cy - 1) { x += cx; y += cy; } }
    let c = 0; let acc = 0; while (c < 16383) { const w = cw(sh, c); if (acc + w > x + 0.5) break; acc += w; c++; }
    const dx = x - acc;
    let r = 0; let accY = 0; while (r < 1048575) { const h = rh(sh, r); if (accY + h > y + 0.5) break; accY += h; r++; }
    return { r, c, dx: Math.round(dx), dy: Math.round(y - accY), w: Math.max(20, Math.round(w)), h: Math.max(20, Math.round(h)) };
  };
  for (const job of frameJobs) {
    const { sh, el, cell } = job;
    const raw = () => { if (cell) { const k = K(cell.r, cell.c); const a = sh._ods.cellFrames.get(k) || []; a.push(serialize(el)); sh._ods.cellFrames.set(k, a); } else sh._ods.frames.push(serialize(el)); };
    if (el.n !== 'draw:frame') { raw(); if (!fid.includes('Drawing shapes – kept, not shown')) fid.push('Drawing shapes – kept, not shown'); continue; }
    const x = toPx(el.a['svg:x']) || 0; const y = toPx(el.a['svg:y']) || 0; const w = toPx(el.a['svg:width']) || 200; const h = toPx(el.a['svg:height']) || 150;
    const img = kid(el, 'draw:image'); const obj = kid(el, 'draw:object');
    if (obj) {
      const href = String(obj.a['xlink:href'] || '').replace(/^\.\//, '').replace(/\/$/, '');
      let root = null;
      if (href && zip) { const f = zip.file(`${href}/content.xml`); if (f) root = parseXml(dec.decode(await f.async('uint8array'))); }
      else root = kid(obj, 'office:document');
      const ch = root ? parseOdfChart(root) : null;
      if (!ch) { raw(); if (!fid.includes('Embedded objects – kept, not shown')) fid.push('Embedded objects – kept, not shown'); continue; }
      ch.id = uid('oc'); ch.anchor = anchorAt(sh, x, y, w, h, cell);
      ch._ods = { dir: href || null, ranges: obj.a['draw:notify-on-update-of-ranges'] || null };
      ch._ods.sig = JSON.stringify(chartSigOds(ch));
      if (href) chartDirs.push(href);
      if (ch._unsupported || ch._partial) fid.push('Some chart types (stock, bubble, net, combined) – shown simplified');
      sh.charts.push(ch);
      continue;
    }
    if (img) {
      let bytesImg = null; let mime = 'image/png'; const href = img.a['xlink:href'];
      if (href && zip) { const f = zip.file(href.replace(/^\.\//, '')); if (f) { bytesImg = await f.async('uint8array'); mime = MIMEX[(href.split('.').pop() || '').toLowerCase()] || 'image/png'; } }
      const bd = kid(img, 'office:binary-data'); if (!bytesImg && bd) bytesImg = b64ToBytes(textOfEl(bd).replace(/\s+/g, ''));
      if (!bytesImg) { raw(); continue; }
      if (bytesImg[0] === 0xff && bytesImg[1] === 0xd8) mime = 'image/jpeg'; else if (bytesImg[0] === 0x89 && bytesImg[1] === 0x50) mime = 'image/png'; else if (bytesImg[0] === 0x47 && bytesImg[1] === 0x49) mime = 'image/gif';
      const src = `data:${mime};base64,${bytesToB64(bytesImg)}`;
      sh.images.push({ id: uid('i'), name: el.a['draw:name'] || undefined, src, anchor: anchorAt(sh, x, y, w, h, cell), _ods: href ? { href: href.replace(/^\.\//, ''), src } : null });
      continue;
    }
    raw();
    if (!fid.includes('Drawing shapes – kept, not shown')) fid.push('Drawing shapes – kept, not shown');
  }
  // active sheet
  const ai = book.sheets.findIndex((s) => s.name === activeName);
  book.active = ai >= 0 ? ai : 0;
  // what to keep when saving
  const files = [];
  if (zip) zip.forEach((p, f) => { files.push({ path: p, dir: f.dir, data: () => f.async('uint8array') }); });
  const manifest = new Map();
  if (zip && zip.file('META-INF/manifest.xml')) { const mx = parseXml(dec.decode(await zip.file('META-INF/manifest.xml').async('uint8array'))); for (const e of all(mx, 'manifest:file-entry')) manifest.set(e.a['manifest:full-path'], e.a['manifest:media-type'] || ''); }
  const confRaw = (() => { const src = rawSettings || flatText || ''; const m = /<config:config-item-set config:name="ooo:configuration-settings">[\s\S]*?<\/config:config-item-set>(?=\s*(<config:config-item-set|<\/office:settings>))/.exec(src); return m ? m[0] : ''; })();
  const contentFonts = kids(kid(contentRoot, 'office:font-face-decls'), 'style:font-face').map(serialize);
  const ns = {}; for (const [k, v] of Object.entries(contentRoot.a)) if (k.startsWith('xmlns:')) ns[k] = v;
  const slice = (re) => { const m = flatText && re.exec(flatText); return m ? m[0] : ''; };
  book._ods = {
    flat: !isZip, files, manifest, stylesXml: rawStyles, metaXml: rawMeta, confRaw, contentFonts, ns, chartDirs,
    before: keep.before, after: keep.after, dbKeep: keep.dbKeep,
    flatStyles: flatText ? `${slice(/<office:font-face-decls[\s\S]*?<\/office:font-face-decls>|<office:font-face-decls\s*\/>/)}${slice(/<office:styles[\s\S]*?<\/office:styles>|<office:styles\s*\/>/)}` : '',
    flatMaster: flatText ? slice(/<office:master-styles[\s\S]*?<\/office:master-styles>|<office:master-styles\s*\/>/) : '',
    flatMeta: flatText ? slice(/<office:meta[\s>][\s\S]*?<\/office:meta>/) : '',
  };
  // a flat file keeps the page styles of its automatic styles too
  if (flatText) book._ods.flatPageLayouts = all(kid(contentRoot, 'office:automatic-styles'), 'style:page-layout').map(serialize).join('');
  book._ods.masters = new Set(masters.keys());
  book.fidelity = fid;
  book.calc.rebuild();
  book.calc.recalc(null, { force: true });
  void ext;
  return book;
}
