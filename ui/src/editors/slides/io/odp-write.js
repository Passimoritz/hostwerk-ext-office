/*
 * OpenDocument presentation writer (odp/otp) – own writer, no DOM. writeOdp(doc, { ext, chartImage, extras }) → Uint8Array.
 * mimetype is the first zip entry and stored uncompressed. Colours are resolved against the theme (ODF has no theme
 * colours); what plain ODF cannot express (theme, layouts, placeholder level styles, multi-stop gradients …) travels in
 * hw:data attributes (namespace urn:hostwerk:office:slides:1) and a meta:user-defined "hostwerk" entry, which readOdp
 * uses for an exact round trip. Other apps ignore them.
 */
import JSZip from 'jszip';
import { esc, escAttr, attrs, tag, XML_HEAD } from '../../_shared/xmltree.js';
import { emuToLength } from '../units.js';
import { resolveHex, resolveAlpha, resolveFont, THEMES } from '../themes.js';
import { shapePath, SHAPES } from '../render/geometry.js';
import { dataUrlToBytes, hashStr, EXT_OF_MIME, isZip } from './util.js';

export const NS = {
  'xmlns:office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  'xmlns:style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  'xmlns:text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  'xmlns:table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  'xmlns:draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  'xmlns:fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  'xmlns:xlink': 'http://www.w3.org/1999/xlink',
  'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
  'xmlns:meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  'xmlns:number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  'xmlns:presentation': 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
  'xmlns:svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  'xmlns:smil': 'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0',
  'xmlns:anim': 'urn:oasis:names:tc:opendocument:xmlns:animation:1.0',
  'xmlns:officeooo': 'http://openoffice.org/2009/office',
  'xmlns:loext': 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
  'xmlns:hw': 'urn:hostwerk:office:slides:1',
  'xmlns:dr3d': 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0',
  'xmlns:chart': 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
  'xmlns:math': 'http://www.w3.org/1998/Math/MathML',
  'xmlns:form': 'urn:oasis:names:tc:opendocument:xmlns:form:1.0',
  'xmlns:script': 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
  'xmlns:ooo': 'http://openoffice.org/2004/office',
  'xmlns:xhtml': 'http://www.w3.org/1999/xhtml',
};
/** full (non-sampling) string hash */
export function fullHash(str) { let h = 2166136261 >>> 0; let k = 5381; for (let i = 0; i < str.length; i++) { const c = str.charCodeAt(i); h ^= c; h = Math.imul(h, 16777619) >>> 0; k = (Math.imul(k, 33) + c) >>> 0; } return h.toString(36) + k.toString(36) + str.length.toString(36); }
/** signature of everything styles.xml is generated from – equal => the original styles.xml can be kept */
export const odfSignature = (doc) => fullHash(JSON.stringify({ size: doc.size, theme: doc.theme, layouts: doc.layouts, footer: doc.footer }));
/** signature of a slide's animations (ids ignored) – equal => the original anim:par can be kept */
export const animSig = (list) => JSON.stringify((list || []).map((a) => { const { id, ...r } = a; return r; })); // eslint-disable-line no-unused-vars
const MIMETYPES = { odp: 'application/vnd.oasis.opendocument.presentation', otp: 'application/vnd.oasis.opendocument.presentation-template' };

const L = (e) => emuToLength(Math.round(Number(e) || 0), 'cm');
const PT = (n) => `${+(+n).toFixed(2)}pt`;
const pct = (n) => `${Math.round(n)}%`;
const secs = (ms) => `${+((Number(ms) || 0) / 1000).toFixed(3)}s`;
const safeId = (id) => { let s = String(id || 'x').replace(/[^A-Za-z0-9_.-]/g, '_'); if (!/^[A-Za-z_]/.test(s)) s = `id${s}`; return s; };
const json = (o) => JSON.stringify(o, (k, v) => (v === undefined ? undefined : v));
const BUILTIN = new Set(THEMES.map((t) => t.id));

const NUM_FMT = { arabicPeriod: ['1', '.', ''], arabicParenR: ['1', ')', ''], arabicParenBoth: ['1', ')', '('], alphaLcPeriod: ['a', '.', ''], alphaUcPeriod: ['A', '.', ''], alphaLcParenR: ['a', ')', ''], alphaUcParenR: ['A', ')', ''], romanUcPeriod: ['I', '.', ''], romanLcPeriod: ['i', '.', ''] };
const MARKERS = {
  triangle: ['Arrow', '0 0 20 30', 'M10 0l-10 30h20z'],
  arrow: ['Line_20_Arrow', '0 0 1122 2243', 'M0 2108v17 17l12 42 30 34 38 21 43 4 29-8 30-21 25-26 13-34 343-1532 339 1520 13 42 29 34 39 21 42 4 42-12 34-30 21-42v-39-12l-4 4-440-1998-9-42-25-39-38-25-43-8-42 8-38 25-26 39-8 42z'],
  stealth: ['Arrow_20_concave', '0 0 1131 1580', 'M1013 1491l118 89-567-1580-564 1580 114-85 136-68 148-46 161-17 161 13 153 46z'],
  oval: ['Circle', '0 0 1131 1131', 'M462 1118l-102-29-102-51-93-72-72-93-51-102-29-102-13-105 13-102 29-106 51-102 72-89 93-72 102-50 102-34 106-9 101 9 106 34 98 50 93 72 72 89 51 102 29 106 13 102-13 105-29 102-51 102-72 93-93 72-98 51-106 29-101 13z'],
  diamond: ['Square_20_45', '0 0 1131 1131', 'M0 564l564 567 567-567-567-564z'],
};
const DASHES = { dash: [1, 400, 0, 0, 300], lgDash: [1, 800, 0, 0, 300], dot: [1, 100, 0, 0, 100], dashDot: [1, 400, 1, 100, 300], sysDash: [1, 300, 0, 0, 100], sysDot: [1, 100, 0, 0, 100], lgDashDot: [1, 800, 1, 100, 300], lgDashDotDot: [1, 800, 2, 100, 300] };
const TRANS = { fade: ['fade', 'crossfade'], push: ['pushWipe', null], wipe: ['barWipe', null], zoom: ['irisWipe', 'rectangle'], cover: ['slideWipe', null], split: ['barnDoorWipe', 'vertical'], dissolve: ['dissolve', 'default'] };
// dir = direction of movement (OOXML style): 'l' moves to the left = comes from the right
const FROM = { l: 'fromRight', r: 'fromLeft', u: 'fromBottom', d: 'fromTop' };
const FLY_SUB = { b: 'from-bottom', t: 'from-top', l: 'from-left', r: 'from-right' };

/** pixel size of a PNG/JPEG/GIF (for fo:clip) */
export function imageSize(b) {
  try {
    if (b[0] === 0x89 && b[1] === 0x50) return { w: (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19]) >>> 0, h: (b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23]) >>> 0 };
    if (b[0] === 0x47 && b[1] === 0x49) return { w: b[6] | b[7] << 8, h: b[8] | b[9] << 8 };
    if (b[0] === 0xff && b[1] === 0xd8) {
      let i = 2;
      while (i < b.length) {
        if (b[i] !== 0xff) { i++; continue; }
        const m = b[i + 1]; const len = b[i + 2] << 8 | b[i + 3];
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { h: b[i + 5] << 8 | b[i + 6], w: b[i + 7] << 8 | b[i + 8] };
        i += 2 + len;
      }
    }
  } catch { /* unknown */ }
  return null;
}

/** automatic style registry of one file */
class Styles {
  constructor(prefix) { this.prefix = prefix; this.map = new Map(); this.xml = []; this.n = {}; }
  add(family, inner, extra = {}, short = family[0]) {
    const key = family + json(extra) + inner;
    if (this.map.has(key)) return this.map.get(key);
    this.n[short] = (this.n[short] || 0) + 1;
    const name = `${this.prefix}${short}${this.n[short]}`;
    this.map.set(key, name);
    this.xml.push(tag('style:style', { 'style:name': name, 'style:family': family, ...extra }, inner));
    return name;
  }
  raw(key, make) { if (this.map.has(key)) return this.map.get(key); this.n.L = (this.n.L || 0) + 1; const name = `${this.prefix}L${this.n.L}`; this.map.set(key, name); this.xml.push(make(name)); return name; }
}

class Ctx {
  constructor(doc, opts) {
    this.doc = doc; this.theme = doc.theme; this.opts = opts; this.extras = opts.extras !== false;
    this.common = new Map(); this.commonXml = []; this.files = new Map(); this.fonts = new Set(); this.idMap = {};
  }
  hex(c, fb = '#000000') { return (c && resolveHex(c, this.theme)) || fb; }
  font(f) { const n = resolveFont(f, this.theme); this.fonts.add(n); return n; }
  file(src, folder = 'Pictures') {
    const d = dataUrlToBytes(src); if (!d) return null;
    const key = hashStr(src);
    if (!this.files.has(key)) this.files.set(key, { path: `${folder}/${key}.${EXT_OF_MIME[d.mime] || (d.mime.split('/')[1] || 'bin').replace(/[^a-z0-9]/g, '')}`, bytes: d.bytes, mime: d.mime });
    return this.files.get(key);
  }
  commonStyle(key, make) { if (!this.common.has(key)) { const name = `hw_${fullHash(key)}`; this.common.set(key, name); this.commonXml.push(make(name)); } return this.common.get(key); }
  mname(l) { return (this.keepStyles && l.master) || safeId(l.id); }
  id(e) { if (!this.idMap[e.id]) this.idMap[e.id] = safeId(e.id); return this.idMap[e.id]; }
}

/* ---------------------------------------------------------------- fills, lines, shadows */
function fillAttrs(fill, ctx) {
  if (!fill) return { 'draw:fill': 'none' };
  if (fill.type === 'solid') { const a = resolveAlpha(fill.color); return { 'draw:fill': 'solid', 'draw:fill-color': ctx.hex(fill.color), 'draw:opacity': a < 0.999 ? pct(a * 100) : undefined }; }
  if (fill.type === 'grad') {
    const st = fill.stops && fill.stops.length ? fill.stops : [{ pos: 0, color: '#FFFFFF' }, { pos: 100, color: '#000000' }];
    const s = ctx.hex(st[0].color); const e = ctx.hex(st[st.length - 1].color);
    const radial = fill.kind === 'radial';
    const ang = ((90 - (Number(fill.angle) || 0)) % 360 + 360) % 360;
    const name = ctx.commonStyle(`g${s}${e}${radial}${ang}`, (n) => tag('draw:gradient', { 'draw:name': n, 'draw:style': radial ? 'radial' : 'linear', 'draw:cx': radial ? '50%' : undefined, 'draw:cy': radial ? '50%' : undefined, 'draw:start-color': s, 'draw:end-color': e, 'draw:start-intensity': '100%', 'draw:end-intensity': '100%', 'draw:angle': `${ang}deg`, 'draw:border': '0%' }, null));
    const a = resolveAlpha(st[0].color);
    return { 'draw:fill': 'gradient', 'draw:fill-gradient-name': name, 'draw:opacity': a < 0.999 ? pct(a * 100) : undefined };
  }
  if (fill.type === 'image' && fill.src) {
    const f = ctx.file(fill.src); if (!f) return { 'draw:fill': 'none' };
    const name = ctx.commonStyle(`i${f.path}`, (n) => tag('draw:fill-image', { 'draw:name': n, 'xlink:href': f.path, 'xlink:type': 'simple', 'xlink:show': 'embed', 'xlink:actuate': 'onLoad' }, null));
    return { 'draw:fill': 'bitmap', 'draw:fill-image-name': name, 'style:repeat': fill.mode === 'tile' ? 'repeat' : 'stretch' };
  }
  return { 'draw:fill': 'none' };
}
function lineAttrs(line, ctx) {
  if (!line) return { 'draw:stroke': 'none' };
  const w = line.width == null ? 9525 : line.width;
  const a = resolveAlpha(line.color);
  const out = { 'draw:stroke': 'solid', 'svg:stroke-color': ctx.hex(line.color), 'svg:stroke-width': L(w), 'svg:stroke-opacity': a < 0.999 ? pct(a * 100) : undefined, 'draw:stroke-linejoin': 'round' };
  if (line.dash && line.dash !== 'solid' && DASHES[line.dash]) {
    const [d1, l1, d2, l2, dist] = DASHES[line.dash];
    out['draw:stroke'] = 'dash';
    ctx.commonStyle(`d${line.dash}`, () => tag('draw:stroke-dash', { 'draw:name': `hw_${line.dash}`, 'draw:style': line.dash.includes('ot') ? 'round' : 'rect', 'draw:dots1': d1, 'draw:dots1-length': `${l1}%`, 'draw:dots2': d2 || undefined, 'draw:dots2-length': d2 ? `${l2}%` : undefined, 'draw:distance': `${dist}%` }, null));
    out['draw:stroke-dash'] = `hw_${line.dash}`;
  }
  const mw = L(Math.max(91440, w * 3.5));
  for (const [end, key] of [['start', line.head], ['end', line.tail]]) {
    if (!key || key === 'none' || !MARKERS[key]) continue;
    const [name, vb, d] = MARKERS[key];
    ctx.commonStyle(`m${key}`, () => tag('draw:marker', { 'draw:name': name, 'svg:viewBox': vb, 'svg:d': d }, null));
    out[`draw:marker-${end}`] = name; out[`draw:marker-${end}-width`] = mw;
  }
  return out;
}
function shadowAttrs(sh, ctx) {
  if (!sh) return { 'draw:shadow': 'hidden' };
  const a = ((Number(sh.dir) || 45) * Math.PI) / 180; const d = Number(sh.dist) || 38100;
  return { 'draw:shadow': 'visible', 'draw:shadow-offset-x': L(Math.cos(a) * d), 'draw:shadow-offset-y': L(Math.sin(a) * d), 'draw:shadow-color': ctx.hex(sh.color || '#000000'), 'draw:shadow-opacity': pct(sh.alpha == null ? 40 : sh.alpha), 'loext:shadow-blur': L(sh.blur || 0) };
}
function textAreaAttrs(b) {
  if (!b) return {};
  const ins = b.insets || {};
  return {
    'draw:textarea-vertical-align': { t: 'top', ctr: 'middle', b: 'bottom' }[b.anchor] || 'top',
    'fo:padding-left': L(ins.l ?? 91440), 'fo:padding-right': L(ins.r ?? 91440), 'fo:padding-top': L(ins.t ?? 45720), 'fo:padding-bottom': L(ins.b ?? 45720),
    'draw:auto-grow-height': b.autofit === 'resize' ? 'true' : 'false', 'draw:auto-grow-width': 'false',
    'style:shrink-to-fit': b.autofit === 'shrink' ? 'true' : undefined,
    'fo:wrap-option': b.wrap === false ? 'no-wrap' : 'wrap',
    'style:writing-mode': b.vert ? 'tb-rl' : undefined,
  };
}

/* ---------------------------------------------------------------- text */
const DEF_LVL = { font: '+mn', size: 18, color: 'tx1', bullet: false, marL: 0, indent: 0 };
function effLvl(b, level) { return { ...DEF_LVL, ...((b && b.style) || {}), ...((b && b.lvl && b.lvl[level]) || {}) }; }
function textProps(r, ctx) {
  const o = {};
  if (r.font) { const f = ctx.font(r.font); o['style:font-name'] = f; o['fo:font-family'] = f.includes(' ') ? `'${f}'` : f; }
  if (r.size) o['fo:font-size'] = PT(r.size);
  if (r.b !== undefined) o['fo:font-weight'] = r.b ? 'bold' : 'normal';
  if (r.i !== undefined) o['fo:font-style'] = r.i ? 'italic' : 'normal';
  if (r.u !== undefined) Object.assign(o, r.u ? { 'style:text-underline-style': 'solid', 'style:text-underline-width': 'auto', 'style:text-underline-color': 'font-color' } : { 'style:text-underline-style': 'none' });
  if (r.s !== undefined) o['style:text-line-through-style'] = r.s ? 'solid' : 'none';
  if (r.color) o['fo:color'] = ctx.hex(r.color);
  if (r.highlight) o['fo:background-color'] = ctx.hex(r.highlight);
  if (r.baseline) o['style:text-position'] = `${r.baseline > 0 ? 33 : -33}% 58%`;
  if (r.caps) o['fo:text-transform'] = 'uppercase';
  return o;
}
function spaced(s) {
  let out = '';
  const parts = String(s).split(/(\t|\r?\n|\v| {2,}|^ )/);
  for (const p of parts) {
    if (!p) continue;
    if (p === '\t') out += '<text:tab/>';
    else if (p === '\n' || p === '\r\n' || p === '\v') out += '<text:line-break/>';
    else if (/^ +$/.test(p)) out += p.length === 1 ? '<text:s/>' : ` <text:s text:c="${p.length - 1}"/>`;
    else out += esc(p);
  }
  return out;
}
function runXml(r, ctx, st) {
  const props = textProps(r, ctx);
  let inner;
  if (r.field === 'slidenum') inner = `<text:page-number>${esc(r.text || '1')}</text:page-number>`;
  else if (r.field === 'datetime') inner = `<text:date>${esc(r.text || '')}</text:date>`;
  else if (r.field === 'footer') inner = '<presentation:footer/>';
  else inner = spaced(r.text || '');
  let x = Object.keys(props).length ? tag('text:span', { 'text:style-name': st.add('text', tag('style:text-properties', props, null), {}, 'T') }, inner) : inner;
  if (r.link) x = tag('text:a', { 'xlink:href': r.link, 'xlink:type': 'simple' }, x);
  return x;
}
function bulletSig(bu) { return bu ? json(bu) : 'none'; }
function listLevelXml(level, bu, lv, ctx) {
  const marL = Number(lv.marL) || 0; const ind = Number(lv.indent) || 0;
  const lab = Math.max(0, -ind) || 285750;
  const props = tag('style:list-level-properties', { 'text:space-before': L(Math.max(0, marL - lab)), 'text:min-label-width': L(lab) }, null);
  if (bu && bu.num) {
    const [fmt, suf, pre] = NUM_FMT[bu.num] || NUM_FMT.arabicPeriod;
    return tag('text:list-level-style-number', { 'text:level': level + 1, 'style:num-format': fmt, 'style:num-suffix': suf, 'style:num-prefix': pre || undefined, 'text:start-value': bu.start && bu.start !== 1 ? bu.start : undefined }, props + (bu.color ? tag('style:text-properties', { 'fo:color': ctx.hex(bu.color) }, null) : ''));
  }
  const tp = { 'fo:font-size': '100%' };
  if (bu && bu.font) tp['fo:font-family'] = bu.font;
  if (bu && bu.color) tp['fo:color'] = ctx.hex(bu.color);
  return tag('text:list-level-style-bullet', { 'text:level': level + 1, 'text:bullet-char': (bu && bu.char) || '•' }, props + tag('style:text-properties', tp, null));
}
/** paragraphs of a body → XML (text:p / nested text:list) */
function parasXml(b, ctx, st, opt = {}) {
  if (!b || !b.paras) return '';
  const items = b.paras.map((p) => {
    const level = Math.max(0, Math.min(8, p.level || 0));
    const lv = effLvl(b, level);
    const bu = p.bullet !== undefined ? p.bullet : lv.bullet;
    const align = p.align || lv.align;
    const pprops = {
      'fo:text-align': align ? { l: 'start', ctr: 'center', r: 'end', just: 'justify', dist: 'justify' }[align] : undefined,
      'fo:margin-top': (p.spcBef ?? lv.spcBef) != null ? PT(p.spcBef ?? lv.spcBef) : undefined,
      'fo:margin-bottom': (p.spcAft ?? lv.spcAft) != null ? PT(p.spcAft ?? lv.spcAft) : undefined,
      'fo:line-height': (p.lnSpc ?? lv.lnSpc) ? pct(p.lnSpc ?? lv.lnSpc) : undefined,
    };
    if (!bu) { const ml = p.marL ?? lv.marL; const ti = p.indent ?? lv.indent; if (ml) pprops['fo:margin-left'] = L(ml); if (ti) pprops['fo:text-indent'] = L(ti); }
    const dflt = { font: lv.font, size: lv.size, color: lv.color, b: lv.b, i: lv.i, u: lv.u, s: lv.s, ...(opt.runDefaults || {}) };
    const tprops = textProps(dflt, ctx);
    const pst = st.add('paragraph', tag('style:paragraph-properties', pprops, null) + tag('style:text-properties', tprops, null), {}, 'P');
    const runs = (p.runs || []).map((r) => runXml(r, ctx, st)).join('');
    return { level, bu, lv, inList: !!bu || level > 0, xml: tag('text:p', { 'text:style-name': pst }, runs || null) };
  });
  let out = '';
  let i = 0;
  while (i < items.length) {
    const it = items[i];
    if (!it.inList) { out += it.xml; i++; continue; }
    // a group of consecutive list paragraphs with consistent bullets per level
    const sig = {}; const lvls = {}; let j = i;
    while (j < items.length && items[j].inList) {
      const s = bulletSig(items[j].bu);
      if (sig[items[j].level] !== undefined && sig[items[j].level] !== s) break;
      sig[items[j].level] = s; lvls[items[j].level] = items[j]; j++;
    }
    const group = items.slice(i, j);
    const key = json(Object.keys(lvls).sort().map((k) => [k, sig[k], lvls[k].lv.marL, lvls[k].lv.indent]));
    const lname = st.raw(`list${key}`, (n) => tag('text:list-style', { 'style:name': n }, Object.keys(lvls).sort().map((k) => listLevelXml(+k, lvls[k].bu, lvls[k].lv, ctx)).join('')));
    const nest = (arr, d) => {
      let s = ''; let k = 0;
      while (k < arr.length) {
        if (arr[k].level === d) { s += tag(arr[k].bu ? 'text:list-item' : 'text:list-header', null, arr[k].xml); k++; } else {
          let m = k; while (m < arr.length && arr[m].level > d) m++;
          s += `<text:list-item><text:list>${nest(arr.slice(k, m), d + 1)}</text:list></text:list-item>`; k = m;
        }
      }
      return s;
    };
    const minL = Math.min(...group.map((g) => g.level));
    let inner = nest(group, minL);
    for (let d = minL; d > 0; d--) inner = `<text:list-item><text:list>${inner}</text:list></text:list-item>`;
    out += tag('text:list', { 'text:style-name': lname }, inner);
    i = j;
  }
  return out;
}

/* ---------------------------------------------------------------- geometry */
function geo(e) {
  const rot = ((Number(e.rot) || 0) % 360 + 360) % 360;
  if (!rot) return { 'svg:width': L(e.w), 'svg:height': L(e.h), 'svg:x': L(e.x), 'svg:y': L(e.y) };
  const t = (rot * Math.PI) / 180; const c = Math.cos(t); const s = Math.sin(t);
  const cx = e.x + e.w / 2; const cy = e.y + e.h / 2;
  const tx = cx + (-e.w / 2) * c - (-e.h / 2) * s; const ty = cy + (-e.w / 2) * s + (-e.h / 2) * c;
  return { 'svg:width': L(e.w), 'svg:height': L(e.h), 'draw:transform': `rotate (${+((((360 - rot) % 360) * Math.PI) / 180).toFixed(6)}) translate (${L(tx)} ${L(ty)})` };
}
/** replace the geometry attributes of a raw element's start tag */
function patchRaw(raw, g) {
  const m = /^<([^\s/>]+)([^>]*?)(\/?)>/.exec(raw);
  if (!m) return raw;
  const rest = m[2].replace(/\s(svg:x|svg:y|svg:width|svg:height|draw:transform)="[^"]*"/g, '');
  return `<${m[1]}${rest}${attrs(g)}${m[3]}>${raw.slice(m[0].length)}`;
}
function commentsXml(s) {
  return (s.comments || []).map((c) => tag('officeooo:annotation', { 'svg:x': L(c.x || 0), 'svg:y': L(c.y || 0), 'svg:width': '0.6cm', 'svg:height': '0.6cm' },
    (c.author ? tag('dc:creator', null, esc(c.author)) : '') + (c.time ? tag('dc:date', null, esc(String(c.time).replace(/Z$/, ''))) : '') + (c.initials ? tag('meta:creator-initials', null, esc(c.initials)) : '')
    + String(c.text || '').split(/\r?\n/).map((l) => tag('text:p', null, spaced(l) || null)).join(''))).join('');
}
function rotateAround(e, cx, cy, deg) {
  const a = (deg * Math.PI) / 180;
  const ex = e.x + e.w / 2 - cx; const ey = e.y + e.h / 2 - cy;
  const dx = cx + ex * Math.cos(a) - ey * Math.sin(a) - (e.x + e.w / 2); const dy = cy + ex * Math.sin(a) + ey * Math.cos(a) - (e.y + e.h / 2);
  const mv = (x) => ({ ...x, x: x.x + dx, y: x.y + dy, children: x.children ? x.children.map(mv) : undefined });
  return { ...mv(e), rot: ((e.rot || 0) + deg) % 360 };
}
function odfPath(d) {
  return d.replace(/([MLCZ])/g, ' $1 ').replace(/-?\d+(\.\d+)?/g, (n) => String(Math.round(+n))).replace(/\s+/g, ' ').trim() + ' N';
}

/* ---------------------------------------------------------------- elements */
const PH_CLASS = { title: 'title', ctrTitle: 'title', subTitle: 'subtitle', pic: 'graphic', dt: 'date-time', ftr: 'footer', sldNum: 'page-number' };
const phClass = (ph) => PH_CLASS[ph] || (/^body/.test(ph) ? 'outline' : 'outline');

function extrasOf(e, ctx, more = {}) {
  if (!ctx.extras) return {};
  const x = { ph: e.ph, phIdx: e.phIdx, name: e.name || undefined, locked: e.locked || undefined, decor: e.decor || undefined, prompt: e.prompt || undefined, textbox: e.textbox || undefined, link: e.link || undefined, flipH: e.flipH || undefined, flipV: e.flipV || undefined, ...more };
  if (e.text) x.text = { lvl: e.text.lvl, style: e.text.style, autofit: e.text.autofit, fontScale: e.text.fontScale, anchor: e.text.anchor, insets: e.text.insets, bullets: e.text.paras.map((p) => (p.bullet === undefined ? 0 : 1)), paraProps: e.text.paras.map((p) => { const o = {}; for (const k of ['align', 'lnSpc', 'spcBef', 'spcAft', 'marL', 'indent']) if (p[k] !== undefined) o[k] = p[k]; return o; }) };
  for (const k of Object.keys(x)) if (x[k] === undefined) delete x[k];
  return Object.keys(x).length ? { 'hw:data': json(x) } : {};
}

function writeEl(e, ctx, st, masterName) {
  try {
    if (e.hidden && !e.preserved) return '';
    if (e.preserved && e.raw && e.rawFmt === 'odf' && ctx.patch && (e.rawParts || []).every((p) => ctx.hasPart(p))) {
      for (const x of e.rawStyles || []) st.raw(`orig${x}`, () => x);
      const b = e.rawBox || {};
      const moved = ['x', 'y', 'w', 'h'].some((k) => Math.abs((b[k] || 0) - (e[k] || 0)) > 1) || Math.abs((b.rot || 0) - (e.rot || 0)) > 0.01;
      return moved ? patchRaw(e.raw, geo(e)) : e.raw;
    }
    const idA = { 'xml:id': ctx.id(e), 'draw:id': ctx.id(e), 'draw:name': e.name || undefined };
    if (e.type === 'group') {
      const kids = e.rot ? e.children.map((c) => rotateAround(c, e.x + e.w / 2, e.y + e.h / 2, e.rot)) : e.children;
      return tag('draw:g', { ...idA, ...extrasOf(e, ctx, { group: { x: e.x, y: e.y, w: e.w, h: e.h, rot: e.rot || 0 } }) }, kids.map((c) => writeEl(c, ctx, st, masterName)).join(''));
    }
    if (e.type === 'line') {
      let x1 = e.x; let y1 = e.y; let x2 = e.x + e.w; let y2 = e.y + e.h;
      if (e.flipH) [x1, x2] = [x2, x1];
      if (e.flipV) [y1, y2] = [y2, y1];
      if (e.rot) { const cx = e.x + e.w / 2; const cy = e.y + e.h / 2; const a = (e.rot * Math.PI) / 180; const r = (x, y) => [cx + (x - cx) * Math.cos(a) - (y - cy) * Math.sin(a), cy + (x - cx) * Math.sin(a) + (y - cy) * Math.cos(a)]; [x1, y1] = r(x1, y1); [x2, y2] = r(x2, y2); }
      const gs = st.add('graphic', tag('style:graphic-properties', { ...lineAttrs(e.line, ctx), 'draw:fill': 'none', ...shadowAttrs(e.shadow, ctx) }, null), {}, 'gr');
      const t = e.geom === 'elbow' ? 'standard' : e.geom === 'curve' ? 'curve' : null;
      return tag(t ? 'draw:connector' : 'draw:line', { ...idA, 'draw:style-name': gs, 'draw:type': t || undefined, 'svg:x1': L(x1), 'svg:y1': L(y1), 'svg:x2': L(x2), 'svg:y2': L(y2), ...extrasOf(e, ctx, { geom: e.geom, line: e.line }) }, '<text:p/>');
    }
    if (e.type === 'image' || (e.type === 'unknown' && e.src) || (e.type === 'chart' && ctx.opts.chartImage && ctx.opts.chartImage(e))) {
      const src = e.type === 'chart' ? ctx.opts.chartImage(e) : e.src;
      const f = ctx.file(src); if (!f) return '';
      const gp = { ...lineAttrs(e.line, ctx), 'draw:fill': 'none', ...shadowAttrs(e.shadow, ctx) };
      if (e.opacity != null && e.opacity < 1) gp['draw:image-opacity'] = pct(e.opacity * 100);
      const cr = e.crop || {};
      if (cr.l || cr.t || cr.r || cr.b) {
        const sz = imageSize(f.bytes);
        if (sz) { const W = sz.w * 9525; const H = sz.h * 9525; gp['fo:clip'] = `rect(${L((cr.t || 0) * H)}, ${L((cr.r || 0) * W)}, ${L((cr.b || 0) * H)}, ${L((cr.l || 0) * W)})`; }
      }
      const gs = st.add('graphic', tag('style:graphic-properties', gp, null), {}, 'gr');
      const more = e.type === 'chart' ? { chart: e.chart } : e.type === 'unknown' ? { unknown: e.what || 'object' } : { crop: e.crop, line: e.line, shadow: e.shadow };
      const phA = e.ph ? { 'presentation:class': 'graphic' } : {};
      return tag('draw:frame', { ...idA, 'draw:style-name': gs, ...phA, ...geo(e), ...extrasOf(e, ctx, more) }, tag('draw:image', { 'xlink:href': f.path, 'xlink:type': 'simple', 'xlink:show': 'embed', 'xlink:actuate': 'onLoad' }, null) + (e.alt ? `<svg:desc>${esc(e.alt)}</svg:desc>` : ''));
    }
    if (e.type === 'media') {
      const f = ctx.file(e.src, 'Media'); if (!f) return '';
      return tag('draw:frame', { ...idA, ...geo(e), ...extrasOf(e, ctx, { media: e.kind }) }, tag('draw:plugin', { 'xlink:href': f.path, 'xlink:type': 'simple', 'xlink:show': 'embed', 'xlink:actuate': 'onLoad', 'draw:mime-type': e.mime || f.mime }, '<draw:param draw:name="Loop" draw:value="false"/><draw:param draw:name="MediaMimeType" draw:value="' + escAttr(e.mime || f.mime) + '"/>'));
    }
    if (e.type === 'table') return tableXml(e, ctx, st, idA);
    if (e.type === 'chart') return chartTableXml(e, ctx, st, idA);
    if (e.type !== 'shape') return '';
    const b = e.text;
    const gp = { ...fillAttrs(e.fill, ctx), ...lineAttrs(e.line, ctx), ...shadowAttrs(e.shadow, ctx), ...textAreaAttrs(b) };
    const isFrame = e.ph || (e.textbox && (e.geom || 'rect') === 'rect');
    if (isFrame) {
      const cls = e.ph ? phClass(e.ph) : null;
      const empty = !b || b.paras.every((p) => !(p.runs || []).some((r) => r.text));
      const fam = cls ? 'presentation' : 'graphic';
      const parent = cls && masterName ? { 'style:parent-style-name': `${masterName}-${cls === 'outline' ? 'outline1' : cls === 'subtitle' ? 'subtitle' : cls === 'title' ? 'title' : 'outline1'}` } : {};
      const gs = st.add(fam, tag('style:graphic-properties', gp, null), parent, fam === 'presentation' ? 'pr' : 'gr');
      const a = { ...idA, [fam === 'presentation' ? 'presentation:style-name' : 'draw:style-name']: gs, ...geo(e) };
      if (cls) { a['presentation:class'] = cls; if (empty) a['presentation:placeholder'] = 'true'; a['presentation:user-transformed'] = 'true'; }
      return tag('draw:frame', { ...a, ...extrasOf(e, ctx, { fill: e.fill && e.fill.type === 'grad' && e.fill.stops.length > 2 ? e.fill : undefined, shadow: e.shadow || undefined }) }, tag('draw:text-box', null, parasXml(b, ctx, st) || null));
    }
    const geom = SHAPES[e.geom] ? e.geom : 'rect';
    const W = Math.max(1, Math.round(e.w / 360)); const H = Math.max(1, Math.round(e.h / 360));
    const p = shapePath(geom, W, H, e.adj);
    const odfType = SHAPES[geom].odf && !SHAPES[geom].odf.startsWith('ooxml-') ? SHAPES[geom].odf : 'non-primitive';
    const gs = st.add('graphic', tag('style:graphic-properties', gp, null), {}, 'gr');
    const eg = tag('draw:enhanced-geometry', { 'svg:viewBox': `0 0 ${W} ${H}`, 'draw:type': odfType, 'draw:enhanced-path': odfPath(p.d), 'draw:mirror-horizontal': e.flipH ? 'true' : undefined, 'draw:mirror-vertical': e.flipV ? 'true' : undefined, 'draw:text-areas': p.text ? `${Math.round(p.text.x)} ${Math.round(p.text.y)} ${Math.round(p.text.x + p.text.w)} ${Math.round(p.text.y + p.text.h)}` : undefined }, null);
    return tag('draw:custom-shape', { ...idA, 'draw:style-name': gs, ...geo(e), ...extrasOf(e, ctx, { geom: e.geom, adj: e.adj, fill: e.fill && (e.fill.type === 'grad' && e.fill.stops.length > 2) ? e.fill : undefined, shadow: e.shadow || undefined }) }, parasXml(b, ctx, st) + eg);
  } catch (err) {
    return `<!-- element ${esc(String(e && e.id))} skipped -->`;
  }
}

function cellFill(t, ri, ci, c) {
  if (c.fill !== undefined) return c.fill;
  const s = t.style || {}; const acc = s.accent || 'accent1';
  if (s.firstRow && ri === 0) return { type: 'solid', color: acc };
  if (s.firstCol && ci === 0) return { type: 'solid', color: acc };
  if (s.bandRows) return { type: 'solid', color: (ri - (s.firstRow ? 1 : 0)) % 2 === 0 ? `${acc}+80` : `${acc}+60` };
  return null;
}
function borderStr(l, ctx) { return l ? `${L(l.width == null ? 9525 : l.width)} solid ${ctx.hex(l.color)}` : 'none'; }
function tableXml(e, ctx, st, idA) {
  const cols = e.cols.map((w) => tag('table:table-column', { 'table:style-name': st.add('table-column', tag('style:table-column-properties', { 'style:column-width': L(w) }, null), {}, 'co') }, null)).join('');
  const implicit = [];
  const rows = e.rows.map((r, ri) => {
    implicit.push(r.cells.map((c) => (c.fill === undefined ? 1 : 0)));
    const cells = r.cells.map((c, ci) => {
      if (c.merged) return '<table:covered-table-cell/>';
      const bd = c.borders || {};
      const head = (e.style && e.style.firstRow && ri === 0) || (e.style && e.style.firstCol && ci === 0);
      const cs = st.add('table-cell', tag('style:graphic-properties', { ...fillAttrs(cellFill(e, ri, ci, c), ctx), ...textAreaAttrs(c.text) }, null) + tag('style:table-cell-properties', { 'fo:border-left': borderStr(bd.l !== undefined ? bd.l : e.border, ctx), 'fo:border-right': borderStr(bd.r !== undefined ? bd.r : e.border, ctx), 'fo:border-top': borderStr(bd.t !== undefined ? bd.t : e.border, ctx), 'fo:border-bottom': borderStr(bd.b !== undefined ? bd.b : e.border, ctx) }, null), {}, 'ce');
      return tag('table:table-cell', { 'table:style-name': cs, 'table:number-columns-spanned': c.span && c.span.c > 1 ? c.span.c : undefined, 'table:number-rows-spanned': c.span && c.span.r > 1 ? c.span.r : undefined }, parasXml(c.text, ctx, st, head && e.style && !e.style.custom ? { runDefaults: { color: 'bg1', b: true } } : {}) || '<text:p/>');
    }).join('');
    return tag('table:table-row', { 'table:style-name': st.add('table-row', tag('style:table-row-properties', { 'style:row-height': L(r.h) }, null), {}, 'ro') }, cells);
  }).join('');
  return tag('draw:frame', { ...idA, ...geo(e), ...extrasOf(e, ctx, { table: { style: e.style, border: e.border, implicit } }) }, tag('table:table', { 'table:use-first-row-styles': e.style && e.style.firstRow ? 'true' : undefined, 'table:use-banding-rows-styles': e.style && e.style.bandRows ? 'true' : undefined }, cols + rows));
}
function chartTableXml(e, ctx, st, idA) {
  const c = e.chart || {}; const cats = c.categories || []; const ser = c.series || [];
  const mk = (txt, bold) => ({ text: { paras: [{ runs: [{ text: String(txt ?? ''), b: bold || undefined }], level: 0 }], anchor: 'ctr', insets: { l: 45720, t: 22860, r: 45720, b: 22860 } } });
  const rows = [{ h: 0, cells: [mk(c.title || '', true), ...ser.map((s) => mk(s.name, true))] }, ...cats.map((cat, i) => ({ h: 0, cells: [mk(cat, true), ...ser.map((s) => mk(s.values[i]))] }))];
  const n = rows[0].cells.length;
  const t = { ...e, cols: Array.from({ length: n }, () => Math.round(e.w / n)), rows: rows.map((r) => ({ ...r, h: Math.round(e.h / rows.length) })), style: { firstRow: true, bandRows: true, accent: (ser[0] && ser[0].color) || 'accent1' }, border: { color: 'bg1', width: 9525 } };
  const x = tableXml(t, ctx, st, idA);
  return ctx.extras ? x.replace(/hw:data="[^"]*"/, `hw:data="${escAttr(json({ chart: e.chart, name: e.name || undefined }))}"`) : x;
}

/* ---------------------------------------------------------------- animations */
function animXml(slide, ctx) {
  const list = (slide.animations || []).filter((a) => ctx.idMap[a.el]);
  if (!list.length) return '';
  const clicks = [];
  for (const a of list) {
    if (!clicks.length || a.trigger === 'click' || !a.trigger) clicks.push({ pars: [] });
    const g = clicks[clicks.length - 1];
    const prev = g.pars[g.pars.length - 1];
    if (a.trigger === 'with' && prev) prev.effects.push(a);
    else { const begin = prev ? prev.end : 0; g.pars.push({ begin, end: begin, effects: [a] }); }
    const p = g.pars[g.pars.length - 1];
    p.end = Math.max(p.end, p.begin + (a.delay || 0) + (a.effect === 'appear' ? 0 : (a.duration || 500)));
  }
  const eff = (a, first) => {
    const id = ctx.idMap[a.el]; const d = secs(a.effect === 'appear' ? 1 : a.duration || 500);
    const set = `<anim:set smil:begin="0s" smil:dur="0.001s" smil:fill="hold" smil:targetElement="${id}" smil:attributeName="visibility" smil:to="visible"/>`;
    const dir = a.dir || 'b';
    let body = set; let pid = 'ooo-entrance-appear'; let sub;
    if (a.effect === 'fade') { pid = 'ooo-entrance-fade-in'; body += `<anim:transitionFilter smil:dur="${d}" smil:targetElement="${id}" smil:type="fade" smil:subtype="crossfade"/>`; }
    else if (a.effect === 'fly') {
      pid = 'ooo-entrance-fly-in'; sub = FLY_SUB[dir] || 'from-bottom';
      const xs = dir === 'l' ? '0-width/2;x' : dir === 'r' ? '1+width/2;x' : 'x;x'; const ys = dir === 't' ? '0-height/2;y' : dir === 'b' ? '1+height/2;y' : 'y;y';
      body += `<anim:animate smil:dur="${d}" smil:fill="hold" smil:targetElement="${id}" smil:attributeName="x" smil:values="${xs}" smil:keyTimes="0;1" presentation:additive="base"/><anim:animate smil:dur="${d}" smil:fill="hold" smil:targetElement="${id}" smil:attributeName="y" smil:values="${ys}" smil:keyTimes="0;1" presentation:additive="base"/>`;
    } else if (a.effect === 'zoom') { pid = 'ooo-entrance-zoom'; sub = 'in'; body += `<anim:animateTransform smil:dur="${d}" smil:fill="hold" smil:targetElement="${id}" smil:values="0,0;1,1" smil:keyTimes="0;1" svg:type="scale"/>`; }
    else if (a.effect === 'wipe') { pid = 'ooo-entrance-wipe'; sub = FLY_SUB[dir] || 'from-bottom'; const v = dir === 'l' || dir === 'r'; body += `<anim:transitionFilter smil:dur="${d}" smil:targetElement="${id}" smil:type="barWipe" smil:subtype="${v ? 'leftToRight' : 'topToBottom'}"${dir === 'r' || dir === 'b' ? ' smil:direction="reverse"' : ''}/>`; }
    const nt = a.trigger === 'with' ? 'with-previous' : a.trigger === 'after' ? 'after-previous' : first ? 'on-click' : 'on-click';
    return tag('anim:par', { 'smil:begin': secs(a.delay || 0), 'smil:fill': 'hold', 'presentation:node-type': nt, 'presentation:preset-class': 'entrance', 'presentation:preset-id': pid, 'presentation:preset-sub-type': sub, 'hw:dir': a.dir || undefined }, body);
  };
  const seq = clicks.map((g, gi) => tag('anim:par', { 'smil:begin': gi === 0 && list[0].trigger && list[0].trigger !== 'click' ? '0s' : 'next' }, g.pars.map((p) => tag('anim:par', { 'smil:begin': secs(p.begin) }, p.effects.map((a, k) => eff(a, k === 0)).join(''))).join(''))).join('');
  return `<anim:par presentation:node-type="timing-root"><anim:seq presentation:node-type="main-sequence">${seq}</anim:seq></anim:par>`;
}

/* ---------------------------------------------------------------- pages */
function pageStyle(slide, ctx, st, isTitle) {
  const p = {};
  if (slide.background) Object.assign(p, fillAttrs(slide.background, ctx), { 'draw:background-size': 'full' });
  const tr = slide.transition || {};
  if (tr.type && tr.type !== 'none' && TRANS[tr.type]) {
    const [type, sub] = TRANS[tr.type];
    p['smil:type'] = type;
    if (tr.type === 'wipe') { const h = tr.dir === 'l' || tr.dir === 'r' || !tr.dir; p['smil:subtype'] = h ? 'leftToRight' : 'topToBottom'; if (tr.dir === 'l' || tr.dir === 'u') p['smil:direction'] = 'reverse'; }
    else p['smil:subtype'] = sub || FROM[tr.dir || 'l'];
    const d = tr.duration || 500;
    p['presentation:transition-speed'] = d < 600 ? 'fast' : d > 1200 ? 'slow' : 'medium';
  }
  if (tr.advanceAfter != null) { p['presentation:transition-type'] = 'automatic'; p['presentation:duration'] = `PT${+(tr.advanceAfter / 1000).toFixed(3)}S`; }
  if (slide.hidden) p['presentation:visibility'] = 'hidden';
  const f = ctx.doc.footer || {};
  const show = !(f.hideOnTitle && isTitle);
  p['presentation:display-footer'] = f.showFooter && show ? 'true' : 'false';
  p['presentation:display-page-number'] = f.showNumber && show ? 'true' : 'false';
  p['presentation:display-date-time'] = f.showDate && show ? 'true' : 'false';
  p['presentation:background-objects-visible'] = slide.showMaster === false ? 'false' : 'true';
  return st.add('drawing-page', tag('style:drawing-page-properties', p, null), {}, 'dp');
}

function notesXml(slide, idx, ctx, st) {
  const lines = String(slide.notes || '').split(/\r?\n/);
  const W = 7560000; const H = 10692000;
  const gs = st.add('presentation', tag('style:graphic-properties', { 'draw:fill': 'none', 'draw:stroke': 'none' }, null), {}, 'pr');
  return tag('presentation:notes', { 'draw:style-name': undefined }, tag('draw:page-thumbnail', { 'svg:width': L(W * 0.8), 'svg:height': L((W * 0.8 * ctx.doc.size.h) / ctx.doc.size.w), 'svg:x': L(W * 0.1), 'svg:y': L(H * 0.08), 'draw:page-number': idx + 1, 'presentation:class': 'page' }, null)
    + tag('draw:frame', { 'presentation:style-name': gs, 'svg:width': L(W * 0.8), 'svg:height': L(H * 0.45), 'svg:x': L(W * 0.1), 'svg:y': L(H * 0.48), 'presentation:class': 'notes', 'presentation:placeholder': slide.notes ? undefined : 'true' }, tag('draw:text-box', null, slide.notes ? lines.map((l) => tag('text:p', null, spaced(l) || null)).join('') : null)));
}

/* ---------------------------------------------------------------- styles.xml pieces */
function presStyles(layout, mname, ctx) {
  const ph = layout.ph || {};
  const title = ph.title || ph.ctrTitle; const bodyPh = ph.body || ph.body2; const sub = ph.subTitle;
  const tp = (lv) => textProps(lv ? { font: lv.font || '+mn', size: lv.size, color: lv.color || 'tx1', b: lv.b, i: lv.i } : { font: '+mn', size: 18, color: 'tx1' }, ctx);
  const gp = { 'draw:fill': 'none', 'draw:stroke': 'none', 'draw:auto-grow-height': 'false' };
  let x = tag('style:style', { 'style:name': `${mname}-title`, 'style:family': 'presentation' }, tag('style:graphic-properties', { ...gp, 'draw:textarea-vertical-align': { t: 'top', ctr: 'middle', b: 'bottom' }[(title && title.anchor) || 'b'] }, null) + tag('style:paragraph-properties', { 'fo:text-align': title && title.lvl && title.lvl[0].align === 'ctr' ? 'center' : 'start' }, null) + tag('style:text-properties', tp(title && title.lvl && title.lvl[0]), null));
  x += tag('style:style', { 'style:name': `${mname}-subtitle`, 'style:family': 'presentation' }, tag('style:graphic-properties', gp, null) + tag('style:text-properties', tp(sub && sub.lvl && sub.lvl[0]), null));
  const lvls = (bodyPh && bodyPh.lvl) || [];
  const list = tag('text:list-style', { 'style:name': `${mname}-outline-list` }, Array.from({ length: 9 }, (_, i) => listLevelXml(i, (lvls[i] && lvls[i].bullet) || { char: '•' }, lvls[i] || { marL: 285750 + i * 457200, indent: -285750 }, ctx)).join(''));
  for (let i = 1; i <= 9; i++) {
    x += tag('style:style', { 'style:name': `${mname}-outline${i}`, 'style:family': 'presentation', 'style:parent-style-name': i > 1 ? `${mname}-outline${i - 1}` : undefined },
      (i === 1 ? tag('style:graphic-properties', gp, list) : '') + tag('style:paragraph-properties', { 'fo:margin-top': lvls[i - 1] && lvls[i - 1].spcBef != null ? PT(lvls[i - 1].spcBef) : undefined }, null) + tag('style:text-properties', tp(lvls[i - 1]), null));
  }
  for (const n of ['notes', 'background', 'backgroundobjects']) x += tag('style:style', { 'style:name': `${mname}-${n}`, 'style:family': 'presentation' }, n === 'notes' ? tag('style:text-properties', { 'fo:font-size': '12pt' }, null) : null);
  return x;
}

function masterXml(layout, ctx, st) {
  const mname = ctx.mname(layout);
  const bg = layout.bg || ctx.theme.bg || { type: 'solid', color: 'bg1' };
  const dp = st.add('drawing-page', tag('style:drawing-page-properties', { ...fillAttrs(bg, ctx), 'draw:background-size': 'full', 'presentation:background-visible': 'true', 'presentation:background-objects-visible': 'true' }, null), {}, 'Mdp');
  const decor = (layout.decor || []).map((e) => writeEl({ ...e, locked: undefined }, ctx, st, null)).join('');
  const ph = layout.ph || {};
  const phFrame = (key, cls) => {
    const d = ph[key]; if (!d) return '';
    return tag('draw:frame', { 'presentation:style-name': `${mname}-${cls === 'outline' ? 'outline1' : cls}`, 'svg:x': L(d.x), 'svg:y': L(d.y), 'svg:width': L(d.w), 'svg:height': L(d.h), 'presentation:class': cls, 'presentation:placeholder': 'true' }, '<draw:text-box/>');
  };
  const f = ctx.doc.footer || {};
  let foot = '';
  if (f.showFooter || f.showNumber || f.showDate) {
    const ff = (key, cls, inner) => { const d = ph[key]; if (!d) return ''; const lv = (d.lvl && d.lvl[0]) || {}; const ps = st.add('paragraph', tag('style:paragraph-properties', { 'fo:text-align': { l: 'start', ctr: 'center', r: 'end' }[lv.align || 'l'] }, null) + tag('style:text-properties', textProps({ font: lv.font || '+mn', size: lv.size || 12, color: lv.color || 'tx2' }, ctx), null), {}, 'MP'); return tag('draw:frame', { 'presentation:style-name': `${mname}-backgroundobjects`, 'svg:x': L(d.x), 'svg:y': L(d.y), 'svg:width': L(d.w), 'svg:height': L(d.h), 'presentation:class': cls }, tag('draw:text-box', null, tag('text:p', { 'text:style-name': ps }, inner))); };
    foot = ff('dt', 'date-time', '<presentation:date-time/>') + ff('ftr', 'footer', '<presentation:footer/>') + ff('sldNum', 'page-number', '<text:page-number>&lt;number&gt;</text:page-number>');
  }
  return tag('style:master-page', { 'style:name': mname, 'style:display-name': layout.name || layout.kind, 'style:page-layout-name': 'PM1', 'draw:style-name': dp, 'hw:layout': layout.id, 'hw:kind': layout.kind },
    loextTheme(ctx.theme) + decor + phFrame(ph.title ? 'title' : 'ctrTitle', 'title') + phFrame(ph.body ? 'body' : 'subTitle', ph.body ? 'outline' : 'subtitle') + foot
    + tag('presentation:notes', { 'style:page-layout-name': 'PM2' }, null));
}

/* ---------------------------------------------------------------- main */
export async function writeOdp(doc, opts = {}) {
  const ext = opts.ext === 'otp' ? 'otp' : 'odp';
  const ctx = new Ctx(doc, opts);
  const orig = opts.original && isZip(opts.original) ? await JSZip.loadAsync(opts.original) : null;
  const origNames = orig ? Object.keys(orig.files).filter((n) => !orig.files[n].dir) : [];
  ctx.patch = !!orig;
  ctx.hasPart = (p) => { const q = String(p).replace(/^\.\//, ''); return origNames.some((n) => n === q || n.startsWith(q.endsWith('/') ? q : `${q}/`)); };
  const origStyles = orig && orig.file('styles.xml') ? await orig.file('styles.xml').async('string') : null;
  ctx.keepStyles = !!(origStyles && doc.odfSig && doc.odfSig === odfSignature(doc));
  if (orig) for (const n of doc.odfFonts || []) ctx.fonts.add(n);
  const cst = new Styles(''); // content.xml automatic styles
  const sst = new Styles('M'); // styles.xml automatic styles
  const layouts = doc.layouts && doc.layouts.length ? doc.layouts : [{ id: 'layout-blank', kind: 'blank', name: 'Blank', ph: {}, decor: [] }];
  const layoutOf = (s) => layouts.find((l) => l.id === s.layoutId) || layouts.find((l) => l.kind === 'blank') || layouts[0];
  const used = [];
  for (const s of doc.slides) { const l = layoutOf(s); if (!used.includes(l)) used.push(l); }
  if (!used.length) used.push(layouts[0]);
  if (ctx.keepStyles && used.some((l) => !l.master)) ctx.keepStyles = false;
  // element ids first (animations reference them)
  const reg = (els) => { for (const e of els || []) { ctx.id(e); if (e.children) reg(e.children); } };
  for (const s of doc.slides) reg(s.elements);

  const masters = used.map((l) => masterXml(l, ctx, sst)).join('');
  const pres = used.map((l) => presStyles(l, ctx.mname(l), ctx)).join('');

  const f = doc.footer || {};
  let decls = '';
  if (f.text) decls += tag('presentation:footer-decl', { 'presentation:name': 'ftr1' }, esc(f.text));
  if (f.showDate) decls += tag('presentation:date-time-decl', { 'presentation:name': 'dtd1', 'presentation:source': f.date ? 'fixed' : 'current' }, f.date ? esc(f.date) : null);

  const pages = doc.slides.map((s, i) => {
    const l = layoutOf(s);
    const dp = pageStyle(s, ctx, cst, l.kind === 'title');
    const els = (s.elements || []).map((e) => writeEl(e, ctx, cst, ctx.mname(l))).join('');
    const extra = ctx.extras ? { 'hw:data': json({ id: s.id, layoutId: s.layoutId, transition: s.transition, showMaster: s.showMaster, timing: s.timing ?? undefined, background: s.background }) } : {};
    return tag('draw:page', { 'draw:name': `page${i + 1}`, 'draw:style-name': dp, 'draw:master-page-name': ctx.mname(l), 'presentation:use-footer-name': f.text ? 'ftr1' : undefined, 'presentation:use-date-time-name': f.showDate ? 'dtd1' : undefined, ...extra },
      els + commentsXml(s, ctx) + (s.timingRaw && s.timingRaw.fmt === 'odf' && s.animSig === animSig(s.animations) ? s.timingRaw.xml : animXml(s, ctx)) + notesXml(s, i, ctx, cst));
  }).join('');

  const show = doc.show || {};
  const settings = tag('presentation:settings', { 'presentation:endless': show.loop ? 'true' : undefined, 'presentation:pause': show.loop ? 'PT0S' : undefined, 'presentation:force-manual': show.useTimings === false ? 'true' : undefined }, null);
  const fontDecls = () => tag('office:font-face-decls', null, [...ctx.fonts].map((n) => tag('style:font-face', { 'style:name': n, 'svg:font-family': n.includes(' ') ? `'${n}'` : n }, null)).join('') || null);

  const content = XML_HEAD + tag('office:document-content', { ...(doc.odfNs || {}), ...NS, 'office:version': '1.3' },
    `${fontDecls()}${tag('office:automatic-styles', null, cst.xml.join('') || null)}${tag('office:body', null, tag('office:presentation', null, decls + pages + settings))}`);

  const W = doc.size.w; const H = doc.size.h;
  const defaultText = textProps({ font: '+mn', size: 18, color: 'tx1' }, ctx);
  const styles = XML_HEAD + tag('office:document-styles', { ...(doc.odfNs || {}), ...NS, 'office:version': '1.3' },
    fontDecls()
    + tag('office:styles', null, ctx.commonXml.join('')
      + tag('style:default-style', { 'style:family': 'graphic' }, tag('style:graphic-properties', { 'svg:stroke-color': ctx.hex('accent1-25'), 'draw:fill-color': ctx.hex('accent1') }, null) + tag('style:paragraph-properties', { 'style:writing-mode': 'lr-tb' }, null) + tag('style:text-properties', defaultText, null))
      + pres)
    + tag('office:automatic-styles', null, tag('style:page-layout', { 'style:name': 'PM1' }, tag('style:page-layout-properties', { 'fo:margin-top': '0cm', 'fo:margin-bottom': '0cm', 'fo:margin-left': '0cm', 'fo:margin-right': '0cm', 'fo:page-width': L(W), 'fo:page-height': L(H), 'style:print-orientation': W >= H ? 'landscape' : 'portrait' }, null))
      + tag('style:page-layout', { 'style:name': 'PM2' }, tag('style:page-layout-properties', { 'fo:margin-top': '0cm', 'fo:margin-bottom': '0cm', 'fo:margin-left': '0cm', 'fo:margin-right': '0cm', 'fo:page-width': '21cm', 'fo:page-height': '29.7cm', 'style:print-orientation': 'portrait' }, null))
      + sst.xml.join(''))
    + tag('office:master-styles', null, masters));
  // fonts may have been added while building styles.xml – rebuild content's font decls
  const contentFinal = content.replace(/<office:font-face-decls[\s\S]*?<\/office:font-face-decls>|<office:font-face-decls\/>/, fontDecls());

  const m = doc.meta || {};
  const metaJson = ctx.extras ? json({ v: 1, theme: doc.theme, layouts: layouts.map((l) => ({ ...l, decor: undefined })), footer: doc.footer, show: doc.show, sections: doc.sections, builtin: BUILTIN.has(doc.theme && doc.theme.id) }) : null;
  const meta = XML_HEAD + tag('office:document-meta', { ...(doc.odfNs || {}), ...NS, 'office:version': '1.3' }, tag('office:meta', null,
    tag('meta:generator', null, 'Hostwerk Office')
    + (m.title ? tag('dc:title', null, esc(m.title)) : '')
    + (m.author ? tag('meta:initial-creator', null, esc(m.author)) + tag('dc:creator', null, esc(m.author)) : '')
    + (m.created ? tag('meta:creation-date', null, esc(String(m.created).replace(/Z$/, ''))) : '')
    + tag('dc:date', null, esc(String(m.modified || new Date().toISOString()).replace(/Z$/, '')))
    + tag('meta:document-statistic', { 'meta:object-count': doc.slides.length }, null)
    + (metaJson ? tag('meta:user-defined', { 'meta:name': 'hostwerk' }, esc(metaJson)) : '')));

  const zip = new JSZip();
  zip.file('mimetype', MIMETYPES[ext], { compression: 'STORE' });
  const entries = [['content.xml', 'text/xml'], ['styles.xml', 'text/xml'], ['meta.xml', 'text/xml'], ['settings.xml', 'text/xml']];
  for (const fl of ctx.files.values()) entries.push([fl.path, fl.mime]);
  if (orig) {
    // patch: every part we do not generate goes back byte-identical (Objects, Basic, Configurations2, Thumbnails, settings …)
    const gen = new Set(['mimetype', 'META-INF/manifest.xml', 'content.xml', 'meta.xml', ...(ctx.keepStyles ? [] : ['styles.xml']), ...[...ctx.files.values()].map((x) => x.path)]);
    const om = orig.file('META-INF/manifest.xml') ? parseManifest(await orig.file('META-INF/manifest.xml').async('string')) : [];
    for (const [p, t] of om) if (p !== '/' && !entries.some((x) => x[0] === p)) entries.push([p, t]);
    for (const n of origNames) {
      if (gen.has(n)) continue;
      zip.file(n, await orig.file(n).async('uint8array'), { compression: /\.(png|jpe?g|gif|mp4|webm|mp3|m4a)$/i.test(n) ? 'STORE' : 'DEFLATE' });
      if (!entries.some((x) => x[0] === n) && !/^META-INF\//.test(n)) entries.push([n, '']);
    }
  }
  const manifest = XML_HEAD + tag('manifest:manifest', { 'xmlns:manifest': 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0', 'manifest:version': '1.3' },
    tag('manifest:file-entry', { 'manifest:full-path': '/', 'manifest:version': '1.3', 'manifest:media-type': MIMETYPES[ext] }, null)
    + entries.map(([p, t]) => tag('manifest:file-entry', { 'manifest:full-path': p, 'manifest:media-type': t }, null)).join(''));
  zip.file('META-INF/manifest.xml', manifest);
  zip.file('content.xml', contentFinal);
  zip.file('styles.xml', ctx.keepStyles ? injectCommon(origStyles, ctx.commonXml) : styles);
  zip.file('meta.xml', meta);
  if (!zip.file('settings.xml')) zip.file('settings.xml', XML_HEAD + tag('office:document-settings', { ...(doc.odfNs || {}), ...NS, 'office:version': '1.3' }, '<office:settings/>'));
  for (const fl of ctx.files.values()) zip.file(fl.path, fl.bytes, { compression: /^image\/(png|jpeg|gif|webp)|^video|^audio/.test(fl.mime) ? 'STORE' : 'DEFLATE' });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: MIMETYPES[ext] });
}

function parseManifest(x) {
  const out = []; const re = /<manifest:file-entry\b([^>]*)>/g; let m;
  while ((m = re.exec(x))) { const p = /manifest:full-path="([^"]*)"/.exec(m[1]); const t = /manifest:media-type="([^"]*)"/.exec(m[1]); if (p) out.push([p[1].replace(/&amp;/g, '&'), t ? t[1] : '']); }
  return out;
}
/** add shared styles (gradients, markers …) the kept styles.xml lacks – unchanged bytes when nothing is missing */
function injectCommon(xml, common) {
  const missing = common.filter((c) => { const n = /draw:name="([^"]*)"/.exec(c); return n && !xml.includes(`draw:name="${n[1]}"`); });
  if (!missing.length) return xml;
  if (xml.includes('</office:styles>')) return xml.replace('</office:styles>', `${missing.join('')}</office:styles>`);
  if (xml.includes('<office:styles/>')) return xml.replace('<office:styles/>', `<office:styles>${missing.join('')}</office:styles>`);
  return xml.replace(/(<office:document-styles[^>]*>)/, `$1<office:styles>${missing.join('')}</office:styles>`);
}
/** LibreOffice 7.6+ theme (loext) so the colours show up as theme colours there */
function loextTheme(theme) {
  if (!theme || !theme.colors) return '';
  const names = { dk1: 'dark1', lt1: 'light1', dk2: 'dark2', lt2: 'light2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hlink: 'hyperlink', folHlink: 'followed-hyperlink' };
  return tag('loext:theme', { 'loext:name': theme.name || 'Theme' }, tag('loext:theme-colors', { 'loext:name': theme.name || 'Theme' }, Object.entries(names).map(([k, n]) => tag('loext:color', { 'loext:name': n, 'loext:color': (theme.colors[k] || '#000000').slice(0, 7).toLowerCase() }, null)).join('')));
}
