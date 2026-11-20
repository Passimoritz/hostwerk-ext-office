/*
 * OpenDocument presentation reader (odp/otp zip and flat fodp). readOdp(bytes, { ext }) → { doc, fidelity: [keys] }.
 * Handles LibreOffice- and PowerPoint-written files; files written by odp-write.js carry hw:data extras for an exact
 * round trip (theme, layouts, placeholder level styles …).
 */
import JSZip from 'jszip';
import { parseXml, kids, kid, all, text as xmlText, serialize } from '../../_shared/xmltree.js';
import { lengthToEmu } from '../units.js';
import { themeById, themeData, buildLayouts, THEMES, parseHex } from '../themes.js';
import { uid, bbox, DEFAULT_INSETS } from '../model.js';
import { ODF_TO_GEOM, resolvePrst, SHAPES } from '../render/geometry.js';
import { pathsOfOdf } from '../render/custom.js';
import { bytesToDataUrl, mimeOfPath, isZip } from './util.js';
import { imageSize, odfSignature, animSig } from './odp-write.js';

/* ---------------------------------------------------------------- preserved (raw) elements */
const STYLE_ATTR = /(^|:)(style-name|text-style-name|list-style-name)$/;
/**
 * Mark an element we cannot model as preserved: its original XML (with its automatic styles renamed o_… so they do
 * not clash with the writer's), the package parts it references and its original box. odp-write emits it unchanged
 * (only svg:x/y/width/height/transform are patched when it was moved or resized).
 */
function preserve(el, e, ctx) {
  if (!ctx.zip || ctx.scope !== 'content') return e;
  const rename = {}; const styles = [];
  const visit = (n) => {
    for (const [k, v] of Object.entries(n.a)) {
      if (!STYLE_ATTR.test(k) || rename[v] !== undefined) continue;
      const hit = Object.keys(ctx.db.auto.content).filter((key) => key.endsWith(`|${v}`)).map((key) => ctx.db.auto.content[key]);
      const list = ctx.db.lists.content[v];
      if (!hit.length && !list) continue;
      const nn = v.startsWith('o_') ? v : `o_${v}`;
      rename[v] = nn;
      for (const st of [...hit, ...(list ? [list] : [])]) { visit(st); styles.push(st); }
    }
    for (const c of n.c || []) if (typeof c !== 'string') visit(c);
  };
  visit(el);
  const clone = (n) => {
    if (typeof n === 'string') return n;
    const a = {};
    for (const [k, v] of Object.entries(n.a)) a[k] = (STYLE_ATTR.test(k) || k === 'style:parent-style-name' || k === 'style:name') && rename[v] ? rename[v] : v;
    return { n: n.n, a, c: (n.c || []).map(clone) };
  };
  const parts = new Set();
  const hrefs = (n) => { if (n.a['xlink:href']) { const h = n.a['xlink:href'].replace(/^\.\//, ''); if (h && !/^[a-z][a-z0-9+.-]*:/i.test(h) && !h.startsWith('#') && !h.startsWith('../')) parts.add(h); } for (const c of n.c || []) if (typeof c !== 'string') hrefs(c); };
  hrefs(el);
  e.preserved = true; e.rawFmt = 'odf';
  e.raw = serialize(clone(el));
  e.rawStyles = styles.map((st) => serialize(clone(st)));
  e.rawParts = [...parts];
  e.rawBox = { x: e.x, y: e.y, w: e.w, h: e.h, rot: e.rot || 0 };
  return e;
}

const desc = (el, out = []) => { for (const c of (el && el.c) || []) if (typeof c !== 'string') { out.push(c); desc(c, out); } return out; };
const len = (v) => (v == null || v === '' ? 0 : lengthToEmu(v));
const ptOf = (v) => { if (v == null || /%$/.test(v)) return undefined; const m = /^\s*([\d.]+)\s*pt\s*$/.exec(v); return m ? +m[1] : +(lengthToEmu(v) / 12700).toFixed(2) || undefined; };
const pctOf = (v) => { const m = /^\s*(-?[\d.]+)\s*%/.exec(v || ''); return m ? +m[1] : null; };
const hexOf = (v) => (v && /^#[0-9a-f]{6}$/i.test(v) ? v.toUpperCase() : null);
const parseJson = (s) => { try { return s ? JSON.parse(s) : null; } catch { return null; } };
const secsOf = (v) => {
  if (!v) return 0;
  const iso = /^P(?:T)?(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/i.exec(v);
  if (iso && /^P/i.test(v)) return (+(iso[1] || 0)) * 3600 + (+(iso[2] || 0)) * 60 + (+(iso[3] || 0));
  const m = /^\s*(-?[\d.]+)\s*(ms|s|min|h)?\s*$/.exec(v); if (!m) return 0;
  return +m[1] * ({ ms: 0.001, s: 1, min: 60, h: 3600 }[m[2] || 's']);
};
const PROP_KIND = { 'style:graphic-properties': 'graphic', 'style:paragraph-properties': 'paragraph', 'style:text-properties': 'text', 'style:drawing-page-properties': 'page', 'style:table-cell-properties': 'cell', 'style:table-column-properties': 'column', 'style:table-row-properties': 'row' };
const NUM_BACK = { '1.': 'arabicPeriod', '1)': 'arabicParenR', '(1)': 'arabicParenBoth', 'a.': 'alphaLcPeriod', 'A.': 'alphaUcPeriod', 'a)': 'alphaLcParenR', 'A)': 'alphaUcParenR', 'I.': 'romanUcPeriod', 'i.': 'romanLcPeriod', 'I)': 'romanUcPeriod', 'i)': 'romanLcPeriod' };
const TRANS_BACK = { fade: 'fade', crossfade: 'fade', pushWipe: 'push', barWipe: 'wipe', slideWipe: 'cover', irisWipe: 'zoom', ellipseWipe: 'zoom', zoom: 'zoom', barnDoorWipe: 'split', dissolve: 'fade', randomBarWipe: 'wipe', blindsWipe: 'wipe', boxWipe: 'zoom', fourBoxWipe: 'split', clockWipe: 'wipe', fanWipe: 'wipe' };
const FROM_BACK = { fromRight: 'l', fromLeft: 'r', fromBottom: 'u', fromTop: 'd' };
const SUB_DIR = { 'from-bottom': 'b', 'from-top': 't', 'from-left': 'l', 'from-right': 'r' };

/* ---------------------------------------------------------------- styles */
class StyleDb {
  constructor() { this.common = {}; this.auto = { content: {}, styles: {} }; this.defaults = {}; this.lists = { content: {}, styles: {}, common: {} }; this.gradients = {}; this.fillImages = {}; this.markers = {}; this.dashes = {}; this.fonts = {}; this.cache = new Map(); }
  addFonts(el) { for (const f of kids(el, 'style:font-face')) this.fonts[f.a['style:name']] = (f.a['svg:font-family'] || f.a['style:name'] || '').replace(/^['"]|['"]$/g, ''); }
  addStyles(el, scope) {
    if (!el) return;
    const target = scope === 'common' ? this.common : this.auto[scope];
    for (const s of el.c) {
      if (typeof s === 'string') continue;
      if (s.n === 'style:style') target[`${s.a['style:family']}|${s.a['style:name']}`] = s;
      else if (s.n === 'style:default-style') this.defaults[s.a['style:family']] = s;
      else if (s.n === 'text:list-style') this.lists[scope][s.a['style:name']] = s;
      else if (s.n === 'draw:gradient') this.gradients[s.a['draw:name']] = s;
      else if (s.n === 'draw:fill-image') this.fillImages[s.a['draw:name']] = s.a['xlink:href'] || s;
      else if (s.n === 'draw:marker') this.markers[s.a['draw:name']] = s;
      else if (s.n === 'draw:stroke-dash') this.dashes[s.a['draw:name']] = s;
    }
  }
  list(name, scope) { return (name && (this.lists[scope][name] || this.lists.common[name] || this.lists.styles[name] || this.lists.content[name])) || null; }
  props(s) {
    const o = {};
    for (const c of kids(s)) { const k = PROP_KIND[c.n]; if (!k) continue; o[k] = { ...c.a }; if (k === 'graphic') { const ls = kid(c, 'text:list-style'); if (ls) o.listStyle = ls; } }
    if (s.a['style:list-style-name']) o.listName = s.a['style:list-style-name'];
    return o;
  }
  /** merged properties of a style (parent chain + default style of the family) */
  get(name, family, scope = 'content', withDefault = true) {
    const key = `${scope}|${family}|${name}|${withDefault}`;
    if (this.cache.has(key)) return this.cache.get(key);
    let out = {};
    const s = name && (this.auto[scope][`${family}|${name}`] || this.common[`${family}|${name}`] || (scope === 'content' && this.auto.styles[`${family}|${name}`]));
    if (s) {
      const parent = s.a['style:parent-style-name'];
      const base = parent ? this.get(parent, family, scope, withDefault) : withDefault && this.defaults[family] ? this.props(this.defaults[family]) : {};
      const own = this.props(s);
      out = merge(base, own);
    } else if (withDefault && this.defaults[family]) out = this.props(this.defaults[family]);
    this.cache.set(key, out);
    return out;
  }
}
function merge(a, b) {
  const o = { ...a };
  for (const k of Object.keys(b)) o[k] = k === 'listStyle' || k === 'listName' ? b[k] : { ...(a[k] || {}), ...b[k] };
  return o;
}

/* ---------------------------------------------------------------- property parsers */
function runProps(t, db) {
  if (!t) return {};
  const o = {};
  const fn = t['style:font-name'] ? db.fonts[t['style:font-name']] || t['style:font-name'] : t['fo:font-family'] ? t['fo:font-family'].split(',')[0].replace(/^\s*['"]|['"]\s*$/g, '').trim() : null;
  if (fn) o.font = fn;
  const sz = ptOf(t['fo:font-size']); if (sz) o.size = sz;
  const fw = t['fo:font-weight']; if (fw) o.b = fw === 'bold' || +fw >= 600;
  const fs = t['fo:font-style']; if (fs) o.i = fs === 'italic' || fs === 'oblique';
  const us = t['style:text-underline-style']; if (us) o.u = us !== 'none';
  const ls = t['style:text-line-through-style'] || t['style:text-line-through-type']; if (ls) o.s = ls !== 'none';
  const c = hexOf(t['fo:color']); if (c) o.color = c;
  const bg = hexOf(t['fo:background-color']); if (bg) o.highlight = bg;
  const tp = t['style:text-position'];
  if (tp) { const v = /^super/.test(tp) ? 30 : /^sub/.test(tp) ? -25 : pctOf(tp); if (v) o.baseline = v; }
  if (t['fo:text-transform'] === 'uppercase' || t['fo:font-variant'] === 'small-caps') o.caps = true;
  return o;
}
function fillOf(g, db, files, kind) {
  if (!g) return undefined;
  const f = g['draw:fill'];
  const alpha = pctOf(g['draw:opacity']);
  const withA = (c) => (alpha != null && alpha < 100 ? `${c}@${Math.round(alpha)}` : c);
  if (f === 'none') return null;
  if (f === 'solid' || (!f && g['draw:fill-color'] && kind === 'shape')) return { type: 'solid', color: withA(hexOf(g['draw:fill-color']) || '#729FCF') };
  if (f === 'gradient') {
    const gr = db.gradients[g['draw:fill-gradient-name']];
    if (!gr) return { type: 'solid', color: withA(hexOf(g['draw:fill-color']) || '#FFFFFF') };
    const a = gr.a; const raw = a['draw:angle'] || '0';
    let ang = /deg$/.test(raw) ? parseFloat(raw) : /rad$/.test(raw) ? (parseFloat(raw) * 180) / Math.PI : parseFloat(raw) / 10;
    ang = ((90 - (ang || 0)) % 360 + 360) % 360;
    const s = hexOf(a['draw:start-color']) || '#FFFFFF'; const e = hexOf(a['draw:end-color']) || '#000000';
    const style = a['draw:style'] || 'linear';
    const cs = (gr.c || []).filter((k) => k && k.n === 'loext:color-stop');
    if (cs.length >= 2) {
      let ms = cs.map((k) => ({ pos: Math.round(parseFloat(k.a['svg:offset'] || '0') * 100), color: withA(hexOf(k.a['loext:color-value']) || '#000000') }));
      const radialish = /radial|ellipsoid/.test(style);
      if (radialish) ms = ms.map((x) => ({ ...x, pos: 100 - x.pos })).reverse();
      return { type: 'grad', angle: Math.round(ang), kind: /radial|ellipsoid|square|rectangular/.test(style) ? 'radial' : 'linear', stops: ms };
    }
    const stops = style === 'axial' ? [{ pos: 0, color: withA(e) }, { pos: 50, color: withA(s) }, { pos: 100, color: withA(e) }] : [{ pos: 0, color: withA(s) }, { pos: 100, color: withA(e) }];
    return { type: 'grad', angle: Math.round(ang), kind: /radial|ellipsoid|square|rectangular/.test(style) ? 'radial' : 'linear', stops: /radial|ellipsoid/.test(style) ? stops.reverse().map((x, i) => ({ ...x, pos: i ? 100 : 0 })) : stops };
  }
  if (f === 'bitmap') {
    const href = db.fillImages[g['draw:fill-image-name']];
    const src = typeof href === 'string' ? files.url(href) : href ? files.binary(href) : null;
    return src ? { type: 'image', src, mode: g['style:repeat'] === 'repeat' ? 'tile' : 'stretch' } : null;
  }
  return undefined;
}
function markerKind(name, db) {
  if (!name) return undefined;
  const n = String(name).replace(/_20_/g, ' ');
  if (/circle|oval/i.test(n)) return 'oval';
  if (/square|diamond/i.test(n)) return 'diamond';
  if (/concave|stealth/i.test(n)) return 'stealth';
  if (/line.*arrow|open|chevron/i.test(n)) return 'arrow';
  return db.markers[name] || /arrow|triangle/i.test(n) ? 'triangle' : 'triangle';
}
function lineOf(g, db, kind) {
  if (!g) return undefined;
  const st = g['draw:stroke'];
  if (st === 'none') return null;
  if (!st && !g['svg:stroke-color'] && kind !== 'line') return undefined;
  const a = pctOf(g['svg:stroke-opacity']);
  const c = hexOf(g['svg:stroke-color']) || '#3465A4';
  const out = { color: a != null && a < 100 ? `${c}@${Math.round(a)}` : c, width: len(g['svg:stroke-width']) || 9525, dash: 'solid' };
  if (st === 'dash') {
    const nm = g['draw:stroke-dash'] || '';
    if (/^hw_/.test(nm)) out.dash = nm.slice(3);
    else { const d = db.dashes[nm]; const da = d ? d.a : {}; out.dash = da['draw:dots2'] ? 'dashDot' : (pctOf(da['draw:dots1-length']) || 0) <= 100 && /%/.test(da['draw:dots1-length'] || '%') ? 'dot' : 'dash'; }
  }
  const h = markerKind(g['draw:marker-start'], db); const t = markerKind(g['draw:marker-end'], db);
  if (h) out.head = h;
  if (t) out.tail = t;
  return out;
}
function shadowOf(g) {
  if (!g || g['draw:shadow'] !== 'visible') return null;
  const x = len(g['draw:shadow-offset-x'] || '0.2cm'); const y = len(g['draw:shadow-offset-y'] || '0.2cm');
  return { color: hexOf(g['draw:shadow-color']) || '#000000', dist: Math.round(Math.hypot(x, y)), dir: Math.round(((Math.atan2(y, x) * 180) / Math.PI + 360) % 360), alpha: pctOf(g['draw:shadow-opacity']) ?? 100, blur: len(g['loext:shadow-blur'] || g['draw:shadow-blur'] || '0') };
}
function areaOf(g) {
  const o = {};
  if (!g) return o;
  const va = g['draw:textarea-vertical-align']; if (va) o.anchor = { top: 't', middle: 'ctr', bottom: 'b' }[va] || 't';
  const p = g['fo:padding'];
  const ins = { l: g['fo:padding-left'] ?? p, t: g['fo:padding-top'] ?? p, r: g['fo:padding-right'] ?? p, b: g['fo:padding-bottom'] ?? p };
  if (Object.values(ins).some((v) => v != null)) o.insets = { l: ins.l != null ? len(ins.l) : DEFAULT_INSETS.l, t: ins.t != null ? len(ins.t) : DEFAULT_INSETS.t, r: ins.r != null ? len(ins.r) : DEFAULT_INSETS.r, b: ins.b != null ? len(ins.b) : DEFAULT_INSETS.b };
  if (g['style:shrink-to-fit'] === 'true' || g['draw:fit-to-size'] === 'shrink-to-fit') o.autofit = 'shrink';
  else if (g['draw:auto-grow-height'] === 'true') o.autofit = 'resize';
  else if (g['draw:auto-grow-height'] === 'false') o.autofit = 'none';
  if (g['fo:wrap-option'] === 'no-wrap') o.wrap = false;
  if (/^tb/.test(g['style:writing-mode'] || '')) o.vert = 'vert';
  return o;
}
function paraPropsOf(p) {
  const o = {};
  if (!p) return o;
  const al = p['fo:text-align']; if (al) o.align = { start: 'l', left: 'l', center: 'ctr', end: 'r', right: 'r', justify: 'just' }[al] || 'l';
  const mt = ptOf(p['fo:margin-top']); if (mt != null && p['fo:margin-top']) o.spcBef = mt || 0;
  const mb = ptOf(p['fo:margin-bottom']); if (mb != null && p['fo:margin-bottom']) o.spcAft = mb || 0;
  const lh = pctOf(p['fo:line-height']); if (lh) o.lnSpc = lh;
  if (p['fo:margin-left']) o.marL = len(p['fo:margin-left']);
  if (p['fo:text-indent']) o.indent = len(p['fo:text-indent']);
  return o;
}
function bulletOf(listEl, level) {
  if (!listEl) return { bullet: { char: '•' } };
  const lv = kids(listEl).find((x) => +x.a['text:level'] === level + 1) || null;
  if (!lv) return { bullet: { char: '•' } };
  const pr = kid(lv, 'style:list-level-properties') || { a: {} };
  const la = kid(pr, 'style:list-level-label-alignment');
  let marL; let indent;
  if (la) { marL = len(la.a['fo:margin-left']); indent = len(la.a['fo:text-indent']); } else if (pr.a['text:space-before'] != null || pr.a['text:min-label-width'] != null) { const sb = len(pr.a['text:space-before']); const mw = len(pr.a['text:min-label-width']); marL = sb + mw; indent = -mw; }
  const tp = kid(lv, 'style:text-properties');
  const color = tp && hexOf(tp.a['fo:color']);
  let bullet;
  if (lv.n === 'text:list-level-style-number') {
    const fmt = lv.a['style:num-format'];
    if (!fmt) bullet = false; else bullet = { num: NUM_BACK[`${lv.a['style:num-prefix'] || ''}${fmt}${lv.a['style:num-suffix'] || ''}`] || NUM_BACK[`${fmt}${lv.a['style:num-suffix'] || '.'}`] || 'arabicPeriod', start: +(lv.a['text:start-value'] || 1) };
  } else if (lv.n === 'text:list-level-style-bullet') {
    const ch = lv.a['text:bullet-char'];
    bullet = ch ? { char: ch } : false;
    if (bullet && tp && tp.a['fo:font-family']) bullet.font = tp.a['fo:font-family'].replace(/['"]/g, '');
  } else bullet = { char: '•' };
  if (bullet && color) bullet.color = color;
  return { bullet, marL, indent };
}

/* ---------------------------------------------------------------- text */
function runsOf(pEl, ctx, base) {
  const runs = [];
  const st = { space: true };
  const push = (t, props) => {
    if (!t) return;
    const last = runs[runs.length - 1];
    if (last && !last.field && !props.field && JSON.stringify({ ...last, text: '' }) === JSON.stringify({ ...props, text: '' })) last.text += t;
    else runs.push({ ...props, text: t });
  };
  const walkN = (el, props) => {
    for (const x of el.c || []) {
      if (typeof x === 'string') {
        let t = x.replace(/[\t\r\n ]+/g, ' ');
        if (st.space) t = t.replace(/^ /, '');
        if (t) { st.space = t.endsWith(' '); push(t, props); }
        continue;
      }
      switch (x.n) {
        case 'text:span': walkN(x, { ...props, ...runProps(ctx.db.get(x.a['text:style-name'], 'text', ctx.scope, false).text, ctx.db) }); break;
        case 'text:a': walkN(x, { ...props, link: x.a['xlink:href'] }); break;
        case 'text:s': push(' '.repeat(+(x.a['text:c'] || 1)), props); st.space = true; break;
        case 'text:tab': push('\t', props); st.space = true; break;
        case 'text:line-break': push('\n', props); st.space = true; break;
        case 'text:page-number': runs.push({ ...props, text: xmlText(x) || '1', field: 'slidenum' }); st.space = false; break;
        case 'text:date': case 'text:time': case 'presentation:date-time': runs.push({ ...props, text: xmlText(x) || '', field: 'datetime' }); st.space = false; break;
        case 'presentation:footer': runs.push({ ...props, text: '', field: 'footer' }); break;
        case 'text:soft-page-break': case 'office:annotation': case 'officeooo:annotation': case 'text:bookmark': case 'text:bookmark-start': case 'text:bookmark-end': break;
        default: walkN(x, props);
      }
    }
  };
  walkN(pEl, base);
  return runs;
}
/** paragraphs of a text container (text-box, shape, cell) */
function parasOf(container, ctx, opt = {}) {
  const out = [];
  const defP = opt.paraStyle;
  const para = (pEl, level, bulletInfo) => {
    const ps = ctx.db.get(pEl.a['text:style-name'] || defP, 'paragraph', ctx.scope, false);
    const runBase = ctx.foreign ? runProps(ps.text, ctx.db) : {};
    const p = { runs: runsOf(pEl, ctx, runBase), level };
    if (ctx.foreign) Object.assign(p, paraPropsOf(ps.paragraph));
    if (bulletInfo) {
      p.bullet = bulletInfo.bullet;
      if (ctx.foreign && bulletInfo.bullet) { if (bulletInfo.marL != null) p.marL = bulletInfo.marL; if (bulletInfo.indent != null) p.indent = bulletInfo.indent; }
    } else p.bullet = false;
    out.push(p);
  };
  const list = (lEl, depth, listEl) => {
    const ls = ctx.db.list(lEl.a['text:style-name'], ctx.scope) || listEl;
    for (const it of kids(lEl)) {
      if (it.n !== 'text:list-item' && it.n !== 'text:list-header') continue;
      let first = true;
      for (const c of kids(it)) {
        if (c.n === 'text:p' || c.n === 'text:h') { para(c, Math.min(8, depth), it.n === 'text:list-item' && first ? bulletOf(ls, Math.min(8, depth)) : { bullet: false }); first = false; } else if (c.n === 'text:list') { list(c, depth + 1, ls); first = false; }
      }
    }
  };
  for (const c of kids(container)) {
    if (c.n === 'text:p' || c.n === 'text:h') para(c, 0, null);
    else if (c.n === 'text:list') list(c, 0, opt.listStyle || null);
  }
  return out;
}

/* ---------------------------------------------------------------- geometry */
function boxOf(el) {
  const w = len(el.a['svg:width']); const h = len(el.a['svg:height']);
  const tr = el.a['draw:transform'];
  if (!tr) return { x: len(el.a['svg:x']), y: len(el.a['svg:y']), w, h, rot: 0 };
  let M = [1, 0, 0, 1, 0, 0];
  const mul = (o, m) => [o[0] * m[0] + o[2] * m[1], o[1] * m[0] + o[3] * m[1], o[0] * m[2] + o[2] * m[3], o[1] * m[2] + o[3] * m[3], o[0] * m[4] + o[2] * m[5] + o[4], o[1] * m[4] + o[3] * m[5] + o[5]];
  const re = /(rotate|translate|scale|skewX|skewY|matrix)\s*\(([^)]*)\)/g;
  let m;
  while ((m = re.exec(tr))) {
    const args = m[2].trim().split(/[\s,]+/);
    let op = [1, 0, 0, 1, 0, 0];
    if (m[1] === 'rotate') { const t = parseFloat(args[0]) || 0; op = [Math.cos(t), -Math.sin(t), Math.sin(t), Math.cos(t), 0, 0]; }
    else if (m[1] === 'translate') op = [1, 0, 0, 1, len(args[0]), len(args[1] || '0')];
    else if (m[1] === 'scale') op = [parseFloat(args[0]) || 1, 0, 0, parseFloat(args[1] || args[0]) || 1, 0, 0];
    else if (m[1] === 'matrix') op = [+args[0], +args[1], +args[2], +args[3], len(args[4]), len(args[5])];
    M = mul(op, M);
  }
  const cx = M[0] * (w / 2) + M[2] * (h / 2) + M[4]; const cy = M[1] * (w / 2) + M[3] * (h / 2) + M[5];
  let rot = (Math.atan2(M[1], M[0]) * 180) / Math.PI;
  rot = Math.round(((rot % 360) + 360) % 360 * 100) / 100;
  if (rot > 359.99) rot = 0;
  return { x: Math.round(cx - w / 2), y: Math.round(cy - h / 2), w, h, rot };
}
function rotateAround(e, cx, cy, deg) {
  const a = (deg * Math.PI) / 180;
  const ex = e.x + e.w / 2 - cx; const ey = e.y + e.h / 2 - cy;
  const dx = cx + ex * Math.cos(a) - ey * Math.sin(a) - (e.x + e.w / 2); const dy = cy + ex * Math.sin(a) + ey * Math.cos(a) - (e.y + e.h / 2);
  const mv = (x) => ({ ...x, x: Math.round(x.x + dx), y: Math.round(x.y + dy), ...(x.children ? { children: x.children.map(mv) } : {}) });
  return { ...mv(e), rot: Math.round(((((e.rot || 0) + deg) % 360) + 360) % 360 * 100) / 100 };
}

/* ---------------------------------------------------------------- elements */
const PH_OF_CLASS = { title: 'title', subtitle: 'subTitle', outline: 'body', graphic: 'pic', 'page-number': 'sldNum', footer: 'ftr', 'date-time': 'dt' };
function elId(el) { return el.a['xml:id'] || el.a['draw:id'] || uid('e'); }

function textBody(el, ctx, g, x, opt = {}) {
  const area = areaOf(g);
  const b = { paras: [], anchor: area.anchor || opt.anchor || 't', insets: area.insets || { ...DEFAULT_INSETS }, autofit: area.autofit || 'none', wrap: area.wrap !== false };
  if (area.vert) b.vert = area.vert;
  b.paras = parasOf(opt.container || el, ctx, { paraStyle: el.a['draw:text-style-name'], listStyle: opt.listStyle });
  if (!b.paras.length) b.paras = [{ runs: [], level: 0 }];
  if (x && x.text) {
    const t = x.text;
    for (const k of ['lvl', 'style', 'autofit', 'fontScale', 'anchor', 'insets']) if (t[k] !== undefined) b[k] = t[k];
    b.paras.forEach((p, i) => { Object.assign(p, (t.paraProps && t.paraProps[i]) || {}); if (t.bullets && t.bullets[i] === 0) delete p.bullet; });
  } else if (opt.lvl) b.lvl = opt.lvl;
  else if (opt.style && Object.keys(opt.style).length) b.style = opt.style;
  return b;
}
const applyExtras = (e, x) => { if (!x) return e; for (const k of ['ph', 'phIdx', 'name', 'locked', 'decor', 'prompt', 'textbox', 'link', 'flipH', 'flipV', 'geom', 'adj', 'shadow', 'crop', 'paths']) if (x[k] !== undefined) e[k] = x[k]; if (x.lineCap && e.line) e.line = { ...e.line, cap: x.lineCap, join: x.lineJoin }; if (x.fill) e.fill = x.fill; if (x.line && e.type === 'line') e.line = x.line; return e; };

/** placeholder level styles from the master's presentation styles (foreign files) */
function phLevels(ctx, cls, ownStyle) {
  const db = ctx.db; const m = ctx.master;
  const own = ownStyle ? runProps(db.get(ownStyle, 'presentation', ctx.scope, false).text, db) : {};
  const out = [];
  for (let i = 0; i < 9; i++) {
    const nm = cls === 'outline' ? `${m}-outline${i + 1}` : `${m}-${cls}`;
    const s = db.get(nm, 'presentation', 'styles');
    const r = { ...runProps(db.get(null, 'graphic', 'styles').text, db), ...runProps(s.text, db) };
    const pp = paraPropsOf(s.paragraph);
    const lv = { font: '+mn', size: 18, color: 'tx1', bullet: false, marL: i * 457200, indent: 0, ...r, ...(i === 0 ? own : {}), ...(pp.align ? { align: pp.align } : {}), ...(pp.spcBef != null ? { spcBef: pp.spcBef } : {}) };
    if (cls === 'outline') { const ls = db.get(`${m}-outline1`, 'presentation', 'styles').listStyle; const bi = bulletOf(ls, i); lv.bullet = bi.bullet; if (bi.marL != null) lv.marL = bi.marL; if (bi.indent != null) lv.indent = bi.indent; }
    out.push(lv);
  }
  return out;
}

function parseEl(el, ctx) {
  try {
    const x = parseJson(el.a['hw:data']);
    const styleName = el.a['draw:style-name'] || el.a['presentation:style-name'];
    const fam = el.a['presentation:style-name'] && !el.a['draw:style-name'] ? 'presentation' : 'graphic';
    const st = ctx.db.get(styleName, fam, ctx.scope);
    const g = st.graphic || {};
    const id = elId(el);
    const common = { id, name: el.a['draw:name'] || '', rot: 0 };
    if (el.n === 'draw:g') {
      let children = kids(el).map((c) => parseEl(c, ctx)).filter(Boolean);
      if (!children.length) return null;
      let box = bbox(children);
      if (x && x.group) { box = x.group; if (box.rot) children = children.map((c) => rotateAround(c, box.x + box.w / 2, box.y + box.h / 2, -box.rot)); }
      return applyExtras({ ...common, type: 'group', x: Math.round(box.x), y: Math.round(box.y), w: Math.round(box.w), h: Math.round(box.h), rot: (x && x.group && x.group.rot) || 0, children }, x);
    }
    if (el.n === 'draw:line' || el.n === 'draw:connector') {
      const x1 = len(el.a['svg:x1']); const y1 = len(el.a['svg:y1']); const x2 = len(el.a['svg:x2']); const y2 = len(el.a['svg:y2']);
      const t = el.a['draw:type'];
      return applyExtras({ ...common, type: 'line', geom: el.n === 'draw:connector' && t === 'standard' ? 'elbow' : t === 'curve' ? 'curve' : 'line', x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH: x2 < x1 || undefined, flipV: y2 < y1 || undefined, line: lineOf(g, ctx.db, 'line') || { color: '#000000', width: 9525, dash: 'solid' }, shadow: shadowOf(g) }, x);
    }
    const box = boxOf(el);
    const cls = el.a['presentation:class'];
    if (el.n === 'draw:frame') {
      if (x && x.chart) return { ...common, ...box, type: 'chart', chart: x.chart, name: x.name || '' };
      const tb = kid(el, 'draw:text-box'); const imgs = kids(el, 'draw:image'); const tbl = kid(el, 'table:table'); const plug = kid(el, 'draw:plugin');
      const obj = kid(el, 'draw:object') || kid(el, 'draw:object-ole');
      if (obj) {
        const isChart = ctx.files.isChart(obj.a['xlink:href']);
        ctx.fidelity.add(isChart ? 'charts' : 'ole');
        const pick = imgs.find((i) => /\.(png|jpe?g|gif|svg|webp|bmp)$/i.test(i.a['xlink:href'] || '')) || imgs[0];
        const src = pick ? (pick.a['xlink:href'] ? ctx.files.url(pick.a['xlink:href']) : ctx.files.binary(pick)) : null;
        return preserve(el, { ...common, ...box, type: 'unknown', what: isChart ? 'Chart' : 'Embedded object', src: src || null }, ctx);
      }
      if (tbl) return parseTable(tbl, ctx, common, box, x);
      if (plug) { const src = ctx.files.url(plug.a['xlink:href']); if (!src) { ctx.fidelity.add('media'); return null; } const mime = plug.a['draw:mime-type'] || mimeOfPath(plug.a['xlink:href']); return applyExtras({ ...common, ...box, type: 'media', kind: x && x.media ? x.media : /^audio/.test(mime) ? 'audio' : 'video', src, mime }, x); }
      if (imgs.length && !tb) {
        const pick = imgs.find((i) => /\.(png|jpe?g|gif|svg|webp|bmp)$/i.test(i.a['xlink:href'] || '')) || imgs[0];
        const src = pick.a['xlink:href'] ? ctx.files.url(pick.a['xlink:href']) : ctx.files.binary(pick);
        if (!src) return null;
        if (x && x.unknown) return { ...common, ...box, type: 'unknown', what: x.unknown, src };
        const e = { ...common, ...box, type: 'image', src, crop: { l: 0, t: 0, r: 0, b: 0 }, opacity: pctOf(g['draw:image-opacity']) != null ? pctOf(g['draw:image-opacity']) / 100 : 1, line: lineOf(g, ctx.db, 'image') || null, shadow: shadowOf(g), alt: xmlText(kid(el, 'svg:desc')) || xmlText(kid(el, 'svg:title')) || '' };
        const clip = /rect\(([^)]*)\)/.exec(g['fo:clip'] || '');
        if (clip) {
          const sz = ctx.files.size(pick.a['xlink:href']);
          const [t, r, b, l] = clip[1].split(/[\s,]+/).map(len);
          if (sz && (t || r || b || l)) { const W = sz.w * 9525; const H = sz.h * 9525; e.crop = { t: t / H, r: r / W, b: b / H, l: l / W }; }
        }
        if (cls === 'graphic') e.ph = 'pic';
        return applyExtras(e, x);
      }
      if (tb || cls) {
        const e = { ...common, ...box, type: 'shape', geom: 'rect', textbox: true, fill: fillOf(g, ctx.db, ctx.files, 'frame') || null, line: lineOf(g, ctx.db, 'frame') || null, shadow: shadowOf(g) };
        const phKey = cls ? PH_OF_CLASS[cls] : null;
        let opt = {};
        if (phKey && ctx.foreign && cls !== 'graphic') opt = { lvl: phLevels(ctx, cls === 'outline' ? 'outline' : cls === 'subtitle' ? 'subtitle' : cls === 'title' ? 'title' : 'outline', el.a['presentation:style-name']), listStyle: cls === 'outline' ? st.listStyle : null };
        else if (ctx.foreign) opt = { style: runProps(st.text, ctx.db), listStyle: st.listStyle };
        e.text = textBody(el, ctx, g, x, { ...opt, container: tb || { c: [] } });
        if (phKey) { e.ph = phKey; if (el.a['presentation:placeholder'] === 'true') e.prompt = ''; }
        return applyExtras(e, x);
      }
      return preserve(el, { ...common, ...box, type: 'unknown', what: 'Object', src: null }, ctx);
    }
    if (['draw:custom-shape', 'draw:rect', 'draw:ellipse', 'draw:circle', 'draw:polygon', 'draw:polyline', 'draw:path', 'draw:regular-polygon', 'draw:caption'].includes(el.n)) {
      let geom = 'rect'; let adj; let keepRaw = false; let customPathsOf = null;
      let flipH; let flipV;
      if (el.n === 'draw:custom-shape') {
        const eg = kid(el, 'draw:enhanced-geometry') || { a: {} };
        const t = eg.a['draw:type'] || '';
        if (ODF_TO_GEOM[t]) geom = ODF_TO_GEOM[t];
        else if (/^ooxml-/.test(t)) geom = resolvePrst(t.slice(6));
        else if (!(x && x.geom)) {
          const cp = pathsOfOdf(eg.a['draw:enhanced-path'], eg.a['svg:viewBox']);
          if (cp) { geom = 'custom'; customPathsOf = cp; } else { ctx.fidelity.add('geometry'); keepRaw = true; }
        }
        if (eg.a['draw:mirror-horizontal'] === 'true') flipH = true;
        if (eg.a['draw:mirror-vertical'] === 'true') flipV = true;
      } else if (el.n === 'draw:ellipse' || el.n === 'draw:circle') geom = 'ellipse';
      else if (el.n === 'draw:rect' && el.a['draw:corner-radius'] && len(el.a['draw:corner-radius']) > 0) { geom = 'roundRect'; adj = [Math.round((len(el.a['draw:corner-radius']) / Math.max(1, Math.min(box.w, box.h))) * 100000)]; }
      else if (el.n !== 'draw:rect') { ctx.fidelity.add('geometry'); keepRaw = true; }
      if (!el.a['svg:width'] && el.a['svg:viewBox']) { /* polygon without size – keep viewBox size */ }
      const e = { ...common, ...box, type: 'shape', geom: SHAPES[geom] || geom === 'custom' ? geom : 'rect', ...(customPathsOf ? { paths: customPathsOf } : {}), fill: fillOf(g, ctx.db, ctx.files, 'shape'), line: lineOf(g, ctx.db, 'shape'), shadow: shadowOf(g) };
      if (e.fill === undefined) e.fill = { type: 'solid', color: '#729FCF' };
      if (e.line === undefined) e.line = { color: '#3465A4', width: 9525, dash: 'solid' };
      if (adj) e.adj = adj;
      if (flipH) e.flipH = true;
      if (flipV) e.flipV = true;
      e.text = textBody(el, ctx, g, x, ctx.foreign ? { anchor: 'ctr', style: runProps(st.text, ctx.db), listStyle: st.listStyle } : { anchor: 'ctr' });
      if (ctx.foreign && !g['draw:textarea-vertical-align']) e.text.anchor = 'ctr';
      if (ctx.foreign && e.text.paras.every((p) => !p.align && !(p.runs || []).length)) e.text.paras.forEach((p) => { p.align = 'ctr'; });
      applyExtras(e, x);
      return keepRaw ? preserve(el, e, ctx) : e;
    }
    if (el.n === 'dr3d:scene' || el.n === 'draw:control' || el.n === 'draw:measure' || el.n === 'draw:page-thumbnail') {
      ctx.fidelity.add(el.n === 'dr3d:scene' ? 'model3d' : 'ole');
      return preserve(el, { ...common, ...box, type: 'unknown', what: el.n === 'dr3d:scene' ? '3D model' : 'Object', src: null }, ctx);
    }
    return null;
  } catch {
    return null;
  }
}

function parseTable(tbl, ctx, common, box, x) {
  const cols = [];
  for (const c of all(tbl, 'table:table-column')) {
    const w = len(ctx.db.get(c.a['table:style-name'], 'table-column', ctx.scope, false).column?.['style:column-width']);
    for (let i = 0; i < +(c.a['table:number-columns-repeated'] || 1); i++) cols.push(w);
  }
  const rowEls = all(tbl, 'table:table-row');
  const rows = rowEls.map((r) => {
    const rh = len(ctx.db.get(r.a['table:style-name'], 'table-row', ctx.scope, false).row?.['style:row-height']);
    const cells = [];
    for (const c of kids(r)) {
      if (c.n === 'table:covered-table-cell') { for (let i = 0; i < +(c.a['table:number-columns-repeated'] || 1); i++) cells.push({ text: { paras: [{ runs: [], level: 0 }], anchor: 'ctr', insets: { ...DEFAULT_INSETS }, autofit: 'none', wrap: true }, merged: true }); continue; }
      if (c.n !== 'table:table-cell') continue;
      const cs = ctx.db.get(c.a['table:style-name'], 'table-cell', ctx.scope);
      const g = cs.graphic || {}; const cp = { ...(cs.paragraph || {}), ...(cs.cell || {}) };
      const fill = fillOf(g, ctx.db, ctx.files, 'cell') ?? (hexOf(cp['fo:background-color']) ? { type: 'solid', color: hexOf(cp['fo:background-color']) } : null);
      const b = textBody(c, ctx, g, null, ctx.foreign ? { anchor: { top: 't', middle: 'ctr', bottom: 'b' }[cp['style:vertical-align']] || 't', style: runProps(cs.text, ctx.db) } : { anchor: 'ctr' });
      const bd = {};
      for (const [k, s] of [['l', 'left'], ['r', 'right'], ['t', 'top'], ['b', 'bottom']]) { const v = cp[`fo:border-${s}`] || cp['fo:border']; if (v) bd[k] = borderOf(v); }
      const cell = { text: b, fill };
      if (Object.keys(bd).length) cell.borders = bd;
      const cs2 = +(c.a['table:number-columns-spanned'] || 1); const rs = +(c.a['table:number-rows-spanned'] || 1);
      if (cs2 > 1 || rs > 1) cell.span = { c: cs2, r: rs };
      cells.push(cell);
    }
    return { h: rh, cells };
  });
  const n = Math.max(cols.length, ...rows.map((r) => r.cells.length));
  while (cols.length < n) cols.push(0);
  const cw = cols.map((w) => w || Math.round(box.w / n));
  rows.forEach((r) => { if (!r.h) r.h = Math.round(box.h / Math.max(1, rows.length)); while (r.cells.length < n) r.cells.push({ text: { paras: [{ runs: [], level: 0 }], anchor: 'ctr', insets: { ...DEFAULT_INSETS }, autofit: 'none', wrap: true }, fill: null }); });
  // covered cells from row spans in later rows
  const e = { ...common, ...box, type: 'table', cols: cw, rows, style: { firstRow: tbl.a['table:use-first-row-styles'] === 'true', bandRows: tbl.a['table:use-banding-rows-styles'] === 'true', firstCol: false, accent: 'accent1' }, border: null };
  const first = rows[0] && rows[0].cells.find((c) => c.borders);
  e.border = first ? first.borders.l || first.borders.t || null : null;
  if (x && x.table) {
    e.style = x.table.style || e.style; e.border = x.table.border !== undefined ? x.table.border : e.border;
    rows.forEach((r, ri) => r.cells.forEach((c, ci) => { if (x.table.implicit && x.table.implicit[ri] && x.table.implicit[ri][ci]) delete c.fill; const same = (a, b) => JSON.stringify(a) === JSON.stringify(b); if (c.borders && ['l', 'r', 't', 'b'].every((k) => same(c.borders[k], e.border))) delete c.borders; }));
  }
  return applyExtras(e, x);
}
function borderOf(v) {
  if (!v || v === 'none') return null;
  const w = /(-?[\d.]+(?:cm|mm|in|pt|px))/.exec(v); const c = /#[0-9a-f]{6}/i.exec(v);
  return { color: c ? c[0].toUpperCase() : '#000000', width: w ? len(w[1]) : 9525, dash: /dash|dot/.test(v) ? 'dash' : 'solid' };
}

/* ---------------------------------------------------------------- pages */
function parsePage(pg, ctx) {
  const dp = ctx.db.get(pg.a['draw:style-name'], 'drawing-page', ctx.scope, false).page || {};
  const px = parseJson(pg.a['hw:data']);
  const slide = { id: (px && px.id) || uid('s'), layoutId: null, hidden: dp['presentation:visibility'] === 'hidden', background: null, elements: [], notes: '', transition: { type: 'none', dir: 'l', duration: 500, advanceAfter: null }, animations: [], comments: [] };
  if (dp['draw:fill'] && dp['draw:fill'] !== 'none') slide.background = fillOf(dp, ctx.db, ctx.files, 'page') || null;
  if (dp['presentation:background-objects-visible'] === 'false') slide.showMaster = false;
  const ty = dp['smil:type'] || dp['presentation:transition-style'];
  if (ty) {
    const type = TRANS_BACK[ty] || (/fade/.test(ty) ? 'fade' : /wipe|uncover|cover/.test(ty) ? 'wipe' : null);
    if (type) {
      const sub = dp['smil:subtype'] || ''; const rev = dp['smil:direction'] === 'reverse';
      let dir = FROM_BACK[sub] || 'l';
      if (type === 'wipe') dir = /topToBottom/i.test(sub) ? (rev ? 'u' : 'd') : (rev ? 'l' : 'r');
      slide.transition = { type, dir, duration: { fast: 500, medium: 1000, slow: 2000 }[dp['presentation:transition-speed']] || 700, advanceAfter: null };
    } else ctx.fidelity.add('transitions');
  }
  if (dp['presentation:transition-type'] === 'automatic') slide.transition.advanceAfter = Math.round(secsOf(dp['presentation:duration']) * 1000);
  ctx.footerFlags = ctx.footerFlags || { f: false, n: false, d: false };
  if (dp['presentation:display-footer'] === 'true') ctx.footerFlags.f = true;
  if (dp['presentation:display-page-number'] === 'true') ctx.footerFlags.n = true;
  if (dp['presentation:display-date-time'] === 'true') ctx.footerFlags.d = true;
  ctx.master = pg.a['draw:master-page-name'] || '';
  const els = []; let bodyN = 0;
  for (const c of kids(pg)) {
    if (c.n === 'presentation:notes') {
      const nf = all(c, 'draw:frame').find((f) => f.a['presentation:class'] === 'notes');
      const tb = nf && kid(nf, 'draw:text-box');
      if (tb) slide.notes = parasOf(tb, { ...ctx, foreign: true }).map((p) => p.runs.map((r) => r.text).join('')).join('\n').replace(/\s+$/, '');
      continue;
    }
    if (c.n === 'anim:par' || c.n === 'anim:seq') { slide.animations = parseAnims(c, ctx); slide.timingRaw = { fmt: 'odf', xml: serialize(c) }; continue; }
    if (c.n === 'officeooo:annotation' || c.n === 'office:annotation') {
      const t = parasOf(c, { ...ctx, foreign: true }).map((p) => p.runs.map((r) => r.text).join('')).join('\n');
      slide.comments.push({ id: uid('c'), author: xmlText(kid(c, 'dc:creator')), initials: xmlText(kid(c, 'meta:creator-initials')) || undefined, time: xmlText(kid(c, 'dc:date')), text: t, x: len(c.a['svg:x']), y: len(c.a['svg:y']) });
      if (!slide.comments[slide.comments.length - 1].initials) delete slide.comments[slide.comments.length - 1].initials;
      continue;
    }
    if (c.n === 'office:forms') continue;
    const e = parseEl(c, ctx);
    if (!e) continue;
    if (e.ph === 'body' && ctx.foreign) { bodyN++; if (bodyN > 1) e.ph = `body${bodyN}`; }
    els.push(e);
  }
  if (ctx.foreign && els.some((e) => e.ph === 'subTitle')) els.forEach((e) => { if (e.ph === 'title') e.ph = 'ctrTitle'; });
  slide.elements = els;
  // ids of animation targets
  const ids = new Set(); const w = (l) => l.forEach((e) => { ids.add(e.id); if (e.children) w(e.children); }); w(els);
  slide.animations = slide.animations.filter((a) => ids.has(a.el));
  if (slide.timingRaw) slide.animSig = animSig(slide.animations);
  if (px) { if (px.comments && !slide.comments.length) slide.comments = px.comments;
    for (const k of ['transition', 'showMaster', 'timing', 'background']) if (px[k] !== undefined) slide[k] = px[k]; slide.layoutId = px.layoutId || null; }
  slide.master = ctx.master;
  return slide;
}

function parseAnims(root, ctx) {
  const out = [];
  for (const p of all({ c: [root] }, 'anim:par')) {
    const pid = p.a['presentation:preset-id'];
    if (!pid) continue;
    const pcls = p.a['presentation:preset-class'] || 'entrance';
    let hwA = null; try { hwA = p.a['hw:anim'] ? JSON.parse(p.a['hw:anim']) : null; } catch { hwA = null; }
    const cls = hwA && hwA.cls ? hwA.cls : pcls === 'exit' ? 'exit' : pcls === 'emphasis' ? 'emph' : pcls === 'motion-path' ? 'path' : pcls === 'entrance' ? 'entr' : null;
    if (!cls) { ctx.fidelity.add('animations'); continue; }
    let effect = hwA && hwA.effect ? hwA.effect : /appear|disappear/.test(pid) ? 'appear' : /fade|dissolve/.test(pid) ? 'fade' : /fly|ascend|descend|rise/.test(pid) ? 'fly' : /zoom|grow|expand/.test(pid) ? 'zoom' : /wipe|wheel|split|blinds/.test(pid) ? 'wipe' : null;
    if (!hwA && cls === 'emph') effect = /spin/.test(pid) ? 'spin' : /grow|shrink/.test(pid) ? 'grow' : /transparen/.test(pid) ? 'transparency' : null;
    if (!hwA && cls === 'path') effect = 'custom';
    if (!effect) { ctx.fidelity.add('animations'); if (cls !== 'entr') continue; }
    const targetEl = desc(p).find((c) => c.a['smil:targetElement']);
    const tgt = targetEl && targetEl.a['smil:targetElement'];
    if (!tgt) continue;
    let dur = 0; for (const c of desc(p)) { const d = secsOf(c.a['smil:dur']); if (d > 0.01) dur = Math.max(dur, d); }
    const nt = p.a['presentation:node-type'];
    const extra = {};
    if (cls !== 'entr') extra.cls = cls;
    if (hwA) for (const k of ['path', 'scale', 'angle', 'opacity']) if (hwA[k] != null) extra[k] = hwA[k];
    if (cls === 'path' && !extra.path) { const m = desc(p).find((c) => c.n === 'anim:animateMotion'); extra.path = `${(m && m.a['svg:path']) || 'M 0 0 L 0.25 0'} E`; }
    out.push({ id: uid('a'), el: tgt, ...extra, effect: effect || 'appear', dir: p.a['hw:dir'] || SUB_DIR[p.a['presentation:preset-sub-type']] || (effect === 'fly' || effect === 'wipe' ? 'b' : undefined), trigger: nt === 'with-previous' ? 'with' : nt === 'after-previous' ? 'after' : 'click', duration: effect === 'appear' ? 0 : Math.round((dur || 0.5) * 1000), delay: Math.round(secsOf(p.a['smil:begin']) * 1000) });
    if (!out[out.length - 1].dir) delete out[out.length - 1].dir;
  }
  return out;
}

/* ---------------------------------------------------------------- files */
function fileStore(zip, cache) {
  const norm = (h) => String(h || '').replace(/^\.\//, '').replace(/^\//, '');
  return {
    url: (h) => { const p = norm(h); return cache[p] ? bytesToDataUrl(cache[p], mimeOfPath(p)) : null; },
    size: (h) => { const p = norm(h); return cache[p] ? imageSize(cache[p]) : null; },
    binary: (el) => { const b = kid(el, 'office:binary-data'); if (!b) return null; const t = xmlText(b).replace(/\s+/g, ''); const head = atob(t.slice(0, 16)); const mime = head.startsWith('\x89PNG') ? 'image/png' : head.startsWith('\xff\xd8') ? 'image/jpeg' : head.startsWith('GIF') ? 'image/gif' : 'image/png'; return `data:${mime};base64,${t}`; },
    isChart: (h) => !!(zip && cache.__charts && cache.__charts.has(norm(h))),
  };
}

/* ---------------------------------------------------------------- main */
export async function readOdp(bytes, { ext } = {}) { // eslint-disable-line no-unused-vars
  const fidelity = new Set();
  const db = new StyleDb();
  let content; let styles; let meta;
  const cache = {};
  let zip = null;
  if (isZip(bytes)) {
    zip = await JSZip.loadAsync(bytes);
    const read = async (p) => { const f = zip.file(p); return f ? f.async('string') : null; };
    content = parseXml(await read('content.xml') || '<office:document-content/>');
    const s = await read('styles.xml'); styles = s ? parseXml(s) : null;
    const m = await read('meta.xml'); meta = m ? parseXml(m) : null;
    cache.__charts = new Set();
    for (const name of Object.keys(zip.files)) {
      const f = zip.files[name]; if (f.dir) continue;
      if (/^(Pictures|Media|media|Images)\//i.test(name) || /\.(png|jpe?g|gif|svg|bmp|webp|mp4|webm|mp3|wav|ogg|m4a|mov)$/i.test(name)) cache[name] = await f.async('uint8array');
      if (/\/content\.xml$/.test(name) && name !== 'content.xml') { const t = await f.async('string'); if (/<office:chart/.test(t)) cache.__charts.add(name.replace(/\/content\.xml$/, '')); }
    }
    if (zip.file('Basic/script-lc.xml') || Object.keys(zip.files).some((n) => /^Basic\/.+\.xml$/.test(n) && !/script-lc|dialog-lc/.test(n))) fidelity.add('macros');
  } else {
    const root = parseXml(typeof bytes === 'string' ? bytes : new TextDecoder().decode(bytes));
    content = root; styles = root; meta = root;
  }
  const flat = content === styles;
  for (const src of [styles, content]) { if (!src) continue; const fd = kid(src, 'office:font-face-decls'); if (fd) db.addFonts(fd); }
  if (styles) { db.addStyles(kid(styles, 'office:styles'), 'common'); db.addStyles(kid(styles, 'office:automatic-styles'), flat ? 'content' : 'styles'); }
  if (!flat) db.addStyles(kid(content, 'office:automatic-styles'), 'content');
  if (flat) db.auto.styles = db.auto.content;
  const files = fileStore(zip, cache);
  const metaEl = meta && kid(meta, 'office:meta');
  const ud = metaEl && kids(metaEl, 'meta:user-defined').find((u) => u.a['meta:name'] === 'hostwerk');
  const hw = ud ? parseJson(xmlText(ud)) : null;
  const ctx = { db, files, fidelity, scope: 'content', foreign: !hw, master: '', zip: !!zip };

  // size from the first master's page layout
  const masterStyles = styles && kid(styles, 'office:master-styles');
  const masterEls = masterStyles ? kids(masterStyles, 'style:master-page') : [];
  let size = { w: 10080000, h: 7560000 };
  const plName = masterEls[0] && masterEls[0].a['style:page-layout-name'];
  const autoS = styles && kid(styles, 'office:automatic-styles');
  const pl = autoS && kids(autoS, 'style:page-layout').find((p) => p.a['style:name'] === plName);
  const plp = pl && kid(pl, 'style:page-layout-properties');
  if (plp && plp.a['fo:page-width']) size = { w: len(plp.a['fo:page-width']), h: len(plp.a['fo:page-height']) };

  // masters → decor, background, placeholder geometry
  const masters = {};
  for (const mp of masterEls) {
    const mctx = { ...ctx, scope: flat ? 'content' : 'styles', master: mp.a['style:name'], foreign: !hw };
    const dp = db.get(mp.a['draw:style-name'], 'drawing-page', mctx.scope, false).page || {};
    const decor = []; const ph = {};
    let themeColors = null;
    for (const c of kids(mp)) {
      if (c.n === 'loext:theme') { const tc = all(c, 'loext:color'); if (tc.length) { themeColors = {}; const map = { dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', hyperlink: 'hlink', 'followed-hyperlink': 'folHlink' }; for (const x of tc) themeColors[map[x.a['loext:name']] || x.a['loext:name']] = (x.a['loext:color'] || '').toUpperCase(); } continue; }
      if (c.n === 'presentation:notes') continue;
      if (c.a['presentation:class']) { const b = boxOf(c); ph[c.a['presentation:class']] = b; continue; }
      const e = parseEl(c, mctx);
      if (e) decor.push({ ...e, locked: true, decor: true });
    }
    masters[mp.a['style:name']] = { name: mp.a['style:display-name'] || mp.a['style:name'], layoutId: mp.a['hw:layout'], kind: mp.a['hw:kind'], bg: dp['draw:fill'] && dp['draw:fill'] !== 'none' ? fillOf(dp, db, files, 'page') : null, decor, ph, themeColors, title: runProps(db.get(`${mp.a['style:name']}-title`, 'presentation', mctx.scope).text, db), body: { ...runProps(db.get(null, 'graphic', mctx.scope).text, db), ...runProps(db.get(`${mp.a['style:name']}-outline1`, 'presentation', mctx.scope).text, db) } };
  }

  const body = kid(content, 'office:body');
  const presEl = body && kid(body, 'office:presentation');
  if (!presEl) throw new Error('This file has no slides (not an OpenDocument presentation).');
  const footers = {}; const dates = {};
  for (const d of kids(presEl, 'presentation:footer-decl')) footers[d.a['presentation:name']] = xmlText(d);
  for (const d of kids(presEl, 'presentation:date-time-decl')) dates[d.a['presentation:name']] = d.a['presentation:source'] === 'fixed' ? xmlText(d) : '';
  const slides = []; let footerText = ''; let footerDate = '';
  for (const pg of kids(presEl, 'draw:page')) {
    const s = parsePage(pg, ctx);
    if (pg.a['presentation:use-footer-name'] && footers[pg.a['presentation:use-footer-name']]) footerText = footers[pg.a['presentation:use-footer-name']];
    if (pg.a['presentation:use-date-time-name'] && dates[pg.a['presentation:use-date-time-name']]) footerDate = dates[pg.a['presentation:use-date-time-name']];
    slides.push(s);
  }
  const settings = kid(presEl, 'presentation:settings');

  // theme + layouts
  let theme; let layouts;
  if (hw && hw.theme) {
    theme = hw.theme;
    if (hw.builtin && THEMES.some((t) => t.id === theme.id)) layouts = buildLayouts(themeById(theme.id), size).map((l) => { const mn = Object.keys(masters).find((k) => masters[k].layoutId === l.id); return mn ? { ...l, master: mn } : l; });
    else {
      layouts = (hw.layouts || []).map((l) => { const m = Object.values(masters).find((x) => x.layoutId === l.id); const mn = Object.keys(masters).find((k) => masters[k].layoutId === l.id); return { ...l, decor: m ? m.decor : [], ...(mn ? { master: mn } : {}) }; });
      if (!layouts.length) layouts = buildLayouts(themeById('office'), size);
    }
  } else {
    const counts = {}; for (const s of slides) counts[s.master] = (counts[s.master] || 0) + 1;
    const main = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0] || Object.keys(masters)[0];
    const mm = masters[main] || { name: 'Imported', ph: {}, decor: [], title: {}, body: {} };
    const td = themeData(themeById('office'));
    theme = { ...td, id: 'imported', name: mm.name || 'Imported', fonts: { major: mm.title.font || td.fonts.major, minor: mm.body.font || mm.title.font || td.fonts.minor } };
    if (mm.themeColors) theme.colors = { ...theme.colors, ...mm.themeColors };
    const bgHex = mm.bg && mm.bg.type === 'solid' ? mm.bg.color : mm.bg && mm.bg.type === 'grad' ? mm.bg.stops[0].color : null;
    const rgb = bgHex && parseHex(bgHex.slice(0, 7));
    if (rgb && (0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b) / 255 < 0.4) theme.dark = true;
    if (mm.bg) theme.bg = mm.bg;
    const base = buildLayouts(theme, size);
    layouts = [];
    const mk = (mname, kind) => {
      const id = `${mname}|${kind}`;
      let l = layouts.find((x) => x.id === id);
      if (l) return l;
      const m = masters[mname] || mm;
      const b = base.find((x) => x.kind === kind) || base[0];
      l = { ...JSON.parse(JSON.stringify(b)), id, name: `${m.name} – ${b.name}`, decor: m.decor || [], bg: m.bg || null, master: masters[mname] ? mname : undefined };
      if (!l.master) delete l.master;
      const tph = m.ph && m.ph.title; const oph = m.ph && m.ph.outline;
      if (tph && tph.w) for (const k of ['titleContent', 'titleOnly', 'twoContent', 'comparison']) if (kind === k && l.ph.title) Object.assign(l.ph.title, { x: tph.x, y: tph.y, w: tph.w, h: tph.h });
      if (oph && oph.w && kind === 'titleContent' && l.ph.body) Object.assign(l.ph.body, { x: oph.x, y: oph.y, w: oph.w, h: oph.h });
      layouts.push(l);
      return l;
    };
    for (const k of ['title', 'titleContent', 'section', 'twoContent', 'comparison', 'titleOnly', 'blank', 'pictureCaption']) mk(main, k);
    for (const s of slides) {
      const keys = new Set(s.elements.map((e) => e.ph).filter(Boolean));
      const kind = keys.has('subTitle') || keys.has('ctrTitle') ? 'title' : keys.has('body4') ? 'comparison' : keys.has('body2') ? 'twoContent' : keys.has('pic') && keys.has('title') ? 'pictureCaption' : keys.has('body') && keys.has('title') ? 'titleContent' : keys.has('title') ? 'titleOnly' : keys.has('body') ? 'titleContent' : 'blank';
      s.layoutId = mk(s.master || main, kind).id;
    }
  }
  for (const s of slides) {
    if (!s.layoutId || !layouts.some((l) => l.id === s.layoutId)) { const m = masters[s.master]; s.layoutId = (m && m.layoutId && layouts.some((l) => l.id === m.layoutId) ? m.layoutId : null) || (layouts.find((l) => l.kind === (m && m.kind)) || layouts.find((l) => l.kind === 'titleContent') || layouts[0]).id; }
    delete s.master;
  }

  const mtxt = (n) => { const e = metaEl && kid(metaEl, n); return e ? xmlText(e) : ''; };
  const ff = ctx.footerFlags || {};
  const doc = {
    size,
    theme,
    layouts,
    slides,
    sections: (hw && hw.sections) || [],
    footer: (hw && hw.footer) || { text: footerText, showFooter: !!(ff.f && footerText), showNumber: !!ff.n, showDate: !!ff.d, date: footerDate, hideOnTitle: true },
    show: (hw && hw.show) || { loop: !!(settings && settings.a['presentation:endless'] === 'true'), autoAdvance: 0, useTimings: !(settings && settings.a['presentation:force-manual'] === 'true') },
    meta: { title: mtxt('dc:title'), author: mtxt('dc:creator') || mtxt('meta:initial-creator'), created: mtxt('meta:creation-date'), modified: mtxt('dc:date') },
  };
  if (zip) {
    // what the writer needs to patch the original package (see writeOdp(doc, { original }))
    doc.odfNs = Object.fromEntries(Object.entries(content.a || {}).filter(([k]) => k.startsWith('xmlns:')));
    const fd = kid(content, 'office:font-face-decls');
    doc.odfFonts = fd ? kids(fd, 'style:font-face').map((x) => x.a['style:name']).filter(Boolean) : [];
    if (styles) doc.odfSig = odfSignature(doc);
  }
  return { doc, fidelity: [...fidelity] };
}
