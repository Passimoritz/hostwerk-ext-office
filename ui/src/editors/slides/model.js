/*
 * Document model of the slides chunk – plain JSON, immutable updates, pure functions (runs in Node for tests).
 * All lengths in EMU (see units.js), angles in degrees, font sizes in pt, colours as refs (see themes.js).
 *
 * Doc = {
 *   size: { w, h },                          slide size
 *   theme: { id, name, dark, colors{dk1…folHlink}, fonts{major,minor}, bg: Fill, clrMap? }
 *   layouts: [Layout]                        see themes.buildLayouts – { id, kind, name, bg, decor:[El], ph:{ key: PhDef } }
 *   slides: [Slide],
 *   sections: [{ id, name, start: slideId }] a section starts at that slide and runs to the next section
 *   footer: { text, showFooter, showNumber, showDate, date: '' (fixed text, empty = today), hideOnTitle: true }
 *   show: { loop, autoAdvance (s, 0 = off), useTimings }
 *   meta: { title, author, created, modified }
 * }
 * Slide = { id, layoutId, hidden, background: Fill|null (null = layout/theme), elements: [El], notes: '' (plain text),
 *           transition: { type: 'none'|'fade'|'push'|'wipe'|'zoom'|'cover'|'split', dir: 'l'|'r'|'u'|'d', duration (ms), advanceAfter (ms|null) },
 *           animations: [{ id, el (element id), effect: 'appear'|'fade'|'fly'|'zoom'|'wipe', dir, trigger: 'click'|'with'|'after', duration (ms), delay (ms) }],
 *           showMaster (bool, default true), timing (rehearsed ms|null), comments: [{ id, author, text, time }] }
 * El (common) = { id, type, name, x, y, w, h, rot, flipH, flipV, locked, hidden, ph?: placeholder key, phIdx?, link? }
 *   shape  { type:'shape', geom (see render/geometry.js), adj?: [n], fill: Fill|null, line: Line|null, shadow: Shadow|null, text?: Body, textbox?: bool }
 *   line   { type:'line', geom:'line'|'elbow'|'curve', line: Line (with head/tail) }   – drawn from (x,y) to (x+w,y+h), flipH/flipV mirror
 *   image  { type:'image', src (data URL), crop: { l,t,r,b } (fractions 0..1), opacity (0..1), line, shadow, alt }
 *   table  { type:'table', cols: [EMU], rows: [{ h, cells: [Cell] }], style: { firstRow, bandRows, firstCol, accent }, border: Line }
 *          Cell = { text: Body, fill: Fill|null|undefined (undefined = from table style), span: { c, r }, merged: bool (covered by a span), borders?: { l,t,r,b: Line|null } }
 *   chart  { type:'chart', chart: { kind: 'bar'|'col'|'line'|'pie'|'doughnut'|'area', categories: [], series: [{ name, values: [], color }], title, legend, labels, stacked } }
 *   group  { type:'group', children: [El] }  – children in absolute slide coordinates; x,y,w,h = bounding box
 *   media  { type:'media', kind: 'video'|'audio', src (data URL), mime, poster? }
 *   unknown{ type:'unknown', what: 'SmartArt'|'OLE object'|…, src? (preview image) } – kept as a picture when possible
 * Body = { paras: [Para], anchor: 't'|'ctr'|'b', insets: { l,t,r,b }, autofit: 'none'|'shrink'|'resize', fontScale?: 0..1,
 *          wrap: true, vert?: 'vert'|'vert270', lvl?: [9 level styles], style?: { run defaults } }
 * Para = { runs: [Run], level, align?, bullet?: false | { char, color?, font? } | { num: 'arabicPeriod'|'alphaLcPeriod'|'alphaUcPeriod'|'romanUcPeriod'|'romanLcPeriod'|'arabicParenR'|'alphaLcParenR', start },
 *          lnSpc? (percent), spcBef? (pt), spcAft? (pt), marL?, indent? (EMU) }
 * Run  = { text, b, i, u, s, font, size (pt), color, highlight, baseline (+30 super / -25 sub), link, field?: 'slidenum'|'datetime'|'footer', caps? }
 * Fill = null | { type:'solid', color } | { type:'grad', angle (deg), kind?: 'linear'|'radial', stops: [{ pos 0..100, color }] } | { type:'image', src, mode: 'stretch'|'tile' }
 * Line = null | { color, width (EMU), dash: 'solid'|'dash'|'dot'|'dashDot'|'lgDash'|'sysDash'|'sysDot', head?, tail? ('none'|'triangle'|'arrow'|'stealth'|'oval'|'diamond') }
 * Shadow = null | { color, blur (EMU), dist (EMU), dir (deg), alpha (0..100) }
 */
import { SIZES } from './units.js';
import { buildLayouts, themeById, themeData, THEMES } from './themes.js';

let seq = 0;
export const uid = (p = 'e') => `${p}${Date.now().toString(36).slice(-4)}${(seq++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;
const clone = (x) => (x == null ? x : JSON.parse(JSON.stringify(x)));
export const deepClone = clone;

export const DEFAULT_INSETS = { l: 91440, t: 45720, r: 91440, b: 45720 };

/* ---------------------------------------------------------------- text */
export const run = (text, props = {}) => ({ text: String(text ?? ''), ...props });
export const para = (text, props = {}, runProps = {}) => ({ runs: text === '' || text == null ? [] : [run(text, runProps)], level: 0, ...props });
export function body(paras = [], props = {}) {
  return { paras: paras.length ? paras : [para('')], anchor: 't', insets: { ...DEFAULT_INSETS }, autofit: 'none', wrap: true, ...props };
}
/** text body from plain text – lines become paragraphs, leading tabs become levels */
export function bodyFromText(str, props = {}, paraProps = {}, runProps = {}) {
  const lines = String(str ?? '').split(/\r?\n/);
  return body(lines.map((l) => { const m = /^(\t*)(.*)$/.exec(l); return para(m[2], { ...paraProps, level: Math.min(8, m[1].length) }, runProps); }), props);
}
export const paraText = (p) => (p.runs || []).map((r) => r.text || '').join('');
export function bodyText(b) {
  if (!b || !b.paras) return '';
  return b.paras.map(paraText).join('\n');
}
export const isEmptyBody = (b) => !b || !b.paras || b.paras.every((p) => !paraText(p));

/** apply run props to every run of a body (and as defaults to empty paragraphs) */
export function mapRuns(b, fn) {
  if (!b) return b;
  return { ...b, paras: b.paras.map((p) => ({ ...p, runs: p.runs.length ? p.runs.map(fn) : p.runs, endProps: fn({ ...(p.endProps || {}), text: '' }) })) };
}
export const setRunProps = (b, props) => mapRuns(b, (r) => cleanRun({ ...r, ...props }));
export const mapParas = (b, fn) => (b ? { ...b, paras: b.paras.map(fn) } : b);
function cleanRun(r) { const o = {}; for (const k of Object.keys(r)) if (r[k] !== undefined && r[k] !== null) o[k] = r[k]; if (o.text === undefined) o.text = ''; return o; }

/* ---------------------------------------------------------------- elements */
const base = (type, x, y, w, h, extra) => ({ id: uid(type[0]), type, name: '', x: Math.round(x), y: Math.round(y), w: Math.max(1, Math.round(w)), h: Math.max(1, Math.round(h)), rot: 0, ...extra });
export const textBox = (x, y, w, h, b, extra = {}) => base('shape', x, y, w, h, { geom: 'rect', textbox: true, fill: null, line: null, text: b || body([para('')], { autofit: 'resize' }), ...extra });
export const shape = (geom, x, y, w, h, extra = {}) => base('shape', x, y, w, h, {
  geom, fill: { type: 'solid', color: 'accent1' }, line: { color: 'accent1-25', width: 12700, dash: 'solid' }, shadow: null,
  text: body([para('', { align: 'ctr' })], { anchor: 'ctr' }), ...extra,
});
export const line = (x, y, w, h, extra = {}) => base('line', x, y, w, h, { geom: 'line', line: { color: 'accent1', width: 19050, dash: 'solid', head: 'none', tail: 'none' }, ...extra });
export const image = (src, x, y, w, h, extra = {}) => base('image', x, y, w, h, { src, crop: { l: 0, t: 0, r: 0, b: 0 }, opacity: 1, line: null, shadow: null, alt: '', ...extra });
export const media = (kind, src, mime, x, y, w, h, extra = {}) => base('media', x, y, w, h, { kind, src, mime, ...extra });
export function table(rows, cols, x, y, w, h, extra = {}) {
  const cw = Math.round(w / cols); const rh = Math.round(h / rows);
  return base('table', x, y, w, h, {
    cols: Array.from({ length: cols }, () => cw),
    rows: Array.from({ length: rows }, () => ({ h: rh, cells: Array.from({ length: cols }, () => cell('')) })),
    style: { firstRow: true, bandRows: true, firstCol: false, accent: 'accent1' }, border: { color: 'bg1', width: 12700, dash: 'solid' }, ...extra,
  });
}
export const cell = (text = '', props = {}) => ({ text: body([para(text)], { anchor: 'ctr', insets: { l: 91440, t: 45720, r: 91440, b: 45720 } }), ...props });
export function chart(kind, x, y, w, h, extra = {}) {
  return base('chart', x, y, w, h, {
    chart: { kind, title: '', legend: true, labels: false, stacked: false, categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5], color: 'accent1' }, { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8], color: 'accent2' }, { name: 'Series 3', values: [2, 2, 3, 5], color: 'accent3' }] },
    ...extra,
  });
}

/** recursively visit elements (groups included) */
export function walk(els, fn, parent = null) { for (const e of els || []) { fn(e, parent); if (e.type === 'group') walk(e.children, fn, e); } }
export function findEl(els, id) { let f = null; walk(els, (e) => { if (e.id === id) f = e; }); return f; }

export function bbox(els) {
  if (!els.length) return { x: 0, y: 0, w: 0, h: 0 };
  let x1 = Infinity; let y1 = Infinity; let x2 = -Infinity; let y2 = -Infinity;
  for (const e of els) { const b = rotatedBox(e); x1 = Math.min(x1, b.x); y1 = Math.min(y1, b.y); x2 = Math.max(x2, b.x + b.w); y2 = Math.max(y2, b.y + b.h); }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}
/** axis-aligned box of a rotated element */
export function rotatedBox(e) {
  const r = (((e.rot || 0) % 180) + 180) % 180;
  if (!r) return { x: e.x, y: e.y, w: e.w, h: e.h };
  const a = (r * Math.PI) / 180; const c = Math.abs(Math.cos(a)); const s = Math.abs(Math.sin(a));
  const w = e.w * c + e.h * s; const h = e.w * s + e.h * c;
  return { x: e.x + e.w / 2 - w / 2, y: e.y + e.h / 2 - h / 2, w, h };
}
export function moveEl(e, dx, dy) {
  const n = { ...e, x: Math.round(e.x + dx), y: Math.round(e.y + dy) };
  if (e.type === 'group') n.children = e.children.map((c) => moveEl(c, dx, dy));
  return n;
}
/** map an element from box `from` into box `to` (scales children of groups, tables, …) */
export function fitEl(e, from, to) {
  const sx = from.w ? to.w / from.w : 1; const sy = from.h ? to.h / from.h : 1;
  const n = { ...e, x: Math.round(to.x + (e.x - from.x) * sx), y: Math.round(to.y + (e.y - from.y) * sy), w: Math.max(1, Math.round(e.w * sx)), h: Math.max(1, Math.round(e.h * sy)) };
  if (e.type === 'group') n.children = e.children.map((c) => fitEl(c, from, to));
  if (e.type === 'table') { n.cols = e.cols.map((c) => Math.max(1, Math.round(c * sx))); n.rows = e.rows.map((r) => ({ ...r, h: Math.max(1, Math.round(r.h * sy)) })); }
  return n;
}
export const resizeEl = (e, box) => fitEl(e, { x: e.x, y: e.y, w: e.w, h: e.h }, box);
/** give every element (recursively) fresh ids, mapping animations later via the returned map */
export function reId(els, map = {}) {
  return els.map((e) => { const id = uid(e.type[0]); map[e.id] = id; const n = { ...e, id }; if (e.type === 'group') n.children = reId(e.children, map); return n; });
}

/* ---------------------------------------------------------------- doc */
export function newDoc({ size = '16:9', theme = 'hostwerk', font, title = '', author = '' } = {}) {
  const sz = typeof size === 'object' ? size : (SIZES[size] || SIZES['16:9']);
  const t = themeById(theme);
  const td = themeData(t);
  if (font && font !== 'Calibri' && t.id === 'office') { td.fonts.minor = font; }
  const now = new Date().toISOString();
  return {
    size: { w: sz.w, h: sz.h },
    theme: td,
    layouts: buildLayouts(t, sz),
    slides: [],
    sections: [],
    footer: { text: '', showFooter: false, showNumber: false, showDate: false, date: '', hideOnTitle: true },
    show: { loop: false, autoAdvance: 0, useTimings: true },
    meta: { title, author, created: now, modified: now },
  };
}

export const layoutOf = (doc, slide) => (doc.layouts || []).find((l) => l.id === slide.layoutId) || (doc.layouts || []).find((l) => l.kind === 'blank') || (doc.layouts || [])[0] || { id: 'none', kind: 'blank', ph: {}, decor: [] };
export const layoutByKind = (doc, kind) => (doc.layouts || []).find((l) => l.kind === kind) || (doc.layouts || [])[0];

/** placeholder element from a layout's ph definition */
export function phElement(key, def, text) {
  const isTitle = /title/i.test(key);
  const lvl = def.lvl ? clone(def.lvl) : undefined;
  if (key === 'pic' && !text) {
    return { ...base('shape', def.x, def.y, def.w, def.h, { geom: 'rect', fill: null, line: null }), ph: key, prompt: def.prompt || '', text: body([para('')], { anchor: def.anchor || 'ctr', lvl }), textbox: true };
  }
  const b = text != null && typeof text === 'object' && text.paras ? { ...text, lvl } : bodyFromText(text || '', { anchor: def.anchor || (isTitle ? 'b' : 't'), lvl, autofit: isTitle ? 'none' : 'shrink' });
  return { ...base('shape', def.x, def.y, def.w, def.h, { geom: 'rect', fill: null, line: null }), ph: key, prompt: def.prompt || '', text: b, textbox: true };
}

export function newSlide(doc, kind = 'titleContent', texts = {}) {
  const layout = layoutByKind(doc, kind);
  const elements = [];
  for (const [key, def] of Object.entries((layout && layout.ph) || {})) {
    if (['dt', 'ftr', 'sldNum'].includes(key)) continue;
    elements.push(phElement(key, def, texts[key]));
  }
  return {
    id: uid('s'), layoutId: layout ? layout.id : null, hidden: false, background: null, elements, notes: texts.notes || '',
    transition: { type: 'none', dir: 'l', duration: 500, advanceAfter: null }, animations: [], comments: [],
  };
}

export const slideIndex = (doc, id) => doc.slides.findIndex((s) => s.id === id);
export const getSlide = (doc, id) => doc.slides.find((s) => s.id === id);
export function slideTitle(slide) {
  const t = (slide.elements || []).find((e) => e.ph && /title/i.test(e.ph) && !/sub/i.test(e.ph)) || null;
  if (t) return bodyText(t.text).split('\n')[0].trim();
  let first = '';
  walk(slide.elements, (e) => { if (!first && e.text) first = bodyText(e.text).split('\n')[0].trim(); });
  return first;
}

const touch = (doc) => ({ ...doc, meta: { ...(doc.meta || {}), modified: new Date().toISOString() } });

export function insertSlides(doc, slides, index = doc.slides.length) {
  const list = doc.slides.slice();
  list.splice(Math.max(0, Math.min(index, list.length)), 0, ...slides);
  return touch({ ...doc, slides: list });
}
export const addSlide = (doc, kind, index, texts) => { const s = newSlide(doc, kind, texts); return { doc: insertSlides(doc, [s], index ?? doc.slides.length), slide: s }; };

export function duplicateSlides(doc, ids) {
  const out = [];
  const list = [];
  for (const s of doc.slides) {
    list.push(s);
    if (ids.includes(s.id)) {
      const map = {};
      const n = { ...clone(s), id: uid('s') };
      n.elements = reId(n.elements, map);
      n.animations = (n.animations || []).filter((a) => map[a.el]).map((a) => ({ ...a, id: uid('a'), el: map[a.el] }));
      list.push(n); out.push(n.id);
    }
  }
  return { doc: touch({ ...doc, slides: list }), ids: out };
}
export function deleteSlides(doc, ids) {
  const keep = doc.slides.filter((s) => !ids.includes(s.id));
  // sections that started at a deleted slide move to the next remaining slide
  const sections = [];
  for (const sec of doc.sections || []) {
    let i = slideIndex(doc, sec.start);
    while (i >= 0 && i < doc.slides.length && ids.includes(doc.slides[i].id)) i++;
    const ns = doc.slides[i];
    if (ns && !sections.some((x) => x.start === ns.id)) sections.push({ ...sec, start: ns.id });
  }
  return touch({ ...doc, slides: keep, sections });
}
/** move the slides `ids` (kept in their order) so that they start at position `to` of the remaining list */
export function moveSlides(doc, ids, to) {
  const moving = doc.slides.filter((s) => ids.includes(s.id));
  const before = doc.slides.slice(0, to).filter((s) => !ids.includes(s.id)).length;
  const rest = doc.slides.filter((s) => !ids.includes(s.id));
  rest.splice(before, 0, ...moving);
  return touch({ ...doc, slides: rest });
}
export const updateSlide = (doc, id, fn) => touch({ ...doc, slides: doc.slides.map((s) => (s.id === id ? (typeof fn === 'function' ? fn(s) : { ...s, ...fn }) : s)) });
export const updateSlides = (doc, ids, fn) => touch({ ...doc, slides: doc.slides.map((s) => (ids.includes(s.id) ? (typeof fn === 'function' ? fn(s) : { ...s, ...fn }) : s)) });

/** update elements (searches groups too); fn(el) → el | patch object */
export function updateEls(doc, slideId, ids, fn) {
  const set = new Set(ids);
  const map = (els) => els.map((e) => {
    let n = e;
    if (set.has(e.id)) n = typeof fn === 'function' ? fn(e) : { ...e, ...fn };
    if (n && n.type === 'group' && n === e) { const ch = map(e.children); if (ch.some((c, i) => c !== e.children[i])) n = { ...e, children: ch }; }
    return n;
  });
  return updateSlide(doc, slideId, (s) => ({ ...s, elements: map(s.elements) }));
}
export const addEls = (doc, slideId, els) => updateSlide(doc, slideId, (s) => ({ ...s, elements: [...s.elements, ...els] }));
export function deleteEls(doc, slideId, ids) {
  const set = new Set(ids);
  const filt = (els) => els.filter((e) => !set.has(e.id)).map((e) => (e.type === 'group' ? { ...e, children: filt(e.children) } : e)).filter((e) => e.type !== 'group' || e.children.length);
  return updateSlide(doc, slideId, (s) => ({ ...s, elements: filt(s.elements), animations: (s.animations || []).filter((a) => !set.has(a.el)) }));
}

/** z-order: 'front' | 'back' | 'forward' | 'backward' (top-level elements) */
export function zorder(doc, slideId, ids, op) {
  return updateSlide(doc, slideId, (s) => {
    let list = s.elements.slice();
    const sel = (e) => ids.includes(e.id);
    if (op === 'front') list = [...list.filter((e) => !sel(e)), ...list.filter(sel)];
    else if (op === 'back') list = [...list.filter(sel), ...list.filter((e) => !sel(e))];
    else if (op === 'forward') { for (let i = list.length - 2; i >= 0; i--) if (sel(list[i]) && !sel(list[i + 1])) [list[i], list[i + 1]] = [list[i + 1], list[i]]; }
    else if (op === 'backward') { for (let i = 1; i < list.length; i++) if (sel(list[i]) && !sel(list[i - 1])) [list[i], list[i - 1]] = [list[i - 1], list[i]]; }
    return { ...s, elements: list };
  });
}

export function groupEls(doc, slideId, ids) {
  const s = getSlide(doc, slideId);
  const members = s.elements.filter((e) => ids.includes(e.id));
  if (members.length < 2) return { doc, id: null };
  const b = bbox(members.map((m) => ({ ...m, rot: 0 })));
  const g = { id: uid('g'), type: 'group', name: '', x: Math.round(b.x), y: Math.round(b.y), w: Math.round(b.w), h: Math.round(b.h), rot: 0, children: members };
  const firstIdx = s.elements.findIndex((e) => ids.includes(e.id));
  const rest = s.elements.filter((e) => !ids.includes(e.id));
  const at = s.elements.slice(0, firstIdx).filter((e) => !ids.includes(e.id)).length;
  rest.splice(at, 0, g);
  return { doc: updateSlide(doc, slideId, { elements: rest }), id: g.id };
}
export function ungroupEls(doc, slideId, ids) {
  const out = [];
  const s = getSlide(doc, slideId);
  const list = [];
  for (const e of s.elements) {
    if (ids.includes(e.id) && e.type === 'group') {
      // a rotated group passes its rotation to the children (around the group centre)
      const kids = e.children.map((c) => (e.rot ? rotateAround(c, e.x + e.w / 2, e.y + e.h / 2, e.rot) : c));
      list.push(...kids); out.push(...kids.map((c) => c.id));
    } else list.push(e);
  }
  return { doc: updateSlide(doc, slideId, { elements: list }), ids: out };
}
function rotateAround(e, cx, cy, deg) {
  const a = (deg * Math.PI) / 180;
  const ex = e.x + e.w / 2 - cx; const ey = e.y + e.h / 2 - cy;
  const nx = cx + ex * Math.cos(a) - ey * Math.sin(a); const ny = cy + ex * Math.sin(a) + ey * Math.cos(a);
  return moveEl({ ...e, rot: ((e.rot || 0) + deg) % 360 }, nx - (e.x + e.w / 2), ny - (e.y + e.h / 2));
}

/** align: 'l'|'c'|'r'|'t'|'m'|'b'; to: 'slide' | 'selection' */
export function alignEls(doc, slideId, ids, how, to = 'selection') {
  const s = getSlide(doc, slideId);
  const els = s.elements.filter((e) => ids.includes(e.id));
  if (!els.length) return doc;
  const ref = to === 'slide' || els.length === 1 ? { x: 0, y: 0, w: doc.size.w, h: doc.size.h } : bbox(els);
  return updateEls(doc, slideId, ids, (e) => {
    const b = rotatedBox(e);
    let dx = 0; let dy = 0;
    if (how === 'l') dx = ref.x - b.x; else if (how === 'c') dx = ref.x + ref.w / 2 - (b.x + b.w / 2); else if (how === 'r') dx = ref.x + ref.w - (b.x + b.w);
    else if (how === 't') dy = ref.y - b.y; else if (how === 'm') dy = ref.y + ref.h / 2 - (b.y + b.h / 2); else if (how === 'b') dy = ref.y + ref.h - (b.y + b.h);
    return moveEl(e, dx, dy);
  });
}
/** distribute horizontally ('h') or vertically ('v') with equal gaps */
export function distributeEls(doc, slideId, ids, dir = 'h') {
  const s = getSlide(doc, slideId);
  const els = s.elements.filter((e) => ids.includes(e.id)).map((e) => ({ e, b: rotatedBox(e) }));
  if (els.length < 3) return doc;
  const k = dir === 'h' ? 'x' : 'y'; const d = dir === 'h' ? 'w' : 'h';
  els.sort((a, b) => a.b[k] - b.b[k]);
  const first = els[0].b; const last = els[els.length - 1].b;
  const total = els.reduce((acc, x) => acc + x.b[d], 0);
  const gap = (last[k] + last[d] - first[k] - total) / (els.length - 1);
  let pos = first[k];
  const target = {};
  for (const x of els) { target[x.e.id] = pos - x.b[k]; pos += x.b[d] + gap; }
  return updateEls(doc, slideId, ids, (e) => moveEl(e, dir === 'h' ? target[e.id] : 0, dir === 'v' ? target[e.id] : 0));
}

/** change the slide size, scaling all elements (scale: 'fit' keeps proportions and centres, 'stretch' fills) */
export function setSize(doc, size, how = 'fit') {
  const from = { x: 0, y: 0, w: doc.size.w, h: doc.size.h };
  let to = { x: 0, y: 0, w: size.w, h: size.h };
  if (how === 'fit') { const f = Math.min(size.w / doc.size.w, size.h / doc.size.h); to = { x: (size.w - doc.size.w * f) / 2, y: (size.h - doc.size.h * f) / 2, w: doc.size.w * f, h: doc.size.h * f }; }
  const f = Math.min(to.w / from.w, to.h / from.h);
  const scaleText = (e) => {
    if (!e.text && e.type !== 'table' && e.type !== 'group') return e;
    if (e.type === 'group') return { ...e, children: e.children.map(scaleText) };
    if (e.type === 'table') return { ...e, rows: e.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, text: scaleBody(c.text, f) })) })) };
    return { ...e, text: scaleBody(e.text, f) };
  };
  const slides = doc.slides.map((s) => ({ ...s, elements: s.elements.map((e) => scaleText(fitEl(e, from, to))) }));
  const imported = !THEMES.some((t) => t.id === doc.theme.id);
  const layouts = imported ? doc.layouts.map((l) => scaleLayout(l, from, to, f)) : buildLayouts(themeById(doc.theme.id), size);
  return touch({ ...doc, size: { w: size.w, h: size.h }, slides, layouts });
}
function scaleBody(b, f) {
  if (!b || Math.abs(f - 1) < 0.01) return b;
  const sc = (v) => (v ? Math.max(4, Math.round(v * f * 2) / 2) : v);
  return { ...b, paras: b.paras.map((p) => ({ ...p, runs: p.runs.map((r) => (r.size ? { ...r, size: sc(r.size) } : r)) })), lvl: b.lvl ? b.lvl.map((l) => ({ ...l, size: sc(l.size) })) : b.lvl };
}
function scaleLayout(l, from, to, f) {
  const ph = {};
  for (const [k, d] of Object.entries(l.ph || {})) { const n = fitEl({ ...d, type: 'shape' }, from, to); ph[k] = { ...d, x: n.x, y: n.y, w: n.w, h: n.h, lvl: d.lvl ? d.lvl.map((v) => ({ ...v, size: v.size ? Math.round(v.size * f) : v.size })) : d.lvl }; }
  return { ...l, ph, decor: (l.decor || []).map((e) => fitEl(e, from, to)) };
}

/** switch to a built-in theme: new layouts, placeholders keep their text but take the new geometry & styles */
export function applyTheme(doc, themeId, { keepColors = false } = {}) {
  const t = themeById(themeId);
  const layouts = buildLayouts(t, doc.size);
  const oldLayouts = doc.layouts || [];
  const slides = doc.slides.map((s) => {
    const old = oldLayouts.find((l) => l.id === s.layoutId);
    const kind = old ? old.kind : 'titleContent';
    const nl = layouts.find((l) => l.kind === kind) || layouts.find((l) => l.kind === 'titleContent');
    const elements = s.elements.map((e) => {
      if (!e.ph || !nl.ph[e.ph]) return e;
      const d = nl.ph[e.ph];
      return { ...e, x: d.x, y: d.y, w: d.w, h: d.h, text: e.text ? { ...e.text, anchor: d.anchor || e.text.anchor, lvl: clone(d.lvl), paras: e.text.paras.map(stripParaStyle) } : e.text };
    });
    return { ...s, layoutId: nl.id, elements };
  });
  const theme = keepColors ? { ...themeData(t), colors: doc.theme.colors } : themeData(t);
  return touch({ ...doc, theme, layouts, slides });
}
/** placeholder text loses imported explicit font/size/colour so the theme shows through */
function stripParaStyle(p) {
  const { align, lnSpc, spcBef, spcAft, marL, indent, ...rest } = p; // eslint-disable-line no-unused-vars
  return { ...rest, runs: p.runs.map((r) => { const { font, size, color, ...k } = r; return k; }) }; // eslint-disable-line no-unused-vars
}
export const setThemeColors = (doc, colors) => touch({ ...doc, theme: { ...doc.theme, colors: { ...doc.theme.colors, ...colors } } });
export const setThemeFonts = (doc, fonts) => touch({ ...doc, theme: { ...doc.theme, fonts: { ...doc.theme.fonts, ...fonts } } });

/** reset slide: placeholders back to layout position and styles, missing placeholders re-added */
export function resetSlide(doc, slideId, kind) {
  return updateSlide(doc, slideId, (s) => {
    const layout = kind ? layoutByKind(doc, kind) : layoutOf(doc, s);
    const used = new Set();
    const elements = [];
    for (const e of s.elements) {
      if (e.ph && layout.ph[e.ph]) {
        const d = layout.ph[e.ph]; used.add(e.ph);
        elements.push({ ...e, x: d.x, y: d.y, w: d.w, h: d.h, rot: 0, text: e.text ? { ...e.text, anchor: d.anchor, lvl: clone(d.lvl), paras: e.text.paras.map(stripParaStyle) } : e.text });
      } else if (e.ph && !layout.ph[e.ph] && e.text && !isEmptyBody(e.text)) {
        // text of a placeholder the new layout lacks: map body→body2 … or keep as a free text box
        const free = Object.keys(layout.ph).find((k) => !used.has(k) && /body/.test(k) && /body/.test(e.ph));
        if (free) { const d = layout.ph[free]; used.add(free); elements.push({ ...e, ph: free, x: d.x, y: d.y, w: d.w, h: d.h, text: { ...e.text, lvl: clone(d.lvl) } }); } else elements.push({ ...e, ph: undefined });
      } else if (!e.ph) elements.push(e);
    }
    for (const [k, d] of Object.entries(layout.ph)) if (!used.has(k) && !['dt', 'ftr', 'sldNum'].includes(k)) elements.push(phElement(k, d));
    return { ...s, layoutId: layout.id, elements };
  });
}

/* ---------------------------------------------------------------- sections */
export function addSection(doc, startId, name) {
  const sections = (doc.sections || []).filter((x) => x.start !== startId);
  sections.push({ id: uid('sec'), name, start: startId });
  sections.sort((a, b) => slideIndex(doc, a.start) - slideIndex(doc, b.start));
  return touch({ ...doc, sections });
}
export const renameSection = (doc, id, name) => touch({ ...doc, sections: doc.sections.map((s) => (s.id === id ? { ...s, name } : s)) });
export const removeSection = (doc, id) => touch({ ...doc, sections: doc.sections.filter((s) => s.id !== id) });
/** [{ section|null, slides: [slide] }] in order */
export function sectionGroups(doc) {
  const secs = (doc.sections || []).map((s) => ({ ...s, i: slideIndex(doc, s.start) })).filter((s) => s.i >= 0).sort((a, b) => a.i - b.i);
  const out = [];
  if (!secs.length || secs[0].i > 0) out.push({ section: null, slides: doc.slides.slice(0, secs.length ? secs[0].i : doc.slides.length) });
  secs.forEach((s, k) => out.push({ section: s, slides: doc.slides.slice(s.i, k + 1 < secs.length ? secs[k + 1].i : doc.slides.length) }));
  return out;
}

/* ---------------------------------------------------------------- text of the whole deck */
export function docToText(doc) {
  const out = [];
  doc.slides.forEach((s, i) => {
    const lines = [];
    walk(s.elements, (e) => {
      if (e.text) { const t = bodyText(e.text).trim(); if (t) lines.push(t); }
      if (e.type === 'table') for (const r of e.rows) lines.push(r.cells.filter((c) => !c.merged).map((c) => bodyText(c.text).trim()).join('\t'));
      if (e.type === 'chart' && e.chart.title) lines.push(e.chart.title);
    });
    out.push(`--- ${i + 1} ---`);
    if (lines.length) out.push(lines.join('\n'));
    if (s.notes && s.notes.trim()) out.push(`[Notes] ${s.notes.trim()}`);
  });
  return out.join('\n\n');
}

/* ---------------------------------------------------------------- history (undo/redo) */
export const HISTORY_MAX = 200;
export const history = (doc) => ({ past: [], doc, future: [], key: null, at: 0 });
/**
 * Apply a new doc. `key`: consecutive changes with the same key within 1.5 s merge into one undo step
 * (typing, nudging with arrow keys, dragging a slider).
 */
export function histPush(h, doc, key = null, now = Date.now()) {
  if (doc === h.doc) return h;
  if (key && key === h.key && now - h.at < 1500 && h.past.length) return { ...h, doc, at: now };
  const past = h.past.concat([h.doc]);
  if (past.length > HISTORY_MAX) past.shift();
  return { past, doc, future: [], key, at: now };
}
export function histUndo(h) {
  if (!h.past.length) return h;
  return { past: h.past.slice(0, -1), doc: h.past[h.past.length - 1], future: [h.doc, ...h.future], key: null, at: 0 };
}
export function histRedo(h) {
  if (!h.future.length) return h;
  return { past: [...h.past, h.doc], doc: h.future[0], future: h.future.slice(1), key: null, at: 0 };
}
/** replace the current doc without an undo step (e.g. live text sync while typing) */
export const histReplace = (h, doc) => ({ ...h, doc });
