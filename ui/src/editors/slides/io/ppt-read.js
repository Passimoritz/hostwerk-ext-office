/*
 * Built-in reader for PowerPoint 97–2003 (.ppt / .pps / .pot, binary "PowerPoint Document" in an OLE compound file).
 * No converter needed. readPpt(bytes, { ext }) → { doc, fidelity } with a doc per model.js.
 *
 * What it reads: Current User → UserEditAtom chain → persist directory → DocumentContainer, masters, slides, notes;
 * slide size, fonts, colour scheme, master text styles (TextMasterStyleAtom), slide texts (SlideListWithText +
 * ClientTextbox, TextChars/TextBytes + StyleTextPropAtom: levels, alignment, bullets, spacing, bold/italic/underline,
 * size, colour, font), OfficeArt shapes (groups, anchors, preset types, fills incl. gradients/pictures, lines with
 * arrow heads, rotation, flips, insets, anchoring), pictures from the BStore / "Pictures" stream (PNG, JPEG, DIB→BMP,
 * EMF/WMF/PICT kept as data URLs), placeholders, backgrounds, master decoration, notes, hidden slides, transitions.
 * Every record/slide is wrapped in try/catch – a damaged file still gives the text of all slides it can find.
 * Units: 1 master unit = 1/576 inch = 1587.5 EMU.
 */
import { readCfb, findStream } from '../../_shared/cfb.js';
import { uid, body, para, newDoc, DEFAULT_INSETS } from '../model.js';
import { buildLayouts, themeById, themeData } from '../themes.js';
import { bytesToDataUrl } from './util.js';

const MU = 1587.5; // EMU per master unit
const mu = (v) => Math.round(v * MU);

/* ---------------------------------------------------------------- old build effects → animations */
// AnimEffect (MS-PPT 2.13.2): 0 cut, 1 random, 2 blinds, 3 checker, 4 cover, 5 dissolve, 6 fade, 7 uncover, 8 random bars,
// 9 strips, 10 wipe, 11 box/zoom, 12 fly, 13 split, 14 flash, 17 diamond, 18 plus, 19 wedge, 26 wheel, 27 circle
const BUILD_FX = { 0: 'appear', 1: 'fade', 5: 'fade', 6: 'fade', 14: 'appear', 4: 'fly', 7: 'fly', 12: 'fly', 2: 'wipe', 3: 'wipe', 8: 'wipe', 9: 'wipe', 10: 'wipe', 11: 'zoom', 17: 'zoom', 18: 'zoom', 27: 'zoom', 13: 'split', 19: 'wheel', 26: 'wheel' };
const FLY_DIR = ['l', 't', 'r', 'b', 'l', 'r', 'l', 'r'];
export function buildAnimations(list) {
  return (list || []).filter((a) => a.order >= 0 && a.build > 0).sort((a, b) => a.order - b.order).map((a) => {
    const effect = BUILD_FX[a.effect] || 'appear';
    const anim = { id: uid('a'), el: a.el, cls: 'entr', effect, trigger: (a.flags & 0x4) ? 'after' : 'click', duration: effect === 'appear' ? 1 : 500, delay: (a.flags & 0x4) ? Math.max(0, a.delay || 0) : 0 };
    if (effect === 'fly') anim.dir = FLY_DIR[a.dir] || 'b';
    return anim;
  });
}
/* ---------------------------------------------------------------- tables (groups of cells + lines) */
export function tableFromGroup(children) {
  const cells = children.filter((c) => c.type === 'shape');
  const lines = children.filter((c) => c.type === 'line');
  if (cells.length < 2 || children.some((c) => c.type !== 'shape' && c.type !== 'line')) return null;
  const tol = 6350; // 0.5 pt
  const uniq = (vals) => { const out = []; for (const v of vals.sort((a, b) => a - b)) if (!out.length || v - out[out.length - 1] > tol) out.push(v); return out; };
  const xs = uniq(cells.flatMap((c) => [c.x, c.x + c.w])); const ys = uniq(cells.flatMap((c) => [c.y, c.y + c.h]));
  const at = (arr, v) => arr.findIndex((x) => Math.abs(x - v) <= tol);
  const nr = ys.length - 1; const nc = xs.length - 1;
  if (nr < 1 || nc < 1 || nr * nc > 2500) return null;
  const grid = Array.from({ length: nr }, () => Array.from({ length: nc }, () => null));
  for (const c of cells) {
    const c1 = at(xs, c.x); const c2 = at(xs, c.x + c.w); const r1 = at(ys, c.y); const r2 = at(ys, c.y + c.h);
    if (c1 < 0 || c2 <= c1 || r1 < 0 || r2 <= r1) return null;
    if (grid[r1][c1]) return null;
    const cell = { text: c.text ? { ...c.text } : body([para('')], { anchor: 't' }) };
    if (c.fill !== undefined) cell.fill = c.fill;
    if (r2 - r1 > 1 || c2 - c1 > 1) cell.span = { r: r2 - r1, c: c2 - c1 };
    grid[r1][c1] = cell;
    for (let r = r1; r < r2; r++) for (let k = c1; k < c2; k++) if (r !== r1 || k !== c1) grid[r][k] = { text: body([para('')]), merged: true };
  }
  // border lines → the edges of the cells they lie on
  const edge = (cell, k, ln) => { if (!cell || cell.merged) return; cell.borders = { ...(cell.borders || {}), [k]: ln }; };
  for (const l of lines) {
    const ln = l.line || null;
    if (l.h <= tol) { const ri = at(ys, l.y); if (ri < 0) continue; for (let c = 0; c < nc; c++) { const mid = (xs[c] + xs[c + 1]) / 2; if (mid < l.x - tol || mid > l.x + l.w + tol) continue; if (ri < nr) edge(grid[ri][c], 't', ln); if (ri > 0) edge(grid[ri - 1][c], 'b', ln); } }
    else if (l.w <= tol) { const ci = at(xs, l.x); if (ci < 0) continue; for (let r = 0; r < nr; r++) { const mid = (ys[r] + ys[r + 1]) / 2; if (mid < l.y - tol || mid > l.y + l.h + tol) continue; if (ci < nc) edge(grid[r][ci], 'l', ln); if (ci > 0) edge(grid[r][ci - 1], 'r', ln); } }
  }
  if (grid.some((row) => row.some((c) => !c))) return null;
  for (const row of grid) for (const c of row) if (!c.merged) { for (const k of ['l', 'r', 't', 'b']) if (c.borders && !(k in c.borders)) c.borders[k] = null; }
  return {
    id: uid('t'), type: 'table', name: '', x: xs[0], y: ys[0], w: xs[nc] - xs[0], h: ys[nr] - ys[0], rot: 0,
    cols: xs.slice(1).map((x, i) => x - xs[i]),
    rows: grid.map((cells2, i) => ({ h: ys[i + 1] - ys[i], cells: cells2 })),
    style: { firstRow: false, bandRows: false, firstCol: false, accent: 'accent1' }, border: undefined,
  };
}

/* ---------------------------------------------------------------- record reader */
const CONTAINER = 0xf;
function makeReader(u8) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const n = u8.length;
  const rec = (off, end = n) => {
    if (off < 0 || off + 8 > end) return null;
    const vi = dv.getUint16(off, true);
    const len = dv.getUint32(off + 4, true);
    const bodyOff = off + 8;
    return { ver: vi & 0xf, inst: vi >>> 4, type: dv.getUint16(off + 2, true), len: Math.min(len, Math.max(0, end - bodyOff)), rawLen: len, off, body: bodyOff, end: Math.min(bodyOff + len, end) };
  };
  const kids = (r) => {
    const out = [];
    if (!r || r.ver !== CONTAINER) return out;
    let p = r.body;
    let guard = 0;
    while (p + 8 <= r.end && guard++ < 100000) {
      const c = rec(p, r.end);
      if (!c) break;
      out.push(c);
      p = c.body + c.rawLen;
      if (c.rawLen > r.end - c.body) break;
    }
    return out;
  };
  const kid = (r, type, inst) => kids(r).find((c) => c.type === type && (inst === undefined || c.inst === inst)) || null;
  const deep = (r, type, out = [], depth = 0) => {
    if (depth > 40) return out;
    for (const c of kids(r)) { if (c.type === type) out.push(c); if (c.ver === CONTAINER) deep(c, type, out, depth + 1); }
    return out;
  };
  const top = () => {
    const out = [];
    let p = 0; let guard = 0;
    while (p + 8 <= n && guard++ < 200000) { const r = rec(p); if (!r) break; out.push(r); p = r.body + r.rawLen; }
    return out;
  };
  const u8at = (o) => (o < n ? u8[o] : 0);
  const u16 = (o) => (o + 2 <= n ? dv.getUint16(o, true) : 0);
  const i16 = (o) => (o + 2 <= n ? dv.getInt16(o, true) : 0);
  const u32 = (o) => (o + 4 <= n ? dv.getUint32(o, true) : 0);
  const i32 = (o) => (o + 4 <= n ? dv.getInt32(o, true) : 0);
  const bytes = (o, len) => u8.subarray(Math.max(0, o), Math.min(n, o + len));
  return { rec, kids, kid, deep, top, u8: u8at, u16, i16, u32, i32, bytes, n };
}

const T = {
  Document: 0x03e8, DocumentAtom: 0x03e9, Slide: 0x03ee, SlideAtom: 0x03ef, Notes: 0x03f0, NotesAtom: 0x03f1, Environment: 0x03f2,
  SlidePersistAtom: 0x03f3, MainMaster: 0x03f8, SSSlideInfoAtom: 0x03f9, VBAInfo: 0x03ff, VBAInfoAtom: 0x0400, PPDrawingGroup: 0x040b, PPDrawing: 0x040c,
  FontCollection: 0x07d6, ColorSchemeAtom: 0x07f0, OEPlaceholderAtom: 0x0bc3, ExObjRefAtom: 0x0bc1, OutlineTextRefAtom: 0x0f9e, TextHeaderAtom: 0x0f9f,
  TextCharsAtom: 0x0fa0, StyleTextPropAtom: 0x0fa1, TextMasterStyleAtom: 0x0fa3, TextBytesAtom: 0x0fa8, FontEntityAtom: 0x0fb7,
  SlideListWithText: 0x0ff0, UserEditAtom: 0x0ff5, CurrentUserAtom: 0x0ff6, PersistDirectoryAtom: 0x1772, ExObjList: 0x0409, ExOleObjStg: 0x1011,
  DggContainer: 0xf000, BStoreContainer: 0xf001, DgContainer: 0xf002, SpgrContainer: 0xf003, SpContainer: 0xf004, FBSE: 0xf007, FSPGR: 0xf009, FSP: 0xf00a,
  FOPT: 0xf00b, ClientTextbox: 0xf00d, ChildAnchor: 0xf00f, ClientAnchor: 0xf010, ClientData: 0xf011, TertiaryFOPT: 0xf122,
};

/* ---------------------------------------------------------------- shapes */
const SPT = {
  0: 'rect', 1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 7: 'parallelogram', 8: 'trapezoid', 9: 'hexagon',
  10: 'octagon', 11: 'plus', 12: 'star5', 13: 'rightArrow', 14: 'rightArrow', 15: 'homePlate', 16: 'cube', 17: 'wedgeRoundRectCallout', 18: 'star12',
  19: 'ellipse', 21: 'roundRect', 22: 'can', 23: 'donut', 41: 'wedgeRectCallout', 42: 'wedgeRectCallout', 43: 'wedgeRectCallout', 44: 'wedgeRectCallout',
  45: 'wedgeRectCallout', 46: 'wedgeRectCallout', 47: 'wedgeRectCallout', 48: 'wedgeRectCallout', 49: 'wedgeRectCallout', 53: 'rect', 54: 'rect',
  55: 'chevron', 56: 'pentagon', 57: 'donut', 58: 'star8', 59: 'star12', 60: 'star12', 61: 'wedgeRectCallout', 62: 'wedgeRoundRectCallout',
  63: 'wedgeEllipseCallout', 64: 'flowChartDocument', 65: 'snip1Rect', 66: 'leftArrow', 67: 'downArrow', 68: 'upArrow', 69: 'leftRightArrow',
  70: 'upDownArrow', 71: 'star8', 72: 'star8', 73: 'lightningBolt', 74: 'heart', 75: 'rect', 76: 'plus', 77: 'rightArrow', 78: 'leftArrow',
  79: 'upArrow', 80: 'downArrow', 81: 'leftRightArrow', 82: 'upDownArrow', 83: 'plus', 84: 'rect', 85: 'rect', 86: 'rect', 87: 'rect', 88: 'rect',
  89: 'upArrow', 90: 'rightArrow', 91: 'rightArrow', 92: 'star12', 93: 'rightArrow', 94: 'notchedRightArrow', 95: 'donut', 96: 'smileyFace',
  97: 'rect', 98: 'rect', 99: 'donut', 100: 'donut', 101: 'upArrow', 102: 'rightArrow', 103: 'leftArrow', 104: 'upArrow', 105: 'downArrow',
  106: 'cloudCallout', 107: 'rect', 108: 'rect', 109: 'flowChartProcess', 110: 'flowChartDecision', 111: 'flowChartInputOutput',
  112: 'flowChartPredefinedProcess', 113: 'flowChartProcess', 114: 'flowChartDocument', 115: 'flowChartDocument', 116: 'flowChartTerminator',
  117: 'hexagon', 118: 'trapezoid', 119: 'trapezoid', 120: 'flowChartConnector', 121: 'snip1Rect', 122: 'flowChartDocument', 123: 'ellipse',
  124: 'ellipse', 125: 'diamond', 126: 'triangle', 127: 'triangle', 128: 'triangle', 129: 'roundRect', 130: 'can', 131: 'can', 132: 'can',
  133: 'flowChartProcess', 134: 'roundRect', 135: 'triangle', 176: 'roundRect', 177: 'homePlate', 183: 'sun', 184: 'moon', 185: 'rect', 186: 'rect',
  187: 'star4', 188: 'flowChartDocument', 202: 'rect',
};
const LINE_SPT = new Set([20, 32, 33, 34, 35, 36, 37, 38, 39, 40]);
const ARROW = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'];
const DASH = ['solid', 'dash', 'dot', 'dashDot', 'dashDot', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDash', 'lgDash'];
const SCHEME_REF = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4'];
const LAYOUT_KIND = { 0: 'title', 1: 'titleContent', 2: 'titleContent', 7: 'titleOnly', 8: 'twoContent', 9: 'titleContent', 10: 'twoContent', 11: 'twoContent', 13: 'twoContent', 14: 'comparison', 15: 'titleContent', 16: 'blank', 17: 'titleContent', 18: 'twoContent' };
/** OOXML layout type (p:sldLayout@type, found in PowerPoint 2007+ RoundTripContentMasterInfo12 records) → our kind */
const OOXML_LAYOUT_KIND = { title: 'title', obj: 'titleContent', tx: 'titleContent', txAndObj: 'twoContent', objAndTx: 'twoContent', secHead: 'section', twoObj: 'twoContent', twoColTx: 'twoContent', twoTxTwoObj: 'comparison', titleOnly: 'titleOnly', blank: 'blank', picTx: 'pictureCaption', objTx: 'titleContent', vertTx: 'titleContent', vertTitleAndTx: 'titleContent', objOnly: 'titleContent' };
const RT_CONTENT_MASTER_INFO12 = 0x041e;
/** slide layout from the placeholders a slide actually has (PowerPoint 2007+ writes geom = SL_Blank for every slide) */
function kindFromPlaceholders(types) {
  const has = (...t) => types.some((x) => t.includes(x));
  const bodies = types.filter((x) => [0x0e, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19].includes(x)).length;
  if (has(0x0f, 0x03) || (has(0x10, 0x04) && !bodies)) return 'title';
  if (has(0x1a)) return 'pictureCaption';
  const title = has(0x0d, 0x01, 0x11);
  if (title && bodies >= 4) return 'comparison';
  if (title && bodies >= 2) return 'twoContent';
  if (title && bodies === 1) return 'titleContent';
  if (title) return 'titleOnly';
  if (bodies) return 'titleContent';
  return 'blank';
}
/** one entry of a (small, embedded) zip package, via its central directory; null if absent/unreadable */
async function zipEntry(z, re) {
  try {
    const dv = new DataView(z.buffer, z.byteOffset, z.byteLength);
    let eocd = -1;
    for (let p = z.length - 22; p >= Math.max(0, z.length - 65557); p--) if (dv.getUint32(p, true) === 0x06054b50) { eocd = p; break; }
    if (eocd < 0) return null;
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    for (let k = 0; k < count && p + 46 <= z.length; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) return null;
      const method = dv.getUint16(p + 10, true); const csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true); const xl = dv.getUint16(p + 30, true); const cl = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(z.subarray(p + 46, p + 46 + nl));
      if (re.test(name)) {
        const start = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
        const data = z.subarray(start, start + csize);
        if (method === 0) return data;
        if (method === 8) return inflateRaw(data);
        return null;
      }
      p += 46 + nl + xl + cl;
    }
  } catch { /* not a zip */ }
  return null;
}
const PH_KEY = { 1: 'title', 2: 'body', 3: 'ctrTitle', 4: 'subTitle', 7: 'dt', 8: 'sldNum', 9: 'ftr', 13: 'title', 14: 'body', 15: 'ctrTitle', 16: 'subTitle', 17: 'title', 18: 'body', 19: 'body', 20: 'body', 21: 'body', 22: 'body', 23: 'body', 24: 'body' };
const hex = (r, g, b) => `#${[r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('')}`.toUpperCase();

/* ---------------------------------------------------------------- text */
function readTextRecords(R, recs) {
  // groups of TextHeaderAtom + text + style in order of appearance
  const out = [];
  let cur = null;
  for (const r of recs) {
    if (r.type === T.TextHeaderAtom) { cur = { type: R.u32(r.body), text: '', style: null }; out.push(cur); continue; }
    if (r.type === T.TextCharsAtom || r.type === T.TextBytesAtom) {
      if (!cur) { cur = { type: 4, text: '', style: null }; out.push(cur); }
      let s = '';
      if (r.type === T.TextCharsAtom) for (let p = r.body; p + 1 < r.end; p += 2) s += String.fromCharCode(R.u16(p));
      else for (let p = r.body; p < r.end; p++) s += String.fromCharCode(R.u8(p));
      cur.text = s;
    } else if (r.type === T.StyleTextPropAtom && cur) cur.style = r;
  }
  return out;
}

function readColorIndex(R, o) {
  const r = R.u8(o); const g = R.u8(o + 1); const b = R.u8(o + 2); const idx = R.u8(o + 3);
  if (idx === 0xfe) return hex(r, g, b);
  if (idx < 8) return { scheme: idx };
  return null;
}

/** TextPFException at o → [props, newOffset] */
function readPF(R, o, end) {
  const m = R.u32(o); o += 4;
  const pf = {};
  if (m & 0x0f) { const f = R.u16(o); o += 2; if (m & 0x1) pf.hasBullet = !!(f & 1); }
  if (m & 0x80) { pf.bulletChar = R.u16(o); o += 2; }
  if (m & 0x10) { pf.bulletFont = R.u16(o); o += 2; }
  if (m & 0x40) { pf.bulletSize = R.i16(o); o += 2; }
  if (m & 0x20) { pf.bulletColor = readColorIndex(R, o); o += 4; }
  if (m & 0x800) { pf.align = R.u16(o); o += 2; }
  if (m & 0x1000) { pf.lineSpacing = R.i16(o); o += 2; }
  if (m & 0x2000) { pf.spaceBefore = R.i16(o); o += 2; }
  if (m & 0x4000) { pf.spaceAfter = R.i16(o); o += 2; }
  if (m & 0x100) { pf.leftMargin = R.u16(o); o += 2; }
  if (m & 0x400) { pf.indent = R.u16(o); o += 2; }
  if (m & 0x8000) o += 2;
  if (m & 0x100000) { const c = R.u16(o); o += 2 + c * 4; }
  if (m & 0x10000) o += 2;
  if (m & 0xe0000) o += 2;
  if (m & 0x200000) o += 2;
  if (o > end) throw new Error('PF overflow');
  return [pf, o];
}
function readCF(R, o, end) {
  const m = R.u32(o); o += 4;
  const cf = {};
  if (m & 0xffff) { const s = R.u16(o); o += 2; if (m & 1) cf.b = !!(s & 1); if (m & 2) cf.i = !!(s & 2); if (m & 4) cf.u = !!(s & 4); }
  if (m & 0x10000) { cf.font = R.u16(o); o += 2; }
  if (m & 0x200000) o += 2;
  if (m & 0x400000) o += 2;
  if (m & 0x800000) o += 2;
  if (m & 0x20000) { cf.size = R.u16(o); o += 2; }
  if (m & 0x40000) { cf.color = readColorIndex(R, o); o += 4; }
  if (m & 0x80000) { cf.baseline = R.i16(o); o += 2; }
  if (m & 0x100000) o += 4;
  if (m & 0x1000000) o += 2;
  if (m & 0x2000000) o += 2;
  if (m & 0x4000000) o += 4;
  if (o > end) throw new Error('CF overflow');
  return [cf, o];
}

/** StyleTextPropAtom → { pf: [{ count, level, pf }], cf: [{ count, cf }] } for a text of length len */
function readStyle(R, r, len) {
  const res = { pf: [], cf: [] };
  if (!r) return res;
  let o = r.body;
  try {
    let total = 0;
    while (total < len + 1 && o + 6 <= r.end) {
      const count = R.u32(o); const level = R.u16(o + 4);
      const [pf, n] = readPF(R, o + 6, r.end);
      res.pf.push({ count, level, pf }); total += count; o = n;
      if (!count) break;
    }
    total = 0;
    while (total < len + 1 && o + 4 <= r.end) {
      const count = R.u32(o);
      const [cf, n] = readCF(R, o + 4, r.end);
      res.cf.push({ count, cf }); total += count; o = n;
      if (!count) break;
    }
  } catch { /* partial style is fine */ }
  return res;
}

/* ---------------------------------------------------------------- main */
export async function readPpt(bytes, { ext } = {}) { // eslint-disable-line no-unused-vars
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const fidelity = new Set();
  const cfb = readCfb(u8);
  const main = findStream(cfb, 'PowerPoint Document');
  if (!main) throw new Error('This is not a PowerPoint 97–2003 file');
  const R = makeReader(main);
  if (cfb.entries.some((e) => /(^|\/)(_?VBA(_PROJECT)?|Macros)(\/|$)/i.test(e.path))) fidelity.add('macros');

  /* persist directory */
  const persist = new Map();
  let docRef = null;
  try {
    const cu = findStream(cfb, 'Current User');
    if (cu && cu.length >= 20) {
      const CR = makeReader(cu);
      let edit = CR.u32(16);
      const seen = new Set();
      while (edit > 0 && edit < R.n && !seen.has(edit) && seen.size < 1000) {
        seen.add(edit);
        const ue = R.rec(edit);
        if (!ue || ue.type !== T.UserEditAtom) break;
        if (docRef == null) docRef = R.u32(ue.body + 16);
        const pd = R.rec(R.u32(ue.body + 12));
        if (pd && pd.type === T.PersistDirectoryAtom) {
          let p = pd.body;
          while (p + 4 <= pd.end) {
            const v = R.u32(p); p += 4;
            const id = v & 0xfffff; const c = v >>> 20;
            for (let k = 0; k < c && p + 4 <= pd.end; k++, p += 4) if (!persist.has(id + k)) persist.set(id + k, R.u32(p));
          }
        }
        edit = R.u32(ue.body + 8);
      }
    }
  } catch { /* fall back to scanning */ }
  const byRef = (ref, type) => { const off = persist.get(ref); if (off == null) return null; const r = R.rec(off); return r && r.type === type ? r : null; };
  let topRecs = null;
  const scan = (type) => { topRecs = topRecs || R.top(); return topRecs.filter((r) => r.type === type); };

  let docRec = docRef != null ? byRef(docRef, T.Document) : null;
  if (!docRec) { const all = scan(T.Document); docRec = all[all.length - 1] || null; }

  /* document atom, fonts, pictures */
  let W = 5760; let H = 4320;
  const fonts = [];
  const blips = [];
  const slwt = {};
  if (docRec) {
    try {
      const da = R.kid(docRec, T.DocumentAtom);
      if (da) { const w = R.i32(da.body); const h = R.i32(da.body + 4); if (w > 100 && h > 100) { W = w; H = h; } }
    } catch { /* default size */ }
    try {
      const env = R.kid(docRec, T.Environment);
      const fc = env && R.kid(env, T.FontCollection);
      for (const fe of R.kids(fc).filter((r) => r.type === T.FontEntityAtom)) {
        let s = '';
        for (let k = 0; k < 32; k++) { const c = R.u16(fe.body + k * 2); if (!c) break; s += String.fromCharCode(c); }
        fonts[fe.inst] = s;
      }
    } catch { /* no fonts */ }
    try {
      const dgg = R.kid(R.kid(docRec, T.PPDrawingGroup), T.DggContainer);
      const bs = dgg && R.kid(dgg, T.BStoreContainer);
      for (const fbse of R.kids(bs).filter((r) => r.type === T.FBSE)) blips.push(fbse);
    } catch { /* no pictures */ }
    for (const l of R.kids(docRec).filter((r) => r.type === T.SlideListWithText)) {
      try { slwt[l.inst] = parseSlwt(R, l); } catch { /* ignore list */ }
    }
    if (R.deep(docRec, T.VBAInfo).length) fidelity.add('macros');
    if (R.deep(docRec, T.ExOleObjStg).length) fidelity.add('ole');
  }
  const size = { w: mu(W), h: mu(H) };

  /* pictures */
  const pictures = findStream(cfb, 'Pictures');
  const PR = pictures ? makeReader(pictures) : null;
  const picCache = new Map();
  const picture = async (pib) => {
    if (!pib || pib > blips.length) return null;
    if (picCache.has(pib)) return picCache.get(pib);
    let url = null;
    try {
      const fbse = blips[pib - 1];
      const cbName = R.u8(fbse.body + 33);
      let blip = null; let BR = R;
      if (fbse.len > 36 + cbName + 8) blip = R.rec(fbse.body + 36 + cbName, fbse.end);
      else if (PR) { blip = PR.rec(R.u32(fbse.body + 28)); BR = PR; }
      if (blip) url = await blipToUrl(BR, blip, fidelity);
    } catch { url = null; }
    picCache.set(pib, url);
    return url;
  };

  /* colour scheme (master) */
  const masters = [];
  if (slwt[1]) for (const e of slwt[1]) { const m = byRef(e.ref, T.MainMaster); if (m) masters.push({ rec: m, id: e.slideId }); }
  if (!masters.length) for (const m of scan(T.MainMaster)) masters.push({ rec: m, id: null });
  const master = masters[0] ? masters[0].rec : null;
  const schemeOf = (container) => {
    const cs = container && R.kids(container).find((r) => r.type === T.ColorSchemeAtom && r.inst === 1);
    if (!cs) return null;
    const out = [];
    for (let k = 0; k < 8; k++) out.push(hex(R.u8(cs.body + k * 4), R.u8(cs.body + k * 4 + 1), R.u8(cs.body + k * 4 + 2)));
    return out;
  };
  const scheme = schemeOf(master) || ['#FFFFFF', '#000000', '#808080', '#000000', '#BBE0E3', '#333399', '#009999', '#99CC00'];

  const office = themeById('office');
  const base = themeData(office);
  const font0 = fonts[0] || 'Arial';
  const theme = {
    ...base, id: 'imported', name: 'PowerPoint 97–2003', dark: false,
    colors: { ...base.colors, lt1: scheme[0], dk1: scheme[1], lt2: scheme[2], dk2: scheme[3], accent1: scheme[4], accent2: scheme[5], accent3: scheme[6], accent4: scheme[7], hlink: scheme[6], folHlink: scheme[7] },
    fonts: { major: font0, minor: font0 },
    bg: { type: 'solid', color: 'bg1' },
  };

  // colour → model ref (scheme index 0..7 → theme ref, or hex from a slide's own scheme)
  const mkColor = (localScheme) => (c) => {
    if (!c) return undefined;
    if (typeof c === 'string') return c;
    if (c.scheme != null) return localScheme ? localScheme[c.scheme] : SCHEME_REF[c.scheme];
    return undefined;
  };

  /* text → body */
  const toParas = (entry, color, lvlBase) => {
    const text = (entry.text || '').replace(/\r$/, '');
    const st = readStyle(R, entry.style, text.length);
    const chars = [...text];
    // per-char style arrays
    const pfAt = []; const cfAt = [];
    let idx = 0;
    for (const r of st.pf) { for (let k = 0; k < r.count && idx <= chars.length; k++) pfAt[idx++] = r; }
    idx = 0;
    for (const r of st.cf) { for (let k = 0; k < r.count && idx <= chars.length; k++) cfAt[idx++] = r.cf; }
    const paras = [];
    let cur = null; let runProps = null; let runText = '';
    const flushRun = () => { if (cur && runText) cur.runs.push({ text: runText, ...runProps }); runText = ''; };
    const startPara = (i) => {
      const r = pfAt[i] || pfAt[pfAt.length - 1];
      const pf = (r && r.pf) || {};
      const p = { runs: [], level: Math.min(8, (r && r.level) || 0) };
      if (pf.align != null) p.align = ['l', 'ctr', 'r', 'just', 'dist'][pf.align] || 'l';
      if (pf.hasBullet === false) p.bullet = false;
      else if (pf.hasBullet) { p.bullet = { char: pf.bulletChar ? String.fromCharCode(pf.bulletChar) : '•' }; const bc = color(pf.bulletColor); if (bc) p.bullet.color = bc; if (pf.bulletFont != null && fonts[pf.bulletFont]) p.bullet.font = fonts[pf.bulletFont]; }
      if (pf.lineSpacing > 0) p.lnSpc = pf.lineSpacing;
      if (pf.spaceBefore != null) p.spcBef = pf.spaceBefore < 0 ? -pf.spaceBefore / 8 : Math.round(pf.spaceBefore * 0.2);
      if (pf.spaceAfter != null) p.spcAft = pf.spaceAfter < 0 ? -pf.spaceAfter / 8 : Math.round(pf.spaceAfter * 0.2);
      if (pf.leftMargin != null) { p.marL = mu(pf.leftMargin); if (pf.indent != null) p.indent = mu(pf.indent - pf.leftMargin); }
      cur = p; paras.push(p); runProps = null; runText = '';
    };
    const propsOf = (cf) => {
      const o = {};
      if (!cf) return o;
      if (cf.b) o.b = true; else if (cf.b === false) o.b = false;
      if (cf.i) o.i = true; else if (cf.i === false) o.i = false;
      if (cf.u) o.u = true;
      if (cf.size) o.size = cf.size;
      if (cf.font != null && fonts[cf.font]) o.font = fonts[cf.font];
      const c = color(cf.color); if (c) o.color = c;
      if (cf.baseline) o.baseline = cf.baseline;
      return o;
    };
    startPara(0);
    for (let i = 0; i < chars.length; i++) {
      const ch = chars[i];
      if (ch === '\r') { flushRun(); startPara(i + 1); continue; }
      const p = propsOf(cfAt[i] || cfAt[cfAt.length - 1]);
      const key = JSON.stringify(p);
      if (runProps === null || key !== JSON.stringify(runProps)) { flushRun(); runProps = p; }
      runText += ch === '\v' ? '\n' : ch;
    }
    flushRun();
    void lvlBase;
    return paras.length ? paras : [para('')];
  };

  /* master text styles → lvl arrays of the layouts */
  const masterColor = mkColor(null);
  const masterStyles = {};
  try {
    for (const r of R.kids(master).filter((x) => x.type === T.TextMasterStyleAtom)) {
      const levels = [];
      const cl = R.u16(r.body);
      let o = r.body + 2;
      for (let k = 0; k < Math.min(cl, 9) && o < r.end; k++) {
        let lvl = k;
        if (r.inst >= 5) { lvl = R.u16(o); o += 2; }
        const [pf, o2] = readPF(R, o, r.end);
        const [cf, o3] = readCF(R, o2, r.end);
        o = o3;
        levels[Math.min(8, lvl)] = { pf, cf };
      }
      masterStyles[r.inst] = levels;
    }
  } catch { /* defaults */ }
  const lvlFrom = (lv, prev) => {
    const out = { ...(prev || {}) };
    if (!lv) return out;
    const { pf, cf } = lv;
    if (cf.size) out.size = cf.size;
    if (cf.font != null && fonts[cf.font]) out.font = fonts[cf.font];
    if (cf.b != null) out.b = cf.b;
    if (cf.i != null) out.i = cf.i;
    const c = masterColor(cf.color); if (c) out.color = c;
    if (pf.align != null) out.align = ['l', 'ctr', 'r', 'just', 'dist'][pf.align] || 'l';
    if (pf.hasBullet === false) out.bullet = false;
    else if (pf.hasBullet) out.bullet = { char: pf.bulletChar ? String.fromCharCode(pf.bulletChar) : '•' };
    if (pf.leftMargin != null) out.marL = mu(pf.leftMargin);
    if (pf.indent != null && pf.leftMargin != null) out.indent = mu(pf.indent - pf.leftMargin);
    if (pf.lineSpacing > 0) out.lnSpc = pf.lineSpacing;
    if (pf.spaceBefore != null) out.spcBef = pf.spaceBefore < 0 ? -pf.spaceBefore / 8 : Math.round(pf.spaceBefore * 0.2);
    return out;
  };
  const applyMaster = (lvls, inst, parentInst) => {
    const src = masterStyles[inst]; const par = parentInst != null ? masterStyles[parentInst] : null;
    if (!src && !par) return lvls;
    return lvls.map((l, i) => { let v = { ...l }; if (par && par[i]) v = lvlFrom(par[i], v); if (src && src[i]) v = lvlFrom(src[i], v); return v; });
  };

  const layouts = buildLayouts(office, size).map((l) => {
    const ph = {};
    for (const [k, d] of Object.entries(l.ph)) {
      let lvl = d.lvl;
      if (k === 'title' || k === 'ctrTitle') lvl = lvl.map((x) => ({ ...x, color: 'tx2' }));
      if (k === 'subTitle') lvl = lvl.map((x) => ({ ...x, color: 'tx1' }));
      if (k === 'title') lvl = applyMaster(lvl, 0);
      else if (k === 'ctrTitle') lvl = applyMaster(lvl, 6, 0);
      else if (k === 'subTitle') lvl = applyMaster(lvl, 5, 1);
      else if (/^body/.test(k)) lvl = applyMaster(lvl, 1);
      ph[k] = { ...d, lvl };
    }
    return { ...l, ph };
  });
  const layoutFor = (kind) => layouts.find((l) => l.kind === kind) || layouts[1];

  /* ---------------- OfficeArt drawing */
  const readOpt = (sp) => {
    const props = new Map();
    for (const r of R.kids(sp).filter((x) => x.type === T.FOPT || x.type === T.TertiaryFOPT)) {
      const n = r.inst;
      let cx = r.body + n * 6;
      for (let k = 0; k < n; k++) {
        const o = r.body + k * 6;
        if (o + 6 > r.end) break;
        const id = R.u16(o); const val = R.i32(o + 2);
        const pid = id & 0x3fff;
        const p = { val, uval: val >>> 0, bid: !!(id & 0x4000) };
        if (id & 0x8000) { p.complex = R.bytes(cx, val >>> 0); cx += val >>> 0; }
        props.set(pid, p);
      }
    }
    return props;
  };
  const colorRef = (c, local) => {
    if (c == null) return undefined;
    const v = c >>> 0;
    const r = v & 0xff; const g = (v >>> 8) & 0xff; const b = (v >>> 16) & 0xff; const fl = (v >>> 24) & 0xff;
    if (fl & 0x08) return local ? local[r] || SCHEME_REF[r] : SCHEME_REF[r] || 'accent1';
    if (fl & 0x10) return undefined; // system colour index – fall back
    return hex(r, g, b);
  };
  const boolProp = (props, id, bit, useBit, def) => {
    const p = props.get(id);
    if (!p) return def;
    const v = p.uval;
    if (v & useBit) return !!(v & bit);
    return def;
  };

  const readSp = (sp) => {
    const out = { kids: R.kids(sp) };
    const fsp = out.kids.find((r) => r.type === T.FSP);
    if (fsp) { out.spt = fsp.inst; out.spid = R.u32(fsp.body); out.flags = R.u32(fsp.body + 4); } else { out.spt = 1; out.flags = 0; }
    const ca = out.kids.find((r) => r.type === T.ClientAnchor);
    if (ca) out.client = ca.len >= 16 ? { t: R.i32(ca.body), l: R.i32(ca.body + 4), r: R.i32(ca.body + 8), b: R.i32(ca.body + 12) } : { t: R.i16(ca.body), l: R.i16(ca.body + 2), r: R.i16(ca.body + 4), b: R.i16(ca.body + 6) };
    const ch = out.kids.find((r) => r.type === T.ChildAnchor);
    if (ch) out.child = { l: R.i32(ch.body), t: R.i32(ch.body + 4), r: R.i32(ch.body + 8), b: R.i32(ch.body + 12) };
    const gr = out.kids.find((r) => r.type === T.FSPGR);
    if (gr) out.grp = { l: R.i32(gr.body), t: R.i32(gr.body + 4), r: R.i32(gr.body + 8), b: R.i32(gr.body + 12) };
    const cd = out.kids.find((r) => r.type === T.ClientData);
    if (cd) {
      const ph = R.deep(cd, T.OEPlaceholderAtom)[0];
      if (ph) out.ph = R.u8(ph.body + 4);
      if (R.deep(cd, T.ExObjRefAtom).length) out.ole = true;
      // old build effects (AnimationInfoAtom, PowerPoint 97–2003)
      const ai = R.deep(cd, 0x0ff1)[0];
      if (ai && ai.len >= 26) out.anim = { flags: R.u16(ai.body + 4), delay: R.i32(ai.body + 12), order: R.i16(ai.body + 16), build: R.u8(ai.body + 20), effect: R.u8(ai.body + 21), dir: R.u8(ai.body + 22) };
    }
    const tb = out.kids.find((r) => r.type === T.ClientTextbox);
    if (tb) {
      const recs = R.kids(tb);
      const ref = recs.find((r) => r.type === T.OutlineTextRefAtom);
      if (ref) out.textRef = R.i32(ref.body);
      const texts = readTextRecords(R, recs);
      if (texts.length) out.text = texts[0];
    }
    out.props = readOpt(sp);
    return out;
  };

  const MU_MAP = (a) => ({ x: mu(a.l), y: mu(a.t), w: mu(a.r - a.l), h: mu(a.b - a.t) });

  // ctx: { map, color, texts (slide texts for OutlineTextRefAtom), used:Set, kindFor }
  const shapeToEl = async (s, box, ctx) => {
    const P = s.props;
    const get = (id) => (P.has(id) ? P.get(id).val : undefined);
    const flipH = !!(s.flags & 0x40); const flipV = !!(s.flags & 0x80);
    let rot = get(0x0004) != null ? get(0x0004) / 65536 : 0;
    rot = ((rot % 360) + 360) % 360;
    let { x, y, w, h } = box;
    if ((rot >= 45 && rot < 135) || (rot >= 225 && rot < 315)) { const cx = x + w / 2; const cy = y + h / 2; [w, h] = [h, w]; x = Math.round(cx - w / 2); y = Math.round(cy - h / 2); }
    const common = { id: uid(s.spt === 75 ? 'i' : 's'), name: '', x, y, w: Math.max(1, w), h: Math.max(1, h), rot: Math.round(rot * 100) / 100 };
    if (flipH) common.flipH = true;
    if (flipV) common.flipV = true;
    const local = ctx.scheme;

    // line / connector
    if (LINE_SPT.has(s.spt) || (s.flags & 0x100)) {
      const lw = get(0x01cb);
      return {
        ...common, type: 'line', geom: s.spt >= 33 && s.spt <= 36 ? 'elbow' : s.spt >= 37 && s.spt <= 40 ? 'curve' : 'line',
        line: { color: colorRef(get(0x01c0), local) || 'tx1', width: lw != null ? lw : 9525, dash: DASH[get(0x01ce) || 0] || 'solid', head: ARROW[get(0x01d0) || 0] || 'none', tail: ARROW[get(0x01d1) || 0] || 'none' },
      };
    }

    // text
    let textEntry = s.text || null;
    if (s.textRef != null && ctx.texts && ctx.texts[s.textRef]) { textEntry = ctx.texts[s.textRef]; ctx.used.add(s.textRef); }
    let phKey = s.ph != null ? PH_KEY[s.ph] : undefined;
    if (phKey && /^body/.test(phKey) && ctx.bodies) { const n = ++ctx.bodies.n; if (n > 1) phKey = `body${n}`; }
    if (phKey === 'title' && ctx.kind === 'title') phKey = 'ctrTitle';
    if (phKey === 'body' && ctx.kind === 'title') phKey = 'subTitle';
    const layout = ctx.layout;
    const phDef = phKey && layout && layout.ph ? layout.ph[phKey] || layout.ph[phKey.replace(/\d+$/, '')] : null;

    const pib = get(0x0104);
    if (s.spt === 75 || (pib && !P.has(0x0186) && !textEntry)) {
      const src = await picture(pib);
      if (s.ole) fidelity.add('ole');
      if (!src) { fidelity.add('effects'); return null; }
      const fr = (id) => { const v = get(id); return v ? Math.max(0, Math.min(1, v / 65536)) : 0; };
      return { ...common, type: 'image', src, crop: { t: fr(0x0100), b: fr(0x0101), l: fr(0x0102), r: fr(0x0103) }, opacity: 1, line: null, shadow: null, alt: '' };
    }

    // fill
    const isText = s.spt === 202 || !!phKey;
    const filled = boolProp(P, 0x01bf, 0x10, 0x100000, !isText);
    let fill = null;
    if (filled) {
      const ft = get(0x0180) || 0;
      const fc = colorRef(get(0x0181), local) || 'accent1';
      const op = get(0x0182);
      const alpha = op != null && op < 65536 ? Math.round((op / 65536) * 100) : 100;
      const withA = (c) => (alpha < 100 ? `${c}@${alpha}` : c);
      if (ft >= 4 && ft <= 8) {
        const back = colorRef(get(0x0183), local) || 'bg1';
        const ang = get(0x018b) != null ? (get(0x018b) / 65536) : 0;
        fill = { type: 'grad', angle: Math.round(((90 - ang) % 360 + 360) % 360), kind: ft === 5 || ft === 6 ? 'radial' : 'linear', stops: [{ pos: 0, color: withA(fc) }, { pos: 100, color: back }] };
      } else if (ft === 2 || ft === 3) {
        const src = await picture(get(0x0186));
        fill = src ? { type: 'image', src, mode: ft === 2 ? 'tile' : 'stretch' } : { type: 'solid', color: withA(fc) };
      } else if (ft === 9) fill = null;
      else fill = { type: 'solid', color: withA(fc) };
    }
    const lined = boolProp(P, 0x01ff, 0x08, 0x80000, !isText);
    const line = lined ? { color: colorRef(get(0x01c0), local) || 'tx1', width: get(0x01cb) != null ? get(0x01cb) : 9525, dash: DASH[get(0x01ce) || 0] || 'solid' } : null;
    const shadowOn = boolProp(P, 0x023f, 0x02, 0x20000, false);
    const shadow = shadowOn ? { color: colorRef(get(0x0201), local) || '#808080', blur: 38100, dist: 38100, dir: 45, alpha: 60 } : null;

    let geom = SPT[s.spt];
    if (!geom) { geom = 'rect'; fidelity.add('geometry'); }
    if (s.spt === 0) fidelity.add('geometry');
    const el = { ...common, type: 'shape', geom, fill, line, shadow };
    if (s.spt === 2 && get(0x0147) != null) el.adj = [Math.round((get(0x0147) / 21600) * 100000)];
    if (s.spt === 202 || (phKey && !filled)) el.textbox = true;
    if (s.ole) fidelity.add('ole');

    const anchorV = get(0x0087);
    const anchor = anchorV == null ? (phDef ? phDef.anchor : (el.textbox ? 't' : 'ctr')) : [ 't', 'ctr', 'b', 't', 'ctr', 'b', 't', 'ctr', 'b', 'ctr'][anchorV] || 't';
    const ins = (id, d) => (get(id) != null ? get(id) : d);
    const tb = {
      anchor, insets: { l: ins(0x0081, DEFAULT_INSETS.l), t: ins(0x0082, DEFAULT_INSETS.t), r: ins(0x0083, DEFAULT_INSETS.r), b: ins(0x0084, DEFAULT_INSETS.b) },
      autofit: 'none', wrap: get(0x0085) !== 2,
    };
    if (phDef && phDef.lvl) tb.lvl = JSON.parse(JSON.stringify(phDef.lvl));
    el.text = body(textEntry ? toParas(textEntry, ctx.color) : [para('', geom === 'rect' && el.textbox ? {} : { align: 'ctr' })], tb);
    if (phKey) {
      if (['dt', 'sldNum', 'ftr'].includes(phKey) && !(textEntry && textEntry.text.trim())) return null;
      el.ph = phKey; el.prompt = phDef ? phDef.prompt || '' : '';
    }
    return el;
  };

  const readDrawing = async (container, ctx) => {
    const out = { elements: [], background: undefined };
    const ppd = R.kid(container, T.PPDrawing);
    const dg = ppd && R.kid(ppd, T.DgContainer);
    if (!dg) return out;
    for (const r of R.kids(dg)) {
      if (r.type === T.SpContainer) {
        try {
          const s = readSp(r);
          if (s.flags & 0x400) out.background = await bgFill(s, ctx);
        } catch { /* ignore background */ }
      }
    }
    const walkGroup = async (spgr, map, depth) => {
      const items = R.kids(spgr).filter((r) => r.type === T.SpContainer || r.type === T.SpgrContainer);
      const res = [];
      for (let k = 0; k < items.length; k++) {
        const r = items[k];
        try {
          if (r.type === T.SpgrContainer) {
            const head = R.kids(r).find((x) => x.type === T.SpContainer);
            if (!head || depth > 20) continue;
            const g = readSp(head);
            const a = g.child || g.client;
            const box = a ? map(a) : null;
            const gr = g.grp;
            let childMap = map;
            if (box && gr && gr.r !== gr.l && gr.b !== gr.t) {
              childMap = (c) => ({ x: Math.round(box.x + ((c.l - gr.l) * box.w) / (gr.r - gr.l)), y: Math.round(box.y + ((c.t - gr.t) * box.h) / (gr.b - gr.t)), w: Math.round(((c.r - c.l) * box.w) / (gr.r - gr.l)), h: Math.round(((c.b - c.t) * box.h) / (gr.b - gr.t)) });
            }
            const children = (await walkGroup(r, childMap, depth + 1)).filter(Boolean);
            if (!children.length) continue;
            // a PowerPoint table is a group (tableProperties 0x39F / tableRowProperties 0x3A0) of cell boxes and border lines
            if ((g.props.has(0x039f) || g.props.has(0x03a0)) && depth < 20) { const tbl = tableFromGroup(children); if (tbl) { res.push(tbl); continue; } }
            if (children.length === 1) { res.push(children[0]); continue; }
            let x1 = Infinity; let y1 = Infinity; let x2 = -Infinity; let y2 = -Infinity;
            for (const c of children) { x1 = Math.min(x1, c.x); y1 = Math.min(y1, c.y); x2 = Math.max(x2, c.x + c.w); y2 = Math.max(y2, c.y + c.h); }
            const rot = g.props.has(4) ? ((g.props.get(4).val / 65536) % 360 + 360) % 360 : 0;
            res.push({ id: uid('g'), type: 'group', name: '', x: x1, y: y1, w: x2 - x1, h: y2 - y1, rot, children });
            continue;
          }
          const s = readSp(r);
          if (k === 0 && (s.flags & 0x1)) continue; // the group's own shape
          if ((s.flags & 0x4) || (s.flags & 0x8) || (s.flags & 0x400)) continue; // patriarch / deleted / background
          const a = (s.flags & 0x2) && s.child ? s.child : s.client || s.child;
          let box = a ? map(a) : null;
          if (!box && s.ph != null) { const key = PH_KEY[s.ph]; const d = key && ctx.layout && ctx.layout.ph[key]; if (d) box = { x: d.x, y: d.y, w: d.w, h: d.h }; }
          if (!box) continue;
          const el = await shapeToEl(s, box, ctx);
          if (el) { res.push(el); if (s.anim && ctx.anims) ctx.anims.push({ el: el.id, ...s.anim }); }
        } catch { /* skip broken shape */ }
      }
      return res;
    };
    const patr = R.kid(dg, T.SpgrContainer);
    if (patr) out.elements = await walkGroup(patr, MU_MAP, 0);
    return out;
  };
  const bgFill = async (s, ctx) => {
    const P = s.props;
    const get = (id) => (P.has(id) ? P.get(id).val : undefined);
    const ft = get(0x0180) || 0;
    const fc = colorRef(get(0x0181), ctx.scheme) || 'bg1';
    if (ft >= 4 && ft <= 8) return { type: 'grad', angle: 90, stops: [{ pos: 0, color: fc }, { pos: 100, color: colorRef(get(0x0183), ctx.scheme) || 'bg1' }] };
    if (ft === 2 || ft === 3) { const src = await picture(get(0x0186)); if (src) return { type: 'image', src, mode: ft === 2 ? 'tile' : 'stretch' }; }
    return { type: 'solid', color: fc };
  };

  /* master decoration + background */
  if (master) {
    try {
      const ctx = { scheme: null, color: masterColor, texts: null, used: new Set(), layout: layouts[1], kind: 'titleContent', bodies: { n: 0 } };
      const d = await readDrawing(master, ctx);
      const decor = d.elements.filter((e) => !e.ph).map((e) => ({ ...e, locked: true, decor: true }));
      for (const l of layouts) l.decor = [...(l.decor || []), ...decor.map((e) => ({ ...e, id: uid('d') }))];
      if (d.background) theme.bg = d.background;
    } catch { /* plain master */ }
  }

  /* slides */
  const slideList = [];
  if (slwt[0] && slwt[0].length) for (const e of slwt[0]) slideList.push({ rec: byRef(e.ref, T.Slide), entry: e });
  if (!slideList.length || slideList.every((s) => !s.rec)) {
    slideList.length = 0;
    for (const r of scan(T.Slide)) slideList.push({ rec: r, entry: null });
  }
  // notes containers by notes id and by slide id
  const notesById = new Map(); const notesBySlide = new Map();
  const addNotes = (r, id) => { if (!r) return; if (id != null) notesById.set(id, r); const na = R.kid(r, T.NotesAtom); if (na) notesBySlide.set(R.u32(na.body), r); };
  if (slwt[2]) for (const e of slwt[2]) addNotes(byRef(e.ref, T.Notes), e.slideId);
  if (!notesById.size) for (const r of scan(T.Notes)) addNotes(r, null);

  // PowerPoint 2007+ stores every slide layout as an extra MainMaster ("content master") carrying the layout's
  // OOXML in a RoundTripContentMasterInfo12 zip; the slide points at it via SlideAtom.masterIdRef
  const masterKind = new Map();
  for (const m of masters) {
    if (m.id == null) continue;
    try {
      const info = R.kid(m.rec, RT_CONTENT_MASTER_INFO12);
      if (!info) continue;
      const xml = await zipEntry(R.bytes(info.body, info.len), /slideLayouts\/[^/]+\.xml$/);
      if (!xml) continue;
      const head = /<(?:\w+:)?sldLayout\b[^>]*>/.exec(new TextDecoder().decode(xml));
      const type = head && /\stype="([^"]+)"/.exec(head[0]);
      masterKind.set(m.id, OOXML_LAYOUT_KIND[type ? type[1] : 'cust'] || null);
    } catch { /* unknown layout */ }
  }
  const slidePhTypes = (sr) => { try { return R.deep(sr, T.OEPlaceholderAtom).map((ph) => R.u8(ph.body + 4)); } catch { return []; } };

  const slides = [];
  for (let si = 0; si < slideList.length; si++) {
    const { rec: sr, entry } = slideList[si];
    const texts = entry ? entry.texts : [];
    const slide = { id: uid('s'), layoutId: layoutFor('titleContent').id, hidden: false, background: null, elements: [], notes: '', transition: { type: 'none', dir: 'l', duration: 500, advanceAfter: null }, animations: [], comments: [] };
    try {
      let kind = 'titleContent'; let flags = 7; let notesId = null;
      const sa = sr && R.kid(sr, T.SlideAtom);
      if (sa) {
        const geom = R.u32(sa.body);
        kind = LAYOUT_KIND[geom] || 'titleContent'; notesId = R.u32(sa.body + 16); flags = R.u16(sa.body + 20);
        const fromMaster = masterKind.get(R.u32(sa.body + 12));
        if (fromMaster) kind = fromMaster;
        else if (geom === 16 && sr) kind = kindFromPlaceholders(slidePhTypes(sr)); // 2007+ files say "blank" for all
      }
      const layout = layoutFor(kind);
      slide.layoutId = layout.id;
      if (!(flags & 1)) slide.showMaster = false;
      const own = !(flags & 2) ? schemeOf(sr) : null;
      const ctx = { scheme: own, color: mkColor(own), texts, used: new Set(), layout, kind, bodies: { n: 0 }, anims: [] };
      if (sr) {
        const d = await readDrawing(sr, ctx);
        slide.elements = d.elements;
        if (d.background && !(flags & 4)) slide.background = d.background;
        const ss = R.kid(sr, T.SSSlideInfoAtom);
        if (ss) Object.assign(slide, transitionOf(R, ss, fidelity));
        slide.animations = buildAnimations(ctx.anims);
      }
      // texts of the slide list not placed by a shape → placeholders at the layout position
      texts.forEach((t, i) => {
        if (ctx.used.has(i) || !t.text.replace(/\r/g, '').trim()) return;
        const key = { 0: kind === 'title' ? 'ctrTitle' : 'title', 6: 'ctrTitle', 5: 'subTitle', 1: kind === 'title' ? 'subTitle' : 'body', 7: 'body', 8: 'body' }[t.type] || 'body';
        let k = key; let n = 1;
        while (slide.elements.some((e) => e.ph === k)) k = `${key.replace(/\d+$/, '')}${++n}`;
        const d = layout.ph[k] || layout.ph[key] || layout.ph.body || { ...layouts[1].ph.body };
        slide.elements.push({ id: uid('s'), type: 'shape', name: '', geom: 'rect', textbox: true, fill: null, line: null, x: d.x, y: d.y, w: d.w, h: d.h, rot: 0, ph: k, prompt: d.prompt || '', text: body(toParas(t, ctx.color), { anchor: d.anchor || 't', lvl: d.lvl ? JSON.parse(JSON.stringify(d.lvl)) : undefined }) });
      });
      // notes
      const nr = (notesId && notesById.get(notesId)) || (entry && notesBySlide.get(entry.slideId)) || null;
      if (nr) slide.notes = await notesText(nr);
    } catch {
      // keep whatever text we have
      if (!slide.elements.length && texts.length) slide.elements = texts.map((t, i) => ({ id: uid('s'), type: 'shape', name: '', geom: 'rect', textbox: true, fill: null, line: null, x: mu(288), y: mu(200 + i * 600), w: mu(W - 576), h: mu(500), rot: 0, text: body([para(t.text.replace(/\r/g, '\n').replace(/\v/g, '\n'))]) }));
    }
    slides.push(slide);
  }

  async function notesText(nr) {
    try {
      const ctx = { scheme: null, color: masterColor, texts: null, used: new Set(), layout: null, kind: 'notes', bodies: null };
      const ppd = R.kid(nr, T.PPDrawing);
      const dg = ppd && R.kid(ppd, T.DgContainer);
      // only the notes BODY placeholder is the notes text – the slide image, slide number ("*"), header, footer
      // and date placeholders of the notes page are not; plain text boxes count only when there is no body
      const body = []; const loose = [];
      for (const sp of R.deep(dg, T.SpContainer)) {
        const s = readSp(sp);
        if (!s.text || !s.text.text.trim()) continue;
        if (s.ph === 0x0c || s.ph === 0x06) body.push(s.text.text);
        else if (!s.ph) loose.push(s.text.text);
      }
      void ctx;
      return (body.length ? body : loose).join('\r').replace(/\r/g, '\n').replace(/\v/g, '\n').trim();
    } catch { return ''; }
  }

  const doc = newDoc({ size, theme: 'office' });
  doc.theme = theme;
  doc.layouts = layouts;
  doc.slides = slides;
  doc.meta = { ...doc.meta, title: '' };
  return { doc, fidelity: [...fidelity] };
}

function parseSlwt(R, l) {
  const out = [];
  let cur = null; let buf = [];
  const flush = () => { if (cur) cur.texts = readTextRecords(R, buf); buf = []; };
  for (const r of R.kids(l)) {
    if (r.type === T.SlidePersistAtom) { flush(); cur = { ref: R.u32(r.body), flags: R.u32(r.body + 4), slideId: R.u32(r.body + 12), texts: [] }; out.push(cur); } else buf.push(r);
  }
  flush();
  return out;
}

function transitionOf(R, ss, fidelity) {
  const time = R.i32(ss.body);
  const dir = R.u8(ss.body + 8); const type = R.u8(ss.body + 9); const fl = R.u16(ss.body + 10); const speed = R.u8(ss.body + 12);
  const map = { 0: 'none', 4: 'cover', 5: 'fade', 6: 'fade', 7: 'cover', 9: 'wipe', 10: 'wipe', 11: 'zoom', 13: 'split', 20: 'push', 23: 'fade' };
  let t = map[type];
  if (!t) { t = 'fade'; fidelity.add('transitions'); }
  const out = { transition: { type: t, dir: ['l', 'u', 'r', 'd'][dir] || 'l', duration: [1000, 750, 500][speed] || 750, advanceAfter: fl & 0x400 ? Math.max(0, time) : null } };
  if (fl & 0x4) out.hidden = true;
  return out;
}

async function inflate(data) {
  if (typeof DecompressionStream === 'undefined') return null;
  for (const fmt of ['deflate', 'deflate-raw']) {
    try {
      const ds = new DecompressionStream(fmt);
      const stream = new Blob([data]).stream().pipeThrough(ds);
      return new Uint8Array(await new Response(stream).arrayBuffer());
    } catch { /* try next */ }
  }
  return null;
}
async function inflateRaw(data) {
  if (typeof DecompressionStream === 'undefined') return null;
  try {
    const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch { return null; }
}

async function blipToUrl(R, blip, fidelity) {
  const two = blip.inst & 1;
  const uids = two ? 32 : 16;
  const t = blip.type;
  if (t === 0xf01a || t === 0xf01b || t === 0xf01c) { // EMF / WMF / PICT
    const h = blip.body + uids;
    const compression = R.u8(h + 32);
    let data = R.bytes(h + 34, blip.end - (h + 34));
    if (compression === 0) data = (await inflate(data)) || data;
    fidelity.add('effects');
    return bytesToDataUrl(data, t === 0xf01a ? 'image/x-emf' : t === 0xf01b ? 'image/x-wmf' : 'image/x-pict');
  }
  const start = blip.body + uids + 1;
  const data = R.bytes(start, blip.end - start);
  if (t === 0xf01d || t === 0xf02a) return bytesToDataUrl(data, 'image/jpeg');
  if (t === 0xf01e) return bytesToDataUrl(data, 'image/png');
  if (t === 0xf029) return bytesToDataUrl(data, 'image/tiff');
  if (t === 0xf01f) return bytesToDataUrl(dibToBmp(data), 'image/bmp');
  return null;
}

function dibToBmp(dib) {
  const dv = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const hs = dv.getUint32(0, true);
  const bits = dib.length >= 16 ? dv.getUint16(14, true) : 24;
  const comp = dib.length >= 20 ? dv.getUint32(16, true) : 0;
  const used = dib.length >= 36 ? dv.getUint32(32, true) : 0;
  const pal = (used || (bits <= 8 ? 1 << bits : 0)) * 4 + (comp === 3 && hs === 40 ? 12 : 0);
  const out = new Uint8Array(14 + dib.length);
  const o = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d;
  o.setUint32(2, out.length, true);
  o.setUint32(10, 14 + hs + pal, true);
  out.set(dib, 14);
  return out;
}
