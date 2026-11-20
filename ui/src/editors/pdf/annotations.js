/*
 * Own annotations of the PDF editor – pure data + geometry + pdf-lib output (no DOM, testable in Node).
 *
 * Every annotation lives in the PDF user space of its page (points, origin bottom-left, y up – exactly what
 * pdf.js' viewport.convertToPdfPoint returns and what pdf-lib draws with). Types:
 *
 *   highlight | underline | strike  { quads: [{ x, y, w, h }], color }
 *   ink        { paths: [[x, y, x, y, …]], color, width }
 *   rect | ellipse { x, y, w, h, color, width, fill? }
 *   line | arrow   { x1, y1, x2, y2, color, width }
 *   text       { x, y (baseline of the first line), text, size, color, rot }
 *   note       { x, y (bottom-left of the icon), text, color }
 *   stamp      { x, y, w, h, text, color, rot }
 *   image      { x, y, w, h, rot, png: Uint8Array, signature? }   (also signatures)
 *
 * `rot` (0/90/180/270) is the page's view rotation when the item was placed – text/stamp/image are drawn turned by
 * it so they appear upright the way the user saw them. (x, y) is then the bottom-left corner in that turned frame.
 */
import {
  rgb, degrees, PDFName, PDFHexString, PDFString, PDFNumber, PDFDict, PDFArray, PDFRef,
  drawRectangle, drawEllipse, drawLine, drawSvgPath, drawText, drawImage,
  pushGraphicsState, popGraphicsState, setGraphicsState, LineCapStyle,
} from 'pdf-lib';
import pako from 'pako';

export const COLORS = ['#ffd400', '#5fd068', '#4ea8ff', '#ff6b9a', '#ff5d3a', '#a46bff', '#111111', '#ffffff'];
export const STAMPS = [
  { text: 'APPROVED', color: '#1d8a4b' }, { text: 'DRAFT', color: '#2463c9' }, { text: 'CONFIDENTIAL', color: '#c42a3c' },
  { text: 'REJECTED', color: '#c42a3c' }, { text: 'PAID', color: '#1d8a4b' }, { text: 'FINAL', color: '#6a46ff' },
  { text: 'COPY', color: '#555555' }, { text: 'RECEIVED', color: '#946000' },
];
export const MARKUP = ['highlight', 'underline', 'strike'];
const SUBTYPE = { highlight: 'Highlight', underline: 'Underline', strike: 'StrikeOut', ink: 'Ink', rect: 'Square', ellipse: 'Circle', line: 'Line', arrow: 'Line', text: 'FreeText', note: 'Text', stamp: 'Stamp', image: 'Stamp' };

let seq = 0;
export const newId = () => `a${Date.now().toString(36)}${(seq++).toString(36)}`;

export function hexToRgb(hex) {
  const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(String(hex || ''));
  if (!m) return [0, 0, 0];
  let h = m[1];
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
}
const col = (hex) => { const [r, g, b] = hexToRgb(hex); return rgb(r, g, b); };

const rad = (d) => (d * Math.PI) / 180;
/** local frame (dx right, dy up, turned by rot) → page space offset */
export function turn(dx, dy, rot = 0) {
  const c = Math.cos(rad(rot)); const s = Math.sin(rad(rot));
  return [dx * c - dy * s, dx * s + dy * c];
}

/** corners of a turned box anchored at its bottom-left (x, y) */
function boxCorners(a) {
  const pts = [[0, 0], [a.w, 0], [a.w, a.h], [0, a.h]];
  return pts.map(([dx, dy]) => { const [ox, oy] = turn(dx, dy, a.rot || 0); return [a.x + ox, a.y + oy]; });
}

export function textLines(a) { return String(a.text || '').split(/\r?\n/); }
/** rough text width (Helvetica average) – the exact width comes from the font when saving */
/** rough text width (Helvetica average; CJK and other full-width characters count one em) */
export const approxWidth = (s, size) => { let w = 0; for (const ch of String(s)) w += /[ᄀ-ᅟ⺀-꓏가-힣豈-﫿︰-﹏＀-｠￠-￦]/.test(ch) || ch.codePointAt(0) > 0xffff ? 1 : 0.52; return w * size; };
export function textBoxSize(a) {
  const lines = textLines(a);
  return { w: Math.max(...lines.map((l) => approxWidth(l, a.size)), a.size), h: lines.length * a.size * 1.2 };
}

/** axis-aligned bounding box in page space: { x, y, w, h } */
export function bbox(a) {
  let pts;
  switch (a.type) {
    case 'highlight': case 'underline': case 'strike':
      pts = a.quads.flatMap((q) => [[q.x, q.y], [q.x + q.w, q.y + q.h]]); break;
    case 'ink': pts = []; for (const p of a.paths) for (let i = 0; i < p.length; i += 2) pts.push([p[i], p[i + 1]]); break;
    case 'rect': case 'ellipse': pts = [[a.x, a.y], [a.x + a.w, a.y + a.h]]; break;
    case 'line': case 'arrow': pts = [[a.x1, a.y1], [a.x2, a.y2]]; break;
    case 'text': {
      const { w, h } = textBoxSize(a);
      // the box starts one ascent above the baseline of the first line
      pts = [[0, a.size * 0.95], [w, a.size * 0.95], [w, a.size * 0.95 - h], [0, a.size * 0.95 - h]].map(([dx, dy]) => { const [ox, oy] = turn(dx, dy, a.rot || 0); return [a.x + ox, a.y + oy]; });
      break;
    }
    case 'note': pts = [[a.x, a.y], [a.x + 20, a.y + 20]]; break;
    default: pts = boxCorners(a);
  }
  if (!pts.length) return { x: 0, y: 0, w: 0, h: 0 };
  const xs = pts.map((p) => p[0]); const ys = pts.map((p) => p[1]);
  const pad = a.width ? a.width / 2 + (a.type === 'arrow' ? a.width * 3 : 0) : 0;
  const x = Math.min(...xs) - pad; const y = Math.min(...ys) - pad;
  return { x, y, w: Math.max(...xs) + pad - x, h: Math.max(...ys) + pad - y };
}

export function center(a) { const b = bbox(a); return [b.x + b.w / 2, b.y + b.h / 2]; }

/** move by (dx, dy) in page space – returns a new object */
export function translate(a, dx, dy) {
  switch (a.type) {
    case 'highlight': case 'underline': case 'strike': return { ...a, quads: a.quads.map((q) => ({ ...q, x: q.x + dx, y: q.y + dy })) };
    case 'ink': return { ...a, paths: a.paths.map((p) => p.map((v, i) => v + (i % 2 ? dy : dx))) };
    case 'line': case 'arrow': return { ...a, x1: a.x1 + dx, y1: a.y1 + dy, x2: a.x2 + dx, y2: a.y2 + dy };
    default: return { ...a, x: a.x + dx, y: a.y + dy };
  }
}

/** scale around the item's centre by factor f (uniform, works for every rotation) */
export function scale(a, f) {
  if (!(f > 0)) return a;
  const [cx, cy] = center(a);
  const sx = (x) => cx + (x - cx) * f; const sy = (y) => cy + (y - cy) * f;
  switch (a.type) {
    case 'highlight': case 'underline': case 'strike': return a; // bound to the text
    case 'ink': return { ...a, paths: a.paths.map((p) => p.map((v, i) => (i % 2 ? sy(v) : sx(v)))) };
    case 'line': case 'arrow': return { ...a, x1: sx(a.x1), y1: sy(a.y1), x2: sx(a.x2), y2: sy(a.y2) };
    case 'note': return a;
    case 'text': return { ...a, x: sx(a.x), y: sy(a.y), size: Math.max(4, Math.min(144, a.size * f)) };
    case 'rect': case 'ellipse': return { ...a, x: sx(a.x), y: sy(a.y), w: a.w * f, h: a.h * f };
    default: return { ...a, x: sx(a.x), y: sy(a.y), w: a.w * f, h: a.h * f };
  }
}

/** merge selection rectangles that sit on the same line into one quad each */
export function mergeQuads(rects) {
  const rs = rects.filter((r) => r.w > 0.5 && r.h > 0.5).map((r) => ({ ...r })).sort((a, b) => (b.y + b.h / 2) - (a.y + a.h / 2) || a.x - b.x);
  const out = [];
  for (const r of rs) {
    const last = out[out.length - 1];
    if (last && Math.abs((last.y + last.h / 2) - (r.y + r.h / 2)) < Math.min(last.h, r.h) * 0.5 && r.x <= last.x + last.w + r.h) {
      const x = Math.min(last.x, r.x); const y = Math.min(last.y, r.y);
      last.w = Math.max(last.x + last.w, r.x + r.w) - x; last.h = Math.max(last.y + last.h, r.y + r.h) - y; last.x = x; last.y = y;
    } else out.push(r);
  }
  return out;
}

/** "M x y L …" in page space for an ink path (y up) – used by the SVG overlay too */
export function inkPath(p) {
  if (p.length < 2) return '';
  let d = `M${p[0].toFixed(2)} ${p[1].toFixed(2)}`;
  if (p.length === 2) d += `L${(p[0] + 0.1).toFixed(2)} ${p[1].toFixed(2)}`;
  for (let i = 2; i < p.length; i += 2) d += `L${p[i].toFixed(2)} ${p[i + 1].toFixed(2)}`;
  return d;
}
/** pdf-lib's drawSvgPath flips y (SVG is y down) – give it the path with y negated */
const flipPath = (p) => inkPath(p.map((v, i) => (i % 2 ? -v : v)));

/** arrow head polygon points (page space) at (x2, y2) */
export function arrowHead(a) {
  const ang = Math.atan2(a.y2 - a.y1, a.x2 - a.x1);
  const len = Math.max(8, (a.width || 2) * 4);
  const p1 = [a.x2 - len * Math.cos(ang - 0.45), a.y2 - len * Math.sin(ang - 0.45)];
  const p2 = [a.x2 - len * Math.cos(ang + 0.45), a.y2 - len * Math.sin(ang + 0.45)];
  return [p1, [a.x2, a.y2], p2];
}

/* ------------------------------------------------------------------------------------------------------------ */
/* pdf-lib output                                                                                                */
/* ------------------------------------------------------------------------------------------------------------ */

/** text → hex string the standard font can show (characters outside WinAnsi become '?') */
export function encodeSafe(font, s) {
  try { return font.encodeText(s); } catch { /* fall through */ }
  let out = '';
  for (const ch of String(s)) { try { font.encodeText(ch); out += ch; } catch { out += '?'; } }
  return font.encodeText(out);
}
function safeWidth(font, s, size) { try { return font.widthOfTextAtSize(s, size); } catch { return approxWidth(s, size); } }

/** res.pick(text) → { font, name } (Unicode fonts) or the single res.font/res.fontName (Helvetica) */
const pickFont = (res, text) => (res.pick ? res.pick(text) : { font: res.font, name: res.fontName });

/**
 * Drawing operators for one annotation in page space.
 * res: { pick(text) → { font, name } | font (PDFFont) + fontName, gs (graphics state name for opacity or null), img (XObject name) }
 */
export function operatorsFor(a, res) {
  const c = col(a.color);
  switch (a.type) {
    case 'highlight':
      return a.quads.flatMap((q) => drawRectangle({ x: q.x, y: q.y, width: q.w, height: q.h, color: c, borderWidth: 0, rotate: degrees(0), xSkew: degrees(0), ySkew: degrees(0), graphicsState: res.gs || undefined }));
    case 'underline': case 'strike':
      return a.quads.flatMap((q) => {
        const t = Math.max(0.8, q.h * 0.07);
        const y = a.type === 'underline' ? q.y + t : q.y + q.h * 0.45;
        return drawLine({ start: { x: q.x, y }, end: { x: q.x + q.w, y }, thickness: t, color: c, lineCap: LineCapStyle.Butt });
      });
    case 'ink':
      return a.paths.flatMap((p) => drawSvgPath(flipPath(p), { x: 0, y: 0, scale: 1, borderColor: c, borderWidth: a.width || 2, borderLineCap: LineCapStyle.Round, rotate: degrees(0), graphicsState: res.gs || undefined }));
    case 'rect':
      return drawRectangle({ x: a.x, y: a.y, width: a.w, height: a.h, borderColor: c, borderWidth: a.width || 2, color: a.fill ? col(a.fill) : undefined, rotate: degrees(0), xSkew: degrees(0), ySkew: degrees(0), graphicsState: res.gs || undefined });
    case 'ellipse':
      return drawEllipse({ x: a.x + a.w / 2, y: a.y + a.h / 2, xScale: a.w / 2, yScale: a.h / 2, borderColor: c, borderWidth: a.width || 2, color: a.fill ? col(a.fill) : undefined, rotate: degrees(0), graphicsState: res.gs || undefined });
    case 'line': case 'arrow': {
      const ops = drawLine({ start: { x: a.x1, y: a.y1 }, end: { x: a.x2, y: a.y2 }, thickness: a.width || 2, color: c, lineCap: LineCapStyle.Round });
      if (a.type === 'arrow') {
        const [p1, p2, p3] = arrowHead(a);
        ops.push(...drawSvgPath(`M${p1[0]} ${-p1[1]}L${p2[0]} ${-p2[1]}L${p3[0]} ${-p3[1]}Z`, { x: 0, y: 0, color: c, borderColor: c, borderWidth: 1, rotate: degrees(0) }));
      }
      return ops;
    }
    case 'text': {
      const lh = a.size * 1.2;
      if (a.raster && res.img) {
        // characters no embedded font has: the text as a picture (drawn with the system fonts), the text itself stays in /Contents
        const { w, h } = textBoxSize(a);
        const [ox, oy] = turn(0, a.size * 0.95 - h, a.rot || 0);
        return drawImage(res.img, { x: a.x + ox, y: a.y + oy, width: w, height: h, rotate: degrees(a.rot || 0), xSkew: degrees(0), ySkew: degrees(0) });
      }
      return textLines(a).flatMap((line, i) => {
        const [ox, oy] = turn(0, -i * lh, a.rot || 0);
        const f = pickFont(res, line);
        return drawText(encodeSafe(f.font, line), { x: a.x + ox, y: a.y + oy, size: a.size, font: f.name, color: c, rotate: degrees(a.rot || 0), xSkew: degrees(0), ySkew: degrees(0) });
      });
    }
    case 'note': {
      const ops = drawRectangle({ x: a.x, y: a.y, width: 20, height: 20, color: col(a.color || '#ffd400'), borderColor: rgb(0.45, 0.35, 0), borderWidth: 0.8, rotate: degrees(0), xSkew: degrees(0), ySkew: degrees(0) });
      for (const k of [0, 1, 2]) ops.push(...drawLine({ start: { x: a.x + 4, y: a.y + 14 - k * 4 }, end: { x: a.x + (k === 2 ? 11 : 16), y: a.y + 14 - k * 4 }, thickness: 1, color: rgb(0.45, 0.35, 0) }));
      return ops;
    }
    case 'stamp': {
      const bw = Math.max(1.5, a.h * 0.07);
      const [ix, iy] = turn(bw / 2, bw / 2, a.rot || 0);
      const ops = drawRectangle({ x: a.x + ix, y: a.y + iy, width: a.w - bw, height: a.h - bw, borderColor: c, borderWidth: bw, rotate: degrees(a.rot || 0), xSkew: degrees(0), ySkew: degrees(0), graphicsState: res.gs || undefined });
      const size = a.h * 0.5;
      const f = pickFont(res, a.text);
      const tw = safeWidth(f.font, a.text, size);
      const [tx, ty] = turn((a.w - tw) / 2, a.h * 0.3, a.rot || 0);
      ops.push(...drawText(encodeSafe(f.font, a.text), { x: a.x + tx, y: a.y + ty, size, font: f.name, color: c, rotate: degrees(a.rot || 0), xSkew: degrees(0), ySkew: degrees(0), graphicsState: res.gs || undefined }));
      return ops;
    }
    case 'image':
      return drawImage(res.img, { x: a.x, y: a.y, width: a.w, height: a.h, rotate: degrees(a.rot || 0), xSkew: degrees(0), ySkew: degrees(0) });
    default: return [];
  }
}

const opacityOf = (a) => (a.type === 'highlight' ? 0.4 : a.opacity != null ? a.opacity : null);

/** private key of our annotations: the editable model as JSON (other apps ignore it, we read it back) */
export const MODEL_KEY = 'HWOffice';
const EDITABLE = new Set(['highlight', 'underline', 'strike', 'ink', 'rect', 'ellipse', 'line', 'arrow', 'text', 'note', 'stamp']);
const utf8Hex = (s) => PDFHexString.of(Array.from(new TextEncoder().encode(s), (b) => b.toString(16).padStart(2, '0')).join(''));

/**
 * fonts: { pick(text) → { font, ref, key } } – the PDF font for a text (Helvetica, or an embedded Unicode font);
 * nameFor(f) → its resource name inside the current resources
 */
function fontUser(fonts, nameFor) {
  const used = new Map();
  return {
    used,
    pick: (text) => {
      const f = fonts.pick(text);
      if (!used.has(f.key)) used.set(f.key, { ref: f.ref, name: nameFor(f) });
      return { font: f.font, name: used.get(f.key).name };
    },
  };
}
/** the old single-font call (font = PDFFont) still works */
const asFonts = (font, fonts) => fonts || { pick: () => ({ font, ref: font.ref, key: 'OfHelv' }) };

/**
 * Writes the annotations of one page. mode 'burn' = into the page content (everyone sees them, nobody can edit them),
 * 'annot' = real PDF annotations with an appearance stream (other apps can show, move and delete them; images and
 * signatures are always burned in so they cannot be removed by accident). Annotations carry our model under
 * /HWOffice, so the editor makes them editable again when the file is opened later (readOwnAnnotations).
 * Returns the number of editable annotations written.
 */
export async function writeAnnotations(doc, page, list, { mode = 'annot', font, fonts, author = '' } = {}) {
  if (!list || !list.length) return 0;
  const ctx = doc.context;
  const F = asFonts(font, fonts);
  const pageNames = new Map();
  let written = 0;
  for (const a of list) {
    let imgRef = null;
    if (a.type === 'image') imgRef = (await doc.embedPng(a.png)).ref;
    else if (a.raster && a.raster.png) imgRef = (await doc.embedPng(a.raster.png)).ref;
    const op = opacityOf(a);
    const gsRef = op != null ? ctx.register(ctx.obj({ Type: 'ExtGState', ca: op, CA: op, ...(a.type === 'highlight' ? { BM: 'Multiply' } : {}) })) : null;
    const asAnnot = mode === 'annot' && a.type !== 'image';
    if (!asAnnot) {
      const gs = gsRef ? page.node.newExtGState('OfGS', gsRef) : null;
      const img = imgRef ? page.node.newXObject('OfImg', imgRef) : null;
      const fu = fontUser(F, (f) => { if (!pageNames.has(f.key)) pageNames.set(f.key, page.node.newFontDictionary(f.key, f.ref)); return pageNames.get(f.key); });
      const ops = operatorsFor(a, { pick: fu.pick, gs, img });
      page.pushOperators(pushGraphicsState(), ...(gs ? [setGraphicsState(gs)] : []), ...ops, popGraphicsState());
      continue;
    }
    // appearance stream in page coordinates: BBox = Rect, identity matrix
    const b = bbox(a);
    const pad = 2;
    const rect = [b.x - pad, b.y - pad, b.x + b.w + pad, b.y + b.h + pad];
    const fu = fontUser(F, (f) => f.key);
    const ops = operatorsFor(a, { pick: fu.pick, gs: gsRef ? 'OfGS' : null, img: imgRef ? 'OfImg' : null });
    const resDict = {};
    if (imgRef) resDict.XObject = { OfImg: imgRef };
    if (fu.used.size) { resDict.Font = {}; for (const [k, v] of fu.used) resDict.Font[k] = v.ref; }
    if (gsRef) resDict.ExtGState = { OfGS: gsRef };
    const ap = ctx.register(ctx.formXObject([pushGraphicsState(), ...(gsRef ? [setGraphicsState('OfGS')] : []), ...ops, popGraphicsState()], { BBox: rect, Resources: resDict }));
    const [r, g, bl] = hexToRgb(a.color || '#000000');
    const dict = {
      Type: 'Annot', Subtype: SUBTYPE[a.type], Rect: rect, F: 4, C: [r, g, bl], AP: { N: ap }, NM: PDFString.of(a.id || newId()),
      M: PDFString.fromDate(new Date()), P: page.ref,
    };
    if (op != null) dict.CA = op;
    if (a.width) dict.BS = { W: a.width, S: 'S' };
    if (a.type === 'highlight' || a.type === 'underline' || a.type === 'strike') {
      dict.QuadPoints = a.quads.flatMap((q) => [q.x, q.y + q.h, q.x + q.w, q.y + q.h, q.x, q.y, q.x + q.w, q.y]);
    } else if (a.type === 'ink') {
      dict.InkList = a.paths;
    } else if (a.type === 'line' || a.type === 'arrow') {
      dict.L = [a.x1, a.y1, a.x2, a.y2];
      if (a.type === 'arrow') dict.LE = ['None', 'ClosedArrow'];
      dict.IC = [r, g, bl];
    } else if (a.type === 'note') {
      dict.Name = 'Comment'; dict.Rect = [a.x, a.y, a.x + 20, a.y + 20];
    } else if (a.type === 'text') {
      dict.DA = PDFString.of(`/Helv ${a.size} Tf ${r.toFixed(3)} ${g.toFixed(3)} ${bl.toFixed(3)} rg`);
      if (a.rot) dict.Rotate = a.rot;
    }
    if (a.fill && (a.type === 'rect' || a.type === 'ellipse')) dict.IC = hexToRgb(a.fill);
    Object.keys(dict).forEach((k) => dict[k] === undefined && delete dict[k]);
    const annot = ctx.obj(dict);
    // texts must be strings, not names
    const contents = a.type === 'stamp' ? a.text : a.text || '';
    if (contents) annot.set(PDFName.of('Contents'), PDFHexString.fromText(contents));
    if (author) annot.set(PDFName.of('T'), PDFHexString.fromText(author));
    if (a.type === 'stamp') annot.set(PDFName.of('Name'), PDFName.of(String(a.text || 'Draft').replace(/[^A-Za-z0-9]/g, '') || 'Draft'));
    if (a.rot && a.type === 'text') annot.set(PDFName.of('Rotate'), PDFNumber.of(a.rot));
    if (EDITABLE.has(a.type)) { const { id, raster, ...m } = a; // eslint-disable-line no-unused-vars
      annot.set(PDFName.of(MODEL_KEY), utf8Hex(JSON.stringify({ v: 1, id, ...m }))); written++; }
    page.node.addAnnot(ctx.register(annot));
  }
  return written;
}

/** marker in the catalog – a stream, so its dictionary never sits inside a compressed object stream (quick byte scan) */
export const MARKER = 'HWOfficeAnnots';
export function markDocument(doc) {
  const ctx = doc.context;
  const s = ctx.stream(new TextEncoder().encode('Hostwerk Office: the annotations of this file stay editable in Hostwerk.'), { Type: MARKER });
  doc.catalog.set(PDFName.of(MARKER), ctx.register(s));
}
function findBytes(bytes, str, from = 0, to = bytes.length) {
  const pat = Array.from(str, (c) => c.charCodeAt(0));
  const f = pat[0];
  outer: for (let i = from, n = Math.min(to, bytes.length) - pat.length; i <= n; i++) {
    if (bytes[i] !== f) continue;
    for (let k = 1; k < pat.length; k++) if (bytes[i + k] !== pat[k]) continue outer;
    return i;
  }
  return -1;
}
/**
 * Does the file (maybe) carry our annotations? Our catalog marker, our /HWOffice key in plain objects, or – when
 * another app re-saved the file with compressed object streams (and dropped the marker) – /HWOffice inside an
 * inflated /ObjStm.
 */
export function hasMarker(bytes) {
  if (findBytes(bytes, `/${MARKER}`) >= 0 || findBytes(bytes, `/${MODEL_KEY}`) >= 0) return true;
  if (bytes.length > 80 * 1024 * 1024) return false;
  let at = 0; let n = 0;
  const key = Array.from(MODEL_KEY, (c) => c.charCodeAt(0));
  while ((at = findBytes(bytes, '/ObjStm', at)) >= 0 && n++ < 2000) {
    const st = findBytes(bytes, 'stream', at, at + 4096);
    if (st < 0) { at += 7; continue; }
    let p = st + 6; if (bytes[p] === 13) p++; if (bytes[p] === 10) p++;
    const end = findBytes(bytes, 'endstream', p);
    if (end < 0) break;
    try {
      const raw = pako.inflate(bytes.subarray(p, end));
      for (let i = 0; i + key.length < raw.length; i++) { if (raw[i] !== key[0]) continue; let ok = true; for (let k = 1; k < key.length; k++) if (raw[i + k] !== key[k]) { ok = false; break; } if (ok) return true; }
    } catch { /* not deflated / damaged */ }
    at = end;
  }
  return false;
}

/**
 * Takes our annotations out of the document again (pdf-lib PDFDocument, changed in place) → [[model …] per page
 * index]. One that another app moved (its Rect no longer matches our model) stays an ordinary annotation.
 */
export function readOwnAnnotations(doc) {
  const out = [];
  doc.getPages().forEach((page, i) => {
    out[i] = [];
    const annots = page.node.lookup(PDFName.of('Annots'));
    if (!(annots instanceof PDFArray)) return;
    const keep = [];
    for (let k = 0; k < annots.size(); k++) {
      const ref = annots.get(k);
      const d = ref instanceof PDFRef ? doc.context.lookup(ref) : ref;
      const raw = d instanceof PDFDict ? d.get(PDFName.of(MODEL_KEY)) : null;
      let m = null;
      if (raw && raw.asBytes) { try { m = JSON.parse(new TextDecoder().decode(raw.asBytes())); } catch { m = null; } }
      if (!m || m.v !== 1 || !EDITABLE.has(m.type)) { keep.push(ref); continue; }
      const { v, ...model0 } = m; // eslint-disable-line no-unused-vars
      let model = model0;
      // another app may have moved it (Rect shifted, same size → move our model along) or recomputed the Rect a bit
      // (a few points of extra margin); a resized one stays an ordinary annotation
      const rect = d.lookup(PDFName.of('Rect'));
      if (rect instanceof PDFArray) {
        const b = bbox(model); const r = rect.asArray().map((x) => (x.asNumber ? x.asNumber() : 0));
        const w = r[2] - r[0]; const h = r[3] - r[1];
        if (m.type === 'note') { if (Math.abs(r[0] - model.x) > 1.5 || Math.abs(r[1] - model.y) > 1.5) model = translate(model, r[0] - model.x, r[1] - model.y); } else {
          const cx = (r[0] + r[2]) / 2 - (b.x + b.w / 2); const cy = (r[1] + r[3]) / 2 - (b.y + b.h / 2);
          if (Math.abs(w - (b.w + 4)) > 8 || Math.abs(h - (b.h + 4)) > 8) { keep.push(ref); continue; }
          if (Math.abs(cx) > 1.5 || Math.abs(cy) > 1.5) model = translate(model, cx, cy);
        }
      }
      // a comment text changed in another app
      if (m.type === 'note' || m.type === 'text') {
        const c = d.lookup(PDFName.of('Contents'));
        const txt = c && c.decodeText ? c.decodeText() : null;
        if (txt != null && txt !== (model.text || '')) model = { ...model, text: txt };
      }
      out[i].push({ ...model, id: model.id || newId() });
      if (ref instanceof PDFRef) doc.context.delete(ref);
    }
    if (keep.length === annots.size()) return;
    if (keep.length) page.node.set(PDFName.of('Annots'), doc.context.obj(keep));
    else page.node.delete(PDFName.of('Annots'));
  });
  return out;
}
