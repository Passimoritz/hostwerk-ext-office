/*
 * Vector PDF of slides without the Office engine (browser). Each slide is rendered once at scale 1 (rotation switched
 * off so everything can be measured), then the DOM is walked: SVG shapes → PDF paths (solid and gradient fills,
 * strokes with dashes/caps/joins), pictures → images (JPEG passes through), table cells → rectangles and borders, text →
 * real, selectable PDF text in the standard fonts (each word placed and horizontally scaled to its measured width).
 * Whatever cannot be drawn as vectors (charts with markers, shadows, clip-path pictures, non-Latin text, video …)
 * becomes a transparent picture of just that object. Rotation is applied per object.
 *   vectorPdf(pages: [{ node, size: { w, h }, model: Map(id → el) }], { title }) → Uint8Array
 */
import { normalizePath, segments } from '../render/custom.js';

/* ---------------------------------------------------------------- small helpers */
const N = (n) => { const v = Math.round(n * 1000) / 1000; return Object.is(v, -0) ? '0' : String(v); };
const enc = new TextEncoder();
function parseColor(c) {
  if (!c) return null;
  let m = /^rgba?\(([^)]+)\)/.exec(c);
  if (m) { const p = m[1].split(/[\s,/]+/).filter(Boolean).map(parseFloat); const a = p.length > 3 ? p[3] : 1; return a <= 0.001 ? null : { r: p[0] / 255, g: p[1] / 255, b: p[2] / 255, a }; }
  m = /^#([0-9a-f]{3,8})$/i.exec(c);
  if (m) { let h = m[1]; if (h.length <= 4) h = h.split('').map((x) => x + x).join(''); const a = h.length === 8 ? parseInt(h.slice(6, 8), 16) / 255 : 1; return { r: parseInt(h.slice(0, 2), 16) / 255, g: parseInt(h.slice(2, 4), 16) / 255, b: parseInt(h.slice(4, 6), 16) / 255, a }; }
  if (c === 'black') return { r: 0, g: 0, b: 0, a: 1 }; if (c === 'white') return { r: 1, g: 1, b: 1, a: 1 };
  return null;
}
const rgb = (c) => `${N(c.r)} ${N(c.g)} ${N(c.b)}`;
// 2D affine [a b c d e f]
const mul = (m, n) => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
const rotAbout = (deg, cx, cy) => { const a = (deg * Math.PI) / 180; const c = Math.cos(a); const s = Math.sin(a); return [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]; };
const cm = (m) => `${m.map(N).join(' ')} cm\n`;
const ID = [1, 0, 0, 1, 0, 0];
// WinAnsi (cp1252) encoding of a string – null when a character is not in it
const CP = { 0x20ac: 128, 0x201a: 130, 0x192: 131, 0x201e: 132, 0x2026: 133, 0x2020: 134, 0x2021: 135, 0x2c6: 136, 0x2030: 137, 0x160: 138, 0x2039: 139, 0x152: 140, 0x17d: 142, 0x2018: 145, 0x2019: 146, 0x201c: 147, 0x201d: 148, 0x2022: 149, 0x2013: 150, 0x2014: 151, 0x2dc: 152, 0x2122: 153, 0x161: 154, 0x203a: 155, 0x153: 156, 0x17e: 158, 0x178: 159 };
function winAnsi(s) {
  const out = [];
  for (const ch of s) { const c = ch.codePointAt(0); if (c === 0x200b || c === 0xfeff) continue; if (c >= 32 && c < 127) out.push(c); else if (c >= 160 && c <= 255) out.push(c); else if (CP[c]) out.push(CP[c]); else if (c === 9 || c === 10) out.push(32); else return null; }
  return out;
}
const pdfStr = (codes) => `(${codes.map((c) => (c === 40 || c === 41 || c === 92 ? `\\${String.fromCharCode(c)}` : c < 32 || c > 126 ? `\\${c.toString(8).padStart(3, '0')}` : String.fromCharCode(c))).join('')})`;
// Helvetica widths (1/1000 em) for 32..126; others 556 (good enough – Tz fits each word to its measured width)
const HW = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584];
const widthOf = (codes, mono) => codes.reduce((a, c) => a + (mono ? 600 : c >= 32 && c <= 126 ? HW[c - 32] : 556), 0) / 1000;
function fontKey(cs) {
  const fam = (cs.fontFamily || '').toLowerCase();
  const bold = (parseInt(cs.fontWeight, 10) || 400) >= 600 || cs.fontWeight === 'bold';
  const it = cs.fontStyle === 'italic' || cs.fontStyle === 'oblique';
  const mono = /mono|courier|consolas|menlo/.test(fam);
  const serif = !mono && /(^|[\s,"'])(serif|times|georgia|cambria|garamond|palatino|book|constantia|sitka)/.test(fam) && !/sans/.test(fam.split(',')[0]);
  const base = mono ? 'Courier' : serif ? 'Times' : 'Helvetica';
  const name = base === 'Times' ? (bold && it ? 'Times-BoldItalic' : bold ? 'Times-Bold' : it ? 'Times-Italic' : 'Times-Roman') : `${base}${bold && it ? '-BoldOblique' : bold ? '-Bold' : it ? '-Oblique' : ''}`;
  return { name, mono };
}
async function deflate(bytes) {
  if (typeof CompressionStream === 'undefined') return null;
  const cs = new CompressionStream('deflate');
  const w = cs.writable.getWriter(); w.write(bytes); w.close();
  return new Uint8Array(await new Response(cs.readable).arrayBuffer());
}
function loadImg(src) { return new Promise((res) => { const i = new Image(); const tm = setTimeout(() => res(null), 8000); i.onload = () => { clearTimeout(tm); res(i); }; i.onerror = () => { clearTimeout(tm); res(null); }; i.src = src; }); }

/* ---------------------------------------------------------------- the document */
class Pdf {
  constructor() { this.objs = []; this.fonts = {}; this.images = new Map(); this.gs = {}; }
  add(body) { this.objs.push(body); return this.objs.length + 2; } // 1 catalog, 2 pages reserved
  font(name) { if (!this.fonts[name]) this.fonts[name] = { id: this.add(`<< /Type /Font /Subtype /Type1 /BaseFont /${name} /Encoding /WinAnsiEncoding >>`), key: `F${Object.keys(this.fonts).length + 1}` }; return this.fonts[name].key; }
  alpha(a) { const k = `GS${Math.round(a * 100)}`; if (!this.gs[k]) this.gs[k] = this.add(`<< /Type /ExtGState /ca ${N(a)} /CA ${N(a)} >>`); return k; }
  async image(src, canvasFor) {
    if (this.images.has(src)) return this.images.get(src);
    let res = null;
    const jpeg = /^data:image\/jpe?g;base64,/i.exec(src);
    const img = await loadImg(src);
    if (!img) return null;
    const w = img.naturalWidth || 1; const h = img.naturalHeight || 1;
    if (jpeg) {
      const b = Uint8Array.from(atob(src.slice(jpeg[0].length)), (c) => c.charCodeAt(0));
      const id = this.add({ dict: `<< /Type /XObject /Subtype /Image /Width ${w} /Height ${h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${b.length} >>`, data: b });
      res = { id, key: `Im${id}`, w, h };
    } else {
      const c = canvasFor ? canvasFor(img) : (() => { const cv = document.createElement('canvas'); const s = Math.min(1, 2400 / Math.max(w, h)); cv.width = Math.max(1, Math.round(w * s)); cv.height = Math.max(1, Math.round(h * s)); cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height); return cv; })();
      res = await this.rgba(c);
    }
    this.images.set(src, res);
    return res;
  }
  /** canvas → image with alpha (Flate, SMask) */
  async rgba(canvas) {
    const w = canvas.width; const h = canvas.height;
    const d = canvas.getContext('2d').getImageData(0, 0, w, h).data;
    const rgbB = new Uint8Array(w * h * 3); const aB = new Uint8Array(w * h); let anyA = false;
    for (let i = 0, j = 0, k = 0; i < d.length; i += 4) { rgbB[j++] = d[i]; rgbB[j++] = d[i + 1]; rgbB[j++] = d[i + 2]; aB[k++] = d[i + 3]; if (d[i + 3] !== 255) anyA = true; }
    const zc = await deflate(rgbB); const za = anyA ? await deflate(aB) : null;
    let smask = '';
    if (anyA) { const sid = this.add({ dict: `<< /Type /XObject /Subtype /Image /Width ${w} /Height ${h} /ColorSpace /DeviceGray /BitsPerComponent 8${za ? ' /Filter /FlateDecode' : ''} /Length ${(za || aB).length} >>`, data: za || aB }); smask = ` /SMask ${sid} 0 R`; }
    const id = this.add({ dict: `<< /Type /XObject /Subtype /Image /Width ${w} /Height ${h} /ColorSpace /DeviceRGB /BitsPerComponent 8${zc ? ' /Filter /FlateDecode' : ''}${smask} /Length ${(zc || rgbB).length} >>`, data: zc || rgbB });
    return { id, key: `Im${id}`, w, h };
  }
  shading(kind, coords, stops) {
    // stitched exponential functions for any number of stops
    const fn = stops.length === 2 ? `<< /FunctionType 2 /Domain [0 1] /C0 [${rgb(stops[0].c)}] /C1 [${rgb(stops[1].c)}] /N 1 >>`
      : `<< /FunctionType 3 /Domain [0 1] /Functions [${stops.slice(1).map((s, i) => `<< /FunctionType 2 /Domain [0 1] /C0 [${rgb(stops[i].c)}] /C1 [${rgb(s.c)}] /N 1 >>`).join(' ')}] /Bounds [${stops.slice(1, -1).map((s) => N(s.o)).join(' ')}] /Encode [${stops.slice(1).map(() => '0 1').join(' ')}] >>`;
    const id = this.add(`<< /ShadingType ${kind === 'radial' ? 3 : 2} /ColorSpace /DeviceRGB /Coords [${coords.map(N).join(' ')}] /Function ${fn} /Extend [true true] >>`);
    return { id, key: `Sh${id}` };
  }
  async build(pages, title) {
    const pageIds = [];
    for (const p of pages) {
      const content = enc.encode(p.content);
      const z = await deflate(content);
      const cid = this.add({ dict: `<<${z ? ' /Filter /FlateDecode' : ''} /Length ${(z || content).length} >>`, data: z || content });
      const res = `<< /Font << ${Object.values(this.fonts).map((f) => `/${f.key} ${f.id} 0 R`).join(' ')} >> /XObject << ${[...this.images.values(), ...p.extraImages].filter(Boolean).map((i) => `/${i.key} ${i.id} 0 R`).join(' ')} >> /ExtGState << ${Object.entries(this.gs).map(([k, v]) => `/${k} ${v} 0 R`).join(' ')} >> /Shading << ${p.shadings.map((s) => `/${s.key} ${s.id} 0 R`).join(' ')} >> >>`;
      pageIds.push(this.add(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${N(p.w)} ${N(p.h)}] /Resources ${res} /Contents ${cid} 0 R >>`));
    }
    const info = this.add(`<< /Producer (Hostwerk Office) /Title ${pdfStr(winAnsi(title || '') || [])} >>`);
    const chunks = []; let len = 0; const off = [];
    const push = (x) => { const b = typeof x === 'string' ? enc.encode(x) : x; chunks.push(b); len += b.length; };
    push('%PDF-1.5\n%\xE2\xE3\xCF\xD3\n');
    const all = [`<< /Type /Catalog /Pages 2 0 R >>`, `<< /Type /Pages /Count ${pageIds.length} /Kids [${pageIds.map((i) => `${i} 0 R`).join(' ')}] >>`, ...this.objs];
    all.forEach((o, i) => { off[i + 1] = len; push(`${i + 1} 0 obj\n`); if (typeof o === 'string') push(o); else { push(`${o.dict}\nstream\n`); push(o.data); push('\nendstream'); } push('\nendobj\n'); });
    const x = len; const n = all.length + 1;
    let xr = `xref\n0 ${n}\n0000000000 65535 f \n`; for (let k = 1; k < n; k++) xr += `${String(off[k]).padStart(10, '0')} 00000 n \n`;
    push(xr); push(`trailer\n<< /Size ${n} /Root 1 0 R /Info ${info} 0 R >>\nstartxref\n${x}\n%%EOF\n`);
    const out = new Uint8Array(len); let o = 0; for (const c of chunks) { out.set(c, o); o += c.length; }
    return out;
  }
}

/* ---------------------------------------------------------------- one page */
function pathOps(d) {
  let out = ''; let cx = 0; let cy = 0;
  for (const s of segments(normalizePath(d))) {
    const p = s.p;
    if (s.c === 'M') { out += `${N(p[0])} ${N(p[1])} m\n`; [cx, cy] = p; }
    else if (s.c === 'L') { out += `${N(p[0])} ${N(p[1])} l\n`; [cx, cy] = p; }
    else if (s.c === 'C') { out += `${p.map(N).join(' ')} c\n`; cx = p[4]; cy = p[5]; }
    else if (s.c === 'Q') { const c1x = cx + (2 / 3) * (p[0] - cx); const c1y = cy + (2 / 3) * (p[1] - cy); const c2x = p[2] + (2 / 3) * (p[0] - p[2]); const c2y = p[3] + (2 / 3) * (p[1] - p[3]); out += `${[c1x, c1y, c2x, c2y, p[2], p[3]].map(N).join(' ')} c\n`; cx = p[2]; cy = p[3]; }
    else out += 'h\n';
  }
  return out;
}
function shapeD(el) {
  const t = el.tagName.toLowerCase(); const a = (k) => parseFloat(el.getAttribute(k) || '0');
  if (t === 'path') return el.getAttribute('d') || '';
  if (t === 'rect') { const x = a('x'); const y = a('y'); const w = a('width'); const hh = a('height'); return `M${x} ${y}H${x + w}V${y + hh}H${x}Z`; }
  if (t === 'circle' || t === 'ellipse') { const cx = a('cx'); const cy = a('cy'); const rx = t === 'circle' ? a('r') : a('rx'); const ry = t === 'circle' ? a('r') : a('ry'); return `M${cx - rx} ${cy}A${rx} ${ry} 0 1 0 ${cx + rx} ${cy}A${rx} ${ry} 0 1 0 ${cx - rx} ${cy}Z`; }
  if (t === 'line') return `M${a('x1')} ${a('y1')}L${a('x2')} ${a('y2')}`;
  if (t === 'polygon' || t === 'polyline') { const p = (el.getAttribute('points') || '').trim().split(/[\s,]+/).map(Number); let d = `M${p[0]} ${p[1]}`; for (let i = 2; i < p.length; i += 2) d += `L${p[i]} ${p[i + 1]}`; return t === 'polygon' ? `${d}Z` : d; }
  return null;
}

/** incremental: addPage() while the page's DOM is mounted, finish() → bytes */
export function createVectorPdf({ raster } = {}) {
  const pdf = new Pdf();
  const out = [];
  const addPage = async (pg) => {
    const { node, size, model, background } = pg;
    const W = size.w; const H = size.h;
    const shadings = []; const extraImages = [];
    let c = `1 0 0 -1 0 ${N(H)} cm\n`; // top-left origin like the slide
    const stage = node.querySelector('.of-slides-stage') || node;
    const O = stage.getBoundingClientRect();
    const rel = (r) => ({ x: r.left - O.left, y: r.top - O.top, w: r.width, h: r.height });
    // background
    if (background && background.css) {
      const bc = parseColor(background.css);
      if (bc) c += `${rgb(bc)} rg 0 0 ${N(W)} ${N(H)} re f\n`;
    }
    if (background && background.grad) {
      const g = background.grad;
      const st = g.stops.map((s) => ({ o: s.pos / 100, c: parseColor(s.css) || { r: 1, g: 1, b: 1 } }));
      const a = ((g.angle || 0) * Math.PI) / 180;
      const sh = pdf.shading(g.kind === 'radial' ? 'radial' : 'linear', g.kind === 'radial' ? [W / 2, H / 2, 0, W / 2, H / 2, Math.hypot(W, H) / 2] : [W / 2 - (Math.cos(a) * W) / 2, H / 2 - (Math.sin(a) * H) / 2, W / 2 + (Math.cos(a) * W) / 2, H / 2 + (Math.sin(a) * H) / 2], st);
      shadings.push(sh); c += `q 0 0 ${N(W)} ${N(H)} re W n /${sh.key} sh Q\n`;
    }
    if (background && background.image) {
      const im = await pdf.image(background.image);
      if (im) c += `q ${N(W)} 0 0 ${N(-H)} 0 ${N(H)} cm /${im.key} Do Q\n`;
    }
    // raster one element (transparent picture) – the fallback for whatever has no vector form here
    const rasterEl = async (elNode, M) => {
      if (!raster) return '';
      const r = rel(elNode.getBoundingClientRect());
      const cv = await raster(elNode, r.w, r.h);
      if (!cv) return '';
      const im = await pdf.rgba(cv);
      extraImages.push(im);
      return `q ${cm(M)}${N(r.w)} 0 0 ${N(-r.h)} ${N(r.x)} ${N(r.y + r.h)} cm /${im.key} Do Q\n`;
    };
    const textOf = (elNode, M) => {
      // every text node → words with their measured boxes
      let s = '';
      const walker = document.createTreeWalker(elNode, NodeFilter.SHOW_TEXT);
      const range = document.createRange();
      for (let tn = walker.nextNode(); tn; tn = walker.nextNode()) {
        const str = tn.nodeValue;
        if (!str || !str.replace(/[\s​]/g, '')) continue;
        const pe = tn.parentElement; const cs = getComputedStyle(pe);
        if (/^vertical/.test(cs.writingMode || '')) throw new Error('vertical');
        if (cs.visibility === 'hidden' || cs.display === 'none') continue;
        const col = parseColor(cs.color) || { r: 0, g: 0, b: 0, a: 1 };
        const fsz = parseFloat(cs.fontSize) || 12;
        const f = fontKey(cs);
        const key = pdf.font(f.name);
        const deco = cs.textDecorationLine || '';
        const re = /\S+/g; let m;
        while ((m = re.exec(str))) {
          const codes = winAnsi(m[0]);
          if (!codes) throw new Error('charset');
          range.setStart(tn, m.index); range.setEnd(tn, m.index + m[0].length);
          const rects = [...range.getClientRects()].filter((q) => q.width > 0.2);
          if (!rects.length) continue;
          const q = rel(rects[0]); const qw = rects.reduce((a, x) => a + x.width, 0);
          const nat = widthOf(codes, f.mono) * fsz;
          const tz = nat > 0 ? Math.max(50, Math.min(200, (qw / nat) * 100)) : 100;
          const base = q.y + q.h / 2 + fsz * 0.33;
          s += `BT /${key} ${N(fsz)} Tf ${N(tz)} Tz ${rgb(col)} rg 1 0 0 -1 ${N(q.x)} ${N(base)} Tm ${pdfStr(codes)} Tj ET\n`;
          if (/underline/.test(deco)) s += `${rgb(col)} RG ${N(Math.max(0.5, fsz / 16))} w ${N(q.x)} ${N(base + fsz * 0.12)} m ${N(q.x + qw)} ${N(base + fsz * 0.12)} l S\n`;
          if (/line-through/.test(deco)) s += `${rgb(col)} RG ${N(Math.max(0.5, fsz / 16))} w ${N(q.x)} ${N(base - fsz * 0.3)} m ${N(q.x + qw)} ${N(base - fsz * 0.3)} l S\n`;
        }
      }
      return s ? `q ${cm(M)}${s}Q\n` : '';
    };
    const svgOf = (svg, M) => {
      let s = '';
      const so = rel(svg.getBoundingClientRect());
      const vb = svg.viewBox && svg.viewBox.baseVal && svg.viewBox.baseVal.width ? svg.viewBox.baseVal : null;
      const sw = parseFloat(svg.getAttribute('width')) || so.w; const shh = parseFloat(svg.getAttribute('height')) || so.h;
      let local = [1, 0, 0, 1, so.x, so.y];
      if (vb) local = mul(local, [sw / vb.width, 0, 0, shh / vb.height, -vb.x * (sw / vb.width), -vb.y * (shh / vb.height)]);
      const tr = svg.style.transform || '';
      if (/scaleX\(-1\)/.test(tr)) local = mul([-1, 0, 0, 1, so.x * 2 + so.w, 0], local);
      if (/scaleY\(-1\)/.test(tr)) local = mul([1, 0, 0, -1, 0, so.y * 2 + so.h], local);
      if (svg.querySelector('[marker-start],[marker-end],[marker-mid],pattern,image,text,filter')) throw new Error('svg-feature');
      if (svg.style.filter && svg.style.filter !== 'none') throw new Error('filter');
      for (const el of svg.querySelectorAll('path,rect,circle,ellipse,line,polygon,polyline')) {
        if (el.closest('defs,marker,clipPath,mask,pattern')) continue;
        const d = shapeD(el); if (!d) continue;
        const cs = getComputedStyle(el);
        const fillV = el.getAttribute('fill') || cs.fill; const strokeV = el.getAttribute('stroke') || cs.stroke;
        const ops = pathOps(d);
        const ctm = el.transform && el.transform.baseVal && el.transform.baseVal.consolidate ? el.transform.baseVal.consolidate() : null;
        const mm = ctm ? [ctm.matrix.a, ctm.matrix.b, ctm.matrix.c, ctm.matrix.d, ctm.matrix.e, ctm.matrix.f] : ID;
        let part = '';
        const op = parseFloat(cs.opacity || '1');
        if (/^url\(/.test(fillV || '')) {
          const id = /#([^)"']+)/.exec(fillV); const g = id && svg.querySelector(`#${CSS.escape(id[1])}`);
          if (!g || !/Gradient$/.test(g.tagName)) throw new Error('fill');
          const stops = [...g.querySelectorAll('stop')].map((st) => ({ o: Math.max(0, Math.min(1, parseFloat(st.getAttribute('offset')) / (String(st.getAttribute('offset')).includes('%') ? 100 : 1))), c: parseColor(st.getAttribute('stop-color') || getComputedStyle(st).stopColor) || { r: 0, g: 0, b: 0 } }));
          if (stops.length < 2) throw new Error('grad');
          const bb = el.getBBox();
          let coords;
          if (/linear/i.test(g.tagName)) { const v = (k, dflt) => { const x = g.getAttribute(k); return x == null ? dflt : parseFloat(x); }; coords = [bb.x + v('x1', 0) * bb.width, bb.y + v('y1', 0) * bb.height, bb.x + v('x2', 1) * bb.width, bb.y + v('y2', 0) * bb.height]; }
          else coords = [bb.x + bb.width / 2, bb.y + bb.height / 2, 0, bb.x + bb.width / 2, bb.y + bb.height / 2, Math.max(bb.width, bb.height) / 2];
          const sh = pdf.shading(/linear/i.test(g.tagName) ? 'linear' : 'radial', coords, stops);
          shadings.push(sh);
          part += `q ${ops}${cs.fillRule === 'evenodd' || el.getAttribute('fill-rule') === 'evenodd' ? 'W*' : 'W'} n /${sh.key} sh Q\n`;
        } else {
          const fc = fillV && fillV !== 'none' ? parseColor(fillV) : null;
          if (fc) part += `${fc.a < 0.999 || op < 0.999 ? `/${pdf.alpha(fc.a * op)} gs ` : ''}${rgb(fc)} rg ${ops}${el.getAttribute('fill-rule') === 'evenodd' ? 'f*' : 'f'}\n`;
        }
        const sc = strokeV && strokeV !== 'none' ? parseColor(strokeV) : null;
        if (sc) {
          const w = parseFloat(el.getAttribute('stroke-width') || cs.strokeWidth) || 1;
          const dash = (el.getAttribute('stroke-dasharray') || '').trim();
          const cap = { butt: 0, round: 1, square: 2 }[el.getAttribute('stroke-linecap') || cs.strokeLinecap] || 0;
          const join = { miter: 0, round: 1, bevel: 2 }[el.getAttribute('stroke-linejoin') || cs.strokeLinejoin] || 0;
          part += `${sc.a < 0.999 || op < 0.999 ? `/${pdf.alpha(sc.a * op)} gs ` : ''}${rgb(sc)} RG ${N(w)} w ${cap} J ${join} j ${dash && dash !== 'none' ? `[${dash.split(/[\s,]+/).map(Number).map(N).join(' ')}] 0 d ` : '[] 0 d '}${ops}S\n`;
        }
        if (part) s += `q ${cm(mm)}${part}Q\n`;
      }
      return s ? `q ${cm(mul(M, local))}${s}Q\n` : '';
    };
    const boxOf = (div, M) => {
      // backgrounds and borders of plain boxes (table cells)
      const cs = getComputedStyle(div); const r = rel(div.getBoundingClientRect()); let s = '';
      const bg = parseColor(cs.backgroundColor);
      if (bg) s += `${bg.a < 0.999 ? `/${pdf.alpha(bg.a)} gs ` : ''}${rgb(bg)} rg ${N(r.x)} ${N(r.y)} ${N(r.w)} ${N(r.h)} re f\n`;
      if (cs.backgroundImage && cs.backgroundImage !== 'none') throw new Error('bgimage');
      for (const [side, x1, y1, x2, y2] of [['Top', r.x, r.y, r.x + r.w, r.y], ['Bottom', r.x, r.y + r.h, r.x + r.w, r.y + r.h], ['Left', r.x, r.y, r.x, r.y + r.h], ['Right', r.x + r.w, r.y, r.x + r.w, r.y + r.h]]) {
        const bw = parseFloat(cs[`border${side}Width`]); const bs = cs[`border${side}Style`]; const bc = parseColor(cs[`border${side}Color`]);
        if (bw > 0 && bs !== 'none' && bc) s += `${rgb(bc)} RG ${N(bw)} w ${N(x1)} ${N(y1)} m ${N(x2)} ${N(y2)} l S\n`;
      }
      return s ? `q ${cm(M)}${s}Q\n` : '';
    };
    const imgOf = async (img, box, M, opacity) => {
      const src = img.currentSrc || img.src; if (!src) return '';
      const im = await pdf.image(src);
      if (!im) throw new Error('img');
      const r = rel(img.getBoundingClientRect()); const b = rel(box.getBoundingClientRect());
      return `q ${cm(M)}${opacity < 0.999 ? `/${pdf.alpha(opacity)} gs ` : ''}${N(b.x)} ${N(b.y)} ${N(b.w)} ${N(b.h)} re W n ${N(r.w)} 0 0 ${N(-r.h)} ${N(r.x)} ${N(r.y + r.h)} cm /${im.key} Do Q\n`;
    };
    // walk the element tree
    const walkEl = async (elNode, Mparent) => {
      const id = elNode.dataset.el; const m = model.get(id) || {};
      const r = rel(elNode.getBoundingClientRect());
      let M = Mparent;
      if (m.rot) M = mul(M, rotAbout(m.rot, r.x + r.w / 2, r.y + r.h / 2));
      if (m.type === 'group') { let s = ''; for (const k of elNode.children) if (k.dataset && k.dataset.el) s += await walkEl(k, M); return s; }
      if (getComputedStyle(elNode).visibility === 'hidden') return '';
      try {
        if (['chart', 'media', 'unknown'].includes(m.type) || elNode.querySelector('video,audio,canvas,foreignObject')) return await rasterEl(elNode, M);
        let s = '';
        if (m.type === 'image') {
          const box = elNode.firstElementChild; const img = elNode.querySelector('img');
          const bcs = box ? getComputedStyle(box) : null;
          if (bcs && ((bcs.clipPath && bcs.clipPath !== 'none') || (bcs.filter && bcs.filter !== 'none') || (bcs.transform && bcs.transform !== 'none') || (bcs.boxShadow && bcs.boxShadow !== 'none'))) return await rasterEl(elNode, M);
          if (img) s += await imgOf(img, box || elNode, M, parseFloat((bcs && bcs.opacity) || '1'));
          return s;
        }
        for (const svg of elNode.querySelectorAll(':scope > div > svg, :scope > svg')) s += svgOf(svg, M);
        for (const td of elNode.querySelectorAll('td')) s += boxOf(td, M);
        s += textOf(elNode, M);
        return s;
      } catch (e) {
        return rasterEl(elNode, M);
      }
    };
    for (const k of stage.children) { if (k.dataset && k.dataset.el) c += await walkEl(k, ID); else if (k.classList && k.classList.contains('of-slides-el')) { try { c += textOf(k, ID); } catch { c += await rasterEl(k, ID); } } }
    out.push({ content: c, w: W, h: H, shadings, extraImages });
  };
  return { addPage, finish: (title = '') => pdf.build(out, title) };
}
export async function vectorPdf(pages, { title = '', raster } = {}) {
  const v = createVectorPdf({ raster });
  for (const p of pages) await v.addPage(p);
  return v.finish(title);
}
