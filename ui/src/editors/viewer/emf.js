/*
 * EMF and WMF pictures → SVG – pure, tested in tests/viewer-emf.test.js. Covers the records that Visio, Office and
 * most programs write for diagrams and clip-art: lines, polygons, Bézier curves, rectangles, ellipses, paths, pens,
 * brushes, fonts, text, mapping modes/world transforms and embedded bitmaps (DIB → BMP data URL). EMF+ comments are
 * skipped (every EMF+ file also carries the GDI records).
 *
 *   isEmf(bytes) · isWmf(bytes)
 *   metafileToSvg(bytes) → { svg, width, height } | null
 */

const S = (v) => (Math.round(v * 100) / 100).toString();
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const rgb = (c) => `#${[c & 255, (c >>> 8) & 255, (c >>> 16) & 255].map((x) => x.toString(16).padStart(2, '0')).join('')}`;

export const isEmf = (b) => b && b.length > 44 && b[0] === 1 && b[1] === 0 && b[2] === 0 && b[3] === 0 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46;
export const isWmf = (b) => b && b.length > 18 && ((b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) || ((b[0] === 1 || b[0] === 2) && b[1] === 0 && b[2] === 9 && b[3] === 0));

function b64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return typeof btoa === 'function' ? btoa(s) : Buffer.from(s, 'binary').toString('base64');
}
/** DIB (BITMAPINFO + bits) → BMP data URL */
function dibUrl(info, bits) {
  const head = new Uint8Array(14); const dv = new DataView(head.buffer);
  head[0] = 0x42; head[1] = 0x4d;
  dv.setUint32(2, 14 + info.length + bits.length, true); dv.setUint32(10, 14 + info.length, true);
  const all = new Uint8Array(14 + info.length + bits.length);
  all.set(head); all.set(info, 14); all.set(bits, 14 + info.length);
  return `data:image/bmp;base64,${b64(all)}`;
}

/** shared drawing state + SVG output */
class Canvas {
  constructor() {
    this.out = []; this.objs = new Map();
    this.st = { pen: { color: '#000', width: 1, style: 0 }, brush: { color: '#fff', style: 0 }, font: { height: 12, weight: 400, italic: false, face: 'Arial', escapement: 0 }, text: '#000', align: 0, bkMode: 2, fillRule: 'evenodd', pos: [0, 0], xf: [1, 0, 0, 1, 0, 0], wo: [0, 0], we: [1, 1], vo: [0, 0], ve: [1, 1], mapMode: 1 };
    this.stack = []; this.path = null;
  }
  // world → device: world transform, then window → viewport (only when the mapping mode scales)
  pt(x, y) {
    const [a, b, c, d, e, f] = this.st.xf;
    let X = a * x + c * y + e; let Y = b * x + d * y + f;
    const s = this.st;
    if (s.mapMode >= 7 && s.we[0] && s.we[1]) { X = (X - s.wo[0]) * (s.ve[0] / s.we[0]) + s.vo[0]; Y = (Y - s.wo[1]) * (s.ve[1] / s.we[1]) + s.vo[1]; } else { X -= s.wo[0] - s.vo[0]; Y -= s.wo[1] - s.vo[1]; }
    return [X, Y];
  }
  scaleLen(v) {
    const [a, b] = this.st.xf; const s = this.st;
    const k = Math.hypot(a, b) * (s.mapMode >= 7 && s.we[0] ? Math.abs(s.ve[0] / s.we[0]) : 1);
    return v * k;
  }
  stroke() { const p = this.st.pen; if (p.style === 5) return 'stroke="none"'; const w = Math.max(this.scaleLen(p.width || 1), 0.5); const dash = p.style === 1 ? ` stroke-dasharray="${S(w * 4)} ${S(w * 2)}"` : p.style === 2 ? ` stroke-dasharray="${S(w)} ${S(w * 2)}"` : ''; return `stroke="${p.color}" stroke-width="${S(w)}"${dash} stroke-linejoin="round"`; }
  fill() { const b = this.st.brush; return b.style === 1 ? 'fill="none"' : `fill="${b.color}"`; }
  d(points, close) { return points.map((p, i) => `${i ? 'L' : 'M'}${S(p[0])} ${S(p[1])}`).join('') + (close ? 'Z' : ''); }
  emit(svgPathD, { fill = true, stroke = true } = {}) {
    if (this.path) { this.path.push(svgPathD); return; }
    this.out.push(`<path d="${svgPathD}" ${fill ? this.fill() : 'fill="none"'} ${stroke ? this.stroke() : 'stroke="none"'} fill-rule="${this.st.fillRule}"/>`);
  }
  poly(pts, close) { if (pts.length) this.emit(this.d(pts.map(([x, y]) => this.pt(x, y)), close), { fill: close }); }
  bezier(pts, fromCurrent) {
    const P = pts.map(([x, y]) => this.pt(x, y));
    let d = ''; let i = 0;
    if (!fromCurrent) { d = `M${S(P[0][0])} ${S(P[0][1])}`; i = 1; } else { const c = this.pt(...this.st.pos); d = `M${S(c[0])} ${S(c[1])}`; }
    for (; i + 2 < P.length; i += 3) d += `C${S(P[i][0])} ${S(P[i][1])} ${S(P[i + 1][0])} ${S(P[i + 1][1])} ${S(P[i + 2][0])} ${S(P[i + 2][1])}`;
    if (pts.length) this.st.pos = pts[pts.length - 1];
    if (this.path) { this.path.push(fromCurrent ? d.replace(/^M[^C]*/, '') : d); return; }
    this.emit(d, { fill: false });
  }
  rect(l, t, r, b, rx = 0, ry = 0) {
    const [x0, y0] = this.pt(l, t); const [x1, y1] = this.pt(r, b);
    const x = Math.min(x0, x1); const y = Math.min(y0, y1);
    this.out.push(`<rect x="${S(x)}" y="${S(y)}" width="${S(Math.abs(x1 - x0))}" height="${S(Math.abs(y1 - y0))}"${rx ? ` rx="${S(this.scaleLen(rx / 2))}" ry="${S(this.scaleLen(ry / 2))}"` : ''} ${this.fill()} ${this.stroke()}/>`);
  }
  ellipse(l, t, r, b) {
    const [x0, y0] = this.pt(l, t); const [x1, y1] = this.pt(r, b);
    this.out.push(`<ellipse cx="${S((x0 + x1) / 2)}" cy="${S((y0 + y1) / 2)}" rx="${S(Math.abs(x1 - x0) / 2)}" ry="${S(Math.abs(y1 - y0) / 2)}" ${this.fill()} ${this.stroke()}/>`);
  }
  text(x, y, str) {
    if (!str) return;
    const f = this.st.font; const [X, Y] = this.pt(x, y);
    const size = Math.max(1, this.scaleLen(Math.abs(f.height) || 12));
    const al = this.st.align;
    const anchor = (al & 6) === 6 ? 'middle' : (al & 2) ? 'end' : 'start';
    const base = (al & 8) ? 'auto' : (al & 24) === 24 ? 'auto' : 'hanging';
    const rot = f.escapement ? ` transform="rotate(${S(-f.escapement / 10)} ${S(X)} ${S(Y)})"` : '';
    this.out.push(`<text x="${S(X)}" y="${S(Y)}" font-family="${esc(f.face)}, Arial, sans-serif" font-size="${S(size)}"${f.weight >= 600 ? ' font-weight="bold"' : ''}${f.italic ? ' font-style="italic"' : ''} fill="${this.st.text}" text-anchor="${anchor}" dominant-baseline="${base}"${rot} xml:space="preserve">${esc(str)}</text>`);
  }
  image(dx, dy, dw, dh, url) {
    const [x0, y0] = this.pt(dx, dy); const [x1, y1] = this.pt(dx + dw, dy + dh);
    const x = Math.min(x0, x1); const y = Math.min(y0, y1); const w = Math.abs(x1 - x0); const h = Math.abs(y1 - y0);
    const flip = (dh < 0) !== (y1 < y0);
    this.out.push(`<image x="${S(x)}" y="${S(y)}" width="${S(w)}" height="${S(h)}" preserveAspectRatio="none" href="${url}"${flip ? ` transform="matrix(1 0 0 -1 0 ${S(2 * y + h)})"` : ''}/>`);
  }
  beginPath() { this.path = []; }
  endPath() { /* kept until fill/stroke */ }
  finishPath(fill, stroke) {
    if (!this.path) return;
    const d = this.path.join('');
    this.path = null;
    if (d) this.out.push(`<path d="${d}" ${fill ? this.fill() : 'fill="none"'} ${stroke ? this.stroke() : 'stroke="none"'} fill-rule="${this.st.fillRule}"/>`);
  }
  save() { this.stack.push(JSON.parse(JSON.stringify(this.st))); }
  restore(n = -1) { let k = n < 0 ? -n : 1; let s = null; while (k-- > 0 && this.stack.length) s = this.stack.pop(); if (s) this.st = s; }
}

/* ---------------- EMF ---------------- */
function readEmf(b) {
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const i32 = (o) => dv.getInt32(o, true); const u32 = (o) => dv.getUint32(o, true);
  const i16 = (o) => dv.getInt16(o, true); const f32 = (o) => dv.getFloat32(o, true);
  const bounds = [i32(8), i32(12), i32(16), i32(20)];
  const cv = new Canvas();
  let off = 0; let guard = 0;
  const pts32 = (o, n) => { const a = []; for (let k = 0; k < n; k++) a.push([i32(o + k * 8), i32(o + k * 8 + 4)]); return a; };
  const pts16 = (o, n) => { const a = []; for (let k = 0; k < n; k++) a.push([i16(o + k * 4), i16(o + k * 4 + 2)]); return a; };
  const utf16 = (o, n) => { let s = ''; for (let k = 0; k < n; k++) s += String.fromCharCode(dv.getUint16(o + k * 2, true)); return s; };
  while (off + 8 <= b.length && guard++ < 500000) {
    const type = u32(off); const size = u32(off + 4);
    if (size < 8 || off + size > b.length) break;
    const r = off + 8; const st = cv.st;
    switch (type) {
      case 14: off = b.length; continue; // EOF
      case 9: st.we = [i32(r), i32(r + 4)]; break; // SETWINDOWEXTEX
      case 10: st.wo = [i32(r), i32(r + 4)]; break;
      case 11: st.ve = [i32(r), i32(r + 4)]; break;
      case 12: st.vo = [i32(r), i32(r + 4)]; break;
      case 17: st.mapMode = u32(r); break;
      case 18: st.bkMode = u32(r); break;
      case 19: st.fillRule = u32(r) === 2 ? 'nonzero' : 'evenodd'; break;
      case 22: st.align = u32(r); break;
      case 24: st.text = rgb(u32(r)); break;
      case 27: st.pos = [i32(r), i32(r + 4)]; if (cv.path) { const p = cv.pt(...st.pos); cv.path.push(`M${S(p[0])} ${S(p[1])}`); } break;
      case 54: { // LINETO
        const to = [i32(r), i32(r + 4)];
        if (cv.path) { const p = cv.pt(...to); cv.path.push(`L${S(p[0])} ${S(p[1])}`); } else cv.poly([st.pos, to], false);
        st.pos = to; break;
      }
      case 33: cv.save(); break;
      case 34: cv.restore(i32(r)); break;
      case 35: st.xf = [f32(r), f32(r + 4), f32(r + 8), f32(r + 12), f32(r + 16), f32(r + 20)]; break;
      case 36: { // MODIFYWORLDTRANSFORM
        const m = [f32(r), f32(r + 4), f32(r + 8), f32(r + 12), f32(r + 16), f32(r + 20)]; const mode = u32(r + 24);
        const mul = (A, B) => [A[0] * B[0] + A[1] * B[2], A[0] * B[1] + A[1] * B[3], A[2] * B[0] + A[3] * B[2], A[2] * B[1] + A[3] * B[3], A[4] * B[0] + A[5] * B[2] + B[4], A[4] * B[1] + A[5] * B[3] + B[5]];
        if (mode === 1) st.xf = [1, 0, 0, 1, 0, 0]; else if (mode === 2) st.xf = mul(m, st.xf); else if (mode === 3) st.xf = mul(st.xf, m); else if (mode === 4) st.xf = m;
        break;
      }
      case 37: { // SELECTOBJECT
        const ix = u32(r);
        if (ix & 0x80000000) {
          const k = ix & 0x7fffffff;
          if (k === 0) st.brush = { color: '#ffffff', style: 0 }; else if (k === 4) st.brush = { color: '#000000', style: 0 }; else if (k === 5) st.brush = { color: '#000', style: 1 }; else if (k === 1 || k === 2 || k === 3) st.brush = { color: ['#c0c0c0', '#808080', '#404040'][k - 1], style: 0 };
          else if (k === 6) st.pen = { color: '#ffffff', width: 1, style: 0 }; else if (k === 7) st.pen = { color: '#000000', width: 1, style: 0 }; else if (k === 8) st.pen = { color: '#000', width: 1, style: 5 };
        } else {
          const o = cv.objs.get(ix);
          if (o && o.kind === 'pen') st.pen = o.v; else if (o && o.kind === 'brush') st.brush = o.v; else if (o && o.kind === 'font') st.font = o.v;
        }
        break;
      }
      case 38: cv.objs.set(u32(r), { kind: 'pen', v: { style: u32(r + 4) & 15, width: i32(r + 8), color: rgb(u32(r + 16)) } }); break; // CREATEPEN
      case 95: cv.objs.set(u32(r), { kind: 'pen', v: { style: u32(r + 20) & 15, width: u32(r + 24), color: rgb(u32(r + 32)) } }); break; // EXTCREATEPEN
      case 39: cv.objs.set(u32(r), { kind: 'brush', v: { style: u32(r + 4), color: rgb(u32(r + 8)) } }); break; // CREATEBRUSHINDIRECT
      case 40: cv.objs.delete(u32(r)); break;
      case 82: { // EXTCREATEFONTINDIRECTW (LOGFONTW at r + 4)
        const lf = r + 4;
        cv.objs.set(u32(r), { kind: 'font', v: { height: i32(lf), escapement: i32(lf + 8), weight: i32(lf + 16), italic: b[lf + 20] !== 0, face: utf16(lf + 28, 32).replace(/\0.*$/, '') || 'Arial' } });
        break;
      }
      case 42: cv.ellipse(i32(r), i32(r + 4), i32(r + 8), i32(r + 12)); break;
      case 43: cv.rect(i32(r), i32(r + 4), i32(r + 8), i32(r + 12)); break;
      case 44: cv.rect(i32(r), i32(r + 4), i32(r + 8), i32(r + 12), i32(r + 16), i32(r + 20)); break;
      case 2: case 85: { const n = u32(r + 16); cv.bezier(type === 2 ? pts32(r + 20, n) : pts16(r + 20, n), false); break; }
      case 5: case 88: { const n = u32(r + 16); cv.bezier(type === 5 ? pts32(r + 20, n) : pts16(r + 20, n), true); break; }
      case 3: case 86: { const n = u32(r + 16); cv.poly(type === 3 ? pts32(r + 20, n) : pts16(r + 20, n), true); break; }
      case 4: case 87: { const n = u32(r + 16); cv.poly(type === 4 ? pts32(r + 20, n) : pts16(r + 20, n), false); break; }
      case 6: case 89: { // POLYLINETO
        const n = u32(r + 16); const P = type === 6 ? pts32(r + 20, n) : pts16(r + 20, n);
        if (cv.path) { for (const p of P) { const q = cv.pt(...p); cv.path.push(`L${S(q[0])} ${S(q[1])}`); } } else cv.poly([st.pos, ...P], false);
        if (P.length) st.pos = P[P.length - 1];
        break;
      }
      case 7: case 8: case 90: case 91: { // POLYPOLYLINE / POLYPOLYGON
        const polys = u32(r + 16); const total = u32(r + 20);
        let po = r + 24 + polys * 4; const wide = type === 7 || type === 8; const close = type === 8 || type === 91;
        let d = '';
        for (let k = 0; k < polys; k++) {
          const n = u32(r + 24 + k * 4);
          const P = (wide ? pts32(po, n) : pts16(po, n)).map(([x, y]) => cv.pt(x, y));
          po += n * (wide ? 8 : 4);
          d += cv.d(P, close);
        }
        if (total) cv.emit(d, { fill: close });
        break;
      }
      case 59: cv.beginPath(); break;
      case 60: cv.endPath(); break;
      case 61: if (cv.path) cv.path.push('Z'); break;
      case 62: cv.finishPath(true, false); break;
      case 63: cv.finishPath(true, true); break;
      case 64: cv.finishPath(false, true); break;
      case 84: case 83: { // EXTTEXTOUTW / A
        const t0 = r + 28; // EMRTEXT
        const x = i32(t0); const y = i32(t0 + 4); const n = u32(t0 + 8); const offStr = u32(t0 + 12);
        const s = type === 84 ? utf16(off + offStr, n) : Array.from(b.subarray(off + offStr, off + offStr + n), (c) => String.fromCharCode(c)).join('');
        cv.text(x, y, s.replace(/\0+$/, ''));
        break;
      }
      case 81: { // STRETCHDIBITS
        const xDest = i32(r + 16); const yDest = i32(r + 20);
        const offBmi = u32(r + 40); const cbBmi = u32(r + 44); const offBits = u32(r + 48); const cbBits = u32(r + 52);
        const cxDest = i32(r + 64); const cyDest = i32(r + 68);
        if (cbBmi && cbBits) cv.image(xDest, yDest, cxDest, cyDest, dibUrl(b.slice(off + offBmi, off + offBmi + cbBmi), b.slice(off + offBits, off + offBits + cbBits)));
        break;
      }
      default: break;
    }
    off += size;
  }
  return { cv, bounds };
}

/* ---------------- WMF ---------------- */
function readWmf(b) {
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const i16 = (o) => dv.getInt16(o, true); const u16 = (o) => dv.getUint16(o, true); const u32 = (o) => dv.getUint32(o, true);
  let off = 0; let bounds = null; let inch = 1440;
  if (u32(0) === 0x9ac6cdd7) { bounds = [i16(6), i16(8), i16(10), i16(12)]; inch = u16(14) || 1440; off = 22; }
  const headerWords = u16(off + 2);
  off += headerWords * 2;
  const cv = new Canvas(); cv.st.mapMode = 8;
  const free = []; let next = 0;
  const addObj = (o) => { const ix = free.length ? free.shift() : next++; cv.objs.set(ix, o); };
  let guard = 0;
  while (off + 6 <= b.length && guard++ < 500000) {
    const size = u32(off) * 2; const fn = u16(off + 4); const r = off + 6;
    if (size < 6 || off + size > b.length) break;
    const st = cv.st;
    const pts = (o, n) => { const a = []; for (let k = 0; k < n; k++) a.push([i16(o + k * 4), i16(o + k * 4 + 2)]); return a; };
    switch (fn) {
      case 0x0000: off = b.length; continue;
      case 0x020b: st.wo = [i16(r + 2), i16(r)]; break;
      case 0x020c: st.we = [i16(r + 2), i16(r)]; if (!st.ve || (st.ve[0] === 1 && st.ve[1] === 1)) st.ve = [i16(r + 2), i16(r)]; break;
      case 0x0209: st.text = rgb(u32(r)); break;
      case 0x012e: st.align = u16(r); break;
      case 0x0106: st.fillRule = u16(r) === 2 ? 'nonzero' : 'evenodd'; break;
      case 0x0214: st.pos = [i16(r + 2), i16(r)]; break;
      case 0x0213: { const to = [i16(r + 2), i16(r)]; cv.poly([st.pos, to], false); st.pos = to; break; }
      case 0x041b: cv.rect(i16(r + 6), i16(r + 4), i16(r + 2), i16(r)); break;
      case 0x061c: cv.rect(i16(r + 10), i16(r + 8), i16(r + 6), i16(r + 4), i16(r + 2), i16(r)); break;
      case 0x0418: cv.ellipse(i16(r + 6), i16(r + 4), i16(r + 2), i16(r)); break;
      case 0x0324: { const n = i16(r); cv.poly(pts(r + 2, n), true); break; }
      case 0x0325: { const n = i16(r); cv.poly(pts(r + 2, n), false); break; }
      case 0x0538: {
        const polys = u16(r); let po = r + 2 + polys * 2; let d = '';
        for (let k = 0; k < polys; k++) { const n = u16(r + 2 + k * 2); d += cv.d(pts(po, n).map(([x, y]) => cv.pt(x, y)), true); po += n * 4; }
        cv.emit(d, { fill: true });
        break;
      }
      case 0x02fa: addObj({ kind: 'pen', v: { style: u16(r) & 15, width: i16(r + 2), color: rgb(u32(r + 6)) } }); break;
      case 0x02fc: addObj({ kind: 'brush', v: { style: u16(r), color: rgb(u32(r + 2)) } }); break;
      case 0x02fb: {
        let face = ''; for (let k = r + 18; k < r + 18 + 32 && b[k]; k++) face += String.fromCharCode(b[k]);
        addObj({ kind: 'font', v: { height: i16(r), escapement: i16(r + 4), weight: i16(r + 8), italic: b[r + 10] !== 0, face: face || 'Arial' } });
        break;
      }
      case 0x00f7: case 0x0142: case 0x01f9: case 0x06ff: addObj({ kind: 'other' }); break; // palette, DIB brush, region …
      case 0x012d: { const o = cv.objs.get(u16(r)); if (o && o.kind === 'pen') st.pen = o.v; else if (o && o.kind === 'brush') st.brush = o.v; else if (o && o.kind === 'font') st.font = o.v; break; }
      case 0x01f0: { const ix = u16(r); cv.objs.delete(ix); free.push(ix); free.sort((x, y) => x - y); break; }
      case 0x001e: cv.save(); break;
      case 0x0127: cv.restore(i16(r)); break;
      case 0x0521: { const n = u16(r); let s = ''; for (let k = 0; k < n; k++) s += String.fromCharCode(b[r + 2 + k]); const po = r + 2 + n + (n & 1); cv.text(i16(po + 2), i16(po), s); break; }
      case 0x0a32: { const y = i16(r); const x = i16(r + 2); const n = i16(r + 4); const opts = u16(r + 6); const so = r + 8 + (opts & 6 ? 8 : 0); let s = ''; for (let k = 0; k < n; k++) s += String.fromCharCode(b[so + k]); cv.text(x, y, s); break; }
      case 0x0f43: { // STRETCHDIB
        const dh = i16(r + 14); const dw = i16(r + 16); const dy = i16(r + 18); const dx = i16(r + 20);
        const dib = b.subarray(r + 22, off + size);
        const hs = dv.getUint32(r + 22, true); const bpp = u16(r + 22 + 14); const used = dv.getUint32(r + 22 + 32, true);
        const pal = bpp <= 8 ? (used || 1 << bpp) * 4 : 0;
        cv.image(dx, dy, dw, dh, dibUrl(dib.slice(0, hs + pal), dib.slice(hs + pal)));
        break;
      }
      case 0x0b41: { // DIBSTRETCHBLT
        const dh = i16(r + 14); const dw = i16(r + 16); const dy = i16(r + 18); const dx = i16(r + 20);
        const dib = b.subarray(r + 22, off + size);
        if (dib.length > 40) { const hs = dv.getUint32(r + 22, true); const bpp = u16(r + 22 + 14); const used = dv.getUint32(r + 22 + 32, true); const pal = bpp <= 8 ? (used || 1 << bpp) * 4 : 0; cv.image(dx, dy, dw, dh, dibUrl(dib.slice(0, hs + pal), dib.slice(hs + pal))); }
        break;
      }
      default: break;
    }
    off += size;
  }
  return { cv, bounds, inch };
}

/** EMF/WMF bytes → { svg, width, height } (width/height in CSS pixels) or null when nothing could be drawn */
export function metafileToSvg(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let cv; let box;
  if (isEmf(b)) { const r = readEmf(b); cv = r.cv; box = r.bounds; } else if (isWmf(b)) {
    const r = readWmf(b); cv = r.cv;
    const s = r.cv.st;
    const p0 = r.bounds ? r.cv.pt(r.bounds[0], r.bounds[1]) : r.cv.pt(s.wo[0], s.wo[1]);
    const p1 = r.bounds ? r.cv.pt(r.bounds[2], r.bounds[3]) : r.cv.pt(s.wo[0] + s.we[0], s.wo[1] + s.we[1]);
    box = [Math.min(p0[0], p1[0]), Math.min(p0[1], p1[1]), Math.max(p0[0], p1[0]), Math.max(p0[1], p1[1])];
  } else return null;
  if (!cv.out.length) return null;
  let [x0, y0, x1, y1] = box;
  if (!(x1 > x0) || !(y1 > y0)) { x0 = 0; y0 = 0; x1 = 100; y1 = 100; }
  const w = x1 - x0 + 1; const h = y1 - y0 + 1;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${S(x0)} ${S(y0)} ${S(w)} ${S(h)}" width="${S(w)}" height="${S(h)}">${cv.out.join('')}</svg>`;
  return { svg, width: w, height: h };
}

/** as a data URL for <img>/<image> */
export function metafileToDataUrl(bytes) {
  const r = metafileToSvg(bytes);
  if (!r) return null;
  return `data:image/svg+xml;base64,${b64(new TextEncoder().encode(r.svg))}`;
}
