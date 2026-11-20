/*
 * EMF / WMF → SVG for display (the file keeps the original metafile). A small renderer for the records Office and
 * most tools write: pens, brushes, fonts, lines, rectangles, ellipses, polygons/polylines/béziers (16/32 bit), paths,
 * text (ExtTextOutW / TextOut), bitmaps (StretchDIBits, BitBlt, StretchBlt → BMP), window/viewport mapping and the
 * world transform. EMF+ records inside comments are skipped (EMF+ dual files carry the plain EMF records too).
 * metaToSvgUrl(dataUrl) → 'data:image/svg+xml…' | null (cached). Pure – runs in Node for tests.
 */

const cache = new Map();
const hex = (c) => `#${[c & 0xff, (c >> 8) & 0xff, (c >> 16) & 0xff].map((v) => v.toString(16).padStart(2, '0')).join('')}`;
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const f2 = (n) => +n.toFixed(2);
function b64(bytes) {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
function fromDataUrl(url) {
  const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(url || '');
  if (!m) return null;
  let bytes;
  if (m[2]) { if (typeof Buffer !== 'undefined') bytes = new Uint8Array(Buffer.from(m[3], 'base64')); else { const s = atob(m[3]); bytes = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i); } }
  else bytes = new TextEncoder().encode(decodeURIComponent(m[3]));
  return { mime: m[1], bytes };
}
/** DIB (BITMAPINFO + bits) → BMP data URL */
function dibToBmp(bytes, offBmi, cbBmi, offBits, cbBits) {
  if (!cbBmi || !cbBits || offBmi + cbBmi > bytes.length || offBits + cbBits > bytes.length) return null;
  const out = new Uint8Array(14 + cbBmi + cbBits);
  const dv = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d; dv.setUint32(2, out.length, true); dv.setUint32(10, 14 + cbBmi, true);
  out.set(bytes.subarray(offBmi, offBmi + cbBmi), 14);
  out.set(bytes.subarray(offBits, offBits + cbBits), 14 + cbBmi);
  return `data:image/bmp;base64,${b64(out)}`;
}

/* ---------------------------------------------------------------- drawing state shared by both formats */
function painter(width, height) {
  const out = [];
  const st = { pen: { color: '#000000', width: 1, none: false }, brush: { color: '#ffffff', none: false }, font: { h: 12, weight: 400, italic: false, face: 'Arial', angle: 0 }, text: '#000000', align: 0, fill: 'evenodd', pos: [0, 0], wo: [0, 0], we: [1, 1], vo: [0, 0], ve: [1, 1], mapped: false, xf: [1, 0, 0, 1, 0, 0] };
  const stack = [];
  const T = (x, y) => {
    const m = st.xf; let X = m[0] * x + m[2] * y + m[4]; let Y = m[1] * x + m[3] * y + m[5];
    if (st.mapped) { X = (X - st.wo[0]) * (st.ve[0] / (st.we[0] || 1)) + st.vo[0]; Y = (Y - st.wo[1]) * (st.ve[1] / (st.we[1] || 1)) + st.vo[1]; }
    return [f2(X), f2(Y)];
  };
  const scale = () => { const sx = st.mapped ? Math.abs(st.ve[0] / (st.we[0] || 1)) : 1; return sx * Math.hypot(st.xf[0], st.xf[1]); };
  const stroke = () => (st.pen.none ? 'stroke="none"' : `stroke="${st.pen.color}" stroke-width="${f2(Math.max(0.5, st.pen.width * scale()))}" stroke-linejoin="round"`);
  const fill = () => (st.brush.none ? 'fill="none"' : `fill="${st.brush.color}"`);
  let path = null;
  return {
    st, T, out,
    save() { stack.push(JSON.stringify(st)); },
    restore() { const s = stack.pop(); if (s) Object.assign(st, JSON.parse(s)); },
    shape(d, closed = true, how = 'both') {
      if (path) { path.push(d); return; }
      const fl = closed && how !== 'stroke' ? fill() : 'fill="none"';
      const sk = how === 'fill' ? 'stroke="none"' : stroke();
      out.push(`<path d="${d}" ${fl} ${sk} fill-rule="${st.fill}"/>`);
    },
    poly(pts, closed) { if (pts.length < 2) return; const d = `M${pts.map((p) => T(p[0], p[1]).join(' ')).join('L')}${closed ? 'Z' : ''}`; this.shape(d, closed); },
    bez(pts, start) { // start point + groups of 3
      const p0 = start || pts.shift(); if (!p0) return;
      let d = `M${T(p0[0], p0[1]).join(' ')}`;
      for (let i = 0; i + 2 < pts.length; i += 3) d += `C${[pts[i], pts[i + 1], pts[i + 2]].map((p) => T(p[0], p[1]).join(' ')).join(' ')}`;
      this.shape(d, false);
    },
    rect(l, t, r, b, rx = 0) { const a = T(l, t); const c = T(r, b); const x = Math.min(a[0], c[0]); const y = Math.min(a[1], c[1]); if (path) { path.push(`M${x} ${y}H${Math.max(a[0], c[0])}V${Math.max(a[1], c[1])}H${x}Z`); return; } out.push(`<rect x="${x}" y="${y}" width="${f2(Math.abs(c[0] - a[0]))}" height="${f2(Math.abs(c[1] - a[1]))}"${rx ? ` rx="${f2(rx * scale() / 2)}"` : ''} ${fill()} ${stroke()}/>`); },
    ellipse(l, t, r, b) { const a = T(l, t); const c = T(r, b); out.push(`<ellipse cx="${f2((a[0] + c[0]) / 2)}" cy="${f2((a[1] + c[1]) / 2)}" rx="${f2(Math.abs(c[0] - a[0]) / 2)}" ry="${f2(Math.abs(c[1] - a[1]) / 2)}" ${fill()} ${stroke()}/>`); },
    moveTo(x, y) { st.pos = [x, y]; if (path) path.push(`M${T(x, y).join(' ')}`); },
    lineTo(x, y) { if (path) path.push(`L${T(x, y).join(' ')}`); else out.push(`<path d="M${T(st.pos[0], st.pos[1]).join(' ')}L${T(x, y).join(' ')}" fill="none" ${stroke()}/>`); st.pos = [x, y]; },
    beginPath() { path = []; }, endPath() { this.pathD = path ? path.join('') : ''; path = null; },
    closeFigure() { if (path) path.push('Z'); },
    drawPath(how) { const d = this.pathD || ''; this.pathD = ''; if (d) this.shape(d, true, how); },
    text(x, y, s) {
      if (!s) return;
      const p = T(x, y); const sz = Math.abs(st.font.h) * scale() || 12;
      const anchor = (st.align & 6) === 6 ? 'middle' : (st.align & 2) ? 'end' : 'start';
      const base = (st.align & 24) === 24 ? 'auto' : (st.align & 8) ? 'auto' : 'hanging';
      const rot = st.font.angle ? ` transform="rotate(${-st.font.angle / 10} ${p[0]} ${p[1]})"` : '';
      out.push(`<text x="${p[0]}" y="${p[1]}" font-family="${esc(st.font.face)}, sans-serif" font-size="${f2(sz)}" font-weight="${st.font.weight}"${st.font.italic ? ' font-style="italic"' : ''} fill="${st.text}" text-anchor="${anchor}" dominant-baseline="${base}"${rot} xml:space="preserve">${esc(s)}</text>`);
    },
    image(l, t, w, h, href) { if (!href) return; const a = T(l, t); const c = T(l + w, t + h); out.push(`<image x="${Math.min(a[0], c[0])}" y="${Math.min(a[1], c[1])}" width="${f2(Math.abs(c[0] - a[0]))}" height="${f2(Math.abs(c[1] - a[1]))}" preserveAspectRatio="none" href="${href}"/>`); },
    svg(vb) { return `<svg xmlns="http://www.w3.org/2000/svg" width="${f2(width)}" height="${f2(height)}" viewBox="${vb.map(f2).join(' ')}" overflow="hidden">${out.join('')}</svg>`; },
  };
}

/* ---------------------------------------------------------------- EMF */
export function emfToSvg(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 88 || dv.getUint32(0, true) !== 1 || dv.getUint32(40, true) !== 0x464d4520) return null;
  const i32 = (o) => dv.getInt32(o, true); const u32 = (o) => dv.getUint32(o, true); const i16 = (o) => dv.getInt16(o, true); const fl = (o) => dv.getFloat32(o, true);
  const bounds = [i32(8), i32(12), i32(16), i32(20)];
  const frame = [i32(24), i32(28), i32(32), i32(36)]; // 0.01 mm
  const devPx = [i32(72), i32(76)]; const devMm = [i32(80), i32(84)];
  // logical frame in device units (the visible area) – fall back to the bounds
  const pxPerMm = [devMm[0] ? devPx[0] / devMm[0] : 3.78, devMm[1] ? devPx[1] / devMm[1] : 3.78];
  let vb = [frame[0] / 100 * pxPerMm[0], frame[1] / 100 * pxPerMm[1], (frame[2] - frame[0]) / 100 * pxPerMm[0], (frame[3] - frame[1]) / 100 * pxPerMm[1]];
  if (!(vb[2] > 0 && vb[3] > 0)) vb = [bounds[0], bounds[1], Math.max(1, bounds[2] - bounds[0]), Math.max(1, bounds[3] - bounds[1])];
  const P = painter(Math.max(1, (frame[2] - frame[0]) / 100 * 3.7795), Math.max(1, (frame[3] - frame[1]) / 100 * 3.7795));
  const S = P.st;
  const objs = {};
  const stock = (i) => { const k = i & 0x7fffffff; if (k === 5 || k === 8) S.pen = { ...S.pen, none: k === 8 }; if (k === 5) S.brush = { color: '#ffffff', none: true }; if (k <= 4) S.brush = { color: ['#ffffff', '#c0c0c0', '#808080', '#404040', '#000000'][k], none: false }; if (k === 6) S.pen = { color: '#ffffff', width: 1, none: false }; if (k === 7) S.pen = { color: '#000000', width: 1, none: false }; };
  let off = 0; let n = 0;
  const pts16 = (o, c) => { const a = []; for (let k = 0; k < c; k++) a.push([i16(o + k * 4), i16(o + k * 4 + 2)]); return a; };
  const pts32 = (o, c) => { const a = []; for (let k = 0; k < c; k++) a.push([i32(o + k * 8), i32(o + k * 8 + 4)]); return a; };
  while (off + 8 <= bytes.length && n++ < 200000) {
    const type = u32(off); const size = u32(off + 4);
    if (size < 8 || off + size > bytes.length) break;
    const o = off + 8;
    try {
      switch (type) {
        case 14: off = bytes.length; continue; // EOF
        case 9: S.we = [i32(o), i32(o + 4)]; S.mapped = true; break; // SETWINDOWEXTEX
        case 10: S.wo = [i32(o), i32(o + 4)]; break;
        case 11: S.ve = [i32(o), i32(o + 4)]; S.mapped = true; break;
        case 12: S.vo = [i32(o), i32(o + 4)]; break;
        case 17: { const mm = u32(o); if (mm === 1) S.mapped = false; break; } // SETMAPMODE
        case 19: S.fill = u32(o) === 2 ? 'nonzero' : 'evenodd'; break;
        case 22: S.align = u32(o); break;
        case 24: S.text = hex(u32(o)); break;
        case 33: P.save(); break;
        case 34: P.restore(); break;
        case 35: S.xf = [fl(o), fl(o + 4), fl(o + 8), fl(o + 12), fl(o + 16), fl(o + 20)]; break;
        case 36: { const m = [fl(o), fl(o + 4), fl(o + 8), fl(o + 12), fl(o + 16), fl(o + 20)]; const mode = u32(o + 24); const a = S.xf;
          if (mode === 1) S.xf = [1, 0, 0, 1, 0, 0];
          else if (mode === 2) S.xf = [m[0] * a[0] + m[1] * a[2], m[0] * a[1] + m[1] * a[3], m[2] * a[0] + m[3] * a[2], m[2] * a[1] + m[3] * a[3], m[4] * a[0] + m[5] * a[2] + a[4], m[4] * a[1] + m[5] * a[3] + a[5]];
          else if (mode === 3) S.xf = [a[0] * m[0] + a[1] * m[2], a[0] * m[1] + a[1] * m[3], a[2] * m[0] + a[3] * m[2], a[2] * m[1] + a[3] * m[3], a[4] * m[0] + a[5] * m[2] + m[4], a[4] * m[1] + a[5] * m[3] + m[5]];
          else if (mode === 4) S.xf = m;
          break; }
        case 37: { const i = u32(o); if (i & 0x80000000) stock(i); else { const ob = objs[i]; if (ob) { if (ob.k === 'pen') S.pen = ob.v; else if (ob.k === 'brush') S.brush = ob.v; else if (ob.k === 'font') S.font = ob.v; } } break; }
        case 38: objs[u32(o)] = { k: 'pen', v: { color: hex(u32(o + 16)), width: Math.max(1, i32(o + 8)), none: (u32(o + 4) & 0xf) === 5 } }; break; // CREATEPEN
        case 95: objs[u32(o)] = { k: 'pen', v: { color: hex(u32(o + 28)), width: Math.max(1, u32(o + 24)), none: (u32(o + 20) & 0xf) === 5 } }; break; // EXTCREATEPEN
        case 39: objs[u32(o)] = { k: 'brush', v: { color: hex(u32(o + 8)), none: u32(o + 4) === 1 } }; break;
        case 82: { // EXTCREATEFONTINDIRECTW
          let face = ''; for (let k = 0; k < 32; k++) { const c = dv.getUint16(o + 4 + 28 + k * 2, true); if (!c) break; face += String.fromCharCode(c); }
          objs[u32(o)] = { k: 'font', v: { h: i32(o + 4), angle: i32(o + 12), weight: i32(o + 20) || 400, italic: !!bytes[o + 24], face: face || 'Arial' } };
          break; }
        case 40: delete objs[u32(o)]; break;
        case 27: P.moveTo(i32(o), i32(o + 4)); break;
        case 54: P.lineTo(i32(o), i32(o + 4)); break;
        case 43: P.rect(i32(o), i32(o + 4), i32(o + 8), i32(o + 12)); break;
        case 44: P.rect(i32(o), i32(o + 4), i32(o + 8), i32(o + 12), i32(o + 16)); break;
        case 42: P.ellipse(i32(o), i32(o + 4), i32(o + 8), i32(o + 12)); break;
        case 86: case 87: P.poly(pts16(o + 20, u32(o + 16)), type === 86); break; // POLYGON16 / POLYLINE16
        case 3: case 4: P.poly(pts32(o + 20, u32(o + 16)), type === 3); break;
        case 85: case 2: P.bez(type === 85 ? pts16(o + 20, u32(o + 16)) : pts32(o + 20, u32(o + 16))); break; // POLYBEZIER
        case 88: case 5: { const pts = type === 88 ? pts16(o + 20, u32(o + 16)) : pts32(o + 20, u32(o + 16)); P.bez(pts, S.pos); if (pts.length) S.pos = pts[pts.length - 1]; break; } // POLYBEZIERTO
        case 89: case 6: { const pts = type === 89 ? pts16(o + 20, u32(o + 16)) : pts32(o + 20, u32(o + 16)); for (const p of pts) P.lineTo(p[0], p[1]); break; } // POLYLINETO
        case 91: case 90: case 8: case 7: { // POLYPOLYGON16 / POLYPOLYLINE16 / 32
          const nPolys = u32(o + 16); const is16 = type === 91 || type === 90; const counts = []; for (let k = 0; k < nPolys; k++) counts.push(u32(o + 24 + k * 4));
          let po = o + 24 + nPolys * 4; let d = '';
          for (const c of counts) { const pts = is16 ? pts16(po, c) : pts32(po, c); po += c * (is16 ? 4 : 8); if (pts.length) d += `M${pts.map((p) => P.T(p[0], p[1]).join(' ')).join('L')}${type === 91 || type === 8 ? 'Z' : ''}`; }
          if (d) P.shape(d, type === 91 || type === 8);
          break; }
        case 59: P.beginPath(); break;
        case 60: P.endPath(); break;
        case 61: P.closeFigure(); break;
        case 62: P.drawPath('fill'); break;
        case 63: P.drawPath('both'); break;
        case 64: P.drawPath('stroke'); break;
        case 84: case 83: { // EXTTEXTOUTW / A
          const r = o + 28; const x = i32(r); const y = i32(r + 4); const nc = u32(r + 8); const os = u32(r + 12);
          let s = '';
          if (type === 84) for (let k = 0; k < nc && off + os + k * 2 + 1 < bytes.length; k++) s += String.fromCharCode(dv.getUint16(off + os + k * 2, true));
          else for (let k = 0; k < nc; k++) s += String.fromCharCode(bytes[off + os + k]);
          P.text(x, y, s.replace(/\0/g, ''));
          break; }
        case 81: { // STRETCHDIBITS
          const href = dibToBmp(bytes, off + u32(o + 40), u32(o + 44), off + u32(o + 48), u32(o + 52));
          P.image(i32(o + 16), i32(o + 20), i32(o + 64), i32(o + 68), href); break; }
        case 76: case 77: { // BITBLT / STRETCHBLT
          const x = i32(o + 16); const y = i32(o + 20); const w = i32(o + 24); const hh = i32(o + 28);
          const offBmi = u32(o + 76); const cbBmi = u32(o + 80); const offBits = u32(o + 84); const cbBits = u32(o + 88);
          if (cbBmi) P.image(x, y, w, hh, dibToBmp(bytes, off + offBmi, cbBmi, off + offBits, cbBits));
          else if (!(u32(o + 32) === 0x00aa0029)) { const c = hex(u32(o + 68)); P.out.push(`<rect x="${P.T(x, y)[0]}" y="${P.T(x, y)[1]}" width="${f2(w)}" height="${f2(hh)}" fill="${c}"/>`); }
          break; }
        default: break;
      }
    } catch { /* skip a broken record */ }
    off += size;
  }
  return P.svg(vb);
}

/* ---------------------------------------------------------------- WMF */
export function wmfToSvg(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let off = 0; let vb = null; let inch = 1440;
  if (bytes.length > 22 && dv.getUint32(0, true) === 0x9ac6cdd7) {
    const l = dv.getInt16(6, true); const t = dv.getInt16(8, true); const r = dv.getInt16(10, true); const b = dv.getInt16(12, true);
    inch = dv.getUint16(14, true) || 1440; vb = [l, t, r - l, b - t]; off = 22;
  }
  if (off + 18 > bytes.length) return null;
  const hdr = dv.getUint16(off + 2, true);
  if (hdr !== 9) return null;
  off += hdr * 2;
  const W = vb ? (vb[2] / inch) * 96 : 400; const H = vb ? (vb[3] / inch) * 96 : 300;
  const P = painter(Math.abs(W) || 400, Math.abs(H) || 300);
  const S = P.st;
  const objs = [];
  const add = (o) => { let i = objs.findIndex((x) => x === null); if (i < 0) i = objs.length; objs[i] = o; };
  const i16 = (o) => dv.getInt16(o, true); const u16 = (o) => dv.getUint16(o, true); const u32 = (o) => dv.getUint32(o, true);
  const col = (o) => hex(u16(o) | (u16(o + 2) << 16));
  let n = 0;
  while (off + 6 <= bytes.length && n++ < 200000) {
    const size = u32(off) * 2; const fn = u16(off + 4); const o = off + 6;
    if (size < 6 || off + size > bytes.length) break;
    try {
      switch (fn) {
        case 0x0000: off = bytes.length; continue;
        case 0x020b: S.wo = [i16(o + 2), i16(o)]; if (!vb) vb = [S.wo[0], S.wo[1], 1, 1]; break;
        case 0x020c: S.we = [i16(o + 2), i16(o)]; if (!vb || vb[2] === 1) vb = [S.wo[0], S.wo[1], S.we[0], S.we[1]]; break;
        case 0x0106: S.fill = u16(o) === 2 ? 'nonzero' : 'evenodd'; break;
        case 0x0209: S.text = col(o); break;
        case 0x012e: S.align = u16(o); break;
        case 0x001e: P.save(); break;
        case 0x0127: P.restore(); break;
        case 0x02fa: add({ k: 'pen', v: { color: col(o + 6), width: Math.max(1, i16(o + 2)), none: (u16(o) & 0xf) === 5 } }); break;
        case 0x02fc: add({ k: 'brush', v: { color: col(o + 2), none: u16(o) === 1 } }); break;
        case 0x02fb: { let face = ''; for (let k = 0; k < 32; k++) { const c = bytes[o + 18 + k]; if (!c) break; face += String.fromCharCode(c); } add({ k: 'font', v: { h: i16(o), angle: i16(o + 4), weight: i16(o + 8) || 400, italic: !!bytes[o + 10], face: face || 'Arial' } }); break; }
        case 0x01f9: case 0x00f7: case 0x06ff: case 0x0142: add({ k: 'other' }); break; // palette, region, pattern brush …
        case 0x012d: { const ob = objs[u16(o)]; if (ob) { if (ob.k === 'pen') S.pen = ob.v; else if (ob.k === 'brush') S.brush = ob.v; else if (ob.k === 'font') S.font = ob.v; } break; }
        case 0x01f0: objs[u16(o)] = null; break;
        case 0x0214: P.moveTo(i16(o + 2), i16(o)); break;
        case 0x0213: P.lineTo(i16(o + 2), i16(o)); break;
        case 0x041b: P.rect(i16(o + 6), i16(o + 4), i16(o + 2), i16(o)); break;
        case 0x061c: P.rect(i16(o + 10), i16(o + 8), i16(o + 6), i16(o + 4), i16(o + 2)); break;
        case 0x0418: P.ellipse(i16(o + 6), i16(o + 4), i16(o + 2), i16(o)); break;
        case 0x0324: case 0x0325: { const c = u16(o); const pts = []; for (let k = 0; k < c; k++) pts.push([i16(o + 2 + k * 4), i16(o + 4 + k * 4)]); P.poly(pts, fn === 0x0324); break; }
        case 0x0538: { const np = u16(o); const counts = []; for (let k = 0; k < np; k++) counts.push(u16(o + 2 + k * 2)); let po = o + 2 + np * 2; let d = '';
          for (const c of counts) { const pts = []; for (let k = 0; k < c; k++) pts.push(P.T(i16(po + k * 4), i16(po + k * 4 + 2))); po += c * 4; if (pts.length) d += `M${pts.map((p) => p.join(' ')).join('L')}Z`; }
          if (d) P.shape(d, true); break; }
        case 0x0521: { const len = u16(o); let s = ''; for (let k = 0; k < len; k++) s += String.fromCharCode(bytes[o + 2 + k]); const po = o + 2 + len + (len & 1); P.text(i16(po + 2), i16(po), s); break; }
        case 0x0a32: { const y = i16(o); const x = i16(o + 2); const len = u16(o + 4); const opts = u16(o + 6); const so = o + 8 + (opts & 6 ? 8 : 0); let s = ''; for (let k = 0; k < len; k++) s += String.fromCharCode(bytes[so + k]); P.text(x, y, s); break; }
        case 0x0f43: { // STRETCHDIB: rop(4) usage(2) srcH srcW srcY srcX destH destW destY destX, DIB
          const dh = i16(o + 14); const dw = i16(o + 16); const dy = i16(o + 18); const dx = i16(o + 20);
          const dib = o + 22; const hs = u32(dib); const bc = u16(dib + 14); const clr = u32(dib + 32);
          const nCol = bc <= 8 ? (clr || (1 << bc)) : (u32(dib + 16) === 3 ? 3 : 0);
          const cbBmi = hs + nCol * 4;
          P.image(dx, dy, dw, dh, dibToBmp(bytes, dib, cbBmi, dib + cbBmi, off + size - (dib + cbBmi))); break; }
        case 0x0b41: { // DIBSTRETCHBLT
          const dh = i16(o + 12); const dw = i16(o + 14); const dy = i16(o + 16); const dx = i16(o + 18);
          const dib = o + 20; if (dib + 40 <= off + size) { const hs = u32(dib); const bc = u16(dib + 14); const clr = u32(dib + 32); const nCol = bc <= 8 ? (clr || (1 << bc)) : 0; const cbBmi = hs + nCol * 4; P.image(dx, dy, dw, dh, dibToBmp(bytes, dib, cbBmi, dib + cbBmi, off + size - (dib + cbBmi))); }
          break; }
        default: break;
      }
    } catch { /* skip */ }
    off += size;
  }
  if (!vb) vb = [0, 0, 1000, 1000];
  S.mapped = false;
  return P.svg([Math.min(vb[0], vb[0] + vb[2]), Math.min(vb[1], vb[1] + vb[3]), Math.abs(vb[2]) || 1, Math.abs(vb[3]) || 1]);
}

/** EMF/WMF data URL → SVG data URL (null when it is not a metafile we can read) */
export function metaToSvgUrl(src) {
  if (!src || typeof src !== 'string' || !/^data:image\/(x-emf|emf|x-wmf|wmf|x-msmetafile)/i.test(src)) return null;
  if (cache.has(src)) return cache.get(src);
  let url = null;
  try {
    const d = fromDataUrl(src);
    let bytes = d && d.bytes;
    if (bytes) {
      const svg = /wmf|msmetafile/i.test(d.mime) && !(bytes.length > 44 && bytes[0] === 1 && bytes[40] === 0x20) ? wmfToSvg(bytes) : emfToSvg(bytes) || wmfToSvg(bytes);
      if (svg) url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
      bytes = null;
    }
  } catch { url = null; }
  if (cache.size > 200) cache.delete(cache.keys().next().value);
  cache.set(src, url);
  return url;
}
export const isMetafile = (src) => typeof src === 'string' && /^data:image\/(x-emf|emf|x-wmf|wmf|x-msmetafile)/i.test(src);
