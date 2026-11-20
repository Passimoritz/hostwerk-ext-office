/*
 * Visio (vsdx / vsdm / vstx) → SVG, pure (no DOM): pages, masters (stencil inheritance incl. group sub-shapes),
 * style sheets, geometry sections (MoveTo/LineTo/ArcTo/EllipticalArcTo/Ellipse, the Rel… rows, NURBSTo, PolylineTo, splines),
 * fills with transparency, line weight/colour/pattern, arrows, text blocks (size, colour, bold/italic, alignment,
 * wrapping, rotation), images (ForeignData) and background pages.
 *
 *   readVsdx(files)               files = { 'visio/pages/pages.xml': '<…>', …, media: { 'visio/media/image1.png': url } }
 *                                 → { pages: [{ name, w, h, svg, texts: [] }] }   (w, h in inches)
 *   geometryToPath(rows, w, h, tf) → SVG path data of one geometry section (tf maps local inches → output coords)
 */
import { parseXml, kids, kid } from '../_shared/xmltree.js';

const S = 72; // output units per inch
const num = (v, d = 0) => { const n = parseFloat(v); return Number.isFinite(n) ? n : d; };

/* ---------------- cells, sections, inheritance ---------------- */

/** <Shape> element → { cells: {N: V}, sections: { Geometry: [{ IX, cells, rows: [{ T, IX, cells, del }] }], Character: … }, text } */
function readSheet(el) {
  const cells = {}; const sections = {};
  for (const c of kids(el)) {
    if (c.n === 'Cell') cells[c.a.N] = c.a.V !== undefined ? c.a.V : c.a.F;
    else if (c.n === 'Section') {
      const name = c.a.N;
      const sec = { IX: c.a.IX != null ? Number(c.a.IX) : 0, del: c.a.Del === '1', cells: {}, rows: [] };
      for (const r of kids(c)) {
        if (r.n === 'Cell') sec.cells[r.a.N] = r.a.V !== undefined ? r.a.V : r.a.F;
        else if (r.n === 'Row') {
          const row = { T: r.a.T || '', N: r.a.N || '', IX: r.a.IX != null ? Number(r.a.IX) : sec.rows.length, del: r.a.Del === '1', cells: {} };
          for (const rc of kids(r, 'Cell')) row.cells[rc.a.N] = rc.a.V !== undefined ? rc.a.V : rc.a.F;
          sec.rows.push(row);
        }
      }
      (sections[name] = sections[name] || []).push(sec);
    }
  }
  return { cells, sections };
}

function mergeRows(base, over) {
  const map = new Map(base.map((r) => [r.IX, r]));
  for (const r of over) {
    if (r.del) { map.delete(r.IX); continue; }
    const b = map.get(r.IX);
    map.set(r.IX, b ? { ...b, T: r.T || b.T, cells: { ...b.cells, ...r.cells } } : r);
  }
  return [...map.values()].sort((a, b) => a.IX - b.IX);
}
function mergeSections(base = [], over = []) {
  const map = new Map(base.map((s) => [s.IX, s]));
  for (const s of over) {
    if (s.del) { map.delete(s.IX); continue; }
    const b = map.get(s.IX);
    map.set(s.IX, b ? { ...s, cells: { ...b.cells, ...s.cells }, rows: mergeRows(b.rows, s.rows) } : s);
  }
  return [...map.values()].sort((a, b) => a.IX - b.IX);
}
function mergeSheet(base, over) {
  if (!base) return over;
  const sections = { ...base.sections };
  for (const [k, v] of Object.entries(over.sections)) sections[k] = k === 'Geometry' || k === 'Character' || k === 'Paragraph' ? mergeSections(base.sections[k], v) : v;
  return { cells: { ...base.cells, ...over.cells }, sections };
}

/* ---------------- colours and patterns ---------------- */

const PALETTE = ['#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#e6e6e6', '#cdcdcd', '#b3b3b3', '#9a9a9a', '#808080', '#666666', '#4d4d4d', '#333333', '#1a1a1a'];
export function color(v, fallback) {
  if (v == null || v === '') return fallback;
  const s = String(v).trim();
  if (/^#[0-9a-f]{6}$/i.test(s)) return s.toLowerCase();
  const rgbm = /^RGB\((\d+),\s*(\d+),\s*(\d+)\)$/i.exec(s);
  if (rgbm) return `#${[rgbm[1], rgbm[2], rgbm[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('')}`;
  if (/^\d+$/.test(s) && PALETTE[Number(s)]) return PALETTE[Number(s)];
  return fallback;
}
const DASH = { 2: [6, 3], 3: [1.5, 3], 4: [6, 3, 1.5, 3], 5: [6, 3, 1.5, 3, 1.5, 3], 6: [9, 3, 6, 3], 7: [12, 3], 9: [4, 4], 10: [1.5, 1.5], 16: [12, 6] };

/* ---------------- geometry ---------------- */

const fmt = (n) => (Math.round(n * 100) / 100).toString();

/** sample an elliptical arc through p0 → (via control point c) → p1 (Visio EllipticalArcTo) into points */
function ellipticalArcPoints(x0, y0, x, y, cx, cy, angle, ratio) {
  // transform into a space where the ellipse is a circle: rotate by -angle, scale x by 1/ratio
  const ca = Math.cos(-angle); const sa = Math.sin(-angle);
  const tr = (px, py) => { const rx = px * ca - py * sa; const ry = px * sa + py * ca; return [rx / (ratio || 1), ry]; };
  const inv = (px, py) => { const sx = px * (ratio || 1); const c2 = Math.cos(angle); const s2 = Math.sin(angle); return [sx * c2 - py * s2, sx * s2 + py * c2]; };
  const [ax, ay] = tr(x0, y0); const [bx, by] = tr(cx, cy); const [ex, ey] = tr(x, y);
  // circle through three points
  const d = 2 * (ax * (by - ey) + bx * (ey - ay) + ex * (ay - by));
  if (Math.abs(d) < 1e-12) return [[x, y]];
  const ux = ((ax * ax + ay * ay) * (by - ey) + (bx * bx + by * by) * (ey - ay) + (ex * ex + ey * ey) * (ay - by)) / d;
  const uy = ((ax * ax + ay * ay) * (ex - bx) + (bx * bx + by * by) * (ax - ex) + (ex * ex + ey * ey) * (bx - ax)) / d;
  const r = Math.hypot(ax - ux, ay - uy);
  let a0 = Math.atan2(ay - uy, ax - ux); let a1 = Math.atan2(by - uy, bx - ux); let a2 = Math.atan2(ey - uy, ex - ux);
  // go from a0 to a2 passing a1
  const norm = (a) => { while (a < 0) a += 2 * Math.PI; while (a >= 2 * Math.PI) a -= 2 * Math.PI; return a; };
  const ccw = norm(a1 - a0) < norm(a2 - a0);
  let sweep = ccw ? norm(a2 - a0) : -norm(a0 - a2);
  if (sweep === 0) sweep = 2 * Math.PI;
  const n = Math.max(8, Math.ceil(Math.abs(sweep) / (Math.PI / 24)));
  const pts = [];
  for (let i = 1; i <= n; i++) { const a = a0 + (sweep * i) / n; pts.push(inv(ux + r * Math.cos(a), uy + r * Math.sin(a))); }
  void a1; void a2;
  return pts;
}

/** numbers inside "NURBS(…)" / "POLYLINE(…)" */
const fnArgs = (s) => { const m = /\(([^)]*)\)/.exec(String(s || '')); return m ? m[1].split(',').map((x) => num(x)) : []; };

/**
 * One geometry section → path data. rows: [{ T, cells }], w/h: shape size (for Rel* rows and NURBS/POLYLINE type 0),
 * tf(x, y) → [X, Y] output coordinates. Returns { d, closed }.
 */
export function geometryToPath(rows, w, h, tf = (x, y) => [x, y]) {
  let d = ''; let cx = 0; let cy = 0; let start = null; let closed = false;
  const P = (x, y) => { const [X, Y] = tf(x, y); return `${fmt(X)} ${fmt(Y)}`; };
  const move = (x, y) => { d += `M${P(x, y)}`; cx = x; cy = y; start = [x, y]; };
  const line = (x, y) => { if (start == null) { move(x, y); return; } d += `L${P(x, y)}`; cx = x; cy = y; };
  for (const r of rows) {
    const c = r.cells; const X = num(c.X); const Y = num(c.Y); const A = num(c.A); const B = num(c.B); const C = num(c.C); const D = num(c.D);
    switch (r.T) {
      case 'MoveTo': move(X, Y); break;
      case 'RelMoveTo': move(X * w, Y * h); break;
      case 'LineTo': line(X, Y); break;
      case 'RelLineTo': line(X * w, Y * h); break;
      case 'ArcTo': {
        // bow A: distance from the chord midpoint to the arc; positive = arc to the left when walking cx,cy → X,Y? (Visio: counter-clockwise)
        if (Math.abs(A) < 1e-9) { line(X, Y); break; }
        const mx = (cx + X) / 2; const my = (cy + Y) / 2; const dx = X - cx; const dy = Y - cy; const len = Math.hypot(dx, dy) || 1;
        // control point on the arc: midpoint moved by A along the chord's normal (to the right of the direction for positive A)
        const nx = dy / len; const ny = -dx / len;
        const px = mx + nx * A; const py = my + ny * A;
        for (const [x, y] of ellipticalArcPoints(cx, cy, X, Y, px, py, 0, 1)) line(x, y);
        cx = X; cy = Y;
        break;
      }
      case 'EllipticalArcTo': case 'RelEllipticalArcTo': {
        const k = r.T === 'RelEllipticalArcTo';
        const ex = k ? X * w : X; const ey = k ? Y * h : Y; const ax = k ? A * w : A; const ay = k ? B * h : B;
        for (const [x, y] of ellipticalArcPoints(cx, cy, ex, ey, ax, ay, C, D)) line(x, y);
        cx = ex; cy = ey;
        break;
      }
      case 'Ellipse': {
        // centre X,Y; (A,B) end of one axis; (C,D) end of the other
        const ux = A - X; const uy = B - Y; const vx = C - X; const vy = D - Y;
        const n = 48;
        for (let i = 0; i <= n; i++) {
          const t = (2 * Math.PI * i) / n;
          const x = X + ux * Math.cos(t) + vx * Math.sin(t); const y = Y + uy * Math.cos(t) + vy * Math.sin(t);
          if (i === 0) move(x, y); else line(x, y);
        }
        d += 'Z'; closed = true;
        break;
      }
      case 'RelCubBezTo': {
        const p1 = tf(A * w, B * h); const p2 = tf(C * w, D * h); const p3 = tf(X * w, Y * h);
        if (start == null) move(cx, cy);
        d += `C${fmt(p1[0])} ${fmt(p1[1])} ${fmt(p2[0])} ${fmt(p2[1])} ${fmt(p3[0])} ${fmt(p3[1])}`; cx = X * w; cy = Y * h;
        break;
      }
      case 'RelQuadBezTo': {
        const p1 = tf(A * w, B * h); const p2 = tf(X * w, Y * h);
        if (start == null) move(cx, cy);
        d += `Q${fmt(p1[0])} ${fmt(p1[1])} ${fmt(p2[0])} ${fmt(p2[1])}`; cx = X * w; cy = Y * h;
        break;
      }
      case 'NURBSTo': {
        // E = NURBS(knotLast, degree, xType, yType, x1, y1, knot1, weight1, …) – approximated through the control points
        const a = fnArgs(c.E);
        const xType = a[2]; const yType = a[3];
        for (let i = 4; i + 3 < a.length; i += 4) line(xType === 0 ? a[i] * w : a[i], yType === 0 ? a[i + 1] * h : a[i + 1]);
        // A/B: second-to-last control point
        if (c.A != null && c.B != null && a.length < 8) line(A, B);
        line(X, Y);
        break;
      }
      case 'PolylineTo': {
        const a = fnArgs(c.A);
        const xType = a[0]; const yType = a[1];
        for (let i = 2; i + 1 < a.length; i += 2) line(xType === 0 ? a[i] * w : a[i], yType === 0 ? a[i + 1] * h : a[i + 1]);
        line(X, Y);
        break;
      }
      case 'SplineStart': case 'SplineKnot': line(X, Y); break;
      case 'InfiniteLine': break;
      default: if (c.X != null && c.Y != null) line(X, Y);
    }
  }
  if (start && Math.abs(cx - start[0]) < 1e-6 && Math.abs(cy - start[1]) < 1e-6 && d && !closed) { d += 'Z'; closed = true; }
  return { d, closed };
}

/* ---------------- transforms ---------------- */

/** local → parent matrix [a, b, c, d, e, f] (x' = a x + c y + e, y' = b x + d y + f) */
function shapeMatrix(cells) {
  const pinX = num(cells.PinX); const pinY = num(cells.PinY);
  const lx = num(cells.LocPinX, num(cells.Width) / 2); const ly = num(cells.LocPinY, num(cells.Height) / 2);
  const ang = num(cells.Angle); const fx = cells.FlipX === '1' ? -1 : 1; const fy = cells.FlipY === '1' ? -1 : 1;
  const co = Math.cos(ang); const si = Math.sin(ang);
  // R * F * (p - L) + Pin
  const a = co * fx; const b = si * fx; const c = -si * fy; const d = co * fy;
  return [a, b, c, d, pinX - (a * lx + c * ly), pinY - (b * lx + d * ly)];
}
const mul = (m, n) => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

/* ---------------- text ---------------- */

/** <Text> element → plain text (fields keep their stored value) */
export function shapeText(el) {
  if (!el) return '';
  let s = '';
  const walk = (n) => { for (const c of n.c || []) { if (typeof c === 'string') s += c; else if (c.n === 'fld' || c.n === 'cp' || c.n === 'pp' || c.n === 'tp') walk(c); else walk(c); } };
  walk(el);
  return s.replace(/\r\n?/g, '\n').replace(/\u2028|\u2029/g, '\n').replace(/\n+$/, '');
}

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

function wrapLines(text, width, size) {
  const out = [];
  const maxChars = Math.max(1, Math.floor(width / (size * 0.52)));
  for (const para of text.split('\n')) {
    if (!width || para.length <= maxChars) { out.push(para); continue; }
    let line = '';
    for (const w of para.split(/\s+/)) {
      if (!line) line = w;
      else if ((`${line} ${w}`).length <= maxChars) line += ` ${w}`;
      else { out.push(line); line = w; }
    }
    if (line) out.push(line);
  }
  return out;
}

/* ---------------- the whole file ---------------- */

function relsOf(files, partPath) {
  const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
  const name = partPath.slice(partPath.lastIndexOf('/') + 1);
  const xml = files[`${dir}_rels/${name}.rels`];
  const map = {};
  if (!xml) return map;
  const root = parseXml(xml);
  for (const r of kids(root, 'Relationship')) {
    let target = r.a.Target || '';
    if (!target.startsWith('/')) {
      const parts = (dir + target).split('/'); const outp = [];
      for (const p of parts) { if (p === '..') outp.pop(); else if (p !== '.') outp.push(p); }
      target = outp.join('/');
    } else target = target.slice(1);
    map[r.a.Id] = target;
  }
  return map;
}
const relId = (el) => { const r = kid(el, 'Rel'); return r ? r.a['r:id'] || Object.entries(r.a).find(([k]) => /:id$/.test(k))?.[1] : null; };

/** theme colours of visio/theme/theme1.xml → { dk1, lt1, dk2, lt2, accent1 … accent6 } */
export function themeColors(xml) {
  const out = {};
  if (!xml) return out;
  for (const m of String(xml).matchAll(/<a:(dk1|lt1|dk2|lt2|accent[1-6])>\s*<a:(?:srgbClr val|sysClr[^>]*?lastClr)="([0-9A-Fa-f]{6})"/g)) out[m[1]] = `#${m[2].toLowerCase()}`;
  return out;
}
/** QuickStyle colour index (0 dk1, 1 lt1, 2…7 accent1…6, 100+ = the same with variants) → colour */
export function quickColor(theme, idx) {
  const i = Number(idx) % 100;
  const key = i === 0 ? 'dk1' : i === 1 ? 'lt1' : i >= 2 && i <= 7 ? `accent${i - 1}` : 'accent1';
  return theme[key] || null;
}
export function shade(hex, f) {
  const n = parseInt(hex.slice(1), 16); const k = (v) => Math.max(0, Math.min(255, Math.round(f < 0 ? v * (1 + f) : v + (255 - v) * f)));
  return `#${[k(n >> 16), k((n >> 8) & 255), k(n & 255)].map((x) => x.toString(16).padStart(2, '0')).join('')}`;
}
/** rounded corners (Visio "Rounding") for paths made of straight lines */
export function roundPath(d, r) {
  if (!(r > 0) || !/^[ML0-9.\s,eZ-]+$/.test(d) || /Z/.test(d.slice(0, -1))) return d;
  const closed = /Z\s*$/.test(d);
  const pts = [...d.matchAll(/[ML]\s*(-?[\d.]+(?:e-?\d+)?)[\s,]+(-?[\d.]+(?:e-?\d+)?)/g)].map((m) => [Number(m[1]), Number(m[2])]);
  if (pts.length < 3) return d;
  const f = (v) => (Math.round(v * 100) / 100).toString();
  let out = `M${f(pts[0][0])} ${f(pts[0][1])}`;
  for (let i = 1; i < pts.length; i++) {
    const p = pts[i];
    if (i === pts.length - 1 && !closed) { out += `L${f(p[0])} ${f(p[1])}`; break; }
    const a = pts[i - 1]; const b = pts[(i + 1) % pts.length];
    const la = Math.hypot(p[0] - a[0], p[1] - a[1]); const lb = Math.hypot(b[0] - p[0], b[1] - p[1]);
    const rr = Math.min(r, la / 2, lb / 2);
    if (rr < 0.01) { out += `L${f(p[0])} ${f(p[1])}`; continue; }
    const p1 = [p[0] + ((a[0] - p[0]) / la) * rr, p[1] + ((a[1] - p[1]) / la) * rr];
    const p2 = [p[0] + ((b[0] - p[0]) / lb) * rr, p[1] + ((b[1] - p[1]) / lb) * rr];
    out += `L${f(p1[0])} ${f(p1[1])}Q${f(p[0])} ${f(p[1])} ${f(p2[0])} ${f(p2[1])}`;
  }
  return closed ? `${out}Z` : out;
}

/* ---------------- orthogonal connector routing ---------------- */
const segHits = (a, b, r) => { // axis-parallel segment a→b crosses box r (shrunk a little)?
  const x1 = Math.min(a[0], b[0]); const x2 = Math.max(a[0], b[0]); const y1 = Math.min(a[1], b[1]); const y2 = Math.max(a[1], b[1]);
  return x2 > r.x1 + 1e-3 && x1 < r.x2 - 1e-3 && y2 > r.y1 + 1e-3 && y1 < r.y2 - 1e-3;
};
const sideOf = (p, r) => { if (!r) return null; const d = [[Math.abs(p[0] - r.x1), 'l'], [Math.abs(p[0] - r.x2), 'r'], [Math.abs(p[1] - r.y1), 'b'], [Math.abs(p[1] - r.y2), 't']]; d.sort((a, b) => a[0] - b[0]); return d[0][1]; };
const DIRV = { l: [-1, 0], r: [1, 0], b: [0, -1], t: [0, 1] };
/**
 * Orthogonal route from p0 to p1 (page inches, y up) around boxes [{x1,y1,x2,y2}] – like Visio's "right-angle"
 * dynamic connector: leave each glued shape perpendicular to its side, then the shortest candidate with one or
 * two bends (or a detour around the boxes) that crosses no other shape. from/to: the glued shapes' boxes.
 */
export function routeOrthogonal(p0, p1, boxes = [], from = null, to = null, stub = 0.2) {
  const s0 = sideOf(p0, from); const s1 = sideOf(p1, to);
  const a = s0 ? [p0[0] + DIRV[s0][0] * stub, p0[1] + DIRV[s0][1] * stub] : p0;
  const b = s1 ? [p1[0] + DIRV[s1][0] * stub, p1[1] + DIRV[s1][1] * stub] : p1;
  const obstacles = boxes.filter((r) => r !== from && r !== to);
  const cands = [];
  const xs = new Set([(a[0] + b[0]) / 2, a[0], b[0]]); const ys = new Set([(a[1] + b[1]) / 2, a[1], b[1]]);
  for (const r of boxes) { xs.add(r.x1 - stub); xs.add(r.x2 + stub); ys.add(r.y1 - stub); ys.add(r.y2 + stub); }
  for (const x of xs) cands.push([a, [x, a[1]], [x, b[1]], b]);
  for (const y of ys) cands.push([a, [a[0], y], [b[0], y], b]);
  const len = (pts) => pts.reduce((n, q, i) => (i ? n + Math.abs(q[0] - pts[i - 1][0]) + Math.abs(q[1] - pts[i - 1][1]) : 0), 0);
  const bends = (pts) => pts.filter((q, i) => i && i < pts.length - 1).length;
  let best = null;
  for (const c of cands) {
    const pts = [p0, ...c, p1].filter((q, i, arr) => !i || Math.abs(q[0] - arr[i - 1][0]) > 1e-6 || Math.abs(q[1] - arr[i - 1][1]) > 1e-6);
    let bad = false;
    for (let i = 1; i < pts.length && !bad; i++) {
      if (Math.abs(pts[i][0] - pts[i - 1][0]) > 1e-6 && Math.abs(pts[i][1] - pts[i - 1][1]) > 1e-6) { bad = true; break; }
      for (const r of obstacles) if (segHits(pts[i - 1], pts[i], r)) { bad = true; break; }
      // the glued shapes themselves: only the stubs may touch them
      if (!bad && i > 1 && i < pts.length - 1) for (const r of [from, to]) if (r && segHits(pts[i - 1], pts[i], r)) { bad = true; break; }
    }
    const score = len(pts) + bends(pts) * 0.05 + (bad ? 1e6 : 0);
    if (!best || score < best.score) best = { pts, score };
  }
  // drop collinear points
  const out = [];
  for (const q of best.pts) { const n = out.length; if (n >= 2 && ((Math.abs(out[n - 2][0] - out[n - 1][0]) < 1e-6 && Math.abs(out[n - 1][0] - q[0]) < 1e-6) || (Math.abs(out[n - 2][1] - out[n - 1][1]) < 1e-6 && Math.abs(out[n - 1][1] - q[1]) < 1e-6))) out[n - 1] = q; else out.push(q); }
  return out;
}

export function readVsdx(files) {
  const theme = themeColors(files['visio/theme/theme1.xml'] || Object.entries(files).find(([k]) => /^visio\/theme\/theme\d+\.xml$/.test(k))?.[1]);
  let gradSeq = 0;
  const media = files.media || {};
  // style sheets
  const styles = new Map();
  if (files['visio/document.xml']) {
    const doc = parseXml(files['visio/document.xml']);
    const ss = kid(doc, 'StyleSheets');
    for (const s of kids(ss, 'StyleSheet')) styles.set(s.a.ID, { ...readSheet(s), parents: [s.a.LineStyle, s.a.FillStyle, s.a.TextStyle] });
  }
  const styleCell = (id, name, depth = 0) => {
    const st = id != null && styles.get(String(id));
    if (!st || depth > 12) return undefined;
    if (st.cells[name] !== undefined) return st.cells[name];
    for (const p of st.parents) { if (p != null && p !== String(id)) { const v = styleCell(p, name, depth + 1); if (v !== undefined) return v; } }
    return undefined;
  };
  const styleChar = (id, depth = 0) => {
    const st = id != null && styles.get(String(id));
    if (!st || depth > 12) return null;
    const ch = st.sections.Character && st.sections.Character[0];
    if (ch && ch.rows[0]) return ch.rows[0].cells;
    return styleChar(st.parents[2], depth + 1);
  };

  // masters: id → Map(shapeId → { sheet, el, children ids }), plus top shape id
  const masters = new Map();
  if (files['visio/masters/masters.xml']) {
    const ms = parseXml(files['visio/masters/masters.xml']);
    const rels = relsOf(files, 'visio/masters/masters.xml');
    for (const m of kids(ms, 'Master')) {
      const target = rels[relId(m)];
      const xml = target && files[target];
      if (!xml) continue;
      const root = parseXml(xml);
      const shapes = new Map(); let top = null;
      const walk = (el, isTop) => {
        for (const sh of kids(kid(el, 'Shapes'), 'Shape')) {
          shapes.set(sh.a.ID, { el: sh, sheet: readSheet(sh), styles: [sh.a.LineStyle, sh.a.FillStyle, sh.a.TextStyle] });
          if (isTop && top == null) top = sh.a.ID;
          walk(sh, false);
        }
      };
      walk(root, true);
      const topCount = kids(kid(root, 'Shapes'), 'Shape').length;
      masters.set(m.a.ID, { shapes, top: topCount === 1 ? top : null, name: m.a.NameU || m.a.Name || '' });
    }
  }

  const pagesXml = files['visio/pages/pages.xml'];
  if (!pagesXml) return { pages: [] };
  const pagesRoot = parseXml(pagesXml);
  const prels = relsOf(files, 'visio/pages/pages.xml');
  const all = kids(pagesRoot, 'Page').map((p) => {
    const ps = readSheet(kid(p, 'PageSheet') || { c: [] });
    return { id: p.a.ID, name: p.a.NameU || p.a.Name || '', background: p.a.Background === '1', backPage: p.a.BackPage, part: prels[relId(p)], w: num(ps.cells.PageWidth, 8.5), h: num(ps.cells.PageHeight, 11) };
  });

  const renderPage = (pg) => {
    const xml = pg.part && files[pg.part];
    const texts = [];
    if (!xml) return { body: '', texts, defs: '' };
    const root = parseXml(xml);
    const H = pg.h;
    const prels2 = relsOf(files, pg.part);
    let body = '';
    let markers = false;
    let gdefs = '';
    // boxes of the top-level 2-D shapes (page inches) and the glue (Connects) of connectors – for routing
    const boxes = new Map();
    for (const sh of kids(kid(root, 'Shapes'), 'Shape')) {
      const c0 = readSheet(sh).cells;
      if (c0.BeginX != null) continue;
      const w0 = num(c0.Width); const h0 = num(c0.Height); if (!w0 || !h0) continue;
      const x0 = num(c0.PinX) - num(c0.LocPinX, w0 / 2); const y0 = num(c0.PinY) - num(c0.LocPinY, h0 / 2);
      boxes.set(sh.a.ID, { x1: x0, y1: y0, x2: x0 + w0, y2: y0 + h0 });
    }
    const glue = new Map(); // connector id → { begin: shape id, end: shape id }
    for (const cn of kids(kid(root, 'Connects'), 'Connect')) {
      const g = glue.get(cn.a.FromSheet) || {}; if (/^Begin/.test(cn.a.FromCell || '')) g.begin = cn.a.ToSheet; else if (/^End/.test(cn.a.FromCell || '')) g.end = cn.a.ToSheet; glue.set(cn.a.FromSheet, g);
    }

    const shapeOut = (el, parentM, masterId, inherited) => {
      const masterAttr = el.a.Master || masterId;
      const mst = masterAttr != null ? masters.get(String(masterAttr)) : null;
      let base = null; let baseStyles = [];
      if (mst) {
        const ref = el.a.MasterShape != null ? mst.shapes.get(el.a.MasterShape) : mst.top != null ? mst.shapes.get(mst.top) : null;
        if (ref) { base = ref.sheet; baseStyles = ref.styles; }
      }
      const own = readSheet(el);
      const sheet = mergeSheet(base || inherited, own);
      const st = [el.a.LineStyle || baseStyles[0], el.a.FillStyle || baseStyles[1], el.a.TextStyle || baseStyles[2]];
      const cell = (name, which) => {
        if (sheet.cells[name] !== undefined && sheet.cells[name] !== 'Themed') return sheet.cells[name];
        const order = which === 'line' ? [st[0], st[1], st[2]] : which === 'fill' ? [st[1], st[0], st[2]] : [st[2], st[0], st[1]];
        for (const s of order) { const v = styleCell(s, name); if (v !== undefined) return v; }
        return undefined;
      };
      const c = sheet.cells;
      const w = num(c.Width); const h = num(c.Height);
      const M = mul(parentM, shapeMatrix(c));
      const tf = (x, y) => { const [X, Y] = apply(M, x, y); return [X * S, (H - Y) * S]; };
      const hidden = cell('NoShow') === '1' || c.Visible === '0';

      if (!hidden) {
        // geometry
        const lp = num(cell('LinePattern', 'line'), 1);
        const lw = Math.max(0.25, num(cell('LineWeight', 'line'), 0.01) * S);
        const lcRaw = cell('LineColor', 'line');
        let lc = color(lcRaw, '#000000');
        if ((lcRaw === 'Themed' || lcRaw == null) && cell('QuickStyleLineColor', 'line') != null && quickColor(theme, cell('QuickStyleLineColor', 'line'))) lc = shade(quickColor(theme, cell('QuickStyleLineColor', 'line')), -0.25);
        const lt = num(cell('LineColorTrans', 'line'), 0);
        const fp = num(cell('FillPattern', 'fill'), 1);
        const fcRaw = cell('FillForegnd', 'fill');
        const qf = cell('QuickStyleFillColor', 'fill');
        let fc = color(fcRaw, '#ffffff');
        if ((fcRaw === 'Themed' || fcRaw == null) && qf != null && quickColor(theme, qf)) fc = quickColor(theme, qf);
        // gradient fill (FillGradientEnabled + FillGradient stops)
        let fillRef = null;
        if (cell('FillGradientEnabled', 'fill') === '1' && sheet.sections.FillGradient && sheet.sections.FillGradient[0]) {
          const stops = sheet.sections.FillGradient[0].rows.filter((r) => !r.del).map((r) => ({ c: color(r.cells.GradientStopColor, fc), p: num(r.cells.GradientStopPosition, 0), tr: num(r.cells.GradientStopColorTrans, 0) }));
          if (stops.length >= 2) {
            let ang = num(cell('FillGradientAngle', 'fill'), Math.PI / 2); if (Math.abs(ang) > 7) ang = (ang * Math.PI) / 180;
            const id = `vsg${++gradSeq}`; const cx = Math.cos(ang) / 2; const cy = Math.sin(ang) / 2;
            gdefs += `<linearGradient id="${id}" x1="${fmt(0.5 - cx)}" y1="${fmt(0.5 + cy)}" x2="${fmt(0.5 + cx)}" y2="${fmt(0.5 - cy)}">${stops.map((x) => `<stop offset="${fmt(Math.max(0, Math.min(1, x.p)))}" stop-color="${x.c}"${x.tr ? ` stop-opacity="${fmt(1 - x.tr)}"` : ''}/>`).join('')}</linearGradient>`;
            fillRef = `url(#${id})`;
          }
        }
        const rounding = num(cell('Rounding', 'line'), 0) * S;
        const ft = num(cell('FillForegndTrans', 'fill'), 0);
        const begin = num(cell('BeginArrow', 'line'), 0); const end = num(cell('EndArrow', 'line'), 0);
        // a connector without its own geometry (files not saved by Visio): route it at right angles around the shapes
        const routed = c.BeginX != null && c.EndX != null && !(own.sections.Geometry && own.sections.Geometry.length) && parentM[0] === 1 && parentM[4] === 0 && parentM[5] === 0 && ![2, 16].includes(num(cell('ShapeRouteStyle', 'line'), 0)) && num(cell('ConLineRouteExt', 'line'), 0) !== 2;
        if (routed) {
          const gl = glue.get(el.a.ID) || {};
          const pts = routeOrthogonal([num(c.BeginX), num(c.BeginY)], [num(c.EndX), num(c.EndY)], [...boxes.values()], boxes.get(gl.begin) || null, boxes.get(gl.end) || null);
          const d = pts.map(([x, y], i) => `${i ? 'L' : 'M'}${fmt(x * S)} ${fmt((H - y) * S)}`).join('');
          const dd = rounding > 0 ? roundPath(d, rounding) : d;
          if (begin > 0 || end > 0) markers = true;
          body += `<path d="${dd}" fill="none" stroke="${lp === 0 ? 'none' : lc}" stroke-width="${fmt(lw)}"${begin > 0 ? ' marker-start="url(#vsa)"' : ''}${end > 0 ? ' marker-end="url(#vsa)"' : ''}/>`;
        }
        for (const g of routed ? [] : sheet.sections.Geometry || []) {
          if (g.cells.NoShow === '1') continue;
          const gp = geometryToPath(g.rows, w, h, tf); const closed = gp.closed; const d = rounding > 0 && gp.d ? roundPath(gp.d, rounding) : gp.d;
          if (!d) continue;
          const noFill = g.cells.NoFill === '1' || !closed || fp === 0;
          const noLine = g.cells.NoLine === '1' || lp === 0;
          const dash = DASH[lp] ? ` stroke-dasharray="${DASH[lp].map((x) => fmt(x * Math.max(1, lw / 1.5))).join(' ')}"` : '';
          const arrows = !closed && !noLine ? `${begin > 0 ? ' marker-start="url(#vsa)"' : ''}${end > 0 ? ' marker-end="url(#vsa)"' : ''}` : '';
          if (arrows) markers = true;
          body += `<path d="${d}" fill="${noFill ? 'none' : fillRef || fc}"${!noFill && ft ? ` fill-opacity="${fmt(1 - ft)}"` : ''} stroke="${noLine ? 'none' : lc}"${lt ? ` stroke-opacity="${fmt(1 - lt)}"` : ''} stroke-width="${fmt(lw)}"${dash}${arrows} stroke-linejoin="round" style="color:${lc}"/>`;
        }
        // images
        const fd = kid(el, 'ForeignData');
        if (fd) {
          const target = prels2[relId(fd)];
          const url = target && media[target];
          const ix = num(c.ImgOffsetX); const iy = num(c.ImgOffsetY); const iw = num(c.ImgWidth, w); const ih = num(c.ImgHeight, h);
          if (url) {
            const [a0, b0] = tf(ix, iy + ih); const [a1, b1] = tf(ix + iw, iy + ih); const [a2, b2] = tf(ix, iy);
            const m = [(a1 - a0) / iw, (b1 - b0) / iw, (a2 - a0) / ih, (b2 - b0) / ih, a0, b0];
            body += `<image href="${esc(url)}" width="${fmt(iw)}" height="${fmt(ih)}" preserveAspectRatio="none" transform="matrix(${m.map(fmt).join(' ')})"/>`;
          }
        }
        // text
        const textEl = kid(el, 'Text');
        const masterRef = mst && (el.a.MasterShape != null ? mst.shapes.get(el.a.MasterShape) : mst.top != null ? mst.shapes.get(mst.top) : null);
        const txt = textEl ? shapeText(textEl) : masterRef ? shapeText(kid(masterRef.el, 'Text')) : '';
        if (txt.trim()) {
          texts.push(txt);
          const ch = (sheet.sections.Character && sheet.sections.Character[0] && sheet.sections.Character[0].rows[0] && sheet.sections.Character[0].rows[0].cells) || styleChar(st[2]) || {};
          const para = (sheet.sections.Paragraph && sheet.sections.Paragraph[0] && sheet.sections.Paragraph[0].rows[0] && sheet.sections.Paragraph[0].rows[0].cells) || {};
          const size = Math.max(4, num(ch.Size, 0.1667) * S);
          const style = num(ch.Style, 0);
          const tcol = color(ch.Color, '#000000');
          const tw = num(c.TxtWidth, w); const th = num(c.TxtHeight, h);
          const tpx = num(c.TxtPinX, w / 2); const tpy = num(c.TxtPinY, h / 2);
          const tlx = num(c.TxtLocPinX, tw / 2); const tly = num(c.TxtLocPinY, th / 2);
          const tang = num(c.TxtAngle);
          const halign = num(para.HorzAlign, 1); const valign = num(cell('VerticalAlign', 'text'), 1);
          const lines = wrapLines(txt, tw * S - 4, size);
          const lh = size * 1.2;
          // text block centre in output coords and its rotation
          const [cxo, cyo] = tf(tpx - tlx + tw / 2 * 1, tpy - tly + th / 2);
          const [rx, ry] = tf(tpx - tlx + tw / 2 + 1, tpy - tly + th / 2);
          const rot = (Math.atan2(ry - cyo, rx - cxo) * 180) / Math.PI - (tang * 180) / Math.PI;
          const total = lines.length * lh;
          const top = valign === 0 ? -th * S / 2 + 2 : valign === 2 ? th * S / 2 - total - 2 : -total / 2;
          const anchor = halign === 0 ? 'start' : halign === 2 ? 'end' : 'middle';
          const x = halign === 0 ? -tw * S / 2 + 2 : halign === 2 ? tw * S / 2 - 2 : 0;
          body += `<text transform="translate(${fmt(cxo)} ${fmt(cyo)}) rotate(${fmt(rot)})" font-size="${fmt(size)}" fill="${tcol}" text-anchor="${anchor}" font-family="Calibri, Arial, sans-serif"${style & 1 ? ' font-weight="700"' : ''}${style & 2 ? ' font-style="italic"' : ''}${style & 4 ? ' text-decoration="underline"' : ''}>`;
          lines.forEach((l, i) => { body += `<tspan x="${fmt(x)}" y="${fmt(top + (i + 0.8) * lh)}">${esc(l) || ' '}</tspan>`; });
          body += '</text>';
        }
      }
      // children (groups): instance children first; master children if the instance has none
      const childEls = kids(kid(el, 'Shapes'), 'Shape');
      if (childEls.length) for (const ch of childEls) shapeOut(ch, M, masterAttr, null);
      else if (mst && el.a.Type === 'Group' && !el.a.MasterShape && mst.top) {
        const topEl = mst.shapes.get(mst.top).el;
        for (const ch of kids(kid(topEl, 'Shapes'), 'Shape')) shapeOut({ ...ch, a: { ...ch.a, MasterShape: ch.a.ID, Master: undefined } }, M, masterAttr, null);
      }
    };

    for (const sh of kids(kid(root, 'Shapes'), 'Shape')) shapeOut(sh, [1, 0, 0, 1, 0, 0], null, null);
    const defs = markers ? '<marker id="vsa" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse" markerUnits="strokeWidth"><path d="M0 0L10 5L0 10z" fill="context-stroke"/></marker>' : '';
    return { body, texts, defs: defs + gdefs };
  };

  const byId = new Map(all.map((p) => [p.id, p]));
  const pages = all.filter((p) => !p.background).map((p) => {
    const layers = [];
    let bg = p.backPage != null ? byId.get(p.backPage) : null;
    const seen = new Set();
    while (bg && !seen.has(bg.id)) { seen.add(bg.id); layers.unshift(renderPage({ ...bg, h: p.h })); bg = bg.backPage != null ? byId.get(bg.backPage) : null; }
    const main = renderPage(p);
    layers.push(main);
    const defs = layers.map((l) => l.defs).join('');
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${fmt(p.w * S)} ${fmt(p.h * S)}" width="${fmt(p.w * S)}" height="${fmt(p.h * S)}"><defs>${defs}</defs><rect width="100%" height="100%" fill="#ffffff"/>${layers.map((l) => l.body).join('')}</svg>`;
    return { name: p.name, w: p.w, h: p.h, svg, texts: main.texts };
  });
  return { pages };
}
