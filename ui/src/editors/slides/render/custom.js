/*
 * Custom geometry (freeform shapes, icons): a shape with geom 'custom' carries
 *   paths: [{ w, h, d, fill: bool (default true), stroke: bool (default true) }]
 * where d is an absolute SVG path using only M L C Q Z in the path's own units (0..w, 0..h). It maps 1:1 to OOXML
 * a:custGeom/a:pathLst and to an ODF draw:enhanced-path. Pure functions (Node tests).
 */

const f = (n) => +(+n).toFixed(2);

/* ---------------------------------------------------------------- SVG path → absolute M/L/C/Q/Z */
const NUMS = /-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/gi;
function tokens(d) {
  const out = [];
  const re = /([MmLlHhVvCcSsQqTtAaZz])|(-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)/g;
  let m;
  while ((m = re.exec(d))) out.push(m[1] || parseFloat(m[2]));
  return out;
}
/** SVG elliptical arc → cubic segments (endpoint parameterisation, SVG spec F.6) */
function arcToCubics(x1, y1, rx, ry, phi, fa, fs, x2, y2) {
  if (!rx || !ry) return [[x1, y1, x2, y2, x2, y2]];
  rx = Math.abs(rx); ry = Math.abs(ry);
  const p = (phi * Math.PI) / 180; const cp = Math.cos(p); const sp = Math.sin(p);
  const dx = (x1 - x2) / 2; const dy = (y1 - y2) / 2;
  const x1p = cp * dx + sp * dy; const y1p = -sp * dx + cp * dy;
  let lam = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lam > 1) { lam = Math.sqrt(lam); rx *= lam; ry *= lam; }
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let co = Math.sqrt(Math.max(0, num / den));
  if (fa === fs) co = -co;
  const cxp = (co * rx * y1p) / ry; const cyp = (-co * ry * x1p) / rx;
  const cx = cp * cxp - sp * cyp + (x1 + x2) / 2; const cy = sp * cxp + cp * cyp + (y1 + y2) / 2;
  const ang = (ux, uy, vx, vy) => { const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy); return a; };
  const t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!fs && dt > 0) dt -= 2 * Math.PI; else if (fs && dt < 0) dt += 2 * Math.PI;
  const segs = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2)));
  const step = dt / segs;
  const out = [];
  const pt = (t) => [cx + rx * Math.cos(t) * cp - ry * Math.sin(t) * sp, cy + rx * Math.cos(t) * sp + ry * Math.sin(t) * cp];
  const der = (t) => [-rx * Math.sin(t) * cp - ry * Math.cos(t) * sp, -rx * Math.sin(t) * sp + ry * Math.cos(t) * cp];
  for (let i = 0; i < segs; i++) {
    const a = t1 + step * i; const b = a + step;
    const k = (4 / 3) * Math.tan(step / 4);
    const [ax, ay] = pt(a); const [bx, by] = pt(b); const [dax, day] = der(a); const [dbx, dby] = der(b);
    out.push([ax + k * dax, ay + k * day, bx - k * dbx, by - k * dby, bx, by]);
  }
  return out;
}
/** any SVG path data → absolute path with M L C Q Z only */
export function normalizePath(d) {
  const tk = tokens(String(d || ''));
  let i = 0; let cmd = null; let x = 0; let y = 0; let sx = 0; let sy = 0; let lc = null; let lq = null;
  const out = [];
  const n = () => tk[i++];
  while (i < tk.length) {
    if (typeof tk[i] === 'string') cmd = tk[i++];
    else if (!cmd) { i++; continue; }
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();
    const ox = rel ? x : 0; const oy = rel ? y : 0;
    if (C === 'Z') { out.push('Z'); x = sx; y = sy; lc = lq = null; cmd = null; continue; }
    if (C === 'M') { x = ox + n(); y = oy + n(); sx = x; sy = y; out.push(`M${f(x)} ${f(y)}`); cmd = rel ? 'l' : 'L'; lc = lq = null; continue; }
    if (C === 'L') { x = ox + n(); y = oy + n(); out.push(`L${f(x)} ${f(y)}`); lc = lq = null; continue; }
    if (C === 'H') { x = ox + n(); out.push(`L${f(x)} ${f(y)}`); lc = lq = null; continue; }
    if (C === 'V') { y = (rel ? y : 0) + n(); out.push(`L${f(x)} ${f(y)}`); lc = lq = null; continue; }
    if (C === 'C' || C === 'S') {
      let x1; let y1;
      if (C === 'C') { x1 = ox + n(); y1 = oy + n(); } else { x1 = lc ? 2 * x - lc[0] : x; y1 = lc ? 2 * y - lc[1] : y; }
      const x2 = ox + n(); const y2 = oy + n(); const ex = ox + n(); const ey = oy + n();
      out.push(`C${f(x1)} ${f(y1)} ${f(x2)} ${f(y2)} ${f(ex)} ${f(ey)}`);
      lc = [x2, y2]; lq = null; x = ex; y = ey; continue;
    }
    if (C === 'Q' || C === 'T') {
      let x1; let y1;
      if (C === 'Q') { x1 = ox + n(); y1 = oy + n(); } else { x1 = lq ? 2 * x - lq[0] : x; y1 = lq ? 2 * y - lq[1] : y; }
      const ex = ox + n(); const ey = oy + n();
      out.push(`Q${f(x1)} ${f(y1)} ${f(ex)} ${f(ey)}`);
      lq = [x1, y1]; lc = null; x = ex; y = ey; continue;
    }
    if (C === 'A') {
      const rx = n(); const ry = n(); const rot = n(); const fa = n() ? 1 : 0; const fs = n() ? 1 : 0; const ex = ox + n(); const ey = oy + n();
      for (const c of arcToCubics(x, y, rx, ry, rot, fa, fs, ex, ey)) out.push(`C${c.map(f).join(' ')}`);
      x = ex; y = ey; lc = lq = null; continue;
    }
    i++; // unknown
  }
  return out.join('');
}

/** parsed segments of a normalized path: [{ c: 'M'|'L'|'C'|'Q'|'Z', p: [numbers] }] */
export function segments(d) {
  const out = [];
  const re = /([MLCQZ])([^MLCQZ]*)/g;
  let m;
  while ((m = re.exec(String(d || '')))) out.push({ c: m[1], p: (m[2].match(NUMS) || []).map(Number) });
  return out;
}

/** svg path for drawing a custom shape at w×h (each sub path scaled from its own units) */
export function customPaths(paths, w, h) {
  return (paths || []).map((p) => {
    const sx = p.w ? w / p.w : 1; const sy = p.h ? h / p.h : 1;
    const d = segments(p.d).map((s) => s.c + s.p.map((v, i) => f(v * (i % 2 ? sy : sx))).join(' ')).join('');
    return { d, fill: p.fill !== false, stroke: p.stroke !== false };
  });
}

/* ---------------------------------------------------------------- OOXML */
/** a:custGeom for the paths (coordinates scaled to integers) */
export function custGeomXml(paths) {
  const k = 100;
  const I = (v) => Math.round(v * k);
  const pl = (paths || []).map((p) => {
    const attrs = [`w="${I(p.w || 1)}"`, `h="${I(p.h || 1)}"`];
    if (p.fill === false) attrs.push('fill="none"');
    if (p.stroke === false) attrs.push('stroke="0"');
    const body = segments(p.d).map((s) => {
      const pts = []; for (let i = 0; i + 1 < s.p.length; i += 2) pts.push(`<a:pt x="${I(s.p[i])}" y="${I(s.p[i + 1])}"/>`);
      if (s.c === 'M') return `<a:moveTo>${pts[0] || '<a:pt x="0" y="0"/>'}</a:moveTo>`;
      if (s.c === 'L') return `<a:lnTo>${pts[0]}</a:lnTo>`;
      if (s.c === 'C') return `<a:cubicBezTo>${pts.slice(0, 3).join('')}</a:cubicBezTo>`;
      if (s.c === 'Q') return `<a:quadBezTo>${pts.slice(0, 2).join('')}</a:quadBezTo>`;
      return '<a:close/>';
    }).join('');
    return `<a:path ${attrs.join(' ')}>${body}</a:path>`;
  }).join('');
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${pl}</a:pathLst></a:custGeom>`;
}

/**
 * paths from an a:custGeom node (xmltree: { n, a, c }). Returns null when the path uses guide names we can't
 * evaluate (the caller keeps the shape as a rectangle and flags it).
 */
export function pathsOfCustGeom(cg) {
  const ln = (n) => n.slice(n.indexOf(':') + 1);
  const kidsOf = (el, name) => (el && el.c ? el.c.filter((x) => typeof x !== 'string' && (!name || ln(x.n) === name)) : []);
  const pl = kidsOf(cg, 'pathLst')[0];
  if (!pl) return null;
  const out = [];
  for (const p of kidsOf(pl, 'path')) {
    const w = Number(p.a.w || 0); const hh = Number(p.a.h || 0);
    let cx = 0; let cy = 0; let ok = true;
    const val = (v) => { const n = Number(v); if (Number.isNaN(n)) ok = false; return n; };
    const parts = [];
    for (const c of kidsOf(p)) {
      const t = ln(c.n);
      const pts = kidsOf(c, 'pt').map((q) => [val(q.a.x), val(q.a.y)]);
      if (t === 'moveTo' && pts[0]) { [cx, cy] = pts[0]; parts.push(`M${f(cx)} ${f(cy)}`); }
      else if (t === 'lnTo' && pts[0]) { [cx, cy] = pts[0]; parts.push(`L${f(cx)} ${f(cy)}`); }
      else if (t === 'cubicBezTo' && pts.length >= 3) { parts.push(`C${pts.slice(0, 3).flat().map(f).join(' ')}`); [cx, cy] = pts[2]; }
      else if (t === 'quadBezTo' && pts.length >= 2) { parts.push(`Q${pts.slice(0, 2).flat().map(f).join(' ')}`); [cx, cy] = pts[1]; }
      else if (t === 'arcTo') {
        const wR = val(c.a.wR); const hR = val(c.a.hR); const st = val(c.a.stAng) / 60000; const sw = val(c.a.swAng) / 60000;
        if (!ok) break;
        const a0 = (st * Math.PI) / 180;
        const ccx = cx - wR * Math.cos(a0); const ccy = cy - hR * Math.sin(a0);
        const segs = Math.max(1, Math.ceil(Math.abs(sw) / 90));
        const step = ((sw / segs) * Math.PI) / 180;
        for (let i = 0; i < segs; i++) {
          const s = a0 + step * i; const e = s + step; const k = (4 / 3) * Math.tan((e - s) / 4);
          const x1 = ccx + wR * (Math.cos(s) - k * Math.sin(s)); const y1 = ccy + hR * (Math.sin(s) + k * Math.cos(s));
          const x2 = ccx + wR * (Math.cos(e) + k * Math.sin(e)); const y2 = ccy + hR * (Math.sin(e) - k * Math.cos(e));
          cx = ccx + wR * Math.cos(e); cy = ccy + hR * Math.sin(e);
          parts.push(`C${[x1, y1, x2, y2, cx, cy].map(f).join(' ')}`);
        }
      } else if (t === 'close') parts.push('Z');
      if (!ok) break;
    }
    if (!ok) return null;
    out.push({ w: w || 1, h: hh || 1, d: parts.join(''), ...(p.a.fill === 'none' ? { fill: false } : {}), ...(p.a.stroke === '0' || p.a.stroke === 'false' ? { stroke: false } : {}) });
  }
  // path size 0 = shape size: the caller passes it in as w/h later (ext of the shape)
  return out.length ? out : null;
}

/* ---------------------------------------------------------------- ODF */
/** draw:enhanced-path + viewBox for custom paths (one viewBox: all sub paths scaled into the first one's units) */
export function odfCustom(paths) {
  const W = Math.max(1, Math.round((paths[0] && paths[0].w) || 100)); const H = Math.max(1, Math.round((paths[0] && paths[0].h) || 100));
  const path = customPaths(paths, W, H).map((p) => `${p.fill ? '' : 'F '}${p.stroke ? '' : 'S '}${p.d.replace(/([MLCQZ])/g, ' $1 ').trim()} N`).join(' ');
  return { viewBox: `0 0 ${W} ${H}`, path };
}
/** parse an ODF enhanced-path with plain numbers (no equations) → paths; null when it uses ?equations or unsupported commands */
export function pathsOfOdf(enhanced, viewBox) {
  if (!enhanced || /[?$]/.test(enhanced)) return null;
  const vb = String(viewBox || '0 0 21600 21600').trim().split(/[\s,]+/).map(Number);
  const W = vb[2] || 21600; const H = vb[3] || 21600; const x0 = vb[0] || 0; const y0 = vb[1] || 0;
  const out = [];
  let cur = { fill: true, stroke: true, parts: [] };
  const re = /([A-Za-z])|(-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)/g;
  const tk = []; let m; while ((m = re.exec(enhanced))) tk.push(m[1] || parseFloat(m[2]));
  let i = 0; let cmd = null;
  const n = () => tk[i++];
  const P = () => { const x = n() - x0; const y = n() - y0; return `${f(x)} ${f(y)}`; };
  while (i < tk.length) {
    if (typeof tk[i] === 'string') cmd = tk[i++];
    if (cmd === 'N') { if (cur.parts.length) out.push(cur); cur = { fill: true, stroke: true, parts: [] }; cmd = null; continue; }
    if (cmd === 'F') { cur.fill = false; cmd = null; continue; }
    if (cmd === 'S') { cur.stroke = false; cmd = null; continue; }
    if (cmd === 'Z') { cur.parts.push('Z'); cmd = null; continue; }
    if (typeof tk[i] !== 'number') { if (cmd && !'MLCQ'.includes(cmd)) return null; if (!cmd) i++; continue; }
    if (cmd === 'M') { cur.parts.push(`M${P()}`); cmd = 'L'; } else if (cmd === 'L') cur.parts.push(`L${P()}`);
    else if (cmd === 'C') cur.parts.push(`C${P()} ${P()} ${P()}`);
    else if (cmd === 'Q') cur.parts.push(`Q${P()} ${P()}`);
    else return null;
  }
  if (cur.parts.length) out.push(cur);
  if (!out.length) return null;
  return out.map((p) => ({ w: W, h: H, d: p.parts.join(''), ...(p.fill ? {} : { fill: false }), ...(p.stroke ? {} : { stroke: false }) }));
}
