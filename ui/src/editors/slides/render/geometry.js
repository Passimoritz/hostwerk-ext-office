/*
 * Preset shape geometry. Ids are the OOXML prstGeom names (so pptx maps 1:1); `odf` is the ODF draw:type.
 * shapePath(geom, w, h, adj) → { d, text } – an SVG path using only M/L/C/Z with absolute coordinates (also valid as an
 * ODF draw:enhanced-path after the numbers) and the text rectangle inside the shape.
 */
const K = 0.5522847498; // bezier circle constant

const f = (n) => +n.toFixed(2);
function ellipsePath(cx, cy, rx, ry) {
  const kx = rx * K; const ky = ry * K;
  return `M${f(cx + rx)} ${f(cy)}C${f(cx + rx)} ${f(cy + ky)} ${f(cx + kx)} ${f(cy + ry)} ${f(cx)} ${f(cy + ry)}C${f(cx - kx)} ${f(cy + ry)} ${f(cx - rx)} ${f(cy + ky)} ${f(cx - rx)} ${f(cy)}C${f(cx - rx)} ${f(cy - ky)} ${f(cx - kx)} ${f(cy - ry)} ${f(cx)} ${f(cy - ry)}C${f(cx + kx)} ${f(cy - ry)} ${f(cx + rx)} ${f(cy - ky)} ${f(cx + rx)} ${f(cy)}Z`;
}
/** arc from angle a0 to a1 (degrees, clockwise in screen space) as cubic segments, without the initial M */
function arcTo(cx, cy, rx, ry, a0, a1) {
  let out = '';
  const segs = Math.max(1, Math.ceil(Math.abs(a1 - a0) / 90));
  const step = (a1 - a0) / segs;
  for (let i = 0; i < segs; i++) {
    const s = ((a0 + step * i) * Math.PI) / 180; const e = ((a0 + step * (i + 1)) * Math.PI) / 180;
    const k = (4 / 3) * Math.tan((e - s) / 4);
    const x1 = cx + rx * (Math.cos(s) - k * Math.sin(s)); const y1 = cy + ry * (Math.sin(s) + k * Math.cos(s));
    const x2 = cx + rx * (Math.cos(e) + k * Math.sin(e)); const y2 = cy + ry * (Math.sin(e) - k * Math.cos(e));
    out += `C${f(x1)} ${f(y1)} ${f(x2)} ${f(y2)} ${f(cx + rx * Math.cos(e))} ${f(cy + ry * Math.sin(e))}`;
  }
  return out;
}
const pt = (cx, cy, rx, ry, a) => [cx + rx * Math.cos((a * Math.PI) / 180), cy + ry * Math.sin((a * Math.PI) / 180)];
const poly = (pts) => `M${pts.map((p) => `${f(p[0])} ${f(p[1])}`).join('L')}Z`;
function roundRect(x, y, w, h, r) {
  r = Math.max(0, Math.min(r, w / 2, h / 2));
  if (!r) return poly([[x, y], [x + w, y], [x + w, y + h], [x, y + h]]);
  return `M${f(x + r)} ${f(y)}L${f(x + w - r)} ${f(y)}${arcTo(x + w - r, y + r, r, r, -90, 0)}L${f(x + w)} ${f(y + h - r)}${arcTo(x + w - r, y + h - r, r, r, 0, 90)}L${f(x + r)} ${f(y + h)}${arcTo(x + r, y + h - r, r, r, 90, 180)}L${f(x)} ${f(y + r)}${arcTo(x + r, y + r, r, r, 180, 270)}Z`;
}
function star(n, w, h, inner) {
  const pts = [];
  for (let i = 0; i < n * 2; i++) {
    const a = -90 + (i * 180) / n;
    const r = i % 2 ? inner : 1;
    pts.push(pt(w / 2, h / 2, (w / 2) * r, (h / 2) * r, a));
  }
  return poly(pts);
}
function regular(n, w, h, rot = -90) {
  const pts = [];
  for (let i = 0; i < n; i++) pts.push(pt(w / 2, h / 2, w / 2, h / 2, rot + (i * 360) / n));
  return poly(pts);
}

const R = (w, h) => ({ x: 0, y: 0, w, h });
const inset = (w, h, fx, fy = fx) => ({ x: w * fx, y: h * fy, w: w * (1 - 2 * fx), h: h * (1 - 2 * fy) });

/** label, group, odf type, path builder */
export const SHAPES = {
  rect: { label: 'Rectangle', group: 'basic', odf: 'rectangle', p: (w, h) => ({ d: poly([[0, 0], [w, 0], [w, h], [0, h]]), text: R(w, h) }) },
  roundRect: { label: 'Rounded rectangle', group: 'basic', odf: 'round-rectangle', adj: [16667], p: (w, h, a) => { const r = Math.min(w, h) * (a[0] / 100000); return { d: roundRect(0, 0, w, h, r), text: inset(w, h, (r * 0.29) / w || 0, (r * 0.29) / h || 0) }; } },
  snip1Rect: { label: 'Snipped rectangle', group: 'basic', odf: 'ooxml-snip1Rect', adj: [16667], p: (w, h, a) => { const s = Math.min(w, h) * (a[0] / 100000); return { d: poly([[0, 0], [w - s, 0], [w, s], [w, h], [0, h]]), text: R(w, h) }; } },
  ellipse: { label: 'Oval', group: 'basic', odf: 'ellipse', p: (w, h) => ({ d: ellipsePath(w / 2, h / 2, w / 2, h / 2), text: inset(w, h, 0.146) }) },
  triangle: { label: 'Triangle', group: 'basic', odf: 'isosceles-triangle', p: (w, h) => ({ d: poly([[w / 2, 0], [w, h], [0, h]]), text: { x: w / 4, y: h / 2, w: w / 2, h: h / 2 } }) },
  rtTriangle: { label: 'Right triangle', group: 'basic', odf: 'right-triangle', p: (w, h) => ({ d: poly([[0, 0], [w, h], [0, h]]), text: { x: w / 12, y: h * 0.58, w: w / 2, h: h * 0.42 } }) },
  diamond: { label: 'Diamond', group: 'basic', odf: 'diamond', p: (w, h) => ({ d: poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]), text: inset(w, h, 0.25) }) },
  parallelogram: { label: 'Parallelogram', group: 'basic', odf: 'parallelogram', adj: [25000], p: (w, h, a) => { const o = Math.min(w, h) * (a[0] / 100000); return { d: poly([[o, 0], [w, 0], [w - o, h], [0, h]]), text: { x: o / 2, y: 0, w: w - o, h } }; } },
  trapezoid: { label: 'Trapezoid', group: 'basic', odf: 'trapezoid', adj: [25000], p: (w, h, a) => { const o = Math.min(w, h) * (a[0] / 100000); return { d: poly([[o, 0], [w - o, 0], [w, h], [0, h]]), text: { x: o / 2, y: 0, w: w - o, h } }; } },
  pentagon: { label: 'Pentagon', group: 'basic', odf: 'pentagon', p: (w, h) => ({ d: regular(5, w, h), text: inset(w, h, 0.2) }) },
  hexagon: { label: 'Hexagon', group: 'basic', odf: 'hexagon', p: (w, h) => ({ d: poly([[w * 0.25, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [w * 0.25, h], [0, h / 2]]), text: inset(w, h, 0.15, 0.1) }) },
  octagon: { label: 'Octagon', group: 'basic', odf: 'octagon', p: (w, h) => { const o = Math.min(w, h) * 0.29; return { d: poly([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]), text: inset(w, h, 0.1) }; } },
  plus: { label: 'Cross', group: 'basic', odf: 'cross', adj: [25000], p: (w, h, a) => { const o = Math.min(w, h) * (a[0] / 100000); return { d: poly([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, h - o], [w - o, h - o], [w - o, h], [o, h], [o, h - o], [0, h - o], [0, o], [o, o]]), text: { x: 0, y: o, w, h: h - 2 * o } }; } },
  donut: { label: 'Ring', group: 'basic', odf: 'ring', adj: [25000], p: (w, h, a) => { const t = Math.min(w, h) * (a[0] / 100000); return { d: ellipsePath(w / 2, h / 2, w / 2, h / 2) + ellipsePath(w / 2, h / 2, Math.max(1, w / 2 - t), Math.max(1, h / 2 - t)), text: inset(w, h, 0.146), evenodd: true }; } },
  frame: { label: 'Frame', group: 'basic', odf: 'frame', adj: [12500], p: (w, h, a) => { const t = Math.min(w, h) * (a[0] / 100000); return { d: poly([[0, 0], [w, 0], [w, h], [0, h]]) + poly([[t, t], [t, h - t], [w - t, h - t], [w - t, t]]), text: inset(w, h, t / w, t / h), evenodd: true }; } },
  can: { label: 'Cylinder', group: 'basic', odf: 'can', p: (w, h) => { const ry = Math.min(h * 0.12, w * 0.25); return { d: `M0 ${f(ry)}${arcTo(w / 2, ry, w / 2, ry, 180, 360)}L${f(w)} ${f(h - ry)}${arcTo(w / 2, h - ry, w / 2, ry, 0, 180)}Z${ellipsePath(w / 2, ry, w / 2, ry)}`, text: { x: 0, y: ry * 2, w, h: h - ry * 3 } }; } },
  cube: { label: 'Cube', group: 'basic', odf: 'cube', p: (w, h) => { const o = Math.min(w, h) * 0.25; return { d: poly([[0, o], [o, 0], [w, 0], [w, h - o], [w - o, h], [0, h]]) + `M0 ${f(o)}L${f(w - o)} ${f(o)}L${f(w)} 0M${f(w - o)} ${f(o)}L${f(w - o)} ${f(h)}`, text: { x: 0, y: o, w: w - o, h: h - o } }; } },
  heart: { label: 'Heart', group: 'basic', odf: 'heart', p: (w, h) => ({ d: `M${f(w / 2)} ${f(h * 0.25)}C${f(w * 0.5)} ${f(-h * 0.08)} ${f(-w * 0.1)} ${f(h * 0.02)} ${f(w * 0.02)} ${f(h * 0.38)}C${f(w * 0.12)} ${f(h * 0.62)} ${f(w * 0.42)} ${f(h * 0.8)} ${f(w / 2)} ${f(h)}C${f(w * 0.58)} ${f(h * 0.8)} ${f(w * 0.88)} ${f(h * 0.62)} ${f(w * 0.98)} ${f(h * 0.38)}C${f(w * 1.1)} ${f(h * 0.02)} ${f(w * 0.5)} ${f(-h * 0.08)} ${f(w / 2)} ${f(h * 0.25)}Z`, text: inset(w, h, 0.2, 0.25) }) },
  lightningBolt: { label: 'Lightning bolt', group: 'basic', odf: 'lightning', p: (w, h) => ({ d: poly([[w * 0.39, 0], [w * 0.62, 0.3 * h], [w * 0.52, 0.34 * h], [w * 0.8, 0.62 * h], [w * 0.7, 0.66 * h], [w, h], [w * 0.42, 0.72 * h], [w * 0.53, 0.68 * h], [w * 0.2, 0.42 * h], [w * 0.32, 0.38 * h], [0, 0.14 * h]]), text: inset(w, h, 0.3) }) },
  moon: { label: 'Moon', group: 'basic', odf: 'moon', p: (w, h) => ({ d: `M${f(w)} 0${arcTo(w, h / 2, w, h / 2, 270, 90)}${arcTo(w, h / 2, w * 0.5, h / 2, 90, 270)}Z`, text: { x: w * 0.1, y: h * 0.25, w: w * 0.4, h: h / 2 } }) },
  smileyFace: { label: 'Smiley face', group: 'basic', odf: 'smiley', p: (w, h) => ({ d: ellipsePath(w / 2, h / 2, w / 2, h / 2) + ellipsePath(w * 0.35, h * 0.38, w * 0.05, h * 0.06) + ellipsePath(w * 0.65, h * 0.38, w * 0.05, h * 0.06) + `M${f(w * 0.28)} ${f(h * 0.62)}C${f(w * 0.4)} ${f(h * 0.78)} ${f(w * 0.6)} ${f(h * 0.78)} ${f(w * 0.72)} ${f(h * 0.62)}`, text: inset(w, h, 0.146) }) },
  sun: { label: 'Sun', group: 'basic', odf: 'sun', p: (w, h) => { let d = ellipsePath(w / 2, h / 2, w * 0.25, h * 0.25); for (let i = 0; i < 8; i++) { const a = i * 45; const [x1, y1] = pt(w / 2, h / 2, w * 0.34, h * 0.34, a - 8); const [x2, y2] = pt(w / 2, h / 2, w * 0.5, h * 0.5, a); const [x3, y3] = pt(w / 2, h / 2, w * 0.34, h * 0.34, a + 8); d += poly([[x1, y1], [x2, y2], [x3, y3]]); } return { d, text: inset(w, h, 0.3) }; } },
  cloud: { label: 'Cloud', group: 'basic', odf: 'cloud', p: (w, h) => ({ d: `M${f(w * 0.2)} ${f(h * 0.85)}C${f(-w * 0.05)} ${f(h * 0.85)} ${f(-w * 0.02)} ${f(h * 0.45)} ${f(w * 0.17)} ${f(h * 0.45)}C${f(w * 0.12)} ${f(h * 0.15)} ${f(w * 0.42)} ${f(h * 0.02)} ${f(w * 0.52)} ${f(h * 0.2)}C${f(w * 0.62)} ${f(h * 0.0)} ${f(w * 0.9)} ${f(h * 0.08)} ${f(w * 0.85)} ${f(h * 0.38)}C${f(w * 1.05)} ${f(h * 0.42)} ${f(w * 1.04)} ${f(h * 0.85)} ${f(w * 0.8)} ${f(h * 0.85)}Z`, text: inset(w, h, 0.18, 0.25) }) },
  star4: { label: '4-point star', group: 'stars', odf: 'star4', p: (w, h) => ({ d: star(4, w, h, 0.38), text: inset(w, h, 0.3) }) },
  star5: { label: '5-point star', group: 'stars', odf: 'star5', p: (w, h) => ({ d: star(5, w, h, 0.4), text: inset(w, h, 0.3) }) },
  star6: { label: '6-point star', group: 'stars', odf: 'star6', p: (w, h) => ({ d: star(6, w, h, 0.55), text: inset(w, h, 0.25) }) },
  star8: { label: '8-point star', group: 'stars', odf: 'star8', p: (w, h) => ({ d: star(8, w, h, 0.72), text: inset(w, h, 0.2) }) },
  star12: { label: '12-point star', group: 'stars', odf: 'ooxml-star12', p: (w, h) => ({ d: star(12, w, h, 0.78), text: inset(w, h, 0.18) }) },
  rightArrow: { label: 'Right arrow', group: 'arrows', odf: 'right-arrow', adj: [50000, 50000], p: (w, h, a) => { const t = h * (a[0] / 100000); const hd = Math.min(w, h * (a[1] / 100000)); return { d: poly([[0, (h - t) / 2], [w - hd, (h - t) / 2], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, (h + t) / 2], [0, (h + t) / 2]]), text: { x: 0, y: (h - t) / 2, w: w - hd / 2, h: t } }; } },
  leftArrow: { label: 'Left arrow', group: 'arrows', odf: 'left-arrow', adj: [50000, 50000], p: (w, h, a) => { const t = h * (a[0] / 100000); const hd = Math.min(w, h * (a[1] / 100000)); return { d: poly([[w, (h - t) / 2], [hd, (h - t) / 2], [hd, 0], [0, h / 2], [hd, h], [hd, (h + t) / 2], [w, (h + t) / 2]]), text: { x: hd / 2, y: (h - t) / 2, w: w - hd / 2, h: t } }; } },
  upArrow: { label: 'Up arrow', group: 'arrows', odf: 'up-arrow', adj: [50000, 50000], p: (w, h, a) => { const t = w * (a[0] / 100000); const hd = Math.min(h, w * (a[1] / 100000)); return { d: poly([[(w - t) / 2, h], [(w - t) / 2, hd], [0, hd], [w / 2, 0], [w, hd], [(w + t) / 2, hd], [(w + t) / 2, h]]), text: { x: (w - t) / 2, y: hd / 2, w: t, h: h - hd / 2 } }; } },
  downArrow: { label: 'Down arrow', group: 'arrows', odf: 'down-arrow', adj: [50000, 50000], p: (w, h, a) => { const t = w * (a[0] / 100000); const hd = Math.min(h, w * (a[1] / 100000)); return { d: poly([[(w - t) / 2, 0], [(w - t) / 2, h - hd], [0, h - hd], [w / 2, h], [w, h - hd], [(w + t) / 2, h - hd], [(w + t) / 2, 0]]), text: { x: (w - t) / 2, y: 0, w: t, h: h - hd / 2 } }; } },
  leftRightArrow: { label: 'Left-right arrow', group: 'arrows', odf: 'left-right-arrow', adj: [50000, 50000], p: (w, h, a) => { const t = h * (a[0] / 100000); const hd = Math.min(w / 2, h * (a[1] / 100000)); return { d: poly([[0, h / 2], [hd, 0], [hd, (h - t) / 2], [w - hd, (h - t) / 2], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, (h + t) / 2], [hd, (h + t) / 2], [hd, h]]), text: { x: hd / 2, y: (h - t) / 2, w: w - hd, h: t } }; } },
  upDownArrow: { label: 'Up-down arrow', group: 'arrows', odf: 'up-down-arrow', adj: [50000, 50000], p: (w, h, a) => { const t = w * (a[0] / 100000); const hd = Math.min(h / 2, w * (a[1] / 100000)); return { d: poly([[w / 2, 0], [w, hd], [(w + t) / 2, hd], [(w + t) / 2, h - hd], [w, h - hd], [w / 2, h], [0, h - hd], [(w - t) / 2, h - hd], [(w - t) / 2, hd], [0, hd]]), text: { x: (w - t) / 2, y: hd / 2, w: t, h: h - hd } }; } },
  chevron: { label: 'Chevron', group: 'arrows', odf: 'chevron', adj: [50000], p: (w, h, a) => { const o = Math.min(w, h * (a[0] / 100000)); return { d: poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h], [o, h / 2]]), text: { x: o, y: 0, w: w - 2 * o, h } }; } },
  homePlate: { label: 'Pentagon arrow', group: 'arrows', odf: 'pentagon-right', adj: [50000], p: (w, h, a) => { const o = Math.min(w, h * (a[0] / 100000)); return { d: poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h]]), text: { x: 0, y: 0, w: w - o / 2, h } }; } },
  notchedRightArrow: { label: 'Notched arrow', group: 'arrows', odf: 'notched-right-arrow', p: (w, h) => { const t = h * 0.5; const hd = Math.min(w, h * 0.5); return { d: poly([[0, (h - t) / 2], [w - hd, (h - t) / 2], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, (h + t) / 2], [0, (h + t) / 2], [hd * 0.5, h / 2]]), text: { x: hd / 2, y: (h - t) / 2, w: w - hd, h: t } }; } },
  wedgeRectCallout: { label: 'Rectangular callout', group: 'callouts', odf: 'rectangular-callout', adj: [-20833, 62500], p: (w, h, a) => { const tx = w / 2 + w * (a[0] / 100000); const ty = h / 2 + h * (a[1] / 100000); const bh = Math.min(h, ty < h ? h : h); return { d: poly([[0, 0], [w, 0], [w, bh], [w * 0.58, bh], [tx, ty], [w * 0.29, bh], [0, bh]]), text: R(w, h) }; } },
  wedgeRoundRectCallout: { label: 'Rounded callout', group: 'callouts', odf: 'round-rectangular-callout', adj: [-20833, 62500], p: (w, h, a) => { const tx = w / 2 + w * (a[0] / 100000); const ty = h / 2 + h * (a[1] / 100000); const r = Math.min(w, h) * 0.16; return { d: `M${f(r)} 0L${f(w - r)} 0${arcTo(w - r, r, r, r, -90, 0)}L${f(w)} ${f(h - r)}${arcTo(w - r, h - r, r, r, 0, 90)}L${f(w * 0.58)} ${f(h)}L${f(tx)} ${f(ty)}L${f(w * 0.29)} ${f(h)}L${f(r)} ${f(h)}${arcTo(r, h - r, r, r, 90, 180)}L0 ${f(r)}${arcTo(r, r, r, r, 180, 270)}Z`, text: inset(w, h, 0.05) }; } },
  wedgeEllipseCallout: { label: 'Oval callout', group: 'callouts', odf: 'round-callout', adj: [-20833, 62500], p: null },
  cloudCallout: { label: 'Cloud callout', group: 'callouts', odf: 'cloud-callout', p: (w, h) => { const c = SHAPES.cloud.p(w, h * 0.8); return { d: c.d + ellipsePath(w * 0.2, h * 0.9, w * 0.05, h * 0.04) + ellipsePath(w * 0.1, h * 0.98, w * 0.025, h * 0.02), text: c.text }; } },
  flowChartProcess: { label: 'Process', group: 'flowchart', odf: 'flowchart-process', p: (w, h) => SHAPES.rect.p(w, h) },
  flowChartDecision: { label: 'Decision', group: 'flowchart', odf: 'flowchart-decision', p: (w, h) => SHAPES.diamond.p(w, h) },
  flowChartTerminator: { label: 'Terminator', group: 'flowchart', odf: 'flowchart-terminator', p: (w, h) => ({ d: roundRect(0, 0, w, h, Math.min(h / 2, w / 2)), text: inset(w, h, 0.1, 0.1) }) },
  flowChartDocument: { label: 'Document', group: 'flowchart', odf: 'flowchart-document', p: (w, h) => ({ d: `M0 0L${f(w)} 0L${f(w)} ${f(h * 0.83)}C${f(w * 0.75)} ${f(h * 0.7)} ${f(w * 0.5)} ${f(h * 0.95)} ${f(w * 0.25)} ${f(h)}C${f(w * 0.12)} ${f(h)} ${f(w * 0.05)} ${f(h * 0.95)} 0 ${f(h * 0.92)}Z`, text: { x: 0, y: 0, w, h: h * 0.8 } }) },
  flowChartInputOutput: { label: 'Data', group: 'flowchart', odf: 'flowchart-data', p: (w, h) => ({ d: poly([[w * 0.2, 0], [w, 0], [w * 0.8, h], [0, h]]), text: inset(w, h, 0.2, 0) }) },
  flowChartPredefinedProcess: { label: 'Predefined process', group: 'flowchart', odf: 'flowchart-predefined-process', p: (w, h) => ({ d: poly([[0, 0], [w, 0], [w, h], [0, h]]) + `M${f(w * 0.12)} 0L${f(w * 0.12)} ${f(h)}M${f(w * 0.88)} 0L${f(w * 0.88)} ${f(h)}`, text: inset(w, h, 0.12, 0) }) },
  flowChartConnector: { label: 'Connector', group: 'flowchart', odf: 'flowchart-connector', p: (w, h) => SHAPES.ellipse.p(w, h) },
};
// oval callout: ellipse with a wedge towards the tail point
SHAPES.wedgeEllipseCallout.p = (w, h, a) => {
  const tx = w / 2 + w * (a[0] / 100000); const ty = h / 2 + h * (a[1] / 100000);
  const ang = (Math.atan2(ty - h / 2, tx - w / 2) * 180) / Math.PI;
  const [x1, y1] = pt(w / 2, h / 2, w / 2, h / 2, ang + 12);
  return { d: `M${f(x1)} ${f(y1)}${arcTo(w / 2, h / 2, w / 2, h / 2, ang + 12, ang + 348)}L${f(tx)} ${f(ty)}Z`, text: inset(w, h, 0.146) };
};

export const SHAPE_GROUPS = [
  { id: 'basic', label: 'Basic shapes' }, { id: 'arrows', label: 'Block arrows' }, { id: 'stars', label: 'Stars' },
  { id: 'callouts', label: 'Callouts' }, { id: 'flowchart', label: 'Flowchart' },
];
export const QUICK_SHAPES = ['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow', 'star5', 'wedgeRoundRectCallout', 'line', 'arrow'];

export function shapePath(geom, w, h, adj) {
  const def = SHAPES[geom] || SHAPES.rect;
  const a = adj && adj.length ? adj.map((v, i) => (v == null ? (def.adj || [])[i] : v)) : (def.adj || []);
  return def.p(Math.max(w, 0.01), Math.max(h, 0.01), a);
}

/** ODF draw:type → our geom */
export const ODF_TO_GEOM = Object.fromEntries(Object.entries(SHAPES).map(([k, v]) => [v.odf, k]));
Object.assign(ODF_TO_GEOM, { 'ooxml-rect': 'rect', 'ooxml-roundRect': 'roundRect', 'ooxml-ellipse': 'ellipse', 'ooxml-triangle': 'triangle', circle: 'ellipse', 'mso-spt100': 'rect', 'flowchart-alternate-process': 'roundRect', 'block-arc': 'donut' });
/** OOXML preset names we do not draw exactly – mapped to the closest one */
export const PRST_ALIAS = {
  roundRect: 'roundRect', round1Rect: 'roundRect', round2SameRect: 'roundRect', round2DiagRect: 'roundRect', snip2SameRect: 'snip1Rect', snipRoundRect: 'roundRect',
  flowChartAlternateProcess: 'roundRect', flowChartManualInput: 'trapezoid', flowChartPreparation: 'hexagon', flowChartOffpageConnector: 'homePlate', flowChartMerge: 'triangle', flowChartExtract: 'triangle',
  star7: 'star6', star10: 'star8', star16: 'star12', star24: 'star12', star32: 'star12', irregularSeal1: 'star8', irregularSeal2: 'star8',
  curvedRightArrow: 'rightArrow', curvedLeftArrow: 'leftArrow', stripedRightArrow: 'rightArrow', bentArrow: 'rightArrow', uturnArrow: 'upArrow', circularArrow: 'donut',
  rightArrowCallout: 'rightArrow', leftArrowCallout: 'leftArrow', upArrowCallout: 'upArrow', downArrowCallout: 'downArrow', quadArrow: 'plus', leftRightUpArrow: 'plus',
  borderCallout1: 'wedgeRectCallout', borderCallout2: 'wedgeRectCallout', callout1: 'wedgeRectCallout', callout2: 'wedgeRectCallout', accentCallout1: 'wedgeRectCallout',
  ellipseRibbon: 'rect', ribbon: 'rect', ribbon2: 'rect', wave: 'flowChartDocument', doubleWave: 'flowChartDocument', bevel: 'rect', plaque: 'roundRect', foldedCorner: 'snip1Rect',
  teardrop: 'ellipse', pie: 'ellipse', chord: 'ellipse', arc: 'ellipse', blockArc: 'donut', noSmoking: 'donut', decagon: 'octagon', dodecagon: 'octagon', heptagon: 'hexagon',
  halfFrame: 'frame', corner: 'frame', diagStripe: 'parallelogram', mathPlus: 'plus', mathMinus: 'rect', mathMultiply: 'plus', mathEqual: 'rect', mathDivide: 'plus', mathNotEqual: 'rect',
  actionButtonBlank: 'rect', flowChartAlternateProcess2: 'roundRect', textRect: 'rect',
};
export const resolvePrst = (p) => (SHAPES[p] ? p : PRST_ALIAS[p] || 'rect');
