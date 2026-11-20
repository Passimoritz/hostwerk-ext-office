/*
 * Canvas painting of the 1.1 cell decorations (render.js calls these): icon sets, sparklines, rich text runs,
 * co-editors' selections and the outline gutter (+/– buttons of grouped rows/columns). Pure canvas code.
 */
import { MAXR, MAXC, parseRange } from './formula/refs.js';

const RYG = { 3: ['#e0403a', '#f5c142', '#3aa55d'], 4: ['#e0403a', '#f5c142', '#9fc54a', '#3aa55d'], 5: ['#e0403a', '#f0883a', '#f5c142', '#9fc54a', '#3aa55d'] };
/** one icon of an icon set, centred in a box of size `s` at (x, y) top-left */
export function drawIcon(ctx, set, i, n, x, y, s) {
  const col = (RYG[n] || RYG[3])[i];
  const gray = /Gray/.test(set);
  const c = gray ? ['#8c8c8c', '#a6a6a6', '#bfbfbf', '#d9d9d9', '#595959'][i] : col;
  const cx = x + s / 2; const cy = y + s / 2; const r = s * 0.42;
  ctx.save();
  ctx.fillStyle = c; ctx.strokeStyle = c; ctx.lineWidth = Math.max(1, s / 9); ctx.lineCap = 'round'; ctx.lineJoin = 'round';
  if (/Arrows/.test(set)) {
    const dirs = n === 3 ? [180, 90, 0] : n === 4 ? [180, 135, 45, 0] : [180, 135, 90, 45, 0]; // 0 = up
    const a = ((dirs[i] - 90) * Math.PI) / 180;
    ctx.translate(cx, cy); ctx.rotate(a + Math.PI / 2);
    ctx.beginPath(); ctx.moveTo(0, -r); ctx.lineTo(r * 0.8, 0); ctx.lineTo(r * 0.3, 0); ctx.lineTo(r * 0.3, r); ctx.lineTo(-r * 0.3, r); ctx.lineTo(-r * 0.3, 0); ctx.lineTo(-r * 0.8, 0); ctx.closePath(); ctx.fill();
  } else if (/Triangles/.test(set)) {
    ctx.beginPath();
    if (i === 0) { ctx.moveTo(cx - r, cy - r * 0.5); ctx.lineTo(cx + r, cy - r * 0.5); ctx.lineTo(cx, cy + r * 0.7); } else if (i === n - 1) { ctx.moveTo(cx - r, cy + r * 0.5); ctx.lineTo(cx + r, cy + r * 0.5); ctx.lineTo(cx, cy - r * 0.7); } else { ctx.rect(cx - r * 0.7, cy - r * 0.2, r * 1.4, r * 0.4); }
    ctx.fill();
  } else if (/Flags/.test(set)) {
    ctx.fillRect(cx - r * 0.7, cy - r, Math.max(1, s / 10), r * 2);
    ctx.beginPath(); ctx.moveTo(cx - r * 0.6, cy - r); ctx.lineTo(cx + r * 0.9, cy - r * 0.55); ctx.lineTo(cx - r * 0.6, cy - r * 0.05); ctx.fill();
  } else if (/Symbols/.test(set)) {
    if (!/Symbols2/.test(set)) { ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fill(); ctx.strokeStyle = '#fff'; }
    ctx.beginPath();
    if (i === n - 1) { ctx.moveTo(cx - r * 0.45, cy); ctx.lineTo(cx - r * 0.1, cy + r * 0.4); ctx.lineTo(cx + r * 0.5, cy - r * 0.4); } else if (i === 0) { ctx.moveTo(cx - r * 0.4, cy - r * 0.4); ctx.lineTo(cx + r * 0.4, cy + r * 0.4); ctx.moveTo(cx + r * 0.4, cy - r * 0.4); ctx.lineTo(cx - r * 0.4, cy + r * 0.4); } else { ctx.moveTo(cx, cy - r * 0.5); ctx.lineTo(cx, cy + r * 0.1); ctx.moveTo(cx, cy + r * 0.45); ctx.lineTo(cx, cy + r * 0.5); }
    ctx.stroke();
  } else if (/Rating|Boxes|Quarters|Stars/.test(set)) {
    const k = i / (n - 1); // fill level 0..1
    ctx.fillStyle = '#4472c4'; ctx.strokeStyle = '#4472c4';
    if (/Quarters/.test(set)) { ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.lineWidth = 1; ctx.stroke(); if (k > 0) { ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2 * k); ctx.closePath(); ctx.fill(); } } else if (/Stars/.test(set)) {
      ctx.fillStyle = k === 1 ? '#f5c142' : k > 0 ? '#f5c142' : '#d9d9d9'; ctx.beginPath();
      for (let p = 0; p < 10; p++) { const rr = p % 2 ? r * 0.45 : r; const a = -Math.PI / 2 + (p * Math.PI) / 5; ctx.lineTo(cx + rr * Math.cos(a), cy + rr * Math.sin(a)); }
      ctx.closePath(); ctx.fill();
    } else for (let b = 0; b < 4; b++) { const h = (r * 2 * (b + 1)) / 4; ctx.globalAlpha = b / 3 <= k + 0.01 ? 1 : 0.25; ctx.fillRect(x + s * 0.12 + b * s * 0.2, y + s * 0.9 - h, s * 0.15, h); }
  } else { // traffic lights / signs: circles (4TrafficLights adds black)
    const cc = n === 4 && /TrafficLights/.test(set) ? ['#262626', '#e0403a', '#f5c142', '#3aa55d'][i] : c;
    ctx.fillStyle = cc;
    if (/Signs/.test(set) && i === 0) { ctx.beginPath(); ctx.moveTo(cx, cy - r); ctx.lineTo(cx + r, cy); ctx.lineTo(cx, cy + r); ctx.lineTo(cx - r, cy); ctx.fill(); } else if (/Signs/.test(set) && i === 1) { ctx.beginPath(); ctx.moveTo(cx, cy - r); ctx.lineTo(cx + r, cy + r * 0.8); ctx.lineTo(cx - r, cy + r * 0.8); ctx.fill(); } else {
      const g = ctx.createRadialGradient(cx - r * 0.3, cy - r * 0.3, r * 0.1, cx, cy, r);
      g.addColorStop(0, '#ffffff99'); g.addColorStop(0.35, cc); g.addColorStop(1, cc);
      ctx.fillStyle = g; ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fill();
      if (/TrafficLights2/.test(set)) { ctx.strokeStyle = '#262626'; ctx.lineWidth = 1; ctx.stroke(); }
    }
  }
  ctx.restore();
}

/** map of the sparkline cells of a sheet → { g (group), it (item) } (cached per version) */
export function sparkMap(sheet) {
  const list = sheet.sparklines || [];
  if (!list.length) return null;
  if (sheet._spk && sheet._spk.v === sheet.version && sheet._spk.n === list) return sheet._spk.m;
  const m = new Map();
  for (const g of list) for (const it of g.items) m.set(it.r * MAXC + it.c, { g, it });
  sheet._spk = { v: sheet.version, n: list, m };
  return m;
}
/** values of a sparkline source (one row or column) */
export function sparkValues(book, ref) {
  const g = parseRange(ref || ''); if (!g) return [];
  const sh = g.sheet ? book.sheetByName(g.sheet) : null; if (!sh) return [];
  const out = [];
  const n = Math.max(g.r2 - g.r1 + 1, g.c2 - g.c1 + 1);
  if (n > 2000) return out;
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const v = book.value(sh, r, c); out.push(typeof v === 'number' ? v : null); }
  return out;
}
export function drawSparkline(ctx, book, sp, rc, z) {
  const { g, it } = sp;
  const vals = sparkValues(book, it.ref);
  const nums = vals.filter((v) => v != null);
  if (!nums.length) return;
  const pad = 3 * z;
  const x0 = rc.x + pad; const y0 = rc.y + pad; const w = rc.w - pad * 2; const h = rc.h - pad * 2;
  if (w < 4 || h < 3) return;
  let min = g.min != null ? g.min : Math.min(...nums); let max = g.max != null ? g.max : Math.max(...nums);
  if (g.type !== 'line') { min = Math.min(0, min); max = Math.max(0, max); }
  if (max === min) { max += 1; min -= 1; }
  const Y = (v) => y0 + h - ((v - min) / (max - min)) * h;
  const n = vals.length;
  const hi = Math.max(...nums); const lo = Math.min(...nums);
  const firstI = vals.findIndex((v) => v != null); let lastI = n - 1; while (lastI > 0 && vals[lastI] == null) lastI--;
  const special = (v, k) => (g.high && v === hi ? g.highColor : g.low && v === lo ? g.lowColor : g.first && k === firstI ? g.firstColor || g.markerColor : g.last && k === lastI ? g.lastColor || g.markerColor : g.negative && v < 0 ? g.neg : null);
  ctx.save();
  ctx.beginPath(); ctx.rect(rc.x, rc.y, rc.w, rc.h); ctx.clip();
  if (g.type === 'line') {
    const X = (k) => x0 + (n === 1 ? w / 2 : (k / (n - 1)) * w);
    ctx.strokeStyle = g.color || '#376092'; ctx.lineWidth = Math.max(1, (g.weight || 0.75) * 1.33 * z); ctx.lineJoin = 'round';
    ctx.beginPath(); let pen = false;
    vals.forEach((v, k) => { if (v == null) { if (g.empty !== 'span') pen = false; return; } if (!pen) { ctx.moveTo(X(k), Y(v)); pen = true; } else ctx.lineTo(X(k), Y(v)); });
    ctx.stroke();
    vals.forEach((v, k) => { if (v == null) return; const col = special(v, k) || (g.markers ? g.markerColor : null); if (!col) return; ctx.fillStyle = col; ctx.beginPath(); ctx.arc(X(k), Y(v), Math.max(1.5, 2 * z), 0, Math.PI * 2); ctx.fill(); });
  } else {
    const bw = w / n;
    const zero = Y(0);
    vals.forEach((v, k) => {
      if (v == null) return;
      ctx.fillStyle = special(v, k) || (v < 0 && g.negative ? g.neg : g.color || '#376092');
      const x = x0 + k * bw + bw * 0.12; const ww = Math.max(1, bw * 0.76);
      if (g.type === 'stacked') { const mid = y0 + h / 2; ctx.fillRect(x, v >= 0 ? y0 : mid, ww, h / 2); } else ctx.fillRect(x, Math.min(zero, Y(v)), ww, Math.max(1, Math.abs(Y(v) - zero)));
    });
  }
  ctx.restore();
}

/** rich text runs on one line (or wrapped by \n); returns true when painted */
export function drawRich(ctx, runs, rc, { z, defFont, fontStack, h, va, pad, s }) {
  const baseSize = ((s && s.size) || defFont.size) * (4 / 3) * z;
  const fontOf = (r) => `${r.i || (s && s.i) ? 'italic ' : ''}${r.b || (s && s.b) ? 'bold ' : ''}${Math.max(1, (r.size ? r.size * (4 / 3) * z : baseSize) * (r.va ? 0.7 : 1)).toFixed(1)}px ${fontStack(r.font || (s && s.font) || defFont.font)}`;
  const lines = [[]];
  for (const r of runs) { const parts = String(r.t).split('\n'); parts.forEach((p, i) => { if (i) lines.push([]); if (p) lines[lines.length - 1].push({ ...r, t: p }); }); }
  const lh = baseSize * 1.22;
  const blockH = lines.length * lh;
  let y = va === 'top' ? rc.y + pad + baseSize * 0.95 : va === 'center' ? rc.y + (rc.h - blockH) / 2 + baseSize * 0.95 : rc.y + rc.h - blockH - pad * 0.6 + baseSize * 0.95 + (lh - baseSize) * 0.35;
  ctx.save(); ctx.beginPath(); ctx.rect(rc.x, rc.y, rc.w, rc.h); ctx.clip();
  ctx.textBaseline = 'alphabetic'; ctx.textAlign = 'left';
  for (const line of lines) {
    const widths = line.map((r) => { ctx.font = fontOf(r); return ctx.measureText(r.t).width; });
    const tw = widths.reduce((a, b) => a + b, 0);
    let x = h === 'right' ? rc.x + rc.w - pad - tw : h === 'center' ? rc.x + (rc.w - tw) / 2 : rc.x + pad;
    line.forEach((r, k) => {
      ctx.font = fontOf(r);
      ctx.fillStyle = r.color || (s && s.color) || '#000000';
      const dy = r.va === 'superscript' ? -baseSize * 0.35 : r.va === 'subscript' ? baseSize * 0.15 : 0;
      ctx.fillText(r.t, x, y + dy);
      if (r.u || r.st) { const lw = widths[k]; if (r.u) ctx.fillRect(x, y + Math.max(1, baseSize * 0.12), lw, Math.max(1, z)); if (r.st) ctx.fillRect(x, y - baseSize * 0.3, lw, Math.max(1, z)); }
      x += widths[k];
    });
    y += lh;
  }
  ctx.restore();
  return true;
}

/** co-editors' selections: [{ ranges, r, c, color, name }] */
export function drawRemote(ctx, g, remote, sheet) {
  for (const p of remote || []) {
    if (p.sheet && p.sheet !== sheet.name) continue;
    ctx.save();
    ctx.strokeStyle = p.color || '#e8590c'; ctx.fillStyle = `${p.color || '#e8590c'}1a`; ctx.lineWidth = 2;
    for (const gg of p.ranges || []) {
      const rr = g.rect(gg.r1, gg.c1, Math.min(gg.r2, MAXR - 1), Math.min(gg.c2, MAXC - 1));
      ctx.fillRect(rr.x, rr.y, rr.w, rr.h);
      ctx.strokeRect(Math.round(rr.x) + 0.5, Math.round(rr.y) + 0.5, Math.round(rr.w) - 1, Math.round(rr.h) - 1);
    }
    if (p.name && p.r != null) {
      const rr = g.rect(p.r, p.c, p.r, p.c);
      ctx.font = '600 10px "Segoe UI", Arial, sans-serif';
      const tw = ctx.measureText(p.name).width + 8;
      ctx.fillStyle = p.color || '#e8590c';
      const ty = rr.y - 15 < g.hh ? rr.y + rr.h : rr.y - 15;
      ctx.fillRect(rr.x - 1, ty, tw, 15);
      ctx.fillStyle = '#ffffff'; ctx.textBaseline = 'middle'; ctx.textAlign = 'left';
      ctx.fillText(p.name, rr.x + 3, ty + 7.5);
    }
    ctx.restore();
  }
}

/* ---------------- outline (grouped rows / columns) ---------------- */
export const outlineLevels = (m) => { let x = 0; for (const v of m.values()) if (v > x) x = v; return Math.min(7, x); };
export const OL = 14; // px per level in the gutter
/** groups of one axis: [{ level, a, b }] (inclusive, contiguous runs with level ≥ k) */
export function outlineGroups(map, levels) {
  const keys = [...map.keys()].sort((a, b) => a - b);
  const out = [];
  for (let k = 1; k <= levels; k++) {
    let a = null; let prev = null;
    for (const i of keys) {
      if ((map.get(i) || 0) >= k) { if (a == null || i !== prev + 1) { if (a != null) out.push({ level: k, a, b: prev }); a = i; } prev = i; }
    }
    if (a != null) out.push({ level: k, a, b: prev });
  }
  return out;
}
/** paint the gutters (left for rows, top for columns) and return the button rects for hit testing */
export function paintOutline(ctx, g, sheet, theme) {
  const btns = [];
  const lr = g.olr; const lc = g.olc;
  if (!lr && !lc) return btns;
  ctx.save();
  ctx.fillStyle = theme.head;
  if (lr) ctx.fillRect(0, 0, g.gx, g.H);
  if (lc) ctx.fillRect(0, 0, g.W, g.gy);
  ctx.strokeStyle = theme.headText; ctx.fillStyle = theme.headText; ctx.lineWidth = 1;
  ctx.font = '600 9px "Segoe UI", Arial, sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  const box = (x, y, label, kind, data) => {
    ctx.fillStyle = '#ffffff'; ctx.fillRect(x - 5, y - 5, 10, 10);
    ctx.strokeRect(x - 4.5, y - 4.5, 9, 9);
    ctx.fillStyle = theme.headText;
    if (label) ctx.fillText(label, x, y + 0.5); else { ctx.fillRect(x - 3, y - 0.5, 6, 1); if (kind === 'expand') ctx.fillRect(x - 0.5, y - 3, 1, 6); }
    btns.push({ x: x - 6, y: y - 6, w: 12, h: 12, ...data });
  };
  // level buttons in the corner
  for (let k = 1; k <= lr + 1 && lr; k++) box(4 + (k - 1) * OL + 3, (g.gy || 0) + g.hh0 / 2, String(k), 'level', { axis: 'row', level: k });
  for (let k = 1; k <= lc + 1 && lc; k++) box((g.gx || 0) + g.hw0 / 2, 4 + (k - 1) * OL + 3, String(k), 'level', { axis: 'col', level: k });
  // row groups: a bracket next to the rows, the +/- button below the group (summary row)
  if (lr) {
    for (const gr of outlineGroups(sheet.outlineRows, lr)) {
      const x = 4 + (gr.level - 1) * OL + 3;
      const collapsed = isCollapsed(sheet, 'row', gr);
      const sumR = gr.b + 1;
      const y1 = g.rowY(gr.a); const y2 = g.rowY(gr.b) + g.rowH(gr.b);
      if (!collapsed && y2 > g.hh && y1 < g.H) { ctx.fillStyle = theme.headText; ctx.fillRect(x, Math.max(g.hh, y1 + 2), 1, Math.max(0, Math.min(g.H, y2) - Math.max(g.hh, y1 + 2))); ctx.fillRect(x, y1 + 2, 4, 1); }
      const by = g.rowY(sumR) + g.rowH(sumR) / 2;
      if (by > g.hh && by < g.H) box(x, by, null, collapsed ? 'expand' : 'collapse', { axis: 'row', group: gr });
    }
  }
  if (lc) {
    for (const gr of outlineGroups(sheet.outlineCols, lc)) {
      const y = 4 + (gr.level - 1) * OL + 3;
      const collapsed = isCollapsed(sheet, 'col', gr);
      const sumC = gr.b + 1;
      const x1 = g.colX(gr.a); const x2 = g.colX(gr.b) + g.colW(gr.b);
      if (!collapsed && x2 > g.hw && x1 < g.W) { ctx.fillStyle = theme.headText; ctx.fillRect(Math.max(g.hw, x1 + 2), y, Math.max(0, Math.min(g.W, x2) - Math.max(g.hw, x1 + 2)), 1); ctx.fillRect(x1 + 2, y, 1, 4); }
      const bx = g.colX(sumC) + g.colW(sumC) / 2;
      if (bx > g.hw && bx < g.W) box(bx, y, null, collapsed ? 'expand' : 'collapse', { axis: 'col', group: gr });
    }
  }
  ctx.restore();
  return btns;
}
export function isCollapsed(sheet, axis, gr) {
  const hid = axis === 'row' ? sheet.hiddenRows : sheet.hiddenCols;
  for (let i = gr.a; i <= gr.b; i++) if (!hid.has(i)) return false;
  return true;
}
