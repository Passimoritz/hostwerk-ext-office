/*
 * Canvas painting of the grid: four quadrants (frozen rows/cols), only visible cells. The sheet area is "paper"
 * (white, black text) in light and dark theme; headers and selection use Hostwerk's tokens (read from CSS variables).
 */
import { MAXR, MAXC, colName } from './formula/refs.js';
import { formatValue } from './format/numfmt.js';
import { isErr } from './formula/errors.js';
import { toDisplay } from './formula/parser.js';
import { cfAt } from './cf.js';
import { tableStyleAt } from './tables.js';
import { drawIcon, sparkMap, drawSparkline, drawRich, drawRemote, paintOutline, outlineLevels, OL } from './paintPlus.js';
import { setDate1904 } from './format/dates.js';

const K = (r, c) => r * MAXC + c;
export const REF_COLORS = ['#2f6fe4', '#d9480f', '#7b2fbe', '#2b8a3e', '#c2255c', '#0b7285', '#e67700', '#5c940d'];
const FALLBACK = '"Segoe UI", Arial, sans-serif';
const fontStack = (f) => `"${String(f || 'Calibri').replace(/"/g, '')}", Calibri, Carlito, ${FALLBACK}`;

const widthCache = new Map();
function measure(ctx, font, text) {
  const k = `${font}\u0001${text}`;
  let w = widthCache.get(k);
  if (w === undefined) { ctx.font = font; w = ctx.measureText(text).width; if (widthCache.size > 20000) widthCache.clear(); widthCache.set(k, w); }
  return w;
}
const BW = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
const DASH = { dotted: [1, 2], hair: [1, 1], dashed: [4, 2], mediumDashed: [6, 3], dashDot: [5, 2, 1, 2], mediumDashDot: [6, 3, 2, 3], dashDotDot: [5, 2, 1, 2, 1, 2], mediumDashDotDot: [6, 3, 2, 3, 2, 3], slantDashDot: [6, 2, 2, 2] };

/** geometry helper shared by painting and hit testing */
export function makeGeo({ sheet, axes, z, scrollX, scrollY, W, H, headings, dpr }) {
  const { rows, cols } = axes;
  const hh0 = headings ? Math.round(22 * Math.min(1.4, Math.max(0.8, z))) : 0;
  const digits = String(Math.min(MAXR, rows.index((scrollY + H) / z) + 2)).length;
  const hw0 = headings ? Math.round((Math.max(3, digits) * 8 + 14) * Math.min(1.4, Math.max(0.8, z))) : 0;
  // outline gutters (grouped rows/columns): one lane per level + the level buttons
  const olr = headings ? outlineLevels(sheet.outlineRows) : 0; const olc = headings ? outlineLevels(sheet.outlineCols) : 0;
  const gx = olr ? (olr + 1) * OL + 2 : 0; const gy = olc ? (olc + 1) * OL + 2 : 0;
  const hw = hw0 + gx; const hh = hh0 + gy;
  const fr = Math.min(sheet.freeze.rows || 0, MAXR - 1); const fc = Math.min(sheet.freeze.cols || 0, MAXC - 1);
  const fy = rows.pos(fr) * z; const fx = cols.pos(fc) * z;
  const g = {
    z, hw, hh, hw0, hh0, gx, gy, olr, olc, fr, fc, fx, fy, W, H, scrollX, scrollY, rows, cols, dpr,
    /** device-pixel-snapped coordinate for 1px lines (crisp at any zoom / DPR) */
    px: (v) => { const d = dpr || 1; const w = Math.max(1, Math.round(d)); return (Math.round(v * d) + (w % 2 ? 0.5 : 0)) / d; },
    lw: Math.max(1, Math.round(dpr || 1)) / (dpr || 1),
    colX: (c) => (c < fc ? hw + cols.pos(c) * z : hw + fx + (cols.pos(c) - cols.pos(fc)) * z - scrollX),
    rowY: (r) => (r < fr ? hh + rows.pos(r) * z : hh + fy + (rows.pos(r) - rows.pos(fr)) * z - scrollY),
    colW: (c) => cols.size(c) * z,
    rowH: (r) => rows.size(r) * z,
    colAt: (x) => (x < hw + fx ? cols.index(Math.max(0, x - hw) / z) : cols.index(cols.pos(fc) + (x - hw - fx + scrollX) / z)),
    rowAt: (y) => (y < hh + fy ? rows.index(Math.max(0, y - hh) / z) : rows.index(rows.pos(fr) + (y - hh - fy + scrollY) / z)),
  };
  // visible index ranges of the scrolled area
  g.c0 = Math.max(fc, cols.index(cols.pos(fc) + scrollX / z));
  g.c1 = Math.min(MAXC - 1, cols.index(cols.pos(fc) + (scrollX + Math.max(0, W - hw - fx)) / z) + 1);
  g.r0 = Math.max(fr, rows.index(rows.pos(fr) + scrollY / z));
  g.r1 = Math.min(MAXR - 1, rows.index(rows.pos(fr) + (scrollY + Math.max(0, H - hh - fy)) / z) + 1);
  /** screen rect of a range (unclipped) */
  g.rect = (r1, c1, r2, c2) => { const x = g.colX(c1); const y = g.rowY(r1); return { x, y, w: g.colX(c2) + g.colW(c2) - x, h: g.rowY(r2) + g.rowH(r2) - y }; };
  return g;
}

/** quadrants: [rowFrom, rowTo, colFrom, colTo, clip] */
function quadrants(g) {
  const q = [];
  const rowsA = g.fr ? [[0, g.fr - 1, g.hh, g.hh + g.fy]] : [];
  const colsA = g.fc ? [[0, g.fc - 1, g.hw, g.hw + g.fx]] : [];
  rowsA.push([g.r0, g.r1, g.hh + g.fy, g.H]);
  colsA.push([g.c0, g.c1, g.hw + g.fx, g.W]);
  for (const [ra, rb, y0, y1] of rowsA) for (const [ca, cb, x0, x1] of colsA) q.push({ ra, rb, ca, cb, clip: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } });
  return q;
}

export function paint(ctx, g, st) {
  const { book, sheet, theme, sel, edit, refs, copy, drag, showFormulas, loc, defFont, readOnly } = st;
  const { dpr, W, H, z } = g;
  setDate1904(book.date1904);
  const tsAt = tableStyleAt(book, sheet);
  const spk = sparkMap(sheet);
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, W, H);
  const cf = cfAt(book, sheet);
  const showGrid = sheet.gridlines !== false;
  // merges index for visible area
  const merges = sheet.merges;
  const covered = new Set();
  const anchorOf = new Map();
  for (const m of merges) { if (m.r2 - m.r1 > 2000 || m.c2 - m.c1 > 200) continue; for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) { const k = K(r, c); if (r === m.r1 && c === m.c1) anchorOf.set(k, m); else covered.add(k); } }
  for (const q of quadrants(g)) {
    ctx.save();
    ctx.beginPath(); ctx.rect(q.clip.x, q.clip.y, q.clip.w, q.clip.h); ctx.clip();
    // gridlines
    if (showGrid) {
      ctx.strokeStyle = '#dadce0'; ctx.lineWidth = g.lw;
      ctx.beginPath();
      for (let c = q.ca; c <= q.cb + 1; c++) { if (g.cols.hidden(c)) continue; const x = g.px(g.colX(c) - 0.5); ctx.moveTo(x, q.clip.y); ctx.lineTo(x, q.clip.y + q.clip.h); }
      for (let r = q.ra; r <= q.rb + 1; r++) { if (g.rows.hidden(r)) continue; const y = g.px(g.rowY(r) - 0.5); ctx.moveTo(q.clip.x, y); ctx.lineTo(q.clip.x + q.clip.w, y); }
      ctx.stroke();
    }
    const items = [];
    // merged cells intersecting this quadrant (anchors may be off screen)
    for (const m of merges) if (m.r1 <= q.rb && m.r2 >= q.ra && m.c1 <= q.cb && m.c2 >= q.ca) items.push({ r: m.r1, c: m.c1, m });
    for (let r = q.ra; r <= q.rb; r++) {
      if (g.rows.hidden(r)) continue;
      for (let c = q.ca; c <= q.cb; c++) {
        if (g.cols.hidden(c)) continue;
        const k = K(r, c);
        if (covered.has(k) || anchorOf.has(k)) continue;
        items.push({ r, c, m: null });
      }
    }
    // pass 1: fills (+ merged white background to hide gridlines)
    const prepared = [];
    for (const it of items) {
      const { r, c, m } = it;
      const k = K(r, c);
      const cell = sheet.cells.get(k);
      const v = cell ? cell.v : sheet.spill.has(k) ? sheet.spill.get(k).v : null;
      const s0 = cell && cell.s;
      const cf0 = cf ? cf(r, c) : null;
      const ts = tsAt ? tsAt(r, c) : null;
      const sp = spk ? spk.get(k) : null;
      if (!cell && v == null && !cf0 && !m && !ts && !sp) continue;
      // table style under the cell's own formatting
      const s = ts ? { ...ts, ...(s0 || {}), fill: (s0 && s0.fill) || ts.fill, color: (s0 && s0.color) || ts.color, b: (s0 && s0.b) || ts.b, bd: (s0 && s0.bd) || ts.bd } : s0;
      const cfs = cf0;
      const rc = m ? g.rect(m.r1, m.c1, m.r2, m.c2) : { x: g.colX(c), y: g.rowY(r), w: g.colW(c), h: g.rowH(r) };
      const fill = (cfs && cfs.fill) || (s && s.fill) || (m ? '#ffffff' : null);
      if (fill) { ctx.fillStyle = fill; ctx.fillRect(rc.x, rc.y, rc.w - (m && !s?.fill && showGrid ? 1 : 0), rc.h - (m && !s?.fill && showGrid ? 1 : 0)); }
      if (cfs && cfs.bar) { const bw = Math.max(0, (rc.w - 4) * cfs.bar.pct); const grd = ctx.createLinearGradient(rc.x + 2, 0, rc.x + 2 + Math.max(bw, 1), 0); grd.addColorStop(0, cfs.bar.color); grd.addColorStop(1, `${cfs.bar.color}33`); ctx.fillStyle = grd; ctx.fillRect(rc.x + 2, rc.y + 2, bw, rc.h - 4); }
      if (sp) drawSparkline(ctx, book, sp, rc, z);
      prepared.push({ r, c, k, cell, v, s, cfs, rc, m });
    }
    // pass 2: text
    for (const p of prepared) {
      const { cell, s, cfs, rc, r, c } = p;
      let v = p.v;
      if (v == null && !(showFormulas && cell && cell.f)) continue;
      if (cfs && cfs.icon) { const isz = Math.min(rc.h - 2, 16 * z); drawIcon(ctx, cfs.icon.set, cfs.icon.i, cfs.icon.n, rc.x + 2 * z, rc.y + (rc.h - isz) / 2, isz); if (cfs.icon.hide) continue; }
      if (cell && cell.rich && !(showFormulas && cell.f) && typeof v === 'string' && !(s && s.rot)) {
        // like plain text: overflow into empty cells on the right
        let w = rc.w; let cc = c + 1;
        while (!(s && s.wrap) && !p.m && cc <= Math.min(q.cb + 2, MAXC - 1)) { const x = sheet.cells.get(K(r, cc)); if ((x && x.v != null && x.v !== '') || sheet.spill.has(K(r, cc)) || covered.has(K(r, cc))) break; w += g.colW(cc); cc++; if (w > 2000) break; }
        drawRich(ctx, cell.rich, { ...rc, w }, { z, defFont, fontStack, h: (s && s.h) || 'left', va: (s && s.va) || 'bottom', pad: 3 * z, s }); continue;
      }
      let text; let color = null;
      if (showFormulas && cell && cell.f) text = `=${toDisplay(cell.f, loc)}`;
      else { const f = formatValue(v, (s && s.nf) || 'General', loc); text = f.text; color = f.color; }
      if (text === '') continue;
      const size = ((s && s.size) || defFont.size) * (4 / 3) * z;
      const font = `${(cfs && cfs.i) || (s && s.i) ? 'italic ' : ''}${(cfs && cfs.b) || (s && s.b) ? 'bold ' : ''}${Math.max(1, size).toFixed(1)}px ${fontStack((s && s.font) || defFont.font)}`;
      ctx.font = font;
      ctx.fillStyle = (cfs && cfs.color) || color || (s && s.color) || '#000000';
      const isNum = typeof v === 'number' && !(showFormulas && cell && cell.f);
      let h = (s && s.h) || 'general';
      if (h === 'general') h = isNum ? 'right' : typeof v === 'boolean' || isErr(v) ? 'center' : 'left';
      if (h === 'centerContinuous') h = 'center';
      const indent = ((s && s.indent) || 0) * 9 * z;
      const pad = 3 * z;
      const va = (s && s.va) || 'bottom';
      const wrap = s && (s.wrap || s.h === 'justify' || s.h === 'distributed');
      const rot = s && s.rot;
      let lines = String(text).split('\n');
      let tw = Math.max(...lines.map((l) => measure(ctx, font, l)));
      if (s && s.h === 'fill' && !wrap) { const one = measure(ctx, font, text) || 1; text = text.repeat(Math.max(1, Math.floor((rc.w - 2 * pad) / one))); lines = [text]; tw = measure(ctx, font, text); }
      if (wrap && !rot) {
        const maxW = rc.w - 2 * pad - indent;
        const out = [];
        for (const para of lines) {
          const words = para.split(/(\s+)/);
          let line = '';
          for (const w of words) {
            const t2 = line + w;
            if (measure(ctx, font, t2.trimEnd()) > maxW && line.trim()) { out.push(line.trimEnd()); line = w.trimStart(); } else line = t2;
            while (measure(ctx, font, line) > maxW && line.length > 1) { let n = line.length - 1; while (n > 1 && measure(ctx, font, line.slice(0, n)) > maxW) n--; out.push(line.slice(0, n)); line = line.slice(n); }
          }
          out.push(line);
        }
        lines = out;
        tw = Math.min(maxW, Math.max(...lines.map((l) => measure(ctx, font, l))));
      }
      const lh = size * 1.22;
      // numbers that don't fit become ####
      if (isNum && !wrap && !rot && tw > rc.w - 2 * pad && (s && s.nf && s.nf !== 'General')) { const hashW = measure(ctx, font, '#') || 6; lines = ['#'.repeat(Math.max(1, Math.floor((rc.w - 2 * pad) / hashW)))]; tw = measure(ctx, font, lines[0]); }
      // clip / overflow into empty neighbours
      let clipL = rc.x; let clipR = rc.x + rc.w;
      if (!wrap && !rot && !isNum && !p.m && tw + 2 * pad + indent > rc.w) {
        const empty = (cc) => { const k = K(r, cc); const x = sheet.cells.get(k); return !(x && (x.v != null && x.v !== '')) && !sheet.spill.has(k) && !covered.has(k) && !anchorOf.has(k); };
        if (h === 'left' || h === 'center') { let cc = c + 1; while (clipR - rc.x < tw + 2 * pad + indent + (h === 'center' ? (tw - rc.w) / 2 : 0) && cc <= Math.min(q.cb + 2, MAXC - 1) && empty(cc)) { clipR = g.colX(cc) + g.colW(cc); cc++; } }
        if (h === 'right' || h === 'center') { let cc = c - 1; while (rc.x + rc.w - clipL < tw + 2 * pad && cc >= Math.max(0, q.ca - 2) && empty(cc)) { clipL = g.colX(cc); cc--; } }
      }
      ctx.save();
      ctx.beginPath(); ctx.rect(clipL, rc.y, clipR - clipL, rc.h); ctx.clip();
      ctx.textBaseline = 'alphabetic';
      if (rot) {
        const deg = rot === 255 ? 0 : rot <= 90 ? rot : -(rot - 90);
        ctx.translate(rc.x + rc.w / 2, rc.y + rc.h / 2);
        if (rot === 255) {
          const chars = String(text).split('');
          ctx.textAlign = 'center';
          chars.forEach((ch, i) => ctx.fillText(ch, 0, (i - (chars.length - 1) / 2) * lh + size * 0.35));
        } else {
          ctx.rotate((-deg * Math.PI) / 180);
          ctx.textAlign = 'center';
          ctx.fillText(lines.join(' '), 0, size * 0.35);
        }
      } else {
        const blockH = lines.length * lh;
        let y0 = va === 'top' ? rc.y + pad + size * 0.95 : va === 'center' || va === 'distributed' || va === 'justify' ? rc.y + (rc.h - blockH) / 2 + size * 0.95 : rc.y + rc.h - blockH - pad * 0.6 + size * 0.95 + (lh - size) * 0.35;
        for (const line of lines) {
          const lw = measure(ctx, font, line);
          let x;
          if (h === 'right') { ctx.textAlign = 'right'; x = rc.x + rc.w - pad - indent; } else if (h === 'center') { ctx.textAlign = 'center'; x = rc.x + rc.w / 2; } else { ctx.textAlign = 'left'; x = rc.x + pad + indent; }
          ctx.fillText(line, x, y0);
          const u = (cfs && cfs.u) || (s && s.u); const stt = (cfs && cfs.st) || (s && s.st);
          if (u || stt || (cell && (cell.link || cell.link_) && !s?.color)) {
            const lx = h === 'right' ? x - lw : h === 'center' ? x - lw / 2 : x;
            ctx.fillRect(lx, y0 + Math.max(1, size * 0.12), lw, Math.max(1, z));
            if (u === 'double') ctx.fillRect(lx, y0 + Math.max(3, size * 0.22), lw, Math.max(1, z));
            if (stt) ctx.fillRect(lx, y0 - size * 0.3, lw, Math.max(1, z));
          }
          y0 += lh;
        }
      }
      ctx.restore();
    }
    // pass 3: borders + note markers
    for (const p of prepared) {
      const { s, rc, cell, cfs } = p;
      const bd = (cfs && cfs.bd) || (s && s.bd);
      if (bd) {
        const edge = (e, x1, y1, x2, y2) => {
          if (!e) return;
          const w = BW[e.s] || 1;
          ctx.strokeStyle = e.c || '#000000';
          if (e.s === 'double') { ctx.lineWidth = 1; ctx.setLineDash([]); const dx = y1 === y2 ? 0 : 1; const dy = y1 === y2 ? 1 : 0; ctx.beginPath(); ctx.moveTo(x1 - dx, y1 - dy); ctx.lineTo(x2 - dx, y2 - dy); ctx.moveTo(x1 + dx, y1 + dy); ctx.lineTo(x2 + dx, y2 + dy); ctx.stroke(); return; }
          ctx.lineWidth = w; ctx.setLineDash(DASH[e.s] || []);
          ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
        };
        const x0 = Math.round(rc.x) - 0.5; const y0 = Math.round(rc.y) - 0.5; const x1 = Math.round(rc.x + rc.w) - 0.5; const y1 = Math.round(rc.y + rc.h) - 0.5;
        edge(bd.t, x0, y0, x1, y0); edge(bd.b, x0, y1, x1, y1); edge(bd.l, x0, y0, x0, y1); edge(bd.r, x1, y0, x1, y1);
        if (bd.d) { if (bd.d.down !== false) edge(bd.d, x0, y0, x1, y1); if (bd.d.up) edge(bd.d, x0, y1, x1, y0); }
        ctx.setLineDash([]);
      }
      if (cell && cell.note) { ctx.fillStyle = '#d93025'; const x = rc.x + rc.w; ctx.beginPath(); ctx.moveTo(x - 6 * Math.min(1.5, z), rc.y); ctx.lineTo(x, rc.y); ctx.lineTo(x, rc.y + 6 * Math.min(1.5, z)); ctx.fill(); }
    }
    // filter buttons
    if (sheet.filter) {
      const f = sheet.filter.range;
      const r = f.r1;
      if (r >= q.ra && r <= q.rb) for (let c = Math.max(f.c1, q.ca); c <= Math.min(f.c2, q.cb); c++) {
        const bx = g.colX(c) + g.colW(c) - 17; const by = g.rowY(r) + g.rowH(r) - 17;
        const on = !!sheet.filter.cols[c];
        ctx.fillStyle = on ? '#e8f0fe' : '#ffffff'; ctx.strokeStyle = on ? '#1a73e8' : '#9aa0a6'; ctx.lineWidth = 1;
        ctx.fillRect(bx, by, 15, 15); ctx.strokeRect(bx + 0.5, by + 0.5, 14, 14);
        ctx.fillStyle = on ? '#1a73e8' : '#5f6368';
        ctx.beginPath();
        if (on) { ctx.moveTo(bx + 3, by + 4); ctx.lineTo(bx + 12, by + 4); ctx.lineTo(bx + 8.5, by + 8); ctx.lineTo(bx + 8.5, by + 12); ctx.lineTo(bx + 6.5, by + 11); ctx.lineTo(bx + 6.5, by + 8); } else { ctx.moveTo(bx + 4, by + 6); ctx.lineTo(bx + 11, by + 6); ctx.lineTo(bx + 7.5, by + 10); }
        ctx.fill();
      }
    }
    // selection, references, copy marquee, drag previews
    const acc = theme.accent || '#6a46ff';
    const strokeRange = (gg, color, width, dash, fillA) => {
      const r2 = Math.min(gg.r2, MAXR - 1); const c2 = Math.min(gg.c2, MAXC - 1);
      if (r2 < q.ra && gg.r1 < q.ra) { /* above */ }
      const rr = g.rect(gg.r1, gg.c1, r2, c2);
      if (fillA) { ctx.fillStyle = fillA; ctx.fillRect(rr.x, rr.y, rr.w, rr.h); }
      ctx.strokeStyle = color; ctx.lineWidth = width; ctx.setLineDash(dash || []);
      ctx.strokeRect(Math.round(rr.x) - 0.5 + (width > 1 ? 0.5 : 0), Math.round(rr.y) - 0.5 + (width > 1 ? 0.5 : 0), Math.round(rr.w) - (width > 1 ? 1 : 0), Math.round(rr.h) - (width > 1 ? 1 : 0));
      ctx.setLineDash([]);
    };
    if (refs) refs.forEach((rf) => { if (rf.sheet === sheet) strokeRange(rf, rf.color, 2, null, `${rf.color}1f`); });
    if (st.spillRange) strokeRange(st.spillRange, '#1a73e8', 1, [3, 2], null);
    if (sel && !(edit && refs && refs.length && st.pointing)) {
      const am = sheet.merges.find((m) => sel.r >= m.r1 && sel.r <= m.r2 && sel.c >= m.c1 && sel.c <= m.c2);
      const ar = am ? g.rect(am.r1, am.c1, am.r2, am.c2) : g.rect(sel.r, sel.c, sel.r, sel.c);
      sel.ranges.forEach((gg, i) => {
        const single = gg.r1 === gg.r2 && gg.c1 === gg.c2;
        if (!single || sel.ranges.length > 1) {
          const rr = g.rect(gg.r1, gg.c1, Math.min(gg.r2, MAXR - 1), Math.min(gg.c2, MAXC - 1));
          ctx.fillStyle = theme.selFill;
          const hasActive = sel.r >= gg.r1 && sel.r <= gg.r2 && sel.c >= gg.c1 && sel.c <= gg.c2 && i === sel.ranges.length - 1;
          if (!hasActive) ctx.fillRect(rr.x, rr.y, rr.w, rr.h);
          else {
            ctx.fillRect(rr.x, rr.y, rr.w, Math.max(0, ar.y - rr.y));
            ctx.fillRect(rr.x, ar.y + ar.h, rr.w, Math.max(0, rr.y + rr.h - ar.y - ar.h));
            ctx.fillRect(rr.x, ar.y, Math.max(0, ar.x - rr.x), ar.h);
            ctx.fillRect(ar.x + ar.w, ar.y, Math.max(0, rr.x + rr.w - ar.x - ar.w), ar.h);
          }
        }
        if (i === sel.ranges.length - 1) {
          if (st.selRect) { const rr = st.selRect; ctx.strokeStyle = acc; ctx.lineWidth = 2; ctx.strokeRect(Math.round(rr.x), Math.round(rr.y), Math.round(rr.w) - 1, Math.round(rr.h) - 1); } else strokeRange(gg, acc, 2, null, null);
        }
      });
      const last = sel.ranges[sel.ranges.length - 1];
      if (sel.ranges.length > 1) { ctx.strokeStyle = acc; ctx.lineWidth = 1; ctx.strokeRect(Math.round(ar.x) + 0.5, Math.round(ar.y) + 0.5, Math.round(ar.w) - 2, Math.round(ar.h) - 2); }
      // fill handle
      if (!readOnly && sel.ranges.length === 1 && !edit) {
        const rr = st.selRect || g.rect(last.r1, last.c1, Math.min(last.r2, MAXR - 1), Math.min(last.c2, MAXC - 1));
        ctx.fillStyle = '#ffffff'; ctx.fillRect(rr.x + rr.w - 4.5, rr.y + rr.h - 4.5, 8, 8);
        ctx.fillStyle = acc; ctx.fillRect(rr.x + rr.w - 3.5, rr.y + rr.h - 3.5, 6, 6);
      }
    }
    if (st.remote && st.remote.length) drawRemote(ctx, g, st.remote, sheet);
    if (copy && copy.sheetId === sheet.id) strokeRange(copy.g, acc, 2, [5, 3], null);
    if (drag && drag.preview) strokeRange(drag.preview, '#5f6368', 2, drag.kind === 'fill' ? [4, 2] : null, null);
    ctx.restore();
  }
  // frozen separators
  ctx.fillStyle = '#9aa0a6';
  if (g.fr) ctx.fillRect(g.hw, g.hh + g.fy - 1, W - g.hw, 1.5);
  if (g.fc) ctx.fillRect(g.hw + g.fx - 1, g.hh, 1.5, H - g.hh);
  if (g.hw || g.hh) paintHeaders(ctx, g, st);
  g.outlineBtns = paintOutline(ctx, g, sheet, theme);
}

function paintHeaders(ctx, g, st) {
  const { theme, sel, sheet } = st;
  const { W, H, hw, hh, z } = g;
  const fs = Math.round(11 * Math.min(1.4, Math.max(0.85, z)));
  ctx.font = `${fs}px ${FALLBACK}`;
  const inSel = (axis, i) => sel && sel.ranges.some((x) => (axis === 'c' ? i >= x.c1 && i <= x.c2 : i >= x.r1 && i <= x.r2));
  const full = (axis, i) => sel && sel.ranges.some((x) => (axis === 'c' ? i >= x.c1 && i <= x.c2 && x.r1 === 0 && x.r2 >= MAXR - 1 : i >= x.r1 && i <= x.r2 && x.c1 === 0 && x.c2 >= MAXC - 1));
  // column headers
  ctx.fillStyle = theme.head; ctx.fillRect(hw, 0, W - hw, hh);
  ctx.fillStyle = theme.head; ctx.fillRect(0, hh, hw, H - hh);
  const cRanges = [...(g.fc ? [[0, g.fc - 1, hw, hw + g.fx]] : []), [g.c0, g.c1, hw + g.fx, W]];
  for (const [a, b, x0, x1] of cRanges) {
    ctx.save(); ctx.beginPath(); ctx.rect(x0, 0, x1 - x0, hh); ctx.clip();
    for (let c = a; c <= b; c++) {
      if (g.cols.hidden(c)) { const x = g.colX(c); ctx.fillStyle = theme.line; ctx.fillRect(x - 1.5, 0, 2, hh); continue; }
      const x = g.colX(c); const w = g.colW(c);
      if (inSel('c', c)) { ctx.fillStyle = full('c', c) ? theme.headSelStrong : theme.headSel; ctx.fillRect(x, 0, w, hh); }
      ctx.fillStyle = inSel('c', c) ? theme.headSelText : theme.headText;
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      if (w > 8) ctx.fillText(st.r1c1 ? String(c + 1) : colName(c), x + w / 2, g.gy + (hh - g.gy) / 2 + 0.5);
      ctx.fillStyle = theme.line; ctx.fillRect(Math.round(x + w) - 1, 0, 1, hh);
    }
    ctx.restore();
  }
  const rRanges = [...(g.fr ? [[0, g.fr - 1, hh, hh + g.fy]] : []), [g.r0, g.r1, hh + g.fy, H]];
  for (const [a, b, y0, y1] of rRanges) {
    ctx.save(); ctx.beginPath(); ctx.rect(0, y0, hw, y1 - y0); ctx.clip();
    for (let r = a; r <= b; r++) {
      if (g.rows.hidden(r)) { const y = g.rowY(r); ctx.fillStyle = theme.line; ctx.fillRect(0, y - 1.5, hw, 2); continue; }
      const y = g.rowY(r); const h = g.rowH(r);
      if (inSel('r', r)) { ctx.fillStyle = full('r', r) ? theme.headSelStrong : theme.headSel; ctx.fillRect(0, y, hw, h); }
      ctx.fillStyle = inSel('r', r) ? theme.headSelText : sheet.filterHidden.has(r + 1) || (sheet.filter && sheet.filterHidden.size && r >= sheet.filter.range.r1 && r <= sheet.filter.range.r2 && r > sheet.filter.range.r1) ? '#1a73e8' : theme.headText;
      ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      if (h > 8) ctx.fillText(String(r + 1), hw - 6, y + h / 2 + 0.5);
      ctx.fillStyle = theme.line; ctx.fillRect(0, Math.round(y + h) - 1, hw, 1);
    }
    ctx.restore();
  }
  ctx.fillStyle = theme.line; ctx.fillRect(hw - 1, 0, 1, H); ctx.fillRect(0, hh - 1, W, 1);
  // corner
  ctx.fillStyle = theme.head; ctx.fillRect(0, 0, hw - 1, hh - 1);
  ctx.fillStyle = theme.headText; ctx.beginPath(); ctx.moveTo(hw - 4, hh - 4); ctx.lineTo(hw - 4, hh - 12); ctx.lineTo(hw - 12, hh - 4); ctx.fill();
}

/** width of the widest content of a column (auto-fit), unzoomed px */
export function autoFitWidth(ctx, book, sheet, c, defFont) {
  let best = 0;
  const b = book.bounds(sheet);
  for (let r = 0; r < Math.min(b.rows, 20000); r++) {
    const k = K(r, c);
    const cell = sheet.cells.get(k);
    const v = cell ? cell.v : sheet.spill.has(k) ? sheet.spill.get(k).v : null;
    if (v == null) continue;
    if (sheet.merges.some((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2 && m.c1 !== m.c2)) continue;
    const s = cell && cell.s;
    const t = formatValue(v, (s && s.nf) || 'General', book.loc).text;
    const font = `${s && s.i ? 'italic ' : ''}${s && s.b ? 'bold ' : ''}${(((s && s.size) || defFont.size) * 4) / 3}px ${fontStack((s && s.font) || defFont.font)}`;
    for (const line of String(t).split('\n')) best = Math.max(best, measure(ctx, font, line) + ((s && s.indent) || 0) * 9);
  }
  if (sheet.filter && sheet.filter.range.r1 >= 0 && c >= sheet.filter.range.c1 && c <= sheet.filter.range.c2) best += 18;
  return best ? Math.ceil(best + 10) : null;
}
export function autoFitHeight(ctx, book, sheet, r, defFont) {
  let best = 0;
  for (const [k, cell] of sheet.cells) {
    if (Math.floor(k / MAXC) !== r || cell.v == null) continue;
    const c = k % MAXC;
    const s = cell.s;
    const size = (((s && s.size) || defFont.size) * 4) / 3;
    let lines = String(formatValue(cell.v, (s && s.nf) || 'General', book.loc).text).split('\n').length;
    if (s && s.wrap) {
      const font = `${s.i ? 'italic ' : ''}${s.b ? 'bold ' : ''}${size}px ${fontStack(s.font || defFont.font)}`;
      const w = (sheet.colW.get(c) ?? sheet.defW) - 6;
      const tw = measure(ctx, font, String(cell.v));
      lines = Math.max(lines, Math.ceil(tw / Math.max(10, w)));
    }
    best = Math.max(best, lines * size * 1.22 + 6);
  }
  return best ? Math.ceil(Math.max(best, sheet.defH)) : null;
}
export { fontStack };
