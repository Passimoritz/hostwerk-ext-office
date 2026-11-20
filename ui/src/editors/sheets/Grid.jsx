/*
 * The virtualised grid: a native scroll area for the scrollbars and a canvas on top that paints only the visible cells
 * (see render.js). All mouse work happens here: selecting, dragging the fill handle, moving ranges, resizing and
 * auto-fitting rows/columns, picking references while a formula is typed, filter buttons, charts and images.
 */
import { useCallback, useEffect, useImperativeHandle, useLayoutEffect, useRef, useState, forwardRef } from 'react';
import { t } from '@hostwerk/i18n';
import { Keep } from '../_shared/keep.jsx';
import { MAXR, MAXC } from './formula/refs.js';
import { axesOf } from './layout.js';
import { makeGeo, paint, autoFitWidth, autoFitHeight } from './render.js';
import { ChartSVG } from './charts.jsx';

const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const norm = (a, b) => ({ r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) });

const reduced = () => typeof matchMedia !== 'undefined' && matchMedia('(prefers-reduced-motion: reduce)').matches;
function readTheme(el) {
  const cs = getComputedStyle(el);
  const v = (n, d) => (cs.getPropertyValue(n) || '').trim() || d;
  const accent = v('--of-sheet-sel', '') || v('--accent', '#6a46ff');
  return {
    accent, head: v('--surface-2', '#f7f8fb'), headText: v('--muted', '#5f677c'), line: v('--line', '#e3e6ee'),
    headSel: v('--accent-soft', 'rgba(106,70,255,.1)'), headSelStrong: `color-mix(in srgb, ${accent} 30%, transparent)`, headSelText: v('--accent-text', '#5a37f0'),
    selFill: 'rgba(26, 115, 232, 0.1)',
  };
}

export const Grid = forwardRef(function Grid(props, apiRef) {
  const { book, sheet, sel, onSel, zoom, readOnly, showFormulas, loc, defFont, edit, refs, pointing, copy, children, version } = props;
  const wrap = useRef(null); const scroller = useRef(null); const canvas = useRef(null);
  const [size, setSize] = useState({ w: 600, h: 400, sbw: 12, sbh: 12 });
  const [scroll, setScroll] = useState({ x: 0, y: 0 });
  const [theme, setTheme] = useState(null);
  const drag = useRef(null);
  const [dragView, setDragView] = useState(null);
  const geoRef = useRef(null);
  const raf = useRef(0);
  const z = zoom / 100;
  const axes = axesOf(sheet);
  const headings = sheet.headings !== false && props.headings !== false;

  // theme tokens (re-read when Hostwerk switches light/dark)
  useEffect(() => {
    if (!wrap.current) return undefined;
    setTheme(readTheme(wrap.current));
    const mo = new MutationObserver(() => setTheme(readTheme(wrap.current)));
    mo.observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme', 'class', 'style'] });
    return () => mo.disconnect();
  }, []);
  useLayoutEffect(() => {
    const el = scroller.current;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight, sbw: el.offsetWidth - el.clientWidth, sbh: el.offsetHeight - el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  // reset scroll when the sheet changes
  useEffect(() => { const el = scroller.current; if (el) { el.scrollTop = 0; el.scrollLeft = 0; setScroll({ x: 0, y: 0 }); } }, [sheet.id]);

  // extent: used area + room, grows while scrolling like Excel
  const b = book.bounds(sheet);
  const [extra, setExtra] = useState({ r: 0, c: 0 });
  const extR = clamp(Math.max(b.rows + 200, (sel ? sel.r : 0) + 60, 400, extra.r), 1, MAXR);
  const extC = clamp(Math.max(b.cols + 12, (sel ? sel.c : 0) + 12, 40, extra.c), 1, MAXC);
  const geo0 = { z };
  const hhEst = headings ? Math.round(22 * Math.min(1.4, Math.max(0.8, z))) : 0;
  const totalH = axes.rows.pos(extR) * z + hhEst + 40;
  const totalW = axes.cols.pos(extC) * z + (headings ? 60 : 0) + 40;
  void geo0;

  const geo = makeGeo({ sheet, axes, z, scrollX: scroll.x, scrollY: scroll.y, W: size.w, H: size.h, headings, dpr: window.devicePixelRatio || 1 });
  geoRef.current = geo;

  const onScroll = () => {
    const el = scroller.current;
    setScroll({ x: el.scrollLeft, y: el.scrollTop });
    if (el.scrollTop + el.clientHeight > el.scrollHeight - 400 && extR < MAXR) setExtra((e) => ({ ...e, r: Math.min(MAXR, extR + 500) }));
    if (el.scrollLeft + el.clientWidth > el.scrollWidth - 300 && extC < MAXC) setExtra((e) => ({ ...e, c: Math.min(MAXC, extC + 20) }));
  };

  // painting
  const draw = useCallback(() => {
    const cv = canvas.current; if (!cv || !theme) return;
    const g = geoRef.current;
    const dpr = window.devicePixelRatio || 1;
    if (cv.width !== Math.round(size.w * dpr) || cv.height !== Math.round(size.h * dpr)) { cv.width = Math.round(size.w * dpr); cv.height = Math.round(size.h * dpr); cv.style.width = `${size.w}px`; cv.style.height = `${size.h}px`; }
    const ctx = cv.getContext('2d');
    let spillRange = null;
    if (sel) {
      const k = sel.r * MAXC + sel.c;
      const sp = sheet.spill.get(k); const anchor = sp ? sp.anchor : sheet.spillAnchors.has(k) ? k : null;
      if (anchor != null) { const d = sheet.spillAnchors.get(anchor); if (d) { const r0 = Math.floor(anchor / MAXC); const c0 = anchor % MAXC; spillRange = { r1: r0, c1: c0, r2: r0 + d.rows - 1, c2: c0 + d.cols - 1 }; } }
    }
    paint(ctx, g, { book, sheet, theme, sel, edit, refs, pointing, copy, drag: dragView, showFormulas, loc, defFont, readOnly, spillRange, remote: props.remote, r1c1: props.r1c1, selRect: selRect.current });
  }, [theme, size, sel, edit, refs, pointing, copy, dragView, showFormulas, loc, defFont, readOnly, book, sheet, scroll, zoom, version, props.remote, props.r1c1]); // eslint-disable-line react-hooks/exhaustive-deps
  // paint in the same frame as React's commit (scrolling stays in step with the scrollbars)
  useLayoutEffect(() => { cancelAnimationFrame(raf.current); draw(); }, [draw]);
  // the selection glides to its new place (keyboard moves; not while dragging, not with reduced motion)
  const selRect = useRef(null); const lastSel = useRef(null); const anim = useRef(0);
  useEffect(() => {
    const g = geoRef.current; if (!sel || !g) return;
    const last = sel.ranges[sel.ranges.length - 1];
    const prev = lastSel.current; lastSel.current = { ...last, sheet: sheet.id };
    if (!prev || prev.sheet !== sheet.id || drag.current || reduced() || sel.ranges.length > 1) return;
    if (Math.abs(prev.r1 - last.r1) + Math.abs(prev.c1 - last.c1) > 12 || (prev.r1 === last.r1 && prev.c1 === last.c1 && prev.r2 === last.r2 && prev.c2 === last.c2)) return;
    if (last.r2 - last.r1 > 50 || last.c2 - last.c1 > 50) return;
    const from = g.rect(prev.r1, prev.c1, prev.r2, prev.c2); const to = g.rect(last.r1, last.c1, last.r2, last.c2);
    const t0 = performance.now(); const D = 110;
    cancelAnimationFrame(anim.current);
    const tick = (now) => {
      const k = Math.min(1, (now - t0) / D); const e = 1 - (1 - k) ** 3;
      selRect.current = k >= 1 ? null : { x: from.x + (to.x - from.x) * e, y: from.y + (to.y - from.y) * e, w: from.w + (to.w - from.w) * e, h: from.h + (to.h - from.h) * e };
      draw();
      if (k < 1) anim.current = requestAnimationFrame(tick);
    };
    anim.current = requestAnimationFrame(tick);
  }, [sel]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- api for the editor ---------- */
  const cellRect = (r, c) => {
    const g = geoRef.current;
    const m = sheet.merges.find((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2);
    return m ? g.rect(m.r1, m.c1, m.r2, m.c2) : g.rect(r, c, r, c);
  };
  const scrollTo = (r, c) => {
    const el = scroller.current; const g = geoRef.current; if (!el || !g) return;
    if (r >= g.fr) {
      const top = axes.rows.pos(r) * z - axes.rows.pos(g.fr) * z; const bot = top + axes.rows.size(r) * z;
      const viewH = size.h - g.hh - g.fy;
      if (top < el.scrollTop) el.scrollTop = top; else if (bot > el.scrollTop + viewH) el.scrollTop = bot - viewH + 2;
    }
    if (c >= g.fc) {
      const left = axes.cols.pos(c) * z - axes.cols.pos(g.fc) * z; const right = left + axes.cols.size(c) * z;
      const viewW = size.w - g.hw - g.fx;
      if (left < el.scrollLeft) el.scrollLeft = left; else if (right > el.scrollLeft + viewW) el.scrollLeft = Math.min(right - viewW + 2, left);
    }
  };
  const pageRows = () => { const g = geoRef.current; return Math.max(1, g.r1 - g.r0 - 1); };
  useImperativeHandle(apiRef, () => ({ cellRect, scrollTo, pageRows, geo: () => geoRef.current, canvas: () => canvas.current, autoFit }), [sheet, size, z, scroll]); // eslint-disable-line react-hooks/exhaustive-deps

  function autoFit(axis, idxs) {
    const ctx = canvas.current.getContext('2d');
    if (axis === 'col') for (const c of idxs) { const w = autoFitWidth(ctx, book, sheet, c, defFont); props.onResize('col', [c], w || sheet.defW); } else for (const r of idxs) { const h = autoFitHeight(ctx, book, sheet, r, defFont); props.onResize('row', [r], h); }
  }

  /* ---------- hit testing ---------- */
  const local = (e) => { const r = canvas.current.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; };
  const hit = (x, y) => {
    const g = geoRef.current;
    if ((g.gx && x < g.gx) || (g.gy && y < g.gy)) return { area: 'gutter' };
    if (x < g.hw && y < g.hh) return { area: 'corner' };
    if (y < g.hh) {
      const c = g.colAt(x);
      const cx = g.colX(c); const right = cx + g.colW(c);
      if (Math.abs(x - right) <= 4) return { area: 'colhead', c, resize: true };
      if (Math.abs(x - cx) <= 3 && c > 0) { let p = c - 1; while (p > 0 && axes.cols.hidden(p)) p--; return { area: 'colhead', c: p, resize: true }; }
      return { area: 'colhead', c };
    }
    if (x < g.hw) {
      const r = g.rowAt(y);
      const ry = g.rowY(r); const bot = ry + g.rowH(r);
      if (Math.abs(y - bot) <= 3) return { area: 'rowhead', r, resize: true };
      if (Math.abs(y - ry) <= 2 && r > 0) { let p = r - 1; while (p > 0 && axes.rows.hidden(p)) p--; return { area: 'rowhead', r: p, resize: true }; }
      return { area: 'rowhead', r };
    }
    const r = g.rowAt(y); const c = g.colAt(x);
    const out = { area: 'cell', r, c };
    if (sheet.filter && r === sheet.filter.range.r1 && c >= sheet.filter.range.c1 && c <= sheet.filter.range.c2) {
      const bx = g.colX(c) + g.colW(c) - 17; const by = g.rowY(r) + g.rowH(r) - 17;
      if (x >= bx && x <= bx + 15 && y >= by && y <= by + 15) out.filterBtn = true;
    }
    if (sel && sel.ranges.length === 1 && !readOnly && !edit) {
      const last = sel.ranges[0];
      const rr = g.rect(last.r1, last.c1, Math.min(last.r2, MAXR - 1), Math.min(last.c2, MAXC - 1));
      if (Math.abs(x - (rr.x + rr.w)) <= 5 && Math.abs(y - (rr.y + rr.h)) <= 5) out.fill = true;
      else if (x >= rr.x - 3 && x <= rr.x + rr.w + 3 && y >= rr.y - 3 && y <= rr.y + rr.h + 3 && (Math.abs(x - rr.x) <= 3 || Math.abs(x - rr.x - rr.w) <= 3 || Math.abs(y - rr.y) <= 3 || Math.abs(y - rr.y - rr.h) <= 3)) out.border = true;
    }
    return out;
  };
  const cellAtPoint = (x, y) => { const g = geoRef.current; return { r: g.rowAt(clamp(y, g.hh, size.h - 1)), c: g.colAt(clamp(x, g.hw, size.w - 1)) }; };

  /* ---------- pointer ---------- */
  const auto = useRef(null);
  const stopAuto = () => { if (auto.current) { clearInterval(auto.current); auto.current = null; } };
  useEffect(() => stopAuto, []);
  const startDrag = (d, e) => {
    drag.current = d;
    const move = (ev) => {
      const p = local(ev);
      d.last = p;
      const el = scroller.current;
      const g = geoRef.current;
      const dx = p.x > size.w - 10 ? 20 : p.x < g.hw + (d.kind === 'col' ? 0 : 0) && d.kind !== 'row' && d.kind !== 'rowsize' ? -20 : 0;
      const dy = p.y > size.h - 10 ? 20 : p.y < g.hh && d.kind !== 'col' && d.kind !== 'colsize' ? -20 : 0;
      stopAuto();
      if ((dx || dy) && !d.kind.endsWith('size')) auto.current = setInterval(() => { el.scrollLeft += dx; el.scrollTop += dy; d.onMove(d.last); }, 40);
      d.onMove(p, ev);
    };
    const up = (ev) => { stopAuto(); window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); try { if (d.onUp) d.onUp(local(ev), ev); } finally { drag.current = null; setDragView(null); } };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    void e;
  };

  const onPointerDown = (e) => {
    if (e.button === 1) return;
    const p = local(e);
    const h = hit(p.x, p.y);
    if (props.onBlankClick) props.onBlankClick();
    if (h.area === 'gutter') { const ob = (geoRef.current.outlineBtns || []).find((b) => p.x >= b.x && p.x <= b.x + b.w && p.y >= b.y && p.y <= b.y + b.h); e.preventDefault(); if (ob && props.onOutline && !readOnly) props.onOutline(ob); return; }
    e.preventDefault(); // keep the keyboard focus in the cell input
    if (props.focus) { props.focus(); setTimeout(() => props.focus && props.focus(), 0); }
    if (e.button === 2) {
      e.preventDefault();
      if (h.area === 'cell' && !(sel && sel.ranges.some((g) => h.r >= g.r1 && h.r <= g.r2 && h.c >= g.c1 && h.c <= g.c2))) onSel({ r: h.r, c: h.c, ranges: [{ r1: h.r, c1: h.c, r2: h.r, c2: h.c }], anchor: { r: h.r, c: h.c } });
      if (h.area === 'colhead' && !(sel && sel.ranges.some((g) => g.r1 === 0 && g.r2 >= MAXR - 1 && h.c >= g.c1 && h.c <= g.c2))) onSel({ r: 0, c: h.c, ranges: [{ r1: 0, c1: h.c, r2: MAXR - 1, c2: h.c }], anchor: { r: 0, c: h.c } });
      if (h.area === 'rowhead' && !(sel && sel.ranges.some((g) => g.c1 === 0 && g.c2 >= MAXC - 1 && h.r >= g.r1 && h.r <= g.r2))) onSel({ r: h.r, c: 0, ranges: [{ r1: h.r, c1: 0, r2: h.r, c2: MAXC - 1 }], anchor: { r: h.r, c: 0 } });
      props.onContextMenu && props.onContextMenu({ x: e.clientX, y: e.clientY }, h);
      return;
    }
    const add = e.ctrlKey || e.metaKey;
    const ext = e.shiftKey;
    if (h.area === 'corner') { onSel({ r: sel ? sel.r : 0, c: sel ? sel.c : 0, ranges: [{ r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 }], anchor: { r: 0, c: 0 } }); return; }
    if ((h.area === 'colhead' || h.area === 'rowhead') && h.resize) {
      if (readOnly) return;
      const isCol = h.area === 'colhead';
      const idx = isCol ? h.c : h.r;
      if (e.detail >= 2) { // double click: auto-fit (all selected columns/rows if the clicked one is selected)
        const selected = sel ? sel.ranges.filter((g) => (isCol ? g.r1 === 0 && g.r2 >= MAXR - 1 && idx >= g.c1 && idx <= g.c2 : g.c1 === 0 && g.c2 >= MAXC - 1 && idx >= g.r1 && idx <= g.r2)) : [];
        const list = [];
        if (selected.length) for (const g of selected) for (let i = isCol ? g.c1 : g.r1; i <= (isCol ? g.c2 : g.r2) && list.length < 500; i++) list.push(i); else list.push(idx);
        autoFit(isCol ? 'col' : 'row', list);
        return;
      }
      const start = isCol ? axes.cols.size(idx) : axes.rows.size(idx);
      const origin = isCol ? p.x : p.y;
      startDrag({ kind: isCol ? 'colsize' : 'rowsize', onMove: (q) => { const d = ((isCol ? q.x : q.y) - origin) / z; setDragView({ size: { axis: isCol ? 'col' : 'row', idx, px: Math.max(isCol ? 4 : 4, start + d) } }); }, onUp: (q) => { const d = ((isCol ? q.x : q.y) - origin) / z; if (Math.abs(d) < 1) return; const px = Math.max(2, Math.round(start + d)); const selected = sel ? sel.ranges.filter((g) => (isCol ? g.r1 === 0 && g.r2 >= MAXR - 1 && idx >= g.c1 && idx <= g.c2 : g.c1 === 0 && g.c2 >= MAXC - 1 && idx >= g.r1 && idx <= g.r2)) : []; const list = []; if (selected.length) for (const g of selected) for (let i = isCol ? g.c1 : g.r1; i <= (isCol ? g.c2 : g.r2) && list.length < 16384; i++) list.push(i); else list.push(idx); props.onResize(isCol ? 'col' : 'row', list, px); } }, e);
      return;
    }
    if (h.area === 'colhead' || h.area === 'rowhead') {
      const isCol = h.area === 'colhead';
      const anchorIdx = ext && sel ? (isCol ? sel.anchor ? sel.anchor.c : sel.c : sel.anchor ? sel.anchor.r : sel.r) : isCol ? h.c : h.r;
      const mk = (i) => (isCol ? { r1: 0, r2: MAXR - 1, c1: Math.min(anchorIdx, i), c2: Math.max(anchorIdx, i) } : { c1: 0, c2: MAXC - 1, r1: Math.min(anchorIdx, i), r2: Math.max(anchorIdx, i) });
      const base = add && sel ? sel.ranges : [];
      if (pointing && props.onPointRef) { props.onPointRef(mk(isCol ? h.c : h.r), 'start'); startDrag({ kind: isCol ? 'col' : 'row', onMove: (q) => { const c = cellAtPoint(q.x, q.y); props.onPointRef(mk(isCol ? c.c : c.r), 'move'); } }, e); return; }
      const apply = (i) => onSel({ r: isCol ? (sel && ext ? sel.r : 0) : i === anchorIdx ? anchorIdx : Math.min(anchorIdx, i), c: isCol ? (i === anchorIdx ? anchorIdx : Math.min(anchorIdx, i)) : sel && ext ? sel.c : 0, ranges: [...base, mk(i)], anchor: isCol ? { r: 0, c: anchorIdx } : { r: anchorIdx, c: 0 } });
      const first = isCol ? h.c : h.r;
      onSel({ r: isCol ? 0 : (ext ? (sel ? sel.r : first) : first), c: isCol ? (ext ? (sel ? sel.c : first) : first) : 0, ranges: [...base, mk(first)], anchor: isCol ? { r: 0, c: anchorIdx } : { r: anchorIdx, c: 0 } });
      startDrag({ kind: isCol ? 'col' : 'row', onMove: (q) => { const c = cellAtPoint(q.x, q.y); apply(isCol ? c.c : c.r); } }, e);
      return;
    }
    // cells
    if (h.filterBtn) { const g = geoRef.current; const r = canvas.current.getBoundingClientRect(); props.onFilterButton && props.onFilterButton(h.c, { left: r.left + g.colX(h.c) + g.colW(h.c) - 17, below: r.top + g.rowY(h.r) + g.rowH(h.r) + 2 }); return; }
    if (pointing && props.onPointRef) {
      const a = { r: h.r, c: h.c };
      props.onPointRef(norm(a, a), ext ? 'extend' : add ? 'add' : 'start');
      startDrag({ kind: 'point', onMove: (q) => { const c = cellAtPoint(q.x, q.y); props.onPointRef(norm(a, c), 'move'); } }, e);
      return;
    }
    if (h.fill && sel) {
      const src = sel.ranges[0];
      startDrag({ kind: 'fill', onMove: (q) => {
        const c = cellAtPoint(q.x, q.y);
        let dst = { ...src };
        const dr = c.r > src.r2 ? c.r - src.r2 : c.r < src.r1 ? c.r - src.r1 : 0;
        const dc = c.c > src.c2 ? c.c - src.c2 : c.c < src.c1 ? c.c - src.c1 : 0;
        if (Math.abs(dr) >= Math.abs(dc) && dr) dst = dr > 0 ? { ...src, r2: c.r } : { ...src, r1: c.r };
        else if (dc) dst = dc > 0 ? { ...src, c2: c.c } : { ...src, c1: c.c };
        drag.current.dst = dst;
        setDragView({ kind: 'fill', preview: dst, tip: props.fillPreview ? props.fillPreview(src, dst) : null, at: q });
      }, onUp: (q, ev) => { const dst = drag.current && drag.current.dst; void q; if (dst && (dst.r1 !== src.r1 || dst.r2 !== src.r2 || dst.c1 !== src.c1 || dst.c2 !== src.c2)) props.onFill(src, dst, ev.ctrlKey || ev.metaKey); } }, e);
      return;
    }
    if (h.border && sel) {
      const src = sel.ranges[0];
      const start = { r: h.r, c: h.c };
      startDrag({ kind: 'move', onMove: (q) => {
        const c = cellAtPoint(q.x, q.y);
        const dr = c.r - start.r; const dc = c.c - start.c;
        const r1 = clamp(src.r1 + dr, 0, MAXR - 1 - (src.r2 - src.r1)); const c1 = clamp(src.c1 + dc, 0, MAXC - 1 - (src.c2 - src.c1));
        const pv = { r1, c1, r2: r1 + src.r2 - src.r1, c2: c1 + src.c2 - src.c1 };
        drag.current.dst = pv;
        setDragView({ kind: 'move', preview: pv });
      }, onUp: (q, ev) => { const pv = drag.current && drag.current.dst; void q; if (pv && (pv.r1 !== src.r1 || pv.c1 !== src.c1)) props.onMove(src, pv.r1, pv.c1, ev.ctrlKey || ev.metaKey); } }, e);
      return;
    }
    if (e.detail === 2) { props.onEditStart && props.onEditStart(h.r, h.c); return; }
    props.onCellDown && props.onCellDown(h, e);
    const anchor = ext && sel ? sel.anchor || { r: sel.r, c: sel.c } : { r: h.r, c: h.c };
    const base = add && sel ? sel.ranges : [];
    const expand = (g) => props.expandMerges ? props.expandMerges(g) : g;
    const set = (c) => onSel({ r: ext && sel ? sel.r : anchor.r, c: ext && sel ? sel.c : anchor.c, ranges: [...base, expand(norm(anchor, c))], anchor });
    set({ r: h.r, c: h.c });
    startDrag({ kind: 'cells', onMove: (q) => set(cellAtPoint(q.x, q.y)) }, e);
  };

  const [cursor, setCursor] = useState('cell');
  const [hover, setHover] = useState(null);
  const onPointerMove = (e) => {
    if (drag.current) return;
    const p = local(e);
    const h = hit(p.x, p.y);
    setCursor(h.area === 'gutter' ? ((geoRef.current.outlineBtns || []).some((b) => p.x >= b.x && p.x <= b.x + b.w && p.y >= b.y && p.y <= b.y + b.h) ? 'pointer' : 'default') : h.area === 'corner' ? 'default' : h.resize ? (h.area === 'colhead' ? 'col-resize' : 'row-resize') : h.area === 'colhead' ? 's-resize' : h.area === 'rowhead' ? 'e-resize' : h.fill ? 'crosshair' : h.border ? 'move' : h.filterBtn ? 'pointer' : 'cell');
    if (h.area === 'cell') {
      const cell = sheet.cells.get(h.r * MAXC + h.c);
      if (cell && (cell.note || ((cell.link || cell.link_) && false))) { if (!hover || hover.r !== h.r || hover.c !== h.c) setHover({ r: h.r, c: h.c }); } else if (hover) setHover(null);
    } else if (hover) setHover(null);
  };
  // mouse wheels move in notches: glide there (touchpads already send smooth, momentum deltas – applied directly)
  const glide = useRef({ x: 0, y: 0, raf: 0 });
  const onWheel = (e) => {
    if (e.ctrlKey) { e.preventDefault(); props.onZoom && props.onZoom(e.deltaY < 0 ? 10 : -10); return; }
    const el = scroller.current;
    const dy = e.deltaMode === 1 ? e.deltaY * 20 : e.deltaMode === 2 ? e.deltaY * el.clientHeight : e.deltaY;
    const dx = e.deltaMode === 1 ? e.deltaX * 20 : e.deltaX || (e.shiftKey ? dy : 0);
    const notch = e.deltaMode !== 0 || (Math.abs(e.deltaY) >= 50 && Number.isInteger(e.deltaY) && !e.deltaX);
    if (!notch || reduced()) { el.scrollTop += e.shiftKey ? 0 : dy; el.scrollLeft += dx; return; }
    const gl = glide.current;
    gl.y += e.shiftKey ? 0 : dy; gl.x += dx;
    if (gl.raf) return;
    const step = () => {
      const sy = Math.abs(gl.y) < 1 ? gl.y : gl.y * 0.28; const sx = Math.abs(gl.x) < 1 ? gl.x : gl.x * 0.28;
      el.scrollTop += sy; el.scrollLeft += sx; gl.y -= sy; gl.x -= sx;
      gl.raf = Math.abs(gl.y) > 0.5 || Math.abs(gl.x) > 0.5 ? requestAnimationFrame(step) : 0;
      if (!gl.raf) { gl.x = 0; gl.y = 0; }
    };
    gl.raf = requestAnimationFrame(step);
  };
  useEffect(() => { const cv = canvas.current; const h = (e) => onWheel(e); cv.addEventListener('wheel', h, { passive: false }); return () => cv.removeEventListener('wheel', h); });

  /* ---------- overlays: charts, images, note tooltip, resize guide ---------- */
  const g = geo;
  const overlayStyle = { left: 0, top: 0, width: size.w, height: size.h };
  const objRect = (a) => {
    const x = g.colX(a.c) + (a.dx || 0) * z; const y = g.rowY(a.r) + (a.dy || 0) * z;
    return { left: x, top: y, width: a.w * z, height: a.h * z };
  };
  const noteCell = hover && sheet.cells.get(hover.r * MAXC + hover.c);
  const noteRect = hover && noteCell ? cellRect(hover.r, hover.c) : null;

  return (
    <div className="of-sheet-gridwrap" ref={wrap}>
      <div className="of-sheet-scroller" ref={scroller} onScroll={onScroll}>
        <div style={{ width: totalW, height: totalH }} />
      </div>
      <canvas ref={canvas} className="of-sheet-canvas" style={{ cursor }} onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerLeave={() => setHover(null)} onContextMenu={(e) => e.preventDefault()}
        role="grid" aria-label={t('Spreadsheet grid')} aria-rowcount={MAXR} aria-colcount={MAXC} />
      <div className="of-sheet-overlay" style={overlayStyle}>
        <div className="of-sheet-objects" style={{ left: g.hw + g.fx, top: g.hh + g.fy, right: 0, bottom: 0 }}>
          <div style={{ position: 'absolute', left: -(g.hw + g.fx), top: -(g.hh + g.fy) }}>
            {props.objects ? props.objects(objRect) : sheet.images.map((im) => { const rc = objRect(im.anchor); return <img key={im.id} src={im.src} alt="" className="of-sheet-image" style={rc} draggable={false} />; })}
            {sheet.charts.map((ch) => <ChartBox key={ch.id} chart={ch} rect={objRect(ch.anchor)} z={z} selected={props.chartSel === ch.id} readOnly={readOnly} data={props.chartData(ch)}
              onSelect={() => props.onChartSel(ch.id)} onChange={(anchor) => props.onChartChange(ch.id, anchor)} onOpen={() => props.onChartOpen(ch.id)} onMenu={(at) => props.onChartMenu(ch.id, at)} />)}
          </div>
        </div>
        {dragView && dragView.size && (dragView.size.axis === 'col'
          ? <div className="of-sheet-guide v" style={{ left: g.colX(dragView.size.idx) + dragView.size.px * z }} />
          : <div className="of-sheet-guide h" style={{ top: g.rowY(dragView.size.idx) + dragView.size.px * z }} />)}
        {dragView && dragView.size && <div className="of-sheet-tip small" style={{ left: dragView.size.axis === 'col' ? g.colX(dragView.size.idx) + dragView.size.px * z + 6 : 4, top: dragView.size.axis === 'col' ? 2 : g.rowY(dragView.size.idx) + dragView.size.px * z + 4 }}>{dragView.size.axis === 'col' ? t('Width: {px} px', { px: Math.round(dragView.size.px) }) : t('Height: {px} px', { px: Math.round(dragView.size.px) })}</div>}
        {dragView && dragView.tip && dragView.at && <div className="of-sheet-tip small is-fill" style={{ left: dragView.at.x + 14, top: dragView.at.y + 16 }}><Keep>{dragView.tip}</Keep></div>}
        {noteRect && noteCell && noteCell.note && <div className="of-sheet-note" style={{ left: noteRect.x + noteRect.w + 6, top: noteRect.y }}>{noteCell.noteAuthor && <b className="small"><Keep>{noteCell.noteAuthor}</Keep></b>}<div className="small"><Keep>{noteCell.note}</Keep></div></div>}
        {children && children(g, cellRect)}
      </div>
    </div>
  );
});

function ChartBox({ chart, rect, z, selected, readOnly, data, onSelect, onChange, onOpen, onMenu }) {
  const [live, setLive] = useState(null);
  const r = live || rect;
  const begin = (e, mode) => {
    e.stopPropagation(); e.preventDefault();
    onSelect();
    if (readOnly) return;
    const sx = e.clientX; const sy = e.clientY; const a0 = { ...chart.anchor };
    const move = (ev) => {
      const dx = (ev.clientX - sx) / z; const dy = (ev.clientY - sy) / z;
      if (mode === 'move') setLive({ ...rect, left: rect.left + dx * z, top: rect.top + dy * z });
      else setLive({ ...rect, width: Math.max(80, rect.width + (mode.includes('e') ? dx * z : 0)), height: Math.max(60, rect.height + (mode.includes('s') ? dy * z : 0)) });
      move.d = { dx, dy };
    };
    const up = () => {
      window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
      setLive(null);
      const d = move.d; if (!d) return;
      if (mode === 'move') onChange({ ...a0, dx: (a0.dx || 0) + d.dx, dy: (a0.dy || 0) + d.dy, _px: true });
      else onChange({ ...a0, w: Math.max(80, a0.w + (mode.includes('e') ? d.dx : 0)), h: Math.max(60, a0.h + (mode.includes('s') ? d.dy : 0)) });
    };
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
  };
  return (
    <div className={`of-sheet-chart${selected ? ' is-sel' : ''}`} style={r} onPointerDown={(e) => begin(e, 'move')} onDoubleClick={onOpen}
      onContextMenu={(e) => { e.preventDefault(); onSelect(); onMenu({ x: e.clientX, y: e.clientY }); }} tabIndex={-1}>
      <ChartSVG chart={chart} data={data} width={Math.max(10, r.width)} height={Math.max(10, r.height)} />
      {selected && !readOnly && ['e', 's', 'se'].map((m) => <span key={m} className={`of-sheet-handle h-${m}`} onPointerDown={(e) => begin(e, m)} />)}
    </div>
  );
}
