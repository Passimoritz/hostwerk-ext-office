/*
 * The editing canvas: one slide, selection, move/resize/rotate with smart guides and snapping, marquee, drawing new
 * shapes, in-place text editing (contentEditable, see render/edit.js), table cell editing.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, createElement as h } from 'react';
import { t } from '@hostwerk/i18n';
import SlideView from './render/SlideView.jsx';
import { bodyToHtml, htmlToBody, refreshParas, saveRange, restoreRange, execRun, setParaProps, caretRunProps } from './render/edit.js';
import { S, EMU_PER_PT } from './units.js';
import { findEl, rotatedBox, bbox, moveEl, fitEl, updateEls, walk, updateSlide } from './model.js';
import { pathSvg, pathPoints } from './anim.js';
import { displaySrc } from './render/SlideView.jsx';

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const toEmu = (px) => px * EMU_PER_PT;

/* ---------------------------------------------------------------- text editor (contentEditable) */
export function TextEditor({ body, theme, w, h: H, ctx, onChange, onExit, apiRef, autoGrow, selectAll }) {
  const ref = useRef(null);
  const range = useRef(null);
  const timer = useRef(null);
  const baseRef = useRef(body);
  const changeRef = useRef(onChange); changeRef.current = onChange;
  const flush = useCallback(() => {
    clearTimeout(timer.current);
    if (!ref.current) return;
    const nb = refreshParas(ref.current, baseRef.current, theme);
    changeRef.current(nb, autoGrow ? ref.current.scrollHeight : null);
  }, [theme, autoGrow]);
  useLayoutEffect(() => {
    const el = ref.current;
    el.innerHTML = bodyToHtml(body, theme, ctx);
    el.focus({ preventScroll: true });
    try { document.execCommand('defaultParagraphSeparator', false, 'div'); } catch { /* ignore */ }
    const sel = window.getSelection();
    const r = document.createRange();
    r.selectNodeContents(el);
    if (!selectAll) r.collapse(false);
    sel.removeAllRanges(); sel.addRange(r);
    const save = () => { const s = saveRange(el); if (s) range.current = s; };
    document.addEventListener('selectionchange', save);
    return () => { document.removeEventListener('selectionchange', save); flush(); };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const schedule = () => { clearTimeout(timer.current); timer.current = setTimeout(flush, 250); };
  if (apiRef) {
    apiRef.current = {
      run: (cmd, value) => { restoreRange(ref.current, range.current); execRun(ref.current, cmd, value, theme); flush(); },
      para: (fn) => { restoreRange(ref.current, range.current); setParaProps(ref.current, fn, theme, baseRef.current); flush(); },
      caret: () => (ref.current ? caretRunProps(ref.current, baseRef.current, theme) : null),
      insertText: (s) => { restoreRange(ref.current, range.current); document.execCommand('insertText', false, s); flush(); },
      flush,
      el: () => ref.current,
    };
  }
  const ins = body.insets || {};
  const style = {
    position: 'absolute', left: 0, top: 0, width: w, minHeight: autoGrow ? undefined : H, height: autoGrow ? 'auto' : H, boxSizing: 'border-box',
    padding: `${S(ins.t ?? 45720)}px ${S(ins.r ?? 91440)}px ${S(ins.b ?? 45720)}px ${S(ins.l ?? 91440)}px`,
    display: 'flex', flexDirection: 'column', justifyContent: body.anchor === 'ctr' ? 'center' : body.anchor === 'b' ? 'flex-end' : 'flex-start',
    whiteSpace: body.wrap === false ? 'pre' : 'pre-wrap', overflowWrap: 'break-word', outline: 'none', cursor: 'text',
  };
  return h('div', {
    ref, className: 'of-slides-edit', contentEditable: true, suppressContentEditableWarning: true, spellCheck: ctx.spellcheck !== false, style,
    onInput: schedule,
    onPointerDown: (e) => e.stopPropagation(),
    onDoubleClick: (e) => e.stopPropagation(),
    onPaste: (e) => {
      e.preventDefault();
      const s = e.clipboardData.getData('text/plain');
      if (s) document.execCommand('insertText', false, s.replace(/\r\n?/g, '\n'));
    },
    onKeyDown: (e) => {
      e.stopPropagation();
      const k = e.key.toLowerCase();
      if (e.key === 'Escape') { e.preventDefault(); flush(); onExit('esc'); return; }
      if (e.key === 'Tab' && !ctx.tableCell) { e.preventDefault(); setParaProps(ref.current, (p) => ({ level: Math.max(0, Math.min(8, (p.level || 0) + (e.shiftKey ? -1 : 1))), marL: undefined, indent: undefined }), theme, baseRef.current); flush(); return; }
      if (e.key === 'Tab' && ctx.tableCell) { e.preventDefault(); flush(); onExit(e.shiftKey ? 'prevCell' : 'nextCell'); return; }
      if ((e.ctrlKey || e.metaKey) && !e.altKey) {
        const map = { b: 'b', i: 'i', u: 'u' };
        if (map[k]) { e.preventDefault(); execRun(ref.current, map[k], null, theme); flush(); return; }
        if (k === 'e' || k === 'l' || k === 'r' || k === 'j') { e.preventDefault(); setParaProps(ref.current, () => ({ align: { e: 'ctr', l: 'l', r: 'r', j: 'just' }[k] }), theme, baseRef.current); flush(); return; }
        if (e.key === ']' || e.key === '[') { e.preventDefault(); const c = caretRunProps(ref.current, baseRef.current, theme); const s = c ? c.run.size : 18; execRun(ref.current, 'size', e.key === ']' ? s + 2 : Math.max(1, s - 2), theme); flush(); return; }
        if (k === 's' || k === 'z' || k === 'y' || k === 'p') { flush(); }
        if (k === 'z' || k === 'y') { e.preventDefault(); onExit(k === 'z' ? 'undo' : 'redo'); }
        if (k === 's') { e.preventDefault(); ctx.save && ctx.save(); }
      }
    },
    onBlur: () => { flush(); },
  });
}

/* ---------------------------------------------------------------- canvas */
export default function Canvas({ doc, slide, index, zoom, setZoom, fitMode, sel, setSel, change, editing, setEditing, tool, setTool, createAt, readOnly, guides, grid, snap = true, textApi, onContext, onPlaceholderPicture, prompt, spellcheck, onSave, onHistory, groupFocus, setGroupFocus, painter, applyPainter, rulers, units = 'cm', cropId, setCropId, peers = [], comments, focusComment, onComment, selAnim }) {
  const wrap = useRef(null);
  const paper = useRef(null);
  const [drag, setDragNow] = useState(null); // live preview { els: { id: el }, guides: [], marquee, create }
  const dragRef = useRef(null);
  const raf = useRef(0); const pending = useRef(null);
  const setDrag = useCallback((v) => {
    if (v === null) { if (raf.current) cancelAnimationFrame(raf.current); raf.current = 0; setDragNow(null); return; }
    pending.current = v;
    if (!raf.current) raf.current = requestAnimationFrame(() => { raf.current = 0; setDragNow(pending.current); });
  }, []);
  useEffect(() => () => { if (raf.current) cancelAnimationFrame(raf.current); }, []);
  const userGuides = doc.guides || [];
  const W = S(doc.size.w); const H = S(doc.size.h);

  // fit to window
  useLayoutEffect(() => {
    if (!fitMode || !wrap.current) return undefined;
    const fit = () => { const r = wrap.current.getBoundingClientRect(); const z = Math.max(0.1, Math.min((r.width - 48) / W, (r.height - 48) / H)); setZoom(Math.round(z * 1000) / 1000, true); };
    fit();
    const ro = new ResizeObserver(fit); ro.observe(wrap.current);
    return () => ro.disconnect();
  }, [fitMode, W, H]); // eslint-disable-line react-hooks/exhaustive-deps

  const ptFromEvent = useCallback((e) => {
    const r = paper.current.getBoundingClientRect();
    return { x: toEmu((e.clientX - r.left) / zoom), y: toEmu((e.clientY - r.top) / zoom) };
  }, [zoom]);

  /** top-level element (or child of the focused group) under the pointer */
  const hitId = (target) => {
    let n = target; const chain = [];
    while (n && n !== paper.current) { if (n.dataset && n.dataset.el) chain.push(n.dataset.el); n = n.parentNode; }
    chain.reverse(); // outermost first
    const top = chain.find((id) => slide.elements.some((e) => e.id === id));
    if (!top) return null;
    if (groupFocus && top === groupFocus) { const child = chain[chain.indexOf(top) + 1]; return child || top; }
    return top;
  };
  const elsOf = (ids) => ids.map((id) => findEl(slide.elements, id)).filter(Boolean);

  /* --------------- snapping */
  const snapLines = useMemo(() => {
    const xs = [0, doc.size.w / 2, doc.size.w]; const ys = [0, doc.size.h / 2, doc.size.h];
    for (const g of doc.guides || []) (g.axis === 'x' ? xs : ys).push(g.pos);
    for (const e of slide.elements) {
      if (sel.includes(e.id)) continue;
      const b = rotatedBox(e);
      xs.push(b.x, b.x + b.w / 2, b.x + b.w); ys.push(b.y, b.y + b.h / 2, b.y + b.h);
    }
    return { xs, ys };
  }, [slide.elements, sel, doc.size, doc.guides]);
  const snapThreshold = toEmu(6 / zoom);
  const gridStep = 360000 / 2; // 0.5 cm
  function snapMove(box, dx, dy, noSnap) {
    if (noSnap || !snap) return { dx, dy, lines: [] };
    const lines = [];
    const pick = (cands, vals, d) => {
      let best = null;
      for (const v of vals) for (const c of cands) { const diff = c - (v + d); if (Math.abs(diff) <= snapThreshold && (!best || Math.abs(diff) < Math.abs(best.diff))) best = { diff, at: c }; }
      return best;
    };
    let gx = null; let gy = null;
    if (grid) { gx = Math.round((box.x + dx) / gridStep) * gridStep - (box.x + dx); gy = Math.round((box.y + dy) / gridStep) * gridStep - (box.y + dy); }
    const bx = pick(snapLines.xs, [box.x, box.x + box.w / 2, box.x + box.w], dx);
    const by = pick(snapLines.ys, [box.y, box.y + box.h / 2, box.y + box.h], dy);
    if (bx) { dx += bx.diff; lines.push({ x: bx.at }); } else if (gx !== null && Math.abs(gx) < snapThreshold) dx += gx;
    if (by) { dy += by.diff; lines.push({ y: by.at }); } else if (gy !== null && Math.abs(gy) < snapThreshold) dy += gy;
    return { dx, dy, lines };
  }
  const snapVal = (v, axis, noSnap) => {
    if (noSnap || !snap) return { v, line: null };
    const list = axis === 'x' ? snapLines.xs : snapLines.ys;
    let best = null; for (const c of list) { const d = c - v; if (Math.abs(d) <= snapThreshold && (!best || Math.abs(d) < Math.abs(best))) best = d; }
    return best === null ? { v, line: null } : { v: v + best, line: axis === 'x' ? { x: v + best } : { y: v + best } };
  };

  /* --------------- pointer */
  const onPointerDown = (e) => {
    if (e.button !== 0) return;
    if (editing && !e.target.closest('.of-slides-edit')) setEditing(null);
    // drawing/moving calls preventDefault (no text selection) – which also keeps the focus where it was (often the
    // ribbon button that just closed, i.e. nowhere). Keep the keyboard on the editor: Ctrl+S, Delete, Ctrl+Z …
    if (!e.target.closest('.of-slides-edit')) { const root = paper.current && paper.current.closest('.of-slides-root'); if (root && !root.contains(document.activeElement)) root.focus({ preventScroll: true }); }
    const p = ptFromEvent(e);
    if (tool && !readOnly) {
      e.preventDefault();
      paper.current.setPointerCapture(e.pointerId);
      dragRef.current = { mode: 'create', start: p, cur: p, tool };
      setDrag({ create: { x: p.x, y: p.y, w: 0, h: 0 } });
      return;
    }
    const gEl = e.target.closest && e.target.closest('[data-guide]');
    if (gEl && !readOnly) {
      e.preventDefault(); e.stopPropagation(); paper.current.setPointerCapture(e.pointerId);
      const g = userGuides.find((x) => x.id === gEl.dataset.guide);
      if (g) { dragRef.current = { mode: 'guide', guide: g, start: p }; return; }
    }
    const handle0 = e.target.dataset && e.target.dataset.handle;
    if (handle0 && /^c-/.test(handle0) && !readOnly) {
      e.preventDefault(); e.stopPropagation(); paper.current.setPointerCapture(e.pointerId);
      const el = findEl(slide.elements, cropId);
      if (el) { dragRef.current = { mode: 'crop', handle: handle0.slice(2), start: p, orig: el }; return; }
    }
    if (handle0 === 'pathEnd' && selAnim && !readOnly) {
      e.preventDefault(); e.stopPropagation(); paper.current.setPointerCapture(e.pointerId);
      dragRef.current = { mode: 'path', start: p, anim: selAnim }; return;
    }
    const handle = e.target.dataset && e.target.dataset.handle;
    if (handle && !readOnly) {
      e.preventDefault(); e.stopPropagation();
      paper.current.setPointerCapture(e.pointerId);
      const els = elsOf(sel);
      dragRef.current = { mode: handle === 'rot' ? 'rotate' : /^p[12]$/.test(handle) ? 'lineEnd' : 'resize', handle, start: p, orig: els, box: els.length === 1 ? { x: els[0].x, y: els[0].y, w: els[0].w, h: els[0].h } : bbox(els.map((x) => ({ ...x, rot: 0 }))) };
      return;
    }
    const id = hitId(e.target);
    if (painter && id) { applyPainter(id); return; }
    if (!id) {
      if (!e.shiftKey && !e.ctrlKey) setSel([]);
      setGroupFocus(null);
      paper.current.setPointerCapture(e.pointerId);
      dragRef.current = { mode: 'marquee', start: p, cur: p, add: e.shiftKey || e.ctrlKey ? sel : [] };
      return;
    }
    let next = sel;
    if (e.shiftKey || e.ctrlKey || e.metaKey) next = sel.includes(id) ? sel.filter((x) => x !== id) : [...sel, id];
    else if (!sel.includes(id)) next = [id];
    setSel(next);
    if (groupFocus && !findEl(slide.elements.find((x) => x.id === groupFocus) ? [slide.elements.find((x) => x.id === groupFocus)] : [], id)) setGroupFocus(null);
    if (readOnly) return;
    const els = elsOf(next).filter((x) => !x.locked);
    if (!els.length) return;
    paper.current.setPointerCapture(e.pointerId);
    dragRef.current = { mode: 'move', start: p, orig: els, box: bbox(els), moved: false, alt: e.altKey, copy: e.ctrlKey };
  };
  const onPointerMove = (e) => {
    const d = dragRef.current;
    if (!d) return;
    const p = ptFromEvent(e);
    const noSnap = e.altKey;
    if (d.mode === 'marquee') { d.cur = p; setDrag({ marquee: norm(d.start, p) }); return; }
    if (d.mode === 'create') {
      d.cur = p;
      let w = p.x - d.start.x; let hh = p.y - d.start.y;
      if (e.shiftKey && !isLineTool(d.tool)) { const m = Math.max(Math.abs(w), Math.abs(hh)); w = Math.sign(w || 1) * m; hh = Math.sign(hh || 1) * m; }
      d.box = isLineTool(d.tool) ? { x: d.start.x, y: d.start.y, w, h: e.shiftKey ? snapAngleBox(w, hh) : hh } : norm(d.start, { x: d.start.x + w, y: d.start.y + hh });
      setDrag({ create: d.box, line: isLineTool(d.tool) }); return;
    }
    const dx = p.x - d.start.x; const dy = p.y - d.start.y;
    if (d.mode === 'guide' || d.mode === 'newguide') {
      const pos = Math.round(d.guide.axis === 'x' ? p.x : p.y);
      d.pos = pos; setDrag({ guide: { ...d.guide, pos }, gone: d.guide.axis === 'x' ? p.x < -toEmu(8 / zoom) || p.x > doc.size.w + toEmu(8 / zoom) : p.y < -toEmu(8 / zoom) || p.y > doc.size.h + toEmu(8 / zoom) });
      return;
    }
    if (d.mode === 'crop') {
      const el = d.orig; const c = { l: 0, t: 0, r: 0, b: 0, ...(el.crop || {}) };
      const fw = el.w / Math.max(0.02, 1 - c.l - c.r); const fh = el.h / Math.max(0.02, 1 - c.t - c.b);
      const n = { ...el, crop: { ...c } };
      const hd = d.handle;
      if (hd.includes('w')) { const v = Math.max(-c.l * fw, Math.min(el.w - 12700 * 4, dx)); n.x = el.x + v; n.w = el.w - v; n.crop.l = c.l + v / fw; }
      if (hd.includes('e')) { const v = Math.max(-(el.w - 12700 * 4), Math.min(c.r * fw, dx)); n.w = el.w + v; n.crop.r = c.r - v / fw; }
      if (hd.includes('n')) { const v = Math.max(-c.t * fh, Math.min(el.h - 12700 * 4, dy)); n.y = el.y + v; n.h = el.h - v; n.crop.t = c.t + v / fh; }
      if (hd.includes('s')) { const v = Math.max(-(el.h - 12700 * 4), Math.min(c.b * fh, dy)); n.h = el.h + v; n.crop.b = c.b - v / fh; }
      for (const k of ['l', 't', 'r', 'b']) n.crop[k] = Math.max(0, Math.min(0.98, n.crop[k]));
      n.x = Math.round(n.x); n.y = Math.round(n.y); n.w = Math.round(n.w); n.h = Math.round(n.h);
      d.result = { [el.id]: n };
      setDrag({ els: d.result });
      return;
    }
    if (d.mode === 'path') {
      const el = findEl(slide.elements, d.anim.el);
      if (!el) return;
      const pts = pathPoints(d.anim.path, 2); const end = pts[pts.length - 1];
      const nx = end[0] + dx / doc.size.w; const ny = end[1] + dy / doc.size.h;
      d.path = scalePath(d.anim.path, end, [nx, ny]);
      setDrag({ path: d.path });
      return;
    }
    if (d.mode === 'move') {
      if (!d.moved && Math.hypot(dx, dy) < toEmu(3 / zoom)) return;
      d.moved = true;
      let mdx = dx; let mdy = dy;
      if (e.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) mdy = 0; else mdx = 0; }
      const s = snapMove(d.box, mdx, mdy, noSnap);
      d.dx = s.dx; d.dy = s.dy;
      setDrag({ els: Object.fromEntries(d.orig.map((x) => [x.id, moveEl(x, s.dx, s.dy)])), lines: s.lines });
      return;
    }
    if (d.mode === 'rotate') {
      const el = d.orig[0];
      const cx = el.x + el.w / 2; const cy = el.y + el.h / 2;
      let a = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
      if (e.shiftKey) a = Math.round(a / 15) * 15; else { const r = Math.round(a / 90) * 90; if (Math.abs(a - r) < 3) a = r; }
      a = ((Math.round(a) % 360) + 360) % 360;
      d.result = { [el.id]: { ...el, rot: a } };
      setDrag({ els: d.result, angle: a });
      return;
    }
    if (d.mode === 'lineEnd') {
      const el = d.orig[0];
      const p1 = { x: el.x + (el.flipH ? el.w : 0), y: el.y + (el.flipV ? el.h : 0) };
      const p2 = { x: el.x + (el.flipH ? 0 : el.w), y: el.y + (el.flipV ? 0 : el.h) };
      let q = { x: snapVal(p.x, 'x', noSnap).v, y: snapVal(p.y, 'y', noSnap).v };
      const other = d.handle === 'p1' ? p2 : p1;
      if (e.shiftKey) { const ww = q.x - other.x; const hh = q.y - other.y; const a = Math.round(Math.atan2(hh, ww) / (Math.PI / 4)) * (Math.PI / 4); const len = Math.hypot(ww, hh); q = { x: other.x + Math.cos(a) * len, y: other.y + Math.sin(a) * len }; }
      const a1 = d.handle === 'p1' ? q : p1; const a2 = d.handle === 'p1' ? p2 : q;
      d.result = { [el.id]: { ...el, x: Math.round(Math.min(a1.x, a2.x)), y: Math.round(Math.min(a1.y, a2.y)), w: Math.max(1, Math.round(Math.abs(a2.x - a1.x))), h: Math.max(1, Math.round(Math.abs(a2.y - a1.y))), flipH: a2.x < a1.x, flipV: a2.y < a1.y } };
      setDrag({ els: d.result });
      return;
    }
    if (d.mode === 'resize') {
      const lines = [];
      if (d.orig.length === 1 && d.orig[0].type !== 'group' && d.orig[0].rot) {
        // rotated single element: resize in its own frame
        const el = d.orig[0];
        const th = ((el.rot || 0) * Math.PI) / 180;
        const cx = el.x + el.w / 2; const cy = el.y + el.h / 2;
        const lx = (p.x - cx) * Math.cos(-th) - (p.y - cy) * Math.sin(-th); const ly = (p.x - cx) * Math.sin(-th) + (p.y - cy) * Math.cos(-th);
        let l = -el.w / 2; let r = el.w / 2; let tt = -el.h / 2; let b = el.h / 2;
        if (d.handle.includes('w')) l = Math.min(lx, r - 12700); if (d.handle.includes('e')) r = Math.max(lx, l + 12700);
        if (d.handle.includes('n')) tt = Math.min(ly, b - 12700); if (d.handle.includes('s')) b = Math.max(ly, tt + 12700);
        if (keepRatio(e, el, d.handle)) { const s = Math.max((r - l) / el.w, (b - tt) / el.h); const nw = el.w * s; const nh = el.h * s; if (d.handle.includes('w')) l = r - nw; else r = l + nw; if (d.handle.includes('n')) tt = b - nh; else b = tt + nh; }
        const ncx = (l + r) / 2; const ncy = (tt + b) / 2;
        const gx = cx + ncx * Math.cos(th) - ncy * Math.sin(th); const gy = cy + ncx * Math.sin(th) + ncy * Math.cos(th);
        const nw = r - l; const nh = b - tt;
        d.result = { [el.id]: fitEl(el, { x: el.x, y: el.y, w: el.w, h: el.h }, { x: gx - nw / 2, y: gy - nh / 2, w: nw, h: nh }) };
        setDrag({ els: d.result });
        return;
      }
      const B = d.box;
      let x1 = B.x; let y1 = B.y; let x2 = B.x + B.w; let y2 = B.y + B.h;
      if (d.handle.includes('w')) { const s = snapVal(B.x + dx, 'x', noSnap); x1 = Math.min(s.v, x2 - 12700); if (s.line) lines.push(s.line); }
      if (d.handle.includes('e')) { const s = snapVal(B.x + B.w + dx, 'x', noSnap); x2 = Math.max(s.v, x1 + 12700); if (s.line) lines.push(s.line); }
      if (d.handle.includes('n')) { const s = snapVal(B.y + dy, 'y', noSnap); y1 = Math.min(s.v, y2 - 12700); if (s.line) lines.push(s.line); }
      if (d.handle.includes('s')) { const s = snapVal(B.y + B.h + dy, 'y', noSnap); y2 = Math.max(s.v, y1 + 12700); if (s.line) lines.push(s.line); }
      if (keepRatio(e, d.orig[0], d.handle, d.orig.length)) {
        const s = Math.max((x2 - x1) / B.w, (y2 - y1) / B.h);
        const nw = B.w * s; const nh = B.h * s;
        if (d.handle.includes('w')) x1 = x2 - nw; else if (d.handle.includes('e')) x2 = x1 + nw; else { x1 = B.x + B.w / 2 - nw / 2; x2 = x1 + nw; }
        if (d.handle.includes('n')) y1 = y2 - nh; else if (d.handle.includes('s')) y2 = y1 + nh; else { y1 = B.y + B.h / 2 - nh / 2; y2 = y1 + nh; }
      }
      if (e.ctrlKey) { const cx = B.x + B.w / 2; const cy = B.y + B.h / 2; const hw = Math.max(Math.abs(x1 - cx), Math.abs(x2 - cx)); const hh = Math.max(Math.abs(y1 - cy), Math.abs(y2 - cy)); if (/[we]/.test(d.handle)) { x1 = cx - hw; x2 = cx + hw; } if (/[ns]/.test(d.handle)) { y1 = cy - hh; y2 = cy + hh; } }
      const to = { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
      d.result = Object.fromEntries(d.orig.map((x) => [x.id, fitEl(x, B, to)]));
      setDrag({ els: d.result, lines });
    }
  };
  const onPointerUp = (e) => {
    const d = dragRef.current;
    dragRef.current = null;
    try { paper.current.releasePointerCapture(e.pointerId); } catch { /* ignore */ }
    setDrag(null);
    if (!d) return;
    if (d.mode === 'marquee') {
      const m = norm(d.start, d.cur);
      if (m.w < 12700 && m.h < 12700) return;
      const hit = slide.elements.filter((x) => { const b = rotatedBox(x); return b.x >= m.x && b.y >= m.y && b.x + b.w <= m.x + m.w && b.y + b.h <= m.y + m.h; }).map((x) => x.id);
      setSel([...new Set([...d.add, ...hit])]);
      return;
    }
    if (d.mode === 'guide' || d.mode === 'newguide') {
      const pos = d.pos ?? d.guide.pos;
      const lim = d.guide.axis === 'x' ? doc.size.w : doc.size.h;
      const gone = pos < -toEmu(8 / zoom) || pos > lim + toEmu(8 / zoom);
      change((dc) => ({ ...dc, guides: gone ? (dc.guides || []).filter((g) => g.id !== d.guide.id) : d.mode === 'newguide' ? [...(dc.guides || []), { ...d.guide, pos }] : (dc.guides || []).map((g) => (g.id === d.guide.id ? { ...g, pos } : g)) }));
      return;
    }
    if (d.mode === 'path') { if (d.path) change((dc) => updateSlide(dc, slide.id, (s) => ({ ...s, animations: (s.animations || []).map((a) => (a.id === d.anim.id ? { ...a, path: d.path } : a)) })), 'path'); return; }
    if (d.mode === 'create') { setTool(null); createAt(d.tool, d.box && (Math.abs(d.box.w) > 25400 || Math.abs(d.box.h) > 25400) ? d.box : { x: d.start.x, y: d.start.y, w: 0, h: 0 }); return; }
    if (d.mode === 'move') {
      if (!d.moved) return;
      const { dx = 0, dy = 0 } = d;
      if (d.copy) { createAt('__copy', { dx, dy, ids: d.orig.map((x) => x.id) }); return; }
      change((dc) => updateEls(dc, slide.id, d.orig.map((x) => x.id), (x) => moveEl(x, dx, dy)), 'move');
      return;
    }
    if (d.result) change((dc) => updateEls(dc, slide.id, Object.keys(d.result), (x) => d.result[x.id] || x), d.mode === 'crop' ? 'crop' : d.mode);
  };

  const onDoubleClick = (e) => {
    if (readOnly) return;
    // pointer capture makes the paper the event target – look at what is really under the pointer
    const under = document.elementFromPoint(e.clientX, e.clientY) || e.target;
    const id = hitId(under);
    if (!id) return;
    const top = slide.elements.find((x) => x.id === id);
    if (top && top.type === 'group' && !top.preserved) {
      // enter the group: select the child under the pointer
      setGroupFocus(top.id);
      let n = under; let child = null;
      while (n && n !== paper.current) { if (n.dataset && n.dataset.el && n.dataset.el !== top.id && top.children.some((c) => c.id === n.dataset.el)) child = n.dataset.el; n = n.parentNode; }
      if (child) setSel([child]);
      return;
    }
    const el = findEl(slide.elements, id);
    if (!el || el.locked || el.preserved) return;
    if ((el.ph === 'pic' || el.photoSlot) && onPlaceholderPicture) { onPlaceholderPicture(el); return; }
    if (el.type === 'image' && setCropId) { setSel([el.id]); setCropId(cropId === el.id ? null : el.id); return; }
    if (el.type === 'table') { const td = under.closest('[data-cell]'); const [r, c] = td ? td.dataset.cell.split(',').map(Number) : [0, 0]; setSel([el.id]); setEditing({ id: el.id, cell: { r, c } }); return; }
    if (el.type === 'shape') { setSel([el.id]); setEditing({ id: el.id }); }
  };

  /* --------------- text editing glue */
  const editCtx = { spellcheck, save: onSave, slideNumber: index + 1 };
  const renderEditor = (el) => h(TextEditor, {
    key: `ed-${el.id}`, body: el.text, theme: doc.theme, w: S(el.w), h: S(el.h), ctx: editCtx, apiRef: textApi, autoGrow: el.text.autofit === 'resize', selectAll: editing && editing.selectAll,
    onChange: (nb, scrollH) => change((dc) => updateEls(dc, slide.id, [el.id], (x) => {
      const n = { ...x, text: { ...nb, fontScale: undefined } };
      if (scrollH && x.text.autofit === 'resize') { const need = Math.round(toEmu(scrollH)); if (Math.abs(need - x.h) > 6350) n.h = need; }
      return n;
    }), `text:${el.id}`),
    onExit: (why) => { if (why === 'undo' || why === 'redo') { setEditing(null); onHistory(why); } else setEditing(null); },
  });
  const renderCellEditor = (el, r, c, cellBody) => {
    const colsW = (el.cols || []).slice(c, c + (((el.rows[r].cells[c] || {}).span || {}).c || 1)).reduce((a, b) => a + b, 0) * (el.w / ((el.cols || []).reduce((a, b) => a + b, 0) || el.w));
    return h(TextEditor, {
      key: `cell-${el.id}-${r}-${c}`, body: cellBody, theme: doc.theme, w: S(colsW), h: S(el.rows[r].h), ctx: { ...editCtx, tableCell: true }, apiRef: textApi,
      onChange: (nb) => change((dc) => updateEls(dc, slide.id, [el.id], (x) => ({ ...x, rows: x.rows.map((row, ri) => (ri !== r ? row : { ...row, cells: row.cells.map((cc, ci) => (ci !== c ? cc : { ...cc, text: { ...nb, lvl: cc.text.lvl } })) })) })), `cell:${el.id}:${r}:${c}`),
      onExit: (why) => {
        if (why === 'nextCell' || why === 'prevCell') {
          const cells = []; el.rows.forEach((row, ri) => row.cells.forEach((cc, ci) => { if (!cc.merged) cells.push({ r: ri, c: ci }); }));
          const i = cells.findIndex((x) => x.r === r && x.c === c);
          const n = cells[(i + (why === 'nextCell' ? 1 : -1) + cells.length) % cells.length];
          setEditing({ id: el.id, cell: n });
        } else if (why === 'undo' || why === 'redo') { setEditing(null); onHistory(why); } else setEditing(null);
      },
    });
  };

  /* --------------- overlay (selection, handles, guides) */
  const live = drag && drag.els ? drag.els : null;
  const shownSlide = live ? { ...slide, elements: slide.elements.map((x) => mapLive(x, live)) } : slide;
  const selEls = sel.map((id) => (live && live[id]) || findEl(shownSlide.elements, id)).filter(Boolean);
  const z = zoom;
  const box = (b) => ({ left: S(b.x) * z, top: S(b.y) * z, width: S(b.w) * z, height: S(b.h) * z });
  const overlay = [];
  if (groupFocus) { const g = findEl(shownSlide.elements, groupFocus); if (g) overlay.push(h('div', { key: 'gf', className: 'of-slides-groupbox', style: { ...box(g), transform: g.rot ? `rotate(${g.rot}deg)` : undefined } })); }
  for (const el of selEls) {
    const st = { ...box(el), transform: el.rot ? `rotate(${el.rot}deg)` : undefined };
    if (el.type === 'line') {
      const p1 = { left: S(el.x + (el.flipH ? el.w : 0)) * z, top: S(el.y + (el.flipV ? el.h : 0)) * z };
      const p2 = { left: S(el.x + (el.flipH ? 0 : el.w)) * z, top: S(el.y + (el.flipV ? 0 : el.h)) * z };
      if (!readOnly && sel.length === 1 && !el.locked) { overlay.push(h('div', { key: `${el.id}p1`, className: 'of-slides-h', 'data-handle': 'p1', style: p1 })); overlay.push(h('div', { key: `${el.id}p2`, className: 'of-slides-h', 'data-handle': 'p2', style: p2 })); } else overlay.push(h('div', { key: el.id, className: 'of-slides-selbox', style: st }));
      continue;
    }
    const isEditing = editing && editing.id === el.id;
    overlay.push(h('div', { key: el.id, className: `of-slides-selbox ${isEditing ? 'editing' : ''} ${el.locked ? 'locked' : ''}`, style: st },
      !readOnly && sel.length === 1 && !el.locked ? [
        ...HANDLES.map((k) => h('div', { key: k, className: `of-slides-h h-${k}`, 'data-handle': k })),
        h('div', { key: 'rot', className: 'of-slides-h h-rot', 'data-handle': 'rot', title: t('Rotate') }),
      ] : null,
      el.preserved ? h('span', { className: 'of-slides-badge' }, t('Kept as is')) : null));
  }
  if (!readOnly && sel.length > 1) {
    const b = bbox(selEls);
    overlay.push(h('div', { key: 'multi', className: 'of-slides-selbox multi', style: box(b) }, HANDLES.map((k) => h('div', { key: k, className: `of-slides-h h-${k}`, 'data-handle': k }))));
  }
  if (drag && drag.lines) for (const [i, l] of drag.lines.entries()) overlay.push(h('div', { key: `g${i}`, className: 'of-slides-guide', style: l.x !== undefined ? { left: S(l.x) * z, top: 0, width: 0, height: H * z } : { top: S(l.y) * z, left: 0, height: 0, width: W * z } }));
  if (drag && drag.marquee) overlay.push(h('div', { key: 'mq', className: 'of-slides-marquee', style: box(drag.marquee) }));
  if (drag && drag.create) {
    const c = drag.create;
    if (drag.line) overlay.push(h('svg', { key: 'cr', className: 'of-slides-createline', style: { left: 0, top: 0, width: W * z, height: H * z } }, h('line', { x1: S(c.x) * z, y1: S(c.y) * z, x2: S(c.x + c.w) * z, y2: S(c.y + c.h) * z })));
    else overlay.push(h('div', { key: 'cr', className: 'of-slides-marquee create', style: box(c) }));
  }
  if (drag && drag.angle !== undefined) { const el = selEls[0]; if (el) overlay.push(h('div', { key: 'ang', className: 'of-slides-angle', style: { left: S(el.x + el.w / 2) * z, top: S(el.y) * z - 44 } }, `${drag.angle}°`)); }
  // crop mode: the whole picture faded, crop handles on the visible part
  const cropEl = cropId ? selEls.find((x) => x.id === cropId) : null;
  if (cropEl && cropEl.src) {
    const c = { l: 0, t: 0, r: 0, b: 0, ...(cropEl.crop || {}) };
    const fw = cropEl.w / Math.max(0.02, 1 - c.l - c.r); const fh = cropEl.h / Math.max(0.02, 1 - c.t - c.b);
    const full = { x: cropEl.x - c.l * fw, y: cropEl.y - c.t * fh, w: fw, h: fh };
    overlay.unshift(h('img', { key: 'cropghost', className: 'of-slides-cropghost', src: displaySrc(cropEl.src), alt: '', draggable: false, style: { ...box(full), transform: cropEl.rot ? `rotate(${cropEl.rot}deg)` : undefined } }));
    overlay.push(h('div', { key: 'cropbox', className: 'of-slides-cropbox', style: { ...box(cropEl), transform: cropEl.rot ? `rotate(${cropEl.rot}deg)` : undefined } }, [h('img', { key: 'img', src: displaySrc(cropEl.src), alt: '', draggable: false, style: { position: 'absolute', left: -c.l * S(fw) * z, top: -c.t * S(fh) * z, width: S(fw) * z, height: S(fh) * z, maxWidth: 'none', pointerEvents: 'none' } }), ...HANDLES.map((k) => h('div', { key: k, className: `of-slides-ch h-${k}`, 'data-handle': `c-${k}`, title: t('Drag to crop') }))]));
  }
  // motion path of the selected animation (drag its end)
  if (selAnim && selAnim.cls === 'path') {
    const el = findEl(shownSlide.elements, selAnim.el);
    if (el) {
      const path = (drag && drag.path) || selAnim.path;
      const cx = S(el.x + el.w / 2) * z; const cy = S(el.y + el.h / 2) * z;
      const d = pathSvg(path, cx, cy, W * z, H * z);
      const pts = pathPoints(path, 2); const end = pts[pts.length - 1];
      overlay.push(h('svg', { key: 'mpath', className: 'of-slides-mpath', style: { left: 0, top: 0, width: W * z, height: H * z } }, h('path', { d }), h('circle', { cx, cy, r: 4, className: 'start' })));
      if (!readOnly) overlay.push(h('div', { key: 'mpathend', className: 'of-slides-h of-slides-pathend', 'data-handle': 'pathEnd', title: t('Drag to change the path'), style: { left: cx + end[0] * W * z, top: cy + end[1] * H * z } }));
    }
  }
  // guides of the presentation (drag to move, drag off the slide to remove)
  for (const g of userGuides) {
    const live = drag && drag.guide && drag.guide.id === g.id ? drag.guide : g;
    overlay.push(h('div', { key: `ug${g.id}`, 'data-guide': g.id, className: `of-slides-uguide ${g.axis} ${drag && drag.guide && drag.guide.id === g.id && drag.gone ? 'gone' : ''}`, title: t('Guide – drag to move, drag off the slide to remove'), style: g.axis === 'x' ? { left: S(live.pos) * z, top: 0, height: H * z } : { top: S(live.pos) * z, left: 0, width: W * z } }));
  }
  if (drag && drag.guide && !userGuides.some((g) => g.id === drag.guide.id)) overlay.push(h('div', { key: 'ugnew', className: `of-slides-uguide ${drag.guide.axis}`, style: drag.guide.axis === 'x' ? { left: S(drag.guide.pos) * z, top: 0, height: H * z } : { top: S(drag.guide.pos) * z, left: 0, width: W * z } }));
  // co-editors' selections
  for (const pr of peers) for (const id of pr.sel || []) {
    const el = findEl(shownSlide.elements, id);
    if (el) overlay.push(h('div', { key: `peer${pr.id}${id}`, className: 'of-slides-peerbox', style: { ...box(el), borderColor: pr.user.color, transform: el.rot ? `rotate(${el.rot}deg)` : undefined } }, h('span', { style: { background: pr.user.color } }, pr.user.name || '')));
  }
  // comment markers
  if (comments) for (const cm of comments) overlay.push(h('button', { key: `cm${cm.id}`, type: 'button', className: `of-slides-cmark ${focusComment === cm.id ? 'on' : ''}`, title: cm.text || '', 'aria-label': t('Comment'), style: { left: Math.min(W * z - 12, S(cm.pos.x) * z), top: Math.min(H * z - 12, S(cm.pos.y) * z) }, onPointerDown: (e) => e.stopPropagation(), onClick: () => onComment && onComment(cm.id) }, String((cm.author || '?')[0]).toUpperCase()));
  if (guides) { overlay.push(h('div', { key: 'gv', className: 'of-slides-guide static', style: { left: (W / 2) * z, top: 0, height: H * z, width: 0 } })); overlay.push(h('div', { key: 'gh', className: 'of-slides-guide static', style: { top: (H / 2) * z, left: 0, width: W * z, height: 0 } })); }

  return (
    <div className={`of-slides-canvas ${tool ? 'drawing' : ''} ${painter ? 'painting' : ''}`} ref={wrap} onContextMenu={(e) => { const id = hitId(e.target); if (id && !sel.includes(id)) setSel([id]); onContext(e, id); }}>
      <div className="of-slides-paperwrap" style={{ width: W * z + 48, height: H * z + 48 }}>
        <div className={`of-slides-paper ${grid ? 'grid' : ''}`} ref={paper} style={{ width: W * z, height: H * z, '--grid': `${S(gridStep) * z}px` }}
          onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp} onDoubleClick={onDoubleClick}>
          <SlideView doc={doc} slide={shownSlide} index={index} scale={z} mode="edit" editingId={editing && editing.id} editingCell={editing && editing.cell} renderEditor={editing && !editing.cell ? renderEditor : null} renderCellEditor={editing && editing.cell ? renderCellEditor : null} prompt={prompt} />
          <div className="of-slides-overlay">{overlay}</div>
          {rulers && <Rulers W={W} H={H} z={z} units={units} onStart={(axis, e) => { e.stopPropagation(); if (readOnly) return; e.preventDefault(); paper.current.setPointerCapture(e.pointerId); const p = ptFromEvent(e); dragRef.current = { mode: 'newguide', guide: { id: `gd${Date.now().toString(36)}`, axis, pos: Math.round(axis === 'x' ? p.x : p.y) }, start: p }; }} sel={selEls} />}
        </div>
      </div>
    </div>
  );
}

function mapLive(el, live) {
  if (live[el.id]) return live[el.id];
  if (el.type === 'group' && el.children.some((c) => live[c.id])) {
    const children = el.children.map((c) => mapLive(c, live));
    const b = bbox(children.map((c) => ({ ...c, rot: 0 })));
    return { ...el, children, ...b };
  }
  return el;
}
const norm = (a, b) => ({ x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y) });
const isLineTool = (tl) => ['line', 'arrow', 'doubleArrow', 'elbow', 'curve'].includes(tl);
const snapAngleBox = (w, hh) => (Math.abs(hh) < Math.abs(w) * 0.4 ? 0 : Math.abs(w) < Math.abs(hh) * 0.4 ? hh : Math.sign(hh) * Math.abs(w));
function keepRatio(e, el, handle, n = 1) {
  const corner = handle.length === 2;
  if (!corner) return false;
  const lockDefault = n === 1 && el && (el.type === 'image' || el.type === 'media' || el.lockRatio);
  return e.shiftKey ? !lockDefault : lockDefault;
}
/** rulers (cm or in) along the top and left edge of the paper; dragging from one creates a guide */
function Rulers({ W, H, z, units, onStart, sel }) {
  const unitPt = units === 'in' ? 72 : 72 / 2.54; // pt per unit
  const step = unitPt * z; const minor = units === 'in' ? 8 : 4;
  const ticks = (len) => { const out = []; for (let i = 0; i * step / minor <= len * z + 1; i++) { const pos = (i * step) / minor; const major = i % minor === 0; out.push({ pos, major, label: major ? String(i / minor) : null }); } return out; };
  const tx = ticks(W); const ty = ticks(H);
  const b = sel && sel.length ? bbox(sel) : null;
  return (
    <>
      <svg className="of-slides-ruler top" width={W * z} height={18} aria-hidden="true" onPointerDown={(e) => onStart('y', e)}>
        {b && <rect x={S(b.x) * z} y={0} width={S(b.w) * z} height={18} className="hl" />}
        {tx.map((k) => <line key={k.pos} x1={k.pos} x2={k.pos} y1={k.major ? 4 : 12} y2={18} />)}
        {tx.filter((k) => k.label && (step > 24 || +k.label % 2 === 0)).map((k) => <text key={`t${k.pos}`} x={k.pos + 2} y={10}>{k.label}</text>)}
      </svg>
      <svg className="of-slides-ruler left" width={18} height={H * z} aria-hidden="true" onPointerDown={(e) => onStart('x', e)}>
        {b && <rect x={0} y={S(b.y) * z} width={18} height={S(b.h) * z} className="hl" />}
        {ty.map((k) => <line key={k.pos} y1={k.pos} y2={k.pos} x1={k.major ? 4 : 12} x2={18} />)}
        {ty.filter((k) => k.label && (step > 24 || +k.label % 2 === 0)).map((k) => <text key={`t${k.pos}`} x={3} y={k.pos + 10}>{k.label}</text>)}
      </svg>
    </>
  );
}
/** move the end of a PowerPoint motion path: lines get the new end, other paths are stretched */
function scalePath(path, end, to) {
  const f = (n) => +n.toFixed(4);
  const nums = String(path).match(/-?\d*\.?\d+(?:e-?\d+)?/gi) || [];
  if (/^\s*M\s*0\s+0\s+L\s+[-\d.e]+\s+[-\d.e]+\s*E?\s*$/i.test(path) || nums.length <= 4) return `M 0 0 L ${f(to[0])} ${f(to[1])} E`;
  const sx = Math.abs(end[0]) > 1e-4 ? to[0] / end[0] : 1; const sy = Math.abs(end[1]) > 1e-4 ? to[1] / end[1] : 1;
  const ox = Math.abs(end[0]) > 1e-4 ? 0 : to[0]; const oy = Math.abs(end[1]) > 1e-4 ? 0 : to[1];
  let k = 0;
  return String(path).replace(/-?\d*\.?\d+(?:e-?\d+)?/gi, (m) => { const v = parseFloat(m); const r = k % 2 === 0 ? v * sx + (v ? 0 : 0) + (k === nums.length - 2 ? ox : 0) : v * sy + (k === nums.length - 1 ? oy : 0); k++; return String(f(r)); });
}
export { walk };
