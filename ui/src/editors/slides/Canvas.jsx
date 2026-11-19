/*
 * The editing canvas: one slide, selection, move/resize/rotate with smart guides and snapping, marquee, drawing new
 * shapes, in-place text editing (contentEditable, see render/edit.js), table cell editing.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, createElement as h } from 'react';
import { t } from '@hostwerk/i18n';
import SlideView from './render/SlideView.jsx';
import { bodyToHtml, htmlToBody, refreshParas, saveRange, restoreRange, execRun, setParaProps, caretRunProps } from './render/edit.js';
import { S, EMU_PER_PT } from './units.js';
import { findEl, rotatedBox, bbox, moveEl, fitEl, updateEls, walk } from './model.js';

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
export default function Canvas({ doc, slide, index, zoom, setZoom, fitMode, sel, setSel, change, editing, setEditing, tool, setTool, createAt, readOnly, guides, grid, snap = true, textApi, onContext, onPlaceholderPicture, prompt, spellcheck, onSave, onHistory, groupFocus, setGroupFocus, painter, applyPainter }) {
  const wrap = useRef(null);
  const paper = useRef(null);
  const [drag, setDrag] = useState(null); // live preview { els: { id: el }, guides: [], marquee, create }
  const dragRef = useRef(null);
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
    for (const e of slide.elements) {
      if (sel.includes(e.id)) continue;
      const b = rotatedBox(e);
      xs.push(b.x, b.x + b.w / 2, b.x + b.w); ys.push(b.y, b.y + b.h / 2, b.y + b.h);
    }
    return { xs, ys };
  }, [slide.elements, sel, doc.size]);
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
    if (d.mode === 'create') { setTool(null); createAt(d.tool, d.box && (Math.abs(d.box.w) > 25400 || Math.abs(d.box.h) > 25400) ? d.box : { x: d.start.x, y: d.start.y, w: 0, h: 0 }); return; }
    if (d.mode === 'move') {
      if (!d.moved) return;
      const { dx = 0, dy = 0 } = d;
      if (d.copy) { createAt('__copy', { dx, dy, ids: d.orig.map((x) => x.id) }); return; }
      change((dc) => updateEls(dc, slide.id, d.orig.map((x) => x.id), (x) => moveEl(x, dx, dy)), 'move');
      return;
    }
    if (d.result) change((dc) => updateEls(dc, slide.id, Object.keys(d.result), (x) => d.result[x.id] || x), d.mode);
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
  if (guides) { overlay.push(h('div', { key: 'gv', className: 'of-slides-guide static', style: { left: (W / 2) * z, top: 0, height: H * z, width: 0 } })); overlay.push(h('div', { key: 'gh', className: 'of-slides-guide static', style: { top: (H / 2) * z, left: 0, width: W * z, height: 0 } })); }

  return (
    <div className={`of-slides-canvas ${tool ? 'drawing' : ''} ${painter ? 'painting' : ''}`} ref={wrap} onContextMenu={(e) => { const id = hitId(e.target); if (id && !sel.includes(id)) setSel([id]); onContext(e, id); }}>
      <div className="of-slides-paperwrap" style={{ width: W * z + 48, height: H * z + 48 }}>
        <div className={`of-slides-paper ${grid ? 'grid' : ''}`} ref={paper} style={{ width: W * z, height: H * z, '--grid': `${S(gridStep) * z}px` }}
          onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp} onDoubleClick={onDoubleClick}>
          <SlideView doc={doc} slide={shownSlide} index={index} scale={z} mode="edit" editingId={editing && editing.id} editingCell={editing && editing.cell} renderEditor={editing && !editing.cell ? renderEditor : null} renderCellEditor={editing && editing.cell ? renderCellEditor : null} prompt={prompt} />
          <div className="of-slides-overlay">{overlay}</div>
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
export { walk };
