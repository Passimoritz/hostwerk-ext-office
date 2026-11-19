/*
 * The one slide renderer – used by the canvas, thumbnails, sorter, presenter, slide show, preview and export.
 * Renders in slide space (1 px = 1 pt) and scales with a CSS transform. User content goes through createElement (h)
 * so Hostwerk's JSX auto-translation never touches slide text.
 */
import { createElement as h, memo, useLayoutEffect, useRef, useState } from 'react';
import { css } from '../themes.js';
import { S } from '../units.js';
import { layoutOf } from '../model.js';
import { shapePath } from './geometry.js';
import { paraProps, runProps, paraCss, runCss, bulletLabels, runText, fontStack } from './text.js';
import Chart from './Chart.jsx';

let gid = 0;
const nextId = () => `g${(gid++).toString(36)}`;

/* ---------------------------------------------------------------- fills */
export function fillCss(fill, theme) {
  if (!fill) return 'transparent';
  if (fill.type === 'solid') return css(fill.color, theme);
  if (fill.type === 'grad') {
    const stops = (fill.stops || []).map((s) => `${css(s.color, theme)} ${s.pos}%`).join(', ');
    if (fill.kind === 'radial') return `radial-gradient(circle, ${stops})`;
    return `linear-gradient(${((fill.angle || 0) + 90) % 360}deg, ${stops})`;
  }
  if (fill.type === 'image') return `url("${fill.src}") center / ${fill.mode === 'tile' ? 'auto' : '100% 100%'} ${fill.mode === 'tile' ? 'repeat' : 'no-repeat'}`;
  return 'transparent';
}
function svgFill(fill, theme, defs, w, h2) {
  if (!fill) return 'none';
  if (fill.type === 'solid') return css(fill.color, theme);
  const id = nextId();
  if (fill.type === 'grad') {
    const stops = (fill.stops || []).map((s, i) => h('stop', { key: i, offset: `${s.pos}%`, stopColor: css(s.color, theme) }));
    if (fill.kind === 'radial') defs.push(h('radialGradient', { key: id, id }, stops));
    else {
      const a = ((fill.angle || 0) * Math.PI) / 180;
      const dx = Math.cos(a) / 2; const dy = Math.sin(a) / 2;
      defs.push(h('linearGradient', { key: id, id, x1: 0.5 - dx, y1: 0.5 - dy, x2: 0.5 + dx, y2: 0.5 + dy }, stops));
    }
    return `url(#${id})`;
  }
  if (fill.type === 'image') {
    defs.push(h('pattern', { key: id, id, patternUnits: 'userSpaceOnUse', width: w, height: h2 }, h('image', { href: fill.src, width: w, height: h2, preserveAspectRatio: 'none' })));
    return `url(#${id})`;
  }
  return 'none';
}
const DASH = { dash: [4, 3], dot: [1, 2], sysDot: [1, 1], sysDash: [3, 1], dashDot: [4, 3, 1, 3], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1] };
const dashArray = (line, wpx) => { const d = DASH[line && line.dash]; return d ? d.map((x) => x * Math.max(1, wpx)).join(' ') : undefined; };
function shadowCss(sh, theme) {
  if (!sh) return undefined;
  const a = ((sh.dir || 45) * Math.PI) / 180; const dist = S(sh.dist || 38100);
  return `drop-shadow(${(Math.cos(a) * dist).toFixed(1)}px ${(Math.sin(a) * dist).toFixed(1)}px ${S(sh.blur || 50800).toFixed(1)}px ${css(`${sh.color || '#000000'}@${sh.alpha ?? 40}`, theme)})`;
}

/* ---------------------------------------------------------------- text */
function TextBody({ body, theme, w, h: H, ctx, mode, prompt, fontScaleOverride }) {
  const ref = useRef(null);
  const [fit, setFit] = useState(1);
  const shrink = body && body.autofit === 'shrink' && !body.fontScale;
  const fs = fontScaleOverride || (body && body.fontScale) || (shrink ? fit : 1);
  useLayoutEffect(() => {
    if (!shrink || !ref.current) return;
    const el = ref.current;
    let f = 1; applyScale(el, 1);
    // shrink until it fits (max 12 steps)
    for (let k = 0; k < 12 && el.scrollHeight > el.clientHeight + 1 && f > 0.3; k++) { f -= 0.06; applyScale(el, f); }
    if (Math.abs(f - fit) > 0.01) setFit(f);
  });
  if (!body) return null;
  const ins = body.insets || {};
  const empty = !(body.paras || []).some((p) => (p.runs || []).some((r) => r.text || r.field));
  const showPrompt = empty && prompt && mode === 'edit';
  const labels = bulletLabels(body, theme);
  const vert = body.vert === 'vert' ? 'vertical-rl' : body.vert === 'vert270' ? 'vertical-rl' : undefined;
  const style = {
    position: 'absolute', left: 0, top: 0, width: w, height: H, boxSizing: 'border-box',
    padding: `${S(ins.t ?? 45720)}px ${S(ins.r ?? 91440)}px ${S(ins.b ?? 45720)}px ${S(ins.l ?? 91440)}px`,
    display: 'flex', flexDirection: 'column', justifyContent: body.anchor === 'ctr' ? 'center' : body.anchor === 'b' ? 'flex-end' : 'flex-start',
    overflow: body.autofit === 'shrink' ? 'hidden' : 'visible', wordWrap: 'break-word', overflowWrap: 'break-word',
    whiteSpace: body.wrap === false ? 'pre' : 'pre-wrap', writingMode: vert, transform: body.vert === 'vert270' ? 'rotate(180deg)' : undefined,
  };
  const paras = showPrompt ? [{ runs: [{ text: prompt }], level: 0, prompt: true }] : body.paras;
  return h('div', { ref, className: 'of-slides-txt', style }, paras.map((p, i) => {
    const pp = paraProps(p, body, theme);
    const lab = showPrompt ? null : labels[i];
    const first = runProps((p.runs || [])[0] || p.endProps || {}, p, body);
    const pcss = paraCss(pp);
    const kids = [];
    if (lab) {
      const bsz = first.size * fs * ((lab.sizePct || 100) / 100);
      kids.push(h('span', { key: 'b', className: 'of-slides-bul', 'data-fs': first.size * ((lab.sizePct || 100) / 100), style: { display: 'inline-block', minWidth: `${Math.max(0, -S(pp.indent))}px`, textIndent: 0, color: css(lab.color || first.color || 'tx1', theme), fontFamily: lab.font ? fontStack(lab.font, theme) : fontStack(first.font, theme), fontSize: `${bsz}px` } }, lab.text));
    }
    const runs = p.runs || [];
    if (!runs.length || runs.every((r) => !(r.text || r.field))) {
      kids.push(h('span', { key: 'e', 'data-fs': parseFloat(runCss(first, theme, 1).fontSize), style: runCss(first, theme, fs) }, '​'));
    } else {
      runs.forEach((r, k) => {
        const rp = runProps(r, p, body);
        const st = runCss(rp, theme, fs);
        if (showPrompt) { st.color = 'rgba(128,128,128,.85)'; }
        const text = runText(r, ctx);
        if (rp.link && mode === 'show') kids.push(h('a', { key: k, style: { ...st, color: st.color, textDecoration: 'underline', cursor: 'pointer' }, 'data-link': rp.link, href: '#', onClick: (e) => { e.preventDefault(); e.stopPropagation(); if (ctx.openLink) ctx.openLink(rp.link); } }, text));
        else kids.push(h('span', { key: k, 'data-fs': parseFloat(runCss(rp, theme, 1).fontSize), style: rp.link ? { ...st, textDecoration: 'underline' } : st }, text));
      });
    }
    return h('div', { key: i, className: 'of-slides-p', style: { ...pcss, margin: 0, marginTop: pcss.marginTop, marginBottom: pcss.marginBottom } }, kids);
  }));
}
function applyScale(el, f) {
  // re-scale the inline font sizes (data-fs = unscaled size from the render), React then renders the final factor
  el.querySelectorAll('[data-fs]').forEach((s) => { const v = parseFloat(s.dataset.fs); if (v) s.style.fontSize = `${v * f}px`; });
}

/* ---------------------------------------------------------------- elements */
function markerDef(kind, color, wpx, defs, start) {
  if (!kind || kind === 'none') return undefined;
  const id = nextId();
  const s = Math.max(3, 3 + wpx * 0.6);
  let shape;
  if (kind === 'oval') shape = h('circle', { cx: 5, cy: 5, r: 4, fill: color });
  else if (kind === 'diamond') shape = h('path', { d: 'M5 0L10 5L5 10L0 5Z', fill: color });
  else if (kind === 'arrow') shape = h('path', { d: 'M0 0L10 5L0 10', fill: 'none', stroke: color, strokeWidth: 2 });
  else if (kind === 'stealth') shape = h('path', { d: 'M0 0L10 5L0 10L3 5Z', fill: color });
  else shape = h('path', { d: 'M0 0L10 5L0 10Z', fill: color });
  defs.push(h('marker', { key: id, id, viewBox: '0 0 10 10', refX: kind === 'oval' || kind === 'diamond' ? 5 : 9, refY: 5, markerWidth: s, markerHeight: s, markerUnits: 'userSpaceOnUse', orient: start ? 'auto-start-reverse' : 'auto' }, shape));
  return `url(#${id})`;
}

function LineEl({ el, theme }) {
  const w = S(el.w); const H = S(el.h);
  const ln = el.line || { color: 'tx1', width: 12700 };
  const color = css(ln.color || 'tx1', theme, '#000');
  const wpx = Math.max(0.5, S(ln.width || 12700));
  const defs = [];
  const x1 = el.flipH ? w : 0; const x2 = el.flipH ? 0 : w; const y1 = el.flipV ? H : 0; const y2 = el.flipV ? 0 : H;
  let d = `M${x1} ${y1}L${x2} ${y2}`;
  if (el.geom === 'elbow') d = `M${x1} ${y1}L${(x1 + x2) / 2} ${y1}L${(x1 + x2) / 2} ${y2}L${x2} ${y2}`;
  if (el.geom === 'curve') d = `M${x1} ${y1}C${(x1 + x2) / 2} ${y1} ${(x1 + x2) / 2} ${y2} ${x2} ${y2}`;
  const ms = markerDef(ln.head, color, wpx, defs, true); const me = markerDef(ln.tail, color, wpx, defs, false);
  return h('svg', { width: Math.max(w, 1), height: Math.max(H, 1), style: { position: 'absolute', left: 0, top: 0, overflow: 'visible' } },
    defs.length ? h('defs', null, defs) : null,
    h('path', { d, fill: 'none', stroke: color, strokeWidth: wpx, strokeDasharray: dashArray(ln, wpx), strokeLinecap: ln.cap || 'butt', markerStart: ms, markerEnd: me }));
}

function ShapeEl({ el, theme, ctx, mode, editing, renderEditor }) {
  const w = S(el.w); const H = S(el.h);
  const defs = [];
  const geo = el.geom && el.geom !== 'rect' || el.fill || el.line ? shapePath(el.geom || 'rect', w, H, el.adj) : null;
  const ln = el.line;
  const wpx = ln ? Math.max(0.25, S(ln.width || 12700)) : 0;
  const fl = [el.flipH && 'scaleX(-1)', el.flipV && 'scaleY(-1)'].filter(Boolean).join(' ');
  const showGeo = geo && (el.fill || ln);
  const ph = el.ph === 'pic' && mode === 'edit' && !el.fill;
  return h('div', { style: { position: 'absolute', inset: 0 } },
    showGeo ? h('svg', { width: w, height: H, viewBox: `0 0 ${w} ${H}`, style: { position: 'absolute', left: 0, top: 0, overflow: 'visible', transform: fl || undefined, filter: shadowCss(el.shadow, theme) } },
      h('defs', null, defs),
      h('path', { d: geo.d, fill: svgFill(el.fill, theme, defs, w, H), fillRule: geo.evenodd ? 'evenodd' : undefined, stroke: ln ? css(ln.color || 'tx1', theme, '#000') : 'none', strokeWidth: wpx || undefined, strokeDasharray: ln ? dashArray(ln, wpx) : undefined, strokeLinejoin: 'round' })) : null,
    ph ? h('div', { className: 'of-slides-picph' }, h('svg', { width: 48, height: 48, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.5 }, h('rect', { x: 3, y: 3, width: 18, height: 18, rx: 2 }), h('circle', { cx: 9, cy: 9, r: 2 }), h('path', { d: 'm21 15-3.1-3.1a2 2 0 0 0-2.8 0L6 21' }))) : null,
    editing && renderEditor ? renderEditor(el) : el.text ? h(TextBody, { body: el.text, theme, w, h: H, ctx, mode, prompt: el.ph && el.ph !== 'pic' ? ctx.prompt && ctx.prompt(el) : null }) : null);
}

function ImageEl({ el, theme }) {
  const w = S(el.w); const H = S(el.h);
  const c = el.crop || {};
  const fw = w / Math.max(0.01, 1 - (c.l || 0) - (c.r || 0)); const fh = H / Math.max(0.01, 1 - (c.t || 0) - (c.b || 0));
  const ln = el.line;
  const geo = el.geom && el.geom !== 'rect' ? shapePath(el.geom, w, H, el.adj) : null;
  return h('div', { style: { position: 'absolute', inset: 0, overflow: 'hidden', opacity: el.opacity ?? 1, clipPath: geo ? `path('${geo.d}')` : undefined, filter: shadowCss(el.shadow, theme), boxShadow: ln ? `inset 0 0 0 ${Math.max(0.5, S(ln.width || 12700))}px ${css(ln.color, theme)}` : undefined, transform: [el.flipH && 'scaleX(-1)', el.flipV && 'scaleY(-1)'].filter(Boolean).join(' ') || undefined } },
    el.src ? h('img', { src: el.src, alt: el.alt || '', draggable: false, style: { position: 'absolute', left: -(c.l || 0) * fw, top: -(c.t || 0) * fh, width: fw, height: fh, maxWidth: 'none', userSelect: 'none', pointerEvents: 'none' } }) : null);
}

function tableCellFill(el, ri, ci, cell, theme) {
  if (cell.fill !== undefined) return cell.fill ? fillCss(cell.fill, theme) : 'transparent';
  const st = el.style || {};
  const acc = st.accent || 'accent1';
  if (st.firstRow && ri === 0) return css(acc, theme);
  if (st.firstCol && ci === 0) return css(acc, theme);
  if (st.bandRows) return css((ri - (st.firstRow ? 1 : 0)) % 2 === 0 ? `${acc}+80` : `${acc}+60`, theme);
  return css(`${acc}+80`, theme);
}
function TableEl({ el, theme, ctx, mode, editingCell, renderCellEditor }) {
  const st = el.style || {};
  const bd = el.border;
  const border = (b) => (b === null ? 'none' : b ? `${Math.max(0.5, S(b.width || 12700))}px ${b.dash && b.dash !== 'solid' ? 'dashed' : 'solid'} ${css(b.color, theme)}` : bd ? `${Math.max(0.5, S(bd.width || 12700))}px solid ${css(bd.color || 'bg1', theme)}` : 'none');
  const totalW = (el.cols || []).reduce((a, b) => a + b, 0) || el.w;
  const sx = el.w / totalW;
  return h('table', { className: 'of-slides-table', style: { position: 'absolute', left: 0, top: 0, width: S(el.w), borderCollapse: 'collapse', tableLayout: 'fixed' } },
    h('colgroup', null, (el.cols || []).map((c, i) => h('col', { key: i, style: { width: S(c * sx) } }))),
    h('tbody', null, (el.rows || []).map((r, ri) => h('tr', { key: ri, style: { height: S(r.h) } }, r.cells.map((cell, ci) => {
      if (cell.merged) return null;
      const header = (st.firstRow && ri === 0) || (st.firstCol && ci === 0);
      const b = cell.borders || {};
      let rowsH = 0; for (let k = 0; k < ((cell.span && cell.span.r) || 1); k++) rowsH += (el.rows[ri + k] || r).h;
      let colsW = 0; for (let k = 0; k < ((cell.span && cell.span.c) || 1); k++) colsW += (el.cols[ci + k] || 0) * sx;
      const body = cell.fill === undefined && header ? { ...cell.text, lvl: (cell.text.lvl && cell.text.lvl.length ? cell.text.lvl : [{}]).map((l) => ({ ...l, b: true, color: 'bg1' })) } : cell.text;
      const isEd = editingCell && editingCell.r === ri && editingCell.c === ci;
      return h('td', { key: ci, colSpan: (cell.span && cell.span.c) || 1, rowSpan: (cell.span && cell.span.r) || 1, 'data-cell': `${ri},${ci}`, style: { background: tableCellFill(el, ri, ci, cell, theme), borderLeft: border(b.l), borderRight: border(b.r), borderTop: border(b.t), borderBottom: border(b.b), padding: 0, position: 'relative', verticalAlign: 'top', height: S(rowsH) } },
        h('div', { style: { position: 'relative', width: S(colsW), height: S(rowsH) } }, isEd && renderCellEditor ? renderCellEditor(el, ri, ci, body) : h(TextBody, { body, theme, w: S(colsW), h: S(rowsH), ctx, mode })));
    })))));
}

function MediaEl({ el, mode }) {
  const w = S(el.w); const H = S(el.h);
  if (mode === 'show' && el.src) {
    if (el.kind === 'audio') return h('audio', { src: el.src, controls: true, style: { position: 'absolute', left: 0, bottom: 0, width: w } });
    return h('video', { src: el.src, controls: true, poster: el.poster || undefined, style: { position: 'absolute', inset: 0, width: w, height: H, objectFit: 'contain', background: '#000' } });
  }
  return h('div', { className: 'of-slides-media', style: { position: 'absolute', inset: 0, background: el.poster ? `url("${el.poster}") center / contain no-repeat #000` : '#1b1f2a' } },
    h('svg', { width: Math.min(w, H) * 0.3, height: Math.min(w, H) * 0.3, viewBox: '0 0 24 24', fill: 'rgba(255,255,255,.9)' }, el.kind === 'audio' ? h('path', { d: 'M9 18V5l12-2v13M9 18a3 3 0 1 1-3-3 3 3 0 0 1 3 3Zm12-2a3 3 0 1 1-3-3 3 3 0 0 1 3 3Z', stroke: 'rgba(255,255,255,.9)', fill: 'none', strokeWidth: 2 }) : h('path', { d: 'M8 5v14l11-7z' })));
}

function UnknownEl({ el }) {
  if (el.src) return h('img', { src: el.src, alt: el.what || '', draggable: false, style: { position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'fill', pointerEvents: 'none' } });
  return h('div', { className: 'of-slides-unknown' }, h('span', null, el.what || 'Object'));
}

export function ElementView({ el, theme, ctx, mode, editingId, renderEditor, elDecor, editingCell, renderCellEditor, origin = { x: 0, y: 0 } }) {
  if (el.hidden && mode !== 'edit') return null;
  const deco = elDecor ? elDecor(el) : null;
  if (deco && deco.skip) return null;
  const style = {
    position: 'absolute', left: S(el.x - origin.x), top: S(el.y - origin.y), width: S(el.w), height: S(el.h),
    transform: el.rot ? `rotate(${el.rot}deg)` : undefined, ...(deco && deco.style),
  };
  let inner = null;
  if (el.type === 'group') {
    inner = (el.children || []).map((c) => h(ElementView, { key: c.id, el: c, theme, ctx, mode, editingId, renderEditor, elDecor, editingCell, renderCellEditor, origin: { x: el.x, y: el.y } }));
  } else if (el.type === 'shape') inner = h(ShapeEl, { el, theme, ctx, mode, editing: editingId === el.id, renderEditor });
  else if (el.type === 'line') inner = h(LineEl, { el, theme });
  else if (el.type === 'image') inner = h(ImageEl, { el, theme });
  else if (el.type === 'table') inner = h(TableEl, { el, theme, ctx, mode, editingCell: editingId === el.id ? editingCell : null, renderCellEditor });
  else if (el.type === 'chart') inner = h(Chart, { el, theme, w: S(el.w), h: S(el.h) });
  else if (el.type === 'media') inner = h(MediaEl, { el, mode });
  else inner = h(UnknownEl, { el });
  return h('div', { className: `of-slides-el${deco && deco.className ? ` ${deco.className}` : ''}`, 'data-el': el.id, style }, inner);
}

/** background of a slide (slide → layout → theme) */
export function slideBackground(doc, slide) {
  const layout = layoutOf(doc, slide);
  return slide.background || (layout && layout.bg) || (doc.theme && doc.theme.bg) || { type: 'solid', color: 'bg1' };
}

function footers(doc, slide, layout, index, ctx) {
  const f = doc.footer || {};
  if (!f.showFooter && !f.showNumber && !f.showDate) return [];
  if (f.hideOnTitle && layout.kind === 'title') return [];
  const out = [];
  const mk = (key, text) => {
    const d = layout.ph && layout.ph[key];
    if (!d) return;
    const lv = (d.lvl && d.lvl[0]) || {};
    out.push(h('div', { key, className: 'of-slides-el', style: { position: 'absolute', left: S(d.x), top: S(d.y), width: S(d.w), height: S(d.h), display: 'flex', alignItems: 'center', justifyContent: lv.align === 'ctr' ? 'center' : lv.align === 'r' ? 'flex-end' : 'flex-start', padding: '0 7px', boxSizing: 'border-box', fontFamily: fontStack(lv.font || '+mn', doc.theme), fontSize: lv.size || 11, color: css(lv.color || 'tx2', doc.theme), whiteSpace: 'nowrap', overflow: 'hidden' } }, text));
  };
  if (f.showDate) mk('dt', f.date || ctx.date);
  if (f.showFooter && f.text) mk('ftr', f.text);
  if (f.showNumber) mk('sldNum', String(index + 1));
  return out;
}

/**
 * <SlideView doc slide index scale mode="edit|thumb|show" … />
 *  editingId/renderEditor: in-place text editing (canvas) · elDecor(el) → { style, className, skip } (animations)
 */
function SlideViewImpl({ doc, slide, index = 0, scale = 1, mode = 'thumb', editingId, renderEditor, elDecor, editingCell, renderCellEditor, prompt, openLink, className = '', style: extra, children, hideIds }) {
  if (!slide) return null;
  const W = S(doc.size.w); const H = S(doc.size.h);
  const layout = layoutOf(doc, slide);
  const ctx = { slideNumber: index + 1, date: new Date().toLocaleDateString(typeof document !== 'undefined' ? document.documentElement.lang || undefined : undefined), footer: (doc.footer || {}).text, prompt, openLink };
  const bg = slideBackground(doc, slide);
  const decor = slide.showMaster === false ? [] : (layout.decor || []);
  return h('div', { className: `of-slides-view ${className}`, style: { width: W * scale, height: H * scale, position: 'relative', overflow: 'hidden', ...extra } },
    h('div', { className: 'of-slides-stage', style: { width: W, height: H, transform: `scale(${scale})`, transformOrigin: '0 0', position: 'absolute', left: 0, top: 0, background: fillCss(bg, doc.theme), overflow: 'hidden' } },
      decor.map((e) => h(ElementView, { key: e.id, el: e, theme: doc.theme, ctx, mode: mode === 'edit' ? 'thumb' : mode })),
      footers(doc, slide, layout, index, ctx),
      slide.elements.filter((e) => !(hideIds && hideIds.has(e.id))).map((e) => h(ElementView, { key: e.id, el: e, theme: doc.theme, ctx, mode, editingId, renderEditor, elDecor, editingCell, renderCellEditor })),
      children));
}
export const SlideView = memo(SlideViewImpl);
export default SlideView;
