/*
 * Page area of the PDF editor: lazy page rendering (canvas + text layer + pdf.js annotation layer for links and
 * forms) and the SVG overlay that shows and edits our own annotations (see annotations.js – page space, y up).
 *
 * Performance (500-page PDFs): pages without a pdf.js proxy yet are placeholders of the first page's size (index.jsx
 * loads proxies in the background and on demand: needProxy); at most two pages render at once, nearest to the middle
 * of the window first; a page leaving the window cancels its render; rendered canvases share a pixel budget
 * (≈ 40 MP) – the ones farthest away are dropped first. Zooming stretches the existing bitmaps at once (CSS
 * transform, 60 fps) and renders sharp at the device pixel ratio ~160 ms after the last zoom step, keeping the point
 * under the mouse (or the middle) in place.
 */
import { useEffect, useLayoutEffect, useMemo, useRef, useState, memo } from 'react';
import { pdfjsLib } from '../_shared/pdfjs.js';
import { bbox, inkPath, arrowHead, textLines, mergeQuads, MARKUP } from './annotations.js';
import { Keep } from '../_shared/keep.jsx';

export const CSS_UNITS = 96 / 72;
const DRAW_TOOLS = new Set(['ink', 'rect', 'ellipse', 'line', 'arrow', 'redact']);
const PLACE_TOOLS = new Set(['text', 'note', 'stamp', 'image', 'signature']);
const PIXEL_BUDGET = 40e6;
const MAX_CANVAS = 16.7e6;

const pngUrls = new WeakMap();
function pngUrl(bytes) {
  if (!pngUrls.has(bytes)) pngUrls.set(bytes, URL.createObjectURL(new Blob([bytes], { type: 'image/png' })));
  return pngUrls.get(bytes);
}

/** total rotation of a model page in the view */
export const pageRotation = (proxy, mp, viewRot) => ((proxy ? proxy.rotate : 0) + (mp.rot || 0) + viewRot) % 360;

/* ---------------- one annotation as SVG (inside the page-space group) ---------------- */
function AnnotShape({ a, selected }) {
  const common = { 'data-aid': a.id, className: `of-pdf-ann${selected ? ' sel' : ''}` };
  const stroke = a.color || '#000';
  switch (a.type) {
    case 'highlight': return <g {...common}>{a.quads.map((q, i) => <rect key={i} x={q.x} y={q.y} width={q.w} height={q.h} fill={stroke} opacity="0.4" style={{ mixBlendMode: 'multiply' }} />)}</g>;
    case 'underline': case 'strike':
      return <g {...common}>{a.quads.map((q, i) => { const t = Math.max(0.8, q.h * 0.07); const y = a.type === 'underline' ? q.y + t : q.y + q.h * 0.45; return <g key={i}><rect x={q.x} y={q.y} width={q.w} height={q.h} fill="transparent" /><line x1={q.x} x2={q.x + q.w} y1={y} y2={y} stroke={stroke} strokeWidth={t} /></g>; })}</g>;
    case 'ink': return <g {...common}>{a.paths.map((p, i) => <path key={i} d={inkPath(p)} fill="none" stroke={stroke} strokeWidth={a.width || 2} strokeLinecap="round" strokeLinejoin="round" />)}</g>;
    case 'rect': return <rect {...common} x={a.x} y={a.y} width={a.w} height={a.h} fill={a.fill || 'transparent'} stroke={stroke} strokeWidth={a.width || 2} />;
    case 'ellipse': return <ellipse {...common} cx={a.x + a.w / 2} cy={a.y + a.h / 2} rx={Math.abs(a.w / 2)} ry={Math.abs(a.h / 2)} fill={a.fill || 'transparent'} stroke={stroke} strokeWidth={a.width || 2} />;
    case 'line': case 'arrow':
      return (
        <g {...common}>
          <line x1={a.x1} y1={a.y1} x2={a.x2} y2={a.y2} stroke="transparent" strokeWidth={Math.max(8, (a.width || 2) * 3)} />
          <line x1={a.x1} y1={a.y1} x2={a.x2} y2={a.y2} stroke={stroke} strokeWidth={a.width || 2} strokeLinecap="round" />
          {a.type === 'arrow' && <polygon points={arrowHead(a).map((p) => p.join(',')).join(' ')} fill={stroke} stroke={stroke} strokeWidth="1" />}
        </g>
      );
    case 'text': {
      const lh = a.size * 1.2;
      return (
        <g {...common} transform={`translate(${a.x} ${a.y}) rotate(${a.rot || 0}) scale(1 -1)`}>
          <text x="0" y="0" fontSize={a.size} fill={stroke} fontFamily="Helvetica, Arial, 'Noto Sans', sans-serif" style={{ whiteSpace: 'pre' }}>
            {textLines(a).map((l, i) => <tspan key={i} x="0" y={i * lh}><Keep>{l || ' '}</Keep></tspan>)}
          </text>
        </g>
      );
    }
    case 'note':
      return (
        <g {...common}>
          <title><Keep>{a.text}</Keep></title>
          <rect x={a.x} y={a.y} width="20" height="20" rx="3" fill={a.color || '#ffd400'} stroke="#735a00" strokeWidth="0.8" />
          {[0, 1, 2].map((k) => <line key={k} x1={a.x + 4} x2={a.x + (k === 2 ? 11 : 16)} y1={a.y + 14 - k * 4} y2={a.y + 14 - k * 4} stroke="#735a00" strokeWidth="1" />)}
        </g>
      );
    case 'stamp': {
      const bw = Math.max(1.5, a.h * 0.07);
      return (
        <g {...common} transform={`translate(${a.x} ${a.y}) rotate(${a.rot || 0})`}>
          <rect x={bw / 2} y={bw / 2} width={a.w - bw} height={a.h - bw} fill="transparent" stroke={stroke} strokeWidth={bw} rx={bw} />
          <text transform="scale(1 -1)" x={a.w / 2} y={-a.h * 0.3} fontSize={a.h * 0.5} fill={stroke} textAnchor="middle" fontWeight="700" fontFamily="Helvetica, Arial, sans-serif"><Keep>{a.text}</Keep></text>
        </g>
      );
    }
    case 'image':
      return (
        <g {...common} transform={`translate(${a.x} ${a.y}) rotate(${a.rot || 0}) scale(1 -1)`}>
          <image href={pngUrl(a.png)} x="0" y={-a.h} width={a.w} height={a.h} preserveAspectRatio="none" />
        </g>
      );
    case 'redact':
      return <rect data-rid={a.id} className={`of-pdf-redact${selected ? ' sel' : ''}`} x={a.x} y={a.y} width={a.w} height={a.h} />;
    default: return null;
  }
}

/* ---------------- a link service that is enough for pdf.js' annotation layer ---------------- */
export function makeLinkService({ onDest, onPage, onNamed, openLink }) {
  return {
    externalLinkEnabled: true,
    externalLinkTarget: 0,
    externalLinkRel: 'noopener noreferrer nofollow',
    isInPresentationMode: false,
    get pagesCount() { return 0; },
    get page() { return 1; },
    set page(v) { onPage(v - 1); },
    rotation: 0,
    addLinkAttributes(link, url) {
      link.href = url;
      link.title = url;
      link.rel = 'noopener noreferrer nofollow';
      link.onclick = (e) => { e.preventDefault(); openLink(url); return false; };
    },
    getDestinationHash: () => '#',
    getAnchorUrl: () => '#',
    goToDestination: (dest) => onDest(dest),
    goToPage: (n) => onPage(n - 1),
    navigateTo: (dest) => onDest(dest),
    executeNamedAction: (a) => onNamed(a),
    executeSetOCGState: () => {},
    setHash: () => {},
    cachePageRef: () => {},
    isPageVisible: () => true,
    isPageCached: () => true,
  };
}

/* ---------------- render scheduler: two at a time, nearest first ---------------- */
function makeScheduler(slots = 2) {
  const queue = new Map(); let running = 0; let seq = 0;
  const pump = () => {
    while (running < slots && queue.size) {
      let best = null;
      for (const [id, job] of queue) if (!best || job.prio() < best[1].prio()) best = [id, job];
      queue.delete(best[0]);
      running++;
      let finished = false;
      const done = () => { if (finished) return; finished = true; running--; pump(); };
      best[1].start(done);
    }
  };
  return {
    request(prio, start) {
      const id = ++seq;
      queue.set(id, { prio, start });
      queue.size && Promise.resolve().then(pump);
      return () => queue.delete(id);
    },
  };
}

/* ---------------- one page ---------------- */
const PageView = memo(function PageView({
  mp, proxy, doc, scale, renderScale, viewRot, near, visible, defaultSize, hits, activeHit, annots, redacts, selectedId, tool, editable, renderForms,
  linkService, onRegister, onPointer, onTextLayer, scheduler, prioOf, onRendered, onContextMenu,
}) {
  const rot = pageRotation(proxy, mp, viewRot);
  const vp = useMemo(() => (proxy ? proxy.getViewport({ scale: renderScale * CSS_UNITS, rotation: rot }) : null), [proxy, renderScale, rot]);
  const box = useRef(null);
  const canvas = useRef(null);
  const textDiv = useRef(null);
  const annDiv = useRef(null);
  const svg = useRef(null);
  const tl = useRef(null);
  const [rendered, setRendered] = useState(null); // key of the rendered state
  const [draft, setDraft] = useState(null);
  const k = scale / renderScale;
  useLayoutEffect(() => { const c = canvas.current; if (c && !c.dataset.init) { c.dataset.init = '1'; c.width = 0; c.height = 0; } }, []);

  // layout size at the current zoom (placeholder: size of the first page)
  let W; let H;
  if (vp) { W = vp.width * k; H = vp.height * k; } else {
    const [dw, dh] = defaultSize || [612, 792];
    const turned = ((mp.rot || 0) + viewRot) % 180 !== 0;
    W = (turned ? dh : dw) * scale * CSS_UNITS; H = (turned ? dw : dh) * scale * CSS_UNITS;
  }

  useEffect(() => { onRegister(mp.id, box.current, vp, k); return () => onRegister(mp.id, null, null, 1); }, [mp.id, vp, k, onRegister]);

  // render canvas + layers when near the window (and again after zoom/rotation)
  useEffect(() => {
    if (!near || !proxy || !vp) return undefined;
    const key = `${renderScale}:${rot}:${renderForms}`;
    if (rendered === key) return undefined;
    const lib = pdfjsLib();
    let cancelled = false; let task = null; let release = null;
    const cancelJob = scheduler.request(() => prioOf(mp.id), (done) => {
      release = done;
      if (cancelled) { done(); return; }
      let dpr = Math.min(window.devicePixelRatio || 1, 2.5);
      const base = vp.width * vp.height;
      if (base * dpr * dpr > MAX_CANVAS) dpr = Math.max(0.5, Math.sqrt(MAX_CANVAS / base));
      const cvp = proxy.getViewport({ scale: renderScale * CSS_UNITS * dpr, rotation: rot });
      const off = document.createElement('canvas');
      off.width = Math.floor(cvp.width); off.height = Math.floor(cvp.height);
      task = proxy.render({ canvasContext: off.getContext('2d', { alpha: false }), viewport: cvp, annotationMode: renderForms ? lib.AnnotationMode.ENABLE_FORMS : lib.AnnotationMode.ENABLE, background: '#ffffff' });
      task.promise.then(async () => {
        if (cancelled) return;
        const c = canvas.current;
        if (!c) return;
        c.width = off.width; c.height = off.height;
        c.getContext('2d', { alpha: false }).drawImage(off, 0, 0);
        off.width = 0; off.height = 0;
        done(); release = null;
        onRendered(mp.id, c.width * c.height, () => { if (canvas.current) { canvas.current.width = 0; canvas.current.height = 0; } setRendered(null); });
        // text layer
        if (tl.current) { try { tl.current.cancel(); } catch { /* done */ } }
        textDiv.current.replaceChildren();
        const layer = new lib.TextLayer({ textContentSource: proxy.streamTextContent({ includeMarkedContent: true, disableNormalization: true }), container: textDiv.current, viewport: vp });
        tl.current = layer;
        try { await layer.render(); } catch { return; }
        if (cancelled) return;
        const end = document.createElement('div'); end.className = 'endOfContent'; textDiv.current.append(end);
        onTextLayer(mp.id, layer);
        // links + form fields
        annDiv.current.replaceChildren();
        try {
          const annotations = await proxy.getAnnotations({ intent: 'display' });
          if (cancelled) return;
          if (annotations.length) {
            const al = new lib.AnnotationLayer({ div: annDiv.current, page: proxy, viewport: vp.clone({ dontFlip: true }), accessibilityManager: null, annotationCanvasMap: null, annotationEditorUIManager: null, structTreeLayer: null });
            await al.render({ annotations, viewport: vp.clone({ dontFlip: true }), linkService, renderForms, annotationStorage: doc.annotationStorage, imageResourcesPath: '', enableScripting: false, hasJSActions: false, fieldObjects: null, downloadManager: null });
          }
        } catch { /* annotations are optional */ }
        if (!cancelled) setRendered(key);
      }).catch(() => { if (release) { release(); release = null; } });
    });
    return () => { cancelled = true; cancelJob(); try { if (task) task.cancel(); } catch { /* done */ } if (release) { release(); release = null; } };
  }, [near, proxy, vp, renderScale, rot, renderForms]); // eslint-disable-line react-hooks/exhaustive-deps

  // search hits inside the text layer
  useEffect(() => {
    const layer = tl.current;
    if (!layer || !rendered) return;
    applyHits(layer, hits || [], activeHit);
  }, [hits, activeHit, rendered]);

  const w = vp ? vp.width : W; const h = vp ? vp.height : H;
  const toPdf = (e) => { const r = svg.current.getBoundingClientRect(); return vp.convertToPdfPoint((e.clientX - r.left) / k, (e.clientY - r.top) / k); };
  const drawing = editable && DRAW_TOOLS.has(tool);
  const placing = editable && PLACE_TOOLS.has(tool);
  const live = editable && (drawing || placing || tool === 'select');

  const down = (e) => {
    if (!vp || e.button > 0) return;
    const pt = toPdf(e);
    const handle = e.target.getAttribute && e.target.getAttribute('data-handle');
    if (handle) { e.stopPropagation(); onPointer({ kind: handle === 'end1' || handle === 'end2' ? handle : 'scale', page: mp.id, id: e.target.getAttribute('data-for'), pt, e, toPdf, rot }); return; }
    const rid = e.target.closest && e.target.closest('[data-rid]');
    if (rid && (tool === 'select' || tool === 'redact' || tool === 'text-select')) { onPointer({ kind: 'redact-pick', page: mp.id, id: rid.getAttribute('data-rid'), pt, e, toPdf }); return; }
    const aid = e.target.closest && e.target.closest('[data-aid]');
    if (aid && (tool === 'select' || tool === 'text-select')) { onPointer({ kind: 'grab', page: mp.id, id: aid.getAttribute('data-aid'), pt, e, toPdf, rot }); return; }
    if (placing) { e.preventDefault(); onPointer({ kind: 'place', page: mp.id, pt, rot, e }); return; }
    if (!drawing) { onPointer({ kind: 'blank', page: mp.id }); return; }
    e.preventDefault();
    svg.current.setPointerCapture(e.pointerId);
    setDraft({ tool, start: pt, pts: [pt[0], pt[1]], cur: pt });
  };
  const move = (e) => {
    if (!draft) return;
    const pt = toPdf(e);
    if (draft.tool === 'ink') {
      const n = draft.pts.length;
      if (Math.hypot(pt[0] - draft.pts[n - 2], pt[1] - draft.pts[n - 1]) < 1.2 / scale) return;
      setDraft({ ...draft, pts: [...draft.pts, pt[0], pt[1]] });
    } else setDraft({ ...draft, cur: pt });
  };
  const up = () => {
    if (!draft) return;
    const d = draft; setDraft(null);
    onPointer({ kind: 'drawn', page: mp.id, draft: d });
  };

  const draftShape = draft ? draftToAnnot(draft, {}) : null;
  const sel = selectedId ? (annots || []).find((a) => a.id === selectedId) : null;
  const sb = sel ? bbox(sel) : null;
  const hr = 5.5 / (renderScale * CSS_UNITS); // handle radius in page units (≈ 5.5 px on screen)
  const corners = sb ? [[sb.x - 2, sb.y - 2, 'nesw'], [sb.x + sb.w + 2, sb.y - 2, 'nwse'], [sb.x + sb.w + 2, sb.y + sb.h + 2, 'nesw'], [sb.x - 2, sb.y + sb.h + 2, 'nwse']] : [];
  const resizable = sel && editable && !MARKUP.includes(sel.type) && sel.type !== 'note' && sel.type !== 'line' && sel.type !== 'arrow';

  return (
    <div className={`of-pdf-page${rendered ? ' ready' : ''}`} ref={box} data-page={mp.id} style={{ width: W, height: H }}
      onContextMenu={(e) => { if (!vp || !onContextMenu) return; const aid = e.target.closest && e.target.closest('[data-aid]'); onContextMenu({ e, page: mp.id, pt: toPdf(e), aid: aid ? aid.getAttribute('data-aid') : null }); }}>
      <div className="of-pdf-layers" style={{ width: w, height: h, '--scale-factor': renderScale * CSS_UNITS, transform: k !== 1 ? `scale(${k})` : undefined }}>
        <canvas ref={canvas} className="of-pdf-canvas" style={{ width: w, height: h }} aria-hidden="true" />
        {!rendered && <div className="of-pdf-page-wait" aria-hidden="true" />}
        <div ref={textDiv} className="textLayer" />
        <div ref={annDiv} className={`annotationLayer${live ? ' of-pdf-passive' : ''}`} />
        {vp && (visible || (annots && annots.length) || (redacts && redacts.length)) && (
          <svg ref={svg} className={`of-pdf-overlay${drawing || placing ? ' draw' : ''}${tool === 'select' ? ' select' : ''}`} width={w} height={h}
            onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerCancel={() => setDraft(null)} style={{ cursor: drawing ? 'crosshair' : placing ? 'copy' : undefined }}>
            <defs>
              <pattern id={`of-hatch-${mp.id}`} width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><line x1="0" y1="0" x2="0" y2="6" className="of-pdf-hatch" /></pattern>
            </defs>
            <g transform={`matrix(${vp.transform.join(' ')})`}>
              {(annots || []).map((a) => <AnnotShape key={a.id} a={a} selected={a.id === selectedId} />)}
              {(redacts || []).map((r) => <g key={r.id}><AnnotShape a={{ ...r, type: 'redact' }} selected={r.id === selectedId} /><rect x={r.x} y={r.y} width={r.w} height={r.h} fill={`url(#of-hatch-${mp.id})`} pointerEvents="none" /></g>)}
              {draftShape && <AnnotShape a={{ ...draftShape, id: 'draft' }} />}
              {sb && (
                <g className="of-pdf-selbox">
                  <rect x={sb.x - 2} y={sb.y - 2} width={sb.w + 4} height={sb.h + 4} fill="none" vectorEffect="non-scaling-stroke" />
                  {resizable && corners.map(([x, y, cur], i) => <circle key={i} data-handle="scale" data-for={sel.id} cx={x} cy={y} r={hr} style={{ cursor: `${cur}-resize` }} />)}
                  {editable && (sel.type === 'line' || sel.type === 'arrow') && <>
                    <circle data-handle="end1" data-for={sel.id} cx={sel.x1} cy={sel.y1} r={hr} style={{ cursor: 'move' }} />
                    <circle data-handle="end2" data-for={sel.id} cx={sel.x2} cy={sel.y2} r={hr} style={{ cursor: 'move' }} />
                  </>}
                </g>
              )}
            </g>
          </svg>
        )}
      </div>
    </div>
  );
});

/** a finished draft → annotation (without id) */
export function draftToAnnot(d, opts) {
  const color = opts.color || d.color || '#e5364a';
  const width = opts.width || d.width || 2;
  const [x1, y1] = d.start; const [x2, y2] = d.cur || d.start;
  switch (d.tool) {
    case 'ink': return { type: 'ink', paths: [d.pts], color, width };
    case 'rect': case 'ellipse': case 'redact': return { type: d.tool, x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), color, width };
    case 'line': case 'arrow': return { type: d.tool, x1, y1, x2, y2, color, width };
    default: return null;
  }
}

/* ---------------- search highlights in a pdf.js text layer ---------------- */
/** hits: [{ parts: [{ i (text item), a, b }] }] – see search.js */
function applyHits(layer, hits, active) {
  const divs = layer.textDivs; const strs = layer.textContentItemsStr;
  divs.forEach((d, i) => { if (d.dataset.hit) { d.textContent = strs[i]; delete d.dataset.hit; } });
  if (!hits.length) return;
  const per = new Map();
  hits.forEach((h, hi) => h.parts.forEach((p) => { if (!per.has(p.i)) per.set(p.i, []); per.get(p.i).push({ a: p.a, b: p.b, active: hi === active }); }));
  let act = null;
  for (const [i, ranges] of per) {
    const d = divs[i]; const s = strs[i];
    if (!d || s == null) continue;
    d.dataset.hit = '1';
    d.replaceChildren();
    let pos = 0;
    ranges.sort((x, y) => x.a - y.a).forEach((r) => {
      if (r.a > pos) d.append(document.createTextNode(s.slice(pos, r.a)));
      const m = document.createElement('span');
      m.className = `of-pdf-hit${r.active ? ' active' : ''}`;
      m.textContent = s.slice(Math.max(r.a, pos), r.b);
      if (r.active && !act) act = m;
      d.append(m); pos = Math.max(pos, r.b);
    });
    if (pos < s.length) d.append(document.createTextNode(s.slice(pos)));
  }
  // scroll only the page area (scrollIntoView would also move Hostwerk's own page)
  const sc = act && act.closest('.of-pdf-scroll');
  if (sc) {
    const r = act.getBoundingClientRect(); const s = sc.getBoundingClientRect();
    if (r.top < s.top + 40 || r.bottom > s.bottom - 40) sc.scrollTop += r.top - s.top - s.height / 2;
    if (r.left < s.left || r.right > s.right) sc.scrollLeft += r.left - s.left - s.width / 2;
  }
}

/* ---------------- the scrolling page area ---------------- */
export default function Viewer({
  pages, proxies, docs, scale, viewRot, scrollMode, current, onCurrent, annots, redactions, selectedId, tool, editable, renderForms,
  hitsByPage, activeHit, linkService, onPointer, onTextSelection, scrollRef, night, onZoom, jump, needProxy, defaultSize, onContextMenu, apiRef, onFirstPaint,
}) {
  const els = useRef(new Map());
  const [visible, setVisible] = useState(() => new Set());
  const [near, setNear] = useState(() => new Set());
  const layers = useRef(new Map());
  const local = useRef(null);
  const el = scrollRef || local;
  const register = useRef((id, node, vp, k) => { if (node) els.current.set(id, { node, vp, k }); else els.current.delete(id); }).current;
  const onTextLayer = useRef((id, layer) => layers.current.set(id, layer)).current;
  const scheduler = useMemo(() => makeScheduler(2), []);
  const centre = useRef(0);
  const prioOf = useRef((id) => { const e = els.current.get(id); if (!e) return 1e9; const horiz = scrollMode === 'horizontal'; const s = horiz ? e.node.offsetLeft + e.node.offsetWidth / 2 : e.node.offsetTop + e.node.offsetHeight / 2; return Math.abs(s - centre.current); });
  prioOf.current = (id) => { const e = els.current.get(id); if (!e) return 1e9; const horiz = scrollMode === 'horizontal'; const s = horiz ? e.node.offsetLeft + e.node.offsetWidth / 2 : e.node.offsetTop + e.node.offsetHeight / 2; return Math.abs(s - centre.current); };
  const prio = useRef((id) => prioOf.current(id)).current;

  // sharp re-render ~160 ms after the last zoom step; until then the bitmaps are stretched
  const [renderScale, setRenderScale] = useState(scale);
  useEffect(() => { if (scale === renderScale) return undefined; const tm = setTimeout(() => setRenderScale(scale), 160); return () => clearTimeout(tm); }, [scale]); // eslint-disable-line react-hooks/exhaustive-deps

  // pixel budget of rendered canvases (LRU, farthest first)
  const renderedPx = useRef(new Map());
  const visibleRef = useRef(visible); visibleRef.current = visible;
  const firstPaint = useRef(onFirstPaint); firstPaint.current = onFirstPaint;
  const onRendered = useRef((id, px, evict) => {
    if (firstPaint.current) { const f = firstPaint.current; firstPaint.current = null; requestAnimationFrame(() => f()); }
    const m = renderedPx.current;
    m.delete(id); m.set(id, { px, evict });
    let total = 0; for (const v of m.values()) total += v.px;
    if (total <= PIXEL_BUDGET) return;
    const order = [...m.keys()].filter((x) => !visibleRef.current.has(x)).sort((a, b) => prioOf.current(b) - prioOf.current(a));
    for (const x of order) { if (total <= PIXEL_BUDGET * 0.8) break; const v = m.get(x); total -= v.px; m.delete(x); try { v.evict(); } catch { /* gone */ } }
  }).current;

  const recompute = useRef(null);
  recompute.current = () => {
    const sc = el.current; if (!sc) return;
    const horiz = scrollMode === 'horizontal';
    const a = horiz ? sc.scrollLeft : sc.scrollTop; const len = horiz ? sc.clientWidth : sc.clientHeight;
    centre.current = a + len / 2;
    const vis = new Set(); const nr = new Set(); let cur = null; let best = Infinity;
    for (const p of pages) {
      const e = els.current.get(p.id); if (!e) continue;
      const s = horiz ? e.node.offsetLeft : e.node.offsetTop; const size = horiz ? e.node.offsetWidth : e.node.offsetHeight;
      if (s + size >= a - len * 1.2 && s <= a + len * 2.2) nr.add(p.id);
      if (s + size >= a && s <= a + len) vis.add(p.id);
      const d = Math.abs(s + Math.min(size, len) / 2 - (a + len / 3));
      if (s <= a + len && s + size >= a && d < best) { best = d; cur = p.id; }
    }
    setVisible((old) => (old.size === vis.size && [...vis].every((x) => old.has(x)) ? old : vis));
    setNear((old) => (old.size === nr.size && [...nr].every((x) => old.has(x)) ? old : nr));
    if (needProxy) for (const id of nr) { const p = pages.find((x) => x.id === id); if (p && !proxies.get(`${p.src}:${p.index}`)) needProxy(p); }
    // LRU: visible pages are the most recent
    for (const id of vis) { const m = renderedPx.current; if (m.has(id)) { const v = m.get(id); m.delete(id); m.set(id, v); } }
    if (cur != null && scrollMode !== 'page') onCurrent(cur, false);
  };

  // zoom: Ctrl+wheel (smooth on touchpads), anchored at the mouse
  const anchor = useRef(null);
  const onZoomRef = useRef(onZoom); onZoomRef.current = onZoom;
  const scaleRef = useRef(scale);
  useEffect(() => {
    const sc = el.current; if (!sc) return undefined;
    let raf = 0;
    const h = () => { cancelAnimationFrame(raf); raf = requestAnimationFrame(() => recompute.current()); };
    sc.addEventListener('scroll', h, { passive: true });
    const ro = new ResizeObserver(h); ro.observe(sc);
    let acc = 1; let zraf = 0; let last = null;
    const wheel = (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      const dy = e.deltaMode === 1 ? e.deltaY * 16 : e.deltaY;
      acc *= Math.exp(-Math.max(-60, Math.min(60, dy)) * 0.0022);
      last = { x: e.clientX, y: e.clientY };
      if (!zraf) zraf = requestAnimationFrame(() => { zraf = 0; const f = acc; acc = 1; anchor.current = { ...last }; onZoomRef.current(f); });
    };
    sc.addEventListener('wheel', wheel, { passive: false });
    return () => { sc.removeEventListener('scroll', h); sc.removeEventListener('wheel', wheel); ro.disconnect(); cancelAnimationFrame(raf); cancelAnimationFrame(zraf); };
  }, [el]);
  useLayoutEffect(() => {
    const sc = el.current; const old = scaleRef.current; scaleRef.current = scale;
    if (!sc || old === scale || scrollMode === 'page') { anchor.current = null; return; }
    const r = sc.getBoundingClientRect();
    const ax = anchor.current ? anchor.current.x - r.left : sc.clientWidth / 2;
    const ay = anchor.current ? anchor.current.y - r.top : sc.clientHeight / 2;
    anchor.current = null;
    const ratio = scale / old;
    sc.scrollLeft = (sc.scrollLeft + ax) * ratio - ax;
    sc.scrollTop = (sc.scrollTop + ay) * ratio - ay;
  }, [scale]); // eslint-disable-line react-hooks/exhaustive-deps
  useLayoutEffect(() => { recompute.current(); }, [pages, scale, viewRot, scrollMode, proxies]);

  // jump to a page (jump = { id, n } changes when the editor asks)
  useLayoutEffect(() => {
    if (!jump || scrollMode === 'page') return;
    const e = els.current.get(jump.id); const sc = el.current;
    if (!e || !sc) return;
    if (scrollMode === 'horizontal') sc.scrollLeft = e.node.offsetLeft - 12;
    else sc.scrollTop = e.node.offsetTop - 12 + (jump.offset || 0) * (e.k || 1);
  }, [jump, scrollMode]); // eslint-disable-line react-hooks/exhaustive-deps

  // text selection → markup
  const mouseUp = () => {
    if (!editable || !MARKUP.includes(tool)) return;
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) return;
    const byPage = selectionQuads(sel);
    const text = sel.toString();
    sel.removeAllRanges();
    for (const [id, qs] of byPage) onTextSelection(id, mergeQuads(qs), text);
  };
  /** selection rectangles per page in PDF space (also used by the context menu) */
  function selectionQuads(sel) {
    const rects = [...sel.getRangeAt(0).getClientRects()];
    const byPage = new Map();
    for (const [id, { node, vp, k }] of els.current) {
      if (!vp) continue;
      const r = node.getBoundingClientRect();
      for (const cr of rects) {
        const cx = cr.left + cr.width / 2; const cy = cr.top + cr.height / 2;
        if (cx < r.left || cx > r.right || cy < r.top || cy > r.bottom || cr.width < 1) continue;
        const [ax, ay] = vp.convertToPdfPoint((cr.left - r.left) / k, (cr.top - r.top) / k);
        const [bx, by] = vp.convertToPdfPoint((cr.right - r.left) / k, (cr.bottom - r.top) / k);
        if (!byPage.has(id)) byPage.set(id, []);
        byPage.get(id).push({ x: Math.min(ax, bx), y: Math.min(ay, by), w: Math.abs(bx - ax), h: Math.abs(by - ay) });
      }
    }
    return byPage;
  }
  if (apiRef) apiRef.current = { selectionQuads, scrollEl: () => el.current };

  const shown = scrollMode === 'page' ? pages.filter((p) => p.id === current) : pages;
  return (
    <div ref={el} className={`of-pdf-scroll of-pdfjs mode-${scrollMode}${night ? ' night' : ''}${tool === 'hand' ? ' hand' : ''}${MARKUP.includes(tool) ? ' markup' : ''}`} onMouseUp={mouseUp} tabIndex={-1}
      onPointerDown={(e) => {
        if (!(tool === 'hand' || e.button === 1) || (e.button !== 0 && e.button !== 1)) return;
        e.preventDefault();
        const sc = el.current; const sx = e.clientX; const sy = e.clientY; const l = sc.scrollLeft; const t = sc.scrollTop;
        sc.classList.add('grabbing');
        const mv = (ev) => { sc.scrollLeft = l - (ev.clientX - sx); sc.scrollTop = t - (ev.clientY - sy); };
        const end = () => { sc.classList.remove('grabbing'); window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', end); };
        window.addEventListener('pointermove', mv); window.addEventListener('pointerup', end);
      }}>
      <div className="of-pdf-pages">
        {shown.map((mp) => {
          const proxy = proxies.get(`${mp.src}:${mp.index}`);
          return (
            <PageView key={mp.id} mp={mp} proxy={proxy} doc={docs[mp.src]} scale={scale} renderScale={renderScale} viewRot={viewRot} defaultSize={defaultSize}
              near={scrollMode === 'page' || near.has(mp.id)} visible={scrollMode === 'page' || visible.has(mp.id)} hits={hitsByPage && hitsByPage.get(mp.id)} activeHit={activeHit && activeHit.page === mp.id ? activeHit.i : -1}
              annots={annots[mp.id]} redacts={redactions && redactions[mp.id]} selectedId={selectedId} tool={tool} editable={editable} renderForms={renderForms && mp.src === 0}
              linkService={linkService} onRegister={register} onPointer={onPointer} onTextLayer={onTextLayer} scheduler={scheduler} prioOf={prio} onRendered={onRendered} onContextMenu={onContextMenu} />
          );
        })}
      </div>
    </div>
  );
}
