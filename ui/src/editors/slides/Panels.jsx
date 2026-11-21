/* Side and alternate views: slide thumbnails with sections, slide sorter, outline, notes page, format pane, comments, animations. */
import { useEffect, useMemo, useRef, useState, createElement as h } from 'react';
import { ChevronDown, ChevronRight, EyeOff, Plus, Trash2, ArrowUp, ArrowDown, MessageSquare, Sparkles, Lock, ArrowRightLeft } from 'lucide-react';
import { t } from '@hostwerk/i18n';
import SlideView from './render/SlideView.jsx';
import { S, emuToCm, cmToEmu, emuToIn, inToEmu } from './units.js';
import { sectionGroups, slideTitle, bodyText, bodyFromText, findEl, updateEls, updateSlide, walk, uid } from './model.js';
import { ColorPop, Num } from './controls.jsx';
import { PaintBucket, PenLine } from 'lucide-react';
import { Raw } from '../_shared/keep.jsx';

/* ---------------------------------------------------------------- lazy slide picture (thumbnails, sorter) */
// One shared IntersectionObserver: a slide is rendered when it comes near the visible area (200-slide decks open fast)
let IO = null; const IO_CB = new Map();
function observe(el, cb) {
  if (typeof IntersectionObserver === 'undefined') { cb(); return () => {}; }
  if (!IO) IO = new IntersectionObserver((list) => { for (const x of list) if (x.isIntersecting) { const f = IO_CB.get(x.target); if (f) { IO_CB.delete(x.target); IO.unobserve(x.target); f(); } } }, { rootMargin: '600px 0px' });
  IO_CB.set(el, cb); IO.observe(el);
  return () => { IO_CB.delete(el); if (IO) IO.unobserve(el); };
}
export function LazySlide({ eager, w, h: hh, children }) {
  const ref = useRef(null);
  const [on, setOn] = useState(!!eager);
  useEffect(() => { if (on || !ref.current) return undefined; return observe(ref.current, () => setOn(true)); }, [on]);
  return on ? children : <div ref={ref} className="of-slides-thumbph" style={{ width: w, height: hh }} />;
}

/* ---------------------------------------------------------------- thumbnails */
export function Thumbs({ doc, curId, selSlides, onSelect, onReorder, onContext, onAdd, onSectionContext, collapsed, setCollapsed, width = 168 }) {
  const W = S(doc.size.w); const H = S(doc.size.h);
  // panel width − border − padding (4 + 8) − number column (20) − gap (6) − outline room (2) = picture width
  const scale = (width - 41) / W;
  const [dropAt, setDropAt] = useState(null);
  const [dragIds, setDragIds] = useState(null);
  const groups = sectionGroups(doc);
  const listRef = useRef(null);
  const curIdx = doc.slides.findIndex((s) => s.id === curId);
  useEffect(() => { const el = listRef.current && listRef.current.querySelector(`[data-slide="${curId}"]`); if (el && el.scrollIntoView) el.scrollIntoView({ block: 'nearest' }); }, [curId]);
  const onDrop = (e) => { e.preventDefault(); const ids = (e.dataTransfer.getData('text/x-of-slides') || '').split(',').filter(Boolean); if (ids.length && dropAt != null) onReorder(ids, dropAt); setDropAt(null); setDragIds(null); };
  return (
    <div className="of-slides-thumbs" ref={listRef} role="listbox" aria-label={t('Slides')} aria-multiselectable="true" onDragOver={(e) => e.preventDefault()} onDrop={onDrop} onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget)) setDropAt(null); }}>
      {groups.map((g, gi) => (
        <div key={g.section ? g.section.id : `none${gi}`} className="of-slides-sec">
          {g.section && (
            <button type="button" className="of-slides-sechead" onClick={() => setCollapsed({ ...collapsed, [g.section.id]: !collapsed[g.section.id] })} onContextMenu={(e) => onSectionContext(e, g.section)}>
              {collapsed[g.section.id] ? <ChevronRight size={14} /> : <ChevronDown size={14} />}
              {h('span', { className: 'grow' }, g.section.name)}<span className="badge">{g.slides.length}</span>
            </button>
          )}
          {!(g.section && collapsed[g.section.id]) && g.slides.map((s) => {
            const i = doc.slides.indexOf(s);
            const selected = selSlides.includes(s.id);
            return (
              <div key={s.id} data-slide={s.id} role="option" aria-selected={selected} tabIndex={s.id === curId ? 0 : -1}
                className={`of-slides-thumb ${s.id === curId ? 'cur' : ''} ${selected ? 'sel' : ''} ${s.hidden ? 'hidden' : ''} ${dropAt === i ? 'drop-before' : ''} ${dropAt === i + 1 && i === doc.slides.length - 1 ? 'drop-after' : ''} ${dragIds && dragIds.includes(s.id) ? 'dragging' : ''}`}
                draggable onDragStart={(e) => { const ids = selected ? selSlides : [s.id]; e.dataTransfer.setData('text/x-of-slides', ids.join(',')); e.dataTransfer.effectAllowed = 'move'; setDragIds(ids); }} onDragEnd={() => { setDragIds(null); setDropAt(null); }}
                onDragOver={(e) => { e.preventDefault(); const r = e.currentTarget.getBoundingClientRect(); setDropAt(e.clientY < r.top + r.height / 2 ? i : i + 1); }}
                onClick={(e) => onSelect(s.id, e)} onContextMenu={(e) => onContext(e, s.id)}>
                <span className="of-slides-thumbnum"><span>{i + 1}</span><ThumbMarks s={s} /></span>
                <Raw as="div" className="of-slides-thumbimg" style={{ width: W * scale, height: H * scale }} title={slideTitle(s) || undefined}><LazySlide eager={Math.abs(i - curIdx) < 12} w={W * scale} h={H * scale}><SlideView doc={doc} slide={s} index={i} scale={scale} mode="thumb" /></LazySlide></Raw>
              </div>
            );
          })}
        </div>
      ))}
      {doc.slides.length === 0 && <div className="small muted of-slides-thumbempty">{t('No slides yet.')}</div>}
      <button type="button" className="btn sm ghost of-slides-addslide" onClick={() => onAdd()}><Plus size={14} />{t('New slide')}</button>
    </div>
  );
}

/** small markers under the slide number – each with a tooltip: hidden, animations, comments */
export function ThumbMarks({ s }) {
  const anims = (s.animations || []).length;
  const comments = (s.comments || []).filter((x) => !x.parent && !x.done).length;
  const mark = (Icon, tip, key) => <Raw as="span" key={key} className="of-slides-mark" title={tip} aria-label={tip} role="img"><Icon size={12} aria-hidden="true" /></Raw>;
  return [
    s.hidden && mark(EyeOff, t('Hidden in the slide show'), 'h'),
    anims > 0 && mark(Sparkles, anims === 1 ? t('1 animation') : t('{n} animations', { n: anims }), 'a'),
    comments > 0 && mark(MessageSquare, comments === 1 ? t('1 comment') : t('{n} comments', { n: comments }), 'c'),
  ];
}

/* ---------------------------------------------------------------- slide sorter */
export function Sorter({ doc, curId, selSlides, onSelect, onOpen, onReorder, onContext, zoom = 1 }) {
  const W = S(doc.size.w); const H = S(doc.size.h);
  const tw = 220 * zoom; const scale = tw / W;
  const [dropAt, setDropAt] = useState(null);
  return (
    <div className="of-slides-sorter" onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); const ids = (e.dataTransfer.getData('text/x-of-slides') || '').split(',').filter(Boolean); if (ids.length && dropAt != null) onReorder(ids, dropAt); setDropAt(null); }}>
      {sectionGroups(doc).map((g, gi) => (
        <div key={g.section ? g.section.id : `n${gi}`} className="of-slides-sortsec">
          {g.section && h('div', { className: 'of-slides-sortsechead' }, g.section.name)}
          <div className="of-slides-sortgrid">
            {g.slides.map((s) => {
              const i = doc.slides.indexOf(s);
              return (
                <div key={s.id} className={`of-slides-sortitem ${s.id === curId ? 'cur' : ''} ${selSlides.includes(s.id) ? 'sel' : ''} ${s.hidden ? 'hidden' : ''} ${dropAt === i ? 'drop-before' : ''}`} draggable
                  onDragStart={(e) => e.dataTransfer.setData('text/x-of-slides', (selSlides.includes(s.id) ? selSlides : [s.id]).join(','))}
                  onDragOver={(e) => { e.preventDefault(); const r = e.currentTarget.getBoundingClientRect(); setDropAt(e.clientX < r.left + r.width / 2 ? i : i + 1); }}
                  onClick={(e) => onSelect(s.id, e)} onDoubleClick={() => onOpen(s.id)} onContextMenu={(e) => onContext(e, s.id)}>
                  <div className="of-slides-sortimg" style={{ width: W * scale, height: H * scale }}><LazySlide eager={i < 24} w={W * scale} h={H * scale}><SlideView doc={doc} slide={s} index={i} scale={scale} mode="thumb" /></LazySlide></div>
                  <div className="small row-tight of-slides-sortmeta"><b>{i + 1}</b><ThumbMarks s={s} />{s.transition && s.transition.type !== 'none' && <Raw as="span" className="of-slides-mark" role="img" title={t('Transition')} aria-label={t('Transition')}><ArrowRightLeft size={12} aria-hidden="true" /></Raw>}{s.transition && s.transition.advanceAfter ? <span className="muted">{Math.round(s.transition.advanceAfter / 100) / 10}s</span> : null}</div>
                </div>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ---------------------------------------------------------------- outline */
function OutlineSlide({ doc, s, i, change, readOnly, onOpen, spellcheck }) {
  const title = s.elements.find((e) => e.ph && /title/i.test(e.ph) && !/sub/i.test(e.ph));
  const bodyEl = s.elements.find((e) => e.ph && /^(body|subTitle)/.test(e.ph));
  const bodyStr = bodyEl ? bodyEl.text.paras.map((p) => `${'\t'.repeat(p.level || 0)}${(p.runs || []).map((r) => r.text).join('')}`).join('\n') : '';
  const setTitle = (v) => title && change((d) => updateEls(d, s.id, [title.id], (e) => ({ ...e, text: mergeText(e.text, v) })), `outline:t:${s.id}`);
  const setBody = (v) => bodyEl && change((d) => updateEls(d, s.id, [bodyEl.id], (e) => ({ ...e, text: mergeText(e.text, v) })), `outline:b:${s.id}`);
  return (
    <div className="of-slides-outitem">
      <button type="button" className="of-slides-outnum" onClick={() => onOpen(s.id)} title={t('Open slide')}>{i + 1}</button>
      <div className="grow stack-tight">
        {title ? <input className="of-slides-outtitle" value={bodyText(title.text)} readOnly={readOnly} spellCheck={spellcheck} aria-label={t('Slide title')} placeholder={t('Click to add title')} onChange={(e) => setTitle(e.target.value)} onKeyDown={(e) => e.stopPropagation()} />
          : h('div', { className: 'of-slides-outtitle muted' }, slideTitle(s) || t('(no title)'))}
        {bodyEl && <textarea className="of-slides-outbody" value={bodyStr} readOnly={readOnly} spellCheck={spellcheck} rows={Math.max(2, bodyStr.split('\n').length)} aria-label={t('Slide text')}
          onChange={(e) => setBody(e.target.value)}
          onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Tab') { e.preventDefault(); const ta = e.target; const st = ta.selectionStart; const v = ta.value; const ls = v.lastIndexOf('\n', st - 1) + 1; const nv = e.shiftKey ? (v[ls] === '\t' ? v.slice(0, ls) + v.slice(ls + 1) : v) : `${v.slice(0, ls)}\t${v.slice(ls)}`; setBody(nv); requestAnimationFrame(() => { ta.selectionStart = ta.selectionEnd = st + (nv.length - v.length); }); } }} />}
      </div>
    </div>
  );
}
/** replace text but keep run formatting of each paragraph's first run */
function mergeText(b, str) {
  const nb = bodyFromText(str);
  const paras = nb.paras.map((p, i) => {
    const old = b.paras[i] || b.paras[b.paras.length - 1] || { runs: [] };
    const { runs, ...pp } = old; // eslint-disable-line no-unused-vars
    const first = (old.runs || [])[0] || {};
    const { text, ...rp } = first; // eslint-disable-line no-unused-vars
    return { ...pp, level: p.level, runs: p.runs.length ? [{ ...rp, text: p.runs[0].text }] : [] };
  });
  return { ...b, paras };
}
export function Outline({ doc, change, readOnly, onOpen, spellcheck }) {
  return <div className="of-slides-outline">{doc.slides.map((s, i) => <OutlineSlide key={s.id} doc={doc} s={s} i={i} change={change} readOnly={readOnly} onOpen={onOpen} spellcheck={spellcheck} />)}</div>;
}

/* ---------------------------------------------------------------- notes page */
export function NotesPage({ doc, slide, index, change, readOnly, spellcheck }) {
  const W = S(doc.size.w);
  // the slide picture follows the paper's width (phones: the paper is narrower than 560 px)
  const paper = useRef(null);
  const [avail, setAvail] = useState(560);
  useEffect(() => {
    const el = paper.current;
    if (!el || typeof ResizeObserver === 'undefined') return undefined;
    const fit = () => { const cs = getComputedStyle(el); setAvail(Math.max(120, Math.min(560, el.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight)))); };
    fit();
    const ro = new ResizeObserver(fit); ro.observe(el);
    return () => ro.disconnect();
  }, []);
  const scale = avail / W;
  return (
    <div className="of-slides-notespage">
      <div className="of-slides-notespaper" ref={paper}>
        <div className="of-slides-notesslide"><SlideView doc={doc} slide={slide} index={index} scale={scale} mode="thumb" /></div>
        <textarea className="of-slides-notesbig" value={slide.notes || ''} readOnly={readOnly} spellCheck={spellcheck} placeholder={t('Click to add notes')} aria-label={t('Notes')}
          onKeyDown={(e) => e.stopPropagation()} onChange={(e) => change((d) => updateSlide(d, slide.id, { notes: e.target.value }), `notes:${slide.id}`)} />
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------- format pane */
export function FormatPane({ doc, slide, sel, change, units = 'cm', onClose, onCrop }) {
  const els = sel.map((id) => findEl(slide.elements, id)).filter(Boolean);
  const el = els[0];
  const toU = (e) => +(units === 'in' ? emuToIn(e) : emuToCm(e)).toFixed(2);
  const fromU = (v) => (units === 'in' ? inToEmu(v) : cmToEmu(v));
  const upd = (fn, key) => change((d) => updateEls(d, slide.id, sel, fn), key);
  if (!el) return <div className="of-slides-pane"><div className="of-slides-panehead"><b className="grow">{t('Format')}</b><button type="button" className="btn xs ghost" onClick={onClose}>{t('Close')}</button></div><p className="small muted">{t('Select an object on the slide to format it.')}</p></div>;
  const hasFill = el.type === 'shape' || el.type === 'image';
  const fill = el.fill;
  const line = el.line;
  return (
    <div className="of-slides-pane">
      <div className="of-slides-panehead"><b className="grow">{t('Format')}</b><button type="button" className="btn xs ghost" onClick={onClose}>{t('Close')}</button></div>
      <div className="of-slides-panesec">
        <div className="small muted">{t('Position and size')}</div>
        <div className="of-slides-grid2">
          <Num label={t('Left')} unit={units} step={0.1} value={toU(el.x)} onChange={(v) => upd((e) => ({ ...e, x: fromU(v) }))} />
          <Num label={t('Top')} unit={units} step={0.1} value={toU(el.y)} onChange={(v) => upd((e) => ({ ...e, y: fromU(v) }))} />
          <Num label={t('Width')} unit={units} step={0.1} min={0.01} value={toU(el.w)} onChange={(v) => upd((e) => { const w = fromU(v); return e.lockRatio || e.type === 'image' ? { ...e, w, h: Math.round(e.h * (w / e.w)) } : { ...e, w }; })} />
          <Num label={t('Height')} unit={units} step={0.1} min={0.01} value={toU(el.h)} onChange={(v) => upd((e) => { const hh = fromU(v); return e.lockRatio || e.type === 'image' ? { ...e, h: hh, w: Math.round(e.w * (hh / e.h)) } : { ...e, h: hh }; })} />
          <Num label={t('Rotation')} unit="°" value={el.rot || 0} onChange={(v) => upd((e) => ({ ...e, rot: ((v % 360) + 360) % 360 }))} />
        </div>
        <label className="row-tight small"><input type="checkbox" checked={!!el.lockRatio} onChange={(e) => upd((x) => ({ ...x, lockRatio: e.target.checked }))} />{t('Lock aspect ratio')}</label>
        <label className="row-tight small"><input type="checkbox" checked={!!el.locked} onChange={(e) => upd((x) => ({ ...x, locked: e.target.checked }))} /><Lock size={12} />{t('Lock position')}</label>
      </div>
      {hasFill && (
        <div className="of-slides-panesec">
          <div className="small muted">{t('Fill')}</div>
          <select className="input" aria-label={t('Fill')} value={!fill ? 'none' : fill.type} onChange={(e) => { const v = e.target.value; upd((x) => ({ ...x, fill: v === 'none' ? null : v === 'solid' ? { type: 'solid', color: (fill && fill.color) || 'accent1' } : v === 'grad' ? { type: 'grad', angle: 90, stops: [{ pos: 0, color: (fill && fill.color) || 'accent1' }, { pos: 100, color: 'accent1-50' }] } : x.fill })); }}>
            <option value="none">{t('No fill')}</option><option value="solid">{t('Solid fill')}</option><option value="grad">{t('Gradient fill')}</option>{fill && fill.type === 'image' && <option value="image">{t('Picture fill')}</option>}
          </select>
          {fill && fill.type === 'solid' && <div className="row-tight"><ColorPop icon={PaintBucket} title={t('Fill colour')} theme={doc.theme} value={fill.color} allowNone={false} onPick={(c) => upd((x) => ({ ...x, fill: { ...x.fill, color: c } }))} /><span className="small">{t('Transparency')}</span><Num label={t('Transparency')} unit="%" min={0} max={100} value={100 - Math.round(alphaOf(fill.color))} onChange={(v) => upd((x) => ({ ...x, fill: { ...x.fill, color: setAlpha(x.fill.color, 100 - v) } }))} /></div>}
          {fill && fill.type === 'grad' && (
            <div className="stack-tight">
              {fill.stops.map((s, i) => <div key={i} className="row-tight"><ColorPop icon={PaintBucket} title={t('Colour')} theme={doc.theme} value={s.color} allowNone={false} onPick={(c) => upd((x) => ({ ...x, fill: { ...x.fill, stops: x.fill.stops.map((y, k) => (k === i ? { ...y, color: c } : y)) } }))} /><Num label={t('Position')} unit="%" min={0} max={100} value={s.pos} onChange={(v) => upd((x) => ({ ...x, fill: { ...x.fill, stops: x.fill.stops.map((y, k) => (k === i ? { ...y, pos: v } : y)) } }))} /></div>)}
              <div className="row-tight"><span className="small">{t('Angle')}</span><Num label={t('Angle')} unit="°" value={fill.angle || 0} onChange={(v) => upd((x) => ({ ...x, fill: { ...x.fill, angle: v } }))} />
                <select className="input" aria-label={t('Type')} value={fill.kind || 'linear'} onChange={(e) => upd((x) => ({ ...x, fill: { ...x.fill, kind: e.target.value } }))}><option value="linear">{t('Linear')}</option><option value="radial">{t('Radial')}</option></select></div>
            </div>
          )}
        </div>
      )}
      {(el.type === 'shape' || el.type === 'line' || el.type === 'image') && (
        <div className="of-slides-panesec">
          <div className="small muted">{t('Line')}</div>
          <div className="row-tight">
            <ColorPop icon={PenLine} title={t('Line colour')} theme={doc.theme} value={line && line.color} noneLabel="No line" onPick={(c) => upd((x) => ({ ...x, line: c ? { width: 12700, dash: 'solid', ...(x.line || {}), color: c } : (x.type === 'line' ? x.line : null) }))} />
            <Num label={t('Line width')} unit="pt" step={0.25} min={0} value={line ? +S(line.width || 12700).toFixed(2) : 0} onChange={(v) => upd((x) => ({ ...x, line: v ? { color: 'tx1', dash: 'solid', ...(x.line || {}), width: Math.round(v * 12700) } : (x.type === 'line' ? x.line : null) }))} />
            <select className="input" aria-label={t('Dash type')} value={(line && line.dash) || 'solid'} onChange={(e) => upd((x) => ({ ...x, line: { color: 'tx1', width: 12700, ...(x.line || {}), dash: e.target.value } }))}>
              {[['solid', 'Solid'], ['dash', 'Dashed'], ['dot', 'Dotted'], ['dashDot', 'Dash dot'], ['lgDash', 'Long dash']].map(([v, l]) => <option key={v} value={v}>{t(l)}</option>)}
            </select>
          </div>
          {el.type === 'line' && (
            <div className="row-tight">
              {['head', 'tail'].map((k) => <select key={k} className="input" aria-label={k === 'head' ? t('Begin arrow') : t('End arrow')} value={(line && line[k]) || 'none'} onChange={(e) => upd((x) => ({ ...x, line: { ...(x.line || {}), [k]: e.target.value } }))}>{[['none', 'None'], ['triangle', 'Arrow'], ['stealth', 'Stealth arrow'], ['arrow', 'Open arrow'], ['oval', 'Oval'], ['diamond', 'Diamond']].map(([v, l]) => <option key={v} value={v}>{`${k === 'head' ? t('Begin') : t('End')}: ${t(l)}`}</option>)}</select>)}
            </div>
          )}
        </div>
      )}
      {(el.type === 'shape' || el.type === 'image') && (
        <div className="of-slides-panesec">
          <label className="row-tight small"><input type="checkbox" checked={!!el.shadow} onChange={(e) => upd((x) => ({ ...x, shadow: e.target.checked ? { color: '#000000', blur: 50800, dist: 38100, dir: 45, alpha: 35 } : null }))} />{t('Shadow')}</label>
        </div>
      )}
      {el.type === 'image' && (
        <div className="of-slides-panesec">
          <div className="small muted">{t('Picture')}</div>
          <div className="row-tight"><span className="small">{t('Transparency')}</span><Num label={t('Transparency')} unit="%" min={0} max={100} value={Math.round((1 - (el.opacity ?? 1)) * 100)} onChange={(v) => upd((x) => ({ ...x, opacity: 1 - v / 100 }))} /></div>
          <div className="small muted">{t('Crop')}</div>
          <div className="of-slides-grid2">
            {[['l', 'Left'], ['t', 'Top'], ['r', 'Right'], ['b', 'Bottom']].map(([k, l]) => <Num key={k} label={t(l)} unit="%" min={0} max={90} value={Math.round(((el.crop || {})[k] || 0) * 100)} onChange={(v) => upd((x) => ({ ...x, crop: { l: 0, t: 0, r: 0, b: 0, ...(x.crop || {}), [k]: v / 100 } }))} />)}
          </div>
          {onCrop && <button type="button" className="btn xs" onClick={onCrop}>{t('Reset picture')}</button>}
          <label className="stack-tight small">{t('Alt text')}<textarea className="input" rows={2} value={el.alt || ''} onKeyDown={(e) => e.stopPropagation()} onChange={(e) => upd((x) => ({ ...x, alt: e.target.value }), `alt:${el.id}`)} /></label>
        </div>
      )}
      {el.text && (
        <div className="of-slides-panesec">
          <div className="small muted">{t('Text box')}</div>
          <select className="input" aria-label={t('Vertical alignment')} value={el.text.anchor || 't'} onChange={(e) => upd((x) => (x.text ? { ...x, text: { ...x.text, anchor: e.target.value } } : x))}>
            <option value="t">{t('Top')}</option><option value="ctr">{t('Middle')}</option><option value="b">{t('Bottom')}</option>
          </select>
          <select className="input" aria-label={t('Autofit')} value={el.text.autofit || 'none'} onChange={(e) => upd((x) => (x.text ? { ...x, text: { ...x.text, autofit: e.target.value, fontScale: undefined } } : x))}>
            <option value="none">{t('Do not autofit')}</option><option value="shrink">{t('Shrink text on overflow')}</option><option value="resize">{t('Resize shape to fit text')}</option>
          </select>
          <label className="row-tight small"><input type="checkbox" checked={el.text.wrap !== false} onChange={(e) => upd((x) => (x.text ? { ...x, text: { ...x.text, wrap: e.target.checked } } : x))} />{t('Wrap text in shape')}</label>
          <div className="small muted">{t('Inner margins')}</div>
          <div className="of-slides-grid2">
            {[['l', 'Left'], ['r', 'Right'], ['t', 'Top'], ['b', 'Bottom']].map(([k, l]) => <Num key={k} label={t(l)} unit={units} step={0.05} min={0} value={toU((el.text.insets || {})[k] ?? (k === 'l' || k === 'r' ? 91440 : 45720))} onChange={(v) => upd((x) => (x.text ? { ...x, text: { ...x.text, insets: { l: 91440, r: 91440, t: 45720, b: 45720, ...(x.text.insets || {}), [k]: fromU(v) } } } : x))} />)}
          </div>
        </div>
      )}
      <div className="of-slides-panesec">
        <label className="stack-tight small">{t('Link')}<input className="input" value={el.link || ''} placeholder="https://…" onKeyDown={(e) => e.stopPropagation()} onChange={(e) => upd((x) => ({ ...x, link: e.target.value || undefined }), `link:${el.id}`)} /></label>
        <label className="stack-tight small">{t('Name')}<input className="input" value={el.name || ''} onKeyDown={(e) => e.stopPropagation()} onChange={(e) => upd((x) => ({ ...x, name: e.target.value }), `name:${el.id}`)} /></label>
      </div>
    </div>
  );
}
function alphaOf(c) { const m = /@(\d+)$/.exec(c || ''); return m ? +m[1] : 100; }
function setAlpha(c, a) { const base = String(c || 'accent1').replace(/@\d+$/, ''); return a >= 100 ? base : `${base}@${Math.round(a)}`; }

/* ---------------------------------------------------------------- comments */
export function CommentsPane({ doc, slide, change, author, readOnly, onClose }) {
  const [text, setText] = useState('');
  const list = slide.comments || [];
  const add = () => { if (!text.trim()) return; change((d) => updateSlide(d, slide.id, (s) => ({ ...s, comments: [...(s.comments || []), { id: uid('c'), author: author || t('You'), text: text.trim(), time: new Date().toISOString() }] }))); setText(''); };
  return (
    <div className="of-slides-pane">
      <div className="of-slides-panehead"><MessageSquare size={14} /><b className="grow">{t('Comments')}</b><button type="button" className="btn xs ghost" onClick={onClose}>{t('Close')}</button></div>
      {!list.length && <p className="small muted">{t('No comments on this slide yet.')}</p>}
      {list.map((c) => (
        <div key={c.id} className="of-slides-comment">
          <div className="row-tight small">{h('b', { className: 'grow' }, c.author || '')}<span className="muted">{c.time ? new Date(c.time).toLocaleString(document.documentElement.lang || undefined, { dateStyle: 'short', timeStyle: 'short' }) : ''}</span>
            {!readOnly && <button type="button" className="btn xs ghost icon" title={t('Delete comment')} aria-label={t('Delete comment')} onClick={() => change((d) => updateSlide(d, slide.id, (s) => ({ ...s, comments: s.comments.filter((x) => x.id !== c.id) })))}><Trash2 size={12} /></button>}</div>
          {h('div', { className: 'small of-slides-commenttext' }, c.text || '')}
        </div>
      ))}
      {!readOnly && <div className="stack-tight"><textarea className="input" rows={3} value={text} placeholder={t('Write a comment …')} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) add(); }} onChange={(e) => setText(e.target.value)} /><button type="button" className="btn sm primary" disabled={!text.trim()} onClick={add}>{t('Add comment')}</button></div>}
    </div>
  );
}

/* ---------------------------------------------------------------- animation pane */
export const EFFECTS = [{ id: 'appear', label: 'Appear' }, { id: 'fade', label: 'Fade' }, { id: 'fly', label: 'Fly in' }, { id: 'zoom', label: 'Zoom' }, { id: 'wipe', label: 'Wipe' }];
export function AnimPane({ doc, slide, sel, setSel, change, onClose, readOnly }) {
  const list = slide.animations || [];
  const nameOf = (id) => { const e = findEl(slide.elements, id); if (!e) return '?'; return e.name || (e.text && bodyText(e.text).split('\n')[0].slice(0, 30)) || t(e.type === 'image' ? 'Picture' : e.type === 'table' ? 'Table' : e.type === 'chart' ? 'Chart' : 'Shape'); };
  const move = (i, d) => change((doc2) => updateSlide(doc2, slide.id, (s) => { const a = s.animations.slice(); const j = i + d; if (j < 0 || j >= a.length) return s; [a[i], a[j]] = [a[j], a[i]]; return { ...s, animations: a }; }));
  return (
    <div className="of-slides-pane">
      <div className="of-slides-panehead"><Sparkles size={14} /><b className="grow">{t('Animation pane')}</b><button type="button" className="btn xs ghost" onClick={onClose}>{t('Close')}</button></div>
      {!list.length && <p className="small muted">{t('Select an object and choose an effect under Animations.')}</p>}
      {list.map((a, i) => (
        <div key={a.id} className={`of-slides-animrow ${sel.includes(a.el) ? 'sel' : ''}`} onClick={() => setSel([a.el])}>
          <span className="badge">{i + 1}</span>
          <span className="grow small">{h('span', null, nameOf(a.el))}<span className="muted"> · {t((EFFECTS.find((x) => x.id === a.effect) || {}).label || a.effect)} · {t(a.trigger === 'with' ? 'With previous' : a.trigger === 'after' ? 'After previous' : 'On click')}</span></span>
          {!readOnly && <>
            <button type="button" className="btn xs ghost icon" title={t('Move up')} aria-label={t('Move up')} onClick={(e) => { e.stopPropagation(); move(i, -1); }}><ArrowUp size={12} /></button>
            <button type="button" className="btn xs ghost icon" title={t('Move down')} aria-label={t('Move down')} onClick={(e) => { e.stopPropagation(); move(i, 1); }}><ArrowDown size={12} /></button>
            <button type="button" className="btn xs ghost icon" title={t('Remove')} aria-label={t('Remove')} onClick={(e) => { e.stopPropagation(); change((d) => updateSlide(d, slide.id, (s) => ({ ...s, animations: s.animations.filter((x) => x.id !== a.id) }))); }}><Trash2 size={12} /></button>
          </>}
        </div>
      ))}
    </div>
  );
}
export { walk, useMemo };
