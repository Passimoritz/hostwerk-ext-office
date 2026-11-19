/*
 * Slide show (real fullscreen, transitions, entrance animations, keyboard, laser/pen, black/white screen, auto-advance,
 * loop, rehearse timings) and the presenter view – in the same window or in a second Hostwerk window synced with
 * BroadcastChannel (the second window is the editor screen #/office-edit?…&present=<token>&window=1).
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { ChevronLeft, ChevronRight, X, Pause, Play, RotateCcw, Square, Monitor, PenLine, MousePointer2, Eraser, Minus, Plus } from 'lucide-react';
import { t } from '@hostwerk/i18n';
import SlideView from './render/SlideView.jsx';
import { S } from './units.js';
import { motionOk } from './motion.js';
import { buildSteps, animDecor, visibleSlides } from './anim.js';
import { Keep } from '../_shared/keep.jsx';

export { buildSteps, animDecor };

/* ---------------------------------------------------------------- stage: fitted slide with transition */
function Stage({ doc, idx, step, playing, black, transitionFrom, fit, openLink, rehearse }) {
  const slide = doc.slides[idx];
  const box = fit;
  const W = S(doc.size.w); const H = S(doc.size.h);
  const scale = Math.min(box.w / W, box.h / H);
  const groups = useMemo(() => (slide ? buildSteps(slide) : []), [slide]);
  const decor = useMemo(() => animDecor(groups, step, playing, motionOk()), [groups, step, playing]);
  if (!slide) return null;
  const tr = slide.transition || {};
  const tType = transitionFrom != null && motionOk() ? tr.type || 'none' : 'none';
  const dur = tr.duration || 500;
  const dir = tr.dir || 'l';
  const prev = transitionFrom != null ? doc.slides[transitionFrom] : null;
  const style = { width: W * scale, height: H * scale };
  const anim = tType === 'none' ? undefined : `of-slides-tr-${tType}${['push', 'wipe', 'cover'].includes(tType) ? `-${dir}` : ''} ${dur}ms cubic-bezier(.3,.7,.2,1) both`;
  const outAnim = tType === 'push' ? `of-slides-tr-pushout-${dir} ${dur}ms cubic-bezier(.3,.7,.2,1) both` : undefined;
  return (
    <div className="of-slides-stagebox" style={style}>
      {prev && tType !== 'none' && <div key={`p${transitionFrom}-${idx}`} className="of-slides-layer" style={{ animation: outAnim }}><SlideView doc={doc} slide={prev} index={transitionFrom} scale={scale} mode="show" /></div>}
      <div key={`c${idx}`} className="of-slides-layer" style={{ animation: anim }}><SlideView doc={doc} slide={slide} index={idx} scale={scale} mode="show" elDecor={decor} openLink={openLink} /></div>
      {black && <div className={`of-slides-blank ${black}`} />}
      {rehearse}
    </div>
  );
}

function useBox(ref) {
  const [b, setB] = useState({ w: 800, h: 450 });
  useEffect(() => {
    if (!ref.current) return undefined;
    const m = () => { const r = ref.current.getBoundingClientRect(); setB({ w: r.width, h: r.height }); };
    m(); const ro = new ResizeObserver(m); ro.observe(ref.current);
    return () => ro.disconnect();
  }, [ref]);
  return b;
}

const fmtTime = (ms) => { const s = Math.max(0, Math.floor(ms / 1000)); const hh = Math.floor(s / 3600); const mm = Math.floor((s % 3600) / 60); const ss = s % 60; return `${hh ? `${hh}:` : ''}${String(mm).padStart(hh ? 2 : 1, '0')}:${String(ss).padStart(2, '0')}`; };

/* ---------------------------------------------------------------- presenter console (same window or remote) */
export function PresenterConsole({ doc, idx, step, black, onCmd, startedAt, embedded }) {
  const [now, setNow] = useState(Date.now());
  const [paused, setPaused] = useState(null);
  const [offset, setOffset] = useState(0);
  const [notesSize, setNotesSize] = useState(18);
  useEffect(() => { const i = setInterval(() => setNow(Date.now()), 500); return () => clearInterval(i); }, []);
  const curRef = useRef(null); const nextRef = useRef(null);
  const cur = useBox(curRef); const nxt = useBox(nextRef);
  const vis = visibleSlides(doc);
  const pos = vis.findIndex((v) => v.i === idx);
  const nextVis = vis[pos + 1];
  const slide = doc.slides[idx];
  const groups = slide ? buildSteps(slide) : [];
  const elapsed = (paused || now) - (startedAt || now) - offset;
  const clock = new Date(now).toLocaleTimeString(document.documentElement.lang || undefined, { hour: '2-digit', minute: '2-digit' });
  return (
    <div className={`of-slides-presenter ${embedded ? 'embedded' : ''}`}>
      <div className="of-slides-pv-top">
        <span className="of-slides-pv-timer mono" aria-label={t('Elapsed time')}>{fmtTime(elapsed)}</span>
        <button type="button" className="btn sm ghost icon" title={paused ? t('Resume timer') : t('Pause timer')} aria-label={paused ? t('Resume timer') : t('Pause timer')} onClick={() => { if (paused) { setOffset(offset + (Date.now() - paused)); setPaused(null); } else setPaused(Date.now()); }}>{paused ? <Play size={14} /> : <Pause size={14} />}</button>
        <button type="button" className="btn sm ghost icon" title={t('Restart timer')} aria-label={t('Restart timer')} onClick={() => { setOffset(Date.now() - (startedAt || Date.now())); setPaused(null); }}><RotateCcw size={14} /></button>
        <span className="grow" />
        <span className="of-slides-pv-clock mono">{clock}</span>
      </div>
      <div className="of-slides-pv-main">
        <div className="of-slides-pv-cur" ref={curRef}>
          {slide && <Stage doc={doc} idx={idx} step={step} playing={false} black={black} fit={cur} />}
        </div>
        <div className="of-slides-pv-side">
          <div className="small muted">{t('Next')}</div>
          <div className="of-slides-pv-next" ref={nextRef}>
            {step < groups.length && slide ? <Stage doc={doc} idx={idx} step={step + 1} playing={false} fit={nxt} /> : nextVis ? <Stage doc={doc} idx={nextVis.i} step={buildSteps(nextVis.s)[0] && buildSteps(nextVis.s)[0].auto ? 1 : 0} playing={false} fit={nxt} /> : <div className="of-slides-pv-end">{t('End of slide show')}</div>}
          </div>
          <div className="row-tight of-slides-pv-notesbar"><span className="small muted grow">{t('Notes')}</span>
            <button type="button" className="btn xs ghost icon" title={t('Smaller text')} aria-label={t('Smaller text')} onClick={() => setNotesSize(Math.max(10, notesSize - 2))}><Minus size={12} /></button>
            <button type="button" className="btn xs ghost icon" title={t('Larger text')} aria-label={t('Larger text')} onClick={() => setNotesSize(Math.min(48, notesSize + 2))}><Plus size={12} /></button>
          </div>
          <div className="of-slides-pv-notes" style={{ fontSize: notesSize }}>{slide && slide.notes ? <Keep>{slide.notes}</Keep> : <span className="muted">{t('No notes for this slide.')}</span>}</div>
        </div>
      </div>
      <div className="of-slides-pv-bottom">
        <button type="button" className="btn sm" onClick={() => onCmd({ cmd: 'prev' })} title={t('Previous')} aria-label={t('Previous')}><ChevronLeft size={16} /></button>
        <span className="small">{t('Slide {n} of {total}', { n: pos + 1, total: vis.length })}</span>
        <button type="button" className="btn sm primary" onClick={() => onCmd({ cmd: 'next' })} title={t('Next')} aria-label={t('Next')}><ChevronRight size={16} /></button>
        <span className="grow" />
        <button type="button" className={`btn sm ghost ${black === 'black' ? 'active' : ''}`} onClick={() => onCmd({ cmd: 'black' })} title={t('Black screen (B)')}><Square size={14} />{t('Black screen')}</button>
        <button type="button" className="btn sm" onClick={() => onCmd({ cmd: 'end' })}><X size={14} />{t('End show')}</button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------- the show */
export default function Show({ doc, start = 0, onExit, presenter = false, loop = false, autoAdvance = 0, openLink, rehearse = false, onTimings, openPresenterWindow }) {
  const root = useRef(null);
  const stageRef = useRef(null);
  const box = useBox(stageRef);
  const vis = useMemo(() => visibleSlides(doc), [doc]);
  const firstIdx = vis.find((v) => v.i >= start) ? vis.find((v) => v.i >= start).i : (vis[0] ? vis[0].i : 0);
  const [st, setSt] = useState(() => ({ idx: firstIdx, step: 0, from: null, playing: true }));
  const [black, setBlack] = useState(null);
  const [ended, setEnded] = useState(false);
  const [cursor, setCursor] = useState(true);
  const [tool, setTool] = useState(null); // 'laser' | 'pen'
  const [laser, setLaser] = useState(null);
  const [mode, setMode] = useState(presenter === 'same' ? 'console' : 'audience');
  const typed = useRef('');
  const startedAt = useRef(Date.now());
  const slideStart = useRef(Date.now());
  const timings = useRef({});
  const penCanvas = useRef(null);
  const chan = useRef(null);
  const [remote, setRemote] = useState(false);

  const groupsOf = useCallback((i) => (doc.slides[i] ? buildSteps(doc.slides[i]) : []), [doc]);
  const enter = useCallback((i, from, atEnd) => {
    const g = groupsOf(i);
    if (rehearse && from != null && doc.slides[from]) timings.current[doc.slides[from].id] = Date.now() - slideStart.current;
    slideStart.current = Date.now();
    const ns = { idx: i, step: atEnd ? g.length : (g[0] && g[0].auto ? 1 : 0), from, playing: !atEnd };
    if (stRef.current) stRef.current = ns;
    setSt(ns);
    setEnded(false);
    if (penCanvas.current) penCanvas.current.getContext('2d').clearRect(0, 0, penCanvas.current.width, penCanvas.current.height);
  }, [groupsOf, rehearse, doc]);
  const stRef = useRef(st); stRef.current = st;
  const next = useCallback(() => {
    if (black) { setBlack(null); return; }
    if (ended) { onExit(); return; }
    const s = stRef.current;
    const g = groupsOf(s.idx);
    if (s.step < g.length) { const ns = { ...s, step: s.step + 1, from: null, playing: true }; stRef.current = ns; setSt(ns); return; }
    const p = vis.findIndex((v) => v.i === s.idx);
    const n = vis[p + 1] || (loop ? vis[0] : null);
    if (!n) { setEnded(true); return; }
    enter(n.i, s.idx);
  }, [black, ended, groupsOf, vis, loop, enter, onExit]);
  const prev = useCallback(() => {
    if (black) { setBlack(null); return; }
    if (ended) { setEnded(false); return; }
    const s = stRef.current;
    const g = groupsOf(s.idx);
    if (s.step > (g[0] && g[0].auto ? 1 : 0)) { const ns = { ...s, step: s.step - 1, from: null, playing: false }; stRef.current = ns; setSt(ns); return; }
    const p = vis.findIndex((v) => v.i === s.idx);
    const n = vis[p - 1];
    if (n) enter(n.i, null, true);
  }, [black, ended, groupsOf, vis, enter]);
  const goto = useCallback((i) => { if (doc.slides[i]) enter(i, st.idx); }, [doc, enter, st.idx]);
  const finish = useCallback(() => {
    if (rehearse && doc.slides[st.idx]) { timings.current[doc.slides[st.idx].id] = Date.now() - slideStart.current; if (onTimings) onTimings({ ...timings.current }); }
    onExit();
  }, [rehearse, doc, st.idx, onTimings, onExit]);

  // fullscreen
  useEffect(() => {
    const el = root.current;
    if (el && el.requestFullscreen && !document.fullscreenElement) el.requestFullscreen({ navigationUI: 'hide' }).catch(() => {});
    const fs = () => { if (!document.fullscreenElement && root.current && root.current.dataset.fs === '1') finish(); if (document.fullscreenElement) root.current.dataset.fs = '1'; };
    document.addEventListener('fullscreenchange', fs);
    return () => { document.removeEventListener('fullscreenchange', fs); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // presenter in a second Hostwerk window
  useEffect(() => {
    if (presenter !== 'window' || typeof BroadcastChannel === 'undefined' || !openPresenterWindow) { if (presenter === 'window') setMode('console'); return undefined; }
    const token = Math.random().toString(36).slice(2, 10);
    const ch = new BroadcastChannel(`of-slides-present-${token}`);
    chan.current = ch;
    ch.onmessage = (ev) => {
      const m = ev.data || {};
      if (m.type === 'hello') { setRemote(true); ch.postMessage({ type: 'doc', doc, startedAt: startedAt.current }); }
      if (m.type === 'cmd') handleCmd(m);
      if (m.type === 'bye') setRemote(false);
    };
    const ok = openPresenterWindow(token);
    if (!ok) setMode('console');
    return () => { try { ch.postMessage({ type: 'end' }); } catch { /* ignore */ } ch.close(); chan.current = null; };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (chan.current) chan.current.postMessage({ type: 'state', idx: st.idx, step: st.step, black, ended }); }, [st.idx, st.step, black, ended, remote]);

  const handleCmdRef = useRef(null);
  handleCmdRef.current = (m) => {
    if (m.cmd === 'next') next(); else if (m.cmd === 'prev') prev(); else if (m.cmd === 'goto') goto(m.idx);
    else if (m.cmd === 'black') setBlack((b) => (b === 'black' ? null : 'black')); else if (m.cmd === 'end') finish();
  };
  function handleCmd(m) { handleCmdRef.current(m); }

  // auto-advance
  useEffect(() => {
    if (ended || black) return undefined;
    const slide = doc.slides[st.idx]; if (!slide) return undefined;
    const g = groupsOf(st.idx);
    if (st.step < g.length) {
      // following "after previous" groups are automatic in PowerPoint only within a group; clicks wait
      return undefined;
    }
    const tr = slide.transition || {};
    const ms = !rehearse && doc.show && doc.show.useTimings === false ? 0 : (tr.advanceAfter != null ? tr.advanceAfter : (autoAdvance ? autoAdvance * 1000 : 0));
    if (!ms || rehearse) return undefined;
    const last = (g[g.length - 1] || { items: [] }).items.reduce((a, it) => Math.max(a, it.delay + it.duration), 0);
    const tm = setTimeout(next, ms + (st.playing ? last : 0));
    return () => clearTimeout(tm);
  }, [st, ended, black, doc, groupsOf, autoAdvance, next, rehearse]);

  // keyboard
  useEffect(() => {
    const k = (e) => {
      if (e.target && /INPUT|TEXTAREA/.test(e.target.tagName)) return;
      const key = e.key;
      if (/^[0-9]$/.test(key)) { typed.current += key; return; }
      if (key === 'Enter' && typed.current) { const n = parseInt(typed.current, 10); typed.current = ''; e.preventDefault(); if (n >= 1 && n <= doc.slides.length) goto(n - 1); return; }
      typed.current = '';
      if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n', 'N'].includes(key)) { e.preventDefault(); next(); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(key) && !e.ctrlKey) { e.preventDefault(); prev(); }
      else if (key === 'Home') { e.preventDefault(); if (vis[0]) enter(vis[0].i, st.idx); }
      else if (key === 'End') { e.preventDefault(); if (vis.length) enter(vis[vis.length - 1].i, st.idx); }
      else if (key === 'b' || key === 'B' || key === '.') { e.preventDefault(); setBlack((b) => (b === 'black' ? null : 'black')); }
      else if (key === 'w' || key === 'W' || key === ',') { e.preventDefault(); setBlack((b) => (b === 'white' ? null : 'white')); }
      else if (key === 'Escape' || key === '-') { e.preventDefault(); finish(); }
      else if (key === 'l' || key === 'L') setTool((x) => (x === 'laser' ? null : 'laser'));
      else if ((key === 'p' || key === 'P') && e.ctrlKey) { e.preventDefault(); setTool((x) => (x === 'pen' ? null : 'pen')); }
      else if (key === 'e' || key === 'E') { if (penCanvas.current) penCanvas.current.getContext('2d').clearRect(0, 0, 99999, 99999); }
      else if (key === 'a' && e.ctrlKey) { e.preventDefault(); setTool(null); }
      else if (key === 's' || key === 'S') { if (!presenter) return; setMode((m) => (m === 'console' ? 'audience' : 'console')); }
      e.stopPropagation();
    };
    window.addEventListener('keydown', k, true);
    return () => window.removeEventListener('keydown', k, true);
  }, [next, prev, goto, finish, enter, vis, st.idx, doc, presenter]);

  // cursor hide
  useEffect(() => {
    let tm = null;
    const mv = () => { setCursor(true); clearTimeout(tm); tm = setTimeout(() => setCursor(false), 2500); };
    window.addEventListener('pointermove', mv);
    mv();
    return () => { window.removeEventListener('pointermove', mv); clearTimeout(tm); };
  }, []);

  // pen
  const drawing = useRef(null);
  const onPenDown = (e) => { if (tool !== 'pen') return; e.stopPropagation(); const c = penCanvas.current; const r = c.getBoundingClientRect(); drawing.current = { x: e.clientX - r.left, y: e.clientY - r.top }; c.setPointerCapture(e.pointerId); };
  const onPenMove = (e) => {
    if (tool === 'laser') setLaser({ x: e.clientX, y: e.clientY });
    if (tool !== 'pen' || !drawing.current) return;
    const c = penCanvas.current; const r = c.getBoundingClientRect(); const ctx = c.getContext('2d');
    const p = { x: e.clientX - r.left, y: e.clientY - r.top };
    ctx.strokeStyle = '#E5364A'; ctx.lineWidth = 3; ctx.lineCap = 'round';
    ctx.beginPath(); ctx.moveTo(drawing.current.x, drawing.current.y); ctx.lineTo(p.x, p.y); ctx.stroke();
    drawing.current = p;
  };
  const onClick = (e) => {
    if (tool === 'pen') return;
    if (e.target.closest('a,video,audio,button')) return;
    next();
  };
  const host = typeof document !== 'undefined' ? document.querySelector('.app') || document.body : null;
  const rehearseBox = rehearse ? <RehearseBox startedAt={startedAt.current} slideStart={slideStart.current} onNext={next} /> : null;
  const content = (
    <div ref={root} className={`of-slides-show ${cursor || tool ? '' : 'nocursor'} ${tool === 'laser' ? 'laser' : ''} ${tool === 'pen' ? 'pen' : ''}`} role="dialog" aria-label={t('Slide show')}
      onContextMenu={(e) => { e.preventDefault(); prev(); }}>
      {mode === 'console' ? (
        <PresenterConsole doc={doc} idx={st.idx} step={st.step} black={black} startedAt={startedAt.current} embedded onCmd={(m) => handleCmd(m)} />
      ) : (
        <div className="of-slides-showstage" ref={stageRef} onClick={onClick} onPointerMove={onPenMove} onPointerDown={onPenDown} onPointerUp={() => { drawing.current = null; }}>
          <Stage doc={doc} idx={st.idx} step={st.step} playing={st.playing} from={st.from} transitionFrom={st.from} black={black} fit={box} openLink={openLink} rehearse={rehearseBox} />
          <canvas ref={penCanvas} className="of-slides-pen" width={box.w} height={box.h} />
          {ended && <div className="of-slides-ended" onClick={(e) => { e.stopPropagation(); finish(); }}>{t('End of slide show – click to exit.')}</div>}
          {tool === 'laser' && laser && <div className="of-slides-laser" style={{ left: laser.x, top: laser.y }} />}
          <div className={`of-slides-showbar ${cursor ? 'on' : ''}`} onClick={(e) => e.stopPropagation()}>
            <button type="button" title={t('Previous')} aria-label={t('Previous')} onClick={prev}><ChevronLeft size={18} /></button>
            <span className="small">{vis.findIndex((v) => v.i === st.idx) + 1} / {vis.length}</span>
            <button type="button" title={t('Next')} aria-label={t('Next')} onClick={next}><ChevronRight size={18} /></button>
            <button type="button" className={tool === 'laser' ? 'on' : ''} title={t('Laser pointer (L)')} aria-label={t('Laser pointer (L)')} onClick={() => setTool(tool === 'laser' ? null : 'laser')}><MousePointer2 size={16} /></button>
            <button type="button" className={tool === 'pen' ? 'on' : ''} title={t('Pen (Ctrl+P)')} aria-label={t('Pen (Ctrl+P)')} onClick={() => setTool(tool === 'pen' ? null : 'pen')}><PenLine size={16} /></button>
            <button type="button" title={t('Erase drawings (E)')} aria-label={t('Erase drawings (E)')} onClick={() => penCanvas.current && penCanvas.current.getContext('2d').clearRect(0, 0, 99999, 99999)}><Eraser size={16} /></button>
            {presenter && <button type="button" title={t('Presenter view (S)')} aria-label={t('Presenter view (S)')} onClick={() => setMode('console')}><Monitor size={16} /></button>}
            <button type="button" title={t('End show (Esc)')} aria-label={t('End show (Esc)')} onClick={finish}><X size={16} /></button>
          </div>
        </div>
      )}
      {mode === 'console' && presenter === 'same' && null}
    </div>
  );
  return host ? createPortal(content, host) : content;
}

function RehearseBox({ startedAt, slideStart, onNext }) {
  const [now, setNow] = useState(Date.now());
  useEffect(() => { const i = setInterval(() => setNow(Date.now()), 250); return () => clearInterval(i); }, []);
  return (
    <div className="of-slides-rehearse" onClick={(e) => e.stopPropagation()}>
      <span className="mono">{fmtTime(now - slideStart)}</span>
      <button type="button" className="btn xs" onClick={onNext}>{t('Next')}</button>
      <span className="mono muted">{fmtTime(now - startedAt)}</span>
    </div>
  );
}

/* ---------------------------------------------------------------- remote presenter window */
export function PresenterRemote({ token }) {
  const [state, setState] = useState(null);
  const [doc, setDoc] = useState(null);
  const [startedAt, setStartedAt] = useState(Date.now());
  const [over, setOver] = useState(false);
  const ch = useRef(null);
  useEffect(() => {
    if (typeof BroadcastChannel === 'undefined') return undefined;
    const c = new BroadcastChannel(`of-slides-present-${token}`);
    ch.current = c;
    c.onmessage = (ev) => {
      const m = ev.data || {};
      if (m.type === 'doc') { setDoc(m.doc); setStartedAt(m.startedAt || Date.now()); }
      if (m.type === 'state') setState(m);
      if (m.type === 'end') { setOver(true); setTimeout(() => { try { window.close(); } catch { /* ignore */ } }, 400); }
    };
    c.postMessage({ type: 'hello' });
    const bye = () => c.postMessage({ type: 'bye' });
    window.addEventListener('beforeunload', bye);
    return () => { window.removeEventListener('beforeunload', bye); c.close(); };
  }, [token]);
  useEffect(() => {
    const k = (e) => {
      if (!ch.current) return;
      if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter'].includes(e.key)) { e.preventDefault(); ch.current.postMessage({ type: 'cmd', cmd: 'next' }); }
      if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(e.key)) { e.preventDefault(); ch.current.postMessage({ type: 'cmd', cmd: 'prev' }); }
      if (e.key === 'b' || e.key === 'B') ch.current.postMessage({ type: 'cmd', cmd: 'black' });
      if (e.key === 'Escape') ch.current.postMessage({ type: 'cmd', cmd: 'end' });
    };
    window.addEventListener('keydown', k);
    return () => window.removeEventListener('keydown', k);
  }, []);
  if (over) return <div className="of-slides-pv-wait">{t('The slide show has ended.')}</div>;
  if (!doc || !state) return <div className="of-slides-pv-wait">{t('Waiting for the slide show …')}</div>;
  return <PresenterConsole doc={doc} idx={state.idx} step={state.step} black={state.black} startedAt={startedAt} onCmd={(m) => ch.current && ch.current.postMessage({ type: 'cmd', ...m })} />;
}
