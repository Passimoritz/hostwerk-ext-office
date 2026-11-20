/*
 * Side panels and bars: navigation (headings), comments + tracked changes, footnotes, find & replace bar,
 * status bar, ruler.
 */
import { useEffect, useRef, useState } from 'react';
import { ChevronUp, ChevronDown, X, Check, Trash2, MessageSquare, RotateCcw, CaseSensitive, WholeWord, Regex, Replace, Minus, Plus } from 'lucide-react';
import { Btn, Seg } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Keep, Raw } from '../_shared/keep.jsx';
import { colorOf } from './extensions/track.js';
import { MM_PX } from './io/pm.js';

const CHANGE_LABEL = (r) => (r.type === 'insertion' ? t('Inserted') : r.type === 'deletion' ? t('Deleted') : r.type === 'pmark' ? (r.attrs.type === 'ins' ? t('New paragraph') : t('Paragraphs joined')) : r.type === 'pchange' ? t('Paragraph formatted') : r.type === 'row' ? (r.attrs.type === 'ins' ? t('Row inserted') : t('Row deleted')) : t('Formatted'));
/** comment text with @mentions highlighted */
export function MentionText({ text }) {
  const parts = String(text || '').split(/(@[p{L}][p{L}p{N}._-]*(?:s[p{Lu}][p{L}.-]*)?)/u);
  return <>{parts.map((p, k) => (k % 2 ? <span key={k} className="of-docs-mention"><Keep>{p}</Keep></span> : <Keep key={k}>{p}</Keep>))}</>;
}
/** "@An" at the caret → names that start with "An" */
function mentionQuery(v) { const m = /(?:^|s)@([p{L}][p{L}p{N}._ -]{0,24})?$/u.exec(v); return m ? (m[1] || '') : null; }
function MentionList({ value, people, onPick }) {
  const q = mentionQuery(value);
  if (q == null) return null;
  const list = people.filter((p) => p.toLowerCase().startsWith(q.toLowerCase())).slice(0, 6);
  if (!list.length) return null;
  return <div className="of-docs-mentions" role="listbox">{list.map((p) => <Raw as="button" type="button" key={p} role="option" className="of-docs-mi" onMouseDown={(e) => e.preventDefault()} onClick={() => onPick(value.replace(/@([p{L}][p{L}p{N}._ -]{0,24})?$/u, `@${p} `))} title={p}>@{p}</Raw>)}</div>;
}
const when = (d) => { if (!d) return ''; const x = new Date(d); return Number.isNaN(x.getTime()) ? '' : x.toLocaleString(document.documentElement.lang === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); };

export function NavPane({ headings, onJump, current, onMove, readOnly }) {
  const [drag, setDrag] = useState(null);
  const [over, setOver] = useState(null);
  const [q, setQ] = useState('');
  const shown = q ? headings.filter((h) => h.text.toLowerCase().includes(q.toLowerCase())) : headings;
  return (
    <nav className="of-docs-side of-docs-nav" aria-label={t('Navigation')}>
      <div className="of-docs-side-head"><b>Navigation</b></div>
      {headings.length > 8 && <div className="of-docs-nav-filter"><input className="input sm" value={q} onChange={(e) => setQ(e.target.value)} placeholder={t('Filter headings')} aria-label={t('Filter headings')} /></div>}
      {headings.length ? (
        <div className="of-docs-side-body">
          {shown.map((h) => (
            <Raw as="button" type="button" key={h.index} draggable={!readOnly && !!onMove && !q}
              className={`of-docs-navitem of-nl${h.level} ${current === h.index ? 'of-on' : ''} ${drag === h.index ? 'of-drag' : ''} ${over === h.index && drag !== h.index ? 'of-over' : ''}`}
              onClick={() => onJump(h.index)} title={readOnly || !onMove ? h.text : `${h.text} – ${t('drag to move this section')}`}
              onDragStart={(e) => { setDrag(h.index); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', h.text); }}
              onDragOver={(e) => { if (drag == null) return; e.preventDefault(); setOver(h.index); }}
              onDragLeave={() => setOver((o) => (o === h.index ? null : o))}
              onDrop={(e) => { e.preventDefault(); if (drag != null && drag !== h.index && onMove) onMove(drag, h.index); setDrag(null); setOver(null); }}
              onDragEnd={() => { setDrag(null); setOver(null); }}>{h.text}</Raw>
          ))}
          {!readOnly && onMove && drag != null && <div className={`of-docs-navend ${over === 'end' ? 'of-over' : ''}`} onDragOver={(e) => { e.preventDefault(); setOver('end'); }} onDrop={(e) => { e.preventDefault(); onMove(drag, null); setDrag(null); setOver(null); }}>{t('Move to the end')}</div>}
        </div>
      ) : <p className="small muted of-docs-side-empty">{t('No headings yet. Use “Heading 1–3” in Home → Styles to build the outline.')}</p>}
    </nav>
  );
}

export function CommentsPane({ comments, anchors, changes, activeId, readOnly, author, onSelect, onChange, onDelete, onReply, onSelectChange, onAccept, onReject, onAcceptAll, onRejectAll, tab, setTab, focusId }) {
  const list = comments.filter((c) => anchors[c.id] != null).sort((a, b) => anchors[a.id] - anchors[b.id]);
  const people = [...new Set([author, ...comments.flatMap((c) => [c.author, ...(c.replies || []).map((r) => r.author)])].filter(Boolean))];
  const [showResolved, setShowResolved] = useState(false);
  const shown = list.filter((c) => showResolved || !c.resolved);
  return (
    <aside className="of-docs-side of-docs-comments" aria-label={t('Comments')}>
      <div className="of-docs-side-head">
        <Seg value={tab} onChange={setTab} options={[{ id: 'comments', label: `${t('Comments')} ${list.length ? `(${list.length})` : ''}` }, { id: 'changes', label: `${t('Changes')} ${changes.length ? `(${changes.length})` : ''}` }]} />
      </div>
      {tab === 'comments' ? (
        <div className="of-docs-side-body">
          {list.some((c) => c.resolved) && <label className="check small"><input type="checkbox" checked={showResolved} onChange={(e) => setShowResolved(e.target.checked)} /> {t('Show resolved')}</label>}
          {!shown.length && <p className="small muted of-docs-side-empty">{readOnly ? t('No comments.') : t('No comments yet. Select text and choose Review → New comment.')}</p>}
          {shown.map((c) => <CommentCard key={c.id} c={c} people={people} active={activeId === c.id} readOnly={readOnly} author={author} autoFocus={focusId === c.id} onSelect={() => onSelect(c.id)} onChange={(patch) => onChange(c.id, patch)} onDelete={() => onDelete(c.id)} onReply={(text) => onReply(c.id, text)} />)}
        </div>
      ) : (
        <div className="of-docs-side-body">
          {!readOnly && changes.length > 0 && <div className="row-tight of-docs-chg-all"><Btn className="xs" icon={Check} onClick={onAcceptAll}>Accept all</Btn><Btn className="xs ghost" icon={X} onClick={onRejectAll}>Reject all</Btn></div>}
          {!changes.length && <p className="small muted of-docs-side-empty">{t('No tracked changes. Turn on Review → Track changes to record edits.')}</p>}
          {changes.map((r, k) => (
            <div key={k} className="card of-docs-chg" style={{ '--of-rev': colorOf(r.attrs.author) }} role="button" tabIndex={0} onClick={() => onSelectChange(r)} onKeyDown={(e) => e.key === 'Enter' && onSelectChange(r)}>
              <div className="of-docs-chg-head"><span className="of-docs-dot" /><b>{r.attrs.author ? <Keep>{r.attrs.author}</Keep> : t('Unknown')}</b><span className="small muted">{when(r.attrs.date)}</span></div>
              <div className="small">{CHANGE_LABEL(r)}: <span className={r.type === 'deletion' ? 'of-docs-del' : r.type === 'insertion' ? 'of-docs-ins' : ''}><Keep>{r.text.length > 80 ? `${r.text.slice(0, 80)}…` : r.text}</Keep></span></div>
              {!readOnly && <div className="row-tight"><Btn className="xs ghost" icon={Check} onClick={(e) => { e.stopPropagation(); onAccept(r); }}>Accept</Btn><Btn className="xs ghost" icon={X} onClick={(e) => { e.stopPropagation(); onReject(r); }}>Reject</Btn></div>}
            </div>
          ))}
        </div>
      )}
    </aside>
  );
}

function CommentCard({ c, people = [], active, readOnly, author, onSelect, onChange, onDelete, onReply, autoFocus }) {
  const [reply, setReply] = useState('');
  const [leaving, setLeaving] = useState(false);
  const resolve = () => { if (c.resolved) { onChange({ resolved: false }); return; } setLeaving(true); setTimeout(() => { setLeaving(false); onChange({ resolved: true }); }, 220); };
  const ref = useRef(null);
  useEffect(() => { if (!autoFocus) return undefined; const tm = setTimeout(() => { if (ref.current) ref.current.focus(); }, 80); return () => clearTimeout(tm); }, [autoFocus]);
  // without an author name in the settings new comments are stored as "Unknown" – they are still the user's own
  const mine = !c.author || c.author === (author || t('Unknown'));
  return (
    <div className={`card of-docs-cmt ${active ? 'of-on' : ''} ${c.resolved ? 'of-resolved' : ''} ${leaving ? 'of-leaving' : ''}`} onClick={onSelect} role="group" aria-label={t('Comment by {name}', { name: c.author || t('Unknown') })}>
      <div className="of-docs-chg-head"><MessageSquare size={13} /><b>{c.author ? <Keep>{c.author}</Keep> : t('Unknown')}</b><span className="small muted">{when(c.date)}</span>{c.resolved && <span className="badge ok">{t('Resolved')}</span>}</div>
      {c.quote && <div className="of-docs-quote small muted">„<Keep>{c.quote.length > 90 ? `${c.quote.slice(0, 90)}…` : c.quote}</Keep>“</div>}
      {readOnly || !mine ? <div className="of-docs-cmt-text"><MentionText text={c.text} /></div> : <textarea ref={ref} className="input of-docs-cmt-input" rows={Math.min(6, Math.max(2, (c.text || '').split('\n').length))} value={c.text} placeholder={t('Write a comment …')} onChange={(e) => onChange({ text: e.target.value })} onClick={(e) => e.stopPropagation()} aria-label="Comment" />}
      {!readOnly && mine && <MentionList value={c.text || ''} people={people} onPick={(v) => onChange({ text: v })} />}
      {(c.replies || []).map((r, k) => <div key={k} className="of-docs-reply"><div className="small"><b>{r.author ? <Keep>{r.author}</Keep> : t('Unknown')}</b> <span className="muted">{when(r.date)}</span></div><div className="of-docs-cmt-text"><MentionText text={r.text} /></div></div>)}
      {!readOnly && active && (
        <div className="stack of-docs-cmt-actions" onClick={(e) => e.stopPropagation()}>
          <div className="row-tight"><input className="input sm grow" value={reply} placeholder={t('Reply …')} onChange={(e) => setReply(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && reply.trim()) { onReply(reply.trim()); setReply(''); } }} aria-label="Reply" /><Btn className="xs" disabled={!reply.trim()} onClick={() => { onReply(reply.trim()); setReply(''); }}>Reply</Btn></div>
          <MentionList value={reply} people={people} onPick={setReply} />
          <div className="row-tight">
            <Btn className="xs ghost" icon={c.resolved ? RotateCcw : Check} onClick={resolve}>{c.resolved ? 'Reopen' : 'Resolve'}</Btn>
            <Btn className="xs ghost" icon={Trash2} onClick={onDelete}>Delete</Btn>
          </div>
        </div>
      )}
    </div>
  );
}

export function FootnotesArea({ notes, readOnly, onChange, focusIndex }) {
  const refs = useRef([]);
  useEffect(() => { if (focusIndex != null && refs.current[focusIndex]) refs.current[focusIndex].focus(); }, [focusIndex]);
  if (!notes.length) return null;
  return (
    <section className="of-docs-notes of-docs-paper" aria-label={t('Footnotes')}>
      <hr />
      {notes.map((n, k) => (
        <div className="of-docs-note" key={k}>
          <sup>{k + 1}</sup>
          {readOnly ? <span><Keep>{n}</Keep></span> : <input ref={(el) => { refs.current[k] = el; }} className="of-docs-note-input" value={n} placeholder={t('Footnote text')} onChange={(e) => onChange(k, e.target.value)} aria-label={t('Footnote {n}', { n: k + 1 })} />}
        </div>
      ))}
    </section>
  );
}

export function FindBar({ editor, mode, setMode, onClose, readOnly }) {
  const [q, setQ] = useState(() => { const { from, to } = editor.state.selection; const s = editor.state.doc.textBetween(from, to, ' '); return s && s.length < 80 ? s : ''; });
  const [r, setR] = useState('');
  const [opts, setOpts] = useState({ matchCase: false, wholeWord: false, regex: false });
  const [, force] = useState(0);
  const input = useRef(null);
  useEffect(() => { if (input.current) { input.current.focus(); input.current.select(); } }, [mode]);
  useEffect(() => { editor.commands.setSearch(q, opts); force((x) => x + 1); }, [q, opts, editor]);
  useEffect(() => () => { editor.commands.setSearch('', {}); }, [editor]);
  const st = editor.storage.ofSearch;
  const count = st.matches.length;
  const cur = count ? st.matches.findIndex((m) => m.from === editor.state.selection.from && m.to === editor.state.selection.to) : -1;
  const next = (d) => { editor.commands.findNext(d); force((x) => x + 1); };
  const tog = (k) => setOpts((o) => ({ ...o, [k]: !o[k] }));
  return (
    <div className="of-docs-find" role="search" onKeyDown={(e) => { if (e.key === 'Escape') { e.preventDefault(); onClose(); } }}>
      <input ref={input} className="input sm" value={q} placeholder={t('Find')} aria-label="Find" onChange={(e) => setQ(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); next(e.shiftKey ? -1 : 1); } }} />
      <span className="small muted of-docs-count">{q ? (count ? t('{a} of {b}', { a: cur + 1 || '–', b: count }) : t('No results')) : ''}</span>
      <button type="button" className={`btn ghost xs icon ${opts.matchCase ? 'of-on' : ''}`} title={t('Match case')} aria-label="Match case" aria-pressed={opts.matchCase} onClick={() => tog('matchCase')}><CaseSensitive /></button>
      <button type="button" className={`btn ghost xs icon ${opts.wholeWord ? 'of-on' : ''}`} title={t('Whole words')} aria-label="Whole words" aria-pressed={opts.wholeWord} onClick={() => tog('wholeWord')}><WholeWord /></button>
      <button type="button" className={`btn ghost xs icon ${opts.regex ? 'of-on' : ''}`} title={t('Regular expression')} aria-label="Regular expression" aria-pressed={opts.regex} onClick={() => tog('regex')}><Regex /></button>
      <button type="button" className="btn ghost xs icon" title={t('Previous (Shift+Enter)')} aria-label="Previous" onClick={() => next(-1)}><ChevronUp /></button>
      <button type="button" className="btn ghost xs icon" title={t('Next (Enter)')} aria-label="Next" onClick={() => next(1)}><ChevronDown /></button>
      {!readOnly && mode !== 'replace' && <button type="button" className="btn ghost xs icon" title={t('Replace (Ctrl+H)')} aria-label="Replace" onClick={() => setMode('replace')}><Replace /></button>}
      {!readOnly && mode === 'replace' && (
        <>
          <input className="input sm" value={r} placeholder={t('Replace with')} aria-label="Replace with" onChange={(e) => setR(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); editor.commands.replaceCurrent(r); force((x) => x + 1); } }} />
          <Btn className="xs" disabled={!count} onClick={() => { editor.commands.replaceCurrent(r); force((x) => x + 1); }}>Replace</Btn>
          <Btn className="xs" disabled={!count} onClick={() => { const n = count; editor.commands.replaceAll(r); force((x) => x + 1); if (n) HUItoast(t('{n} replaced', { n })); }}>Replace all</Btn>
        </>
      )}
      <span className="grow" />
      <button type="button" className="btn ghost xs icon" title={t('Close (Esc)')} aria-label="Close" onClick={onClose}><X /></button>
    </div>
  );
}
let HUItoast = () => {};
export function setToast(fn) { HUItoast = fn; }

export function StatusBar({ page, pages, stats, lang, zoom, setZoom, fitWidth, track, changes, layout, setLayout, goal, onWords, peers }) {
  const pct = goal ? Math.min(100, Math.round((stats.words / goal) * 100)) : 0;
  return (
    <footer className="of-docs-status" aria-label={t('Status')}>
      <span>{t('Page {a} of {b}', { a: page, b: pages })}</span>
      <button type="button" className="of-docs-words" onClick={onWords} title={t('Word count and goal')}>
        {goal ? <span className="of-docs-goal" style={{ '--of-goal': `${pct * 3.6}deg` }} aria-hidden="true" /> : null}
        {goal ? t('{n} of {g} words', { n: stats.words.toLocaleString(), g: goal.toLocaleString() }) : t('{n} words', { n: stats.words.toLocaleString() })}
      </button>
      {peers && peers.length > 0 && <span className="of-docs-peers" title={peers.map((p) => p.name).join(', ')}>{peers.slice(0, 5).map((p, k) => <span key={k} className="of-docs-peer" style={{ background: p.color }}><Keep>{(p.name || '?').slice(0, 1).toUpperCase()}</Keep></span>)}</span>}
      <span className="of-hide-sm">{t('{n} characters', { n: stats.chars.toLocaleString() })}</span>
      <span className="of-hide-sm">{lang}</span>
      {track && <span className="badge warn">{t('Track changes on')}</span>}
      {changes > 0 && <span className="of-hide-sm">{t('{n} changes', { n: changes })}</span>}
      <span className="grow" />
      <Seg value={layout} onChange={setLayout} options={[{ id: 'print', label: 'Print layout' }, { id: 'web', label: 'Web layout' }]} />
      <button type="button" className="btn ghost xs icon" title={t('Zoom out')} aria-label="Zoom out" onClick={() => setZoom(Math.max(25, zoom - 10))}><Minus /></button>
      <input type="range" min="25" max="300" step="5" value={zoom} onChange={(e) => setZoom(Number(e.target.value))} aria-label={t('Zoom')} className="of-docs-zoomrange" />
      <button type="button" className="btn ghost xs icon" title={t('Zoom in')} aria-label="Zoom in" onClick={() => setZoom(Math.min(300, zoom + 10))}><Plus /></button>
      <button type="button" className="btn ghost xs" title={t('Fit to width')} onClick={fitWidth}>{zoom}%</button>
    </footer>
  );
}

/** horizontal ruler with draggable left/right margins (mm in the model, cm or inch ticks) */
export function Ruler({ page, box, units, readOnly, onMargins }) {
  const w = box.width * MM_PX;
  const ref = useRef(null);
  const drag = (side) => (e) => {
    if (readOnly) return;
    e.preventDefault();
    const r = ref.current.getBoundingClientRect();
    const k = r.width / w;
    const move = (ev) => {
      const mm = ((ev.clientX - r.left) / k) / MM_PX;
      const m = { ...page.margins };
      if (side === 'left') m.left = Math.max(0, Math.min(box.width - m.right - 20, Math.round(mm)));
      else m.right = Math.max(0, Math.min(box.width - m.left - 20, Math.round(box.width - mm)));
      onMargins(m, false);
    };
    const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); onMargins(null, true); };
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
  };
  const step = units === 'in' ? 25.4 : 10;
  const ticks = [];
  for (let x = 0, n = 0; x <= box.width + 0.1; x += step / (units === 'in' ? 8 : 4), n++) {
    const major = n % (units === 'in' ? 8 : 4) === 0;
    ticks.push(<span key={n} className={`of-tick ${major ? 'of-major' : ''}`} style={{ left: x * MM_PX }}>{major && n ? Math.round(x / step) : ''}</span>);
  }
  return (
    <div className="of-docs-ruler" ref={ref} style={{ width: w }} aria-hidden="true">
      <span className="of-ruler-margin" style={{ left: 0, width: page.margins.left * MM_PX }} />
      <span className="of-ruler-margin" style={{ right: 0, width: page.margins.right * MM_PX }} />
      {ticks}
      <span className="of-ruler-handle" style={{ left: page.margins.left * MM_PX }} title={t('Left margin – drag to change')} onPointerDown={drag('left')} />
      <span className="of-ruler-handle" style={{ left: (box.width - page.margins.right) * MM_PX }} title={t('Right margin – drag to change')} onPointerDown={drag('right')} />
    </div>
  );
}
