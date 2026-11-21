/*
 * Office 1.1 panes and dialogs of the slides editor: icon library, full outline view, threaded comments, animation
 * pane (entrance/emphasis/exit/motion paths), Designer, compare versions, presentation from a document, live fields, AI.
 */
import { useEffect, useMemo, useRef, useState, createElement as h } from 'react';
import { Modal, Btn, Callout, Loading, Empty } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import {
  Search, MessageSquare, Trash2, Reply, Check, RotateCcw, Sparkles, ArrowUp, ArrowDown, GripVertical, Plus, Wand2, FileText, History, Radio, Bot,
  LogIn, Star, LogOut, Route, ChevronDown, ChevronRight, Palette, X,
} from 'lucide-react';
import SlideView from './render/SlideView.jsx';
import { S } from './units.js';
import { extra } from './lazy.js';
import { normalizePath } from './render/custom.js';
import { bodyText, bodyFromText, findEl, updateEls, updateSlide, uid, slideTitle, insertSlides, moveSlides, newSlide, walk } from './model.js';
import { ANIM_EFFECTS, CLASS_LABEL, effectLabel, DEFAULT_PATHS } from './anim.js';
import { designIdeas, compareDecks, outlineFromMd, outlineFromDocxXml, outlineFromOdtXml, planSlides, slidesFromPlans, outlineText, plansFromAi } from './features.js';
import { aiAsk, fieldCatalog, fieldValue, listVersions, versionBytes, pickDocument } from './host11.js';
import { Keep, Raw } from '../_shared/keep.jsx';
import { keyLabel } from '../_shared/keys.js';

/* ---------------------------------------------------------------- icon library */
const DE_WORDS = { arrow: 'pfeil', user: 'person nutzer', users: 'team leute', mail: 'brief email', phone: 'telefon', calendar: 'kalender termin', clock: 'uhr zeit', lock: 'schloss sicher', key: 'schlüssel', heart: 'herz', star: 'stern', home: 'haus', car: 'auto', plane: 'flugzeug', sun: 'sonne', moon: 'mond', cloud: 'wolke', leaf: 'blatt', rocket: 'rakete', lightbulb: 'idee glühbirne', chart: 'diagramm', money: 'geld', 'piggy-bank': 'sparschwein', euro: 'geld', trophy: 'pokal', target: 'ziel', settings: 'einstellungen zahnrad', search: 'suche lupe', trash: 'papierkorb löschen', globe: 'welt', earth: 'erde welt', book: 'buch', file: 'datei', folder: 'ordner', check: 'haken ok', shield: 'schild sicherheit', truck: 'lkw lieferung', coffee: 'kaffee', gift: 'geschenk', music: 'musik', camera: 'kamera foto', image: 'bild', handshake: 'handschlag partner', building: 'gebäude firma', factory: 'fabrik', store: 'laden geschäft', 'shopping-cart': 'einkaufswagen', 'graduation-cap': 'schule studium', brain: 'gehirn', puzzle: 'puzzle', flame: 'feuer', droplet: 'tropfen wasser', mountain: 'berg', 'tree-pine': 'baum', flower: 'blume', bike: 'fahrrad', train: 'zug bahn', bus: 'bus', ship: 'schiff', bell: 'glocke', megaphone: 'megafon', smile: 'lächeln', 'thumbs-up': 'daumen gut', wifi: 'wlan', server: 'server', database: 'datenbank', laptop: 'laptop', smartphone: 'handy', bot: 'roboter ki', sparkles: 'funkeln ki', zap: 'blitz energie', battery: 'batterie akku' };
export function IconPicker({ onPick, close }) {
  const [q, setQ] = useState('');
  const [open, setOpen] = useState(null);
  const [ICONS, setIcons] = useState(null);
  const [failed, setFailed] = useState(false);
  useEffect(() => { extra().then((m) => setIcons(m.ICONS)).catch(() => setFailed(true)); }, []);
  const ql = q.trim().toLowerCase();
  const match = (name) => !ql || name.includes(ql) || Object.entries(DE_WORDS).some(([k, v]) => name.includes(k) && v.includes(ql));
  const cats = Object.entries(ICONS || {}).map(([c, icons]) => [c, Object.entries(icons).filter(([n]) => match(n))]).filter(([, l]) => l.length);
  const ref = useRef(null);
  useEffect(() => { if (ref.current) ref.current.focus(); }, []);
  return (
    <div className="of-slides-icons">
      <div className="of-slides-icons-search"><Search size={14} aria-hidden="true" /><input ref={ref} className="input" value={q} placeholder={t('Search icons')} aria-label={t('Search icons')} onKeyDown={(e) => e.stopPropagation()} onChange={(e) => setQ(e.target.value)} /></div>
      <div className="of-slides-icons-list">
        {!ICONS && !failed && <Loading />}
        {failed && <div className="small muted">{t('The icon library could not be loaded.')}</div>}
        {ICONS && !cats.length && <div className="small muted">{t('No icon matches “{q}”.', { q })}</div>}
        {cats.map(([c, list]) => (
          <div key={c} className="of-slides-icons-cat">
            <div className="small muted">{t(c)}</div>
            <div className="of-slides-icons-grid">
              {(ql || open === c ? list : list.slice(0, 18)).map(([name, d]) => (
                <Raw as="button" key={name} type="button" className="of-slides-iconbtn" title={name} aria-label={name} onClick={() => { close && close(); onPick(name, d); }}>
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d={d} /></svg>
                </Raw>
              ))}
              {!ql && list.length > 18 && open !== c && <button type="button" className="btn xs ghost" onClick={() => setOpen(c)}>{t('{n} more', { n: list.length - 18 })}</button>}
            </div>
          </div>
        ))}
      </div>
      <div className="small muted of-slides-icons-foot">{t('Icons: Lucide (ISC licence) – they stay editable shapes, also in PowerPoint.')}</div>
    </div>
  );
}
/** a custom-shape element for an icon */
export function iconElement(name, d, { x, y, size, color = 'accent1' }) {
  const w = Math.round(size); const nd = normalizePath(d);
  return { id: uid('i'), type: 'shape', name: `Icon ${name}`, x: Math.round(x), y: Math.round(y), w, h: w, rot: 0, geom: 'custom', lockRatio: true, paths: [{ w: 24, h: 24, d: nd, fill: false }], fill: null, line: { color, width: Math.max(6350, Math.round(w / 12)), dash: 'solid', cap: 'round', join: 'round' }, shadow: null };
}

/* ---------------------------------------------------------------- outline view */
const isTitleEl = (e) => e.ph && /title/i.test(e.ph) && !/sub/i.test(e.ph);
function textEls(s) { const out = []; walk(s.elements, (e) => { if (e.type === 'shape' && e.text && !isTitleEl(e) && !e.preserved && (e.ph || bodyText(e.text).trim())) out.push(e); }); return out; }
function mergeBody(b, str) {
  const nb = bodyFromText(str);
  const paras = nb.paras.map((p, i) => {
    const old = b.paras[i] || b.paras[b.paras.length - 1] || { runs: [] };
    const { runs, ...pp } = old; // eslint-disable-line no-unused-vars
    const first = (old.runs || [])[0] || {};
    const { text, live, ...rp } = first; // eslint-disable-line no-unused-vars
    return { ...pp, level: p.level, runs: p.runs.length ? [{ ...rp, text: p.runs[0].text }] : [] };
  });
  return { ...b, paras };
}
const bodyLines = (b) => b.paras.map((p) => `${'\t'.repeat(p.level || 0)}${(p.runs || []).map((r) => r.text).join('')}`).join('\n');
function tabKey(e, set) {
  if (e.key !== 'Tab') return;
  e.preventDefault();
  const ta = e.target; const st = ta.selectionStart; const v = ta.value; const ls = v.lastIndexOf('\n', st - 1) + 1;
  const nv = e.shiftKey ? (v[ls] === '\t' ? v.slice(0, ls) + v.slice(ls + 1) : v) : `${v.slice(0, ls)}\t${v.slice(ls)}`;
  set(nv);
  requestAnimationFrame(() => { ta.selectionStart = ta.selectionEnd = Math.max(ls, st + (nv.length - v.length)); });
}
export function OutlineView({ doc, change, readOnly, onOpen, spellcheck, curId }) {
  const [drag, setDrag] = useState(null);
  const [over, setOver] = useState(null);
  const addAfter = (i) => { const s = newSlide(doc, 'titleContent'); change((d) => insertSlides(d, [s], i + 1)); setTimeout(() => { const el = document.querySelector(`[data-outline="${s.id}"] input`); if (el) el.focus(); }, 30); };
  return (
    <div className="of-slides-outline" role="list" aria-label={t('Outline view')}>
      <p className="small muted of-slides-outhint">{t('Tab and Shift+Tab change the level · Enter in a title adds a slide · drag ⠿ to reorder.')}</p>
      {doc.slides.map((s, i) => {
        const title = s.elements.find(isTitleEl);
        const els = textEls(s);
        return (
          <div key={s.id} data-outline={s.id} role="listitem" className={`of-slides-outitem ${curId === s.id ? 'cur' : ''} ${drag === s.id ? 'dragging' : ''} ${over === i ? 'drop-before' : ''}`}
            onDragOver={(e) => { if (!drag) return; e.preventDefault(); const r = e.currentTarget.getBoundingClientRect(); setOver(e.clientY < r.top + r.height / 2 ? i : i + 1); }}
            onDrop={(e) => { e.preventDefault(); if (drag && over != null) change((d) => moveSlides(d, [drag], over)); setDrag(null); setOver(null); }}>
            {!readOnly && <span className="of-slides-outgrip" draggable title={t('Drag to reorder')} aria-label={t('Drag to reorder')} onDragStart={(e) => { e.dataTransfer.setData('text/plain', s.id); e.dataTransfer.effectAllowed = 'move'; setDrag(s.id); }} onDragEnd={() => { setDrag(null); setOver(null); }}><GripVertical size={14} /></span>}
            <button type="button" className="of-slides-outnum" onClick={() => onOpen(s.id)} title={t('Open slide')}>{i + 1}</button>
            <div className="grow stack-tight">
              {title ? (
                <input className="of-slides-outtitle" value={bodyText(title.text)} readOnly={readOnly} spellCheck={spellcheck} aria-label={t('Slide title')} placeholder={t('Click to add title')}
                  onChange={(e) => change((d) => updateEls(d, s.id, [title.id], (x) => ({ ...x, text: mergeBody(x.text, e.target.value) })), `outline:t:${s.id}`)}
                  onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter' && !readOnly) { e.preventDefault(); addAfter(i); } }} />
              ) : els.length ? null : h('div', { className: 'of-slides-outtitle muted' }, slideTitle(s) || t('(no title)'))}
              {/* no title placeholder (text boxes only): the first text box is the slide's title – shown once, as the title */}
              {els.map((el, k) => {
                const v = bodyLines(el.text);
                const set = (nv) => change((d) => updateEls(d, s.id, [el.id], (x) => ({ ...x, text: mergeBody(x.text, nv) })), `outline:${el.id}`);
                return <textarea key={el.id} className={`of-slides-outbody${!title && k === 0 ? ' is-lead' : ''}`} value={v} readOnly={readOnly} spellCheck={spellcheck} rows={Math.max(1, v.split('\n').length)} aria-label={t('Slide text')} placeholder={t('Click to add text')} onChange={(e) => set(e.target.value)} onKeyDown={(e) => { e.stopPropagation(); tabKey(e, set); }} />;
              })}
            </div>
          </div>
        );
      })}
      {!readOnly && <button type="button" className="btn sm ghost" onClick={() => addAfter(doc.slides.length - 1)}><Plus size={14} />{t('New slide')}</button>}
    </div>
  );
}

/* ---------------------------------------------------------------- comments (threads) */
const when = (s) => { try { return s ? new Date(s).toLocaleString(document.documentElement.lang || undefined, { dateStyle: 'short', timeStyle: 'short' }) : ''; } catch { return ''; } };
export function CommentsPane({ doc, slide, change, author, readOnly, onClose, selEl, focusId, setFocusId }) {
  const [text, setText] = useState('');
  const [reply, setReply] = useState({});
  const [showDone, setShowDone] = useState(true);
  const list = slide.comments || [];
  const tops = list.filter((c) => !c.parent || !list.some((x) => x.id === c.parent));
  const me = author || t('You');
  const upd = (fn) => change((d) => updateSlide(d, slide.id, (s) => ({ ...s, comments: fn(s.comments || []) })));
  const add = () => { if (!text.trim()) return; const pos = selEl ? { x: selEl.x + selEl.w, y: selEl.y } : { x: Math.round(doc.size.w * 0.02), y: Math.round(doc.size.h * 0.04) }; const id = uid('c'); upd((l) => [...l, { id, author: me, text: text.trim(), time: new Date().toISOString(), pos }]); setText(''); if (setFocusId) setFocusId(id); };
  const addReply = (pid) => { const v = (reply[pid] || '').trim(); if (!v) return; upd((l) => [...l, { id: uid('c'), parent: pid, author: me, text: v, time: new Date().toISOString() }]); setReply({ ...reply, [pid]: '' }); };
  const shown = tops.filter((c) => showDone || !c.done);
  return (
    <div className="of-slides-pane">
      <div className="of-slides-panehead"><MessageSquare size={14} /><b className="grow">{t('Comments')}</b>
        {tops.some((c) => c.done) && <button type="button" className={`btn xs ghost ${showDone ? 'active' : ''}`} onClick={() => setShowDone(!showDone)}>{t('Resolved')}</button>}
        <button type="button" className="btn xs ghost" onClick={onClose}>{t('Close')}</button></div>
      {!readOnly && <div className="stack-tight of-slides-newcomment">
        <textarea className="input" rows={2} value={text} placeholder={selEl ? t('Comment on the selected object …') : t('Write a comment …')} aria-label={t('New comment')} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) add(); }} onChange={(e) => setText(e.target.value)} />
        <div className="row-tight"><span className="small muted grow">{keyLabel('Ctrl+Enter')}</span><button type="button" className="btn sm primary" disabled={!text.trim()} onClick={add}>{t('Post')}</button></div>
      </div>}
      {!shown.length && <p className="small muted">{t('No comments on this slide yet.')}</p>}
      {shown.map((c) => {
        const replies = list.filter((x) => x.parent === c.id);
        return (
          <div key={c.id} className={`of-slides-thread ${c.done ? 'done' : ''} ${focusId === c.id ? 'focus' : ''}`} onClick={() => setFocusId && setFocusId(c.id)}>
            {[c, ...replies].map((x, k) => (
              <div key={x.id} className={`of-slides-comment ${k ? 'reply' : ''}`}>
                <div className="row-tight small">
                  <span className="of-slides-avatar" style={{ background: avatarColor(x.author) }} aria-hidden="true">{initials(x.author)}</span>
                  {h('b', { className: 'grow' }, x.author || '')}<span className="muted">{when(x.time)}</span>
                  {!readOnly && <button type="button" className="btn xs ghost icon" title={t('Delete')} aria-label={t('Delete')} onClick={(e) => { e.stopPropagation(); upd((l) => l.filter((y) => y.id !== x.id && (k || y.parent !== c.id))); }}><Trash2 size={12} /></button>}
                </div>
                {h('div', { className: 'small of-slides-commenttext' }, x.text || '')}
              </div>
            ))}
            {!readOnly && <div className="row-tight of-slides-replyrow">
              <input className="input" value={reply[c.id] || ''} placeholder={t('Reply …')} aria-label={t('Reply')} onClick={(e) => e.stopPropagation()} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter') addReply(c.id); }} onChange={(e) => setReply({ ...reply, [c.id]: e.target.value })} />
              <button type="button" className="btn xs ghost icon" title={t('Reply')} aria-label={t('Reply')} disabled={!(reply[c.id] || '').trim()} onClick={(e) => { e.stopPropagation(); addReply(c.id); }}><Reply size={13} /></button>
              <button type="button" className="btn xs ghost icon" title={c.done ? t('Reopen') : t('Resolve')} aria-label={c.done ? t('Reopen') : t('Resolve')} onClick={(e) => { e.stopPropagation(); upd((l) => l.map((y) => (y.id === c.id ? { ...y, done: !c.done || undefined } : y))); }}>{c.done ? <RotateCcw size={13} /> : <Check size={13} />}</button>
            </div>}
          </div>
        );
      })}
    </div>
  );
}
const initials = (n) => String(n || '?').split(/\s+/).map((w) => w[0] || '').join('').slice(0, 2).toUpperCase();
const AV = ['#E5484D', '#3E63DD', '#30A46C', '#F76B15', '#8E4EC6', '#12A594', '#D6409F', '#AD7F58'];
export const avatarColor = (n) => { let x = 0; for (const c of String(n || '')) x = (x * 31 + c.charCodeAt(0)) >>> 0; return AV[x % AV.length]; };

/* ---------------------------------------------------------------- animation pane */
const CLS_ICON = { entr: LogIn, emph: Star, exit: LogOut, path: Route };
export function AnimPane({ doc, slide, sel, setSel, change, onClose, readOnly, selAnimId, setSelAnimId, onPreview }) {
  const list = slide.animations || [];
  const nameOf = (id) => { const e = findEl(slide.elements, id); if (!e) return '?'; return e.name || (e.text && bodyText(e.text).split('\n')[0].slice(0, 30)) || t(e.type === 'image' ? 'Picture' : e.type === 'table' ? 'Table' : e.type === 'chart' ? 'Chart' : 'Shape'); };
  const move = (i, d) => change((doc2) => updateSlide(doc2, slide.id, (s) => { const a = s.animations.slice(); const j = i + d; if (j < 0 || j >= a.length) return s; [a[i], a[j]] = [a[j], a[i]]; return { ...s, animations: a }; }));
  // timeline: start of each effect within its click group
  let t0 = 0; let last = 0; let click = 0;
  const rows = list.map((a, i) => { if (i === 0 || a.trigger === 'click') { t0 = 0; last = 0; if (a.trigger === 'click' || i === 0) click++; } const start = (a.trigger === 'after' ? last : a.trigger === 'with' ? t0 : 0) + (a.delay || 0); t0 = start; last = Math.max(last, start + (a.duration || 500)); return { a, i, start, click }; });
  const span = Math.max(1000, ...rows.map((r) => r.start + (r.a.duration || 500)));
  return (
    <div className="of-slides-pane">
      <div className="of-slides-panehead"><Sparkles size={14} /><b className="grow">{t('Animation pane')}</b>
        {onPreview && list.length > 0 && <button type="button" className="btn xs" onClick={onPreview}>{t('Play')}</button>}
        <button type="button" className="btn xs ghost" onClick={onClose}>{t('Close')}</button></div>
      {!list.length && <p className="small muted">{t('Select an object and choose an effect under Animations.')}</p>}
      {rows.map(({ a, i, start, click: ck }) => {
        const Icon = CLS_ICON[a.cls || 'entr'] || Sparkles;
        return (
          <div key={a.id} className={`of-slides-animrow cls-${a.cls || 'entr'} ${selAnimId === a.id || (!selAnimId && sel.includes(a.el)) ? 'sel' : ''}`} onClick={() => { setSel([a.el]); if (setSelAnimId) setSelAnimId(a.id); }}>
            <span className="badge">{a.trigger === 'click' || i === 0 ? ck : ''}</span>
            <Icon size={13} className="of-slides-animic" aria-label={t(CLASS_LABEL[a.cls || 'entr'])} />
            <span className="grow small of-slides-animtxt">{h('span', null, nameOf(a.el))}<span className="muted"> · {t(effectLabel(a))}</span>
              <span className="of-slides-animbar" aria-hidden="true"><i style={{ left: `${(start / span) * 100}%`, width: `${Math.max(2, ((a.duration || 500) / span) * 100)}%` }} /></span></span>
            {!readOnly && <>
              <button type="button" className="btn xs ghost icon" title={t('Move up')} aria-label={t('Move up')} onClick={(e) => { e.stopPropagation(); move(i, -1); }}><ArrowUp size={12} /></button>
              <button type="button" className="btn xs ghost icon" title={t('Move down')} aria-label={t('Move down')} onClick={(e) => { e.stopPropagation(); move(i, 1); }}><ArrowDown size={12} /></button>
              <button type="button" className="btn xs ghost icon" title={t('Remove')} aria-label={t('Remove')} onClick={(e) => { e.stopPropagation(); change((d) => updateSlide(d, slide.id, (s) => ({ ...s, animations: s.animations.filter((x) => x.id !== a.id) }))); }}><Trash2 size={12} /></button>
            </>}
          </div>
        );
      })}
    </div>
  );
}
/** ribbon gallery content: effects of one class */
export function EffectGallery({ cls, current, onPick, close }) {
  return (
    <div className="of-slides-effgallery">
      {Object.entries(ANIM_EFFECTS).filter(([k]) => !cls || k === cls).map(([k, list]) => (
        <div key={k} className="stack-tight">
          <div className="small muted">{t(CLASS_LABEL[k])}</div>
          <div className="of-slides-effrow">
            {list.map(([id, label]) => { const Icon = CLS_ICON[k]; return <button key={id} type="button" className={`of-slides-effbtn cls-${k} ${current && (current.cls || 'entr') === k && current.effect === id ? 'on' : ''}`} onClick={() => { if (close) close(); onPick({ cls: k, effect: id, ...(k === 'path' ? { path: DEFAULT_PATHS[id] } : {}) }); }}><Icon size={16} aria-hidden="true" /><span className="small">{t(label)}</span></button>; })}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ---------------------------------------------------------------- Designer */
export function DesignerPane({ doc, slide, onApply, onClose }) {
  const ideas = useMemo(() => designIdeas(doc, slide), [doc, slide]);
  const W = S(doc.size.w);
  return (
    <div className="of-slides-pane">
      <div className="of-slides-panehead"><Wand2 size={14} /><b className="grow">{t('Designer')}</b><button type="button" className="btn xs ghost" onClick={onClose}>{t('Close')}</button></div>
      {!ideas.length && <p className="small muted">{t('Add a title, some points or pictures – Designer then suggests layouts.')}</p>}
      <div className="of-slides-ideas">
        {ideas.map((d) => (
          <button key={d.id} type="button" className="of-slides-idea" title={t(d.name)} onClick={() => onApply(d)}>
            <SlideView doc={doc} slide={{ ...slide, elements: d.elements, background: d.background || slide.background, animations: [] }} scale={232 / W} mode="thumb" />
            <span className="small">{t(d.name)}</span>
          </button>
        ))}
      </div>
      <p className="small muted">{t('Works offline – uses your theme colours. Ctrl+Z undoes it.')}</p>
    </div>
  );
}

/* ---------------------------------------------------------------- compare versions */
export function CompareDialog({ doc, file, host, readAny, onClose, onGoto }) {
  const [versions, setVersions] = useState(null);
  const [err, setErr] = useState(null);
  const [old, setOld] = useState(null);
  const [busy, setBusy] = useState(false);
  const [only, setOnly] = useState(true);
  useEffect(() => { listVersions(host, file).then((v) => setVersions(v.slice().sort((a, b) => (b.time || 0) - (a.time || 0)))).catch((e) => setErr(e.message)); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const load = async (v) => { setBusy(true); try { const bytes = await versionBytes(host, file, v.id); const r = await readAny(bytes, file.ext); setOld({ v, doc: r.doc }); } catch (e) { setErr(e.message); } setBusy(false); };
  const pickLocal = async () => { const f = await pickDocument(host, '.pptx,.pptm,.odp,.ppt,.potx,.ppsx'); if (!f) return; setBusy(true); try { const r = await readAny(f.bytes, f.ext); setOld({ v: { label: f.name }, doc: r.doc }); } catch (e) { setErr(e.message); } setBusy(false); };
  const rows = useMemo(() => (old ? compareDecks(old.doc, doc) : []), [old, doc]);
  const shown = rows.filter((r) => !only || r.status !== 'same');
  const W = S(doc.size.w); const sc = 220 / W;
  const STATUS = { same: 'Unchanged', changed: 'Changed', added: 'Added', removed: 'Removed', moved: 'Moved' };
  const mark = (d, s, ids, cls) => (ids && ids.length ? (s.elements || []).filter((e) => ids.includes(e.id)).map((e) => h('div', { key: e.id, className: `of-slides-diffbox ${cls}`, style: { left: S(e.x) * sc, top: S(e.y) * sc, width: S(e.w) * sc, height: S(e.h) * sc } })) : null);
  return (
    <Modal title={t('Compare with a version')} onClose={onClose} size="xl" footer={<Btn onClick={onClose}>{t('Close')}</Btn>}>
      {err && <Callout kind="warn">{h(Keep, null, String(err))}</Callout>}
      {!old && (
        <div className="stack">
          <p className="small muted">{t('Pick an older version – each slide is compared and the changes are marked.')}</p>
          {/* small loading line – Hostwerk's dialog never shrinks again, a big spinner would leave an empty area */}
          {!versions && !err && <div className="row-tight small muted"><Loading inline />{t('Loading versions …')}</div>}
          {versions && !versions.length && <div className="of-slides-noversions"><History size={18} aria-hidden="true" /><span className="small">{t('This file has no saved versions yet.')}</span></div>}
          {versions && versions.length > 0 && <div className="of-slides-verlist">
            {(versions || []).map((v) => <button key={v.id} type="button" className="of-slides-verrow" disabled={busy} onClick={() => load(v)}><History size={14} />{h('span', { className: 'grow' }, `${when(v.time)}${v.label ? ` · ${v.label}` : ''}`)}<span className="small muted">{Math.round((v.size || 0) / 1024)} KB</span></button>)}
          </div>}
          <button type="button" className="btn sm" onClick={pickLocal} disabled={busy}><FileText size={14} />{t('Compare with another file …')}</button>
          {busy && <div className="row-tight small muted"><Loading inline /></div>}
        </div>
      )}
      {old && (
        <div className="stack">
          <div className="row-tight"><span className="grow small">{t('{n} of {total} slides differ', { n: rows.filter((r) => r.status !== 'same').length, total: rows.length })}</span>
            <label className="row-tight small"><input type="checkbox" checked={only} onChange={(e) => setOnly(e.target.checked)} />{t('Only changes')}</label>
            <button type="button" className="btn xs ghost" onClick={() => setOld(null)}>{t('Other version')}</button></div>
          <div className="of-slides-difflist">
            {!shown.length && <Empty icon={Check} title={t('No differences')} />}
            {shown.map((r, k) => (
              <div key={k} className={`of-slides-diffrow st-${r.status}`}>
                <div className="of-slides-diffhead small"><b>{t(STATUS[r.status])}</b>{r.b != null && <button type="button" className="btn xs ghost" onClick={() => { onGoto(doc.slides[r.b].id); onClose(); }}>{t('Go to slide {n}', { n: r.b + 1 })}</button>}
                  {r.changes && r.changes.notes && <span className="muted">· {t('Notes changed')}</span>}</div>
                <div className="of-slides-diffpair">
                  <div className="of-slides-diffside">{r.a != null ? <div style={{ position: 'relative' }}><SlideView doc={old.doc} slide={old.doc.slides[r.a]} index={r.a} scale={sc} mode="thumb" />{r.changes && mark(old.doc, old.doc.slides[r.a], r.changes.removed, 'del')}</div> : <div className="of-slides-diffnone small muted">{t('not in the old version')}</div>}<div className="small muted">{t('Before')}</div></div>
                  <div className="of-slides-diffside">{r.b != null ? <div style={{ position: 'relative' }}><SlideView doc={doc} slide={doc.slides[r.b]} index={r.b} scale={sc} mode="thumb" />{r.changes && mark(doc, doc.slides[r.b], [...r.changes.added], 'add')}{r.changes && mark(doc, doc.slides[r.b], r.changes.changed, 'chg')}</div> : <div className="of-slides-diffnone small muted">{t('deleted')}</div>}<div className="small muted">{t('Now')}</div></div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </Modal>
  );
}

/* ---------------------------------------------------------------- presentation from a document */
async function outlineOfFile(f) {
  const ext = f.ext;
  if (ext === 'md' || ext === 'markdown' || ext === 'txt') return outlineFromMd(new TextDecoder().decode(f.bytes));
  const { default: JSZip } = await import('jszip');
  const z = await JSZip.loadAsync(f.bytes);
  if (ext === 'docx' || ext === 'docm' || ext === 'dotx') return outlineFromDocxXml(await z.file('word/document.xml').async('string'), z.file('word/styles.xml') ? await z.file('word/styles.xml').async('string') : '');
  if (ext === 'odt' || ext === 'ott') return outlineFromOdtXml(await z.file('content.xml').async('string'));
  throw new Error(t('This file type is not supported – use .docx, .odt, .md or .txt.'));
}
export function FromDocDialog({ doc, host, ai, onClose, onCreate }) {
  const [file, setFile] = useState(null);
  const [outline, setOutline] = useState(null);
  const [err, setErr] = useState(null);
  const [useAi, setUseAi] = useState(false);
  const [busy, setBusy] = useState(false);
  const [mode, setMode] = useState('append');
  const pick = async () => {
    setErr(null);
    const f = await pickDocument(host, '.docx,.odt,.md,.markdown,.txt');
    if (!f) return;
    try { const o = await outlineOfFile(f); setFile(f); setOutline(o); } catch (e) { setErr(e.message); }
  };
  useEffect(() => { pick(); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const plans = useMemo(() => (outline ? planSlides(outline, { title: file && file.name.replace(/\.[^.]+$/, '') }) : []), [outline, file]);
  const create = async () => {
    setBusy(true); setErr(null);
    try {
      let p = plans;
      if (useAi) {
        const ans = await aiAsk(host, `Turn this document into a presentation. Answer ONLY with JSON: an array of slides, each { "title": string, "bullets": [string, max 5, short], "notes": string (speaker notes, 2–4 sentences) }. The first slide is the title slide { "title", "subtitle" }. Use the language of the document. Document:\n\n${outlineText(outline)}`, { system: 'You write clear, concise presentation slides.', json: true });
        const ap = plansFromAi(ans);
        if (!ap) throw new Error(t('The AI answer could not be used – the slides were made from the headings instead.'));
        p = ap;
      }
      onCreate(slidesFromPlans(doc, p), mode);
      onClose();
    } catch (e) { setErr(e.message); if (useAi) { onCreate(slidesFromPlans(doc, plans), mode); onClose(); } }
    setBusy(false);
  };
  return (
    <Modal title={t('Presentation from a document')} onClose={onClose} footer={<><Btn onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" disabled={!plans.length || busy} onClick={create}>{busy ? t('Creating …') : t('Create {n} slides', { n: plans.length })}</Btn></>}>
      <div className="stack">
        {err && <Callout kind="warn">{h(Keep, null, String(err))}</Callout>}
        <div className="row-tight"><button type="button" className="btn sm" onClick={pick}><FileText size={14} />{file ? t('Other document …') : t('Choose a document …')}</button>{file && h('span', { className: 'small muted grow' }, file.name)}</div>
        {outline && !plans.length && <p className="small">{t('No headings or text found in this document.')}</p>}
        {plans.length > 0 && (
          <>
            <div className="of-slides-planlist">
              {plans.slice(0, 40).map((p, i) => <div key={i} className="of-slides-planrow small"><b>{i + 1}</b>{h('span', { className: 'grow' }, p.title || '…')}<span className="muted">{p.kind === 'title' ? t('Title slide') : p.kind === 'section' ? t('Section') : t('{n} points', { n: p.bullets.length })}</span></div>)}
            </div>
            <label className="row-tight small"><input type="radio" checked={mode === 'append'} onChange={() => setMode('append')} />{t('Add after the current slide')}</label>
            <label className="row-tight small"><input type="radio" checked={mode === 'replace'} onChange={() => setMode('replace')} />{t('Replace all slides')}</label>
            {ai && <label className="row-tight small"><input type="checkbox" checked={useAi} onChange={(e) => setUseAi(e.target.checked)} /><Bot size={14} />{t('Summarise with AI (short points and speaker notes)')}</label>}
          </>
        )}
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------- live fields */
export function LiveFieldDialog({ host, onClose, onInsert }) {
  const [list, setList] = useState(null);
  const [err, setErr] = useState(null);
  const [q, setQ] = useState('');
  useEffect(() => { fieldCatalog(host).then(setList).catch((e) => setErr(e.message)); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const pick = async (f) => { let v = null; try { v = await fieldValue(host, f.id, f.params); } catch { v = null; } onInsert(f, v); onClose(); };
  const shown = (list || []).filter((f) => !q || `${f.label} ${f.group || ''} ${f.id}`.toLowerCase().includes(q.toLowerCase()));
  return (
    <Modal title={t('Insert a live field')} onClose={onClose} footer={<Btn onClick={onClose}>{t('Cancel')}</Btn>}>
      <div className="stack">
        <p className="small muted">{t('A live field shows a value from Hostwerk (for example players online, free disk space, the latest backup). It refreshes during the slide show; the last value is saved in the file, so PowerPoint shows it too.')}</p>
        {err && <Callout kind="warn">{t('Live fields are not available right now ({error}).', { error: err })}</Callout>}
        {!list && !err && <Loading />}
        {list && <input className="input" value={q} placeholder={t('Search')} aria-label={t('Search')} onKeyDown={(e) => e.stopPropagation()} onChange={(e) => setQ(e.target.value)} />}
        <div className="of-slides-verlist">
          {shown.map((f) => <Raw as="button" key={`${f.id}${JSON.stringify(f.params || {})}`} type="button" className="of-slides-verrow" title={f.description || f.label} onClick={() => pick(f)}><Radio size={14} />{h('span', { className: 'grow' }, f.label || f.id)}{f.group && h('span', { className: 'small muted' }, f.group)}</Raw>)}
        </div>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------- AI */
export const AI_ACTIONS = [
  { id: 'notes', label: 'Write speaker notes', prompt: (s) => `Write speaker notes (3–6 sentences, natural spoken language, same language as the slide) for this slide. Answer with the notes only.\n\n${s}` },
  { id: 'shorter', label: 'Make shorter', text: true, prompt: (s) => `Rewrite this slide text shorter and crisper. Keep one line per point, keep the language. Answer with the text only.\n\n${s}` },
  { id: 'longer', label: 'Expand', text: true, prompt: (s) => `Expand this slide text a little with concrete detail. Keep one line per point, keep the language. Answer with the text only.\n\n${s}` },
  { id: 'formal', label: 'More formal', text: true, prompt: (s) => `Rewrite this slide text in a more formal, professional tone. Keep one line per point, keep the language. Answer with the text only.\n\n${s}` },
  { id: 'fix', label: 'Fix spelling and grammar', text: true, prompt: (s) => `Correct spelling and grammar of this slide text. Change nothing else. Keep one line per point. Answer with the text only.\n\n${s}` },
];
export function AiDialog({ host, slide, selEl, onClose, onNotes, onText }) {
  const [busy, setBusy] = useState(null);
  const [err, setErr] = useState(null);
  const [custom, setCustom] = useState('');
  const slideStr = () => { const out = []; walk(slide.elements, (e) => { if (e.text) { const s = bodyText(e.text).trim(); if (s) out.push(s); } }); return out.join('\n'); };
  const run = async (act, freePrompt) => {
    setBusy(act ? act.id : 'custom'); setErr(null);
    try {
      const src = act && act.text && selEl && selEl.text ? bodyText(selEl.text) : slideStr();
      const ans = (await aiAsk(host, freePrompt ? `${freePrompt}\n\nSlide:\n${src}\n\nAnswer with the text only.` : act.prompt(src))).trim();
      if (!act || !act.text) { if (act && act.id === 'notes') onNotes(ans); else if (selEl && selEl.text) onText(ans); else onNotes(ans); } else onText(ans);
      onClose();
    } catch (e) { setErr(e.message); }
    setBusy(null);
  };
  return (
    <Modal title={t('AI for this slide')} onClose={onClose} footer={<Btn onClick={onClose}>{t('Close')}</Btn>}>
      <div className="stack">
        {err && <Callout kind="warn">{h(Keep, null, String(err))}</Callout>}
        <p className="small muted">{t('Uses the AI you set up in Hostwerk. Ctrl+Z undoes every change.')}</p>
        <div className="of-slides-aiacts">
          {AI_ACTIONS.map((a) => <button key={a.id} type="button" className="btn sm" disabled={!!busy || (a.text && !(selEl && selEl.text))} onClick={() => run(a)}><Sparkles size={14} />{busy === a.id ? t('Working …') : t(a.label)}</button>)}
        </div>
        {!(selEl && selEl.text) && <p className="small muted">{t('Select a text box to rewrite its text.')}</p>}
        <label className="stack-tight small">{t('/ai – your own instruction')}
          <textarea className="input" rows={2} value={custom} placeholder={t('For example: turn this into three questions for the audience')} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter' && (e.ctrlKey || e.metaKey) && custom.trim()) run(null, custom.trim()); }} onChange={(e) => setCustom(e.target.value)} /></label>
        <div className="row-tight"><span className="grow small muted">{selEl && selEl.text ? t('The answer replaces the selected text.') : t('The answer goes into the speaker notes.')}</span><button type="button" className="btn sm primary" disabled={!custom.trim() || !!busy} onClick={() => run(null, custom.trim())}>{busy === 'custom' ? t('Working …') : t('Run')}</button></div>
      </div>
    </Modal>
  );
}
export { ChevronDown, ChevronRight, Palette, X };
