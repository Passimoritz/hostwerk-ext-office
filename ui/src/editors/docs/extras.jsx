/*
 * Interaction helpers of the docs editor: the "/" command menu (and the same commands for Hostwerk's command
 * palette), the floating formatting bubble on a selection, the right-click menu with icons and shortcuts, the page
 * boxes of the print layout (only the visible ones are drawn – 300+ pages stay light) with rich headers/footers, and
 * the page thumbnails strip.
 */
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  Heading1, Heading2, Heading3, List, ListOrdered, ListTodo, Table, Image, Code, Quote, Sigma, Workflow, CalendarDays, Braces, TableOfContents, SeparatorHorizontal,
  Sparkles, Minus, MessageSquarePlus, Superscript, SquareDashed, Pilcrow, Bold, Italic, Underline, Strikethrough, Highlighter, Link, RemoveFormatting, Scissors, Copy,
  ClipboardPaste, ClipboardType, BetweenHorizontalStart, BetweenHorizontalEnd, BetweenVerticalStart, BetweenVerticalEnd, Trash2, TableCellsMerge, TableCellsSplit, Wand2,
  Languages, FileText, CheckSquare, ChevronDownSquare, TextCursorInput, SquareSplitVertical, Type, Smile, Omega,
} from 'lucide-react';
import { t } from '@hostwerk/i18n';
import { keyLabel } from '../_shared/keys.js';
import { Keep, Raw } from '../_shared/keep.jsx';
import { docToHtml } from './io/render-html.js';
import { pageKey } from './extensions/paginate.js';

/* ---------------- commands ("/" menu, palette, context menu) ---------------- */
/** every insertable thing: { id, title, group, icon, keys?, words, run(act, editor) } */
export function docsCommands({ ai = false } = {}) {
  const c = (e) => e.chain().focus();
  return [
    { id: 'p', title: 'Text', group: 'Basic', icon: Type, words: 'paragraph normal text absatz', run: (a, e) => c(e).setParagraphStyle('normal').run() },
    { id: 'h1', title: 'Heading 1', group: 'Basic', icon: Heading1, keys: 'Ctrl+Alt+1', words: 'title h1 überschrift', run: (a, e) => c(e).setParagraphStyle('h1').run() },
    { id: 'h2', title: 'Heading 2', group: 'Basic', icon: Heading2, keys: 'Ctrl+Alt+2', words: 'h2 überschrift', run: (a, e) => c(e).setParagraphStyle('h2').run() },
    { id: 'h3', title: 'Heading 3', group: 'Basic', icon: Heading3, keys: 'Ctrl+Alt+3', words: 'h3 überschrift', run: (a, e) => c(e).setParagraphStyle('h3').run() },
    { id: 'ul', title: 'Bulleted list', group: 'Basic', icon: List, keys: 'Ctrl+Shift+L', words: 'bullets list aufzählung', run: (a, e) => c(e).toggleBulletList().run() },
    { id: 'ol', title: 'Numbered list', group: 'Basic', icon: ListOrdered, keys: 'Ctrl+Shift+7', words: 'numbers list nummerierung', run: (a, e) => c(e).toggleOrderedList().run() },
    { id: 'todo', title: 'Checklist', group: 'Basic', icon: ListTodo, words: 'task todo checkbox checkliste aufgaben', run: (a, e) => c(e).toggleTaskList().run() },
    { id: 'quote', title: 'Quote', group: 'Basic', icon: Quote, words: 'blockquote zitat', run: (a, e) => c(e).setParagraphStyle('quote').run() },
    { id: 'code', title: 'Code', group: 'Basic', icon: Code, words: 'code block programm quelltext', run: (a, e) => c(e).setParagraphStyle('code').run() },
    { id: 'callout', title: 'Info box', group: 'Basic', icon: SquareDashed, words: 'callout note box hinweis kasten', run: (a, e) => c(e).insertCallout('info').run() },
    { id: 'table', title: 'Table', group: 'Insert', icon: Table, words: 'table grid tabelle', run: (a, e) => c(e).insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run() },
    { id: 'image', title: 'Picture', group: 'Insert', icon: Image, words: 'image picture photo bild foto', run: (a) => a.imageFile() },
    { id: 'eq', title: 'Formula', group: 'Insert', icon: Sigma, words: 'equation math latex formel gleichung', run: (a) => a.equation(null, '', false) },
    { id: 'eqblock', title: 'Formula on its own line', group: 'Insert', icon: Sigma, words: 'equation display math latex formel', run: (a) => a.equation(null, '', true) },
    { id: 'diagram', title: 'Diagram from text', group: 'Insert', icon: Workflow, words: 'mermaid flowchart chart diagramm ablauf gantt', run: (a) => a.diagram() },
    { id: 'date', title: 'Date (updates)', group: 'Fields', icon: CalendarDays, words: 'date today datum heute field', run: (a) => a.quickField('date') },
    { id: 'field', title: 'Live field …', group: 'Fields', icon: Braces, words: 'field live value hostwerk cell zelle feld wert', run: (a) => a.field() },
    { id: 'page', title: 'Page number', group: 'Fields', icon: FileText, words: 'page number seitenzahl', run: (a) => a.quickField('page') },
    { id: 'checkbox', title: 'Check box', group: 'Form', icon: CheckSquare, words: 'checkbox form kontrollkästchen formular', run: (a) => a.control('checkbox') },
    { id: 'dropdown', title: 'Drop-down list', group: 'Form', icon: ChevronDownSquare, words: 'dropdown select list auswahl formular', run: (a) => a.control('dropdown') },
    { id: 'datepick', title: 'Date picker', group: 'Form', icon: CalendarDays, words: 'date picker datum auswahl formular', run: (a) => a.control('date') },
    { id: 'textfield', title: 'Text field', group: 'Form', icon: TextCursorInput, words: 'text field input eingabe formular', run: (a) => a.control('text') },
    { id: 'toc', title: 'Table of contents', group: 'Document', icon: TableOfContents, words: 'toc contents inhaltsverzeichnis', run: (a) => a.toc() },
    { id: 'pagebreak', title: 'Page break', group: 'Document', icon: SeparatorHorizontal, keys: 'Ctrl+Enter', words: 'page break seitenumbruch', run: (a) => a.pageBreak() },
    { id: 'section', title: 'Section break', group: 'Document', icon: SquareSplitVertical, words: 'section landscape abschnitt querformat', run: (a) => a.sectionBreak() },
    { id: 'hr', title: 'Horizontal line', group: 'Document', icon: Minus, words: 'line divider rule linie trennlinie', run: (a, e) => c(e).setHorizontalRule().run() },
    { id: 'footnote', title: 'Footnote', group: 'Document', icon: Superscript, keys: 'Ctrl+Alt+F', words: 'footnote fußnote', run: (a) => a.footnote() },
    { id: 'comment', title: 'Comment', group: 'Document', icon: MessageSquarePlus, keys: 'Ctrl+Alt+M', words: 'comment kommentar', run: (a) => a.newComment() },
    { id: 'symbol', title: 'Symbol', group: 'Document', icon: Omega, words: 'symbol sonderzeichen', run: (a) => a.symbols() },
    { id: 'emoji', title: 'Emoji', group: 'Document', icon: Smile, words: 'emoji smiley', run: (a, e) => c(e).insertContent('🙂').run() },
    ...(ai ? [
      { id: 'ai', title: 'Write with AI …', group: 'AI', icon: Sparkles, words: 'ai ki write schreiben', run: (a) => a.ai('write') },
      { id: 'aicont', title: 'Continue writing (AI)', group: 'AI', icon: Wand2, words: 'ai ki continue weiterschreiben', run: (a) => a.ai('continue') },
      { id: 'aisum', title: 'Summarise (AI)', group: 'AI', icon: FileText, words: 'ai ki summary zusammenfassen', run: (a) => a.ai('summarize') },
      { id: 'aitr', title: 'Translate (AI)', group: 'AI', icon: Languages, words: 'ai ki translate übersetzen', run: (a) => a.ai('translate') },
    ] : []),
  ];
}
const norm = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
export function filterCommands(list, q) {
  const n = norm(q).trim();
  if (!n) return list;
  return list.map((x) => {
    const title = norm(t(x.title)); const en = norm(x.title);
    const score = title.startsWith(n) || en.startsWith(n) ? 3 : title.includes(n) || en.includes(n) ? 2 : norm(x.words).includes(n) ? 1 : 0;
    return { x, score };
  }).filter((r) => r.score).sort((a, b) => b.score - a.score).map((r) => r.x);
}

/* ---------------- "/" menu ---------------- */
export function SlashMenu({ editor, state, commands, act }) {
  const [sel, setSel] = useState(0);
  const list = useMemo(() => filterCommands(commands, state ? state.query : ''), [commands, state && state.query]); // eslint-disable-line react-hooks/exhaustive-deps
  const box = useRef(null);
  useEffect(() => { setSel(0); }, [state && state.query]); // eslint-disable-line react-hooks/exhaustive-deps
  const run = (cmd) => {
    if (!cmd || !state) return;
    const to = editor.state.selection.from;
    editor.chain().focus().deleteRange({ from: state.from, to }).closeSlash().run();
    setTimeout(() => cmd.run(act, editor), 0);
  };
  // keys come from the editor (the caret stays in the text)
  useEffect(() => {
    if (!editor) return undefined;
    editor.storage.ofSlash.onKey = (e) => {
      if (e.key === 'ArrowDown') { setSel((s) => Math.min(list.length - 1, s + 1)); return true; }
      if (e.key === 'ArrowUp') { setSel((s) => Math.max(0, s - 1)); return true; }
      if (e.key === 'Enter' || e.key === 'Tab') { if (list.length) { run(list[Math.min(sel, list.length - 1)]); return true; } return false; }
      if (e.key === 'Escape') { editor.commands.closeSlash(); return true; }
      if (e.key === ' ' && !list.length) { editor.commands.closeSlash(); return false; }
      return false;
    };
    return () => { editor.storage.ofSlash.onKey = null; };
  });
  useLayoutEffect(() => { const el = box.current && box.current.querySelector('.of-on'); if (el && el.scrollIntoView) el.scrollIntoView({ block: 'nearest' }); }, [sel]);
  if (!state || !state.coords) return null;
  const host = document.querySelector('.app') || document.body;
  const below = state.coords.bottom + 6;
  const top = below + 320 > window.innerHeight ? Math.max(8, state.coords.top - 326) : below;
  let group = null;
  return createPortal(
    <div ref={box} className="of-docs-slash" role="listbox" aria-label={t('Insert')} style={{ left: Math.min(state.coords.left, window.innerWidth - 300), top }} onMouseDown={(e) => e.preventDefault()}>
      {!list.length && <div className="of-docs-slash-empty small muted">{t('Nothing found – keep typing or press Esc')}</div>}
      {list.map((x, k) => {
        const head = x.group !== group ? (group = x.group) : null;
        const I = x.icon;
        return (
          <div key={x.id}>
            {head && <div className="of-docs-slash-group">{t(head)}</div>}
            <button type="button" role="option" aria-selected={k === sel} className={`of-docs-slash-item ${k === sel ? 'of-on' : ''}`} onMouseEnter={() => setSel(k)} onClick={() => run(x)}>
              <span className="of-docs-slash-ic"><I size={16} aria-hidden="true" /></span><span className="grow">{t(x.title)}</span>{x.keys && <span className="of-mi-keys">{keyLabel(x.keys)}</span>}
            </button>
          </div>
        );
      })}
    </div>, host);
}

/* ---------------- formatting bubble ---------------- */
export function Bubble({ editor, act, readOnly, rootRef, ai }) {
  const [pos, setPos] = useState(null);
  const [down, setDown] = useState(false);
  useEffect(() => {
    if (!editor) return undefined;
    const upd = () => {
      const { state, view } = editor;
      const sel = state.selection;
      if (sel.empty || readOnly || !view.hasFocus() || sel.node || editor.isActive('codeBlock')) { setPos(null); return; }
      try {
        const a = view.coordsAtPos(sel.from); const b = view.coordsAtPos(sel.to, -1);
        const root = rootRef.current && rootRef.current.getBoundingClientRect();
        if (!root) return;
        const top = Math.min(a.top, b.top) - root.top - 46;
        const left = ((a.top === b.top ? (a.left + b.right) / 2 : a.left + 120)) - root.left;
        setPos({ top: top < 4 ? Math.max(a.bottom, b.bottom) - root.top + 8 : top, left: Math.max(130, Math.min(left, root.width - 170)) });
      } catch { setPos(null); }
    };
    const md = () => { setDown(true); setPos(null); };
    const mu = () => { setDown(false); setTimeout(upd, 0); };
    let tm = null;
    const later = () => { clearTimeout(tm); tm = setTimeout(upd, 160); };
    editor.on('selectionUpdate', later); editor.on('blur', () => setTimeout(() => { if (!editor.view.hasFocus()) setPos(null); }, 150));
    const dom = editor.view.dom;
    const scroller = dom.closest('.of-docs-scroll');
    const sc = () => { setPos(null); later(); };
    dom.addEventListener('mousedown', md); window.addEventListener('mouseup', mu); if (scroller) scroller.addEventListener('scroll', sc, { passive: true });
    return () => { editor.off('selectionUpdate', later); dom.removeEventListener('mousedown', md); window.removeEventListener('mouseup', mu); if (scroller) scroller.removeEventListener('scroll', sc); clearTimeout(tm); };
  }, [editor, readOnly]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!pos || down || !editor) return null;
  const c = () => editor.chain().focus();
  const B = ({ icon: I, title, keys, on, onClick }) => <button type="button" className={`of-docs-bb ${on ? 'of-on' : ''}`} title={keys ? `${t(title)} (${keyLabel(keys)})` : t(title)} aria-label={t(title)} aria-pressed={on === undefined ? undefined : !!on} onMouseDown={(e) => e.preventDefault()} onClick={onClick}><I size={15} aria-hidden="true" /></button>;
  return (
    <div className="of-docs-bubble" style={{ top: pos.top, left: pos.left }} role="toolbar" aria-label={t('Formatting')}>
      <B icon={Bold} title="Bold" keys="Ctrl+B" on={editor.isActive('bold')} onClick={() => c().toggleBold().run()} />
      <B icon={Italic} title="Italic" keys="Ctrl+I" on={editor.isActive('italic')} onClick={() => c().toggleItalic().run()} />
      <B icon={Underline} title="Underline" keys="Ctrl+U" on={editor.isActive('underline')} onClick={() => c().toggleUnderline().run()} />
      <B icon={Strikethrough} title="Strikethrough" keys="Alt+Shift+5" on={editor.isActive('strike')} onClick={() => c().toggleStrike().run()} />
      <B icon={Highlighter} title="Highlight" on={editor.isActive('highlight')} onClick={() => (editor.isActive('highlight') ? c().unsetHighlight().run() : c().setHighlight({ color: '#FFFF00' }).run())} />
      <span className="of-docs-bsep" />
      <B icon={Heading2} title="Heading 2" on={editor.isActive('heading', { level: 2 })} onClick={() => c().setParagraphStyle(editor.isActive('heading', { level: 2 }) ? 'normal' : 'h2').run()} />
      <B icon={List} title="Bullets" on={editor.isActive('bulletList')} onClick={() => c().toggleBulletList().run()} />
      <B icon={Link} title="Link" keys="Ctrl+Shift+K" on={editor.isActive('link')} onClick={act.link} />
      <B icon={MessageSquarePlus} title="Comment" keys="Ctrl+Alt+M" onClick={act.newComment} />
      <B icon={RemoveFormatting} title="Clear formatting" keys="Ctrl+Space" onClick={() => c().unsetAllMarks().run()} />
      {ai && <><span className="of-docs-bsep" /><button type="button" className="of-docs-bb of-docs-bb-ai" title={t('Rewrite with AI')} onMouseDown={(e) => e.preventDefault()} onClick={() => act.ai('rewrite')}><Sparkles size={15} aria-hidden="true" /><span>{t('AI')}</span></button></>}
    </div>
  );
}

/* ---------------- context menu ---------------- */
export function ContextMenu({ at, editor, act, readOnly, ai, onClose }) {
  const box = useRef(null);
  const [pos, setPos] = useState({ left: at.x, top: at.y });
  useLayoutEffect(() => {
    const el = box.current; if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({ left: Math.max(6, Math.min(at.x, window.innerWidth - r.width - 6)), top: Math.max(6, Math.min(at.y, window.innerHeight - r.height - 6)) });
  }, [at]);
  useEffect(() => {
    const down = (e) => { if (box.current && !box.current.contains(e.target)) onClose(); };
    const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); onClose(); } };
    window.addEventListener('mousedown', down, true); window.addEventListener('keydown', key, true); window.addEventListener('blur', onClose);
    const first = box.current && box.current.querySelector('button:not([disabled])'); if (first) first.focus();
    return () => { window.removeEventListener('mousedown', down, true); window.removeEventListener('keydown', key, true); window.removeEventListener('blur', onClose); };
  }, [onClose]);
  const c = () => editor.chain().focus();
  const empty = editor.state.selection.empty;
  const inTable = editor.isActive('table');
  const img = editor.isActive('image');
  const items = [
    ...(!readOnly ? [{ icon: Scissors, label: 'Cut', keys: 'Ctrl+X', off: empty, run: () => act.clipboard('cut') }] : []),
    { icon: Copy, label: 'Copy', keys: 'Ctrl+C', off: empty, run: () => act.clipboard('copy') },
    ...(!readOnly ? [{ icon: ClipboardPaste, label: 'Paste', keys: 'Ctrl+V', run: () => act.paste(false) }, { icon: ClipboardType, label: 'Paste as plain text', keys: 'Ctrl+Shift+V', run: () => act.paste(true) }] : []),
    ...(!readOnly ? [null, { icon: MessageSquarePlus, label: 'New comment', keys: 'Ctrl+Alt+M', off: empty, run: act.newComment }, { icon: Link, label: 'Link …', keys: 'Ctrl+Shift+K', run: act.link }, { icon: Pilcrow, label: 'Paragraph …', run: act.paragraphDialog }] : []),
    ...(!readOnly && ai ? [null, { icon: Sparkles, label: empty ? 'Write with AI …' : 'Rewrite with AI …', run: () => act.ai(empty ? 'write' : 'rewrite') }, ...(!empty ? [{ icon: Languages, label: 'Translate (AI)', run: () => act.ai('translate') }] : [])] : []),
    ...(!readOnly && inTable ? [null,
      { icon: BetweenHorizontalStart, label: 'Insert row above', run: () => c().addRowBefore().run() },
      { icon: BetweenHorizontalEnd, label: 'Insert row below', run: () => c().addRowAfter().run() },
      { icon: BetweenVerticalStart, label: 'Insert column left', run: () => c().addColumnBefore().run() },
      { icon: BetweenVerticalEnd, label: 'Insert column right', run: () => c().addColumnAfter().run() },
      { icon: TableCellsMerge, label: 'Merge cells', off: !editor.can().mergeCells(), run: () => c().mergeCells().run() },
      { icon: TableCellsSplit, label: 'Split cell', off: !editor.can().splitCell(), run: () => c().splitCell().run() },
      { icon: Trash2, label: 'Delete row', run: act.deleteRow },
      { icon: Trash2, label: 'Delete column', run: () => c().deleteColumn().run() },
      { icon: Trash2, label: 'Delete table', run: () => c().deleteTable().run() }] : []),
    ...(!readOnly && img ? [null, { icon: Image, label: 'Alternative text …', run: act.altText }, { icon: Image, label: 'Replace picture …', run: act.imageFile }] : []),
  ];
  while (items.length && items[items.length - 1] === null) items.pop();
  return createPortal(
    <div ref={box} className="of-docs-ctx" role="menu" style={pos} onContextMenu={(e) => e.preventDefault()} title={t('Hold Ctrl while right-clicking for the browser menu (spelling suggestions)')}>
      {items.map((x, k) => (x ? (
        <button key={k} type="button" role="menuitem" className="of-docs-mi" disabled={x.off} onMouseDown={(e) => e.preventDefault()} onClick={() => { onClose(); x.run(); }}>
          <x.icon size={14} aria-hidden="true" /><span className="grow">{t(x.label)}</span>{x.keys && <span className="of-mi-keys">{keyLabel(x.keys)}</span>}
        </button>
      ) : <hr key={k} />))}
    </div>, document.querySelector('.app') || document.body);
}

/* ---------------- page boxes (print layout) ---------------- */
/** which header/footer a page shows: first page of a section, even pages, or the default one */
export function hfFor(model, sectionAttrs, page, indexInSection, kind) {
  const src = sectionAttrs && sectionAttrs.hf && Object.keys(sectionAttrs.hf).length ? sectionAttrs.hf : model;
  const titlePage = sectionAttrs ? sectionAttrs.titlePage : model.titlePage;
  if (indexInSection === 0 && titlePage) return src[`${kind}First`] || (sectionAttrs ? null : null);
  if (model.evenOdd && (page + 1) % 2 === 0) return src[`${kind}Even`] || src[kind] || null;
  return src[kind] || (sectionAttrs ? model[kind] : null) || null;
}
const htmlCache = new Map();
function hfHtml(x, page, pages, title) {
  if (!x) return '';
  if (!x.doc) {
    const s = String(x.text || '').replace(/\{page\}/g, String(page)).replace(/\{pages\}/g, String(pages)).replace(/\{title\}/g, title || '').replace(/\{date\}/g, new Date().toLocaleDateString());
    return `<p style="text-align:${x.align || 'center'};margin:0">${s.replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch]))}</p>`;
  }
  const key = `${page}|${pages}|${x.text}|${x.doc.content ? x.doc.content.length : 0}|${JSON.stringify(x.doc).length}`;
  let h = htmlCache.get(key);
  if (h == null) { h = docToHtml(x.doc, { fields: { page, pages } }); if (htmlCache.size > 400) htmlCache.clear(); htmlCache.set(key, h); }
  return h;
}
export function PagesLayer({ pages, scrollRef, zoom, model, sectionList, title, gap }) {
  const [range, setRange] = useState([0, 6]);
  useEffect(() => {
    const el = scrollRef.current;
    if (!el || !pages || !pages.length) return undefined;
    let raf = 0;
    const upd = () => {
      raf = 0;
      const z = zoom / 100;
      const top = el.scrollTop / z - 400; const bottom = (el.scrollTop + el.clientHeight) / z + 400;
      let a = 0; while (a < pages.length - 1 && pages[a].top + pages[a].h < top) a++;
      let b = a; while (b < pages.length - 1 && pages[b].top < bottom) b++;
      setRange((r) => (r[0] === a && r[1] === b ? r : [a, b]));
    };
    const on = () => { if (!raf) raf = requestAnimationFrame(upd); };
    upd();
    el.addEventListener('scroll', on, { passive: true });
    window.addEventListener('resize', on);
    return () => { el.removeEventListener('scroll', on); window.removeEventListener('resize', on); if (raf) cancelAnimationFrame(raf); };
  }, [pages, zoom, scrollRef]);
  if (!pages || !pages.length) return null;
  const out = [];
  const firstOfSection = [];
  for (let i = 0; i < pages.length; i++) firstOfSection[i] = i === 0 || pages[i - 1].section !== pages[i].section ? i : firstOfSection[i - 1];
  for (let i = range[0]; i <= Math.min(range[1], pages.length - 1); i++) {
    const p = pages[i];
    const sa = sectionList[p.section] || null;
    const idx = i - firstOfSection[i];
    const head = hfFor(model, sa, i, idx, 'header');
    const foot = hfFor(model, sa, i, idx, 'footer');
    out.push(
      <div key={i} className="of-docs-page" style={{ top: p.top, height: p.h, width: p.w }} aria-hidden="true" title={t('Double-click the top or bottom margin to edit the header or footer')}>
        {head && <div className="of-docs-hf of-hf-top of-docs-paper" style={{ top: Math.max(6, p.mt * 0.42), left: p.ml, right: p.mr }} dangerouslySetInnerHTML={{ __html: hfHtml(head, i + 1, pages.length, title) }} />}
        {foot && <div className="of-docs-hf of-hf-bottom of-docs-paper" style={{ bottom: Math.max(6, p.mb * 0.4), left: p.ml, right: p.mr }} dangerouslySetInnerHTML={{ __html: hfHtml(foot, i + 1, pages.length, title) }} />}
        <span className="of-docs-pageno">{i + 1}</span>
      </div>,
    );
  }
  void gap;
  return <>{out}</>;
}

/* ---------------- thumbnails ---------------- */
export function Thumbnails({ editor, pages, current, onJump }) {
  const texts = useMemo(() => {
    if (!editor || !pages || !pages.length) return [];
    const st = pageKey.getState(editor.state);
    const layout = (st && st.layout) || [];
    const per = pages.map(() => []);
    for (const b of layout) {
      const n = editor.state.doc.nodeAt(b.pos);
      if (!n || !per[b.page] || per[b.page].length > 7) continue;
      const txt = n.textContent.slice(0, 90);
      per[b.page].push({ h: n.type.name === 'heading', tbl: n.type.name === 'table', txt });
    }
    return per;
  }, [editor, pages]);
  if (!pages || !pages.length) return <aside className="of-docs-side of-docs-thumbs"><p className="small muted of-docs-side-empty">{t('Thumbnails appear in the print layout.')}</p></aside>;
  return (
    <aside className="of-docs-side of-docs-thumbs" aria-label={t('Pages')}>
      <div className="of-docs-side-head"><b>Pages</b><span className="small muted">{pages.length}</span></div>
      <div className="of-docs-side-body">
        {pages.map((p, i) => (
          <Raw as="button" type="button" key={i} className={`of-docs-thumb ${current === i + 1 ? 'of-on' : ''}`} onClick={() => onJump(i)} title={t('Page {n}', { n: i + 1 })}>
            <span className="of-docs-thumb-paper" style={{ aspectRatio: `${p.w} / ${p.h}` }}>
              {(texts[i] || []).map((x, k) => <span key={k} className={x.h ? 'of-th-h' : x.tbl ? 'of-th-t' : 'of-th-p'}><Keep>{x.txt}</Keep></span>)}
            </span>
            <span className="small">{i + 1}</span>
          </Raw>
        ))}
      </div>
    </aside>
  );
}
