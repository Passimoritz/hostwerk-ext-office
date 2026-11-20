/*
 * The documents editor: tiptap + our extensions, ribbon, print layout with real pages (lines and table rows break
 * over pages, sections with their own page setup, rich headers/footers), web layout, ruler, navigation pane, comments /
 * changes / thumbnails panels, footnotes, find & replace, "/" menu, formatting bubble, block handles, context menu,
 * smooth zoom, live fields, formulas, diagrams, compare, AI, co-editing, status bar, saver/exporter/print for the shell.
 * Everything Hostwerk-global (Ctrl+K, Ctrl+J, Ctrl+1…9, Ctrl+N, Ctrl+, …) stays free; our shortcuts only act inside
 * the editor area.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import Document from '@tiptap/extension-document';
import Strike from '@tiptap/extension-strike';
import Underline from '@tiptap/extension-underline';
import TextAlign from '@tiptap/extension-text-align';
import TextStyle from '@tiptap/extension-text-style';
import Color from '@tiptap/extension-color';
import Highlight from '@tiptap/extension-highlight';
import FontFamily from '@tiptap/extension-font-family';
import LinkExt from '@tiptap/extension-link';
import Table from '@tiptap/extension-table';
import TableRow from '@tiptap/extension-table-row';
import TableCell from '@tiptap/extension-table-cell';
import TableHeader from '@tiptap/extension-table-header';
import TaskList from '@tiptap/extension-task-list';
import TaskItem from '@tiptap/extension-task-item';
import SubscriptExt from '@tiptap/extension-subscript';
import SuperscriptExt from '@tiptap/extension-superscript';
import Placeholder from '@tiptap/extension-placeholder';
import Typography from '@tiptap/extension-typography';
import { TextSelection } from '@tiptap/pm/state';
import { Loading, Callout, Btn, Modal, Spinner, toast, promptDialog, confirmDialog } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Ribbon, fontCss } from './Ribbon.jsx';
import { NavPane, CommentsPane, FootnotesArea, FindBar, StatusBar, Ruler, setToast } from './panels.jsx';
import { SymbolDialog, LinkDialog, ParagraphDialog, WordCountDialog, MacrosSheet } from './dialogs.jsx';
import { EquationEditor, FieldDialog, DiagramDialog, AiDialog, HeaderFooterEditor, SectionDialog, Compare, registerCompareTarget, fieldValue } from './dialogs2.jsx';
import { SlashMenu, Bubble, ContextMenu, PagesLayer, Thumbnails, docsCommands } from './extras.jsx';
import { FontSize, ParagraphFormat, RunKeep, WordKeys } from './extensions/format.js';
import { PageBreak, Footnote, Toc, Callout as CalloutNode, Preserved, ImageEx, Comment } from './extensions/nodes.js';
import { ContentControl, SdtMark, ControlBlock, Field, SectionBreak, Diagram, EquationEx, RevisionAttrs, formatDate } from './extensions/rich.js';
import { Insertion, Deletion, FormatChange, TrackChanges, changeRanges } from './extensions/track.js';
import { Search, FormattingMarks } from './extensions/view.js';
import { Pagination, geometry, pageKey } from './extensions/paginate.js';
import { Slash, slashKey } from './extensions/slash.js';
import { BlockHandles } from './extensions/handles.js';
import { readDocument, writeDocument, EXPORT_FORMATS, MIME } from './io/index.js';
import { standaloneHtml } from './io/render-html.js';
import { mdToHtml } from './io/md.js';
import { PAGE_SIZES, MM_PX, PT_PX, pageBox, uid, sizeName, walk, inlineText, TEXTBLOCKS } from './io/pm.js';
import { vbaModules } from './io/vba.js';
import { renderLatex, compareLib } from './lazy.js';
import { office } from '../../lib/api.js';

setToast((m) => toast(m));
// Ctrl+1/2/5 are Word's line spacing keys (single/double/1.5) – inside the editor they stay with Word (ARCHITECTURE.md
// "Shortcuts"); the other digits keep jumping to Hostwerk's pages
const HOSTWERK_KEYS = /^(ctrl\+(k|j|n|,|[346789])|ctrl\+shift\+(s|x|r)|alt\+arrow(left|right)|f1)$/;
const LINE_KEYS = { 1: 1, 2: 2, 5: 1.5 };
const comboOf = (e) => `${e.ctrlKey || e.metaKey ? 'ctrl+' : ''}${e.altKey ? 'alt+' : ''}${e.shiftKey ? 'shift+' : ''}${String(e.key).toLowerCase()}`;
const FIDELITY = {
  macros: 'Macros (kept unchanged, never run here)', smartart: 'SmartArt graphics (kept unchanged)', chart: 'Charts (kept unchanged)', shape: 'Shapes and drawings (kept unchanged)', textbox: 'Text boxes (kept – double-click to edit their text)', ole: 'Embedded objects (kept unchanged)', object: 'Other embedded content (kept unchanged)',
  field: 'Fields (kept, result shown)', contentControls: 'Some content controls (text kept, the control itself is removed)', trackedChanges: 'Tracked changes (shown and kept – accept or reject them in Review)', paragraphChanges: 'Paragraph formatting changes (shown – accept or reject them in Review)',
  sections: 'Several sections with their own page setup (kept)', headerLayout: 'Pictures or tables in the header/footer (kept as long as you do not edit the header/footer here)', endnotes: 'Endnotes (shown as footnotes)', signature: 'Digital signature (becomes invalid when you save)', images: 'Pictures in formats the browser cannot show (EMF/WMF)', altChunk: 'Embedded HTML/RTF chunks', indexes: 'Indexes (shown as text)', binaryFormat: 'Word 97–2003 format – saved as .docx, or back to .doc with the Office engine',
};
const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : v; } catch { return d; } };
const lsSet = (k, v) => { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch { /* private window */ } };
const locOf = (f) => (f && f.root ? { root: f.root, path: f.path } : { path: (f && f.path) || '' });

/** the editor's extensions (also used to build the schema for co-editing) */
function buildExtensions({ readOnly, ds = {}, lang = 'de', actRef, collab = null }) {
  const smart = ds.smartQuotes !== false;
  const de = String(lang || '').startsWith('de');
  return [
    // section breaks only between top-level blocks (never inside a table cell or a list)
    Document.extend({ content: '(block | sectionBlock)+' }),
    StarterKit.configure({ document: false, strike: false, heading: { levels: [1, 2, 3, 4, 5, 6] }, history: collab ? false : { depth: 500 } }),
    Strike.extend({ addKeyboardShortcuts: () => ({}) }),
    Underline, TextStyle, Color, FontFamily, FontSize,
    Highlight.configure({ multicolor: true }),
    TextAlign.extend({ addKeyboardShortcuts: () => ({}) }).configure({ types: ['heading', 'paragraph'] }),
    LinkExt.extend({ inclusive: false }).configure({ openOnClick: false, autolink: ds.autoLinks !== false, linkOnPaste: true, HTMLAttributes: { target: null, rel: null, class: 'of-docs-link' } }),
    Table.configure({ resizable: !readOnly, lastColumnResizable: false, allowTableNodeSelection: true }), TableRow, TableCell, TableHeader,
    TaskList, TaskItem.configure({ nested: true }),
    SubscriptExt.extend({ addKeyboardShortcuts: () => ({}) }), SuperscriptExt.extend({ addKeyboardShortcuts: () => ({}) }),
    Placeholder.configure({ placeholder: readOnly ? '' : t('Type “/” for commands – or just start writing') }),
    // no automatic ² ½ ≠ … – they would break LaTeX typed between $ signs
    ...(smart ? [Typography.configure({ superscriptTwo: false, superscriptThree: false, oneHalf: false, oneQuarter: false, threeQuarters: false, notEqual: false, multiplication: false, plusMinus: false, ...(de ? { openDoubleQuote: '„', closeDoubleQuote: '“', openSingleQuote: '‚', closeSingleQuote: '‘' } : {}) })] : []),
    ParagraphFormat, RunKeep, WordKeys.configure({ onLink: () => actRef.current.link(), onPageBreak: () => actRef.current.pageBreak(), onTrack: () => { if (!readOnly) actRef.current.toggleTrack(); }, onMarks: () => actRef.current.toggleMarks() }),
    PageBreak, Footnote, Toc.configure({ onJump: (i) => actRef.current.jump(i), onUpdate: () => actRef.current.updateToc() }),
    EquationEx.configure({ onEdit: (pos, latex, display) => actRef.current.equation(pos, latex, display) }),
    CalloutNode, Preserved, ImageEx.configure({ allowBase64: true, inline: true }), Comment,
    ContentControl.configure({ onEdit: (pos) => actRef.current.controlText(pos) }), SdtMark, ControlBlock,
    Field.configure({ onEdit: (pos, node) => actRef.current.field(pos, node) }),
    SectionBreak.configure({ onEdit: (pos) => actRef.current.sectionSetup(pos) }),
    Diagram.configure({ onEdit: (pos, node) => actRef.current.diagram(pos, node) }),
    RevisionAttrs,
    Insertion, Deletion, FormatChange, TrackChanges, Search, FormattingMarks, Pagination, Slash, BlockHandles.configure({ enabled: !readOnly }),
    ...(collab ? [collab.extension] : []),
  ];
}

export function DocsEditor({ file, bytes, readOnly, settings = {}, host = {} }) {
  const [loaded, setLoaded] = useState(null);
  const [error, setError] = useState(null);
  const [collab, setCollab] = useState(null);
  const [joining, setJoining] = useState(false);
  const ext = String((file && file.converted && file.converted.to) || (file && file.ext) || 'docx').toLowerCase();
  useEffect(() => {
    let alive = true;
    setLoaded(null); setError(null);
    const t0 = performance.now();
    readDocument(bytes || new Uint8Array(0), ext, settings).then((r) => { if (alive) setLoaded({ ...r, raw: bytes, readMs: Math.round(performance.now() - t0) }); }, (e) => { if (alive) setError({ text: e.message || String(e), code: e.code || null }); });
    return () => { alive = false; };
  }, [bytes, ext]); // eslint-disable-line react-hooks/exhaustive-deps
  // co-editing: join by itself when someone else is already editing this file
  useEffect(() => {
    if (!loaded || readOnly || collab || !host.collab || typeof host.collab.connect !== 'function' || !file || file.converted) return undefined;
    let alive = true;
    (async () => {
      try {
        const r = await fetch(`/api/office/collab/state?${new URLSearchParams(locOf(file))}`, { credentials: 'same-origin' });
        if (!r.ok) return;
        const j = await r.json();
        if (alive && j && (Array.isArray(j.peers) ? j.peers.length : Number(j.peers) || 0) > 0) start(loaded.model, loaded.html != null ? null : loaded.model.doc, true);
      } catch { /* no co-editing server */ }
    })();
    return () => { alive = false; };
  }, [loaded]); // eslint-disable-line react-hooks/exhaustive-deps
  const start = async (model, doc, auto = false) => {
    if (joining) return;
    setJoining(true);
    try {
      const { joinCollab, collabExtension } = await import('./collab.js');
      const actRef = { current: new Proxy({}, { get: () => () => {} }) };
      const exts = buildExtensions({ readOnly, ds: settings.docs || {}, actRef });
      const session = await joinCollab(host, { extensions: exts, initialDoc: doc || model.doc, model });
      session.extension = collabExtension(session);
      setLoaded((l) => ({ ...l, model, html: null }));
      setCollab(session);
      if (!auto) toast(t('Co-editing is on – everyone who opens this file now edits with you.'));
    } catch (e) { if (!auto) toast.error(e); } finally { setJoining(false); }
  };
  const stop = () => { if (collab) { collab.destroy(); if (host.collab && host.collab.disconnect) host.collab.disconnect(); } setCollab(null); };
  useEffect(() => () => { if (collab) collab.destroy(); }, [collab]);
  if (error) {
    // the readers write English messages – show them in Hostwerk's language; a Word 6/95 file needs the Office engine
    const old = error.code === 'DOC_TOO_OLD';
    const locked = error.code === 'ENCRYPTED';
    return (
      <div className="of-docs of-docs-center">
        <Callout kind={old || locked ? 'info' : 'err'} title={old ? 'This document needs the Office engine' : locked ? 'This document is password-protected' : 'This document could not be opened'} dismissible={false}>
          <div className="stack of-docs-gap">
            <span>{locked ? t('Office password protection is not opened here. Open it in the desktop app, remove the password there (File → Info → Protect document) – or keep it safe in Hostwerk instead: “Encrypt with the vault” in the file’s menu protects it without a password prompt.') : t(error.text)}</span>
            {old && typeof host.ensureEngine === 'function' && <div><Btn className="primary sm" onClick={() => host.ensureEngine(t('“{name}” is a Word 6/95 document – the Office engine opens it exactly.', { name: (file && file.name) || '' }))}>Set up the Office engine …</Btn></div>}
          </div>
        </Callout>
      </div>
    );
  }
  if (!loaded) return <div className="of-docs of-docs-center"><Loading text="Opening document …" /></div>;
  return <EditorBody key={collab ? 'collab' : 'solo'} file={file} ext={ext} initial={loaded} readOnly={!!readOnly} settings={settings} host={host} collab={collab} joining={joining} onCollab={(on, model, doc) => (on ? start(model, doc) : stop())} />;
}

function EditorBody({ file, ext, initial, readOnly, settings, host, collab, joining, onCollab }) {
  const ds = settings.docs || {};
  const gs = settings.general || {};
  const author = gs.author || '';
  const hostLang = host.lang || (document.documentElement.lang === 'de' ? 'de' : 'en');
  const [model, setModel] = useState(() => initial.model);
  const modelRef = useRef(model);
  modelRef.current = model;
  const [comments, setComments] = useState(() => (initial.model.comments || []).map((c) => ({ ...c })));
  const commentsRef = useRef(comments);
  commentsRef.current = comments;
  const [tab, setTab] = useState('home');
  const [layout, setLayout] = useState(() => (['md', 'markdown', 'txt', 'html', 'htm', 'xhtml'].includes(ext) ? 'web' : 'print'));
  const [zoom, setZoomRaw] = useState(() => Number(ds.zoom) || 100);
  const [ruler, setRuler] = useState(ds.ruler !== false);
  const [nav, setNav] = useState(false);
  const [side, setSide] = useState(() => ((initial.model.comments || []).length ? 'comments' : null));
  const [sideTab, setSideTab] = useState('comments');
  const [marks, setMarks] = useState(!!ds.showFormatting);
  const [spell, setSpell] = useState(gs.spellcheck !== false);
  const [lang, setLang] = useState(() => { const l = gs.language && gs.language !== 'auto' ? gs.language : hostLang; return l.length === 2 ? (l === 'de' ? 'de-DE' : 'en-US') : l; });
  const [find, setFind] = useState(null);
  const [dialog, setDialog] = useState(null);
  const [focus, setFocus] = useState(false);
  const [pages, setPages] = useState(null);
  const [track, setTrack] = useState(!!initial.model.trackChanges);
  const [painter, setPainter] = useState(null);
  const [focusComment, setFocusComment] = useState(null);
  const [focusNote, setFocusNote] = useState(null);
  const [slash, setSlash] = useState(null);
  const [ctx, setCtx] = useState(null);
  const [mini, setMini] = useState(() => lsGet('of-docs-ribbon-min', '0') === '1');
  const goalKey = `of-docs-goal:${(file && (file.root || '') + (file.path || file.name)) || ''}`;
  const [goal, setGoalRaw] = useState(() => Number(lsGet(goalKey, 0)) || 0);
  const [peers, setPeers] = useState([]);
  const [, setTick] = useState(0);
  const rootRef = useRef(null);
  const scrollRef = useRef(null);
  const sheetRef = useRef(null);
  const page = model.page;
  const box = pageBox(page);
  const GAP = 28;
  const contentW = (box.width - page.margins.left - page.margins.right) * MM_PX;
  const aiReady = !!(host.ai && typeof host.ai.ask === 'function' && (typeof host.ai.ready !== 'function' || host.ai.ready()));
  const actRef = useRef(null);
  const leader = () => !collab || collab.isLeader();

  const extensions = useMemo(() => buildExtensions({ readOnly, ds, lang, actRef, collab }), []); // eslint-disable-line react-hooks/exhaustive-deps
  const editor = useEditor({
    editable: !readOnly,
    // the chrome re-renders on its own schedule (below) – never once per keystroke
    shouldRerenderOnTransaction: false,
    extensions,
    content: collab ? null : initial.html != null ? initial.html : model.doc,
    editorProps: {
      attributes: { class: 'of-docs-pm', spellcheck: spell ? 'true' : 'false', lang },
      handleClick: (view, pos, e) => {
        const target = e.target && e.target.closest && e.target.closest('a[href]');
        if (target && (e.ctrlKey || e.metaKey || readOnly)) { openLink(target.getAttribute('href')); return true; }
        const fn = e.target && e.target.closest && e.target.closest('.of-docs-fn');
        if (fn) { const all = [...view.dom.querySelectorAll('.of-docs-fn')]; setFocusNote(all.indexOf(fn)); }
        return false;
      },
      handleDoubleClickOn: (view, pos, node) => {
        if (node && node.type.name === 'preserved' && node.attrs.kind === 'textbox' && !readOnly) { actRef.current.textbox(pos, node); return true; }
        return false;
      },
      handlePaste: (view, e) => {
        const files = [...((e.clipboardData && e.clipboardData.files) || [])].filter((f) => /^image\//.test(f.type));
        if (!files.length || readOnly) return false;
        e.preventDefault();
        for (const f of files) readImage(f).then((src) => insertImage(src));
        return true;
      },
      handleDrop: (view, e) => {
        const files = [...((e.dataTransfer && e.dataTransfer.files) || [])].filter((f) => /^image\//.test(f.type));
        if (!files.length || readOnly) return false;
        e.preventDefault();
        const at = view.posAtCoords({ left: e.clientX, top: e.clientY });
        for (const f of files) readImage(f).then((src) => editorRef.current && editorRef.current.chain().focus().insertContentAt(at ? at.pos : editorRef.current.state.selection.from, { type: 'image', attrs: { src } }).run());
        return true;
      },
    },
    onCreate: ({ editor: ed }) => {
      ed.storage.trackChanges.enabled = !!initial.model.trackChanges && !readOnly;
      ed.storage.trackChanges.author = author || t('Unknown');
      ed.storage.wordKeys.base = model.defaults.fontSize || 11;
      if (marks) ed.commands.setFormattingMarks(true);
    },
    onUpdate: ({ transaction }) => {
      if (transaction.getMeta('ofPages') || transaction.getMeta('ofSilent') || !transaction.docChanged) return;
      // co-editing: remote edits are saved by the leader only
      if (collab && transaction.getMeta('y-sync$') && !leader()) return;
      markDirty();
    },
  }, []);
  const editorRef = useRef(editor);
  editorRef.current = editor;

  const dirtyRef = useRef(false);
  const markDirty = useCallback(() => { if (readOnly) return; if (collab && !leader()) return; dirtyRef.current = true; if (host.dirty) host.dirty(true); }, [readOnly, host, collab]); // eslint-disable-line react-hooks/exhaustive-deps
  const openLink = (href) => {
    if (!href) return;
    if (href.startsWith('#')) { const want = decodeURIComponent(href.slice(1)).toLowerCase(); const hs = info.headings; const h = hs.find((x) => x.text.toLowerCase().replace(/\s+/g, '-') === want || x.text.toLowerCase() === want); if (h) actRef.current.jump(h.index); return; }
    if (host.openLink) host.openLink(href); else toast(t('Links open in Hostwerk – this Hostwerk version cannot open them yet.'));
  };
  const readImage = (f) => new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsDataURL(f); });
  const insertImage = (src, extra = {}) => { const ed = editorRef.current; if (!ed) return; ed.chain().focus().setImage({ src, ...extra }).run(); };

  // editor DOM attributes that change at runtime
  useEffect(() => { if (!editor) return; editor.view.dom.setAttribute('spellcheck', spell ? 'true' : 'false'); editor.view.dom.setAttribute('lang', lang); }, [editor, spell, lang]);
  useEffect(() => { if (editor) editor.setEditable(!readOnly); }, [editor, readOnly]);
  useEffect(() => { if (editor) { editor.storage.trackChanges.enabled = track && !readOnly; } }, [editor, track, readOnly]);
  useEffect(() => { if (!editor) return undefined; editor.storage.ofSlash.onChange = setSlash; return () => { editor.storage.ofSlash.onChange = null; }; }, [editor]);

  /* ---------- derived state: computed when typing pauses (never on every keystroke of a long document) ---------- */
  const computeInfo = () => {
    const ed = editorRef.current;
    const empty = { stats: { words: 0, chars: 0, charsNoSpaces: 0, paragraphs: 0 }, headings: [], footnotes: [], anchors: { a: {}, quotes: {} }, changes: [], sections: [] };
    if (!ed) return empty;
    const doc = ed.state.doc;
    let words = 0; let chars = 0; let nospace = 0; let paragraphs = 0;
    const headings = []; const footnotes = []; const a = {}; const quotes = {}; const sections = [];
    let hk = 0;
    const WRE = /[\p{L}\p{N}][\p{L}\p{N}'’\-_.@]*/gu;
    doc.descendants((n, pos) => {
      if (n.type.name === 'sectionBreak') { sections.push(n.attrs); return false; }
      if (n.isTextblock) {
        const s = n.textContent;
        if (s.trim()) paragraphs++;
        if (n.type.name === 'heading') { const text = s.trim(); if (text) headings.push({ level: n.attrs.level || 1, text, index: hk }); hk++; }
        let deleted = 0;
        n.forEach((c) => { if (c.isText && c.marks.some((m) => m.type.name === 'deletion')) deleted += c.text.length; });
        const live = deleted ? n.content.content.filter((c) => !(c.isText && c.marks.some((m) => m.type.name === 'deletion'))).map((c) => c.text || '').join('') : s;
        words += (live.match(WRE) || []).length; chars += live.length; nospace += live.replace(/\s/g, '').length;
      }
      if (n.type.name === 'footnote') footnotes.push(n.attrs.text || '');
      if (n.isText) for (const m of n.marks) if (m.type.name === 'comment') { if (a[m.attrs.id] == null) a[m.attrs.id] = pos; quotes[m.attrs.id] = (quotes[m.attrs.id] || '') + n.text; }
      return true;
    });
    const changes = changeRanges(doc).map((r) => ({ ...r, text: r.type === 'pmark' ? '¶' : doc.textBetween(Math.max(0, r.from), Math.min(doc.content.size, r.to), ' ', ' ') }));
    return { stats: { words, chars, charsNoSpaces: nospace, paragraphs }, headings, footnotes, anchors: { a, quotes }, changes, sections };
  };
  const [info, setInfo] = useState(computeInfo);
  useEffect(() => {
    if (!editor) return undefined;
    let timer = null; let idle = null;
    const run = () => { timer = null; const ric = window.requestIdleCallback || ((f) => setTimeout(f, 1)); idle = ric(() => setInfo(computeInfo()), { timeout: 600 }); };
    const big = () => editor.state.doc.content.size > 150000;
    const h = ({ transaction }) => { if (!transaction.docChanged || transaction.getMeta('ofPages')) return; if (timer) clearTimeout(timer); timer = setTimeout(run, big() ? 700 : 250); };
    editor.on('transaction', h);
    setInfo(computeInfo());
    return () => { editor.off('transaction', h); if (timer) clearTimeout(timer); if (idle && window.cancelIdleCallback) window.cancelIdleCallback(idle); };
  }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps
  const { stats: st, headings, footnotes, anchors, changes } = info;

  // pagination: one geometry per section (the page setup of each section break, the last one = the document's)
  const sectionList = useMemo(() => [...info.sections, null], [info.sections]);
  const sectionsKey = JSON.stringify(info.sections.map((s) => [s.type, s.page && s.page.width, s.page && s.page.height, s.page && s.page.orientation, s.page && s.page.margins])) + JSON.stringify([page.width, page.height, page.orientation, page.margins]);
  useEffect(() => {
    if (!editor) return;
    editor.storage.ofPages.onPages = (p) => setPages(p);
    const geoms = [...info.sections.map((s) => geometry(s.page || page, pageBox(s.page || page))), geometry(page, box)];
    editor.commands.setPagination({ enabled: layout === 'print' && (page.columns || 1) === 1, sections: geoms, gap: GAP });
  }, [editor, layout, sectionsKey, page.columns]); // eslint-disable-line react-hooks/exhaustive-deps
  // re-render the chrome on selection changes (cheap – the derived state above is not recomputed)
  useEffect(() => {
    if (!editor) return undefined;
    // selection moves: next frame · typing: when the typing pauses (ribbon/status bar would cost a frame per key)
    let raf = 0; let typing = null;
    const now = () => { if (!raf) raf = requestAnimationFrame(() => { raf = 0; setTick((x) => x + 1); }); };
    const h = ({ transaction }) => {
      if (transaction && transaction.getMeta('ofPages')) return;
      if (transaction && transaction.docChanged) { clearTimeout(typing); typing = setTimeout(() => { typing = null; now(); }, 220); return; }
      if (typing) return;
      now();
    };
    const sel = () => { if (!typing) now(); };
    editor.on('selectionUpdate', sel); editor.on('transaction', h);
    return () => { editor.off('selectionUpdate', sel); editor.off('transaction', h); if (raf) cancelAnimationFrame(raf); clearTimeout(typing); };
  }, [editor]);
  // fidelity report to the shell
  useEffect(() => {
    const items = (initial.model.fidelity || []).map((k) => t(FIDELITY[k] || k));
    if (host.fidelity) host.fidelity({ items, kept: true });
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  // html/md sources: pick up the imported document into the model once
  useEffect(() => { if (editor && initial.html != null) setModel((m) => ({ ...m, doc: editor.getJSON() })); }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps
  // the document is usable
  useEffect(() => {
    if (!editor) return;
    requestAnimationFrame(() => { if (host.perf) host.perf('interactive'); });
    window.__ofDocsPerf = { read: initial.readMs, blocks: editor.state.doc.childCount, pagination: () => editor.storage.ofPages.stats };
  }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps

  const setZoom = (z) => setZoomRaw(Math.max(25, Math.min(400, Math.round(z))));
  const fitWidth = () => { const el = scrollRef.current; if (!el) return; const avail = el.clientWidth - 48; setZoom((avail / (box.width * MM_PX)) * 100); };
  // narrow window: start with the page fitted to the width (no sideways scrolling)
  useEffect(() => {
    const h = setTimeout(() => { const el = scrollRef.current; if (el && el.clientWidth < 700 && el.clientWidth - 48 < box.width * MM_PX * (zoom / 100)) fitWidth(); }, 150);
    return () => clearTimeout(h);
  }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps
  // smooth zoom: Ctrl + mouse wheel / pinch on the touchpad, around the pointer
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return undefined;
    let pending = null; let raf = 0;
    const apply = () => {
      raf = 0;
      const { factor, x, y } = pending; pending = null;
      const z1 = zoomRef.current; const z2 = Math.max(25, Math.min(400, z1 * factor));
      if (Math.abs(z2 - z1) < 0.5) return;
      const r = el.getBoundingClientRect();
      const px = x - r.left; const py = y - r.top;
      const sx = el.scrollLeft; const sy = el.scrollTop;
      setZoomRaw(Math.round(z2));
      requestAnimationFrame(() => { el.scrollLeft = (sx + px) * (z2 / z1) - px; el.scrollTop = (sy + py) * (z2 / z1) - py; });
    };
    const wheel = (e) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      const f = Math.exp(-Math.max(-60, Math.min(60, e.deltaY)) * 0.0045);
      pending = { factor: (pending ? pending.factor : 1) * f, x: e.clientX, y: e.clientY };
      if (!raf) raf = requestAnimationFrame(apply);
    };
    el.addEventListener('wheel', wheel, { passive: false });
    return () => { el.removeEventListener('wheel', wheel); if (raf) cancelAnimationFrame(raf); };
  }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- live fields: refreshed on open (not a change of yours – the file is not marked as changed) ---------- */
  const refreshFields = useCallback(async ({ silent = true } = {}) => {
    const ed = editorRef.current;
    if (!ed) return 0;
    const list = [];
    ed.state.doc.descendants((n, pos) => { if (n.type.name === 'field') list.push({ pos, node: n }); });
    if (!list.length) return 0;
    const ctxv = { title: modelRef.current.meta.title || (file && file.name) || '', author: modelRef.current.meta.author || author, filename: (file && file.name) || '', words: info.stats.words, pages: pages ? pages.length : 1 };
    const values = await Promise.all(list.map(async ({ node }) => {
      const a = node.attrs;
      try {
        if (a.kind === 'hostwerk' || a.kind === 'cell') { if (!host.fields || typeof host.fields.value !== 'function') return null; const v = await host.fields.value(a.source, a.params); return { value: String((v && (v.text != null ? v.text : v.value)) ?? ''), updatedAt: (v && v.updatedAt) || new Date().toISOString() }; }
        if (['date', 'time', 'title', 'author', 'filename', 'words', 'pages'].includes(a.kind)) { const fmt = /\\@ "([^"]+)"/.exec(a.instr || ''); return { value: a.kind === 'date' && /SAVEDATE|CREATEDATE|PRINTDATE/i.test(a.instr || '') ? a.value : fieldValue(a.kind, { ...ctxv, format: fmt ? fmt[1].replace(/'/g, '') : 'dd.MM.yyyy' }) }; }
      } catch { return null; }
      return null;
    }));
    const cur = editorRef.current;
    if (!cur) return 0;
    const tr = cur.state.tr; let n = 0;
    list.forEach(({ pos, node }, k) => { const v = values[k]; if (!v || v.value === node.attrs.value) return; const m = tr.mapping.map(pos); const now = tr.doc.nodeAt(m); if (now && now.type.name === 'field') { tr.setNodeMarkup(m, undefined, { ...now.attrs, ...v }); n++; } });
    if (n) { if (silent) tr.setMeta('ofSilent', true).setMeta('addToHistory', false); cur.view.dispatch(tr); }
    return n;
  }, [info.stats.words, pages]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (!editor) return undefined; const h = setTimeout(() => refreshFields().catch(() => {}), 900); return () => clearTimeout(h); }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- co-editing: model parts (comments, page setup, headers) through the shared map ---------- */
  useEffect(() => {
    if (!collab) return undefined;
    const m = collab.meta;
    const pull = () => {
      const patch = {};
      for (const k of ['page', 'header', 'footer', 'headerFirst', 'footerFirst', 'headerEven', 'footerEven', 'titlePage', 'evenOdd']) if (m.has(k)) patch[k] = m.get(k);
      if (Object.keys(patch).length) setModel((x) => ({ ...x, ...patch }));
      if (m.has('comments')) setComments(m.get('comments') || []);
    };
    pull();
    m.observe(pull);
    const upd = () => setPeers(collab.peers());
    collab.awareness.on('change', upd); upd();
    let wasLeader = collab.isLeader();
    const lead = () => { const now = collab.isLeader(); if (now && !wasLeader && dirtyRef.current === false) { /* took over saving */ if (host.dirty) host.dirty(true); } wasLeader = now; };
    collab.awareness.on('change', lead);
    return () => { m.unobserve(pull); collab.awareness.off('change', upd); collab.awareness.off('change', lead); };
  }, [collab]); // eslint-disable-line react-hooks/exhaustive-deps
  const share = (k, v) => { if (collab) collab.meta.set(k, v == null ? null : JSON.parse(JSON.stringify(v))); };
  useEffect(() => { if (collab) share('comments', comments); }, [comments]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- the current document as a model ---------- */
  const buildModel = () => {
    const ed = editorRef.current;
    const doc = ed.getJSON();
    const anchored = new Set();
    ed.state.doc.descendants((n) => { for (const m of n.marks) if (m.type.name === 'comment') anchored.add(m.attrs.id); });
    const m = modelRef.current;
    return { ...m, doc, comments: commentsRef.current.filter((c) => anchored.has(c.id)).map(({ quote, ...c }) => c), trackChanges: track, meta: { ...m.meta, author: m.meta.author || author } };
  };
  /** formulas edited/added here get their MathML (→ real Office maths in Word, MathML objects in ODT) */
  const prepareModel = async (m) => {
    const jobs = [];
    walk(m.doc, (n) => { if (n.type === 'equation' && n.attrs && n.attrs.latex && !n.attrs.mathml && !(n.attrs.omml && n.attrs.ommlLatex === n.attrs.latex)) jobs.push(n); });
    for (const n of jobs) { try { const r = await renderLatex(n.attrs.latex, !!n.attrs.display); if (r.mathml) n.attrs = { ...n.attrs, mathml: r.mathml }; } catch { /* offline: the simple converter writes it */ } }
    return m;
  };

  /* ---------- saver / exporter / status ---------- */
  useEffect(() => {
    if (!editor) return;
    if (host.setSaver) host.setSaver(async () => {
      const m = await prepareModel(buildModel());
      let target = ext;
      if (['doc', 'dot', 'wps', 'wri', 'pages'].includes(target)) {
        const docx = await writeDocument(m, 'docx');
        if (host.hasConverter && host.convert) { try { return await host.convert(docx, 'docx', target === 'dot' ? 'dot' : 'doc'); } catch (e) { toast.error(e); } }
        const ok = await confirmDialog({ title: 'Save as Word document (.docx)?', text: t('Without the Office engine this file cannot be written back in the old .{ext} format. A .docx file with the same name is saved instead – the original stays as it is.', { ext: target }), confirm: 'Save as .docx' });
        if (!ok) throw new Error(t('Not saved'));
        return { bytes: docx, ext: 'docx' };
      }
      if (!['docx', 'docm', 'dotx', 'dotm', 'odt', 'ott', 'fodt', 'rtf', 'txt', 'md', 'markdown', 'html', 'htm', 'xhtml'].includes(target)) target = 'docx';
      const out = await writeDocument(m, target);
      dirtyRef.current = false;
      return target === ext ? out : { bytes: out, ext: target };
    });
    if (host.setExporter) host.setExporter({
      formats: EXPORT_FORMATS.map((f) => ({ ext: f.ext, label: f.ext === 'pdf' && !host.hasConverter ? t('PDF (print dialog)') : t(f.label) })),
      run: async (x) => {
        const m = await prepareModel(buildModel());
        if (x === 'pdf') {
          if (host.hasConverter && host.convert) return host.convert(await writeDocument(m, 'docx'), 'docx', 'pdf');
          doPrint(); return null;
        }
        return writeDocument(m, x);
      },
    });
  }, [editor, track]); // eslint-disable-line react-hooks/exhaustive-deps

  const doPrint = () => {
    const m = buildModel();
    const html = standaloneHtml(m, { print: true });
    if (host.print) host.print(html); else toast.error(t('Printing is not available here.'));
  };

  /* ---------- caret page (from the pagination layout) ---------- */
  const caretPage = (() => {
    if (!editor || layout !== 'print' || !pages || !pages.length) return 1;
    const lay = (pageKey.getState(editor.state) || {}).layout || [];
    const head = editor.state.selection.head;
    let lo = 0; let hi = lay.length - 1; let k = -1;
    while (lo <= hi) { const mid = (lo + hi) >> 1; if (lay[mid].pos <= head) { k = mid; lo = mid + 1; } else hi = mid - 1; }
    return k < 0 ? 1 : Math.min(pages.length, lay[k].page + 1);
  })();
  const pageCount = layout === 'print' && pages ? pages.length : 1;
  useEffect(() => {
    if (host.status) host.status({ text: t('Page {a} of {b} · {n} words', { a: caretPage, b: pageCount, n: st.words.toLocaleString() }), trackChanges: track, changes: changes.length });
  }, [caretPage, pageCount, st.words, track, changes.length, layout]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- actions ---------- */
  const updateModel = (patch) => { setModel((m) => ({ ...m, ...patch })); for (const [k, v] of Object.entries(patch)) share(k, v); markDirty(); };
  const setPage = (patch) => { const p = { ...modelRef.current.page, ...patch }; if (patch.width || patch.height) p.size = sizeName(p.width, p.height); updateModel({ page: p }); };
  const headingPos = (index) => { let k = 0; let target = null; editorRef.current.state.doc.descendants((n, pos) => { if (target != null) return false; if (n.type.name === 'heading') { if (k === index) target = pos; k++; return false; } return true; }); return target; };
  const jump = (index) => {
    const ed = editorRef.current;
    const target = headingPos(index);
    if (target == null) return;
    ed.chain().focus().setTextSelection(target + 1).run();
    const dom = ed.view.nodeDOM(target);
    if (dom && dom.scrollIntoView) dom.scrollIntoView({ block: 'start', behavior: 'smooth' });
  };
  const updateToc = () => {
    const ed = editorRef.current;
    const hs = info.headings.filter((h) => h.level <= 3).map((h) => ({ level: h.level, text: h.text, index: h.index }));
    const tr = ed.state.tr;
    ed.state.doc.descendants((n, pos) => { if (n.type.name === 'toc') tr.setNodeMarkup(pos, undefined, { ...n.attrs, entries: hs }); });
    ed.view.dispatch(tr);
  };
  /** a top-level heading with everything up to the next heading of the same or a higher level */
  const sectionRange = (index) => {
    const doc = editorRef.current.state.doc;
    let k = 0; let start = null; let level = 0; let end = null;
    doc.forEach((n, pos) => {
      if (end != null) return;
      if (n.type.name === 'heading') {
        if (start != null && n.attrs.level <= level) { end = pos; return; }
        if (k === index && start == null) { start = pos; level = n.attrs.level; }
        k++;
      }
    });
    if (start == null) return null;
    return { from: start, to: end != null ? end : doc.content.size };
  };
  const insertTop = (node, pos) => { const ed = editorRef.current; ed.chain().focus().insertContentAt(pos, node).run(); };
  const act = {
    jump, updateToc,
    moveSection: (from, to) => {
      const ed = editorRef.current;
      const src = sectionRange(from);
      if (!src) { toast(t('Only headings at the top level of the document can be moved.')); return; }
      const dest = to == null ? ed.state.doc.content.size : headingPos(to);
      if (dest == null || (dest >= src.from && dest <= src.to)) return;
      const slice = ed.state.doc.slice(src.from, src.to);
      const tr = ed.state.tr;
      if (dest > src.to) { tr.insert(dest, slice.content); tr.delete(src.from, src.to); } else { tr.delete(src.from, src.to); tr.insert(dest, slice.content); }
      ed.view.dispatch(tr.scrollIntoView());
      toast(t('Section moved – Ctrl+Z moves it back.'));
    },
    clipboard: (kind) => { editorRef.current.commands.focus(); try { document.execCommand(kind); } catch { toast(t('Use Ctrl+C / Ctrl+X')); } },
    paste: async (plain) => {
      const ed = editorRef.current;
      try {
        if (plain) { const text = await navigator.clipboard.readText(); const parts = text.replace(/\r\n?/g, '\n').split('\n'); ed.chain().focus().insertContent(parts.length > 1 ? parts.map((p) => ({ type: 'paragraph', content: p ? [{ type: 'text', text: p }] : [] })) : { type: 'text', text }).run(); return; }
        const items = await navigator.clipboard.read();
        for (const it of items) {
          const imgType = it.types.find((x) => /^image\//.test(x));
          if (imgType) { insertImage(await readImage(await it.getType(imgType))); return; }
          if (it.types.includes('text/html')) { const { sanitize } = await import('./io/html.js'); ed.chain().focus().insertContent(sanitize(await (await it.getType('text/html')).text())).run(); return; }
          if (it.types.includes('text/plain')) { ed.chain().focus().insertContent(await (await it.getType('text/plain')).text()).run(); return; }
        }
      } catch { toast(t('The browser did not allow reading the clipboard – press Ctrl+V (or Ctrl+Shift+V for plain text).')); }
    },
    painter: () => {
      const ed = editorRef.current;
      if (painter) { setPainter(null); return; }
      const $f = ed.state.selection.$from;
      setPainter({ marks: (ed.state.storedMarks || $f.marks()).filter((m) => !['comment', 'insertion', 'deletion', 'link', 'sdt'].includes(m.type.name)), para: { ...$f.parent.attrs } });
    },
    fontStep: (dir) => { const ed = editorRef.current; const sizes = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]; const cur = parseFloat(ed.getAttributes('textStyle').fontSize) || model.defaults.fontSize || 11; const next = dir > 0 ? sizes.find((x) => x > cur) || cur + 4 : [...sizes].reverse().find((x) => x < cur) || Math.max(1, cur - 1); ed.chain().focus().setFontSize(`${next}pt`).run(); },
    changeCase: (mode) => {
      const ed = editorRef.current;
      const { from, to } = ed.state.selection;
      if (from === to) return;
      const tr = ed.state.tr;
      let first = true;
      ed.state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText) return;
        const s = Math.max(from, pos); const e = Math.min(to, pos + node.nodeSize);
        const txt = node.text.slice(s - pos, e - pos);
        let out = txt;
        if (mode === 'upper') out = txt.toLocaleUpperCase(lang); else if (mode === 'lower') out = txt.toLocaleLowerCase(lang);
        else if (mode === 'title') out = txt.toLocaleLowerCase(lang).replace(/(^|[\s\-–(„"'])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase(lang));
        else if (mode === 'toggle') out = [...txt].map((ch) => (ch === ch.toLocaleUpperCase(lang) ? ch.toLocaleLowerCase(lang) : ch.toLocaleUpperCase(lang))).join('');
        else if (mode === 'sentence') { out = txt.toLocaleLowerCase(lang).replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (m, a, b, off) => (off === 0 && !first ? m : a + b.toLocaleUpperCase(lang))); }
        first = false;
        if (out !== txt) tr.replaceWith(s, e, ed.state.schema.text(out, node.marks));
      });
      tr.setSelection(TextSelection.create(tr.doc, from, to));
      ed.view.dispatch(tr);
      ed.commands.focus();
    },
    listStyle: (ls) => {
      const ed = editorRef.current;
      if (!ed.isActive('orderedList')) ed.chain().focus().toggleOrderedList().run();
      ed.chain().focus().updateAttributes('orderedList', { listStyle: ls === 'decimal' ? null : ls }).run();
    },
    toggleMarks: () => { const v = !marks; setMarks(v); editorRef.current.commands.setFormattingMarks(v); },
    paragraphDialog: () => setDialog({ kind: 'paragraph' }),
    find: (mode) => setFind(mode),
    print: doPrint,
    pageBreak: () => editorRef.current.chain().focus().setPageBreak().run(),
    sectionBreak: (type = 'nextPage') => {
      const ed = editorRef.current;
      const $p = ed.state.selection.$from;
      const pos = $p.depth >= 1 ? $p.after(1) : ed.state.doc.content.size;
      const m = modelRef.current;
      ed.chain().focus().insertContentAt(pos, [{ type: 'sectionBreak', attrs: { type, page: JSON.parse(JSON.stringify(m.page)), titlePage: false } }, { type: 'paragraph' }]).run();
      toast(t('Section break inserted – click “Page setup of this section …” on it for landscape pages or other margins.'));
    },
    sectionSetup: (pos) => { const n = editorRef.current.state.doc.nodeAt(pos); if (n) setDialog({ kind: 'section', pos, attrs: n.attrs }); },
    customTable: async () => {
      const v = await promptDialog({ title: 'Insert table', label: 'Columns × rows', value: '3 × 3', confirm: 'Insert', validate: (x) => (/^\s*\d{1,2}\s*[x×*]\s*\d{1,3}\s*$/i.test(x) ? null : t('For example 4 × 6')) });
      if (!v) return;
      const [cols, rows] = v.split(/[x×*]/i).map((n) => Math.max(1, Math.min(60, parseInt(n, 10))));
      editorRef.current.chain().focus().insertTable({ rows, cols, withHeaderRow: true }).run();
    },
    deleteRow: () => { const ed = editorRef.current; if (track) ed.chain().focus().trackedDeleteRow().run(); else ed.chain().focus().deleteRow().run(); },
    imageFile: () => {
      const input = document.createElement('input');
      input.type = 'file'; input.accept = 'image/*';
      input.onchange = async () => { const f = input.files && input.files[0]; if (!f) return; const src = await readImage(f); const ed = editorRef.current; if (ed.isActive('image')) ed.chain().focus().updateAttributes('image', { src, ox: null }).run(); else insertImage(src, { alt: f.name.replace(/\.[^.]+$/, '') }); };
      input.click();
    },
    imageUrl: async () => {
      const url = await promptDialog({ title: 'Picture from a web address', label: 'Address of the picture', placeholder: 'https://…', confirm: 'Insert', validate: (x) => (/^https?:\/\/\S+$/i.test(x.trim()) ? null : t('Enter an address starting with https://')) });
      if (!url) return;
      let src = url.trim();
      try { const r = await fetch(src, { mode: 'cors' }); if (r.ok) { const b = await r.blob(); if (/^image\//.test(b.type)) src = await readImage(b); } } catch { /* stays linked */ }
      insertImage(src);
    },
    imageClipboard: async () => {
      try {
        const items = await navigator.clipboard.read();
        for (const it of items) { const ty = it.types.find((x) => /^image\//.test(x)); if (ty) { insertImage(await readImage(await it.getType(ty))); return; } }
        toast(t('There is no picture on the clipboard.'));
      } catch { toast(t('The browser did not allow reading the clipboard – press Ctrl+V.')); }
    },
    image: (attrs) => editorRef.current.chain().focus().updateAttributes('image', attrs).run(),
    imageWidth: async () => { const ed = editorRef.current; const v = await promptDialog({ title: 'Picture width', label: 'Width in pixels', value: String(ed.getAttributes('image').width || ''), confirm: 'Save', validate: (x) => (/^\d{1,4}$/.test(x.trim()) ? null : t('A number, e.g. 320')) }); if (v) act.image({ width: Number(v) }); },
    altText: async () => { const ed = editorRef.current; const v = await promptDialog({ title: 'Alternative text', text: t('Describe the picture for people who cannot see it.'), label: 'Description', value: ed.getAttributes('image').alt || '', confirm: 'Save' }); if (v != null) act.image({ alt: v }); },
    link: () => {
      const ed = editorRef.current;
      if (readOnly) return;
      ed.chain().focus().extendMarkRange('link').run();
      const { from, to } = ed.state.selection;
      setDialog({ kind: 'link', href: ed.getAttributes('link').href || '', text: ed.state.doc.textBetween(from, to, ' ') });
    },
    newComment: () => {
      const ed = editorRef.current;
      if (ed.state.selection.empty) { toast(t('Select the text you want to comment on first.')); return; }
      const id = uid('c');
      setComments((cs) => [...cs, { id, author: author || t('Unknown'), date: new Date().toISOString(), text: '', resolved: false, replies: [] }]);
      const end = ed.state.selection.to;
      ed.chain().setComment(id).setTextSelection(end).run();
      setSide('comments'); setSideTab('comments'); setFocusComment(id);
      markDirty();
    },
    footnote: () => { const ed = editorRef.current; ed.chain().focus().insertFootnote('').run(); setTimeout(() => setFocusNote(indexOfFootnoteBefore(ed)), 30); },
    toc: () => { const ed = editorRef.current; const hs = info.headings.filter((h) => h.level <= 3).map((h) => ({ level: h.level, text: h.text, index: h.index })); ed.chain().focus().insertContent({ type: 'toc', attrs: { title: t('Contents'), entries: hs } }).run(); },
    headerFooter: () => setDialog({ kind: 'hf' }),
    symbols: () => setDialog({ kind: 'symbols' }),
    equation: (pos = null, latex = '', display = false) => setDialog({ kind: 'equation', pos, latex, display }),
    diagram: (pos = null, node = null) => setDialog({ kind: 'diagram', pos, attrs: node ? node.attrs : null }),
    field: (pos = null, node = null) => setDialog({ kind: 'field', pos, attrs: node ? node.attrs : null }),
    quickField: (kind) => {
      const fmt = 'dd.MM.yyyy';
      const value = kind === 'page' ? String(caretPage) : fieldValue(kind, { format: fmt, pages: pageCount, title: model.meta.title || (file && file.name), author });
      editorRef.current.chain().focus().insertContent({ type: 'field', attrs: { kind, value, instr: kind === 'date' ? `DATE \\@ "${fmt}"` : '', updatedAt: new Date().toISOString() } }).run();
    },
    refreshFields: async () => { const n = await refreshFields({ silent: false }); toast(n ? t('{n} fields updated', { n }) : t('All fields are up to date.')); },
    control: async (kind) => {
      const ed = editorRef.current;
      if (kind === 'dropdown') {
        const v = await promptDialog({ title: 'Drop-down list', label: 'Choices, separated by commas', value: t('Yes, No, Maybe'), confirm: 'Insert' });
        if (!v) return;
        const items = v.split(/[,;\n]/).map((x) => x.trim()).filter(Boolean).map((x) => ({ text: x, value: x }));
        ed.chain().focus().insertContent({ type: 'contentControl', attrs: { kind: 'dropdown', items, value: '' } }).run();
      } else if (kind === 'date') ed.chain().focus().insertContent({ type: 'contentControl', attrs: { kind: 'date', value: '', format: 'dd.MM.yyyy' } }).run();
      else if (kind === 'checkbox') ed.chain().focus().insertContent({ type: 'contentControl', attrs: { kind: 'checkbox', checked: false } }).run();
      else ed.chain().focus().insertContent({ type: 'contentControl', attrs: { kind: 'text', value: '', placeholder: true } }).run();
    },
    controlText: async (pos) => {
      const ed = editorRef.current;
      const n = ed.state.doc.nodeAt(pos);
      if (!n) return;
      const v = await promptDialog({ title: n.attrs.alias || t('Text field'), label: 'Text', value: n.attrs.value || '', confirm: 'Save' });
      if (v == null) return;
      // from now on the text is normal text inside the control (the sdt mark carries it)
      const mark = ed.state.schema.marks.sdt.create({ id: uid('s'), kind: 'text', alias: n.attrs.alias, tag: n.attrs.tag, pr: n.attrs.pr ? n.attrs.pr.replace(/<w:showingPlcHdr\s*\/>/, '') : null, end: n.attrs.end });
      const tr = ed.state.tr;
      if (v) tr.replaceWith(pos, pos + n.nodeSize, ed.state.schema.text(v, [...n.marks, mark])); else tr.setNodeMarkup(pos, undefined, { ...n.attrs, value: '' });
      ed.view.dispatch(tr);
    },
    textbox: async (pos, node) => {
      const v = await promptDialog({ title: 'Text box', text: t('The box keeps its shape and position – only its text changes.'), label: 'Text', value: node.attrs.text || '', confirm: 'Save', multiline: true });
      if (v == null) return;
      const ed = editorRef.current;
      ed.view.dispatch(ed.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, text: v, textEdited: true }));
    },
    ai: (mode = 'write') => {
      if (!aiReady) { toast(t('Set up an AI provider in Hostwerk’s settings first.')); return; }
      const ed = editorRef.current;
      const { from, to } = ed.state.selection;
      const selection = from === to ? '' : ed.state.doc.textBetween(from, to, '\n', ' ');
      const context = ed.state.doc.textBetween(Math.max(0, from - 6000), from, '\n', ' ');
      const all = mode === 'summarize' || mode === 'translate' ? (selection || ed.state.doc.textBetween(0, ed.state.doc.content.size, '\n', ' ')) : context;
      setDialog({ kind: 'ai', mode, selection, context: mode === 'summarize' || mode === 'translate' ? all : context, range: { from, to } });
    },
    pageSize: (k) => setPage({ size: k, ...PAGE_SIZES[k] }),
    page: (patch) => setPage(patch),
    marginsDialog: () => setDialog({ kind: 'section', pos: null, attrs: { type: 'nextPage', page: modelRef.current.page, titlePage: modelRef.current.titlePage } }),
    toggleSpell: () => setSpell((v) => !v),
    setLang: (l) => setLang(l),
    wordCount: () => setDialog({ kind: 'count' }),
    side: (which) => { if (side && sideTab === which) setSide(null); else { setSide('comments'); setSideTab(which); } },
    thumbs: () => setSide((s) => (s === 'thumbs' ? null : 'thumbs')),
    toggleTrack: () => { const v = !track; setTrack(v); toast(v ? t('Track changes is on – your edits are marked.') : t('Track changes is off.')); markDirty(); },
    compare: () => setDialog({ kind: 'versions' }),
    collab: () => { if (collab) { onCollab(false); return; } onCollab(true, buildModel(), editorRef.current.getJSON()); },
    macros: async () => {
      const k = modelRef.current.keep || {};
      let modules = []; let err = null;
      try {
        if (k.docx && k.docx.zip) {
          const { default: JSZip } = await import('jszip');
          const zip = await JSZip.loadAsync(k.docx.zip);
          const f = Object.keys(zip.files).find((n) => /vbaProject\.bin$/i.test(n));
          if (f) modules = vbaModules(await zip.file(f).async('uint8array'));
        } else if (k.doc && initial.raw) modules = vbaModules(initial.raw, 'Macros');
      } catch (e) { err = t('The macro project could not be decoded: {m}', { m: e.message }); }
      setDialog({ kind: 'macros', modules, err });
    },
    setLayout: (l) => setLayout(l),
    focusMode: () => setFocus(true),
    toggleRuler: () => setRuler((v) => !v),
    toggleNav: () => setNav((v) => !v),
    zoom: (z) => setZoom(z),
    fitWidth,
    toggleMini: () => setMini((v) => { lsSet('of-docs-ribbon-min', v ? '0' : '1'); return !v; }),
    evenColumns: () => { const ed = editorRef.current; const tr = ed.state.tr; const $p = ed.state.selection.$from; for (let d = $p.depth; d > 0; d--) { if ($p.node(d).type.name === 'table') { const start = $p.before(d); ed.state.doc.nodesBetween(start, start + $p.node(d).nodeSize, (n, pos) => { if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') tr.setNodeMarkup(pos, undefined, { ...n.attrs, colwidth: null }); }); break; } } ed.view.dispatch(tr); },
    columnWidth: async () => { const v = await promptDialog({ title: 'Column width', label: 'Width in pixels', value: '120', confirm: 'Save', validate: (x) => (/^\d{2,4}$/.test(x.trim()) ? null : t('A number, e.g. 120')) }); if (v) editorRef.current.chain().focus().setCellAttribute('colwidth', [Number(v)]).run(); },
  };
  actRef.current = act;
  const indexOfFootnoteBefore = (ed) => { let k = -1; const pos = ed.state.selection.from; ed.state.doc.descendants((n, p) => { if (n.type.name === 'footnote' && p < pos) k++; }); return Math.max(0, k); };
  const commands = useMemo(() => docsCommands({ ai: aiReady }), [aiReady]);

  // command palette (Alt+Q) and dropped files
  useEffect(() => {
    if (!editor) return undefined;
    const offs = [];
    if (host.palette && typeof host.palette.register === 'function' && !readOnly) offs.push(host.palette.register(commands.map((c) => ({ id: `docs:${c.id}`, title: t(c.title), group: t(c.group), keys: c.keys, run: () => c.run(actRef.current, editorRef.current) }))));
    if (typeof host.insertFiles === 'function' && !readOnly) offs.push(host.insertFiles(async (files, at) => {
      const pos = at ? editorRef.current.view.posAtCoords({ left: at.x, top: at.y }) : null;
      for (const f of files) {
        if (!/^image\//.test(f.mime || '')) { toast(t('Only pictures can be dropped into a document.')); continue; }
        const b = await f.bytes();
        const src = await readImage(new Blob([b], { type: f.mime }));
        editorRef.current.chain().focus().insertContentAt(pos ? pos.pos : editorRef.current.state.selection.from, { type: 'image', attrs: { src, alt: (f.name || '').replace(/\.[^.]+$/, '') } }).run();
      }
    }));
    return () => offs.forEach((f) => typeof f === 'function' && f());
  }, [editor, commands]); // eslint-disable-line react-hooks/exhaustive-deps
  // the shell's compare view: "Apply as tracked changes"
  useEffect(() => {
    if (!editor || readOnly) return undefined;
    return registerCompareTarget(async (docA, docB) => {
      const { applyAsTrackedChanges } = await compareLib();
      const tracked = applyAsTrackedChanges(docA, docB, { author: author || t('Unknown'), date: new Date().toISOString() });
      editorRef.current.commands.setContent(tracked, true);
      setTrack(true); setSide('comments'); setSideTab('changes');
      markDirty();
    });
  }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps

  // format painter: apply on the next selection
  useEffect(() => {
    if (!editor || !painter) return undefined;
    const up = () => setTimeout(() => {
      const ed = editorRef.current;
      if (!ed || ed.state.selection.empty) return;
      const ch = ed.chain().focus().unsetAllMarks();
      for (const m of painter.marks) ch.setMark(m.type.name, m.attrs);
      const { textAlign, lineHeight, spaceBefore, spaceAfter, indent, firstLine } = painter.para;
      ch.setParagraphFormat({ lineHeight, spaceBefore, spaceAfter, indent, firstLine });
      if (textAlign) ch.setTextAlign(textAlign);
      ch.run();
      setPainter(null);
    }, 0);
    editor.view.dom.addEventListener('mouseup', up);
    return () => editor.view.dom.removeEventListener('mouseup', up);
  }, [editor, painter]);

  // keyboard: our shortcuts in the editor area; never swallow Hostwerk's global ones
  useEffect(() => {
    const el = rootRef.current;
    if (!el) return undefined;
    const key = (e) => {
      const combo = comboOf(e);
      if (HOSTWERK_KEYS.test(combo)) return;
      // the editor already handled it (its own keymap) – keep it inside the editor area
      if (e.defaultPrevented) { if (e.ctrlKey || e.metaKey || e.altKey) e.stopPropagation(); return; }
      const mod = e.ctrlKey || e.metaKey;
      const k = String(e.key).toLowerCase();
      let handled = true;
      if (mod && !e.altKey && !e.shiftKey && k === 's') { if (host.save) host.save(); }
      else if (mod && !e.altKey && !e.shiftKey && k === 'p') doPrint();
      else if (mod && !e.altKey && !e.shiftKey && k === 'f') setFind('find');
      else if (mod && !e.altKey && !e.shiftKey && k === 'h') { if (!readOnly) setFind('replace'); }
      else if (mod && e.shiftKey && !e.altKey && k === 'v' && !readOnly) act.paste(true);
      else if (mod && e.shiftKey && (k === '8' || k === '*' || e.code === 'Digit8')) act.toggleMarks();
      else if (mod && !e.altKey && !e.shiftKey && LINE_KEYS[k] && !readOnly) { const ed = editorRef.current; if (ed) ed.chain().focus().setParagraphFormat({ lineHeight: LINE_KEYS[k] }).run(); }
      else if (mod && e.altKey && k === 'm' && !readOnly) act.newComment();
      else if (mod && e.altKey && k === 'f' && !readOnly) act.footnote();
      else if (mod && e.shiftKey && k === 'e' && !readOnly) act.toggleTrack();
      else if (mod && !e.altKey && !e.shiftKey && (k === '0' || e.code === 'Digit0') && e.target.closest && !e.target.closest('input,textarea')) setZoom(100);
      else if (e.key === 'F9' && !readOnly) act.refreshFields();
      else if (e.key === 'Escape' && focus) setFocus(false);
      else if (e.key === 'Escape' && painter) setPainter(null);
      else handled = false;
      if (handled) { e.preventDefault(); e.stopPropagation(); return; }
      // everything ProseMirror handled stays inside the editor
      if (e.defaultPrevented && (mod || e.altKey)) e.stopPropagation();
    };
    el.addEventListener('keydown', key);
    return () => el.removeEventListener('keydown', key);
  });

  /* ---------- render ---------- */
  if (!editor) return <div className="of-docs of-docs-center"><Loading /></div>;
  const print = layout === 'print';
  const d = model.defaults || {};
  const paperVars = {
    '--of-font': fontCss(d.font || ds.font || 'Calibri'), '--of-size': `${d.fontSize || 11}pt`, '--of-lh': d.lineHeight || 1.15, '--of-after': `${d.spaceAfter != null ? d.spaceAfter : 8}pt`,
    '--of-paper': page.color || '#ffffff',
  };
  const lastPage = pages && pages.length ? pages[pages.length - 1] : null;
  const sheetH = print ? (lastPage ? lastPage.top + lastPage.h : box.height * MM_PX) : null;
  const sheetW = print ? Math.max(box.width * MM_PX, ...(pages || []).map((p) => p.w)) : null;
  // the body column: content box of the last section (others get their width from the pagination)
  const first = pages && pages[0];
  const bodyTop = first ? first.mt : page.margins.top * MM_PX;
  const sideOpen = side && !focus;
  // the ruler shows the section of the caret's page (only the last section's margins are dragged here)
  const caretSec = pages && pages[caretPage - 1] ? pages[caretPage - 1].section : info.sections.length;
  const rulerPage = caretSec < info.sections.length && info.sections[caretSec].page ? info.sections[caretSec].page : page;
  const onContext = (e) => {
    if (e.ctrlKey || e.metaKey || !editor.view.dom.contains(e.target)) return;
    e.preventDefault();
    const pos = editor.view.posAtCoords({ left: e.clientX, top: e.clientY });
    const { from, to } = editor.state.selection;
    if (pos && (pos.pos < from || pos.pos > to)) editor.commands.setTextSelection(pos.pos);
    setCtx({ x: e.clientX, y: e.clientY });
  };
  return (
    <div ref={rootRef} className={`of-docs ${focus ? 'of-focus' : ''} ${painter ? 'of-painting' : ''} ${readOnly ? 'of-ro' : ''}`} style={paperVars}>
      {!focus && <Ribbon editor={editor} readOnly={readOnly} tab={tab} setTab={setTab} act={act} host={host}
        s={{ painter, marks, page, spell, lang, track, changes: changes.length, comments: !!side && side === 'comments' && sideTab === 'comments', layout, ruler, nav, zoom, hasMacros: (initial.model.fidelity || []).includes('macros'), contentWidth: contentW, defaultFont: d.font, defaultSize: d.fontSize, mini, ai: aiReady, collab: !!collab, joining, canCollab: !!(host.collab && host.collab.connect) && !readOnly && !(file && file.converted), thumbs: side === 'thumbs' }} />}
      {find && <FindBar editor={editor} mode={find} setMode={setFind} readOnly={readOnly} onClose={() => { setFind(null); editor.commands.focus(); }} />}
      {focus && <div className="of-docs-focusbar"><Btn className="sm ghost" onClick={() => setFocus(false)}>Leave focus mode (Esc)</Btn></div>}
      <div className="of-docs-main">
        {nav && !focus && <NavPane headings={headings} onJump={jump} onMove={readOnly ? null : act.moveSection} readOnly={readOnly} />}
        <div className={`of-docs-scroll ${print ? 'of-print' : 'of-web'}`} ref={scrollRef} onContextMenu={onContext} onMouseDown={(e) => { if (e.target === e.currentTarget) editor.commands.focus('end'); }}>
          <div className="of-docs-zoom" style={{ zoom: zoom / 100 }}>
            {print && ruler && !focus && <Ruler page={rulerPage} box={pageBox(rulerPage)} units={gs.units} readOnly={readOnly || rulerPage !== page} onMargins={(m, done) => { if (m) setModel((x) => ({ ...x, page: { ...x.page, margins: m } })); if (done) { share('page', { ...modelRef.current.page, margins: m || modelRef.current.page.margins }); markDirty(); } }} />}
            {print ? (
              <div className="of-docs-sheet" ref={sheetRef} style={{ width: sheetW, height: sheetH }} onDoubleClick={(ev) => { if (!readOnly && ev.target.classList && ev.target.classList.contains('of-docs-page')) act.headerFooter(); }}>
                <PagesLayer pages={pages} scrollRef={scrollRef} zoom={zoom} model={model} sectionList={sectionList} title={model.meta.title || (file && file.name) || ''} gap={GAP} />
                {!pages && <div className="of-docs-page" style={{ top: 0, height: box.height * MM_PX, width: box.width * MM_PX }} aria-hidden="true" />}
                <div className="of-docs-body of-docs-paper" style={{ top: bodyTop, left: page.margins.left * MM_PX, width: contentW, columnCount: (page.columns || 1) > 1 ? page.columns : undefined }}>
                  <EditorContent editor={editor} />
                </div>
              </div>
            ) : (
              <div className="of-docs-webpage of-docs-paper" style={{ columnCount: (page.columns || 1) > 1 ? page.columns : undefined }}>
                <EditorContent editor={editor} />
              </div>
            )}
            <div className="of-docs-after" style={{ width: print ? box.width * MM_PX : undefined }}>
              <FootnotesArea notes={footnotes} readOnly={readOnly} focusIndex={focusNote} onChange={(k, text) => {
                let i = -1; const tr = editor.state.tr;
                editor.state.doc.descendants((n, pos) => { if (n.type.name === 'footnote') { i++; if (i === k) tr.setNodeMarkup(pos, undefined, { ...n.attrs, text }); } });
                editor.view.dispatch(tr);
              }} />
            </div>
          </div>
        </div>
        {sideOpen && side === 'thumbs' && <Thumbnails editor={editor} pages={print ? pages : null} current={caretPage} onJump={(i) => { const el = scrollRef.current; if (el && pages && pages[i]) el.scrollTo({ top: pages[i].top * (zoom / 100), behavior: 'smooth' }); }} />}
        {sideOpen && side === 'comments' && <CommentsPane tab={sideTab} setTab={setSideTab} readOnly={readOnly} author={author}
          comments={comments.map((c) => ({ ...c, quote: anchors.quotes[c.id] }))} anchors={anchors.a} changes={changes} activeId={(editor.state.selection.$from.marks().find((m) => m.type.name === 'comment') || { attrs: {} }).attrs.id} focusId={focusComment}
          onSelect={(id) => { const from = anchors.a[id]; if (from != null) { editor.chain().focus().setTextSelection(from).run(); const dom = editor.view.domAtPos(from + 1); const el = dom.node.nodeType === 1 ? dom.node : dom.node.parentElement; if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center', behavior: 'smooth' }); } }}
          onChange={(id, patch) => { setComments((cs) => cs.map((c) => (c.id === id ? { ...c, ...patch } : c))); markDirty(); }}
          onDelete={(id) => { editor.chain().focus().removeComment(id).run(); setComments((cs) => cs.filter((c) => c.id !== id)); markDirty(); }}
          onReply={(id, text) => { setComments((cs) => cs.map((c) => (c.id === id ? { ...c, replies: [...(c.replies || []), { author: author || t('Unknown'), date: new Date().toISOString(), text }] } : c))); markDirty(); }}
          onSelectChange={(r) => { editor.chain().focus().setTextSelection({ from: Math.max(1, r.from), to: Math.max(1, Math.min(editor.state.doc.content.size, r.to)) }).scrollIntoView().run(); }}
          onAccept={(r) => editor.chain().focus().setTextSelection({ from: Math.max(1, r.from), to: Math.max(1, Math.min(editor.state.doc.content.size, r.to)) }).acceptChange().run()}
          onReject={(r) => editor.chain().focus().setTextSelection({ from: Math.max(1, r.from), to: Math.max(1, Math.min(editor.state.doc.content.size, r.to)) }).rejectChange().run()}
          onAcceptAll={() => editor.chain().focus().acceptAllChanges().run()} onRejectAll={() => editor.chain().focus().rejectAllChanges().run()}
          onClose={() => setSide(null)} />}
      </div>
      {!focus && <StatusBar page={caretPage} pages={pageCount} stats={st} lang={lang} zoom={zoom} setZoom={setZoom} fitWidth={fitWidth} track={track} changes={changes.length} layout={layout} setLayout={setLayout} goal={goal} onWords={() => setDialog({ kind: 'count' })} peers={peers} />}

      {!readOnly && <SlashMenu editor={editor} state={slash} commands={commands} act={act} />}
      <Bubble editor={editor} act={act} readOnly={readOnly} rootRef={rootRef} ai={aiReady} />
      {ctx && <ContextMenu at={ctx} editor={editor} act={act} readOnly={readOnly} ai={aiReady} onClose={() => setCtx(null)} />}

      {dialog && dialog.kind === 'symbols' && <SymbolDialog onClose={() => setDialog(null)} onInsert={(ch) => editor.chain().focus().insertContent(ch).run()} />}
      {dialog && dialog.kind === 'link' && <LinkDialog initial={dialog} onClose={() => setDialog(null)}
        onRemove={() => { editor.chain().focus().extendMarkRange('link').unsetLink().run(); setDialog(null); }}
        onSave={(href, text) => {
          const { from, to } = editor.state.selection;
          const cur = editor.state.doc.textBetween(from, to, ' ');
          if (from === to || (text && text !== cur)) editor.chain().focus().insertContent({ type: 'text', text: text || href, marks: [{ type: 'link', attrs: { href } }] }).run();
          else editor.chain().focus().extendMarkRange('link').setLink({ href }).run();
          setDialog(null);
        }} />}
      {dialog && dialog.kind === 'hf' && <HeaderFooterEditor model={model} onClose={() => setDialog(null)} onSave={(patch) => { updateModel(patch); setDialog(null); }} />}
      {dialog && dialog.kind === 'section' && <SectionDialog attrs={dialog.attrs} units={gs.units} onClose={() => setDialog(null)} onSave={(v) => {
        if (dialog.pos == null) { setPage(v.page); updateModel({ titlePage: v.titlePage }); } else { const n = editor.state.doc.nodeAt(dialog.pos); if (n) editor.view.dispatch(editor.state.tr.setNodeMarkup(dialog.pos, undefined, { ...n.attrs, ...v })); }
        setDialog(null);
      }} />}
      {dialog && dialog.kind === 'paragraph' && <ParagraphDialog attrs={{ ...editor.getAttributes('paragraph'), ...editor.getAttributes('heading') }} onClose={() => setDialog(null)} onSave={(a) => { editor.chain().focus().setParagraphFormat(a).run(); setDialog(null); }} />}
      {dialog && dialog.kind === 'count' && <WordCountDialog stats={st} pages={pageCount} goal={goal} onGoal={(g) => { setGoalRaw(g); lsSet(goalKey, g ? String(g) : null); }} selection={(() => { const { from, to } = editor.state.selection; if (from === to) return null; const s = editor.state.doc.textBetween(from, to, '\n', ' '); return { words: (s.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-_.@]*/gu) || []).length, chars: s.length }; })()} onClose={() => setDialog(null)} />}
      {dialog && dialog.kind === 'equation' && <EquationEditor latex={dialog.latex} display={dialog.display} onClose={() => setDialog(null)} onSave={(latex, display, mathml) => {
        const attrs = { latex, display, mathml: mathml || null, omml: null, ommlLatex: null };
        if (dialog.pos != null) { const n = editor.state.doc.nodeAt(dialog.pos); if (n) editor.view.dispatch(editor.state.tr.setNodeMarkup(dialog.pos, undefined, { ...n.attrs, ...attrs, omml: n.attrs.ommlLatex === latex ? n.attrs.omml : null, ommlLatex: n.attrs.ommlLatex === latex ? n.attrs.ommlLatex : null })); } else editor.chain().focus().insertContent({ type: 'equation', attrs }).run();
        setDialog(null);
      }} />}
      {dialog && dialog.kind === 'field' && <FieldDialog host={host} initial={dialog.attrs} ctx={{ page: caretPage, pages: pageCount, title: model.meta.title || (file && file.name), author: model.meta.author || author, filename: file && file.name, words: st.words }} onClose={() => setDialog(null)} onSave={(attrs) => {
        if (dialog.pos != null) { const n = editor.state.doc.nodeAt(dialog.pos); if (n) editor.view.dispatch(editor.state.tr.setNodeMarkup(dialog.pos, undefined, { ...n.attrs, ...attrs, ox: null })); } else editor.chain().focus().insertContent({ type: 'field', attrs }).run();
        setDialog(null);
      }} />}
      {dialog && dialog.kind === 'diagram' && <DiagramDialog host={host} initial={dialog.attrs} onClose={() => setDialog(null)} onSave={(attrs) => {
        if (dialog.pos != null) { const n = editor.state.doc.nodeAt(dialog.pos); if (n) editor.view.dispatch(editor.state.tr.setNodeMarkup(dialog.pos, undefined, { ...n.attrs, ...attrs })); } else editor.chain().focus().insertContent({ type: 'diagram', attrs }).run();
        setDialog(null);
      }} />}
      {dialog && dialog.kind === 'ai' && <AiDialog host={host} mode={dialog.mode} selection={dialog.selection} context={dialog.context} onClose={() => setDialog(null)} onInsert={(text, replace) => {
        // one transaction → one Ctrl+Z
        const html = mdToHtml(text);
        const r = dialog.range;
        const ch = editor.chain().focus();
        if (replace && r && r.from !== r.to) ch.insertContentAt({ from: r.from, to: r.to }, html); else ch.insertContentAt(r ? r.to : editor.state.selection.to, html);
        ch.run();
        setDialog(null);
      }} />}
      {dialog && dialog.kind === 'versions' && <VersionPicker host={host} file={file} onClose={() => setDialog(null)} buildCurrent={async () => writeDocument(await prepareModel(buildModel()), ['docx', 'odt', 'rtf', 'md', 'html', 'txt'].includes(ext) ? ext : 'docx')} ext={ext} />}
      {dialog && dialog.kind === 'macros' && <MacrosSheet modules={dialog.modules} error={dialog.err} onClose={() => setDialog(null)} />}
    </div>
  );
}

/** Review → Compare: pick a version (or a file) – the shell's compare view shows the differences */
function VersionPicker({ host, file, ext, onClose, buildCurrent }) {
  const [list, setList] = useState(null);
  const [own, setOwn] = useState(null);
  useEffect(() => { let alive = true; office.versions(locOf(file)).then((r) => { if (alive) setList(Array.isArray(r) ? r : (r && r.versions) || []); }, () => alive && setList([])); return () => { alive = false; }; }, [file]);
  const go = async (a) => {
    if (typeof host.compare === 'function') { onClose(); try { await host.compare(a, 'current'); } catch (e) { toast.error(e); } return; }
    // older shell: our own view
    const bytes = a.bytes || await office.versionBytes(locOf(file), a.id);
    setOwn({ a: { bytes, label: a.label || t('Version') }, b: { bytes: await buildCurrent(), label: t('Now') } });
  };
  const pickFile = () => { const i = document.createElement('input'); i.type = 'file'; i.accept = '.docx,.docm,.dotx,.odt,.rtf,.doc,.md,.html,.txt'; i.onchange = async () => { const f = i.files && i.files[0]; if (!f) return; go({ bytes: new Uint8Array(await f.arrayBuffer()), label: f.name }); }; i.click(); };
  if (own) return <Modal size="xwide" title="Compare" onClose={onClose} footer={<Btn onClick={onClose}>Close</Btn>}><Compare a={own.a} b={own.b} ext={ext} host={host} /></Modal>;
  return (
    <Modal title="Compare with …" subtitle={t('Choose an older version of this file or another file – the differences show word by word and can become tracked changes.')} onClose={onClose} footer={<><Btn onClick={pickFile}>Another file …</Btn><span className="grow" /><Btn onClick={onClose}>Cancel</Btn></>}>
      {!list ? <Spinner /> : !list.length ? <p className="muted">{t('No saved versions yet – they appear after the first save.')}</p> : (
        <div className="stack of-docs-verlist">
          {list.slice(0, 40).map((v) => <button type="button" key={v.id} className="of-docs-mi" onClick={() => go({ id: v.id, time: v.time, label: v.label || new Date(v.time).toLocaleString() })}><span className="grow">{new Date(v.time).toLocaleString()}</span>{v.label && <span className="pill">{v.label}</span>}<span className="muted small">{Math.max(1, Math.round((v.size || 0) / 1024))} KB</span></button>)}
        </div>
      )}
    </Modal>
  );
}

export { MIME, PT_PX, Compare, formatDate, inlineText, TEXTBLOCKS };
