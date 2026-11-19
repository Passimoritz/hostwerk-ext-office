/*
 * The documents editor: tiptap + our extensions, ribbon, print layout with real page boxes (header/footer, page
 * numbers), web layout, ruler, navigation pane, comments/changes panel, footnotes, find & replace, status bar,
 * saver/exporter/print for the shell. Everything Hostwerk-global (Ctrl+K, Ctrl+J, Ctrl+1…9, Ctrl+N, Ctrl+, …)
 * stays free; our shortcuts only act inside the editor area.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
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
import { Loading, Callout, Btn, toast, promptDialog, confirmDialog } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Ribbon, fontCss } from './Ribbon.jsx';
import { NavPane, CommentsPane, FootnotesArea, FindBar, StatusBar, Ruler, setToast } from './panels.jsx';
import { SymbolDialog, LinkDialog, HeaderFooterDialog, MarginsDialog, ParagraphDialog, WordCountDialog, EquationDialog, MacrosSheet } from './dialogs.jsx';
import { FontSize, ParagraphFormat, RunKeep, WordKeys } from './extensions/format.js';
import { PageBreak, Footnote, Toc, Equation, Callout as CalloutNode, Preserved, ImageEx, Comment } from './extensions/nodes.js';
import { Insertion, Deletion, FormatChange, TrackChanges, changeRanges } from './extensions/track.js';
import { Search, FormattingMarks, Pagination } from './extensions/view.js';
import { readDocument, writeDocument, EXPORT_FORMATS, MIME } from './io/index.js';
import { standaloneHtml } from './io/render-html.js';
import { PAGE_SIZES, MM_PX, PT_PX, pageBox, stats as docStats, headingsOf, uid, fillFields, sizeName } from './io/pm.js';
import { vbaModules } from './io/vba.js';
import { Keep } from '../_shared/keep.jsx';

setToast((m) => toast(m));
// Ctrl+1/2/5 are Word's line spacing keys (single/double/1.5) – inside the editor they stay with Word (ARCHITECTURE.md
// "Shortcuts"); the other digits keep jumping to Hostwerk's pages
const HOSTWERK_KEYS = /^(ctrl\+(k|j|n|,|[346789])|ctrl\+shift\+(s|x|r)|alt\+arrow(left|right)|f1)$/;
const LINE_KEYS = { 1: 1, 2: 2, 5: 1.5 };
const comboOf = (e) => `${e.ctrlKey || e.metaKey ? 'ctrl+' : ''}${e.altKey ? 'alt+' : ''}${e.shiftKey ? 'shift+' : ''}${String(e.key).toLowerCase()}`;
const FIDELITY = {
  macros: 'Macros (kept unchanged, never run here)', smartart: 'SmartArt graphics (kept unchanged)', chart: 'Charts (kept unchanged)', shape: 'Shapes and drawings (kept unchanged)', textbox: 'Text boxes (kept unchanged)', ole: 'Embedded objects (kept unchanged)', object: 'Other embedded content (kept unchanged)',
  field: 'Fields (kept, result shown)', contentControls: 'Content controls (text kept, the control itself is removed)', trackedChanges: 'Tracked changes (shown and kept – accept or reject them in Review)', paragraphChanges: 'Paragraph formatting changes (accepted when saving)',
  sections: 'Several sections with their own page setup (kept)', headerLayout: 'Pictures or tables in the header/footer (kept as long as you do not edit the header/footer here)', endnotes: 'Endnotes (shown as footnotes)', signature: 'Digital signature (becomes invalid when you save)', images: 'Pictures in formats the browser cannot show (EMF/WMF)', altChunk: 'Embedded HTML/RTF chunks', indexes: 'Indexes (shown as text)', binaryFormat: 'Word 97–2003 format – saved as .docx, or back to .doc with the Office engine',
};

export function DocsEditor({ file, bytes, readOnly, settings = {}, host = {} }) {
  const [loaded, setLoaded] = useState(null);
  const [error, setError] = useState(null);
  const ext = String((file && file.converted && file.converted.to) || (file && file.ext) || 'docx').toLowerCase();
  useEffect(() => {
    let alive = true;
    setLoaded(null); setError(null);
    readDocument(bytes || new Uint8Array(0), ext, settings).then((r) => { if (alive) setLoaded({ ...r, raw: bytes }); }, (e) => { if (alive) setError({ text: e.message || String(e), code: e.code || null }); });
    return () => { alive = false; };
  }, [bytes, ext]); // eslint-disable-line react-hooks/exhaustive-deps
  if (error) {
    // the readers write English messages – show them in Hostwerk's language; a Word 6/95 file needs the Office engine
    const old = error.code === 'DOC_TOO_OLD';
    return (
      <div className="of-docs of-docs-center">
        <Callout kind={old ? 'info' : 'err'} title={old ? 'This document needs the Office engine' : 'This document could not be opened'} dismissible={false}>
          <div className="stack of-docs-gap">
            <span>{t(error.text)}</span>
            {old && typeof host.ensureEngine === 'function' && <div><Btn className="primary sm" onClick={() => host.ensureEngine(t('“{name}” is a Word 6/95 document – the Office engine opens it exactly.', { name: (file && file.name) || '' }))}>Set up the Office engine …</Btn></div>}
          </div>
        </Callout>
      </div>
    );
  }
  if (!loaded) return <div className="of-docs of-docs-center"><Loading text="Opening document …" /></div>;
  return <EditorBody file={file} ext={ext} initial={loaded} readOnly={!!readOnly} settings={settings} host={host} />;
}

function EditorBody({ file, ext, initial, readOnly, settings, host }) {
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
  const [pages, setPages] = useState(1);
  const [track, setTrack] = useState(!!initial.model.trackChanges);
  const [painter, setPainter] = useState(null);
  const [focusComment, setFocusComment] = useState(null);
  const [focusNote, setFocusNote] = useState(null);
  const [, setTick] = useState(0);
  const rootRef = useRef(null);
  const scrollRef = useRef(null);
  const sheetRef = useRef(null);
  const page = model.page;
  const box = pageBox(page);
  const P = box.height * MM_PX;
  const GAP = 24;
  const contentW = (box.width - page.margins.left - page.margins.right) * MM_PX;
  const contentH = (box.height - page.margins.top - page.margins.bottom) * MM_PX;
  const smart = ds.smartQuotes !== false;
  const de = (lang || '').startsWith('de');

  const editor = useEditor({
    editable: !readOnly,
    extensions: [
      StarterKit.configure({ strike: false, heading: { levels: [1, 2, 3, 4, 5, 6] }, history: { depth: 500 } }),
      Strike.extend({ addKeyboardShortcuts: () => ({}) }),
      Underline, TextStyle, Color, FontFamily, FontSize,
      Highlight.configure({ multicolor: true }),
      TextAlign.extend({ addKeyboardShortcuts: () => ({}) }).configure({ types: ['heading', 'paragraph'] }),
      LinkExt.extend({ inclusive: false }).configure({ openOnClick: false, autolink: ds.autoLinks !== false, linkOnPaste: true, HTMLAttributes: { target: null, rel: null, class: 'of-docs-link' } }),
      Table.configure({ resizable: !readOnly, lastColumnResizable: false, allowTableNodeSelection: true }), TableRow, TableCell, TableHeader,
      TaskList, TaskItem.configure({ nested: true }),
      SubscriptExt.extend({ addKeyboardShortcuts: () => ({}) }), SuperscriptExt.extend({ addKeyboardShortcuts: () => ({}) }),
      Placeholder.configure({ placeholder: readOnly ? '' : t('Start typing – # for a heading, - for a list, [ ] for a checklist') }),
      ...(smart ? [Typography.configure(de ? { openDoubleQuote: '„', closeDoubleQuote: '“', openSingleQuote: '‚', closeSingleQuote: '‘' } : {})] : []),
      ParagraphFormat, RunKeep, WordKeys.configure({ onLink: () => actRef.current.link(), onPageBreak: () => actRef.current.pageBreak(), onTrack: () => { if (!readOnly) actRef.current.toggleTrack(); }, onMarks: () => actRef.current.toggleMarks() }),
      PageBreak, Footnote, Toc.configure({ onJump: (i) => actRef.current.jump(i), onUpdate: () => actRef.current.updateToc() }), Equation.configure({ onEdit: (pos, latex) => actRef.current.equation(pos, latex) }),
      CalloutNode, Preserved, ImageEx.configure({ allowBase64: true, inline: true }), Comment,
      Insertion, Deletion, FormatChange, TrackChanges, Search, FormattingMarks, Pagination,
    ],
    content: initial.html != null ? initial.html : model.doc,
    editorProps: {
      attributes: { class: 'of-docs-pm', spellcheck: spell ? 'true' : 'false', lang },
      handleClick: (view, pos, e) => {
        const target = e.target && e.target.closest && e.target.closest('a[href]');
        if (target && (e.ctrlKey || e.metaKey || readOnly)) { openLink(target.getAttribute('href')); return true; }
        const fn = e.target && e.target.closest && e.target.closest('.of-docs-fn');
        if (fn) { const all = [...view.dom.querySelectorAll('.of-docs-fn')]; setFocusNote(all.indexOf(fn)); }
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
    onUpdate: ({ transaction }) => { if (!transaction.getMeta('ofPages') && transaction.docChanged) markDirty(); },
  }, []);
  const editorRef = useRef(editor);
  editorRef.current = editor;

  const dirtyRef = useRef(false);
  const markDirty = useCallback(() => { if (readOnly) return; dirtyRef.current = true; if (host.dirty) host.dirty(true); }, [readOnly, host]);
  const openLink = (href) => {
    if (!href) return;
    if (href.startsWith('#')) { const want = decodeURIComponent(href.slice(1)).toLowerCase(); const hs = headingsOf(editorRef.current.getJSON()); const h = hs.find((x) => x.text.toLowerCase().replace(/\s+/g, '-') === want || x.text.toLowerCase() === want); if (h) actRef.current.jump(h.index); return; }
    if (host.openLink) host.openLink(href); else toast(t('Links open in Hostwerk – this Hostwerk version cannot open them yet.'));
  };
  const readImage = (f) => new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsDataURL(f); });
  const insertImage = (src, extra = {}) => { const ed = editorRef.current; if (!ed) return; ed.chain().focus().setImage({ src, ...extra }).run(); };

  // editor DOM attributes that change at runtime
  useEffect(() => { if (!editor) return; editor.view.dom.setAttribute('spellcheck', spell ? 'true' : 'false'); editor.view.dom.setAttribute('lang', lang); }, [editor, spell, lang]);
  useEffect(() => { if (editor) editor.setEditable(!readOnly); }, [editor, readOnly]);
  useEffect(() => { if (editor) { editor.storage.trackChanges.enabled = track && !readOnly; } }, [editor, track, readOnly]);
  // pagination
  useEffect(() => {
    if (!editor) return;
    editor.storage.ofPages.onPages = setPages;
    editor.commands.setPagination({ enabled: layout === 'print' && (page.columns || 1) === 1, pageH: P, contentH, gap: GAP });
  }, [editor, layout, P, contentH, page.columns]);
  // re-render the chrome on selection changes
  useEffect(() => {
    if (!editor) return undefined;
    const h = () => setTick((x) => x + 1);
    editor.on('selectionUpdate', h); editor.on('transaction', h);
    return () => { editor.off('selectionUpdate', h); editor.off('transaction', h); };
  }, [editor]);
  // fidelity report to the shell
  useEffect(() => {
    const items = (initial.model.fidelity || []).map((k) => t(FIDELITY[k] || k));
    if (host.fidelity) host.fidelity({ items, kept: true });
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  // html/md sources: pick up the imported document into the model once
  useEffect(() => { if (editor && initial.html != null) setModel((m) => ({ ...m, doc: editor.getJSON() })); }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps

  const setZoom = (z) => setZoomRaw(Math.max(25, Math.min(300, Math.round(z))));
  const fitWidth = () => { const el = scrollRef.current; if (!el) return; const avail = el.clientWidth - 48; setZoom((avail / (box.width * MM_PX)) * 100); };
  // narrow window: start with the page fitted to the width (no sideways scrolling)
  useEffect(() => {
    const h = setTimeout(() => { const el = scrollRef.current; if (el && el.clientWidth < 700 && el.clientWidth - 48 < box.width * MM_PX * (zoom / 100)) fitWidth(); }, 150);
    return () => clearTimeout(h);
  }, [editor]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- the current document as a model ---------- */
  const buildModel = () => {
    const ed = editorRef.current;
    const doc = ed.getJSON();
    const anchored = new Set();
    ed.state.doc.descendants((n) => { for (const m of n.marks) if (m.type.name === 'comment') anchored.add(m.attrs.id); });
    const m = modelRef.current;
    return { ...m, doc, comments: commentsRef.current.filter((c) => anchored.has(c.id)).map(({ quote, ...c }) => c), trackChanges: track, meta: { ...m.meta, author: m.meta.author || author } };
  };

  /* ---------- saver / exporter / status ---------- */
  useEffect(() => {
    if (!editor) return;
    if (host.setSaver) host.setSaver(async () => {
      const m = buildModel();
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
        const m = buildModel();
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

  /* ---------- derived state ---------- */
  const json = editor ? editor.state.doc : null;
  const st = useMemo(() => (editor ? docStats(editor.getJSON()) : { words: 0, chars: 0, charsNoSpaces: 0, paragraphs: 0 }), [json]); // eslint-disable-line react-hooks/exhaustive-deps
  const headings = useMemo(() => (editor ? headingsOf(editor.getJSON()) : []), [json]); // eslint-disable-line react-hooks/exhaustive-deps
  const footnotes = useMemo(() => { const out = []; if (editor) editor.state.doc.descendants((n) => { if (n.type.name === 'footnote') out.push(n.attrs.text || ''); }); return out; }, [json]); // eslint-disable-line react-hooks/exhaustive-deps
  const anchors = useMemo(() => {
    const a = {}; const quotes = {};
    if (editor) editor.state.doc.descendants((n, pos) => { if (!n.isText) return; for (const m of n.marks) if (m.type.name === 'comment') { if (a[m.attrs.id] == null) a[m.attrs.id] = pos; quotes[m.attrs.id] = (quotes[m.attrs.id] || '') + n.text; } });
    return { a, quotes };
  }, [json]); // eslint-disable-line react-hooks/exhaustive-deps
  const changes = useMemo(() => (editor ? changeRanges(editor.state.doc).map((r) => ({ ...r, text: editor.state.doc.textBetween(r.from, r.to, ' ', ' ') })) : []), [json]); // eslint-disable-line react-hooks/exhaustive-deps
  const activeComment = editor ? (editor.state.selection.$from.marks().find((m) => m.type.name === 'comment') || {}).attrs : null;
  const caretPage = (() => {
    if (!editor || layout !== 'print' || !sheetRef.current) return 1;
    try {
      const c = editor.view.coordsAtPos(editor.state.selection.head);
      const r = sheetRef.current.getBoundingClientRect();
      const k = r.width / (sheetRef.current.offsetWidth || 1);
      return Math.min(pages, Math.max(1, Math.floor((c.top - r.top) / k / (P + GAP)) + 1));
    } catch { return 1; }
  })();
  useEffect(() => {
    if (host.status) host.status({ text: t('Page {a} of {b} · {n} words', { a: caretPage, b: layout === 'print' ? pages : 1, n: st.words.toLocaleString() }), trackChanges: track, changes: changes.length });
  }, [caretPage, pages, st.words, track, changes.length, layout]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- actions ---------- */
  const updateModel = (patch) => { setModel((m) => ({ ...m, ...patch })); markDirty(); };
  const setPage = (patch) => { const p = { ...modelRef.current.page, ...patch }; if (patch.width || patch.height) p.size = sizeName(p.width, p.height); updateModel({ page: p }); };
  const jump = (index) => {
    const ed = editorRef.current;
    let k = 0; let target = null;
    ed.state.doc.descendants((n, pos) => { if (target != null) return false; if (n.type.name === 'heading') { if (k === index) target = pos; k++; return false; } return true; });
    if (target == null) return;
    ed.chain().focus().setTextSelection(target + 1).run();
    const dom = ed.view.nodeDOM(target);
    if (dom && dom.scrollIntoView) dom.scrollIntoView({ block: 'start', behavior: 'smooth' });
  };
  const updateToc = () => {
    const ed = editorRef.current;
    const hs = headingsOf(ed.getJSON(), 3).map((h) => ({ level: h.level, text: h.text, index: h.index }));
    const tr = ed.state.tr;
    ed.state.doc.descendants((n, pos) => { if (n.type.name === 'toc') tr.setNodeMarkup(pos, undefined, { ...n.attrs, entries: hs }); });
    ed.view.dispatch(tr);
  };
  const act = {
    jump, updateToc,
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
      setPainter({ marks: (ed.state.storedMarks || $f.marks()).filter((m) => !['comment', 'insertion', 'deletion', 'link'].includes(m.type.name)), para: { ...$f.parent.attrs } });
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
    customTable: async () => {
      const v = await promptDialog({ title: 'Insert table', label: 'Columns × rows', value: '3 × 3', confirm: 'Insert', validate: (x) => (/^\s*\d{1,2}\s*[x×*]\s*\d{1,3}\s*$/i.test(x) ? null : t('For example 4 × 6')) });
      if (!v) return;
      const [cols, rows] = v.split(/[x×*]/i).map((n) => Math.max(1, Math.min(60, parseInt(n, 10))));
      editorRef.current.chain().focus().insertTable({ rows, cols, withHeaderRow: true }).run();
    },
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
    footnote: () => { const ed = editorRef.current; ed.chain().focus().insertFootnote('').run(); setTimeout(() => setFocusNote(footnotesRef.current.length - 1 >= 0 ? indexOfFootnoteBefore(ed) : 0), 30); },
    toc: () => { const ed = editorRef.current; const hs = headingsOf(ed.getJSON(), 3).map((h) => ({ level: h.level, text: h.text, index: h.index })); ed.chain().focus().insertContent({ type: 'toc', attrs: { title: t('Contents'), entries: hs } }).run(); },
    headerFooter: () => setDialog({ kind: 'hf' }),
    symbols: () => setDialog({ kind: 'symbols' }),
    equation: (pos = null, latex = '') => setDialog({ kind: 'equation', pos, latex }),
    pageSize: (k) => setPage({ size: k, ...PAGE_SIZES[k] }),
    page: (patch) => setPage(patch),
    marginsDialog: () => setDialog({ kind: 'margins' }),
    toggleSpell: () => setSpell((v) => !v),
    setLang: (l) => setLang(l),
    wordCount: () => setDialog({ kind: 'count' }),
    side: (which) => { if (side && sideTab === which) setSide(null); else { setSide('comments'); setSideTab(which); } },
    toggleTrack: () => { const v = !track; setTrack(v); toast(v ? t('Track changes is on – your edits are marked.') : t('Track changes is off.')); markDirty(); },
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
    evenColumns: () => { const ed = editorRef.current; const tr = ed.state.tr; const $p = ed.state.selection.$from; for (let d = $p.depth; d > 0; d--) { if ($p.node(d).type.name === 'table') { const start = $p.before(d); ed.state.doc.nodesBetween(start, start + $p.node(d).nodeSize, (n, pos) => { if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') tr.setNodeMarkup(pos, undefined, { ...n.attrs, colwidth: null }); }); break; } } ed.view.dispatch(tr); },
    columnWidth: async () => { const v = await promptDialog({ title: 'Column width', label: 'Width in pixels', value: '120', confirm: 'Save', validate: (x) => (/^\d{2,4}$/.test(x.trim()) ? null : t('A number, e.g. 120')) }); if (v) editorRef.current.chain().focus().setCellAttribute('colwidth', [Number(v)]).run(); },
  };
  const actRef = useRef(act);
  actRef.current = act;
  const footnotesRef = useRef(footnotes);
  footnotesRef.current = footnotes;
  const indexOfFootnoteBefore = (ed) => { let k = -1; const pos = ed.state.selection.from; ed.state.doc.descendants((n, p) => { if (n.type.name === 'footnote' && p < pos) k++; }); return Math.max(0, k); };

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
  const hf = (x, i, where) => {
    if (!x || !x.text || (model.titlePage && i === 0)) return null;
    return <div className={`of-docs-hf of-hf-${where} of-al-${x.align || 'center'}`} style={{ [where === 'top' ? 'top' : 'bottom']: Math.max(4, (where === 'top' ? page.margins.top : page.margins.bottom) * MM_PX * 0.4), left: page.margins.left * MM_PX, right: page.margins.right * MM_PX }}><Keep>{fillFields(x.text, { page: i + 1, pages, title: model.meta.title || (file && file.name) || '', date: new Date().toLocaleDateString(lang) })}</Keep></div>;
  };
  const sheetH = print ? pages * (P + GAP) - GAP : null;
  const sideOpen = side && !focus;
  return (
    <div ref={rootRef} className={`of-docs ${focus ? 'of-focus' : ''} ${painter ? 'of-painting' : ''} ${readOnly ? 'of-ro' : ''}`} style={paperVars}>
      {!focus && <Ribbon editor={editor} readOnly={readOnly} tab={tab} setTab={setTab} act={act}
        s={{ painter, marks, page, spell, lang, track, changes: changes.length, comments: !!side && sideTab === 'comments', layout, ruler, nav, zoom, hasMacros: (initial.model.fidelity || []).includes('macros'), contentWidth: contentW, defaultFont: d.font, defaultSize: d.fontSize }} />}
      {find && <FindBar editor={editor} mode={find} setMode={setFind} readOnly={readOnly} onClose={() => { setFind(null); editor.commands.focus(); }} />}
      {focus && <div className="of-docs-focusbar"><Btn className="sm ghost" onClick={() => setFocus(false)}>Leave focus mode (Esc)</Btn></div>}
      <div className="of-docs-main">
        {nav && !focus && <NavPane headings={headings} onJump={jump} />}
        <div className={`of-docs-scroll ${print ? 'of-print' : 'of-web'}`} ref={scrollRef} onMouseDown={(e) => { if (e.target === e.currentTarget) editor.commands.focus('end'); }}>
          <div className="of-docs-zoom" style={{ zoom: zoom / 100 }}>
            {print && ruler && !focus && <Ruler page={page} box={box} units={gs.units} readOnly={readOnly} onMargins={(m, done) => { if (m) setModel((x) => ({ ...x, page: { ...x.page, margins: m } })); if (done) markDirty(); }} />}
            {print ? (
              <div className="of-docs-sheet" ref={sheetRef} style={{ width: box.width * MM_PX, height: sheetH }}>
                {Array.from({ length: pages }, (_, i) => (
                  <div key={i} className="of-docs-page" style={{ top: i * (P + GAP), height: P, width: box.width * MM_PX }} aria-hidden="true">
                    {hf(model.header, i, 'top')}
                    {hf(model.footer, i, 'bottom')}
                  </div>
                ))}
                <div className="of-docs-body of-docs-paper" style={{ top: page.margins.top * MM_PX, left: page.margins.left * MM_PX, width: contentW, columnCount: (page.columns || 1) > 1 ? page.columns : undefined }}>
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
        {sideOpen && <CommentsPane tab={sideTab} setTab={setSideTab} readOnly={readOnly} author={author}
          comments={comments.map((c) => ({ ...c, quote: anchors.quotes[c.id] }))} anchors={anchors.a} changes={changes} activeId={activeComment && activeComment.id} focusId={focusComment}
          onSelect={(id) => { const from = anchors.a[id]; if (from != null) { editor.chain().focus().setTextSelection(from).run(); const dom = editor.view.domAtPos(from + 1); const el = dom.node.nodeType === 1 ? dom.node : dom.node.parentElement; if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center', behavior: 'smooth' }); } }}
          onChange={(id, patch) => { setComments((cs) => cs.map((c) => (c.id === id ? { ...c, ...patch } : c))); markDirty(); }}
          onDelete={(id) => { editor.chain().focus().removeComment(id).run(); setComments((cs) => cs.filter((c) => c.id !== id)); markDirty(); }}
          onReply={(id, text) => { setComments((cs) => cs.map((c) => (c.id === id ? { ...c, replies: [...(c.replies || []), { author: author || t('Unknown'), date: new Date().toISOString(), text }] } : c))); markDirty(); }}
          onSelectChange={(r) => { editor.chain().focus().setTextSelection({ from: r.from, to: r.to }).scrollIntoView().run(); }}
          onAccept={(r) => editor.chain().focus().setTextSelection({ from: r.from, to: r.to }).acceptChange().run()}
          onReject={(r) => editor.chain().focus().setTextSelection({ from: r.from, to: r.to }).rejectChange().run()}
          onAcceptAll={() => editor.chain().focus().acceptAllChanges().run()} onRejectAll={() => editor.chain().focus().rejectAllChanges().run()}
          onClose={() => setSide(null)} />}
      </div>
      {!focus && <StatusBar page={caretPage} pages={print ? pages : 1} stats={st} lang={lang} zoom={zoom} setZoom={setZoom} fitWidth={fitWidth} track={track} changes={changes.length} layout={layout} setLayout={setLayout} />}

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
      {dialog && dialog.kind === 'hf' && <HeaderFooterDialog header={model.header} footer={model.footer} titlePage={model.titlePage} onClose={() => setDialog(null)} onSave={(h, f, first) => { updateModel({ header: h, footer: f, titlePage: first }); setDialog(null); }} />}
      {dialog && dialog.kind === 'margins' && <MarginsDialog page={{ ...page, width: box.width, height: box.height }} units={gs.units} onClose={() => setDialog(null)} onSave={(v) => { const land = v.width > v.height; setPage({ margins: v.margins, width: Math.min(v.width, v.height), height: Math.max(v.width, v.height), orientation: land ? 'landscape' : 'portrait' }); setDialog(null); }} />}
      {dialog && dialog.kind === 'paragraph' && <ParagraphDialog attrs={{ ...editor.getAttributes('paragraph'), ...editor.getAttributes('heading') }} onClose={() => setDialog(null)} onSave={(a) => { editor.chain().focus().setParagraphFormat(a).run(); setDialog(null); }} />}
      {dialog && dialog.kind === 'count' && <WordCountDialog stats={st} pages={print ? pages : 1} selection={(() => { const { from, to } = editor.state.selection; if (from === to) return null; const s = editor.state.doc.textBetween(from, to, '\n', ' '); return { words: (s.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-_.@]*/gu) || []).length, chars: s.length }; })()} onClose={() => setDialog(null)} />}
      {dialog && dialog.kind === 'equation' && <EquationDialog latex={dialog.latex} onClose={() => setDialog(null)} onSave={(latex) => {
        if (dialog.pos != null) { const n = editor.state.doc.nodeAt(dialog.pos); if (n) editor.view.dispatch(editor.state.tr.setNodeMarkup(dialog.pos, undefined, { ...n.attrs, latex })); } else editor.chain().focus().insertContent({ type: 'equation', attrs: { latex } }).run();
        setDialog(null);
      }} />}
      {dialog && dialog.kind === 'macros' && <MacrosSheet modules={dialog.modules} error={dialog.err} onClose={() => setDialog(null)} />}
    </div>
  );
}

export { MIME, PT_PX };
