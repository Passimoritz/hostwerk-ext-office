/*
 * PDF chunk (kind 'pdf'): viewer with pdf.js, editing with pdf-lib. Formats: pdf, and xps/oxps after the shell
 * converted them (file.converted). Contract: ARCHITECTURE.md → "Editor chunk contract".
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Lock, FileWarning, X, ChevronUp, ChevronDown, CaseSensitive, WholeWord } from 'lucide-react';
import { Btn, Modal, Field, Callout, Empty, Loading, toast, confirmDialog, promptDialog } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { openPdf, pdfjsLib, pdfText, renderPageToCanvas } from '../_shared/pdfjs.js';
import { zipStore } from '../_shared/zipwrite.js';
import Viewer, { CSS_UNITS, makeLinkService, pageRotation, draftToAnnot } from './Viewer.jsx';
import Sidebar, { fmtSize, fmtPdfDate } from './Sidebar.jsx';
import Toolbar from './Toolbar.jsx';
import SignatureDialog from './signature.jsx';
import { buildDocument, createBlank, imagesToPdf, loadPdf, readMeta, listSignatures, inspect, splitPdf, SIZES } from './edit.js';
import { newId, translate, scale as scaleAnnot, center, STAMPS, approxWidth, MARKUP, turn } from './annotations.js';
import { collectWidgets, valuesFromStorage, summarize } from './forms.js';
import { pageIndex, findInPage } from './search.js';
import '../_shared/pdfjs-layers.css';
import './style.css';
import { Keep, Raw } from '../_shared/keep.jsx';

const PDF_MIME = 'application/pdf';
const baseName = (n) => String(n || 'document').replace(/\.[^.]+$/, '');
const ls = { get: (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } }, set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* off */ } } };

/** read files chosen in a picker */
function pickFiles(accept, multiple = true) {
  return new Promise((resolve) => {
    const i = document.createElement('input');
    i.type = 'file'; i.accept = accept; i.multiple = multiple;
    i.onchange = async () => resolve(await Promise.all([...i.files].map(async (f) => ({ name: f.name, type: f.type, bytes: new Uint8Array(await f.arrayBuffer()) }))));
    i.addEventListener('cancel', () => resolve([]));
    i.click();
  });
}
/** any browser image → PNG bytes (+ size) */
async function toPng(bytes, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type: type || 'image/png' }));
  try {
    const img = await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = () => rej(new Error(t('This image cannot be read.'))); im.src = url; });
    const c = document.createElement('canvas'); c.width = img.naturalWidth; c.height = img.naturalHeight;
    c.getContext('2d').drawImage(img, 0, 0);
    const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
    return { png: new Uint8Array(await blob.arrayBuffer()), w: img.naturalWidth, h: img.naturalHeight };
  } finally { URL.revokeObjectURL(url); }
}
const isPdf = (b) => b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46;
const isJpgOrPng = (b) => (b[0] === 0xff && b[1] === 0xd8) || (b[0] === 0x89 && b[1] === 0x50);

/* ---------------- dialogs ---------------- */

function Properties({ file, info, meta, editable, pageSize, count, onSave, onClose }) {
  const [m, setM] = useState({ title: meta.title || '', author: meta.author || '', subject: meta.subject || '', keywords: meta.keywords || '' });
  const changed = ['title', 'author', 'subject', 'keywords'].some((k) => (m[k] || '') !== (meta[k] || ''));
  const row = (k, v) => (v ? <tr><th>{k}</th><td><Keep>{v}</Keep></td></tr> : null);
  return (
    <Modal title="Document properties" subtitle={<Keep>{file.name}</Keep>} onClose={onClose} dirty={changed}
      footer={<><Btn onClick={onClose}>{editable ? 'Cancel' : 'Close'}</Btn>{editable && <Btn className="primary" disabled={!changed} onClick={() => { onSave(m); onClose(); }}>Apply</Btn>}</>}>
      <div className="stack">
        {['title', 'author', 'subject', 'keywords'].map((k) => (
          <Field key={k} label={{ title: t('Title'), author: t('Author'), subject: t('Subject'), keywords: t('Keywords') }[k]} hint={k === 'keywords' ? t('Separate with commas') : undefined}>
            <input className="input" value={m[k]} disabled={!editable} onChange={(e) => setM({ ...m, [k]: e.target.value })} aria-label={k} />
          </Field>
        ))}
        <table className="table of-pdf-props"><tbody>
          {row(t('File size'), fmtSize(file.size || 0))}
          {row(t('Pages'), String(count))}
          {row(t('Page size'), pageSize)}
          {row(t('PDF version'), info.PDFFormatVersion)}
          {row(t('Created with'), meta.creator || info.Creator)}
          {row(t('PDF producer'), meta.producer || info.Producer)}
          {row(t('Created'), info.CreationDate ? fmtPdfDate(info.CreationDate) : '')}
          {row(t('Modified'), info.ModDate ? fmtPdfDate(info.ModDate) : '')}
          {row(t('Protected'), info.IsEncrypted ? t('Yes') : '')}
          {row(t('Linearized (fast web view)'), info.IsLinearized ? t('Yes') : '')}
          {row(t('PDF/A'), info.pdfa ? t('Yes – kept when saving') : '')}
        </tbody></table>
      </div>
    </Modal>
  );
}

function Presenter({ pages, proxies, start, onClose }) {
  const box = useRef(null);
  const canvas = useRef(null);
  const [i, setI] = useState(start);
  useEffect(() => {
    const el = box.current;
    if (el.requestFullscreen) el.requestFullscreen().catch(() => {});
    const fs = () => { if (!document.fullscreenElement) onClose(); };
    document.addEventListener('fullscreenchange', fs);
    el.focus();
    return () => { document.removeEventListener('fullscreenchange', fs); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); };
  }, [onClose]);
  useEffect(() => {
    const mp = pages[i]; const proxy = mp && proxies.get(`${mp.src}:${mp.index}`);
    if (!proxy) return;
    const rot = pageRotation(proxy, mp, 0);
    const base = proxy.getViewport({ scale: 1, rotation: rot });
    const s = Math.min(window.innerWidth / base.width, window.innerHeight / base.height) * Math.min(2, window.devicePixelRatio || 1);
    renderPageToCanvas(proxy, canvas.current, { scale: s, rotation: rot - proxy.rotate }).catch(() => {});
  }, [i, pages, proxies]);
  const go = (d) => setI((x) => Math.max(0, Math.min(pages.length - 1, x + d)));
  return (
    <div ref={box} className="of-pdf-present" tabIndex={-1} role="dialog" aria-label="Presentation"
      onKeyDown={(e) => { if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(e.key)) { e.preventDefault(); go(1); } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(e.key)) { e.preventDefault(); go(-1); } else if (e.key === 'Home') setI(0); else if (e.key === 'End') setI(pages.length - 1); else if (e.key === 'Escape') onClose(); }}
      onClick={(e) => go(e.clientX < window.innerWidth / 4 ? -1 : 1)}>
      <canvas ref={canvas} />
      <span className="of-pdf-present-no">{i + 1} / {pages.length}</span>
    </div>
  );
}

/* ---------------- the editor ---------------- */

function Editor({ file, bytes, readOnly, settings, host }) {
  const ps = (settings && settings.pdf) || {};
  const [phase, setPhase] = useState('loading');
  const [error, setError] = useState('');
  const [retry, setRetry] = useState(0);
  const sources = useRef([]); // [{ bytes, doc }]
  const [proxies, setProxies] = useState(() => new Map());
  const [model, setModelState] = useState(null);
  const modelRef = useRef(null);
  const hist = useRef({ past: [], future: [] });
  const [histTick, setHistTick] = useState(0);
  const changed = useRef(false);
  const formsChanged = useRef(false);
  const [pdfInfo, setPdfInfo] = useState({ info: {}, meta: {}, encrypted: false, insp: null, outline: [], attachments: [], signatures: [], formInfo: null });
  const widgets = useRef(null);

  const [tab, setTab] = useState('home');
  const [tool, setToolState] = useState('text-select');
  const [color, setColor] = useState(() => ls.get('office-pdf-color', '#e5364a'));
  const [width, setWidth] = useState(2);
  const [fontSize, setFontSize] = useState(12);
  const [stamp, setStamp] = useState('APPROVED');
  const [annotMode, setAnnotModeState] = useState(() => ls.get('office-pdf-annot-mode', 'annot'));
  const [zoom, setZoom] = useState(() => { const z = ps.zoom; return z && z !== 'auto' && !Number.isNaN(Number(z)) ? Number(z) / (Number(z) > 10 ? 100 : 1) : z || 'auto'; });
  const [viewRot, setViewRot] = useState(0);
  const [scroll, setScroll] = useState(ps.scroll || 'vertical');
  // narrow windows: the sidebar would cover the pages – start closed there
  const [sidebar, setSidebar] = useState(() => ps.sidebar !== false && !(typeof matchMedia === 'function' && matchMedia('(max-width: 600px)').matches));
  const [sideTab, setSideTab] = useState('pages');
  const [night, setNight] = useState(false);
  const [current, setCurrent] = useState(null);
  const [selection, setSelection] = useState(() => new Set());
  const [selectedAnnot, setSelectedAnnot] = useState(null);
  const [pageInput, setPageInput] = useState('1');
  const [jump, setJump] = useState(null);
  const [dialog, setDialog] = useState(null);
  const [pending, setPending] = useState(null); // image/signature waiting to be placed
  const [presenting, setPresenting] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWord, setWholeWord] = useState(false);
  const [hits, setHits] = useState([]); // [{ page, i, parts }]
  const [activeHit, setActiveHit] = useState(-1);
  const textIdx = useRef(new Map());
  const [box, setBox] = useState({ w: 900, h: 700 });
  const root = useRef(null);
  const scrollRef = useRef(null);

  const editable = !readOnly && phase === 'ready' && !pdfInfo.encrypted;

  /* ----- loading ----- */
  useEffect(() => {
    let gone = false;
    setPhase('loading');
    (async () => {
      try {
        const doc = await openPdf(bytes, {
          onPassword: async (wrong) => {
            if (gone) return null;
            setPhase('password');
            const pw = await promptDialog({ title: wrong ? t('Wrong password – try again') : t('This PDF is protected'), text: t('Enter the password to open {name}.', { name: file.name }), label: t('Password'), password: true, confirm: t('Open') });
            if (pw != null) setPhase('loading');
            return pw;
          },
        });
        if (gone) { doc.destroy(); return; }
        sources.current = [{ bytes, doc }];
        const map = new Map();
        const all = await Promise.all([...Array(doc.numPages).keys()].map((i) => doc.getPage(i + 1)));
        all.forEach((p, i) => map.set(`0:${i}`, p));
        const pages = all.map((_, index) => ({ id: `p${index}`, src: 0, index, rot: 0 }));
        const m = { pages, annots: {}, meta: null, flatten: false };
        modelRef.current = m; setModelState(m); setProxies(map); setCurrent(pages[0] ? pages[0].id : null);
        hist.current = { past: [], future: [] }; changed.current = false; formsChanged.current = false;
        doc.annotationStorage.onSetModified = () => { formsChanged.current = true; host.dirty(true); };
        setPhase('ready');
        // details in the background
        const meta = await doc.getMetadata().catch(() => ({ info: {} }));
        const outline = await doc.getOutline().catch(() => null);
        const att = await doc.getAttachments().catch(() => null);
        let encrypted = !!(meta.info && meta.info.IsEncrypted);
        let lib = null;
        try { lib = await loadPdf(bytes); } catch { encrypted = true; }
        const insp = lib ? inspect(lib) : null;
        const signatures = lib ? listSignatures(lib) : [];
        widgets.current = await collectWidgets(doc).catch(() => new Map());
        if (gone) return;
        setPdfInfo({ info: { ...(meta.info || {}), pdfa: insp && insp.pdfa }, meta: lib ? readMeta(lib) : {}, encrypted, insp, outline: outline || [], attachments: att ? Object.values(att) : [], signatures, formInfo: summarize(widgets.current) });
        const items = [];
        if (encrypted && !readOnly) items.push({ level: 'warn', text: t('This PDF is protected – you can read, search and print it, editing is off.') });
        if (insp && insp.signatures.length) items.push({ level: 'warn', text: t('This PDF is digitally signed. Saving changes makes the signature invalid.') });
        if (insp && insp.xfa) items.push({ level: 'info', text: t('This PDF contains an XFA form – the basic form is shown and filled; the XFA part is kept as it is.') });
        if (insp && insp.javascript) items.push({ level: 'info', text: t('Scripts inside this PDF are not run here.') });
        if (host.fidelity) host.fidelity({ items });
      } catch (e) {
        if (gone) return;
        const lib = pdfjsLib();
        if (e && e.name === 'PasswordException') { setPhase('locked'); return; }
        setError(e instanceof lib.InvalidPDFException ? t('This file is not a valid PDF or is damaged.') : (e && e.message) || String(e));
        setPhase('error');
      }
    })();
    return () => { gone = true; };
  }, [bytes, retry]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => () => { for (const s of sources.current) { try { s.doc.destroy(); } catch { /* gone */ } } }, []);

  /* ----- model + history ----- */
  const commit = useCallback((next) => {
    if (!next || next === modelRef.current) return;
    hist.current.past.push(modelRef.current);
    if (hist.current.past.length > 200) hist.current.past.shift();
    hist.current.future = [];
    modelRef.current = next; setModelState(next); setHistTick((x) => x + 1);
    changed.current = true; host.dirty(true);
  }, [host]);
  const undo = () => { const h = hist.current; if (!h.past.length) return; h.future.push(modelRef.current); const m = h.past.pop(); modelRef.current = m; setModelState(m); setHistTick((x) => x + 1); changed.current = true; host.dirty(true); setSelectedAnnot(null); };
  const redo = () => { const h = hist.current; if (!h.future.length) return; h.past.push(modelRef.current); const m = h.future.pop(); modelRef.current = m; setModelState(m); setHistTick((x) => x + 1); changed.current = true; host.dirty(true); };

  /* ----- writing ----- */
  const build = useCallback(async (opts = {}) => {
    const m = modelRef.current;
    if (!m) return bytes;
    if (!changed.current && !formsChanged.current && !opts.force) return bytes;
    const s0 = sources.current[0];
    let forms = null;
    if (formsChanged.current) {
      if (!widgets.current) widgets.current = await collectWidgets(s0.doc);
      forms = valuesFromStorage(widgets.current, s0.doc.annotationStorage.getAll() || {});
    }
    return buildDocument({ sources: sources.current.map((s) => s.bytes), pages: opts.pages || m.pages, annots: m.annots, forms, flatten: m.flatten, meta: m.meta, mode: annotMode, author: (settings && settings.general && settings.general.author) || '' });
  }, [bytes, annotMode, settings]);

  /** pages as PNG (scale 2 ≈ 144 dpi), from the edited document */
  const renderPngs = useCallback(async (ids) => {
    const out = await build();
    const doc = out === bytes ? sources.current[0].doc : await openPdf(out);
    const m = modelRef.current;
    const res = [];
    try {
      for (const id of ids) {
        const n = m.pages.findIndex((p) => p.id === id);
        const page = out === bytes ? proxies.get(`0:${m.pages[n].index}`) : await doc.getPage(n + 1);
        const c = document.createElement('canvas');
        await renderPageToCanvas(page, c, { scale: 2 });
        const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
        res.push({ n: n + 1, bytes: new Uint8Array(await blob.arrayBuffer()) });
      }
    } finally { if (doc !== sources.current[0].doc) doc.destroy(); }
    return res;
  }, [build, bytes, proxies]);

  const text = useCallback(async () => {
    const out = await build();
    const doc = out === bytes ? sources.current[0].doc : await openPdf(out);
    try { return await pdfText(doc); } finally { if (doc !== sources.current[0].doc) doc.destroy(); }
  }, [build, bytes]);

  useEffect(() => {
    if (phase !== 'ready') return;
    host.setSaver(async () => build());
    const formats = [
      { ext: 'pdf', label: t('PDF document') },
      { ext: 'png', label: t('PNG image (current page)') },
      { ext: 'zip', label: t('PNG images of all pages (zip)') },
      { ext: 'txt', label: t('Plain text') },
    ];
    host.setExporter({
      formats,
      run: async (ext) => {
        if (ext === 'pdf') return build();
        if (ext === 'png') return (await renderPngs([currentRef.current]))[0].bytes;
        if (ext === 'zip') { const imgs = await renderPngs(modelRef.current.pages.map((p) => p.id)); const w = String(imgs.length).length; return zipStore(imgs.map((x) => ({ name: `${baseName(file.name)}-${String(x.n).padStart(w, '0')}.png`, bytes: x.bytes }))); }
        if (ext === 'txt') return new TextEncoder().encode(await text());
        throw new Error(t('This format is not available here.'));
      },
    });
  }, [phase, build, renderPngs, text]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ----- layout / zoom ----- */
  useEffect(() => {
    const el = scrollRef.current; if (!el) return undefined;
    const ro = new ResizeObserver(() => setBox({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, [phase, presenting]);

  const currentRef = useRef(current); currentRef.current = current;
  const pages = model ? model.pages : [];
  const curIdx = Math.max(0, pages.findIndex((p) => p.id === current));
  const curPage = pages[curIdx];
  const curProxy = curPage ? proxies.get(`${curPage.src}:${curPage.index}`) : null;
  const scale = useMemo(() => {
    if (typeof zoom === 'number') return zoom;
    if (!curProxy) return 1;
    const vp = curProxy.getViewport({ scale: CSS_UNITS, rotation: pageRotation(curProxy, curPage, viewRot) });
    const horizontalPad = 40; const verticalPad = 32;
    const fitW = (box.w - horizontalPad) / vp.width; const fitH = (box.h - verticalPad) / vp.height;
    if (zoom === 'page-fit') return Math.max(0.1, Math.min(fitW, fitH));
    if (zoom === 'page-width') return Math.max(0.1, fitW);
    const landscape = vp.width > vp.height;
    return Math.max(0.1, Math.min(landscape ? Math.min(fitW, fitH) : fitW, 1.25));
  }, [zoom, curProxy, curPage, viewRot, box]);
  const roundScale = Math.round(scale * 1000) / 1000;

  const zoomStep = useCallback((d) => {
    setZoom((z) => {
      const s = typeof z === 'number' ? z : roundScale;
      const steps = [0.25, 0.33, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];
      if (d > 0) return steps.find((x) => x > s + 0.001) || 4;
      return [...steps].reverse().find((x) => x < s - 0.001) || 0.25;
    });
    setJump({ id: currentRef.current, n: Date.now() });
  }, [roundScale]);

  useEffect(() => { setPageInput(String(curIdx + 1)); host.status(t('Page {n} of {count}', { n: curIdx + 1, count: pages.length })); }, [curIdx, pages.length]); // eslint-disable-line react-hooks/exhaustive-deps

  const goPage = (i, offset) => { const p = pages[Math.max(0, Math.min(pages.length - 1, i))]; if (!p) return; setCurrent(p.id); setJump({ id: p.id, n: Date.now(), offset }); };
  const onCurrent = useCallback((id) => setCurrent(id), []);

  /* ----- links, outline ----- */
  const goDest = useCallback(async (srcIdx, dest) => {
    const doc = sources.current[srcIdx] && sources.current[srcIdx].doc;
    if (!doc) return;
    try {
      const explicit = typeof dest === 'string' ? await doc.getDestination(dest) : dest;
      if (!Array.isArray(explicit)) return;
      const ref = explicit[0];
      const index = typeof ref === 'object' && ref ? await doc.getPageIndex(ref) : Number(ref);
      const m = modelRef.current;
      const n = m.pages.findIndex((p) => p.src === srcIdx && p.index === index);
      if (n < 0) { toast(t('That page is no longer in this document.'), 'info'); return; }
      // explicit[1] = /XYZ left top zoom → scroll to "top"
      let offset = 0;
      const proxy = proxies.get(`${srcIdx}:${index}`);
      if (explicit[1] && explicit[1].name === 'XYZ' && proxy && explicit[3] != null) {
        const vp = proxy.getViewport({ scale: roundScale * CSS_UNITS, rotation: pageRotation(proxy, m.pages[n], viewRot) });
        offset = Math.max(0, vp.convertToViewportPoint(explicit[2] || 0, explicit[3])[1] - 8);
      }
      goPage(n, offset);
    } catch { /* broken destination */ }
  }, [proxies, roundScale, viewRot]); // eslint-disable-line react-hooks/exhaustive-deps

  const openLink = useCallback((url) => {
    if (!/^(https?:|mailto:)/i.test(url)) { toast(t('This link type is not opened for safety: {url}', { url }), 'warn'); return; }
    if (host.openLink) host.openLink(url);
    else toast(t('Links open in Hostwerk once the Office shell supports it: {url}', { url }), 'info');
  }, [host]);

  const linkService = useMemo(() => makeLinkService({
    onDest: (d) => goDest(0, d),
    onPage: (i) => { const n = modelRef.current.pages.findIndex((p) => p.src === 0 && p.index === i); if (n >= 0) goPage(n); },
    onNamed: (a) => { const n = pages.findIndex((p) => p.id === currentRef.current); if (a === 'NextPage') goPage(n + 1); else if (a === 'PrevPage') goPage(n - 1); else if (a === 'FirstPage') goPage(0); else if (a === 'LastPage') goPage(pages.length - 1); else if (a === 'Print') doPrint(); },
    openLink,
  }), [goDest, pages, openLink]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ----- annotations ----- */
  const setAnnots = (pageId, fn) => {
    const m = modelRef.current;
    const list = fn(m.annots[pageId] || []);
    commit({ ...m, annots: { ...m.annots, [pageId]: list } });
  };
  const addAnnot = (pageId, a) => { const id = newId(); setAnnots(pageId, (l) => [...l, { ...a, id }]); return id; };
  const findAnnot = (id) => { for (const [pid, list] of Object.entries(modelRef.current.annots)) { const a = list.find((x) => x.id === id); if (a) return { pid, a }; } return null; };
  const deleteAnnot = () => { const f = selectedAnnot && findAnnot(selectedAnnot); if (!f) return; setAnnots(f.pid, (l) => l.filter((x) => x.id !== f.a.id)); setSelectedAnnot(null); };

  const editText = async (pid, a) => {
    const v = await promptDialog({ title: a.type === 'note' ? t('Note') : t('Text'), label: t('Text'), value: a.text || '', confirm: t('Apply') });
    if (v == null) return;
    if (!v.trim()) { setAnnots(pid, (l) => l.filter((x) => x.id !== a.id)); return; }
    setAnnots(pid, (l) => l.map((x) => (x.id === a.id ? { ...x, text: v } : x)));
  };

  const onPointer = useCallback(async (ev) => {
    const m = modelRef.current;
    if (ev.kind === 'blank') { setSelectedAnnot(null); return; }
    if (ev.kind === 'drawn') {
      const a = draftToAnnot(ev.draft, { color, width });
      if (!a) return;
      const tiny = a.type === 'ink' ? a.paths[0].length < 4 : a.type === 'line' || a.type === 'arrow' ? Math.hypot(a.x2 - a.x1, a.y2 - a.y1) < 3 : a.w < 3 || a.h < 3;
      if (tiny) return;
      addAnnot(ev.page, a);
      return;
    }
    if (ev.kind === 'place') {
      const [x, y] = ev.pt; const rot = ev.rot;
      /** bottom-left corner of a box of height h whose top-left is where the user clicked */
      const below = (h) => { const [ox, oy] = turn(0, -h, rot); return [x + ox, y + oy]; };
      if (tool === 'text') {
        const v = await promptDialog({ title: t('Text box'), label: t('Text'), value: '', confirm: t('Add') });
        if (!v || !v.trim()) return;
        const [bx, by] = below(fontSize * 0.95);
        const id = addAnnot(ev.page, { type: 'text', x: bx, y: by, text: v, size: fontSize, color, rot });
        setSelectedAnnot(id);
      } else if (tool === 'note') {
        const v = await promptDialog({ title: t('Sticky note'), label: t('Note'), value: '', confirm: t('Add') });
        if (!v || !v.trim()) return;
        addAnnot(ev.page, { type: 'note', x: x - 10, y: y - 10, text: v, color: '#ffd400' });
      } else if (tool === 'stamp') {
        const st = STAMPS.find((s) => s.text === stamp) || STAMPS[0];
        const label = t(st.text);
        const h = 30; const w = approxWidth(label, h * 0.5) * 1.25 + 24;
        const [bx, by] = below(h);
        const id = addAnnot(ev.page, { type: 'stamp', x: bx, y: by, w, h, text: label, color: st.color, rot });
        setSelectedAnnot(id);
      } else if ((tool === 'image' || tool === 'signature') && pending) {
        const maxW = pending.signature ? 160 : 240;
        const s = Math.min(1, maxW / pending.w);
        const w = pending.w * s * (pending.signature ? 1 : 0.75); const h = pending.h * s * (pending.signature ? 1 : 0.75);
        const [bx, by] = below(h);
        const id = addAnnot(ev.page, { type: 'image', x: bx, y: by, w, h, rot, png: pending.png, signature: !!pending.signature });
        setSelectedAnnot(id); setPending(null); setToolState('select');
      }
      return;
    }
    if (ev.kind === 'grab' || ev.kind === 'scale') {
      const f = findAnnot(ev.id); if (!f) return;
      setSelectedAnnot(ev.id);
      if (!editable) return;
      const start = ev.pt; const orig = f.a; const pid = f.pid; let last = orig; let moved = false;
      const c = center(orig);
      const d0 = Math.hypot(start[0] - c[0], start[1] - c[1]) || 1;
      const mv = (e2) => {
        const p = ev.toPdf(e2);
        moved = true;
        last = ev.kind === 'grab' ? translate(orig, p[0] - start[0], p[1] - start[1]) : scaleAnnot(orig, Math.max(0.1, Math.hypot(p[0] - c[0], p[1] - c[1]) / d0));
        const mm = modelRef.current;
        const next = { ...mm, annots: { ...mm.annots, [pid]: mm.annots[pid].map((x) => (x.id === orig.id ? last : x)) } };
        modelRef.current = next; setModelState(next);
      };
      const end = () => {
        window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', end);
        if (!moved) {
          // click on a text/note: second click edits
          if ((orig.type === 'text' || orig.type === 'note') && ev.e.detail >= 2) editText(pid, orig);
          return;
        }
        // one history step for the whole drag
        const after = modelRef.current; modelRef.current = m; commit(after);
      };
      window.addEventListener('pointermove', mv); window.addEventListener('pointerup', end);
    }
  }, [tool, color, width, fontSize, stamp, pending, editable, commit]); // eslint-disable-line react-hooks/exhaustive-deps

  const onTextSelection = useCallback((pageId, quads, txt) => {
    if (!quads.length) return;
    addAnnot(pageId, { type: tool, quads, color: tool === 'highlight' ? (color === '#e5364a' ? '#ffd400' : color) : color, text: txt });
  }, [tool, color]); // eslint-disable-line react-hooks/exhaustive-deps

  const setTool = (id) => { setToolState(id); if (id !== 'select') setSelectedAnnot(null); if (id !== 'image' && id !== 'signature') setPending(null); };

  /* ----- pages ----- */
  const targets = () => (selection.size > 1 ? pages.filter((p) => selection.has(p.id)).map((p) => p.id) : [current]);
  const rotatePages = (ids, d) => { const m = modelRef.current; const s = new Set(ids); commit({ ...m, pages: m.pages.map((p) => (s.has(p.id) ? { ...p, rot: (((p.rot || 0) + d) % 360 + 360) % 360 } : p)) }); };
  const deletePages = async (ids) => {
    const m = modelRef.current;
    if (ids.length >= m.pages.length) { toast(t('A PDF needs at least one page.'), 'warn'); return; }
    if (ids.length > 1 && !(await confirmDialog({ title: t('Delete {n} pages?', { n: ids.length }), text: t('You can undo this with Ctrl+Z.'), confirm: t('Delete'), danger: true }))) return;
    const s = new Set(ids);
    const left = m.pages.filter((p) => !s.has(p.id));
    commit({ ...m, pages: left });
    if (s.has(current)) { const n = Math.min(curIdx, left.length - 1); setCurrent(left[n].id); }
    setSelection(new Set());
  };
  const movePagesTo = (ids, targetId, after) => {
    const m = modelRef.current; const s = new Set(ids);
    if (s.has(targetId)) return;
    const moving = m.pages.filter((p) => s.has(p.id)); const rest = m.pages.filter((p) => !s.has(p.id));
    let at = rest.findIndex((p) => p.id === targetId); if (after) at += 1;
    commit({ ...m, pages: [...rest.slice(0, at), ...moving, ...rest.slice(at)] });
  };
  const movePages = (d) => {
    const m = modelRef.current; const ids = new Set(targets());
    const arr = [...m.pages];
    const idx = arr.map((p, i) => (ids.has(p.id) ? i : -1)).filter((i) => i >= 0);
    if (!idx.length || (d < 0 && idx[0] === 0) || (d > 0 && idx[idx.length - 1] === arr.length - 1)) return;
    const order = d < 0 ? idx : [...idx].reverse();
    for (const i of order) { const j = i + d; [arr[i], arr[j]] = [arr[j], arr[i]]; }
    commit({ ...m, pages: arr });
    setJump({ id: current, n: Date.now() });
  };
  /** add a PDF (bytes) as a new source and insert its pages at position `at` */
  const addSource = async (pdfBytes, at) => {
    const doc = await openPdf(pdfBytes);
    const src = sources.current.length;
    sources.current.push({ bytes: pdfBytes, doc });
    const all = await Promise.all([...Array(doc.numPages).keys()].map((i) => doc.getPage(i + 1)));
    setProxies((old) => { const n = new Map(old); all.forEach((p, i) => n.set(`${src}:${i}`, p)); return n; });
    const m = modelRef.current;
    const added = all.map((_, index) => ({ id: `s${src}p${index}`, src, index, rot: 0 }));
    const pos = at == null ? m.pages.length : at;
    commit({ ...m, pages: [...m.pages.slice(0, pos), ...added, ...m.pages.slice(pos)] });
    return added;
  };
  const insertBlank = async () => {
    const vp = curProxy ? curProxy.getViewport({ scale: 1 }) : null;
    const size = vp ? [vp.width, vp.height] : SIZES.A4;
    const added = await addSource(await createBlank({ size }), curIdx + 1);
    setCurrent(added[0].id); setJump({ id: added[0].id, n: Date.now() });
  };
  const filesToPdf = async (files) => {
    const out = [];
    for (const f of files) {
      if (isPdf(f.bytes)) out.push(f.bytes);
      else if (/^image\//.test(f.type) || /\.(png|jpe?g|gif|webp|bmp)$/i.test(f.name)) {
        const img = isJpgOrPng(f.bytes) ? f.bytes : (await toPng(f.bytes, f.type)).png;
        out.push(await imagesToPdf([{ bytes: img }]));
      } else toast(t('{name} is neither a PDF nor an image.', { name: f.name }), 'warn');
    }
    return out;
  };
  const insertFile = async (atEnd = false) => {
    const files = await pickFiles('application/pdf,.pdf,image/*', true);
    if (!files.length) return;
    let at = atEnd ? null : curIdx + 1; let first = null;
    for (const b of await filesToPdf(files)) {
      try {
        const added = await addSource(b, at);
        if (!first) first = added[0];
        if (at != null) at += added.length;
      } catch (e) { toast.error(e.name === 'PasswordException' ? t('Protected PDFs cannot be inserted.') : e.message); }
    }
    if (first) { setCurrent(first.id); setJump({ id: first.id, n: Date.now() }); toast.success(t('Pages inserted')); }
  };
  const extract = async () => {
    const ids = targets();
    const keep = pages.filter((p) => ids.includes(p.id));
    const out = await build({ pages: keep, force: true });
    host.download(out, `${baseName(file.name)} (${t('pages')} ${keep.map((p) => pages.indexOf(p) + 1).join(', ')}).pdf`, PDF_MIME);
  };
  const split = async () => {
    const v = await promptDialog({ title: t('Split PDF'), text: t('Every part gets this many pages. You get a zip with all parts.'), label: t('Pages per part'), value: '1', confirm: t('Split'), validate: (x) => (/^\d+$/.test(x) && Number(x) >= 1 ? null : t('Enter a number of 1 or more')) });
    if (v == null) return;
    const all = await build({ force: true });
    const parts = await splitPdf(all, Number(v));
    const w = String(parts.length).length;
    host.download(zipStore(parts.map((b, i) => ({ name: `${baseName(file.name)}-${String(i + 1).padStart(w, '0')}.pdf`, bytes: b }))), `${baseName(file.name)}-split.zip`, 'application/zip');
  };

  /* ----- misc actions ----- */
  async function doPrint() {
    const out = await build();
    host.print(new Blob([out], { type: PDF_MIME }));
  }
  const toWord = async () => {
    const ok = await confirmDialog({ title: t('Convert to Word?'), text: t('The Office engine reads a PDF like a drawing: every line becomes its own text box, fonts are approximated. Good for copying text and small changes – for longer editing, prefer the original document.'), confirm: t('Convert') });
    if (!ok) return;
    const out = await host.convert(await build(), 'pdf', 'docx');
    host.download(out, `${baseName(file.name)}.docx`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  };
  const pickImage = async () => {
    const [f] = await pickFiles('image/*', false);
    if (!f) return;
    const p = await toPng(f.bytes, f.type);
    setPending({ ...p, signature: false }); setToolState('image');
    toast(t('Click on the page where the image should go.'), 'info');
  };

  /* ----- search ----- */
  const runSearch = useCallback(async (q) => {
    if (!q) { setHits([]); setActiveHit(-1); return; }
    const m = modelRef.current; const out = [];
    for (const mp of m.pages) {
      const key = `${mp.src}:${mp.index}`;
      if (!textIdx.current.has(key)) { const proxy = proxies.get(key); if (!proxy) continue; textIdx.current.set(key, pageIndex((await proxy.getTextContent({ disableNormalization: true })).items)); }
      for (const h of findInPage(textIdx.current.get(key), q, { matchCase, wholeWord })) out.push({ page: mp.id, ...h });
    }
    setHits(out);
    const startAt = out.findIndex((h) => m.pages.findIndex((p) => p.id === h.page) >= curIdx);
    setActiveHit(out.length ? Math.max(0, startAt) : -1);
  }, [proxies, matchCase, wholeWord, curIdx]);
  useEffect(() => { if (!searchOpen) return undefined; const tm = setTimeout(() => runSearch(query), 250); return () => clearTimeout(tm); }, [query, matchCase, wholeWord, searchOpen]); // eslint-disable-line react-hooks/exhaustive-deps
  const hitsByPage = useMemo(() => {
    const m = new Map();
    hits.forEach((h) => { if (!m.has(h.page)) m.set(h.page, []); m.get(h.page).push(h); });
    return m;
  }, [hits]);
  const active = useMemo(() => {
    if (activeHit < 0 || !hits[activeHit]) return null;
    const h = hits[activeHit];
    return { page: h.page, i: hitsByPage.get(h.page).indexOf(h) };
  }, [activeHit, hits, hitsByPage]);
  useEffect(() => { if (active && active.page !== currentRef.current) { const n = pages.findIndex((p) => p.id === active.page); goPage(n); } }, [active]); // eslint-disable-line react-hooks/exhaustive-deps
  const stepHit = (d) => { if (!hits.length) return; setActiveHit((i) => (i + d + hits.length) % hits.length); };
  const searchInput = useRef(null);
  const openSearch = () => { setSearchOpen(true); setTimeout(() => searchInput.current && searchInput.current.select(), 0); };

  /* ----- keyboard ----- */
  const onKey = (e) => {
    const k = e.key; const mod = e.ctrlKey || e.metaKey;
    const inField = /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName) || e.target.isContentEditable;
    if (mod && k.toLowerCase() === 'f') { e.preventDefault(); openSearch(); return; }
    if (mod && k.toLowerCase() === 'p') { e.preventDefault(); doPrint(); return; }
    if (mod && k.toLowerCase() === 's') { e.preventDefault(); host.save(); return; }
    if (mod && (k === '+' || k === '=')) { e.preventDefault(); zoomStep(1); return; }
    if (mod && k === '-') { e.preventDefault(); zoomStep(-1); return; }
    if (mod && k === '0') { e.preventDefault(); setZoom('page-fit'); return; }
    if (inField) return;
    if (editable && mod && k.toLowerCase() === 'z' && !e.shiftKey) { e.preventDefault(); undo(); return; }
    if (editable && mod && (k.toLowerCase() === 'y' || (k.toLowerCase() === 'z' && e.shiftKey))) { e.preventDefault(); redo(); return; }
    if (mod) return;
    if (k === 'F5') { e.preventDefault(); setPresenting(true); return; }
    if ((k === 'Delete' || k === 'Backspace') && editable && selectedAnnot) { e.preventDefault(); deleteAnnot(); return; }
    if (k === 'Escape') { setSelectedAnnot(null); setPending(null); if (tool !== 'text-select') setToolState('text-select'); if (searchOpen) setSearchOpen(false); return; }
    if (k === 'PageDown' && scroll !== 'vertical') { e.preventDefault(); goPage(curIdx + 1); return; }
    if (k === 'PageUp' && scroll !== 'vertical') { e.preventDefault(); goPage(curIdx - 1); return; }
    if ((k === 'ArrowRight' || k === 'ArrowDown') && scroll === 'page') { e.preventDefault(); goPage(curIdx + 1); return; }
    if ((k === 'ArrowLeft' || k === 'ArrowUp') && scroll === 'page') { e.preventDefault(); goPage(curIdx - 1); return; }
    if (k === 'Home' && e.target === scrollRef.current) { e.preventDefault(); goPage(0); return; }
    if (k === 'End' && e.target === scrollRef.current) { e.preventDefault(); goPage(pages.length - 1); return; }
    if (k.toLowerCase() === 'h') setTool(tool === 'hand' ? 'text-select' : 'hand');
    else if (k.toLowerCase() === 'r' && !e.shiftKey) setViewRot((r) => (r + 90) % 360);
    else if (editable && k.toLowerCase() === 'v') setTool('select');
    else if (editable && k.toLowerCase() === 'p') { setTab('annotate'); setTool('ink'); } else if (editable && k.toLowerCase() === 't') { setTab('annotate'); setTool('text'); }
  };
  useEffect(() => { if (phase === 'ready' && root.current) root.current.focus({ preventScroll: true }); }, [phase]);

  /* ----- render ----- */
  if (phase === 'loading' || phase === 'password') return <div className="of-pdf-root"><Loading text={t('Opening PDF …')} /></div>;
  if (phase === 'locked') return <div className="of-pdf-root"><Empty icon={Lock} title={t('This PDF is protected')} text={t('It opens with its password.')} action={<Btn className="primary" onClick={() => setRetry((x) => x + 1)}>Enter password …</Btn>} /></div>;
  if (phase === 'error') return <div className="of-pdf-root"><Empty icon={FileWarning} title={t('This PDF cannot be shown')} text={error} action={<Btn onClick={() => host.download(bytes, file.name, PDF_MIME)}>Download</Btn>} /></div>;

  const h = hist.current;
  void histTick;
  const st = {
    tab, editable, tool, color, width, fontSize, stamp, annotMode, zoom, scroll, sidebar, night, searchOpen,
    pageNo: curIdx + 1, pageCount: pages.length, pageInput, canUndo: h.past.length > 0, canRedo: h.future.length > 0,
    selectedAnnot, hasConverter: !!host.hasConverter, selCount: selection.size, formInfo: pdfInfo.formInfo, flatten: model.flatten,
  };
  const act = {
    setTab, setTool, setColor: (c) => { setColor(c); ls.set('office-pdf-color', c); }, setWidth, setFontSize, setStamp,
    setAnnotMode: (v) => { setAnnotModeState(v); ls.set('office-pdf-annot-mode', v); if (changed.current) host.dirty(true); },
    setZoom: (z) => { setZoom(z); setJump({ id: current, n: Date.now() }); }, zoomStep, setScroll: (v) => { setScroll(v); setJump({ id: current, n: Date.now() }); },
    toggleSidebar: () => setSidebar((x) => !x), toggleNight: () => setNight((x) => !x), toggleSearch: () => (searchOpen ? setSearchOpen(false) : openSearch()),
    prev: () => goPage(curIdx - 1), next: () => goPage(curIdx + 1), setPageInput,
    goInput: () => { const n = parseInt(pageInput, 10); if (n >= 1 && n <= pages.length) goPage(n - 1); else setPageInput(String(curIdx + 1)); },
    rotateView: () => setViewRot((r) => (r + 90) % 360), present: () => setPresenting(true), print: doPrint,
    properties: () => setDialog('props'), toWord, undo, redo, deleteAnnot, pickImage,
    insertBlank, insertFile: () => insertFile(false), mergeFiles: () => insertFile(true), rotatePages: (d) => rotatePages(targets(), d),
    movePages, deletePages: () => deletePages(targets()), extract, split,
    signature: () => setDialog('sign'), setFlatten: (v) => commit({ ...modelRef.current, flatten: v }),
  };
  const meta = { ...pdfInfo.meta, ...(model.meta || {}) };
  const pageSize = curProxy ? (() => { const v = curProxy.getViewport({ scale: 1 }); return `${Math.round(v.width / 72 * 25.4)} × ${Math.round(v.height / 72 * 25.4)} mm`; })() : '';
  const curSigned = pdfInfo.insp && pdfInfo.insp.signatures.length > 0;

  return (
    <div className={`of-pdf-root${night ? ' night' : ''}`} ref={root} tabIndex={-1} onKeyDown={onKey}>
      <Toolbar st={st} act={act} />
      {searchOpen && (
        <div className="of-pdf-search" role="search">
          <input ref={searchInput} className="input" value={query} placeholder="Search in the document" aria-label="Search in the document" autoFocus
            onChange={(e) => setQuery(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); stepHit(e.shiftKey ? -1 : 1); } else if (e.key === 'Escape') { setSearchOpen(false); setHits([]); root.current.focus(); } }} />
          <button type="button" className={`btn sm ghost icon${matchCase ? ' on' : ''}`} aria-pressed={matchCase} title="Match case" aria-label="Match case" onClick={() => setMatchCase(!matchCase)}><CaseSensitive size={16} /></button>
          <button type="button" className={`btn sm ghost icon${wholeWord ? ' on' : ''}`} aria-pressed={wholeWord} title="Whole words" aria-label="Whole words" onClick={() => setWholeWord(!wholeWord)}><WholeWord size={16} /></button>
          <span className="small muted of-pdf-hitcount">{query ? (hits.length ? t('{i} of {n}', { i: activeHit + 1, n: hits.length }) : t('Nothing found.')) : ''}</span>
          <button type="button" className="btn sm ghost icon" title="Previous match (Shift+Enter)" aria-label="Previous match" disabled={!hits.length} onClick={() => stepHit(-1)}><ChevronUp size={16} /></button>
          <button type="button" className="btn sm ghost icon" title="Next match (Enter)" aria-label="Next match" disabled={!hits.length} onClick={() => stepHit(1)}><ChevronDown size={16} /></button>
          <button type="button" className="btn sm ghost icon" title="Close search (Esc)" aria-label="Close search" onClick={() => { setSearchOpen(false); setHits([]); }}><X size={16} /></button>
        </div>
      )}
      {curSigned && editable && !host.fidelity && (changed.current || formsChanged.current) && <div className="of-pdf-callout"><Callout kind="warn" title={t('Signed PDF')}>{t('Saving changes makes the digital signature invalid.')}</Callout></div>}
      {pending && <div className="of-pdf-callout"><Callout kind="info" dismissible={false}>{t('Click on the page where it should go. Esc cancels.')}</Callout></div>}
      <div className="of-pdf-body">
        {sidebar && (
          <Sidebar tab={sideTab} onTab={setSideTab} pages={pages} proxies={proxies} current={current} selection={selection}
            onSelect={(s, id) => { setSelection(s); const n = pages.findIndex((p) => p.id === id); goPage(n); }}
            editable={editable} onMove={movePagesTo} onRotate={rotatePages} onDelete={deletePages}
            outline={pdfInfo.outline} onOutline={(it) => { if (it.dest) goDest(0, it.dest); else if (it.url) openLink(it.url); }}
            attachments={pdfInfo.attachments} onAttachment={(a) => host.download(a.content, a.filename, 'application/octet-stream')} signatures={pdfInfo.signatures} />
        )}
        <Viewer pages={pages} proxies={proxies} docs={sources.current.map((s) => s.doc)} scale={roundScale} viewRot={viewRot} scrollMode={scroll}
          current={current} onCurrent={onCurrent} annots={model.annots} selectedId={selectedAnnot} tool={tool} editable={editable} renderForms={!readOnly && !pdfInfo.encrypted}
          hitsByPage={hitsByPage} activeHit={active} linkService={linkService} onPointer={onPointer} onTextSelection={onTextSelection}
          scrollRef={scrollRef} night={night} onZoomWheel={zoomStep} jump={jump} />
      </div>
      <div className="of-pdf-status small muted">
        <span>{t('Page {n} of {count}', { n: curIdx + 1, count: pages.length })}</span>
        <span>{Math.round(roundScale * 100)} %</span>
        {viewRot ? <span>{viewRot}°</span> : null}
        {readOnly ? <span>{t('Read only')}</span> : pdfInfo.encrypted ? <span><Lock size={12} /> {t('Protected – editing is off')}</span> : null}
        {file.converted && <span>{t('Converted from {ext}', { ext: String(file.converted.from).toUpperCase() })}</span>}
        <span className="grow" />
        {MARKUP.includes(tool) && <span>{t('Select text to mark it.')}</span>}
        {tool === 'select' && <span>{t('Drag to move · corner to resize · double-click to edit text · Del deletes')}</span>}
      </div>
      {dialog === 'props' && <Properties file={file} info={pdfInfo.info} meta={meta} editable={editable} pageSize={pageSize} count={pages.length} onClose={() => setDialog(null)} onSave={(m) => commit({ ...modelRef.current, meta: m })} />}
      {dialog === 'sign' && <SignatureDialog defaultName={ps.signatureName || (settings && settings.general && settings.general.author) || ''} onClose={() => setDialog(null)} onPick={(sig) => { setPending({ ...sig, signature: true }); setToolState('signature'); }} />}
      {presenting && <Presenter pages={pages} proxies={proxies} start={curIdx} onClose={() => setPresenting(false)} />}
    </div>
  );
}

/* ---------------- file browser preview ---------------- */
function Preview({ bytes }) {
  const c = useRef(null);
  const [state, setState] = useState('loading');
  useEffect(() => {
    let doc = null; let gone = false;
    openPdf(bytes).then(async (d) => {
      doc = d; if (gone) return;
      const page = await d.getPage(1);
      const base = page.getViewport({ scale: 1 });
      await renderPageToCanvas(page, c.current, { scale: Math.min(2, 480 / base.width) });
      if (!gone) setState('ok');
    }).catch((e) => { if (!gone) setState(e && e.name === 'PasswordException' ? 'locked' : 'err'); });
    return () => { gone = true; if (doc) doc.destroy(); };
  }, [bytes]);
  return (
    <div className="of-pdf-preview">
      <canvas ref={c} hidden={state !== 'ok'} />
      {state === 'loading' && <Loading inline />}
      {state === 'locked' && <span className="small muted"><Lock size={14} /> {t('Protected PDF')}</span>}
      {state === 'err' && <span className="small muted">{t('No preview')}</span>}
    </div>
  );
}

export default {
  kind: 'pdf',
  Editor,
  Preview,
  create: async ({ template, settings } = {}) => {
    if (template && template !== 'blank') throw new Error(t('Unknown template: {id}', { id: template }));
    const size = settings && settings.docs && SIZES[settings.docs.pageSize] ? settings.docs.pageSize : 'A4';
    return createBlank({ size, landscape: !!(settings && settings.docs && settings.docs.orientation === 'landscape') });
  },
  toText: async (bytes) => {
    const doc = await openPdf(bytes);
    try { return await pdfText(doc); } finally { doc.destroy(); }
  },
};
