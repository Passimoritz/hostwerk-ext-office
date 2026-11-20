/*
 * PDF chunk (kind 'pdf'): viewer with pdf.js, editing with pdf-lib. Formats: pdf, and xps/oxps after the shell
 * converted them (file.converted). Contract: ARCHITECTURE.md → "Editor chunk contract" and "Office 1.1".
 *
 * Opening: protected PDFs are decrypted with our own security handler (security.js) – they stay editable and are
 * saved with the same protection; our own annotations are taken out of the file again (annotations.js
 * readOwnAnnotations) so they stay editable; pdf.js gets the first page at once, the other pages load in the
 * background and on demand (Viewer.jsx renders progressively). Tools: OCR, real redaction, compare, compress,
 * protect (tools.jsx).
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Lock, FileWarning, X, ChevronUp, ChevronDown, CaseSensitive, WholeWord, Copy, Highlighter, Underline, Strikethrough, EyeOff, Search,
  Type, StickyNote, Stamp, RotateCw, RotateCcw, Trash2, ZoomIn, ZoomOut, Maximize, Printer, Info, Pencil, KeyRound, Languages,
} from 'lucide-react';
import { Btn, Modal, Field, Callout, Empty, Loading, toast, confirmDialog, promptDialog } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { pdfjsLib, pdfText, renderPageToCanvas } from '../_shared/pdfjs.js';
import { zipStore } from '../_shared/zipwrite.js';
import Viewer, { CSS_UNITS, makeLinkService, pageRotation, draftToAnnot } from './Viewer.jsx';
import Sidebar, { fmtSize, fmtPdfDate } from './Sidebar.jsx';
import Toolbar from './Toolbar.jsx';
import SignatureDialog from './signature.jsx';
import ContextMenu from './menu.jsx';
import { ProtectDialog, CompressDialog, OcrDialog, RedactPanel, RedactResult, CompareDialog, CompareView, recodeImage } from './tools.jsx';
import { buildDocument, createBlank, imagesToPdf, loadPdf, listSignatures, inspect, splitPdf, SIZES, missingGlyphs } from './edit.js';
import { newId, translate, scale as scaleAnnot, center, STAMPS, approxWidth, MARKUP, turn, readOwnAnnotations, hasMarker, textBoxSize, textLines } from './annotations.js';
import { collectWidgets, valuesFromStorage } from './forms.js';
import { pageIndex, findInPage } from './search.js';
import { isEncryptedBytes, openProtected, encryptPdf } from './security.js';
import { openDoc, cjkAssets, needsCmaps, latinFontBytes, cjkFontBytes, forgetCjk } from './open.js';
import { ensureComponent } from './components.js';
import { createTesseract, OCR_PACKS } from './ocr.js';
import { addTextLayer, imageWordsToPdf } from './textlayer.js';
import { findSensitive, hitRects, keptWords, applyRedactions, verifyRedaction, scanFile } from './redact.js';
import { compressPdf } from './compress.js';
import '../_shared/pdfjs-layers.css';
import './style.css';
import './style-11.css';
import { Keep } from '../_shared/keep.jsx';

const PDF_MIME = 'application/pdf';
const baseName = (n) => String(n || 'document').replace(/\.[^.]+$/, '');
const ls = { get: (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } }, set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* off */ } } };
const idle = (fn) => (typeof requestIdleCallback === 'function' ? requestIdleCallback(fn, { timeout: 2000 }) : setTimeout(fn, 200));

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
const canvasBytes = async (c, type = 'image/png', q) => new Uint8Array(await (await new Promise((r) => c.toBlob(r, type, q))).arrayBuffer());

/** text with characters no font covers → PNG drawn with the system fonts (4× resolution) */
async function rasterText(a) {
  const { w, h } = textBoxSize(a);
  const k = 4; const c = document.createElement('canvas');
  c.width = Math.max(1, Math.ceil(w * k)); c.height = Math.max(1, Math.ceil(h * k));
  const ctx = c.getContext('2d');
  ctx.fillStyle = a.color || '#000'; ctx.textBaseline = 'alphabetic';
  ctx.font = `${a.size * k}px "Noto Sans CJK SC", "Microsoft YaHei", "Yu Gothic", "Malgun Gothic", "PingFang SC", sans-serif`;
  textLines(a).forEach((l, i) => ctx.fillText(l, 0, (a.size * 0.95 + i * a.size * 1.2) * k, c.width));
  return { png: await canvasBytes(c), w, h };
}

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
          {row(t('Protected'), info.protection || '')}
          {row(t('Linearized (fast web view)'), info.IsLinearized ? t('Yes') : '')}
          {row(t('PDF/A'), info.pdfa ? t('Yes – kept when saving') : '')}
        </tbody></table>
      </div>
    </Modal>
  );
}

function Presenter({ pages, getProxy, start, onClose }) {
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
    const mp = pages[i]; if (!mp) return;
    let gone = false;
    getProxy(mp).then((proxy) => {
      if (gone || !proxy) return;
      const rot = pageRotation(proxy, mp, 0);
      const base = proxy.getViewport({ scale: 1, rotation: rot });
      const s = Math.min(window.innerWidth / base.width, window.innerHeight / base.height) * Math.min(2, window.devicePixelRatio || 1);
      renderPageToCanvas(proxy, canvas.current, { scale: s, rotation: rot - proxy.rotate }).catch(() => {});
    });
    return () => { gone = true; };
  }, [i, pages, getProxy]);
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
  const sources = useRef([]); // [{ bytes, doc, name }]
  const [proxies, setProxies] = useState(() => new Map());
  const proxyBuf = useRef(new Map()); const proxyFlush = useRef(0); const proxyWait = useRef(new Map());
  const [defaultSize, setDefaultSize] = useState([595, 842]);
  const [model, setModelState] = useState(null);
  const modelRef = useRef(null);
  const hist = useRef({ past: [], future: [] });
  const [histTick, setHistTick] = useState(0);
  const changed = useRef(false);
  const formsChanged = useRef(false);
  const [pdfInfo, setPdfInfo] = useState({ info: {}, meta: {}, outline: [], attachments: [], signatures: [], formInfo: null, insp: null });
  const [access, setAccess] = useState({ protectedFile: false, canEdit: true, viewOnly: false }); // viewOnly: pdf.js opened it, our handler could not
  const [needFonts, setNeedFonts] = useState(false);
  const widgets = useRef(null);

  const [tab, setTab] = useState('home');
  const [tool, setToolState] = useState('text-select');
  const [color, setColorState] = useState(() => ls.get('office-pdf-color', '#e5364a'));
  const [width, setWidthState] = useState(2);
  const [fontSize, setFontSizeState] = useState(12);
  const [stamp, setStamp] = useState('APPROVED');
  const [annotMode, setAnnotModeState] = useState(() => ls.get('office-pdf-annot-mode', 'annot'));
  const [zoom, setZoom] = useState(() => { const z = ps.zoom; return z && z !== 'auto' && !Number.isNaN(Number(z)) ? Number(z) / (Number(z) > 10 ? 100 : 1) : z || 'auto'; });
  const [viewRot, setViewRot] = useState(0);
  const [scroll, setScroll] = useState(ps.scroll || 'vertical');
  // narrow windows: the sidebar would cover the pages – start closed there
  const [sidebar, setSidebar] = useState(() => ps.sidebar !== false && !(typeof matchMedia === 'function' && matchMedia('(max-width: 600px)').matches));
  const [sideTab, setSideTab] = useState('pages');
  const [ribbonOpen, setRibbonOpen] = useState(() => ls.get('office-pdf-ribbon', true));
  const [night, setNight] = useState(false);
  const [current, setCurrent] = useState(null);
  const [selection, setSelection] = useState(() => new Set());
  const [selectedAnnot, setSelectedAnnot] = useState(null);
  const [pageInput, setPageInput] = useState('1');
  const [jump, setJump] = useState(null);
  const [dialog, setDialog] = useState(null);
  const [panel, setPanel] = useState(null); // 'redact'
  const [busy, setBusy] = useState(null);
  const [menu, setMenu] = useState(null);
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
  const viewerApi = useRef(null);
  const perf = useRef({ t0: 0, painted: false });

  const editable = !readOnly && phase === 'ready' && access.canEdit && !access.viewOnly;

  /* ----- proxies (pdf.js pages): batched, on demand ----- */
  const flushProxies = () => {
    proxyFlush.current = 0;
    const buf = proxyBuf.current; if (!buf.size) return;
    proxyBuf.current = new Map();
    setProxies((old) => { const n = new Map(old); for (const [k, v] of buf) n.set(k, v); return n; });
  };
  const putProxy = (key, p) => { proxyBuf.current.set(key, p); if (!proxyFlush.current) proxyFlush.current = requestAnimationFrame(flushProxies); };
  const proxiesRef = useRef(proxies); proxiesRef.current = proxies;
  const getProxy = useCallback((mp) => {
    const key = `${mp.src}:${mp.index}`;
    const have = proxiesRef.current.get(key) || proxyBuf.current.get(key);
    if (have) return Promise.resolve(have);
    if (proxyWait.current.has(key)) return proxyWait.current.get(key);
    const s = sources.current[mp.src];
    if (!s) return Promise.resolve(null);
    const p = s.doc.getPage(mp.index + 1).then((pg) => { putProxy(key, pg); proxyWait.current.delete(key); return pg; }, () => { proxyWait.current.delete(key); return null; });
    proxyWait.current.set(key, p);
    return p;
  }, []);
  const needProxy = useCallback((mp) => { getProxy(mp); }, [getProxy]);
  /** all pages of a source in the background, a few at a time between frames */
  const loadAllProxies = (src) => {
    const s = sources.current[src]; if (!s) return;
    let i = 0;
    const step = () => {
      if (sources.current[src] !== s) return;
      const end = Math.min(s.doc.numPages, i + 24);
      const batch = [];
      for (; i < end; i++) batch.push(getProxy({ src, index: i }));
      Promise.all(batch).then(() => { if (i < s.doc.numPages) setTimeout(step, 0); });
    };
    idle(step);
  };

  /* ----- loading ----- */
  useEffect(() => {
    let gone = false;
    setPhase('loading');
    perf.current = { t0: performance.now(), painted: false };
    (async () => {
      try {
        let src = bytes; let protection = null; let canEdit = true; let viewOnly = false; let protectedFile = false; let pwUsed = null;
        // 1. protection (our handler: the file stays editable and is saved with the same passwords)
        if (isEncryptedBytes(bytes)) {
          let r = null; let wrong = false;
          for (;;) {
            try { r = await openProtected(bytes, pwUsed || ''); break; } catch (e) {
              if (e.code === 'PLAIN') break;
              if (e.code !== 'PASSWORD') { viewOnly = true; break; }
              if (gone) return;
              setPhase('password');
              const pw = await promptDialog({ title: wrong || pwUsed ? t('Wrong password – try again') : t('This PDF is protected'), text: t('Enter the password to open {name}.', { name: file.name }), label: t('Password'), password: true, confirm: t('Open') });
              if (pw == null) { if (!gone) setPhase('locked'); return; }
              wrong = true; pwUsed = pw; setPhase('loading');
            }
          }
          if (r) {
            protectedFile = true; src = r.plain; protection = { keep: r.protection };
            canEdit = r.protection.owner || r.protection.permissions.modify || r.protection.permissions.annotate;
          }
        }
        if (gone) return;
        // 2. our own annotations become editable again
        let own = null;
        if (hasMarker(src)) {
          try { const d = await loadPdf(src); own = readOwnAnnotations(d); if (own.some((l) => l.length)) src = await d.save({ updateFieldAppearances: false }); else own = null; } catch { own = null; }
        }
        // 3. pdf.js (CMaps of the component when it is there)
        const cjk = await cjkAssets({ wait: needsCmaps(src) ? 8000 : 400 });
        const doc = await openDoc(src, {
          cjk, password: viewOnly ? pwUsed || undefined : undefined,
          onPassword: viewOnly ? async (w) => { if (gone) return null; setPhase('password'); const pw = await promptDialog({ title: w ? t('Wrong password – try again') : t('This PDF is protected'), text: t('Enter the password to open {name}.', { name: file.name }), label: t('Password'), password: true, confirm: t('Open') }); if (pw != null) setPhase('loading'); return pw; } : undefined,
        });
        if (gone) { doc.destroy(); return; }
        for (const s of sources.current) { try { s.doc.destroy(); } catch { /* gone */ } }
        sources.current = [{ bytes: src, doc, name: file.name }];
        proxyBuf.current = new Map(); proxyWait.current = new Map();
        const first = await doc.getPage(1);
        const fv = first.getViewport({ scale: 1 });
        const map = new Map([['0:0', first]]);
        const pages = [...Array(doc.numPages).keys()].map((index) => ({ id: `p${index}`, src: 0, index, rot: 0 }));
        const annots = {};
        if (own) own.forEach((l, i) => { if (l.length && pages[i]) annots[pages[i].id] = l; });
        const m = { pages, annots, meta: null, flatten: false, protection, redact: {} };
        modelRef.current = m; setModelState(m); setProxies(map); setDefaultSize([fv.width, fv.height]); setCurrent(pages[0] ? pages[0].id : null);
        hist.current = { past: [], future: [] }; changed.current = false; formsChanged.current = false; widgets.current = null; textIdx.current = new Map();
        doc.annotationStorage.onSetModified = () => { formsChanged.current = true; host.dirty(true); };
        setAccess({ protectedFile, canEdit, viewOnly });
        setNeedFonts(!cjk && needsCmaps(src));
        setPhase('ready');
        loadAllProxies(0);
        // details in the background (pdf.js only – pdf-lib parses the whole file, that waits for idle time)
        const [meta, outline, att, fo, js] = await Promise.all([doc.getMetadata().catch(() => ({ info: {} })), doc.getOutline().catch(() => null), doc.getAttachments().catch(() => null), doc.getFieldObjects().catch(() => null), doc.getJSActions().catch(() => null)]);
        if (gone) return;
        const info = meta.info || {};
        const pdfa = !!(meta.metadata && (() => { try { return meta.metadata.get('pdfaid:part'); } catch { return null; } })());
        const protLabel = protection ? `${t('Yes')} · ${({ rc4: 'RC4', aes128: 'AES-128', aes256: 'AES-256' })[protection.keep.method] || ''}` : viewOnly ? t('Yes') : '';
        setPdfInfo((old) => ({ ...old, info: { ...info, pdfa, protection: protLabel }, meta: { title: info.Title || '', author: info.Author || '', subject: info.Subject || '', keywords: info.Keywords || '', creator: info.Creator || '', producer: info.Producer || '' }, outline: outline || [], attachments: att ? Object.values(att) : [], formInfo: { fields: fo ? Object.keys(fo).length : 0 } }));
        const items = [];
        if (viewOnly && !readOnly) items.push({ level: 'warn', text: t('This PDF uses a protection that can only be read here – editing is off.') });
        else if (protectedFile && !canEdit && !readOnly) items.push({ level: 'info', text: t('This PDF only allows reading and printing. Enter the permissions password (Tools → Protect) to edit it.') });
        if (info.IsSignaturesPresent) items.push({ level: 'warn', text: t('This PDF is digitally signed. Saving changes makes the signature invalid.') });
        if (info.IsXFAPresent) items.push({ level: 'info', text: t('This PDF contains an XFA form – the basic form is shown and filled; the XFA part is kept as it is.') });
        if (js && Object.keys(js).length) items.push({ level: 'info', text: t('Scripts inside this PDF are not run here.') });
        if (host.fidelity) host.fidelity({ items });
        if (info.IsSignaturesPresent || pdfa) {
          idle(async () => {
            try { const lib = await loadPdf(src); if (gone) return; setPdfInfo((old) => ({ ...old, signatures: listSignatures(lib), insp: inspect(lib) })); } catch { /* not for pdf-lib */ }
          });
        }
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
  const onFirstPaint = useCallback(() => {
    if (perf.current.painted) return;
    perf.current.painted = true;
    const ms = Math.round(performance.now() - perf.current.t0);
    try { window.__officePdfOpen = [...(window.__officePdfOpen || []), { ms, pages: sources.current[0] ? sources.current[0].doc.numPages : 0, size: bytes.length }]; } catch { /* ignore */ }
    if (host.perf) host.perf('interactive');
  }, [bytes, host]);

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

  /** fonts for the annotation texts: bundled Latin font, the component's CJK fonts when needed; the rest as pictures */
  const prepareAnnots = async (annots) => {
    const latin = await latinFontBytes();
    const fonts = [latin];
    let miss = missingGlyphs(annots, fonts);
    if (miss.length) {
      try { fonts.push(...await cjkFontBytes(host, t('Some annotation texts use characters ({chars}) that need the font add-on.', { chars: miss.slice(0, 6).join(' ') }))); } catch { /* no add-on */ }
      miss = missingGlyphs(annots, fonts);
    }
    if (!miss.length) return { annots, fonts };
    const set = new Set(miss); const out = {};
    for (const [pid, list] of Object.entries(annots)) {
      out[pid] = await Promise.all(list.map(async (a) => ((a.type === 'text') && [...String(a.text || '')].some((ch) => set.has(ch)) ? { ...a, raster: await rasterText(a) } : a)));
    }
    return { annots: out, fonts };
  };

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
    const prepared = await prepareAnnots(opts.annots || m.annots);
    let out = await buildDocument({ sources: sources.current.map((s) => s.bytes), names: sources.current.map((s, i) => (i ? s.name : '')), pages: opts.pages || m.pages, annots: prepared.annots, fonts: prepared.fonts, forms, flatten: m.flatten, meta: m.meta, mode: opts.mode || annotMode, author: (settings && settings.general && settings.general.author) || '' });
    if (m.protection && !opts.plain) out = await encryptPdf(out, m.protection.keep ? { keep: m.protection.keep } : m.protection);
    return out;
  }, [bytes, annotMode, settings]); // eslint-disable-line react-hooks/exhaustive-deps

  /** heavy tools (OCR, redaction, compression) produce a new document – it replaces the open one (not saved yet) */
  const rebase = async (plain) => {
    let own = null; let src = plain;
    if (hasMarker(plain)) { const d = await loadPdf(plain); own = readOwnAnnotations(d); src = await d.save({ updateFieldAppearances: false }); }
    const doc = await openDoc(src, { cjk: await cjkAssets() });
    for (const s of sources.current) { try { s.doc.destroy(); } catch { /* gone */ } }
    sources.current = [{ bytes: src, doc, name: file.name }];
    proxyBuf.current = new Map(); proxyWait.current = new Map();
    const first = await doc.getPage(1);
    const pages = [...Array(doc.numPages).keys()].map((index) => ({ id: `r${Date.now().toString(36)}p${index}`, src: 0, index, rot: 0 }));
    const annots = {};
    if (own) own.forEach((l, i) => { if (l.length) annots[pages[i].id] = l; });
    const curIdx0 = Math.max(0, modelRef.current.pages.findIndex((p) => p.id === currentRef.current));
    const m = { ...modelRef.current, pages, annots, redact: {}, flatten: false };
    hist.current = { past: [], future: [] };
    modelRef.current = m; setModelState(m); setProxies(new Map([['0:0', first]])); setHistTick((x) => x + 1);
    widgets.current = null; textIdx.current = new Map(); formsChanged.current = false;
    doc.annotationStorage.onSetModified = () => { formsChanged.current = true; host.dirty(true); };
    changed.current = true; host.dirty(true);
    const target = pages[Math.min(curIdx0, pages.length - 1)];
    setCurrent(target.id); setJump({ id: target.id, n: Date.now() });
    loadAllProxies(0);
  };

  /** pages as PNG (scale 2 ≈ 144 dpi), from the edited document */
  const withDoc = async (fn, opts) => {
    const out = await build(opts);
    const same = out === bytes && sources.current[0].bytes === bytes;
    const doc = same ? sources.current[0].doc : await openDoc(m0Plain(out), { cjk: await cjkAssets() });
    try { return await fn(doc); } finally { if (!same) doc.destroy(); }
  };
  // a protected result is opened from its plain build
  const m0Plain = (b) => b;
  const renderPngs = useCallback(async (ids) => withDoc(async (doc) => {
    const m = modelRef.current; const res = [];
    for (const id of ids) {
      const n = m.pages.findIndex((p) => p.id === id);
      const page = await doc.getPage(n + 1);
      const c = document.createElement('canvas');
      await renderPageToCanvas(page, c, { scale: 2 });
      res.push({ n: n + 1, bytes: await canvasBytes(c) });
    }
    return res;
  }, { plain: true }), [build, bytes]); // eslint-disable-line react-hooks/exhaustive-deps
  const text = useCallback(async () => withDoc((doc) => pdfText(doc), { plain: true }), [build, bytes]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (phase !== 'ready') return;
    host.setSaver(async () => build());
    host.setExporter({
      formats: [
        { ext: 'pdf', label: t('PDF document') },
        { ext: 'png', label: t('PNG image (current page)') },
        { ext: 'zip', label: t('PNG images of all pages (zip)') },
        { ext: 'txt', label: t('Plain text') },
      ],
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
  }, [phase, presenting, panel]);

  const currentRef = useRef(current); currentRef.current = current;
  const pages = model ? model.pages : [];
  const curIdx = Math.max(0, pages.findIndex((p) => p.id === current));
  const curPage = pages[curIdx];
  const curProxy = curPage ? proxies.get(`${curPage.src}:${curPage.index}`) : null;
  const firstProxy = pages[0] ? proxies.get(`${pages[0].src}:${pages[0].index}`) : null;
  const fitProxy = curProxy || firstProxy;
  const scale = useMemo(() => {
    if (typeof zoom === 'number') return zoom;
    if (!fitProxy) return 1;
    const vp = fitProxy.getViewport({ scale: CSS_UNITS, rotation: pageRotation(fitProxy, curPage || pages[0], viewRot) });
    const horizontalPad = 40; const verticalPad = 32;
    const fitW = (box.w - horizontalPad) / vp.width; const fitH = (box.h - verticalPad) / vp.height;
    if (zoom === 'page-fit') return Math.max(0.1, Math.min(fitW, fitH));
    if (zoom === 'page-width') return Math.max(0.1, fitW);
    const landscape = vp.width > vp.height;
    return Math.max(0.1, Math.min(landscape ? Math.min(fitW, fitH) : fitW, 1.25));
  }, [zoom, fitProxy, curPage, viewRot, box]); // eslint-disable-line react-hooks/exhaustive-deps
  const roundScale = Math.round(scale * 1000) / 1000;
  const scaleRef = useRef(roundScale); scaleRef.current = roundScale;

  const zoomStep = useCallback((d) => {
    setZoom((z) => {
      const s = typeof z === 'number' ? z : scaleRef.current;
      const steps = [0.25, 0.33, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5, 6, 8];
      if (d > 0) return steps.find((x) => x > s + 0.001) || 8;
      return [...steps].reverse().find((x) => x < s - 0.001) || 0.25;
    });
  }, []);
  const zoomBy = useCallback((f) => setZoom((z) => Math.max(0.1, Math.min(8, Math.round((typeof z === 'number' ? z : scaleRef.current) * f * 1000) / 1000))), []);

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
      let offset = 0;
      const proxy = await getProxy(m.pages[n]);
      if (explicit[1] && explicit[1].name === 'XYZ' && proxy && explicit[3] != null) {
        const vp = proxy.getViewport({ scale: scaleRef.current * CSS_UNITS, rotation: pageRotation(proxy, m.pages[n], viewRot) });
        offset = Math.max(0, vp.convertToViewportPoint(explicit[2] || 0, explicit[3])[1] - 8);
      }
      goPage(n, offset);
    } catch { /* broken destination */ }
  }, [viewRot, getProxy]); // eslint-disable-line react-hooks/exhaustive-deps

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
  const updateAnnot = (id, patch) => { const f = findAnnot(id); if (!f) return; setAnnots(f.pid, (l) => l.map((x) => (x.id === id ? { ...x, ...patch } : x))); };
  const deleteAnnot = (id = selectedAnnot) => {
    const f = id && findAnnot(id);
    if (f) { setAnnots(f.pid, (l) => l.filter((x) => x.id !== f.a.id)); setSelectedAnnot(null); return; }
    const r = id && findRedact(id);
    if (r) { const m = modelRef.current; commit({ ...m, redact: { ...m.redact, [r.pid]: m.redact[r.pid].filter((x) => x.id !== id) } }); setSelectedAnnot(null); }
  };
  const findRedact = (id) => { for (const [pid, list] of Object.entries(modelRef.current.redact || {})) { const a = list.find((x) => x.id === id); if (a) return { pid, a }; } return null; };

  const editText = async (pid, a) => {
    const v = await promptDialog({ title: a.type === 'note' ? t('Note') : t('Text'), label: t('Text'), value: a.text || '', confirm: t('Apply'), multiline: true });
    if (v == null) return;
    if (!v.trim()) { setAnnots(pid, (l) => l.filter((x) => x.id !== a.id)); return; }
    setAnnots(pid, (l) => l.map((x) => (x.id === a.id ? { ...x, text: v } : x)));
  };

  /** place a text box / note / stamp / image at a point of a page */
  const placeAt = async (what, pageId, [x, y], rot) => {
    const below = (h) => { const [ox, oy] = turn(0, -h, rot); return [x + ox, y + oy]; };
    if (what === 'text') {
      const v = await promptDialog({ title: t('Text box'), label: t('Text'), value: '', confirm: t('Add'), multiline: true });
      if (!v || !v.trim()) return;
      const [bx, by] = below(fontSize * 0.95);
      setSelectedAnnot(addAnnot(pageId, { type: 'text', x: bx, y: by, text: v, size: fontSize, color, rot }));
    } else if (what === 'note') {
      const v = await promptDialog({ title: t('Sticky note'), label: t('Note'), value: '', confirm: t('Add'), multiline: true });
      if (!v || !v.trim()) return;
      addAnnot(pageId, { type: 'note', x: x - 10, y: y - 10, text: v, color: '#ffd400' });
    } else if (what === 'stamp') {
      const st = STAMPS.find((s) => s.text === stamp) || STAMPS[0];
      const label = t(st.text);
      const h = 30; const w = approxWidth(label, h * 0.5) * 1.25 + 24;
      const [bx, by] = below(h);
      setSelectedAnnot(addAnnot(pageId, { type: 'stamp', x: bx, y: by, w, h, text: label, color: st.color, rot }));
    } else if ((what === 'image' || what === 'signature') && pending) {
      const maxW = pending.signature ? 160 : 240;
      const s = Math.min(1, maxW / pending.w);
      const w = pending.w * s * (pending.signature ? 1 : 0.75); const h = pending.h * s * (pending.signature ? 1 : 0.75);
      const [bx, by] = below(h);
      setSelectedAnnot(addAnnot(pageId, { type: 'image', x: bx, y: by, w, h, rot, png: pending.png, signature: !!pending.signature }));
      setPending(null); setToolState('select');
    }
  };

  const onPointer = useCallback(async (ev) => {
    const m = modelRef.current;
    if (ev.kind === 'blank') { setSelectedAnnot(null); return; }
    if (ev.kind === 'drawn') {
      const a = draftToAnnot(ev.draft, { color, width });
      if (!a) return;
      const tiny = a.type === 'ink' ? a.paths[0].length < 4 : a.type === 'line' || a.type === 'arrow' ? Math.hypot(a.x2 - a.x1, a.y2 - a.y1) < 3 : a.w < 3 || a.h < 3;
      if (tiny) return;
      if (a.type === 'redact') { addRedact(ev.page, { x: a.x, y: a.y, w: a.w, h: a.h }); return; }
      addAnnot(ev.page, a);
      return;
    }
    if (ev.kind === 'place') { await placeAt(tool, ev.page, ev.pt, ev.rot); return; }
    if (ev.kind === 'redact-pick') { setSelectedAnnot(ev.id); return; }
    if (ev.kind === 'grab' || ev.kind === 'scale' || ev.kind === 'end1' || ev.kind === 'end2') {
      const f = findAnnot(ev.id); if (!f) return;
      setSelectedAnnot(ev.id);
      if (!editable) return;
      const start = ev.pt; const orig = f.a; const pid = f.pid; let moved = false;
      const c = center(orig);
      const d0 = Math.hypot(start[0] - c[0], start[1] - c[1]) || 1;
      let raf = 0; let lastE = null;
      const apply = () => {
        raf = 0;
        const p = ev.toPdf(lastE);
        let next;
        if (ev.kind === 'grab') next = translate(orig, p[0] - start[0], p[1] - start[1]);
        else if (ev.kind === 'end1') next = { ...orig, x1: p[0], y1: p[1] };
        else if (ev.kind === 'end2') next = { ...orig, x2: p[0], y2: p[1] };
        else next = scaleAnnot(orig, Math.max(0.1, Math.hypot(p[0] - c[0], p[1] - c[1]) / d0));
        const mm = modelRef.current;
        const nm = { ...mm, annots: { ...mm.annots, [pid]: mm.annots[pid].map((x) => (x.id === orig.id ? next : x)) } };
        modelRef.current = nm; setModelState(nm);
      };
      const mv = (e2) => { moved = true; lastE = e2; if (!raf) raf = requestAnimationFrame(apply); };
      const end = () => {
        window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', end);
        if (raf) { cancelAnimationFrame(raf); apply(); }
        if (!moved) {
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
  // colour / width / size also change the selected annotation (contextual options)
  const selA = selectedAnnot ? findAnnot(selectedAnnot) : null;
  const setColor = (c) => { setColorState(c); ls.set('office-pdf-color', c); if (selA && editable && selA.a.type !== 'image') updateAnnot(selA.a.id, { color: c }); };
  const setWidth = (w) => { setWidthState(w); if (selA && editable && selA.a.width) updateAnnot(selA.a.id, { width: w }); };
  const setFontSize = (s) => { setFontSizeState(s); if (selA && editable && selA.a.type === 'text') updateAnnot(selA.a.id, { size: s }); };

  /* ----- redaction marks ----- */
  const addRedact = (pageId, r, extra = {}) => { const m = modelRef.current; const id = newId(); commit({ ...m, redact: { ...m.redact, [pageId]: [...((m.redact || {})[pageId] || []), { id, ...r, ...extra }] } }); return id; };
  const pageItems = async (mp) => {
    const key = `${mp.src}:${mp.index}`;
    const proxy = await getProxy(mp);
    const items = (await proxy.getTextContent({ disableNormalization: true })).items;
    if (!textIdx.current.has(key)) textIdx.current.set(key, pageIndex(items));
    return { items, idx: textIdx.current.get(key) };
  };
  const markSensitive = async ({ kinds, terms }) => {
    setBusy('find');
    try {
      const m = modelRef.current; const red = { ...(m.redact || {}) }; let n = 0;
      for (const mp of m.pages) {
        const { items, idx } = await pageItems(mp);
        for (const h of findSensitive(idx.text, { kinds, terms })) {
          for (const part of findInPage(idx, h.value, { matchCase: true })) {
            if (part.start !== h.start) continue;
            for (const p of part.parts) {
              const r = hitRects(items.filter((x) => x.str !== undefined)[p.i], p.a, p.b, 1.2);
              const list = red[mp.id] || [];
              if (list.some((x) => Math.abs(x.x - r.x) < 0.5 && Math.abs(x.y - r.y) < 0.5)) continue;
              red[mp.id] = [...list, { id: newId(), ...r, value: h.value, kind: h.kind }]; n++;
            }
          }
        }
      }
      if (n) commit({ ...m, redact: red });
      return n;
    } finally { setBusy(null); }
  };
  const applyRedaction = async () => {
    const m = modelRef.current;
    const marked = m.pages.map((p, i) => ({ p, i, rects: (m.redact || {})[p.id] || [] })).filter((x) => x.rects.length);
    if (!marked.length) return;
    const ok = await confirmDialog({ title: t('Remove the marked content for good?'), text: t('Text, graphics and image pixels under {n} marks on {p} pages are removed. Pages without pictures keep their text and graphics; pages with pictures become pictures with an invisible text layer of the remaining text. You can still undo it until you save.', { n: marked.reduce((s, x) => s + x.rects.length, 0), p: marked.length }), confirm: t('Redact'), danger: true });
    if (!ok) return;
    setBusy('apply');
    try {
      const plain = await build({ force: true, plain: true, mode: 'burn' });
      const doc = await openDoc(plain, { cjk: await cjkAssets() });
      const secrets = [...new Set(marked.flatMap((x) => x.rects.map((r) => r.value).filter(Boolean)))];
      const jobs = [];
      try {
        for (const { i, rects } of marked) {
          const page = await doc.getPage(i + 1);
          const s = 200 / 72;
          const vp = page.getViewport({ scale: s, rotation: 0 });
          const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
          const ctx = c.getContext('2d', { alpha: false });
          await page.render({ canvasContext: ctx, viewport: vp, background: '#ffffff', annotationMode: pdfjsLib().AnnotationMode.ENABLE }).promise;
          ctx.fillStyle = '#000';
          for (const r of rects) { const [x0, y0, x1, y1] = vp.convertToViewportRectangle([r.x, r.y, r.x + r.w, r.y + r.h]); ctx.fillRect(Math.floor(Math.min(x0, x1)) - 1, Math.floor(Math.min(y0, y1)) - 1, Math.ceil(Math.abs(x1 - x0)) + 2, Math.ceil(Math.abs(y1 - y0)) + 2); }
          const items = (await page.getTextContent({ disableNormalization: true })).items.filter((x) => x.str !== undefined);
          const idx = pageIndex(items);
          const removed = new Map();
          for (const r of rects) {
            if (!r.value) continue;
            for (const h of findInPage(idx, r.value, { matchCase: true })) for (const p of h.parts) { const hr = hitRects(items[p.i], p.a, p.b, 0); if (hr.x < r.x + r.w && hr.x + hr.w > r.x && hr.y < r.y + r.h && hr.y + hr.h > r.y) { if (!removed.has(p.i)) removed.set(p.i, []); removed.get(p.i).push([p.a, p.b]); } }
          }
          jobs.push({ index: i, rects, view: page.view, image: { bytes: await canvasBytes(c, 'image/jpeg', 0.9) }, words: keptWords(items, rects, removed) });
          c.width = 0; c.height = 0;
        }
      } finally { doc.destroy(); }
      let out = await applyRedactions(plain, jobs);
      // check: extract the text of the result again
      const check = async (b) => {
        const d = await openDoc(b); const ex = [];
        try { for (let k = 1; k <= d.numPages; k++) { const items = (await (await d.getPage(k)).getTextContent()).items.filter((x) => x.str !== undefined); ex.push({ items, text: items.map((x) => x.str).join(' ') }); } } finally { d.destroy(); }
        const marks = {}; for (const j of jobs) marks[j.index] = j.rects;
        const v = verifyRedaction(ex, { secrets, marks });
        const raw = await scanFile(b, secrets);
        return { ok: v.ok && !raw.length, leaks: v.leaks.length + raw.length };
      };
      let res = await check(out);
      if (!res.ok) { out = await applyRedactions(plain, jobs.map((j) => ({ ...j, words: [], vector: false }))); res = { ...(await check(out)), retried: true }; }
      await rebase(out);
      setPanel(null); setTool('text-select');
      setDialog({ type: 'redacted', result: { ok: res.ok && !res.retried, n: jobs.reduce((s, j) => s + j.rects.length, 0), p: jobs.length, leaks: res.leaks } });
    } catch (e) { toast.error(e.message || String(e)); } finally { setBusy(null); }
  };

  /* ----- OCR ----- */
  const runOcr = async ({ scope, lang, onProgress, stopped }) => {
    onProgress({ loading: true, done: 0, total: 1, page: 0 });
    const comp = await ensureComponent(host, 'ocr', t('Text recognition needs the OCR engine (≈ 23 MB download, German and English).'));
    if (!comp) throw new Error(t('The OCR engine is not installed.'));
    const m = modelRef.current;
    const plain = await build({ force: true, plain: true });
    const doc = await openDoc(plain, { cjk: await cjkAssets() });
    // optional language packs (French, Spanish, Italian, Dutch, Polish, Turkish) are components of their own
    const langBases = {};
    for (const l of lang.split('+').filter((x) => OCR_PACKS.includes(x))) {
      const pack = await ensureComponent(host, `ocr-${l}`, t('Text recognition in this language needs its language pack (one-time download).'));
      if (!pack) throw new Error(t('The language pack is not installed.'));
      langBases[l] = pack.base;
    }
    const engine = await createTesseract({ base: comp.base, langs: lang, langBases });
    let words = 0; let done = 0; const layers = {};
    try {
      const all = m.pages.map((p, i) => i);
      let target = scope === 'current' ? (selection.size > 1 ? m.pages.map((p, i) => (selection.has(p.id) ? i : -1)).filter((i) => i >= 0) : [curIdx]) : all;
      if (scope === 'auto') {
        const left = [];
        for (const i of target) { const tc = await (await doc.getPage(i + 1)).getTextContent(); if (tc.items.filter((x) => x.str && x.str.trim()).length < 3) left.push(i); }
        target = left;
      }
      if (!target.length) return { words: 0, pages: 0 };
      for (const i of target) {
        if (stopped()) break;
        onProgress({ done, total: target.length, page: i + 1 });
        const page = await doc.getPage(i + 1);
        const vp = page.getViewport({ scale: 300 / 72, rotation: 0 });
        const k = Math.min(1, Math.sqrt(36e6 / (vp.width * vp.height)));
        const v2 = page.getViewport({ scale: (300 / 72) * k, rotation: 0 });
        const c = document.createElement('canvas'); c.width = Math.round(v2.width); c.height = Math.round(v2.height);
        await page.render({ canvasContext: c.getContext('2d', { willReadFrequently: true }), viewport: v2, background: '#ffffff' }).promise;
        const boxes = await engine.recognize(c);
        c.width = 0; c.height = 0;
        layers[i] = imageWordsToPdf(boxes, (x, y) => v2.convertToPdfPoint(x, y));
        words += layers[i].length; done++;
      }
    } finally { engine.close(); doc.destroy(); }
    onProgress({ done, total: done, page: 0 });
    if (words) {
      const lib = await loadPdf(plain);
      const pgs = lib.getPages();
      for (const [i, w] of Object.entries(layers)) if (w.length) await addTextLayer(lib, pgs[Number(i)], w);
      await rebase(await lib.save({ useObjectStreams: true, updateFieldAppearances: false }));
    }
    return { words, pages: done };
  };

  /* ----- compress / protect ----- */
  const runCompress = async (lv) => {
    const plain = await build({ force: true, plain: true });
    return compressPdf(plain, { maxPx: lv.maxPx, recode: (input, o) => recodeImage(input, { ...o, quality: lv.quality }) });
  };
  const setProtection = (spec) => commit({ ...modelRef.current, protection: spec });
  const unlockOwner = async () => {
    const pw = await promptDialog({ title: t('Permissions password'), text: t('Enter the permissions password to edit this PDF.'), label: t('Password'), password: true, confirm: t('Unlock') });
    if (pw == null) return;
    try {
      const r = await openProtected(bytes, pw);
      if (!r.protection.owner) { toast.error(t('That is not the permissions password.')); return; }
      modelRef.current = { ...modelRef.current, protection: { keep: r.protection } }; setModelState(modelRef.current);
      setAccess((a) => ({ ...a, canEdit: true }));
      toast.success(t('Editing is unlocked.'));
    } catch { toast.error(t('That is not the permissions password.')); }
  };
  const saveDecryptedCopy = async () => {
    const plain = await build({ force: true, plain: true });
    if (host.saveCopy) await host.saveCopy(plain, 'pdf', `${baseName(file.name)} (${t('unprotected')}).pdf`);
    else host.download(plain, `${baseName(file.name)} (${t('unprotected')}).pdf`, PDF_MIME);
  };
  const compare = async () => {
    const [f] = await pickFiles('application/pdf,.pdf', false);
    if (!f) return;
    if (host.compare) { try { await host.compare({ bytes: f.bytes, label: f.name }); return; } catch { /* the editor's own dialog */ } }
    setDialog({ type: 'compare', other: f });
  };

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
  const addSource = async (pdfBytes, at, name = '') => {
    let b = pdfBytes;
    if (isEncryptedBytes(b)) {
      const pw = await promptDialog({ title: t('This PDF is protected'), text: t('Enter the password to open {name}.', { name }), label: t('Password'), password: true, confirm: t('Open') });
      if (pw == null) return [];
      b = (await openProtected(b, pw)).plain;
    }
    const doc = await openDoc(b, { cjk: await cjkAssets() });
    const src = sources.current.length;
    sources.current.push({ bytes: b, doc, name: baseName(name) });
    const m = modelRef.current;
    const added = [...Array(doc.numPages).keys()].map((index) => ({ id: `s${src}p${index}`, src, index, rot: 0 }));
    const pos = at == null ? m.pages.length : at;
    commit({ ...m, pages: [...m.pages.slice(0, pos), ...added, ...m.pages.slice(pos)] });
    await getProxy(added[0]);
    loadAllProxies(src);
    return added;
  };
  const insertBlank = async () => {
    const vp = curProxy ? curProxy.getViewport({ scale: 1 }) : null;
    const size = vp ? [vp.width, vp.height] : SIZES.A4;
    const added = await addSource(await createBlank({ size }), curIdx + 1, '');
    setCurrent(added[0].id); setJump({ id: added[0].id, n: Date.now() });
  };
  const filesToPdf = async (files) => {
    const out = [];
    for (const f of files) {
      if (isPdf(f.bytes)) out.push({ bytes: f.bytes, name: f.name });
      else if (/^image\//.test(f.type) || /\.(png|jpe?g|gif|webp|bmp)$/i.test(f.name)) {
        const img = isJpgOrPng(f.bytes) ? f.bytes : (await toPng(f.bytes, f.type)).png;
        out.push({ bytes: await imagesToPdf([{ bytes: img }]), name: '' });
      } else toast(t('{name} is neither a PDF nor an image.', { name: f.name }), 'warn');
    }
    return out;
  };
  const insertFiles = async (files, atEnd = false) => {
    let at = atEnd ? null : curIdx + 1; let first = null;
    for (const b of await filesToPdf(files)) {
      try {
        const added = await addSource(b.bytes, at, b.name);
        if (!added.length) continue;
        if (!first) first = added[0];
        if (at != null) at += added.length;
      } catch (e) { toast.error(e.code === 'PASSWORD' ? t('Wrong password') : e.message); }
    }
    if (first) { setCurrent(first.id); setJump({ id: first.id, n: Date.now() }); toast.success(t('Pages inserted')); }
  };
  const insertFile = async (atEnd = false) => {
    const files = await pickFiles('application/pdf,.pdf,image/*', true);
    if (files.length) await insertFiles(files, atEnd);
  };
  // files dropped from Dateien or the desktop
  useEffect(() => {
    if (!editable || !host.insertFiles) return undefined;
    return host.insertFiles(async (list) => {
      const files = [];
      for (const f of list) { if (/\.(pdf|png|jpe?g|gif|webp|bmp)$/i.test(f.name)) files.push({ name: f.name, type: f.mime || '', bytes: await f.bytes() }); }
      if (files.length) await insertFiles(files, false); else toast(t('Only PDFs and pictures can be inserted.'), 'info');
    });
  }, [editable]); // eslint-disable-line react-hooks/exhaustive-deps
  const extract = async () => {
    const ids = targets();
    const keep = pages.filter((p) => ids.includes(p.id));
    const out = await build({ pages: keep, force: true });
    host.download(out, `${baseName(file.name)} (${t('pages')} ${keep.map((p) => pages.indexOf(p) + 1).join(', ')}).pdf`, PDF_MIME);
  };
  const split = async () => {
    const v = await promptDialog({ title: t('Split PDF'), text: t('Every part gets this many pages. You get a zip with all parts.'), label: t('Pages per part'), value: '1', confirm: t('Split'), validate: (x) => (/^\d+$/.test(x) && Number(x) >= 1 ? null : t('Enter a number of 1 or more')) });
    if (v == null) return;
    const all = await build({ force: true, plain: true });
    const parts = await splitPdf(all, Number(v));
    const w = String(parts.length).length;
    host.download(zipStore(parts.map((b, i) => ({ name: `${baseName(file.name)}-${String(i + 1).padStart(w, '0')}.pdf`, bytes: b }))), `${baseName(file.name)}-split.zip`, 'application/zip');
  };

  /* ----- misc actions ----- */
  async function doPrint() {
    const out = await build({ plain: true });
    host.print(new Blob([out], { type: PDF_MIME }));
  }
  const toWord = async () => {
    const ok = await confirmDialog({ title: t('Convert to Word?'), text: t('The Office engine reads a PDF like a drawing: every line becomes its own text box, fonts are approximated. Good for copying text and small changes – for longer editing, prefer the original document.'), confirm: t('Convert') });
    if (!ok) return;
    const out = await host.convert(await build({ plain: true }), 'pdf', 'docx');
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
      if (!textIdx.current.has(key)) { const proxy = await getProxy(mp); if (!proxy) continue; textIdx.current.set(key, pageIndex((await proxy.getTextContent({ disableNormalization: true })).items)); }
      for (const h of findInPage(textIdx.current.get(key), q, { matchCase, wholeWord })) out.push({ page: mp.id, ...h });
    }
    setHits(out);
    const startAt = out.findIndex((h) => m.pages.findIndex((p) => p.id === h.page) >= curIdx);
    setActiveHit(out.length ? Math.max(0, startAt) : -1);
  }, [getProxy, matchCase, wholeWord, curIdx]);
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
  const openSearch = (q) => { setSearchOpen(true); if (typeof q === 'string') setQuery(q); setTimeout(() => searchInput.current && searchInput.current.select(), 0); };

  /* ----- right-click menu ----- */
  const onContextMenu = useCallback(({ e, page, pt, aid }) => {
    e.preventDefault();
    const sel = window.getSelection();
    const selText = sel && !sel.isCollapsed ? sel.toString().trim() : '';
    const markSel = (type) => {
      const byPage = viewerApi.current ? viewerApi.current.selectionQuads(sel) : new Map();
      for (const [id, qs] of byPage) {
        if (type === 'redact') for (const q of qs) addRedact(id, { x: q.x - 1, y: q.y - 1, w: q.w + 2, h: q.h + 2 }, { value: selText });
        else addAnnot(id, { type, quads: qs, color: type === 'highlight' ? '#ffd400' : color, text: selText });
      }
      sel.removeAllRanges();
    };
    const mp = pages.find((p) => p.id === page);
    const rot = mp ? pageRotation(proxies.get(`${mp.src}:${mp.index}`), mp, viewRot) : 0;
    const a = aid ? findAnnot(aid) : null;
    const items = [];
    if (selText) {
      items.push({ icon: Copy, label: t('Copy'), keys: 'Ctrl+C', run: () => { try { navigator.clipboard.writeText(selText); } catch { document.execCommand('copy'); } } });
      if (editable) items.push({ icon: Highlighter, label: t('Highlight'), run: () => markSel('highlight') }, { icon: Underline, label: t('Underline'), run: () => markSel('underline') }, { icon: Strikethrough, label: t('Strike through'), run: () => markSel('strike') }, { icon: EyeOff, label: t('Mark for redaction'), run: () => { markSel('redact'); setPanel('redact'); } });
      items.push({ icon: Search, label: t('Search for “{text}”', { text: selText.length > 24 ? `${selText.slice(0, 24)}…` : selText }), keys: 'Ctrl+F', run: () => openSearch(selText) }, '-');
    }
    if (a && editable) {
      if (a.a.type === 'text' || a.a.type === 'note') items.push({ icon: Pencil, label: t('Edit text …'), run: () => editText(a.pid, a.a) });
      items.push({ icon: Trash2, label: t('Delete annotation'), keys: 'Del', danger: true, run: () => deleteAnnot(a.a.id) }, '-');
    }
    if (editable && pt) {
      items.push({ icon: Type, label: t('Add text here'), keys: 'T', run: () => placeAt('text', page, pt, rot) }, { icon: StickyNote, label: t('Add a note here'), run: () => placeAt('note', page, pt, rot) }, { icon: Stamp, label: t('Stamp “{name}” here', { name: t(stamp) }), run: () => placeAt('stamp', page, pt, rot) }, '-');
      items.push({ icon: RotateCcw, label: t('Rotate page left'), run: () => rotatePages([page], -90) }, { icon: RotateCw, label: t('Rotate page right'), run: () => rotatePages([page], 90) }, { icon: Trash2, label: t('Delete page'), danger: true, disabled: pages.length <= 1, run: () => deletePages([page]) }, '-');
    }
    items.push({ icon: ZoomIn, label: t('Zoom in'), keys: 'Ctrl++', run: () => zoomStep(1) }, { icon: ZoomOut, label: t('Zoom out'), keys: 'Ctrl+-', run: () => zoomStep(-1) }, { icon: Maximize, label: t('Whole page'), keys: 'Ctrl+0', run: () => setZoom('page-fit') }, '-', { icon: Printer, label: t('Print …'), keys: 'Ctrl+P', run: () => doPrint() }, { icon: Info, label: t('Document properties'), run: () => setDialog('props') });
    setMenu({ x: e.clientX, y: e.clientY, items });
  }, [pages, proxies, viewRot, editable, color, stamp]); // eslint-disable-line react-hooks/exhaustive-deps

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
    if (editable && selectedAnnot && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(k)) {
      const f = findAnnot(selectedAnnot);
      if (f) { e.preventDefault(); const st = e.shiftKey ? 10 : 1; const dx = k === 'ArrowLeft' ? -st : k === 'ArrowRight' ? st : 0; const dy = k === 'ArrowDown' ? -st : k === 'ArrowUp' ? st : 0; setAnnots(f.pid, (l) => l.map((x) => (x.id === f.a.id ? translate(x, dx, dy) : x))); return; }
    }
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
  const redactCount = Object.values(model.redact || {}).reduce((s, l) => s + l.length, 0);
  const st = {
    tab, editable, tool, color, width, fontSize, stamp, annotMode, zoom, scroll, sidebar, night, searchOpen, ribbonOpen,
    pageNo: curIdx + 1, pageCount: pages.length, pageInput, canUndo: h.past.length > 0, canRedo: h.future.length > 0,
    selectedAnnot, selectedType: selA ? selA.a.type : null, hasConverter: !!host.hasConverter, selCount: selection.size, formInfo: pdfInfo.formInfo, flatten: model.flatten,
    protectedNow: !!model.protection, redactOpen: panel === 'redact', redactCount, readOnly: !!readOnly, locked: access.protectedFile && !access.canEdit && !readOnly,
  };
  const act = {
    setTab, setTool, setColor, setWidth, setFontSize, setStamp,
    setAnnotMode: (v) => { setAnnotModeState(v); ls.set('office-pdf-annot-mode', v); if (changed.current) host.dirty(true); },
    setZoom: (z) => setZoom(z), zoomStep, setScroll: (v) => { setScroll(v); setJump({ id: current, n: Date.now() }); },
    toggleSidebar: () => setSidebar((x) => !x), toggleNight: () => setNight((x) => !x), toggleSearch: () => (searchOpen ? setSearchOpen(false) : openSearch()),
    toggleRibbon: () => setRibbonOpen((x) => { ls.set('office-pdf-ribbon', !x); return !x; }),
    prev: () => goPage(curIdx - 1), next: () => goPage(curIdx + 1), setPageInput,
    goInput: () => { const n = parseInt(pageInput, 10); if (n >= 1 && n <= pages.length) goPage(n - 1); else setPageInput(String(curIdx + 1)); },
    rotateView: () => setViewRot((r) => (r + 90) % 360), present: () => setPresenting(true), print: doPrint,
    properties: () => setDialog('props'), toWord, undo, redo, deleteAnnot: () => deleteAnnot(), pickImage,
    insertBlank, insertFile: () => insertFile(false), mergeFiles: () => insertFile(true), rotatePages: (d) => rotatePages(targets(), d),
    movePages, deletePages: () => deletePages(targets()), extract, split,
    signature: () => setDialog('sign'), setFlatten: (v) => commit({ ...modelRef.current, flatten: v }),
    ocr: () => setDialog('ocr'), redact: () => { setPanel((p) => (p === 'redact' ? null : 'redact')); setTool(panel === 'redact' ? 'text-select' : 'redact'); },
    compare, compress: () => setDialog('compress'), protect: () => setDialog('protect'), unlock: unlockOwner,
  };
  const meta = { ...pdfInfo.meta, ...(model.meta || {}) };
  const pageSize = curProxy ? (() => { const v = curProxy.getViewport({ scale: 1 }); return `${Math.round(v.width / 72 * 25.4)} × ${Math.round(v.height / 72 * 25.4)} mm`; })() : '';
  const curSigned = pdfInfo.info.IsSignaturesPresent;
  const pageNoOf = (pid) => pages.findIndex((p) => p.id === pid) + 1;

  return (
    <div className={`of-pdf-root${night ? ' night' : ''}`} ref={root} tabIndex={-1} onKeyDown={onKey}>
      <Toolbar st={st} act={act} />
      {searchOpen && (
        <div className="of-pdf-search" role="search">
          <Search size={15} className="muted" />
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
      {needFonts && <div className="of-pdf-callout"><Callout kind="info" icon={Languages}><span className="row">{t('This PDF uses Asian fonts that are not inside the file. With the font add-on its text shows correctly.')} <Btn className="sm" onClick={async () => { const c = await ensureComponent(host, 'cjk-fonts', t('Fonts and character maps for Chinese, Japanese and Korean PDFs.')); if (c) { forgetCjk(); setRetry((x) => x + 1); } }}>Install fonts</Btn></span></Callout></div>}
      {st.locked && <div className="of-pdf-callout"><Callout kind="info" icon={KeyRound}><span className="row">{t('This PDF only allows reading and printing.')} <Btn className="sm" onClick={unlockOwner}>Enter permissions password …</Btn></span></Callout></div>}
      <div className="of-pdf-body">
        {sidebar && (
          <Sidebar tab={sideTab} onTab={setSideTab} pages={pages} proxies={proxies} needProxy={needProxy} defaultSize={defaultSize} current={current} selection={selection}
            onSelect={(s, id) => { setSelection(s); const n = pages.findIndex((p) => p.id === id); goPage(n); }}
            editable={editable} onMove={movePagesTo} onRotate={rotatePages} onDelete={deletePages}
            outline={pdfInfo.outline} onOutline={(it) => { if (it.dest) goDest(0, it.dest); else if (it.url) openLink(it.url); }}
            attachments={pdfInfo.attachments} onAttachment={(a) => host.download(a.content, a.filename, 'application/octet-stream')} signatures={pdfInfo.signatures} />
        )}
        <Viewer pages={pages} proxies={proxies} docs={sources.current.map((s) => s.doc)} scale={roundScale} viewRot={viewRot} scrollMode={scroll}
          current={current} onCurrent={onCurrent} annots={model.annots} redactions={model.redact} selectedId={selectedAnnot} tool={tool} editable={editable} renderForms={!readOnly && !access.viewOnly}
          hitsByPage={hitsByPage} activeHit={active} linkService={linkService} onPointer={onPointer} onTextSelection={onTextSelection}
          scrollRef={scrollRef} night={night} onZoom={zoomBy} jump={jump} needProxy={needProxy} defaultSize={defaultSize} onContextMenu={onContextMenu} apiRef={viewerApi} onFirstPaint={onFirstPaint} />
        {panel === 'redact' && editable && (
          <RedactPanel marks={model.redact || {}} pageNo={pageNoOf} busy={busy} onFind={markSensitive} onApply={applyRedaction}
            onRemove={(id) => deleteAnnot(id)} onClear={() => commit({ ...modelRef.current, redact: {} })}
            onGo={(mk) => { setSelectedAnnot(mk.id); goPage(pageNoOf(mk.pid) - 1); }} />
        )}
      </div>
      <div className="of-pdf-status small muted">
        <span>{t('Page {n} of {count}', { n: curIdx + 1, count: pages.length })}</span>
        <span>{Math.round(roundScale * 100)} %</span>
        {viewRot ? <span>{viewRot}°</span> : null}
        {readOnly ? <span>{t('Read only')}</span> : access.viewOnly ? <span><Lock size={12} /> {t('Protected – editing is off')}</span> : model.protection ? <span className="of-pdf-badge"><Lock size={12} /> {t('Protected')}</span> : null}
        {file.converted && <span>{t('Converted from {ext}', { ext: String(file.converted.from).toUpperCase() })}</span>}
        <span className="grow" />
        {MARKUP.includes(tool) && <span>{t('Select text to mark it.')}</span>}
        {tool === 'redact' && <span>{t('Drag over what should disappear · right-click selected text to mark it')}</span>}
        {tool === 'select' && <span>{t('Drag to move · corners resize · arrow keys nudge · double-click edits text · Del deletes')}</span>}
      </div>
      {dialog === 'props' && <Properties file={file} info={pdfInfo.info} meta={meta} editable={editable} pageSize={pageSize} count={pages.length} onClose={() => setDialog(null)} onSave={(m) => commit({ ...modelRef.current, meta: m })} />}
      {dialog === 'sign' && <SignatureDialog host={host} defaultName={ps.signatureName || (settings && settings.general && settings.general.author) || ''} onClose={() => setDialog(null)} onPick={(sig) => { setPending({ ...sig, signature: true }); setToolState('signature'); }} />}
      {dialog === 'protect' && <ProtectDialog protection={model.protection} onApply={setProtection} onClose={() => setDialog(null)} onSaveDecrypted={model.protection ? saveDecryptedCopy : null} />}
      {dialog === 'compress' && <CompressDialog run={runCompress} hasConverter={!!host.hasConverter} onApply={(b) => rebase(b)} onClose={() => setDialog(null)} />}
      {dialog === 'ocr' && <OcrDialog count={pages.length} current={curIdx + 1} selection={selection.size} run={runOcr} onClose={() => setDialog(null)} />}
      {dialog && dialog.type === 'redacted' && <RedactResult result={dialog.result} onClose={() => setDialog(null)} />}
      {dialog && dialog.type === 'compare' && <CompareDialog name={file.name} other={dialog.other} getBytes={() => build({ plain: true })} pick={async () => (await pickFiles('application/pdf,.pdf', false))[0]} onClose={() => setDialog(null)} />}
      {presenting && <Presenter pages={pages} getProxy={getProxy} start={curIdx} onClose={() => setPresenting(false)} />}
      {menu && <ContextMenu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} />}
    </div>
  );
}

/* ---------------- file browser preview + thumbnail ---------------- */
function Preview({ bytes }) {
  const c = useRef(null);
  const [state, setState] = useState('loading');
  useEffect(() => {
    let doc = null; let gone = false;
    (async () => {
      if (isEncryptedBytes(bytes)) { try { bytes = (await openProtected(bytes, '')).plain; } catch (e) { if (e.code === 'PASSWORD') throw Object.assign(new Error('locked'), { name: 'PasswordException' }); } } // eslint-disable-line no-param-reassign
      const d = await openDoc(bytes);
      doc = d; if (gone) return;
      const page = await d.getPage(1);
      const base = page.getViewport({ scale: 1 });
      await renderPageToCanvas(page, c.current, { scale: Math.min(2, 480 / base.width) });
      if (!gone) setState('ok');
    })().catch((e) => { if (!gone) setState(e && e.name === 'PasswordException' ? 'locked' : 'err'); });
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

async function thumbnail({ bytes }) {
  let b = bytes;
  if (isEncryptedBytes(b)) { try { b = (await openProtected(b, '')).plain; } catch { return null; } }
  const doc = await openDoc(b);
  try {
    const page = await doc.getPage(1);
    const base = page.getViewport({ scale: 1 });
    const c = document.createElement('canvas');
    await renderPageToCanvas(page, c, { scale: 320 / base.width });
    return c;
  } finally { doc.destroy(); }
}

export default {
  kind: 'pdf',
  Editor,
  Preview,
  Compare: CompareView,
  thumbnail,
  create: async ({ template, settings } = {}) => {
    if (template && template !== 'blank') throw new Error(t('Unknown template: {id}', { id: template }));
    const size = settings && settings.docs && SIZES[settings.docs.pageSize] ? settings.docs.pageSize : 'A4';
    return createBlank({ size, landscape: !!(settings && settings.docs && settings.docs.orientation === 'landscape') });
  },
  toText: async (bytes) => {
    let b = bytes;
    if (isEncryptedBytes(b)) { try { b = (await openProtected(b, '')).plain; } catch { return ''; } }
    const doc = await openDoc(b);
    try { return await pdfText(doc); } finally { doc.destroy(); }
  },
};
