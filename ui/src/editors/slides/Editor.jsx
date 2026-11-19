/*
 * Presentation editor: state + history, panels, keyboard, clipboard, saver/exporter, presentation mode.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, createElement as h } from 'react';
import { createPortal } from 'react-dom';
import { Loading, Empty, Callout, Btn, confirmDialog, promptDialog } from '@hostwerk/ui';
import { toast } from '../../lib/toast';
import { contextMenu } from '@hostwerk/kit';
import { t } from '@hostwerk/i18n';
import { Presentation, Plus, Copy, Trash2, EyeOff, Scissors, ClipboardPaste, BringToFront, SendToBack, Group, Ungroup, Type, PanelRight, FolderPlus, Pencil, LayoutTemplate, CopyPlus, Link2, Sparkles, StickyNote, FileWarning, Download, Minus } from 'lucide-react';
import Ribbon, { SIZES_PT } from './Ribbon.jsx';
import Canvas from './Canvas.jsx';
import Show, { PresenterRemote } from './Present.jsx';
import { Thumbs, Sorter, Outline, NotesPage, FormatPane, CommentsPane, AnimPane } from './Panels.jsx';
import { FooterDialog, SizeDialog, ChartDataDialog } from './dialogs.jsx';
import SlideView from './render/SlideView.jsx';
import { readAny, writeAny, saveFormat, OOXML, ODF } from './io/index.js';
import { nodeToCanvas, canvasBytes, imagePdf, printHtml } from './io/export.js';
import { FIDELITY } from './io/util.js';
import { tableOp } from './table.js';
import { S, EMU_PER_PT, cmToEmu } from './units.js';
import { runProps, paraProps } from './render/text.js';
import {
  history, histPush, histUndo, histRedo, histReplace, newDoc, addSlide, insertSlides, duplicateSlides, deleteSlides, moveSlides, updateSlide, updateSlides,
  updateEls, addEls, deleteEls, zorder, groupEls, ungroupEls, alignEls, distributeEls, setSize, applyTheme, setThemeColors, setThemeFonts, resetSlide,
  addSection, renameSection, removeSection, sectionGroups, findEl, walk, bbox, moveEl, reId, uid, textBox, shape, line as lineEl, image as imageEl,
  media as mediaEl, table as tableEl, chart as chartEl, body, para, setRunProps, mapParas, mapRuns, layoutByKind, slideTitle, newSlide, deepClone,
} from './model.js';
import { SHAPES } from './render/geometry.js';
import { keyLabel } from '../_shared/keys.js';

/* ---------------------------------------------------------------- module clipboard (shared by all open decks) */
let CLIP = null;

const RAW_KEYS = ['spid', 'raw', 'rawRels', 'rawSig', 'rawXfrm', 'rawNs', 'rawSrc', 'rawDoc', 'rawFmt', 'rawStyles', 'rawParts', 'rawBox'];
function stripRaw(e) {
  const n = { ...e };
  if (!e.preserved) for (const k of RAW_KEYS) delete n[k]; else delete n.spid;
  if (n.children) n.children = n.children.map(stripRaw);
  return n;
}
/** keep group boxes in sync with their children */
function fitGroups(els) {
  return els.map((e) => {
    if (e.type !== 'group') return e;
    const children = fitGroups(e.children);
    const b = bbox(children.map((c) => ({ ...c, rot: 0 })));
    const changed = children.some((c, i) => c !== e.children[i]);
    if (!changed && Math.abs(b.x - e.x) < 2 && Math.abs(b.y - e.y) < 2 && Math.abs(b.w - e.w) < 2 && Math.abs(b.h - e.h) < 2) return e;
    return { ...e, children, x: Math.round(b.x), y: Math.round(b.y), w: Math.round(b.w), h: Math.round(b.h) };
  });
}
const readFile = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(f); });
const imgSize = (src) => new Promise((res) => { const i = new Image(); i.onload = () => res({ w: i.naturalWidth || 400, h: i.naturalHeight || 300 }); i.onerror = () => res({ w: 400, h: 300 }); i.src = src; });
function pickFile(accept, multiple = false) {
  return new Promise((res) => {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = accept; inp.multiple = multiple;
    inp.onchange = () => res([...(inp.files || [])]);
    inp.click();
  });
}
const caseOf = (s, mode) => (mode === 'upper' ? s.toUpperCase() : mode === 'lower' ? s.toLowerCase() : mode === 'title' ? s.toLowerCase().replace(/(^|\s)(\p{L})/gu, (m, a, b) => a + b.toUpperCase()) : s.toLowerCase().replace(/(^\s*\p{L})|([.!?]\s+\p{L})/gu, (m) => m.toUpperCase()));

/* ---------------------------------------------------------------- offscreen renderer (export) */
function useOffscreen() {
  const [job, setJob] = useState(null);
  const ref = useRef(null);
  const queue = useRef(Promise.resolve());
  const render = useCallback((doc, slide, index) => {
    const p = queue.current.then(() => new Promise((resolve) => setJob({ doc, slide, index, resolve })));
    queue.current = p.catch(() => {});
    return p;
  }, []);
  useLayoutEffect(() => {
    if (!job || !ref.current) return;
    // wait a frame so images decode and autofit settles
    const imgs = [...ref.current.querySelectorAll('img')];
    Promise.all(imgs.map((i) => (i.complete ? null : new Promise((r) => { i.onload = r; i.onerror = r; })))).then(() => setTimeout(() => { const node = ref.current && ref.current.firstChild; job.resolve(node); setJob(null); }, 40));
  }, [job]);
  const host = typeof document !== 'undefined' ? document.body : null;
  const el = job && host ? createPortal(<div className="of-slides-offscreen" ref={ref} aria-hidden="true"><SlideView doc={job.doc} slide={job.slide} index={job.index} scale={1} mode="show" /></div>, host) : null;
  return [el, render];
}

/* ---------------------------------------------------------------- entry: presenter remote or editor */
export default function Editor(props) {
  const token = useMemo(() => { const m = /[?&]present=([a-z0-9]+)/i.exec(typeof location !== 'undefined' ? location.hash : ''); return m ? m[1] : null; }, []);
  if (token) return <div className="of-slides-root"><PresenterRemote token={token} /></div>;
  return <SlidesEditor {...props} />;
}

function SlidesEditor({ file, bytes, readOnly: ro0, settings = {}, host = {} }) {
  const ss = settings.slides || {};
  const gen = settings.general || {};
  const units = gen.units === 'in' ? 'in' : 'cm';
  const srcExt = (file.converted && file.converted.to) || file.ext;
  const [phase, setPhase] = useState('loading');
  const [error, setError] = useState(null);
  const [hist, setHist] = useState(null);
  const histRef = useRef(null); histRef.current = hist;
  const doc = hist && hist.doc;
  const original = useRef({ bytes: null, fmt: null });
  const savedDoc = useRef(null);
  const explicitSave = useRef(false); // Ctrl+S in the editor: ask about "Save as a copy" again even after "Cancel"
  const [fidelity, setFidelity] = useState([]);
  const [curId, setCurId] = useState(null);
  const [selSlides, setSelSlides] = useState([]);
  const [sel, setSel] = useState([]);
  const [editing, setEditing] = useState(null);
  const [groupFocus, setGroupFocus] = useState(null);
  const [tool, setTool] = useState(null);
  const [zoom, setZoomS] = useState(1);
  const [fit, setFit] = useState(true);
  const [view, setView] = useState('normal');
  const [tab, setTab] = useState('home');
  const [pane, setPane] = useState(null);
  const [notesPane, setNotesPane] = useState(ss.notes !== false);
  const [thumbs, setThumbs] = useState(true);
  const [guides, setGuides] = useState(false);
  const [grid, setGrid] = useState(false);
  const [snap, setSnap] = useState(true);
  const [spell, setSpell] = useState(gen.spellcheck !== false);
  const [presenterView, setPresenterView] = useState(ss.presenterView !== false);
  const [present, setPresent] = useState(null);
  const [dialog, setDialog] = useState(null);
  const [painter, setPainter] = useState(null);
  const [collapsed, setCollapsed] = useState({});
  const [caretTick, setCaretTick] = useState(0);
  const [notesH, setNotesH] = useState(110);
  const textApi = useRef(null);
  const rootRef = useRef(null);
  const [offscreen, renderOff] = useOffscreen();
  const readOnly = !!ro0;

  /* ---------------- load */
  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        let d; let fid = []; let pres = []; let fmt = srcExt;
        if (bytes && bytes.length) {
          const r = await readAny(bytes, srcExt);
          d = r.doc; fid = r.fidelity; pres = r.preserved; fmt = r.fmt;
          original.current = { bytes, fmt };
        } else {
          d = newDoc({ size: ss.size || '16:9', theme: ss.theme || 'hostwerk', font: ss.font });
          d = addSlide(d, 'title').doc;
        }
        if (!d.slides.length) d = addSlide(d, 'titleContent').doc;
        if (!alive) return;
        const items = [...new Set([...(fid || []), ...(pres || [])])].map((k) => t(FIDELITY[k] || k));
        setFidelity(items);
        if (host.fidelity) { try { host.fidelity({ items, kept: !(fid || []).filter((k) => k !== 'signature').length }); } catch { /* old shell */ } }
        savedDoc.current = d;
        setHist(history(d));
        setCurId(d.slides[0].id);
        setSelSlides([d.slides[0].id]);
        setPhase('ready');
      } catch (e) {
        if (!alive) return;
        setError(e); setPhase('error');
      }
    })();
    return () => { alive = false; };
  }, [bytes, srcExt]); // eslint-disable-line react-hooks/exhaustive-deps

  const slideIdx = doc ? Math.max(0, doc.slides.findIndex((s) => s.id === curId)) : 0;
  const slide = doc ? doc.slides[slideIdx] : null;
  const selEls = useMemo(() => (slide ? sel.map((id) => findEl(slide.elements, id)).filter(Boolean) : []), [slide, sel]);

  /* ---------------- change / history */
  const change = useCallback((fn, key = null) => {
    if (readOnly) return;
    setHist((hs) => {
      if (!hs) return hs;
      let nd = fn(hs.doc);
      if (!nd || nd === hs.doc) return hs;
      // keep group boxes right after child edits
      nd = { ...nd, slides: nd.slides.map((s) => { const f = fitGroups(s.elements); return f.some((e, i) => e !== s.elements[i]) ? { ...s, elements: f } : s; }) };
      return histPush(hs, nd, key);
    });
  }, [readOnly]);
  const undo = useCallback(() => { if (editing && textApi.current) textApi.current.flush(); setEditing(null); setHist((hs) => (hs ? histUndo(hs) : hs)); }, [editing]);
  const redo = useCallback(() => { setEditing(null); setHist((hs) => (hs ? histRedo(hs) : hs)); }, []);

  // dirty + status
  useEffect(() => {
    if (!doc) return;
    if (host.dirty) host.dirty(doc !== savedDoc.current);
  }, [doc]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!doc || !host.status) return;
    const n = doc.slides.length;
    host.status(`${t('Slide {n} of {total}', { n: slideIdx + 1, total: n })} · ${doc.theme.name || ''}`);
  }, [doc && doc.slides.length, slideIdx, doc && doc.theme.name]); // eslint-disable-line react-hooks/exhaustive-deps
  // keep current slide valid
  useEffect(() => { if (doc && !doc.slides.some((s) => s.id === curId) && doc.slides.length) setCurId(doc.slides[Math.min(slideIdx, doc.slides.length - 1)].id); }, [doc, curId, slideIdx]);
  useEffect(() => { setSel([]); setEditing(null); setGroupFocus(null); }, [curId]);
  useEffect(() => { if (slide) setSel((s) => s.filter((id) => findEl(slide.elements, id))); }, [slide]);
  // ribbon state follows the caret while editing
  useEffect(() => {
    if (!editing) return undefined;
    const f = () => setCaretTick((x) => x + 1);
    document.addEventListener('selectionchange', f);
    return () => document.removeEventListener('selectionchange', f);
  }, [editing]);

  /* ---------------- save & export */
  const writeFmt = useMemo(() => {
    if (file.converted) return file.converted.to;
    const e = String(file.ext || '').toLowerCase();
    if ([...OOXML, 'odp', 'otp'].includes(e)) return e;
    return saveFormat(e);
  }, [file.ext, file.converted]);
  const chartImage = useCallback(async () => null, []);
  const docRef = useRef(doc); docRef.current = doc;
  const serialize = useCallback(async (fmt) => {
    if (textApi.current && textApi.current.flush) textApi.current.flush();
    await new Promise((r) => setTimeout(r, 0));
    const d = histRef.current.doc;
    return writeAny(d, fmt, { original: original.current.bytes, originalFmt: original.current.fmt, chartImage });
  }, [chartImage]);
  const renderSlides = useCallback(async (indices, scale = 2, type = 'image/png') => {
    const d = histRef.current.doc;
    const out = [];
    for (const i of indices) {
      const node = await renderOff(d, d.slides[i], i);
      const c = await nodeToCanvas(node, S(d.size.w), S(d.size.h), scale);
      out.push({ canvas: c, bytes: await canvasBytes(c, type, 0.9) });
    }
    return out;
  }, [renderOff]);
  const pdfBytes = useCallback(async () => {
    const d = histRef.current.doc;
    if (host.hasConverter && host.convert) {
      try { return await host.convert(await serialize('pptx'), 'pptx', 'pdf'); } catch { /* fall back to the built-in raster PDF */ }
    }
    const idx = d.slides.map((s, i) => i).filter((i) => !d.slides[i].hidden);
    const imgs = await renderSlides(idx.length ? idx : [0], 2, 'image/jpeg');
    return imagePdf(imgs.map((x) => ({ jpeg: x.bytes, pxW: x.canvas.width, pxH: x.canvas.height })), S(d.size.w), S(d.size.h), d.meta && d.meta.title);
  }, [host, serialize, renderSlides]);
  const exportRun = useCallback(async (ext) => {
    const d = histRef.current.doc;
    if (ext === 'pdf') return pdfBytes();
    if (ext === 'png') { const [r] = await renderSlides([slideIdxRef.current], 2); return r.bytes; }
    if (ext === 'zip') {
      const { default: JSZip } = await import('jszip');
      const z = new JSZip();
      const all = await renderSlides(d.slides.map((s, i) => i), 2);
      all.forEach((r, i) => z.file(`${t('Slide')} ${String(i + 1).padStart(2, '0')}.png`, r.bytes));
      return z.generateAsync({ type: 'uint8array' });
    }
    try { return await serialize(ext); } catch (e) {
      // a download is a new file: a regenerated package is fine, but say what it lacks
      if (e && e.code === 'PPTX_PATCH_FAILED' && typeof e.regenerate === 'function') {
        toast.warn(t('Exported without the parts this editor does not support (for example SmartArt, embedded objects or macros).'));
        return e.regenerate();
      }
      throw e;
    }
  }, [pdfBytes, renderSlides, serialize]);
  const slideIdxRef = useRef(0); slideIdxRef.current = slideIdx;
  /*
   * The original package could not be patched (writePptx → err.code 'PPTX_PATCH_FAILED'): saving over the file would
   * drop parts this editor does not model, so the save fails visibly and we offer "Save as a copy" – a regenerated
   * package saved as "<name> (copy).<ext>" next to the original with host.saveCopy() – the editor goes on with the copy.
   * Never silent.
   */
  const patchAsk = useRef({ open: false, declined: false });
  const offerCopy = useCallback(async (err) => {
    const st = patchAsk.current;
    if (st.open || (st.declined && !explicitSave.current)) return;
    explicitSave.current = false;
    st.open = true;
    let ok = false;
    try {
      ok = await confirmDialog({
        title: t('Save as a copy?'),
        text: t('This presentation cannot be saved over the original without losing parts this editor does not support (for example SmartArt, embedded objects or macros). Save your changes as a copy next to it – the original stays unchanged.'),
        confirm: t('Save as a copy'),
      });
    } finally { st.open = false; }
    if (!ok) { st.declined = true; return; }
    st.declined = false;
    try {
      const bytes = await err.regenerate();
      const { base, ext } = splitName(file.name, writeFmt);
      const name = t('{name} (copy)', { name: base }) + `.${ext}`;
      if (host.saveCopy) {
        // the shell creates the copy next to the original and goes on with it in this editor (the copy is then the
        // package later saves patch – a regenerated one, so they work)
        original.current = { bytes, fmt: ext };
        await host.saveCopy(bytes, ext, name);
        savedDoc.current = histRef.current.doc;
      } else if (host.openAs) {
        pendingCopy.name = name; // older shells: the editor opened by openAs saves it at once (next to the original)
        if (host.dirty) host.dirty(false); // the changes are in the copy – no "leave without saving?" question
        await host.openAs({ bytes, ext, name });
      } else if (host.download) host.download(bytes, name, MIME_OF[ext] || 'application/octet-stream');
    } catch (e) { pendingCopy.name = null; toast.error(e); }
  }, [file.name, writeFmt, host]);
  const saveOrOffer = useCallback(async (fmt) => {
    try { return await serialize(fmt); } catch (e) {
      if (e && e.code === 'PPTX_PATCH_FAILED' && typeof e.regenerate === 'function') {
        setTimeout(() => { offerCopy(e); }, 0);
        throw Object.assign(new Error(t('Not saved – this presentation can only be saved as a copy (the original stays unchanged).')), { code: e.code, cause: e });
      }
      throw e;
    }
  }, [serialize, offerCopy]);
  useEffect(() => {
    if (phase !== 'ready') return;
    if (host.setSaver && !readOnly) {
      host.setSaver(async () => {
        const out = await saveOrOffer(writeFmt);
        savedDoc.current = histRef.current.doc;
        return writeFmt === String(file.ext).toLowerCase() ? out : { bytes: out, ext: writeFmt };
      }, { ext: writeFmt });
      // opened by "Save as a copy": create the copy now
      if (pendingCopy.name && pendingCopy.name === file.name && host.save) { pendingCopy.name = null; setTimeout(() => host.save(), 0); }
    }
    if (host.setExporter) {
      host.setExporter({
        formats: [
          { ext: 'pptx', label: t('PowerPoint presentation (.pptx)') }, { ext: 'odp', label: t('OpenDocument presentation (.odp)') },
          { ext: 'pdf', label: t('PDF') }, { ext: 'png', label: t('PNG image of the current slide') }, { ext: 'zip', label: t('All slides as PNG images (.zip)') },
          ...(MACRO_EXT(file.ext) ? [{ ext: 'pptm', label: t('PowerPoint with macros (.pptm)') }] : []),
          { ext: 'potx', label: t('PowerPoint template (.potx)') }, { ext: 'ppsx', label: t('PowerPoint show (.ppsx)') },
        ],
        run: exportRun,
      });
    }
  }, [phase, readOnly, writeFmt, serialize, exportRun]); // eslint-disable-line react-hooks/exhaustive-deps

  const print = useCallback(async () => {
    const d = histRef.current.doc;
    const idx = d.slides.map((s, i) => i).filter((i) => !d.slides[i].hidden);
    const imgs = await renderSlides(idx, 2, 'image/jpeg');
    const urls = imgs.map((x) => x.canvas.toDataURL('image/jpeg', 0.9));
    if (host.print) host.print(printHtml(urls, S(d.size.w), S(d.size.h), { title: d.meta && d.meta.title }));
  }, [renderSlides, host]);

  /* ---------------- helpers */
  const center = (w, hh) => ({ x: Math.round((doc.size.w - w) / 2), y: Math.round((doc.size.h - hh) / 2) });
  const addToSlide = (els, { edit = false } = {}) => {
    change((d) => addEls(d, slide.id, els));
    setSel(els.map((e) => e.id));
    if (edit) setTimeout(() => setEditing({ id: els[0].id }), 0);
  };
  const forTextEls = (fn) => (e) => {
    if (e.type === 'table') return { ...e, rows: e.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, text: fn(c.text) })) })) };
    if (e.type === 'group') return { ...e, children: e.children.map((c) => forTextEls(fn)(c)) };
    return e.text ? { ...e, text: fn(e.text) } : e;
  };
  const runCmd = (cmd, value, modelFn) => {
    if (editing && textApi.current) { textApi.current.run(cmd, value); return; }
    if (!sel.length) return;
    change((d) => updateEls(d, slide.id, sel, forTextEls(modelFn)));
  };
  const paraCmd = (patchFn) => {
    if (editing && textApi.current) { textApi.current.para(patchFn); return; }
    if (!sel.length) return;
    change((d) => updateEls(d, slide.id, sel, forTextEls((b) => mapParas(b, (p) => { const np = { ...p, ...patchFn(p) }; for (const k of Object.keys(np)) if (np[k] === undefined) delete np[k]; return np; }))));
  };

  /* ---------------- ribbon state */
  const fmt = useMemo(() => {
    const o = {};
    if (!doc) return o;
    const e = selEls[0];
    let rp = null; let pp = null; let bodyB = null;
    if (editing && textApi.current && textApi.current.caret) { const cr = textApi.current.caret(); if (cr) { rp = cr.run; pp = cr.para; } }
    if (e) {
      bodyB = e.text || (e.type === 'table' ? e.rows[0].cells[0].text : null);
      if (!rp && bodyB) { const p = bodyB.paras[0] || { runs: [] }; rp = runProps(p.runs[0] || p.endProps || {}, p, bodyB); pp = paraProps(p, bodyB, doc.theme); o.rawFont = (p.runs[0] || {}).font; }
      o.fill = e.fill && e.fill.type === 'solid' ? e.fill.color : null;
      o.line = e.line ? e.line.color : null;
      o.shadow = !!e.shadow;
      o.anchor = bodyB && bodyB.anchor;
    }
    if (rp) Object.assign(o, { font: rp.font && rp.font[0] === '+' ? (rp.font.startsWith('+mj') ? doc.theme.fonts.major : doc.theme.fonts.minor) : rp.font, fontRaw: rp.font, size: rp.size, b: rp.b, i: rp.i, u: rp.u, s: rp.s, color: rp.color, highlight: rp.highlight, baseline: rp.baseline || 0 });
    if (pp) Object.assign(o, { align: pp.align, lnSpc: pp.lnSpc, bullet: pp.bullet ? (pp.bullet.num ? 'num' : 'char') : null });
    return o;
  }, [doc, selEls, editing, caretTick]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------------- insertion */
  const createAt = (tl, box) => {
    if (tl === '__copy') {
      const els = box.ids.map((id) => findEl(slide.elements, id)).filter(Boolean);
      const copies = reId(els.map(stripRaw)).map((x) => moveEl(x, box.dx, box.dy));
      addToSlide(copies);
      return;
    }
    const isLine = ['line', 'arrow', 'doubleArrow', 'elbow', 'curve'].includes(tl);
    if (isLine) {
      const w = box.w || 1828800; const hh = box.w || box.h ? box.h : 0;
      const x1 = box.x; const y1 = box.y; const x2 = box.x + w; const y2 = box.y + hh;
      const e = lineEl(Math.min(x1, x2), Math.min(y1, y2), Math.max(1, Math.abs(x2 - x1)), Math.max(1, Math.abs(y2 - y1)), { flipH: x2 < x1, flipV: y2 < y1, geom: tl === 'elbow' ? 'elbow' : tl === 'curve' ? 'curve' : 'line' });
      e.line = { ...e.line, tail: tl === 'arrow' || tl === 'doubleArrow' || tl === 'elbow' || tl === 'curve' ? 'triangle' : 'none', head: tl === 'doubleArrow' ? 'triangle' : 'none' };
      addToSlide([e]);
      return;
    }
    const def = tl === 'textbox' ? { w: cmToEmu(8), h: cmToEmu(1.2) } : { w: cmToEmu(5), h: tl === 'rect' || tl === 'roundRect' ? cmToEmu(3) : cmToEmu(5) };
    const b = box.w > 0 ? box : { ...box, w: def.w, h: def.h };
    if (tl === 'textbox') {
      const e = textBox(b.x, b.y, b.w, b.h, body([para('', {}, {})], { autofit: 'resize', wrap: box.w > 0 }));
      addToSlide([e], { edit: true });
      return;
    }
    const e = shape(SHAPES[tl] ? tl : 'rect', b.x, b.y, b.w, b.h);
    e.text.paras[0].endProps = { color: 'lt1' };
    e.text.lvl = Array.from({ length: 9 }, () => ({ color: 'lt1', align: 'ctr', size: 18 }));
    addToSlide([e]);
  };

  const insertPicture = async (src, target) => {
    const sz = await imgSize(src);
    if (target) {
      // fill the placeholder box, crop to cover
      const ar = sz.w / sz.h; const br = target.w / target.h;
      const crop = ar > br ? { l: (1 - br / ar) / 2, r: (1 - br / ar) / 2, t: 0, b: 0 } : { t: (1 - ar / br) / 2, b: (1 - ar / br) / 2, l: 0, r: 0 };
      const e = imageEl(src, target.x, target.y, target.w, target.h, { crop });
      change((d) => updateSlide(d, slide.id, (s) => ({ ...s, elements: s.elements.map((x) => (x.id === target.id ? e : x)) })));
      setSel([e.id]);
      return;
    }
    const maxW = doc.size.w * 0.7; const maxH = doc.size.h * 0.7;
    let w = sz.w * 9525; let hh = sz.h * 9525;
    const f = Math.min(1, maxW / w, maxH / hh); w *= f; hh *= f;
    const c = center(w, hh);
    addToSlide([imageEl(src, c.x, c.y, w, hh)]);
  };
  const phTarget = () => { const e = selEls[0]; return e && (e.ph === 'pic' || e.photoSlot) ? e : null; };

  /* ---------------- clipboard */
  const copySel = (cut = false) => {
    if (!sel.length) return;
    const els = selEls.map((e) => deepClone(e));
    const text = els.map((e) => (e.text ? e.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n') : '')).filter(Boolean).join('\n');
    CLIP = { type: 'els', els, theme: doc.theme, marker: `of-slides:${uid('c')}`, text };
    try { navigator.clipboard.writeText(text || ' ').catch(() => {}); } catch { /* ignore */ }
    if (cut) { change((d) => deleteEls(d, slide.id, sel)); setSel([]); }
  };
  const pasteEls = (offset = true) => {
    if (!CLIP || CLIP.type !== 'els') return false;
    const shift = offset && CLIP.els.some((e) => slide.elements.some((x) => x.x === e.x && x.y === e.y)) ? cmToEmu(0.5) : 0;
    const els = reId(CLIP.els.map(stripRaw)).map((e) => moveEl(e, shift, shift));
    if (shift) CLIP = { ...CLIP, els: CLIP.els.map((e) => moveEl(e, shift, shift)) };
    addToSlide(els);
    return true;
  };
  const pasteFromSystem = async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          const type = it.types.find((x) => x.startsWith('image/'));
          if (type) { const blob = await it.getType(type); await insertPicture(await readFile(blob), phTarget()); return true; }
        }
        for (const it of items) if (it.types.includes('text/plain')) {
          const s = await (await it.getType('text/plain')).text();
          if (CLIP && CLIP.text === s && pasteEls()) return true;
          if (s.trim()) { pasteText(s); return true; }
        }
      }
    } catch { /* permission denied → internal clipboard */ }
    return pasteEls();
  };
  const pasteText = (s) => {
    const w = cmToEmu(16); const c = center(w, cmToEmu(2));
    addToSlide([textBox(c.x, c.y, w, cmToEmu(2), body(s.split(/\r?\n/).map((l) => para(l)), { autofit: 'resize' }))]);
  };
  useEffect(() => {
    const onPaste = async (e) => {
      if (readOnly || editing || !rootRef.current || !rootRef.current.contains(document.activeElement) || /INPUT|TEXTAREA/.test(document.activeElement.tagName)) return;
      if (view !== 'normal' && view !== 'sorter') return;
      const cd = e.clipboardData;
      if (!cd) return;
      if (CLIP && CLIP.type === 'slides' && (view === 'sorter' || document.activeElement.closest('.of-slides-thumbs'))) { e.preventDefault(); pasteSlides(); return; }
      const file0 = [...(cd.files || [])].find((f) => /^image\//.test(f.type));
      if (file0) { e.preventDefault(); await insertPicture(await readFile(file0), phTarget()); return; }
      const s = cd.getData('text/plain');
      e.preventDefault();
      if (CLIP && CLIP.type === 'els' && (!s || s === CLIP.text || s === ' ')) { pasteEls(); return; }
      if (s && s.trim()) pasteText(s);
      else pasteEls();
    };
    document.addEventListener('paste', onPaste);
    return () => document.removeEventListener('paste', onPaste);
  });
  const copySlides = (cut = false) => {
    const ids = selSlides.length ? selSlides : [curId];
    CLIP = { type: 'slides', slides: doc.slides.filter((s) => ids.includes(s.id)).map((s) => deepClone(s)), marker: uid('c') };
    try { navigator.clipboard.writeText(CLIP.slides.map(slideTitle).join('\n') || ' ').catch(() => {}); } catch { /* ignore */ }
    if (cut) delSlides(ids, true);
  };
  const pasteSlides = () => {
    if (!CLIP || CLIP.type !== 'slides') return;
    const sl = CLIP.slides.map((s) => {
      const map = {};
      const n = { ...deepClone(s), id: uid('s'), elements: reId(s.elements.map(stripRaw), map) };
      n.animations = (n.animations || []).filter((a) => map[a.el]).map((a) => ({ ...a, id: uid('a'), el: map[a.el] }));
      if (!doc.layouts.some((l) => l.id === n.layoutId)) { const lk = layoutByKind(doc, 'titleContent'); n.layoutId = lk && lk.id; }
      delete n.timingRaw; delete n.srcPart; delete n.notesPart; delete n.commentsPart; delete n.transitionRaw; delete n.bgRaw;
      return n;
    });
    change((d) => insertSlides(d, sl, slideIdx + 1));
    setCurId(sl[0].id); setSelSlides(sl.map((s) => s.id));
  };

  /* ---------------- slides */
  const newSlideCmd = (kind) => {
    const layout = slide && doc.layouts.find((l) => l.id === slide.layoutId);
    const k = kind || (layout && layout.kind === 'title' ? 'titleContent' : (layout && layout.kind) || 'titleContent');
    const r = addSlide(doc, k, slideIdx + 1);
    const s = { ...r.slide, transition: { ...r.slide.transition, type: ss.transition && ss.transition !== 'none' ? (ss.transition === 'slide' ? 'push' : ss.transition) : (slide && slide.transition && slide.transition.type) || 'none' } };
    change((d) => insertSlides(d, [s], slideIdx + 1));
    setCurId(s.id); setSelSlides([s.id]);
  };
  const delSlides = async (ids, silent) => {
    if (!ids.length) return;
    if (ids.length >= doc.slides.length) { toast.warn(t('A presentation needs at least one slide.')); return; }
    if (!silent && ids.length > 1 && gen.confirmDelete !== false) { const ok = await confirmDialog({ title: t('Delete {n} slides?', { n: ids.length }), text: t('You can undo this with Ctrl+Z.'), confirm: t('Delete'), danger: true }); if (!ok) return; }
    const firstIdx = doc.slides.findIndex((s) => ids.includes(s.id));
    change((d) => deleteSlides(d, ids));
    const rest = doc.slides.filter((s) => !ids.includes(s.id));
    const n = rest[Math.min(firstIdx, rest.length - 1)];
    if (n) { setCurId(n.id); setSelSlides([n.id]); }
  };
  const dupSlides = () => { const ids = selSlides.length ? selSlides : [curId]; const r = duplicateSlides(doc, ids); change(() => r.doc); if (r.ids[0]) { setCurId(r.ids[0]); setSelSlides(r.ids); } };
  const selectSlide = (id, e) => {
    if (e && (e.ctrlKey || e.metaKey)) setSelSlides((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
    else if (e && e.shiftKey) { const a = doc.slides.findIndex((s) => s.id === curId); const b = doc.slides.findIndex((s) => s.id === id); const [x, y] = [Math.min(a, b), Math.max(a, b)]; setSelSlides(doc.slides.slice(x, y + 1).map((s) => s.id)); }
    else setSelSlides([id]);
    setCurId(id);
  };
  const curSection = useMemo(() => { if (!doc) return null; const g = sectionGroups(doc).find((x) => x.slides.some((s) => s.id === curId)); return g && g.section; }, [doc, curId]);

  /* ---------------- presentation */
  const openPresenterWindow = useCallback((token) => {
    if (!file || !file.path) return false;
    const q = [file.root ? `root=${encodeURIComponent(file.root)}` : '', `path=${encodeURIComponent(file.path)}`, 'ro=1', `present=${token}`, 'window=1'].filter(Boolean).join('&');
    try { const w = window.open(`${location.pathname}${location.search}#/office-edit?${q}`, '_blank'); return !!w; } catch { return false; }
  }, [file]);
  const startShow = (from) => {
    if (editing && textApi.current) textApi.current.flush();
    setEditing(null);
    const start = from === 'start' || from === 'rehearse' ? 0 : slideIdx;
    setPresent({ start, presenter: from === 'presenter' ? 'same' : (presenterView && from !== 'rehearse' ? 'window' : false), rehearse: from === 'rehearse' });
  };

  /* ---------------- commands for the ribbon and menus */
  const c = {
    doc, slide, selEls, fmt, readOnly, curSection,
    state: { tool, painter, pane, view, guides, grid, snap, notesPane, thumbs, fit, spellcheck: spell, presenterView, canUndo: hist && hist.past.length > 0, canRedo: hist && hist.future.length > 0, cell: editing && editing.cell },
    undo, redo, setTool: (x) => setTool((cur) => (cur === x ? null : x)),
    present: startShow,
    newSlide: newSlideCmd,
    setLayout: (kind) => change((d) => resetSlide(d, slide.id, kind)),
    resetSlide: () => change((d) => resetSlide(d, slide.id)),
    addSection: async () => { const name = await promptDialog({ title: t('Add section'), label: t('Section name'), value: t('Untitled section'), confirm: t('Add') }); if (name) change((d) => addSection(d, curId, name)); },
    renameSection: async () => { if (!curSection) return; const name = await promptDialog({ title: t('Rename section'), label: t('Section name'), value: curSection.name, confirm: t('Rename') }); if (name) change((d) => renameSection(d, curSection.id, name)); },
    removeSection: () => curSection && change((d) => removeSection(d, curSection.id)),
    paste: () => pasteFromSystem(), cut: () => copySel(true), copy: () => copySel(false),
    formatPainter: () => {
      if (painter) { setPainter(null); return; }
      const e = selEls[0]; if (!e) return;
      const b = e.text; const p = b && b.paras[0]; const r = p && (p.runs[0] || p.endProps);
      const { text, ...rp } = r || {}; // eslint-disable-line no-unused-vars
      setPainter({ fill: e.fill, line: e.line, shadow: e.shadow, run: rp, para: p ? { align: p.align, lnSpc: p.lnSpc, spcBef: p.spcBef, spcAft: p.spcAft } : null, type: e.type });
    },
    fontFamily: (f) => runCmd('font', f, (b) => setRunProps(b, { font: f })),
    fontSize: (n) => runCmd('size', n, (b) => setRunProps(b, { size: n })),
    growFont: (dir) => {
      const cur = fmt.size || 18;
      const list = SIZES_PT;
      const n = dir > 0 ? (list.find((x) => x > cur) || cur + 8) : ([...list].reverse().find((x) => x < cur) || Math.max(1, cur - 1));
      if (editing && textApi.current) { textApi.current.run('size', n); return; }
      change((d) => updateEls(d, slide.id, sel, forTextEls((b) => mapRuns(b, (r) => { const rpp = runProps(r, null, b); const s = rpp.size || cur; const ns = dir > 0 ? (list.find((x) => x > s) || s + 8) : ([...list].reverse().find((x) => x < s) || Math.max(1, s - 1)); return { ...r, size: ns }; }))));
    },
    clearFormat: () => runCmd('removeFormat', null, (b) => mapRuns(b, (r) => ({ text: r.text, ...(r.field ? { field: r.field } : {}), ...(r.link ? { link: r.link } : {}) }))),
    toggleRun: (k) => runCmd(k, null, (b) => setRunProps(b, { [k]: !fmt[k] })),
    baseline: (v) => runCmd(v > 0 ? 'sup' : 'sub', null, (b) => setRunProps(b, { baseline: (fmt.baseline > 0) === (v > 0) && fmt.baseline ? undefined : v })),
    changeCase: (mode) => {
      if (editing && textApi.current) { const s = String(window.getSelection() || ''); if (s) textApi.current.insertText(caseOf(s, mode)); return; }
      change((d) => updateEls(d, slide.id, sel, forTextEls((b) => mapRuns(b, (r) => ({ ...r, text: caseOf(r.text || '', mode) })))));
    },
    fontColor: (col) => runCmd('color', col, (b) => setRunProps(b, { color: col })),
    highlight: (col) => runCmd('highlight', col, (b) => setRunProps(b, { highlight: col || undefined })),
    bullets: (ch) => paraCmd(() => ({ bullet: ch ? { char: ch } : false, marL: undefined, indent: undefined, ...(ch ? { marLAuto: undefined } : {}) })),
    numbering: (sc) => paraCmd(() => ({ bullet: sc ? { num: sc } : false })),
    level: (d) => paraCmd((p) => ({ level: Math.max(0, Math.min(8, (p.level || 0) + d)), marL: undefined, indent: undefined })),
    lineSpacing: (v) => paraCmd(() => ({ lnSpc: v })),
    align: (a) => paraCmd(() => ({ align: a })),
    anchor: (a) => change((d) => updateEls(d, slide.id, sel, (e) => (e.text ? { ...e, text: { ...e.text, anchor: a } } : e))),
    arrange: (op) => change((d) => zorder(d, slide.id, sel, op)),
    group: () => { const r = groupEls(doc, slide.id, sel); if (r.id) { change(() => r.doc); setSel([r.id]); } },
    ungroup: () => { const r = ungroupEls(doc, slide.id, sel); change(() => r.doc); setSel(r.ids); },
    alignEls: (how) => change((d) => alignEls(d, slide.id, sel, how, sel.length > 1 ? 'selection' : 'slide')),
    distribute: (dir) => change((d) => distributeEls(d, slide.id, sel, dir)),
    rotate: (deg) => change((d) => updateEls(d, slide.id, sel, (e) => ({ ...e, rot: (((e.rot || 0) + deg) % 360 + 360) % 360 }))),
    flip: (ax) => change((d) => updateEls(d, slide.id, sel, (e) => (ax === 'h' ? { ...e, flipH: !e.flipH } : { ...e, flipV: !e.flipV }))),
    toggleLock: () => { const on = !selEls.some((e) => e.locked); change((d) => updateEls(d, slide.id, sel, (e) => ({ ...e, locked: on }))); },
    shapeFill: (col) => change((d) => updateEls(d, slide.id, sel, (e) => (e.type === 'shape' || e.type === 'table' ? (e.type === 'table' ? { ...e, rows: e.rows.map((r) => ({ ...r, cells: r.cells.map((x) => ({ ...x, fill: col ? { type: 'solid', color: col } : null })) })) } : { ...e, fill: col ? { type: 'solid', color: col } : null }) : e))),
    shapeLine: (col) => change((d) => updateEls(d, slide.id, sel, (e) => (e.type === 'line' ? { ...e, line: { ...e.line, color: col || e.line.color } } : { ...e, line: col ? { width: 12700, dash: 'solid', ...(e.line || {}), color: col } : null }))),
    lineWidth: (w) => change((d) => updateEls(d, slide.id, sel, (e) => ({ ...e, line: { color: 'tx1', dash: 'solid', ...(e.line || {}), width: Math.round(w * EMU_PER_PT) } }))),
    lineDash: (dash) => change((d) => updateEls(d, slide.id, sel, (e) => ({ ...e, line: { color: 'tx1', width: 12700, ...(e.line || {}), dash } }))),
    toggleShadow: () => { const on = !fmt.shadow; change((d) => updateEls(d, slide.id, sel, (e) => ({ ...e, shadow: on ? { color: '#000000', blur: 50800, dist: 38100, dir: 45, alpha: 35 } : null }))); },
    quickStyle: (q) => change((d) => updateEls(d, slide.id, sel, (e) => {
      if (e.type !== 'shape') return e;
      const fill = q.fill === 'grad1' ? { type: 'grad', angle: 45, stops: [{ pos: 0, color: 'accent1' }, { pos: 100, color: 'accent2' }] } : q.fill === 'grad2' ? { type: 'grad', angle: 45, stops: [{ pos: 0, color: 'accent3' }, { pos: 100, color: 'accent5' }] } : q.fill ? { type: 'solid', color: q.fill } : null;
      return { ...e, fill, line: q.line ? { color: q.line, width: 19050, dash: 'solid' } : null, text: e.text ? setRunProps(e.text, { color: q.text }) : e.text };
    })),
    togglePane: (p, force) => setPane((cur) => (force ? p : cur === p ? null : p)),
    insertTable: (r, cc) => { const w = Math.min(doc.size.w * 0.86, cc * cmToEmu(4)); const hh = r * cmToEmu(1.1); const p = center(w, hh); addToSlide([tableEl(r, cc, p.x, p.y, w, hh)]); },
    pictureFromFile: async () => { const fs = await pickFile('image/*', true); let target = phTarget(); for (const f of fs) { await insertPicture(await readFile(f), target); target = null; } },
    pictureFromClipboard: async () => { const ok = await pasteFromSystem(); if (!ok) toast.warn(t('The clipboard contains no picture.')); },
    pictureFromUrl: async () => {
      const url = await promptDialog({ title: t('Picture from a web address'), label: t('Address of the picture'), placeholder: 'https://…', confirm: t('Insert') });
      if (!url) return;
      try { const r = await fetch(url); if (!r.ok) throw new Error(String(r.status)); const blob = await r.blob(); if (!/^image\//.test(blob.type)) throw new Error(t('That address is not a picture.')); await insertPicture(await readFile(blob), phTarget()); } catch (e) { toast.error(t('Could not load the picture ({error}). Save it on your computer and insert it from there.', { error: e.message })); }
    },
    insertEmoji: (em) => { if (editing && textApi.current) { textApi.current.insertText(em); return; } const s = cmToEmu(3); const p = center(s, s); addToSlide([textBox(p.x, p.y, s, s, body([para(em, { align: 'ctr' }, { size: 72 })], { anchor: 'ctr', autofit: 'none' }))]); },
    insertSymbol: (sy) => { if (editing && textApi.current) { textApi.current.insertText(sy); return; } const p = center(cmToEmu(2), cmToEmu(1.5)); addToSlide([textBox(p.x, p.y, cmToEmu(2), cmToEmu(1.5), body([para(sy, { align: 'ctr' }, { size: 32 })], { autofit: 'resize' }))]); },
    insertChart: (kind) => { const w = doc.size.w * 0.6; const hh = doc.size.h * 0.6; const p = center(w, hh); const e = chartEl(kind, p.x, p.y, w, hh); addToSlide([e]); setTimeout(() => setDialog({ type: 'chart', id: e.id }), 50); },
    insertLink: async () => {
      const cur = selEls[0] && selEls[0].link;
      const url = await promptDialog({ title: t('Insert link'), label: t('Address'), value: cur || 'https://', confirm: t('Insert') });
      if (url == null) return;
      if (editing && textApi.current) { textApi.current.run('link', url && url !== 'https://' ? url : null); return; }
      change((d) => updateEls(d, slide.id, sel, (e) => ({ ...e, link: url && url !== 'https://' ? url : undefined })));
    },
    footerDialog: () => setDialog({ type: 'footer' }),
    insertMedia: async (kind) => {
      const [f] = await pickFile(kind === 'video' ? 'video/*' : 'audio/*');
      if (!f) return;
      if (f.size > 200 * 1024 * 1024) { toast.error(t('This file is too large to embed (over 200 MB).')); return; }
      if (f.size > 50 * 1024 * 1024) toast.warn(t('Large media files make the presentation slow to open and save.'));
      const src = await readFile(f);
      const w = kind === 'video' ? doc.size.w * 0.6 : cmToEmu(2.5); const hh = kind === 'video' ? w * 9 / 16 : cmToEmu(2.5);
      const p = center(w, hh);
      addToSlide([mediaEl(kind, src, f.type || (kind === 'video' ? 'video/mp4' : 'audio/mpeg'), p.x, p.y, w, hh, { name: f.name })]);
    },
    applyTheme: async (id) => {
      if (doc.theme.id === 'imported') { const ok = await confirmDialog({ title: t('Apply a new theme?'), text: t('The design of this file (masters and layouts) is replaced. Text and pictures stay.'), confirm: t('Apply theme') }); if (!ok) return; }
      change((d) => applyTheme(d, id));
    },
    themeColors: (colors) => change((d) => setThemeColors(d, colors)),
    themeFonts: (fonts) => change((d) => setThemeFonts(d, fonts)),
    slideSize: (sz) => change((d) => setSize(d, sz, 'fit')),
    customSize: () => setDialog({ type: 'size' }),
    setBackground: (fill) => change((d) => updateSlides(d, selSlides.length ? selSlides : [curId], { background: fill })),
    backgroundToAll: () => { const fill = slide.background; change((d) => ({ ...d, slides: d.slides.map((s) => ({ ...s, background: fill })) })); },
    backgroundPicture: async () => { const [f] = await pickFile('image/*'); if (!f) return; const src = await readFile(f); change((d) => updateSlides(d, selSlides.length ? selSlides : [curId], { background: { type: 'image', src, mode: 'stretch' } })); },
    hideBackgroundGraphics: (on) => change((d) => updateSlides(d, selSlides.length ? selSlides : [curId], { showMaster: on ? false : undefined })),
    setTransition: (patch) => change((d) => updateSlides(d, selSlides.length ? selSlides : [curId], (s) => ({ ...s, transition: { type: 'none', dir: 'l', duration: 500, advanceAfter: null, ...(s.transition || {}), ...patch } })), 'transition'),
    transitionToAll: () => { const tr = slide.transition; change((d) => ({ ...d, slides: d.slides.map((s) => ({ ...s, transition: { ...tr } })) })); toast.success(t('Applied to all slides')); },
    previewTransition: () => setPresent({ start: slideIdx, presenter: false, preview: true }),
    setAnim: (patch) => {
      if (!sel.length) return;
      change((d) => updateSlide(d, slide.id, (s) => {
        let anims = (s.animations || []).slice();
        for (const id of sel) {
          const i = anims.findIndex((a) => a.el === id);
          if (i >= 0) anims[i] = { ...anims[i], ...patch };
          else anims.push({ id: uid('a'), el: id, effect: 'fade', dir: 'b', trigger: 'click', duration: 500, delay: 0, ...patch });
        }
        anims = anims.map((a) => (a.effect === 'appear' ? { ...a, duration: a.duration || 1 } : a));
        return { ...s, animations: anims };
      }), 'anim');
    },
    removeAnim: () => change((d) => updateSlide(d, slide.id, (s) => ({ ...s, animations: (s.animations || []).filter((a) => !sel.includes(a.el)) }))),
    moveAnim: (dir) => change((d) => updateSlide(d, slide.id, (s) => { const a = (s.animations || []).slice(); const i = a.findIndex((x) => sel.includes(x.el)); const j = i + dir; if (i < 0 || j < 0 || j >= a.length) return s; [a[i], a[j]] = [a[j], a[i]]; return { ...s, animations: a }; })),
    previewAnims: () => setPresent({ start: slideIdx, presenter: false, preview: true }),
    toggleHidden: () => { const ids = selSlides.length ? selSlides : [curId]; const on = !slide.hidden; change((d) => updateSlides(d, ids, { hidden: on })); },
    setShow: (patch) => change((d) => ({ ...d, show: { loop: false, autoAdvance: 0, useTimings: true, ...(d.show || {}), ...patch } })),
    setPresenterView,
    toggleSpell: () => setSpell((x) => !x),
    setView: (v) => { setEditing(null); setView(v); },
    toggleGuides: () => setGuides((x) => !x), toggleGrid: () => setGrid((x) => !x), toggleSnap: () => setSnap((x) => !x),
    toggleNotes: () => setNotesPane((x) => !x), toggleThumbs: () => setThumbs((x) => !x),
    zoomBy: (d) => { setFit(false); setZoomS((z) => Math.max(0.1, Math.min(4, +(z * (d > 0 ? 1.15 : 1 / 1.15)).toFixed(3)))); },
    zoomFit: () => setFit(true),
    tableOp: (op, value) => {
      const e = selEls[0]; if (!e || e.type !== 'table') return;
      const at = (editing && editing.cell) || { r: 0, c: 0 };
      if (editing && textApi.current) textApi.current.flush();
      setTimeout(() => {
        change((d) => updateEls(d, slide.id, [e.id], (x) => tableOp(x, op, at, value) || x));
        if (/^del/.test(op)) setEditing(null);
      }, 0);
    },
    tableStyle: (patch) => change((d) => updateEls(d, slide.id, sel, (e) => (e.type === 'table' ? { ...e, style: { ...(e.style || {}), ...patch } } : e))),
    chartSet: (patch, key) => change((d) => updateEls(d, slide.id, sel, (e) => (e.type === 'chart' ? { ...e, chart: { ...e.chart, ...patch } } : e)), key),
    chartData: () => selEls[0] && setDialog({ type: 'chart', id: selEls[0].id }),
    replacePicture: async () => { const [f] = await pickFile('image/*'); if (!f) return; const src = await readFile(f); change((d) => updateEls(d, slide.id, sel, (e) => ({ ...e, src, crop: { l: 0, t: 0, r: 0, b: 0 } }))); },
    resetPicture: () => change((d) => updateEls(d, slide.id, sel, (e) => ({ ...e, crop: { l: 0, t: 0, r: 0, b: 0 }, opacity: 1, geom: undefined, shadow: null, line: null }))),
    cropToShape: (g) => change((d) => updateEls(d, slide.id, sel, (e) => (e.type === 'image' && SHAPES[g] ? { ...e, geom: g } : e))),
    print,
  };

  const applyPainter = (id) => {
    const p = painter; setPainter(null);
    change((d) => updateEls(d, slide.id, [id], (e) => {
      const n = { ...e };
      if (e.type === 'shape' && p.type === 'shape') { n.fill = p.fill; n.line = p.line; n.shadow = p.shadow; }
      if (e.type === 'line' && p.line) n.line = { ...e.line, color: p.line.color, width: p.line.width, dash: p.line.dash };
      if (n.text && p.run) n.text = mapParas(setRunProps(n.text, p.run), (q) => ({ ...q, ...(p.para || {}) }));
      return n;
    }));
    setSel([id]);
  };

  /* ---------------- keyboard */
  const onKeyDown = (e) => {
    if (present || dialog) return;
    const tag = e.target.tagName;
    if (/INPUT|TEXTAREA|SELECT/.test(tag) || e.target.isContentEditable) return;
    const k = e.key.toLowerCase();
    const ctrl = e.ctrlKey || e.metaKey;
    const inThumbs = !!e.target.closest('.of-slides-thumbs, .of-slides-sorter');
    if (e.key === 'F5') { e.preventDefault(); startShow(e.shiftKey ? 'current' : 'start'); return; }
    if (ctrl && k === 's') { e.preventDefault(); explicitSave.current = true; if (host.save) host.save(); return; }
    if (ctrl && k === 'p') { e.preventDefault(); print(); return; }
    if (readOnly) {
      if (['pagedown', 'arrowdown'].includes(k) || (inThumbs && k === 'arrowright')) { e.preventDefault(); const n = doc.slides[slideIdx + 1]; if (n) selectSlide(n.id); }
      if (['pageup', 'arrowup'].includes(k) || (inThumbs && k === 'arrowleft')) { e.preventDefault(); const n = doc.slides[slideIdx - 1]; if (n) selectSlide(n.id); }
      if (ctrl && k === 'c') copySel();
      return;
    }
    if (ctrl && k === 'z') { e.preventDefault(); undo(); return; }
    if (ctrl && (k === 'y' || (e.shiftKey && k === 'z'))) { e.preventDefault(); redo(); return; }
    if (ctrl && k === 'm') { e.preventDefault(); newSlideCmd(); return; }
    if (inThumbs || view === 'sorter') {
      if (k === 'delete' || k === 'backspace') { e.preventDefault(); delSlides(selSlides.length ? selSlides : [curId]); return; }
      if (ctrl && k === 'd') { e.preventDefault(); dupSlides(); return; }
      if (ctrl && k === 'c') { e.preventDefault(); copySlides(); return; }
      if (ctrl && k === 'x') { e.preventDefault(); copySlides(true); return; }
      if (ctrl && k === 'a') { e.preventDefault(); setSelSlides(doc.slides.map((s) => s.id)); return; }
      if (k === 'arrowdown' || k === 'arrowright' || k === 'arrowup' || k === 'arrowleft') {
        e.preventDefault();
        const d = k === 'arrowdown' || k === 'arrowright' ? 1 : -1;
        if (e.ctrlKey && e.shiftKey) { change((dd) => moveSlides(dd, selSlides.length ? selSlides : [curId], Math.max(0, slideIdx + (d > 0 ? 2 : -1)))); return; }
        const n = doc.slides[slideIdx + d]; if (n) selectSlide(n.id, e.shiftKey ? { shiftKey: true } : null);
        return;
      }
      if (k === 'enter') { e.preventDefault(); newSlideCmd(); return; }
      return;
    }
    if (view !== 'normal') return;
    if (k === 'escape') { if (tool) setTool(null); else if (painter) setPainter(null); else if (groupFocus) { setSel([groupFocus]); setGroupFocus(null); } else setSel([]); return; }
    if (ctrl && k === 'a') { e.preventDefault(); setSel(slide.elements.filter((x) => !x.locked).map((x) => x.id)); return; }
    if (ctrl && k === 'c') { e.preventDefault(); copySel(); return; }
    if (ctrl && k === 'x') { e.preventDefault(); copySel(true); return; }
    if (ctrl && k === 'd') { e.preventDefault(); if (sel.length) { const copies = reId(selEls.map(stripRaw)).map((x) => moveEl(x, cmToEmu(0.5), cmToEmu(0.5))); addToSlide(copies); } else dupSlides(); return; }
    if (ctrl && k === 'g') { e.preventDefault(); if (e.shiftKey) c.ungroup(); else c.group(); return; }
    if (ctrl && k === 'k') { e.preventDefault(); c.insertLink(); return; }
    if (ctrl && (e.key === ']' || e.key === '[' || e.key === '}' || e.key === '{')) {
      e.preventDefault();
      if (sel.length && selEls.some((x) => x.text)) c.growFont(e.key === ']' || e.key === '}' ? 1 : -1);
      else c.arrange(e.shiftKey ? (e.key === '}' || e.key === ']' ? 'front' : 'back') : e.key === ']' ? 'forward' : 'backward');
      return;
    }
    if (ctrl && ['b', 'i', 'u'].includes(k) && sel.length) { e.preventDefault(); c.toggleRun(k); return; }
    if (ctrl && ['e', 'l', 'r', 'j'].includes(k) && sel.length) { e.preventDefault(); c.align({ e: 'ctr', l: 'l', r: 'r', j: 'just' }[k]); return; }
    if ((k === 'delete' || k === 'backspace') && sel.length) { e.preventDefault(); change((d) => deleteEls(d, slide.id, sel.filter((id) => !(findEl(slide.elements, id) || {}).locked))); setSel([]); return; }
    if (k.startsWith('arrow') && sel.length) {
      e.preventDefault();
      const step = e.ctrlKey ? 12700 : cmToEmu(0.1) * (e.shiftKey ? 10 : 1);
      const dx = k === 'arrowleft' ? -step : k === 'arrowright' ? step : 0; const dy = k === 'arrowup' ? -step : k === 'arrowdown' ? step : 0;
      change((d) => updateEls(d, slide.id, sel, (x) => (x.locked ? x : moveEl(x, dx, dy))), 'nudge');
      return;
    }
    if (k.startsWith('arrow') || k === 'pagedown' || k === 'pageup') {
      if (k === 'pagedown' || k === 'arrowdown' || k === 'arrowright') { e.preventDefault(); const n = doc.slides[slideIdx + 1]; if (n) selectSlide(n.id); }
      if (k === 'pageup' || k === 'arrowup' || k === 'arrowleft') { e.preventDefault(); const n = doc.slides[slideIdx - 1]; if (n) selectSlide(n.id); }
      return;
    }
    if (k === 'tab' && slide.elements.length) { e.preventDefault(); const list = slide.elements; const i = list.findIndex((x) => x.id === sel[0]); const n = list[(i + (e.shiftKey ? -1 : 1) + list.length) % list.length]; setSel([n.id]); return; }
    if ((k === 'enter' || k === 'f2') && selEls.length === 1 && (selEls[0].text || selEls[0].type === 'table') && !selEls[0].preserved) { e.preventDefault(); setEditing(selEls[0].type === 'table' ? { id: selEls[0].id, cell: { r: 0, c: 0 } } : { id: selEls[0].id, selectAll: true }); return; }
    // typing on a selected text shape starts editing (replaces nothing – caret at the end)
    if (e.key.length === 1 && !ctrl && !e.altKey && selEls.length === 1 && selEls[0].text && !selEls[0].preserved) {
      setEditing({ id: selEls[0].id });
    }
  };

  /* ---------------- context menus */
  const elMenu = (e, id) => {
    if (readOnly) { contextMenu(e, [{ label: t('Copy'), icon: Copy, onClick: () => copySel() }]); return; }
    if (!id) {
      contextMenu(e, [
        { label: t('Paste'), icon: ClipboardPaste, hint: keyLabel('Ctrl+V'), onClick: () => pasteFromSystem() },
        { label: t('Layout …'), icon: LayoutTemplate, onClick: () => setTab('home') },
        { label: t('Format background …'), onClick: () => setTab('design') },
        { label: t('New slide'), icon: Plus, hint: keyLabel('Ctrl+M'), onClick: () => newSlideCmd() },
      ]);
      return;
    }
    const el = findEl(slide.elements, id) || {};
    contextMenu(e, [
      { label: t('Cut'), icon: Scissors, hint: keyLabel('Ctrl+X'), onClick: () => copySel(true) },
      { label: t('Copy'), icon: Copy, hint: keyLabel('Ctrl+C'), onClick: () => copySel() },
      { label: t('Paste'), icon: ClipboardPaste, hint: keyLabel('Ctrl+V'), onClick: () => pasteFromSystem() },
      { label: t('Duplicate'), icon: CopyPlus, hint: keyLabel('Ctrl+D'), onClick: () => addToSlide(reId(selEls.map(stripRaw)).map((x) => moveEl(x, cmToEmu(0.5), cmToEmu(0.5)))) },
      (el.text || el.type === 'table') && !el.preserved && { label: t('Edit text'), icon: Type, hint: keyLabel('F2'), onClick: () => setEditing(el.type === 'table' ? { id, cell: { r: 0, c: 0 } } : { id }) },
      { label: t('Bring to front'), icon: BringToFront, onClick: () => c.arrange('front') },
      { label: t('Send to back'), icon: SendToBack, onClick: () => c.arrange('back') },
      sel.length > 1 && { label: t('Group'), icon: Group, hint: keyLabel('Ctrl+G'), onClick: c.group },
      el.type === 'group' && { label: t('Ungroup'), icon: Ungroup, hint: keyLabel('Ctrl+Shift+G'), onClick: c.ungroup },
      { label: t('Link …'), icon: Link2, hint: keyLabel('Ctrl+K'), onClick: c.insertLink },
      { label: t('Add animation'), icon: Sparkles, onClick: () => { c.setAnim({ effect: 'fade' }); setTab('animations'); } },
      { label: t('Format …'), icon: PanelRight, onClick: () => setPane('format') },
      { label: t('Delete'), icon: Trash2, danger: true, hint: keyLabel('Del'), onClick: () => { change((d) => deleteEls(d, slide.id, sel)); setSel([]); } },
    ]);
  };
  const slideMenu = (e, id) => {
    if (!selSlides.includes(id)) { setSelSlides([id]); setCurId(id); }
    const ids = selSlides.includes(id) ? selSlides : [id];
    const s = doc.slides.find((x) => x.id === id);
    if (readOnly) { contextMenu(e, [{ label: t('Present from here'), icon: Presentation, onClick: () => { setCurId(id); setTimeout(() => startShow('current')); } }]); return; }
    contextMenu(e, [
      { label: t('New slide'), icon: Plus, hint: keyLabel('Ctrl+M'), onClick: () => newSlideCmd() },
      { label: t('Duplicate slide'), icon: CopyPlus, hint: keyLabel('Ctrl+D'), onClick: dupSlides },
      { label: t('Cut'), icon: Scissors, onClick: () => copySlides(true) },
      { label: t('Copy'), icon: Copy, onClick: () => copySlides() },
      CLIP && CLIP.type === 'slides' && { label: t('Paste slides'), icon: ClipboardPaste, onClick: pasteSlides },
      { label: s && s.hidden ? t('Show slide') : t('Hide slide'), icon: EyeOff, onClick: () => change((d) => updateSlides(d, ids, { hidden: !(s && s.hidden) })) },
      { label: t('Add section …'), icon: FolderPlus, onClick: async () => { const name = await promptDialog({ title: t('Add section'), label: t('Section name'), value: t('Untitled section'), confirm: t('Add') }); if (name) change((d) => addSection(d, id, name)); } },
      { label: t('Present from here'), icon: Presentation, onClick: () => { setCurId(id); setTimeout(() => startShow('current')); } },
      { label: ids.length > 1 ? t('Delete {n} slides', { n: ids.length }) : t('Delete slide'), icon: Trash2, danger: true, onClick: () => delSlides(ids) },
    ]);
  };
  const sectionMenu = (e, sec) => {
    if (readOnly) return;
    contextMenu(e, [
      { label: t('Rename section …'), icon: Pencil, onClick: async () => { const name = await promptDialog({ title: t('Rename section'), label: t('Section name'), value: sec.name, confirm: t('Rename') }); if (name) change((d) => renameSection(d, sec.id, name)); } },
      { label: t('Remove section'), icon: Minus, onClick: () => change((d) => removeSection(d, sec.id)) },
    ]);
  };

  /* ---------------- render */
  if (phase === 'loading') return <div className="of-slides-root of-slides-center"><Loading text={t('Opening presentation …')} /></div>;
  if (phase === 'error') {
    return (
      <div className="of-slides-root of-slides-center">
        <Empty icon={FileWarning} title={t('This presentation could not be opened')} text={`${t('The file may be damaged or in a format the built-in editor cannot read.')} (${error && error.message})`}
          action={host.download ? <Btn icon={Download} onClick={() => host.download(bytes, file.name)}>{t('Download')}</Btn> : null} />
      </div>
    );
  }
  const panes = pane && slide ? (
    pane === 'format' ? <FormatPane doc={doc} slide={slide} sel={sel} change={change} units={units} onClose={() => setPane(null)} onCrop={c.resetPicture} />
      : pane === 'comments' ? <CommentsPane doc={doc} slide={slide} change={change} author={gen.author} readOnly={readOnly} onClose={() => setPane(null)} />
        : pane === 'anim' ? <AnimPane doc={doc} slide={slide} sel={sel} setSel={setSel} change={change} readOnly={readOnly} onClose={() => setPane(null)} /> : null
  ) : null;
  const promptFn = (e) => t(e.prompt || '');
  const zoomPct = Math.round(zoom * 100);
  return (
    <div className={`of-slides-root ${readOnly ? 'ro' : ''}`} ref={rootRef} onKeyDown={onKeyDown} tabIndex={-1}>
      {!host.fidelity && fidelity.length > 0 && <div className="of-slides-callout"><Callout kind="info" title={t('Some parts can’t be edited here')}>{t('They are kept unchanged when you save: {list}.', { list: fidelity.join(', ') })}</Callout></div>}
      {!readOnly ? <Ribbon c={c} tab={tab} setTab={setTab} /> : (
        <div className="of-slides-robar">
          <span className="small muted grow">{t('Read only – you can browse, present, print and export.')}</span>
          <Btn className="sm" icon={Presentation} onClick={() => startShow('start')}>{t('From beginning')}</Btn>
          <Btn className="sm" onClick={() => startShow('current')}>{t('From current slide')}</Btn>
        </div>
      )}
      <div className="of-slides-body">
        {view === 'normal' && thumbs && (
          <Thumbs doc={doc} curId={curId} selSlides={selSlides} onSelect={selectSlide} onReorder={(ids, to) => !readOnly && change((d) => moveSlides(d, ids, to))} onContext={slideMenu} onAdd={() => !readOnly && newSlideCmd()} onSectionContext={sectionMenu} collapsed={collapsed} setCollapsed={setCollapsed} />
        )}
        <div className="of-slides-main">
          {view === 'normal' && slide && (
            <>
              <Canvas doc={doc} slide={slide} index={slideIdx} zoom={zoom} setZoom={(z, auto) => { setZoomS(z); if (!auto) setFit(false); }} fitMode={fit} sel={sel} setSel={setSel} change={change}
                editing={editing} setEditing={setEditing} tool={tool} setTool={setTool} createAt={createAt} readOnly={readOnly} guides={guides} grid={grid} snap={snap}
                textApi={textApi} onContext={elMenu} onPlaceholderPicture={async (el) => { const [f] = await pickFile('image/*'); if (f) insertPicture(await readFile(f), el); }}
                prompt={promptFn} spellcheck={spell} onSave={() => host.save && host.save()} onHistory={(w) => (w === 'undo' ? undo() : redo())} groupFocus={groupFocus} setGroupFocus={setGroupFocus}
                painter={painter} applyPainter={applyPainter} />
              {notesPane && (
                <div className="of-slides-notes" style={{ height: notesH }}>
                  <div className="of-slides-notesgrip" role="separator" aria-label={t('Resize notes')} onPointerDown={(e) => { const y0 = e.clientY; const h0 = notesH; const mv = (ev) => setNotesH(Math.max(48, Math.min(400, h0 - (ev.clientY - y0)))); const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); }; window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up); }} />
                  <textarea className="of-slides-notesta" value={slide.notes || ''} readOnly={readOnly} spellCheck={spell} placeholder={t('Click to add notes')} aria-label={t('Speaker notes')}
                    onKeyDown={(e) => { e.stopPropagation(); if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); if (host.save) host.save(); } }}
                    onChange={(e) => change((d) => updateSlide(d, slide.id, { notes: e.target.value }), `notes:${slide.id}`)} />
                </div>
              )}
            </>
          )}
          {view === 'sorter' && <Sorter doc={doc} curId={curId} selSlides={selSlides} onSelect={selectSlide} onOpen={(id) => { setCurId(id); setView('normal'); }} onReorder={(ids, to) => !readOnly && change((d) => moveSlides(d, ids, to))} onContext={slideMenu} zoom={Math.max(0.6, Math.min(2, zoom * 1.4))} />}
          {view === 'outline' && <Outline doc={doc} change={change} readOnly={readOnly} spellcheck={spell} onOpen={(id) => { setCurId(id); setView('normal'); }} />}
          {view === 'notes' && slide && <NotesPage doc={doc} slide={slide} index={slideIdx} change={change} readOnly={readOnly} spellcheck={spell} />}
        </div>
        {panes && <div className="of-slides-side">{panes}</div>}
      </div>
      <div className="of-slides-status">
        <span className="small">{t('Slide {n} of {total}', { n: slideIdx + 1, total: doc.slides.length })}</span>
        {curSection && h('span', { className: 'small muted' }, `· ${curSection.name}`)}
        {tool && <span className="small muted">· {t('Drag on the slide to draw – Esc cancels')}</span>}
        {painter && <span className="small muted">· {t('Click an object to apply the formatting')}</span>}
        <span className="grow" />
        <button type="button" className={`btn xs ghost ${notesPane ? 'active' : ''}`} onClick={() => setNotesPane((x) => !x)} title={t('Notes')}><StickyNote size={12} />{t('Notes')}</button>
        <button type="button" className={`btn xs ghost ${pane === 'comments' ? 'active' : ''}`} onClick={() => c.togglePane('comments')} title={t('Comments')}>{t('Comments')}{slide && (slide.comments || []).length ? ` (${slide.comments.length})` : ''}</button>
        <span className="of-slides-viewbtns">
          {[['normal', 'Normal'], ['sorter', 'Slide sorter'], ['outline', 'Outline view'], ['notes', 'Notes page']].map(([v, l]) => <button key={v} type="button" className={`btn xs ghost ${view === v ? 'active' : ''}`} onClick={() => c.setView(v)}>{t(l)}</button>)}
        </span>
        <button type="button" className="btn xs ghost" onClick={() => startShow('current')} title={t('Present from the current slide (Shift+F5)')}><Presentation size={12} /></button>
        <button type="button" className="btn xs ghost icon" title={t('Zoom out')} aria-label={t('Zoom out')} onClick={() => c.zoomBy(-1)}><Minus size={12} /></button>
        <input type="range" className="of-slides-zoom" min={10} max={400} value={zoomPct} aria-label={t('Zoom')} onChange={(e) => { setFit(false); setZoomS(+e.target.value / 100); }} />
        <button type="button" className="btn xs ghost icon" title={t('Zoom in')} aria-label={t('Zoom in')} onClick={() => c.zoomBy(1)}><Plus size={12} /></button>
        <button type="button" className={`btn xs ghost ${fit ? 'active' : ''}`} onClick={c.zoomFit} title={t('Fit to window')}>{zoomPct}%</button>
      </div>
      {present && (
        <Show doc={doc} start={present.start} presenter={present.preview ? false : present.presenter} rehearse={present.rehearse} loop={!present.preview && !!((doc.show && doc.show.loop) || ss.loop)} autoAdvance={present.preview ? 0 : (doc.show && doc.show.autoAdvance) || ss.autoAdvance || 0}
          openLink={(u) => (host.openLink ? host.openLink(u) : null)} openPresenterWindow={openPresenterWindow}
          onExit={() => { setPresent(null); setTimeout(() => rootRef.current && rootRef.current.focus(), 0); }}
          onTimings={async (tm) => {
            if (readOnly) return;
            const ok = await confirmDialog({ title: t('Keep the new slide timings?'), text: t('The slides then advance automatically in the slide show.'), confirm: t('Keep timings') });
            if (ok) change((d) => ({ ...d, slides: d.slides.map((s) => (tm[s.id] ? { ...s, transition: { type: 'none', dir: 'l', duration: 500, ...(s.transition || {}), advanceAfter: Math.round(tm[s.id] / 100) * 100 } } : s)) }));
          }} />
      )}
      {dialog && dialog.type === 'footer' && <FooterDialog footer={doc.footer} onClose={() => setDialog(null)} onApply={(f) => change((d) => ({ ...d, footer: f }))} />}
      {dialog && dialog.type === 'size' && <SizeDialog size={doc.size} units={units} onClose={() => setDialog(null)} onApply={(sz, how) => change((d) => setSize(d, sz, how))} />}
      {dialog && dialog.type === 'chart' && (() => { const el = slide && findEl(slide.elements, dialog.id); return el ? <ChartDataDialog chart={el.chart} onClose={() => setDialog(null)} onApply={(ch) => change((d) => updateEls(d, slide.id, [el.id], (x) => ({ ...x, chart: ch })))} /> : null; })()}
      {offscreen}
    </div>
  );
}
/** "Save as a copy" (PPTX_PATCH_FAILED): the name of the copy whose editor should save it as soon as it is ready */
const pendingCopy = { name: null };
const MIME_OF = { pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12', ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow', potx: 'application/vnd.openxmlformats-officedocument.presentationml.template' };
const splitName = (name, fallbackExt) => { const m = /^(.*?)(?:\.([^.]+))?$/.exec(String(name || '')); return { base: m[1] || String(name || ''), ext: fallbackExt || (m[2] || 'pptx').toLowerCase() }; };
const MACRO_EXT = (e) => ['pptm', 'ppsm', 'potm'].includes(String(e || '').toLowerCase());
export { newSlide, histReplace };
