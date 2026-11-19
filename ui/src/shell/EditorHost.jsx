/*
 * Editor host: a full-height view inside the Hostwerk page (#/ext/office/edit?path=…[&root=…][&app=…][&ro=1][&then=…])
 * or its own Hostwerk window (screen #/office-edit?…&window=1). Slim top bar (back, icon + name to rename, save state,
 * track changes, status line, Save, Versions, Share, Download as …, Print, more, close); below it the editor chunk
 * (ARCHITECTURE.md "Editor chunk contract") or the full editor (ONLYOFFICE Document Server) embedded in the page.
 *
 * The `host` object given to the chunk:
 *   setSaver(fn, { ext }?)  fn → Uint8Array | { bytes, ext } in file.ext (or file.converted.to / the nearest native format)
 *   setExporter({ formats, run }) · dirty(bool) · save() · convert(bytes, from, to) · hasConverter · download(bytes, name, mime)
 *   print(html | Blob) · status(text | { text, trackChanges }) · fidelity({ items: [string | { label }] }) · openLink(url)
 *   close() · lang
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowLeft, Save, History, Link2, Download, Printer, MoreHorizontal, X, Info, Monitor, Sparkles, ExternalLink, FolderOpen, LayoutTemplate,
  AlertTriangle, RotateCcw, CheckCircle2, Loader2, Cpu, FileWarning, GitCompare, AppWindow, Lock, CloudOff,
} from 'lucide-react';
import { get } from '@hostwerk/api';
import { Btn, Modal, Loading, Empty, Callout, toast, confirmDialog } from '@hostwerk/ui';
import { contextMenu, ErrorBoundary, useAiContext } from '@hostwerk/kit';
import { t } from '@hostwerk/i18n';
import { office, download, BASE } from '../lib/api';
import { loadEditor } from '../lib/loader';
import { FORMATS, KINDS, EXPORTS, extOf } from '../lib/formats';
import { openPlan, convertTargets, splitExt, basename, dirname, validateName, saveStateText, editRoute, nativeExtFor, toQuery } from './util';
import { FileIcon, Keep, fmtTime, uiLang, openLink, openHostwerkWindow } from './common';
import { loadSettings, loadStatus, hasConverterIn, settingsNow, changed } from './store';
import { EXT_ID, libRoute, openInDesktop, revealEntry, saveAsTemplate, filesRoute } from './ops';
import { ShareDialog, InfoDialog, VersionsSheet } from './dialogs';
import { IntegrationEditor } from './integrations';
import { ensureEngine, EngineDialogHost, FullEditorDialogHost, ensureFullEditor, openEngineDialog } from './engine';

const GOOGLE = { gdoc: { kind: 'docs', ext: 'docx' }, gsheet: { kind: 'sheets', ext: 'xlsx' }, gslides: { kind: 'slides', ext: 'pptx' } };
// Hostwerk's contextMenu shows labels as given – translate them here (t() leaves already translated text alone)
const menuAt = (btn, items) => { const r = btn.getBoundingClientRect(); contextMenu({ preventDefault() {}, stopPropagation() {}, clientX: Math.max(8, r.right - 240), clientY: r.bottom + 4 }, items.map((x) => (x && typeof x === 'object' && typeof x.label === 'string' ? { ...x, label: t(x.label) } : x))); };
const run = (fn) => () => { Promise.resolve().then(fn).catch((e) => toast.error(e)); };

/** Print an HTML string or a PDF blob through a hidden iframe in this window */
function printContent(content, holder) {
  const fr = document.createElement('iframe');
  fr.className = 'of-shell-print-frame';
  fr.setAttribute('aria-hidden', 'true');
  fr.title = 'print';
  let urlToFree = null;
  if (typeof content === 'string') fr.srcdoc = content;
  else { urlToFree = URL.createObjectURL(content instanceof Blob ? content : new Blob([content], { type: 'application/pdf' })); fr.src = urlToFree; }
  fr.onload = () => setTimeout(() => {
    try { fr.contentWindow.focus(); fr.contentWindow.print(); } catch (e) { toast.error(t('Printing did not start: {error}', { error: e.message })); }
    setTimeout(() => { fr.remove(); if (urlToFree) URL.revokeObjectURL(urlToFree); }, 60000);
  }, 350);
  (holder || document.body).appendChild(fr);
}

/**
 * A NEW file next to another one: POST /office/create in the library or a folder of Hostwerk's Files (first free name,
 * never overwrites). Returns the entry { path, name, root?, mtime }.
 */
async function createNextTo({ root, dir }, name, data) {
  // the backend picks the first free name in the folder (library or a root of Hostwerk's Files) – never overwrites
  return office.create(root ? { root, dir: dir || '' } : dir || '', name, data);
}

export default function EditorHost({ target, opts = {}, bare = false }) {
  const [file, setFile] = useState(() => ({ root: target.root || undefined, path: target.path, name: basename(target.path) }));
  const [phase, setPhase] = useState('loading'); // loading | ready | error | integration
  const [error, setError] = useState(null);
  const [chunk, setChunk] = useState(null);
  const [bytes, setBytes] = useState(null);
  const [plan, setPlan] = useState(null);
  const [readOnly, setReadOnly] = useState(!!opts.ro);
  const [settings, setSettings] = useState(null);
  const [hasConv, setHasConv] = useState(false);
  const [state, setState] = useState('idle');
  const [savedAt, setSavedAt] = useState(null);
  const [saveErr, setSaveErr] = useState('');
  const [statusLine, setStatusLine] = useState('');
  const [track, setTrack] = useState(null);
  const [fidelity, setFidelity] = useState(null);
  const [fidelityHidden, setFidelityHidden] = useState(false);
  const [notice, setNotice] = useState(null); // calm banner: { kind, text, action }
  const [exporter, setExporterState] = useState(null);
  const [dialog, setDialog] = useState(null);
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState('');
  const [lockInfo, setLockInfo] = useState(null);
  const [gen, setGen] = useState(0); // bump = reload the editor with fresh bytes
  const [conflict, setConflict] = useState(null);
  const [choice, setChoice] = useState(null); // ask how to save an imported/converted file: { resolve }
  const saver = useRef(null);
  const saverExt = useRef(null);
  const mtime = useRef(null);
  const dirtyRef = useRef(false);
  const editSeq = useRef(0);
  const saving = useRef(false);
  const again = useRef(false);
  const autoTimer = useRef(null);
  const choiceRef = useRef(null); // 'copy' | 'back' for this session
  // host.openAs(): bytes shown in an editor that are not a file yet – { google: driveId, ext } (saves back to Google
  // Drive) or { root?, dir, name } ("Save" creates a new file next to the original)
  const virtualRef = useRef(null);
  const [virtual, setVirtual] = useState(null);
  const [vgen, setVgen] = useState(0);
  const fileRef = useRef(file); fileRef.current = file;
  const planRef = useRef(plan); planRef.current = plan;
  const boxRef = useRef(null);
  // narrow window (Hostwerk ≤ 600 px shows its navigation above the page): bring the editor to the top of the view
  // and give it exactly the height of Hostwerk's scroll area (the page scrolls past Hostwerk's navigation)
  useEffect(() => {
    const el = boxRef.current;
    if (bare || !el || typeof matchMedia !== 'function') return undefined;
    const mq = matchMedia('(max-width: 600px)');
    const main = el.closest('.main');
    const fit = () => { el.style.height = mq.matches && main ? `${Math.max(480, main.clientHeight)}px` : ''; };
    fit();
    const ro = main && typeof ResizeObserver !== 'undefined' ? new ResizeObserver(fit) : null;
    if (ro) ro.observe(main);
    if (mq.addEventListener) mq.addEventListener('change', fit);
    const h = setTimeout(() => { if (mq.matches) el.scrollIntoView({ block: 'start' }); }, 60);
    return () => { clearTimeout(h); if (ro) ro.disconnect(); if (mq.removeEventListener) mq.removeEventListener('change', fit); el.style.height = ''; };
  }, [bare]);
  const thenDone = useRef(false);
  const ext = extOf(file.name);
  const fmt = FORMATS[ext] || null;
  const kind = (plan && plan.kind) || (fmt && fmt.kind) || 'viewer';
  const fromFiles = !!file.root;
  // address of this editor (screen #/office-edit or page #/ext/office/edit) – keeps "back" (folder in Files) and "os"
  const keep = { ...(opts.back ? { back: opts.back } : {}), ...(opts.os ? { os: 1 } : {}) };
  // opened by a click in Files (not from Explorer, not a fresh window): Back = one step back in the history
  const cameFrom = useRef(!opts.os && typeof history !== 'undefined' && history.length > 1);
  const routeOf = (f, extra = {}) => (bare ? `/office-edit?${toQuery({ ...(f.root ? { root: f.root } : {}), path: f.path, ...keep, ...extra })}` : editRoute(EXT_ID, f, { ...keep, ...extra }));
  // presenter console of the slides chunk (#/office-edit?…&present=<token>&window=1): no top bar, the chunk fills the window
  const presenting = !!opts.present;

  useAiContext(phase === 'ready' ? { kind: 'text', id: `office:${file.root || ''}:${file.path}`, label: file.name, text: t('Office file “{name}” is open in the editor{status}', { name: file.name, status: statusLine ? ` (${statusLine})` : '' }) } : null);

  /* ---------------- loading ---------------- */
  useEffect(() => {
    let alive = true;
    (async () => {
      setPhase('loading'); setError(null); setChunk(null); setExporterState(null); saver.current = null; setFidelity(null); setNotice(null);
      const s = await loadSettings();
      const st = await loadStatus().catch(() => null);
      const conv = hasConverterIn(st);
      if (!alive) return;
      setSettings(s); setHasConv(conv);
      if (!fmt) throw Object.assign(new Error(t('Office does not know the format “.{ext}”.', { ext })), { code: 'UNKNOWN' });
      if (opts.app === 'onlyoffice') { setPhase('integration'); return; }
      let p = openPlan(ext, { readOnly: !!opts.ro });
      let g = null;
      if (GOOGLE[ext]) {
        g = await get('/office/google/state').catch(() => null);
        if (g && g.connected) p = { mode: 'google', kind: GOOGLE[ext].kind, openExt: GOOGLE[ext].ext, saveExt: GOOGLE[ext].ext };
        else setNotice({ kind: 'info', text: t('Connect Google Drive under Office settings → Integrations to open and edit this Google file right here.'), action: { label: t('Connect Google Drive …'), href: `#${libRoute('/settings?group=integrations')}` } });
      }
      // a Google pointer without a connection: the viewer gets the pointer itself (the backend answers 409 otherwise)
      const r = await office.read(fileRef.current, { pointer: (FORMATS[ext] || {}).family === 'google' && p.mode !== 'google' });
      if (!alive) return;
      virtualRef.current = null; setVirtual(null);
      mtime.current = r.mtime;
      let data = r.bytes;
      let converted = null;
      if (p.mode === 'google') converted = { from: ext, to: p.openExt, google: true };
      if (p.mode === 'convert') {
        if (conv) {
          try { data = await office.convert(r.bytes, ext, p.openExt); converted = { from: ext, to: p.openExt }; } catch (e) { setNotice({ kind: 'warn', text: t('Converting with the Office engine failed ({error}) – the built-in reader shows the file instead.', { error: e.message }) }); p = { ...p, mode: 'import', openExt: ext }; }
        } else {
          // no engine yet: the chunk's built-in reader opens the original bytes; saving makes a native copy
          p = { ...p, mode: 'import', openExt: ext };
          ensureEngine(t('“{name}” is an older format. The Office engine opens it exactly and can save it back as .{ext}.', { name: fileRef.current.name, ext }));
          setNotice({ kind: 'info', text: t('Opened with the built-in reader. For a perfect result Office can set up its engine once (it asks you).'), action: { label: t('Set up the Office engine …'), onClick: () => openEngineDialog() } });
        }
      }
      let c = null;
      try { c = await loadEditor(p.kind); } catch (e) {
        if (p.kind !== 'viewer') {
          try { c = await loadEditor('viewer'); p = { ...p, mode: 'view' }; setNotice({ kind: 'warn', text: t('The editor for this kind of file is not available right now – the file is shown read-only.') }); } catch { c = null; }
        }
        if (!c) throw Object.assign(new Error(t('The editor could not be loaded: {error}', { error: e.message || String(e) })), { code: 'NO_CHUNK' });
      }
      if (!alive) return;
      if (!c || !c.Editor) throw Object.assign(new Error(t('The editor for this kind of file is not available yet.')), { code: 'NO_CHUNK' });
      setPlan(p);
      setReadOnly(p.mode === 'view' || !!opts.ro);
      setFile((f) => ({ ...f, ext, size: r.bytes.length, mtime: r.mtime, ...(converted ? { converted } : {}) }));
      setBytes(data); setChunk(c);
      setState(p.mode === 'view' || opts.ro ? 'readonly' : 'idle');
      dirtyRef.current = false;
      setPhase('ready');
    })().catch((e) => { if (alive) { setError(e); setPhase('error'); } });
    return () => { alive = false; };
  }, [file.path, file.root, gen, opts.app]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------------- soft lock ---------------- */
  useEffect(() => {
    if (phase !== 'ready' || readOnly || virtual) return undefined;
    const f = { root: file.root, path: file.path };
    const lock = () => office.lock(f).then((r) => { const l = r && (r.lock || r); if (l && l.other) setLockInfo(l); }).catch((e) => { if (e.status === 409 || e.status === 423) setLockInfo({ by: (e.lock && e.lock.by) || (e.data && e.data.lock && e.data.lock.by) || null, text: e.message }); });
    lock();
    const h = setInterval(lock, 120000);
    // window closed or reloaded: React cleanups do not run – release the lock with a request that survives the page
    const bye = () => { try { fetch(`${BASE}/unlock`, { method: 'POST', keepalive: true, credentials: 'same-origin', headers: { 'content-type': 'application/json' }, body: JSON.stringify(f) }); } catch { /* the lock expires on its own */ } };
    window.addEventListener('pagehide', bye);
    return () => { clearInterval(h); window.removeEventListener('pagehide', bye); office.unlock(f).catch(() => {}); };
  }, [phase, readOnly, file.path, file.root, virtual]);

  /* ---------------- saving ---------------- */
  const askChoice = () => new Promise((resolve) => setChoice({ resolve }));
  const writeTo = async (f, data, { force = false, auto = false } = {}) => {
    const r = await office.write(f, data, { mtime: force ? undefined : mtime.current, force, label: auto ? undefined : undefined });
    mtime.current = r.mtime || Date.now();
    return r;
  };
  const switchTo = (entry, newExt) => {
    const f = { root: fileRef.current.root, path: entry.path, name: entry.name || basename(entry.path) };
    setPlan((p) => ({ ...p, mode: 'edit', saveExt: newExt }));
    setFile((cur) => ({ ...cur, ...f, ext: newExt, converted: undefined }));
    try { history.replaceState(null, '', `#${routeOf(f)}`); } catch { /* ignore */ }
  };
  const saveNow = useCallback(async ({ auto = false, force = false, data: givenData = null } = {}) => {
    const p = planRef.current;
    if (!saver.current || !p || readOnly) return false;
    if (saving.current) { again.current = true; return false; }
    saving.current = true;
    const seq = editSeq.current;
    setState('saving'); setSaveErr('');
    let data = givenData;
    try {
      let outExt = saverExt.current || p.saveExt;
      if (!data) {
        const out = await saver.current();
        if (out && out.bytes) { data = out.bytes; outExt = out.ext || outExt; } else data = out;
      }
      if (!(data instanceof Uint8Array)) data = new Uint8Array(data);
      const f = fileRef.current;
      const origExt = extOf(f.name);
      const v = virtualRef.current;
      if (v && v.google) {
        await office.writeGoogle(v.google, data, outExt || v.ext);
      } else if (v) {
        // opened with host.openAs (e.g. iWork "Edit as docx"): the first explicit save creates the file – never autosave
        if (auto) { saving.current = false; setState('dirty'); return false; }
        const entry = await createNextTo({ root: v.root, dir: v.dir }, `${splitExt(v.name).base}.${outExt || extOf(v.name)}`, data);
        mtime.current = entry.mtime || Date.now();
        virtualRef.current = null; setVirtual(null);
        switchTo(entry, outExt || extOf(v.name));
        toast.success(t('Saved as “{name}” – the original stays unchanged', { name: entry.name }));
      } else if (outExt !== origExt && p.mode !== 'google') {
        let how = choiceRef.current;
        if (!how) {
          if (auto) { saving.current = false; setState('dirty'); return false; } // ask only when the user saves
          how = await askChoice();
          if (!how) { saving.current = false; setState('dirty'); return false; }
          choiceRef.current = how;
        }
        if (how === 'back') {
          let back;
          try { back = await office.convert(data, outExt, origExt); } catch (e) {
            choiceRef.current = null;
            const ready = await ensureEngine(t('Saving back as .{ext} needs the Office engine.', { ext: origExt }));
            if (!ready) throw new Error(t('Saving as .{ext} needs the Office engine – it is being set up. Save as .{native} meanwhile, or try again in a few minutes.', { ext: origExt, native: outExt }));
            back = await office.convert(data, outExt, origExt);
          }
          await writeTo(f, back, { force, auto });
        } else {
          const name = `${splitExt(f.name).base}.${outExt}`;
          const entry = await createNextTo({ root: f.root, dir: dirname(f.path) }, name, data);
          mtime.current = entry.mtime || Date.now();
          choiceRef.current = null;
          switchTo(entry, outExt);
          toast.success(t('Saved as “{name}” – the original stays unchanged', { name: entry.name || name }));
        }
      } else {
        await writeTo(f, data, { force, auto });
      }
      setSavedAt(Date.now());
      if (editSeq.current === seq) { dirtyRef.current = false; setState('saved'); } else setState('dirty');
      changed({ saved: f.path });
      saving.current = false;
      if (again.current) { again.current = false; if (dirtyRef.current) setTimeout(() => saveNow({ auto: true }), 50); }
      return true;
    } catch (e) {
      saving.current = false;
      if (e.status === 409) { setState('conflict'); setConflict({ data, mtime: e.mtime }); return false; }
      setState('error'); setSaveErr(e.message || String(e));
      if (!auto) toast.error(e);
      return false;
    }
  }, [readOnly]); // eslint-disable-line react-hooks/exhaustive-deps

  const markDirty = useCallback((on) => {
    if (readOnly) return;
    if (on) {
      editSeq.current++;
      dirtyRef.current = true;
      if (!saving.current) setState('dirty');
      const g = settingsNow().general;
      clearTimeout(autoTimer.current);
      if (g.autosave) autoTimer.current = setTimeout(() => { if (dirtyRef.current) saveNow({ auto: true }); }, Math.max(5, Number(g.autosaveSeconds) || 20) * 1000);
    } else { dirtyRef.current = false; setState((s) => (s === 'dirty' ? 'saved' : s)); }
  }, [readOnly, saveNow]);
  useEffect(() => () => clearTimeout(autoTimer.current), []);

  /* ---------------- the host object ---------------- */
  const hostRef = useRef(null);
  const exporterRef = useRef(null);
  const exportBytes = async (toExt) => {
    const p = planRef.current;
    const ex = exporterRef.current;
    if (ex && ex.formats.some((f) => f.ext === toExt)) return ex.run(toExt);
    if (!saver.current) throw new Error(t('Nothing to export yet.'));
    const out = await saver.current();
    const data = out && out.bytes ? out.bytes : out;
    const fromExt = (out && out.ext) || saverExt.current || p.saveExt;
    if (fromExt === toExt) return data;
    try { return await office.convert(data, fromExt, toExt); } catch (e) {
      await ensureEngine(t('Exporting as .{ext} needs the Office engine.', { ext: toExt }));
      throw e;
    }
  };
  const doClose = async () => {
    if (dirtyRef.current && !(await leaveGuard())) return;
    leave();
  };
  /**
   * Back: own Hostwerk window → close it; opened from Files (files.open, "back" = #/files?root=…&path=…&select=…) →
   * that folder (history when we came from there, else the address – a file from Explorer has no history here);
   * otherwise the previous page or the Office folder in Files.
   */
  const leave = () => {
    if (opts.window) { try { window.close(); return; } catch { /* not a window */ } }
    const back = typeof opts.back === 'string' && /^#?\/files(\?|$)/.test(opts.back) ? `#${opts.back.replace(/^#/, '')}` : null;
    if (back) { if (!opts.os && cameFrom.current) history.back(); else location.replace(back); return; }
    if (opts.back === 'history' || (fromFiles && history.length > 1)) { history.back(); return; }
    location.replace(filesRoute(fileRef.current));
  };
  /**
   * host.openAs({ bytes, ext, name, google? }) – show bytes in the matching editor as a file that is not saved yet
   * (viewer: Google "Open in Office", iWork "Edit as docx/pptx/xlsx"). google = { id }: saving uploads back to Google
   * Drive; otherwise the first "Save" creates "<name>.<ext>" next to the original (library folder or file-browser folder).
   */
  const openAs = async ({ bytes: data, ext: toExt, name, google } = {}) => {
    const x = String(toExt || '').toLowerCase();
    const fm = FORMATS[x];
    if (!fm || !(data instanceof Uint8Array || data instanceof ArrayBuffer)) throw new Error(t('Office does not know the format “.{ext}”.', { ext: x }));
    if (dirtyRef.current && !(await leaveGuard())) return false;
    const cur = fileRef.current;
    const origExt = extOf(cur.name);
    // the pointer file itself (.gdoc …) while Google is connected now: the host opens it the normal way (saves to Google)
    if (google && GOOGLE[origExt] && GOOGLE[origExt].ext === x) { setGen((n) => n + 1); return true; }
    const c = await loadEditor(fm.kind);
    if (!c || !c.Editor) throw new Error(t('The editor for this kind of file is not available yet.'));
    const nm = name || `${splitExt(cur.name).base}.${x}`;
    const v = google && google.id ? { google: String(google.id), ext: x, name: nm } : { root: cur.root, dir: dirname(cur.path), name: nm };
    virtualRef.current = v; setVirtual(v);
    saver.current = null; saverExt.current = null; exporterRef.current = null; setExporterState(null); setFidelity(null);
    choiceRef.current = null;
    const p = { mode: v.google ? 'google' : 'edit', kind: fm.kind, openExt: x, saveExt: x };
    setPlan(p); planRef.current = p;
    setReadOnly(fm.edit === 'view');
    setFile((f) => ({ ...f, name: nm, ext: x, size: data.byteLength, converted: v.google ? { from: origExt, to: x, google: true } : undefined }));
    setBytes(data instanceof Uint8Array ? data : new Uint8Array(data)); setChunk(c);
    setNotice(v.google ? null : { kind: 'info', text: t('Not saved yet – “Save” creates “{name}” next to the original, which stays unchanged.', { name: nm }) });
    dirtyRef.current = !v.google;
    setState(fm.edit === 'view' ? 'readonly' : v.google ? 'idle' : 'dirty');
    setVgen((n) => n + 1);
    return true;
  };
  /**
   * host.saveCopy(bytes, ext, name) – save bytes as a NEW file next to the original (first free name, never overwrites)
   * and go on editing that copy in this editor; the original stays unchanged (e.g. slides' PPTX_PATCH_FAILED). → entry
   */
  const saveCopy = async (data, toExt, name) => {
    const cur = fileRef.current;
    const x = String(toExt || extOf(cur.name)).toLowerCase();
    if (!FORMATS[x]) throw new Error(t('Office does not know the format “.{ext}”.', { ext: x }));
    const raw = data instanceof Uint8Array ? data : new Uint8Array(data);
    const nm = name ? (extOf(name) === x ? name : `${name}.${x}`) : `${t('{name} (copy)', { name: splitExt(cur.name).base })}.${x}`;
    const entry = await createNextTo({ root: cur.root, dir: dirname(cur.path) }, nm, raw);
    mtime.current = entry.mtime || Date.now();
    dirtyRef.current = false; choiceRef.current = null;
    setSavedAt(Date.now()); setState('saved');
    switchTo(entry, x);
    changed({ created: entry.path });
    toast.success(t('Saved as “{name}” – the original stays unchanged', { name: entry.name || nm }));
    return entry;
  };
  if (!hostRef.current) hostRef.current = {};
  Object.assign(hostRef.current, {
    setSaver: (fn, o = {}) => { saver.current = typeof fn === 'function' ? fn : null; saverExt.current = o && o.ext ? o.ext : null; },
    setExporter: (x) => { exporterRef.current = x && Array.isArray(x.formats) && typeof x.run === 'function' ? x : null; setExporterState(exporterRef.current); },
    dirty: (on) => markDirty(!!on),
    save: () => saveNow(),
    convert: async (data, from, to) => {
      try { return await office.convert(data, from, to); } catch (e) {
        if (e.status === 501) ensureEngine(t('Converting from .{from} to .{to} needs the Office engine.', { from, to }));
        throw e;
      }
    },
    hasConverter: hasConv,
    download: (data, name, mime) => download(data, name, mime),
    print: (content) => printContent(content, boxRef.current),
    status: (x) => {
      if (x && typeof x === 'object') { setStatusLine(String(x.text || '')); if ('trackChanges' in x) setTrack(x.trackChanges ? { on: true, count: x.changes ?? x.trackCount ?? null } : { on: false, count: x.changes ?? null }); }
      else setStatusLine(String(x || ''));
    },
    fidelity: (x) => { const items = (x && Array.isArray(x.items) ? x.items : []).map((i) => (typeof i === 'string' ? i : i && (i.label || i.text))).filter(Boolean); setFidelity(items.length ? { items, kept: x.kept !== false } : null); setFidelityHidden(false); },
    openLink: (u) => openLink(u),
    openAs: (x) => openAs(x).catch((e) => { toast.error(e); throw e; }),
    saveCopy: (data, toExt, name) => saveCopy(data, toExt, name).catch((e) => { toast.error(e); throw e; }),
    /** the Office engine dialog; when it is ready the file opens again (now through the engine) → ready */
    ensureEngine: async (reason) => { const ok = await openEngineDialog(reason || ''); if (ok) { await loadStatus(true).catch(() => {}); setGen((n) => n + 1); } return !!ok; },
    close: () => doClose(),
    lang: uiLang(),
  });

  // messages of the full editor's frame: { event: 'ready' | 'modified' | 'saved' | 'error' | 'close' }
  const onFrameEvent = useCallback((d) => {
    if (d.event === 'modified') setState('dirty'); // ONLYOFFICE saves by itself (callback to the backend)
    if (d.event === 'saved') { setState('saved'); setSavedAt(Date.now()); changed({ saved: fileRef.current.path }); }
    if (d.event === 'close') {
      leave();
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------------- leave guard ---------------- */
  const leaveRef = useRef(null);
  const leaveGuard = () => new Promise((resolve) => { leaveRef.current = resolve; setDialog({ type: 'leave' }); });
  useEffect(() => {
    const before = (e) => { if (dirtyRef.current) { e.preventDefault(); e.returnValue = ''; } };
    // links elsewhere in Hostwerk (sidebar, search results …) while there are unsaved changes: ask first
    const click = async (e) => {
      if (!dirtyRef.current || e.defaultPrevented || e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey) return;
      const a = e.target.closest && e.target.closest('a[href^="#"]');
      if (!a || (boxRef.current && boxRef.current.contains(a))) return;
      e.preventDefault(); e.stopPropagation();
      const href = a.getAttribute('href');
      if (await leaveGuard()) location.hash = href.replace(/^#/, '');
    };
    window.addEventListener('beforeunload', before);
    document.addEventListener('click', click, true);
    return () => {
      window.removeEventListener('beforeunload', before);
      document.removeEventListener('click', click, true);
      // left another way (shortcut, back button) – keep the work: save in the background
      if (dirtyRef.current && saver.current) { const name = fileRef.current.name; saveNow({ auto: false }).then((ok) => { if (ok) toast.success(t('Changes to “{name}” saved', { name })); }); }
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  // Ctrl+S in the top bar (the chunk handles it inside the editor area)
  const onKey = (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && !e.altKey) { e.preventDefault(); e.stopPropagation(); if (!readOnly) saveNow(); }
    // Ctrl+R inside the editor area is the editors' own key (align right / fill right, as in Word and Excel). Where an
    // editor does not use it (e.g. while typing in a cell), Hostwerk's "reload the interface" must not throw away
    // unsaved changes: with changes it is swallowed here (Hostwerk's listener sits on window, Electron's menu
    // accelerator only fires for keys the page did not handle).
    if (e.ctrlKey && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'r' && dirtyRef.current) { e.preventDefault(); e.stopPropagation(); }
  };
  // The same keys when nothing has the focus (a ribbon button that closed, a click on an area that prevents focus):
  // the key then goes to <body>, past the editor – Ctrl+S must still save and Ctrl+R must not reload over changes
  const onKeyRef = useRef(onKey);
  onKeyRef.current = onKey;
  useEffect(() => {
    const k = (e) => { const a = document.activeElement; if (!e.defaultPrevented && (!a || a === document.body || a === document.documentElement)) onKeyRef.current(e); };
    window.addEventListener('keydown', k, true);
    return () => window.removeEventListener('keydown', k, true);
  }, []);

  /* ---------------- "then": convert after opening (library "Convert to …" without converter) ---------------- */
  useEffect(() => {
    const th = String(opts.then || '');
    if (thenDone.current || phase !== 'ready' || !th.startsWith('convert:')) return undefined;
    const to = th.slice(8);
    const ready = exporter && exporter.formats.some((f) => f.ext === to);
    const timer = setTimeout(async () => {
      if (thenDone.current) return;
      thenDone.current = true;
      try {
        const out = await exportBytes(to);
        if (out == null) return; // handled by the chunk (print dialog) – nothing to store
        const f = fileRef.current;
        const name = `${splitExt(f.name).base}.${to}`;
        const created = await createNextTo({ root: f.root, dir: dirname(f.path) }, name, out);
        changed();
        toast(t('Converted – saved as “{name}”', { name: created.name || name }), 'success', 9000, { action: { label: t('Open'), onClick: () => { location.hash = routeOf(created); } } });
      } catch (e) { toast.error(e); }
    }, ready ? 50 : 2500);
    return () => clearTimeout(timer);
  }, [phase, exporter]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------------- actions ---------------- */
  const rename = async () => {
    const next = nameDraft.trim();
    setRenaming(false);
    if (!next || next === file.name) return;
    const m = validateName(next); if (m) { toast.error(t(m)); return; }
    const withExt = extOf(next) ? next : `${next}.${ext}`;
    try {
      const e = await office.rename(file.root ? { root: file.root, path: file.path } : file.path, withExt);
      const f = { root: (e && e.root) || undefined, path: (e && e.path) || [dirname(file.path), withExt].filter(Boolean).join('/'), name: (e && e.name) || withExt };
      setFile((cur) => ({ ...cur, ...f }));
      try { history.replaceState(null, '', `#${routeOf(f)}`); } catch { /* ignore */ }
      changed();
      toast.success(t('Renamed to “{name}”', { name: f.name }));
    } catch (e) { toast.error(e); }
  };
  const downloadAs = async (toExt) => {
    const out = await exportBytes(toExt);
    if (out == null) return; // the chunk handled it itself (e.g. the print dialog for "PDF" without the engine)
    download(out, `${splitExt(file.name).base}.${toExt}`, (FORMATS[toExt] || {}).mime);
  };
  const clientFormats = exporter ? exporter.formats.map((f) => f.ext) : [];
  const targets = useMemo(() => {
    const own = ext;
    const list = convertTargets(fmt && fmt.kind !== 'viewer' ? (plan && plan.saveExt) || own : own, { hasConverter: true, client: clientFormats.length ? clientFormats : (EXPORTS[kind] || {}).client || [] });
    if (!list.some((x) => x.ext === own) && own && fmt && fmt.kind !== 'viewer') list.unshift({ ext: own, label: fmt.label, via: 'client' });
    return list;
  }, [exporter, plan, ext]); // eslint-disable-line react-hooks/exhaustive-deps
  const downloadMenu = (btn) => menuAt(btn, [
    !virtual && { label: t('Saved file (.{ext})', { ext }), icon: Download, onClick: run(async () => { const r = await office.read(file); download(r.bytes, file.name); }) },
    !virtual && '-',
    ...targets.map((x) => ({ label: `${t(x.label)} (.${x.ext})`, icon: Download, hint: x.via === 'server' && !hasConv ? t('engine') : undefined, onClick: run(() => downloadAs(x.ext)) })),
  ].filter(Boolean));
  const print = async () => {
    try {
      const canPdf = (exporter && exporter.formats.some((f) => f.ext === 'pdf')) || ext === 'pdf';
      if (ext === 'pdf' && !dirtyRef.current) { const r = await office.read(file); printContent(new Blob([r.bytes], { type: 'application/pdf' }), boxRef.current); return; }
      if (!canPdf && !hasConv) { toast(t('Use the print button of the editor (Ctrl+P) – or set up the Office engine for print-ready PDFs.')); return; }
      const out = await exportBytes('pdf');
      if (out == null) return; // the chunk opened its own print dialog
      printContent(new Blob([out], { type: 'application/pdf' }), boxRef.current);
    } catch (e) { toast.error(e); }
  };
  const openFull = async () => {
    if (dirtyRef.current) { const ok = await saveNow(); if (!ok && dirtyRef.current) return; }
    const app = await ensureFullEditor();
    if (app) location.hash = routeOf(file, { app });
  };
  const openWindow = async () => {
    if (dirtyRef.current) { const ok = await saveNow(); if (!ok) return; }
    openHostwerkWindow(`/office-edit?${editRoute(EXT_ID, file).split('?')[1]}`);
  };
  const moreMenu = (btn) => menuAt(btn, [
    !virtual && { label: 'Open in the full editor', icon: Sparkles, onClick: run(openFull) },
    !virtual && { label: 'Open in the desktop app', icon: Monitor, onClick: run(async () => { if (dirtyRef.current) await saveNow(); await openInDesktop(file); }) },
    !bare && !virtual && { label: 'Open in its own window', icon: AppWindow, onClick: run(openWindow) },
    !virtual && '-',
    !virtual && { label: 'Info', icon: Info, onClick: () => setDialog({ type: 'info' }) },
    !virtual && { label: 'Save as template …', icon: LayoutTemplate, onClick: run(() => saveAsTemplate(file)) },
    { label: 'Show in Explorer', icon: FolderOpen, onClick: run(() => revealEntry(file)) },
    phase === 'integration' && { label: 'Back to the built-in editor', icon: RotateCcw, onClick: () => { location.hash = routeOf(file); } },
  ].filter(Boolean));

  /* ---------------- render ---------------- */
  const ss = saveStateText(phase === 'integration' ? (state === 'dirty' ? 'dirty' : 'external') : state, { savedAt, error: saveErr, app: 'ONLYOFFICE' });
  const saveIcon = state === 'saving' ? Loader2 : state === 'saved' ? CheckCircle2 : state === 'error' || state === 'conflict' ? AlertTriangle : state === 'readonly' ? Lock : null;
  const SI = saveIcon;
  const top = (
    <header className="of-shell-ebar" role="toolbar" aria-label={t('Document')}>
      <button type="button" className="btn ghost icon sm" title={t(opts.back ? 'Back to the folder' : 'Back')} aria-label={t(opts.back ? 'Back to the folder' : 'Back')} onClick={() => doClose()}><ArrowLeft size={16} /></button>
      <FileIcon entry={{ name: file.name }} />
      {renaming ? (
        <input className="input of-shell-ename-input" autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={rename} onKeyDown={(e) => { if (e.key === 'Enter') rename(); if (e.key === 'Escape') setRenaming(false); }} aria-label={t('File name')} />
      ) : (
        <button type="button" className="of-shell-ename" disabled={readOnly || !!virtual} title={virtual ? t('Not saved yet') : t('Rename')} onClick={() => { setNameDraft(file.name); setRenaming(true); }}><Keep>{file.name}</Keep></button>
      )}
      <span className={`of-shell-save c-${ss.tone}`} role="status" aria-live="polite">{SI && <SI size={13} className={state === 'saving' ? 'spin' : ''} aria-hidden="true" />}{t(ss.text, ss.vars)}</span>
      {track && track.on && <span className="badge warn" title={t('Changes are recorded – accept or reject them in the Review tab')}><GitCompare size={11} aria-hidden="true" />{track.count != null ? t('Track changes · {n}', { n: track.count }) : t('Track changes')}</span>}
      {file.converted && !file.converted.google && <span className="badge" title={t('Opened via the Office engine from .{from}', { from: file.converted.from })}>{`.${file.converted.from} → .${file.converted.to}`}</span>}
      {file.converted && file.converted.google && <span className="badge accent">Google Drive</span>}
      {statusLine && <span className="of-shell-estatus tiny muted ellipsis" title={statusLine}>{statusLine}</span>}
      <span className="grow" />
      {phase !== 'integration' && !readOnly && <Btn className="sm" icon={Save} disabled={!chunk || state === 'saving'} title={t('Save (Ctrl+S)')} onClick={() => saveNow()}>Save</Btn>}
      {!virtual && <Btn className="sm ghost icon" icon={History} title="Versions" aria-label="Versions" onClick={() => setDialog({ type: 'versions' })} />}
      {!virtual && settingsNow().sharing.enabled && <Btn className="sm ghost icon" icon={Link2} title="Share link …" aria-label="Share link …" onClick={() => setDialog({ type: 'share' })} />}
      <button type="button" className="btn sm ghost icon" title={t('Download as …')} aria-label={t('Download as …')} onClick={(e) => downloadMenu(e.currentTarget)}><Download size={15} /></button>
      <Btn className="sm ghost icon" icon={Printer} title="Print" aria-label="Print" onClick={print} />
      <button type="button" className="btn sm ghost icon" title={t('More')} aria-label={t('More')} onClick={(e) => moreMenu(e.currentTarget)}><MoreHorizontal size={15} /></button>
      <button type="button" className="btn sm ghost icon" title={t('Close')} aria-label={t('Close')} onClick={() => doClose()}><X size={16} /></button>
    </header>
  );
  const banners = (
    <>
      {lockInfo && <div className="of-shell-banner is-warn"><Lock size={14} aria-hidden="true" /><span className="small grow">{lockInfo.by || !lockInfo.text ? t('This file is also open elsewhere ({by}). Saving may overwrite the other changes – Office asks if it notices.', { by: lockInfo.by || '?' }) : `${lockInfo.text} ${t('Saving may overwrite the other changes – Office asks if it notices.')}`}</span><button type="button" className="btn xs ghost icon" aria-label={t('Hide')} title={t('Hide')} onClick={() => setLockInfo(null)}><X size={12} /></button></div>}
      {notice && <div className={`of-shell-banner is-${notice.kind}`}><Info size={14} aria-hidden="true" /><span className="small grow">{notice.text}</span>{notice.action && (notice.action.href ? <a className="btn xs" href={notice.action.href}>{notice.action.label}</a> : <button type="button" className="btn xs" onClick={notice.action.onClick}>{notice.action.label}</button>)}<button type="button" className="btn xs ghost icon" aria-label={t('Hide')} title={t('Hide')} onClick={() => setNotice(null)}><X size={12} /></button></div>}
      {fidelity && !fidelityHidden && <div className="of-shell-banner is-info"><FileWarning size={14} aria-hidden="true" /><span className="small grow">{t(fidelity.kept ? 'Some parts can’t be edited here and are kept unchanged: {list}.' : 'Some parts can’t be edited here: {list}.', { list: fidelity.items.slice(0, 6).join(', ') })}</span><button type="button" className="btn xs" onClick={run(openFull)}><Sparkles size={12} />{t('Open in the full editor')}</button><button type="button" className="btn xs ghost icon" aria-label={t('Hide')} title={t('Hide')} onClick={() => setFidelityHidden(true)}><X size={12} /></button></div>}
    </>
  );
  let body;
  if (phase === 'loading') body = <Loading text={t('Opening “{name}” …', { name: file.name })} />;
  else if (phase === 'error') body = <OpenError error={error} file={file} onRetry={() => setGen((n) => n + 1)} onReadOnly={() => { location.hash = editRoute(EXT_ID, file, { ro: 1 }); setGen((n) => n + 1); }} onFull={openFull} />;
  else if (phase === 'integration') body = <IntegrationEditor file={file} app={opts.app} onEvent={onFrameEvent} />;
  else {
    const E = chunk.Editor;
    body = (
      <ErrorBoundary compact where={`office editor ${kind}`} resetKey={gen}>
        <E key={`${file.root || ''}|${file.path}|${gen}|${vgen}`} file={file} bytes={bytes} readOnly={readOnly} settings={settings} host={hostRef.current} />
      </ErrorBoundary>
    );
  }
  return (
    <div className={`of-shell-editor ${bare ? 'is-bare' : ''} ${presenting ? 'is-present' : ''}`} ref={boxRef} onKeyDown={onKey}>
      {!presenting && top}
      {!presenting && banners}
      <div className="of-shell-ebody">{body}</div>
      {dialog && dialog.type === 'versions' && <VersionsSheet file={file} onClose={() => setDialog(null)} beforeRestore={async () => { if (dirtyRef.current) { const ok = await confirmDialog({ title: 'Discard unsaved changes?', text: 'Restoring a version replaces what you see now. Your unsaved changes would be lost – save first to keep them as a version.', confirm: 'Discard and restore', danger: true }); if (!ok) return false; dirtyRef.current = false; } return true; }} onRestored={() => { setDialog(null); dirtyRef.current = false; setGen((n) => n + 1); }} />}
      {dialog && dialog.type === 'share' && <ShareDialog entry={file} onClose={() => setDialog(null)} />}
      {dialog && dialog.type === 'info' && <InfoDialog entry={file} onClose={() => setDialog(null)} />}
      {dialog && dialog.type === 'leave' && <LeaveDialog name={file.name} onSave={async () => { const ok = await saveNow(); if (ok) { setDialog(null); leaveRef.current && leaveRef.current(true); } }} onDiscard={() => { dirtyRef.current = false; setDialog(null); leaveRef.current && leaveRef.current(true); }} onCancel={() => { setDialog(null); leaveRef.current && leaveRef.current(false); }} />}
      {conflict && <ConflictDialog name={file.name} onOverwrite={async () => { const c = conflict; setConflict(null); await saveNow({ force: true, data: c.data }); }}
        onCopy={async () => { const c = conflict; setConflict(null); const { base, ext: e2 } = splitExt(file.name); const name = `${base} (${t('my version')}).${e2}`; if (file.root) { await office.write({ root: file.root, path: [dirname(file.path), name].filter(Boolean).join('/') }, c.data, { force: true }); } else await office.create(dirname(file.path), name, c.data); changed(); dirtyRef.current = false; setState('saved'); toast.success(t('Your version is saved as “{name}”', { name })); }}
        onReload={() => { setConflict(null); dirtyRef.current = false; setGen((n) => n + 1); }} onCancel={() => { setConflict(null); setState('dirty'); }} />}
      {choice && <SaveChoiceDialog file={file} outExt={saverExt.current || (plan && plan.saveExt) || nativeExtFor(ext)} hasConverter={hasConv} onPick={(v) => { const c = choice; setChoice(null); c.resolve(v); }} />}
      <EngineDialogHost />
      <FullEditorDialogHost />
    </div>
  );
}

function OpenError({ error, file, onRetry, onReadOnly, onFull }) {
  const msg = (error && error.message) || String(error || '');
  const noChunk = error && error.code === 'NO_CHUNK';
  return (
    <Empty icon={error && error.status === 404 ? CloudOff : AlertTriangle} title={error && error.status === 404 ? 'The file is not there any more' : noChunk ? 'The editor could not be loaded' : 'The file could not be opened'} text={msg}
      action={<div className="row-tight of-shell-wrap of-shell-center">
        <Btn className="primary" icon={RotateCcw} onClick={onRetry}>Try again</Btn>
        {noChunk && <Btn icon={Lock} onClick={onReadOnly}>Open read-only</Btn>}
        <Btn icon={Sparkles} onClick={onFull}>Open in the full editor</Btn>
        <Btn className="ghost" icon={Monitor} onClick={() => openInDesktop(file)}>Open in the desktop app</Btn>
        <Btn className="ghost" icon={Download} onClick={() => office.read(file).then((r) => download(r.bytes, file.name))}>Download</Btn>
      </div>} />
  );
}

function LeaveDialog({ name, onSave, onDiscard, onCancel }) {
  return (
    <Modal size="dialog" title={t('Save the changes to “{name}”?', { name })} onClose={onCancel}
      footer={<><Btn className="ghost" onClick={onCancel}>Cancel</Btn><Btn className="danger" onClick={onDiscard}>Don’t save</Btn><Btn className="primary" icon={Save} onClick={onSave}>Save</Btn></>}>
      <p className="small of-shell-p">Your changes are lost if you leave without saving.</p>
    </Modal>
  );
}
function ConflictDialog({ name, onOverwrite, onCopy, onReload, onCancel }) {
  return (
    <Modal title={t('“{name}” was changed elsewhere', { name })} subtitle="Someone or another program saved the file after you opened it" onClose={onCancel}
      footer={<><Btn className="ghost" onClick={onCancel}>Cancel</Btn><Btn onClick={onReload} icon={RotateCcw}>Load theirs</Btn><Btn onClick={onCopy}>Save mine as a copy</Btn><Btn className="primary" onClick={onOverwrite}>Overwrite with mine</Btn></>}>
      <div className="stack">
        <p className="small of-shell-p"><b>Overwrite with mine</b> – your version replaces theirs (theirs stays under Versions).</p>
        <p className="small of-shell-p"><b>Save mine as a copy</b> – both stay: yours as a new file next to it.</p>
        <p className="small of-shell-p"><b>Load theirs</b> – you see the other version; your unsaved changes are dropped.</p>
      </div>
    </Modal>
  );
}
function SaveChoiceDialog({ file, outExt, hasConverter, onPick }) {
  const ext = extOf(file.name);
  const base = splitExt(file.name).base;
  return (
    <Modal title={t('How should “{name}” be saved?', { name: file.name })} onClose={() => onPick(null)} footer={<Btn onClick={() => onPick(null)}>Cancel</Btn>}>
      <div className="stack">
        <button type="button" className="of-shell-choice is-btn" onClick={() => onPick('copy')}>
          <FileIcon name={`x.${outExt}`} size="md" />
          <span className="stack of-shell-gap-1 grow"><b>{t('Save as “{name}”', { name: `${base}.${outExt}` })}</b><span className="small muted">{t('Recommended – a new file in the modern format, the original .{ext} stays as it is.', { ext })}</span></span>
        </button>
        <button type="button" className="of-shell-choice is-btn" onClick={() => onPick('back')}>
          <FileIcon name={file.name} size="md" />
          <span className="stack of-shell-gap-1 grow"><b>{t('Keep .{ext}', { ext })}</b><span className="small muted">{hasConverter ? t('The Office engine converts it back into the original format.') : t('Needs the Office engine – Office sets it up once (it asks you).')}</span></span>
          {!hasConverter && <Cpu size={16} className="muted" aria-hidden="true" />}
        </button>
      </div>
    </Modal>
  );
}
