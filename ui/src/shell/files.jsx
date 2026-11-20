/*
 * Office inside Hostwerk's Files (SDK 3.15) – no own page in the sidebar:
 *   files.open    click / Enter on an office file in ANY folder opens the editor in front (screen #/office-edit, Back =
 *                 the folder), also for files from Explorer ("Open with" → from: 'os')
 *   files.new     "New" menu: Document, Spreadsheet, Presentation, PDF (default format, created and opened at once),
 *                 "From template or other format …" (the New dialog)
 *   files.actions right-click: Open in Office / full editor, Convert, Share link, Versions, Favourite, Tags, Template, Info
 *   files.marks   star (favourite), link (shared), lock (open in an editor) behind the name
 *   files.folder  the "Office" section in the Office documents root: new buttons, recent files, favourites, shared,
 *                 tags, templates, Office trash and full-text search
 * Dialogs open in Office's dialog host – mounted once in Hostwerk's sidebar footer (renders nothing itself, its Modals go
 * to .app), so they work in every folder; a screen fallback covers a missing host.
 */
import { memo, useCallback, useEffect, useRef, useState } from 'react';
import {
  FileText, Sheet, Presentation, FileType, LayoutTemplate, Pencil, Sparkles, FileOutput, Link2, History, Star, Tag, Info, Monitor, Clock, Share2,
  Trash2, RotateCcw, Search, Settings2, Lock, FilePlus2, X, Plus, Upload, FolderOpen, ChevronLeft, ChevronRight,
} from 'lucide-react';
import { Btn, Empty } from '@hostwerk/ui';
import * as HAPI from '@hostwerk/api';
import { contextMenu } from '@hostwerk/kit';
import { t, pl } from '@hostwerk/i18n';
import { office, isMissing } from '../lib/api';
import { toast } from '../lib/toast';
import { FORMATS, OFFICE_RE, TEMPLATES, KINDS, extOf } from '../lib/formats';
import { KIND_COLOR } from '../lib/fileTypes';
import { prefetchEditor, prefetchFor, warmEditor, kindFor } from '../lib/loader';
import { keyLabel } from '../editors/_shared/keys.js';
import { FileIcon, Thumb, Keep, Raw, Highlight, fmtWhen, useDebounced } from './common';
import { settingsNow, useChanged, useOfficeStatus, loadSettings, changed } from './store';
import { convertTargets, defaultExt, defaultBaseName, dirname, basename, toQuery } from './util';
import { LIB_ROOT, canon, openEditor, filesRoute, createFile, newFromOwnTemplate, convertEntry, openInDesktop, toggleFavorite, saveAsTemplate, restoreTrash, purgeTrash, emptyTrash, uploadFiles } from './ops';
import { NewDialog, ShareDialog, InfoDialog, TagsDialog, VersionsSheet, ConvertDialog, TemplateCard } from './dialogs';
import { EngineDialogHost, FullEditorDialogHost, ensureFullEditor } from './engine';
import { DefaultAppHint, checkOsNotice } from './osCard';
import { FeatureHost } from './features';
import { featureItems } from './fileFeatures';
import ThumbRenderer from './thumbHost';

const KIND_NEW = [['docs', 'Document', FileText], ['sheets', 'Spreadsheet', Sheet], ['slides', 'Presentation', Presentation], ['pdf', 'PDF', FileType]];
const run = (fn) => () => { Promise.resolve().then(fn).catch((e) => toast.error(e)); };
const trItems = (items) => items.filter(Boolean).map((x) => (x && typeof x === 'object' && typeof x.label === 'string' ? { ...x, label: t(x.label), ...(x.hint ? { hint: keyLabel(x.hint) } : {}) } : x));

/* ---------------- which files Office opens by click ---------------- */
// Hostwerk's own text editor stays for plain text and web files; Markdown and CSV open in Office only outside the
// folders of web servers, projects and apps (there the code editor is the better tool) – and always from Explorer
const NEVER = new Set(['txt', 'html', 'htm', 'xhtml', 'xml', 'json', 'svg', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'zip', 'tsv', 'tab', 'prn']);
const DEV_TEXT = new Set(['md', 'markdown', 'csv']);
const DEV_ROOT = /^(htdocs|web:|projekte$|app:|backups$|logs$|config$)/;
/** Root of the folder Files shows right now (it keeps it in the address) – null elsewhere (e.g. a file from Explorer) */
export function currentFilesRoot(hash = typeof location !== 'undefined' ? location.hash : '') {
  const m = /^#\/files\?(.*)$/.exec(String(hash || ''));
  return m ? new URLSearchParams(m[1]).get('root') : null;
}
export function opensInOffice(name, { root = currentFilesRoot(), settings = settingsNow() } = {}) {
  if (settings && settings.general && settings.general.openFromFiles === false) return false;
  const ext = extOf(name);
  if (!ext || NEVER.has(ext) || !FORMATS[ext] || !OFFICE_RE.test(String(name))) return false;
  if (DEV_TEXT.has(ext) && root && DEV_ROOT.test(root)) return false;
  return true;
}

/** files.open → the editor in front; respects "Double-click opens files in" (built-in / full editor / desktop app) */
export async function openFromFiles({ root, path, back, from }) {
  const f = canon({ root, path, name: basename(path) });
  const s = await loadSettings().catch(() => settingsNow());
  const mode = (s && s.integrations && s.integrations.defaultOpen) || 'builtin';
  if (mode === 'desktop' && from !== 'os') { await openInDesktop(f); return; }
  const extra = { back: back || filesRoute(f), ...(from === 'os' ? { os: 1 } : {}) };
  if (mode === 'onlyoffice') { const app = await ensureFullEditor(); if (app) extra.app = app; }
  openEditor(f, extra);
}

/* ---------------- dialog host ---------------- */
let dlg = null;
const dsubs = new Set();
let hosts = 0;
/** Show an Office dialog where the user is (false = no host mounted – the caller falls back to the screen) */
export function openDialog(d) {
  if (!hosts) {
    if (d && d.type === 'new') { location.hash = `/office-edit?${toQuery({ dialog: 'new', kind: d.kind, gallery: d.gallery ? 1 : '', root: d.root || '', dir: d.dir || '', back: filesRoute({ root: d.root || LIB_ROOT, path: d.dir || '', dir: true }) })}`; return true; }
    return false;
  }
  dlg = d; dsubs.forEach((f) => f(d)); return true;
}
const closeDialog = () => { dlg = null; dsubs.forEach((f) => f(null)); };
const refreshFiles = () => { try { window.dispatchEvent(new Event('hw:files-marks')); window.dispatchEvent(new Event('hw:files-reload')); } catch { /* old browser */ } };

/** One Office dialog (also used by the screen fallback) */
export function DialogView({ d, onClose, hasConverter }) {
  if (!d) return null;
  if (d.type === 'new') return <NewDialog kind={d.kind || 'docs'} dir={d.dir || ''} root={d.root && d.root !== LIB_ROOT ? d.root : null} rootLabel={d.rootLabel || ''} templateId={d.templateId || null} gallery={!!d.gallery} onClose={onClose} onCreated={(e) => { onClose(); refreshFiles(); openEditor(e, { back: filesRoute(e) }); }} />;
  if (d.type === 'share') return <ShareDialog entry={d.entry} onClose={() => { onClose(); refreshFiles(); }} />;
  if (d.type === 'info') return <InfoDialog entry={d.entry} onClose={onClose} />;
  if (d.type === 'tags') return <TagsDialog entries={d.entries} onClose={() => { onClose(); refreshFiles(); }} />;
  if (d.type === 'versions') return <VersionsSheet file={d.entry} onClose={onClose} onRestored={() => { onClose(); }} onCompare={(v) => { onClose(); openEditor(d.entry, { back: filesRoute(d.entry), compare: v.id }); }} />;
  if (d.type === 'convert') return <ConvertDialog entry={d.entry} hasConverter={hasConverter} onClose={onClose} />;
  return null;
}
/** Mounted once in Hostwerk (slot sidebar.footer) – renders nothing but the dialogs (Hostwerk's Modal portals to .app) */
export function OfficeDialogHost() {
  const [d, setD] = useState(dlg);
  const { hasConverter } = useOfficeStatus();
  useEffect(() => { hosts++; dsubs.add(setD); checkOsNotice(); return () => { hosts--; dsubs.delete(setD); }; }, []);
  return <><DialogView d={d} onClose={closeDialog} hasConverter={hasConverter} /><EngineDialogHost /><FullEditorDialogHost /><FeatureHost /><ThumbRenderer /></>;
}

/* ---------------- New ---------------- */
/** "Document" in the New menu: default format, first free name, opened at once (rename in the editor's title bar) */
export async function quickNew(kind, { root, path }) {
  const s = settingsNow();
  const ext = defaultExt(kind, s.general.newFormat) || (kind === 'pdf' ? 'pdf' : 'docx');
  const where = canon({ root, path: path || '' });
  const e = await createFile({ kind, ext, name: t(defaultBaseName(kind)), dir: where.path || '', root: where.root || null });
  const f = where.root ? { ...e, root: e.root || where.root } : e;
  openEditor(f, { back: filesRoute(f) });
  return f;
}
export const filesNew = [
  ...KIND_NEW.map(([kind, label, icon], i) => ({ id: `new-${kind}`, title: label, icon, group: 'office', order: i + 1, run: ({ root, path }) => quickNew(kind, { root, path }) })),
  { id: 'new-more', title: 'From template or other format …', icon: LayoutTemplate, group: 'office', order: 9, run: ({ root, path }) => { const w = canon({ root, path: path || '' }); openDialog({ type: 'new', kind: 'docs', gallery: true, root: w.root || LIB_ROOT, dir: w.path || ''}); } },
];

/* ---------------- right-click actions ---------------- */
const isOfficeFile = (e) => !!e && !e.dir && OFFICE_RE.test(e.name || e.path || '') && !NEVER.has(extOf(e.name || e.path));
const entryOf = (e) => canon({ root: e.root, path: e.path, name: e.name || basename(e.path) });
// favourites of the open folder, kept from files.marks – so the menu shows "Add" or "Remove"
const favCache = new Map();
const favKey = (root, path) => `${root || ''}|${path}`;
const isFav = (e) => favCache.get(favKey(e.root, e.path)) === true;
const dialogOr = (d) => { if (!openDialog(d)) toast(t('Open the file in Office for this.')); };
// Files already offers "Open with “Office”" (files.open) – the own entry only when the click does not open Office
const clickOpens = (e) => {
  let off = [];
  try { const st = typeof HAPI.getStore === 'function' ? HAPI.getStore() : null; const ui = st && st.settings && st.settings.general && st.settings.general.ui; off = (ui && ui.filesOpenOff) || []; } catch { off = []; }
  return !off.includes('office') && opensInOffice(e.name || basename(e.path), { root: e.root });
};
export const fileActions = [
  { id: 'open', title: 'Open in Office', icon: Pencil, when: (e) => isOfficeFile(e) && !clickOpens(e), run: (e) => openFromFiles({ root: e.root, path: e.path, back: filesRoute({ root: e.root, path: e.path }) }) },
  { id: 'full', title: 'Open in the full editor', icon: Sparkles, when: (e) => isOfficeFile(e) && (FORMATS[extOf(e.name || e.path)] || {}).kind !== 'viewer', run: async (e) => { const app = await ensureFullEditor(); if (app) openEditor(entryOf(e), { app, back: filesRoute(e) }); } },
  { id: 'desktop', title: 'Open in the desktop app', icon: Monitor, when: isOfficeFile, run: (e) => openInDesktop(entryOf(e)).catch((er) => toast.error(er)) },
  { id: 'convert-pdf', title: 'Convert to PDF', icon: FileOutput, when: (e) => isOfficeFile(e) && extOf(e.name || e.path) !== 'pdf' && (FORMATS[extOf(e.name || e.path)] || {}).kind !== 'viewer', run: (e) => { convertEntry(entryOf(e), 'pdf').then(refreshFiles).catch((er) => toast.error(er)); } },
  { id: 'convert', title: 'Convert to …', icon: FileOutput, when: (e) => isOfficeFile(e) && convertTargets(extOf(e.name || e.path), { hasConverter: true }).length > 0, run: (e) => dialogOr({ type: 'convert', entry: entryOf(e) }) },
  { id: 'share', title: 'Share link …', icon: Link2, when: (e) => isOfficeFile(e) && settingsNow().sharing.enabled, run: (e) => dialogOr({ type: 'share', entry: entryOf(e) }) },
  { id: 'versions', title: 'Versions …', icon: History, when: isOfficeFile, run: (e) => dialogOr({ type: 'versions', entry: entryOf(e) }) },
  { id: 'fav-add', title: 'Add to favourites', icon: Star, when: (e) => isOfficeFile(e) && !isFav(e), run: (e) => toggleFavorite([entryOf(e)], true).then(refreshFiles).catch((er) => toast.error(er)) },
  { id: 'fav-remove', title: 'Remove from favourites', icon: Star, when: (e) => isOfficeFile(e) && isFav(e), run: (e) => toggleFavorite([entryOf(e)], false).then(refreshFiles).catch((er) => toast.error(er)) },
  { id: 'tags', title: 'Tags, colour and note …', icon: Tag, when: isOfficeFile, run: (e) => dialogOr({ type: 'tags', entries: [entryOf(e)] }) },
  { id: 'template', title: 'Save as template …', icon: LayoutTemplate, when: (e) => isOfficeFile(e) && ['docs', 'sheets', 'slides'].includes((FORMATS[extOf(e.name || e.path)] || {}).kind), run: (e) => saveAsTemplate(entryOf(e)).catch((er) => toast.error(er)) },
  { id: 'info', title: 'Office info', icon: Info, when: isOfficeFile, run: (e) => dialogOr({ type: 'info', entry: entryOf(e) }) },
  // Office 1.1: clean, publish, report, rules, forms, encrypt/decrypt, older states (fileFeatures.js – same list as in the editor)
  ...['clean', 'publish', 'report', 'rules', 'forms', 'encrypt', 'decrypt', 'backups'].map((id) => {
    const item = (e) => featureItems(entryOf(e)).find((x) => x.id === id);
    const proto = featureItems({ name: id === 'rules' || id === 'forms' ? 'x.xlsx' : id === 'decrypt' ? 'x.hwvault' : 'x.docx', path: 'x' }).find((x) => x.id === id) || {};
    return { id: `f-${id}`, title: proto.title || id, icon: proto.icon, when: (e) => !!e && !e.dir && (isOfficeFile(e) || id === 'decrypt') && !!item(e), run: (e) => { const x = item(e); if (x) x.run(); } };
  }),
];

/* ---------------- marks behind the name ---------------- */
// one request per folder for all three marks (cached for a moment)
const markCache = new Map();
function folderMarks(root, path) {
  const k = `${root}|${path}`;
  const hit = markCache.get(k);
  if (hit && Date.now() - hit.t < 1500) return hit.p;
  const p = office.marks({ root, path }).then((r) => {
    const names = (r && r.names) || {};
    for (const [n, m] of Object.entries(names)) favCache.set(favKey(root, path ? `${path}/${n}` : n), !!m.favorite);
    return names;
  }).catch(() => ({}));
  markCache.set(k, { t: Date.now(), p });
  return p;
}
const pick = (fn) => async ({ root, path }) => { const names = await folderMarks(root, path || ''); const out = {}; for (const [n, m] of Object.entries(names)) { const v = fn(m); if (v) out[n] = v; } return out; };
export const fileMarks = [
  { id: 'favourite', icon: Star, marks: pick((m) => m.favorite && { title: t('Favourite') }) },
  { id: 'shared', icon: Link2, marks: pick((m) => m.shared && { title: t('Shared by link (read-only)') }) },
  { id: 'locked', icon: Lock, marks: pick((m) => m.lock && { title: t('Open in an editor by {by}', { by: m.lock.by || '?' }) }) },
];

/* ---------------- the "Office" section in the Office documents root ---------------- */
const VIEWS = [['recent', 'Recent', Clock], ['favorites', 'Favourites', Star], ['shared', 'Shared', Share2], ['tags', 'Tags', Tag], ['templates', 'Templates', LayoutTemplate], ['trash', 'Trash', Trash2]];
const remembered = (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : v; } catch { return d; } };
const remember = (k, v) => { try { localStorage.setItem(k, v); } catch { /* this window only */ } };
let focusSearch = false;
/** From Hostwerk's "New" menu / shortcuts (page actions): open a view or the search of the section */
export function sectionIntent({ view, search } = {}) { if (view) remember('office:section-view', view); if (search) focusSearch = true; }

// names of Hostwerk's roots (Files' own list) – "Webprojekte (htdocs)" instead of "htdocs" in the lists
let rootNames = null;
function useRootNames() {
  const [m, setM] = useState(rootNames);
  useEffect(() => { if (!rootNames && typeof HAPI.get === 'function') HAPI.get('/files/roots').then((l) => { rootNames = new Map((Array.isArray(l) ? l : []).map((r) => [r.id, r.label])); setM(rootNames); }, () => {}); }, []);
  return m;
}
/** where a listed file lives: its folder in the library, or "<root name> › folder" */
const placeOf = (e, names) => { const d = dirname(e.path); if (!e.root) return d; const n = (names && names.get(e.root)) || (/^opened:/.test(e.root) ? '' : e.root); return [n, d].filter(Boolean).join(' › '); };
const NEW_TILES = [['docs', 'Document', 'Letter, report, notes'], ['sheets', 'Spreadsheet', 'Budget, list, calculation'], ['slides', 'Presentation', 'Talk, pitch, photo show'], ['pdf', 'PDF', 'Form, scan, signature']];
const CAROUSEL = TEMPLATES.filter((x) => x.id !== 'blank' && ['docs', 'sheets', 'slides'].includes(x.kind));
const PAGE = 12;

export function OfficeSection({ ctx }) {
  const names = useRootNames();
  const dir = (ctx && ctx.path) || '';
  const [view, setView0] = useState(() => remembered('office:section-view', 'recent'));
  const setView = useCallback((v) => { setView0(v); remember('office:section-view', v); }, []);
  const [data, setData] = useState(null);
  const [tags, setTags] = useState([]);
  const [tag, setTag] = useState('');
  const [counts, setCounts] = useState({});
  const [q, setQ] = useState('');
  const [semantic, setSemantic] = useState(() => remembered('office:semantic', '0') === '1');
  const dq = useDebounced(q.trim(), 280);
  const [hits, setHits] = useState(null);
  const [more, setMore] = useState(1);
  const [over, setOver] = useState(false);
  const [up, setUp] = useState(null); // upload progress { done, total, name }
  const [favs, setFavs] = useState({}); // optimistic favourites: key → bool
  const searchRef = useRef(null);
  const load = useCallback(async () => {
    try {
      if (view === 'templates') { const r = await office.templates(); setData((Array.isArray(r) ? r : (r && r.templates) || []).map((x) => ({ ...x, template: true }))); }
      else if (view === 'tags') { const r = await office.tags(); const list = (Array.isArray(r) ? r : (r && r.tags) || []).map((x) => (typeof x === 'string' ? x : x.tag || x.name)).filter(Boolean); setTags(list); const tg = tag && list.includes(tag) ? tag : list[0] || ''; setTag(tg); setData(tg ? (await office.list({ view: 'all', tag: tg })).entries || [] : []); }
      else { const r = await office.list({ view }); setData((r.entries || []).filter((e) => !e.dir || view === 'trash')); if (r.counts) setCounts(r.counts); }
      setFavs({});
    } catch { setData([]); }
  }, [view, tag]);
  useEffect(() => { setData(null); setMore(1); load(); }, [load]);
  useChanged(() => { load(); }, [load]);
  // the Office root is open: the editors of the recent files get ready in idle time (the next click opens at once)
  useEffect(() => { if (view !== 'recent' || !data) return; const kinds = [...new Set(data.slice(0, 6).map((e) => kindFor(e.name)).filter(Boolean))].slice(0, 2); kinds.forEach((k) => warmEditor(k)); }, [view, data]);
  useEffect(() => {
    if (dq.length < 2) { setHits(null); return undefined; }
    let alive = true;
    setHits((h) => h || null);
    const go = semantic ? office.semanticSearch(dq, {}).catch((e) => { if (isMissing(e)) { toast(t('Search by meaning needs a newer Office server part – searching the words instead.')); setSemantic(false); remember('office:semantic', '0'); return office.search(dq, {}); } throw e; }) : office.search(dq, {});
    go.then((r) => { if (alive) setHits(r.hits || []); }, () => alive && setHits([]));
    return () => { alive = false; };
  }, [dq, semantic]);
  useEffect(() => { if (focusSearch && searchRef.current) { focusSearch = false; searchRef.current.focus(); } }, []);
  const where = { root: LIB_ROOT, dir };
  const openE = useCallback((e) => openEditor(e, { back: filesRoute(e) }), []);
  // own template: ask for the name, create the file in this folder and open it (no detour through the New dialog)
  const tplBusy = useRef(false);
  const fromTemplate = (x) => {
    if (tplBusy.current) return; // click + double click: one name dialog
    tplBusy.current = true;
    newFromOwnTemplate(x, { root: null, dir }).then((e) => { if (e) { refreshFiles(); openE(e); } }, (er) => toast.error(er)).finally(() => { tplBusy.current = false; });
  };
  const isFavNow = (e) => { const k = `${e.root || ''}|${e.path}`; return k in favs ? favs[k] : !!e.favorite; };
  const toggleFav = async (e) => {
    const k = `${e.root || ''}|${e.path}`;
    const on = !isFavNow(e);
    setFavs((f) => ({ ...f, [k]: on })); // optimistic
    try { await office.meta(e.root ? { root: e.root, path: e.path } : e.path, { favorite: on }); changed(); refreshFiles(); } catch (er) { setFavs((f) => ({ ...f, [k]: !on })); toast.error(er); }
  };
  const menu = (ev, e) => {
    ev.preventDefault();
    if (view === 'trash' && !hits) { contextMenu(ev, trItems([{ label: 'Restore', icon: RotateCcw, onClick: run(() => restoreTrash([e])) }, '-', { label: 'Delete forever …', icon: Trash2, danger: true, onClick: run(() => purgeTrash([e])) }])); return; }
    if (e.template) { contextMenu(ev, trItems([{ label: 'New from this template …', icon: FilePlus2, onClick: () => fromTemplate(e) }, { label: 'More templates and formats …', icon: LayoutTemplate, onClick: () => openDialog({ type: 'new', kind: e.kind || (FORMATS[e.ext] || {}).kind || 'docs', gallery: true, ...where }) }])); return; }
    const fe = { root: e.root, path: e.path, name: e.name };
    contextMenu(ev, trItems([
      { label: 'Open', icon: Pencil, hint: 'Enter', onClick: () => openE(e) },
      { label: 'Show in its folder', icon: FolderOpen, onClick: () => { location.hash = filesRoute(e).slice(1); } },
      { label: isFavNow(e) ? 'Remove from favourites' : 'Add to favourites', icon: Star, onClick: () => toggleFav(e) },
      '-',
      ...fileActions.filter((a) => !['open', 'fav-add', 'fav-remove'].includes(a.id) && a.when({ ...fe, dir: false })).map((a) => ({ label: a.title, icon: a.icon, onClick: () => a.run({ ...fe, dir: false }) })),
    ]));
  };
  /* drop files on the section: from the desktop → upload into this folder; from Files → a copy here */
  const dragKind = (e) => { const ty = [...((e.dataTransfer && e.dataTransfer.types) || [])]; return ty.includes('application/x-hostwerk-files') ? 'hw' : ty.includes('Files') ? 'os' : null; };
  const onDragOver = (e) => { if (!dragKind(e)) return; e.preventDefault(); e.stopPropagation(); e.dataTransfer.dropEffect = 'copy'; if (!over) setOver(true); };
  const onDrop = async (e) => {
    const how = dragKind(e);
    setOver(false);
    if (!how) return;
    e.preventDefault(); e.stopPropagation();
    if (how === 'os') { await uploadFiles(e.dataTransfer.files, dir, setUp).catch((er) => toast.error(er)); refreshFiles(); return; }
    let d = null; try { d = JSON.parse(e.dataTransfer.getData('application/x-hostwerk-files')); } catch { d = null; }
    if (!d || !Array.isArray(d.paths) || !d.paths.length) return;
    if (d.root === LIB_ROOT) return; // already in the library – Files moves it itself
    try {
      for (const p of d.paths) { const { bytes } = await office.read({ root: d.root, path: p }); await office.create(dir, basename(p), bytes); }
      changed(); refreshFiles();
      toast.success(pl(d.paths.length, 'Copied to the Office documents', '{n} files copied to the Office documents'));
    } catch (er) { toast.error(er); }
  };
  const list = hits || data;
  const EMPTY = {
    recent: ['No document opened yet', 'Create one with the tiles above – or drop files here.', Clock],
    favorites: ['No favourites yet', 'Click the star on a file to find it here.', Star],
    shared: ['Nothing shared', 'Right-click a file → “Share link …” for a read-only link.', Share2],
    tags: ['No tags yet', 'Right-click a file → “Tags, colour and note …”.', Tag],
    templates: ['No own templates yet', 'Right-click a document → “Save as template …”.', LayoutTemplate],
    trash: ['The trash is empty', 'Files you delete in Files are in Hostwerk’s trash.', Trash2],
  }[view];
  const cards = !hits && view !== 'trash' && view !== 'templates';
  const shown = list ? list.slice(0, cards ? PAGE * more : 400) : null;
  return (
    <div className={`of-shell-fsec ${over ? 'is-over' : ''}`} onDragOver={onDragOver} onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget)) setOver(false); }} onDrop={onDrop}>
      <DefaultAppHint />
      <section className="of-shell-hero" aria-label={t('New')}>
        <div className="of-shell-hero-head">
          <div className="stack of-shell-gap-0 of-shell-min0">
            <b className="of-shell-hero-title">{t('Create something new')}</b>
            <span className="small muted">{t('Opens right away – the name can be changed in the editor.')}</span>
          </div>
          <button type="button" className="btn sm ghost" onClick={() => openDialog({ type: 'new', kind: 'docs', gallery: true, ...where })}><LayoutTemplate size={14} aria-hidden="true" />{t('All templates')}</button>
        </div>
        <div className="of-shell-newtiles" role="group" aria-label={t('New')}>
          {NEW_TILES.map(([kind, label, sub]) => (
            <button key={kind} type="button" className={`of-shell-newtile of-shell-k-${kind}`} style={{ '--of-k': KIND_COLOR[kind] }} onClick={run(() => quickNew(kind, { root: LIB_ROOT, path: dir }))}
              onMouseEnter={() => prefetchEditor(kind)} onFocus={() => prefetchEditor(kind)} title={t('New {kind} in this folder', { kind: t(label) })}>
              <span className="of-shell-newtile-ic"><FileIcon name={`x.${defaultExt(kind, settingsNow().general.newFormat) || kind}`} size={32} /><Plus size={12} className="of-shell-newtile-plus" aria-hidden="true" /></span>
              <span className="stack of-shell-gap-0 of-shell-min0"><b className="small">{t(label)}</b><span className="tiny muted ellipsis">{t(sub)}</span></span>
            </button>
          ))}
        </div>
        {CAROUSEL.length > 0 && <TemplateStrip where={where} />}
      </section>

      <div className="of-shell-fsec-search">
        <Search size={16} aria-hidden="true" />
        <input ref={searchRef} className="input" type="search" value={q} onChange={(e) => setQ(e.target.value)} onKeyDown={(e) => { if (e.key === 'Escape' && q) { e.stopPropagation(); setQ(''); } }}
          placeholder={t(semantic ? 'Describe what you are looking for …' : 'Search in all Office documents (also the text) …')} aria-label={t('Search in all Office documents')} />
        {q && <button type="button" className="btn ghost xs icon" title={t('Clear')} aria-label={t('Clear')} onClick={() => setQ('')}><X size={12} /></button>}
        <button type="button" className={`of-shell-semantic ${semantic ? 'on' : ''}`} aria-pressed={semantic} title={t('Search by meaning – finds documents about a topic even without the exact words (uses Hostwerk’s AI)')}
          onClick={() => { const v = !semantic; setSemantic(v); remember('office:semantic', v ? '1' : '0'); }}><Sparkles size={14} aria-hidden="true" /><span>{t('By meaning')}</span></button>
      </div>

      {!hits && <div className="of-shell-tabs" role="tablist" aria-label={t('Views')}>
        {VIEWS.map(([id, label, I]) => {
          const n = id === 'trash' ? counts.trash : id === 'favorites' ? counts.favorites : id === 'shared' ? counts.shared : null;
          return <button key={id} type="button" role="tab" aria-selected={view === id} className={`of-shell-tab ${view === id ? 'on' : ''}`} onClick={() => setView(id)}><I size={14} aria-hidden="true" />{t(label)}{n ? <span className="of-shell-count">{n}</span> : null}</button>;
        })}
      </div>}
      {!hits && view === 'tags' && tags.length > 0 && <div className="row-tight of-shell-wrap">{tags.map((x) => <Raw as="button" key={x} type="button" className={`of-shell-tag is-btn ${x === tag ? 'on' : ''}`} onClick={() => setTag(x)}><Keep>{x}</Keep></Raw>)}</div>}
      {hits && <div className="small muted of-shell-hitsinfo">{hits.length ? pl(hits.length, '{n} hit', '{n} hits') : ''}{semantic ? ` · ${t('by meaning')}` : ''}</div>}

      {up && <div className="of-shell-upload"><Upload size={14} aria-hidden="true" /><span className="small grow ellipsis">{t('Uploading {name} …', { name: up.name })}</span><span className="tiny muted">{`${up.done + 1}/${up.total}`}</span></div>}

      {!list ? <SectionSkeleton cards={cards || view === 'templates'} /> : list.length === 0 ? (
        hits ? <Empty icon={Search} title="Nothing found" text={semantic ? t('Try other words – or switch “By meaning” off for an exact search.') : t('Check the spelling – or switch on “By meaning”.')} />
          : <Empty icon={EMPTY[2]} title={EMPTY[0]} text={t(EMPTY[1])} action={view === 'recent' ? <Btn className="primary" icon={FilePlus2} onClick={run(() => quickNew('docs', { root: LIB_ROOT, path: dir }))}>New document</Btn> : null} />
      ) : cards ? (
        <>
          <ul className="of-shell-cards" aria-label={t(VIEWS.find((v) => v[0] === view)[1])}>
            {shown.map((e) => <RecentCard key={`${e.root || ''}|${e.id || e.path}`} e={e} place={placeOf(e, names)} fav={isFavNow(e)} onOpen={openE} onMenu={menu} onFav={toggleFav} />)}
          </ul>
          {list.length > shown.length && <div className="of-shell-center row-tight"><Btn className="sm ghost" onClick={() => setMore((m) => m + 1)}>{t('Show more ({n})', { n: list.length - shown.length })}</Btn></div>}
        </>
      ) : view === 'templates' && !hits ? (
        <div className="of-shell-tpls is-wide">{shown.map((x) => <TemplateCard key={x.id} kind={x.kind || (FORMATS[x.ext] || {}).kind || 'docs'} ext={x.ext} name={x.name} cat={x.category || t('My templates')} own on={false} onPick={() => fromTemplate(x)} onOpen={() => fromTemplate(x)} onMenu={(ev) => menu(ev, x)} />)}</div>
      ) : (
        <VirtualRows items={shown} height={hits ? 58 : 44} render={(e) => (
          <li key={`${e.root || ''}|${e.id || e.path}`} className="of-shell-row is-click" tabIndex={0} style={{ height: hits ? 58 : 44 }}
            onClick={() => (view === 'trash' && !hits ? restoreTrash([e]).catch((er) => toast.error(er)) : openE(e))}
            onMouseEnter={() => prefetchFor(e.name)} onKeyDown={(ev) => { if (ev.key === 'Enter') ev.currentTarget.click(); }} onContextMenu={(ev) => menu(ev, e)}>
            <FileIcon entry={e} size={20} />
            <div className="grow of-shell-min0">
              <div className="small ellipsis"><b>{hits ? <Highlight value={e.name} q={dq} /> : <Keep>{e.name}</Keep>}</b> <span className="tiny muted"><Keep>{view === 'trash' && !hits ? dirname(e.from || '') : placeOf(e, names)}</Keep></span></div>
              {hits && e.snippet && <div className="tiny muted of-shell-snippet">… <Highlight value={e.snippet} q={dq} /> …</div>}
            </div>
            {hits && e.score != null && <span className="of-shell-score" title={t('How well it matches')}><i style={{ width: `${Math.round(Math.max(0.1, Math.min(1, e.score)) * 100)}%` }} /></span>}
            <span className="tiny muted">{view === 'trash' && !hits ? t('Restore') : fmtWhen(e.opened || e.mtime)}</span>
          </li>
        )} />
      )}
      {!hits && view === 'trash' && (data || []).length > 0 && <div className="row-tight"><Btn className="sm ghost" icon={Trash2} onClick={run(() => emptyTrash(data.length))}>Empty trash …</Btn></div>}
      <div className="row-tight of-shell-wrap tiny muted of-shell-fsec-foot">
        <span className="grow">{t('Tip: drop files here to add them · Alt+Q in a document shows every command')}</span>
        <a className="link-btn tiny" href="#/settings/office"><Settings2 size={12} aria-hidden="true" /> {t('Office settings')}</a>
      </div>
      {over && <div className="of-shell-drop is-section" aria-hidden="true"><div className="of-shell-drop-card"><Upload size={22} /><b>{t('Drop to add to the Office documents')}</b></div></div>}
    </div>
  );
}
export const officeSectionWhen = (ctx) => !!ctx && ctx.root === LIB_ROOT;

/** one recent/favourite/shared file as a card with its first page */
const RecentCard = memo(function RecentCard({ e, place, fav, onOpen, onMenu, onFav }) {
  const kind = kindFor(e.name) || 'viewer';
  return (
    <li className={`of-shell-card of-shell-k-${kind}`} style={{ '--of-k': KIND_COLOR[kind] || KIND_COLOR.viewer }}>
      <button type="button" className="of-shell-card-main" onClick={() => onOpen(e)} onContextMenu={(ev) => onMenu(ev, e)} onMouseEnter={() => prefetchFor(e.name)} onFocus={() => prefetchFor(e.name)}
        onKeyDown={(ev) => { if (ev.key === 'ContextMenu' || (ev.shiftKey && ev.key === 'F10')) { const r = ev.currentTarget.getBoundingClientRect(); onMenu({ preventDefault() {}, clientX: r.left + 20, clientY: r.top + 20 }, e); } }}>
        <span className="of-shell-card-thumb"><Thumb entry={e} /></span>
        <span className="of-shell-card-meta">
          <FileIcon entry={e} size={16} />
          <span className="stack of-shell-gap-0 of-shell-min0 grow">
            <Raw as="b" className="small ellipsis" title={e.name}>{e.name}</Raw>
            <span className="tiny muted ellipsis"><Keep>{[fmtWhen(e.opened || e.mtime), place].filter(Boolean).join(' · ')}</Keep></span>
          </span>
        </span>
      </button>
      <button type="button" className={`of-shell-card-fav ${fav ? 'on' : ''}`} aria-pressed={fav} title={t(fav ? 'Remove from favourites' : 'Add to favourites')} aria-label={t(fav ? 'Remove from favourites' : 'Add to favourites')} onClick={() => onFav(e)}><Star size={14} /></button>
      {e.shared && <span className="of-shell-card-badge" title={t('Shared by link (read-only)')}><Link2 size={12} aria-hidden="true" /></span>}
      {e.lock && <span className="of-shell-card-badge is-lock" title={t('Open in an editor by {by}', { by: e.lock.by || '?' })}><Lock size={12} aria-hidden="true" /></span>}
    </li>
  );
});

/** built-in templates with their real first page – a click opens the New dialog with it chosen */
function TemplateStrip({ where }) {
  const strip = useRef(null);
  const scroll = (d) => { const el = strip.current; if (el) el.scrollBy({ left: d * el.clientWidth * 0.8, behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' }); };
  return (
    <div className="of-shell-carousel">
      <button type="button" className="btn ghost icon sm of-shell-carousel-btn is-prev" aria-label={t('Previous templates')} title={t('Previous templates')} onClick={() => scroll(-1)}><ChevronLeft size={16} /></button>
      <div className="of-shell-carousel-track" ref={strip}>
        {CAROUSEL.map((x) => {
          const ext = defaultExt(x.kind, settingsNow().general.newFormat);
          return <TemplateCard key={`${x.kind}:${x.id}`} kind={x.kind} ext={ext} tplId={x.id} name={t(x.name)} cat={t(KINDS[x.kind].label)} on={false}
            onPick={() => openDialog({ type: 'new', kind: x.kind, templateId: x.id, gallery: true, ...where })} onOpen={() => openDialog({ type: 'new', kind: x.kind, templateId: x.id, gallery: true, ...where })} />;
        })}
      </div>
      <button type="button" className="btn ghost icon sm of-shell-carousel-btn is-next" aria-label={t('More templates')} title={t('More templates')} onClick={() => scroll(1)}><ChevronRight size={16} /></button>
    </div>
  );
}

function SectionSkeleton({ cards }) {
  return cards ? <ul className="of-shell-cards" aria-busy="true">{[0, 1, 2, 3, 4, 5].map((i) => <li key={i} className="of-shell-card is-skel"><span className="of-shell-card-thumb sk" /><span className="of-shell-card-meta"><i className="sk of-shell-sk-row" /></span></li>)}</ul>
    : <ul className="of-shell-rows" aria-busy="true">{[0, 1, 2, 3].map((i) => <li key={i} className="of-shell-row"><i className="sk of-shell-sk-row" style={{ width: `${85 - i * 12}%` }} /></li>)}</ul>;
}

/**
 * Long lists (search hits, trash) – only the rows in view are in the DOM (fixed row height); short lists render whole.
 */
function VirtualRows({ items, height, render, max = 420 }) {
  const box = useRef(null);
  const [top, setTop] = useState(0);
  if (items.length <= 40) return <ul className="of-shell-rows of-shell-fsec-list">{items.map(render)}</ul>;
  const view = Math.min(max, items.length * height);
  const first = Math.max(0, Math.floor(top / height) - 6);
  const last = Math.min(items.length, Math.ceil((top + view) / height) + 6);
  return (
    <div ref={box} className="of-shell-vlist" style={{ height: view }} onScroll={(e) => setTop(e.currentTarget.scrollTop)}>
      <ul className="of-shell-rows" style={{ height: items.length * height, paddingTop: first * height, boxSizing: 'border-box' }}>{items.slice(first, last).map(render)}</ul>
    </div>
  );
}

