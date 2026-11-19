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
import { useEffect, useState } from 'react';
import {
  FileText, Sheet, Presentation, FileType, LayoutTemplate, Pencil, Sparkles, FileOutput, Link2, History, Star, Tag, Info, Monitor, Clock, Share2,
  Trash2, RotateCcw, Search, Settings2, Lock, FilePlus2, X,
} from 'lucide-react';
import { Btn, Seg, Loading } from '@hostwerk/ui';
import * as HAPI from '@hostwerk/api';
import { contextMenu } from '@hostwerk/kit';
import { t } from '@hostwerk/i18n';
import { office } from '../lib/api';
import { toast } from '../lib/toast';
import { FORMATS, OFFICE_RE, extOf } from '../lib/formats';
import { keyLabel } from '../editors/_shared/keys.js';
import { FileIcon, Keep, Raw, Highlight, fmtWhen, useDebounced } from './common';
import { settingsNow, useChanged, useOfficeStatus, loadSettings } from './store';
import { convertTargets, defaultExt, defaultBaseName, dirname, basename, toQuery } from './util';
import { LIB_ROOT, canon, openEditor, filesRoute, createFile, convertEntry, openInDesktop, toggleFavorite, saveAsTemplate, restoreTrash, purgeTrash, emptyTrash } from './ops';
import { NewDialog, ShareDialog, InfoDialog, TagsDialog, VersionsSheet, ConvertDialog } from './dialogs';
import { EngineDialogHost, FullEditorDialogHost, ensureFullEditor } from './engine';
import { DefaultAppHint, checkOsNotice } from './osCard';

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
  if (d.type === 'new') return <NewDialog kind={d.kind || 'docs'} dir={d.dir || ''} root={d.root && d.root !== LIB_ROOT ? d.root : null} rootLabel={d.rootLabel || ''} gallery={!!d.gallery} onClose={onClose} onCreated={(e) => { onClose(); refreshFiles(); openEditor(e, { back: filesRoute(e) }); }} />;
  if (d.type === 'share') return <ShareDialog entry={d.entry} onClose={() => { onClose(); refreshFiles(); }} />;
  if (d.type === 'info') return <InfoDialog entry={d.entry} onClose={onClose} />;
  if (d.type === 'tags') return <TagsDialog entries={d.entries} onClose={() => { onClose(); refreshFiles(); }} />;
  if (d.type === 'versions') return <VersionsSheet file={d.entry} onClose={onClose} onRestored={() => { onClose(); toast.success('Version restored'); }} />;
  if (d.type === 'convert') return <ConvertDialog entry={d.entry} hasConverter={hasConverter} onClose={onClose} />;
  return null;
}
/** Mounted once in Hostwerk (slot sidebar.footer) – renders nothing but the dialogs (Hostwerk's Modal portals to .app) */
export function OfficeDialogHost() {
  const [d, setD] = useState(dlg);
  const { hasConverter } = useOfficeStatus();
  useEffect(() => { hosts++; dsubs.add(setD); checkOsNotice(); return () => { hosts--; dsubs.delete(setD); }; }, []);
  return <><DialogView d={d} onClose={closeDialog} hasConverter={hasConverter} /><EngineDialogHost /><FullEditorDialogHost /></>;
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

export function OfficeSection({ ctx }) {
  const names = useRootNames();
  const dir = (ctx && ctx.path) || '';
  const [view, setView0] = useState(() => remembered('office:section-view', 'recent'));
  const setView = (v) => { setView0(v); remember('office:section-view', v); };
  const [data, setData] = useState(null);
  const [tags, setTags] = useState([]);
  const [tag, setTag] = useState('');
  const [counts, setCounts] = useState({});
  const [q, setQ] = useState('');
  const dq = useDebounced(q.trim(), 300);
  const [hits, setHits] = useState(null);
  const load = async () => {
    try {
      if (view === 'templates') { const r = await office.templates(); setData((Array.isArray(r) ? r : (r && r.templates) || []).map((x) => ({ ...x, template: true }))); }
      else if (view === 'tags') { const r = await office.tags(); const list = (Array.isArray(r) ? r : (r && r.tags) || []).map((x) => (typeof x === 'string' ? x : x.tag || x.name)).filter(Boolean); setTags(list); const tg = tag && list.includes(tag) ? tag : list[0] || ''; setTag(tg); setData(tg ? (await office.list({ view: 'all', tag: tg })).entries || [] : []); }
      else { const r = await office.list({ view }); setData((r.entries || []).filter((e) => !e.dir || view === 'trash')); if (r.counts) setCounts(r.counts); }
    } catch { setData([]); }
  };
  useEffect(() => { setData(null); load(); }, [view, tag]); // eslint-disable-line react-hooks/exhaustive-deps
  useChanged(() => { load(); }, [view, tag]);
  useEffect(() => {
    if (dq.length < 2) { setHits(null); return undefined; }
    let alive = true;
    office.search(dq, {}).then((r) => { if (alive) setHits(r.hits || []); }, () => alive && setHits([]));
    return () => { alive = false; };
  }, [dq]);
  const where = { root: LIB_ROOT, dir };
  const openE = (e) => openEditor(e, { back: filesRoute(e) });
  const menu = (ev, e) => {
    ev.preventDefault();
    if (view === 'trash') { contextMenu(ev, trItems([{ label: 'Restore', icon: RotateCcw, onClick: run(() => restoreTrash([e])) }, '-', { label: 'Delete forever …', icon: Trash2, danger: true, onClick: run(() => purgeTrash([e])) }])); return; }
    if (e.template) { contextMenu(ev, trItems([{ label: 'New from this template', icon: FilePlus2, onClick: () => openDialog({ type: 'new', kind: e.kind || (FORMATS[e.ext] || {}).kind || 'docs', gallery: true, ...where }) }])); return; }
    const fe = { root: e.root, path: e.path, name: e.name };
    contextMenu(ev, trItems([
      { label: 'Open', icon: Pencil, hint: 'Enter', onClick: () => openE(e) },
      { label: 'Show in its folder', icon: FileText, onClick: () => { location.hash = filesRoute(e).slice(1); } },
      '-',
      ...fileActions.filter((a) => a.id !== 'open' && a.when({ ...fe, dir: false })).map((a) => ({ label: a.title, icon: a.icon, onClick: () => a.run({ ...fe, dir: false }) })),
    ]));
  };
  const list = hits || data;
  const empty = {
    recent: 'Files you open in Office show up here.', favorites: 'Right-click a file → “Add to favourites”.', shared: 'Right-click a file → “Share link …” for a read-only link.',
    tags: 'Right-click a file → “Tags, colour and note …”.', templates: 'Right-click a document → “Save as template …”.', trash: 'Nothing in Office’s trash. Files you delete in Files are in Hostwerk’s trash.',
  }[view];
  return (
    <div className="stack of-shell-fsec">
      <DefaultAppHint />
      <div className="of-shell-fsec-new" role="group" aria-label={t('New')}>
        {KIND_NEW.map(([kind, label, I]) => (
          <button key={kind} type="button" className={`of-shell-fsec-newbtn of-shell-k-${kind}`} onClick={run(() => quickNew(kind, { root: LIB_ROOT, path: dir }))} title={t('New {kind} in this folder', { kind: t(label) })}>
            <span className={`of-shell-icon of-shell-k-${kind} of-shell-icon-md`}><I size={20} /></span><span className="small">{t(label)}</span>
          </button>
        ))}
        <button type="button" className="of-shell-fsec-newbtn" onClick={() => openDialog({ type: 'new', kind: 'docs', gallery: true, ...where })} title={t('From template or other format …')}>
          <span className="of-shell-icon of-shell-icon-md"><LayoutTemplate size={20} /></span><span className="small">{t('Template …')}</span>
        </button>
      </div>
      <div className="of-shell-fsec-search">
        <Search size={14} aria-hidden="true" />
        <input className="input sm" type="search" value={q} autoFocus={focusSearch} onFocus={() => { focusSearch = false; }} onChange={(e) => setQ(e.target.value)} placeholder={t('Search in all Office documents (also the text) …')} aria-label={t('Search in all Office documents')} />
        {q && <button type="button" className="btn ghost xs icon" title={t('Clear')} aria-label={t('Clear')} onClick={() => setQ('')}><X size={12} /></button>}
      </div>
      {!hits && <div className="of-shell-fsec-views">
        <Seg value={view} onChange={setView} options={VIEWS.map(([id, label]) => ({ id, label: id === 'trash' && counts.trash ? `${t(label)} (${counts.trash})` : id === 'favorites' && counts.favorites ? `${t(label)} (${counts.favorites})` : t(label) }))} />
      </div>}
      {!hits && view === 'tags' && tags.length > 0 && <div className="row-tight of-shell-wrap">{tags.map((x) => <Raw as="button" key={x} type="button" className={`of-shell-tag is-btn ${x === tag ? 'on' : ''}`} onClick={() => setTag(x)}><Keep>{x}</Keep></Raw>)}</div>}
      {!list ? <Loading inline /> : list.length === 0 ? <p className="small muted of-shell-p">{hits ? t('Nothing found.') : t(empty)}</p> : (
        <ul className="of-shell-rows of-shell-fsec-list">
          {list.slice(0, 40).map((e) => (
            <li key={`${e.root || ''}|${e.id || e.path}`} className="of-shell-row is-click" tabIndex={0}
              onClick={() => (view === 'trash' && !hits ? restoreTrash([e]).catch((er) => toast.error(er)) : e.template ? openDialog({ type: 'new', kind: e.kind || 'docs', gallery: true, ...where }) : openE(e))}
              onKeyDown={(ev) => { if (ev.key === 'Enter') ev.currentTarget.click(); }} onContextMenu={(ev) => menu(ev, e)}>
              <FileIcon entry={e} />
              <div className="grow of-shell-min0">
                <div className="small ellipsis"><b>{hits ? <Highlight value={e.name} q={dq} /> : <Keep>{e.name}</Keep>}</b> {!e.template && <span className="tiny muted"><Keep>{view === 'trash' && !hits ? dirname(e.from || '') : placeOf(e, names)}</Keep></span>}</div>
                {hits && e.snippet && <div className="tiny muted of-shell-snippet">… <Highlight value={e.snippet} q={dq} /> …</div>}
              </div>
              {e.favorite && <Star size={12} className="of-shell-fav" aria-label={t('Favourite')} />}
              <span className="tiny muted">{view === 'trash' && !hits ? t('Restore') : e.template ? <Keep>{e.category || ''}</Keep> : fmtWhen(e.opened || e.mtime)}</span>
            </li>
          ))}
        </ul>
      )}
      {!hits && view === 'trash' && (data || []).length > 0 && <div className="row-tight"><Btn className="sm ghost" icon={Trash2} onClick={run(() => emptyTrash(data.length))}>Empty trash …</Btn></div>}
      <div className="row-tight of-shell-wrap tiny muted">
        <a className="link-btn tiny" href="#/settings/office"><Settings2 size={12} aria-hidden="true" /> {t('Office settings')}</a>
      </div>
    </div>
  );
}
export const officeSectionWhen = (ctx) => !!ctx && ctx.root === LIB_ROOT;
