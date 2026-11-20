/*
 * Office inside the rest of Hostwerk: the preview in Files (files.preview – right-click → "Preview", or every office
 * file when the click is switched off under Settings → Advanced → Open files), a dashboard tile with recent documents
 * (dashboard.tiles) and the editor as a full view (screen office-edit: from Files, from Explorer, own windows).
 * Everything else of Files – files.open / files.new / files.actions / files.marks / the Office section – is in files.jsx.
 */
import { useEffect, useState } from 'react';
import { Pencil, Monitor, FileText, Sparkles, FilePlus2, ArrowLeft } from 'lucide-react';
import { useWindowChrome } from './chrome';
import { Btn, TileEmpty } from '@hostwerk/ui';
import { t, pl } from '@hostwerk/i18n';
import { office } from '../lib/api';
import { toast } from '../lib/toast';
import { OFFICE_RE, formatOf, isTemplateName } from '../lib/formats';
import { KIND_COLOR } from '../lib/fileTypes';
import { openTemplateFile } from './templates';
import { FilePreviewBox } from './preview';
import { FileIcon, Keep, fmtWhen } from './common';
import { settingsNow, loadSettings, useChanged, useOfficeStatus } from './store';
import { LIB_ROOT, canon, openEditor, openInDesktop, filesRoute } from './ops';
import { parseQuery, basename, fmtBytes } from './util';
import EditorHost from './EditorHost';
import { ensureFullEditor } from './engine';
import { DialogView } from './files';
import ThumbRenderer from './thumbHost';
import { FileTypeIcon } from '../lib/fileIcons';
import { familyOf, iconInfo } from '../lib/fileTypes';
import { canThumb, ensureThumb } from '../lib/thumbs';
import { prefetchFor } from '../lib/loader';

/*
 * Coloured icons and first-page thumbnails for Hostwerk's Files (SDK 3.16 slots files.icon / files.thumb – older
 * Hostwerk versions ignore unknown slots, so nothing breaks there):
 *   files.icon  { id, match: { test(name) }, Component({ name, size, entry }), info(name) → { color, family, label } }
 *   files.thumb { id, match: { test(name) }, thumb({ root, path, name, mtime, size }) → Promise<url|null> }
 * A thumbnail request also fetches the matching editor chunk in the background (the click after it opens fast).
 */
// only the formats Office itself opens or edits – text, code, images, archives, databases and folders keep Hostwerk's own icons
const ICON_FAMILIES = new Set(['word', 'sheet', 'slides', 'pdf', 'mail', 'contact', 'calendar', 'ebook', 'diagram', 'note', 'publisher', 'project', 'google', 'iwork']);
const iconMatch = { test: (name) => ICON_FAMILIES.has(familyOf(String(name || ''))) };
function FilesIcon({ name, entry, size = 16 }) { return <FileTypeIcon name={name || (entry && entry.name)} entry={entry && entry.dir ? entry : undefined} size={size} showExt={size >= 48} />; }
export const filesIcon = [{ id: 'office', title: 'Office', match: iconMatch, Component: FilesIcon, info: (name) => { const i = iconInfo(name); return { color: i.color, family: i.family, label: isTemplateName(name) ? ((formatOf(name) || {}).label || i.label) : i.label }; } }];
export const filesThumb = [{
  id: 'office', title: 'Office', match: { test: (name) => canThumb({ name: String(name || '') }) },
  thumb: (x) => { const name = x.name || basename(x.path); prefetchFor(name); return ensureThumb(canon({ root: x.root, path: x.path, name, mtime: x.mtime, size: x.size })); },
}];

loadSettings().catch(() => {}); // the slots ask settingsNow() synchronously (openFromFiles)

/** files.preview `match`: a RegExp-like object, so the setting "Open office files from Files" applies live */
export const officeMatch = { test: (p) => settingsNow().general.openFromFiles !== false && OFFICE_RE.test(String(p || '')) };

/** "In Word öffnen" when Microsoft Office is there (status.converters.msoffice), else the system's app */
const toastError = (e) => toast.error(e);
const DESKTOP_APP = { docs: 'word', sheets: 'excel', slides: 'powerpoint' };
const APP_NAME = { word: 'Word', excel: 'Excel', powerpoint: 'PowerPoint' };
export function desktopLabel(kind, status) {
  const ms = status && status.converters && status.converters.msoffice;
  const app = DESKTOP_APP[kind];
  return app && ms && ms[app] ? t('Open in {app}', { app: APP_NAME[app] }) : t('Open in the desktop app');
}

/**
 * Files' preview pane (files.preview – right-click → "Preview", or a click when Office does not open files): the
 * document as it is – all pages in print layout, the grid with its sheet tabs, the slides, the PDF – fitted to the
 * width and scrollable, with big "Edit" and "Open in Word …" buttons; templates offer "New document" first.
 */
export function FilesPreview({ root, path }) {
  const name = basename(path);
  const f = canon({ root, path, name });
  const fmt = formatOf(name);
  const { status } = useOfficeStatus();
  const [info, setInfo] = useState(null);
  useEffect(() => { setInfo(null); office.info(f).then((i) => setInfo(i || null), () => {}); }, [root, path]); // eslint-disable-line react-hooks/exhaustive-deps
  const back = filesRoute({ root, path });
  const tpl = isTemplateName(name);
  const kind = fmt ? fmt.kind : 'viewer';
  const facts = [fmt ? t(fmt.label) : '', info && info.size != null ? fmtBytes(info.size) : '',
    info && info.pages ? pl(info.pages, '{n} page', '{n} pages') : info && info.slides ? pl(info.slides, '{n} slide', '{n} slides') : info && info.sheets ? pl(Array.isArray(info.sheets) ? info.sheets.length : info.sheets, '{n} sheet', '{n} sheets') : '',
    info && info.mtime ? fmtWhen(info.mtime) : ''].filter(Boolean).join(' · ');
  const edit = () => openEditor(f, { back, ...(fmt && fmt.edit === 'view' ? { ro: 1 } : {}) });
  const newFrom = () => openTemplateFile(f, { openEditor, filesRoute, editTemplate: (x) => openEditor(x, { back }) }).catch((e) => toastError(e));
  const desk = !status || status.desktop !== false;
  return (
    <div className={`of-shell-fpreview of-shell-k-${kind}`} style={{ '--of-k': KIND_COLOR[kind] || KIND_COLOR.viewer }}>
      <div className="of-shell-fpreview-bar">
        <FileIcon name={name} size={32} />
        <div className="stack of-shell-gap-0 grow of-shell-min0">
          <b className="ellipsis"><Keep>{name}</Keep></b>
          <span className="tiny muted ellipsis"><Keep>{facts}</Keep></span>
        </div>
        <div className="row-tight of-shell-wrap of-shell-fpreview-acts">
          {tpl && <Btn className="primary" icon={FilePlus2} onClick={newFrom}>New document</Btn>}
          {fmt && fmt.edit !== 'view' && <Btn className={tpl ? '' : 'primary'} icon={Pencil} onClick={edit}>{tpl ? t('Edit template') : t('Edit')}</Btn>}
          {fmt && fmt.edit === 'view' && <Btn className="primary" icon={FileText} onClick={edit}>Open in Office</Btn>}
          {desk && fmt && <Btn icon={Monitor} onClick={() => openInDesktop(f).catch((e) => toastError(e))}>{desktopLabel(kind, status)}</Btn>}
          {fmt && kind !== 'viewer' && <Btn className="ghost icon" icon={Sparkles} title="Open in the full editor" aria-label="Open in the full editor" onClick={async () => { const app = await ensureFullEditor(); if (app) openEditor(f, { app, back }); }} />}
        </div>
      </div>
      <FilePreviewBox key={`${root}|${path}`} file={{ ...f, size: info ? info.size : undefined, mtime: info ? info.mtime : undefined }} pane />
    </div>
  );
}

/** Dashboard tile: recent documents */
export function RecentTile() {
  const [list, setList] = useState(null);
  const load = () => office.list({ view: 'recent' }).then((r) => setList((r.entries || []).filter((e) => !e.dir).slice(0, 6)), () => setList([]));
  useEffect(() => { load(); }, []);
  useChanged(load, []);
  if (!list) return null;
  if (!list.length) return <TileEmpty icon={FileText} text="No documents yet." action="New document" href="#/ext/office?action=new&kind=docs" />;
  return (
    <div className="card-body list-rows">
      {list.map((e) => <a key={`${e.root || ''}|${e.path}`} className="list-row-item" href={`#${filesRoute(e)}`} onClick={(ev) => { ev.preventDefault(); openEditor(e, { back: filesRoute(e) }); }}><FileIcon entry={e} /><span className="grow ellipsis"><Keep>{e.name}</Keep></span><span className="small muted">{fmtWhen(e.mtime)}</span></a>)}
    </div>
  );
}

/** slim title bar of a screen without the editor (New dialog fallback, nothing chosen): back + drag region in the desktop app */
function WindowBar({ onBack }) {
  const c = useWindowChrome();
  return (
    <div className={`of-shell-wbar is-drag ${c.on ? 'is-wco' : ''}`} style={c.on ? { '--of-wco-right': `${c.right}px`, '--of-wco-h': `${c.height}px` } : undefined}>
      <button type="button" className="btn ghost icon sm" title={t('Back to Files')} aria-label={t('Back to Files')} onClick={onBack}><ArrowLeft size={16} /></button>
      <FileIcon name="x.docx" size={16} /><b className="small">Office</b>
    </div>
  );
}

/**
 * Screen #/office-edit?path=…[&root=…][&back=#/files?…][&os=1] – the editor in front, without Hostwerk's sidebar (from
 * Files, from Explorer, "Open in its own window"). ?dialog=new&… = the New dialog when Office's dialog host is missing.
 */
export function EditorScreen() {
  const [q, setQ] = useState(() => parseQuery((location.hash.split('?')[1]) || ''));
  const { hasConverter } = useOfficeStatus();
  useEffect(() => { const h = () => setQ(parseQuery((location.hash.split('?')[1]) || '')); window.addEventListener('hashchange', h); return () => window.removeEventListener('hashchange', h); }, []);
  const back = () => { location.replace(q.back && /^#?\/files/.test(q.back) ? `#${String(q.back).replace(/^#/, '')}` : filesRoute({ root: LIB_ROOT, dir: true })); };
  if (q.dialog === 'new') return <div className="of-shell-screen of-shell-screen-dialog"><WindowBar onBack={back} /><DialogView d={{ type: 'new', kind: q.kind || 'docs', gallery: !!q.gallery, root: q.root || LIB_ROOT, dir: q.dir || '' }} hasConverter={hasConverter} onClose={back} /><ThumbRenderer /></div>;
  if (!q.path) return <div className="of-shell-screen"><WindowBar onBack={back} /><div className="of-shell-screen-empty"><TileEmpty icon={FileText} text="No file selected." action="Open Office documents" href={filesRoute({ root: LIB_ROOT, dir: true })} /></div></div>;
  const target = canon({ root: q.root, path: q.path });
  return <div className="of-shell-screen"><EditorHost key={`${target.root || ''}|${target.path}|${q.app || ''}`} target={target} opts={q} bare /></div>;
}
