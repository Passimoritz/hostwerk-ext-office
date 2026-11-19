/*
 * Office inside the rest of Hostwerk: the preview in Files (files.preview – right-click → "Preview", or every office
 * file when the click is switched off under Settings → Advanced → Open files), a dashboard tile with recent documents
 * (dashboard.tiles) and the editor as a full view (screen office-edit: from Files, from Explorer, own windows).
 * Everything else of Files – files.open / files.new / files.actions / files.marks / the Office section – is in files.jsx.
 */
import { useEffect, useState } from 'react';
import { Pencil, Monitor, FileText, Sparkles } from 'lucide-react';
import { Btn, TileEmpty } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { office } from '../lib/api';
import { OFFICE_RE, formatOf } from '../lib/formats';
import { FilePreviewBox } from './preview';
import { FileIcon, Keep, fmtWhen } from './common';
import { settingsNow, loadSettings, useChanged, useOfficeStatus } from './store';
import { LIB_ROOT, canon, openEditor, openInDesktop, filesRoute } from './ops';
import { parseQuery, basename, fmtBytes } from './util';
import EditorHost from './EditorHost';
import { ensureFullEditor } from './engine';
import { DialogView } from './files';

loadSettings().catch(() => {}); // the slots ask settingsNow() synchronously (openFromFiles)

/** files.preview `match`: a RegExp-like object, so the setting "Open office files from Files" applies live */
export const officeMatch = { test: (p) => settingsNow().general.openFromFiles !== false && OFFICE_RE.test(String(p || '')) };

export function FilesPreview({ root, path }) {
  const name = basename(path);
  const f = canon({ root, path, name });
  const fmt = formatOf(name);
  const [size, setSize] = useState(null);
  useEffect(() => { office.info(f).then((i) => setSize(i && i.size), () => {}); }, [root, path]); // eslint-disable-line react-hooks/exhaustive-deps
  const back = filesRoute({ root, path });
  return (
    <div className="of-shell-fpreview">
      <div className="of-shell-fpreview-bar">
        <FileIcon name={name} />
        <span className="small grow ellipsis"><b><Keep>{name}</Keep></b> <span className="muted tiny">{fmt ? t(fmt.label) : ''}{size != null ? ` · ${fmtBytes(size)}` : ''}</span></span>
        {fmt && fmt.edit !== 'view' && <Btn className="sm primary" icon={Pencil} onClick={() => openEditor(f, { back })}>Edit in Office</Btn>}
        {fmt && fmt.edit === 'view' && <Btn className="sm" icon={FileText} onClick={() => openEditor(f, { back, ro: 1 })}>Open in Office</Btn>}
        <Btn className="sm ghost icon" icon={Sparkles} title="Open in the full editor" aria-label="Open in the full editor" onClick={async () => { const app = await ensureFullEditor(); if (app) openEditor(f, { app, back }); }} />
        <Btn className="sm ghost icon" icon={Monitor} title="Open in the desktop app" aria-label="Open in the desktop app" onClick={() => openInDesktop(f)} />
      </div>
      <FilePreviewBox file={f} editorFallback />
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

/**
 * Screen #/office-edit?path=…[&root=…][&back=#/files?…][&os=1] – the editor in front, without Hostwerk's sidebar (from
 * Files, from Explorer, "Open in its own window"). ?dialog=new&… = the New dialog when Office's dialog host is missing.
 */
export function EditorScreen() {
  const [q, setQ] = useState(() => parseQuery((location.hash.split('?')[1]) || ''));
  const { hasConverter } = useOfficeStatus();
  useEffect(() => { const h = () => setQ(parseQuery((location.hash.split('?')[1]) || '')); window.addEventListener('hashchange', h); return () => window.removeEventListener('hashchange', h); }, []);
  const back = () => { location.replace(q.back && /^#?\/files/.test(q.back) ? `#${String(q.back).replace(/^#/, '')}` : filesRoute({ root: LIB_ROOT, dir: true })); };
  if (q.dialog === 'new') return <div className="of-shell-screen of-shell-screen-dialog"><DialogView d={{ type: 'new', kind: q.kind || 'docs', gallery: !!q.gallery, root: q.root || LIB_ROOT, dir: q.dir || '' }} hasConverter={hasConverter} onClose={back} /></div>;
  if (!q.path) return <div className="of-shell-screen-empty"><TileEmpty icon={FileText} text="No file selected." action="Open Office documents" href={filesRoute({ root: LIB_ROOT, dir: true })} /></div>;
  const target = canon({ root: q.root, path: q.path });
  return <div className="of-shell-screen"><EditorHost key={`${target.root || ''}|${target.path}|${q.app || ''}`} target={target} opts={q} bare /></div>;
}
