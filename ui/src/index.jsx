/*
 * Office – native Hostwerk extension, interface module (shell). Office lives in Hostwerk's Files ("Dateien", SDK 3.15)
 * and has no own entry in the sidebar (manifest page.nav: false):
 *   root "Office documents" (server: host.files.addRoot) with the Office section (files.folder), files.open (click opens
 *   the editor in front), files.new ("New" menu), files.actions (right-click), files.marks, files.preview
 *   screen #/office-edit?…        the editor (Back = the folder in Files) · ?dialog=new = New dialog fallback
 *   page #/ext/office/settings     Office settings (also the tab "Office" in Hostwerk's settings)
 *   page #/ext/office/edit?…       the editor inside Hostwerk's frame (older links)
 *   page #/ext/office[?action=…]   Hostwerk's "New" menu, shortcuts, quick search: new · new-docs … · template · upload ·
 *                                  folder · search · trash · open&target= – everything leads into Files
 */
import './lib/toastfix';
import { useEffect } from 'react';
import { FileText, Settings2 } from 'lucide-react';
import EditorHost from './shell/EditorHost';
import OfficeSettings from './shell/Settings';
import { FilesPreview, RecentTile, EditorScreen, officeMatch, filesIcon, filesThumb } from './shell/slots';
import { fileActions, fileMarks, filesNew, OfficeSection, OfficeDialogHost, officeSectionWhen, opensInOffice, openFromFiles, openDialog, sectionIntent } from './shell/files';
import { parseQuery, openTarget } from './shell/util';
import { LIB_ROOT, filesRoute, openEditor, canon } from './shell/ops';
import { KINDS } from './lib/formats';
import './style.css';

const toFiles = (dir = '') => { location.replace(filesRoute({ root: LIB_ROOT, path: dir, dir: true })); };
/** Actions of the page address → into Files (the Office documents root) */
function pageAction(query) {
  const p = parseQuery(query);
  const a = p.action || '';
  if (a === 'open' && p.target) { const f = canon(openTarget(p)); location.replace(`#${filesRoute(f)}`); setTimeout(() => openEditor(f), 0); return; }
  const kind = p.kind || (a.startsWith('new-') ? a.slice(4) : '');
  if (a === 'trash') sectionIntent({ view: 'trash' });
  if (a === 'search') sectionIntent({ search: true });
  toFiles();
  if (a === 'new' || a.startsWith('new-')) setTimeout(() => openDialog({ type: 'new', kind: KINDS[kind] ? kind : 'docs', root: LIB_ROOT, dir: '' }), 50);
  else if (a === 'template' || a === 'templates') setTimeout(() => openDialog({ type: 'new', kind: KINDS[kind] ? kind : 'docs', gallery: true, root: LIB_ROOT, dir: '' }), 50);
}

function OfficePage({ path = [], query = '' }) {
  const sub = Array.isArray(path) ? path[0] : path;
  if (sub === 'edit') {
    const q = parseQuery(query);
    if (q.path) { const target = canon({ root: q.root, path: q.path }); return <div className="page full bleed of-shell-editor-page"><EditorHost key={`${target.root || ''}|${target.path}|${q.app || ''}|${q.ro || ''}`} target={target} opts={q} /></div>; }
  }
  if (sub === 'settings') return <OfficeSettings query={query} />;
  // the former library page: Office now lives in Files
  return <ToFiles query={query} />;
}
function ToFiles({ query }) {
  useEffect(() => { pageAction(query); }, [query]);
  return null;
}

const newAt = (kind) => () => { toFiles(); setTimeout(() => openDialog({ type: 'new', kind, root: LIB_ROOT, dir: '' }), 50); };

export default {
  pages: { main: OfficePage },
  screens: { 'office-edit': { Component: EditorScreen } },
  slots: {
    // click / Enter on an office file in any folder of Files (and files from Explorer, from: 'os')
    'files.open': [{ id: 'office', title: 'Office', icon: FileText, priority: 20, match: { test: (name) => opensInOffice(name) }, open: (x) => openFromFiles(x) }],
    'files.new': filesNew,
    'files.preview': [{ id: 'office', title: 'Office', match: officeMatch, Component: FilesPreview }],
    'files.actions': fileActions,
    'files.marks': fileMarks,
    // SDK 3.16 (feature-detected by Hostwerk: older versions ignore these): coloured icons + first-page thumbnails
    'files.icon': filesIcon,
    'files.thumb': filesThumb,
    'dashboard.tiles': [{ id: 'recent', title: 'Recent documents', icon: FileText, size: 'third', Component: RecentTile }],
    'settings.tabs': [{ id: 'office', title: 'Office', icon: Settings2, Component: OfficeSettings }],
    // renders nothing but Office's dialogs (New, Share, Versions …) wherever they are opened from
    'sidebar.footer': [{ id: 'dialogs', title: 'Office', Component: OfficeDialogHost }],
  },
  sections: {
    'files.folder': [{ id: 'office', title: 'Office', icon: FileText, order: 10, when: officeSectionWhen, Component: OfficeSection }],
  },
  shortcuts: [
    { id: 'open', title: 'Open Office documents', keys: 'ctrl+alt+o', run: () => { location.hash = filesRoute({ root: LIB_ROOT, dir: true }).slice(1); } },
    { id: 'new-document', title: 'New document', run: newAt('docs') },
    { id: 'new-spreadsheet', title: 'New spreadsheet', run: newAt('sheets') },
    { id: 'new-presentation', title: 'New presentation', run: newAt('slides') },
    { id: 'search', title: 'Search in Office', run: () => { sectionIntent({ search: true }); location.hash = filesRoute({ root: LIB_ROOT, dir: true }).slice(1); } },
  ],
};
