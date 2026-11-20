/*
 * Templates like Word/Excel/PowerPoint/LibreOffice (1.3 – server/templates.js, server/routesTpl.js):
 *   SaveTemplateDialog   "Save as template …" – name, category (sub folder), where (template folder / next to the file /
 *                        both), "use for new …" – writes a REAL template (dotx, xltx, potx, ott, ots, otp …)
 *   TemplateGallery      the Templates view of the Office section: grouped by app, real thumbnails, default badge,
 *                        right-click: new from it, edit, rename, default, show in Files, delete
 *   TemplatesPane        Settings → Templates: own folder, where "Save as template" writes, default per app, imported
 *                        folders (read-only – e.g. %APPDATA%\Microsoft\Templates, LibreOffice) with suggestions
 *   openTemplateFile()   double-click on a template FILE in Files: a new document from it (like Word), opened at once
 */
import { useEffect, useMemo, useState } from 'react';
import { LayoutTemplate, FilePlus2, Pencil, Star, StarOff, FolderOpen, Trash2, Plus, X, FolderInput, Check, Lock } from 'lucide-react';
import { Modal, Btn, Field, Seg, Switch, Callout, Empty, Loading, SettingRow, PathInput, confirmDialog, promptDialog } from '@hostwerk/ui';
import { contextMenu } from '@hostwerk/kit';
import { t } from '@hostwerk/i18n';
import { office } from '../lib/api';
import { toast } from '../lib/toast';
import { FORMATS, KINDS, extOf, templateExtFor, DOC_OF } from '../lib/formats';
import { KIND_COLOR } from '../lib/fileTypes';
import { ownTemplateThumb } from '../lib/thumbs';
import { FileIcon, Keep, Raw } from './common';
import { settingsNow, changed } from './store';
import { splitExt, basename } from './util';
import { registerFeature } from './features';

const APP_KINDS = [['docs', 'Documents'], ['sheets', 'Spreadsheets'], ['slides', 'Presentations']];
const kindOfTpl = (x) => x.kind || (FORMATS[x.ext] || {}).kind || 'docs';
const listOf = (r) => (Array.isArray(r) ? r : (r && r.templates) || []);
const trItems = (items) => items.filter(Boolean).map((x) => (x && typeof x === 'object' && typeof x.label === 'string' ? { ...x, label: t(x.label) } : x));

/* ---------------- Save as template ---------------- */
function SaveTemplateDialog({ entry, onDone }) {
  const s = settingsNow();
  const ext = extOf(entry.name || entry.path);
  const kind = (FORMATS[ext] || {}).kind;
  const tplExt = templateExtFor(entry.name || entry.path);
  const [name, setName] = useState(splitExt(entry.name || basename(entry.path)).base);
  const [cat, setCat] = useState('');
  const [to, setTo] = useState((s.templates && s.templates.saveTo) || 'folder');
  const [asDefault, setAsDefault] = useState(false);
  const [cats, setCats] = useState([]);
  const [dir, setDir] = useState('');
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState(null);
  useEffect(() => { office.templates().then((r) => { setCats([...new Set(listOf(r).filter((x) => x.source === 'own' && x.category).map((x) => x.category))]); setDir((r && r.dir) || ''); }, () => {}); }, []);
  const fmtLabel = tplExt && FORMATS[tplExt] ? t(FORMATS[tplExt].label) : '';
  const save = async () => {
    if (!name.trim()) return;
    setBusy(true); setErr(null);
    try {
      const r = await office.saveTemplate(entry, name.trim(), cat.trim(), to);
      if (asDefault && r && r.template) await office.defaultTemplate(r.template.id, true);
      changed({ templates: true });
      try { window.dispatchEvent(new Event('hw:files-reload')); } catch { /* old browser */ }
      toast.success(t('Template “{name}” saved', { name: name.trim() }));
      onDone(r);
    } catch (e) { setErr(e.message || String(e)); setBusy(false); }
  };
  return (
    <Modal title="Save as template" onClose={() => onDone(null)}
      footer={<><Btn onClick={() => onDone(null)}>Cancel</Btn><Btn className="primary" icon={LayoutTemplate} disabled={!name.trim() || busy} busy={busy} onClick={save}>Save template</Btn></>}>
      <div className="stack">
        <p className="small muted of-shell-p">{t('Saved as a real template ({format}, .{ext}) – Word, Excel, PowerPoint and LibreOffice open it as a template, and a double-click creates a new document from it.', { format: fmtLabel, ext: tplExt || '' })}</p>
        <Field label="Template name"><Raw as="input" className="input" autoFocus value={name} onFocus={(e) => e.target.select()} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') save(); }} /></Field>
        <Field label="Where">
          <Seg value={to} onChange={setTo} options={[{ id: 'folder', label: t('Template folder') }, { id: 'next', label: t('Next to the file') }, { id: 'both', label: t('Both') }]} />
        </Field>
        {to !== 'next' && <>
          <Field label="Category (optional)" hint={dir ? t('A sub folder of {dir}', { dir }) : undefined}>
            <Raw as="input" className="input" list="of-shell-tplcats" value={cat} placeholder={t('e.g. Letters')} onChange={(e) => setCat(e.target.value)} />
            <datalist id="of-shell-tplcats">{cats.map((c) => <option key={c} value={c} />)}</datalist>
          </Field>
          {kind && KINDS[kind] && ['docs', 'sheets', 'slides'].includes(kind) && <SettingRow title={t('Use for new blank {kind}', { kind: t(KINDS[kind].plural) })} text={t('Like Normal.dotm in Word: “New” starts from this template’s styles and page setup.')}><Switch checked={asDefault} onChange={setAsDefault} /></SettingRow>}
        </>}
        {err && <Callout kind="err" dismissible={false}>{err}</Callout>}
      </div>
    </Modal>
  );
}
registerFeature('save-template', SaveTemplateDialog);

/* ---------------- a template file opened from Files ---------------- */
/** like Word: a new document from the template, in the template's folder (template folders: the library root) – opened at once */
export async function openTemplateFile(f, { openEditor, filesRoute, editTemplate }) {
  const e = await office.fromTemplate(f);
  changed({ created: e && e.path });
  try { window.dispatchEvent(new Event('hw:files-reload')); } catch { /* old browser */ }
  toast(t('New document “{name}” from the template', { name: e.name }), 'success', 9000, { action: { label: t('Edit the template instead'), onClick: () => editTemplate(f) } });
  openEditor(e, { back: filesRoute(e) });
  return e;
}

/* ---------------- gallery (Office section → Templates) ---------------- */
/** a listed template with its real first page */
export function OwnTemplateCard({ x, on = false, onPick, onOpen, onMenu, compact = false }) {
  const kind = kindOfTpl(x);
  const [img, setImg] = useState(null);
  useEffect(() => { let alive = true; setImg(null); ownTemplateThumb(x).then((u) => { if (alive) setImg(u); }); return () => { alive = false; }; }, [x.id, x.mtime]); // eslint-disable-line react-hooks/exhaustive-deps
  const keyMenu = (ev) => { if (onMenu && (ev.key === 'ContextMenu' || (ev.shiftKey && ev.key === 'F10'))) { ev.preventDefault(); const r = ev.currentTarget.getBoundingClientRect(); onMenu({ preventDefault() {}, clientX: r.left + 24, clientY: r.top + 24 }); } };
  return (
    <div className={`of-shell-tpl ${on ? 'on' : ''} ${compact ? '' : 'is-own'}`} role="option" aria-selected={on} tabIndex={0} title={x.full || x.name}
      onClick={onPick} onDoubleClick={onOpen} onContextMenu={onMenu} onKeyDown={(e) => { if (e.key === 'Enter') onOpen(); else if (e.key === ' ') { e.preventDefault(); onPick(); } else keyMenu(e); }}>
      <div className={`of-shell-tpl-paper of-shell-k-${kind} ${img ? 'has-img' : ''} is-${kind}`} style={{ '--of-k': KIND_COLOR[kind] }}>
        {img ? <img src={img} alt="" decoding="async" /> : <><FileIcon name={`x.${x.ext}`} size={20} /><span className="of-shell-tpl-lines" aria-hidden="true"><i /><i /><i /></span></>}
        {x.default && <span className="of-shell-tpl-badge" title={t('Default for new files')}><Star size={11} aria-hidden="true" />{t('Default')}</span>}
        {x.readOnly && <span className="of-shell-tpl-ro" title={t('From an imported folder – read-only')}><Lock size={11} aria-hidden="true" /></span>}
      </div>
      <b className="small ellipsis"><Keep>{x.name}</Keep></b>
      <span className="tiny muted ellipsis"><Keep>{[x.category, x.source === 'folder' ? x.folderName : '', `.${x.ext}`].filter(Boolean).join(' · ')}</Keep></span>
      {on && <span className="of-shell-tpl-check" aria-hidden="true"><Check size={12} /></span>}
    </div>
  );
}

/** grouped by app; actions through `act` = { create(x), edit(x), reveal(x) } from the section */
export function TemplateGallery({ list, act, reload }) {
  const groups = APP_KINDS.map(([k, label]) => [k, label, list.filter((x) => kindOfTpl(x) === k)]).filter(([, , l]) => l.length);
  const menu = (ev, x) => {
    ev.preventDefault();
    const kind = kindOfTpl(x);
    contextMenu(ev, trItems([
      { label: 'New document from this template', icon: FilePlus2, onClick: () => act.create(x) },
      x.path != null && { label: 'Edit template', icon: Pencil, onClick: () => act.edit(x) },
      '-',
      x.default ? { label: 'Stop using as default', icon: StarOff, onClick: () => setDefault(x, false) } : { label: t('Use for new blank {kind}', { kind: t(KINDS[kind].plural) }), icon: Star, onClick: () => setDefault(x, true) },
      !x.readOnly && { label: 'Rename …', icon: Pencil, onClick: () => rename(x) },
      x.path != null && { label: 'Show in Files', icon: FolderOpen, onClick: () => act.reveal(x) },
      !x.readOnly && '-',
      !x.readOnly && { label: 'Delete template …', icon: Trash2, danger: true, onClick: () => remove(x) },
    ]));
  };
  const setDefault = async (x, on) => { try { await office.defaultTemplate(x.id, on); changed({ templates: true }); reload(); toast.success(on ? t('New {kind} start from “{name}”', { kind: t(KINDS[kindOfTpl(x)].plural), name: x.name }) : t('New files start blank again')); } catch (e) { toast.error(e); } };
  const rename = async (x) => {
    const n = await promptDialog({ title: 'Rename template', label: 'Name', value: x.name, confirm: 'Rename', validate: (v) => (v.trim() ? null : t('Please enter a name.')) });
    if (!n) return;
    try { await office.renameTemplate(x.id, n.trim()); changed({ templates: true }); reload(); } catch (e) { toast.error(e); }
  };
  const remove = async (x) => {
    if (!(await confirmDialog({ title: t('Delete the template “{name}”?', { name: x.name }), text: x.path != null ? t('It goes to the Office trash. Files made from it stay.') : t('Files made from it stay.'), confirm: 'Delete', danger: true }))) return;
    try { await office.removeTemplate(x.id); changed({ templates: true }); reload(); } catch (e) { toast.error(e); }
  };
  return (
    <div className="stack of-shell-tplgal">
      {groups.map(([k, label, l]) => (
        <section key={k} className="stack of-shell-gap-1" aria-label={t(label)}>
          <div className="of-shell-sec-title" style={{ '--of-k': KIND_COLOR[k] }}><FileIcon name={`x.${{ docs: 'dotx', sheets: 'xltx', slides: 'potx' }[k]}`} size={16} />{t(label)}<span className="tiny muted">{l.length}</span></div>
          <div className="of-shell-tpls is-wide" role="listbox" aria-label={t(label)}>
            {l.map((x) => <OwnTemplateCard key={x.id} x={x} onPick={() => act.create(x)} onOpen={() => act.create(x)} onMenu={(ev) => menu(ev, x)} />)}
          </div>
        </section>
      ))}
    </div>
  );
}

/* ---------------- Settings → Templates ---------------- */
export function TemplatesPane({ v, set }) {
  const [info, setInfo] = useState(null);
  const [sugg, setSugg] = useState(null);
  const [add, setAdd] = useState('');
  const load = () => office.templates().then((r) => setInfo({ list: listOf(r), dir: r.dir || '', folders: r.folders || [] }), () => setInfo({ list: [], dir: '', folders: [] }));
  useEffect(() => { load(); office.templateFolders().then((r) => setSugg((r && r.folders) || []), () => setSugg([])); }, []);
  const folders = v('templates.folders') || [];
  const addFolder = (p) => { const x = String(p || '').trim(); if (!x) return; if (!/^([a-zA-Z]:[\\/]|\/|\\\\)/.test(x)) { toast.error(t('Enter a full path, e.g. D:\\Documents\\Office.')); return; } if (!folders.some((f) => f.toLowerCase() === x.toLowerCase())) set('templates.folders')([...folders, x]); setAdd(''); };
  const byKind = useMemo(() => Object.fromEntries(APP_KINDS.map(([k]) => [k, ((info && info.list) || []).filter((x) => kindOfTpl(x) === k)])), [info]);
  const known = (p) => ((info && info.folders) || []).find((f) => f.path.toLowerCase() === String(p).toLowerCase());
  return <>
    <div className="card"><div className="card-head"><LayoutTemplate size={16} className="of-shell-chead" aria-hidden="true" /><h3>My templates</h3></div><div className="card-body stack">
      <Field label="Template folder" hint={t('Empty = the folder “Vorlagen” in the Office documents ({dir}). Sub folders become categories.', { dir: (info && info.dir) || '…' })}>
        <PathInput mode="folder" title={t('Template folder')} value={v('templates.dir') || ''} onChange={set('templates.dir')} placeholder={(info && info.dir) || t('Vorlagen in the Office documents')} />
      </Field>
      <SettingRow title="“Save as template” saves" text="In the template folder it appears under “New”; next to the file it stays where you work.">
        <Seg value={v('templates.saveTo') || 'folder'} onChange={set('templates.saveTo')} options={[{ id: 'folder', label: t('In the template folder') }, { id: 'next', label: t('Next to the file') }, { id: 'both', label: t('Both') }]} />
      </SettingRow>
    </div></div>
    <div className="card"><div className="card-head"><Star size={16} className="of-shell-chead" aria-hidden="true" /><h3>Default templates</h3></div><div className="card-body">
      <p className="small muted of-shell-p">{t('Like Normal.dotm in Word: a new blank file starts from this template – its styles, fonts and page setup.')}</p>
      {APP_KINDS.map(([k, label]) => (
        <SettingRow key={k} title={t(label)}>
          <select className="input select of-shell-select" value={v(`templates.defaults.${k}`) || ''} onChange={(e) => set(`templates.defaults.${k}`)(e.target.value)} aria-label={t(label)}>
            <option value="">{t('Blank (Office’s own)')}</option>
            {(byKind[k] || []).map((x) => <option key={x.id} value={x.id}>{`${x.name} (.${x.ext})${x.category ? ` · ${x.category}` : ''}`}</option>)}
          </select>
        </SettingRow>
      ))}
    </div></div>
    <div className="card"><div className="card-head"><FolderInput size={16} className="of-shell-chead" aria-hidden="true" /><h3>Template folders of other apps</h3></div><div className="card-body stack">
      <p className="small muted of-shell-p">{t('Templates in these folders appear in the gallery – used where they are (read-only, nothing is copied). E.g. Word’s personal templates, “Custom Office Templates” or LibreOffice’s template folder.')}</p>
      {folders.length > 0 && <ul className="of-shell-rows">{folders.map((p) => { const k = known(p); return (
        <li key={p} className="of-shell-row">
          <FolderOpen size={16} aria-hidden="true" />
          <div className="grow of-shell-min0"><div className="small ellipsis mono"><Keep>{p}</Keep></div><div className="tiny muted">{k ? (k.exists ? t('{n} templates', { n: k.count }) : t('Folder not found')) : t('Save to apply')}</div></div>
          <button type="button" className="btn ghost xs icon" title={t('Remove')} aria-label={t('Remove')} onClick={() => set('templates.folders')(folders.filter((x) => x !== p))}><X size={12} /></button>
        </li>); })}</ul>}
      <div className="row-tight">
        <div className="grow"><PathInput mode="folder" title={t('Add a template folder')} value={add} onChange={setAdd} placeholder={t('Folder with templates …')} /></div>
        <Btn className="sm" icon={Plus} disabled={!add.trim()} onClick={() => addFolder(add)}>Add</Btn>
      </div>
      {!sugg ? <Loading inline /> : sugg.filter((x) => !folders.some((f) => f.toLowerCase() === x.path.toLowerCase())).length > 0 && (
        <div className="stack of-shell-gap-1">
          <span className="tiny muted">{t('Found on this computer:')}</span>
          {sugg.filter((x) => !folders.some((f) => f.toLowerCase() === x.path.toLowerCase())).map((x) => (
            <div key={x.path} className="row-tight of-shell-wrap">
              <span className="small grow of-shell-min0"><b>{t(x.label)}</b> <span className="tiny muted mono ellipsis"><Keep>{x.path}</Keep></span> <span className="tiny muted">{t('{n} templates', { n: x.count })}</span></span>
              <Btn className="xs" icon={Plus} onClick={() => addFolder(x.path)}>Add</Btn>
            </div>
          ))}
        </div>
      )}
      {sugg && sugg.length === 0 && folders.length === 0 && <span className="tiny muted">{t('No template folders of Word or LibreOffice found on this computer.')}</span>}
    </div></div>
  </>;
}

export { DOC_OF };
