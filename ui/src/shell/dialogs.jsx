/*
 * Dialogs of the shell: New file / template gallery, Move/Copy, Share, Info, Tags/colour/note, Versions.
 */
import { useEffect, useMemo, useState } from 'react';
import { FilePlus2, FolderPlus, Link2, Copy, Trash2, Lock, Download, RotateCcw, Tag, History, Pencil, LayoutTemplate, Check, GitCompare } from 'lucide-react';
import { Modal, Btn, Seg, Field, Switch, Loading, Empty, Callout, toast, confirmDialog, promptDialog } from '@hostwerk/ui';
import { t, pl } from '@hostwerk/i18n';
import { get } from '@hostwerk/api';
import { office } from '../lib/api';
import { TEMPLATES, KINDS, FORMATS, extOf } from '../lib/formats';
import { formatChoices, defaultExt, defaultBaseName, validateName, ensureExt, uniqueName, crumbsOf, parseTags, shareExpiry, fmtBytes, COLORS, splitExt, convertTargets } from './util';
import { FolderTree } from './tree';
import { FileIcon, Thumb, Keep, Raw, fmtDateTime, fmtWhen, fmtTime, copyText, fmtNum, uiLang } from './common';
import { settingsNow, changed } from './store';
import { createFile, newFolder, moveEntries, copyEntries, convertEntry } from './ops';
import { FilePreviewBox } from './preview';
import { download } from '../lib/api';
import { loadEditor, prefetchEditor } from '../lib/loader';
import { templateThumb } from '../lib/thumbs';
import { CHUNK_VERSION } from '../lib/version';
import { QrImage } from './features';
const NEW_KINDS = ['docs', 'sheets', 'slides', 'pdf'];

/* ---------------- New file / template gallery ---------------- */
/**
 * Left: the four apps (coloured tiles) and the templates with real thumbnails; right: a live preview of the chosen
 * template, name, format and folder. root: a folder of Hostwerk's Files outside the library (files.new) – then the
 * folder is fixed, rootLabel names it.
 */
export function NewDialog({ kind: startKind = 'docs', dir: startDir = '', root = null, rootLabel = '', templateId = null, gallery = false, siblings = [], onClose, onCreated }) {
  const s = settingsNow();
  const [kind, setKind] = useState(NEW_KINDS.includes(startKind) ? startKind : 'docs');
  const [ext, setExt] = useState(defaultExt(startKind, s.general.newFormat));
  const [name, setName] = useState(t(defaultBaseName(startKind)));
  const [touched, setTouched] = useState(false);
  const [dir, setDir] = useState(startDir);
  const [pick, setPick] = useState(false);
  const [tpl, setTpl] = useState(templateId ? { id: templateId, own: false } : { id: 'blank', own: false });
  const [own, setOwn] = useState(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState(null);
  const [place, setPlace] = useState(rootLabel);
  const [live, setLive] = useState(null); // { key, bytes } – preview of the chosen template
  useEffect(() => { office.templates().then((r) => setOwn(Array.isArray(r) ? r : (r && r.templates) || []), () => setOwn([])); }, []);
  // name of the Files root the file goes to (Hostwerk's list of roots)
  useEffect(() => { if (root && !rootLabel) get('/files/roots').then((l) => { const r = (Array.isArray(l) ? l : []).find((x) => x.id === root); if (r) setPlace(r.label); }, () => {}); }, [root, rootLabel]);
  const switchKind = (k) => {
    setKind(k); setExt(defaultExt(k, s.general.newFormat)); setTpl({ id: 'blank', own: false });
    if (!touched) setName(t(defaultBaseName(k)));
  };
  const choices = formatChoices(kind, s.general.newFormat);
  const builtIn = TEMPLATES.filter((x) => x.kind === kind);
  const mine = (own || []).filter((x) => (x.kind || (FORMATS[x.ext] || {}).kind) === kind);
  const finalExt = tpl.own ? ((mine.find((m) => m.id === tpl.id) || {}).ext || ext) : ext;
  const fileName = ensureExt(name.trim(), finalExt);
  const nameErr = validateName(fileName);
  // live preview of the chosen template: the chunk builds it (built-in) or the server sends it (own)
  const liveKey = `${kind}|${tpl.own ? 'o' : 'b'}:${tpl.id}|${finalExt}`;
  useEffect(() => {
    let alive = true;
    const h = setTimeout(async () => {
      try {
        let bytes;
        if (tpl.own) bytes = await office.templateBytes(tpl.id);
        else { const c = await loadEditor(kind); if (!c || !c.create) return; bytes = await c.create({ ext: finalExt, template: tpl.id, settings: settingsNow() }); }
        if (alive) setLive({ key: liveKey, bytes });
      } catch { if (alive) setLive({ key: liveKey, bytes: null }); }
    }, 120);
    return () => { alive = false; clearTimeout(h); };
  }, [liveKey]); // eslint-disable-line react-hooks/exhaustive-deps
  const create = async () => {
    setBusy(true); setErr(null);
    try {
      const e = await createFile({ kind, ext: finalExt, name: dir === startDir ? uniqueName(fileName, siblings) : fileName, dir, root, template: tpl.own ? { id: tpl.id, own: true, ext: finalExt } : tpl });
      onCreated(e);
    } catch (e) { setErr(e.message || String(e)); setBusy(false); }
  };
  const ownTitle = (x) => x.name;
  return (
    <Modal size="xwide" title={gallery ? 'New from template' : 'New file'} onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" icon={FilePlus2} disabled={!!nameErr || busy} busy={busy} onClick={create}>Create and open</Btn></>}>
      <div className="of-shell-newdlg">
        <div className="stack of-shell-newdlg-left">
          <div className="of-shell-kinds" role="radiogroup" aria-label={t('Kind')}>
            {NEW_KINDS.map((k) => (
              <button key={k} type="button" role="radio" aria-checked={kind === k} className={`of-shell-kind of-shell-k-${k} ${kind === k ? 'on' : ''}`} onClick={() => switchKind(k)} onMouseEnter={() => prefetchEditor(k)}>
                <FileIcon name={`x.${defaultExt(k, s.general.newFormat) || k}`} size={32} /><b>{t(KINDS[k].label)}</b>
              </button>
            ))}
          </div>
          <div className="of-shell-sec-title"><LayoutTemplate size={14} aria-hidden="true" /> Template</div>
          <div className="of-shell-tpls" role="listbox" aria-label={t('Template')}>
            {builtIn.map((x) => <TemplateCard key={`b:${x.id}`} kind={kind} ext={finalExt} tplId={x.id} name={t(x.name)} cat={t(x.category)} on={!tpl.own && tpl.id === x.id} onPick={() => setTpl({ id: x.id, own: false })} onOpen={create} blank={x.id === 'blank'} />)}
            {mine.map((x) => <TemplateCard key={`o:${x.id}`} kind={kind} ext={x.ext} name={ownTitle(x)} cat={x.category || t('My templates')} own on={tpl.own && tpl.id === x.id} onPick={() => setTpl({ id: x.id, own: true })} onOpen={create} onRemove={async () => { if (await confirmDialog({ title: t('Remove the template “{name}”?', { name: x.name }), text: 'Files made from it stay.', confirm: 'Remove', danger: true })) { await office.removeTemplate(x.id); setOwn((o) => o.filter((y) => y.id !== x.id)); if (tpl.id === x.id) setTpl({ id: 'blank', own: false }); } }} />)}
          </div>
          {own && !mine.length && <p className="tiny muted of-shell-p">Own templates: right-click a file → “Save as template”.</p>}
        </div>
        <div className="stack of-shell-newdlg-right">
          <div className={`of-shell-newprev of-shell-k-${kind}`}>
            {live && live.key === liveKey && live.bytes ? <FilePreviewBox file={{ name: `preview.${finalExt}`, path: '', mtime: 0 }} bytes={live.bytes} compact /> : <div className="of-shell-skel-page" aria-busy="true" />}
          </div>
          <Field label="Name"><Raw as="input" className="input" autoFocus value={name} onFocus={(e) => e.target.select()} onChange={(e) => { setName(e.target.value); setTouched(true); }} onKeyDown={(e) => { if (e.key === 'Enter' && !nameErr) create(); }} /></Field>
          {nameErr && name && <span className="small c-err">{t(nameErr)}</span>}
          <Field label="Format" hint={tpl.own ? t('The template decides the format.') : undefined}>
            <select className="input select" value={finalExt} disabled={tpl.own} onChange={(e) => setExt(e.target.value)}>
              {choices.map((c) => <option key={c.ext} value={c.ext}>{`${t(c.label)} (.${c.ext})${c.isDefault ? ` · ${t('default')}` : ''}`}</option>)}
            </select>
          </Field>
          <Field label="Folder">
            {root ? <span className="of-shell-path ellipsis"><Keep>{[place || root, ...crumbsOf(dir).map((c) => c.name)].join(' › ')}</Keep></span> : <>
              <div className="row-tight"><span className="of-shell-path grow ellipsis"><Keep>{[t('Office documents'), ...crumbsOf(dir).map((c) => c.name)].join(' › ')}</Keep></span><Btn className="sm" onClick={() => setPick(!pick)}>{pick ? 'Done' : 'Change …'}</Btn></div>
              {pick && <div className="of-shell-pick"><FolderTree id="new" value={dir} onSelect={setDir} /></div>}
            </>}
          </Field>
          {err && <Callout kind="err" dismissible={false}>{err}</Callout>}
        </div>
      </div>
    </Modal>
  );
}
/** a template with its real first page (built-ins: rendered once and kept in this browser) */
export function TemplateCard({ kind, ext, tplId, name, cat, on, own, blank, onPick, onOpen, onRemove }) {
  const [img, setImg] = useState(null);
  useEffect(() => {
    let alive = true;
    setImg(null);
    if (!own && tplId) templateThumb({ kind, id: tplId, ext, lang: uiLang(), version: CHUNK_VERSION, settings: settingsNow() }).then((u) => { if (alive) setImg(u); });
    return () => { alive = false; };
  }, [kind, tplId, ext, own]);
  return (
    <div className={`of-shell-tpl ${on ? 'on' : ''}`} role="option" aria-selected={on} tabIndex={0} onClick={onPick} onDoubleClick={onOpen} onKeyDown={(e) => { if (e.key === 'Enter') onOpen(); if (e.key === ' ') { e.preventDefault(); onPick(); } }}>
      <div className={`of-shell-tpl-paper of-shell-k-${kind} ${img ? 'has-img' : ''} is-${kind}`}>
        {img ? <img src={img} alt="" decoding="async" /> : blank ? <FilePlus2 size={22} aria-hidden="true" /> : <><FileIcon name={`x.${ext}`} size={20} /><span className="of-shell-tpl-lines" aria-hidden="true"><i /><i /><i /></span></>}
      </div>
      <b className="small ellipsis"><Keep>{name}</Keep></b>
      <span className="tiny muted ellipsis"><Keep>{cat}</Keep></span>
      {on && <span className="of-shell-tpl-check" aria-hidden="true"><Check size={12} /></span>}
      {own && onRemove && <button type="button" className="btn ghost xs icon of-shell-tpl-x" title={t('Remove template')} aria-label={t('Remove template')} onClick={(e) => { e.stopPropagation(); onRemove(); }}><Trash2 size={12} /></button>}
    </div>
  );
}

/* ---------------- Move / copy ---------------- */
export function MoveCopyDialog({ entries, mode = 'move', start = '', onClose }) {
  const [to, setTo] = useState(start);
  const blocked = (d) => mode === 'move' && entries.some((e) => e.dir && (d === e.path || d.startsWith(`${e.path}/`)));
  const go = async () => { if (mode === 'move') await moveEntries(entries, to); else await copyEntries(entries, to); onClose(); };
  const title = mode === 'move' ? (entries.length === 1 ? t('Move “{name}” to …', { name: entries[0].name }) : t('Move {n} items to …', { n: entries.length })) : (entries.length === 1 ? t('Copy “{name}” to …', { name: entries[0].name }) : t('Copy {n} items to …', { n: entries.length }));
  return (
    <Modal title={title} onClose={onClose} footer={<>
      <Btn className="ghost" icon={FolderPlus} onClick={async () => { const e = await newFolder(to); if (e && e.path) setTo(e.path); }}>New folder …</Btn>
      <span className="grow" />
      <Btn onClick={onClose}>Cancel</Btn>
      <Btn className="primary" disabled={blocked(to)} onClick={go}>{mode === 'move' ? 'Move here' : 'Copy here'}</Btn>
    </>}>
      <div className="of-shell-pick is-tall"><FolderTree id={`pick-${mode}`} value={to} onSelect={setTo} disabled={blocked} /></div>
      <p className="small muted of-shell-p">{t('Target: {path}', { path: [t('Library'), ...crumbsOf(to).map((c) => c.name)].join(' › ') })}</p>
    </Modal>
  );
}

/* ---------------- Share ---------------- */
const DAYS = [{ id: '1', label: '1 day' }, { id: '7', label: '7 days' }, { id: '30', label: '30 days' }, { id: '0', label: 'No end' }];
export const shareUrl = (sh) => (sh && (sh.url || sh.link)) || (sh && sh.token ? `${location.origin}/api/office/pub/${sh.token}` : '');
export function ShareDialog({ entry, onClose }) {
  const s = settingsNow();
  const [list, setList] = useState(null);
  const [days, setDays] = useState(String(s.sharing.defaultDays));
  const [usePw, setUsePw] = useState(false);
  const [pw, setPw] = useState('');
  const [allow, setAllow] = useState(!!s.sharing.allowDownload);
  const [made, setMade] = useState(null);
  const load = () => office.shares().then((r) => setList((Array.isArray(r) ? r : (r && r.shares) || []).filter((x) => x.path === entry.path)), () => setList([]));
  useEffect(() => { load(); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  if (!s.sharing.enabled) return <Modal title="Share link" onClose={onClose} footer={<Btn onClick={onClose}>Close</Btn>}><Empty icon={Link2} title="Sharing is off" text="Switch it on under Office settings → Sharing." action={<a className="btn" href="#/ext/office/settings?group=sharing" onClick={onClose}>Open settings</a>} /></Modal>;
  const create = async () => {
    const r = await office.share(entry.root ? entry : entry.path, { expires: shareExpiry(days), password: usePw && pw ? pw : undefined, allowDownload: allow });
    setMade(r); load(); changed();
    const u = shareUrl(r);
    if (u) { try { await navigator.clipboard.writeText(u); toast.success('Link created and copied'); } catch { toast.success('Link created'); } }
  };
  const revoke = async (x) => {
    if (!(await confirmDialog({ title: 'Delete this link?', text: 'Whoever has it can no longer open the file. The file stays.', confirm: 'Delete link', danger: true }))) return;
    await office.unshare(x.id); load(); changed(); if (made && made.id === x.id) setMade(null);
    toast.success('Link deleted');
  };
  return (
    <Modal title={t('Share “{name}”', { name: entry.name })} subtitle="Read-only link – anyone with the link can view the file" onClose={onClose}
      footer={<><Btn onClick={onClose}>Close</Btn><Btn className="primary" icon={Link2} disabled={usePw && pw.length < 4} onClick={create}>Create link</Btn></>}>
      <div className="stack">
        {made && shareUrl(made) && (
          <div className="of-shell-qrbox hw-pop">
            <QrImage value={shareUrl(made)} size={120} />
            <div className="stack of-shell-gap-2 grow of-shell-min0">
              <b className="small">Link created and copied – share it or let someone scan the code</b>
              <div className="of-shell-link"><Check size={16} className="c-ok" aria-hidden="true" /><Raw as="input" className="input mono grow" readOnly value={shareUrl(made)} onFocus={(e) => e.target.select()} aria-label={t('Link')} /><Btn className="sm" icon={Copy} onClick={() => copyText(shareUrl(made))}>Copy</Btn></div>
            </div>
          </div>
        )}
        <Field label="Valid for"><Seg value={days} onChange={setDays} options={DAYS} /></Field>
        <div className="setting-row"><div className="txt"><b>Protect with a password</b><span>Visitors enter it before they see the file.</span></div><Switch checked={usePw} onChange={setUsePw} /></div>
        {usePw && <Field label="Password" hint="At least 4 characters – pass it on separately."><input className="input" type="text" autoComplete="new-password" value={pw} onChange={(e) => setPw(e.target.value)} /></Field>}
        <div className="setting-row"><div className="txt"><b>Allow download</b><span>Otherwise the file can only be viewed.</span></div><Switch checked={allow} onChange={setAllow} /></div>
        <div>
          <div className="of-shell-sec-title"><Link2 size={14} aria-hidden="true" /> Links of this file</div>
          {!list ? <Loading inline /> : !list.length ? <p className="small muted of-shell-p">No link yet.</p> : (
            <ul className="of-shell-rows">{list.map((x) => <ShareRow key={x.id} x={x} onRevoke={() => revoke(x)} />)}</ul>
          )}
        </div>
      </div>
    </Modal>
  );
}
export function ShareRow({ x, onRevoke, showName = false }) {
  const expired = x.expires && x.expires < Date.now();
  return (
    <li className={`of-shell-row ${expired ? 'is-old' : ''}`}>
      {x.password ? <Lock size={14} className="muted" aria-label={t('With password')} /> : <Link2 size={14} className="muted" aria-hidden="true" />}
      <div className="grow of-shell-min0">
        {showName && <b className="small ellipsis">{x.name || x.path}</b>}
        <div className="tiny muted ellipsis">{expired ? t('expired') : x.expires ? t('until {date}', { date: fmtDateTime(x.expires) }) : t('no end date')} · {x.allowDownload ? t('download allowed') : t('view only')}{x.views != null ? ` · ${pl(x.views, '{n} view', '{n} views')}` : ''}</div>
      </div>
      {shareUrl(x) && !expired && <Btn className="xs ghost icon" icon={Copy} title="Copy link" aria-label="Copy link" onClick={() => copyText(shareUrl(x))} />}
      <Btn className="xs ghost icon" icon={Trash2} title="Delete link" aria-label="Delete link" onClick={onRevoke} />
    </li>
  );
}

/* ---------------- Info ---------------- */
export function InfoDialog({ entry, onClose }) {
  const [info, setInfo] = useState(null);
  const [err, setErr] = useState(null);
  useEffect(() => { office.info(entry).then(setInfo, (e) => setErr(e.message)); }, [entry.path, entry.root]);
  const f = FORMATS[extOf(entry.name)] || {};
  const rows = info ? [
    ['Name', info.name || entry.name], ['Type', f.label ? `${t(f.label)} (.${extOf(entry.name)})` : extOf(entry.name)], ['Location', entry.root ? `${entry.root}: ${entry.path}` : `${t('Office documents')} / ${entry.path}`],
    ['Size', fmtBytes(info.size ?? entry.size)], ['Modified', fmtDateTime(info.mtime || entry.mtime)], ['Created', fmtDateTime(info.ctime || info.created || entry.ctime)],
    info.title && ['Title', info.title], info.author && ['Author', info.author], info.modified && ['Last saved in the file', fmtDateTime(Date.parse(info.modified) || info.modified)],
    info.pages != null && ['Pages', fmtNum(info.pages)], info.words != null && ['Words', fmtNum(info.words)], info.sheets != null && ['Sheets', fmtNum(info.sheets)], info.slides != null && ['Slides', fmtNum(info.slides)],
    info.lock && ['Open in', t('{by} since {since}', { by: info.lock.by || '?', since: fmtWhen(info.lock.since) })],
    info.shared && ['Shared', t('yes – read-only link')], (info.tags || []).length > 0 && ['Tags', info.tags.join(', ')], info.note && ['Note', info.note],
  ].filter(Boolean) : [];
  return (
    <Modal title={<Keep>{entry.name}</Keep>} subtitle="File information" onClose={onClose} footer={<Btn onClick={onClose}>Close</Btn>}>
      <div className="of-shell-info">
        <div className="of-shell-info-thumb"><Thumb entry={{ ...entry, mtime: (info && info.mtime) || entry.mtime, thumb: info ? info.thumb : entry.thumb }} /></div>
        {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !info ? <div className="stack of-shell-gap-2 grow">{[80, 60, 70, 50].map((w) => <i key={w} className="sk of-shell-sk-row" style={{ width: `${w}%` }} />)}</div> : (
          <dl className="of-shell-dl grow">{rows.map(([k, v]) => <div key={k} className="of-shell-dl-row"><dt className="small muted">{t(k)}</dt><dd className="small"><Keep>{String(v)}</Keep></dd></div>)}</dl>
        )}
      </div>
    </Modal>
  );
}

/* ---------------- Tags, colour, note ---------------- */
export function TagsDialog({ entries, onClose }) {
  const one = entries.length === 1 ? entries[0] : null;
  const [tags, setTags] = useState(one ? [...(one.tags || [])] : []);
  const [input, setInput] = useState('');
  const [color, setColor] = useState(one ? one.color || '' : null);
  const [note, setNote] = useState(one ? one.note || '' : '');
  const [all, setAll] = useState([]);
  useEffect(() => { office.tags().then((r) => setAll((Array.isArray(r) ? r : (r && r.tags) || []).map((x) => (typeof x === 'string' ? x : x.tag || x.name)).filter(Boolean)), () => {}); }, []);
  const add = (v) => { const n = parseTags(v); if (n.length) setTags((cur) => parseTags([...cur, ...n].join(','))); setInput(''); };
  // optimistic: the dialog closes at once, the lists update when the server confirmed (an error comes as a toast)
  const save = () => {
    const pending = parseTags(input);
    onClose();
    (async () => {
      try {
        for (const e of entries) {
          const next = one ? parseTags([...tags, ...pending].join(',')) : parseTags([...(e.tags || []), ...tags, ...pending].join(','));
          await office.meta(e.root ? e : e.path, { tags: next, ...(color !== null ? { color } : {}), ...(one ? { note } : {}) });
        }
        changed();
        try { window.dispatchEvent(new Event('hw:files-marks')); } catch { /* old browser */ }
      } catch (e) { toast.error(e); }
    })();
  };
  const suggest = all.filter((x) => !tags.some((y) => y.toLowerCase() === x.toLowerCase()) && (!input || x.toLowerCase().includes(input.toLowerCase()))).slice(0, 12);
  return (
    <Modal title={one ? t('Tags and note – {name}', { name: one.name }) : t('Tags for {n} items', { n: entries.length })} onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" onClick={save}>Save</Btn></>}>
      <div className="stack">
        <Field label="Tags" hint={one ? 'Enter or comma adds a tag.' : 'The tags are added to every selected item.'}>
          <div className="of-shell-tagbox">
            {tags.map((x) => <span key={x} className="chip"><Keep>{x}</Keep><button type="button" aria-label={t('Remove tag {tag}', { tag: x })} title={t('Remove')} onClick={() => setTags(tags.filter((y) => y !== x))}>×</button></span>)}
            <input className="of-shell-taginput" value={input} placeholder={tags.length ? '' : t('e.g. invoices, 2026')} onChange={(e) => { if (/[,;]$/.test(e.target.value)) add(e.target.value); else setInput(e.target.value); }} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); add(input); } if (e.key === 'Backspace' && !input && tags.length) setTags(tags.slice(0, -1)); }} aria-label={t('New tag')} />
          </div>
          {suggest.length > 0 && <div className="row-tight of-shell-wrap">{suggest.map((x) => <button key={x} type="button" className="chip-toggle" onClick={() => add(x)}><Tag size={11} aria-hidden="true" /><Keep>{x}</Keep></button>)}</div>}
        </Field>
        <Field label="Colour">
          <div className="row-tight of-shell-wrap" role="radiogroup" aria-label={t('Colour')}>
            {COLORS.map((c) => <button key={c || 'none'} type="button" role="radio" aria-checked={color === c} className={`of-shell-swatch of-shell-c-${c || 'none'} ${color === c ? 'on' : ''}`} title={t(c ? c[0].toUpperCase() + c.slice(1) : 'No colour')} aria-label={t(c ? c[0].toUpperCase() + c.slice(1) : 'No colour')} onClick={() => setColor(c)} />)}
          </div>
        </Field>
        {one && <Field label="Note"><textarea className="input" rows={3} value={note} maxLength={2000} onChange={(e) => setNote(e.target.value)} placeholder={t('Only you see it – e.g. “send to the tax office by Friday”')} /></Field>}
      </div>
    </Modal>
  );
}

/* ---------------- Versions ---------------- */
/**
 * Timeline on the left, the chosen version and the current file side by side on the right (wide windows; narrow: the
 * version only). onCompare(version) – the editor's compare view (host.compare); Files: opens the editor with it.
 */
export function VersionsSheet({ file, onClose, onRestored, beforeRestore, onCompare }) {
  const [list, setList] = useState(null);
  const [err, setErr] = useState(null);
  const [cur, setCur] = useState(null);
  const [bytes, setBytes] = useState(null);
  const [now, setNow] = useState(null);
  const [side, setSide] = useState(() => { try { return localStorage.getItem('office:versions-side') !== '0'; } catch { return true; } });
  const load = () => office.versions(file).then((r) => { const l = Array.isArray(r) ? r : (r && r.versions) || []; const s = [...l].sort((a, b) => b.time - a.time); setList(s); setCur((c) => c || s[0] || null); }, (e) => setErr(e.message));
  useEffect(() => { load(); }, [file.path, file.root]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { setBytes(null); if (cur) office.versionBytes(file, cur.id).then(setBytes, (e) => toast.error(e)); }, [cur && cur.id]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (side && !now) office.read(file).then((r) => setNow(r.bytes), () => {}); }, [side]); // eslint-disable-line react-hooks/exhaustive-deps
  const toggleSide = () => { const v = !side; setSide(v); try { localStorage.setItem('office:versions-side', v ? '1' : '0'); } catch { /* this window only */ } };
  const restore = async (v) => {
    if (beforeRestore && !(await beforeRestore())) return;
    if (!(await confirmDialog({ title: t('Restore the version from {when}?', { when: fmtDateTime(v.time) }), text: 'The current content is kept as a version too, so you can go back.', confirm: 'Restore' }))) return;
    await office.restoreVersion(file, v.id);
    toast.success('Version restored');
    changed();
    if (onRestored) onRestored(v); else load();
  };
  const label = async (v) => {
    const l = await promptDialog({ title: 'Name this version', label: 'Name', value: v.label || '', placeholder: t('e.g. Sent to the customer'), confirm: 'Save' });
    if (l == null) return;
    // optimistic: the name shows at once, the server follows
    setList((xs) => xs.map((x) => (x.id === v.id ? { ...x, label: l.trim() } : x)));
    if (cur && cur.id === v.id) setCur({ ...cur, label: l.trim() });
    try { await office.labelVersion(file, v.id, l.trim()); } catch (e) { toast.error(e); load(); }
  };
  const dl = async (v) => { const b = await office.versionBytes(file, v.id); const { base, ext } = splitExt(file.name); download(b, `${base} (${fmtDateTime(v.time).replace(/[/:]/g, '-')}).${ext}`); };
  // group the timeline by day
  const days = useMemo(() => { const m = []; for (const v of list || []) { const d = new Date(v.time).toDateString(); const g = m[m.length - 1]; if (g && g.d === d) g.items.push(v); else m.push({ d, when: v.time, items: [v] }); } return m; }, [list]);
  return (
    <Modal size="xwide" title="Versions" subtitle={<Keep>{file.name}</Keep>} onClose={onClose}
      footer={cur ? <>
        <Btn className="ghost" icon={Pencil} onClick={() => label(cur)}>Name …</Btn>
        <Btn className="ghost" icon={Download} onClick={() => dl(cur)}>Download</Btn>
        <span className="grow" />
        <Btn onClick={onClose}>Close</Btn>
        {onCompare && <Btn icon={GitCompare} onClick={() => onCompare(cur)}>Compare with now</Btn>}
        <Btn className="primary" icon={RotateCcw} onClick={() => restore(cur)}>Restore</Btn>
      </> : <Btn onClick={onClose}>Close</Btn>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !list ? <div className="of-shell-versions"><div className="of-shell-timeline">{[0, 1, 2, 3].map((i) => <i key={i} className="sk of-shell-sk-row" />)}</div><div className="of-shell-skel-page" /></div> : !list.length ? (
        <Empty icon={History} title="No versions yet" text={t('Every save keeps the previous content as a version ({n} at most, for {days} days).', { n: settingsNow().general.keepVersions, days: settingsNow().general.versionDays })} />
      ) : (
        <div className="of-shell-versions">
          <ul className="of-shell-timeline" role="listbox" aria-label={t('Versions')}>
            {days.map((g) => (
              <li key={g.d} className="of-shell-tl-day">
                <span className="of-shell-tl-head tiny muted">{fmtWhen(g.when).replace(/\s+\d{1,2}:\d{2}.*$/, '') || fmtDateTime(g.when)}</span>
                <ul>
                  {g.items.map((v) => (
                    <li key={v.id} role="option" aria-selected={cur && cur.id === v.id} tabIndex={0} className={cur && cur.id === v.id ? 'on' : ''} onClick={() => setCur(v)}
                      onKeyDown={(e) => { if (e.key === 'Enter') setCur(v); if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); const i = list.indexOf(v) + (e.key === 'ArrowDown' ? 1 : -1); if (list[i]) { setCur(list[i]); const el = e.currentTarget.closest('.of-shell-timeline').querySelectorAll('[role=option]')[i]; if (el) el.focus(); } } }}>
                      <span className="of-shell-tl-dot" aria-hidden="true" />
                      <span className="stack of-shell-gap-0 of-shell-min0">
                        <b className="small">{fmtTime(v.time)}{v.label ? <Raw as="span" className="badge accent" title={v.label}>{v.label}</Raw> : null}{v === list[0] && <span className="badge">newest</span>}</b>
                        <span className="tiny muted">{fmtBytes(v.size)} · {v.auto ? t('autosave') : t('saved')}</span>
                      </span>
                    </li>
                  ))}
                </ul>
              </li>
            ))}
          </ul>
          <div className="of-shell-vside">
            <div className="of-shell-vbar">
              <span className="small grow"><b>{cur ? t('Version of {when}', { when: fmtDateTime(cur.time) }) : ''}</b></span>
              <label className="row-tight small of-shell-vtoggle"><Switch checked={side} onChange={toggleSide} />{t('Next to the current file')}</label>
            </div>
            <div className={`of-shell-vpanes ${side ? 'is-side' : ''}`}>
              <div className="of-shell-vpane">{side && <span className="of-shell-vlabel tiny">{t('Version')}</span>}{bytes ? <FilePreviewBox file={{ ...file, mtime: cur && cur.time }} bytes={bytes} editorFallback /> : <div className="of-shell-skel-page" aria-busy="true" />}</div>
              {side && <div className="of-shell-vpane"><span className="of-shell-vlabel tiny is-now">{t('Now')}</span>{now ? <FilePreviewBox file={file} bytes={now} editorFallback /> : <div className="of-shell-skel-page" aria-busy="true" />}</div>}
            </div>
          </div>
        </div>
      )}
    </Modal>
  );
}

/* ---------------- Convert to … ---------------- */
export function ConvertDialog({ entry, hasConverter, onClose }) {
  const targets = convertTargets(extOf(entry.name), { hasConverter: true });
  return (
    <Modal title={t('Convert “{name}”', { name: entry.name })} subtitle="A copy in the new format is saved next to the file" onClose={onClose} footer={<Btn onClick={onClose}>Cancel</Btn>}>
      <div className="of-shell-convert">
        {targets.map((x) => (
          <button key={x.ext} type="button" className="of-shell-convert-btn" onClick={() => { onClose(); convertEntry(entry, x.ext, { via: x.via === 'client' && !hasConverter ? 'client' : 'server' }).catch((e) => toast.error(e)); }}>
            <FileIcon name={`x.${x.ext}`} size="md" />
            <span className="stack of-shell-gap-0"><b className="small">.{x.ext}</b><span className="tiny muted">{t(x.label)}</span>{x.via === 'server' && !hasConverter && <span className="tiny muted">{t('needs the Office engine')}</span>}</span>
          </button>
        ))}
      </div>
    </Modal>
  );
}
