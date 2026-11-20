/*
 * File operations of the library (used by the page, the context menu, the editor host and the file-browser slots).
 * Every change ends with changed() so open lists refresh; destructive steps offer "Undo" where the backend allows it.
 */
import { confirmDialog, promptDialog } from '@hostwerk/ui';
import * as HAPI from '@hostwerk/api';
import { toast } from '../lib/toast';
import { t, pl } from '@hostwerk/i18n';
import { office } from '../lib/api';
import { loadEditor, warmEditor, kindFor } from '../lib/loader';
import { markOpen } from '../lib/perf';
import { FORMATS, extOf } from '../lib/formats';
import { changed, settingsNow, loadStatus, hasConverterIn } from './store';
import { basename, dirname, ensureExt, splitExt, validateName, fileKey, toQuery } from './util';
import { downloadFrom } from './common';
import { resetTree } from './tree';
import { ensureEngine } from './engine';

export const EXT_ID = 'office';
/** The Office library as a root of Hostwerk's Files (server: host.files.addRoot({ id: 'library' })) */
export const LIB_ROOT = `ext:${EXT_ID}:library`;
export const libRoute = (extra = '') => `/ext/${EXT_ID}${extra}`;
/** A file of the library root is a library file (same favourites, versions, shares) – { root: LIB_ROOT, path } → { path } */
export const canon = (f) => (f && f.root === LIB_ROOT ? { ...f, root: undefined, path: f.path || '' } : f);
/** Address of a folder in Hostwerk's Files (with the file marked) – library files live in the Office root */
export function filesRoute(f = {}, { select = true } = {}) {
  const isDir = !!f.dir;
  const p = String(f.path || '');
  const dir = isDir ? p : p.split('/').slice(0, -1).join('/');
  const q = new URLSearchParams();
  q.set('root', f.root || LIB_ROOT);
  if (dir) q.set('path', dir);
  if (select && !isDir && p) q.set('select', p.split('/').pop());
  return `#/files?${q.toString()}`;
}
/** The editor as a full view (screen) – Back returns to `back` (the folder in Files) */
export const screenRoute = (f, extra = {}) => `/office-edit?${toQuery({ ...(f.root ? { root: f.root } : {}), path: f.path, back: filesRoute(f), ...extra })}`;
/**
 * Page change with the View Transitions API (Files ↔ editor: a short zoom/fade, 160 ms) – plain change where the API
 * is missing or motion is reduced. The new view counts as there when the editor host mounted (event office:view) or
 * after 400 ms at the latest.
 */
export function viewTransition(update, type = 'office-open') {
  const reduced = typeof matchMedia === 'function' && matchMedia('(prefers-reduced-motion: reduce)').matches;
  if (reduced || typeof document === 'undefined' || typeof document.startViewTransition !== 'function' || document.hidden) { update(); return; }
  // the editor host announces itself (office:view); back to Files: React renders the folder within a few ms
  const ready = () => new Promise((res) => { const h = () => { window.removeEventListener('office:view', h); res(); }; window.addEventListener('office:view', h); setTimeout(h, type === 'office-open' ? 300 : 80); });
  try { document.startViewTransition({ update: async () => { update(); await ready(); }, types: [type] }); } catch {
    try { document.startViewTransition(async () => { update(); await ready(); }); } catch { update(); }
  }
}
export function openEditor(f, extra = {}) {
  const c = canon(f);
  markOpen(c.name || basename(c.path));
  warmEditor(kindFor(c.name || c.path));
  viewTransition(() => { location.hash = screenRoute(c, extra); }, 'office-open');
}
// names of Hostwerk's Files roots ("Webprojekte (htdocs)") – breadcrumb of the editor, places in lists
let rootNames = null;
export async function rootLabel(root) {
  if (!root || root === LIB_ROOT) return t('Office documents');
  if (!rootNames) rootNames = HAPI.get('/files/roots').then((l) => new Map((Array.isArray(l) ? l : []).map((r) => [r.id, r.label])), () => new Map());
  const m = await rootNames;
  return m.get(root) || (/^opened:/.test(root) ? '' : root);
}
const errMsg = (e) => (e && e.message) || String(e);

/* ---------------- New files ---------------- */
const HTML_BLANK = '<!doctype html><html><head><meta charset="utf-8"><title></title></head><body><p></p></body></html>';
/** Bytes of an empty file when the editor chunk is not there: plain formats directly, the rest via the converter */
async function blankBytes(kind, ext) {
  const enc = (s) => new TextEncoder().encode(s);
  if (['txt', 'md', 'markdown', 'csv', 'tsv'].includes(ext)) return enc('');
  if (['html', 'htm'].includes(ext)) return enc(HTML_BLANK);
  if (ext === 'rtf') return enc('{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Calibri;}}\\f0\\fs22 \\par}');
  const st = await loadStatus().catch(() => null);
  if (hasConverterIn(st)) {
    if (kind === 'docs' || kind === 'pdf') return office.convert(enc(HTML_BLANK), 'html', ext);
    if (kind === 'sheets') return office.convert(enc(' \n'), 'csv', ext);
  }
  ensureEngine(t('The Office engine can create {ext} files while the built-in editor for them is not available.', { ext: ext.toUpperCase() }));
  const e = new Error(t('The editor for this kind of file is not available yet, so Office cannot create an empty {ext} file right now. Set up the Office engine (it asks you once) or try again after the update.', { ext: ext.toUpperCase() }));
  e.code = 'NO_CHUNK';
  throw e;
}
/**
 * Creates a file and returns its entry. template: { id, own } – own templates come from the backend, built-ins from the
 * chunk (create({ ext, template, settings })).
 */
export async function createFile({ kind, ext, name, dir = '', root = null, template = null }) {
  let bytes; let useExt = ext;
  if (template && template.own) {
    bytes = await office.templateBytes(template.id);
    useExt = template.ext || ext;
  } else {
    let chunk = null;
    try { chunk = await loadEditor(kind); } catch { chunk = null; }
    if (chunk && typeof chunk.create === 'function') bytes = await chunk.create({ ext, template: (template && template.id) || 'blank', settings: settingsNow() });
    else bytes = await blankBytes(kind, ext);
  }
  const entry = await office.create(root && root !== LIB_ROOT ? { root, dir } : dir, ensureExt(name, useExt), bytes);
  changed({ created: entry && entry.path });
  return entry;
}
export async function newFolder(dir) {
  const name = await promptDialog({ title: 'New folder', label: 'Name', value: t('New folder'), confirm: 'Create', validate: (v) => { const m = validateName(v); return m ? t(m) : null; } });
  if (!name) return null;
  const e = await office.mkdir(dir, name.trim());
  resetTree(); changed();
  toast.success(t('Folder “{name}” created', { name: name.trim() }));
  return e;
}
export async function uploadFiles(files, dir, onProgress) {
  const list = [...files].filter((f) => f && (f.size > 0 || f.type)); // folders dropped from Explorer come as empty entries
  let done = 0; const out = []; const failed = [];
  for (const f of list) {
    if (onProgress) onProgress({ done, total: list.length, name: f.name });
    try { out.push(await office.upload(dir, f)); } catch (e) { failed.push(`${f.name}: ${errMsg(e)}`); }
    done++;
  }
  if (onProgress) onProgress(null);
  changed();
  if (out.length) toast.success(pl(out.length, '{n} file uploaded', '{n} files uploaded'));
  if (failed.length) toast.error(t('Could not upload: {list}', { list: failed.join(' · ') }));
  if (!list.length && files.length) toast.warn(t('Folders cannot be uploaded this way – open the folder and drop its files.'));
  return out;
}

/* ---------------- Rename, duplicate, move, copy ---------------- */
export async function renameEntry(entry, siblings = []) {
  const { base, ext } = splitExt(entry.name);
  const name = await promptDialog({
    title: entry.dir ? 'Rename folder' : 'Rename file', label: 'New name', value: entry.name, confirm: 'Rename',
    validate: (v) => { const m = validateName(v, { existing: siblings, current: entry.name }); return m ? t(m) : null; },
  });
  if (!name || name.trim() === entry.name) return null;
  let next = name.trim();
  // the extension decides which editor opens the file – keep it unless the user typed another known one
  if (!entry.dir && ext && !FORMATS[extOf(next)] && base !== next) next = ensureExt(next, ext);
  const e = await office.rename(entry.root ? entry : entry.path, next);
  if (entry.dir) resetTree();
  changed({ renamed: entry.path });
  toast.success(t('Renamed to “{name}”', { name: next }));
  return e;
}
export async function duplicate(entries) {
  const r = await office.copy(entries.map((e) => e.path), dirname(entries[0].path));
  changed();
  toast.success(pl(entries.length, 'Duplicated', '{n} items duplicated'));
  return r;
}
export async function moveEntries(entries, to) {
  const from = entries.map((e) => dirname(e.path));
  if (entries.some((e) => e.dir && (to === e.path || to.startsWith(`${e.path}/`)))) throw new Error(t('A folder cannot be moved into itself.'));
  if (entries.every((e, i) => from[i] === to)) return null;
  const r = await office.move(entries.map((e) => e.path), to);
  resetTree(); changed();
  const moved = (r && (r.entries || (Array.isArray(r) ? r : null))) || null;
  toast(pl(entries.length, 'Moved to “{to}”', '{n} items moved to “{to}”', { to: to || t('Library') }), 'success', 8000, moved ? {
    action: { label: t('Undo'), onClick: async () => { try { for (let i = 0; i < moved.length; i++) await office.move([moved[i].path], from[i] || ''); resetTree(); changed(); toast('Move undone'); } catch (e) { toast.error(e); } } },
  } : null);
  return r;
}
export async function copyEntries(entries, to) {
  const r = await office.copy(entries.map((e) => e.path), to);
  resetTree(); changed();
  toast.success(pl(entries.length, 'Copied to “{to}”', '{n} items copied to “{to}”', { to: to || t('Library') }));
  return r;
}

/* ---------------- Trash ---------------- */
export async function trashEntries(entries) {
  const s = settingsNow();
  if (s.general.confirmDelete) {
    const ok = await confirmDialog({
      title: entries.length === 1 ? t('Move “{name}” to the trash?', { name: entries[0].name }) : t('Move {n} items to the trash?', { n: entries.length }),
      text: t('You can restore it from the trash for {days} days.', { days: s.general.trashDays }), confirm: 'Move to trash', danger: true,
    });
    if (!ok) return null;
  }
  const paths = entries.map((e) => e.path);
  const r = await office.remove(paths);
  resetTree(); changed();
  // the backend answers with the trash ids (ids or entries); if not, look them up in the trash list
  let ids = (r && (r.ids || (r.entries || []).map((e) => e.id))) || [];
  if (!ids.length) { try { const tr = await office.list({ view: 'trash' }); ids = (tr.entries || []).filter((e) => paths.includes(e.from)).map((e) => e.id); } catch { /* no undo then */ } }
  toast(pl(entries.length, '“{name}” is in the trash', '{n} items are in the trash', { name: entries[0].name }), 'success', 9000, ids.length ? {
    action: { label: t('Undo'), onClick: async () => { try { await office.restore(ids); resetTree(); changed(); toast('Restored'); } catch (e) { toast.error(e); } } },
  } : null);
  return r;
}
export async function restoreTrash(entries) {
  await office.restore(entries.map((e) => e.id));
  resetTree(); changed();
  toast.success(pl(entries.length, 'Restored', '{n} items restored'));
}
export async function purgeTrash(entries) {
  const ok = await confirmDialog({ title: entries.length === 1 ? t('Delete “{name}” for good?', { name: entries[0].name }) : t('Delete {n} items for good?', { n: entries.length }), text: 'This cannot be undone.', confirm: 'Delete forever', danger: true });
  if (!ok) return;
  await office.purge(entries.map((e) => e.id));
  changed();
  toast.success(pl(entries.length, 'Deleted', '{n} items deleted'));
}
export async function emptyTrash(count) {
  const ok = await confirmDialog({ title: 'Empty the trash?', text: count ? t('{n} items will be deleted for good. This cannot be undone.', { n: count }) : 'Everything in the trash will be deleted for good. This cannot be undone.', confirm: 'Empty trash', danger: true });
  if (!ok) return;
  await office.purge();
  changed();
  toast.success('Trash emptied');
}

/* ---------------- Download, convert, desktop ---------------- */
export function downloadEntries(entries) {
  if (entries.length === 1 && !entries[0].dir) { downloadFrom(office.downloadUrl(entries[0]), entries[0].name); return; }
  downloadFrom(office.zipUrl(entries.map((e) => e.path)), entries.length === 1 ? `${entries[0].name}.zip` : 'Office.zip');
}
export const downloadZip = (entries) => downloadFrom(office.zipUrl(entries.map((e) => e.path)), entries.length === 1 ? `${entries[0].name}.zip` : 'Office.zip');
/**
 * "Convert to …" of a library file. Server converter when installed (result next to the file), otherwise the
 * editor chunk does it: the editor opens and saves the copy (then=convert:<ext>).
 */
export async function convertEntry(entry, toExt, { via = 'server' } = {}) {
  const ext = extOf(entry.name);
  const st = await loadStatus().catch(() => null);
  if (via === 'client') { openEditor(entry, { then: `convert:${toExt}` }); return null; }
  if (!hasConverterIn(st)) {
    const ready = await ensureEngine(t('Converting to {ext} needs the Office engine.', { ext: toExt.toUpperCase() }));
    if (!ready) { toast(t('The conversion starts as soon as the Office engine is ready – try again then.')); return null; }
  }
  const id = toast.progress(t('Converting “{name}” to {ext} …', { name: entry.name, ext: toExt.toUpperCase() }));
  try {
    const { bytes } = await office.read(entry);
    const out = await office.convert(bytes, ext, toExt);
    const name = `${splitExt(entry.name).base}.${toExt}`;
    // next to the file – library or any folder of Hostwerk's Files; the backend never overwrites (first free name)
    const created = await office.create(entry.root ? { root: entry.root, dir: dirname(entry.path) } : dirname(entry.path), name, out);
    toast.dismiss(id);
    changed();
    window.dispatchEvent(new Event('hw:files-marks'));
    toast(t('Converted – saved as “{name}”', { name: created.name || name }), 'success', 8000, { action: { label: t('Open'), onClick: () => openEditor(created) } });
    return created;
  } catch (e) { toast.dismiss(id); throw e; }
}
export async function openInDesktop(f) {
  await office.openDesktop(f);
  toast(t('“{name}” opens in the desktop app', { name: f.name || basename(f.path) }));
}
export async function revealEntry(f) { await office.reveal(f); }
export async function toggleFavorite(entries, on) {
  for (const e of entries) await office.meta(e.root ? e : e.path, { favorite: on });
  changed();
  toast.success(on ? (entries.length === 1 ? 'Added to favourites' : t('{n} items added to favourites', { n: entries.length })) : 'Removed from favourites');
}
export async function saveAsTemplate(entry) {
  const r = await promptDialog({ title: 'Save as template', text: 'The template appears under “New” and in the template gallery.', label: 'Template name', value: splitExt(entry.name).base, confirm: 'Save template', validate: (v) => (v.trim() ? null : t('Please enter a name.')) });
  if (!r) return;
  await office.saveTemplate(entry, r.trim(), t('My templates'));
  changed({ templates: true });
  toast.success(t('Template “{name}” saved', { name: r.trim() }));
}
/** Copy a file from Hostwerk's file browser into the library root */
export async function copyToLibrary(f) {
  const { bytes } = await office.read(f);
  const e = await office.create('', basename(f.path), bytes);
  changed();
  toast(t('Copied to the Office library as “{name}”', { name: e.name || basename(f.path) }), 'success', 8000, { action: { label: t('Open'), onClick: () => openEditor(e) } });
  return e;
}
export { fileKey };
