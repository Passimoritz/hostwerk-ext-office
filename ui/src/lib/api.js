/*
 * Shared client for the Office backend (all routes under /api/office/…, see ARCHITECTURE.md).
 * JSON calls go through Hostwerk's helpers (same login, error toasts, other machines); binary calls use fetch
 * (Hostwerk patches window.fetch so /api/… also works when another machine is selected).
 */
import { get, post, put, del } from '@hostwerk/api';

export const BASE = '/api/office';
const q = (o) => Object.entries(o || {}).filter(([, v]) => v !== undefined && v !== null && v !== '').map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`).join('&');
export const url = (p, params) => `${BASE}${p}${params && q(params) ? `?${q(params)}` : ''}`;

async function check(res) {
  if (res.ok) return res;
  let msg = `${res.status} ${res.statusText}`;
  try { const j = await res.json(); msg = j.error || j.message || msg; } catch { /* not json */ }
  const e = new Error(msg); e.status = res.status; throw e;
}

/** File location: { path } inside the library, or { root, path } for a file from Hostwerk's file browser. */
const loc = (f) => (f.root ? { root: f.root, path: f.path } : { path: f.path });

export const office = {
  // JSON (paths without /api – Hostwerk's helpers add it)
  status: () => get('/office/status'),
  settings: () => get('/office/settings'),
  saveSettings: (s) => put('/office/settings', s),
  list: (params) => get(`/office/files?${q(params)}`),
  search: (text, params) => get(`/office/search?${q({ q: text, ...params })}`),
  mkdir: (dir, name) => post('/office/folder', { dir, name }),
  /** path: library path, or a file { root, path } of Hostwerk's Files (metadata and versions follow) */
  rename: (path, name) => post('/office/rename', path && typeof path === 'object' ? { ...loc(path), name } : { path, name }),
  move: (paths, to) => post('/office/move', { paths, to }),
  copy: (paths, to) => post('/office/copy', { paths, to }),
  remove: (paths) => post('/office/trash', { paths }),
  restore: (ids) => post('/office/trash/restore', { ids }),
  purge: (ids) => post('/office/trash/purge', { ids }), // ids omitted = empty the trash
  meta: (path, patch) => put('/office/meta', { ...(path && typeof path === 'object' ? loc(path) : { path }), ...patch }), // { favorite, tags, color, note }
  tags: () => get('/office/tags'),
  /** favourite / shared / lock marks of the office files in a folder of Hostwerk's Files → { names: { name: {…} } } */
  marks: ({ root, path }) => get(`/office/marks?${q({ root, path })}`),
  /** "Open with" / default app (SDK 3.15) */
  os: () => get('/office/os'),
  versions: (f) => get(`/office/versions?${q(loc(f))}`),
  restoreVersion: (f, id) => post('/office/versions/restore', { ...loc(f), id }),
  labelVersion: (f, id, label) => put('/office/versions/label', { ...loc(f), id, label }),
  templates: () => get('/office/templates'),
  saveTemplate: (f, name, category) => post('/office/templates', { ...loc(f), name, category }),
  removeTemplate: (id) => del(`/office/templates/${encodeURIComponent(id)}`),
  shares: () => get('/office/shares'),
  share: (path, opts) => post('/office/shares', { ...(path && typeof path === 'object' ? loc(path) : { path }), ...opts }), // { expires, password, allowDownload }
  unshare: (id) => del(`/office/shares/${encodeURIComponent(id)}`),
  openDesktop: (f) => post('/office/open-desktop', loc(f)),
  reveal: (f) => post('/office/reveal', loc(f)),
  info: (f) => get(`/office/info?${q(loc(f))}`), // { name, size, mtime, ext, kind, pages?, words?, author?, … }
  lock: (f) => post('/office/lock', loc(f)),
  unlock: (f) => post('/office/unlock', loc(f)),

  // Binary
  /** pointer: for .gdoc/.gsheet/.gslides/.gdraw the JSON pointer file itself (else the backend exports the Google file) */
  async read(f, { pointer = false } = {}) { const r = await check(await fetch(url('/raw', { ...loc(f), pointer: pointer ? 1 : undefined }), { credentials: 'same-origin' })); return { bytes: new Uint8Array(await r.arrayBuffer()), mtime: Number(r.headers.get('x-mtime')) || null }; },
  /** Upload bytes back to a Google Drive file (by its Drive id) – ext = format of the bytes (docx, xlsx, pptx, pdf …) */
  async writeGoogle(id, bytes, ext) {
    const r = await check(await fetch(url('/google/raw', { id, ext }), { method: 'PUT', credentials: 'same-origin', headers: { 'content-type': 'application/octet-stream' }, body: bytes }));
    return r.json();
  },
  async write(f, bytes, { mtime, force = false, label } = {}) {
    const r = await check(await fetch(url('/raw', { ...loc(f), mtime, force: force ? 1 : undefined, label }), { method: 'PUT', credentials: 'same-origin', headers: { 'content-type': 'application/octet-stream' }, body: bytes }));
    return r.json(); // { ok, mtime, size, version }
  },
  /** dir: library folder, or { root, dir } = a folder of Hostwerk's Files */
  async create(dir, name, bytes) {
    const where = dir && typeof dir === 'object' ? { root: dir.root, dir: dir.dir || '' } : { dir };
    const r = await check(await fetch(url('/create', { ...where, name }), { method: 'POST', credentials: 'same-origin', headers: { 'content-type': 'application/octet-stream' }, body: bytes || new Uint8Array() }));
    return r.json(); // entry of the new file (name may get " (2)")
  },
  async upload(dir, file) {
    const r = await check(await fetch(url('/create', { dir, name: file.name, keep: 1 }), { method: 'POST', credentials: 'same-origin', headers: { 'content-type': 'application/octet-stream' }, body: file }));
    return r.json();
  },
  /** Convert with LibreOffice / Microsoft Office on this PC (backend). Returns bytes of the target format. */
  async convert(bytes, fromExt, toExt) {
    const r = await check(await fetch(url('/convert', { from: fromExt, to: toExt }), { method: 'POST', credentials: 'same-origin', headers: { 'content-type': 'application/octet-stream' }, body: bytes }));
    return new Uint8Array(await r.arrayBuffer());
  },
  async versionBytes(f, id) { const r = await check(await fetch(url('/versions/raw', { ...loc(f), id }), { credentials: 'same-origin' })); return new Uint8Array(await r.arrayBuffer()); },
  async templateBytes(id) { const r = await check(await fetch(url(`/templates/${encodeURIComponent(id)}/raw`), { credentials: 'same-origin' })); return new Uint8Array(await r.arrayBuffer()); },
  downloadUrl: (f) => url('/raw', { ...loc(f), download: 1 }),
  zipUrl: (paths) => url('/zip', { paths: JSON.stringify(paths) }),
  assetUrl: (name) => url(`/assets/${name}`),
};

/** Save bytes as a download in the browser / Hostwerk window. */
export function download(bytes, name, mime = 'application/octet-stream') {
  const blob = bytes instanceof Blob ? bytes : new Blob([bytes], { type: mime });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}
