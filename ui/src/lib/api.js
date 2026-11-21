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
  let j = null;
  try { j = await res.json(); msg = j.error || j.message || msg; } catch { /* not json */ }
  const e = new Error(msg); e.status = res.status; if (j && j.mtime != null) e.mtime = j.mtime; throw e;
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
  marks: ({ root, path, types, client }) => get(`/office/marks?${q({ root, path, types: types ? 1 : undefined, client })}`),
  /** "Open with" / default app (SDK 3.15) */
  os: () => get('/office/os'),
  versions: (f) => get(`/office/versions?${q(loc(f))}`),
  restoreVersion: (f, id) => post('/office/versions/restore', { ...loc(f), id }),
  labelVersion: (f, id, label) => put('/office/versions/label', { ...loc(f), id, label }),
  templates: () => get('/office/templates'),
  /** "Save as template" (real dotx/xltx/potx/ott … – the server switches the package type); to: 'folder' | 'next' | 'both' → { template, entries } */
  saveTemplate: (f, name, category, to) => post('/office/templates', { ...loc(f), name, category, ...(to ? { to } : {}) }),
  renameTemplate: (id, name) => post(`/office/templates/${encodeURIComponent(id)}/rename`, { name }),
  /** default template of its app (new blank documents start from it) – on: false removes it */
  defaultTemplate: (id, on = true) => post(`/office/templates/${encodeURIComponent(id)}/default`, { on }),
  /** usual template folders on this computer → { folders: [{ path, label, count, added }] } */
  templateFolders: () => get('/office/templates/suggest'),
  /** a template FILE (double-click in Files) → a new document next to it (template folders: the library root) → Entry */
  fromTemplate: (f, name) => post('/office/fromtemplate', { ...loc(f), ...(name ? { name } : {}) }),
  /** new document from a listed template → Entry (dir/root: where; default the library root) */
  newFromTemplate: (id, { root, dir, name } = {}) => post(`/office/templates/${encodeURIComponent(id)}/new`, { ...(root ? { root } : {}), ...(dir != null ? { dir } : {}), ...(name ? { name } : {}) }),
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
  /** password-protected Office files (server/encrypt.js): state → { encrypted: 'ooxml'|'odf'|null, supported, vault: { available, locked, hasEntry } } */
  encryptState: (f) => get(`/office/encrypt/state?${q(loc(f))}`),
  /** the decrypted bytes for the editor (the file itself stays encrypted); throws status 403 = wrong password, 404 = not in the vault */
  async decryptBytes(f, { password = '', fromVault = false } = {}) {
    // soft: the server answers a wrong password with 200 { ok: false } (no red 403 in the console) – older servers still send 403
    const r = await check(await fetch(url('/decrypt'), { method: 'POST', credentials: 'same-origin', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ ...loc(f), ...(password ? { password } : { fromVault: !!fromVault }), mode: 'bytes', soft: true }) }));
    if (/application\/json/.test(r.headers.get('content-type') || '')) {
      const j = await r.json().catch(() => ({}));
      const e = new Error(j.error || 'The password is not correct.'); e.status = j.reason === 'vault' ? 404 : 403; throw e;
    }
    return { bytes: new Uint8Array(await r.arrayBuffer()), mtime: Number(r.headers.get('x-mtime')) || null };
  },
  /** save an opened encrypted file: plain bytes in, the server encrypts again with the same password (or the vault's) */
  async writeEncrypted(f, bytes, { mtime, force = false, password = '' } = {}) {
    const r = await check(await fetch(url('/encrypt/raw', { ...loc(f), mtime, force: force ? 1 : undefined }), { method: 'PUT', credentials: 'same-origin', headers: { 'content-type': 'application/octet-stream', ...(password ? { 'x-office-password': encodeURIComponent(password) } : {}) }, body: bytes }));
    return r.json(); // { ok, mtime, size, version }
  },
  /** Windows' "How do you want to open this file?" for one kind; bytes = an empty file of its type (sample) → { opened, ext, state } */
  async osOpenWith(kind, bytes) {
    const r = await check(await fetch(url('/os/openwith', { kind }), { method: 'POST', credentials: 'same-origin', headers: { 'content-type': 'application/octet-stream' }, body: bytes || new Uint8Array() }));
    return r.json();
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
  /** bytes of a NEW document made from a listed template → { bytes, ext } (dotx → docx …) */
  async templateDoc(id) { const r = await check(await fetch(url(`/templates/${encodeURIComponent(id)}/doc`), { credentials: 'same-origin' })); return { bytes: new Uint8Array(await r.arrayBuffer()), ext: r.headers.get('x-ext') || '' }; },
  templateThumbUrl: (id, v) => url(`/templates/${encodeURIComponent(id)}/thumb`, { v }),
  async putTemplateThumb(id, blob) { const r = await check(await fetch(url(`/templates/${encodeURIComponent(id)}/thumb`), { method: 'PUT', credentials: 'same-origin', headers: { 'content-type': blob.type || 'image/webp' }, body: blob })); return r.json().catch(() => ({ ok: true })); },
  async templateBytes(id) { const r = await check(await fetch(url(`/templates/${encodeURIComponent(id)}/raw`), { credentials: 'same-origin' })); return new Uint8Array(await r.arrayBuffer()); },
  downloadUrl: (f) => url('/raw', { ...loc(f), download: 1 }),
  zipUrl: (paths) => url('/zip', { paths: JSON.stringify(paths) }),
  assetUrl: (name) => url(`/assets/${name}`),

  /* ---------------- 1.1 (ARCHITECTURE.md "Office 1.1 – expansion contract" / "Host object 1.1") ---------------- */
  /** generated first-page thumbnail (image/webp or png) – stored next to the metadata, dropped when the file changes */
  async putThumb(f, blob, mtime) {
    const r = await check(await fetch(url('/thumb', { ...loc(f), mtime }), { method: 'PUT', credentials: 'same-origin', headers: { 'content-type': blob.type || 'image/webp' }, body: blob }));
    return r.json().catch(() => ({ ok: true }));
  },
  thumbUrl: (f, v) => url('/thumb', { ...loc(f), v }),
  /** search by meaning (AI embeddings on the server) – same hits as search() plus score */
  semanticSearch: (text, params) => get(`/office/search?${q({ q: text, semantic: 1, ...params })}`),
  data: {
    sources: () => get('/office/data/sources').then(list('sources')),
    query: (source, params) => get(`/office/data/query?${q({ source, ...(params || {}) })}`),
  },
  db: {
    servers: () => get('/office/db/servers').then(list('servers')),
    tables: (server, db) => get(`/office/db/tables?${q({ server, db })}`).then(list('tables')),
    rows: ({ server, db, table, offset, limit, sort, filter }) => get(`/office/db/rows?${q({ server, db, table, offset, limit, sort, filter: filter && typeof filter === 'object' ? JSON.stringify(filter) : filter })}`),
    write: ({ server, db, table, changes }) => post('/office/db/rows', { server, db, table, changes }),
  },
  fields: {
    list: () => get('/office/fields').then(list('fields')),
    value: (id, params) => get(`/office/fields/value?${q({ id, params: params ? JSON.stringify(params) : undefined })}`),
  },
  forms: {
    list: (f) => get(`/office/forms?${q(loc(f))}`).then(list('forms')),
    create: (f, form) => post('/office/forms', { ...loc(f), ...form }),
    remove: (id) => del(`/office/forms/${encodeURIComponent(id)}`),
  },
  rules: {
    get: (f) => get(`/office/rules?${q(loc(f))}`).then(list('rules')),
    set: (f, rules) => put(`/office/rules?${q(loc(f))}`, { rules }),
  },
  components: {
    list: () => get('/office/components').then(list('components')),
    install: (id) => post('/office/components/install', { id }),
    base: (id) => `${BASE}/components/${encodeURIComponent(id)}/`,
  },
  ai: (body) => post('/office/ai', body),
  vault: {
    state: () => get('/office/vault'),
    encrypt: (f) => post('/office/vault/encrypt', loc(f)),
    decrypt: (f) => post('/office/vault/decrypt', loc(f)),
  },
  clean: {
    preview: (f) => post('/office/clean/preview', loc(f)),
    run: (f, remove, { copy = true } = {}) => post('/office/clean', { ...loc(f), remove, copy }),
  },
  publish: {
    targets: () => get('/office/publish/targets').then(list('targets')),
    list: (f) => get(`/office/publish?${q(loc(f))}`).then(list('publications')),
    run: (f, { target, title, path }) => post('/office/publish', { ...loc(f), target, title, path }),
    remove: (id) => del(`/office/publish/${encodeURIComponent(id)}`),
  },
  reports: {
    list: (f) => get(`/office/reports?${q(f ? loc(f) : {})}`).then(list('reports')),
    save: (f, report) => post('/office/reports', { ...loc(f), ...report }),
    remove: (id) => del(`/office/reports/${encodeURIComponent(id)}`),
    run: (id) => post(`/office/reports/${encodeURIComponent(id)}/run`, {}),
  },
  backups: {
    list: (f) => get(`/office/backups?${q(loc(f))}`).then(list('backups')),
    async bytes(f, id) { const r = await check(await fetch(url('/backups/raw', { ...loc(f), id }), { credentials: 'same-origin' })); return new Uint8Array(await r.arrayBuffer()); },
    restore: (f, id) => post('/office/backups/restore', { ...loc(f), id }),
  },
  /** websocket address of a server route (co-editing, remote control) – same host, ws/wss */
  // Hostwerk hands logged-in websocket upgrades to extensions only under /ws/* (never /api/*) – server/collab.js, remote.js
  wsUrl: (p, params) => { const u = new URL(url(p, params).replace(/^\/api\/office\//, '/ws/office/'), location.href); u.protocol = u.protocol === 'https:' ? 'wss:' : 'ws:'; return u.href; },
};
/** list responses: the array itself or { <key>: [...] } (the server may wrap it) */
function list(key) { return (r) => (Array.isArray(r) ? r : (r && Array.isArray(r[key]) ? r[key] : (r && Array.isArray(r.items) ? r.items : []))); }
/** true when an error means "this server has no such route yet" (older Office server / feature not built) */
export const isMissing = (e) => !!e && (e.status === 404 || e.status === 501 || /\b(404|501)\b/.test(String(e.message || '')));

/** Save bytes as a download in the browser / Hostwerk window. */
export function download(bytes, name, mime = 'application/octet-stream') {
  const blob = bytes instanceof Blob ? bytes : new Blob([bytes], { type: mime });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}
