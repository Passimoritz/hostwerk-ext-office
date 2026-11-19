'use strict';
/*
 * Google Drive inside Hostwerk (no external browser):
 *   connect   OAuth 2.0 for installed apps with PKCE and a loopback redirect (http://127.0.0.1:<Hostwerk port>/api/office/google/callback).
 *             The Google sign-in page opens in an in-app Hostwerk window (host.ctx.app.openWebWindow({ url, waitFor })),
 *             which resolves with the redirect URL – the code is taken from there.
 *   tokens    sealed with Hostwerk's secret store (core/secrets) in <data>/google.json; access tokens refresh themselves
 *   files     list folders and files, export Google-native files (docx/xlsx/pptx/pdf/odt…), download others,
 *             upload back (Google-native files are updated with conversion), import into the library
 *   pointers  .gdoc / .gsheet / .gslides … files of Google Drive for desktop ({ doc_id } / { resource_id } / { url })
 */
const crypto = require('crypto');
const { readJson, writeJson } = require('./store');

const AUTH = 'https://accounts.google.com/o/oauth2/v2/auth';
const TOKEN = 'https://oauth2.googleapis.com/token';
const API = 'https://www.googleapis.com/drive/v3';
const UPLOAD = 'https://www.googleapis.com/upload/drive/v3';
const SCOPE = 'https://www.googleapis.com/auth/drive openid email profile';

const G = 'application/vnd.google-apps.';
/** Google-native types → default export + all export targets (extension → mime) */
const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', odt: 'application/vnd.oasis.opendocument.text', rtf: 'application/rtf', pdf: 'application/pdf', txt: 'text/plain', html: 'application/zip', epub: 'application/epub+zip', md: 'text/markdown',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', ods: 'application/vnd.oasis.opendocument.spreadsheet', csv: 'text/csv', tsv: 'text/tab-separated-values',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', odp: 'application/vnd.oasis.opendocument.presentation',
  png: 'image/png', jpg: 'image/jpeg', svg: 'image/svg+xml',
};
const NATIVE = {
  [`${G}document`]: { ext: 'docx', kind: 'docs', pointer: 'gdoc', exports: ['docx', 'odt', 'rtf', 'pdf', 'txt', 'epub', 'md'] },
  [`${G}spreadsheet`]: { ext: 'xlsx', kind: 'sheets', pointer: 'gsheet', exports: ['xlsx', 'ods', 'csv', 'tsv', 'pdf'] },
  [`${G}presentation`]: { ext: 'pptx', kind: 'slides', pointer: 'gslides', exports: ['pptx', 'odp', 'pdf', 'txt'] },
  [`${G}drawing`]: { ext: 'pdf', kind: 'viewer', pointer: 'gdraw', exports: ['pdf', 'png', 'jpg', 'svg'] },
};
const FOLDER = `${G}folder`;

const b64u = (b) => Buffer.from(b).toString('base64url');
/** PKCE pair */
function pkce() { const verifier = b64u(crypto.randomBytes(48)); return { verifier, challenge: b64u(crypto.createHash('sha256').update(verifier).digest()) }; }
function authUrl({ clientId, redirectUri, challenge, state, hint }) {
  const q = new URLSearchParams({ client_id: clientId, redirect_uri: redirectUri, response_type: 'code', scope: SCOPE, code_challenge: challenge, code_challenge_method: 'S256', state, access_type: 'offline', prompt: 'consent', ...(hint ? { login_hint: hint } : {}) });
  return `${AUTH}?${q}`;
}
/** .gdoc/.gsheet/… content → Drive file id or null */
function parsePointer(text) {
  let j; try { j = JSON.parse(String(text || '').replace(/^﻿/, '')); } catch { return null; }
  if (!j || typeof j !== 'object') return null;
  const ok = (id) => (/^[\w-]{10,200}$/.test(String(id || '')) ? String(id) : null);
  if (j.doc_id) return ok(j.doc_id);
  if (j.resource_id) return ok(String(j.resource_id).split(':').pop());
  if (j.url) { const m = String(j.url).match(/[?&]id=([\w-]+)|\/d\/([\w-]+)/); if (m) return ok(m[1] || m[2]); }
  return null;
}
/** Drive file → entry for the UI */
function entryOf(f) {
  const n = NATIVE[f.mimeType];
  return { id: f.id, name: f.name, dir: f.mimeType === FOLDER, mimeType: f.mimeType, google: !!n, kind: n ? n.kind : null, ext: n ? n.ext : (/\.([a-z0-9]+)$/i.exec(f.name || '') || [])[1]?.toLowerCase() || '', exports: n ? n.exports : [], size: f.size ? Number(f.size) : 0, mtime: f.modifiedTime ? Date.parse(f.modifiedTime) : 0, owner: f.owners && f.owners[0] ? f.owners[0].displayName : null, shared: !!f.shared, starred: !!f.starred, parents: f.parents || [], webViewLink: undefined };
}

const fail = (msg, status = 400, vars) => Object.assign(new Error(msg), { status, vars });

class Drive {
  /** deps: { file, seal, open, settings(), openWindow({ url, title, partition, waitFor }), redirectUri(), fetch } */
  constructor(deps) { this.d = deps; this.fetch = deps.fetch || globalThis.fetch; this.pending = null; }
  store() { return readJson(this.d.file, {}); }
  save(v) { writeJson(this.d.file, v); }
  client() {
    const g = (this.d.settings().integrations || {}).google || {};
    return { clientId: g.clientId || '', clientSecret: g.clientSecret || '' };
  }
  tokens() { const s = this.store(); if (!s.tokens) return null; try { return JSON.parse(this.d.open(s.tokens) || 'null'); } catch { return null; } }
  setTokens(t) { const s = this.store(); s.tokens = t ? this.d.seal(JSON.stringify(t)) : undefined; if (!t) delete s.user; this.save(s); }
  state() {
    const c = this.client(); const s = this.store();
    return { configured: !!c.clientId, connected: !!this.tokens(), account: s.user || null, user: s.user || null, connecting: !!(this.pending && this.pending.until > Date.now()), inApp: this.d.canWindow(), redirectUri: this.d.redirectUri(),
      guide: ['Open console.cloud.google.com → APIs & Services → enable the "Google Drive API".', 'Credentials → Create credentials → OAuth client ID → Application type "Desktop app".', 'Copy the client ID and the client secret into Office → Settings → Integrations → Google Drive.', 'OAuth consent screen: add your Google account as a test user (or publish the app).', 'Click "Connect" – the Google sign-in opens in a Hostwerk window.'] };
  }
  /** Start a sign-in → { url } (pending for 10 minutes) */
  begin() {
    const c = this.client();
    if (!c.clientId) throw fail('Enter the client ID of your Google OAuth client (Desktop app) in the Office settings first.', 409);
    const { verifier, challenge } = pkce();
    const state = b64u(crypto.randomBytes(18));
    const redirectUri = this.d.redirectUri();
    this.pending = { state, verifier, redirectUri, until: Date.now() + 10 * 60000 };
    return { url: authUrl({ clientId: c.clientId, redirectUri, challenge, state }), redirectUri };
  }
  /** Redirect URL (with ?code=&state=) → tokens stored → state() */
  async finish(back) {
    const p = this.pending;
    if (!p || p.until < Date.now()) throw fail('The Google sign-in could not be verified – please try again.', 400);
    const u = new URL(back);
    if (u.searchParams.get('state') !== p.state) throw fail('The Google sign-in could not be verified – please try again.', 400);
    this.pending = null;
    if (u.searchParams.get('error')) throw fail('Google refused the sign-in: {error}', 400, { error: u.searchParams.get('error') });
    const code = u.searchParams.get('code');
    if (!code) throw fail('The Google sign-in was closed.', 409);
    const t = await this.tokenCall({ grant_type: 'authorization_code', code, code_verifier: p.verifier, redirect_uri: p.redirectUri });
    this.setTokens({ access: t.access_token, refresh: t.refresh_token, exp: Date.now() + (t.expires_in || 3600) * 1000 - 60000 });
    try { const me = await this.json(`${API}/about?fields=user(displayName,emailAddress,photoLink)`); const s = this.store(); s.user = { name: me.user.displayName, email: me.user.emailAddress }; this.save(s); } catch { /* optional */ }
    return this.state();
  }
  /**
   * Connect: with the desktop app the sign-in opens in a Hostwerk window and this resolves when it is done; without it
   * → { url } for the shell (openInApp) and the loopback callback route finishes the sign-in.
   */
  async connect() {
    if (this.pending && this.pending.until > Date.now() && this.pending.window) throw fail('A Google sign-in is already open.', 409);
    const { url, redirectUri } = this.begin();
    if (!this.d.canWindow()) return { ...this.state(), url };
    this.pending.window = true;
    let back;
    try { back = await this.d.openWindow({ url, title: 'Google', partition: 'persist:hw-auth-google', waitFor: redirectUri, width: 520, height: 720 }); } catch (e) { this.pending = null; throw e; }
    if (!back) { this.pending = null; throw fail('The Google sign-in was closed.', 409); }
    return this.finish(back);
  }
  async disconnect() {
    const t = this.tokens();
    if (t && (t.refresh || t.access)) { try { await this.fetch(`https://oauth2.googleapis.com/revoke?token=${encodeURIComponent(t.refresh || t.access)}`, { method: 'POST' }); } catch { /* offline – forget anyway */ } }
    this.setTokens(null);
    return this.state();
  }
  async tokenCall(params) {
    const c = this.client();
    const body = new URLSearchParams({ client_id: c.clientId, ...(c.clientSecret ? { client_secret: c.clientSecret } : {}), ...params });
    const r = await this.fetch(TOKEN, { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw fail('Google refused the sign-in: {error}', r.status === 400 ? 401 : 502, { error: j.error_description || j.error || r.status });
    return j;
  }
  async access() {
    const t = this.tokens();
    if (!t) throw fail('Connect your Google account first (Office → Settings → Integrations).', 401);
    if (t.access && t.exp > Date.now()) return t.access;
    if (!t.refresh) throw fail('The Google connection has expired – please connect again.', 401);
    const n = await this.tokenCall({ grant_type: 'refresh_token', refresh_token: t.refresh });
    const next = { ...t, access: n.access_token, exp: Date.now() + (n.expires_in || 3600) * 1000 - 60000, refresh: n.refresh_token || t.refresh };
    this.setTokens(next);
    return next.access;
  }
  async call(url, opts = {}) {
    const r = await this.fetch(url, { ...opts, headers: { ...(opts.headers || {}), authorization: `Bearer ${await this.access()}` } });
    if (r.status === 401) { const t = this.tokens(); if (t) { t.exp = 0; this.setTokens(t); } }
    if (!r.ok) { let msg = `${r.status}`; try { const j = await r.json(); msg = (j.error && j.error.message) || msg; } catch { /* not json */ } throw fail('Google Drive: {error}', r.status === 404 ? 404 : r.status === 403 ? 403 : 502, { error: msg }); }
    return r;
  }
  async json(url, opts) { return (await this.call(url, opts)).json(); }
  /** Folder listing / search → { folder, entries, next } */
  async list({ folder = 'root', q = '', view = '', pageToken = '' } = {}) {
    if (!/^[\w-]{1,200}$/.test(folder)) throw fail('Invalid folder', 400);
    const esc = (s) => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    const parts = ['trashed = false'];
    if (q) parts.push(`(name contains '${esc(q)}' or fullText contains '${esc(q)}')`);
    else if (view === 'shared') parts.push('sharedWithMe = true');
    else if (view === 'starred') parts.push('starred = true');
    else if (view === 'recent') { /* ordered below */ } else parts.push(`'${folder}' in parents`);
    const params = new URLSearchParams({ q: parts.join(' and '), pageSize: '200', fields: 'nextPageToken,files(id,name,mimeType,modifiedTime,size,parents,owners(displayName),shared,starred)', orderBy: view === 'recent' ? 'viewedByMeTime desc' : 'folder,name', supportsAllDrives: 'true', includeItemsFromAllDrives: 'true', ...(pageToken ? { pageToken } : {}) });
    const j = await this.json(`${API}/files?${params}`);
    return { folder, entries: (j.files || []).filter((f) => !/shortcut$/.test(f.mimeType) || true).map(entryOf), next: j.nextPageToken || null };
  }
  async meta(id) {
    if (!/^[\w-]{10,200}$/.test(String(id || ''))) throw fail('Invalid Google file id', 400);
    const f = await this.json(`${API}/files/${id}?fields=id,name,mimeType,modifiedTime,size,parents,owners(displayName),shared,starred,capabilities(canEdit)&supportsAllDrives=true`);
    return { ...entryOf(f), canEdit: !!(f.capabilities && f.capabilities.canEdit) };
  }
  /** Bytes of a file: Google-native → export as `as` (default per type), others → download → { data, name, ext, mime } */
  async read(id, as) {
    const m = await this.meta(id);
    const n = NATIVE[m.mimeType];
    if (m.dir) throw fail('This is a folder.', 400);
    if (n) {
      const ext = String(as || n.ext).toLowerCase();
      if (!n.exports.includes(ext) || !MIME[ext]) throw fail('Google cannot export this file as {ext}.', 400, { ext });
      const r = await this.call(`${API}/files/${id}/export?mimeType=${encodeURIComponent(MIME[ext])}`);
      return { data: Buffer.from(await r.arrayBuffer()), name: `${m.name}.${ext}`, ext, mime: MIME[ext], meta: m };
    }
    const r = await this.call(`${API}/files/${id}?alt=media&supportsAllDrives=true`);
    return { data: Buffer.from(await r.arrayBuffer()), name: m.name, ext: m.ext, mime: r.headers.get('content-type') || 'application/octet-stream', meta: m };
  }
  /** Upload new content (ext = format of the bytes). Google-native files are converted by Drive. */
  async write(id, data, ext) {
    const m = await this.meta(id);
    if (m.dir) throw fail('This is a folder.', 400);
    const mime = MIME[String(ext || m.ext).toLowerCase()] || 'application/octet-stream';
    const r = await this.json(`${UPLOAD}/files/${id}?uploadType=media&supportsAllDrives=true&fields=id,name,mimeType,modifiedTime,size`, { method: 'PATCH', headers: { 'content-type': mime }, body: data });
    return entryOf(r);
  }
}

module.exports = { Drive, pkce, authUrl, parsePointer, entryOf, NATIVE, MIME, SCOPE };
