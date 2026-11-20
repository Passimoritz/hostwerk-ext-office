/*
 * Components on demand (ARCHITECTURE.md → "Office 1.1 – expansion contract"): heavy parts that are downloaded once
 * (pinned SHA-256) into <data>/office-components/<id>/ and served by GET /office/components/<id>/<file>.
 *
 *   ensureComponent(host, id, reason)  → { base, files } when installed (asks once through the shell, size shown)
 *   componentState(id)                 → { installed, files?, size? } | null
 *   componentFile(base, names)         → first file of `names` that exists (URL), or null
 *
 * The shell offers host.components.ensure(id, { reason }) – until it exists we talk to the backend directly.
 */
import { t } from '@hostwerk/i18n';
import { confirmDialog } from '@hostwerk/ui';

const API = '/api/office/components';
let stateCache = null;

async function json(url, opts) {
  const r = await fetch(url, { credentials: 'same-origin', ...opts });
  if (!r.ok) { let msg = ''; try { msg = (await r.json()).error || ''; } catch { /* text */ } throw Object.assign(new Error(msg || `HTTP ${r.status}`), { status: r.status }); }
  return r.json();
}

export async function componentState(id, { fresh = false } = {}) {
  try {
    if (!stateCache || fresh) stateCache = await json(API);
    const list = Array.isArray(stateCache) ? stateCache : stateCache.components || Object.entries(stateCache).map(([k, v]) => ({ id: k, ...v }));
    return list.find((c) => c.id === id) || null;
  } catch { return null; }
}
export const isReady = (s) => !!(s && (s.installed === true || s.state === 'ready' || s.state === 'installed'));
export const componentBase = (id) => `${API}/${encodeURIComponent(id)}/`;

/** ask the shell (or the backend) to make the component available → { base, files } or null (declined / failed) */
export async function ensureComponent(host, id, reason) {
  let s = await componentState(id, { fresh: true });
  if (isReady(s)) return { base: componentBase(id), files: s.files || null };
  if (host && host.components && typeof host.components.ensure === 'function') {
    const r = await host.components.ensure(id, { reason });
    if (!r) return null;
    s = await componentState(id, { fresh: true });
    return { base: componentBase(id), files: (s && s.files) || null };
  }
  if (!s) return null; // backend without components
  const mb = s.download || s.size;
  const size = mb ? ` (≈ ${Math.round(mb / 1048576)} MB)` : '';
  // the shell's dialog is not there – ask with Hostwerk's own dialog
  const ok = await confirmDialog({ title: s.label || t('Add-on'), text: `${reason || ''} ${t('Download this add-on once{size}? It is checked and kept in Hostwerk’s data folder.', { size })}`, confirm: t('Download') });
  if (!ok) return null;
  await json(`${API}/install`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ id }) });
  for (let i = 0; i < 600; i++) {
    s = await componentState(id, { fresh: true });
    if (isReady(s)) return { base: componentBase(id), files: s.files || null };
    if (s && (s.error || s.state === 'error')) throw new Error(s.error || t('The add-on could not be installed.'));
    await new Promise((r) => setTimeout(r, 1000));
  }
  throw new Error(t('Installing the add-on takes too long – try again later.'));
}

/** first of the candidate file names the component really has (uses its file list, else asks the server) */
export async function componentFile(comp, names) {
  if (!comp) return null;
  const files = Array.isArray(comp.files) ? new Set(comp.files.map((f) => (typeof f === 'string' ? f : f.name || f.path))) : null;
  for (const n of names) {
    if (files) { if (files.has(n)) return comp.base + n; continue; }
    try { const r = await fetch(comp.base + n, { method: 'HEAD', credentials: 'same-origin' }); if (r.ok) return comp.base + n; } catch { /* next */ }
  }
  return null;
}

/** fetch a component file as bytes (gzip-compressed files are unpacked) */
export async function componentBytes(url) {
  const r = await fetch(url, { credentials: 'same-origin' });
  if (!r.ok) throw new Error(`HTTP ${r.status}`);
  let b = new Uint8Array(await r.arrayBuffer());
  if (b[0] === 0x1f && b[1] === 0x8b && typeof DecompressionStream !== 'undefined') {
    b = new Uint8Array(await new Response(new Blob([b]).stream().pipeThrough(new DecompressionStream('gzip'))).arrayBuffer());
  }
  return b;
}
