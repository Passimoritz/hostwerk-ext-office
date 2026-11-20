/*
 * Loads an editor chunk (ui/assets/<kind>.js, served by GET /api/office/assets/<kind>.js) on first use and keeps it
 * (the module stays warm for the whole session – opening the next file of the same kind needs no network and no parse).
 * A chunk's default export follows the editor contract in ARCHITECTURE.md ({ kind, Editor, create, toText, … }).
 *
 *   prefetchEditor(kind)          fetch the chunk (+ css, + shared libs) without running it – <link rel=modulepreload>,
 *                                 on hover/focus of an office file
 *   warmEditor(kind)              load AND run it in idle time (Dateien opened the Office root, a file is selected)
 *   prefetchFor(name)             the same by file name
 */
import { office } from './api';
import { CHUNK_VERSION as VERSION } from './version';
import { FORMATS, extOf } from './formats';

const cache = new Map();
const hinted = new Set();
// shared libraries a chunk loads first (tools/build.js) – preloading them too saves one round trip
const LIBS = { docs: ['lib-jszip'], sheets: ['lib-jszip'], slides: ['lib-jszip'], viewer: ['lib-jszip', 'lib-pdfjs'], pdf: ['lib-pdfjs'] };
const KNOWN = new Set(['docs', 'sheets', 'slides', 'pdf', 'viewer']);

const assetHref = (file) => `${office.assetUrl(file)}?v=${VERSION}`;

function css(kind) {
  const id = `office-css-${kind}`;
  if (document.getElementById(id)) return;
  const l = document.createElement('link');
  l.id = id; l.rel = 'stylesheet'; l.href = assetHref(`${kind}.css`);
  l.onerror = () => l.remove(); // chunks without own styles have no .css – that is fine
  document.head.appendChild(l);
}

export function loadEditor(kind) {
  if (!cache.has(kind)) {
    css(kind);
    const t0 = performance.now();
    const p = import(/* @vite-ignore */ assetHref(`${kind}.js`)).then((m) => {
      try { performance.measure(`office:chunk:${kind}`, { start: t0, end: performance.now() }); } catch { /* old browser */ }
      return m.default || m;
    });
    p.catch(() => cache.delete(kind)); // try again next time (e.g. after an update)
    cache.set(kind, p);
  }
  return cache.get(kind);
}
export const isLoaded = (kind) => cache.has(kind);

function hint(file, rel = 'modulepreload') {
  const href = assetHref(file);
  if (hinted.has(href)) return;
  hinted.add(href);
  const l = document.createElement('link');
  l.rel = rel; l.href = href;
  if (rel === 'preload') l.as = 'style';
  l.onerror = () => l.remove();
  document.head.appendChild(l);
}
/** fetch a chunk without running it (cheap – hover/focus) */
export function prefetchEditor(kind) {
  if (!KNOWN.has(kind) || cache.has(kind) || typeof document === 'undefined') return;
  hint(`${kind}.js`);
  for (const l of LIBS[kind] || []) hint(`${l}.js`);
}
/** load and run a chunk when the browser is idle (the next open is instant) */
export function warmEditor(kind) {
  if (!KNOWN.has(kind) || cache.has(kind)) return;
  const go = () => { loadEditor(kind).catch(() => {}); };
  if (typeof requestIdleCallback === 'function') requestIdleCallback(go, { timeout: 4000 }); else setTimeout(go, 400);
}
export function kindFor(name) { const f = FORMATS[extOf(name || '')]; return f ? f.kind : null; }
export function prefetchFor(name, { warm = false } = {}) { const k = kindFor(name); if (!k) return; if (warm) warmEditor(k); else prefetchEditor(k); }
