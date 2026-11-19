/*
 * Loads an editor chunk (ui/assets/<kind>.js, served by GET /api/office/assets/<kind>.js) on first use and keeps it.
 * A chunk's default export follows the editor contract in ARCHITECTURE.md ({ kind, Editor, create, toText, … }).
 */
import { office } from './api';
import { CHUNK_VERSION as VERSION } from './version';

const cache = new Map();

function css(kind) {
  const id = `office-css-${kind}`;
  if (document.getElementById(id)) return;
  const l = document.createElement('link');
  l.id = id; l.rel = 'stylesheet'; l.href = `${office.assetUrl(`${kind}.css`)}?v=${VERSION}`;
  l.onerror = () => l.remove(); // chunks without own styles have no .css – that is fine
  document.head.appendChild(l);
}

export function loadEditor(kind) {
  if (!cache.has(kind)) {
    css(kind);
    const p = import(/* @vite-ignore */ `${office.assetUrl(`${kind}.js`)}?v=${VERSION}`).then((m) => m.default || m);
    p.catch(() => cache.delete(kind)); // try again next time (e.g. after an update)
    cache.set(kind, p);
  }
  return cache.get(kind);
}
