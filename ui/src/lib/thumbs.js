/*
 * First-page thumbnails of office files.
 *   thumbSrc(entry)            → address of a known thumbnail (made here, or Entry.thumb from the server) or null
 *   requestThumb(entry, cb)    → queue it: the server's (embedded / stored) thumbnail first, else render it here in idle
 *                                time – ONE file at a time, offscreen – and upload it (PUT /office/thumb); cb(url) → cancel()
 *   ensureThumb(entry)         → Promise<url|null> (files.thumb slot of Hostwerk's Files)
 * Rendering: the chunk's own thumbnail({ bytes, ext }) when it has one, else its Preview component mounted by the
 * renderer (shell/thumbHost.jsx – a React portal into a hidden container) and rasterised (canvas of pdf.js, or the DOM
 * with inlined styles through an SVG foreignObject).
 */
import { apiUrl } from '@hostwerk/api';
import { office, isMissing, url } from './api';
import { loadEditor } from './loader';
import { FORMATS, extOf } from './formats';

export const THUMB_W = 320;
const MAX_BYTES = 20 * 1048576;
const KINDS = new Set(['docs', 'sheets', 'slides', 'pdf']);
const keyOf = (e) => `${e.root || ''}|${e.path}`;
const loc = (e) => (e.root ? { root: e.root, path: e.path } : { path: e.path });
const made = new Map(); // key → { mtime, url }
const tried = new Map(); // key → mtime that failed (no retry for the same content)
const serverHas = new Map(); // key|mtime → Promise<url|null>
let serverStores = true; // PUT /office/thumb exists (false after a 404 – keep thumbnails in memory only)

export const canThumb = (e) => {
  if (!e || e.dir) return false;
  const f = FORMATS[extOf(e.name || e.path || '')];
  return !!f && KINDS.has(f.kind) && !(e.size > MAX_BYTES) && !/^(txt|md|markdown|csv|tsv|html?|xhtml|json|xml)$/.test(extOf(e.name || e.path || ''));
};
export function thumbSrc(e) {
  if (!e || e.dir) return null;
  const m = made.get(keyOf(e));
  if (m && (!e.mtime || !m.mtime || m.mtime === e.mtime)) return m.url;
  if (e.thumb) return apiUrl(office.thumbUrl(loc(e), e.mtime || ''));
  return null;
}
export function forgetThumb(e) { made.delete(keyOf(e)); }

/* ---------------- queue ---------------- */
const queue = []; // { e, cbs: Set }
let running = false;
const renderers = new Set(); // mounted thumbHost renderers: (job) => Promise<Blob|null>
const listeners = new Set();
export function onThumb(fn) { listeners.add(fn); return () => listeners.delete(fn); }
const idle = (ms = 1200) => new Promise((res) => (typeof requestIdleCallback === 'function' ? requestIdleCallback(() => res(), { timeout: ms }) : setTimeout(res, 60)));

/** the shell's offscreen renderer registers here (only one renders at a time – the last mounted one) */
export function registerRenderer(fn) { renderers.add(fn); pump(); return () => renderers.delete(fn); }

export function requestThumb(e, cb) {
  if (!canThumb(e)) return () => {};
  const k = keyOf(e);
  const have = thumbSrc(e);
  if (have) { cb(have); return () => {}; }
  if (tried.get(k) === (e.mtime || 0)) return () => {};
  let job = queue.find((j) => j.k === k);
  if (!job) { job = { k, e, cbs: new Set() }; queue.push(job); }
  job.cbs.add(cb);
  pump();
  return () => { job.cbs.delete(cb); if (!job.cbs.size && !job.started) { const i = queue.indexOf(job); if (i >= 0) queue.splice(i, 1); } };
}
export function ensureThumb(e) {
  return new Promise((resolve) => {
    const have = thumbSrc(e);
    if (have) { resolve(have); return; }
    const off = requestThumb(e, (u) => { off(); resolve(u); });
    setTimeout(() => { off(); resolve(thumbSrc(e)); }, 30000);
  });
}

async function serverThumb(e) {
  const k = `${keyOf(e)}|${e.mtime || ''}`;
  if (!serverHas.has(k)) {
    const u = apiUrl(office.thumbUrl(loc(e), e.mtime || ''));
    serverHas.set(k, fetch(apiUrl(url('/thumb', { ...loc(e), v: e.mtime || '', probe: 1 })), { credentials: 'same-origin' }).then((r) => (r.status === 200 && /^image\//.test(r.headers.get('content-type') || '') ? u : null), () => null));
  }
  return serverHas.get(k);
}

async function pump() {
  if (running || !queue.length) return;
  running = true;
  try {
    // no renderer mounted (yet): the jobs wait – registerRenderer() starts the queue again
    while (queue.length && renderers.size) {
      await idle();
      const job = queue.shift();
      if (!job || !job.cbs.size) continue;
      job.started = true;
      const t0 = performance.now();
      let url = null;
      if (job.make) {
        try { url = await job.make(); } catch { url = null; }
        job.cbs.forEach((cb) => { try { cb(url); } catch { /* ignore */ } });
        await new Promise((r) => setTimeout(r, 60));
        continue;
      }
      try {
        url = await serverThumb(job.e);
        if (!url) url = await renderOne(job.e);
      } catch { url = null; }
      if (url) made.set(job.k, { mtime: job.e.mtime || 0, url });
      else tried.set(job.k, job.e.mtime || 0);
      job.cbs.forEach((cb) => { try { cb(url); } catch { /* one bad listener */ } });
      listeners.forEach((fn) => { try { fn(job.e, url, performance.now() - t0); } catch { /* ignore */ } });
      // give the page room between two files
      await new Promise((r) => setTimeout(r, 120));
    }
  } finally { running = false; }
}

/*
 * Thumbnails of built-in templates (New dialog, template carousel): the chunk creates the file, the renderer paints it.
 * Kept for the session (object URLs) and – small webp data URLs – in this browser's localStorage per version.
 */
const tplMade = new Map();
const TPL_KEY = 'office:tpl-thumbs';
let tplStore = null;
const tplRead = () => { if (!tplStore) { try { tplStore = JSON.parse(localStorage.getItem(TPL_KEY) || '{}') || {}; } catch { tplStore = {}; } } return tplStore; };
export function templateThumb({ kind, id, ext, lang = '', version = '', settings = {} }) {
  const k = `${kind}|${id}|${ext}|${lang}|${version}`;
  if (tplMade.has(k)) return tplMade.get(k);
  const stored = tplRead()[k];
  if (stored) { const p = Promise.resolve(stored); tplMade.set(k, p); return p; }
  const p = new Promise((resolve) => {
    queue.push({ k: `tpl:${k}`, cbs: new Set([resolve]), make: async () => {
      const chunk = await loadEditor(kind);
      if (!chunk || typeof chunk.create !== 'function') return null;
      const bytes = await chunk.create({ ext, template: id, settings });
      const blob = await paint(chunk, bytes, ext, `template.${ext}`);
      if (!blob) return null;
      const url = await new Promise((r) => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.onerror = () => r(null); fr.readAsDataURL(blob); });
      if (url && url.length < 60000) { try { const all = tplRead(); all[k] = url; const keys = Object.keys(all); if (keys.length > 80) delete all[keys[0]]; localStorage.setItem(TPL_KEY, JSON.stringify(all)); } catch { /* full */ } }
      return url;
    } });
    pump();
  });
  tplMade.set(k, p);
  return p;
}

async function paint(chunk, bytes, ext, name) {
  let blob = null;
  if (chunk && typeof chunk.thumbnail === 'function') {
    try { blob = await toBlob(await chunk.thumbnail({ bytes, ext, width: THUMB_W })); } catch { blob = null; }
  }
  if (!blob && chunk && chunk.Preview) {
    const r = [...renderers].pop();
    if (!r) return null;
    blob = await r({ chunk, bytes, ext, name });
  }
  return blob;
}

async function renderOne(e) {
  const fmt = FORMATS[extOf(e.name || e.path)];
  const { bytes, mtime } = await office.read(loc(e));
  const chunk = await loadEditor(fmt.kind);
  const blob = await paint(chunk, bytes, extOf(e.name || e.path), e.name || e.path);
  if (!blob) return null;
  const url = URL.createObjectURL(blob);
  if (serverStores) {
    office.putThumb(loc(e), blob, mtime || e.mtime).catch((er) => { if (isMissing(er)) serverStores = false; });
  }
  return url;
}

async function toBlob(x) {
  if (!x) return null;
  if (x instanceof Blob) return x;
  if (typeof HTMLCanvasElement !== 'undefined' && x instanceof HTMLCanvasElement) return canvasBlob(fitCanvas(x));
  return null;
}
export function canvasBlob(c) {
  return new Promise((res) => { try { c.toBlob((b) => res(b), 'image/webp', 0.82); } catch { res(null); } });
}
/** scale a canvas down to THUMB_W (keeps the aspect, at most 1.5 × as tall as wide) */
export function fitCanvas(src, w = THUMB_W) {
  const ratio = Math.min(1.5, src.height / Math.max(1, src.width));
  const c = document.createElement('canvas');
  c.width = w; c.height = Math.round(w * ratio);
  const g = c.getContext('2d');
  g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  const sh = Math.min(src.height, src.width * ratio);
  g.drawImage(src, 0, 0, src.width, sh, 0, 0, c.width, c.height);
  return c;
}

/* ---------------- DOM → image ---------------- */
// computed properties worth copying – enough for paper, text, tables, slides (SVG) and images
const PROPS = ['display', 'position', 'top', 'left', 'right', 'bottom', 'width', 'height', 'min-width', 'min-height', 'max-width', 'max-height',
  'box-sizing', 'margin-top', 'margin-right', 'margin-bottom', 'margin-left', 'padding-top', 'padding-right', 'padding-bottom', 'padding-left',
  'border-top', 'border-right', 'border-bottom', 'border-left', 'border-radius', 'border-collapse', 'background-color', 'background-image',
  'color', 'font-family', 'font-size', 'font-weight', 'font-style', 'line-height', 'letter-spacing', 'text-align', 'text-decoration', 'text-transform',
  'white-space', 'vertical-align', 'overflow', 'opacity', 'transform', 'transform-origin', 'flex-direction', 'flex-wrap', 'align-items',
  'justify-content', 'gap', 'flex', 'grid-template-columns', 'grid-template-rows', 'table-layout', 'list-style-type', 'z-index',
  'fill', 'stroke', 'stroke-width', 'font-variant', 'text-indent', 'word-break', 'overflow-wrap', 'box-shadow', 'aspect-ratio', 'object-fit'];
const MAX_NODES = 2500;

/** the biggest canvas inside el (pdf.js pages), if it covers most of the preview */
function bigCanvas(el) {
  let best = null;
  for (const c of el.querySelectorAll('canvas')) if (c.width * c.height > (best ? best.width * best.height : 40000)) best = c;
  return best;
}

/** el → Blob (webp) of its first page/visible area, THUMB_W wide. Yields between batches (no long tasks). */
export async function rasterize(el) {
  const cv = bigCanvas(el);
  if (cv) { try { return await canvasBlob(fitCanvas(cv)); } catch { /* tainted – fall back to the DOM */ } }
  const page = el.firstElementChild || el;
  const box = page.getBoundingClientRect();
  const w = Math.max(1, Math.round(box.width));
  const h = Math.max(1, Math.round(Math.min(box.height, box.width * 1.5)));
  const clone = page.cloneNode(true);
  const src = [page, ...page.querySelectorAll('*')];
  const dst = [clone, ...clone.querySelectorAll('*')];
  for (let i = 0; i < src.length; i++) {
    if (i && i % 300 === 0) await new Promise((r) => setTimeout(r, 0));
    const d = dst[i];
    if (!d) break;
    if (i > MAX_NODES) { d.remove(); continue; }
    const s = src[i];
    const cs = getComputedStyle(s);
    let css = '';
    for (const p of PROPS) { const v = cs.getPropertyValue(p); if (v && v !== 'normal' && v !== 'auto' && v !== 'none' && v !== '0px') css += `${p}:${v};`; else if (p === 'display' && v === 'none') css += 'display:none;'; }
    if (/url\(/.test(css) && !/url\("?data:/.test(css)) css = css.replace(/background-image:[^;]*;/, '');
    d.setAttribute('style', css);
    d.removeAttribute('class');
    if (s.tagName === 'IMG') {
      try {
        if (!/^data:/.test(s.src) && s.complete && s.naturalWidth) { const c = document.createElement('canvas'); c.width = s.naturalWidth; c.height = s.naturalHeight; c.getContext('2d').drawImage(s, 0, 0); d.setAttribute('src', c.toDataURL('image/png')); }
      } catch { d.removeAttribute('src'); }
    }
    if (s.tagName === 'CANVAS') {
      try { const img = document.createElement('img'); img.setAttribute('src', s.toDataURL('image/png')); img.setAttribute('style', css); d.replaceWith(img); dst[i] = img; } catch { /* tainted */ }
    }
  }
  clone.style.margin = '0';
  clone.style.transform = 'none';
  const xml = new XMLSerializer().serializeToString(clone);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><foreignObject x="0" y="0" width="${w}" height="${h}"><div xmlns="http://www.w3.org/1999/xhtml" style="width:${w}px;height:${h}px;overflow:hidden;background:#fff">${xml}</div></foreignObject></svg>`;
  const img = new Image();
  img.decoding = 'async';
  await new Promise((res, rej) => { img.onload = res; img.onerror = () => rej(new Error('svg')); img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`; });
  const c = document.createElement('canvas');
  const scale = THUMB_W / w;
  c.width = THUMB_W; c.height = Math.round(h * scale);
  const g = c.getContext('2d');
  g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  g.drawImage(img, 0, 0, c.width, c.height);
  return canvasBlob(c);
}

/** wait until el stops changing (quiet for `quiet` ms, at most `max` ms), then for its images and the fonts */
export function settled(el, { quiet = 300, max = 5000 } = {}) {
  return new Promise((resolve) => {
    let t = null;
    const done = async () => {
      mo.disconnect(); clearTimeout(t); clearTimeout(hard);
      try { await Promise.race([document.fonts ? document.fonts.ready : null, new Promise((r) => setTimeout(r, 800))]); } catch { /* ignore */ }
      await Promise.all([...el.querySelectorAll('img')].map((i) => (i.complete ? null : new Promise((r) => { i.onload = r; i.onerror = r; setTimeout(r, 1500); }))));
      resolve();
    };
    const kick = () => { clearTimeout(t); t = setTimeout(done, quiet); };
    const mo = new MutationObserver(kick);
    mo.observe(el, { subtree: true, childList: true, attributes: true, characterData: true });
    const hard = setTimeout(done, max);
    kick();
  });
}
