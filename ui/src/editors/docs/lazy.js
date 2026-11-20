/*
 * Heavy parts of the docs editor load on demand – they are not in docs.js:
 *   math()     → ui/assets/docs-math.js (temml, LaTeX → MathML) next to this chunk, same ?v=
 *   compareLib() → the version compare, in the same sub-chunk
 *   mermaid(h) → the "mermaid" component of the Office backend (host.components.ensure – downloaded once, hash-pinned)
 * Each loads once; failures are cached for a short while only (offline → try again later).
 */
const cache = {};
function sibling(file) {
  const me = new URL(import.meta.url);
  const u = new URL(file, me);
  u.search = me.search;
  return u.href;
}
function once(key, fn) {
  if (cache[key] && !cache[key].failed) return cache[key].p;
  const entry = { p: null, failed: false };
  entry.p = fn().catch((e) => { entry.failed = true; throw e; });
  cache[key] = entry;
  return entry.p;
}

export const math = () => once('math', async () => (await import(/* @vite-ignore */ sibling('docs-math.js'))).default);
/** compareDocs / applyAsTrackedChanges (diff-match-patch) */
export const compareLib = () => math().then((m) => m.compare);
/** the Word dialogs (paragraph, font, table properties, borders, sort) */
export const wordPack = () => math().then((m) => m.word);

/** LaTeX → MathML (null + error text when it cannot be parsed) */
export async function renderLatex(latex, display = false) {
  const m = await math();
  return m.tryLatexToMathml(latex, { display });
}

/* ---------------- Mermaid (component on demand) ---------------- */
let mermaidReady = null;
function loadScript(src) {
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = src; s.async = true;
    s.onload = () => res(); s.onerror = () => rej(new Error(`Could not load ${src}`));
    document.head.append(s);
  });
}
/** the mermaid API – asks the shell to provide the component first (it asks the user once, with the size) */
export function mermaid(host, { ask = true } = {}) {
  if (mermaidReady) return mermaidReady;
  const p = (async () => {
    if (typeof window !== 'undefined' && window.mermaid && window.mermaid.render) return window.mermaid;
    const comp = host && host.components;
    if (!comp || typeof comp.ensure !== 'function') throw Object.assign(new Error('Diagrams need the "Mermaid" component – this Hostwerk version cannot provide it yet.'), { code: 'NO_COMPONENTS' });
    const r = await comp.ensure('mermaid', { ask });
    if (!r) throw Object.assign(new Error('The diagram component is not installed.'), { code: 'NOT_INSTALLED' });
    const url = typeof r === 'string' ? r : r.url || r.main || (r.base ? `${String(r.base).replace(/\/$/, '')}/${r.file || 'mermaid.min.js'}` : '/api/office/components/mermaid/mermaid.min.js');
    let api = null;
    if (/\.m?js(\?|$)/.test(url) && /esm|\.mjs/.test(url)) { const mod = await import(/* @vite-ignore */ url); api = mod.default || mod; } else { await loadScript(url); api = window.mermaid; }
    if (!api || !api.render) throw new Error('The diagram component could not be started.');
    api.initialize({ startOnLoad: false, securityLevel: 'strict', theme: 'default', fontFamily: 'Calibri, Carlito, "Segoe UI", sans-serif', htmlLabels: false, flowchart: { htmlLabels: false } });
    return api;
  })();
  mermaidReady = p.catch((e) => { mermaidReady = null; throw e; });
  return mermaidReady;
}
let mid = 0;
/** source → { svg, width, height } */
export async function renderMermaid(host, source, opts) {
  const api = await mermaid(host, opts);
  const { svg } = await api.render(`of-mmd-${Date.now().toString(36)}${mid++}`, String(source || ''));
  const m = /viewBox="[\d.\-]+ [\d.\-]+ ([\d.]+) ([\d.]+)"/.exec(svg);
  return { svg, width: m ? Math.round(Number(m[1])) : 400, height: m ? Math.round(Number(m[2])) : 300 };
}
/** SVG text → PNG data URL (what Word / Google show; 2× for sharpness) */
export async function svgToPng(svg, width, height, scale = 2) {
  const img = new Image();
  const src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg.replace(/<svg\b/, (s) => (/\swidth=/.test(svg.slice(0, 300)) ? s : `${s} width="${width}" height="${height}"`)))}`;
  img.src = src;
  await img.decode();
  const c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(width * scale)); c.height = Math.max(1, Math.round(height * scale));
  const g = c.getContext('2d');
  g.fillStyle = '#ffffff'; g.fillRect(0, 0, c.width, c.height);
  g.drawImage(img, 0, 0, c.width, c.height);
  return c.toDataURL('image/png');
}
