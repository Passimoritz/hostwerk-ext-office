/*
 * OCR with the tesseract engine of the "ocr" component (tesseract.js-core: emscripten build of Tesseract 5 + deu/eng
 * LSTM data). The engine runs in a classic Web Worker created from a Blob (no eval; the component's script is
 * loaded with importScripts from Hostwerk's own address). Pure part: parseTsv (tested in tests/pdf-ocr.test.js).
 *
 *   parseTsv(tsv, minConf)      → [{ text, conf, x0, y0, x1, y1, line }]   (pixel boxes, top-left origin)
 *   createOcr({ core, wasm?, data: { deu: url, eng: url }, langs }) → { recognize(rgba, w, h) → words, close() }
 */

/** Tesseract TSV (level page/block/par/line/word …) → words */
export function parseTsv(tsv, minConf = 25) {
  const out = [];
  for (const line of String(tsv || '').split(/\r?\n/)) {
    const c = line.split('\t');
    if (c.length < 12 || c[0] !== '5') continue;
    const text = c.slice(11).join('\t').trim();
    const conf = Number(c[10]);
    if (!text || !(conf >= minConf)) continue;
    const x = Number(c[6]); const y = Number(c[7]); const w = Number(c[8]); const h = Number(c[9]);
    if (!(w > 0 && h > 0)) continue;
    out.push({ text, conf, x0: x, y0: y, x1: x + w, y1: y + h, line: `${c[2]}.${c[3]}.${c[4]}` });
  }
  return out;
}

/** words of one line share one height – looks calmer when the text is selected in other programs */
export function evenLines(words) {
  const by = new Map();
  for (const w of words) { if (!by.has(w.line)) by.set(w.line, []); by.get(w.line).push(w); }
  for (const list of by.values()) {
    const y0 = Math.min(...list.map((w) => w.y0)); const y1 = Math.max(...list.map((w) => w.y1));
    for (const w of list) { w.y0 = y0; w.y1 = y1; }
  }
  return words;
}

const WORKER = `
let core = null, api = null;
const post = (m) => self.postMessage(m);
self.onmessage = async (e) => {
  const m = e.data;
  try {
    if (m.type === 'init') {
      importScripts(m.core);
      const factory = self.TesseractCore || self.createTesseractCore;
      if (!factory) throw new Error('TesseractCore missing');
      core = await factory({ locateFile: (p) => (m.wasm && /\\.wasm$/.test(p) ? m.wasm : new URL(p, m.core).href), print: () => {}, printErr: () => {} });
      try { core.FS.mkdir('/tessdata'); } catch (x) {}
      for (const [lang, bytes] of m.data) core.FS.writeFile('/tessdata/' + lang + '.traineddata', new Uint8Array(bytes));
      api = new core.TessBaseAPI();
      let st;
      try { st = api.Init('/tessdata', m.langs, 1); } catch (x) { st = api.Init('/tessdata', m.langs); }
      if (st) throw new Error('Tesseract Init failed (' + st + ')');
      post({ type: 'ready' });
    } else if (m.type === 'recognize') {
      const px = new Uint8Array(m.rgba);
      const ptr = core._malloc(px.length);
      core.HEAPU8.set(px, ptr);
      api.SetImage(ptr, m.w, m.h, 4, m.w * 4);
      try { api.SetSourceResolution(m.dpi || 300); } catch (x) {}
      api.Recognize(null);
      let tsv = api.GetTSVText(0);
      if (typeof tsv === 'number') tsv = core.UTF8ToString(tsv);
      api.Clear();
      core._free(ptr);
      post({ type: 'result', id: m.id, tsv });
    }
  } catch (err) { post({ type: 'error', id: m.id, error: String(err && err.message || err) }); }
};
`;

export function createOcr({ core, wasm, data, langs = 'deu+eng' }) {
  const url = URL.createObjectURL(new Blob([WORKER], { type: 'text/javascript' }));
  const w = new Worker(url);
  URL.revokeObjectURL(url);
  let seq = 0; const wait = new Map();
  let ready; let fail;
  const readyP = new Promise((res, rej) => { ready = res; fail = rej; });
  w.onmessage = (e) => {
    const m = e.data;
    if (m.type === 'ready') ready();
    else if (m.type === 'error' && m.id == null) fail(new Error(m.error));
    else if (wait.has(m.id)) { const p = wait.get(m.id); wait.delete(m.id); if (m.type === 'error') p.rej(new Error(m.error)); else p.res(m.tsv); }
  };
  w.onerror = (e) => { const err = new Error(e.message || 'OCR worker failed'); fail(err); for (const p of wait.values()) p.rej(err); wait.clear(); };
  const abs = (u) => new URL(u, location.href).href;
  const entries = Object.entries(data).map(([k, v]) => [k, v]);
  w.postMessage({ type: 'init', core: abs(core), wasm: wasm ? abs(wasm) : null, data: entries, langs }, entries.map(([, v]) => v));
  return {
    ready: readyP,
    async recognize(rgba, width, height, dpi) {
      await readyP;
      const id = ++seq;
      const p = new Promise((res, rej) => wait.set(id, { res, rej }));
      w.postMessage({ type: 'recognize', id, rgba: rgba.buffer, w: width, h: height, dpi }, [rgba.buffer]);
      return evenLines(parseTsv(await p));
    },
    close() { w.terminate(); },
  };
}

/** words of tesseract.js' result (tsv when present, else the blocks tree) */
export function wordsFromData(data) {
  if (data && typeof data.tsv === 'string' && data.tsv.trim()) return evenLines(parseTsv(data.tsv));
  const out = [];
  let li = 0;
  for (const b of (data && data.blocks) || []) {
    for (const p of b.paragraphs || []) {
      for (const l of p.lines || []) {
        li++;
        for (const w of l.words || []) {
          if (!w.text || !w.text.trim() || !(w.confidence >= 25) || !w.bbox) continue;
          out.push({ text: w.text.trim(), conf: w.confidence, x0: w.bbox.x0, y0: w.bbox.y0, x1: w.bbox.x1, y1: w.bbox.y1, line: String(li) });
        }
      }
    }
  }
  return evenLines(out);
}

function loadScript(url) {
  return new Promise((res, rej) => {
    if (globalThis.Tesseract) { res(globalThis.Tesseract); return; }
    const s = document.createElement('script');
    s.src = url; s.async = true;
    s.onload = () => (globalThis.Tesseract ? res(globalThis.Tesseract) : rej(new Error('Tesseract missing')));
    s.onerror = () => rej(new Error('The OCR engine could not be loaded.'));
    document.head.append(s);
  });
}

/**
 * The OCR component's tesseract.js (tesseract.min.js + worker.min.js + core/ + lang/) – everything from Hostwerk's
 * own address, language data uncompressed, no caching in IndexedDB. Falls back to driving the core directly.
 *   createTesseract({ base, langs, langBases, onLog }) → { recognize(canvas) → words, close() }
 * langBases: { fra: '<base of the ocr-fra pack>/' … } – languages from the optional packs (their data is handed over
 * directly, the engine's own deu/eng come from its lang/ folder).
 */
const SIMD = (() => { try { return WebAssembly.validate(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0, 1, 5, 1, 96, 0, 1, 123, 3, 2, 1, 0, 10, 10, 1, 8, 0, 65, 0, 253, 15, 253, 98, 11])); } catch { return false; } })();
export const OCR_PACKS = ['fra', 'spa', 'ita', 'nld', 'pol', 'tur'];
export async function createTesseract({ base, langs = 'deu+eng', langBases = {}, onLog }) {
  const extra = {};
  for (const l of langs.split('+')) if (langBases[l]) { const r = await fetch(`${langBases[l]}lang/${l}.traineddata`, { credentials: 'same-origin' }); if (!r.ok) throw new Error(`OCR language ${l} is missing`); extra[l] = new Uint8Array(await r.arrayBuffer()); }
  const abs = (u) => new URL(u, location.href).href;
  // the component ships the (simd-)lstm builds – name the file, tesseract.js would otherwise look for other variants
  const coreFile = `core/tesseract-core-${SIMD ? 'simd-' : ''}lstm.wasm.js`;
  try {
    const T = await loadScript(abs(`${base}tesseract.min.js`));
    const worker = await T.createWorker(langs.split('+').map((l) => (extra[l] ? { code: l, data: extra[l] } : l)), 1, {
      workerPath: abs(`${base}worker.min.js`), corePath: abs(`${base}${coreFile}`), langPath: abs(`${base}lang`), gzip: false, cacheMethod: 'none',
      logger: (m) => { if (onLog) onLog(m); },
    });
    return {
      async recognize(canvas) { const { data } = await worker.recognize(canvas, {}, { tsv: true, blocks: true, text: false }); return wordsFromData(data); },
      close() { worker.terminate(); },
    };
  } catch (e) {
    // direct use of the core (no tesseract.js wrapper)
    const lang = {};
    for (const l of langs.split('+')) { if (extra[l]) { lang[l] = extra[l].buffer; continue; } const r = await fetch(`${base}lang/${l}.traineddata`, { credentials: 'same-origin' }); if (!r.ok) throw e; lang[l] = await r.arrayBuffer(); }
    const ocr = createOcr({ core: `${base}${coreFile}`, data: lang, langs });
    await ocr.ready;
    return {
      async recognize(canvas) { const ctx = canvas.getContext('2d'); const img = ctx.getImageData(0, 0, canvas.width, canvas.height); return ocr.recognize(new Uint8Array(img.data.buffer), canvas.width, canvas.height, 300); },
      close() { ocr.close(); },
    };
  }
}
