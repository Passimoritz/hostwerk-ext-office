'use strict';
/*
 * Components on demand: heavy parts that do not fit into the 10 MB package (OCR engine + language data, Mermaid,
 * pdf.js CMaps + standard fonts). Downloaded once from PINNED urls, every download checked against a PINNED SHA-256
 * (nothing is written, unpacked or served without a matching hash), unpacked into <data>/office-components/<id>/,
 * served read-only. Nothing here is ever executed on the server – the files are only handed to the browser.
 *
 *   new Components({ dir, fetchImpl, offline, log, broadcast })
 *   .list() → [{ id, label, license, size, download, state: 'missing'|'installing'|'ready'|'error', progress, error, version }]
 *   .install(id) → starts (Promise, progress via broadcast { type: 'office-component', id, state, progress })
 *   .remove(id) · .file(id, name) → { full, type, sha256, size } | null
 *   untar(buf) → Map(name → Buffer) (ustar/pax, regular files only)
 */
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const crypto = require('crypto');

const NPM = 'https://registry.npmjs.org';
const GFONTS = 'https://fonts.gstatic.com/s';
const TESSDATA = 'https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/87416418657359cb625c412a48b6e1d6d41c29bd';
/* Hashes determined 2026-11-20 by downloading once (npm tarballs also cross-checked with the registry's sha512 integrity). */
const REGISTRY = {
  ocr: {
    label: 'Text recognition (OCR)', version: 'tesseract.js 7.0.0 · tessdata_fast 4.1 (deu, eng)', license: 'Apache-2.0',
    sources: [
      { url: `${NPM}/tesseract.js/-/tesseract.js-7.0.0.tgz`, sha256: '9a93bf51c3387f945d10a24bf8b3a4bf2e45c7c7161b8242aafbaf9d3c4b606a', size: 616435, tar: { 'package/dist/tesseract.min.js': 'tesseract.min.js', 'package/dist/worker.min.js': 'worker.min.js' } },
      { url: `${NPM}/tesseract.js-core/-/tesseract.js-core-7.0.0.tgz`, sha256: 'ba584355515eaff877552022853c0e71f2cb70466e759d1d6940484929718ee0', size: 16935308, tar: {
        'package/tesseract-core-simd-lstm.wasm.js': 'core/tesseract-core-simd-lstm.wasm.js', 'package/tesseract-core-lstm.wasm.js': 'core/tesseract-core-lstm.wasm.js',
        'package/tesseract-core-simd-lstm.js': 'core/tesseract-core-simd-lstm.js', 'package/tesseract-core-simd-lstm.wasm': 'core/tesseract-core-simd-lstm.wasm',
        'package/tesseract-core-lstm.js': 'core/tesseract-core-lstm.js', 'package/tesseract-core-lstm.wasm': 'core/tesseract-core-lstm.wasm' } },
      { url: `${TESSDATA}/deu.traineddata`, sha256: '19d219bbb6672c869d20a9636c6816a81eb9a71796cb93ebe0cb1530e2cdb22d', size: 1525436, file: 'lang/deu.traineddata' },
      { url: `${TESSDATA}/eng.traineddata`, sha256: '7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2', size: 4113088, file: 'lang/eng.traineddata' },
    ],
  },
  mermaid: {
    label: 'Diagrams from text (Mermaid)', version: 'mermaid 11.12.0', license: 'MIT',
    sources: [{ url: `${NPM}/mermaid/-/mermaid-11.12.0.tgz`, sha256: '93492e4d609be83653a41c435f4cbd7e74419f7fa6b4fc90a8c983195f64af82', size: 14365637, tar: { 'package/dist/mermaid.min.js': 'mermaid.min.js', 'package/LICENSE': 'LICENSE' } }],
  },
  'cjk-fonts': {
    label: 'PDF fonts for Asian languages (CMaps, standard fonts, Noto Sans SC/JP/KR)', version: 'pdfjs-dist 4.10.38 · Noto Sans SC v41, JP v57, KR v40', license: 'Apache-2.0 (CMaps: BSD-3-Clause Adobe, fonts: Foxit BSD / Liberation OFL, Noto Sans CJK: OFL-1.1)',
    sources: [
      { url: `${NPM}/pdfjs-dist/-/pdfjs-dist-4.10.38.tgz`, sha256: '1011b38553532d7078c59f26b15a471f8dae00f101b60e2add9b8511737a1ce0', size: 10346793, tarDirs: { 'package/cmaps/': 'cmaps/', 'package/standard_fonts/': 'standard_fonts/' } },
      // static TrueType (glyf) instances, weight 400, exactly as Google Fonts serves them (versioned URLs) – the PDF editor subsets them
      { url: `${GFONTS}/notosanssc/v41/k3kCo84MPvpLmixcA63oeAL7Iqp5IZJF9bmaG9_FnYw.ttf`, sha256: '450625c8d46ab3df97b7904ded955ec2746d17ec76740cb1e91d1ba63a0f89af', size: 10540644, file: 'NotoSansSC-Regular.ttf' },
      { url: `${GFONTS}/notosansjp/v57/-F6jfjtqLzI2JPCgQBnw7HFyzSD-AsregP8VFBEj75s.ttf`, sha256: 'f037d714109f22d57ce75fd9a6b9d23439c903061fb11106c070b2dcc03f2021', size: 5703608, file: 'NotoSansJP-Regular.ttf' },
      { url: `${GFONTS}/notosanskr/v40/PbyxFmXiEBPT4ITbgNA5Cgms3VYcOA-vvnIzzuoyeLQ.ttf`, sha256: '84e2b8acadd0a01af9f0de6b0a2d29117fb5b0eabd7c2f8353d9d2da0461d5fa', size: 6170740, file: 'NotoSansKR-Regular.ttf' },
      { url: 'https://raw.githubusercontent.com/google/fonts/a85815a42757630ce188fdad368c2dfc444d4773/ofl/notosanssc/OFL.txt', sha256: '1c05c68c34f9708415aada51f17e1b0092d2cea709bf4a94cd38114f9e73d7d9', size: 4388, file: 'OFL.txt' },
    ],
  },
};
/** other names the UI may use for the same component */
const ALIAS = { 'pdf-cmaps': 'cjk-fonts', 'pdf-fonts': 'cjk-fonts', cjk: 'cjk-fonts', 'ocr-deu': 'ocr', 'ocr-eng': 'ocr', tesseract: 'ocr' };
const TYPES = { js: 'text/javascript; charset=utf-8', mjs: 'text/javascript; charset=utf-8', wasm: 'application/wasm', traineddata: 'application/octet-stream', bcmap: 'application/octet-stream', pfb: 'application/octet-stream', ttf: 'font/ttf', txt: 'text/plain; charset=utf-8', '': 'text/plain; charset=utf-8' };
const SAFE_FILE = /^[a-z0-9_-]+(\/[A-Za-z0-9_.-]+){0,2}$|^[A-Za-z0-9_.-]+$/;
const sha = (b) => crypto.createHash('sha256').update(b).digest('hex');

/** Regular files of a tar archive (ustar + pax/gnu long names) */
function untar(buf, { max = 400 * 1048576 } = {}) {
  const out = new Map(); let p = 0; let longName = null; let total = 0;
  while (p + 512 <= buf.length) {
    const h = buf.subarray(p, p + 512);
    if (h.every((b) => b === 0)) break;
    const str = (o, l) => h.toString('utf8', o, o + l).replace(/\0[\s\S]*$/, '');
    const size = parseInt(str(124, 12).trim() || '0', 8);
    if (!Number.isFinite(size) || size < 0) throw new Error('Broken tar archive');
    const type = String.fromCharCode(h[156] || 48);
    const prefix = h.toString('utf8', 257, 263).startsWith('ustar') ? str(345, 155) : '';
    let name = longName || (prefix ? `${prefix}/${str(0, 100)}` : str(0, 100));
    longName = null;
    const body = buf.subarray(p + 512, p + 512 + size);
    if (type === 'L') longName = body.toString('utf8').replace(/\0[\s\S]*$/, '');
    else if (type === 'x') { const m = /\d+ path=([^\n]*)\n/.exec(body.toString('utf8')); if (m) longName = m[1]; }
    else if (type === '0' || type === '\0' || type === '7') { total += size; if (total > max) throw new Error('Archive too large'); name = name.replace(/^\.\//, ''); out.set(name, Buffer.from(body)); }
    p += 512 + Math.ceil(size / 512) * 512;
  }
  return out;
}

class Components {
  constructor({ dir, fetchImpl = null, offline = () => false, log = () => {}, broadcast = () => {}, registry = REGISTRY }) {
    this.dir = dir; this.fetch = fetchImpl || ((u, o) => fetch(u, o)); this.offline = offline; this.log = log; this.bc = broadcast; this.reg = registry;
    this.jobs = new Map();
  }
  canon(id) { const s = String(id || ''); return Object.prototype.hasOwnProperty.call(ALIAS, s) && this.reg[ALIAS[s]] ? ALIAS[s] : s; }
  def(id0) { const id = this.canon(id0); const d = Object.prototype.hasOwnProperty.call(this.reg, id) ? this.reg[id] : null; if (!d) throw Object.assign(new Error('Unknown component'), { status: 404 }); return d; }
  root(id0) { const id = this.canon(id0); this.def(id); return path.join(this.dir, id); }
  manifest(id0) { const id = this.canon(id0); try { const m = JSON.parse(fs.readFileSync(path.join(this.root(id), 'installed.json'), 'utf8')); return m && m.version === this.def(id).version ? m : null; } catch { return null; } }
  view(id0) {
    const id = this.canon(id0); const d = this.def(id); const j = this.jobs.get(id); const m = this.manifest(id);
    return { id, label: d.label, version: d.version, license: d.license, download: d.sources.reduce((s, x) => s + x.size, 0), size: m ? m.size : null,
      installed: !!m, aliases: Object.keys(ALIAS).filter((k) => ALIAS[k] === id),
      state: j && j.state === 'installing' ? 'installing' : m ? 'ready' : j && j.state === 'error' ? 'error' : 'missing', progress: j ? j.progress : m ? 100 : 0, error: j && j.state === 'error' ? j.error : null,
      files: m ? Object.keys(m.files) : [], fileCount: m ? Object.keys(m.files).length : 0, installedAt: m ? m.at : null, base: `/api/office/components/${id}/` };
  }
  list() { return Object.keys(this.reg).map((id) => this.view(id)); }
  status(id, patch) { const j = { ...(this.jobs.get(id) || {}), ...patch }; this.jobs.set(id, j); try { this.bc({ type: 'office-component', id, state: j.state, progress: j.progress, error: j.error || null }); } catch { /* egal */ } }

  /** Download with a size cap, hash while streaming; resolves with the verified bytes */
  async download(src, onBytes) {
    if (this.offline()) throw Object.assign(new Error('Hostwerk is in offline mode – downloads are switched off.'), { status: 409 });
    const r = await this.fetch(src.url, { redirect: 'follow', signal: AbortSignal.timeout(15 * 60000), headers: { 'User-Agent': 'Hostwerk-Office' } });
    if (!r.ok) throw new Error(`HTTP ${r.status} for ${new URL(src.url).hostname}`);
    const h = crypto.createHash('sha256'); const parts = []; let n = 0;
    const cap = src.size + 1024;
    if (r.body && typeof r.body.getReader === 'function') {
      const rd = r.body.getReader();
      for (;;) { const { done, value } = await rd.read(); if (done) break; n += value.length; if (n > cap) { try { rd.cancel(); } catch { /* egal */ } throw new Error('The download is larger than expected – stopped.'); } h.update(value); parts.push(Buffer.from(value)); onBytes(value.length); }
    } else { const b = Buffer.from(await r.arrayBuffer()); n = b.length; if (n > cap) throw new Error('The download is larger than expected – stopped.'); h.update(b); parts.push(b); onBytes(b.length); }
    const got = h.digest('hex');
    if (got !== src.sha256) throw Object.assign(new Error('The download does not match its pinned checksum – nothing was installed.'), { code: 'EHASH', got });
    return Buffer.concat(parts);
  }

  install(id0) {
    const id = this.canon(id0); const d = this.def(id);
    const cur = this.jobs.get(id);
    if (cur && cur.state === 'installing') return cur.promise;
    const total = d.sources.reduce((s, x) => s + x.size, 0);
    let done = 0;
    this.status(id, { state: 'installing', progress: 0, error: null });
    const promise = (async () => {
      const files = {}; // dest → Buffer
      let last = 0;
      for (const src of d.sources) {
        const buf = await this.download(src, (k) => { done += k; const p = Math.min(95, Math.floor((done / total) * 95)); if (p !== last) { last = p; this.status(id, { progress: p }); } });
        if (src.file) files[src.file] = buf;
        else {
          const tar = untar(zlib.gunzipSync(buf, { maxOutputLength: 400 * 1048576 }));
          for (const [from, to] of Object.entries(src.tar || {})) { if (!tar.has(from)) throw new Error(`File ${from} missing in the download`); files[to] = tar.get(from); }
          for (const [pre, to] of Object.entries(src.tarDirs || {})) for (const [n, b] of tar) if (n.startsWith(pre) && !n.endsWith('/') && !n.slice(pre.length).includes('/')) files[to + n.slice(pre.length)] = b;
        }
      }
      for (const n of Object.keys(files)) if (!SAFE_FILE.test(n) || n.split('/').includes('..')) throw new Error(`Unsafe file name ${n}`);
      // write into a fresh folder, then swap (an old installation stays until the new one is complete)
      const root = this.root(id); const tmp = `${root}.~new`;
      fs.rmSync(tmp, { recursive: true, force: true });
      const man = { id, version: d.version, at: Date.now(), size: 0, files: {} };
      for (const [n, b] of Object.entries(files)) { const f = path.join(tmp, ...n.split('/')); fs.mkdirSync(path.dirname(f), { recursive: true }); fs.writeFileSync(f, b); man.files[n] = { sha256: sha(b), size: b.length }; man.size += b.length; }
      fs.writeFileSync(path.join(tmp, 'installed.json'), JSON.stringify(man));
      fs.rmSync(root, { recursive: true, force: true });
      fs.renameSync(tmp, root);
      this.status(id, { state: 'ready', progress: 100 });
      this.log('component installed', id, d.version);
      return this.view(id);
    })().catch((e) => { this.status(id, { state: 'error', error: e.message }); this.log('component install failed', id, e.message); throw e; });
    promise.catch(() => {});
    this.jobs.get(id).promise = promise;
    return promise;
  }
  remove(id0) { const id = this.canon(id0); const root = this.root(id); this.jobs.delete(id); fs.rmSync(root, { recursive: true, force: true }); fs.rmSync(`${root}.~new`, { recursive: true, force: true }); return this.view(id); }
  /** A served file: only names listed in installed.json (the hashes were checked on install) */
  file(id0, name) {
    const id = this.canon(id0); const m = this.manifest(id);
    const n = String(name || '');
    if (!m || !Object.prototype.hasOwnProperty.call(m.files, n) || !SAFE_FILE.test(n)) return null;
    const full = path.join(this.root(id), ...n.split('/'));
    const ext = (/\.([a-z0-9]+)$/i.exec(n) || [])[1] || '';
    return { full, type: TYPES[ext.toLowerCase()] || 'application/octet-stream', ...m.files[n] };
  }
}

module.exports = { Components, REGISTRY, untar, TYPES, ALIAS };
