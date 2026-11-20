/*
 * RTF bodies of Outlook .msg files – pure, tested in tests/viewer-rtf.test.js.
 *
 *   decompressRtf(bytes)   → Uint8Array   compressed RTF (MS-OXRTFCP, "LZFu"/"MELA") → plain RTF bytes
 *   rtfToHtml(rtf)         → { html, text, fromHtml }   small RTF reader: paragraphs, bold/italic/underline/strike,
 *                            colours, font sizes, \u unicode, \'hh code pages, tables as rows, hyperlinks; Outlook's
 *                            encapsulated HTML (\fromhtml1, MS-OXRTFEX) is unpacked to the original HTML
 */

const PREBUF = '{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx';
export const PREBUF_LENGTH = PREBUF.length;

export function decompressRtf(bytes) {
  const d = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (d.length < 16) throw new Error('Compressed RTF is too short');
  const dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
  const compSize = dv.getUint32(0, true); const rawSize = dv.getUint32(4, true); const type = dv.getUint32(8, true);
  if (type === 0x414c454d) return d.slice(16, 16 + rawSize); // MELA – stored
  if (type !== 0x75465a4c) throw new Error('Unknown compressed RTF type');
  const dict = new Uint8Array(4096);
  for (let i = 0; i < PREBUF.length; i++) dict[i] = PREBUF.charCodeAt(i);
  let wp = PREBUF.length;
  const out = new Uint8Array(rawSize + 16); let o = 0;
  const end = Math.min(d.length, compSize + 4);
  let p = 16;
  outer: while (p < end) {
    const ctl = d[p++];
    for (let bit = 0; bit < 8; bit++) {
      if (p >= end) break outer;
      if (!(ctl & (1 << bit))) {
        const b = d[p++];
        if (o < out.length) out[o++] = b;
        dict[wp] = b; wp = (wp + 1) & 4095;
      } else {
        if (p + 1 >= end) break outer;
        const ref = (d[p] << 8) | d[p + 1]; p += 2;
        const off = ref >> 4; const len = (ref & 15) + 2;
        if (off === wp) break outer;
        for (let k = 0; k < len; k++) {
          const b = dict[(off + k) & 4095];
          if (o < out.length) out[o++] = b;
          dict[wp] = b; wp = (wp + 1) & 4095;
        }
      }
    }
  }
  return out.slice(0, Math.min(o, rawSize));
}

/** literal-only "compression" (valid LZFu, used by the tests and as a reference) */
export function compressRtfStored(bytes) {
  const out = []; let wp = PREBUF.length;
  for (let i = 0; i < bytes.length; i += 8) {
    const chunk = bytes.slice(i, i + 8);
    let ctl = 0; const body = [...chunk];
    if (chunk.length < 8) { ctl = 1 << chunk.length; body.push((wp + chunk.length) >> 4 & 0xff, (((wp + chunk.length) & 15) << 4) & 0xff); }
    out.push(ctl, ...body);
    wp = (wp + chunk.length) & 4095;
  }
  if (bytes.length % 8 === 0) out.push(1, (wp >> 4) & 0xff, (wp & 15) << 4);
  const head = new Uint8Array(16); const dv = new DataView(head.buffer);
  dv.setUint32(0, out.length + 12, true); dv.setUint32(4, bytes.length, true); dv.setUint32(8, 0x75465a4c, true);
  return Uint8Array.from([...head, ...out]);
}

/* ---------------- RTF → HTML ---------------- */

const CP = { 1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257', 1258: 'windows-1258', 874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 437: 'ibm866', 850: 'windows-1252', 10000: 'macintosh', 65001: 'utf-8' };
const CHARSET_CP = { 0: 1252, 1: 1252, 2: 1252, 77: 10000, 128: 932, 129: 949, 134: 936, 136: 950, 161: 1253, 162: 1254, 163: 1258, 177: 1255, 178: 1256, 186: 1257, 204: 1251, 222: 874, 238: 1250 };
const SKIP = new Set(['fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'header', 'footer', 'headerl', 'headerr', 'headerf', 'footerl', 'footerr', 'footerf', 'listtable', 'listoverridetable', 'revtbl', 'rsidtbl', 'generator', 'xmlnstbl', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'object', 'fldinst', 'bkmkstart', 'bkmkend', 'mmathPr', 'pgdsctbl', 'filetbl']);
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const decoders = new Map();
function decodeByte(b, cp) {
  const name = CP[cp] || 'windows-1252';
  if (!decoders.has(name)) { try { decoders.set(name, new TextDecoder(name)); } catch { decoders.set(name, new TextDecoder('windows-1252')); } }
  return decoders.get(name).decode(Uint8Array.of(b));
}

/** tokens: { t: 'open' | 'close' | 'word' | 'text' | 'hex', … } */
function* tokens(s) {
  let i = 0; const n = s.length;
  while (i < n) {
    const c = s[i];
    if (c === '{') { yield { t: 'open' }; i++; continue; }
    if (c === '}') { yield { t: 'close' }; i++; continue; }
    if (c === '\\') {
      const nx = s[i + 1];
      if (nx === "'") { yield { t: 'hex', v: parseInt(s.slice(i + 2, i + 4), 16) }; i += 4; continue; }
      if (/[a-zA-Z]/.test(nx || '')) {
        let j = i + 1; while (j < n && /[a-zA-Z]/.test(s[j])) j++;
        const word = s.slice(i + 1, j);
        let k = j; if (s[k] === '-' || /\d/.test(s[k] || '')) { k++; while (k < n && /\d/.test(s[k])) k++; }
        const param = k > j ? Number(s.slice(j, k)) : null;
        if (s[k] === ' ') k++;
        yield { t: 'word', w: word, p: param };
        i = k; continue;
      }
      // control symbols
      if (nx === '*') { yield { t: 'star' }; i += 2; continue; }
      if (nx === '~') { yield { t: 'text', v: ' ' }; i += 2; continue; }
      if (nx === '-') { i += 2; continue; }
      if (nx === '_') { yield { t: 'text', v: '‑' }; i += 2; continue; }
      if (nx === '\n' || nx === '\r') { yield { t: 'word', w: 'par', p: null }; i += 2; continue; }
      yield { t: 'text', v: nx || '' }; i += 2; continue;
    }
    if (c === '\r' || c === '\n') { i++; continue; }
    let j = i; while (j < n && !'{}\\\r\n'.includes(s[j])) j++;
    yield { t: 'text', v: s.slice(i, j) }; i = j;
  }
}

export function rtfToHtml(input) {
  const s = typeof input === 'string' ? input : Array.from(input, (b) => String.fromCharCode(b)).join('');
  const fromHtml = /\\fromhtml1/.test(s.slice(0, 2000));
  const fonts = new Map(); const colors = [];
  let cp = 1252;
  const m = /\\ansicpg(\d+)/.exec(s.slice(0, 1000)); if (m) cp = Number(m[1]);
  // font table (charsets) and colour table first
  const ft = /\{\\fonttbl([\s\S]*?)\}\s*(?:\{\\colortbl|\{\\\*|\{\\stylesheet|\\)/.exec(s);
  if (ft) for (const f of ft[1].matchAll(/\\f(\d+)[^;{}]*?\\fcharset(\d+)/g)) fonts.set(Number(f[1]), CHARSET_CP[Number(f[2])] || cp);
  const ct = /\{\\colortbl\s*;?([\s\S]*?)\}/.exec(s);
  if (ct) { colors.push(null); for (const c of ct[1].split(';')) { const r = /\\red(\d+)/.exec(c); const g = /\\green(\d+)/.exec(c); const b = /\\blue(\d+)/.exec(c); if (r || g || b) colors.push(`rgb(${r ? r[1] : 0},${g ? g[1] : 0},${b ? b[1] : 0})`); } }

  const st = { skip: 0, b: false, i: false, u: false, s: false, fs: 0, cf: 0, f: 0, uc: 1, htmlrtf: false, inHtmltag: false, inst: false };
  const stack = [];
  let html = ''; let text = ''; let ucSkip = 0;
  let pendingStar = false; let first = false;
  let open = ''; // currently open inline style span
  let field = null; // { depth } of the open link
  let pendingUrl = null;
  const styleKey = () => (fromHtml ? '' : [st.b && 'font-weight:bold', st.i && 'font-style:italic', st.u && 'text-decoration:underline', st.s && 'text-decoration:line-through', st.fs && st.fs !== 24 && `font-size:${st.fs / 2}pt`, st.cf && colors[st.cf] && `color:${colors[st.cf]}`].filter(Boolean).join(';'));
  const emit = (str) => {
    if (!str) return;
    if (fromHtml) { if (st.inHtmltag || !st.htmlrtf) html += st.inHtmltag ? str : esc(str); if (!st.htmlrtf) text += str; return; }
    const k = styleKey();
    if (k !== open) { if (open) html += '</span>'; if (k) html += `<span style="${k}">`; open = k; }
    html += esc(str); text += str;
  };
  const raw = (h) => { if (!fromHtml) { if (open) { html += '</span>'; open = ''; } html += h; } };
  let cell = false;
  for (const tk of tokens(s)) {
    if (tk.t === 'open') { stack.push({ ...st }); pendingStar = false; first = true; continue; }
    if (tk.t === 'close') {
      const prev = stack.pop();
      if (field && stack.length < field.depth) { raw('</a>'); field = null; }
      if (prev) Object.assign(st, prev);
      first = false; continue;
    }
    if (tk.t === 'star') { pendingStar = true; continue; }
    const wasFirst = first; first = false;
    if (st.skip) continue;
    if (tk.t === 'word') {
      const { w, p } = tk;
      if (wasFirst && (SKIP.has(w) || (pendingStar && !(fromHtml && (w === 'htmltag' || w === 'mhtmltag'))))) {
        if (w === 'fldinst') { st.inst = true; pendingStar = false; continue; }
        st.skip = 1; pendingStar = false; continue;
      }
      pendingStar = false;
      if (fromHtml && (w === 'htmltag' || w === 'mhtmltag')) { st.inHtmltag = true; continue; }
      switch (w) {
        case 'htmlrtf': st.htmlrtf = p !== 0; break;
        case 'par': case 'sect': if (fromHtml) { if (st.inHtmltag || !st.htmlrtf) { html += st.inHtmltag ? '\r\n' : ''; } text += '\n'; } else { raw('<br>'); text += '\n'; } break;
        case 'line': emit(fromHtml ? '\n' : ''); if (!fromHtml) raw('<br>'); break;
        case 'tab': emit('\t'); break;
        case 'cell': raw('</td><td>'); text += '\t'; cell = true; break;
        case 'row': raw('</td></tr><tr><td>'); text += '\n'; break;
        case 'trowd': if (!cell) raw(''); break;
        case 'b': st.b = p !== 0; break;
        case 'i': st.i = p !== 0; break;
        case 'ul': st.u = p !== 0; break;
        case 'ulnone': st.u = false; break;
        case 'strike': st.s = p !== 0; break;
        case 'plain': st.b = st.i = st.u = st.s = false; st.fs = 0; st.cf = 0; break;
        case 'fs': st.fs = p || 0; break;
        case 'cf': st.cf = p || 0; break;
        case 'f': st.f = p || 0; break;
        case 'uc': st.uc = p == null ? 1 : p; break;
        case 'u': { let c = p; if (c < 0) c += 65536; emit(String.fromCharCode(c)); ucSkip = st.uc; break; }
        case 'emdash': emit('—'); break;
        case 'endash': emit('–'); break;
        case 'bullet': emit('•'); break;
        case 'lquote': emit('‘'); break;
        case 'rquote': emit('’'); break;
        case 'ldblquote': emit('“'); break;
        case 'rdblquote': emit('”'); break;
        case 'field': break;
        case 'fldrslt': if (pendingUrl && !fromHtml) { raw(`<a href="${esc(pendingUrl)}">`); field = { depth: stack.length }; pendingUrl = null; } break;
        default: break;
      }
      continue;
    }
    if (tk.t === 'hex') {
      if (ucSkip > 0) { ucSkip--; continue; }
      if (fromHtml && st.htmlrtf && !st.inHtmltag) continue;
      emit(decodeByte(tk.v, fonts.get(st.f) || cp));
      continue;
    }
    if (tk.t === 'text') {
      let v = tk.v;
      if (ucSkip > 0) { const k = Math.min(ucSkip, v.length); v = v.slice(k); ucSkip -= k; }
      if (st.inst) { const hl = /HYPERLINK\s+"([^"]+)"/.exec(v); if (hl && /^(https?:|mailto:)/i.test(hl[1])) pendingUrl = hl[1]; continue; }
      emit(v);
    }
  }
  if (open) html += '</span>';
  if (fromHtml) return { html, text: text.replace(/\n{3,}/g, '\n\n').trim(), fromHtml: true };
  html = html.replace(/<\/td><tr><td>$/, '');
  if (cell) html = `<table class="table of-view-rtftable"><tr><td>${html}</td></tr></table>`.replace(/<tr><td><\/td><\/tr><\/table>$/, '</table>');
  return { html: `<div class="of-view-rtf">${html}</div>`, text: text.replace(/\n{3,}/g, '\n\n').trim(), fromHtml: false };
}
