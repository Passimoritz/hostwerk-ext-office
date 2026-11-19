/*
 * CSV / TSV: text decoding with BOM + encoding detection, delimiter sniffing, RFC 4180 parse and write.
 * Pure – no DOM, works in Node and the browser.
 */

// windows-1252 bytes 0x80–0x9F (undefined positions → U+FFFD)
const CP1252 = [0x20AC, 0xFFFD, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0xFFFD, 0x017D, 0xFFFD,
  0xFFFD, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0xFFFD, 0x017E, 0x0178];
const CP1252_REV = new Map(CP1252.map((u, i) => [u, 0x80 + i]).filter(([u]) => u !== 0xFFFD));

const norm = (e) => {
  const x = String(e || 'auto').toLowerCase().replace(/_/g, '-');
  if (['utf8', 'utf-8'].includes(x)) return 'utf-8';
  if (['cp1252', 'windows1252', 'windows-1252', 'ansi'].includes(x)) return 'windows-1252';
  if (['latin1', 'latin-1', 'iso-8859-1', 'iso8859-1'].includes(x)) return 'iso-8859-1';
  if (['utf-16', 'utf16', 'utf-16le', 'utf16le'].includes(x)) return 'utf-16le';
  if (['utf-16be', 'utf16be'].includes(x)) return 'utf-16be';
  if (x === 'utf-8-bom') return 'utf-8';
  return x;
};

function manualSingleByte(bytes, cp1252) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 8192) {
    const part = bytes.subarray(i, i + 8192);
    const codes = new Array(part.length);
    for (let j = 0; j < part.length; j++) { const b = part[j]; codes[j] = cp1252 && b >= 0x80 && b <= 0x9F ? CP1252[b - 0x80] : b; }
    s += String.fromCharCode(...codes);
  }
  return s;
}

function decodeWith(bytes, enc) {
  if (enc === 'windows-1252' || enc === 'iso-8859-1') return manualSingleByte(bytes, enc === 'windows-1252');
  try { return new TextDecoder(enc).decode(bytes); } catch { /* label unknown */ }
  if (enc === 'utf-16le' || enc === 'utf-16be') {
    let s = '';
    for (let i = 0; i + 1 < bytes.length; i += 2) s += String.fromCharCode(enc === 'utf-16le' ? bytes[i] | (bytes[i + 1] << 8) : (bytes[i] << 8) | bytes[i + 1]);
    return s;
  }
  return manualSingleByte(bytes, true);
}

function validUtf8(bytes) {
  try { new TextDecoder('utf-8', { fatal: true }).decode(bytes); return true; } catch { return false; }
}

/** bytes → { text, encoding } – BOM wins, 'auto' = UTF-8 if valid else windows-1252; the BOM is stripped. */
export function decodeText(bytes, encoding = 'auto') {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes || []);
  let enc = norm(encoding);
  let start = 0;
  if (b[0] === 0xEF && b[1] === 0xBB && b[2] === 0xBF) { enc = 'utf-8'; start = 3; }
  else if (b[0] === 0xFF && b[1] === 0xFE) { enc = 'utf-16le'; start = 2; }
  else if (b[0] === 0xFE && b[1] === 0xFF) { enc = 'utf-16be'; start = 2; }
  const body = b.subarray(start);
  if (enc === 'auto') enc = validUtf8(body) ? 'utf-8' : 'windows-1252';
  let text = decodeWith(body, enc);
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  return { text, encoding: enc };
}

/** text → bytes in 'utf-8' | 'utf-8-bom' | 'windows-1252' | 'iso-8859-1' | 'utf-16le' | 'utf-16be' (UTF-16 with BOM) */
export function encodeText(text, encoding = 'utf-8') {
  const s = String(text ?? '');
  const raw = String(encoding || 'utf-8').toLowerCase();
  const enc = raw === 'utf-8-bom' ? raw : norm(raw);
  if (enc === 'utf-8' || enc === 'utf-8-bom' || enc === 'auto') {
    const u = new TextEncoder().encode(s);
    if (enc !== 'utf-8-bom') return u;
    const out = new Uint8Array(u.length + 3); out.set([0xEF, 0xBB, 0xBF]); out.set(u, 3); return out;
  }
  if (enc === 'utf-16le' || enc === 'utf-16be') {
    const out = new Uint8Array(2 + s.length * 2);
    const le = enc === 'utf-16le';
    out[0] = le ? 0xFF : 0xFE; out[1] = le ? 0xFE : 0xFF;
    for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); out[2 + i * 2] = le ? c & 255 : c >> 8; out[3 + i * 2] = le ? c >> 8 : c & 255; }
    return out;
  }
  const cp = enc !== 'iso-8859-1';
  const out = [];
  for (const ch of s) {
    const c = ch.codePointAt(0);
    if (c < 0x80 || (c >= 0xA0 && c <= 0xFF) || (!cp && c <= 0xFF)) out.push(c);
    else if (cp && CP1252_REV.has(c)) out.push(CP1252_REV.get(c));
    else out.push(0x3F);
  }
  return Uint8Array.from(out);
}

const CANDIDATES = [',', ';', '\t', '|'];

/** Guess the delimiter from the first ~20 lines (counting outside quotes). */
export function detectDelimiter(text) {
  const lines = [];
  let cur = { ',': 0, ';': 0, '\t': 0, '|': 0 };
  let inQ = false;
  const s = String(text || '');
  for (let i = 0; i < s.length && lines.length < 20; i++) {
    const ch = s[i];
    if (ch === '"') inQ = !inQ;
    else if (!inQ && (ch === '\n' || ch === '\r')) {
      if (ch === '\r' && s[i + 1] === '\n') i++;
      lines.push(cur); cur = { ',': 0, ';': 0, '\t': 0, '|': 0 };
    } else if (!inQ && ch in cur) cur[ch]++;
  }
  if (lines.length < 20 && Object.values(cur).some((n) => n > 0)) lines.push(cur);
  let best = ','; let bestScore = 0;
  for (const d of CANDIDATES) {
    const counts = lines.map((l) => l[d]);
    const withAny = counts.filter((n) => n > 0);
    if (!withAny.length) continue;
    // most frequent count among lines → consistency
    const freq = new Map();
    for (const n of withAny) freq.set(n, (freq.get(n) || 0) + 1);
    const [mode, hits] = [...freq.entries()].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0];
    const score = (hits / lines.length) * 100 + Math.min(mode, 50) / 10;
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}

/** RFC 4180 parse → string[][]; ragged rows allowed, a trailing newline adds no row. */
export function parseCSV(text, { delimiter = 'auto', quote = '"', maxRows = Infinity } = {}) {
  const s = String(text ?? '');
  const d = !delimiter || delimiter === 'auto' ? detectDelimiter(s) : delimiter;
  const rows = [];
  let row = [];
  let field = '';
  let inQ = false;
  let quoted = false;
  let i = 0;
  const n = s.length;
  const endRow = () => { row.push(field); rows.push(row); row = []; field = ''; quoted = false; };
  while (i < n) {
    const ch = s[i];
    if (inQ) {
      if (ch === quote) {
        if (s[i + 1] === quote) { field += quote; i += 2; continue; }
        inQ = false; i++; continue;
      }
      field += ch; i++; continue;
    }
    if (ch === quote && (field === '' && !quoted)) { inQ = true; quoted = true; i++; continue; }
    if (ch === d) { row.push(field); field = ''; quoted = false; i++; continue; }
    if (ch === '\r' || ch === '\n') {
      endRow();
      if (rows.length >= maxRows) return rows;
      i += ch === '\r' && s[i + 1] === '\n' ? 2 : 1;
      continue;
    }
    field += ch; i++;
  }
  if (field !== '' || row.length || quoted) endRow();
  return rows.length > maxRows ? rows.slice(0, maxRows) : rows;
}

/** rows → CSV text; fields with delimiter, quote, CR/LF or outer spaces are quoted. */
export function writeCSV(rows, { delimiter = ',', eol = '\r\n', quoteAll = false } = {}) {
  const needs = (f) => f.includes(delimiter) || f.includes('"') || f.includes('\n') || f.includes('\r') || /^\s|\s$/.test(f);
  return (rows || []).map((r) => (r || []).map((v) => {
    const f = v == null ? '' : String(v);
    return quoteAll || needs(f) ? `"${f.replace(/"/g, '""')}"` : f;
  }).join(delimiter)).join(eol) + ((rows || []).length ? eol : '');
}

/** For an import dialog: decode, detect the delimiter, first 20 rows. */
export function sniff(bytes, { encoding = 'auto', delimiter = 'auto' } = {}) {
  const { text, encoding: enc } = decodeText(bytes, encoding);
  const d = !delimiter || delimiter === 'auto' ? detectDelimiter(text) : delimiter;
  return { text, encoding: enc, delimiter: d, preview: parseCSV(text, { delimiter: d, maxRows: 20 }) };
}
