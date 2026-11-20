/*
 * Codepage table for SheetJS (XLSX.set_cptable): old .xls (BIFF2–5) and .dbf files store text in a Windows/DOS/Mac
 * codepage. Decoding uses the browser's own TextDecoder (all WHATWG encodings incl. Shift-JIS, GBK, Big5, EUC-KR,
 * windows-125x, KOI8, Mac) – nothing to bundle; only the DOS codepages 437 and 850 (common in dBASE files) are tables.
 * Encoding (writing .xls/.dbf) works for every single-byte codepage; double-byte ones are built lazily on first use.
 */
const LABEL = {
  874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 866: 'ibm866', 1200: 'utf-16le', 1201: 'utf-16be', 65001: 'utf-8',
  10000: 'macintosh', 10007: 'x-mac-cyrillic', 20866: 'koi8-r', 21866: 'koi8-u', 20932: 'euc-jp', 54936: 'gb18030', 51932: 'euc-jp', 50220: 'iso-2022-jp',
  28592: 'iso-8859-2', 28593: 'iso-8859-3', 28594: 'iso-8859-4', 28595: 'iso-8859-5', 28596: 'iso-8859-6', 28597: 'iso-8859-7', 28598: 'iso-8859-8', 28603: 'iso-8859-13', 28605: 'iso-8859-15',
};
for (let i = 1250; i <= 1258; i++) LABEL[i] = `windows-${i}`;
const HIGH = {
  437: 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ',
  850: 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜø£Ø×ƒáíóúñÑªº¿®¬½¼¡«»░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐└┴┬├─┼ãÃ╚╔╩╦╠═╬¤ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀ÓßÔÒõÕµþÞÚÛÙýÝ¯´­±‗¾¶§÷¸°¨·¹³²■ ',
};
const decoders = new Map();
const single = new Map(); // cp → 256-char string
function table(cp) {
  let s = single.get(cp);
  if (s !== undefined) return s;
  s = null;
  if (HIGH[cp]) s = Array.from({ length: 128 }, (_, i) => String.fromCharCode(i)).join('') + HIGH[cp];
  else if (cp === 28591 || cp === 20127) s = Array.from({ length: 256 }, (_, i) => String.fromCharCode(i)).join('');
  else if (LABEL[cp] && ![932, 936, 949, 950, 1200, 1201, 65001, 20932, 51932, 54936, 50220].includes(cp)) {
    try { const d = new TextDecoder(LABEL[cp]); s = ''; for (let i = 0; i < 256; i++) s += d.decode(new Uint8Array([i])) || '�'; } catch { s = null; }
  }
  single.set(cp, s);
  return s;
}
function decode(cp, data) {
  const bytes = typeof data === 'string' ? Uint8Array.from(data, (c) => c.charCodeAt(0) & 255) : data instanceof Uint8Array ? data : Uint8Array.from(data || []);
  const tb = table(cp);
  if (tb) { let o = ''; for (let i = 0; i < bytes.length; i++) o += tb[bytes[i]]; return o; }
  const label = LABEL[cp] || 'windows-1252';
  let d = decoders.get(label);
  if (!d) { try { d = new TextDecoder(label); } catch { d = new TextDecoder('windows-1252'); } decoders.set(label, d); }
  return d.decode(bytes);
}
const reverse = new Map();
function encode(cp, str, mode) {
  const s = String(str ?? '');
  let out;
  if (cp === 65001) out = [...new TextEncoder().encode(s)];
  else if (cp === 1200 || cp === 1201) { out = []; for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); if (cp === 1200) out.push(c & 255, c >> 8); else out.push(c >> 8, c & 255); } } else {
    let rev = reverse.get(cp);
    if (!rev) {
      rev = new Map();
      const tb = table(cp);
      if (tb) for (let i = 255; i >= 0; i--) rev.set(tb[i], [i]);
      else { // double-byte: decode every lead/trail pair once
        const d = new TextDecoder(LABEL[cp] || 'windows-1252');
        for (let i = 0; i < 128; i++) rev.set(String.fromCharCode(i), [i]);
        for (let a = 0x81; a <= 0xfe; a++) for (let b = 0x40; b <= 0xfe; b++) { const ch = d.decode(new Uint8Array([a, b])); if (ch.length === 1 && ch !== '�' && !rev.has(ch)) rev.set(ch, [a, b]); }
      }
      reverse.set(cp, rev);
    }
    out = [];
    for (const ch of s) { const b = rev.get(ch); if (b) out.push(...b); else out.push(63); }
  }
  if (mode === 'str') return out.map((b) => String.fromCharCode(b)).join('');
  return mode === 'arr' ? out : typeof Buffer !== 'undefined' && mode === 'buf' ? Buffer.from(out) : out;
}
export const cptable = { utils: { decode, encode, hascp: (cp) => !!(LABEL[cp] || HIGH[cp] || cp === 28591 || cp === 20127) } };
