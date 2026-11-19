/*
 * Best-effort text extraction for formats we cannot render without the Office engine – pure, testable.
 *
 *   snappyUncompress(u8)        raw Snappy block
 *   iwaUnframe(u8)              Apple IWA framing (chunks: 0x00 + 3-byte length + Snappy block) → bytes
 *   protoStrings(u8)            UTF-8 strings of length-delimited protobuf fields (heuristic)
 *   iworkText(iwaBytesList)     readable text of a Pages / Keynote / Numbers document
 *   binaryStrings(u8)           UTF-16LE and ASCII/Latin-1 runs of a binary file (vsd, pub, one, mpp, mdb …)
 */

export function snappyUncompress(src) {
  let p = 0; let len = 0; let shift = 0;
  for (;;) { const b = src[p++]; len |= (b & 0x7f) << shift; if (!(b & 0x80)) break; shift += 7; if (shift > 28) throw new Error('bad snappy length'); }
  const out = new Uint8Array(len); let o = 0;
  while (p < src.length && o < len) {
    const tag = src[p++]; const type = tag & 3;
    if (type === 0) {
      let l = tag >> 2;
      if (l >= 60) { const n = l - 59; l = 0; for (let i = 0; i < n; i++) l |= src[p++] << (8 * i); }
      l += 1;
      out.set(src.subarray(p, p + l), o); p += l; o += l;
    } else {
      let l; let off;
      if (type === 1) { l = ((tag >> 2) & 7) + 4; off = ((tag >> 5) << 8) | src[p++]; } else if (type === 2) { l = (tag >> 2) + 1; off = src[p] | (src[p + 1] << 8); p += 2; } else { l = (tag >> 2) + 1; off = (src[p] | (src[p + 1] << 8) | (src[p + 2] << 16) | (src[p + 3] << 24)) >>> 0; p += 4; }
      if (!off || off > o) throw new Error('bad snappy offset');
      for (let i = 0; i < l; i++, o++) out[o] = out[o - off];
    }
  }
  return o === len ? out : out.subarray(0, o);
}

export function iwaUnframe(u8) {
  const parts = []; let p = 0; let total = 0;
  while (p + 4 <= u8.length) {
    const type = u8[p]; const len = u8[p + 1] | (u8[p + 2] << 8) | (u8[p + 3] << 16);
    p += 4;
    if (type !== 0 || p + len > u8.length) break;
    try { const d = snappyUncompress(u8.subarray(p, p + len)); parts.push(d); total += d.length; } catch { break; }
    p += len;
  }
  const out = new Uint8Array(total); let o = 0;
  for (const d of parts) { out.set(d, o); o += d.length; }
  return out;
}

const dec = typeof TextDecoder !== 'undefined' ? new TextDecoder('utf-8', { fatal: true }) : null;
const readable = (s) => {
  if (!s || /[\u0000-\u0008\u000e-\u001f�]/.test(s)) return false;
  const letters = (s.match(/[\p{L}\p{N}]/gu) || []).length;
  return letters >= Math.max(2, s.length * 0.5);
};

/** strings in protobuf length-delimited fields: tag byte with wire type 2, varint length, UTF-8 bytes */
export function protoStrings(u8, min = 2) {
  const out = [];
  for (let i = 0; i < u8.length - 2; i++) {
    const tag = u8[i];
    if ((tag & 7) !== 2 || tag < 8) continue;
    let p = i + 1; let len = 0; let shift = 0; let ok = true;
    for (let k = 0; k < 4; k++) { const b = u8[p++]; len |= (b & 0x7f) << shift; if (!(b & 0x80)) break; shift += 7; if (k === 3) ok = false; }
    if (!ok || len < min || len > 1 << 22 || p + len > u8.length) continue;
    let s;
    try { s = dec.decode(u8.subarray(p, p + len)); } catch { continue; }
    if (!readable(s)) continue;
    out.push(s);
    i = p + len - 1;
  }
  return out;
}

const looksTechnical = (s) => /^[\w.-]+$/.test(s) && !/\s/.test(s) && (/[A-Z][a-z]+[A-Z]/.test(s) || /[._]/.test(s) || /^[a-z]+[A-Z]/.test(s) || s.length < 4);
/** Pages/Keynote/Numbers: text from the .iwa parts (body text first, then smaller strings), de-duplicated */
export function iworkText(parts) {
  const seen = new Set(); const out = [];
  for (const bytes of parts) {
    let raw;
    try { raw = iwaUnframe(bytes); } catch { continue; }
    for (const s of protoStrings(raw, 2)) {
      const t = s.replace(/￼/g, '').trim();
      if (!t || seen.has(t) || looksTechnical(t)) continue;
      if (/^(Default|Body|Title|Heading|Caption|Footnote|Header|Footer|Bullet|Normal|None|Basic|Free Form)\b/.test(t) && t.length < 25) continue;
      seen.add(t); out.push(t);
    }
  }
  return out.join('\n');
}

/** text runs in a binary file (UTF-16LE first – Office binaries store text that way) */
export function binaryStrings(u8, { min = 5, max = 4000 } = {}) {
  const out = []; const seen = new Set();
  const add = (s) => { const t = s.replace(/\s+/g, ' ').trim(); if (t.length >= min && readable(t) && !seen.has(t)) { seen.add(t); out.push(t); } };
  let cur = '';
  for (let i = 0; i + 1 < u8.length && out.length < max; i += 2) {
    const c = u8[i] | (u8[i + 1] << 8);
    if ((c >= 0x20 && c < 0xd800 && c !== 0xfffe && c !== 0xffff) || c === 9 || c === 10 || c === 13) cur += String.fromCharCode(c);
    else { if (cur.length >= min) add(cur); cur = ''; }
  }
  if (cur.length >= min) add(cur);
  // odd offsets too (UTF-16 runs that start at an odd byte)
  cur = '';
  for (let i = 1; i + 1 < u8.length && out.length < max; i += 2) {
    const c = u8[i] | (u8[i + 1] << 8);
    if (c >= 0x20 && c < 0x2000) cur += String.fromCharCode(c); else { if (cur.length >= min) add(cur); cur = ''; }
  }
  const utf16 = out.length;
  cur = '';
  for (let i = 0; i < u8.length && out.length < max; i++) {
    const c = u8[i];
    if ((c >= 0x20 && c < 0x7f) || (c >= 0xc0 && c <= 0xff) || c === 9) cur += String.fromCharCode(c);
    else { if (cur.length >= min + 3 && /\s/.test(cur)) add(cur); cur = ''; }
  }
  return { strings: out, utf16 };
}
