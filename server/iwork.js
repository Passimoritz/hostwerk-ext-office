'use strict';
/*
 * Apple iWork (Pages, Keynote, Numbers – 2013 and newer: a zip with Index/*.iwa, older: a package with QuickLook/Preview.pdf).
 * Best effort, read only:
 *   previewOf(z)   → { pdf?: Buffer, image?: { type, data } } embedded preview
 *   iwaText(z)     → text: IWA chunks are Snappy-compressed protobuf; strings are taken from length-delimited fields
 *                    that look like human text (the body text of Pages/Keynote lives in one such field)
 */
/** Raw Snappy block (no framing, no CRC) → Buffer */
function snappy(buf) {
  let p = 0; let len = 0; let shift = 0;
  for (;;) { const b = buf[p++]; len |= (b & 0x7f) << shift; if (!(b & 0x80)) break; shift += 7; if (shift > 28) throw new Error('snappy: bad length'); }
  if (len > 128 * 1048576) throw new Error('snappy: too large');
  const out = Buffer.alloc(len); let o = 0;
  while (p < buf.length && o < len) {
    const tag = buf[p++];
    const t = tag & 3;
    if (t === 0) {
      let n = tag >> 2;
      if (n >= 60) { const k = n - 59; n = 0; for (let i = 0; i < k; i++) n |= buf[p + i] << (8 * i); p += k; }
      n += 1;
      buf.copy(out, o, p, p + n); o += n; p += n;
    } else {
      let n; let off;
      if (t === 1) { n = ((tag >> 2) & 7) + 4; off = ((tag >> 5) << 8) | buf[p++]; } else if (t === 2) { n = (tag >> 2) + 1; off = buf.readUInt16LE(p); p += 2; } else { n = (tag >> 2) + 1; off = buf.readUInt32LE(p); p += 4; }
      if (!off || off > o) throw new Error('snappy: bad offset');
      for (let i = 0; i < n && o < len; i++, o++) out[o] = out[o - off];
    }
  }
  return out.subarray(0, o);
}
/** IWA file: chunks [0x00, len (3 bytes LE), snappy data] → protobuf bytes */
function iwaDecode(buf) {
  const parts = []; let p = 0;
  while (p + 4 <= buf.length) {
    const type = buf[p]; const n = buf[p + 1] | (buf[p + 2] << 8) | (buf[p + 3] << 16); p += 4;
    if (type !== 0) break;
    parts.push(snappy(buf.subarray(p, p + n))); p += n;
  }
  return Buffer.concat(parts);
}
const looksLikeText = (s) => {
  if (s.length < 2 || s.includes('�')) return false;
  const letters = (s.match(/[\p{L}\p{N}\s.,;:!?'"()\-–—…€$%&/]/gu) || []).length;
  return letters / s.length > 0.92 && /\p{L}{2}/u.test(s) && !/^[a-z]+(\.[a-z]+)+$/i.test(s) && !/^[A-Z][A-Za-z]+Archive$/.test(s);
};
/** Strings from protobuf bytes (heuristic: wire type 2 fields with readable UTF-8) */
function protoStrings(buf, out = [], depth = 0) {
  let p = 0;
  const varint = () => { let v = 0; let s = 0; for (let i = 0; i < 10 && p < buf.length; i++) { const b = buf[p++]; v += (b & 0x7f) * 2 ** s; if (!(b & 0x80)) return v; s += 7; } return -1; };
  while (p < buf.length) {
    const key = varint(); if (key < 0) return out;
    const wt = key % 8;
    if (wt === 0) { if (varint() < 0) return out; } else if (wt === 1) p += 8; else if (wt === 5) p += 4;
    else if (wt === 2) {
      const n = varint(); if (n < 0 || p + n > buf.length) return out;
      const sub = buf.subarray(p, p + n); p += n;
      const s = sub.toString('utf8');
      if (looksLikeText(s)) out.push(s); else if (depth < 6 && n > 2) protoStrings(sub, out, depth + 1);
    } else return out;
  }
  return out;
}
function iwaText(z) {
  const names = [...z.keys()].filter((n) => /^Index\/.*\.iwa$/.test(n)).sort((a, b) => (/Document\.iwa$/.test(b) ? 1 : 0) - (/Document\.iwa$/.test(a) ? 1 : 0));
  const seen = new Set(); const out = [];
  for (const n of names) {
    let pb; try { pb = iwaDecode(z.get(n).read()); } catch { continue; }
    for (const s of protoStrings(pb)) { const v = s.trim(); if (v && !seen.has(v)) { seen.add(v); out.push(v); } }
  }
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}
function previewOf(z) {
  const res = {};
  for (const n of ['QuickLook/Preview.pdf', 'preview.pdf']) if (z.get(n)) { res.pdf = z.get(n).read(); break; }
  for (const [n, type] of [['preview.jpg', 'image/jpeg'], ['QuickLook/Thumbnail.jpg', 'image/jpeg'], ['preview-web.jpg', 'image/jpeg'], ['QuickLook/Thumbnail.png', 'image/png']]) if (z.get(n)) { res.image = { type, data: z.get(n).read() }; break; }
  return res;
}

module.exports = { snappy, iwaDecode, protoStrings, iwaText, previewOf, looksLikeText };
