/*
 * Minimal TrueType/OpenType support for Unicode text in PDF annotations – pure, no fontkit (≈ 700 KB).
 *
 *   parseFont(bytes)            → font { unitsPerEm, ascent, descent, capHeight, bbox, italicAngle, name, cff, gid(cp), advance(gid) }
 *   subsetFont(font, gids)      → TrueType bytes with only these glyphs (glyph ids unchanged – "blanking" subset,
 *                                  composite glyphs keep their parts); CFF fonts (OTTO) are returned unchanged
 *   UnicodeFont                 → collects the text drawn with a font, gives hex strings + widths, embeds a
 *                                  Type0/CIDFontType2 (Identity-H) font with ToUnicode into a pdf-lib document
 *   canEncodeWinAnsi(text)      → true when Helvetica (WinAnsi) can show the text
 */
import { PDFName, PDFString, PDFHexString } from 'pdf-lib';

const u16 = (d, o) => (d[o] << 8) | d[o + 1];
const i16 = (d, o) => { const v = u16(d, o); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (d, o) => ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0;
const tag = (d, o) => String.fromCharCode(d[o], d[o + 1], d[o + 2], d[o + 3]);

/* WinAnsi = Latin-1 printable + these from 0x80–0x9f */
const WIN_EXTRA = new Set([0x20ac, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x017d, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x017e, 0x0178]);
export function canEncodeWinAnsi(text) {
  for (const ch of String(text || '')) {
    const c = ch.codePointAt(0);
    if (c === 10 || c === 13 || c === 9) continue;
    if ((c >= 0x20 && c <= 0x7e) || (c >= 0xa0 && c <= 0xff) || WIN_EXTRA.has(c)) continue;
    return false;
  }
  return true;
}

export function parseFont(bytes) {
  const d = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let base = 0;
  if (tag(d, 0) === 'ttcf') base = u32(d, 12); // first font of a collection
  const sfnt = tag(d, base);
  const n = u16(d, base + 4);
  const tables = {};
  for (let i = 0; i < n; i++) {
    const o = base + 12 + i * 16;
    tables[tag(d, o)] = { offset: u32(d, o + 8), length: u32(d, o + 12) };
  }
  const T = (t) => tables[t];
  if (!T('head') || !T('hhea') || !T('hmtx') || !T('maxp') || !T('cmap')) throw new Error('Not a usable TrueType/OpenType font');
  const head = T('head').offset;
  const unitsPerEm = u16(d, head + 18);
  const bbox = [i16(d, head + 36), i16(d, head + 38), i16(d, head + 40), i16(d, head + 42)];
  const locFormat = i16(d, head + 50);
  const hhea = T('hhea').offset;
  const ascent = i16(d, hhea + 4); const descent = i16(d, hhea + 6);
  const numHM = u16(d, hhea + 34);
  const numGlyphs = u16(d, T('maxp').offset + 4);
  const hmtx = T('hmtx').offset;
  let capHeight = ascent * 0.7;
  const os2 = T('OS/2');
  if (os2 && u16(d, os2.offset) >= 2 && os2.length >= 90) capHeight = i16(d, os2.offset + 88);
  let italicAngle = 0;
  if (T('post')) italicAngle = i16(d, T('post').offset + 4);
  // cmap: prefer format 12 (full Unicode), else format 4 (BMP)
  const cm = T('cmap').offset;
  const subs = [];
  for (let i = 0, k = u16(d, cm + 2); i < k; i++) {
    const p = u16(d, cm + 4 + i * 8); const e = u16(d, cm + 6 + i * 8); const off = cm + u32(d, cm + 8 + i * 8);
    subs.push({ p, e, off, fmt: u16(d, off) });
  }
  const pick = subs.find((s) => s.fmt === 12 && (s.p === 3 || s.p === 0)) || subs.find((s) => s.fmt === 4 && ((s.p === 3 && (s.e === 1 || s.e === 10)) || s.p === 0)) || subs.find((s) => s.fmt === 4);
  if (!pick) throw new Error('The font has no Unicode character map');
  const map = new Map();
  if (pick.fmt === 12) {
    const groups = u32(d, pick.off + 12);
    for (let g = 0; g < groups; g++) {
      const o = pick.off + 16 + g * 12;
      const s = u32(d, o); const e = u32(d, o + 4); const gid = u32(d, o + 8);
      for (let c = s; c <= e && c - s < 0x10000; c++) map.set(c, gid + c - s);
    }
  } else {
    const o = pick.off; const segX2 = u16(d, o + 6); const seg = segX2 / 2;
    const ends = o + 14; const starts = ends + segX2 + 2; const deltas = starts + segX2; const ranges = deltas + segX2;
    for (let s = 0; s < seg; s++) {
      const end = u16(d, ends + s * 2); const start = u16(d, starts + s * 2); const delta = i16(d, deltas + s * 2); const ro = u16(d, ranges + s * 2);
      for (let c = start; c <= end && c !== 0xffff; c++) {
        let g;
        if (!ro) g = (c + delta) & 0xffff;
        else { const p = ranges + s * 2 + ro + (c - start) * 2; g = u16(d, p); if (g) g = (g + delta) & 0xffff; }
        if (g) map.set(c, g);
      }
    }
  }
  let name = 'Font';
  const nm = T('name');
  if (nm) {
    const o = nm.offset; const cnt = u16(d, o + 2); const so = o + u16(d, o + 4);
    for (let i = 0; i < cnt; i++) {
      const r = o + 6 + i * 12;
      if (u16(d, r + 6) !== 6) continue; // PostScript name
      const pid = u16(d, r); const len = u16(d, r + 8); const off = so + u16(d, r + 10);
      let s = '';
      if (pid === 3 || pid === 0) for (let k = 0; k + 1 < len; k += 2) s += String.fromCharCode(u16(d, off + k));
      else for (let k = 0; k < len; k++) s += String.fromCharCode(d[off + k]);
      if (s) { name = s; break; }
    }
  }
  name = name.replace(/[^A-Za-z0-9_-]/g, '') || 'Font';
  const advance = (g) => u16(d, hmtx + 4 * Math.min(g, numHM - 1));
  return {
    bytes: d, base, sfnt, tables, unitsPerEm, bbox, ascent, descent, capHeight, italicAngle, numGlyphs, locFormat, name,
    cff: !!T('CFF ') || sfnt === 'OTTO', map, gid: (cp) => map.get(cp) || 0, advance,
  };
}

/** glyph data range of glyph g */
function glyphRange(f, g) {
  const loca = f.tables.loca.offset; const d = f.bytes;
  const a = f.locFormat ? u32(d, loca + g * 4) : u16(d, loca + g * 2) * 2;
  const b = f.locFormat ? u32(d, loca + g * 4 + 4) : u16(d, loca + g * 2 + 2) * 2;
  return [f.tables.glyf.offset + a, f.tables.glyf.offset + b];
}

/** add the parts of composite glyphs */
function closeGlyphs(f, gids) {
  const set = new Set([0, ...gids]);
  const todo = [...set];
  const d = f.bytes;
  while (todo.length) {
    const g = todo.pop();
    if (g >= f.numGlyphs) continue;
    const [a, b] = glyphRange(f, g);
    if (b - a < 10 || i16(d, a) >= 0) continue;
    let o = a + 10;
    for (;;) {
      const flags = u16(d, o); const part = u16(d, o + 2);
      if (!set.has(part)) { set.add(part); todo.push(part); }
      o += 4 + (flags & 1 ? 4 : 2);
      if (flags & 8) o += 2; else if (flags & 0x40) o += 4; else if (flags & 0x80) o += 8;
      if (!(flags & 0x20)) break;
    }
  }
  return set;
}

function checksum(b) {
  let s = 0;
  for (let i = 0; i < b.length; i += 4) s = (s + ((b[i] << 24) | ((b[i + 1] || 0) << 16) | ((b[i + 2] || 0) << 8) | (b[i + 3] || 0))) >>> 0;
  return s;
}
function writeSfnt(tables) {
  const tags = Object.keys(tables).sort();
  const n = tags.length;
  let es = 0; while ((1 << (es + 1)) <= n) es++;
  const sr = (1 << es) * 16;
  let size = 12 + n * 16;
  for (const t of tags) size += (tables[t].length + 3) & ~3;
  const out = new Uint8Array(size);
  const w16 = (o, v) => { out[o] = (v >>> 8) & 255; out[o + 1] = v & 255; };
  const w32 = (o, v) => { out[o] = v >>> 24; out[o + 1] = (v >>> 16) & 255; out[o + 2] = (v >>> 8) & 255; out[o + 3] = v & 255; };
  w32(0, 0x00010000); w16(4, n); w16(6, sr); w16(8, es); w16(10, n * 16 - sr);
  let off = 12 + n * 16;
  tags.forEach((t, i) => {
    const b = tables[t]; const r = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[r + k] = t.charCodeAt(k);
    w32(r + 4, checksum(b)); w32(r + 8, off); w32(r + 12, b.length);
    out.set(b, off); off += (b.length + 3) & ~3;
  });
  return out;
}

/** TrueType subset that keeps glyph ids (unused glyphs become empty) – what a CIDFontType2 with Identity needs */
/** a glyph without its TrueType instructions (hinting) – much smaller, renders the same at print resolutions */
function unhint(g) {
  if (g.length < 12) return g;
  const nc = i16(g, 0);
  if (nc >= 0) {
    const il = u16(g, 10 + nc * 2);
    const at = 12 + nc * 2;
    const out = new Uint8Array(g.length - il);
    out.set(g.subarray(0, at)); out[at - 2] = 0; out[at - 1] = 0;
    out.set(g.subarray(at + il), at);
    return out;
  }
  let o = 10; let any = 0; const at = [];
  for (;;) {
    at.push(o); const flags = u16(g, o); any |= flags;
    o += 4 + (flags & 1 ? 4 : 2);
    if (flags & 8) o += 2; else if (flags & 0x40) o += 4; else if (flags & 0x80) o += 8;
    if (!(flags & 0x20)) break;
  }
  if (!(any & 0x100)) return g;
  const out = g.slice(0, o);
  for (const p of at) { const f2 = u16(out, p) & ~0x100; out[p] = f2 >> 8; out[p + 1] = f2 & 255; }
  return out;
}

export function subsetFont(f, gids, { keepCmap = false, stripHints = false } = {}) {
  if (f.cff || !f.tables.glyf || !f.tables.loca) return f.bytes;
  const keep = closeGlyphs(f, gids);
  const d = f.bytes;
  const parts = []; let total = 0;
  const loca = new Uint8Array((f.numGlyphs + 1) * 4);
  const wl = (i, v) => { loca[i * 4] = v >>> 24; loca[i * 4 + 1] = (v >>> 16) & 255; loca[i * 4 + 2] = (v >>> 8) & 255; loca[i * 4 + 3] = v & 255; };
  for (let g = 0; g < f.numGlyphs; g++) {
    wl(g, total);
    if (!keep.has(g)) continue;
    const [a, b] = glyphRange(f, g);
    if (b <= a) continue;
    const src = stripHints ? unhint(d.subarray(a, b)) : d.subarray(a, b);
    const len = (src.length + 3) & ~3;
    const part = new Uint8Array(len); part.set(src);
    parts.push(part); total += len;
  }
  wl(f.numGlyphs, total);
  const glyf = new Uint8Array(total); let o = 0;
  for (const p of parts) { glyf.set(p, o); o += p.length; }
  const tab = (t) => (f.tables[t] ? d.slice(f.tables[t].offset, f.tables[t].offset + f.tables[t].length) : null);
  const head = tab('head');
  head[50] = 0; head[51] = 1; // long loca
  head[8] = 0; head[9] = 0; head[10] = 0; head[11] = 0; // checkSumAdjustment
  const tables = { head, hhea: tab('hhea'), maxp: tab('maxp'), hmtx: tab('hmtx'), loca, glyf };
  for (const t of [...(stripHints ? [] : ['cvt ', 'fpgm', 'prep']), 'OS/2', 'post', 'name', ...(keepCmap ? ['cmap'] : [])]) { const b = tab(t); if (b) tables[t] = b; }
  if (tables.post && tables.post.length > 32) { tables.post = tables.post.slice(0, 32); tables.post[0] = 0; tables.post[1] = 3; tables.post[2] = 0; tables.post[3] = 0; } // format 3: no glyph names
  const out = writeSfnt(tables);
  // checkSumAdjustment over the whole file
  const adj = (0xb1b0afba - checksum(out)) >>> 0;
  const hOff = (() => { const n = (out[4] << 8) | out[5]; for (let i = 0; i < n; i++) { const r = 12 + i * 16; if (tag(out, r) === 'head') return u32(out, r + 8); } return -1; })();
  if (hOff >= 0) { out[hOff + 8] = adj >>> 24; out[hOff + 9] = (adj >>> 16) & 255; out[hOff + 10] = (adj >>> 8) & 255; out[hOff + 11] = adj & 255; }
  return out;
}

const hex4 = (n) => n.toString(16).padStart(4, '0');
function utf16hex(cp) {
  if (cp < 0x10000) return hex4(cp);
  const v = cp - 0x10000;
  return hex4(0xd800 + (v >> 10)) + hex4(0xdc00 + (v & 0x3ff));
}

/**
 * One embedded Unicode font. Use it like a pdf-lib PDFFont in our operators (encodeText, widthOfTextAtSize, ref);
 * call embed() once after all text was drawn (it writes the subset with exactly the glyphs used).
 */
export class UnicodeFont {
  constructor(doc, fonts) {
    this.doc = doc;
    this.fonts = fonts; // [parsed font] – the first that has a glyph wins (per character, but one font per run)
    this.font = fonts[0];
    this.used = new Map(); // gid → codepoint string
    this.ref = doc.context.nextRef();
    this.embedded = false;
  }
  covers(text) { for (const ch of String(text)) { const c = ch.codePointAt(0); if (c > 31 && !this.font.gid(c)) return false; } return true; }
  gids(text) {
    const out = [];
    for (const ch of String(text)) {
      const c = ch.codePointAt(0);
      if (c < 32) continue;
      const g = this.font.gid(c);
      out.push(g);
      if (g && !this.used.has(g)) this.used.set(g, ch);
    }
    return out;
  }
  encodeText(text) { return PDFHexString.of(this.gids(text).map(hex4).join('')); }
  widthOfTextAtSize(text, size) {
    let w = 0;
    for (const ch of String(text)) { const c = ch.codePointAt(0); if (c >= 32) w += this.font.advance(this.font.gid(c)); }
    return (w / this.font.unitsPerEm) * size;
  }
  heightAtSize(size) { return ((this.font.ascent - this.font.descent) / this.font.unitsPerEm) * size; }
  embed() {
    if (this.embedded) return this.ref;
    this.embedded = true;
    const ctx = this.doc.context; const f = this.font; const k = 1000 / f.unitsPerEm;
    const gids = [...this.used.keys()].sort((a, b) => a - b);
    const data = subsetFont(f, gids);
    const tagName = `${String.fromCharCode(...[...Array(6)].map(() => 65 + Math.floor(Math.random() * 26)))}+${f.name}`;
    const file = ctx.flateStream(data, f.cff ? { Subtype: 'OpenType' } : { Length1: data.length });
    const fileRef = ctx.register(file);
    const fd = ctx.obj({
      Type: 'FontDescriptor', FontName: tagName, Flags: 32, FontBBox: f.bbox.map((v) => Math.round(v * k)), ItalicAngle: f.italicAngle,
      Ascent: Math.round(f.ascent * k), Descent: Math.round(f.descent * k), CapHeight: Math.round(f.capHeight * k), StemV: 80,
    });
    fd.set(PDFName.of(f.cff ? 'FontFile3' : 'FontFile2'), fileRef);
    const W = [];
    for (const g of gids) W.push(g, [Math.round(f.advance(g) * k)]);
    const cid = ctx.obj({
      Type: 'Font', Subtype: f.cff ? 'CIDFontType0' : 'CIDFontType2', BaseFont: tagName, FontDescriptor: ctx.register(fd), W, DW: 1000,
    });
    cid.set(PDFName.of('CIDSystemInfo'), ctx.obj({ Registry: PDFString.of('Adobe'), Ordering: PDFString.of('Identity'), Supplement: 0 }));
    if (!f.cff) cid.set(PDFName.of('CIDToGIDMap'), PDFName.of('Identity'));
    // ToUnicode (copy/search in every PDF program)
    const lines = gids.map((g) => `<${hex4(g)}> <${utf16hex(this.used.get(g).codePointAt(0))}>`);
    let body = '';
    for (let i = 0; i < lines.length; i += 100) { const part = lines.slice(i, i + 100); body += `${part.length} beginbfchar\n${part.join('\n')}\nendbfchar\n`; }
    const cmap = `/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n${body}endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`;
    const tu = ctx.register(ctx.flateStream(new TextEncoder().encode(cmap)));
    const type0 = ctx.obj({ Type: 'Font', Subtype: 'Type0', BaseFont: tagName, Encoding: 'Identity-H', DescendantFonts: [ctx.register(cid)], ToUnicode: tu });
    ctx.assign(this.ref, type0);
    return this.ref;
  }
}

async function inflate(bytes) {
  const ds = new DecompressionStream('deflate');
  const out = new Response(new Blob([bytes]).stream().pipeThrough(ds));
  return new Uint8Array(await out.arrayBuffer());
}
/** WOFF 1.0 → TrueType/OpenType bytes (WOFF2 is not needed: our bundled font is WOFF 1) */
export async function woffToSfnt(bytes) {
  const d = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (tag(d, 0) !== 'wOFF') return d;
  const n = u16(d, 12);
  const tables = {};
  for (let i = 0; i < n; i++) {
    const o = 44 + i * 20;
    const off = u32(d, o + 4); const comp = u32(d, o + 8); const orig = u32(d, o + 12);
    const raw = d.subarray(off, off + comp);
    tables[tag(d, o)] = comp < orig ? await inflate(raw) : raw.slice();
  }
  return writeSfnt(tables);
}
/** "data:…;base64,…" (esbuild dataurl loader) → bytes */
export function dataUrlBytes(u) {
  const s = atob(String(u).split(',')[1] || '');
  const a = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) a[i] = s.charCodeAt(i);
  return a;
}
