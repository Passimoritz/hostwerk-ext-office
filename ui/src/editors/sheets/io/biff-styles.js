/*
 * Cell formatting of old Excel 97–2003 files (.xls / .xlt, BIFF8) – SheetJS CE reads values, formulas and number
 * formats but no fonts, borders or alignment. This small reader walks the Workbook stream itself: FONT, XF and
 * PALETTE records of the globals, and the XF index of every cell record (plus COLINFO / ROW defaults) per sheet.
 *
 *   readBiffStyles(bytes) → { sheets: [Map<'r,c', xfStyle>], cols: [Map<c, xfStyle>] } | null
 *   xfStyle = partial model style (font, size, b, i, u, st, color, fill, bd, h, va, wrap, rot, indent, shrink, unlocked, hiddenF)
 */
import { readCfb, isCfb, findStream } from '../../_shared/cfb.js';
import { BUILTIN_FORMATS } from '../format/numfmt.js';

const DEF_PAL = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
const BORDER = [null, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
const HALIGN = [null, 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const VALIGN = ['top', 'center', null, 'justify', 'distributed'];

function records(u, start, end, cb) {
  let p = start;
  while (p + 4 <= end) {
    const type = u[p] | (u[p + 1] << 8); const len = u[p + 2] | (u[p + 3] << 8);
    if (cb(type, p + 4, len) === false) return p;
    p += 4 + len;
    if (type === 0x000a) return p; // EOF
  }
  return p;
}
const u16 = (u, p) => u[p] | (u[p + 1] << 8);
const u32 = (u, p) => (u[p] | (u[p + 1] << 8) | (u[p + 2] << 16) | (u[p + 3] << 24)) >>> 0;

function shortString(u, p) {
  const cch = u[p]; const flags = u[p + 1]; let s = '';
  if (flags & 1) for (let i = 0; i < cch; i++) s += String.fromCharCode(u16(u, p + 2 + i * 2));
  else for (let i = 0; i < cch; i++) s += String.fromCharCode(u[p + 2 + i]);
  return s;
}

export function readBiffStyles(bytes) {
  const u0 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (!isCfb(u0)) return null;
  let cfb;
  try { cfb = readCfb(u0); } catch { return null; }
  const u = findStream(cfb, 'Workbook') || findStream(cfb, 'Book');
  if (!u) return null;
  if (u16(u, 4) !== 0x0600) return null; // BIFF8 only
  const fonts = []; const xfs = []; const formats = new Map(); const pal = DEF_PAL.slice(); const sheetsAt = [];
  records(u, 0, u.length, (t, p, len) => {
    if (t === 0x0031 && len >= 15) { // FONT
      const grbit = u16(u, p + 2);
      fonts.push({ size: u16(u, p) / 20, i: !!(grbit & 2), st: !!(grbit & 8), icv: u16(u, p + 4), b: u16(u, p + 6) >= 600, sss: u16(u, p + 8), uls: u[p + 10], name: shortString(u, p + 14) });
      if (fonts.length === 4) fonts.push(null); // index 4 does not exist
    } else if (t === 0x041e && len >= 5) { const cch = u16(u, p + 2); const hi = u[p + 4] & 1; let str = ''; for (let i = 0; i < cch; i++) str += String.fromCharCode(hi ? u16(u, p + 5 + i * 2) : u[p + 5 + i]); formats.set(u16(u, p), str); } // FORMAT
    else if (t === 0x00e0 && len >= 20) xfs.push(u.subarray(p, p + 20)); // XF
    else if (t === 0x0092) { const n = u16(u, p); for (let i = 0; i < n && i < 56; i++) pal[i] = [u[p + 2 + i * 4], u[p + 3 + i * 4], u[p + 4 + i * 4]].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase(); } // PALETTE (overrides 8..63)
    else if (t === 0x0085) sheetsAt.push({ pos: u32(u, p), type: u[p + 5] }); // BOUNDSHEET8
  });
  // icv 0..7: the fixed EGA colours, 8..63: the palette (PALETTE changes it)
  const col = (icv) => (icv < 8 ? `#${DEF_PAL[icv]}` : icv < 64 ? `#${pal[icv - 8]}` : null);
  const def = fonts[0] || { name: 'Arial', size: 10 };
  const cache = new Map();
  const styleOf = (ix) => {
    if (cache.has(ix)) return cache.get(ix);
    const x = xfs[ix]; if (!x) { cache.set(ix, null); return null; }
    const s = {};
    const f = fonts[u16(x, 0)] || null;
    if (f) {
      if (f.name && f.name !== def.name) s.font = f.name;
      if (f.size && f.size !== def.size) s.size = f.size;
      if (f.b) s.b = true; if (f.i) s.i = true; if (f.st) s.st = true;
      if (f.uls === 1 || f.uls === 0x21) s.u = true; else if (f.uls === 2 || f.uls === 0x22) s.u = 'double';
      if (f.sss === 1) s.vert = 'superscript'; else if (f.sss === 2) s.vert = 'subscript';
      const c = f.icv !== 0x7fff && f.icv !== 64 ? col(f.icv) : null; if (c && c !== '#000000') s.color = c;
    }
    // number format: the built-in ones in their English form (a German Excel stores 4–8 localized), own codes as written
    const ifmt = u16(x, 2);
    if (ifmt === 14 || ifmt === 22) s.nf = ifmt; else if (ifmt && ifmt < 164 && BUILTIN_FORMATS[ifmt]) s.nf = BUILTIN_FORMATS[ifmt]; else if (ifmt >= 164 && formats.has(ifmt)) s.nf = formats.get(ifmt);
    const prot = u16(x, 4);
    if (!(prot & 1)) s.unlocked = true;
    if (prot & 2) s.hiddenF = true;
    const al = x[6];
    const h = HALIGN[al & 7]; if (h) s.h = h;
    if (al & 8) s.wrap = true;
    const v = VALIGN[(al >> 4) & 7]; if (v) s.va = v;
    const rot = x[7]; if (rot) s.rot = rot;
    const ind = x[8]; if (ind & 15) s.indent = ind & 15; if (ind & 16) s.shrink = true;
    const b1 = u32(x, 10); const b2 = u32(x, 14); const b3 = u16(x, 18);
    const bd = {};
    const side = (k, dg, icv) => { if (BORDER[dg]) bd[k] = { s: BORDER[dg], c: col(icv) || '#000000' }; };
    side('l', b1 & 15, (b1 >>> 16) & 0x7f); side('r', (b1 >>> 4) & 15, (b1 >>> 23) & 0x7f); side('t', (b1 >>> 8) & 15, b2 & 0x7f); side('b', (b1 >>> 12) & 15, (b2 >>> 7) & 0x7f);
    const diag = (b1 >>> 30) & 3; const dgDiag = (b2 >>> 21) & 15;
    if (diag && BORDER[dgDiag]) bd.d = { s: BORDER[dgDiag], c: col((b2 >>> 14) & 0x7f) || '#000000', down: !!(diag & 1), up: !!(diag & 2) };
    if (Object.keys(bd).length) s.bd = bd;
    const fls = (b2 >>> 26) & 0x3f;
    if (fls) { const fg = col(b3 & 0x7f); const bg = col((b3 >>> 7) & 0x7f); const c = fls === 1 ? fg : fg || bg; if (c) s.fill = c; }
    cache.set(ix, s);
    return s;
  };
  const sheets = []; const cols = [];
  for (const sh of sheetsAt) {
    const cells = new Map(); const cdef = new Map();
    if (sh.type !== 0 || sh.pos >= u.length) { sheets.push(cells); cols.push(cdef); continue; }
    let depth = 0;
    records(u, sh.pos, u.length, (t, p, len) => {
      if (t === 0x0809) { depth++; return true; } // BOF (nested: charts)
      if (t === 0x000a) { depth--; return depth > 0 ? true : false; }
      if (depth > 1) return true;
      const put = (r, c, ix) => { const s = styleOf(ix); if (s && Object.keys(s).length) cells.set(`${r},${c}`, s); };
      switch (t) {
        case 0x00fd: case 0x0204: case 0x0203: case 0x027e: case 0x0201: case 0x0006: case 0x0205: put(u16(u, p), u16(u, p + 2), u16(u, p + 4)); break;
        case 0x00bd: { const r = u16(u, p); const c1 = u16(u, p + 2); const n = (len - 6) / 6; for (let i = 0; i < n; i++) put(r, c1 + i, u16(u, p + 4 + i * 6)); break; }
        case 0x00be: { const r = u16(u, p); const c1 = u16(u, p + 2); const n = (len - 6) / 2; for (let i = 0; i < n; i++) put(r, c1 + i, u16(u, p + 4 + i * 2)); break; }
        case 0x007d: { const c1 = u16(u, p); const c2 = Math.min(u16(u, p + 2), 255); const s = styleOf(u16(u, p + 6)); if (s && Object.keys(s).length && c2 - c1 < 256) for (let c = c1; c <= c2; c++) cdef.set(c, s); break; }
        default: break;
      }
      return true;
    });
    sheets.push(cells); cols.push(cdef);
  }
  return { sheets, cols, defaultFont: { font: def.name, size: def.size } };
}
