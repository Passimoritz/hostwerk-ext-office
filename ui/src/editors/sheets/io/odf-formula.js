/*
 * OpenFormula (ODF 1.2 part 2) ↔ our formula storage for ods/ots/fods.
 *
 * SheetJS CE converts formulas for OpenDocument with a few regexes: it writes `of:=Data![.D7]*2` for a reference to
 * another sheet (Excel and LibreOffice give #REF!/#NAME?) and reads `of:=[Data.D7]` back as `[Data.D7]`; it also
 * swaps ',' and ';' inside string literals. So we let SheetJS write a placeholder formula per cell and put our own
 * OpenFormula text into the XML afterwards, and when reading we take the raw `table:formula` attributes from
 * content.xml ourselves.
 */
import * as XLSX from 'xlsx';
import { tokenize } from '../formula/parser.js';
import { colName, colIndex, quoteSheet, MAXR, MAXC } from '../formula/refs.js';
import { crc32 } from '../../_shared/zipwrite.js';

/** functions LibreOffice stores with the COM.MICROSOFT. prefix (newer than ODF 1.2) */
const COM_MS = new Set(['IFS', 'MAXIFS', 'MINIFS', 'SWITCH', 'CONCAT', 'TEXTJOIN', 'XLOOKUP', 'XMATCH', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'LET', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT', 'VSTACK', 'HSTACK', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND']);

const odfSheet = (name) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${String(name).replace(/'/g, "''")}'`);
const part = (col, row, ac, ar) => `${col == null ? '' : `${ac ? '$' : ''}${colName(col)}`}${row == null ? '' : `${ar ? '$' : ''}${row + 1}`}`;

/** one reference token of our tokenizer → `[Sheet.A1:.B2]` */
function refToOdf(ref) {
  const sh = ref.sheet != null ? `$${odfSheet(ref.sheet)}` : '';
  if (ref.kind === 'cols') return `[${sh}.${part(ref.c1, null, ref.ac1)}:.${part(ref.c2, null, ref.ac2)}]`;
  if (ref.kind === 'rows') return `[${sh}.${part(null, ref.r1, false, ref.ar1)}:.${part(null, ref.r2, false, ref.ar2)}]`;
  const a = `${sh}.${part(ref.c1, ref.r1, ref.ac1, ref.ar1)}`;
  if (ref.kind === 'range') return `[${a}:.${part(ref.c2, ref.r2, ref.ac2, ref.ar2)}]`;
  return `[${a}]`;
}

/** storage formula (English, no '=') → `of:=…` */
export function toOpenFormula(formula) {
  let brace = 0;
  const out = tokenize(String(formula || '')).map((t) => {
    switch (t.t) {
      case 'ref': return refToOdf(t.ref);
      case 'func': return COM_MS.has(t.v) || (/\./.test(t.v) && !/^(COM|ORG)\./.test(t.v)) ? `COM.MICROSOFT.${t.v}` : t.v;
      case 'name': return t.v; // sheet-scoped names have no OpenFormula form: the name alone
      case 'bool': return t.v ? 'TRUE()' : 'FALSE()';
      case 'sep': return ';';
      case 'lb': brace++; return '{';
      case 'rb': brace = Math.max(0, brace - 1); return '}';
      case 'col': return ';';
      case 'row': return '|';
      case 'at': return '';
      default: return t.text;
    }
  }).join('');
  return `of:=${out}`;
}

/* ---------------- reading ---------------- */

const unq = (s) => s.replace(/''/g, "'");
/** `$'My Sheet'.$A$1` / `.A1` / `.A` / `.1` → { sheet, col, row, ac, ar } */
function parsePart(p) {
  const m = /^(?:\$?(?:'((?:[^']|'')*)'|([^.']*)))?\.(\$?)([A-Za-z]{0,3})(\$?)(\d*)$/.exec(p);
  if (!m) return null;
  const sheet = m[1] != null ? unq(m[1]) : m[2] ? m[2] : null;
  const col = m[4] ? colIndex(m[4]) : null; const row = m[6] ? Number(m[6]) - 1 : null;
  if (col == null && row == null) return null;
  if ((col != null && col >= MAXC) || (row != null && (row < 0 || row >= MAXR))) return null;
  return { sheet, col, row, ac: !!m[3], ar: !!m[5] };
}
const xl = (p) => `${p.col == null ? '' : `${p.ac ? '$' : ''}${colName(p.col)}`}${p.row == null ? '' : `${p.ar ? '$' : ''}${p.row + 1}`}`;

/** inner text of `[…]` → Excel-style reference text, or null */
function odfRefToExcel(inner) {
  if (/#REF!/i.test(inner)) return '#REF!';
  // split at ':' outside quotes
  const parts = []; let cur = ''; let q = false;
  for (const ch of inner) { if (ch === "'") q = !q; if (ch === ':' && !q) { parts.push(cur); cur = ''; } else cur += ch; }
  parts.push(cur);
  if (parts.length > 2) return null;
  const a = parsePart(parts[0].trim());
  if (!a) return null;
  const pre = a.sheet != null ? `${quoteSheet(a.sheet)}!` : '';
  if (parts.length === 1) return a.col != null && a.row != null ? pre + xl(a) : null;
  const b = parsePart(parts[1].trim());
  if (!b) return null;
  return `${pre}${xl(a)}:${xl(b)}`;
}

/** `of:=[Data.D7]*2` → `Data!D7*2` (storage form); null for formulas in another dialect we can't read */
export function fromOpenFormula(f) {
  let s = String(f || '');
  const ns = /^([a-z]+):/.exec(s);
  if (ns) {
    if (ns[1] === 'msoxl') return s.slice(6).replace(/^=/, ''); // Excel syntax as is
    if (ns[1] !== 'of') return null;
    s = s.slice(3);
  }
  s = s.replace(/^==?/, '');
  let out = ''; let brace = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') { let j = i + 1; while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; } out += s.slice(i, j + 1); i = j; continue; }
    if (ch === "'") { let j = i + 1; while (j < s.length) { if (s[j] === "'") { if (s[j + 1] === "'") { j += 2; continue; } break; } j++; } out += s.slice(i, j + 1); i = j; continue; }
    if (ch === '[') {
      let j = i + 1; let q = false;
      while (j < s.length && (q || s[j] !== ']')) { if (s[j] === "'") q = !q; j++; }
      const r = odfRefToExcel(s.slice(i + 1, j));
      if (r == null) return null;
      out += r; i = j; continue;
    }
    if (ch === '#') { const e = /^#[A-Za-z0-9/]+[!?]?/.exec(s.slice(i)); if (e) { out += e[0]; i += e[0].length - 1; continue; } }
    if (ch === '{') { brace++; out += ch; continue; }
    if (ch === '}') { brace = Math.max(0, brace - 1); out += ch; continue; }
    if (ch === ';') { out += ','; continue; }
    if (ch === '|' && brace) { out += ';'; continue; }
    if (ch === '~') { out += ','; continue; } // union
    if (ch === '!') { out += ' '; continue; } // intersection
    out += ch;
  }
  return out.replace(/\b(?:COM\.MICROSOFT|ORG\.OPENOFFICE|ORG\.LIBREOFFICE)\.([A-Z][A-Z0-9.]*\s*\()/gi, '$1')
    .replace(/\b(TRUE|FALSE)\(\s*\)/gi, (m, b) => b.toUpperCase());
}

/* ---------------- content.xml access ---------------- */

const decodeXml = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|quot|amp|lt|gt|apos);/gi, (m, e) => (e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : { quot: '"', amp: '&', lt: '<', gt: '>', apos: "'" }[e.toLowerCase()]));
const encodeXml = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const attr = (a, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(a); return m ? decodeXml(m[1]) : null; };

/** the content XML of an ods/ots (zip) or fods (flat XML) file, or null */
export function odfContentXml(bytes) {
  const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (u[0] === 0x50 && u[1] === 0x4b) {
    try {
      const z = XLSX.CFB.read(u, { type: 'array' });
      const e = XLSX.CFB.find(z, 'content.xml');
      return e && e.content ? new TextDecoder().decode(e.content instanceof Uint8Array ? e.content : new Uint8Array(e.content)) : null;
    } catch { return null; }
  }
  const t = new TextDecoder().decode(u.subarray(0, Math.min(u.length, 4096)));
  return /<office:document[\s>]/.test(t) ? new TextDecoder().decode(u) : null;
}

/** raw formulas per sheet: [Map('r,c' → 'of:=…')] in table order */
export function odfRawFormulas(xml) {
  const sheets = [];
  if (!xml) return sheets;
  const re = /<(\/?)table:(table|table-row|table-cell|covered-table-cell)(?=[\s/>])([^>]*?)(\/?)>/g;
  let m; let cur = null; let r = 0; let c = 0; let rowRep = 1; let inCell = 0; let depth = 0;
  while ((m = re.exec(xml))) {
    const [, close, tag, a, self] = m;
    if (inCell) { // nested content of a cell (sub-tables are not supported): wait for the cell's end
      if (tag === 'table-cell' || tag === 'covered-table-cell') { if (close) inCell--; else if (!self) inCell++; }
      continue;
    }
    if (tag === 'table') {
      if (close) { depth--; if (!depth) cur = null; continue; }
      if (self) continue;
      depth++;
      if (depth === 1) { cur = new Map(); sheets.push(cur); r = 0; }
      continue;
    }
    if (!cur || depth !== 1) continue;
    if (tag === 'table-row') {
      if (close) { r += rowRep; continue; }
      rowRep = Math.max(1, Number(attr(a, 'table:number-rows-repeated')) || 1); c = 0;
      if (self) r += rowRep;
      continue;
    }
    if (close) continue;
    const rep = Math.max(1, Number(attr(a, 'table:number-columns-repeated')) || 1);
    const f = tag === 'table-cell' ? attr(a, 'table:formula') : null;
    if (f && rep * rowRep <= 20000) for (let rr = 0; rr < rowRep; rr++) for (let k = 0; k < rep; k++) cur.set(`${r + rr},${c + k}`, f);
    c += rep;
    if (!self) inCell = 1;
  }
  return sheets;
}

/* ---------------- writing ---------------- */

export const PLACEHOLDER = (n) => `HWFORMULA_${n}_`;
const PH_RE = /of:=HWFORMULA_(\d+)_/g;

/** put the real OpenFormula texts (list index = placeholder number) into the XML SheetJS wrote */
export function fillFormulas(xml, formulas) {
  return xml.replace(PH_RE, (m, n) => encodeXml(toOpenFormula(formulas[Number(n)])));
}

/** minimal zip writer: `mimetype` first and stored (ODF 1.2 part 3, 3.3), the rest deflated */
function zipOdf(files) {
  const enc = new TextEncoder();
  const chunks = []; const central = []; let offset = 0;
  for (const f of files) {
    const name = enc.encode(f.name); const data = f.bytes;
    const store = f.name === 'mimetype';
    const comp = store ? data : XLSX.CFB.utils._deflateRaw(data);
    const body = comp instanceof Uint8Array ? comp : new Uint8Array(comp);
    const crc = crc32(data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, store ? 0 : 8, true);
    lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true); lh.setUint32(14, crc, true); lh.setUint32(18, body.length, true);
    lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
    chunks.push(new Uint8Array(lh.buffer), name, body);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, store ? 0 : 8, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true); ch.setUint32(16, crc, true);
    ch.setUint32(20, body.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true);
    ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), name);
    offset += 30 + name.length + body.length;
  }
  const cdSize = central.reduce((n, b) => n + b.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  const all = [...chunks, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((n, b) => n + b.length, 0));
  let p = 0; for (const b of all) { out.set(b, p); p += b.length; }
  return out;
}

/** SheetJS ods/fods output + formula list → final bytes */
export function patchOdfOutput(bytes, formulas, flat) {
  const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (flat) return new TextEncoder().encode(fillFormulas(new TextDecoder().decode(u), formulas));
  const z = XLSX.CFB.read(u, { type: 'array' });
  const root = z.FullPaths[0];
  const files = [];
  z.FullPaths.forEach((p, i) => {
    const fi = z.FileIndex[i];
    if (fi.type !== 2 || !fi.content || p.endsWith('\u0001Sh33tJ5')) return;
    let data = fi.content instanceof Uint8Array ? fi.content : new Uint8Array(fi.content);
    const name = p.slice(root.length);
    if (name === 'content.xml') data = new TextEncoder().encode(fillFormulas(new TextDecoder().decode(data), formulas));
    files.push({ name, bytes: data });
  });
  files.sort((a, b) => (a.name === 'mimetype' ? -1 : b.name === 'mimetype' ? 1 : 0));
  return zipOdf(files);
}
