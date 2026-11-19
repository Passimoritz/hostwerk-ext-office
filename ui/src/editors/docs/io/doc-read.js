/*
 * Word 97–2003 (.doc / .dot) reader without any converter: CFB → FIB → piece table (text) → CHPX/PAPX formatting
 * (bold, italic, underline, strike, super/sub, size, font, colour, highlight, alignment, indents, spacing, line
 * height, page breaks, headings by style), tables (cell/row marks), lists (PlfLst/PlfLfo number formats), fields
 * (result text, HYPERLINK → link), footnotes, comments (text + author), tracked changes (insert/delete marks with
 * author), inline pictures (PNG/JPEG blips from the Data stream), page size/margins/orientation/columns (last
 * section), header/footer text and title/author from SummaryInformation. Macros are listed via vba.js, never run.
 * Not supported: Word 6/95 and older files (older FIB – error code DOC_TOO_OLD, the engine opens them), password-protected files, floating drawings (noted as fidelity).
 */
import { readCfb } from '../../_shared/cfb.js';
import { defaultModel, sizeName, twipToMm, normalizeInline, fixDoc, dataUrl, imageSize } from './pm.js';
import { groupBlocks } from './docx-read.js';

const ICO = [null, '#000000', '#0000FF', '#00FFFF', '#00FF00', '#FF00FF', '#FF0000', '#FFFF00', '#FFFFFF', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#C0C0C0'];
const NFC = { 0: { kind: 'ordered' }, 1: { kind: 'ordered', listStyle: 'upper-roman' }, 2: { kind: 'ordered', listStyle: 'lower-roman' }, 3: { kind: 'ordered', listStyle: 'upper-alpha' }, 4: { kind: 'ordered', listStyle: 'lower-alpha' }, 23: { kind: 'bullet' } };

/** split a grpprl into [sprm, operandBytes] */
function sprms(u, start, end) {
  const out = [];
  let p = start;
  while (p + 2 <= end) {
    const sprm = u[p] | (u[p + 1] << 8);
    p += 2;
    const spra = sprm >> 13;
    let len; let dataStart = p;
    if (spra === 0 || spra === 1) len = 1; else if (spra === 2 || spra === 4 || spra === 5) len = 2; else if (spra === 3) len = 4; else if (spra === 7) len = 3;
    else if (sprm === 0xd608) { len = (u[p] | (u[p + 1] << 8)) + 1; dataStart = p + 2; }
    else { len = u[p] + 1; dataStart = p + 1; }
    if (p + len > end) break;
    out.push([sprm, u.subarray(dataStart, p + len)]);
    p += len;
  }
  return out;
}
const i16 = (b) => { const v = b[0] | (b[1] << 8); return v > 32767 ? v - 65536 : v; };
const u16 = (b) => b[0] | (b[1] << 8);

/** Word 6/95 and older: the shell offers the Office engine (it converts the file before the editor sees it) */
export const DOC_TOO_OLD = 'DOC_TOO_OLD';
function tooOld() {
  return Object.assign(new Error('This Word 6/95 document is too old for the built-in reader. The Office engine opens it – set it up once with “Set up the Office engine …” and open the file again, or use the desktop app.'), { code: DOC_TOO_OLD, needsEngine: true });
}

export function readDoc(input, settings) {
  const cfb = readCfb(input);
  const wd = cfb.streams.get('WordDocument');
  if (!wd) throw new Error('This is not a Word document.');
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  const U16 = (o) => dv.getUint16(o, true); const U32 = (o) => dv.getUint32(o, true); const I32 = (o) => dv.getInt32(o, true);
  const wIdent = U16(0); const nFib = U16(2);
  // Word for Windows 2.0 (0xa59b/0xa59c), Word 6/95 (0xa5dc, nFib 101–104) and Word for Mac 5.x (0xa5db) have an
  // older FIB; Word 97 and later always write wIdent 0xa5ec with nFib ≥ 0xc1 (betas 0xc0)
  if ([0xa59b, 0xa59c, 0xa5db, 0xa5dc].includes(wIdent) || (wIdent === 0xa5ec && nFib < 0xc0)) throw tooOld();
  if (wIdent !== 0xa5ec) throw new Error('This is not a Word 97–2003 document.');
  const flags = U16(0x0a);
  if (flags & 0x0100) throw new Error('This Word document is password-protected – it can only be opened with the desktop app.');
  const table = cfb.streams.get(flags & 0x0200 ? '1Table' : '0Table');
  if (!table) throw new Error('This Word document is damaged (no table stream).');
  const data = cfb.streams.get('Data') || new Uint8Array(0);
  const tdv = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const tU16 = (o) => tdv.getUint16(o, true); const tU32 = (o) => tdv.getUint32(o, true); const tI32 = (o) => tdv.getInt32(o, true);
  const csw = U16(0x20);
  const lwStart = 0x22 + csw * 2 + 2;
  const lw = (k) => I32(lwStart + k * 4);
  const cslw = U16(0x22 + csw * 2);
  const fcStart = lwStart + cslw * 4 + 2;
  const fcLcb = (k) => ({ fc: U32(fcStart + k * 8), lcb: U32(fcStart + k * 8 + 4) });
  const ccpText = lw(3); const ccpFtn = lw(4); const ccpHdd = lw(5); const ccpAtn = lw(7);

  const model = defaultModel(settings);
  model.defaults.spaceAfter = 0; model.defaults.lineHeight = 1;
  const fid = new Set();

  // ---- piece table → full text with cp→fc mapping
  const clx = fcLcb(33);
  let p = clx.fc; const clxEnd = clx.fc + clx.lcb;
  while (p < clxEnd && table[p] === 0x01) p += 3 + tU16(p + 1);
  if (table[p] !== 0x02) throw new Error('This Word document is damaged (no piece table).');
  const lcbPcd = tU32(p + 1);
  const plc = p + 5;
  const nPieces = (lcbPcd - 4) / 12;
  const pieces = [];
  for (let k = 0; k < nPieces; k++) {
    const cpS = tU32(plc + k * 4); const cpE = tU32(plc + (k + 1) * 4);
    const pcd = plc + (nPieces + 1) * 4 + k * 8;
    let fc = tU32(pcd + 2);
    const compressed = !!(fc & 0x40000000);
    fc &= 0x3fffffff;
    if (compressed) fc /= 2;
    pieces.push({ cpS, cpE, fc, compressed });
  }
  const cp1252 = new TextDecoder('windows-1252');
  let text = '';
  for (const pc of pieces) {
    const n = pc.cpE - pc.cpS;
    text += pc.compressed ? cp1252.decode(wd.subarray(pc.fc, pc.fc + n)) : new TextDecoder('utf-16le').decode(wd.subarray(pc.fc, pc.fc + n * 2));
  }
  const fcOf = (cp) => { for (const pc of pieces) if (cp >= pc.cpS && cp < pc.cpE) return pc.fc + (cp - pc.cpS) * (pc.compressed ? 1 : 2); return -1; };

  // ---- FKPs (character and paragraph runs by fc)
  const fkps = (idx, kind) => {
    const { fc, lcb } = fcLcb(idx);
    const out = [];
    if (!lcb) return out;
    const n = (lcb - 4) / 8;
    for (let k = 0; k < n; k++) {
      const pn = tU32(fc + (n + 1) * 4 + k * 4) & 0x3fffff;
      const base = pn * 512;
      if (base + 512 > wd.length) continue;
      const cnt = wd[base + 511];
      for (let r = 0; r < cnt; r++) {
        const fs = U32(base + r * 4); const fe = U32(base + (r + 1) * 4);
        if (kind === 'chp') {
          const off = wd[base + (cnt + 1) * 4 + r] * 2;
          const grp = off ? sprms(wd, base + off + 1, base + off + 1 + wd[base + off]) : [];
          out.push({ fs, fe, grp });
        } else {
          const off = wd[base + (cnt + 1) * 4 + r * 13] * 2;
          let grp = []; let istd = 0;
          if (off) {
            let cb = wd[base + off]; let s = base + off + 1;
            let size;
            if (cb === 0) { cb = wd[s]; s += 1; size = cb * 2; } else size = cb * 2 - 1;
            istd = U16(s);
            grp = sprms(wd, s + 2, s + size);
          }
          out.push({ fs, fe, grp, istd });
        }
      }
    }
    out.sort((a, b) => a.fs - b.fs);
    return out;
  };
  const chpRuns = fkps(12, 'chp');
  const papRuns = fkps(13, 'pap');
  const findRun = (runs, fc) => { let lo = 0; let hi = runs.length - 1; while (lo <= hi) { const m = (lo + hi) >> 1; if (fc < runs[m].fs) hi = m - 1; else if (fc >= runs[m].fe) lo = m + 1; else return runs[m]; } return null; };

  // ---- styles (names, heading level)
  const styles = [];
  try {
    const { fc, lcb } = fcLcb(1);
    if (lcb) {
      const cbStshi = tU16(fc);
      const cstd = tU16(fc + 2);
      const cbBase = tU16(fc + 4);
      let q = fc + 2 + cbStshi;
      for (let k = 0; k < cstd && q < fc + lcb; k++) {
        const cb = tU16(q);
        if (cb) {
          const s = q + 2;
          const sti = tU16(s) & 0x0fff;
          const nameAt = s + cbBase;
          const len = tU16(nameAt);
          let name = '';
          for (let j = 0; j < len && j < 128; j++) name += String.fromCharCode(tU16(nameAt + 2 + j * 2));
          styles[k] = { sti, name: name.toLowerCase() };
        }
        q += 2 + cb;
      }
    }
  } catch { /* styles are optional */ }
  const styleRole = (istd) => {
    const s = styles[istd];
    if (!s) return {};
    if (s.sti >= 1 && s.sti <= 9) return { heading: Math.min(6, s.sti) };
    const hm = /^(?:heading|überschrift)\s*(\d)/.exec(s.name); if (hm) return { heading: Math.min(6, Number(hm[1])) };
    if (s.sti === 62 || s.name === 'title' || s.name === 'titel') return { pstyle: 'title' };
    if (s.sti === 74 || s.name === 'subtitle' || s.name === 'untertitel') return { pstyle: 'subtitle' };
    if (/^toc \d|^verzeichnis \d/.test(s.name)) return { toc: true };
    return {};
  };

  // ---- fonts
  const fonts = [];
  try {
    const { fc, lcb } = fcLcb(15);
    if (lcb) {
      let q = fc; const cData = tU16(q) === 0xffff ? tU16(q + 2) : tU16(q);
      q += tU16(fc) === 0xffff ? 6 : 4;
      for (let k = 0; k < cData; k++) {
        const cb = table[q];
        let name = ''; let j = q + 40;
        while (j + 1 < q + 1 + cb) { const c = tU16(j); if (!c) break; name += String.fromCharCode(c); j += 2; }
        fonts.push(name);
        q += cb + 1;
      }
    }
  } catch { /* optional */ }
  // revision authors
  const authors = [];
  try {
    const { fc, lcb } = fcLcb(51);
    if (lcb) {
      const ext = tU16(fc) === 0xffff; let q = fc + (ext ? 2 : 0);
      const n = tU16(q); q += 4;
      for (let k = 0; k < n; k++) { const cch = tU16(q); let s = ''; for (let j = 0; j < cch; j++) s += String.fromCharCode(tU16(q + 2 + j * 2)); authors.push(s); q += 2 + cch * 2; }
    }
  } catch { /* optional */ }
  // lists: ilfo → { lsid } → levels nfc
  const lfo = []; const lists = new Map();
  try {
    const lst = fcLcb(73); const lf = fcLcb(74);
    if (lst.lcb) {
      let q = lst.fc; const cLst = tU16(q); q += 2;
      const heads = [];
      for (let k = 0; k < cLst; k++) { heads.push({ lsid: tI32(q), simple: !!(table[q + 26] & 1) }); q += 28; }
      for (const h of heads) {
        const levels = [];
        for (let l = 0; l < (h.simple ? 1 : 9); l++) {
          const start = tI32(q); const nfc = table[q + 4];
          const cbChpx = table[q + 24]; const cbPapx = table[q + 25];
          q += 28 + cbPapx + cbChpx;
          const cch = tU16(q); q += 2 + cch * 2;
          levels.push({ start, nfc });
        }
        lists.set(h.lsid, levels);
      }
    }
    if (lf.lcb) { const n = tI32(lf.fc); for (let k = 0; k < n; k++) lfo.push(tI32(lf.fc + 4 + k * 16)); }
  } catch { /* lists fall back to bullets */ }
  const listOf = (ilfo, ilvl) => {
    if (!ilfo || ilfo >= 2000) return null;
    const levels = lists.get(lfo[ilfo - 1]);
    const lv = levels ? levels[Math.min(ilvl, levels.length - 1)] : null;
    if (lv && lv.nfc === 0xff) return null;
    const k = lv ? NFC[lv.nfc] || { kind: 'ordered' } : { kind: 'bullet' };
    return { ...k, start: lv && lv.start > 0 ? lv.start : 1, numId: `d${ilfo}`, level: ilvl };
  };

  // ---- footnotes and comments (sub-documents)
  const plcPairs = (idx, dataSize) => {
    const { fc, lcb } = fcLcb(idx);
    if (!lcb) return { cps: [], data: [] };
    const n = dataSize ? Math.floor((lcb - 4) / (4 + dataSize)) : (lcb / 4) - 1;
    const cps = []; for (let k = 0; k <= n; k++) cps.push(tU32(fc + k * 4));
    const dat = []; for (let k = 0; k < n; k++) dat.push(fc + (n + 1) * 4 + k * dataSize);
    return { cps, data: dat };
  };
  const clean = (s) => s.replace(/\x13[^\x14\x15]*\x14?/g, '').replace(/[\x00-\x08\x0b-\x1f]/g, ' ').replace(/\s+/g, ' ').trim();
  const ftnRef = plcPairs(2, 2); const ftnTxt = plcPairs(3, 0);
  const footText = new Map();
  ftnRef.cps.slice(0, -1).forEach((cp, k) => { const a = ftnTxt.cps[k]; const b = ftnTxt.cps[k + 1]; if (a != null && b != null) footText.set(cp, clean(text.slice(ccpText + a, ccpText + b))); });
  const atnRef = plcPairs(4, 30); const atnTxt = plcPairs(5, 0);
  const owners = [];
  try { const { fc, lcb } = fcLcb(36); let q = fc; while (q < fc + lcb) { const cch = tU16(q); let s = ''; for (let j = 0; j < cch; j++) s += String.fromCharCode(tU16(q + 2 + j * 2)); owners.push(s); q += 2 + cch * 2; } } catch { /* optional */ }
  const annAt = new Map();
  const atnBase = ccpText + ccpFtn + ccpHdd;
  atnRef.cps.slice(0, -1).forEach((cp, k) => {
    const a = atnTxt.cps[k]; const b = atnTxt.cps[k + 1];
    const ibst = atnRef.data[k] != null ? tU16(atnRef.data[k] + 20) : 0;
    if (a != null && b != null) annAt.set(cp, { author: owners[ibst] || '', text: clean(text.slice(atnBase + a, atnBase + b)) });
  });

  // ---- pictures in the Data stream
  const picture = (loc) => {
    if (!data.length || loc <= 0 || loc + 68 > data.length) return null;
    const ddv = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const lcb = ddv.getInt32(loc, true);
    const end = Math.min(data.length, loc + Math.max(0, lcb));
    const dxa = ddv.getInt16(loc + 28, true); const mx = ddv.getUint16(loc + 32, true) || 1000;
    let width = dxa > 0 ? Math.round((dxa * mx) / 1000 / 15) : null;
    for (let k = loc + 68; k < end - 8; k++) {
      if (data[k] === 0x89 && data[k + 1] === 0x50 && data[k + 2] === 0x4e && data[k + 3] === 0x47) {
        let e = k + 8;
        while (e + 8 < end) { const len = (data[e] << 24 | data[e + 1] << 16 | data[e + 2] << 8 | data[e + 3]) >>> 0; const ty = String.fromCharCode(data[e + 4], data[e + 5], data[e + 6], data[e + 7]); e += 12 + len; if (ty === 'IEND') break; }
        const bytes = data.slice(k, Math.min(e, end));
        if (!width) { const s = imageSize(bytes); width = s ? Math.min(s.width, 600) : null; }
        return { src: dataUrl(bytes, 'image/png'), width };
      }
      if (data[k] === 0xff && data[k + 1] === 0xd8 && data[k + 2] === 0xff) {
        let e = end; for (let j = end - 2; j > k; j--) if (data[j] === 0xff && data[j + 1] === 0xd9) { e = j + 2; break; }
        const bytes = data.slice(k, e);
        if (!width) { const s = imageSize(bytes); width = s ? Math.min(s.width, 600) : null; }
        return { src: dataUrl(bytes, 'image/jpeg'), width };
      }
    }
    fid.add('images');
    return null;
  };

  // ---- walk the main text
  const flat = [];
  let inl = [];
  let row = null; let rows = null; let cellBlocks = [];
  const fieldStack = [];
  const charProps = (cp) => {
    const r = findRun(chpRuns, fcOf(cp));
    const c = {};
    for (const [s, op] of (r ? r.grp : [])) {
      switch (s) {
        case 0x0835: c.b = op[0] === 1 || op[0] === 0x81; break;
        case 0x0836: c.i = op[0] === 1 || op[0] === 0x81; break;
        case 0x0837: case 0x2a53: c.strike = op[0] === 1; break;
        case 0x2a3e: c.u = op[0] !== 0; break;
        case 0x4a43: c.size = u16(op) / 2; break;
        case 0x2a42: c.color = ICO[op[0]] || null; break;
        case 0x6870: if (op[3] === 0) c.color = `#${[op[0], op[1], op[2]].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`; break;
        case 0x4a4f: c.font = fonts[u16(op)] || null; break;
        case 0x2a48: c.sup = op[0] === 1; c.sub = op[0] === 2; break;
        case 0x2a0c: c.hl = ICO[op[0]] || null; break;
        case 0x6a03: c.pic = (op[0] | op[1] << 8 | op[2] << 16 | op[3] << 24) >>> 0; break;
        case 0x0855: c.spec = op[0] === 1; break;
        case 0x0800: c.del = op[0] === 1; break;
        case 0x0801: c.ins = op[0] === 1; break;
        case 0x4804: c.rAuthor = u16(op); break;
        case 0x083c: c.hidden = op[0] === 1; break;
        case 0x0806: c.ole = op[0] === 1; break;
        default:
      }
    }
    return c;
  };
  const marksOf = (c) => {
    const m = [];
    if (c.b) m.push({ type: 'bold' }); if (c.i) m.push({ type: 'italic' }); if (c.u) m.push({ type: 'underline' }); if (c.strike) m.push({ type: 'strike' });
    if (c.sup) m.push({ type: 'superscript' }); if (c.sub) m.push({ type: 'subscript' });
    const ts = { fontFamily: c.font && c.font !== model.defaults.font ? c.font : null, fontSize: c.size && c.size !== model.defaults.fontSize ? `${c.size}pt` : null, color: c.color && c.color !== '#000000' ? c.color : null };
    if (ts.fontFamily || ts.fontSize || ts.color) m.push({ type: 'textStyle', attrs: ts });
    if (c.hl) m.push({ type: 'highlight', attrs: { color: c.hl } });
    for (const f of fieldStack) if (f.href && f.phase === 'result') m.push({ type: 'link', attrs: { href: f.href } });
    if (c.ins) m.push({ type: 'insertion', attrs: { id: `r${c.rAuthor || 0}`, author: authors[c.rAuthor] || '', date: null } });
    if (c.del) m.push({ type: 'deletion', attrs: { id: `r${c.rAuthor || 0}`, author: authors[c.rAuthor] || '', date: null } });
    return m;
  };
  const paraProps = (cp) => {
    const r = findRun(papRuns, fcOf(cp));
    const a = { istd: r ? r.istd : 0 };
    for (const [s, op] of (r ? r.grp : [])) {
      switch (s) {
        case 0x2403: case 0x2461: a.jc = op[0]; break;
        case 0x2416: a.inTable = op[0] === 1; break;
        case 0x2417: a.ttp = op[0] === 1; break;
        case 0x6649: a.itap = op[0] | op[1] << 8; if (a.itap > 0) a.inTable = true; break;
        case 0x244b: a.inTable = op[0] === 1; break;
        case 0x244c: a.ttp = op[0] === 1; break;
        case 0x460b: a.ilfo = u16(op); break;
        case 0x260a: a.ilvl = op[0]; break;
        case 0x840f: case 0x845e: a.left = i16(op); break;
        case 0x8411: case 0x8460: a.first = i16(op); break;
        case 0xa413: a.before = u16(op); break;
        case 0xa414: a.after = u16(op); break;
        case 0x6412: { const dya = i16(op); const mult = op[2] | op[3] << 8; if (mult && dya > 0) a.line = Math.round((dya / 240) * 100) / 100; break; }
        case 0x2640: a.outline = op[0]; break;
        case 0x2407: a.pageBreak = op[0] === 1; break;
        default:
      }
    }
    return a;
  };
  const endParagraph = (endCp, kind) => {
    const pa = paraProps(endCp);
    const role = styleRole(pa.istd);
    const attrs = {};
    if (pa.jc === 1) attrs.textAlign = 'center'; else if (pa.jc === 2) attrs.textAlign = 'right'; else if (pa.jc === 3 || pa.jc === 4) attrs.textAlign = 'justify';
    if (pa.before) attrs.spaceBefore = Math.round(pa.before / 20);
    if (pa.after != null) attrs.spaceAfter = Math.round(pa.after / 20);
    if (pa.line) attrs.lineHeight = pa.line;
    if (pa.left) attrs.indent = Math.round(pa.left / 20);
    if (pa.first) attrs.firstLine = Math.round(pa.first / 20);
    if (role.pstyle) attrs.pstyle = role.pstyle;
    const level = role.heading || (pa.outline != null && pa.outline < 6 ? pa.outline + 1 : 0);
    const content = normalizeInline(inl);
    inl = [];
    const node = level ? { type: 'heading', attrs: { level, ...attrs }, content } : { type: 'paragraph', attrs, content };
    if (!content.length) delete node.content;
    if (pa.inTable) {
      cellBlocks.push(node);
      if (kind === 'cell') {
        if (pa.ttp) { // row end mark
          if (row && row.length) { rows = rows || []; rows.push({ type: 'tableRow', content: row }); }
          row = null; cellBlocks = [];
        } else { row = row || []; row.push({ type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: null, backgroundColor: null }, content: cellBlocks.length ? cellBlocks : [{ type: 'paragraph' }] }); cellBlocks = []; }
      }
      return;
    }
    if (rows) { flat.push({ node: { type: 'table', content: rows } }); rows = null; }
    if (pa.pageBreak) flat.push({ node: { type: 'pageBreak' } });
    let list = null;
    if (pa.ilfo) { list = listOf(pa.ilfo, pa.ilvl || 0); if (list) { delete attrs.indent; delete attrs.firstLine; } }
    flat.push({ node, list });
  };

  for (let cp = 0; cp < ccpText && cp < text.length; cp++) {
    const ch = text[cp];
    const code = ch.charCodeAt(0);
    if (code === 0x0d) { endParagraph(cp, 'par'); continue; }
    if (code === 0x07) { endParagraph(cp, 'cell'); continue; }
    if (code === 0x0c) { if (inl.length) endParagraph(cp, 'par'); if (rows) { flat.push({ node: { type: 'table', content: rows } }); rows = null; } flat.push({ node: { type: 'pageBreak' } }); continue; }
    if (code === 0x13) { fieldStack.push({ instr: '', phase: 'instr', start: inl.length }); continue; }
    if (code === 0x14) { const f = fieldStack[fieldStack.length - 1]; if (f) { f.phase = 'result'; const m = /HYPERLINK\s+(?:\\l\s+)?"([^"]+)"/i.exec(f.instr); if (m) f.href = m[1]; } continue; }
    if (code === 0x15) { fieldStack.pop(); continue; }
    const f = fieldStack[fieldStack.length - 1];
    if (f && f.phase === 'instr') { f.instr += ch; continue; }
    const c = charProps(cp);
    if (c.hidden && !c.del) continue;
    if (code === 0x0b) { inl.push({ type: 'hardBreak' }); continue; }
    if (c.spec) {
      if (code === 0x01 && !c.ole) { const pic = picture(c.pic); if (pic) inl.push({ type: 'image', attrs: { src: pic.src, alt: '', title: null, width: pic.width, height: null, align: null, wrap: null } }); continue; }
      if (code === 0x01 && c.ole) { fid.add('ole'); continue; }
      if (code === 0x08) { fid.add('shape'); continue; }
      if (code === 0x02) { if (footText.has(cp)) inl.push({ type: 'footnote', attrs: { text: footText.get(cp) } }); continue; }
      if (code === 0x05) {
        const an = annAt.get(cp);
        if (an) {
          const id = `w${cp}`;
          model.comments.push({ id, author: an.author, date: null, text: an.text, resolved: false, replies: [] });
          const last = [...inl].reverse().find((n) => n.type === 'text');
          if (last) last.marks = [...(last.marks || []), { type: 'comment', attrs: { id } }];
        }
        continue;
      }
      continue;
    }
    let t = ch;
    if (code === 0x1e) t = '‑'; else if (code === 0x1f) continue; else if (code < 0x20 && code !== 0x09) continue;
    // merge with the previous text node of the same cp formatting run
    inl.push({ type: 'text', text: t, marks: marksOf(c) });
  }
  if (inl.length) endParagraph(ccpText - 1, 'par');
  if (rows) flat.push({ node: { type: 'table', content: rows } });
  model.doc = fixDoc({ type: 'doc', content: groupBlocks(flat) });

  // ---- page setup (last section)
  try {
    const sed = plcPairs(6, 12);
    const last = sed.data[sed.data.length - 1];
    if (last != null) {
      const fcSepx = tI32(last + 2);
      if (fcSepx > 0 && fcSepx + 2 < wd.length) {
        const cb = U16(fcSepx);
        let w = null; let h = null; let land = false;
        for (const [s, op] of sprms(wd, fcSepx + 2, fcSepx + 2 + cb)) {
          if (s === 0xb01f) w = u16(op); else if (s === 0xb020) h = u16(op);
          else if (s === 0xb021) model.page.margins.left = twipToMm(u16(op)); else if (s === 0xb022) model.page.margins.right = twipToMm(u16(op));
          else if (s === 0x9023) model.page.margins.top = twipToMm(Math.abs(i16(op))); else if (s === 0x9024) model.page.margins.bottom = twipToMm(Math.abs(i16(op)));
          else if (s === 0x301d) land = op[0] === 2; else if (s === 0x500b) model.page.columns = Math.min(3, u16(op) + 1);
          else if (s === 0x300a) model.titlePage = op[0] === 1;
        }
        if (w && h) { const W = twipToMm(w); const H = twipToMm(h); model.page.width = Math.min(W, H); model.page.height = Math.max(W, H); model.page.size = sizeName(W, H); model.page.orientation = land || W > H ? 'landscape' : 'portrait'; }
      }
    }
  } catch { /* defaults */ }
  // ---- header/footer (first section: odd header = story 7, odd footer = story 9)
  try {
    const hdd = plcPairs(11, 0);
    const base = ccpText + ccpFtn;
    const story = (k) => { const a = hdd.cps[k]; const b = hdd.cps[k + 1]; if (a == null || b == null || b <= a) return ''; return text.slice(base + a, base + b).replace(/\x13\s*PAGE\b[^\x14\x15]*\x14?[^\x15]*\x15/gi, '{page}').replace(/\x13\s*NUMPAGES\b[^\x14\x15]*\x14?[^\x15]*\x15/gi, '{pages}'); };
    const h = clean(story(7)); const f = clean(story(9));
    if (h) model.header = { text: h, align: 'center' };
    if (f) model.footer = { text: f, align: 'center' };
    if (ccpHdd && (h || f)) fid.add('headerLayout');
  } catch { /* optional */ }
  void ccpAtn;
  // ---- title / author
  try {
    const si = cfb.streams.get('\u0005SummaryInformation');
    if (si) {
      const sdv = new DataView(si.buffer, si.byteOffset, si.byteLength);
      const off = sdv.getUint32(44, true);
      const count = sdv.getUint32(off + 4, true);
      for (let k = 0; k < count; k++) {
        const id = sdv.getUint32(off + 8 + k * 8, true); const po = off + sdv.getUint32(off + 12 + k * 8, true);
        if (sdv.getUint32(po, true) !== 30) continue;
        const len = sdv.getUint32(po + 4, true);
        const s = cp1252.decode(si.subarray(po + 8, po + 8 + len)).replace(/\0+$/, '');
        if (id === 2) model.meta.title = s; else if (id === 4) model.meta.author = s; else if (id === 3) model.meta.subject = s; else if (id === 5) model.meta.keywords = s;
      }
    }
  } catch { /* optional */ }
  if ([...cfb.streams.keys()].some((k) => /^Macros\//i.test(k))) fid.add('macros');
  if ([...cfb.streams.keys()].some((k) => /^ObjectPool\//i.test(k))) fid.add('ole');
  if (chpRuns.some((r) => r.grp.some(([s]) => s === 0x0800 || s === 0x0801))) fid.add('trackedChanges');
  model.fidelity = [...fid, 'binaryFormat'];
  model.keep.doc = { cfb: true };
  return model;
}
