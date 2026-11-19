'use strict';
/*
 * Legacy binary office files (OLE compound file / CFB) – read only, for search and the AI:
 *   readCfb(buf)        → Map(stream path → Buffer), e.g. 'WordDocument', '1Table', 'PowerPoint Document'
 *   docText(buf)        → text of a Word 97–2003 document (.doc/.dot) through the piece table
 *   pptText(buf)        → { text, slides } of a PowerPoint 97–2003 file (.ppt/.pps/.pot): text atoms per slide
 *   isCfb(buf)          → signature check
 * Nothing is interpreted beyond text records – macros or embedded objects are never touched.
 */
const SIG = Buffer.from('d0cf11e0a1b11ae1', 'hex');
const END = 0xfffffffe; const FREE = 0xffffffff;
const isCfb = (buf) => !!buf && buf.length >= 512 && buf.subarray(0, 8).equals(SIG);

function readCfb(buf) {
  if (!isCfb(buf)) throw new Error('Not an OLE compound file');
  const shift = buf.readUInt16LE(0x1e); const ss = 1 << shift;
  const miniSs = 1 << buf.readUInt16LE(0x20);
  const nFat = buf.readUInt32LE(0x2c);
  const dirStart = buf.readUInt32LE(0x30);
  const cutoff = buf.readUInt32LE(0x38);
  const miniFatStart = buf.readUInt32LE(0x3c);
  let difatSector = buf.readUInt32LE(0x44);
  const sector = (n) => { const o = (n + 1) * ss; return o + ss <= buf.length ? buf.subarray(o, o + ss) : Buffer.alloc(0); };
  // FAT sectors: 109 in the header, more in DIFAT sectors
  const fatSectors = [];
  for (let i = 0; i < 109 && fatSectors.length < nFat; i++) { const s = buf.readUInt32LE(0x4c + i * 4); if (s !== FREE) fatSectors.push(s); }
  let guard = 0;
  while (difatSector !== END && difatSector !== FREE && fatSectors.length < nFat && guard++ < 10000) {
    const d = sector(difatSector);
    if (!d.length) break;
    for (let i = 0; i < ss / 4 - 1 && fatSectors.length < nFat; i++) { const s = d.readUInt32LE(i * 4); if (s !== FREE) fatSectors.push(s); }
    difatSector = d.readUInt32LE(ss - 4);
  }
  const fat = [];
  for (const s of fatSectors) { const d = sector(s); for (let i = 0; i + 4 <= d.length; i += 4) fat.push(d.readUInt32LE(i)); }
  const chain = (start, table) => { const out = []; let s = start; const seen = new Set(); while (s !== END && s !== FREE && s < table.length && !seen.has(s)) { seen.add(s); out.push(s); s = table[s]; } return out; };
  const big = (start, size) => { const parts = chain(start, fat).map(sector); const b = Buffer.concat(parts); return size != null ? b.subarray(0, size) : b; };
  const dir = big(dirStart);
  const entries = [];
  for (let o = 0; o + 128 <= dir.length; o += 128) {
    const nl = dir.readUInt16LE(o + 0x40);
    entries.push({ name: dir.toString('utf16le', o, o + Math.max(0, nl - 2)), type: dir[o + 0x42], left: dir.readUInt32LE(o + 0x44), right: dir.readUInt32LE(o + 0x48), child: dir.readUInt32LE(o + 0x4c), start: dir.readUInt32LE(o + 0x74), size: dir.readUInt32LE(o + 0x78) });
  }
  const root = entries[0];
  if (!root) return new Map();
  const ministream = big(root.start, root.size);
  const miniFat = [];
  for (const s of chain(miniFatStart, fat)) { const d = sector(s); for (let i = 0; i + 4 <= d.length; i += 4) miniFat.push(d.readUInt32LE(i)); }
  const mini = (start, size) => Buffer.concat(chain(start, miniFat).map((s) => ministream.subarray(s * miniSs, s * miniSs + miniSs))).subarray(0, size);
  const out = new Map();
  // Walk the red-black tree of each storage
  const walk = (idx, prefix, depth) => {
    if (idx === FREE || idx >= entries.length || depth > 64) return;
    const e = entries[idx];
    walk(e.left, prefix, depth + 1); walk(e.right, prefix, depth + 1);
    const p = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.type === 2) out.set(p, e.size < cutoff ? mini(e.start, e.size) : big(e.start, e.size));
    else if (e.type === 1) walk(e.child, p, depth + 1);
  };
  walk(root.child, '', 0);
  return out;
}

const CP1252 = { 0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…', 0x86: '†', 0x87: '‡', 0x88: 'ˆ', 0x89: '‰', 0x8a: 'Š', 0x8b: '‹', 0x8c: 'Œ', 0x8e: 'Ž', 0x91: '‘', 0x92: '’', 0x93: '“', 0x94: '”', 0x95: '•', 0x96: '–', 0x97: '—', 0x98: '˜', 0x99: '™', 0x9a: 'š', 0x9b: '›', 0x9c: 'œ', 0x9e: 'ž', 0x9f: 'Ÿ' };
const cp1252 = (b) => { let s = ''; for (const c of b) s += CP1252[c] || String.fromCharCode(c); return s; };

/** Word control characters → plain text (field instructions dropped, results kept) */
function cleanWord(s) {
  let out = ''; let depth = 0; const inInstr = [];
  for (const ch of s) {
    const c = ch.charCodeAt(0);
    if (c === 0x13) { depth++; inInstr.push(true); continue; }
    if (c === 0x14) { if (inInstr.length) inInstr[inInstr.length - 1] = false; continue; }
    if (c === 0x15) { if (depth) { depth--; inInstr.pop(); } continue; }
    if (inInstr.some(Boolean)) continue;
    if (c === 0x0d || c === 0x0b || c === 0x0c) out += '\n';
    else if (c === 0x07) out += '\t';
    else if (c === 0x09) out += '\t';
    else if (c === 0x1e) out += '-';
    else if (c === 0x1f || c === 0x01 || c === 0x08 || c < 0x09) continue;
    else if (c < 0x20) continue;
    else out += ch;
  }
  return out.replace(/\t+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

/** Text of a .doc (Word 97–2003) through the FIB → Clx → piece table */
function docText(buf) {
  const s = readCfb(buf);
  const wd = s.get('WordDocument');
  if (!wd || wd.length < 0x200) throw new Error('No WordDocument stream');
  if (wd.readUInt16LE(0) !== 0xa5ec) throw new Error('Not a Word document');
  const flags = wd.readUInt16LE(0x0a);
  if (flags & 0x0100) throw new Error('The document is encrypted');
  const table = s.get(flags & 0x0200 ? '1Table' : '0Table');
  // FibBase (32) → csw + fibRgW → cslw + fibRgLw → cbRgFcLcb + fibRgFcLcb
  let p = 32;
  const csw = wd.readUInt16LE(p); p += 2 + csw * 2;
  const cslw = wd.readUInt16LE(p); p += 2 + cslw * 4;
  const ccpText = wd.readInt32LE(32 + 2 + csw * 2 + 2 + 3 * 4); // fibRgLw.ccpText
  p += 2;
  const fcClx = wd.readUInt32LE(p + 66 * 8); const lcbClx = wd.readUInt32LE(p + 66 * 8 + 4);
  let text = '';
  if (table && lcbClx > 0 && fcClx + lcbClx <= table.length) {
    const clx = table.subarray(fcClx, fcClx + lcbClx);
    let i = 0;
    while (i < clx.length && clx[i] === 0x01) i += 3 + clx.readUInt16LE(i + 1); // Prc
    if (clx[i] === 0x02) {
      const lcb = clx.readUInt32LE(i + 1); const plc = clx.subarray(i + 5, i + 5 + lcb);
      const n = Math.floor((lcb - 4) / 12);
      for (let k = 0; k < n; k++) {
        const cpA = plc.readUInt32LE(k * 4); const cpB = plc.readUInt32LE((k + 1) * 4);
        const pcd = (n + 1) * 4 + k * 8;
        const fcRaw = plc.readUInt32LE(pcd + 2);
        const count = cpB - cpA;
        if (count <= 0 || count > 50e6) continue;
        if (fcRaw & 0x40000000) { const fc = (fcRaw & 0x3fffffff) / 2; text += cp1252(wd.subarray(fc, fc + count)); } else text += wd.toString('utf16le', fcRaw, fcRaw + count * 2);
      }
    }
  }
  if (!text) { // very old files without a piece table: 8-bit text right after the FIB
    const fcMin = wd.readUInt32LE(0x18); const fcMac = wd.readUInt32LE(0x1c);
    text = cp1252(wd.subarray(fcMin, Math.min(wd.length, fcMac)));
  }
  if (ccpText > 0 && ccpText < text.length) text = text.slice(0, ccpText); // main document only (no headers/footnotes)
  return cleanWord(text);
}

/** Text of a PowerPoint 97–2003 file: text atoms of the slides (master placeholders skipped) */
function pptText(buf) {
  const s = readCfb(buf);
  const doc = s.get('PowerPoint Document');
  if (!doc) throw new Error('No PowerPoint Document stream');
  // Text lives either in the outline (SlideListWithText) or – PowerPoint 2007+ – in the shapes of each SlideContainer
  const listed = []; const drawn = []; const notes = []; let cur = null;
  const add = (t) => { const v = t.replace(/\r/g, '\n').replace(/\u000b/g, '\n').trim(); if (v && cur) cur.push(v); };
  const walk = (start, end, ctx, depth) => {
    let o = start;
    while (o + 8 <= end && depth < 32) {
      const verInst = doc.readUInt16LE(o); const type = doc.readUInt16LE(o + 2); const len = doc.readUInt32LE(o + 4);
      const body = o + 8; const stop = Math.min(end, body + len);
      if ((verInst & 0x0f) === 0x0f) {
        const inst = verInst >> 4;
        if (type === 0x03f8) { /* MainMaster – template texts */ } else if (type === 0x03ee) { cur = []; drawn.push(cur); walk(body, stop, ctx, depth + 1); cur = null; } else if (type === 0x03f0) { cur = []; notes.push(cur); walk(body, stop, ctx, depth + 1); cur = null; } else if (type === 0x0ff0) { walk(body, stop, { list: inst }, depth + 1); cur = null; } else walk(body, stop, ctx, depth + 1);
      } else if (type === 0x03f3 && ctx.list === 0) { cur = []; listed.push(cur); } // SlidePersistAtom in the slide list
      else if (type === 0x0fa0) add(doc.toString('utf16le', body, stop));
      else if (type === 0x0fa8) add(cp1252(doc.subarray(body, stop)));
      o = body + len;
    }
  };
  walk(0, doc.length, {}, 0);
  const n = Math.max(listed.length, drawn.length);
  const parts = [];
  for (let i = 0; i < n; i++) {
    const seen = new Set(); const lines = [];
    for (const t of [...(listed[i] || []), ...(drawn[i] || [])]) if (!seen.has(t)) { seen.add(t); lines.push(t); }
    parts.push(`--- ${i + 1} ---\n${lines.join('\n')}`);
  }
  const nt = notes.map((x) => x.filter((t) => !/^\d+$/.test(t)).join('\n')).filter(Boolean);
  if (nt.length) parts.push(nt.join('\n'));
  return { text: parts.join('\n\n').trim(), slides: listed.length || drawn.length };
}

module.exports = { readCfb, docText, pptText, isCfb, cleanWord, cp1252 };
