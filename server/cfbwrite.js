'use strict';
/*
 * Minimal OLE compound file (CFB, version 3, 512-byte sectors) WRITER – used for encrypted OOXML packages
 * (EncryptionInfo + EncryptedPackage + \x06DataSpaces). Streams below 4096 bytes go into the mini stream (readers decide
 * the location by size, so this is required). Reading stays in cfb.js (readCfb).
 *
 *   writeCfb({ 'EncryptionInfo': Buffer, '\x06DataSpaces/Version': Buffer, … }) → Buffer
 */
const SS = 512; const MINI = 64; const CUTOFF = 4096;
const END = 0xfffffffe; const FREE = 0xffffffff; const FATSECT = 0xfffffffd; const NOSTREAM = 0xffffffff;

/** CFB name order: shorter first, then case-insensitive (upper-case) comparison */
function cmpName(a, b) { if (a.length !== b.length) return a.length - b.length; const A = a.toUpperCase(); const B = b.toUpperCase(); return A < B ? -1 : A > B ? 1 : 0; }

function writeCfb(files) {
  // tree of storages
  const root = { name: 'Root Entry', type: 5, kids: new Map() };
  for (const [p, data] of Object.entries(files)) {
    const parts = p.split('/').filter(Boolean);
    let cur = root;
    for (let i = 0; i < parts.length - 1; i++) {
      if (!cur.kids.has(parts[i])) cur.kids.set(parts[i], { name: parts[i], type: 1, kids: new Map() });
      cur = cur.kids.get(parts[i]);
    }
    const name = parts[parts.length - 1];
    if (!name || name.length > 31) throw new Error(`cfb: invalid stream name ${p}`);
    cur.kids.set(name, { name, type: 2, data: Buffer.isBuffer(data) ? data : Buffer.from(data) });
  }
  // flatten into directory entries
  const entries = [];
  const add = (node) => { node.idx = entries.length; entries.push(node); node.left = NOSTREAM; node.right = NOSTREAM; node.child = NOSTREAM; if (node.kids) for (const k of node.kids.values()) add(k); };
  add(root);
  // children of each storage: balanced binary tree by cmpName (all black)
  const build = (list) => { if (!list.length) return NOSTREAM; const m = list.length >> 1; const n = list[m]; n.left = build(list.slice(0, m)); n.right = build(list.slice(m + 1)); return n.idx; };
  for (const e of entries) if (e.kids) e.child = build([...e.kids.values()].sort((a, b) => cmpName(a.name, b.name)));

  // mini stream
  const miniFat = []; const miniParts = [];
  for (const e of entries) {
    if (e.type !== 2 || e.data.length >= CUTOFF) continue;
    const n = Math.ceil(e.data.length / MINI);
    if (!n) { e.start = END; continue; }
    e.start = miniFat.length;
    for (let i = 0; i < n; i++) miniFat.push(i === n - 1 ? END : miniFat.length + 1);
    const pad = Buffer.alloc(n * MINI); e.data.copy(pad); miniParts.push(pad);
  }
  const ministream = Buffer.concat(miniParts);
  const big = entries.filter((e) => e.type === 2 && e.data.length >= CUTOFF);
  const secs = (n) => Math.ceil(n / SS);
  const dirSecs = secs(entries.length * 128);
  const miniFatSecs = secs(miniFat.length * 4);
  const miniStreamSecs = secs(ministream.length);
  const bigSecs = big.reduce((s, e) => s + secs(e.data.length), 0);
  const dataSecs = dirSecs + miniFatSecs + miniStreamSecs + bigSecs;
  // FAT sectors needed (each holds 128 entries, they map themselves too); DIFAT beyond 109 not needed below ~6.8 MB → support it anyway
  let fatSecs = 1; let difatSecs = 0;
  for (;;) { difatSecs = fatSecs > 109 ? Math.ceil((fatSecs - 109) / 127) : 0; const need = Math.ceil((dataSecs + fatSecs + difatSecs) / 128); if (need <= fatSecs) break; fatSecs = need; }
  const total = fatSecs + difatSecs + dataSecs;
  const fat = new Array(fatSecs * 128).fill(FREE);
  let s = 0;
  const fatStart = s; for (let i = 0; i < fatSecs; i++) fat[s++] = FATSECT;
  const difatStart = s; for (let i = 0; i < difatSecs; i++) fat[s++] = 0xfffffffc;
  const chainOf = (n) => { const start = s; for (let i = 0; i < n; i++) { fat[s] = i === n - 1 ? END : s + 1; s++; } return n ? start : END; };
  const dirStart = chainOf(dirSecs);
  const miniFatStart = chainOf(miniFatSecs);
  root.start = chainOf(miniStreamSecs); root.size = ministream.length;
  for (const e of big) e.start = chainOf(secs(e.data.length));
  if (s !== total) throw new Error('cfb: layout error');

  const out = Buffer.alloc(SS * (1 + total));
  // header
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(out, 0);
  out.writeUInt16LE(0x003e, 0x18); out.writeUInt16LE(0x0003, 0x1a); out.writeUInt16LE(0xfffe, 0x1c);
  out.writeUInt16LE(9, 0x1e); out.writeUInt16LE(6, 0x20);
  out.writeUInt32LE(0, 0x28); // directory sectors (0 for v3)
  out.writeUInt32LE(fatSecs, 0x2c); out.writeUInt32LE(dirStart, 0x30); out.writeUInt32LE(0, 0x34);
  out.writeUInt32LE(CUTOFF, 0x38); out.writeUInt32LE(miniFatSecs ? miniFatStart : END, 0x3c); out.writeUInt32LE(miniFatSecs, 0x40);
  out.writeUInt32LE(difatSecs ? difatStart : END, 0x44); out.writeUInt32LE(difatSecs, 0x48);
  for (let i = 0; i < 109; i++) out.writeUInt32LE(i < fatSecs ? fatStart + i : FREE, 0x4c + i * 4);
  const sec = (n) => SS * (n + 1);
  // DIFAT sectors
  for (let d = 0; d < difatSecs; d++) {
    const o = sec(difatStart + d);
    for (let i = 0; i < 127; i++) { const k = 109 + d * 127 + i; out.writeUInt32LE(k < fatSecs ? fatStart + k : FREE, o + i * 4); }
    out.writeUInt32LE(d === difatSecs - 1 ? END : difatStart + d + 1, o + 508);
  }
  // FAT
  for (let i = 0; i < fat.length; i++) out.writeUInt32LE(fat[i] >>> 0, sec(fatStart) + i * 4);
  // directory
  const dirBuf = Buffer.alloc(dirSecs * SS);
  for (let i = 0; i < dirSecs * 4; i++) { const o = i * 128; if (i >= entries.length) { dirBuf.writeUInt32LE(NOSTREAM, o + 0x44); dirBuf.writeUInt32LE(NOSTREAM, o + 0x48); dirBuf.writeUInt32LE(NOSTREAM, o + 0x4c); continue; }
    const e = entries[i];
    const nm = Buffer.from(`${e.name}\0`, 'utf16le'); nm.copy(dirBuf, o, 0, Math.min(64, nm.length));
    dirBuf.writeUInt16LE(Math.min(64, nm.length), o + 0x40); dirBuf[o + 0x42] = e.type; dirBuf[o + 0x43] = 1; // black
    dirBuf.writeUInt32LE(e.left >>> 0, o + 0x44); dirBuf.writeUInt32LE(e.right >>> 0, o + 0x48); dirBuf.writeUInt32LE(e.child >>> 0, o + 0x4c);
    const st = e.type === 2 ? (e.data.length ? e.start : END) : e.type === 5 ? (e.size ? e.start : END) : 0;
    dirBuf.writeUInt32LE(st >>> 0, o + 0x74);
    dirBuf.writeUInt32LE(e.type === 2 ? e.data.length : e.type === 5 ? e.size : 0, o + 0x78);
  }
  dirBuf.copy(out, sec(dirStart));
  if (miniFatSecs) { const b = Buffer.alloc(miniFatSecs * SS, 0xff); miniFat.forEach((v, i) => b.writeUInt32LE(v >>> 0, i * 4)); b.copy(out, sec(miniFatStart)); }
  if (miniStreamSecs) ministream.copy(out, sec(root.start));
  for (const e of big) e.data.copy(out, sec(e.start));
  return out;
}

module.exports = { writeCfb, cmpName };
