/*
 * Read-only reader for OLE compound files (CFB / "structured storage": .doc, .xls, .ppt, vbaProject.bin, …).
 * Dependency-free, works in the browser and in Node. Written by the docs chunk, free for every editor.
 *
 *   readCfb(bytes) → { streams: Map<path, Uint8Array>, entries: [{ path, name, type: 'storage'|'stream', size }] }
 *
 * Paths use '/' between storages and never start with '/': 'WordDocument', '1Table', 'Macros/VBA/dir',
 * 'VBA/Module1'. Lookups are case-sensitive; use findStream(cfb, name) for a case-insensitive search.
 * Throws Error('Not a compound file') for other bytes.
 */
const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE = 0xffffffff; const END = 0xfffffffe;

export function isCfb(bytes) {
  const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return u.length >= 512 && SIG.every((b, i) => u[i] === b);
}

export function readCfb(input) {
  const u = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (!isCfb(u)) throw new Error('Not a compound file');
  const dv = new DataView(u.buffer, u.byteOffset, u.byteLength);
  const u32 = (o) => dv.getUint32(o, true);
  const u16 = (o) => dv.getUint16(o, true);
  const shift = u16(0x1e);
  const sz = 1 << shift;
  const miniShift = u16(0x20);
  const miniSz = 1 << miniShift;
  const nFat = u32(0x2c);
  const dirStart = u32(0x30);
  const cutoff = u32(0x38) || 4096;
  const miniFatStart = u32(0x3c);
  const nMiniFat = u32(0x40);
  let difatSector = u32(0x44);
  const nDifat = u32(0x48);
  const off = (s) => (s + 1) * sz;
  const maxSector = Math.floor((u.length - sz) / sz);

  // DIFAT → FAT sector list
  const fatSectors = [];
  for (let k = 0; k < 109 && fatSectors.length < nFat; k++) { const s = u32(0x4c + k * 4); if (s !== FREE) fatSectors.push(s); }
  for (let d = 0; d < nDifat && difatSector !== END && difatSector !== FREE && difatSector <= maxSector; d++) {
    const base = off(difatSector);
    for (let k = 0; k < sz / 4 - 1 && fatSectors.length < nFat; k++) { const s = u32(base + k * 4); if (s !== FREE) fatSectors.push(s); }
    difatSector = u32(base + sz - 4);
  }
  const perSector = sz / 4;
  const fat = new Uint32Array(fatSectors.length * perSector);
  fatSectors.forEach((s, k) => { const base = off(s); for (let j = 0; j < perSector && base + j * 4 + 4 <= u.length; j++) fat[k * perSector + j] = u32(base + j * 4); });

  const chain = (start, table, limit) => {
    const out = [];
    let s = start;
    const seen = new Set();
    while (s !== END && s !== FREE && s < table.length && !seen.has(s) && out.length < limit) { seen.add(s); out.push(s); s = table[s]; }
    return out;
  };
  const readChain = (start, size) => {
    const secs = chain(start, fat, Math.ceil((size || u.length) / sz) + 1);
    const buf = new Uint8Array(secs.length * sz);
    secs.forEach((s, k) => { const o = off(s); buf.set(u.subarray(o, Math.min(o + sz, u.length)), k * sz); });
    return size != null ? buf.subarray(0, size) : buf;
  };

  // directory
  const dirBytes = readChain(dirStart, null);
  const entries = [];
  for (let k = 0; k + 128 <= dirBytes.length; k += 128) {
    const e = new DataView(dirBytes.buffer, dirBytes.byteOffset + k, 128);
    const nameLen = e.getUint16(0x40, true);
    let name = '';
    for (let j = 0; j < Math.max(0, nameLen / 2 - 1) && j < 32; j++) name += String.fromCharCode(e.getUint16(j * 2, true));
    entries.push({ name, type: e.getUint8(0x42), left: e.getUint32(0x44, true), right: e.getUint32(0x48, true), child: e.getUint32(0x4c, true), start: e.getUint32(0x74, true), size: e.getUint32(0x78, true) });
  }
  const root = entries[0];
  if (!root) throw new Error('Damaged compound file');
  // mini stream + mini FAT
  const miniStream = root.start !== END ? readChain(root.start, root.size) : new Uint8Array(0);
  const mfBytes = nMiniFat ? readChain(miniFatStart, nMiniFat * sz) : new Uint8Array(0);
  const miniFat = new Uint32Array(mfBytes.length / 4);
  for (let k = 0; k < miniFat.length; k++) miniFat[k] = mfBytes[k * 4] | mfBytes[k * 4 + 1] << 8 | mfBytes[k * 4 + 2] << 16 | (mfBytes[k * 4 + 3] << 24 >>> 0);
  const readMini = (start, size) => {
    const secs = chain(start, miniFat, Math.ceil(size / miniSz) + 1);
    const buf = new Uint8Array(secs.length * miniSz);
    secs.forEach((s, k) => buf.set(miniStream.subarray(s * miniSz, s * miniSz + miniSz), k * miniSz));
    return buf.subarray(0, size);
  };

  const streams = new Map();
  const list = [];
  const visit = (idx, prefix, depth) => {
    if (idx === FREE || idx >= entries.length || depth > 64) return;
    const stack = [idx];
    const seen = new Set();
    // in-order walk over the sibling tree
    while (stack.length) {
      const i = stack.pop();
      if (i === FREE || i >= entries.length || seen.has(i)) continue;
      seen.add(i);
      const e = entries[i];
      stack.push(e.left, e.right);
      const path = prefix ? `${prefix}/${e.name}` : e.name;
      if (e.type === 2) {
        const data = e.size < cutoff ? readMini(e.start, e.size) : readChain(e.start, e.size);
        streams.set(path, data);
        list.push({ path, name: e.name, type: 'stream', size: e.size });
      } else if (e.type === 1) {
        list.push({ path, name: e.name, type: 'storage', size: 0 });
        visit(e.child, path, depth + 1);
      }
    }
  };
  visit(root.child, '', 0);
  return { streams, entries: list };
}

/** first stream whose path ends with `name` (case-insensitive), e.g. findStream(cfb, 'VBA/dir') */
export function findStream(cfb, name) {
  const want = String(name).toLowerCase();
  for (const [p, data] of cfb.streams) { const l = p.toLowerCase(); if (l === want || l.endsWith(`/${want}`)) return data; }
  return null;
}
