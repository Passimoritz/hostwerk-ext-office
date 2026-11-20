'use strict';
/*
 * Patch a zip package without touching the parts we do not change: unchanged entries are copied byte for byte
 * (compressed data, CRC, method, order, extra fields dropped), changed ones are replaced, new ones appended,
 * null removes. Used by row append, clean-before-sharing and ODF encryption – "nothing gets lost".
 *
 *   entries(buf)                      → [{ name, method, flags, crc, csize, size, raw (compressed bytes), time, date }]
 *   patchZip(buf, { name: Buffer|string|null|{ data, store, raw } }, { order }) → Buffer
 */
const zlib = require('zlib');
const { crc32 } = require('./zip');

const MAX32 = 0xffffffff;

function entries(buf) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Not a zip file');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  if (p === MAX32) throw new Error('ZIP64 archives are not supported');
  const out = [];
  for (let i = 0; i < count; i++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) throw new Error('Broken zip directory');
    const flags = buf.readUInt16LE(p + 8); const method = buf.readUInt16LE(p + 10);
    const time = buf.readUInt16LE(p + 12); const date = buf.readUInt16LE(p + 14);
    const crc = buf.readUInt32LE(p + 16); const csize = buf.readUInt32LE(p + 20); const size = buf.readUInt32LE(p + 24);
    const nl = buf.readUInt16LE(p + 28); const xl = buf.readUInt16LE(p + 30); const cl = buf.readUInt16LE(p + 32);
    const attr = buf.readUInt32LE(p + 38); const off = buf.readUInt32LE(p + 42);
    const nameBuf = buf.subarray(p + 46, p + 46 + nl);
    const name = nameBuf.toString(flags & 0x0800 ? 'utf8' : 'latin1');
    p += 46 + nl + xl + cl;
    if (off + 30 > buf.length || buf.readUInt32LE(off) !== 0x04034b50) throw new Error(`Broken zip entry ${name}`);
    const start = off + 30 + buf.readUInt16LE(off + 26) + buf.readUInt16LE(off + 28);
    if (start + csize > buf.length) throw new Error(`Broken zip entry ${name}`);
    out.push({ name, nameBuf, method, flags, crc, csize, size, time, date, attr, raw: buf.subarray(start, start + csize) });
  }
  return out;
}

/** Inflated data of an entry from entries() (cap against zip bombs) */
function inflate(e, max = 256 * 1048576) {
  if (e.flags & 1) throw new Error('Encrypted zip entries are not supported');
  if (e.method === 0) return Buffer.from(e.raw);
  if (e.method === 8) { if (e.size > max) throw new Error(`Zip entry too large: ${e.name}`); return zlib.inflateRawSync(e.raw, { maxOutputLength: Math.max(1, Math.min(max, e.size || max)) }); }
  throw new Error(`Unsupported zip method ${e.method}`);
}

function dos(d = new Date()) { const y = Math.max(1980, d.getFullYear()); return { time: (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2), date: ((y - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate() }; }

/** Fresh entry record. v: Buffer|string or { data, store } or { raw, method, crc, size } (already compressed/encrypted) */
function makeEntry(name, v) {
  const now = dos();
  if (v && v.raw) return { name, nameBuf: Buffer.from(name, 'utf8'), method: v.method || 0, flags: 0x0800, crc: v.crc >>> 0, csize: v.raw.length, size: v.size, ...now, attr: 0, raw: v.raw };
  const o = v && !Buffer.isBuffer(v) && typeof v === 'object' && !(v instanceof Uint8Array) ? v : { data: v };
  const data = Buffer.isBuffer(o.data) ? o.data : typeof o.data === 'string' ? Buffer.from(o.data, 'utf8') : Buffer.from(o.data || []);
  let method = 0; let raw = data;
  if (!o.store && data.length > 64) { const d = zlib.deflateRawSync(data, { level: 6 }); if (d.length < data.length) { method = 8; raw = d; } }
  return { name, nameBuf: Buffer.from(name, 'utf8'), method, flags: 0x0800, crc: crc32(data), csize: raw.length, size: data.length, ...now, attr: 0, raw };
}

function write(list) {
  const parts = []; const central = []; let offset = 0;
  for (const e of list) {
    if (e.raw.length >= MAX32 || offset >= MAX32) throw new Error('zip: archive too large (more than 4 GB)');
    const flags = (e.flags & 0x0800) | (e.flags & 0x0006); // keep UTF-8 + deflate level bits; data descriptors are resolved
    const h = Buffer.alloc(30);
    h.writeUInt32LE(0x04034b50, 0); h.writeUInt16LE(20, 4); h.writeUInt16LE(flags, 6); h.writeUInt16LE(e.method, 8);
    h.writeUInt16LE(e.time, 10); h.writeUInt16LE(e.date, 12); h.writeUInt32LE(e.crc >>> 0, 14); h.writeUInt32LE(e.raw.length, 18);
    h.writeUInt32LE(e.size, 22); h.writeUInt16LE(e.nameBuf.length, 26); h.writeUInt16LE(0, 28);
    parts.push(h, e.nameBuf, e.raw);
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(20, 4); c.writeUInt16LE(20, 6); c.writeUInt16LE(flags, 8); c.writeUInt16LE(e.method, 10);
    c.writeUInt16LE(e.time, 12); c.writeUInt16LE(e.date, 14); c.writeUInt32LE(e.crc >>> 0, 16); c.writeUInt32LE(e.raw.length, 20); c.writeUInt32LE(e.size, 24);
    c.writeUInt16LE(e.nameBuf.length, 28); c.writeUInt32LE(e.attr || 0, 38); c.writeUInt32LE(offset, 42);
    central.push(Buffer.concat([c, e.nameBuf]));
    offset += 30 + e.nameBuf.length + e.raw.length;
  }
  const cd = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(list.length, 8); end.writeUInt16LE(list.length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...parts, cd, end]);
}

function patchZip(buf, changes = {}, { first = [] } = {}) {
  const list = entries(buf);
  const seen = new Set();
  const out = [];
  for (const e of list) {
    seen.add(e.name);
    if (!Object.prototype.hasOwnProperty.call(changes, e.name)) { out.push(e); continue; }
    const v = changes[e.name];
    if (v === null || v === undefined) continue;
    const ne = makeEntry(e.name, v); ne.time = e.time; ne.date = e.date; out.push(ne);
  }
  for (const [n, v] of Object.entries(changes)) if (!seen.has(n) && v != null) out.push(makeEntry(n, v));
  // e.g. ODF: "mimetype" must be the first entry
  for (const n of [...first].reverse()) { const i = out.findIndex((e) => e.name === n); if (i > 0) out.unshift(out.splice(i, 1)[0]); }
  return write(out);
}

module.exports = { entries, inflate, patchZip, makeEntry, writeZipEntries: write };
