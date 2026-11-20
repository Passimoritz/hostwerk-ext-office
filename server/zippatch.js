'use strict';
/*
 * Patch a zip package without touching the parts we do not change: unchanged entries are copied byte for byte
 * (compressed data, CRC, method, order, extra fields dropped), changed ones are replaced, new ones appended,
 * null removes; ZIP64 archives are read and written (when the size needs it). Used by row append, clean-before-sharing and ODF encryption – "nothing gets lost".
 *
 *   entries(buf)                      → [{ name, method, flags, crc, csize, size, raw (compressed bytes), time, date }]
 *   patchZip(buf, { name: Buffer|string|null|{ data, store, raw } }, { order }) → Buffer
 */
const zlib = require('zlib');
const { crc32, centralDir, localHeader, centralHeader, endRecords } = require('./zip');


function entries(buf) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  return centralDir(buf).map((e) => ({ name: e.name, nameBuf: e.nameBuf, method: e.method, flags: e.flags, crc: e.crc, csize: e.csize, size: e.size, time: e.time, date: e.date, attr: e.attr, raw: buf.subarray(e.start, e.start + e.csize) }));
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
    const flags = (e.flags & 0x0800) | (e.flags & 0x0006); // keep UTF-8 + deflate level bits; data descriptors are resolved
    const rec = { ...e, flags, csize: e.raw.length };
    const h = localHeader(rec);
    parts.push(h, e.raw);
    central.push(centralHeader(rec, offset));
    offset += h.length + e.raw.length;
  }
  const cd = Buffer.concat(central);
  return Buffer.concat([...parts, cd, endRecords(list.length, cd.length, offset)]);
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
