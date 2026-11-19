'use strict';
/*
 * Small ZIP writer and reader (no dependencies, Node's zlib only).
 *   ZipWriter: entries one after another into a sink (Buffer chunks) – used for zip downloads (streamed) and for
 *              building test fixtures / office files. UTF-8 names, deflate or store, folders. No ZIP64 (≤ 4 GB).
 *   readZip:   central directory → Map(name → { size, method, read() }) – with a cap on the unpacked size (zip bombs).
 */
const zlib = require('zlib');

let TABLE = null;
function crc32(buf) {
  if (typeof zlib.crc32 === 'function') return zlib.crc32(buf) >>> 0;
  if (!TABLE) { TABLE = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; TABLE[n] = c >>> 0; } }
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** Date → DOS time/date words */
function dosTime(d) {
  const t = d instanceof Date && !Number.isNaN(d.getTime()) ? d : new Date();
  const y = Math.max(1980, Math.min(2107, t.getFullYear()));
  return { time: (t.getHours() << 11) | (t.getMinutes() << 5) | Math.floor(t.getSeconds() / 2), date: ((y - 1980) << 9) | ((t.getMonth() + 1) << 5) | t.getDate() };
}
const MAX32 = 0xffffffff;
/** Names in the archive: forward slashes, no leading slash, no "..". */
function cleanName(name) {
  const parts = String(name || '').replace(/\\/g, '/').split('/').filter((p) => p && p !== '.' && p !== '..');
  return parts.join('/');
}

class ZipWriter {
  /** @param {(chunk: Buffer) => any} sink called synchronously for every chunk (return value is passed back from add) */
  constructor(sink) {
    this.chunks = sink ? null : [];
    this.sink = sink || ((b) => { this.chunks.push(b); return true; });
    this.offset = 0;
    this.central = [];
    this.names = new Set();
  }
  _write(b) { this.offset += b.length; return this.sink(b); }
  /**
   * Add a file. data: Buffer | Uint8Array | string. opts: { mtime: Date|number, store: bool, level }
   * → the sink's return value (false = wait for 'drain' when writing to a stream)
   */
  add(name, data, opts = {}) {
    const n = cleanName(name);
    if (!n) throw new Error('zip: empty entry name');
    if (this.names.has(n)) throw new Error(`zip: duplicate entry ${n}`);
    this.names.add(n);
    const raw = Buffer.isBuffer(data) ? data : typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data || []);
    const crc = crc32(raw);
    let method = 0; let body = raw;
    if (!opts.store && raw.length > 64) {
      const d = zlib.deflateRawSync(raw, { level: opts.level != null ? opts.level : 6 });
      if (d.length < raw.length) { method = 8; body = d; }
    }
    if (raw.length >= MAX32 || body.length >= MAX32 || this.offset >= MAX32) throw new Error('zip: archive too large (more than 4 GB)');
    return this._entry(n, raw.length, body, crc, method, opts.mtime, false);
  }
  /** Add an (empty) folder entry */
  addDir(name, opts = {}) {
    const n = `${cleanName(name)}/`;
    if (n === '/' || this.names.has(n)) return true;
    this.names.add(n);
    return this._entry(n, 0, Buffer.alloc(0), 0, 0, opts.mtime, true);
  }
  _entry(name, size, body, crc, method, mtime, isDir) {
    const nameBuf = Buffer.from(name, 'utf8');
    const { time, date } = dosTime(mtime instanceof Date ? mtime : mtime ? new Date(mtime) : new Date());
    const h = Buffer.alloc(30);
    h.writeUInt32LE(0x04034b50, 0); h.writeUInt16LE(20, 4); h.writeUInt16LE(0x0800, 6); h.writeUInt16LE(method, 8);
    h.writeUInt16LE(time, 10); h.writeUInt16LE(date, 12); h.writeUInt32LE(crc, 14); h.writeUInt32LE(body.length, 18);
    h.writeUInt32LE(size, 22); h.writeUInt16LE(nameBuf.length, 26); h.writeUInt16LE(0, 28);
    const offset = this.offset;
    this._write(h); this._write(nameBuf);
    const r = body.length ? this._write(body) : true;
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(20, 4); c.writeUInt16LE(20, 6); c.writeUInt16LE(0x0800, 8); c.writeUInt16LE(method, 10);
    c.writeUInt16LE(time, 12); c.writeUInt16LE(date, 14); c.writeUInt32LE(crc, 16); c.writeUInt32LE(body.length, 20); c.writeUInt32LE(size, 24);
    c.writeUInt16LE(nameBuf.length, 28); c.writeUInt16LE(0, 30); c.writeUInt16LE(0, 32); c.writeUInt16LE(0, 34); c.writeUInt16LE(0, 36);
    c.writeUInt32LE(isDir ? 0x10 : 0, 38); c.writeUInt32LE(offset, 42);
    this.central.push(Buffer.concat([c, nameBuf]));
    return r;
  }
  /** Write the central directory. Without a sink → the whole archive as Buffer. */
  end() {
    const start = this.offset;
    for (const c of this.central) this._write(c);
    const e = Buffer.alloc(22);
    e.writeUInt32LE(0x06054b50, 0); e.writeUInt16LE(0, 4); e.writeUInt16LE(0, 6);
    e.writeUInt16LE(Math.min(0xffff, this.central.length), 8); e.writeUInt16LE(Math.min(0xffff, this.central.length), 10);
    e.writeUInt32LE(this.offset - start, 12); e.writeUInt32LE(start, 16); e.writeUInt16LE(0, 20);
    this._write(e);
    return this.chunks ? Buffer.concat(this.chunks) : null;
  }
}

/** Build a zip in memory: files = { name: Buffer|string } or [{ name, data, store }] */
function makeZip(files) {
  const z = new ZipWriter();
  const list = Array.isArray(files) ? files : Object.entries(files).map(([name, data]) => ({ name, data }));
  for (const f of list) { if (f.dir || f.name.endsWith('/')) z.addDir(f.name); else z.add(f.name, f.data, { store: f.store, mtime: f.mtime }); }
  return z.end();
}

/**
 * Read a zip from a Buffer. → Map(name → { name, size, csize, method, dir, read(): Buffer })
 * maxSize: cap per entry (unpacked) – throws beyond. Names are taken as UTF-8 (flag 11) or latin1.
 */
function readZip(buf, { maxSize = 256 * 1048576 } = {}) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  const min = Math.max(0, buf.length - 65557);
  let eocd = -1;
  for (let i = buf.length - 22; i >= min; i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Not a zip file');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  if (p === MAX32) throw new Error('ZIP64 archives are not supported');
  const out = new Map();
  for (let i = 0; i < count; i++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) throw new Error('Broken zip directory');
    const flags = buf.readUInt16LE(p + 8);
    const method = buf.readUInt16LE(p + 10);
    const csize = buf.readUInt32LE(p + 20);
    const size = buf.readUInt32LE(p + 24);
    const nl = buf.readUInt16LE(p + 28); const xl = buf.readUInt16LE(p + 30); const cl = buf.readUInt16LE(p + 32);
    const off = buf.readUInt32LE(p + 42);
    const name = buf.toString(flags & 0x0800 ? 'utf8' : 'latin1', p + 46, p + 46 + nl);
    p += 46 + nl + xl + cl;
    const read = () => {
      if (off + 30 > buf.length || buf.readUInt32LE(off) !== 0x04034b50) throw new Error(`Broken zip entry ${name}`);
      const start = off + 30 + buf.readUInt16LE(off + 26) + buf.readUInt16LE(off + 28);
      const data = buf.subarray(start, start + csize);
      if (flags & 1) throw new Error('Encrypted zip entries are not supported');
      if (method === 0) return Buffer.from(data);
      if (method === 8) {
        if (size > maxSize) throw new Error(`Zip entry too large: ${name}`);
        return zlib.inflateRawSync(data, { maxOutputLength: Math.max(1, Math.min(maxSize, size || maxSize)) });
      }
      throw new Error(`Unsupported zip method ${method}`);
    };
    out.set(name, { name, size, csize, method, dir: name.endsWith('/'), read });
  }
  return out;
}

/** True when the buffer starts like a zip (local header) */
const isZip = (buf) => !!buf && buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b && (buf[2] === 3 || buf[2] === 5) && (buf[3] === 4 || buf[3] === 6);

module.exports = { ZipWriter, makeZip, readZip, crc32, isZip, cleanName };
