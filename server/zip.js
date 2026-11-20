'use strict';
/*
 * Small ZIP writer and reader (no dependencies, Node's zlib only).
 *   ZipWriter: entries one after another into a sink (Buffer chunks) – used for zip downloads (streamed), library
 *              snapshots and test fixtures / office files. UTF-8 names, deflate or store, folders. ZIP64 records are
 *              written only when needed (an entry or offset ≥ 4 GB, more than 65 535 entries) – small archives stay
 *              classic zips that every office program reads.
 *   readZip:   central directory → Map(name → { size, method, read() }) – with a cap on the unpacked size (zip bombs).
 *              Reads ZIP64 archives (end record + locator, extra field 0x0001) – also small ones some writers make.
 *   centralDir(buf) → [{ name, nameBuf, flags, method, time, date, crc, csize, size, attr, off, start }] (shared with
 *              zippatch.js) · localHeader/centralHeader/endRecords – the record writers (ZIP64 when needed)
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
const MAX16 = 0xffff;
/** Names in the archive: forward slashes, no leading slash, no "..". */
function cleanName(name) {
  const parts = String(name || '').replace(/\\/g, '/').split('/').filter((p) => p && p !== '.' && p !== '..');
  return parts.join('/');
}
const u64 = (n) => { const b = Buffer.alloc(8); b.writeBigUInt64LE(BigInt(n)); return b; };
const r64 = (buf, p) => { const v = buf.readBigUInt64LE(p); if (v > BigInt(Number.MAX_SAFE_INTEGER)) throw new Error('Broken zip directory'); return Number(v); };

/* ---------------- record writers (shared with zippatch.js) ---------------- */
/** local file header (+ name, + ZIP64 extra with both sizes when one of them needs it) */
function localHeader({ nameBuf, flags, method, time, date, crc, csize, size }) {
  const z64 = csize >= MAX32 || size >= MAX32;
  const extra = z64 ? Buffer.concat([Buffer.from([1, 0, 16, 0]), u64(size), u64(csize)]) : Buffer.alloc(0);
  const h = Buffer.alloc(30);
  h.writeUInt32LE(0x04034b50, 0); h.writeUInt16LE(z64 ? 45 : 20, 4); h.writeUInt16LE(flags, 6); h.writeUInt16LE(method, 8);
  h.writeUInt16LE(time, 10); h.writeUInt16LE(date, 12); h.writeUInt32LE(crc >>> 0, 14);
  h.writeUInt32LE(z64 ? MAX32 : csize, 18); h.writeUInt32LE(z64 ? MAX32 : size, 22);
  h.writeUInt16LE(nameBuf.length, 26); h.writeUInt16LE(extra.length, 28);
  return Buffer.concat([h, nameBuf, extra]);
}
/** central directory header; ZIP64 extra holds exactly the fields that overflow (size, csize, offset – in that order) */
function centralHeader({ nameBuf, flags, method, time, date, crc, csize, size, attr = 0 }, offset) {
  const parts = [];
  if (size >= MAX32) parts.push(u64(size));
  if (csize >= MAX32) parts.push(u64(csize));
  if (offset >= MAX32) parts.push(u64(offset));
  const body = Buffer.concat(parts);
  const extra = parts.length ? Buffer.concat([Buffer.from([1, 0, body.length, 0]), body]) : Buffer.alloc(0);
  const c = Buffer.alloc(46);
  c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(parts.length ? 45 : 20, 4); c.writeUInt16LE(parts.length ? 45 : 20, 6); c.writeUInt16LE(flags, 8); c.writeUInt16LE(method, 10);
  c.writeUInt16LE(time, 12); c.writeUInt16LE(date, 14); c.writeUInt32LE(crc >>> 0, 16);
  c.writeUInt32LE(csize >= MAX32 ? MAX32 : csize, 20); c.writeUInt32LE(size >= MAX32 ? MAX32 : size, 24);
  c.writeUInt16LE(nameBuf.length, 28); c.writeUInt16LE(extra.length, 30); c.writeUInt32LE(attr >>> 0, 38); c.writeUInt32LE(offset >= MAX32 ? MAX32 : offset, 42);
  return Buffer.concat([c, nameBuf, extra]);
}
/** end of central directory; with ZIP64 end record + locator before it when count/size/offset need it */
function endRecords(count, cdSize, cdStart) {
  const z64 = count >= MAX16 || cdSize >= MAX32 || cdStart >= MAX32;
  const e = Buffer.alloc(22);
  e.writeUInt32LE(0x06054b50, 0);
  e.writeUInt16LE(z64 ? MAX16 : count, 8); e.writeUInt16LE(z64 ? MAX16 : count, 10);
  e.writeUInt32LE(z64 ? MAX32 : cdSize, 12); e.writeUInt32LE(z64 ? MAX32 : cdStart, 16);
  if (!z64) return e;
  const r = Buffer.alloc(56);
  r.writeUInt32LE(0x06064b50, 0); r.writeBigUInt64LE(44n, 4); r.writeUInt16LE(45, 12); r.writeUInt16LE(45, 14);
  r.writeBigUInt64LE(BigInt(count), 24); r.writeBigUInt64LE(BigInt(count), 32); r.writeBigUInt64LE(BigInt(cdSize), 40); r.writeBigUInt64LE(BigInt(cdStart), 48);
  const loc = Buffer.alloc(20);
  loc.writeUInt32LE(0x07064b50, 0); loc.writeBigUInt64LE(BigInt(cdStart + cdSize), 8); loc.writeUInt32LE(1, 16);
  return Buffer.concat([r, loc, e]);
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
    const e = { nameBuf, flags: 0x0800, method, time, date, crc, csize: body.length, size, attr: isDir ? 0x10 : 0 };
    const offset = this.offset;
    this._write(localHeader(e));
    const r = body.length ? this._write(body) : true;
    this.central.push(centralHeader(e, offset));
    return r;
  }
  /** Write the central directory. Without a sink → the whole archive as Buffer. */
  end() {
    const start = this.offset;
    for (const c of this.central) this._write(c);
    this._write(endRecords(this.central.length, this.offset - start, start));
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
 * The central directory (classic or ZIP64) → entries with the start of their data.
 * Throws on a broken directory; local headers are checked.
 */
function centralDir(buf) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  const min = Math.max(0, buf.length - 65557);
  let eocd = -1;
  for (let i = buf.length - 22; i >= min; i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Not a zip file');
  let count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  // ZIP64: locator right before the end record → the ZIP64 end record with the real values
  if (eocd >= 20 && buf.readUInt32LE(eocd - 20) === 0x07064b50) {
    const at = r64(buf, eocd - 12);
    if (at + 56 > buf.length || buf.readUInt32LE(at) !== 0x06064b50) throw new Error('Broken zip directory');
    count = r64(buf, at + 32); p = r64(buf, at + 48);
  } else if (p === MAX32 || count === MAX16) {
    if (p === MAX32) throw new Error('Broken zip directory');
  }
  const out = [];
  for (let i = 0; i < count; i++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) throw new Error('Broken zip directory');
    const flags = buf.readUInt16LE(p + 8); const method = buf.readUInt16LE(p + 10);
    const time = buf.readUInt16LE(p + 12); const date = buf.readUInt16LE(p + 14);
    const crc = buf.readUInt32LE(p + 16); let csize = buf.readUInt32LE(p + 20); let size = buf.readUInt32LE(p + 24);
    const nl = buf.readUInt16LE(p + 28); const xl = buf.readUInt16LE(p + 30); const cl = buf.readUInt16LE(p + 32);
    const attr = buf.readUInt32LE(p + 38); let off = buf.readUInt32LE(p + 42);
    const nameBuf = buf.subarray(p + 46, p + 46 + nl);
    const name = nameBuf.toString(flags & 0x0800 ? 'utf8' : 'latin1');
    if (size === MAX32 || csize === MAX32 || off === MAX32) {
      // ZIP64 extended information (0x0001): only the fields that are 0xFFFFFFFF, in this order
      let x = p + 46 + nl; const xe = x + xl;
      while (x + 4 <= xe) {
        const id = buf.readUInt16LE(x); const len = buf.readUInt16LE(x + 2);
        if (id === 1) {
          let q = x + 4;
          if (size === MAX32 && q + 8 <= x + 4 + len) { size = r64(buf, q); q += 8; }
          if (csize === MAX32 && q + 8 <= x + 4 + len) { csize = r64(buf, q); q += 8; }
          if (off === MAX32 && q + 8 <= x + 4 + len) { off = r64(buf, q); q += 8; }
          break;
        }
        x += 4 + len;
      }
    }
    p += 46 + nl + xl + cl;
    if (off + 30 > buf.length || buf.readUInt32LE(off) !== 0x04034b50) throw new Error(`Broken zip entry ${name}`);
    const start = off + 30 + buf.readUInt16LE(off + 26) + buf.readUInt16LE(off + 28);
    if (start + csize > buf.length) throw new Error(`Broken zip entry ${name}`);
    out.push({ name, nameBuf, flags, method, time, date, crc, csize, size, attr, off, start });
  }
  return out;
}

/**
 * Read a zip from a Buffer. → Map(name → { name, size, csize, method, dir, read(): Buffer })
 * maxSize: cap per entry (unpacked) – throws beyond. Names are taken as UTF-8 (flag 11) or latin1.
 */
function readZip(buf, { maxSize = 256 * 1048576 } = {}) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  const out = new Map();
  for (const e of centralDir(buf)) {
    const { name, flags, method, csize, size, start } = e;
    const read = () => {
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

module.exports = { ZipWriter, makeZip, readZip, centralDir, localHeader, centralHeader, endRecords, crc32, isZip, cleanName, MAX32 };
