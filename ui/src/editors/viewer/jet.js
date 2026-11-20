/*
 * Read-only reader for Microsoft Access databases without Access or the Office engine: Jet 3 (Access 97 .mdb),
 * Jet 4 (Access 2000–2003 .mdb) and ACE (.accdb, Access 2007+). Pure, runs in Node tests.
 *
 *   readJet(bytes, { maxRows }) → { version, format, tables: [{ name, columns: [{ name, type }], rows, rowCount, truncated }],
 *                                  queries, forms, reports, macros, modules, linked }
 *
 * What it does: the catalog (MSysObjects, table definition page 2), each user table's definition page (columns:
 * type, fixed offset, variable index, length, flags; column names) and every data page that belongs to the table
 * (rows with fixed + variable columns, null mask, Jet 4 Unicode compression, memo fields stored inline or in long
 * value pages). Not supported (shown as such): encrypted / password-protected databases, OLE objects and
 * attachments (shown as their size), linked tables (their source is named), query SQL.
 */

const T = { BOOL: 1, BYTE: 2, INT: 3, LONG: 4, MONEY: 5, FLOAT: 6, DOUBLE: 7, DATETIME: 8, BINARY: 9, TEXT: 10, OLE: 11, MEMO: 12, GUID: 15, NUMERIC: 16, COMPLEX: 18, BIGINT: 19 };
export const JET_TYPES = { 1: 'Yes/No', 2: 'Byte', 3: 'Integer', 4: 'Long', 5: 'Currency', 6: 'Single precision', 7: 'Double precision', 8: 'Date/Time', 9: 'Binary', 10: 'Text', 11: 'OLE object', 12: 'Memo', 15: 'GUID', 16: 'Decimal', 18: 'Attachment', 19: 'Large number' };

const F3 = { page: 2048, rowsOff: 8, ncolsOff: 25, nridxOff: 31, colsStart: 43, ridxSize: 8, colSize: 18, cFlags: 13, cLen: 16, cNum: 1, cVar: 3, cFixed: 14, cRowNum: 5, nameLen1: true };
const F4 = { page: 4096, rowsOff: 12, ncolsOff: 45, nridxOff: 51, colsStart: 63, ridxSize: 12, colSize: 25, cFlags: 15, cLen: 23, cNum: 5, cVar: 7, cFixed: 21, cRowNum: 9, nameLen1: false };

const u16 = (b, o) => b[o] | (b[o + 1] << 8);
const i16 = (b, o) => { const v = u16(b, o); return v & 0x8000 ? v - 0x10000 : v; };
const u24 = (b, o) => b[o] | (b[o + 1] << 8) | (b[o + 2] << 16);
const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
const i32 = (b, o) => b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24);

function latin1(b) { let s = ''; for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i] < 0x80 || b[i] >= 0xa0 ? b[i] : CP1252[b[i] - 0x80] || b[i]); return s; }
const CP1252 = [0x20ac, 0, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0, 0x17d, 0, 0, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0, 0x17e, 0x178];
function ucs2(b) { let s = ''; for (let i = 0; i + 1 < b.length; i += 2) s += String.fromCharCode(u16(b, i)); return s; }
/** Jet 4 text: UTF-16LE, or "compressed" (FF FE prefix, 0x00 switches between 1-byte and 2-byte chars) */
function jet4Text(b) {
  if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) {
    let s = ''; let compressed = true; let i = 2;
    while (i < b.length) {
      if (b[i] === 0) { compressed = !compressed; i++; continue; }
      if (compressed) { s += String.fromCharCode(b[i]); i++; } else { if (i + 1 >= b.length) break; s += String.fromCharCode(u16(b, i)); i += 2; }
    }
    return s;
  }
  return ucs2(b);
}

export function readJet(input, { maxRows = 5000, maxPages = 200000 } = {}) {
  const b = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (b.length < 2048 || b[0] !== 0 || b[1] !== 1 || b[2] !== 0 || b[3] !== 0) throw new Error('Not an Access database');
  const sig = latin1(b.subarray(4, 19));
  if (!/Standard (Jet|ACE) DB/.test(sig)) throw new Error('Not an Access database');
  const ver = b[0x14];
  const jet3 = ver === 0;
  const F = jet3 ? F3 : F4;
  const P = F.page;
  const pages = Math.min(Math.floor(b.length / P), maxPages);
  const text = jet3 ? latin1 : jet4Text;
  const page = (n) => (n >= 0 && n < pages ? b.subarray(n * P, (n + 1) * P) : null);

  /* ---- table definition (may continue over several pages) ---- */
  function tdef(n) {
    const pg = page(n); if (!pg || pg[0] !== 2) return null;
    const parts = [];
    let cur = n; let guard = 0;
    while (cur && guard++ < 1000) { const p = page(cur); if (!p || p[0] !== 2) break; parts.push(parts.length ? p.subarray(8) : p); cur = u32(p, 4); }
    const d = new Uint8Array(parts.reduce((s, p) => s + p.length, 0)); let o = 0; for (const p of parts) { d.set(p, o); o += p.length; }
    const numCols = u16(d, F.ncolsOff);
    const numRidx = u32(d, F.nridxOff);
    let p = F.colsStart + numRidx * F.ridxSize;
    const cols = [];
    for (let i = 0; i < numCols; i++) {
      const c = d.subarray(p, p + F.colSize);
      cols.push({ type: c[0], num: u16(c, F.cNum), varIdx: u16(c, F.cVar), rowNum: u16(c, F.cRowNum), flags: c[F.cFlags], fixedOff: u16(c, F.cFixed), len: u16(c, F.cLen) });
      p += F.colSize;
    }
    for (const c of cols) {
      if (F.nameLen1) { const l = d[p]; c.name = latin1(d.subarray(p + 1, p + 1 + l)); p += 1 + l; } else { const l = u16(d, p); c.name = ucs2(d.subarray(p + 2, p + 2 + l)); p += 2 + l; }
    }
    return { page: n, cols, rowCount: u32(d, F.rowsOff + 4 - 4 + (jet3 ? 4 : 4)), fixed: (c) => !!(c.flags & 1) };
  }

  /* ---- rows ---- */
  function rowBounds(pg) {
    const nOff = jet3 ? 8 : 12;
    const n = u16(pg, nOff);
    const out = [];
    for (let i = 0; i < n && i < 1000; i++) {
      const raw = u16(pg, nOff + 2 + i * 2);
      const start = raw & 0x1fff;
      const end = i === 0 ? P - 1 : (u16(pg, nOff + 2 + (i - 1) * 2) & 0x1fff) - 1;
      out.push({ start, end, deleted: !!(raw & 0x8000), lookup: !!(raw & 0x4000) });
    }
    return out;
  }
  function lval(ptr) { // row pointer: row (1 byte) + page (3 bytes)
    const rn = ptr & 0xff; const pn = ptr >>> 8;
    const pg = page(pn); if (!pg || pg[0] !== 1) return null;
    const rb = rowBounds(pg)[rn]; if (!rb) return null;
    return pg.subarray(rb.start, rb.end + 1);
  }
  function memo(d) {
    if (d.length < 12) return null;
    const head = u32(d, 0); const len = head & 0x3fffffff;
    if (head & 0x80000000) return d.subarray(12, 12 + len);
    if (head & 0x40000000) { const r = lval(u32(d, 4)); return r ? r.subarray(0, len) : null; }
    // chain of long value pages: each row starts with the pointer to the next one
    const parts = []; let got = 0; let ptr = u32(d, 4); let guard = 0;
    while (ptr && got < len && guard++ < 10000) { const r = lval(ptr); if (!r) break; const piece = r.subarray(4); parts.push(piece); got += piece.length; ptr = u32(r, 0); }
    const out = new Uint8Array(Math.min(got, len)); let o = 0; for (const x of parts) { const take = Math.min(x.length, out.length - o); out.set(x.subarray(0, take), o); o += take; }
    return out;
  }
  function value(c, d) {
    switch (c.type) {
      case T.BYTE: return d[0];
      case T.INT: return i16(d, 0);
      case T.LONG: case T.COMPLEX: return i32(d, 0);
      case T.MONEY: { const lo = u32(d, 0); const hi = i32(d, 4); return Math.round((hi * 4294967296 + lo)) / 10000; }
      case T.FLOAT: return Math.round(new DataView(d.buffer, d.byteOffset, 4).getFloat32(0, true) * 1e6) / 1e6;
      case T.DOUBLE: return new DataView(d.buffer, d.byteOffset, 8).getFloat64(0, true);
      case T.BIGINT: return Number(new DataView(d.buffer, d.byteOffset, 8).getBigInt64(0, true));
      case T.DATETIME: { const v = new DataView(d.buffer, d.byteOffset, 8).getFloat64(0, true); const ms = Math.round((v - 25569) * 86400000); return new Date(ms); }
      case T.TEXT: return text(d);
      case T.MEMO: { const m = memo(d); return m ? text(m) : ''; }
      case T.GUID: { const h = (o, n, rev) => { const x = Array.from(d.subarray(o, o + n), (v) => v.toString(16).padStart(2, '0')); return (rev ? x.reverse() : x).join(''); }; return `{${h(0, 4, true)}-${h(4, 2, true)}-${h(6, 2, true)}-${h(8, 2)}-${h(10, 6)}}`.toUpperCase(); }
      case T.NUMERIC: { // 17 bytes: precision, scale, sign, 16 bytes magnitude (4 little-endian 32-bit words, most significant first)
        const scale = d[1]; const neg = d[2] === 0x80; let v = 0n; for (let k = 0; k < 4; k++) v = (v << 32n) | BigInt(u32(d, 3 + k * 4)); const s = v.toString().padStart(scale + 1, '0'); const r = scale ? `${s.slice(0, -scale)}.${s.slice(-scale)}` : s; return Number(neg ? `-${r}` : r);
      }
      case T.OLE: case T.BINARY: { const m = c.type === T.OLE ? memo(d) : d; return { binary: m ? m.length : d.length }; }
      default: return { binary: d.length };
    }
  }
  function rowValues(def, pg, rb) {
    const r = pg.subarray(rb.start, rb.end + 1);
    const len = r.length; if (len < 2) return null;
    const ncols = jet3 ? r[0] : u16(r, 0);
    const mask = Math.ceil(ncols / 8);
    const nullAt = (k) => { const byte = r[len - mask + (k >> 3)]; return !(byte & (1 << (k & 7))); };
    let varCount; const varOff = [];
    if (jet3) {
      varCount = r[len - mask - 1];
      let jumps = Math.floor(len / 256); const colPtr = len - mask - jumps - 2;
      if (Math.floor((colPtr - varCount) / 256) < jumps) jumps--;
      let used = 0;
      for (let i = 0; i <= varCount; i++) {
        while (used < jumps && i === r[len - mask - used - 2]) used++;
        varOff.push(r[colPtr - i] + used * 256);
      }
    } else {
      varCount = u16(r, len - mask - 2);
      for (let i = 0; i <= varCount; i++) varOff.push(u16(r, len - mask - 4 - i * 2));
    }
    const pre = jet3 ? 1 : 2;
    return def.cols.map((c) => {
      if (c.type === T.BOOL) return c.rowNum < ncols ? !nullAt(c.rowNum) : null;
      if (c.rowNum >= ncols || nullAt(c.rowNum)) return null;
      if (c.flags & 1) { const s = pre + c.fixedOff; return value(c, r.subarray(s, s + c.len)); }
      if (c.varIdx >= varCount) return null;
      const s = varOff[c.varIdx]; const e = varOff[c.varIdx + 1];
      if (!(e >= s) || e > len) return null;
      return value(c, r.subarray(s, e));
    });
  }
  /** all rows of a table: every data page owned by its definition page */
  function rows(def, limit) {
    const out = []; let total = 0;
    for (let n = 1; n < pages; n++) {
      const pg = page(n);
      if (pg[0] !== 1 || u32(pg, 4) !== def.page) continue;
      for (const rb of rowBounds(pg)) {
        if (rb.deleted || rb.lookup || rb.end < rb.start) continue;
        total++;
        if (out.length >= limit) continue;
        try { const v = rowValues(def, pg, rb); if (v) out.push(v); } catch { /* damaged row */ }
      }
    }
    return { rows: out, total };
  }

  // catalog
  const cat = tdef(2);
  if (!cat) throw new Error('The database catalog could not be read (encrypted or damaged database)');
  const ci = (name) => cat.cols.findIndex((c) => c.name.toLowerCase() === name);
  const iName = ci('name'); const iType = ci('type'); const iFlags = ci('flags'); const iId = ci('id'); const iDb = ci('database'); const iForeign = ci('foreignname');
  const objs = rows(cat, 100000).rows;
  const res = { version: ver, format: jet3 ? 'Jet 3 (Access 97)' : ver === 1 ? 'Jet 4 (Access 2000–2003)' : 'ACE (Access 2007+)', tables: [], queries: [], forms: [], reports: [], macros: [], modules: [], linked: [] };
  for (const o of objs) {
    const name = o[iName]; const type = o[iType]; const flags = o[iFlags] || 0;
    if (typeof name !== 'string' || /^(MSys|USys|~)/i.test(name)) continue;
    if (type === 1 && !(flags & 0x80000002)) res.tables.push({ name, id: (o[iId] || 0) & 0xffffff });
    else if (type === 5) res.queries.push(name);
    else if (type === 4 || type === 6) res.linked.push({ name, source: [o[iDb], o[iForeign]].filter((x) => typeof x === 'string' && x).join(' · ') });
    else if (type === -32768) res.forms.push(name);
    else if (type === -32764) res.reports.push(name);
    else if (type === -32766) res.macros.push(name);
    else if (type === -32761) res.modules.push(name);
  }
  res.tables = res.tables.sort((a, x) => a.name.localeCompare(x.name)).map((t) => {
    const def = tdef(t.id);
    if (!def) return { name: t.name, columns: [], rows: [], rowCount: 0, error: true };
    const r = rows(def, maxRows);
    return { name: t.name, columns: def.cols.map((c) => ({ name: c.name, type: c.type })), rows: r.rows, rowCount: r.total, truncated: r.total > r.rows.length };
  });
  for (const k of ['queries', 'forms', 'reports', 'macros', 'modules']) res[k].sort((a, x) => a.localeCompare(x));
  return res;
}

/** a cell for display */
export function jetText(v) {
  if (v == null) return '';
  if (v instanceof Date) return Number.isNaN(+v) ? '' : v.toISOString().replace('T', ' ').replace(/(\s00:00:00)?\.000Z$/, '').replace(/Z$/, '');
  if (typeof v === 'boolean') return v ? '✓' : '–';
  if (typeof v === 'object' && v.binary != null) return `‹${v.binary} B›`;
  return String(v);
}
