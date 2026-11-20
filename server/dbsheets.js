'use strict';
/*
 * Database tables as sheets – through Hostwerk's own database layer (ctx.db = DbApi of the core: connections, cached
 * drivers, read-only switches). Nothing is concatenated from user input: table and column names are checked against
 * the real lists of the server and quoted, values always go as parameters.
 *
 *   new DbSheets(ctxDb)
 *   .servers() → [{ id, name, type, kind, engine, running, readOnly, builtin }]
 *   .databases(server) → [{ name, system }]
 *   .tables(server, db) → [{ name, view, rows }]
 *   .rows(server, db, table, { offset, limit, sort, dir, filter }) → { columns: [{ id, label, type, key }], rows, total, primaryKey, readOnly }
 *   .write(server, db, table, changes) → { ok, applied, results }   – one transaction (SQL), all or nothing
 */
const MAX_CHANGES = 500;
const fail = (status, msg, vars) => Object.assign(new Error(msg), { status, vars });
const okValue = (v) => v === null || ['string', 'number', 'boolean'].includes(typeof v);
const OPS = new Set(['=', '!=', '<', '>', '<=', '>=', 'contains', 'LIKE', 'NOT LIKE', 'IS NULL', 'IS NOT NULL']);

/** filter: JSON [{ col, op, value }] or "col op value" list → validated list */
function parseFilter(f) {
  let list = f;
  if (typeof f === 'string') { if (!f.trim()) return []; try { list = JSON.parse(f); } catch { throw fail(400, 'Invalid filter'); } }
  if (!Array.isArray(list)) list = list ? [list] : [];
  if (list.length > 20) throw fail(400, 'Too many filters');
  return list.map((x) => {
    const op = String((x && x.op) || '=');
    if (!OPS.has(op)) throw fail(400, 'Invalid filter');
    if (x.value !== undefined && !okValue(x.value)) throw fail(400, 'Invalid filter');
    return { col: String(x.col || ''), op, value: x.value == null ? x.value : String(x.value).slice(0, 1000) };
  });
}

const quote = {
  mysql: (n) => `\`${String(n).replace(/`/g, '``')}\``,
  postgres: (n) => String(n).split('.').map((p) => `"${p.replace(/"/g, '""')}"`).join('.'),
  sqlite: (n) => `"${String(n).replace(/"/g, '""')}"`,
};
const pcol = (n) => `"${String(n).replace(/"/g, '""')}"`;

class DbSheets {
  constructor(db, { lib = null } = {}) { this.db = db; this.lib = lib; }
  api() { if (!this.db) throw fail(501, 'Hostwerk\'s database studio is not available.'); return this.db; }
  servers() {
    return (this.api().connections() || []).filter((c) => c.kind === 'sql' || c.kind === 'mongo').map((c) => ({ id: c.id, name: c.name || c.id, type: c.type, kind: c.kind, dialect: c.dialect || null, engine: c.engine || c.type, running: c.running !== false, readOnly: !!c.readOnly, builtin: !!c.builtin }));
  }
  conn(server) {
    const s = String(server || '');
    if (!/^[\w:.-]{1,80}$/.test(s)) throw fail(400, 'Invalid server');
    const c = this.servers().find((x) => x.id === s);
    if (!c) throw fail(404, 'Database server not found');
    if (!c.running) throw fail(409, 'The database server {name} is not running.', { name: c.name });
    const { d } = this.api().driver(c.id);
    return { c, d };
  }
  async databases(server) {
    const { c, d } = this.conn(server);
    const list = await d.databases();
    return (list || []).map((x) => (typeof x === 'string' ? { name: x, system: false } : { name: x.name, system: !!x.system })).filter((x) => c.type !== 'sqlite' || x.name === 'main');
  }
  async checkDb(server, db) { const list = await this.databases(server); const n = String(db || (list[0] && list[0].name) || ''); if (!list.some((x) => x.name === n)) throw fail(404, 'Database not found'); return n; }
  async tables(server, db) {
    const { c, d } = this.conn(server);
    const name = await this.checkDb(server, db);
    if (c.kind === 'mongo') return (await d.collections(name)).map((x) => ({ name: x.name, view: !!x.view, rows: x.count != null ? x.count : null }));
    return (await d.tables(name)).map((t) => ({ name: t.name, view: !!t.view || /view/i.test(t.type || ''), rows: t.rowsApprox != null ? Number(t.rowsApprox) : null }));
  }
  async checkTable(server, db, table) {
    const list = await this.tables(server, db);
    const t = list.find((x) => x.name === String(table || ''));
    if (!t) throw fail(404, 'Table not found');
    return t;
  }

  async rows(server, db0, table, { offset = 0, limit = 100, sort = '', dir = 'asc', filter = '' } = {}) {
    const { c, d } = this.conn(server);
    const db = await this.checkDb(server, db0);
    const t = await this.checkTable(server, db, table);
    const lim = Math.min(Math.max(Number(limit) || 100, 1), 1000);
    const off = Math.max(Number(offset) || 0, 0);
    const filters = parseFilter(filter);
    if (c.kind === 'mongo') return this.mongoRows(c, d, db, t.name, { off, lim, sort, dir, filters });
    // the core's table browser pages by page number – offset must be a multiple of the page size for exact paging
    const size = lim; const page = Math.floor(off / size) + 1;
    const r = await d.tableRows(db, t.name, { page, size, sort: String(sort || ''), dir: dir === 'desc' ? 'desc' : 'asc', filters });
    const skip = off - (page - 1) * size;
    let rows = r.rows.slice(skip);
    if (skip && rows.length < lim) { const more = await d.tableRows(db, t.name, { page: page + 1, size, sort: String(sort || ''), dir: dir === 'desc' ? 'desc' : 'asc', filters }); rows = rows.concat(more.rows.slice(0, lim - rows.length)); }
    const pk = r.pk || [];
    return { columns: r.columns.map((x) => ({ id: x.name, label: x.name, type: x.dataType || x.type || null, key: x.key || null, nullable: !!x.nullable })), rows, total: r.total, primaryKey: pk, readOnly: c.readOnly || !pk.length || t.view, offset: off, limit: lim };
  }

  async mongoRows(c, d, db, coll, { off, lim, sort, dir, filters }) {
    const f = {};
    for (const x of filters) {
      if (!/^[\w.-]{1,100}$/.test(x.col) || x.col.startsWith('$')) throw fail(400, 'Invalid filter');
      if (x.op === '=') f[x.col] = x.value;
      else if (x.op === 'contains') f[x.col] = { $regex: String(x.value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), $options: 'i' };
      else if (x.op === 'IS NULL') f[x.col] = null;
      else if (x.op === 'IS NOT NULL') f[x.col] = { $ne: null };
      else if (['<', '>', '<=', '>=', '!='].includes(x.op)) f[x.col] = { [{ '<': '$lt', '>': '$gt', '<=': '$lte', '>=': '$gte', '!=': '$ne' }[x.op]]: Number.isFinite(Number(x.value)) ? Number(x.value) : x.value };
    }
    const s = sort && /^[\w.-]{1,100}$/.test(sort) ? { [sort]: dir === 'desc' ? -1 : 1 } : {};
    const r = await d.find(db, coll, { filter: JSON.stringify(f), sort: JSON.stringify(s), skip: off, limit: Math.min(lim, 500) });
    const keys = ['_id'];
    for (const doc of r.docs || []) for (const k of Object.keys(doc)) if (!keys.includes(k) && keys.length < 200) keys.push(k);
    const val = (v) => (v === undefined ? null : v !== null && typeof v === 'object' ? JSON.stringify(v) : v);
    return { columns: keys.map((k) => ({ id: k, label: k, type: null, key: k === '_id' ? 'PRI' : null })), rows: (r.docs || []).map((doc) => keys.map((k) => val(doc[k]))), total: r.total, primaryKey: ['_id'], readOnly: c.readOnly, offset: off, limit: lim, mongo: true };
  }

  /** changes: [{ op: 'update'|'insert'|'delete', key: { pkCol: value }, values: { col: value } }] */
  async write(server, db0, table, changes) {
    const { c, d } = this.conn(server);
    if (c.readOnly) throw fail(403, 'This database connection is read-only.');
    try { this.api().guard(c.id); } catch (e) { throw fail(403, 'This database connection is read-only.'); }
    const db = await this.checkDb(server, db0);
    const t = await this.checkTable(server, db, table);
    if (t.view) throw fail(409, 'Views cannot be changed.');
    if (!Array.isArray(changes) || !changes.length) throw fail(400, 'Nothing to save');
    if (changes.length > MAX_CHANGES) throw fail(413, 'At most {n} changes at once.', { n: MAX_CHANGES });
    if (c.kind === 'mongo') return this.mongoWrite(d, db, t.name, changes);
    const st = await d.structure(db, t.name);
    const cols = new Set((st.columns || []).map((x) => x.name));
    const pk = (st.columns || []).filter((x) => x.key === 'PRI').map((x) => x.name);
    if (!pk.length) throw fail(409, 'This table has no primary key – it can only be read.');
    const list = changes.map((ch, i) => {
      const op = String((ch && ch.op) || '');
      if (!['update', 'insert', 'delete'].includes(op)) throw fail(400, 'Invalid change #{n}', { n: i + 1 });
      const values = (ch && ch.values) || {};
      const key = (ch && ch.key) || {};
      for (const [k, v] of Object.entries(values)) { if (!cols.has(k)) throw fail(400, 'Unknown column {name}', { name: k }); if (!okValue(v)) throw fail(400, 'Invalid value in column {name}', { name: k }); }
      if (op !== 'insert') {
        if (Object.keys(key).length !== pk.length || !pk.every((k) => Object.prototype.hasOwnProperty.call(key, k))) throw fail(400, 'Every change needs the full primary key ({keys}).', { keys: pk.join(', ') });
        for (const v of Object.values(key)) if (!okValue(v) || v === null) throw fail(400, 'Invalid primary key');
      }
      if (op === 'update' && !Object.keys(values).length) throw fail(400, 'Invalid change #{n}', { n: i + 1 });
      return { op, key, values };
    });
    const dialect = c.type === 'sqlite' ? 'sqlite' : c.dialect === 'postgres' || c.type === 'postgres' ? 'postgres' : 'mysql';
    const Q = quote[dialect];
    const tbl = dialect === 'mysql' ? `${Q(db)}.${Q(t.name)}` : Q(t.name);
    const stmts = list.map(({ op, key, values }) => {
      const vc = Object.keys(values); const kc = Object.keys(key);
      const ph = dialect === 'postgres' ? (i) => `$${i}` : () => '?';
      if (op === 'insert') return vc.length ? { sql: `INSERT INTO ${tbl} (${vc.map(dialect === 'postgres' ? pcol : Q).join(', ')}) VALUES (${vc.map((_, i) => ph(i + 1)).join(', ')})`, params: vc.map((k) => values[k]) } : { sql: dialect === 'mysql' ? `INSERT INTO ${tbl} () VALUES ()` : `INSERT INTO ${tbl} DEFAULT VALUES`, params: [] };
      const qc = dialect === 'postgres' ? pcol : Q;
      const where = kc.map((k, i) => `${qc(k)} = ${ph(vc.length * (op === 'update' ? 1 : 0) + i + 1)}`).join(' AND ');
      if (op === 'delete') return { sql: `DELETE FROM ${tbl} WHERE ${where}${dialect === 'mysql' ? ' LIMIT 1' : ''}`, params: kc.map((k) => key[k]) };
      return { sql: `UPDATE ${tbl} SET ${vc.map((k, i) => `${qc(k)} = ${ph(i + 1)}`).join(', ')} WHERE ${where}${dialect === 'mysql' ? ' LIMIT 1' : ''}`, params: [...vc.map((k) => values[k]), ...kc.map((k) => key[k])] };
    });
    const results = [];
    const check = (n, i) => { if (list[i].op !== 'insert' && n !== 1) throw fail(409, 'Change #{n} did not find its row (changed by someone else?) – nothing was saved.', { n: i + 1 }); };
    if (dialect === 'mysql') {
      await d.withConn(db, async (conn) => {
        await conn.query('START TRANSACTION');
        try { for (let i = 0; i < stmts.length; i++) { const [r] = await conn.query(stmts[i].sql, stmts[i].params); check(r.affectedRows, i); results.push({ affected: r.affectedRows, insertId: r.insertId || null }); } await conn.query('COMMIT'); } catch (e) { await conn.query('ROLLBACK').catch(() => {}); throw e; }
      });
    } else if (dialect === 'postgres') {
      const client = await d.pool(db).connect();
      try {
        await client.query('BEGIN');
        try { for (let i = 0; i < stmts.length; i++) { const r = await client.query(stmts[i].sql, stmts[i].params); check(r.rowCount, i); results.push({ affected: r.rowCount }); } await client.query('COMMIT'); } catch (e) { await client.query('ROLLBACK').catch(() => {}); throw e; }
      } finally { client.release(); }
    } else {
      d.db.exec('BEGIN');
      try { for (let i = 0; i < stmts.length; i++) { const r = d.db.prepare(stmts[i].sql).run(...stmts[i].params); check(Number(r.changes), i); results.push({ affected: Number(r.changes), insertId: r.lastInsertRowid != null ? String(r.lastInsertRowid) : null }); } d.db.exec('COMMIT'); } catch (e) { try { d.db.exec('ROLLBACK'); } catch { /* egal */ } throw e; }
    }
    return { ok: true, applied: results.length, results };
  }

  async mongoWrite(d, db, coll, changes) {
    const EJSON = this.lib && this.lib.BSON && this.lib.BSON.EJSON;
    if (!EJSON) throw fail(501, 'MongoDB support is not available.');
    const col = d.db(db).collection(coll);
    const idOf = (k) => { const v = k && k._id; if (v == null) throw fail(400, 'Every change needs the full primary key ({keys}).', { keys: '_id' }); try { return EJSON.parse(typeof v === 'string' && /^\s*\{/.test(v) ? v : JSON.stringify(v), { relaxed: true }); } catch { return v; } };
    const clean = (values) => { const o = {}; for (const [k, v] of Object.entries(values || {})) { if (k.startsWith('$') || k.includes('\0') || k === '_id') throw fail(400, 'Unknown column {name}', { name: k }); if (!okValue(v)) throw fail(400, 'Invalid value in column {name}', { name: k }); o[k] = v; } return o; };
    const results = [];
    // MongoDB without a replica set has no transactions – changes are applied one by one (checked before)
    const prepared = changes.map((ch) => ({ op: ch.op, id: ch.op === 'insert' ? null : idOf(ch.key), values: clean(ch.values) }));
    for (const p of prepared) {
      if (p.op === 'insert') { const r = await col.insertOne(p.values); results.push({ insertId: String(r.insertedId) }); }
      else if (p.op === 'delete') results.push({ affected: (await col.deleteOne({ _id: p.id })).deletedCount });
      else if (p.op === 'update') results.push({ affected: (await col.updateOne({ _id: p.id }, { $set: p.values })).matchedCount });
      else throw fail(400, 'Invalid change');
    }
    return { ok: true, applied: results.length, results, atomic: false };
  }
}

module.exports = { DbSheets, parseFilter, quote };
