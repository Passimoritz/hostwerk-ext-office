'use strict';
/*
 * Co-editing relay (Yjs sync protocol + awareness, y-protocols framing) – one room per file.
 * The server keeps the document as the log of all Yjs updates (Yjs updates are commutative and idempotent, so replaying
 * the log gives the full state). The log is compacted now and then with the small Yjs build shipped in server/vendor/
 * (yjs.cjs, MIT): all updates are applied to a fresh Y.Doc (garbage collection on – deleted text keeps no content) and
 * encoded again as ONE update. Only a room whose compacted state is still over the cap asks the leader to save and
 * everyone to reload ("reset").
 *
 * Messages (binary, lib0 varUint framing):
 *   [0, 0, sv]  sync step 1 · [0, 1, update] sync step 2 · [0, 2, update] update
 *   [1, awarenessUpdate] awareness · [3] query awareness
 *   [100, jsonString] control (ours): server → client { type: 'hello', id, leader, peers, user }, { type: 'leader', leader },
 *     { type: 'peers', peers: [{ id, user, leader }] }, { type: 'saved', mtime, by }, { type: 'reset', reason },
 *     { type: 'error', error }; client → server { type: 'saved', mtime } (leader only)
 * The leader (first client of a room, then the oldest one) saves the file through the normal saver (PUT /office/raw).
 */
const MSG_SYNC = 0; const MSG_AWARENESS = 1; const MSG_QUERY = 3; const MSG_CONTROL = 100;
const LIMITS = { clients: 24, rooms: 200, logBytes: 48 * 1048576, message: 16 * 1048576, emptyGraceMs: 120000 };

/* ---------- lib0 encoding ---------- */
function readVarUint(buf, pos) { let n = 0; let mult = 1; for (let i = 0; i < 8; i++) { if (pos.p >= buf.length) throw new Error('short'); const b = buf[pos.p++]; n += (b & 127) * mult; if (b < 128) return n; mult *= 128; } throw new Error('varint too long'); }
function writeVarUint(n) { const out = []; let v = n; while (v > 127) { out.push((v % 128) | 128); v = Math.floor(v / 128); } out.push(v); return Buffer.from(out); }
function readBytes(buf, pos) { const len = readVarUint(buf, pos); if (pos.p + len > buf.length) throw new Error('short'); const b = buf.subarray(pos.p, pos.p + len); pos.p += len; return b; }
const varBytes = (b) => Buffer.concat([writeVarUint(b.length), b]);
const readString = (buf, pos) => readBytes(buf, pos).toString('utf8');
const varString = (s) => varBytes(Buffer.from(String(s), 'utf8'));

const syncMsg = (type, data) => Buffer.concat([writeVarUint(MSG_SYNC), writeVarUint(type), varBytes(data)]);
const controlMsg = (o) => Buffer.concat([writeVarUint(MSG_CONTROL), varString(JSON.stringify(o))]);
/** awareness update → [{ client, clock, state (JSON string) }] */
function parseAwareness(u) { const pos = { p: 0 }; const n = readVarUint(u, pos); const out = []; for (let i = 0; i < n && i < 1000; i++) out.push({ client: readVarUint(u, pos), clock: readVarUint(u, pos), state: readString(u, pos) }); return out; }
function encodeAwareness(list) { return Buffer.concat([writeVarUint(list.length), ...list.flatMap((e) => [writeVarUint(e.client), writeVarUint(e.clock), varString(e.state)])]); }
const awarenessMsg = (list) => Buffer.concat([writeVarUint(MSG_AWARENESS), varBytes(encodeAwareness(list))]);

// a development checkout has yjs in node_modules (the same copy the tests' clients use – two copies in one process warn);
// the package uses its own small build
let Y = null;
try { Y = require('yjs'); } catch {
  // our own private copy: Yjs warns ("already imported") when another copy runs in the same process (Hostwerk or another
  // extension). The copies never share objects here, so the marker is set aside while loading and put back after.
  const MARK = '__ $YJS$ __'; const had = Object.prototype.hasOwnProperty.call(globalThis, MARK); const before = globalThis[MARK];
  try { delete globalThis[MARK]; Y = require('./vendor/yjs.cjs'); } catch { /* no compaction */ } finally { if (had) globalThis[MARK] = before; else delete globalThis[MARK]; }
}

class Room {
  constructor(key, opts) { this.key = key; this.opts = opts; this.peers = new Map(); this.log = []; this.bytes = 0; this.compacted = 0; this.compactions = 0; this.aw = new Map(); this.seq = 0; this.emptySince = null; this.mtime = null; }
  get leader() { for (const p of this.peers.values()) return p; return null; }
  send(p, buf) { try { p.send(buf); } catch { /* closed */ } }
  others(from, buf) { for (const p of this.peers.values()) if (p !== from) this.send(p, buf); }
  peerList() { const l = this.leader; return [...this.peers.values()].map((p) => ({ id: p.id, user: p.user, leader: p === l })); }
  addUpdate(u) {
    if (!u.length || (u.length === 2 && u[0] === 0 && u[1] === 0)) return; // empty update
    this.log.push(Buffer.from(u)); this.bytes += u.length;
    // compact after many small updates or when the log doubled since the last compaction (+ 1 MB)
    if (Y && (this.log.length > 300 || this.bytes > this.compacted * 2 + 1048576)) this.compact();
    if (this.bytes > this.opts.logBytes) {
      if (Y) this.compact();
      if (this.bytes > this.opts.logBytes) this.reset('too-large');
    }
  }
  compact() {
    if (!Y || this.log.length < 2) { this.compacted = this.bytes; return; }
    try {
      const doc = new Y.Doc(); // gc on: deleted content becomes small "deleted" structs
      for (const u of this.log) Y.applyUpdate(doc, new Uint8Array(u));
      const m = Buffer.from(Y.encodeStateAsUpdate(doc)); // pending (not yet integrable) structs are kept in it
      doc.destroy();
      this.log = [m]; this.bytes = m.length; this.compactions++;
    } catch {
      try { const m = Buffer.from(Y.mergeUpdates(this.log.map((b) => new Uint8Array(b)))); this.log = [m]; this.bytes = m.length; } catch { /* keep the log */ }
    }
    this.compacted = this.bytes;
  }
  reset(reason) { const l = this.leader; for (const p of this.peers.values()) this.send(p, controlMsg({ type: 'reset', reason, leader: p === l })); this.log = []; this.bytes = 0; }
  join(p) {
    this.emptySince = null;
    const wasEmpty = !this.peers.size;
    this.peers.set(p.id, p);
    p.clients = new Set();
    this.send(p, controlMsg({ type: 'hello', id: p.id, leader: this.leader === p, peers: this.peers.size, user: p.user, fresh: wasEmpty && !this.log.length, mtime: this.mtime }));
    // ask for the client's state, then give it everything we have
    this.send(p, syncMsg(0, Buffer.from([0])));
    for (const u of this.log) this.send(p, syncMsg(1, u));
    if (this.aw.size) this.send(p, awarenessMsg([...this.aw.values()].filter((e) => e.state !== 'null')));
    this.others(p, controlMsg({ type: 'peers', peers: this.peerList() }));
  }
  message(p, data) {
    const buf = Buffer.isBuffer(data) ? data : Buffer.from(data);
    const pos = { p: 0 };
    const type = readVarUint(buf, pos);
    if (type === MSG_SYNC) {
      const st = readVarUint(buf, pos); const body = readBytes(buf, pos);
      if (st === 0) { if (!this.log.length) this.send(p, syncMsg(1, Buffer.from([0, 0]))); for (const u of this.log) this.send(p, syncMsg(1, u)); return; }
      if (st === 1 || st === 2) { this.addUpdate(body); this.others(p, syncMsg(2, body)); return; }
      return;
    }
    if (type === MSG_AWARENESS) {
      const u = readBytes(buf, pos);
      const list = parseAwareness(u);
      for (const e of list) { const cur = this.aw.get(e.client); if (!cur || e.clock >= cur.clock) { this.aw.set(e.client, e); p.clients.add(e.client); } }
      if (this.aw.size > 500) for (const k of [...this.aw.keys()].slice(0, this.aw.size - 500)) this.aw.delete(k);
      this.others(p, awarenessMsg(list));
      return;
    }
    if (type === MSG_QUERY) { if (this.aw.size) this.send(p, awarenessMsg([...this.aw.values()].filter((e) => e.state !== 'null'))); return; }
    if (type === MSG_CONTROL) {
      let o; try { o = JSON.parse(readString(buf, pos)); } catch { return; }
      if (o && o.type === 'saved' && this.leader === p) { this.mtime = Number(o.mtime) || null; this.others(p, controlMsg({ type: 'saved', mtime: this.mtime, by: p.user })); }
    }
  }
  leave(p) {
    const wasLeader = this.leader === p;
    this.peers.delete(p.id);
    // awareness: the client's cursors disappear for the others
    const gone = [...(p.clients || [])].map((c) => { const cur = this.aw.get(c); const e = { client: c, clock: (cur ? cur.clock : 0) + 1, state: 'null' }; this.aw.delete(c); return e; });
    if (gone.length) this.others(p, awarenessMsg(gone));
    const l = this.leader;
    if (wasLeader && l) this.send(l, controlMsg({ type: 'leader', leader: true }));
    if (this.peers.size) for (const q of this.peers.values()) this.send(q, controlMsg({ type: 'peers', peers: this.peerList() }));
    else this.emptySince = Date.now();
  }
}

class Collab {
  constructor(opts = {}) { this.opts = { ...LIMITS, ...opts }; this.rooms = new Map(); this.n = 0; }
  room(key) {
    let r = this.rooms.get(key);
    if (!r) {
      this.sweep();
      if (this.rooms.size >= this.opts.rooms) throw Object.assign(new Error('Too many co-editing sessions at once.'), { status: 503 });
      r = new Room(key, this.opts); this.rooms.set(key, r);
    }
    return r;
  }
  /** peer: { send(buf), close(code, reason) } → { id, message(data), close() } */
  connect(key, peer, { user = 'Hostwerk' } = {}) {
    const r = this.room(key);
    if (r.peers.size >= this.opts.clients) throw Object.assign(new Error('Too many people in this file at once.'), { status: 429 });
    const p = { id: `c${++this.n}`, user: String(user).slice(0, 60), send: (b) => peer.send(b) };
    r.join(p);
    return {
      id: p.id, room: r,
      message: (data) => { if (data.length > this.opts.message) { peer.close(1009, 'too big'); return; } try { r.message(p, data); } catch { peer.close(1003, 'bad message'); } },
      close: () => r.leave(p),
    };
  }
  /** drop rooms that stayed empty longer than the grace time (a reload keeps the room) */
  sweep(now = Date.now()) { for (const [k, r] of this.rooms) if (!r.peers.size && r.emptySince && now - r.emptySince > this.opts.emptyGraceMs) this.rooms.delete(k); }
  /** the file was saved / changed outside the room: tell everybody (they reload when no peer is the writer) */
  external(key, mtime) { const r = this.rooms.get(key); if (r && r.mtime !== mtime) { r.mtime = mtime; for (const p of r.peers.values()) r.send(p, controlMsg({ type: 'saved', mtime, by: null })); } }
  rename(a, b) { const r = this.rooms.get(a); if (r && !this.rooms.has(b)) { this.rooms.delete(a); r.key = b; this.rooms.set(b, r); } }
  stats() { return [...this.rooms.values()].map((r) => ({ key: r.key, peers: r.peers.size, bytes: r.bytes, updates: r.log.length, compactions: r.compactions })); }
}

module.exports = { Collab, Room, readVarUint, writeVarUint, varBytes, varString, syncMsg, controlMsg, awarenessMsg, parseAwareness, encodeAwareness, MSG_CONTROL, hasYjs: !!Y };
