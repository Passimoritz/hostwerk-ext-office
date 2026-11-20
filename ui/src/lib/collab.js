/*
 * Co-editing provider – part of the shared library chunk ui/assets/lib-yjs.js (tools/build.js), so the shell and every
 * editor chunk use ONE Yjs instance. Protocol = y-websocket's (binary frames: 0 sync step 1/2/update, 1 awareness,
 * 3 query awareness) over /api/office/collab?path=…[&root=…]; the server relays and keeps a Y.Doc per open file.
 *
 *   const p = new OfficeProvider(wsUrl, doc, { user: { name, color, id } })
 *   p.awareness · p.synced (Promise) · p.status · p.on('status'|'synced'|'peers'|'control'|'saved'|'leader'|'reset', fn) · p.off(…) · p.destroy()
 *   p.leader (the server's leader = the client that saves) · p.hello ({ id, leader, peers, user, fresh, mtime } of the server)
 *   p.reportSaved(mtime) – the leader tells the room after its save (the others then know the new file time)
 */
import * as Y from 'yjs';
import * as syncProtocol from 'y-protocols/sync';
import * as awarenessProtocol from 'y-protocols/awareness';
import * as encoding from 'lib0/encoding';
import * as decoding from 'lib0/decoding';

const MSG_SYNC = 0;
const MSG_AWARENESS = 1;
const MSG_QUERY_AWARENESS = 3;
const MSG_CONTROL = 100; // Office's own control channel (server/collab.js): varString JSON

export class OfficeProvider {
  constructor(url, doc, { user = null, awareness = null, connect = true } = {}) {
    this.url = url; this.doc = doc;
    this.awareness = awareness || new awarenessProtocol.Awareness(doc);
    if (user) this.awareness.setLocalStateField('user', user);
    this.status = 'connecting';
    this.listeners = new Map();
    this.retries = 0;
    this.closed = false;
    this.leader = false;
    this.hello = null;
    this.synced = new Promise((res) => { this.resolveSynced = res; });
    this.onUpdate = (update, origin) => {
      if (origin === this) return;
      const e = encoding.createEncoder(); encoding.writeVarUint(e, MSG_SYNC); syncProtocol.writeUpdate(e, update); this.send(encoding.toUint8Array(e));
    };
    this.onAwareness = ({ added, updated, removed }) => {
      const changedClients = added.concat(updated, removed);
      const e = encoding.createEncoder(); encoding.writeVarUint(e, MSG_AWARENESS);
      encoding.writeVarUint8Array(e, awarenessProtocol.encodeAwarenessUpdate(this.awareness, changedClients));
      this.send(encoding.toUint8Array(e));
      this.emit('peers', this.peers());
    };
    this.onUnload = () => awarenessProtocol.removeAwarenessStates(this.awareness, [doc.clientID], 'window unload');
    doc.on('update', this.onUpdate);
    this.awareness.on('update', this.onAwareness);
    if (typeof window !== 'undefined') window.addEventListener('pagehide', this.onUnload);
    if (connect) this.connect();
  }
  on(evt, fn) { if (!this.listeners.has(evt)) this.listeners.set(evt, new Set()); this.listeners.get(evt).add(fn); return () => this.off(evt, fn); }
  off(evt, fn) { const s = this.listeners.get(evt); if (s) s.delete(fn); }
  emit(evt, v) { const s = this.listeners.get(evt); if (s) s.forEach((f) => { try { f(v); } catch { /* listener error */ } }); }
  setStatus(s) { if (this.status !== s) { this.status = s; this.emit('status', s); } }
  /** other people in this file: [{ clientId, user }] */
  peers() { const out = []; this.awareness.getStates().forEach((st, id) => { if (id !== this.doc.clientID && st && st.user) out.push({ clientId: id, user: st.user }); }); return out; }
  send(buf) { const ws = this.ws; if (ws && ws.readyState === 1) ws.send(buf); }
  connect() {
    if (this.closed) return;
    this.setStatus('connecting');
    const ws = new WebSocket(this.url);
    ws.binaryType = 'arraybuffer';
    this.ws = ws;
    ws.onopen = () => {
      this.retries = 0;
      this.setStatus('connected');
      const e = encoding.createEncoder(); encoding.writeVarUint(e, MSG_SYNC); syncProtocol.writeSyncStep1(e, this.doc); this.send(encoding.toUint8Array(e));
      if (this.awareness.getLocalState() !== null) {
        const a = encoding.createEncoder(); encoding.writeVarUint(a, MSG_AWARENESS); encoding.writeVarUint8Array(a, awarenessProtocol.encodeAwarenessUpdate(this.awareness, [this.doc.clientID])); this.send(encoding.toUint8Array(a));
      }
    };
    ws.onmessage = (ev) => {
      const d = decoding.createDecoder(new Uint8Array(ev.data));
      const e = encoding.createEncoder();
      const type = decoding.readVarUint(d);
      if (type === MSG_SYNC) {
        encoding.writeVarUint(e, MSG_SYNC);
        const st = syncProtocol.readSyncMessage(d, e, this.doc, this);
        if (st === syncProtocol.messageYjsSyncStep2 && this.resolveSynced) { this.resolveSynced(true); this.resolveSynced = null; this.emit('synced', true); }
        if (encoding.length(e) > 1) this.send(encoding.toUint8Array(e));
      } else if (type === MSG_AWARENESS) {
        awarenessProtocol.applyAwarenessUpdate(this.awareness, decoding.readVarUint8Array(d), this);
      } else if (type === MSG_CONTROL) {
        let m = null;
        try { m = JSON.parse(decoding.readVarString(d)); } catch { m = null; }
        if (m && typeof m === 'object') {
          if (m.type === 'hello') { this.hello = m; this.leader = !!m.leader; this.emit('leader', this.leader); }
          if (m.type === 'leader') { this.leader = !!m.leader; this.emit('leader', this.leader); }
          if (m.type === 'reset') { this.leader = !!m.leader; }
          this.emit('control', m);
          if (m.type === 'saved' || m.type === 'reset' || m.type === 'peers') this.emit(m.type, m);
        }
      } else if (type === MSG_QUERY_AWARENESS) {
        encoding.writeVarUint(e, MSG_AWARENESS);
        encoding.writeVarUint8Array(e, awarenessProtocol.encodeAwarenessUpdate(this.awareness, [...this.awareness.getStates().keys()]));
        this.send(encoding.toUint8Array(e));
      }
    };
    ws.onclose = () => {
      this.ws = null;
      awarenessProtocol.removeAwarenessStates(this.awareness, [...this.awareness.getStates().keys()].filter((id) => id !== this.doc.clientID), this);
      if (this.closed) return;
      this.setStatus('offline');
      // back off: 1 s, 2 s, 4 s … at most 30 s
      const ms = Math.min(30000, 1000 * 2 ** Math.min(5, this.retries++));
      this.timer = setTimeout(() => this.connect(), ms);
    };
    ws.onerror = () => { /* onclose follows */ };
  }
  reportSaved(mtime) { const e = encoding.createEncoder(); encoding.writeVarUint(e, MSG_CONTROL); encoding.writeVarString(e, JSON.stringify({ type: 'saved', mtime })); this.send(encoding.toUint8Array(e)); }
  destroy() {
    this.closed = true;
    clearTimeout(this.timer);
    this.onUnload();
    this.doc.off('update', this.onUpdate);
    this.awareness.off('update', this.onAwareness);
    if (typeof window !== 'undefined') window.removeEventListener('pagehide', this.onUnload);
    try { if (this.ws) this.ws.close(); } catch { /* closed */ }
    this.setStatus('offline');
    this.listeners.clear();
  }
}

export { Y, syncProtocol, awarenessProtocol };
export default { Y, OfficeProvider, syncProtocol, awarenessProtocol };
