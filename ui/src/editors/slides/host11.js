/*
 * Office 1.1 host features (ARCHITECTURE.md "Host object 1.1"): AI, live fields, phone remote, co-editing, versions,
 * picking a document. The members are always present on a 1.1 shell; their server routes may be missing (calls throw
 * with status 404/501) – everything here degrades to "not available" instead of breaking the editor. A 1.0 shell
 * (no members at all) gets the same calm fallbacks.
 */
import { office } from '../../lib/api.js';

const API = '/api/office';
const fileQuery = (file) => new URLSearchParams({ ...(file && file.root ? { root: file.root } : {}), path: (file && file.path) || '' }).toString();
async function json(url, opts) {
  const r = await fetch(url, { credentials: 'same-origin', ...opts });
  if (!r.ok) { let msg = `${r.status}`; try { const j = await r.json(); msg = j.error || j.message || msg; } catch { /* not json */ } throw Object.assign(new Error(msg), { status: r.status }); }
  return r.json();
}

export function available(host = {}) {
  let ai = false;
  try { ai = !!(host.ai && (typeof host.ai.ready === 'function' ? host.ai.ready() : typeof host.ai.ask === 'function')); } catch { ai = false; }
  return { ai, fields: true, remote: !!(host.remote && host.remote.start) || typeof WebSocket !== 'undefined', collab: !!(host.collab && host.collab.connect), compare: typeof host.compare === 'function' };
}

/* ---------------------------------------------------------------- AI */
/** ask Hostwerk's AI (the user's provider) → answer text */
export async function aiAsk(host, prompt, { system, json: wantJson, text } = {}) {
  const a = host && host.ai;
  if (!a || typeof a.ask !== 'function') throw new Error('AI is not available');
  const r = await a.ask({ prompt, system, ...(text ? { text } : {}), ...(wantJson ? { json: true } : {}) });
  if (r && typeof r === 'object') return r.json && wantJson ? JSON.stringify(r.json) : String(r.text ?? r.content ?? '');
  return String(r || '');
}

/* ---------------------------------------------------------------- live fields */
export async function fieldCatalog(host) {
  if (host && host.fields && typeof host.fields.list === 'function') return host.fields.list();
  const r = await office.fields.list();
  return Array.isArray(r) ? r : r.fields || [];
}
export async function fieldValue(host, id, params) {
  let v;
  if (host && host.fields && typeof host.fields.value === 'function') v = await host.fields.value(id, params);
  else v = await office.fields.value(id, params);
  if (v && typeof v === 'object') return v.text != null ? v.text : v.value;
  return v;
}

/* ---------------------------------------------------------------- phone remote */
/**
 * remote-control session over host.remote.start → { url, qr (data URL or null), send(state), close() };
 * onCmd({ cmd: 'next'|'prev'|'goto'|'laser', idx?, x?, y? }), onPeer({ count })
 */
export async function startRemote(host, { onCmd, onPeer, title, notes } = {}) {
  if (!host || !host.remote || typeof host.remote.start !== 'function') throw new Error('Remote control is not available');
  const s = await host.remote.start({ title, notes });
  const offs = [];
  const on = (evt, fn) => { if (typeof s.on === 'function') { const off = s.on(evt, fn); if (typeof off === 'function') offs.push(off); } };
  on('next', () => onCmd && onCmd({ cmd: 'next' }));
  on('prev', () => onCmd && onCmd({ cmd: 'prev' }));
  on('goto', (v) => onCmd && onCmd({ cmd: 'goto', idx: typeof v === 'object' && v ? v.idx ?? v.index : Number(v) }));
  on('black', () => onCmd && onCmd({ cmd: 'black' }));
  on('blank', () => onCmd && onCmd({ cmd: 'black' })); // the server's name for it (phone page "Black")
  on('first', () => onCmd && onCmd({ cmd: 'goto', idx: 0 }));
  on('end', () => onCmd && onCmd({ cmd: 'end' }));
  on('laser', (v) => onCmd && onCmd({ cmd: 'laser', ...(v && typeof v === 'object' && v.on !== false ? { x: v.x, y: v.y } : { x: null }) }));
  on('peers', (v) => onPeer && onPeer({ count: typeof v === 'number' ? v : Array.isArray(v) ? v.length : (v && v.count) || 0 }));
  return {
    url: s.url, qr: s.qr || null, token: s.token,
    // the server's state shape (server/remote.js): slide = 0-based position among the shown slides, total, title, notes, blank, timer
    send(state) {
      if (!state) return;
      const slide = state.pos != null ? Math.max(0, state.pos - 1) : state.index != null ? state.index : state.idx || 0;
      try { s.send({ type: 'state', slide, total: state.total || 0, title: state.title || '', notes: state.notes || '', blank: !!state.black, timer: state.startedAt ? { running: !state.ended, elapsed: Math.max(0, Date.now() - state.startedAt), startedAt: state.startedAt } : null }); } catch { /* closed */ }
    },
    close() { offs.forEach((f) => f()); try { s.stop(); } catch { /* stopped */ } },
  };
}

/* ---------------------------------------------------------------- co-editing */
/**
 * join co-editing of this file: host.collab.connect({ doc, user }) with a Y.Doc of the shared yjs.
 * → { ydoc, awareness, synced: Promise, isLeader(), close() } | null when the server part is missing.
 */
export async function joinCollab(host, Y, me) {
  if (!host || !host.collab || typeof host.collab.connect !== 'function' || !Y) return null;
  try { if (typeof host.collab.available === 'function' && !(await host.collab.available())) return null; } catch { return null; }
  const ydoc = new Y.Doc();
  const p = await host.collab.connect({ doc: ydoc, user: me });
  if (!p) return null;
  const awareness = p.awareness;
  return {
    ydoc, awareness, provider: p,
    synced: Promise.race([p.synced || Promise.resolve(true), new Promise((res) => setTimeout(() => res(false), 4000))]),
    // the co-editor with the smallest client id saves the file (one saver – no conflicts)
    isLeader() { if (!awareness) return true; let min = ydoc.clientID; awareness.getStates().forEach((st, id) => { if (st && st.user && id < min) min = id; }); return min === ydoc.clientID; },
    close() { try { if (host.collab.disconnect) host.collab.disconnect(); else p.destroy(); } catch { /* gone */ } },
  };
}

/* ---------------------------------------------------------------- versions */
export async function listVersions(host, file) {
  const r = await office.versions(file);
  return Array.isArray(r) ? r : (r && r.versions) || [];
}
export function versionBytes(host, file, id) { return office.versionBytes(file, id); }

/* ---------------------------------------------------------------- pick a file */
/** pick a document from this computer → { bytes, name, ext } | null */
export function pickDocument(host, accept) {
  return new Promise((res) => {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = accept;
    inp.onchange = async () => { const f = inp.files && inp.files[0]; if (!f) { res(null); return; } res({ bytes: new Uint8Array(await f.arrayBuffer()), name: f.name, ext: f.name.split('.').pop().toLowerCase() }); };
    inp.click();
  });
}
