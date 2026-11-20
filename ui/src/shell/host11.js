/*
 * Host object 1.1 (ARCHITECTURE.md "Host object 1.1") – the members the editor host adds for the chunks:
 * data, db, fields, forms, rules, collab, remote, ai, components, compare, palette, insertFiles, toast, perf.
 * extendHost(host, ctx) fills them in place (the host object stays the same object for the chunk's whole life).
 *   ctx: { file() → current file, chunk() → chunk module, current() → Promise<Uint8Array> (saver), onPeers(list),
 *          onCommands(list), onDropTarget(fn|null), onInteractive() }
 */
import * as HAPI from '@hostwerk/api';
import * as KIT from '@hostwerk/kit';
import { toast } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { office, isMissing, BASE } from '../lib/api';
import { CHUNK_VERSION } from '../lib/version';
import { extOf } from '../lib/formats';
import { askFeature } from './features';
import { fmtDateTime } from './common';

const loc = (f) => (f.root ? { root: f.root, path: f.path } : { path: f.path });
const missing = (what) => Object.assign(new Error(t('{what} needs a newer Office server part – update Office and try again.', { what })), { status: 501 });
const wrap = (what, p) => p.catch((e) => { if (isMissing(e)) throw missing(what); throw e; });

/* ---------------- live data subscriptions (one websocket listener per source) ---------------- */
function subscribe(source, params, cb) {
  let alive = true;
  let last = 0;
  const run = () => office.data.query(source, params).then((r) => { if (alive) { last = Date.now(); cb(r); } }, () => {});
  run();
  const off = typeof HAPI.on === 'function' ? HAPI.on('office-data', (m) => { if (!m || !m.source || m.source === source) run(); }) : null;
  const timer = setInterval(() => { if (!document.hidden && Date.now() - last > 29000) run(); }, 30000);
  return () => { alive = false; clearInterval(timer); if (typeof off === 'function') off(); };
}

/* ---------------- Yjs (shared library chunk lib-yjs.js) ---------------- */
let yjsLib = null;
export function loadYjs() {
  if (!yjsLib) {
    yjsLib = import(/* @vite-ignore */ `${office.assetUrl('lib-yjs.js')}?v=${CHUNK_VERSION}`).then((m) => {
      const d = m.default;
      const L = globalThis.__officeLib || (globalThis.__officeLib = {});
      L.yjs = L.yjs || d.Y; L.ysync = L.ysync || d.syncProtocol; L.yawareness = L.yawareness || d.awarenessProtocol;
      return d;
    });
    yjsLib.catch(() => { yjsLib = null; });
  }
  return yjsLib;
}
const COLORS = ['#2b6fd6', '#14995a', '#e0662b', '#a24bc0', '#0e8f86', '#dc3545', '#b0761a', '#3d5bc4'];
export function me() {
  let name = '';
  try { const s = typeof HAPI.getStore === 'function' ? HAPI.getStore() : null; name = (s && (s.user && (s.user.name || s.user.login))) || (s && s.settings && s.settings.general && s.settings.general.name) || ''; } catch { name = ''; }
  let id = '';
  try { id = localStorage.getItem('office:me') || ''; if (!id) { id = Math.random().toString(36).slice(2, 10); localStorage.setItem('office:me', id); } } catch { id = Math.random().toString(36).slice(2, 10); }
  const n = [...id].reduce((s, c) => s + c.charCodeAt(0), 0);
  return { id, name: name || t('Guest {n}', { n: (n % 90) + 10 }), color: COLORS[n % COLORS.length], ...(name ? {} : { guest: true }) };
}

/* components on demand: the server answers { id, state: 'ready', installed, files: [names], base, aliases } */
export const findComponent = (list, id) => (list || []).find((x) => x.id === id || (Array.isArray(x.aliases) && x.aliases.includes(id))) || null;
export const componentReady = (c) => !!(c && (c.state === 'ready' || c.installed === true));
const componentResult = (c) => ({ id: c.id, base: new URL(c.base || office.components.base(c.id), location.href).href, files: Array.isArray(c.files) ? c.files : [] });

export function extendHost(host, ctx) {
  const file = () => ctx.file();
  Object.defineProperty(host, 'file', { configurable: true, enumerable: true, get: () => { const f = file(); return { name: f.name, path: f.path, root: f.root, ext: extOf(f.name) }; } });

  host.data = {
    sources: () => wrap(t('Live data'), office.data.sources()),
    query: (source, params) => wrap(t('Live data'), office.data.query(source, params)),
    subscribe: (source, params, cb) => subscribe(source, params || {}, cb),
    pick: (o = {}) => askFeature('data-pick', o),
  };
  host.db = {
    servers: () => wrap(t('Databases'), office.db.servers()),
    tables: (server, db) => wrap(t('Databases'), office.db.tables(server, db)),
    rows: (o) => wrap(t('Databases'), office.db.rows(o)),
    pick: () => askFeature('db-pick'),
    /** asks first (n changes) – null when the user cancels */
    write: async ({ server, db, table, changes }) => {
      if (!Array.isArray(changes) || !changes.length) return { ok: true, changed: 0 };
      const ok = await askFeature('db-confirm', { table, changes });
      if (!ok) return null;
      const r = await wrap(t('Databases'), office.db.write({ server, db, table, changes }));
      toast.success(t('{n} changes written to “{table}”', { n: changes.length, table }));
      return r;
    },
  };
  host.fields = {
    list: () => wrap(t('Fields'), office.fields.list()),
    value: (id, params) => wrap(t('Fields'), office.fields.value(id, params)),
    pick: () => askFeature('field-pick'),
  };
  host.forms = {
    list: () => wrap(t('Forms'), office.forms.list(loc(file()))),
    create: (form) => wrap(t('Forms'), office.forms.create(loc(file()), form)),
    open: (o = {}) => askFeature('forms', { file: file(), ...o }),
  };
  host.rules = {
    get: () => wrap(t('Rules'), office.rules.get(loc(file()))),
    set: (rules) => wrap(t('Rules'), office.rules.set(loc(file()), rules)),
    open: (o = {}) => askFeature('rules', { file: file(), ...o }),
  };
  let provider = null;
  host.collab = {
    available: async () => { try { return (await fetch(`${BASE}/collab/state?${new URLSearchParams(loc(file()))}`, { credentials: 'same-origin' })).ok; } catch { return false; } },
    /** connect a Y.Doc (of the shared yjs) to the co-editing relay → provider */
    connect: async ({ doc, awareness = null, user = null } = {}) => {
      const lib = await loadYjs();
      if (!doc) doc = new lib.Y.Doc();
      if (provider) provider.destroy();
      const u0 = user || me();
      provider = new lib.OfficeProvider(office.wsUrl('/collab', loc(file())), doc, { user: u0, awareness });
      // no name in the browser ("Guest 23"): the server knows who is logged in (owner name or the team member)
      if (!user && u0.guest) provider.on('control', (m) => { if (m && m.type === 'hello' && m.user && provider) provider.awareness.setLocalStateField('user', { ...u0, name: String(m.user).slice(0, 60), guest: false }); });
      provider.on('peers', (list) => ctx.onPeers && ctx.onPeers(list));
      provider.on('status', (s) => ctx.onPeers && ctx.onPeers(provider.peers(), s));
      // any save of the file (the leader's or from outside): the editor host learns the new file time, so a later save of
      // this editor (when it becomes the leader) is no false "changed on disk" conflict
      provider.on('saved', (m) => { if (m && m.mtime && ctx.onRemoteSaved) ctx.onRemoteSaved(Number(m.mtime), m.by || null); });
      provider.on('reset', (m) => { if (ctx.onCollabReset) ctx.onCollabReset(m); });
      return provider;
    },
    disconnect: () => { if (provider) { provider.destroy(); provider = null; ctx.onPeers && ctx.onPeers([], 'offline'); } },
    me,
  };
  host.remote = {
    /** remote control of a slide show → session { token, url, qr, on, send, stop } */
    start: async ({ notes = null, title = '' } = {}) => {
      const r = await wrap(t('Remote control'), HAPI.post('/office/remote', { ...loc(file()), title }));
      const ws = new WebSocket(office.wsUrl('/remote', { token: r.token, role: 'presenter' }));
      const ls = new Map();
      const on = (evt, fn) => { if (!ls.has(evt)) ls.set(evt, new Set()); ls.get(evt).add(fn); return () => ls.get(evt).delete(fn); };
      ws.onmessage = (ev) => { let m = null; try { m = JSON.parse(ev.data); } catch { return; } const s = ls.get(m.type); if (s) s.forEach((f) => f(m.value ?? m)); };
      const send = (msg) => { if (ws.readyState === 1) ws.send(JSON.stringify(msg)); };
      if (notes) ws.addEventListener('open', () => send({ type: 'notes', value: notes }));
      const session = { ...r, on, send, stop: () => { try { ws.close(); } catch { /* closed */ } HAPI.del(`/office/remote/${encodeURIComponent(r.token)}`).catch(() => {}); } };
      host.remote.current = session;
      return session;
    },
    dialog: (session) => askFeature('remote', { session: session || host.remote.current }),
    current: null,
  };
  host.ai = {
    ready: () => { try { return typeof KIT.aiReadyNow === 'function' ? !!KIT.aiReadyNow() : false; } catch { return false; } },
    quick: (task, text, { context = '' } = {}) => HAPI.post(`/ai/quick/${encodeURIComponent(task)}`, { text, lang: host.lang, context }).then((r) => r.text),
    ask: (body) => wrap(t('AI'), office.ai({ ...body, lang: host.lang, ...loc(file()) })),
    panel: (question, context) => { if (typeof KIT.openAiPanel === 'function') return KIT.openAiPanel({ question, context: context || t('Office file “{name}”', { name: file().name }) }); return false; },
  };
  host.components = {
    state: () => wrap(t('Components'), office.components.list()),
    /** → { id, base } when installed (asks once with size + progress), null when declined */
    ensure: async (id, { reason = '', label = '' } = {}) => {
      let list = [];
      try { list = await office.components.list(); } catch (e) { if (isMissing(e)) throw missing(t('Components')); throw e; }
      const c = findComponent(list, id);
      const cid = c ? c.id : id; // aliases (ocr-deu → ocr, pdf-cmaps → cjk-fonts) → the real id
      if (componentReady(c)) return componentResult(c);
      const r = await askFeature('component', { id: cid, label: label || (c && c.label) || id, reason });
      return r ? componentResult(r) : null;
    },
  };
  /** host.compare(a, b = 'current') – a/b: { id } (version), 'current', { bytes, label } */
  host.compare = async (a, b = 'current') => {
    const f = file();
    const resolve = async (x) => {
      if (x === 'current') { const bytes = await ctx.current(); return { bytes, label: t('Now (not saved yet)') }; }
      if (x && x.bytes) return { bytes: x.bytes instanceof Uint8Array ? x.bytes : new Uint8Array(x.bytes), label: x.label || '' };
      if (x && x.id) return { bytes: await office.versionBytes(loc(f), x.id), label: x.label || (x.time ? fmtDateTime(x.time) : t('Version')) };
      throw new Error(t('Nothing to compare.'));
    };
    const [A, B] = await Promise.all([resolve(a), resolve(b)]);
    return askFeature('compare', { a: A, b: B, chunk: ctx.chunk(), ext: extOf(f.name), host });
  };
  host.palette = { register: (cmds) => { ctx.onCommands(Array.isArray(cmds) ? cmds : []); return () => ctx.onCommands([]); }, open: () => ctx.openPalette && ctx.openPalette() };
  host.insertFiles = (cb) => { ctx.onDropTarget(typeof cb === 'function' ? cb : null); return () => ctx.onDropTarget(null); };
  host.toast = (text, { kind = 'info', action = null, undo = null, ms } = {}) => {
    const act = undo ? { label: t('Undo'), onClick: undo } : action ? { label: action.label, onClick: action.run || action.onClick } : null;
    if (kind === 'error') return toast.error(text);
    return toast(text, kind === 'success' ? 'success' : kind === 'warn' ? 'warn' : 'info', ms || (act ? 8000 : 4500), act ? { action: act } : null);
  };
  host.perf = (mark) => { if (mark === 'interactive' && ctx.onInteractive) ctx.onInteractive(); };
  return host;
}
