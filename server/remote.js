'use strict';
/*
 * Phone as remote control for a slide show.
 *   presenter (logged in): POST /office/remote → { token, url, urls, qr (svg), expires } and the websocket
 *     /ws/office/remote?token=… (JSON text frames):
 *       presenter → server { type: 'state', slide, total, title, notes, blank, timer: { running, elapsed, startedAt } }
 *         or single values { type: 'slide', value: n | { index, total, title } } · { type: 'notes', value } · { type: 'timer', value }
 *         · { type: 'blank', value } · { type: 'end' }
 *       server → presenter { type: 'next'|'prev'|'first'|'last'|'goto'|'blank'|'laser', value?, cmd, n?, x?, y?, on? }
 *         (goto: value = slide index, laser: value = { x, y, on } 0…1, blank: value = bool) · { type: 'peers', value: phones, count }
 *   a logged-in phone may also use the websocket with role=phone (same commands as JSON, gets { type: 'state', … }).
 *   phone (public, token = the QR code): GET /office/pub/remote/:token (page), …/app.js, …/events (Server-Sent Events
 *     with the state), POST …/cmd { cmd, n, x, y, on }.
 * Hostwerk only hands /ws/* upgrades of logged-in users to extensions – the phone therefore uses plain HTTP (SSE + POST).
 */
const crypto = require('crypto');

const TTL = 12 * 3600000;
const CMDS = new Set(['next', 'prev', 'first', 'last', 'goto', 'blank', 'laser']);
const TOKEN = /^[A-Za-z0-9_-]{32}$/;

class Remote {
  constructor({ now = Date.now } = {}) { this.now = now; this.sessions = new Map(); }
  create(user) {
    this.sweep();
    if (this.sessions.size > 100) throw Object.assign(new Error('Too many remote controls at once.'), { status: 503 });
    const token = crypto.randomBytes(24).toString('base64url');
    this.sessions.set(token, { token, user: String(user || '').slice(0, 60), created: this.now(), expires: this.now() + TTL, presenter: null, phones: new Set(), state: { slide: 0, total: 0, title: '', notes: '', blank: false, timer: null }, hits: [] });
    return this.sessions.get(token);
  }
  get(token) { if (!TOKEN.test(String(token || ''))) return null; const s = this.sessions.get(token); if (!s || s.expires < this.now()) return null; return s; }
  end(token) { const s = this.sessions.get(token); if (!s) return; for (const ph of s.phones) { try { ph.end({ type: 'end' }); } catch { /* egal */ } } this.sessions.delete(token); }
  sweep() { const now = this.now(); for (const [k, s] of this.sessions) if (s.expires < now) this.end(k); }
  /** presenter: { send(obj), close() } */
  attach(token, presenter) {
    const s = this.get(token); if (!s) return null;
    if (s.presenter) { try { s.presenter.close(); } catch { /* egal */ } }
    s.presenter = presenter;
    presenter.send({ type: 'peers', value: s.phones.size, count: s.phones.size });
    return {
      message: (txt) => {
        let m; try { m = JSON.parse(String(txt)); } catch { return; }
        if (!m || typeof m !== 'object') return;
        if (m.type === 'end') { this.end(token); return; }
        if (m.type === 'slide' || m.type === 'notes' || m.type === 'timer' || m.type === 'blank' || m.type === 'title') {
          const v = m.value; const cur = { ...s.state };
          if (m.type === 'slide') { if (v && typeof v === 'object') { cur.slide = v.index != null ? v.index : v.slide; if (v.total != null) cur.total = v.total; if (v.title != null) cur.title = v.title; if (v.notes != null) cur.notes = v.notes; } else cur.slide = v; }
          else if (m.type === 'timer') cur.timer = v; else cur[m.type] = v;
          m = { ...cur, type: 'state' };
        }
        if (m.type !== 'state') return;
        const t = m.timer && typeof m.timer === 'object' ? { running: !!m.timer.running, elapsed: Math.max(0, Number(m.timer.elapsed) || 0), startedAt: Number(m.timer.startedAt) || null } : null;
        s.state = { slide: Math.max(0, Math.floor(Number(m.slide) || 0)), total: Math.max(0, Math.floor(Number(m.total) || 0)), title: String(m.title || '').slice(0, 200), notes: String(m.notes || '').slice(0, 20000), blank: !!m.blank, timer: t, at: this.now() };
        for (const ph of s.phones) { try { ph.send({ type: 'state', ...s.state }); } catch { /* egal */ } }
      },
      close: () => { if (s.presenter === presenter) s.presenter = null; for (const ph of s.phones) { try { ph.send({ type: 'presenter', connected: false }); } catch { /* egal */ } } },
    };
  }
  /** phone: { send(obj), end(obj) } → off() */
  phone(token, ph) {
    const s = this.get(token); if (!s) return null;
    if (s.phones.size >= 5) throw Object.assign(new Error('Too many phones'), { status: 429 });
    s.phones.add(ph);
    ph.send({ type: 'state', ...s.state, presenter: !!s.presenter });
    if (s.presenter) s.presenter.send({ type: 'peers', value: s.phones.size, count: s.phones.size });
    return () => { s.phones.delete(ph); if (s.presenter) { try { s.presenter.send({ type: 'peers', value: s.phones.size, count: s.phones.size }); } catch { /* egal */ } } };
  }
  command(token, body0) {
    const body = body0 && !body0.cmd && body0.type ? { ...body0, cmd: body0.type } : body0;
    const s = this.get(token); if (!s) throw Object.assign(new Error('This remote control has ended.'), { status: 404 });
    const now = this.now();
    s.hits = s.hits.filter((t) => now - t < 1000); if (s.hits.length >= 20) throw Object.assign(new Error('Too many commands'), { status: 429 });
    const cmd = String((body && body.cmd) || '');
    if (!CMDS.has(cmd)) throw Object.assign(new Error('Unknown command'), { status: 400 });
    const out = { type: cmd, cmd };
    const v = body.value && typeof body.value === 'object' ? body.value : {};
    if (cmd === 'goto') { out.n = Math.max(0, Math.floor(Number(body.n != null ? body.n : body.value) || 0)); out.value = out.n; }
    if (cmd === 'laser') { out.on = (body.on != null ? body.on : v.on) !== false; out.x = Math.min(1, Math.max(0, Number(body.x != null ? body.x : v.x) || 0)); out.y = Math.min(1, Math.max(0, Number(body.y != null ? body.y : v.y) || 0)); out.value = { x: out.x, y: out.y, on: out.on }; }
    if (cmd === 'blank') { const b = body.on != null ? body.on : body.value; out.on = b === undefined || b === null ? !s.state.blank : !!b; out.value = out.on; }
    if (!s.presenter) throw Object.assign(new Error('The slide show is not connected.'), { status: 409 });
    s.hits.push(now);
    s.presenter.send(out);
    return { ok: true };
  }
}

const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const T = {
  de: { title: 'Fernbedienung', next: 'Weiter', prev: 'Zurück', blank: 'Schwarz', laser: 'Laser', notes: 'Notizen', none: 'Keine Notizen', waiting: 'Warte auf die Präsentation …', ended: 'Die Präsentation ist beendet.', slide: 'Folie' },
  en: { title: 'Remote control', next: 'Next', prev: 'Back', blank: 'Black', laser: 'Laser', notes: 'Notes', none: 'No notes', waiting: 'Waiting for the presentation …', ended: 'The presentation has ended.', slide: 'Slide' },
};
function phonePage(lang, base) {
  const t = T[lang] || T.en;
  return `<!doctype html><html lang="${lang}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no"><meta name="robots" content="noindex"><meta name="theme-color" content="#111"><title>${t.title}</title>
<style>:root{color-scheme:dark}*{box-sizing:border-box}body{margin:0;background:#111;color:#eee;font:16px/1.4 system-ui,-apple-system,"Segoe UI",sans-serif;display:flex;flex-direction:column;min-height:100vh;-webkit-user-select:none;user-select:none}
header{display:flex;justify-content:space-between;align-items:center;padding:12px 16px;border-bottom:1px solid #2a2a2a}h1{font-size:15px;margin:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#count,#timer{font-variant-numeric:tabular-nums;color:#aaa}.pad{display:grid;grid-template-columns:1fr 2fr;gap:12px;padding:16px}
button{font:inherit;font-size:20px;border:0;border-radius:14px;padding:26px 8px;background:#2a2d34;color:#fff;touch-action:manipulation}button:active{background:#3a3e48}
#next{background:#2563eb}#next:active{background:#1d4ed8}.row{display:flex;gap:12px;padding:0 16px}.row button{flex:1;font-size:16px;padding:14px}button[aria-pressed="true"]{background:#b45309}
#laserpad{margin:12px 16px;border:1px dashed #444;border-radius:12px;height:120px;display:none;touch-action:none}#notes{flex:1;margin:12px 16px 16px;padding:12px;background:#1b1d22;border-radius:12px;white-space:pre-wrap;overflow:auto;font-size:17px}
.msg{padding:24px;text-align:center;color:#aaa}@media (prefers-reduced-motion:no-preference){button{transition:background .12s}}</style></head>
<body data-base="${esc(base)}"><header><h1 id="title">${t.title}</h1><span id="count"></span><span id="timer"></span></header>
<div class="pad"><button id="prev" aria-label="${t.prev}">‹ ${t.prev}</button><button id="next" aria-label="${t.next}">${t.next} ›</button></div>
<div class="row"><button id="blank" aria-pressed="false">${t.blank}</button><button id="laser" aria-pressed="false">${t.laser}</button></div>
<div id="laserpad" aria-label="${t.laser}"></div><div id="notes" aria-label="${t.notes}"><div class="msg">${t.waiting}</div></div>
<script type="application/json" id="t">${JSON.stringify(t).replace(/</g, '\\u003c')}</script><script src="${esc(base)}/app.js"></script></body></html>`;
}
const PHONE_JS = `(function () {
  var base = document.body.getAttribute('data-base'); var T = JSON.parse(document.getElementById('t').textContent);
  var $ = function (id) { return document.getElementById(id); }; var st = {}; var laser = false; var last = 0;
  function send(o) { return fetch(base + '/cmd', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(o), credentials: 'omit' }).catch(function () {}); }
  function vib() { if (navigator.vibrate) navigator.vibrate(15); }
  $('next').onclick = function () { vib(); send({ cmd: 'next' }); }; $('prev').onclick = function () { vib(); send({ cmd: 'prev' }); };
  $('blank').onclick = function () { send({ cmd: 'blank' }); };
  $('laser').onclick = function () { laser = !laser; $('laser').setAttribute('aria-pressed', String(laser)); $('laserpad').style.display = laser ? 'block' : 'none'; if (!laser) send({ cmd: 'laser', on: false }); };
  var pad = $('laserpad');
  pad.addEventListener('pointermove', function (e) { if (!laser) return; var n = Date.now(); if (n - last < 50) return; last = n; var r = pad.getBoundingClientRect(); send({ cmd: 'laser', on: true, x: (e.clientX - r.left) / r.width, y: (e.clientY - r.top) / r.height }); });
  function fmt(ms) { var s = Math.floor(ms / 1000); var m = Math.floor(s / 60); return m + ':' + String(s % 60).padStart(2, '0'); }
  setInterval(function () { var t = st.timer; if (!t) { $('timer').textContent = ''; return; } $('timer').textContent = fmt(t.running && t.startedAt ? t.elapsed + (Date.now() - t.at0) : t.elapsed); }, 500);
  function show(s) {
    if (s.type === 'end') { $('notes').innerHTML = ''; var d = document.createElement('div'); d.className = 'msg'; d.textContent = T.ended; $('notes').appendChild(d); return; }
    if (s.type !== 'state') return; st = s; if (s.timer) s.timer.at0 = Date.now();
    $('title').textContent = s.title || T.title; $('count').textContent = s.total ? T.slide + ' ' + (s.slide + 1) + ' / ' + s.total : '';
    $('blank').setAttribute('aria-pressed', String(!!s.blank));
    $('notes').textContent = s.notes || T.none;
  }
  var es = new EventSource(base + '/events'); es.onmessage = function (e) { try { show(JSON.parse(e.data)); } catch (x) {} };
})();`;

module.exports = { Remote, phonePage, PHONE_JS, TOKEN };
