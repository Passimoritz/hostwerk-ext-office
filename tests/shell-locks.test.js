// Soft lock of an open editor (ui/src/lib/locks.js) – node --test. Every way of closing releases the lock with a
// keepalive request, in order (seq), and refreshes the marks of Files; the heartbeat stops after close and pauses
// while the window is hidden (tray), taking the lock again when it shows.
import test from 'node:test';
import assert from 'node:assert/strict';
import { editorLock, sendLock, lockSince, refreshMarks, CLIENT } from '../ui/src/lib/locks.js';

function env() {
  const win = new EventTarget();
  const doc = new EventTarget();
  doc.visibilityState = 'visible';
  const sent = [];
  let marks = 0;
  win.addEventListener('hw:files-marks', () => { marks++; });
  const send = (op, body, o) => { sent.push({ op, body, keepalive: !!(o && o.keepalive) }); return Promise.resolve(op === 'lock' ? { ok: true, lock: { by: 'Anna', since: 1, other: false } } : { ok: true, released: true }); };
  const flush = () => new Promise((r) => setTimeout(r, 5));
  return { win, doc, sent, send, flush, marks: () => marks };
}
const file = { root: 'documents', path: 'Projekte - Claude/ISR/Plan v0.3.docx' };

test('Back / Close / route change (unmount): unlock with keepalive, newer seq, marks refreshed, no heartbeat after', async () => {
  const e = env();
  const l = editorLock({ file, send: e.send, win: e.win, doc: e.doc, every: 20 });
  l.start();
  await e.flush();
  assert.equal(e.sent[0].op, 'lock');
  assert.deepEqual([e.sent[0].body.root, e.sent[0].body.path, e.sent[0].body.client], [file.root, file.path, CLIENT]);
  await new Promise((r) => setTimeout(r, 50));
  const locks = e.sent.filter((x) => x.op === 'lock').length;
  assert.ok(locks >= 2, 'heartbeat renews');
  await l.stop();
  await e.flush();
  const last = e.sent[e.sent.length - 1];
  assert.equal(last.op, 'unlock'); assert.equal(last.keepalive, true);
  assert.equal(last.body.sid, e.sent[0].body.sid);
  assert.ok(last.body.seq > Math.max(...e.sent.filter((x) => x.op === 'lock').map((x) => x.body.seq)), 'unlock is the newest request');
  assert.equal(e.marks(), 1, 'Files marks refreshed after the unlock');
  const n = e.sent.length;
  await new Promise((r) => setTimeout(r, 60));
  assert.equal(e.sent.length, n, 'no heartbeat after close');
  e.win.dispatchEvent(new Event('pagehide'));
  assert.equal(e.sent.length, n, 'listeners removed');
});

test('tab / window closed (pagehide) and closing without unsaved changes (beforeunload): keepalive unlock', async () => {
  const e = env();
  const l = editorLock({ file, send: e.send, win: e.win, doc: e.doc, every: 100000 });
  l.start();
  e.win.dispatchEvent(new Event('pagehide'));
  assert.equal(e.sent[e.sent.length - 1].op, 'unlock', 'sent synchronously in the handler');
  assert.equal(e.sent[e.sent.length - 1].keepalive, true);
  // back from the back/forward cache: locked again
  const ev = new Event('pageshow'); ev.persisted = true; e.win.dispatchEvent(ev);
  assert.equal(e.sent[e.sent.length - 1].op, 'lock');
  e.win.dispatchEvent(new Event('beforeunload'));
  assert.equal(e.sent[e.sent.length - 1].op, 'unlock');
  l.stop();
});

test('beforeunload with unsaved changes does not release (the page may stay) – pagehide does', () => {
  const e = env();
  const l = editorLock({ file, send: e.send, win: e.win, doc: e.doc, every: 100000, dirty: () => true });
  l.start();
  e.win.dispatchEvent(new Event('beforeunload'));
  assert.equal(e.sent[e.sent.length - 1].op, 'lock');
  e.win.dispatchEvent(new Event('pagehide'));
  assert.equal(e.sent[e.sent.length - 1].op, 'unlock');
  l.stop();
});

test('hidden window (tray): no renewal – the lock expires; shown again: locked at once', async () => {
  const e = env();
  const l = editorLock({ file, send: e.send, win: e.win, doc: e.doc, every: 10 });
  l.start();
  e.doc.visibilityState = 'hidden';
  const n = e.sent.length;
  await new Promise((r) => setTimeout(r, 50));
  assert.equal(e.sent.length, n, 'no heartbeat while hidden');
  e.doc.visibilityState = 'visible';
  e.doc.dispatchEvent(new Event('visibilitychange'));
  assert.equal(e.sent[e.sent.length - 1].op, 'lock');
  l.stop();
});

test('a failing unlock never throws; another user holding the file reaches onResult', async () => {
  const e = env();
  const got = [];
  const send = (op) => (op === 'lock' ? Promise.reject(Object.assign(new Error('Ben is editing this file right now.'), { status: 409, lock: { by: 'Ben' } })) : Promise.reject(new Error('offline')));
  const l = editorLock({ file, send, win: e.win, doc: e.doc, every: 100000, onResult: (r, er) => got.push([r, er && er.status]) });
  l.start();
  await e.flush();
  assert.deepEqual(got, [[null, 409]]);
  assert.equal(await l.stop(), null);
});

test('sendLock: POST JSON with keepalive, errors carry status and lock', async () => {
  const calls = [];
  const ok = await sendLock('unlock', { path: 'a.docx' }, { keepalive: true, fetchFn: async (u, o) => { calls.push([u, o]); return { ok: true, status: 200, json: async () => ({ ok: true }) }; } });
  assert.deepEqual(ok, { ok: true });
  assert.equal(calls[0][0], '/api/office/unlock');
  assert.equal(calls[0][1].keepalive, true); assert.equal(calls[0][1].method, 'POST');
  await assert.rejects(sendLock('lock', {}, { fetchFn: async () => ({ ok: false, status: 409, json: async () => ({ error: 'busy', lock: { by: 'Ben' } }) }) }), (er) => er.status === 409 && er.lock.by === 'Ben');
});

test('lockSince: time today, date and time before; refreshMarks dispatches hw:files-marks', () => {
  const now = new Date(2026, 10, 20, 15, 0).getTime();
  assert.match(lockSince(new Date(2026, 10, 20, 14, 5).getTime(), 'de', now), /^14:05$/);
  assert.match(lockSince(new Date(2026, 10, 19, 9, 7).getTime(), 'de', now), /19\.11\.?, 09:07/);
  assert.equal(lockSince(0, 'de', now), '');
  const w = new EventTarget(); let n = 0; w.addEventListener('hw:files-marks', () => n++);
  refreshMarks(w); assert.equal(n, 1);
});
