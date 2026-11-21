/*
 * Soft lock of an open editor (server/meta.js, POST /api/office/lock|unlock) – and the lock marks in Hostwerk's Files.
 *
 * - CLIENT: this window (one JS context); Files asks for the marks with it and never shows the locks of its own window
 *   (when Files shows here, this window's editor is closed).
 * - sid/seq: one open editor and the order of its requests – the server ignores a heartbeat that arrives after the
 *   unlock and never lets the late unlock of a closed editor release the lock of the next one.
 * - The lock is renewed every 2 minutes while the page is visible (a window hidden in the tray lets it expire, it is
 *   taken again as soon as the window shows); released on unmount (Back, Close, route change), `pagehide` (tab or window
 *   closed, reload) and `beforeunload` without unsaved changes – with `fetch(…, { keepalive: true })`, which outlives the
 *   page. Afterwards the marks of Files are refreshed (`hw:files-marks`) – here, and through the server's broadcast
 *   `office-lock` in every other window.
 * No imports from Hostwerk: node --test runs this file (tests/shell-locks.test.js).
 */
const rnd = () => Math.random().toString(36).slice(2, 10);
export const CLIENT = `w${Date.now().toString(36)}${rnd()}`;
export const LOCK_EVERY = 120000;

/** Files: refresh the marks of the open folder (Office's mark cache listens to the same event) */
export function refreshMarks(win = typeof window !== 'undefined' ? window : null) {
  try { if (win) win.dispatchEvent(new Event('hw:files-marks')); } catch { /* old browser */ }
}

/** POST /api/office/<op> – keepalive: the request outlives the page (closing tab/window) → parsed JSON; throws { status, lock } */
export async function sendLock(op, body, { fetchFn = typeof fetch !== 'undefined' ? fetch : null, keepalive = false } = {}) {
  if (!fetchFn) throw new Error('fetch is not available');
  const res = await fetchFn(`/api/office/${op}`, { method: 'POST', keepalive, credentials: 'same-origin', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
  let j = null;
  try { j = await res.json(); } catch { /* empty */ }
  if (!res.ok) throw Object.assign(new Error((j && j.error) || `${res.status}`), { status: res.status, lock: j && j.lock ? j.lock : null });
  return j || {};
}

/**
 * The lock of one open editor. file: { root?, path } · onResult(result | null, error | null) after every lock request ·
 * dirty(): unsaved changes (beforeunload may still be cancelled then – pagehide releases) · send(op, body, { keepalive })
 * → start() / stop() (stop = released, returns the unlock promise).
 */
export function editorLock({ file, onResult, dirty = () => false, send = (op, body, o) => sendLock(op, body, o), win = typeof window !== 'undefined' ? window : null, doc = typeof document !== 'undefined' ? document : null, every = LOCK_EVERY, client = CLIENT, sid = `e${rnd()}${rnd()}` } = {}) {
  let seq = 0;
  let alive = false;
  let timer = null;
  const body = () => ({ ...(file.root ? { root: file.root } : {}), path: file.path, client, sid, seq: ++seq });
  const hidden = () => !!(doc && doc.visibilityState === 'hidden');
  const lock = () => {
    if (!alive || hidden()) return null;
    return send('lock', body(), {}).then((r) => { if (alive && onResult) onResult(r, null); return r; }, (e) => { if (alive && onResult) onResult(null, e); return null; });
  };
  const release = () => {
    // sent synchronously (pagehide: nothing after this handler is sure to run); a failed unlock expires on its own
    let p;
    try { p = Promise.resolve(send('unlock', body(), { keepalive: true })).catch(() => null); } catch { p = Promise.resolve(null); }
    p.then(() => refreshMarks(win));
    return p;
  };
  const onPageHide = () => { release(); };
  const onPageShow = (e) => { if (e && e.persisted) lock(); }; // back from the back/forward cache
  const onBeforeUnload = () => { if (!dirty()) release(); };
  const onVisible = () => { if (!hidden()) lock(); };
  return {
    start() {
      if (alive) return;
      alive = true;
      lock();
      timer = setInterval(() => lock(), every);
      if (win) { win.addEventListener('pagehide', onPageHide); win.addEventListener('pageshow', onPageShow); win.addEventListener('beforeunload', onBeforeUnload); }
      if (doc) doc.addEventListener('visibilitychange', onVisible);
    },
    stop() {
      if (!alive) return Promise.resolve(null);
      alive = false;
      clearInterval(timer); timer = null;
      if (win) { win.removeEventListener('pagehide', onPageHide); win.removeEventListener('pageshow', onPageShow); win.removeEventListener('beforeunload', onBeforeUnload); }
      if (doc) doc.removeEventListener('visibilitychange', onVisible);
      return release();
    },
    get sid() { return sid; },
  };
}

/** "since" of a lock → "14:05" today, else "19.11., 14:05" (lang: Hostwerk's language) */
export function lockSince(since, lang = 'en', now = Date.now()) {
  const d = new Date(Number(since) || 0);
  if (!since || Number.isNaN(d.getTime())) return '';
  const time = d.toLocaleTimeString(lang, { hour: '2-digit', minute: '2-digit' });
  const n = new Date(now);
  if (d.toDateString() === n.toDateString()) return time;
  return `${d.toLocaleDateString(lang, { day: 'numeric', month: 'numeric' })}, ${time}`;
}
