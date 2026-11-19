'use strict';
/*
 * Read-only share links for other people: random token (192 bit), optional expiry, optional password (scrypt),
 * allowDownload, view counter. The public page (host.publicRouter) shows PDFs, images and text inline, office files as
 * PDF through the converter (cached per file version), everything else as a download.
 */
const crypto = require('crypto');

const TOKEN = /^[A-Za-z0-9_-]{32}$/;
const newToken = () => crypto.randomBytes(24).toString('base64url');

function hashPassword(pw) {
  const salt = crypto.randomBytes(16);
  const hash = crypto.scryptSync(String(pw), salt, 32, { N: 16384, r: 8, p: 1 });
  return `scrypt$${salt.toString('base64')}$${hash.toString('base64')}`;
}
function checkPassword(stored, pw) {
  const m = /^scrypt\$([^$]+)\$([^$]+)$/.exec(String(stored || ''));
  if (!m) return false;
  const want = Buffer.from(m[2], 'base64');
  const got = crypto.scryptSync(String(pw || ''), Buffer.from(m[1], 'base64'), want.length, { N: 16384, r: 8, p: 1 });
  return crypto.timingSafeEqual(want, got);
}
/** Expiry from the request: timestamp (ms), or days (≤ 3650), 0/null = never → ms or null; throws on the past */
function expiryOf(v, defaultDays, now = Date.now()) {
  if (v === undefined) return defaultDays ? now + defaultDays * 86400000 : null;
  if (v === null || v === '' || v === 0 || v === false) return null;
  const n = Number(v);
  if (!Number.isFinite(n) || n < 0) throw Object.assign(new Error('Invalid expiry date'), { status: 400 });
  const t = n <= 3650 ? now + n * 86400000 : n;
  if (t <= now) throw Object.assign(new Error('The expiry date must be in the future.'), { status: 400 });
  return t;
}
const expired = (s, now = Date.now()) => !!(s && s.expires && s.expires <= now);
/** Valid share for a token or null */
function find(list, token, now = Date.now()) {
  if (!TOKEN.test(String(token || ''))) return null;
  const s = (list || []).find((x) => x.token === token);
  return s && !expired(s, now) ? s : null;
}
/** Cookie value that proves the password was entered (bound to token + password hash) */
const unlockValue = (secret, s) => crypto.createHmac('sha256', secret).update(`${s.token}\n${s.hash || ''}`).digest('base64url');
/** Public view of a share (no hash) */
const view = (s, now = Date.now()) => ({ id: s.id, token: s.token, path: s.path, root: s.root || null, name: s.name, created: s.created, expires: s.expires || null, expired: expired(s, now), password: !!s.hash, allowDownload: !!s.allowDownload, views: s.views || 0, lastView: s.lastView || null });

const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const SIZE = (n) => { const u = ['B', 'KB', 'MB', 'GB']; let i = 0; let x = n; while (x >= 1024 && i < u.length - 1) { x /= 1024; i++; } return `${x.toFixed(x >= 10 || i === 0 ? 0 : 1)} ${u[i]}`; };
const T = {
  de: { download: 'Herunterladen', password: 'Passwort', open: 'Öffnen', locked: 'Diese Datei ist mit einem Passwort geschützt.', wrong: 'Das Passwort stimmt nicht.', noPreview: 'Für diese Datei gibt es keine Vorschau.', noDownload: 'Der Absender hat das Herunterladen nicht erlaubt.', footer: 'Geteilt mit Hostwerk – nur zum Ansehen.', until: 'Gültig bis', gone: 'Dieser Link existiert nicht oder ist abgelaufen.', gone2: 'Frag die Person, die ihn dir geschickt hat, nach einem neuen Link.' },
  en: { download: 'Download', password: 'Password', open: 'Open', locked: 'This file is protected with a password.', wrong: 'The password is not correct.', noPreview: 'There is no preview for this file.', noDownload: 'The sender did not allow downloads.', footer: 'Shared with Hostwerk – view only.', until: 'Valid until', gone: 'This link does not exist or has expired.', gone2: 'Ask the person who sent it to you for a new link.' },
};
const CSS = `:root{color-scheme:light dark;--bg:#f6f7f9;--panel:#fff;--text:#14161a;--muted:#5d6470;--line:#e2e5ea;--accent:#2563eb}
@media (prefers-color-scheme:dark){:root{--bg:#0f1115;--panel:#171a21;--text:#e8eaf0;--muted:#9aa1ad;--line:#2a2f3a;--accent:#6c9bff}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;min-height:100vh;display:flex;flex-direction:column}
header{display:flex;align-items:center;gap:12px;padding:12px 20px;background:var(--panel);border-bottom:1px solid var(--line)}
h1{font-size:16px;margin:0;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.muted{color:var(--muted);font-size:13px}
.btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;border-radius:8px;border:1px solid var(--line);background:var(--panel);color:var(--text);text-decoration:none;font:inherit;cursor:pointer;min-height:36px}
.btn.primary{background:var(--accent);border-color:var(--accent);color:#fff}.btn:focus-visible,input:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
main{flex:1;display:flex;flex-direction:column;padding:16px;gap:12px;min-height:0}
.frame{flex:1;min-height:70vh;border:1px solid var(--line);border-radius:12px;background:#fff;width:100%}
img.view{max-width:100%;height:auto;margin:0 auto;display:block;border-radius:8px;border:1px solid var(--line);background:#fff}
pre{white-space:pre-wrap;word-break:break-word;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px;margin:0;font:13px/1.55 ui-monospace,Consolas,monospace;max-height:75vh;overflow:auto}
.card{max-width:420px;margin:12vh auto;background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:24px;display:flex;flex-direction:column;gap:12px}
input{font:inherit;padding:9px 12px;border-radius:8px;border:1px solid var(--line);background:var(--bg);color:var(--text)}.err{color:#dc2626}
footer{padding:10px 20px;text-align:center}`;
const shell = (lang, title, head, body) => `<!doctype html><html lang="${lang}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><meta name="robots" content="noindex"><title>${esc(title)}</title><style>${CSS}</style></head><body>${head}${body}</body></html>`;

/**
 * Public page. o: { lang, share, base (URL of this share), mode: 'pdf'|'image'|'text'|'none', text?, size, error? }
 */
function page(o) {
  const t = T[o.lang] || T.en;
  const s = o.share;
  const dl = s.allowDownload ? `<a class="btn primary" href="${esc(o.base)}/download" download>${t.download}</a>` : '';
  const head = `<header><h1 title="${esc(s.name)}">${esc(s.name)}</h1><span class="muted">${esc(SIZE(o.size || 0))}</span>${dl}</header>`;
  let content;
  if (o.mode === 'pdf') content = `<iframe class="frame" src="${esc(o.base)}/${o.converted ? 'pdf' : 'file'}#view=FitH" title="${esc(s.name)}"></iframe>`;
  else if (o.mode === 'image') content = `<img class="view" src="${esc(o.base)}/file" alt="${esc(s.name)}">`;
  else if (o.mode === 'text') content = `<pre>${esc(o.text)}</pre>`;
  else content = `<div class="card"><strong>${t.noPreview}</strong>${s.allowDownload ? dl : `<span class="muted">${t.noDownload}</span>`}</div>`;
  const until = s.expires ? ` ${t.until} ${esc(new Date(s.expires).toLocaleDateString(o.lang === 'de' ? 'de-DE' : 'en-US'))}.` : '';
  return shell(o.lang, s.name, head, `<main>${content}</main><footer class="muted">${t.footer}${until}</footer>`);
}
function passwordPage(lang, share, wrong) {
  const t = T[lang] || T.en;
  return shell(lang, share.name, '', `<main><form class="card" method="post"><strong>${esc(share.name)}</strong><span class="muted">${t.locked}</span><label>${t.password}<br><input type="password" name="password" autocomplete="current-password" required autofocus style="width:100%"></label>${wrong ? `<span class="err" role="alert">${t.wrong}</span>` : ''}<button class="btn primary" type="submit">${t.open}</button></form></main>`);
}
function gonePage(lang) {
  const t = T[lang] || T.en;
  return shell(lang, t.gone, '', `<main><div class="card"><strong>${t.gone}</strong><span class="muted">${t.gone2}</span></div></main>`);
}
const langOf = (req) => (/^de\b/i.test(String(req.headers['accept-language'] || '')) ? 'de' : 'en');

module.exports = { newToken, hashPassword, checkPassword, expiryOf, expired, find, unlockValue, view, page, passwordPage, gonePage, langOf, esc, TOKEN };
