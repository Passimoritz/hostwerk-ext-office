'use strict';
/*
 * Full editor in the browser: helpers for the ONLYOFFICE Document Server (the only full editor – see ARCHITECTURE.md).
 *   tokens   access tokens for the public file/callback routes are HMAC-SHA256 signed JSON { k: file key, r: root,
 *            p: path, w: can write, exp } – short-lived, checked on every public call (no session, no cookie)
 *   jwt      HS256 sign/verify for ONLYOFFICE (config + callback) with the configured secret
 */
const crypto = require('crypto');
const http = require('http');
const https = require('https');

const b64u = (b) => Buffer.from(b).toString('base64url');
const fromB64u = (s) => Buffer.from(String(s || ''), 'base64url');
const hmac = (secret, data) => crypto.createHmac('sha256', secret).update(data).digest();
const safeEq = (a, b) => a.length === b.length && crypto.timingSafeEqual(a, b);

/** Signed token for a file: payload { k, r, p, w } + exp (ms) */
function signToken(secret, payload, ttlMs = 10 * 3600000, now = Date.now()) {
  const body = b64u(JSON.stringify({ ...payload, exp: now + ttlMs }));
  return `${body}.${b64u(hmac(secret, body))}`;
}
/** → payload or null (bad signature, expired, wrong file id) */
function verifyToken(secret, token, { now = Date.now(), fileId } = {}) {
  const m = /^([A-Za-z0-9_-]+)\.([A-Za-z0-9_-]+)$/.exec(String(token || ''));
  if (!m) return null;
  if (!safeEq(fromB64u(m[2]), hmac(secret, m[1]))) return null;
  let p; try { p = JSON.parse(fromB64u(m[1]).toString('utf8')); } catch { return null; }
  if (!p || typeof p.exp !== 'number' || p.exp < now) return null;
  if (fileId != null && fileIdOf(p.r, p.p) !== fileId) return null;
  return p;
}
/** File id in the public URLs: base64url of "root\npath" (root '' = library) */
const fileIdOf = (root, p) => b64u(`${root || ''}\n${p || ''}`);
function parseFileId(id) {
  if (!/^[A-Za-z0-9_-]{1,2000}$/.test(String(id || ''))) return null;
  const s = fromB64u(id).toString('utf8');
  const i = s.indexOf('\n');
  return i < 0 ? null : { root: s.slice(0, i) || null, path: s.slice(i + 1) };
}

/* ---------------- JWT (HS256) ---------------- */
function jwtSign(payload, secret) {
  const head = b64u(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  const body = b64u(JSON.stringify(payload));
  return `${head}.${body}.${b64u(hmac(secret, `${head}.${body}`))}`;
}
function jwtVerify(token, secret, now = Date.now()) {
  const m = /^([A-Za-z0-9_-]+)\.([A-Za-z0-9_-]+)\.([A-Za-z0-9_-]+)$/.exec(String(token || ''));
  if (!m || !secret) return null;
  let head; try { head = JSON.parse(fromB64u(m[1]).toString('utf8')); } catch { return null; }
  if (!head || head.alg !== 'HS256') return null;
  if (!safeEq(fromB64u(m[3]), hmac(secret, `${m[1]}.${m[2]}`))) return null;
  let p; try { p = JSON.parse(fromB64u(m[2]).toString('utf8')); } catch { return null; }
  if (p && typeof p.exp === 'number' && p.exp * 1000 < now) return null;
  return p;
}

/** GET a URL (http/https) → { status, headers, body: Buffer } with size limit and timeout */
function fetchUrl(url, { timeout = 20000, max = 512 * 1048576, headers = {} } = {}) {
  return new Promise((resolve, reject) => {
    let u; try { u = new URL(url); } catch { reject(new Error('Invalid URL')); return; }
    if (!/^https?:$/.test(u.protocol)) { reject(new Error('Invalid URL')); return; }
    const lib = u.protocol === 'https:' ? https : http;
    const req = lib.get(u, { timeout, headers: { 'user-agent': 'Hostwerk-Office', ...headers } }, (res) => {
      const parts = []; let n = 0;
      res.on('data', (c) => { n += c.length; if (n > max) { req.destroy(new Error('Response too large')); return; } parts.push(c); });
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: Buffer.concat(parts) }));
      res.on('error', reject);
    });
    req.on('timeout', () => req.destroy(new Error('Timeout')));
    req.on('error', reject);
  });
}

/** ONLYOFFICE document type for an extension */
function ooType(ext) {
  const kind = (require('./formats').FORMATS[ext] || {}).kind;
  if (kind === 'sheets') return 'cell';
  if (kind === 'slides') return 'slide';
  if (kind === 'pdf') return 'pdf';
  return 'word';
}
/** ONLYOFFICE document key: unique per file version, [0-9a-zA-Z.=_-], ≤ 128 */
const ooKey = (fileKey, mtime) => crypto.createHash('sha256').update(`${fileKey}\n${Math.floor(mtime || 0)}`).digest('base64url').slice(0, 40);

module.exports = { signToken, verifyToken, fileIdOf, parseFileId, jwtSign, jwtVerify, fetchUrl, ooType, ooKey };
