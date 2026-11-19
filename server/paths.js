'use strict';
/*
 * Path safety for the library (and for files from Hostwerk's roots):
 *   normRel('a\\b/../c')  → throws (no "..")      normRel('/a//b/') → 'a/b'
 *   checkName(name)       → valid file/folder name (no / \ : * ? " < > |, no control chars, no Windows reserved names,
 *                           no trailing dot/space, ≤ 200 chars) – throws an Error with status 400 otherwise
 *   inside(base, full)    → full is base or below it (case-insensitive on Windows)
 *   resolveIn(base, rel)  → absolute path below base (also checks symlinks of the existing part)
 *   uniqueName(name, exists) → "Report.docx" → "Report (2).docx" …
 */
const fs = require('fs');
const path = require('path');

const WIN = process.platform === 'win32';
const RESERVED = /^(con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³])(\..*)?$/i;
const BAD_CHARS = /[<>:"/\\|?*\u0000-\u001f]/;

const fail = (msg, status = 400) => Object.assign(new Error(msg), { status });

/** Valid single name (file or folder) – returns the trimmed name or throws */
function checkName(name) {
  const n = String(name == null ? '' : name).normalize('NFC').trim();
  if (!n) throw fail('Please enter a name.');
  if (n === '.' || n === '..') throw fail('This name is not allowed.');
  if (n.length > 200) throw fail('The name is too long (at most 200 characters).');
  if (BAD_CHARS.test(n)) throw fail('A name must not contain any of these characters: \\ / : * ? " < > |');
  if (RESERVED.test(n)) throw fail('This name is reserved by Windows – please choose another one.');
  if (/[. ]$/.test(n)) throw fail('A name must not end with a dot or a space.');
  return n;
}

/** Make a name safe instead of refusing it (for "Save as" of generated names) */
function safeName(name, fallback = 'Untitled') {
  let n = String(name == null ? '' : name).normalize('NFC').replace(/[<>:"/\\|?*\u0000-\u001f]+/g, '-').trim().replace(/[. ]+$/, '');
  if (n.length > 200) { const e = path.extname(n); n = n.slice(0, 200 - e.length) + e; }
  if (!n || n === '.' || n === '..') n = fallback;
  if (RESERVED.test(n)) n = `_${n}`;
  return n;
}

/** Normalised relative path ('' = root) – every segment checked */
function normRel(rel) {
  if (rel == null) return '';
  const s = String(rel);
  if (s.includes('\0')) throw fail('Invalid path');
  if (/^[a-zA-Z]:/.test(s) || /^\\\\/.test(s)) throw fail('Invalid path');
  const parts = s.replace(/\\/g, '/').split('/').filter((p) => p !== '' && p !== '.');
  for (const p of parts) {
    if (p === '..') throw fail('Invalid path', 403);
    checkName(p);
  }
  return parts.join('/');
}

const norm = (p) => (WIN ? path.resolve(p).toLowerCase() : path.resolve(p));
/** Is full the same as base or below it? */
function inside(base, full) {
  const rel = path.relative(norm(base), norm(full));
  return rel === '' || (!!rel && !rel.startsWith('..') && !path.isAbsolute(rel));
}

/** Absolute path of rel below base; throws 403 when it would leave base (also through symlinks/junctions) */
function resolveIn(base, rel) {
  const r = normRel(rel);
  const b = path.resolve(base);
  const full = path.resolve(b, ...r.split('/').filter(Boolean));
  if (!inside(b, full)) throw fail('Path outside the library', 403);
  // Symlinks: the longest existing part must stay inside the real base
  let probe = full;
  while (!fs.existsSync(probe) && probe !== b && inside(b, probe)) probe = path.dirname(probe);
  try {
    const realBase = fs.realpathSync(b);
    const realProbe = fs.realpathSync(probe);
    if (!inside(realBase, realProbe)) throw fail('Path outside the library', 403);
  } catch (e) { if (e.status) throw e; /* base missing yet – nothing to follow */ }
  return full;
}

/** "a/b/c.docx" → { dir: 'a/b', name: 'c.docx' } */
function splitRel(rel) {
  const r = normRel(rel);
  const i = r.lastIndexOf('/');
  return i < 0 ? { dir: '', name: r } : { dir: r.slice(0, i), name: r.slice(i + 1) };
}
const joinRel = (...parts) => parts.filter((p) => p !== '' && p != null).join('/');

/** Stem and extension, "Report.final.docx" → ['Report.final', '.docx']; folders/no ext → [name, ''] */
function stemExt(name, isDir = false) {
  const ext = isDir ? '' : path.extname(name);
  return ext && ext !== name ? [name.slice(0, -ext.length), ext] : [name, ''];
}

/**
 * First free name: exists(name) → bool. "Report.docx" → "Report (2).docx"; an existing " (n)" counter is continued.
 */
function uniqueName(name, exists, isDir = false) {
  if (!exists(name)) return name;
  const [stem0, ext] = stemExt(name, isDir);
  const m = /^(.*) \((\d+)\)$/.exec(stem0);
  const stem = m ? m[1] : stem0;
  for (let i = m ? Number(m[2]) + 1 : 2; i < 10000; i++) {
    const n = `${stem} (${i})${ext}`;
    if (!exists(n)) return n;
  }
  return `${stem}-${Date.now()}${ext}`;
}

/** Hidden or temporary office files that the library never lists (~$x.docx, .~lock.x#, .hidden, Thumbs.db) */
const isJunk = (name) => /^~\$|^\.~lock\.|^\.|^(thumbs\.db|desktop\.ini)$/i.test(name) || /\.tmp$/i.test(name);

module.exports = { checkName, safeName, normRel, inside, resolveIn, splitRel, joinRel, stemExt, uniqueName, isJunk, fail, WIN };
