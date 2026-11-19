/*
 * Pure helpers of the Office shell (no React, no Hostwerk modules) – unit-tested in tests/shell-*.test.js.
 * Texts returned here are English; the UI translates them with t().
 */
import { FORMATS, KINDS, EXPORTS, extOf } from '../lib/formats.js';

/* ---------------- Paths ---------------- */
export const basename = (p) => String(p || '').split('/').filter(Boolean).pop() || '';
export const dirname = (p) => { const a = String(p || '').split('/').filter(Boolean); a.pop(); return a.join('/'); };
export const joinPath = (dir, name) => [String(dir || '').replace(/^\/+|\/+$/g, ''), String(name || '').replace(/^\/+/, '')].filter(Boolean).join('/');
/** "Report 2024.docx" → { base: 'Report 2024', ext: 'docx' } (dot-files and names without dot have no extension) */
export function splitExt(name) {
  const s = String(name || '');
  const m = /^(.+)\.([A-Za-z0-9]{1,12})$/.exec(s);
  return m ? { base: m[1], ext: m[2].toLowerCase() } : { base: s, ext: '' };
}
/** Breadcrumbs of a library folder: '' → [], 'a/b' → [{ name:'a', dir:'a' }, { name:'b', dir:'a/b' }] */
export function crumbsOf(dir) {
  const parts = String(dir || '').split('/').filter(Boolean);
  return parts.map((name, i) => ({ name, dir: parts.slice(0, i + 1).join('/') }));
}
/** Is `inner` the same as or inside `outer`? (moving a folder into itself is not allowed) */
export const isInside = (inner, outer) => { const a = String(inner || ''); const b = String(outer || ''); return a === b || (b !== '' && a.startsWith(`${b}/`)) || b === ''; };
/** Key of a file for maps and selection: library path or "<root>|<path>" */
export const fileKey = (f) => (f && f.root ? `${f.root}|${f.path}` : (f && (f.id && f.deletedAt ? `trash:${f.id}` : f.path)) || '');

/* ---------------- Names ---------------- */
const RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;
/**
 * Checks a file or folder name. Returns null when fine, otherwise an English message (translate with t()).
 * existing: names in the same folder (case-insensitive, like Windows).
 */
export function validateName(name, { existing = [], current = '' } = {}) {
  const s = String(name || '');
  if (!s.trim()) return 'Please enter a name.';
  if (s.length > 200) return 'The name is too long (at most 200 characters).';
  if (/[<>:"/\\|?*\u0000-\u001f]/.test(s)) return 'These characters are not allowed: < > : " / \\ | ? *';
  if (s === '.' || s === '..') return 'This name is not allowed.';
  if (/[. ]$/.test(s)) return 'The name must not end with a dot or a space.';
  if (/^\s/.test(s)) return 'The name must not start with a space.';
  if (RESERVED.test(s)) return 'Windows reserves this name – please pick another one.';
  const low = s.toLowerCase();
  if (low !== String(current || '').toLowerCase() && existing.some((e) => String(e).toLowerCase() === low)) return 'A file with this name already exists here.';
  return null;
}
/** Adds/replaces the extension: ensureExt('Letter', 'docx') → 'Letter.docx', ensureExt('a.odt', 'docx') → 'a.docx' when replace */
export function ensureExt(name, ext, { replace = false } = {}) {
  const s = String(name || '').trim();
  if (!ext) return s;
  const cur = splitExt(s);
  if (cur.ext === ext.toLowerCase()) return s;
  if (replace && cur.ext && FORMATS[cur.ext]) return `${cur.base}.${ext}`;
  return `${s}.${ext}`;
}
/** Unique name in a folder: "Report.docx" → "Report (2).docx" when taken (case-insensitive). */
export function uniqueName(name, existing = []) {
  const taken = new Set(existing.map((e) => String(e).toLowerCase()));
  if (!taken.has(String(name).toLowerCase())) return name;
  const { base, ext } = splitExt(name);
  const stem = base.replace(/ \(\d+\)$/, '');
  for (let i = 2; i < 10000; i++) {
    const n = `${stem} (${i})${ext ? `.${ext}` : ''}`;
    if (!taken.has(n.toLowerCase())) return n;
  }
  return `${stem} (${Date.now()})${ext ? `.${ext}` : ''}`;
}
/** Default name of a new file: "Document", "Spreadsheet" … (English – translate, then ensureExt + uniqueName) */
export const defaultBaseName = (kind) => ({ docs: 'Document', sheets: 'Spreadsheet', slides: 'Presentation', pdf: 'PDF document' }[kind] || 'File');

/* ---------------- Entries: kind, sort, filter ---------------- */
export const kindOfEntry = (e) => (e && e.dir ? 'folder' : (e && e.kind) || (FORMATS[extOf(e && e.name)] || {}).kind || 'viewer');
export const familyOfEntry = (e) => (e && e.family) || (FORMATS[extOf(e && e.name)] || {}).family || 'other';
const KIND_ORDER = { folder: 0, docs: 1, sheets: 2, slides: 3, pdf: 4, viewer: 5 };
const collator = typeof Intl !== 'undefined' ? new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' }) : null;
const cmpName = (a, b) => (collator ? collator.compare(a.name || '', b.name || '') : String(a.name).localeCompare(String(b.name)));
export const SORTS = [
  { id: 'name', label: 'Name' },
  { id: 'modified', label: 'Modified' },
  { id: 'size', label: 'Size' },
  { id: 'type', label: 'Type' },
];
/**
 * Sorted copy; folders always first. key: name | modified | size | type, dir: asc | desc.
 * Ties fall back to the name, so the order is stable for equal values.
 */
export function sortEntries(list, key = 'name', dir = 'asc') {
  const m = dir === 'desc' ? -1 : 1;
  const by = {
    name: cmpName,
    modified: (a, b) => (a.mtime || 0) - (b.mtime || 0),
    size: (a, b) => (a.size || 0) - (b.size || 0),
    type: (a, b) => (KIND_ORDER[kindOfEntry(a)] - KIND_ORDER[kindOfEntry(b)]) || String(extOf(a.name)).localeCompare(String(extOf(b.name))),
  }[key] || cmpName;
  return [...(list || [])].sort((a, b) => {
    if (!!a.dir !== !!b.dir) return a.dir ? -1 : 1;
    return m * by(a, b) || cmpName(a, b);
  });
}
/** Lower case without accents – "Übersicht" matches "ubersicht" */
export const fold = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
/**
 * Instant filter of the listed entries: q (name, tags, note), kinds ['docs' …], families ['ms' …].
 * Folders pass the kind/family filter only when no filter is set (they hold files of every kind).
 */
export function filterEntries(list, { q = '', kinds = [], families = [] } = {}) {
  const words = fold(q).split(/\s+/).filter(Boolean);
  return (list || []).filter((e) => {
    if (kinds.length && (e.dir || !kinds.includes(kindOfEntry(e)))) return false;
    if (families.length && (e.dir || !families.includes(familyOfEntry(e)))) return false;
    if (!words.length) return true;
    const hay = fold([e.name, ...(e.tags || []), e.note || ''].join(' '));
    return words.every((w) => hay.includes(w));
  });
}
/** Counts per kind and family of a listing (for the filter chips) */
export function countBy(list) {
  const kinds = {}; const families = {};
  for (const e of list || []) { if (e.dir) continue; const k = kindOfEntry(e); const f = familyOfEntry(e); kinds[k] = (kinds[k] || 0) + 1; families[f] = (families[f] || 0) + 1; }
  return { kinds, families };
}
/**
 * Splits a text into parts with the search words marked: highlightParts('Annual report', 'rep') →
 * [{ text: 'Annual ', hit: false }, { text: 'rep', hit: true }, { text: 'ort', hit: false }]. Accent-insensitive.
 * HTML in snippets is removed first (the result is rendered as text, never as HTML).
 */
export function highlightParts(text, q) {
  const s = String(text || '').replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
  const words = fold(q).split(/\s+/).filter((w) => w.length > 0);
  if (!words.length || !s) return [{ text: s, hit: false }];
  const f = fold(s);
  // fold() keeps the length for common Latin letters (é → e); if it does not, mark nothing rather than wrong letters
  if (f.length !== s.length) return [{ text: s, hit: false }];
  const marks = new Array(s.length).fill(false);
  for (const w of words) { let i = f.indexOf(w); while (i >= 0) { for (let k = i; k < i + w.length; k++) marks[k] = true; i = f.indexOf(w, i + w.length); } }
  const out = [];
  for (let i = 0; i < s.length; i++) {
    const last = out[out.length - 1];
    if (last && last.hit === marks[i]) last.text += s[i]; else out.push({ text: s[i], hit: marks[i] });
  }
  return out;
}

/* ---------------- Selection ---------------- */
/**
 * Next selection after a click. order = keys in display order; mods = { ctrl, shift }.
 * Returns { sel: Set, anchor }.
 */
export function clickSelect(order, sel, anchor, key, { ctrl = false, shift = false } = {}) {
  if (shift && anchor != null && order.includes(anchor)) {
    const a = order.indexOf(anchor); const b = order.indexOf(key);
    const range = order.slice(Math.min(a, b), Math.max(a, b) + 1);
    const next = ctrl ? new Set([...sel, ...range]) : new Set(range);
    return { sel: next, anchor };
  }
  if (ctrl) {
    const next = new Set(sel);
    if (next.has(key)) next.delete(key); else next.add(key);
    return { sel: next, anchor: key };
  }
  return { sel: new Set([key]), anchor: key };
}
/** Arrow keys: index of the next item (grid: cols per row; list: cols = 1) */
export function moveIndex(i, key, count, cols = 1) {
  if (!count) return -1;
  if (i < 0) return 0;
  const d = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: cols, ArrowUp: -cols, Home: -Infinity, End: Infinity }[key];
  if (d === undefined) return i;
  if (d === -Infinity) return 0;
  if (d === Infinity) return count - 1;
  return Math.max(0, Math.min(count - 1, i + d));
}

/* ---------------- Formats: new files, conversion, saving ---------------- */
/** Extension of a new file of this kind with the user's choice (settings.general.newFormat 'ms' | 'odf') */
export const defaultExt = (kind, newFormat = 'ms') => ((KINDS[kind] || {}).newExt || {})[newFormat === 'odf' ? 'odf' : 'ms'] || '';
/**
 * Formats a new file of `kind` can be created in – the chunk writes them itself (edit: 'native').
 * The default format comes first, the other main family second, then the rest by family.
 */
export function formatChoices(kind, newFormat = 'ms') {
  const first = defaultExt(kind, newFormat);
  const second = defaultExt(kind, newFormat === 'odf' ? 'ms' : 'odf');
  const fam = { ms: 0, odf: 1, open: 2, apple: 3, google: 4, other: 5 };
  const list = Object.entries(FORMATS)
    .filter(([ext, f]) => f.kind === kind && f.edit === 'native' && !['markdown', 'htm'].includes(ext))
    .map(([ext, f]) => ({ ext, label: f.label, family: f.family }));
  const rank = (x) => (x.ext === first ? -2 : x.ext === second ? -1 : fam[x.family] ?? 9);
  return list.sort((a, b) => rank(a) - rank(b) || a.ext.localeCompare(b.ext)).map((x) => ({ ...x, isDefault: x.ext === first }));
}
/**
 * The nearest format the built-in editor writes for a file that is only imported (docm → docx, ott → odt, numbers → xlsx …).
 * Same kind; OpenDocument stays OpenDocument, everything else becomes Microsoft (or the open counterpart).
 */
export function nativeExtFor(ext) {
  const e = String(ext || '').toLowerCase();
  const f = FORMATS[e];
  if (!f) return '';
  if (f.edit === 'native') return e;
  if (f.edit === 'convert' && f.via) return f.via;
  if (['xhtml'].includes(e)) return 'html';
  if (['tab', 'prn'].includes(e)) return e === 'tab' ? 'tsv' : 'csv';
  return defaultExt(f.kind, f.family === 'odf' ? 'odf' : 'ms') || '';
}
/**
 * Targets for "Convert to …" / "Download as …" of a file: [{ ext, label, via: 'client' | 'server' }].
 * client = the editor chunk writes it (formats.js EXPORTS[kind].client, or the formats the chunk registered),
 * server = LibreOffice / Microsoft Office on this PC (only when one is installed). The file's own format is left out.
 */
export function convertTargets(ext, { hasConverter = false, client = null } = {}) {
  const f = FORMATS[String(ext || '').toLowerCase()];
  if (!f) return [];
  const ex = EXPORTS[f.kind];
  if (!ex) return hasConverter && f.via ? [{ ext: f.via, label: (FORMATS[f.via] || {}).label || f.via.toUpperCase(), via: 'server' }] : [];
  const own = String(ext).toLowerCase();
  const clientList = client || ex.client;
  const out = [];
  const seen = new Set([own]);
  for (const e of clientList) if (!seen.has(e)) { seen.add(e); out.push({ ext: e, label: (FORMATS[e] || {}).label || e.toUpperCase(), via: 'client' }); }
  if (hasConverter) for (const e of ex.server) if (!seen.has(e)) { seen.add(e); out.push({ ext: e, label: (FORMATS[e] || {}).label || e.toUpperCase(), via: 'server' }); }
  return out;
}
/**
 * How the editor host opens a file: { mode: 'edit' | 'import' | 'convert' | 'view' | 'unknown', kind, openExt, saveExt }
 *   openExt  format of the bytes the chunk gets (convert: the `via` format)
 *   saveExt  format the chunk's saver returns (import/convert: the nearest native format)
 */
export function openPlan(ext, { readOnly = false } = {}) {
  const e = String(ext || '').toLowerCase();
  const f = FORMATS[e];
  if (!f) return { mode: 'unknown', kind: null, openExt: e, saveExt: e };
  if (readOnly || f.edit === 'view') return { mode: 'view', kind: f.kind, openExt: e, saveExt: e };
  if (f.edit === 'convert') return { mode: 'convert', kind: f.kind, openExt: f.via || e, saveExt: f.via || e };
  if (f.edit === 'import') return { mode: 'import', kind: f.kind, openExt: e, saveExt: nativeExtFor(e) };
  return { mode: 'edit', kind: f.kind, openExt: e, saveExt: e };
}

/* ---------------- Query strings and addresses ---------------- */
export function parseQuery(s) {
  const out = {};
  for (const part of String(s || '').replace(/^[?#]/, '').split('&')) {
    if (!part) continue;
    const i = part.indexOf('=');
    const k = decodeURIComponent((i < 0 ? part : part.slice(0, i)).replace(/\+/g, ' '));
    const v = i < 0 ? '' : decodeURIComponent(part.slice(i + 1).replace(/\+/g, ' '));
    if (k) out[k] = v;
  }
  return out;
}
export const toQuery = (o) => Object.entries(o || {}).filter(([, v]) => v !== undefined && v !== null && v !== '' && v !== false).map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v === true ? 1 : v)}`).join('&');
/**
 * File of an `?action=open&target=<path>[&root=<root>]` address (quick search, AI links) → { root?, path }.
 * Older links put the root in front ("<root>:<path>"); root ids may contain ":" themselves (own:ab12, web:x), so the
 * root is everything up to the last ":" of the first path segment.
 */
export function openTarget(q) {
  const target = String((q && q.target) || '');
  if (q && q.root) return { root: String(q.root), path: target };
  const first = target.split('/')[0];
  const i = first.lastIndexOf(':');
  if (i > 0 && !/^[a-z]:$/i.test(first)) return { root: first.slice(0, i), path: target.slice(i + 1) };
  return { path: target };
}
/** Address of the editor host: /ext/office/edit?path=…[&root=…][&app=…] */
export const editRoute = (id, f, extra = {}) => `/ext/${encodeURIComponent(id || 'office')}/edit?${toQuery({ ...(f.root ? { root: f.root } : {}), path: f.path, ...extra })}`;

/* ---------------- Numbers, dates ---------------- */
export function fmtBytes(n) {
  const v = Number(n) || 0;
  if (v < 1024) return `${v} B`;
  const u = ['KB', 'MB', 'GB', 'TB'];
  let x = v / 1024; let i = 0;
  while (x >= 1024 && i < u.length - 1) { x /= 1024; i++; }
  const s = `${x >= 100 ? Math.round(x) : x >= 10 ? x.toFixed(1) : x.toFixed(1)} ${u[i]}`.replace('.0 ', ' ');
  // German interface: decimal comma like the rest of Hostwerk ("9,5 KB")
  const lang = typeof document !== 'undefined' && document.documentElement ? document.documentElement.lang || '' : '';
  return /^de/i.test(lang) ? s.replace('.', ',') : s;
}
/** Expiry timestamp of a share link: days = 0 → never (null) */
export const shareExpiry = (days, now = Date.now()) => (Number(days) > 0 ? now + Number(days) * 86400000 : null);
/** "12:03" today, "Yesterday 12:03" (key 'yesterday'), otherwise a date – returns { key, time, date } for the UI to phrase */
export function relDay(ts, now = Date.now()) {
  const d = new Date(ts); const n = new Date(now);
  const day = (x) => new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime();
  const diff = Math.round((day(n) - day(d)) / 86400000);
  return { key: diff === 0 ? 'today' : diff === 1 ? 'yesterday' : diff < 7 && diff > 0 ? 'week' : 'date', diff };
}
/** Tag list from user input: "a, b ,a" → ['a', 'b'] (trimmed, unique, at most 40 characters each) */
export function parseTags(input) {
  const seen = new Set(); const out = [];
  for (const raw of String(input || '').split(/[,;\n]/)) { const t = raw.trim().slice(0, 40); if (t && !seen.has(t.toLowerCase())) { seen.add(t.toLowerCase()); out.push(t); } }
  return out;
}
export const COLORS = ['', 'red', 'orange', 'yellow', 'green', 'blue', 'purple', 'gray'];

/* ---------------- Save state line ---------------- */
/**
 * What the top bar says about saving: { text, vars, tone } with English text.
 * state: idle | dirty | saving | saved | error | conflict | readonly | external
 */
export function saveStateText(state, { savedAt = null, error = '', app = '' } = {}) {
  switch (state) {
    case 'saving': return { text: 'Saving …', tone: 'muted' };
    case 'dirty': return { text: 'Unsaved changes', tone: 'warn' };
    case 'saved': return savedAt ? { text: 'Saved {time}', vars: { time: new Date(savedAt).toTimeString().slice(0, 5) }, tone: 'ok' } : { text: 'Saved', tone: 'ok' };
    case 'error': return { text: 'Not saved – {error}', vars: { error }, tone: 'err' };
    case 'conflict': return { text: 'Changed elsewhere', tone: 'err' };
    case 'readonly': return { text: 'Read only', tone: 'muted' };
    case 'external': return { text: 'Saved by {app}', vars: { app }, tone: 'muted' };
    default: return { text: 'No changes', tone: 'muted' };
  }
}
