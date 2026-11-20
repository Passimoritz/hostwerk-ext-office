'use strict';
/*
 * Settings of the Office extension: defaults exactly as in ARCHITECTURE.md, deep merge of a patch with validation.
 *   defaults(libraryDir) → full object
 *   merge(current, patch) → { settings, errors: [key…] }  – invalid values are not taken (errors lists them)
 *   normalize(stored, libraryDir) → stored values on top of the defaults, everything invalid dropped
 *                                   (migration of removed values – see migrate())
 */
const path = require('path');

function defaults(libraryDir = '') {
  return {
    libraryDir,
    general: { autosave: true, autosaveSeconds: 20, newFormat: 'ms', author: '', spellcheck: true, language: 'auto',
      confirmDelete: true, trashDays: 30, keepVersions: 30, versionDays: 90, view: 'grid', sort: 'modified',
      openFromFiles: true, recentCount: 20, units: 'cm' },
    docs: { font: 'Calibri', fontSize: 11, lineHeight: 1.15, paragraphSpacing: 8, pageSize: 'A4', orientation: 'portrait',
      margins: { top: 25, right: 25, bottom: 20, left: 25 }, ruler: true, zoom: 100, smartQuotes: true, autoLinks: true, showFormatting: false },
    sheets: { font: 'Calibri', fontSize: 11, gridlines: true, formulaBar: true, headings: true, autoCalc: true, zoom: 100,
      decimal: 'auto', dateFormat: 'auto', csvDelimiter: 'auto', csvEncoding: 'utf-8', r1c1: false, functionNames: 'auto' },
    slides: { size: '16:9', theme: 'hostwerk', font: 'Calibri', notes: true, presenterView: true, transition: 'none', loop: false, autoAdvance: 0 },
    pdf: { zoom: 'auto', sidebar: true, scroll: 'vertical', signatureName: '' },
    // autoEngine (Full fidelity): private LibreOffice in <data>/office-engine – ask once, set up silently, or never
    converters: { libreoffice: '', msoffice: true, preferred: 'auto', timeout: 120, autoEngine: 'ask' },
    // hostUrl (addition, see "Open points"): address under which ONLYOFFICE reaches Hostwerk ('' = from the request)
    // google: own OAuth client (Desktop app) for Google Drive – see GET /office/google/state for the guide
    integrations: { onlyoffice: { url: '', secret: '' }, defaultOpen: 'builtin', hostUrl: '', google: { clientId: '', clientSecret: '' }, osOpenWith: true },
    sharing: { enabled: true, defaultDays: 7, allowDownload: true },
    // 1.3 templates: dir '' = <library>/Vorlagen · saveTo: where "Save as template" writes · folders: imported template
    // folders (read-only, absolute paths) · defaults: template id per app that new blank documents start from
    templates: { dir: '', saveTo: 'folder', folders: [], defaults: { docs: '', sheets: '', slides: '' } },
  };
}

const ENUMS = {
  'general.newFormat': ['ms', 'odf'],
  'general.view': ['grid', 'list'],
  'general.sort': ['name', 'modified', 'size', 'kind', 'type', 'created', '-name', '-modified', '-size', '-kind', '-type', '-created'],
  'general.units': ['cm', 'in'],
  'docs.pageSize': ['A4', 'Letter', 'A5', 'Legal'],
  'docs.orientation': ['portrait', 'landscape'],
  'sheets.decimal': ['auto', '.', ','],
  'sheets.functionNames': ['auto', 'en', 'de'],
  'slides.size': ['16:9', '4:3', 'A4'],
  'slides.transition': ['none', 'fade', 'slide'],
  'pdf.scroll': ['vertical', 'horizontal', 'page'],
  'converters.preferred': ['auto', 'libreoffice', 'msoffice'],
  'converters.autoEngine': ['ask', 'auto', 'off'],
  'integrations.defaultOpen': ['builtin', 'onlyoffice', 'desktop'],
  'templates.saveTo': ['folder', 'next', 'both'],
};
const RANGES = {
  'general.autosaveSeconds': [3, 3600], 'general.trashDays': [1, 3650], 'general.keepVersions': [0, 1000], 'general.versionDays': [1, 3650],
  'general.recentCount': [1, 200],
  'docs.fontSize': [4, 400], 'docs.lineHeight': [0.5, 5], 'docs.paragraphSpacing': [0, 200], 'docs.zoom': [10, 500],
  'docs.margins.top': [0, 200], 'docs.margins.right': [0, 200], 'docs.margins.bottom': [0, 200], 'docs.margins.left': [0, 200],
  'sheets.fontSize': [4, 400], 'sheets.zoom': [10, 500],
  'slides.autoAdvance': [0, 3600],
  'converters.timeout': [10, 3600],
  'sharing.defaultDays': [0, 3650],
};
/** Custom checks for strings (value → ok) */
const URL_OK = (v) => v === '' || /^https?:\/\/[^\s"'<>]+$/i.test(v);
const STRINGS = {
  'general.language': (v) => /^(auto|[a-z]{2,3}(-[A-Za-z]{2,4})?)$/.test(v),
  'converters.libreoffice': (v) => v === '' || path.isAbsolute(v),
  'integrations.onlyoffice.url': URL_OK,
  'integrations.hostUrl': URL_OK,
  'pdf.zoom': (v) => ['auto', 'page-fit', 'page-width'].includes(v) || /^\d{2,3}$/.test(v),
  'integrations.google.clientId': (v) => v === '' || /^[\w.-]{10,200}\.apps\.googleusercontent\.com$/.test(v),
  'integrations.google.clientSecret': (v) => /^[\w-]{0,100}$/.test(v),
  'slides.theme': (v) => /^[a-z0-9_-]{1,40}$/i.test(v),
  'templates.dir': (v) => v === '' || (path.isAbsolute(v) && !v.includes('\0')),
  'templates.defaults.docs': (v) => /^[\w-]{0,64}$/.test(v),
  'templates.defaults.sheets': (v) => /^[\w-]{0,64}$/.test(v),
  'templates.defaults.slides': (v) => /^[\w-]{0,64}$/.test(v),
};
/** Lists of absolute folder paths (templates.folders): at most 20, no duplicates */
function checkFolders(v) {
  if (!Array.isArray(v) || v.length > 20) return [false];
  const out = [];
  for (const x of v) {
    if (typeof x !== 'string') return [false];
    const p = x.trim();
    if (!p || p.length > 1000 || p.includes('\0') || !path.isAbsolute(p)) return [false];
    const r = path.resolve(p);
    if (!out.some((y) => y.toLowerCase() === r.toLowerCase())) out.push(r);
  }
  return [true, out];
}

const isObj = (x) => !!x && typeof x === 'object' && !Array.isArray(x);

/** One value against its default → [ok, value] */
function check(key, def, v) {
  if (key === 'pdf.zoom' && typeof v === 'number') v = String(Math.round(v)); // zoom may also be a number
  if (typeof def === 'boolean') return typeof v === 'boolean' ? [true, v] : [false];
  if (typeof def === 'number') {
    const n = typeof v === 'string' && v.trim() !== '' ? Number(v) : v;
    if (typeof n !== 'number' || !Number.isFinite(n)) return [false];
    const r = RANGES[key];
    if (r && (n < r[0] || n > r[1])) return [false];
    return [true, Math.round(n * 100) / 100];
  }
  if (typeof def === 'string') {
    if (typeof v !== 'string') return [false];
    const s = key === 'integrations.onlyoffice.secret' ? v : v.trim();
    if (s.length > 1000) return [false];
    if (ENUMS[key] && !ENUMS[key].includes(s)) return [false];
    if (STRINGS[key] && !STRINGS[key](s)) return [false];
    return [true, /url$/i.test(key) ? s.replace(/\/+$/, '') : key === 'templates.dir' && s ? path.resolve(s) : s];
  }
  return [false];
}

/** Deep merge patch into base (copy) – only keys that exist in the defaults; errors collects rejected keys */
function mergeInto(base, def, patch, prefix, errors) {
  const out = { ...base };
  for (const [k, v] of Object.entries(patch || {})) {
    if (!Object.prototype.hasOwnProperty.call(def, k)) continue; // unknown key → ignored
    const key = prefix ? `${prefix}.${k}` : k;
    if (isObj(def[k])) {
      if (!isObj(v)) { errors.push(key); continue; }
      out[k] = mergeInto(isObj(base[k]) ? base[k] : def[k], def[k], v, key, errors);
      continue;
    }
    if (Array.isArray(def[k])) {
      const [ok, val] = key === 'templates.folders' ? checkFolders(v) : [false];
      if (ok) out[k] = val; else errors.push(key);
      continue;
    }
    if (key === 'libraryDir') {
      if (typeof v === 'string' && v.trim() && path.isAbsolute(v.trim()) && !v.includes('\0')) out[k] = path.resolve(v.trim());
      else errors.push(key);
      continue;
    }
    const [ok, val] = check(key, def[k], v);
    if (ok) out[k] = val; else errors.push(key);
  }
  return out;
}

/**
 * Older stored settings → current shape. ONLYOFFICE is the only full editor now: a stored defaultOpen that named a
 * full editor which no longer exists becomes 'onlyoffice' (stored values were always validated, so any other string
 * can only be that removed editor); its old integrations.<app> key is dropped (keys missing from the defaults are
 * never taken, so it disappears from the loaded settings and from the next save).
 */
function migrate(stored) {
  if (!isObj(stored) || !isObj(stored.integrations)) return stored;
  const integrations = { ...stored.integrations };
  const d = integrations.defaultOpen;
  if (typeof d === 'string' && d && !ENUMS['integrations.defaultOpen'].includes(d)) integrations.defaultOpen = 'onlyoffice';
  return { ...stored, integrations };
}

/** Stored settings on top of the defaults (invalid stored values fall back to the default) */
function normalize(stored, libraryDir) {
  const def = defaults(libraryDir);
  const s = mergeInto(def, def, isObj(stored) ? migrate(stored) : {}, '', []);
  if (!s.libraryDir) s.libraryDir = libraryDir;
  return s;
}

/** PUT: current + patch → { settings, errors } */
function merge(current, patch, libraryDir) {
  const def = defaults(libraryDir || (current && current.libraryDir) || '');
  const errors = [];
  const settings = mergeInto(normalize(current, def.libraryDir), def, isObj(patch) ? patch : {}, '', errors);
  return { settings, errors };
}

module.exports = { defaults, merge, normalize, migrate, ENUMS, RANGES };
