/*
 * The settings object of ARCHITECTURE.md – defaults (used until the backend answers and for "Reset"), deep merge,
 * path get/set, diff and validation. Pure – unit-tested in tests/shell-settings.test.js.
 */
export const DEFAULT_SETTINGS = {
  libraryDir: '',
  general: { autosave: true, autosaveSeconds: 20, newFormat: 'ms', author: '', spellcheck: true, language: 'auto', confirmDelete: true, trashDays: 30, keepVersions: 30, versionDays: 90, view: 'grid', sort: 'modified', openFromFiles: true, recentCount: 20, units: 'cm' },
  docs: { font: 'Calibri', fontSize: 11, lineHeight: 1.15, paragraphSpacing: 8, pageSize: 'A4', orientation: 'portrait', margins: { top: 25, right: 25, bottom: 20, left: 25 }, ruler: true, zoom: 100, smartQuotes: true, autoLinks: true, showFormatting: false },
  sheets: { font: 'Calibri', fontSize: 11, gridlines: true, formulaBar: true, headings: true, autoCalc: true, zoom: 100, decimal: 'auto', dateFormat: 'auto', csvDelimiter: 'auto', csvEncoding: 'utf-8', r1c1: false, functionNames: 'auto' },
  slides: { size: '16:9', theme: 'hostwerk', font: 'Calibri', notes: true, presenterView: true, transition: 'none', loop: false, autoAdvance: 0 },
  pdf: { zoom: 'auto', sidebar: true, scroll: 'vertical', signatureName: '' },
  converters: { libreoffice: '', msoffice: true, preferred: 'auto', timeout: 120, autoEngine: 'ask' },
  integrations: { onlyoffice: { url: '', secret: '' }, defaultOpen: 'builtin', hostUrl: '', google: { clientId: '', clientSecret: '' }, osOpenWith: true },
  sharing: { enabled: true, defaultDays: 7, allowDownload: true },
  templates: { dir: '', saveTo: 'folder', folders: [], defaults: { docs: '', sheets: '', slides: '' } },
};

const isObj = (v) => v && typeof v === 'object' && !Array.isArray(v);
/** Deep merge: objects are merged, everything else (arrays, numbers …) replaced. Never mutates. */
export function mergeSettings(base, patch) {
  if (!isObj(patch)) return patch === undefined ? base : patch;
  const out = { ...(isObj(base) ? base : {}) };
  for (const [k, v] of Object.entries(patch)) out[k] = isObj(v) ? mergeSettings(out[k], v) : v;
  return out;
}
/** Settings from the backend completed with the defaults (older backends may miss newer fields) */
export const withDefaults = (s) => mergeSettings(DEFAULT_SETTINGS, s || {});
export const getIn = (obj, path) => String(path).split('.').reduce((o, k) => (o == null ? undefined : o[k]), obj);
/** Immutable set: setIn(s, 'docs.margins.top', 30) */
export function setIn(obj, path, value) {
  const [k, ...rest] = String(path).split('.');
  const cur = isObj(obj) ? obj : {};
  return { ...cur, [k]: rest.length ? setIn(cur[k], rest.join('.'), value) : value };
}
/** Paths whose values differ ('general.autosave', 'docs.margins.top' …) */
export function diffSettings(a, b, prefix = '') {
  const out = [];
  const keys = new Set([...Object.keys(isObj(a) ? a : {}), ...Object.keys(isObj(b) ? b : {})]);
  for (const k of keys) {
    const p = prefix ? `${prefix}.${k}` : k;
    const x = isObj(a) ? a[k] : undefined; const y = isObj(b) ? b[k] : undefined;
    if (isObj(x) || isObj(y)) out.push(...diffSettings(x, y, p));
    else if (JSON.stringify(x) !== JSON.stringify(y)) out.push(p);
  }
  return out;
}
/** Only the changed parts as a nested patch (PUT /office/settings merges) */
export function patchOf(original, draft) {
  let patch = {};
  for (const p of diffSettings(original, draft)) patch = setIn(patch, p, getIn(draft, p));
  return patch;
}

/** Number limits: path → [min, max] */
export const LIMITS = {
  'general.autosaveSeconds': [5, 600],
  'general.trashDays': [1, 365],
  'general.keepVersions': [1, 500],
  'general.versionDays': [1, 3650],
  'general.recentCount': [5, 100],
  'docs.fontSize': [6, 96],
  'docs.lineHeight': [0.8, 3],
  'docs.paragraphSpacing': [0, 72],
  'docs.margins.top': [0, 100], 'docs.margins.right': [0, 100], 'docs.margins.bottom': [0, 100], 'docs.margins.left': [0, 100],
  'docs.zoom': [25, 400],
  'sheets.fontSize': [6, 96],
  'sheets.zoom': [25, 400],
  'slides.autoAdvance': [0, 600],
  'converters.timeout': [10, 1800],
  'sharing.defaultDays': [0, 365],
};
const URLS = ['integrations.onlyoffice.url'];
/** { path: { text: 'English message', vars } } for every invalid field – empty object = fine (translate with t(text, vars)) */
export function validateSettings(s) {
  const err = {};
  for (const [p, [min, max]] of Object.entries(LIMITS)) {
    const v = getIn(s, p);
    if (v === undefined) continue;
    if (typeof v !== 'number' || !Number.isFinite(v)) err[p] = { text: 'Please enter a number.' };
    else if (v < min || v > max) err[p] = { text: 'Between {min} and {max}.', vars: { min, max } };
  }
  for (const p of URLS) {
    const v = String(getIn(s, p) || '').trim();
    if (v && !/^https?:\/\/[^\s/$.?#].[^\s]*$/i.test(v)) err[p] = { text: 'Enter an address starting with http:// or https://.' };
  }
  const gid = String(getIn(s, 'integrations.google.clientId') || '').trim();
  if (gid && !/^[\w.-]{10,200}\.apps\.googleusercontent\.com$/.test(gid)) err['integrations.google.clientId'] = { text: 'A client ID ends with .apps.googleusercontent.com.' };
  const lib = String(getIn(s, 'libraryDir') || '');
  if (lib && !/^([a-zA-Z]:[\\/]|\/|\\\\)/.test(lib)) err.libraryDir = { text: 'Enter a full path, e.g. D:\\Documents\\Office.' };
  return err;
}
/** Is this integration set up? (URL entered) */
export const integrationReady = (s, app) => !!String(getIn(s, `integrations.${app}.url`) || '').trim();
