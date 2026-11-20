'use strict';
/*
 * Templates like Word/Excel/PowerPoint/LibreOffice (1.3):
 *   - real template formats: "Save as template" turns docx → dotx, docm → dotm, xlsx → xltx, xlsm → xltm, pptx → potx,
 *     pptm → potm, odt → ott, ods → ots, odp → otp, odg → otg by switching the package's main content type (OOXML:
 *     [Content_Types].xml) or the ODF mimetype (+ the manifest's root entry) – every other part stays byte for byte
 *     (zippatch). A new document from a template is the reverse (dotx → docx …, timestamps of the core properties reset).
 *   - where templates live: the own template folder (settings.templates.dir, default <library>/Vorlagen – "Templates"
 *     in English), folders the user imports read-only (settings.templates.folders – e.g. %APPDATA%\Microsoft\Templates,
 *     "Custom Office Templates", LibreOffice's template folder) and the templates Office 1.0–1.2 kept in its data folder
 *     (templates.json, still listed).
 *   - default template per app (settings.templates.defaults.docs|sheets|slides = template id): "New blank document"
 *     starts from it, like Word's Normal.dotm.
 * Template id: 'f' + sha1(absolute path) (stable while the file stays where it is) or the stored id of a 1.0 template.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { FORMATS, extOf } = require('./formats');
const { patchZip, entries, inflate } = require('./zippatch');
const { isZip } = require('./zip');
const P = require('./paths');

/** document ext → template ext (and templates map to themselves) */
const TEMPLATE_OF = {
  docx: 'dotx', docm: 'dotm', dotx: 'dotx', dotm: 'dotm', doc: 'dot', dot: 'dot',
  xlsx: 'xltx', xlsm: 'xltm', xltx: 'xltx', xltm: 'xltm', xls: 'xlt', xlt: 'xlt',
  pptx: 'potx', pptm: 'potm', ppsx: 'potx', ppsm: 'potm', potx: 'potx', potm: 'potm', ppt: 'pot', pot: 'pot',
  odt: 'ott', ott: 'ott', fodt: 'ott', ods: 'ots', ots: 'ots', fods: 'ots', odp: 'otp', otp: 'otp', fodp: 'otp', odg: 'otg', otg: 'otg',
};
/** template ext → document ext (macro templates keep their macros: dotm → docm …) */
const DOC_OF = { dotx: 'docx', dotm: 'docm', dot: 'doc', xltx: 'xlsx', xltm: 'xlsm', xlt: 'xls', potx: 'pptx', potm: 'pptm', pot: 'ppt', ott: 'odt', ots: 'ods', otp: 'odp', otg: 'odg' };
const isTemplateExt = (e) => Object.prototype.hasOwnProperty.call(DOC_OF, String(e || '').toLowerCase());
/** formats "Save as template" can write as a real template (binary 97–2003 and flat ODF are not converted) */
const CAN_SAVE = new Set(['docx', 'docm', 'dotx', 'dotm', 'xlsx', 'xlsm', 'xltx', 'xltm', 'pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'odt', 'ott', 'ods', 'ots', 'odp', 'otp', 'odg', 'otg']);

/** main part content types of the OOXML packages */
const MAIN = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  docm: 'application/vnd.ms-word.document.macroEnabled.main+xml',
  dotm: 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  xltm: 'application/vnd.ms-excel.template.macroEnabled.main+xml',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  potx: 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
  pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
  potm: 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
  ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
  ppsm: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
};
const MAIN_TYPES = new Set(Object.values(MAIN));
const ODF = new Set(['odt', 'ott', 'ods', 'ots', 'odp', 'otp', 'odg', 'otg']);
const fail = (msg, status = 400) => { const e = new Error(msg); e.status = status; return e; };
const esc = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Package bytes → the same package as `toExt` (template ↔ document). OOXML: the main part's content type in
 * [Content_Types].xml; ODF: `mimetype` (first, stored) and the manifest's "/" entry. `fresh`: a new document from a
 * template – created/modified of the core properties (OOXML) / meta.xml (ODF) become now.
 * Binary 97–2003 files and anything that is not a package come back unchanged.
 */
function convertPackage(buf, toExt, { fresh = false, now = new Date() } = {}) {
  const to = String(toExt || '').toLowerCase();
  buf = Buffer.isBuffer(buf) ? buf : Buffer.from(buf || []);
  if (!isZip(buf)) return buf;
  const list = entries(buf);
  const byName = new Map(list.map((e) => [e.name, e]));
  const changes = {};
  const iso = now.toISOString().replace(/\.\d{3}Z$/, 'Z');
  if (MAIN[to]) {
    const ct = byName.get('[Content_Types].xml');
    if (!ct) throw fail('This file is not an Office package.');
    const xml = inflate(ct).toString('utf8');
    let hit = false;
    const out = xml.replace(/<Override\b[^>]*>/g, (tag) => {
      const m = /ContentType="([^"]+)"/.exec(tag);
      if (!m || !MAIN_TYPES.has(m[1])) return tag;
      hit = true;
      return tag.replace(m[0], `ContentType="${MAIN[to]}"`);
    });
    if (!hit) throw fail('The main part of this package was not found.');
    if (out !== xml) changes['[Content_Types].xml'] = out;
    if (fresh) {
      const core = byName.get('docProps/core.xml');
      if (core) {
        const cx = inflate(core).toString('utf8');
        const nx = cx.replace(/(<dcterms:(created|modified)\b[^>]*>)[^<]*(<\/dcterms:\2>)/g, `$1${iso}$3`);
        if (nx !== cx) changes['docProps/core.xml'] = nx;
      }
    }
    return Object.keys(changes).length ? patchZip(buf, changes) : buf;
  }
  if (ODF.has(to)) {
    const mime = (FORMATS[to] || {}).mime;
    if (!byName.has('mimetype') && !byName.has('META-INF/manifest.xml')) throw fail('This file is not an OpenDocument package.');
    changes.mimetype = { data: mime, store: true };
    const mf = byName.get('META-INF/manifest.xml');
    if (mf) {
      const mx = inflate(mf).toString('utf8');
      const nx = mx.replace(/<manifest:file-entry\b[^>]*manifest:full-path="\/"[^>]*>/, (tag) => tag.replace(/manifest:media-type="[^"]*"/, `manifest:media-type="${mime}"`));
      if (nx !== mx) changes['META-INF/manifest.xml'] = nx;
    }
    if (fresh) {
      const meta = byName.get('meta.xml');
      if (meta) {
        const mx = inflate(meta).toString('utf8');
        const local = iso.replace(/Z$/, '');
        const nx = mx.replace(/(<meta:creation-date>)[^<]*(<\/meta:creation-date>)/, `$1${local}$2`).replace(/(<dc:date>)[^<]*(<\/dc:date>)/, `$1${local}$2`);
        if (nx !== mx) changes['meta.xml'] = nx;
      }
    }
    return patchZip(buf, changes, { first: ['mimetype'] });
  }
  return buf;
}
/** the main content type / ODF mimetype a package declares (tests, checks) */
function packageType(buf) {
  buf = Buffer.isBuffer(buf) ? buf : Buffer.from(buf || []);
  if (!isZip(buf)) return null;
  const list = entries(buf);
  const m = list.find((e) => e.name === 'mimetype');
  if (m) return inflate(m).toString('utf8').trim();
  const ct = list.find((e) => e.name === '[Content_Types].xml');
  if (!ct) return null;
  const xml = inflate(ct).toString('utf8');
  for (const t of MAIN_TYPES) if (new RegExp(`ContentType="${esc(t)}"`).test(xml)) return t;
  return null;
}

/* ---------------- where templates live ---------------- */
const KINDS = new Set(['docs', 'sheets', 'slides']);
const PLAIN = /^(txt|md|markdown|csv|tsv|tab|prn|html?|xhtml|json|xml)$/;
/** can this file be used as a template? (real template formats everywhere; inside a template folder also documents – e.g. Google Docs templates exported as .docx) */
function usable(name, inFolder) {
  const e = extOf(name);
  const f = FORMATS[e];
  if (!f || !KINDS.has(f.kind) || P.isJunk(name)) return false;
  if (isTemplateExt(e)) return true;
  return !!inFolder && f.edit !== 'view' && !PLAIN.test(e);
}
const idOf = (full) => `f${crypto.createHash('sha1').update(path.resolve(full).toLowerCase()).digest('hex').slice(0, 20)}`;

/** folders on this computer that usually hold templates (Settings → Templates → "Suggestions") */
function knownFolders(env = process.env, home = os.homedir(), platform = process.platform) {
  const out = [];
  const add = (p, label) => { if (p) out.push({ path: path.resolve(p), label }); };
  if (platform === 'win32') {
    const appdata = env.APPDATA || path.join(home, 'AppData', 'Roaming');
    add(path.join(appdata, 'Microsoft', 'Templates'), 'Microsoft Office – personal templates (Normal.dotm)');
    for (const docs of [path.join(home, 'Documents'), path.join(home, 'Dokumente'), path.join(home, 'OneDrive', 'Documents'), path.join(home, 'OneDrive', 'Dokumente')]) {
      add(path.join(docs, 'Custom Office Templates'), 'Custom Office Templates');
      add(path.join(docs, 'Benutzerdefinierte Office-Vorlagen'), 'Custom Office Templates');
    }
    add(path.join(appdata, 'LibreOffice', '4', 'user', 'template'), 'LibreOffice – my templates');
    for (const pf of [env.ProgramFiles, env['ProgramFiles(x86)']].filter(Boolean)) add(path.join(pf, 'LibreOffice', 'share', 'template', 'common'), 'LibreOffice – templates that come with it');
  } else if (platform === 'darwin') {
    add(path.join(home, 'Library', 'Group Containers', 'UBF8T346G9.Office', 'User Content', 'Templates'), 'Microsoft Office – personal templates');
    add(path.join(home, 'Library', 'Application Support', 'LibreOffice', '4', 'user', 'template'), 'LibreOffice – my templates');
    add('/Applications/LibreOffice.app/Contents/Resources/template/common', 'LibreOffice – templates that come with it');
  } else {
    add(path.join(home, '.config', 'libreoffice', '4', 'user', 'template'), 'LibreOffice – my templates');
    add(path.join(home, 'Templates'), 'Templates (home folder)');
    add(path.join(home, 'Vorlagen'), 'Templates (home folder)');
    add('/usr/lib/libreoffice/share/template/common', 'LibreOffice – templates that come with it');
  }
  const seen = new Set();
  return out.filter((x) => { const k = x.path.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; });
}

/** template files in a folder (and its sub folders, 3 levels – they become the categories) */
function scan(dir, { inFolder = true, max = 800 } = {}) {
  const out = [];
  const walk = (d, rel, depth) => {
    if (out.length >= max || depth > 3) return;
    let list;
    try { list = fs.readdirSync(d, { withFileTypes: true }); } catch { return; }
    for (const de of list) {
      if (out.length >= max) break;
      if (de.name.startsWith('.') || P.isJunk(de.name)) continue;
      const full = path.join(d, de.name);
      if (de.isDirectory()) { if (!/^(wizard|internal|~)/i.test(de.name)) walk(full, rel ? `${rel}/${de.name}` : de.name, depth + 1); continue; }
      if (!de.isFile() || !usable(de.name, inFolder)) continue;
      let st; try { st = fs.statSync(full); } catch { continue; }
      out.push({ full, rel: rel ? `${rel}/${de.name}` : de.name, sub: rel, st });
    }
  };
  walk(dir, '', 0);
  return out;
}

class TemplateStore {
  /** deps: { settings(), saveSettings(patch), legacy (Templates of library.js), lang() → 'de'|'en' } */
  constructor(deps) { this.d = deps; this.cache = null; }
  /** the own template folder (created on first use) */
  ownDir(create = false) {
    const s = this.d.settings();
    const set = (s.templates && s.templates.dir) || '';
    let dir = set;
    if (!dir) {
      const base = s.libraryDir;
      const de = path.join(base, 'Vorlagen'); const en = path.join(base, 'Templates');
      dir = fs.existsSync(de) ? de : fs.existsSync(en) ? en : (this.d.lang && this.d.lang() === 'en' ? en : de);
    }
    if (create) fs.mkdirSync(dir, { recursive: true });
    return dir;
  }
  folders() { const s = this.d.settings(); return ((s.templates && s.templates.folders) || []).filter((x) => typeof x === 'string' && path.isAbsolute(x)); }
  /** is this absolute path inside a template folder (own or imported)? */
  inTemplateFolder(full) { return [this.ownDir(), ...this.folders()].some((d) => P.inside(d, full)); }
  invalidate() { this.cache = null; }
  /** every template: own folder, imported folders, 1.0 store → [{ id, name, category, ext, kind, family, size, mtime, source, folder, readOnly, full, path?, default }] */
  list() {
    if (this.cache && Date.now() - this.cache.t < 3000) return this.cache.list;
    const s = this.d.settings();
    const defs = (s.templates && s.templates.defaults) || {};
    const lib = s.libraryDir;
    const out = [];
    const seen = new Set();
    const push = (full, st, extra) => {
      const id = idOf(full);
      if (seen.has(id)) return;
      seen.add(id);
      const e = extOf(full);
      const f = FORMATS[e] || {};
      const libRel = lib && P.inside(lib, full) ? path.relative(lib, full).split(path.sep).join('/') : null;
      out.push({ id, name: path.basename(full, path.extname(full)), ext: e, kind: f.kind || null, family: f.family || null, size: st.size, mtime: Math.floor(st.mtimeMs), created: Math.floor(st.birthtimeMs || st.ctimeMs), full, ...(libRel != null ? { path: libRel } : {}), template: isTemplateExt(e), ...extra });
    };
    const own = this.ownDir();
    for (const x of scan(own)) push(x.full, x.st, { source: 'own', category: x.sub || '', folder: own, readOnly: false });
    for (const dir of this.folders()) {
      for (const x of scan(dir)) push(x.full, x.st, { source: 'folder', category: x.sub || '', folder: dir, readOnly: true });
    }
    if (this.d.legacy) {
      for (const t of this.d.legacy.list()) {
        const p = this.d.legacy.path(t.id);
        let st; try { st = fs.statSync(p); } catch { continue; }
        out.push({ id: t.id, name: t.name, ext: t.ext, kind: t.kind, family: t.family, size: st.size, mtime: Math.floor(st.mtimeMs), created: t.created, full: p, template: isTemplateExt(t.ext), source: 'legacy', category: t.category && t.category !== 'My templates' ? t.category : '', folder: '', readOnly: false });
      }
    }
    for (const t of out) t.default = !!t.kind && defs[t.kind] === t.id;
    out.sort((a, b) => String(a.kind).localeCompare(String(b.kind)) || a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
    this.cache = { t: Date.now(), list: out };
    return out;
  }
  get(id) { const t = this.list().find((x) => x.id === String(id || '')); if (!t) throw fail('Template not found', 404); return t; }
  /** the template of `kind` that new blank documents start from, or null */
  defaultFor(kind) { const id = (((this.d.settings().templates || {}).defaults) || {})[kind]; if (!id) return null; return this.list().find((x) => x.id === id) || null; }

  /** bytes of a new DOCUMENT made from the template at `full` → { data, ext } */
  instantiate(full) {
    const e = extOf(full);
    const to = DOC_OF[e] || e;
    const data = fs.readFileSync(full);
    return { data: to === e ? data : convertPackage(data, to, { fresh: true }), ext: to };
  }
  /** bytes of the file as a real TEMPLATE → { data, ext } (throws for formats without a template form) */
  asTemplate(full) {
    const e = extOf(full);
    if (!CAN_SAVE.has(e)) throw fail('This format has no template form – save it as a Word, Excel, PowerPoint or OpenDocument file first.');
    const to = TEMPLATE_OF[e];
    const data = fs.readFileSync(full);
    return { data: to === e ? data : convertPackage(data, to), ext: to };
  }
  /** write a template into the own folder (category = sub folder) → absolute path */
  addOwn(data, ext, name, category = '') {
    const base = this.ownDir(true);
    const sub = String(category || '').split(/[\\/]/).map((x) => P.safeName(x, '')).filter(Boolean).slice(0, 2).join(path.sep);
    const dir = sub ? path.join(base, sub) : base;
    fs.mkdirSync(dir, { recursive: true });
    const n = P.uniqueName(`${P.safeName(name, 'Template')}.${ext}`, (x) => fs.existsSync(path.join(dir, x)));
    const full = path.join(dir, n);
    fs.writeFileSync(full, data, { flag: 'wx' });
    this.invalidate();
    return full;
  }
  rename(id, name) {
    const t = this.get(id);
    if (t.readOnly) throw fail('Templates from imported folders are read-only.', 403);
    const clean = P.safeName(String(name || '').trim(), '');
    if (!clean) throw fail('Please enter a name.');
    if (t.source === 'legacy') {
      const list = this.d.legacy.list().map((x) => (x.id === t.id ? { ...x, name: clean } : x));
      require('./store').writeJson(this.d.legacy.file, list);
      this.invalidate();
      return this.get(t.id);
    }
    const dir = path.dirname(t.full);
    const n = P.uniqueName(`${clean}.${t.ext}`, (x) => x.toLowerCase() !== path.basename(t.full).toLowerCase() && fs.existsSync(path.join(dir, x)));
    const to = path.join(dir, n);
    if (to !== t.full) fs.renameSync(t.full, to);
    // the default template follows its file
    const newId = idOf(to);
    const defs = { ...(((this.d.settings().templates || {}).defaults) || {}) };
    if (t.kind && defs[t.kind] === t.id && this.d.saveSettings) this.d.saveSettings({ templates: { defaults: { [t.kind]: newId } } });
    this.invalidate();
    return this.get(newId);
  }
}

module.exports = { TemplateStore, convertPackage, packageType, knownFolders, scan, usable, idOf, isTemplateExt, TEMPLATE_OF, DOC_OF, CAN_SAVE, MAIN };
