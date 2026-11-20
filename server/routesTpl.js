'use strict';
/*
 * Template routes (1.3 – server/templates.js does the work):
 *   GET    /office/templates                    { templates: [...], dir, folders: [{ path, label, exists, count }], saveTo, defaults }
 *   GET    /office/templates/suggest            { folders: [{ path, label, count, added }] } – usual template folders on this computer
 *   POST   /office/templates {path|root+path, name, category?, to?: folder|next|both}  "Save as template" (real dotx/xltx/potx/ott …)
 *                                               → { template?, entries: [Entry] }
 *   GET    /office/templates/:id/raw            the template file itself
 *   GET    /office/templates/:id/doc            raw bytes of a NEW document made from it (header x-ext)
 *   POST   /office/templates/:id/new {root?, dir?, name?}  create that document (library root by default) → Entry
 *   POST   /office/templates/:id/rename {name}  own templates only
 *   POST   /office/templates/:id/default {on}   default template of its app (new blank documents start from it)
 *   DELETE /office/templates/:id                own: to the Office trash (library) / deleted; imported folders: 403
 *   GET    /office/templates/:id/thumb · PUT …   embedded or rendered thumbnail (rendered ones keyed tpl:<id>)
 *   POST   /office/fromtemplate {root?, path}   a template FILE (double-click in Files) → new document next to it
 *                                               (template folders: the library root), named after the template → Entry
 */
const fs = require('fs');
const path = require('path');
const X = require('./extract');
const { extOf, mimeOf } = require('./formats');
const P = require('./paths');
const T = require('./templates');

function templateRoutes(host, o, { raw, err, b }) {
  const { router: r } = host;
  const lib = o.library;
  const store = o.tplStore;
  const s = () => o.settings();
  const pub = (t) => { const { full, folder, ...rest } = t; return { ...rest, folderName: folder ? path.basename(folder) : '', full }; };
  const folderInfo = () => {
    const known = T.knownFolders();
    return store.folders().map((p) => {
      let exists = false; let count = 0;
      try { exists = fs.statSync(p).isDirectory(); } catch { exists = false; }
      if (exists) count = T.scan(p, { max: 400 }).length;
      const k = known.find((x) => x.path.toLowerCase() === p.toLowerCase());
      return { path: p, label: k ? k.label : path.basename(p), exists, count };
    });
  };
  /** where a new document from a template goes: next to it, unless it sits in a template folder (then the library root) */
  const createNear = (tplFull, where, name, data) => {
    if (where && where.dir != null) return lib.create(where.root ? { root: where.root, dir: where.dir } : where.dir, name, data);
    return lib.create('', name, data);
  };

  r.get('/office/templates', (req, res) => {
    const st = s().templates || {};
    res.json({ templates: store.list().map(pub), dir: store.ownDir(), folders: folderInfo(), saveTo: st.saveTo || 'folder', defaults: st.defaults || {} });
  });
  r.get('/office/templates/suggest', (req, res) => {
    const have = new Set(store.folders().map((x) => x.toLowerCase()));
    const out = [];
    for (const k of T.knownFolders()) {
      let ok = false; try { ok = fs.statSync(k.path).isDirectory(); } catch { ok = false; }
      if (!ok) continue;
      out.push({ ...k, count: T.scan(k.path, { max: 400 }).length, added: have.has(k.path.toLowerCase()) });
    }
    res.json({ folders: out });
  });
  r.post('/office/templates', (req, res) => {
    const body = req.body || {};
    const f = lib.file(b(req));
    const to = ['folder', 'next', 'both'].includes(body.to) ? body.to : (s().templates || {}).saveTo || 'folder';
    const { data, ext } = store.asTemplate(f.full);
    const name = String(body.name || '').trim() || path.basename(f.full, path.extname(f.full));
    const entries = [];
    let template = null;
    if (to === 'folder' || to === 'both') {
      const full = store.addOwn(data, ext, name, body.category || '');
      template = store.list().find((x) => x.full === full) || null;
      if (template && template.path != null) { try { entries.push(lib.entryOf({ path: template.path })); } catch { /* listed below */ } }
    }
    if (to === 'next' || to === 'both') {
      const dir = f.root ? { root: f.root, dir: path.posix.dirname(f.rel) === '.' ? '' : path.posix.dirname(f.rel) } : (path.posix.dirname(f.rel) === '.' ? '' : path.posix.dirname(f.rel));
      entries.push(lib.create(dir, `${P.safeName(name, 'Template')}.${ext}`, data));
      store.invalidate();
    }
    res.json({ template: template ? pub(template) : null, entries });
  });
  r.get('/office/templates/:id/raw', (req, res) => {
    const t = store.get(req.params.id);
    res.setHeader('Content-Type', mimeOf(t.full));
    res.setHeader('Cache-Control', 'no-store');
    fs.createReadStream(t.full).on('error', (e) => res.destroy(e)).pipe(res);
  });
  r.get('/office/templates/:id/doc', (req, res) => {
    const t = store.get(req.params.id);
    const { data, ext } = store.instantiate(t.full);
    res.setHeader('Content-Type', mimeOf(`x.${ext}`));
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('x-ext', ext);
    res.setHeader('Access-Control-Expose-Headers', 'x-ext');
    res.end(data);
  });
  r.post('/office/templates/:id/new', (req, res) => {
    const t = store.get(req.params.id);
    const body = req.body || {};
    const { data, ext } = store.instantiate(t.full);
    const name = `${P.safeName(String(body.name || '').trim() || t.name, 'Document')}.${ext}`;
    const where = body.dir != null || body.root ? { root: body.root ? String(body.root) : null, dir: String(body.dir || '') } : null;
    res.json(createNear(t.full, where, name, data));
  });
  r.post('/office/templates/:id/rename', (req, res) => res.json(pub(store.rename(req.params.id, (req.body || {}).name))));
  r.post('/office/templates/:id/default', (req, res) => {
    const t = store.get(req.params.id);
    if (!t.kind) throw err(400, 'This template has no app.');
    const on = (req.body || {}).on !== false;
    const cur = ((s().templates || {}).defaults || {})[t.kind];
    o.saveSettings({ templates: { defaults: { [t.kind]: on ? t.id : (cur === t.id ? '' : cur || '') } } });
    store.invalidate();
    res.json(pub(store.get(t.id)));
  });
  r.delete('/office/templates/:id', (req, res) => {
    const t = store.get(req.params.id);
    if (t.readOnly) throw err(403, 'Templates from imported folders are read-only.');
    if (t.source === 'legacy') o.templates.remove(t.id);
    else if (t.path != null) lib.toTrash([t.path]);
    else fs.rmSync(t.full, { force: true });
    const defs = (s().templates || {}).defaults || {};
    if (t.kind && defs[t.kind] === t.id) o.saveSettings({ templates: { defaults: { [t.kind]: '' } } });
    store.invalidate();
    res.json({ ok: true });
  });
  // thumbnails: embedded (Office/LibreOffice save one) or one the UI rendered (key tpl:<id>, per mtime)
  r.get('/office/templates/:id/thumb', (req, res) => {
    const t = store.get(req.params.id);
    let th = null;
    try { if (t.size <= 64 * 1048576) th = X.thumb(fs.readFileSync(t.full), t.ext); } catch { th = null; }
    if (!th && o.thumbs) th = o.thumbs.get(`tpl:${t.id}`, t.mtime);
    if (!th) { if (req.query.probe) return res.status(204).end(); throw err(404, 'No thumbnail'); }
    res.setHeader('Content-Type', th.type);
    res.setHeader('Cache-Control', 'private, max-age=60');
    res.end(th.data);
  });
  r.put('/office/templates/:id/thumb', raw, (req, res) => {
    const t = store.get(req.params.id);
    const data = Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0);
    const isPng = data.length >= 33 && data.readUInt32BE(0) === 0x89504e47;
    const isWebp = data.length >= 16 && data.toString('latin1', 0, 4) === 'RIFF' && data.toString('latin1', 8, 12) === 'WEBP';
    if (!(isPng || isWebp) || data.length > 2 * 1048576) throw err(400, 'A thumbnail must be a PNG or WebP image up to 2 MB.');
    if (!o.thumbs) throw err(501, 'Thumbnails are not stored here.');
    o.thumbs.put(`tpl:${t.id}`, t.mtime, data, isPng ? 'image/png' : 'image/webp');
    res.json({ ok: true });
  });
  r.post('/office/fromtemplate', (req, res) => {
    const f = lib.file(b(req));
    const e = extOf(f.full);
    if (!T.usable(path.basename(f.full), true)) throw err(400, 'This file cannot be used as a template.');
    const { data, ext } = store.instantiate(f.full);
    const name = `${P.safeName(String((req.body || {}).name || '').trim() || path.basename(f.full, path.extname(f.full)), 'Document')}.${ext}`;
    const parent = path.posix.dirname(f.rel) === '.' ? '' : path.posix.dirname(f.rel);
    // a template folder (own/imported) is no place for documents – those go to the library root
    const where = store.inTemplateFolder(f.full) ? null : { root: f.root, dir: parent };
    let entry;
    try { entry = createNear(f.full, where, name, data); } catch (x) {
      if (!where || x.status === 400) throw x;
      entry = createNear(f.full, null, name, data); // read-only folder → library
    }
    res.json({ ...entry, from: e });
  });
}

module.exports = { templateRoutes };
