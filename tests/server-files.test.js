'use strict';
/*
 * Office inside Hostwerk's Files (SDK 3.15): the library as a root (paths of the root are library paths), files in any
 * root (create with free names, rename keeps metadata/versions), and "Open with" (osreg.js) against Hostwerk's own
 * core/filetypes.js with a SIMULATED registry – the real Windows registry is never touched. Without a Hostwerk checkout
 * (GitHub CI) the core part runs against a small stand-in of host.fileTypes instead.
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const SET = require('../server/settings');
const { Meta } = require('../server/meta');
const { Versions } = require('../server/versions');
const { Trash } = require('../server/trash');
const { Library } = require('../server/library');
const { OsReg, osTypes, definition, kindStates, OS_EXTS, REG_ID, APP_NAME } = require('../server/osreg');
const { FORMATS } = require('../server/formats');

const made = [];
const tmp = () => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'office-files-')); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });
const coreDir = [process.env.HOSTWERK_DIR, path.join(__dirname, '..', '..', '..', 'hostwerk')].filter(Boolean).find((d) => fs.existsSync(path.join(d, 'server', 'core', 'filetypes.js')));

/** Hostwerk's files core in small: roots → folders, resolve() refuses leaving the root */
function filesCore(roots) {
  return {
    resolve(root, rel) {
      const base = roots[root];
      if (!base) throw Object.assign(new Error('Unbekannter Bereich'), { status: 400 });
      const full = path.resolve(base, `.${path.sep}${String(rel || '')}`);
      if (full !== path.resolve(base) && !full.startsWith(path.resolve(base) + path.sep)) throw Object.assign(new Error('outside'), { status: 403 });
      return { base, full };
    },
  };
}
function mkLib(dir) {
  const libDir = path.join(dir, 'office');
  const other = path.join(dir, 'web');
  fs.mkdirSync(libDir, { recursive: true }); fs.mkdirSync(other, { recursive: true });
  const settings = SET.normalize({}, libDir);
  const meta = new Meta(path.join(dir, 'meta.json'));
  const versions = new Versions(path.join(dir, 'versions'));
  const lib = new Library({ settings: () => settings, files: filesCore({ 'ext:office:library': libDir, 'own:ab12': other }), meta, versions, trash: new Trash(path.join(dir, 'trash')), index: null, sharedKeys: () => new Set() });
  return { lib, meta, versions, libDir, other };
}

test('Files: a file of the "Office documents" root is the same library file (favourites, versions)', () => {
  const { lib, meta, libDir } = mkLib(tmp());
  fs.mkdirSync(path.join(libDir, 'Work'));
  fs.writeFileSync(path.join(libDir, 'Work', 'Plan.docx'), 'x');
  const l = lib.loc({ root: 'ext:office:library', path: 'Work/Plan.docx' });
  assert.deepStrictEqual([l.key, l.root, l.rel], ['Work/Plan.docx', null, 'Work/Plan.docx']);
  meta.set('Work/Plan.docx', { favorite: true });
  const e = lib.entryOf({ root: 'ext:office:library', path: 'Work/Plan.docx' });
  assert.strictEqual(e.root, undefined, 'no root – a library entry');
  assert.strictEqual(e.favorite, true);
  assert.throws(() => lib.loc({ root: 'ext:office:library', path: '../x.docx' }), (x) => x.status === 403);
});

test('Files: create in any root with a free name, rename keeps metadata and versions', () => {
  const { lib, meta, versions, other } = mkLib(tmp());
  fs.mkdirSync(path.join(other, 'docs'));
  const a = lib.create({ root: 'own:ab12', dir: 'docs' }, 'Document.docx', Buffer.from('a'));
  const b = lib.create({ root: 'own:ab12', dir: 'docs' }, 'Document.docx', Buffer.from('b'));
  assert.deepStrictEqual([a.root, a.path, b.name], ['own:ab12', 'docs/Document.docx', 'Document (2).docx']);
  assert.strictEqual(fs.readFileSync(path.join(other, 'docs', 'Document.docx'), 'utf8'), 'a', 'never overwritten');
  meta.set('@own:ab12/docs/Document.docx', { favorite: true, tags: ['Q4'] });
  lib.write({ root: 'own:ab12', path: 'docs/Document.docx' }, Buffer.from('a2'), {});
  assert.strictEqual(versions.list('@own:ab12/docs/Document.docx').length, 1);
  const r = lib.rename({ root: 'own:ab12', path: 'docs/Document.docx' }, 'Offer.docx');
  assert.deepStrictEqual([r.root, r.path, r.favorite, r.tags], ['own:ab12', 'docs/Offer.docx', true, ['Q4']]);
  assert.strictEqual(versions.list('@own:ab12/docs/Offer.docx').length, 1, 'versions follow the rename');
  assert.ok(fs.existsSync(path.join(other, 'docs', 'Offer.docx')));
  assert.throws(() => lib.rename({ root: 'own:ab12', path: 'docs/Offer.docx' }, 'Document (2).docx'), (x) => x.status === 409);
});

test('"Open with": the type list is exactly what Office opens, every type with MIME and kind', () => {
  const types = osTypes();
  assert.strictEqual(types.length, OS_EXTS.length);
  for (const x of types) {
    assert.ok(FORMATS[x.ext], x.ext);
    assert.ok(/^[a-z0-9][\w.+-]{0,60}\/[\w.+-]{1,100}$/i.test(x.mime), `${x.ext}: ${x.mime}`);
    assert.ok(['docs', 'sheets', 'slides', 'pdf', 'viewer'].includes(x.kind), x.ext);
    assert.ok(/^[^\u0000-\u001f"%\\/]{2,80}$/.test(x.progIdLabel), x.progIdLabel);
  }
  for (const e of ['docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'pdf', 'doc', 'xls', 'ppt', 'rtf', 'csv', 'md', 'epub', 'vsdx', 'eml', 'msg', 'vcf', 'ics']) assert.ok(OS_EXTS.includes(e), e);
  for (const e of ['txt', 'html', 'exe', 'js', 'json']) assert.ok(!OS_EXTS.includes(e), `never ${e}`);
  const d = definition((s) => `DE ${s}`);
  assert.deepStrictEqual([d.id, d.appName], [REG_ID, APP_NAME]);
  assert.ok(d.types[0].progIdLabel.startsWith('DE '), 'labels in the language of Hostwerk');
  assert.deepStrictEqual(kindStates([{ ext: 'docx', kind: 'docs', registered: true, isDefault: true }, { ext: 'odt', kind: 'docs', registered: true, isDefault: false }, { ext: 'pdf', kind: 'pdf', registered: true, isDefault: true }]).docs.state, 'partial');
  assert.strictEqual(kindStates([{ ext: 'pdf', kind: 'pdf', registered: true, isDefault: true }]).pdf.state, 'default');
  assert.strictEqual(kindStates([{ ext: 'xlsx', kind: 'sheets', registered: true, isDefault: false }]).sheets.state, 'openwith');
  assert.strictEqual(kindStates([]).slides.state, 'none');
});

/* ---------- a simulated Windows registry (reg.exe query/delete and .reg import) ---------- */
function fakeRegistry() {
  const keys = new Map(); // lower-case key → { name, values: Map }
  const norm = (k) => k.toLowerCase();
  const ensure = (k) => { const parts = k.split('\\'); for (let i = 1; i <= parts.length; i++) { const sub = parts.slice(0, i).join('\\'); if (!keys.has(norm(sub))) keys.set(norm(sub), { name: sub, values: new Map() }); } return keys.get(norm(k)); };
  for (const k of ['HKEY_CURRENT_USER\\Software\\Classes\\Applications', 'HKEY_CURRENT_USER\\Software\\RegisteredApplications', 'HKEY_CURRENT_USER\\Software\\Classes\\.docx\\OpenWithProgids']) ensure(k).values.set('', '');
  ensure('HKEY_CURRENT_USER\\Software\\Classes\\.docx\\OpenWithProgids').values.set('Word.Document.12', '');
  const children = (k) => [...keys.keys()].filter((x) => x.startsWith(`${norm(k)}\\`) && !x.slice(norm(k).length + 1).includes('\\'));
  const delTree = (k) => { for (const x of [...keys.keys()]) if (x === norm(k) || x.startsWith(`${norm(k)}\\`)) keys.delete(x); };
  const un = (s) => s.replace(/\\(["\\])/g, '$1');
  const opened = [];
  return {
    keys, opened, platform: 'win32',
    snapshot: () => JSON.stringify([...keys.entries()].sort().map(([k, v]) => [k, [...v.values.entries()].sort()])),
    setUserChoice: (ext, progId) => ensure(`HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.${ext}\\UserChoice`).values.set('ProgId', progId),
    reg: async (args) => {
      const [op, k] = args;
      const e = keys.get(norm(k));
      if (op === 'query') {
        if (!e) return { code: 1, stdout: '' };
        let vals = [...e.values.entries()];
        if (args[2] === '/ve') vals = vals.filter(([n]) => n === '');
        else if (args[2] === '/v') { vals = vals.filter(([n]) => n === args[3]); if (!vals.length) return { code: 1, stdout: '' }; }
        const lines = ['', e.name, ...vals.map(([n, v]) => `    ${n === '' ? '(Standard)' : n}    REG_SZ    ${v}`), ''];
        if (args.length === 2) lines.push(...children(k).map((c) => keys.get(c).name));
        return { code: 0, stdout: lines.join('\r\n') };
      }
      if (op === 'delete') { if (!e) return { code: 1 }; delTree(k); return { code: 0 }; }
      return { code: 1 };
    },
    regImport: async (txt) => {
      let cur = null;
      for (const l of txt.split('\r\n')) {
        let m;
        if ((m = /^\[-(.+)\]$/.exec(l))) { delTree(m[1]); cur = null; } else if ((m = /^\[(.+)\]$/.exec(l))) cur = ensure(m[1]);
        else if ((m = /^(@|"((?:[^"\\]|\\.)*)")=-$/.exec(l))) cur.values.delete(m[1] === '@' ? '' : un(m[2]));
        else if ((m = /^(@|"((?:[^"\\]|\\.)*)")="((?:[^"\\]|\\.)*)"$/.exec(l))) cur.values.set(m[1] === '@' ? '' : un(m[2]), un(m[3]));
      }
      return { code: 0 };
    },
    notifyAssoc: async () => ({ code: 0 }),
    openUri: async (u) => { opened.push(u); return { code: 0 }; },
  };
}

/** host.fileTypes like Hostwerk's services/extapi.js builds it – over the core FileTypes with the simulated registry */
function hostWith({ packaged = true } = {}) {
  const dir = tmp();
  const io = fakeRegistry();
  let ft = null;
  if (coreDir) {
    const { FileTypes } = require(path.join(coreDir, 'server', 'core', 'filetypes.js'));
    const launch = { exe: 'C:\\Program Files\\Hostwerk\\Hostwerk.exe', packaged, appPath: 'C:\\src\\hostwerk' };
    ft = new FileTypes({ paths: { data: dir }, app: { launch } }, { io });
  }
  const before = io.snapshot();
  const fileTypes = ft ? {
    register: (def, o) => ft.register('office', def, { allowDev: !!(o && o.allowDev), setDefault: !!(o && o.setDefault) }),
    unregister: (id) => ft.unregister('office', id),
    state: (id) => ft.state('office', id),
    openDefaultsSettings: (id) => ft.openDefaultsSettings('office', id),
  } : (() => {
    // stand-in without a Hostwerk checkout: same answers, no registry
    let reg = null;
    return {
      register: async (def) => { if (!packaged) throw Object.assign(new Error('dev'), { code: 'EDEVBUILD' }); reg = def; return {}; },
      unregister: async () => { reg = null; return {}; },
      state: async () => (reg ? { registered: true, types: reg.types.map((x) => ({ ...x, registered: true, isDefault: false })) } : { registered: false, types: [], ok: true, dev: !packaged }),
      openDefaultsSettings: async () => { io.opened.push(`ms-settings:defaultapps?registeredAppUser=${encodeURIComponent(APP_NAME)}`); return { opened: true }; },
    };
  })();
  let settings = SET.normalize({}, dir);
  const host = { version: '1.0.0', t: (s) => s, supports: (n) => n === 'fileTypes', fileTypes };
  const reg = new OsReg({ host, file: path.join(dir, 'os.json'), settings: () => settings, log: () => {} });
  return { reg, io, before, host, setSettings: (p) => { settings = SET.merge(settings, p, dir).settings; } };
}

test('"Open with": registered once per version, notice once, removed without a trace (simulated registry)', async (t) => {
  if (!coreDir) t.diagnostic('Hostwerk core not found – running against a stand-in of host.fileTypes (set HOSTWERK_DIR for the real one)');
  const { reg, io, before, host, setSettings } = hostWith();
  assert.deepStrictEqual(await reg.sync(), { registered: true });
  let st = await reg.state();
  assert.ok(st.registered && st.enabled && st.notice, 'registered, notice pending');
  for (const k of ['docs', 'sheets', 'slides', 'pdf']) assert.strictEqual(st.kinds[k].state, 'openwith', k);
  if (coreDir) {
    assert.ok(io.keys.has('hkey_current_user\\software\\classes\\hostwerk.office.docx\\shell\\open\\command'), 'ProgID written');
    assert.ok(io.keys.get('hkey_current_user\\software\\classes\\.docx\\openwithprogids').values.has('Word.Document.12'), 'Word stays');
    // the user picked Hostwerk in Windows' settings
    for (const e of OS_EXTS.filter((x) => FORMATS[x].kind === 'pdf')) io.setUserChoice(e, `Hostwerk.office.${e}`);
    assert.strictEqual((await reg.state()).kinds.pdf.state, 'default');
  }
  reg.seen();
  assert.strictEqual((await reg.state()).notice, false);
  assert.deepStrictEqual(await reg.sync(), { current: true }, 'nothing again for the same version');
  // removed by the user (Hostwerk → Settings → Advanced): not forced back for this version
  await reg.unregister();
  assert.deepStrictEqual(await reg.sync(), { current: true });
  assert.strictEqual((await reg.state()).registered, false);
  // an update registers again (without a second notice)
  host.version = '1.0.1';
  assert.deepStrictEqual(await reg.sync(), { registered: true });
  assert.strictEqual((await reg.state()).notice, false);
  // Windows' "Default apps" page for Hostwerk Office
  const r = await reg.openDefaults();
  assert.ok(io.opened.some((u) => u === 'ms-settings:defaultapps?registeredAppUser=Hostwerk%20Office'), io.opened.join());
  assert.ok(r.state.registered);
  // switched off in the settings → removed
  setSettings({ integrations: { osOpenWith: false } });
  assert.deepStrictEqual(await reg.sync(), { off: true });
  assert.strictEqual((await reg.state()).registered, false);
  if (coreDir) {
    // only the UserChoice values the "user" set by hand stay – everything Hostwerk wrote is gone
    const left = [...io.keys.keys()].filter((k) => /hostwerk/.test(k) && !/userchoice/.test(k));
    assert.deepStrictEqual(left, []);
    for (const e of OS_EXTS.filter((x) => FORMATS[x].kind === 'pdf')) io.keys.delete(`hkey_current_user\\software\\microsoft\\windows\\currentversion\\explorer\\fileexts\\.${e}\\userchoice`);
    const now = JSON.parse(io.snapshot()).filter(([k]) => !k.startsWith('hkey_current_user\\software\\microsoft')); // the test's own UserChoice keys
    assert.deepStrictEqual(now, JSON.parse(before), 'registry exactly as before');
  }
});

test('"Open with": a developer start answers EDEVBUILD – silent at start, clear on the button', async () => {
  const { reg, io } = hostWith({ packaged: false });
  const r = await reg.sync();
  assert.strictEqual(r.error, 'EDEVBUILD');
  const st = await reg.state();
  assert.strictEqual(st.registered, false);
  assert.strictEqual(st.code, 'EDEVBUILD');
  await assert.rejects(() => reg.register(), (e) => e.code === 'EDEVBUILD' && e.status === 409);
  assert.ok(![...io.keys.keys()].some((k) => /hostwerk/.test(k)), 'nothing written');
});

test('"Open with": without the desktop app nothing happens and the state says why', async () => {
  const dir = tmp();
  const reg = new OsReg({ host: { version: '1', supports: () => false }, file: path.join(dir, 'os.json'), settings: () => SET.normalize({}, dir) });
  assert.deepStrictEqual(await reg.sync(), { skipped: true });
  const st = await reg.state();
  assert.strictEqual(st.supported, false);
  assert.ok(['ENODESKTOP', 'ENOTSUPPORTED'].includes(st.code));
  await assert.rejects(() => reg.register(), (e) => e.code === 'ENODESKTOP');
});

test('server modules load (routes, index with uninstall hook)', () => {
  assert.strictEqual(typeof require('../server/routes').routes, 'function');
  const idx = require('../server/index');
  for (const k of ['activate', 'deactivate', 'uninstall']) assert.strictEqual(typeof idx[k], 'function', k);
});
