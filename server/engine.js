'use strict';
/*
 * Private Office engine: an own LibreOffice in <data>/office-engine – no system install, no admin rights.
 *   Windows  official MSI from download.documentfoundation.org (SHA-256 from the mirror list), unpacked with
 *            msiexec /a <msi> /qn TARGETDIR=<dir> (administrative install = portable copy)
 *   Linux    official deb archive (…_Linux_x86-64_deb.tar.gz) unpacked with tar + dpkg-deb -x (or ar + tar)
 *   macOS    official dmg: hdiutil attach → copy LibreOffice.app → detach
 * The current stable version comes from the "stable" folder listing (newest version), PINNED is the fallback.
 * Updates: staged next to the engine (office-engine.new), probed with a real conversion, then swapped – also through
 * Hostwerk's security updates (host.ctx.patchUpdates).
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile, spawn } = require('child_process');

const BASE = 'https://download.documentfoundation.org/libreoffice/stable/';
const PINNED = '25.8.7';
const SIZE_HINT = 380 * 1048576;

/** Versions from the directory listing, newest first */
function parseVersions(html) {
  const set = new Set();
  for (const m of String(html || '').matchAll(/href="(\d+\.\d+\.\d+(?:\.\d+)?)\/"/g)) set.add(m[1]);
  return [...set].sort(cmpVersion).reverse();
}
function cmpVersion(a, b) {
  const x = String(a || '0').split('.').map(Number); const y = String(b || '0').split('.').map(Number);
  for (let i = 0; i < Math.max(x.length, y.length); i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d < 0 ? -1 : 1; }
  return 0;
}
/** Download file for a version on this platform → { url, name, kind: 'msi' | 'deb' | 'dmg' } or null */
function asset(version, platform = process.platform, arch = process.arch) {
  const v = version;
  if (platform === 'win32') { const a = arch === 'arm64' ? ['aarch64', 'aarch64'] : ['x86_64', 'x86-64']; const name = `LibreOffice_${v}_Win_${a[1]}.msi`; return { url: `${BASE}${v}/win/${a[0]}/${name}`, name, kind: 'msi' }; }
  if (platform === 'darwin') { const a = arch === 'arm64' ? ['aarch64', 'aarch64'] : ['x86_64', 'x86-64']; const name = `LibreOffice_${v}_MacOS_${a[1]}.dmg`; return { url: `${BASE}${v}/mac/${a[0]}/${name}`, name, kind: 'dmg' }; }
  if (platform === 'linux') { const a = arch === 'arm64' ? ['aarch64', 'aarch64'] : ['x86_64', 'x86-64']; const name = `LibreOffice_${v}_Linux_${a[1]}_deb.tar.gz`; return { url: `${BASE}${v}/deb/${a[0]}/${name}`, name, kind: 'deb' }; }
  return null;
}
/** SHA-256 from the mirrorbrain page (<file>.mirrorlist) */
function parseSha256(html) {
  const s = String(html || '');
  const i = s.search(/SHA-256/i);
  const m = (i >= 0 ? s.slice(i) : s).match(/\b([0-9a-f]{64})\b/i);
  return m ? m[1].toLowerCase() : null;
}
/**
 * msiexec arguments for a portable (administrative) install. msiexec parses its own command line, so the values are
 * quoted here and passed verbatim (paths are ours and must not contain quotes).
 */
function msiArgs(msi, target, log) {
  for (const p of [msi, target, log]) if (/["\r\n]/.test(p)) throw new Error('Invalid path for msiexec');
  return ['/a', `"${msi}"`, '/qn', `TARGETDIR="${target}"`, '/l*v', `"${log}"`];
}

const run = (cmd, args, opts = {}) => new Promise((resolve) => {
  const p = spawn(cmd, args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'], ...opts });
  let out = '';
  p.stdout.on('data', (c) => { out += c; }); p.stderr.on('data', (c) => { out += c; });
  const t = opts.timeout ? setTimeout(() => { try { p.kill(); } catch { /* weg */ } }, opts.timeout) : null;
  p.on('error', (e) => { if (t) clearTimeout(t); resolve({ code: -1, out: e.message }); });
  p.on('close', (code) => { if (t) clearTimeout(t); resolve({ code, out }); });
});
const which = (cmd) => new Promise((r) => execFile(process.platform === 'win32' ? 'where' : 'which', [cmd], { windowsHide: true }, (e, out) => r(e ? null : String(out).split(/\r?\n/)[0].trim() || null)));
function dirSize(p) { let n = 0; try { const st = fs.statSync(p); if (!st.isDirectory()) return st.size; for (const e of fs.readdirSync(p)) n += dirSize(path.join(p, e)); } catch { /* weg */ } return n; }

class Engine {
  /**
   * deps: { dir, fetchText(url), downloadFile(url, dest, { sha256, onProgress }), runJob(title, fn, opts), findSoffice(dir),
   *         probe(sofficePath) → Promise, t(text, vars), log }
   */
  constructor(deps) {
    this.d = deps;
    this.dir = deps.dir;
    this.job = null;
    this.state = { installing: false, progress: 0, text: '', error: null };
    this.latestCache = null;
  }
  info() { try { return JSON.parse(fs.readFileSync(path.join(this.dir, 'engine.json'), 'utf8')); } catch { return null; } }
  soffice(dir = this.dir) { return this.d.findSoffice(dir); }
  status() {
    const i = this.info(); const exe = i ? this.soffice() : null;
    return { installed: !!(i && exe), version: i ? i.version : null, path: exe, installedAt: i ? i.installedAt : null, size: i ? i.size || null : null, sizeHint: SIZE_HINT,
      installing: this.state.installing, progress: this.state.progress, text: this.state.text, error: this.state.error, jobId: this.job ? this.job.id : null,
      latest: this.latestCache ? this.latestCache.version : null, supported: !!asset(PINNED), platform: process.platform };
  }
  /** Newest stable version + download + checksum */
  async latest({ fresh = false } = {}) {
    if (!fresh && this.latestCache && Date.now() - this.latestCache.t < 6 * 3600000) return this.latestCache;
    let version = PINNED;
    try { const list = parseVersions(await this.d.fetchText(BASE)); if (list.length) version = list[0]; } catch (e) { this.d.log && this.d.log('engine: version list failed –', e.message); }
    const a = asset(version);
    if (!a) throw Object.assign(new Error('The Office engine is not available for this system.'), { status: 501 });
    let sha256 = null;
    try { sha256 = parseSha256(await this.d.fetchText(`${a.url}.mirrorlist`)); } catch { /* below */ }
    if (!sha256) throw Object.assign(new Error('The checksum of the Office engine could not be loaded – please try again later.'), { status: 502 });
    this.latestCache = { version, ...a, sha256, t: Date.now() };
    return this.latestCache;
  }
  /** Install or update as a Hostwerk job → job */
  install({ update = false } = {}) {
    if (this.job && this.state.installing) return this.job;
    const T = this.d.t;
    const title = update ? T('Update the Office engine') : T('Set up the Office engine');
    this.state = { installing: true, progress: 0, text: T('Starting …'), error: null };
    const job = this.d.runJob(title, async (j) => {
      const step = (pct, text) => { this.state.progress = pct; this.state.text = text; j.progress(pct, text); };
      try {
        step(2, T('Looking for the current version …'));
        const l = await this.latest({ fresh: true });
        const cur = this.info();
        if (cur && cmpVersion(cur.version, l.version) >= 0 && this.soffice()) { step(100, T('The Office engine is up to date.')); return { version: cur.version, unchanged: true }; }
        const staged = await this.stage(l, j, step);
        step(92, T('Checking the engine …'));
        await this.d.probe(this.soffice(staged));
        this.activate(staged, l.version);
        step(100, T('The Office engine is ready.'));
        return { version: l.version };
      } catch (e) {
        this.state.error = e.message;
        throw e;
      } finally { this.state.installing = false; this.job = null; }
    }, { key: 'office-engine' });
    this.job = job;
    return job;
  }
  /** Download + unpack into <dir>.new → path of the staged folder */
  async stage(l, job, step = () => {}) {
    const T = this.d.t;
    const staged = `${this.dir}.new`;
    const dl = path.join(path.dirname(this.dir), 'office-engine-download');
    fs.rmSync(staged, { recursive: true, force: true });
    fs.mkdirSync(staged, { recursive: true }); fs.mkdirSync(dl, { recursive: true });
    const file = path.join(dl, l.name);
    step(5, T('Downloading the Office engine ({version}) …', { version: l.version }));
    await this.d.downloadFile(l.url, file, { sha256: l.sha256, onProgress: (rec, total) => step(5 + Math.round(70 * (total ? rec / total : 0)), T('Downloading the Office engine – {mb} MB', { mb: Math.round(rec / 1048576) })) });
    step(78, T('Unpacking …'));
    if (l.kind === 'msi') {
      const log = path.join(dl, 'msiexec.log');
      const r = await run('msiexec.exe', msiArgs(file, staged, log), { windowsVerbatimArguments: true, timeout: 30 * 60000 });
      if (r.code !== 0) throw new Error(T('Unpacking the Office engine failed (msiexec code {code}).', { code: r.code }));
      for (const f of fs.readdirSync(staged)) if (/\.msi$/i.test(f)) fs.rmSync(path.join(staged, f), { force: true }); // admin image keeps a copy of the msi
    } else if (l.kind === 'deb') {
      const tmp = path.join(dl, 'debs'); fs.rmSync(tmp, { recursive: true, force: true }); fs.mkdirSync(tmp);
      let r = await run('tar', ['-xzf', file, '-C', tmp], { timeout: 20 * 60000 });
      if (r.code !== 0) throw new Error(T('Unpacking the Office engine failed ({step}).', { step: 'tar' }));
      const debs = [];
      const walk = (d) => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else if (/\.deb$/.test(e.name)) debs.push(p); } };
      walk(tmp);
      const dpkg = await which('dpkg-deb');
      for (const deb of debs) {
        if (dpkg) r = await run(dpkg, ['-x', deb, staged], { timeout: 10 * 60000 });
        else {
          const x = path.join(tmp, 'ar'); fs.rmSync(x, { recursive: true, force: true }); fs.mkdirSync(x);
          r = await run('ar', ['x', deb], { cwd: x, timeout: 10 * 60000 });
          const data = fs.readdirSync(x).find((f) => /^data\.tar/.test(f));
          if (r.code === 0 && data) r = await run('tar', ['-xf', path.join(x, data), '-C', staged], { timeout: 10 * 60000 });
        }
        if (r.code !== 0) throw new Error(T('Unpacking the Office engine failed ({step}).', { step: path.basename(deb) }));
      }
      fs.rmSync(tmp, { recursive: true, force: true });
    } else if (l.kind === 'dmg') {
      const mnt = path.join(dl, 'mnt'); fs.mkdirSync(mnt, { recursive: true });
      let r = await run('hdiutil', ['attach', '-nobrowse', '-readonly', '-mountpoint', mnt, file], { timeout: 10 * 60000 });
      if (r.code !== 0) throw new Error(T('Unpacking the Office engine failed ({step}).', { step: 'hdiutil' }));
      try { r = await run('cp', ['-R', path.join(mnt, 'LibreOffice.app'), staged], { timeout: 20 * 60000 }); } finally { await run('hdiutil', ['detach', mnt, '-force'], { timeout: 120000 }); }
      if (r.code !== 0) throw new Error(T('Unpacking the Office engine failed ({step}).', { step: 'cp' }));
    }
    fs.rmSync(dl, { recursive: true, force: true });
    if (!this.soffice(staged)) throw new Error(T('The Office engine was unpacked but LibreOffice was not found in it.'));
    return staged;
  }
  /** Swap the staged folder in */
  activate(staged, version) {
    const old = `${this.dir}.old`;
    fs.rmSync(old, { recursive: true, force: true });
    if (fs.existsSync(this.dir)) fs.renameSync(this.dir, old);
    fs.renameSync(staged, this.dir);
    fs.writeFileSync(path.join(this.dir, 'engine.json'), JSON.stringify({ version, installedAt: Date.now(), size: dirSize(this.dir) }));
    fs.rm(old, { recursive: true, force: true }, () => {});
  }
  remove() {
    if (this.state.installing) throw Object.assign(new Error('The Office engine is being set up – please wait until it is done.'), { status: 409 });
    fs.rmSync(this.dir, { recursive: true, force: true });
    fs.rmSync(`${this.dir}.new`, { recursive: true, force: true });
    fs.rmSync(`${this.dir}.old`, { recursive: true, force: true });
  }
  /** Source for Hostwerk's security updates (host.ctx.patchUpdates.register) */
  patchSource() {
    return {
      label: 'Office engine (LibreOffice)', labels: { de: 'Office-Engine (LibreOffice)', en: 'Office engine (LibreOffice)' },
      items: () => { const i = this.info(); return i ? [{ key: 'office-engine', label: 'Office engine (LibreOffice)', version: i.version }] : []; },
      updates: async () => {
        const i = this.info(); if (!i) return [];
        const l = await this.latest().catch(() => null);
        // only the same release line (x.y) – bigger jumps are offered in the Office settings
        if (!l || cmpVersion(l.version, i.version) <= 0 || l.version.split('.').slice(0, 2).join('.') !== String(i.version).split('.').slice(0, 2).join('.')) return [];
        return [{ key: 'office-engine', label: 'Office engine (LibreOffice)', current: i.version, latest: l.version, target: l }];
      },
      plan: (u) => {
        let staged = null;
        return {
          install: async () => { staged = await this.stage(u.target, null); },
          probe: async () => { await this.d.probe(this.soffice(staged)); },
          activate: async () => { this.activate(staged, u.target.version); },
          discard: async () => { if (staged) fs.rmSync(staged, { recursive: true, force: true }); },
        };
      },
    };
  }
}

module.exports = { Engine, parseVersions, parseSha256, asset, msiArgs, cmpVersion, BASE, PINNED };
