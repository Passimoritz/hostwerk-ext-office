'use strict';
/*
 * Hostwerk under the system's "Open with" for office files (SDK 3.15, host.fileTypes – the core writes the registry /
 * .desktop file, never this module). Windows: per user, no admin; the default stays the user's choice – Windows 10/11
 * only lets the user pick it (Settings → Default apps), so "Make default" opens that page. Linux: .desktop file, the
 * default only on request (xdg-mime, with consent). macOS: not possible at runtime.
 *
 * Registered once per install/upgrade (marker = version + list of types in <data>/office-data/os.json); a registration
 * the user removed himself (Hostwerk → Settings → Advanced → Open files) is not forced back until the next update.
 * Setting integrations.osOpenWith (default on) turns it off and removes it. Developer starts answer EDEVBUILD – kept silent.
 */
const crypto = require('crypto');
const { FORMATS } = require('./formats');
const { readJson, writeJson } = require('./store');

const REG_ID = 'office';
const APP_NAME = 'Hostwerk Office';
/** Exactly the formats Office edits or shows well (each with mime + kind from formats.js) */
const OS_EXTS = [
  'docx', 'docm', 'dotx', 'doc', 'odt', 'ott', 'rtf', 'md',
  'xlsx', 'xlsm', 'xltx', 'xls', 'ods', 'ots', 'csv',
  'pptx', 'pptm', 'ppsx', 'potx', 'ppt', 'odp', 'otp',
  'pdf', 'epub', 'vsdx', 'pages', 'numbers', 'key', 'eml', 'msg', 'vcf', 'ics',
];
const KINDS = ['docs', 'sheets', 'slides', 'pdf', 'viewer'];
/** Names without characters the registry text refuses (quotes, %, \ and /) */
const clean = (s, max) => String(s || '').replace(/["%\\/]+/g, ' ').replace(/[\u0000-\u001f\u007f]/g, '').replace(/\s+/g, ' ').trim().slice(0, max);

/** The types for host.fileTypes.register: [{ ext, mime, kind, progIdLabel }] – t = translation of the label */
function osTypes(t = (x) => x) {
  return OS_EXTS.map((ext) => {
    const f = FORMATS[ext];
    if (!f) throw new Error(`osreg: unknown format .${ext}`);
    return { ext, mime: f.mime, kind: f.kind, progIdLabel: clean(`${t(f.label)} (Hostwerk)`, 80) };
  });
}
const typesHash = () => crypto.createHash('sha1').update(OS_EXTS.join(',')).digest('hex').slice(0, 10);
const definition = (t = (x) => x) => ({
  id: REG_ID, appName: APP_NAME,
  description: clean(t('Documents, spreadsheets, presentations and PDFs in Hostwerk'), 200),
  types: osTypes(t),
});

/**
 * Per kind: 'default' (Hostwerk opens all of them by double-click), 'partial' (some), 'openwith' (only under "Open
 * with"), 'none' (not registered). types = state(id).types of the core.
 */
function kindStates(types) {
  const out = {};
  for (const k of KINDS) {
    const list = (types || []).filter((x) => (x.kind || (FORMATS[x.ext] || {}).kind) === k);
    if (!list.length) { out[k] = { state: 'none', total: 0, defaults: 0 }; continue; }
    const reg = list.filter((x) => x.registered);
    const def = list.filter((x) => x.isDefault);
    out[k] = { state: def.length && def.length === list.length ? 'default' : def.length ? 'partial' : reg.length ? 'openwith' : 'none', total: list.length, defaults: def.length, exts: list.map((x) => x.ext) };
  }
  return out;
}

class OsReg {
  /** deps: { host, file (os.json), settings(), log } */
  constructor({ host, file, settings, log = () => {} }) { this.host = host; this.file = file; this.settings = settings; this.log = log; this.busy = Promise.resolve(); }
  get ft() { return this.host.fileTypes; }
  supported() { try { return typeof this.host.supports === 'function' && this.host.supports('fileTypes') && !!this.host.fileTypes; } catch { return false; } }
  enabled() { const s = this.settings(); return !(s && s.integrations && s.integrations.osOpenWith === false); }
  marker() { return readJson(this.file, {}); }
  saveMarker(patch) { const m = { ...this.marker(), ...patch }; writeJson(this.file, m); return m; }
  want() { return `${this.host.version || '0'}|${typesHash()}`; }
  t(s) { try { return this.host.t ? this.host.t(s) : s; } catch { return s; } }
  serial(fn) { const p = this.busy.then(fn, fn); this.busy = p.catch(() => {}); return p; }

  /** State for the shell: { supported, platform, code, dev, enabled, registered, stale, kinds, types, notice, lastError } */
  async state() {
    const m = this.marker();
    const base = { supported: this.supported(), platform: process.platform, enabled: this.enabled(), notice: !!m.notice, lastError: m.lastError || null, appName: APP_NAME };
    if (!base.supported) {
      // older/other Hostwerk, server without desktop, macOS: say why
      const code = process.platform === 'darwin' ? 'ENOTSUPPORTED' : 'ENODESKTOP';
      return { ...base, code, registered: false, kinds: kindStates([]), types: [] };
    }
    let st = null;
    try { st = await this.ft.state(REG_ID); } catch (e) { st = { registered: false, types: [], code: e.code || null, reason: e.message }; }
    const code = st.code || (st.ok === false ? 'ENOTSUPPORTED' : null) || (m.lastError === 'EDEVBUILD' && !st.registered ? 'EDEVBUILD' : null);
    return { ...base, code, dev: !!st.dev, registered: !!st.registered, stale: !!st.stale, kinds: kindStates(st.types), types: (st.types || []).map((x) => ({ ext: x.ext, kind: x.kind, registered: !!x.registered, isDefault: !!x.isDefault })) };
  }

  /** Register now (button or start). opts: { setDefault (Linux), notice (undefined = only the very first time) } → state() */
  register({ setDefault = false, notice } = {}) {
    return this.serial(async () => {
      if (!this.supported()) throw Object.assign(new Error(this.t('“Open with” needs the Hostwerk desktop app on Windows or Linux.')), { status: 409, code: 'ENODESKTOP' });
      try {
        await this.ft.register(definition((s) => this.t(s)), { setDefault: !!setDefault && process.platform === 'linux' });
        const first = !this.marker().version;
        this.saveMarker({ version: this.want(), at: Date.now(), lastError: null, removed: false, ...((notice === undefined ? first : notice) ? { notice: true } : {}) });
      } catch (e) {
        this.saveMarker({ lastError: e.code || 'ERROR' });
        if (e.code === 'EDEVBUILD') throw Object.assign(new Error(this.t('Hostwerk runs from source (developer start) – “Open with” is only set up by the installed app.')), { status: 409, code: 'EDEVBUILD' });
        throw e;
      }
      return this.state();
    });
  }
  /** Remove (button, setting off, uninstall) → state() */
  unregister({ byUser = true } = {}) {
    return this.serial(async () => {
      if (this.supported()) { try { await this.ft.unregister(REG_ID); } catch (e) { if (!/unknown|missing/i.test(e.message || '')) throw e; } }
      this.saveMarker({ removed: !!byUser, notice: false });
      return this.state();
    });
  }
  /** Windows: Settings → Default apps (the user picks Hostwerk there – Windows' rule). Linux: xdg-mime with consent. */
  async openDefaults() {
    if (!this.supported()) throw Object.assign(new Error(this.t('“Open with” needs the Hostwerk desktop app on Windows or Linux.')), { status: 409, code: 'ENODESKTOP' });
    if (process.platform === 'linux') { await this.register({ setDefault: true }); return { opened: false, linux: true, state: await this.state() }; }
    let st = await this.state();
    if (!st.registered) st = await this.register();
    const r = await this.ft.openDefaultsSettings(REG_ID);
    return { ...r, state: st };
  }
  seen() { this.saveMarker({ notice: false }); return { ok: true }; }

  /**
   * At start and after the setting changed: register once per install/upgrade, remove when switched off. Never throws
   * (developer start = EDEVBUILD, kept silent and shown as state).
   */
  async sync() {
    try {
      if (!this.supported()) return { skipped: true };
      const m = this.marker();
      if (!this.enabled()) {
        const st = await this.ft.state(REG_ID).catch(() => null);
        if (st && st.registered) await this.unregister({ byUser: false });
        return { off: true };
      }
      if (m.version === this.want()) return { current: true }; // done for this version – also when the user removed it himself
      await this.register();
      return { registered: true };
    } catch (e) {
      if (e.code !== 'EDEVBUILD') this.log('open with', e.message);
      return { error: e.code || e.message };
    }
  }
}

module.exports = { OsReg, osTypes, definition, kindStates, typesHash, OS_EXTS, REG_ID, APP_NAME };
