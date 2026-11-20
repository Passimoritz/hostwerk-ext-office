'use strict';
/*
 * Windows: which app opens a file type by double-click, and the two ways Windows lets the USER make Hostwerk Office the
 * default (Windows 10/11 never let an app set it itself – UserChoice carries a hash only Windows writes):
 *
 *   1. Windows' own "How do you want to open this file?" dialog for one file type – OpenWith.exe <sample file>
 *      (rundll32 shell32.dll,OpenAs_RunDLL on systems without it). The user picks "Hostwerk Office" → "Always".
 *      Works on every Windows 10/11 build, also when Hostwerk Office is missing from Settings → Default apps.
 *   2. Settings → Default apps – opened through ShellExecute (rundll32 url.dll,FileProtocolHandler), because
 *      explorer.exe drops a ms-settings: link that carries a query (…?registeredAppUser=…): Settings then does not move
 *      at all. Windows only jumps to the app's page when it lists the app; otherwise it shows the overview, where the
 *      user can search a file type (".docx") – the card says so.
 *
 * State: HKCU\…\Explorer\FileExts\.<ext> – Windows 11 24H2+ keeps the choice in UserChoiceLatest\ProgId (value ProgId,
 * a subkey!), older builds in UserChoiceLatest (value ProgId) or UserChoice (value ProgId). Read only, never written.
 * One "reg query … /s" over all of FileExts (≈ 0.3 s) instead of one process per type; cached for a moment.
 */
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');

const FILE_EXTS = 'HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts';
/** Main type per kind: the sample file the dialog is opened for */
const KIND_EXT = { docs: 'docx', sheets: 'xlsx', slides: 'pptx', pdf: 'pdf' };
const SAFE_EXT = /^[a-z0-9]{1,10}$/;

/**
 * Output of `reg query <FileExts> /s` → Map(ext → ProgId the user chose, or null).
 * Order: UserChoiceLatest\ProgId → UserChoiceLatest → UserChoice (the newest place Windows wrote wins).
 */
function parseFileExts(stdout) {
  const raw = new Map(); // ext → { sub, latest, legacy }
  let cur = null;
  for (const line of String(stdout || '').split(/\r?\n/)) {
    const t = line.trim();
    if (/^HKEY_/i.test(t)) {
      const m = /\\FileExts\\\.([^\\]+)\\(UserChoiceLatest\\ProgId|UserChoiceLatest|UserChoice)$/i.exec(t);
      cur = m ? { ext: m[1].toLowerCase(), where: /\\ProgId$/i.test(m[2]) ? 'sub' : /Latest/i.test(m[2]) ? 'latest' : 'legacy' } : null;
      continue;
    }
    if (!cur) continue;
    const v = /^ {4}(ProgId) {4}REG_[A-Z_]+ {4}(.*)$/i.exec(line);
    if (!v || !v[2].trim()) continue;
    const rec = raw.get(cur.ext) || {};
    rec[cur.where] = v[2].trim();
    raw.set(cur.ext, rec);
  }
  const out = new Map();
  for (const [ext, r] of raw) out.set(ext, r.sub || r.latest || r.legacy || null);
  return out;
}

/** isDefault per type: the core's answer OR the ProgId Windows really has (the core only reads UserChoiceLatest's value) */
function mergeChoices(types, choices) {
  if (!choices) return types;
  return (types || []).map((x) => {
    const pid = x.progId || `Hostwerk.office.${x.ext}`;
    const chosen = choices.get(String(x.ext).toLowerCase()) || null;
    return { ...x, userChoice: chosen, isDefault: !!x.isDefault || chosen === pid };
  });
}

/** Program and arguments for Windows' "Open with" dialog of one file (pure – tested) */
function openWithCommand(file, { systemRoot = 'C:\\Windows', hasOpenWith = true } = {}) {
  const f = String(file || '');
  if (!path.win32.isAbsolute(f) || /["\r\n\0]/.test(f)) throw Object.assign(new Error('openWith: unusable path'), { status: 400 });
  const sys = path.win32.join(systemRoot, 'System32');
  if (hasOpenWith) return { cmd: path.win32.join(sys, 'OpenWith.exe'), args: [f], verbatim: false };
  // OpenAs_RunDLL takes the rest of the command line as the path – unquoted, so no argument quoting by Node
  return { cmd: path.win32.join(sys, 'rundll32.exe'), args: [`shell32.dll,OpenAs_RunDLL ${f}`], verbatim: true };
}
/** ShellExecute of a ms-settings: link (pure – tested). The query keeps its space: Settings does not decode %20. */
function settingsCommand(appName, { systemRoot = 'C:\\Windows' } = {}) {
  const name = String(appName || '').replace(/[^\w .()-]/g, '').trim();
  const uri = name ? `ms-settings:defaultapps?registeredAppUser=${name}` : 'ms-settings:defaultapps';
  return { cmd: path.win32.join(systemRoot, 'System32', 'rundll32.exe'), args: ['url.dll,FileProtocolHandler', uri], uri };
}

class OsDefault {
  /** deps: { dir (sample files), run(cmd, args, opts) → { code, stdout }, spawn(cmd, args, opts), platform, systemRoot } */
  constructor({ dir, run = null, spawn = null, platform = process.platform, systemRoot = process.env.SystemRoot || 'C:\\Windows' }) {
    this.dir = dir; this.platform = platform; this.systemRoot = systemRoot;
    this.run = run || ((cmd, args) => new Promise((resolve) => execFile(cmd, args, { windowsHide: true, timeout: 15000, maxBuffer: 32 * 1048576, encoding: 'latin1' }, (e, so) => resolve({ code: e ? 1 : 0, stdout: String(so || '') }))));
    this.spawn = spawn || ((cmd, args, opts) => { const cp = require('child_process').spawn(cmd, args, { detached: true, stdio: 'ignore', windowsHide: false, ...opts }); cp.on('error', () => {}); cp.unref(); return cp; });
    this.cache = null;
  }
  get win() { return this.platform === 'win32'; }
  /** Map(ext → ProgId) or null (not Windows / reg failed) – cached 1.5 s (the card polls) */
  async choices({ fresh = false } = {}) {
    if (!this.win) return null;
    if (!fresh && this.cache && Date.now() - this.cache.at < 1500) return this.cache.map;
    const r = await this.run(path.win32.join(this.systemRoot, 'System32', 'reg.exe'), ['query', FILE_EXTS, '/s']).catch(() => ({ code: 1 }));
    const map = r.code === 0 ? parseFileExts(r.stdout) : null;
    this.cache = { at: Date.now(), map };
    return map;
  }
  /** Writes the sample file of `ext` (bytes from the editor chunk, so "Just once" opens a real empty document) → path */
  sample(ext, bytes, name = 'Hostwerk Office') {
    const e = String(ext || '').toLowerCase();
    if (!SAFE_EXT.test(e)) throw Object.assign(new Error('openWith: unknown type'), { status: 400 });
    fs.mkdirSync(this.dir, { recursive: true });
    const base = String(name).replace(/[^\w .()-]/g, '').trim() || 'Sample';
    const file = path.join(this.dir, `${base}.${e}`);
    if (bytes && bytes.length) fs.writeFileSync(file, bytes);
    else if (!fs.existsSync(file)) fs.writeFileSync(file, Buffer.alloc(0));
    return file;
  }
  /** Windows' "How do you want to open this file?" for one type → { opened, ext, file, cmd } */
  openWith(ext, bytes, name) {
    if (!this.win) throw Object.assign(new Error('Only on Windows.'), { status: 409, code: 'ENOTWINDOWS' });
    const file = this.sample(ext, bytes, name);
    const hasOpenWith = fs.existsSync(path.win32.join(this.systemRoot, 'System32', 'OpenWith.exe'));
    const c = openWithCommand(file, { systemRoot: this.systemRoot, hasOpenWith });
    this.spawn(c.cmd, c.args, { windowsVerbatimArguments: c.verbatim });
    this.cache = null;
    return { opened: true, ext: String(ext).toLowerCase(), file, cmd: path.basename(c.cmd) };
  }
  /** Settings → Default apps (ShellExecute) → { opened, uri } */
  openSettings(appName) {
    if (!this.win) throw Object.assign(new Error('Only on Windows.'), { status: 409, code: 'ENOTWINDOWS' });
    const c = settingsCommand(appName, { systemRoot: this.systemRoot });
    this.spawn(c.cmd, c.args, {});
    this.cache = null;
    return { opened: true, uri: c.uri };
  }
}

module.exports = { OsDefault, parseFileExts, mergeChoices, openWithCommand, settingsCommand, KIND_EXT, FILE_EXTS };
