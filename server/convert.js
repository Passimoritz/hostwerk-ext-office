'use strict';
/*
 * Converters on this PC:
 *   LibreOffice  soffice --headless --convert-to <fmt> --outdir <dir> <file> with a private profile
 *                (-env:UserInstallation=file:///…) so it also runs while the user has LibreOffice open.
 *                Found in the settings path, the private Office engine (<data>/office-engine), PATH or the usual places.
 *   MS Office    (Windows) Word / Excel / PowerPoint / Visio through COM in a PowerShell script. The script is a fixed
 *                file; paths and format numbers come in a JSON job file (never interpolated into a command line).
 *                Macros are disabled (AutomationSecurity = force disable), alerts off, the app is quit and released in
 *                finally – unless it was already running (PowerPoint shares one instance: then only our file is closed).
 * Only one conversion per app at a time (queue), every run has a timeout and its own temp folder (always removed).
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, execFile } = require('child_process');
const { pathToFileURL } = require('url');

const WIN = process.platform === 'win32';
const MAC = process.platform === 'darwin';

/* ---------------- Format tables (MS Office) ---------------- */
// Word: WdSaveFormat
const WORD_SAVE = { docx: 16, docm: 13, dotx: 14, dotm: 15, doc: 0, dot: 1, rtf: 6, txt: 2, html: 10, htm: 10, odt: 23, pdf: 17, xps: 18, mht: 9, mhtml: 9, xml: 19 };
const WORD_OPEN = ['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'txt', 'html', 'htm', 'xhtml', 'mht', 'mhtml', 'odt', 'pdf', 'xml', 'wpd', 'wps'];
// Excel: XlFileFormat (pdf/xps via ExportAsFixedFormat)
const EXCEL_SAVE = { xlsx: 51, xlsm: 52, xlsb: 50, xls: 56, xltx: 54, xltm: 53, xlt: 17, csv: 62, tsv: -4158, txt: -4158, ods: 60, html: 44, htm: 44, dif: 9, slk: 2, prn: 36, xml: 46, pdf: 'pdf', xps: 'xps' };
const EXCEL_OPEN = ['xlsx', 'xlsm', 'xlsb', 'xls', 'xltx', 'xltm', 'xlt', 'csv', 'tsv', 'tab', 'txt', 'ods', 'html', 'htm', 'dif', 'slk', 'prn', 'xml', 'dbf', 'wk1', 'wks', 'qpw'];
// PowerPoint: PpSaveAsFileType
const PPT_SAVE = { pptx: 24, pptm: 25, ppt: 1, ppsx: 28, ppsm: 29, pps: 7, potx: 26, potm: 27, pot: 5, odp: 35, pdf: 32, xps: 33 };
const PPT_OPEN = ['pptx', 'pptm', 'ppt', 'ppsx', 'ppsm', 'pps', 'potx', 'potm', 'pot', 'odp'];
// Visio: export to PDF/XPS, SaveAs for its own formats
const VISIO_SAVE = { pdf: 'pdf', xps: 'xps', vsdx: 'save', vsd: 'save', svg: 'save', png: 'save' };
const VISIO_OPEN = ['vsd', 'vsdx', 'vsdm', 'vstx', 'vss', 'vssx', 'vst'];
const MS = {
  word: { open: WORD_OPEN, save: WORD_SAVE, exe: 'WINWORD.EXE' },
  excel: { open: EXCEL_OPEN, save: EXCEL_SAVE, exe: 'EXCEL.EXE' },
  powerpoint: { open: PPT_OPEN, save: PPT_SAVE, exe: 'POWERPNT.EXE' },
  visio: { open: VISIO_OPEN, save: VISIO_SAVE, exe: 'VISIO.EXE' },
};

/* ---------------- LibreOffice ---------------- */
const LO_IN = ['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'odt', 'ott', 'fodt', 'sxw', 'rtf', 'txt', 'html', 'htm', 'xhtml', 'wpd', 'wps', 'wri', 'pages', 'abw', 'lwp', 'hwp', 'md',
  'xlsx', 'xlsm', 'xlsb', 'xltx', 'xltm', 'xls', 'xlt', 'ods', 'ots', 'fods', 'sxc', 'csv', 'tsv', 'tab', 'dif', 'slk', 'dbf', 'numbers', 'wk1', 'wks', 'qpw', 'et', 'prn',
  'pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'ppt', 'pps', 'pot', 'odp', 'otp', 'fodp', 'sxi', 'key', 'dps',
  'odg', 'otg', 'fodg', 'vsd', 'vsdx', 'vsdm', 'vstx', 'pub', 'svg', 'pdf', 'xps', 'oxps', 'odf', 'epub', 'cdr'];
const LO_OUT = ['pdf', 'docx', 'doc', 'odt', 'rtf', 'txt', 'html', 'htm', 'epub', 'xlsx', 'xls', 'ods', 'csv', 'fods', 'fodt', 'pptx', 'ppt', 'odp', 'fodp', 'odg', 'png', 'svg', 'jpg', 'xhtml', 'dotx', 'xltx', 'potx', 'docm', 'xlsm', 'pptm', 'ppsx', 'pps', 'xlsb', 'tsv', 'dif', 'slk', 'wps'];
const KIND_OF = (ext) => { const f = require('./formats').FORMATS[ext]; return f ? f.kind : null; };
/** --convert-to argument: special filters where the plain extension picks a bad default */
function loFilter(from, to) {
  const kind = KIND_OF(from);
  if (to === 'csv') return 'csv:Text - txt - csv (StarCalc):44,34,76,1';
  if (to === 'tsv') return 'tsv:Text - txt - csv (StarCalc):9,34,76,1';
  if (to === 'txt') return kind === 'sheets' ? 'txt:Text - txt - csv (StarCalc):9,34,76,1' : 'txt:Text (encoded):UTF8';
  if (to === 'html' || to === 'htm') return kind === 'sheets' ? 'html:HTML (StarCalc)' : kind === 'slides' ? 'html:impress_html_Export' : 'html:XHTML Writer File:UTF8';
  if (to === 'pdf' && from === 'pdf') return 'pdf';
  return to;
}
/** LibreOffice reads PDFs into Draw – for "PDF → document" use the Writer import */
const loInfilter = (from, to) => (from === 'pdf' && ['docx', 'doc', 'odt', 'rtf', 'txt'].includes(to) ? 'writer_pdf_import' : null);

const exists = (p) => { try { return !!p && fs.statSync(p).isFile(); } catch { return false; } };
const run = (cmd, args, opts = {}) => new Promise((resolve) => {
  execFile(cmd, args, { timeout: opts.timeout || 15000, windowsHide: true, maxBuffer: 4 * 1048576, ...opts }, (err, stdout, stderr) => resolve({ code: err ? (typeof err.code === 'number' ? err.code : 1) : 0, stdout: String(stdout || ''), stderr: String(stderr || ''), err }));
});

/** soffice on PATH (Windows: soffice.exe/.com) */
function onPath(names) {
  const dirs = String(process.env.PATH || '').split(path.delimiter).filter(Boolean);
  for (const d of dirs) for (const n of names) { const p = path.join(d.replace(/^"|"$/g, ''), n); if (exists(p)) return p; }
  return null;
}
function loCandidates(engineDir) {
  const pf = [process.env.ProgramFiles, process.env['ProgramFiles(x86)'], 'C:\\Program Files', 'C:\\Program Files (x86)'].filter(Boolean);
  const list = [];
  if (engineDir) list.push(...engineFiles(engineDir));
  if (WIN) for (const p of pf) list.push(path.join(p, 'LibreOffice', 'program', 'soffice.exe'));
  if (MAC) list.push('/Applications/LibreOffice.app/Contents/MacOS/soffice', path.join(os.homedir(), 'Applications/LibreOffice.app/Contents/MacOS/soffice'));
  if (!WIN && !MAC) list.push('/usr/bin/soffice', '/usr/bin/libreoffice', '/usr/local/bin/soffice', '/opt/libreoffice/program/soffice', '/snap/bin/libreoffice', '/var/lib/flatpak/exports/bin/org.libreoffice.LibreOffice');
  return list;
}
/** Where soffice sits inside the private engine folder (Windows admin install / Linux deb extract / macOS app) */
function engineFiles(dir) {
  if (WIN) return [path.join(dir, 'LibreOffice', 'program', 'soffice.exe'), path.join(dir, 'program', 'soffice.exe')];
  if (MAC) return [path.join(dir, 'LibreOffice.app', 'Contents', 'MacOS', 'soffice')];
  const out = [path.join(dir, 'program', 'soffice')];
  try { for (const d of fs.readdirSync(path.join(dir, 'opt'))) if (/^libreoffice/.test(d)) out.push(path.join(dir, 'opt', d, 'program', 'soffice')); } catch { /* not there */ }
  return out;
}
/** Version of a LibreOffice installation: program/version.ini (Windows/Linux) or Info.plist / --version */
async function loVersion(exe) {
  const prog = path.dirname(exe);
  for (const f of [path.join(prog, 'version.ini'), path.join(prog, 'versionrc'), path.join(prog, '..', 'Resources', 'version.plist')]) {
    try {
      const t = fs.readFileSync(f, 'utf8');
      const m = t.match(/MsiProductVersion=([\d.]+)/) || t.match(/ProductVersion=([\d.]+)/) || t.match(/<string>(\d+\.\d+\.\d+[\d.]*)<\/string>/);
      if (m) return m[1];
    } catch { /* next */ }
  }
  if (WIN) return null; // soffice --version prints nothing on Windows
  const r = await run(exe, ['--version'], { timeout: 20000 });
  const m = r.stdout.match(/LibreOffice\s+([\d.]+)/);
  return m ? m[1] : null;
}
/** → { path, version, source: 'settings' | 'engine' | 'system' } or null */
async function detectLibreOffice({ override, engineDir } = {}) {
  if (override && exists(override)) return { path: override, version: await loVersion(override), source: 'settings' };
  const engine = engineDir ? engineFiles(engineDir).find(exists) : null;
  const sys = onPath(WIN ? ['soffice.exe', 'soffice.com'] : ['soffice', 'libreoffice']) || loCandidates(null).find(exists);
  // System LibreOffice first (kept up to date by the user), the private engine otherwise
  const p = sys || engine;
  if (!p) return null;
  let version = await loVersion(p);
  // the administrative (portable) install leaves version.ini without a version – the engine remembers it itself
  if (!version && p === engine) { try { version = JSON.parse(fs.readFileSync(path.join(engineDir, 'engine.json'), 'utf8')).version || null; } catch { /* unknown */ } }
  return { path: p, version, source: p === engine ? 'engine' : 'system' };
}

/* ---------------- Microsoft Office ---------------- */
async function regAppPath(exe) {
  for (const hive of ['HKLM', 'HKCU']) {
    const r = await run('reg.exe', ['query', `${hive}\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${exe}`, '/ve'], { timeout: 8000 });
    const m = r.stdout.match(/REG_SZ\s+(.+\.exe)/i);
    if (m && exists(m[1].trim().replace(/^"|"$/g, ''))) return m[1].trim().replace(/^"|"$/g, '');
  }
  return null;
}
/** → { word, excel, powerpoint, visio } (paths or null) or null when none / not Windows */
async function detectMsOffice() {
  if (!WIN) return null;
  const pf = [...new Set([process.env.ProgramFiles, process.env['ProgramFiles(x86)'], 'C:\\Program Files', 'C:\\Program Files (x86)'].filter(Boolean))];
  const dirs = [];
  for (const p of pf) for (const v of ['Office16', 'Office15', 'Office14']) { dirs.push(path.join(p, 'Microsoft Office', 'root', v)); dirs.push(path.join(p, 'Microsoft Office', v)); }
  const out = {};
  for (const [app, d] of Object.entries(MS)) {
    out[app] = dirs.map((x) => path.join(x, d.exe)).find(exists) || await regAppPath(d.exe);
  }
  return Object.values(out).some(Boolean) ? out : null;
}

/* ---------------- Plan ---------------- */
/** Which MS app can do from → to? → app name or null */
function msApp(from, to, ms) {
  if (!ms) return null;
  const order = KIND_OF(from) === 'sheets' ? ['excel', 'word', 'powerpoint', 'visio'] : KIND_OF(from) === 'slides' ? ['powerpoint', 'word', 'excel', 'visio'] : VISIO_OPEN.includes(from) ? ['visio'] : ['word', 'excel', 'powerpoint', 'visio'];
  return order.find((a) => ms[a] && MS[a].open.includes(from) && MS[a].save[to] !== undefined && !(a === 'visio' && to === from)) || null;
}
const loCan = (from, to, lo) => !!lo && LO_IN.includes(from) && LO_OUT.includes(to);
/**
 * Converter for from → to. avail: { lo, ms, msEnabled }, preferred: auto | libreoffice | msoffice
 * → { tool: 'lo' } | { tool: 'ms', app } | null
 */
function plan(from, to, avail, preferred = 'auto') {
  const f = String(from || '').toLowerCase(); const t = String(to || '').toLowerCase();
  const ms = avail.msEnabled === false ? null : avail.ms;
  const app = msApp(f, t, ms);
  const lo = loCan(f, t, avail.lo);
  if (preferred === 'libreoffice' && lo) return { tool: 'lo' };
  if (preferred === 'msoffice' && app) return { tool: 'ms', app };
  // auto: Microsoft formats → MS Office (best fidelity), OpenDocument & the rest → LibreOffice
  const fam = (require('./formats').FORMATS[f] || {}).family;
  const odf = fam === 'odf' || /^(odt|ods|odp|odg)$/.test(t);
  if (app && (!odf || !lo)) return { tool: 'ms', app };
  if (lo) return { tool: 'lo' };
  if (app) return { tool: 'ms', app };
  return null;
}

/* ---------------- Running ---------------- */
const queues = new Map(); // tool/app → Promise (one conversion at a time)
function queued(key, fn) {
  const prev = queues.get(key) || Promise.resolve();
  const next = prev.catch(() => {}).then(fn);
  queues.set(key, next.catch(() => {}));
  return next;
}
const children = new Set(); // running processes (killed on shutdown)
function spawnTimed(cmd, args, { timeout, env, cwd, onSpawn } = {}) {
  return new Promise((resolve) => {
    let out = ''; let err = ''; let done = false;
    const p = spawn(cmd, args, { windowsHide: true, env: env || process.env, cwd, stdio: ['ignore', 'pipe', 'pipe'] });
    children.add(p);
    if (onSpawn) onSpawn(p);
    const t = setTimeout(() => { if (!done) { try { p.kill(); } catch { /* weg */ } if (WIN && p.pid) execFile('taskkill.exe', ['/PID', String(p.pid), '/T', '/F'], { windowsHide: true }, () => {}); } }, timeout);
    p.stdout.on('data', (c) => { out += c; if (out.length > 1e6) out = out.slice(-1e6); });
    p.stderr.on('data', (c) => { err += c; if (err.length > 1e6) err = err.slice(-1e6); });
    const fin = (code, e) => { if (done) return; done = true; clearTimeout(t); children.delete(p); resolve({ code, stdout: out, stderr: err || (e ? e.message : ''), timedOut: code === null }); };
    p.on('error', (e) => fin(-1, e));
    p.on('close', (code) => fin(code));
  });
}
const tmpDir = () => fs.mkdtempSync(path.join(os.tmpdir(), 'hw-office-'));
const failure = (msg, status = 500, vars) => Object.assign(new Error(msg), { status, vars, raw: msg });

/** LibreOffice: input file → output file in outDir */
async function runLo(exe, input, from, to, outDir, { timeout, profileDir }) {
  const profile = pathToFileURL(profileDir).href;
  const infilter = loInfilter(from, to);
  const args = [`-env:UserInstallation=${profile}`, '--headless', '--invisible', '--nologo', '--nodefault', '--norestore', '--nolockcheck', ...(infilter ? [`--infilter=${infilter}`] : []), '--convert-to', loFilter(from, to), '--outdir', outDir, input];
  const env = { ...process.env, SAL_USE_VCLPLUGIN: 'svp', HOME: process.env.HOME || os.homedir() };
  const r = await spawnTimed(exe, args, { timeout, env, cwd: path.dirname(exe) });
  if (r.timedOut) throw failure('The conversion took too long and was stopped.', 504);
  const files = fs.readdirSync(outDir).filter((f) => !f.startsWith('.'));
  if (!files.length) throw failure('LibreOffice could not convert the file. {detail}', 422, { detail: (r.stderr || r.stdout).trim().split('\n').slice(-2).join(' ') });
  return path.join(outDir, files.find((f) => f.toLowerCase().endsWith(`.${to}`)) || files[0]);
}

const PS_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$job = Get-Content -LiteralPath $env:HW_OFFICE_JOB -Raw -Encoding UTF8 | ConvertFrom-Json
$in = [string]$job.input; $out = [string]$job.output; $app = [string]$job.app; $fmt = $job.format
$names = @{ word = 'WINWORD'; excel = 'EXCEL'; powerpoint = 'POWERPNT'; visio = 'VISIO' }
$procName = $names[$app]
$before = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$o = $null; $doc = $null; $own = $false; $code = 0
try {
  switch ($app) {
    'word' { $o = New-Object -ComObject Word.Application }
    'excel' { $o = New-Object -ComObject Excel.Application }
    'powerpoint' { $o = New-Object -ComObject PowerPoint.Application }
    'visio' { $o = New-Object -ComObject Visio.InvisibleApp }
  }
  Start-Sleep -Milliseconds 200
  $new = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Where-Object { $before -notcontains $_ })
  if ($new.Count -gt 0) { $own = $true; Set-Content -LiteralPath $job.pidFile -Value ($new -join ',') -Encoding ASCII }
  try { $o.AutomationSecurity = 3 } catch {}
  switch ($app) {
    'word' {
      $o.Visible = $false; $o.DisplayAlerts = 0
      try { $o.Options.ConfirmConversions = $false } catch {}
      $doc = $o.Documents.Open($in, $false, $true, $false, 'hw-no-password', 'hw-no-password')
      if ([string]$fmt -eq '17') { $doc.ExportAsFixedFormat($out, 17) }
      elseif ([string]$fmt -eq '2') { $doc.SaveAs2($out, 2, $false, '', $false, '', $false, $false, $false, $false, $false, 65001) }
      else { $doc.SaveAs2($out, [int]$fmt) }
      $doc.Close(0)
    }
    'excel' {
      $o.Visible = $false; $o.DisplayAlerts = $false; $o.AskToUpdateLinks = $false; $o.EnableEvents = $false
      $doc = $o.Workbooks.Open($in, 0, $true, [Type]::Missing, 'hw-no-password')
      try { $doc.CheckCompatibility = $false } catch {}
      if ([string]$fmt -eq 'pdf') { $doc.ExportAsFixedFormat(0, $out) }
      elseif ([string]$fmt -eq 'xps') { $doc.ExportAsFixedFormat(1, $out) }
      else { $doc.SaveAs($out, [int]$fmt) }
      $doc.Close($false)
    }
    'powerpoint' {
      try { $o.DisplayAlerts = 1 } catch {}
      $doc = $o.Presentations.Open($in, -1, 0, 0)
      $doc.SaveAs($out, [int]$fmt)
      $doc.Close()
    }
    'visio' {
      try { $o.AlertResponse = 7 } catch {}
      $doc = $o.Documents.OpenEx($in, 2 + 8 + 128)
      if ([string]$fmt -eq 'pdf') { $doc.ExportAsFixedFormat(1, $out, 1, 0) }
      elseif ([string]$fmt -eq 'xps') { $doc.ExportAsFixedFormat(2, $out, 1, 0) }
      else { $doc.SaveAs($out) }
      $doc.Close()
    }
  }
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  $code = 2
} finally {
  if ($doc) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {} }
  if ($o) {
    if ($own) { try { $o.Quit() } catch {} }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {}
  }
  $doc = $null; $o = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
exit $code
`;

/** MS Office through COM: input file → output file */
async function runMs(app, input, to, outFile, { timeout, scriptDir }) {
  const fmt = MS[app].save[to];
  const dir = path.dirname(outFile);
  const script = path.join(scriptDir, 'office-convert.ps1');
  if (!fs.existsSync(script) || fs.readFileSync(script, 'utf8') !== PS_SCRIPT) { fs.mkdirSync(scriptDir, { recursive: true }); fs.writeFileSync(script, PS_SCRIPT); }
  const jobFile = path.join(dir, 'job.json'); const pidFile = path.join(dir, 'pid.txt');
  fs.writeFileSync(jobFile, JSON.stringify({ input, output: outFile, app, format: fmt, pidFile }), 'utf8');
  const env = { ...process.env, HW_OFFICE_JOB: jobFile };
  const r = await spawnTimed('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script], { timeout, env });
  if (r.timedOut || r.code !== 0) {
    // our own (new) Office process is left over → end it (never the user's)
    try { for (const pid of fs.readFileSync(pidFile, 'utf8').split(',').map((x) => Number(x.trim())).filter(Boolean)) execFile('taskkill.exe', ['/PID', String(pid), '/T', '/F'], { windowsHide: true }, () => {}); } catch { /* no pid file */ }
  }
  if (r.timedOut) throw failure('The conversion took too long and was stopped.', 504);
  if (r.code !== 0 || !exists(outFile)) throw failure('Microsoft Office could not convert the file. {detail}', 422, { detail: r.stderr.trim().split('\n').slice(-2).join(' ') });
  return outFile;
}

/**
 * Convert bytes. ctx: { detect: async () => ({ lo, ms, msEnabled }), preferred, timeoutSec, profileDir, scriptDir }
 * → { data: Buffer, tool, app? }
 */
async function convert(bytes, from, to, ctx) {
  const f = String(from || '').toLowerCase().replace(/^\./, ''); const t = String(to || '').toLowerCase().replace(/^\./, '');
  if (!/^[a-z0-9]{1,8}$/.test(f) || !/^[a-z0-9]{1,8}$/.test(t)) throw failure('Unknown format', 400);
  const avail = await ctx.detect();
  const p = plan(f, t, avail, ctx.preferred);
  if (!p) {
    const none = !avail.lo && !(avail.ms && avail.msEnabled !== false);
    throw failure(none ? 'No converter is available. Set up the Office engine (LibreOffice) in the Office settings or install Microsoft Office.' : 'This conversion ({from} → {to}) is not possible with the converters on this PC.', 501, { from: f, to: t });
  }
  const dir = tmpDir();
  try {
    const inDir = path.join(dir, 'in'); const outDir = path.join(dir, 'out');
    fs.mkdirSync(inDir); fs.mkdirSync(outDir);
    const input = path.join(inDir, `document.${f}`);
    fs.writeFileSync(input, bytes);
    const timeout = Math.max(10, Number(ctx.timeoutSec) || 120) * 1000;
    let file;
    if (p.tool === 'lo') file = await queued('lo', () => runLo(avail.lo.path, input, f, t, outDir, { timeout, profileDir: ctx.profileDir }));
    else file = await queued(`ms:${p.app}`, () => runMs(p.app, input, t, path.join(outDir, `document.${t}`), { timeout, scriptDir: ctx.scriptDir }));
    return { data: fs.readFileSync(file), tool: p.tool === 'lo' ? 'libreoffice' : 'msoffice', app: p.app };
  } finally {
    setTimeout(() => fs.rm(dir, { recursive: true, force: true }, () => {}), 500).unref();
  }
}

/** All targets reachable from an extension with the given converters → [ext] */
function targets(from, avail) {
  const set = new Set();
  const f = String(from || '').toLowerCase();
  if (avail.lo && LO_IN.includes(f)) for (const t of LO_OUT) set.add(t);
  if (avail.ms && avail.msEnabled !== false) for (const [app, d] of Object.entries(MS)) if (avail.ms[app] && d.open.includes(f)) for (const t of Object.keys(d.save)) set.add(t);
  set.delete(f);
  return [...set];
}

function killAll() { for (const p of children) { try { p.kill(); } catch { /* weg */ } } children.clear(); }
const hash = (b) => crypto.createHash('sha1').update(b).digest('hex');

module.exports = { detectLibreOffice, detectMsOffice, plan, convert, targets, loFilter, loInfilter, engineFiles, MS, LO_IN, LO_OUT, killAll, spawnTimed, PS_SCRIPT, hash };
