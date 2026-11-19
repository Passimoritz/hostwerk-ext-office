'use strict';
/*
 * Open a file with the system's default app / show it in the file manager – arguments only, never a shell string.
 * Windows: explorer.exe <file> / explorer.exe /select,<file> (no cmd.exe – "&" in a name would be a command there).
 * Never a web browser (binding rule: everything stays inside Hostwerk): Google pointer files, links and web pages are
 * refused with 409 and a hint; executables and scripts are never opened.
 */
const { spawn } = require('child_process');
const path = require('path');

const BROWSER = /\.(gdoc|gsheet|gslides|gdraw|gform|gsite|gmap|gjam|glink|url|webloc|website|desktop|html?|xhtml|mht|mhtml|svg)$/i;
const EXEC = /\.(exe|com|bat|cmd|ps1|psm1|vbs|vbe|js|jse|wsf|wsh|msi|msp|scr|lnk|pif|reg|cpl|hta|jar|sh|app|command|py|rb|pl|appref-ms|msc|inf|scf|ws|gadget|application|library-ms|settingcontent-ms)$/i;

/** null = ok, else { status, error } why it must not be opened */
function refuse(file) {
  if (BROWSER.test(file)) return { status: 409, error: 'This file would open in a web browser – open it here in Office instead.' };
  if (EXEC.test(file)) return { status: 403, error: 'Programs and scripts are never opened from Office.' };
  return null;
}
function launch(cmd, args) {
  return new Promise((resolve, reject) => {
    const p = spawn(cmd, args, { detached: true, stdio: 'ignore', windowsHide: false });
    p.on('error', reject);
    p.unref();
    setTimeout(resolve, 150);
  });
}
async function open(file) {
  if (process.platform === 'win32') return launch('explorer.exe', [path.resolve(file)]);
  if (process.platform === 'darwin') return launch('open', ['--', file]);
  return launch('xdg-open', [file]);
}
async function reveal(file) {
  if (process.platform === 'win32') return launch('explorer.exe', [`/select,${path.resolve(file)}`]);
  if (process.platform === 'darwin') return launch('open', ['-R', '--', file]);
  return launch('xdg-open', [path.dirname(file)]);
}

module.exports = { open, reveal, refuse, BROWSER, EXEC };
