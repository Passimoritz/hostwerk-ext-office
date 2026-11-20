'use strict';
/*
 * Default app on Windows (server/osdefault.js): reading the user's choice (UserChoice / UserChoiceLatest / the
 * UserChoiceLatest\ProgId subkey of Windows 11 24H2+), merging it into the core's state, and the exact command lines
 * of Windows' "Open with" dialog and the Settings link. Nothing here touches the real registry or starts a program.
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { OsDefault, parseFileExts, mergeChoices, openWithCommand, settingsCommand, KIND_EXT } = require('../server/osdefault');
const { OsReg, kindStates } = require('../server/osreg');

const FE = 'HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts';
// shaped exactly like `reg query <FileExts> /s` on Windows 11 26H2 (German), 2026-11-20
const REG_OUT = [
  '',
  `${FE}\\.docx`,
  '',
  `${FE}\\.docx\\OpenWithList`,
  '    a    REG_SZ    WINWORD.EXE',
  '    MRUList    REG_SZ    a',
  '',
  `${FE}\\.docx\\OpenWithProgids`,
  '    Word.Document.12    REG_NONE    ',
  '',
  `${FE}\\.pdf\\UserChoice`,
  '    ProgId    REG_SZ    ChromePDF',
  '    Hash    REG_SZ    nkdJipRtyes=',
  '',
  `${FE}\\.pdf\\UserChoiceLatest`,
  '    Hash    REG_SZ    npk4fCvhL+E=',
  '',
  `${FE}\\.pdf\\UserChoiceLatest\\ProgId`,
  '    ProgId    REG_SZ    Hostwerk.office.pdf',
  '',
  `${FE}\\.xlsx\\UserChoice`,
  '    ProgId    REG_SZ    Hostwerk.office.xlsx',
  '    Hash    REG_SZ    abc=',
  '',
  `${FE}\\.PPTX\\UserChoiceLatest`,
  '    ProgId    REG_SZ    Hostwerk.office.pptx',
  '    Hash    REG_SZ    def=',
  '',
  `${FE}\\.odt\\UserChoice`,
  '    ProgId    REG_SZ    Hostwerk.office.odt',
  '',
  `${FE}\\.odt\\UserChoiceLatest\\ProgId`,
  '    ProgId    REG_SZ    LibreOffice.WriterDocument.1',
  '',
].join('\r\n');

test('UserChoiceLatest\\ProgId beats UserChoiceLatest beats UserChoice; untouched types have no choice', () => {
  const m = parseFileExts(REG_OUT);
  assert.strictEqual(m.get('pdf'), 'Hostwerk.office.pdf', 'Windows 11 24H2+: subkey ProgId wins over the older UserChoice');
  assert.strictEqual(m.get('xlsx'), 'Hostwerk.office.xlsx', 'Windows 10: UserChoice only');
  assert.strictEqual(m.get('pptx'), 'Hostwerk.office.pptx', 'UserChoiceLatest value, extension lower-cased');
  assert.strictEqual(m.get('odt'), 'LibreOffice.WriterDocument.1', 'the newest place decides – also against Hostwerk');
  assert.strictEqual(m.has('docx'), false, 'only OpenWithList/OpenWithProgids – no choice');
  assert.strictEqual(parseFileExts('').size, 0);
});

test('merge: the real choice turns "Open with" into default – also when the core read nothing', () => {
  const core = [
    { ext: 'pdf', kind: 'pdf', progId: 'Hostwerk.office.pdf', registered: true, isDefault: false },
    { ext: 'docx', kind: 'docs', progId: 'Hostwerk.office.docx', registered: true, isDefault: false },
    { ext: 'odt', kind: 'docs', progId: 'Hostwerk.office.odt', registered: true, isDefault: false },
    { ext: 'xlsx', kind: 'sheets', registered: true, isDefault: true },
  ];
  const t = mergeChoices(core, parseFileExts(REG_OUT));
  assert.deepStrictEqual(t.map((x) => [x.ext, x.isDefault]), [['pdf', true], ['docx', false], ['odt', false], ['xlsx', true]]);
  assert.strictEqual(mergeChoices(core, null), core, 'no reading → the core decides');
  const k = kindStates(t);
  assert.strictEqual(k.pdf.state, 'default');
  assert.strictEqual(k.docs.state, 'openwith');
  assert.strictEqual(k.sheets.state, 'default');
});

test('command lines: OpenWith.exe with the file as one argument, rundll32 fallback unquoted, Settings via ShellExecute', () => {
  const f = 'C:\\Users\\Anna Berger\\Hostwerk\\data\\office-data\\openwith\\Hostwerk Office sample.docx';
  assert.deepStrictEqual(openWithCommand(f), { cmd: 'C:\\Windows\\System32\\OpenWith.exe', args: [f], verbatim: false });
  assert.deepStrictEqual(openWithCommand(f, { hasOpenWith: false, systemRoot: 'D:\\Win' }), { cmd: 'D:\\Win\\System32\\rundll32.exe', args: [`shell32.dll,OpenAs_RunDLL ${f}`], verbatim: true });
  assert.throws(() => openWithCommand('relative\\x.docx'));
  assert.throws(() => openWithCommand('C:\\x".docx'));
  assert.throws(() => openWithCommand('C:\\x\n.docx'));
  const s = settingsCommand('Hostwerk Office');
  assert.deepStrictEqual(s.args, ['url.dll,FileProtocolHandler', 'ms-settings:defaultapps?registeredAppUser=Hostwerk Office'], 'the space stays – Settings does not decode %20');
  assert.strictEqual(s.cmd, 'C:\\Windows\\System32\\rundll32.exe');
  assert.strictEqual(settingsCommand('a&b"c%').uri, 'ms-settings:defaultapps?registeredAppUser=abc', 'no characters a shell or URI could read');
  assert.deepStrictEqual(KIND_EXT, { docs: 'docx', sheets: 'xlsx', slides: 'pptx', pdf: 'pdf' });
});

function fake(platform = 'win32') {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'of-osd-'));
  const calls = { run: [], spawn: [] };
  const osd = new OsDefault({ dir: path.join(dir, 'openwith'), platform, systemRoot: dir,
    run: async (cmd, args) => { calls.run.push([path.basename(cmd), ...args]); return { code: 0, stdout: REG_OUT }; },
    spawn: (cmd, args, opts) => { calls.spawn.push({ cmd: path.basename(cmd), args, opts }); } });
  return { dir, osd, calls };
}

test('OsDefault: one reg query for all types (cached), sample file written, dialog started – never on other systems', async () => {
  const { dir, osd, calls } = fake();
  const a = await osd.choices();
  await osd.choices();
  assert.strictEqual(calls.run.length, 1, 'cached between two polls');
  assert.deepStrictEqual(calls.run[0], ['reg.exe', 'query', FE, '/s']);
  assert.strictEqual(a.get('pdf'), 'Hostwerk.office.pdf');
  const bytes = Buffer.from('PK fake docx');
  const r = osd.openWith('docx', bytes, 'Hostwerk Office sample');
  assert.ok(fs.readFileSync(r.file).equals(bytes), 'the empty document from the editor chunk');
  assert.strictEqual(path.basename(r.file), 'Hostwerk Office sample.docx');
  // no OpenWith.exe in the fake system folder → rundll32 with the unquoted path
  assert.strictEqual(calls.spawn[0].cmd, 'rundll32.exe');
  assert.deepStrictEqual(calls.spawn[0].args, [`shell32.dll,OpenAs_RunDLL ${r.file}`]);
  assert.strictEqual(calls.spawn[0].opts.windowsVerbatimArguments, true);
  fs.mkdirSync(path.join(dir, 'System32'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'System32', 'OpenWith.exe'), '');
  osd.openWith('pdf');
  assert.strictEqual(calls.spawn[1].cmd, 'OpenWith.exe');
  assert.strictEqual(fs.statSync(path.join(dir, 'openwith', 'Hostwerk Office.pdf')).size, 0, 'without bytes: an empty placeholder');
  assert.throws(() => osd.openWith('../x'), /unknown type/);
  const s = osd.openSettings('Hostwerk Office');
  assert.strictEqual(s.uri, 'ms-settings:defaultapps?registeredAppUser=Hostwerk Office');
  const lin = fake('linux');
  assert.strictEqual(await lin.osd.choices(), null);
  assert.throws(() => lin.osd.openWith('docx'), (e) => e.code === 'ENOTWINDOWS');
  assert.strictEqual(lin.calls.run.length + lin.calls.spawn.length, 0);
});

test('OsReg with OsDefault: state shows the real choice, openWith registers first and opens the dialog of the kind', async () => {
  const { dir, osd, calls } = fake();
  let reg = null;
  const fileTypes = {
    register: async (def) => { reg = def; return {}; },
    unregister: async () => { reg = null; return {}; },
    state: async () => (reg ? { registered: true, types: reg.types.map((x) => ({ ext: x.ext, kind: x.kind, progId: `Hostwerk.office.${x.ext}`, registered: true, isDefault: false })) } : { registered: false, types: [] }),
    openDefaultsSettings: async () => { throw new Error('the core path must not be used on Windows'); },
  };
  const host = { version: '1.3.1', t: (s) => s, supports: (n) => n === 'fileTypes', fileTypes };
  const settings = { integrations: { osOpenWith: true } };
  const r = new OsReg({ host, file: path.join(dir, 'os.json'), settings: () => settings, osdefault: osd });
  const o = await r.openWith('slides', Buffer.from('x'));
  assert.ok(reg, 'registered before the dialog (Hostwerk Office must be in its list)');
  assert.strictEqual(o.kind, 'slides');
  assert.ok(/\.pptx$/.test(o.file));
  assert.strictEqual(calls.spawn.length, 1);
  const st = await r.state();
  assert.strictEqual(st.openWithDialog, true);
  assert.strictEqual(st.kinds.pdf.state, 'default', '.pdf chosen (UserChoiceLatest\\ProgId)');
  assert.strictEqual(st.kinds.docs.state, 'openwith', 'nothing of the documents chosen');
  assert.strictEqual(st.kinds.slides.state, 'partial', '.pptx chosen (UserChoiceLatest)');
  assert.strictEqual(st.kinds.sheets.state, 'partial');
  await assert.rejects(() => r.openWith('viewer'), /unknown kind/);
  const d = await r.openDefaults();
  assert.strictEqual(d.uri, 'ms-settings:defaultapps?registeredAppUser=Hostwerk Office');
});
