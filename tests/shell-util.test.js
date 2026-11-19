// Pure helpers of the Office shell (ui/src/shell/util.js) – node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  basename, dirname, joinPath, splitExt, crumbsOf, validateName, ensureExt, uniqueName, sortEntries, filterEntries, countBy, highlightParts,
  clickSelect, moveIndex, defaultExt, formatChoices, nativeExtFor, convertTargets, openPlan, parseQuery, toQuery, editRoute, fmtBytes, shareExpiry,
  parseTags, saveStateText, fileKey, fold, openTarget,
} from '../ui/src/shell/util.js';

test('paths', () => {
  assert.equal(basename('a/b/c.docx'), 'c.docx');
  assert.equal(dirname('a/b/c.docx'), 'a/b');
  assert.equal(dirname('c.docx'), '');
  assert.equal(joinPath('a/', 'b.txt'), 'a/b.txt');
  assert.equal(joinPath('', 'b.txt'), 'b.txt');
  assert.deepEqual(splitExt('Report 2024.DOCX'), { base: 'Report 2024', ext: 'docx' });
  assert.deepEqual(splitExt('README'), { base: 'README', ext: '' });
  assert.deepEqual(crumbsOf('a/b'), [{ name: 'a', dir: 'a' }, { name: 'b', dir: 'a/b' }]);
  assert.deepEqual(crumbsOf(''), []);
  assert.equal(fileKey({ root: 'htdocs', path: 'x.docx' }), 'htdocs|x.docx');
  assert.equal(fileKey({ path: 'x.docx' }), 'x.docx');
});

test('validateName', () => {
  assert.equal(validateName('Letter.docx'), null);
  assert.ok(validateName(''));
  assert.ok(validateName('   '));
  assert.ok(validateName('a/b'));
  assert.ok(validateName('a:b'));
  assert.ok(validateName('name.'));
  assert.ok(validateName('CON'));
  assert.ok(validateName('nul.txt'));
  assert.ok(validateName('x'.repeat(201)));
  assert.ok(validateName('Plan.xlsx', { existing: ['plan.XLSX'] }));
  assert.equal(validateName('Plan.xlsx', { existing: ['plan.xlsx'], current: 'Plan.xlsx' }), null);
});

test('ensureExt and uniqueName', () => {
  assert.equal(ensureExt('Letter', 'docx'), 'Letter.docx');
  assert.equal(ensureExt('Letter.docx', 'docx'), 'Letter.docx');
  assert.equal(ensureExt('a.odt', 'docx', { replace: true }), 'a.docx');
  assert.equal(ensureExt('v1.2', 'docx'), 'v1.2.docx');
  assert.equal(uniqueName('Report.docx', []), 'Report.docx');
  assert.equal(uniqueName('Report.docx', ['report.docx']), 'Report (2).docx');
  assert.equal(uniqueName('Report.docx', ['Report.docx', 'Report (2).docx']), 'Report (3).docx');
  assert.equal(uniqueName('Report (2).docx', ['Report (2).docx']), 'Report (3).docx');
});

const E = [
  { name: 'b.xlsx', mtime: 3, size: 10 },
  { name: 'A.docx', mtime: 1, size: 30, tags: ['Invoice'] },
  { name: 'Folder', dir: true, mtime: 0 },
  { name: 'c10.pdf', mtime: 2, size: 20, note: 'für Steuer' },
  { name: 'c9.pdf', mtime: 2, size: 5 },
];
test('sortEntries: folders first, numeric names, all keys', () => {
  assert.deepEqual(sortEntries(E, 'name', 'asc').map((e) => e.name), ['Folder', 'A.docx', 'b.xlsx', 'c9.pdf', 'c10.pdf']);
  assert.deepEqual(sortEntries(E, 'modified', 'desc').map((e) => e.name), ['Folder', 'b.xlsx', 'c9.pdf', 'c10.pdf', 'A.docx']);
  assert.deepEqual(sortEntries(E, 'size', 'asc').map((e) => e.name), ['Folder', 'c9.pdf', 'b.xlsx', 'c10.pdf', 'A.docx']);
  assert.deepEqual(sortEntries(E, 'type', 'asc').map((e) => e.name), ['Folder', 'A.docx', 'b.xlsx', 'c9.pdf', 'c10.pdf']);
  assert.equal(E[0].name, 'b.xlsx', 'does not mutate');
});
test('filterEntries and countBy', () => {
  assert.deepEqual(filterEntries(E, { q: 'invoice' }).map((e) => e.name), ['A.docx']);
  assert.deepEqual(filterEntries(E, { q: 'fur' }).map((e) => e.name), ['c10.pdf'], 'accent-insensitive note');
  assert.deepEqual(filterEntries(E, { kinds: ['pdf'] }).map((e) => e.name), ['c10.pdf', 'c9.pdf']);
  assert.deepEqual(filterEntries(E, { families: ['ms'] }).map((e) => e.name), ['b.xlsx', 'A.docx']);
  assert.equal(filterEntries(E, {}).length, 5);
  assert.deepEqual(countBy(E).kinds, { sheets: 1, docs: 1, pdf: 2 });
  assert.equal(fold('Übersicht'), 'ubersicht');
});
test('highlightParts', () => {
  assert.deepEqual(highlightParts('Annual report', 'rep'), [{ text: 'Annual ', hit: false }, { text: 'rep', hit: true }, { text: 'ort', hit: false }]);
  assert.deepEqual(highlightParts('<b>x</b> &amp; y', ''), [{ text: 'x & y', hit: false }]);
  assert.equal(highlightParts('Café au lait', 'cafe').filter((p) => p.hit)[0].text, 'Café');
  assert.equal(highlightParts('nothing', 'zzz').some((p) => p.hit), false);
});
test('clickSelect and moveIndex', () => {
  const order = ['a', 'b', 'c', 'd'];
  let r = clickSelect(order, new Set(), null, 'b');
  assert.deepEqual([...r.sel], ['b']);
  r = clickSelect(order, r.sel, r.anchor, 'd', { shift: true });
  assert.deepEqual([...r.sel], ['b', 'c', 'd']);
  r = clickSelect(order, r.sel, r.anchor, 'c', { ctrl: true });
  assert.deepEqual([...r.sel], ['b', 'd']);
  assert.equal(moveIndex(-1, 'ArrowDown', 4), 0);
  assert.equal(moveIndex(0, 'ArrowDown', 10, 4), 4);
  assert.equal(moveIndex(9, 'ArrowRight', 10, 4), 9);
  assert.equal(moveIndex(5, 'Home', 10), 0);
  assert.equal(moveIndex(5, 'End', 10), 9);
});
test('formats for new files and saving', () => {
  assert.equal(defaultExt('docs', 'ms'), 'docx');
  assert.equal(defaultExt('sheets', 'odf'), 'ods');
  const docs = formatChoices('docs', 'odf');
  assert.equal(docs[0].ext, 'odt'); assert.ok(docs[0].isDefault);
  assert.equal(docs[1].ext, 'docx');
  assert.ok(docs.every((d) => d.ext !== 'doc'), 'legacy formats are not offered for new files');
  assert.equal(nativeExtFor('docm'), 'docx');
  assert.equal(nativeExtFor('ott'), 'odt');
  assert.equal(nativeExtFor('doc'), 'docx');
  assert.equal(nativeExtFor('numbers'), 'xlsx');
  assert.equal(nativeExtFor('docx'), 'docx');
  assert.equal(openPlan('doc').mode, 'convert');
  assert.equal(openPlan('doc').openExt, 'docx');
  assert.equal(openPlan('docm').mode, 'import');
  assert.equal(openPlan('docm').saveExt, 'docx');
  assert.equal(openPlan('docx').mode, 'edit');
  assert.equal(openPlan('docx', { readOnly: true }).mode, 'view');
  assert.equal(openPlan('eml').mode, 'view');
  assert.equal(openPlan('xyz').mode, 'unknown');
});
test('convertTargets', () => {
  const t1 = convertTargets('docx', { hasConverter: false });
  assert.ok(!t1.some((x) => x.ext === 'docx'));
  assert.ok(t1.some((x) => x.ext === 'pdf' && x.via === 'client'));
  assert.ok(!t1.some((x) => x.ext === 'doc'));
  const t2 = convertTargets('docx', { hasConverter: true });
  assert.ok(t2.some((x) => x.ext === 'doc' && x.via === 'server'));
  assert.equal(t2.filter((x) => x.ext === 'pdf').length, 1, 'no duplicates');
  assert.deepEqual(convertTargets('vsdx', { hasConverter: true }).map((x) => x.ext), ['pdf']);
  assert.deepEqual(convertTargets('unknown'), []);
});
test('query strings and routes', () => {
  assert.deepEqual(parseQuery('?action=new&kind=docs&x'), { action: 'new', kind: 'docs', x: '' });
  assert.deepEqual(parseQuery('path=a%2Fb%20c.docx'), { path: 'a/b c.docx' });
  assert.equal(toQuery({ a: 1, b: '', c: null, d: true, e: false }), 'a=1&d=1');
  assert.equal(editRoute('office', { path: 'a/b.docx' }), '/ext/office/edit?path=a%2Fb.docx');
  assert.equal(editRoute('office', { root: 'mc:1', path: 'x.xlsx' }, { app: 'onlyoffice' }), '/ext/office/edit?root=mc%3A1&path=x.xlsx&app=onlyoffice');
});
test('small formatters', () => {
  assert.equal(fmtBytes(512), '512 B');
  assert.equal(fmtBytes(1536), '1.5 KB');
  assert.equal(fmtBytes(1048576), '1 MB');
  assert.equal(shareExpiry(0), null);
  assert.equal(shareExpiry(2, 1000), 1000 + 2 * 86400000);
  assert.deepEqual(parseTags('a, b ,A;c\n'), ['a', 'b', 'c']);
  assert.equal(saveStateText('dirty').text, 'Unsaved changes');
  assert.equal(saveStateText('saved', { savedAt: new Date(2026, 0, 1, 12, 3).getTime() }).vars.time, '12:03');
  assert.equal(saveStateText('error', { error: 'x' }).tone, 'err');
});

test('open links of quick search and AI (action=open)', () => {
  assert.deepEqual(openTarget({ target: 'Reports/2026.docx' }), { path: 'Reports/2026.docx' });
  assert.deepEqual(openTarget({ target: 'a/b.docx', root: 'own:ab12' }), { root: 'own:ab12', path: 'a/b.docx' });
  // older form <root>:<path> – root ids may contain a colon themselves
  assert.deepEqual(openTarget({ target: 'own:ab12:a/b.docx' }), { root: 'own:ab12', path: 'a/b.docx' });
  assert.deepEqual(openTarget({ target: 'htdocs:x.docx' }), { root: 'htdocs', path: 'x.docx' });
});
