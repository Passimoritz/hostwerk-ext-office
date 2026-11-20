// Settings model of the Office shell (ui/src/shell/settingsModel.js) – node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_SETTINGS, mergeSettings, withDefaults, getIn, setIn, diffSettings, patchOf, validateSettings, integrationReady } from '../ui/src/shell/settingsModel.js';

test('defaults cover every group of the contract', () => {
  for (const g of ['general', 'docs', 'sheets', 'slides', 'pdf', 'converters', 'integrations', 'sharing']) assert.ok(DEFAULT_SETTINGS[g], g);
  assert.equal(DEFAULT_SETTINGS.general.autosaveSeconds, 20);
  assert.equal(DEFAULT_SETTINGS.converters.autoEngine, 'ask');
});
test('mergeSettings / withDefaults', () => {
  const s = withDefaults({ general: { autosave: false }, docs: { margins: { top: 30 } } });
  assert.equal(s.general.autosave, false);
  assert.equal(s.general.autosaveSeconds, 20);
  assert.equal(s.docs.margins.top, 30);
  assert.equal(s.docs.margins.left, 25);
  assert.equal(DEFAULT_SETTINGS.docs.margins.top, 25, 'defaults untouched');
  assert.deepEqual(mergeSettings({ a: [1, 2] }, { a: [3] }), { a: [3] });
});
test('getIn / setIn / diff / patch', () => {
  const a = withDefaults(null);
  const b = setIn(a, 'docs.margins.top', 40);
  assert.equal(getIn(b, 'docs.margins.top'), 40);
  assert.equal(getIn(a, 'docs.margins.top'), 25);
  assert.equal(b.general, a.general, 'unchanged branches are shared');
  assert.deepEqual(diffSettings(a, b), ['docs.margins.top']);
  assert.deepEqual(patchOf(a, setIn(b, 'general.author', 'Anna')), { docs: { margins: { top: 40 } }, general: { author: 'Anna' } });
  assert.deepEqual(diffSettings(a, a), []);
});
test('validateSettings', () => {
  const ok = withDefaults(null);
  assert.deepEqual(validateSettings(ok), {});
  const bad = setIn(setIn(setIn(ok, 'general.autosaveSeconds', 1), 'integrations.onlyoffice.url', 'onlyoffice'), 'libraryDir', 'relative/dir');
  const e = validateSettings(bad);
  assert.equal(e['general.autosaveSeconds'].text, 'Between {min} and {max}.');
  assert.deepEqual(e['general.autosaveSeconds'].vars, { min: 5, max: 600 });
  assert.ok(e['integrations.onlyoffice.url']);
  assert.ok(e.libraryDir);
  assert.ok(validateSettings(setIn(ok, 'docs.fontSize', ''))['docs.fontSize']);
  assert.deepEqual(validateSettings(setIn(ok, 'libraryDir', 'D:\\Office')), {});
  assert.equal(integrationReady(setIn(ok, 'integrations.onlyoffice.url', 'https://oo'), 'onlyoffice'), true);
  assert.equal(integrationReady(ok, 'onlyoffice'), false);
});

// 1.3.1: sections of the settings page, which section a field belongs to, what saves only on leaving a field
import { SETTING_GROUPS, groupOfPath, LAZY_PATHS, GROUP_RESET } from '../ui/src/shell/settingsModel.js';
test('sections: unique ids with label + sentence; every validated field maps to an existing section', () => {
  const ids = SETTING_GROUPS.map((g) => g.id);
  assert.equal(new Set(ids).size, ids.length);
  for (const g of SETTING_GROUPS) assert.ok(g.label && g.text && /\.$/.test(g.text), g.id);
  assert.equal(ids[1], 'os', 'the default app right after General');
  const bad = validateSettings(setIn(setIn(setIn(DEFAULT_SETTINGS, 'libraryDir', 'relative'), 'integrations.onlyoffice.url', 'x'), 'docs.margins.top', 999));
  for (const p of Object.keys(bad)) assert.ok(ids.includes(groupOfPath(p)), p);
  assert.equal(groupOfPath('libraryDir'), 'storage');
  assert.equal(groupOfPath('general.trashDays'), 'storage');
  assert.equal(groupOfPath('general.autosave'), 'general');
  assert.equal(groupOfPath('integrations.onlyoffice.url'), 'fulleditor');
  assert.equal(groupOfPath('integrations.google.clientId'), 'google');
  assert.equal(groupOfPath('integrations.osOpenWith'), 'os');
  assert.equal(groupOfPath('docs.margins.top'), 'docs');
});
test('lazy paths and reset lists only name real settings', () => {
  for (const p of LAZY_PATHS) assert.notEqual(getIn(DEFAULT_SETTINGS, p), undefined, p);
  assert.ok(LAZY_PATHS.includes('libraryDir'), 'a half-typed path never moves the library');
  for (const [g, list] of Object.entries(GROUP_RESET)) { assert.ok(SETTING_GROUPS.some((x) => x.id === g), g); for (const p of list) assert.notEqual(getIn(DEFAULT_SETTINGS, p), undefined, p); }
});
