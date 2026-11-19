// Shortcut labels and colour names (ui/src/editors/_shared/keys.js, colors.js) – node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { keyLabel, withKeys, keysInText } from '../ui/src/editors/_shared/keys.js';
import { colorName, COLOR_NAMES } from '../ui/src/editors/_shared/colors.js';

test('keyLabel in German and English', () => {
  assert.equal(keyLabel('Ctrl+Shift+S', 'de'), 'Strg+Umschalt+S');
  assert.equal(keyLabel('Ctrl+Shift+S', 'en'), 'Ctrl+Shift+S');
  assert.equal(keyLabel('Del', 'de'), 'Entf');
  assert.equal(keyLabel('Ctrl++', 'de'), 'Strg++');
  assert.equal(keyLabel('Ctrl+-', 'de'), 'Strg+-');
  assert.equal(keyLabel('Ctrl+Shift+=', 'de'), 'Strg+Umschalt+=');
  assert.equal(keyLabel('Shift+Enter', 'de'), 'Umschalt+Eingabe');
  assert.equal(keyLabel('Ctrl+PageDown', 'de'), 'Strg+Bild ab');
  assert.equal(keyLabel('Ctrl+Y / Ctrl+Shift+Z', 'de'), 'Strg+Y / Strg+Umschalt+Z');
  assert.equal(keyLabel('F2', 'de'), 'F2');
  assert.equal(keyLabel('Alt+Shift+5', 'de'), 'Alt+Umschalt+5');
  assert.equal(keyLabel('', 'de'), '');
  assert.equal(keyLabel(undefined, 'de'), '');
});

test('withKeys and keysInText', () => {
  assert.equal(withKeys('Speichern', 'Ctrl+S', 'de'), 'Speichern (Strg+S)');
  assert.equal(withKeys('Save', 'Ctrl+S', 'en'), 'Save (Ctrl+S)');
  assert.equal(withKeys('Fett', undefined, 'de'), 'Fett');
  assert.equal(keysInText('Press Ctrl+V or Del', 'de'), 'Press Strg+V or Entf');
  assert.equal(keysInText('Press Ctrl+V', 'en'), 'Press Ctrl+V');
});

test('colorName gives real names, never a hex code', () => {
  assert.equal(colorName('#ffd400'), 'Yellow');
  assert.equal(colorName('#FF0000'), 'Red');
  assert.equal(colorName('#ff5d3a'), 'Red');
  assert.equal(colorName('#111111'), 'Black');
  assert.equal(colorName('#ffffff'), 'White');
  assert.equal(colorName('#0070C0'), 'Blue');
  assert.equal(colorName('#00B050'), 'Green');
  assert.equal(colorName('#fff'), 'White');
  assert.equal(colorName('nonsense'), '');
  for (const c of ['#123456', '#abcdef', '#7f7f00', '#e0e0ff']) assert.ok(COLOR_NAMES.includes(colorName(c)), c);
});

test('every colour name has a German translation', () => {
  const de = JSON.parse(readFileSync(new URL('../locales/de.json', import.meta.url), 'utf8'));
  const missing = COLOR_NAMES.filter((n) => !de[n]);
  assert.deepEqual(missing, []);
});
