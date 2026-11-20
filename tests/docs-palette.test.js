/*
 * Docs command palette (1.2): the ribbon's formatting, styles, alignment, lists, insert, page setup, review and zoom
 * commands are in the palette with localised shortcut labels; every command calls something the editor really has.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { paletteCommands, allPaletteCommands } from '../ui/src/editors/docs/palette.js';
import { keyLabel } from '../ui/src/editors/_shared/keys.js';

const src = fs.readFileSync(new URL('../ui/src/editors/docs/Editor.jsx', import.meta.url), 'utf8');
const actBlock = src.slice(src.indexOf('const act = {'), src.indexOf('actRef.current = act;'));

/** a fake editor that records the chain it gets */
function fakeEditor() {
  const calls = [];
  const chain = new Proxy({}, { get: (o, k) => (k === 'run' ? () => { calls.push('run'); return true; } : (...args) => { calls.push([k, ...args]); return chain; }) });
  return { calls, chain: () => chain, isActive: () => false, commands: new Proxy({}, { get: (o, k) => (...args) => { calls.push([k, ...args]); return true; } }) };
}

test('palette: formatting, styles, alignment, lists, insert, page setup, track changes and zoom are there', () => {
  const list = paletteCommands();
  const ids = list.map((c) => c.id);
  assert.strictEqual(new Set(ids).size, ids.length, 'ids unique');
  for (const want of ['bold', 'italic', 'underline', 'strike', 'clear', 'normal', 'left', 'center', 'right', 'justify', 'indent', 'tablesize', 'imagefile', 'pagesetup', 'landscape', 'track', 'acceptall', 'zoomin', 'zoomout', 'zoom100', 'fit']) assert.ok(ids.includes(want), want);
  for (const c of list) {
    assert.ok(c.title && c.group && typeof c.run === 'function' && c.words, c.id);
    if (c.keys) assert.ok(!/Ctrl\+(K|J|,|[1-9])$/.test(c.keys), `${c.id}: no Hostwerk global key`);
  }
  // shortcut labels in German
  const byId = Object.fromEntries(list.map((c) => [c.id, c]));
  assert.strictEqual(keyLabel(byId.bold.keys, 'de'), 'Strg+B');
  assert.strictEqual(keyLabel(byId.track.keys, 'de'), 'Strg+Umschalt+E');
  assert.strictEqual(keyLabel(byId.clear.keys, 'de'), 'Strg+Leertaste');
  // with the "/" commands: no duplicate ids
  const all = allPaletteCommands([{ id: 'bold', title: 'x', group: 'y', run: () => {} }, { id: 'h1', title: 'Heading 1', group: 'Basic', run: () => {} }]);
  assert.strictEqual(all.filter((c) => c.id === 'bold').length, 1);
  assert.ok(all.some((c) => c.id === 'h1'));
});

test('palette: every command uses an existing act method or editor command', () => {
  for (const c of paletteCommands()) {
    const used = [];
    const act = new Proxy({}, { get: (o, k) => (...args) => { used.push(String(k)); return args; } });
    const ed = fakeEditor();
    c.run(act, ed);
    for (const m of used) assert.match(actBlock, new RegExp(`\\n\\s{4}${m}[:,(]|\\n\\s{4}${m}\\s*[,:]|[ {]${m},`), `${c.id}: act.${m} exists in Editor.jsx`);
    if (!used.length) assert.ok(ed.calls.length, `${c.id} does something`);
  }
  // the editor registers the full list with the shell (words for the ranking)
  assert.match(src, /host\.palette\.register\(allPaletteCommands\(commands\)\.map\(\(c\) => [\s\S]{0,200}words: c\.words/);
});
