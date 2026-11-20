import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { readPpt, buildAnimations, tableFromGroup } from '../ui/src/editors/slides/io/ppt-read.js';

const SRC = fs.readFileSync(new URL('./fixtures/ppt-table-anim.ppt', import.meta.url));
const txt = (c) => c.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n');

test('ppt: a PowerPoint table arrives as a real table (not a group)', async () => {
  const { doc } = await readPpt(SRC, { ext: 'ppt' });
  const tbl = doc.slides[0].elements.find((e) => e.type === 'table');
  assert.ok(tbl, 'table element');
  assert.equal(tbl.rows.length, 2); assert.equal(tbl.cols.length, 3);
  assert.deepEqual(tbl.rows.map((r) => r.cells.map(txt)), [['Name', 'Qty', 'Price'], ['Apple', '3', '1.50']]);
  assert.ok(!doc.slides[0].elements.some((e) => e.type === 'group'));
  assert.equal(Math.round(tbl.cols.reduce((a, b) => a + b, 0)), tbl.w);
});

test('ppt: old build effects become animations', async () => {
  const { doc } = await readPpt(SRC, { ext: 'ppt' });
  const s = doc.slides[0];
  assert.equal(s.animations.length, 2);
  const [a, b] = s.animations;
  assert.equal(a.effect, 'fly'); assert.equal(a.dir, 'l'); assert.equal(a.trigger, 'click');
  assert.equal(b.trigger, 'after');
  assert.ok(s.elements.some((e) => e.id === a.el && /Fly/.test(txt(e))));
});

test('build list mapping and table grid edge cases', () => {
  const an = buildAnimations([{ el: 'x', flags: 0, delay: 0, order: 2, build: 1, effect: 6, dir: 0 }, { el: 'y', flags: 4, delay: 1500, order: 1, build: 1, effect: 10, dir: 0 }, { el: 'z', flags: 0, delay: 0, order: -1, build: 0, effect: 0, dir: 0 }]);
  assert.deepEqual(an.map((a) => [a.el, a.effect, a.trigger, a.delay]), [['y', 'wipe', 'after', 1500], ['x', 'fade', 'click', 0]]);
  const box = (x, y, w, h, t = '') => ({ type: 'shape', x, y, w, h, text: { paras: [{ runs: [{ text: t }], level: 0 }] } });
  const t = tableFromGroup([box(0, 0, 200000, 100000, 'merged'), box(0, 100000, 100000, 100000, 'a'), box(100000, 100000, 100000, 100000, 'b')]);
  assert.deepEqual(t.rows[0].cells[0].span, { r: 1, c: 2 }); assert.equal(t.rows[0].cells[1].merged, true);
  assert.equal(tableFromGroup([box(0, 0, 10, 10), { type: 'image', x: 0, y: 0, w: 1, h: 1 }]), null);
});
