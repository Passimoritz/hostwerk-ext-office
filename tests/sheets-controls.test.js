import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import JSZip from 'jszip';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { readXlsx, writeXlsx } from '../ui/src/editors/sheets/io/xlsx.js';
import { applyControl, controlItems, controlState } from '../ui/src/editors/sheets/formcontrols.js';
import { recalcDataTables } from '../ui/src/editors/sheets/whatif.js';

const loc = makeLocale({ lang: 'en' });
const SRC = fs.readFileSync(new URL('./fixtures/excel-controls.xlsx', import.meta.url));
const K = (r, c) => r * 16384 + c;

test('form controls of an Excel file are read', async () => {
  const b = await readXlsx(SRC, { loc, ext: 'xlsx' });
  const s = b.sheets[0];
  const types = s.controls.map((c) => c.type);
  assert.deepEqual(types, ['checkbox', 'dropdown', 'option', 'option', 'spinner', 'button']);
  const [cb, dd, o1, o2, sp, bt] = s.controls;
  assert.equal(cb.text, 'Paid'); assert.equal(cb.checked, 1); assert.equal(cb.link, '$C$1');
  assert.equal(dd.sel, 2); assert.deepEqual(controlItems(b, s, dd), ['Apple', 'Pear', 'Plum']);
  assert.equal(o2.link, '$C$3'); assert.equal(o1.group, o2.group); assert.equal(o2.checked, 1);
  assert.equal(sp.val, 3); assert.equal(sp.max, 10);
  assert.equal(bt.text, 'Go');
  assert.ok(!b.fidelity.some((f) => /Form controls/.test(f)));
});

test('toggling / picking writes the linked cell and the control parts', async () => {
  const b = await readXlsx(SRC, { loc, ext: 'xlsx' });
  const s = b.sheets[0];
  const [cb, dd, o1, , sp] = s.controls;
  assert.ok(applyControl(b, s, cb.id, { toggle: true }));
  assert.equal(b.value(s, 0, 2), false);
  assert.ok(applyControl(b, s, dd.id, { sel: 3 }));
  assert.equal(b.value(s, 1, 2), 3);
  assert.ok(applyControl(b, s, o1.id, { toggle: true }));
  assert.equal(b.value(s, 2, 2), 1);
  assert.equal(controlState(b, s, s.controls[3]).checked, 0);
  assert.ok(applyControl(b, s, sp.id, { step: 1 }));
  assert.equal(b.value(s, 3, 2), 4);
  b.undo();
  assert.equal(b.value(s, 3, 2), 3);
  b.redo();
  const z = await JSZip.loadAsync(await writeXlsx(b, { ext: 'xlsx' }));
  const p = async (n) => z.file(`xl/ctrlProps/ctrlProp${n}.xml`).async('string');
  assert.ok(!/checked=/.test(await p(1)));
  assert.match(await p(2), /sel="3"/);
  assert.match(await p(3), /checked="Checked"/);
  assert.ok(!/checked=/.test(await p(4)));
  assert.match(await p(5), /val="4"/);
  const vml = await z.file(Object.keys(z.files).find((f) => /vmlDrawing/.test(f) && !/rels/.test(f))).async('string');
  assert.match(vml, /<x:Sel>3<\/x:Sel>/);
});

test('what-if data table: formula kept on save, results follow the inputs', async () => {
  const b = await readXlsx(SRC, { loc, ext: 'xlsx' });
  const s = b.sheets[0];
  assert.equal(s.dataTables.length, 1);
  assert.equal(b.value(s, 3, 5), 40);
  b.input(s, 3, 4, '25');
  recalcDataTables(b);
  assert.equal(b.value(s, 3, 5), 50);
  assert.equal(b.value(s, 2, 5), 20);
  assert.equal(b.value(s, 0, 4), 1); // input cell restored
  const out = await writeXlsx(b, { ext: 'xlsx' });
  const xml = await (await JSZip.loadAsync(out)).file('xl/worksheets/sheet1.xml').async('string');
  assert.match(xml, /<c r="F3"><f t="dataTable" ref="F3:F5" dt2D="0" dtr="0" r1="E1"\/><v>20<\/v><\/c>/);
  assert.match(xml, /<c r="F4"><v>50<\/v><\/c>/);
  void K;
});
