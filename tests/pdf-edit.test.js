import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import {
  createBlank, mergePdfs, reorderPages, rotatePages, deletePages, insertBlankPage, insertPdf, extractPages, splitPdf,
  getMetadata, setMetadata, buildDocument, imagesToPdf, listFields, listSignatures, inspect, loadPdf,
} from '../ui/src/editors/pdf/edit.js';
import { bbox, translate, scale, mergeQuads, hexToRgb, turn } from '../ui/src/editors/pdf/annotations.js';

/** a PDF whose pages have distinct widths (100, 110, 120 …) so the order can be checked */
async function make(n, start = 100) {
  const d = await PDFDocument.create();
  const f = await d.embedFont(StandardFonts.Helvetica);
  for (let i = 0; i < n; i++) d.addPage([start + i * 10, 200]).drawText(`Page ${i + 1}`, { x: 10, y: 100, size: 12, font: f });
  return d.save();
}
const widths = async (b) => (await PDFDocument.load(b)).getPages().map((p) => Math.round(p.getWidth()));
const rotations = async (b) => (await PDFDocument.load(b)).getPages().map((p) => p.getRotation().angle);
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR4nGP8z8DwnwEIGBkZGRkZGQAhFwQBWGSfzwAAAABJRU5ErkJggg==', 'base64'));

test('blank A4', async () => {
  const d = await PDFDocument.load(await createBlank());
  assert.equal(d.getPageCount(), 1);
  assert.deepEqual([Math.round(d.getPage(0).getWidth()), Math.round(d.getPage(0).getHeight())], [595, 842]);
});
test('merge keeps order', async () => {
  assert.deepEqual(await widths(await mergePdfs([await make(2, 100), await make(3, 300)])), [100, 110, 300, 310, 320]);
});
test('reorder and delete', async () => {
  const b = await make(4);
  assert.deepEqual(await widths(await reorderPages(b, [3, 0, 2, 1])), [130, 100, 120, 110]);
  assert.deepEqual(await widths(await deletePages(b, [1, 2])), [100, 130]);
});
test('rotate adds to the existing rotation', async () => {
  const b = await rotatePages(await make(3), [0, 2], 90);
  assert.deepEqual(await rotations(b), [90, 0, 90]);
  assert.deepEqual(await rotations(await rotatePages(b, [0], 270)), [0, 0, 90]);
});
test('insert blank page and other PDF, extract, split', async () => {
  const b = await make(2);
  assert.deepEqual(await widths(await insertBlankPage(b, 1)), [100, 100, 110]);
  assert.deepEqual(await widths(await insertPdf(b, await make(1, 500), 1)), [100, 500, 110]);
  assert.deepEqual(await widths(await extractPages(await make(5), [4, 1])), [140, 110]);
  const parts = await splitPdf(await make(5), 2);
  assert.deepEqual(await Promise.all(parts.map(widths)), [[100, 110], [120, 130], [140]]);
});
test('metadata round trip', async () => {
  const b = await setMetadata(await make(1), { title: 'Quartalsbericht – Ü', author: 'Lena', subject: 'Zahlen', keywords: 'a, b' });
  const m = await getMetadata(b);
  assert.equal(m.title, 'Quartalsbericht – Ü');
  assert.equal(m.author, 'Lena');
  assert.equal(m.subject, 'Zahlen');
  assert.match(m.keywords, /a/);
});
test('buildDocument: pages from two sources, rotation, annotations, metadata', async () => {
  const a = await make(2, 100); const c = await make(2, 400);
  const annots = {
    x: [{ type: 'highlight', quads: [{ x: 10, y: 95, w: 40, h: 14 }], color: '#ffd400' }, { type: 'text', x: 20, y: 50, text: 'Hallo €', size: 12, color: '#ff0000', rot: 0 }, { type: 'note', x: 5, y: 5, text: 'Bitte prüfen' }],
    y: [{ type: 'ink', paths: [[1, 1, 20, 20, 30, 5]], color: '#000000', width: 2 }, { type: 'image', x: 10, y: 10, w: 20, h: 20, rot: 90, png: PNG }, { type: 'stamp', x: 10, y: 120, w: 80, h: 24, text: 'APPROVED', color: '#1d8a4b', rot: 0 }, { type: 'arrow', x1: 0, y1: 0, x2: 50, y2: 50, width: 2, color: '#000000' }],
  };
  const out = await buildDocument({ sources: [a, c], pages: [{ id: 'y', src: 1, index: 1, rot: 90 }, { id: 'x', src: 0, index: 0, rot: 0 }], annots, meta: { title: 'T' } });
  const d = await PDFDocument.load(out);
  assert.deepEqual(d.getPages().map((p) => Math.round(p.getWidth())), [410, 100]);
  assert.deepEqual(d.getPages().map((p) => p.getRotation().angle), [90, 0]);
  assert.equal(d.getTitle(), 'T');
  // page 2 ('x') has 3 real annotations (highlight, free text, note)
  assert.equal(d.getPage(1).node.Annots().size(), 3);
  // burn mode: no annotation objects
  const burned = await PDFDocument.load(await buildDocument({ sources: [a], pages: [{ id: 'x', src: 0, index: 0, rot: 0 }], annots, mode: 'burn' }));
  assert.equal((burned.getPage(0).node.Annots()?.size() || 0), 0);
});
test('forms: fill, keep and flatten', async () => {
  const d = await PDFDocument.create();
  const p = d.addPage([300, 300]);
  const form = d.getForm();
  const tf = form.createTextField('name'); tf.addToPage(p, { x: 10, y: 250, width: 200, height: 20 });
  const cb = form.createCheckBox('ok'); cb.addToPage(p, { x: 10, y: 220, width: 15, height: 15 });
  const rg = form.createRadioGroup('size'); rg.addOptionToPage('S', p, { x: 10, y: 190 }); rg.addOptionToPage('L', p, { x: 40, y: 190 });
  const dd = form.createDropdown('city'); dd.addOptions(['Berlin', 'Köln']); dd.addToPage(p, { x: 10, y: 150, width: 100, height: 20 });
  const b = await d.save();
  const out = await buildDocument({ sources: [b], pages: [{ id: 'a', src: 0, index: 0, rot: 0 }], forms: { name: 'Lena', ok: true, size: 'L', city: 'Berlin' } });
  const f = Object.fromEntries(listFields(await loadPdf(out)).map((x) => [x.name, x.value]));
  assert.deepEqual(f, { name: 'Lena', ok: true, size: 'L', city: 'Berlin' });
  const flat = await buildDocument({ sources: [b], pages: [{ id: 'a', src: 0, index: 0, rot: 0 }], forms: { name: 'Lena' }, flatten: true });
  assert.equal(listFields(await loadPdf(flat)).length, 0);
  assert.deepEqual(listSignatures(await loadPdf(b)), []);
  assert.equal(inspect(await loadPdf(b)).xfa, false);
});
test('images to PDF', async () => {
  const d = await PDFDocument.load(await imagesToPdf([{ bytes: PNG }, { bytes: PNG }]));
  assert.equal(d.getPageCount(), 2);
});
test('annotation geometry', () => {
  assert.deepEqual(hexToRgb('#ff0000'), [1, 0, 0]);
  const [x, y] = turn(10, 0, 90); assert.ok(Math.abs(x) < 1e-9 && Math.abs(y - 10) < 1e-9);
  const r = { type: 'rect', x: 10, y: 10, w: 20, h: 10, width: 2 };
  assert.deepEqual(bbox(r), { x: 9, y: 9, w: 22, h: 12 });
  assert.deepEqual(translate(r, 5, -5), { ...r, x: 15, y: 5 });
  const s = scale(r, 2); assert.equal(s.w, 40); assert.equal(s.x, 0);
  const q = mergeQuads([{ x: 0, y: 0, w: 10, h: 10 }, { x: 11, y: 1, w: 10, h: 9 }, { x: 0, y: -20, w: 10, h: 10 }]);
  assert.equal(q.length, 2); assert.equal(q[0].w, 21);
});
