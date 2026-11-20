import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, rgb } from 'pdf-lib';
import { diffWords, comparePages, pixelDiff } from '../ui/src/editors/pdf/compare.js';
import { compressPdf, listImages, canRecode } from '../ui/src/editors/pdf/compress.js';

test('word diff', () => {
  const d = diffWords('Der Preis beträgt 100 Euro netto.', 'Der Preis beträgt 120 Euro brutto.');
  assert.deepEqual(d.filter((x) => x.op < 0).map((x) => x.text.trim()), ['100', 'netto.']);
  assert.deepEqual(d.filter((x) => x.op > 0).map((x) => x.text.trim()), ['120', 'brutto.']);
  assert.equal(d.filter((x) => x.op === 0).map((x) => x.text).join(''), 'Der Preis beträgt  Euro ');
});

test('pages: an inserted page is found, the others are matched', () => {
  const a = ['Einleitung und Ziele des Projekts', 'Kosten und Zeitplan für das Jahr', 'Anhang mit Kontakten'];
  const b = ['Einleitung und Ziele des Projekts', 'Ganz neue Seite mit Risiken und Annahmen', 'Kosten und Zeitplan für das Jahr 2027', 'Anhang mit Kontakten'];
  const r = comparePages(a, b);
  assert.deepEqual(r.map((x) => [x.a, x.b]), [[0, 0], [null, 1], [1, 2], [2, 3]]);
  assert.ok(r[0].same && r[3].same && !r[2].same);
});

test('pixel diff', () => {
  const A = new Uint8ClampedArray(4 * 4 * 4).fill(255); const B = A.slice();
  B[0] = 0; B[1] = 0; B[2] = 0;
  const d = pixelDiff(A, B, 4, 4);
  assert.equal(d.count, 1); assert.equal(d.mask[0], 1);
});

test('compress: big raw image is handed to the recoder and replaced', async () => {
  const d = await PDFDocument.create();
  const p = d.addPage([400, 400]);
  p.drawRectangle({ x: 0, y: 0, width: 10, height: 10, color: rgb(1, 0, 0) });
  // an uncompressed-ish 300×300 RGB image (noise → big)
  const w = 300; const h = 300; const px = new Uint8Array(w * h * 3).map(() => Math.floor(Math.random() * 256));
  const ref = d.context.register(d.context.flateStream(px, { Type: 'XObject', Subtype: 'Image', Width: w, Height: h, ColorSpace: 'DeviceRGB', BitsPerComponent: 8 }));
  p.node.newXObject('Im1', ref);
  const bytes = await d.save();
  const imgs = listImages(await PDFDocument.load(bytes));
  assert.equal(imgs.length, 1); assert.ok(canRecode(imgs[0]));
  let seen = null;
  const fakeJpeg = new Uint8Array(1000).fill(7);
  const r = await compressPdf(bytes, { recode: async (input) => { seen = input; return { data: fakeJpeg, w: 150, h: 150, gray: false }; } });
  assert.equal(seen.kind, 'raw'); assert.equal(seen.w, 300); assert.equal(seen.comps, 3);
  assert.equal(r.images, 1);
  assert.ok(r.after < r.before / 10);
  const im2 = listImages(await PDFDocument.load(r.bytes))[0];
  assert.equal(im2.filter, 'DCTDecode'); assert.equal(im2.w, 150);
});
