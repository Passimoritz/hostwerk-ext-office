import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, StandardFonts, rgb, PDFName, PDFDict } from 'pdf-lib';
import { applyRedactions, verifyRedaction, scanFile, hitRects } from '../ui/src/editors/pdf/redact.js';
import { pageIndex, findInPage } from '../ui/src/editors/pdf/search.js';

const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4//8/AAX+Av4N70a4AAAAAElFTkSuQmCC', 'base64'));
async function extract(bytes) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false, verbosity: 0 }).promise;
  const out = [];
  for (let i = 1; i <= doc.numPages; i++) { const p = await doc.getPage(i); const items = (await p.getTextContent()).items; const ops = await p.getOperatorList(); out.push({ items, text: items.map((x) => x.str).join(' '), view: p.view, ops }); }
  await doc.destroy();
  return out;
}
async function src(withImage = false) {
  const d = await PDFDocument.create();
  const f = await d.embedFont(StandardFonts.Helvetica);
  const p = d.addPage([400, 300]);
  p.drawText('Kunde: Erika Mustermann wohnt hier', { x: 20, y: 250, size: 14, font: f });
  p.drawText('Betrag 120,00 EUR', { x: 20, y: 190, size: 14, font: f });
  p.drawRectangle({ x: 20, y: 100, width: 60, height: 40, color: rgb(0, 0, 1) });
  p.drawRectangle({ x: 300, y: 100, width: 20, height: 20, color: rgb(1, 0, 0) }); // inside a mark → not painted
  if (withImage) p.drawImage(await d.embedPng(PNG), { x: 200, y: 20, width: 10, height: 10 });
  return d.save();
}
async function marks(bytes, secret) {
  const [p0] = await extract(bytes);
  const idx = pageIndex(p0.items); const rects = [];
  for (const h of findInPage(idx, secret, { matchCase: true })) for (const part of h.parts) rects.push(hitRects(p0.items[part.i], part.a, part.b));
  rects.push({ x: 295, y: 95, w: 30, h: 30 });
  return { rects, view: p0.view };
}

test('vector redaction: marked glyphs and paths gone, the rest stays text and vectors', async () => {
  const s = await src();
  const { rects, view } = await marks(s, 'Erika Mustermann');
  const jobs = [{ index: 0, rects, image: { bytes: PNG, png: true }, words: [], view }];
  const out = await applyRedactions(s, jobs);
  assert.equal(jobs[0].done, 'vector');
  const [p] = await extract(out);
  assert.match(p.text, /Kunde:/); assert.match(p.text, /wohnt hier/); assert.match(p.text, /Betrag 120,00 EUR/);
  assert.ok(!/Erika|Mustermann/.test(p.text), p.text);
  assert.ok(verifyRedaction([p], { secrets: ['Erika Mustermann'], marks: { 0: rects } }).ok);
  assert.deepEqual(await scanFile(out, ['Erika Mustermann']), []);
  // no picture on the page: still vector content, the page has no XObjects
  const d = await PDFDocument.load(out);
  const res = d.getPages()[0].node.Resources();
  assert.ok(!(res.lookup(PDFName.of('XObject')) instanceof PDFDict) || res.lookup(PDFName.of('XObject')).keys().length === 0);
  // "wohnt" keeps its place (TJ offsets instead of the removed glyphs): 20 + width of "Kunde: Erika Mustermann " ≈ 185.2
  const w = p.items.find((x) => x.str.startsWith('wohnt'));
  assert.ok(w && Math.abs(w.transform[4] - 185.2) < 1.5, JSON.stringify(p.items.map((x) => [x.str, x.transform[4]])));
});

test('pages with pictures still take the picture method', async () => {
  const s = await src(true);
  const { rects, view } = await marks(s, 'Erika Mustermann');
  const jobs = [{ index: 0, rects, image: { bytes: PNG, png: true }, words: [], view }];
  await applyRedactions(s, jobs);
  assert.equal(jobs[0].done, undefined);
});
