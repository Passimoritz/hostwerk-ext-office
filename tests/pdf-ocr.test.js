import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument } from 'pdf-lib';
import { parseTsv, evenLines } from '../ui/src/editors/pdf/ocr.js';
import { addTextLayer, imageWordsToPdf, itemsToWords } from '../ui/src/editors/pdf/textlayer.js';

const TSV = [
  'level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext',
  '1\t1\t0\t0\t0\t0\t0\t0\t1000\t1414\t-1\t',
  '4\t1\t1\t1\t1\t0\t100\t100\t500\t40\t-1\t',
  '5\t1\t1\t1\t1\t1\t100\t100\t180\t40\t96.5\tRechnung',
  '5\t1\t1\t1\t1\t2\t300\t104\t120\t34\t91\tNr.',
  '5\t1\t1\t1\t1\t3\t440\t102\t160\t38\t12\t~~',
  '5\t1\t1\t1\t2\t1\t100\t200\t260\t40\t88\tMüller',
].join('\n');

test('TSV → words (low confidence dropped, lines evened)', () => {
  const w = evenLines(parseTsv(TSV));
  assert.deepEqual(w.map((x) => x.text), ['Rechnung', 'Nr.', 'Müller']);
  assert.equal(w[0].y0, 100); assert.equal(w[1].y0, 100); assert.equal(w[1].y1, 140);
  assert.equal(w[2].x1, 360);
});

async function extract(bytes) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false, verbosity: 0 }).promise;
  const page = await doc.getPage(1);
  const tc = await page.getTextContent();
  await doc.destroy();
  return tc.items;
}

test('invisible text layer: pdf.js finds the words at their places (also on a rotated page)', async () => {
  for (const rotate of [0, 90]) {
    const doc = await PDFDocument.create();
    const page = doc.addPage([595, 842]);
    // image of 1000×1414 px over the page; toPdf maps pixels (top-left) → points (bottom-left)
    const sx = 595 / 1000; const sy = 842 / 1414;
    const toPdf = rotate ? (x, y) => [x * sx, 842 - y * sy] : (x, y) => [x * sx, 842 - y * sy];
    const words = imageWordsToPdf(parseTsv(TSV), toPdf);
    await addTextLayer(doc, page, words);
    if (rotate) page.setRotation({ type: 'degrees', angle: 90 });
    const items = await extract(await doc.save());
    const text = items.map((i) => i.str).join('');
    assert.match(text, /Rechnung/);
    assert.match(text, /Müller/);
    const it = items.find((i) => i.str.startsWith('Rechnung'));
    assert.ok(Math.abs(it.transform[4] - 100 * sx) < 2, `x of the first word (${it.transform[4]})`);
    // the width of the text matches the box (± 15 %)
    assert.ok(Math.abs(it.width - 180 * sx * (9 / 8.4)) / (180 * sx) < 0.2, `width ${it.width}`);
  }
});

test('itemsToWords → text layer keeps a page\'s text', async () => {
  const items = [{ str: 'Hallo Welt', transform: [12, 0, 0, 12, 50, 700], width: 60, height: 12 }];
  const doc = await PDFDocument.create();
  const page = doc.addPage([595, 842]);
  await addTextLayer(doc, page, itemsToWords(items));
  const out = await extract(await doc.save());
  assert.equal(out.map((x) => x.str).join('').trim(), 'Hallo Welt');
  assert.ok(Math.abs(out[0].transform[5] - 700) < 3);
});
