import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import { findSensitive, hitRects, keptWords, applyRedactions, verifyRedaction, scanFile, collectGarbage } from '../ui/src/editors/pdf/redact.js';
import { pageIndex, findInPage } from '../ui/src/editors/pdf/search.js';

// 1×1 white PNG – the "rendered page" (the editor renders it with pdf.js and paints the marks black)
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4//8/AAX+Av4N70a4AAAAAElFTkSuQmCC', 'base64'));

test('patterns: IBAN (checksum), e-mail, phone, card (Luhn), names, own terms', () => {
  const text = 'Bitte an DE89 3704 0044 0532 0130 00 überweisen, nicht DE00 1234 5678 9012 3456 78. Mail: erika.muster@example.org, Tel. +49 30 1234567 oder 030/123 45 67. Karte 4111 1111 1111 1111. Sehr geehrte Frau Mustermann, Projekt Falke.';
  const hits = findSensitive(text, { kinds: ['iban', 'email', 'phone', 'card', 'name'], terms: ['projekt falke'] });
  const v = hits.map((h) => `${h.kind}:${h.value}`);
  assert.ok(v.includes('iban:DE89 3704 0044 0532 0130 00'), v.join(' | '));
  assert.ok(!v.some((x) => x.includes('DE00')), 'invalid IBAN ignored');
  assert.ok(v.includes('email:erika.muster@example.org'));
  assert.ok(v.some((x) => x.startsWith('phone:+49 30 1234567')));
  assert.ok(v.some((x) => x.startsWith('phone:030/123 45 67')));
  assert.ok(v.includes('card:4111 1111 1111 1111'));
  assert.ok(v.includes('name:Mustermann'));
  assert.ok(v.includes('term:Projekt Falke'));
});

async function sample() {
  const d = await PDFDocument.create();
  const f = await d.embedFont(StandardFonts.Helvetica);
  const p = d.addPage([400, 300]);
  p.drawText('Kunde: Erika Mustermann', { x: 20, y: 250, size: 14, font: f });
  p.drawText('IBAN DE89 3704 0044 0532 0130 00', { x: 20, y: 220, size: 14, font: f });
  p.drawText('Betrag 120,00 EUR', { x: 20, y: 190, size: 14, font: f });
  p.drawRectangle({ x: 20, y: 100, width: 60, height: 40, color: rgb(0, 0, 1) });
  d.addPage([400, 300]).drawText('Seite zwei bleibt', { x: 20, y: 250, size: 14, font: f });
  return d.save();
}
async function extractAll(bytes) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false, verbosity: 0 }).promise;
  const out = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const p = await doc.getPage(i);
    const items = (await p.getTextContent()).items;
    out.push({ items, text: items.map((x) => x.str).join(' '), view: p.view });
  }
  await doc.destroy();
  return out;
}

test('redaction removes the text under the marks for real and keeps the rest searchable', async () => {
  const src = await sample();
  const pages = await extractAll(src);
  const p0 = pages[0];
  // marks: the IBAN found by the pattern search + the name, mapped to rectangles
  const idx = pageIndex(p0.items);
  const secrets = ['DE89 3704 0044 0532 0130 00', 'Erika Mustermann'];
  const rects = []; const removed = new Map();
  for (const s of secrets) {
    for (const h of findInPage(idx, s, { matchCase: true })) {
      for (const part of h.parts) {
        rects.push(hitRects(p0.items[part.i], part.a, part.b));
        if (!removed.has(part.i)) removed.set(part.i, []);
        removed.get(part.i).push([part.a, part.b]);
      }
    }
  }
  assert.equal(rects.length, 2);
  const words = keptWords(p0.items, rects, removed);
  assert.ok(words.some((w) => w.text === 'Betrag'));
  assert.ok(!words.some((w) => /Mustermann|3704/.test(w.text)));
  const out = await applyRedactions(src, [{ index: 0, rects, image: { bytes: PNG, png: true }, words, view: p0.view }]);
  const after = await extractAll(out);
  assert.match(after[0].text, /Betrag/);
  assert.match(after[0].text, /Kunde:/);
  assert.match(after[1].text, /Seite zwei bleibt/);
  const v = verifyRedaction(after, { secrets, marks: { 0: rects } });
  assert.ok(v.ok, JSON.stringify(v.leaks));
  assert.deepEqual(await scanFile(out, secrets), []);
  // the check itself catches a leak in the original
  const bad = verifyRedaction(pages, { secrets, marks: { 0: rects } });
  assert.ok(!bad.ok && bad.leaks.length >= 2);
  assert.ok((await scanFile(src, secrets)).length >= 1);
});

test('garbage collection drops objects nobody refers to', async () => {
  const d = await PDFDocument.create();
  d.addPage();
  d.context.register(d.context.obj({ Secret: 'x' }));
  const before = d.context.enumerateIndirectObjects().length;
  assert.equal(collectGarbage(d), 1);
  assert.equal(d.context.enumerateIndirectObjects().length, before - 1);
});
