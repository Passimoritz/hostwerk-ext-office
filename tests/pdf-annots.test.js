import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { PDFDocument, StandardFonts, PDFName } from 'pdf-lib';
import { buildDocument, loadPdf, missingGlyphs, listFields } from '../ui/src/editors/pdf/edit.js';
import { readOwnAnnotations, hasMarker } from '../ui/src/editors/pdf/annotations.js';
import { parseFont, subsetFont, canEncodeWinAnsi, woffToSfnt } from '../ui/src/editors/pdf/ttf.js';

const LATIN = new URL('../ui/src/editors/pdf/fonts/latin-sans.woff', import.meta.url);
const LIB = new URL('../node_modules/pdfjs-dist/standard_fonts/LiberationSans-Regular.ttf', import.meta.url);

async function pdfjsPages(bytes) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false, verbosity: 0 }).promise;
  const out = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const p = await doc.getPage(i);
    const ops = await p.getOperatorList({ annotationMode: 2 }).catch(() => null);
    out.push({ text: (await p.getTextContent()).items.map((x) => x.str).join(''), annots: await p.getAnnotations(), ops: ops ? ops.fnArray.length : 0 });
  }
  const outline = await doc.getOutline();
  const fields = await doc.getFieldObjects();
  await doc.destroy();
  return { pages: out, outline, fields };
}

test('WinAnsi detection', () => {
  assert.ok(canEncodeWinAnsi('Grüße – 5 € „gut“'));
  assert.ok(!canEncodeWinAnsi('Łódź'));
  assert.ok(!canEncodeWinAnsi('日本'));
});

test('bundled Latin font: WOFF → TrueType with Polish, Czech, Greek, Cyrillic glyphs', async () => {
  const f = parseFont(await woffToSfnt(new Uint8Array(fs.readFileSync(LATIN))));
  for (const ch of 'ŁłČčŠšŽžŐőΩωЖж€') assert.ok(f.gid(ch.codePointAt(0)) > 0, ch);
  assert.ok(f.advance(f.gid(65)) > 0);
});

test('subset keeps glyph ids and is much smaller', () => {
  const f = parseFont(new Uint8Array(fs.readFileSync(LIB)));
  const g = [f.gid(65), f.gid(66)];
  const sub = subsetFont(f, g);
  assert.ok(sub.length < fs.statSync(LIB).size / 2);
  const s = parseFont(subsetFont(f, g, { keepCmap: true }));
  assert.equal(s.gid(65), g[0]);
  assert.equal(s.numGlyphs, f.numGlyphs);
});

async function base(n = 1) {
  const d = await PDFDocument.create();
  const font = await d.embedFont(StandardFonts.Helvetica);
  for (let i = 0; i < n; i++) d.addPage([400, 300]).drawText(`Seite ${i + 1}`, { x: 20, y: 250, size: 12, font });
  return d.save();
}

test('Unicode annotation text is embedded with a real font (copyable, not "?")', async () => {
  const latin = await woffToSfnt(new Uint8Array(fs.readFileSync(LATIN)));
  const annots = { p0: [{ id: 'a1', type: 'text', x: 30, y: 150, text: 'Łódź – Žluťoučký Ωmega Жук', size: 14, color: '#d0021b', rot: 0 }] };
  assert.deepEqual(missingGlyphs(annots, [latin]), []);
  assert.deepEqual(missingGlyphs({ p0: [{ type: 'text', text: '漢字' }] }, [latin]), ['漢', '字']);
  for (const mode of ['burn', 'annot']) {
    const out = await buildDocument({ sources: [await base()], pages: [{ id: 'p0', src: 0, index: 0, rot: 0 }], annots, mode, fonts: [latin] });
    const r = await pdfjsPages(out);
    if (mode === 'burn') assert.match(r.pages[0].text, /Łódź – Žluťoučký Ωmega Жук/);
    else {
      const a = r.pages[0].annots.find((x) => x.subtype === 'FreeText');
      assert.ok(a, 'FreeText annotation');
      assert.match(a.contentsObj.str, /Łódź/);
    }
    assert.ok(!Buffer.from(out).includes('LiberationSans-Regular.ttf'));
  }
});

test('own annotations stay editable after saving and opening again', async () => {
  const annots = {
    p0: [
      { id: 'h1', type: 'highlight', quads: [{ x: 18, y: 246, w: 50, h: 14 }], color: '#ffd400', text: 'Seite 1' },
      { id: 'r1', type: 'rect', x: 100, y: 50, w: 80, h: 40, color: '#e5364a', width: 2 },
      { id: 't1', type: 'text', x: 30, y: 120, text: 'Hallo', size: 12, color: '#111111', rot: 0 },
    ],
  };
  const out = await buildDocument({ sources: [await base()], pages: [{ id: 'p0', src: 0, index: 0, rot: 0 }], annots, mode: 'annot' });
  assert.ok(hasMarker(out), 'marker is visible without parsing');
  assert.ok(!hasMarker(await base()));
  const doc = await loadPdf(out);
  const back = readOwnAnnotations(doc);
  assert.equal(back[0].length, 3);
  assert.deepEqual(back[0].map((a) => a.type).sort(), ['highlight', 'rect', 'text']);
  assert.equal(back[0].find((a) => a.type === 'rect').w, 80);
  assert.equal(doc.getPage(0).node.lookup(PDFName.of('Annots')), undefined, 'taken out of the page');
  // other apps' annotations stay
  const other = await PDFDocument.load(out);
  const ctx = other.context;
  other.getPage(0).node.addAnnot(ctx.register(ctx.obj({ Type: 'Annot', Subtype: 'Square', Rect: [1, 1, 5, 5] })));
  const d2 = await loadPdf(await other.save());
  assert.equal(readOwnAnnotations(d2)[0].length, 3);
  assert.equal(d2.getPage(0).node.lookup(PDFName.of('Annots')).size(), 1);
});

test('inserted pages keep their form fields and bookmarks', async () => {
  // second PDF: two pages, a text field, an outline with two entries
  const d = await PDFDocument.create();
  const p1 = d.addPage([400, 300]); const p2 = d.addPage([400, 300]);
  const form = d.getForm();
  const tf = form.createTextField('Name'); tf.setText('Erika'); tf.addToPage(p2, { x: 20, y: 20, width: 200, height: 24 });
  const ctx = d.context;
  const outlineRef = ctx.nextRef(); const i1 = ctx.nextRef(); const i2 = ctx.nextRef();
  ctx.assign(i1, ctx.obj({ Title: 'Kapitel A'.split('').length ? ctx.obj('Kapitel A') : '', Parent: outlineRef, Next: i2, Dest: [p1.ref, 'XYZ', null, null, null] }));
  ctx.assign(i2, ctx.obj({ Parent: outlineRef, Prev: i1, Dest: [p2.ref, 'Fit'] }));
  const { PDFHexString } = await import('pdf-lib');
  ctx.lookup(i1).set(PDFName.of('Title'), PDFHexString.fromText('Kapitel A'));
  ctx.lookup(i2).set(PDFName.of('Title'), PDFHexString.fromText('Kapitel B'));
  ctx.assign(outlineRef, ctx.obj({ Type: 'Outlines', First: i1, Last: i2, Count: 2 }));
  d.catalog.set(PDFName.of('Outlines'), outlineRef);
  const other = await d.save();
  // base also has a field called "Name"
  const b = await PDFDocument.create(); const bp = b.addPage([400, 300]);
  const bf = b.getForm().createTextField('Name'); bf.addToPage(bp, { x: 20, y: 20, width: 100, height: 20 });
  const baseBytes = await b.save();
  const out = await buildDocument({
    sources: [baseBytes, other], names: ['', 'Anhang.pdf'],
    pages: [{ id: 'p0', src: 0, index: 0, rot: 0 }, { id: 's1p0', src: 1, index: 0, rot: 0 }, { id: 's1p1', src: 1, index: 1, rot: 0 }],
  });
  const fields = listFields(await loadPdf(out));
  assert.deepEqual(fields.map((f) => f.name).sort(), ['Name', 'Name (2)']);
  assert.equal(fields.find((f) => f.name === 'Name (2)').value, 'Erika');
  const r = await pdfjsPages(out);
  assert.equal(r.outline.length, 1);
  assert.equal(r.outline[0].title, 'Anhang.pdf');
  assert.deepEqual(r.outline[0].items.map((x) => x.title), ['Kapitel A', 'Kapitel B']);
  assert.ok(r.fields && r.fields['Name (2)']);
});
