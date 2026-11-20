import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, StandardFonts, PDFName, PDFArray, PDFNumber, PDFHexString } from 'pdf-lib';
import { buildDocument, loadPdf } from '../ui/src/editors/pdf/edit.js';
import { readOwnAnnotations, hasMarker } from '../ui/src/editors/pdf/annotations.js';

async function base() {
  const d = await PDFDocument.create();
  const font = await d.embedFont(StandardFonts.Helvetica);
  d.addPage([400, 300]).drawText('Page', { x: 20, y: 250, size: 12, font });
  return d.save();
}
/** what another viewer does when it re-saves: drops our catalog marker, moves one annotation, edits a note, object streams */
async function foreignResave(bytes, { dx = 0, dy = 0, note = null, grow = 0 } = {}) {
  const d = await PDFDocument.load(bytes);
  const mk = d.catalog.get(PDFName.of('HWOfficeAnnots'));
  d.catalog.delete(PDFName.of('HWOfficeAnnots'));
  if (mk) d.context.delete(mk);
  const annots = d.getPages()[0].node.lookup(PDFName.of('Annots'), PDFArray);
  for (let i = 0; i < annots.size(); i++) {
    const a = annots.lookup(i);
    const r = a.lookup(PDFName.of('Rect'), PDFArray).asArray().map((x) => x.asNumber());
    if (a.get(PDFName.of('Subtype')).toString() === '/Text' && note != null) a.set(PDFName.of('Contents'), PDFHexString.fromText(note));
    a.set(PDFName.of('Rect'), d.context.obj([r[0] + dx, r[1] + dy, r[2] + dx + grow, r[3] + dy + grow].map((v) => PDFNumber.of(v))));
  }
  return d.save({ useObjectStreams: true });
}

const annots = { p0: [
  { id: 'r1', type: 'rect', x: 50, y: 50, w: 100, h: 40, color: '#ff0000', width: 2 },
  { id: 'n1', type: 'note', x: 300, y: 200, text: 'first', color: '#ffd400' },
] };

test('own annotations stay editable after another app re-saved the file (object streams, marker gone)', async () => {
  const ours = await buildDocument({ sources: [await base()], pages: [{ id: 'p0', src: 0, index: 0, rot: 0 }], annots, mode: 'annot' });
  const foreign = await foreignResave(ours, { dx: 10, dy: -5, note: 'changed elsewhere' });
  assert.equal(new TextDecoder('latin1').decode(foreign).includes('/HWOfficeAnnots'), false);
  assert.equal(new TextDecoder('latin1').decode(foreign).includes('/HWOffice'), false); // compressed into an object stream
  assert.ok(hasMarker(foreign));
  const own = readOwnAnnotations(await loadPdf(foreign));
  const r = own[0].find((a) => a.id === 'r1'); const n = own[0].find((a) => a.id === 'n1');
  assert.ok(r && n);
  assert.equal(r.x, 60); assert.equal(r.y, 45); assert.equal(r.w, 100);
  assert.equal(n.text, 'changed elsewhere'); assert.equal(n.x, 310);
});

test('a resized foreign annotation stays an ordinary annotation; files without our key are not loaded', async () => {
  const ours = await buildDocument({ sources: [await base()], pages: [{ id: 'p0', src: 0, index: 0, rot: 0 }], annots: { p0: [annots.p0[0]] }, mode: 'annot' });
  const own = readOwnAnnotations(await loadPdf(await foreignResave(ours, { grow: 40 })));
  assert.equal(own[0].length, 0);
  assert.equal(hasMarker(await base()), false);
  const plainObjStm = await (await PDFDocument.load(await base())).save({ useObjectStreams: true });
  assert.equal(hasMarker(plainObjStm), false);
});
