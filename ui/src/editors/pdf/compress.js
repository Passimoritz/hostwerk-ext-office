/*
 * Make a PDF smaller in the browser: big pictures are scaled down and stored as JPEG again, everything is saved with
 * compressed object streams. The picture work (decode → canvas → JPEG) is done by `recode`, which the editor
 * provides (browser APIs); the PDF part is pure and tested in tests/pdf-compress.test.js.
 *
 *   listImages(doc)                         → [{ ref, w, h, filter, cs, bpc, size, smask }]
 *   compressPdf(bytes, { recode, maxPx })   → { bytes, before, after, images }
 *        recode({ kind: 'jpeg', data } | { kind: 'raw', data, w, h, comps }, { maxPx }) → { data (jpeg), w, h, gray } | null
 */
import { PDFDocument, PDFName, PDFRawStream, PDFNumber, PDFArray, PDFRef, decodePDFRawStream } from 'pdf-lib';

const look = (ctx, o) => (o instanceof PDFRef ? ctx.lookup(o) : o);
const nameOf = (o) => (o instanceof PDFName ? o.decodeText() : o instanceof PDFArray && o.size() ? nameOf(o.get(0)) : '');

export function listImages(doc) {
  const ctx = doc.context; const out = [];
  for (const [ref, obj] of ctx.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    const d = obj.dict;
    if (d.get(PDFName.of('Subtype')) !== PDFName.of('Image')) continue;
    const num = (k) => { const v = look(ctx, d.get(PDFName.of(k))); return v instanceof PDFNumber ? v.asNumber() : 0; };
    const filter = nameOf(look(ctx, d.get(PDFName.of('Filter'))));
    const csObj = look(ctx, d.get(PDFName.of('ColorSpace')));
    out.push({
      ref, w: num('Width'), h: num('Height'), bpc: num('BitsPerComponent') || 8, filter, cs: nameOf(csObj), size: obj.contents.length,
      smask: !!d.get(PDFName.of('SMask')), mask: !!d.get(PDFName.of('Mask')), decode: !!d.get(PDFName.of('Decode')), imageMask: !!look(ctx, d.get(PDFName.of('ImageMask'))),
      filters: (() => { const f = look(ctx, d.get(PDFName.of('Filter'))); return f instanceof PDFArray ? f.size() : f ? 1 : 0; })(),
    });
  }
  return out;
}

/** which images we can safely re-encode */
export function canRecode(im) {
  if (im.imageMask || im.decode || im.mask || im.bpc !== 8 || im.filters > 1) return false;
  if (im.filter === 'DCTDecode') return im.cs === 'DeviceRGB' || im.cs === 'DeviceGray' || im.cs === 'ICCBased' || im.cs === '';
  if (im.filter === 'FlateDecode' || im.filter === '') return im.cs === 'DeviceRGB' || im.cs === 'DeviceGray';
  return false;
}

export async function compressPdf(bytes, { recode, maxPx = 2000, minSize = 40000 } = {}) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false });
  const ctx = doc.context;
  let images = 0;
  if (recode) {
    for (const im of listImages(doc)) {
      if (!canRecode(im) || im.size < minSize) continue;
      const stream = ctx.lookup(im.ref);
      let input;
      if (im.filter === 'DCTDecode') input = { kind: 'jpeg', data: stream.contents };
      else {
        let raw; try { raw = decodePDFRawStream(stream).decode(); } catch { continue; }
        const comps = im.cs === 'DeviceGray' ? 1 : 3;
        if (raw.length < im.w * im.h * comps) continue;
        const pred = look(ctx, stream.dict.get(PDFName.of('DecodeParms')));
        if (pred) continue; // PNG predictors – leave them
        input = { kind: 'raw', data: raw, w: im.w, h: im.h, comps };
      }
      let res = null;
      try { res = await recode(input, { maxPx }); } catch { res = null; }
      if (!res || !res.data || res.data.length >= im.size * 0.9) continue;
      const d = stream.dict;
      d.set(PDFName.of('Filter'), PDFName.of('DCTDecode'));
      d.delete(PDFName.of('DecodeParms'));
      d.set(PDFName.of('Width'), PDFNumber.of(res.w)); d.set(PDFName.of('Height'), PDFNumber.of(res.h));
      d.set(PDFName.of('BitsPerComponent'), PDFNumber.of(8));
      d.set(PDFName.of('ColorSpace'), PDFName.of(res.gray ? 'DeviceGray' : 'DeviceRGB'));
      ctx.assign(im.ref, PDFRawStream.of(d, res.data));
      images++;
    }
  }
  const out = await doc.save({ useObjectStreams: true, updateFieldAppearances: false });
  return { bytes: out.length < bytes.length ? out : bytes, before: bytes.length, after: Math.min(out.length, bytes.length), images };
}
