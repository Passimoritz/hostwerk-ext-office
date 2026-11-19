/*
 * PDF operations with pdf-lib – pure functions on bytes (no DOM), unit-tested in tests/pdf-edit.test.js.
 *
 * The editor keeps a page model { sources: [bytes], pages: [{ id, src, index, rot }], annots: { pageId: [...] },
 * forms: { fieldName: value }, meta } and writes it with buildDocument(): the first source is the base document (its
 * AcroForm, outline, XMP/PDF-A metadata, attachments stay), pages of other sources are copied in.
 */
import {
  PDFDocument, PDFName, PDFDict, PDFString, PDFHexString, StandardFonts, PageSizes, degrees,
  PDFTextField, PDFCheckBox, PDFRadioGroup, PDFDropdown, PDFOptionList, PDFSignature,
} from 'pdf-lib';
import { writeAnnotations } from './annotations.js';

export const SIZES = { A4: PageSizes.A4, Letter: PageSizes.Letter, A5: PageSizes.A5, Legal: PageSizes.Legal };
const LOAD = { updateMetadata: false, ignoreEncryption: false, throwOnInvalidObject: false };
const u8 = (b) => (b instanceof Uint8Array ? b : new Uint8Array(b));
const norm = (d) => (((d % 360) + 360) % 360);

export async function loadPdf(bytes) { return PDFDocument.load(u8(bytes), LOAD); }

/** a new PDF with `pages` empty pages (default one A4 portrait page) */
export async function createBlank({ size = 'A4', pages = 1, landscape = false } = {}) {
  const doc = await PDFDocument.create();
  let [w, h] = Array.isArray(size) ? size : SIZES[size] || SIZES.A4;
  if (landscape) [w, h] = [h, w];
  for (let i = 0; i < pages; i++) doc.addPage([w, h]);
  doc.setProducer('Hostwerk Office');
  doc.setCreator('Hostwerk Office');
  return doc.save();
}

export async function pageCount(bytes) { return (await loadPdf(bytes)).getPageCount(); }

/** join several PDFs in the given order */
export async function mergePdfs(list) {
  const out = await PDFDocument.create();
  for (const b of list) {
    const src = await loadPdf(b);
    const pages = await out.copyPages(src, src.getPageIndices());
    pages.forEach((p) => out.addPage(p));
  }
  return out.save();
}

/** keep the pages in `order` (original indices; leaving some out deletes them) */
export async function reorderPages(bytes, order) {
  return buildDocument({ sources: [bytes], pages: order.map((index, i) => ({ id: `p${i}`, src: 0, index, rot: 0 })) });
}

export async function rotatePages(bytes, indices, delta) {
  const doc = await loadPdf(bytes);
  for (const i of indices) { const p = doc.getPage(i); p.setRotation(degrees(norm(p.getRotation().angle + delta))); }
  return doc.save();
}

export async function deletePages(bytes, indices) {
  const drop = new Set(indices);
  const n = await pageCount(bytes);
  return reorderPages(bytes, [...Array(n).keys()].filter((i) => !drop.has(i)));
}

export async function insertBlankPage(bytes, at, size) {
  const doc = await loadPdf(bytes);
  const ref = doc.getPageCount() ? doc.getPage(Math.max(0, Math.min(at, doc.getPageCount()) - 1)).getSize() : null;
  const dims = size ? (Array.isArray(size) ? size : SIZES[size]) : ref ? [ref.width, ref.height] : SIZES.A4;
  doc.insertPage(Math.max(0, Math.min(at, doc.getPageCount())), dims);
  return doc.save();
}

export async function insertPdf(bytes, other, at) {
  const doc = await loadPdf(bytes);
  const src = await loadPdf(other);
  const pages = await doc.copyPages(src, src.getPageIndices());
  let pos = Math.max(0, Math.min(at == null ? doc.getPageCount() : at, doc.getPageCount()));
  for (const p of pages) doc.insertPage(pos++, p);
  return doc.save();
}

/** new PDF with only these pages (in this order) */
export async function extractPages(bytes, indices) {
  const src = await loadPdf(bytes);
  const out = await PDFDocument.create();
  const pages = await out.copyPages(src, indices);
  pages.forEach((p) => out.addPage(p));
  return out.save();
}

/** split into parts of `size` pages each → [Uint8Array] */
export async function splitPdf(bytes, size = 1) {
  const n = await pageCount(bytes);
  const parts = [];
  for (let i = 0; i < n; i += size) parts.push(await extractPages(bytes, [...Array(Math.min(size, n - i)).keys()].map((k) => i + k)));
  return parts;
}

const isPng = (b) => b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47;
const isJpg = (b) => b[0] === 0xff && b[1] === 0xd8;

/**
 * images → PDF. images: [{ bytes (png or jpg) }]. fit 'image' = page as large as the image (72 dpi → 1 px = 0.75 pt,
 * capped to A4 scale), 'a4' = every image centred on an A4 page with a margin.
 */
export async function imagesToPdf(images, { fit = 'a4', margin = 24 } = {}) {
  const doc = await PDFDocument.create();
  for (const im of images) {
    const b = u8(im.bytes);
    const img = isPng(b) ? await doc.embedPng(b) : isJpg(b) ? await doc.embedJpg(b) : null;
    if (!img) throw new Error('Only PNG and JPEG images can be put into a PDF directly.');
    if (fit === 'image') {
      const w = img.width * 0.75; const h = img.height * 0.75;
      const page = doc.addPage([w, h]);
      page.drawImage(img, { x: 0, y: 0, width: w, height: h });
    } else {
      const land = img.width > img.height;
      const [pw, ph] = land ? [SIZES.A4[1], SIZES.A4[0]] : SIZES.A4;
      const s = Math.min((pw - 2 * margin) / img.width, (ph - 2 * margin) / img.height, 1.5);
      const w = img.width * s; const h = img.height * s;
      doc.addPage([pw, ph]).drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
    }
  }
  doc.setProducer('Hostwerk Office');
  return doc.save();
}

/* ---------------- metadata ---------------- */

export function readMeta(doc) {
  const g = (f) => { try { return doc[f]() || ''; } catch { return ''; } };
  return { title: g('getTitle'), author: g('getAuthor'), subject: g('getSubject'), keywords: g('getKeywords'), creator: g('getCreator'), producer: g('getProducer'), created: (() => { try { return doc.getCreationDate() || null; } catch { return null; } })(), modified: (() => { try { return doc.getModificationDate() || null; } catch { return null; } })() };
}
export async function getMetadata(bytes) { return readMeta(await loadPdf(bytes)); }

function applyMeta(doc, meta) {
  if (!meta) return;
  if (meta.title != null) doc.setTitle(String(meta.title), { showInWindowTitleBar: !!meta.title });
  if (meta.author != null) doc.setAuthor(String(meta.author));
  if (meta.subject != null) doc.setSubject(String(meta.subject));
  if (meta.keywords != null) doc.setKeywords(String(meta.keywords).split(/[,;]\s*/).filter(Boolean));
}

export async function setMetadata(bytes, meta) {
  const doc = await loadPdf(bytes);
  applyMeta(doc, meta);
  doc.setModificationDate(new Date());
  return doc.save();
}

/* ---------------- forms and signatures ---------------- */

/** fields of the document: [{ name, type, value, options, readOnly }] */
export function listFields(doc) {
  let form;
  try { form = doc.getForm(); } catch { return []; }
  return form.getFields().map((f) => {
    const name = f.getName();
    const base = { name, readOnly: f.isReadOnly() };
    try {
      if (f instanceof PDFTextField) return { ...base, type: 'text', value: f.getText() || '' };
      if (f instanceof PDFCheckBox) return { ...base, type: 'checkbox', value: f.isChecked() };
      if (f instanceof PDFRadioGroup) return { ...base, type: 'radio', value: f.getSelected() || '', options: f.getOptions() };
      if (f instanceof PDFDropdown) return { ...base, type: 'dropdown', value: f.getSelected()[0] || '', options: f.getOptions() };
      if (f instanceof PDFOptionList) return { ...base, type: 'list', value: f.getSelected(), options: f.getOptions() };
      if (f instanceof PDFSignature) return { ...base, type: 'signature' };
    } catch { /* odd field */ }
    return { ...base, type: 'other' };
  });
}

/** set values by field name; returns names that could not be set */
export function applyFormValues(doc, values, { flatten = false } = {}) {
  const failed = [];
  let form;
  try { form = doc.getForm(); } catch { return failed; }
  for (const [name, v] of Object.entries(values || {})) {
    let f;
    try { f = form.getField(name); } catch { failed.push(name); continue; }
    try {
      if (f instanceof PDFTextField) f.setText(v == null ? '' : String(v));
      else if (f instanceof PDFCheckBox) { if (v === true || v === 'On' || v === 'Yes') f.check(); else f.uncheck(); }
      else if (f instanceof PDFRadioGroup) { if (v) f.select(String(v)); else f.clear(); }
      else if (f instanceof PDFDropdown) { if (v === '' || v == null) f.clear(); else f.select(Array.isArray(v) ? String(v[0]) : String(v)); }
      else if (f instanceof PDFOptionList) { const list = (Array.isArray(v) ? v : [v]).filter((x) => x != null && x !== '').map(String); if (list.length) f.select(list); else f.clear(); }
    } catch { failed.push(name); }
  }
  if (flatten) form.flatten({ updateFieldAppearances: true });
  return failed;
}

const str = (o) => (o instanceof PDFString || o instanceof PDFHexString ? o.decodeText() : o ? String(o) : '');
/** digital signatures (shown as information – validity is not checked): [{ field, name, reason, location, date }] */
export function listSignatures(doc) {
  const out = [];
  try {
    for (const f of doc.getForm().getFields()) {
      if (!(f instanceof PDFSignature)) continue;
      const v = f.acroField.dict.lookup(PDFName.of('V'));
      const d = v instanceof PDFDict ? v : null;
      out.push({
        field: f.getName(), signed: !!d,
        name: d ? str(d.lookup(PDFName.of('Name'))) : '', reason: d ? str(d.lookup(PDFName.of('Reason'))) : '',
        location: d ? str(d.lookup(PDFName.of('Location'))) : '', date: d ? str(d.lookup(PDFName.of('M'))) : '',
        filter: d ? String(d.lookup(PDFName.of('SubFilter')) || d.lookup(PDFName.of('Filter')) || '').replace('/', '') : '',
      });
    }
  } catch { /* no form */ }
  return out;
}

/** what the file contains that a rewrite changes or the editor does not show (for the fidelity notice) */
export function inspect(doc) {
  const cat = doc.catalog;
  const has = (k) => { try { return !!cat.lookup(PDFName.of(k)); } catch { return false; } };
  let xfa = false; let pdfa = false;
  try { const af = cat.lookup(PDFName.of('AcroForm')); xfa = !!(af instanceof PDFDict && af.lookup(PDFName.of('XFA'))); } catch { /* none */ }
  try {
    const m = cat.lookup(PDFName.of('Metadata'));
    if (m && m.getContents) { const s = new TextDecoder().decode(m.getContents()); pdfa = /pdfaid:part/.test(s); }
  } catch { /* compressed or none */ }
  return { signatures: listSignatures(doc).filter((s) => s.signed), xfa, pdfa, javascript: has('Names') && (() => { try { const n = cat.lookup(PDFName.of('Names')); return !!(n instanceof PDFDict && n.lookup(PDFName.of('JavaScript'))); } catch { return false; } })(), metadataXmp: has('Metadata') };
}

/* ---------------- the whole document ---------------- */

/**
 * Writes the page model. Options:
 *   sources   [Uint8Array]           – 0 is the base document
 *   pages     [{ id, src, index, rot }]
 *   annots    { [pageId]: [annotation] }   (see annotations.js)
 *   forms     { fieldName: value }   – values for the base document's AcroForm
 *   flatten   bool                   – burn form fields into the pages
 *   meta      { title, author, subject, keywords }
 *   mode      'annot' | 'burn'       – how annotations are written
 */
export async function buildDocument({ sources, pages, annots = {}, forms = null, flatten = false, meta = null, mode = 'annot', author = '' }) {
  const base = await loadPdf(sources[0]);
  const basePages = base.getPages();
  // values first – the fields still find their widgets on the original pages
  if (forms && Object.keys(forms).length) applyFormValues(base, forms, { flatten: false });
  if (flatten) { try { base.getForm().flatten({ updateFieldAppearances: true }); } catch { /* no form */ } }

  // pages from other sources (one copyPages call per source)
  const copied = new Map();
  const extra = new Map();
  for (const p of pages) if (p.src !== 0) { if (!extra.has(p.src)) extra.set(p.src, []); extra.get(p.src).push(p.index); }
  for (const [src, idx] of extra) {
    const d = await loadPdf(sources[src]);
    const uniq = [...new Set(idx)];
    const got = await base.copyPages(d, uniq);
    uniq.forEach((ix, k) => copied.set(`${src}:${ix}`, got[k]));
  }
  const used = new Set();
  const out = [];
  for (const p of pages) {
    let page;
    if (p.src === 0) {
      if (used.has(p.index)) [page] = await base.copyPages(base, [p.index]);
      else { page = basePages[p.index]; used.add(p.index); }
    } else page = copied.get(`${p.src}:${p.index}`);
    out.push({ p, page });
  }

  // fields whose widgets all sit on removed pages are removed with them
  const keep = new Set(out.map((o) => o.page.ref.toString()));
  if (used.size < basePages.length) {
    try {
      const form = base.getForm();
      for (const f of form.getFields()) {
        const ws = f.acroField.getWidgets();
        if (ws.length && ws.every((w) => { const pr = w.P(); return pr && !keep.has(pr.toString()); })) { try { form.removeField(f); } catch { /* keep it */ } }
      }
    } catch { /* no form */ }
  }

  for (let i = base.getPageCount() - 1; i >= 0; i--) base.removePage(i);
  for (const { page } of out) base.addPage(page);

  const needFont = out.some(({ p }) => (annots[p.id] || []).length);
  const font = needFont ? await base.embedFont(StandardFonts.Helvetica) : null;
  for (const { p, page } of out) {
    if (p.rot) page.setRotation(degrees(norm(page.getRotation().angle + p.rot)));
    const list = annots[p.id];
    if (list && list.length) await writeAnnotations(base, page, list, { mode, font, author });
  }
  applyMeta(base, meta);
  base.setModificationDate(new Date());
  try { return await base.save(); } catch (e) {
    // characters the standard font cannot show in a field appearance: let the viewer draw the fields itself
    if (!forms) throw e;
    try { base.getForm().acroForm.dict.set(PDFName.of('NeedAppearances'), base.context.obj(true)); } catch { /* none */ }
    return base.save({ updateFieldAppearances: false });
  }
}

/** quick helper for tests and the "Extract" action: indices of the model's pages → model entries */
export const modelFromCount = (n) => [...Array(n).keys()].map((index) => ({ id: `p${index}`, src: 0, index, rot: 0 }));

