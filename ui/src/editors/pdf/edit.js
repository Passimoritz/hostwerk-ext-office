/*
 * PDF operations with pdf-lib – pure functions on bytes (no DOM), unit-tested in tests/pdf-edit.test.js.
 *
 * The editor keeps a page model { sources: [bytes], pages: [{ id, src, index, rot }], annots: { pageId: [...] },
 * forms: { fieldName: value }, meta } and writes it with buildDocument(): the first source is the base document (its
 * AcroForm, outline, XMP/PDF-A metadata, attachments stay), pages of other sources are copied in.
 */
import {
  PDFDocument, PDFName, PDFDict, PDFString, PDFHexString, StandardFonts, PageSizes, degrees, PDFArray, PDFRef, PDFNumber,
  PDFTextField, PDFCheckBox, PDFRadioGroup, PDFDropdown, PDFOptionList, PDFSignature, PDFObjectCopier,
} from 'pdf-lib';
import { writeAnnotations, markDocument, MARKER } from './annotations.js';
import { canEncodeWinAnsi, parseFont, UnicodeFont } from './ttf.js';

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
export async function buildDocument({ sources, pages, annots = {}, forms = null, flatten = false, meta = null, mode = 'annot', author = '', fonts = [], names = [] }) {
  const base = await loadPdf(sources[0]);
  const basePages = base.getPages();
  // values first – the fields still find their widgets on the original pages
  if (forms && Object.keys(forms).length) applyFormValues(base, forms, { flatten: false });
  if (flatten) { try { base.getForm().flatten({ updateFieldAppearances: true }); } catch { /* no form */ } }

  // pages from other sources (one copyPages call per source)
  const copied = new Map();
  const extra = new Map();
  for (const p of pages) if (p.src !== 0) { if (!extra.has(p.src)) extra.set(p.src, []); extra.get(p.src).push(p.index); }
  const adopt = [];
  for (const [src, idx] of extra) {
    const d = await loadPdf(sources[src]);
    const uniq = [...new Set(idx)];
    const got = await base.copyPages(d, uniq);
    uniq.forEach((ix, k) => copied.set(`${src}:${ix}`, got[k]));
    adopt.push({ d, got, refMap: new Map(uniq.map((ix, k) => [d.getPage(ix).ref.toString(), got[k].ref])), title: names[src] || '' });
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
  // inserted pages keep their form fields and bookmarks
  adopt.forEach((a, i) => {
    try { adoptForms(base, a.d, a.got, i + 2); } catch { /* the fields still show, only not as one form */ }
    try { adoptOutline(base, a.d, a.refMap, a.title); } catch { /* no bookmarks */ }
  });

  const needFont = out.some(({ p }) => (annots[p.id] || []).length);
  const font = needFont ? await base.embedFont(StandardFonts.Helvetica) : null;
  const picker = needFont ? fontPicker(base, font, fonts) : null;
  let editable = 0;
  for (const { p, page } of out) {
    if (p.rot) page.setRotation(degrees(norm(page.getRotation().angle + p.rot)));
    const list = annots[p.id];
    if (list && list.length) editable += await writeAnnotations(base, page, list, { mode, fonts: picker, author });
  }
  if (picker) picker.finish();
  if (editable) markDocument(base);
  else if (base.catalog.get(PDFName.of(MARKER))) base.catalog.delete(PDFName.of(MARKER));
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


/* ---------------- fonts for annotation text ---------------- */

/**
 * Helvetica (WinAnsi) when it can show the text, else the first Unicode font (bundled Latin/Greek/Cyrillic font,
 * then the CJK fonts of the component) that has every character – embedded as a subset at the end (finish()).
 */
export function fontPicker(doc, helv, fonts = []) {
  const parsed = []; const unis = [];
  const get = (i) => { if (parsed[i] === undefined) { try { parsed[i] = parseFont(fonts[i]); } catch { parsed[i] = null; } } return parsed[i]; };
  const uni = (i) => { if (!unis[i]) unis[i] = new UnicodeFont(doc, [get(i)]); return { font: unis[i], ref: unis[i].ref, key: `OfUni${i}` }; };
  const helvetica = { font: helv, ref: helv.ref, key: 'OfHelv' };
  return {
    pick(text) {
      if (canEncodeWinAnsi(text)) return helvetica;
      let best = -1; let bestMissing = Infinity;
      for (let i = 0; i < fonts.length; i++) {
        const f = get(i); if (!f) continue;
        let missing = 0; for (const ch of String(text)) { const c = ch.codePointAt(0); if (c > 31 && !f.gid(c)) missing++; }
        if (!missing) return uni(i);
        if (missing < bestMissing) { bestMissing = missing; best = i; }
      }
      return best >= 0 ? uni(best) : helvetica;
    },
    finish() { unis.forEach((u) => u && u.embed()); },
  };
}

/** characters of the annotation texts that neither Helvetica nor the given fonts can show */
export function missingGlyphs(annots, fonts = []) {
  const parsed = fonts.map((b) => { try { return parseFont(b); } catch { return null; } }).filter(Boolean);
  const miss = new Set();
  for (const list of Object.values(annots || {})) {
    for (const a of list || []) {
      if (a.type !== 'text' && a.type !== 'stamp') continue;
      for (const ch of String(a.text || '')) {
        const c = ch.codePointAt(0);
        if (c < 32 || canEncodeWinAnsi(ch)) continue;
        if (!parsed.some((f) => f.gid(c))) miss.add(ch);
      }
    }
  }
  return [...miss];
}

/* ---------------- forms and bookmarks of inserted pages ---------------- */

const look = (ctx, o) => (o instanceof PDFRef ? ctx.lookup(o) : o);

/** root fields of the widgets on the copied pages → into the base document's AcroForm (renamed on a name clash) */
export function adoptForms(base, src, pages, n = 2) {
  const ctx = base.context;
  const roots = new Map();
  for (const page of pages) {
    const annots = look(ctx, page.node.get(PDFName.of('Annots')));
    if (!(annots instanceof PDFArray)) continue;
    for (let i = 0; i < annots.size(); i++) {
      let ref = annots.get(i); let d = look(ctx, ref);
      if (!(d instanceof PDFDict) || d.get(PDFName.of('Subtype')) !== PDFName.of('Widget')) continue;
      let guard = 0;
      while (d.get(PDFName.of('Parent')) instanceof PDFRef && guard++ < 32) { ref = d.get(PDFName.of('Parent')); d = ctx.lookup(ref); }
      if (ref instanceof PDFRef) roots.set(ref.toString(), { ref, d });
    }
  }
  if (!roots.size) return 0;
  let acro = look(ctx, base.catalog.get(PDFName.of('AcroForm')));
  if (!(acro instanceof PDFDict)) { acro = ctx.obj({ Fields: [] }); base.catalog.set(PDFName.of('AcroForm'), ctx.register(acro)); }
  let fields = look(ctx, acro.get(PDFName.of('Fields')));
  if (!(fields instanceof PDFArray)) { fields = ctx.obj([]); acro.set(PDFName.of('Fields'), fields); }
  const nameOf = (d) => { const t = d.get(PDFName.of('T')); return t && t.decodeText ? t.decodeText() : ''; };
  const taken = new Set();
  for (let i = 0; i < fields.size(); i++) { const d = look(ctx, fields.get(i)); if (d instanceof PDFDict) taken.add(nameOf(d)); }
  let added = 0;
  for (const { ref, d } of roots.values()) {
    if (fields.asArray().some((x) => x instanceof PDFRef && x.toString() === ref.toString())) continue;
    let name = nameOf(d);
    if (name && taken.has(name)) { let k = n; while (taken.has(`${name} (${k})`)) k++; name = `${name} (${k})`; d.set(PDFName.of('T'), PDFHexString.fromText(name)); }
    taken.add(name);
    fields.push(ref); added++;
  }
  // default appearance + resources (fonts) of the source form
  const sAcro = look(src.context, src.catalog.get(PDFName.of('AcroForm')));
  if (sAcro instanceof PDFDict) {
    const copier = PDFObjectCopier.for(src.context, ctx);
    if (!acro.get(PDFName.of('DA')) && sAcro.get(PDFName.of('DA'))) acro.set(PDFName.of('DA'), copier.copy(look(src.context, sAcro.get(PDFName.of('DA')))));
    const sDR = look(src.context, sAcro.get(PDFName.of('DR')));
    if (sDR instanceof PDFDict) {
      const dr = look(ctx, acro.get(PDFName.of('DR')));
      if (!(dr instanceof PDFDict)) acro.set(PDFName.of('DR'), copier.copy(sDR));
      else {
        const sFont = look(src.context, sDR.get(PDFName.of('Font')));
        let fontD = look(ctx, dr.get(PDFName.of('Font')));
        if (sFont instanceof PDFDict) {
          if (!(fontD instanceof PDFDict)) { fontD = ctx.obj({}); dr.set(PDFName.of('Font'), fontD); }
          for (const [k, v] of sFont.entries()) if (!fontD.get(k)) fontD.set(k, copier.copy(v));
        }
      }
    }
  }
  return added;
}

function nameTree(ctx, node, out = new Map(), depth = 0) {
  const n = look(ctx, node);
  if (!(n instanceof PDFDict) || depth > 32) return out;
  const names = look(ctx, n.get(PDFName.of('Names')));
  if (names instanceof PDFArray) for (let i = 0; i + 1 < names.size(); i += 2) { const k = look(ctx, names.get(i)); out.set(k && k.decodeText ? k.decodeText() : String(k), names.get(i + 1)); }
  const kids = look(ctx, n.get(PDFName.of('Kids')));
  if (kids instanceof PDFArray) for (let i = 0; i < kids.size(); i++) nameTree(ctx, kids.get(i), out, depth + 1);
  return out;
}

function link(parentRef, parent, kids, open) {
  kids.forEach((k, i) => {
    k.dict.set(PDFName.of('Parent'), parentRef);
    if (i > 0) { k.dict.set(PDFName.of('Prev'), kids[i - 1].ref); kids[i - 1].dict.set(PDFName.of('Next'), k.ref); }
  });
  parent.set(PDFName.of('First'), kids[0].ref);
  parent.set(PDFName.of('Last'), kids[kids.length - 1].ref);
  parent.set(PDFName.of('Count'), PDFNumber.of(open ? kids.length : -kids.length));
}

/** bookmarks of the source that lead to copied pages → appended to the base outline (under the source's name) */
export function adoptOutline(base, src, refMap, title = '') {
  const sctx = src.context; const ctx = base.context;
  const root = look(sctx, src.catalog.get(PDFName.of('Outlines')));
  if (!(root instanceof PDFDict)) return 0;
  let named = null;
  const namedDest = (key) => {
    if (!named) {
      named = new Map();
      const dests = look(sctx, src.catalog.get(PDFName.of('Dests')));
      if (dests instanceof PDFDict) for (const [k, v] of dests.entries()) named.set(k.decodeText(), v);
      const names = look(sctx, src.catalog.get(PDFName.of('Names')));
      if (names instanceof PDFDict) nameTree(sctx, names.get(PDFName.of('Dests')), named);
    }
    let v = look(sctx, named.get(key));
    if (v instanceof PDFDict) v = look(sctx, v.get(PDFName.of('D')));
    return v;
  };
  const destOf = (item) => {
    let dest = look(sctx, item.get(PDFName.of('Dest')));
    if (!dest) { const a = look(sctx, item.get(PDFName.of('A'))); if (a instanceof PDFDict && a.get(PDFName.of('S')) === PDFName.of('GoTo')) dest = look(sctx, a.get(PDFName.of('D'))); }
    if (dest && !(dest instanceof PDFArray) && dest.decodeText) dest = namedDest(dest.decodeText());
    if (!(dest instanceof PDFArray) || !(dest.get(0) instanceof PDFRef)) return null;
    const to = refMap.get(dest.get(0).toString());
    if (!to) return null;
    const rest = dest.asArray().slice(1).map((x) => { const v = look(sctx, x); return v instanceof PDFName || v instanceof PDFNumber ? v : ctx.obj(null); });
    return ctx.obj([to, ...rest]);
  };
  // copy a chain of siblings → [{ ref, dict }]; items without a kept page and without kept children are dropped
  const copyChain = (first, depth = 0) => {
    const items = []; let cur = look(sctx, first); let guard = 0;
    while (cur instanceof PDFDict && guard++ < 5000 && depth < 32) {
      const kids = copyChain(cur.get(PDFName.of('First')), depth + 1);
      const dest = destOf(cur);
      if (dest || kids.length) {
        const t = look(sctx, cur.get(PDFName.of('Title')));
        const dict = ctx.obj({});
        dict.set(PDFName.of('Title'), PDFHexString.fromText(t && t.decodeText ? t.decodeText() : ''));
        if (dest) dict.set(PDFName.of('Dest'), dest);
        const ref = ctx.register(dict);
        if (kids.length) link(ref, dict, kids, true);
        items.push({ ref, dict });
      }
      cur = look(sctx, cur.get(PDFName.of('Next')));
    }
    return items;
  };
  let items = copyChain(root.get(PDFName.of('First')));
  if (!items.length) return 0;
  if (title) {
    const dict = ctx.obj({}); dict.set(PDFName.of('Title'), PDFHexString.fromText(title));
    const first = items.find((x) => x.dict.get(PDFName.of('Dest')));
    if (first) dict.set(PDFName.of('Dest'), first.dict.get(PDFName.of('Dest')));
    const ref = ctx.register(dict);
    link(ref, dict, items, false);
    items = [{ ref, dict }];
  }
  let brootRef = base.catalog.get(PDFName.of('Outlines'));
  let broot = look(ctx, brootRef);
  if (!(broot instanceof PDFDict)) { broot = ctx.obj({ Type: 'Outlines' }); brootRef = ctx.register(broot); base.catalog.set(PDFName.of('Outlines'), brootRef); }
  const lastRef = broot.get(PDFName.of('Last'));
  const last = look(ctx, lastRef);
  items.forEach((it, i) => {
    it.dict.set(PDFName.of('Parent'), brootRef);
    if (i > 0) { it.dict.set(PDFName.of('Prev'), items[i - 1].ref); items[i - 1].dict.set(PDFName.of('Next'), it.ref); }
  });
  if (last instanceof PDFDict) { last.set(PDFName.of('Next'), items[0].ref); items[0].dict.set(PDFName.of('Prev'), lastRef); } else broot.set(PDFName.of('First'), items[0].ref);
  broot.set(PDFName.of('Last'), items[items.length - 1].ref);
  const cnt = look(ctx, broot.get(PDFName.of('Count')));
  broot.set(PDFName.of('Count'), PDFNumber.of(Math.abs(cnt instanceof PDFNumber ? cnt.asNumber() : 0) + items.length));
  return items.length;
}
