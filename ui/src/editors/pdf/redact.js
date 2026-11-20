/*
 * Real redaction – pure parts (pdf-lib, no DOM), tested in tests/pdf-redact.test.js.
 *
 * A page with marks is rewritten completely: its content (text, vector graphics, images) is replaced by one picture
 * of the page in which the marked areas are black (rendered by the editor with pdf.js), plus an invisible text layer
 * with the page's text WITHOUT the marked characters – the page stays searchable, nothing under a mark survives.
 * Annotations/form fields under a mark are removed, the structure tree (alternate texts) goes, objects nobody refers
 * to any more (the old content streams, fonts, images) are dropped from the file. verifyRedaction() checks the
 * result by extracting the text again.
 *
 *   findSensitive(text, { kinds, terms })            → [{ start, end, kind, value }]
 *   hitRects(item, a, b)                              → PDF rectangle of characters a…b of a pdf.js text item
 *   keptWords(items, marks, removed)                  → words (textlayer.js) of the text that stays
 *   applyRedactions(bytes, pages)                     → bytes   pages: [{ index, rects, image: { bytes, png? }, words }]
 *   verifyRedaction(extracted, { secrets, marks })    → { ok, leaks }
 */
import { PDFDocument, PDFName, PDFDict, PDFArray, PDFRef, PDFRawStream, PDFNumber, PDFHexString, StandardFonts, decodePDFRawStream } from 'pdf-lib';
import { addTextLayer } from './textlayer.js';

/* ---------------- what to look for ---------------- */

function ibanOk(s) {
  const v = s.replace(/\s+/g, '').toUpperCase();
  if (!/^[A-Z]{2}\d{2}[A-Z0-9]{10,30}$/.test(v)) return false;
  const r = (v.slice(4) + v.slice(0, 4)).replace(/[A-Z]/g, (c) => String(c.charCodeAt(0) - 55));
  let m = 0; for (const ch of r) m = (m * 10 + Number(ch)) % 97;
  return m === 1;
}
function luhn(s) {
  const d = s.replace(/\D/g, '');
  if (d.length < 13 || d.length > 19) return false;
  let sum = 0;
  for (let i = 0; i < d.length; i++) { let n = Number(d[d.length - 1 - i]); if (i % 2) { n *= 2; if (n > 9) n -= 9; } sum += n; }
  return sum % 10 === 0;
}

export const KINDS = {
  email: { label: 'E-mail addresses', re: /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g },
  iban: { label: 'IBANs', re: /\b[A-Z]{2}\d{2}(?: ?[A-Z0-9]{4}){2,7}(?: ?[A-Z0-9]{1,4})?\b/g, ok: ibanOk },
  phone: { label: 'Phone numbers', re: /(?<![\w/.,-])(?:\+\d{1,3}[ ./-]?|00\d{1,3}[ ./-]?|0)(?:\(0\)[ ]?)?\d{2,5}(?:[ ./-]?\d{2,}){1,4}(?![\w/-])/g, ok: (s) => { const n = s.replace(/\D/g, '').length; return n >= 7 && n <= 15; } },
  card: { label: 'Card numbers', re: /\b(?:\d[ -]?){12,18}\d\b/g, ok: luhn },
  name: { label: 'Names after a salutation', re: /\b(?:Herr|Herrn|Frau|Mr\.?|Mrs\.?|Ms\.?|Dr\.|Prof\.)\s+(?:(?:Dr\.|Prof\.)\s+)?([A-ZÄÖÜ][a-zäöüß]+(?:[- ][A-ZÄÖÜ][a-zäöüß]+)?)/g, group: 1 },
  date: { label: 'Dates of birth and other dates', re: /\b(?:\d{1,2}\.\s?\d{1,2}\.\s?(?:19|20)\d{2}|(?:19|20)\d{2}-\d{2}-\d{2})\b/g },
};

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** matches in a text (kinds = ids of KINDS, terms = own words/names, case-insensitive) – sorted, not overlapping */
export function findSensitive(text, { kinds = ['email', 'iban', 'phone'], terms = [] } = {}) {
  const out = [];
  for (const k of kinds) {
    const def = KINDS[k]; if (!def) continue;
    def.re.lastIndex = 0;
    let m;
    while ((m = def.re.exec(text))) {
      let start = m.index; let value = m[0];
      if (def.group) { value = m[def.group]; start = m.index + m[0].lastIndexOf(value); }
      value = value.replace(/[\s.,;:]+$/, '');
      if (def.ok && !def.ok(value)) continue;
      out.push({ start, end: start + value.length, kind: k, value });
    }
  }
  for (const term of terms.map((x) => String(x).trim()).filter(Boolean)) {
    const re = new RegExp(escapeRe(term).replace(/\s+/g, '\\s+'), 'gi');
    let m;
    while ((m = re.exec(text))) out.push({ start: m.index, end: m.index + m[0].length, kind: 'term', value: m[0] });
  }
  out.sort((a, b) => a.start - b.start || b.end - a.end);
  const res = [];
  for (const h of out) { const last = res[res.length - 1]; if (last && h.start < last.end) { if (h.end > last.end) last.end = h.end; continue; } res.push({ ...h }); }
  return res;
}

/* ---------------- geometry of text items ---------------- */

/** frame of a pdf.js text item: origin (baseline start), unit direction, length, height */
function frame(it) {
  const [a, b, c, d, e, f] = it.transform;
  const sx = Math.hypot(a, b) || 1;
  const h = Math.hypot(c, d) || it.height || sx;
  const len = it.width || sx * String(it.str).length * 0.5;
  return { ox: e, oy: f, ux: a / sx, uy: b / sx, len, h };
}
/** PDF rectangle (axis-aligned, generous) of characters a…b of a text item – positions are proportional */
export function hitRects(it, a, b, pad = 0.8) {
  const fr = frame(it); const n = Math.max(1, String(it.str).length);
  const s = (fr.len * a) / n; const e = (fr.len * b) / n;
  const desc = fr.h * 0.25; const asc = fr.h * 0.95;
  const pts = [[s, -desc], [e, -desc], [e, asc], [s, asc]].map(([x, y]) => [fr.ox + fr.ux * x - fr.uy * y, fr.oy + fr.uy * x + fr.ux * y]);
  const xs = pts.map((p) => p[0]); const ys = pts.map((p) => p[1]);
  const x = Math.min(...xs) - pad; const y = Math.min(...ys) - pad;
  return { x, y, w: Math.max(...xs) + pad - x, h: Math.max(...ys) + pad - y };
}
const hit = (r, s) => r.x < s.x + s.w && r.x + r.w > s.x && r.y < s.y + s.h && r.y + r.h > s.y;
const centreIn = (r, s) => { const cx = r.x + r.w / 2; const cy = r.y + r.h / 2; return cx >= s.x && cx <= s.x + s.w && cy >= s.y && cy <= s.y + s.h; };

/**
 * The text that stays: per item the characters not removed (removed: Map item index → [[a, b]]) whose box centre
 * is not inside a mark → words (runs of kept characters) for the invisible text layer.
 */
export function keptWords(items, marks = [], removed = new Map()) {
  const words = [];
  items.forEach((it, i) => {
    if (!it || !it.str || !it.transform) return;
    const str = String(it.str); const n = str.length;
    const drop = new Uint8Array(n);
    for (const [a, b] of removed.get(i) || []) for (let k = Math.max(0, a); k < Math.min(n, b); k++) drop[k] = 1;
    for (let k = 0; k < n; k++) if (!drop[k] && marks.some((m) => centreIn(hitRects(it, k, k + 1, 0), m) || hit(hitRects(it, k + 0.2, k + 0.8, 0), m))) drop[k] = 1;
    const fr = frame(it);
    let k = 0;
    while (k < n) {
      if (drop[k] || /\s/.test(str[k])) { k++; continue; }
      let e = k; while (e < n && !drop[e] && !/\s/.test(str[e])) e++;
      const s0 = (fr.len * k) / n; const s1 = (fr.len * e) / n; const desc = fr.h * 0.2;
      const P = (x, y) => [fr.ox + fr.ux * x - fr.uy * y, fr.oy + fr.uy * x + fr.ux * y];
      words.push({ text: str.slice(k, e), bl: P(s0, -desc), br: P(s1, -desc), tl: P(s0, fr.h * 0.95) });
      k = e;
    }
  });
  return words;
}

/* ---------------- writing ---------------- */

const look = (ctx, o) => (o instanceof PDFRef ? ctx.lookup(o) : o);
function rectOf(ctx, d) {
  const r = look(ctx, d.get(PDFName.of('Rect')));
  if (!(r instanceof PDFArray) || r.size() < 4) return null;
  const v = r.asArray().map((x) => { const n = look(ctx, x); return n instanceof PDFNumber ? n.asNumber() : 0; });
  return { x: Math.min(v[0], v[2]), y: Math.min(v[1], v[3]), w: Math.abs(v[2] - v[0]), h: Math.abs(v[3] - v[1]) };
}

/** drop every object that can no longer be reached from the trailer (old page contents, fonts, images …) */
export function collectGarbage(doc) {
  const ctx = doc.context;
  const seen = new Set(); const todo = [];
  const visit = (o) => {
    if (o instanceof PDFRef) { const k = o.toString(); if (seen.has(k)) return; seen.add(k); todo.push(ctx.lookup(o)); return; }
    if (o instanceof PDFDict) { for (const [, v] of o.entries()) visit(v); return; }
    if (o instanceof PDFArray) { for (let i = 0; i < o.size(); i++) visit(o.get(i)); return; }
    if (o && o.dict instanceof PDFDict) visit(o.dict);
  };
  const ti = ctx.trailerInfo;
  [ti.Root, ti.Info, ti.Encrypt, ti.ID].forEach((x) => x && visit(x));
  while (todo.length) visit(todo.pop());
  let dropped = 0;
  for (const [ref] of ctx.enumerateIndirectObjects()) if (!seen.has(ref.toString())) { ctx.delete(ref); dropped++; }
  return dropped;
}

/**
 * pages: [{ index, rects: [{x,y,w,h}], image: { bytes, png? }, words, view: [x0, y0, x1, y1] }]
 * (image = the page rendered without rotation over `view` with the rectangles painted black)
 */
export async function applyRedactions(bytes, pages, { removeStructure = true } = {}) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false });
  const ctx = doc.context;
  const all = doc.getPages();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (const p of pages) {
    const page = all[p.index];
    if (!page || !p.rects || !p.rects.length) continue;
    const node = page.node;
    // annotations under a mark (links, comments, form widgets with their values)
    const annots = look(ctx, node.get(PDFName.of('Annots')));
    if (annots instanceof PDFArray) {
      const keep = [];
      for (let i = 0; i < annots.size(); i++) {
        const d = look(ctx, annots.get(i));
        const r = d instanceof PDFDict ? rectOf(ctx, d) : null;
        if (!r || !p.rects.some((m) => hit(r, m))) { keep.push(annots.get(i)); continue; }
        if (d instanceof PDFDict && d.get(PDFName.of('Subtype')) === PDFName.of('Widget')) {
          let f = d; let guard = 0;
          while (f && guard++ < 32) { f.delete(PDFName.of('V')); f.delete(PDFName.of('AP')); f.delete(PDFName.of('DV')); f = look(ctx, f.get(PDFName.of('Parent'))); if (!(f instanceof PDFDict)) break; }
        }
      }
      if (keep.length) node.set(PDFName.of('Annots'), ctx.obj(keep)); else node.delete(PDFName.of('Annots'));
    }
    for (const k of ['Contents', 'Resources', 'Thumb', 'PieceInfo', 'Metadata', 'StructParents', 'Group', 'B']) node.delete(PDFName.of(k));
    node.set(PDFName.of('Resources'), ctx.obj({}));
    const img = p.image.png ? await doc.embedPng(p.image.bytes) : await doc.embedJpg(p.image.bytes);
    const [x0, y0, x1, y1] = p.view || [0, 0, page.getWidth(), page.getHeight()];
    page.drawImage(img, { x: x0, y: y0, width: x1 - x0, height: y1 - y0 });
    await addTextLayer(doc, page, p.words || [], { font });
  }
  if (removeStructure) {
    for (const k of ['StructTreeRoot', 'MarkInfo']) doc.catalog.delete(PDFName.of(k));
    for (const page of all) page.node.delete(PDFName.of('StructParents'));
  }
  collectGarbage(doc);
  return doc.save({ useObjectStreams: true, updateFieldAppearances: false });
}

/* ---------------- checking ---------------- */

const squash = (s) => String(s || '').normalize('NFKC').replace(/[\s­​]+/g, '').toLowerCase();

/**
 * extracted: [{ text, items: [{ str, transform, width, height }] }] of the result (pdf.js getTextContent per page)
 * secrets: strings that must be gone, marks: { pageIndex: [rect] } – nothing may sit inside a mark.
 */
export function verifyRedaction(extracted, { secrets = [], marks = {} } = {}) {
  const leaks = [];
  const wanted = [...new Set(secrets.map(squash).filter((s) => s.length >= 2))];
  extracted.forEach((pg, i) => {
    const flat = squash(pg.text != null ? pg.text : (pg.items || []).map((x) => x.str).join(''));
    for (const s of wanted) if (flat.includes(s)) leaks.push({ page: i, text: s, why: 'text' });
    for (const m of marks[i] || []) {
      for (const it of pg.items || []) {
        if (!it.str || !it.str.trim() || !it.transform) continue;
        const n = it.str.length;
        for (let k = 0; k < n; k++) {
          if (/\s/.test(it.str[k])) continue;
          const shrunk = { x: m.x + 1, y: m.y + 1, w: Math.max(0, m.w - 2), h: Math.max(0, m.h - 2) };
          if (centreIn(hitRects(it, k, k + 1, 0), shrunk)) { leaks.push({ page: i, text: it.str, why: 'area' }); k = n; }
        }
      }
    }
  });
  return { ok: leaks.length === 0, leaks };
}

/** raw-bytes check as a second net: no secret in any (decoded) stream or string of the file */
export async function scanFile(bytes, secrets) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false });
  const wanted = [...new Set(secrets.map(squash).filter((s) => s.length >= 3))];
  const found = new Set();
  const latin = (b) => { let s = ''; for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]); return s; };
  const check = (s) => { const q = squash(s); for (const w of wanted) if (q.includes(w)) found.add(w); };
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (obj instanceof PDFRawStream) {
      let data = obj.contents;
      try { data = decodePDFRawStream(obj).decode(); } catch { /* image or unknown filter */ }
      if (data.length < 4e6) {
        const s = latin(data);
        check(s);
        // strings of the content stream ((…) and <…>, simple fonts) joined – text is often split into pieces
        let joined = ''; const re = /<([0-9A-Fa-f\s]+)>|\(((?:\\.|[^\\)])*)\)/g; let m;
        while ((m = re.exec(s))) {
          if (m[1] != null) { const h = m[1].replace(/\s+/g, ''); for (let i = 0; i + 1 < h.length; i += 2) joined += String.fromCharCode(parseInt(h.slice(i, i + 2), 16)); } else joined += m[2].replace(/\\(.)/g, '$1');
        }
        if (joined) check(joined);
      }
    }
    const walk = (o) => {
      if (o instanceof PDFHexString || (o && o.constructor && o.decodeText && !(o instanceof PDFName))) { try { check(o.decodeText()); } catch { /* bytes */ } }
      if (o instanceof PDFDict) for (const [, v] of o.entries()) walk(v);
      if (o instanceof PDFArray) for (let i = 0; i < o.size(); i++) walk(o.get(i));
    };
    walk(obj instanceof PDFRawStream ? obj.dict : obj);
  }
  return [...found];
}
