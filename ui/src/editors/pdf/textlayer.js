/*
 * Invisible text layer (text render mode 3) – what OCR adds to scanned pages and what redaction keeps of a page that
 * was flattened to an image. Pure (pdf-lib only), tested in tests/pdf-ocr.test.js: after writing, pdf.js finds and
 * copies the words, and every other PDF program does too.
 *
 *   words: [{ text, bl: [x, y], br: [x, y], tl: [x, y] }]   – corners in PDF user space (bottom-left, bottom-right,
 *           top-left of the word box; any rotation)
 *   textLayerOperators(words, font, fontName)  → pdf-lib operators
 *   addTextLayer(doc, page, words, { font })   → appends them to the page (Helvetica when no font is given)
 *   imageWordsToPdf(words, toPdf)              → OCR/pixel boxes { text, x0, y0, x1, y1 } → words (toPdf: px → PDF point)
 */
import {
  StandardFonts, pushGraphicsState, popGraphicsState, beginText, endText, setTextRenderingMode,
  TextRenderingMode, setFontAndSize, setTextMatrix, setCharacterSqueeze, showText,
} from 'pdf-lib';
import { canEncodeWinAnsi } from './ttf.js';

const r3 = (v) => Math.round(v * 1000) / 1000;

/** WinAnsi-safe text for Helvetica (characters it cannot show become a space, so word positions stay) */
function safeText(font, s) {
  if (canEncodeWinAnsi(s) || typeof font.covers === 'function') return s;
  return Array.from(s, (ch) => (canEncodeWinAnsi(ch) ? ch : ' ')).join('');
}

export function textLayerOperators(words, font, fontName) {
  const ops = [pushGraphicsState(), beginText(), setTextRenderingMode(TextRenderingMode.Invisible)];
  let lastSize = -1;
  for (const w of words) {
    const raw = String(w.text || '').replace(/\s+/g, ' ').trim();
    if (!raw) continue;
    const text = `${safeText(font, raw)} `; // trailing space: copy/paste keeps word gaps
    const dx = w.br[0] - w.bl[0]; const dy = w.br[1] - w.bl[1];
    const len = Math.hypot(dx, dy);
    const ux = w.tl[0] - w.bl[0]; const uy = w.tl[1] - w.bl[1];
    const height = Math.hypot(ux, uy);
    if (len < 0.5 || height < 0.5) continue;
    const size = Math.max(1, height * 0.9);
    const natural = font.widthOfTextAtSize(text, size) || 1;
    const squeeze = Math.max(1, Math.min(1000, (len / natural) * 100 * (text.length / (text.length - 0.6))));
    const cos = dx / len; const sin = dy / len;
    // baseline a little above the box bottom (descenders)
    const bx = w.bl[0] + (ux / height) * height * 0.18; const by = w.bl[1] + (uy / height) * height * 0.18;
    if (size !== lastSize) { ops.push(setFontAndSize(fontName, r3(size))); lastSize = size; }
    ops.push(setCharacterSqueeze(r3(squeeze)), setTextMatrix(r3(cos), r3(sin), r3(-sin), r3(cos), r3(bx), r3(by)), showText(font.encodeText(text)));
  }
  ops.push(endText(), popGraphicsState());
  return ops;
}

export async function addTextLayer(doc, page, words, { font } = {}) {
  if (!words || !words.length) return 0;
  const f = font || await doc.embedFont(StandardFonts.Helvetica);
  const name = page.node.newFontDictionary('OfOcr', f.ref);
  page.pushOperators(...textLayerOperators(words, f, name));
  return words.length;
}

/** pixel boxes (top-left origin) → words in PDF space; toPdf([x, y]) converts one pixel point */
export function imageWordsToPdf(boxes, toPdf) {
  return boxes.filter((b) => b.text && b.x1 > b.x0 && b.y1 > b.y0).map((b) => ({
    text: b.text, bl: toPdf(b.x0, b.y1), br: toPdf(b.x1, b.y1), tl: toPdf(b.x0, b.y0),
  }));
}

/** pdf.js text items (with transform + width) → words with corners, for keeping the text of a flattened page */
export function itemsToWords(items) {
  const out = [];
  for (const it of items) {
    if (!it || !it.str || !it.str.trim() || !it.transform) continue;
    const [a, b, c, d, e, f] = it.transform;
    const h = Math.hypot(c, d) || Math.hypot(a, b) || it.height || 10;
    const len = it.width || Math.hypot(a, b) * it.str.length * 0.5;
    const ca = Math.hypot(a, b) ? a / Math.hypot(a, b) : 1; const sa = Math.hypot(a, b) ? b / Math.hypot(a, b) : 0;
    const desc = h * 0.2;
    const bl = [e + sa * desc, f - ca * desc];
    out.push({ text: it.str, bl, br: [bl[0] + ca * len, bl[1] + sa * len], tl: [bl[0] - sa * h, bl[1] + ca * h] });
  }
  return out;
}

