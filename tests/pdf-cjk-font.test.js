// The "cjk-fonts" component (Noto Sans SC/JP/KR, static TrueType, downloaded on demand – server/components.js): CJK
// annotation text is embedded as a subset and stays copyable. Runs only when the fonts are on this PC:
//   OFFICE_CJK_FONTS=<folder with NotoSansSC-Regular.ttf, NotoSansJP-Regular.ttf, NotoSansKR-Regular.ttf> node --test tests/pdf-cjk-font.test.js
// (e.g. <Hostwerk data>/office-components/cjk-fonts after installing the component) – skipped otherwise (CI has no download).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { buildDocument, missingGlyphs } from '../ui/src/editors/pdf/edit.js';
import { parseFont } from '../ui/src/editors/pdf/ttf.js';

const DIR = process.env.OFFICE_CJK_FONTS || '';
const FILES = ['NotoSansSC-Regular.ttf', 'NotoSansJP-Regular.ttf', 'NotoSansKR-Regular.ttf'];
const have = DIR && FILES.every((f) => fs.existsSync(path.join(DIR, f)));

test('CJK annotation text with the component fonts (subset, copyable)', { skip: have ? false : 'OFFICE_CJK_FONTS not set' }, async () => {
  const fonts = FILES.map((f) => new Uint8Array(fs.readFileSync(path.join(DIR, f))));
  for (const b of fonts) { const f = parseFont(b); assert.ok(!f.cff, 'TrueType outlines (glyf)'); assert.ok(f.gid(0x4e2d), 'has 中'); }
  const texts = ['中文测试 汉字', '日本語のテキスト', '한국어 텍스트'];
  const annots = { p0: texts.map((text, i) => ({ id: `a${i}`, type: 'text', x: 20, y: 200 - i * 40, text, size: 14, color: '#000000', rot: 0 })) };
  assert.deepEqual(missingGlyphs(annots, fonts), []);
  const d = await PDFDocument.create(); const h = await d.embedFont(StandardFonts.Helvetica); d.addPage([400, 300]).drawText('Seite', { x: 20, y: 270, size: 12, font: h });
  const out = await buildDocument({ sources: [await d.save()], pages: [{ id: 'p0', src: 0, index: 0, rot: 0 }], annots, mode: 'burn', fonts });
  assert.ok(out.length < 1.5 * 1048576, `subset, not the whole 10 MB font (${out.length} bytes)`);
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: out.slice(), isEvalSupported: false, verbosity: 0 }).promise;
  const text = (await (await doc.getPage(1)).getTextContent()).items.map((x) => x.str).join('');
  await doc.destroy();
  for (const s of texts) assert.ok(text.replace(/\s/g, '').includes(s.replace(/\s/g, '')), `copyable: ${s} in ${text}`);
});
