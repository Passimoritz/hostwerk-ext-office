/*
 * pdf.js set-up shared by the pdf and viewer chunks (each bundles its own copy of pdf.js; the worker is the separate
 * asset ui/assets/pdfworker.js, built from ui/src/editors/pdfworker). No eval: isEvalSupported is always false.
 *
 *   pdfjsLib()                      → the pdf.js module with GlobalWorkerOptions.workerSrc set
 *   openPdf(bytes, { password, onPassword(wrong) → Promise<string|null> })  → PDFDocumentProxy
 *   renderPageToCanvas(page, canvas, { scale, rotation, signal })
 *   pdfText(doc)                    → plain text of all pages (pages separated by a blank line)
 */
import * as pdfjs from 'pdfjs-dist';

let configured = false;
export function pdfjsLib() {
  if (!configured) {
    configured = true;
    try {
      const me = new URL(import.meta.url);
      const w = new URL('pdfworker.js', me);
      w.search = me.search;
      pdfjs.GlobalWorkerOptions.workerSrc = w.href;
    } catch { /* not in a browser module – tests never get here */ }
  }
  return pdfjs;
}

export function openPdf(bytes, { password, onPassword } = {}) {
  const lib = pdfjsLib();
  const task = lib.getDocument({
    data: bytes.slice(), // pdf.js transfers (detaches) the buffer it gets
    password, isEvalSupported: false, enableXfa: false, useSystemFonts: true, fontExtraProperties: false,
  });
  if (onPassword) {
    task.onPassword = (update, reason) => {
      Promise.resolve(onPassword(reason === lib.PasswordResponses.INCORRECT_PASSWORD)).then((pw) => {
        if (pw == null) task.destroy(); else update(pw);
      }, () => task.destroy());
    };
  }
  return task.promise;
}

export async function renderPageToCanvas(page, canvas, { scale = 1, rotation = 0, background = '#ffffff' } = {}) {
  const viewport = page.getViewport({ scale, rotation: (page.rotate + rotation) % 360 });
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const task = page.render({ canvasContext: canvas.getContext('2d'), viewport, background });
  await task.promise;
  return viewport;
}

/** text of one page, lines kept where pdf.js reports an end of line */
export async function pageText(page) {
  const tc = await page.getTextContent();
  let s = '';
  for (const it of tc.items) { if (it.str === undefined) continue; s += it.str; if (it.hasEOL) s += '\n'; }
  return s.replace(/[ \t]+\n/g, '\n').trim();
}

export async function pdfText(doc, { maxPages = 2000 } = {}) {
  const out = [];
  for (let i = 1; i <= Math.min(doc.numPages, maxPages); i++) out.push(await pageText(await doc.getPage(i)));
  return out.join('\n\n');
}
