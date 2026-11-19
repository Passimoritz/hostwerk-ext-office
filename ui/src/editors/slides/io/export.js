/*
 * Raster export (browser): a rendered slide DOM node → SVG foreignObject → canvas → PNG/JPEG; a minimal image PDF
 * writer (one JPEG page per slide) for "PDF" without the Office engine; print HTML.
 */

/** canvas of a slide node (node = the .of-slides-view rendered at scale 1, size in slide px = pt) */
export async function nodeToCanvas(node, w, h, scale = 2) {
  const clone = node.cloneNode(true);
  clone.style.transform = 'none';
  clone.style.margin = '0';
  // media players can't be painted – keep their poster boxes
  clone.querySelectorAll('video,audio').forEach((m) => m.remove());
  const xml = new XMLSerializer().serializeToString(clone);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w * scale}" height="${h * scale}" viewBox="0 0 ${w} ${h}"><foreignObject x="0" y="0" width="${w}" height="${h}">${xml}</foreignObject></svg>`;
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  const img = new Image();
  img.decoding = 'sync';
  await new Promise((res, rej) => { img.onload = res; img.onerror = () => rej(new Error('Could not render the slide as an image')); img.src = url; });
  const c = document.createElement('canvas');
  c.width = Math.round(w * scale); c.height = Math.round(h * scale);
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, c.width, c.height);
  ctx.drawImage(img, 0, 0, c.width, c.height);
  return c;
}
export async function canvasBytes(canvas, type = 'image/png', quality = 0.92) {
  const blob = await new Promise((res) => canvas.toBlob(res, type, quality));
  if (!blob) throw new Error('Could not create the image');
  return new Uint8Array(await blob.arrayBuffer());
}

/** PDF with one full-page JPEG per page. pages: [{ jpeg: Uint8Array, pxW, pxH }], page size in pt */
export function imagePdf(pages, wPt, hPt, title = '') {
  const enc = new TextEncoder();
  const chunks = []; let len = 0;
  const offsets = [];
  const push = (x) => { const b = typeof x === 'string' ? enc.encode(x) : x; chunks.push(b); len += b.length; };
  const obj = (n, body) => { offsets[n] = len; push(`${n} 0 obj\n`); if (typeof body === 'function') body(); else push(body); push('\nendobj\n'); };
  push('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');
  const n = pages.length;
  // 1 catalog, 2 pages, 3 info, then per page: page, content, image
  const kids = pages.map((_, i) => `${4 + i * 3} 0 R`).join(' ');
  obj(1, '<< /Type /Catalog /Pages 2 0 R >>');
  obj(2, `<< /Type /Pages /Count ${n} /Kids [${kids}] >>`);
  const esc = (s) => String(s).replace(/[\\()]/g, (c) => `\\${c}`).replace(/[^\x20-\x7e]/g, '?');
  obj(3, `<< /Producer (Hostwerk Office) /Title (${esc(title)}) >>`);
  pages.forEach((p, i) => {
    const pn = 4 + i * 3; const cn = pn + 1; const im = pn + 2;
    obj(pn, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${wPt.toFixed(2)} ${hPt.toFixed(2)}] /Resources << /XObject << /Im${i} ${im} 0 R >> >> /Contents ${cn} 0 R >>`);
    const content = `q ${wPt.toFixed(2)} 0 0 ${hPt.toFixed(2)} 0 0 cm /Im${i} Do Q`;
    obj(cn, `<< /Length ${content.length} >>\nstream\n${content}\nendstream`);
    obj(im, () => { push(`<< /Type /XObject /Subtype /Image /Width ${p.pxW} /Height ${p.pxH} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`); push(p.jpeg); push('\nendstream'); });
  });
  const xref = len;
  const total = 4 + n * 3;
  let x = `xref\n0 ${total}\n0000000000 65535 f \n`;
  for (let k = 1; k < total; k++) x += `${String(offsets[k]).padStart(10, '0')} 00000 n \n`;
  push(x);
  push(`trailer\n<< /Size ${total} /Root 1 0 R /Info 3 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
  const out = new Uint8Array(len); let o = 0;
  for (const c of chunks) { out.set(c, o); o += c.length; }
  return out;
}

/** print HTML: one slide image per page (optionally with notes) */
export function printHtml(images, wPt, hPt, { title = '', notes = null } = {}) {
  const landscape = wPt >= hPt;
  const page = notes ? 'A4 portrait' : `${wPt}pt ${hPt}pt`;
  const esc = (s) => String(s || '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const body = images.map((src, i) => (notes
    ? `<section class="n"><img src="${src}" alt=""><div class="t">${esc(notes[i]).replace(/\n/g, '<br>')}</div></section>`
    : `<section><img src="${src}" alt=""></section>`)).join('');
  return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}</title><style>
@page { size: ${page}; margin: 0; }
html, body { margin: 0; padding: 0; background: #fff; }
section { width: ${notes ? '210mm' : `${wPt}pt`}; height: ${notes ? '297mm' : `${hPt}pt`}; page-break-after: always; break-after: page; overflow: hidden; display: flex; flex-direction: column; align-items: center; }
section:last-child { page-break-after: auto; break-after: auto; }
section > img { width: 100%; height: ${notes ? 'auto' : '100%'}; display: block; object-fit: contain; }
section.n { padding: 15mm; box-sizing: border-box; gap: 8mm; }
section.n > img { border: 1px solid #999; }
.t { font: 12pt/1.4 'Segoe UI', Arial, sans-serif; color: #000; align-self: stretch; }
</style></head><body data-orientation="${landscape ? 'landscape' : 'portrait'}">${body}</body></html>`;
}
