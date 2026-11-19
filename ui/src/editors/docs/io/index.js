/*
 * Format dispatcher: readDocument(bytes, ext) → { model, html? } and writeDocument(model, ext) → Uint8Array.
 * The content decides, not the file name: a .doc that is really a docx/rtf/html (common) opens correctly.
 */
import { defaultModel, utf8, encUtf8, docText, fixDoc, parseDataUrl, walk, headingsOf } from './pm.js';
import { readDocx, docxToText } from './docx-read.js';
import { writeDocx } from './docx-write.js';
import { readOdt } from './odt-read.js';
import { writeOdt } from './odt-write.js';
import { parseRtf, writeRtf } from './rtf.js';
import { readDoc } from './doc-read.js';
import { mdToHtml, htmlToMd } from './md.js';
import { docToHtml, standaloneHtml } from './render-html.js';

export const NATIVE_WRITE = ['docx', 'docm', 'dotx', 'dotm', 'odt', 'ott', 'fodt', 'rtf', 'txt', 'md', 'markdown', 'html', 'htm', 'xhtml'];
export const EXPORT_FORMATS = [
  { ext: 'docx', label: 'Word document (.docx)' },
  { ext: 'odt', label: 'OpenDocument text (.odt)' },
  { ext: 'rtf', label: 'Rich Text (.rtf)' },
  { ext: 'html', label: 'Web page (.html)' },
  { ext: 'md', label: 'Markdown (.md)' },
  { ext: 'txt', label: 'Plain text (.txt)' },
  { ext: 'pdf', label: 'PDF' },
];
export const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', docm: 'application/vnd.ms-word.document.macroEnabled.12', dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template', dotm: 'application/vnd.ms-word.template.macroEnabled.12',
  odt: 'application/vnd.oasis.opendocument.text', ott: 'application/vnd.oasis.opendocument.text-template', fodt: 'application/vnd.oasis.opendocument.text-flat-xml', rtf: 'application/rtf', txt: 'text/plain', md: 'text/markdown', markdown: 'text/markdown', html: 'text/html', htm: 'text/html', xhtml: 'application/xhtml+xml', pdf: 'application/pdf', doc: 'application/msword',
};

export function sniff(bytes, ext) {
  const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes || []);
  if (u[0] === 0x50 && u[1] === 0x4b) {
    const head = new TextDecoder('latin1').decode(u.subarray(0, Math.min(u.length, 4000)));
    if (/mimetypeapplication\/vnd\.oasis\.opendocument/.test(head)) return 'odt';
    return ['odt', 'ott'].includes(ext) ? 'odt' : 'docx';
  }
  if (u[0] === 0xd0 && u[1] === 0xcf && u[2] === 0x11 && u[3] === 0xe0) return 'doc';
  const start = new TextDecoder('latin1').decode(u.subarray(0, 600)).replace(/^﻿|^ï»¿/, '').trimStart();
  if (start.startsWith('{\\rtf')) return 'rtf';
  if (/^<\?xml[\s\S]*<office:document/.test(start) || /^<office:document/.test(start)) return 'fodt';
  if (['html', 'htm', 'xhtml'].includes(ext) || /^<!doctype html|^<html/i.test(start)) return 'html';
  if (['md', 'markdown'].includes(ext)) return 'md';
  if (['docx', 'docm', 'dotx', 'dotm', 'odt', 'ott', 'doc', 'dot'].includes(ext) && !u.length) return 'empty';
  return ext === 'rtf' ? 'rtf' : 'txt';
}

/** → { model, html } – html is set for HTML/Markdown sources (the editor parses it with its schema) */
export async function readDocument(bytes, ext, settings) {
  const kind = sniff(bytes, ext);
  const mk = () => defaultModel(settings);
  switch (kind) {
    case 'docx': return { model: await readDocx(bytes, settings, { ext: ['docx', 'docm', 'dotx', 'dotm'].includes(ext) ? ext : 'docx' }) };
    case 'odt': return { model: await readOdt(bytes, settings, { ext: ext === 'ott' ? 'ott' : 'odt' }) };
    case 'fodt': return { model: await readOdt(bytes, settings, { ext: 'fodt' }) };
    case 'doc': return { model: readDoc(bytes, settings) };
    case 'rtf': return { model: parseRtf(new TextDecoder('latin1').decode(bytes), settings) };
    case 'html': {
      const { importHtml } = await import('./html.js');
      const r = importHtml(utf8(bytes));
      const model = mk();
      if (r.meta) for (const k of ['page', 'header', 'footer', 'titlePage', 'comments', 'defaults', 'meta']) if (r.meta[k] != null) model[k] = r.meta[k];
      if (r.title && !model.meta.title) model.meta.title = r.title;
      return { model, html: r.html };
    }
    case 'md': {
      const { importHtml } = await import('./html.js');
      return { model: mk(), html: importHtml(mdToHtml(utf8(bytes))).html };
    }
    case 'empty': return { model: mk() };
    default: return { model: textModel(utf8(bytes), settings) };
  }
}

export function textModel(text, settings) {
  const model = defaultModel(settings);
  model.defaults.spaceAfter = 0;
  const content = [];
  for (const page of String(text).replace(/\r\n?/g, '\n').split('\f')) {
    if (content.length) content.push({ type: 'pageBreak' });
    for (const line of page.split('\n')) content.push(line ? { type: 'paragraph', content: [{ type: 'text', text: line }] } : { type: 'paragraph' });
  }
  model.doc = fixDoc({ type: 'doc', content });
  return model;
}

/** browser only: images other formats cannot hold → PNG; remote images → embedded */
export async function prepareImages(model, ext) {
  if (typeof document === 'undefined') return model;
  const needPng = ['docx', 'docm', 'dotx', 'dotm', 'rtf'].includes(ext);
  const jobs = [];
  walk(model.doc, (n) => {
    if (n.type !== 'image' || !n.attrs || !n.attrs.src) return;
    const d = parseDataUrl(n.attrs.src);
    if ((d && needPng && !['png', 'jpg', 'gif', 'bmp'].includes(d.ext)) || (!d && /^https?:/.test(n.attrs.src))) jobs.push(n);
  });
  for (const n of jobs) {
    try {
      const png = await toPng(n.attrs.src, !needPng);
      if (png) n.attrs = { ...n.attrs, src: png };
    } catch { /* keep as is */ }
  }
  return model;
}
async function toPng(src, keepFormat) {
  if (/^https?:/.test(src)) {
    const r = await fetch(src, { mode: 'cors' });
    const blob = await r.blob();
    const url = await new Promise((res) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.readAsDataURL(blob); });
    if (keepFormat && /^data:image\/(png|jpeg|gif)/.test(url)) return url;
    src = url;
  }
  const img = new Image();
  img.src = src;
  await img.decode();
  const c = document.createElement('canvas');
  c.width = img.naturalWidth || 300; c.height = img.naturalHeight || 200;
  c.getContext('2d').drawImage(img, 0, 0);
  return c.toDataURL('image/png');
}

/** refresh every table of contents from the headings */
export function refreshToc(doc) {
  const hs = headingsOf(doc, 3);
  walk(doc, (n) => { if (n.type === 'toc') n.attrs = { ...n.attrs, entries: hs.map((h) => ({ level: h.level, text: h.text, index: h.index })) }; });
  return doc;
}

export async function writeDocument(model, ext) {
  const e = String(ext || 'docx').toLowerCase();
  switch (e) {
    case 'docx': case 'docm': case 'dotx': case 'dotm': return writeDocx(await prepareImages(model, e), { ext: e });
    case 'odt': case 'ott': case 'fodt': return writeOdt(await prepareImages(model, e), { ext: e });
    case 'rtf': return encUtf8(writeRtf(await prepareImages(model, e)));
    case 'html': case 'htm': return encUtf8(standaloneHtml(await prepareImages(model, e)));
    case 'xhtml': { const { toXhtml } = await import('./html.js'); return encUtf8(toXhtml(standaloneHtml(await prepareImages(model, e)))); }
    case 'md': case 'markdown': return encUtf8(htmlToMd(docToHtml(model.doc, { exportMode: true })));
    case 'txt': return encUtf8(docText(model.doc).replace(/\n/g, '\r\n'));
    default: throw new Error(`Cannot write .${e}`);
  }
}

export async function documentText(bytes, ext) {
  const kind = sniff(bytes, ext);
  if (kind === 'docx') return docxToText(bytes);
  if (kind === 'html') return utf8(bytes).replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ').replace(/<br\s*\/?>|<\/(p|div|h\d|li|tr)>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/\n{3,}/g, '\n\n').trim();
  if (kind === 'md' || kind === 'txt') return utf8(bytes);
  const { model } = await readDocument(bytes, ext);
  return docText(model.doc);
}
