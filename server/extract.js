'use strict';
/*
 * Text, facts and thumbnails from office files – without office programs:
 *   text(buf, ext, libs)  → plain text (docx/docm/dotx/dotm, xlsx/xlsm/xltx, pptx/pptm/ppsx/potx, odt/ods/odp/ott/ots/otp,
 *                           fodt/fods/fodp, rtf, txt/md/csv/tsv, html, xml, json, pdf (libs.unpdf), eml (libs.mailparser))
 *   info(buf, ext, libs)  → { pages?, words?, sheets?, slides?, author?, title?, created?, modified? }
 *   thumb(buf, ext)       → { type, data } embedded thumbnail (docProps/thumbnail.*, Thumbnails/thumbnail.png) or null
 * Everything here only READS the file – nothing in a file is ever executed.
 */
const { readZip, isZip } = require('./zip');
const cfb = require('./cfb');
const iwork = require('./iwork');

const MAX_TEXT = 400000; // characters per file in the search index / for the AI

const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'", nbsp: ' ' };
function decode(s) {
  return String(s).replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e) => {
    if (e[0] === '#') { const n = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10); try { return String.fromCodePoint(n); } catch { return ''; } }
    return ENT[e.toLowerCase()] != null ? ENT[e.toLowerCase()] : m;
  });
}
const tidy = (s) => String(s).replace(/\r\n?/g, '\n').replace(/[ \t ]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
const clip = (s) => (s.length > MAX_TEXT ? s.slice(0, MAX_TEXT) : s);

/** Bytes → string (UTF-8 / UTF-16 with BOM, else UTF-8 with latin1 fallback for broken sequences) */
function bytesToString(buf) {
  if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) return buf.subarray(2).toString('utf16le');
  if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff) { const b = Buffer.from(buf.subarray(2)); b.swap16(); return b.toString('utf16le'); }
  const s = buf.toString('utf8').replace(/^﻿/, '');
  return s.includes('�') && !buf.toString('latin1').includes('�') ? buf.toString('latin1') : s;
}

const zipText = (z, name) => { const e = z.get(name); return e ? e.read().toString('utf8') : ''; };
const num = (a, b) => Number((a.match(/(\d+)\.xml$/) || [])[1] || 0) - Number((b.match(/(\d+)\.xml$/) || [])[1] || 0);

/** WordprocessingML → text (paragraphs, tabs, breaks) */
function wordXml(xml) {
  return decode(xml
    .replace(/<w:tab\/>|<w:tab [^>]*\/>/g, '\t')
    .replace(/<w:(br|cr)[^>]*\/>/g, '\n')
    .replace(/<\/w:p>/g, '\n')
    .replace(/<w:delText[^>]*>[\s\S]*?<\/w:delText>/g, '') // deleted (tracked) text
    .replace(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g, '') // field codes
    .replace(/<[^>]+>/g, ''));
}
function docxText(z) {
  const names = [...z.keys()];
  const parts = ['word/document.xml', ...names.filter((n) => /^word\/(header|footer)\d*\.xml$/.test(n)).sort(), 'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'];
  return tidy(parts.map((n) => wordXml(zipText(z, n))).filter((t) => t.trim()).join('\n\n'));
}

/** SpreadsheetML → text: sheet name, then one line per row (cells separated by tabs) */
function xlsxText(z) {
  const shared = [];
  const ss = zipText(z, 'xl/sharedStrings.xml');
  for (const m of ss.matchAll(/<si>([\s\S]*?)<\/si>/g)) shared.push(decode([...m[1].replace(/<rPh[\s\S]*?<\/rPh>/g, '').matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]).join('')));
  const wb = zipText(z, 'xl/workbook.xml');
  const names = [...wb.matchAll(/<sheet [^>]*name="([^"]*)"/g)].map((m) => decode(m[1]));
  const sheets = [...z.keys()].filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort(num);
  const out = [];
  sheets.forEach((n, i) => {
    const xml = zipText(z, n);
    const rows = [];
    for (const r of xml.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)) {
      const cells = [];
      for (const c of r[1].matchAll(/<c([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const attrs = c[1] || ''; const body = c[2] || '';
        const t = (attrs.match(/\st="([^"]+)"/) || [])[1];
        let v = '';
        if (t === 'inlineStr') v = decode([...body.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]).join(''));
        else { const raw = (body.match(/<v>([\s\S]*?)<\/v>/) || [])[1]; if (raw != null) v = t === 's' ? shared[Number(raw)] || '' : decode(raw); }
        if (v !== '') cells.push(v);
      }
      if (cells.length) rows.push(cells.join('\t'));
    }
    out.push(`${names[i] || `Sheet ${i + 1}`}\n${rows.join('\n')}`);
  });
  if (!sheets.length && shared.length) out.push(shared.join('\n'));
  return tidy(out.join('\n\n'));
}

/** DrawingML (slides/notes) → text, one line per paragraph */
const drawingXml = (xml) => decode(xml.replace(/<a:br\/>/g, '\n').replace(/<\/a:p>/g, '\n').replace(/<a:t>([\s\S]*?)<\/a:t>|<[^>]+>/g, (m, t) => (t != null ? t : '')));
function pptxText(z) {
  const slides = [...z.keys()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort(num);
  return tidy(slides.map((n, i) => {
    const notes = zipText(z, `ppt/notesSlides/notesSlide${(n.match(/(\d+)\.xml$/) || [])[1]}.xml`);
    const nt = notes ? tidy(drawingXml(notes)).replace(/^\d+$/m, '') : '';
    return `--- ${i + 1} ---\n${tidy(drawingXml(zipText(z, n)))}${nt.trim() ? `\n\n${nt.trim()}` : ''}`;
  }).join('\n\n'));
}

/** ODF content.xml (or a flat .fodt/.fods/.fodp file) → text */
function odfXml(xml) {
  const body = (xml.match(/<office:body>([\s\S]*)<\/office:body>/) || [])[1] || xml;
  return tidy(decode(body
    .replace(/<text:s(?: text:c="(\d+)")?\/>/g, (m, c) => ' '.repeat(Number(c) || 1))
    .replace(/<text:tab\/>/g, '\t')
    .replace(/<text:line-break\/>/g, '\n')
    .replace(/<\/text:(p|h)>/g, '\n')
    .replace(/<\/table:table-cell>/g, '\t')
    .replace(/<\/table:table-row>/g, '\n')
    .replace(/<draw:page [^>]*draw:name="([^"]*)"[^>]*>/g, '\n--- $1 ---\n')
    .replace(/<text:tracked-changes>[\s\S]*?<\/text:tracked-changes>/g, '')
    .replace(/<office:annotation>[\s\S]*?<\/office:annotation>/g, '')
    .replace(/<[^>]+>/g, '')).replace(/\t+\n/g, '\n'));
}

/** RTF → text (control words removed, \'xx and \uN decoded) */
function rtfText(buf) {
  let s = buf.toString('latin1');
  s = s.replace(/\{\\\*[^{}]*(\{[^{}]*\}[^{}]*)*\}/g, '');
  s = s.replace(/\{\\(fonttbl|colortbl|stylesheet|info|listtable|listoverridetable|rsidtbl|generator)([^{}]|\{[^{}]*(\{[^{}]*\}[^{}]*)*\})*\}/g, '');
  s = s.replace(/\\'([0-9a-f]{2})/gi, (m, h) => Buffer.from([parseInt(h, 16)]).toString('latin1'));
  s = s.replace(/\\u(-?\d+)\??/g, (m, n) => String.fromCharCode((Number(n) + 65536) % 65536));
  s = s.replace(/\\(par|line|row)\b ?/g, '\n').replace(/\\(tab|cell)\b ?/g, '\t').replace(/\\[a-z]+-?\d* ?/gi, '').replace(/\\([{}\\])/g, '$1').replace(/[{}]/g, '');
  return tidy(s);
}

/** HTML → text */
const htmlText = (s) => tidy(decode(String(s).replace(/<(script|style|head|template)[\s\S]*?<\/\1>/gi, '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|h[1-6]|li|tr|section|article|header|footer|blockquote|pre)>/gi, '\n').replace(/<\/t[dh]>/gi, '\t').replace(/<[^>]+>/g, '')));

/** Simple e-mail text without mailparser: headers Subject/From + body (quoted-printable/base64 parts not decoded) */
function emlFallback(buf) {
  const s = bytesToString(buf);
  const i = s.search(/\r?\n\r?\n/);
  const head = i > 0 ? s.slice(0, i) : '';
  const pick = (h) => ((head.match(new RegExp(`^${h}:\\s*(.*)$`, 'mi')) || [])[1] || '').trim();
  return tidy(`${pick('Subject')}\n${pick('From')}\n\n${i > 0 ? s.slice(i) : s}`);
}

const DOCX = /^(docx|docm|dotx|dotm)$/;
const XLSX = /^(xlsx|xlsm|xltx|xltm)$/;
const PPTX = /^(pptx|pptm|ppsx|ppsm|potx|potm)$/;
const ODF = /^(odt|ods|odp|ott|ots|otp|odg|otg)$/;
const FLAT = /^(fodt|fods|fodp|fodg)$/;
const PLAIN = /^(txt|md|markdown|csv|tsv|tab|json|xml|vcf|ics|prn|dif|slk)$/;
/** Can text be extracted from this extension? */
const LEGACY = /^(doc|dot|ppt|pps|pot|xls|xlt)$/;
const IWORK = /^(pages|key|numbers)$/;
const canText = (ext) => LEGACY.test(ext) || IWORK.test(ext) || DOCX.test(ext) || XLSX.test(ext) || PPTX.test(ext) || ODF.test(ext) || FLAT.test(ext) || PLAIN.test(ext) || /^(rtf|html|htm|xhtml|pdf|eml)$/.test(ext);

async function pdfText(buf, unpdf) {
  if (!unpdf) return '';
  const pdf = await unpdf.getDocumentProxy(new Uint8Array(buf));
  try {
    const { text } = await unpdf.extractText(pdf, { mergePages: false });
    return tidy((Array.isArray(text) ? text : [text]).map((t, i) => `--- ${i + 1} ---\n${String(t).trim()}`).join('\n\n'));
  } finally { try { await pdf.destroy(); } catch { /* egal */ } }
}

/** Plain text of a file. libs: { unpdf, mailparser } (optional) */
async function text(buf, ext, libs = {}) {
  const e = String(ext || '').toLowerCase();
  if (DOCX.test(e) || XLSX.test(e) || PPTX.test(e) || ODF.test(e)) {
    if (!isZip(buf)) return '';
    const z = readZip(buf, { maxSize: 64 * 1048576 });
    if (DOCX.test(e)) return clip(docxText(z));
    if (XLSX.test(e)) return clip(xlsxText(z));
    if (PPTX.test(e)) return clip(pptxText(z));
    return clip(odfXml(zipText(z, 'content.xml')));
  }
  if (FLAT.test(e)) return clip(odfXml(bytesToString(buf)));
  if (LEGACY.test(e)) {
    if (!cfb.isCfb(buf)) return '';
    return clip(/^(doc|dot)$/.test(e) ? cfb.docText(buf) : /^(xls|xlt)$/.test(e) ? tidy(cfb.xlsText(buf).text) : cfb.pptText(buf).text);
  }
  if (IWORK.test(e)) {
    if (!isZip(buf)) return '';
    const z = readZip(buf, { maxSize: 128 * 1048576 });
    const t = iwork.iwaText(z);
    if (t) return clip(t);
    const pv = iwork.previewOf(z);
    return pv.pdf ? clip(await pdfText(pv.pdf, libs.unpdf)) : '';
  }
  if (e === 'rtf') return clip(rtfText(buf));
  if (/^(html|htm|xhtml)$/.test(e)) return clip(htmlText(bytesToString(buf)));
  if (PLAIN.test(e)) return clip(tidy(bytesToString(buf)));
  if (e === 'pdf') return clip(await pdfText(buf, libs.unpdf));
  if (e === 'eml') {
    if (libs.mailparser && libs.mailparser.simpleParser) {
      const m = await libs.mailparser.simpleParser(buf);
      return clip(tidy(`${m.subject || ''}\n${(m.from && m.from.text) || ''}\n\n${m.text || (m.html ? htmlText(m.html) : '')}`));
    }
    return clip(emlFallback(buf));
  }
  return '';
}

const words = (s) => (String(s).match(/[\p{L}\p{N}][\p{L}\p{N}'’_-]*/gu) || []).length;
const tag = (xml, name) => { const m = xml.match(new RegExp(`<${name}(?:\\s[^>]*)?>([\\s\\S]*?)</${name}>`)); return m ? decode(m[1]).trim() : undefined; };
const date = (s) => { if (!s) return undefined; const t = Date.parse(s); return Number.isFinite(t) ? t : undefined; };
const clean = (o) => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined && v !== '' && !(typeof v === 'number' && !Number.isFinite(v))));

/** Facts from the file's own metadata (docProps/core.xml + app.xml, ODF meta.xml, PDF info) */
async function info(buf, ext, libs = {}) {
  const e = String(ext || '').toLowerCase();
  if ((DOCX.test(e) || XLSX.test(e) || PPTX.test(e)) && isZip(buf)) {
    const z = readZip(buf, { maxSize: 64 * 1048576 });
    const core = zipText(z, 'docProps/core.xml'); const app = zipText(z, 'docProps/app.xml');
    const n = (t) => { const v = tag(app, t); return v != null && v !== '' ? Number(v) : undefined; };
    const out = { title: tag(core, 'dc:title'), author: tag(core, 'dc:creator'), lastModifiedBy: tag(core, 'cp:lastModifiedBy'), created: date(tag(core, 'dcterms:created')), modified: date(tag(core, 'dcterms:modified')), application: tag(app, 'Application') };
    if (DOCX.test(e)) { out.pages = n('Pages'); out.words = n('Words'); out.characters = n('Characters'); if (out.words == null) out.words = words(docxText(z)); }
    if (XLSX.test(e)) out.sheets = [...zipText(z, 'xl/workbook.xml').matchAll(/<sheet [^>]*\/?>/g)].length || [...z.keys()].filter((x) => /^xl\/worksheets\/sheet\d+\.xml$/.test(x)).length;
    if (PPTX.test(e)) { out.slides = n('Slides') || [...z.keys()].filter((x) => /^ppt\/slides\/slide\d+\.xml$/.test(x)).length; out.words = n('Words'); out.notes = n('Notes'); if (!out.words) out.words = words(pptxText(z).replace(/^--- \d+ ---$/gm, '')); } // some writers store Words=0
    return clean(out);
  }
  if ((ODF.test(e) && isZip(buf)) || FLAT.test(e)) {
    let xml; let z = null;
    if (FLAT.test(e)) xml = bytesToString(buf); else { z = readZip(buf, { maxSize: 64 * 1048576 }); xml = zipText(z, 'meta.xml'); }
    const stat = (xml.match(/<meta:document-statistic([^>]*)\/?>/) || [])[1] || '';
    const sa = (k) => { const m = stat.match(new RegExp(`meta:${k}="(\\d+)"`)); return m ? Number(m[1]) : undefined; };
    const out = { title: tag(xml, 'dc:title'), author: tag(xml, 'meta:initial-creator') || tag(xml, 'dc:creator'), lastModifiedBy: tag(xml, 'dc:creator'), created: date(tag(xml, 'meta:creation-date')), modified: date(tag(xml, 'dc:date')), application: tag(xml, 'meta:generator') };
    if (/^(odt|ott|fodt)$/.test(e)) { out.pages = sa('page-count'); out.words = sa('word-count'); out.characters = sa('character-count'); }
    if (/^(ods|ots|fods)$/.test(e)) { out.sheets = sa('table-count'); if (!out.sheets) { const c = z ? zipText(z, 'content.xml') : xml; out.sheets = (c.match(/<table:table /g) || []).length; } }
    if (/^(odp|otp|fodp|odg|otg|fodg)$/.test(e)) { const c = z ? zipText(z, 'content.xml') : xml; out.slides = (c.match(/<draw:page /g) || []).length; }
    if (out.words == null && /^(odt|ott|fodt)$/.test(e)) out.words = words(z ? odfXml(zipText(z, 'content.xml')) : odfXml(xml));
    return clean(out);
  }
  if (/^(ppt|pps|pot)$/.test(e) && cfb.isCfb(buf)) { const r = cfb.pptText(buf); return clean({ slides: r.slides, words: words(r.text) }); }
  if (/^(xls|xlt)$/.test(e) && cfb.isCfb(buf)) { try { return clean({ sheets: cfb.xlsText(buf).sheets }); } catch { return {}; } }
  if (e === 'pdf' && libs.unpdf) {
    const pdf = await libs.unpdf.getDocumentProxy(new Uint8Array(buf));
    try {
      const out = { pages: pdf.numPages };
      try { const m = await pdf.getMetadata(); const i = (m && m.info) || {}; out.title = i.Title; out.author = i.Author; out.application = i.Creator || i.Producer; out.created = pdfDate(i.CreationDate); out.modified = pdfDate(i.ModDate); } catch { /* no metadata */ }
      return clean(out);
    } finally { try { await pdf.destroy(); } catch { /* egal */ } }
  }
  if (canText(e)) { const t = await text(buf, e, libs); return clean({ words: words(t), characters: t.length, lines: t ? t.split('\n').length : 0 }); }
  return {};
}
/** PDF date "D:20240131120000+01'00'" → ms */
function pdfDate(s) {
  const m = /^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?([Zz+-])?(\d{2})?'?(\d{2})?/.exec(String(s || ''));
  if (!m) return undefined;
  const t = Date.UTC(+m[1], (+m[2] || 1) - 1, +m[3] || 1, +m[4] || 0, +m[5] || 0, +m[6] || 0);
  const off = m[7] === '+' || m[7] === '-' ? (m[7] === '+' ? 1 : -1) * ((+m[8] || 0) * 60 + (+m[9] || 0)) * 60000 : 0;
  return t - off;
}

/** Embedded thumbnail → { type, data } or null (only PNG/JPEG/GIF – WMF/EMF are skipped) */
function thumb(buf, ext) {
  const e = String(ext || '').toLowerCase();
  if (IWORK.test(e) && isZip(buf)) { const pv = iwork.previewOf(readZip(buf, { maxSize: 32 * 1048576 })); return pv.image || null; }
  if (!isZip(buf) || !(DOCX.test(e) || XLSX.test(e) || PPTX.test(e) || ODF.test(e) || /^(vsdx|vsdm|vstx|xlsb|epub)$/.test(e))) return null;
  const z = readZip(buf, { maxSize: 16 * 1048576 });
  for (const [name, type] of [['Thumbnails/thumbnail.png', 'image/png'], ['docProps/thumbnail.jpeg', 'image/jpeg'], ['docProps/thumbnail.jpg', 'image/jpeg'], ['docProps/thumbnail.png', 'image/png'], ['docProps/thumbnail.gif', 'image/gif']]) {
    const x = z.get(name);
    if (x && x.size > 0 && x.size < 16 * 1048576) return { type, data: x.read() };
  }
  return null;
}
/** Cheap check: does the zip have a thumbnail entry? */
function hasThumb(buf, ext) { try { return !!thumb(buf, ext); } catch { return false; } }

/** Snippet around the first hit of q (case-insensitive) */
function snippet(textValue, q, radius = 80) {
  const t = String(textValue || '');
  const terms = String(q || '').toLowerCase().split(/\s+/).filter(Boolean);
  const low = t.toLowerCase();
  let at = -1; let len = 0;
  for (const term of terms) { const i = low.indexOf(term); if (i >= 0 && (at < 0 || i < at)) { at = i; len = term.length; } }
  if (at < 0) return t.slice(0, radius * 2).replace(/\s+/g, ' ').trim();
  const start = Math.max(0, at - radius); const end = Math.min(t.length, at + len + radius);
  return `${start > 0 ? '…' : ''}${t.slice(start, end).replace(/\s+/g, ' ').trim()}${end < t.length ? '…' : ''}`;
}

module.exports = { text, info, thumb, hasThumb, snippet, canText, words, decode, rtfText, htmlText, odfXml, docxText, xlsxText, pptxText, bytesToString, pdfDate, MAX_TEXT };
