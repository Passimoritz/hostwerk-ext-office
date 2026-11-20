/*
 * The docs chunk keeps a document as ProseMirror/tiptap JSON (the editor's own format) plus a small model around it.
 * Every reader produces this, every writer consumes it – pure data, so all converters run in Node (unit tests) too.
 *
 *   model = {
 *     doc:      { type: 'doc', content: [...] }            tiptap JSON (node/mark names see SCHEMA below)
 *     page:     { size, width, height (mm, portrait), orientation, margins: { top, right, bottom, left } (mm), columns, color }
 *     header / footer: { text, align }                     text may contain {page} {pages} {date} {title}
 *     titlePage: bool                                      no header/footer on the first page
 *     comments: [{ id, author, date, text, resolved, replies: [{ author, date, text }] }]
 *     meta:     { title, author, subject, keywords, created, modified }
 *     defaults: { font, fontSize (pt), lineHeight, spaceAfter (pt) }
 *     fidelity: ['macros', 'smartart', …]                  what the source had that we cannot fully keep
 *     keep:     { … }                                      original parts to round-trip (docm macros …)
 *   }
 *
 * SCHEMA
 *   blocks: paragraph {textAlign,lineHeight,spaceBefore,spaceAfter,indent,firstLine,pstyle} · heading {level,+same}
 *           blockquote · codeBlock {language} · bulletList {listStyle} · orderedList {start,listStyle} · listItem
 *           taskList · taskItem {checked} · horizontalRule · pageBreak · toc {entries,title} · callout {kind}
 *           table · tableRow · tableCell|tableHeader {colspan,rowspan,colwidth,backgroundColor}
 *   inline: text · hardBreak · image {src,alt,title,width,height,align,wrap} · footnote {text} · equation {latex}
 *   marks:  bold italic underline strike code subscript superscript link {href} highlight {color}
 *           textStyle {color,fontFamily,fontSize} comment {id} insertion {id,author,date} deletion {id,author,date}
 */

export const PAGE_SIZES = {
  A4: { width: 210, height: 297 },
  A5: { width: 148, height: 210 },
  A3: { width: 297, height: 420 },
  Letter: { width: 215.9, height: 279.4 },
  Legal: { width: 215.9, height: 355.6 },
};

/**
 * Word's "single" line of a font = its ascent + descent + line gap (not 1 × the font size like CSS line-height: 1).
 * A Word line spacing "multiple m" is m × this factor × the font size – the factor of the common fonts:
 */
const LINE_FACTOR = { calibri: 1.2207, 'calibri light': 1.2207, carlito: 1.2207, aptos: 1.2234, 'aptos display': 1.2234, 'aptos narrow': 1.2234, arial: 1.149, 'liberation sans': 1.149, helvetica: 1.149, 'times new roman': 1.149, 'liberation serif': 1.149, cambria: 1.1724, 'courier new': 1.133, consolas: 1.1709, georgia: 1.1362, verdana: 1.2153, tahoma: 1.2075, 'segoe ui': 1.3301, 'trebuchet ms': 1.1626, garamond: 1.1259, 'century gothic': 1.2251, 'palatino linotype': 1.3271, 'franklin gothic medium': 1.1338, 'lucida sans': 1.1612, candara: 1.2207, constantia: 1.2207, corbel: 1.2207, 'comic sans ms': 1.3936, impact: 1.2192 };
export function lineFactor(font) { const f = String(font || '').split(',')[0].replace(/["']/g, '').trim().toLowerCase(); return LINE_FACTOR[f] || 1.17; }

export const MM_PX = 96 / 25.4;
export const PT_PX = 96 / 72;
export const TWIP_MM = 25.4 / 1440;
export const mmToTwip = (mm) => Math.round((mm / 25.4) * 1440);
export const twipToMm = (tw) => Math.round(((Number(tw) || 0) / 1440) * 25.4 * 10) / 10;

export function defaultModel(settings = {}) {
  const d = (settings && settings.docs) || {};
  const g = (settings && settings.general) || {};
  const size = PAGE_SIZES[d.pageSize] ? d.pageSize : 'A4';
  const m = d.margins || {};
  return {
    doc: { type: 'doc', content: [{ type: 'paragraph' }] },
    page: {
      size, ...PAGE_SIZES[size], orientation: d.orientation === 'landscape' ? 'landscape' : 'portrait',
      margins: { top: num(m.top, 25), right: num(m.right, 25), bottom: num(m.bottom, 20), left: num(m.left, 25) },
      columns: 1, color: null,
    },
    header: { text: '', align: 'center' },
    footer: { text: '', align: 'center' },
    titlePage: false,
    comments: [],
    meta: { title: '', author: g.author || '', subject: '', keywords: '' },
    defaults: { font: d.font || 'Calibri', fontSize: num(d.fontSize, 11), lineHeight: num(d.lineHeight, 1.15), spaceAfter: num(d.paragraphSpacing, 8) },
    fidelity: [],
    keep: {},
  };
}
const num = (v, def) => (Number.isFinite(Number(v)) && v !== '' && v !== null ? Number(v) : def);

/** page width/height in mm as shown (orientation applied) */
export function pageBox(page) {
  const w = page.width || 210; const h = page.height || 297;
  return page.orientation === 'landscape' ? { width: Math.max(w, h), height: Math.min(w, h) } : { width: Math.min(w, h), height: Math.max(w, h) };
}
/** find the named size for width/height in mm (either orientation) */
export function sizeName(w, h) {
  const a = Math.min(w, h); const b = Math.max(w, h);
  for (const [k, v] of Object.entries(PAGE_SIZES)) if (Math.abs(v.width - a) < 2 && Math.abs(v.height - b) < 2) return k;
  return 'custom';
}

/* ---------------- builders ---------------- */
export const txt = (text, marks) => (text ? (marks && marks.length ? { type: 'text', text, marks } : { type: 'text', text }) : null);
const clean = (arr) => (arr || []).flat(Infinity).filter(Boolean).map((x) => (typeof x === 'string' ? { type: 'text', text: x } : x)).filter((x) => x.type !== 'text' || x.text);
export function para(content, attrs) { const c = clean(Array.isArray(content) ? content : [content]); const n = { type: 'paragraph' }; if (attrs && Object.keys(attrs).length) n.attrs = attrs; if (c.length) n.content = c; return n; }
export function heading(level, content, attrs) { const n = para(content, attrs); n.type = 'heading'; n.attrs = { level, ...(attrs || {}) }; return n; }
export const mark = (type, attrs) => (attrs ? { type, attrs } : { type });
export const b = (t) => txt(t, [mark('bold')]);
export const i = (t) => txt(t, [mark('italic')]);
export function bullets(items, ordered = false, attrs) {
  return { type: ordered ? 'orderedList' : 'bulletList', ...(attrs ? { attrs } : {}), content: items.map((it) => ({ type: 'listItem', content: Array.isArray(it) && it.length && it[0] && it[0].type && it[0].type !== 'text' ? it : [para(it)] })) };
}
export function tasks(items) { return { type: 'taskList', content: items.map(([checked, t]) => ({ type: 'taskItem', attrs: { checked: !!checked }, content: [para(t)] })) }; }
export function cell(content, attrs = {}, header = false) {
  const blocks = (Array.isArray(content) ? content : [content]).filter(Boolean);
  const c = blocks.length && blocks[0].type && blocks[0].type !== 'text' ? blocks : [para(blocks)];
  return { type: header ? 'tableHeader' : 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: null, ...attrs }, content: c };
}
export function table(rows, { header = true, widths = null } = {}) {
  return {
    type: 'table',
    content: rows.map((r, ri) => ({ type: 'tableRow', content: r.map((c, ci) => (c && c.type && (c.type === 'tableCell' || c.type === 'tableHeader') ? c : cell(c, widths && widths[ci] ? { colwidth: [widths[ci]] } : {}, header && ri === 0))) })),
  };
}

/* ---------------- walking ---------------- */
export function walk(node, fn, parent = null, depth = 0) {
  if (!node) return;
  if (fn(node, parent, depth) === false) return;
  for (const c of node.content || []) walk(c, fn, node, depth + 1);
}

export const BLOCKS = new Set(['sectionBreak', 'controlBlock', 'paragraph', 'heading', 'blockquote', 'codeBlock', 'bulletList', 'orderedList', 'listItem', 'taskList', 'taskItem', 'horizontalRule', 'pageBreak', 'toc', 'callout', 'table', 'tableRow', 'tableCell', 'tableHeader']);
export const TEXTBLOCKS = new Set(['paragraph', 'heading', 'codeBlock']);

/** plain text of an inline run (text, hard breaks, footnotes as [n] markers are left out) */
export function inlineText(node, { deleted = false } = {}) {
  let s = '';
  for (const c of node.content || []) {
    if (c.type === 'text') { if (!deleted && (c.marks || []).some((m) => m.type === 'deletion')) continue; s += c.text; } else if (c.type === 'hardBreak') s += '\n';
    else if (c.type === 'equation') s += (c.attrs && c.attrs.latex) || '';
    else if (c.type === 'field') s += (c.attrs && c.attrs.value) || '';
    else if (c.type === 'contentControl') s += c.attrs && c.attrs.kind === 'checkbox' ? (c.attrs.checked ? '☒' : '☐') : (c.attrs && c.attrs.value) || '';
    else if (c.content) s += inlineText(c, { deleted });
  }
  return s;
}

/** plain text of a whole document – paragraphs separated by newlines, lists with "• " / "1. ", tables tab-separated */
export function docText(doc) {
  const out = [];
  const block = (n, prefix = '') => {
    switch (n.type) {
      case 'paragraph': case 'heading': case 'codeBlock': out.push(prefix + inlineText(n)); break;
      case 'bulletList': case 'taskList': for (const li of n.content || []) listItem(li, n.type === 'taskList' ? ((li.attrs && li.attrs.checked) ? '[x] ' : '[ ] ') : '• ', prefix); break;
      case 'orderedList': { let k = (n.attrs && n.attrs.start) || 1; for (const li of n.content || []) listItem(li, `${k++}. `, prefix); break; }
      case 'table': for (const r of n.content || []) out.push(prefix + (r.content || []).map((c) => (c.content || []).map((p) => inlineText(p)).join(' ')).join('\t')); break;
      case 'horizontalRule': out.push(prefix + '———'); break;
      case 'pageBreak': out.push('\f'); break;
      case 'sectionBreak': if (!n.attrs || n.attrs.type !== 'continuous') out.push('\f'); break;
      case 'toc': for (const e of (n.attrs && n.attrs.entries) || []) out.push(`${'  '.repeat(Math.max(0, (e.level || 1) - 1))}${e.text}`); break;
      default: for (const c of n.content || []) block(c, prefix);
    }
  };
  const listItem = (li, bullet, prefix) => {
    (li.content || []).forEach((c, k) => (k === 0 && TEXTBLOCKS.has(c.type) ? out.push(prefix + bullet + inlineText(c)) : block(c, prefix + '   ')));
  };
  for (const n of (doc && doc.content) || []) block(n);
  const notes = [];
  walk(doc, (n) => { if (n.type === 'footnote') notes.push(n.attrs.text || ''); });
  if (notes.length) { out.push('', '———'); notes.forEach((t, k) => out.push(`[${k + 1}] ${t}`)); }
  return out.join('\n').replace(/\n\f\n/g, '\n\f');
}

/** headings for TOC / navigation: [{ level, text, index }] (index = n-th heading in document order) */
export function headingsOf(doc, maxLevel = 6) {
  const out = [];
  let k = 0;
  walk(doc, (n) => {
    if (n.type === 'heading') { const level = (n.attrs && n.attrs.level) || 1; if (level <= maxLevel) out.push({ level, text: inlineText(n).trim(), index: k }); k++; return false; }
    if (n.type === 'paragraph' && n.attrs && n.attrs.pstyle === 'title') return false;
    return undefined;
  });
  return out.filter((h) => h.text);
}

/** word / character / paragraph counts of a document */
export function stats(doc) {
  const text = docText(doc);
  const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-_.@]*/gu) || []).length;
  let paragraphs = 0;
  walk(doc, (n) => { if (TEXTBLOCKS.has(n.type) && inlineText(n).trim()) paragraphs++; });
  return { words, chars: text.replace(/[\n\f]/g, '').length, charsNoSpaces: text.replace(/\s/g, '').length, paragraphs };
}

/** footnotes in document order: [text] */
export function footnotesOf(doc) { const out = []; walk(doc, (n) => { if (n.type === 'footnote') out.push((n.attrs && n.attrs.text) || ''); }); return out; }

/** ids of comments that are still anchored somewhere in the document */
export function commentIdsIn(doc) {
  const ids = new Set();
  walk(doc, (n) => { for (const m of n.marks || []) if (m.type === 'comment' && m.attrs && m.attrs.id) ids.add(String(m.attrs.id)); });
  return ids;
}

/** merge neighbouring text nodes with equal marks, drop empty ones (readers produce many small runs) */
export function normalizeInline(content) {
  const out = [];
  for (const c of content || []) {
    if (!c) continue;
    if (c.type === 'text') {
      if (!c.text) continue;
      const prev = out[out.length - 1];
      if (prev && prev.type === 'text' && sameMarks(prev.marks, c.marks)) { prev.text += c.text; continue; }
      const n = { type: 'text', text: c.text };
      if (c.marks && c.marks.length) n.marks = c.marks;
      out.push(n);
    } else out.push(c);
  }
  return out;
}
export function sameMarks(a, b) {
  const x = a || []; const y = b || [];
  if (x.length !== y.length) return false;
  const key = (m) => m.type + JSON.stringify(m.attrs || {});
  const s = new Set(x.map(key));
  return y.every((m) => s.has(key(m)));
}

/** make a document valid for the schema: list items/cells start with a paragraph, no empty containers … */
export function fixDoc(doc) {
  const fixBlock = (n) => {
    if (!n || !n.type) return null;
    if (n.type === 'text') return n;
    if (n.content) n.content = n.content.map(fixBlock).filter(Boolean);
    if (TEXTBLOCKS.has(n.type)) {
      if (n.content) n.content = n.type === 'codeBlock' ? n.content.filter((c) => c.type === 'text').map((c) => ({ type: 'text', text: c.text })) : normalizeInline(n.content);
      if (n.content && !n.content.length) delete n.content;
      return n;
    }
    if (['listItem', 'taskItem'].includes(n.type)) {
      if (!n.content || !n.content.length || n.content[0].type !== 'paragraph') n.content = [{ type: 'paragraph' }, ...(n.content || [])];
    }
    if (['tableCell', 'tableHeader', 'blockquote', 'callout', 'controlBlock'].includes(n.type) && (!n.content || !n.content.length)) n.content = [{ type: 'paragraph' }];
    if (['bulletList', 'orderedList', 'taskList', 'table', 'tableRow'].includes(n.type) && (!n.content || !n.content.length)) return null;
    if (n.type === 'table') {
      // every row needs the same number of columns – pad short rows (merged layouts from other apps)
      const width = (r) => (r.content || []).reduce((s, c) => s + ((c.attrs && c.attrs.colspan) || 1), 0);
      const rows = n.content.filter((r) => r.type === 'tableRow');
      const max = Math.max(...rows.map(width));
      for (const r of rows) for (let k = width(r); k < max; k++) r.content.push(cell(''));
    }
    return n;
  };
  const out = fixBlock(doc) || { type: 'doc' };
  out.type = 'doc';
  if (!out.content || !out.content.length) out.content = [{ type: 'paragraph' }];
  // inline nodes directly in the doc (sloppy readers) → wrap
  const fixed = [];
  let run = null;
  for (const c of out.content) {
    if (['text', 'hardBreak', 'image', 'footnote', 'equation', 'field', 'contentControl', 'diagram', 'preserved'].includes(c.type)) { if (!run) { run = { type: 'paragraph', content: [] }; fixed.push(run); } run.content.push(c); } else { run = null; fixed.push(c); }
  }
  out.content = fixed;
  return out;
}

/* ---------------- small shared helpers ---------------- */
export function firstFont(f) { return String(f || '').split(',')[0].trim().replace(/^["']|["']$/g, ''); }
export function ptOf(fontSize) { const m = /([\d.]+)\s*(pt|px)?/.exec(String(fontSize || '')); if (!m) return null; return m[2] === 'px' ? Math.round((Number(m[1]) * 0.75) * 2) / 2 : Number(m[1]); }
export function hexColor(c) {
  if (!c) return null;
  const s = String(c).trim();
  let m = /^#?([0-9a-f]{6})$/i.exec(s); if (m) return m[1].toUpperCase();
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(s); if (m) return (m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  m = /rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/i.exec(s); if (m) return [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('').toUpperCase();
  const named = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', cyan: '00FFFF', magenta: 'FF00FF', gray: '808080', grey: '808080', orange: 'FFA500', darkblue: '00008B', darkred: '8B0000', darkgreen: '006400', lightgray: 'D3D3D3' };
  return named[s.toLowerCase()] || null;
}
export const uid = (p = 'c') => `${p}${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
export function bytesToBase64(u8) {
  if (typeof Buffer !== 'undefined') return Buffer.from(u8).toString('base64');
  let s = ''; const CH = 0x8000;
  for (let k = 0; k < u8.length; k += CH) s += String.fromCharCode.apply(null, u8.subarray(k, k + CH));
  return btoa(s);
}
export function base64ToBytes(b64) {
  if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(b64, 'base64'));
  const s = atob(b64); const u = new Uint8Array(s.length); for (let k = 0; k < s.length; k++) u[k] = s.charCodeAt(k); return u;
}
export function mimeOf(name) {
  const e = String(name).toLowerCase().split('.').pop();
  return { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', emf: 'image/x-emf', wmf: 'image/x-wmf', tif: 'image/tiff', tiff: 'image/tiff' }[e] || 'application/octet-stream';
}
export function dataUrl(bytes, mime) { return `data:${mime};base64,${bytesToBase64(bytes)}`; }
/** { mime, bytes, ext } of a data: URL (null for other URLs) */
export function parseDataUrl(src) {
  const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(String(src || ''));
  if (!m) return null;
  const mime = m[1] || 'application/octet-stream';
  const bytes = m[2] ? base64ToBytes(m[3]) : new TextEncoder().encode(decodeURIComponent(m[3]));
  const ext = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/webp': 'webp' }[mime] || 'bin';
  return { mime, bytes, ext };
}
/** pixel size of PNG/JPEG/GIF/BMP bytes (null if unknown) */
export function imageSize(u8) {
  try {
    if (u8[0] === 0x89 && u8[1] === 0x50) return { width: (u8[16] << 24 | u8[17] << 16 | u8[18] << 8 | u8[19]) >>> 0, height: (u8[20] << 24 | u8[21] << 16 | u8[22] << 8 | u8[23]) >>> 0 };
    if (u8[0] === 0x47 && u8[1] === 0x49) return { width: u8[6] | u8[7] << 8, height: u8[8] | u8[9] << 8 };
    if (u8[0] === 0x42 && u8[1] === 0x4d) return { width: u8[18] | u8[19] << 8 | u8[20] << 16, height: Math.abs((u8[22] | u8[23] << 8 | u8[24] << 16 | u8[25] << 24) >> 0) };
    if (u8[0] === 0xff && u8[1] === 0xd8) {
      let k = 2;
      while (k < u8.length) {
        if (u8[k] !== 0xff) return null;
        const m = u8[k + 1];
        const len = u8[k + 2] << 8 | u8[k + 3];
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { height: u8[k + 5] << 8 | u8[k + 6], width: u8[k + 7] << 8 | u8[k + 8] };
        k += 2 + len;
      }
    }
  } catch { /* unknown */ }
  return null;
}
export function utf8(bytes) {
  if (typeof bytes === 'string') return bytes;
  const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (u[0] === 0xef && u[1] === 0xbb && u[2] === 0xbf) return new TextDecoder('utf-8').decode(u.subarray(3));
  if (u[0] === 0xff && u[1] === 0xfe) return new TextDecoder('utf-16le').decode(u.subarray(2));
  if (u[0] === 0xfe && u[1] === 0xff) return new TextDecoder('utf-16be').decode(u.subarray(2));
  try { return new TextDecoder('utf-8', { fatal: true }).decode(u); } catch { return new TextDecoder('windows-1252').decode(u); }
}
export const encUtf8 = (s) => new TextEncoder().encode(s);
export const escHtml = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
/** fill {page} {pages} {date} {title} in header/footer text (for HTML/print: page numbers via CSS counters) */
export function fillFields(text, { page = '', pages = '', title = '', date = '' } = {}) {
  return String(text || '').replace(/\{page\}/g, page).replace(/\{pages\}/g, pages).replace(/\{title\}/g, title).replace(/\{date\}/g, date);
}

/** comparable key of a header/footer (changed → the writers rewrite that part) */
export function hfKey(x) { return x ? JSON.stringify({ t: x.text || '', a: x.align || '', d: x.doc || null }) : ''; }
/** does a header/footer have anything to show? */
export function hfHas(x) { if (!x) return false; if (x.doc) { let any = false; walk(x.doc, (n) => { if (any) return false; if ((n.type === 'text' && n.text.trim()) || ['image', 'field', 'table', 'preserved', 'equation'].includes(n.type)) any = true; return undefined; }); return any; } return !!(x.text && x.text.trim()); }
