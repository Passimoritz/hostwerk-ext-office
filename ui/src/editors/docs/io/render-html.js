/*
 * tiptap JSON → HTML. One renderer for: HTML export, print/PDF, Markdown export (through turndown), previews and the
 * clipboard. The markup matches the parse rules of the editor extensions, so exported HTML opens again with all
 * features (footnotes, comments, TOC, equations, page breaks, call-outs, page setup in a <meta>).
 */
import { escHtml as esc, pageBox, fillFields, footnotesOf } from './pm.js';
import { equationHtml } from './equation.js';

const styleOf = (o) => Object.entries(o).filter(([, v]) => v !== null && v !== undefined && v !== '').map(([k, v]) => `${k}:${v}`).join(';');

export function blockStyle(a = {}) {
  return styleOf({
    'text-align': a.textAlign && a.textAlign !== 'left' ? a.textAlign : null,
    'line-height': a.lineHeight || null,
    'margin-top': a.spaceBefore != null ? `${a.spaceBefore}pt` : null,
    'margin-bottom': a.spaceAfter != null ? `${a.spaceAfter}pt` : null,
    'margin-left': a.indent ? `${a.indent}pt` : null,
    'text-indent': a.firstLine ? `${a.firstLine}pt` : null,
  });
}
const attr = (name, v) => (v === null || v === undefined || v === '' || v === false ? '' : ` ${name}="${esc(v === true ? 'true' : v)}"`);

function markOpen(m) {
  const a = m.attrs || {};
  switch (m.type) {
    case 'bold': return ['<strong>', '</strong>'];
    case 'italic': return ['<em>', '</em>'];
    case 'underline': return ['<u>', '</u>'];
    case 'strike': return ['<s>', '</s>'];
    case 'code': return ['<code>', '</code>'];
    case 'subscript': return ['<sub>', '</sub>'];
    case 'superscript': return ['<sup>', '</sup>'];
    case 'link': return [`<a href="${esc(a.href || '')}"${a.target ? ` target="${esc(a.target)}" rel="noopener noreferrer"` : ''}>`, '</a>'];
    case 'highlight': return [`<mark${a.color ? ` data-color="${esc(a.color)}" style="background-color:${esc(a.color)};color:inherit"` : ''}>`, '</mark>'];
    case 'textStyle': { const s = styleOf({ color: a.color, 'font-family': a.fontFamily, 'font-size': a.fontSize }); return s ? [`<span style="${esc(s)}">`, '</span>'] : ['', '']; }
    case 'comment': return [`<span class="of-docs-comment" data-comment-id="${esc(a.id)}">`, '</span>'];
    case 'insertion': return [`<ins class="of-docs-ins"${attr('data-id', a.id)}${attr('data-author', a.author)}${attr('data-date', a.date)}>`, '</ins>'];
    case 'deletion': return [`<del class="of-docs-del"${attr('data-id', a.id)}${attr('data-author', a.author)}${attr('data-date', a.date)}>`, '</del>'];
    default: return ['', ''];
  }
}

export function renderInline(nodes, ctx) {
  let s = '';
  for (const n of nodes || []) {
    if (n.type === 'text') {
      let t = esc(n.text).replace(/\n/g, '<br>');
      const ms = n.marks || [];
      for (let k = ms.length - 1; k >= 0; k--) { const [o, c] = markOpen(ms[k]); t = o + t + c; }
      s += t;
    } else if (n.type === 'hardBreak') s += '<br>';
    else if (n.type === 'image') {
      const a = n.attrs || {};
      const st = styleOf({
        width: a.width ? `${a.width}px` : null, 'max-width': '100%', height: 'auto',
        float: a.wrap === 'square' && (a.align === 'left' || a.align === 'right') ? a.align : null,
        display: a.align === 'center' ? 'block' : null, margin: a.align === 'center' ? '0 auto' : a.wrap === 'square' && a.align === 'left' ? '0 12px 8px 0' : a.wrap === 'square' && a.align === 'right' ? '0 0 8px 12px' : null,
      });
      s += `<img src="${esc(a.src || '')}"${attr('alt', a.alt)}${attr('title', a.title)}${attr('width', a.width)}${attr('data-align', a.align)}${attr('data-wrap', a.wrap)} style="${st}">`;
    } else if (n.type === 'footnote') {
      ctx.fn = (ctx.fn || 0) + 1;
      const k = ctx.fn;
      s += ctx.export
        ? `<sup data-type="footnote" data-text="${esc(n.attrs.text || '')}" class="of-docs-fn" id="fnref${k}"><a href="#fn${k}">${k}</a></sup>`
        : `<sup data-type="footnote" data-text="${esc(n.attrs.text || '')}" class="of-docs-fn">${k}</sup>`;
    } else if (n.type === 'equation') s += `<span data-type="equation" data-latex="${esc(n.attrs.latex || '')}" class="of-docs-eq">${equationHtml(n.attrs.latex || '')}</span>`;
    else if (n.content) s += renderInline(n.content, ctx);
  }
  return s;
}

function listStyleCss(ls) { return ls && ls !== 'outline' ? `list-style-type:${ls}` : ''; }

export function renderBlocks(nodes, ctx) {
  let s = '';
  for (const n of nodes || []) s += renderBlock(n, ctx);
  return s;
}
function renderBlock(n, ctx) {
  const a = n.attrs || {};
  switch (n.type) {
    case 'paragraph': {
      const st = blockStyle(a);
      const inner = renderInline(n.content, ctx);
      return `<p${a.pstyle ? ` data-style="${esc(a.pstyle)}" class="of-p-${esc(a.pstyle)}"` : ''}${st ? ` style="${st}"` : ''}>${inner || (ctx.export ? '<br>' : '')}</p>`;
    }
    case 'heading': { const l = Math.min(6, Math.max(1, a.level || 1)); const st = blockStyle(a); return `<h${l}${st ? ` style="${st}"` : ''}${ctx.ids ? ` id="h${ctx.h = (ctx.h || 0) + 1}"` : ''}>${renderInline(n.content, ctx)}</h${l}>`; }
    case 'blockquote': return `<blockquote>${renderBlocks(n.content, ctx)}</blockquote>`;
    case 'codeBlock': return `<pre><code${a.language ? ` class="language-${esc(a.language)}"` : ''}>${esc((n.content || []).map((c) => c.text || '').join(''))}</code></pre>`;
    case 'bulletList': return `<ul${attr('data-list-style', a.listStyle)}${a.listStyle ? ` style="${listStyleCss(a.listStyle)}"` : ''}>${renderBlocks(n.content, ctx)}</ul>`;
    case 'orderedList': return `<ol${a.start && a.start !== 1 ? ` start="${a.start}"` : ''}${attr('data-list-style', a.listStyle)}${a.listStyle === 'outline' ? ' class="of-list-outline"' : a.listStyle ? ` style="${listStyleCss(a.listStyle)}"` : ''}>${renderBlocks(n.content, ctx)}</ol>`;
    case 'listItem': return `<li>${renderBlocks(n.content, ctx)}</li>`;
    case 'taskList': return `<ul data-type="taskList" class="of-tasks">${renderBlocks(n.content, ctx)}</ul>`;
    case 'taskItem': return `<li data-type="taskItem" data-checked="${a.checked ? 'true' : 'false'}">${ctx.export ? `<input type="checkbox" disabled${a.checked ? ' checked' : ''}> ` : ''}${renderBlocks(n.content, ctx)}</li>`;
    case 'horizontalRule': return '<hr>';
    case 'pageBreak': return '<div data-type="page-break" class="of-docs-pagebreak" style="break-after:page;page-break-after:always"></div>';
    case 'toc': {
      const es = a.entries || [];
      ctx.tocs = (ctx.tocs || 0) + 1;
      return `<nav data-type="toc" class="of-docs-toc" data-title="${esc(a.title || '')}" data-entries="${esc(JSON.stringify(es))}"><p class="of-toc-title">${esc(a.title || '')}</p>${es.map((e) => `<p class="of-toc-l${e.level || 1}">${ctx.ids ? `<a href="#h${(e.index || 0) + 1}">${esc(e.text)}</a>` : esc(e.text)}</p>`).join('')}</nav>`;
    }
    case 'callout': return `<div data-type="callout" data-kind="${esc(a.kind || 'info')}" class="of-docs-callout of-callout-${esc(a.kind || 'info')}">${renderBlocks(n.content, ctx)}</div>`;
    case 'table': {
      // column widths from the first row
      const first = (n.content && n.content[0] && n.content[0].content) || [];
      const cols = [];
      for (const c of first) { const cw = (c.attrs && c.attrs.colwidth) || []; for (let k = 0; k < ((c.attrs && c.attrs.colspan) || 1); k++) cols.push(cw[k] || null); }
      const colgroup = cols.some(Boolean) ? `<colgroup>${cols.map((w) => `<col${w ? ` style="width:${w}px"` : ''}>`).join('')}</colgroup>` : '';
      return `<table class="of-docs-table">${colgroup}<tbody>${(n.content || []).map((r) => `<tr>${(r.content || []).map((c) => {
        const ca = c.attrs || {};
        const tag = c.type === 'tableHeader' ? 'th' : 'td';
        return `<${tag}${ca.colspan > 1 ? ` colspan="${ca.colspan}"` : ''}${ca.rowspan > 1 ? ` rowspan="${ca.rowspan}"` : ''}${ca.colwidth ? ` colwidth="${ca.colwidth.join(',')}"` : ''}${ca.backgroundColor ? ` data-background="${esc(ca.backgroundColor)}" style="background-color:${esc(ca.backgroundColor)}"` : ''}>${renderBlocks(c.content, ctx)}</${tag}>`;
      }).join('')}</tr>`).join('')}</tbody></table>`;
    }
    default: return n.content ? renderBlocks(n.content, ctx) : '';
  }
}

/** body HTML of a document (export: footnote list + links; ids: heading anchors for the TOC) */
export function docToHtml(doc, { exportMode = false, ids = false } = {}) {
  const ctx = { export: exportMode, ids };
  let html = renderBlocks((doc && doc.content) || [], ctx);
  if (exportMode) {
    const notes = footnotesOf(doc);
    if (notes.length) html += `<section class="of-docs-footnotes"><hr><ol>${notes.map((t, k) => `<li id="fn${k + 1}">${esc(t)} <a href="#fnref${k + 1}" aria-label="Back">↩</a></li>`).join('')}</ol></section>`;
  }
  return html;
}

/** CSS for exported/printed documents (paper look, Word-like defaults) */
export function documentCss(model, { print = false } = {}) {
  const d = model.defaults || {};
  const p = model.page || {};
  const box = pageBox(p);
  const m = p.margins || { top: 25, right: 25, bottom: 20, left: 25 };
  const font = d.font ? `${d.font.includes(',') ? d.font : `"${d.font}"`}, Carlito, "Segoe UI", Arial, sans-serif` : 'Calibri, Carlito, sans-serif';
  const hf = (x, where) => {
    if (!x || !x.text) return '';
    const parts = String(x.text).split(/(\{page\}|\{pages\})/).filter(Boolean).map((t) => (t === '{page}' ? 'counter(page)' : t === '{pages}' ? 'counter(pages)' : JSON.stringify(fillFields(t, { title: (model.meta && model.meta.title) || '', date: new Date().toLocaleDateString() })))).join(' ');
    const pos = `${where}-${x.align === 'left' ? 'left' : x.align === 'right' ? 'right' : 'center'}`;
    return `@${pos}{content:${parts};font:9pt ${font};color:#555}`;
  };
  return `
@page{size:${box.width}mm ${box.height}mm;margin:${m.top}mm ${m.right}mm ${m.bottom}mm ${m.left}mm;${hf(model.header, 'top')}${hf(model.footer, 'bottom')}}
${model.titlePage ? '@page:first{@top-left{content:none}@top-center{content:none}@top-right{content:none}@bottom-left{content:none}@bottom-center{content:none}@bottom-right{content:none}}' : ''}
html{background:${print ? '#fff' : '#f0f0f0'}}
body{font-family:${font};font-size:${d.fontSize || 11}pt;line-height:${d.lineHeight || 1.15};color:#000;${print ? 'margin:0;' : `max-width:${box.width}mm;margin:24px auto;padding:${m.top}mm ${m.right}mm ${m.bottom}mm ${m.left}mm;box-sizing:border-box;box-shadow:0 2px 12px rgba(0,0,0,.15);`}background:${p.color || '#fff'};${p.columns > 1 ? `column-count:${p.columns};column-gap:12mm;` : ''}-webkit-print-color-adjust:exact;print-color-adjust:exact}
p{margin:0 0 ${d.spaceAfter != null ? d.spaceAfter : 8}pt}
h1,h2,h3,h4,h5,h6{margin:12pt 0 4pt;line-height:1.2;break-after:avoid}
h1{font-size:16pt;color:#1f3864}h2{font-size:13pt;color:#2f5496}h3{font-size:12pt;color:#1f3763}h4{font-size:11pt;font-style:italic;color:#2f5496}h5{font-size:11pt;color:#2f5496}h6{font-size:11pt;color:#1f3763}
.of-p-title{font-size:28pt;line-height:1.1;margin-bottom:4pt;letter-spacing:-.5pt}
.of-p-subtitle{font-size:14pt;color:#5a5a5a;margin-bottom:12pt}
.of-p-listparagraph{margin-left:36pt}
blockquote{margin:8pt 24pt;padding-left:10pt;border-left:3px solid #b4b4b4;color:#404040;font-style:italic}
pre{font-family:Consolas,"Courier New",monospace;font-size:10pt;background:#f4f4f4;padding:6pt 8pt;border-radius:3px;white-space:pre-wrap}
code{font-family:Consolas,"Courier New",monospace;font-size:.95em}
table{border-collapse:collapse;width:100%;margin:6pt 0;break-inside:auto}
td,th{border:1px solid #999;padding:3pt 6pt;vertical-align:top}th{background:#f2f2f2;text-align:left}
td>p:last-child,th>p:last-child{margin-bottom:0}
ul.of-tasks{list-style:none;padding-left:4pt}
ol.of-list-outline{counter-reset:ofl;list-style:none}ol.of-list-outline>li{counter-increment:ofl}ol.of-list-outline>li::before{content:counters(ofl,".") ". "}
.of-docs-pagebreak{break-after:page;height:0}
.of-docs-fn{font-size:.7em}
.of-docs-footnotes{font-size:9pt;margin-top:18pt}.of-docs-footnotes hr{width:30%;margin-left:0}
.of-docs-toc{margin:0 0 12pt}.of-toc-title{font-size:14pt;font-weight:600;color:#2f5496}.of-toc-l2{margin-left:12pt}.of-toc-l3{margin-left:24pt}.of-toc-l4,.of-toc-l5,.of-toc-l6{margin-left:36pt}
.of-docs-toc a{color:inherit;text-decoration:none}
.of-docs-callout{border:1px solid #9bb7e0;background:#eef4fc;border-radius:4px;padding:6pt 10pt;margin:6pt 0}
.of-callout-warn{border-color:#e0c07a;background:#fdf6e6}.of-callout-ok{border-color:#8fcfb1;background:#ecf8f2}.of-callout-note{border-color:#c8c8c8;background:#f6f6f6}
.of-docs-eq{font-family:"Cambria Math","Cambria",serif}.of-eq-frac{display:inline-flex;flex-direction:column;vertical-align:middle;text-align:center;font-size:.9em}.of-eq-frac>span:first-child{border-bottom:1px solid}.of-eq-sqrt>span{border-top:1px solid}
ins.of-docs-ins{color:#1a5fb4;text-decoration:underline}del.of-docs-del{color:#c01c28}
img{max-width:100%}
a{color:#0563c1}
`;
}

/** a complete standalone HTML file (images are already data: URLs) */
export function standaloneHtml(model, { print = false, title } = {}) {
  const body = docToHtml(model.doc, { exportMode: true, ids: true });
  const t = title || (model.meta && model.meta.title) || 'Document';
  const meta = { page: model.page, header: model.header, footer: model.footer, titlePage: model.titlePage, comments: model.comments, defaults: model.defaults, meta: model.meta };
  const lang = (typeof document !== 'undefined' && document.documentElement.lang) || 'en';
  return `<!DOCTYPE html>
<html lang="${esc(lang)}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="Hostwerk Office">
${model.meta && model.meta.author ? `<meta name="author" content="${esc(model.meta.author)}">\n` : ''}<meta name="of-docs" content="${esc(JSON.stringify(meta))}">
<title>${esc(t)}</title>
<style>${documentCss(model, { print })}</style>
</head>
<body>
${body}
</body>
</html>`;
}
