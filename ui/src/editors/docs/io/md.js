/*
 * Markdown ↔ HTML. Reading: marked (GFM: tables, task lists, strikethrough, autolinks). Writing: turndown + GFM
 * plugin with rules for our own nodes (check lists, footnotes as [^n], equations as $…$, page breaks, call-outs,
 * table of contents, tracked changes shown as the final text).
 */
import { marked } from 'marked';
import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';

export function mdToHtml(md) {
  // footnote definitions "[^1]: text" → our footnote nodes
  const defs = {};
  const body = String(md || '').replace(/^\[\^([^\]]+)\]:\s*(.+)$/gm, (m, id, t) => { defs[id] = t.trim(); return ''; });
  let html = marked.parse(body, { gfm: true, breaks: false, async: false });
  html = html.replace(/\[\^([^\]]+)\]/g, (m, id) => (defs[id] != null ? `<sup data-type="footnote" data-text="${defs[id].replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}">${id}</sup>` : m));
  // $…$ equations (inline, not across lines, not money like "$5 and $6")
  html = html.replace(/(^|[^\\$\w])\$([^$\n<>]*[\\^_{}][^$\n<>]*)\$(?!\d)/g, (m, pre, tex) => `${pre}<span data-type="equation" data-latex="${tex.replace(/"/g, '&quot;')}"></span>`);
  // GFM task lists → tiptap task lists
  html = html.replace(/<ul>\s*((?:<li>\s*<input[^>]*type="checkbox"[^>]*>[\s\S]*?<\/li>\s*)+)<\/ul>/g, (m, items) => `<ul data-type="taskList">${items.replace(/<li>\s*<input([^>]*)>\s*/g, (x, attrs) => `<li data-type="taskItem" data-checked="${/checked/.test(attrs) ? 'true' : 'false'}">`)}</ul>`);
  // explicit page breaks written by us
  html = html.replace(/<div style="page-break-after:\s*always;?"><\/div>/g, '<div data-type="page-break"></div>');
  return html;
}

let td = null;
function service() {
  if (td) return td;
  td = new TurndownService({ headingStyle: 'atx', codeBlockStyle: 'fenced', bulletListMarker: '-', emDelimiter: '*', strongDelimiter: '**', hr: '---', linkStyle: 'inlined' });
  td.use(gfm);
  // compact list markers ("- item", "1. item") – added first, so the more specific rules below win
  td.addRule('listItem', {
    filter: 'li',
    replacement: (content, node, options) => {
      let prefix = `${options.bulletListMarker} `;
      const parent = node.parentNode;
      if (parent && parent.nodeName === 'OL') { const start = parent.getAttribute('start'); const index = Array.prototype.indexOf.call(parent.children, node); prefix = `${(start ? Number(start) : 1) + index}. `; }
      const body = content.replace(/^\n+/, '').replace(/\n+$/, '\n').replace(/\n/gm, `\n${' '.repeat(prefix.length)}`);
      return prefix + body + (node.nextSibling && !/\n$/.test(body) ? '\n' : '');
    },
  });
  td.addRule('taskItem', { filter: (n) => n.nodeName === 'LI' && n.getAttribute('data-type') === 'taskItem', replacement: (content, node) => `- [${node.getAttribute('data-checked') === 'true' ? 'x' : ' '}] ${content.replace(/^\s+/, '').replace(/\n+$/, '')}\n` });
  td.addRule('taskList', { filter: (n) => n.nodeName === 'UL' && n.getAttribute('data-type') === 'taskList', replacement: (content) => `\n\n${content}\n\n` });
  td.addRule('checkbox', { filter: (n) => n.nodeName === 'INPUT', replacement: () => '' });
  td.addRule('footnote', {
    filter: (n) => n.nodeName === 'SUP' && n.getAttribute('data-type') === 'footnote',
    replacement: (content, node) => { td.__notes = td.__notes || []; td.__notes.push(node.getAttribute('data-text') || ''); return `[^${td.__notes.length}]`; },
  });
  td.addRule('footnoteSection', { filter: (n) => n.nodeName === 'SECTION' && /of-docs-footnotes/.test(n.getAttribute('class') || ''), replacement: () => '' });
  td.addRule('equation', { filter: (n) => n.nodeName === 'SPAN' && n.getAttribute('data-type') === 'equation', replacement: (c, node) => `$${node.getAttribute('data-latex') || ''}$` });
  td.addRule('pageBreak', { filter: (n) => n.nodeName === 'DIV' && n.getAttribute('data-type') === 'page-break', replacement: () => '\n\n<div style="page-break-after:always"></div>\n\n' });
  td.addRule('toc', {
    filter: (n) => n.nodeName === 'NAV' && n.getAttribute('data-type') === 'toc',
    replacement: (c, node) => { let es = []; try { es = JSON.parse(node.getAttribute('data-entries') || '[]'); } catch { /* none */ } const title = node.getAttribute('data-title'); return `\n\n${title ? `**${title}**\n\n` : ''}${es.map((e) => `${'  '.repeat(Math.max(0, (e.level || 1) - 1))}- ${e.text}`).join('\n')}\n\n`; },
  });
  td.addRule('callout', { filter: (n) => n.nodeName === 'DIV' && n.getAttribute('data-type') === 'callout', replacement: (content) => `\n\n${content.trim().split('\n').map((l) => `> ${l}`).join('\n')}\n\n` });
  td.addRule('deletion', { filter: (n) => n.nodeName === 'DEL' && /of-docs-del/.test(n.getAttribute('class') || ''), replacement: () => '' });
  td.addRule('underline', { filter: ['u'], replacement: (c) => `<u>${c}</u>` });
  td.addRule('mark', { filter: ['mark'], replacement: (c) => `==${c}==` });
  td.addRule('subsup', { filter: (n) => (n.nodeName === 'SUB' || n.nodeName === 'SUP') && !n.getAttribute('data-type'), replacement: (c, node) => `<${node.nodeName.toLowerCase()}>${c}</${node.nodeName.toLowerCase()}>` });
  td.addRule('spanPassthrough', { filter: (n) => (n.nodeName === 'SPAN' || n.nodeName === 'INS') && !n.getAttribute('data-type'), replacement: (c) => c });
  return td;
}

export function htmlToMd(html) {
  const s = service();
  s.__notes = [];
  let md = s.turndown(String(html || ''));
  if (s.__notes.length) md += `\n\n${s.__notes.map((t, k) => `[^${k + 1}]: ${t}`).join('\n')}\n`;
  return md.replace(/\n{3,}/g, '\n\n').trim() + '\n';
}
