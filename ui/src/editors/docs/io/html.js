/*
 * HTML import (html/htm/xhtml, pasted HTML, Markdown output): our own page setup comes back from the "of-docs" meta,
 * the rest is sanitised with DOMPurify (no scripts, no event handlers, no remote frames) and normalised for the
 * editor (GFM task lists, the exported footnote list, Word/Google Docs clipboard quirks). Browser only.
 */
import DOMPurify from 'dompurify';

export function readMeta(html) {
  const m = /<meta\s+name="of-docs"\s+content="([^"]*)"/i.exec(html);
  if (!m) return null;
  try { return JSON.parse(m[1].replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')); } catch { return null; }
}

export function sanitize(html) {
  return DOMPurify.sanitize(String(html || ''), {
    WHOLE_DOCUMENT: false,
    ADD_ATTR: ['colwidth', 'data-type', 'data-text', 'data-latex', 'data-entries', 'data-title', 'data-kind', 'data-comment-id', 'data-author', 'data-date', 'data-id', 'data-checked', 'data-color', 'data-background', 'data-align', 'data-wrap', 'data-style', 'data-list-style', 'start'],
    FORBID_TAGS: ['style', 'script', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'textarea', 'select', 'link', 'meta', 'base'],
    ALLOW_DATA_ATTR: true,
  });
}

/** clean HTML for the editor: { html, meta, title } */
export function importHtml(raw) {
  const src = String(raw || '');
  const meta = readMeta(src);
  const title = (/<title>([^<]*)<\/title>/i.exec(src) || [])[1] || '';
  const bodyMatch = /<body[^>]*>([\s\S]*)<\/body>/i.exec(src);
  const clean = sanitize(bodyMatch ? bodyMatch[1] : src);
  const doc = new DOMParser().parseFromString(`<div id="of-root">${clean}</div>`, 'text/html');
  const root = doc.getElementById('of-root');
  for (const s of root.querySelectorAll('section.of-docs-footnotes')) s.remove();
  // GFM-style task lists (<li><input type=checkbox>)
  for (const ul of root.querySelectorAll('ul')) {
    const items = [...ul.children].filter((x) => x.tagName === 'LI');
    if (items.length && items.every((li) => li.querySelector(':scope > input[type="checkbox"], :scope > p > input[type="checkbox"]'))) {
      ul.setAttribute('data-type', 'taskList');
      for (const li of items) { const cb = li.querySelector('input[type="checkbox"]'); li.setAttribute('data-type', 'taskItem'); li.setAttribute('data-checked', cb && cb.checked ? 'true' : 'false'); if (cb) cb.remove(); }
    }
  }
  // Word/Google clipboard: <b style="font-weight:normal" id="docs-internal-guid…"> wrapper
  for (const b of root.querySelectorAll('b[id^="docs-internal-guid"]')) { while (b.firstChild) b.parentNode.insertBefore(b.firstChild, b); b.remove(); }
  // links never open a new browser window inside Hostwerk
  for (const a of root.querySelectorAll('a[target]')) a.removeAttribute('target');
  return { html: root.innerHTML, meta, title };
}

/** XHTML needs self-closed void elements */
export function toXhtml(html) {
  return html.replace(/<!DOCTYPE html>/i, '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE html>').replace('<html ', '<html xmlns="http://www.w3.org/1999/xhtml" ')
    .replace(/<(br|hr|img|col|input|meta|link)\b([^>]*?)\s*\/?>/gi, '<$1$2 />');
}
