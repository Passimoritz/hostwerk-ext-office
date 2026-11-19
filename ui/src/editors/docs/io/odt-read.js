/*
 * OpenDocument text reader: odt / ott (zip) and fodt (flat XML). Own parser of content.xml/styles.xml:
 * paragraphs and headings with their styles (alignment, spacing, indents, line height, page breaks), spans (bold,
 * italic, underline, strike, super/sub, font, size, colour, background), lists (bullets, numbers, letters, roman,
 * outline, ☐/☑ check lists, nesting), tables (widths, merges, cell colours, header rows), images (Pictures/ or
 * embedded binary data), links, footnotes/endnotes, comments with replies and "resolved", tracked changes
 * (insertions, deletions, format changes), table of contents, sections, page layout and header/footer.
 * Frames we cannot edit (objects, charts, shapes, text boxes) and fields become `preserved` nodes.
 */
import JSZip from 'jszip';
import { parseXml, kids, kid, all, text as xtext, serialize } from '../../_shared/xmltree.js';
import { defaultModel, sizeName, normalizeInline, fixDoc, dataUrl, mimeOf, uid } from './pm.js';

/** length string → mm */
export function lenMm(v) {
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(String(v || '').trim());
  if (!m) return null;
  const n = Number(m[1]);
  return { cm: n * 10, mm: n, in: n * 25.4, pt: (n * 25.4) / 72, pc: (n * 25.4) / 6, px: (n * 25.4) / 96 }[m[2] || 'cm'];
}
export const lenPt = (v) => { const mm = lenMm(v); return mm == null ? null : Math.round(((mm / 25.4) * 72) * 10) / 10; };
const decodeName = (s) => String(s || '').replace(/_([0-9a-f]{2})_/gi, (m, h) => String.fromCharCode(parseInt(h, 16)));

export async function readOdt(input, settings, { ext = 'odt' } = {}) {
  const model = defaultModel(settings);
  const fid = new Set();
  let contentRoot; let stylesRoot; let metaRoot = null; let zip = null; let flatXml = null;
  const pictures = {};
  const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
  const isZip = u8[0] === 0x50 && u8[1] === 0x4b;
  if (isZip) {
    zip = await JSZip.loadAsync(u8);
    const c = zip.file('content.xml');
    if (!c) throw new Error('This file has no OpenDocument text inside.');
    contentRoot = parseXml(await c.async('string'));
    const s = zip.file('styles.xml');
    stylesRoot = s ? parseXml(await s.async('string')) : { n: 'x', a: {}, c: [] };
    const mt = zip.file('meta.xml');
    if (mt) metaRoot = parseXml(await mt.async('string'));
  } else {
    flatXml = new TextDecoder('utf-8').decode(u8);
    contentRoot = parseXml(flatXml);
    stylesRoot = contentRoot;
    metaRoot = contentRoot;
  }
  const picture = async (href, frame) => {
    if (!href) {
      const bin = frame && all(frame, 'office:binary-data')[0];
      return bin ? `data:image/png;base64,${xtext(bin).replace(/\s+/g, '')}` : null;
    }
    if (/^(https?:|data:)/.test(href)) return href;
    if (pictures[href] !== undefined) return pictures[href];
    const f = zip && zip.file(href.replace(/^\.\//, ''));
    const mime = mimeOf(href);
    pictures[href] = f && /^image\/(png|jpeg|gif|bmp|svg\+xml|webp)$/.test(mime) ? dataUrl(await f.async('uint8array'), mime) : null;
    return pictures[href];
  };

  // ---------- styles ----------
  const styles = { paragraph: {}, text: {}, 'table-cell': {}, 'table-column': {}, table: {}, graphic: {} };
  const listStyles = {};
  const fonts = {};
  const readStyles = (root) => {
    for (const fd of all(root, 'style:font-face')) fonts[fd.a['style:name']] = (fd.a['svg:font-family'] || fd.a['style:name']).replace(/^'|'$/g, '');
    for (const container of [...all(root, 'office:styles'), ...all(root, 'office:automatic-styles')]) {
      for (const s of kids(container, 'style:style')) {
        const fam = s.a['style:family'];
        if (!styles[fam]) styles[fam] = {};
        styles[fam][s.a['style:name']] = {
          name: s.a['style:name'], display: (s.a['style:display-name'] || decodeName(s.a['style:name'])).toLowerCase(), parent: s.a['style:parent-style-name'],
          p: (kid(s, 'style:paragraph-properties') || {}).a || {}, t: (kid(s, 'style:text-properties') || {}).a || {},
          cell: (kid(s, 'style:table-cell-properties') || {}).a || {}, col: (kid(s, 'style:table-column-properties') || {}).a || {},
          outline: s.a['style:default-outline-level'], listStyle: s.a['style:list-style-name'], auto: container.n === 'office:automatic-styles',
          masterPage: s.a['style:master-page-name'] || null,
        };
      }
      for (const l of kids(container, 'text:list-style')) {
        const levels = {};
        for (const lv of kids(l, '*')) {
          const k = Number(lv.a['text:level'] || 1) - 1;
          if (lv.n === 'text:list-level-style-bullet') levels[k] = { kind: /[☐☑☒□]/.test(lv.a['text:bullet-char'] || '') ? 'task' : 'bullet' };
          else if (lv.n === 'text:list-level-style-number') {
            const f = lv.a['style:num-format'] || '1';
            levels[k] = f === '' ? { kind: 'bullet' } : { kind: 'ordered', listStyle: { a: 'lower-alpha', A: 'upper-alpha', i: 'lower-roman', I: 'upper-roman' }[f] || (Number(lv.a['text:display-levels'] || 1) > 1 ? 'outline' : null), start: Number(lv.a['text:start-value'] || 1) };
          } else levels[k] = { kind: 'bullet' };
        }
        listStyles[l.a['style:name']] = levels;
      }
    }
  };
  readStyles(stylesRoot);
  if (contentRoot !== stylesRoot) readStyles(contentRoot);
  const chain = (fam, name) => { const out = []; let s = styles[fam] && styles[fam][name]; let g = 0; while (s && g++ < 20) { out.unshift(s); s = styles[fam][s.parent]; } return out; };
  const merged = (fam, name, key) => chain(fam, name).reduce((acc, s) => ({ ...acc, ...s[key] }), {});
  const defaultPara = all(stylesRoot, 'style:default-style').find((d) => d.a['style:family'] === 'paragraph');
  if (defaultPara) {
    const tp = (kid(defaultPara, 'style:text-properties') || {}).a || {};
    if (tp['style:font-name']) model.defaults.font = fonts[tp['style:font-name']] || tp['style:font-name'];
    if (tp['fo:font-size']) model.defaults.fontSize = lenPt(tp['fo:font-size']) || model.defaults.fontSize;
  }
  const std = styles.paragraph.Standard || styles.paragraph.Default_20_Paragraph_20_Style;
  if (std) {
    const tp = merged('paragraph', std.name, 't');
    if (tp['style:font-name']) model.defaults.font = fonts[tp['style:font-name']] || tp['style:font-name'];
    if (tp['fo:font-size'] && !/%/.test(tp['fo:font-size'])) model.defaults.fontSize = lenPt(tp['fo:font-size']) || model.defaults.fontSize;
    const pp = merged('paragraph', std.name, 'p');
    model.defaults.spaceAfter = pp['fo:margin-bottom'] ? lenPt(pp['fo:margin-bottom']) : 0;
    model.defaults.lineHeight = /%$/.test(pp['fo:line-height'] || '') ? Number(pp['fo:line-height'].slice(0, -1)) / 100 : 1;
  } else { model.defaults.spaceAfter = 0; model.defaults.lineHeight = 1; }
  const roleOf = (name) => {
    for (const s of chain('paragraph', name).reverse()) {
      const d = s.display;
      const hm = /^(?:heading|überschrift)\s*(\d)$/.exec(d); if (hm) return { heading: Number(hm[1]) };
      if (d === 'title' || d === 'titel') return { pstyle: 'title' };
      if (d === 'subtitle' || d === 'untertitel') return { pstyle: 'subtitle' };
      if (d === 'quotations' || d === 'zitat' || d === 'quote') return { quote: true };
      if (d === 'preformatted text' || d === 'vorformatierter text' || d === 'source text') return { code: true };
      if (d === 'horizontal line' || d === 'horizontale linie') return { hr: true };
      if (/^contents \d|^inhaltsverzeichnis \d/.test(d)) return { toc: true };
      if (d === 'standard' || d === 'default paragraph style' || d === 'text body' || d === 'textkörper' || d === 'body text') return {};
    }
    return {};
  };
  const textMarks = (t) => {
    const m = [];
    if (t['fo:font-weight'] && t['fo:font-weight'] !== 'normal' && Number(t['fo:font-weight']) !== 400) m.push({ type: 'bold' });
    if (t['fo:font-style'] === 'italic' || t['fo:font-style'] === 'oblique') m.push({ type: 'italic' });
    if (t['style:text-underline-style'] && t['style:text-underline-style'] !== 'none') m.push({ type: 'underline' });
    if (t['style:text-line-through-style'] && t['style:text-line-through-style'] !== 'none') m.push({ type: 'strike' });
    const pos = t['style:text-position'];
    if (pos) { const f = pos.split(/\s+/)[0]; if (f === 'super' || parseFloat(f) > 0) m.push({ type: 'superscript' }); else if (f === 'sub' || parseFloat(f) < 0) m.push({ type: 'subscript' }); }
    const font = t['style:font-name'] ? fonts[t['style:font-name']] || t['style:font-name'] : t['fo:font-family'] ? t['fo:font-family'].replace(/^'|'$/g, '') : null;
    const size = t['fo:font-size'] && !/%/.test(t['fo:font-size']) ? lenPt(t['fo:font-size']) : null;
    const color = t['fo:color'] && t['fo:color'] !== '#000000' ? t['fo:color'] : null;
    const attrs = { fontFamily: font && font !== model.defaults.font ? font : null, fontSize: size && size !== model.defaults.fontSize ? `${size}pt` : null, color };
    if (attrs.fontFamily || attrs.fontSize || attrs.color) m.push({ type: 'textStyle', attrs });
    const bg = t['fo:background-color'];
    if (bg && bg !== 'transparent' && bg !== '#ffffff') m.push({ type: 'highlight', attrs: { color: bg } });
    return m;
  };
  const paraAttrs = (p) => {
    const a = {};
    const al = p['fo:text-align'];
    if (al) a.textAlign = { center: 'center', end: 'right', right: 'right', justify: 'justify' }[al] || null;
    if (!a.textAlign) delete a.textAlign;
    if (p['fo:margin-top'] != null) a.spaceBefore = lenPt(p['fo:margin-top']);
    if (p['fo:margin-bottom'] != null) a.spaceAfter = lenPt(p['fo:margin-bottom']);
    if (p['fo:margin-left'] && lenPt(p['fo:margin-left'])) a.indent = lenPt(p['fo:margin-left']);
    if (p['fo:text-indent'] && lenPt(p['fo:text-indent'])) a.firstLine = lenPt(p['fo:text-indent']);
    if (/%$/.test(p['fo:line-height'] || '')) a.lineHeight = Number(p['fo:line-height'].slice(0, -1)) / 100;
    return a;
  };

  // ---------- comments, tracked changes ----------
  const regions = {};
  for (const r of all(contentRoot, 'text:changed-region')) {
    const id = r.a['text:id'] || r.a['xml:id'];
    const ch = kids(r, '*')[0];
    if (!ch) continue;
    const info = kid(ch, 'office:change-info');
    regions[id] = { kind: ch.n.replace('text:', ''), author: info ? xtext(kid(info, 'dc:creator')) : '', date: info ? xtext(kid(info, 'dc:date')) : null, deleted: ch.n === 'text:deletion' ? kids(ch, '*').filter((x) => x.n !== 'office:change-info') : [] };
  }
  if (Object.keys(regions).length) fid.add('trackedChanges');
  const openChanges = new Set();
  const openComments = new Set();
  const commentByName = {};

  // ---------- inline ----------
  async function inlines(el, ctx, out) {
    for (const c of el.c || []) {
      if (typeof c === 'string') { addText(c, ctx, out); continue; }
      switch (c.n) {
        case 'text:span': {
          const sm = textMarks(merged('text', c.a['text:style-name'], 't'));
          if (/^OfEq/.test(c.a['text:style-name'] || '')) { out.push({ type: 'equation', attrs: { latex: xtext(c) } }); break; }
          await inlines(c, { ...ctx, marks: [...ctx.marks.filter((m) => !sm.some((s) => s.type === m.type)), ...sm] }, out);
          break;
        }
        case 'text:a': await inlines(c, { ...ctx, marks: [...ctx.marks, { type: 'link', attrs: { href: c.a['xlink:href'] || '' } }] }, out); break;
        case 'text:s': addText(' '.repeat(Number(c.a['text:c'] || 1)), ctx, out, true); break;
        case 'text:tab': addText('\t', ctx, out, true); break;
        case 'text:line-break': out.push({ type: 'hardBreak' }); break;
        case 'text:soft-page-break': case 'text:bookmark': case 'text:bookmark-start': case 'text:bookmark-end': case 'text:reference-mark': case 'text:reference-mark-start': case 'text:reference-mark-end': case 'text:alphabetical-index-mark': break;
        case 'text:note': {
          const body = kid(c, 'text:note-body');
          if (c.a['text:note-class'] === 'endnote') fid.add('endnotes');
          out.push({ type: 'footnote', attrs: { text: body ? kids(body, '*').map(xtext).join(' ').replace(/\s+/g, ' ').trim() : '' } });
          break;
        }
        case 'office:annotation': {
          const name = c.a['office:name'];
          const id = `n${uid('')}`;
          const textOf = kids(c, 'text:p').map(xtext).join('\n').trim();
          const parent = c.a['loext:parent-name'] && commentByName[c.a['loext:parent-name']];
          if (parent) { parent.replies.push({ author: xtext(kid(c, 'dc:creator')), date: xtext(kid(c, 'dc:date')) || null, text: textOf }); break; }
          const cm = { id, author: xtext(kid(c, 'dc:creator')), date: xtext(kid(c, 'dc:date')) || null, text: textOf, resolved: c.a['loext:resolved'] === 'true', replies: [] };
          model.comments.push(cm);
          if (name) { commentByName[name] = cm; openComments.add(id); } else {
            // point comment: anchor to the word before it
            const last = [...out].reverse().find((n) => n.type === 'text');
            if (last) last.marks = [...(last.marks || []), { type: 'comment', attrs: { id } }];
            else ctx.pendingComment = id;
          }
          break;
        }
        case 'office:annotation-end': { const cm = commentByName[c.a['office:name']]; if (cm) openComments.delete(cm.id); break; }
        case 'text:change-start': openChanges.add(c.a['text:change-id']); break;
        case 'text:change-end': openChanges.delete(c.a['text:change-id']); break;
        case 'text:change': {
          const r = regions[c.a['text:change-id']];
          if (r && r.kind === 'deletion') {
            const t = r.deleted.map(xtext).join(' ');
            if (t) out.push({ type: 'text', text: t, marks: [...ctx.marks, { type: 'deletion', attrs: { id: c.a['text:change-id'], author: r.author, date: r.date } }] });
          }
          break;
        }
        case 'draw:frame': case 'draw:a': out.push(await frame(c.n === 'draw:a' ? kid(c, 'draw:frame') || c : c, ctx)); break;
        case 'draw:custom-shape': case 'draw:rect': case 'draw:ellipse': case 'draw:line': case 'draw:g': case 'draw:polygon': case 'draw:path': case 'draw:connector':
          fid.add('shape'); out.push({ type: 'preserved', attrs: { format: 'odt', kind: 'shape', label: 'Shape', text: all(c, 'text:p').map(xtext).join(' '), src: null, width: null, xml: serialize(c) } }); break;
        case 'text:page-number': case 'text:page-count': case 'text:date': case 'text:time': case 'text:title': case 'text:subject': case 'text:author-name': case 'text:file-name': case 'text:chapter': case 'text:sequence': case 'text:variable-get': case 'text:user-field-get': case 'text:bookmark-ref': case 'text:reference-ref': case 'text:sequence-ref': case 'text:note-ref': case 'text:word-count': case 'text:character-count':
          out.push({ type: 'preserved', attrs: { format: 'odt', kind: 'field', label: c.n.replace('text:', ''), text: xtext(c), src: null, width: null, xml: serialize(c) } }); break;
        case 'text:meta': case 'text:ruby-base': case 'text:hidden-text': await inlines(c, ctx, out); break;
        case 'text:ruby': await inlines(kid(c, 'text:ruby-base') || { c: [] }, ctx, out); break;
        default: if (c.c && c.c.length) await inlines(c, ctx, out);
      }
    }
  }
  function addText(raw, ctx, out, literal = false) {
    const t = literal ? raw : raw.replace(/[\r\n\t ]+/g, ' ');
    if (!t) return;
    const marks = [...ctx.marks];
    for (const id of openChanges) { const r = regions[id]; if (!r) continue; if (r.kind === 'insertion') marks.push({ type: 'insertion', attrs: { id, author: r.author, date: r.date } }); if (r.kind === 'format-change') marks.push({ type: 'formatChange', attrs: { id, author: r.author, date: r.date, old: '' } }); }
    for (const id of openComments) marks.push({ type: 'comment', attrs: { id } });
    if (ctx.pendingComment) { marks.push({ type: 'comment', attrs: { id: ctx.pendingComment } }); ctx.pendingComment = null; }
    out.push({ type: 'text', text: t, marks });
  }
  async function frame(f, ctx) {
    const img = kid(f, 'draw:image');
    const obj = kid(f, 'draw:object') || kid(f, 'draw:object-ole');
    const tb = kid(f, 'draw:text-box');
    const w = f.a['svg:width'] ? Math.round((lenMm(f.a['svg:width']) || 0) * (96 / 25.4)) : null;
    const h = f.a['svg:height'] ? Math.round((lenMm(f.a['svg:height']) || 0) * (96 / 25.4)) : null;
    const gs = merged('graphic', f.a['draw:style-name'], 'p');
    if (img && !obj && !tb) {
      const src = await picture(img.a['xlink:href'], img);
      if (src) {
        const anchor = f.a['text:anchor-type'];
        let align = null; let wrap = null;
        if (anchor && anchor !== 'as-char') { const hp = gs['style:horizontal-pos']; align = hp === 'right' ? 'right' : hp === 'center' ? 'center' : 'left'; if (gs['style:wrap'] && !['none', 'run-through'].includes(gs['style:wrap'])) wrap = 'square'; }
        return { type: 'image', attrs: { src, alt: xtext(kid(f, 'svg:desc')) || xtext(kid(f, 'svg:title')) || '', title: null, width: w, height: h, align, wrap } };
      }
    }
    const kind = obj ? (/chart/i.test(serialize(obj)) ? 'chart' : 'ole') : tb ? 'textbox' : 'object';
    fid.add(kind);
    const fallbackImg = kids(f, 'draw:image')[0];
    const src = fallbackImg ? await picture(fallbackImg.a['xlink:href'], fallbackImg) : null;
    void ctx;
    return { type: 'preserved', attrs: { format: 'odt', kind, label: { chart: 'Chart', ole: 'Embedded object', textbox: 'Text box', object: 'Object' }[kind], text: tb ? all(tb, 'text:p').map(xtext).join('\n') : '', src, width: w, xml: serialize(f) } };
  }

  // ---------- blocks ----------
  async function paragraphNode(el, extra = {}) {
    const sname = el.a['text:style-name'];
    const role = roleOf(sname);
    const pp = merged('paragraph', sname, 'p');
    const attrs = paraAttrs(pp);
    const own = (styles.paragraph[sname] || {});
    if (role.pstyle) attrs.pstyle = role.pstyle;
    // run formatting set on custom paragraph styles
    const tp = merged('paragraph', sname, 't');
    const base = role.heading || role.pstyle ? [] : textMarks(tp);
    const ctx = { marks: base };
    const out = [];
    await inlines(el, ctx, out);
    const level = el.n === 'text:h' ? Number(el.a['text:outline-level'] || role.heading || 1) : role.heading || 0;
    let node;
    if (role.hr && !out.length) node = { type: 'horizontalRule' };
    else if (role.code) node = { type: 'codeBlock', attrs: {}, content: [{ type: 'text', text: out.map((x) => x.text || (x.type === 'hardBreak' ? '\n' : '')).join('') }].filter((x) => x.text) };
    else if (level && level <= 6) node = { type: 'heading', attrs: { level, ...attrs }, content: normalizeInline(out) };
    else node = { type: 'paragraph', attrs, content: normalizeInline(out) };
    if (node.content && !node.content.length) delete node.content;
    // keep the original style when nothing about it was changed (patch writer)
    if (node.attrs && sname) node.attrs.oo = { style: sname, sig: sigOf(node), auto: !!own.auto };
    const res = [];
    if (pp['fo:break-before'] === 'page') res.push({ type: 'pageBreak' });
    if (!(pp['fo:break-after'] === 'page' && !out.length && node.type === 'paragraph')) res.push(node);
    if (pp['fo:break-after'] === 'page') res.push({ type: 'pageBreak' });
    void extra;
    return { nodes: res, quote: !!role.quote, toc: !!role.toc };
  }
  async function list(el, depth, inherited) {
    const styleName = el.a['text:style-name'] || inherited;
    const levels = listStyles[styleName] || {};
    const lv = levels[depth] || levels[0] || { kind: 'bullet' };
    const items = [];
    let start = lv.start || 1;
    const li = kids(el, 'text:list-item').concat(kids(el, 'text:list-header'));
    const firstSv = li[0] && li[0].a['text:start-value'];
    if (firstSv) start = Number(firstSv);
    let kind = lv.kind;
    for (const item of kids(el, '*').filter((x) => x.n === 'text:list-item' || x.n === 'text:list-header')) {
      const content = [];
      let checked = null;
      for (const c of kids(item, '*')) {
        if (c.n === 'text:p' || c.n === 'text:h') {
          const r = await paragraphNode(c);
          for (const n of r.nodes) {
            if (n.type === 'heading') { n.type = 'paragraph'; delete n.attrs.level; }
            if (n.type === 'paragraph' && n.attrs) { delete n.attrs.indent; delete n.attrs.firstLine; }
            content.push(n);
          }
        } else if (c.n === 'text:list') content.push(await list(c, depth + 1, styleName));
        else if (c.n === 'table:table') content.push(...await tableNode(c));
      }
      // ☐/☑ written as text (our writer, LibreOffice users) → task item
      const first = content[0];
      const ft = first && first.content && first.content[0];
      if (ft && ft.type === 'text') { const m = /^([☐☑☒])\s?/.exec(ft.text); if (m) { checked = m[1] !== '☐'; ft.text = ft.text.slice(m[0].length); if (!ft.text) first.content.shift(); } }
      if (checked !== null && kind !== 'ordered') kind = 'task';
      items.push({ content, checked });
    }
    if (kind === 'task') return { type: 'taskList', content: items.map((it) => ({ type: 'taskItem', attrs: { checked: !!it.checked }, content: it.content.length ? it.content : [{ type: 'paragraph' }] })) };
    const node = kind === 'ordered' ? { type: 'orderedList', attrs: { start, listStyle: lv.listStyle || null }, content: [] } : { type: 'bulletList', attrs: { listStyle: null }, content: [] };
    node.content = items.map((it) => ({ type: 'listItem', content: it.content.length ? it.content : [{ type: 'paragraph' }] }));
    return node;
  }
  async function tableNode(el) {
    const name = el.a['table:name'] || '';
    const cols = [];
    const colEls = [];
    const collectCols = (x) => { for (const c of kids(x, '*')) { if (c.n === 'table:table-column') colEls.push(c); else if (c.n === 'table:table-columns' || c.n === 'table:table-header-columns' || c.n === 'table:table-column-group') collectCols(c); } };
    collectCols(el);
    for (const c of colEls) {
      const w = merged('table-column', c.a['table:style-name'], 'col')['style:column-width'];
      const px = w ? Math.round((lenMm(w) || 0) * (96 / 25.4)) : null;
      for (let k = 0; k < Number(c.a['table:number-columns-repeated'] || 1) && cols.length < 64; k++) cols.push(px);
    }
    const rows = [];
    const rowEls = [];
    const collectRows = (x, header) => { for (const c of kids(x, '*')) { if (c.n === 'table:table-row') rowEls.push({ el: c, header }); else if (c.n === 'table:table-header-rows') collectRows(c, true); else if (c.n === 'table:table-rows' || c.n === 'table:table-row-group') collectRows(c, header); } };
    collectRows(el, false);
    for (const { el: r, header } of rowEls) {
      const row = { type: 'tableRow', content: [] };
      let col = 0;
      for (const c of kids(r, '*')) {
        if (c.n === 'table:covered-table-cell') { col += Number(c.a['table:number-columns-repeated'] || 1); continue; }
        if (c.n !== 'table:table-cell') continue;
        const rep = Math.min(64, Number(c.a['table:number-columns-repeated'] || 1));
        for (let k = 0; k < rep; k++) {
          const span = Number(c.a['table:number-columns-spanned'] || 1);
          const cp = merged('table-cell', c.a['table:style-name'], 'cell');
          const bg = cp['fo:background-color'] && cp['fo:background-color'] !== 'transparent' ? cp['fo:background-color'] : null;
          const content = await blocksOf(c);
          const widths = cols.slice(col, col + span);
          row.content.push({ type: header ? 'tableHeader' : 'tableCell', attrs: { colspan: span, rowspan: Number(c.a['table:number-rows-spanned'] || 1), colwidth: widths.length === span && widths.every(Boolean) ? widths : null, backgroundColor: bg }, content: content.length ? content : [{ type: 'paragraph' }] });
          col += span;
        }
      }
      if (row.content.length) rows.push(row);
    }
    if (!rows.length) return [];
    if (/^of-callout-/.test(name) && rows.length === 1 && rows[0].content.length === 1) return [{ type: 'callout', attrs: { kind: name.replace('of-callout-', '').replace(/\d+$/, '') || 'info' }, content: rows[0].content[0].content }];
    return [{ type: 'table', content: rows }];
  }
  async function blocksOf(container) {
    const out = [];
    let quote = null;
    const push = (n, isQuote) => {
      if (isQuote) { if (!quote) { quote = { type: 'blockquote', content: [] }; out.push(quote); } quote.content.push(n); return; }
      quote = null; out.push(n);
    };
    for (const c of kids(container, '*')) {
      switch (c.n) {
        case 'text:p': case 'text:h': {
          const r = await paragraphNode(c);
          if (r.toc) { const t = (r.nodes.find((n) => n.content) || {}).content; if (t) push({ type: 'paragraph', content: t }); break; }
          for (const n of r.nodes) push(n, r.quote && n.type !== 'pageBreak');
          break;
        }
        case 'text:list': push(await list(c, 0, null)); break;
        case 'table:table': for (const t of await tableNode(c)) push(t); break;
        case 'text:section': for (const b of await blocksOf(c)) push(b); break;
        case 'text:table-of-content': {
          const body = kid(c, 'text:index-body');
          const titleEl = body && kid(body, 'text:index-title');
          const title = titleEl ? kids(titleEl, 'text:p').map(xtext).join(' ').trim() : '';
          const entries = [];
          for (const p of body ? kids(body, 'text:p') : []) {
            const st = styles.paragraph[p.a['text:style-name']];
            const m = st && /(\d)\s*$/.exec(chain('paragraph', st.name).map((s) => s.display).reverse().find((d) => /contents|verzeichnis/.test(d)) || '');
            const t = xtext(p).replace(/\t+\s*\d*\s*$/, '').replace(/\t/g, ' ').trim();
            if (t) entries.push({ level: m ? Number(m[1]) : 1, text: t });
          }
          push({ type: 'toc', attrs: { title, entries } });
          break;
        }
        case 'text:index-title': case 'text:illustration-index': case 'text:alphabetical-index': case 'text:bibliography': case 'text:user-index': case 'text:object-index': case 'text:table-index': {
          const body = kid(c, 'text:index-body');
          for (const b of await blocksOf(body || c)) push(b);
          fid.add('indexes');
          break;
        }
        case 'draw:frame': { const f = await frame(c, {}); push({ type: 'paragraph', content: [f] }); break; }
        case 'text:tracked-changes': case 'text:sequence-decls': case 'text:variable-decls': case 'text:user-field-decls': case 'text:dde-connection-decls': case 'office:forms': case 'text:soft-page-break': case 'table:calculation-settings': case 'table:content-validations': case 'table:label-ranges': break;
        case 'text:change': case 'text:change-start': case 'text:change-end': { const tmp = []; await inlines({ c: [c] }, { marks: [] }, tmp); if (tmp.length) push({ type: 'paragraph', content: tmp }); break; }
        default: break;
      }
    }
    return out;
  }

  const officeText = all(contentRoot, 'office:text')[0];
  if (!officeText) throw new Error('This OpenDocument file has no text body.');
  const content = await blocksOf(officeText);
  model.doc = fixDoc({ type: 'doc', content });

  // ---------- page layout, header/footer ----------
  const masters = all(stylesRoot, 'style:master-page');
  const master = masters.find((m) => m.a['style:name'] === 'Standard') || masters[0];
  const layouts = all(stylesRoot, 'style:page-layout');
  const layout = master && layouts.find((l) => l.a['style:name'] === master.a['style:page-layout-name']);
  const pl = layout && kid(layout, 'style:page-layout-properties');
  if (pl) {
    const a = pl.a;
    const w = lenMm(a['fo:page-width']); const h = lenMm(a['fo:page-height']);
    if (w && h) {
      model.page.width = Math.round(Math.min(w, h) * 10) / 10; model.page.height = Math.round(Math.max(w, h) * 10) / 10;
      model.page.orientation = a['style:print-orientation'] === 'landscape' || w > h ? 'landscape' : 'portrait';
      model.page.size = sizeName(w, h);
    }
    const mm = (k, d) => { const v = lenMm(a[k]); return v == null ? d : Math.round(v * 10) / 10; };
    model.page.margins = { top: mm('fo:margin-top', 20), right: mm('fo:margin-right', 20), bottom: mm('fo:margin-bottom', 20), left: mm('fo:margin-left', 20) };
    // header/footer heights are part of the margin in ODF
    const hp = kid(kid(layout, 'style:header-style') || { c: [] }, 'style:header-footer-properties');
    const fp = kid(kid(layout, 'style:footer-style') || { c: [] }, 'style:header-footer-properties');
    if (hp && hp.a['svg:height']) model.page.margins.top = Math.round((model.page.margins.top + (lenMm(hp.a['svg:height']) || 0) + (lenMm(hp.a['fo:margin-bottom']) || 0)) * 10) / 10;
    if (fp && fp.a['svg:height']) model.page.margins.bottom = Math.round((model.page.margins.bottom + (lenMm(fp.a['svg:height']) || 0) + (lenMm(fp.a['fo:margin-top']) || 0)) * 10) / 10;
    const cols = kid(pl, 'style:columns');
    if (cols && Number(cols.a['fo:column-count']) > 1) model.page.columns = Math.min(3, Number(cols.a['fo:column-count']));
    if (a['fo:background-color'] && a['fo:background-color'] !== 'transparent' && a['fo:background-color'] !== '#ffffff') model.page.color = a['fo:background-color'];
  }
  const hfText = (el) => {
    let s = '';
    const visit = (x) => { for (const c of x.c || []) { if (typeof c === 'string') s += c; else if (c.n === 'text:page-number') s += '{page}'; else if (c.n === 'text:page-count') s += '{pages}'; else if (c.n === 'text:tab') s += ' '; else if (c.n === 'text:s') s += ' '; else visit(c); } };
    visit(el);
    return s.replace(/\s+/g, ' ').trim();
  };
  for (const [tagName, key] of [['style:header', 'header'], ['style:footer', 'footer']]) {
    const el = master && kid(master, tagName);
    if (!el || el.a['style:display'] === 'false') continue;
    const p = all(el, 'text:p')[0];
    const al = p ? merged('paragraph', p.a['text:style-name'], 'p')['fo:text-align'] : null;
    const tabbed = p && all(p, 'text:tab').length > 0;
    model[key] = { text: hfText(el), align: al === 'center' || tabbed ? 'center' : al === 'end' || al === 'right' ? 'right' : 'left' };
  }
  if (master && kid(master, 'style:header-first')) model.titlePage = true; // ODF 1.3 (LibreOffice 7+)
  else if (master && (model.header || model.footer)) {
    // ODF 1.2 way (our writer, older LibreOffice "First Page" style): the first paragraph starts a master page without
    // header/footer whose next page style is the main one
    const first = (officeText.c || []).find((x) => x && typeof x === 'object' && (x.n === 'text:p' || x.n === 'text:h'));
    const st = first && styles.paragraph[first.a['text:style-name']];
    const fmName = st && (st.masterPage || (styles.paragraph[st.parent] || {}).masterPage);
    const fm = fmName && fmName !== master.a['style:name'] && masters.find((m) => m.a['style:name'] === fmName);
    const shows = (el) => el && el.a['style:display'] !== 'false' && hfText(el);
    if (fm && (!fm.a['style:next-style-name'] || fm.a['style:next-style-name'] === master.a['style:name'])
      && !shows(kid(fm, 'style:header')) && !shows(kid(fm, 'style:footer'))) model.titlePage = true;
  }

  // ---------- meta ----------
  if (metaRoot) {
    const g = (n) => { const e = all(metaRoot, n)[0]; return e ? xtext(e) : ''; };
    model.meta = { ...model.meta, title: g('dc:title'), author: g('meta:initial-creator') || g('dc:creator') || model.meta.author, subject: g('dc:subject'), keywords: all(metaRoot, 'meta:keyword').map(xtext).join(', '), created: g('meta:creation-date'), modified: g('dc:date') };
  }
  if (zip && Object.keys(zip.files).some((f) => /^Basic\/|^Scripts\//.test(f))) fid.add('macros');
  if (zip && zip.file('META-INF/documentsignatures.xml')) fid.add('signature');
  const settingsXml = zip && zip.file('settings.xml') ? await zip.file('settings.xml').async('string') : '';
  if (/<text:tracked-changes\b[^>]*text:track-changes="true"/.test(xtextSafe(contentRoot)) || /RecordChanges"[^>]*>true</.test(settingsXml)) model.trackChanges = true;
  model.fidelity = [...fid];
  model.keep.odt = { zip: isZip ? u8 : null, flat: flatXml, ext, headerFooter: { header: { ...model.header }, footer: { ...model.footer } } };
  return model;
}
const xtextSafe = (root) => { const tc = all(root, 'text:tracked-changes')[0]; return tc ? serialize({ n: tc.n, a: tc.a, c: [] }) : ''; };

/** signature of the modeled paragraph formatting (to know if a kept ODF style still fits) */
export function sigOf(node) {
  const a = node.attrs || {};
  return JSON.stringify([node.type, a.level || 0, a.textAlign || null, a.spaceBefore ?? null, a.spaceAfter ?? null, a.indent || null, a.firstLine || null, a.lineHeight || null, a.pstyle || null]);
}

export async function odtToText(bytes) {
  const m = await readOdt(bytes);
  const { docText } = await import('./pm.js');
  return docText(m.doc);
}
