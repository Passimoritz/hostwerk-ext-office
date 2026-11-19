/*
 * OpenDocument text writer: odt, ott (template mimetype) and fodt (flat XML, images inline).
 *
 * Lossless: with the original package (model.keep.odt) only content.xml, meta.xml (dates/title), the manifest (new
 * pictures) and – when page setup or header/footer changed – styles.xml are rewritten; the original automatic styles,
 * declarations, settings.xml, Basic/ macros, embedded objects and everything else stay as they were. Paragraphs whose
 * formatting you did not change keep their original style; preserved nodes are written back verbatim.
 * The package follows ODF: "mimetype" first and stored uncompressed.
 */
import JSZip from 'jszip';
import { esc, escAttr } from '../../_shared/xmltree.js';
import { pageBox, firstFont, ptOf, hexColor, parseDataUrl, imageSize, fillFields, walk, TEXTBLOCKS } from './pm.js';
import { equationText } from './equation.js';
import { sigOf } from './odt-read.js';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0', table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink', dc: 'http://purl.org/dc/elements/1.1/', meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0', svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  loext: 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0', officeooo: 'http://openoffice.org/2009/office',
};
const nsAttrs = (list) => list.map((p) => `xmlns:${p}="${NS[p]}"`).join(' ');
const ALL = Object.keys(NS);
const XMLH = '<?xml version="1.0" encoding="UTF-8"?>\n';
const MIME = { odt: 'application/vnd.oasis.opendocument.text', ott: 'application/vnd.oasis.opendocument.text-template', fodt: 'application/vnd.oasis.opendocument.text' };
const cm = (mm) => `${(Math.round(mm * 100) / 1000).toFixed(3)}cm`;
const pt = (v) => `${Math.round(v * 10) / 10}pt`;
const LIST_FMT = { decimal: '1', 'lower-alpha': 'a', 'upper-alpha': 'A', 'lower-roman': 'i', 'upper-roman': 'I', outline: '1' };

/** common styles we rely on (added to styles.xml when missing) */
function commonStyles(model) {
  const d = model.defaults || {};
  const head = (l, size) => `<style:style style:name="Heading_20_${l}" style:display-name="Heading ${l}" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="${l}" style:class="text"><style:paragraph-properties fo:margin-top="${l === 1 ? '0.423cm' : '0.247cm'}" fo:margin-bottom="0.106cm" fo:keep-with-next="always"/><style:text-properties fo:font-size="${size}pt" fo:font-weight="bold" fo:color="${l <= 2 ? '#2f5496' : '#1f3763'}"${l === 4 ? ' fo:font-style="italic"' : ''}/></style:style>`;
  return {
    Standard: `<style:style style:name="Standard" style:family="paragraph" style:class="text"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="${cm(((d.spaceAfter != null ? d.spaceAfter : 8) * 25.4) / 72)}" fo:line-height="${Math.round((d.lineHeight || 1.15) * 100)}%"/></style:style>`,
    Heading: '<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text"><style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:keep-with-next="always"/></style:style>',
    Text_20_body: '<style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text"/>',
    Heading_20_1: head(1, 16), Heading_20_2: head(2, 13), Heading_20_3: head(3, 12), Heading_20_4: head(4, 11), Heading_20_5: head(5, 11), Heading_20_6: head(6, 11),
    Title: '<style:style style:name="Title" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Subtitle" style:class="chapter"><style:paragraph-properties fo:margin-bottom="0.141cm"/><style:text-properties fo:font-size="28pt"/></style:style>',
    Subtitle: '<style:style style:name="Subtitle" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:class="chapter"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.423cm"/><style:text-properties fo:font-size="14pt" fo:color="#5a5a5a"/></style:style>',
    Quotations: '<style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-left="1cm" fo:margin-right="1cm" fo:margin-top="0.2cm" fo:margin-bottom="0.2cm"/><style:text-properties fo:font-style="italic" fo:color="#404040"/></style:style>',
    Preformatted_20_Text: '<style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:background-color="#f4f4f4"/><style:text-properties style:font-name="Consolas" fo:font-family="Consolas" fo:font-size="10pt"/></style:style>',
    Footnote: '<style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-left="0.6cm" fo:margin-bottom="0cm" fo:text-indent="-0.6cm"/><style:text-properties fo:font-size="10pt"/></style:style>',
    Header: '<style:style style:name="Header" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:text-properties fo:font-size="9pt"/></style:style>',
    Footer: '<style:style style:name="Footer" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:text-properties fo:font-size="9pt"/></style:style>',
    Table_20_Contents: '<style:style style:name="Table_20_Contents" style:display-name="Table Contents" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-bottom="0cm"/></style:style>',
    Contents_20_Heading: '<style:style style:name="Contents_20_Heading" style:display-name="Contents Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index"><style:text-properties fo:font-size="14pt" fo:font-weight="bold" fo:color="#2f5496"/></style:style>',
    Contents_20_1: '<style:style style:name="Contents_20_1" style:display-name="Contents 1" style:family="paragraph" style:parent-style-name="Standard" style:class="index"><style:paragraph-properties fo:margin-bottom="0.1cm"/></style:style>',
    Contents_20_2: '<style:style style:name="Contents_20_2" style:display-name="Contents 2" style:family="paragraph" style:parent-style-name="Standard" style:class="index"><style:paragraph-properties fo:margin-left="0.5cm" fo:margin-bottom="0.1cm"/></style:style>',
    Contents_20_3: '<style:style style:name="Contents_20_3" style:display-name="Contents 3" style:family="paragraph" style:parent-style-name="Standard" style:class="index"><style:paragraph-properties fo:margin-left="1cm" fo:margin-bottom="0.1cm"/></style:style>',
    Horizontal_20_Line: '<style:style style:name="Horizontal_20_Line" style:display-name="Horizontal Line" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-bottom="0.3cm" fo:border-bottom="0.06pt solid #808080" fo:padding="0cm"/><style:text-properties fo:font-size="6pt"/></style:style>',
  };
}

export async function writeOdt(model, { ext = 'odt' } = {}) {
  const keep = model.keep && model.keep.odt;
  const flat = ext === 'fodt';
  const zip = keep && keep.zip && !flat ? await JSZip.loadAsync(keep.zip) : new JSZip();
  const patch = !!(keep && (keep.zip || keep.flat));
  const origContent = keep && keep.zip && zip.file('content.xml') ? await zip.file('content.xml').async('string') : keep && keep.flat ? keep.flat : null;
  let origStyles = keep && keep.zip && zip.file('styles.xml') ? await zip.file('styles.xml').async('string') : keep && keep.flat ? keep.flat : null;

  const takenNames = new Set(origContent ? [...origContent.matchAll(/style:name="([^"]+)"/g)].map((m) => m[1]) : []);
  const auto = []; // automatic style xml
  const autoKey = new Map();
  let pN = 0; let tN = 0; let lN = 0; let tbN = 0; let frN = 0; let imgN = 0;
  const fontsUsed = new Set([firstFont((model.defaults && model.defaults.font) || 'Calibri')]);
  const autoStyle = (prefix, family, body, extraAttrs = '') => {
    const key = `${family}|${extraAttrs}|${body}`;
    if (autoKey.has(key)) return autoKey.get(key);
    let name;
    do { name = `Of${prefix}${prefix === 'P' ? ++pN : prefix === 'T' ? ++tN : prefix === 'Tb' ? ++tbN : ++frN}`; } while (takenNames.has(name));
    auto.push(`<style:style style:name="${name}" style:family="${family}"${extraAttrs}>${body}</style:style>`);
    autoKey.set(key, name);
    return name;
  };
  const pictures = []; // { path, bytes, mime }
  const picPath = new Map();
  const picture = (src) => {
    if (picPath.has(src)) return picPath.get(src);
    const d = parseDataUrl(src);
    let res = null;
    if (d) {
      const path = `Pictures/of_image${++imgN}.${d.ext}`;
      pictures.push({ path, bytes: d.bytes, mime: d.mime, src });
      res = { href: path, size: imageSize(d.bytes), data: d };
    } else if (/^https?:/.test(src)) res = { href: src, size: null };
    picPath.set(src, res);
    return res;
  };

  // text style for a mark set
  const textStyleOf = (marks) => {
    let props = '';
    for (const m of marks || []) {
      const a = m.attrs || {};
      if (m.type === 'bold') props += ' fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"';
      else if (m.type === 'italic') props += ' fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"';
      else if (m.type === 'underline' || (m.type === 'link' && false)) props += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
      else if (m.type === 'strike') props += ' style:text-line-through-style="solid" style:text-line-through-type="single"';
      else if (m.type === 'superscript') props += ' style:text-position="super 58%"';
      else if (m.type === 'subscript') props += ' style:text-position="sub 58%"';
      else if (m.type === 'code') { props += ' style:font-name="Consolas"'; fontsUsed.add('Consolas'); }
      else if (m.type === 'textStyle') {
        if (a.fontFamily) { const f = firstFont(a.fontFamily); fontsUsed.add(f); props += ` style:font-name="${escAttr(f)}"`; }
        const p = ptOf(a.fontSize); if (p) props += ` fo:font-size="${p}pt" style:font-size-asian="${p}pt" style:font-size-complex="${p}pt"`;
        const c = hexColor(a.color); if (c) props += ` fo:color="#${c.toLowerCase()}"`;
      } else if (m.type === 'highlight') { const c = hexColor(a.color || '#ffff00') || 'FFFF00'; props += ` fo:background-color="#${c.toLowerCase()}"`; }
    }
    return props ? autoStyle('T', 'text', `<style:text-properties${props}/>`) : null;
  };
  const paraStyleOf = (node, parent, extra = '') => {
    const a = node.attrs || {};
    // unchanged formatting → keep the original style name (keeps everything we do not model)
    if (patch && a.oo && a.oo.style && a.oo.sig === sigOf(node) && !extra) return a.oo.auto ? keepAuto(a.oo.style) : a.oo.style;
    let pp = '';
    if (a.textAlign && a.textAlign !== 'left') pp += ` fo:text-align="${a.textAlign === 'right' ? 'end' : a.textAlign}"`;
    if (a.spaceBefore != null) pp += ` fo:margin-top="${pt(a.spaceBefore)}"`;
    if (a.spaceAfter != null) pp += ` fo:margin-bottom="${pt(a.spaceAfter)}"`;
    if (a.indent) pp += ` fo:margin-left="${pt(a.indent)}"`;
    if (a.firstLine) pp += ` fo:text-indent="${pt(a.firstLine)}"`;
    if (a.lineHeight) pp += ` fo:line-height="${Math.round(a.lineHeight * 100)}%"`;
    pp += extra;
    const base = a.oo && a.oo.style && !a.oo.auto && patch ? a.oo.style : parent;
    if (!pp) return base;
    return autoStyle('P', 'paragraph', `<style:paragraph-properties${pp}/>`, ` style:parent-style-name="${escAttr(base)}"`);
  };
  // original automatic styles of content.xml are kept as they are (names stay valid)
  const keepAuto = (name) => name;

  // ---------- comments / changes ----------
  const comments = new Map((model.comments || []).map((c) => [String(c.id), c]));
  const range = new Map();
  let counter = 0;
  walk(model.doc, (n, parent) => {
    if (parent && TEXTBLOCKS.has(parent.type)) {
      for (const m of n.marks || []) if (m.type === 'comment' && comments.has(String(m.attrs.id))) { const r = range.get(String(m.attrs.id)) || { first: counter, last: counter }; r.last = counter; range.set(String(m.attrs.id), r); }
      counter++;
    }
  });
  const startsAt = new Map(); const endsAt = new Map();
  for (const [id, r] of range) { (startsAt.get(r.first) || startsAt.set(r.first, []).get(r.first)).push(id); (endsAt.get(r.last) || endsAt.set(r.last, []).get(r.last)).push(id); }
  const regions = [];
  let chN = 0;
  const ix = { n: 0 };
  const info = (a) => `<office:change-info><dc:creator>${esc(a.author || 'Unknown')}</dc:creator><dc:date>${esc(a.date || new Date().toISOString().replace(/\.\d+Z$/, ''))}</dc:date></office:change-info>`;
  let footN = 0;

  const textXml = (s) => {
    let o = '';
    const parts = String(s).split(/(\t|\n| {2,})/);
    for (const p of parts) {
      if (!p) continue;
      if (p === '\t') o += '<text:tab/>';
      else if (p === '\n') o += '<text:line-break/>';
      else if (/^ {2,}$/.test(p)) o += ` <text:s text:c="${p.length - 1}"/>`;
      else o += esc(p).replace(/^ /, '<text:s/>');
    }
    return o;
  };
  const inlineXml = (content) => {
    let s = '';
    let openIns = null;
    const closeIns = () => { if (openIns) { s += `<text:change-end text:change-id="${openIns}"/>`; openIns = null; } };
    for (const n of content || []) {
      const at = ix.n++;
      for (const id of startsAt.get(at) || []) s += annotation(comments.get(id), id);
      const ins = n.type === 'text' && (n.marks || []).find((m) => m.type === 'insertion');
      const del = n.type === 'text' && (n.marks || []).find((m) => m.type === 'deletion');
      if (!ins || (openIns && openIns !== ins.__ch)) closeIns();
      if (ins && !openIns) { const id = `ofct${++chN}`; regions.push(`<text:changed-region xml:id="${id}" text:id="${id}"><text:insertion>${info(ins.attrs || {})}</text:insertion></text:changed-region>`); s += `<text:change-start text:change-id="${id}"/>`; openIns = id; }
      if (n.type === 'text') {
        if (del) {
          const id = `ofct${++chN}`;
          regions.push(`<text:changed-region xml:id="${id}" text:id="${id}"><text:deletion>${info(del.attrs || {})}<text:p>${esc(n.text)}</text:p></text:deletion></text:changed-region>`);
          s += `<text:change text:change-id="${id}"/>`;
        } else {
          const ts = textStyleOf((n.marks || []).filter((m) => !['link', 'comment', 'insertion', 'deletion', 'formatChange', 'runKeep'].includes(m.type)));
          let body = textXml(n.text);
          if (ts) body = `<text:span text:style-name="${ts}">${body}</text:span>`;
          const link = (n.marks || []).find((m) => m.type === 'link');
          if (link) body = `<text:a xlink:type="simple" xlink:href="${escAttr(link.attrs.href || '')}">${body}</text:a>`;
          s += body;
        }
      } else if (n.type === 'hardBreak') s += '<text:line-break/>';
      else if (n.type === 'image') s += imageXml(n.attrs || {});
      else if (n.type === 'footnote') { footN++; s += `<text:note text:id="ftn${footN}" text:note-class="footnote"><text:note-citation>${footN}</text:note-citation><text:note-body><text:p text:style-name="Footnote">${textXml((n.attrs && n.attrs.text) || '')}</text:p></text:note-body></text:note>`; }
      else if (n.type === 'equation') s += `<text:span text:style-name="OfEq">${esc(equationText((n.attrs && n.attrs.latex) || ''))}</text:span>`;
      else if (n.type === 'preserved') {
        const a = n.attrs || {};
        if (patch && a.format === 'odt' && a.xml) s += a.xml;
        else if (a.src) s += imageXml({ src: a.src, width: a.width });
        else if (a.text) s += textXml(a.text);
      }
      for (const id of endsAt.get(at) || []) s += `<office:annotation-end office:name="of${escAttr(id)}"/>${((comments.get(id) || {}).replies || []).map((r, k) => `<office:annotation-end office:name="of${escAttr(id)}r${k}"/>`).join('')}`;
    }
    closeIns();
    return s;
  };
  const annotation = (c, id) => {
    if (!c) return '';
    const one = (name, author, date, text, parent) => `<office:annotation office:name="${escAttr(name)}"${parent ? ` loext:parent-name="${escAttr(parent)}"` : ''}${c.resolved ? ' loext:resolved="true"' : ''}><dc:creator>${esc(author || '')}</dc:creator>${date ? `<dc:date>${esc(String(date).replace(/Z$/, ''))}</dc:date>` : ''}${String(text || '').split('\n').map((l) => `<text:p>${textXml(l)}</text:p>`).join('')}</office:annotation>`;
    return one(`of${id}`, c.author, c.date, c.text) + (c.replies || []).map((r, k) => one(`of${id}r${k}`, r.author, r.date, r.text, `of${id}`)).join('');
  };
  const imageXml = (a) => {
    const p = a.src && picture(a.src);
    if (!p) return '';
    const w = a.width || (p.size ? Math.min(p.size.width, 600) : 300);
    const h = a.height && a.width ? a.height : p.size ? Math.round(w * (p.size.height / p.size.width)) : Math.round(w * 0.66);
    const floated = a.wrap === 'square' && (a.align === 'left' || a.align === 'right');
    const centered = a.align === 'center' && !floated;
    const gstyle = autoStyle('Fr', 'graphic', `<style:graphic-properties style:wrap="${floated ? 'parallel' : 'none'}" style:vertical-pos="top" style:vertical-rel="${floated || centered ? 'paragraph' : 'baseline'}" style:horizontal-pos="${floated ? a.align : centered ? 'center' : 'from-left'}" style:horizontal-rel="paragraph" fo:margin-left="0.2cm" fo:margin-right="0.2cm" style:mirror="none" fo:clip="rect(0cm, 0cm, 0cm, 0cm)" draw:luminance="0%" draw:contrast="0%" draw:red="0%" draw:green="0%" draw:blue="0%" draw:gamma="100%" draw:color-inversion="false" draw:image-opacity="100%" draw:color-mode="standard"/>`, ' style:parent-style-name="Graphics"');
    const inner = flat && p.data ? `<draw:image><office:binary-data>${b64(p.data.bytes)}</office:binary-data></draw:image>` : `<draw:image xlink:href="${escAttr(p.href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`;
    return `<draw:frame draw:style-name="${gstyle}" draw:name="Image${imgN}_${frN}" text:anchor-type="${floated || centered ? 'paragraph' : 'as-char'}" svg:width="${cm((w * 25.4) / 96)}" svg:height="${cm((h * 25.4) / 96)}" draw:z-index="0">${inner}${a.alt ? `<svg:desc>${esc(a.alt)}</svg:desc>` : ''}</draw:frame>`;
  };

  // ---------- blocks ----------
  const listStyles = new Map();
  const listStyle = (kind) => {
    if (listStyles.has(kind)) return listStyles.get(kind);
    const name = `OfL${++lN}`;
    let lv = '';
    for (let l = 1; l <= 10; l++) {
      const pos = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${cm(6.35 * (l + 1))}" fo:text-indent="-0.635cm" fo:margin-left="${cm(6.35 * (l + 1))}"/></style:list-level-properties>`;
      if (kind === 'bullet') lv += `<text:list-level-style-bullet text:level="${l}" text:bullet-char="${['•', '◦', '▪'][(l - 1) % 3]}">${pos}</text:list-level-style-bullet>`;
      else if (kind === 'task') lv += `<text:list-level-style-bullet text:level="${l}" text:bullet-char=" ">${pos}</text:list-level-style-bullet>`;
      else lv += `<text:list-level-style-number text:level="${l}" style:num-suffix="." style:num-format="${LIST_FMT[kind] || '1'}"${kind === 'outline' ? ` text:display-levels="${l}"` : ''}>${pos}</text:list-level-style-number>`;
    }
    auto.push(`<text:list-style style:name="${name}">${lv}</text:list-style>`);
    listStyles.set(kind, name);
    return name;
  };
  const out = [];
  let pendingBreak = false;
  const brk = ' fo:break-before="page"';
  const para = (n, parent, tag = 'text:p', extraAttr = '') => {
    const extra = pendingBreak ? brk : '';
    pendingBreak = false;
    const st = paraStyleOf(n, parent, extra);
    return `<${tag} text:style-name="${escAttr(st)}"${extraAttr}>${inlineXml(n.content)}</${tag}>`;
  };
  const block = (n, ctx = {}) => {
    const a = n.attrs || {};
    if (pendingBreak && !['paragraph', 'heading', 'codeBlock', 'pageBreak', 'blockquote', 'bulletList', 'orderedList', 'taskList'].includes(n.type)) { pendingBreak = false; out.push(`<text:p text:style-name="${autoStyle('P', 'paragraph', '<style:paragraph-properties fo:break-after="page"/>', ' style:parent-style-name="Standard"')}"/>`); }
    switch (n.type) {
      case 'paragraph': {
        const parent = ctx.quote ? 'Quotations' : a.pstyle === 'title' ? 'Title' : a.pstyle === 'subtitle' ? 'Subtitle' : ctx.cell ? 'Table_20_Contents' : 'Standard';
        out.push(para(n, parent)); break;
      }
      case 'heading': { const l = Math.min(6, a.level || 1); out.push(para(n, `Heading_20_${l}`, 'text:h', ` text:outline-level="${l}"`)); break; }
      case 'blockquote': for (const c of n.content || []) block(c, { ...ctx, quote: true }); break;
      case 'codeBlock': {
        const code = (n.content || []).map((c) => { ix.n++; return c.text || ''; }).join('');
        const extra = pendingBreak ? autoStyle('P', 'paragraph', `<style:paragraph-properties${brk}/>`, ' style:parent-style-name="Preformatted_20_Text"') : 'Preformatted_20_Text';
        pendingBreak = false;
        out.push(`<text:p text:style-name="${extra}">${textXml(code)}</text:p>`);
        break;
      }
      case 'bulletList': case 'orderedList': case 'taskList': {
        const kind = n.type === 'bulletList' ? 'bullet' : n.type === 'taskList' ? 'task' : (a.listStyle || 'decimal');
        out.push(ctx.inList ? '<text:list>' : `<text:list text:style-name="${listStyle(kind)}">`);
        (n.content || []).forEach((li, k) => {
          out.push(`<text:list-item${k === 0 && n.type === 'orderedList' && a.start && a.start !== 1 ? ` text:start-value="${a.start}"` : ''}>`);
          (li.content || []).forEach((c, j) => {
            if (j === 0 && c.type === 'paragraph' && n.type === 'taskList') {
              const box = li.attrs && li.attrs.checked ? '☑' : '☐';
              const extra = pendingBreak ? brk : ''; pendingBreak = false;
              out.push(`<text:p text:style-name="${escAttr(paraStyleOf(c, 'Standard', extra))}">${box} ${inlineXml(c.content)}</text:p>`);
            } else if (['bulletList', 'orderedList', 'taskList'].includes(c.type)) block(c, { ...ctx, inList: c.type === n.type && (c.attrs || {}).listStyle === a.listStyle });
            else block(c, ctx);
          });
          out.push('</text:list-item>');
        });
        out.push('</text:list>');
        break;
      }
      case 'horizontalRule': out.push('<text:p text:style-name="Horizontal_20_Line"/>'); break;
      case 'pageBreak': pendingBreak = true; break;
      case 'toc': {
        const es = a.entries || [];
        out.push(`<text:table-of-content text:style-name="OfSect1" text:protected="true" text:name="Table of Contents1"><text:table-of-content-source text:outline-level="3"><text:index-title-template text:style-name="Contents_20_Heading">${esc(a.title || '')}</text:index-title-template>${[1, 2, 3].map((l) => `<text:table-of-content-entry-template text:outline-level="${l}" text:style-name="Contents_20_${l}"><text:index-entry-link-start text:style-name="Index_20_Link"/><text:index-entry-chapter/><text:index-entry-text/><text:index-entry-tab-stop style:type="right" style:leader-char="."/><text:index-entry-page-number/><text:index-entry-link-end/></text:table-of-content-entry-template>`).join('')}</text:table-of-content-source><text:index-body>${a.title ? `<text:index-title text:style-name="OfSect1" text:name="Table of Contents1_Head"><text:p text:style-name="Contents_20_Heading">${esc(a.title)}</text:p></text:index-title>` : ''}${es.map((e) => `<text:p text:style-name="Contents_20_${Math.min(3, e.level || 1)}">${esc(e.text)}</text:p>`).join('')}</text:index-body></text:table-of-content>`);
        break;
      }
      case 'callout': {
        const color = { warn: '#fdf6e6', ok: '#ecf8f2', note: '#f6f6f6' }[a.kind] || '#eef4fc';
        const border = { warn: '#e0c07a', ok: '#8fcfb1', note: '#c8c8c8' }[a.kind] || '#9bb7e0';
        const cs = autoStyle('Tb', 'table-cell', `<style:table-cell-properties fo:background-color="${color}" fo:padding="0.15cm" fo:border="0.5pt solid ${border}"/>`);
        const ts = autoStyle('Tb', 'table', `<style:table-properties style:width="${cm(usable())}" table:align="margins"/>`);
        tableCount++;
        out.push(`<table:table table:name="of-callout-${escAttr(a.kind || 'info')}${tableCount}" table:style-name="${ts}"><table:table-column/><table:table-row><table:table-cell table:style-name="${cs}" office:value-type="string">`);
        for (const c of n.content || []) block(c, { cell: true });
        out.push('</table:table-cell></table:table-row></table:table>');
        break;
      }
      case 'table': table(n); break;
      default: for (const c of n.content || []) block(c, ctx);
    }
  };
  const usable = () => { const b = pageBox(model.page); return b.width - model.page.margins.left - model.page.margins.right; };
  let tableCount = 0;
  const table = (n) => {
    const rows = n.content || [];
    const grid = [];
    rows.forEach((r, ri) => {
      grid[ri] = grid[ri] || [];
      let ci = 0;
      for (const c of r.content || []) {
        while (grid[ri][ci]) ci++;
        const cs = (c.attrs && c.attrs.colspan) || 1; const rs = (c.attrs && c.attrs.rowspan) || 1;
        for (let y = 0; y < rs; y++) { grid[ri + y] = grid[ri + y] || []; for (let x = 0; x < cs; x++) grid[ri + y][ci + x] = { cell: c, origin: x === 0 && y === 0 }; }
        ci += cs;
      }
    });
    const cols = Math.max(1, ...grid.map((r) => r.length));
    const widths = Array(cols).fill(null);
    for (const r of grid) for (let k = 0; k < cols; k++) { const g = r[k]; if (g && g.origin && g.cell.attrs && g.cell.attrs.colwidth) g.cell.attrs.colwidth.forEach((w, j) => { if (w && widths[k + j] == null) widths[k + j] = w; }); }
    tableCount++;
    const totalPx = (usable() * 96) / 25.4;
    const known = widths.filter(Boolean).reduce((s, w) => s + w, 0);
    const free = widths.filter((w) => !w).length;
    const each = free ? Math.max(30, (totalPx - known) / free) : 0;
    const W = widths.map((w) => w || each);
    const sum = W.reduce((s, w) => s + w, 0);
    const ts = autoStyle('Tb', 'table', `<style:table-properties style:width="${cm((sum * 25.4) / 96)}" table:align="left"/>`);
    out.push(`<table:table table:name="Table${tableCount}" table:style-name="${ts}">`);
    for (const w of W) out.push(`<table:table-column table:style-name="${autoStyle('Tb', 'table-column', `<style:table-column-properties style:column-width="${cm((w * 25.4) / 96)}"/>`)}"/>`);
    let inHeader = false;
    grid.forEach((r, ri) => {
      const isHeader = ri === 0 && (rows[0].content || []).every((c) => c.type === 'tableHeader');
      if (isHeader && !inHeader) { out.push('<table:table-header-rows>'); inHeader = true; }
      out.push('<table:table-row>');
      for (let k = 0; k < cols; k++) {
        const g = r[k];
        if (!g || !g.origin) { out.push(g ? '<table:covered-table-cell/>' : '<table:table-cell office:value-type="string"><text:p text:style-name="Table_20_Contents"/></table:table-cell>'); continue; }
        const ca = g.cell.attrs || {};
        const bg = hexColor(ca.backgroundColor);
        const cs = autoStyle('Tb', 'table-cell', `<style:table-cell-properties fo:padding="0.097cm" fo:border="0.5pt solid #999999"${bg ? ` fo:background-color="#${bg.toLowerCase()}"` : ''}/>`);
        out.push(`<table:table-cell table:style-name="${cs}" office:value-type="string"${ca.colspan > 1 ? ` table:number-columns-spanned="${ca.colspan}"` : ''}${ca.rowspan > 1 ? ` table:number-rows-spanned="${ca.rowspan}"` : ''}>`);
        const content = g.cell.content && g.cell.content.length ? g.cell.content : [{ type: 'paragraph' }];
        for (const c of content) block(c, { cell: true });
        out.push('</table:table-cell>');
      }
      out.push('</table:table-row>');
      if (isHeader && inHeader) { out.push('</table:table-header-rows>'); inHeader = false; }
    });
    out.push('</table:table>');
  };
  for (const n of (model.doc && model.doc.content) || []) block(n);
  if (pendingBreak) out.push(`<text:p text:style-name="${autoStyle('P', 'paragraph', `<style:paragraph-properties${brk}/>`, ' style:parent-style-name="Standard"')}"/>`);

  // "different first page" (title page without header/footer): the ODF 1.2 way – a first-page master page
  // (no header/footer, next page "Standard") started by the first paragraph. <style:header-first> is ODF 1.3 only and
  // makes Microsoft Word reject the whole file.
  let firstPageMaster = '';
  const hasHf = (model.header && model.header.text) || (model.footer && model.footer.text);
  if (model.titlePage && hasHf && !patch) {
    const m = /^<text:(p|h)\b[^>]*?text:style-name="([^"]*)"/.exec(out[0] || '');
    if (m) {
      const st = m[2];
      let name = 'OfFirstPageP'; let k = 1; while (takenNames.has(name) || auto.some((x) => x.includes(`style:name="${name}"`))) name = `OfFirstPageP${++k}`;
      const own = auto.find((x) => x.startsWith(`<style:style style:name="${st}" style:family="paragraph"`));
      auto.push(own
        ? own.replace(`style:name="${st}"`, `style:name="${name}" style:master-page-name="OfFirstPage"`)
        : `<style:style style:name="${name}" style:family="paragraph" style:parent-style-name="${escAttr(st)}" style:master-page-name="OfFirstPage"/>`);
      out[0] = out[0].replace(`text:style-name="${st}"`, `text:style-name="${name}"`);
      firstPageMaster = '<style:master-page style:name="OfFirstPage" style:display-name="First Page" style:page-layout-name="OfPm1" style:next-style-name="Standard"/>';
    }
  }

  // ---------- content.xml ----------
  const pick = (xml, tagName) => { if (!xml) return ''; const m = new RegExp(`<${tagName}\\b[\\s\\S]*?</${tagName}>|<${tagName}\\b[^>]*/>`).exec(xml); return m ? m[0] : ''; };
  const inner = (el) => el.replace(/^<[^>]+>/, '').replace(/<\/[^>]+>$/, '').replace(/^<[^>]+\/>$/, '');
  const fontDecls = () => {
    const orig = inner(pick(flat ? origStyles : origContent, 'office:font-face-decls'));
    const have = new Set([...orig.matchAll(/style:name="([^"]+)"/g)].map((m) => m[1]));
    return `<office:font-face-decls>${orig}${[...fontsUsed].filter((f) => !have.has(f)).map((f) => `<style:font-face style:name="${escAttr(f)}" svg:font-family="${escAttr(/\s/.test(f) ? `'${f}'` : f)}"/>`).join('')}</office:font-face-decls>`;
  };
  auto.push('<style:style style:name="OfEq" style:family="text"><style:text-properties fo:font-style="italic" style:font-name="Cambria Math"/></style:style>');
  auto.push('<style:style style:name="OfSect1" style:family="section"><style:section-properties style:editable="false"/></style:style>');
  fontsUsed.add('Cambria Math');
  const origAuto = patch ? inner(pick(origContent, 'office:automatic-styles')) : '';
  // the original office:text "prelude" (variable/sequence declarations, forms) stays
  let prelude = '';
  if (patch && origContent) {
    const ot = pick(origContent, 'office:text');
    for (const t of ['office:forms', 'text:variable-decls', 'text:sequence-decls', 'text:user-field-decls', 'text:dde-connection-decls']) prelude += pick(ot, t);
  }
  const tracked = regions.length ? `<text:tracked-changes${model.trackChanges ? ' text:track-changes="true"' : ''}>${regions.join('')}</text:tracked-changes>` : model.trackChanges ? '<text:tracked-changes text:track-changes="true"/>' : '';
  const bodyXml = `<office:body><office:text>${prelude}${tracked}${out.join('')}</office:text></office:body>`;
  const autoXml = `<office:automatic-styles>${origAuto}${auto.join('')}</office:automatic-styles>`;
  const scripts = patch ? pick(origContent, 'office:scripts') : '';

  // ---------- styles.xml ----------
  const box = pageBox(model.page);
  const mg = model.page.margins;
  const hfP = (x, cls) => {
    const al = x.align === 'right' ? 'end' : x.align === 'center' ? 'center' : 'start';
    const body = String(x.text).split(/(\{page\}|\{pages\}|\{date\})/).filter(Boolean).map((p) => (p === '{page}' ? '<text:page-number text:select-page="current">1</text:page-number>' : p === '{pages}' ? '<text:page-count>1</text:page-count>' : p === '{date}' ? `<text:date>${esc(new Date().toLocaleDateString())}</text:date>` : textXml(fillFields(p, { title: (model.meta && model.meta.title) || '' })))).join('');
    return `<text:p text:style-name="OfMP${cls}${al}">${body}</text:p>`;
  };
  const hfStyles = ['start', 'center', 'end'].flatMap((al) => ['Header', 'Footer'].map((cls) => `<style:style style:name="OfMP${cls}${al}" style:family="paragraph" style:parent-style-name="${cls}"><style:paragraph-properties fo:text-align="${al}"/></style:style>`)).join('');
  const layoutProps = `fo:page-width="${cm(box.width)}" fo:page-height="${cm(box.height)}" style:print-orientation="${model.page.orientation === 'landscape' ? 'landscape' : 'portrait'}" fo:margin-top="${cm(mg.top)}" fo:margin-bottom="${cm(mg.bottom)}" fo:margin-left="${cm(mg.left)}" fo:margin-right="${cm(mg.right)}"${model.page.color ? ` fo:background-color="${escAttr(model.page.color)}"` : ''} style:writing-mode="lr-tb"`;
  const colsXml = model.page.columns > 1 ? `<style:columns fo:column-count="${model.page.columns}" fo:column-gap="0.5cm"/>` : '';
  const hfMaster = `${model.header && model.header.text ? `<style:header>${hfP(model.header, 'Header')}</style:header>` : ''}${model.footer && model.footer.text ? `<style:footer>${hfP(model.footer, 'Footer')}</style:footer>` : ''}`;
  const common = commonStyles(model);
  let stylesXml;
  if (patch && keep.zip && origStyles && !flat) {
    stylesXml = origStyles;
    // missing common styles
    const missing = Object.entries(common).filter(([name]) => !new RegExp(`style:name="${name}"`).test(stylesXml)).map(([, x]) => x).join('');
    stylesXml = stylesXml.replace('</office:styles>', `${missing}</office:styles>`);
    // page layout of the master page
    const master = /<style:master-page\b[^>]*style:name="Standard"[^>]*>|<style:master-page\b[^>]*>/.exec(stylesXml);
    const layoutName = master && /style:page-layout-name="([^"]+)"/.exec(master[0]);
    if (layoutName) {
      stylesXml = stylesXml.replace(new RegExp(`(<style:page-layout\\b[^>]*style:name="${layoutName[1]}"[^>]*>[\\s\\S]*?<style:page-layout-properties\\b)([^>]*?)(/?>)`), (m, a, attrs, end) => {
        const keepAttrs = attrs.replace(/\s(fo:page-width|fo:page-height|style:print-orientation|fo:margin-top|fo:margin-bottom|fo:margin-left|fo:margin-right|fo:background-color|style:writing-mode)="[^"]*"/g, '');
        return `${a}${keepAttrs} ${layoutProps}${end}`;
      });
    }
    // header/footer only when you changed them
    const o = keep.headerFooter || {};
    const changed = (k) => !o[k] || o[k].text !== ((model[k] && model[k].text) || '') || o[k].align !== ((model[k] && model[k].align) || 'center');
    if (master && (changed('header') || changed('footer'))) {
      stylesXml = stylesXml.replace(/(<style:master-page\b[^>]*style:name="Standard"[^>]*>|<style:master-page\b[^>]*>)([\s\S]*?)(<\/style:master-page>)/, (m, open, body, close) => {
        let b = body;
        if (changed('header')) b = b.replace(/<style:header\b[\s\S]*?<\/style:header>|<style:header\b[^>]*\/>/, '').replace(/<style:header-first\b[\s\S]*?<\/style:header-first>|<style:header-first\b[^>]*\/>/, '');
        if (changed('footer')) b = b.replace(/<style:footer\b[\s\S]*?<\/style:footer>|<style:footer\b[^>]*\/>/, '').replace(/<style:footer-first\b[\s\S]*?<\/style:footer-first>|<style:footer-first\b[^>]*\/>/, '');
        const h = changed('header') && model.header && model.header.text ? `<style:header>${hfP(model.header, 'Header')}</style:header>` : '';
        const f = changed('footer') && model.footer && model.footer.text ? `<style:footer>${hfP(model.footer, 'Footer')}</style:footer>` : '';
        return `${open}${h}${b}${f}${close}`;
      });
      if (!/OfMPHeader/.test(origStyles)) stylesXml = /<office:automatic-styles\/>/.test(stylesXml) ? stylesXml.replace('<office:automatic-styles/>', `<office:automatic-styles>${hfStyles}</office:automatic-styles>`) : stylesXml.replace('</office:automatic-styles>', `${hfStyles}</office:automatic-styles>`);
    }
  } else {
    const d = model.defaults || {};
    const f = firstFont(d.font || 'Calibri');
    const stylesBody = `<office:styles><style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf"/></style:default-style><style:default-style style:family="paragraph"><style:paragraph-properties fo:hyphenation-ladder-count="no-limit" style:writing-mode="page"/><style:text-properties style:font-name="${escAttr(f)}" fo:font-size="${d.fontSize || 11}pt" fo:language="${(typeof document !== 'undefined' && document.documentElement.lang === 'de') ? 'de' : 'en'}" fo:country="${(typeof document !== 'undefined' && document.documentElement.lang === 'de') ? 'DE' : 'US'}" style:font-size-asian="${d.fontSize || 11}pt" style:font-size-complex="${d.fontSize || 11}pt"/></style:default-style><style:default-style style:family="table"><style:table-properties table:border-model="collapsing"/></style:default-style>${Object.values(common).join('')}<style:style style:name="Graphics" style:family="graphic"><style:graphic-properties text:anchor-type="paragraph" svg:x="0cm" svg:y="0cm" style:wrap="dynamic" style:number-wrapped-paragraphs="no-limit" style:wrap-contour="false" style:vertical-pos="top" style:vertical-rel="paragraph" style:horizontal-pos="center" style:horizontal-rel="paragraph"/></style:style><style:style style:name="Index_20_Link" style:display-name="Index Link" style:family="text"/><text:outline-style style:name="Outline">${[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((l) => `<text:outline-level-style text:level="${l}" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style>`).join('')}</text:outline-style><text:notes-configuration text:note-class="footnote" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/></office:styles>`;
    const autoS = `<office:automatic-styles><style:page-layout style:name="OfPm1"><style:page-layout-properties ${layoutProps}>${colsXml}</style:page-layout-properties><style:header-style><style:header-footer-properties fo:min-height="0cm" fo:margin-bottom="0.3cm"/></style:header-style><style:footer-style><style:header-footer-properties fo:min-height="0cm" fo:margin-top="0.3cm"/></style:footer-style></style:page-layout>${hfStyles}</office:automatic-styles>`;
    const masterS = `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="OfPm1">${hfMaster}</style:master-page>${firstPageMaster}</office:master-styles>`;
    stylesXml = { stylesBody, autoS, masterS };
  }

  // ---------- meta ----------
  const meta = model.meta || {};
  const now = new Date().toISOString().replace(/\.\d+Z$/, '');
  const metaBody = `<meta:generator>Hostwerk Office</meta:generator>${meta.title ? `<dc:title>${esc(meta.title)}</dc:title>` : ''}${meta.subject ? `<dc:subject>${esc(meta.subject)}</dc:subject>` : ''}${meta.author ? `<meta:initial-creator>${esc(meta.author)}</meta:initial-creator><dc:creator>${esc(meta.author)}</dc:creator>` : ''}<meta:creation-date>${esc(meta.created || now)}</meta:creation-date><dc:date>${now}</dc:date>${meta.keywords ? String(meta.keywords).split(/\s*,\s*/).filter(Boolean).map((k) => `<meta:keyword>${esc(k)}</meta:keyword>`).join('') : ''}`;
  const mime = MIME[ext] || MIME.odt;

  if (flat) {
    const s = typeof stylesXml === 'string' ? null : stylesXml;
    const origOffStyles = patch ? pick(keep.flat, 'office:styles') : '';
    const origMaster = patch ? pick(keep.flat, 'office:master-styles') : '';
    const fs = s || { stylesBody: '', autoS: '', masterS: '' };
    const stylesPart = origOffStyles ? origOffStyles.replace('</office:styles>', `${Object.entries(common).filter(([n]) => !origOffStyles.includes(`style:name="${n}"`)).map(([, x]) => x).join('')}</office:styles>`) : fs.stylesBody;
    const pageLayout = `<style:page-layout style:name="OfPm1"><style:page-layout-properties ${layoutProps}>${colsXml}</style:page-layout-properties></style:page-layout>${hfStyles}`;
    const xml = `${XMLH}<office:document ${nsAttrs(ALL)} office:version="1.3" office:mimetype="${mime}"><office:meta>${metaBody}</office:meta>${fontDecls()}${stylesPart}<office:automatic-styles>${origAuto}${auto.join('')}${pageLayout}</office:automatic-styles>${origMaster && !(model.header && model.header.text) && !(model.footer && model.footer.text) ? origMaster.replace(/style:page-layout-name="[^"]*"/, 'style:page-layout-name="OfPm1"') : `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="OfPm1">${hfMaster}</style:master-page>${firstPageMaster}</office:master-styles>`}${bodyXml}</office:document>`;
    return new TextEncoder().encode(xml);
  }

  const contentXml = `${XMLH}<office:document-content ${nsAttrs(ALL)} office:version="1.3">${scripts}${fontDecls()}${autoXml}${bodyXml}</office:document-content>`;
  zip.file('content.xml', contentXml);
  if (typeof stylesXml === 'string') zip.file('styles.xml', stylesXml);
  else zip.file('styles.xml', `${XMLH}<office:document-styles ${nsAttrs(ALL)} office:version="1.3">${fontDecls()}${stylesXml.stylesBody}${stylesXml.autoS}${stylesXml.masterS}</office:document-styles>`);
  let metaXml = patch && zip.file('meta.xml') ? await zip.file('meta.xml').async('string') : null;
  if (metaXml) {
    metaXml = metaXml.replace(/<dc:date>[^<]*<\/dc:date>/, `<dc:date>${now}</dc:date>`);
    if (meta.title) metaXml = /<dc:title>/.test(metaXml) ? metaXml.replace(/<dc:title>[^<]*<\/dc:title>/, `<dc:title>${esc(meta.title)}</dc:title>`) : metaXml.replace('</office:meta>', `<dc:title>${esc(meta.title)}</dc:title></office:meta>`);
  } else metaXml = `${XMLH}<office:document-meta ${nsAttrs(['office', 'meta', 'dc', 'xlink'])} office:version="1.3"><office:meta>${metaBody}</office:meta></office:document-meta>`;
  zip.file('meta.xml', metaXml);
  for (const p of pictures) zip.file(p.path, p.bytes);
  // manifest
  let man = patch && zip.file('META-INF/manifest.xml') ? await zip.file('META-INF/manifest.xml').async('string') : `${XMLH}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mime}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
  man = man.replace(/(manifest:full-path="\/"[^>]*manifest:media-type=")[^"]*(")/, `$1${mime}$2`).replace(/(manifest:media-type=")[^"]*("[^>]*manifest:full-path="\/")/, `$1${mime}$2`);
  for (const p of pictures) if (!man.includes(`full-path="${p.path}"`)) man = man.replace('</manifest:manifest>', `<manifest:file-entry manifest:full-path="${p.path}" manifest:media-type="${p.mime}"/></manifest:manifest>`);
  zip.file('META-INF/manifest.xml', man);
  // mimetype first, stored
  const outZip = new JSZip();
  outZip.file('mimetype', mime, { compression: 'STORE' });
  for (const n of Object.keys(zip.files)) { if (n === 'mimetype' || zip.files[n].dir) continue; outZip.file(n, zip.files[n].async('uint8array'), { binary: true, createFolders: false }); }
  return outZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: mime });
}

function b64(u8) {
  if (typeof Buffer !== 'undefined') return Buffer.from(u8).toString('base64');
  let s = ''; for (let k = 0; k < u8.length; k += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(k, k + 0x8000));
  return btoa(s);
}
