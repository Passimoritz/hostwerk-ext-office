/*
 * docx / docm / dotx / dotm writer (own WordprocessingML writer).
 *
 * Lossless: when the document came from a Word file (model.keep.docx), the ORIGINAL package is patched – only
 * word/document.xml, comments, footnotes, numbering (appended), styles (missing ones appended), settings (two flags),
 * core properties and – if you changed them – header/footer are touched; every other part (vbaProject.bin, customXml,
 * embeddings, SmartArt, charts, themes, fonts, glossary …) stays byte-identical. Unknown paragraph/run properties and
 * `preserved` nodes are written back verbatim. Without an original a complete fresh package is built.
 */
import JSZip from 'jszip';
import { parseXml, kids, serialize, esc, escAttr } from '../../_shared/xmltree.js';
import { pageBox, mmToTwip, firstFont, ptOf, hexColor, parseDataUrl, imageSize, walk, TEXTBLOCKS, fillFields, headingsOf, hfKey, hfHas } from './pm.js';
import { equationOmml } from './equation.js';
import { mathmlToOmml } from './mathml.js';
import { paraPropsXml, runPropsXml } from './wordprops.js';
import { bordersXml, marginsXml, lookXml, tableStyleXml, builtinTableStyles, TBL_BORDER_ORDER, TC_BORDER_ORDER } from './tables.js';

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
};
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  docm: 'application/vnd.ms-word.document.macroEnabled.main+xml',
  dotm: 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  commentsExtended: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
};
const RT = {
  styles: `${REL}/styles`, settings: `${REL}/settings`, numbering: `${REL}/numbering`, footnotes: `${REL}/footnotes`, comments: `${REL}/comments`,
  commentsExtended: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', header: `${REL}/header`, footer: `${REL}/footer`,
  image: `${REL}/image`, hyperlink: `${REL}/hyperlink`, officeDocument: `${REL}/officeDocument`, endnotes: `${REL}/endnotes`,
  vbaProject: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
  core: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', app: `${REL}/extended-properties`,
};
const XMLH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const HL_NAMES = { FFFF00: 'yellow', '00FF00': 'green', '00FFFF': 'cyan', FF00FF: 'magenta', '0000FF': 'blue', FF0000: 'red', '000080': 'darkBlue', '008080': 'darkCyan', '008000': 'darkGreen', '800080': 'darkMagenta', '800000': 'darkRed', '808000': 'darkYellow', '808080': 'darkGray', C0C0C0: 'lightGray', '000000': 'black' };
const RPR_ORDER = ['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo', 'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath', 'w:rPrChange'];
const PPR_ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'];
const TBLPR_ORDER = ['w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual', 'w:tblStyleRowBandSize', 'w:tblStyleColBandSize', 'w:tblW', 'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders', 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription', 'w:tblPrChange'];
const TRPR_ORDER = ['w:cnfStyle', 'w:divId', 'w:gridBefore', 'w:gridAfter', 'w:wBefore', 'w:wAfter', 'w:cantSplit', 'w:trHeight', 'w:tblHeader', 'w:tblCellSpacing', 'w:jc', 'w:hidden', 'w:ins', 'w:del', 'w:trPrChange'];
const TCPR_ORDER = ['w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders', 'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText', 'w:vAlign', 'w:hideMark', 'w:headers', 'w:cellIns', 'w:cellDel', 'w:cellMerge', 'w:tcPrChange'];
const SECT_ORDER = ['w:headerReference', 'w:footerReference', 'w:footnotePr', 'w:endnotePr', 'w:type', 'w:pgSz', 'w:pgMar', 'w:paperSrc', 'w:pgBorders', 'w:lnNumType', 'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign', 'w:noEndnote', 'w:titlePg', 'w:textDirection', 'w:bidi', 'w:rtlGutter', 'w:docGrid', 'w:printerSettings', 'w:sectPrChange'];
const ordered = (map, order) => {
  const entries = [...map.entries()];
  const idx = (n) => { const k = order.indexOf(n); return k < 0 ? order.length - 1.5 : k; };
  return entries.sort((a, b) => idx(a[0]) - idx(b[0])).map(([, x]) => (Array.isArray(x) ? x.join('') : x)).join('');
};
const t = (s) => `<w:t xml:space="preserve">${esc(s)}</w:t>`;
const isoNow = () => new Date().toISOString().replace(/\.\d+Z$/, 'Z');

/** style definitions we may need (name → xml builder with the real Normal id) */
function styleDefs(normalId, font) {
  const head = (l, sz, color, before, after, extra = '') => `<w:style w:type="paragraph" w:styleId="Heading${l}"><w:name w:val="heading ${l}"/><w:basedOn w:val="${normalId}"/><w:next w:val="${normalId}"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="${after}"/><w:outlineLvl w:val="${l - 1}"/></w:pPr><w:rPr>${font ? `<w:rFonts w:ascii="${escAttr(font)}" w:hAnsi="${escAttr(font)}"/>` : ''}${extra}<w:color w:val="${color}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`;
  return {
    'heading 1': head(1, 32, '2F5496', 360, 80),
    'heading 2': head(2, 26, '2F5496', 160, 80),
    'heading 3': head(3, 24, '1F3763', 160, 80),
    'heading 4': head(4, 22, '2F5496', 80, 40, '<w:i/>'),
    'heading 5': head(5, 22, '2F5496', 80, 40),
    'heading 6': head(6, 22, '1F3763', 40, 0),
    title: `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="${normalId}"/><w:next w:val="${normalId}"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>`,
    subtitle: `<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="${normalId}"/><w:next w:val="${normalId}"/><w:uiPriority w:val="11"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:color w:val="595959"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>`,
    quote: `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="${normalId}"/><w:next w:val="${normalId}"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>`,
    'toc heading': `<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="${normalId}"/><w:next w:val="${normalId}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>`,
    'toc 1': `<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="${normalId}"/><w:next w:val="${normalId}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="60"/></w:pPr></w:style>`,
    'toc 2': `<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="${normalId}"/><w:next w:val="${normalId}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="60"/><w:ind w:left="220"/></w:pPr></w:style>`,
    'toc 3': `<w:style w:type="paragraph" w:styleId="TOC3"><w:name w:val="toc 3"/><w:basedOn w:val="${normalId}"/><w:next w:val="${normalId}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="60"/><w:ind w:left="440"/></w:pPr></w:style>`,
    'list paragraph': `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="${normalId}"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>`,
  };
}
const STYLE_ID = { 'heading 1': 'Heading1', 'heading 2': 'Heading2', 'heading 3': 'Heading3', 'heading 4': 'Heading4', 'heading 5': 'Heading5', 'heading 6': 'Heading6', title: 'Title', subtitle: 'Subtitle', quote: 'Quote', 'list paragraph': 'ListParagraph', 'toc heading': 'TOCHeading', 'toc 1': 'TOC1', 'toc 2': 'TOC2', 'toc 3': 'TOC3' };

export async function writeDocx(model, { ext = 'docx' } = {}) {
  const keep = model.keep && model.keep.docx;
  const patch = !!(keep && keep.zip);
  const zip = patch ? await JSZip.loadAsync(keep.zip) : new JSZip();
  const docPath = patch ? keep.docPath : 'word/document.xml';
  const dir = docPath.replace(/[^/]+$/, '');
  const relsPath = docPath.replace(/([^/]+)$/, '_rels/$1.rels');
  const read = async (p) => { const f = zip.file(p); return f ? f.async('string') : null; };

  // ---- relationships of the main part
  let relsXml = (patch && await read(relsPath)) || `${XMLH}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const relList = kids(parseXml(relsXml), 'Relationship').map((r) => ({ id: r.a.Id, type: r.a.Type, target: r.a.Target, mode: r.a.TargetMode }));
  let ridN = 1;
  const newRels = [];
  // while a header/footer part is written its relationships go to that part's own .rels (scope)
  let scope = null;
  const addRel = (type, target, external = false) => {
    if (scope) { const id = `rIdH${scope.n++}`; scope.rels.push({ id, type, target, mode: external ? 'External' : undefined }); return id; }
    let id; do { id = `rIdOf${ridN++}`; } while (relList.some((r) => r.id === id));
    const rel = { id, type, target, mode: external ? 'External' : undefined };
    relList.push(rel); newRels.push(rel);
    return id;
  };
  const relOf = (type) => relList.find((r) => r.type === type);
  const removed = new Set();
  const removeRel = (rel) => { removed.add(rel.id); const k = relList.indexOf(rel); if (k >= 0) relList.splice(k, 1); };
  const partOf = (rel) => (rel.target.startsWith('/') ? rel.target.slice(1) : dir + rel.target);

  // ---- content types
  let ctXml = (patch && await read('[Content_Types].xml')) || `${XMLH}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`;
  const ensureDefault = (e, type) => { if (!new RegExp(`Extension="${e}"`, 'i').test(ctXml)) ctXml = ctXml.replace('</Types>', `<Default Extension="${e}" ContentType="${type}"/></Types>`); };
  const ensureOverride = (part, type) => { if (!ctXml.includes(`PartName="/${part}"`)) ctXml = ctXml.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`); };
  const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const dropOverride = (part) => { ctXml = ctXml.replace(new RegExp(`<Override[^>]*PartName="/${reEsc(part)}"[^>]*/>`, 'i'), ''); };
  /** set (or replace) the content type of one part */
  const setOverride = (part, type) => {
    const re = new RegExp(`<Override\\b[^>]*PartName="/${reEsc(part)}"[^>]*/>`, 'i');
    if (re.test(ctXml)) ctXml = ctXml.replace(re, `<Override PartName="/${part}" ContentType="${type}"/>`);
    else ensureOverride(part, type);
  };

  // ---- styles: find ids by name, append missing ones
  let stylesRel = relOf(RT.styles);
  let stylesXml = stylesRel ? await read(partOf(stylesRel)) : null;
  const idByName = {}; let normalId = 'Normal';
  const styleIds = new Set(); const usedTableStyles = new Set(); let tableNormalId = null;
  if (stylesXml) {
    for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
      const id = /w:styleId="([^"]*)"/.exec(m[1]); const nm = /<w:name w:val="([^"]*)"/.exec(m[2]);
      if (id) styleIds.add(id[1]);
      if (id && /w:default="1"/.test(m[1]) && /w:type="table"/.test(m[1])) tableNormalId = id[1];
      if (id && nm) idByName[nm[1].toLowerCase()] = id[1];
      if (id && /w:default="1"/.test(m[1]) && /w:type="paragraph"/.test(m[1])) normalId = id[1];
    }
  }
  const defs = styleDefs(normalId, null);
  const missingStyles = new Set();
  const styleId = (name) => { if (idByName[name]) return idByName[name]; missingStyles.add(name); return STYLE_ID[name]; };

  // ---- numbering
  const absDefs = new Map(); // kind → abstractNumId
  const numDefs = []; // { numId, abs, start }
  let numN = 9100;
  const LIST_FMT = { bullet: null, decimal: 'decimal', 'lower-alpha': 'lowerLetter', 'upper-alpha': 'upperLetter', 'lower-roman': 'lowerRoman', 'upper-roman': 'upperRoman', outline: 'decimal' };
  const absFor = (kind) => { if (!absDefs.has(kind)) absDefs.set(kind, 9001 + absDefs.size); return absDefs.get(kind); };
  const newNum = (kind, start, level = 0) => { const numId = numN++; numDefs.push({ numId, abs: absFor(kind), start, level }); return numId; };
  const absXml = (kind, id) => {
    let lv = '';
    for (let l = 0; l < 9; l++) {
      const ind = `<w:pPr><w:ind w:left="${720 + l * 360}" w:hanging="360"/></w:pPr>`;
      if (kind === 'bullet') lv += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${['•', '◦', '▪'][l % 3]}"/><w:lvlJc w:val="left"/>${ind}<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/></w:rPr></w:lvl>`;
      else lv += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${LIST_FMT[kind] || 'decimal'}"/><w:lvlText w:val="${kind === 'outline' ? Array.from({ length: l + 1 }, (_, k) => `%${k + 1}.`).join('') : `%${l + 1}.`}"/><w:lvlJc w:val="left"/>${ind}</w:lvl>`;
    }
    return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="${kind === 'bullet' ? 'hybridMultilevel' : 'multilevel'}"/>${lv}</w:abstractNum>`;
  };

  // ---- media
  const mediaRid = new Map(); let mediaN = 1; let drawingId = 5000;
  const imageRid = (src0) => {
    const src = src0; const key = `${scope ? scope.part : ''}|${src0}`;
    if (mediaRid.has(key)) return mediaRid.get(key);
    const d = parseDataUrl(src);
    let res = null;
    if (d && ['png', 'jpg', 'gif', 'bmp'].includes(d.ext)) {
      let name; do { name = `media/of_image${mediaN++}.${d.ext}`; } while (zip.file(dir + name));
      zip.file(dir + name, d.bytes);
      ensureDefault(d.ext, d.mime);
      if (d.ext === 'jpg') ensureDefault('jpeg', 'image/jpeg');
      res = { rid: addRel(RT.image, name), size: imageSize(d.bytes) };
    } else if (!d && /^https?:/i.test(src)) res = { rid: addRel(RT.image, src, true), size: null, link: true };
    mediaRid.set(key, res);
    return res;
  };
  const hyperRid = new Map();
  const linkRid = (href) => { const key = `${scope ? scope.part : ''}|${href}`; if (!hyperRid.has(key)) hyperRid.set(key, addRel(RT.hyperlink, href, true)); return hyperRid.get(key); };

  // ---- comments: ids and ranges (first/last inline node index per comment)
  const comments = (model.comments || []).map((c, k) => ({ ...c, wid: k }));
  const byId = new Map(comments.map((c) => [String(c.id), c]));
  const range = new Map(); // id → { first, last }
  let counter = 0;
  walk(model.doc, (n, parent) => {
    if (parent && TEXTBLOCKS.has(parent.type)) {
      for (const m of n.marks || []) if (m.type === 'comment' && byId.has(String(m.attrs.id))) { const r = range.get(String(m.attrs.id)) || { first: counter, last: counter }; r.last = counter; range.set(String(m.attrs.id), r); }
      counter++;
    }
  });
  let replyId = comments.length;
  for (const c of comments) c.replyIds = (c.replies || []).map(() => replyId++);
  const used = comments.filter((c) => range.has(String(c.id)));
  const startsAt = new Map(); const endsAt = new Map();
  for (const c of used) { const r = range.get(String(c.id)); (startsAt.get(r.first) || startsAt.set(r.first, []).get(r.first)).push(c); (endsAt.get(r.last) || endsAt.set(r.last, []).get(r.last)).push(c); }

  // ---- footnotes
  const footnotes = [];
  let revN = 1000;
  const ix = { n: 0 };

  // ---- table of contents: written with its current entries filled in (field result), each linked to a hidden
  // _Toc bookmark on its heading – no "dirty" flag / updateFields, so Word opens it without the "update fields?" prompt
  const allHeadings = headingsOf(model.doc, 9);
  const tocEntries = (a) => {
    const list = a.entries && a.entries.length ? a.entries : allHeadings.filter((h) => h.level <= 3);
    const taken = new Set();
    return list.map((e) => {
      let index = e.index != null && allHeadings.some((h) => h.index === e.index) ? e.index : null;
      if (index == null) { const h = allHeadings.find((x) => x.text === String(e.text || '').trim() && !taken.has(x.index)); if (h) index = h.index; }
      if (index != null) taken.add(index);
      return { level: e.level || 1, text: e.text || '', index };
    });
  };
  const tocTargets = new Set();
  walk(model.doc, (n) => { if (n.type === 'toc') for (const e of tocEntries(n.attrs || {})) if (e.index != null) tocTargets.add(e.index); });
  const tocBookmark = (index) => `_Toc${String(index + 1).padStart(9, '0')}`;
  let headingN = 0; let bookmarkN = 90000;

  /* ============ inline content ============ */
  const rPrOf = (marks, extra = {}) => {
    const map = new Map();
    let link = null; let ins = null; let del = null; let sdt = null;
    for (const m of marks || []) {
      const a = m.attrs || {};
      switch (m.type) {
        case 'runKeep': for (const [nm, x] of a.ox || []) map.set(nm, x); break;
        case 'bold': map.set('w:b', '<w:b/>'); map.set('w:bCs', '<w:bCs/>'); break;
        case 'italic': map.set('w:i', '<w:i/>'); map.set('w:iCs', '<w:iCs/>'); break;
        case 'underline': map.set('w:u', '<w:u w:val="single"/>'); break;
        case 'strike': map.set('w:strike', '<w:strike/>'); break;
        case 'superscript': map.set('w:vertAlign', '<w:vertAlign w:val="superscript"/>'); break;
        case 'subscript': map.set('w:vertAlign', '<w:vertAlign w:val="subscript"/>'); break;
        case 'code': map.set('w:rFonts', '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>'); break;
        case 'textStyle': {
          if (a.fontFamily) { const f = escAttr(firstFont(a.fontFamily)); map.set('w:rFonts', `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`); }
          const pt = ptOf(a.fontSize); if (pt) { map.set('w:sz', `<w:sz w:val="${Math.round(pt * 2)}"/>`); map.set('w:szCs', `<w:szCs w:val="${Math.round(pt * 2)}"/>`); }
          const c = hexColor(a.color); if (c) map.set('w:color', `<w:color w:val="${c}"/>`);
          for (const [k, v] of runPropsXml(a)) map.set(k, v);
          break;
        }
        case 'highlight': { const c = hexColor(a.color || '#FFFF00') || 'FFFF00'; if (HL_NAMES[c]) map.set('w:highlight', `<w:highlight w:val="${HL_NAMES[c]}"/>`); else map.set('w:shd', `<w:shd w:val="clear" w:color="auto" w:fill="${c}"/>`); break; }
        case 'link': link = a.href; break;
        case 'insertion': ins = a; break;
        case 'deletion': del = a; break;
        case 'sdt': sdt = a; break;
        case 'formatChange': map.set('w:rPrChange', `<w:rPrChange w:id="${revN++}" w:author="${escAttr(a.author || '')}"${a.date ? ` w:date="${escAttr(a.date)}"` : ''}>${a.old || '<w:rPr/>'}</w:rPrChange>`); break;
        default:
      }
    }
    if (link && !map.has('w:color')) map.set('w:color', '<w:color w:val="0563C1"/>');
    if (link && !map.has('w:u')) map.set('w:u', '<w:u w:val="single"/>');
    if (extra.sup) map.set('w:vertAlign', '<w:vertAlign w:val="superscript"/>');
    return { rpr: map.size ? `<w:rPr>${ordered(map, RPR_ORDER)}</w:rPr>` : '', link, ins, del, sdt };
  };
  const textRuns = (text, rpr, deleted) => {
    let s = '';
    for (const part of String(text).split(/(\t|\n)/)) {
      if (!part) continue;
      if (part === '\t') s += `<w:r>${rpr}<w:tab/></w:r>`;
      else if (part === '\n') s += `<w:r>${rpr}<w:br/></w:r>`;
      else s += `<w:r>${rpr}${deleted ? `<w:delText xml:space="preserve">${esc(part)}</w:delText>` : t(part)}</w:r>`;
    }
    return s;
  };
  const drawingXml = (a, rid, size) => {
    const w = a.width || (size ? Math.min(size.width, 600) : 300);
    const h = a.height && a.width ? a.height : size ? Math.round(w * (size.height / size.width)) : Math.round(w * 0.66);
    const cx = Math.round(w * 9525); const cy = Math.round(h * 9525);
    const id = drawingId++;
    const graphic = `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"${a.alt ? ` descr="${escAttr(a.alt)}"` : ''}/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip ${rid.link ? 'r:link' : 'r:embed'}="${rid.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
    const docPr = `<wp:docPr id="${id}" name="Picture ${id}"${a.alt ? ` descr="${escAttr(a.alt)}"` : ''}${a.title ? ` title="${escAttr(a.title)}"` : ''}/>`;
    if (a.wrap === 'square' && (a.align === 'left' || a.align === 'right')) {
      return `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>${a.align}</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>${docPr}<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>${graphic}</wp:anchor></w:drawing>`;
    }
    return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${docPr}<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>${graphic}</wp:inline></w:drawing>`;
  };

  /* ---- content controls (check box, drop-down, date, plain text) ---- */
  let sdtId = 700000;
  const controlXml = (a, rpr) => {
    const kind = a.kind || 'text';
    let pr = a.pr || '';
    if (!pr) {
      const type = kind === 'checkbox' ? `<w14:checkbox><w14:checked w14:val="${a.checked ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>`
        : kind === 'dropdown' || kind === 'combo' ? `<w:${kind === 'combo' ? 'comboBox' : 'dropDownList'}>${(a.items || []).map((it) => `<w:listItem w:displayText="${escAttr(it.text)}" w:value="${escAttr(it.value != null ? it.value : it.text)}"/>`).join('')}</w:${kind === 'combo' ? 'comboBox' : 'dropDownList'}>`
          : kind === 'date' ? `<w:date${a.fullDate ? ` w:fullDate="${escAttr(a.fullDate)}"` : ''}><w:dateFormat w:val="${escAttr(a.format || 'dd.MM.yyyy')}"/><w:lid w:val="de-DE"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>` : '<w:text/>';
      pr = `<w:sdtPr>${a.alias ? `<w:alias w:val="${escAttr(a.alias)}"/>` : ''}${a.tag ? `<w:tag w:val="${escAttr(a.tag)}"/>` : ''}<w:id w:val="${sdtId++}"/>${!a.value && kind !== 'checkbox' ? '<w:showingPlcHdr/>' : ''}${type}</w:sdtPr>`;
    } else {
      if (kind === 'checkbox') pr = pr.replace(/(<w14:checked\b[^>]*w14:val=")[^"]*(")/, `$1${a.checked ? 1 : 0}$2`);
      if (kind === 'date' && a.fullDate) pr = /w:fullDate="/.test(pr) ? pr.replace(/w:fullDate="[^"]*"/, `w:fullDate="${escAttr(a.fullDate)}"`) : pr.replace(/<w:date\b/, `<w:date w:fullDate="${escAttr(a.fullDate)}"`);
      if (a.value && !a.placeholder) pr = pr.replace(/<w:showingPlcHdr\s*\/>/, '');
    }
    const body = kind === 'checkbox'
      ? `<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr>${t(a.checked ? '☒' : '☐')}</w:r>`
      : `<w:r>${rpr}${t(a.value || (a.placeholder ? '' : ' '))}</w:r>`;
    return `<w:sdt>${pr}${a.end || ''}<w:sdtContent>${body}</w:sdtContent></w:sdt>`;
  };
  /* ---- fields: Word's own (fldSimple, cached result) and Hostwerk's live values (a locked plain-text control) ---- */
  const FIELD_INSTR = { page: 'PAGE', pages: 'NUMPAGES', date: 'DATE \\@ "dd.MM.yyyy"', time: 'TIME \\@ "HH:mm"', author: 'AUTHOR', title: 'TITLE', filename: 'FILENAME', subject: 'SUBJECT', words: 'NUMWORDS' };
  const fieldXml = (a, rpr) => {
    const value = a.value == null ? '' : String(a.value);
    if (a.ox) {
      const was = [...String(a.ox).matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
      if (was === value) return a.ox;
    }
    if (a.kind === 'hostwerk' || a.kind === 'cell') {
      const tag = `hostwerk-field:${JSON.stringify({ k: a.kind, s: a.source || null, p: a.params || null, l: a.label || '' })}`;
      return `<w:sdt><w:sdtPr>${rpr}<w:alias w:val="${escAttr(a.label || 'Hostwerk')}"/><w:tag w:val="${escAttr(tag)}"/><w:id w:val="${sdtId++}"/><w:lock w:val="contentLocked"/><w:text/></w:sdtPr><w:sdtContent><w:r>${rpr}${t(value || ' ')}</w:r></w:sdtContent></w:sdt>`;
    }
    const instr = a.instr || FIELD_INSTR[a.kind] || 'DATE';
    return `<w:fldSimple w:instr=" ${escAttr(instr.trim())} "><w:r>${rpr}${t(value || ' ')}</w:r></w:fldSimple>`;
  };
  /** a kept text box whose text was edited here: new paragraphs in every w:txbxContent (Word + VML fallback) */
  const textboxEdited = (xml, text) => xml.replace(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/g, (m, inner) => {
    const ppr = (/<w:pPr>[\s\S]*?<\/w:pPr>/.exec(inner) || [''])[0].replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, '');
    const rpr = (/<w:r>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(inner) || [null, ''])[1];
    return `<w:txbxContent>${String(text).split('\n').map((line) => `<w:p>${ppr}${line ? `<w:r>${rpr}${t(line)}</w:r>` : ''}</w:p>`).join('')}</w:txbxContent>`;
  });

  /** inline nodes → paragraph content XML */
  const inlineXml = (content, base = {}) => {
    const items = []; // { xml, link, ins, del, marker }
    for (const n of content || []) {
      const at = ix.n++;
      for (const c of startsAt.get(at) || []) items.push({ marker: true, xml: `<w:commentRangeStart w:id="${c.wid}"/>${c.replyIds.map((r) => `<w:commentRangeStart w:id="${r}"/>`).join('')}` });
      const marks = [...(base.marks || []), ...(n.marks || [])];
      if (n.type === 'text') {
        const { rpr, link, ins, del, sdt } = rPrOf(marks);
        items.push({ xml: textRuns(n.text, rpr, !!del), link, ins, del, sdt });
      } else if (n.type === 'hardBreak') items.push({ xml: '<w:r><w:br/></w:r>' });
      else if (n.type === 'contentControl') { const r = rPrOf(marks); items.push({ xml: controlXml(n.attrs || {}, r.rpr), ins: r.ins, del: r.del }); }
      else if (n.type === 'field') { const r = rPrOf(marks); items.push({ xml: fieldXml(n.attrs || {}, r.rpr), ins: r.ins, del: r.del }); }
      else if (n.type === 'diagram') {
        const a = n.attrs || {};
        const r = a.src && imageRid(a.src);
        if (r) items.push({ xml: `<w:r>${drawingXml({ width: a.width, height: a.height, alt: a.source || '', title: `hostwerk-diagram:${a.kind || 'mermaid'}` }, r, r.size)}</w:r>` });
        else if (a.source) items.push({ xml: textRuns(a.source, '', false) });
      }
      else if (n.type === 'image') {
        const a = n.attrs || {};
        if (patch && a.ox && a.ox.xml && a.ox.src === (a.src || '').length && a.ox.width === a.width) items.push({ xml: a.ox.xml });
        else { const r = a.src && imageRid(a.src); if (r) items.push({ xml: `<w:r>${drawingXml(a, r, r.size)}</w:r>` }); }
      } else if (n.type === 'footnote') {
        const id = footnotes.length + 1;
        footnotes.push(n.attrs || {});
        items.push({ xml: `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>` });
      } else if (n.type === 'equation') {
        const a = n.attrs || {};
        // unchanged Word maths stays byte-identical; edited/new: MathML (temml) → OMML, else the simple converter
        let x = a.omml && a.ommlLatex === a.latex ? a.omml : a.mathml ? mathmlToOmml(a.mathml) : equationOmml(a.latex || '');
        if (a.display && !/^<m:oMathPara/.test(x)) x = `<m:oMathPara>${x}</m:oMathPara>`;
        if (!a.display && /^<m:oMathPara/.test(x)) x = x.replace(/^<m:oMathPara>(?:<m:oMathParaPr>[\s\S]*?<\/m:oMathParaPr>)?/, '').replace(/<\/m:oMathPara>$/, '');
        items.push({ xml: x });
      } else if (n.type === 'preserved') {
        const a = n.attrs || {};
        if (patch && a.format === 'docx' && a.xml && a.textEdited) items.push({ xml: textboxEdited(a.xml, a.text || '') });
        else if (patch && a.format === 'docx' && a.xml) items.push({ xml: a.xml });
        else if (a.src && imageRid(a.src)) { const r = imageRid(a.src); items.push({ xml: `<w:r>${drawingXml({ width: a.width }, r, r.size)}</w:r>` }); }
        else if (a.text) items.push({ xml: textRuns(a.text, '', false) });
      }
      for (const c of endsAt.get(at) || []) items.push({ marker: true, xml: `<w:commentRangeEnd w:id="${c.wid}"/>${c.replyIds.map((r) => `<w:commentRangeEnd w:id="${r}"/>`).join('')}<w:r><w:commentReference w:id="${c.wid}"/></w:r>${c.replyIds.map((r) => `<w:r><w:commentReference w:id="${r}"/></w:r>`).join('')}` });
    }
    // group: content control (w:sdt) > hyperlink > ins/del > runs
    let s = '';
    let k = 0;
    const nextSdt = (j) => { while (j < items.length && items[j].marker) j++; return j < items.length ? items[j].sdt : null; };
    while (k < items.length) {
      const sd = items[k].marker ? null : items[k].sdt;
      if (!sd) { const inner = []; while (k < items.length && !(items[k].sdt && !items[k].marker)) inner.push(items[k++]); s += linkGroups(inner); continue; }
      const inner = [];
      // comment markers between two runs of the same control stay inside it
      while (k < items.length && ((!items[k].marker && items[k].sdt && items[k].sdt.id === sd.id) || (items[k].marker && (nextSdt(k) || {}).id === sd.id))) inner.push(items[k++]);
      s += `<w:sdt>${sd.pr || '<w:sdtPr/>'}${sd.end || ''}<w:sdtContent>${linkGroups(inner)}</w:sdtContent></w:sdt>`;
    }
    return s;
  };
  const linkGroups = (items) => {
    let s = '';
    let k = 0;
    while (k < items.length) {
      const it = items[k];
      if (it.link) {
        let inner = [];
        const href = it.link;
        while (k < items.length && items[k].link === href) inner.push(items[k++]);
        s += `<w:hyperlink ${href.startsWith('#') ? `w:anchor="${escAttr(href.slice(1))}"` : `r:id="${linkRid(href)}"`} w:history="1">${groupRev(inner)}</w:hyperlink>`;
        inner = null;
        continue;
      }
      const inner = [];
      while (k < items.length && !items[k].link) inner.push(items[k++]);
      s += groupRev(inner);
    }
    return s;
  };
  const groupRev = (items) => {
    let s = ''; let k = 0;
    while (k < items.length) {
      const it = items[k];
      const rev = it.ins || it.del;
      if (!rev || it.marker) { s += it.xml; k++; continue; }
      const kind = it.ins ? 'ins' : 'del';
      const key = `${kind}:${rev.id}:${rev.author}`;
      let inner = '';
      while (k < items.length && !items[k].marker && (items[k].ins || items[k].del) && `${items[k].ins ? 'ins' : 'del'}:${(items[k].ins || items[k].del).id}:${(items[k].ins || items[k].del).author}` === key) inner += items[k++].xml;
      const tag = rev.move ? (kind === 'ins' ? 'w:moveTo' : 'w:moveFrom') : `w:${kind}`;
      s += `<${tag} w:id="${revN++}" w:author="${escAttr(rev.author || 'Unknown')}"${rev.date ? ` w:date="${escAttr(rev.date)}"` : ''}>${inner}</${tag}>`;
    }
    return s;
  };

  /* ============ blocks ============ */
  const pPrOf = (a = {}, opts = {}) => {
    const map = new Map();
    for (const [nm, x] of (a.ox && a.ox.ppr) || []) map.set(nm, x);
    if (opts.style) map.set('w:pStyle', `<w:pStyle w:val="${escAttr(opts.style)}"/>`);
    else if (a.styleId) map.set('w:pStyle', `<w:pStyle w:val="${escAttr(a.styleId)}"/>`);
    if (opts.num) map.set('w:numPr', `<w:numPr><w:ilvl w:val="${opts.num.level}"/><w:numId w:val="${opts.num.id}"/></w:numPr>`);
    if (a.spaceBefore != null || a.spaceAfter != null || a.lineHeight || (a.lineRule && a.lineExact) || opts.spacing) {
      const sp = { ...(opts.spacing || {}) };
      if (a.spaceBefore != null) sp.before = Math.round(a.spaceBefore * 20);
      if (a.spaceAfter != null) sp.after = Math.round(a.spaceAfter * 20);
      if (a.lineRule && a.lineExact) { sp.line = Math.round(a.lineExact * 20); sp.lineRule = a.lineRule; } else if (a.lineHeight) { sp.line = Math.round(a.lineHeight * 240); sp.lineRule = 'auto'; }
      map.set('w:spacing', `<w:spacing${sp.before != null ? ` w:before="${sp.before}"` : ''}${sp.after != null ? ` w:after="${sp.after}"` : ''}${sp.line ? ` w:line="${sp.line}" w:lineRule="${sp.lineRule || 'auto'}"` : ''}/>`);
    }
    const left = (a.indent || 0) * 20 + (opts.indent || 0);
    for (const [k, v] of paraPropsXml(a)) map.set(k, v);
    if (left || a.firstLine || a.rightIndent) map.set('w:ind', `<w:ind w:left="${Math.round(left)}"${a.rightIndent ? ` w:right="${Math.round(a.rightIndent * 20)}"` : ''}${a.firstLine < 0 ? ` w:hanging="${Math.round(-a.firstLine * 20)}"` : a.firstLine ? ` w:firstLine="${Math.round(a.firstLine * 20)}"` : ''}/>`);
    if (a.textAlign && a.textAlign !== 'left') map.set('w:jc', `<w:jc w:val="${a.textAlign === 'justify' ? 'both' : a.textAlign}"/>`);
    for (const [k, v] of Object.entries(opts.extra || {})) map.set(k, v);
    // tracked paragraph mark (Enter / joined paragraphs) and tracked paragraph formatting
    if (a.pmark && (a.pmark.type === 'ins' || a.pmark.type === 'del')) {
      const rev = `<w:${a.pmark.type} w:id="${revN++}" w:author="${escAttr(a.pmark.author || 'Unknown')}"${a.pmark.date ? ` w:date="${escAttr(a.pmark.date)}"` : ''}/>`;
      const old = map.get('w:rPr');
      map.set('w:rPr', old ? old.replace(/^<w:rPr>/, `<w:rPr>${rev}`).replace(/^<w:rPr\/>$/, `<w:rPr>${rev}</w:rPr>`) : `<w:rPr>${rev}</w:rPr>`);
    }
    if (a.pchange) map.set('w:pPrChange', `<w:pPrChange w:id="${revN++}" w:author="${escAttr(a.pchange.author || 'Unknown')}"${a.pchange.date ? ` w:date="${escAttr(a.pchange.date)}"` : ''}>${a.pchange.old || '<w:pPr/>'}</w:pPrChange>`);
    return map.size ? `<w:pPr>${ordered(map, PPR_ORDER)}</w:pPr>` : '';
  };
  let out = [];
  let curPage = null; // page setup of the section being written (tables/boxes take its text width)
  const usable = () => { const pg = curPage || model.page; const b = pageBox(pg); const mg = pg.margins || model.page.margins; return mmToTwip(b.width - mg.left - mg.right); };

  const block = (n, ctx = {}) => {
    const a = n.attrs || {};
    switch (n.type) {
      case 'paragraph': {
        const style = ctx.quote ? styleId('quote') : a.pstyle === 'title' ? styleId('title') : a.pstyle === 'subtitle' ? styleId('subtitle') : a.pstyle === 'listparagraph' ? styleId('list paragraph') : ctx.style || null;
        out.push(`<w:p>${pPrOf(a, { style, num: ctx.num, indent: ctx.indent, extra: ctx.pextra })}${ctx.prefix || ''}${inlineXml(n.content, ctx)}</w:p>`);
        break;
      }
      case 'heading': {
        const hIndex = headingN++;
        const mark = tocTargets.has(hIndex) ? { id: bookmarkN++, name: tocBookmark(hIndex) } : null;
        const inner = inlineXml(n.content);
        out.push(`<w:p>${pPrOf(a, { style: styleId(`heading ${Math.min(6, a.level || 1)}`), num: ctx.num, extra: ctx.pextra })}${mark ? `<w:bookmarkStart w:id="${mark.id}" w:name="${mark.name}"/>${inner}<w:bookmarkEnd w:id="${mark.id}"/>` : inner}</w:p>`);
        break;
      }
      case 'blockquote': for (const c of n.content || []) block(c, { ...ctx, quote: true }); break;
      case 'codeBlock': {
        const code = (n.content || []).map((c) => { ix.n++; return c.text || ''; }).join('');
        for (const line of code.split('\n')) out.push(`<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F4F4F4"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="${ctx.indent || 0}"/></w:pPr>${line ? `<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/></w:rPr>${t(line)}</w:r>` : ''}</w:p>`);
        break;
      }
      case 'bulletList': case 'orderedList': {
        const level = ctx.listLevel || 0;
        const kind = n.type === 'bulletList' ? 'bullet' : (a.listStyle || 'decimal');
        const numId = patch && a.ox && a.ox.numId ? a.ox.numId : newNum(kind, n.type === 'orderedList' ? (a.start || 1) : 1, level);
        for (const li of n.content || []) {
          (li.content || []).forEach((c, k) => {
            if (k === 0 && c.type === 'paragraph') block(c, { ...ctx, num: { id: numId, level }, style: styleId('list paragraph'), listLevel: level });
            else if (['bulletList', 'orderedList', 'taskList'].includes(c.type)) block(c, { ...ctx, listLevel: level + 1 });
            else block(c, { ...ctx, indent: 720 + level * 360, num: null });
          });
        }
        break;
      }
      case 'taskList': {
        const level = ctx.listLevel || 0;
        for (const li of n.content || []) {
          const checked = !!(li.attrs && li.attrs.checked);
          (li.content || []).forEach((c, k) => {
            if (k === 0 && c.type === 'paragraph') {
              const box = `<w:sdt><w:sdtPr><w14:checkbox><w14:checked w14:val="${checked ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr>${t(checked ? '☒' : '☐')}</w:r></w:sdtContent></w:sdt><w:r>${t(' ')}</w:r>`;
              block(c, { ...ctx, prefix: box, indent: level * 360, num: null });
            } else if (['bulletList', 'orderedList', 'taskList'].includes(c.type)) block(c, { ...ctx, listLevel: level + 1 });
            else block(c, { ...ctx, indent: 360 + level * 360 });
          });
        }
        break;
      }
      case 'horizontalRule': out.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'); break;
      case 'pageBreak': out.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>'); break;
      case 'toc': {
        const es = tocEntries(a);
        out.push('<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>');
        if (a.title) out.push(`<w:p><w:pPr><w:pStyle w:val="${styleId('toc heading')}"/></w:pPr><w:r>${t(a.title)}</w:r></w:p>`);
        // the field result is the entry list our editor shows (no page numbers – Word adds them on "Update field")
        const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
        const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
        if (!es.length) out.push(`<w:p>${begin}<w:r>${t('No table of contents entries found.')}</w:r>${end}</w:p>`);
        es.forEach((e, k) => {
          const run = `<w:r>${t(e.text)}</w:r>`;
          const body = e.index != null ? `<w:hyperlink w:anchor="${tocBookmark(e.index)}" w:history="1">${run}</w:hyperlink>` : run;
          out.push(`<w:p><w:pPr><w:pStyle w:val="${styleId(`toc ${Math.min(3, e.level || 1)}`)}"/></w:pPr>${k === 0 ? begin : ''}${body}${k === es.length - 1 ? end : ''}</w:p>`);
        });
        out.push('</w:sdtContent></w:sdt>');
        break;
      }
      case 'callout': {
        const color = { warn: 'FDF6E6', ok: 'ECF8F2', note: 'F6F6F6' }[a.kind] || 'EEF4FC';
        const border = { warn: 'E0C07A', ok: '8FCFB1', note: 'C8C8C8' }[a.kind] || '9BB7E0';
        const b = (s) => `<w:${s} w:val="single" w:sz="8" w:space="0" w:color="${border}"/>`;
        out.push(`<w:tbl><w:tblPr><w:tblW w:w="${usable()}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right'].map(b).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/><w:tblCaption w:val="of-callout:${escAttr(a.kind || 'info')}"/></w:tblPr><w:tblGrid><w:gridCol w:w="${usable()}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${usable()}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="${color}"/></w:tcPr>`);
        cellContent(n.content);
        out.push('</w:tc></w:tr></w:tbl>');
        break;
      }
      case 'table': table(n); break;
      case 'controlBlock': {
        out.push(`<w:sdt>${a.pr || `<w:sdtPr>${a.alias ? `<w:alias w:val="${escAttr(a.alias)}"/>` : ''}<w:id w:val="${sdtId++}"/></w:sdtPr>`}${a.end || ''}<w:sdtContent>`);
        const inner = n.content && n.content.length ? n.content : [{ type: 'paragraph' }];
        for (const c of inner) block(c, ctx);
        out.push('</w:sdtContent></w:sdt>');
        break;
      }
      // a section break whose paragraph could not carry it (after a table, a list …): an empty paragraph with the sectPr
      case 'sectionBreak': if (!ctx.top) break; out.push(`<w:p><w:pPr>${sectionOf(a)}</w:pPr></w:p>`); break;
      default: for (const c of n.content || []) block(c, ctx);
    }
  };
  const cellContent = (content) => {
    const list = content && content.length ? content : [{ type: 'paragraph' }];
    for (const c of list) block(c, {});
    if (list[list.length - 1].type === 'table') out.push('<w:p/>');
  };
  const table = (n) => {
    const rows = n.content || [];
    const ta = n.attrs || {};
    const grid = [];
    rows.forEach((r, ri) => {
      grid[ri] = grid[ri] || [];
      let ci = 0;
      for (const c of r.content || []) {
        while (grid[ri][ci]) ci++;
        const cs = (c.attrs && c.attrs.colspan) || 1; const rs = (c.attrs && c.attrs.rowspan) || 1;
        for (let y = 0; y < rs; y++) { grid[ri + y] = grid[ri + y] || []; for (let x = 0; x < cs; x++) grid[ri + y][ci + x] = { cell: c, origin: y === 0 && x === 0, cont: y > 0 && x === 0, covered: x > 0, cs, rs }; }
        ci += cs;
      }
    });
    const cols = Math.max(1, ...grid.map((r) => r.length));
    const widths = Array(cols).fill(null);
    for (const r of grid) for (let k = 0; k < cols; k++) { const g = r[k]; if (g && g.origin && g.cell.attrs && g.cell.attrs.colwidth) g.cell.attrs.colwidth.forEach((w, j) => { if (w && widths[k + j] == null) widths[k + j] = Math.round(w * 15); }); }
    const known = widths.filter(Boolean).reduce((s, w) => s + w, 0);
    const free = widths.filter((w) => !w).length;
    const each = free ? Math.max(500, Math.round((usable() - known) / free)) : 0;
    const W2 = widths.map((w) => w || each);
    // tblPr: the original properties, the modelled ones replaced by the attributes (they are the truth)
    const pr = new Map();
    const orig = ta.ox && ta.ox.tblPr ? parseXml(ta.ox.tblPr) : null;
    const MODELED = ['w:tblStyle', 'w:tblW', 'w:jc', 'w:tblInd', 'w:tblBorders', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook'];
    if (orig) for (const c of kids(orig, '*')) if (!MODELED.includes(c.n)) pr.set(c.n, serialize(c));
    const fresh = !orig && !ta.styleId && !ta.borders;
    if (ta.styleId && (patch || builtinTableStyles()[ta.styleId])) { pr.set('w:tblStyle', `<w:tblStyle w:val="${escAttr(ta.styleId)}"/>`); if (!styleIds.has(ta.styleId)) usedTableStyles.add(ta.styleId); }
    const tw = ta.width;
    pr.set('w:tblW', tw && tw.type === 'pct' ? `<w:tblW w:w="${Math.round(tw.w)}" w:type="pct"/>` : tw && tw.type === 'dxa' && tw.w > 0 ? `<w:tblW w:w="${Math.round(tw.w)}" w:type="dxa"/>` : '<w:tblW w:w="0" w:type="auto"/>');
    if (ta.align) pr.set('w:jc', `<w:jc w:val="${ta.align === 'right' ? 'right' : 'center'}"/>`);
    if (ta.indent) pr.set('w:tblInd', `<w:tblInd w:w="${Math.round(ta.indent)}" w:type="dxa"/>`);
    if (ta.borders) pr.set('w:tblBorders', bordersXml('tblBorders', ta.borders, TBL_BORDER_ORDER));
    else if (fresh) pr.set('w:tblBorders', `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="999999"/>`).join('')}</w:tblBorders>`);
    if (ta.layout === 'fixed' || (fresh && !ta.width)) pr.set('w:tblLayout', '<w:tblLayout w:type="fixed"/>');
    if (ta.cellMargins) pr.set('w:tblCellMar', marginsXml('tblCellMar', ta.cellMargins));
    else if (fresh) pr.set('w:tblCellMar', '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>');
    pr.set('w:tblLook', lookXml(ta.look || { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true }));
    out.push(`<w:tbl><w:tblPr>${ordered(pr, TBLPR_ORDER)}</w:tblPr><w:tblGrid>${W2.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`);
    grid.forEach((r, ri) => {
      const row = rows[ri] || { content: [] };
      const ra = row.attrs || {};
      const isHeader = (row.content || []).length && (row.content || []).every((c) => c.type === 'tableHeader');
      const rc = ra.change;
      const trPr = new Map();
      for (const [nm, x] of ra.ox || []) trPr.set(nm, x);
      if (ra.cantSplit) trPr.set('w:cantSplit', '<w:cantSplit/>');
      if (ra.height) trPr.set('w:trHeight', `<w:trHeight w:val="${Math.round(ra.height)}"${ra.heightRule === 'exact' ? ' w:hRule="exact"' : ''}/>`);
      if (isHeader) trPr.set('w:tblHeader', '<w:tblHeader/>');
      if (rc && (rc.type === 'ins' || rc.type === 'del')) trPr.set(`w:${rc.type}`, `<w:${rc.type} w:id="${revN++}" w:author="${escAttr(rc.author || 'Unknown')}"${rc.date ? ` w:date="${escAttr(rc.date)}"` : ''}/>`);
      out.push(`<w:tr>${trPr.size ? `<w:trPr>${ordered(trPr, TRPR_ORDER)}</w:trPr>` : ''}`);
      for (let k = 0; k < cols; k++) {
        const g = r[k];
        if (g && g.covered) continue;
        const span = g ? g.cs : 1;
        const w = W2.slice(k, k + span).reduce((s, x) => s + x, 0);
        const ca = (g && g.cell.attrs) || {};
        const bg = hexColor(ca.backgroundColor);
        const tc = new Map();
        for (const [nm, x] of ca.ox || []) tc.set(nm, x);
        // a percentage/auto preferred width stays as Word wrote it while the column was not resized here
        const oldW = tc.get('w:tcW');
        if (!(oldW && /w:type="(pct|auto)"/.test(oldW) && ta.layout !== 'fixed')) tc.set('w:tcW', `<w:tcW w:w="${w}" w:type="dxa"/>`);
        if (span > 1) tc.set('w:gridSpan', `<w:gridSpan w:val="${span}"/>`);
        if (g && g.rs > 1) tc.set('w:vMerge', g.origin ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>');
        if (ca.borders) tc.set('w:tcBorders', bordersXml('tcBorders', ca.borders, TC_BORDER_ORDER));
        if (bg) tc.set('w:shd', `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`);
        if (ca.noWrap) tc.set('w:noWrap', '<w:noWrap/>');
        if (ca.margins) tc.set('w:tcMar', marginsXml('tcMar', ca.margins));
        if (ca.textDir) tc.set('w:textDirection', `<w:textDirection w:val="${ca.textDir}"/>`);
        if (ca.vAlign) tc.set('w:vAlign', `<w:vAlign w:val="${ca.vAlign}"/>`);
        const tcPr = `<w:tcPr>${ordered(tc, TCPR_ORDER)}</w:tcPr>`;
        if (!g || g.cont) { out.push(`<w:tc>${tcPr}<w:p/></w:tc>`); continue; }
        out.push(`<w:tc>${tcPr}`);
        cellContent(g.cell.content);
        out.push('</w:tc>');
      }
      out.push('</w:tr>');
    });
    out.push('</w:tbl>');
  };

  /* ============ sections: page setup per section, header/footer references ============ */
  function sectionXml(a, refs = null) {
    const page = a.page || model.page;
    const box = pageBox(page); const mg = page.margins || model.page.margins;
    const sect = new Map();
    let parsed = null;
    if (a.sectPr) { parsed = parseXml(a.sectPr); for (const c of kids(parsed, '*')) { const arr = sect.get(c.n) || []; arr.push(serialize(c)); sect.set(c.n, arr); } }
    // header/footer references are only valid inside their original package
    if (!patch) { sect.delete('w:headerReference'); sect.delete('w:footerReference'); }
    if (refs) for (const [tg, list] of Object.entries(refs)) { if (list.length) sect.set(tg, list); else sect.delete(tg); }
    sect.set('w:pgSz', [`<w:pgSz w:w="${mmToTwip(box.width)}" w:h="${mmToTwip(box.height)}"${page.orientation === 'landscape' ? ' w:orient="landscape"' : ''}/>`]);
    const oldMar = parsed && kids(parsed, 'w:pgMar')[0];
    sect.set('w:pgMar', [`<w:pgMar w:top="${mmToTwip(mg.top)}" w:right="${mmToTwip(mg.right)}" w:bottom="${mmToTwip(mg.bottom)}" w:left="${mmToTwip(mg.left)}" w:header="${(oldMar && oldMar.a['w:header']) || 708}" w:footer="${(oldMar && oldMar.a['w:footer']) || 708}" w:gutter="${(oldMar && oldMar.a['w:gutter']) || 0}"/>`]);
    const oldCols = parsed && kids(parsed, 'w:cols')[0];
    if ((page.columns || 1) > 1) sect.set('w:cols', [`<w:cols w:num="${page.columns}" w:space="708"/>`]);
    else if (!oldCols || Number(oldCols.a['w:num'] || 1) > 1) sect.set('w:cols', ['<w:cols w:space="708"/>']);
    if (a.titlePage) sect.set('w:titlePg', ['<w:titlePg/>']); else sect.delete('w:titlePg');
    if (a.type && a.type !== 'nextPage') sect.set('w:type', [`<w:type w:val="${escAttr(a.type)}"/>`]); else sect.delete('w:type');
    if (!sect.has('w:docGrid') && !patch) sect.set('w:docGrid', ['<w:docGrid w:linePitch="360"/>']);
    return `<w:sectPr>${ordered(sect, SECT_ORDER)}</w:sectPr>`;
  }
  // headers/footers (default, first page, even pages): rich content from their own document, written into their own
  // parts with their own relationships; unchanged ones keep their original part
  const hfPlainXml = (x) => {
    const fld = (instr, res) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r>${t(res)}</w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const body = String(x.text || '').split(/(\{page\}|\{pages\}|\{date\})/).filter(Boolean).map((p) => (p === '{page}' ? fld('PAGE', '1') : p === '{pages}' ? fld('NUMPAGES', '1') : p === '{date}' ? fld('DATE', new Date().toLocaleDateString()) : `<w:r><w:rPr><w:sz w:val="18"/></w:rPr>${t(fillFields(p, { title: (model.meta && model.meta.title) || '' }))}</w:r>`)).join('');
    return `<w:p><w:pPr>${x.align && x.align !== 'left' ? `<w:jc w:val="${x.align}"/>` : ''}</w:pPr>${body}</w:p>`;
  };
  const writeHfPart = (x, kind) => {
    const tg = kind === 'header' ? 'w:hdr' : 'w:ftr';
    let part; let k = 1; do { part = `of_${kind}${k++}.xml`; } while (zip.file(dir + part));
    scope = { rels: [], n: 1, part };
    let body = ''; let rels = [];
    try {
      if (x.doc && x.doc.content && x.doc.content.length) {
        const saved = out; out = [];
        for (const c of x.doc.content) block(c, {});
        body = out.join(''); out = saved;
        if (!/<\/w:p>$/.test(body)) body += '<w:p/>';
      } else body = hfPlainXml(x);
    } finally { rels = scope.rels; scope = null; }
    zip.file(dir + part, `${XMLH}<${tg} ${Object.entries(NS).map(([p, u]) => `xmlns:${p}="${u}"`).join(' ')} mc:Ignorable="w14 w15">${body}</${tg}>`);
    if (rels.length) zip.file(`${dir}_rels/${part}.rels`, `${XMLH}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${escAttr(r.target)}"${r.mode ? ` TargetMode="${r.mode}"` : ''}/>`).join('')}</Relationships>`);
    ensureOverride(dir + part, CT[kind]);
    return addRel(RT[kind], part);
  };
  const hfRefs = { 'w:headerReference': [], 'w:footerReference': [] };
  if (patch && keep.sectPr) for (const c of kids(parseXml(keep.sectPr), '*')) if (hfRefs[c.n]) hfRefs[c.n].push(serialize(c));
  {
    // header content must not take comment ranges, TOC bookmarks or footnotes of the body
    const savedIx = ix.n; const savedHeading = headingN; const savedNotes = footnotes.length;
    ix.n = -1e9;
    for (const [key, type] of [['header', 'default'], ['footer', 'default'], ['headerFirst', 'first'], ['footerFirst', 'first'], ['headerEven', 'even'], ['footerEven', 'even']]) {
      const kind = key.startsWith('header') ? 'header' : 'footer';
      const refTag = `w:${kind}Reference`;
      const x = model[key];
      const orig = patch ? ((keep.hf && keep.hf[key]) || (keep[key] && key.length === 6 ? hfKey(keep[key]) : '')) : null;
      if (patch && hfKey(x) === (orig || '')) continue;
      hfRefs[refTag] = hfRefs[refTag].filter((r) => { const ty = (/w:type="([^"]+)"/.exec(r) || [null, 'default'])[1]; return ty !== type; });
      if (hfHas(x)) hfRefs[refTag].push(`<${refTag} w:type="${type}" r:id="${writeHfPart(x, kind)}"/>`);
    }
    ix.n = savedIx; headingN = savedHeading; footnotes.length = savedNotes;
  }
  // a section break made here (no sectPr of its own) shows the same headers as the last section – Word would show none
  const sectionOf = (a) => sectionXml(a.sectPr ? a : { ...a, sectPr: null }, a.sectPr ? null : hfRefs);

  const top = (model.doc && model.doc.content) || [];
  const secPageAt = []; { let pg = model.page; for (let k = top.length - 1; k >= 0; k--) { if (top[k].type === 'sectionBreak' && top[k].attrs && top[k].attrs.page) pg = top[k].attrs.page; secPageAt[k] = pg; } }
  for (let k = 0; k < top.length; k++) {
    const n = top[k]; const nx = top[k + 1];
    curPage = secPageAt[k];
    // Word keeps a section's properties in its last paragraph
    if (nx && nx.type === 'sectionBreak' && (n.type === 'paragraph' || n.type === 'heading')) { block(n, { pextra: { 'w:sectPr': sectionOf(nx.attrs || {}) } }); k++; continue; }
    block(n, { top: true });
  }
  if (!out.length) out.push('<w:p/>');

  /* ============ final section (headers/footers were written before the body) ============ */
  const sectXml = sectionXml({ page: model.page, titlePage: model.titlePage, type: null, sectPr: patch ? keep.sectPr : null }, hfRefs);

  /* ============ document.xml ============ */
  let rootTag = (patch && keep.rootTag) || `<w:document ${Object.entries(NS).map(([p, u]) => `xmlns:${p}="${u}"`).join(' ')} mc:Ignorable="w14 w15">`;
  for (const p of ['w', 'r', 'wp', 'm', 'w14', 'mc']) if (!rootTag.includes(`xmlns:${p}=`)) rootTag = rootTag.replace(/>$/, ` xmlns:${p}="${NS[p]}">`);
  const bgXml = model.page.color ? `<w:background w:color="${hexColor(model.page.color) || 'FFFFFF'}"/>` : '';
  zip.file(docPath, `${XMLH}${rootTag}${bgXml}<w:body>${out.join('')}${sectXml}</w:body></w:document>`);
  // main part type follows the TARGET extension (a docm saved as .docx must not keep the macro-enabled type, a docx
  // saved as .dotx needs the template type – Word refuses both otherwise)
  setOverride(docPath, CT[ext] || CT.docx);
  if (patch && !/m$/.test(ext)) {
    // macro-free target: drop the VBA project (and what hangs off it – vbaData.xml) the way Word does on "save as docx"
    const resolve = (base, target) => {
      if (target.startsWith('/')) return target.slice(1);
      const segs = (base + target).split('/'); const outSegs = [];
      for (const s of segs) { if (s === '..') outSegs.pop(); else if (s && s !== '.') outSegs.push(s); }
      return outSegs.join('/');
    };
    const dropPart = async (p, seen = new Set()) => {
      if (seen.has(p)) return; seen.add(p);
      const pRels = p.replace(/([^/]+)$/, '_rels/$1.rels');
      const rx = await read(pRels);
      if (rx) for (const r of kids(parseXml(rx), 'Relationship')) if (r.a.TargetMode !== 'External' && r.a.Target) await dropPart(resolve(p.replace(/[^/]+$/, ''), r.a.Target), seen);
      zip.remove(pRels); zip.remove(p); dropOverride(p);
    };
    for (const r of relList.filter((x) => x.type === RT.vbaProject)) { await dropPart(partOf(r)); removeRel(r); }
    const binLeft = Object.keys(zip.files).some((n) => /\.bin$/i.test(n) && !zip.files[n].dir);
    if (!binLeft) ctXml = ctXml.replace(/<Default\b[^>]*\/>/gi, (m) => (/Extension="bin"/i.test(m) && /vbaProject/i.test(m) ? '' : m));
  }

  /* ============ footnotes ============ */
  const fnRel = relOf(RT.footnotes);
  if (footnotes.length || fnRel) {
    const path = fnRel ? partOf(fnRel) : `${dir}footnotes.xml`;
    let seps = '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
    let rootF = `<w:footnotes xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:m="${NS.m}" xmlns:wp="${NS.wp}" xmlns:w14="${NS.w14}" xmlns:mc="${NS.mc}" mc:Ignorable="w14">`;
    if (fnRel) {
      const old = await read(path);
      if (old) {
        const r = /<w:footnotes\b[^>]*>/.exec(old); if (r) rootF = r[0];
        const s = [...old.matchAll(/<w:footnote\b[^>]*w:type="(?:separator|continuationSeparator|continuationNotice)"[\s\S]*?<\/w:footnote>/g)].map((m) => m[0]).join('');
        if (s) seps = s;
      }
    }
    const notes = footnotes.map((f, k) => {
      const id = k + 1;
      if (patch && f.ox && f.ox.xml && f.ox.text === (f.text || '')) return `<w:footnote w:id="${id}">${f.ox.xml}</w:footnote>`;
      return `<w:footnote w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="20"/></w:rPr><w:footnoteRef/></w:r><w:r><w:rPr><w:sz w:val="20"/></w:rPr>${t(` ${f.text || ''}`)}</w:r></w:p></w:footnote>`;
    }).join('');
    zip.file(path, `${XMLH}${rootF}${seps}${notes}</w:footnotes>`);
    if (!fnRel) { addRel(RT.footnotes, path.slice(dir.length)); ensureOverride(path, CT.footnotes); }
  }

  /* ============ comments (+ replies, resolved) ============ */
  const cRel = relOf(RT.comments);
  const exRel = relOf(RT.commentsExtended);
  // durable ids (commentsIds) and UTC dates (commentsExtensible) are rewritten for the comments we write – the ids of
  // comments read from the file stay the same (Word/Teams follow a comment thread by its durable id)
  const idsRel = relList.find((r) => /commentsIds$/.test(r.type)); const cexRel = relList.find((r) => /commentsExtensible$/.test(r.type));
  for (const r of [idsRel, cexRel]) if (r) { const p = partOf(r); zip.remove(p); dropOverride(p); removeRel(r); }
  if (used.length || cRel) {
    let pid = 0x10000001;
    const para = () => (pid++).toString(16).toUpperCase().padStart(8, '0');
    const exs = []; const cids = []; const cexs = [];
    const takenDurable = new Set();
    const durableOf = (want, seed) => {
      let d = want && /^[0-9A-F]{8}$/i.test(want) && parseInt(want, 16) < 0x7fffffff && !takenDurable.has(want.toUpperCase()) ? want.toUpperCase() : null;
      if (!d) { let h = 0x2a5f3c1; for (const ch of String(seed)) h = (Math.imul(h ^ ch.charCodeAt(0), 16777619) >>> 0); do { d = ((h >>> 1) || 1).toString(16).toUpperCase().padStart(8, '0'); h = (Math.imul(h, 2654435761) + 7) >>> 0; } while (takenDurable.has(d)); }
      takenDurable.add(d);
      return d;
    };
    const one = (wid, author, date, text, parentPara, done, durable) => {
      const lines = String(text || '').split('\n');
      let last = '';
      const ps = lines.map((line, k) => { const id = para(); last = id; return `<w:p w14:paraId="${id}" w14:textId="77777777"><w:pPr><w:spacing w:after="0"/></w:pPr>${k === 0 ? '<w:r><w:annotationRef/></w:r>' : ''}<w:r>${t(line)}</w:r></w:p>`; }).join('');
      exs.push(`<w15:commentEx w15:paraId="${last}"${parentPara ? ` w15:paraIdParent="${parentPara}"` : ''} w15:done="${done ? 1 : 0}"/>`);
      const d = durableOf(durable, `${wid}|${author}|${date}|${text}`);
      cids.push(`<w16cid:commentId w16cid:paraId="${last}" w16cid:durableId="${d}"/>`);
      if (date) cexs.push(`<w16cex:commentExtensible w16cex:durableId="${d}" w16cex:dateUtc="${escAttr(new Date(date).toISOString().replace(/\.\d+Z$/, 'Z'))}"/>`);
      return { xml: `<w:comment w:id="${wid}" w:author="${escAttr(author || 'Unknown')}"${date ? ` w:date="${escAttr(date)}"` : ''} w:initials="${escAttr(initials(author))}">${ps}</w:comment>`, para: last };
    };
    let xml = '';
    for (const c of used) {
      const main = one(c.wid, c.author, c.date, c.text, null, c.resolved, c.durableId);
      xml += main.xml;
      (c.replies || []).forEach((r, k) => { xml += one(c.replyIds[k], r.author, r.date, r.text, main.para, c.resolved, r.durableId).xml; });
    }
    if (used.length) {
      const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid'; const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
      const idsPath = `${dir}commentsIds.xml`; const cexPath = `${dir}commentsExtensible.xml`;
      zip.file(idsPath, `${XMLH}<w16cid:commentsIds xmlns:mc="${NS.mc}" xmlns:w16cid="${W16CID}" mc:Ignorable="w16cid">${cids.join('')}</w16cid:commentsIds>`);
      addRel('http://schemas.microsoft.com/office/2016/09/relationships/commentsIds', 'commentsIds.xml'); ensureOverride(idsPath, 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml');
      zip.file(cexPath, `${XMLH}<w16cex:commentsExtensible xmlns:mc="${NS.mc}" xmlns:w16cex="${W16CEX}" mc:Ignorable="w16cex">${cexs.join('')}</w16cex:commentsExtensible>`);
      addRel('http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible', 'commentsExtensible.xml'); ensureOverride(cexPath, 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml');
    }
    const cPath = cRel ? partOf(cRel) : `${dir}comments.xml`;
    zip.file(cPath, `${XMLH}<w:comments xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:w14="${NS.w14}" xmlns:mc="${NS.mc}" mc:Ignorable="w14">${xml}</w:comments>`);
    if (!cRel) { addRel(RT.comments, cPath.slice(dir.length)); ensureOverride(cPath, CT.comments); }
    const exPath = exRel ? partOf(exRel) : `${dir}commentsExtended.xml`;
    zip.file(exPath, `${XMLH}<w15:commentsEx xmlns:w15="${NS.w15}" xmlns:mc="${NS.mc}" mc:Ignorable="w15">${exs.join('')}</w15:commentsEx>`);
    if (!exRel) { addRel(RT.commentsExtended, exPath.slice(dir.length)); ensureOverride(exPath, CT.commentsExtended); }
  }

  /* ============ numbering ============ */
  if (numDefs.length) {
    const nRel = relOf(RT.numbering);
    const nPath = nRel ? partOf(nRel) : `${dir}numbering.xml`;
    let nx = (nRel && await read(nPath)) || `${XMLH}<w:numbering xmlns:w="${NS.w}"></w:numbering>`;
    const abs = [...absDefs.entries()].map(([kind, id]) => absXml(kind, id)).join('');
    const nums = numDefs.map((d) => `<w:num w:numId="${d.numId}"><w:abstractNumId w:val="${d.abs}"/><w:lvlOverride w:ilvl="${d.level || 0}"><w:startOverride w:val="${d.start || 1}"/></w:lvlOverride></w:num>`).join('');
    const firstNum = nx.search(/<w:num\b/);
    if (firstNum >= 0) nx = nx.slice(0, firstNum) + abs + nx.slice(firstNum); else nx = nx.replace(/<w:numIdMacAtCleanup\b|<\/w:numbering>/, (m) => abs + m);
    nx = nx.replace(/(<w:numIdMacAtCleanup\b[^>]*\/>)?<\/w:numbering>/, (m) => nums + m);
    zip.file(nPath, nx);
    if (!nRel) { addRel(RT.numbering, nPath.slice(dir.length)); ensureOverride(nPath, CT.numbering); }
  }

  /* ============ styles ============ */
  const fontName = firstFont((model.defaults && model.defaults.font) || 'Calibri');
  if (!stylesXml) {
    const d = model.defaults || {};
    const sz = Math.round((d.fontSize || 11) * 2);
    stylesXml = `${XMLH}<w:styles xmlns:w="${NS.w}" xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" mc:Ignorable="w14"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${escAttr(fontName)}" w:eastAsia="${escAttr(fontName)}" w:hAnsi="${escAttr(fontName)}" w:cs="${escAttr(fontName)}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:lang w:val="${(typeof document !== 'undefined' && document.documentElement.lang === 'de') ? 'de-DE' : 'en-US'}" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="${Math.round((d.spaceAfter != null ? d.spaceAfter : 8) * 20)}" w:line="${Math.round((d.lineHeight || 1.15) * 240)}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>`;
    stylesRel = { target: 'styles.xml' };
    addRel(RT.styles, 'styles.xml');
    ensureOverride(`${dir}styles.xml`, CT.styles);
    for (const name of Object.keys(STYLE_ID)) missingStyles.add(name);
  }
  if (!tableNormalId && /w:styleId="TableNormal"/.test(stylesXml)) tableNormalId = 'TableNormal';
  const tableDefs = [...usedTableStyles].map((id) => builtinTableStyles()[id]).filter(Boolean).map((def) => tableStyleXml(def, tableNormalId));
  if (missingStyles.size || tableDefs.length) {
    const all = styleDefs(normalId, null);
    stylesXml = stylesXml.replace('</w:styles>', `${[...missingStyles].map((n) => all[n]).join('')}${tableDefs.join('')}</w:styles>`);
    zip.file(stylesRel.target.startsWith('/') ? stylesRel.target.slice(1) : dir + stylesRel.target, stylesXml);
  }
  void defs;

  /* ============ settings ============ */
  const sRel = relOf(RT.settings);
  let sx = sRel ? await read(partOf(sRel)) : null;
  const before = (xml, names, snippet) => {
    let at = -1;
    for (const n of names) { const k = xml.indexOf(`<${n}`); if (k >= 0 && (at < 0 || k < at)) at = k; }
    if (at < 0) at = xml.lastIndexOf('</w:settings>');
    return xml.slice(0, at) + snippet + xml.slice(at);
  };
  const newSettings = !sx;
  // a Word file without settings keeps having none (a compatibility mode of our own would change Word's layout of it)
  if (!sx && patch) sx = `${XMLH}<w:settings xmlns:w="${NS.w}" xmlns:m="${NS.m}"></w:settings>`;
  if (!sx) {
    sx = `${XMLH}<w:settings xmlns:w="${NS.w}" xmlns:m="${NS.m}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
  }
  // settings may only reference separator notes that exist: Word reports the file as damaged when footnotePr /
  // endnotePr name w:footnote / w:endnote ids but there is no footnotes / endnotes part
  for (const [kind, type] of [['footnote', RT.footnotes], ['endnote', RT.endnotes]]) {
    const rel = relOf(type);
    if (rel && zip.file(partOf(rel))) continue;
    sx = sx.replace(new RegExp(`<w:${kind}Pr>([\\s\\S]*?)</w:${kind}Pr>`), (m, inner) => {
      const rest = inner.replace(new RegExp(`<w:${kind}\\b[^>]*/>`, 'g'), '');
      return rest.trim() ? `<w:${kind}Pr>${rest}</w:${kind}Pr>` : '';
    });
  }
  // never ask Word to update all fields on open (modal prompt, hangs unattended conversions) – our TOC is filled in
  sx = sx.replace(/<w:updateFields\b[^>]*\/>/g, '');
  sx = sx.replace(/<w:trackRevisions\b[^>]*\/>/g, '');
  if (model.trackChanges) sx = before(sx, ['w:doNotTrackMoves', 'w:doNotTrackFormatting', 'w:documentProtection', 'w:autoFormatOverride', 'w:styleLockTheme', 'w:styleLockQFSet', 'w:defaultTabStop', 'w:autoHyphenation', 'w:hyphenationZone', 'w:characterSpacingControl', 'w:compat', 'w:rsids', 'm:mathPr', 'w:themeFontLang', 'w:clrSchemeMapping', 'w:shapeDefaults', 'w:decimalSymbol', 'w:listSeparator', 'w14:', 'w15:'], '<w:trackRevisions/>');
  sx = sx.replace(/<w:evenAndOddHeaders\b[^>]*\/>/g, '');
  if (model.evenOdd) sx = before(sx, ['w:bookFoldRevPrinting', 'w:bookFoldPrinting', 'w:bookFoldPrintingSheets', 'w:drawingGridHorizontalSpacing', 'w:drawingGridVerticalSpacing', 'w:displayHorizontalDrawingGridEvery', 'w:displayVerticalDrawingGridEvery', 'w:doNotUseMarginsForDrawingGridOrigin', 'w:drawingGridHorizontalOrigin', 'w:drawingGridVerticalOrigin', 'w:doNotShadeFormData', 'w:noPunctuationKerning', 'w:characterSpacingControl', 'w:printTwoOnOne', 'w:strictFirstAndLastChars', 'w:noLineBreaksAfter', 'w:noLineBreaksBefore', 'w:savePreviewPicture', 'w:doNotValidateAgainstSchema', 'w:saveInvalidXml', 'w:ignoreMixedContent', 'w:alwaysShowPlaceholderText', 'w:doNotDemarcateInvalidXml', 'w:saveXmlDataOnly', 'w:useXSLTWhenSaving', 'w:saveThroughXslt', 'w:showXMLTags', 'w:alwaysMergeEmptyNamespace', 'w:updateFields', 'w:hdrShapeDefaults', 'w:footnotePr', 'w:endnotePr', 'w:compat', 'w:docVars', 'w:rsids', 'm:mathPr', 'w:attachedSchema', 'w:themeFontLang', 'w:clrSchemeMapping', 'w:doNotIncludeSubdocsInStats', 'w:doNotAutoCompressPictures', 'w:forceUpgrade', 'w:captions', 'w:readModeInkLockDown', 'w:smartTagType', 'sl:schemaLibrary', 'w:shapeDefaults', 'w:doNotEmbedSmartTags', 'w:decimalSymbol', 'w:listSeparator', 'w14:', 'w15:'], '<w:evenAndOddHeaders/>');
  if (model.page.color && !/<w:displayBackgroundShape\b/.test(sx)) sx = before(sx, ['w:printPostScriptOverText', 'w:printFractionalCharacterWidth', 'w:printFormsData', 'w:embedTrueTypeFonts', 'w:embedSystemFonts', 'w:saveSubsetFonts', 'w:saveFormsData', 'w:mirrorMargins', 'w:alignBordersAndEdges', 'w:bordersDoNotSurroundHeader', 'w:bordersDoNotSurroundFooter', 'w:gutterAtTop', 'w:hideSpellingErrors', 'w:hideGrammaticalErrors', 'w:activeWritingStyle', 'w:proofState', 'w:formsDesign', 'w:attachedTemplate', 'w:linkStyles', 'w:stylePaneFormatFilter', 'w:stylePaneSortMethod', 'w:documentType', 'w:mailMerge', 'w:revisionView', 'w:trackRevisions', 'w:doNotTrackMoves', 'w:defaultTabStop'], '<w:displayBackgroundShape/>');
  if (!newSettings || /<w:settings[^>]*>\s*<[a-z]/.test(sx)) {
    zip.file(sRel ? partOf(sRel) : `${dir}settings.xml`, sx);
    if (newSettings) { addRel(RT.settings, 'settings.xml'); ensureOverride(`${dir}settings.xml`, CT.settings); }
  }

  /* ============ relationships, package parts ============ */
  for (const id of removed) relsXml = relsXml.replace(new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*/>`), '');
  relsXml = relsXml.replace('</Relationships>', `${newRels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${escAttr(r.target)}"${r.mode ? ` TargetMode="${r.mode}"` : ''}/>`).join('')}</Relationships>`);
  zip.file(relsPath, relsXml);
  ensureDefault('png', 'image/png');

  // core properties
  const meta = model.meta || {};
  let core = patch ? await read('docProps/core.xml') : null;
  const now = isoNow();
  const coreTag = (n, v) => `<${n}>${esc(v)}</${n}>`;
  if (!core) {
    core = `${XMLH}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${meta.title ? coreTag('dc:title', meta.title) : ''}${meta.subject ? coreTag('dc:subject', meta.subject) : ''}${coreTag('dc:creator', meta.author || '')}${meta.keywords ? coreTag('cp:keywords', meta.keywords) : ''}${coreTag('cp:lastModifiedBy', meta.author || '')}<cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
    if (!patch || !zip.file('docProps/core.xml')) {
      let rr = (patch && await read('_rels/.rels')) || `${XMLH}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT.officeDocument}" Target="${docPath}"/></Relationships>`;
      if (!rr.includes('core-properties')) rr = rr.replace('</Relationships>', `<Relationship Id="rIdOfCore" Type="${RT.core}" Target="docProps/core.xml"/></Relationships>`);
      if (!patch) rr = rr.replace('</Relationships>', `<Relationship Id="rIdOfApp" Type="${RT.app}" Target="docProps/app.xml"/></Relationships>`);
      zip.file('_rels/.rels', rr);
      ensureOverride('docProps/core.xml', CT.core);
    }
  } else {
    const setTag = (xml, n, v, add = true) => (new RegExp(`<${n}\\b[^>]*>[\\s\\S]*?</${n}>|<${n}\\b[^>]*/>`).test(xml) ? xml.replace(new RegExp(`(<${n}\\b[^>]*?)(?:/>|>[\\s\\S]*?</${n}>)`), (m, open) => `${open.replace(/\/$/, '')}>${esc(v)}</${n}>`) : add ? xml.replace('</cp:coreProperties>', `${coreTag(n, v)}</cp:coreProperties>`) : xml);
    if (meta.title) core = setTag(core, 'dc:title', meta.title);
    if (meta.author) core = setTag(core, 'cp:lastModifiedBy', meta.author);
    core = setTag(core, 'dcterms:modified', now, false);
  }
  zip.file('docProps/core.xml', core);
  if (!patch) {
    zip.file('docProps/app.xml', `${XMLH}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Hostwerk Office</Application><DocSecurity>0</DocSecurity></Properties>`);
    ensureOverride('docProps/app.xml', CT.app);
    if (!zip.file('_rels/.rels')) zip.file('_rels/.rels', `${XMLH}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT.officeDocument}" Target="${docPath}"/><Relationship Id="rIdOfCore" Type="${RT.core}" Target="docProps/core.xml"/><Relationship Id="rIdOfApp" Type="${RT.app}" Target="docProps/app.xml"/></Relationships>`);
  }
  zip.file('[Content_Types].xml', ctXml);
  // [Content_Types].xml first, like Word does
  const ordered2 = new JSZip();
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  names.sort((x, y) => (x === '[Content_Types].xml' ? -1 : y === '[Content_Types].xml' ? 1 : x === '_rels/.rels' ? -1 : y === '_rels/.rels' ? 1 : 0));
  for (const n of names) ordered2.file(n, zip.files[n].async('uint8array'), { binary: true, createFolders: false });
  return ordered2.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

const initials = (name) => String(name || '?').split(/\s+/).filter(Boolean).map((w) => w[0]).join('').slice(0, 3).toUpperCase() || '?';
