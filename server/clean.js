'use strict';
/*
 * "Clean before sharing": remove personal information and hidden content from a copy of a file.
 *   docx/docm/dotx · xlsx/xlsm · pptx/pptm · odt/ods/odp/odg · pdf · jpg/png (EXIF)
 *
 *   cleanFile(buf, ext, opts) → { data, report: { author, comments, revisions, hidden, custom, thumbnail, rsid, exif,
 *                                  paths, info }, warnings: [] }
 *   opts (all default true): author, comments, revisions (accept tracked changes), hidden (hidden text), custom
 *   (custom properties), thumbnail, rsid (revision ids), exif (image metadata), templates (template / file paths)
 * Pure: unchanged package parts are copied byte for byte (zippatch). Nothing of the content is executed.
 */
const ZP = require('./zippatch');
const S = require('./sheet');

const DEFAULTS = { author: true, comments: true, revisions: true, hidden: true, custom: true, thumbnail: true, rsid: true, exif: true, templates: true };
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Balanced elements <tag …>…</tag> or <tag …/> → [{ start, end, open, inner }] (outermost first, nested skipped) */
function elements(xml, tag) {
  const out = [];
  const re = new RegExp(`<(/?)${esc(tag)}(?=[\\s/>])([^>]*?)(/?)>`, 'g');
  let depth = 0; let cur = null; let m;
  while ((m = re.exec(xml))) {
    if (m[1]) { if (depth > 0 && --depth === 0 && cur) { out.push({ start: cur.start, end: re.lastIndex, open: cur.open, inner: xml.slice(cur.innerStart, m.index) }); cur = null; } continue; }
    if (m[3]) { if (depth === 0) out.push({ start: m.index, end: re.lastIndex, open: m[0], inner: '' }); continue; }
    if (depth === 0) cur = { start: m.index, open: m[0], innerStart: re.lastIndex };
    depth++;
  }
  return out;
}
/** Replace every outermost element <tag> by fn(el) (string) – from the end, so offsets stay valid */
function replaceEl(xml, tag, fn) {
  let n = 0;
  const list = elements(xml, tag);
  for (let i = list.length - 1; i >= 0; i--) { const e = list[i]; const r = fn(e); if (r === undefined || r === null) continue; if (r !== xml.slice(e.start, e.end)) n++; xml = xml.slice(0, e.start) + r + xml.slice(e.end); }
  return { xml, n };
}
const strip = (xml, tag) => replaceEl(xml, tag, () => '');
const unwrap = (xml, tag) => replaceEl(xml, tag, (e) => e.inner);

/* ---------------- images ---------------- */
function cleanJpeg(b) {
  if (!(b[0] === 0xff && b[1] === 0xd8)) return null;
  const parts = [b.subarray(0, 2)]; let i = 2; let removed = 0;
  while (i + 4 <= b.length) {
    if (b[i] !== 0xff) return null;
    const mk = b[i + 1];
    if (mk === 0xda || mk === 0xd9) { parts.push(b.subarray(i)); break; }
    const len = b.readUInt16BE(i + 2);
    if (len < 2 || i + 2 + len > b.length) return null;
    // APP1 (Exif, XMP), APP12/13 (Photoshop/IPTC), COM – APP0 (JFIF), APP2 (ICC), APP14 (Adobe) stay
    if (mk === 0xe1 || mk === 0xec || mk === 0xed || mk === 0xfe) removed++; else parts.push(b.subarray(i, i + 2 + len));
    i += 2 + len;
  }
  return removed ? Buffer.concat(parts) : null;
}
function cleanPng(b) {
  if (b.length < 8 || b.readUInt32BE(0) !== 0x89504e47) return null;
  const parts = [b.subarray(0, 8)]; let i = 8; let removed = 0;
  while (i + 12 <= b.length) {
    const len = b.readUInt32BE(i); const type = b.toString('latin1', i + 4, i + 8);
    if (i + 12 + len > b.length) return null;
    if (/^(tEXt|zTXt|iTXt|eXIf|tIME)$/.test(type)) removed++; else parts.push(b.subarray(i, i + 12 + len));
    i += 12 + len;
    if (type === 'IEND') break;
  }
  return removed ? Buffer.concat(parts) : null;
}
const cleanImage = (name, data) => (/\.jpe?g$/i.test(name) ? cleanJpeg(data) : /\.png$/i.test(name) ? cleanPng(data) : null);

/* ---------------- OOXML ---------------- */
function dropPart(ctx, name) {
  if (!ctx.z.has(name)) return false;
  ctx.changes[name] = null;
  ctx.ct = ctx.ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${esc(name)}"[^>]*/>`, 'g'), '');
  // relationships pointing to it
  for (const [rn, e] of ctx.z) {
    if (!/_rels\/[^/]*\.rels$/.test(rn) || ctx.changes[rn] === null) continue;
    const base = rn.replace(/_rels\/([^/]*)\.rels$/, '$1');
    let x = ctx.changes[rn] != null ? ctx.changes[rn] : ZP.inflate(e).toString('utf8');
    const before = x;
    x = x.replace(/<Relationship\b[^>]*\/>/g, (r) => { const t = (/\bTarget="([^"]*)"/.exec(r) || [])[1]; const mode = (/TargetMode="External"/.test(r)); if (!t || mode) return r; const full = t.startsWith('/') ? t.slice(1) : resolve(base, t); return full === name ? '' : r; });
    if (x !== before) ctx.changes[rn] = x;
  }
  if (ctx.changes[`${name.replace(/([^/]*)$/, '_rels/$1')}.rels`] === undefined && ctx.z.has(name.replace(/([^/]*)$/, '_rels/$1.rels'))) ctx.changes[name.replace(/([^/]*)$/, '_rels/$1.rels')] = null;
  return true;
}
function resolve(base, t) { const parts = base.split('/').slice(0, -1); for (const p of t.split('/')) { if (p === '..') parts.pop(); else if (p && p !== '.') parts.push(p); } return parts.join('/'); }
const get = (ctx, n) => (ctx.changes[n] != null ? String(ctx.changes[n]) : ctx.z.has(n) ? ZP.inflate(ctx.z.get(n)).toString('utf8') : null);
const set = (ctx, n, x) => { if (get(ctx, n) !== x) ctx.changes[n] = x; };
const relsOf = (ctx, n) => { const rp = n.replace(/([^/]*)$/, '_rels/$1.rels'); const x = get(ctx, rp) || ''; return [...x.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => ({ id: (/\bId="([^"]*)"/.exec(m[0]) || [])[1], type: (/\bType="([^"]*)"/.exec(m[0]) || [])[1] || '', target: (/\bTarget="([^"]*)"/.exec(m[0]) || [])[1] || '', ext: /TargetMode="External"/.test(m[0]) })).map((r) => ({ ...r, full: r.target.startsWith('/') ? r.target.slice(1) : resolve(n, r.target) })); };

function cleanCore(ctx, o, rep) {
  const n = 'docProps/core.xml'; let x = get(ctx, n); if (!x || !o.author) return;
  const b = x;
  x = x.replace(/<(dc:creator|cp:lastModifiedBy)\b[^>]*>[\s\S]*?<\/\1>|<(dc:creator|cp:lastModifiedBy)\b[^>]*\/>/g, '').replace(/<cp:lastPrinted\b[^>]*>[\s\S]*?<\/cp:lastPrinted>/g, '').replace(/<cp:revision\b[^>]*>[\s\S]*?<\/cp:revision>/g, '');
  if (x !== b) { set(ctx, n, x); rep.author = true; }
}
function cleanApp(ctx, o, rep) {
  const n = 'docProps/app.xml'; let x = get(ctx, n); if (!x) return;
  const b = x;
  if (o.author) x = x.replace(/<(Company|Manager|HyperlinkBase)\b[^>]*>[\s\S]*?<\/\1>|<(Company|Manager|HyperlinkBase)\s*\/>/g, '').replace(/<TotalTime>\d*<\/TotalTime>/, '<TotalTime>0</TotalTime>');
  if (o.templates) x = x.replace(/<Template>[\s\S]*?<\/Template>/, (t) => (/[\\/]/.test(t) ? `<Template>${t.replace(/<\/?Template>/g, '').split(/[\\/]/).pop()}</Template>` : t));
  if (x !== b) { set(ctx, n, x); rep.author = true; }
}

function cleanWordXml(x, o, rep) {
  if (o.comments) {
    const b = x;
    x = x.replace(/<w:comment(RangeStart|RangeEnd)\b[^>]*\/>/g, '');
    x = replaceEl(x, 'w:r', (e) => (/<w:commentReference\b/.test(e.inner) && !/<w:t\b/.test(e.inner) ? '' : undefined)).xml;
    x = x.replace(/<w:commentReference\b[^>]*\/>/g, '');
    if (x !== b) rep.comments += 1;
  }
  if (o.revisions) {
    let r;
    // deleted table rows (accept = row gone)
    r = replaceEl(x, 'w:tr', (e) => (/<w:trPr\b[^>]*>(?:(?!<\/w:trPr>)[\s\S])*<w:del\b/.test(e.inner) ? '' : undefined)); x = r.xml; rep.revisions += r.n;
    for (const t of ['w:del', 'w:moveFrom']) { r = strip(x, t); x = r.xml; rep.revisions += r.n; }
    for (const t of ['w:ins', 'w:moveTo']) { r = unwrap(x, t); x = r.xml; rep.revisions += r.n; }
    for (const t of ['w:rPrChange', 'w:pPrChange', 'w:sectPrChange', 'w:tblPrChange', 'w:tblGridChange', 'w:trPrChange', 'w:tcPrChange', 'w:numberingChange', 'w:cellIns', 'w:cellDel', 'w:cellMerge']) { r = strip(x, t); x = r.xml; rep.revisions += r.n; }
    x = x.replace(/<w:(moveFromRangeStart|moveFromRangeEnd|moveToRangeStart|moveToRangeEnd|customXmlInsRangeStart|customXmlInsRangeEnd|customXmlDelRangeStart|customXmlDelRangeEnd)\b[^>]*\/>/g, () => { rep.revisions++; return ''; });
  }
  if (o.hidden) {
    const r = replaceEl(x, 'w:r', (e) => (/<w:rPr\b[^>]*>(?:(?!<\/w:rPr>)[\s\S])*<w:vanish(\s+w:val="(1|true|on)")?\s*\/>/.test(e.inner) ? '' : undefined));
    x = r.xml; rep.hidden += r.n;
  }
  if (o.rsid) { const b = x.length; x = x.replace(/\sw:rsid\w*="[0-9A-Fa-f]*"/g, ''); if (x.length !== b) rep.rsid++; }
  return x;
}

function cleanOoxml(buf, ext, o) {
  const list = ZP.entries(buf);
  const ctx = { z: new Map(list.map((e) => [e.name, e])), changes: {}, ct: '' };
  ctx.ct = get(ctx, '[Content_Types].xml') || '';
  const ct0 = ctx.ct;
  const rep = { author: false, comments: 0, revisions: 0, hidden: 0, custom: false, thumbnail: false, rsid: 0, exif: 0, paths: 0 };
  const warnings = [];
  cleanCore(ctx, o, rep); cleanApp(ctx, o, rep);
  if (o.custom && dropPart(ctx, 'docProps/custom.xml')) rep.custom = true;
  if (o.thumbnail) for (const n of ctx.z.keys()) if (/^docProps\/thumbnail\.\w+$/i.test(n) && dropPart(ctx, n)) rep.thumbnail = true;
  if (/^(docx|docm|dotx|dotm)$/.test(ext)) {
    for (const n of ctx.z.keys()) {
      if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments|glossary\/document)\.xml$/.test(n)) continue;
      const x = get(ctx, n); const y = cleanWordXml(x, o, rep); if (y !== x) set(ctx, n, y);
    }
    if (o.comments) for (const n of ['word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml', 'word/people.xml']) if (dropPart(ctx, n)) rep.comments = Math.max(rep.comments, 1);
    let st = get(ctx, 'word/settings.xml');
    if (st) {
      if (o.rsid) st = st.replace(/<w:rsids>[\s\S]*?<\/w:rsids>/, () => { rep.rsid++; return ''; });
      if (o.templates) st = st.replace(/<w:attachedTemplate\b[^>]*\/>/, () => { rep.paths++; return ''; });
      set(ctx, 'word/settings.xml', st);
      // attached template relationship (a path on the author's PC)
      if (o.templates) { const rp = 'word/_rels/settings.xml.rels'; const rx = get(ctx, rp); if (rx) set(ctx, rp, rx.replace(/<Relationship\b[^>]*attachedTemplate[^>]*\/>/g, '')); }
    }
  } else if (/^(xlsx|xlsm|xltx|xltm)$/.test(ext)) {
    for (const n of [...ctx.z.keys()].filter((k) => /^xl\/worksheets\/sheet\d*\.xml$/.test(k))) {
      let x = get(ctx, n);
      for (const r of relsOf(ctx, n)) {
        if (!o.comments) break;
        if (/\/comments$|\/threadedComment$/.test(r.type)) { if (dropPart(ctx, r.full)) rep.comments++; }
        else if (/\/vmlDrawing$/.test(r.type)) {
          const v = get(ctx, r.full) || '';
          const shapes = (v.match(/<v:shape\b/g) || []).length; const notes = (v.match(/ObjectType="Note"/g) || []).length;
          if (shapes && shapes === notes) { dropPart(ctx, r.full); x = x.replace(new RegExp(`<legacyDrawing\\b[^>]*r:id="${esc(r.id)}"[^>]*/>`), ''); }
        }
      }
      if (o.rsid) x = x.replace(/\sxr:uid="\{[^"]*\}"/g, () => { rep.rsid++; return ''; });
      if (/<sheetState|state="hidden"|hidden="1"/.test(x)) rep.hiddenRows = true;
      set(ctx, n, x);
    }
    if (o.comments) for (const n of [...ctx.z.keys()].filter((k) => /^xl\/persons\/person\w*\.xml$/.test(k))) dropPart(ctx, n);
    let wb = get(ctx, 'xl/workbook.xml');
    if (wb) {
      if (o.templates) wb = wb.replace(/<mc:AlternateContent\b[^>]*>(?:(?!<\/mc:AlternateContent>)[\s\S])*?x15ac:absPath[\s\S]*?<\/mc:AlternateContent>/g, () => { rep.paths++; return ''; });
      if (o.rsid) wb = wb.replace(/<xr:revisionPtr\b[^>]*\/>/g, () => { rep.rsid++; return ''; }).replace(/\sxr:uid="\{[^"]*\}"/g, '');
      if (/<sheet\b[^>]*state="(hidden|veryHidden)"/.test(wb)) warnings.push('hiddenSheets');
      set(ctx, 'xl/workbook.xml', wb);
    }
  } else if (/^(pptx|pptm|ppsx|potx)$/.test(ext)) {
    if (o.comments) {
      for (const n of [...ctx.z.keys()].filter((k) => /^ppt\/comments\/[\w-]+\.xml$/.test(k))) if (dropPart(ctx, n)) rep.comments++;
      for (const n of ['ppt/commentAuthors.xml', 'ppt/authors.xml']) dropPart(ctx, n);
    }
    if (o.rsid) for (const n of [...ctx.z.keys()].filter((k) => /^ppt\/(slides\/slide\d+|presentation)\.xml$/.test(k))) { const x = get(ctx, n); const y = x.replace(/<p14:creationId\b[^>]*\/>/g, '').replace(/<p:extLst>\s*<p:ext uri="\{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E\}">\s*<\/p:ext>\s*<\/p:extLst>/g, ''); if (y !== x) { set(ctx, n, y); rep.rsid++; } }
    if (/show="0"/.test([...ctx.z.keys()].filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k)).map((k) => (get(ctx, k) || '').slice(0, 400)).join(''))) warnings.push('hiddenSlides');
  }
  if (o.exif) for (const [n, e] of ctx.z) if (/^(word|xl|ppt)\/media\/[^/]+\.(jpe?g|png)$/i.test(n)) { const c = cleanImage(n, ZP.inflate(e)); if (c) { ctx.changes[n] = { data: c, store: /\.jpe?g$/i.test(n) }; rep.exif++; } }
  if (ctx.ct !== ct0) ctx.changes['[Content_Types].xml'] = ctx.ct;
  if (/vbaProject\.bin/.test([...ctx.z.keys()].join('\n'))) warnings.push('macros');
  return { data: ZP.patchZip(buf, ctx.changes), report: rep, warnings };
}

/* ---------------- ODF ---------------- */
function cleanOdf(buf, ext, o) {
  const list = ZP.entries(buf);
  const z = new Map(list.map((e) => [e.name, e]));
  const changes = {};
  const rep = { author: false, comments: 0, revisions: 0, hidden: 0, custom: false, thumbnail: false, rsid: 0, exif: 0, paths: 0 };
  const warnings = [];
  const get0 = (n) => (z.has(n) ? ZP.inflate(z.get(n)).toString('utf8') : null);
  let man = get0('META-INF/manifest.xml') || '';
  if (/encryption-data/.test(man)) throw Object.assign(new Error('Decrypt the file first, then clean it.'), { status: 409 });
  let meta = get0('meta.xml');
  if (meta) {
    const b = meta;
    if (o.author) meta = meta.replace(/<(meta:initial-creator|dc:creator|meta:printed-by)\b[^>]*>[\s\S]*?<\/\1>/g, '').replace(/<meta:print-date\b[^>]*>[\s\S]*?<\/meta:print-date>/g, '').replace(/<meta:editing-cycles>\d*<\/meta:editing-cycles>/g, '').replace(/<meta:editing-duration>[^<]*<\/meta:editing-duration>/g, '');
    if (o.custom) { const c0 = meta; meta = meta.replace(/<meta:user-defined\b[^>]*(?:\/>|>[\s\S]*?<\/meta:user-defined>)/g, ''); if (c0 !== meta) rep.custom = true; }
    if (o.templates) meta = meta.replace(/<meta:template\b[^>]*\/>/g, () => { rep.paths++; return ''; });
    if (meta !== b) { changes['meta.xml'] = meta; rep.author = rep.author || /creator|printed/.test(b) !== /creator|printed/.test(meta) || true; }
  }
  for (const n of ['content.xml', 'styles.xml']) {
    let x = get0(n); if (!x) continue; const b = x;
    if (o.comments) { let r = strip(x, 'office:annotation'); x = r.xml; rep.comments += r.n; x = x.replace(/<office:annotation-end\b[^>]*\/>/g, ''); }
    if (o.revisions) {
      const r = strip(x, 'text:tracked-changes'); x = r.xml;
      const m = (x.match(/<text:change(-start|-end)?\b[^>]*\/>/g) || []).length; x = x.replace(/<text:change(-start|-end)?\b[^>]*\/>/g, '');
      rep.revisions += m + (r.n ? 0 : 0);
      if (ext === 'ods' || ext === 'ots') x = x.replace(/<table:tracked-changes\b[\s\S]*?<\/table:tracked-changes>/g, () => { rep.revisions++; return ''; });
    }
    if (o.hidden) {
      let r = strip(x, 'text:hidden-text'); x = r.xml; rep.hidden += r.n;
      r = replaceEl(x, 'text:p', (e) => (/<text:hidden-paragraph\b[^>]*text:is-hidden="true"/.test(e.inner) ? '' : undefined)); x = r.xml; rep.hidden += r.n;
      if (/text:display="none"/.test(x)) warnings.push('hiddenStyles');
    }
    if (x !== b) changes[n] = x;
  }
  if (o.thumbnail) for (const n of z.keys()) if (/^Thumbnails\//.test(n)) { changes[n] = null; man = man.replace(new RegExp(`<manifest:file-entry\\b[^>]*manifest:full-path="${esc(n)}"[^>]*/>`), ''); rep.thumbnail = true; }
  if (o.exif) for (const [n, e] of z) if (/^Pictures\/[^/]+\.(jpe?g|png)$/i.test(n)) { const c = cleanImage(n, ZP.inflate(e)); if (c) { changes[n] = { data: c, store: true }; rep.exif++; } }
  if (rep.thumbnail) changes['META-INF/manifest.xml'] = man;
  if (/^Basic\//m.test([...z.keys()].join('\n'))) warnings.push('macros');
  return { data: ZP.patchZip(buf, changes, { first: ['mimetype'] }), report: rep, warnings };
}

/* ---------------- PDF: in place, same length (xref offsets stay valid) ---------------- */
function cleanPdf(buf, o) {
  const b = Buffer.from(buf);
  const s = b.toString('latin1');
  const rep = { author: false, comments: 0, revisions: 0, hidden: 0, custom: false, thumbnail: false, rsid: 0, exif: 0, paths: 0, info: 0 };
  const warnings = [];
  if (/\/Encrypt\s/.test(s)) throw Object.assign(new Error('Encrypted PDFs cannot be cleaned.'), { status: 409 });
  const blank = (start, end) => { for (let i = start; i < end; i++) b[i] = 0x20; };
  // Info dictionary values: /Author (…) /Creator <…> … → spaces inside the delimiters
  const keys = o.author ? ['Author', 'Creator', 'Producer', 'Company', 'SourceModified', 'Manager'] : [];
  for (const k of keys) {
    const re = new RegExp(`/${k}\\s*(\\((?:\\\\.|[^\\\\)])*\\)|<[0-9A-Fa-f\\s]*>)`, 'g');
    let m;
    while ((m = re.exec(s))) { const v = m[1]; const st = m.index + m[0].length - v.length; blank(st + 1, st + v.length - 1); rep.info++; rep.author = true; }
  }
  // XMP packets (uncompressed): text of personal properties → spaces
  if (o.author) {
    const xmpRe = /<x:xmpmeta\b[\s\S]*?<\/x:xmpmeta>/g; let m;
    while ((m = xmpRe.exec(s))) {
      const x = m[0];
      const inner = /<(dc:creator|xmp:CreatorTool|pdf:Producer|xmpMM:History|xmpMM:DocumentID|xmpMM:InstanceID|pdfx:[\w-]+|photoshop:AuthorsPosition)\b[^>]*>([\s\S]*?)<\/\1>|\s(xmp:CreatorTool|pdf:Producer|xmpMM:DocumentID|xmpMM:InstanceID)="([^"]*)"/g;
      let k;
      while ((k = inner.exec(x))) {
        const val = k[2] != null ? k[2] : k[4];
        const off = m.index + k.index + (k[2] != null ? k[0].indexOf('>') : k[0].indexOf('"')) + 1;
        // keep the XML valid: blank only the text between tags
        let depth = 0;
        for (let i = 0; i < val.length; i++) { const c = val.charCodeAt(i); if (c === 0x3c) depth++; if (!depth) b[off + i] = 0x20; if (c === 0x3e) depth--; }
        rep.info++;
      }
    }
  }
  if (/\/ObjStm\b/.test(s) && !rep.info) warnings.push('pdfCompressed');
  else if (/\/ObjStm\b/.test(s)) warnings.push('pdfPartial');
  if (o.comments && /\/Annots\b/.test(s)) warnings.push('pdfAnnotations');
  if (/\/JavaScript\b|\/JS\b/.test(s)) warnings.push('pdfScripts');
  return { data: b, report: rep, warnings };
}

function cleanFile(buf, ext, opts = {}) {
  const o = { ...DEFAULTS, ...Object.fromEntries(Object.entries(opts || {}).filter(([k]) => k in DEFAULTS).map(([k, v]) => [k, v !== false])) };
  const e = String(ext || '').toLowerCase();
  const data = Buffer.isBuffer(buf) ? buf : Buffer.from(buf);
  if (/^(docx|docm|dotx|dotm|xlsx|xlsm|xltx|xltm|pptx|pptm|ppsx|potx)$/.test(e)) return cleanOoxml(data, e, o);
  if (/^(odt|ods|odp|odg|ott|ots|otp)$/.test(e)) return cleanOdf(data, e, o);
  if (e === 'pdf') return cleanPdf(data, o);
  if (/^(jpe?g|png)$/.test(e)) { const c = cleanImage(`x.${e}`, data); return { data: c || data, report: { exif: c ? 1 : 0 }, warnings: [] }; }
  throw Object.assign(new Error('This format cannot be cleaned.'), { status: 415 });
}

module.exports = { cleanFile, elements, replaceEl, cleanJpeg, cleanPng, DEFAULTS, S };
