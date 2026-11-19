'use strict';
const { makeZip } = require('../server/zip');
const CT = (over) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${over}</Types>`;
const REL = (t, target) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${t}" Target="${target}"/></Relationships>`;
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
function docx(text) {
  return makeZip({
    '[Content_Types].xml': CT('<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'),
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`,
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${String(text).split('\n').map((l) => `<w:p><w:r><w:t xml:space="preserve">${l}</w:t></w:r></w:p>`).join('')}</w:body></w:document>`,
    'docProps/core.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Test title</dc:title><dc:creator>Anna</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created></cp:coreProperties>',
    'docProps/app.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Pages>1</Pages><Words>4</Words></Properties>',
  });
}
function xlsx(rows) {
  const shared = []; const idx = (s) => { let i = shared.indexOf(s); if (i < 0) { shared.push(s); i = shared.length - 1; } return i; };
  const sheet = rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) => { const ref = String.fromCharCode(65 + ci) + (ri + 1); return typeof v === 'number' ? `<c r="${ref}"><v>${v}</v></c>` : `<c r="${ref}" t="s"><v>${idx(v)}</v></c>`; }).join('')}</row>`).join('');
  return makeZip({
    '[Content_Types].xml': CT('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'),
    '_rels/.rels': REL(OD, 'xl/workbook.xml'),
    'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheet}<row r="${rows.length + 1}"><c r="A${rows.length + 1}" t="inlineStr"><is><t>Inline note</t></is></c></row></sheetData></worksheet>`,
    'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${shared.length}" uniqueCount="${shared.length}">${shared.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>`,
  });
}
function pptx(slides) {
  const n = slides.length;
  const files = {
    '[Content_Types].xml': CT(`<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slides.map((s, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`),
    '_rels/.rels': REL(OD, 'ppt/presentation.xml'),
    'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst>${slides.map((s, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
    'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Slides>${n}</Slides></Properties>`,
  };
  slides.forEach((s, i) => {
    files[`ppt/slides/slide${i + 1}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody>${s.text.map((t) => `<a:p><a:r><a:t>${t}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
    if (s.notes) files[`ppt/notesSlides/notesSlide${i + 1}.xml`] = `<?xml version="1.0" encoding="UTF-8"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>${s.notes}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
  });
  return makeZip(files);
}
function odt(paragraphs, meta = {}) {
  return makeZip([
    { name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', store: true },
    { name: 'META-INF/manifest.xml', data: '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>' },
    { name: 'content.xml', data: `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2"><office:body><office:text>${paragraphs.map((p, i) => (i === 0 ? `<text:h text:outline-level="1">${p}</text:h>` : `<text:p>${p.replace(/  /g, '<text:s/> ')}</text:p>`)).join('')}</office:text></office:body></office:document-content>` },
    { name: 'meta.xml', data: `<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:title>${meta.title || ''}</dc:title><meta:initial-creator>${meta.author || ''}</meta:initial-creator><meta:document-statistic meta:page-count="${meta.pages || 1}" meta:word-count="${meta.words || 0}"/></office:meta></office:document-meta>` },
    { name: 'Thumbnails/thumbnail.png', data: Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex'), store: true },
  ]);
}
/* ---------- OLE compound file (CFB) builder for .doc / .ppt tests ---------- */
const END = 0xfffffffe; const FREE = 0xffffffff;
/** streams: { name: Buffer } (each padded to ≥ 4096 bytes so no mini stream is needed) */
function cfb(streams) {
  const list = Object.entries(streams).map(([name, data]) => { const n = Math.max(4096, Math.ceil(data.length / 512) * 512); const b = Buffer.alloc(n); data.copy(b); return { name, data: b, size: Math.max(4096, data.length) }; });
  const fat = [0xfffffffd, END];
  let next = 2;
  for (const s of list) { s.start = next; const n = s.data.length / 512; for (let i = 0; i < n; i++) fat.push(i === n - 1 ? END : next + i + 1); next += n; }
  const fatSec = Buffer.alloc(512, 0xff); fat.forEach((v, i) => fatSec.writeUInt32LE(v >>> 0, i * 4));
  const dir = Buffer.alloc(512);
  const entry = (i, name, type, child, right, start, size) => { const o = i * 128; const nb = Buffer.from(`${name}\0`, 'utf16le'); nb.copy(dir, o); dir.writeUInt16LE(nb.length, o + 0x40); dir[o + 0x42] = type; dir[o + 0x43] = 1; dir.writeUInt32LE(FREE, o + 0x44); dir.writeUInt32LE(right, o + 0x48); dir.writeUInt32LE(child, o + 0x4c); dir.writeUInt32LE(start, o + 0x74); dir.writeUInt32LE(size, o + 0x78); };
  entry(0, 'Root Entry', 5, list.length ? 1 : FREE, FREE, END, 0);
  list.forEach((s, i) => entry(i + 1, s.name, 2, FREE, i + 2 <= list.length ? i + 2 : FREE, s.start, s.size));
  const h = Buffer.alloc(512);
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(h, 0);
  h.writeUInt16LE(0x3e, 0x18); h.writeUInt16LE(3, 0x1a); h.writeUInt16LE(0xfffe, 0x1c); h.writeUInt16LE(9, 0x1e); h.writeUInt16LE(6, 0x20);
  h.writeUInt32LE(1, 0x2c); h.writeUInt32LE(1, 0x30); h.writeUInt32LE(4096, 0x38); h.writeUInt32LE(END, 0x3c); h.writeUInt32LE(0, 0x40); h.writeUInt32LE(END, 0x44); h.writeUInt32LE(0, 0x48);
  for (let i = 0; i < 109; i++) h.writeUInt32LE(i === 0 ? 0 : FREE, 0x4c + i * 4);
  return Buffer.concat([h, fatSec, dir, ...list.map((s) => s.data)]);
}
/** Word 97 document with one compressed (8-bit) piece */
function doc(text) {
  const wd = Buffer.alloc(4096);
  wd.writeUInt16LE(0xa5ec, 0); wd.writeUInt16LE(0x0200, 0x0a); // fWhichTblStm → 1Table
  wd.writeUInt16LE(14, 32); wd.writeUInt16LE(22, 62); wd.writeInt32LE(text.length, 76); wd.writeUInt16LE(93, 152);
  const at = 2048; Buffer.from(text, 'latin1').copy(wd, at);
  const table = Buffer.alloc(64);
  table[0] = 0x02; table.writeUInt32LE(16, 1); table.writeUInt32LE(0, 5); table.writeUInt32LE(text.length, 9); table.writeUInt16LE(0, 13); table.writeUInt32LE(((at * 2) | 0x40000000) >>> 0, 15); table.writeUInt16LE(0, 19);
  wd.writeUInt32LE(0, 154 + 66 * 8); wd.writeUInt32LE(21, 154 + 66 * 8 + 4);
  return cfb({ WordDocument: wd, '1Table': table });
}
/** PowerPoint 97 file: slide list with one text atom per slide */
function ppt(slides) {
  const rec = (verInst, type, body) => { const h = Buffer.alloc(8); h.writeUInt16LE(verInst, 0); h.writeUInt16LE(type, 2); h.writeUInt32LE(body.length, 4); return Buffer.concat([h, body]); };
  const items = slides.flatMap((t, i) => [rec(0, 0x03f3, Buffer.alloc(20)), i % 2 ? rec(0, 0x0fa8, Buffer.from(t, 'latin1')) : rec(0, 0x0fa0, Buffer.from(t, 'utf16le'))]);
  const list = rec(0x000f, 0x0ff0, Buffer.concat(items));
  return cfb({ 'PowerPoint Document': rec(0x000f, 0x03e8, list) });
}
module.exports = { docx, xlsx, pptx, odt, cfb, doc, ppt };
