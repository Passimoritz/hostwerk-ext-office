/*
 * Docs chunk – pure conversion logic (node --test, no browser):
 * RTF, ODT/FODT, DOCX (fresh + lossless patch), Markdown, templates, CFB + Word 97 reader, VBA decompression.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { defaultModel, para, heading, bullets, tasks, table, b, txt, docText, walk } from '../ui/src/editors/docs/io/pm.js';
import { parseRtf, writeRtf } from '../ui/src/editors/docs/io/rtf.js';
import { readOdt } from '../ui/src/editors/docs/io/odt-read.js';
import { writeOdt } from '../ui/src/editors/docs/io/odt-write.js';
import { readDocx } from '../ui/src/editors/docs/io/docx-read.js';
import { writeDocx } from '../ui/src/editors/docs/io/docx-write.js';
import { mdToHtml, htmlToMd } from '../ui/src/editors/docs/io/md.js';
import { docToHtml, standaloneHtml } from '../ui/src/editors/docs/io/render-html.js';
import { writeDocument, readDocument, documentText, sniff } from '../ui/src/editors/docs/io/index.js';
import { templateModel } from '../ui/src/editors/docs/templates.js';
import { readCfb } from '../ui/src/editors/_shared/cfb.js';
import { readDoc } from '../ui/src/editors/docs/io/doc-read.js';
import { decompressVba } from '../ui/src/editors/docs/io/vba.js';
import { equationText, equationOmml } from '../ui/src/editors/docs/io/equation.js';

const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR42mP8z8DwnwEIGBkZGBgYAAAw/gH+3x3s5wAAAABJRU5ErkJggg==', 'base64'));

function sample() {
  const m = defaultModel({ general: { author: 'Ann' } });
  m.doc = { type: 'doc', content: [
    heading(1, 'Hello world'),
    para(['Some ', b('bold'), ' and ', txt('commented', [{ type: 'comment', attrs: { id: 'x1' } }]), ' text', { type: 'footnote', attrs: { text: 'A note' } }, txt(' inserted', [{ type: 'insertion', attrs: { id: '1', author: 'Ann', date: '2024-01-02T03:04:00Z' } }]), txt(' gone', [{ type: 'deletion', attrs: { id: '2', author: 'Bob', date: null } }])], { textAlign: 'center' }),
    bullets(['one', 'two']),
    bullets(['first', 'second'], true),
    tasks([[true, 'done'], [false, 'open']]),
    table([['A', 'B'], ['1', '2']]),
    para([{ type: 'image', attrs: { src: `data:image/png;base64,${Buffer.from(PNG).toString('base64')}`, alt: 'dot', width: 40 } }]),
    { type: 'pageBreak' },
    { type: 'toc', attrs: { title: 'Contents', entries: [{ level: 1, text: 'Hello world' }] } },
    para(['E: ', { type: 'equation', attrs: { latex: '\\frac{a}{b}+x^2' } }]),
    { type: 'callout', attrs: { kind: 'warn' }, content: [para('Careful')] },
  ] };
  m.comments = [{ id: 'x1', author: 'Ann', date: '2024-01-02T03:04:00Z', text: 'Check this', resolved: true, replies: [{ author: 'Bob', date: null, text: 'Done' }] }];
  m.header = { text: 'Page {page} of {pages}', align: 'right' };
  m.page.orientation = 'landscape';
  m.page.margins = { top: 20, right: 15, bottom: 20, left: 30 };
  return m;
}
const types = (doc) => { const t = {}; walk(doc, (n) => { t[n.type] = (t[n.type] || 0) + 1; }); return t; };
const marksOf = (doc, type) => { const out = []; walk(doc, (n) => { for (const m of n.marks || []) if (m.type === type) out.push({ text: n.text, attrs: m.attrs }); }); return out; };

test('RTF: write → parse keeps text, formatting, lists, tables, footnotes, header, page', () => {
  const rtf = writeRtf(sample());
  assert.match(rtf, /^\{\\rtf1/);
  const m = parseRtf(rtf);
  const text = docText(m.doc);
  for (const s of ['Hello world', 'Some bold', 'one', 'second', 'done', 'Careful', 'A note']) assert.ok(text.includes(s), `missing ${s}`);
  const t = types(m.doc);
  assert.equal(t.heading, 1);
  assert.ok(t.bulletList >= 1 && t.orderedList >= 1 && t.taskList >= 1 && t.table >= 1 && t.footnote === 1 && t.pageBreak === 1 && t.image === 1);
  assert.ok(marksOf(m.doc, 'bold').some((x) => x.text === 'bold'));
  assert.ok(marksOf(m.doc, 'insertion').some((x) => x.attrs.author === 'Ann'));
  assert.equal(m.header.text, 'Page {page} of {pages}');
  assert.equal(m.page.orientation, 'landscape');
  assert.equal(m.page.margins.left, 30);
  assert.equal(m.comments.length, 1);
});

test('RTF: reads Word-style input (fonts, colours, unicode, codepage, hyperlinks)', () => {
  const src = String.raw`{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0 Calibri;}{\f1 Arial;}}{\colortbl ;\red255\green0\blue0;}\paperw11906\paperh16838
\pard\qc\b Titel\b0\par
\pard {\f1\fs28\cf1 Gr\'fc\'dfe} \u8364? {\field{\*\fldinst{HYPERLINK "https://example.org"}}{\fldrslt{Link}}}\par}`;
  const m = parseRtf(src);
  const [p1, p2] = m.doc.content;
  assert.equal(p1.attrs.textAlign, 'center');
  assert.equal(p1.content[0].text, 'Titel');
  const text = docText(m.doc);
  assert.ok(text.includes('Grüße') && text.includes('€') && text.includes('Link'));
  const ts = marksOf(m.doc, 'textStyle').find((x) => x.text === 'Grüße');
  assert.equal(ts.attrs.fontFamily, 'Arial'); assert.equal(ts.attrs.fontSize, '14pt'); assert.equal(ts.attrs.color, '#FF0000');
  assert.equal(marksOf(m.doc, 'link')[0].attrs.href, 'https://example.org');
  assert.equal(m.page.size, 'A4');
  void p2;
});

test('ODT: write → read round trip (comments with replies, tracked changes, footnotes, page setup)', async () => {
  const bytes = await writeOdt(sample());
  const zip = await JSZip.loadAsync(bytes);
  // ODF: mimetype is the first entry and stored uncompressed
  assert.equal(Buffer.from(bytes.subarray(30, 38)).toString(), 'mimetype');
  assert.equal(bytes[8] | (bytes[9] << 8), 0, 'mimetype must be STORED');
  assert.equal(await zip.file('mimetype').async('string'), 'application/vnd.oasis.opendocument.text');
  for (const f of ['content.xml', 'styles.xml', 'meta.xml', 'META-INF/manifest.xml']) assert.ok(zip.file(f), f);
  assert.ok(Object.keys(zip.files).some((f) => f.startsWith('Pictures/')));
  const m = await readOdt(bytes);
  const t = types(m.doc);
  assert.ok(t.heading === 1 && t.bulletList && t.orderedList && t.taskList && t.table && t.footnote === 1 && t.toc === 1 && t.callout === 1 && t.image === 1 && t.pageBreak === 1);
  assert.equal(m.comments.length, 1);
  assert.equal(m.comments[0].replies.length, 1);
  assert.equal(m.comments[0].resolved, true);
  assert.ok(marksOf(m.doc, 'insertion').some((x) => x.text.includes('inserted') && x.attrs.author === 'Ann'));
  assert.ok(marksOf(m.doc, 'deletion').some((x) => x.text.includes('gone') && x.attrs.author === 'Bob'));
  assert.equal(m.header.text, 'Page {page} of {pages}');
  assert.equal(m.page.orientation, 'landscape');
  assert.equal(m.page.margins.left, 30);
  // template mimetype
  const ott = await writeOdt(sample(), { ext: 'ott' });
  assert.equal(await (await JSZip.loadAsync(ott)).file('mimetype').async('string'), 'application/vnd.oasis.opendocument.text-template');
});

test('ODT: patching keeps unknown parts byte-identical', async () => {
  const first = await JSZip.loadAsync(await writeOdt(sample()));
  first.file('Basic/Standard/Module1.xml', '<macro>keep me</macro>');
  first.file('Object 1/content.xml', '<x>embedded</x>');
  const orig = await first.generateAsync({ type: 'uint8array' });
  const m = await readOdt(orig);
  m.doc.content[0].content[0].text = 'Changed';
  const out = await JSZip.loadAsync(await writeOdt(m));
  assert.equal(await out.file('Basic/Standard/Module1.xml').async('string'), '<macro>keep me</macro>');
  assert.equal(await out.file('Object 1/content.xml').async('string'), '<x>embedded</x>');
  const again = await readOdt(await out.generateAsync({ type: 'uint8array' }));
  assert.equal(again.doc.content[0].content[0].text, 'Changed');
  assert.ok(again.fidelity.includes('macros'));
});

test('FODT: flat XML round trip', async () => {
  const bytes = await writeOdt(sample(), { ext: 'fodt' });
  assert.match(Buffer.from(bytes).toString('utf8'), /office:mimetype="application\/vnd\.oasis\.opendocument\.text"/);
  assert.equal(sniff(bytes, 'fodt'), 'fodt');
  const m = await readOdt(bytes);
  assert.ok(docText(m.doc).includes('Hello world'));
  assert.equal(types(m.doc).image, 1);
});

test('DOCX: fresh package is a valid zip with document.xml containing the text', async () => {
  const bytes = await writeDocx(sample());
  const zip = await JSZip.loadAsync(bytes);
  for (const f of ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml', 'word/settings.xml', 'word/numbering.xml', 'word/footnotes.xml', 'word/comments.xml', 'word/commentsExtended.xml', 'docProps/core.xml']) assert.ok(zip.file(f), f);
  assert.equal(Object.keys(zip.files)[0], '[Content_Types].xml');
  const doc = await zip.file('word/document.xml').async('string');
  for (const s of ['Hello world', 'bold', 'w:ins ', 'w:delText', 'commentRangeStart', 'footnoteReference', 'w:orient="landscape"', 'w:tblCaption w:val="of-callout:warn"', 'm:oMath', 'TOC \\o']) assert.ok(doc.includes(s), `document.xml lacks ${s}`);
  assert.match(await zip.file('docProps/core.xml').async('string'), /<dc:creator>Ann<\/dc:creator>/);
  const m = await readDocx(bytes);
  const t = types(m.doc);
  assert.ok(t.heading === 1 && t.bulletList && t.orderedList && t.taskList === 1 && t.table && t.footnote === 1 && t.toc === 1 && t.callout === 1 && t.image === 1 && t.pageBreak === 1 && t.equation === 1);
  assert.equal(m.comments[0].text, 'Check this');
  assert.equal(m.comments[0].replies[0].text, 'Done');
  assert.equal(m.comments[0].resolved, true);
  assert.ok(marksOf(m.doc, 'deletion').some((x) => x.attrs.author === 'Bob'));
  assert.equal(m.header.text, 'Page {page} of {pages}');
  assert.equal(m.page.orientation, 'landscape');
  assert.equal(m.doc.content[1].attrs.textAlign, 'center');
  assert.equal(equationText('\\frac{a}{b}+x^2'), 'a/b+x²');
  assert.match(equationOmml('x^2'), /<m:sSup>/);
});

test('DOCX: lossless patch keeps macros, customXml and unknown properties', async () => {
  // a "foreign" docm: our package + vbaProject + customXml + an unknown run/paragraph property
  const zip = await JSZip.loadAsync(await writeDocx(sample()));
  zip.file('word/vbaProject.bin', new Uint8Array([1, 2, 3, 4, 5]));
  zip.file('customXml/item1.xml', '<root>data</root>');
  let rels = await zip.file('word/_rels/document.xml.rels').async('string');
  rels = rels.replace('</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', rels);
  let doc = await zip.file('word/document.xml').async('string');
  doc = doc.replace('<w:p><w:pPr><w:pStyle w:val="Heading1"/>', '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:keepLines/><w:pBdr><w:top w:val="single" w:sz="4"/></w:pBdr>');
  doc = doc.replace('<w:t xml:space="preserve">Hello world</w:t>', '<w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:rPr><w:caps/></w:rPr><w:t>world</w:t>');
  zip.file('word/document.xml', doc);
  const orig = await zip.generateAsync({ type: 'uint8array' });
  const m = await readDocx(orig, {}, { ext: 'docm' });
  assert.ok(m.fidelity.includes('macros'));
  m.doc.content.push(para('Added later'));
  const out = await JSZip.loadAsync(await writeDocx(m, { ext: 'docm' }));
  assert.deepEqual([...await out.file('word/vbaProject.bin').async('uint8array')], [1, 2, 3, 4, 5]);
  assert.equal(await out.file('customXml/item1.xml').async('string'), '<root>data</root>');
  assert.match(await out.file('word/_rels/document.xml.rels').async('string'), /vbaProject/);
  const d2 = await out.file('word/document.xml').async('string');
  assert.ok(d2.includes('Added later'));
  assert.ok(d2.includes('<w:caps/>'), 'unknown run property kept');
  assert.ok(/<w:pStyle w:val="Heading1"\/><w:keepLines\/><w:pBdr>/.test(d2), 'unknown paragraph properties kept in schema order');
});

test('DOCX: preserved nodes (SmartArt/shape) are written back verbatim', async () => {
  const zip = await JSZip.loadAsync(await writeDocx(sample()));
  let doc = await zip.file('word/document.xml').async('string');
  const shape = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:txbx><w:txbxContent><w:p><w:r><w:t>In the box</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>';
  doc = doc.replace('<w:body>', `<w:body><w:p>${shape}</w:p>`);
  zip.file('word/document.xml', doc);
  const m = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  const kept = []; walk(m.doc, (n) => { if (n.type === 'preserved') kept.push(n.attrs); });
  assert.equal(kept.length, 1);
  assert.equal(kept[0].kind, 'textbox');
  assert.equal(kept[0].text, 'In the box');
  const out = await JSZip.loadAsync(await writeDocx(m));
  assert.ok((await out.file('word/document.xml').async('string')).includes('<w:t>In the box</w:t></w:r></w:p></w:txbxContent>'));
});

test('Markdown: md → html → md round trip', () => {
  const md = '# Title\n\nSome **bold** and *italic* text with a [link](https://example.org).\n\n- one\n- two\n\n1. first\n2. second\n\n- [x] done\n- [ ] open\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n> quoted\n\n```\ncode\n```\n\nA note[^1].\n\n[^1]: The note text\n';
  const html = mdToHtml(md);
  assert.match(html, /<h1>Title<\/h1>/);
  assert.match(html, /data-type="taskList"/);
  assert.match(html, /data-type="footnote" data-text="The note text"/);
  const back = htmlToMd(html);
  for (const s of ['# Title', '**bold**', '*italic*', '[link](https://example.org)', '- one', '1. first', '- [x] done', '- [ ] open', '| A | B |', '> quoted', '```', '[^1]: The note text']) assert.ok(back.includes(s), `md lacks ${s}\n${back}`);
});

test('HTML export carries page setup, comments and footnotes', () => {
  const html = standaloneHtml(sample());
  assert.match(html, /<meta name="of-docs" content="/);
  assert.match(html, /@page\{size:297mm 210mm/);
  assert.match(html, /counter\(page\)/);
  assert.match(html, /class="of-docs-footnotes"/);
  assert.match(docToHtml(sample().doc), /data-type="toc"/);
});

test('templates: every docs template in every writable format', async () => {
  for (const id of ['blank', 'letter', 'invoice', 'cv', 'report', 'minutes', 'notes']) {
    for (const lang of ['de', 'en']) {
      const m = templateModel(id, { general: { author: 'Ann' } }, lang);
      for (const ext of ['docx', 'odt', 'rtf', 'md', 'html', 'txt']) {
        const bytes = await writeDocument(m, ext);
        assert.ok(bytes.length > 0 || id === 'blank', `${id}.${ext}`);
        const text = await documentText(bytes, ext);
        if (id === 'report') assert.ok(text.includes(lang === 'de' ? 'Jahresbericht' : 'Annual report'), `${id}.${ext} text`);
      }
    }
  }
  const r = templateModel('report', {}, 'en');
  const toc = r.doc.content.find((n) => n.type === 'toc');
  assert.ok(toc.attrs.entries.length >= 5);
});

test('readDocument sniffs the real format (a .doc that is really RTF)', async () => {
  const rtf = new TextEncoder().encode('{\\rtf1\\ansi Hello\\par}');
  const { model } = await readDocument(rtf, 'doc');
  assert.equal(docText(model.doc).trim(), 'Hello');
});

/* ---------- CFB + Word 97 ---------- */
function buildCfb(streams) {
  const SEC = 512; const END = 0xfffffffe; const FREE = 0xffffffff;
  const names = Object.keys(streams);
  const secCount = names.map((n) => Math.ceil(streams[n].length / SEC));
  const total = 2 + secCount.reduce((a, b) => a + b, 0);
  const out = new Uint8Array(SEC * (total + 1));
  const dv = new DataView(out.buffer);
  [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].forEach((v, i) => { out[i] = v; });
  dv.setUint16(0x18, 0x3e, true); dv.setUint16(0x1a, 3, true); dv.setUint16(0x1c, 0xfffe, true); dv.setUint16(0x1e, 9, true); dv.setUint16(0x20, 6, true);
  dv.setUint32(0x2c, 1, true); dv.setUint32(0x30, 1, true); dv.setUint32(0x38, 4096, true); dv.setUint32(0x3c, END, true); dv.setUint32(0x44, END, true);
  for (let k = 0; k < 109; k++) dv.setUint32(0x4c + k * 4, k === 0 ? 0 : FREE, true);
  const fat = (k, v) => dv.setUint32(SEC + k * 4, v, true);
  for (let k = 0; k < 128; k++) fat(k, FREE);
  fat(0, 0xfffffffd); fat(1, END);
  let next = 2; const starts = [];
  names.forEach((n, i) => { starts.push(next); for (let k = 0; k < secCount[i]; k++) fat(next + k, k === secCount[i] - 1 ? END : next + k + 1); out.set(streams[n], SEC * (next + 1)); next += secCount[i]; });
  const dir = (idx, name, type, child, right, start, size) => {
    const o = SEC * 2 + idx * 128;
    for (let k = 0; k < name.length; k++) dv.setUint16(o + k * 2, name.charCodeAt(k), true);
    dv.setUint16(o + 0x40, (name.length + 1) * 2, true); out[o + 0x42] = type; out[o + 0x43] = 1;
    dv.setUint32(o + 0x44, FREE, true); dv.setUint32(o + 0x48, right, true); dv.setUint32(o + 0x4c, child, true);
    dv.setUint32(o + 0x74, start, true); dv.setUint32(o + 0x78, size, true);
  };
  dir(0, 'Root Entry', 5, 1, FREE, END, 0);
  names.forEach((n, i) => dir(i + 1, n, 2, FREE, i + 2 <= names.length ? i + 2 : FREE, starts[i], streams[n].length));
  for (let k = names.length + 1; k < 4; k++) { const o = SEC * 2 + k * 128; dv.setUint32(o + 0x44, FREE, true); dv.setUint32(o + 0x48, FREE, true); dv.setUint32(o + 0x4c, FREE, true); }
  return out;
}
function buildWord97() {
  const text = 'Hello World\rSecond para\r';
  const wd = new Uint8Array(5120); const w = new DataView(wd.buffer);
  w.setUint16(0, 0xa5ec, true); w.setUint16(2, 0xc1, true); w.setUint16(0x0a, 0, true);
  w.setUint16(0x20, 14, true); w.setUint16(0x3e, 22, true); w.setInt32(0x40 + 12, text.length, true); w.setUint16(0x98, 93, true);
  const fcText = 0x800;
  wd.set(new TextEncoder().encode(text), fcText);
  // CHPX FKP at page 8: run 0 "Hello" bold, run 1 rest plain
  const base = 8 * 512;
  w.setUint32(base, fcText, true); w.setUint32(base + 4, fcText + 5, true); w.setUint32(base + 8, fcText + text.length, true);
  wd[base + 12] = 0x80; wd[base + 13] = 0; wd[base + 511] = 2;
  wd.set([3, 0x35, 0x08, 0x01], base + 0x100);
  const tbl = new Uint8Array(4096); const t = new DataView(tbl.buffer);
  // Clx at 0x100: Pcdt with one piece (compressed 8-bit text)
  tbl[0x100] = 2; t.setUint32(0x101, 16, true); t.setUint32(0x105, 0, true); t.setUint32(0x109, text.length, true); t.setUint16(0x10d, 0, true); t.setUint32(0x10f, (fcText * 2) | 0x40000000, true); t.setUint16(0x113, 0, true);
  // PlcBteChpx at 0x200: fcs + one PnFkpChpx
  t.setUint32(0x200, fcText, true); t.setUint32(0x204, fcText + text.length, true); t.setUint32(0x208, 8, true);
  const fl = (k, fc, lcb) => { w.setUint32(0x9a + k * 8, fc, true); w.setUint32(0x9a + k * 8 + 4, lcb, true); };
  fl(33, 0x100, 21); fl(12, 0x200, 12);
  return buildCfb({ WordDocument: wd, '0Table': tbl });
}

test('CFB reader lists streams', () => {
  const cfb = readCfb(buildWord97());
  assert.deepEqual([...cfb.streams.keys()].sort(), ['0Table', 'WordDocument']);
  assert.equal(cfb.streams.get('WordDocument').length, 5120);
  assert.throws(() => readCfb(new Uint8Array(600)), /Not a compound file/);
});

test('Word 97 reader: piece table text and CHPX bold', () => {
  const m = readDoc(buildWord97());
  assert.equal(m.doc.content.length, 2);
  assert.equal(docText(m.doc), 'Hello World\nSecond para');
  const first = m.doc.content[0].content;
  assert.equal(first[0].text, 'Hello');
  assert.deepEqual(first[0].marks, [{ type: 'bold' }]);
  assert.equal(first[1].text, ' World');
});

test('VBA: MS-OVBA decompression (spec example)', () => {
  const comp = Uint8Array.from([0x01, 0x2f, 0xb0, 0x00, 0x23, 0x61, 0x61, 0x61, 0x62, 0x63, 0x64, 0x65, 0x82, 0x66, 0x00, 0x70, 0x61, 0x67, 0x68, 0x69, 0x6a, 0x01, 0x38, 0x08, 0x61, 0x6b, 0x6c, 0x00, 0x30, 0x6d, 0x6e, 0x6f, 0x70, 0x06, 0x71, 0x02, 0x70, 0x04, 0x10, 0x72, 0x73, 0x74, 0x75, 0x76, 0x10, 0x77, 0x78, 0x79, 0x7a, 0x00, 0x3c]);
  assert.equal(Buffer.from(decompressVba(comp)).toString('latin1'), '#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa');
});

/* ---------- regressions of the Microsoft Office check (Word 16) ---------- */
const plain = () => { const m = defaultModel({}); m.doc = { type: 'doc', content: [heading(1, 'One'), para('Text'), heading(2, 'Two')] }; return m; };
const part = async (bytes, name) => { const z = await JSZip.loadAsync(bytes); const f = z.file(name); return f ? f.async('string') : null; };

test('Office #1: settings name separator footnotes only when footnotes.xml exists', async () => {
  const none = await writeDocx(plain());
  assert.equal(await part(none, 'word/footnotes.xml'), null);
  assert.doesNotMatch(await part(none, 'word/settings.xml'), /<w:footnote w:id=|<w:endnote w:id=/);
  // a foreign package whose settings name separators although the notes parts are gone
  const z = await JSZip.loadAsync(await writeDocx(plain()));
  z.file('word/settings.xml', (await z.file('word/settings.xml').async('string')).replace('</w:settings>', '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr></w:settings>'));
  const m = await readDocx(await z.generateAsync({ type: 'uint8array' }));
  m.doc.content.push(para('Added'));
  const out = await writeDocx(m);
  const st = await part(out, 'word/settings.xml');
  assert.equal(await part(out, 'word/footnotes.xml'), null);
  assert.doesNotMatch(st, /<w:footnote w:id=|<w:endnote w:id=/);
  const withNotes = await writeDocx(sample());
  assert.ok(await part(withNotes, 'word/footnotes.xml'));
  assert.match(await part(withNotes, 'word/settings.xml'), /<w:footnote w:id="0"\/>/);
});

test('Office #2: a docm saved as docx/dotx gets the right main type and loses the VBA chain', async () => {
  const zip = await JSZip.loadAsync(await writeDocx(sample(), { ext: 'docm' }));
  zip.file('word/vbaProject.bin', new Uint8Array([1, 2, 3]));
  zip.file('word/vbaData.xml', '<wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>');
  zip.file('word/_rels/vbaProject.bin.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/wordVbaData" Target="vbaData.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels').async('string')).replace('</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('</Types>', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/></Types>'));
  const orig = await zip.generateAsync({ type: 'uint8array' });
  const m = await readDocx(orig, {}, { ext: 'docm' });
  for (const [ext, type] of [['docx', 'document'], ['dotx', 'template']]) {
    const out = await JSZip.loadAsync(await writeDocx(m, { ext }));
    const ct = await out.file('[Content_Types].xml').async('string');
    assert.ok(ct.includes(`PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}.main+xml"`), `${ext}: ${ct}`);
    assert.doesNotMatch(ct, /macroEnabled|vbaProject|vbaData/, ext);
    assert.equal(out.file('word/vbaProject.bin'), null, ext);
    assert.equal(out.file('word/vbaData.xml'), null, ext);
    assert.equal(out.file('word/_rels/vbaProject.bin.rels'), null, ext);
    assert.doesNotMatch(await out.file('word/_rels/document.xml.rels').async('string'), /vbaProject/, ext);
  }
  const keep = await JSZip.loadAsync(await writeDocx(m, { ext: 'docm' }));
  assert.ok((await keep.file('[Content_Types].xml').async('string')).includes('PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"'));
  assert.ok(keep.file('word/vbaProject.bin'));
});

test('Office #6: ODT title page = first-page master (no header-first), read back from ours and from header-first', async () => {
  const m = plain();
  m.header = { text: 'Head {page}', align: 'left' }; m.titlePage = true;
  for (const ext of ['odt', 'fodt']) {
    const bytes = await writeOdt(m, { ext });
    const all = ext === 'odt' ? (await part(bytes, 'content.xml')) + (await part(bytes, 'styles.xml')) : Buffer.from(bytes).toString('utf8');
    assert.doesNotMatch(all, /<style:header-first\b/, ext);
    assert.match(all, /<style:master-page style:name="OfFirstPage"[^>]*style:next-style-name="Standard"/, ext);
    assert.match(all, /style:master-page-name="OfFirstPage"/, ext);
    const r = await readOdt(bytes);
    assert.equal(r.titlePage, true, ext);
    assert.equal(r.header.text, 'Head {page}', ext);
  }
  // without a title page nothing of that is written
  const n = plain(); n.header = { text: 'Head', align: 'left' };
  const nb = await writeOdt(n);
  assert.doesNotMatch(await part(nb, 'styles.xml'), /OfFirstPage/);
  assert.notEqual((await readOdt(nb)).titlePage, true);
  // other producers (LibreOffice 7+): <style:header-first> inside the master
  const z = await JSZip.loadAsync(nb);
  z.file('styles.xml', (await z.file('styles.xml').async('string')).replace(/(<style:master-page style:name="Standard"[^>]*>)/, '$1<style:header-first><text:p/></style:header-first>'));
  assert.equal((await readOdt(await z.generateAsync({ type: 'uint8array' }))).titlePage, true);
});

test('Office #8: RTF declares \\ftnbj so Word keeps footnotes', () => {
  assert.match(writeRtf(sample()), /\\ftnbj/);
});

test('Office #11: TOC written with entries and _Toc bookmarks, no dirty field / updateFields', async () => {
  const m = plain();
  m.doc.content.unshift({ type: 'toc', attrs: { title: 'Contents', entries: [] } });
  const bytes = await writeDocx(m);
  const doc = await part(bytes, 'word/document.xml');
  assert.doesNotMatch(doc, /w:dirty=/);
  assert.doesNotMatch(await part(bytes, 'word/settings.xml'), /updateFields/);
  assert.match(doc, /TOC \\o/);
  const marks = [...doc.matchAll(/<w:bookmarkStart w:id="\d+" w:name="(_Toc\d+)"\/>/g)].map((x) => x[1]);
  assert.ok(marks.length >= 2, 'bookmarks on the headings');
  for (const bm of marks) assert.ok(doc.includes(`w:anchor="${bm}"`), `entry links to ${bm}`);
  assert.ok(doc.indexOf('>One<') < doc.lastIndexOf('>One<'), 'entry text filled in before the heading');
});

test('Office #9: Word 6/95 files get the engine message (DOC_TOO_OLD), not "not a Word document"', () => {
  for (const [wIdent, nFib] of [[0xa5dc, 101], [0xa5dc, 104], [0xa59b, 45], [0xa5ec, 0x68]]) {
    const wd = new Uint8Array(5120); const w = new DataView(wd.buffer);
    w.setUint16(0, wIdent, true); w.setUint16(2, nFib, true);
    const bytes = buildCfb({ WordDocument: wd });
    assert.throws(() => readDoc(bytes), (e) => e.code === 'DOC_TOO_OLD' && /too old for the built-in reader/.test(e.message) && /Office engine/.test(e.message), `${wIdent.toString(16)}/${nFib}`);
  }
  const wd = new Uint8Array(5120); new DataView(wd.buffer).setUint16(0, 0x1234, true);
  assert.throws(() => readDoc(buildCfb({ WordDocument: wd })), /not a Word 97–2003 document/);
});
