/*
 * Docs chunk 1.1 – the structures added in the 1.1 round (node --test, no browser):
 * sections with their own page setup, rich headers/footers (formatting, fields, pictures, tables, first/even pages),
 * content controls (check box, drop-down, date, plain text, block), live fields, equations (MathML → OMML / ODF formula
 * object), tracked paragraph marks / paragraph formatting / table rows, commentsIds + commentsExtensible, diagrams,
 * encrypted files, version compare.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import temml from 'temml';
import { defaultModel, para, heading, table, b, txt, walk, docText } from '../ui/src/editors/docs/io/pm.js';
import { readDocx } from '../ui/src/editors/docs/io/docx-read.js';
import { writeDocx } from '../ui/src/editors/docs/io/docx-write.js';
import { readOdt } from '../ui/src/editors/docs/io/odt-read.js';
import { writeOdt } from '../ui/src/editors/docs/io/odt-write.js';
import { readDocument } from '../ui/src/editors/docs/io/index.js';
import { mathmlToOmml, mathmlToLatex, texAnnotation } from '../ui/src/editors/docs/io/mathml.js';
import { compareDocs, applyAsTrackedChanges } from '../ui/src/editors/docs/io/compare.js';
import { docToHtml } from '../ui/src/editors/docs/io/render-html.js';

const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR42mP8z8DwnwEIGBkZGBgYAAAw/gH+3x3s5wAAAAJJRU5ErkJggg==', 'base64'));
const PNG_URL = `data:image/png;base64,${Buffer.from(PNG).toString('base64')}`;
const ml = (latex, display = false) => temml.renderToString(latex, { xml: true, annotate: true, displayMode: display });
const nodesOf = (doc, type) => { const out = []; walk(doc, (n) => { if (n.type === type) out.push(n); }); return out; };
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"';

/** a small Word package with the given body (and optional extra parts/rels) */
async function docxOf(body, { parts = {}, rels = '', settings = '' } = {}) {
  const z = new JSZip();
  z.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`);
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>${rels}</Relationships>`);
  z.file('word/settings.xml', `<?xml version="1.0" encoding="UTF-8"?><w:settings ${W_NS}>${settings}<w:defaultTabStop w:val="708"/></w:settings>`);
  z.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`);
  for (const [p, x] of Object.entries(parts)) z.file(p, x);
  return z.generateAsync({ type: 'uint8array' });
}
const docXml = async (bytes) => (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');

/* ---------------- equations ---------------- */
test('equations: MathML (temml) → OMML with fractions, operators, roots, fences, matrices, accents', () => {
  const o = mathmlToOmml(ml(String.raw`\frac{a}{b} + \sum_{i=1}^{n} i^2 + \sqrt[3]{x} + \left( \frac{1}{2} \right) + \hat{v} + \begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}`));
  assert.match(o, /^<m:oMath>/);
  for (const tag of ['<m:f>', '<m:nary>', '<m:chr m:val="∑"/>', '<m:rad><m:deg>', '<m:d>', '<m:acc>', '<m:m>', '<m:sSup>']) assert.ok(o.includes(tag), `missing ${tag}`);
  assert.equal((o.match(/<m:mr>/g) || []).length, 2);
  assert.match(mathmlToOmml(ml('x', true), { display: true }), /^<m:oMathPara><m:oMath>/);
  assert.equal(texAnnotation(ml(String.raw`\frac{a}{b}`)), String.raw`\frac{a}{b}`);
  assert.equal(mathmlToLatex('<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><mn>2</mn></mfrac></math>'), String.raw`\frac{a}{2}`);
});

test('equations: docx writes OMML from MathML (inline + display), reads back LaTeX and display; unchanged OMML stays', async () => {
  const m = defaultModel();
  m.doc = { type: 'doc', content: [
    para(['E: ', { type: 'equation', attrs: { latex: String.raw`\frac{a}{b}`, mathml: ml(String.raw`\frac{a}{b}`) } }]),
    para([{ type: 'equation', attrs: { latex: String.raw`\int_0^1 x\,dx`, mathml: ml(String.raw`\int_0^1 x\,dx`, true), display: true } }]),
  ] };
  const bytes = await writeDocx(m);
  const x = await docXml(bytes);
  assert.match(x, /<m:oMath><m:f><m:num>/);
  assert.match(x, /<m:oMathPara><m:oMath><m:nary>/);
  const back = await readDocx(bytes);
  const eqs = nodesOf(back.doc, 'equation');
  assert.equal(eqs.length, 2);
  assert.equal(eqs[0].attrs.latex, String.raw`\frac{a}{b}`);
  assert.equal(eqs[1].attrs.display, true);
  assert.match(eqs[1].attrs.latex, /^\\int_0\^1/);
  // a second save without edits writes the original OMML again
  const again = await docXml(await writeDocx(back));
  assert.ok(again.includes(eqs[0].attrs.omml));
});

test('equations: ODT gets a real formula object (MathML sub-document), read back with its LaTeX', async () => {
  const m = defaultModel();
  m.doc = { type: 'doc', content: [para(['x = ', { type: 'equation', attrs: { latex: String.raw`\sqrt{2}`, mathml: ml(String.raw`\sqrt{2}`) } }])] };
  const odt = await writeOdt(m);
  const z = await JSZip.loadAsync(odt);
  assert.ok(z.file('Object 1/content.xml'), 'formula object written');
  assert.match(await z.file('Object 1/content.xml').async('string'), /<msqrt>/);
  const man = await z.file('META-INF/manifest.xml').async('string');
  assert.match(man, /full-path="Object 1\/" manifest:media-type="application\/vnd\.oasis\.opendocument\.formula"/);
  assert.match(await z.file('content.xml').async('string'), /<draw:object xlink:href="\.\/Object 1"/);
  const back = await readOdt(odt);
  const eq = nodesOf(back.doc, 'equation')[0];
  assert.equal(eq.attrs.latex, String.raw`\sqrt{2}`);
  assert.match(eq.attrs.mathml, /<msqrt>/);
  // flat ODT: inline math:math
  const fodt = new TextDecoder().decode(await writeOdt(m, { ext: 'fodt' }));
  assert.match(fodt, /<draw:object><math:math[^>]*><math:semantics>/);
});

/* ---------------- content controls ---------------- */
const SDT_BODY = `<w:p><w:r><w:t xml:space="preserve">Agree: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="1"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt></w:p>
<w:p><w:sdt><w:sdtPr><w:alias w:val="Colour"/><w:tag w:val="colour"/><w:id w:val="2"/><w:dropDownList><w:listItem w:displayText="Red" w:value="r"/><w:listItem w:displayText="Blue" w:value="b"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>Red</w:t></w:r></w:sdtContent></w:sdt></w:p>
<w:p><w:sdt><w:sdtPr><w:id w:val="3"/><w:date w:fullDate="2024-05-01T00:00:00Z"><w:dateFormat w:val="dd.MM.yyyy"/><w:lid w:val="de-DE"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr><w:sdtContent><w:r><w:t>01.05.2024</w:t></w:r></w:sdtContent></w:sdt></w:p>
<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Name"/><w:id w:val="4"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:i/></w:rPr><w:t>Ann</w:t></w:r></w:sdtContent></w:sdt></w:p>
<w:sdt><w:sdtPr><w:alias w:val="Terms"/><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Block text</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;

test('content controls: check box, drop-down, date, plain text and block controls are kept and editable', async () => {
  const m = await readDocx(await docxOf(SDT_BODY));
  const cc = nodesOf(m.doc, 'contentControl');
  assert.deepEqual(cc.map((n) => n.attrs.kind), ['checkbox', 'dropdown', 'date']);
  assert.equal(cc[0].attrs.checked, false);
  assert.deepEqual(cc[1].attrs.items, [{ text: 'Red', value: 'r' }, { text: 'Blue', value: 'b' }]);
  assert.ok(cc[1].marks.some((x) => x.type === 'bold'), 'run formatting of the control kept');
  assert.equal(cc[2].attrs.fullDate, '2024-05-01T00:00:00Z');
  const name = [];
  walk(m.doc, (n) => { if (n.type === 'text' && (n.marks || []).some((x) => x.type === 'sdt')) name.push(n.text); });
  assert.deepEqual(name, ['Ann']);
  assert.equal(nodesOf(m.doc, 'controlBlock').length, 1);
  assert.ok(!(m.fidelity || []).includes('contentControls'));
  // edit: tick the box, pick Blue, change the date, rename
  cc[0].attrs.checked = true; cc[1].attrs.value = 'Blue'; cc[2].attrs.fullDate = '2025-12-24T00:00:00Z'; cc[2].attrs.value = '24.12.2025';
  walk(m.doc, (n) => { if (n.type === 'text' && n.text === 'Ann') n.text = 'Bob'; });
  const out = await writeDocx(m);
  const x = await docXml(out);
  assert.match(x, /<w14:checked w14:val="1"\/>/);
  assert.match(x, /<w:t xml:space="preserve">☒<\/w:t>/);
  assert.match(x, /w:fullDate="2025-12-24T00:00:00Z"/);
  assert.match(x, /<w:alias w:val="Name"\/><w:id w:val="4"\/><w:text\/><\/w:sdtPr><w:sdtContent><w:r><w:rPr><w:i\/><w:iCs\/><\/w:rPr><w:t xml:space="preserve">Bob<\/w:t>/);
  assert.match(x, /<w:sdt><w:sdtPr><w:alias w:val="Terms"\/><w:id w:val="5"\/><\/w:sdtPr><w:sdtContent><w:p>/);
  const again = await readDocx(out);
  const cc2 = nodesOf(again.doc, 'contentControl');
  assert.equal(cc2[0].attrs.checked, true);
  assert.equal(cc2[1].attrs.value, 'Blue');
  assert.equal(cc2[2].attrs.value, '24.12.2025');
  assert.ok(docText(again.doc).includes('Name: Bob'));
});

test('content controls made here get a complete sdtPr (Word opens them as controls)', async () => {
  const m = defaultModel();
  m.doc = { type: 'doc', content: [para(['Pick: ', { type: 'contentControl', attrs: { kind: 'dropdown', value: 'Yes', items: [{ text: 'Yes', value: 'y' }, { text: 'No', value: 'n' }] } }, ' ', { type: 'contentControl', attrs: { kind: 'checkbox', checked: true } }, ' ', { type: 'contentControl', attrs: { kind: 'date', value: '', fullDate: null, format: 'dd.MM.yyyy' } }])] };
  const x = await docXml(await writeDocx(m));
  assert.match(x, /<w:dropDownList><w:listItem w:displayText="Yes" w:value="y"\/><w:listItem w:displayText="No" w:value="n"\/><\/w:dropDownList>/);
  assert.match(x, /<w14:checkbox><w14:checked w14:val="1"\/>/);
  assert.match(x, /<w:showingPlcHdr\/><w:date><w:dateFormat w:val="dd.MM.yyyy"\/>/);
  const back = await readDocx(await writeDocx(m));
  assert.deepEqual(nodesOf(back.doc, 'contentControl').map((n) => n.attrs.kind), ['dropdown', 'checkbox', 'date']);
});

/* ---------------- fields ---------------- */
test('live fields: Word fields with cached values, Hostwerk values as locked controls; read back', async () => {
  const m = defaultModel();
  m.doc = { type: 'doc', content: [para([
    'Today ', { type: 'field', attrs: { kind: 'date', value: '20.11.2026' } },
    ', CPU ', { type: 'field', attrs: { kind: 'hostwerk', source: 'resources.cpu', params: { unit: '%' }, label: 'CPU load', value: '12 %' } },
    ', cell ', { type: 'field', attrs: { kind: 'cell', source: 'cell', params: { path: 'Budget.xlsx', sheet: 'Sum', cell: 'B4' }, label: 'Budget!B4', value: '1.250 €' } },
  ])] };
  const bytes = await writeDocx(m);
  const x = await docXml(bytes);
  assert.match(x, /<w:fldSimple w:instr=" DATE \\@ &quot;dd\.MM\.yyyy&quot; "><w:r><w:t xml:space="preserve">20\.11\.2026<\/w:t><\/w:r><\/w:fldSimple>/);
  assert.match(x, /<w:alias w:val="CPU load"\/><w:tag w:val="hostwerk-field:\{&quot;k&quot;:&quot;hostwerk&quot;/);
  assert.match(x, /<w:lock w:val="contentLocked"\/>/);
  const back = await readDocx(bytes);
  const fs = nodesOf(back.doc, 'field');
  assert.deepEqual(fs.map((f) => f.attrs.kind), ['date', 'hostwerk', 'cell']);
  assert.equal(fs[1].attrs.source, 'resources.cpu');
  assert.deepEqual(fs[2].attrs.params, { path: 'Budget.xlsx', sheet: 'Sum', cell: 'B4' });
  assert.equal(fs[2].attrs.value, '1.250 €');
  // unchanged → the original XML again; a refreshed value → a new result
  fs[0].attrs.value = '21.11.2026';
  const x2 = await docXml(await writeDocx(back));
  assert.match(x2, />21\.11\.2026</);
  assert.ok(x2.includes(fs[1].attrs.ox));
  // ODT: real ODF fields
  const odt = await JSZip.loadAsync(await writeOdt(m));
  assert.match(await odt.file('content.xml').async('string'), /<text:date>20\.11\.2026<\/text:date>/);
});

test('fields: PAGE / NUMPAGES / complex DATE fields become live fields', async () => {
  const body = '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE \\@ "d.M.yyyy" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1.2.2020</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:sectPr/>';
  const m = await readDocx(await docxOf(body));
  const fs = nodesOf(m.doc, 'field');
  assert.deepEqual(fs.map((f) => [f.attrs.kind, f.attrs.value]), [['page', '3'], ['date', '1.2.2020']]);
  assert.match(await docXml(await writeDocx(m)), /<w:instrText xml:space="preserve"> DATE \\@ "d\.M\.yyyy" <\/w:instrText>/);
});

/* ---------------- sections + headers ---------------- */
const HDR = `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W_NS}><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">ACME </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:hdr>`;
test('sections: own page setup per section, header with formatting, field and table, round trip', async () => {
  const body = `<w:p><w:r><w:t>Portrait part</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr><w:r><w:t>End of part 1</w:t></w:r></w:p><w:p><w:r><w:t>Landscape part</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:type w:val="nextPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;
  const src = await docxOf(body, { parts: { 'word/header1.xml': HDR }, rels: '<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' });
  const m = await readDocx(src);
  const sb = nodesOf(m.doc, 'sectionBreak');
  assert.equal(sb.length, 1);
  assert.equal(sb[0].attrs.page.orientation, 'portrait');
  assert.equal(m.page.orientation, 'landscape');
  assert.ok(Math.abs(sb[0].attrs.page.margins.top - 25) < 0.2);
  // the header is a document of its own: bold text, a PAGE field, a table
  const h = m.header;
  assert.equal(h.align, 'right');
  assert.ok(nodesOf(h.doc, 'text').some((n) => n.text === 'ACME ' && n.marks.some((x) => x.type === 'bold')));
  assert.equal(nodesOf(h.doc, 'field')[0].attrs.kind, 'page');
  assert.equal(nodesOf(h.doc, 'table').length, 1);
  assert.ok(!(m.fidelity || []).includes('headerLayout'));
  // edit the first section's orientation, save, read again
  sb[0].attrs.page = { ...sb[0].attrs.page, orientation: 'landscape' };
  const out = await writeDocx(m);
  const x = await docXml(out);
  const sects = x.match(/<w:sectPr>[\s\S]*?<\/w:sectPr>/g);
  assert.equal(sects.length, 2);
  assert.match(sects[0], /w:orient="landscape"/);
  assert.match(sects[0], /<w:headerReference w:type="default" r:id="rIdH1"\/>/, 'unchanged header part kept');
  assert.match(x, /<w:p><w:pPr><w:sectPr>[\s\S]*?<\/w:sectPr><\/w:pPr><w:r><w:t xml:space="preserve">End of part 1/);
  const again = await readDocx(out);
  assert.equal(nodesOf(again.doc, 'sectionBreak')[0].attrs.page.orientation, 'landscape');
});

test('headers/footers: rich content, picture with its own rels, first/even pages, evenAndOddHeaders', async () => {
  const m = defaultModel();
  m.doc = { type: 'doc', content: [para('Body')] };
  m.header = { text: 'Logo Report', align: 'left', doc: { type: 'doc', content: [para([{ type: 'image', attrs: { src: PNG_URL, width: 20 } }, ' ', b('Report')]), table([['a', 'b']], { header: false })] } };
  m.footer = { text: 'Page {page}', align: 'center', doc: { type: 'doc', content: [para(['Page ', { type: 'field', attrs: { kind: 'page', value: '1' } }, ' of ', { type: 'field', attrs: { kind: 'pages', value: '1' } }], { textAlign: 'center' })] } };
  m.headerFirst = { text: 'Title page', align: 'center', doc: { type: 'doc', content: [para('Title page')] } };
  m.footerEven = { text: 'even', align: 'left', doc: { type: 'doc', content: [para(txt('even', [{ type: 'italic' }]))] } };
  m.titlePage = true; m.evenOdd = true;
  const bytes = await writeDocx(m);
  const z = await JSZip.loadAsync(bytes);
  const x = await z.file('word/document.xml').async('string');
  assert.match(x, /<w:headerReference w:type="default" r:id="[^"]+"\/><w:headerReference w:type="first"/);
  assert.match(x, /<w:footerReference w:type="even"/);
  assert.match(x, /<w:titlePg\/>/);
  const hName = Object.keys(z.files).find((n) => /^word\/of_header\d\.xml$/.test(n) && /Report/.test(''));
  void hName;
  const parts = await Promise.all(Object.keys(z.files).filter((n) => /^word\/of_(header|footer)\d+\.xml$/.test(n)).map(async (n) => [n, await z.file(n).async('string')]));
  const hdr = parts.find(([, s]) => s.includes('Report'));
  assert.ok(hdr, 'rich header written');
  assert.match(hdr[1], /<w:b\/>/);
  assert.match(hdr[1], /<w:tbl>/);
  assert.match(hdr[1], /<a:blip r:embed="rIdH1"\/>/);
  assert.match(await z.file(hdr[0].replace('word/', 'word/_rels/') + '.rels').async('string'), /Id="rIdH1"[^>]*Target="media\/of_image\d+\.png"/);
  const ftr = parts.find(([, s]) => s.includes('NUMPAGES'));
  assert.match(ftr[1], /<w:fldSimple w:instr=" PAGE "/);
  assert.match(await z.file('word/settings.xml').async('string'), /<w:evenAndOddHeaders\/>/);
  const back = await readDocx(bytes);
  assert.ok(nodesOf(back.header.doc, 'image').length === 1);
  assert.equal(back.headerFirst.text, 'Title page');
  assert.ok(back.footerEven);
  assert.equal(back.evenOdd, true);
  assert.deepEqual(nodesOf(back.footer.doc, 'field').map((f) => f.attrs.kind), ['page', 'pages']);
  // html for the page boxes fills page numbers
  assert.match(docToHtml(back.footer.doc, { fields: { page: 7, pages: 9 } }), />7<\/span> of <span[^>]*>9</);
});

/* ---------------- tracked changes ---------------- */
test('track changes: paragraph marks, paragraph formatting and table rows are written and read back', async () => {
  const m = defaultModel();
  const rev = { id: '1', author: 'Ann', date: '2026-11-20T10:00:00Z' };
  m.doc = { type: 'doc', content: [
    para('Split here', { pmark: { type: 'ins', ...rev } }),
    para('joined', { pmark: { type: 'del', ...rev } }),
    para('centred now', { textAlign: 'center', pchange: { ...rev, old: '<w:pPr><w:jc w:val="left"/></w:pPr>' } }),
    { type: 'table', content: [{ type: 'tableRow', content: [{ type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: null }, content: [para('old row')] }] }, { type: 'tableRow', attrs: { change: { type: 'ins', ...rev } }, content: [{ type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: null }, content: [para('new row')] }] }] },
  ] };
  const bytes = await writeDocx(m);
  const x = await docXml(bytes);
  assert.match(x, /<w:pPr><w:rPr><w:ins w:id="\d+" w:author="Ann" w:date="2026-11-20T10:00:00Z"\/><\/w:rPr><\/w:pPr><w:r><w:t xml:space="preserve">Split here/);
  assert.match(x, /<w:rPr><w:del w:id="\d+" w:author="Ann"/);
  assert.match(x, /<w:jc w:val="center"\/><w:pPrChange w:id="\d+" w:author="Ann" w:date="2026-11-20T10:00:00Z"><w:pPr><w:jc w:val="left"\/><\/w:pPr><\/w:pPrChange>/);
  assert.match(x, /<w:trPr><w:ins w:id="\d+" w:author="Ann"/);
  const back = await readDocx(bytes);
  const ps = back.doc.content;
  assert.equal(ps[0].attrs.pmark.type, 'ins');
  assert.equal(ps[1].attrs.pmark.type, 'del');
  assert.equal(ps[2].attrs.pchange.author, 'Ann');
  assert.match(ps[2].attrs.pchange.old, /w:jc w:val="left"/);
  assert.equal(ps[3].content[1].attrs.change.type, 'ins');
  assert.ok(!JSON.stringify(ps[0].attrs.ox || {}).includes('w:ins'), 'the mark is modelled, not kept twice');
});

/* ---------------- comments ---------------- */
test('comments: commentsIds + commentsExtensible written, durable ids survive a round trip', async () => {
  const m = defaultModel();
  m.doc = { type: 'doc', content: [para(['A ', txt('remark', [{ type: 'comment', attrs: { id: 'c1' } }])])] };
  m.comments = [{ id: 'c1', author: 'Ann', date: '2026-11-20T10:00:00Z', text: 'Check', resolved: false, replies: [{ author: 'Bob', date: '2026-11-20T11:00:00Z', text: 'Done' }] }];
  const bytes = await writeDocx(m);
  const z = await JSZip.loadAsync(bytes);
  const ids = await z.file('word/commentsIds.xml').async('string');
  const cex = await z.file('word/commentsExtensible.xml').async('string');
  const durable = [...ids.matchAll(/w16cid:durableId="([0-9A-F]{8})"/g)].map((r) => r[1]);
  assert.equal(durable.length, 2);
  assert.ok(durable.every((d) => parseInt(d, 16) < 0x7fffffff));
  assert.match(cex, /w16cex:dateUtc="2026-11-20T11:00:00Z"/);
  const rels = await z.file('word/_rels/document.xml.rels').async('string');
  assert.match(rels, /relationships\/commentsIds" Target="commentsIds\.xml"/);
  assert.match(rels, /relationships\/commentsExtensible" Target="commentsExtensible\.xml"/);
  assert.match(await z.file('[Content_Types].xml').async('string'), /commentsIds\+xml/);
  const back = await readDocx(bytes);
  assert.equal(back.comments[0].durableId, durable[0]);
  const z2 = await JSZip.loadAsync(await writeDocx(back));
  assert.ok((await z2.file('word/commentsIds.xml').async('string')).includes(durable[0]));
});

/* ---------------- diagrams ---------------- */
test('diagrams: written as a picture with the source in its alt text (docx + odt), read back as diagrams', async () => {
  const m = defaultModel();
  const source = 'graph TD\n  A --> B';
  m.doc = { type: 'doc', content: [para([{ type: 'diagram', attrs: { kind: 'mermaid', source, src: PNG_URL, width: 120, height: 80 } }])] };
  const bytes = await writeDocx(m);
  const x = await docXml(bytes);
  assert.match(x, /<wp:docPr id="\d+" name="Picture \d+" descr="graph TD&#10;  A --&gt; B" title="hostwerk-diagram:mermaid"\/>/);
  const d = nodesOf((await readDocx(bytes)).doc, 'diagram')[0];
  assert.equal(d.attrs.source, source);
  assert.match(d.attrs.src, /^data:image\/png/);
  const od = nodesOf((await readOdt(await writeOdt(m))).doc, 'diagram')[0];
  assert.equal(od.attrs.source, source);
});

/* ---------------- text boxes ---------------- */
test('text boxes: the text of a kept text box can be edited, the shape stays', async () => {
  const box = '<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="900000" cy="400000"/><wp:docPr id="9" name="Text Box 9"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Old</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
  const m = await readDocx(await docxOf(`<w:p>${box}</w:p><w:sectPr/>`));
  const p = nodesOf(m.doc, 'preserved')[0];
  assert.equal(p.attrs.kind, 'textbox');
  assert.equal(p.attrs.text, 'Old');
  p.attrs.text = 'New\nsecond line'; p.attrs.textEdited = true;
  const x = await docXml(await writeDocx(m));
  assert.match(x, /<w:txbxContent><w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">New<\/w:t><\/w:r><\/w:p><w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">second line/);
  assert.match(x, /<wps:wsp/);
});

/* ---------------- encrypted files ---------------- */
function buildCfb(streams) {
  const SEC = 512; const END = 0xfffffffe; const FREE = 0xffffffff;
  const names = Object.keys(streams);
  const secCount = names.map((n) => Math.ceil(streams[n].length / SEC));
  const total = 2 + secCount.reduce((a, c) => a + c, 0);
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
test('encrypted Word files (docx in an encrypted package, protected .doc) give ENCRYPTED for the vault', async () => {
  const enc = buildCfb({ EncryptionInfo: new Uint8Array(4096), EncryptedPackage: new Uint8Array(4608) });
  await assert.rejects(() => readDocument(enc, 'docx'), (e) => e.code === 'ENCRYPTED');
});

/* ---------------- compare ---------------- */
test('compare: word-level diff of two versions, applied as real tracked changes', async () => {
  const a = { type: 'doc', content: [heading(1, 'Plan'), para('The quick brown fox jumps.'), para('Unchanged line.'), para('Removed paragraph.')] };
  const bDoc = { type: 'doc', content: [heading(1, 'Plan'), para('The quick red fox jumps high.'), para('Unchanged line.'), para('A new paragraph.')] };
  const diff = compareDocs(a, bDoc);
  assert.equal(diff.stats.unchanged, 2);
  assert.ok(diff.stats.changed >= 1);
  const words = diff.blocks.find((x) => x.kind === 'changed').parts;
  assert.ok(words.some((p) => p.op === -1 && p.text.includes('brown')));
  assert.ok(words.some((p) => p.op === 1 && p.text.includes('red')));
  const tracked = applyAsTrackedChanges(a, bDoc, { author: 'Ann', date: '2026-11-20T10:00:00Z' });
  const m = defaultModel(); m.doc = tracked;
  const x = await docXml(await writeDocx(m));
  assert.match(x, /<w:del w:id="\d+" w:author="Ann"[^>]*><w:r><w:delText xml:space="preserve">brown/);
  assert.match(x, /<w:ins w:id="\d+" w:author="Ann"[^>]*><w:r><w:t xml:space="preserve">red/);
  assert.match(x, /<w:delText xml:space="preserve">Removed<\/w:delText>/);
  // accepting everything gives the new version's text
  const accepted = JSON.parse(JSON.stringify(tracked));
  walk(accepted, (n) => { if (n.content) n.content = n.content.filter((c) => !(c.marks || []).some((mk) => mk.type === 'deletion')); });
  const txtOf = (d) => docText(d).replace(/\s+/g, ' ').trim();
  assert.equal(txtOf({ ...accepted, content: accepted.content.filter((p) => !(p.attrs && p.attrs.pmark && p.attrs.pmark.type === 'del' && !(p.content || []).length)) }), txtOf(bDoc));
});

/* ---------------- track changes helpers (pure) ---------------- */
test('track changes: paragraph formatting before/after survives pPr XML (reject restores it)', async () => {
  const { pPrXml, attrsOfPPr } = await import('../ui/src/editors/docs/extensions/track.js');
  const a = { textAlign: 'center', lineHeight: 1.5, spaceBefore: 6, spaceAfter: 12, indent: 36, firstLine: -18 };
  assert.deepEqual(attrsOfPPr(pPrXml(a)), a);
  assert.deepEqual(attrsOfPPr('<w:pPr/>'), { textAlign: null, lineHeight: null, spaceBefore: null, spaceAfter: null, indent: null, firstLine: null });
});

test('every writable format accepts the 1.1 structures (no crash, text kept)', async () => {
  const { writeDocument } = await import('../ui/src/editors/docs/io/index.js');
  const m = defaultModel();
  m.doc = { type: 'doc', content: [
    para(['A ', { type: 'field', attrs: { kind: 'date', value: '20.11.2026' } }, ' B ', { type: 'contentControl', attrs: { kind: 'dropdown', value: 'Blue', items: [{ text: 'Blue', value: 'b' }] } }, ' C ', { type: 'contentControl', attrs: { kind: 'checkbox', checked: true } }, ' ', { type: 'equation', attrs: { latex: 'x^2', display: false } }]),
    { type: 'controlBlock', attrs: { kind: 'rich' }, content: [para('inside control')] },
    { type: 'sectionBreak', attrs: { type: 'nextPage', page: { ...m.page } } },
    para([{ type: 'diagram', attrs: { kind: 'mermaid', source: 'graph TD\nA-->B', src: PNG_URL, width: 40, height: 40 } }, ' after']),
  ] };
  for (const ext of ['docx', 'docm', 'dotx', 'odt', 'ott', 'fodt', 'rtf', 'html', 'xhtml', 'md', 'txt']) {
    const out = await writeDocument(JSON.parse(JSON.stringify(m)), ext);
    assert.ok(out.length > 50, ext);
    if (['rtf', 'html', 'md', 'txt', 'fodt'].includes(ext)) { const s = new TextDecoder().decode(out); assert.ok(/inside control/.test(s) && /20\.11\.2026/.test(s), `${ext} keeps the text`); }
  }
});
