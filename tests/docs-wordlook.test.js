// Docs 1.3.2: a real Word compliance matrix (alternating portrait/landscape sections, 35 tables with repeated header
// rows, coloured heading styles, a logo, ■ bullets, "left ⇥ right" headers) looked and paginated differently from
// Word. Synthetic fixtures with the same structures – never the real document.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { getSchema } from '@tiptap/core';
import Document from '@tiptap/extension-document';
import Paragraph from '@tiptap/extension-paragraph';
import Text from '@tiptap/extension-text';
import Bold from '@tiptap/extension-bold';
import { EditorState } from '@tiptap/pm/state';
import { readDocx } from '../ui/src/editors/docs/io/docx-read.js';
import { writeDocx } from '../ui/src/editors/docs/io/docx-write.js';
import { walk, styleLookVars, hfDist } from '../ui/src/editors/docs/io/pm.js';
import { cellCss, tableLook, sidePad, borderPx } from '../ui/src/editors/docs/io/tables.js';
import { paraCss, paraBoxCss } from '../ui/src/editors/docs/io/wordprops.js';
import { docToHtml, tabbedParagraph } from '../ui/src/editors/docs/io/render-html.js';
import { pager, pageAfter, firstChange, keepsNext, pageBreakCounts } from '../ui/src/editors/docs/extensions/pagemodel.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const run = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const sect = (w, h, orient, refs = '') => `<w:sectPr>${refs}<w:pgSz w:w="${w}" w:h="${h}" w:orient="${orient}"/><w:pgMar w:top="1300" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>`;
const cell = (w, t, fill) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${w}"/><w:tcBorders><w:top w:val="single" w:color="BFBFBF" w:sz="4"/><w:left w:val="single" w:color="BFBFBF" w:sz="4"/><w:bottom w:val="single" w:color="BFBFBF" w:sz="4"/><w:right w:val="single" w:color="BFBFBF" w:sz="4"/></w:tcBorders>${fill ? `<w:shd w:fill="${fill}" w:color="auto" w:val="clear"/>` : ''}<w:tcMar><w:top w:type="dxa" w:w="60"/><w:left w:type="dxa" w:w="90"/><w:bottom w:type="dxa" w:w="60"/><w:right w:type="dxa" w:w="90"/></w:tcMar><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:spacing w:after="30"/></w:pPr>${run(t, '<w:sz w:val="16"/>')}</w:p></w:tc>`;
const row = (cells, head) => `<w:tr><w:trPr><w:cantSplit/>${head ? '<w:tblHeader/>' : ''}</w:trPr>${cells}</w:tr>`;
const matrix = (n) => `<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="14570"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="12570"/></w:tblGrid>${row(cell(2000, 'Zeile', 'C8000A') + cell(12570, 'Inhalt', 'C8000A'), true)}${Array.from({ length: n }, (_, k) => row(cell(2000, `R-${k + 1}`) + cell(12570, `Text der Zeile ${k + 1}`, k % 2 ? 'F2F2F2' : null))).join('')}</w:tbl>`;
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="333333"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="276"/></w:pPr></w:pPrDefault></w:docDefaults>`
  + '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="200" w:before="0"/></w:pPr><w:rPr><w:b/><w:color w:val="C8000A"/><w:sz w:val="52"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:color="FF0000" w:sz="18" w:space="6"/></w:pBdr><w:spacing w:after="240" w:before="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="C8000A"/><w:sz w:val="34"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:pBdr><w:left w:val="single" w:color="FF0000" w:sz="30" w:space="8"/></w:pBdr><w:spacing w:after="120" w:before="280"/><w:ind w:left="170"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="C8000A"/><w:sz w:val="26"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style></w:styles>';
const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="■"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="260"/></w:pPr><w:rPr><w:color w:val="C8000A"/><w:sz w:val="14"/></w:rPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="☐"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>`;
const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}><w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:color="FF0000" w:sz="8" w:space="4"/></w:pBdr><w:tabs><w:tab w:val="right" w:pos="14570"/></w:tabs></w:pPr>${run('Bericht · Version 1')}<w:r><w:tab/></w:r>${run('Klassifizierung: intern', '<w:b/><w:color w:val="C8000A"/>')}</w:p></w:hdr>`;

async function fixture() {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>');
  z.file('word/styles.xml', STYLES);
  z.file('word/numbering.xml', NUMBERING);
  z.file('word/header1.xml', HEADER);
  const body = [
    `<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>${run('Abgleich')}</w:p>`,
    `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${run('Punkt eins')}</w:p>`,
    `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>${run('Zu erledigen')}</w:p>`,
    `<w:p><w:pPr><w:keepNext w:val="false"/><w:pBdr><w:left w:val="single" w:color="FF0000" w:sz="36" w:space="10"/></w:pBdr><w:shd w:fill="F2F2F2" w:color="auto" w:val="clear"/><w:spacing w:after="160"/><w:ind w:left="240" w:right="120"/></w:pPr>${run('Hinweis im Kasten')}</w:p>`,
    // the section ends in a paragraph that holds a page break and the section break (as the real document does)
    `<w:p><w:pPr>${sect(11906, 16838, 'portrait', '<w:headerReference w:type="default" r:id="rId3"/>')}</w:pPr><w:r><w:br w:type="page"/></w:r></w:p>`,
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run('5 Abgleich je Anforderung')}</w:p>`,
    `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${run('5.1 Kapitel 1')}</w:p>`,
    `<w:p><w:pPr><w:keepNext/><w:spacing w:after="120"/></w:pPr>${run('23 Zeilen')}</w:p>`,
    matrix(30),
  ].join('');
  z.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${sect(16838, 11906, 'landscape', '<w:headerReference w:type="default" r:id="rId3"/>')}</w:body></w:document>`);
  return z.generateAsync({ type: 'uint8array' });
}

test('docx: the document\'s own heading / title styles, text colour and header distance are read for the page look', async () => {
  const m = await readDocx(await fixture());
  assert.equal(m.defaults.color, '#333333');
  const L = m.styleLook;
  assert.deepEqual({ size: L.h1.size, color: L.h1.color, b: L.h1.b, before: L.h1.before, after: L.h1.after }, { size: 17, color: '#C8000A', b: true, before: 6, after: 12 });
  assert.match(L.h1.pborders.bottom, /^single:18:FF0000:6/);
  assert.equal(L.h2.indent, 8.5);
  assert.match(L.h2.pborders.left, /^single:30:FF0000:8/);
  assert.deepEqual({ size: L.title.size, color: L.title.color }, { size: 26, color: '#C8000A' });
  // header / footer distance from the paper edge (pgMar w:header / w:footer = 567 twips = 10 mm)
  assert.equal(m.page.headerMm, 10);
  assert.equal(m.page.footerMm, 10);
  const px = hfDist(null, m, 'headerMm', 86.6);
  assert.ok(Math.abs(px - 37.8) < 0.2, `header at ${px}px`);
  assert.equal(hfDist(null, { page: {} }, 'footerMm', 75), 30); // no distance stored → a share of the margin as before
});

test('styleLookVars: heading look as CSS variables – sizes, colours, spacing, Word line height, borders outside the indent', () => {
  const v = styleLookVars({ h1: { size: 17, color: '#C8000A', b: true, before: 6, after: 12, pborders: { bottom: 'single:18:FF0000:6' } }, h2: { size: 13, indent: 8.5, pborders: { left: 'single:30:FF0000:8' } }, title: { size: 26 } }, { font: 'Arial', lineHeight: 1.15 }, (f) => f);
  assert.equal(v['--of-h1-size'], '17pt');
  assert.equal(v['--of-h1-weight'], '700');
  assert.equal(v['--of-h1-mt'], '6pt');
  assert.equal(v['--of-h1-mb'], '12pt');
  assert.equal(v['--of-h1-lh'], String(Math.round(1.15 * 1.149 * 1000) / 1000)); // Word "multiple 1.15" of Arial
  assert.equal(v['--of-h1-bb'], '3px solid #FF0000');
  assert.equal(v['--of-h1-pb'], '6pt');
  assert.equal(v['--of-h2-bl'], '5px solid #FF0000');
  assert.equal(v['--of-h2-pl'], '8pt');
  assert.equal(v['--of-h2-mlx'], 'calc(0.5pt - 5px)'); // indent 8.5 pt − space 8 pt − the border
  assert.equal(v['--of-title-ls'], 'normal');
  assert.deepEqual(styleLookVars(null), {});
});

test('lists: ■ bullets are square bullets (not check boxes), ☐ stays a check list', async () => {
  const m = await readDocx(await fixture());
  const lists = [];
  walk(m.doc, (n) => { if (['bulletList', 'taskList', 'orderedList'].includes(n.type)) lists.push({ type: n.type, style: n.attrs && n.attrs.listStyle }); });
  assert.deepEqual(lists, [{ type: 'bulletList', style: 'square' }, { type: 'taskList', style: undefined }]);
});

test('paragraph borders like Word: space as padding, left/right border outside the indent, no fixed 1pt/4pt padding', () => {
  const css = paraCss({ pborders: { left: 'single:36:FF0000:10' }, pshading: '#F2F2F2', indent: 12, rightIndent: 6 });
  assert.match(css, /border-left:6px solid #FF0000/);
  assert.match(css, /padding-left:10pt/);
  assert.match(css, /margin-left:calc\(2pt - 6px\) !important/);
  assert.doesNotMatch(css, /padding:1pt 4pt/);
  assert.deepEqual(paraBoxCss({ top: 'nil', bottom: 'single:8:FF0000:4' }), { 'border-bottom': '1.3px solid #FF0000', 'padding-bottom': '4pt' });
});

test('table cells: Word cell margins with ½ pt borders – the extra width of a 1px CSS border is taken from the padding', () => {
  assert.equal(borderPx('single:4:BFBFBF:0'), 1);
  // 90 twips = 6 px; Word draws 0.667 px of border, CSS 1 px (half of it inside the cell) → 6 − (1 − 0.667) / 2
  assert.equal(sidePad(90, 'single:4:BFBFBF:0'), '5.833px');
  assert.equal(sidePad(90, 'nil'), '6px');
  assert.equal(sidePad(90, null), '6px');
  const css = cellCss({ borders: { top: 'single:4:BFBFBF:0', left: 'single:4:BFBFBF:0', bottom: 'single:4:BFBFBF:0', right: 'single:4:BFBFBF:0' }, margins: { top: 60, left: 90, bottom: 60, right: 90 } });
  assert.match(css, /padding-top:3\.833px/);
  assert.match(css, /padding-left:5\.833px/);
  const look = tableLook({ borders: { insideV: 'single:4:000000:0', insideH: 'single:4:000000:0' }, cellMargins: { left: 108, right: 108, top: 0, bottom: 0 } });
  assert.equal(look.vars['--tc-p'], '0px 7.033px 0px 7.033px');
});

test('header with "left ⇥ right" and a bottom border: rendered spread over the line (the right part was glued to the left one)', async () => {
  const m = await readDocx(await fixture());
  const h = m.header && m.header.doc ? m.header : Object.values(m).find((x) => x && x.doc && JSON.stringify(x.doc).includes('Klassifizierung'));
  assert.ok(h && h.doc, 'header read');
  const html = docToHtml(h.doc);
  assert.match(html, /display:flex;justify-content:space-between/);
  assert.match(html, /<span>[^]*Bericht[^]*<\/span><span>[^]*Klassifizierung/);
  assert.match(html, /border-bottom:1\.3px solid #FF0000;padding-bottom:4pt/);
  assert.equal(tabbedParagraph({ tabs: [{ pos: 200, align: 'left' }] }, [{ type: 'text', text: 'a\tb' }]), null); // left stops: the editor's own tab layout
});

test('saving: the re-read look survives, nothing of the matrix gets lost (bullets stay bullets, headers repeat, cells keep fills)', async () => {
  const orig = await fixture();
  const m = await readDocx(orig);
  const out = await writeDocx(m);
  const z = await JSZip.loadAsync(out);
  const doc = await z.file('word/document.xml').async('string');
  assert.equal((doc.match(/<w:tblHeader\/>/g) || []).length, 1);
  assert.equal((doc.match(/<w:cantSplit\/>/g) || []).length, 31);
  assert.equal((doc.match(/w:fill="F2F2F2"/g) || []).length >= 15, true);
  assert.match(doc, /<w:numId w:val="1"\/>/); // the ■ list keeps its numbering
  assert.equal(await z.file('word/styles.xml').async('string'), STYLES); // the styles part is untouched
  assert.equal(await z.file('word/header1.xml').async('string'), HEADER);
  const again = await readDocx(out);
  assert.deepEqual(again.styleLook, m.styleLook);
});

test('pagination: the space after a paragraph does not push the next block onto the next page uncounted', () => {
  const MM = 96 / 25.4;
  const A4 = { pageW: 210 * MM, pageH: 297 * MM, mt: 25 * MM, mb: 20 * MM, ml: 25 * MM, mr: 25 * MM, contentW: 160 * MM, contentH: 252 * MM };
  const P = pager([A4], 24); P.add(0);
  const E = P.E(0);
  // a paragraph whose text ends 3 px above the page end and whose space after runs past it stays on page 0 –
  // counting the space made the next block "already on page 1" without a spacer (drawn over the footer)
  assert.equal(pageAfter(P, 0, E - 3), 0);
  assert.equal(P.pages.length, 1);
  assert.equal(pageAfter(P, 0, E + 40), 1); // text that really runs over (a picture taller than the rest)
});

test('pagination: formatting changes (marks) mark the layout dirty from their position', () => {
  const schema = getSchema([Document, Paragraph, Text, Bold]);
  const doc = schema.node('doc', null, Array.from({ length: 5 }, (_, k) => schema.node('paragraph', null, schema.text(`Absatz ${k}`))));
  const state = EditorState.create({ schema, doc });
  const tr = state.tr.addMark(32, 35, schema.marks.bold.create());
  assert.equal(tr.mapping.maps.every((m) => { let n = 0; m.forEach(() => { n++; }); return n === 0; }), true); // empty step maps
  assert.equal(firstChange(tr, Infinity), 32);
  const typed = state.tr.insertText('x', 12);
  assert.equal(firstChange(typed, Infinity), 12);
  assert.equal(firstChange(state.tr, Infinity), Infinity);
});

test('pagination: a page break right before a next-page section break makes no empty page; keep-with-next rules', () => {
  const n = (name, attrs = {}) => ({ type: { name }, attrs });
  assert.equal(pageBreakCounts(n('sectionBreak', { type: 'nextPage' })), false);
  assert.equal(pageBreakCounts(n('sectionBreak', { type: 'continuous' })), true);
  assert.equal(pageBreakCounts(n('paragraph')), true);
  assert.equal(pageBreakCounts(undefined), true);
  assert.equal(keepsNext(n('heading', { level: 2 })), true); // Word's heading styles keep with the next paragraph
  assert.equal(keepsNext(n('heading', { level: 2, keepNext: false })), false);
  assert.equal(keepsNext(n('paragraph', { keepNext: true })), true);
  assert.equal(keepsNext(n('paragraph', {})), false);
});
