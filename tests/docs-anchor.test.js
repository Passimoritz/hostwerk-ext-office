/* Docs 1.2: floating Word text boxes / shapes (wp:anchor) – position read, shown, moved and written back */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { anchorOf, moveAnchor, floatStyle } from '../ui/src/editors/docs/io/anchor.js';

const run = (pos, wrap = '<wp:wrapSquare wrapText="bothSides"/>') => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${pos}<wp:extent cx="1905000" cy="952500"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}<wp:docPr id="1" name="Textfeld 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>Kasten</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
const OFF = '<wp:positionH relativeFrom="column"><wp:posOffset>952500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV>';

test('anchorOf: offsets, size, wrap, alignment', () => {
  assert.deepEqual(anchorOf(run(OFF)), { x: 100, y: 20, w: 200, h: 100, relH: 'column', relV: 'paragraph', alignH: null, alignV: null, wrap: 'square', behind: false });
  const al = anchorOf(run('<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>', '<wp:wrapNone/>'));
  assert.equal(al.alignH, 'right'); assert.equal(al.x, null); assert.equal(al.wrap, 'none'); assert.equal(al.relV, 'page');
  assert.equal(anchorOf('<w:r><w:drawing><wp:inline/></w:drawing></w:r>'), null);
});

test('floatStyle: square wraps left/right, top-and-bottom is a block, none is absolute; page offsets lose the margin', () => {
  assert.deepEqual(floatStyle(anchorOf(run(OFF)), { contentW: 600, marginLeft: 94 }), { mode: 'float-left', marginLeft: 100, top: 20 });
  const right = anchorOf(run('<wp:positionH relativeFrom="page"><wp:posOffset>3810000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'));
  assert.deepEqual(floatStyle(right, { contentW: 600, marginLeft: 94 }), { mode: 'float-right', marginRight: 94, top: 0 }); // 400 px from the page edge = 306 in the column
  assert.equal(floatStyle(anchorOf(run(OFF, '<wp:wrapTopAndBottom/>'))).mode, 'block');
  const none = floatStyle(anchorOf(run(OFF, '<wp:wrapNone/>')), { contentW: 600 });
  assert.deepEqual(none, { mode: 'absolute', left: 100, top: 20 });
  // vertical positions from the page keep the box at its paragraph
  assert.equal(floatStyle({ ...anchorOf(run(OFF)), relV: 'page', y: 500 }).top, 0);
});

test('moveAnchor: new offsets from the column / paragraph in EMU, everything else byte-identical', () => {
  const x0 = run('<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>');
  const moved = moveAnchor(x0, 150, 30);
  assert.match(moved, /<wp:positionH relativeFrom="column"><wp:posOffset>1428750<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>285750<\/wp:posOffset><\/wp:positionV>/);
  const strip = (s) => s.replace(/<wp:position[HV]\b[\s\S]*?<\/wp:position[HV]>/g, '');
  assert.equal(strip(moved), strip(x0));
  assert.deepEqual(anchorOf(moved), { ...anchorOf(x0), x: 150, y: 30, relH: 'column', relV: 'paragraph', alignH: null, alignV: null });
});

test('docx: an anchored text box is read as a floating kept node; moved, it is written back with the new offsets', async () => {
  const { default: JSZip } = await import('jszip');
  const { readDocx } = await import('../ui/src/editors/docs/io/docx-read.js');
  const { writeDocx } = await import('../ui/src/editors/docs/io/docx-write.js');
  const { walk } = await import('../ui/src/editors/docs/io/pm.js');
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
  const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body><w:p>${run(OFF)}<w:r><w:t>Text daneben</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const m = await readDocx(await z.generateAsync({ type: 'uint8array' }));
  let kept = null; walk(m.doc, (n) => { if (n.type === 'preserved') kept = n; });
  assert.ok(kept); assert.equal(kept.attrs.kind, 'textbox');
  assert.deepEqual({ x: kept.attrs.float.x, y: kept.attrs.float.y, wrap: kept.attrs.float.wrap }, { x: 100, y: 20, wrap: 'square' });
  kept.attrs.xml = moveAnchor(kept.attrs.xml, 40, 8); // what the node view does after a drag
  const out = await (await JSZip.loadAsync(await writeDocx(m))).file('word/document.xml').async('string');
  assert.match(out, /<wp:positionH relativeFrom="column"><wp:posOffset>381000<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>76200<\/wp:posOffset><\/wp:positionV>/);
  assert.match(out, /<w:t>Kasten<\/w:t>/); assert.match(out, /Text daneben/);
});
