import test from 'node:test';
import assert from 'node:assert/strict';
import { parseGoogleFile, googleId, googleUrl, googleExport } from '../ui/src/editors/viewer/google.js';
import { geometryToPath, readVsdx, color, shapeText } from '../ui/src/editors/viewer/vsdx.js';
import { snappyUncompress, iwaUnframe, protoStrings, binaryStrings } from '../ui/src/editors/viewer/extract.js';
import { parseXml } from '../ui/src/editors/_shared/xmltree.js';

test('Google files → id and URL', () => {
  const j = parseGoogleFile(new TextEncoder().encode('{"url": "https://docs.google.com/open?id=1AbCdEfGhIjKlMn", "doc_id": "1AbCdEfGhIjKlMn", "email": "lena@gmail.com"}'));
  assert.equal(googleId(j), '1AbCdEfGhIjKlMn');
  assert.equal(googleUrl('gdoc', j), 'https://docs.google.com/open?id=1AbCdEfGhIjKlMn');
  assert.equal(googleUrl('gsheet', { doc_id: 'XYZ1234567890' }), 'https://docs.google.com/spreadsheets/d/XYZ1234567890/edit');
  assert.equal(googleUrl('gslides', { resource_id: 'presentation:PRES1234567890' }), 'https://docs.google.com/presentation/d/PRES1234567890/edit');
  assert.equal(googleUrl('gmap', { doc_id: 'MAP1234567890' }), 'https://www.google.com/maps/d/edit?mid=MAP1234567890');
  assert.equal(googleUrl('glink', { doc_id: 'F1234567890' }), 'https://drive.google.com/open?id=F1234567890');
  assert.equal(googleUrl('gdoc', { url: 'https://evil.example.com/x' }), ''); // foreign hosts are never used
  assert.equal(googleId({ url: 'https://docs.google.com/document/d/DOC_abcdefghij/edit' }), 'DOC_abcdefghij');
  assert.deepEqual(googleExport('gsheet'), { as: 'xlsx', kind: 'sheets' });
  assert.equal(googleExport('gform'), null);
});

test('Visio geometry → SVG path', () => {
  const rect = [{ T: 'MoveTo', cells: { X: '0', Y: '0' } }, { T: 'LineTo', cells: { X: '2', Y: '0' } }, { T: 'LineTo', cells: { X: '2', Y: '1' } }, { T: 'LineTo', cells: { X: '0', Y: '1' } }, { T: 'LineTo', cells: { X: '0', Y: '0' } }];
  assert.deepEqual(geometryToPath(rect, 2, 1), { d: 'M0 0L2 0L2 1L0 1L0 0Z', closed: true });
  const rel = [{ T: 'RelMoveTo', cells: { X: '0', Y: '0' } }, { T: 'RelLineTo', cells: { X: '1', Y: '1' } }];
  assert.equal(geometryToPath(rel, 4, 2).d, 'M0 0L4 2');
  const ell = geometryToPath([{ T: 'Ellipse', cells: { X: '1', Y: '1', A: '2', B: '1', C: '1', D: '2' } }], 2, 2);
  assert.ok(ell.closed && ell.d.startsWith('M2 1') && ell.d.endsWith('Z'));
  // a half circle via ArcTo: all points stay within radius 1 of the centre
  const arc = geometryToPath([{ T: 'MoveTo', cells: { X: '0', Y: '0' } }, { T: 'ArcTo', cells: { X: '2', Y: '0', A: '1' } }], 2, 1).d;
  const pts = arc.slice(1).split(/[ML]/).map((p) => p.trim().split(' ').map(Number));
  assert.ok(pts.length > 8 && pts.every(([x, y]) => Math.abs(Math.hypot(x - 1, y) - 1) < 0.02));
  const bez = geometryToPath([{ T: 'RelMoveTo', cells: { X: '0', Y: '0' } }, { T: 'RelCubBezTo', cells: { X: '1', Y: '0', A: '0', B: '1', C: '1', D: '1' } }], 10, 10).d;
  assert.equal(bez, 'M0 0C0 10 10 10 10 0');
  const tf = (x, y) => [x * 72, (1 - y) * 72];
  assert.equal(geometryToPath(rect.slice(0, 2), 2, 1, tf).d, 'M0 72L144 72');
  assert.equal(color('#FF0000'), '#ff0000');
  assert.equal(color('2'), '#ff0000');
  assert.equal(color('THEMEVAL()', '#123456'), '#123456');
  assert.equal(shapeText(parseXml('<Text><cp IX="0"/>Hallo <fld IX="0">Feld</fld>\nWelt</Text>')), 'Hallo Feld\nWelt');
});

test('Visio file with a master, a style and a page', () => {
  const files = {
    'visio/document.xml': '<VisioDocument><StyleSheets><StyleSheet ID="0"><Cell N="LineWeight" V="0.01"/><Cell N="LineColor" V="#000000"/><Cell N="FillForegnd" V="#ffffff"/></StyleSheet></StyleSheets></VisioDocument>',
    'visio/masters/masters.xml': '<Masters xmlns:r="r"><Master ID="2" NameU="Box"><Rel r:id="rId1"/></Master></Masters>',
    'visio/masters/_rels/masters.xml.rels': '<Relationships><Relationship Id="rId1" Target="master1.xml"/></Relationships>',
    'visio/masters/master1.xml': '<MasterContents><Shapes><Shape ID="5" LineStyle="0" FillStyle="0"><Cell N="Width" V="1"/><Cell N="Height" V="1"/><Cell N="FillForegnd" V="#00ff00"/><Section N="Geometry" IX="0"><Row T="RelMoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row><Row T="RelLineTo" IX="2"><Cell N="X" V="1"/><Cell N="Y" V="0"/></Row><Row T="RelLineTo" IX="3"><Cell N="X" V="1"/><Cell N="Y" V="1"/></Row><Row T="RelLineTo" IX="4"><Cell N="X" V="0"/><Cell N="Y" V="1"/></Row><Row T="RelLineTo" IX="5"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row></Section><Text>Box</Text></Shape></Shapes></MasterContents>',
    'visio/pages/pages.xml': '<Pages xmlns:r="r"><Page ID="0" NameU="Seite-1"><PageSheet><Cell N="PageWidth" V="4"/><Cell N="PageHeight" V="3"/></PageSheet><Rel r:id="rId1"/></Page></Pages>',
    'visio/pages/_rels/pages.xml.rels': '<Relationships><Relationship Id="rId1" Target="page1.xml"/></Relationships>',
    'visio/pages/page1.xml': '<PageContents><Shapes><Shape ID="1" Master="2"><Cell N="PinX" V="2"/><Cell N="PinY" V="1.5"/><Cell N="Width" V="2"/><Cell N="Height" V="1"/></Shape><Shape ID="2"><Cell N="PinX" V="1"/><Cell N="PinY" V="1"/><Cell N="Width" V="1"/><Cell N="Height" V="0"/><Cell N="EndArrow" V="4"/><Section N="Geometry" IX="0"><Cell N="NoFill" V="1"/><Row T="MoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row><Row T="LineTo" IX="2"><Cell N="X" V="1"/><Cell N="Y" V="0"/></Row></Section></Shape></Shapes></PageContents>',
  };
  const { pages } = readVsdx(files);
  assert.equal(pages.length, 1);
  assert.equal(pages[0].name, 'Seite-1');
  // master geometry scaled to the instance (2 × 1 inch at pin 2/1.5 → x 72..216, y 72..144 in a 216 high page)
  assert.match(pages[0].svg, /<path d="M72 144L216 144L216 72L72 72L72 144Z" fill="#00ff00"/);
  assert.match(pages[0].svg, />Box</);
  assert.match(pages[0].svg, /marker-end="url\(#vsa\)"/);
  assert.deepEqual(pages[0].texts, ['Box']);
});

test('Snappy, IWA framing and protobuf strings', () => {
  // snappy block: length 11, literal "hello " (6), copy len 5 offset 6 → "hello hello"
  const block = Uint8Array.from([11, (6 - 1) << 2, ...new TextEncoder().encode('hello '), 0x01 | ((5 - 4) << 2), 6]);
  assert.equal(new TextDecoder().decode(snappyUncompress(block)), 'hello hello');
  const framed = Uint8Array.from([0, block.length, 0, 0, ...block]);
  assert.equal(new TextDecoder().decode(iwaUnframe(framed)), 'hello hello');
  const msg = Uint8Array.from([0x08, 0x01, 0x12, 5, ...new TextEncoder().encode('Hallo'), 0x1a, 3, 1, 2, 3]);
  assert.deepEqual(protoStrings(msg), ['Hallo']);
  const bin = Uint8Array.from([7, 0, 0, 0, ...[...'Quarterly plan'].flatMap((c) => [c.charCodeAt(0), 0]), 0, 0, 7]);
  assert.ok(binaryStrings(bin).strings.includes('Quarterly plan'));
});
