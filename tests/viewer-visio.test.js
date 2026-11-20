import test from 'node:test';
import assert from 'node:assert/strict';
import { readVsdx, themeColors, quickColor, roundPath, shade } from '../ui/src/editors/viewer/vsdx.js';

const THEME = '<a:theme><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2></a:clrScheme></a:themeElements></a:theme>';

test('theme colours and QuickStyle indices', () => {
  const th = themeColors(THEME);
  assert.deepEqual(th, { dk1: '#000000', lt1: '#ffffff', accent1: '#4472c4', accent2: '#ed7d31' });
  assert.equal(quickColor(th, 2), '#4472c4');
  assert.equal(quickColor(th, 103), '#ed7d31');
  assert.equal(shade('#ffffff', -0.5), '#808080');
});

test('rounded corners for straight connector paths', () => {
  assert.equal(roundPath('M0 0L10 0L10 10', 2), 'M0 0L8 0Q10 0 10 2L10 10');
  assert.equal(roundPath('M0 0C1 1 2 2 3 3', 2), 'M0 0C1 1 2 2 3 3');
});

const rect = '<Section N="Geometry" IX="0"><Row T="RelMoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row><Row T="RelLineTo" IX="2"><Cell N="X" V="1"/><Cell N="Y" V="0"/></Row><Row T="RelLineTo" IX="3"><Cell N="X" V="1"/><Cell N="Y" V="1"/></Row><Row T="RelLineTo" IX="4"><Cell N="X" V="0"/><Cell N="Y" V="1"/></Row><Row T="RelLineTo" IX="5"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row></Section>';
test('themed fill, gradient fill', () => {
  const files = {
    'visio/theme/theme1.xml': THEME,
    'visio/pages/pages.xml': '<Pages xmlns:r="r"><Page ID="0" NameU="P"><PageSheet><Cell N="PageWidth" V="4"/><Cell N="PageHeight" V="3"/></PageSheet><Rel r:id="rId1"/></Page></Pages>',
    'visio/pages/_rels/pages.xml.rels': '<Relationships><Relationship Id="rId1" Target="page1.xml"/></Relationships>',
    'visio/pages/page1.xml': `<PageContents><Shapes>
      <Shape ID="1"><Cell N="PinX" V="1"/><Cell N="PinY" V="1"/><Cell N="Width" V="1"/><Cell N="Height" V="1"/><Cell N="FillForegnd" V="Themed"/><Cell N="QuickStyleFillColor" V="3"/>${rect}</Shape>
      <Shape ID="2"><Cell N="PinX" V="3"/><Cell N="PinY" V="1"/><Cell N="Width" V="1"/><Cell N="Height" V="1"/><Cell N="FillGradientEnabled" V="1"/><Cell N="FillGradientAngle" V="0"/>
        <Section N="FillGradient"><Row IX="0"><Cell N="GradientStopColor" V="#ff0000"/><Cell N="GradientStopPosition" V="0"/></Row><Row IX="1"><Cell N="GradientStopColor" V="#0000ff"/><Cell N="GradientStopPosition" V="1"/></Row></Section>${rect}</Shape>
    </Shapes></PageContents>`,
  };
  const svg = readVsdx(files).pages[0].svg;
  assert.match(svg, /fill="#ed7d31"/);
  assert.match(svg, /<linearGradient id="vsg1"[^>]*><stop offset="0" stop-color="#ff0000"\/><stop offset="1" stop-color="#0000ff"\/>/);
  assert.match(svg, /fill="url\(#vsg1\)"/);
});

test('connectors without geometry are routed at right angles around shapes', async () => {
  const { routeOrthogonal } = await import('../ui/src/editors/viewer/vsdx.js');
  const A = { x1: 0, y1: 0, x2: 1, y2: 1 }; const B = { x1: 4, y1: 0, x2: 5, y2: 1 }; const O = { x1: 2, y1: -1, x2: 3, y2: 2 };
  const pts = routeOrthogonal([1, 0.5], [4, 0.5], [A, B, O], A, B);
  // every segment is horizontal or vertical and none crosses the obstacle
  for (let i = 1; i < pts.length; i++) assert.ok(pts[i][0] === pts[i - 1][0] || pts[i][1] === pts[i - 1][1], JSON.stringify(pts));
  for (let i = 1; i < pts.length; i++) { const [a, b] = [pts[i - 1], pts[i]]; const crosses = Math.max(a[0], b[0]) > O.x1 && Math.min(a[0], b[0]) < O.x2 && Math.max(a[1], b[1]) > O.y1 && Math.min(a[1], b[1]) < O.y2; assert.ok(!crosses, JSON.stringify(pts)); }
  assert.deepEqual(pts[0], [1, 0.5]); assert.deepEqual(pts[pts.length - 1], [4, 0.5]);
  // free way: straight
  assert.deepEqual(routeOrthogonal([1, 0.5], [4, 0.5], [A, B], A, B), [[1, 0.5], [4, 0.5]]);
  // in a page: a connector shape with only Begin/End cells and glue
  const files = {
    'visio/pages/pages.xml': '<Pages xmlns:r="r"><Page ID="0" NameU="P"><PageSheet><Cell N="PageWidth" V="6"/><Cell N="PageHeight" V="4"/></PageSheet><Rel r:id="rId1"/></Page></Pages>',
    'visio/pages/_rels/pages.xml.rels': '<Relationships><Relationship Id="rId1" Target="page1.xml"/></Relationships>',
    'visio/pages/page1.xml': `<PageContents><Shapes>
      <Shape ID="1"><Cell N="PinX" V="0.5"/><Cell N="PinY" V="0.5"/><Cell N="Width" V="1"/><Cell N="Height" V="1"/>${rect}</Shape>
      <Shape ID="2"><Cell N="PinX" V="4.5"/><Cell N="PinY" V="2.5"/><Cell N="Width" V="1"/><Cell N="Height" V="1"/>${rect}</Shape>
      <Shape ID="3"><Cell N="BeginX" V="1"/><Cell N="BeginY" V="0.5"/><Cell N="EndX" V="4"/><Cell N="EndY" V="2.5"/><Cell N="EndArrow" V="13"/></Shape>
    </Shapes><Connects><Connect FromSheet="3" FromCell="BeginX" ToSheet="1"/><Connect FromSheet="3" FromCell="EndX" ToSheet="2"/></Connects></PageContents>`,
  };
  const svg = readVsdx(files).pages[0].svg;
  const m = /<path d="(M[^"]+)" fill="none"[^>]*marker-end/.exec(svg);
  assert.ok(m, svg);
  assert.equal((m[1].match(/L/g) || []).length >= 2, true, m[1]); // at least one bend
});
