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
