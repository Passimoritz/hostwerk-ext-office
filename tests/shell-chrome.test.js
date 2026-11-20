// Window chrome of the editor screen (ui/src/shell/chrome.js) and the template maps (ui/src/lib/formats.js) – node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import { overlayGeometry, isOwnWindow } from '../ui/src/shell/chrome.js';
import { TEMPLATE_OF, DOC_OF, isTemplateName, templateExtFor } from '../ui/src/lib/formats.js';

test('window controls overlay: free space on the right, own windows and the browser need none', () => {
  const wco = { visible: true, getTitlebarAreaRect: () => ({ x: 0, y: 0, width: 1300, height: 40 }) };
  assert.deepEqual(overlayGeometry({ wco, innerWidth: 1440, zoom: 1, desktop: true, own: false, full: false }), { on: true, right: 140, height: 40 });
  // Hostwerk's interface zoom (125 %) – CSS px of the page
  assert.deepEqual(overlayGeometry({ wco, innerWidth: 1440, zoom: 1.25, desktop: true, own: false, full: false }), { on: true, right: 112, height: 32 });
  // no WCO API: Hostwerk's own 150 px
  assert.deepEqual(overlayGeometry({ wco: null, innerWidth: 1440, zoom: 1, desktop: true, own: false, full: false }), { on: true, right: 150, height: 40 });
  assert.equal(overlayGeometry({ wco, innerWidth: 1440, zoom: 1, desktop: false, own: false, full: false }).on, false);
  assert.equal(overlayGeometry({ wco, innerWidth: 1440, zoom: 1, desktop: true, own: true, full: false }).on, false);
  assert.equal(overlayGeometry({ wco, innerWidth: 1440, zoom: 1, desktop: true, own: false, full: true }).on, false);
  assert.equal(overlayGeometry({ wco: { visible: false }, innerWidth: 1440, zoom: 1, desktop: true, own: false, full: false }).on, false);
  assert.ok(isOwnWindow('#/office-edit?path=a.docx&window=1'));
  assert.ok(!isOwnWindow('#/office-edit?path=window=1x.docx'));
});

test('template formats', () => {
  assert.equal(templateExtFor('Brief.docx'), 'dotx');
  assert.equal(templateExtFor('Tabelle.ods'), 'ots');
  assert.equal(templateExtFor('x.csv'), null);
  assert.ok(isTemplateName('Normal.dotm') && isTemplateName('a.OTP') && !isTemplateName('a.docx'));
  for (const [doc, tpl] of Object.entries(TEMPLATE_OF)) if (doc !== tpl && DOC_OF[tpl] && !['ppsx', 'ppsm'].includes(doc)) assert.equal(DOC_OF[tpl], doc);
});
