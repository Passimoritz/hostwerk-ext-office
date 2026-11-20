import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import JSZip from 'jszip';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { readXlsx, writeXlsx } from '../ui/src/editors/sheets/io/xlsx.js';
import { patchChartXml } from '../ui/src/editors/sheets/io/xlsxChartPatch.js';

const loc = makeLocale({ lang: 'en' });
const SRC = fs.readFileSync(new URL('./fixtures/excel-chart.xlsx', import.meta.url));
const K = (r, c) => r * 16384 + c;

async function edit(fn) {
  const b = await readXlsx(SRC, { loc, ext: 'xlsx' });
  const s = b.sheets[0];
  fn(b, s, s.charts[0]);
  const out = await writeXlsx(b, { ext: 'xlsx' });
  const z = await JSZip.loadAsync(out);
  return { xml: await z.file('xl/charts/chart1.xml').async('string'), out };
}

test('edited Excel chart: title, series range and colour are patched into the original part', async () => {
  const orig = await (await JSZip.loadAsync(SRC)).file('xl/charts/chart1.xml').async('string');
  const { xml, out } = await edit((b, s, ch) => {
    b.input(s, 0, 2, 'More'); b.input(s, 1, 2, '5'); b.input(s, 2, 2, '6'); b.input(s, 3, 2, '7');
    ch.title = 'Fruit 2026';
    ch.series = [{ ...ch.series[0], valRef: 'Tabelle1!$C$2:$C$4', nameRef: 'Tabelle1!$C$1', color: '#ff0000' }];
  });
  assert.match(xml, /<a:t>Fruit 2026<\/a:t>/);
  assert.match(xml, /<c:f>Tabelle1!\$C\$2:\$C\$4<\/c:f>/);
  assert.match(xml, /<c:f>Tabelle1!\$C\$1<\/c:f>/);
  assert.match(xml, /<c:v>5<\/c:v>/);
  assert.match(xml, /<a:srgbClr val="FF0000"\/>/);
  // everything else of Excel's part is still there
  for (const keep of ['c14:style val="102"', 'c16r3:dataDisplayOptions16', '<c:gapWidth val="75"/>', '<c:overlap val="-25"/>', '<c:legendPos val="t"/>', 'c16:uniqueId', '<c:printSettings>']) assert.ok(xml.includes(keep), keep);
  assert.ok(orig.length > 0);
  const back = await readXlsx(out, { loc, ext: 'xlsx' });
  assert.equal(back.sheets[0].charts[0].title, 'Fruit 2026');
  assert.equal(back.sheets[0].charts[0].series[0].valRef, 'Tabelle1!$C$2:$C$4');
});

test('edited Excel chart: type change column → line / bar and stacking keep the part', async () => {
  const line = (await edit((b, s, ch) => { ch.type = 'line'; })).xml;
  assert.match(line, /<c:lineChart><c:grouping val="standard"\/>/);
  assert.ok(!/barDir|gapWidth|invertIfNegative/.test(line));
  assert.match(line, /<c:marker val="1"\/><c:axId/);
  assert.ok(line.includes('c14:style val="102"'));
  const bar = (await edit((b, s, ch) => { ch.type = 'bar'; ch.stacked = 'stacked'; })).xml;
  assert.match(bar, /<c:barDir val="bar"\/><c:grouping val="stacked"\/>/);
  assert.match(bar, /<c:overlap val="100"\/>/);
  assert.match(bar, /<c:catAx>[\s\S]*?<c:axPos val="l"\/>/);
});

test('edited Excel chart: legend off, axis title added, a second series cloned', async () => {
  const { xml } = await edit((b, s, ch) => {
    b.input(s, 1, 2, '1'); b.input(s, 2, 2, '2'); b.input(s, 3, 2, '3');
    ch.legend = 'none'; ch.yTitle = 'Pieces';
    ch.series = [...ch.series, { name: 'Two', nameRef: null, catRef: 'Tabelle1!$A$2:$A$4', valRef: 'Tabelle1!$C$2:$C$4', color: null }];
  });
  assert.ok(!xml.includes('<c:legend>'));
  assert.match(xml, /<c:valAx>[\s\S]*<a:t>Pieces<\/a:t>/);
  assert.equal((xml.match(/<c:ser>/g) || []).length, 2);
  assert.match(xml, /<c:tx><c:v>Two<\/c:v><\/c:tx>/);
  assert.match(xml, /<c:idx val="101"\/><c:order val="1"\/>/);
});

test('unsupported chart edits fall back (null)', () => {
  const x = '<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:pieChart><c:ser/></c:pieChart></c:plotArea></c:chart></c:chartSpace>';
  assert.equal(patchChartXml(x, { type: 'pie', series: [{}] }, { type: 'column', series: [{}] }, {}), null);
  assert.equal(patchChartXml(x, { type: 'pie', labels: false, series: [{}] }, { type: 'pie', labels: true, series: [{}] }, {}), null);
  void K;
});
