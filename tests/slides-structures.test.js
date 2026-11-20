'use strict';
/* Office 1.1 structures of the slides chunk: node --test tests/slides-structures.test.js
 * chart workbook embed, threaded comments, emphasis/exit/motion-path animations, kept foreign effects, Morph,
 * custom geometry (icons), multi-stop ODF gradients, flat ODF export, EMF/WMF display. */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const SRC = path.join(__dirname, '..', 'ui', 'src', 'editors');
const load = (p) => import(require('url').pathToFileURL(path.join(SRC, p)).href);
const JSZip = require('jszip');

async function deck(build) {
  const m = await load('slides/model.js');
  let d = m.newDoc({});
  d = m.addSlide(d, 'blank').doc;
  d = m.addSlide(d, 'blank').doc;
  return build(m, d) || d;
}

test('chart: written with an embedded workbook that holds the data ("Edit data" in PowerPoint)', async () => {
  const io = await load('slides/io/index.js');
  const d = await deck((m, d0) => m.addEls(d0, d0.slides[0].id, [m.chart('col', 100000, 100000, 6000000, 3500000)]));
  const bytes = await io.writeAny(d, 'pptx');
  const z = await JSZip.loadAsync(bytes);
  const chart = await z.file('ppt/charts/chart1.xml').async('string');
  assert.match(chart, /<c:externalData r:id="rId1"><c:autoUpdate val="0"\/><\/c:externalData><\/c:chartSpace>$/);
  assert.match(chart, /<c:f>Sheet1!\$B\$2:\$B\$5<\/c:f><c:numCache>/);
  assert.match(chart, /<c:f>Sheet1!\$A\$2:\$A\$5<\/c:f><c:strCache>/);
  const rels = await z.file('ppt/charts/_rels/chart1.xml.rels').async('string');
  assert.match(rels, /relationships\/package" Target="\.\.\/embeddings\/Microsoft_Excel_Worksheet1\.xlsx"/);
  const ct = await z.file('[Content_Types].xml').async('string');
  assert.match(ct, /Extension="xlsx" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet"/);
  const XLSX = require('xlsx');
  const wb = XLSX.read(await z.file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx').async('uint8array'));
  const rows = XLSX.utils.sheet_to_json(wb.Sheets.Sheet1, { header: 1 });
  assert.deepEqual(rows[0].slice(1), ['Series 1', 'Series 2', 'Series 3']);
  assert.deepEqual(rows[1], ['Q1', 4.3, 2.4, 2]);
  // read back: same data
  const { doc } = await io.readAny(bytes, 'pptx');
  const ch = doc.slides[0].elements.find((e) => e.type === 'chart').chart;
  assert.deepEqual(ch.series[1].values, [2.4, 4.4, 1.8, 2.8]);
});

test('comments: replies are written as threads (p15:threadingInfo) with positions and read back', async () => {
  const io = await load('slides/io/index.js');
  const d = await deck((m, d0) => m.updateSlide(d0, d0.slides[0].id, { comments: [
    { id: 'c1', author: 'Anna Berg', text: 'Is this number right?', time: '2026-11-01T10:00:00Z', pos: { x: 1587500, y: 3175000 } },
    { id: 'c2', author: 'Ben', text: 'Yes – checked.', time: '2026-11-01T11:00:00Z', parent: 'c1' },
    { id: 'c3', author: 'Anna Berg', text: 'Second thread', time: '2026-11-02T09:00:00Z' },
  ] }));
  const bytes = await io.writeAny(d, 'pptx');
  const z = await JSZip.loadAsync(bytes);
  const cm = await z.file('ppt/comments/comment1.xml').async('string');
  assert.match(cm, /<p15:threadingInfo [^>]*><p15:parentCm authorId="0" idx="1"\/><\/p15:threadingInfo>/);
  assert.match(cm, /<p:pos x="1000" y="2000"\/>/);
  const { doc } = await io.readAny(bytes, 'pptx');
  const list = doc.slides[0].comments;
  assert.equal(list.length, 3);
  const top = list.find((c) => c.text === 'Is this number right?');
  const reply = list.find((c) => c.text === 'Yes – checked.');
  assert.equal(reply.parent, top.id);
  assert.equal(reply.author, 'Ben');
  assert.deepEqual(top.pos, { x: 1587500, y: 3175000 });
  assert.equal(list.find((c) => c.text === 'Second thread').parent, undefined);
});

test('animations: entrance, emphasis, exit and motion paths round-trip through pptx and odp', async () => {
  const io = await load('slides/io/index.js');
  let ids;
  const d = await deck((m, d0) => {
    const a = m.shape('rect', 900000, 900000, 1800000, 1200000); const b = m.shape('ellipse', 4000000, 900000, 1500000, 1500000);
    ids = [a.id, b.id];
    let x = m.addEls(d0, d0.slides[0].id, [a, b]);
    x = m.updateSlide(x, x.slides[0].id, { animations: [
      { id: 'a1', el: a.id, cls: 'entr', effect: 'float', trigger: 'click', duration: 600, delay: 0 },
      { id: 'a2', el: a.id, cls: 'emph', effect: 'spin', trigger: 'click', duration: 1000, delay: 0, angle: 180 },
      { id: 'a3', el: b.id, cls: 'emph', effect: 'grow', trigger: 'with', duration: 800, delay: 200, scale: 1.25 },
      { id: 'a4', el: b.id, cls: 'path', effect: 'line', trigger: 'after', duration: 2000, delay: 0, path: 'M 0 0 L 0.2 0.1 E' },
      { id: 'a5', el: a.id, cls: 'exit', effect: 'fly', dir: 'l', trigger: 'click', duration: 500, delay: 0 },
      { id: 'a6', el: b.id, cls: 'exit', effect: 'fade', trigger: 'with', duration: 500, delay: 0 },
    ] });
    return x;
  });
  const bytes = await io.writeAny(d, 'pptx');
  const z = await JSZip.loadAsync(bytes);
  const s1 = await z.file('ppt/slides/slide1.xml').async('string');
  assert.match(s1, /presetID="42" presetClass="entr"/);
  assert.match(s1, /presetID="8" presetClass="emph"[\s\S]*<p:animRot by="10800000">/);
  assert.match(s1, /presetClass="path"[\s\S]*<p:animMotion origin="layout" path="M 0 0 L 0\.2 0\.1 E"/);
  assert.match(s1, /presetID="2" presetClass="exit"/);
  assert.match(s1, /<p:animEffect transition="out" filter="fade">/);
  // unique cTn ids
  const cids = [...s1.matchAll(/<p:cTn id="(\d+)"/g)].map((x) => x[1]);
  assert.equal(new Set(cids).size, cids.length);
  const strip = (l) => l.map(({ id, el, ...r }) => r); // eslint-disable-line no-unused-vars
  const { doc } = await io.readAny(bytes, 'pptx');
  const got = doc.slides[0].animations;
  assert.deepEqual(strip(got), strip(d.slides[0].animations).map((a) => ({ ...a, dir: a.dir })).map((a) => { if (!a.dir) delete a.dir; return a; }));
  // odp
  const ob = await io.writeAny(d, 'odp');
  const r2 = await io.readAny(ob, 'odp');
  const o = r2.doc.slides[0].animations;
  assert.deepEqual(o.map((a) => [a.cls || 'entr', a.effect]), [['entr', 'float'], ['emph', 'spin'], ['emph', 'grow'], ['path', 'line'], ['exit', 'fly'], ['exit', 'fade']]);
  assert.equal(o[3].path, 'M 0 0 L 0.2 0.1 E');
  assert.equal(o[2].scale, 1.25);
  assert.ok(ids.length === 2);
});

test('animations: an effect we cannot edit is kept when the others change', async () => {
  const { animationsOf, timingXml } = await load('slides/io/pptx-anim.js');
  const { parseXml } = await load('_shared/xmltree.js');
  const foreign = '<p:par><p:cTn id="5" presetID="35" presetClass="emph" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="6" dur="250" autoRev="1"/><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>';
  const xml = `<p:timing xmlns:p="p"><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3"><p:childTnLst><p:par><p:cTn id="4"><p:childTnLst>${foreign}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
  let n = 0;
  const list = animationsOf(parseXml(xml), { 4: 'eA' }, () => `a${n++}`);
  assert.equal(list.length, 1);
  assert.equal(list[0].effect, 'raw');
  const out = timingXml([...list, { id: 'x', el: 'eB', cls: 'entr', effect: 'fade', trigger: 'click', duration: 500, delay: 0 }], { eA: 9, eB: 10 }, new Set());
  assert.match(out, /presetID="35" presetClass="emph"/);
  assert.match(out, /<p:spTgt spid="9"\/>/);
  assert.match(out, /presetID="10" presetClass="entr"/);
});

test('morph transition: written with a fade fallback and read back', async () => {
  const io = await load('slides/io/index.js');
  const d = await deck((m, d0) => m.updateSlide(d0, d0.slides[1].id, { transition: { type: 'morph', option: 'object', dir: 'l', duration: 1500, advanceAfter: null } }));
  const bytes = await io.writeAny(d, 'pptx');
  const s2 = await (await JSZip.loadAsync(bytes)).file('ppt/slides/slide2.xml').async('string');
  assert.match(s2, /<mc:Choice xmlns:p159="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2015\/09\/main" Requires="p159"><p:transition [^>]*><p159:morph option="byObject"\/><\/p:transition><\/mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade\/>/);
  const { doc } = await io.readAny(bytes, 'pptx');
  assert.equal(doc.slides[1].transition.type, 'morph');
  assert.equal(doc.slides[1].transition.duration, 1500);
  const { morphPairs } = await load('slides/anim.js');
  const a = { elements: [{ id: '1', type: 'shape', name: 'Logo', x: 0 }, { id: '2', type: 'shape', text: { paras: [{ runs: [{ text: 'Hi' }] }] } }] };
  const b = { elements: [{ id: '3', type: 'shape', name: 'Logo', x: 5 }, { id: '4', type: 'shape', text: { paras: [{ runs: [{ text: 'Hi' }] }] } }, { id: '5', type: 'image', src: 'x' }] };
  assert.deepEqual(morphPairs(a, b).map((p) => [p.from.id, p.to.id]), [['1', '3'], ['2', '4']]);
});

test('custom geometry (icons): custGeom in pptx, enhanced path in odp, SVG arcs normalised', async () => {
  const io = await load('slides/io/index.js');
  const cg = await load('slides/render/custom.js');
  const d0 = cg.normalizePath('M12 2a10 10 0 1 0 0 20a10 10 0 1 0 0-20zM8 12h8');
  assert.match(d0, /^M12 2C/);
  assert.ok(!/[aAhH]/.test(d0));
  const paths = [{ w: 24, h: 24, d: d0, fill: false }];
  const d = await deck((m, x) => m.addEls(x, x.slides[0].id, [{ ...m.shape('rect', 100000, 100000, 914400, 914400), geom: 'custom', paths, fill: null, line: { color: 'accent1', width: 19050, dash: 'solid', cap: 'round', join: 'round' }, text: undefined }]));
  const bytes = await io.writeAny(d, 'pptx');
  const s1 = await (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml').async('string');
  assert.match(s1, /<a:custGeom><a:avLst\/><a:gdLst\/><a:ahLst\/><a:cxnLst\/><a:rect l="0" t="0" r="r" b="b"\/><a:pathLst><a:path w="2400" h="2400" fill="none"><a:moveTo><a:pt x="1200" y="200"\/><\/a:moveTo><a:cubicBezTo>/);
  assert.match(s1, /<a:ln w="19050" cap="rnd">.*<a:round\/>/);
  const r = await io.readAny(bytes, 'pptx');
  const e = r.doc.slides[0].elements[0];
  assert.equal(e.geom, 'custom');
  assert.equal(e.paths[0].fill, false);
  assert.equal(e.line.cap, 'round');
  assert.equal(cg.segments(e.paths[0].d).length, cg.segments(d0).length);
  const ob = await io.writeAny(d, 'odp');
  const r2 = await io.readAny(ob, 'odp');
  assert.equal(r2.doc.slides[0].elements[0].geom, 'custom');
  // a foreign ODF enhanced path without equations
  const p = cg.pathsOfOdf('M 0 0 L 100 0 100 100 Z N F M 10 10 L 20 20 N', '0 0 100 100');
  assert.equal(p.length, 2);
  assert.equal(p[1].fill, false);
});

test('odp: gradients with more than two stops (LibreOffice color-stops) and flat ODF export', async () => {
  const io = await load('slides/io/index.js');
  const fill = { type: 'grad', angle: 0, stops: [{ pos: 0, color: '#FF0000' }, { pos: 50, color: '#00FF00' }, { pos: 100, color: '#0000FF' }] };
  const d = await deck((m, x) => {
    let y = m.addEls(x, x.slides[0].id, [{ ...m.shape('rect', 100000, 100000, 3000000, 2000000), fill }]);
    y = m.addEls(y, y.slides[0].id, [m.image('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==', 4000000, 100000, 914400, 914400)]);
    return y;
  });
  const ob = await io.writeAny(d, 'odp');
  const st = await (await JSZip.loadAsync(ob)).file('styles.xml').async('string');
  assert.match(st, /<loext:color-stop svg:offset="0\.5" loext:color-type="rgb" loext:color-value="#00ff00"\/>/);
  const flat = await io.writeAny(d, 'fodp');
  const xml = new TextDecoder().decode(flat);
  assert.match(xml, /^<\?xml[^>]*\?>\n<office:document [^>]*office:mimetype="application\/vnd\.oasis\.opendocument\.presentation"/);
  assert.match(xml, /<office:binary-data>iVBOR/);
  assert.ok(!/xlink:href="Pictures\//.test(xml));
  const r = await io.readAny(flat, 'fodp');
  assert.equal(r.fmt, 'fodp');
  assert.equal(r.doc.slides.length, 2);
  const els = r.doc.slides[0].elements;
  const g = els.find((e) => e.fill && e.fill.type === 'grad');
  assert.deepEqual(g.fill.stops.map((s) => s.pos), [0, 50, 100]);
  assert.ok(els.some((e) => e.type === 'image' && /^data:image\/png/.test(e.src)));
});

test('metafiles: WMF and EMF records become SVG for display', async () => {
  const { wmfToSvg, emfToSvg, metaToSvgUrl } = await load('slides/render/metafile.js');
  // WMF (placeable header + header + CreateBrush + Select + Rectangle + EOF)
  const w = [];
  const u16 = (v) => w.push(v & 0xff, (v >> 8) & 0xff); const u32 = (v) => { u16(v & 0xffff); u16((v >>> 16) & 0xffff); };
  u32(0x9ac6cdd7); u16(0); u16(0); u16(0); u16(1000); u16(500); u16(1440); u32(0); u16(0);
  u16(1); u16(9); u16(0x300); u32(30); u16(1); u32(8); u16(0);
  u32(7); u16(0x02fc); u16(0); u32(0x0000ff); u16(0); // brush red
  u32(4); u16(0x012d); u16(0);
  u32(7); u16(0x041b); u16(400); u16(900); u16(100); u16(100);
  u32(3); u16(0);
  const svg = wmfToSvg(new Uint8Array(w));
  assert.match(svg, /viewBox="0 0 1000 500"/);
  assert.match(svg, /<rect x="100" y="100" width="800" height="300" fill="#ff0000"/);
  // EMF: header + CREATEBRUSHINDIRECT + SELECTOBJECT + RECTANGLE + EOF
  const e = [];
  const E32 = (v) => { e.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff); };
  E32(1); E32(108); E32(0); E32(0); E32(99); E32(49); E32(0); E32(0); E32(2645); E32(1322); E32(0x464d4520); E32(0x10000); E32(0); E32(4); E32(0); E32(0); E32(0);
  E32(100); E32(50); E32(26); E32(13); E32(0); E32(0); E32(0); E32(0); E32(0); E32(0);
  while (e.length < 108) e.push(0);
  E32(39); E32(24); E32(1); E32(0); E32(0x00ff00); E32(0);
  E32(37); E32(12); E32(1);
  E32(43); E32(24); E32(10); E32(10); E32(60); E32(40);
  E32(14); E32(20); E32(0); E32(0x10); E32(20);
  const es = emfToSvg(new Uint8Array(e));
  assert.match(es, /<rect x="10" y="10" width="50" height="30" fill="#00ff00"/);
  const url = metaToSvgUrl(`data:image/x-wmf;base64,${Buffer.from(w).toString('base64')}`);
  assert.match(url, /^data:image\/svg\+xml/);
  assert.equal(metaToSvgUrl('data:image/png;base64,AAAA'), null);
});
