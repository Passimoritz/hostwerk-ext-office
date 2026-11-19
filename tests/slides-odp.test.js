'use strict';
/* ODP writer → reader round trip and flat ODP (fodp) parsing of the slides chunk. */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const SRC = path.join(__dirname, '..', 'ui', 'src', 'editors', 'slides');
const imp = (p) => import(`file://${path.join(SRC, p).replace(/\\/g, '/')}`);
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const near = (a, b, tol = 400) => assert.ok(Math.abs(a - b) <= tol, `${a} ≈ ${b}`);

async function sampleDoc() {
  const M = await imp('model.js');
  let doc = M.newDoc({ theme: 'hostwerk', title: 'Test deck', author: 'Tester' });
  doc = M.addSlide(doc, 'title', 0, { ctrTitle: 'Hello ODP', subTitle: 'Round trip' }).doc;
  doc = M.addSlide(doc, 'titleContent', 1, { title: 'Agenda' }).doc;
  const s2 = doc.slides[1];
  const bodyEl = s2.elements.find((e) => e.ph === 'body');
  doc = M.updateEls(doc, s2.id, [bodyEl.id], (e) => ({ ...e, text: { ...e.text, paras: [
    { runs: [{ text: 'First', b: true }, { text: ' point', color: '#FF0000' }], level: 0, bullet: { char: '•' } },
    { runs: [{ text: 'Sub  item', i: true, size: 20 }], level: 1, bullet: { char: '–' } },
    { runs: [{ text: 'Second' }], level: 0, bullet: { num: 'arabicPeriod', start: 1 } },
  ] } }));
  const star = M.shape('star5', 914400, 914400, 1828800, 1371600, { rot: 30, fill: { type: 'grad', angle: 45, stops: [{ pos: 0, color: '#FF0000' }, { pos: 100, color: '#0000FF' }] } });
  star.text = M.bodyFromText('Star', { anchor: 'ctr' }, { align: 'ctr' });
  const ln = M.line(3000000, 4000000, 2000000, 500000, { line: { color: '#00AA00', width: 25400, dash: 'dash', head: 'none', tail: 'triangle' } });
  const img = M.image(PNG, 6000000, 1000000, 1000000, 1000000, { opacity: 0.5, alt: 'red dot' });
  const tbl = M.table(2, 3, 1000000, 5000000, 6000000, 1000000);
  tbl.rows[0].cells[0] = M.cell('Merged', { span: { c: 2, r: 1 } });
  tbl.rows[0].cells[1] = { ...M.cell(''), merged: true };
  tbl.rows[1].cells[2] = M.cell('C', { fill: { type: 'solid', color: '#FFFF00' } });
  const r1 = M.shape('rect', 8000000, 4000000, 500000, 500000);
  const r2 = M.shape('ellipse', 8800000, 4200000, 500000, 700000);
  doc = M.addEls(doc, s2.id, [star, ln, img, tbl, r1, r2]);
  const g = M.groupEls(doc, s2.id, [r1.id, r2.id]);
  doc = g.doc;
  doc = M.updateSlide(doc, s2.id, (s) => ({ ...s, notes: 'Speaker notes\nLine 2', transition: { type: 'fade', dir: 'l', duration: 700, advanceAfter: 5000 }, animations: [{ id: 'a1', el: star.id, effect: 'fly', dir: 'b', trigger: 'click', duration: 600, delay: 0 }, { id: 'a2', el: img.id, effect: 'fade', trigger: 'after', duration: 400, delay: 100 }] }));
  doc = M.addSlide(doc, 'blank', 2).doc;
  doc = M.updateSlide(doc, doc.slides[2].id, { hidden: true, background: { type: 'solid', color: '#112233' } });
  return { doc, ids: { star: star.id, ln: ln.id, img: img.id, tbl: tbl.id, group: g.id } };
}

for (const extras of [true, false]) {
  test(`odp write → read round trip (extras ${extras})`, async () => {
    const { writeOdp } = await imp('io/odp-write.js');
    const { readOdp } = await imp('io/odp-read.js');
    const { doc } = await sampleDoc();
    const bytes = await writeOdp(doc, { extras });
    const { doc: back } = await readOdp(bytes, { ext: 'odp' });
    assert.equal(back.slides.length, 3);
    near(back.size.w, doc.size.w); near(back.size.h, doc.size.h);
    const all = (s) => { const out = []; const w = (l) => l.forEach((e) => { out.push(e); if (e.children) w(e.children); }); w(s.elements); return out; };
    const txt = (e) => e.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
    const s1 = back.slides[0];
    assert.ok(all(s1).some((e) => e.text && txt(e) === 'Hello ODP' && /title/i.test(e.ph)));
    assert.ok(all(s1).some((e) => e.text && txt(e) === 'Round trip' && e.ph === 'subTitle'));
    const s2 = back.slides[1];
    const body = all(s2).find((e) => e.ph === 'body');
    assert.equal(txt(body), 'First point\nSub  item\nSecond');
    assert.deepEqual(body.text.paras.map((p) => p.level), [0, 1, 0]);
    assert.equal(body.text.paras[0].runs[0].b, true);
    assert.equal(body.text.paras[0].runs[1].color, '#FF0000');
    assert.equal(body.text.paras[1].runs[0].i, true);
    assert.equal(body.text.paras[1].runs[0].size, 20);
    assert.deepEqual(body.text.paras[0].bullet, { char: '•' });
    assert.equal(body.text.paras[1].bullet.char, '–');
    assert.equal(body.text.paras[2].bullet.num, 'arabicPeriod');
    const star = all(s2).find((e) => e.type === 'shape' && e.geom === 'star5');
    assert.ok(star, 'star shape');
    near(star.x, 914400); near(star.y, 914400); near(star.w, 1828800); near(star.h, 1371600);
    assert.ok(Math.abs(star.rot - 30) < 0.1, `rot ${star.rot}`);
    assert.equal(star.fill.type, 'grad');
    assert.equal(star.fill.angle, 45);
    assert.equal(star.fill.stops[0].color, '#FF0000');
    assert.equal(txt(star), 'Star');
    const ln = all(s2).find((e) => e.type === 'line');
    near(ln.x, 3000000); near(ln.w, 2000000); near(ln.h, 500000);
    assert.equal(ln.line.tail, 'triangle');
    assert.equal(ln.line.dash, 'dash');
    assert.equal(ln.line.color, '#00AA00');
    const img = all(s2).find((e) => e.type === 'image');
    assert.equal(img.src, PNG);
    assert.equal(img.opacity, 0.5);
    const tbl = all(s2).find((e) => e.type === 'table');
    assert.equal(tbl.rows.length, 2); assert.equal(tbl.cols.length, 3);
    assert.deepEqual(tbl.rows[0].cells[0].span, { c: 2, r: 1 });
    assert.equal(tbl.rows[0].cells[1].merged, true);
    assert.equal(txt(tbl.rows[0].cells[0]), 'Merged');
    assert.equal(tbl.rows[1].cells[2].fill.color, '#FFFF00');
    const grp = s2.elements.find((e) => e.type === 'group');
    assert.ok(grp && grp.children.length === 2, 'group');
    assert.equal(s2.notes, 'Speaker notes\nLine 2');
    assert.equal(s2.transition.type, 'fade');
    assert.equal(s2.transition.advanceAfter, 5000);
    assert.equal(s2.animations.length, 2);
    assert.equal(s2.animations[0].effect, 'fly');
    assert.equal(s2.animations[0].dir, 'b');
    assert.equal(s2.animations[0].el, star.id);
    assert.equal(s2.animations[1].effect, 'fade');
    assert.equal(s2.animations[1].trigger, 'after');
    assert.equal(s2.animations[1].duration, 400);
    assert.equal(back.slides[2].hidden, true);
    assert.equal(back.slides[2].background.color, '#112233');
    assert.equal(back.meta.title, 'Test deck');
    assert.equal(back.meta.author, 'Tester');
    if (extras) {
      assert.equal(back.theme.id, 'hostwerk');
      assert.equal(back.slides[1].id, doc.slides[1].id);
      assert.equal(back.slides[1].transition.duration, 700);
      assert.deepEqual(back.layouts.map((l) => l.id), doc.layouts.map((l) => l.id));
      assert.deepEqual(body.text.lvl, doc.slides[1].elements.find((e) => e.ph === 'body').text.lvl);
      assert.equal(tbl.rows[0].cells[2].fill, undefined, 'implicit table-style fill stays implicit');
    } else {
      assert.equal(back.theme.id, 'imported');
      assert.ok(back.layouts.some((l) => l.id === back.slides[1].layoutId));
    }
  });
}

test('odp zip: mimetype first + stored, all parts parse, manifest complete', async () => {
  const { writeOdp } = await imp('io/odp-write.js');
  const { parseXml, all } = await imp('../_shared/xmltree.js');
  const JSZip = require('jszip');
  const { doc } = await sampleDoc();
  const bytes = await writeOdp(doc);
  assert.deepEqual([...bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
  assert.equal(bytes[8] | (bytes[9] << 8), 0, 'mimetype stored');
  const nameLen = bytes[26] | (bytes[27] << 8);
  assert.equal(Buffer.from(bytes.slice(30, 30 + nameLen)).toString(), 'mimetype');
  assert.equal(Buffer.from(bytes.slice(30 + nameLen + (bytes[28] | (bytes[29] << 8)), 30 + nameLen + (bytes[28] | (bytes[29] << 8)) + 39)).toString(), 'application/vnd.oasis.opendocument.pres');
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  for (const n of names.filter((x) => /\.xml$/.test(x))) {
    const s = await zip.file(n).async('string');
    const root = parseXml(s);
    assert.ok(root && root.n, `${n} parses`);
    assert.equal((s.match(/</g) || []).length - (s.match(/<\?|<!--/g) || []).length > 0, true);
  }
  const man = parseXml(await zip.file('META-INF/manifest.xml').async('string'));
  const listed = new Set(all(man, 'manifest:file-entry').map((e) => e.a['manifest:full-path']));
  for (const n of names) if (n !== 'mimetype' && n !== 'META-INF/manifest.xml') assert.ok(listed.has(n), `manifest lists ${n}`);
  assert.ok(names.some((n) => /^Pictures\/.+\.png$/.test(n)));
  // styles.xml has one master page per used layout, content has 3 pages
  const content = await zip.file('content.xml').async('string');
  assert.equal((content.match(/<draw:page /g) || []).length, 3);
});

test('patch save keeps unknown parts, preserved elements, raw animations and comments', async () => {
  const { writeOdp } = await imp('io/odp-write.js');
  const { readOdp } = await imp('io/odp-read.js');
  const JSZip = require('jszip');
  const { doc } = await sampleDoc();
  const zip = await JSZip.loadAsync(await writeOdp(doc));
  const extra = {
    'Object 1/content.xml': '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart chart:class="chart:bar"/></office:chart></office:body></office:document-content>',
    'Object 1/styles.xml': '<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
    'ObjectReplacements/Object 1': new Uint8Array([1, 2, 3, 4, 250, 251]),
    'Basic/script-lc.xml': '<?xml version="1.0" encoding="UTF-8"?><library:libraries xmlns:library="http://openoffice.org/2000/library"><library:library library:name="Standard" library:link="false"/></library:libraries>',
    'Basic/Standard/Module1.xml': '<?xml version="1.0" encoding="UTF-8"?><script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">Sub Main\nMsgBox "hi"\nEnd Sub</script:module>',
    'Configurations2/accelerator/current.xml': '',
  };
  for (const [p, v] of Object.entries(extra)) zip.file(p, v);
  let man = await zip.file('META-INF/manifest.xml').async('string');
  man = man.replace('</manifest:manifest>', '<manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/vnd.oasis.opendocument.chart"/><manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/></manifest:manifest>');
  zip.file('META-INF/manifest.xml', man);
  let content = await zip.file('content.xml').async('string');
  content = content.replace('</office:automatic-styles>', '<style:style style:name="gr77" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#123456"/></style:style></office:automatic-styles>');
  const raw = '<draw:frame draw:style-name="gr77" draw:id="chart1" xml:id="chart1" svg:x="1cm" svg:y="1cm" svg:width="5cm" svg:height="4cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><draw:image xlink:href="./ObjectReplacements/Object 1" xlink:type="simple"/></draw:frame>'
    + '<draw:polygon draw:style-name="gr77" svg:x="7cm" svg:y="1cm" svg:width="3cm" svg:height="3cm" svg:viewBox="0 0 3000 3000" draw:points="0,0 3000,0 1500,3000"/>'
    + '<officeooo:annotation svg:x="2cm" svg:y="2cm"><dc:creator>Ann</dc:creator><dc:date>2026-01-02T03:04:05</dc:date><text:p>Check this</text:p></officeooo:annotation>'
    + '<anim:par presentation:node-type="timing-root"><anim:seq presentation:node-type="main-sequence"><anim:par smil:begin="next"><anim:par smil:begin="0s"><anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="on-click" presentation:preset-class="emphasis" presentation:preset-id="ooo-emphasis-spin"><anim:animateTransform smil:dur="2s" smil:fill="hold" smil:targetElement="chart1" svg:type="rotate" smil:by="360"/></anim:par></anim:par></anim:par></anim:seq></anim:par>';
  const p1 = content.indexOf('<presentation:notes');
  content = content.slice(0, p1) + raw + content.slice(p1);
  zip.file('content.xml', content);
  const original = await zip.generateAsync({ type: 'uint8array' });

  const { doc: d1, fidelity } = await readOdp(original);
  assert.ok(fidelity.includes('charts'));
  const s1 = d1.slides[0];
  const ch = s1.elements.find((e) => e.preserved && e.type === 'unknown');
  assert.ok(ch && ch.what === 'Chart' && ch.rawParts.includes('Object 1') && ch.rawStyles.some((x) => /o_gr77/.test(x)));
  const poly = s1.elements.find((e) => e.preserved && e.type === 'shape');
  assert.ok(poly, 'polygon preserved');
  assert.equal(s1.comments.length, 1); assert.equal(s1.comments[0].author, 'Ann'); assert.equal(s1.comments[0].text, 'Check this');
  assert.ok(s1.timingRaw && /ooo-emphasis-spin/.test(s1.timingRaw.xml));

  const same = async (z1, z2, p) => assert.deepEqual(Buffer.from(await z2.file(p).async('uint8array')), Buffer.from(await z1.file(p).async('uint8array')), `${p} byte-identical`);
  const out = await writeOdp(d1, { original });
  const z0 = await JSZip.loadAsync(original); const z2 = await JSZip.loadAsync(out);
  for (const p of [...Object.keys(extra), 'settings.xml', 'styles.xml']) await same(z0, z2, p);
  const c2 = await z2.file('content.xml').async('string');
  assert.ok(c2.includes('xlink:href="./Object 1"') && c2.includes('style:name="o_gr77"') && c2.includes('draw:points="0,0 3000,0 1500,3000"'));
  assert.ok(c2.includes('ooo-emphasis-spin') && c2.includes('Check this'));
  const m2 = await z2.file('META-INF/manifest.xml').async('string');
  assert.ok(m2.includes('Basic/script-lc.xml') && m2.includes('Object 1/content.xml'));
  // second round trip keeps everything again
  const { doc: d2 } = await readOdp(out);
  assert.ok(d2.slides[0].elements.some((e) => e.preserved && e.what === 'Chart'));
  assert.equal(d2.slides[0].comments[0].text, 'Check this');

  // moved preserved element: only the geometry is patched
  const moved = JSON.parse(JSON.stringify(d1));
  const mc = moved.slides[0].elements.find((e) => e.preserved && e.type === 'unknown');
  mc.x += 360000;
  const c3 = await (await JSZip.loadAsync(await writeOdp(moved, { original }))).file('content.xml').async('string');
  assert.ok(/<draw:frame draw:style-name="o_gr77"[^>]*svg:x="2cm"[^>]*><draw:object xlink:href="\.\/Object 1"/.test(c3));
  // changed theme: styles.xml regenerated, unknown parts still kept
  const themed = JSON.parse(JSON.stringify(d1));
  themed.theme.colors.accent1 = '#00FF00';
  const z4 = await JSZip.loadAsync(await writeOdp(themed, { original }));
  assert.notEqual(await z4.file('styles.xml').async('string'), await z0.file('styles.xml').async('string'));
  await same(z0, z4, 'Basic/Standard/Module1.xml');
  // changed animations: the raw timing is replaced by the written one
  const anim = JSON.parse(JSON.stringify(d1));
  anim.slides[0].animations = [{ id: 'x', el: anim.slides[0].elements[0].id, effect: 'fade', trigger: 'click', duration: 500, delay: 0 }];
  const c5 = await (await JSZip.loadAsync(await writeOdp(anim, { original }))).file('content.xml').async('string');
  assert.ok(!c5.includes('ooo-emphasis-spin') && c5.includes('ooo-entrance-fade-in'));
  // without the original package a preserved element falls back (chart without picture is dropped, not broken)
  const plain = await (await JSZip.loadAsync(await writeOdp(d1))).file('content.xml').async('string');
  assert.ok(!plain.includes('Object 1'));
});

test('fodp (flat ODF) is parsed', async () => {
  const { readOdp } = await imp('io/odp-read.js');
  const fodp = `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation">
 <office:meta><dc:title>Flat</dc:title></office:meta>
 <office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'"/></office:font-face-decls>
 <office:styles>
  <style:default-style style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#729fcf"/><style:text-properties fo:font-size="18pt" style:font-name="Liberation Sans"/></style:default-style>
  <style:style style:name="Default-title" style:family="presentation"><style:text-properties fo:font-size="44pt" fo:color="#222222"/></style:style>
  <style:style style:name="Default-outline1" style:family="presentation"><style:graphic-properties><text:list-style style:name="LO"><text:list-level-style-bullet text:level="1" text:bullet-char="●"/><text:list-level-style-bullet text:level="2" text:bullet-char="–"/></text:list-style></style:graphic-properties><style:text-properties fo:font-size="32pt"/></style:style>
  <style:style style:name="Default-outline2" style:family="presentation" style:parent-style-name="Default-outline1"><style:text-properties fo:font-size="28pt"/></style:style>
 </office:styles>
 <office:automatic-styles>
  <style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="25.4cm" fo:page-height="19.05cm"/></style:page-layout>
  <style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" presentation:duration="PT3S" smil:type="pushWipe" smil:subtype="fromRight" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"/></style:style>
  <style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff8000" svg:stroke-color="#000000" draw:stroke="solid"/></style:style>
  <style:style style:name="pr1" style:family="presentation" style:parent-style-name="Default-title"/>
  <style:style style:name="pr2" style:family="presentation" style:parent-style-name="Default-outline1"/>
  <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
 </office:automatic-styles>
 <office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>
 <office:body><office:presentation>
  <draw:page draw:name="p1" draw:style-name="dp1" draw:master-page-name="Default">
   <draw:frame presentation:style-name="pr1" presentation:class="title" svg:x="1cm" svg:y="1cm" svg:width="20cm" svg:height="3cm"><draw:text-box><text:p>Flat <text:span text:style-name="T1">title</text:span></text:p></draw:text-box></draw:frame>
   <draw:frame presentation:style-name="pr2" presentation:class="outline" svg:x="1cm" svg:y="5cm" svg:width="20cm" svg:height="10cm"><draw:text-box><text:list><text:list-item><text:p>One</text:p><text:list><text:list-item><text:p>Two</text:p></text:list-item></text:list></text:list-item></text:list></draw:text-box></draw:frame>
   <draw:custom-shape draw:style-name="gr1" svg:width="4cm" svg:height="2cm" draw:transform="rotate (1.5707963) translate (10cm 15cm)"><text:p>Oval</text:p><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="ellipse"/></draw:custom-shape>
   <presentation:notes><draw:frame presentation:class="notes"><draw:text-box><text:p>My notes</text:p></draw:text-box></draw:frame></presentation:notes>
  </draw:page>
 </office:presentation></office:body>
</office:document>`;
  const { doc } = await readOdp(new TextEncoder().encode(fodp), { ext: 'fodp' });
  assert.equal(doc.slides.length, 1);
  assert.equal(doc.size.w, 9144000);
  const s = doc.slides[0];
  const title = s.elements.find((e) => e.ph === 'title');
  assert.equal(title.text.paras[0].runs.map((r) => r.text).join(''), 'Flat title');
  assert.equal(title.text.paras[0].runs[1].b, true);
  assert.equal(title.text.lvl[0].size, 44);
  const body = s.elements.find((e) => e.ph === 'body');
  assert.deepEqual(body.text.paras.map((p) => [p.level, p.bullet.char]), [[0, '●'], [1, '–']]);
  assert.equal(body.text.lvl[1].size, 28);
  const oval = s.elements.find((e) => e.geom === 'ellipse');
  assert.equal(oval.fill.color, '#FF8000');
  assert.ok(Math.abs(oval.rot - 270) < 0.01, `rot ${oval.rot}`);
  // rotate 90° ccw around top-left, then translate: centre = (10cm + 1cm, 15cm - 2cm)
  assert.ok(Math.abs(oval.x + oval.w / 2 - 11 * 360000) < 50 && Math.abs(oval.y + oval.h / 2 - 13 * 360000) < 50);
  assert.equal(s.notes, 'My notes');
  assert.equal(s.transition.type, 'push');
  assert.equal(s.transition.advanceAfter, 3000);
  assert.equal(doc.meta.title, 'Flat');
  assert.equal(doc.theme.fonts.minor, 'Liberation Sans');
});
