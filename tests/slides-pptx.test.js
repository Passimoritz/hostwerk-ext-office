'use strict';
/* PPTX import/export of the slides chunk: node --test tests/slides-pptx.test.js */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const SRC = path.join(__dirname, '..', 'ui', 'src', 'editors');
const load = (p) => import(require('url').pathToFileURL(path.join(SRC, p)).href);
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';
const near = (a, b, tol = 2000) => assert.ok(Math.abs(a - b) <= tol, `${a} ≈ ${b}`);
const allEls = (els, out = []) => { for (const e of els) { out.push(e); if (e.type === 'group') allEls(e.children, out); } return out; };
const textOf = (e) => (e.text ? e.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n') : '');

test('reads a pptx generated with pptxgenjs (text, positions, shape, table, notes)', async () => {
  const PptxGenJS = require('pptxgenjs');
  const { readPptx } = await load('slides/io/pptx-read.js');
  const p = new PptxGenJS();
  p.layout = 'LAYOUT_WIDE';
  const s = p.addSlide();
  s.addText('Hello title', { x: 1, y: 0.5, w: 8, h: 1, fontSize: 32, bold: true, color: '1F4E79' });
  s.addText([{ text: 'First point', options: { bullet: true } }, { text: 'Second point', options: { bullet: true, indentLevel: 1 } }], { x: 1, y: 1.6, w: 6, h: 2, fontSize: 20 });
  s.addShape(p.ShapeType.ellipse, { x: 9, y: 2, w: 2, h: 1, fill: { color: 'FF0000' }, line: { color: '0000FF', width: 2 } });
  s.addTable([[{ text: 'A' }, { text: 'B' }], [{ text: 'C' }, { text: 'D' }]], { x: 1, y: 4, w: 4 });
  s.addNotes('My notes here');
  const s2 = p.addSlide();
  s2.addText('Second slide', { x: 0.5, y: 0.5, w: 5, h: 1 });
  const bytes = await p.write({ outputType: 'uint8array' });
  const { doc } = await readPptx(bytes, { ext: 'pptx' });
  assert.equal(doc.size.w, 12192000);
  assert.equal(doc.size.h, 6858000);
  assert.equal(doc.slides.length, 2);
  const els = doc.slides[0].elements;
  const title = els.find((e) => textOf(e) === 'Hello title');
  assert.ok(title, 'title text found');
  near(title.x, 914400); near(title.y, 457200); near(title.w, 8 * 914400); near(title.h, 914400);
  const run = title.text.paras[0].runs[0];
  assert.equal(run.size, 32);
  assert.equal(run.b, true);
  assert.equal(run.color, '#1F4E79');
  const list = els.find((e) => /First point/.test(textOf(e)));
  assert.equal(list.text.paras.length, 2);
  assert.equal(list.text.paras[1].level, 1);
  assert.ok(list.text.paras[0].bullet && list.text.paras[0].bullet.char, 'bullet char');
  const ell = els.find((e) => e.type === 'shape' && e.geom === 'ellipse');
  assert.ok(ell, 'ellipse');
  assert.deepEqual(ell.fill, { type: 'solid', color: '#FF0000' });
  assert.equal(ell.line.color, '#0000FF');
  near(ell.x, 9 * 914400);
  const tbl = els.find((e) => e.type === 'table');
  assert.ok(tbl, 'table');
  assert.equal(tbl.rows.length, 2);
  assert.equal(tbl.cols.length, 2);
  assert.equal(textOf(tbl.rows[1].cells[1]), 'D');
  assert.equal(doc.slides[0].notes, 'My notes here');
  assert.equal(doc.slides[1].notes, '');
  assert.ok(doc.layouts.length >= 1);
  assert.equal(doc.theme.id, 'imported');
  assert.equal(doc.theme.colors.accent1, '#4472C4');
});

async function sampleDoc() {
  const M = await load('slides/model.js');
  let doc = M.newDoc({ size: '16:9', theme: 'hostwerk', title: 'Deck' });
  doc = M.addSlide(doc, 'title', 0, { ctrTitle: 'Round trip', subTitle: 'Subtitle here' }).doc;
  let r = M.addSlide(doc, 'titleContent', 1, { title: 'Agenda', body: 'One\n\tTwo\nThree', notes: 'Speaker line 1\nline 2' });
  doc = r.doc;
  const s2 = r.slide.id;
  const sh = M.shape('roundRect', 914400, 1828800, 1828800, 914400, { adj: [25000], shadow: { color: '#000000', blur: 50800, dist: 38100, dir: 45, alpha: 40 } });
  sh.text = M.body([M.para('Shape text', { align: 'ctr' }, { b: true, color: 'lt1', size: 20 })], { anchor: 'ctr' });
  const a = M.shape('ellipse', 0, 0, 914400, 914400, { fill: { type: 'solid', color: 'accent2+40' } });
  const b = M.shape('rect', 914400, 914400, 914400, 914400, { fill: { type: 'grad', angle: 90, stops: [{ pos: 0, color: 'accent1' }, { pos: 100, color: '#FF000080' }] } });
  const ln = M.line(100000, 5000000, 3000000, 0, { line: { color: 'tx1', width: 25400, dash: 'dash', head: 'none', tail: 'triangle' } });
  const img = M.image(PNG, 6000000, 1000000, 2000000, 1500000, { crop: { l: 0.1, t: 0, r: 0.2, b: 0 }, opacity: 0.5, alt: 'Pixel' });
  doc = M.addEls(doc, s2, [sh, a, b, ln, img]);
  doc = M.groupEls(doc, s2, [a.id, b.id]).doc;
  r = M.addSlide(doc, 'titleOnly', 2, { title: 'Data' });
  doc = r.doc;
  const s3 = r.slide.id;
  const tb = M.table(3, 3, 914400, 1828800, 5486400, 1371600);
  tb.rows[0].cells[0] = M.cell('Head', { span: { c: 2, r: 1 } });
  tb.rows[0].cells[1] = { ...M.cell(''), merged: true };
  tb.rows[1].cells[2] = M.cell('X', { fill: { type: 'solid', color: '#00FF00' } });
  const ch = M.chart('col', 6400800, 1828800, 4572000, 3429000);
  ch.chart.title = 'Sales';
  doc = M.addEls(doc, s3, [tb, ch]);
  doc = M.updateSlide(doc, s3, (s) => ({ ...s, hidden: true, transition: { type: 'push', dir: 'u', duration: 800, advanceAfter: 3000 } }));
  doc = M.updateSlide(doc, s2, (s) => ({
    ...s,
    transition: { type: 'fade', dir: 'l', duration: 500, advanceAfter: null },
    animations: [{ id: 'a1', el: sh.id, effect: 'fade', trigger: 'click', duration: 500, delay: 0 }, { id: 'a2', el: img.id, effect: 'fly', dir: 'l', trigger: 'after', duration: 700, delay: 0 }],
  }));
  doc = M.addSection(doc, doc.slides[0].id, 'Intro');
  doc = M.addSection(doc, s3, 'Numbers');
  doc = { ...doc, show: { ...doc.show, loop: true } };
  return { doc, ids: { s2, s3, sh: sh.id, img: img.id } };
}

test('writePptx → readPptx round trip', async () => {
  const { writePptx } = await load('slides/io/pptx-write.js');
  const { readPptx } = await load('slides/io/pptx-read.js');
  const { doc } = await sampleDoc();
  const bytes = await writePptx(doc);
  const { doc: back } = await readPptx(bytes, { ext: 'pptx' });
  assert.equal(back.slides.length, 3);
  assert.deepEqual(back.size, doc.size);
  // title slide placeholders
  const t = back.slides[0].elements.find((e) => e.ph === 'ctrTitle');
  assert.equal(textOf(t), 'Round trip');
  assert.equal(textOf(back.slides[0].elements.find((e) => e.ph === 'subTitle')), 'Subtitle here');
  assert.equal(back.layouts.find((l) => l.id === back.slides[0].layoutId).kind, 'title');
  // content slide
  const s2 = back.slides[1];
  assert.equal(back.layouts.find((l) => l.id === s2.layoutId).kind, 'titleContent');
  const body = s2.elements.find((e) => e.ph === 'body');
  assert.equal(textOf(body), 'One\nTwo\nThree');
  assert.equal(body.text.paras[1].level, 1);
  assert.ok(body.text.lvl[0].bullet && body.text.lvl[0].bullet.char, 'bullets inherited from layout');
  assert.equal(s2.notes, 'Speaker line 1\nline 2');
  const sh = s2.elements.find((e) => textOf(e) === 'Shape text');
  assert.equal(sh.geom, 'roundRect');
  assert.deepEqual(sh.adj, [25000]);
  assert.equal(sh.text.paras[0].runs[0].b, true);
  assert.equal(sh.text.paras[0].runs[0].color, 'lt1');
  assert.equal(sh.text.paras[0].align, 'ctr');
  assert.equal(sh.shadow.blur, 50800);
  const g = s2.elements.find((e) => e.type === 'group');
  assert.equal(g.children.length, 2);
  assert.deepEqual(g.children.map((c) => [c.x, c.y, c.w, c.h]), [[0, 0, 914400, 914400], [914400, 914400, 914400, 914400]]);
  assert.deepEqual(g.children[0].fill, { type: 'solid', color: 'accent2+40' });
  assert.equal(g.children[1].fill.type, 'grad');
  assert.equal(g.children[1].fill.angle, 90);
  assert.equal(g.children[1].fill.stops[1].color, '#FF0000@50');
  const ln = s2.elements.find((e) => e.type === 'line');
  assert.equal(ln.line.tail, 'triangle');
  assert.equal(ln.line.dash, 'dash');
  assert.equal(ln.line.width, 25400);
  const img = s2.elements.find((e) => e.type === 'image');
  assert.ok(img.src.startsWith('data:image/png;base64,'));
  assert.equal(img.opacity, 0.5);
  assert.equal(img.crop.r, 0.2);
  assert.equal(img.alt, 'Pixel');
  assert.equal(s2.transition.type, 'fade');
  assert.equal(s2.animations.length, 2);
  assert.equal(s2.animations[0].effect, 'fade');
  assert.equal(s2.animations[0].el, sh.id);
  assert.equal(s2.animations[1].effect, 'fly');
  assert.equal(s2.animations[1].dir, 'l');
  assert.equal(s2.animations[1].trigger, 'after');
  assert.equal(s2.animations[1].el, img.id);
  // data slide
  const s3 = back.slides[2];
  assert.equal(s3.hidden, true);
  assert.equal(s3.transition.type, 'push');
  assert.equal(s3.transition.dir, 'u');
  assert.equal(s3.transition.duration, 800);
  assert.equal(s3.transition.advanceAfter, 3000);
  const tb = s3.elements.find((e) => e.type === 'table');
  assert.equal(tb.rows.length, 3);
  assert.deepEqual(tb.rows[0].cells[0].span, { c: 2, r: 1 });
  assert.equal(tb.rows[0].cells[1].merged, true);
  assert.equal(textOf(tb.rows[0].cells[0]), 'Head');
  assert.deepEqual(tb.rows[1].cells[2].fill, { type: 'solid', color: '#00FF00' });
  assert.equal(tb.style.firstRow, true);
  const ch = s3.elements.find((e) => e.type === 'chart');
  assert.equal(ch.chart.kind, 'col');
  assert.equal(ch.chart.title, 'Sales');
  assert.deepEqual(ch.chart.categories, ['Q1', 'Q2', 'Q3', 'Q4']);
  assert.equal(ch.chart.series.length, 3);
  assert.deepEqual(ch.chart.series[0].values, [4.3, 2.5, 3.5, 4.5]);
  // sections, show, theme
  assert.deepEqual(back.sections.map((s) => s.name), ['Intro', 'Numbers']);
  assert.equal(back.sections[1].start, s3.id);
  assert.equal(back.show.loop, true);
  assert.equal(back.theme.colors.accent1, doc.theme.colors.accent1);
  assert.equal(back.theme.fonts.major, doc.theme.fonts.major);
  assert.equal(back.meta.title, 'Deck');
});

test('written package: every XML part parses, content types cover every part, rels resolve', async () => {
  const JSZip = require('jszip');
  const { parseXml, kids } = await load('_shared/xmltree.js');
  const { writePptx } = await load('slides/io/pptx-write.js');
  const { doc } = await sampleDoc();
  const zip = await JSZip.loadAsync(await writePptx(doc));
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  assert.equal(names[0], '[Content_Types].xml');
  const ct = parseXml(await zip.file('[Content_Types].xml').async('string'));
  const defaults = new Set(kids(ct, 'Default').map((d) => d.a.Extension.toLowerCase()));
  const overrides = new Set(kids(ct, 'Override').map((o) => o.a.PartName));
  for (const n of names) {
    if (n === '[Content_Types].xml') continue;
    const ext = n.split('.').pop().toLowerCase();
    assert.ok(overrides.has(`/${n}`) || defaults.has(ext), `content type for ${n}`);
    if (/\.(xml|rels)$/.test(n)) {
      const s = await zip.file(n).async('string');
      const x = parseXml(s);
      assert.ok(x && x.n, `parses ${n}`);
      if (n.endsWith('.rels')) {
        const base = n === '_rels/.rels' ? '' : n.replace(/_rels\/([^/]+)\.rels$/, '$1');
        for (const r of kids(x, 'Relationship')) {
          if (r.a.TargetMode === 'External') continue;
          const dir = base.includes('/') ? base.slice(0, base.lastIndexOf('/')) : '';
          const parts = dir ? dir.split('/') : [];
          for (const seg of r.a.Target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
          const target = r.a.Target.startsWith('/') ? r.a.Target.slice(1) : parts.join('/');
          assert.ok(zip.file(target), `${n}: ${r.a.Target} exists`);
        }
      }
    }
  }
  for (const o of overrides) assert.ok(zip.file(o.slice(1)), `override part exists ${o}`);
  assert.ok(zip.file('ppt/notesMasters/notesMaster1.xml'));
  assert.ok(zip.file('ppt/charts/chart1.xml'));
  assert.equal(names.filter((n) => /^ppt\/media\//.test(n)).length, 1, 'image stored once');
});

test('round trip keeps original master/layout parts of an imported deck', async () => {
  const PptxGenJS = require('pptxgenjs');
  const JSZip = require('jszip');
  const { readPptx } = await load('slides/io/pptx-read.js');
  const { writePptx } = await load('slides/io/pptx-write.js');
  const p = new PptxGenJS();
  p.defineSlideMaster({ title: 'BRANDED', background: { color: 'F1F1F1' }, objects: [{ rect: { x: 0, y: 0, w: '100%', h: 0.4, fill: { color: '003366' } } }] });
  const s = p.addSlide({ masterName: 'BRANDED' });
  s.addText('Kept', { x: 1, y: 1, w: 4, h: 1 });
  const orig = await p.write({ outputType: 'uint8array' });
  const { doc } = await readPptx(orig, { ext: 'pptx' });
  const out = await writePptx({ ...doc, slides: [...doc.slides, { ...doc.slides[0], id: 'copy', notes: 'added notes' }] }, { original: orig });
  const zip = await JSZip.loadAsync(out);
  const origZip = await JSZip.loadAsync(orig);
  for (const n of Object.keys(origZip.files).filter((x) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(x))) {
    assert.equal(await zip.file(n).async('string'), await origZip.file(n).async('string'), `${n} unchanged`);
  }
  const { doc: back } = await readPptx(out, { ext: 'pptx' });
  assert.equal(back.slides.length, 2);
  assert.equal(back.slides[1].notes, 'added notes');
  assert.ok(back.slides[0].elements.some((e) => textOf(e) === 'Kept'));
  const layout = back.layouts.find((l) => l.id === back.slides[0].layoutId);
  assert.ok(layout.decor.some((d) => d.fill && d.fill.color === '#003366'), 'master decoration read');
});

test('theme colour mods (schemeClr lumMod/lumOff/alpha) become refs', async () => {
  const { colorOf } = await load('slides/io/pptx-read.js');
  const { parseXml } = await load('_shared/xmltree.js');
  const theme = { colors: { accent1: '#4472C4', dk1: '#000000', lt1: '#FFFFFF' } };
  const c = (s) => colorOf(parseXml(`<a:solidFill xmlns:a="x">${s}</a:solidFill>`), { theme });
  assert.equal(c('<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>'), 'accent1+40');
  assert.equal(c('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>'), 'accent1-25');
  assert.equal(c('<a:schemeClr val="tx1"><a:alpha val="50000"/></a:schemeClr>'), 'tx1@50');
  assert.equal(c('<a:srgbClr val="ff8800"/>'), '#FF8800');
  assert.match(c('<a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr>'), /^#[0-9A-F]{6}$/);
  assert.equal(c('<a:sysClr val="windowText" lastClr="000000"/>'), '#000000');
  const { colorXml } = await load('slides/io/pptx-write.js');
  assert.equal(colorXml('accent1+40', theme), '<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>');
  assert.equal(colorXml('accent1-25@50', theme), '<a:schemeClr val="accent1"><a:lumMod val="75000"/><a:alpha val="50000"/></a:schemeClr>');
  assert.equal(colorXml('#112233', theme), '<a:srgbClr val="112233"/>');
});

/* ---------------------------------------------------------------- full fidelity: patching the original package */
async function patchZip(bytes, fn) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(bytes);
  await fn(zip);
  return zip.generateAsync({ type: 'uint8array' });
}
const VBA = new Uint8Array(Array.from({ length: 300 }, (_, i) => (i * 7 + 3) & 255));
async function withMacros(bytes) {
  return patchZip(bytes, async (zip) => {
    zip.file('ppt/vbaProject.bin', VBA);
    let rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    rels = rels.replace('</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
    zip.file('ppt/_rels/presentation.xml.rels', rels);
    let ct = await zip.file('[Content_Types].xml').async('string');
    ct = ct.replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"')
      .replace('application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml');
    zip.file('[Content_Types].xml', ct);
    zip.file('customXml/item1.xml', '<root xmlns="urn:x">custom data</root>');
    let root = await zip.file('_rels/.rels').async('string');
    root = root.replace('</Relationships>', '<Relationship Id="rIdCx" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="customXml/item1.xml"/></Relationships>');
    zip.file('_rels/.rels', root);
  });
}

test('pptm: vbaProject.bin and customXml survive byte-identical, pptx drops the macros', async () => {
  const JSZip = require('jszip');
  const { writePptx } = await load('slides/io/pptx-write.js');
  const { readPptx } = await load('slides/io/pptx-read.js');
  const { doc: d0 } = await sampleDoc();
  const pptm = await withMacros(await writePptx(d0));
  const { doc, fidelity, preserved } = await readPptx(pptm, { ext: 'pptm' });
  assert.deepEqual(fidelity, []);
  assert.ok(preserved.includes('macros'));
  const edited = { ...doc, slides: doc.slides.map((s, i) => (i === 0 ? { ...s, notes: 'changed' } : s)) };
  const out = await JSZip.loadAsync(await writePptx(edited, { original: pptm, ext: 'pptm' }));
  assert.deepEqual(await out.file('ppt/vbaProject.bin').async('uint8array'), VBA);
  assert.equal(await out.file('customXml/item1.xml').async('string'), '<root xmlns="urn:x">custom data</root>');
  const ct = await out.file('[Content_Types].xml').async('string');
  assert.match(ct, /macroEnabled\.main\+xml/);
  assert.match(await out.file('ppt/_rels/presentation.xml.rels').async('string'), /vbaProject/);
  // as plain pptx: macros dropped, custom XML still kept
  const plain = await JSZip.loadAsync(await writePptx(edited, { original: pptm, ext: 'pptx' }));
  assert.equal(plain.file('ppt/vbaProject.bin'), null);
  assert.doesNotMatch(await plain.file('[Content_Types].xml').async('string'), /macroEnabled/);
  assert.ok(plain.file('customXml/item1.xml'));
  // a built-in theme applied: masters are regenerated, unrelated parts still kept
  const M = await load('slides/model.js');
  const themed = M.applyTheme(edited, 'ocean');
  const t = await JSZip.loadAsync(await writePptx(themed, { original: pptm, ext: 'pptm' }));
  assert.deepEqual(await t.file('ppt/vbaProject.bin').async('uint8array'), VBA);
  assert.ok(t.file('customXml/item1.xml'));
  const { doc: back } = await readPptx(await t.generateAsync({ type: 'uint8array' }), { ext: 'pptm' });
  assert.equal(back.slides.length, 3);
  assert.equal(back.theme.colors.accent1, themed.theme.colors.accent1);
  assert.equal(back.slides[0].notes, 'changed');
});

const FRAME = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="77" name="Mystery 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="3000000" cy="1500000"/></p:xfrm><a:graphic><a:graphicData uri="urn:example:mystery"><my:thing xmlns:my="urn:example:mystery" r:id="rIdMy" answer="42"/></a:graphicData></a:graphic></p:graphicFrame>';
async function withMystery(bytes) {
  return patchZip(bytes, async (zip) => {
    let s = await zip.file('ppt/slides/slide2.xml').async('string');
    s = s.replace('</p:spTree>', FRAME + '</p:spTree>');
    zip.file('ppt/slides/slide2.xml', s);
    zip.file('ppt/embeddings/mystery1.bin', new Uint8Array([1, 2, 3, 4]));
    let r = await zip.file('ppt/slides/_rels/slide2.xml.rels').async('string');
    r = r.replace('</Relationships>', '<Relationship Id="rIdMy" Type="http://example.com/mystery" Target="../embeddings/mystery1.bin"/></Relationships>');
    zip.file('ppt/slides/_rels/slide2.xml.rels', r);
  });
}

test('unknown graphicFrame is kept raw (with its part), moving it patches only the xfrm', async () => {
  const JSZip = require('jszip');
  const { writePptx } = await load('slides/io/pptx-write.js');
  const { readPptx } = await load('slides/io/pptx-read.js');
  const { doc: d0 } = await sampleDoc();
  const orig = await withMystery(await writePptx(d0));
  const { doc } = await readPptx(orig, { ext: 'pptx' });
  const el = doc.slides[1].elements.find((e) => e.name === 'Mystery 1');
  assert.ok(el && el.preserved && el.raw && el.spid === 77, 'preserved element');
  assert.equal(el.rawRels.length, 1);
  assert.equal(el.rawRels[0].target, 'ppt/embeddings/mystery1.bin');
  let out = await JSZip.loadAsync(await writePptx(doc, { original: orig }));
  let s2 = await out.file('ppt/slides/slide2.xml').async('string');
  assert.ok(s2.includes('answer="42"') && s2.includes('uri="urn:example:mystery"'));
  assert.ok(s2.includes('<a:off x="1000000" y="2000000"/>'));
  const rid = /<my:thing[^>]*r:id="([^"]+)"/.exec(s2)[1];
  const rels = await out.file('ppt/slides/_rels/slide2.xml.rels').async('string');
  assert.ok(new RegExp('Id="' + rid + '"[^>]*Target="../embeddings/mystery1.bin"').test(rels), 'rel remapped');
  assert.deepEqual(await out.file('ppt/embeddings/mystery1.bin').async('uint8array'), new Uint8Array([1, 2, 3, 4]));
  // moved + slide order changed
  const M = await load('slides/model.js');
  let moved = M.updateEls(doc, doc.slides[1].id, [el.id], (e) => ({ ...e, x: e.x + 500000 }));
  moved = M.moveSlides(moved, [moved.slides[1].id], 0);
  out = await JSZip.loadAsync(await writePptx(moved, { original: orig }));
  s2 = await out.file('ppt/slides/slide1.xml').async('string');
  assert.ok(s2.includes('<a:off x="1500000" y="2000000"/>') && s2.includes('answer="42"'));
  assert.ok(out.file('ppt/embeddings/mystery1.bin'));
});

test('timing XML survives byte-identical while animations are unchanged; edits regenerate it', async () => {
  const JSZip = require('jszip');
  const { writePptx } = await load('slides/io/pptx-write.js');
  const { readPptx } = await load('slides/io/pptx-read.js');
  const { doc: d0 } = await sampleDoc();
  const orig = await patchZip(await writePptx(d0), async (zip) => {
    let s = await zip.file('ppt/slides/slide2.xml').async('string');
    s = s.replace('presetID="10" presetClass="entr"', 'presetID="10" presetClass="entr" custom="1"');
    zip.file('ppt/slides/slide2.xml', s);
  });
  const timingOf = async (bytes, n = 2) => { const z = await JSZip.loadAsync(bytes); const s = await z.file('ppt/slides/slide' + n + '.xml').async('string'); return s.slice(s.indexOf('<p:timing'), s.indexOf('</p:timing>') + 11); };
  const t0 = await timingOf(orig);
  assert.ok(t0.includes('custom="1"'));
  const { doc } = await readPptx(orig, { ext: 'pptx' });
  assert.ok(doc.slides[1].timingRaw && doc.slides[1].animSig);
  const M = await load('slides/model.js');
  const edited = M.addEls(doc, doc.slides[1].id, [M.shape('star5', 0, 0, 500000, 500000)]);
  const out = await writePptx(edited, { original: orig });
  assert.equal(await timingOf(out), t0);
  const { doc: back } = await readPptx(out, { ext: 'pptx' });
  assert.equal(back.slides[1].animations.length, 2);
  assert.equal(back.slides[1].elements.filter((e) => e.geom === 'star5').length, 1);
  const z = await JSZip.loadAsync(out);
  assert.match(await z.file('ppt/slides/slide2.xml').async('string'), /<p:fade\/>/);
  // change an animation → regenerated
  const changed = M.updateSlide(doc, doc.slides[1].id, (s) => ({ ...s, animations: s.animations.slice(0, 1).map((a) => ({ ...a, duration: 900 })) }));
  const t2 = await timingOf(await writePptx(changed, { original: orig }));
  assert.notEqual(t2, t0);
  assert.ok(!t2.includes('custom="1"') && t2.includes('dur="900"'));
  // an unchanged shape keeps its raw XML + id; an edited one is regenerated with the same id
  const s2 = doc.slides[1];
  const sh = s2.elements.find((e) => e.text && e.text.paras.some((p) => p.runs.some((r) => r.text === 'Shape text')));
  assert.ok(sh.raw && sh.spid);
  const ed = M.updateEls(doc, s2.id, [sh.id], (e) => ({ ...e, text: { ...e.text, paras: [{ ...e.text.paras[0], runs: [{ ...e.text.paras[0].runs[0], text: 'New text' }] }] } }));
  const { doc: b2 } = await readPptx(await writePptx(ed, { original: orig }), { ext: 'pptx' });
  const nsh = b2.slides[1].elements.find((e) => e.spid === sh.spid);
  assert.equal(nsh.text.paras[0].runs[0].text, 'New text');
});

test('comments round trip: unchanged part kept, edits written as legacy comments', async () => {
  const JSZip = require('jszip');
  const { writePptx } = await load('slides/io/pptx-write.js');
  const { readPptx } = await load('slides/io/pptx-read.js');
  const { doc: d0 } = await sampleDoc();
  const withC = { ...d0, slides: d0.slides.map((s, i) => (i === 1 ? { ...s, comments: [{ id: 'c1', author: 'Lena Ott', text: 'Check numbers', time: '2026-01-02T03:04:05Z' }] } : s)) };
  const bytes = await writePptx(withC);
  const { doc } = await readPptx(bytes, { ext: 'pptx' });
  assert.equal(doc.slides[1].comments.length, 1);
  assert.equal(doc.slides[1].comments[0].author, 'Lena Ott');
  assert.equal(doc.slides[1].comments[0].text, 'Check numbers');
  const z0 = await JSZip.loadAsync(bytes);
  const part = Object.keys(z0.files).find((n) => /^ppt\/comments\/.+\.xml$/.test(n));
  const again = await JSZip.loadAsync(await writePptx(doc, { original: bytes }));
  assert.equal(await again.file(part).async('string'), await z0.file(part).async('string'));
  const more = { ...doc, slides: doc.slides.map((s, i) => (i === 1 ? { ...s, comments: [...s.comments, { id: 'c2', author: 'Max', text: 'Done' }] } : s)) };
  const { doc: b } = await readPptx(await writePptx(more, { original: bytes }), { ext: 'pptx' });
  assert.deepEqual(b.slides[1].comments.map((c) => c.author + ':' + c.text), ['Lena Ott:Check numbers', 'Max:Done']);
});
