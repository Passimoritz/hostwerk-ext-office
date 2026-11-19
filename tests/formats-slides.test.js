/* Regressions of the Microsoft Office check (PowerPoint 16) for the pptx writer: findings #3, #4, #5. */
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { buildTemplate } from '../ui/src/editors/slides/templates.js';
import { writePptx } from '../ui/src/editors/slides/io/pptx-write.js';
import { readPptx } from '../ui/src/editors/slides/io/pptx-read.js';

const phs = (xml) => [...xml.matchAll(/<p:ph\b([^>]*?)\/?>/g)].map((m) => ({ type: (/type="([^"]+)"/.exec(m[1]) || [])[1] || 'obj', idx: (/idx="([^"]+)"/.exec(m[1]) || [])[1] || '0' }));

test('Office #3: master placeholders are typed (title/body/dt/ftr/sldNum), layouts match their slides', async () => {
  const d = buildTemplate('pitch', { lang: 'en' });
  const z = await JSZip.loadAsync(await writePptx(d));
  const master = await z.file('ppt/slideMasters/slideMaster1.xml').async('string');
  assert.deepEqual(phs(master), [{ type: 'title', idx: '0' }, { type: 'body', idx: '1' }, { type: 'dt', idx: '2' }, { type: 'ftr', idx: '3' }, { type: 'sldNum', idx: '4' }]);
  const rels = (f) => z.file(f.replace(/([^/]+)$/, '_rels/$1.rels')).async('string');
  for (const f of Object.keys(z.files).filter((x) => /^ppt\/slides\/slide\d+\.xml$/.test(x))) {
    const slidePh = phs(await z.file(f).async('string'));
    const layoutTarget = /Target="\.\.\/slideLayouts\/([^"]+)"/.exec(await rels(f))[1];
    const layoutPh = phs(await z.file(`ppt/slideLayouts/${layoutTarget}`).async('string'));
    // every placeholder of the slide finds its layout placeholder (same idx, and the same type where the slide names one)
    for (const p of slidePh) {
      const hit = layoutPh.find((l) => l.idx === p.idx && (p.type === 'obj' || l.type === p.type || (p.type === 'body' && l.type === 'obj')));
      assert.ok(hit, `${f}: ${p.type}/${p.idx} not on ${layoutTarget} (${JSON.stringify(layoutPh)})`);
    }
  }
});

test('Office #4: macro and show/template formats get their own main content type', async () => {
  const d = buildTemplate('blank', { lang: 'en' });
  const want = {
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
    potx: 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
    pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
    ppsm: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
    potm: 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
  };
  for (const [ext, type] of Object.entries(want)) {
    const z = await JSZip.loadAsync(await writePptx(d, { ext }));
    const ct = await z.file('[Content_Types].xml').async('string');
    assert.ok(ct.includes(`PartName="/ppt/presentation.xml" ContentType="${type}"`), ext);
  }
});

test('Office #5: a failed patch never falls back silently – PPTX_PATCH_FAILED with regenerate()', async () => {
  const d = buildTemplate('blank', { lang: 'en' });
  const good = await writePptx(d);
  const { doc } = await readPptx(good, { ext: 'pptx' });
  // an "original" that is no usable package any more
  const z = await JSZip.loadAsync(good);
  z.remove('ppt/presentation.xml');
  const broken = await z.generateAsync({ type: 'uint8array' });
  let err = null;
  try { await writePptx(doc, { original: broken, ext: 'pptx' }); } catch (e) { err = e; }
  assert.ok(err, 'must throw');
  assert.equal(err.code, 'PPTX_PATCH_FAILED');
  assert.equal(typeof err.regenerate, 'function');
  assert.match(err.message, /save a copy/);
  const fresh = await err.regenerate();
  const back = await readPptx(fresh, { ext: 'pptx' });
  assert.equal(back.doc.slides.length, doc.slides.length);
  // the unbroken original still patches
  const ok = await writePptx(doc, { original: good, ext: 'pptx' });
  assert.ok(ok.length > 0);
});

test('Office #5: imported slide backgrounds keep their theme colour index (bgRef n)', async () => {
  const d = buildTemplate('blank', { lang: 'en' });
  const z = await JSZip.loadAsync(await writePptx(d));
  const slide = await z.file('ppt/slides/slide1.xml').async('string');
  const bg = slide.replace('<p:cSld>', '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg2"/></p:bgRef></p:bg>');
  z.file('ppt/slides/slide1.xml', bg);
  const orig = await z.generateAsync({ type: 'uint8array' });
  const { doc } = await readPptx(orig, { ext: 'pptx' });
  const edited = { ...doc, slides: doc.slides.map((s) => ({ ...s, notes: 'touched' })) };
  const out = await JSZip.loadAsync(await writePptx(edited, { original: orig, ext: 'pptx' }));
  const s1 = await out.file('ppt/slides/slide1.xml').async('string');
  assert.ok(/<p:bg>/.test(s1), 'background kept');
  assert.doesNotMatch(s1, /undefined|NaN/);
});
