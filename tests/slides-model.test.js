'use strict';
/* Slides chunk – pure logic: units, theme colours, model operations, history, templates, table ops, steps, PDF writer. */
const test = require('node:test');
const assert = require('node:assert/strict');

const S = '../ui/src/editors/slides/';
const load = (p) => import(S + p);

test('units: EMU conversions and ODF lengths', async () => {
  const u = await load('units.js');
  assert.equal(u.inToEmu(1), 914400);
  assert.equal(u.cmToEmu(2.54), 914400);
  assert.equal(u.ptToEmu(1), 12700);
  assert.equal(u.emuToPt(12700 * 18), 18);
  assert.equal(u.lengthToEmu('2.54cm'), 914400);
  assert.equal(u.lengthToEmu('1in'), 914400);
  assert.equal(u.lengthToEmu('72pt'), 914400);
  assert.equal(u.lengthToEmu('25.4mm'), 914400);
  assert.equal(u.emuToLength(914400, 'cm'), '2.54cm');
  assert.equal(u.sizePreset(12192000, 6858000), '16:9');
  assert.equal(u.sizePreset(9144000, 6858000), '4:3');
  assert.equal(u.sizePreset(1234, 5678), 'custom');
  assert.equal(u.degToAng(90), 5400000);
  assert.equal(u.angToDeg(5400000), 90);
});

test('themes: colour refs resolve with lumMod/lumOff, alpha and dark mapping', async () => {
  const th = await load('themes.js');
  const theme = { colors: { ...th.DEFAULT_COLORS, accent1: '#4472C4' } };
  assert.equal(th.resolveHex('accent1', theme), '#4472C4');
  assert.equal(th.resolveHex('#ff0000', theme), '#FF0000');
  // lighter 40 % = lumMod 60 % + lumOff 40 % (PowerPoint's "Lighter 40 %") – #4472C4 → #8FAADC
  assert.equal(th.resolveHex('accent1+40', theme), '#8FAADC');
  // darker 25 % = lumMod 75 % – #4472C4 → #2F5597
  assert.equal(th.resolveHex('accent1-25', theme), '#2F5597');
  assert.equal(th.resolveHex('tx1', theme), '#000000');
  assert.equal(th.resolveHex('tx1', { ...theme, dark: true }), '#FFFFFF');
  assert.equal(th.css('accent1@50', theme), 'rgba(68,114,196,0.5)');
  assert.equal(th.resolveFont('+mj', { fonts: { major: 'Georgia', minor: 'Arial' } }), 'Georgia');
  assert.equal(th.resolveFont('+mn-lt', { fonts: { major: 'Georgia', minor: 'Arial' } }), 'Arial');
  for (const t of th.THEMES) {
    const layouts = th.buildLayouts(t, { w: 12192000, h: 6858000 });
    assert.equal(layouts.length, 8, t.id);
    assert.ok(layouts.find((l) => l.kind === 'title').ph.ctrTitle);
  }
});

test('model: slides add, duplicate, reorder, delete with sections', async () => {
  const m = await load('model.js');
  let d = m.newDoc({ theme: 'office' });
  const ids = [];
  for (const k of ['title', 'titleContent', 'twoContent', 'section']) { const r = m.addSlide(d, k, undefined, { title: k, ctrTitle: k }); d = r.doc; ids.push(r.slide.id); }
  assert.equal(d.slides.length, 4);
  assert.equal(m.slideTitle(d.slides[0]), 'title');
  d = m.moveSlides(d, [ids[3]], 0);
  assert.deepEqual(d.slides.map((s) => s.id), [ids[3], ids[0], ids[1], ids[2]]);
  d = m.moveSlides(d, [ids[3], ids[0]], 4);
  assert.deepEqual(d.slides.map((s) => s.id), [ids[1], ids[2], ids[3], ids[0]]);
  const dup = m.duplicateSlides(d, [ids[1]]);
  assert.equal(dup.doc.slides.length, 5);
  assert.equal(dup.doc.slides[1].id, dup.ids[0]);
  assert.notEqual(dup.doc.slides[1].elements[0].id, dup.doc.slides[0].elements[0].id);
  d = m.addSection(dup.doc, ids[3], 'Part 2');
  const groups = m.sectionGroups(d);
  assert.equal(groups.length, 2);
  assert.equal(groups[1].section.name, 'Part 2');
  d = m.deleteSlides(d, [ids[3]]);
  assert.equal(d.slides.length, 4);
  assert.equal(d.sections[0].start, ids[0], 'section moves to the next slide');
});

test('model: group, ungroup, z-order, align, distribute, resize', async () => {
  const m = await load('model.js');
  let d = m.newDoc({});
  d = m.addSlide(d, 'blank').doc;
  const sid = d.slides[0].id;
  const a = m.shape('rect', 0, 0, 100, 100); const b = m.shape('ellipse', 300, 200, 100, 50); const c = m.shape('rect', 1000, 0, 100, 100);
  d = m.addEls(d, sid, [a, b, c]);
  const g = m.groupEls(d, sid, [a.id, b.id]);
  d = g.doc;
  const grp = d.slides[0].elements.find((e) => e.id === g.id);
  assert.equal(grp.type, 'group');
  assert.deepEqual([grp.x, grp.y, grp.w, grp.h], [0, 0, 400, 250]);
  assert.equal(d.slides[0].elements[0].id, g.id, 'group keeps the z position of its first member');
  d = m.updateEls(d, sid, [g.id], (e) => m.moveEl(e, 10, 20));
  assert.equal(m.findEl(d.slides[0].elements, b.id).x, 310);
  d = m.updateEls(d, sid, [g.id], (e) => m.resizeEl(e, { x: 10, y: 20, w: 800, h: 500 }));
  const bb = m.findEl(d.slides[0].elements, b.id);
  assert.deepEqual([bb.x, bb.y, bb.w, bb.h], [610, 420, 200, 100]);
  const u = m.ungroupEls(d, sid, [g.id]);
  d = u.doc;
  assert.equal(d.slides[0].elements.length, 3);
  d = m.zorder(d, sid, [c.id], 'back');
  assert.equal(d.slides[0].elements[0].id, c.id);
  d = m.zorder(d, sid, [c.id], 'forward');
  assert.equal(d.slides[0].elements[1].id, c.id);
  d = m.alignEls(d, sid, [a.id, c.id], 't');
  assert.equal(m.findEl(d.slides[0].elements, a.id).y, m.findEl(d.slides[0].elements, c.id).y);
  d = m.distributeEls(d, sid, [a.id, b.id, c.id], 'h');
  const xs = [a, b, c].map((e) => m.findEl(d.slides[0].elements, e.id)).sort((p, q) => p.x - q.x);
  assert.equal(Math.round(xs[1].x - (xs[0].x + xs[0].w)), Math.round(xs[2].x - (xs[1].x + xs[1].w)));
  const rb = m.rotatedBox({ x: 0, y: 0, w: 200, h: 100, rot: 90 });
  assert.deepEqual([Math.round(rb.x), Math.round(rb.y), Math.round(rb.w), Math.round(rb.h)], [50, -50, 100, 200]);
});

test('model: undo/redo history with merging keys', async () => {
  const m = await load('model.js');
  let d = m.newDoc({});
  let h = m.history(d);
  d = m.addSlide(d, 'title').doc;
  h = m.histPush(h, d, null, 1000);
  const d2 = m.updateSlide(d, d.slides[0].id, { notes: 'a' });
  h = m.histPush(h, d2, 'notes', 2000);
  const d3 = m.updateSlide(d2, d.slides[0].id, { notes: 'ab' });
  h = m.histPush(h, d3, 'notes', 2500); // merged
  assert.equal(h.past.length, 2);
  h = m.histUndo(h);
  assert.equal(h.doc.slides[0].notes, '');
  h = m.histRedo(h);
  assert.equal(h.doc.slides[0].notes, 'ab');
  h = m.histUndo(m.histUndo(h));
  assert.equal(h.doc.slides.length, 0);
  assert.equal(m.histUndo(h), h, 'nothing more to undo');
  const d4 = m.addSlide(h.doc, 'blank').doc;
  h = m.histPush(h, d4, null, 9000);
  assert.equal(h.future.length, 0, 'a new change clears redo');
});

test('model: theme switch keeps text, size change scales', async () => {
  const m = await load('model.js');
  let d = m.newDoc({ theme: 'hostwerk' });
  d = m.addSlide(d, 'titleContent', undefined, { title: 'Hello', body: 'One\n\tTwo' }).doc;
  const body = d.slides[0].elements.find((e) => e.ph === 'body');
  assert.equal(body.text.paras[1].level, 1);
  d = m.applyTheme(d, 'midnight');
  assert.equal(d.theme.id, 'midnight');
  assert.equal(m.slideTitle(d.slides[0]), 'Hello');
  const t0 = d.slides[0].elements.find((e) => e.ph === 'title');
  d = m.setSize(d, { w: 9144000, h: 6858000 }, 'fit');
  const t1 = d.slides[0].elements.find((e) => e.ph === 'title');
  assert.ok(t1.w < t0.w);
  assert.equal(d.size.w, 9144000);
  d = m.resetSlide(d, d.slides[0].id, 'twoContent');
  assert.ok(d.slides[0].elements.some((e) => e.ph === 'body2'));
  assert.ok(m.docToText(d).includes('Hello'));
});

test('templates: every template builds and survives pptx and odp', async () => {
  const T = await load('templates.js');
  const io = await load('io/index.js');
  for (const id of ['blank', 'pitch', 'lesson', 'photo']) {
    for (const lang of ['en', 'de']) {
      const d = T.buildTemplate(id, { lang });
      assert.ok(d.slides.length >= 1);
      for (const fmt of ['pptx', 'odp']) {
        const bytes = await io.writeAny(d, fmt);
        const r = await io.readAny(bytes, fmt);
        assert.equal(r.doc.slides.length, d.slides.length, `${id} ${fmt}`);
      }
    }
  }
  assert.equal(T.buildTemplate('pitch', { lang: 'en' }).slides.length, 8);
  assert.equal(T.buildTemplate('lesson', { lang: 'de' }).slides.length, 6);
  assert.equal(io.saveFormat('pptm'), 'pptm');
  assert.equal(io.saveFormat('ppt'), 'pptx');
  assert.equal(io.saveFormat('fodp'), 'fodp'); // 1.1: flat ODF is written as flat ODF
});

test('table operations: insert, delete, merge, split', async () => {
  const m = await load('model.js');
  const { tableOp } = await load('table.js');
  let t = m.table(2, 3, 0, 0, 3000, 2000);
  t = tableOp(t, 'rowBelow', { r: 0, c: 0 });
  assert.equal(t.rows.length, 3);
  t = tableOp(t, 'colRight', { r: 0, c: 2 });
  assert.equal(t.cols.length, 4);
  assert.ok(t.rows.every((r) => r.cells.length === 4));
  t = tableOp(t, 'mergeRight', { r: 0, c: 0 });
  t = tableOp(t, 'mergeDown', { r: 0, c: 0 });
  assert.deepEqual(t.rows[0].cells[0].span, { c: 2, r: 2 });
  assert.equal(t.rows[1].cells[1].merged, true);
  t = tableOp(t, 'split', { r: 1, c: 1 });
  assert.equal(t.rows[0].cells[0].span, undefined);
  assert.equal(t.rows[1].cells[1].merged, false);
  t = tableOp(t, 'delCol', { r: 0, c: 0 });
  assert.equal(t.cols.length, 3);
  assert.equal(tableOp(m.table(1, 1, 0, 0, 10, 10), 'delRow', { r: 0, c: 0 }), null);
});

test('slide show steps: click groups with with/after', async () => {
  const { buildSteps } = await load('anim.js');
  const slide = { elements: [{ id: 'a' }, { id: 'b' }, { id: 'c' }, { id: 'd' }], animations: [
    { el: 'a', effect: 'fade', trigger: 'with', duration: 400 },
    { el: 'b', effect: 'fly', trigger: 'click', duration: 500 },
    { el: 'c', effect: 'fade', trigger: 'after', duration: 300 },
    { el: 'd', effect: 'appear', trigger: 'with', duration: 1, delay: 100 },
    { el: 'zz', effect: 'fade', trigger: 'click' },
  ] };
  const g = buildSteps(slide);
  assert.equal(g.length, 2);
  assert.equal(g[0].auto, true);
  assert.equal(g[1].items[1].delay, 500, 'after previous waits for the fly-in');
  assert.equal(g[1].items[2].delay, 600, 'with previous starts with c (+100 ms delay)');
});

test('geometry: every preset produces a finite path', async () => {
  const g = await load('render/geometry.js');
  for (const k of Object.keys(g.SHAPES)) {
    const p = g.shapePath(k, 120, 80);
    assert.ok(p.d && !/NaN|Infinity/.test(p.d), k);
    assert.ok(/^[MLCZ0-9 .\-]+$/.test(p.d), `${k} uses only M/L/C/Z`);
  }
  assert.equal(g.resolvePrst('round2SameRect'), 'roundRect');
  assert.equal(g.resolvePrst('nonsense'), 'rect');
  assert.equal(g.ODF_TO_GEOM['right-arrow'], 'rightArrow');
});

test('image PDF writer produces a parseable PDF', async () => {
  const { imagePdf } = await load('io/export.js');
  // 1×1 white JPEG
  const jpeg = Uint8Array.from(Buffer.from('/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAAA//EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AN//Z', 'base64'));
  const pdf = imagePdf([{ jpeg, pxW: 1, pxH: 1 }, { jpeg, pxW: 1, pxH: 1 }], 960, 540, 'Test (deck)');
  const { PDFDocument } = require('pdf-lib');
  const doc = await PDFDocument.load(pdf);
  assert.equal(doc.getPageCount(), 2);
  const { width, height } = doc.getPage(0).getSize();
  assert.equal(Math.round(width), 960);
  assert.equal(Math.round(height), 540);
});
