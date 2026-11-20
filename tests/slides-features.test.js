'use strict';
/* Office 1.1 features of the slides chunk (pure parts): node --test tests/slides-features.test.js
 * presentation from a document, Designer, compare versions, live fields, co-editing state, QR code, guides,
 * slide-show animation states. */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const SRC = path.join(__dirname, '..', 'ui', 'src', 'editors');
const load = (p) => import(require('url').pathToFileURL(path.join(SRC, p)).href);

test('presentation from a document: Markdown, docx and odt outlines become slides', async () => {
  const f = await load('slides/features.js');
  const m = await load('slides/model.js');
  const md = '# Quarterly report\nA short intro sentence. More detail follows here.\n\n## Results\n- Revenue up 12 %\n- Costs flat\n  - Energy down\n\n## Next steps\nWe will hire two people. Then we launch.\n';
  const o = f.outlineFromMd(md);
  assert.deepEqual(o.map((x) => x.kind), ['h', 'p', 'h', 'li', 'li', 'li', 'h', 'p']);
  const plans = f.planSlides(o);
  assert.equal(plans[0].kind, 'title');
  assert.equal(plans[0].title, 'Quarterly report');
  assert.equal(plans[0].subtitle, 'A short intro sentence.');
  const res = plans.find((p) => p.title === 'Results');
  assert.deepEqual(res.bullets, [{ text: 'Revenue up 12 %', level: 0 }, { text: 'Costs flat', level: 0 }, { text: 'Energy down', level: 1 }]);
  assert.deepEqual(plans.find((p) => p.title === 'Next steps').bullets, [{ text: 'We will hire two people.', level: 0 }]);
  const doc = m.newDoc({});
  const slides = f.slidesFromPlans(doc, plans);
  assert.equal(slides.length, 3);
  const body = slides[1].elements.find((e) => e.ph && /^body/.test(e.ph));
  assert.equal(body.text.paras[2].level, 1);
  // docx
  const docXml = '<w:document xmlns:w="w"><w:body><w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>My deck</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Intro</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>Point </w:t></w:r><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>Plain text.</w:t></w:r></w:p></w:body></w:document>';
  const od = f.outlineFromDocxXml(docXml);
  assert.deepEqual(od, [{ kind: 'title', level: 0, text: 'My deck' }, { kind: 'h', level: 1, text: 'Intro' }, { kind: 'li', level: 0, text: 'Point one' }, { kind: 'p', level: 0, text: 'Plain text.' }]);
  // odt
  const odt = '<office:document-content xmlns:office="o" xmlns:text="t"><office:body><office:text><text:h text:outline-level="1">Kapitel</text:h><text:list><text:list-item><text:p>Erster Punkt</text:p></text:list-item></text:list><text:p>Absatz</text:p></office:text></office:body></office:document-content>';
  assert.deepEqual(f.outlineFromOdtXml(odt).map((x) => [x.kind, x.text]), [['h', 'Kapitel'], ['li', 'Erster Punkt'], ['p', 'Absatz']]);
  // AI answer
  const ai = f.plansFromAi('Here you go:\n[{"title":"Deck","subtitle":"Sub"},{"title":"Point","bullets":["a","b"],"notes":"Say hi"}]');
  assert.equal(ai[0].kind, 'title');
  assert.deepEqual(ai[1].bullets.map((b) => b.text), ['a', 'b']);
  assert.equal(ai[1].notes, 'Say hi');
  assert.equal(f.plansFromAi('no json'), null);
});

test('Designer: arrangements for pictures and points stay inside the slide', async () => {
  const f = await load('slides/features.js');
  const m = await load('slides/model.js');
  let d = m.newDoc({});
  d = m.addSlide(d, 'titleContent', undefined, { title: 'Our plan', body: 'Plan\nBuild\nTest\nShip' }).doc;
  const ideas = f.designIdeas(d, d.slides[0]);
  assert.ok(ideas.length >= 3, `ideas: ${ideas.map((x) => x.id)}`);
  for (const i of ideas) for (const e of i.elements) { assert.ok(e.x >= -1 && e.y >= -1 && e.x + e.w <= d.size.w + 2 && e.y + e.h <= d.size.h + 2, `${i.id}: element inside`); }
  assert.ok(ideas.some((i) => i.id === 'cards' && i.elements.filter((e) => /^Card/.test(e.name)).length === 4));
  // a picture
  const img = m.image('data:image/png;base64,AAAA', 0, 0, 4000000, 3000000);
  d = m.addEls(d, d.slides[0].id, [img]);
  const pi = f.designIdeas(d, d.slides[0]);
  assert.ok(pi.some((i) => i.id === 'pic-right'));
  const pr = pi.find((i) => i.id === 'pic-right').elements.find((e) => e.type === 'image');
  assert.ok(pr.crop.t > 0 || pr.crop.l > 0 || (pr.crop.t === 0 && pr.crop.l === 0));
  assert.equal(pi[0].elements.filter((e) => e.type === 'image').length, 1);
});

test('compare versions: changed, added, removed and moved slides', async () => {
  const f = await load('slides/features.js');
  const m = await load('slides/model.js');
  let a = m.newDoc({});
  a = m.addSlide(a, 'titleContent', undefined, { title: 'One', body: 'alpha' }).doc;
  a = m.addSlide(a, 'titleContent', undefined, { title: 'Two', body: 'beta' }).doc;
  a = m.addSlide(a, 'titleContent', undefined, { title: 'Three', body: 'gamma' }).doc;
  a = m.addSlide(a, 'titleContent', undefined, { title: 'Four', body: 'delta' }).doc;
  let b = JSON.parse(JSON.stringify(a));
  // change slide Two's body, delete Three, add Five, move Four before One
  const two = b.slides[1];
  const body = two.elements.find((e) => e.ph === 'body');
  body.text.paras[0].runs[0].text = 'beta changed';
  b.slides = [b.slides[3], b.slides[0], b.slides[1]];
  b = m.addSlide(b, 'titleContent', undefined, { title: 'Five', body: 'epsilon' }).doc;
  const rows = f.compareDecks(a, b);
  const by = (st) => rows.filter((r) => r.status === st);
  assert.equal(by('removed').length, 1);
  assert.equal(by('removed')[0].a, 2);
  assert.equal(by('added').length, 1);
  assert.equal(b.slides[by('added')[0].b].elements.length > 0, true);
  const ch = by('changed');
  assert.equal(ch.length, 1);
  assert.deepEqual(ch[0].changes.changed, [body.id]);
  assert.ok(by('moved').length + by('same').length === 2);
});

test('live fields: runs keep id/params, values refresh, pptx keeps the last value', async () => {
  const f = await load('slides/features.js');
  const m = await load('slides/model.js');
  const io = await load('slides/io/index.js');
  let d = m.newDoc({});
  d = m.addSlide(d, 'blank').doc;
  const box = f.liveFieldBox(d, { id: 'minecraft.players', params: { server: 'mc1' }, label: 'Players online' }, 3);
  d = m.addEls(d, d.slides[0].id, [box]);
  const runs = f.liveRuns(d);
  assert.equal(runs.length, 1);
  assert.equal(runs[0].key, 'minecraft.players|{"server":"mc1"}');
  const d2 = f.applyLiveValues(d, { [runs[0].key]: 7 });
  assert.notEqual(d2, d);
  assert.equal(f.applyLiveValues(d2, { [runs[0].key]: 7 }), d2);
  const bytes = await io.writeAny(d2, 'pptx');
  const JSZip = require('jszip');
  const xml = await (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml').async('string');
  assert.match(xml, /<a:fld id="\{[^"]+\}" type="hostwerk\.live\?id=minecraft\.players&amp;params=[^"]*"><a:rPr[^>]*\/><a:t>7<\/a:t><\/a:fld>/);
  const r = await io.readAny(bytes, 'pptx');
  const run = r.doc.slides[0].elements[0].text.paras[0].runs[0];
  assert.deepEqual(run.live, { id: 'minecraft.players', params: { server: 'mc1' }, label: 'Players online' });
  assert.equal(run.text, '7');
});

test('co-editing: element-level shared state converges between two editors', async () => {
  const Y = require('yjs');
  const c = await load('slides/collab.js');
  const m = await load('slides/model.js');
  let d = m.newDoc({});
  d = m.addSlide(d, 'titleContent', undefined, { title: 'Hello' }).doc;
  const a = new Y.Doc(); const b = new Y.Doc();
  a.on('update', (u) => Y.applyUpdate(b, u)); b.on('update', (u) => Y.applyUpdate(a, u));
  c.seed(null, a, d);
  const fromB = c.readShared(b);
  assert.equal(fromB.slides.length, 1);
  assert.equal(fromB.slides[0].elements.length, d.slides[0].elements.length);
  // editor B moves an element, editor A adds a slide – both see both
  const el = fromB.slides[0].elements[0];
  const bDoc = m.updateEls(fromB, fromB.slides[0].id, [el.id], (x) => ({ ...x, x: x.x + 1000 }));
  let pushed = 0; b.on('update', () => pushed++);
  c.pushDiff(b, fromB, bDoc);
  assert.equal(pushed, 1);
  const aDoc = m.addSlide(c.readShared(a), 'blank').doc;
  c.pushDiff(a, c.readShared(a), aDoc);
  const final = c.readShared(b);
  assert.equal(final.slides.length, 2);
  assert.equal(final.slides[0].elements[0].x, el.x + 1000);
  assert.deepEqual(c.readShared(a).slides.map((s) => s.id), final.slides.map((s) => s.id));
  // unchanged doc: nothing is sent
  let n = 0; a.on('update', () => n++);
  c.pushDiff(a, final, final);
  assert.equal(n, 0);
});

test('QR code: valid sizes and the remote URL encodes', async () => {
  const { qrMatrix, qrSvg } = await load('slides/qr.js');
  assert.equal(qrMatrix('HELLO').length, 21);
  assert.equal(qrMatrix('https://hw.local:8443/api/office/remote/0123456789abcdef').length, 33);
  assert.match(qrSvg('x'), /^<svg [^>]*viewBox="0 0 29 29"/);
  assert.throws(() => qrMatrix('x'.repeat(400)));
});

test('guides round-trip through pptx (p15:sldGuideLst)', async () => {
  const m = await load('slides/model.js');
  const io = await load('slides/io/index.js');
  let d = m.newDoc({});
  d = m.addSlide(d, 'blank').doc;
  d = { ...d, guides: [{ id: 'a', axis: 'x', pos: 6096000 }, { id: 'b', axis: 'y', pos: 1587500 }] };
  const bytes = await io.writeAny(d, 'pptx');
  const JSZip = require('jszip');
  const pres = await (await JSZip.loadAsync(bytes)).file('ppt/presentation.xml').async('string');
  assert.match(pres, /<p15:guide id="1" pos="3840">/);
  assert.match(pres, /<p15:guide id="2" orient="horz" pos="1000">/);
  const r = await io.readAny(bytes, 'pptx');
  assert.deepEqual(r.doc.guides.map((g) => [g.axis, g.pos]), [['x', 6096000], ['y', 1587500]]);
});

test('slide show states: exit hides, emphasis keeps, motion paths move', async () => {
  const a = await load('slides/anim.js');
  const slide = { elements: [{ id: 'x' }, { id: 'y' }], animations: [
    { id: '1', el: 'x', cls: 'entr', effect: 'fade', trigger: 'click', duration: 500 },
    { id: '2', el: 'y', cls: 'path', effect: 'line', trigger: 'click', duration: 1000, path: 'M 0 0 L 0.5 0 E' },
    { id: '3', el: 'x', cls: 'exit', effect: 'fade', trigger: 'click', duration: 500 },
    { id: '4', el: 'y', cls: 'emph', effect: 'grow', trigger: 'with', duration: 500, scale: 2 },
  ] };
  const g = a.buildSteps(slide);
  assert.equal(g.length, 3);
  const at = (step, playing = false) => a.animDecor(g, step, playing, true, { W: 1000, H: 500 });
  assert.equal(at(0)({ id: 'x' }).style.visibility, 'hidden');
  assert.equal(at(0)({ id: 'y' }), null);
  assert.equal(at(1)({ id: 'x' }), null);
  assert.equal(at(2)({ id: 'y' }).style.translate, '500.0px 0.0px');
  assert.equal(at(3)({ id: 'x' }).style.visibility, 'hidden');
  assert.equal(at(3)({ id: 'y' }).style.scale, '2');
  const jobs = a.motionJobs(g, 2, { W: 1000, H: 500 });
  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].frames[jobs[0].frames.length - 1].translate, '500.0px 0.0px');
  assert.match(at(3, true)({ id: 'x' }).style.animation, /^of-slides-fade 500ms .* reverse both$/);
});
