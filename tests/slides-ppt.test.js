'use strict';
/* PowerPoint 97–2003 reader: a synthetic .ppt built in the test (tiny CFB writer + PPT records). */
const test = require('node:test');
const assert = require('node:assert');

const load = () => import('../ui/src/editors/slides/io/ppt-read.js');

/* ---------------- records */
const u16 = (v) => { const b = Buffer.alloc(2); b.writeUInt16LE(v & 0xffff); return b; };
const i16 = (v) => { const b = Buffer.alloc(2); b.writeInt16LE(v); return b; };
const u32 = (v) => { const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0); return b; };
const i32 = (v) => { const b = Buffer.alloc(4); b.writeInt32LE(v); return b; };
function rec(type, body, inst = 0, ver = 0) { const b = Buffer.concat(body ? [].concat(body) : []); return Buffer.concat([u16((inst << 4) | ver), u16(type), u32(b.length), b]); }
const box = (type, kids, inst = 0) => rec(type, kids, inst, 0xf);
const utf16 = (s) => Buffer.from(s, 'utf16le');

const textChars = (s) => rec(0x0fa0, [utf16(s)]);
const textHeader = (t) => rec(0x0f9f, [u32(t)]);
/** bold + size + centred style for a text of length n */
const styleBoldCentered = (n, size) => rec(0x0fa1, [u32(n + 1), u16(0), u32(0x800), u16(1), u32(n + 1), u32(0x1 | 0x20000), u16(1), u16(size)]);
const slidePersist = (ref, slideId) => rec(0x03f3, [u32(ref), u32(0), i32(0), u32(slideId), u32(0)]);
const slideAtom = (geom, masterId, notesId, flags) => rec(0x03ef, [u32(geom), Buffer.alloc(8), u32(masterId), u32(notesId), u16(flags), u16(0)]);
const fsp = (spt, spid, flags) => rec(0xf00a, [u32(spid), u32(flags)], spt, 2);
const fopt = (props) => rec(0xf00b, props.map(([id, v]) => Buffer.concat([u16(id), i32(v)])), props.length, 3);
const anchor = (t, l, r, b) => rec(0xf010, [i16(t), i16(l), i16(r), i16(b)]);
const sp = (kids) => box(0xf004, kids);
function drawing(shapes) {
  const patriarch = sp([rec(0xf009, [i32(0), i32(0), i32(0), i32(0)], 0, 1), fsp(0, 1024, 0x5)]);
  return box(0x040c, [box(0xf002, [rec(0xf008, [u32(3), u32(1026)]), box(0xf003, [patriarch, ...shapes])])]);
}

function buildPpt({ blank2007 = false } = {}) {
  const W = 5760; const H = 4320;
  const docAtom = rec(0x03e9, [i32(W), i32(H), i32(H), i32(W), i32(1), i32(2), u32(0), u32(0), u16(1), u16(0), Buffer.alloc(4)], 0, 1);
  const fontEntity = rec(0x0fb7, [Buffer.concat([utf16('Arial'), Buffer.alloc(64 - 10)]), Buffer.alloc(4)], 0);
  const env = box(0x03f2, [box(0x07d6, [fontEntity])]);
  const t1 = 'Quarterly review';
  const slwt0 = box(0x0ff0, [
    slidePersist(3, 256), textHeader(6), textChars(t1), styleBoldCentered(t1.length, 40), textHeader(5), textChars('Second line\rThird line'),
    slidePersist(4, 257), textHeader(0), textChars('Agenda'), textHeader(1), textChars('Point one\rPoint two'),
  ], 0);
  const slwt1 = box(0x0ff0, [slidePersist(2, 0x80000000)], 1);
  const slwt2 = box(0x0ff0, [slidePersist(5, 300)], 2);
  const fbse = rec(0xf007, [Buffer.from([6, 6]), Buffer.alloc(16), u16(0xff), u32(PNG.length + 25), u32(1), u32(0), Buffer.from([0, 0, 0, 0])], 6, 2);
  const dgg = box(0x040b, [box(0xf000, [box(0xf001, [fbse], 1)])]);
  const docC = box(0x03e8, [docAtom, env, dgg, slwt1, slwt0, slwt2]);
  const scheme = Buffer.concat(['FFFFFF', '000000', '808080', '1F3864', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000'].map((h) => Buffer.from(h + '00', 'hex')));
  const masterC = box(0x03f8, [slideAtom(1, 0, 0, 0), rec(0x07f0, [scheme], 1), drawing([])]);
  // slide 1: title placeholder (text from the slide list), red rectangle, a text box with its own text
  const titlePh = sp([fsp(1, 1025, 0xa00), anchor(500, 400, 5360, 1500), box(0xf011, [rec(0x0bc3, [u32(0), Buffer.from([15, 0]), u16(0)])]), box(0xf00d, [rec(0x0f9e, [i32(0)])])]);
  const rect = sp([fsp(1, 1026, 0xa00), fopt([[0x0181, 0x000000ff], [0x01bf, 0x00100010], [0x01c0, 0x08000001], [0x01cb, 12700]]), anchor(2000, 576, 1728, 2864)]);
  const tbox = sp([fsp(202, 1027, 0xa00), fopt([[0x01bf, 0x00100000]]), anchor(3000, 2000, 5000, 3400), box(0xf00d, [textHeader(4), textChars('Hello box')])]);
  const pic = sp([fsp(75, 1029, 0xa00), fopt([[0x4104, 1]]), anchor(3600, 100, 600, 4200)]);
  const slide1 = box(0x03ee, [slideAtom(0, 0x80000000, 300, 7), drawing([titlePh, rect, tbox, pic])]);
  // slide 2: no shapes at all – texts only in the slide list, hidden + fade transition
  const ss = rec(0x03f9, [i32(3000), u32(0), Buffer.from([0, 6]), u16(0x404), Buffer.from([2, 0, 0, 0])]);
  // blank2007: PowerPoint 2007+ writes geom 16 (SL_Blank) for every slide – the layout comes from the placeholders
  const titleOnly = sp([fsp(1, 1030, 0xa00), anchor(300, 300, 5400, 1200), box(0xf011, [rec(0x0bc3, [u32(0), Buffer.from([0x0d, 0]), u16(0)])]), box(0xf00d, [rec(0x0f9e, [i32(0)])])]);
  const slide2 = blank2007 ? box(0x03ee, [slideAtom(16, 0x80000000, 0, 7), ss, drawing([titleOnly])]) : box(0x03ee, [slideAtom(1, 0x80000000, 0, 7), ss, drawing([])]);
  const notesBody = sp([fsp(1, 1028, 0xa00), anchor(100, 100, 5000, 4000), box(0xf011, [rec(0x0bc3, [u32(1), Buffer.from([12, 0]), u16(0)])]), box(0xf00d, [textHeader(2), textChars('Speaker notes one')])]);
  // Office #10: the slide number placeholder of the notes page shows "*" – it is not part of the notes
  const notesNum = sp([fsp(1, 1031, 0xa00), anchor(4100, 3000, 5000, 4300), box(0xf011, [rec(0x0bc3, [u32(5), Buffer.from([0x08, 0]), u16(0)])]), box(0xf00d, [textHeader(4), textChars('*')])]);
  const notesImg = sp([fsp(1, 1032, 0xa00), anchor(100, 100, 5000, 2000), box(0xf011, [rec(0x0bc3, [u32(0), Buffer.from([0x0b, 0]), u16(0)])])]);
  const notesC = box(0x03f0, [rec(0x03f1, [u32(256), u16(0), u16(0)], 0, 1), drawing([notesImg, notesBody, notesNum])]);
  const parts = [docC, masterC, slide1, slide2, notesC];
  const offsets = []; let pos = 0;
  for (const p of parts) { offsets.push(pos); pos += p.length; }
  const persistOff = pos;
  const persistDir = rec(0x1772, [u32(1 | (5 << 20)), ...offsets.map(u32)]);
  const userEditOff = persistOff + persistDir.length;
  const userEdit = rec(0x0ff5, [u32(256), u16(0), Buffer.from([0, 3]), u32(0), u32(persistOff), u32(1), u32(6), u16(1), u16(0)]);
  const stream = Buffer.concat([...parts, persistDir, userEdit]);
  const cu = rec(0x0ff6, [u32(0x14), u32(0xe391c05f), u32(userEditOff), u16(0), u16(0x3f4), Buffer.from([3, 0]), u16(0)]);
  return { stream, cu };
}

const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8cfc0f01f0005000201a3c2e0a30000000049454e44ae426082', 'hex');
const picturesStream = () => rec(0xf01e, [Buffer.alloc(16), Buffer.from([0xff]), PNG], 0x6e0, 0);

/* ---------------- minimal CFB writer (512-byte sectors, no mini stream) */
function cfb(streams) {
  const SEC = 512; const END = 0xfffffffe; const FREE = 0xffffffff;
  const pad = (b) => { const n = Math.max(4096, Math.ceil(b.length / SEC) * SEC); const o = Buffer.alloc(n); b.copy(o); return o; };
  const data = streams.map(([name, b]) => ({ name, b: pad(b) }));
  const fat = [0xfffffffd, END];
  let next = 2;
  for (const d of data) { d.start = next; const n = d.b.length / SEC; for (let k = 0; k < n; k++) fat.push(k === n - 1 ? END : next + k + 1); next += n; }
  while (fat.length < SEC / 4) fat.push(FREE);
  const header = Buffer.alloc(SEC);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(header, 0);
  header.writeUInt16LE(0x3e, 24); header.writeUInt16LE(3, 26); header.writeUInt16LE(0xfffe, 28); header.writeUInt16LE(9, 30); header.writeUInt16LE(6, 32);
  header.writeUInt32LE(1, 44); header.writeUInt32LE(1, 48); header.writeUInt32LE(4096, 56); header.writeUInt32LE(END, 60); header.writeUInt32LE(0, 64); header.writeUInt32LE(END, 68); header.writeUInt32LE(0, 72);
  for (let k = 0; k < 109; k++) header.writeUInt32LE(k === 0 ? 0 : FREE, 76 + k * 4);
  const fatSec = Buffer.alloc(SEC); fat.forEach((v, k) => fatSec.writeUInt32LE(v >>> 0, k * 4));
  const dir = Buffer.alloc(SEC);
  const entry = (i, name, type, start, size, child, right) => {
    const o = i * 128; const nm = utf16(name); nm.copy(dir, o); dir.writeUInt16LE(nm.length + 2, o + 64); dir[o + 66] = type; dir[o + 67] = 1;
    dir.writeUInt32LE(FREE, o + 68); dir.writeUInt32LE(right >>> 0, o + 72); dir.writeUInt32LE(child >>> 0, o + 76); dir.writeUInt32LE(start >>> 0, o + 116); dir.writeUInt32LE(size, o + 120);
  };
  entry(0, 'Root Entry', 5, END, 0, 1, FREE);
  data.forEach((d, k) => entry(k + 1, d.name, 2, d.start, d.b.length, FREE, k + 2 <= data.length ? k + 2 : FREE));
  return new Uint8Array(Buffer.concat([header, fatSec, dir, ...data.map((d) => d.b)]));
}

test('ppt: synthetic file – slides, texts, shapes, notes, transitions', async () => {
  const { readPpt } = await load();
  const { stream, cu } = buildPpt();
  const bytes = cfb([['Current User', cu], ['PowerPoint Document', stream], ['Pictures', picturesStream()]]);
  const { doc } = await readPpt(bytes, { ext: 'ppt' });
  assert.strictEqual(doc.slides.length, 2);
  assert.strictEqual(doc.size.w, Math.round(5760 * 1587.5));
  assert.strictEqual(doc.size.h, Math.round(4320 * 1587.5));
  assert.strictEqual(doc.theme.id, 'imported');
  assert.strictEqual(doc.theme.colors.accent1, '#4472C4');
  assert.strictEqual(doc.theme.fonts.minor, 'Arial');

  const s1 = doc.slides[0];
  const text = (e) => e.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
  const title = s1.elements.find((e) => e.ph === 'ctrTitle');
  assert.ok(title, 'title placeholder');
  assert.strictEqual(text(title), 'Quarterly review');
  assert.strictEqual(title.text.paras[0].runs[0].b, true);
  assert.strictEqual(title.text.paras[0].runs[0].size, 40);
  assert.strictEqual(title.text.paras[0].align, 'ctr');
  assert.strictEqual(title.x, Math.round(400 * 1587.5));
  const rect = s1.elements.find((e) => e.type === 'shape' && e.geom === 'rect' && !e.textbox);
  assert.ok(rect, 'rectangle');
  assert.deepStrictEqual([rect.x, rect.y, rect.w, rect.h], [576, 2000, 1152, 864].map((v) => Math.round(v * 1587.5)));
  assert.deepStrictEqual(rect.fill, { type: 'solid', color: '#FF0000' });
  assert.strictEqual(rect.line.color, 'tx1');
  assert.strictEqual(rect.line.width, 12700);
  const tb = s1.elements.find((e) => e.textbox && !e.ph);
  assert.strictEqual(text(tb), 'Hello box');
  assert.strictEqual(tb.fill, null);
  // subtitle text from the slide list (not placed by a shape)
  const sub = s1.elements.find((e) => e.ph === 'subTitle');
  assert.strictEqual(text(sub), 'Second line\nThird line');
  assert.strictEqual(s1.notes, 'Speaker notes one');
  const img = s1.elements.find((e) => e.type === 'image');
  assert.ok(img, 'picture');
  assert.strictEqual(img.src, 'data:image/png;base64,' + PNG.toString('base64'));
  assert.strictEqual(img.y, Math.round(3600 * 1587.5));

  const s2 = doc.slides[1];
  assert.strictEqual(text(s2.elements.find((e) => e.ph === 'title')), 'Agenda');
  assert.strictEqual(text(s2.elements.find((e) => e.ph === 'body')), 'Point one\nPoint two');
  assert.strictEqual(s2.hidden, true);
  assert.strictEqual(s2.transition.type, 'fade');
  assert.strictEqual(s2.transition.advanceAfter, 3000);
  assert.strictEqual(s2.transition.duration, 500);
});

test('ppt: garbage in the PowerPoint Document stream does not throw', async () => {
  const { readPpt } = await load();
  let seed = 7;
  const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed & 0xff; };
  const junk = Buffer.alloc(20000); for (let i = 0; i < junk.length; i++) junk[i] = rnd();
  const cu = rec(0x0ff6, [u32(0x14), u32(0xe391c05f), u32(1234), u16(0)]);
  const { doc } = await readPpt(cfb([['Current User', cu], ['PowerPoint Document', junk]]));
  assert.ok(doc && Array.isArray(doc.slides));
  // truncated real file
  const { stream } = buildPpt();
  const half = stream.subarray(0, Math.floor(stream.length / 2));
  const r2 = await readPpt(cfb([['Current User', cu], ['PowerPoint Document', half]]));
  assert.ok(Array.isArray(r2.doc.slides));
});

test('ppt: a file without the PowerPoint stream is rejected with a clear error', async () => {
  const { readPpt } = await load();
  await assert.rejects(() => readPpt(cfb([['WordDocument', Buffer.alloc(10)]])), /PowerPoint/);
});

test('Office #10: notes are only the notes-body placeholder; 2007+ blank geometry takes the layout from the placeholders', async () => {
  const { readPpt } = await load();
  const { stream, cu } = buildPpt({ blank2007: true });
  const { doc } = await readPpt(cfb([['Current User', cu], ['PowerPoint Document', stream], ['Pictures', picturesStream()]]), { ext: 'ppt' });
  assert.strictEqual(doc.slides[0].notes, 'Speaker notes one');
  const lay = (s) => doc.layouts.find((l) => l.id === s.layoutId);
  assert.strictEqual(lay(doc.slides[1]).kind, 'titleOnly');
  assert.strictEqual(doc.slides[1].transition.type, 'fade');
});
