import test from 'node:test';
import assert from 'node:assert/strict';
import { decompressRtf, compressRtfStored, rtfToHtml, PREBUF_LENGTH } from '../ui/src/editors/viewer/rtf.js';
import { metafileToSvg, isEmf, isWmf } from '../ui/src/editors/viewer/emf.js';

const H = (s) => Uint8Array.from(s.trim().split(/\s+/).map((x) => parseInt(x, 16)));

test('LZFu: dictionary prefix has the length of the spec (207)', () => {
  assert.equal(PREBUF_LENGTH, 207);
});

test('LZFu: example of MS-OXRTFCP (with back references)', () => {
  const comp = H('2d 00 00 00 2b 00 00 00 4c 5a 46 75 f1 c5 c7 a7 03 00 0a 00 72 63 70 67 31 32 35 42 32 0a f3 20 68 65 6c 09 00 20 62 77 05 b0 6c 64 7d 0a 80 0f a0');
  assert.equal(new TextDecoder().decode(decompressRtf(comp)), '{\\rtf1\\ansi\\ansicpg1252\\pard hello world}\r\n');
});

test('LZFu: literal stream and MELA (stored) round trip', () => {
  for (const s of ['{\\rtf1 x}', '{\\rtf1\\ansi exactly8}', 'a'.repeat(5000)]) {
    const b = new TextEncoder().encode(s);
    assert.equal(new TextDecoder().decode(decompressRtf(compressRtfStored(b))), s);
  }
  const mela = new Uint8Array(16 + 3); const dv = new DataView(mela.buffer);
  dv.setUint32(0, 15, true); dv.setUint32(4, 3, true); dv.setUint32(8, 0x414c454d, true); mela.set([65, 66, 67], 16);
  assert.equal(new TextDecoder().decode(decompressRtf(mela)), 'ABC');
});

test('RTF → HTML: formatting, colours, unicode, code page, links', () => {
  const rtf = '{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}{\\f1\\fcharset204 Arial Cyr;}}{\\colortbl ;\\red255\\green0\\blue0;}'
    + '\\pard Hallo \\b fett\\b0 , \\i kursiv\\i0  und \\cf1 rot\\cf0 .\\par Gr\\\'fc\\\'dfe \\u8364?  {\\f1 \\\'c4\\\'e0}\\par'
    + '{\\field{\\*\\fldinst{HYPERLINK "https://example.org"}}{\\fldrslt{Beispiel}}}\\par}';
  const r = rtfToHtml(rtf);
  assert.match(r.html, /<span style="font-weight:bold">fett<\/span>/);
  assert.match(r.html, /<span style="font-style:italic">kursiv<\/span>/);
  assert.match(r.html, /color:rgb\(255,0,0\)">rot/);
  assert.match(r.text, /Grüße €/);
  assert.match(r.text, /Да/);
  assert.match(r.html, /<a href="https:\/\/example.org">Beispiel<\/a>/);
  assert.ok(!/Arial/.test(r.text), 'font table not shown');
});

test('RTF with encapsulated HTML (\\fromhtml1) gives the original HTML back', () => {
  const rtf = '{\\rtf1\\ansi\\fbidis\\ansicpg1252\\deff0\\fromhtml1 {\\fonttbl{\\f0\\fswiss Arial;}}'
    + '{\\*\\htmltag19 <html>}{\\*\\htmltag34 <body>}\\htmlrtf {\\htmlrtf0 '
    + '{\\*\\htmltag64 <p>}Hallo \\htmlrtf\\b\\htmlrtf0{\\*\\htmltag84 <b>}Welt{\\*\\htmltag92 </b>}\\htmlrtf0 \\htmlrtf\\par\\htmlrtf0{\\*\\htmltag72 </p>}'
    + '\\htmlrtf }\\htmlrtf0 {\\*\\htmltag42 </body>}{\\*\\htmltag27 </html>}}';
  const r = rtfToHtml(rtf);
  assert.ok(r.fromHtml);
  assert.match(r.html, /<p>Hallo <b>Welt<\/b>/);
  assert.match(r.html, /<\/html>/);
});

/* ---------------- EMF / WMF ---------------- */
function emf(records) {
  const head = new Uint8Array(108); const dv = new DataView(head.buffer);
  dv.setUint32(0, 1, true); dv.setUint32(4, 108, true);
  [0, 0, 199, 99].forEach((v, i) => dv.setInt32(8 + i * 4, v, true));
  head.set([0x20, 0x45, 0x4d, 0x46], 40);
  const body = [head, ...records, rec(14, [0, 0, 20])];
  const n = body.reduce((a, b) => a + b.length, 0); const out = new Uint8Array(n); let o = 0;
  for (const b of body) { out.set(b, o); o += b.length; }
  return out;
}
function rec(type, ints, extra) {
  const size = 8 + ints.length * 4 + (extra ? extra.length : 0);
  const b = new Uint8Array((size + 3) & ~3); const dv = new DataView(b.buffer);
  dv.setUint32(0, type, true); dv.setUint32(4, b.length, true);
  ints.forEach((v, i) => dv.setInt32(8 + i * 4, v, true));
  if (extra) b.set(extra, 8 + ints.length * 4);
  return b;
}
test('EMF → SVG: pen, brush, rectangle, polyline, text', () => {
  const text = 'Hi';
  const t16 = new Uint8Array(text.length * 2); text.split('').forEach((c, i) => { t16[i * 2] = c.charCodeAt(0); });
  // EXTTEXTOUTW: bounds(16) mode(4) exScale(4) eyScale(4) EMRTEXT: ref x,y n offString options rect(16) offDx
  const offString = 8 + 16 + 12 + 40;
  const textRec = rec(84, [0, 0, 0, 0, 1, 0, 0, 10, 50, text.length, offString, 0, 0, 0, 0, 0, 0], t16);
  const data = emf([
    rec(38, [1, 0, 3, 0, 0x0000ff]), // CREATEPEN #1 solid width 3 red
    rec(37, [1]),
    rec(39, [2, 0, 0x00ff00, 0]), // CREATEBRUSHINDIRECT #2 green
    rec(37, [2]),
    rec(43, [10, 10, 90, 60]), // RECTANGLE
    rec(87, [0, 0, 0, 0, 3, (0) | (0 << 16), (100) | (50 << 16), (150) | (10 << 16)]), // POLYLINE16
    textRec,
  ]);
  assert.ok(isEmf(data));
  const r = metafileToSvg(data);
  assert.match(r.svg, /<rect x="10" y="10" width="80" height="50" fill="#00ff00" stroke="#ff0000" stroke-width="3"/);
  assert.match(r.svg, /<path d="M0 0L100 50L150 10" fill="none"/);
  assert.match(r.svg, />Hi<\/text>/);
  assert.equal(r.width, 200);
});

test('WMF (placeable) → SVG', () => {
  const words = [];
  const push16 = (...v) => v.forEach((x) => words.push(x & 0xffff));
  const recW = (fn, params) => { push16(((3 + params.length) & 0xffff), ((3 + params.length) >> 16), fn, ...params); };
  // placeable header (22 bytes) is written separately
  push16(1, 9, 0x0300, 0, 0, 0, 0, 0, 0); // META header (9 words; sizes not checked)
  recW(0x020c, [100, 200]); // SETWINDOWEXT (y, x)
  recW(0x02fc, [0, 0x00ff, 0x0000, 0]); // CREATEBRUSHINDIRECT style 0, color 0x0000ff (red)
  recW(0x012d, [0]);
  recW(0x041b, [80, 150, 20, 10]); // RECTANGLE bottom right top left
  recW(0x0000, []);
  const body = new Uint8Array(words.length * 2); const dv = new DataView(body.buffer);
  words.forEach((w, i) => dv.setUint16(i * 2, w, true));
  const head = new Uint8Array(22); const hv = new DataView(head.buffer);
  hv.setUint32(0, 0x9ac6cdd7, true); hv.setInt16(10, 200, true); hv.setInt16(12, 100, true); hv.setUint16(14, 1440, true);
  const data = new Uint8Array(22 + body.length); data.set(head); data.set(body, 22);
  assert.ok(isWmf(data));
  const r = metafileToSvg(data);
  assert.match(r.svg, /<rect x="10" y="20" width="140" height="60" fill="#ff0000"/);
});
