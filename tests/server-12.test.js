'use strict';
/* Office 1.2 server work – pure logic (no Hostwerk needed): form dates, .xls text, ZIP64, Yjs log compaction,
 * report recipients, wrong-password answers, deleted paragraph marks in "clean". */
const test = require('node:test');
const assert = require('node:assert');
const { makeZip, readZip } = require('../server/zip');
const ZP = require('../server/zippatch');
const S = require('../server/sheet');
const FX = require('./server-fixtures');

/* ---------------- 1. form time stamps as real date/time cells ---------------- */
test('sheet: date values become real dates – xlsx serial + date style, ods date-value, csv ISO', () => {
  const stamp = { $date: '2026-11-20T14:03:05' };
  const day = { $date: '2026-03-01' };
  assert.deepStrictEqual(S.dateValue(stamp), { y: 2026, mo: 11, d: 20, h: 14, mi: 3, s: 5, time: true });
  assert.strictEqual(S.dateValue({ $date: '2026-13-01' }), null);
  assert.strictEqual(S.dateText(stamp), '2026-11-20 14:03:05');
  assert.strictEqual(S.dateText(day), '2026-03-01');
  assert.match(S.localStamp(new Date(2026, 0, 2, 3, 4, 5)).$date, /^2026-01-02T03:04:05$/);
  // Excel serials: 1900 system (1899-12-30 = 0) and 1904 system
  assert.strictEqual(S.excelSerial(S.dateValue({ $date: '1900-03-01' })), 61);
  assert.strictEqual(Math.round(S.excelSerial(S.dateValue(stamp)) * 1e6) / 1e6, Math.round((46346 + (14 * 3600 + 3 * 60 + 5) / 86400) * 1e6) / 1e6);
  assert.strictEqual(S.excelSerial(S.dateValue({ $date: '1904-01-02' }), true), 1);

  // xlsx without styles.xml: a minimal one is added (relationship + content type), cells get s=…
  const x = FX.xlsx([['Zeit', 'Name'], ['x', 'y']]);
  const out = S.appendRow(x, 'xlsx', { values: [stamp, 'Anna', day] });
  const z = readZip(out.data);
  const sheet = z.get('xl/worksheets/sheet1.xml').read().toString();
  const styles = z.get('xl/styles.xml').read().toString();
  const xfs = [...styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].matchAll(/<xf\b[^>]*\/>/g)].map((m) => m[0]);
  const cA = /<c r="A4" s="(\d+)"><v>([\d.]+)<\/v><\/c>/.exec(sheet); const cC = /<c r="C4" s="(\d+)"><v>([\d.]+)<\/v><\/c>/.exec(sheet);
  assert.ok(cA && cC, sheet);
  assert.match(xfs[Number(cA[1])], /numFmtId="22"/);
  assert.match(xfs[Number(cC[1])], /numFmtId="14"/);
  assert.strictEqual(Number(cC[2]), 46082);
  assert.ok(sheet.includes('<cols><col min="1" max="1" width="17" customWidth="1"/><col min="3" max="3" width="11" customWidth="1"/></cols><sheetData>'), 'date columns wide enough – no #####');
  assert.match(z.get('[Content_Types].xml').read().toString(), /\/xl\/styles\.xml/);
  assert.match(z.get('xl/_rels/workbook.xml.rels').read().toString(), /Target="styles\.xml"/);
  assert.strictEqual(S.cellValue(S.readBook(out.data, 'xlsx'), '', 'B4'), 'Anna');
  // a second answer reuses the styles (no new xf)
  const out2 = S.appendRow(out.data, 'xlsx', { values: [stamp, 'Bernd'] });
  const styles2 = readZip(out2.data).get('xl/styles.xml').read().toString();
  assert.strictEqual(styles2, styles);
  assert.strictEqual((readZip(out2.data).get('xl/worksheets/sheet1.xml').read().toString().match(/<col /g) || []).length, 2);

  // xlsx with styles.xml and the 1904 date system
  const withStyles = ZP.patchZip(x, {
    'xl/styles.xml': '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="4" fontId="1" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>',
    'xl/_rels/workbook.xml.rels': readZip(x).get('xl/_rels/workbook.xml.rels').read().toString().replace('</Relationships>', '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
    'xl/workbook.xml': readZip(x).get('xl/workbook.xml').read().toString().replace('<sheets>', '<workbookPr date1904="1"/><sheets>'),
  });
  const o3 = S.appendRow(withStyles, 'xlsx', { values: [{ $date: '1904-01-03' }] });
  const z3 = readZip(o3.data);
  assert.match(z3.get('xl/styles.xml').read().toString(), /<cellXfs count="3">.*<xf numFmtId="14"[^>]*\/><\/cellXfs>/);
  assert.match(z3.get('xl/worksheets/sheet1.xml').read().toString(), /<c r="A4" s="2"><v>2<\/v><\/c>/);

  // ods: office:date-value with a date style in the automatic styles
  const content = '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="t" xmlns:text="x" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:automatic-styles/><office:body><office:spreadsheet><table:table table:name="Antworten"><table:table-row><table:table-cell office:value-type="string"><text:p>Zeit</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>';
  const ods = makeZip([{ name: 'mimetype', data: 'application/vnd.oasis.opendocument.spreadsheet', store: true }, { name: 'content.xml', data: content }]);
  const o4 = S.appendRow(ods, 'ods', { sheet: 'Antworten', values: [stamp, day, 'Text'] });
  const c4 = readZip(o4.data).get('content.xml').read().toString();
  assert.match(c4, /<number:date-style style:name="NOfDateTime">[\s\S]*<number:seconds/);
  assert.match(c4, /table:style-name="ceOfDateTime" office:value-type="date" office:date-value="2026-11-20T14:03:05"><text:p>20\.11\.2026 14:03:05<\/text:p>/);
  assert.match(c4, /table:style-name="ceOfDate" office:value-type="date" office:date-value="2026-03-01"/);
  assert.match(c4, /xmlns:number=/);
  assert.strictEqual(S.cellValue(S.readBook(o4.data, 'ods'), 'Antworten', 'A2'), '2026-11-20T14:03:05');
  // a second append reuses the styles
  const c5 = readZip(S.appendRow(o4.data, 'ods', { sheet: 'Antworten', values: [stamp] }).data).get('content.xml').read().toString();
  assert.strictEqual((c5.match(/style:name="ceOfDateTime"/g) || []).length, 1);

  // csv: ISO text Excel/LibreOffice recognise as a date
  const csv = S.appendRow(Buffer.from('Zeit;Name\r\n'), 'csv', { values: [stamp, 'Anna'] });
  assert.strictEqual(csv.data.toString(), 'Zeit;Name\r\n2026-11-20 14:03:05;Anna\r\n');
});

/* ---------------- 8. .xls (BIFF8) text for the search index ---------------- */
const rec = (type, ...parts) => { const d = Buffer.concat(parts.map((p) => (Buffer.isBuffer(p) ? p : Buffer.from(p)))); const h = Buffer.alloc(4); h.writeUInt16LE(type, 0); h.writeUInt16LE(d.length, 2); return Buffer.concat([h, d]); };
const u16 = (v) => { const b = Buffer.alloc(2); b.writeUInt16LE(v); return b; };
const u32 = (v) => { const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0); return b; };
const f64 = (v) => { const b = Buffer.alloc(8); b.writeDoubleLE(v); return b; };
const cell = (r, c) => Buffer.concat([u16(r), u16(c), u16(15)]);
/** a BIFF8 workbook: globals with an SST split over CONTINUE records, two sheets with every cell record kind */
function biff8() {
  const long = 'Überweisung März – ' + 'ä'.repeat(30); // the CONTINUE splits inside the characters, 8-bit → 16-bit
  // SST: "Kunde" (8-bit), long (first 10 chars 8-bit in the SST record, rest UTF-16 in CONTINUE with flags 1),
  // rich "Fett" with 1 run – its run array goes into a second CONTINUE (no flags byte there)
  const sst1 = Buffer.concat([u32(3), u32(3), u16(5), Buffer.from([0]), Buffer.from('Kunde', 'latin1'), u16(long.length), Buffer.from([0]), Buffer.from(long.slice(0, 10), 'latin1')]);
  const cont1 = Buffer.concat([Buffer.from([1]), Buffer.from(long.slice(10), 'utf16le'), u16(4), Buffer.from([0x08]), u16(1), Buffer.from('Fett', 'latin1')]);
  const cont2 = Buffer.concat([u16(0), u16(1)]);
  const bs = (pos, name) => rec(0x0085, u32(pos), Buffer.from([0, 0]), Buffer.from([name.length, 1]), Buffer.from(name, 'utf16le'));
  const head = rec(0x0809, u16(0x0600), u16(5), Buffer.alloc(12));
  const build = (p1, p2) => Buffer.concat([head, bs(p1, 'Rechnungen'), bs(p2, 'Notizen'), rec(0x00fc, sst1), rec(0x003c, cont1), rec(0x003c, cont2), rec(0x000a)]);
  const g0 = build(0, 0);
  const sheet1 = Buffer.concat([rec(0x0809, u16(0x0600), u16(0x10), Buffer.alloc(12)),
    rec(0x00fd, cell(0, 0), u32(0)), rec(0x00fd, cell(0, 1), u32(1)),
    rec(0x0203, cell(1, 0), f64(1234.5)),
    rec(0x027e, cell(1, 1), u32(((42 << 2) | 2) >>> 0)), // RK integer 42
    rec(0x00bd, u16(2), u16(0), u16(15), u32(((7 << 2) | 3) >>> 0), u16(15), u32(((250 << 2) | 2) >>> 0), u16(1)), // 0.07 and 250
    rec(0x0006, cell(3, 0), Buffer.from([0, 0, 0, 0, 0, 0, 0xff, 0xff]), u16(0), u32(0), u16(0)), rec(0x0207, u16(9), Buffer.from([0]), Buffer.from('Ergebnis!', 'latin1')),
    rec(0x0205, cell(3, 1), Buffer.from([1, 0])),
    rec(0x00fd, cell(4, 0), u32(2)),
    rec(0x000a)]);
  const sheet2 = Buffer.concat([rec(0x0809, u16(0x0600), u16(0x10), Buffer.alloc(12)), rec(0x0204, cell(0, 0), u16(6), Buffer.from([1]), Buffer.from('Notiz✓', 'utf16le')), rec(0x000a)]);
  const g = build(g0.length, g0.length + sheet1.length);
  return { wb: Buffer.concat([g, sheet1, sheet2]), long };
}
test('xls: BIFF8 text – SST over CONTINUE records (flags byte, rich runs), LABELSST, LABEL, numbers, RK, MULRK, formula strings', async () => {
  const cfb = require('../server/cfb');
  const X = require('../server/extract');
  const { wb, long } = biff8();
  const buf = FX.cfb({ Workbook: wb });
  const r = cfb.xlsText(buf);
  assert.strictEqual(r.sheets, 2);
  const lines = r.text.split('\n');
  assert.deepStrictEqual(lines.slice(0, 6), ['Rechnungen', `Kunde\t${long}`, '1234.5\t42', '0.07\t250', 'Ergebnis!\tTRUE', 'Fett']);
  assert.ok(r.text.includes('Notizen\nNotiz✓'));
  assert.ok(X.canText('xls'));
  assert.match(await X.text(buf, 'xls'), /Überweisung März/);
  assert.strictEqual((await X.info(buf, 'xls')).sheets, 2);
  // encrypted workbooks (FILEPASS) give no text, no crash in the index
  const enc = FX.cfb({ Workbook: Buffer.concat([rec(0x0809, u16(0x0600), u16(5), Buffer.alloc(12)), rec(0x002f, u16(1)), rec(0x000a)]) });
  assert.throws(() => cfb.xlsText(enc), /encrypted/);
  // a workbook written by SheetJS (BIFF8 without SST) reads too
  const XLSX = require('xlsx');
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Größe', 'Preis'], ['Brot', 2.5], ['Käse €', true]]), 'Daten');
  const t2 = cfb.xlsText(XLSX.write(book, { type: 'buffer', bookType: 'biff8' })).text;
  assert.strictEqual(t2, 'Daten\nGröße\tPreis\nBrot\t2.5\nKäse €\tTRUE');
});

/* ---------------- 7. ZIP64 ---------------- */
/** a small archive written the way streaming writers do it: every size/offset in the ZIP64 extra field, ZIP64 end records */
function zip64Always(files) {
  const zlib = require('zlib');
  const { crc32 } = require('../server/zip');
  const b64 = (n) => { const b = Buffer.alloc(8); b.writeBigUInt64LE(BigInt(n)); return b; };
  const parts = []; const central = []; let off = 0;
  for (const [name, text] of Object.entries(files)) {
    const data = Buffer.from(text); const raw = zlib.deflateRawSync(data); const nb = Buffer.from(name);
    const lx = Buffer.concat([Buffer.from([1, 0, 16, 0]), b64(data.length), b64(raw.length)]);
    const h = Buffer.alloc(30); h.writeUInt32LE(0x04034b50, 0); h.writeUInt16LE(45, 4); h.writeUInt16LE(0x0800, 6); h.writeUInt16LE(8, 8); h.writeUInt32LE(crc32(data), 14); h.writeUInt32LE(0xffffffff, 18); h.writeUInt32LE(0xffffffff, 22); h.writeUInt16LE(nb.length, 26); h.writeUInt16LE(lx.length, 28);
    const cx = Buffer.concat([Buffer.from([1, 0, 24, 0]), b64(data.length), b64(raw.length), b64(off)]);
    const c = Buffer.alloc(46); c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(45, 4); c.writeUInt16LE(45, 6); c.writeUInt16LE(0x0800, 8); c.writeUInt16LE(8, 10); c.writeUInt32LE(crc32(data), 16); c.writeUInt32LE(0xffffffff, 20); c.writeUInt32LE(0xffffffff, 24); c.writeUInt16LE(nb.length, 28); c.writeUInt16LE(cx.length, 30); c.writeUInt32LE(0xffffffff, 42);
    parts.push(h, nb, lx, raw); central.push(Buffer.concat([c, nb, cx]));
    off += 30 + nb.length + lx.length + raw.length;
  }
  const cd = Buffer.concat(central);
  // ZIP64 end record + locator by hand (the classic end record says "see ZIP64")
  const r =Buffer.alloc(56); r.writeUInt32LE(0x06064b50, 0); r.writeBigUInt64LE(44n, 4); r.writeUInt16LE(45, 12); r.writeUInt16LE(45, 14); r.writeBigUInt64LE(BigInt(central.length), 24); r.writeBigUInt64LE(BigInt(central.length), 32); r.writeBigUInt64LE(BigInt(cd.length), 40); r.writeBigUInt64LE(BigInt(off), 48);
  const loc = Buffer.alloc(20); loc.writeUInt32LE(0x07064b50, 0); loc.writeBigUInt64LE(BigInt(off + cd.length), 8); loc.writeUInt32LE(1, 16);
  const e = Buffer.alloc(22); e.writeUInt32LE(0x06054b50, 0); e.writeUInt16LE(0xffff, 8); e.writeUInt16LE(0xffff, 10); e.writeUInt32LE(0xffffffff, 12); e.writeUInt32LE(0xffffffff, 16);
  return Buffer.concat([...parts, cd, r, loc, e]);
}
test('zip64: archives with ZIP64 records are read (readZip, zip patch, sheet, extract), patched and stay readable', async () => {
  const z64 = zip64Always({ 'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Hallo ZIP64</w:t></w:r></w:p></w:body></w:document>', 'docProps/app.xml': '<Properties><Pages>1</Pages></Properties>' });
  const z = readZip(z64);
  assert.strictEqual(z.get('word/document.xml').read().toString().includes('Hallo ZIP64'), true);
  assert.strictEqual(ZP.entries(z64).length, 2);
  const X = require('../server/extract');
  assert.match(await X.text(z64, 'docx'), /Hallo ZIP64/);
  // patching keeps the raw entries and writes a classic zip again (small)
  const out = ZP.patchZip(z64, { 'docProps/app.xml': '<Properties/>' });
  assert.strictEqual(readZip(out).get('word/document.xml').read().toString().includes('Hallo ZIP64'), true);
  assert.ok(!out.includes(Buffer.from([0x50, 0x4b, 0x06, 0x06])), 'no ZIP64 end record for a small archive');
});
test('zip64: the writer switches to ZIP64 records only when needed (> 65535 entries, offsets ≥ 4 GB)', async (t) => {
  const { ZipWriter, centralHeader, localHeader, endRecords, MAX32 } = require('../server/zip');
  // small: classic
  const small = makeZip({ 'a.txt': 'x' });
  assert.ok(!small.includes(Buffer.from([0x50, 0x4b, 0x06, 0x06])));
  // 70 000 entries → ZIP64 end records, every entry still readable (own reader, zip patch, and Hostwerk's yauzl)
  const w = new ZipWriter();
  for (let i = 0; i < 70000; i++) w.add(`f/${i}.txt`, String(i), { store: true });
  const big = w.end();
  assert.ok(big.includes(Buffer.from([0x50, 0x4b, 0x06, 0x06])), 'ZIP64 end record');
  const m = readZip(big);
  assert.strictEqual(m.size, 70000);
  assert.strictEqual(m.get('f/69999.txt').read().toString(), '69999');
  assert.strictEqual(ZP.entries(ZP.patchZip(big, { 'f/1.txt': 'eins' })).length, 70000);
  // headers for entries/offsets beyond 4 GB (no 4 GB buffer needed): extra field 0x0001 with exactly the big values
  const e = { nameBuf: Buffer.from('big.bin'), flags: 0x0800, method: 0, time: 0, date: 0, crc: 1, csize: MAX32 + 10, size: MAX32 + 10 };
  const lh = localHeader(e);
  assert.strictEqual(lh.readUInt32LE(18), MAX32); assert.strictEqual(lh.readUInt16LE(28), 20); assert.strictEqual(Number(lh.readBigUInt64LE(30 + 7 + 4)), MAX32 + 10);
  const ch = centralHeader({ ...e, csize: 5, size: 5 }, MAX32 + 99);
  assert.strictEqual(ch.readUInt32LE(42), MAX32); assert.strictEqual(ch.readUInt16LE(30), 12); assert.strictEqual(Number(ch.readBigUInt64LE(46 + 7 + 4)), MAX32 + 99);
  const er = endRecords(3, 100, MAX32 + 5);
  assert.strictEqual(er.length, 56 + 20 + 22); assert.strictEqual(Number(er.readBigUInt64LE(48)), MAX32 + 5);
  // Hostwerk's own reader agrees
  const dirs = [process.env.HOSTWERK_DIR, require('path').join(__dirname, '..', '..', '..', 'hostwerk')].filter(Boolean);
  let yauzl = null; for (const d of dirs) { try { yauzl = require(require.resolve('yauzl', { paths: [d] })); break; } catch { /* next */ } }
  if (!yauzl) { t.diagnostic('yauzl cross-check skipped'); return; }
  const n = await new Promise((ok, bad) => yauzl.fromBuffer(big, { lazyEntries: true }, (er2, zf) => { if (er2) return bad(er2); let k = 0; zf.on('entry', () => { k++; zf.readEntry(); }); zf.on('end', () => ok(k)); zf.on('error', bad); zf.readEntry(); }));
  assert.strictEqual(n, 70000);
});

/* ---------------- 6. co-editing room: the Yjs update log is compacted ---------------- */
test('collab: a long session is compacted (deleted text gone), late joiners get the same state, no reset', (t) => {
  let Y, enc;
  try { Y = require('yjs'); enc = require('lib0/encoding'); } catch { t.skip('yjs not installed'); return; }
  const C = require('../server/collab');
  const relay = new C.Collab({ logBytes: 2 * 1048576 });
  const sent = [];
  const conn = relay.connect('Lang.docx', { send: (b) => sent.push(b), close: () => {} }, { user: 'Anna' });
  const doc = new Y.Doc(); const text = doc.getText('t');
  doc.on('update', (u) => { const e = enc.createEncoder(); enc.writeVarUint(e, 0); enc.writeVarUint(e, 2); enc.writeVarUint8Array(e, u); conn.message(Buffer.from(enc.toUint8Array(e))); });
  // ~ 6 MB of typing and deleting again (more than the 2 MB cap without compaction)
  const block = 'Lorem ipsum dolor sit amet, '.repeat(200);
  let total = 0;
  for (let i = 0; i < 1000; i++) { text.insert(text.length, block); text.delete(0, block.length - 10); total += block.length; }
  const room = conn.room;
  assert.ok(total > 5 * 1048576);
  assert.ok(room.compactions > 0, 'compacted');
  assert.ok(room.bytes < 1048576, `log ${room.bytes} bytes`);
  assert.ok(!sent.some((b) => { try { const s = b.toString('utf8'); return s.includes('"reset"'); } catch { return false; } }), 'no reset');
  // a late joiner replays the log
  const late = new Y.Doc();
  for (const u of room.log) Y.applyUpdate(late, new Uint8Array(u));
  assert.strictEqual(late.getText('t').toString(), text.toString());
  conn.close();
});
test('collab: the shipped Yjs build (server/vendor/yjs.cjs) compacts on its own (no node_modules)', () => {
  const { execFileSync } = require('child_process');
  const path = require('path');
  const code = `
    const Y = require(${JSON.stringify(path.join(__dirname, '..', 'server', 'vendor', 'yjs.cjs'))});
    const a = new Y.Doc(); const t = a.getText('t'); const log = [];
    a.on('update', (u) => log.push(u));
    for (let i = 0; i < 300; i++) { t.insert(0, 'abcdefghij'.repeat(50)); t.delete(0, 400); }
    const d = new Y.Doc(); for (const u of log) Y.applyUpdate(d, u);
    const m = Y.encodeStateAsUpdate(d);
    const e = new Y.Doc(); Y.applyUpdate(e, m);
    const before = log.reduce((n, u) => n + u.length, 0);
    process.stdout.write(JSON.stringify({ same: e.getText('t').toString() === t.toString(), before, after: m.length, merge: typeof Y.mergeUpdates }));`;
  const out = JSON.parse(execFileSync(process.execPath, ['-e', code], { cwd: require('os').tmpdir(), encoding: 'utf8' }));
  assert.strictEqual(out.same, true);
  assert.strictEqual(out.merge, 'function');
  assert.ok(out.after < out.before / 3, JSON.stringify(out)); // 30 000 characters stay (300 × 100)
  const size = require('fs').statSync(path.join(__dirname, '..', 'server', 'vendor', 'yjs.cjs')).size;
  assert.ok(size < 100 * 1024, `vendored yjs ${size} bytes`);
  assert.ok(!/\beval\(|new Function/.test(require('fs').readFileSync(path.join(__dirname, '..', 'server', 'vendor', 'yjs.cjs'), 'utf8')));
});

/* ---------------- 9. clean: accepting a deleted paragraph mark joins the paragraphs ---------------- */
test('clean docx: deleted paragraph marks join the paragraphs (chains, bookmarks between, last in a cell keeps)', () => {
  const { cleanFile } = require('../server/clean');
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const del = '<w:del w:id="9" w:author="Anna" w:date="2026-11-20T10:00:00Z"/>';
  const body = `<w:p><w:pPr><w:jc w:val="center"/><w:rPr>${del}</w:rPr></w:pPr><w:r><w:t xml:space="preserve">Eins </w:t></w:r></w:p>
<w:bookmarkStart w:id="1" w:name="b"/><w:p><w:pPr><w:pPrChange w:id="3" w:author="Anna"><w:pPr/></w:pPrChange><w:rPr>${del}</w:rPr></w:pPr><w:r><w:t xml:space="preserve">zwei </w:t></w:r><w:del w:id="4" w:author="Anna"><w:r><w:delText>weg</w:delText></w:r></w:del></w:p><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>drei</w:t></w:r></w:p>`
    + `<w:tbl><w:tr><w:tc><w:p><w:pPr><w:rPr>${del}</w:rPr></w:pPr><w:r><w:t>Zelle</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Ende</w:t></w:r></w:p><w:sectPr/>`;
  const z = makeZip({ '[Content_Types].xml': '<Types/>', 'word/document.xml': `<w:document ${W}><w:body>${body}</w:body></w:document>` });
  const r = cleanFile(z, 'docx', {});
  const x = readZip(r.data).get('word/document.xml').read().toString();
  const paras = [...x.matchAll(/<w:p>([\s\S]*?)<\/w:p>/g)].map((m) => m[1]);
  const text = (p) => [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
  assert.deepStrictEqual(paras.map(text), ['Eins zwei drei', 'Zelle', 'B', 'Ende']);
  assert.match(paras[0], /^<w:pPr><w:jc w:val="center"\/>/); // the first paragraph's properties stay
  assert.match(paras[0], /<w:bookmarkStart w:id="1" w:name="b"\/>/); // the bookmark moved along
  assert.ok(!/<w:del\b|delText|pPrChange/.test(x));
  assert.ok(r.report.revisions >= 3);
});

/* ---------------- 5. scheduled reports: own recipients per report ---------------- */
test('reports: recipients per report (Hostwerk, alert channels, e-mail addresses, Discord/ntfy targets), format, one failure does not stop the others', async () => {
  const fs = require('fs'); const os = require('os'); const path = require('path');
  const { Reports, normRecipients } = require('../server/extras');
  const { Notifier } = require('../server/notify');
  assert.deepStrictEqual(normRecipients([{ channel: 'mail', to: 'a@x.de; b@y.org' }, { channel: 'hostwerk', to: 'ignored' }, { channel: 'all' }, { channel: 'ntfy', to: 'office-berichte' }]),
    [{ channel: 'mail', to: 'a@x.de, b@y.org' }, { channel: 'hostwerk', to: '' }, { channel: 'alerts', to: '' }, { channel: 'ntfy', to: 'office-berichte' }]);
  assert.throws(() => normRecipients([{ channel: 'mail', to: 'kein-mail' }]), /not a valid e-mail/);
  assert.throws(() => normRecipients([{ channel: 'discord', to: 'https://evil.example/api/webhooks/1/x' }]), /Discord/);
  assert.throws(() => normRecipients([{ channel: 'sms', to: '1' }]), /Unknown channel/);
  assert.deepStrictEqual(normRecipients(null, 'all'), [{ channel: 'alerts', to: '' }]);

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-rep-'));
  const sent = []; const converted = [];
  const notifier = { send: async (ch, msg, opt) => { if (ch === 'discord') throw new Error('Discord: HTTP 500'); sent.push({ ch, to: opt.to, name: msg.attachment.name, type: msg.attachment.type }); } };
  const rep = new Reports({ file: path.join(dir, 'r.json'), readFile: async () => ({ buf: Buffer.from('PK'), ext: 'xlsx' }), convert: async (b, f, t) => { converted.push([f, t]); return { data: Buffer.from(`${t}!`) }; }, notifier, saveCopy: () => ({ name: 'copy.pdf' }), t: (s, v) => String(s).replace(/\{(\w+)\}/g, (m, k) => (v && v[k] != null ? v[k] : m)), log: () => {} });
  const it = rep.add({ key: 'Budget.xlsx', path: 'Budget.xlsx', name: 'Budget.xlsx', schedule: { every: 'weekly', time: '08:00' }, format: 'xlsx', recipients: [{ channel: 'mail', to: 'chef@firma.de' }, { channel: 'discord', to: 'https://discord.com/api/webhooks/123/abc-DEF' }, { channel: 'ntfy', to: 'https://ntfy.example.org/berichte' }, { channel: 'alerts' }] });
  assert.strictEqual(it.schedule.at, '08:00');
  const r = await rep.run(it.id);
  assert.strictEqual(r.ok, true);
  assert.deepStrictEqual(converted, []); // xlsx stays xlsx
  assert.deepStrictEqual(sent.map((x) => [x.ch, x.to]), [['mail', 'chef@firma.de'], ['ntfy', 'https://ntfy.example.org/berichte'], ['all', null]]);
  assert.match(sent[0].name, /^Budget \d{4}-\d{2}-\d{2}\.xlsx$/);
  assert.match(r.error, /Discord: HTTP 500/);
  assert.strictEqual(rep.list()[0].lastResults.length, 4);
  // PDF (default): converted, sent; an old entry with one channel still works
  rep.update(it.id, { format: 'pdf' });
  sent.length = 0; await rep.run(it.id);
  assert.deepStrictEqual(converted, [['xlsx', 'pdf']]);
  assert.strictEqual(sent[0].type, 'application/pdf');
  rep.items.push({ id: 'old', key: 'a.docx', path: 'a.docx', name: 'a.docx', schedule: { every: 'daily', time: '07:00', weekday: 1, day: 1 }, channel: 'hostwerk', on: true });
  assert.deepStrictEqual(rep.list().find((x) => x.id === 'old').recipients, [{ channel: 'hostwerk', to: '' }]);
  sent.length = 0; await rep.run('old');
  assert.deepStrictEqual(sent.map((x) => x.ch), ['hostwerk']);

  // the notifier: own targets instead of the alert addresses; ntfy topic on the alert server
  const calls = [];
  const realFetch = global.fetch;
  global.fetch = async (u, o) => { calls.push({ u: String(u), o }); return { ok: true, status: 200 }; };
  try {
    const n = new Notifier({ host: { t: (s) => s, ctx: { settings: { get: () => ({ live: { alerts: { ntfy: 'https://ntfy.home.lan/alarm', discord: '' } } }) } } } });
    await n.send('ntfy', { title: 'T', text: 'x' }, { to: 'berichte' });
    await n.send('discord', { title: 'T', text: 'x' }, { to: 'https://discord.com/api/webhooks/1/abc' });
    await n.send('ntfy', { title: 'T', text: 'x' });
    assert.deepStrictEqual(calls.map((c) => c.u), ['https://ntfy.home.lan/berichte', 'https://discord.com/api/webhooks/1/abc', 'https://ntfy.home.lan/alarm']);
    await assert.rejects(() => n.send('discord', { title: 'T' }), /No Discord webhook/);
    // e-mail with attachment: one message per address, into Hostwerk's mail catcher without a relay
    const raws = [];
    const n2 = new Notifier({ host: { t: (s) => s, core: (m) => (m === 'smtp-client' ? { sendRaw: async (o) => raws.push(o) } : null), ctx: { services: { mail: { cfg: { smtpPort: 2525 } } }, settings: { get: () => ({ live: { alerts: {} }, mail: {} }) } } } });
    await n2.send('mail', { title: 'Bericht', text: 'anbei', attachment: { name: 'B.pdf', data: Buffer.from('%PDF'), type: 'application/pdf' } }, { to: 'a@x.de, b@y.de' });
    assert.deepStrictEqual(raws.map((x) => [x.to, x.port, x.host]), [['a@x.de', 2525, '127.0.0.1'], ['b@y.de', 2525, '127.0.0.1']]);
    assert.match(raws[0].raw, /X-Hostwerk-Internal: 1/); assert.match(raws[0].raw, /Content-Disposition: attachment/);
  } finally { global.fetch = realFetch; }
});
