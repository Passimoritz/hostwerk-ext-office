import test from 'node:test';
import assert from 'node:assert/strict';
import { decodeText, encodeText, detectDelimiter, parseCSV, writeCSV, sniff } from '../ui/src/editors/sheets/io/csv.js';

test('parse quotes, embedded delimiters, newlines and doubled quotes', () => {
  const t = 'a,"b,c","say ""hi"""\r\n1,"line1\nline2",3\r\n';
  assert.deepEqual(parseCSV(t, { delimiter: ',' }), [['a', 'b,c', 'say "hi"'], ['1', 'line1\nline2', '3']]);
});
test('CR / LF / CRLF, ragged rows, empty fields, no trailing empty row', () => {
  assert.deepEqual(parseCSV('a,b\rc\nd,,e\r\n', { delimiter: ',' }), [['a', 'b'], ['c'], ['d', '', 'e']]);
  assert.deepEqual(parseCSV('', { delimiter: ',' }), []);
  assert.deepEqual(parseCSV('x,\n', { delimiter: ',' }), [['x', '']]);
});
test('delimiter detection', () => {
  assert.equal(detectDelimiter('Name;Betrag\nA;1,5\nB;2,25\n'), ';');
  assert.equal(detectDelimiter('a\tb\tc\n1\t2\t3\n'), '\t');
  assert.equal(detectDelimiter('a|b\n1|2\n'), '|');
  assert.equal(detectDelimiter('a,b\n1,2\n'), ',');
  assert.equal(detectDelimiter('"x;y",b\n"1;2",3\n'), ',');
  assert.equal(detectDelimiter('single'), ',');
  assert.deepEqual(parseCSV('a;b\n1,5;2'), [['a', 'b'], ['1,5', '2']]);
});
test('maxRows', () => {
  assert.equal(parseCSV('1\n2\n3\n4', { delimiter: ',', maxRows: 2 }).length, 2);
});
test('write quotes only where needed', () => {
  assert.equal(writeCSV([['a', 'b;c', 'x"y'], [1, null, ' s']], { delimiter: ';' }), 'a;"b;c";"x""y"\r\n1;;" s"\r\n');
  assert.equal(writeCSV([['a', 'b']], { quoteAll: true, eol: '\n' }), '"a","b"\n');
  const rows = [['ä', 'line\nbreak', 'q"']];
  assert.deepEqual(parseCSV(writeCSV(rows), { delimiter: ',' }), rows);
});
test('windows-1252 round trip with umlauts and euro sign', () => {
  const s = 'Größe;Preis €;„Zitat“';
  const b = encodeText(s, 'windows-1252');
  assert.equal(b.length, s.length);
  assert.equal(b[2], 0xF6); assert.equal(b[12], 0x80);
  assert.deepEqual(decodeText(b, 'windows-1252'), { text: s, encoding: 'windows-1252' });
  assert.equal(decodeText(b).encoding, 'windows-1252'); // invalid UTF-8 → auto falls back
  assert.equal(decodeText(b).text, s);
  assert.equal(encodeText('日', 'windows-1252')[0], 0x3F);
});
test('iso-8859-1 and UTF-8 / BOM handling', () => {
  assert.equal(decodeText(encodeText('Müller', 'iso-8859-1'), 'iso-8859-1').text, 'Müller');
  const bom = encodeText('Müller', 'utf-8-bom');
  assert.deepEqual([...bom.subarray(0, 3)], [0xEF, 0xBB, 0xBF]);
  assert.deepEqual(decodeText(bom), { text: 'Müller', encoding: 'utf-8' });
  assert.deepEqual(decodeText(encodeText('Müller', 'utf-8')), { text: 'Müller', encoding: 'utf-8' });
});
test('UTF-16 LE/BE with BOM', () => {
  for (const e of ['utf-16le', 'utf-16be']) {
    const b = encodeText('Straße;€', e);
    assert.deepEqual(decodeText(b), { text: 'Straße;€', encoding: e });
  }
});
test('sniff', () => {
  const r = sniff(encodeText('a;b\n1;2\n', 'utf-8-bom'));
  assert.equal(r.delimiter, ';'); assert.equal(r.encoding, 'utf-8');
  assert.deepEqual(r.preview, [['a', 'b'], ['1', '2']]);
});
