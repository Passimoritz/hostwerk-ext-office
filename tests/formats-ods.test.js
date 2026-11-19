import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { Book } from '../ui/src/editors/sheets/model.js';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { readSheetJS, writeSheetJS } from '../ui/src/editors/sheets/io/sheetjs.js';
import { toOpenFormula, fromOpenFormula, odfRawFormulas } from '../ui/src/editors/sheets/io/odf-formula.js';

// finding #7 of the Microsoft Office check: SheetJS CE wrote `of:=Data![.D7]*2` (Excel: #NAME?/#REF!)
const loc = makeLocale({ lang: 'en' });
const K = (r, c) => r * 16384 + c;

function book() {
  const b = Book.blank({ loc });
  const s = b.sheets[0];
  b.renameSheet(s, 'Summary');
  const d = b.addSheet('Data');
  const m = b.addSheet('My Sheet');
  b.run('fill', () => {
    b.input(d, 2, 0, 'Apple'); b.input(d, 3, 0, 'Pear'); b.input(d, 4, 0, 'Plum'); b.input(d, 5, 0, 'Fig');
    b.input(d, 6, 3, '17.4');
    b.input(m, 0, 0, '5');
    b.input(s, 0, 1, '=Data!D7*2');
    b.input(s, 1, 1, '=COUNTA(Data!A3:A6)');
    b.input(s, 2, 1, '=CONCATENATE(Data!A3,"-, ;|",Data!A4)');
    b.input(s, 3, 1, "='My Sheet'!$A$1+1");
    b.input(s, 4, 1, '=SUM(Data!$D:$D)');
    b.input(s, 5, 1, '=IF(TRUE,SUM({1,2;3,4}),0)');
  });
  return b;
}

test('toOpenFormula writes OpenFormula references', () => {
  assert.equal(toOpenFormula('Data!D7*2'), 'of:=[$Data.D7]*2');
  assert.equal(toOpenFormula('COUNTA(Data!A3:A6)'), 'of:=COUNTA([$Data.A3:.A6])');
  assert.equal(toOpenFormula("'My Sheet'!$A$1+A2"), "of:=[$'My Sheet'.$A$1]+[.A2]");
  assert.equal(toOpenFormula('SUM(A:A,3:4)'), 'of:=SUM([.A:.A];[.3:.4])');
  assert.equal(toOpenFormula('CONCATENATE("a,b;c",B1)'), 'of:=CONCATENATE("a,b;c";[.B1])');
  assert.equal(toOpenFormula('SUM({1,2;3,4})'), 'of:=SUM({1;2|3;4})');
  assert.equal(toOpenFormula('IFS(A1>1,1,TRUE,0)'), 'of:=COM.MICROSOFT.IFS([.A1]>1;1;TRUE();0)');
});

test('fromOpenFormula reads LibreOffice/Excel OpenFormula', () => {
  assert.equal(fromOpenFormula('of:=[Data.D7]*2'), 'Data!D7*2');
  assert.equal(fromOpenFormula('of:=[$Data.$D$7]*2'), 'Data!$D$7*2');
  assert.equal(fromOpenFormula('of:=COUNTA([$Data.A3:.A6])'), 'COUNTA(Data!A3:A6)');
  assert.equal(fromOpenFormula("of:=[$'My Sheet'.A1]+[.B2]"), "'My Sheet'!A1+B2");
  assert.equal(fromOpenFormula('of:=SUM([.A:.A];[.3:.4])'), 'SUM(A:A,3:4)');
  assert.equal(fromOpenFormula('of:=CONCATENATE("a;b|c";[.B1])'), 'CONCATENATE("a;b|c",B1)');
  assert.equal(fromOpenFormula('of:=SUM({1;2|3;4})'), 'SUM({1,2;3,4})');
  assert.equal(fromOpenFormula('of:=COM.MICROSOFT.IFS([.A1]>1;1;TRUE();0)'), 'IFS(A1>1,1,TRUE,0)');
  assert.equal(fromOpenFormula('of:=[.#REF!]+1'), '#REF!+1');
  assert.equal(fromOpenFormula('of:=IFERROR(#N/A;#DIV/0!)'), 'IFERROR(#N/A,#DIV/0!)');
  assert.equal(fromOpenFormula('of:=SUM([.A1:.B2] ! [.B1:.C2])'), 'SUM(A1:B2   B1:C2)');
  assert.equal(fromOpenFormula('msoxl:=Data!A1'), 'Data!A1');
  assert.equal(fromOpenFormula('oooc:=[.A1]'), null);
});

for (const ext of ['ods', 'fods']) {
  test(`${ext}: cross-sheet formulas are written as OpenFormula and read back`, async () => {
    const b = book();
    const bytes = writeSheetJS(b, ext);
    let xml;
    if (ext === 'ods') {
      // mimetype first and stored (ODF 1.2 part 3)
      assert.equal(new TextDecoder().decode(bytes.subarray(30, 38)), 'mimetype');
      assert.equal(bytes[8], 0);
      const z = await JSZip.loadAsync(bytes);
      assert.equal(await z.file('mimetype').async('string'), 'application/vnd.oasis.opendocument.spreadsheet');
      xml = await z.file('content.xml').async('string');
    } else xml = new TextDecoder().decode(bytes);
    const fs = [...xml.matchAll(/table:formula="([^"]*)"/g)].map((m) => m[1]);
    assert.ok(fs.includes('of:=[$Data.D7]*2'), fs.join(' '));
    assert.ok(fs.includes('of:=COUNTA([$Data.A3:.A6])'));
    assert.ok(fs.includes('of:=CONCATENATE([$Data.A3];&quot;-, ;|&quot;;[$Data.A4])'));
    assert.ok(fs.includes("of:=[$'My Sheet'.$A$1]+1"));
    assert.ok(fs.includes('of:=SUM([$Data.$D:.$D])'));
    assert.ok(!fs.some((f) => /HWFORMULA|!\[/.test(f)), fs.join(' '));
    const r = readSheetJS(bytes, { loc, ext });
    const s = r.sheets[0];
    assert.equal(s.cells.get(K(0, 1)).f, 'Data!D7*2');
    assert.equal(s.cells.get(K(1, 1)).f, 'COUNTA(Data!A3:A6)');
    assert.equal(s.cells.get(K(2, 1)).f, 'CONCATENATE(Data!A3,"-, ;|",Data!A4)');
    assert.equal(s.cells.get(K(3, 1)).f, "'My Sheet'!$A$1+1");
    assert.equal(r.value(s, 0, 1), 34.8);
    assert.equal(r.value(s, 1, 1), 4);
    assert.equal(r.value(s, 2, 1), 'Apple-, ;|Pear');
    assert.equal(r.value(s, 3, 1), 6);
    assert.equal(r.value(s, 4, 1), 17.4);
    assert.equal(r.value(s, 5, 1), 10);
  });
}

test('named ranges survive ods', () => {
  const b = book();
  const d = b.sheets[1];
  b.names.set('FRUIT|', { name: 'Fruit', ref: 'Data!$A$3:$A$6', scope: null, comment: '' });
  b.names.set('TAX|', { name: 'Tax', ref: '0.19', scope: null, comment: '' });
  const r = readSheetJS(writeSheetJS(b, 'ods'), { loc, ext: 'ods' });
  const n = [...r.names.values()].find((x) => x.name === 'Fruit');
  assert.ok(n, 'Fruit kept');
  assert.equal(n.ref.replace(/\$/g, ''), 'Data!A3:A6');
  assert.ok(d);
});

test('raw formula scan follows repeated rows/columns and covered cells', () => {
  const xml = '<office:spreadsheet><table:table table:name="A"><table:table-column table:number-columns-repeated="3"/>'
    + '<table:table-row table:number-rows-repeated="2"><table:table-cell/></table:table-row>'
    + '<table:table-row><table:table-cell table:number-columns-repeated="2"/><table:covered-table-cell/><table:table-cell table:formula="of:=[.A1]"><text:p>1</text:p></table:table-cell></table:table-row>'
    + '</table:table><table:table table:name="B"><table:table-row><table:table-cell table:formula="of:=1"/></table:table-row></table:table></office:spreadsheet>';
  const s = odfRawFormulas(xml);
  assert.equal(s.length, 2);
  assert.equal(s[0].get('2,3'), 'of:=[.A1]');
  assert.equal(s[1].get('0,0'), 'of:=1');
});
