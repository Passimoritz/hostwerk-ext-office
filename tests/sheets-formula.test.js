import test from 'node:test';
import assert from 'node:assert/strict';
import { Book } from '../ui/src/editors/sheets/model.js';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { parse, shiftFormula, adjustFormula, toInternal, toDisplay, moveFormulaRefs, renameSheetInFormula, toOoxml, fromOoxml } from '../ui/src/editors/sheets/formula/parser.js';
import { isErr } from '../ui/src/editors/sheets/formula/errors.js';

const en = makeLocale({ lang: 'en' });
const de = makeLocale({ lang: 'de' });

function book(data, loc = en) {
  const b = Book.blank({ loc });
  const s = b.sheets[0];
  b.run('t', () => { for (const [a, v] of Object.entries(data)) { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; b.input(s, +m[2] - 1, c, String(v)); } });
  return b;
}
const val = (b, a, sheet = 0) => { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; const v = b.value(b.sheets[sheet], +m[2] - 1, c); return isErr(v) ? v.code : v; };
const calc = (f, data = {}) => { const b = book({ ...data, Z99: f }); return val(b, 'Z99'); };

test('parser precedence and operators', () => {
  assert.equal(calc('=1+2*3'), 7);
  assert.equal(calc('=(1+2)*3'), 9);
  assert.equal(calc('=-2^2'), 4);
  assert.equal(calc('=2^3^2'), 64);
  assert.equal(calc('=10%'), 0.1);
  assert.equal(calc('="a"&1&TRUE'), 'a1TRUE');
  assert.equal(calc('=1/0'), '#DIV/0!');
  assert.equal(calc('=0.1+0.2'), 0.3);
  assert.equal(calc('=1<2'), true);
  assert.equal(calc('="abc"="ABC"'), true);
  assert.equal(calc('="3"+4'), 7);
  assert.equal(calc('="x"+1'), '#VALUE!');
  assert.equal(calc('=FOO(1)'), '#NAME?');
  assert.equal(calc('=undefinedname'), '#NAME?');
  assert.throws(() => parse('1+'));
});

test('ranges, absolute refs, cross sheet, whole columns', () => {
  const b = book({ A1: 1, A2: 2, A3: 3, B1: '=SUM(A1:A3)', B2: '=SUM(A:A)', B3: '=$A$1+A$2+$A3', C1: '=AVERAGE(A1:A3)', C2: '=COUNT(A1:A3,"x",1)' });
  assert.equal(val(b, 'B1'), 6); assert.equal(val(b, 'B2'), 6); assert.equal(val(b, 'B3'), 6); assert.equal(val(b, 'C1'), 2); assert.equal(val(b, 'C2'), 4);
  const s2 = b.addSheet('Sheet 2');
  b.run('t', () => { b.input(s2, 0, 0, "='Sheet1'!A3*2"); b.input(s2, 0, 1, '=Sheet1!B1+1'); b.input(s2, 0, 2, "=SUM(Sheet1!A1:A2)"); });
  assert.equal(val(b, 'A1', 1), 6); assert.equal(val(b, 'B1', 1), 7); assert.equal(val(b, 'C1', 1), 3);
  b.run('t', () => b.input(b.sheets[0], 0, 0, '10'));
  assert.equal(val(b, 'B1'), 15); assert.equal(val(b, 'B1', 1), 16, 'incremental recalc across sheets');
});

test('functions', () => {
  const d = { A1: 'apple', A2: 'pear', A3: 'apple', B1: 10, B2: 20, B3: 30, C1: 1, C2: 2, C3: 3 };
  assert.equal(calc('=SUMIF(A1:A3,"apple",B1:B3)', d), 40);
  assert.equal(calc('=COUNTIF(B1:B3,">15")', d), 2);
  assert.equal(calc('=COUNTIFS(A1:A3,"ap*",B1:B3,">=30")', d), 1);
  assert.equal(calc('=AVERAGEIFS(B1:B3,A1:A3,"apple")', d), 20);
  assert.equal(calc('=VLOOKUP("pear",A1:B3,2,FALSE)', d), 20);
  assert.equal(calc('=VLOOKUP("kiwi",A1:B3,2,FALSE)', d), '#N/A');
  assert.equal(calc('=VLOOKUP(2.5,C1:C3,1)', d), 2);
  assert.equal(calc('=INDEX(B1:B3,MATCH("pear",A1:A3,0))', d), 20);
  assert.equal(calc('=XLOOKUP("apple",A1:A3,B1:B3,"none",0,-1)', d), 30);
  assert.equal(calc('=XLOOKUP("x",A1:A3,B1:B3,"none")', d), 'none');
  assert.equal(calc('=SUM(INDEX(B1:B3,2):B3)', d), 50);
  assert.equal(calc('=IF(B1>5,"big","small")', d), 'big');
  assert.equal(calc('=IFERROR(1/0,"x")'), 'x');
  assert.equal(calc('=IFS(B1>50,"a",B1>5,"b")', d), 'b');
  assert.equal(calc('=ROUND(2.345,2)'), 2.35);
  assert.equal(calc('=ROUND(-2.5,0)'), -3);
  assert.equal(calc('=MOD(-1,3)'), 2);
  assert.equal(calc('=SUMPRODUCT(B1:B3,C1:C3)', d), 140);
  assert.equal(calc('=TEXT(1234.5,"#,##0.00")'), '1,234.50');
  assert.equal(calc('=LEFT("Hello",2)&UPPER("x")'), 'HeX');
  assert.equal(calc('=DATE(2024,2,30)'), 45352);
  assert.equal(calc('=YEAR(DATE(2024,12,31))'), 2024);
  assert.equal(calc('=EOMONTH(DATE(2024,1,15),1)'), 45351);
  assert.equal(calc('=DATEDIF(DATE(2020,1,1),DATE(2024,6,1),"Y")'), 4);
  assert.equal(calc('=ROWS(A1:B5)*COLUMNS(A1:B5)'), 10);
  assert.equal(calc('=SUM(OFFSET(B1,1,0,2,1))', d), 50);
  assert.equal(calc('=INDIRECT("B"&2)', d), 20);
  assert.equal(calc('=SUBTOTAL(9,B1:B3)', d), 60);
  assert.equal(calc('=AGGREGATE(14,6,B1:B3,1)', d), 30);
  assert.equal(calc('=MEDIAN(C1:C3,"x")'.replace(',"x"', ''), d), 2);
  assert.equal(Math.round(calc('=STDEV.S(B1:B3)', d)), 10);
  assert.equal(Math.round(calc('=PMT(0.05/12,60,10000)') * 100) / 100, -188.71);
  assert.equal(calc('=ISBLANK(Q1)'), true);
  assert.equal(calc('=AND(TRUE,1)'), true);
  assert.equal(calc('=TEXTJOIN("-",TRUE,A1:A3)', d), 'apple-pear-apple');
  assert.equal(calc('=LET(x,2,y,3,x*y)'), 6);
});

test('dynamic arrays spill and #SPILL!', () => {
  const b = book({ A1: 3, A2: 1, A3: 2, C1: '=SORT(A1:A3)', D1: '=SEQUENCE(2,2)', F1: '=UNIQUE({1;1;2})', H1: '=SUM(C1#)' });
  assert.equal(val(b, 'C1'), 1); assert.equal(val(b, 'C2'), 2); assert.equal(val(b, 'C3'), 3);
  assert.equal(val(b, 'E2'), 4); assert.equal(val(b, 'F2'), 2); assert.equal(val(b, 'H1'), 6);
  b.run('t', () => b.input(b.sheets[0], 1, 4, 'x')); // E2 blocks D1's spill
  assert.equal(val(b, 'D1'), '#SPILL!');
  b.run('t', () => b.input(b.sheets[0], 1, 4, ''));
  assert.equal(val(b, 'D1'), 1);
});

test('circular references are detected', () => {
  const b = book({ A1: '=B1+1', B1: '=A1+1', C1: '=5' });
  assert.equal(val(b, 'A1'), '#CIRC!');
  assert.ok(b.calc.circular);
  assert.equal(val(b, 'C1'), 5);
  const b2 = book({ A1: '=SUM(A1:A3)' });
  assert.equal(val(b2, 'A1'), '#CIRC!');
});

test('long dependency chains do not overflow the stack', () => {
  const b = Book.blank({ loc: en });
  const s = b.sheets[0];
  b.run('t', () => { b.input(s, 0, 0, '1'); for (let r = 1; r < 20000; r++) b.input(s, r, 0, `=A${r}+1`); });
  assert.equal(b.value(s, 19999, 0), 20000);
  b.run('t', () => b.input(s, 0, 0, '5'));
  assert.equal(b.value(s, 19999, 0), 20004);
});

test('reference shifting: copy, insert/delete, move, rename', () => {
  assert.equal(shiftFormula('A1+$B$2+B$3+$C4', 1, 1), 'B2+$B$2+C$3+$C5');
  assert.equal(shiftFormula('A1', -1, 0), '#REF!');
  assert.equal(shiftFormula('SUM(A:A)', 3, 1), 'SUM(B:B)');
  const ins = (f, axis, at, count, here = 'S') => adjustFormula(f, { target: 'S', here, axis, at, count });
  assert.equal(ins('A5+A1', 'row', 2, 2), 'A7+A1');
  assert.equal(ins('SUM(A1:A5)', 'row', 2, 3), 'SUM(A1:A8)');
  assert.equal(ins('SUM(A1:A5)', 'row', 1, -2), 'SUM(A1:A3)');
  assert.equal(ins('A3', 'row', 1, -3), '#REF!');
  assert.equal(ins('C1+$C$1', 'col', 0, 1), 'D1+$D$1');
  assert.equal(ins('Other!A5', 'row', 0, 1), 'Other!A5');
  assert.equal(ins("'S'!A5", 'row', 0, 1, 'X'), 'S!A6');
  assert.equal(moveFormulaRefs('A1+B1', { here: 'S', srcSheet: 'S', src: { r1: 0, c1: 0, r2: 0, c2: 0 }, dr: 2, dc: 2 }), 'C3+B1');
  assert.equal(renameSheetInFormula("Data!A1+'Data'!B2", 'Data', 'My Data'), "'My Data'!A1+'My Data'!B2");
  // via the book: insert a row updates formulas
  const b = book({ A1: 1, A2: 2, A3: '=SUM(A1:A2)' });
  b.insert(b.sheets[0], 'row', 1, 1);
  assert.equal(b.sheets[0].cells.get(3 * 16384).f, 'SUM(A1:A3)');
  assert.equal(val(b, 'A4'), 3);
  b.undo();
  assert.equal(val(b, 'A3'), 3);
  assert.equal(b.sheets[0].cells.get(2 * 16384).f, 'SUM(A1:A2)');
});

test('German input: decimal comma and semicolons', () => {
  assert.equal(toInternal('SUM(A1;2,5)', de), 'SUM(A1,2.5)');
  assert.equal(toInternal('SUMME(A1;2,5)', de), 'SUM(A1,2.5)');
  assert.equal(toDisplay('ROUND(A1,2.5)', de), 'RUNDEN(A1;2,5)');
  assert.equal(toDisplay('ROUND(A1,2.5)', makeLocale({ lang: 'de', functionNames: 'en' })), 'ROUND(A1;2,5)');
  assert.equal(toInternal('IF(A1>1;"a;b";1)', en), 'IF(A1>1,"a;b",1)');
  const b = book({ A1: '1,5', B1: '=A1*2', C1: '=ROUND(1,25;1)' }, de);
  assert.equal(val(b, 'A1'), 1.5); assert.equal(val(b, 'B1'), 3); assert.equal(val(b, 'C1'), 1.3);
});

test('OOXML prefixes', () => {
  assert.equal(fromOoxml('_xlfn.XLOOKUP(A1,B:B,C:C)'), 'XLOOKUP(A1,B:B,C:C)');
  assert.equal(toOoxml('XLOOKUP(A1,B:B,C:C)+SUM(A1)'), '_xlfn.XLOOKUP(A1,B:B,C:C)+SUM(A1)');
  assert.equal(toOoxml('SORT(A1:A3)'), '_xlfn._xlws.SORT(A1:A3)');
});

test('undo / redo of edits', () => {
  const b = book({ A1: 1 });
  b.run('t', () => b.input(b.sheets[0], 0, 0, '2'));
  assert.equal(val(b, 'A1'), 2);
  b.undo(); assert.equal(val(b, 'A1'), 1);
  b.redo(); assert.equal(val(b, 'A1'), 2);
});

test('fill series and formula fill', () => {
  const b = book({ A1: 1, A2: 2, B1: '=A1*2' });
  const s = b.sheets[0];
  b.fill(s, { r1: 0, c1: 0, r2: 1, c2: 0 }, { r1: 0, c1: 0, r2: 4, c2: 0 });
  assert.equal(val(b, 'A5'), 5);
  b.fill(s, { r1: 0, c1: 1, r2: 0, c2: 1 }, { r1: 0, c1: 1, r2: 4, c2: 1 });
  assert.equal(s.cells.get(4 * 16384 + 1).f, 'A5*2');
  assert.equal(val(b, 'B5'), 10);
});

test('sort and filter', () => {
  const b = book({ A1: 'Name', A2: 'b', A3: 'a', A4: 'c', B2: 2, B3: 1, B4: 3 });
  const s = b.sheets[0];
  b.sort(s, { r1: 0, c1: 0, r2: 3, c2: 1 }, [{ col: 1, desc: true }], { header: true });
  assert.equal(val(b, 'A2'), 'c'); assert.equal(val(b, 'A4'), 'a');
  b.setFilter(s, { r1: 0, c1: 0, r2: 3, c2: 1 });
  b.setFilterColumn(s, 1, { cond: { op: 'gt', v: '1' } });
  assert.deepEqual([...s.filterHidden], [3]);
});
