// German function names, logical values and error codes in Sheets (input/display only – storage stays English)
import test from 'node:test';
import assert from 'node:assert/strict';
import { Book } from '../ui/src/editors/sheets/model.js';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { formatValue } from '../ui/src/editors/sheets/format/numfmt.js';
import { parseInput } from '../ui/src/editors/sheets/format/parse.js';
import { toInternal, toDisplay, toOoxml } from '../ui/src/editors/sheets/formula/parser.js';
import { FUNCTION_NAMES } from '../ui/src/editors/sheets/formula/functions.js';
import { FUNC_INFO, callAt } from '../ui/src/editors/sheets/formula/funcinfo.js';
import { FN_DE, localFn, internalFn, localErr, internalErr, localArgList, isExcelName } from '../ui/src/editors/sheets/formula/i18n.js';
import { ERR, isErr } from '../ui/src/editors/sheets/formula/errors.js';

const de = makeLocale({ lang: 'de' });
const en = makeLocale({ lang: 'en' });
const deEnNames = makeLocale({ lang: 'de', functionNames: 'en' });
const enDeNames = makeLocale({ lang: 'en', functionNames: 'de' });

test('every Excel function of the engine has its German name; no German name hides another function', () => {
  const engine = new Set(FUNCTION_NAMES);
  for (const n of FUNC_INFO.keys()) assert.ok(FN_DE[n], `funcinfo ${n} has no German name`);
  for (const [enName, deName] of Object.entries(FN_DE)) {
    assert.equal(deName, deName.toUpperCase(), deName);
    if (deName !== enName) assert.ok(!engine.has(deName), `${deName} (${enName}) is also an English function name`);
  }
  const seen = new Map();
  for (const [enName, deName] of Object.entries(FN_DE)) { assert.ok(!seen.has(deName), `${deName} used for ${seen.get(deName)} and ${enName}`); seen.set(deName, enName); }
  // engine names without a German name are only the engine's own aliases (CEILINGMATH, CHIDIST.RT …) of a real name
  const extra = new Set(['BINOMDIST.RANGE', 'CHIDIST.RT', 'CHIINV.RT', 'FDIST.RT', 'FINV.RT', 'TDIST.2T', 'TDIST.RT', 'TINV.2T', 'ISO.CEILING.MATH', 'ISO.CEILING.PRECISE', 'LOGNORMINV', 'POISSONDIST', 'WEIBULLDIST']);
  const squash = new Set([...Object.keys(FN_DE), ...extra].map((n) => n.replace(/\./g, '')));
  for (const n of FUNCTION_NAMES.filter((x) => !FN_DE[x])) assert.ok(squash.has(n) || extra.has(n), `${n} needs a German name`);
  assert.ok(isExcelName('vlookup') && !isExcelName('CEILINGMATH'));
});

test('the names the user asked for', () => {
  const want = { IF: 'WENN', SUM: 'SUMME', VLOOKUP: 'SVERWEIS', XLOOKUP: 'XVERWEIS', COUNTIF: 'ZÄHLENWENN', ROUND: 'RUNDEN', TODAY: 'HEUTE',
    NOW: 'JETZT', AVERAGE: 'MITTELWERT', COUNT: 'ANZAHL', COUNTA: 'ANZAHL2', CONCATENATE: 'VERKETTEN', TEXTJOIN: 'TEXTVERKETTEN', INDEX: 'INDEX',
    MATCH: 'VERGLEICH', IFERROR: 'WENNFEHLER', AND: 'UND', OR: 'ODER', NOT: 'NICHT', MAX: 'MAX', MIN: 'MIN', LEFT: 'LINKS', RIGHT: 'RECHTS',
    MID: 'TEIL', LEN: 'LÄNGE', TRIM: 'GLÄTTEN', DATE: 'DATUM', YEAR: 'JAHR', MONTH: 'MONAT', DAY: 'TAG', WEEKDAY: 'WOCHENTAG' };
  for (const [e, d] of Object.entries(want)) { assert.equal(localFn(e, 'de'), d); assert.equal(internalFn(d.toLowerCase(), 'de'), e); }
  assert.equal(localFn('SUM', 'en'), 'SUM');
  assert.equal(internalFn('SUMME', 'en'), 'SUMME', 'English mode does not know German names');
  assert.equal(internalFn('SUM', 'de'), 'SUM', 'English names still work in German');
  // checked against a real German Excel 16 (Range.Formula → Range.FormulaLocal)
  for (const [e, d] of Object.entries({ DOLLAR: 'DM', FORECAST: 'SCHÄTZER', WRAPCOLS: 'SPALTENUMBRUCH', WRAPROWS: 'ZEILENUMBRUCH', 'FORECAST.LINEAR': 'PROGNOSE.LINEAR', OFFSET: 'BEREICH.VERSCHIEBEN' })) { assert.equal(localFn(e, 'de'), d, e); assert.equal(internalFn(d, 'de'), e); }
  assert.equal(internalFn('WÄHRUNG', 'de'), 'DOLLAR', 'also accepted on input');
});

test('input → storage → display', () => {
  assert.equal(toInternal('WENN(A1>1;SVERWEIS(B1;C:D;2;FALSCH);"wahr")', de), 'IF(A1>1,VLOOKUP(B1,C:D,2,FALSE),"wahr")');
  assert.equal(toDisplay('IF(A1>1,VLOOKUP(B1,C:D,2,FALSE),"WAHR")', de), 'WENN(A1>1;SVERWEIS(B1;C:D;2;FALSCH);"WAHR")');
  assert.equal(toInternal('zählenwenn(A:A;"x")+länge(B1)', de), 'COUNTIF(A:A,"x")+LEN(B1)');
  assert.equal(toInternal('WAHR()', de), 'TRUE()');
  assert.equal(toInternal('WENNFEHLER(1/0;#NV)', de), 'IFERROR(1/0,#N/A)');
  assert.equal(toDisplay('IFERROR(1/0,#N/A)+ISNA(#VALUE!)', de), 'WENNFEHLER(1/0;#NV)+ISTNV(#WERT!)');
  assert.equal(toInternal('BEREICH.VERSCHIEBEN(A1;1;1)+GDA2(1;2;3;4)+ZW(1;2;3)', de), 'OFFSET(A1,1,1)+DB(1,2,3,4)+FV(1,2,3)');
  assert.equal(toInternal("Tabelle1!A1+'Mein Blatt'!B2", de), "Tabelle1!A1+'Mein Blatt'!B2", 'sheet names untouched');
  assert.equal(toInternal('"SUMME(1)"&WENN(1;2)', de), '"SUMME(1)"&IF(1,2)', 'strings untouched');
  // the file keeps English names (OOXML prefixes too)
  assert.equal(toOoxml(toInternal('XVERWEIS(A1;B:B;C:C)', de)), '_xlfn.XLOOKUP(A1,B:B,C:C)');
  // setting: English names in the German interface, German names in the English interface
  assert.equal(toDisplay('SUM(A1,1.5)', deEnNames), 'SUM(A1;1,5)');
  assert.equal(toDisplay('SUM(A1,1.5)', enDeNames), 'SUMME(A1,1.5)');
  assert.equal(toDisplay('SUM(A1,1.5)', en), 'SUM(A1,1.5)');
  assert.equal(toInternal('SUMME(1.5)', enDeNames), 'SUM(1.5)');
});

test('German formulas calculate; cells show German errors and logical values', () => {
  const b = Book.blank({ loc: de });
  const s = b.sheets[0];
  b.run('t', () => {
    b.input(s, 0, 0, '10'); b.input(s, 1, 0, '20');
    b.input(s, 0, 1, '=SUMME(A1:A2)');
    b.input(s, 1, 1, '=WENN(A1>5;"groß";"klein")');
    b.input(s, 2, 1, '=1/0');
    b.input(s, 3, 1, '=SVERWEIS(99;A1:A2;1;FALSCH)');
    b.input(s, 4, 1, '=UND(WAHR;NICHT(FALSCH))');
    b.input(s, 5, 1, '=RUNDEN(MITTELWERT(A1:A2)/7;2)');
    b.input(s, 6, 1, '#NV');
  });
  assert.equal(b.value(s, 0, 1), 30);
  assert.equal(b.value(s, 1, 1), 'groß');
  assert.equal(s.cells.get(1 * 16384 + 1).f, 'IF(A1>5,"groß","klein")');
  assert.ok(isErr(b.value(s, 2, 1)));
  assert.equal(formatValue(b.value(s, 2, 1), 'General', de).text, '#DIV/0!');
  assert.equal(formatValue(b.value(s, 3, 1), 'General', de).text, '#NV');
  assert.equal(formatValue(b.value(s, 3, 1), 'General', en).text, '#N/A');
  assert.equal(formatValue(b.value(s, 4, 1), 'General', de).text, 'WAHR');
  assert.equal(b.value(s, 5, 1), 2.14);
  assert.equal(b.value(s, 6, 1).code, ERR.NA.code);
  assert.equal(b.editText(s, 1, 1, (f) => toDisplay(f, de)), '=WENN(A1>5;"groß";"klein")');
  assert.equal(b.editText(s, 3, 1, (f) => toDisplay(f, de)), '=SVERWEIS(99;A1:A2;1;FALSCH)');
  assert.equal(b.editText(s, 6, 1), '#NV');
});

test('error codes and values typed into cells', () => {
  for (const [d, e] of [['#NV', '#N/A'], ['#WERT!', '#VALUE!'], ['#BEZUG!', '#REF!'], ['#ZAHL!', '#NUM!'], ['#ÜBERLAUF!', '#SPILL!'], ['#DIV/0!', '#DIV/0!'], ['#NAME?', '#NAME?'], ['#NULL!', '#NULL!']]) {
    assert.equal(internalErr(d), e); assert.equal(localErr(e, 'de'), d); assert.equal(localErr(e, 'en'), e);
    assert.equal(parseInput(d, de).v.code, e);
  }
  assert.equal(internalErr('#nv'), '#N/A');
  assert.equal(parseInput('wahr', de).v, true);
  assert.equal(parseInput('FALSCH', deEnNames).v, false, 'German UI accepts WAHR/FALSCH also with English names');
  assert.equal(parseInput('WAHR', en).v, 'WAHR');
});

test('argument hints and call detection with German names', () => {
  assert.deepEqual(localArgList('number1, [number2], …', 'de'), ['Zahl1', '[Zahl2]', '…']);
  assert.deepEqual(localArgList('logical_test, [value_if_true], [value_if_false]', 'de'), ['Wahrheitstest', '[Wert_wenn_wahr]', '[Wert_wenn_falsch]']);
  assert.deepEqual(localArgList('lookup_value, table_array, col_index_num, [range_lookup]', 'de'), ['Suchkriterium', 'Matrix', 'Spaltenindex', '[Bereich_Verweis]']);
  assert.deepEqual(localArgList('number1, [number2], …', 'en'), ['number1', '[number2]', '…']);
  assert.deepEqual(callAt('=ZÄHLENWENN(A:A;', 16, ';'), { name: 'ZÄHLENWENN', arg: 1 });
  assert.deepEqual(callAt('=BEREICH.VERSCHIEBEN(A1;1', 24, ';'), { name: 'BEREICH.VERSCHIEBEN', arg: 1 });
});
