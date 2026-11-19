import test from 'node:test';
import assert from 'node:assert/strict';
import { formatValue, changeDecimals, fmtKind, isDateFormat, isTimeOnly, decimalsForValue, makeNumberFormat, makeCurrencyFormat, ACCOUNTING, FORMAT_PRESETS, BUILTIN_FORMATS } from '../ui/src/editors/sheets/format/numfmt.js';
import { parseInput, parseNumberLoose } from '../ui/src/editors/sheets/format/parse.js';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { serialFromParts, partsFromSerial, dateToSerial, serialToDate } from '../ui/src/editors/sheets/format/dates.js';
import { ERR } from '../ui/src/editors/sheets/formula/errors.js';

const en = makeLocale({ lang: 'en' });
const de = makeLocale({ lang: 'de' });
const f = (v, fmt, loc = en) => formatValue(v, fmt, loc).text;

test('locale', () => {
  assert.equal(de.decimal, ','); assert.equal(de.group, '.'); assert.equal(de.argSep, ';'); assert.equal(de.dateOrder, 'DMY');
  assert.equal(en.argSep, ','); assert.equal(en.dateOrder, 'MDY');
  const x = makeLocale({ lang: 'en', decimal: ',' });
  assert.equal(x.group, '.'); assert.equal(x.argSep, ';');
  assert.equal(makeLocale({ lang: 'de', dateFormat: 'yyyy-mm-dd' }).dateOrder, 'YMD');
});

test('serial dates', () => {
  assert.equal(serialFromParts(1900, 1, 1), 1);
  assert.equal(serialFromParts(1900, 2, 28), 59);
  assert.equal(serialFromParts(1900, 2, 29), 60);
  assert.equal(serialFromParts(1900, 3, 1), 61);
  assert.equal(serialFromParts(2023, 3, 15), 45000);
  assert.equal(serialFromParts(2024, 13, 1), serialFromParts(2025, 1, 1));
  assert.equal(serialFromParts(2024, 3, 0), serialFromParts(2024, 2, 29));
  assert.equal(serialFromParts(24, 1, 1), serialFromParts(1924, 1, 1));
  assert.deepEqual(partsFromSerial(45000.5), { y: 2023, m: 3, d: 15, h: 12, mi: 0, s: 0, ms: 0, dow: 3 });
  assert.equal(partsFromSerial(60).d, 29);
  assert.equal(partsFromSerial(1).dow, 0);
  const d = new Date(2026, 10, 18, 13, 45);
  assert.equal(serialToDate(dateToSerial(d)).getTime(), d.getTime());
});

test('General', () => {
  assert.equal(f(0, 'General'), '0');
  assert.equal(f(1234.5, 'General'), '1234.5');
  assert.equal(f(1234.5, 'General', de), '1234,5');
  assert.equal(f(1 / 3, 'General'), '0.333333333');
  assert.equal(f(-2.5, 'General'), '-2.5');
  assert.equal(f(123456789012, 'General'), '1.23457E+11');
  assert.equal(f(0.00001, 'General'), '1E-05');
  assert.equal(f(0.1 + 0.2, null), '0.3');
  assert.equal(f('abc', 'General'), 'abc');
  assert.equal(f(true, 'General', de), 'WAHR');
  assert.equal(f(ERR.DIV0, '0.00'), '#DIV/0!');
  assert.equal(f(null, '0.00'), '');
});

test('numbers and rounding', () => {
  assert.equal(f(3.14159, '0.00'), '3.14');
  assert.equal(f(1.005, '0.00'), '1.01');
  assert.equal(f(2.5, '0'), '3');
  assert.equal(f(-2.5, '0'), '-3');
  assert.equal(f(1234567.891, '#,##0.00'), '1,234,567.89');
  assert.equal(f(1234567.891, '#,##0.00', de), '1.234.567,89');
  assert.equal(f(0.5, '#.00'), '.50');
  assert.equal(f(5, '00000'), '00005');
  assert.equal(f(1234567, '#,##0,'), '1,235');
  assert.equal(f(1234567, '0.0,,'), '1.2');
  assert.equal(f(1.5, '0.0#'), '1.5');
  assert.equal(f(1.5, '0.???'), '1.5  ');
  assert.equal(f(-5, '0.00'), '-5.00');
  assert.equal(f(1e21, '0'), '1000000000000000000000');
});

test('sections, colours, conditions, text', () => {
  const code = '#,##0.00;[Red]-#,##0.00';
  assert.deepEqual(formatValue(-1234.5, code, en), { text: '-1,234.50', color: '#FF0000' });
  assert.deepEqual(formatValue(1234.5, code, de), { text: '1.234,50', color: null });
  assert.equal(f(-5, '0;(0)'), '(5)');
  assert.equal(f(0, '0;-0;"zero"'), 'zero');
  assert.equal(f('hi', '0;-0;0;"<"@">"'), '<hi>');
  assert.equal(f('hi', '@'), 'hi');
  assert.equal(f(150, '[>100]"big";[<=100]"small"'), 'big');
  assert.equal(f(50, '[>100]"big";[<=100]"small"'), 'small');
  assert.equal(formatValue(5, '[Blue][>=0]0;[Red]0', en).color, '#0000FF');
  assert.equal(formatValue(5, '[Color10]0', en).color, '#008000');
  assert.equal(f(5, ';;;'), '');
  assert.equal(f(12, '0 "pcs"'), '12 pcs');
  assert.equal(f(12, '0\\h'), '12h');
});

test('currency, accounting, percent, scientific', () => {
  assert.equal(f(1234.5, de.currencyFmt, de), '1.234,50 €');
  assert.equal(f(1234.5, en.currencyFmt), '$1,234.50');
  assert.equal(f(-1234.5, en.currencyFmt), '-$1,234.50');
  assert.equal(f(1234.5, '[$€-407] #,##0.00'), '€ 1,234.50');
  assert.equal(f(1234.5, ACCOUNTING(de), de).trim(), '1.234,50 €');
  assert.equal(f(-1234.5, ACCOUNTING(en)).trim(), '$(1,234.50)');
  assert.equal(f(0, ACCOUNTING(en)).replace(/\s+/g, ''), '$-');
  assert.equal(f(0.256, '0%'), '26%');
  assert.equal(f(0.256, '0.0%', de), '25,6%');
  assert.equal(f(12345, '0.00E+00'), '1.23E+04');
  assert.equal(f(0.000123, '0.00E+00', de), '1,23E-04');
  assert.equal(f(12345, '##0.0E+0'), '12.3E+3');
  assert.equal(f(1234.5, BUILTIN_FORMATS[4]), '1,234.50');
});

test('fractions', () => {
  assert.equal(f(1.5, '# ?/?'), '1 1/2');
  assert.equal(f(0.75, '# ?/?'), '3/4');
  assert.equal(f(1.25, '# ??/??'), '1  1/4 ');
  assert.equal(f(0.3333, '# ??/??').trim(), '1/3');
  assert.equal(f(3.14159, '# ??/??').trim(), '3 14/99');
  assert.equal(f(3.14159, '# ?/?').trim(), '3 1/7');
  assert.equal(f(0.375, '?/8'), '3/8');
  assert.equal(f(2, '# ?/?').trimEnd(), '2');
  assert.equal(f(-1.5, '# ?/?'), '-1 1/2');
});

test('dates and times', () => {
  const d = serialFromParts(2026, 11, 18, 13, 45, 30);
  assert.equal(f(d, 'dd.mm.yyyy', de), '18.11.2026');
  assert.equal(f(d, 'd. mmmm yyyy', de), '18. November 2026');
  assert.equal(f(serialFromParts(2026, 3, 5), 'd. mmmm yyyy', de), '5. März 2026');
  assert.equal(f(serialFromParts(2026, 3, 5), 'dd. mmm yy', de), '05. Mrz 26');
  assert.equal(f(d, 'mmm d, yyyy'), 'Nov 18, 2026');
  assert.equal(f(d, 'dddd', de), 'Mittwoch');
  assert.equal(f(d, 'ddd', en), 'Wed');
  assert.equal(f(d, 'mmmmm'), 'N');
  assert.equal(f(d, 'h:mm AM/PM'), '1:45 PM');
  assert.equal(f(d, 'hh:mm:ss'), '13:45:30');
  assert.equal(f(d, 'yyyy-mm-dd hh:mm'), '2026-11-18 13:45');
  assert.equal(f(1.5 + 1 / 1440, '[h]:mm'), '36:01');
  assert.equal(f(0.5 / 86400 * 3, 'mm:ss.0', de), '00:01,5');
  assert.equal(f(d, 14, de), '18.11.2026');
  assert.equal(f(d, 14, en), '11/18/2026');
  assert.equal(f(-1, 'dd.mm.yyyy'), '########');
  assert.equal(f(0.99999999, 'hh:mm'), '00:00');
});

test('format kinds and helpers', () => {
  assert.ok(isDateFormat('dd.mm.yyyy')); assert.ok(isDateFormat(14)); assert.ok(!isDateFormat('0.00'));
  assert.ok(isTimeOnly('hh:mm')); assert.ok(!isTimeOnly('dd.mm.yyyy hh:mm'));
  assert.equal(fmtKind('General'), 'general');
  assert.equal(fmtKind('#,##0.00;[Red]-#,##0.00'), 'number');
  assert.equal(fmtKind('#,##0.00 €'), 'currency');
  assert.equal(fmtKind(ACCOUNTING(en)), 'accounting');
  assert.equal(fmtKind('0%'), 'percent');
  assert.equal(fmtKind('0.00E+00'), 'scientific');
  assert.equal(fmtKind('# ?/?'), 'fraction');
  assert.equal(fmtKind('hh:mm'), 'time');
  assert.equal(fmtKind('@'), 'text');
  assert.equal(FORMAT_PRESETS(de).length, 11);
  assert.equal(makeNumberFormat({ decimals: 1, thousands: true, negRed: true }), '#,##0.0;[Red]-#,##0.0');
  assert.equal(makeCurrencyFormat({ symbol: '$', decimals: 2, after: false }), '$#,##0.00');
});

test('changeDecimals', () => {
  assert.equal(changeDecimals('0.00', 1), '0.000');
  assert.equal(changeDecimals('0.00', -1), '0.0');
  assert.equal(changeDecimals('0.0', -1), '0');
  assert.equal(changeDecimals('0', -1), '0');
  assert.equal(changeDecimals('#,##0', 1), '#,##0.0');
  assert.equal(changeDecimals('#,##0.00;[Red]-#,##0.00', 1), '#,##0.000;[Red]-#,##0.000');
  assert.equal(changeDecimals('0%', 2), '0.00%');
  assert.equal(changeDecimals('#,##0.00 "€"', -2), '#,##0 "€"');
  assert.equal(changeDecimals('General', 1, 1.5), '0.00');
  assert.equal(changeDecimals('General', -1, 1.25), '0.0');
  assert.equal(changeDecimals('dd.mm.yyyy', 1), 'dd.mm.yyyy');
  assert.equal(decimalsForValue(3), 0);
  assert.equal(decimalsForValue(1.125), 3);
});

test('parseInput en', () => {
  assert.deepEqual(parseInput('1234.5', en), { v: 1234.5 });
  assert.deepEqual(parseInput('1,234.5', en), { v: 1234.5, fmt: '#,##0.00' });
  assert.deepEqual(parseInput('-12', en), { v: -12 });
  assert.deepEqual(parseInput('(12)', en), { v: -12 });
  assert.deepEqual(parseInput('12%', en), { v: 0.12, fmt: '0%' });
  assert.deepEqual(parseInput('12.5%', en), { v: 0.125, fmt: '0.0%' });
  assert.deepEqual(parseInput('$1,200.50', en), { v: 1200.5, fmt: '$#,##0.00' });
  assert.deepEqual(parseInput('1e5', en), { v: 100000, fmt: '0.00E+00' });
  assert.deepEqual(parseInput('11/18/2026', en), { v: serialFromParts(2026, 11, 18), fmt: en.shortDate });
  assert.deepEqual(parseInput('2026-11-18', en), { v: serialFromParts(2026, 11, 18), fmt: en.shortDate });
  assert.deepEqual(parseInput('3/4', en).fmt, en.shortDate);
  assert.deepEqual(parseInput('1 1/2', en), { v: 1.5, fmt: '# ?/?' });
  assert.deepEqual(parseInput('1:30 PM', en), { v: 13.5 / 24, fmt: en.time });
  assert.deepEqual(parseInput('Nov 18, 2026', en).v, serialFromParts(2026, 11, 18));
  assert.deepEqual(parseInput('true', en), { v: true });
  assert.deepEqual(parseInput("'0123", en), { v: '0123' });
  assert.deepEqual(parseInput('#N/A', en), { v: ERR.NA });
  assert.deepEqual(parseInput('hello', en), { v: 'hello' });
  assert.deepEqual(parseInput('1,5', en), { v: '1,5' });
  assert.deepEqual(parseInput('=A1', en), { v: '=A1' });
  assert.deepEqual(parseInput('', en), { v: null });
});

test('parseInput de', () => {
  assert.deepEqual(parseInput('1234,5', de), { v: 1234.5 });
  assert.deepEqual(parseInput('1.234,5', de), { v: 1234.5, fmt: '#,##0.00' });
  assert.deepEqual(parseInput('1.234', de), { v: 1234, fmt: '#,##0' });
  assert.deepEqual(parseInput('12 €', de), { v: 12, fmt: '#,##0 €' });
  assert.deepEqual(parseInput('12,50 €', de), { v: 12.5, fmt: '#,##0.00 €' });
  assert.deepEqual(parseInput('-3,5%', de), { v: -0.035, fmt: '0.0%' });
  assert.deepEqual(parseInput('18.11.2026', de), { v: serialFromParts(2026, 11, 18), fmt: de.shortDate });
  assert.deepEqual(parseInput('18.11.26', de).v, serialFromParts(2026, 11, 18));
  assert.deepEqual(parseInput('18.11.', de).v, serialFromParts(new Date().getFullYear(), 11, 18));
  assert.deepEqual(parseInput('18.11.2026 13:45', de), { v: serialFromParts(2026, 11, 18, 13, 45), fmt: de.dateTime });
  assert.deepEqual(parseInput('31.02.2026', de), { v: '31.02.2026' });
  assert.deepEqual(parseInput('3/4', de), { v: 0.75, fmt: '# ?/?' });
  assert.deepEqual(parseInput('13:45', de), { v: (13 * 60 + 45) / 1440, fmt: 'hh:mm' });
  assert.deepEqual(parseInput('13:45:10', de).fmt, 'hh:mm:ss');
  assert.deepEqual(parseInput('25:30', de).fmt, '[h]:mm:ss');
  assert.deepEqual(parseInput('WAHR', de), { v: true });
  assert.deepEqual(parseInput('falsch', de), { v: false });
  assert.deepEqual(parseInput('18. November 2026', de).v, serialFromParts(2026, 11, 18));
});

test('parseNumberLoose', () => {
  assert.equal(parseNumberLoose('1.234,56', de), 1234.56);
  assert.equal(parseNumberLoose('1,234.56', en), 1234.56);
  assert.equal(parseNumberLoose('  -7 ', en), -7);
  assert.equal(parseNumberLoose('12,5 %', de), 0.125);
  assert.equal(parseNumberLoose('€ 3,00', de), 3);
  assert.equal(parseNumberLoose('abc', de), null);
  assert.equal(parseNumberLoose('', de), null);
});
