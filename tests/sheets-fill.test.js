import test from 'node:test';
import assert from 'node:assert/strict';
import { fillSeries, detectSeriesKind } from '../ui/src/editors/sheets/fill.js';

// tiny 1900-system date helpers (epoch 1899-12-30 for serials >= 61)
const DAY = 86400000; const EPOCH = Date.UTC(1899, 11, 30);
const serialFromParts = (y, m, d, h = 0, mi = 0, s = 0) => (Date.UTC(y, m - 1, d, h, mi, s) - EPOCH) / DAY;
const partsFromSerial = (n) => { const t = new Date(EPOCH + Math.round(n * DAY)); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate(), h: t.getUTCHours(), mi: t.getUTCMinutes(), s: t.getUTCSeconds() }; };
const D = { partsFromSerial, serialFromParts };
const N = (...v) => v.map((x) => ({ v: x }));
const dt = (...s) => s.map((x) => ({ v: serialFromParts(...x), date: true }));
const ymd = (arr) => arr.map((n) => { const p = partsFromSerial(n); return `${p.y}-${p.m}-${p.d}`; });

test('numbers', () => {
  assert.deepEqual(fillSeries(N(1, 2), 3), [3, 4, 5]);
  assert.deepEqual(fillSeries(N(2, 4), 2), [6, 8]);
  assert.deepEqual(fillSeries(N(5), 2), [5, 5]);
  assert.deepEqual(fillSeries(N(0.1, 0.2), 2), [0.3, 0.4]);
  assert.deepEqual(fillSeries(N(1, 2), 2, { copy: true }), [1, 2]);
  assert.equal(detectSeriesKind(N(1, 2)), 'linear');
  assert.equal(detectSeriesKind(N(1)), 'copy');
});
test('dates', () => {
  assert.deepEqual(ymd(fillSeries(dt([2024, 1, 30]), 3, D)), ['2024-1-31', '2024-2-1', '2024-2-2']);
  assert.deepEqual(ymd(fillSeries(dt([2024, 1, 1], [2024, 1, 8]), 1, D)), ['2024-1-15']);
  assert.equal(detectSeriesKind(dt([2024, 1, 31], [2024, 2, 29]), D), 'month');
  assert.deepEqual(ymd(fillSeries(dt([2024, 1, 31], [2024, 2, 29]), 3, D)), ['2024-3-31', '2024-4-30', '2024-5-31']);
  assert.deepEqual(ymd(fillSeries(dt([2023, 1, 31], [2023, 2, 28]), 1, D)), ['2023-3-31']);
  assert.deepEqual(ymd(fillSeries(dt([2020, 5, 1], [2021, 5, 1]), 2, D)), ['2022-5-1', '2023-5-1']);
  assert.equal(detectSeriesKind(dt([2020, 5, 1], [2021, 5, 1]), D), 'year');
  // without helpers: day steps
  assert.equal(detectSeriesKind(dt([2024, 1, 31], [2024, 2, 29])), 'date');
});
test('weekday and month names', () => {
  assert.deepEqual(fillSeries(N('Mon'), 3), ['Tue', 'Wed', 'Thu']);
  assert.deepEqual(fillSeries(N('Saturday'), 3), ['Sunday', 'Monday', 'Tuesday']);
  assert.deepEqual(fillSeries(N('MONDAY'), 1), ['TUESDAY']);
  assert.deepEqual(fillSeries(N('montag', 'mittwoch'), 2), ['freitag', 'sonntag']);
  assert.deepEqual(fillSeries(N('Januar'), 2), ['Februar', 'März']);
  assert.deepEqual(fillSeries(N('Nov'), 3), ['Dec', 'Jan', 'Feb']);
  assert.deepEqual(fillSeries(N('Okt'), 2), ['Nov', 'Dez']);
  assert.equal(detectSeriesKind(N('Mon')), 'weekday');
  assert.equal(detectSeriesKind(N('March')), 'monthname');
});
test('quarters and text with numbers', () => {
  assert.deepEqual(fillSeries(N('Q1'), 4), ['Q2', 'Q3', 'Q4', 'Q1']);
  assert.deepEqual(fillSeries(N('Quartal 3'), 2), ['Quartal 4', 'Quartal 1']);
  assert.deepEqual(fillSeries(N('Item 9'), 2), ['Item 10', 'Item 11']);
  assert.deepEqual(fillSeries(N('Week 01'), 2), ['Week 02', 'Week 03']);
  assert.deepEqual(fillSeries(N('Row 2', 'Row 4'), 2), ['Row 6', 'Row 8']);
  assert.equal(detectSeriesKind(N('Item 1')), 'textnum');
  assert.equal(detectSeriesKind(N('Raum 3a')), 'copy');
});
test('cycles', () => {
  assert.deepEqual(fillSeries(N('a', 'b'), 3), ['a', 'b', 'a']);
  assert.deepEqual(fillSeries(N('x'), 2), ['x', 'x']);
  assert.deepEqual(fillSeries(N(true, false), 3), [true, false, true]);
  assert.deepEqual(fillSeries(N(1, 'a'), 3), [1, 'a', 1]);
  assert.deepEqual(fillSeries(N('A 1', 'B 2'), 2), ['A 1', 'B 2']);
  assert.equal(detectSeriesKind(N(1, 'a')), 'cycle');
});
