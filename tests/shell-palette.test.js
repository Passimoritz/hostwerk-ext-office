// Command palette ranking, compare fallback diff and perf summary of the shell – node --test
import test from 'node:test';
import assert from 'node:assert/strict';
import { score, rank, pushRecent, fold } from '../ui/src/shell/paletteModel.js';
import { lineDiff } from '../ui/src/shell/diff.js';
import { perfSummary } from '../ui/src/lib/perf.js';

const CMDS = [
  { id: 'bold', title: 'Bold' }, { id: 'table', title: 'Insert table' }, { id: 'comment', title: 'Insert comment' },
  { id: 'pdf', title: 'Download as PDF (.pdf)', group: 'File' }, { id: 'share', title: 'Share link …', group: 'File' },
  { id: 'einf', title: 'Einfügen' }, { id: 'hidden', title: 'Hidden', when: () => false },
];

test('score: exact > prefix > word start > substring > initials > scattered', () => {
  const s = (q) => score(q, 'Insert table');
  assert.ok(score('insert table', 'Insert table') > s('insert'));
  assert.ok(s('insert') > s('table'));
  assert.ok(s('table') > s('ert'));
  assert.ok(s('ert') > s('it'));
  assert.ok(s('it') > 0);
  assert.equal(score('xyz', 'Insert table'), 0);
  assert.equal(score('', 'anything'), 1);
});

test('umlauts and case are folded', () => {
  assert.equal(fold('Einfügen'), 'einfugen');
  assert.ok(score('einfug', 'Einfügen') > 0);
  assert.ok(score('EINFÜ', 'einfügen') > 0);
});

test('rank: filters, hides when() = false, multi-word queries, recent first', () => {
  assert.deepEqual(rank(CMDS, 'ins tab').map((c) => c.id)[0], 'table');
  assert.ok(!rank(CMDS, '').some((c) => c.id === 'hidden'));
  assert.equal(rank(CMDS, 'pdf')[0].id, 'pdf');
  assert.equal(rank(CMDS, 'file share')[0].id, 'share', 'the group counts too');
  const r = rank(CMDS, '', ['share', 'bold']);
  assert.deepEqual(r.slice(0, 2).map((c) => c.id), ['share', 'bold']);
  assert.deepEqual(pushRecent(['a', 'b', 'c'], 'b'), ['b', 'a', 'c']);
  assert.equal(pushRecent(Array.from({ length: 20 }, (_, i) => String(i)), 'x').length, 12);
});

test('lineDiff', () => {
  const d = lineDiff('a\nb\nc\nd', 'a\nB\nc\nd\ne');
  assert.deepEqual(d.map((x) => `${x.op}${x.text}`), ['=a', '-b', '+B', '=c', '=d', '+e']);
  assert.deepEqual(lineDiff('same', 'same'), [{ op: '=', text: 'same' }]);
  assert.deepEqual(lineDiff('', 'x').filter((x) => x.op !== '=').map((x) => x.op), ['-', '+']);
});

test('perfSummary: median and last per kind', () => {
  const s = perfSummary({ docs: [300, 100, 200], sheets: [50] });
  assert.deepEqual(s.docs, { n: 3, median: 200, last: 200 });
  assert.deepEqual(s.sheets, { n: 1, median: 50, last: 50 });
});
