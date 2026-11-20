// Print layout page model (ui/src/editors/docs/extensions/pagemodel.js): where lines and table rows go when a page is
// full. Regression for "table row drawn across the page gap / under the next page's header" (ISR-Abgleich, 1.2).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pager, planLines, planTable } from '../ui/src/editors/docs/extensions/pagemodel.js';

// A4 portrait at 96 dpi, margins 2.5 / 2 cm: text area 971.3 px high
const MM = 96 / 25.4;
const A4 = { pageW: 210 * MM, pageH: 297 * MM, mt: 25 * MM, mb: 20 * MM, ml: 25 * MM, mr: 25 * MM, contentW: 160 * MM, contentH: 252 * MM };
const GAP = 24;
const newPager = () => { const P = pager([A4], GAP); P.add(0); return P; };

/** lines of `n` text lines of height h starting at y */
const lines = (y, n, h = 16) => Array.from({ length: n }, (_, k) => ({ top: y + k * h, bottom: y + (k + 1) * h }));

/** final y of every line of a planned table: row pushes move the whole row, cell spacers move a cell from that line on */
function place(rows, res, shift = 0) {
  const out = [];
  let added = 0;
  rows.forEach((row, j) => {
    const base0 = shift + added;
    let rowPush = 0;
    for (const it of res.items) if (it.kind === 'row' && it.row === j) rowPush += it.height;
    const cellShift = row.cells.map(() => 0);
    const growth = [];
    row.cells.forEach((cell, c) => {
      const ls = cell.lines;
      const ys = ls.map((l, k) => {
        let s = rowPush;
        for (const it of res.items) if (it.kind === 'cell' && it.row === j && it.cell === c && it.line <= k) s += it.height;
        return { top: l.top + base0 + s, bottom: l.bottom + base0 + s, row: j, cell: c, k };
      });
      cellShift[c] = ys.length ? ys[ys.length - 1].bottom - (ls[ls.length - 1].bottom + base0) : rowPush;
      growth.push(cell.bottom + base0 + cellShift[c]);
      out.push(...ys);
    });
    const nb = Math.max(row.bottom + base0 + rowPush, ...growth);
    added += nb - (row.bottom + base0);
  });
  return out;
}
/** every line inside the text area of one page (never in the gap, never in a margin) */
function assertInsidePages(P, ys) {
  for (const y of ys) {
    const p = P.pages.findIndex((_, i) => y.top >= P.C(i) - 0.6 && y.bottom <= P.E(i) + 0.6);
    assert.ok(p >= 0, `line ${y.k} of row ${y.row} cell ${y.cell} at ${y.top.toFixed(1)}–${y.bottom.toFixed(1)} lies across a page end / in a margin`);
  }
}
const row = (top, cellLineCounts, { cantSplit = false, pad = 4, h = 16 } = {}) => {
  const cells = cellLineCounts.map((n) => ({ lines: lines(top + pad, n, h), bottom: top + pad + n * h + pad }));
  const bottom = Math.max(...cells.map((c) => c.bottom));
  return { top, bottom, cantSplit, cells };
};
/** consecutive rows from y */
function table(y, specs) { const rows = []; let t = y; for (const s of specs) { const r = row(t, s.n, s); rows.push(r); t = r.bottom; } return rows; }

test('pager: text area, paper edges and the gap', () => {
  const P = newPager();
  const p1 = P.next(0);
  assert.equal(P.C(0), 0);
  assert.ok(Math.abs(P.E(0) - 252 * MM) < 0.01);
  assert.ok(Math.abs(P.edge(0) - (P.E(0) + 20 * MM)) < 0.01);
  assert.ok(Math.abs(P.top(p1) - (P.edge(0) + GAP)) < 0.01);
  assert.ok(Math.abs(P.C(p1) - (P.top(p1) + 25 * MM)) < 0.01);
});

test('lines: a paragraph breaks between lines, the rest starts at the next page\'s text top', () => {
  const P = newPager();
  const ls = lines(P.E(0) - 40, 6);
  const r = planLines(ls, { shift: 0, page: 0, P });
  assert.equal(r.whole, false);
  assert.equal(r.spacers.length, 1);
  assert.equal(r.spacers[0].line, 2); // two lines fit (32 px of 40)
  assert.ok(Math.abs(ls[2].top + r.spacers[0].height - P.C(1)) < 0.2);
  // not even the first line fits → the block moves
  assert.equal(planLines(lines(P.E(0) - 10, 3), { shift: 0, page: 0, P: newPager() }).whole, true);
});

test('table: a row that does not fit moves whole (cantSplit) – header repeated above it', () => {
  const P = newPager();
  const head = row(P.E(0) - 300, [1]);
  const rows = [head, ...table(head.bottom, Array.from({ length: 10 }, () => ({ n: [2, 3], cantSplit: true })))];
  const res = planTable(rows, { shift: 0, page: 0, P, headN: 1, headH: head.bottom - head.top });
  assert.equal(res.whole, false);
  const moved = res.items.filter((i) => i.kind === 'row');
  assert.equal(moved.length, 1);
  const it = moved[0];
  // the row's top lands below the repeated header at the next page's text top
  assert.ok(Math.abs(rows[it.row].top + it.height - (P.C(1) + it.head)) < 0.2);
  assert.ok(it.head > 0);
  assertInsidePages(P, place(rows, res));
});

test('table: a row that may break is split between its lines (Word: "allow row to break across pages")', () => {
  const P = newPager();
  const rows = table(P.E(0) - 120, [{ n: [1, 2] }, { n: [2, 14, 5] }, { n: [1, 1] }]);
  const res = planTable(rows, { shift: 0, page: 0, P });
  const cells = res.items.filter((i) => i.kind === 'cell');
  assert.ok(cells.length >= 1, 'the long cell continues on the next page');
  assert.equal(res.items.filter((i) => i.kind === 'row').length, 0, 'nothing moved whole');
  assert.ok(cells[0].band && cells[0].band.cut > cells[0].band.cutTop, 'a band covers the rest of the page');
  assertInsidePages(P, place(rows, res));
});

test('table: rows taller than a page are broken over several pages, never across a gap', () => {
  const P = newPager();
  const rows = table(200, [{ n: [1] }, { n: [3, 150, 2], cantSplit: true }, { n: [2, 2] }]); // 150 lines ≈ 2.5 pages
  const res = planTable(rows, { shift: 0, page: 0, P, headN: 1, headH: rows[0].bottom - rows[0].top });
  assert.equal(res.whole, false);
  assert.ok(res.items.filter((i) => i.kind === 'cell').length >= 2);
  assert.ok(P.pages.length >= 3);
  assertInsidePages(P, place(rows, res));
});

test('table: header row that does not fit at the page end → the whole table moves', () => {
  const P = newPager();
  const rows = table(P.E(0) - 20, [{ n: [2] }, { n: [2] }, { n: [2] }]);
  assert.equal(planTable(rows, { shift: 0, page: 0, P, headN: 1, headH: 40 }).whole, true);
  // a cantSplit first body row that cannot stay with its header moves the table too
  const P2 = newPager();
  const r2 = table(P2.E(0) - 60, [{ n: [1] }, { n: [4], cantSplit: true }]);
  assert.equal(planTable(r2, { shift: 0, page: 0, P: P2, headN: 1, headH: 24 }).whole, true);
  // forced (the table already starts at a page top): it is laid out anyway
  const P3 = newPager();
  const r3 = table(0, [{ n: [1] }, ...Array.from({ length: 80 }, () => ({ n: [1, 2], cantSplit: true }))]);
  const res = planTable(r3, { shift: 0, page: 0, P: P3, headN: 1, headH: 24, force: true });
  assert.equal(res.whole, false);
  assertInsidePages(P3, place(r3, res));
});

test('table: random tables never put a line into a gap or margin', () => {
  let seed = 7;
  const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
  for (let t = 0; t < 60; t++) {
    const P = newPager();
    const n = 5 + Math.floor(rnd() * 40);
    const specs = Array.from({ length: n }, () => ({ n: Array.from({ length: 1 + Math.floor(rnd() * 5) }, () => 1 + Math.floor(rnd() * (rnd() < 0.08 ? 90 : 12))), cantSplit: rnd() < 0.5 }));
    specs.unshift({ n: [1, 1] });
    const rows = table(rnd() * 700, specs);
    const headN = rnd() < 0.5 ? 1 : 0;
    let res = planTable(rows, { shift: 0, page: 0, P, headN, headH: headN ? rows[0].bottom - rows[0].top : 0 });
    let shift = 0;
    if (res.whole) { shift = P.C(P.next(0)) - rows[0].top; res = planTable(rows, { shift, page: 1, P, headN, headH: headN ? rows[0].bottom - rows[0].top : 0, force: true }); }
    assertInsidePages(P, place(rows, res, shift));
  }
});
