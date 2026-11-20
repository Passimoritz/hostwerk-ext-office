/*
 * What-if data tables (Excel: Data → What-If Analysis → Data Table, xlsx <f t="dataTable" ref dt2D dtr r1 r2/>).
 * The table formula itself stays in the file (cell._dtF, written back unchanged); this module recalculates the result
 * cells like Excel does: put each input value into the input cell, recalculate, take the formula's result, restore.
 *
 *   sheet.dataTables = [{ ref: { r1, c1, r2, c2 }, dt2D, dtr, r1: 'E1', r2: null }]
 *   recalcDataTables(book) → number of cells whose value changed (results are calc output: no undo step, not dirty)
 */
import { MAXC, parseRange } from './formula/refs.js';

const K = (r, c) => r * MAXC + c;
const MAX_CELLS = 5000;

function withInput(book, sh, inputs, fn) {
  const saved = inputs.map(({ r, c }) => sh.cells.get(K(r, c)));
  const seeds = inputs.map(({ r, c }) => ({ sheet: sh, r, c }));
  try {
    return fn((vals) => {
      inputs.forEach(({ r, c }, i) => { const old = saved[i]; const n = { ...(old || {}), v: vals[i] }; delete n.f; sh.cells.set(K(r, c), n); book.calc.link(sh, K(r, c), n); });
      book.calc.recalc(seeds, { force: true });
    });
  } finally {
    inputs.forEach(({ r, c }, i) => { const k = K(r, c); if (saved[i]) sh.cells.set(k, saved[i]); else sh.cells.delete(k); book.calc.link(sh, k, saved[i] || null); });
    book.calc.recalc(seeds, { force: true });
  }
}

export function recalcDataTables(book) {
  let changed = 0;
  for (const sh of book.sheets) {
    for (const dt of sh.dataTables || []) {
      const g = dt.ref;
      if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > MAX_CELLS) continue;
      const a = parseRange(dt.r1 || ''); const b = dt.r2 ? parseRange(dt.r2) : null;
      if (!a) continue;
      const results = [];
      if (dt.dt2D) {
        // row input cell r1 gets the values of the row above, column input cell r2 those of the column left
        if (!b) continue;
        withInput(book, sh, [{ r: a.r1, c: a.c1 }, { r: b.r1, c: b.c1 }], (set) => {
          for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
            set([book.value(sh, g.r1 - 1, c), book.value(sh, r, g.c1 - 1)]);
            results.push([r, c, book.value(sh, g.r1 - 1, g.c1 - 1)]);
          }
        });
      } else if (dt.dtr) {
        // row-oriented: input values in the row above, formulas in the column left
        withInput(book, sh, [{ r: a.r1, c: a.c1 }], (set) => {
          for (let c = g.c1; c <= g.c2; c++) { set([book.value(sh, g.r1 - 1, c)]); for (let r = g.r1; r <= g.r2; r++) results.push([r, c, book.value(sh, r, g.c1 - 1)]); }
        });
      } else {
        // column-oriented: input values in the column left, formulas in the row above
        withInput(book, sh, [{ r: a.r1, c: a.c1 }], (set) => {
          for (let r = g.r1; r <= g.r2; r++) { set([book.value(sh, r, g.c1 - 1)]); for (let c = g.c1; c <= g.c2; c++) results.push([r, c, book.value(sh, g.r1 - 1, c)]); }
        });
      }
      const seeds = [];
      for (const [r, c, v] of results) {
        const k = K(r, c); const old = sh.cells.get(k);
        if (!old || !old._dt || old.f) continue; // overwritten by the user
        const ov = old ? old.v : null;
        if (ov === v || (v && ov && v.code && v.code === ov.code)) continue;
        sh.cells.set(k, { ...(old || {}), v });
        seeds.push({ sheet: sh, r, c });
        changed++;
      }
      if (seeds.length) book.calc.recalc(seeds, { force: true });
    }
  }
  return changed;
}
