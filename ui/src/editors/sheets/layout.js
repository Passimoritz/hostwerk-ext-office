/* Row/column geometry with sparse custom sizes and hidden entries (O(log n) position ↔ index). Unzoomed px. */
import { MAXR, MAXC } from './formula/refs.js';

export class Axis {
  constructor(def, sizes, hiddenSets, max) {
    this.def = def;
    this.max = max;
    const custom = new Map();
    for (const [i, s] of sizes) if (i < max) custom.set(i, s);
    for (const set of hiddenSets) for (const i of set) if (i < max) custom.set(i, 0);
    this.idx = [...custom.keys()].sort((a, b) => a - b);
    this.sz = this.idx.map((i) => custom.get(i));
    this.cum = new Float64Array(this.idx.length + 1);
    for (let k = 0; k < this.idx.length; k++) this.cum[k + 1] = this.cum[k] + (this.sz[k] - def);
    this.map = custom;
  }
  lb(i) { let lo = 0; let hi = this.idx.length; while (lo < hi) { const m = (lo + hi) >> 1; if (this.idx[m] < i) lo = m + 1; else hi = m; } return lo; }
  pos(i) { return i * this.def + this.cum[this.lb(i)]; }
  size(i) { const s = this.map.get(i); return s === undefined ? this.def : s; }
  hidden(i) { return this.map.get(i) === 0; }
  /** index at position p (first index whose span contains p) */
  index(p) {
    if (p <= 0) return this.firstVisible(0);
    let lo = 0; let hi = this.max - 1;
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (this.pos(m) <= p) lo = m; else hi = m - 1; }
    // skip zero-size entries
    while (lo < this.max - 1 && this.size(lo) === 0) lo++;
    return lo;
  }
  firstVisible(i) { while (i < this.max - 1 && this.size(i) === 0) i++; return i; }
  next(i, step = 1) { // next visible index in direction
    let j = i + step;
    while (j >= 0 && j < this.max && this.size(j) === 0) j += step;
    return j < 0 || j >= this.max ? i : j;
  }
}

export function axesOf(sheet) {
  const key = sheet.version;
  if (sheet._axes && sheet._axes.key === key) return sheet._axes;
  const rows = new Axis(sheet.defH, sheet.rowH, [sheet.hiddenRows, sheet.filterHidden], MAXR);
  const cols = new Axis(sheet.defW, sheet.colW, [sheet.hiddenCols], MAXC);
  sheet._axes = { key, rows, cols };
  return sheet._axes;
}
