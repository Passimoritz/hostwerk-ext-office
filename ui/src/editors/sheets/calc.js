/*
 * Recalculation: parse cache, dependency graph (cell → precedents, reverse index for dependents), incremental
 * recalculation in topological order (iterative, no recursion limits), circular reference detection, dynamic array
 * spilling. Values of formula cells are kept in cell.v.
 */
import { parse, collectDeps } from './formula/parser.js';
import { evaluate, finalize, isRef, structuredRef } from './formula/eval.js';
import { setDate1904 } from './format/dates.js';
import { FUNCTIONS } from './formula/functions.js';
import { ERR, isErr } from './formula/errors.js';
import { MAXC, MAXR, parseRange } from './formula/refs.js';

const K = (r, c) => r * MAXC + c;

export class Calc {
  constructor(book) {
    this.book = book;
    this.ast = new Map(); // formula text → AST | Error
    this.prec = new Map(); // id → { sheet, key, refs:[{sheet,r1,c1,r2,c2}], names, volatile }
    this.single = new Map(); // `${sheetId}:${key}` → Set(id)
    this.ranges = new Map(); // sheetId → [{ r1, c1, r2, c2, id }]
    this.fcols = new Map(); // sheetId → Map(col → Set(row)) of formula cells
    this.volatile = new Set();
    this.nameUsers = new Set();
    this.tableUsers = new Set(); // formulas with structured references (relinked when a table changes)
    this.circular = null;
    this.iterations = 0; // iterations used by the last iterative calculation
    this.pending = new Set(); // manual mode: dirty seeds
  }

  compile(f) {
    let a = this.ast.get(f);
    if (a === undefined) {
      try { a = parse(f); } catch (e) { a = e; }
      if (this.ast.size > 50000) this.ast.clear();
      this.ast.set(f, a);
    }
    return a;
  }

  /* ---------- graph ---------- */
  id(sheet, key) { return `${sheet.id}:${key}`; }
  unlink(id) {
    const p = this.prec.get(id);
    if (!p) return;
    for (const g of p.refs) {
      if (g.single) { const s = this.single.get(`${g.sheet.id}:${K(g.r1, g.c1)}`); if (s) s.delete(id); } else {
        const list = this.ranges.get(g.sheet.id);
        if (list) { const i = list.findIndex((x) => x.id === id && x.g === g); if (i >= 0) list.splice(i, 1); }
      }
    }
    const fc = this.fcols.get(p.sheet.id);
    if (fc) { const s = fc.get(p.key % MAXC); if (s) s.delete(Math.floor(p.key / MAXC)); }
    this.volatile.delete(id); this.nameUsers.delete(id); this.tableUsers.delete(id);
    this.prec.delete(id);
  }
  link(sheet, key, cell) {
    const id = this.id(sheet, key);
    this.unlink(id);
    if (!cell || !cell.f) return;
    const ast = this.compile(cell.f);
    const refs = [];
    let names = [];
    let vol = false;
    if (!(ast instanceof Error)) {
      const d = collectDeps(ast);
      vol = d.volatile;
      names = d.names;
      const addRef = (n, baseSheet) => {
        if (n.sheet2) { for (const x of this.book.sheetsBetween(n.sheet, n.sheet2)) addRef({ ...n, sheet: x.name, sheet2: null }, baseSheet); return; }
        const sh = n.sheet ? this.book.sheetByName(n.sheet) : baseSheet;
        if (!sh) return;
        const r1 = Math.min(n.r1, n.r2); const r2 = Math.max(n.r1, n.r2); const c1 = Math.min(n.c1, n.c2); const c2 = Math.max(n.c1, n.c2);
        refs.push({ sheet: sh, r1, c1, r2, c2, single: r1 === r2 && c1 === c2 });
      };
      for (const n of d.refs) addRef(n, sheet);
      if (d.srefs) {
        const cx = this.ctxFor(sheet, Math.floor(key / MAXC), key % MAXC);
        for (const n of d.srefs) { const g = structuredRef(n, cx); if (isRef(g)) refs.push({ sheet: g.sheet, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, single: g.single }); }
        this.tableUsers.add(this.id(sheet, key));
      }
      // names: add their precedents (one level deep is enough for typical use; deeper names → volatile)
      for (const nm of names) {
        const def = this.book.nameDef(nm.v, sheet);
        if (!def) continue;
        const nast = this.compile(def.ref);
        if (nast instanceof Error) continue;
        const nd = collectDeps(nast);
        for (const n of nd.refs) addRef(n, def.scope || sheet);
        if (nd.names.length || nd.volatile) vol = true;
      }
    }
    const p = { sheet, key, refs, volatile: vol };
    this.prec.set(id, p);
    for (const g of refs) {
      if (g.single) { const k = `${g.sheet.id}:${K(g.r1, g.c1)}`; let s = this.single.get(k); if (!s) this.single.set(k, (s = new Set())); s.add(id); } else {
        let list = this.ranges.get(g.sheet.id); if (!list) this.ranges.set(g.sheet.id, (list = [])); list.push({ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, id, g });
      }
    }
    let fc = this.fcols.get(sheet.id); if (!fc) this.fcols.set(sheet.id, (fc = new Map()));
    const col = key % MAXC; let s = fc.get(col); if (!s) fc.set(col, (s = new Set())); s.add(Math.floor(key / MAXC));
    if (vol) this.volatile.add(id);
    if (names.length) this.nameUsers.add(id);
  }

  /** Full rebuild of the graph (after load or structural changes) */
  rebuild() {
    this.prec.clear(); this.single.clear(); this.ranges.clear(); this.fcols.clear(); this.volatile.clear(); this.nameUsers.clear(); this.tableUsers.clear();
    for (const sh of this.book.sheets) for (const [k, cell] of sh.cells) if (cell.f) this.link(sh, k, cell);
  }

  dependents(sheet, r, c, out) {
    const s = this.single.get(`${sheet.id}:${K(r, c)}`);
    if (s) for (const id of s) out.push(id);
    const list = this.ranges.get(sheet.id);
    if (list) for (const g of list) if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) out.push(g.id);
  }

  /** formula cells inside a range on a sheet (for ordering) */
  formulasIn(sheet, g, cb) {
    const fc = this.fcols.get(sheet.id);
    if (!fc) return;
    for (const [col, rows] of fc) {
      if (col < g.c1 || col > g.c2) continue;
      if (rows.size < g.r2 - g.r1 + 1) { for (const r of rows) if (r >= g.r1 && r <= g.r2) cb(`${sheet.id}:${K(r, col)}`); } else for (let r = g.r1; r <= g.r2; r++) if (rows.has(r)) cb(`${sheet.id}:${K(r, col)}`);
    }
  }

  /* ---------- evaluation ---------- */
  ctxFor(sheet, r, c) {
    const book = this.book;
    const ctx = {
      sheet, r, c, loc: book.loc, fn: FUNCTIONS,
      get: (sh, rr, cc) => book.value(sh, rr, cc),
      sheetByName: (n) => book.sheetByName(n),
      sheetIndex: (sh) => book.sheets.indexOf(sh),
      sheetsBetween: (a, b) => book.sheetsBetween(a, b),
      table: (name, sh, rr, cc) => book.tableDef(name, sh, rr, cc),
      sheetCount: () => book.sheets.length,
      name: (n, sh) => { const d = book.nameDef(n, sh); if (!d) return null; const ast = this.compile(d.ref); return ast instanceof Error ? null : { ast, sheet: d.scope }; },
      bounds: (sh) => book.bounds(sh),
      formulaAt: (sh, rr, cc) => { const x = sh.cells.get(K(rr, cc)); return x && x.f; },
      rowHidden: (sh, rr) => (sh.filterHidden && sh.filterHidden.has(rr) ? 'filter' : sh.hiddenRows.has(rr) ? 'manual' : false),
      spillOf: (sh, rr, cc) => sh.spillAnchors.get(K(rr, cc)),
      onLink: null,
    };
    return ctx;
  }

  evalCell(sheet, key, cell) {
    const r = Math.floor(key / MAXC); const c = key % MAXC;
    const ast = this.compile(cell.f);
    if (ast instanceof Error) return { v: cell._cached !== undefined && cell._cachedF === cell.f ? cell._cached : ERR.NAME };
    const ctx = this.ctxFor(sheet, r, c);
    let link = null;
    ctx.onLink = (u) => { link = u; };
    let v;
    try { v = finalize(evaluate(ast, ctx), ctx); } catch (e) { v = e && e.name === 'RangeError' ? ERR.CALC : ERR.VALUE; }
    // functions/links we can't compute (external workbooks, rare functions): keep the value the file had
    if (isErr(v) && v.code === '#NAME?' && cell._cached !== undefined && cell._cachedF === cell.f) v = cell._cached;
    return { v, link };
  }

  /** Evaluate a formula text in a context without a cell (conditional formats, validation, names) */
  evalText(f, sheet, r = 0, c = 0) {
    const ast = this.compile(f);
    if (ast instanceof Error) return ERR.NAME;
    const ctx = this.ctxFor(sheet, r, c);
    try { const v = evaluate(ast, ctx); return isRef(v) && v.single ? ctx.get(v.sheet, v.r1, v.c1) : v; } catch { return ERR.VALUE; }
  }

  /**
   * Recalculate. seeds: array of { sheet, r, c } that changed (values or formulas); null → everything.
   * Returns the set of ids that were evaluated.
   */
  recalc(seeds, { force = false } = {}) {
    if (!this.book.autoCalc && !force) { if (seeds) for (const s of seeds) this.pending.add(s); else this.pending.add(null); return new Set(); }
    if (force && this.pending.size) { if (this.pending.has(null)) seeds = null; else seeds = [...(seeds || []), ...this.pending]; this.pending.clear(); }
    let todo;
    if (!seeds || seeds.length > 3000) todo = new Set(this.prec.keys());
    else {
      todo = new Set();
      const stack = [];
      for (const s of seeds) {
        const id = this.id(s.sheet, K(s.r, s.c));
        if (this.prec.has(id)) todo.add(id);
        this.dependents(s.sheet, s.r, s.c, stack);
      }
      while (stack.length) {
        const id = stack.pop();
        if (todo.has(id)) continue;
        todo.add(id);
        if (todo.size > 20000) { todo = new Set(this.prec.keys()); stack.length = 0; break; }
        const p = this.prec.get(id);
        if (p) this.dependents(p.sheet, Math.floor(p.key / MAXC), p.key % MAXC, stack);
        // cells spilled by this formula are dependencies of others too
        const sp = p && p.sheet.spillAnchors.get(p.key);
        if (sp) for (let rr = 0; rr < sp.rows; rr++) for (let cc = 0; cc < sp.cols; cc++) if (rr || cc) this.dependents(p.sheet, Math.floor(p.key / MAXC) + rr, (p.key % MAXC) + cc, stack);
      }
    }
    for (const v of this.volatile) todo.add(v);
    this.circular = null;
    setDate1904(this.book.date1904);
    const done = new Set();
    for (let pass = 0; pass < 8; pass++) {
      const changedSpill = this.run(todo);
      for (const id of todo) done.add(id);
      if (!changedSpill.length) break;
      todo = new Set();
      const stack = [];
      for (const s of changedSpill) this.dependents(s.sheet, s.r, s.c, stack);
      for (const id of stack) todo.add(id);
      if (!todo.size) break;
    }
    return done;
  }

  /** topological evaluation of the ids in `todo`; returns spilled cells whose values changed */
  run(todo) {
    const order = [];
    const state = new Map(); // id → 1 visiting, 2 done
    const circ = new Set();
    for (const start of todo) {
      if (state.get(start)) continue;
      const stack = [[start, null]];
      while (stack.length) {
        const top = stack[stack.length - 1];
        const [id] = top;
        if (top[1] === null) {
          if (state.get(id) === 2) { stack.pop(); continue; }
          state.set(id, 1);
          const pre = [];
          const p = this.prec.get(id);
          if (p) for (const g of p.refs) {
            if (g.single) { const pid = `${g.sheet.id}:${K(g.r1, g.c1)}`; if (todo.has(pid)) pre.push(pid); } else this.formulasIn(g.sheet, g, (pid) => { if (todo.has(pid)) pre.push(pid); });
          }
          top[1] = pre; top[2] = 0;
        }
        const pre = top[1];
        if (top[2] < pre.length) {
          const pid = pre[top[2]++];
          const st = state.get(pid);
          if (st === 1) { // cycle: mark everything on the stack from pid
            if (pid === id && this.isSelfRangeOnly(id, pid)) continue;
            let k = stack.length - 1;
            while (k >= 0 && stack[k][0] !== pid) { circ.add(stack[k][0]); k--; }
            circ.add(pid);
          } else if (!st) stack.push([pid, null]);
          continue;
        }
        state.set(id, 2);
        order.push(id);
        stack.pop();
      }
    }
    const spillChanged = [];
    const it = this.book.iterate;
    if (circ.size && it && it.on) return this.iterate(order, circ, it, spillChanged);
    for (const id of order) {
      const p = this.prec.get(id);
      if (!p) continue;
      const cell = p.sheet.cells.get(p.key);
      if (!cell || !cell.f) continue;
      if (circ.has(id)) { this.setResult(p.sheet, p.key, cell, ERR.CIRC, null, spillChanged); if (!this.circular) this.circular = { sheet: p.sheet, key: p.key }; continue; }
      const { v, link } = this.evalCell(p.sheet, p.key, cell);
      this.setResult(p.sheet, p.key, cell, v, link, spillChanged);
    }
    return spillChanged;
  }
  /** iterative calculation (File → Options → Formulas in Excel): circular cells start from their last value and are
   *  evaluated again until no value changes more than `delta` or `count` passes were made */
  iterate(order, circ, it, spillChanged) {
    const max = Math.max(1, Math.min(32767, it.count || 100)); const delta = it.delta ?? 0.001;
    for (const id of circ) { const p = this.prec.get(id); const cell = p && p.sheet.cells.get(p.key); if (cell && (typeof cell.v !== 'number' && typeof cell.v !== 'string' && typeof cell.v !== 'boolean')) cell.v = 0; }
    let n = 0;
    for (; n < max; n++) {
      let change = 0;
      for (const id of order) {
        const p = this.prec.get(id);
        if (!p) continue;
        const cell = p.sheet.cells.get(p.key);
        if (!cell || !cell.f) continue;
        const old = cell.v;
        const { v, link } = this.evalCell(p.sheet, p.key, cell);
        this.setResult(p.sheet, p.key, cell, v, link, spillChanged);
        if (circ.has(id)) change = Math.max(change, typeof v === 'number' && typeof old === 'number' ? Math.abs(v - old) : v === old ? 0 : Infinity);
      }
      if (change <= delta) { n++; break; }
    }
    this.iterations = n;
    return spillChanged;
  }
  isSelfRangeOnly() { return false; }

  setResult(sheet, key, cell, v, link, spillChanged) {
    const r0 = Math.floor(key / MAXC); const c0 = key % MAXC;
    const old = sheet.spillAnchors.get(key);
    if (Array.isArray(v)) {
      const rows = v.length; const cols = v[0].length;
      let blocked = r0 + rows > MAXR || c0 + cols > MAXC;
      for (let r = 0; r < rows && !blocked; r++) for (let c = 0; c < cols; c++) {
        if (!r && !c) continue;
        const k = K(r0 + r, c0 + c);
        const other = sheet.cells.get(k);
        if (other && (other.v != null || other.f)) { blocked = true; break; }
        const sp = sheet.spill.get(k);
        if (sp && sp.anchor !== key) { blocked = true; break; }
      }
      if (old) this.clearSpill(sheet, key, old, spillChanged);
      if (blocked) { cell.v = ERR.SPILL; cell.link_ = null; sheet.spillBlocked.set(key, { rows, cols }); return; }
      sheet.spillBlocked.delete(key);
      sheet.spillAnchors.set(key, { rows, cols });
      for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
        if (!r && !c) continue;
        const k = K(r0 + r, c0 + c);
        sheet.spill.set(k, { v: v[r][c], anchor: key });
        spillChanged.push({ sheet, r: r0 + r, c: c0 + c });
      }
      cell.v = v[0][0];
    } else {
      if (old) this.clearSpill(sheet, key, old, spillChanged);
      sheet.spillBlocked.delete(key);
      cell.v = v;
    }
    cell.link_ = link || null;
  }
  clearSpill(sheet, key, sp, spillChanged) {
    const r0 = Math.floor(key / MAXC); const c0 = key % MAXC;
    for (let r = 0; r < sp.rows; r++) for (let c = 0; c < sp.cols; c++) {
      if (!r && !c) continue;
      const k = K(r0 + r, c0 + c);
      const s = sheet.spill.get(k);
      if (s && s.anchor === key) { sheet.spill.delete(k); spillChanged.push({ sheet, r: r0 + r, c: c0 + c }); }
    }
    sheet.spillAnchors.delete(key);
  }
  /** When a cell is written into a spill area, the anchor must recalc (→ #SPILL!) */
  spillAnchorsAt(sheet, r, c) {
    const out = [];
    const s = sheet.spill.get(K(r, c));
    if (s) out.push(s.anchor);
    for (const [k, d] of sheet.spillBlocked) { const r0 = Math.floor(k / MAXC); const c0 = k % MAXC; if (r >= r0 && r < r0 + d.rows && c >= c0 && c < c0 + d.cols && k !== K(r, c)) out.push(k); }
    return out;
  }
}

export { parseRange, isErr };
