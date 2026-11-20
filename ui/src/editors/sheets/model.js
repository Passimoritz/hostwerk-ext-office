/*
 * Workbook model + undo/redo + all editing operations (pure – no React, runs in Node tests).
 *
 * Cell = { v, f?, s?, note?, link?, _si?, arr? } – cells are treated as immutable (always replaced), except that the
 * calc engine writes results of formula cells into cell.v.  s = interned style object (see StyleTable).
 * Style = { font, size, b, i, u, st, color, fill, h, va, wrap, rot, indent, nf, bd: { t, b, l, r: { s, c } }, _xf }
 */
import { Calc } from './calc.js';
import { MAXR, MAXC, key as K, parseRange, rangeText, quoteSheet } from './formula/refs.js';
import { shiftFormula, adjustFormula, moveFormulaRefs, renameSheetInFormula, toInternal, tokenize } from './formula/parser.js';
import { isErr, ERR } from './formula/errors.js';
import { localErr } from './formula/i18n.js';
import { compare } from './formula/eval.js';
import { cmpSort, criteria } from './formula/functions.js';
import { parseInput } from './format/parse.js';
import { formatValue, isDateFormat } from './format/numfmt.js';
import { fillSeries } from './fill.js';
import { partsFromSerial, serialFromParts, setDate1904 } from './format/dates.js';

let UID = 1;
export const uid = (p = 'x') => `${p}${Date.now().toString(36)}${(UID++).toString(36)}`;
const R = (k) => Math.floor(k / MAXC);
const C = (k) => k % MAXC;

/* ---------------- styles ---------------- */
export class StyleTable {
  constructor() { this.map = new Map(); }
  intern(s) {
    if (!s) return null;
    const o = {};
    for (const k of Object.keys(s).sort()) { const v = s[k]; if (v !== undefined && v !== null && v !== false && v !== '' && !(k === 'nf' && v === 'General')) o[k] = v; }
    if (o.bd) { const b = {}; for (const e of ['t', 'b', 'l', 'r', 'd']) if (o.bd[e] && o.bd[e].s) b[e] = o.bd[e]; if (Object.keys(b).length) o.bd = b; else delete o.bd; }
    const keys = Object.keys(o);
    if (!keys.length || (keys.length === 1 && keys[0] === '_xf' && false)) return null;
    const id = JSON.stringify(o);
    let x = this.map.get(id);
    if (!x) { x = Object.freeze(o); this.map.set(id, x); }
    return x;
  }
  /** merge patch into a style; patch values of undefined remove a key */
  merge(s, patch) {
    const o = { ...(s || {}) };
    delete o._xf; // a modified style is no longer the file's original xf
    for (const [k, v] of Object.entries(patch)) { if (v === undefined) delete o[k]; else o[k] = v; }
    return this.intern(o);
  }
}

/* ---------------- sheets ---------------- */
export function newSheet(name) {
  return {
    id: uid('s'), name, color: null, hidden: false,
    cells: new Map(), colW: new Map(), rowH: new Map(), hiddenCols: new Set(), hiddenRows: new Set(), filterHidden: new Set(),
    outlineRows: new Map(), outlineCols: new Map(),
    defW: 64, defH: 20,
    merges: [], freeze: { rows: 0, cols: 0 }, gridlines: true, headings: true, zoom: null,
    cf: [], dv: [], filter: null, charts: [], images: [], tables: [], pivots: [], sparklines: [], shapes: [], slicers: [], view: null, print: { area: null, orientation: 'portrait', fit: 'width', gridlines: false, headings: false },
    spill: new Map(), spillAnchors: new Map(), spillBlocked: new Map(),
    sel: null, _xml: null, version: 0,
  };
}
const jc = (x) => JSON.parse(JSON.stringify(x));
const cloneSheetProps = (s) => ({
  name: s.name, color: s.color, hidden: s.hidden, colW: new Map(s.colW), rowH: new Map(s.rowH), hiddenCols: new Set(s.hiddenCols), hiddenRows: new Set(s.hiddenRows),
  filterHidden: new Set(s.filterHidden), outlineRows: new Map(s.outlineRows), outlineCols: new Map(s.outlineCols), defW: s.defW, defH: s.defH,
  merges: s.merges.map((m) => ({ ...m })), freeze: { ...s.freeze }, gridlines: s.gridlines, headings: s.headings,
  cf: JSON.parse(JSON.stringify(s.cf)), dv: JSON.parse(JSON.stringify(s.dv)), filter: s.filter ? JSON.parse(JSON.stringify(s.filter)) : null,
  charts: JSON.parse(JSON.stringify(s.charts)), images: s.images.map((i) => ({ ...i, anchor: { ...i.anchor } })), print: { ...s.print },
  tables: jc(s.tables || []), pivots: jc(s.pivots || []), sparklines: jc(s.sparklines || []), shapes: jc(s.shapes || []), view: s.view ? jc(s.view) : null,
});

/* iterate cells of ranges (only existing cells when the range is huge) */
function forRange(sheet, g, cb, { create = true, bounds } = {}) {
  const area = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
  if (area > 300000 || !create) {
    const keys = [];
    for (const k of sheet.cells.keys()) { const r = R(k); const c = C(k); if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) keys.push(k); }
    if (create && bounds) { // also create cells inside the used area (formatting a column of a used sheet)
      const r2 = Math.min(g.r2, bounds.rows - 1); const c2 = Math.min(g.c2, bounds.cols - 1);
      if ((r2 - g.r1 + 1) * (c2 - g.c1 + 1) <= 300000) { const have = new Set(keys); for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) { const k = K(r, c); if (!have.has(k)) keys.push(k); } }
    }
    for (const k of keys) cb(R(k), C(k), k);
    return;
  }
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) cb(r, c, K(r, c));
}

/* ---------------- the book ---------------- */
export class Book {
  constructor({ loc, settings = {} } = {}) {
    this.loc = loc;
    this.settings = settings;
    this.sheets = [];
    this.names = new Map(); // UPPER(name)|scopeId → { name, ref, scope: sheet|null, comment, hidden }
    this.styles = new StyleTable();
    this.date1904 = false;
    this.iterate = { on: false, count: 100, delta: 0.001 }; // iterative calculation (Excel: calcPr iterate)
    this.autoCalc = settings.autoCalc !== false;
    this.active = 0;
    this.props = {};
    this.pkg = null; // original package info (xlsx patching)
    this.fidelity = [];
    this.calc = new Calc(this);
    this.undoStack = []; this.redoStack = [];
    this.tx = null;
    this.version = 0;
    this.listeners = new Set();
    this.boundsCache = new Map();
  }
  static blank({ loc, settings, sheetName = 'Sheet1' } = {}) {
    const b = new Book({ loc, settings });
    b.sheets.push(newSheet(sheetName));
    b.applyDefaults();
    return b;
  }
  applyDefaults() {
    const s = this.settings || {};
    for (const sh of this.sheets) { if (s.gridlines === false && !sh._xml) sh.gridlines = false; }
  }
  /** 1904 date system (serial 0 = 1904-01-01) – dates.js follows the book that was set last */
  get date1904() { return this._d1904; }
  set date1904(v) { this._d1904 = !!v; setDate1904(this._d1904); }
  /** sheets of a 3D reference Sheet1:Sheet3 (in tab order, either direction) */
  sheetsBetween(a, b) {
    const i = this.sheets.indexOf(this.sheetByName(a)); const j = this.sheets.indexOf(this.sheetByName(b));
    if (i < 0 || j < 0) return [];
    return this.sheets.slice(Math.min(i, j), Math.max(i, j) + 1);
  }
  /** all tables (ListObjects) with their sheet */
  allTables() { const out = []; for (const sh of this.sheets) for (const t of sh.tables || []) out.push({ t, sheet: sh }); return out; }
  /** column names of a table: header cells (when it has a header row) or the stored names */
  tableColumns(sheet, t) {
    const g = t.range; const out = [];
    for (let c = g.c1; c <= g.c2; c++) {
      const j = c - g.c1;
      let n = t.header ? this.value(sheet, g.r1, c) : null;
      n = n == null || n === '' ? (t.columns && t.columns[j]) || `Column${j + 1}` : typeof n === 'number' ? this.text(sheet, g.r1, c) : String(n);
      out.push(n);
    }
    return out;
  }
  /** table by name, or the table containing (sheet, r, c) when name is empty → { name, range, header, totals, columns, sheet } */
  tableDef(name, sheet, r, c) {
    let hit = null;
    if (name) { const lo = String(name).toLowerCase(); hit = this.allTables().find((x) => String(x.t.name).toLowerCase() === lo); } else if (sheet) {
      const t = (sheet.tables || []).find((x) => r >= x.range.r1 && r <= x.range.r2 && c >= x.range.c1 && c <= x.range.c2);
      if (t) hit = { t, sheet };
    }
    if (!hit) return null;
    return { ...hit.t, header: hit.t.header !== false, totals: !!hit.t.totals, sheet: hit.sheet, columns: this.tableColumns(hit.sheet, { ...hit.t, header: hit.t.header !== false }) };
  }
  on(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }
  emit(info) { this.version++; for (const l of this.listeners) l(info); }

  /* ---------- lookup ---------- */
  sheetByName(name) { const lo = String(name).toLowerCase(); return this.sheets.find((s) => s.name.toLowerCase() === lo) || null; }
  sheetById(id) { return this.sheets.find((s) => s.id === id) || null; }
  nameDef(name, sheet) {
    const up = String(name).toUpperCase();
    if (sheet) { const d = this.names.get(`${up}|${sheet.id}`); if (d) return d; }
    return this.names.get(`${up}|`) || null;
  }
  cell(sheet, r, c) { return sheet.cells.get(K(r, c)) || null; }
  value(sheet, r, c) {
    const k = K(r, c);
    const cell = sheet.cells.get(k);
    if (cell) return cell.v === undefined ? null : cell.v;
    const sp = sheet.spill.get(k);
    return sp ? sp.v : null;
  }
  bounds(sheet) {
    let b = this.boundsCache.get(sheet.id);
    if (!b || b.v !== sheet.version) {
      let rows = 0; let cols = 0;
      for (const k of sheet.cells.keys()) { const r = R(k) + 1; const c = C(k) + 1; if (r > rows) rows = r; if (c > cols) cols = c; }
      for (const k of sheet.spill.keys()) { const r = R(k) + 1; const c = C(k) + 1; if (r > rows) rows = r; if (c > cols) cols = c; }
      b = { rows, cols, v: sheet.version };
      this.boundsCache.set(sheet.id, b);
    }
    return b;
  }
  /** used range with content (not only styles) */
  usedRange(sheet) {
    let r1 = Infinity; let c1 = Infinity; let r2 = -1; let c2 = -1;
    const add = (k) => { const r = R(k); const c = C(k); if (r < r1) r1 = r; if (c < c1) c1 = c; if (r > r2) r2 = r; if (c > c2) c2 = c; };
    for (const [k, cell] of sheet.cells) if (cell.v != null || cell.f || cell.s) add(k);
    for (const k of sheet.spill.keys()) add(k);
    for (const m of sheet.merges) { add(K(m.r1, m.c1)); add(K(m.r2, m.c2)); }
    return r2 < 0 ? null : { r1, c1, r2, c2 };
  }
  dataRange(sheet) {
    let r2 = -1; let c2 = -1;
    for (const [k, cell] of sheet.cells) if (cell.v != null || cell.f) { if (R(k) > r2) r2 = R(k); if (C(k) > c2) c2 = C(k); }
    for (const k of sheet.spill.keys()) { if (R(k) > r2) r2 = R(k); if (C(k) > c2) c2 = C(k); }
    return r2 < 0 ? null : { r1: 0, c1: 0, r2, c2 };
  }
  styleOf(sheet, r, c) { const x = sheet.cells.get(K(r, c)); return (x && x.s) || null; }
  /** formatted display text of a cell */
  text(sheet, r, c) {
    const v = this.value(sheet, r, c);
    if (v == null) return '';
    const s = this.styleOf(sheet, r, c);
    return formatValue(v, (s && s.nf) || 'General', this.loc).text;
  }

  /* ---------- transactions / undo ---------- */
  begin(label, meta = {}) {
    if (this.tx) { this.tx.depth++; return this.tx; }
    this.tx = { label, depth: 1, cells: [], snaps: new Map(), book: null, seeds: [], structural: false, meta };
    return this.tx;
  }
  /** snapshot sheet props (and cells if withCells) before changing them */
  touch(sheet, withCells = false) {
    const tx = this.tx;
    if (!tx) throw new Error('touch outside a transaction');
    let s = tx.snaps.get(sheet);
    if (!s) { s = { props: cloneSheetProps(sheet), cells: null }; tx.snaps.set(sheet, s); }
    if (withCells && !s.cells) s.cells = new Map(sheet.cells);
    sheet.version++;
  }
  touchBook() {
    if (!this.tx.book) this.tx.book = { sheets: this.sheets.slice(), names: new Map(this.names), active: this.active };
  }
  end(info = {}) {
    const tx = this.tx;
    if (!tx) return;
    if (--tx.depth > 0) return;
    this.tx = null;
    if (!tx.cells.length && !tx.snaps.size && !tx.book) return;
    // after-state of snapshots
    for (const [sheet, s] of tx.snaps) { s.after = { props: cloneSheetProps(sheet), cells: s.cells ? new Map(sheet.cells) : null }; }
    if (tx.book) tx.bookAfter = { sheets: this.sheets.slice(), names: new Map(this.names), active: this.active };
    this.undoStack.push(tx);
    if (this.undoStack.length > 100) this.undoStack.shift();
    this.redoStack = [];
    this.afterChange(tx);
    this.emit({ label: tx.label, ...info });
  }
  afterChange(tx) {
    if (tx.structural || tx.book) { this.calc.rebuild(); this.calc.recalc(null); } else this.calc.recalc(tx.seeds);
  }
  run(label, fn, meta) { this.begin(label, meta); try { return fn(); } finally { this.end(); } }
  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }
  undo() {
    const tx = this.undoStack.pop();
    if (!tx) return null;
    if (tx.book) { this.sheets = tx.book.sheets.slice(); this.names = new Map(tx.book.names); this.active = tx.book.active; }
    for (const [sheet, s] of tx.snaps) { Object.assign(sheet, cloneSheetProps({ ...sheet, ...s.props })); if (s.cells) sheet.cells = new Map(s.cells); sheet.version++; }
    for (let i = tx.cells.length - 1; i >= 0; i--) { const [sheet, k, before] = tx.cells[i]; if (before) sheet.cells.set(k, before); else sheet.cells.delete(k); sheet.version++; }
    this.redoStack.push(tx);
    this.calc.rebuild(); this.calc.recalc(null);
    this.emit({ label: tx.label, undo: true, meta: tx.meta });
    return tx;
  }
  redo() {
    const tx = this.redoStack.pop();
    if (!tx) return null;
    for (const [sheet, k, , after] of tx.cells) { if (after) sheet.cells.set(k, after); else sheet.cells.delete(k); sheet.version++; }
    for (const [sheet, s] of tx.snaps) { Object.assign(sheet, cloneSheetProps({ ...sheet, ...s.after.props })); if (s.after.cells) sheet.cells = new Map(s.after.cells); sheet.version++; }
    if (tx.bookAfter) { this.sheets = tx.bookAfter.sheets.slice(); this.names = new Map(tx.bookAfter.names); this.active = tx.bookAfter.active; }
    this.undoStack.push(tx);
    this.calc.rebuild(); this.calc.recalc(null);
    this.emit({ label: tx.label, redo: true, meta: tx.meta });
    return tx;
  }

  /** the one primitive for cell writes */
  setCell(sheet, r, c, cell) {
    const k = K(r, c);
    const before = sheet.cells.get(k) || null;
    let after = cell && (cell.v != null || cell.f || cell.s || cell.note || cell.link) ? cell : null;
    if (after && after.v === undefined) after = { ...after, v: null };
    if (before === after) return;
    const tx = this.tx;
    const snap = tx && tx.snaps.get(sheet);
    if (tx && !(snap && snap.cells)) tx.cells.push([sheet, k, before, after]);
    if (after) sheet.cells.set(k, after); else sheet.cells.delete(k);
    sheet.version++;
    if (tx) {
      tx.seeds.push({ sheet, r, c });
      for (const a of this.calc.spillAnchorsAt(sheet, r, c)) tx.seeds.push({ sheet, r: R(a), c: C(a) });
    }
    if (!tx || !tx.structural) this.calc.link(sheet, k, after);
  }
  patchCell(sheet, r, c, patch) {
    const old = sheet.cells.get(K(r, c));
    const n = { ...(old || {}), ...patch };
    for (const x of Object.keys(patch)) if (patch[x] === undefined) delete n[x];
    if ('v' in patch || 'f' in patch) delete n._si;
    this.setCell(sheet, r, c, n);
  }

  /* ---------- typing ---------- */
  /** what the user typed into a cell (already without IME junk). Returns the new cell. */
  input(sheet, r, c, text, { asFormula } = {}) {
    const old = sheet.cells.get(K(r, c)) || {};
    const s = String(text ?? '');
    let cell;
    if ((s.startsWith('=') && s.length > 1) || asFormula) {
      let f = toInternal(s.replace(/^=/, ''), this.loc);
      // close open parentheses like Excel does (=SUM(A1:A3 → =SUM(A1:A3))
      let depth = 0; let q = false;
      for (const ch of f) { if (ch === '"') q = !q; else if (!q && ch === '(') depth++; else if (!q && ch === ')') depth--; }
      if (depth > 0 && !q) f += ')'.repeat(depth);
      cell = { ...old, f, v: null };
      delete cell._si;
      if (old.s && old.s.nf === '@') { cell = { ...old, v: s, f: undefined }; delete cell.f; }
    } else {
      if (old.s && old.s.nf === '@') cell = { ...old, v: s === '' ? null : s };
      else {
        const p = parseInput(s, this.loc);
        cell = { ...old, v: p.v };
        if (p.fmt && (!old.s || !old.s.nf || (isDateFormat(p.fmt) !== isDateFormat(old.s.nf) && !(isDateFormat(p.fmt) && isDateFormat(old.s.nf))))) cell.s = this.styles.merge(old.s, { nf: p.fmt });
      }
      delete cell.f; delete cell._si;
    }
    this.setCell(sheet, r, c, cell);
    return cell;
  }
  /** raw text for the editor (formula with '=' in the user's locale, or the value) */
  editText(sheet, r, c, toDisplay) {
    const cell = sheet.cells.get(K(r, c));
    if (!cell) return '';
    if (cell.f) return `=${toDisplay ? toDisplay(cell.f) : cell.f}`;
    const v = cell.v;
    if (v == null) return '';
    if (typeof v === 'string') return (cell.s && cell.s.nf === '@') || !/^(=|[-+]?\d|TRUE$|FALSE$|WAHR$|FALSCH$|#)/i.test(v) ? v : `'${v}`;
    if (typeof v === 'boolean') return v ? this.loc.trueStr : this.loc.falseStr;
    if (isErr(v)) return localErr(v.code, this.loc);
    const nf = cell.s && cell.s.nf;
    if (nf && isDateFormat(nf)) {
      const hasTime = Math.abs(v - Math.floor(v)) > 1e-9;
      const date = Math.floor(v) !== 0 || !hasTime;
      return formatValue(v, date && hasTime ? `${this.loc.shortDate} ${this.loc.timeSec}` : date ? this.loc.shortDate : this.loc.timeSec, this.loc).text;
    }
    if (nf && /%/.test(nf) && !/"[^"]*%/.test(nf)) return `${String(Number((v * 100).toPrecision(15))).replace('.', this.loc.decimal)}%`;
    return String(Number(v.toPrecision(15))).replace('.', this.loc.decimal);
  }

  /* ---------- clearing & formatting ---------- */
  clear(sheet, ranges, what = 'contents') {
    this.run(what === 'all' ? 'Clear all' : what === 'formats' ? 'Clear formats' : 'Clear contents', () => {
      for (const g of ranges) forRange(sheet, g, (r, c, k) => {
        const cell = sheet.cells.get(k);
        if (!cell) return;
        if (what === 'all') this.setCell(sheet, r, c, null);
        else if (what === 'formats') this.patchCell(sheet, r, c, { s: undefined });
        else if (what === 'notes') this.patchCell(sheet, r, c, { note: undefined });
        else if (what === 'links') this.patchCell(sheet, r, c, { link: undefined });
        else this.patchCell(sheet, r, c, { v: null, f: undefined });
      }, { create: false });
      if (what === 'all' || what === 'formats') {
        this.touch(sheet);
        sheet.cf = sheet.cf.map((rule) => ({ ...rule, ranges: subtractRanges(rule.ranges, ranges) })).filter((rule) => rule.ranges.length);
        if (what === 'all') sheet.merges = sheet.merges.filter((m) => !ranges.some((g) => m.r1 >= g.r1 && m.r2 <= g.r2 && m.c1 >= g.c1 && m.c2 <= g.c2));
      }
    });
  }
  /** patch: object or fn(style) → patch */
  format(sheet, ranges, patch, label = 'Format cells') {
    this.run(label, () => {
      for (const g of ranges) forRange(sheet, g, (r, c, k) => {
        const cell = sheet.cells.get(k);
        const p = typeof patch === 'function' ? patch(cell && cell.s, r, c, cell) : patch;
        if (!p) return;
        const s = this.styles.merge(cell && cell.s, p);
        if (cell && cell.s === s) return;
        if (!cell && !s) return;
        this.setCell(sheet, r, c, { ...(cell || { v: null }), s });
      }, { bounds: this.bounds(sheet) });
    });
  }
  /** border presets: all | outside | inside | none | top | bottom | left | right | thickBox | bottomDouble | thickBottom | topBottom */
  borders(sheet, g, preset, line = { s: 'thin', c: '#000000' }) {
    const L = line;
    const thick = { ...L, s: 'medium' };
    this.format(sheet, [g], (s, r, c) => {
      const bd = { ...((s && s.bd) || {}) };
      const top = r === g.r1; const bot = r === g.r2; const left = c === g.c1; const right = c === g.c2;
      switch (preset) {
        case 'none': return { bd: undefined };
        case 'all': Object.assign(bd, { t: L, b: L, l: L, r: L }); break;
        case 'outside': if (top) bd.t = L; if (bot) bd.b = L; if (left) bd.l = L; if (right) bd.r = L; break;
        case 'thickBox': if (top) bd.t = thick; if (bot) bd.b = thick; if (left) bd.l = thick; if (right) bd.r = thick; break;
        case 'inside': if (!top) bd.t = L; if (!bot) bd.b = L; if (!left) bd.l = L; if (!right) bd.r = L; break;
        case 'top': if (top) bd.t = L; else return null; break;
        case 'bottom': if (bot) bd.b = L; else return null; break;
        case 'left': if (left) bd.l = L; else return null; break;
        case 'right': if (right) bd.r = L; else return null; break;
        case 'bottomDouble': if (bot) bd.b = { ...L, s: 'double' }; else return null; break;
        case 'thickBottom': if (bot) bd.b = thick; else return null; break;
        case 'topBottom': if (top) bd.t = L; if (bot) bd.b = L; if (!top && !bot) return null; break;
        default: return null;
      }
      return { bd };
    }, 'Borders');
  }

  /* ---------- merges ---------- */
  mergeAt(sheet, r, c) {
    for (const m of sheet.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return null;
  }
  merge(sheet, g, { center = false, across = false } = {}) {
    this.run('Merge cells', () => {
      this.touch(sheet);
      const parts = across ? Array.from({ length: g.r2 - g.r1 + 1 }, (_, i) => ({ r1: g.r1 + i, r2: g.r1 + i, c1: g.c1, c2: g.c2 })) : [g];
      for (const p of parts) {
        if (p.r1 === p.r2 && p.c1 === p.c2) continue;
        sheet.merges = sheet.merges.filter((m) => !(m.r1 <= p.r2 && p.r1 <= m.r2 && m.c1 <= p.c2 && p.c1 <= m.c2));
        // keep the top-left value only (Excel does the same after a warning)
        forRange(sheet, p, (r, c) => { if (r === p.r1 && c === p.c1) return; const cell = sheet.cells.get(K(r, c)); if (cell && (cell.v != null || cell.f)) this.patchCell(sheet, r, c, { v: null, f: undefined }); }, { create: false });
        sheet.merges.push({ ...p });
        if (center) { const cell = sheet.cells.get(K(p.r1, p.c1)); this.setCell(sheet, p.r1, p.c1, { ...(cell || { v: null }), s: this.styles.merge(cell && cell.s, { h: 'center' }) }); }
      }
    });
  }
  unmerge(sheet, g) {
    this.run('Unmerge cells', () => { this.touch(sheet); sheet.merges = sheet.merges.filter((m) => !(m.r1 <= g.r2 && g.r1 <= m.r2 && m.c1 <= g.c2 && g.c1 <= m.c2)); });
  }

  /* ---------- formula rewriting across the book ---------- */
  rewriteAll(fn, { names = true, extras = true } = {}) {
    // fn(formula, hereSheet) → new formula
    for (const sh of this.sheets) {
      for (const [k, cell] of sh.cells) {
        if (!cell.f) continue;
        const f = fn(cell.f, sh);
        if (f !== cell.f) this.setCell(sh, R(k), C(k), { ...cell, f });
      }
      if (extras) {
        let touched = false;
        const tch = () => { if (!touched) { this.touch(sh); touched = true; } };
        sh.cf.forEach((rule) => { for (const key of ['f1', 'f2', 'formula']) if (typeof rule[key] === 'string' && rule[key]) { const f = fn(rule[key], sh); if (f !== rule[key]) { tch(); rule[key] = f; } } });
        sh.dv.forEach((rule) => { for (const key of ['f1', 'f2']) if (typeof rule[key] === 'string' && rule[key] && !/^".*"$/.test(rule[key])) { const f = fn(rule[key], sh); if (f !== rule[key]) { tch(); rule[key] = f; } } });
        (sh.sparklines || []).forEach((sg) => sg.items.forEach((it) => { if (it.ref) { const f = fn(it.ref, sh); if (f !== it.ref) { tch(); it.ref = f; sg._dirty = true; } } }));
        (sh.pivots || []).forEach((pv) => { if (pv.source) { const f = fn(pv.source, sh); if (f !== pv.source) { tch(); pv.source = f; pv._dirty = true; } } });
        sh.charts.forEach((ch) => {
          if (ch.range) { const f = fn(ch.range, sh); if (f !== ch.range) { tch(); ch.range = f; ch._dirty = true; } }
          (ch.series || []).forEach((s) => { for (const key of ['nameRef', 'catRef', 'valRef']) if (s[key]) { const f = fn(s[key], sh); if (f !== s[key]) { tch(); s[key] = f; ch._dirty = true; } } });
        });
      }
    }
    if (names) {
      let changed = false;
      const nn = new Map();
      for (const [k, d] of this.names) { const f = fn(d.ref, d.scope || this.sheets[0]); if (f !== d.ref) changed = true; nn.set(k, f === d.ref ? d : { ...d, ref: f }); }
      if (changed) { this.touchBook(); this.names = nn; }
    }
  }

  /* ---------- insert / delete rows & columns ---------- */
  insert(sheet, axis, at, count) { this.shiftAxis(sheet, axis, at, count, count > 0 ? (axis === 'row' ? 'Insert rows' : 'Insert columns') : axis === 'row' ? 'Delete rows' : 'Delete columns'); }
  remove(sheet, axis, at, count) { this.insert(sheet, axis, at, -count); }
  shiftAxis(sheet, axis, at, count, label) {
    this.run(label, () => {
      const tx = this.tx; tx.structural = true;
      for (const sh of this.sheets) this.touch(sh, true);
      const row = axis === 'row';
      const del = count < 0 ? -count : 0;
      const nc = new Map();
      for (const [k, cell] of sheet.cells) {
        let r = R(k); let c = C(k);
        const i = row ? r : c;
        if (del && i >= at && i < at + del) continue;
        let n = i >= at ? i + count : i;
        if (n >= (row ? MAXR : MAXC)) continue;
        if (row) r = n; else c = n;
        nc.set(K(r, c), cell);
      }
      sheet.cells = nc;
      sheet.spill = new Map(); sheet.spillAnchors = new Map(); sheet.spillBlocked = new Map();
      const shiftMap = (m) => { const o = new Map(); for (const [i, v] of m) { if (del && i >= at && i < at + del) continue; o.set(i >= at ? i + count : i, v); } return o; };
      const shiftSet = (s) => { const o = new Set(); for (const i of s) { if (del && i >= at && i < at + del) continue; o.add(i >= at ? i + count : i); } return o; };
      if (row) { sheet.rowH = shiftMap(sheet.rowH); sheet.hiddenRows = shiftSet(sheet.hiddenRows); sheet.filterHidden = shiftSet(sheet.filterHidden); sheet.outlineRows = shiftMap(sheet.outlineRows); } else { sheet.colW = shiftMap(sheet.colW); sheet.hiddenCols = shiftSet(sheet.hiddenCols); sheet.outlineCols = shiftMap(sheet.outlineCols); }
      const adjG = (g) => {
        const lo = row ? 'r1' : 'c1'; const hi = row ? 'r2' : 'c2';
        const n = { ...g };
        if (count > 0) { if (n[lo] >= at) n[lo] += count; if (n[hi] >= at) n[hi] += count; return n; }
        const end = at + del - 1;
        if (n[lo] >= at && n[hi] <= end) return null;
        if (n[lo] > end) n[lo] -= del; else if (n[lo] >= at) n[lo] = at;
        if (n[hi] > end) n[hi] -= del; else if (n[hi] >= at) n[hi] = at - 1;
        return n[hi] < n[lo] ? null : n;
      };
      sheet.merges = sheet.merges.map(adjG).filter((m) => m && (m.r1 !== m.r2 || m.c1 !== m.c2));
      sheet.cf = sheet.cf.map((rule) => ({ ...rule, ranges: rule.ranges.map(adjG).filter(Boolean) })).filter((rule) => rule.ranges.length);
      sheet.dv = sheet.dv.map((rule) => ({ ...rule, ranges: rule.ranges.map(adjG).filter(Boolean) })).filter((rule) => rule.ranges.length);
      if (sheet.filter) { const g = adjG(sheet.filter.range); if (!g) sheet.filter = null; else { if (!row && sheet.filter.cols) { const nc2 = {}; for (const [ci, v] of Object.entries(sheet.filter.cols)) { const i = +ci; if (del && i >= at && i < at + del) continue; nc2[i >= at ? i + count : i] = v; } sheet.filter.cols = nc2; } sheet.filter.range = g; } }
      for (const a of [...sheet.charts.map((x) => x.anchor), ...sheet.images.map((x) => x.anchor), ...(sheet.shapes || []).map((x) => x.anchor)]) { if (!a) continue; const p = row ? 'r' : 'c'; if (a[p] >= at) a[p] = Math.max(at, a[p] + count); }
      sheet.tables = (sheet.tables || []).map((tb) => { const g = adjG(tb.range); return g ? { ...tb, range: g, _dirty: true } : null; }).filter(Boolean);
      sheet.pivots = (sheet.pivots || []).map((pv) => { const g = pv.range && adjG(pv.range); return g ? { ...pv, range: g, at: { r: g.r1, c: g.c1 }, _moved: true } : pv; });
      sheet.sparklines = (sheet.sparklines || []).map((sg) => ({ ...sg, items: sg.items.map((it) => { const g = adjG({ r1: it.r, c1: it.c, r2: it.r, c2: it.c }); return g ? { ...it, r: g.r1, c: g.c1 } : null; }).filter(Boolean) })).filter((sg) => sg.items.length);
      if (sheet.print.area) { const g = parseRange(sheet.print.area); const n = g && adjG(g); sheet.print.area = n ? rangeText({ ...n }, false) : null; }
      if (sheet.freeze && ((row && sheet.freeze.rows > at) || (!row && sheet.freeze.cols > at))) { if (row) sheet.freeze.rows = Math.max(at, sheet.freeze.rows + count); else sheet.freeze.cols = Math.max(at, sheet.freeze.cols + count); }
      const spec = { target: sheet.name, axis, at, count };
      // rewrite formulas everywhere (cells were moved above; rewrite with setCell bypassing diff since snapshots cover it)
      const fx = (f, sh) => adjustFormula(f, { ...spec, here: sh.name });
      this.rewriteAll(fx, { names: true, extras: true });
    });
  }
  /** insert/delete cells with shifting inside the selection's rows/cols */
  shiftCells(sheet, g, dir, insert) {
    // dir: 'down'|'right' for insert, 'up'|'left' for delete
    this.run(insert ? 'Insert cells' : 'Delete cells', () => {
      this.tx.structural = true;
      for (const sh of this.sheets) this.touch(sh, true);
      const vertical = dir === 'down' || dir === 'up';
      const n = vertical ? g.r2 - g.r1 + 1 : g.c2 - g.c1 + 1;
      const nc = new Map();
      const moved = [];
      for (const [k, cell] of sheet.cells) {
        let r = R(k); let c = C(k);
        const inBand = vertical ? c >= g.c1 && c <= g.c2 : r >= g.r1 && r <= g.r2;
        if (inBand) {
          const i = vertical ? r : c; const start = vertical ? g.r1 : g.c1;
          if (insert) { if (i >= start) { if (vertical) r += n; else c += n; } } else if (i >= start && i < start + n) continue; else if (i >= start + n) { if (vertical) r -= n; else c -= n; }
        }
        if (r >= MAXR || c >= MAXC) continue;
        nc.set(K(r, c), cell);
        moved.push(k);
      }
      sheet.cells = nc;
      sheet.spill = new Map(); sheet.spillAnchors = new Map(); sheet.spillBlocked = new Map();
      // references: the band from the start (or end) moves
      const band = vertical ? { r1: insert ? g.r1 : g.r2 + 1, r2: MAXR - 1, c1: g.c1, c2: g.c2 } : { r1: g.r1, r2: g.r2, c1: insert ? g.c1 : g.c2 + 1, c2: MAXC - 1 };
      const dr = vertical ? (insert ? n : -n) : 0; const dc = vertical ? 0 : insert ? n : -n;
      for (const sh of this.sheets) for (const [k, cell] of sh.cells) if (cell.f) { const f = moveFormulaRefs(cell.f, { here: sh.name, srcSheet: sheet.name, src: band, dr, dc }); if (f !== cell.f) sh.cells.set(k, { ...cell, f }); }
    });
  }

  /* ---------- row/col sizes & visibility ---------- */
  setSizes(sheet, axis, idxs, px) {
    this.run(axis === 'row' ? 'Row height' : 'Column width', () => {
      this.touch(sheet);
      const m = axis === 'row' ? sheet.rowH : sheet.colW;
      for (const i of idxs) { if (px == null) m.delete(i); else m.set(i, Math.max(2, Math.round(px))); }
    });
  }
  setHidden(sheet, axis, idxs, hidden) {
    this.run(hidden ? 'Hide' : 'Unhide', () => {
      this.touch(sheet);
      const s = axis === 'row' ? sheet.hiddenRows : sheet.hiddenCols;
      for (const i of idxs) { if (hidden) s.add(i); else s.delete(i); }
      if (!hidden && axis === 'row') for (const i of idxs) sheet.filterHidden.delete(i);
    });
  }
  setOutline(sheet, axis, g, delta) {
    this.run(delta > 0 ? 'Group' : 'Ungroup', () => {
      this.touch(sheet);
      const m = axis === 'row' ? sheet.outlineRows : sheet.outlineCols;
      const [a, b] = axis === 'row' ? [g.r1, g.r2] : [g.c1, g.c2];
      for (let i = a; i <= b; i++) { const l = Math.max(0, Math.min(7, (m.get(i) || 0) + delta)); if (l) m.set(i, l); else m.delete(i); }
    });
  }
  setSheetProps(sheet, patch, label = 'Sheet') {
    this.run(label, () => { this.touch(sheet); Object.assign(sheet, patch); });
  }

  /* ---------- sheets ---------- */
  uniqueSheetName(base) {
    let n = base; let i = 2;
    while (this.sheetByName(n)) n = `${base.replace(/\s*\(\d+\)$/, '')} (${i++})`;
    return n;
  }
  nextSheetName(prefix = 'Sheet') { let i = this.sheets.length + 1; while (this.sheetByName(`${prefix}${i}`)) i++; return `${prefix}${i}`; }
  addSheet(name, at = this.sheets.length) {
    return this.run('Insert sheet', () => {
      this.touchBook();
      const sh = newSheet(name || this.nextSheetName());
      if (this.settings.gridlines === false) sh.gridlines = false;
      this.sheets.splice(at, 0, sh);
      this.active = at;
      return sh;
    });
  }
  renameSheet(sheet, name) {
    name = String(name).trim().slice(0, 31);
    if (!name || /[\\/?*[\]:]/.test(name) || (this.sheetByName(name) && this.sheetByName(name) !== sheet)) return false;
    const old = sheet.name;
    this.run('Rename sheet', () => {
      this.touchBook();
      this.tx.structural = true;
      this.touch(sheet);
      sheet.name = name;
      this.rewriteAll((f) => renameSheetInFormula(f, old, name));
    });
    return true;
  }
  duplicateSheet(sheet) {
    return this.run('Duplicate sheet', () => {
      this.touchBook();
      const sh = newSheet(this.uniqueSheetName(sheet.name));
      Object.assign(sh, cloneSheetProps(sheet), { name: sh.name });
      sh.cells = new Map(sheet.cells);
      sh.charts = sh.charts.map((c) => ({ ...c, id: uid('c'), _orig: null }));
      sh.images = sh.images.map((i) => ({ ...i, id: uid('i'), _orig: null }));
      const i = this.sheets.indexOf(sheet) + 1;
      this.sheets.splice(i, 0, sh);
      this.active = i;
      return sh;
    });
  }
  deleteSheet(sheet) {
    if (this.sheets.filter((s) => !s.hidden).length <= 1 && !sheet.hidden) return false;
    this.run('Delete sheet', () => {
      this.touchBook();
      this.tx.structural = true;
      const name = sheet.name.toLowerCase();
      const i = this.sheets.indexOf(sheet);
      this.sheets.splice(i, 1);
      this.rewriteAll((f) => (tokenize(f).some((t) => t.t === 'ref' && t.ref.sheet && t.ref.sheet.toLowerCase() === name) ? tokenize(f).map((t) => (t.t === 'ref' && t.ref.sheet && t.ref.sheet.toLowerCase() === name ? '#REF!' : t.text)).join('') : f));
      for (const [k, d] of [...this.names]) if (d.scope === sheet) this.names.delete(k);
      this.active = Math.max(0, Math.min(this.active >= i ? this.active - 1 : this.active, this.sheets.length - 1));
      if (this.sheets[this.active].hidden) this.active = this.sheets.findIndex((s) => !s.hidden);
    });
    return true;
  }
  moveSheet(from, to) {
    if (from === to) return;
    this.run('Move sheet', () => {
      this.touchBook();
      const [sh] = this.sheets.splice(from, 1);
      this.sheets.splice(to, 0, sh);
      this.active = to;
    });
  }

  /* ---------- names ---------- */
  setName(name, ref, scope = null, comment = '', oldKey = null) {
    if (!/^[A-Za-z_\\\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^R\d*C\d*$/i.test(name)) return false;
    this.run('Name', () => {
      this.touchBook();
      this.tx.structural = true;
      this.names = new Map(this.names);
      if (oldKey) this.names.delete(oldKey);
      this.names.set(`${name.toUpperCase()}|${scope ? scope.id : ''}`, { name, ref: toInternal(String(ref).replace(/^=/, ''), this.loc), scope, comment });
    });
    return true;
  }
  deleteName(k) { this.run('Delete name', () => { this.touchBook(); this.tx.structural = true; this.names = new Map(this.names); this.names.delete(k); }); }

  /* ---------- copy / paste ---------- */
  /** copy a rectangular block (cells + merges) */
  copyBlock(sheet, g) {
    const rows = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const row = [];
      for (let c = g.c1; c <= g.c2; c++) {
        const cell = sheet.cells.get(K(r, c));
        const sp = !cell && sheet.spill.get(K(r, c));
        row.push(cell ? { ...cell } : sp ? { v: sp.v, spilled: true } : null);
      }
      rows.push(row);
    }
    const merges = sheet.merges.filter((m) => m.r1 >= g.r1 && m.r2 <= g.r2 && m.c1 >= g.c1 && m.c2 <= g.c2).map((m) => ({ r1: m.r1 - g.r1, r2: m.r2 - g.r1, c1: m.c1 - g.c1, c2: m.c2 - g.c1 }));
    const colW = []; for (let c = g.c1; c <= g.c2; c++) colW.push(sheet.colW.get(c));
    return { sheetId: sheet.id, sheetName: sheet.name, g: { ...g }, rows, merges, colW };
  }
  /**
   * paste a block at (r, c). mode: { what: 'all'|'values'|'formulas'|'formats'|'noBorders'|'colWidths', transpose, skipBlanks, cut }
   * fills a larger selection with repeats when it is a multiple of the block (like Excel)
   */
  paste(sheet, target, blk, mode = {}) {
    const what = mode.what || 'all';
    let rows = blk.rows;
    let merges = blk.merges;
    if (mode.transpose) { rows = rows[0].map((_, j) => rows.map((row) => row[j])); merges = merges.map((m) => ({ r1: m.c1, r2: m.c2, c1: m.r1, c2: m.r2 })); }
    const h = rows.length; const w = rows[0].length;
    const th = target.r2 - target.r1 + 1; const tw = target.c2 - target.c1 + 1;
    const ry = th % h === 0 && th > h ? th / h : 1; const rx = tw % w === 0 && tw > w ? tw / w : 1;
    const src = blk.sheetId ? this.sheetById(blk.sheetId) : null;
    this.run(mode.cut ? 'Move' : 'Paste', () => {
      if (mode.cut && src) {
        // cut: clear the source first, then rewrite references to the moved block everywhere
        this.tx.structural = true;
        for (const sh of this.sheets) this.touch(sh, true);
        for (let r = blk.g.r1; r <= blk.g.r2; r++) for (let c = blk.g.c1; c <= blk.g.c2; c++) src.cells.delete(K(r, c));
        src.merges = src.merges.filter((m) => !(m.r1 >= blk.g.r1 && m.r2 <= blk.g.r2 && m.c1 >= blk.g.c1 && m.c2 <= blk.g.c2));
      }
      for (let y = 0; y < ry; y++) for (let x = 0; x < rx; x++) {
        const r0 = target.r1 + y * h; const c0 = target.c1 + x * w;
        if (what === 'all' || what === 'formats') this.touch(sheet);
        if (what === 'all' || what === 'formats' || what === 'noBorders') sheet.merges = sheet.merges.filter((m) => !(m.r1 <= r0 + h - 1 && r0 <= m.r2 && m.c1 <= c0 + w - 1 && c0 <= m.c2));
        for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
          const r = r0 + i; const c = c0 + j;
          if (r >= MAXR || c >= MAXC) continue;
          const sc = rows[i][j];
          if (mode.skipBlanks && (!sc || (sc.v == null && !sc.f))) continue;
          const old = sheet.cells.get(K(r, c));
          let n;
          const srcR = blk.g.r1 + (mode.transpose ? j : i); const srcC = blk.g.c1 + (mode.transpose ? i : j);
          const shiftedF = sc && sc.f && !mode.cut ? shiftFormula(sc.f, r - srcR, c - srcC) : sc && sc.f;
          switch (what) {
            case 'values': n = { ...(old || {}), v: sc ? sc.v : null }; delete n.f; delete n._si; break;
            case 'formulas': n = { ...(old || {}), v: sc ? sc.v : null, f: shiftedF || undefined }; if (!n.f) delete n.f; delete n._si; break;
            case 'formats': n = { ...(old || { v: null }), s: sc ? sc.s : undefined }; break;
            case 'noBorders': n = sc ? { ...sc, f: shiftedF, s: sc.s && sc.s.bd ? this.styles.merge(sc.s, { bd: undefined }) : sc.s } : null; break;
            case 'colWidths': n = old; break;
            default: n = sc ? { ...sc, f: shiftedF } : null;
          }
          if (n) { delete n.spilled; if (!n.f) delete n.f; if (mode.cut && src && src !== sheet) delete n._si; }
          if (n && n.f && !mode.cut) delete n._si;
          this.setCell(sheet, r, c, n);
        }
        if (what === 'all' || what === 'formats') for (const m of merges) sheet.merges.push({ r1: m.r1 + r0, r2: m.r2 + r0, c1: m.c1 + c0, c2: m.c2 + c0 });
        if (what === 'colWidths' || (what === 'all' && mode.colWidths)) { this.touch(sheet); blk.colW.forEach((wd, j) => { if (wd) sheet.colW.set(c0 + j, wd); }); }
      }
      if (mode.cut && src) {
        const dr = target.r1 - blk.g.r1; const dc = target.c1 - blk.g.c1;
        for (const sh of this.sheets) for (const [k, cell] of sh.cells) if (cell.f) {
          const f = moveFormulaRefs(cell.f, { here: sh.name, srcSheet: src.name, src: blk.g, dr, dc, dstSheet: sheet !== src ? sheet.name : null });
          if (f !== cell.f) sh.cells.set(k, { ...cell, f });
        }
      }
    });
    return { r1: target.r1, c1: target.c1, r2: target.r1 + h * ry - 1, c2: target.c1 + w * rx - 1 };
  }
  /** plain text grid (from the clipboard) at (r, c) – values are parsed like typed input */
  pasteText(sheet, r0, c0, grid, { transpose = false } = {}) {
    const g = transpose ? grid[0].map((_, j) => grid.map((row) => row[j] ?? '')) : grid;
    this.run('Paste', () => { g.forEach((row, i) => row.forEach((t, j) => { if (r0 + i < MAXR && c0 + j < MAXC) this.input(sheet, r0 + i, c0 + j, t); })); });
    return { r1: r0, c1: c0, r2: r0 + g.length - 1, c2: c0 + Math.max(...g.map((x) => x.length)) - 1 };
  }

  /* ---------- fill (handle, Ctrl+D/R) ---------- */
  /** src: the selected block; dst: the full extended block (contains src). dir from src to the new cells. */
  /** what the last cell of a fill drag would show (the fill handle's preview), without changing anything */
  fillPreview(sheet, src, dst, toDisplay) {
    const down = dst.r2 > src.r2; const up = dst.r1 < src.r1; const right = dst.c2 > src.c2; const left = dst.c1 < src.c1;
    const vertical = down || up;
    if (!vertical && !right && !left) return null;
    const line = vertical ? src.c1 : src.r1;
    const n = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
    const cells = []; const items = [];
    for (let i = 0; i < n; i++) { const r = vertical ? src.r1 + i : line; const c = vertical ? line : src.c1 + i; const cell = sheet.cells.get(K(r, c)); cells.push({ cell, r, c }); items.push({ v: cell ? cell.v : null, date: !!(cell && cell.s && cell.s.nf && isDateFormat(cell.s.nf)) }); }
    const reverse = up || left;
    const count = vertical ? (down ? dst.r2 - src.r2 : src.r1 - dst.r1) : right ? dst.c2 - src.c2 : src.c1 - dst.c1;
    if (count <= 0) return null;
    const k = count - 1;
    const r = vertical ? (down ? src.r2 + count : src.r1 - count) : line; const c = vertical ? line : right ? src.c2 + count : src.c1 - count;
    const s = cells[reverse ? n - 1 - (k % n) : k % n];
    if (s.cell && s.cell.f) return `=${toDisplay ? toDisplay(shiftFormula(s.cell.f, r - s.r, c - s.c)) : shiftFormula(s.cell.f, r - s.r, c - s.c)}`;
    const values = fillSeries(reverse ? items.slice().reverse() : items, count, { partsFromSerial, serialFromParts });
    const v = values ? values[k] : s.cell ? s.cell.v : null;
    if (v == null) return null;
    return formatValue(v, (s.cell && s.cell.s && s.cell.s.nf) || 'General', this.loc).text;
  }
  fill(sheet, src, dst, { copy = false } = {}) {
    const down = dst.r2 > src.r2; const up = dst.r1 < src.r1; const right = dst.c2 > src.c2; const left = dst.c1 < src.c1;
    const vertical = down || up;
    this.run('Fill', () => {
      const lines = vertical ? [src.c1, src.c2] : [src.r1, src.r2];
      for (let line = lines[0]; line <= lines[1]; line++) {
        const items = []; const cells = [];
        const n = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
        for (let i = 0; i < n; i++) {
          const r = vertical ? src.r1 + i : line; const c = vertical ? line : src.c1 + i;
          const cell = sheet.cells.get(K(r, c));
          cells.push({ cell, r, c });
          items.push({ v: cell ? cell.v : null, date: !!(cell && cell.s && cell.s.nf && isDateFormat(cell.s.nf)) });
        }
        const reverse = up || left;
        const count = vertical ? (down ? dst.r2 - src.r2 : src.r1 - dst.r1) : right ? dst.c2 - src.c2 : src.c1 - dst.c1;
        const hasFormula = cells.some((x) => x.cell && x.cell.f);
        const seq = reverse ? items.slice().reverse() : items;
        const values = hasFormula || copy ? null : fillSeries(seq, count, { copy, partsFromSerial, serialFromParts });
        for (let k = 0; k < count; k++) {
          const off = k + 1;
          const r = vertical ? (down ? src.r2 + off : src.r1 - off) : line;
          const c = vertical ? line : right ? src.c2 + off : src.c1 - off;
          const srcIdx = reverse ? (n - 1 - (k % n)) : k % n;
          const s = cells[srcIdx];
          const cell = s.cell;
          let nc;
          if (cell && cell.f) nc = { ...cell, f: shiftFormula(cell.f, r - s.r, c - s.c), v: null };
          else nc = cell ? { ...cell, v: values ? values[k] : cell.v } : values && values[k] != null ? { v: values[k] } : null;
          if (nc) { delete nc._si; delete nc.note; }
          this.setCell(sheet, r, c, nc);
        }
      }
    });
  }

  /* ---------- move by drag ---------- */
  moveRange(sheet, g, r, c, { copy = false } = {}) {
    const blk = this.copyBlock(sheet, g);
    const h = g.r2 - g.r1; const w = g.c2 - g.c1;
    return this.paste(sheet, { r1: r, c1: c, r2: r + h, c2: c + w }, blk, { cut: !copy });
  }

  /* ---------- sort ---------- */
  /** keys: [{ col (absolute index), desc, by: 'value'|'color', color }], header: first row excluded */
  sort(sheet, g, keys, { header = false, byCols = false } = {}) {
    this.run('Sort', () => {
      this.tx.structural = true;
      for (const sh of this.sheets) this.touch(sh, true);
      const r1 = g.r1 + (header && !byCols ? 1 : 0); const c1 = g.c1 + (header && byCols ? 1 : 0);
      const lines = [];
      const n = byCols ? g.c2 - c1 + 1 : g.r2 - r1 + 1;
      for (let i = 0; i < n; i++) {
        const cells = [];
        if (byCols) for (let r = g.r1; r <= g.r2; r++) cells.push(sheet.cells.get(K(r, c1 + i)) || null); else for (let c = g.c1; c <= g.c2; c++) cells.push(sheet.cells.get(K(r1 + i, c)) || null);
        lines.push({ i, cells });
      }
      const val = (line, key) => { const idx = byCols ? key.col - g.r1 : key.col - g.c1; const cell = line.cells[idx]; return cell ? cell.v : null; };
      lines.sort((a, b) => {
        for (const key of keys) {
          let x; let y;
          if (key.by === 'color') { const idx = byCols ? key.col - g.r1 : key.col - g.c1; const fa = a.cells[idx] && a.cells[idx].s && a.cells[idx].s.fill; const fb = b.cells[idx] && b.cells[idx].s && b.cells[idx].s.fill; x = fa === key.color ? 0 : 1; y = fb === key.color ? 0 : 1; if (x !== y) return x - y; continue; }
          x = val(a, key); y = val(b, key);
          const blank = (v) => v == null || v === '';
          if (blank(x) || blank(y)) { if (blank(x) && blank(y)) continue; return blank(x) ? 1 : -1; } // blanks always last
          const cmp = cmpSort(x, y);
          if (cmp) return key.desc ? -cmp : cmp;
        }
        return a.i - b.i;
      });
      lines.forEach((line, ni) => {
        line.cells.forEach((cell, j) => {
          const r = byCols ? g.r1 + j : r1 + ni; const c = byCols ? c1 + ni : g.c1 + j;
          const oldR = byCols ? r : r1 + line.i; const oldC = byCols ? c1 + line.i : c;
          let nc = cell;
          if (cell && cell.f) nc = { ...cell, f: shiftFormula(cell.f, r - oldR, c - oldC) };
          if (nc) sheet.cells.set(K(r, c), nc); else sheet.cells.delete(K(r, c));
        });
      });
    });
  }

  /* ---------- data tools ---------- */
  removeDuplicates(sheet, g, cols, header) {
    const r1 = g.r1 + (header ? 1 : 0);
    const seen = new Set(); const keep = []; let removed = 0;
    for (let r = r1; r <= g.r2; r++) {
      const k = cols.map((c) => { const v = this.value(sheet, r, c); return typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${v}`; }).join('\u0000');
      if (seen.has(k)) { removed++; continue; }
      seen.add(k); keep.push(r);
    }
    if (!removed) return 0;
    this.run('Remove duplicates', () => {
      this.tx.structural = true;
      this.touch(sheet, true);
      const rows = keep.map((r) => { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(sheet.cells.get(K(r, c)) || null); return row; });
      for (let r = r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) sheet.cells.delete(K(r, c));
      rows.forEach((row, i) => row.forEach((cell, j) => { if (cell) sheet.cells.set(K(r1 + i, g.c1 + j), cell); }));
    });
    return removed;
  }
  textToColumns(sheet, g, { delimiters = [','], other = '', treatConsecutive = false, quote = '"', fixed = null }) {
    this.run('Text to columns', () => {
      for (let r = g.r1; r <= g.r2; r++) {
        const v = this.value(sheet, r, g.c1);
        if (typeof v !== 'string') continue;
        let parts;
        if (fixed && fixed.length) { parts = []; let p = 0; for (const w of [...fixed, Infinity]) { parts.push(v.slice(p, p + w)); p += w; } parts = parts.filter((x, i) => i < fixed.length || x); } else {
          const ds = [...delimiters, ...(other ? [other] : [])].map((d) => (d === 'tab' ? '\t' : d === 'space' ? ' ' : d));
          parts = []; let cur = ''; let q = false;
          for (let i = 0; i < v.length; i++) {
            const ch = v[i];
            if (ch === quote && quote) { q = !q; continue; }
            if (!q && ds.includes(ch)) { if (!(treatConsecutive && cur === '' && parts.length && ds.includes(v[i - 1]))) parts.push(cur); cur = ''; continue; }
            cur += ch;
          }
          parts.push(cur);
        }
        parts.forEach((p, j) => this.input(sheet, r, g.c1 + j, p.trim() === p ? p : p.trim()));
      }
    });
  }

  /* ---------- find & replace ---------- */
  /** opts: { query, matchCase, whole, inFormulas, allSheets, sheet, regex } → [{ sheet, r, c }] in row order */
  findAll(opts) {
    const { query, matchCase, whole, inFormulas } = opts;
    if (!query) return [];
    const esc = query.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*');
    const re = new RegExp(whole ? `^${esc}$` : esc, matchCase ? '' : 'i');
    const out = [];
    const sheets = opts.allSheets ? this.sheets : [opts.sheet];
    for (const sh of sheets) {
      const keys = [...sh.cells.keys()].sort((a, b) => a - b);
      for (const k of keys) {
        const cell = sh.cells.get(k);
        const t = inFormulas && cell.f ? `=${cell.f}` : cell.v == null ? '' : this.text(sh, R(k), C(k));
        if (t && re.test(t)) out.push({ sheet: sh, r: R(k), c: C(k) });
      }
    }
    return out;
  }
  replaceAll(opts, replacement, only = null) {
    const hits = only || this.findAll(opts);
    const esc = opts.query.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
    const re = new RegExp(opts.whole ? `^${esc}$` : esc, opts.matchCase ? 'g' : 'gi');
    let n = 0;
    this.run('Replace', () => {
      for (const h of hits) {
        const cell = h.sheet.cells.get(K(h.r, h.c));
        if (!cell) continue;
        if (cell.f) { if (!opts.inFormulas) continue; const f = cell.f.replace(re, replacement); if (f !== cell.f) { this.setCell(h.sheet, h.r, h.c, { ...cell, f, v: null }); n++; } continue; }
        if (typeof cell.v === 'string') { const v = cell.v.replace(re, replacement); if (v !== cell.v) { this.input(h.sheet, h.r, h.c, v); n++; } } else if (typeof cell.v === 'number') { const t = this.editText(h.sheet, h.r, h.c); const v = t.replace(re, replacement); if (v !== t) { this.input(h.sheet, h.r, h.c, v); n++; } }
      }
    });
    return n;
  }

  /* ---------- AutoFilter ---------- */
  setFilter(sheet, g) {
    this.run(g ? 'Filter' : 'Remove filter', () => {
      this.touch(sheet);
      sheet.filter = g ? { range: { ...g }, cols: {} } : null;
      sheet.filterHidden = new Set();
    });
  }
  /** crit: { values: string[] | null (allowed display texts, '' = blanks), cond: { op, v, join, op2, v2 } | null, sort } */
  setFilterColumn(sheet, col, crit) {
    this.run('Filter', () => {
      this.touch(sheet);
      if (!sheet.filter) return;
      const cols = { ...sheet.filter.cols };
      if (crit && (crit.values || crit.cond)) cols[col] = crit; else delete cols[col];
      sheet.filter = { ...sheet.filter, cols };
      this.applyFilter(sheet);
    });
  }
  applyFilter(sheet) {
    const f = sheet.filter;
    sheet.filterHidden = new Set();
    if (!f) return;
    const tests = Object.entries(f.cols).map(([c, crit]) => [Number(c), filterTest(crit, this.loc)]);
    if (!tests.length) return;
    const r2 = Math.min(f.range.r2, Math.max(f.range.r1, this.bounds(sheet).rows - 1));
    for (let r = f.range.r1 + 1; r <= r2; r++) {
      for (const [c, test] of tests) if (!test(this.value(sheet, r, c), this.text(sheet, r, c))) { sheet.filterHidden.add(r); break; }
    }
  }
}

function filterTest(crit, loc) {
  const set = crit.values ? new Set(crit.values.map((x) => String(x).toLowerCase())) : null;
  const cond = crit.cond;
  const one = (op, v) => {
    if (op === 'top' || op === 'bottom') return () => true;
    const pat = { eq: `=${v}`, ne: `<>${v}`, gt: `>${v}`, lt: `<${v}`, ge: `>=${v}`, le: `<=${v}`, begins: `=${v}*`, ends: `=*${v}`, contains: `=*${v}*`, notContains: `<>*${v}*` }[op];
    return pat ? criteria(pat, loc) : () => true;
  };
  const c1 = cond && one(cond.op, cond.v); const c2 = cond && cond.op2 && one(cond.op2, cond.v2);
  return (v, text) => {
    if (set && !set.has(String(text).toLowerCase())) return false;
    if (c1) { const a = c1(v); if (c2) { const b = c2(v); return cond.join === 'or' ? a || b : a && b; } return a; }
    return true;
  };
}

/** ranges minus cut-out ranges (coarse: removes fully covered parts only) */
function subtractRanges(list, cut) {
  const out = [];
  for (const g of list) { if (cut.some((x) => x.r1 <= g.r1 && x.r2 >= g.r2 && x.c1 <= g.c1 && x.c2 >= g.c2)) continue; out.push(g); }
  return out;
}

export { K as cellKey, R as rowOfKey, C as colOfKey, compare, quoteSheet, ERR };
