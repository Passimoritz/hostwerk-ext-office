/*
 * Word tables in the editor: the attributes of io/tables.js on table / row / cell, a table view that lays the table out
 * like Word (column widths from the grid, fixed vs. autofit, preferred width, indent, alignment, table style look) and
 * the table tools of the ribbon (style, borders, shading, margins, alignment, row height, header rows, sort, split,
 * distribute, AutoFit, text ↔ table, formulas).
 */
import { Extension } from '@tiptap/core';
import { TableView } from '@tiptap/extension-table';
import { TextSelection } from '@tiptap/pm/state';
import { CellSelection, TableMap } from '@tiptap/pm/tables';
import { tableLook, tableWidthCss, cellCss, rowCss, builtinTableStyles, wordTableCss } from '../io/tables.js';

// the table-style CSS (shared with the HTML/PDF export) – once per page
if (typeof document !== 'undefined' && !document.getElementById('of-docs-wt-css')) { const st = document.createElement('style'); st.id = 'of-docs-wt-css'; st.textContent = wordTableCss('.of-docs-pm ') + wordTableCss('.of-docs-tstyles '); document.head.append(st); }

const keep = () => ({ default: null, parseHTML: () => null, renderHTML: () => ({}) });

export const TableFormat = Extension.create({
  name: 'tableFormat',
  addGlobalAttributes() {
    return [
      { types: ['table'], attributes: { layout: keep(), width: keep(), indent: keep(), align: keep(), borders: keep(), cellMargins: keep(), styleId: keep(), look: keep(), tstyle: keep() } },
      // one attribute renders the style of all (tiptap hands every renderHTML the whole attribute object)
      { types: ['tableRow'], attributes: { height: { default: null, parseHTML: () => null, renderHTML: (a) => { const s = rowCss(a); return s ? { style: s } : {}; } }, heightRule: keep(), cantSplit: keep(), ox: keep() } },
      { types: ['tableCell', 'tableHeader'], attributes: { vAlign: { default: null, parseHTML: () => null, renderHTML: (a) => { const s = cellCss(a); return s ? { style: s } : {}; } }, borders: keep(), margins: keep(), textDir: keep(), noWrap: keep(), ox: keep() } },
    ];
  },
});

/** grid columns (px) of a ProseMirror table node */
function gridCols(node) {
  const cols = [];
  const row = node.firstChild;
  if (!row) return cols;
  row.forEach((c) => { const a = c.attrs; for (let k = 0; k < (a.colspan || 1); k++) cols.push((a.colwidth && a.colwidth[k]) || null); });
  return cols;
}

/** the table's DOM: prosemirror-tables' view (column resizing keeps working) + Word's width rules and style look */
export class OfTableView extends TableView {
  constructor(node, cellMinWidth) { super(node, cellMinWidth); this.applyWord(node); }
  update(node) { if (!super.update(node)) return false; this.applyWord(node); return true; }
  applyWord(node) {
    const a = node.attrs || {};
    const t = this.table;
    const look = tableLook(a);
    const word = !!(look || a.layout || a.width || a.indent || a.align);
    t.className = look ? look.cls.join(' ') : '';
    this.dom.className = `tableWrapper${word ? ' of-wtw' : ''}`;
    for (const p of [...t.style]) if (p.startsWith('--')) t.style.removeProperty(p);
    if (look) for (const [k, v] of Object.entries(look.vars)) t.style.setProperty(k, v);
    if (!word) return;
    const w = tableWidthCss(a, gridCols(node));
    for (const p of ['width', 'min-width', 'margin-left', 'margin-right', 'table-layout']) t.style.removeProperty(p);
    for (const [k, v] of Object.entries(w.table)) t.style.setProperty(k, v);
    const cols = [...this.colgroup.children];
    cols.forEach((c, i) => { c.style.removeProperty('min-width'); if (w.cols[i]) c.style.width = w.cols[i]; });
  }
}

/* ---------------- commands ---------------- */
const findTable = (state) => {
  const $p = state.selection.$from;
  for (let d = $p.depth; d > 0; d--) if ($p.node(d).type.name === 'table') return { node: $p.node(d), pos: $p.before(d), depth: d };
  return null;
};
/** positions of the selected cells (a cell selection, or the cell of the cursor) */
function selectedCells(state) {
  const sel = state.selection;
  if (sel instanceof CellSelection) { const out = []; sel.forEachCell((n, pos) => out.push({ node: n, pos })); return out; }
  const $p = sel.$from;
  for (let d = $p.depth; d > 0; d--) { const n = $p.node(d); if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') return [{ node: n, pos: $p.before(d) }]; }
  return [];
}
function selectedRows(state) {
  const t = findTable(state); if (!t) return [];
  const cells = selectedCells(state);
  const rows = new Map();
  for (const c of cells) { const $c = state.doc.resolve(c.pos); const rowPos = $c.before($c.depth); rows.set(rowPos, $c.node($c.depth)); }
  return [...rows].map(([pos, node]) => ({ pos, node }));
}

export const TableTools = Extension.create({
  name: 'tableTools',
  addCommands() {
    return {
      setTableAttrs: (attrs) => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        if (dispatch) dispatch(tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, ...attrs }));
        return true;
      },
      /** Word table style from the gallery (or the document's own) – spec for the display */
      setTableStyle: (id, spec) => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        const def = builtinTableStyles()[id];
        if (dispatch) dispatch(tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, styleId: id, tstyle: spec || (def && def.spec) || { cond: {} }, borders: null }));
        return true;
      },
      setTableLook: (patch) => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        const look = { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true, ...(t.node.attrs.look || {}), ...patch };
        if (dispatch) dispatch(tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, look, tstyle: t.node.attrs.tstyle || { cond: {} } }));
        return true;
      },
      setCellsAttrs: (attrs) => ({ state, tr, dispatch }) => {
        const cells = selectedCells(state); if (!cells.length) return false;
        for (const c of cells) tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, ...(typeof attrs === 'function' ? attrs(c.node.attrs) : attrs) });
        if (dispatch) dispatch(tr);
        return true;
      },
      setRowsAttrs: (attrs) => ({ state, tr, dispatch }) => {
        const rows = selectedRows(state); if (!rows.length) return false;
        for (const r of rows) tr.setNodeMarkup(r.pos, undefined, { ...r.node.attrs, ...attrs });
        if (dispatch) dispatch(tr);
        return true;
      },
      /** header rows: rows 0..n-1 become header cells (repeated on every page), the rest normal cells */
      setHeaderRows: (n) => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        const { schema } = state;
        let pos = t.pos + 1;
        t.node.forEach((row, _o, ri) => {
          let cp = pos + 1;
          row.forEach((cell) => {
            const want = ri < n ? schema.nodes.tableHeader : schema.nodes.tableCell;
            if (cell.type !== want) tr.setNodeMarkup(cp, want, cell.attrs);
            cp += cell.nodeSize;
          });
          pos += row.nodeSize;
        });
        if (dispatch) dispatch(tr);
        return true;
      },
      /** same width for the selected (or all) columns, the total kept */
      distributeColumns: () => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        const map = TableMap.get(t.node);
        const cols = gridCols(t.node);
        const total = cols.reduce((s, w) => s + (w || 0), 0) || 600;
        const each = Math.round((total / map.width) * 100) / 100;
        t.node.descendants((n, p) => { if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') tr.setNodeMarkup(t.pos + 1 + p, undefined, { ...n.attrs, colwidth: Array(n.attrs.colspan || 1).fill(each) }); return n.type.name === 'table' || n.type.name === 'tableRow'; });
        if (dispatch) dispatch(tr);
        return true;
      },
      /** all selected rows as high as the highest of them (measured) */
      distributeRows: (heightTw) => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        let pos = t.pos + 1;
        t.node.forEach((row) => { tr.setNodeMarkup(pos, undefined, { ...row.attrs, height: heightTw, heightRule: null }); pos += row.nodeSize; });
        if (dispatch) dispatch(tr);
        return true;
      },
      /** AutoFit: contents (no widths – the browser fits the text), window (100 % of the text width), fixed (current widths) */
      autoFit: (mode, widths) => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        const a = t.node.attrs;
        if (mode === 'window') tr.setNodeMarkup(t.pos, undefined, { ...a, layout: null, width: { type: 'pct', w: 5000 } });
        else if (mode === 'fixed') {
          tr.setNodeMarkup(t.pos, undefined, { ...a, layout: 'fixed', width: null });
          if (widths) t.node.descendants((n, p) => { if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') { const k = widths.cellCol.get(p); if (k != null) tr.setNodeMarkup(t.pos + 1 + p, undefined, { ...n.attrs, colwidth: widths.cols.slice(k, k + (n.attrs.colspan || 1)) }); } return n.type.name === 'table' || n.type.name === 'tableRow'; });
        } else {
          tr.setNodeMarkup(t.pos, undefined, { ...a, layout: null, width: null });
          if (widths) t.node.descendants((n, p) => { if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') { const k = widths.cellCol.get(p); if (k != null) tr.setNodeMarkup(t.pos + 1 + p, undefined, { ...n.attrs, colwidth: widths.cols.slice(k, k + (n.attrs.colspan || 1)) }); } return n.type.name === 'table' || n.type.name === 'tableRow'; });
        }
        if (dispatch) dispatch(tr);
        return true;
      },
      /** split the table above the current row (Word: Layout → Split table) */
      splitTable: () => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        const $p = state.selection.$from;
        const rowIndex = $p.index(t.depth);
        if (rowIndex < 1) return false;
        const rows = []; t.node.forEach((r) => rows.push(r));
        const a = t.node.type.create(t.node.attrs, rows.slice(0, rowIndex));
        const b = t.node.type.create(t.node.attrs, rows.slice(rowIndex));
        const p = state.schema.nodes.paragraph.create();
        tr.replaceWith(t.pos, t.pos + t.node.nodeSize, [a, p, b]);
        tr.setSelection(TextSelection.near(tr.doc.resolve(t.pos + a.nodeSize + 1)));
        if (dispatch) dispatch(tr.scrollIntoView());
        return true;
      },
      /** sort the body rows by one column (text or numbers, de/en aware), header rows stay on top */
      sortTable: ({ col = 0, desc = false, numeric = null } = {}) => ({ state, tr, dispatch }) => {
        const t = findTable(state); if (!t) return false;
        const rows = []; t.node.forEach((r) => rows.push(r));
        const head = []; while (rows.length && rows[0].childCount && Array.from({ length: rows[0].childCount }, (_, i) => rows[0].child(i)).every((c) => c.type.name === 'tableHeader')) head.push(rows.shift());
        if (rows.some((r) => { let span = false; r.forEach((c) => { if ((c.attrs.rowspan || 1) > 1) span = true; }); return span; })) return false;
        const key = (r) => (col < r.childCount ? r.child(col).textContent.trim() : '');
        const num = (s) => { const v = parseFloat(String(s).replace(/\s/g, '').replace(/\.(?=\d{3}\b)/g, '').replace(',', '.')); return Number.isFinite(v) ? v : null; };
        const isNum = numeric != null ? numeric : rows.every((r) => !key(r) || num(key(r)) != null);
        const coll = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
        rows.sort((x, y) => { const a = key(x); const b = key(y); const r = isNum ? (num(a) ?? -Infinity) - (num(b) ?? -Infinity) : coll.compare(a, b); return desc ? -r : r; });
        const nt = t.node.type.create(t.node.attrs, [...head, ...rows]);
        tr.replaceWith(t.pos, t.pos + t.node.nodeSize, nt);
        if (dispatch) dispatch(tr);
        return true;
      },
    };
  },
});

/**
 * Word's table formula fields: =SUM(ABOVE|LEFT|BELOW|RIGHT), AVERAGE, MIN, MAX, COUNT, PRODUCT – the value for the field
 * at `pos` (numbers read de/en style from the cells of that direction); null when the field is not in a table
 */
export function formulaValue(doc, pos, instr) {
  const m = /^=\s*(SUM|AVERAGE|MIN|MAX|COUNT|PRODUCT)\s*\(\s*(ABOVE|LEFT|BELOW|RIGHT)\s*\)/i.exec(String(instr || '').trim());
  if (!m) return null;
  const $p = doc.resolve(pos);
  let td = -1;
  for (let d = $p.depth; d > 0; d--) if ($p.node(d).type.name === 'table') { td = d; break; }
  if (td < 0) return null;
  const table = $p.node(td); const tstart = $p.start(td);
  const map = TableMap.get(table);
  const cellPos = $p.before(td + 2) - tstart;
  const rect = map.findCell(cellPos);
  const dir = m[2].toUpperCase();
  const vals = [];
  const read = (r, c) => { const p = map.map[r * map.width + c]; const n = table.nodeAt(p); if (!n) return; const v = parseFloat(n.textContent.replace(/\s/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.').replace(/[^\d.\-]/g, '')); if (Number.isFinite(v)) vals.push(v); };
  if (dir === 'ABOVE') for (let r = rect.top - 1; r >= 0; r--) read(r, rect.left);
  if (dir === 'BELOW') for (let r = rect.bottom; r < map.height; r++) read(r, rect.left);
  if (dir === 'LEFT') for (let c = rect.left - 1; c >= 0; c--) read(rect.top, c);
  if (dir === 'RIGHT') for (let c = rect.right; c < map.width; c++) read(rect.top, c);
  const f = m[1].toUpperCase();
  const v = f === 'SUM' ? vals.reduce((a, b) => a + b, 0) : f === 'AVERAGE' ? (vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0) : f === 'MIN' ? Math.min(...vals) : f === 'MAX' ? Math.max(...vals) : f === 'COUNT' ? vals.length : vals.reduce((a, b) => a * b, 1);
  return Number.isFinite(v) ? String(Math.round(v * 100) / 100).replace('.', (typeof document !== 'undefined' && document.documentElement.lang === 'de') ? ',' : '.') : '0';
}
