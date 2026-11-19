/* Table operations on a table element (pure). Cells: { text, fill, span: { c, r }, merged }. */
import { cell, body, para } from './model.js';

const copyRow = (row) => ({ ...row, cells: row.cells.map((c) => ({ ...cell(''), fill: c.fill, text: { ...cell('').text, lvl: c.text && c.text.lvl } })) });
const resize = (el, rows, cols) => ({ ...el, rows, cols, w: cols.reduce((a, b) => a + b, 0) * (el.w / ((el.cols || []).reduce((a, b) => a + b, 0) || el.w)), h: rows.reduce((a, r) => a + r.h, 0) });

/** owner of a cell (top-left of its merge area) */
export function ownerOf(el, r, c) {
  for (let ri = r; ri >= 0; ri--) for (let ci = c; ci >= 0; ci--) {
    const x = el.rows[ri] && el.rows[ri].cells[ci];
    if (x && !x.merged) { const sp = x.span || {}; if (ri + (sp.r || 1) > r && ci + (sp.c || 1) > c) return { r: ri, c: ci }; }
  }
  return { r, c };
}
function unmergeAll(el, pred) {
  const rows = el.rows.map((row) => ({ ...row, cells: row.cells.map((x) => ({ ...x })) }));
  rows.forEach((row, ri) => row.cells.forEach((x, ci) => {
    if (x.merged || !x.span) return;
    if (pred(ri, ci, x.span)) { for (let a = 0; a < (x.span.r || 1); a++) for (let b = 0; b < (x.span.c || 1); b++) { const y = rows[ri + a] && rows[ri + a].cells[ci + b]; if (y) { y.merged = false; } } delete x.span; }
  }));
  return { ...el, rows };
}
export function tableOp(el, op, at = { r: 0, c: 0 }, value) {
  const { r, c } = at;
  if (op === 'rowAbove' || op === 'rowBelow') {
    const i = op === 'rowAbove' ? r : r + 1;
    const base = unmergeAll(el, (ri, ci, sp) => ri < i && ri + (sp.r || 1) > i);
    const rows = base.rows.slice(); rows.splice(i, 0, copyRow(base.rows[Math.min(r, base.rows.length - 1)]));
    return resize(base, rows, base.cols);
  }
  if (op === 'colLeft' || op === 'colRight') {
    const i = op === 'colLeft' ? c : c + 1;
    const base = unmergeAll(el, (ri, ci, sp) => ci < i && ci + (sp.c || 1) > i);
    const cw = base.cols[Math.min(c, base.cols.length - 1)];
    const cols = base.cols.slice(); cols.splice(i, 0, cw);
    const rows = base.rows.map((row) => { const cells = row.cells.slice(); cells.splice(i, 0, { ...cell(''), fill: (row.cells[c] || {}).fill }); return { ...row, cells }; });
    return resize(base, rows, cols);
  }
  if (op === 'delRow') {
    if (el.rows.length <= 1) return null;
    const base = unmergeAll(el, (ri, ci, sp) => ri <= r && ri + (sp.r || 1) > r);
    return resize(base, base.rows.filter((_, k) => k !== r), base.cols);
  }
  if (op === 'delCol') {
    if (el.cols.length <= 1) return null;
    const base = unmergeAll(el, (ri, ci, sp) => ci <= c && ci + (sp.c || 1) > c);
    return resize(base, base.rows.map((row) => ({ ...row, cells: row.cells.filter((_, k) => k !== c) })), base.cols.filter((_, k) => k !== c));
  }
  if (op === 'mergeRight' || op === 'mergeDown') {
    const o = ownerOf(el, r, c);
    const own = el.rows[o.r].cells[o.c];
    const sp = { c: (own.span && own.span.c) || 1, r: (own.span && own.span.r) || 1 };
    if (op === 'mergeRight') { if (o.c + sp.c >= el.cols.length) return el; sp.c += 1; } else { if (o.r + sp.r >= el.rows.length) return el; sp.r += 1; }
    const rows = el.rows.map((row) => ({ ...row, cells: row.cells.map((x) => ({ ...x })) }));
    const texts = [];
    for (let a = 0; a < sp.r; a++) for (let b = 0; b < sp.c; b++) {
      const y = rows[o.r + a].cells[o.c + b];
      if (a === 0 && b === 0) continue;
      if (!y.merged && y.text && y.text.paras.some((p) => p.runs.some((x) => x.text))) texts.push(...y.text.paras);
      y.merged = true; delete y.span;
    }
    const tgt = rows[o.r].cells[o.c];
    tgt.span = sp;
    if (texts.length) tgt.text = { ...tgt.text, paras: [...tgt.text.paras.filter((p) => p.runs.some((x) => x.text)), ...texts] };
    return { ...el, rows };
  }
  if (op === 'split') {
    const o = ownerOf(el, r, c);
    return unmergeAll(el, (ri, ci) => ri === o.r && ci === o.c);
  }
  if (op === 'fill') {
    const o = ownerOf(el, r, c);
    return { ...el, rows: el.rows.map((row, ri) => (ri !== o.r ? row : { ...row, cells: row.cells.map((x, ci) => (ci !== o.c ? x : { ...x, fill: value ? { type: 'solid', color: value } : null })) })) };
  }
  return el;
}
export const emptyCellBody = () => body([para('')]);
