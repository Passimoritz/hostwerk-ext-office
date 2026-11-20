/*
 * Excel tables (ListObjects): create / resize / totals row / style / convert to range, the look of the built-in table
 * styles (TableStyleLight1…21, Medium1…28, Dark1…11 – approximated from the theme accents) and the table part XML.
 * Pure – runs in Node tests. A table = { id, name, range, header, totals, style, stripes, colStripes, firstCol, lastCol,
 * columns: [names], totalsFn: { colIndex: 'sum'|'average'|'count'|'countNums'|'max'|'min'|'stdDev'|'var'|'none' }, path? }.
 */
import { MAXC, rangeA1, colName } from './formula/refs.js';
import { escCol } from './formula/parser.js';
import { esc, escText } from './io/xml.js';

// root element of the xlsx table part – XML, not an HTML table (Hostwerk's UI lint looks for HTML table tags)
const TABLE_TAG = 'table';
const K = (r, c) => r * MAXC + c;
export const TABLE_STYLES = [
  ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
  ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
  ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
];
export const SUBTOTAL_NO = { average: 101, count: 103, countNums: 102, max: 104, min: 105, stdDev: 107, sum: 109, var: 110 };
export const TOTAL_FNS = ['none', 'sum', 'average', 'count', 'countNums', 'max', 'min', 'stdDev', 'var'];
const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

const mix = (hex, w, base = 'FFFFFF') => { // w = share of hex (0..1)
  const a = parseInt(hex, 16); const b = parseInt(base, 16);
  const ch = (x, s) => (x >> s) & 255;
  const m = (s) => Math.round(ch(a, s) * w + ch(b, s) * (1 - w));
  return `#${((m(16) << 16) | (m(8) << 8) | m(0)).toString(16).padStart(6, '0')}`;
};

/** colours of a built-in table style: { head, headText, band, border, total, totalText, firstCol } */
export function tableLook(style, theme) {
  const th = theme && theme.length >= 10 ? theme : DEFAULT_THEME;
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(style || 'TableStyleMedium2');
  if (!m) return tableLook('TableStyleMedium2', theme);
  const n = +m[2]; const k = (n - 1) % 7; // 0 = black/grey, 1..6 = accent 1..6
  const acc = k === 0 ? '000000' : th[3 + k];
  const fam = m[1];
  if (fam === 'Light') {
    const grp = Math.floor((n - 1) / 7); // 0: lines, 1: header filled, 2: grid
    if (grp === 1) return { head: `#${acc}`, headText: '#FFFFFF', band: mix(acc, 0.2), border: `#${acc}`, total: null, totalText: '#000000', bold: true, line: `#${acc}` };
    return { head: null, headText: k ? `#${acc}` : '#000000', band: mix(acc, 0.2), border: `#${acc}`, total: null, totalText: '#000000', bold: true, line: `#${acc}`, grid: grp === 2 };
  }
  if (fam === 'Medium') {
    const grp = Math.floor((n - 1) / 7);
    if (grp === 0) return { head: `#${acc}`, headText: '#FFFFFF', band: mix(acc, 0.2), border: mix(acc, 0.6), total: null, totalText: '#000000', bold: true, line: `#${acc}` };
    if (grp === 1) return { head: `#${acc}`, headText: '#FFFFFF', band: mix(acc, 0.4), bandOdd: mix(acc, 0.2), border: '#FFFFFF', total: `#${acc}`, totalText: '#FFFFFF', bold: true };
    if (grp === 2) return { head: '#000000', headText: '#FFFFFF', band: mix(acc, 0.4), bandOdd: mix(acc, 0.2), border: '#FFFFFF', total: null, totalText: '#000000', bold: true };
    return { head: mix(acc, 0.4), headText: '#000000', band: mix(acc, 0.2), border: mix(acc, 0.6), total: null, totalText: '#000000', bold: true };
  }
  return { head: '#000000', headText: '#FFFFFF', band: mix(acc, 0.75, '000000'), bandOdd: mix(acc, 1, '000000'), text: '#FFFFFF', border: null, total: '#000000', totalText: '#FFFFFF', bold: true };
}

/** per-cell look of the tables on a sheet (no own fill/colour wins over it): fn(r, c) → { fill, color, b, bd } | null */
export function tableStyleAt(book, sheet) {
  const list = (sheet.tables || []).filter((t) => t.style !== 'none');
  if (!list.length) return null;
  const theme = book.pkg && book.pkg.theme;
  const looks = list.map((t) => ({ t, L: tableLook(t.style, theme) }));
  return (r, c) => {
    for (const { t, L } of looks) {
      const g = t.range;
      if (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2) continue;
      const hdr = t.header !== false; const tot = !!t.totals;
      if (hdr && r === g.r1) return { fill: L.head, color: L.headText, b: L.bold, bd: L.line && !L.head ? { b: { s: 'thin', c: L.line } } : null };
      if (tot && r === g.r2) return { fill: L.total, color: L.totalText, b: true, bd: { t: { s: 'double', c: L.line || L.head || '#000000' } } };
      const i = r - g.r1 - (hdr ? 1 : 0);
      const band = t.stripes !== false && i % 2 === 0 ? L.band : L.bandOdd || null;
      const colBand = t.colStripes && (c - g.c1) % 2 === 0 ? L.band : null;
      const first = t.firstCol && c === g.c1; const last = t.lastCol && c === g.c2;
      return { fill: colBand || band, color: L.text || null, b: first || last, bd: L.grid ? { b: { s: 'thin', c: L.border }, r: { s: 'thin', c: L.border } } : null };
    }
    return null;
  };
}

/** a free table name (Table1, Table2 … or Tabelle1 …) */
export function uniqueTableName(book, base = 'Table') {
  const used = new Set(book.allTables().map((x) => String(x.t.name).toLowerCase()));
  for (const n of book.names.values()) used.add(String(n.name).toLowerCase());
  let i = 1; while (used.has(`${base}${i}`.toLowerCase())) i++;
  return `${base}${i}`;
}
export const validTableName = (n) => /^[A-Za-z_\\À-￿][\w.À-￿]{0,254}$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^R\d*C\d*$/i.test(n);

/** header names must be unique, non-empty texts (Excel repairs the file otherwise) */
function headerNames(book, sheet, g) {
  const seen = new Set(); const out = [];
  for (let c = g.c1; c <= g.c2; c++) {
    let n = book.text(sheet, g.r1, c).trim() || `Column${c - g.c1 + 1}`;
    let k = 2; const base = n; while (seen.has(n.toLowerCase())) n = `${base}${k++}`;
    seen.add(n.toLowerCase()); out.push(n);
  }
  return out;
}

/** Insert → Table / Home → Format as table. header=false inserts a header row with Column1 … above the range. */
export function addTable(book, sheet, g0, { header = true, style = 'TableStyleMedium2', name, base = 'Table' } = {}) {
  let g = { ...g0 };
  const clash = (sheet.tables || []).some((t) => t.range.r1 <= g.r2 && g.r1 <= t.range.r2 && t.range.c1 <= g.c2 && g.c1 <= t.range.c2);
  if (clash) return null;
  const nm = name && validTableName(name) ? name : uniqueTableName(book, base);
  let tb = null;
  book.run('Format as table', () => {
    book.touch(sheet);
    if (!header) { book.shiftCells(sheet, { r1: g.r1, c1: g.c1, r2: g.r1, c2: g.c2 }, 'down', true); g = { ...g, r2: g.r2 + 1 }; for (let c = g.c1; c <= g.c2; c++) book.input(sheet, g.r1, c, `Column${c - g.c1 + 1}`); }
    const names = headerNames(book, sheet, g);
    names.forEach((n, j) => { const cur = book.value(sheet, g.r1, g.c1 + j); if (cur !== n) { const cell = sheet.cells.get(K(g.r1, g.c1 + j)) || {}; book.setCell(sheet, g.r1, g.c1 + j, { ...cell, v: n, f: undefined }); } });
    tb = { id: `t${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`, name: nm, range: g, header: true, totals: false, style, stripes: true, columns: names, totalsFn: {}, _new: true };
    book.touch(sheet);
    sheet.tables = [...(sheet.tables || []), tb];
    book.tx.structural = true;
  });
  return tb;
}
const findT = (sheet, id) => (sheet.tables || []).find((t) => t.id === id || t.name === id);
export function setTableProps(book, sheet, id, patch, label = 'Table') {
  book.run(label, () => {
    book.touch(sheet);
    sheet.tables = (sheet.tables || []).map((t) => (t.id === id || t.name === id ? { ...t, ...patch, _dirty: true } : t));
    book.tx.structural = true;
  });
}
export function resizeTable(book, sheet, id, g) { setTableProps(book, sheet, id, { range: g }, 'Resize table'); }
/** totals row on/off – writes SUBTOTAL formulas like Excel (sum under the last numeric column, a "Total" label first) */
export function toggleTotals(book, sheet, id, on) {
  const t = findT(sheet, id); if (!t || !!t.totals === !!on) return;
  book.run(on ? 'Total row' : 'Remove total row', () => {
    book.touch(sheet, true);
    const g = { ...t.range };
    const cols = book.tableColumns(sheet, t);
    if (on) {
      const r = g.r2 + 1;
      const below = [...sheet.cells.keys()].some((k) => Math.floor(k / MAXC) === r && k % MAXC >= g.c1 && k % MAXC <= g.c2);
      if (below) book.shiftCells(sheet, { r1: r, c1: g.c1, r2: r, c2: g.c2 }, 'down', true);
      const fns = { ...(t.totalsFn || {}) };
      let lastNum = -1;
      for (let c = g.c1; c <= g.c2; c++) { const v = book.value(sheet, g.r2, c); if (typeof v === 'number') lastNum = c; }
      if (!Object.keys(fns).length && lastNum >= 0) fns[lastNum - g.c1] = 'sum';
      book.input(sheet, r, g.c1, fns[0] ? '' : 'Total');
      for (const [j, fn] of Object.entries(fns)) if (fn && fn !== 'none') book.setCell(sheet, r, g.c1 + +j, { ...(sheet.cells.get(K(r, g.c1 + +j)) || {}), f: `SUBTOTAL(${SUBTOTAL_NO[fn]},${t.name}[${escCol(cols[+j])}])`, v: null });
      sheet.tables = sheet.tables.map((x) => (x === t ? { ...x, totals: true, totalsFn: fns, totalsLabel: fns[0] ? null : 'Total', range: { ...g, r2: r }, _dirty: true } : x));
    } else {
      for (let c = g.c1; c <= g.c2; c++) book.setCell(sheet, g.r2, c, null);
      sheet.tables = sheet.tables.map((x) => (x === t ? { ...x, totals: false, range: { ...g, r2: g.r2 - 1 }, _dirty: true } : x));
    }
    book.tx.structural = true;
  });
}
export function setTotalFn(book, sheet, id, j, fn) {
  const t = findT(sheet, id); if (!t || !t.totals) return;
  const cols = book.tableColumns(sheet, t);
  book.run('Total row', () => {
    book.touch(sheet);
    const r = t.range.r2; const c = t.range.c1 + j;
    const cell = sheet.cells.get(K(r, c)) || {};
    if (!fn || fn === 'none') book.setCell(sheet, r, c, { ...cell, f: undefined, v: null });
    else book.setCell(sheet, r, c, { ...cell, f: `SUBTOTAL(${SUBTOTAL_NO[fn]},${t.name}[${escCol(cols[j])}])`, v: null });
    sheet.tables = sheet.tables.map((x) => (x === t ? { ...x, totalsFn: { ...(x.totalsFn || {}), [j]: fn }, _dirty: true } : x));
  });
}
/** Convert to range: structured references in formulas become normal references */
export function convertToRange(book, sheet, id) {
  const t = findT(sheet, id); if (!t) return;
  book.run('Convert to range', () => {
    book.touch(sheet);
    sheet.tables = sheet.tables.filter((x) => x !== t);
    book.tx.structural = true;
  });
}
/** table that grows when typing right below / right of it (Excel's auto-expansion) */
export function autoExpand(book, sheet, r, c) {
  for (const t of sheet.tables || []) {
    const g = t.range;
    if (!t.totals && r === g.r2 + 1 && c >= g.c1 && c <= g.c2) { resizeTable(book, sheet, t.id, { ...g, r2: r }); return t; }
    if (c === g.c2 + 1 && r >= g.r1 && r <= g.r2 && r === g.r1 && t.header !== false) { resizeTable(book, sheet, t.id, { ...g, c2: c }); return t; }
  }
  return null;
}

/* ---------------- XML ---------------- */
const FN_XML = { sum: 'sum', average: 'average', count: 'count', countNums: 'countNums', max: 'max', min: 'min', stdDev: 'stdDev', var: 'var' };
/** a complete table part for a new table (id = workbook-unique number) */
export function tableXml(book, sheet, t, id) {
  const g = t.range;
  const cols = book.tableColumns(sheet, t);
  const af = t.header !== false ? `<autoFilter ref="${rangeA1({ ...g, r2: t.totals ? g.r2 - 1 : g.r2 })}"/>` : '';
  const tc = cols.map((n, j) => {
    const fn = t.totals && t.totalsFn && t.totalsFn[j];
    const label = t.totals && j === 0 && !fn ? ` totalsRowLabel="${esc(t.totalsLabel || 'Total')}"` : '';
    return `<tableColumn id="${j + 1}" name="${esc(n)}"${label}${fn && fn !== 'none' ? ` totalsRowFunction="${FN_XML[fn] || 'sum'}"` : ''}/>`;
  }).join('');
  const st = t.style === 'none' ? '' : `<tableStyleInfo name="${esc(t.style || 'TableStyleMedium2')}" showFirstColumn="${t.firstCol ? 1 : 0}" showLastColumn="${t.lastCol ? 1 : 0}" showRowStripes="${t.stripes === false ? 0 : 1}" showColumnStripes="${t.colStripes ? 1 : 0}"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${TABLE_TAG} xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${esc(t.name)}" displayName="${esc(t.name)}" ref="${rangeA1(g)}"${t.header === false ? ' headerRowCount="0"' : ''}${t.totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>${af}<tableColumns count="${cols.length}">${tc}</tableColumns>${st}</table>`;
}
/** patch an existing table part (keeps everything we don't model: calculated column formulas, dxfs, extLst …) */
export function patchTableXml(xml, book, sheet, t) {
  const g = t.range;
  let x = xml.replace(/(<table\b[^>]*?\sref=")[^"]*"/, `$1${rangeA1(g)}"`);
  x = x.replace(/(<table\b[^>]*?\sname=")[^"]*"/, `$1${esc(t.name)}"`).replace(/(<table\b[^>]*?\sdisplayName=")[^"]*"/, `$1${esc(t.name)}"`);
  x = x.replace(/\stotalsRowCount="\d+"/, '').replace(/\stotalsRowShown="\d+"/, '').replace(/<table\b/, `<table${t.totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}`);
  const afRef = rangeA1({ ...g, r2: t.totals ? g.r2 - 1 : g.r2 });
  if (/<autoFilter\b/.test(x)) x = x.replace(/(<autoFilter\b[^>]*?\sref=")[^"]*"/, `$1${afRef}"`); else if (t.header !== false) x = x.replace(/(<table\b[^>]*>)/, `$1<autoFilter ref="${afRef}"/>`);
  const cols = book.tableColumns(sheet, t);
  const old = [...x.matchAll(/<tableColumn\b[\s\S]*?(?:\/>|<\/tableColumn>)/g)].map((m) => m[0]);
  const items = cols.map((n, j) => {
    const fn = t.totals && t.totalsFn && t.totalsFn[j];
    let e = old[j] || `<tableColumn id="${j + 1}" name=""/>`;
    e = e.replace(/\sname="[^"]*"/, ` name="${esc(n)}"`).replace(/\stotalsRowFunction="[^"]*"/, '').replace(/\stotalsRowLabel="[^"]*"/, '');
    if (fn && fn !== 'none') e = e.replace(/<tableColumn\b/, `<tableColumn totalsRowFunction="${FN_XML[fn] || 'sum'}"`);
    else if (t.totals && j === 0) e = e.replace(/<tableColumn\b/, `<tableColumn totalsRowLabel="${esc(t.totalsLabel || 'Total')}"`);
    return e;
  });
  const ids = new Set(); let maxId = 0; items.forEach((e) => { const m = /\sid="(\d+)"/.exec(e); if (m) maxId = Math.max(maxId, +m[1]); });
  const fixed = items.map((e) => { let m = /\sid="(\d+)"/.exec(e); if (!m || ids.has(m[1])) { e = e.replace(/\sid="\d+"/, '').replace(/<tableColumn\b/, `<tableColumn id="${++maxId}"`); m = /\sid="(\d+)"/.exec(e); } ids.add(m[1]); return e; });
  x = x.replace(/<tableColumns\b[\s\S]*?<\/tableColumns>/, `<tableColumns count="${cols.length}">${fixed.join('')}</tableColumns>`);
  if (t.style && t.style !== 'none') {
    const st = `<tableStyleInfo name="${esc(t.style)}" showFirstColumn="${t.firstCol ? 1 : 0}" showLastColumn="${t.lastCol ? 1 : 0}" showRowStripes="${t.stripes === false ? 0 : 1}" showColumnStripes="${t.colStripes ? 1 : 0}"/>`;
    x = /<tableStyleInfo\b/.test(x) ? x.replace(/<tableStyleInfo\b[^>]*\/>/, st) : x.replace(/(<\/tableColumns>)/, `$1${st}`);
  }
  return x;
}
/** reader: table part → table model */
export function readTableXml(xml, path) {
  const head = /<table\b[^>]*>/.exec(xml); if (!head) return null;
  const a = {}; for (const m of head[0].matchAll(/([\w:]+)="([^"]*)"/g)) a[m[1]] = m[2];
  const cols = [...xml.matchAll(/<tableColumn\b([^>]*?)(?:\/>|>)/g)].map((m) => { const o = {}; for (const q of m[1].matchAll(/([\w:]+)="([^"]*)"/g)) o[q[1]] = q[2]; return o; });
  const st = /<tableStyleInfo\b([^>]*)\/>/.exec(xml); const sa = {}; if (st) for (const m of st[1].matchAll(/([\w:]+)="([^"]*)"/g)) sa[m[1]] = m[2];
  const un = (s) => String(s || '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&').replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
  const totalsFn = {}; cols.forEach((c, j) => { if (c.totalsRowFunction && c.totalsRowFunction !== 'none' && c.totalsRowFunction !== 'custom') totalsFn[j] = c.totalsRowFunction; });
  return {
    id: `t${a.id || Math.random().toString(36).slice(2, 6)}`, xmlId: +a.id || 0, path, name: un(a.displayName || a.name), refText: a.ref,
    header: a.headerRowCount !== '0', totals: +(a.totalsRowCount || 0) > 0, columns: cols.map((c) => un(c.name)), totalsFn,
    totalsLabel: cols[0] && cols[0].totalsRowLabel ? un(cols[0].totalsRowLabel) : null,
    style: st ? sa.name || 'none' : 'none', stripes: sa.showRowStripes !== '0', colStripes: sa.showColumnStripes === '1', firstCol: sa.showFirstColumn === '1', lastCol: sa.showLastColumn === '1',
  };
}
export { colName, escText };
