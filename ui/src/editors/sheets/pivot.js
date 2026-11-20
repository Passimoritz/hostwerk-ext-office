/*
 * Pivot tables: model, calculation (tabular layout like Excel's "Show in tabular form", no subtotals), writing the
 * result into the sheet, and the OOXML parts (pivotCacheDefinition + pivotCacheRecords + pivotTableDefinition, written
 * with refreshOnLoad so Excel rebuilds its own layout – and can refresh it later from the same source range).
 * Pure – runs in Node tests.
 *
 * pv = { id, name, source: "Data!$A$1:$D$40", at: { r, c }, rows: [field], cols: [field], values: [{ field, fn, name? }],
 *        filters: [{ field, exclude: [items] }], style, grandRows, grandCols, range (output area), path?, cachePath?, cacheId? }
 */
import { MAXC, parseRange, rangeA1, colName } from './formula/refs.js';
import { isErr } from './formula/errors.js';
import { esc, escText } from './io/xml.js';

const K = (r, c) => r * MAXC + c;
export const PIVOT_FNS = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'];
const CAPTION = { sum: 'Sum of {f}', count: 'Count of {f}', average: 'Average of {f}', max: 'Max of {f}', min: 'Min of {f}', product: 'Product of {f}', countNums: 'Count of {f}', stdDev: 'StdDev of {f}', stdDevp: 'StdDevp of {f}', var: 'Var of {f}', varp: 'Varp of {f}' };
export const DEFAULT_LABELS = { grand: 'Grand Total', values: 'Values', blank: '(blank)', caption: (fn, f) => CAPTION[fn].replace('{f}', f) };

/** source data → { fields: [names], records: [[v…]], sheet, g } or { error } */
export function pivotSource(book, pv) {
  const g = parseRange(pv.source || '');
  if (!g) return { error: 'The source range is not valid.' };
  const sheet = g.sheet ? book.sheetByName(g.sheet) : book.sheets[book.active];
  if (!sheet) return { error: 'The source sheet does not exist.' };
  const b = book.bounds(sheet);
  const r2 = Math.min(g.r2, b.rows - 1); const c2 = Math.min(g.c2, b.cols - 1);
  if (r2 < g.r1 || c2 < g.c1) return { error: 'The source range is empty.' };
  const fields = []; const seen = new Set();
  for (let c = g.c1; c <= c2; c++) {
    let n = book.text(sheet, g.r1, c).trim();
    if (!n) return { error: 'Every column of the source needs a heading.' };
    const base = n; let k = 2; while (seen.has(n.toLowerCase())) n = `${base}${k++}`;
    seen.add(n.toLowerCase()); fields.push(n);
  }
  const records = [];
  for (let r = g.r1 + 1; r <= r2; r++) {
    const row = []; let any = false;
    for (let c = g.c1; c <= c2; c++) { let v = book.value(sheet, r, c); if (v === '') v = null; if (v != null) any = true; row.push(v); }
    if (any) records.push(row);
  }
  return { fields, records, sheet, g: { r1: g.r1, c1: g.c1, r2, c2 } };
}

const itemKey = (v) => (v == null ? '\u0000' : isErr(v) ? `\u0001${v.code}` : typeof v === 'number' ? `n${v}` : typeof v === 'boolean' ? `b${v}` : `s${String(v).toLowerCase()}`);
const cmpItem = (a, b) => {
  const t = (v) => (v == null ? 3 : typeof v === 'number' ? 0 : typeof v === 'boolean' ? 2 : isErr(v) ? 4 : 1);
  const ta = t(a); const tb = t(b);
  if (ta !== tb) return ta - tb;
  if (ta === 0) return a - b;
  if (ta === 1) return String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
  return 0;
};
/** unique items of a field (sorted like Excel) */
export function fieldItems(src, fi) {
  const m = new Map();
  for (const rec of src.records) { const v = rec[fi]; const k = itemKey(v); if (!m.has(k)) m.set(k, v); }
  return [...m.values()].sort(cmpItem);
}
function agg(fn, list) {
  const nums = list.filter((v) => typeof v === 'number');
  switch (fn) {
    case 'count': return list.filter((v) => v != null && v !== '').length;
    case 'countNums': return nums.length;
    case 'average': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
    case 'max': return nums.length ? Math.max(...nums) : null;
    case 'min': return nums.length ? Math.min(...nums) : null;
    case 'product': return nums.length ? nums.reduce((a, b) => a * b, 1) : null;
    case 'stdDev': case 'var': case 'stdDevp': case 'varp': {
      const p = fn.endsWith('p'); const n = nums.length; if (n < (p ? 1 : 2)) return null;
      const mean = nums.reduce((a, b) => a + b, 0) / n; const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (p ? n : n - 1);
      return fn.startsWith('stdDev') ? Math.sqrt(ss) : ss;
    }
    default: return nums.reduce((a, b) => a + b, 0);
  }
}

/**
 * Calculate: → { rows: [[{ v, k }]] (k: 'cap' | 'fld' | 'item' | 'val' | 'grand' | 'gval' | 'page' | null), loc: {
 * rowOff (page area height), firstHeaderRow, firstDataRow, firstDataCol, w, h }, src, fi } or { error }
 */
export function computePivot(book, pv, labels = DEFAULT_LABELS) {
  const src = pivotSource(book, pv);
  if (src.error) return src;
  const fi = (n) => src.fields.findIndex((f) => f.toLowerCase() === String(n).toLowerCase());
  const rows = (pv.rows || []).map(fi).filter((i) => i >= 0);
  const cols = (pv.cols || []).map(fi).filter((i) => i >= 0);
  const vals = (pv.values || []).map((v) => ({ ...v, i: fi(v.field) })).filter((v) => v.i >= 0);
  const filters = (pv.filters || []).map((f) => ({ ...f, i: fi(f.field) })).filter((f) => f.i >= 0);
  // filtered records
  const recs = src.records.filter((rec) => filters.every((f) => !f.exclude || !f.exclude.length || !f.exclude.some((x) => itemKey(x) === itemKey(rec[f.i]))) && (pv.itemFilters ? Object.entries(pv.itemFilters).every(([field, ex]) => { const i = fi(field); return i < 0 || !ex.some((x) => itemKey(x) === itemKey(rec[i])); }) : true));
  const keyOf = (rec, idx) => idx.map((i) => itemKey(rec[i])).join('\u0002');
  const combos = (idx) => {
    const m = new Map();
    for (const rec of recs) { const k = keyOf(rec, idx); if (!m.has(k)) m.set(k, idx.map((i) => rec[i])); }
    return [...m.entries()].sort((a, b) => { for (let j = 0; j < idx.length; j++) { const d = cmpItem(a[1][j], b[1][j]); if (d) return d; } return 0; });
  };
  const rowKeys = rows.length ? combos(rows) : [['', []]];
  const colKeys = cols.length ? combos(cols) : [['', []]];
  // buckets
  const cell = new Map();
  const push = (k, rec) => { let a = cell.get(k); if (!a) cell.set(k, (a = [])); a.push(rec); };
  for (const rec of recs) {
    const rk = rows.length ? keyOf(rec, rows) : ''; const ck = cols.length ? keyOf(rec, cols) : '';
    push(`${rk}\u0003${ck}`, rec); push(`${rk}\u0003*`, rec); push(`*\u0003${ck}`, rec); push('*\u0003*', rec);
  }
  const value = (rk, ck, v) => { const list = cell.get(`${rk}\u0003${ck}`); return list ? agg(v.fn || 'sum', list.map((rec) => rec[v.i])) : null; };
  const nR = rows.length; const nC = cols.length; const V = vals.length;
  const multi = V > 1;
  const dataColList = []; // { ck, v }
  for (const [ck] of colKeys) for (const v of (V ? vals : [null])) dataColList.push({ ck, v });
  const grandCols = nC > 0 && pv.grandCols !== false && V > 0;
  if (grandCols) for (const v of vals) dataColList.push({ ck: '*', v, grand: true });
  const firstDataCol = Math.max(1, nR);
  const lbl = (x) => (x == null ? labels.blank : x);
  const capOf = (v) => v.name || labels.caption(v.fn || 'sum', src.fields[v.i]);
  const out = [];
  // like Excel: several value fields form an extra column level "Values"
  const levels = nC + (multi ? 1 : 0);
  const H = levels ? 1 + levels : 1;
  const W = firstDataCol + (V || nC ? dataColList.length : 0);
  for (let i = 0; i < H; i++) out.push(new Array(W).fill(null));
  if (levels) {
    out[0][0] = { v: V === 1 ? capOf(vals[0]) : null, k: 'cap' };
    cols.forEach((ci, j) => { out[0][firstDataCol + j] = { v: src.fields[ci], k: 'fld' }; });
    if (multi) out[0][firstDataCol + nC] = { v: labels.values, k: 'fld' };
    dataColList.forEach((d, j) => {
      const x = firstDataCol + j;
      if (d.grand) { out[1][x] = { v: multi ? `${labels.grand}: ${capOf(d.v)}` : labels.grand, k: 'grand' }; return; }
      const key = colKeys.find((q) => q[0] === d.ck)[1];
      key.forEach((item, lv) => { const prev = dataColList[j - 1]; const same = prev && !prev.grand && prev.ck.split('\u0002').slice(0, lv + 1).join('\u0002') === d.ck.split('\u0002').slice(0, lv + 1).join('\u0002'); if (!same) out[1 + lv][x] = { v: lbl(item), k: 'item' }; });
      if (multi) out[1 + nC][x] = { v: capOf(d.v), k: 'item' };
    });
  } else if (V) dataColList.forEach((d, j) => { out[0][firstDataCol + j] = { v: capOf(d.v), k: 'cap' }; });
  rows.forEach((ri, j) => { out[H - 1][j] = { v: src.fields[ri], k: 'fld' }; });
  if (!nR && !nC && V === 1) out[0][0] = { v: null, k: null };
  // data rows
  let prev = null;
  for (const [rk, key] of rowKeys) {
    if (!nR && !V) break;
    const line = new Array(W).fill(null);
    key.forEach((item, lv) => { if (!prev || prev.slice(0, lv + 1).some((x, z) => itemKey(x) !== itemKey(key[z]))) line[lv] = { v: lbl(item), k: 'item' }; });
    prev = key;
    if (!nR) line[0] = { v: null };
    dataColList.forEach((d, j) => { if (d.v) line[firstDataCol + j] = { v: value(nR ? rk : '*', d.grand ? '*' : nC ? d.ck : '*', d.v), k: d.grand ? 'gval' : 'val' }; });
    out.push(line);
  }
  if (nR && V && pv.grandRows !== false) {
    const line = new Array(W).fill(null);
    line[0] = { v: labels.grand, k: 'grand' };
    dataColList.forEach((d, j) => { line[firstDataCol + j] = { v: value('*', d.grand ? '*' : nC ? d.ck : '*', d.v), k: 'gval' }; });
    out.push(line);
  }
  // page (filter) fields above, one empty row between
  const page = filters.map((f) => [{ v: src.fields[f.i], k: 'fld' }, { v: f.exclude && f.exclude.length ? labels.multiple || '(Multiple Items)' : labels.all || '(All)', k: 'page' }]);
  const rowOff = page.length ? page.length + 1 : 0;
  return { rows: out, page, loc: { rowOff, firstHeaderRow: 1, firstDataRow: H, firstDataCol, w: W, h: out.length }, src, fi: { rows, cols, vals, filters } };
}

/** the whole area a pivot occupies (page fields included) */
export const pivotArea = (pv, res) => ({ r1: pv.at.r, c1: pv.at.c, r2: pv.at.r + res.loc.rowOff + res.loc.h - 1, c2: pv.at.c + Math.max(res.loc.w, res.page.length ? 2 : 0) - 1 });

/** write the result into the sheet (inside a transaction). Old output is cleared first. Returns the new area. */
export function renderPivot(book, sheet, pv, res, styles) {
  const S = styles || pivotStyles(book);
  if (pv.range) for (let r = pv.range.r1; r <= pv.range.r2; r++) for (let c = pv.range.c1; c <= pv.range.c2; c++) if (sheet.cells.has(K(r, c))) book.setCell(sheet, r, c, null);
  const r0 = pv.at.r; const c0 = pv.at.c;
  res.page.forEach((line, i) => line.forEach((x, j) => book.setCell(sheet, r0 + i, c0 + j, { v: x.v, s: j ? S.page : S.fld })));
  const top = r0 + res.loc.rowOff;
  res.rows.forEach((line, i) => line.forEach((x, j) => {
    const head = i < res.loc.firstDataRow;
    const grandRow = i === res.rows.length - 1 && line[0] && line[0].k === 'grand';
    const s = head ? S.head : grandRow ? S.grand : x && x.k === 'gval' ? S.gcol : j < res.loc.firstDataCol ? S.label : S.val;
    if ((x && x.v != null) || s) book.setCell(sheet, top + i, c0 + j, { v: x ? x.v : null, s });
  }));
  return pivotArea(pv, res);
}
export function pivotStyles(book) {
  const I = (o) => book.styles.intern(o);
  return {
    head: I({ b: true, fill: '#DDEBF7', bd: { b: { s: 'thin', c: '#9BC2E6' } } }), fld: I({ b: true }), page: I({ bd: { b: { s: 'thin', c: '#9BC2E6' } } }),
    label: null, val: null, grand: I({ b: true, fill: '#DDEBF7', bd: { t: { s: 'thin', c: '#9BC2E6' } } }), gcol: I({ b: true }),
  };
}

/* ---------------- OOXML ---------------- */
const SUBTOTAL_ATTR = { sum: null, count: 'count', average: 'average', max: 'max', min: 'min', product: 'product', countNums: 'countNums', stdDev: 'stdDev', stdDevp: 'stdDevp', var: 'var', varp: 'varp' };
const shared = (items, axis) => {
  const nums = items.filter((v) => typeof v === 'number'); const strs = items.filter((v) => typeof v === 'string' || typeof v === 'boolean' || isErr(v));
  const blank = items.some((v) => v == null);
  const a = [];
  if (!strs.length && nums.length) a.push('containsSemiMixedTypes="0"', 'containsString="0"');
  if (nums.length) { a.push('containsNumber="1"'); if (nums.every(Number.isInteger)) a.push('containsInteger="1"'); a.push(`minValue="${Math.min(...nums)}"`, `maxValue="${Math.max(...nums)}"`); }
  if (blank) a.push('containsBlank="1"');
  if (strs.length && nums.length) a.push('containsMixedTypes="1"');
  if (!axis) return `<sharedItems${a.length ? ` ${a.join(' ')}` : ''}/>`;
  const x = items.map((v) => (v == null ? '<m/>' : typeof v === 'number' ? `<n v="${v}"/>` : typeof v === 'boolean' ? `<b v="${v ? 1 : 0}"/>` : isErr(v) ? `<e v="${esc(v.code)}"/>` : `<s v="${esc(v)}"/>`)).join('');
  return `<sharedItems${a.length ? ` ${a.join(' ')}` : ''} count="${items.length}">${x}</sharedItems>`;
};
/**
 * The three parts of one pivot. srcSheetName/srcRef = source (A1 without $), loc = output area of the table body
 * (page fields excluded). keepExt = extLst of an original cache definition (slicers refer to it).
 */
export function pivotParts(pv, res, { cacheId, srcSheetName, srcRef, location, keepExt = '', recordsRid = 'rId1' }) {
  const { src, fi } = res;
  const axisSet = new Set([...fi.rows, ...fi.cols, ...fi.filters.map((f) => f.i)]);
  const itemsOf = src.fields.map((_, i) => (axisSet.has(i) ? fieldItems(src, i) : null));
  const cacheFields = src.fields.map((n, i) => `<cacheField name="${esc(n)}" numFmtId="0">${shared(itemsOf[i] || src.records.map((r) => r[i]), !!itemsOf[i])}</cacheField>`).join('');
  const cacheDef = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${recordsRid}" refreshOnLoad="1" refreshedBy="Hostwerk Office" refreshedDate="${(Date.now() / 86400000 + 25569).toFixed(6)}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${src.records.length}"><cacheSource type="worksheet"><worksheetSource ref="${srcRef}" sheet="${esc(srcSheetName)}"/></cacheSource><cacheFields count="${src.fields.length}">${cacheFields}</cacheFields>${keepExt}</pivotCacheDefinition>`;
  const idx = itemsOf.map((items) => (items ? new Map(items.map((v, j) => [itemKey(v), j])) : null));
  const recXml = src.records.map((rec) => `<r>${rec.map((v, i) => (idx[i] ? `<x v="${idx[i].get(itemKey(v))}"/>` : v == null ? '<m/>' : typeof v === 'number' ? `<n v="${v}"/>` : typeof v === 'boolean' ? `<b v="${v ? 1 : 0}"/>` : isErr(v) ? `<e v="${esc(v.code)}"/>` : `<s v="${esc(v)}"/>`)).join('')}</r>`).join('');
  const records = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="${src.records.length}">${recXml}</pivotCacheRecords>`;
  const dataIdx = new Set(fi.vals.map((v) => v.i));
  const excluded = (i) => { const f = fi.filters.find((x) => x.i === i); const ex = (f && f.exclude) || (pv.itemFilters && pv.itemFilters[src.fields[i]]) || []; return new Set(ex.map(itemKey)); };
  const pf = src.fields.map((_, i) => {
    const axis = fi.rows.includes(i) ? 'axisRow' : fi.cols.includes(i) ? 'axisCol' : fi.filters.some((f) => f.i === i) ? 'axisPage' : null;
    const a = `${axis ? ` axis="${axis}"` : ''}${dataIdx.has(i) ? ' dataField="1"' : ''} compact="0" outline="0" showAll="0"${axis && axis !== 'axisPage' ? ' defaultSubtotal="0"' : ''}${axis === 'axisPage' && excluded(i).size ? ' multipleItemSelectionAllowed="1"' : ''}`;
    if (!axis) return `<pivotField${a}/>`;
    const ex = excluded(i);
    const items = itemsOf[i].map((v, j) => `<item${ex.has(itemKey(v)) ? ' h="1"' : ''} x="${j}"/>`).join('') + (axis === 'axisPage' ? '<item t="default"/>' : '');
    return `<pivotField${a}><items count="${itemsOf[i].length + (axis === 'axisPage' ? 1 : 0)}">${items}</items></pivotField>`;
  }).join('');
  const multi = fi.vals.length > 1;
  const rowFields = fi.rows.length ? `<rowFields count="${fi.rows.length}">${fi.rows.map((i) => `<field x="${i}"/>`).join('')}</rowFields>` : '';
  const colList = [...fi.cols, ...(multi ? [-2] : [])];
  const colFields = colList.length ? `<colFields count="${colList.length}">${colList.map((i) => `<field x="${i}"/>`).join('')}</colFields>` : '';
  const pageFields = fi.filters.length ? `<pageFields count="${fi.filters.length}">${fi.filters.map((f) => `<pageField fld="${f.i}" hier="-1"/>`).join('')}</pageFields>` : '';
  const dataFields = fi.vals.length ? `<dataFields count="${fi.vals.length}">${fi.vals.map((v) => `<dataField name="${esc(v.name || DEFAULT_LABELS.caption(v.fn || 'sum', src.fields[v.i]))}" fld="${v.i}"${SUBTOTAL_ATTR[v.fn || 'sum'] ? ` subtotal="${SUBTOTAL_ATTR[v.fn]}"` : ''} baseField="0" baseItem="0"/>`).join('')}</dataFields>` : '';
  const L = res.loc;
  const table = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${esc(pv.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" rowGrandTotals="${pv.grandRows === false ? 0 : 1}" colGrandTotals="${pv.grandCols === false ? 0 : 1}" itemPrintTitles="1" createdVersion="6" indent="0" compact="0" compactData="0" outline="1" outlineData="1" gridDropZones="1" multipleFieldFilters="0"><location ref="${location}" firstHeaderRow="${L.firstHeaderRow}" firstDataRow="${L.firstDataRow}" firstDataCol="${L.firstDataCol}"${fi.filters.length ? ` rowPageCount="${fi.filters.length}" colPageCount="1"` : ''}/><pivotFields count="${src.fields.length}">${pf}</pivotFields>${rowFields}${colFields}${pageFields}${dataFields}<pivotTableStyleInfo name="${esc(pv.style || 'PivotStyleLight16')}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
  return { cacheDef, records, table };
}

/** reader: pivotTableDefinition + its cache definition → model (or a read-only stub for features we can't rebuild) */
export function readPivot(tableXml, cacheXml, { path, cachePath, sheetName }) {
  const A = (tag, xml) => { const m = new RegExp(`<${tag}\\b([^>]*?)/?>`).exec(xml); const o = {}; if (m) for (const q of m[1].matchAll(/([\w:]+)="([^"]*)"/g)) o[q[1]] = q[2]; return o; };
  const un = (s) => String(s || '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const head = A('pivotTableDefinition', tableXml);
  const loc = A('location', tableXml);
  const ws = A('worksheetSource', cacheXml || '');
  const names = [...(cacheXml || '').matchAll(/<cacheField\b[^>]*?\sname="([^"]*)"/g)].map((m) => un(m[1]));
  const sharedLists = [...(cacheXml || '').matchAll(/<cacheField\b[\s\S]*?(?:<\/cacheField>|\/>)/g)].map((m) => { const si = /<sharedItems\b[^>]*>([\s\S]*?)<\/sharedItems>/.exec(m[0]); if (!si) return []; return [...si[1].matchAll(/<(s|n|b|m|e|d)\b(?:[^>]*?\sv="([^"]*)")?[^>]*\/?>/g)].map((q) => (q[1] === 'm' ? null : q[1] === 'n' ? +q[2] : q[1] === 'b' ? q[2] === '1' : un(q[2]))); });
  const xs = (tag) => { const m = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(tableXml); return m ? [...m[1].matchAll(/<field\b[^>]*\sx="(-?\d+)"/g)].map((q) => +q[1]) : []; };
  const pfs = [...tableXml.matchAll(/<pivotField\b[\s\S]*?(?:<\/pivotField>|\/>)/g)].map((m) => m[0]);
  const hidden = {};
  pfs.forEach((x, i) => { const h = [...x.matchAll(/<item\b[^>]*\sh="1"[^>]*\sx="(\d+)"|<item\b[^>]*\sx="(\d+)"[^>]*\sh="1"/g)].map((q) => +(q[1] || q[2])); if (h.length && names[i]) hidden[names[i]] = h.map((j) => (sharedLists[i] || [])[j]); });
  const pages = [...tableXml.matchAll(/<pageField\b([^>]*)\/>/g)].map((m) => { const o = {}; for (const q of m[1].matchAll(/([\w:]+)="([^"]*)"/g)) o[q[1]] = q[2]; return o; });
  const dfs = [...tableXml.matchAll(/<dataField\b([^>]*)\/?>/g)].map((m) => { const o = {}; for (const q of m[1].matchAll(/([\w:]+)="([^"]*)"/g)) o[q[1]] = q[2]; return o; });
  const g = parseRange(loc.ref || 'A1');
  const complex = /<calculatedItems|<fieldGroup|<cacheHierarchies|<calculatedItem\b|<formula\b|<pivotHierarchies/.test(tableXml + (cacheXml || '')) || !ws.ref || dfs.some((d) => d.showDataAs);
  const filters = pages.map((p) => { const n = names[+p.fld]; const item = p.item != null ? (sharedLists[+p.fld] || [])[+p.item] : undefined; const all = sharedLists[+p.fld] || []; return { field: n, exclude: item !== undefined ? all.filter((v) => v !== item) : hidden[n] || [] }; });
  const itemFilters = {}; for (const [n, ex] of Object.entries(hidden)) if (!filters.some((f) => f.field === n)) itemFilters[n] = ex;
  const fn = (d) => d.subtotal || 'sum';
  const rowOff = pages.length ? pages.length + 1 : 0;
  return {
    id: `pv${Math.random().toString(36).slice(2, 8)}`, name: un(head.name || 'PivotTable1'), cacheId: +head.cacheId, path, cachePath,
    source: ws.ref ? `${quoteName(un(ws.sheet || sheetName))}!${absRef(ws.ref)}` : null, sourceName: ws.name ? un(ws.name) : null,
    rows: xs('rowFields').filter((x) => x >= 0).map((x) => names[x]), cols: xs('colFields').filter((x) => x >= 0).map((x) => names[x]),
    values: dfs.map((d) => ({ field: names[+d.fld], fn: fn(d), name: un(d.name) })), filters, itemFilters,
    style: (A('pivotTableStyleInfo', tableXml).name) || 'PivotStyleLight16', grandRows: head.rowGrandTotals !== '0', grandCols: head.colGrandTotals !== '0',
    at: g ? { r: Math.max(0, g.r1 - rowOff), c: g.c1 } : { r: 0, c: 0 }, range: g ? { r1: Math.max(0, g.r1 - rowOff), c1: g.c1, r2: g.r2, c2: g.c2 } : null,
    readOnly: complex, _orig: true,
  };
}
const quoteName = (n) => (/^[A-Za-z_][\w.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${String(n).replace(/'/g, "''")}'`);
const absRef = (ref) => String(ref).replace(/\$/g, '').replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
export const pivotSig = (pv) => { const { id, path, cachePath, cacheId, range, _orig, _sig, _dirty, _moved, readOnly, ...rest } = pv; return JSON.stringify(rest); };
export { rangeA1, colName, escText };
