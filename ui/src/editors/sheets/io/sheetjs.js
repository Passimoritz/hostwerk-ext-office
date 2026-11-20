/*
 * Everything that isn't OOXML goes through SheetJS CE: xls/xlt/xlsb/ods/ots/fods/sxc/numbers/wk1/wks/qpw/dif/slk/dbf/prn
 * reading (values, formulas, number formats, merges, sizes, hidden rows/cols, notes, links, names) and xls/ods/fods/
 * dif/slk/dbf/xlsb writing (SheetJS CE writes no cell styles – formatting is kept only in xlsx, or via the office
 * engine when the shell has one: see formats.js).
 */
import * as XLSX from 'xlsx';
import { cptable } from './codepage.js';
import { Book, newSheet } from '../model.js';
import { MAXC } from '../formula/refs.js';
import { fromOoxml, toOoxml } from '../formula/parser.js';
import { errOf, isErr, ERROR_TYPE } from '../formula/errors.js';
import { formatValue, isDateFormat } from '../format/numfmt.js';
import { serialToDate } from '../format/dates.js';
import { readBiffStyles } from './biff-styles.js';
import { odfContentXml, odfRawFormulas, fromOpenFormula, PLACEHOLDER, patchOdfOutput } from './odf-formula.js';

const K = (r, c) => r * MAXC + c;
// old .xls / .dbf codepages (TextDecoder-backed table, see codepage.js)
if (XLSX.set_cptable) XLSX.set_cptable(cptable);
const ERRNUM = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA' };
const ERRCODE = Object.fromEntries(Object.entries(ERRNUM).map(([k, v]) => [v, +k]));

export function readSheetJS(bytes, { loc, settings = {}, ext } = {}) {
  const wb = XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true, cellDates: false, cellText: true, sheetStubs: false, dense: false, WTF: false });
  const book = new Book({ loc, settings });
  book.date1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
  // OpenDocument: SheetJS's own OpenFormula conversion loses references to other sheets – read the raw formulas
  const odfRaw = ['ods', 'ots', 'fods'].includes(ext) ? odfRawFormulas(odfContentXml(bytes)) : null;
  const styleCache = new Map();
  const nfStyle = (z) => { if (!z || z === 'General') return null; let s = styleCache.get(z); if (!s) { s = book.styles.intern({ nf: z }); styleCache.set(z, s); } return s; };
  wb.SheetNames.forEach((name, i) => {
    const ws = wb.Sheets[name];
    const sh = newSheet(name);
    const meta = wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[i];
    if (meta && meta.Hidden) sh.hidden = meta.Hidden === 2 ? 'veryHidden' : 'hidden';
    if (ws) {
      for (const a of Object.keys(ws)) {
        if (a[0] === '!') continue;
        const c = ws[a];
        const p = XLSX.utils.decode_cell(a);
        const cell = {};
        if (c.t === 'n') cell.v = c.v; else if (c.t === 's' || c.t === 'str') cell.v = c.v == null ? '' : String(c.v);
        else if (c.t === 'b') cell.v = !!c.v;
        else if (c.t === 'e') cell.v = errOf(c.w || ERRNUM[c.v] || '#N/A');
        else if (c.t === 'd') cell.v = c.v instanceof Date ? (c.v - Date.UTC(1899, 11, 30)) / 86400000 : null;
        else cell.v = null;
        if (c.f) {
          const raw = odfRaw && odfRaw[i] && odfRaw[i].get(`${p.r},${p.c}`);
          const of = raw != null ? fromOpenFormula(raw) : null;
          cell.f = fromOoxml(of != null ? of : String(c.f).replace(/^=/, '')); cell._cached = cell.v; cell._cachedF = cell.f;
        }
        if (c.F && c.F !== a && !c.f) { /* array formula member */ }
        let s = nfStyle(typeof c.z === 'number' ? XLSX.SSF._table[c.z] : c.z);
        // .xls fills SheetJS CE reads (cellStyles): keep them
        const fg = c.s && c.s.patternType === 'solid' && c.s.fgColor && c.s.fgColor.rgb;
        if (fg && /^[0-9A-F]{6}$/i.test(String(fg).slice(-6)) && String(fg).slice(-6).toUpperCase() !== 'FFFFFF') s = book.styles.intern({ ...(s || {}), fill: `#${String(fg).slice(-6)}` });
        if (s) cell.s = s;
        if (c.l && c.l.Target) cell.link = c.l.Target;
        if (c.c && c.c.length) { cell.note = c.c.map((x) => x.t).join('\n'); cell.noteAuthor = c.c[0].a || ''; }
        if (cell.v == null && !cell.f && !cell.s && !cell.note) continue;
        sh.cells.set(K(p.r, p.c), cell);
      }
      for (const m of ws['!merges'] || []) sh.merges.push({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c });
      (ws['!cols'] || []).forEach((col, c) => { if (!col) return; const px = col.wpx || (col.wch ? Math.round(col.wch * 7 + 5) : col.width ? Math.round(col.width * 7 + 5) : null); if (px) sh.colW.set(c, px); if (col.hidden) sh.hiddenCols.add(c); if (col.level) sh.outlineCols.set(c, col.level); });
      (ws['!rows'] || []).forEach((row, r) => { if (!row) return; const px = row.hpx || (row.hpt ? Math.round((row.hpt * 4) / 3) : null); if (px) sh.rowH.set(r, px); if (row.hidden) sh.hiddenRows.add(r); if (row.level) sh.outlineRows.set(r, row.level); });
      if (ws['!autofilter'] && ws['!autofilter'].ref) { const g = XLSX.utils.decode_range(ws['!autofilter'].ref); sh.filter = { range: { r1: g.s.r, c1: g.s.c, r2: g.e.r, c2: g.e.c }, cols: {} }; }
    }
    if (settings.gridlines === false) sh.gridlines = false;
    book.sheets.push(sh);
  });
  for (const n of (wb.Workbook && wb.Workbook.Names) || []) {
    if (!n.Name || n.Name.startsWith('_xlnm.') || !n.Ref) continue;
    const scope = n.Sheet != null ? book.sheets[n.Sheet] || null : null;
    book.names.set(`${n.Name.toUpperCase()}|${scope ? scope.id : ''}`, { name: n.Name, ref: fromOoxml(n.Ref.replace(/^=/, '')), scope, comment: n.Comment || '' });
  }
  if (!book.sheets.length) book.sheets.push(newSheet('Sheet1'));
  // .xls: fonts, fills, borders and alignment from the BIFF8 records (SheetJS CE gives number formats only)
  let biff = null;
  if (ext === 'xls' || ext === 'xlt') { try { biff = readBiffStyles(bytes); } catch { biff = null; } }
  if (biff) {
    book.props.defaultFont = biff.defaultFont;
    biff.sheets.forEach((cells, i) => {
      const sh = book.sheets[i]; if (!sh) return;
      for (const [rc, st] of cells) {
        const [r, c] = rc.split(',').map(Number);
        const k = K(r, c); const cell = sh.cells.get(k);
        const merged = book.styles.intern({ ...(cell && cell.s && cell.s.nf ? { nf: cell.s.nf } : {}), ...st });
        if (!merged) continue;
        sh.cells.set(k, cell ? { ...cell, s: merged } : { v: null, s: merged });
      }
    });
  }
  book.fidelity = [];
  if (['ods', 'ots', 'fods', 'sxc'].includes(ext)) book.fidelity.push('Cell formatting, charts and images of OpenDocument files are not read by the built-in reader (values, formulas, number formats, merges and sizes are)');
  else if (biff) book.fidelity.push('Charts, images and conditional formats of this older Excel format are not read by the built-in reader (values, formulas and cell formatting are)');
  else if (['xls', 'xlt', 'xlsb'].includes(ext)) book.fidelity.push('Cell formatting, charts and images of this older Excel format are not read by the built-in reader (values, formulas, number formats, merges and sizes are)');
  book.calc.rebuild();
  book.calc.recalc(null, { force: true });
  return book;
}

/** Book → SheetJS workbook object */
export function toSheetJS(book, { valuesOnly = false, odf = false, formulas = null } = {}) {
  const wb = XLSX.utils.book_new();
  wb.Workbook = { Sheets: [], Names: [], WBProps: { date1904: book.date1904 } };
  book.sheets.forEach((sh, i) => {
    const ws = {};
    let r2 = 0; let c2 = 0;
    const put = (k, cell, spilled) => {
      const r = Math.floor(k / MAXC); const c = k % MAXC;
      const v = cell.v;
      const o = {};
      if (typeof v === 'number') { o.t = 'n'; o.v = v; } else if (typeof v === 'boolean') { o.t = 'b'; o.v = v; } else if (isErr(v)) { o.t = 'e'; o.v = ERRCODE[v.code] ?? 0x0f; o.w = v.code === '#CIRC!' ? '#VALUE!' : v.code; } else if (v == null) { if (!cell.f) return; o.t = 'z'; } else { o.t = 's'; o.v = String(v); }
      if (cell.f && !valuesOnly && !spilled) {
        // OpenDocument: a placeholder SheetJS passes through untouched; patchOdfOutput() puts the real OpenFormula in
        if (formulas) { o.f = PLACEHOLDER(formulas.length); formulas.push(cell.f); } else o.f = toOoxml(cell.f).replace(/_xlfn\.(_xlws\.)?/g, '');
      }
      const nf = cell.s && cell.s.nf;
      if (nf) o.z = typeof nf === 'number' ? (nf === 14 ? book.loc.shortDate : nf === 22 ? book.loc.dateTime : XLSX.SSF._table[nf]) : nf;
      if (odf && o.z) {
        // the OpenDocument writer only knows simple number styles: dates become real dates, other codes are dropped
        if (o.t === 'n' && isDateFormat(o.z)) { const d = serialToDate(o.v); o.t = 'd'; o.v = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds())); }
        if (!/^(General|0|0\.0+|#,##0(\.0+)?|0%|0\.0+%)$/.test(o.z)) delete o.z;
      }
      if (cell.link && !valuesOnly) o.l = { Target: cell.link };
      if (cell.note && !valuesOnly) o.c = [{ a: cell.noteAuthor || '', t: cell.note }];
      ws[XLSX.utils.encode_cell({ r, c })] = o;
      if (r > r2) r2 = r; if (c > c2) c2 = c;
    };
    for (const [k, cell] of sh.cells) put(k, cell, false);
    for (const [k, sp] of sh.spill) if (!sh.cells.has(k)) put(k, { v: sp.v }, true);
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r2, c: c2 } });
    if (sh.merges.length) ws['!merges'] = sh.merges.map((m) => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
    const cols = []; for (const [c, w] of sh.colW) cols[c] = { wpx: w }; for (const c of sh.hiddenCols) cols[c] = { ...(cols[c] || { wpx: sh.defW }), hidden: true };
    if (cols.length) ws['!cols'] = cols;
    const rows = []; for (const [r, h] of sh.rowH) rows[r] = { hpx: h }; for (const r of [...sh.hiddenRows, ...sh.filterHidden]) rows[r] = { ...(rows[r] || {}), hidden: true };
    if (rows.length) ws['!rows'] = rows;
    if (sh.filter) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: sh.filter.range.r1, c: sh.filter.range.c1 }, e: { r: sh.filter.range.r2, c: sh.filter.range.c2 } }) };
    let name = sh.name.slice(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, name);
    wb.Workbook.Sheets[i] = { Hidden: sh.hidden === 'veryHidden' ? 2 : sh.hidden ? 1 : 0 };
  });
  for (const d of book.names.values()) {
    if (d.raw) continue;
    // OpenDocument named ranges are cell-range addresses only (SheetJS writes no named expressions)
    if (odf && !/^(?:(?:'(?:[^']|'')+'|[^'!:]+)!)?\$?[A-Z]*\$?\d*(?::\$?[A-Z]*\$?\d*)?$/i.test(String(d.ref).trim())) continue;
    wb.Workbook.Names.push({ Name: d.name, Ref: toOoxml(d.ref).replace(/_xlfn\./g, ''), ...(d.scope ? { Sheet: book.sheets.indexOf(d.scope) } : {}) });
  }
  return wb;
}

const BOOKTYPE = { xls: 'biff8', xlt: 'biff8', ods: 'ods', ots: 'ods', fods: 'fods', dif: 'dif', slk: 'sylk', dbf: 'dbf', xlsb: 'xlsb', html: 'html', txt: 'txt', prn: 'prn', numbers: 'numbers' };
export function writeSheetJS(book, ext, { sheet } = {}) {
  const odf = ['ods', 'ots', 'fods'].includes(ext);
  const formulas = odf ? [] : null;
  const wb = toSheetJS(book, { valuesOnly: ['dif', 'dbf', 'prn'].includes(ext), odf, formulas });
  const single = ['dif', 'slk', 'dbf', 'prn', 'txt', 'html'].includes(ext);
  if (single) { // one-sheet formats: the active (or given) sheet
    const idx = sheet != null ? sheet : book.active;
    const nm = wb.SheetNames[idx];
    wb.SheetNames = [nm]; wb.Sheets = { [nm]: wb.Sheets[nm] }; wb.Workbook.Sheets = [wb.Workbook.Sheets[idx]]; wb.Workbook.Names = [];
  }
  const raw = XLSX.write(wb, { bookType: BOOKTYPE[ext] || ext, type: 'array', compression: true });
  const out = raw instanceof Uint8Array ? raw : new Uint8Array(raw);
  return odf ? patchOdfOutput(out, formulas, ext === 'fods') : out;
}

/** quick text of a file (toText / search): TSV of the first sheets */
export function sheetjsText(bytes, maxSheets = 3) {
  const wb = XLSX.read(bytes, { type: 'array', cellFormula: false, cellNF: false, cellText: true, sheetRows: 2000 });
  return wb.SheetNames.slice(0, maxSheets).map((n) => `${wb.SheetNames.length > 1 ? `# ${n}\n` : ''}${XLSX.utils.sheet_to_csv(wb.Sheets[n], { FS: '\t', blankrows: false })}`).join('\n\n');
}
/** first ~rows×cols formatted texts of the first sheet (Preview) */
export function sheetjsGrid(bytes, rows = 50, cols = 15, loc = null) {
  const wb = XLSX.read(bytes, { type: 'array', cellText: true, cellNF: true, sheetRows: rows });
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const out = [];
  for (let r = 0; r < rows; r++) { const row = []; for (let c = 0; c < cols; c++) { const x = ws && ws[XLSX.utils.encode_cell({ r, c })]; row.push(x ? { t: loc && x.t === 'n' ? formatValue(x.v, x.z || 'General', loc).text : x.w != null ? x.w : String(x.v ?? ''), n: x.t === 'n' } : null); } out.push(row); }
  while (out.length && out[out.length - 1].every((x) => !x)) out.pop();
  let w = 0; for (const row of out) row.forEach((x, i) => { if (x) w = Math.max(w, i + 1); });
  return { name, sheets: wb.SheetNames, rows: out.map((r) => r.slice(0, Math.max(w, 1))) };
}

export { ERROR_TYPE };
