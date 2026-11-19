/* Open/save dispatcher for every spreadsheet format + CSV/TSV and HTML (export, print) writers.
   hostwerk-allow: tables – the HTML writer makes a stand-alone file for browsers, printers and other programs, not Hostwerk UI */
import { Book, newSheet } from '../model.js';
import { MAXC, colName } from '../formula/refs.js';
import { readXlsx, writeXlsx } from './xlsx.js';
import { readSheetJS, writeSheetJS } from './sheetjs.js';
import { sniff, parseCSV, writeCSV, encodeText } from './csv.js';
import { parseInput, parseNumberLoose } from '../format/parse.js';
import { makeLocale } from '../format/locale.js';
import { formatValue } from '../format/numfmt.js';
import { esc } from './xml.js';

export const OOXML = ['xlsx', 'xlsm', 'xltx', 'xltm'];
export const CSV = ['csv', 'tsv', 'tab'];
const K = (r, c) => r * MAXC + c;

/** settings.sheets.csvDelimiter → character ('auto' = by locale) */
export function csvDelimiter(setting, loc, ext) {
  if (ext === 'tsv' || ext === 'tab') return '\t';
  if (setting === 'tab' || setting === '\t') return '\t';
  if ([',', ';', '|'].includes(setting)) return setting;
  return loc && loc.decimal === ',' ? ';' : ',';
}

/** opts: { loc, settings (sheets section), csv: { delimiter, encoding, header, locale } } */
export async function openBytes(bytes, ext, opts = {}) {
  const { loc, settings = {} } = opts;
  if (OOXML.includes(ext)) {
    try { return await readXlsx(bytes, { loc, settings, ext }); } catch (e) {
      const b = readSheetJS(bytes, { loc, settings, ext });
      b.fidelity.unshift('This file could not be read completely – formatting may be missing');
      b.fidelity.push(String(e.message || e));
      return b;
    }
  }
  if (CSV.includes(ext)) return csvToBook(bytes, { loc, settings, ext, ...(opts.csv || {}) });
  return readSheetJS(bytes, { loc, settings, ext });
}

export function csvToBook(bytes, { loc, settings = {}, ext = 'csv', delimiter, encoding, numberLocale, header = false, name = 'Sheet1', keepText = false } = {}) {
  const enc = encoding || (settings.csvEncoding && settings.csvEncoding !== 'auto' ? settings.csvEncoding : 'auto');
  const delim = delimiter || (ext === 'tsv' || ext === 'tab' ? '\t' : settings.csvDelimiter && settings.csvDelimiter !== 'auto' ? csvDelimiter(settings.csvDelimiter, loc, ext) : 'auto');
  const sn = sniff(bytes, { encoding: enc, delimiter: delim });
  const rows = parseCSV(sn.text, { delimiter: sn.delimiter });
  const nloc = numberLocale ? makeLocale({ lang: numberLocale === ',' ? 'de' : 'en', decimal: numberLocale }) : loc;
  const book = new Book({ loc, settings });
  const sh = newSheet(name);
  if (settings.gridlines === false) sh.gridlines = false;
  const styles = new Map();
  rows.forEach((row, r) => row.forEach((t, c) => {
    if (t === '' || c >= MAXC) return;
    let cell;
    if (keepText || (header && r === 0)) cell = { v: t };
    else {
      const p = parseInput(t, nloc);
      cell = { v: p.v };
      if (p.fmt) { let s = styles.get(p.fmt); if (!s) { s = book.styles.intern({ nf: p.fmt }); styles.set(p.fmt, s); } cell.s = s; }
      if (typeof p.v === 'string' && p.v.startsWith('=')) cell.v = t; // never run formulas from a CSV file
    }
    sh.cells.set(K(r, c), cell);
  }));
  if (header && rows.length) { const bold = book.styles.intern({ b: true }); for (let c = 0; c < (rows[0] || []).length; c++) { const k = K(0, c); const x = sh.cells.get(k); if (x) sh.cells.set(k, { ...x, s: bold }); } sh.freeze = { rows: 1, cols: 0 }; }
  // column widths from the content (like opening a CSV in Excel, but a bit kinder)
  const widths = [];
  rows.slice(0, 200).forEach((row) => row.forEach((t, c) => { widths[c] = Math.max(widths[c] || 0, Math.min(60, String(t).length)); }));
  widths.forEach((w, c) => { if (w > 9) sh.colW.set(c, Math.min(400, Math.round(w * 7 + 12))); });
  book.sheets.push(sh);
  book.csv = { encoding: sn.encoding, delimiter: sn.delimiter };
  book.calc.rebuild();
  return book;
}

/** what a CSV cell should contain: displayed text like Excel ("as shown") */
function csvText(book, sh, r, c) {
  const v = book.value(sh, r, c);
  if (v == null) return '';
  return book.text(sh, r, c);
}
export function bookToCSV(book, { sheet = book.active, delimiter = ',', encoding = 'utf-8', eol = '\r\n' } = {}) {
  const sh = book.sheets[sheet];
  const g = book.dataRange(sh);
  const rows = [];
  if (g) for (let r = 0; r <= g.r2; r++) { const row = []; for (let c = 0; c <= g.c2; c++) row.push(csvText(book, sh, r, c)); rows.push(row); }
  return encodeText(writeCSV(rows, { delimiter, eol }), encoding);
}

/* ---------------- HTML (export + print) ---------------- */
const BORDER = { thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', medium: '2px solid', mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dotted', mediumDashDotDot: '2px dotted', slantDashDot: '2px dashed', thick: '3px solid', double: '3px double' };
function cellCss(s, loc, color, defFont) {
  if (!s && !color) return '';
  const st = [];
  s = s || {};
  if (s.font && s.font !== defFont.font) st.push(`font-family:'${s.font.replace(/'/g, '')}'`);
  if (s.size && s.size !== defFont.size) st.push(`font-size:${s.size}pt`);
  if (s.b) st.push('font-weight:bold');
  if (s.i) st.push('font-style:italic');
  const deco = [s.u ? 'underline' : '', s.st ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) st.push(`text-decoration:${deco}`);
  if (color || s.color) st.push(`color:${color || s.color}`);
  if (s.fill) st.push(`background:${s.fill}`);
  if (s.h && s.h !== 'general') st.push(`text-align:${s.h === 'centerContinuous' ? 'center' : s.h === 'distributed' || s.h === 'fill' ? 'justify' : s.h}`);
  if (s.va) st.push(`vertical-align:${s.va === 'center' ? 'middle' : s.va === 'top' ? 'top' : 'bottom'}`);
  if (s.wrap) st.push('white-space:pre-wrap'); else st.push('white-space:nowrap');
  if (s.indent) st.push(`padding-left:${s.indent * 9 + 3}px`);
  if (s.bd) for (const [k, p] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) if (s.bd[k]) st.push(`border-${p}:${BORDER[s.bd[k].s] || '1px solid'} ${s.bd[k].c || '#000'}`);
  return st.join(';');
}

/**
 * HTML document of a sheet (whole used range, a given range, or the print area).
 * opts: { sheet, range, gridlines, headings, orientation, fit, title, forPrint }
 */
export function bookToHTML(book, opts = {}) {
  const si = opts.sheet ?? book.active;
  const sh = book.sheets[si];
  const loc = book.loc;
  const defFont = book.props.defaultFont || { font: book.settings.font || 'Calibri', size: book.settings.fontSize || 11 };
  const used = book.usedRange(sh);
  const g = opts.range || (sh.print.area ? (() => { const m = /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(sh.print.area); if (!m) return null; const ci = (x) => x.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return { r1: +m[2] - 1, c1: ci(m[1]), r2: +m[4] - 1, c2: ci(m[3]) }; })() : null) || (used ? { r1: 0, c1: 0, r2: used.r2, c2: used.c2 } : { r1: 0, c1: 0, r2: 0, c2: 0 });
  const grid = opts.gridlines ?? sh.print.gridlines;
  const heads = opts.headings ?? sh.print.headings;
  const covered = new Set();
  const spans = new Map();
  for (const m of sh.merges) {
    if (m.r1 < g.r1 || m.c1 < g.c1 || m.r1 > g.r2 || m.c1 > g.c2) continue;
    spans.set(K(m.r1, m.c1), { rs: Math.min(m.r2, g.r2) - m.r1 + 1, cs: Math.min(m.c2, g.c2) - m.c1 + 1 });
    for (let r = m.r1; r <= Math.min(m.r2, g.r2); r++) for (let c = m.c1; c <= Math.min(m.c2, g.c2); c++) if (r !== m.r1 || c !== m.c1) covered.add(K(r, c));
  }
  let colgroup = heads ? '<col style="width:36px">' : '';
  let totalW = heads ? 36 : 0;
  for (let c = g.c1; c <= g.c2; c++) { if (sh.hiddenCols.has(c)) continue; const w = sh.colW.get(c) ?? sh.defW; totalW += w; colgroup += `<col style="width:${w}px">`; }
  let body = '';
  if (heads) { body += '<tr><th></th>'; for (let c = g.c1; c <= g.c2; c++) if (!sh.hiddenCols.has(c)) body += `<th>${colName(c)}</th>`; body += '</tr>'; }
  for (let r = g.r1; r <= g.r2; r++) {
    if (sh.hiddenRows.has(r) || sh.filterHidden.has(r)) continue;
    const h = sh.rowH.get(r) ?? sh.defH;
    body += `<tr style="height:${h}px">${heads ? `<th>${r + 1}</th>` : ''}`;
    for (let c = g.c1; c <= g.c2; c++) {
      if (sh.hiddenCols.has(c)) continue;
      const k = K(r, c);
      if (covered.has(k)) continue;
      const cell = sh.cells.get(k);
      const v = book.value(sh, r, c);
      const s = cell && cell.s;
      let text = ''; let color = null;
      if (v != null) { const f = formatValue(v, (s && s.nf) || 'General', loc); text = f.text; color = f.color; }
      const num = typeof v === 'number' && !(s && s.h);
      const span = spans.get(k);
      const css = cellCss(s, loc, color, defFont) + (num ? ';text-align:right' : '') + (typeof v === 'boolean' || (v && v.code) ? ';text-align:center' : '');
      const link = cell && (cell.link || cell.link_);
      const content = esc(text).replace(/\n/g, '<br>');
      body += `<td${span ? `${span.rs > 1 ? ` rowspan="${span.rs}"` : ''}${span.cs > 1 ? ` colspan="${span.cs}"` : ''}` : ''}${css ? ` style="${esc(css)}"` : ''}>${link && !link.startsWith('#') ? `<a href="${esc(link)}">${content}</a>` : content}</td>`;
    }
    body += '</tr>';
  }
  const fit = opts.fit ?? sh.print.fit;
  const orient = opts.orientation || sh.print.orientation || 'portrait';
  const pageW = orient === 'landscape' ? 1047 : 718; // A4 minus margins in CSS px
  const scale = fit && fit !== 'none' && totalW > pageW ? pageW / totalW : 1;
  const title = esc(opts.title || sh.name);
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title><style>
@page { size: A4 ${orient}; margin: 15mm 12mm; }
html,body { background:#fff; color:#000; margin:0; }
body { font-family:'${defFont.font}',Calibri,Arial,sans-serif; font-size:${defFont.size}pt; }
table { border-collapse:collapse; table-layout:fixed; width:${totalW}px; ${scale < 1 ? `transform:scale(${scale.toFixed(4)}); transform-origin:0 0;` : ''} }
td { padding:1px 3px; overflow:hidden; vertical-align:bottom; ${grid ? 'border:1px solid #d0d0d0;' : ''} }
th { background:#f0f0f0; border:1px solid #c0c0c0; font-weight:normal; font-size:9pt; color:#333; }
tr { page-break-inside:avoid; }
a { color:#0563c1; }
</style></head><body><table><colgroup>${colgroup}</colgroup>${body}</table></body></html>`;
}

/**
 * Book → bytes in ext. opts: { getValues (charts), sheet, settings, convert(bytes, from, to)?, preferConverter }
 * Returns { bytes, note? }.
 */
export async function saveBytes(book, ext, opts = {}) {
  if (OOXML.includes(ext)) return { bytes: await writeXlsx(book, { ext, getValues: opts.getValues }) };
  if (CSV.includes(ext) || ext === 'txt') {
    const delimiter = ext === 'csv' ? (book.csv && book.csv.delimiter && book.csv.delimiter !== '\t' ? book.csv.delimiter : csvDelimiter(opts.settings && opts.settings.csvDelimiter, book.loc, ext)) : '\t';
    const encoding = (book.csv && book.csv.encoding) || (opts.settings && opts.settings.csvEncoding && opts.settings.csvEncoding !== 'auto' ? opts.settings.csvEncoding : 'utf-8');
    return { bytes: bookToCSV(book, { sheet: opts.sheet ?? book.active, delimiter, encoding: ext === 'txt' ? 'utf-16le' : encoding }) };
  }
  if (ext === 'html' || ext === 'htm') return { bytes: new TextEncoder().encode(bookToHTML(book, { sheet: opts.sheet })) };
  // formats SheetJS writes without formatting: let the office engine convert a styled xlsx when it is there
  if (opts.convert && opts.preferConverter !== false && ['ods', 'xls', 'fods', 'ots', 'xlt', 'xlsb'].includes(ext)) {
    try {
      const x = await writeXlsx(book, { ext: 'xlsx', getValues: opts.getValues });
      const out = await opts.convert(x, 'xlsx', ext);
      if (out && out.length) return { bytes: out instanceof Uint8Array ? out : new Uint8Array(out) };
    } catch { /* fall back to the built-in writer */ }
  }
  const note = ['ods', 'fods', 'ots'].includes(ext) ? 'Saved without cell formatting (the built-in OpenDocument writer keeps values, formulas, number formats, merges and sizes)' : ext === 'xls' || ext === 'xlt' ? 'Saved as values without formulas and formatting (old .xls format) – use .xlsx to keep everything' : null;
  return { bytes: writeSheetJS(book, ext, { sheet: opts.sheet }), note };
}
