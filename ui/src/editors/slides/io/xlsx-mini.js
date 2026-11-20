/*
 * The embedded workbook of a chart (PowerPoint's "Edit data" opens it in Excel): a minimal, valid xlsx with the
 * chart's table on Sheet1 – A1 empty, B1… series names, A2… categories, values below. Synchronous (STORE zip) so
 * the pptx writer can call it while building a slide. Also the cell references the chart XML points to.
 */
import { zipStore } from '../../_shared/zipwrite.js';

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
/** column letters of a 0-based index */
export const colName = (i) => { let s = ''; let n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
/** references for series i (0-based): { name, cat, val } like Sheet1!$B$1 */
export function chartRefs(i, nCats) {
  const c = colName(i + 1);
  const last = Math.max(2, nCats + 1);
  return { name: `Sheet1!$${c}$1`, cat: `Sheet1!$A$2:$A$${last}`, val: `Sheet1!$${c}$2:$${c}$${last}` };
}

export function chartWorkbook(chart) {
  const cats = (chart && chart.categories) || [];
  const series = (chart && chart.series) || [];
  const rows = [];
  const cell = (r, c, v, isNum) => {
    const ref = `${colName(c)}${r}`;
    if (v == null || v === '') return '';
    if (isNum && Number.isFinite(Number(v))) return `<c r="${ref}"><v>${Number(v)}</v></c>`;
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  };
  rows.push(`<row r="1">${series.map((s, i) => cell(1, i + 1, s.name || `Series ${i + 1}`)).join('')}</row>`);
  cats.forEach((cat, k) => {
    rows.push(`<row r="${k + 2}">${cell(k + 2, 0, cat)}${series.map((s, i) => cell(k + 2, i + 1, (s.values || [])[k], true)).join('')}</row>`);
  });
  const lastCol = colName(Math.max(1, series.length));
  const dim = `A1:${lastCol}${Math.max(1, cats.length + 1)}`;
  const H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const files = {
    '[Content_Types].xml': `${H}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels': `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `${H}<workbook xmlns="${NS}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `${H}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `${H}<styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    'xl/worksheets/sheet1.xml': `${H}<worksheet xmlns="${NS}" xmlns:r="${R}"><dimension ref="${dim}"/><sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/></worksheet>`,
  };
  const enc = new TextEncoder();
  return zipStore(Object.entries(files).map(([name, s]) => ({ name, bytes: enc.encode(s) })));
}
