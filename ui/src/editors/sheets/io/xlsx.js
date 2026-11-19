/*
 * Own OOXML (xlsx/xlsm/xltx/xltm) reader and lossless writer.
 *
 * Reading builds the Book (values, formulas incl. shared/array, styles, merges, sizes, freeze, CF, validation, notes,
 * hyperlinks, filters, print setup, images, charts, defined names, tables) and keeps the original package.
 * Writing PATCHES the original package: only the parts we own are regenerated (workbook sheet list + names, styles
 * lists (original entries kept by index, new ones appended), shared strings (original entries kept), the edited parts
 * of each worksheet (sheetData, cols, merges, views, CF, validation, hyperlinks, filter, print setup), comments only
 * when notes changed, our own charts). Everything else – vbaProject.bin, pivot tables + caches, unchanged charts,
 * sparklines and x14 extensions, slicers, tables, external links, customXml, drawings/images, themes, printer
 * settings – stays byte-identical. calcChain.xml is dropped (Excel rebuilds it) and fullCalcOnLoad is set.
 */
import JSZip from 'jszip';
import { Book, newSheet, uid } from '../model.js';
import { MAXC, colIndex, colName, addr, parseRange, rangeA1 } from '../formula/refs.js';
import { shiftFormula, fromOoxml, toOoxml } from '../formula/parser.js';
import { errOf, isErr } from '../formula/errors.js';
import { BUILTIN_FORMATS } from '../format/numfmt.js';
import { readChartsFromXlsx, writeChartsIntoXlsx, patchOriginalCharts } from './xlsxCharts.js';
import { resolveChartData } from '../chartData.js';
import { esc, escText, unesc, unx, attrs, elements, first, attrOf, topChildren, textOf, parseRels, relsPathOf, resolvePath, relTarget, REL, addRel, nextRid, removeRels, ctOverride, ctDefault, ctRemove } from './xml.js';

const K = (r, c) => r * MAXC + c;
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const MAIN_CT = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  xltm: 'application/vnd.ms-excel.template.macroEnabled.main+xml',
};
const CT = {
  sheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  sst: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  metadata: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
};
const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
const pt2px = (pt) => Math.round((pt * 4) / 3);
const px2pt = (px) => Math.round(px * 0.75 * 100) / 100;
const w2px = (w) => Math.round(w * 7 + 5);
const px2w = (px) => Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);

/* ---------------- colours ---------------- */
function applyTint(hex, tint) {
  if (!tint) return hex;
  let r = parseInt(hex.slice(0, 2), 16) / 255; let g = parseInt(hex.slice(2, 4), 16) / 255; let b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b); const min = Math.min(r, g, b);
  let h = 0; let s = 0; let l = (max + min) / 2;
  if (max !== min) { const d = max - min; s = l > 0.5 ? d / (2 - max - min) : d / (max + min); h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  if (s === 0) { r = g = b = l; } else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
  const x = (v) => Math.round(v * 255).toString(16).padStart(2, '0');
  return (x(r) + x(g) + x(b)).toUpperCase();
}
function colorOf(a, theme) {
  if (!a) return null;
  let hex = null;
  if (a.rgb) hex = a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb.slice(-6);
  else if (a.theme != null) hex = theme[+a.theme] || null;
  else if (a.indexed != null) { const i = +a.indexed; hex = i === 64 || i === 65 ? null : INDEXED[i] || null; }
  if (!hex) return null;
  return `#${applyTint(hex.toUpperCase(), a.tint ? +a.tint : 0)}`;
}
const colorEl = (xml, tag, theme) => { const e = first(xml, tag); return e ? colorOf(attrs(e.a), theme) : null; };
const rgb = (c) => `FF${String(c).replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6)}`;

/* ---------------- styles: read ---------------- */
function readStyles(xml, theme) {
  const out = { numFmts: new Map(), fonts: [], fills: [], borders: [], xfs: [], dxfs: [], raw: { numFmts: [], fonts: [], fills: [], borders: [], xfs: [], dxfs: [] }, xml };
  if (!xml) return out;
  const sec = (name) => first(xml, name);
  const nf = sec('numFmts');
  if (nf) for (const e of elements(nf.inner, 'numFmt')) { const a = attrs(e.a); out.numFmts.set(+a.numFmtId, a.formatCode); out.raw.numFmts.push(nf.inner.slice(e.start, e.end)); }
  const fonts = sec('fonts');
  if (fonts) for (const e of elements(fonts.inner, 'font')) {
    const x = e.inner; const f = {};
    const v = (t) => { const q = first(x, t); return q ? attrs(q.a) : null; };
    const b = v('b'); const i = v('i'); const u = v('u'); const st = v('strike');
    if (b && b.val !== '0' && b.val !== 'false') f.b = true;
    if (i && i.val !== '0' && i.val !== 'false') f.i = true;
    if (u && u.val !== 'none') f.u = u.val === 'double' || u.val === 'doubleAccounting' ? 'double' : true;
    if (st && st.val !== '0' && st.val !== 'false') f.st = true;
    const sz = v('sz'); if (sz) f.size = +sz.val;
    const nm = v('name'); if (nm) f.font = nm.val;
    const va = v('vertAlign'); if (va && va.val !== 'baseline') f.vert = va.val;
    const col = colorEl(x, 'color', theme); if (col && col !== '#000000') f.color = col;
    out.fonts.push(f); out.raw.fonts.push(fonts.inner.slice(e.start, e.end));
  }
  const fills = sec('fills');
  if (fills) for (const e of elements(fills.inner, 'fill')) {
    const p = first(e.inner, 'patternFill');
    let fill = null;
    if (p) { const a = attrs(p.a); if (a.patternType && a.patternType !== 'none') fill = colorEl(p.inner, 'fgColor', theme) || (a.patternType === 'solid' ? null : colorEl(p.inner, 'bgColor', theme)); } else if (first(e.inner, 'gradientFill')) { const s = elements(e.inner, 'stop'); if (s.length) fill = colorEl(s[0].inner, 'color', theme); }
    out.fills.push(fill); out.raw.fills.push(fills.inner.slice(e.start, e.end));
  }
  const borders = sec('borders');
  if (borders) for (const e of elements(borders.inner, 'border')) {
    const bd = {};
    for (const [tag, k] of [['left', 'l'], ['right', 'r'], ['top', 't'], ['bottom', 'b'], ['diagonal', 'd']]) {
      const q = first(e.inner, tag);
      if (!q) continue;
      const a = attrs(q.a);
      if (a.style && a.style !== 'none') bd[k] = { s: a.style, c: colorEl(q.inner, 'color', theme) || '#000000' };
    }
    const ba = attrs(e.a);
    if (bd.d) { bd.d.up = ba.diagonalUp === '1'; bd.d.down = ba.diagonalDown === '1'; }
    out.borders.push(Object.keys(bd).length ? bd : null); out.raw.borders.push(borders.inner.slice(e.start, e.end));
  }
  const xfs = sec('cellXfs');
  if (xfs) for (const e of elements(xfs.inner, 'xf')) {
    const a = attrs(e.a); const al = first(e.inner, 'alignment'); const pr = first(e.inner, 'protection');
    out.xfs.push({ numFmtId: +(a.numFmtId || 0), fontId: +(a.fontId || 0), fillId: +(a.fillId || 0), borderId: +(a.borderId || 0), al: al ? attrs(al.a) : null, pr: pr ? attrs(pr.a) : null, quote: a.quotePrefix === '1' });
    out.raw.xfs.push(xfs.inner.slice(e.start, e.end));
  }
  const dxfs = sec('dxfs');
  if (dxfs) for (const e of elements(dxfs.inner, 'dxf')) {
    const d = {};
    const font = first(e.inner, 'font');
    if (font) { if (first(font.inner, 'b')) d.b = true; if (first(font.inner, 'i')) d.i = true; if (first(font.inner, 'strike')) d.st = true; if (first(font.inner, 'u')) d.u = true; const c = colorEl(font.inner, 'color', theme); if (c) d.color = c; }
    const p = first(e.inner, 'patternFill');
    if (p) { const c = colorEl(p.inner, 'bgColor', theme) || colorEl(p.inner, 'fgColor', theme); if (c) d.fill = c; }
    const nfe = first(e.inner, 'numFmt'); if (nfe) d.nf = attrs(nfe.a).formatCode;
    out.dxfs.push(d); out.raw.dxfs.push(dxfs.inner.slice(e.start, e.end));
  }
  return out;
}
function styleFromXf(st, i, defFont, book) {
  const x = st.xfs[i];
  if (!x) return null;
  const s = { _xf: i };
  const f = st.fonts[x.fontId] || {};
  for (const k of ['b', 'i', 'u', 'st', 'color', 'vert']) if (f[k]) s[k] = f[k];
  if (f.font && f.font !== defFont.font) s.font = f.font;
  if (f.size && f.size !== defFont.size) s.size = f.size;
  if (st.fills[x.fillId]) s.fill = st.fills[x.fillId];
  if (st.borders[x.borderId]) s.bd = st.borders[x.borderId];
  if (x.numFmtId) s.nf = x.numFmtId === 14 || x.numFmtId === 22 ? x.numFmtId : st.numFmts.get(x.numFmtId) || BUILTIN_FORMATS[x.numFmtId] || 'General';
  if (x.al) {
    if (x.al.horizontal && x.al.horizontal !== 'general') s.h = x.al.horizontal;
    if (x.al.vertical && x.al.vertical !== 'bottom') s.va = x.al.vertical;
    if (x.al.wrapText === '1' || x.al.wrapText === 'true') s.wrap = true;
    if (x.al.textRotation && x.al.textRotation !== '0') s.rot = +x.al.textRotation;
    if (x.al.indent && x.al.indent !== '0') s.indent = +x.al.indent;
    if (x.al.shrinkToFit === '1') s.shrink = true;
  }
  if (x.pr && x.pr.locked === '0') s.unlocked = true;
  if (x.pr && x.pr.hidden === '1') s.hiddenF = true;
  if (x.quote) s.quote = true;
  return book.styles.intern(s);
}

/* ---------------- reading ---------------- */
async function str(zip, p) { const f = zip.file(p); return f ? f.async('string') : null; }

export async function readXlsx(bytes, { loc, settings = {}, ext = 'xlsx' } = {}) {
  const zip = await JSZip.loadAsync(bytes);
  const book = new Book({ loc, settings });
  const fid = [];
  const ctXml = (await str(zip, '[Content_Types].xml')) || '';
  const rootRels = parseRels(await str(zip, '_rels/.rels'), '');
  const wbPath = (rootRels.find((r) => r.type === REL.officeDoc) || { path: 'xl/workbook.xml' }).path;
  const wbXml = await str(zip, wbPath);
  if (!wbXml) throw new Error('This is not an Excel workbook (xl/workbook.xml is missing).');
  const wbRels = parseRels(await str(zip, relsPathOf(wbPath)), wbPath);
  const wbPr = attrOf(first(wbXml, 'workbookPr'));
  book.date1904 = wbPr.date1904 === '1' || wbPr.date1904 === 'true';
  // theme
  let theme = DEFAULT_THEME.slice();
  const themeRel = wbRels.find((r) => r.type === REL.theme);
  const themeXml = themeRel && (await str(zip, themeRel.path));
  if (themeXml) {
    const cs = first(themeXml, 'clrScheme');
    if (cs) {
      const get = (n) => { const e = first(cs.inner, n); if (!e) return null; const s = first(e.inner, 'srgbClr'); if (s) return attrs(s.a).val; const y = first(e.inner, 'sysClr'); return y ? attrs(y.a).lastClr : null; };
      const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      theme = order.map((n, i) => (get(n) || DEFAULT_THEME[i]).toUpperCase());
    }
  }
  const stylesRel = wbRels.find((r) => r.type === REL.styles);
  const st = readStyles(stylesRel && (await str(zip, stylesRel.path)), theme);
  const defFont = { font: (st.fonts[0] && st.fonts[0].font) || 'Calibri', size: (st.fonts[0] && st.fonts[0].size) || 11 };
  book.props.defaultFont = defFont;
  const xfStyle = st.xfs.map((_, i) => (i === 0 ? null : styleFromXf(st, i, defFont, book)));
  // xf 0 may carry formatting too (rare) – keep it as a style when it differs from the defaults
  if (st.xfs[0]) { const s0 = styleFromXf(st, 0, defFont, book); if (s0 && Object.keys(s0).length > 1) xfStyle[0] = s0; }
  // shared strings
  const sstRel = wbRels.find((r) => r.type === REL.sst);
  const sstXml = sstRel && (await str(zip, sstRel.path));
  const sst = []; const sstRaw = [];
  if (sstXml) for (const e of elements(sstXml, 'si')) { sst.push(textOf(e.inner)); sstRaw.push(sstXml.slice(e.start, e.end)); }
  // metadata (dynamic arrays)
  const metaRel = wbRels.find((r) => r.type === REL.metadata);
  const metaXml = metaRel && (await str(zip, metaRel.path));
  // sheets
  const sheetTags = elements(first(wbXml, 'sheets').inner, 'sheet');
  const names = first(wbXml, 'definedNames');
  const nameEls = names ? elements(names.inner, 'definedName') : [];
  const charts = await readChartsFromXlsx(zip).catch(() => new Map());
  const info = new Map();
  const sheetsList = [];
  for (const t of sheetTags) {
    const a = attrs(t.a);
    const rid = a['r:id'] || a.id;
    const rel = wbRels.find((r) => r.id === rid);
    if (!rel || rel.type !== REL.sheet) { fid.push(`Chart sheet or dialog sheet "${a.name}" (kept, not shown)`); sheetsList.push({ keep: true, tag: t, a, rel }); continue; }
    const sh = newSheet(a.name);
    sh.hidden = a.state === 'hidden' ? 'hidden' : a.state === 'veryHidden' ? 'veryHidden' : false;
    const xml = await str(zip, rel.path);
    const srels = parseRels(await str(zip, relsPathOf(rel.path)), rel.path);
    const meta = { path: rel.path, rid, sheetId: a.sheetId, rels: srels, notesSig: '', tables: [], hasMeta: !!metaXml };
    await readSheet(zip, xml || '', sh, { book, st, xfStyle, sst, srels, meta, fid, theme, defFont, charts: charts.get(a.name) || [] });
    info.set(sh.id, meta);
    book.sheets.push(sh);
    sheetsList.push({ sheet: sh });
  }
  // defined names
  for (const e of nameEls) {
    const a = attrs(e.a);
    const ref = fromOoxml(unesc(e.inner));
    const scope = a.localSheetId != null ? book.sheets[+a.localSheetId] || null : null;
    if (a.name === '_xlnm.Print_Area' && scope) { scope.print.area = ref.replace(/^.*!/, '').replace(/\$/g, ''); continue; }
    if (a.name === '_xlnm._FilterDatabase') continue;
    book.names.set(`${a.name.toUpperCase()}|${scope ? scope.id : ''}`, { name: a.name, ref, scope, comment: a.comment || '', hidden: a.hidden === '1', raw: a.name.startsWith('_xlnm.') || a.name.startsWith('_xlfn.') });
  }
  const bv = attrOf(first(wbXml, 'workbookView'));
  book.active = Math.min(book.sheets.length - 1, Math.max(0, +(bv.activeTab || 0)));
  const calcPr = attrOf(first(wbXml, 'calcPr'));
  if (calcPr.calcMode === 'manual') book.fileManualCalc = true;
  // fidelity notes (everything listed is kept unchanged when saving)
  const files = Object.keys(zip.files);
  const has = (re) => files.some((f) => re.test(f));
  if (has(/vbaProject\.bin$/i)) fid.push('Macros (VBA) – kept, not run');
  if (has(/pivotTables\//)) fid.push('Pivot tables – shown with their stored values, kept; refresh them in Excel');
  if (has(/slicers?\//)) fid.push('Slicers – kept');
  if (has(/externalLinks\//)) fid.push('Links to other workbooks – kept, show their stored values');
  if (has(/threadedComments\//)) fid.push('Threaded comments – kept; edits here change the classic notes');
  if (has(/ctrlProps\/|activeX\//)) fid.push('Form controls / ActiveX – kept, not shown');
  if (has(/model\/item\.data|connections\.xml|queryTables\//)) fid.push('Data model / queries – kept, not refreshed');
  if (/sparklineGroup/.test(ctXml) || false) fid.push('Sparklines – kept, not shown');
  const origCharts = new Map(book.sheets.map((sh) => [sh.id, sh.charts.slice()]));
  book.pkg = { origCharts, bytes, ext, wbPath, wbXml, wbRels, ctXml, st, sstRaw, sst, info, metaXml, theme, sheetsList, origSheetIds: new Set(book.sheets.map((s) => s.id)) };
  book.fidelity = [...new Set(fid)];
  book.calc.rebuild();
  // keep the file's cached values where we can't compute (unknown functions) – our values replace them otherwise
  book.calc.recalc(null, { force: true });
  restoreCachedForUnknown(book);
  return book;
}
function restoreCachedForUnknown(book) {
  for (const sh of book.sheets) for (const cell of sh.cells.values()) if (cell.f && cell._cached !== undefined && isErr(cell.v) && cell.v.code === '#NAME?' && !(isErr(cell._cached) && cell._cached.code === '#NAME?')) cell.v = cell._cached;
}

async function readSheet(zip, xml, sh, ctx) {
  const { book, st, xfStyle, sst, srels, meta, fid, theme } = ctx;
  const pr = first(xml, 'sheetPr');
  if (pr) { const tc = first(pr.inner, 'tabColor'); if (tc) sh.color = colorOf(attrs(tc.a), theme); }
  const view = first(xml, 'sheetView');
  if (view) {
    const a = attrs(view.a);
    if (a.showGridLines === '0' || a.showGridLines === 'false') sh.gridlines = false;
    if (a.showRowColHeaders === '0') sh.headings = false;
    if (a.zoomScale) sh.zoom = +a.zoomScale;
    if (a.rightToLeft === '1') sh.rtl = true;
    const pane = first(view.inner, 'pane');
    if (pane) { const p = attrs(pane.a); if (p.state === 'frozen' || p.state === 'frozenSplit') { sh.freeze = { rows: Math.round(+(p.ySplit || 0)), cols: Math.round(+(p.xSplit || 0)) }; } }
    const sel = first(view.inner, 'selection');
    if (sel) { const s = attrs(sel.a); const g = s.activeCell && parseRange(s.activeCell); if (g) sh.sel = { r: g.r1, c: g.c1 }; }
  }
  const fmt = attrOf(first(xml, 'sheetFormatPr'));
  if (fmt.defaultRowHeight) sh.defH = pt2px(+fmt.defaultRowHeight);
  if (fmt.defaultColWidth) sh.defW = w2px(+fmt.defaultColWidth); else if (fmt.baseColWidth) sh.defW = w2px(+fmt.baseColWidth + 0.71);
  const cols = first(xml, 'cols');
  sh._colXf = new Map();
  if (cols) for (const e of elements(cols.inner, 'col')) {
    const a = attrs(e.a);
    const lo = +a.min - 1; const hi = Math.min(+a.max - 1, MAXC - 1);
    const wpx = a.width != null ? w2px(+a.width) : null;
    const span = hi - lo + 1;
    if (span > 2000) { if (wpx != null && a.customWidth === '1' && hi >= MAXC - 2) sh.defW = sh.defW; continue; }
    for (let c = lo; c <= hi; c++) {
      if (wpx != null && (a.customWidth === '1' || Math.abs(wpx - sh.defW) > 1)) sh.colW.set(c, wpx);
      if (a.hidden === '1' || a.hidden === 'true') sh.hiddenCols.add(c);
      if (a.outlineLevel) sh.outlineCols.set(c, +a.outlineLevel);
      if (a.style && a.style !== '0') sh._colXf.set(c, +a.style);
    }
  }
  // autofilter first (rows hidden by a filter are "filtered", not manually hidden)
  const af = first(xml, 'autoFilter');
  let afRange = null;
  if (af) {
    afRange = parseRange(attrs(af.a).ref);
    if (afRange) {
      sh.filter = { range: afRange, cols: {} };
      for (const fc of elements(af.inner, 'filterColumn')) {
        const ci = afRange.c1 + +attrs(fc.a).colId;
        const fl = first(fc.inner, 'filters');
        const cf = first(fc.inner, 'customFilters');
        if (fl) { const vals = elements(fl.inner, 'filter').map((e) => attrs(e.a).val); if (attrs(fl.a).blank === '1') vals.push(''); sh.filter.cols[ci] = { values: vals }; } else if (cf) {
          const list = elements(cf.inner, 'customFilter').map((e) => attrs(e.a));
          const op = (o) => ({ equal: 'eq', notEqual: 'ne', greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le' }[o || 'equal'] || 'eq');
          const fixOp = (o, v) => { if (op(o) === 'eq' && /^\*.*\*$/.test(v)) return ['contains', v.slice(1, -1)]; if (op(o) === 'eq' && v.endsWith('*')) return ['begins', v.slice(0, -1)]; if (op(o) === 'eq' && v.startsWith('*')) return ['ends', v.slice(1)]; return [op(o), v]; };
          const [o1, v1] = fixOp(list[0] && list[0].operator, list[0] ? list[0].val : '');
          const c2 = list[1] ? fixOp(list[1].operator, list[1].val) : null;
          sh.filter.cols[ci] = { cond: { op: o1, v: v1, join: attrs(cf.a).and === '1' ? 'and' : 'or', op2: c2 && c2[0], v2: c2 && c2[1] } };
        } else sh.filter.cols[ci] = { raw: fc.inner };
      }
    }
  }
  // sheet data
  const sd = first(xml, 'sheetData');
  const shared = new Map();
  if (sd) {
    const data = sd.inner;
    const rowRe = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let rm; let rowIdx = -1;
    while ((rm = rowRe.exec(data))) {
      const ra = attrs(rm[1]);
      rowIdx = ra.r ? +ra.r - 1 : rowIdx + 1;
      const r = rowIdx;
      if (ra.ht && (ra.customHeight === '1' || ra.customHeight === 'true')) sh.rowH.set(r, pt2px(+ra.ht));
      else if (ra.ht && Math.abs(pt2px(+ra.ht) - sh.defH) > 2) sh.rowH.set(r, pt2px(+ra.ht));
      if (ra.hidden === '1' || ra.hidden === 'true') { if (afRange && r > afRange.r1 && r <= afRange.r2) sh.filterHidden.add(r); else sh.hiddenRows.add(r); }
      if (ra.outlineLevel) sh.outlineRows.set(r, +ra.outlineLevel);
      const rowXf = ra.customFormat === '1' && ra.s ? +ra.s : null;
      if (rowXf) { if (!sh._rowXf) sh._rowXf = new Map(); sh._rowXf.set(r, rowXf); }
      const body = rm[3];
      if (!body) continue;
      cellRe.lastIndex = 0;
      let cm; let colIdx = -1;
      while ((cm = cellRe.exec(body))) {
        const ca = attrs(cm[1]);
        let c;
        if (ca.r) { const m = /^([A-Z]+)(\d+)$/.exec(ca.r); c = colIndex(m[1]); } else c = colIdx + 1;
        colIdx = c;
        const inner = cm[2] || '';
        const cell = {};
        const sIdx = ca.s ? +ca.s : 0;
        const style = xfStyle[sIdx];
        if (style) cell.s = style;
        const vEl = /<v>([\s\S]*?)<\/v>/.exec(inner);
        const raw = vEl ? unesc(vEl[1]) : null;
        const t = ca.t || 'n';
        let v = null;
        if (t === 's') { const i = +raw; v = sst[i] ?? ''; cell._si = i; } else if (t === 'inlineStr') { const is = /<is>([\s\S]*?)<\/is>/.exec(inner); v = is ? textOf(is[1]) : ''; } else if (t === 'str') v = raw == null ? '' : unx(raw);
        else if (t === 'b') v = raw === '1' || raw === 'true';
        else if (t === 'e') v = raw ? errOf(raw) : null;
        else if (t === 'd') { const d = raw ? new Date(raw) : null; v = d && !Number.isNaN(+d) ? (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - Date.UTC(1899, 11, 30)) / 86400000 : null; } else v = raw == null || raw === '' ? null : Number(raw);
        cell.v = v;
        const fEl = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(inner);
        if (fEl) {
          const fa = attrs(fEl[1]);
          const ftext = fEl[2] != null ? unx(unesc(fEl[2])) : '';
          if (fa.t === 'shared') {
            if (ftext) { shared.set(fa.si, { f: fromOoxml(ftext), r, c }); cell.f = fromOoxml(ftext); } else { const m = shared.get(fa.si); if (m) cell.f = shiftFormula(m.f, r - m.r, c - m.c); }
          } else if (fa.t === 'array') {
            cell.f = fromOoxml(ftext);
            const g = parseRange(fa.ref);
            if (g && (g.r2 > g.r1 || g.c2 > g.c1)) { cell.arr = rangeA1(g); if (!sh._arr) sh._arr = []; sh._arr.push(g); }
          } else if (fa.t === 'dataTable') { fid.push('What-if data tables – values kept, not recalculated'); } else if (ftext) cell.f = fromOoxml(ftext);
          if (cell.f) { cell._cached = v; cell._cachedF = cell.f; }
        }
        if (cell.v == null && !cell.f && !cell.s) continue;
        sh.cells.set(K(r, c), cell);
      }
    }
  }
  // cells covered by legacy array formulas / spills keep only the anchor (their values come from the spill)
  if (sh._arr) for (const g of sh._arr) for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { if (r === g.r1 && c === g.c1) continue; const k = K(r, c); const x = sh.cells.get(k); if (x && !x.f) { if (x.s) sh.cells.set(k, { v: null, s: x.s }); else sh.cells.delete(k); } }
  // merges
  const mc = first(xml, 'mergeCells');
  if (mc) for (const e of elements(mc.inner, 'mergeCell')) { const g = parseRange(attrs(e.a).ref); if (g) sh.merges.push({ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }); }
  // conditional formatting
  for (const cfe of elements(xml, 'conditionalFormatting')) {
    const a = attrs(cfe.a);
    if (/x14:|xm:/.test(cfe.name)) continue;
    const ranges = String(a.sqref || '').split(/\s+/).map(parseRange).filter(Boolean);
    for (const re of elements(cfe.inner, 'cfRule', true)) {
      const ra = attrs(re.a);
      const fs = elements(re.inner, 'formula').map((x) => fromOoxml(unesc(x.inner)));
      const style = ra.dxfId != null ? { ...(st.dxfs[+ra.dxfId] || {}) } : null;
      const rule = { id: uid('cf'), ranges, priority: +(ra.priority || 0), stop: ra.stopIfTrue === '1', style };
      const OPS = { greaterThan: 'gt', lessThan: 'lt', between: 'between', notBetween: 'notBetween', equal: 'eq', notEqual: 'ne', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le' };
      switch (ra.type) {
        case 'cellIs': Object.assign(rule, { type: 'cell', op: OPS[ra.operator] || 'eq', f1: fs[0] || '', f2: fs[1] || '' }); break;
        case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': Object.assign(rule, { type: 'text', op: { containsText: 'contains', notContainsText: 'notContains', beginsWith: 'begins', endsWith: 'ends' }[ra.type], text: ra.text || '' }); break;
        case 'duplicateValues': rule.type = 'dup'; break;
        case 'uniqueValues': rule.type = 'unique'; break;
        case 'top10': Object.assign(rule, { type: ra.bottom === '1' ? 'bottom' : 'top', rank: +(ra.rank || 10), percent: ra.percent === '1' }); break;
        case 'aboveAverage': Object.assign(rule, { type: ra.aboveAverage === '0' ? 'below' : 'above' }); break;
        case 'containsBlanks': rule.type = 'blank'; break;
        case 'notContainsBlanks': rule.type = 'notBlank'; break;
        case 'containsErrors': rule.type = 'error'; break;
        case 'expression': Object.assign(rule, { type: 'formula', formula: fs[0] || '' }); break;
        case 'colorScale': {
          const cs = first(re.inner, 'colorScale');
          const cfvo = elements(cs.inner, 'cfvo').map((x) => attrs(x.a));
          const colors = elements(cs.inner, 'color').map((x) => colorOf(attrs(x.a), theme) || '#FFFFFF');
          Object.assign(rule, { type: 'scale', colors, stops: cfvo.map((x) => ({ type: x.type, val: x.val })) });
          break;
        }
        case 'dataBar': {
          const db = first(re.inner, 'dataBar');
          Object.assign(rule, { type: 'bar', color: colorEl(db.inner, 'color', theme) || '#638EC6', stops: elements(db.inner, 'cfvo').map((x) => { const q = attrs(x.a); return { type: q.type, val: q.val }; }) });
          break;
        }
        default: Object.assign(rule, { type: 'raw', kind: ra.type }); if (ra.type === 'iconSet') fid.push('Icon sets in conditional formatting – kept, not shown'); break;
      }
      rule._raw = cfe.inner.slice(re.start, re.end);
      rule._rawSig = JSON.stringify(sigOf(rule));
      sh.cf.push(rule);
    }
  }
  // data validation
  const dvs = first(xml, 'dataValidations');
  if (dvs) for (const e of elements(dvs.inner, 'dataValidation')) {
    const a = attrs(e.a);
    const f1 = first(e.inner, 'formula1'); const f2 = first(e.inner, 'formula2');
    sh.dv.push({
      id: uid('dv'), ranges: String(a.sqref || '').split(/\s+/).map(parseRange).filter(Boolean), type: a.type || 'any', op: a.operator || 'between',
      f1: f1 ? fromOoxml(unesc(f1.inner)) : '', f2: f2 ? fromOoxml(unesc(f2.inner)) : '', allowBlank: a.allowBlank === '1', dropdown: a.showDropDown !== '1',
      prompt: a.showInputMessage === '1' || a.prompt ? { title: a.promptTitle || '', text: a.prompt || '' } : null,
      error: { style: a.errorStyle || 'stop', title: a.errorTitle || '', text: a.error || '', show: a.showErrorMessage === '1' },
    });
  }
  // hyperlinks
  const hl = first(xml, 'hyperlinks');
  if (hl) for (const e of elements(hl.inner, 'hyperlink')) {
    const a = attrs(e.a);
    const g = parseRange(a.ref); if (!g) continue;
    const rel = a['r:id'] && srels.find((x) => x.id === a['r:id']);
    const link = rel ? rel.target : a.location ? `#${a.location}` : null;
    if (!link) continue;
    const k = K(g.r1, g.c1);
    const cell = sh.cells.get(k) || { v: null };
    sh.cells.set(k, { ...cell, link });
  }
  // print
  const po = attrOf(first(xml, 'printOptions'));
  if (po.gridLines === '1' || po.gridLines === 'true') sh.print.gridlines = true;
  if (po.headings === '1' || po.headings === 'true') sh.print.headings = true;
  const ps = attrOf(first(xml, 'pageSetup'));
  if (ps.orientation) sh.print.orientation = ps.orientation;
  const fitPage = pr && /fitToPage="(1|true)"/.test(pr.inner);
  sh.print.fit = fitPage ? (ps.fitToHeight === '0' ? 'width' : 'page') : 'none';
  // notes
  const cRel = srels.find((r) => r.type === REL.comments);
  if (cRel) {
    const cx = await str(zip, cRel.path);
    if (cx) {
      const authors = elements(first(cx, 'authors')?.inner || '', 'author').map((x) => unesc(x.inner));
      for (const e of elements(cx, 'comment')) {
        const a = attrs(e.a); const g = parseRange(a.ref); if (!g) continue;
        const text = textOf(e.inner);
        const k = K(g.r1, g.c1);
        const cell = sh.cells.get(k) || { v: null };
        sh.cells.set(k, { ...cell, note: text, noteAuthor: authors[+a.authorId] || '' });
      }
    }
  }
  meta.notesSig = notesSig(sh);
  // images in the drawing (kept by the writer, shown here)
  const dRel = srels.find((r) => r.type === REL.drawing);
  if (dRel) {
    const dx = await str(zip, dRel.path);
    if (dx) {
      const drels = parseRels(await str(zip, relsPathOf(dRel.path)), dRel.path);
      for (const an of [...elements(dx, 'twoCellAnchor', true), ...elements(dx, 'oneCellAnchor', true)]) {
        const pic = first(an.inner, 'pic', true);
        if (!pic) { if (!/<(\w+:)?graphicFrame/.test(an.inner)) { if (/<(\w+:)?sp\b|<(\w+:)?grpSp\b|<(\w+:)?cxnSp\b/.test(an.inner)) fid.push('Shapes and text boxes – kept, not shown'); } continue; }
        const blip = /r:embed="([^"]+)"/.exec(pic.inner);
        const rel = blip && drels.find((x) => x.id === blip[1]);
        if (!rel || !rel.path || !zip.file(rel.path)) continue;
        const b64 = await zip.file(rel.path).async('base64');
        const extm = /\.(\w+)$/.exec(rel.path); const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', emf: 'image/emf', wmf: 'image/wmf' }[(extm ? extm[1] : '').toLowerCase()] || 'image/png';
        const pt = (tag) => { const e = first(an.inner, tag); if (!e) return null; const g = (t) => +((first(e.inner, t) || { inner: '0' }).inner); return { c: g('col'), r: g('row'), dx: Math.round(g('colOff') / 9525), dy: Math.round(g('rowOff') / 9525) }; };
        const from = pt('from'); const to = pt('to');
        const ext = first(an.inner, 'ext');
        const size = to ? null : ext ? { w: Math.round(+attrs(ext.a).cx / 9525), h: Math.round(+attrs(ext.a).cy / 9525) } : { w: 200, h: 150 };
        sh.images.push({ id: uid('i'), src: `data:${mime};base64,${b64}`, anchor: anchorFrom(sh, from, to, size), _orig: true });
      }
    }
  }
  // charts (from xlsxCharts) → anchor in the editor's form
  for (const ch of ctx.charts) {
    const a = ch.anchor || {};
    ch.anchor = anchorFrom(sh, a.from || { c: 0, r: 0, dx: 0, dy: 0 }, a.to || null, a.ext || { w: 480, h: 288 });
    ch.stacked = ch.stacked || false;
    ch._sig = JSON.stringify(chartSig(ch));
    if (ch._unsupported) fid.push('Some chart types (radar, stock, surface, bubble) – kept, shown simplified');
    sh.charts.push(ch);
  }
  // tables (ListObjects) – kept, refs patched on save
  for (const r of srels.filter((x) => x.type === REL.table)) {
    const tx = await str(zip, r.path);
    if (!tx) continue;
    const ta = attrs(/<table\b[^>]*>/.exec(tx)[0]);
    const g = parseRange(ta.ref);
    if (g) meta.tables.push({ path: r.path, g, name: ta.displayName || ta.name, header: ta.headerRowCount !== '0' });
    sh.tables = sh.tables || [];
    if (g) sh.tables.push({ path: r.path, name: ta.displayName || ta.name, range: g });
  }
  if (srels.some((x) => x.type === REL.pivot)) sh.hasPivot = true;
}
function anchorFrom(sh, from, to, size) {
  const colX = (c) => { let x = 0; for (let i = 0; i < c; i++) x += sh.hiddenCols.has(i) ? 0 : sh.colW.get(i) ?? sh.defW; return x; };
  const rowY = (r) => { let y = 0; for (let i = 0; i < r; i++) y += sh.hiddenRows.has(i) ? 0 : sh.rowH.get(i) ?? sh.defH; return y; };
  let w = size ? size.w : 0; let h = size ? size.h : 0;
  if (to) { w = colX(to.c) + to.dx - colX(from.c) - from.dx; h = rowY(to.r) + to.dy - rowY(from.r) - from.dy; }
  return { r: from.r, c: from.c, dx: from.dx, dy: from.dy, w: Math.max(20, w), h: Math.max(20, h) };
}
const notesSig = (sh) => { const o = []; for (const [k, c] of sh.cells) if (c.note) o.push(`${k}\u0001${c.note}`); return o.sort().join('\u0002'); };
const sigOf = (rule) => { const { id, _raw, _rawSig, ...rest } = rule; return rest; };
export const chartSig = (ch) => { const { id, _orig, _sig, _xl, _dirty, ...rest } = ch; return rest; };

/* ---------------- styles: write ---------------- */
function fontXml(s, def) {
  let x = '<font>';
  if (s.b) x += '<b/>';
  if (s.i) x += '<i/>';
  if (s.st) x += '<strike/>';
  if (s.u) x += s.u === 'double' ? '<u val="double"/>' : '<u/>';
  if (s.vert) x += `<vertAlign val="${s.vert}"/>`;
  x += `<sz val="${s.size || def.size}"/>`;
  x += s.color ? `<color rgb="${rgb(s.color)}"/>` : '<color theme="1"/>';
  x += `<name val="${esc(s.font || def.font)}"/><family val="2"/>`;
  return `${x}</font>`;
}
const fillXml = (c) => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb(c)}"/><bgColor indexed="64"/></patternFill></fill>`;
function borderXml(bd) {
  const one = (tag, e) => (e ? `<${tag} style="${e.s}"><color rgb="${rgb(e.c || '#000000')}"/></${tag}>` : `<${tag}/>`);
  const d = bd.d;
  return `<border${d && d.up ? ' diagonalUp="1"' : ''}${d && d.down ? ' diagonalDown="1"' : ''}>${one('left', bd.l)}${one('right', bd.r)}${one('top', bd.t)}${one('bottom', bd.b)}${one('diagonal', bd.d)}</border>`;
}

class StyleWriter {
  constructor(st, defFont) {
    this.def = defFont;
    this.numFmts = st ? st.raw.numFmts.slice() : [];
    this.fonts = st && st.raw.fonts.length ? st.raw.fonts.slice() : [fontXml({}, defFont)];
    this.fills = st && st.raw.fills.length ? st.raw.fills.slice() : ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = st && st.raw.borders.length ? st.raw.borders.slice() : ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.xfs = st && st.raw.xfs.length ? st.raw.xfs.slice() : ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.dxfs = st ? st.raw.dxfs.slice() : [];
    this.origXfs = st ? st.raw.xfs.length : 0;
    this.nfIds = new Map(); // code → id
    if (st) for (const [id, code] of st.numFmts) this.nfIds.set(code, id);
    this.nextNf = Math.max(163, ...(st ? [...st.numFmts.keys()] : [])) + 1;
    this.cache = new Map();
    this.index = (arr, x) => { let i = arr.indexOf(x); if (i < 0) { arr.push(x); i = arr.length - 1; } return i; };
    this.changed = false;
  }
  nfId(nf) {
    if (nf == null || nf === 'General') return 0;
    if (typeof nf === 'number') return nf;
    for (const [id, code] of Object.entries(BUILTIN_FORMATS)) if (code === nf) return +id;
    let id = this.nfIds.get(nf);
    if (id == null) { id = this.nextNf++; this.nfIds.set(nf, id); this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${esc(nf)}"/>`); this.changed = true; }
    return id;
  }
  xf(s) {
    if (!s) return 0;
    if (s._xf != null && s._xf < this.origXfs) return s._xf;
    let i = this.cache.get(s);
    if (i != null) return i;
    this.changed = true;
    const fontId = this.index(this.fonts, fontXml(s, this.def));
    const fillId = s.fill ? this.index(this.fills, fillXml(s.fill)) : 0;
    const borderId = s.bd ? this.index(this.borders, borderXml(s.bd)) : 0;
    const numFmtId = this.nfId(s.nf);
    const al = [];
    if (s.h) al.push(`horizontal="${s.h}"`);
    if (s.va) al.push(`vertical="${s.va}"`);
    if (s.wrap) al.push('wrapText="1"');
    if (s.rot) al.push(`textRotation="${s.rot}"`);
    if (s.indent) al.push(`indent="${s.indent}"`);
    if (s.shrink) al.push('shrinkToFit="1"');
    const prot = s.unlocked || s.hiddenF ? `<protection${s.unlocked ? ' locked="0"' : ''}${s.hiddenF ? ' hidden="1"' : ''}/>` : '';
    const x = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''} applyFont="1"${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${al.length ? ' applyAlignment="1"' : ''}${prot ? ' applyProtection="1"' : ''}${s.quote ? ' quotePrefix="1"' : ''}>${al.length ? `<alignment ${al.join(' ')}/>` : ''}${prot}</xf>`;
    i = this.index(this.xfs, x);
    this.cache.set(s, i);
    return i;
  }
  dxf(style) {
    if (!style) return null;
    let x = '<dxf>';
    if (style.b || style.i || style.st || style.u || style.color) x += `<font>${style.b ? '<b/>' : ''}${style.i ? '<i/>' : ''}${style.st ? '<strike/>' : ''}${style.u ? '<u/>' : ''}${style.color ? `<color rgb="${rgb(style.color)}"/>` : ''}</font>`;
    if (style.nf) x += `<numFmt numFmtId="${this.nfId(style.nf)}" formatCode="${esc(String(style.nf))}"/>`;
    if (style.fill) x += `<fill><patternFill patternType="solid"><fgColor rgb="${rgb(style.fill)}"/><bgColor rgb="${rgb(style.fill)}"/></patternFill></fill>`;
    if (style.bd) x += borderXml(style.bd);
    x += '</dxf>';
    this.changed = true;
    return this.index(this.dxfs, x);
  }
  xml(original) {
    const sec = (name, items, extra = '') => `<${name} count="${items.length}"${extra}>${items.join('')}</${name}>`;
    const parts = {
      numFmts: this.numFmts.length ? sec('numFmts', this.numFmts) : null,
      fonts: sec('fonts', this.fonts), fills: sec('fills', this.fills), borders: sec('borders', this.borders),
      cellXfs: sec('cellXfs', this.xfs), dxfs: sec('dxfs', this.dxfs),
    };
    const base = original || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${NS_MAIN}"><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    return assemble(base, parts, ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst']);
  }
}

/** replace / insert direct children of the root element keeping the schema order; unknown children stay where they are */
function assemble(xml, parts, order) {
  const tc = topChildren(xml);
  const slot = (name) => { const i = order.indexOf(name); return i < 0 ? null : i; };
  const items = [];
  let last = -1;
  tc.kids.forEach((k, n) => {
    const s = slot(k.name);
    const pos = s == null ? last + 0.001 * (n + 1) : s;
    if (s != null) last = s;
    if (k.name in parts) return; // replaced
    items.push({ pos, n, xml: k.xml });
  });
  for (const [name, x] of Object.entries(parts)) {
    if (x == null) continue;
    const s = slot(name);
    const list = Array.isArray(x) ? x : [x];
    list.forEach((xx, j) => items.push({ pos: s == null ? 999 : s, n: 10000 + j, xml: xx }));
  }
  items.sort((a, b) => a.pos - b.pos || a.n - b.n);
  return tc.head + items.map((i) => i.xml).join('') + tc.tail;
}

/* ---------------- writing ---------------- */
const SHEET_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
const BLANK_SHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
const META_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="${NS_MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;

function skeleton(ext) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${MAIN_CT[ext] || MAIN_CT.xlsx}"/><Override PartName="/xl/styles.xml" ContentType="${CT.styles}"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Hostwerk Office</Application></Properties>');
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><workbookPr/><bookViews><workbookView/></bookViews><sheets/><calcPr calcId="191029"/></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL.styles}" Target="styles.xml"/></Relationships>`);
  return zip;
}

/**
 * Book → xlsx bytes. ext: xlsx | xlsm | xltx | xltm. With book.pkg (opened from an OOXML file) the original package
 * is patched; otherwise a fresh package is written.
 */
/** default value reader for chart data: "Sheet!$A$1:$B$5" → { values, texts } */
export const bookValues = (book) => (ref) => {
  const g = parseRange(ref); if (!g) return null;
  const sh = g.sheet ? book.sheetByName(g.sheet) : book.sheets[book.active]; if (!sh) return null;
  const values = []; const texts = [];
  for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) { const a = []; const tx = []; for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 200); c++) { a.push(book.value(sh, r, c)); tx.push(book.text(sh, r, c)); } values.push(a); texts.push(tx); }
  return { values, texts };
};

export async function writeXlsx(book, { ext = 'xlsx', getValues } = {}) {
  if (!getValues) getValues = bookValues(book);
  const pkg = book.pkg;
  const zip = pkg ? await JSZip.loadAsync(pkg.bytes) : skeleton(ext);
  const wbPath = pkg ? pkg.wbPath : 'xl/workbook.xml';
  let ct = await str(zip, '[Content_Types].xml');
  let wbXml = await str(zip, wbPath);
  let wbRelsXml = await str(zip, relsPathOf(wbPath));
  const wbRels = parseRels(wbRelsXml, wbPath);
  const stylesRel = wbRels.find((r) => r.type === REL.styles);
  const stylesPath = stylesRel ? stylesRel.path : 'xl/styles.xml';
  const defFont = (book.props.defaultFont) || { font: book.settings.font || 'Calibri', size: book.settings.fontSize || 11 };
  const sw = new StyleWriter(pkg && pkg.st, defFont);
  // shared strings
  const sstRaw = pkg ? pkg.sstRaw.slice() : [];
  const sstText = pkg ? pkg.sst : [];
  const sstIdx = new Map();
  sstRaw.forEach((x, i) => { if (!/<r>/.test(x) && !sstIdx.has(sstText[i])) sstIdx.set(sstText[i], i); });
  let sstCount = 0;
  const sIndex = (cell) => {
    sstCount++;
    if (cell._si != null && sstText[cell._si] === cell.v) return cell._si;
    let i = sstIdx.get(cell.v);
    if (i == null) { i = sstRaw.length; sstRaw.push(`<si><t${/^\s|\s$|\n/.test(cell.v) ? ' xml:space="preserve"' : ''}>${escText(cell.v)}</t></si>`); sstIdx.set(cell.v, i); }
    return i;
  };
  // macro handling for the target ext
  const macro = ext === 'xlsm' || ext === 'xltm';
  const vbaRel = wbRels.find((r) => r.type === REL.vba);
  if (!macro && vbaRel) { wbRelsXml = removeRels(wbRelsXml, (a) => a.Type === REL.vba); zip.remove(vbaRel.path); ct = ctRemove(ct, vbaRel.path); }
  ct = ctOverride(ct, wbPath, MAIN_CT[ext] || MAIN_CT.xlsx);
  // calc chain is dropped (Excel rebuilds it)
  const cc = wbRels.find((r) => r.type === REL.calcChain);
  if (cc) { zip.remove(cc.path); ct = ctRemove(ct, cc.path); wbRelsXml = removeRels(wbRelsXml, (a) => a.Type === REL.calcChain); }
  // dynamic array metadata
  let metaOk = !!(pkg && pkg.metaXml && /XLDAPR/.test(pkg.metaXml));
  const anySpill = book.sheets.some((s) => s.spillAnchors.size);
  if (!metaOk && anySpill && !wbRels.some((r) => r.type === REL.metadata)) {
    zip.file('xl/metadata.xml', META_XML);
    const rid = nextRid(wbRelsXml);
    wbRelsXml = addRel(wbRelsXml, rid, REL.metadata, 'metadata.xml');
    ct = ctOverride(ct, 'xl/metadata.xml', CT.metadata);
    metaOk = true;
  }
  const metaIdx = metaOk ? (pkg && pkg.metaXml ? xldaprIndex(pkg.metaXml) : 1) : 0;
  // sheets
  const sheetEntries = [];
  const keptOther = pkg ? pkg.sheetsList.filter((x) => x.keep) : [];
  let nextSheetNo = 1; zip.forEach((p) => { const m = /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(p); if (m) nextSheetNo = Math.max(nextSheetNo, +m[1] + 1); });
  let maxSheetId = 0;
  if (pkg) for (const m of pkg.info.values()) maxSheetId = Math.max(maxSheetId, +m.sheetId || 0);
  for (const x of keptOther) maxSheetId = Math.max(maxSheetId, +x.a.sheetId || 0);
  // deleted sheets
  if (pkg) for (const [id, m] of pkg.info) if (!book.sheets.some((s) => s.id === id)) { zip.remove(m.path); ct = ctRemove(ct, m.path); wbRelsXml = removeRels(wbRelsXml, (a) => a.Id === m.rid); }
  const allCharts = [];
  for (let si = 0; si < book.sheets.length; si++) {
    const sh = book.sheets[si];
    let m = pkg && pkg.info.get(sh.id);
    if (!m) {
      const path = `xl/worksheets/sheet${nextSheetNo++}.xml`;
      const rid = nextRid(wbRelsXml);
      wbRelsXml = addRel(wbRelsXml, rid, REL.sheet, relTarget(wbPath, path));
      ct = ctOverride(ct, path, CT.sheet);
      zip.file(path, BLANK_SHEET);
      m = { path, rid, sheetId: String(++maxSheetId), rels: [], notesSig: '', tables: [], fresh: true };
    }
    const res = await writeSheet(zip, sh, si, m, { book, sw, sIndex, metaIdx, ct });
    ct = res.ct;
    sheetEntries.push({ sh, m });
    allCharts.push({ sh, m });
  }
  // workbook.xml: sheets, names, views, calc
  const sheetsXml = `<sheets>${sheetEntries.map(({ sh, m }) => `<sheet name="${esc(sh.name)}" sheetId="${m.sheetId}"${sh.hidden ? ` state="${sh.hidden === 'veryHidden' ? 'veryHidden' : 'hidden'}"` : ''} r:id="${m.rid}"/>`).join('')}${keptOther.map((x) => wbXml.slice(x.tag.start, x.tag.end)).join('')}</sheets>`;
  const dn = [];
  for (const d of book.names.values()) {
    const scope = d.scope ? book.sheets.indexOf(d.scope) : -1;
    if (d.scope && scope < 0) continue;
    dn.push(`<definedName name="${esc(d.name)}"${scope >= 0 ? ` localSheetId="${scope}"` : ''}${d.hidden ? ' hidden="1"' : ''}${d.comment ? ` comment="${esc(d.comment)}"` : ''}>${escText(toOoxml(d.ref))}</definedName>`);
  }
  book.sheets.forEach((sh, i) => {
    const q = `'${sh.name.replace(/'/g, "''")}'`;
    if (sh.print.area) { const g = parseRange(sh.print.area); if (g) dn.push(`<definedName name="_xlnm.Print_Area" localSheetId="${i}">${escText(`${q}!$${colName(g.c1)}$${g.r1 + 1}:$${colName(g.c2)}$${g.r2 + 1}`)}</definedName>`); }
    if (sh.filter) { const g = sh.filter.range; dn.push(`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${escText(`${q}!$${colName(g.c1)}$${g.r1 + 1}:$${colName(g.c2)}$${g.r2 + 1}`)}</definedName>`); }
  });
  const wbParts = { sheets: sheetsXml, definedNames: dn.length ? `<definedNames>${dn.join('')}</definedNames>` : null };
  const calc = first(wbXml, 'calcPr');
  let calcTag = calc ? wbXml.slice(calc.start, calc.start + wbXml.slice(calc.start).indexOf('>') + 1) : '<calcPr calcId="191029"/>';
  if (calc && calc.inner) calcTag = wbXml.slice(calc.start, calc.end);
  calcTag = calcTag.replace(/\s(calcMode|fullCalcOnLoad)="[^"]*"/g, '').replace(/<((?:\w+:)?calcPr)\b/, `<$1${book.autoCalc ? '' : ' calcMode="manual"'} fullCalcOnLoad="1"`);
  wbParts.calcPr = calcTag;
  const bv = first(wbXml, 'bookViews');
  if (bv) {
    const v = first(bv.inner, 'workbookView');
    if (v) {
      let tag = bv.inner.slice(v.start, v.end);
      tag = tag.replace(/\sactiveTab="\d+"/, '').replace(/\sfirstSheet="(\d+)"/, (x, n) => (+n >= book.sheets.length ? '' : x)).replace(/<((?:\w+:)?workbookView)\b/, `<$1 activeTab="${book.active}"`);
      wbParts.bookViews = wbXml.slice(bv.start, bv.end).replace(bv.inner, bv.inner.slice(0, v.start) + tag + bv.inner.slice(v.end));
    }
  }
  wbXml = assemble(wbXml, wbParts, ['fileVersion', 'fileSharing', 'workbookPr', 'AlternateContent', 'revisionPtr', 'absPath', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst']);
  if (book.date1904 && !/date1904/.test(wbXml)) wbXml = wbXml.replace(/<workbookPr\b/, '<workbookPr date1904="1"');
  zip.file(wbPath, wbXml);
  // shared strings
  const sstRel = parseRels(wbRelsXml, wbPath).find((r) => r.type === REL.sst);
  let sstPath = sstRel ? sstRel.path : null;
  if (sstRaw.length) {
    if (!sstPath) { sstPath = 'xl/sharedStrings.xml'; wbRelsXml = addRel(wbRelsXml, nextRid(wbRelsXml), REL.sst, relTarget(wbPath, sstPath)); ct = ctOverride(ct, sstPath, CT.sst); }
    zip.file(sstPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${sstCount}" uniqueCount="${sstRaw.length}">${sstRaw.join('')}</sst>`);
  }
  // styles
  if (sw.changed || !pkg) zip.file(stylesPath, sw.xml(pkg && pkg.st.xml));
  if (!pkg) ct = ctOverride(ct, stylesPath, CT.styles);
  zip.file(relsPathOf(wbPath), wbRelsXml);
  zip.file('[Content_Types].xml', ct);
  // charts: new ones are added, edited/moved/deleted originals patched, unchanged originals untouched
  const list = []; const orig = [];
  for (const { sh } of allCharts) {
    const add = []; const data = [];
    for (const ch of sh.charts) {
      const d = getValues ? resolveChartData(ch, getValues) : { categories: [], series: [] };
      ch._xl = xlAnchor(sh, ch.anchor);
      if (ch._orig && pkg && pkg.origSheetIds.has(sh.id)) {
        const changed = ch._sig !== JSON.stringify(chartSig(ch));
        if (changed) orig.push({ chart: ch, data: d, rewrite: JSON.stringify({ ...chartSig(ch), anchor: null }) !== JSON.stringify({ ...JSON.parse(ch._sig), anchor: null }) });
      } else { add.push(ch); data.push(d); }
    }
    if (pkg) for (const o of (pkg.origCharts && pkg.origCharts.get(sh.id)) || []) if (!sh.charts.includes(o)) orig.push({ chart: o, deleted: true });
    if (add.length) list.push({ sheetName: sh.name, charts: add, data });
  }
  if (orig.length) await patchOriginalCharts(zip, orig);
  if (list.length) await writeChartsIntoXlsx(zip, list);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
function xldaprIndex(meta) {
  const types = elements(first(meta, 'metadataTypes')?.inner || '', 'metadataType').map((e) => attrs(e.a).name);
  const t = types.indexOf('XLDAPR');
  if (t < 0) return 0;
  const bks = elements(first(meta, 'cellMetadata')?.inner || '', 'bk');
  const i = bks.findIndex((b) => new RegExp(`<rc t="${t + 1}"`).test(b.inner));
  return i < 0 ? 0 : i + 1;
}
function xlAnchor(sh, a) {
  if (!a) return { from: { c: 0, r: 0, dx: 0, dy: 0 }, to: { c: 6, r: 15, dx: 0, dy: 0 } };
  const cw = (c) => (sh.hiddenCols.has(c) ? 0 : sh.colW.get(c) ?? sh.defW);
  const rh = (r) => (sh.hiddenRows.has(r) ? 0 : sh.rowH.get(r) ?? sh.defH);
  let c = a.c; let x = a.dx + a.w; while (x > cw(c) && c < MAXC - 1) { x -= cw(c); c++; }
  let r = a.r; let y = a.dy + a.h; while (y > rh(r) && r < 1048575) { y -= rh(r); r++; }
  return { from: { c: a.c, r: a.r, dx: a.dx, dy: a.dy }, to: { c, r, dx: Math.round(x), dy: Math.round(y) } };
}

async function writeSheet(zip, sh, si, m, { book, sw, sIndex, metaIdx, ct }) {
  let xml = (await str(zip, m.path)) || BLANK_SHEET;
  const relsPath = relsPathOf(m.path);
  let rels = await str(zip, relsPath);
  const parts = {};
  // sheetPr: tab colour + fitToPage, other content kept
  {
    const pr = first(xml, 'sheetPr');
    let inner = pr ? pr.inner : '';
    let a = pr ? pr.a : '';
    inner = inner.replace(/<tabColor\b[^>]*\/>/, '');
    if (sh.color) inner = `<tabColor rgb="${rgb(sh.color)}"/>${inner}`;
    const fit = sh.print.fit && sh.print.fit !== 'none';
    if (/<pageSetUpPr\b/.test(inner)) inner = inner.replace(/<pageSetUpPr\b([^>]*?)\/>/, (x, at) => `<pageSetUpPr${at.replace(/\s*fitToPage="[^"]*"/, '')}${fit ? ' fitToPage="1"' : ''}/>`);
    else if (fit) inner = inner.replace(/(<\/?outlinePr\b[^>]*>)?$/, (x) => `${x || ''}<pageSetUpPr fitToPage="1"/>`);
    if (sh.outlineRows.size || sh.outlineCols.size) { if (!/<outlinePr\b/.test(inner)) inner = inner.replace(/(<tabColor\b[^>]*\/>)?/, (x) => `${x || ''}<outlinePr summaryBelow="1" summaryRight="1"/>`); }
    // keep schema order inside sheetPr: tabColor, outlinePr, pageSetUpPr
    const tab = /<tabColor\b[^>]*\/>/.exec(inner); const ol = /<outlinePr\b[^>]*\/>/.exec(inner); const psu = /<pageSetUpPr\b[^>]*\/>/.exec(inner);
    inner = `${tab ? tab[0] : ''}${ol ? ol[0] : ''}${psu ? psu[0] : ''}`;
    parts.sheetPr = inner || a.trim() ? `<sheetPr${a}>${inner}</sheetPr>`.replace(/<sheetPr([^>]*)><\/sheetPr>/, '<sheetPr$1/>') : null;
  }
  // used range
  let r2 = -1; let c2 = -1;
  for (const k of [...sh.cells.keys(), ...sh.spill.keys()]) { const r = Math.floor(k / MAXC); const c = k % MAXC; if (r > r2) r2 = r; if (c > c2) c2 = c; }
  parts.dimension = `<dimension ref="${r2 < 0 ? 'A1' : `A1:${addr(r2, c2)}`}"/>`;
  // views (keep the original sheetView attributes we don't model)
  {
    const v = first(xml, 'sheetView');
    const a = v ? attrs(v.a) : {};
    a.workbookViewId = a.workbookViewId || '0';
    if (sh.gridlines === false) a.showGridLines = '0'; else delete a.showGridLines;
    if (sh.headings === false) a.showRowColHeaders = '0'; else delete a.showRowColHeaders;
    if (sh.zoom && sh.zoom !== 100) a.zoomScale = String(sh.zoom); else delete a.zoomScale;
    if (si === book.active) a.tabSelected = '1'; else delete a.tabSelected;
    let inner = '';
    const fr = sh.freeze || {};
    if (fr.rows || fr.cols) {
      const tl = addr(fr.rows || 0, fr.cols || 0);
      const pane = fr.rows && fr.cols ? 'bottomRight' : fr.rows ? 'bottomLeft' : 'topRight';
      inner += `<pane${fr.cols ? ` xSplit="${fr.cols}"` : ''}${fr.rows ? ` ySplit="${fr.rows}"` : ''} topLeftCell="${tl}" activePane="${pane}" state="frozen"/>`;
      const ac = sh.sel ? addr(sh.sel.r, sh.sel.c) : tl;
      inner += `<selection pane="${pane}" activeCell="${ac}" sqref="${ac}"/>`;
    } else if (sh.sel) { const ac = addr(sh.sel.r, sh.sel.c); inner += `<selection activeCell="${ac}" sqref="${ac}"/>`; }
    const order = ['tabSelected', 'showGridLines', 'showRowColHeaders', 'rightToLeft', 'zoomScale', 'workbookViewId'];
    const keys = Object.keys(a).sort((x, y) => (order.indexOf(x) < 0 ? 50 : order.indexOf(x)) - (order.indexOf(y) < 0 ? 50 : order.indexOf(y)));
    parts.sheetViews = `<sheetViews><sheetView ${keys.map((k) => `${k}="${esc(a[k])}"`).join(' ')}>${inner}</sheetView></sheetViews>`;
  }
  {
    const old = attrOf(first(xml, 'sheetFormatPr'));
    const maxOl = Math.max(0, ...sh.outlineRows.values()); const maxOc = Math.max(0, ...sh.outlineCols.values());
    parts.sheetFormatPr = `<sheetFormatPr${old.baseColWidth ? ` baseColWidth="${old.baseColWidth}"` : ''}${Math.abs(sh.defW - 64) > 1 || old.defaultColWidth ? ` defaultColWidth="${px2w(sh.defW)}"` : ''} defaultRowHeight="${px2pt(sh.defH)}"${sh.defH !== 20 ? ' customHeight="1"' : ''}${maxOl ? ` outlineLevelRow="${maxOl}"` : ''}${maxOc ? ` outlineLevelCol="${maxOc}"` : ''}/>`;
  }
  // cols
  {
    const idx = new Set([...sh.colW.keys(), ...sh.hiddenCols, ...sh.outlineCols.keys(), ...(sh._colXf ? sh._colXf.keys() : [])]);
    const list = [...idx].sort((a, b) => a - b);
    const out = [];
    let i = 0;
    while (i < list.length) {
      const c = list[i];
      const sig = (x) => `${sh.colW.get(x)}|${sh.hiddenCols.has(x)}|${sh.outlineCols.get(x)}|${sh._colXf && sh._colXf.get(x)}`;
      let j = i;
      while (j + 1 < list.length && list[j + 1] === list[j] + 1 && sig(list[j + 1]) === sig(c)) j++;
      const w = sh.colW.get(c);
      out.push(`<col min="${c + 1}" max="${list[j] + 1}" width="${px2w(w ?? sh.defW)}"${sh._colXf && sh._colXf.get(c) ? ` style="${sh._colXf.get(c)}"` : ''}${sh.hiddenCols.has(c) ? ' hidden="1"' : ''}${w != null ? ' customWidth="1"' : ''}${sh.outlineCols.get(c) ? ` outlineLevel="${sh.outlineCols.get(c)}"` : ''}/>`);
      i = j + 1;
    }
    parts.cols = out.length ? `<cols>${out.join('')}</cols>` : null;
  }
  // sheetData
  {
    const rows = new Map();
    const put = (k, item) => { const r = Math.floor(k / MAXC); let x = rows.get(r); if (!x) rows.set(r, (x = [])); x.push([k % MAXC, item]); };
    for (const [k, cell] of sh.cells) put(k, cell);
    for (const [k, sp] of sh.spill) if (!sh.cells.has(k)) put(k, { v: sp.v, spilled: true });
    for (const r of [...sh.rowH.keys(), ...sh.hiddenRows, ...sh.filterHidden, ...sh.outlineRows.keys(), ...(sh._rowXf ? sh._rowXf.keys() : [])]) if (!rows.has(r)) rows.set(r, []);
    const out = [];
    for (const r of [...rows.keys()].sort((a, b) => a - b)) {
      const cells = rows.get(r).sort((a, b) => a[0] - b[0]);
      const h = sh.rowH.get(r);
      const hidden = sh.hiddenRows.has(r) || sh.filterHidden.has(r);
      const ol = sh.outlineRows.get(r);
      const rx = sh._rowXf && sh._rowXf.get(r);
      let x = `<row r="${r + 1}"${rx ? ` s="${rx}" customFormat="1"` : ''}${h != null ? ` ht="${px2pt(h)}" customHeight="1"` : ''}${hidden ? ' hidden="1"' : ''}${ol ? ` outlineLevel="${ol}"` : ''}>`;
      for (const [c, cell] of cells) x += cellXml(sh, r, c, cell, { sw, sIndex, metaIdx });
      out.push(`${x}</row>`);
    }
    parts.sheetData = out.length ? `<sheetData>${out.join('')}</sheetData>` : '<sheetData/>';
  }
  // autofilter
  if (sh.filter) {
    const g = sh.filter.range;
    let inner = '';
    for (const [ci, crit] of Object.entries(sh.filter.cols || {}).sort((a, b) => a[0] - b[0])) {
      const id = +ci - g.c1;
      if (crit.raw) { inner += `<filterColumn colId="${id}">${crit.raw}</filterColumn>`; continue; }
      if (crit.values) { const blank = crit.values.includes(''); inner += `<filterColumn colId="${id}"><filters${blank ? ' blank="1"' : ''}>${crit.values.filter((v) => v !== '').map((v) => `<filter val="${esc(v)}"/>`).join('')}</filters></filterColumn>`; continue; }
      if (crit.cond) {
        const OP = { eq: 'equal', ne: 'notEqual', gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual' };
        const one = (op, v) => (op === 'contains' ? `<customFilter val="*${esc(v)}*"/>` : op === 'notContains' ? `<customFilter operator="notEqual" val="*${esc(v)}*"/>` : op === 'begins' ? `<customFilter val="${esc(v)}*"/>` : op === 'ends' ? `<customFilter val="*${esc(v)}"/>` : `<customFilter operator="${OP[op] || 'equal'}" val="${esc(v)}"/>`);
        inner += `<filterColumn colId="${id}"><customFilters${crit.cond.join === 'and' && crit.cond.op2 ? ' and="1"' : ''}>${one(crit.cond.op, crit.cond.v)}${crit.cond.op2 ? one(crit.cond.op2, crit.cond.v2) : ''}</customFilters></filterColumn>`;
      }
    }
    parts.autoFilter = `<autoFilter ref="${rangeA1(g)}"${inner ? `>${inner}</autoFilter>` : '/>'}`;
  } else parts.autoFilter = null;
  parts.mergeCells = sh.merges.length ? `<mergeCells count="${sh.merges.length}">${sh.merges.map((g) => `<mergeCell ref="${rangeA1(g)}"/>`).join('')}</mergeCells>` : null;
  // conditional formatting
  {
    const blocks = [];
    let prio = 1;
    for (const rule of sh.cf) {
      if (!rule.ranges.length) continue;
      const sq = rule.ranges.map(rangeA1).join(' ');
      if (rule._raw && rule._rawSig === JSON.stringify(sigOf(rule))) { blocks.push(`<conditionalFormatting sqref="${sq}">${rule._raw.replace(/priority="\d+"/, `priority="${prio++}"`)}</conditionalFormatting>`); continue; }
      const x = cfRuleXml(rule, prio++, sw);
      if (x) blocks.push(`<conditionalFormatting sqref="${sq}">${x}</conditionalFormatting>`);
    }
    parts.conditionalFormatting = blocks.length ? blocks : null;
  }
  // data validation
  if (sh.dv.length) {
    const TYPES = ['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'];
    const items = sh.dv.filter((d) => d.ranges.length && TYPES.includes(d.type)).map((d) => {
      const op = d.type === 'list' || d.type === 'custom' ? '' : d.op && d.op !== 'between' ? ` operator="${d.op}"` : '';
      const p = d.prompt; const e = d.error || {};
      return `<dataValidation type="${d.type}"${d.type !== 'list' && d.type !== 'custom' && e.style && e.style !== 'stop' ? ` errorStyle="${e.style}"` : d.type === 'list' && e.style && e.style !== 'stop' ? ` errorStyle="${e.style}"` : ''}${op}${d.allowBlank ? ' allowBlank="1"' : ''}${d.dropdown === false ? ' showDropDown="1"' : ''}${p ? ' showInputMessage="1"' : ''}${e.show !== false ? ' showErrorMessage="1"' : ''}${e.title ? ` errorTitle="${esc(e.title)}"` : ''}${e.text ? ` error="${esc(e.text)}"` : ''}${p && p.title ? ` promptTitle="${esc(p.title)}"` : ''}${p && p.text ? ` prompt="${esc(p.text)}"` : ''} sqref="${d.ranges.map(rangeA1).join(' ')}">${d.f1 ? `<formula1>${escText(toOoxml(d.f1))}</formula1>` : ''}${d.f2 && /between/i.test(d.op || 'between') ? `<formula2>${escText(toOoxml(d.f2))}</formula2>` : ''}</dataValidation>`;
    });
    parts.dataValidations = items.length ? `<dataValidations count="${items.length}">${items.join('')}</dataValidations>` : null;
  } else parts.dataValidations = null;
  // hyperlinks (relationships are rebuilt)
  {
    rels = removeRels(rels, (a) => a.Type === REL.hyperlink);
    const items = [];
    for (const [k, cell] of [...sh.cells].sort((a, b) => a[0] - b[0])) {
      if (!cell.link) continue;
      const ref = addr(Math.floor(k / MAXC), k % MAXC);
      if (cell.link.startsWith('#')) items.push(`<hyperlink ref="${ref}" location="${esc(cell.link.slice(1))}" display="${esc(String(cell.v ?? cell.link.slice(1)))}"/>`);
      else { const rid = nextRid(rels); rels = addRel(rels, rid, REL.hyperlink, cell.link, true); items.push(`<hyperlink ref="${ref}" r:id="${rid}"/>`); }
    }
    parts.hyperlinks = items.length ? `<hyperlinks>${items.join('')}</hyperlinks>` : null;
  }
  // print options / setup
  parts.printOptions = sh.print.gridlines || sh.print.headings ? `<printOptions${sh.print.headings ? ' headings="1"' : ''}${sh.print.gridlines ? ' gridLines="1"' : ''}/>` : null;
  {
    const ps = first(xml, 'pageSetup');
    const a = ps ? attrs(ps.a) : {};
    a.orientation = sh.print.orientation || 'portrait';
    if (sh.print.fit === 'width') { a.fitToWidth = '1'; a.fitToHeight = '0'; } else if (sh.print.fit === 'page') { a.fitToWidth = '1'; a.fitToHeight = '1'; } else { delete a.fitToWidth; delete a.fitToHeight; }
    if (!ps && !(sh.print.fit && sh.print.fit !== 'none') && a.orientation === 'portrait') parts.pageSetup = null;
    else parts.pageSetup = `<pageSetup ${Object.entries(a).map(([k, v]) => `${k}="${esc(v)}"`).join(' ')}/>`;
    if (!first(xml, 'pageMargins')) parts.pageMargins = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
  }
  // notes (only when they changed)
  const sig = notesSig(sh);
  if (sig !== m.notesSig) {
    const res = writeNotes(zip, sh, m, rels, ct, xml);
    rels = res.rels; ct = res.ct;
    if (res.legacy !== undefined) parts.legacyDrawing = res.legacy;
  }
  // tables: patch refs and header names
  for (const t of m.tables || []) {
    const st = (sh.tables || []).find((x) => x.path === t.path);
    const tx = await str(zip, t.path);
    if (!tx) continue;
    if (!st) continue;
    const g = st.range;
    let nx = tx.replace(/(<table\b[^>]*?\sref=")[^"]*"/, `$1${rangeA1(g)}"`).replace(/(<autoFilter\b[^>]*?\sref=")[^"]*"/, `$1${rangeA1(g)}"`);
    if (t.header) {
      const n = g.c2 - g.c1 + 1;
      const tcs = elements(nx, 'tableColumn');
      const names = [];
      for (let j = 0; j < n; j++) { const v = book.value(sh, g.r1, g.c1 + j); names.push(v == null || v === '' ? `Column${j + 1}` : String(typeof v === 'number' ? book.text(sh, g.r1, g.c1 + j) : v)); }
      if (tcs.length === n) { let off = 0; tcs.forEach((e, j) => { const old = nx.slice(e.start + off, e.end + off); const nw = old.replace(/\sname="[^"]*"/, ` name="${esc(names[j])}"`); nx = nx.slice(0, e.start + off) + nw + nx.slice(e.end + off); off += nw.length - old.length; }); } else {
        const cols = names.map((nm, j) => `<tableColumn id="${j + 1}" name="${esc(nm)}"/>`).join('');
        nx = nx.replace(/<tableColumns\b[\s\S]*?<\/tableColumns>/, `<tableColumns count="${n}">${cols}</tableColumns>`);
      }
    }
    if (nx !== tx) zip.file(t.path, nx);
  }
  xml = assemble(xml, parts, SHEET_ORDER);
  if (!/xmlns:r=/.test(xml.slice(0, 600))) xml = xml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS_R}"`);
  zip.file(m.path, xml);
  if (rels && /<Relationship\b/.test(rels)) zip.file(relsPath, rels); else if (zip.file(relsPath)) zip.remove(relsPath);
  return { ct };
}

function cellXml(sh, r, c, cell, { sw, sIndex, metaIdx }) {
  const ref = addr(r, c);
  const s = sw.xf(cell.s);
  const sa = s ? ` s="${s}"` : '';
  let v = cell.v;
  if (cell.f) {
    const sp = sh.spillAnchors.get(r * MAXC + c);
    const arrRef = sp ? rangeA1({ r1: r, c1: c, r2: r + sp.rows - 1, c2: c + sp.cols - 1 }) : null;
    const f = `<f${arrRef ? ` t="array" ref="${arrRef}"` : ''}>${escText(toOoxml(cell.f))}</f>`;
    const cm = arrRef && metaIdx ? ` cm="${metaIdx}"` : '';
    if (isErr(v)) return v.code === '#CIRC!' ? `<c r="${ref}"${sa}${cm}>${f}<v>0</v></c>` : `<c r="${ref}"${sa} t="e"${cm}>${f}<v>${esc(v.code)}</v></c>`;
    if (typeof v === 'string') return `<c r="${ref}"${sa} t="str"${cm}>${f}<v>${escText(v)}</v></c>`;
    if (typeof v === 'boolean') return `<c r="${ref}"${sa} t="b"${cm}>${f}<v>${v ? 1 : 0}</v></c>`;
    if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}"${sa}${cm}>${f}<v>${v}</v></c>`;
    return `<c r="${ref}"${sa}${cm}>${f}</c>`;
  }
  if (v == null || v === '') return s ? `<c r="${ref}"${sa}/>` : '';
  if (typeof v === 'number') return Number.isFinite(v) ? `<c r="${ref}"${sa}><v>${v}</v></c>` : `<c r="${ref}"${sa} t="e"><v>#NUM!</v></c>`;
  if (typeof v === 'boolean') return `<c r="${ref}"${sa} t="b"><v>${v ? 1 : 0}</v></c>`;
  if (isErr(v)) return `<c r="${ref}"${sa} t="e"><v>${esc(v.code === '#CIRC!' ? '#VALUE!' : v.code)}</v></c>`;
  return `<c r="${ref}"${sa} t="s"><v>${sIndex(cell.spilled ? { v: String(v) } : cell)}</v></c>`;
}

function cfRuleXml(rule, prio, sw) {
  const dxf = rule.style && Object.keys(rule.style).length ? ` dxfId="${sw.dxf(rule.style)}"` : '';
  const p = ` priority="${prio}"${rule.stop ? ' stopIfTrue="1"' : ''}`;
  const g = rule.ranges[0];
  const tl = addr(g.r1, g.c1);
  const F = (f) => `<formula>${escText(toOoxml(String(f)))}</formula>`;
  const qt = (t) => String(t).replace(/"/g, '""');
  switch (rule.type) {
    case 'cell': {
      const OP = { gt: 'greaterThan', lt: 'lessThan', between: 'between', notBetween: 'notBetween', eq: 'equal', ne: 'notEqual', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual' };
      return `<cfRule type="cellIs"${dxf}${p} operator="${OP[rule.op] || 'equal'}">${F(rule.f1 || '0')}${/between/i.test(rule.op) ? F(rule.f2 || '0') : ''}</cfRule>`;
    }
    case 'text': {
      const t = rule.text || '';
      const m = { contains: ['containsText', 'containsText', `NOT(ISERROR(SEARCH("${qt(t)}",${tl})))`], notContains: ['notContainsText', 'notContains', `ISERROR(SEARCH("${qt(t)}",${tl}))`], begins: ['beginsWith', 'beginsWith', `LEFT(${tl},LEN("${qt(t)}"))="${qt(t)}"`], ends: ['endsWith', 'endsWith', `RIGHT(${tl},LEN("${qt(t)}"))="${qt(t)}"`] }[rule.op || 'contains'];
      return `<cfRule type="${m[0]}"${dxf}${p} operator="${m[1]}" text="${esc(t)}">${F(m[2])}</cfRule>`;
    }
    case 'dup': return `<cfRule type="duplicateValues"${dxf}${p}/>`;
    case 'unique': return `<cfRule type="uniqueValues"${dxf}${p}/>`;
    case 'top': case 'bottom': return `<cfRule type="top10"${dxf}${p}${rule.percent ? ' percent="1"' : ''}${rule.type === 'bottom' ? ' bottom="1"' : ''} rank="${rule.rank || 10}"/>`;
    case 'above': return `<cfRule type="aboveAverage"${dxf}${p}/>`;
    case 'below': return `<cfRule type="aboveAverage"${dxf}${p} aboveAverage="0"/>`;
    case 'blank': return `<cfRule type="containsBlanks"${dxf}${p}>${F(`LEN(TRIM(${tl}))=0`)}</cfRule>`;
    case 'notBlank': return `<cfRule type="notContainsBlanks"${dxf}${p}>${F(`LEN(TRIM(${tl}))>0`)}</cfRule>`;
    case 'error': return `<cfRule type="containsErrors"${dxf}${p}>${F(`ISERROR(${tl})`)}</cfRule>`;
    case 'formula': return `<cfRule type="expression"${dxf}${p}>${F(rule.formula || 'FALSE')}</cfRule>`;
    case 'scale': {
      const cols = rule.colors && rule.colors.length ? rule.colors : ['#F8696B', '#FFEB84', '#63BE7B'];
      const stops = rule.stops && rule.stops.length === cols.length ? rule.stops : cols.length === 3 ? [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }];
      return `<cfRule type="colorScale"${p}><colorScale>${stops.map((s) => `<cfvo type="${s.type}"${s.val != null ? ` val="${esc(s.val)}"` : ''}/>`).join('')}${cols.map((c) => `<color rgb="${rgb(c)}"/>`).join('')}</colorScale></cfRule>`;
    }
    case 'bar': {
      const stops = rule.stops && rule.stops.length === 2 ? rule.stops : [{ type: 'min' }, { type: 'max' }];
      return `<cfRule type="dataBar"${p}><dataBar>${stops.map((s) => `<cfvo type="${s.type}"${s.val != null ? ` val="${esc(s.val)}"` : ''}/>`).join('')}<color rgb="${rgb(rule.color || '#638EC6')}"/></dataBar></cfRule>`;
    }
    default: return rule._raw || null;
  }
}

function writeNotes(zip, sh, m, rels, ct, xml) {
  const notes = [...sh.cells].filter(([, c]) => c.note).sort((a, b) => a[0] - b[0]);
  const oldC = (m.rels || []).find((r) => r.type === REL.comments);
  const oldV = (m.rels || []).find((r) => r.type === REL.vml);
  if (!notes.length) {
    if (oldC) { zip.remove(oldC.path); ct = ctRemove(ct, oldC.path); rels = removeRels(rels, (a) => a.Id === oldC.id); }
    if (oldV && oldC) { zip.remove(oldV.path); rels = removeRels(rels, (a) => a.Id === oldV.id); return { rels, ct, legacy: null }; }
    return { rels, ct };
  }
  const authors = [...new Set(notes.map(([, c]) => c.noteAuthor || ''))];
  const cXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${NS_MAIN}"><authors>${authors.map((a) => `<author>${escText(a)}</author>`).join('')}</authors><commentList>${notes.map(([k, c]) => `<comment ref="${addr(Math.floor(k / MAXC), k % MAXC)}" authorId="${authors.indexOf(c.noteAuthor || '')}"><text><r><t xml:space="preserve">${escText(c.note)}</t></r></text></comment>`).join('')}</commentList></comments>`;
  const shapes = notes.map(([k], i) => {
    const r = Math.floor(k / MAXC); const c = k % MAXC;
    return `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${r}, 2, ${c + 3}, 15, ${r + 4}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
  }).join('');
  const vXml = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes}</xml>`;
  let n = 1; zip.forEach((p) => { const mm = /^xl\/comments(\d+)\.xml$/.exec(p); if (mm) n = Math.max(n, +mm[1] + 1); });
  const cPath = oldC ? oldC.path : `xl/comments${n}.xml`;
  let vn = 1; zip.forEach((p) => { const mm = /^xl\/drawings\/vmlDrawing(\d+)\.vml$/.exec(p); if (mm) vn = Math.max(vn, +mm[1] + 1); });
  const vPath = oldV ? oldV.path : `xl/drawings/vmlDrawing${vn}.vml`;
  zip.file(cPath, cXml);
  zip.file(vPath, vXml);
  ct = ctOverride(ct, cPath, CT.comments);
  ct = ctDefault(ct, 'vml', CT.vml);
  let legacy;
  if (!oldC) { const rid = nextRid(rels); rels = addRel(rels, rid, REL.comments, relTarget(m.path, cPath)); }
  if (!oldV) { const rid = nextRid(rels); rels = addRel(rels, rid, REL.vml, relTarget(m.path, vPath)); legacy = `<legacyDrawing r:id="${rid}"/>`; } else if (!/<legacyDrawing\b/.test(xml)) legacy = `<legacyDrawing r:id="${oldV.id}"/>`;
  return { rels, ct, legacy };
}

/** list of the VBA modules for the read-only "Macros" panel (needs the CFB reader – see io/vba.js) */
export async function vbaProjectBytes(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const f = Object.keys(zip.files).find((p) => /vbaProject\.bin$/i.test(p));
  return f ? zip.file(f).async('uint8array') : null;
}
