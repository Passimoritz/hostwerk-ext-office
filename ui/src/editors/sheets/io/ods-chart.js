/*
 * Charts of OpenDocument spreadsheets: an embedded chart object ("Object 1/content.xml", or inline in a flat file)
 * ↔ the editor's chart model (type, stacking, title, legend, axis titles, series with name/category/value ranges,
 * colours, data labels, trendlines). Cell ranges use ODF addresses ("Sheet1.$A$2:.$A$5").
 */
import { kids, kid, all, text as textOfEl } from '../../_shared/xmltree.js';
import { parseRangeRef } from '../chartData.js';
import { xesc, cm } from './ods-styles.js';

const colName = (c) => { let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
const odfSheet = (n) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${String(n).replace(/'/g, "''")}'`);

/** "'My Sheet'!$A$1:$B$4" → "'My Sheet'.$A$1:'My Sheet'.$B$4" (defaultSheet when the ref has none) */
export function excelToOdfRange(ref, defaultSheet) {
  const g = parseRangeRef(ref); if (!g) return null;
  const s = odfSheet(g.sheet || defaultSheet);
  const a = `${s}.$${colName(g.c1)}$${g.r1 + 1}`;
  return g.r1 === g.r2 && g.c1 === g.c2 ? a : `${a}:.$${colName(g.c2)}$${g.r2 + 1}`; // the form Excel and LibreOffice write
}
/** ODF cell (range) address → Excel ref ("Sheet1.A1:Sheet1.B5", "$Sheet1.$A$1:.$B$5", "'a b'.A1") */
export function odfToExcelRange(addr) {
  if (!addr) return null;
  const first = String(addr).trim().split(/\s+/)[0];
  const m = /^\$?(?:'((?:[^']|'')+)'|([^.':]+))?\.(\$?[A-Z]+\$?\d+)(?::\$?(?:'(?:[^']|'')+'|[^.':]+)?\.(\$?[A-Z]+\$?\d+))?$/i.exec(first);
  if (!m) return null;
  const sheet = m[1] != null ? m[1].replace(/''/g, "'") : m[2] || null;
  const abs = (x) => x.replace(/^\$?([A-Z]+)\$?(\d+)$/i, '$$$1$$$2').toUpperCase();
  const q = sheet ? (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) && !/^[A-Za-z]{1,3}\d+$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`) : null;
  return `${q ? `${q}!` : ''}${abs(m[3])}${m[4] && abs(m[4]) !== abs(m[3]) ? `:${abs(m[4])}` : ''}`;
}

const CLASS = { column: 'chart:bar', bar: 'chart:bar', line: 'chart:line', area: 'chart:area', pie: 'chart:circle', doughnut: 'chart:ring', scatter: 'chart:scatter' };
const LEG = { right: 'end', left: 'start', top: 'top', bottom: 'bottom' };
const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';

/**
 * chart model + resolved data ({ categories, series: [{ name, values, color, refs }] }) → the chart document's inner
 * parts: { automatic, body } (office:automatic-styles content and office:body content) and `ranges` (notify list).
 */
export function chartParts(ch, data, sheetName, size) {
  const type = ch.type || 'column';
  const cls = CLASS[type] || 'chart:bar';
  const pie = type === 'pie' || type === 'doughnut';
  const S = data.series || [];
  const auto = [];
  auto.push('<style:style style:name="ch" style:family="chart"><style:graphic-properties draw:stroke="none" draw:fill="solid" draw:fill-color="#ffffff"/></style:style>');
  auto.push('<style:style style:name="ti" style:family="chart"><style:text-properties fo:font-size="13pt" fo:font-weight="bold"/></style:style>');
  auto.push(`<style:style style:name="pa" style:family="chart"><style:chart-properties${type === 'bar' ? ' chart:vertical="true"' : ''}${ch.stacked === 'stacked' ? ' chart:stacked="true"' : ch.stacked === 'percent' ? ' chart:percentage="true"' : ''}${type === 'line' || type === 'scatter' ? ` chart:symbol-type="${type === 'scatter' ? 'automatic' : 'none'}"` : ''}${type === 'scatter' ? ' chart:lines="false"' : ''} chart:include-hidden-cells="false"/></style:style>`);
  auto.push('<style:style style:name="ax" style:family="chart"><style:chart-properties chart:display-label="true"/><style:graphic-properties svg:stroke-color="#b3b3b3"/></style:style>');
  const ranges = [];
  const series = S.map((s, i) => {
    const col = String(s.color || '#4472c4').toLowerCase();
    const src = (ch.series && ch.series[i]) || {};
    const labels = ch.labels || src.labels;
    auto.push(`<style:style style:name="se${i}" style:family="chart"><style:chart-properties${labels ? ' chart:data-label-number="value"' : ''}${pie ? ' chart:data-label-number="percentage"' : ''}/><style:graphic-properties draw:fill="solid" draw:fill-color="${col}" svg:stroke-color="${col}"${type === 'line' || type === 'scatter' ? ' svg:stroke-width="0.079cm"' : ''}/></style:style>`);
    const tr = ch.trend || src.trendline;
    if (tr && !pie) auto.push(`<style:style style:name="tr${i}" style:family="chart"><style:chart-properties chart:regression-type="${tr === 'exp' ? 'exponential' : tr === 'log' ? 'logarithmic' : tr === 'poly' ? 'polynomial' : tr === 'power' ? 'power' : 'linear'}"/><style:graphic-properties svg:stroke-color="${col}" draw:stroke="dash"/></style:style>`);
    const val = s.refs && s.refs.val ? excelToOdfRange(s.refs.val, sheetName) : null;
    const lab = s.refs && s.refs.name ? excelToOdfRange(s.refs.name, sheetName) : null;
    if (val) ranges.push(val);
    const xv = type === 'scatter' && s.refs && s.refs.cat ? excelToOdfRange(s.refs.cat, sheetName) : null;
    const n = Math.max(1, (s.values || []).length);
    return `<chart:series chart:style-name="se${i}"${val ? ` chart:values-cell-range-address="${xesc(val)}"` : ''}${lab ? ` chart:label-cell-address="${xesc(lab)}"` : ''} chart:class="${cls}">${xv ? `<chart:domain table:cell-range-address="${xesc(xv)}"/>` : ''}${tr && !pie ? `<chart:regression-curve chart:style-name="tr${i}"/>` : ''}<chart:data-point chart:repeated="${n}"/></chart:series>`;
  });
  const cat0 = S.find((s) => s.refs && s.refs.cat);
  const catRange = cat0 ? excelToOdfRange(cat0.refs.cat, sheetName) : null;
  if (catRange) ranges.unshift(catRange);
  const title = ch.title ? `<chart:title chart:style-name="ti"><text:p>${xesc(ch.title)}</text:p></chart:title>` : '';
  const legend = ch.legend && ch.legend !== 'none' ? `<chart:legend chart:legend-position="${LEG[ch.legend] || 'end'}" style:legend-expansion="high"/>` : '';
  const axTitle = (t) => (t ? `<chart:title><text:p>${xesc(t)}</text:p></chart:title>` : '');
  const axes = pie ? '' : `<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ax">${axTitle(ch.xTitle)}${catRange && type !== 'scatter' ? `<chart:categories table:cell-range-address="${xesc(catRange)}"/>` : ''}</chart:axis><chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ax">${axTitle(ch.yTitle)}<chart:grid chart:class="major"/></chart:axis>`;
  // local table with the cached values (readers without access to the sheet show these)
  const cats = data.categories || [];
  const rows = Math.max(cats.length, ...S.map((s) => (s.values || []).length), 0);
  const cell = (v) => (typeof v === 'number' && Number.isFinite(v) ? `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>` : `<table:table-cell office:value-type="string"><text:p>${xesc(v ?? '')}</text:p></table:table-cell>`);
  let local = `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns><table:table-column table:number-columns-repeated="${Math.max(1, S.length)}"/></table:table-columns>`;
  local += `<table:table-header-rows><table:table-row><table:table-cell/>${S.map((s) => cell(s.name)).join('')}</table:table-row></table:table-header-rows><table:table-rows>`;
  for (let r = 0; r < rows; r++) local += `<table:table-row>${cell(cats[r] ?? String(r + 1))}${S.map((s) => cell((s.values || [])[r] ?? null)).join('')}</table:table-row>`;
  local += '</table:table-rows></table:table>';
  const plotRange = ranges.length ? ranges.join(' ') : '';
  const body = `<office:chart><chart:chart svg:width="${cm(size.w)}" svg:height="${cm(size.h)}" chart:class="${cls}" chart:style-name="ch">${title}${legend}<chart:plot-area chart:style-name="pa"${plotRange ? ` table:cell-range-address="${xesc(plotRange)}"` : ''} chart:data-source-has-labels="${cat0 ? 'both' : 'row'}">${axes}${series.join('')}</chart:plot-area>${ranges.length ? '' : local}</chart:chart></office:chart>`;
  return { automatic: auto.join(''), body, ranges };
}
/** stand-alone chart content.xml (packaged files) */
export function chartContentXml(parts) {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS} office:version="1.2"><office:automatic-styles>${parts.automatic}</office:automatic-styles><office:body>${parts.body}</office:body></office:document-content>`;
}
/** inline chart document of a flat file (inside draw:object) */
export function chartInlineXml(parts) {
  return `<office:document ${NS} office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.chart"><office:automatic-styles>${parts.automatic}</office:automatic-styles><office:body>${parts.body}</office:body></office:document>`;
}

/* ---------------- reading ---------------- */
const KIND = { 'chart:bar': 'column', 'chart:line': 'line', 'chart:area': 'area', 'chart:circle': 'pie', 'chart:ring': 'doughnut', 'chart:scatter': 'scatter' };
const LEG_R = { end: 'right', start: 'left', top: 'top', bottom: 'bottom', 'top-end': 'right', 'bottom-end': 'right', 'top-start': 'left', 'bottom-start': 'left' };
const cleanCls = (c) => (c ? (c.includes(':') ? `chart:${c.split(':').pop()}` : `chart:${c}`) : null);
/** parsed chart document root (office:document-content / office:document) → chart model (or null) */
export function parseOdfChart(root) {
  const chartEl = all(root, 'chart:chart')[0]; if (!chartEl) return null;
  const styles = new Map();
  for (const st of all(root, 'style:style')) styles.set(st.a['style:name'], st);
  const prop = (name, tag) => { const st = styles.get(name); const p = st && kid(st, tag); return p ? p.a : {}; };
  const cls = cleanCls(chartEl.a['chart:class']);
  let type = KIND[cls] || null;
  const ch = { type: type || 'column', stacked: false, title: '', legend: 'none', xTitle: '', yTitle: '', range: null, byRows: false, series: [], colors: null };
  if (!type) ch._unsupported = true;
  const title = kid(chartEl, 'chart:title'); if (title) ch.title = kids(title, 'text:p').map(textOfEl).join(' ');
  const leg = kid(chartEl, 'chart:legend'); if (leg) ch.legend = LEG_R[leg.a['chart:legend-position'] || 'end'] || 'right';
  const plot = kid(chartEl, 'chart:plot-area'); if (!plot) return null;
  const pp = prop(plot.a['chart:style-name'], 'style:chart-properties');
  if (pp['chart:vertical'] === 'true' && ch.type === 'column') ch.type = 'bar';
  if (pp['chart:percentage'] === 'true') ch.stacked = 'percent'; else if (pp['chart:stacked'] === 'true') ch.stacked = 'stacked';
  let cats = null;
  for (const ax of kids(plot, 'chart:axis')) {
    const t = kid(ax, 'chart:title'); const tt = t ? kids(t, 'text:p').map(textOfEl).join(' ') : '';
    if (ax.a['chart:dimension'] === 'x') { if (tt) ch.xTitle = tt; const c = kid(ax, 'chart:categories'); if (c) cats = odfToExcelRange(c.a['table:cell-range-address']); } else if (ax.a['chart:dimension'] === 'y' && tt && !ch.yTitle) ch.yTitle = tt;
  }
  const sers = kids(plot, 'chart:series');
  sers.forEach((s) => {
    const sc = cleanCls(s.a['chart:class']);
    if (sc && KIND[sc] && KIND[sc] !== ch.type && !(KIND[sc] === 'column' && ch.type === 'bar')) ch._partial = true;
    const gp = prop(s.a['chart:style-name'], 'style:graphic-properties');
    const cp = prop(s.a['chart:style-name'], 'style:chart-properties');
    const dom = kid(s, 'chart:domain');
    const ser = { name: '', nameRef: odfToExcelRange(s.a['chart:label-cell-address']), catRef: dom ? odfToExcelRange(dom.a['table:cell-range-address']) : cats, valRef: odfToExcelRange(s.a['chart:values-cell-range-address']), color: null };
    const col = ch.type === 'line' || ch.type === 'scatter' ? gp['svg:stroke-color'] || gp['draw:fill-color'] : gp['draw:fill-color'];
    if (col && !/^(pie|doughnut)$/.test(ch.type)) ser.color = col.toLowerCase();
    if (cp['chart:data-label-number'] === 'value') ser.labels = true;
    const rc = kid(s, 'chart:regression-curve');
    if (rc) { const rt = prop(rc.a['chart:style-name'], 'style:chart-properties')['chart:regression-type']; if (rt && rt !== 'none') ser.trendline = rt === 'exponential' ? 'exp' : rt === 'logarithmic' ? 'log' : rt === 'polynomial' ? 'poly' : rt === 'power' ? 'power' : 'linear'; }
    if (!ser.valRef) { // values only in the local table: keep them as a literal-less series (cached in the model)
      ser._local = true;
    }
    ch.series.push(ser);
  });
  // series without ranges: take the plot area's range (first column = categories) when there is one
  if (ch.series.some((s) => !s.valRef) && plot.a['table:cell-range-address']) {
    const r = odfToExcelRange(plot.a['table:cell-range-address']);
    if (r) { ch.series = []; ch.range = r; }
  }
  ch.series = ch.series.filter((s) => s.valRef || !s._local);
  if (!ch.series.length && !ch.range) return null;
  return ch;
}
