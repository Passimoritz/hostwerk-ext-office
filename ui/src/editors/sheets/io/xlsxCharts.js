/*
 * Charts in xlsx files (DrawingML). exceljs neither reads nor writes charts, so this module works directly on the
 * zip (an injected JSZip instance): readChartsFromXlsx parses the basic chart kinds into the editor's chart model,
 * writeChartsIntoXlsx adds chart parts + drawing anchors to a zip that exceljs wrote. String parsing only (runs in Node).
 */
import { PALETTE } from '../chartData.js';

const NS = {
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
};
const T_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const T_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const EMU = 9525;

/* ---------------- small XML helpers ---------------- */
export const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => String(s ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d)).replace(/&amp;/g, '&');
/** drop namespace prefixes from element names (attributes keep theirs) */
const strip = (xml) => xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, '<$1');
const attrs = (tag) => { const o = {}; tag.replace(/([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g, (_, k, __, a, b) => { o[k] = unesc(a ?? b); }); return o; };
/** first <tag …>…</tag> (or self-closing) → { attrs, body } */
function el(xml, tag) {
  const re = new RegExp(`<${tag}(\\s[^>]*?)?(/?)>`, 'g');
  const m = re.exec(xml);
  if (!m) return null;
  if (m[2] === '/') return { attrs: attrs(m[1] || ''), body: '', start: m.index, end: re.lastIndex };
  const close = `</${tag}>`; let depth = 1; let i = re.lastIndex;
  const open = new RegExp(`<${tag}(\\s[^>]*?)?(/?)>`, 'g');
  while (depth > 0) {
    const ci = xml.indexOf(close, i); if (ci < 0) return { attrs: attrs(m[1] || ''), body: xml.slice(re.lastIndex), start: m.index, end: xml.length };
    open.lastIndex = i; let om;
    while ((om = open.exec(xml)) && om.index < ci) { if (om[2] !== '/') depth++; }
    depth--; i = ci + close.length;
    if (depth === 0) return { attrs: attrs(m[1] || ''), body: xml.slice(re.lastIndex, ci), start: m.index, end: i };
  }
  return null;
}
function els(xml, tag) { const out = []; let rest = xml; for (;;) { const e = el(rest, tag); if (!e) break; out.push(e); rest = rest.slice(e.end); } return out; }
const val = (xml, tag) => { const e = el(xml, tag); return e ? e.attrs.val : undefined; };

const dirOf = (p) => p.slice(0, p.lastIndexOf('/') + 1);
function resolve(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(base) + target).split('/'); const out = [];
  for (const x of parts) { if (x === '..') out.pop(); else if (x !== '.') out.push(x); }
  return out.join('/');
}
const relsPath = (p) => `${dirOf(p)}_rels/${p.slice(p.lastIndexOf('/') + 1)}.rels`;
async function text(zip, p) { const f = zip.file(p); return f ? f.async('string') : null; }
async function rels(zip, part) {
  const x = await text(zip, relsPath(part)); const out = [];
  if (x) for (const m of x.match(/<Relationship\b[^>]*>/g) || []) { const a = attrs(m); out.push({ id: a.Id, type: a.Type, target: a.Target, mode: a.TargetMode, path: a.TargetMode === 'External' ? null : resolve(part, a.Target) }); }
  return out;
}
/** [{ name, path }] of the workbook's worksheets in order */
async function sheetParts(zip) {
  const wbPath = 'xl/workbook.xml';
  const wb = await text(zip, wbPath); if (!wb) return [];
  const rs = await rels(zip, wbPath);
  return (wb.match(/<(\w+:)?sheet\b[^>]*>/g) || []).map((t) => { const a = attrs(t); const rel = rs.find((r) => r.id === a['r:id'] || r.id === Object.entries(a).find(([k]) => /:id$/.test(k))?.[1]); return { name: a.name, path: rel && rel.path }; }).filter((s) => s.path);
}

/* ---------------- reading ---------------- */
const KIND = { barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter',
  radarChart: null, stockChart: null, surfaceChart: null, surface3DChart: null, bubbleChart: null };

function richText(xml) {
  if (!xml) return '';
  const ts = (xml.match(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g) || []).map((t) => unesc(t.replace(/<[^>]+>/g, '')));
  if (ts.length) return ts.join('');
  const vs = (xml.match(/<v>([\s\S]*?)<\/v>/g) || []).map((t) => unesc(t.replace(/<[^>]+>/g, '')));
  return vs.join(' ');
}
const refOf = (xml) => { if (!xml) return null; const f = el(xml, 'f'); return f ? unesc(f.body).trim() : null; };

function parseChart(raw) {
  const x = strip(raw);
  const chartEl = el(x, 'chart'); if (!chartEl) return null;
  const plot = el(chartEl.body, 'plotArea'); if (!plot) return null;
  const outside = chartEl.body.slice(0, plot.start) + chartEl.body.slice(plot.end);
  const groups = [];
  for (const k of Object.keys(KIND)) for (const g of els(plot.body, k)) groups.push({ k, g, at: plot.body.indexOf(`<${k}`) });
  groups.sort((a, b) => a.at - b.at);
  if (!groups.length) return null;
  const { k, g } = groups[0];
  const chart = { type: KIND[k] || 'column', stacked: false, title: '', legend: 'none', xTitle: '', yTitle: '', range: null, byRows: false, series: [], colors: null };
  if (KIND[k] == null) chart._unsupported = true;
  if (groups.length > 1) chart._partial = true;
  if (k.startsWith('bar')) chart.type = val(g.body, 'barDir') === 'bar' ? 'bar' : 'column';
  const grouping = val(g.body, 'grouping');
  if (grouping === 'stacked') chart.stacked = 'stacked'; else if (grouping === 'percentStacked') chart.stacked = 'percent';
  const title = el(outside, 'title');
  if (title) chart.title = richText(title.body);
  { const gl = g.body.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, ''); const dl = el(gl, 'dLbls'); if (dl && /<(?:c:)?showVal val="1"/.test(dl.body)) chart.labels = true; }
  const leg = el(outside, 'legend');
  if (leg) chart.legend = ({ r: 'right', b: 'bottom', t: 'top', l: 'left', tr: 'right' })[val(leg.body, 'legendPos') || 'r'] || 'right';
  // axis titles
  const axes = [...els(plot.body, 'catAx').map((a) => ({ a, cat: true })), ...els(plot.body, 'dateAx').map((a) => ({ a, cat: true })), ...els(plot.body, 'valAx').map((a) => ({ a, cat: false }))];
  for (const { a, cat } of axes) {
    const t = el(a.body, 'title'); if (!t) continue;
    const pos = val(a.body, 'axPos');
    const isX = chart.type === 'scatter' ? pos === 'b' || pos === 't' : cat;
    if (isX) chart.xTitle = richText(t.body); else chart.yTitle = richText(t.body);
  }
  for (const s of els(g.body, 'ser')) {
    const own = s.body.replace(/<dPt>[\s\S]*?<\/dPt>/g, '').replace(/<dLbls>[\s\S]*?<\/dLbls>/g, '').replace(/<marker>[\s\S]*?<\/marker>/g, '').replace(/<trendline>[\s\S]*?<\/trendline>/g, '');
    const tx = el(own, 'tx'); const cat = el(own, 'cat') || el(own, 'xVal'); const v = el(own, 'val') || el(own, 'yVal');
    const ser = { name: '', nameRef: null, catRef: refOf(cat && cat.body), valRef: refOf(v && v.body), color: null };
    const dl = el(s.body, 'dLbls'); if (dl && /<(?:c:)?showVal val="1"/.test(dl.body)) ser.labels = true;
    const tl = el(s.body, 'trendline'); if (tl) ser.trendline = val(tl.body, 'trendlineType') || 'linear';
    if (tx) { ser.nameRef = refOf(tx.body); ser.name = richText(tx.body); }
    const sp = el(own, 'spPr');
    if (sp && chart.type !== 'pie' && chart.type !== 'doughnut') {
      const fillSrc = chart.type === 'line' || chart.type === 'scatter' ? (el(sp.body, 'ln') || sp).body : sp.body.replace(/<ln[\s>][\s\S]*?<\/ln>/g, '');
      const m = /<solidFill>\s*<srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(fillSrc);
      if (m) ser.color = `#${m[1].toLowerCase()}`;
    }
    chart.series.push(ser);
  }
  return chart;
}

function anchorOf(a, isTwo) {
  const pt = (b) => b && { c: +(el(b.body, 'col')?.body || 0), r: +(el(b.body, 'row')?.body || 0), dx: Math.round(+(el(b.body, 'colOff')?.body || 0) / EMU), dy: Math.round(+(el(b.body, 'rowOff')?.body || 0) / EMU) };
  const from = pt(el(a, 'from'));
  if (isTwo) return { from, to: pt(el(a, 'to')) };
  const ext = el(a, 'ext');
  return { from, ext: { w: Math.round(+(ext?.attrs.cx || 0) / EMU), h: Math.round(+(ext?.attrs.cy || 0) / EMU) } };
}

/** → Map<sheetName, chart[]> (anchor: { from: { c, r, dx, dy }, to } or { from, ext: { w, h } }, offsets in px) */
export async function readChartsFromXlsx(zip) {
  const out = new Map();
  let n = 0;
  for (const sh of await sheetParts(zip)) {
    const srels = await rels(zip, sh.path);
    for (const dr of srels.filter((r) => r.type === T_DRAWING && r.path)) {
      const dx = await text(zip, dr.path); if (!dx) continue;
      const drels = await rels(zip, dr.path);
      const x = strip(dx);
      const anchors = [...els(x, 'twoCellAnchor').map((a) => ({ a, two: true })), ...els(x, 'oneCellAnchor').map((a) => ({ a, two: false }))];
      for (const { a, two } of anchors) {
        const m = /<chart\b[^>]*\b[\w]*:?id="([^"]+)"/.exec(a.body); if (!m) continue;
        const rel = drels.find((r) => r.id === m[1] && r.type === T_CHART); if (!rel) continue;
        const cx = await text(zip, rel.path); if (!cx) continue;
        const chart = parseChart(cx); if (!chart) continue;
        chart.id = `xl${++n}`;
        chart.anchor = anchorOf(a.body, two);
        chart._orig = { drawingPath: dr.path, rid: m[1], chartPath: rel.path };
        if (!out.has(sh.name)) out.set(sh.name, []);
        out.get(sh.name).push(chart);
      }
    }
  }
  return out;
}

/* ---------------- writing ---------------- */
const hex = (c, i) => String(c || PALETTE[i % PALETTE.length]).replace('#', '').toUpperCase().slice(0, 6);
const fill = (c) => `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>`;
const rich = (t, sz = 1400, b = 1) => `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="${b}"/></a:pPr><a:r><a:rPr lang="en-US" sz="${sz}" b="${b}"/><a:t>${esc(t)}</a:t></a:r></a:p></c:rich></c:tx>`;
const titleXml = (t, sz) => `<c:title>${rich(t, sz)}<c:overlay val="0"/></c:title>`;
const strCache = (arr) => `<c:strCache><c:ptCount val="${arr.length}"/>${arr.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v ?? '')}</c:v></c:pt>`).join('')}</c:strCache>`;
const numCache = (arr) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${arr.length}"/>${arr.map((v, i) => (typeof v === 'number' && Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numCache>`;
const numSrc = (ref, arr) => (ref ? `<c:numRef><c:f>${esc(ref)}</c:f>${numCache(arr)}</c:numRef>` : `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${arr.length}"/>${arr.map((v, i) => (typeof v === 'number' && Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numLit>`);
const strSrc = (ref, arr) => (ref ? `<c:strRef><c:f>${esc(ref)}</c:f>${strCache(arr)}</c:strRef>` : `<c:strLit><c:ptCount val="${arr.length}"/>${arr.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v ?? '')}</c:v></c:pt>`).join('')}</c:strLit>`);
const serTx = (s) => (s.refs && s.refs.name ? `<c:tx><c:strRef><c:f>${esc(s.refs.name)}</c:f>${strCache([s.name])}</c:strRef></c:tx>` : `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`);
/** series extras: data labels and trendline (they go right before <c:cat> in every series type) */
const serExtras = (chart, s, i) => {
  const src = chart.series && chart.series[i];
  const lbl = chart.labels || (src && src.labels);
  const tr = chart.trend || (src && src.trendline);
  return `${lbl ? '<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : ''}${tr ? `<c:trendline><c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="${hex(s.color, i)}"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></c:spPr><c:trendlineType val="${typeof tr === 'string' ? tr : 'linear'}"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>` : ''}`;
};
const noLbls = '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';

function axis(kind, id, cross, pos, title, opts = {}) {
  const t = title ? titleXml(title, 1000) : '';
  const common = `<c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${opts.grid ? '<c:majorGridlines><c:spPr><a:ln w="9525">' + fill('D9D9D9') + '</a:ln></c:spPr></c:majorGridlines>' : ''}${t}`;
  if (kind === 'cat') return `<c:catAx>${common}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  return `<c:valAx>${common}<c:numFmt formatCode="${opts.pct ? '0%' : 'General'}" sourceLinked="${opts.pct ? 0 : 1}"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="${opts.between || 'between'}"/></c:valAx>`;
}

/** DrawingML chart part for one chart + its resolved data */
export function chartXml(chart, data) {
  const type = chart.type || 'column';
  const S = (data && data.series) || [];
  const cats = (data && data.categories) || [];
  const grouping = chart.stacked === 'percent' ? 'percentStacked' : chart.stacked ? 'stacked' : null;
  const catXml = (s) => {
    const ref = s.refs && s.refs.cat;
    if (!cats.length) return '';
    return cats.every((c) => typeof c === 'number') ? `<c:cat>${numSrc(ref, cats)}</c:cat>` : `<c:cat>${strSrc(ref, cats.map((c) => String(c ?? '')))}</c:cat>`;
  };
  const valXml = (s) => `<c:val>${numSrc(s.refs && s.refs.val, s.values)}</c:val>`;
  let group = ''; let axes = '';
  const A1 = 111111111; const A2 = 222222222;
  const pie = type === 'pie' || type === 'doughnut';
  if (type === 'column' || type === 'bar') {
    const sers = S.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${serTx(s)}<c:spPr>${fill(hex(s.color, i))}</c:spPr><c:invertIfNegative val="0"/>${serExtras(chart, s, i)}${catXml(s)}${valXml(s)}</c:ser>`).join('');
    group = `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping || 'clustered'}"/><c:varyColors val="0"/>${sers}${noLbls}<c:gapWidth val="150"/>${grouping ? '<c:overlap val="100"/>' : ''}<c:axId val="${A1}"/><c:axId val="${A2}"/></c:barChart>`;
    axes = axis('cat', A1, A2, type === 'bar' ? 'l' : 'b', chart.xTitle) + axis('val', A2, A1, type === 'bar' ? 'b' : 'l', chart.yTitle, { grid: true, pct: chart.stacked === 'percent' });
  } else if (type === 'line') {
    const sers = S.map((s, i) => { const c = hex(s.color, i); return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${serTx(s)}<c:spPr><a:ln w="28575" cap="rnd">${fill(c)}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill('FFFFFF')}<a:ln w="12700">${fill(c)}</a:ln></c:spPr></c:marker>${serExtras(chart, s, i)}${catXml(s)}${valXml(s)}<c:smooth val="0"/></c:ser>`; }).join('');
    group = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${noLbls}<c:marker val="1"/><c:axId val="${A1}"/><c:axId val="${A2}"/></c:lineChart>`;
    axes = axis('cat', A1, A2, 'b', chart.xTitle) + axis('val', A2, A1, 'l', chart.yTitle, { grid: true });
  } else if (type === 'area') {
    const sers = S.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${serTx(s)}<c:spPr>${fill(hex(s.color, i))}</c:spPr>${serExtras(chart, { ...s, }, i).replace(/<c:trendline>[\s\S]*<\/c:trendline>/, '')}${catXml(s)}${valXml(s)}</c:ser>`).join('');
    group = `<c:areaChart><c:grouping val="${grouping || 'standard'}"/><c:varyColors val="0"/>${sers}${noLbls}<c:axId val="${A1}"/><c:axId val="${A2}"/></c:areaChart>`;
    axes = axis('cat', A1, A2, 'b', chart.xTitle) + axis('val', A2, A1, 'l', chart.yTitle, { grid: true, pct: chart.stacked === 'percent', between: 'midCat' });
  } else if (pie) {
    const colors = (chart.colors && chart.colors.length ? chart.colors : (data && data.pointColors) || PALETTE);
    const sers = S.slice(0, type === 'pie' ? 1 : S.length).map((s, i) => {
      const dpts = s.values.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr>${fill(hex(colors[k % colors.length], k))}<a:ln w="12700">${fill('FFFFFF')}</a:ln></c:spPr></c:dPt>`).join('');
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${serTx(s)}${dpts}<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:dLblPos val="${type === 'pie' ? 'ctr' : 'ctr'}"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>${catXml(s)}${valXml(s)}</c:ser>`;
    }).join('');
    group = type === 'pie' ? `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
    // dLblPos is not allowed for doughnut labels
    if (type === 'doughnut') group = group.replace(/<c:dLblPos val="[^"]*"\/>/g, '');
  } else if (type === 'scatter') {
    const sers = S.map((s, i) => { const c = hex(s.color, i); const xs = (s.refs && s.refs.cat) || cats.length ? `<c:xVal>${numSrc(s.refs && s.refs.cat, cats.map((v) => (typeof v === 'number' ? v : null)))}</c:xVal>` : '';
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${serTx(s)}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${fill(c)}<a:ln w="9525">${fill(c)}</a:ln></c:spPr></c:marker>${xs}<c:yVal>${numSrc(s.refs && s.refs.val, s.values)}</c:yVal><c:smooth val="0"/></c:ser>`; }).join('');
    group = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${noLbls}<c:axId val="${A1}"/><c:axId val="${A2}"/></c:scatterChart>`;
    axes = axis('val', A1, A2, 'b', chart.xTitle, { between: 'midCat' }) + axis('val', A2, A1, 'l', chart.yTitle, { grid: true, between: 'midCat' });
  }
  const legend = chart.legend && chart.legend !== 'none' ? `<c:legend><c:legendPos val="${({ right: 'r', bottom: 'b', top: 't', left: 'l' })[chart.legend] || 'r'}"/><c:overlay val="0"/></c:legend>` : '';
  const title = chart.title ? `${titleXml(chart.title, 1400)}<c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${group}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr>${fill('FFFFFF')}<a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>`;
}

export const anchorXml = (pfx, id, rid, from, to, name) => {
  const p = pfx ? `${pfx}:` : '';
  const pt = (t, a) => `<${p}${t}><${p}col>${Math.max(0, a.c | 0)}</${p}col><${p}colOff>${Math.max(0, Math.round((a.dx || 0) * EMU))}</${p}colOff><${p}row>${Math.max(0, a.r | 0)}</${p}row><${p}rowOff>${Math.max(0, Math.round((a.dy || 0) * EMU))}</${p}rowOff></${p}${t}>`;
  return `<${p}twoCellAnchor editAs="oneCell">${pt('from', from)}${pt('to', to)}<${p}graphicFrame macro=""><${p}nvGraphicFramePr><${p}cNvPr id="${id}" name="${esc(name)}"/><${p}cNvGraphicFramePr/></${p}nvGraphicFramePr><${p}xfrm><a:off xmlns:a="${NS.a}" x="0" y="0"/><a:ext xmlns:a="${NS.a}" cx="0" cy="0"/></${p}xfrm><a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData></a:graphic></${p}graphicFrame><${p}clientData/></${p}twoCellAnchor>`;
};

function nextNum(zip, re) { let n = 0; zip.forEach((p) => { const m = re.exec(p); if (m) n = Math.max(n, +m[1]); }); return n + 1; }
function nextRid(relsXml) { let n = 0; for (const m of (relsXml || '').matchAll(/Id="rId(\d+)"/g)) n = Math.max(n, +m[1]); return `rId${n + 1}`; }
function addRel(relsXml, id, type, target) {
  const x = relsXml || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rel}"></Relationships>`;
  return x.replace(/<\/Relationships>\s*$/, `<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"/></Relationships>`).replace(/<Relationships([^>]*)\/>\s*$/, `<Relationships$1><Relationship Id="${id}" Type="${type}" Target="${esc(target)}"/></Relationships>`);
}
function relTarget(from, to) { // relative path from part `from` to part `to`
  const a = dirOf(from).split('/').filter(Boolean); const b = to.split('/');
  let i = 0; while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return `${'../'.repeat(a.length - i)}${b.slice(i).join('/')}`;
}
const SHEET_AFTER = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/**
 * Adds charts to an xlsx zip written by exceljs.
 * list = [{ sheetName, charts: [chart & { _xl: { from, to } }], data: [{ categories, series }] }]
 */
export async function writeChartsIntoXlsx(zip, list) {
  const parts = await sheetParts(zip);
  let ct = await text(zip, '[Content_Types].xml');
  const override = (p, type) => { if (!ct.includes(`PartName="/${p}"`)) ct = ct.replace(/<\/Types>\s*$/, `<Override PartName="/${p}" ContentType="${type}"/></Types>`); };
  for (const item of list || []) {
    if (!item || !item.charts || !item.charts.length) continue;
    const sh = parts.find((p) => p.name === item.sheetName); if (!sh) continue;
    let sx = await text(zip, sh.path); if (!sx) continue;
    const sRelsPath = relsPath(sh.path);
    let sRels = await text(zip, sRelsPath);
    // existing drawing of this sheet?
    let drawingPath = null;
    const dm = /<(\w+:)?drawing\s[^>]*?(\w+:)?id="([^"]+)"/.exec(sx);
    if (dm) { const r = (await rels(zip, sh.path)).find((x) => x.id === dm[3]); if (r && r.path) drawingPath = r.path; }
    let dx; let pfx = 'xdr';
    if (drawingPath && zip.file(drawingPath)) {
      dx = await text(zip, drawingPath);
      const root = /<(?:(\w+):)?wsDr\b/.exec(dx); pfx = root ? root[1] || '' : 'xdr';
    } else {
      drawingPath = `xl/drawings/drawing${nextNum(zip, /^xl\/drawings\/drawing(\d+)\.xml$/)}.xml`;
      dx = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`;
      zip.file(drawingPath, dx);
      const rid = nextRid(sRels);
      sRels = addRel(sRels, rid, T_DRAWING, relTarget(sh.path, drawingPath));
      zip.file(sRelsPath, sRels);
      if (!/<worksheet\b[^>]*xmlns:r=/.test(sx)) sx = sx.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS.r}"`);
      const tag = `<drawing r:id="${rid}"/>`;
      let at = -1;
      for (const t of SHEET_AFTER) { const i = sx.search(new RegExp(`<${t}[\\s/>]`)); if (i >= 0 && (at < 0 || i < at)) at = i; }
      if (at < 0) at = sx.lastIndexOf('</worksheet>');
      sx = sx.slice(0, at) + tag + sx.slice(at);
      zip.file(sh.path, sx);
      override(drawingPath, CT_DRAWING);
    }
    const dRelsPath = relsPath(drawingPath);
    let dRels = await text(zip, dRelsPath);
    let nextId = 1; for (const m of dx.matchAll(/cNvPr\b[^>]*\bid="(\d+)"/g)) nextId = Math.max(nextId, +m[1] + 1);
    let anchors = '';
    item.charts.forEach((chart, i) => {
      const chartPath = `xl/charts/chart${nextNum(zip, /^xl\/charts\/chart(\d+)\.xml$/)}.xml`;
      zip.file(chartPath, chartXml(chart, item.data && item.data[i]));
      override(chartPath, CT_CHART);
      const rid = nextRid(dRels);
      dRels = addRel(dRels, rid, T_CHART, relTarget(drawingPath, chartPath));
      const xl = chart._xl || { from: { c: 0, r: 0, dx: 0, dy: 0 }, to: { c: 6, r: 15, dx: 0, dy: 0 } };
      anchors += anchorXml(pfx, nextId, rid, xl.from, xl.to, chart.title || `Chart ${nextId}`);
      nextId++;
    });
    const p = pfx ? `${pfx}:` : '';
    if (new RegExp(`</${p}wsDr>`).test(dx)) dx = dx.replace(new RegExp(`</${p}wsDr>\\s*$`), `${anchors}</${p}wsDr>`);
    else dx = dx.replace(new RegExp(`<${p}wsDr([^>]*)/>`), `<${p}wsDr$1>${anchors}</${p}wsDr>`);
    zip.file(drawingPath, dx);
    zip.file(dRelsPath, dRels);
  }
  zip.file('[Content_Types].xml', ct);
  return zip;
}

/**
 * Original Excel charts the user edited, moved or deleted: rewrite their chart part (edited) and patch / remove the
 * anchor in the drawing. Unchanged charts are never touched (byte-identical).
 * items = [{ chart (with _orig and _xl), data, deleted }]
 */
export async function patchOriginalCharts(zip, items) {
  const byDrawing = new Map();
  for (const it of items || []) { const o = it.chart && it.chart._orig; if (!o || !o.drawingPath) continue; if (!byDrawing.has(o.drawingPath)) byDrawing.set(o.drawingPath, []); byDrawing.get(o.drawingPath).push(it); }
  for (const [path, list] of byDrawing) {
    let dx = await text(zip, path); if (dx == null) continue;
    const root = /<(?:(\w+):)?wsDr\b/.exec(dx); const pfx = root ? root[1] || '' : 'xdr';
    for (const it of list) {
      const { rid, chartPath } = it.chart._orig;
      const re = /<((?:\w+:)?)(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/\1\2>/g;
      let m; let hit = null;
      while ((m = re.exec(dx))) { if (m[0].includes(`:id="${rid}"`) && /<(\w+:)?chart\b/.test(m[0])) { hit = m; break; } }
      if (!hit) continue;
      let repl = '';
      if (!it.deleted) {
        const idm = /cNvPr\b[^>]*\bid="(\d+)"/.exec(hit[0]);
        const xl = it.chart._xl || { from: { c: 0, r: 0, dx: 0, dy: 0 }, to: { c: 6, r: 15, dx: 0, dy: 0 } };
        repl = anchorXml(pfx, idm ? +idm[1] : 1000, rid, xl.from, xl.to, it.chart.title || 'Chart');
        if (it.rewrite) zip.file(chartPath, chartXml(it.chart, it.data));
      }
      dx = dx.slice(0, hit.index) + repl + dx.slice(hit.index + hit[0].length);
    }
    zip.file(path, dx);
  }
}
