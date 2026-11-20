/*
 * Edited Excel charts: patch the ORIGINAL chart part instead of writing a simple new one. Everything the editor does
 * not model (styles, effects, layout, number formats, extLst, data label formatting, trendline styles …) stays as it
 * is; only what changed is touched:
 *   title, axis titles, legend (position / on / off), series (formula refs, name, colour, caches; added series are
 *   cloned from the last one, removed ones dropped), stacking, type changes between column / bar / line / area.
 * Anything else (pie ↔ others, scatter, labels or trendline toggled …) → null: the caller writes the simple form.
 */
const escX = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** element <P:tag …>…</P:tag> (or self-closing) starting the search at `from`; depth-aware */
function find(xml, P, tag, from = 0, until = xml.length) {
  const open = new RegExp(`<${P}${tag}(?=[\\s/>])[^>]*?(/?)>`, 'g');
  open.lastIndex = from;
  const m = open.exec(xml);
  if (!m || m.index >= until) return null;
  if (m[1] === '/') return { start: m.index, end: open.lastIndex, inner: '', innerStart: open.lastIndex };
  const close = `</${P}${tag}>`;
  let depth = 1; let i = open.lastIndex;
  const any = new RegExp(`<${P}${tag}(?=[\\s/>])[^>]*?(/?)>|</${P}${tag}>`, 'g');
  any.lastIndex = i;
  let x;
  while ((x = any.exec(xml))) {
    if (x[0] === close) { depth--; if (!depth) return { start: m.index, end: any.lastIndex, inner: xml.slice(open.lastIndex, x.index), innerStart: open.lastIndex }; } else if (x[1] !== '/') depth++;
  }
  return null;
  // eslint-disable-next-line no-unreachable
  void i;
}
function findAll(xml, P, tag, from = 0, until = xml.length) {
  const out = []; let at = from;
  for (;;) { const e = find(xml, P, tag, at, until); if (!e) break; out.push(e); at = e.end; }
  return out;
}
const splice = (s, a, b, ins) => s.slice(0, a) + ins + s.slice(b);

const GROUP = { column: 'barChart', bar: 'barChart', line: 'lineChart', area: 'areaChart' };
const SIMPLE = new Set(['column', 'bar', 'line', 'area']);

function richTitle(P, A, text) {
  return `<${P}title><${P}tx><${P}rich><${A}bodyPr/><${A}lstStyle/><${A}p><${A}pPr><${A}defRPr/></${A}pPr><${A}r><${A}rPr lang="en-US"/><${A}t>${escX(text)}</${A}t></${A}r></${A}p></${P}rich></${P}tx><${P}overlay val="0"/></${P}title>`;
}
/** replace the text of a <c:title> keeping its formatting (first run gets the text, other runs go) */
function setTitleText(P, A, titleXml, text) {
  const rich = /<(\w+:)?rich>/.test(titleXml);
  if (!rich) return richTitle(P, A, text);
  let first = true;
  let x = titleXml.replace(new RegExp(`<${A}r>[\\s\\S]*?</${A}r>`, 'g'), (r) => { if (!first) return ''; first = false; return r.replace(new RegExp(`<${A}t>[\\s\\S]*?</${A}t>`), `<${A}t>${escX(text)}</${A}t>`); });
  if (first) x = x.replace(new RegExp(`</${A}p>`), `<${A}r><${A}rPr lang="en-US"/><${A}t>${escX(text)}</${A}t></${A}r></${A}p>`);
  // only the first paragraph keeps text
  let pi = 0; x = x.replace(new RegExp(`<${A}p>[\\s\\S]*?</${A}p>`, 'g'), (p) => (pi++ ? '' : p));
  return x;
}
function cacheStr(P, arr) { return `<${P}strCache><${P}ptCount val="${arr.length}"/>${arr.map((v, i) => `<${P}pt idx="${i}"><${P}v>${escX(v ?? '')}</${P}v></${P}pt>`).join('')}</${P}strCache>`; }
function cacheNum(P, arr, fmt) { return `<${P}numCache><${P}formatCode>${escX(fmt || 'General')}</${P}formatCode><${P}ptCount val="${arr.length}"/>${arr.map((v, i) => (typeof v === 'number' && Number.isFinite(v) ? `<${P}pt idx="${i}"><${P}v>${v}</${P}v></${P}pt>` : '')).join('')}</${P}numCache>`; }

/** set the reference (and cache) of a <c:tx>/<c:cat>/<c:val> element body */
function setRef(P, el, ref, values, kind) {
  const fm = /<(\w+:)?formatCode>([\s\S]*?)<\/(\w+:)?formatCode>/.exec(el);
  const fmt = fm ? fm[2] : 'General';
  const numeric = kind === 'val' || (kind === 'cat' && values.length && values.every((v) => typeof v === 'number'));
  if (!ref) return kind === 'tx' ? `<${P}v>${escX(values[0] ?? '')}</${P}v>` : numeric ? `<${P}numLit><${P}formatCode>General</${P}formatCode><${P}ptCount val="${values.length}"/>${values.map((v, i) => (typeof v === 'number' ? `<${P}pt idx="${i}"><${P}v>${v}</${P}v></${P}pt>` : '')).join('')}</${P}numLit>` : `<${P}strLit><${P}ptCount val="${values.length}"/>${values.map((v, i) => `<${P}pt idx="${i}"><${P}v>${escX(v)}</${P}v></${P}pt>`).join('')}</${P}strLit>`;
  // keep multi-level refs as they are when the ref did not change
  return numeric ? `<${P}numRef><${P}f>${escX(ref)}</${P}f>${cacheNum(P, values, fmt)}</${P}numRef>` : `<${P}strRef><${P}f>${escX(ref)}</${P}f>${cacheStr(P, values.map((v) => (v == null ? '' : String(v))))}</${P}strRef>`;
}
const refIn = (body) => { const m = /<(\w+:)?f>([\s\S]*?)<\/(\w+:)?f>/.exec(body || ''); return m ? m[2].replace(/&amp;/g, '&').replace(/&apos;/g, "'").replace(/&quot;/g, '"').trim() : null; };

function patchSeries(xml, P, A, gStart, gEnd, chart, data, oldChart) {
  let g = xml.slice(gStart, gEnd);
  let sers = findAll(g, P, 'ser');
  if (!sers.length) return null;
  const want = chart.series || [];
  const D = (data && data.series) || [];
  // add / remove series
  if (want.length > sers.length) {
    const last = g.slice(sers[sers.length - 1].start, sers[sers.length - 1].end);
    let ins = '';
    for (let i = sers.length; i < want.length; i++) ins += last.replace(new RegExp(`<${P}extLst>[\\s\\S]*?</${P}extLst>`), '').replace(/<(\w+:)?idx val="\d+"\/>/, `<${P}idx val="${i + 100}"/>`).replace(/<(\w+:)?order val="\d+"\/>/, `<${P}order val="${i}"/>`);
    g = splice(g, sers[sers.length - 1].end, sers[sers.length - 1].end, ins);
    sers = findAll(g, P, 'ser');
  } else if (want.length < sers.length) {
    for (let i = sers.length - 1; i >= want.length; i--) g = splice(g, sers[i].start, sers[i].end, '');
    sers = findAll(g, P, 'ser');
  }
  for (let i = sers.length - 1; i >= 0; i--) {
    const s = sers[i]; let body = g.slice(s.start, s.end);
    const w = want[i] || {}; const o = (oldChart.series || [])[i] || {}; const d = D[i] || { values: [] };
    const fresh = i >= (oldChart.series || []).length;
    // name
    if (fresh || w.nameRef !== o.nameRef || w.name !== o.name) {
      const tx = find(body, P, 'tx');
      const nx = `<${P}tx>${setRef(P, tx ? tx.inner : '', w.nameRef || null, [d.name ?? w.name ?? ''], 'tx')}</${P}tx>`;
      if (tx) body = splice(body, tx.start, tx.end, nx);
      else { const ord = find(body, P, 'order'); if (ord) body = splice(body, ord.end, ord.end, nx); }
    }
    for (const [tags, key, oldKey, kind] of [[['cat', 'xVal'], 'catRef', 'catRef', 'cat'], [['val', 'yVal'], 'valRef', 'valRef', 'val']]) {
      if (!fresh && w[key] === o[oldKey]) continue;
      let el = null; let tag = tags[0];
      for (const t of tags) { el = find(body, P, t); if (el) { tag = t; break; } }
      const values = kind === 'cat' ? (data.categories || []) : (d.values || []);
      const nx = `<${P}${tag}>${setRef(P, el ? el.inner : '', w[key] || null, values, kind)}</${P}${tag}>`;
      if (el) body = splice(body, el.start, el.end, nx);
      else if (kind === 'cat') { const v = find(body, P, 'val') || find(body, P, 'yVal'); if (v) body = splice(body, v.start, v.start, nx); }
    }
    if (fresh || (w.color || null) !== (o.color || null)) {
      if (w.color) {
        const hex = w.color.replace('#', '').toUpperCase().slice(0, 6);
        const line = chart.type === 'line';
        const sp = find(body, P, 'spPr');
        const fillX = `<${A}solidFill><${A}srgbClr val="${hex}"/></${A}solidFill>`;
        if (sp) {
          let spx = body.slice(sp.start, sp.end);
          if (line) { const ln = find(spx, A, 'ln'); if (ln) { const lx = spx.slice(ln.start, ln.end).replace(new RegExp(`<${A}(solidFill|noFill|gradFill)[\\s\\S]*?</${A}\\1>|<${A}noFill/>`), ''); spx = splice(spx, ln.start, ln.end, lx.replace(/(<(\w+:)?ln\b[^>]*?)(\/?)>/, (m0, a, b, c) => (c ? `${a}>${fillX}</${A}ln>` : `${a}>${fillX}`))); } else spx = spx.replace(/(<(\w+:)?spPr\b[^>]*>)/, `$1<${A}ln w="28575"><${A}solidFill><${A}srgbClr val="${hex}"/></${A}solidFill></${A}ln>`); } else {
            const f = new RegExp(`<${A}(solidFill|gradFill|pattFill)\\b[\\s\\S]*?</${A}\\1>|<${A}noFill/>`).exec(spx);
            const lnAt = spx.indexOf(`<${A}ln`);
            if (f && (lnAt < 0 || f.index < lnAt)) spx = splice(spx, f.index, f.index + f[0].length, fillX);
            else spx = spx.replace(/(<(\w+:)?spPr\b[^>]*>)/, `$1${fillX}`);
          }
          body = splice(body, sp.start, sp.end, spx);
        } else {
          const anchor = find(body, P, 'tx') || find(body, P, 'order');
          if (anchor) body = splice(body, anchor.end, anchor.end, `<${P}spPr>${line ? `<${A}ln w="28575">${fillX}</${A}ln>` : fillX}</${P}spPr>`);
        }
      }
    }
    g = splice(g, s.start, s.end, body);
  }
  return xml.slice(0, gStart) + g + xml.slice(gEnd);
}

/** change the chart group type (column/bar/line/area) in place; returns xml or null */
function retype(xml, P, from, to, stacked) {
  const fromTag = GROUP[from]; const toTag = GROUP[to];
  const g = find(xml, P, fromTag); if (!g) return null;
  let inner = g.inner;
  const openTag = xml.slice(g.start, g.innerStart).replace(fromTag, toTag);
  const grp = stacked === 'stacked' ? 'stacked' : stacked === 'percent' ? 'percentStacked' : (toTag === 'barChart' ? 'clustered' : 'standard');
  const drop = (tags, s) => tags.reduce((x, t) => x.replace(new RegExp(`<${P}${t}\\b[^>]*/>|<${P}${t}\\b[^>]*>[\\s\\S]*?</${P}${t}>`, 'g'), ''), s);
  if (fromTag !== toTag) {
    if (fromTag === 'barChart') inner = drop(['barDir', 'gapWidth', 'overlap', 'serLines', 'shape'], inner).replace(new RegExp(`<${P}ser>[\\s\\S]*?</${P}ser>`, 'g'), (s) => drop(['invertIfNegative', 'pictureOptions', 'shape'], s));
    if (fromTag === 'lineChart') inner = drop(['hiLowLines', 'upDownBars', 'marker', 'smooth'], inner);
    if (fromTag === 'areaChart') inner = drop([], inner);
    if (toTag === 'barChart') {
      inner = `<${P}barDir val="${to === 'bar' ? 'bar' : 'col'}"/>${inner}`;
      const ax = inner.indexOf(`<${P}axId`);
      inner = splice(inner, ax, ax, `<${P}gapWidth val="150"/>${grp !== 'clustered' ? `<${P}overlap val="100"/>` : ''}`);
    }
    if (toTag === 'lineChart') { const ax = inner.indexOf(`<${P}axId`); inner = splice(inner, ax, ax, `<${P}marker val="1"/>`); }
  } else if (toTag === 'barChart') {
    inner = inner.replace(/<(\w+:)?barDir val="\w+"\/>/, `<${P}barDir val="${to === 'bar' ? 'bar' : 'col'}"/>`);
  }
  // grouping
  if (/<(\w+:)?grouping\b/.test(inner)) inner = inner.replace(/<(\w+:)?grouping val="\w+"\/>/, `<${P}grouping val="${grp}"/>`);
  else { const bd = find(inner, P, 'barDir'); const at = bd ? bd.end : 0; inner = splice(inner, at, at, `<${P}grouping val="${grp}"/>`); }
  if (toTag === 'barChart') {
    const hasOv = /<(\w+:)?overlap\b/.test(inner);
    if (grp !== 'clustered' && hasOv) inner = inner.replace(/<(\w+:)?overlap val="-?\d+"\/>/, `<${P}overlap val="100"/>`);
    else if (grp !== 'clustered') { const gw = find(inner, P, 'gapWidth'); if (gw) inner = splice(inner, gw.end, gw.end, `<${P}overlap val="100"/>`); }
    else if (hasOv && /<(\w+:)?overlap val="100"\/>/.test(inner)) inner = inner.replace(/<(\w+:)?overlap val="100"\/>/, '');
  }
  let out = splice(xml, g.start, g.end, `${openTag}${inner}</${P}${toTag}>`);
  // bar ↔ column: the axes swap sides
  if (fromTag === 'barChart' && toTag === 'barChart' && from !== to) out = out.replace(new RegExp(`(<${P}catAx>[\\s\\S]*?<${P}axPos val=")(\\w)(")`), (m0, a, p, c) => `${a}${to === 'bar' ? 'l' : 'b'}${c}`).replace(new RegExp(`(<${P}valAx>[\\s\\S]*?<${P}axPos val=")(\\w)(")`), (m0, a, p, c) => `${a}${to === 'bar' ? 'b' : 'l'}${c}`);
  return out;
}

function setAxisTitle(xml, P, A, axTag, text) {
  const ax = find(xml, P, axTag); if (!ax) return xml;
  let body = xml.slice(ax.start, ax.end);
  const t = find(body, P, 'title');
  if (t) body = text ? splice(body, t.start, t.end, setTitleText(P, A, body.slice(t.start, t.end), text)) : splice(body, t.start, t.end, '');
  else if (text) {
    const after = find(body, P, 'minorGridlines') || find(body, P, 'majorGridlines') || find(body, P, 'axPos');
    if (after) body = splice(body, after.end, after.end, richTitle(P, A, text).replace(`<${P}overlay val="0"/>`, `<${P}overlay val="0"/>`));
  }
  return splice(xml, ax.start, ax.end, body);
}

/**
 * original chart XML + the chart as read (oldChart = parsed _sig) + the edited chart + resolved data → patched XML,
 * or null when the change cannot be expressed as a patch.
 */
export function patchChartXml(xml, oldChart, chart, data) {
  if (!xml || !oldChart || !chart) return null;
  const root = /<(?:(\w+):)?chartSpace\b/.exec(xml); if (!root) return null;
  const P = root[1] ? `${root[1]}:` : '';
  const am = /xmlns:(\w+)="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/main"/.exec(xml);
  const A = am ? `${am[1]}:` : 'a:';
  const keys = new Set([...Object.keys(oldChart), ...Object.keys(chart)].filter((k) => !k.startsWith('_') && k !== 'id' && k !== 'anchor'));
  const changed = [...keys].filter((k) => JSON.stringify(oldChart[k] ?? null) !== JSON.stringify(chart[k] ?? null));
  const ok = new Set(['title', 'legend', 'xTitle', 'yTitle', 'series', 'type', 'stacked']);
  if (changed.some((k) => !ok.has(k))) return null;
  if (oldChart._partial || chart.range || !(chart.series && chart.series.length)) return null;
  if ((oldChart.series || []).some((s, i) => chart.series[i] && (s.labels !== chart.series[i].labels || s.trendline !== chart.series[i].trendline))) return null;
  if (changed.includes('type') && !(SIMPLE.has(oldChart.type) && SIMPLE.has(chart.type))) return null;
  let x = xml;
  // type / stacking
  if (changed.includes('type') || changed.includes('stacked')) { x = retype(x, P, oldChart.type, chart.type, chart.stacked); if (!x) return null; }
  // series
  if (changed.includes('series')) {
    const plot = find(x, P, 'plotArea'); if (!plot) return null;
    const tag = GROUP[chart.type] || ({ pie: 'pieChart', doughnut: 'doughnutChart', scatter: 'scatterChart' })[chart.type];
    const g = find(x, P, tag, plot.start, plot.end) || find(x, P, tag.replace('Chart', '3DChart'), plot.start, plot.end);
    if (!g) return null;
    x = patchSeries(x, P, A, g.start, g.end, chart, data, oldChart);
    if (!x) return null;
  }
  // title
  if (changed.includes('title')) {
    const ch = find(x, P, 'chart'); if (!ch) return null;
    const plot = find(x, P, 'plotArea');
    const t = find(x, P, 'title', ch.innerStart, plot ? plot.start : ch.end);
    if (t) x = chart.title ? splice(x, t.start, t.end, setTitleText(P, A, x.slice(t.start, t.end), chart.title)) : splice(x, t.start, t.end, '');
    else if (chart.title) x = splice(x, ch.innerStart, ch.innerStart, richTitle(P, A, chart.title));
    x = x.replace(/<(\w+:)?autoTitleDeleted val="\w+"\/>/, `<${P}autoTitleDeleted val="${chart.title ? 0 : 1}"/>`);
  }
  if (changed.includes('xTitle')) x = setAxisTitle(x, P, A, find(x, P, 'catAx') ? 'catAx' : find(x, P, 'dateAx') ? 'dateAx' : 'valAx', chart.xTitle);
  if (changed.includes('yTitle')) { const vals = findAll(x, P, 'valAx'); if (vals.length) x = setAxisTitle(x, P, A, 'valAx', chart.yTitle); }
  if (changed.includes('legend')) {
    const POS = { right: 'r', left: 'l', top: 't', bottom: 'b' };
    const lg = find(x, P, 'legend');
    if (chart.legend === 'none' || !chart.legend) { if (lg) x = splice(x, lg.start, lg.end, ''); } else if (lg) {
      let body = x.slice(lg.start, lg.end);
      body = /<(\w+:)?legendPos\b/.test(body) ? body.replace(/<(\w+:)?legendPos val="\w+"\/>/, `<${P}legendPos val="${POS[chart.legend] || 'r'}"/>`) : body.replace(/(<(\w+:)?legend\b[^>]*>)/, `$1<${P}legendPos val="${POS[chart.legend] || 'r'}"/>`);
      body = body.replace(/<(\w+:)?layout>[\s\S]*?<\/(\w+:)?layout>/, ''); // a manual layout belongs to the old position
      x = splice(x, lg.start, lg.end, body);
    } else {
      const plot = find(x, P, 'plotArea');
      if (plot) x = splice(x, plot.end, plot.end, `<${P}legend><${P}legendPos val="${POS[chart.legend] || 'r'}"/><${P}overlay val="0"/></${P}legend>`);
    }
  }
  return x;
}
