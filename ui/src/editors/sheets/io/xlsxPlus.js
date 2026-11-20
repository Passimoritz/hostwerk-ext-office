/*
 * The 1.1 structures of the xlsx reader/writer (kept apart from xlsx.js): tables, pivot tables + caches, x14 sparklines,
 * icon sets, pictures and shapes in drawings (insert / move / resize / delete, originals patched in place), slicers
 * (read, shown), rich text runs of shared strings, iterative calculation, and the Hostwerk sidecar part
 * (customXml/item*.xml: views, live-data and database bindings, forms – Excel shows the plain values and keeps the part).
 * String-based like xlsx.js, runs in Node.
 */
import { parseRange, rangeA1, MAXC, colName } from '../formula/refs.js';
import { esc, escText, unesc, attrs, elements, first, parseRels, relsPathOf, relTarget, addRel, nextRid, removeRels, ctOverride, ctDefault, ctRemove } from './xml.js';
import { readTableXml, tableXml, patchTableXml } from '../tables.js';
import { readPivot, computePivot, pivotParts, pivotSig } from '../pivot.js';

export const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
export const NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const URI_SPARK = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
const EXT_ORDER = ['{78C0D931-6437-407d-A8EE-F0AAD7539E65}', '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}', URI_SPARK];
export const REL2 = {
  pivotCache: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
  pivotRecords: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
  pivotTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
  table: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  customXml: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
  customXmlProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps',
  slicer: 'http://schemas.microsoft.com/office/2007/relationships/slicer',
  slicerCache: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
};
const CT2 = {
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  pivotCache: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotRecords: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  customXmlProps: 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
};
const IMG_CT = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };
const EMU = 9525;
const K = (r, c) => r * MAXC + c;
const UTF8 = new TextDecoder('utf-8');
async function str(zip, p) { const f = zip.file(p); if (!f) return null; const u = await f.async('uint8array'); return UTF8.decode(u); } // native UTF-8 decoding: much faster than JSZip's string path for big sheets
const nextNum = (zip, re) => { let n = 0; zip.forEach((p) => { const m = re.exec(p); if (m) n = Math.max(n, +m[1]); }); return n + 1; };
const colorOfA = (a, theme) => { if (!a) return null; if (a.rgb) return `#${a.rgb.slice(-6)}`; if (a.theme != null && theme) { const i = +a.theme; const map = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]; return `#${theme[map[i]] || '000000'}`; } return null; };

/* ---------------- rich text (shared strings) ---------------- */
/** <si> inner XML → runs [{ t, b, i, u, st, color, size, font }] or null when it is plain text */
export function richRuns(siXml) {
  if (!/<r>/.test(siXml)) return null;
  const runs = [];
  for (const m of siXml.matchAll(/<r>([\s\S]*?)<\/r>/g)) {
    const pr = /<rPr>([\s\S]*?)<\/rPr>/.exec(m[1]);
    const tx = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/.exec(m[1]);
    const run = { t: unesc(tx ? tx[1] : '').replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16))) };
    if (pr) {
      const p = pr[1];
      if (/<b(\s+val="(1|true)")?\s*\/>/.test(p)) run.b = true;
      if (/<i(\s+val="(1|true)")?\s*\/>/.test(p)) run.i = true;
      if (/<strike(\s+val="(1|true)")?\s*\/>/.test(p)) run.st = true;
      const u = /<u(?:\s+val="(\w+)")?\s*\/>/.exec(p); if (u) run.u = u[1] === 'double' ? 'double' : true;
      const sz = /<sz\s+val="([\d.]+)"/.exec(p); if (sz) run.size = +sz[1];
      const f = /<rFont\s+val="([^"]*)"/.exec(p); if (f) run.font = unesc(f[1]);
      const c = /<color\s+rgb="(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{6})"/.exec(p); if (c) run.color = `#${c[1]}`;
      const va = /<vertAlign\s+val="(\w+)"/.exec(p); if (va && va[1] !== 'baseline') run.va = va[1];
    }
    runs.push(run);
  }
  return runs.length ? runs : null;
}
/** runs → <si> XML */
export function richXml(runs) {
  return `<si>${runs.map((r) => {
    const p = [r.b ? '<b/>' : '', r.i ? '<i/>' : '', r.st ? '<strike/>' : '', r.u ? `<u${r.u === 'double' ? ' val="double"' : ''}/>` : '', r.va ? `<vertAlign val="${r.va}"/>` : '', r.size ? `<sz val="${r.size}"/>` : '', r.color ? `<color rgb="FF${r.color.replace('#', '').toUpperCase()}"/>` : '', r.font ? `<rFont val="${esc(r.font)}"/>` : ''].join('');
    return `<r>${p ? `<rPr>${p}</rPr>` : ''}<t xml:space="preserve">${escText(r.t)}</t></r>`;
  }).join('')}</si>`;
}
export const runsText = (runs) => runs.map((r) => r.t).join('');

/* ---------------- icon sets ---------------- */
export const ICON_SETS = ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '3Stars', '3Triangles', '5Boxes'];
export function readIconSet(ruleInner) {
  const is = /<iconSet\b([^>]*)>([\s\S]*?)<\/iconSet>/.exec(ruleInner) || /<iconSet\b([^>]*)\/>/.exec(ruleInner);
  if (!is) return null;
  const a = attrs(is[1]);
  const stops = [...(is[2] || '').matchAll(/<cfvo\b([^>]*)\/>/g)].map((m) => { const q = attrs(m[1]); return { type: q.type, val: q.val, gte: q.gte !== '0' }; });
  return { set: a.iconSet || '3TrafficLights1', stops, reverse: a.reverse === '1', showValue: a.showValue !== '0' };
}
export function iconSetXml(rule, prio) {
  const n = +String(rule.set || '3TrafficLights1')[0] || 3;
  const stops = rule.stops && rule.stops.length === n ? rule.stops : Array.from({ length: n }, (_, i) => ({ type: 'percent', val: String(Math.round((i * 100) / n)) }));
  return `<cfRule type="iconSet" priority="${prio}"${rule.stop ? ' stopIfTrue="1"' : ''}><iconSet${rule.set && rule.set !== '3TrafficLights1' ? ` iconSet="${rule.set}"` : ''}${rule.showValue === false ? ' showValue="0"' : ''}${rule.reverse ? ' reverse="1"' : ''}>${stops.map((s) => `<cfvo type="${s.type || 'percent'}" val="${esc(s.val ?? 0)}"${s.gte === false ? ' gte="0"' : ''}/>`).join('')}</iconSet></cfRule>`;
}

/* ---------------- sparklines (x14) ---------------- */
export function readSparklines(sheetXml, theme) {
  const out = [];
  const ext = /<(\w+:)?ext\b[^>]*uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}"[^>]*>[\s\S]*?<\/\1?ext>/i.exec(sheetXml);
  if (!ext) return out;
  for (const m of ext[0].matchAll(/<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g)) {
    const a = attrs(m[1]);
    const col = (tag) => { const e = new RegExp(`<x14:${tag}\\b([^>]*)/>`).exec(m[2]); return e ? colorOfA(attrs(e[1]), theme) : null; };
    const items = [...m[2].matchAll(/<x14:sparkline>\s*<xm:f>([\s\S]*?)<\/xm:f>\s*<xm:sqref>([\s\S]*?)<\/xm:sqref>\s*<\/x14:sparkline>/g)].map((q) => { const g = parseRange(q[2].trim()); return g ? { ref: unesc(q[1]).trim(), r: g.r1, c: g.c1 } : null; }).filter(Boolean);
    const g = {
      id: `sp${Math.random().toString(36).slice(2, 8)}`, type: a.type === 'column' ? 'column' : a.type === 'stacked' ? 'stacked' : 'line',
      color: col('colorSeries') || '#376092', neg: col('colorNegative') || '#D00000', markerColor: col('colorMarkers') || '#D00000', highColor: col('colorHigh') || '#D00000', lowColor: col('colorLow') || '#D00000', firstColor: col('colorFirst'), lastColor: col('colorLast'),
      markers: a.markers === '1', high: a.high === '1', low: a.low === '1', first: a.first === '1', last: a.last === '1', negative: a.negative === '1',
      weight: a.lineWeight ? +a.lineWeight : 0.75, empty: a.displayEmptyCellsAs || 'gap', min: a.manualMin != null ? +a.manualMin : null, max: a.manualMax != null ? +a.manualMax : null,
      items, _raw: m[0],
    };
    g._sig = sparkSig(g);
    out.push(g);
  }
  return out;
}
export const sparkSig = (g) => { const { id, _raw, _sig, _dirty, ...rest } = g; return JSON.stringify(rest); };
const rgbA = (c) => `FF${String(c || '#000000').replace('#', '').toUpperCase()}`;
export function sparkGroupXml(g) {
  const flags = ['markers', 'high', 'low', 'first', 'last', 'negative'].filter((k) => g[k]).map((k) => ` ${k}="1"`).join('');
  return `<x14:sparklineGroup${g.type !== 'line' ? ` type="${g.type}"` : ''}${g.weight && g.weight !== 0.75 ? ` lineWeight="${g.weight}"` : ''} displayEmptyCellsAs="${g.empty || 'gap'}"${flags}${g.min != null ? ` minAxisType="custom" manualMin="${g.min}"` : ''}${g.max != null ? ` maxAxisType="custom" manualMax="${g.max}"` : ''}><x14:colorSeries rgb="${rgbA(g.color)}"/><x14:colorNegative rgb="${rgbA(g.neg)}"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="${rgbA(g.markerColor)}"/><x14:colorFirst rgb="${rgbA(g.firstColor || g.markerColor)}"/><x14:colorLast rgb="${rgbA(g.lastColor || g.markerColor)}"/><x14:colorHigh rgb="${rgbA(g.highColor)}"/><x14:colorLow rgb="${rgbA(g.lowColor)}"/><x14:sparklines>${g.items.map((it) => `<x14:sparkline><xm:f>${escText(it.ref)}</xm:f><xm:sqref>${colName(it.c)}${it.r + 1}</xm:sqref></x14:sparkline>`).join('')}</x14:sparklines></x14:sparklineGroup>`;
}
/** new extLst of a worksheet with the sparkline ext replaced (other exts kept, schema order respected); null = no extLst */
export function sheetExtLst(sheetXml, groups, origGroups) {
  const changed = groups.length !== origGroups.length || groups.some((g, i) => !g._raw || g._sig !== sparkSig(g) || origGroups[i] !== g._raw);
  const cur = /<extLst>([\s\S]*?)<\/extLst>/.exec(sheetXml);
  if (!changed) return undefined; // untouched
  const exts = cur ? [...cur[1].matchAll(/<((?:\w+:)?)ext\b[^>]*>[\s\S]*?<\/\1ext>/g)].map((m) => ({ xml: m[0], uri: (/uri="([^"]+)"/.exec(m[0]) || [])[1] || '' })) : [];
  const rest = exts.filter((e) => e.uri.toUpperCase() !== URI_SPARK.toUpperCase());
  if (groups.length) rest.push({ uri: URI_SPARK, xml: `<ext uri="${URI_SPARK}" xmlns:x14="${NS_X14}"><x14:sparklineGroups xmlns:xm="${NS_XM}">${groups.map((g) => (g._raw && g._sig === sparkSig(g) ? g._raw : sparkGroupXml(g))).join('')}</x14:sparklineGroups></ext>` });
  const rank = (u) => { const i = EXT_ORDER.findIndex((x) => x.toUpperCase() === u.toUpperCase()); return i < 0 ? 10 : i; };
  rest.sort((a, b) => rank(a.uri) - rank(b.uri));
  return rest.length ? `<extLst>${rest.map((e) => e.xml).join('')}</extLst>` : null;
}

/* ---------------- drawings: pictures, shapes, slicers ---------------- */
const anchorRe = /<((?:\w+:)?)(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/\1\2>/g;
function anchorPos(an, pfx) {
  const pt = (tag) => { const e = new RegExp(`<${pfx}${tag}>([\\s\\S]*?)</${pfx}${tag}>`).exec(an); if (!e) return null; const g = (t) => +((new RegExp(`<${pfx}${t}>(-?\\d+)</${pfx}${t}>`).exec(e[1]) || [0, 0])[1]); return { c: g('col'), r: g('row'), dx: Math.round(g('colOff') / EMU), dy: Math.round(g('rowOff') / EMU) }; };
  const ext = /<(?:\w+:)?ext\s+cx="(\d+)"\s+cy="(\d+)"/.exec(an);
  return { from: pt('from'), to: pt('to'), size: ext ? { w: Math.round(+ext[1] / EMU), h: Math.round(+ext[2] / EMU) } : null };
}
/** shapes (sp) and slicers of a drawing → { shapes, slicers }; pictures are read by xlsx.js (it adds _orig ids via picIds) */
export function readDrawingExtras(dx, drawingPath, anchorFrom, theme) {
  const shapes = []; const slicers = []; const pics = [];
  const root = /<(?:(\w+):)?wsDr\b/.exec(dx); const pfx = root && root[1] ? `${root[1]}:` : '';
  for (const m of dx.matchAll(anchorRe)) {
    const an = m[0];
    const pos = anchorPos(an, pfx);
    if (!pos.from) continue;
    const anchor = anchorFrom(pos.from, m[2] === 'twoCellAnchor' ? pos.to : null, pos.size || { w: 120, h: 80 });
    const id = (/cNvPr\b[^>]*\bid="(\d+)"/.exec(an) || [])[1];
    const name = unesc((/cNvPr\b[^>]*\bname="([^"]*)"/.exec(an) || [])[1] || '');
    if (/<(\w+:)?pic\b/.test(an)) { pics.push({ id, anchor }); continue; }
    const sl = /<sle:slicer\b[^>]*\bname="([^"]*)"/.exec(an);
    if (sl) { slicers.push({ id: `sl${id}`, name: unesc(sl[1]), anchor }); continue; }
    if (/<(\w+:)?graphicFrame\b/.test(an) || /<(\w+:)?grpSp\b/.test(an) || /<(\w+:)?cxnSp\b/.test(an)) continue;
    const sp = /<(\w+:)?sp\b[\s\S]*?<\/\1sp>/.exec(an); if (!sp) continue;
    const prst = (/<a:prstGeom\s+prst="(\w+)"/.exec(sp[0]) || [])[1] || 'rect';
    const spPr = (/<(?:\w+:)?spPr\b[\s\S]*?<\/(?:\w+:)?spPr>/.exec(sp[0]) || [''])[0];
    const fillM = /<a:solidFill>\s*<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(spPr.replace(/<a:ln\b[\s\S]*?<\/a:ln>/, ''));
    const lnM = /<a:ln\b[\s\S]*?<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(spPr);
    const text = [...sp[0].matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((p) => [...p[1].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((t) => unesc(t[1])).join('')).join('\n');
    const noFill = /<a:noFill\/>/.test(spPr.replace(/<a:ln\b[\s\S]*?<\/a:ln>/, ''));
    const sh = { id: `shp${id}`, kind: /TextBox="1"/.test(an) ? 'textbox' : prst, anchor, fill: noFill ? null : fillM ? `#${fillM[1]}` : /<(\w+:)?style>/.test(sp[0]) ? '#4472C4' : null, line: lnM ? `#${lnM[1]}` : /<a:ln\b[^>]*>\s*<a:noFill/.test(spPr) ? null : '#2F528F', text, name, _orig: { drawingPath, cid: id } };
    sh._sig = shapeSig(sh);
    shapes.push(sh);
  }
  return { shapes, slicers, pics };
}
export const shapeSig = (s) => JSON.stringify({ k: s.kind, a: s.anchor, f: s.fill, l: s.line, t: s.text, c: s.textColor, sz: s.fontSize });
export const picSig = (p) => JSON.stringify({ a: p.anchor, s: p.src.length });

function xlAnchorOf(sh, a) {
  const cw = (c) => (sh.hiddenCols.has(c) ? 0 : sh.colW.get(c) ?? sh.defW);
  const rh = (r) => (sh.hiddenRows.has(r) ? 0 : sh.rowH.get(r) ?? sh.defH);
  let c = a.c; let x = (a.dx || 0) + a.w; while (x > cw(c) && c < MAXC - 1) { x -= cw(c); c++; }
  let r = a.r; let y = (a.dy || 0) + a.h; while (y > rh(r) && r < 1048575) { y -= rh(r); r++; }
  return { from: { c: a.c, r: a.r, dx: a.dx || 0, dy: a.dy || 0 }, to: { c, r, dx: Math.round(x), dy: Math.round(y) } };
}
const ptXml = (p, t, a) => `<${p}${t}><${p}col>${Math.max(0, a.c | 0)}</${p}col><${p}colOff>${Math.max(0, Math.round((a.dx || 0) * EMU))}</${p}colOff><${p}row>${Math.max(0, a.r | 0)}</${p}row><${p}rowOff>${Math.max(0, Math.round((a.dy || 0) * EMU))}</${p}rowOff></${p}${t}>`;
const PRST = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rightArrow: 'rightArrow', line: 'line', textbox: 'rect', star5: 'star5', diamond: 'diamond', callout: 'wedgeRoundRectCallout', wedgeRoundRectCallout: 'wedgeRoundRectCallout' };
function shapeXml(p, s, id, xl) {
  const fill = s.fill ? `<a:solidFill><a:srgbClr val="${s.fill.replace('#', '').toUpperCase()}"/></a:solidFill>` : '<a:noFill/>';
  const ln = s.line ? `<a:ln w="12700"><a:solidFill><a:srgbClr val="${s.line.replace('#', '').toUpperCase()}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>';
  const color = (s.textColor || (s.fill && s.kind !== 'textbox' ? '#FFFFFF' : '#000000')).replace('#', '').toUpperCase();
  const paras = String(s.text || '').split('\n').map((l) => `<a:p>${s.kind !== 'textbox' ? '<a:pPr algn="ctr"/>' : ''}${l ? `<a:r><a:rPr lang="de-DE" sz="${Math.round((s.fontSize || 11) * 100)}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${escText(l)}</a:t></a:r>` : `<a:endParaRPr lang="de-DE" sz="${Math.round((s.fontSize || 11) * 100)}"/>`}</a:p>`).join('');
  return `<${p}twoCellAnchor editAs="oneCell">${ptXml(p, 'from', xl.from)}${ptXml(p, 'to', xl.to)}<${p}sp macro="" textlink=""><${p}nvSpPr><${p}cNvPr id="${id}" name="${esc(s.name || `${s.kind === 'textbox' ? 'TextBox' : 'Shape'} ${id}`)}"/><${p}cNvSpPr${s.kind === 'textbox' ? ' txBox="1"' : ''}/></${p}nvSpPr><${p}spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(s.anchor.w * EMU)}" cy="${Math.round(s.anchor.h * EMU)}"/></a:xfrm><a:prstGeom prst="${PRST[s.kind] || 'rect'}"><a:avLst/></a:prstGeom>${fill}${ln}</${p}spPr><${p}txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="${s.kind === 'textbox' ? 't' : 'ctr'}"/><a:lstStyle/>${paras}</${p}txBody></${p}sp><${p}clientData/></${p}twoCellAnchor>`;
}
function picXml(p, im, id, rid, xl) {
  return `<${p}twoCellAnchor editAs="oneCell">${ptXml(p, 'from', xl.from)}${ptXml(p, 'to', xl.to)}<${p}pic><${p}nvPicPr><${p}cNvPr id="${id}" name="${esc(im.name || `Picture ${id}`)}"${im.alt ? ` descr="${esc(im.alt)}"` : ''}/><${p}cNvPicPr><a:picLocks noChangeAspect="1"/></${p}cNvPicPr></${p}nvPicPr><${p}blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></${p}blipFill><${p}spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(im.anchor.w * EMU)}" cy="${Math.round(im.anchor.h * EMU)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></${p}spPr></${p}pic><${p}clientData/></${p}twoCellAnchor>`;
}
const b64bytes = (b64) => { if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(b64, 'base64')); const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };

/**
 * Pictures and shapes of one sheet into its drawing: new ones added, moved/resized/edited originals patched in
 * place, deleted originals removed, untouched originals byte-identical. Returns the updated content types.
 */
export async function writeDrawingObjects(zip, sh, sheetPath, ct, orig) {
  const imgs = sh.images || []; const shapes = sh.shapes || [];
  const origIds = (orig && orig.ids) || new Set(); // cNvPr ids we read (pictures + shapes)
  const live = new Set([...imgs.filter((i) => i._orig && i._orig.cid).map((i) => i._orig.cid), ...shapes.filter((s) => s._orig).map((s) => s._orig.cid)]);
  const fresh = [...imgs.filter((i) => !i._orig), ...shapes.filter((s) => !s._orig)];
  const changed = [...imgs.filter((i) => i._orig && i._orig.cid && i._sig !== picSig(i)), ...shapes.filter((s) => s._orig && s._sig !== shapeSig(s))];
  const removed = [...origIds].filter((id) => !live.has(id));
  if (!fresh.length && !changed.length && !removed.length) return ct;
  let sx = await str(zip, sheetPath); if (!sx) return ct;
  const sRelsPath = relsPathOf(sheetPath);
  let sRels = await str(zip, sRelsPath);
  let drawingPath = null;
  const dm = /<(\w+:)?drawing\s[^>]*?(\w+:)?id="([^"]+)"/.exec(sx);
  if (dm) { const r = parseRels(sRels, sheetPath).find((x) => x.id === dm[3]); if (r) drawingPath = r.path; }
  let dx;
  if (drawingPath && zip.file(drawingPath)) dx = await str(zip, drawingPath);
  else {
    if (!fresh.length) return ct;
    drawingPath = `xl/drawings/drawing${nextNum(zip, /^xl\/drawings\/drawing(\d+)\.xml$/)}.xml`;
    dx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>';
    const rid = nextRid(sRels);
    sRels = addRel(sRels, rid, REL2.drawing, relTarget(sheetPath, drawingPath));
    zip.file(sRelsPath, sRels);
    if (!/<worksheet\b[^>]*xmlns:r=/.test(sx)) sx = sx.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    const after = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
    let at = -1; for (const tg of after) { const i = sx.search(new RegExp(`<${tg}[\\s/>]`)); if (i >= 0 && (at < 0 || i < at)) at = i; }
    if (at < 0) at = sx.lastIndexOf('</worksheet>');
    sx = `${sx.slice(0, at)}<drawing r:id="${rid}"/>${sx.slice(at)}`;
    zip.file(sheetPath, sx);
    ct = ctOverride(ct, drawingPath, CT2.drawing);
  }
  const root = /<(?:(\w+):)?wsDr\b/.exec(dx); const p = root && root[1] ? `${root[1]}:` : '';
  if (!/xmlns:a=/.test(dx.slice(0, 600))) dx = dx.replace(/<((?:\w+:)?wsDr)\b/, '<$1 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"');
  const dRelsPath = relsPathOf(drawingPath);
  let dRels = await str(zip, dRelsPath);
  // patch / remove originals
  const byCid = new Map(changed.map((o) => [o._orig.cid, o]));
  dx = dx.replace(anchorRe, (an) => {
    const cid = (/cNvPr\b[^>]*\bid="(\d+)"/.exec(an) || [])[1];
    if (!cid) return an;
    if (removed.includes(cid) && origIds.has(cid)) return '';
    const o = byCid.get(cid); if (!o) return an;
    const xl = xlAnchorOf(sh, o.anchor);
    if (o.kind) return shapeXml(p, o, cid, xl); // shape: regenerate
    // picture: new position/size, everything else (effects, crop) kept
    let x = an.replace(/^<((?:\w+:)?)(oneCellAnchor|absoluteAnchor)\b([^>]*)>/, `<$1twoCellAnchor editAs="oneCell">`).replace(/<\/((?:\w+:)?)(oneCellAnchor|absoluteAnchor)>$/, '</$1twoCellAnchor>');
    x = x.replace(new RegExp(`<${p}pos\\b[^>]*/>`), '').replace(new RegExp(`<${p}ext\\s+cx="\\d+"\\s+cy="\\d+"\\s*/>`), '');
    x = x.replace(new RegExp(`<${p}from>[\\s\\S]*?</${p}from>(\\s*<${p}to>[\\s\\S]*?</${p}to>)?`), ptXml(p, 'from', xl.from) + ptXml(p, 'to', xl.to));
    if (!new RegExp(`<${p}from>`).test(x)) x = x.replace(/^(<[^>]+>)/, `$1${ptXml(p, 'from', xl.from)}${ptXml(p, 'to', xl.to)}`);
    x = x.replace(/(<a:ext\s+)cx="\d+"\s+cy="\d+"/, `$1cx="${Math.round(o.anchor.w * EMU)}" cy="${Math.round(o.anchor.h * EMU)}"`);
    return x;
  });
  // new objects
  let nextId = 1; for (const m of dx.matchAll(/cNvPr\b[^>]*\bid="(\d+)"/g)) nextId = Math.max(nextId, +m[1] + 1);
  let add = '';
  for (const o of fresh) {
    const xl = xlAnchorOf(sh, o.anchor);
    const id = nextId++;
    if (o.kind) { add += shapeXml(p, o, id, xl); continue; }
    const m = /^data:image\/([\w+]+);base64,(.*)$/.exec(o.src || ''); if (!m) continue;
    const ext = m[1] === 'jpeg' ? 'jpeg' : m[1] === 'svg+xml' ? 'svg' : m[1];
    const media = `xl/media/image${nextNum(zip, /^xl\/media\/image(\d+)\.\w+$/)}.${ext}`;
    zip.file(media, b64bytes(m[2]));
    ct = ctDefault(ct, ext, IMG_CT[ext] || `image/${ext}`);
    const rid = nextRid(dRels);
    dRels = addRel(dRels, rid, REL2.image, relTarget(drawingPath, media));
    add += picXml(p, o, id, rid, xl);
  }
  if (add) {
    if (new RegExp(`</${p}wsDr>`).test(dx)) dx = dx.replace(new RegExp(`</${p}wsDr>\\s*$`), `${add}</${p}wsDr>`);
    else dx = dx.replace(new RegExp(`<${p}wsDr([^>]*)/>`), `<${p}wsDr$1>${add}</${p}wsDr>`);
  }
  zip.file(drawingPath, dx);
  if (dRels) zip.file(dRelsPath, dRels);
  return ct;
}

/* ---------------- tables ---------------- */
export async function readTables(zip, srels, relTable) {
  const out = [];
  for (const r of srels.filter((x) => x.type === relTable)) {
    const tx = await str(zip, r.path); if (!tx) continue;
    const t = readTableXml(tx, r.path); if (!t) continue;
    const g = parseRange(t.refText); if (!g) continue;
    t.range = { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }; t.rid = r.id; delete t.refText;
    out.push(t);
  }
  return out;
}
export const tableSig = (t, cols) => JSON.stringify({ n: t.name, g: t.range, h: t.header, t: t.totals, s: t.style, st: t.stripes, cs: t.colStripes, f: t.firstCol, l: t.lastCol, fn: t.totalsFn, c: cols });
/** write the tables of a sheet; returns { rels, ct, tableParts } (tableParts = XML or null, undefined = unchanged) */
export async function writeTables(zip, book, sh, sheetPath, rels, ct, state) {
  const orig = sh._origTables || [];
  const list = sh.tables || [];
  let touched = false;
  for (const o of orig) if (!list.some((t) => t.path === o.path)) { zip.remove(o.path); ct = ctRemove(ct, o.path); rels = removeRels(rels, (a) => a.Id === o.rid); touched = true; }
  for (const t of list) {
    const cols = book.tableColumns(sh, t);
    if (t.path && zip.file(t.path)) {
      if (t._sig === tableSig(t, cols)) continue;
      const tx = await str(zip, t.path);
      zip.file(t.path, patchTableXml(tx, book, sh, t));
      continue;
    }
    const n = nextNum(zip, /^xl\/tables\/table(\d+)\.xml$/);
    const path = `xl/tables/table${n}.xml`;
    zip.file(path, tableXml(book, sh, t, ++state.maxTableId));
    ct = ctOverride(ct, path, CT2.table);
    const rid = nextRid(rels);
    rels = addRel(rels, rid, REL2.table, relTarget(sheetPath, path));
    t.path = path; t.rid = rid; t._sig = tableSig(t, cols);
    touched = true;
  }
  if (!touched) return { rels, ct, tableParts: undefined };
  const ids = list.map((t) => t.rid).filter(Boolean);
  return { rels, ct, tableParts: ids.length ? `<tableParts count="${ids.length}">${ids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>` : null };
}
export async function maxTableId(zip) { let m = 0; for (const p of Object.keys(zip.files)) if (/^xl\/tables\/table\d+\.xml$/.test(p)) { const x = await str(zip, p); const q = x && /<table\b[^>]*?\sid="(\d+)"/.exec(x); if (q) m = Math.max(m, +q[1]); } return m; }

/* ---------------- pivots ---------------- */
/** read the pivot tables of every sheet (call after all sheets are read) */
export async function readPivots(zip, book, info, wbRels, wbXml) {
  const caches = new Map(); // cacheId → path
  const pc = first(wbXml, 'pivotCaches');
  if (pc) for (const e of elements(pc.inner, 'pivotCache')) { const a = attrs(e.a); const rel = wbRels.find((r) => r.id === a['r:id']); if (rel) caches.set(+a.cacheId, rel.path); }
  for (const sh of book.sheets) {
    const m = info.get(sh.id); if (!m) continue;
    for (const r of m.rels.filter((x) => x.type === REL2.pivotTable)) {
      const tx = await str(zip, r.path); if (!tx) continue;
      const cid = +((/cacheId="(\d+)"/.exec(tx) || [])[1]);
      const cachePath = caches.get(cid);
      const cx = cachePath ? await str(zip, cachePath) : null;
      const pv = readPivot(tx, cx, { path: r.path, cachePath, sheetName: sh.name });
      pv.rid = r.id;
      pv._sig = pivotSig(pv);
      sh.pivots.push(pv);
    }
  }
}
/**
 * Pivot tables: new ones get cache + records + table parts and a workbook pivotCache entry; edited/refreshed originals
 * are rebuilt in their parts (cache extLst kept – slicers point into it); deleted ones removed.
 * Returns { ct, wbRelsXml, pivotCaches (XML | undefined) }.
 */
export async function writePivots(zip, book, entries, { ct, wbRelsXml, wbXml, wbPath }) {
  const pcEl = first(wbXml, 'pivotCaches');
  let cacheList = pcEl ? elements(pcEl.inner, 'pivotCache').map((e) => { const a = attrs(e.a); return { id: +a.cacheId, rid: a['r:id'] }; }) : [];
  let maxCache = Math.max(0, ...cacheList.map((c) => c.id));
  let touched = false;
  const labels = book.pivotLabels;
  for (const { sh, m } of entries) {
    let rels = await str(zip, relsPathOf(m.path));
    const origs = sh._origPivots || [];
    for (const o of origs) if (!(sh.pivots || []).some((p) => p.path === o.path)) {
      zip.remove(o.path); ct = ctRemove(ct, o.path); rels = removeRels(rels, (a) => a.Id === o.rid);
      const rp = relsPathOf(o.path); if (zip.file(rp)) zip.remove(rp);
      if (o.cachePath && !book.sheets.some((s) => (s.pivots || []).some((p) => p.cachePath === o.cachePath))) {
        const recRel = parseRels(await str(zip, relsPathOf(o.cachePath)), o.cachePath).find((x) => x.type === REL2.pivotRecords);
        if (recRel) { zip.remove(recRel.path); ct = ctRemove(ct, recRel.path); }
        zip.remove(o.cachePath); ct = ctRemove(ct, o.cachePath); if (zip.file(relsPathOf(o.cachePath))) zip.remove(relsPathOf(o.cachePath));
        const ce = cacheList.find((c) => c.id === o.cacheId);
        if (ce) { wbRelsXml = removeRels(wbRelsXml, (a) => a.Id === ce.rid); cacheList = cacheList.filter((c) => c !== ce); }
      }
      touched = true;
    }
    for (const pv of sh.pivots || []) {
      if (pv.readOnly) continue;
      if (pv.path && pv._sig === pivotSig(pv) && !pv._dirty && !pv._moved) continue;
      const res = computePivot(book, pv, labels);
      if (res.error) continue;
      const srcSheet = res.sheet || res.src.sheet;
      const srcRef = rangeA1(res.src.g);
      const top = pv.at.r + res.loc.rowOff;
      const location = rangeA1({ r1: top, c1: pv.at.c, r2: top + res.loc.h - 1, c2: pv.at.c + res.loc.w - 1 });
      if (pv.path && pv.cachePath && zip.file(pv.cachePath)) {
        const old = await str(zip, pv.cachePath);
        const keepExt = (/<extLst>[\s\S]*<\/extLst>/.exec(old) || [''])[0];
        const cRels = parseRels(await str(zip, relsPathOf(pv.cachePath)), pv.cachePath);
        let rec = cRels.find((x) => x.type === REL2.pivotRecords);
        let recRid = rec && rec.id;
        if (!rec) {
          const rp = `xl/pivotCache/pivotCacheRecords${nextNum(zip, /^xl\/pivotCache\/pivotCacheRecords(\d+)\.xml$/)}.xml`;
          recRid = 'rId1'; zip.file(relsPathOf(pv.cachePath), addRel(null, recRid, REL2.pivotRecords, relTarget(pv.cachePath, rp)));
          ct = ctOverride(ct, rp, CT2.pivotRecords); rec = { path: rp };
        }
        const parts = pivotParts(pv, res, { cacheId: pv.cacheId, srcSheetName: srcSheet.name, srcRef, location, keepExt, recordsRid: recRid });
        zip.file(pv.cachePath, parts.cacheDef); zip.file(rec.path, parts.records); zip.file(pv.path, parts.table);
      } else {
        const cid = ++maxCache;
        const n = nextNum(zip, /^xl\/pivotCache\/pivotCacheDefinition(\d+)\.xml$/);
        const cachePath = `xl/pivotCache/pivotCacheDefinition${n}.xml`; const recPath = `xl/pivotCache/pivotCacheRecords${n}.xml`;
        const tPath = `xl/pivotTables/pivotTable${nextNum(zip, /^xl\/pivotTables\/pivotTable(\d+)\.xml$/)}.xml`;
        const parts = pivotParts(pv, res, { cacheId: cid, srcSheetName: srcSheet.name, srcRef, location, recordsRid: 'rId1' });
        zip.file(cachePath, parts.cacheDef); zip.file(recPath, parts.records); zip.file(tPath, parts.table);
        zip.file(relsPathOf(cachePath), addRel(null, 'rId1', REL2.pivotRecords, relTarget(cachePath, recPath)));
        zip.file(relsPathOf(tPath), addRel(null, 'rId1', REL2.pivotCache, relTarget(tPath, cachePath)));
        ct = ctOverride(ct, cachePath, CT2.pivotCache); ct = ctOverride(ct, recPath, CT2.pivotRecords); ct = ctOverride(ct, tPath, CT2.pivotTable);
        const wrid = nextRid(wbRelsXml);
        wbRelsXml = addRel(wbRelsXml, wrid, REL2.pivotCache, relTarget(wbPath, cachePath));
        cacheList.push({ id: cid, rid: wrid });
        const srid = nextRid(rels);
        rels = addRel(rels, srid, REL2.pivotTable, relTarget(m.path, tPath));
        Object.assign(pv, { path: tPath, cachePath, cacheId: cid, rid: srid });
      }
      pv._sig = pivotSig(pv); pv._dirty = false; pv._moved = false;
      touched = true;
    }
    if (rels) zip.file(relsPathOf(m.path), rels);
  }
  return { ct, wbRelsXml, pivotCaches: touched ? (cacheList.length ? `<pivotCaches>${cacheList.map((c) => `<pivotCache cacheId="${c.id}" r:id="${c.rid}"/>`).join('')}</pivotCaches>` : null) : undefined };
}

/* ---------------- slicers (shown, kept) ---------------- */
export async function readSlicerInfo(zip, book) {
  // slicer part (caption, cache name) → slicer cache (sourceName = field, pivot cache id or table)
  const caches = new Map();
  for (const p of Object.keys(zip.files)) {
    if (!/^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(p)) continue;
    const x = await str(zip, p); if (!x) continue;
    const a = attrs((/<slicerCacheDefinition\b([^>]*)>/.exec(x) || [0, ''])[1]);
    const sel = [...x.matchAll(/<i\b[^>]*\sx="(\d+)"([^>]*)\/>/g)].map((m) => ({ x: +m[1], s: /\ss="1"/.test(m[2]) }));
    const tbl = /<x15:tableSlicerCache\b[^>]*tableId="(\d+)"[^>]*column="(\d+)"/.exec(x);
    caches.set(a.name, { field: unesc(a.sourceName || ''), sel, pivotCacheId: +((/pivotCacheId="(\d+)"/.exec(x) || [])[1] || -1), tableId: tbl ? +tbl[1] : null, column: tbl ? +tbl[2] : null });
  }
  const slicers = new Map();
  for (const p of Object.keys(zip.files)) {
    if (!/^xl\/slicers\/slicer\d+\.xml$/.test(p)) continue;
    const x = await str(zip, p); if (!x) continue;
    for (const m of x.matchAll(/<slicer\b([^>]*)\/?>/g)) { const a = attrs(m[1]); slicers.set(unesc(a.name), { caption: unesc(a.caption || a.name), cache: caches.get(a.cache) || null, style: a.style || 'SlicerStyleLight1' }); }
  }
  for (const sh of book.sheets) for (const s of sh.slicers || []) Object.assign(s, slicers.get(s.name) || { caption: s.name });
}

/* ---------------- Hostwerk sidecar part ---------------- */
const SIDECAR_NS = 'urn:hostwerk:office:1';
export async function readSidecar(zip) {
  for (const p of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(p)) continue;
    const x = await str(zip, p);
    if (x && x.includes(SIDECAR_NS)) { const m = /<!\[CDATA\[([\s\S]*?)\]\]>/.exec(x); try { return { path: p, data: JSON.parse(m ? m[1] : '{}') }; } catch { return { path: p, data: {} }; } }
  }
  return null;
}
/** write (or remove) the sidecar: data = { sheets: { name: { view, live, db, form } }, v: 1 } */
export async function writeSidecar(zip, data, { ct, wbRelsXml, wbPath, path }) {
  const empty = !data || !Object.keys(data.sheets || {}).length;
  if (empty) {
    if (path && zip.file(path)) {
      const rp = relsPathOf(path); const pr = parseRels(await str(zip, rp), path).find((x) => x.type === REL2.customXmlProps);
      if (pr) { zip.remove(pr.path); ct = ctRemove(ct, pr.path); }
      zip.remove(path); if (zip.file(rp)) zip.remove(rp);
      wbRelsXml = removeRels(wbRelsXml, (a) => a.Type === REL2.customXml && relTarget(wbPath, path) === a.Target);
    }
    return { ct, wbRelsXml };
  }
  const json = JSON.stringify(data).replace(/]]>/g, ']]]]><![CDATA[>');
  const body = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<hw:office xmlns:hw="${SIDECAR_NS}"><![CDATA[${json}]]></hw:office>`;
  if (path && zip.file(path)) { zip.file(path, body); return { ct, wbRelsXml }; }
  const n = nextNum(zip, /^customXml\/item(\d+)\.xml$/);
  const item = `customXml/item${n}.xml`; const props = `customXml/itemProps${n}.xml`;
  const guid = '{7B8E2C3A-4F1D-4E6B-9A2C-5D3F1E0B8A71}';
  zip.file(item, body);
  zip.file(props, `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="${guid}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${SIDECAR_NS}"/></ds:schemaRefs></ds:datastoreItem>`);
  zip.file(relsPathOf(item), addRel(null, 'rId1', REL2.customXmlProps, `itemProps${n}.xml`));
  ct = ctOverride(ct, props, CT2.customXmlProps);
  wbRelsXml = addRel(wbRelsXml, nextRid(wbRelsXml), REL2.customXml, relTarget(wbPath, item));
  return { ct, wbRelsXml, path: item };
}

/* ---------------- calcPr ---------------- */
export function readCalcPr(a) { return { on: a.iterate === '1' || a.iterate === 'true', count: a.iterateCount ? +a.iterateCount : 100, delta: a.iterateDelta ? +a.iterateDelta : 0.001 }; }
export function calcPrAttrs(tag, it) {
  let x = tag.replace(/\s(iterate|iterateCount|iterateDelta)="[^"]*"/g, '');
  if (it && it.on) x = x.replace(/<((?:\w+:)?calcPr)\b/, `<$1 iterate="1"${it.count && it.count !== 100 ? ` iterateCount="${it.count}"` : ''}${it.delta != null && it.delta !== 0.001 ? ` iterateDelta="${it.delta}"` : ''}`);
  return x;
}
export { K };
