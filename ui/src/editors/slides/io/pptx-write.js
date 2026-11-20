/*
 * PPTX writer of the slides chunk – own OOXML writer (jszip), browser + Node.
 *   writePptx(doc, { original }) → Uint8Array
 * With `original` (bytes of the imported pptx) and an imported theme the original theme/master/layouts and unknown
 * parts are kept; only presentation.xml, slides, notes, charts, content types and rels are regenerated.
 */
import JSZip from 'jszip';
import { parseXml, kid, kids, tag, esc, escAttr, XML_HEAD, serialize } from '../../_shared/xmltree.js';
import { parseColorRef, resolveHex } from '../themes.js';
import { SHAPES } from '../render/geometry.js';
import { custGeomXml } from '../render/custom.js';
import { chartWorkbook, chartRefs } from './xlsx-mini.js';
import { timingXml } from './pptx-anim.js';
import { walk } from '../model.js';
import { elementSig, animSig, sigOf } from './pptx-read.js';
import { dataUrlToBytes, EXT_OF_MIME, hashStr, extOfPath } from './util.js';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.';
const CTYPES = {
  slide: `${CT}presentationml.slide+xml`, layout: `${CT}presentationml.slideLayout+xml`, master: `${CT}presentationml.slideMaster+xml`,
  notes: `${CT}presentationml.notesSlide+xml`, notesMaster: `${CT}presentationml.notesMaster+xml`, theme: `${CT}theme+xml`, chart: `${CT}drawingml.chart+xml`,
  pres: `${CT}presentationml.presentation.main+xml`, presProps: `${CT}presentationml.presProps+xml`, viewProps: `${CT}presentationml.viewProps+xml`,
  tableStyles: `${CT}presentationml.tableStyles+xml`, core: 'application/vnd.openxmlformats-package.core-properties+xml', app: `${CT}extended-properties+xml`,
};
const NSDECL = { 'xmlns:a': NS_A, 'xmlns:r': NS_R, 'xmlns:p': NS_P };
const SCHEME = new Set(['tx1', 'bg1', 'tx2', 'bg2', 'dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink', 'phClr']);
const TABLE_STYLE = ['{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}', '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}', '{00A15C55-8517-42AA-B614-E9B94910E393}', '{7DF18680-E054-41AD-8BC1-D1AEF772440D}', '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}', '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}'];
export const TABLE_STYLE_IDS = TABLE_STYLE;
const KIND_TO_TYPE = { title: 'title', titleContent: 'obj', section: 'secHead', twoContent: 'twoObj', comparison: 'twoTxTwoObj', titleOnly: 'titleOnly', blank: 'blank', pictureCaption: 'picTx', custom: 'cust' };

const I = (n) => Math.round(Number(n) || 0);
const guid = () => `{${'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => { const r = (Math.random() * 16) | 0; return (c === 'x' ? r : (r & 3) | 8).toString(16); }).toUpperCase()}}`;

/* ---------------------------------------------------------------- colours, fills, lines */
export function colorXml(ref, theme) {
  const p = parseColorRef(ref);
  if (!p) return '<a:srgbClr val="000000"/>';
  const mods = [];
  let base = p.base;
  if (base[0] === '#') {
    if (base.length === 9) mods.push(`<a:alpha val="${Math.round((parseInt(base.slice(7, 9), 16) / 255) * p.alpha) * 1000}"/>`);
    if (p.mod) base = resolveHex(ref, theme) || base;
  } else if (!SCHEME.has(base)) base = resolveHex(ref, theme) || '#000000';
  if (base[0] !== '#' && p.mod > 0) mods.push(`<a:lumMod val="${(100 - p.mod) * 1000}"/><a:lumOff val="${p.mod * 1000}"/>`);
  else if (base[0] !== '#' && p.mod < 0) mods.push(`<a:lumMod val="${(100 + p.mod) * 1000}"/>`);
  if (p.alpha < 100 && !(p.base[0] === '#' && p.base.length === 9)) mods.push(`<a:alpha val="${p.alpha * 1000}"/>`);
  const inner = mods.join('');
  if (base[0] === '#') return tag('a:srgbClr', { val: base.slice(1, 7).toUpperCase() }, inner || null);
  return tag('a:schemeClr', { val: base }, inner || null);
}
const solid = (c, theme) => `<a:solidFill>${colorXml(c, theme)}</a:solidFill>`;

function fillXml(f, ctx) {
  if (f === undefined) return '';
  if (!f) return '<a:noFill/>';
  if (f.type === 'solid') return solid(f.color, ctx.theme);
  if (f.type === 'grad') {
    const stops = (f.stops || []).map((s) => `<a:gs pos="${I((s.pos || 0) * 1000)}">${colorXml(s.color, ctx.theme)}</a:gs>`).join('');
    const shade = f.kind === 'radial' ? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' : `<a:lin ang="${I((((f.angle || 0) % 360) + 360) % 360 * 60000)}" scaled="0"/>`;
    return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst>${shade}</a:gradFill>`;
  }
  if (f.type === 'image' && f.src) {
    const rid = ctx.part.media(f.src);
    if (!rid) return '<a:noFill/>';
    return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}"/>${f.mode === 'tile' ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>' : '<a:stretch><a:fillRect/></a:stretch>'}</a:blipFill>`;
  }
  return '<a:noFill/>';
}
const DASH = { solid: 'solid', dash: 'dash', dot: 'sysDot', dashDot: 'dashDot', lgDash: 'lgDash', sysDash: 'sysDash', sysDot: 'sysDot' };
const END = { triangle: 'triangle', arrow: 'arrow', stealth: 'stealth', oval: 'oval', diamond: 'diamond' };
function lineXml(l, ctx, tagName = 'a:ln') {
  if (l === undefined) return '';
  if (!l) return `<${tagName}><a:noFill/></${tagName}>`;
  const inner = [solid(l.color || 'tx1', ctx.theme), `<a:prstDash val="${DASH[l.dash] || 'solid'}"/>`];
  if (l.join === 'round') inner.push('<a:round/>');
  if (l.head && END[l.head]) inner.push(`<a:headEnd type="${END[l.head]}"/>`);
  if (l.tail && END[l.tail]) inner.push(`<a:tailEnd type="${END[l.tail]}"/>`);
  return tag(tagName, { w: I(l.width || 12700), cap: l.cap === 'round' ? 'rnd' : l.cap === 'square' ? 'sq' : undefined }, inner);
}
function shadowXml(s, ctx) {
  if (!s) return '';
  const p = parseColorRef(s.color || '#000000');
  const c = colorXml(`${p ? p.base : '#000000'}${p && p.mod ? (p.mod > 0 ? `+${p.mod}` : p.mod) : ''}@${Math.round(s.alpha == null ? 40 : s.alpha)}`, ctx.theme);
  return `<a:effectLst><a:outerShdw blurRad="${I(s.blur ?? 50800)}" dist="${I(s.dist ?? 38100)}" dir="${I((s.dir ?? 45) * 60000)}" algn="tl" rotWithShape="0">${c}</a:outerShdw></a:effectLst>`;
}

/* ---------------------------------------------------------------- text */
const fontTok = (f) => (f === '+mj' ? '+mj-lt' : f === '+mn' ? '+mn-lt' : f);
function rPrInner(r, ctx) {
  const out = [];
  if (r.color) out.push(solid(r.color, ctx.theme));
  if (r.highlight) out.push(`<a:highlight>${colorXml(r.highlight, ctx.theme)}</a:highlight>`);
  if (r.font) { const f = fontTok(r.font); out.push(`<a:latin typeface="${escAttr(f)}"/>`, `<a:ea typeface="${escAttr(f.replace('-lt', '-ea'))}"/>`, `<a:cs typeface="${escAttr(f.replace('-lt', '-cs'))}"/>`); }
  if (r.link && ctx.part) { const rid = ctx.part.rel('hyperlink', r.link, true); out.push(`<a:hlinkClick r:id="${rid}"/>`); }
  return out;
}
function rPrAttrs(r) {
  return {
    lang: r.lang, sz: r.size ? I(r.size * 100) : undefined, b: r.b === undefined ? undefined : r.b ? 1 : 0, i: r.i === undefined ? undefined : r.i ? 1 : 0,
    u: r.u === undefined ? undefined : r.u ? (typeof r.u === 'string' ? r.u : 'sng') : 'none', strike: r.s === undefined ? undefined : r.s ? 'sngStrike' : 'noStrike',
    cap: r.caps ? 'all' : undefined, baseline: r.baseline ? I(r.baseline * 1000) : undefined, dirty: 0,
  };
}
const rPr = (r, ctx, name = 'a:rPr') => tag(name, rPrAttrs(r), rPrInner(r, ctx));

function bulletXml(b, ctx) {
  if (b === undefined) return '';
  if (b === false || b === null) return '<a:buNone/>';
  const out = [];
  if (b.color) out.push(`<a:buClr>${colorXml(b.color, ctx.theme)}</a:buClr>`);
  if (b.sizePct) out.push(`<a:buSzPct val="${I(b.sizePct * 1000)}"/>`);
  if (b.num) { out.push(`<a:buFont typeface="+mj-lt"/>`, tag('a:buAutoNum', { type: b.num, startAt: b.start && b.start !== 1 ? b.start : undefined })); return out.join(''); }
  out.push(`<a:buFont typeface="${escAttr(b.font || 'Arial')}"/>`, `<a:buChar char="${escAttr(b.char || '•')}"/>`);
  return out.join('');
}
const ALIGN = { l: 'l', ctr: 'ctr', r: 'r', just: 'just', c: 'ctr', center: 'ctr' };
function pPrInner(p, ctx) {
  const out = [];
  if (p.lnSpc) out.push(`<a:lnSpc><a:spcPct val="${I(p.lnSpc * 1000)}"/></a:lnSpc>`);
  if (p.spcBef != null) out.push(`<a:spcBef><a:spcPts val="${I(p.spcBef * 100)}"/></a:spcBef>`);
  if (p.spcAft != null) out.push(`<a:spcAft><a:spcPts val="${I(p.spcAft * 100)}"/></a:spcAft>`);
  out.push(bulletXml(p.bullet, ctx));
  return out;
}
/** level style → <a:lvlNpPr> */
function lvlXml(l, i, ctx) {
  if (!l) return '';
  const attrs = { marL: l.marL != null ? I(l.marL) : undefined, indent: l.indent != null ? I(l.indent) : undefined, algn: l.align ? ALIGN[l.align] : undefined };
  const inner = pPrInner(l, ctx);
  inner.push(tag('a:defRPr', rPrAttrs({ size: l.size, b: l.b, i: l.i, u: l.u }), rPrInner({ color: l.color, font: l.font }, { ...ctx, part: null })));
  return tag(`a:lvl${i + 1}pPr`, attrs, inner);
}
const lstStyle = (lvl, ctx) => (lvl && lvl.length ? `<a:lstStyle>${lvl.slice(0, 9).map((l, i) => lvlXml(l, i, ctx)).join('')}</a:lstStyle>` : '<a:lstStyle/>');

// soft line breaks inside a run: \n, vertical tab, U+2028
const BREAK_RE = new RegExp('[\\n\\x0B\\u2028]');
function runsXml(runs, ctx) {
  const out = [];
  for (const r of runs || []) {
    if (r.live && r.live.id) {
      const q = new URLSearchParams({ id: r.live.id, ...(r.live.params ? { params: JSON.stringify(r.live.params) } : {}), ...(r.live.label ? { label: r.live.label } : {}) }).toString();
      out.push(`<a:fld id="${guid()}" type="${escAttr(`hostwerk.live?${q}`)}">${rPr(r, ctx)}<a:t>${esc(r.text || '')}</a:t></a:fld>`);
      continue;
    }
    if (r.field) {
      const type = r.field === 'slidenum' ? 'slidenum' : r.field === 'datetime' ? 'datetime1' : null;
      if (type) { out.push(`<a:fld id="${guid()}" type="${type}">${rPr(r, ctx)}<a:t>${esc(r.field === 'slidenum' ? (ctx.slideNo || r.text || '‹#›') : r.text || '')}</a:t></a:fld>`); continue; }
    }
    const parts = String(r.text ?? '').split(BREAK_RE);
    parts.forEach((t, k) => {
      if (k) out.push(tag('a:br', null, rPr(r, ctx)));
      if (t !== '' || parts.length === 1) out.push(`<a:r>${rPr(r, ctx)}<a:t>${esc(t)}</a:t></a:r>`);
    });
  }
  return out.join('');
}
function paraXml(p, ctx) {
  const attrs = { lvl: p.level ? p.level : undefined, marL: p.marL != null ? I(p.marL) : undefined, indent: p.indent != null ? I(p.indent) : undefined, algn: p.align ? ALIGN[p.align] : undefined };
  const inner = pPrInner(p, ctx).join('');
  const ppr = Object.values(attrs).some((v) => v !== undefined) || inner ? tag('a:pPr', attrs, inner || null) : '';
  const end = p.endProps ? rPr(p.endProps, ctx, 'a:endParaRPr') : '<a:endParaRPr dirty="0"/>';
  return `<a:p>${ppr}${runsXml(p.runs, ctx)}${end}</a:p>`;
}
function bodyPrXml(b) {
  const ins = b.insets || {};
  const a = { wrap: b.wrap === false ? 'none' : 'square', lIns: I(ins.l ?? 91440), tIns: I(ins.t ?? 45720), rIns: I(ins.r ?? 91440), bIns: I(ins.b ?? 45720), anchor: b.anchor || 't', vert: b.vert, rtlCol: 0 };
  const fit = b.autofit === 'shrink' ? tag('a:normAutofit', { fontScale: b.fontScale && b.fontScale < 1 ? I(b.fontScale * 100000) : undefined }) : b.autofit === 'resize' ? '<a:spAutoFit/>' : '<a:noAutofit/>';
  return tag('a:bodyPr', a, fit);
}
function txBodyXml(b, ctx, name = 'p:txBody') {
  if (!b) return tag(name, null, ['<a:bodyPr/>', '<a:lstStyle/>', '<a:p><a:endParaRPr dirty="0"/></a:p>']);
  const paras = (b.paras && b.paras.length ? b.paras : [{ runs: [] }]).map((p) => paraXml(p, ctx)).join('');
  return tag(name, null, [bodyPrXml(b), lstStyle(b.lvl, ctx), paras]);
}

/* ---------------------------------------------------------------- elements */
function xfrm(e, name = 'a:xfrm', extra = '') {
  return tag(name, { rot: e.rot ? I((((e.rot % 360) + 360) % 360) * 60000) : undefined, flipH: e.flipH ? 1 : undefined, flipV: e.flipV ? 1 : undefined }, `<a:off x="${I(e.x)}" y="${I(e.y)}"/><a:ext cx="${Math.max(0, I(e.w))}" cy="${Math.max(0, I(e.h))}"/>${extra}`);
}
function geomXml(geom, adj) {
  const def = SHAPES[geom];
  const prst = def ? geom : 'rect';
  const defaults = (def && def.adj) || [];
  const gd = (adj || []).map((v, i) => (v == null ? '' : `<a:gd name="${defaults.length > 1 ? `adj${i + 1}` : 'adj'}" fmla="val ${I(v)}"/>`)).join('');
  return `<a:prstGeom prst="${prst}"><a:avLst>${gd}</a:avLst></a:prstGeom>`;
}
function cNvPr(e, ctx, id) {
  const inner = e.link ? `<a:hlinkClick r:id="${ctx.part.rel('hyperlink', e.link, true)}"/>` : null;
  return tag('p:cNvPr', { id, name: e.name || `${e.type} ${id}`, descr: e.alt || undefined, hidden: e.hidden ? 1 : undefined }, inner);
}
/**
 * placeholder type/idx for a ph key.
 * OOXML / PowerPoint rules: a slide MASTER may only hold typed placeholders (title, body, dt, ftr, sldNum – an
 * untyped <p:ph idx="1"/> means "obj", which PowerPoint rejects on a master with "file is corrupt"); masters use the
 * classic idx 1–4. Layouts may hold untyped content (obj) placeholders; slides inherit from the layout by idx.
 */
const MASTER_PH = { title: { type: 'title' }, ctrTitle: { type: 'title' }, body: { type: 'body', idx: 1 }, dt: { type: 'dt', idx: 2 }, ftr: { type: 'ftr', idx: 3 }, sldNum: { type: 'sldNum', idx: 4 } };
export function phAttrs(key, def, layoutKind) {
  if (layoutKind === 'master') return { ...(MASTER_PH[key] || MASTER_PH.body) };
  if (def && (def.type || def.idx != null)) return { type: def.type && def.type !== 'obj' ? def.type : undefined, idx: def.idx != null ? def.idx : undefined };
  const objKinds = ['titleContent', 'twoContent'];
  switch (key) {
    case 'title': return { type: 'title' };
    case 'ctrTitle': return { type: 'ctrTitle' };
    case 'subTitle': return { type: 'subTitle', idx: 1 };
    case 'pic': return { type: 'pic', idx: 5 };
    case 'dt': return { type: 'dt', idx: 10 };
    case 'ftr': return { type: 'ftr', idx: 11 };
    case 'sldNum': return { type: 'sldNum', idx: 12 };
    default: {
      const n = key === 'body' ? 1 : parseInt(key.replace(/\D/g, ''), 10) || 1;
      return { type: objKinds.includes(layoutKind) && !(layoutKind === 'comparison') ? undefined : 'body', idx: n };
    }
  }
}

/* ---------------------------------------------------------------- preserved (raw) elements */
const XF_SAME = (e, r) => r && I(e.x) === I(r.x) && I(e.y) === I(r.y) && I(e.w) === I(r.w) && I(e.h) === I(r.h) && Math.abs((e.rot || 0) - (r.rot || 0)) < 0.01 && !!e.flipH === !!r.flipH && !!e.flipV === !!r.flipV;
function setXfrm(x, e, ns) {
  x.a = { ...x.a };
  for (const k of ['rot', 'flipH', 'flipV']) delete x.a[k];
  if (e.rot) x.a.rot = String(I((((e.rot % 360) + 360) % 360) * 60000));
  if (e.flipH) x.a.flipH = '1';
  if (e.flipV) x.a.flipV = '1';
  const keep = (x.c || []).filter((c) => typeof c !== 'string' && /:(chOff|chExt)$/.test(c.n));
  x.c = [{ n: `${ns}:off`, a: { x: String(I(e.x)), y: String(I(e.y)) }, c: [] }, { n: `${ns}:ext`, a: { cx: String(Math.max(0, I(e.w))), cy: String(Math.max(0, I(e.h))) }, c: [] }, ...keep];
}
const lname = (n) => n.slice(n.indexOf(':') + 1);
const kidL = (el, name) => (el.c || []).find((c) => typeof c !== 'string' && lname(c.n) === name);
/** patch own xfrm (and cNvPr id) of a parsed raw element node; AlternateContent: every branch */
function patchRaw(node, e, id, moved) {
  const t = lname(node.n);
  if (t === 'AlternateContent') { for (const br of node.c || []) if (typeof br !== 'string') for (const c of br.c || []) if (typeof c !== 'string') patchRaw(c, e, id, moved); return; }
  const nv = (node.c || []).find((c) => typeof c !== 'string' && /^nv/.test(lname(c.n)));
  const cnv = nv && kidL(nv, 'cNvPr');
  if (cnv && id != null) cnv.a = { ...cnv.a, id: String(id) };
  if (!moved) return;
  const prKey = t === 'grpSp' ? 'grpSpPr' : 'spPr';
  if (t === 'graphicFrame') { let x = kidL(node, 'xfrm'); if (!x) { x = { n: 'p:xfrm', a: {}, c: [] }; node.c.splice(node.c.indexOf(nv) + 1, 0, x); } setXfrm(x, e, 'a'); return; }
  const pr = kidL(node, prKey);
  if (!pr) return;
  let x = kidL(pr, 'xfrm');
  if (!x) { x = { n: 'a:xfrm', a: {}, c: [] }; pr.c.unshift(x); }
  setXfrm(x, e, 'a');
}
function rawElXml(e, ctx) {
  const placed = ctx.raw.place(e.raw, e.rawRels);
  if (!placed) return null;
  const node = parseXml(placed);
  if (!node) return null;
  const id = ctx.assign(e);
  walk(e.children || [], (c) => { if (c.spid) ctx.usedSpids.add(c.spid); });
  patchRaw(node, e, id !== e.spid ? id : null, !XF_SAME(e, e.rawXfrm));
  Object.assign(ctx.ns, e.rawNs || {});
  return serialize(node);
}

function elXml(e, ctx) {
  try {
    if (e.raw && ctx.raw && ctx.raw.usable(e) && (e.preserved || elementSig(e) === e.rawSig)) {
      const x = rawElXml(e, ctx);
      if (x) return x;
    }
    const id = ctx.assign(e);
    switch (e.type) {
      case 'group': {
        const kidsXml = (e.children || []).map((c) => elXml(c, ctx)).join('');
        return `<p:grpSp><p:nvGrpSpPr>${cNvPr(e, ctx, id)}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrm(e, 'a:xfrm', `<a:chOff x="${I(e.x)}" y="${I(e.y)}"/><a:chExt cx="${I(e.w)}" cy="${I(e.h)}"/>`)}</p:grpSpPr>${kidsXml}</p:grpSp>`;
      }
      case 'line': {
        const prst = e.geom === 'elbow' ? 'bentConnector3' : e.geom === 'curve' ? 'curvedConnector3' : 'straightConnector1';
        return `<p:cxnSp><p:nvCxnSpPr>${cNvPr(e, ctx, id)}<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm(e)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${lineXml(e.line || { color: 'tx1', width: 12700 }, ctx)}${shadowXml(e.shadow, ctx)}</p:spPr></p:cxnSp>`;
      }
      case 'image': case 'unknown': {
        if (!e.src) return shapeXml({ ...e, type: 'shape', geom: 'rect', fill: { type: 'solid', color: 'bg2' }, line: { color: 'tx2', width: 9525 }, text: { paras: [{ runs: [{ text: e.what || '' }], align: 'ctr' }], anchor: 'ctr' } }, ctx, id);
        const rid = ctx.part.media(e.src);
        const c = e.crop || {};
        const src = c.l || c.t || c.r || c.b ? tag('a:srcRect', { l: c.l ? I(c.l * 100000) : undefined, t: c.t ? I(c.t * 100000) : undefined, r: c.r ? I(c.r * 100000) : undefined, b: c.b ? I(c.b * 100000) : undefined }) : '';
        const alpha = e.opacity != null && e.opacity < 1 ? `<a:alphaModFix amt="${I(e.opacity * 100000)}"/>` : '';
        const ph = e.ph ? tag('p:ph', phAttrs(e.ph, ctx.layoutPh(e.ph), ctx.layoutKind)) : '';
        return `<p:pic><p:nvPicPr>${cNvPr(e, ctx, id)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr>${ph}</p:nvPr></p:nvPicPr><p:blipFill>${tag('a:blip', { 'r:embed': rid }, alpha || null)}${src}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(e)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${e.line ? lineXml(e.line, ctx) : ''}${shadowXml(e.shadow, ctx)}</p:spPr></p:pic>`;
      }
      case 'media': {
        const m = dataUrlToBytes(e.src);
        if (!m) return '';
        const target = ctx.addMediaFile(e.src);
        const r1 = ctx.part.rel(e.kind === 'audio' ? 'audio' : 'video', `../media/${target}`);
        const r2 = ctx.part.rel('http://schemas.microsoft.com/office/2007/relationships/media', `../media/${target}`);
        const poster = ctx.part.media(e.poster || POSTER_PNG);
        ctx.media.push({ spid: id, kind: e.kind });
        return `<p:pic><p:nvPicPr>${cNvPr(e, ctx, id)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:${e.kind === 'audio' ? 'audioFile' : 'videoFile'} r:link="${r1}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="${NS_P14}" r:embed="${r2}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${poster}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(e)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      }
      case 'table': return tableXml(e, ctx, id);
      case 'chart': return chartFrameXml(e, ctx, id);
      default: return shapeXml(e, ctx, id);
    }
  } catch (err) {
    return '';
  }
}

function shapeXml(e, ctx, id) {
  const ph = e.ph ? tag('p:ph', phAttrs(e.ph, ctx.layoutPh(e.ph), ctx.layoutKind)) : '';
  const nv = `<p:nvSpPr>${cNvPr(e, ctx, id)}${e.textbox && !e.ph ? '<p:cNvSpPr txBox="1"/>' : e.ph ? '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' : '<p:cNvSpPr/>'}<p:nvPr>${ph}</p:nvPr></p:nvSpPr>`;
  const sp = `<p:spPr>${xfrm(e)}${e.geom === 'custom' && e.paths && e.paths.length ? custGeomXml(e.paths) : geomXml(e.geom || 'rect', e.adj)}${fillXml(e.fill === undefined ? null : e.fill, ctx)}${lineXml(e.line === undefined ? null : e.line, ctx)}${shadowXml(e.shadow, ctx)}</p:spPr>`;
  return `<p:sp>${nv}${sp}${txBodyXml(e.text, ctx)}</p:sp>`;
}

function tableXml(e, ctx, id) {
  const st = e.style || {};
  const accent = /^accent([1-6])$/.exec(st.accent || 'accent1');
  const styleId = TABLE_STYLE[accent ? accent[1] - 1 : 0];
  const rows = e.rows || [];
  // origin of covered cells
  const origin = rows.map((r) => r.cells.map(() => null));
  rows.forEach((r, ri) => r.cells.forEach((c, ci) => {
    if (c.merged) return;
    const sp = c.span || { c: 1, r: 1 };
    for (let y = ri; y < ri + (sp.r || 1) && y < rows.length; y++) for (let x = ci; x < ci + (sp.c || 1) && x < r.cells.length; x++) if (y !== ri || x !== ci) origin[y][x] = [ri, ci];
  }));
  const border = e.border;
  const tr = rows.map((r, ri) => {
    const tcs = r.cells.map((c, ci) => {
      const o = origin[ri][ci];
      const sp = c.span || {};
      const attrs = o ? { hMerge: ci > o[1] ? 1 : undefined, vMerge: ri > o[0] ? 1 : undefined } : { gridSpan: sp.c > 1 ? sp.c : undefined, rowSpan: sp.r > 1 ? sp.r : undefined };
      const b = c.text || {};
      const ins = b.insets || {};
      const bd = c.borders || {};
      const ln = ['l', 'r', 't', 'b'].map((k) => { const l = bd[k] !== undefined ? bd[k] : border; return l === undefined ? '' : lineXml(l, ctx, `a:ln${k.toUpperCase()}`); }).join('');
      const tcPr = tag('a:tcPr', { marL: I(ins.l ?? 91440), marR: I(ins.r ?? 91440), marT: I(ins.t ?? 45720), marB: I(ins.b ?? 45720), anchor: b.anchor || undefined }, [ln, c.fill === undefined ? '' : fillXml(c.fill, ctx)]);
      return tag('a:tc', attrs, [txBodyXml(b.paras ? b : null, ctx, 'a:txBody'), tcPr]);
    }).join('');
    return `<a:tr h="${I(r.h)}">${tcs}</a:tr>`;
  }).join('');
  const tblPr = tag('a:tblPr', { firstRow: st.firstRow ? 1 : undefined, bandRow: st.bandRows ? 1 : undefined, firstCol: st.firstCol ? 1 : undefined, lastRow: st.lastRow ? 1 : undefined }, `<a:tableStyleId>${styleId}</a:tableStyleId>`);
  return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(e, ctx, id)}<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>${xfrm(e, 'p:xfrm')}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>${tblPr}<a:tblGrid>${(e.cols || []).map((w) => `<a:gridCol w="${I(w)}"/>`).join('')}</a:tblGrid>${tr}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

/* ---------------------------------------------------------------- charts */
function chartFrameXml(e, ctx, id) {
  // free part names: a patched package still holds the charts of unchanged (raw) chart frames
  const path = ctx.pkg ? ctx.pkg.free((k) => `ppt/charts/chart${k}.xml`) : `ppt/charts/chart${++ctx.chartNo.n}.xml`;
  // the embedded workbook – PowerPoint's "Edit data" opens it in Excel
  const wb = ctx.pkg ? ctx.pkg.free((k) => `ppt/embeddings/Microsoft_Excel_Worksheet${k}.xlsx`) : null;
  if (wb) { ctx.files[wb] = chartWorkbook(e.chart || {}); if (ctx.pkg.defaults && !ctx.pkg.defaults.xlsx) ctx.pkg.defaults.xlsx = `${CT}spreadsheetml.sheet`; }
  ctx.files[path] = XML_HEAD + chartPartXml(e.chart || {}, ctx, !!wb);
  ctx.types[`/${path}`] = CTYPES.chart;
  if (wb) ctx.files[relsPathOf(path)] = relsXml([{ id: 'rId1', type: `${RT}package`, target: `../embeddings/${basename(wb)}` }]);
  const rid = ctx.part.rel('chart', relTarget(ctx.part.path, path));
  return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(e, ctx, id)}<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${xfrm(e, 'p:xfrm')}<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}
const txt = (s) => `<c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(s)}</a:t></a:r></a:p></c:rich>`;
function chartPartXml(ch, ctx, withData = false) {
  const kind = ch.kind || 'col';
  const cats = ch.categories || [];
  const strLit = (arr) => `<c:strLit><c:ptCount val="${arr.length}"/>${arr.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strLit>`;
  const numLit = (arr) => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${arr.length}"/>${arr.map((v, i) => (v == null || v === '' || Number.isNaN(Number(v)) ? '' : `<c:pt idx="${i}"><c:v>${Number(v)}</c:v></c:pt>`)).join('')}</c:numLit>`;
  const dl = ch.labels ? '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : '';
  const pie = kind === 'pie' || kind === 'doughnut';
  const ser = (ch.series || []).map((s, i) => {
    const color = s.color || `accent${(i % 6) + 1}`;
    const sp = kind === 'line' ? `<c:spPr><a:ln w="28575" cap="rnd">${solid(color, ctx.theme)}<a:round/></a:ln></c:spPr>` : `<c:spPr>${solid(color, ctx.theme)}</c:spPr>`;
    const dpt = pie ? cats.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr>${solid(`accent${(k % 6) + 1}`, ctx.theme)}<a:ln w="19050">${solid('bg1', ctx.theme)}</a:ln></c:spPr></c:dPt>`).join('') : '';
    const ref = chartRefs(i, cats.length);
    const nm = s.name || `Series ${i + 1}`;
    const tx = withData ? `<c:tx><c:strRef><c:f>${ref.name}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(nm)}</c:v></c:pt></c:strCache></c:strRef></c:tx>` : `<c:tx><c:v>${esc(nm)}</c:v></c:tx>`;
    const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}`;
    const cache = (lit, tagN) => lit.replace(new RegExp(`^<c:${tagN}Lit>`), `<c:${tagN}Cache>`).replace(new RegExp(`</c:${tagN}Lit>$`), `</c:${tagN}Cache>`);
    const data = withData
      ? `<c:cat><c:strRef><c:f>${ref.cat}</c:f>${cache(strLit(cats.map(String)), 'str')}</c:strRef></c:cat><c:val><c:numRef><c:f>${ref.val}</c:f>${cache(numLit(s.values || []), 'num')}</c:numRef></c:val>`
      : `<c:cat>${strLit(cats.map(String))}</c:cat><c:val>${numLit(s.values || [])}</c:val>`;
    if (kind === 'line') return `<c:ser>${head}<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>${dl}${data}<c:smooth val="0"/></c:ser>`;
    if (pie) return `<c:ser>${head}${dpt}${dl}${data}</c:ser>`;
    if (kind === 'area') return `<c:ser>${head}${dl}${data}</c:ser>`;
    return `<c:ser>${head}<c:invertIfNegative val="0"/>${dl}${data}</c:ser>`;
  }).join('');
  const grouping = ch.stacked ? 'stacked' : kind === 'line' || kind === 'area' ? 'standard' : 'clustered';
  const ax = '<c:axId val="50010001"/><c:axId val="50010002"/>';
  let plot;
  if (kind === 'pie') plot = `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
  else if (kind === 'doughnut') plot = `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
  else if (kind === 'line') plot = `<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${ser}<c:marker val="1"/>${ax}</c:lineChart>`;
  else if (kind === 'area') plot = `<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${ser}${ax}</c:areaChart>`;
  else plot = `<c:barChart><c:barDir val="${kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${ser}<c:gapWidth val="150"/>${ch.stacked ? '<c:overlap val="100"/>' : ''}${ax}</c:barChart>`;
  const axes = pie ? '' : `<c:catAx><c:axId val="50010001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${kind === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="50010002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${kind === 'bar' ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525">${solid('tx1+85', ctx.theme)}</a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const title = ch.title ? `<c:title><c:tx>${txt(ch.title)}</c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = ch.legend !== false ? `<c:legend><c:legendPos val="${pie ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend>` : '';
  return `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200">${solid('tx1+25', ctx.theme)}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>${withData ? '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>' : ''}</c:chartSpace>`;
}

/* ---------------------------------------------------------------- transitions & timing */
const SPD = (ms) => (ms <= 500 ? 'fast' : ms <= 750 ? 'med' : 'slow');
function transitionXml(t) {
  if (!t || (!t.type || t.type === 'none') && !t.advanceAfter) return '';
  const dir = ['l', 'r', 'u', 'd'].includes(t.dir) ? t.dir : undefined;
  const child = { fade: '<p:fade/>', push: tag('p:push', { dir }), wipe: tag('p:wipe', { dir }), cover: tag('p:cover', { dir }), zoom: '<p:zoom/>', split: '<p:split orient="horz" dir="out"/>', morph: '<p:fade/>' }[t.type] || '';
  const adv = t.advanceAfter ? I(t.advanceAfter) : undefined;
  const dur = I(t.duration || 500);
  if (t.type === 'morph') {
    // Morph (PowerPoint 2019/365) – older versions fade
    return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">${tag('p:transition', { spd: 'slow', 'p14:dur': dur, advTm: adv, 'xmlns:p14': NS_P14 }, `<p159:morph option="${t.option === 'word' ? 'byWord' : t.option === 'char' ? 'byChar' : 'byObject'}"/>`)}</mc:Choice><mc:Fallback>${tag('p:transition', { spd: 'slow', advTm: adv }, '<p:fade/>')}</mc:Fallback></mc:AlternateContent>`;
  }
  return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="${NS_P14}" Requires="p14">${tag('p:transition', { spd: SPD(dur), 'p14:dur': dur, advTm: adv }, child || null)}</mc:Choice><mc:Fallback>${tag('p:transition', { spd: SPD(dur), advTm: adv }, child || null)}</mc:Fallback></mc:AlternateContent>`;
}
/* ---------------------------------------------------------------- boilerplate parts */
// 1×1 grey PNG as video poster when none is given
const POSTER_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const NV_GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const relsXml = (rels) => `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((r) => tag('Relationship', { Id: r.id, Type: r.type, Target: r.target, TargetMode: r.external ? 'External' : undefined })).join('')}</Relationships>`;

function themeXml(theme, name) {
  const c = theme.colors || {};
  const clr = (k) => `<a:${k}><a:srgbClr val="${String(c[k] || '#000000').replace('#', '').slice(0, 6).toUpperCase()}"/></a:${k}>`;
  const font = (f) => `<a:latin typeface="${escAttr(f)}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  const fonts = theme.fonts || {};
  const ph = (inner) => `<a:schemeClr val="phClr">${inner}</a:schemeClr>`;
  return `${XML_HEAD}<a:theme xmlns:a="${NS_A}" name="${escAttr(name || theme.name || 'Theme')}"><a:themeElements><a:clrScheme name="${escAttr(theme.name || 'Custom')}">${['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(clr).join('')}</a:clrScheme><a:fontScheme name="${escAttr(theme.name || 'Custom')}"><a:majorFont>${font(fonts.major || 'Calibri Light')}</a:majorFont><a:minorFont>${font(fonts.minor || 'Calibri')}</a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill>${ph('')}</a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${ph('<a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/>')}</a:gs><a:gs pos="50000">${ph('<a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/>')}</a:gs><a:gs pos="100000">${ph('<a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/>')}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${ph('<a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/>')}</a:gs><a:gs pos="50000">${ph('<a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/>')}</a:gs><a:gs pos="100000">${ph('<a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/>')}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph('')}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph('')}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph('')}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill>${ph('')}</a:solidFill><a:solidFill>${ph('<a:tint val="95000"/><a:satMod val="170000"/>')}</a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${ph('<a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/>')}</a:gs><a:gs pos="50000">${ph('<a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/>')}</a:gs><a:gs pos="100000">${ph('<a:shade val="63000"/><a:satMod val="120000"/>')}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}
const LVL_TAGS = (fn) => Array.from({ length: 9 }, (_, i) => fn(i)).join('');
const DEFAULT_TEXT_STYLE = `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${LVL_TAGS((i) => `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`)}</p:defaultTextStyle>`;
function clrMapAttrs(theme) {
  const m = theme.dark ? { bg1: 'dk1', tx1: 'lt1', bg2: 'dk2', tx2: 'lt2' } : { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
  Object.assign(m, theme.clrMap || {});
  return { ...m, accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hlink: 'hlink', folHlink: 'folHlink' };
}
function notesMasterXml() {
  const sp = (id, name, type, idx, x, y, w, h, extra = '') => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${tag('p:ph', { type, idx })}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${extra}</p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
  return `${XML_HEAD}<p:notesMaster ${Object.entries(NSDECL).map(([k, v]) => `${k}="${v}"`).join(' ')}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${NV_GRP}${sp(2, 'Slide Image Placeholder 1', 'sldImg', 2, 685800, 1143000, 5486400, 3086100, '<a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>')}${sp(3, 'Notes Placeholder 2', 'body', 1, 685800, 4400550, 5486400, 3600450)}${sp(4, 'Slide Number Placeholder 3', 'sldNum', 5, 3884613, 8685213, 2971800, 458787)}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle>${LVL_TAGS((i) => `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`)}</p:notesStyle></p:notesMaster>`;
}
function notesSlideXml(notes, bodyIdx = 1) {
  const paras = String(notes || '').split(/\r?\n/).map((l) => `<a:p>${l ? `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${esc(l)}</a:t></a:r>` : ''}<a:endParaRPr lang="en-US" dirty="0"/></a:p>`).join('');
  return `${XML_HEAD}<p:notes ${Object.entries(NSDECL).map(([k, v]) => `${k}="${v}"`).join(' ')}><p:cSld><p:spTree>${NV_GRP}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="${bodyIdx}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}
function presPropsXml(doc) {
  const s = doc.show || {};
  return `${XML_HEAD}<p:presentationPr ${Object.entries(NSDECL).map(([k, v]) => `${k}="${v}"`).join(' ')}>${tag('p:showPr', { loop: s.loop ? 1 : undefined, useTimings: s.useTimings === false ? 0 : undefined, showNarration: 1 }, '<p:present/><p:sldAll/><p:penClr><a:prstClr val="red"/></p:penClr>')}</p:presentationPr>`;
}
const VIEW_PROPS = `${XML_HEAD}<p:viewPr ${Object.entries(NSDECL).map(([k, v]) => `${k}="${v}"`).join(' ')}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0"><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`;
const TABLE_STYLES = `${XML_HEAD}<a:tblStyleLst xmlns:a="${NS_A}" def="${TABLE_STYLE[0]}"/>`;
const iso = (d) => { try { return new Date(d || Date.now()).toISOString().replace(/\.\d+Z$/, 'Z'); } catch { return new Date().toISOString().replace(/\.\d+Z$/, 'Z'); } };
function coreXml(doc) {
  const m = doc.meta || {};
  return `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(m.title || '')}</dc:title><dc:creator>${esc(m.author || '')}</dc:creator><cp:lastModifiedBy>${esc(m.author || '')}</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">${iso(m.created)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${iso(m.modified)}</dcterms:modified></cp:coreProperties>`;
}
const appXml = (doc) => `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Application>Hostwerk Office</Application><PresentationFormat>Custom</PresentationFormat><Slides>${doc.slides.length}</Slides><Notes>${doc.slides.filter((s) => s.notes && s.notes.trim()).length}</Notes><HiddenSlides>${doc.slides.filter((s) => s.hidden).length}</HiddenSlides><AppVersion>16.0000</AppVersion></Properties>`;

/* ---------------------------------------------------------------- package building */
const MAIN_TYPES = {
  pptx: CTYPES.pres, pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
  ppsx: `${CT}presentationml.slideshow.main+xml`, ppsm: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
  potx: `${CT}presentationml.template.main+xml`, potm: 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
};
const MACRO_EXT = /^(pptm|ppsm|potm)$/i;
const dirname = (p) => p.slice(0, p.lastIndexOf('/'));
const basename = (p) => p.slice(p.lastIndexOf('/') + 1);
function relPath(fromDir, toDir) {
  const a = fromDir ? fromDir.split('/') : []; const b = toDir ? toDir.split('/') : [];
  let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++;
  return [...a.slice(i).map(() => '..'), ...b.slice(i)].join('/') || '.';
}
const relTarget = (fromPart, toPart) => { const d = relPath(dirname(fromPart), dirname(toPart)); return d === '.' ? basename(toPart) : `${d}/${basename(toPart)}`; };
const normPath = (base, target) => {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base ? base.split('/').slice(0, -1) : [];
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.' && seg !== '') parts.push(seg); }
  return parts.join('/');
};
const relsPathOf = (p) => { const i = p.lastIndexOf('/'); return `${p.slice(0, i)}/_rels/${p.slice(i + 1)}.rels`; };
const asStr = (v) => (typeof v === 'string' ? v : new TextDecoder().decode(v));

function makePart(pkg, path) {
  const rels = [];
  const part = {
    path, rels,
    rel(type, target, external = false) {
      const t = type.includes('://') ? type : RT + type;
      const ex = rels.find((r) => r.type === t && r.target === target);
      if (ex) return ex.id;
      let k = rels.length + 1;
      while (rels.some((r) => r.id === `rId${k}`)) k++;
      const id = `rId${k}`;
      rels.push({ id, type: t, target, external });
      return id;
    },
    media(src) {
      const file = pkg.addMediaFile(src);
      if (!file) return null;
      return part.rel('image', relTarget(path, `ppt/media/${file}`));
    },
  };
  return part;
}

function newPkg(existing = new Set()) {
  const pkg = {
    files: {}, types: {}, orig: {}, mediaByHash: {}, mediaNo: 0, chartNo: { n: 0 }, existing, defaults: {},
    free(make) { let n = 0; let p; do { p = make(++n); } while (pkg.existing.has(p) || pkg.files[p] !== undefined); return p; },
    addMediaFile(src) {
      if (!src) return null;
      const h = hashStr(src) + src.length.toString(36) + src.slice(-24);
      if (pkg.mediaByHash[h]) return pkg.mediaByHash[h];
      const m = dataUrlToBytes(src);
      if (!m) return null;
      const ext = EXT_OF_MIME[m.mime] || (m.mime.split('/')[1] || 'bin').replace(/[^a-z0-9]/g, '');
      const kind = /^(video|audio)/.test(m.mime) ? 'media' : 'image';
      const p = pkg.free((n) => `ppt/media/${kind}${n}.${ext}`);
      pkg.files[p] = m.bytes;
      if (!pkg.defaults[ext]) pkg.defaults[ext] = m.mime;
      pkg.mediaByHash[h] = basename(p);
      return basename(p);
    },
  };
  return pkg;
}

/* ---- comments (legacy p:cmLst + commentAuthors.xml) */
function authorsOf(pkg) {
  if (pkg.authors) return pkg.authors;
  const a = { list: [], dirty: false, path: 'ppt/commentAuthors.xml' };
  const f = pkg.files[a.path];
  if (f !== undefined) {
    try { for (const x of kids(parseXml(asStr(f)), 'p:cmAuthor')) a.list.push({ id: parseInt(x.a.id, 10) || 0, name: x.a.name || '', initials: x.a.initials || '', lastIdx: parseInt(x.a.lastIdx, 10) || 0, clrIdx: parseInt(x.a.clrIdx, 10) || 0 }); } catch { /* fresh */ }
  }
  pkg.authors = a;
  return a;
}
function writeComments(pkg, part, comments) {
  const A = authorsOf(pkg);
  // threads: a reply (parent = id of a top comment) follows its thread; PowerPoint 2013+ links it with p15:threadingInfo
  const tops = comments.filter((c) => !c.parent || !comments.some((x) => x.id === c.parent));
  const ordered = [];
  for (const c of tops) { ordered.push(c); for (const r of comments) if (r.parent === c.id && r !== c) ordered.push(r); }
  const key = {};
  const items = ordered.map((c, i) => {
    const name = c.author || 'Author';
    let au = A.list.find((x) => x.name === name);
    if (!au) { au = { id: A.list.reduce((m, x) => Math.max(m, x.id + 1), 0), name, initials: name.split(/\s+/).map((w) => w[0] || '').join('').slice(0, 3).toUpperCase(), lastIdx: 0, clrIdx: A.list.length % 8 }; A.list.push(au); }
    au.lastIdx += 1;
    A.dirty = true;
    key[c.id] = { authorId: au.id, idx: au.lastIdx };
    const par = c.parent && key[c.parent];
    const top = par ? ordered.find((x) => x.id === c.parent) : c;
    const pos = (top && top.pos) || { x: (10 + i * 10) * 1587.5, y: (10 + i * 10) * 1587.5 };
    const thread = par ? `<p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" timeZoneBias="0"><p15:parentCm authorId="${par.authorId}" idx="${par.idx}"/></p15:threadingInfo></p:ext></p:extLst>` : '';
    return `<p:cm authorId="${au.id}" dt="${iso(c.time || Date.now()).replace('Z', '.000')}" idx="${au.lastIdx}"><p:pos x="${I(pos.x / 1587.5)}" y="${I(pos.y / 1587.5)}"/><p:text>${esc(c.text || '')}</p:text>${thread}</p:cm>`;
  });
  const p = pkg.free((n) => `ppt/comments/comment${n}.xml`);
  pkg.files[p] = `${XML_HEAD}<p:cmLst ${Object.entries(NSDECL).map(([k, v]) => `${k}="${v}"`).join(' ')}>${items.join('')}</p:cmLst>`;
  pkg.types[`/${p}`] = `${CT}presentationml.comments+xml`;
  part.rel('comments', relTarget(part.path, p));
}
function finishAuthors(pkg, presRels) {
  const A = pkg.authors;
  if (!A || !A.dirty) return;
  pkg.files[A.path] = `${XML_HEAD}<p:cmAuthorLst ${Object.entries(NSDECL).map(([k, v]) => `${k}="${v}"`).join(' ')}>${A.list.map((a) => tag('p:cmAuthor', { id: a.id, name: a.name, initials: a.initials, lastIdx: a.lastIdx, clrIdx: a.clrIdx })).join('')}</p:cmAuthorLst>`;
  pkg.types[`/${A.path}`] = `${CT}presentationml.commentAuthors+xml`;
  if (!presRels.some((r) => /commentAuthors$/.test(r.type))) presRels.push({ id: `rId${presRels.length + 1}x`, type: `${RT}commentAuthors`, target: 'commentAuthors.xml' });
}

/* ---- one slide (+ notes, comments) – shared by both writers */
function writeSlide(pkg, doc, slide, n, o) {
  const path = o.slidePath || `ppt/slides/slide${n}.xml`;
  const part = makePart(pkg, path);
  part.rel('slideLayout', o.layoutTarget);
  const src = o.src;
  const layout = o.layout;
  const usable = (x) => !!(src && x && x.rawDoc && x.rawDoc === doc.sourceId);
  // ids: imported elements keep their original shape ids (timing XML refers to them)
  const reserved = new Set();
  walk(slide.elements, (e) => { if (e.spid) reserved.add(e.spid); });
  if (usable(slide)) for (const x of slide.timingSpids || []) reserved.add(x);
  const usedSpids = new Set();
  let counter = 1;
  const fresh = () => { do { counter++; } while (reserved.has(counter) || usedSpids.has(counter)); usedSpids.add(counter); return counter; };
  const ctx = {
    theme: doc.theme, part, spids: {}, media: [], files: pkg.files, types: pkg.types, chartNo: pkg.chartNo, slideNo: n, pkg, ns: {}, usedSpids,
    nextId: fresh,
    assign: (e) => { let id; if (e.spid && !usedSpids.has(e.spid)) { id = e.spid; usedSpids.add(id); } else id = fresh(); ctx.spids[e.id] = id; return id; },
    addMediaFile: (s) => pkg.addMediaFile(s), layoutKind: layout ? layout.kind : 'custom', layoutPh: (k) => (layout && layout.ph ? layout.ph[k] : null),
    raw: src ? {
      usable,
      place(raw, rawRels) {
        const map = {};
        for (const r of rawRels || []) {
          if (r.external) { map[r.id] = part.rel(r.type, r.target, true); continue; }
          let t = r.target;
          if (/^ppt\/slides\/[^/]+\.xml$/.test(t)) { t = o.slideMap[t]; if (!t) return null; }
          else if (pkg.files[t] === undefined) return null;
          map[r.id] = part.rel(r.type, relTarget(path, t));
        }
        return raw.replace(/(\sr:[A-Za-z]+=")([^"]*)"/g, (m, a, v) => (map[v] ? `${a}${map[v]}"` : m));
      },
    } : null,
  };
  const textIds = new Set();
  walk(slide.elements, (e) => { if (e.type === 'shape' && e.text && e.text.paras && e.text.paras.some((p) => (p.runs || []).length)) textIds.add(e.id); });
  let tree = (slide.elements || []).map((e) => elXml(e, ctx)).join('');
  // footer placeholders from the presentation's footer settings
  const f = doc.footer || {};
  const isTitle = layout && layout.kind === 'title';
  if (layout && layout.ph && !(isTitle && f.hideOnTitle)) {
    const fp = (key, text, field) => {
      const d = layout.ph[key];
      if (!d) return '';
      return shapeXml({ id: `f-${key}`, type: 'shape', geom: 'rect', name: key, x: d.x, y: d.y, w: d.w, h: d.h, ph: key, fill: null, line: null, textbox: true, text: { paras: [{ runs: [{ text, field }], align: (d.lvl && d.lvl[0] && d.lvl[0].align) || undefined }], anchor: d.anchor || 'ctr', lvl: d.lvl } }, ctx, ctx.nextId());
    };
    if (f.showDate) tree += fp('dt', f.date || new Date().toLocaleDateString(), f.date ? undefined : 'datetime');
    if (f.showFooter && f.text) tree += fp('ftr', f.text);
    if (f.showNumber) tree += fp('sldNum', String(n), 'slidenum');
  }
  const sUse = usable(slide);
  const placeS = (x) => (x && sUse ? ctx.raw.place(x, slide.rawRels) : null);
  let bg = '';
  if (sUse && slide.bgRaw && sigOf(slide.background) === slide.bgSig) bg = placeS(slide.bgRaw) || '';
  if (!bg && slide.background) bg = `<p:bg><p:bgPr>${fillXml(slide.background, ctx)}<a:effectLst/></p:bgPr></p:bg>`;
  let trans = '';
  if (sUse && slide.transitionRaw && sigOf(slide.transition) === slide.transitionSig) trans = placeS(slide.transitionRaw) || '';
  if (!trans) trans = transitionXml(slide.transition);
  let timing = '';
  if (sUse && slide.timingRaw) {
    const spidOf = {};
    walk(slide.elements, (e) => { if (e.spid) spidOf[e.id] = e.spid; });
    if (animSig(slide.animations, spidOf) === slide.animSig && (slide.timingSpids || []).every((x) => usedSpids.has(x))) timing = placeS(slide.timingRaw) || '';
  }
  if (!timing) timing = timingXml(slide.animations, ctx.spids, textIds);
  const ex = sUse && slide.rawExtra ? slide.rawExtra : {};
  const cExtra = [ex.custData, ex.controls, ex.cSldExt].map(placeS).filter(Boolean).join('');
  const sExt = placeS(ex.sldExt) || '';
  if (sUse) Object.assign(ctx.ns, slide.rawNs || {});
  const attrs = { ...NSDECL };
  for (const [k, v] of Object.entries(ctx.ns)) if (!attrs[`xmlns:${k}`]) attrs[`xmlns:${k}`] = v;
  if (sUse && slide.rawIgnorable) { const ign = slide.rawIgnorable.split(/\s+/).filter((p) => attrs[`xmlns:${p}`]); if (ign.length && attrs['xmlns:mc']) attrs['mc:Ignorable'] = ign.join(' '); }
  Object.assign(attrs, { show: slide.hidden ? 0 : undefined, showMasterSp: slide.showMaster === false ? 0 : undefined });
  const name = sUse && slide.name ? ` name="${escAttr(slide.name)}"` : '';
  pkg.files[path] = `${XML_HEAD}${tag('p:sld', attrs, [`<p:cSld${name}>${bg}<p:spTree>${NV_GRP}${tree}</p:spTree>${cExtra}</p:cSld>`, '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>', trans, timing, sExt])}`;
  pkg.types[`/${path}`] = CTYPES.slide;
  // notes: the original notes slide when the text is unchanged, else a fresh one
  if (slide.notes && slide.notes.trim() && o.notesMasterPath) {
    const np = `ppt/notesSlides/notesSlide${n}.xml`;
    const orig = sUse && slide.notesPart && sigOf(slide.notes) === slide.notesSig ? pkg.orig[slide.notesPart] : undefined;
    if (orig !== undefined) {
      pkg.files[np] = orig;
      let rels = [];
      const ro = pkg.orig[relsPathOf(slide.notesPart)];
      if (ro !== undefined) rels = kids(parseXml(asStr(ro)), 'Relationship').map((r) => ({ id: r.a.Id, type: r.a.Type, target: r.a.Target, external: r.a.TargetMode === 'External' }));
      rels = rels.filter((r) => r.external || /\/(slide|notesMaster)$/.test(r.type) || pkg.files[normPath(slide.notesPart, r.target)] !== undefined).map((r) => (/\/slide$/.test(r.type) ? { ...r, target: relTarget(np, path) } : /\/notesMaster$/.test(r.type) ? { ...r, target: relTarget(np, o.notesMasterPath) } : r));
      if (!rels.some((r) => /\/notesMaster$/.test(r.type))) rels.push({ id: 'rIdNm', type: `${RT}notesMaster`, target: relTarget(np, o.notesMasterPath) });
      pkg.files[relsPathOf(np)] = relsXml(rels);
    } else {
      pkg.files[np] = notesSlideXml(slide.notes, o.notesBodyIdx);
      pkg.files[relsPathOf(np)] = relsXml([{ id: 'rId1', type: `${RT}notesMaster`, target: relTarget(np, o.notesMasterPath) }, { id: 'rId2', type: `${RT}slide`, target: relTarget(np, path) }]);
    }
    pkg.types[`/${np}`] = CTYPES.notes;
    part.rel('notesSlide', relTarget(path, np));
  }
  // comments: original part when unchanged
  const comments = slide.comments || [];
  if (sUse && slide.commentsPart && sigOf(comments) === slide.commentsSig && pkg.files[slide.commentsPart.path] !== undefined) {
    part.rel(slide.commentsPart.type, relTarget(path, slide.commentsPart.path));
  } else if (comments.length) writeComments(pkg, part, comments);
  pkg.files[relsPathOf(path)] = relsXml(part.rels);
}

const GUIDES_URI = '{EFAFB233-063F-42B5-8137-9DF3F51BA10A}';
function guidesXml(doc) {
  const g = doc.guides || [];
  if (!g.length) return '';
  return `<p:ext uri="${GUIDES_URI}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main">${g.map((x, i) => `<p15:guide id="${i + 1}"${x.axis === 'y' ? ' orient="horz"' : ''} pos="${I(x.pos / 1587.5)}"><p15:clr><a:srgbClr val="A4A3A4"/></p15:clr></p15:guide>`).join('')}</p15:sldGuideLst></p:ext>`;
}
function sectionsXml(doc, ids) {
  const groups = [];
  const secs = (doc.sections || []).map((s) => ({ ...s, i: doc.slides.findIndex((x) => x.id === s.start) })).filter((s) => s.i >= 0).sort((a, b) => a.i - b.i);
  if (!secs.length) return '';
  if (secs[0].i > 0) groups.push({ name: 'Default Section', from: 0, to: secs[0].i });
  secs.forEach((s, k) => groups.push({ name: s.name || 'Section', from: s.i, to: k + 1 < secs.length ? secs[k + 1].i : doc.slides.length }));
  return `<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="${NS_P14}">${groups.map((g) => `<p14:section name="${escAttr(g.name)}" id="${guid()}"><p14:sldIdLst>${ids.slice(g.from, g.to).map((id) => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst></p14:section>`).join('')}</p14:sectionLst></p:ext>`;
}

function contentTypes(pkg, overrides, origDefaults = {}) {
  const defaults = { rels: 'application/vnd.openxmlformats-package.relationships+xml', xml: 'application/xml', ...origDefaults, ...pkg.defaults };
  for (const p of Object.keys(pkg.files)) {
    const ext = extOfPath(p);
    if (!defaults[ext] && !overrides[`/${p}`]) defaults[ext] = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bin: 'application/vnd.openxmlformats-officedocument.oleObject', xlsx: `${CT}spreadsheetml.sheet`, fntdata: 'application/x-fontdata', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml', vml: `${CT}vmlDrawing`, mp4: 'video/mp4' }[ext] || 'application/octet-stream';
  }
  const ov = Object.entries(overrides).filter(([p]) => pkg.files[p.slice(1)] !== undefined);
  return `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${Object.entries(defaults).map(([e, t]) => `<Default Extension="${escAttr(e)}" ContentType="${escAttr(t)}"/>`).join('')}${ov.map(([p, t]) => `<Override PartName="${escAttr(p)}" ContentType="${escAttr(t)}"/>`).join('')}</Types>`;
}

async function zipUp(pkg) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', pkg.files['[Content_Types].xml']);
  for (const [p, v] of Object.entries(pkg.files)) if (p !== '[Content_Types].xml') zip.file(p, v);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

/** theme + master + layouts from doc.theme/doc.layouts; returns targets */
function genMasters(pkg, doc, sfx = '') {
  const theme = doc.theme || {};
  const layouts = doc.layouts && doc.layouts.length ? doc.layouts : [{ id: 'layout-blank', kind: 'blank', name: 'Blank', ph: {}, decor: [] }];
  const masterPath = `ppt/slideMasters/slideMaster${sfx}1.xml`;
  const themePath = `ppt/theme/theme${sfx}1.xml`;
  const masterPart = makePart(pkg, masterPath);
  const ctl = layouts.find((l) => l.kind === 'titleContent') || layouts[0];
  const seq = () => { let i = 1; return () => ++i; };
  const mctx = { theme, part: masterPart, spids: {}, media: [], nextId: seq(), files: pkg.files, types: pkg.types, chartNo: pkg.chartNo, pkg, ns: {}, addMediaFile: (s) => pkg.addMediaFile(s), layoutKind: ctl.kind, layoutPh: (k) => (ctl.ph || {})[k] };
  mctx.assign = (e) => { const id = mctx.nextId(); mctx.spids[e.id] = id; return id; };
  const phShape = (key, d, ctx, kind) => {
    const id = ctx.nextId();
    const prompt = d.prompt || (key === 'title' ? 'Click to edit Master title style' : '');
    const runs = key === 'sldNum' ? [{ text: '‹#›', field: 'slidenum' }] : prompt ? [{ text: prompt }] : [];
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escAttr(key)} ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${tag('p:ph', phAttrs(key, kind === 'master' ? null : d, kind))}</p:nvPr></p:nvSpPr><p:spPr>${xfrm(d)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>${txBodyXml({ paras: [{ runs }], anchor: d.anchor || 't', insets: d.insets, autofit: 'none', lvl: d.lvl }, { ...ctx, slideNo: '‹#›' })}</p:sp>`;
  };
  const masterPh = ['title', 'body', 'dt', 'ftr', 'sldNum'].map((k) => (ctl.ph && ctl.ph[k] ? phShape(k, ctl.ph[k], mctx, 'master') : '')).join('');
  const layoutIds = [];
  const layoutTargets = [];
  layouts.forEach((l, i) => {
    const path = `ppt/slideLayouts/slideLayout${sfx}${i + 1}.xml`;
    const part = makePart(pkg, path);
    part.rel('slideMaster', relTarget(path, masterPath));
    const ctx = { ...mctx, part, nextId: seq(), layoutKind: l.kind, layoutPh: (k) => (l.ph || {})[k] };
    ctx.assign = (e) => { const id = ctx.nextId(); ctx.spids[e.id] = id; return id; };
    const decor = (l.decor || []).map((e) => elXml({ ...e, ph: undefined, raw: undefined }, ctx)).join('');
    const ph = Object.entries(l.ph || {}).map(([k, d]) => phShape(k, d, ctx, l.kind)).join('');
    const bg = l.bg ? `<p:bg><p:bgPr>${fillXml(l.bg, ctx)}<a:effectLst/></p:bgPr></p:bg>` : '';
    pkg.files[path] = `${XML_HEAD}${tag('p:sldLayout', { ...NSDECL, type: KIND_TO_TYPE[l.kind] || 'cust', preserve: 1 }, [`<p:cSld name="${escAttr(l.name || l.kind)}">${bg}<p:spTree>${NV_GRP}${decor}${ph}</p:spTree></p:cSld>`, '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'])}`;
    pkg.files[relsPathOf(path)] = relsXml(part.rels);
    pkg.types[`/${path}`] = CTYPES.layout;
    layoutIds.push(masterPart.rel('slideLayout', relTarget(masterPath, path)));
    layoutTargets.push(path);
  });
  masterPart.rel('theme', relTarget(masterPath, themePath));
  const styleOf = (lvl) => (lvl || []).map((x, i) => lvlXml(x, i, { theme })).join('');
  const titleLvl = (ctl.ph && ctl.ph.title && ctl.ph.title.lvl) || [{ font: '+mj', size: 40, color: 'tx1', bullet: false }];
  const bodyLvl = (ctl.ph && ctl.ph.body && ctl.ph.body.lvl) || Array.from({ length: 9 }, (_, i) => ({ font: '+mn', size: 24, color: 'tx1', bullet: { char: '•' }, marL: 285750 + i * 457200, indent: -285750 }));
  const otherLvl = Array.from({ length: 9 }, (_, i) => ({ font: '+mn', size: 18, color: 'tx1', marL: i * 457200 }));
  const mbg = theme.bg || { type: 'solid', color: 'bg1' };
  pkg.files[masterPath] = `${XML_HEAD}${tag('p:sldMaster', NSDECL, [`<p:cSld><p:bg><p:bgPr>${fillXml(mbg, mctx)}<a:effectLst/></p:bgPr></p:bg><p:spTree>${NV_GRP}${masterPh}</p:spTree></p:cSld>`, tag('p:clrMap', clrMapAttrs(theme)), `<p:sldLayoutIdLst>${layoutIds.map((r, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="${r}"/>`).join('')}</p:sldLayoutIdLst>`, `<p:txStyles><p:titleStyle>${styleOf(titleLvl.slice(0, 1))}</p:titleStyle><p:bodyStyle>${styleOf(bodyLvl)}</p:bodyStyle><p:otherStyle>${styleOf(otherLvl)}</p:otherStyle></p:txStyles>`])}`;
  pkg.files[relsPathOf(masterPath)] = relsXml(masterPart.rels);
  pkg.types[`/${masterPath}`] = CTYPES.master;
  pkg.files[themePath] = themeXml(theme);
  pkg.types[`/${themePath}`] = CTYPES.theme;
  return { masterPath, themePath, layouts, layoutTargets };
}
function genNotesMaster(pkg, doc, sfx = '') {
  const path = `ppt/notesMasters/notesMaster${sfx}1.xml`;
  const themePath = `ppt/theme/theme${sfx}Notes.xml`;
  pkg.files[path] = notesMasterXml();
  pkg.files[themePath] = themeXml(doc.theme || {}, 'Notes Theme');
  pkg.files[relsPathOf(path)] = relsXml([{ id: 'rId1', type: `${RT}theme`, target: relTarget(path, themePath) }]);
  pkg.types[`/${path}`] = CTYPES.notesMaster;
  pkg.types[`/${themePath}`] = CTYPES.theme;
  return path;
}

/* ---------------------------------------------------------------- full writer (new file) */
function writeFull(doc, ext = 'pptx') {
  const pkg = newPkg();
  const { masterPath, themePath, layouts, layoutTargets } = genMasters(pkg, doc);
  const hasNotes = doc.slides.some((s) => s.notes && s.notes.trim());
  const notesMasterPath = hasNotes ? genNotesMaster(pkg, doc) : null;
  const presRels = [{ id: 'rId1', type: `${RT}slideMaster`, target: relTarget('ppt/presentation.xml', masterPath) }];
  const sldIds = [];
  doc.slides.forEach((s, i) => {
    const li = Math.max(0, layouts.findIndex((l) => l.id === s.layoutId));
    writeSlide(pkg, doc, s, i + 1, { layoutTarget: relTarget(`ppt/slides/slide${i + 1}.xml`, layoutTargets[li]), layout: layouts[li], notesMasterPath, notesBodyIdx: 1 });
    presRels.push({ id: `rId${presRels.length + 1}`, type: `${RT}slide`, target: `slides/slide${i + 1}.xml` });
    sldIds.push(256 + i);
  });
  const slideRids = presRels.slice(1).map((r) => r.id);
  let notesRid = null;
  if (notesMasterPath) { notesRid = `rId${presRels.length + 1}`; presRels.push({ id: notesRid, type: `${RT}notesMaster`, target: relTarget('ppt/presentation.xml', notesMasterPath) }); }
  for (const [t, target] of [['presProps', 'presProps.xml'], ['viewProps', 'viewProps.xml'], ['theme', relTarget('ppt/presentation.xml', themePath)], ['tableStyles', 'tableStyles.xml']]) presRels.push({ id: `rId${presRels.length + 1}`, type: RT + t, target });
  finishAuthors(pkg, presRels);
  presRels.forEach((r, i) => { if (/x$/.test(r.id)) r.id = `rId${presRels.length + 1 + i}`; });
  const sec = sectionsXml(doc, sldIds) + guidesXml(doc);
  pkg.files['ppt/presentation.xml'] = `${XML_HEAD}${tag('p:presentation', { ...NSDECL, saveSubsetFonts: 1 }, [
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>',
    notesRid ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesRid}"/></p:notesMasterIdLst>` : '',
    sldIds.length ? `<p:sldIdLst>${sldIds.map((id, i) => `<p:sldId id="${id}" r:id="${slideRids[i]}"/>`).join('')}</p:sldIdLst>` : '',
    `<p:sldSz cx="${I(doc.size.w)}" cy="${I(doc.size.h)}"/>`, '<p:notesSz cx="6858000" cy="9144000"/>', DEFAULT_TEXT_STYLE,
    sec ? `<p:extLst>${sec}</p:extLst>` : '',
  ])}`;
  pkg.files['ppt/_rels/presentation.xml.rels'] = relsXml(presRels);
  pkg.files['ppt/presProps.xml'] = presPropsXml(doc);
  pkg.files['ppt/viewProps.xml'] = VIEW_PROPS;
  pkg.files['ppt/tableStyles.xml'] = TABLE_STYLES;
  pkg.files['docProps/core.xml'] = coreXml(doc);
  pkg.files['docProps/app.xml'] = appXml(doc);
  pkg.files['_rels/.rels'] = relsXml([{ id: 'rId1', type: `${RT}officeDocument`, target: 'ppt/presentation.xml' }, { id: 'rId2', type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' }, { id: 'rId3', type: `${RT}extended-properties`, target: 'docProps/app.xml' }]);
  Object.assign(pkg.types, { '/ppt/presentation.xml': MAIN_TYPES[String(ext).toLowerCase()] || CTYPES.pres, '/ppt/presProps.xml': CTYPES.presProps, '/ppt/viewProps.xml': CTYPES.viewProps, '/ppt/tableStyles.xml': CTYPES.tableStyles, '/docProps/core.xml': CTYPES.core, '/docProps/app.xml': CTYPES.app });
  pkg.files['[Content_Types].xml'] = contentTypes(pkg, pkg.types);
  return zipUp(pkg);
}

/* ---------------------------------------------------------------- patching writer (imported file) */
async function writePatched(doc, original, ext) {
  const src = await JSZip.loadAsync(original);
  const names = Object.keys(src.files).filter((n) => !src.files[n].dir);
  const pkg = newPkg(new Set(names));
  const macros = MACRO_EXT.test(ext || '');
  const keepMaster = !!(doc.theme && doc.theme.id === 'imported');
  for (const n of names) pkg.orig[n] = await src.file(n).async('uint8array');
  const readRels = (p) => { const f = pkg.orig[p === '' ? '_rels/.rels' : relsPathOf(p)]; if (f === undefined) return []; return kids(parseXml(asStr(f)), 'Relationship').map((r) => ({ id: r.a.Id, type: r.a.Type, target: r.a.Target, external: r.a.TargetMode === 'External' })); };
  const ct = parseXml(asStr(pkg.orig['[Content_Types].xml']));
  const origOverrides = {};
  for (const o of kids(ct, 'Override')) origOverrides[o.a.PartName] = o.a.ContentType;
  const origDefaults = {};
  for (const d of kids(ct, 'Default')) origDefaults[d.a.Extension.toLowerCase()] = d.a.ContentType;
  const rootRels = readRels('');
  const mainRel = rootRels.find((r) => /officeDocument$/.test(r.type));
  const presPath = mainRel ? normPath('', mainRel.target) : 'ppt/presentation.xml';
  if (presPath !== 'ppt/presentation.xml') throw new Error('unusual package layout');
  const presRels = readRels(presPath);
  const isSig = (n) => /^_xmlsignatures\//.test(n);
  const isVba = (n) => /^ppt\/(vbaProject\.bin|vbaData\.xml)$/i.test(n) || /^ppt\/_rels\/vbaProject\.bin\.rels$/i.test(n);
  const masterSet = new Set();
  if (!keepMaster) {
    for (const r of presRels) if (/slideMaster$/.test(r.type)) masterSet.add(normPath(presPath, r.target));
  }
  for (const n of names) {
    if (/^ppt\/(slides|notesSlides)\//.test(n) || n === '[Content_Types].xml' || n === presPath || n === 'ppt/_rels/presentation.xml.rels' || isSig(n)) continue;
    if (isVba(n) && !macros) continue;
    pkg.files[n] = pkg.orig[n];
  }
  // root rels without the digital signature origin
  pkg.files['_rels/.rels'] = relsXml(rootRels.filter((r) => !/digital-signature/.test(r.type)));
  // masters / layouts
  let layoutTargets = null; let layouts = doc.layouts || [];
  let master = null;
  if (!keepMaster) master = genMasters(pkg, doc, 'hw');
  if (master) { layoutTargets = master.layoutTargets; layouts = master.layouts; }
  const origLayouts = names.filter((n) => /^ppt\/slideLayouts\/[^/]+\.xml$/.test(n));
  // notes master: original one if present
  const nmRel = presRels.find((r) => /notesMaster$/.test(r.type));
  let notesMasterPath = null; let notesBodyIdx = 1;
  if (nmRel && pkg.files[normPath(presPath, nmRel.target)] !== undefined) {
    notesMasterPath = normPath(presPath, nmRel.target);
    try { const ph = findAll(parseXml(asStr(pkg.files[notesMasterPath])), 'p:ph').find((p) => p.a.type === 'body'); if (ph && ph.a.idx) notesBodyIdx = ph.a.idx; } catch { /* default */ }
  } else if (doc.slides.some((s) => s.notes && s.notes.trim())) notesMasterPath = genNotesMaster(pkg, doc, 'hw');
  // slides
  const slideMap = {};
  doc.slides.forEach((s, i) => { if (s.srcPart && s.rawDoc === doc.sourceId && !slideMap[s.srcPart]) slideMap[s.srcPart] = `ppt/slides/slide${i + 1}.xml`; });
  const ids = [];
  doc.slides.forEach((s, i) => {
    let layout; let lp;
    if (master) { const li = Math.max(0, layouts.findIndex((l) => l.id === s.layoutId)); layout = layouts[li]; lp = layoutTargets[li]; }
    else { layout = layouts.find((l) => l.id === s.layoutId) || null; lp = layout && layout.part && pkg.files[layout.part] !== undefined ? layout.part : origLayouts[0]; }
    writeSlide(pkg, doc, s, i + 1, { layoutTarget: relTarget(`ppt/slides/slide${i + 1}.xml`, lp), layout, notesMasterPath, notesBodyIdx, src, slideMap });
    ids.push(256 + i);
  });
  // presentation rels
  const DROP = /\/(slide|vbaProject|digital-signature\/origin)$/;
  const keep = presRels.filter((r) => {
    if (/\/slide$/.test(r.type)) return false;
    if (/vbaProject$/.test(r.type)) return macros;
    if (master && (/slideMaster$/.test(r.type) || /\/theme$/.test(r.type))) return false;
    if (/notesMaster$/.test(r.type)) return normPath(presPath, r.target) === notesMasterPath;
    return !DROP.test(r.type);
  });
  const relOut = [];
  const idMap = {};
  let n = 1;
  for (const r of keep) { const id = `rId${n++}`; idMap[r.id] = id; relOut.push({ ...r, id }); }
  let ourMasterRid = null;
  if (master) {
    ourMasterRid = `rId${n++}`; relOut.push({ id: ourMasterRid, type: `${RT}slideMaster`, target: relTarget(presPath, master.masterPath) });
    relOut.push({ id: `rId${n++}`, type: `${RT}theme`, target: relTarget(presPath, master.themePath) });
  }
  if (notesMasterPath && !relOut.some((r) => /notesMaster$/.test(r.type))) relOut.push({ id: `rId${n++}`, type: `${RT}notesMaster`, target: relTarget(presPath, notesMasterPath) });
  const oldSlideRid = {};
  for (const r of presRels) if (/\/slide$/.test(r.type)) oldSlideRid[r.id] = normPath(presPath, r.target);
  const slideRids = doc.slides.map((_, i) => { const id = `rId${n++}`; relOut.push({ id, type: `${RT}slide`, target: `slides/slide${i + 1}.xml` }); return id; });
  const newRidOfPart = {};
  doc.slides.forEach((s, i) => { if (s.srcPart && s.rawDoc === doc.sourceId && !newRidOfPart[s.srcPart]) newRidOfPart[s.srcPart] = slideRids[i]; });
  for (const [old, p] of Object.entries(oldSlideRid)) if (newRidOfPart[p]) idMap[old] = newRidOfPart[p];
  finishAuthors(pkg, relOut);
  relOut.forEach((r) => { if (/x$/.test(r.id)) r.id = `rId${n++}`; });
  // presentation.xml: patch the original
  const pres = parseXml(asStr(pkg.orig[presPath]));
  const remap = (el) => { if (typeof el === 'string') return; for (const k of Object.keys(el.a)) if (/^r:/.test(k) && el.a[k] in idMap) el.a[k] = idMap[el.a[k]]; (el.c || []).forEach(remap); };
  pres.c = pres.c.filter((x) => typeof x === 'string' || !['p:sldIdLst', 'p:notesMasterIdLst'].includes(x.n) && !(master && x.n === 'p:sldMasterIdLst'));
  remap(pres);
  // custom shows: drop references to deleted slides
  const cs = kid(pres, 'p:custShowLst');
  if (cs) for (const show of kids(cs, 'p:custShow')) { const l = kid(show, 'p:sldLst'); if (l) l.c = l.c.filter((x) => typeof x === 'string' || Object.values(newRidOfPart).includes(x.a['r:id'])); }
  const order = ['p:sldMasterIdLst', 'p:notesMasterIdLst', 'p:handoutMasterIdLst', 'p:sldIdLst', 'p:sldSz', 'p:notesSz'];
  const insertAt = (name, el) => {
    const idx = order.indexOf(name);
    let pos = 0;
    pres.c.forEach((x, k) => { if (typeof x !== 'string' && order.indexOf(x.n) >= 0 && order.indexOf(x.n) < idx) pos = k + 1; });
    pres.c.splice(pos, 0, el);
  };
  if (master) insertAt('p:sldMasterIdLst', { n: 'p:sldMasterIdLst', a: {}, c: [{ n: 'p:sldMasterId', a: { id: '2147483648', 'r:id': ourMasterRid }, c: [] }] });
  const nmKept = relOut.find((r) => /notesMaster$/.test(r.type));
  if (nmKept) insertAt('p:notesMasterIdLst', { n: 'p:notesMasterIdLst', a: {}, c: [{ n: 'p:notesMasterId', a: { 'r:id': nmKept.id }, c: [] }] });
  if (ids.length) insertAt('p:sldIdLst', { n: 'p:sldIdLst', a: {}, c: ids.map((id, i) => ({ n: 'p:sldId', a: { id: String(id), 'r:id': slideRids[i] }, c: [] })) });
  const sz = kid(pres, 'p:sldSz');
  if (sz) { sz.a.cx = String(I(doc.size.w)); sz.a.cy = String(I(doc.size.h)); delete sz.a.type; }
  const ext0 = kid(pres, 'p:extLst');
  if (ext0) ext0.c = ext0.c.filter((x) => typeof x === 'string' || (x.a.uri !== '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}' && x.a.uri !== GUIDES_URI));
  const sec = sectionsXml(doc, ids) + guidesXml(doc);
  let presXml = XML_HEAD + serialize(pres);
  if (sec) presXml = ext0 ? presXml.replace(/<\/p:extLst>(?![\s\S]*<\/p:extLst>)/, `${sec}</p:extLst>`) : presXml.replace(/<\/p:presentation>\s*$/, `<p:extLst>${sec}</p:extLst></p:presentation>`);
  pkg.files[presPath] = presXml;
  pkg.files['ppt/_rels/presentation.xml.rels'] = relsXml(relOut);
  // presProps: loop
  if (pkg.files['ppt/presProps.xml'] !== undefined) {
    try {
      const pp = parseXml(asStr(pkg.files['ppt/presProps.xml']));
      let show = kid(pp, 'p:showPr');
      const want = !!(doc.show && doc.show.loop);
      if (!show && want) { show = { n: 'p:showPr', a: {}, c: [{ n: 'p:present', a: {}, c: [] }, { n: 'p:sldAll', a: {}, c: [] }] }; pp.c.unshift(show); }
      if (show && (show.a.loop === '1') !== want) { if (want) show.a.loop = '1'; else delete show.a.loop; pkg.files['ppt/presProps.xml'] = XML_HEAD + serialize(pp); }
    } catch { /* keep */ }
  } else pkg.files['ppt/presProps.xml'] = presPropsXml(doc);
  // docProps/core: patch title + modified only
  if (pkg.files['docProps/core.xml'] !== undefined) {
    let c = asStr(pkg.files['docProps/core.xml']);
    const m = doc.meta || {};
    if (/<dc:title>[\s\S]*?<\/dc:title>|<dc:title\/>/.test(c)) c = c.replace(/<dc:title>[\s\S]*?<\/dc:title>|<dc:title\/>/, `<dc:title>${esc(m.title || '')}</dc:title>`);
    else if (m.title) c = c.replace(/<\/cp:coreProperties>/, `<dc:title>${esc(m.title)}</dc:title></cp:coreProperties>`);
    c = c.replace(/(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${iso(Date.now())}$2`);
    pkg.files['docProps/core.xml'] = c;
  } else pkg.files['docProps/core.xml'] = coreXml(doc);
  // garbage-collect unreachable parts (old masters/layouts, media of deleted slides …)
  const reach = new Set();
  const queue = [''];
  while (queue.length) {
    const p = queue.shift();
    const rp = p === '' ? '_rels/.rels' : relsPathOf(p);
    if (pkg.files[rp] === undefined) continue;
    reach.add(rp);
    for (const r of kids(parseXml(asStr(pkg.files[rp])), 'Relationship')) {
      if (r.a.TargetMode === 'External') continue;
      const t = normPath(p, r.a.Target);
      if (!reach.has(t)) { reach.add(t); queue.push(t); }
    }
  }
  for (const p of Object.keys(pkg.files)) if (!reach.has(p)) delete pkg.files[p];
  // content types
  const types = { ...origOverrides, ...pkg.types, '/docProps/core.xml': CTYPES.core };
  const e = String(ext || '').toLowerCase();
  types[`/${presPath}`] = MAIN_TYPES[e] || (/macroEnabled/.test(origOverrides[`/${presPath}`] || '') && !macros ? CTYPES.pres : origOverrides[`/${presPath}`] || CTYPES.pres);
  pkg.files['[Content_Types].xml'] = contentTypes(pkg, types, origDefaults);
  return zipUp(pkg);
}
function findAll(el, name, out = []) { for (const x of el.c || []) { if (typeof x === 'string') continue; if (x.n === name) out.push(x); findAll(x, name, out); } return out; }

/**
 * writePptx(doc, { original, ext }) – with `original` (the bytes the doc was read from) the original package is patched:
 * unknown parts stay byte-identical, unchanged imported elements/timing/transitions/notes/comments are written back raw.
 * `ext` decides the main content type (pptm/ppsm/potm keep macros, pptx drops them).
 * If patching fails it throws (err.code 'PPTX_PATCH_FAILED', err.regenerate() → a fresh package without the unknown
 * parts, for "save as copy"); `regenerate: true` asks for that fresh package directly.
 */
export async function writePptx(doc, { original, ext = 'pptx', regenerate = false } = {}) {
  if (original && !regenerate) {
    try { return await writePatched(doc, original, ext); } catch (e) {
      // NEVER fall back silently: a regenerated package drops every part we do not model (SmartArt, embeddings,
      // macros, custom XML …). Fail the save visibly; the caller can offer "save as copy" via err.regenerate().
      const err = new Error(`This presentation could not be saved losslessly (${(e && e.message) || e}). Saving it would drop parts this editor cannot rebuild – save a copy instead.`);
      err.code = 'PPTX_PATCH_FAILED';
      err.cause = e;
      err.fidelity = { items: ['Parts of the original file that this editor does not model (e.g. SmartArt, embedded objects, macros, custom XML)'], kept: false };
      err.regenerate = () => writeFull(doc, ext);
      throw err;
    }
  }
  return writeFull(doc, ext);
}
