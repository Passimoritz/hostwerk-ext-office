/*
 * PPTX reader of the slides chunk (pptx/pptm/ppsx/ppsm/potx/potm) – browser + Node.
 *   readPptx(bytes, { ext }) → { doc, fidelity: [FIDELITY keys] }
 * Placeholder text styles are resolved (master txStyles → master ph → layout ph → slide lstStyle) into element.text.lvl;
 * only explicit slide-level paragraph/run props stay on paragraphs/runs. Colours become refs ('accent1+40', '#RRGGBB@50' …).
 */
import JSZip from 'jszip';
import { parseXml, serialize, text as textOf } from '../../_shared/xmltree.js';
import { resolveHex, lumMod as lumModHex, tint as tintHex, shade as shadeHex, DEFAULT_COLORS } from '../themes.js';
import { uid, DEFAULT_INSETS } from '../model.js';
import { resolvePrst, SHAPES } from '../render/geometry.js';
import { pathsOfCustGeom } from '../render/custom.js';
import { animationsOf as readAnimations } from './pptx-anim.js';
import { bytesToDataUrl, mimeOfPath } from './util.js';

/* ---------------------------------------------------------------- signatures (unchanged → write back raw) */
/** JSON with sorted keys (undefined skipped) */
export function stable(v) {
  if (v === undefined) return 'null';
  if (v === null || typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return `[${v.map(stable).join(',')}]`;
  return `{${Object.keys(v).sort().filter((k) => v[k] !== undefined).map((k) => `${JSON.stringify(k)}:${stable(v[k])}`).join(',')}}`;
}
/** full FNV-1a style hash (every character counts) */
export function fnv(s) {
  let h1 = 2166136261 >>> 0; let h2 = 5381;
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); h1 = Math.imul(h1 ^ c, 16777619) >>> 0; h2 = (Math.imul(h2, 33) + c) >>> 0; }
  return `${h1.toString(36)}.${h2.toString(36)}.${s.length.toString(36)}`;
}
export const sigOf = (v) => fnv(stable(v));
const SIG_SKIP = new Set(['id', 'x', 'y', 'w', 'h', 'rot', 'flipH', 'flipV', 'raw', 'rawRels', 'rawNs', 'rawXfrm', 'rawSig', 'rawDoc', 'rawSrc', 'spid', 'preserved', 'locked', '_xf']);
function sigObj(e, pb) {
  const o = {};
  for (const k of Object.keys(e)) { if (SIG_SKIP.has(k) || e[k] === undefined) continue; o[k] = k === 'children' && Array.isArray(e.children) ? e.children.map((c) => sigObj(c, e)) : e[k]; }
  if (pb) { const W = pb.w || 1; const H = pb.h || 1; o._rel = [(e.x - pb.x) / W, (e.y - pb.y) / H, e.w / W, e.h / H].map((v) => Math.round(v * 1e4)).concat([Math.round((e.rot || 0) * 100), e.flipH ? 1 : 0, e.flipV ? 1 : 0]); }
  return o;
}
/** signature of an element's content (position/size/rotation excluded; group children relative to the group) */
export const elementSig = (e) => sigOf(sigObj(e, null));
/** signature of a slide's animations with targets as original shape ids */
export const animSig = (anims, spidOf) => sigOf((anims || []).map((a) => [spidOf[a.el] || null, a.cls || 'entr', a.effect, a.dir || null, a.trigger || 'click', a.duration ?? null, a.delay ?? 0, a.path || null, a.scale || null, a.angle || null, a.opacity ?? null, a.raw ? a.raw.length : null]));
const usedNs = (raw, rootNs) => {
  const out = {};
  if (!rootNs) return out;
  const re = /<\/?([A-Za-z_][\w.-]*):|\s([A-Za-z_][\w.-]*):[\w.-]+=/g;
  let m;
  while ((m = re.exec(raw))) { const pf = m[1] || m[2]; if (pf !== 'xmlns' && rootNs[pf]) out[pf] = rootNs[pf]; }
  return out;
};
const rootNsOf = (el) => { const o = {}; for (const [k, v] of Object.entries((el && el.a) || {})) if (k.startsWith('xmlns:')) o[k.slice(6)] = v; return o; };
const relIdsIn = (raw) => { const ids = new Set(); raw.replace(/\sr:[A-Za-z]+="([^"]*)"/g, (m, v) => { ids.add(v); return m; }); return [...ids]; };

/* ---------------------------------------------------------------- xml helpers (prefix-agnostic) */
const ln = (n) => { if (typeof n !== 'string') return ''; const i = n.indexOf(':'); return i < 0 ? n : n.slice(i + 1); };
const K = (el, name) => { if (!el || !el.c) return null; for (const x of el.c) if (typeof x !== 'string' && ln(x.n) === name) return x; return null; };
const KS = (el, name) => (el && el.c ? el.c.filter((x) => typeof x !== 'string' && (!name || ln(x.n) === name)) : []);
const P = (el, path) => { let c = el; for (const p of path.split('/')) { if (!c) return null; c = K(c, p); } return c; };
const FIND = (el, name, out = []) => { for (const x of (el && el.c) || []) { if (typeof x === 'string') continue; if (ln(x.n) === name) out.push(x); FIND(x, name, out); } return out; };
const num = (v, d = 0) => { const n = parseInt(v, 10); return Number.isFinite(n) ? n : d; };
const bool = (v) => v === '1' || v === 'true' || v === 'on';
/** pick mc:AlternateContent branch: Choice for p14/a14 (we understand those), else Fallback */
function unwrapAlt(children, prefer = 'Fallback') {
  const out = [];
  for (const x of children) {
    if (typeof x === 'string') continue;
    if (ln(x.n) === 'AlternateContent') {
      const choice = K(x, 'Choice'); const fb = K(x, 'Fallback');
      const pick = prefer === 'Choice' ? (choice || fb) : (fb || choice);
      if (pick) out.push(...unwrapAlt(pick.c || [], prefer));
    } else out.push(x);
  }
  return out;
}

const dirOf = (p) => p.slice(0, p.lastIndexOf('/'));
function normPath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base ? base.split('/') : [];
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.' && seg !== '') parts.push(seg); }
  return parts.join('/');
}
const relsPathOf = (p) => { const i = p.lastIndexOf('/'); return `${p.slice(0, i)}/_rels/${p.slice(i + 1)}.rels`; };

/* ---------------------------------------------------------------- colours */
const PRST = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', gray: '#808080', grey: '#808080', darkGray: '#A9A9A9', lightGray: '#D3D3D3', orange: '#FFA500', purple: '#800080', navy: '#000080', cyan: '#00FFFF', magenta: '#FF00FF' };
const COLOR_TAGS = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'];
export function colorOf(node, ctx) {
  if (!node || typeof node !== 'object') return null;
  const c = COLOR_TAGS.includes(ln(node.n)) ? node : KS(node).find((x) => COLOR_TAGS.includes(ln(x.n)));
  if (!c || !c.a) return null;
  const t = ln(c.n);
  let base;
  if (t === 'srgbClr') base = `#${String(c.a.val || '000000').toUpperCase()}`;
  else if (t === 'sysClr') base = `#${String(c.a.lastClr || (c.a.val === 'window' ? 'FFFFFF' : '000000')).toUpperCase()}`;
  else if (t === 'prstClr') base = PRST[c.a.val] || '#000000';
  else if (t === 'scrgbClr') { const f = (v) => Math.round(Math.min(1, Math.max(0, num(v) / 100000)) ** (1 / 2.2) * 255).toString(16).padStart(2, '0'); base = `#${f(c.a.r)}${f(c.a.g)}${f(c.a.b)}`.toUpperCase(); }
  else if (t === 'hslClr') base = '#808080';
  else {
    base = c.a.val || 'tx1';
    if (base === 'phClr') base = (ctx && ctx.phClr) || 'accent1';
  }
  const mods = {};
  for (const m of KS(c)) mods[ln(m.n)] = num(m.a.val);
  let ref = base;
  const keys = Object.keys(mods).filter((k) => k !== 'alpha');
  const simple = base[0] !== '#' && /^[a-z]+\d?$/i.test(base) && !/[+\-@]/.test(base);
  if (keys.length && simple && keys.every((k) => k === 'lumMod' || k === 'lumOff')) {
    const lm = mods.lumMod ?? 100000; const lo = mods.lumOff ?? 0;
    if (lo > 0 && Math.abs(lm + lo - 100000) < 600) ref = `${base}+${Math.round(lo / 1000)}`;
    else if (!lo && lm < 100000) ref = `${base}-${Math.round(100 - lm / 1000)}`;
    else ref = bake(base, mods, ctx);
  } else if (keys.length) ref = bake(base, mods, ctx);
  if (mods.alpha !== undefined && mods.alpha < 100000) {
    if (/@\d+$/.test(ref)) ref = ref.replace(/@\d+$/, '');
    ref = `${ref}@${Math.round(mods.alpha / 1000)}`;
  }
  return ref;
}
function bake(base, mods, ctx) {
  let hex = base[0] === '#' ? base.slice(0, 7) : resolveHex(base, ctx && ctx.theme) || '#000000';
  if (mods.tint !== undefined) hex = tintHex(hex, mods.tint / 100000);
  if (mods.shade !== undefined) hex = shadeHex(hex, mods.shade / 100000);
  if (mods.lumMod !== undefined || mods.lumOff !== undefined) hex = lumModHex(hex, (mods.lumMod ?? 100000) / 100000, (mods.lumOff ?? 0) / 100000);
  return hex.toUpperCase();
}

function fillOf(spPr, ctx) {
  if (!spPr) return undefined;
  for (const x of KS(spPr)) {
    const t = ln(x.n);
    if (t === 'noFill') return null;
    if (t === 'solidFill') return { type: 'solid', color: colorOf(x, ctx) || '#000000' };
    if (t === 'gradFill') {
      const stops = KS(K(x, 'gsLst'), 'gs').map((g) => ({ pos: Math.round(num(g.a.pos) / 1000), color: colorOf(g, ctx) || '#000000' })).sort((a, b) => a.pos - b.pos);
      const lin = K(x, 'lin');
      const path = K(x, 'path');
      return { type: 'grad', kind: path ? 'radial' : 'linear', angle: lin ? Math.round(num(lin.a.ang) / 60000) : 90, stops: stops.length ? stops : [{ pos: 0, color: '#FFFFFF' }, { pos: 100, color: '#000000' }] };
    }
    if (t === 'blipFill') { const src = ctx.blip(K(x, 'blip')); return src ? { type: 'image', src, mode: K(x, 'tile') ? 'tile' : 'stretch' } : null; }
    if (t === 'pattFill') return { type: 'solid', color: colorOf(K(x, 'fgClr'), ctx) || '#000000' };
    if (t === 'grpFill') return ctx.grpFill === undefined ? null : ctx.grpFill;
  }
  return undefined;
}
const DASHES = { solid: 'solid', dash: 'dash', sysDash: 'sysDash', dot: 'dot', sysDot: 'sysDot', dashDot: 'dashDot', lgDash: 'lgDash', lgDashDot: 'dashDot', lgDashDotDot: 'dashDot', sysDashDot: 'dashDot', sysDashDotDot: 'dashDot' };
const ENDS = { triangle: 'triangle', arrow: 'arrow', stealth: 'stealth', oval: 'oval', diamond: 'diamond' };
function lineOf(lnNode, ctx, styleLine) {
  if (!lnNode) return styleLine === undefined ? undefined : styleLine;
  if (K(lnNode, 'noFill')) return null;
  const fill = K(lnNode, 'solidFill') || K(lnNode, 'gradFill');
  const color = fill ? (ln(fill.n) === 'gradFill' ? colorOf(K(K(fill, 'gsLst'), 'gs'), ctx) : colorOf(fill, ctx)) : styleLine ? styleLine.color : null;
  if (!color) return styleLine ? { ...styleLine, width: lnNode.a.w ? num(lnNode.a.w) : styleLine.width } : undefined;
  const dash = K(lnNode, 'prstDash');
  const head = K(lnNode, 'headEnd'); const tail = K(lnNode, 'tailEnd');
  const l = { color, width: lnNode.a.w ? num(lnNode.a.w) : (styleLine && styleLine.width) || 9525, dash: dash ? DASHES[dash.a.val] || 'solid' : 'solid' };
  if (lnNode.a.cap === 'rnd') l.cap = 'round'; else if (lnNode.a.cap === 'sq') l.cap = 'square';
  if (K(lnNode, 'round')) l.join = 'round';
  if (head && ENDS[head.a.type]) l.head = ENDS[head.a.type];
  if (tail && ENDS[tail.a.type]) l.tail = ENDS[tail.a.type];
  return l;
}
function shadowOf(spPr, ctx) {
  const eff = K(spPr, 'effectLst');
  if (!eff) return undefined;
  for (const t of ['glow', 'reflection', 'softEdge']) if (K(eff, t)) ctx.flag('effects');
  const s = K(eff, 'outerShdw');
  if (!s) return null;
  const c = colorOf(s, ctx) || '#000000';
  const m = /@(\d+)$/.exec(c);
  return { color: c.replace(/@\d+$/, ''), blur: num(s.a.blurRad), dist: num(s.a.dist), dir: Math.round(num(s.a.dir) / 60000), alpha: m ? num(m[1]) : 100 };
}

/* ---------------------------------------------------------------- text */
const ALG = { l: 'l', ctr: 'ctr', r: 'r', just: 'just', dist: 'just', justLow: 'just' };
function spc(node, size) {
  if (!node) return undefined;
  const pts = K(node, 'spcPts'); if (pts) return num(pts.a.val) / 100;
  const pct = K(node, 'spcPct'); if (pct) return Math.round(((num(pct.a.val) / 100000) * (size || 18) * 1.2) * 10) / 10;
  return undefined;
}
function runPropsOf(rPr, ctx) {
  const r = {};
  if (!rPr) return r;
  const a = rPr.a;
  if (a.sz) r.size = num(a.sz) / 100;
  if (a.b !== undefined) r.b = bool(a.b);
  if (a.i !== undefined) r.i = bool(a.i);
  if (a.u !== undefined) r.u = a.u !== 'none';
  if (a.strike !== undefined) r.s = a.strike !== 'noStrike';
  if (a.baseline) r.baseline = Math.round(num(a.baseline) / 1000);
  if (a.cap === 'all') r.caps = true;
  const fill = K(rPr, 'solidFill') || (K(rPr, 'gradFill') && K(K(K(rPr, 'gradFill'), 'gsLst'), 'gs'));
  if (fill) r.color = colorOf(fill, ctx);
  if (K(rPr, 'noFill')) r.color = '#FFFFFF@0';
  const hl = K(rPr, 'highlight'); if (hl) r.highlight = colorOf(hl, ctx);
  const latin = K(rPr, 'latin'); if (latin && latin.a.typeface) r.font = fontTok(latin.a.typeface);
  const link = K(rPr, 'hlinkClick'); if (link && link.a['r:id'] && ctx.rels[link.a['r:id']]) r.link = ctx.rels[link.a['r:id']].target;
  return r;
}
const fontTok = (f) => (/^\+mj/.test(f) ? '+mj' : /^\+mn/.test(f) ? '+mn' : f);
/** level style from <a:lvlNpPr>/<a:pPr> */
function lvlOf(pPr, ctx, base = {}) {
  if (!pPr) return { ...base };
  const l = { ...base };
  const a = pPr.a;
  if (a.marL !== undefined) l.marL = num(a.marL);
  if (a.indent !== undefined) l.indent = num(a.indent);
  if (a.algn) l.align = ALG[a.algn] || 'l';
  const d = runPropsOf(K(pPr, 'defRPr'), ctx);
  delete d.link;
  Object.assign(l, d);
  const lnSpc = K(pPr, 'lnSpc');
  if (lnSpc) { const pct = K(lnSpc, 'spcPct'); if (pct) l.lnSpc = Math.round(num(pct.a.val) / 1000); }
  const sb = spc(K(pPr, 'spcBef'), l.size); if (sb !== undefined) l.spcBef = sb;
  const sa = spc(K(pPr, 'spcAft'), l.size); if (sa !== undefined) l.spcAft = sa;
  const bu = bulletOf(pPr, ctx); if (bu !== undefined) l.bullet = bu;
  return l;
}
function bulletOf(pPr, ctx) {
  if (K(pPr, 'buNone')) return false;
  const ch = K(pPr, 'buChar'); const an = K(pPr, 'buAutoNum'); const blip = K(pPr, 'buBlip');
  if (!ch && !an && !blip) return undefined;
  const b = an ? { num: an.a.type || 'arabicPeriod', start: an.a.startAt ? num(an.a.startAt) : 1 } : { char: ch ? ch.a.char || '•' : '•' };
  const clr = K(pPr, 'buClr'); if (clr) b.color = colorOf(clr, ctx);
  const font = K(pPr, 'buFont'); if (font && ch && font.a.typeface && !/^\+/.test(font.a.typeface)) b.font = font.a.typeface;
  const sz = K(pPr, 'buSzPct'); if (sz) b.sizePct = Math.round(num(sz.a.val) / 1000);
  return b;
}
function lstOf(lst, ctx, base) {
  const out = base ? base.map((x) => ({ ...x })) : Array.from({ length: 9 }, () => ({}));
  if (!lst) return out;
  for (let i = 0; i < 9; i++) { const n = K(lst, `lvl${i + 1}pPr`); if (n) out[i] = lvlOf(n, ctx, out[i]); }
  const def = K(lst, 'defPPr'); if (def) for (let i = 0; i < 9; i++) out[i] = { ...lvlOf(def, ctx), ...out[i] };
  return out;
}
const mergeLvls = (a, b) => a.map((x, i) => ({ ...x, ...((b && b[i]) || {}) }));

function parasOf(txBody, ctx) {
  const paras = [];
  for (const p of KS(txBody, 'p')) {
    const pPr = K(p, 'pPr');
    const para = { runs: [], level: pPr && pPr.a.lvl ? Math.min(8, num(pPr.a.lvl)) : 0 };
    if (pPr) {
      const l = lvlOf(pPr, ctx);
      for (const k of ['align', 'marL', 'indent', 'lnSpc', 'spcBef', 'spcAft', 'bullet']) if (l[k] !== undefined) para[k] = l[k];
    }
    const kidsList = unwrapAlt(p.c || []);
    for (const x of kidsList) {
      const t = ln(x.n);
      if (t === 'r') para.runs.push({ text: textOf(K(x, 't')), ...runPropsOf(K(x, 'rPr'), ctx) });
      else if (t === 'br') para.runs.push({ text: '\n', ...runPropsOf(K(x, 'rPr'), ctx) });
      else if (t === 'fld') {
        const type = x.a.type || '';
        const field = type === 'slidenum' ? 'slidenum' : /^datetime/.test(type) ? 'datetime' : undefined;
        if (/^hostwerk\.live\?/.test(type)) {
          const q = new URLSearchParams(type.slice(14));
          let params; try { params = q.get('params') ? JSON.parse(q.get('params')) : undefined; } catch { params = undefined; }
          para.runs.push({ text: textOf(K(x, 't')), ...runPropsOf(K(x, 'rPr'), ctx), live: { id: q.get('id') || '', ...(params ? { params } : {}), ...(q.get('label') ? { label: q.get('label') } : {}) } });
          continue;
        }
        if (!field) ctx.flag('fields');
        para.runs.push({ text: textOf(K(x, 't')), ...runPropsOf(K(x, 'rPr'), ctx), ...(field ? { field } : {}) });
      } else if (t === 'm' || t === 'oMathPara' || t === 'oMath') ctx.flag('math');
    }
    if (FIND(p, 'm').length) ctx.flag('math');
    const end = K(p, 'endParaRPr');
    if (end) { const e = runPropsOf(end, ctx); if (Object.keys(e).length) para.endProps = e; }
    paras.push(para);
  }
  return paras.length ? paras : [{ runs: [], level: 0 }];
}
function bodyPrOf(bp, base = {}) {
  const b = { ...base };
  if (!bp) return b;
  const a = bp.a;
  if (a.anchor) b.anchor = a.anchor === 'ctr' ? 'ctr' : a.anchor === 'b' ? 'b' : 't';
  b.insets = { ...(b.insets || DEFAULT_INSETS) };
  if (a.lIns !== undefined) b.insets.l = num(a.lIns);
  if (a.tIns !== undefined) b.insets.t = num(a.tIns);
  if (a.rIns !== undefined) b.insets.r = num(a.rIns);
  if (a.bIns !== undefined) b.insets.b = num(a.bIns);
  if (a.wrap) b.wrap = a.wrap !== 'none';
  if (a.vert && a.vert !== 'horz') b.vert = a.vert === 'vert270' ? 'vert270' : 'vert';
  const nf = K(bp, 'normAutofit');
  if (nf) { b.autofit = 'shrink'; if (nf.a.fontScale) b.fontScale = num(nf.a.fontScale) / 100000; else delete b.fontScale; }
  else if (K(bp, 'spAutoFit')) b.autofit = 'resize';
  else if (K(bp, 'noAutofit')) b.autofit = 'none';
  return b;
}

/* ---------------------------------------------------------------- geometry */
function xfrmOf(x) {
  if (!x) return null;
  const off = K(x, 'off'); const ext = K(x, 'ext');
  const r = { x: num(off && off.a.x), y: num(off && off.a.y), w: num(ext && ext.a.cx), h: num(ext && ext.a.cy), rot: x.a.rot ? Math.round((num(x.a.rot) / 60000) * 100) / 100 : 0, flipH: bool(x.a.flipH), flipV: bool(x.a.flipV) };
  const co = K(x, 'chOff'); const ce = K(x, 'chExt');
  if (co || ce) r.ch = { x: num(co && co.a.x), y: num(co && co.a.y), w: num(ce && ce.a.cx), h: num(ce && ce.a.cy) };
  return r;
}
const IDT = (b) => b;
function applyT(T, b) { const p = T({ x: b.x, y: b.y, w: b.w, h: b.h }); return { ...b, ...p }; }
function groupT(T, g) {
  const ch = g.ch && g.ch.w && g.ch.h ? g.ch : { x: g.x, y: g.y, w: g.w || 1, h: g.h || 1 };
  const sx = ch.w ? g.w / ch.w : 1; const sy = ch.h ? g.h / ch.h : 1;
  return (b) => T({ x: g.x + (b.x - ch.x) * sx, y: g.y + (b.y - ch.y) * sy, w: b.w * sx, h: b.h * sy });
}
const R = (b) => ({ x: Math.round(b.x), y: Math.round(b.y), w: Math.max(0, Math.round(b.w)), h: Math.max(0, Math.round(b.h)) });

/* ---------------------------------------------------------------- placeholders */
const phInfo = (nvPr) => { const ph = K(nvPr, 'ph'); if (!ph) return null; return { type: ph.a.type || 'obj', idx: ph.a.idx !== undefined ? num(ph.a.idx) : undefined, hasIdx: ph.a.idx !== undefined }; };
const TITLE = (t) => t === 'title' || t === 'ctrTitle';
function matchPh(list, info) {
  if (!list || !info) return null;
  if (info.hasIdx) { const m = list.find((p) => p.info.hasIdx && p.info.idx === info.idx && (TITLE(p.info.type) === TITLE(info.type))); if (m) return m; }
  const m = list.find((p) => p.info.type === info.type);
  if (m) return m;
  if (info.type === 'ctrTitle') return list.find((p) => p.info.type === 'title') || null;
  if (info.type === 'subTitle' || info.type === 'obj' || info.type === 'body') return list.find((p) => p.info.type === 'body' || p.info.type === 'obj') || null;
  return null;
}
const styleKind = (type) => (TITLE(type) ? 'title' : ['body', 'obj', 'subTitle', 'tbl', 'chart', 'dgm', 'media', 'clipArt', 'pic'].includes(type) ? 'body' : 'other');

/* ---------------------------------------------------------------- reader */
export async function readPptx(bytes, { ext } = {}) {
  const zip = await JSZip.loadAsync(bytes);
  const flags = new Set();
  const flag = (k) => flags.add(k);
  const lost = new Set();
  const sourceId = uid('src');
  const has = (p) => !!zip.file(p);
  const str = async (p) => (has(p) ? zip.file(p).async('string') : null);
  const xml = async (p) => { const s = await str(p); return s ? parseXml(s) : null; };
  const relsCache = {};
  const rels = async (p) => {
    if (relsCache[p]) return relsCache[p];
    const x = await xml(relsPathOf(p));
    const out = {};
    for (const r of KS(x, 'Relationship')) out[r.a.Id] = { type: r.a.Type, target: r.a.TargetMode === 'External' ? r.a.Target : normPath(dirOf(p), r.a.Target), external: r.a.TargetMode === 'External' };
    return (relsCache[p] = out);
  };
  const mediaCache = {};
  const media = async (path) => {
    if (mediaCache[path] !== undefined) return mediaCache[path];
    const f = zip.file(path);
    mediaCache[path] = f ? bytesToDataUrl(await f.async('uint8array'), mimeOfPath(path)) : null;
    return mediaCache[path];
  };
  if (/^(pptm|ppsm|potm)$/i.test(ext || '') || Object.keys(zip.files).some((n) => /vbaProject\.bin$/i.test(n))) flag('macros');
  if (Object.keys(zip.files).some((n) => /^_xmlsignatures\//.test(n))) lost.add('signature');

  // presentation
  const rootRels = await xml('_rels/.rels');
  const main = KS(rootRels, 'Relationship').find((r) => /officeDocument$/.test(r.a.Type));
  const presPath = main ? normPath('', main.a.Target) : 'ppt/presentation.xml';
  const pres = await xml(presPath);
  if (!pres) throw new Error('Not a PowerPoint file');
  const presRels = await rels(presPath);
  const sz = K(pres, 'sldSz');
  const size = { w: num(sz && sz.a.cx, 12192000), h: num(sz && sz.a.cy, 6858000) };

  // theme (of the first master)
  const masterIds = KS(K(pres, 'sldMasterIdLst'), 'sldMasterId');
  const masterPath = masterIds.length && presRels[masterIds[0].a['r:id']] ? presRels[masterIds[0].a['r:id']].target : Object.values(presRels).find((r) => /slideMaster$/.test(r.type))?.target;
  const masterRels = masterPath ? await rels(masterPath) : {};
  const themeRel = Object.values(masterRels).find((r) => /theme$/.test(r.type)) || Object.values(presRels).find((r) => /theme$/.test(r.type));
  const themeX = themeRel ? await xml(themeRel.target) : null;
  const colors = { ...DEFAULT_COLORS };
  const cs = P(themeX, 'themeElements/clrScheme');
  for (const k of Object.keys(DEFAULT_COLORS)) { const n = K(cs, k); if (n) { const c = colorOf(n, null); if (c && c[0] === '#') colors[k] = c.slice(0, 7); } }
  const fs = P(themeX, 'themeElements/fontScheme');
  const fonts = { major: (P(fs, 'majorFont/latin') || { a: {} }).a.typeface || 'Calibri Light', minor: (P(fs, 'minorFont/latin') || { a: {} }).a.typeface || 'Calibri' };
  const master = masterPath ? await xml(masterPath) : null;
  const cm = K(master, 'clrMap');
  const theme = { id: 'imported', name: (themeX && themeX.a.name) || 'Imported', dark: false, colors, fonts, bg: { type: 'solid', color: 'bg1' } };
  if (cm) {
    const std = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
    if (cm.a.bg1 === 'dk1' && cm.a.tx1 === 'lt1') theme.dark = true;
    else if (['bg1', 'tx1', 'bg2', 'tx2'].some((k) => cm.a[k] && cm.a[k] !== std[k])) theme.clrMap = { bg1: cm.a.bg1, tx1: cm.a.tx1, bg2: cm.a.bg2, tx2: cm.a.tx2 };
  }
  const fmt = P(themeX, 'themeElements/fmtScheme');
  const fillStyles = KS(K(fmt, 'fillStyleLst'));
  const bgStyles = KS(K(fmt, 'bgFillStyleLst'));
  const lnStyles = KS(K(fmt, 'lnStyleLst'), 'ln');

  /* ---- context factory per part */
  const mkCtx = async (partPath) => {
    const r = await rels(partPath);
    const ctx = { theme, rels: r, flag, blipMap: {} };
    ctx.blip = (b) => (b && b.a['r:embed'] && ctx.blipMap[b.a['r:embed']]) || null;
    // preload images of this part
    for (const [id, rel] of Object.entries(r)) if (!rel.external && /\/(image|media|video|audio)$/.test(rel.type)) ctx.blipMap[id] = await media(rel.target);
    return ctx;
  };
  const styleRefFill = (ref, ctx, list = fillStyles, offset = 1) => {
    if (!ref) return undefined;
    const idx = num(ref.a.idx);
    if (!idx) return null;
    const c = colorOf(ref, ctx);
    const st = list[idx - offset];
    if (!st) return c ? { type: 'solid', color: c } : undefined;
    return fillOf({ c: [st] }, { ...ctx, phClr: c || 'accent1' });
  };
  const bgOf = (cSld, ctx) => {
    const bg = K(cSld, 'bg');
    if (!bg) return null;
    const pr = K(bg, 'bgPr');
    if (pr) return fillOf(pr, ctx) || null;
    const ref = K(bg, 'bgRef');
    // idx 1001+ points into bgFillStyleLst (PowerPoint writes <p:bgRef idx="1001"> for every default background)
    if (ref) { const idx = num(ref.a.idx); return idx >= 1001 ? styleRefFill({ n: ref.n, a: { ...ref.a, idx: String(idx - 1000) }, c: ref.c }, ctx, bgStyles) || null : styleRefFill(ref, ctx) || null; }
    return null;
  };

  /* ---- master/layout placeholder lists */
  const txStyles = K(master, 'txStyles');
  const presDefault = lstOf(K(pres, 'defaultTextStyle'), { theme, rels: {}, flag });
  const mctx = master ? await mkCtx(masterPath) : null;
  const masterStyle = {
    title: mergeLvls(presDefault, lstOf(K(txStyles, 'titleStyle'), mctx || { theme, flag })),
    body: mergeLvls(presDefault, lstOf(K(txStyles, 'bodyStyle'), mctx || { theme, flag })),
    other: mergeLvls(presDefault, lstOf(K(txStyles, 'otherStyle'), mctx || { theme, flag })),
  };
  const phList = (spTree, ctx) => {
    const out = [];
    for (const sp of unwrapAlt((spTree && spTree.c) || [])) {
      if (!['sp', 'pic', 'graphicFrame'].includes(ln(sp.n))) continue;
      const nv = KS(sp).find((x) => /^nv/.test(ln(x.n)));
      const info = phInfo(K(nv, 'nvPr'));
      if (!info) continue;
      const spPr = K(sp, 'spPr');
      const tx = K(sp, 'txBody');
      out.push({ info, node: sp, xfrm: xfrmOf(K(spPr, 'xfrm') || K(sp, 'xfrm')), lst: tx ? lstOf(K(tx, 'lstStyle'), ctx) : null, bodyPr: tx ? K(tx, 'bodyPr') : null, prompt: tx ? KS(tx, 'p').map((p) => FIND(p, 't').map(textOf).join('')).join('\n') : '' });
    }
    return out;
  };
  const masterPh = master ? phList(P(master, 'cSld/spTree'), mctx) : [];
  theme.bg = (master && bgOf(K(master, 'cSld'), mctx)) || { type: 'solid', color: 'bg1' };

  /** resolved geometry/body/level styles of a placeholder through the chain */
  const resolvePh = (info, layoutPhs, own) => {
    const lay = matchPh(layoutPhs, info);
    const mas = matchPh(masterPh, lay ? lay.info : info) || matchPh(masterPh, info);
    let lvl = masterStyle[styleKind(info.type)].map((x) => ({ ...x }));
    if (mas && mas.lst) lvl = mergeLvls(lvl, mas.lst);
    if (lay && lay.lst) lvl = mergeLvls(lvl, lay.lst);
    if (own) lvl = mergeLvls(lvl, own);
    let body = bodyPrOf(mas && mas.bodyPr, { anchor: TITLE(info.type) ? 'ctr' : 't', insets: { ...DEFAULT_INSETS }, autofit: 'none', wrap: true });
    if (lay && lay.bodyPr) body = bodyPrOf(lay.bodyPr, body);
    const xf = (lay && lay.xfrm) || (mas && mas.xfrm) || null;
    return { lay, mas, lvl, body, xf };
  };

  /* ---- element parsing */
  /** children with the mc:AlternateContent they came from (raw keeps the whole AlternateContent) */
  const nodesOf = (children) => {
    const out = [];
    for (const x of children || []) {
      if (typeof x === 'string') continue;
      if (ln(x.n) === 'AlternateContent') { const inner = unwrapAlt([x]); if (inner.length === 1) out.push({ node: inner[0], wrap: x }); else for (const n of inner) out.push({ node: n, wrap: null }); } else out.push({ node: x, wrap: null });
    }
    return out;
  };
  const attachRaw = (e, node, wrap, ctx, opts) => {
    const raw = serialize(wrap || node);
    e.raw = raw;
    e.rawRels = relIdsIn(raw).filter((id) => ctx.rels[id]).map((id) => ({ id, type: ctx.rels[id].type, target: ctx.rels[id].target, external: !!ctx.rels[id].external }));
    e.rawNs = usedNs(raw, opts.raw.ns);
    e.rawDoc = sourceId;
    e.rawSrc = opts.raw.part;
    const nv = KS(node).find((x) => /^nv/.test(ln(x.n)));
    const cnv = nv && K(nv, 'cNvPr');
    if (cnv && cnv.a.id) e.spid = num(cnv.a.id);
    const xf = e._xf;
    e.rawXfrm = xf ? { x: Math.round(xf.x), y: Math.round(xf.y), w: Math.round(xf.w), h: Math.round(xf.h), rot: xf.rot || 0, flipH: !!xf.flipH, flipV: !!xf.flipV } : null;
    if (wrap && FIND(wrap, 'contentPart').length) { e.preserved = true; flag('ink'); }
    delete e._xf;
    e.rawSig = elementSig(e);
  };
  const parseElR = async (node, wrap, ctx, opts, T) => {
    const e = await parseEl(node, ctx, opts, T);
    if (e && opts.raw) { try { attachRaw(e, node, wrap, ctx, opts); } catch { /* no raw */ } }
    if (e) delete e._xf;
    return e;
  };
  const parseTree = async (spTree, ctx, opts) => {
    const out = [];
    for (const { node, wrap } of nodesOf((spTree && spTree.c) || [])) {
      try { const e = await parseElR(node, wrap, ctx, opts, IDT); if (e) out.push(e); } catch (err) { /* never let one shape break the import */ }
    }
    return out;
  };
  const common = (nv, ctx, opts) => {
    const c = K(nv, 'cNvPr') || { a: {} };
    const e = { id: uid('e'), name: c.a.name || '' };
    if (c.a.descr) e.alt = c.a.descr;
    if (bool(c.a.hidden)) e.hidden = true;
    const hl = K(c, 'hlinkClick');
    if (hl && hl.a['r:id'] && ctx.rels[hl.a['r:id']]) e.link = ctx.rels[hl.a['r:id']].target;
    if (opts.idMap && c.a.id) opts.idMap[c.a.id] = e.id;
    return e;
  };
  const geomOf = (spPr, e) => {
    const pg = K(spPr, 'prstGeom');
    if (pg) {
      const prst = pg.a.prst || 'rect';
      e.geom = resolvePrst(prst);
      if (!SHAPES[prst]) flag('geometry');
      const gds = KS(K(pg, 'avLst'), 'gd');
      if (gds.length && SHAPES[e.geom] && SHAPES[e.geom].adj) {
        const names = SHAPES[e.geom].adj.length > 1 ? SHAPES[e.geom].adj.map((_, i) => `adj${i + 1}`) : ['adj'];
        const adj = names.map((n) => { const g = gds.find((x) => x.a.name === n || (n === 'adj' && x.a.name === 'adj1')); const m = g && /val\s+(-?\d+)/.exec(g.a.fmla || ''); return m ? num(m[1]) : null; });
        if (adj.some((v) => v != null)) e.adj = adj;
      }
      return prst;
    }
    if (K(spPr, 'custGeom')) {
      const paths = pathsOfCustGeom(K(spPr, 'custGeom'));
      if (paths) { e.geom = 'custom'; e.paths = paths; return 'cust'; }
      e.geom = 'rect'; flag('geometry'); return 'cust';
    }
    e.geom = 'rect';
    return 'rect';
  };
  const textOfBody = (tx, ctx, base, lvlBase) => {
    if (!tx) return undefined;
    const b = bodyPrOf(K(tx, 'bodyPr'), base || { anchor: 't', insets: { ...DEFAULT_INSETS }, autofit: 'none', wrap: true });
    const own = lstOf(K(tx, 'lstStyle'), ctx);
    b.lvl = lvlBase ? mergeLvls(lvlBase, own) : own;
    b.paras = parasOf(tx, ctx);
    return { paras: b.paras, anchor: b.anchor || 't', insets: b.insets || { ...DEFAULT_INSETS }, autofit: b.autofit || 'none', ...(b.fontScale ? { fontScale: b.fontScale } : {}), wrap: b.wrap !== false, ...(b.vert ? { vert: b.vert } : {}), lvl: b.lvl };
  };

  async function parseEl(node, ctx, opts, T) {
    const t = ln(node.n);
    if (t === 'sp' || t === 'cxnSp') {
      const nv = K(node, t === 'sp' ? 'nvSpPr' : 'nvCxnSpPr');
      const nvPr = K(nv, 'nvPr');
      const info = phInfo(nvPr);
      if (info && opts.skipPh) return null;
      const e = common(nv, ctx, opts);
      const spPr = K(node, 'spPr');
      const style = K(node, 'style');
      let xf = xfrmOf(K(spPr, 'xfrm'));
      let ph = null;
      if (info && opts.layoutPhs) {
        const tx = K(node, 'txBody');
        ph = resolvePh(info, opts.layoutPhs, tx ? lstOf(K(tx, 'lstStyle'), ctx) : null);
        if (!xf && ph.xf) xf = ph.xf;
        const key = ph.lay && opts.keyOf ? opts.keyOf(ph.lay) : opts.keyFor ? opts.keyFor(info) : null;
        if (key) e.ph = key;
        if (info.hasIdx) e.phIdx = info.idx;
        if (ph.lay && ph.lay.prompt) e.prompt = ph.lay.prompt;
      }
      if (!xf) xf = { x: 0, y: 0, w: 0, h: 0, rot: 0 };
      e._xf = xf;
      Object.assign(e, R(applyT(T, xf)));
      e.rot = xf.rot || 0;
      if (xf.flipH) e.flipH = true;
      if (xf.flipV) e.flipV = true;
      const locks = K(K(nv, t === 'sp' ? 'cNvSpPr' : 'cNvCxnSpPr'), 'spLocks'); if (locks && bool(locks.a.noMove)) e.locked = true;
      const styleLine = style && K(style, 'lnRef') && num(K(style, 'lnRef').a.idx) ? (() => { const r = K(style, 'lnRef'); const c = colorOf(r, ctx); const st = lnStyles[num(r.a.idx) - 1]; return { color: c || 'accent1', width: st && st.a.w ? num(st.a.w) : 12700, dash: 'solid' }; })() : undefined;
      const prst = geomOf(spPr, e);
      if (e.paths) e.paths = e.paths.map((p) => ({ ...p, w: p.w > 1 ? p.w : xf.w || 1, h: p.h > 1 ? p.h : xf.h || 1 }));
      if (t === 'cxnSp' || prst === 'line' || /Connector/.test(prst)) {
        e.type = 'line';
        e.geom = /bent/.test(prst) ? 'elbow' : /curved/.test(prst) ? 'curve' : 'line';
        delete e.adj;
        e.line = lineOf(K(spPr, 'ln'), ctx, styleLine) || styleLine || { color: 'tx1', width: 9525, dash: 'solid' };
        e.line = { head: 'none', tail: 'none', ...e.line };
        const sh = shadowOf(spPr, ctx); if (sh) e.shadow = sh;
        return e;
      }
      e.type = 'shape';
      if (bool((K(nv, 'cNvSpPr') || { a: {} }).a.txBox)) e.textbox = true;
      let fill = fillOf(spPr, ctx);
      if (fill === undefined) fill = style ? styleRefFill(K(style, 'fillRef'), ctx) : undefined;
      if (fill === undefined && ph && ph.lay) fill = fillOf(K(ph.lay.node, 'spPr'), ctx);
      e.fill = fill === undefined ? null : fill;
      let line = lineOf(K(spPr, 'ln'), ctx, styleLine);
      if (line === undefined) line = styleLine;
      e.line = line === undefined ? null : line;
      const sh = shadowOf(spPr, ctx);
      e.shadow = sh || null;
      if (K(spPr, 'scene3d') || K(spPr, 'sp3d')) flag('effects');
      const tx = K(node, 'txBody');
      if (info && ph) {
        e.text = textOfBody(tx, ctx, ph.body, ph.lvl) || { paras: [{ runs: [], level: 0 }], anchor: ph.body.anchor || 't', insets: ph.body.insets, autofit: ph.body.autofit || 'none', wrap: true, lvl: ph.lvl };
        if (tx) { const b = bodyPrOf(K(tx, 'bodyPr'), ph.body); e.text.anchor = b.anchor || 't'; e.text.insets = b.insets; e.text.autofit = b.autofit || 'none'; if (b.fontScale) e.text.fontScale = b.fontScale; }
        e.textbox = true;
      } else {
        let lvlBase = presDefault.map((x) => ({ ...x }));
        const fr = style && K(style, 'fontRef');
        if (fr) { const c = colorOf(fr, ctx); if (c) lvlBase = lvlBase.map((x) => ({ ...x, color: c })); }
        e.text = textOfBody(tx, ctx, { anchor: 't', insets: { ...DEFAULT_INSETS }, autofit: 'none', wrap: true }, lvlBase);
        if (!e.text) e.text = { paras: [{ runs: [], level: 0, align: 'ctr' }], anchor: 'ctr', insets: { ...DEFAULT_INSETS }, autofit: 'none', wrap: true, lvl: lvlBase };
      }
      if (opts.decor) { e.decor = true; e.locked = true; }
      return e;
    }
    if (t === 'pic') {
      const nv = K(node, 'nvPicPr');
      const nvPr = K(nv, 'nvPr');
      if (phInfo(nvPr) && opts.skipPh) return null;
      const e = common(nv, ctx, opts);
      const spPr = K(node, 'spPr');
      let xf = xfrmOf(K(spPr, 'xfrm'));
      const info = phInfo(nvPr);
      if (info && opts.layoutPhs) { const ph = resolvePh(info, opts.layoutPhs); if (!xf) xf = ph.xf; if (ph.lay && opts.keyOf) e.ph = opts.keyOf(ph.lay); }
      xf = xf || { x: 0, y: 0, w: 0, h: 0, rot: 0 };
      e._xf = xf;
      Object.assign(e, R(applyT(T, xf)), { rot: xf.rot || 0 });
      if (xf.flipH) e.flipH = true;
      if (xf.flipV) e.flipV = true;
      const bf = K(node, 'blipFill');
      const blip = K(bf, 'blip');
      const src = ctx.blip(blip);
      const sr = K(bf, 'srcRect');
      const crop = { l: 0, t: 0, r: 0, b: 0 };
      if (sr) for (const k of ['l', 't', 'r', 'b']) crop[k] = sr.a[k] ? num(sr.a[k]) / 100000 : 0;
      const am = K(blip, 'alphaModFix');
      // video/audio
      const vf = K(nvPr, 'videoFile') || K(nvPr, 'audioFile');
      if (vf) {
        const mediaEl = FIND(nvPr, 'media')[0];
        const rel = mediaEl && ctx.rels[mediaEl.a['r:embed']] ? ctx.rels[mediaEl.a['r:embed']] : vf.a['r:link'] && ctx.rels[vf.a['r:link']];
        if (rel && !rel.external && has(rel.target)) {
          const data = await media(rel.target);
          return { ...e, type: 'media', kind: ln(vf.n) === 'audioFile' ? 'audio' : 'video', src: data, mime: mimeOfPath(rel.target), poster: src || undefined };
        }
        flag('media');
      }
      const line = lineOf(K(spPr, 'ln'), ctx);
      return { ...e, type: 'image', src, crop, opacity: am ? num(am.a.amt, 100000) / 100000 : 1, line: line || null, shadow: shadowOf(spPr, ctx) || null, alt: e.alt || '' };
    }
    if (t === 'grpSp') {
      const nv = K(node, 'nvGrpSpPr');
      const e = common(nv, ctx, opts);
      const xf = xfrmOf(P(node, 'grpSpPr/xfrm')) || { x: 0, y: 0, w: 0, h: 0, rot: 0 };
      const abs = R(applyT(T, xf));
      const CT = groupT(T, xf);
      const gctx = { ...ctx, grpFill: fillOf(K(node, 'grpSpPr'), ctx) };
      const children = [];
      for (const { node: c, wrap } of nodesOf(node.c || [])) {
        if (!['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame'].includes(ln(c.n))) continue;
        try { const ce = await parseElR(c, wrap, gctx, { ...opts, layoutPhs: null }, CT); if (ce) children.push(ce); } catch { /* skip */ }
      }
      if (!children.length) return null;
      return { ...e, _xf: xf, type: 'group', ...abs, rot: xf.rot || 0, children, ...(opts.decor ? { decor: true, locked: true } : {}) };
    }
    if (t === 'graphicFrame') {
      const nv = K(node, 'nvGraphicFramePr');
      if (phInfo(K(nv, 'nvPr')) && opts.skipPh) return null;
      const e = common(nv, ctx, opts);
      const xf = xfrmOf(K(node, 'xfrm')) || { x: 0, y: 0, w: 0, h: 0 };
      e._xf = xf;
      Object.assign(e, R(applyT(T, xf)), { rot: 0 });
      const gd = P(node, 'graphic/graphicData');
      const uri = (gd && gd.a.uri) || '';
      if (/table$/.test(uri)) return { ...e, ...tableOf(K(gd, 'tbl'), ctx) };
      if (/chart$/.test(uri)) {
        flag('charts');
        const c = K(gd, 'chart');
        const rel = c && ctx.rels[c.a['r:id']];
        const ch = rel ? await chartOf(rel.target) : null;
        if (ch) return { ...e, type: 'chart', chart: ch };
        return { ...e, type: 'unknown', what: 'Chart', preserved: true };
      }
      if (/diagram$/.test(uri)) {
        flag('smartart');
        const dm = Object.values(ctx.rels).find((r) => /diagramDrawing$/.test(r.type));
        const rIds = K(gd, 'relIds');
        // the drawing part belonging to this diagram: same folder as the data part, matching number
        const dataRel = rIds && ctx.rels[rIds.a['r:dm']];
        let drawing = null;
        if (dataRel) { const nm = /(\d+)\.xml$/.exec(dataRel.target); drawing = Object.values(ctx.rels).find((r) => /diagramDrawing$/.test(r.type) && nm && r.target.endsWith(`${nm[1]}.xml`)) || dm; }
        if (drawing && has(drawing.target)) {
          const dx = await xml(drawing.target);
          const dctx = await mkCtx(drawing.target);
          const kidsT = (b) => T({ x: xf.x + b.x, y: xf.y + b.y, w: b.w, h: b.h });
          const children = [];
          for (const c of unwrapAlt(P(dx, 'spTree')?.c || [])) { try { const ce = await parseEl(c, dctx, { ...opts, layoutPhs: null, idMap: null, raw: null }, kidsT); if (ce) { delete ce._xf; children.push(ce); } } catch { /* skip */ } }
          if (children.length) return { ...e, type: 'group', what: 'SmartArt', preserved: true, children };
        }
        return { ...e, type: 'unknown', what: 'SmartArt', preserved: true };
      }
      if (/ole$/.test(uri) || FIND(gd, 'oleObj').length) {
        flag('ole');
        const pic = FIND(gd, 'pic')[0];
        if (pic) { const p = await parseEl(pic, ctx, { ...opts, idMap: null, raw: null }, () => R(applyT(T, xf))); if (p) return { ...p, id: e.id, name: e.name, ...R(applyT(T, xf)), _xf: xf, preserved: true, what: 'Embedded object' }; }
        return { ...e, type: 'unknown', what: 'Embedded object', preserved: true };
      }
      if (/model3d/i.test(uri)) flag('model3d');
      return { ...e, type: 'unknown', what: 'Object', preserved: true };
    }
    if (t === 'contentPart') { flag('ink'); return null; }
    return null;
  }

  function tableOf(tbl, ctx) {
    const pr = K(tbl, 'tblPr') || { a: {} };
    const sid = textOf(K(pr, 'tableStyleId')).trim().toUpperCase();
    const STY = ['{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}', '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}', '{00A15C55-8517-42AA-B614-E9B94910E393}', '{7DF18680-E054-41AD-8BC1-D1AEF772440D}', '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}', '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}'];
    const ai = STY.indexOf(sid);
    const style = { firstRow: bool(pr.a.firstRow), bandRows: bool(pr.a.bandRow), firstCol: bool(pr.a.firstCol), accent: `accent${ai >= 0 ? ai + 1 : 1}` };
    if (!sid) { style.firstRow = false; style.bandRows = false; }
    const cols = KS(K(tbl, 'tblGrid'), 'gridCol').map((g) => num(g.a.w));
    const lvlBase = Array.from({ length: 9 }, (_, i) => ({ size: 18, color: 'tx1', font: '+mn', marL: i * 457200, bullet: false }));
    const rows = KS(tbl, 'tr').map((tr) => ({
      h: num(tr.a.h),
      cells: KS(tr, 'tc').map((tc) => {
        const tcPr = K(tc, 'tcPr') || { a: {} };
        const b = textOfBody(K(tc, 'txBody'), ctx, { anchor: 't', insets: { l: num(tcPr.a.marL, 91440), r: num(tcPr.a.marR, 91440), t: num(tcPr.a.marT, 45720), b: num(tcPr.a.marB, 45720) }, autofit: 'none', wrap: true }, lvlBase) || { paras: [{ runs: [], level: 0 }], anchor: 't', insets: { ...DEFAULT_INSETS }, autofit: 'none', wrap: true };
        if (tcPr.a.anchor) b.anchor = tcPr.a.anchor === 'ctr' ? 'ctr' : tcPr.a.anchor === 'b' ? 'b' : 't';
        b.insets = { l: num(tcPr.a.marL, 91440), r: num(tcPr.a.marR, 91440), t: num(tcPr.a.marT, 45720), b: num(tcPr.a.marB, 45720) };
        const cell = { text: b };
        const f = fillOf(tcPr, ctx);
        if (f !== undefined) cell.fill = f;
        if (tc.a.gridSpan || tc.a.rowSpan) cell.span = { c: num(tc.a.gridSpan, 1), r: num(tc.a.rowSpan, 1) };
        if (bool(tc.a.hMerge) || bool(tc.a.vMerge)) cell.merged = true;
        const borders = {};
        let any = false;
        for (const k of ['L', 'R', 'T', 'B']) { const l = K(tcPr, `ln${k}`); if (l) { borders[k.toLowerCase()] = lineOf(l, ctx) ?? null; any = true; } }
        if (any) cell.borders = borders;
        return cell;
      }),
    }));
    return { type: 'table', cols, rows, style, border: undefined };
  }

  async function chartOf(path) {
    const cx = await xml(path);
    if (!cx) return null;
    const plot = FIND(cx, 'plotArea')[0];
    if (!plot) return null;
    const ctx = { theme, flag };
    const TYPES = { barChart: 'col', bar3DChart: 'col', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'line', radarChart: 'line', ofPieChart: 'pie' };
    const node = KS(plot).find((x) => TYPES[ln(x.n)]);
    if (!node) return null;
    let kind = TYPES[ln(node.n)];
    if (kind === 'col' && K(node, 'barDir') && K(node, 'barDir').a.val === 'bar') kind = 'bar';
    if (/3D/.test(node.n) || ['scatterChart', 'radarChart'].includes(ln(node.n))) flag('charts');
    const grouping = K(node, 'grouping');
    const pts = (n) => {
      const cache = n && (FIND(n, 'strCache')[0] || FIND(n, 'numCache')[0] || FIND(n, 'strLit')[0] || FIND(n, 'numLit')[0] || FIND(n, 'lvl')[0]);
      if (!cache) return [];
      const cnt = num((K(cache, 'ptCount') || { a: {} }).a.val, 0);
      const out = Array.from({ length: cnt }, () => '');
      for (const p of KS(cache, 'pt')) out[num(p.a.idx)] = textOf(K(p, 'v'));
      return out;
    };
    let categories = [];
    const series = KS(node, 'ser').map((s, i) => {
      if (!categories.length) categories = pts(K(s, 'cat') || K(s, 'xVal'));
      const tx = K(s, 'tx');
      const name = tx ? (FIND(tx, 'v')[0] ? textOf(FIND(tx, 'v')[0]) : textOf(tx)) : `Series ${i + 1}`;
      const sp = K(s, 'spPr');
      const fill = sp && (K(sp, 'solidFill') || P(sp, 'ln/solidFill'));
      return { name, values: pts(K(s, 'val') || K(s, 'yVal')).map((v) => (v === '' ? null : Number(v))), color: (fill && colorOf(fill, ctx)) || `accent${(i % 6) + 1}` };
    });
    const title = FIND(cx, 'title')[0];
    const autoDel = FIND(cx, 'autoTitleDeleted')[0];
    let titleText = title ? FIND(title, 't').map(textOf).join('') : '';
    if (title && !titleText && !(autoDel && bool(autoDel.a.val)) && series.length === 1) titleText = series[0].name;
    const dl = FIND(node, 'showVal')[0];
    return { kind, title: titleText, legend: !!FIND(cx, 'legend')[0], labels: !!(dl && bool(dl.a.val)), stacked: !!(grouping && /stacked/i.test(grouping.a.val)), categories: categories.map(String), series };
  }

  /* ---- layouts */
  const LKIND = { title: 'title', obj: 'titleContent', tx: 'titleContent', secHead: 'section', twoObj: 'twoContent', twoColTx: 'twoContent', twoTxTwoObj: 'comparison', titleOnly: 'titleOnly', blank: 'blank', picTx: 'pictureCaption', objTx: 'pictureCaption' };
  const keyAssigner = () => {
    let n = 0;
    return (info) => {
      if (['title', 'ctrTitle', 'subTitle', 'pic', 'dt', 'ftr', 'sldNum'].includes(info.type)) return info.type;
      n++;
      return n === 1 ? 'body' : `body${n}`;
    };
  };
  const layoutRelIds = KS(K(master, 'sldLayoutIdLst'), 'sldLayoutId');
  const layoutPaths = layoutRelIds.map((l) => masterRels[l.a['r:id']]).filter(Boolean).map((r) => r.target);
  for (const r of Object.values(masterRels)) if (/slideLayout$/.test(r.type) && !layoutPaths.includes(r.target)) layoutPaths.push(r.target);
  const layouts = [];
  const layoutInfo = {};
  const masterDecor = master ? await parseTree(P(master, 'cSld/spTree'), mctx, { skipPh: true, decor: true }) : [];
  for (const lp of layoutPaths) {
    try {
      const lx = await xml(lp);
      if (!lx) continue;
      const lctx = await mkCtx(lp);
      const phs = phList(P(lx, 'cSld/spTree'), lctx);
      const assign = keyAssigner();
      const keyMap = new Map();
      const ph = {};
      for (const p of phs) {
        const key = assign(p.info);
        if (ph[key]) continue;
        keyMap.set(p, key);
        const r = resolvePh(p.info, phs);
        const xf = r.xf || { x: 0, y: 0, w: size.w, h: size.h };
        ph[key] = { x: xf.x, y: xf.y, w: xf.w, h: xf.h, anchor: r.body.anchor || 't', lvl: r.lvl, prompt: p.prompt || '', type: p.info.type, idx: p.info.idx };
      }
      const showMaster = lx.a.showMasterSp !== '0';
      const decor = [...(showMaster ? masterDecor : []), ...(await parseTree(P(lx, 'cSld/spTree'), lctx, { skipPh: true, decor: true }))];
      const kind = LKIND[lx.a.type] || (lx.a.type ? 'custom' : 'custom');
      const id = `layout-${layouts.length + 1}`;
      layouts.push({ id, kind, name: (K(lx, 'cSld') || { a: {} }).a.name || kind, part: lp, bg: bgOf(K(lx, 'cSld'), lctx) || theme.bg, decor, ph });
      layoutInfo[lp] = { id, phs, keyOf: (p) => keyMap.get(p) || null };
    } catch { /* skip broken layout */ }
  }
  if (!layouts.length) layouts.push({ id: 'layout-1', kind: 'blank', name: 'Blank', bg: theme.bg, decor: [], ph: {} });

  /* ---- comment authors (legacy + modern) */
  const authors = {};
  for (const x of FIND(await xml('ppt/commentAuthors.xml'), 'cmAuthor')) authors[x.a.id] = x.a.name || '';
  const authRel = Object.values(presRels).find((r) => /authors$/i.test(r.type));
  for (const x of FIND(await xml(authRel ? authRel.target : 'ppt/authors.xml'), 'author')) authors[x.a.id] = x.a.name || '';

  /* ---- slides */
  const slideIds = KS(K(pres, 'sldIdLst'), 'sldId');
  const numToSlide = {};
  const slides = [];
  for (const s of slideIds) {
    const rel = presRels[s.a['r:id']];
    if (!rel || !has(rel.target)) continue;
    try {
      const sp = rel.target;
      const sstr = await str(sp);
      const sx = parseXml(sstr);
      const sctx = await mkCtx(sp);
      const rootNs = rootNsOf(sx);
      const lrel = Object.values(sctx.rels).find((r) => /slideLayout$/.test(r.type));
      const li = lrel ? layoutInfo[lrel.target] : null;
      const idMap = {};
      const used = new Set();
      const opts = {
        idMap, layoutPhs: li ? li.phs : [], keyOf: (p) => { const k = li && li.keyOf(p); if (!k || used.has(k)) return undefined; used.add(k); return k; },
        keyFor: (info) => (TITLE(info.type) ? info.type : null),
        raw: { ns: rootNs, part: sp },
      };
      const elements = await parseTree(P(sx, 'cSld/spTree'), sctx, opts);
      const slide = {
        id: uid('s'), layoutId: li ? li.id : layouts[0].id, hidden: sx.a.show === '0', background: bgOf(K(sx, 'cSld'), sctx), elements, notes: '',
        transition: { type: 'none', dir: 'l', duration: 500, advanceAfter: null }, animations: [], comments: [],
      };
      if (sx.a.showMasterSp === '0') slide.showMaster = false;
      slide.srcPart = sp;
      slide.rawDoc = sourceId;
      const cSld = K(sx, 'cSld');
      if (cSld && cSld.a.name) slide.name = cSld.a.name;
      if (sx.a['mc:Ignorable']) slide.rawIgnorable = sx.a['mc:Ignorable'];
      const slideRaws = [];
      const bgNode = K(cSld, 'bg');
      if (bgNode) { slide.bgRaw = serialize(bgNode); slide.bgSig = sigOf(slide.background); slideRaws.push(slide.bgRaw); }
      const extra = {};
      for (const [k, nm] of [['custData', 'custDataLst'], ['controls', 'controls'], ['cSldExt', 'extLst']]) { const x = K(cSld, nm); if (x) { extra[k] = serialize(x); slideRaws.push(extra[k]); } }
      const sExt = KS(sx).find((x) => ln(x.n) === 'extLst');
      if (sExt) { extra.sldExt = serialize(sExt); slideRaws.push(extra.sldExt); }
      if (Object.keys(extra).length) slide.rawExtra = extra;
      // notes
      const nrel = Object.values(sctx.rels).find((r) => /notesSlide$/.test(r.type));
      if (nrel && has(nrel.target)) {
        const nx = await xml(nrel.target);
        const body = FIND(nx, 'sp').find((x) => { const i = phInfo(K(K(x, 'nvSpPr'), 'nvPr')); return i && i.type === 'body'; });
        if (body) slide.notes = KS(K(body, 'txBody'), 'p').map((p) => unwrapAlt(p.c || []).map((r) => (ln(r.n) === 'br' ? '\n' : ['r', 'fld'].includes(ln(r.n)) ? textOf(K(r, 't')) : '')).join('')).join('\n').replace(/\s+$/, '');
        slide.notesPart = nrel.target;
        slide.notesSig = sigOf(slide.notes);
      }
      // comments (legacy + modern)
      for (const r of Object.values(sctx.rels)) {
        if (!/comments$/i.test(r.type) || !has(r.target)) continue;
        const cx = await xml(r.target);
        const byKey = {};
        for (const c of KS(cx, 'cm')) {
          const txt = K(c, 'txBody') ? KS(K(c, 'txBody'), 'p').map((p) => FIND(p, 't').map(textOf).join('')).join('\n') : textOf(K(c, 'text'));
          const cid = uid('c');
          const cm = { id: cid, author: authors[c.a.authorId] || c.a.authorId || '', text: txt, time: c.a.dt || c.a.created || '' };
          const pos = K(c, 'pos');
          if (pos && pos.a.x != null) cm.pos = { x: Math.round(num(pos.a.x) * 1587.5), y: Math.round(num(pos.a.y) * 1587.5) };
          if (c.a.status === 'resolved' || c.a.status === 'closed') cm.done = true;
          const thr = FIND(c, 'parentCm')[0];
          if (thr && byKey[`${thr.a.authorId}/${thr.a.idx}`]) cm.parent = byKey[`${thr.a.authorId}/${thr.a.idx}`];
          if (c.a.idx != null) byKey[`${c.a.authorId}/${c.a.idx}`] = cid;
          slide.comments.push(cm);
          for (const rp of KS(K(c, 'replyLst'), 'reply')) slide.comments.push({ id: uid('c'), parent: cid, author: authors[rp.a.authorId] || rp.a.authorId || '', text: KS(K(rp, 'txBody'), 'p').map((p) => FIND(p, 't').map(textOf).join('')).join('\n'), time: rp.a.created || '' });
        }
        if (!slide.commentsPart) slide.commentsPart = { path: r.target, type: r.type };
      }
      if (slide.commentsPart) slide.commentsSig = sigOf(slide.comments);
      // transition
      const trs = unwrapAlt(sx.c || [], 'Choice').filter((x) => ln(x.n) === 'transition');
      if (trs.length) slide.transition = transitionOf(trs[0]);
      const trNode = KS(sx).find((x) => ln(x.n) === 'transition' || (ln(x.n) === 'AlternateContent' && FIND(x, 'transition').length));
      if (trNode) { slide.transitionRaw = serialize(trNode); slide.transitionSig = sigOf(slide.transition); slideRaws.push(slide.transitionRaw); }
      // animations (+ the original timing XML, written back unchanged while the animations are untouched)
      const timing = K(sx, 'timing');
      if (timing) {
        slide.animations = animationsOf(timing, idMap);
        const i0 = sstr.indexOf('<p:timing');
        const i1 = sstr.indexOf('</p:timing>', i0);
        slide.timingRaw = i0 >= 0 && i1 > i0 ? sstr.slice(i0, i1 + 11) : serialize(timing);
        slide.timingSpids = [...new Set(FIND(timing, 'spTgt').map((x) => num(x.a.spid)).filter(Boolean))];
        const spidOf = {};
        for (const [sid, eid] of Object.entries(idMap)) spidOf[eid] = num(sid);
        slide.animSig = animSig(slide.animations, spidOf);
        slideRaws.push(slide.timingRaw);
      }
      const allRaw = slideRaws.join('');
      slide.rawRels = relIdsIn(allRaw).filter((id) => sctx.rels[id]).map((id) => ({ id, type: sctx.rels[id].type, target: sctx.rels[id].target, external: !!sctx.rels[id].external }));
      slide.rawNs = usedNs(allRaw, rootNs);
      numToSlide[s.a.id] = slide.id;
      slides.push(slide);
    } catch (err) { /* skip broken slide */ }
  }

  function transitionOf(tr) {
    const spd = { fast: 500, med: 750, slow: 1000 }[tr.a.spd] || 700;
    const dur = tr.a['p14:dur'] ? num(tr.a['p14:dur']) : tr.a.dur ? num(tr.a.dur) : spd;
    const out = { type: 'none', dir: 'l', duration: dur, advanceAfter: tr.a.advTm ? num(tr.a.advTm) : null };
    const c = KS(tr).find((x) => !['sndAc', 'extLst'].includes(ln(x.n)));
    if (c && ln(c.n) === 'morph') { out.type = 'morph'; out.option = c.a.option === 'byWord' ? 'word' : c.a.option === 'byChar' ? 'char' : 'object'; return out; }
    if (c) {
      const t = ln(c.n);
      const MAP = { fade: 'fade', push: 'push', wipe: 'wipe', zoom: 'zoom', cover: 'cover', split: 'split', pull: 'cover', cut: 'none', dissolve: 'fade', randomBar: 'wipe', strips: 'wipe', blinds: 'wipe', checker: 'fade', circle: 'zoom', diamond: 'zoom', plus: 'zoom', wedge: 'wipe', wheel: 'wipe', newsflash: 'zoom', random: 'fade' };
      out.type = MAP[t] || 'fade';
      if (!MAP[t] || !['fade', 'push', 'wipe', 'zoom', 'cover', 'split', 'cut'].includes(t)) flag('transitions');
      const d = c.a.dir;
      if (['l', 'r', 'u', 'd'].includes(d)) out.dir = d;
      else if (d === 'horz' || d === 'vert') out.dir = d === 'horz' ? 'l' : 'u';
      if (t === 'fade' && bool(c.a.thruBlk)) flag('transitions');
    } else if (out.advanceAfter == null) out.type = 'none';
    return out;
  }
  function animationsOf(timing, idMap) {
    return readAnimations(timing, idMap, uid, flag);
  }

  /* ---- sections */
  const sections = [];
  const secList = FIND(K(pres, 'extLst'), 'sectionLst')[0];
  for (const sec of KS(secList, 'section')) {
    const first = FIND(sec, 'sldId').map((x) => numToSlide[x.a.id]).find(Boolean);
    if (first) sections.push({ id: uid('sec'), name: sec.a.name || '', start: first });
  }

  const guides = FIND(K(pres, 'extLst'), 'sldGuideLst')[0];
  const guideList = KS(guides, 'guide').map((g, i) => ({ id: `gd${i + 1}`, axis: g.a.orient === 'horz' ? 'y' : 'x', pos: Math.round(num(g.a.pos) * 1587.5) }));

  /* ---- show settings & meta */
  const pp = await xml('ppt/presProps.xml');
  const showPr = K(pp, 'showPr');
  const core = await xml('docProps/core.xml');
  const metaText = (n) => { const x = core && FIND(core, n)[0]; return x ? textOf(x).trim() : ''; };
  const doc = {
    size, theme, layouts, slides, sections,
    footer: { text: '', showFooter: false, showNumber: false, showDate: false, date: '', hideOnTitle: true },
    show: { loop: !!(showPr && bool(showPr.a.loop)), autoAdvance: 0, useTimings: !(showPr && showPr.a.useTimings === '0') },
    meta: { title: metaText('title'), author: metaText('creator'), created: metaText('created') || new Date().toISOString(), modified: metaText('modified') || new Date().toISOString() },
  };
  if (guideList.length) doc.guides = guideList;
  doc.sourceId = sourceId;
  // fidelity: what is lost when saving back to pptx with `original`; preserved: kept as it is but not editable here
  return { doc, fidelity: [...lost], preserved: [...flags] };
}
