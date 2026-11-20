/*
 * Word tables – pure helpers shared by the docx reader/writer, the editor's table view and the HTML renderer.
 *
 * Model (tiptap JSON attrs):
 *   table     { layout: 'fixed'|null (autofit), width: { type: 'pct'|'dxa'|'auto', w } | null, indent (twips), align,
 *               borders: { top, left, bottom, right, insideH, insideV } (border strings), cellMargins: { top, left,
 *               bottom, right } (twips), styleId, look: { firstRow, lastRow, firstColumn, lastColumn, noHBand, noVBand },
 *               tstyle (display spec of the table style – never written), ox: { tblPr } (original XML) }
 *   tableRow  { height (twips), heightRule: 'atLeast'|'exact', cantSplit, ox }
 *   cell      { vAlign, borders, margins (twips), textDir, noWrap, ox, + colspan/rowspan/colwidth/backgroundColor }
 * A border string is "val:sz:color" (Word units: sz in eighths of a point, colour hex or "auto"), "nil" = no border.
 */

export const SIDES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
const LOOK_KEYS = ['firstRow', 'lastRow', 'firstColumn', 'lastColumn', 'noHBand', 'noVBand'];

/* ---------------- colours ---------------- */
const hex2 = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0').toUpperCase();
function rgb(hex) { const h = String(hex || '000000').replace('#', ''); return [0, 2, 4].map((k) => parseInt(h.slice(k, k + 2), 16) || 0); }
function toHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b); const mn = Math.min(r, g, b); const l = (mx + mn) / 2;
  if (mx === mn) return [0, 0, l];
  const d = mx - mn; const s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  const h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}
function fromHsl([h, s, l]) {
  if (!s) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q;
  const f = (t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}
/** Word's "lighter n %" theme tint (lumMod/lumOff): tint(accent, 0.2) = the 20 % shade of the colour gallery */
export function tint(hex, amount) { const [h, s, l] = toHsl(rgb(hex)); return fromHsl([h, s, l * amount + (1 - amount)]).map(hex2).join(''); }

/* ---------------- borders ---------------- */
const CSS_STYLE = { single: 'solid', thick: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed', dotDash: 'dashed', dotDotDash: 'dotted', triple: 'double', thinThickSmallGap: 'double', thickThinSmallGap: 'double', thinThickMediumGap: 'double', thickThinMediumGap: 'double', wave: 'solid', doubleWave: 'double', threeDEmboss: 'ridge', threeDEngrave: 'groove', outset: 'outset', inset: 'inset' };
export function parseBorder(s) {
  if (!s) return null;
  const [val = 'single', sz = '4', color = 'auto', space = '0'] = String(s).split(':');
  return { val, sz: Number(sz) || 0, color, space: Number(space) || 0 };
}
export const isNone = (s) => !s || /^(nil|none)\b/.test(String(s));
/** border string → CSS border shorthand ("none" for nil/none) */
export function borderCss(s) {
  if (isNone(s)) return 'none';
  const b = parseBorder(s);
  const style = CSS_STYLE[b.val] || 'solid';
  let px = Math.max(1, Math.round(((b.sz || 4) / 8) * (4 / 3) * 10) / 10);
  if (b.val === 'thick') px = Math.max(px, 2);
  if (style === 'double') px = Math.max(3, px * 3);
  return `${px}px ${style} ${b.color && b.color !== 'auto' ? `#${b.color}` : '#000'}`;
}
/** <w:top w:val=… /> elements of a w:tblBorders / w:tcBorders / w:pBdr element → { side: string } */
export function bordersFromEl(el, kidsOf) {
  if (!el) return null;
  const o = {};
  for (const c of kidsOf(el)) {
    const name = c.n.replace(/^w:/, '');
    const side = name === 'start' ? 'left' : name === 'end' ? 'right' : name;
    if (!['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'between', 'bar', 'tl2br', 'tr2bl'].includes(side)) continue;
    const v = c.a['w:val'] || 'single';
    o[side] = v === 'nil' || v === 'none' ? 'nil' : `${v}:${c.a['w:sz'] || 4}:${c.a['w:color'] || 'auto'}:${c.a['w:space'] || 0}`;
  }
  return Object.keys(o).length ? o : null;
}
const bXml = (tag, s) => {
  if (s == null) return '';
  if (isNone(s)) return `<w:${tag} w:val="nil"/>`;
  const b = parseBorder(s);
  return `<w:${tag} w:val="${b.val}" w:sz="${b.sz || 4}" w:space="${b.space || 0}" w:color="${b.color || 'auto'}"/>`;
};
/** { side: string } → <w:tblBorders>/<w:tcBorders>/<w:pBdr> XML (Word's child order) */
export function bordersXml(tag, borders, order = SIDES) {
  if (!borders) return '';
  const inner = order.map((s) => bXml(s, borders[s])).join('');
  return inner ? `<w:${tag}>${inner}</w:${tag}>` : '';
}
export const TBL_BORDER_ORDER = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
export const TC_BORDER_ORDER = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'];
export const P_BORDER_ORDER = ['top', 'left', 'bottom', 'right', 'between', 'bar'];

/** cell margins element → { top, left, bottom, right } in twips (dxa only) */
export function marginsFromEl(el, kidsOf) {
  if (!el) return null;
  const o = {};
  for (const c of kidsOf(el)) {
    const name = c.n.replace(/^w:/, '');
    const side = name === 'start' ? 'left' : name === 'end' ? 'right' : name;
    if (!['top', 'left', 'bottom', 'right'].includes(side)) continue;
    if ((c.a['w:type'] || 'dxa') === 'dxa' && c.a['w:w'] != null) o[side] = Number(c.a['w:w']) || 0;
  }
  return Object.keys(o).length ? o : null;
}
export function marginsXml(tag, m) {
  if (!m) return '';
  const inner = ['top', 'left', 'bottom', 'right'].filter((s) => m[s] != null).map((s) => `<w:${s} w:w="${Math.round(m[s])}" w:type="dxa"/>`).join('');
  return inner ? `<w:${tag}>${inner}</w:${tag}>` : '';
}

/* ---------------- table styles (display spec from the style XML) ---------------- */
const COND = ['wholeTable', 'firstRow', 'lastRow', 'firstCol', 'lastCol', 'band1Vert', 'band2Vert', 'band1Horz', 'band2Horz', 'neCell', 'nwCell', 'seCell', 'swCell'];
/**
 * one w:style (type table) element → partial spec; `base` = the spec of its basedOn style (merged under it)
 * spec = { b: borders, cm: cell margins, fill, color, bold, pAfter (pt), pLine (multiple), cond: { firstRow: { fill, color, bold, italic, caps, b } … } }
 */
export function specFromStyle(el, { kid, kids }, base = null) {
  const spec = base ? JSON.parse(JSON.stringify(base)) : { cond: {} };
  if (!spec.cond) spec.cond = {};
  const w = (e, n) => (e ? kid(e, `w:${n}`) : null);
  const on = (e) => !!e && !['0', 'false', 'off'].includes(String(e.a['w:val'] == null ? 'true' : e.a['w:val']));
  const part = (target, pPr, rPr, tblPr, tcPr) => {
    if (tblPr) {
      const b = bordersFromEl(w(tblPr, 'tblBorders'), kids); if (b) target.b = { ...(target.b || {}), ...b };
      const m = marginsFromEl(w(tblPr, 'tblCellMar'), kids); if (m) target.cm = { ...(target.cm || {}), ...m };
    }
    if (tcPr) {
      const b = bordersFromEl(w(tcPr, 'tcBorders'), kids); if (b) target.b = { ...(target.b || {}), ...b };
      const shd = w(tcPr, 'shd'); const f = shd && shd.a['w:fill'];
      if (f && f !== 'auto') target.fill = f.toUpperCase(); else if (shd && f === 'auto') target.fill = null;
    }
    if (rPr) {
      if (w(rPr, 'b')) target.bold = on(w(rPr, 'b'));
      if (w(rPr, 'i')) target.italic = on(w(rPr, 'i'));
      if (w(rPr, 'caps')) target.caps = on(w(rPr, 'caps'));
      const c = w(rPr, 'color'); if (c && c.a['w:val'] && c.a['w:val'] !== 'auto') target.color = c.a['w:val'].toUpperCase();
    }
    if (pPr) {
      const sp = w(pPr, 'spacing');
      if (sp && sp.a['w:after'] != null) target.pAfter = Number(sp.a['w:after']) / 20;
      if (sp && sp.a['w:line'] && (sp.a['w:lineRule'] || 'auto') === 'auto') target.pLine = Number(sp.a['w:line']) / 240;
      const jc = w(pPr, 'jc'); if (jc) target.align = jc.a['w:val'];
    }
  };
  part(spec, w(el, 'pPr'), w(el, 'rPr'), w(el, 'tblPr'), w(el, 'tcPr'));
  const rb = w(w(el, 'tblPr'), 'tblStyleRowBandSize'); if (rb) spec.rowBand = Number(rb.a['w:val']) || 1;
  for (const c of kids(el, 'w:tblStylePr')) {
    const type = c.a['w:type'];
    if (!COND.includes(type)) continue;
    if (type === 'wholeTable') { part(spec, w(c, 'pPr'), w(c, 'rPr'), w(c, 'tblPr'), w(c, 'tcPr')); continue; }
    const t = spec.cond[type] || (spec.cond[type] = {});
    part(t, w(c, 'pPr'), w(c, 'rPr'), w(c, 'tblPr'), w(c, 'tcPr'));
  }
  return spec;
}

/* ---------------- Word's built-in table styles (the gallery) ---------------- */
// Office theme 2013–2022 accents
export const ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const single = (c, sz = 4) => `single:${sz}:${c}:0`;
/** gallery definitions: id → { name, spec } (spec in the same shape specFromStyle produces) */
function builtins() {
  const out = {};
  const add = (id, name, spec, group) => { out[id] = { id, name, spec: { cond: {}, pAfter: 0, pLine: 1, ...spec }, group }; };
  const all = (c, sz = 4) => ({ top: single(c, sz), left: single(c, sz), bottom: single(c, sz), right: single(c, sz), insideH: single(c, sz), insideV: single(c, sz) });
  add('TableGrid', 'Table Grid', { b: all('auto') }, 'plain');
  add('TableGridLight', 'Grid Table Light', { b: all('BFBFBF') }, 'plain');
  add('PlainTable1', 'Plain Table 1', { b: all('BFBFBF'), cond: { firstRow: { bold: true }, lastRow: { bold: true, b: { top: 'double:4:BFBFBF:0' } }, firstCol: { bold: true }, lastCol: { bold: true }, band1Horz: { fill: 'F2F2F2' }, band1Vert: { fill: 'F2F2F2' } } }, 'plain');
  add('PlainTable2', 'Plain Table 2', { b: { top: single('7F7F7F'), bottom: single('7F7F7F') }, cond: { firstRow: { bold: true, b: { bottom: single('7F7F7F') } }, lastRow: { bold: true, b: { top: single('7F7F7F') } }, firstCol: { bold: true }, lastCol: { bold: true }, band1Horz: { b: { top: single('7F7F7F'), bottom: single('7F7F7F') } }, band1Vert: { b: { left: single('7F7F7F'), right: single('7F7F7F') } } } }, 'plain');
  add('PlainTable3', 'Plain Table 3', { cond: { firstRow: { bold: true, caps: true, b: { bottom: single('7F7F7F') } }, lastRow: { bold: true, caps: true }, firstCol: { bold: true, caps: true, b: { right: single('7F7F7F') } }, lastCol: { bold: true, caps: true }, band1Horz: { fill: 'F2F2F2' }, band1Vert: { fill: 'F2F2F2' } } }, 'plain');
  add('PlainTable4', 'Plain Table 4', { cond: { firstRow: { bold: true }, lastRow: { bold: true }, firstCol: { bold: true }, lastCol: { bold: true }, band1Horz: { fill: 'F2F2F2' }, band1Vert: { fill: 'F2F2F2' } } }, 'plain');
  const fam = (n) => {
    const a = n ? ACCENTS[n - 1] : '000000';
    const sfx = n ? `-Accent${n}` : '';
    const nm = n ? ` Accent ${n}` : '';
    const t20 = n ? tint(a, 0.2) : 'CCCCCC'; const t40 = n ? tint(a, 0.4) : '999999'; const t60 = n ? tint(a, 0.6) : '666666';
    add(`GridTable1Light${sfx}`, `Grid Table 1 Light${nm}`, { b: all(n ? t40 : '999999'), cond: { firstRow: { bold: true, b: { bottom: single(n ? t60 : '666666', 12) } }, lastRow: { bold: true, b: { top: `double:2:${n ? t60 : '666666'}:0` } }, firstCol: { bold: true }, lastCol: { bold: true } } }, 'grid');
    add(`GridTable4${sfx}`, `Grid Table 4${nm}`, { b: all(t60), cond: { firstRow: { bold: true, color: 'FFFFFF', fill: a, b: { top: single(a), left: single(a), bottom: single(a), right: single(a), insideH: 'nil', insideV: 'nil' } }, lastRow: { bold: true, b: { top: `double:4:${a}:0` } }, firstCol: { bold: true }, lastCol: { bold: true }, band1Vert: { fill: t20 }, band1Horz: { fill: t20 } } }, 'grid');
    add(`GridTable5Dark${sfx}`, `Grid Table 5 Dark${nm}`, { b: all('FFFFFF'), fill: t20, cond: { firstRow: { bold: true, color: 'FFFFFF', fill: a }, lastRow: { bold: true, color: 'FFFFFF', fill: a }, firstCol: { bold: true, color: 'FFFFFF', fill: a }, lastCol: { bold: true, color: 'FFFFFF', fill: a }, band1Vert: { fill: t40 }, band1Horz: { fill: t40 } } }, 'grid');
    add(`ListTable3${sfx}`, `List Table 3${nm}`, { b: { top: single(a), left: single(a), bottom: single(a), right: single(a) }, cond: { firstRow: { bold: true, color: 'FFFFFF', fill: a }, lastRow: { bold: true, b: { top: single(a) } }, firstCol: { bold: true }, lastCol: { bold: true }, band1Horz: { b: { top: single(a), bottom: single(a) } }, band1Vert: { b: { left: single(a), right: single(a) } } } }, 'list');
    add(`ListTable4${sfx}`, `List Table 4${nm}`, { b: { top: single(t60), left: single(t60), bottom: single(t60), right: single(t60), insideH: single(t60) }, cond: { firstRow: { bold: true, color: 'FFFFFF', fill: a, b: { top: single(a), left: single(a), bottom: single(a), right: single(a), insideH: 'nil' } }, lastRow: { bold: true, b: { top: `double:4:${t60}:0` } }, firstCol: { bold: true }, lastCol: { bold: true }, band1Vert: { fill: t20 }, band1Horz: { fill: t20 } } }, 'list');
  };
  for (let n = 0; n <= 6; n++) fam(n);
  return out;
}
let BUILTIN = null;
export function builtinTableStyles() { if (!BUILTIN) BUILTIN = builtins(); return BUILTIN; }
/** a built-in style id by its (English) name, case-insensitive */
export function builtinByName(name) { const n = String(name || '').toLowerCase(); return Object.values(builtinTableStyles()).find((s) => s.name.toLowerCase() === n) || null; }

/** spec → w:style XML for styles.xml (written when a document uses a built-in style it does not define yet) */
export function tableStyleXml(def, baseId = null) {
  const s = def.spec;
  const rPr = (t) => `${t.bold ? '<w:b/><w:bCs/>' : ''}${t.italic ? '<w:i/><w:iCs/>' : ''}${t.caps ? '<w:caps/>' : ''}${t.color ? `<w:color w:val="${t.color}"/>` : ''}`;
  const tcPr = (t) => `${bordersXml('tcBorders', t.b, TC_BORDER_ORDER)}${t.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${t.fill}"/>` : ''}`;
  let x = `<w:style w:type="table" w:styleId="${def.id}"><w:name w:val="${def.name}"/>${baseId ? `<w:basedOn w:val="${baseId}"/>` : ''}<w:uiPriority w:val="${def.id === 'TableGrid' ? 39 : 49}"/>`;
  x += `<w:pPr><w:spacing w:after="${Math.round((s.pAfter || 0) * 20)}" w:line="${Math.round((s.pLine || 1) * 240)}" w:lineRule="auto"/></w:pPr>`;
  x += `<w:tblPr>${def.id !== 'TableGrid' ? '<w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>' : ''}${bordersXml('tblBorders', s.b, TBL_BORDER_ORDER)}</w:tblPr>`;
  if (s.fill) x += `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${s.fill}"/></w:tcPr>`;
  for (const type of ['firstRow', 'lastRow', 'firstCol', 'lastCol', 'band1Vert', 'band2Vert', 'band1Horz', 'band2Horz']) {
    const t = s.cond[type];
    if (!t) continue;
    const r = rPr(t); const c = tcPr(t);
    x += `<w:tblStylePr w:type="${type}">${r ? `<w:rPr>${r}</w:rPr>` : ''}${c ? `<w:tblPr/><w:tcPr>${c}</w:tcPr>` : ''}</w:tblStylePr>`;
  }
  return `${x}</w:style>`;
}

/* ---------------- display: CSS variables + classes for the <table> ---------------- */
const pad = (v) => `${Math.round(((Number(v) || 0) / 20) * (4 / 3) * 10) / 10}px`;
/**
 * table attrs → { cls: [..], vars: { '--x': value } } – the CSS in style.css (.of-wt) turns them into Word's look:
 * outside/inside borders, cell margins, table-style conditional formats (header row, banded rows, first column …)
 */
export function tableLook(a = {}) {
  const spec = a.tstyle || null;
  if (!spec && !a.borders && !a.cellMargins && !a.styleId) return null;
  const vars = {}; const cls = ['of-wt'];
  const b = { ...((spec && spec.b) || {}), ...(a.borders || {}) };
  const side = { top: '--tb-t', left: '--tb-l', bottom: '--tb-b', right: '--tb-r', insideH: '--tb-ih', insideV: '--tb-iv' };
  for (const [k, v] of Object.entries(side)) vars[v] = borderCss(b[k]);
  const cm = { top: 0, left: 108, bottom: 0, right: 108, ...((spec && spec.cm) || {}), ...(a.cellMargins || {}) };
  vars['--tc-p'] = `${pad(cm.top)} ${pad(cm.right)} ${pad(cm.bottom)} ${pad(cm.left)}`;
  if (spec) {
    if (spec.pAfter != null) vars['--tp-after'] = `${spec.pAfter}pt`;
    if (spec.pLine != null) vars['--tp-lh'] = String(spec.pLine);
    if (spec.fill) { vars['--tf'] = `#${spec.fill}`; cls.push('of-tf'); }
    if (spec.color) { vars['--tcol'] = `#${spec.color}`; cls.push('of-tcol'); }
    const look = { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true, ...(a.look || {}) };
    const c = spec.cond || {};
    const cond = (key, on, prefix) => {
      const t = c[key];
      if (!t || !on) return;
      if (t.fill) { vars[`--${prefix}-bg`] = `#${t.fill}`; cls.push(`of-${prefix}-bg`); }
      if (t.color) { vars[`--${prefix}-c`] = `#${t.color}`; cls.push(`of-${prefix}-c`); }
      if (t.bold != null) { vars[`--${prefix}-w`] = t.bold ? '700' : '400'; cls.push(`of-${prefix}-w`); }
      if (t.italic) cls.push(`of-${prefix}-i`);
      if (t.caps) cls.push(`of-${prefix}-caps`);
      if (t.b) for (const [k, v] of Object.entries(t.b)) { if (!['top', 'bottom', 'left', 'right'].includes(k)) continue; vars[`--${prefix}-b${k[0]}`] = borderCss(v); cls.push(`of-${prefix}-b${k[0]}`); }
      // header row: the inside borders inside it
      if (t.b && (t.b.insideV || t.b.insideH)) { vars[`--${prefix}-iv`] = borderCss(t.b.insideV); cls.push(`of-${prefix}-iv`); }
    };
    cond('band1Horz', !look.noHBand, 'bh');
    cond('band1Vert', !look.noVBand, 'bv');
    cond('firstCol', look.firstColumn, 'fc');
    cond('lastCol', look.lastColumn, 'lc');
    cond('lastRow', look.lastRow, 'lr');
    cond('firstRow', look.firstRow, 'fr');
    if (look.firstRow) cls.push('of-lk-fr');
  }
  return { cls, vars };
}

/* ---------------- widths ---------------- */
/**
 * how the table is laid out in CSS (Word semantics):
 *   fixed layout          → exact column widths (grid), table width = tblW (dxa) or the grid sum; may run into the margin
 *   autofit + tblW pct    → the given share of the text width, columns in proportion to the grid
 *   autofit (auto / dxa)  → the grid widths, shrunk in proportion when they are wider than the text width
 * cols = column widths in px (grid), null entries = unknown. Returns { table: {width, ...}, cols: [css width] }.
 */
export function tableWidthCss(a = {}, cols = []) {
  const known = cols.filter((w) => w > 0);
  const total = known.reduce((s, w) => s + w, 0);
  const all = cols.length && known.length === cols.length;
  const w = a.width || null;
  const out = { table: {}, cols: cols.map((c) => (c > 0 ? `${c}px` : '')) };
  const indent = a.indent ? Math.round((a.indent / 20) * (4 / 3)) : 0;
  if (indent) out.table['margin-left'] = `${indent}px`;
  if (a.align === 'center') { out.table['margin-left'] = 'auto'; out.table['margin-right'] = 'auto'; } else if (a.align === 'right') out.table['margin-left'] = 'auto';
  if (a.layout === 'fixed') {
    const tw = w && w.type === 'dxa' && w.w > 0 ? Math.round((w.w / 20) * (4 / 3)) : w && w.type === 'pct' && w.w > 0 ? null : (all ? total : null);
    if (w && w.type === 'pct' && w.w > 0) { out.table.width = `${pct(w.w)}%`; if (all) out.cols = cols.map((c) => `${Math.round((c / total) * 10000) / 100}%`); } else if (tw) { out.table.width = `${tw}px`; out.table['min-width'] = `${tw}px`; }
    out.table['table-layout'] = 'fixed';
    return out;
  }
  if (w && w.type === 'pct' && w.w > 0) {
    out.table.width = `${pct(w.w)}%`;
    if (all) out.cols = cols.map((c) => `${Math.round((c / total) * 10000) / 100}%`);
    out.table['table-layout'] = 'fixed';
    return out;
  }
  // autofit: the grid as stored by Word, shrunk to the text width when wider
  const pref = w && w.type === 'dxa' && w.w > 0 ? Math.round((w.w / 20) * (4 / 3)) : total;
  if (all && pref) {
    out.table.width = `min(${pref}px, 100%)`;
    out.cols = cols.map((c) => `${Math.round((c / total) * 10000) / 100}%`);
    out.table['table-layout'] = 'fixed';
  }
  return out;
}
const pct = (v) => { const s = String(v); return s.endsWith('%') ? Number(s.slice(0, -1)) : Math.round(Number(v) / 50 * 100) / 100; };

/** grid columns (px) of a table node (first row, colspans spread) */
export function gridOf(node) {
  const first = node && node.content && node.content[0];
  const cols = [];
  for (const c of (first && first.content) || []) {
    const a = c.attrs || {};
    for (let k = 0; k < (a.colspan || 1); k++) cols.push((a.colwidth && a.colwidth[k]) || null);
  }
  return cols;
}

/** CSS of one cell from its attrs (vertical alignment, own borders, own margins, text direction) */
export function cellCss(a = {}) {
  const s = [];
  if (a.vAlign) s.push(`vertical-align:${a.vAlign === 'center' ? 'middle' : a.vAlign === 'bottom' ? 'bottom' : 'top'}`);
  if (a.borders) for (const k of ['top', 'left', 'bottom', 'right']) if (a.borders[k] != null) s.push(`border-${k}:${borderCss(a.borders[k])}`);
  if (a.margins) for (const k of ['top', 'left', 'bottom', 'right']) if (a.margins[k] != null) s.push(`padding-${k}:${pad(a.margins[k])}`);
  if (a.textDir === 'btLr') s.push('writing-mode:vertical-rl;transform:rotate(180deg)');
  else if (a.textDir === 'tbRl') s.push('writing-mode:vertical-rl');
  if (a.noWrap) s.push('white-space:nowrap');
  return s.join(';');
}
/** CSS of one row (height exact / at least) */
export function rowCss(a = {}) {
  if (!a.height) return '';
  const px = Math.round((a.height / 20) * (4 / 3));
  return a.heightRule === 'exact' ? `height:${px}px;max-height:${px}px` : `height:${px}px`;
}

/**
 * the CSS behind tableLook() – `p` = the scope prefix ('.of-docs-pm ' in the editor, '' in the HTML/PDF export).
 * Order = Word's priority: whole table < banded rows/columns < first/last column < last row < header row.
 */
export function wordTableCss(p = '') {
  const T = `${p}table.of-wt`; const C = `${T} > tbody > tr`;
  const cond = (cls, sel, prop, v) => `${T}.${cls} > tbody > ${sel} { ${prop}: var(${v}); }`;
  return [
    `${T} { width: auto; margin: 0; overflow: visible; border-collapse: collapse; }`,
    `${C} > * { border: 0; border-top: var(--tb-ih); border-bottom: var(--tb-ih); border-left: var(--tb-iv); border-right: var(--tb-iv); padding: var(--tc-p); background: var(--of-paper, #fff); font-weight: inherit; overflow-wrap: anywhere; text-align: inherit; }`,
    `${C} > * > p, ${C} > * > p:last-child { margin-bottom: var(--tp-after, var(--of-after, 8pt)); }`,
    `${C} > * > :is(p, h1, h2, h3, h4, h5, h6) { line-height: calc(var(--tp-lh) * var(--of-lf, 1)); }`,
    `${C}:first-child > * { border-top: var(--tb-t); }`,
    `${C}:last-child > * { border-bottom: var(--tb-b); }`,
    `${C} > :first-child { border-left: var(--tb-l); }`,
    `${C} > :last-child { border-right: var(--tb-r); }`,
    `${T}.of-tf > tbody > tr > * { background: var(--tf); }`,
    `${T}.of-tcol > tbody > tr > * { color: var(--tcol); }`,
    `${T}.of-lk-fr.of-bh-bg > tbody > tr:nth-child(even) > *, ${T}:not(.of-lk-fr).of-bh-bg > tbody > tr:nth-child(odd) > * { background: var(--bh-bg); }`,
    `${T}.of-lk-fr.of-bh-bt > tbody > tr:nth-child(even) > *, ${T}:not(.of-lk-fr).of-bh-bt > tbody > tr:nth-child(odd) > * { border-top: var(--bh-bt); }`,
    `${T}.of-lk-fr.of-bh-bb > tbody > tr:nth-child(even) > *, ${T}:not(.of-lk-fr).of-bh-bb > tbody > tr:nth-child(odd) > * { border-bottom: var(--bh-bb); }`,
    cond('of-bv-bg', 'tr > :nth-child(even)', 'background', '--bv-bg'),
    cond('of-fc-bg', 'tr > :first-child', 'background', '--fc-bg'),
    cond('of-fc-c', 'tr > :first-child', 'color', '--fc-c'),
    cond('of-fc-w', 'tr > :first-child', 'font-weight', '--fc-w'),
    `${T}.of-fc-caps > tbody > tr > :first-child { text-transform: uppercase; }`,
    cond('of-fc-br', 'tr > :first-child', 'border-right', '--fc-br'),
    cond('of-lc-bg', 'tr > :last-child', 'background', '--lc-bg'),
    cond('of-lc-c', 'tr > :last-child', 'color', '--lc-c'),
    cond('of-lc-w', 'tr > :last-child', 'font-weight', '--lc-w'),
    cond('of-lr-bg', 'tr:last-child > *', 'background', '--lr-bg'),
    cond('of-lr-c', 'tr:last-child > *', 'color', '--lr-c'),
    cond('of-lr-w', 'tr:last-child > *', 'font-weight', '--lr-w'),
    cond('of-lr-bt', 'tr:last-child > *', 'border-top', '--lr-bt'),
    cond('of-fr-bg', 'tr:first-child > *', 'background', '--fr-bg'),
    cond('of-fr-c', 'tr:first-child > *', 'color', '--fr-c'),
    cond('of-fr-w', 'tr:first-child > *', 'font-weight', '--fr-w'),
    `${T}.of-fr-i > tbody > tr:first-child > * { font-style: italic; }`,
    `${T}.of-fr-caps > tbody > tr:first-child > * { text-transform: uppercase; }`,
    `${T}.of-fr-iv > tbody > tr:first-child > * { border-left: var(--fr-iv); border-right: var(--fr-iv); }`,
    cond('of-fr-bt', 'tr:first-child > *', 'border-top', '--fr-bt'),
    cond('of-fr-bb', 'tr:first-child > *', 'border-bottom', '--fr-bb'),
    cond('of-fr-bl', 'tr:first-child > :first-child', 'border-left', '--fr-bl'),
    cond('of-fr-br', 'tr:first-child > :last-child', 'border-right', '--fr-br'),
  ].join('\n');
}

/** look flags from a w:tblLook element (both the attribute and the old hex form) */
export function lookFromEl(el) {
  if (!el) return null;
  const v = parseInt(el.a['w:val'] || '0', 16) || 0;
  const flag = (k, bit) => (el.a[`w:${k}`] != null ? ['1', 'true', 'on'].includes(String(el.a[`w:${k}`])) : !!(v & bit));
  return { firstRow: flag('firstRow', 0x20), lastRow: flag('lastRow', 0x40), firstColumn: flag('firstColumn', 0x80), lastColumn: flag('lastColumn', 0x100), noHBand: flag('noHBand', 0x200), noVBand: flag('noVBand', 0x400) };
}
export function lookXml(l) {
  if (!l) return '';
  let v = 0; const bits = { firstRow: 0x20, lastRow: 0x40, firstColumn: 0x80, lastColumn: 0x100, noHBand: 0x200, noVBand: 0x400 };
  for (const k of LOOK_KEYS) if (l[k]) v |= bits[k];
  return `<w:tblLook w:val="${v.toString(16).toUpperCase().padStart(4, '0')}" ${LOOK_KEYS.map((k) => `w:${k}="${l[k] ? 1 : 0}"`).join(' ')}/>`;
}
