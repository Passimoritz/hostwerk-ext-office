/*
 * Text styling shared by the renderer (React), the in-place editor (contentEditable HTML) and its parser.
 * Slide space: 1 CSS px = 1 pt; lengths in EMU are divided by 12700.
 */
import { css, resolveFont, resolveHex } from '../themes.js';
import { S } from '../units.js';

export const FONT_FALLBACK = ", 'Segoe UI', system-ui, -apple-system, Arial, sans-serif";
export const fontStack = (f, theme) => {
  const n = resolveFont(f, theme);
  return `'${String(n).replace(/'/g, '')}'${FONT_FALLBACK}`;
};

/** effective level style: element lvl → defaults */
export function levelStyle(body, level, theme) {
  const lv = (body && body.lvl && (body.lvl[level] || body.lvl[body.lvl.length - 1])) || {};
  const st = (body && body.style) || {};
  return { font: '+mn', size: 18, color: 'tx1', marL: level * 457200, indent: 0, lnSpc: 100, spcBef: 0, spcAft: 0, align: 'l', bullet: false, ...st, ...lv };
}

/** effective paragraph props */
export function paraProps(p, body, theme) {
  const lv = levelStyle(body, p.level || 0, theme);
  const pick = (k) => (p[k] !== undefined && p[k] !== null ? p[k] : lv[k]);
  return {
    lv,
    align: pick('align') || 'l',
    lnSpc: pick('lnSpc') || 100,
    spcBef: pick('spcBef') || 0,
    spcAft: pick('spcAft') || 0,
    marL: pick('marL') || 0,
    indent: pick('indent') || 0,
    bullet: p.bullet !== undefined ? p.bullet : lv.bullet,
  };
}

/** effective run props (p.endProps used for empty paragraphs) */
export function runProps(r, p, body) {
  const lv = levelStyle(body, (p && p.level) || 0);
  const o = { font: lv.font, size: lv.size, color: lv.color, b: !!lv.b, i: !!lv.i, u: !!lv.u, s: !!lv.s, caps: lv.caps };
  for (const k of ['font', 'size', 'color', 'b', 'i', 'u', 's', 'highlight', 'baseline', 'link', 'caps', 'spc']) if (r && r[k] !== undefined && r[k] !== null) o[k] = r[k];
  return o;
}

const ALIGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' };

export function paraCss(pp, scale = 1) {
  return {
    textAlign: ALIGN[pp.align] || 'left',
    lineHeight: `${(pp.lnSpc / 100) * 1.2}`,
    marginTop: `${pp.spcBef * scale}px`,
    marginBottom: `${pp.spcAft * scale}px`,
    paddingLeft: `${Math.max(0, S(pp.marL))}px`,
    textIndent: `${S(pp.indent)}px`,
    minHeight: '1em',
    margin: undefined,
  };
}

export function runCss(rp, theme, fontScale = 1) {
  const st = {
    fontFamily: fontStack(rp.font, theme),
    fontSize: `${+(rp.size * fontScale).toFixed(2)}px`,
    color: css(rp.color || 'tx1', theme, '#000'),
  };
  if (rp.b) st.fontWeight = 700;
  if (rp.i) st.fontStyle = 'italic';
  const deco = [rp.u && 'underline', rp.s && 'line-through'].filter(Boolean).join(' ');
  if (deco) st.textDecoration = deco;
  if (rp.highlight) st.backgroundColor = css(rp.highlight, theme);
  if (rp.baseline) { st.verticalAlign = rp.baseline > 0 ? 'super' : 'sub'; st.fontSize = `${+(rp.size * fontScale * 0.65).toFixed(2)}px`; }
  if (rp.caps === 'all') st.textTransform = 'uppercase';
  if (rp.caps === 'small') st.fontVariant = 'small-caps';
  if (rp.spc) st.letterSpacing = `${rp.spc / 100}px`;
  return st;
}

const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
const roman = (n) => { let s = ''; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; };
const alpha = (n) => { let s = ''; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };
export function numLabel(scheme, n) {
  const sc = scheme || 'arabicPeriod';
  let core = String(n);
  if (/^alphaLc/.test(sc)) core = alpha(n); else if (/^alphaUc/.test(sc)) core = alpha(n).toUpperCase();
  else if (/^romanLc/.test(sc)) core = roman(n); else if (/^romanUc/.test(sc)) core = roman(n).toUpperCase();
  if (/ParenBoth$/.test(sc)) return `(${core})`;
  if (/ParenR$/.test(sc)) return `${core})`;
  if (/Plain$/.test(sc)) return core;
  return `${core}.`;
}

/** bullet labels for all paragraphs of a body: [{ text, color, font } | null] */
export function bulletLabels(body, theme) {
  const counters = [];
  return (body.paras || []).map((p) => {
    const pp = paraProps(p, body, theme);
    const lvl = p.level || 0;
    counters.length = lvl + 1;
    const b = pp.bullet;
    if (!b || (!(p.runs || []).some((r) => r.text) && !p.forceBullet)) { if (!b) counters[lvl] = 0; return null; }
    if (b.num) {
      const n = counters[lvl] ? counters[lvl] + 1 : (b.start || 1);
      counters[lvl] = n;
      return { text: numLabel(b.num, n), color: b.color, font: b.font, sizePct: b.sizePct };
    }
    counters[lvl] = 0;
    return { text: b.char || '•', color: b.color, font: b.font, sizePct: b.sizePct };
  });
}

/** text of a run, fields resolved */
export function runText(r, ctx = {}) {
  if (r.field === 'slidenum') return String(ctx.slideNumber || r.text || '#');
  if (r.field === 'datetime') return ctx.date || new Date().toLocaleDateString(ctx.locale || undefined);
  if (r.field === 'footer') return ctx.footer || r.text || '';
  return r.text || '';
}

export const cssColor = (c, theme) => css(c, theme);
export const hexColor = (c, theme) => resolveHex(c, theme);
