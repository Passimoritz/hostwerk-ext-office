/*
 * Word paragraph and character properties beyond the basics – pure helpers shared by the docx reader/writer, the
 * editor extensions (extensions/wordfmt.js) and the HTML renderer.
 *
 * paragraph/heading attrs: rightIndent (pt), lineRule ('exact'|'atLeast'), lineExact (pt), keepNext, keepLines,
 *   widow (false = widow/orphan control off), outline (0–8 body text level of a normal paragraph), pborders
 *   ({ top, left, bottom, right, between } border strings), pshading (hex), tabs ([{ pos (pt), align, leader }]),
 *   dropCap ({ lines, type: 'drop'|'margin' })
 * textStyle attrs: letterSpacing (pt), scale (%), position (pt, + raised / − lowered), caps ('all'|'small'),
 *   dstrike, hidden, lang (BCP 47)
 */
import { borderCss, bordersFromEl, bordersXml, P_BORDER_ORDER, parseBorder, isNone, borderPx } from './tables.js';

const r2 = (v) => Math.round(v * 100) / 100;
const TAB_ALIGN = { left: 'left', start: 'left', right: 'right', end: 'right', center: 'center', decimal: 'decimal', bar: 'bar', clear: 'clear', num: 'left' };
const LEADER_CHAR = { dot: '.', hyphen: '-', underscore: '_', heavy: '_', middleDot: '·', none: '' };

/** pPr → the extra paragraph attrs (only those present) */
export function paraPropsFromPPr(pPr, { kid, kids }) {
  const o = {};
  if (!pPr) return o;
  const w = (n) => kid(pPr, `w:${n}`);
  const on = (e) => !!e && !['0', 'false', 'off'].includes(String(e.a['w:val'] == null ? 'true' : e.a['w:val']));
  const ind = w('ind');
  if (ind) { const r = ind.a['w:right'] != null ? ind.a['w:right'] : ind.a['w:end']; if (r != null && Number(r)) o.rightIndent = r2(Number(r) / 20); }
  const sp = w('spacing');
  if (sp && sp.a['w:line'] && ['exact', 'atLeast'].includes(sp.a['w:lineRule'])) { o.lineRule = sp.a['w:lineRule']; o.lineExact = r2(Number(sp.a['w:line']) / 20); }
  if (w('keepNext')) o.keepNext = on(w('keepNext')) || null;
  if (w('keepLines')) o.keepLines = on(w('keepLines')) || null;
  if (w('widowControl') && !on(w('widowControl'))) o.widow = false;
  const ol = w('outlineLvl'); if (ol && Number(ol.a['w:val']) < 9) o.outline = Number(ol.a['w:val']);
  const bd = bordersFromEl(w('pBdr'), kids); if (bd) o.pborders = bd;
  const shd = w('shd'); if (shd && shd.a['w:fill'] && !['auto', 'FFFFFF'].includes(shd.a['w:fill'])) o.pshading = `#${shd.a['w:fill']}`;
  const tabs = w('tabs');
  if (tabs) { const list = kids(tabs, 'w:tab').filter((t) => t.a['w:val'] !== 'clear').map((t) => ({ pos: r2(Number(t.a['w:pos'] || 0) / 20), align: TAB_ALIGN[t.a['w:val']] || 'left', leader: t.a['w:leader'] && t.a['w:leader'] !== 'none' ? t.a['w:leader'] : null })); if (list.length) o.tabs = list; }
  const fp = w('framePr');
  if (fp && (fp.a['w:dropCap'] === 'drop' || fp.a['w:dropCap'] === 'margin')) o.dropCap = { type: fp.a['w:dropCap'], lines: Number(fp.a['w:lines'] || 3) };
  return o;
}
/** the pPr children these attrs write (name → xml); undefined values are left alone */
export function paraPropsXml(a = {}) {
  const m = new Map();
  if (a.keepNext) m.set('w:keepNext', '<w:keepNext/>');
  if (a.keepLines) m.set('w:keepLines', '<w:keepLines/>');
  if (a.widow === false) m.set('w:widowControl', '<w:widowControl w:val="0"/>');
  if (a.pborders) m.set('w:pBdr', bordersXml('pBdr', a.pborders, P_BORDER_ORDER));
  if (a.pshading) m.set('w:shd', `<w:shd w:val="clear" w:color="auto" w:fill="${String(a.pshading).replace('#', '').toUpperCase()}"/>`);
  if (a.tabs && a.tabs.length) m.set('w:tabs', `<w:tabs>${[...a.tabs].sort((x, y) => x.pos - y.pos).map((t) => `<w:tab w:val="${t.align === 'left' ? 'left' : t.align}"${t.leader ? ` w:leader="${t.leader}"` : ''} w:pos="${Math.round(t.pos * 20)}"/>`).join('')}</w:tabs>`);
  if (a.outline != null && a.outline < 9) m.set('w:outlineLvl', `<w:outlineLvl w:val="${a.outline}"/>`);
  if (a.dropCap) m.set('w:framePr', `<w:framePr w:dropCap="${a.dropCap.type || 'drop'}" w:lines="${a.dropCap.lines || 3}" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>`);
  return m;
}
/** the pPr children paraPropsXml manages (removed from the kept original when the attrs say otherwise) */
export const PARA_PROP_TAGS = ['w:keepNext', 'w:keepLines', 'w:widowControl', 'w:pBdr', 'w:shd', 'w:tabs', 'w:outlineLvl', 'w:framePr'];

/** CSS for the editor/export */
/**
 * paragraph borders like Word: each border sits `space` pt away from the text (padding) and, left and right, OUTSIDE the
 * indent – the text keeps its indent, the border and the shading reach into the margin (Word: indent − space − width).
 * → { 'border-left': …, 'padding-left': …, 'margin-left': … } (CSS property → value)
 */
export function paraBoxCss(pborders, indent = 0, rightIndent = 0) {
  const o = {};
  if (!pborders) return o;
  for (const k of ['top', 'left', 'bottom', 'right']) {
    const b = pborders[k];
    if (!b || isNone(b)) continue;
    const sp = parseBorder(b).space || 0;
    o[`border-${k}`] = borderCss(b);
    o[`padding-${k}`] = `${sp}pt`;
    if (k === 'left' || k === 'right') o[`margin-${k}`] = `calc(${r2((k === 'left' ? indent : rightIndent) || 0) - sp}pt - ${borderPx(b)}px)`;
  }
  return o;
}

export function paraCss(a = {}) {
  const s = [];
  if (a.rightIndent) s.push(`margin-right:${a.rightIndent}pt`);
  if (a.lineRule === 'exact' && a.lineExact) s.push(`line-height:${a.lineExact}pt`);
  else if (a.lineRule === 'atLeast' && a.lineExact) s.push(`line-height:max(${a.lineExact}pt, calc(${a.lineHeight || 1} * var(--of-lf, 1) * 1em))`);
  if (a.pborders) for (const [k, v] of Object.entries(paraBoxCss(a.pborders, a.indent, a.rightIndent))) s.push(`${k}:${v}${k.startsWith('margin') ? ' !important' : ''}`);
  if (a.pshading) s.push(`background-color:${a.pshading}`);
  if (a.keepNext) s.push('break-after:avoid');
  if (a.keepLines) s.push('break-inside:avoid');
  return s.join(';');
}

/** rPr → extra textStyle attrs */
export function runPropsFromRPr(rPr, { kid }) {
  const o = {};
  if (!rPr) return o;
  const w = (n) => kid(rPr, `w:${n}`);
  const on = (e) => !!e && !['0', 'false', 'off'].includes(String(e.a['w:val'] == null ? 'true' : e.a['w:val']));
  const sp = w('spacing'); if (sp && Number(sp.a['w:val'])) o.letterSpacing = r2(Number(sp.a['w:val']) / 20);
  const sc = w('w'); if (sc && Number(sc.a['w:val']) && Number(sc.a['w:val']) !== 100) o.scale = Number(sc.a['w:val']);
  const pos = w('position'); if (pos && Number(pos.a['w:val'])) o.position = r2(Number(pos.a['w:val']) / 2);
  if (w('caps') && on(w('caps'))) o.caps = 'all'; else if (w('smallCaps') && on(w('smallCaps'))) o.caps = 'small';
  if (w('dstrike') && on(w('dstrike'))) o.dstrike = true;
  if (w('vanish') && on(w('vanish'))) o.hidden = true;
  const lang = w('lang'); if (lang && lang.a['w:val']) o.lang = lang.a['w:val'];
  return o;
}
export const RUN_PROP_TAGS = ['w:spacing', 'w:w', 'w:position', 'w:caps', 'w:smallCaps', 'w:dstrike', 'w:vanish', 'w:lang'];
export function runPropsXml(a = {}) {
  const m = new Map();
  if (a.letterSpacing) m.set('w:spacing', `<w:spacing w:val="${Math.round(a.letterSpacing * 20)}"/>`);
  if (a.scale && a.scale !== 100) m.set('w:w', `<w:w w:val="${Math.round(a.scale)}"/>`);
  if (a.position) m.set('w:position', `<w:position w:val="${Math.round(a.position * 2)}"/>`);
  if (a.caps === 'all') m.set('w:caps', '<w:caps/>'); else if (a.caps === 'small') m.set('w:smallCaps', '<w:smallCaps/>');
  if (a.dstrike) m.set('w:dstrike', '<w:dstrike/>');
  if (a.hidden) m.set('w:vanish', '<w:vanish/>');
  if (a.lang) m.set('w:lang', `<w:lang w:val="${a.lang}"/>`);
  return m;
}
export function runCss(a = {}) {
  const s = [];
  if (a.letterSpacing) s.push(`letter-spacing:${a.letterSpacing}pt`);
  if (a.scale && a.scale !== 100) s.push(`display:inline-block;transform:scaleX(${a.scale / 100});transform-origin:left`);
  if (a.position) s.push(`position:relative;top:${-a.position}pt`);
  if (a.caps === 'all') s.push('text-transform:uppercase'); else if (a.caps === 'small') s.push('font-variant:small-caps');
  if (a.dstrike) s.push('text-decoration:line-through double');
  if (a.hidden) s.push('opacity:.45;text-decoration:underline dotted');
  return s.join(';');
}

/** a tab leader as the characters Word fills the gap with (plain-text exports) */
export const leaderChar = (l) => LEADER_CHAR[l] || '';
