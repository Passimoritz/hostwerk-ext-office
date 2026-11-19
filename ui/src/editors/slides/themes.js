/*
 * Themes (colours, fonts, background, decoration) and layouts of the slides chunk. Pure – no React, runs in Node.
 *
 * Colour strings used everywhere in the model:
 *   '#RRGGBB' | '#RRGGBBAA'                      explicit
 *   'accent1'…'accent6','dk1','lt1','dk2','lt2','tx1','bg1','tx2','bg2','hlink','folHlink'   theme colour
 *   optional modifiers: '+40' lighter by 40 % (OOXML lumMod 60000 lumOff 40000), '-25' darker by 25 % (lumMod 75000),
 *   '@50' alpha 50 %.  Example: 'accent1+40@50'
 */
import { EMU_PER_PT } from './units.js';

export const SCHEME_KEYS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

const hex2 = (n) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0');
export function parseHex(c) {
  let s = String(c || '').replace('#', '').trim();
  if (s.length === 3) s = s.split('').map((x) => x + x).join('');
  if (!/^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(s)) return null;
  return { r: parseInt(s.slice(0, 2), 16), g: parseInt(s.slice(2, 4), 16), b: parseInt(s.slice(4, 6), 16), a: s.length === 8 ? parseInt(s.slice(6, 8), 16) / 255 : 1 };
}
export const toHex = ({ r, g, b }) => `#${hex2(r)}${hex2(g)}${hex2(b)}`.toUpperCase();

export function rgbToHsl({ r, g, b }) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b); const min = Math.min(r, g, b);
  let h = 0; let s = 0; const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  return { h, s, l };
}
export function hslToRgb({ h, s, l }) {
  if (s === 0) return { r: l * 255, g: l * 255, b: l * 255 };
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q;
  return { r: f(p, q, h + 1 / 3) * 255, g: f(p, q, h) * 255, b: f(p, q, h - 1 / 3) * 255 };
}
/** OOXML luminance modulation: lumMod/lumOff in 0..1 */
export function lumMod(hex, mod = 1, off = 0) {
  const rgb = parseHex(hex); if (!rgb) return hex;
  const hsl = rgbToHsl(rgb);
  hsl.l = Math.max(0, Math.min(1, hsl.l * mod + off));
  return toHex(hslToRgb(hsl));
}
export function mixHex(a, b, t) {
  const x = parseHex(a); const y = parseHex(b); if (!x || !y) return a;
  return toHex({ r: x.r + (y.r - x.r) * t, g: x.g + (y.g - x.g) * t, b: x.b + (y.b - x.b) * t });
}
export function shade(hex, f) { const x = parseHex(hex); if (!x) return hex; return toHex({ r: x.r * f, g: x.g * f, b: x.b * f }); }
export function tint(hex, f) { const x = parseHex(hex); if (!x) return hex; return toHex({ r: x.r + (255 - x.r) * (1 - f), g: x.g + (255 - x.g) * (1 - f), b: x.b + (255 - x.b) * (1 - f) }); }

const COLOR_RE = /^([a-zA-Z]+[0-9]?)([+-]\d+)?(@\d+)?$/;
/** split a colour ref into { base, mod, alpha } (base: scheme key or hex) */
export function parseColorRef(c) {
  if (!c) return null;
  const s = String(c);
  if (s[0] === '#') {
    const m = /^(#[0-9a-fA-F]{6,8})([+-]\d+)?(@\d+)?$/.exec(s);
    if (!m) return { base: s, mod: 0, alpha: 100 };
    return { base: m[1], mod: m[2] ? parseInt(m[2], 10) : 0, alpha: m[3] ? parseInt(m[3].slice(1), 10) : 100 };
  }
  const m = COLOR_RE.exec(s);
  if (!m) return { base: s, mod: 0, alpha: 100 };
  return { base: m[1], mod: m[2] ? parseInt(m[2], 10) : 0, alpha: m[3] ? parseInt(m[3].slice(1), 10) : 100 };
}

/** colour map: tx1/bg1/tx2/bg2 → dk/lt (dark themes swap) */
export function mapScheme(key, theme) {
  const dark = theme && theme.dark;
  const m = dark ? { tx1: 'lt1', bg1: 'dk1', tx2: 'lt2', bg2: 'dk2' } : { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
  if (theme && theme.clrMap && theme.clrMap[key]) return theme.clrMap[key];
  return m[key] || key;
}

/** resolve a colour ref to '#RRGGBB' (alpha separately via resolveAlpha) */
export function resolveHex(c, theme) {
  const p = parseColorRef(c);
  if (!p) return null;
  let base = p.base;
  if (base[0] !== '#') {
    const k = mapScheme(base, theme);
    base = (theme && theme.colors && theme.colors[k]) || DEFAULT_COLORS[k] || '#000000';
  }
  let hex = base.length === 9 ? base.slice(0, 7) : base;
  if (p.mod > 0) hex = lumMod(hex, 1 - p.mod / 100, p.mod / 100);
  else if (p.mod < 0) hex = lumMod(hex, 1 + p.mod / 100, 0);
  return hex.toUpperCase();
}
export function resolveAlpha(c) {
  const p = parseColorRef(c);
  if (!p) return 1;
  let a = p.alpha / 100;
  if (p.base[0] === '#' && p.base.length === 9) a *= parseInt(p.base.slice(7, 9), 16) / 255;
  return a;
}
/** CSS colour for a ref (rgba when transparent) */
export function css(c, theme, fallback = 'transparent') {
  if (!c) return fallback;
  const hex = resolveHex(c, theme);
  if (!hex) return fallback;
  const a = resolveAlpha(c);
  if (a >= 0.999) return hex;
  const x = parseHex(hex);
  return `rgba(${x.r},${x.g},${x.b},${+a.toFixed(3)})`;
}
export const withAlpha = (c, pct) => { const p = parseColorRef(c); if (!p) return c; return `${p.base}${p.mod ? (p.mod > 0 ? `+${p.mod}` : p.mod) : ''}${pct < 100 ? `@${Math.round(pct)}` : ''}`; };

export const DEFAULT_COLORS = {
  dk1: '#000000', lt1: '#FFFFFF', dk2: '#44546A', lt2: '#E7E6E6', accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5',
  accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47', hlink: '#0563C1', folHlink: '#954F72',
};

/* ---------------------------------------------------------------- built-in themes */
// Decoration: shapes in fractions of the slide (x, y, w, h), drawn behind the content (master / layout shapes).
const D = (geom, x, y, w, h, fill, extra = {}) => ({ geom, x, y, w, h, fill, ...extra });
const grad = (angle, ...stops) => ({ type: 'grad', angle, stops: stops.map((c, i) => ({ pos: Math.round((i / (stops.length - 1)) * 100), color: c })) });
const solid = (color) => ({ type: 'solid', color });

export const THEMES = [
  {
    id: 'hostwerk', name: 'Hostwerk',
    colors: { dk1: '#151925', lt1: '#FFFFFF', dk2: '#3E4558', lt2: '#EEF0F6', accent1: '#6A46FF', accent2: '#19C7A8', accent3: '#5F7BFF', accent4: '#F5B84A', accent5: '#FF5D6C', accent6: '#2FD39A', hlink: '#5A37F0', folHlink: '#8B6CFF' },
    fonts: { major: 'Segoe UI Semibold', minor: 'Segoe UI' },
    bg: solid('lt1'),
    titleBg: grad(135, '#8B6CFF', '#5F7BFF', '#19C7A8'),
    titleText: 'lt1', titleAlign: 'l', titleColor: 'tx1',
    deco: {
      title: [D('ellipse', 0.72, -0.25, 0.5, 0.9, solid('#FFFFFF@10')), D('ellipse', 0.82, 0.55, 0.3, 0.55, solid('#FFFFFF@8'))],
      section: [D('rect', 0, 0, 0.012, 1, grad(90, 'accent1', 'accent2')), D('ellipse', 0.8, 0.62, 0.32, 0.6, solid('accent1@8'))],
      content: [D('rect', 0.07, 0.205, 0.06, 0.008, grad(0, 'accent1', 'accent2'))],
    },
  },
  {
    id: 'office', name: 'Office',
    colors: { ...DEFAULT_COLORS, dk2: '#0E2841', lt2: '#E8E8E8', accent1: '#156082', accent2: '#E97132', accent3: '#196B24', accent4: '#0F9ED5', accent5: '#A02B93', accent6: '#4EA72E', hlink: '#467886', folHlink: '#96607D' },
    fonts: { major: 'Aptos Display', minor: 'Aptos' },
    bg: solid('lt1'), titleAlign: 'ctr', deco: {},
  },
  {
    id: 'midnight', name: 'Midnight', dark: true,
    colors: { dk1: '#0F172A', lt1: '#F8FAFC', dk2: '#1E293B', lt2: '#CBD5E1', accent1: '#38BDF8', accent2: '#F59E0B', accent3: '#A78BFA', accent4: '#34D399', accent5: '#F472B6', accent6: '#FB7185', hlink: '#7DD3FC', folHlink: '#C4B5FD' },
    fonts: { major: 'Segoe UI Semibold', minor: 'Segoe UI' },
    bg: grad(160, '#0F172A', '#111C33'), titleBg: grad(135, '#0B1222', '#1E293B'), titleAlign: 'l', titleColor: 'tx1',
    deco: {
      title: [D('rect', 0.07, 0.62, 0.12, 0.012, solid('accent1')), D('ellipse', 0.68, -0.3, 0.6, 1.0, solid('accent1@7'))],
      section: [D('rect', 0.07, 0.62, 0.12, 0.012, solid('accent2'))],
      content: [D('rect', 0, 0, 1, 0.008, grad(0, 'accent1', 'accent3'))],
    },
  },
  {
    id: 'slate', name: 'Slate',
    colors: { dk1: '#1F2937', lt1: '#FFFFFF', dk2: '#374151', lt2: '#F3F4F6', accent1: '#2563EB', accent2: '#0EA5E9', accent3: '#10B981', accent4: '#F59E0B', accent5: '#EF4444', accent6: '#8B5CF6', hlink: '#2563EB', folHlink: '#7C3AED' },
    fonts: { major: 'Segoe UI Light', minor: 'Segoe UI' },
    bg: solid('#F8FAFC'), titleBg: solid('#1F2937'), titleText: 'lt1', titleAlign: 'l',
    deco: {
      title: [D('rect', 0, 0, 0.035, 1, solid('accent1'))],
      section: [D('rect', 0, 0, 0.035, 1, solid('accent1'))],
      content: [D('rect', 0, 0, 0.012, 1, solid('accent1'))],
    },
  },
  {
    id: 'forest', name: 'Forest',
    colors: { dk1: '#1B2A22', lt1: '#FFFFFF', dk2: '#2F4A3A', lt2: '#EEF3EC', accent1: '#2E7D4F', accent2: '#A3B18A', accent3: '#D4A373', accent4: '#588157', accent5: '#BC6C25', accent6: '#606C38', hlink: '#2E7D4F', folHlink: '#606C38' },
    fonts: { major: 'Georgia', minor: 'Segoe UI' },
    bg: solid('#FBFCF8'), titleBg: grad(150, '#2E7D4F', '#1B4332'), titleText: 'lt1', titleAlign: 'ctr',
    deco: {
      title: [D('rect', 0.42, 0.6, 0.16, 0.006, solid('#FFFFFF@70'))],
      section: [D('rect', 0, 0.9, 1, 0.1, solid('accent1'))],
      content: [D('rect', 0, 0.965, 1, 0.035, solid('accent1'))],
    },
  },
  {
    id: 'sunset', name: 'Sunset',
    colors: { dk1: '#2B1B2F', lt1: '#FFFFFF', dk2: '#5B2A4E', lt2: '#FFF3EC', accent1: '#F25F5C', accent2: '#FFB347', accent3: '#B5489C', accent4: '#247BA0', accent5: '#70C1B3', accent6: '#FFE066', hlink: '#B5489C', folHlink: '#7A3E65' },
    fonts: { major: 'Trebuchet MS', minor: 'Segoe UI' },
    bg: solid('#FFFBF8'), titleBg: grad(120, '#FFB347', '#F25F5C', '#B5489C'), titleText: 'lt1', titleAlign: 'ctr',
    deco: {
      title: [D('ellipse', -0.15, 0.6, 0.5, 0.8, solid('#FFFFFF@12'))],
      section: [D('rect', 0, 0, 1, 0.012, grad(0, 'accent2', 'accent1', 'accent3'))],
      content: [D('rect', 0, 0, 1, 0.012, grad(0, 'accent2', 'accent1', 'accent3'))],
    },
  },
  {
    id: 'ocean', name: 'Ocean',
    colors: { dk1: '#0B2545', lt1: '#FFFFFF', dk2: '#13315C', lt2: '#EEF4ED', accent1: '#1D7EC2', accent2: '#00A6A6', accent3: '#8DA9C4', accent4: '#F4A259', accent5: '#5BC0EB', accent6: '#9BC53D', hlink: '#1D7EC2', folHlink: '#13315C' },
    fonts: { major: 'Segoe UI Semibold', minor: 'Segoe UI' },
    bg: solid('lt1'), titleBg: grad(180, '#13315C', '#0B2545'), titleText: 'lt1', titleAlign: 'l',
    deco: {
      title: [D('rect', 0, 0.82, 1, 0.18, solid('accent1')), D('rect', 0, 0.8, 1, 0.02, solid('accent2'))],
      section: [D('rect', 0, 0, 1, 0.35, solid('lt2'))],
      content: [D('rect', 0, 0, 1, 0.025, solid('accent1'))],
    },
  },
  {
    id: 'mono', name: 'Mono',
    colors: { dk1: '#111111', lt1: '#FFFFFF', dk2: '#444444', lt2: '#F2F2F2', accent1: '#111111', accent2: '#E63946', accent3: '#6C757D', accent4: '#ADB5BD', accent5: '#343A40', accent6: '#E9ECEF', hlink: '#E63946', folHlink: '#6C757D' },
    fonts: { major: 'Arial Black', minor: 'Arial' },
    bg: solid('lt1'), titleAlign: 'l',
    deco: { title: [D('rect', 0.07, 0.6, 0.08, 0.016, solid('accent2'))], section: [D('rect', 0.07, 0.6, 0.08, 0.016, solid('accent2'))], content: [] },
  },
  {
    id: 'paper', name: 'Paper',
    colors: { dk1: '#2D2A26', lt1: '#FFFFFF', dk2: '#5C544B', lt2: '#F6F1E7', accent1: '#9C6644', accent2: '#7F5539', accent3: '#B08968', accent4: '#DDB892', accent5: '#6B705C', accent6: '#A5A58D', hlink: '#9C6644', folHlink: '#7F5539' },
    fonts: { major: 'Cambria', minor: 'Calibri' },
    bg: solid('#FAF6EE'), titleAlign: 'ctr',
    deco: {
      title: [D('rect', 0.04, 0.06, 0.92, 0.88, null, { line: { color: 'accent1', width: 19050 } })],
      section: [D('rect', 0.04, 0.06, 0.92, 0.88, null, { line: { color: 'accent1', width: 9525 } })],
      content: [D('rect', 0.07, 0.205, 0.86, 0.003, solid('accent3'))],
    },
  },
];

export const themeById = (id) => THEMES.find((t) => t.id === id) || THEMES[0];

/** theme data as stored in the document (no functions, no decoration) */
export function themeData(t) {
  return { id: t.id, name: t.name, dark: !!t.dark, colors: { ...t.colors }, fonts: { ...t.fonts }, bg: t.bg || solid('bg1') };
}

/** resolve '+mj-lt' / '+mn-lt' / '+mj' / '+mn' tokens */
export function resolveFont(f, theme) {
  if (!f) return (theme && theme.fonts && theme.fonts.minor) || 'Calibri';
  if (f === '+mj' || f === '+mj-lt') return (theme && theme.fonts && theme.fonts.major) || 'Calibri Light';
  if (f === '+mn' || f === '+mn-lt') return (theme && theme.fonts && theme.fonts.minor) || 'Calibri';
  return f;
}

/* ---------------------------------------------------------------- layouts */
export const LAYOUT_KINDS = [
  { kind: 'title', name: 'Title slide' },
  { kind: 'titleContent', name: 'Title and content' },
  { kind: 'section', name: 'Section header' },
  { kind: 'twoContent', name: 'Two content' },
  { kind: 'comparison', name: 'Comparison' },
  { kind: 'titleOnly', name: 'Title only' },
  { kind: 'blank', name: 'Blank' },
  { kind: 'pictureCaption', name: 'Picture with caption' },
];

const BUL = ['•', '–', '•', '–', '•', '–', '•', '–', '•'];
/** per-level body styles (PowerPoint-like defaults) */
export function bodyLevels(size0 = 24, { bullets = true, color = 'tx1' } = {}) {
  const out = [];
  for (let i = 0; i < 9; i++) {
    const size = Math.max(12, size0 - Math.min(i, 3) * 2 - (i > 3 ? 0 : 0));
    out.push({ font: '+mn', size, color, bullet: bullets ? { char: BUL[i] } : false, marL: bullets ? 285750 + i * 457200 : i * 457200, indent: bullets ? -285750 : 0, spcBef: i === 0 ? 10 : 5, lnSpc: 100 });
  }
  return out;
}
export const plainLevels = (style) => Array.from({ length: 9 }, (_, i) => ({ font: '+mn', size: 18, color: 'tx1', bullet: false, marL: i * 457200, indent: 0, spcBef: 0, lnSpc: 100, ...style }));

/**
 * Build the layouts of a theme for a slide size. A layout:
 * { id, kind, name, bg: Fill|null, decor: [El], ph: { key: { x, y, w, h, anchor, align, lvl: [9 level styles], prompt } } }
 */
export function buildLayouts(theme, size) {
  const t = typeof theme === 'string' ? themeById(theme) : (THEMES.find((x) => x.id === theme.id) || theme);
  const W = size.w; const H = size.h;
  const R = (x, y, w, h) => ({ x: Math.round(x * W), y: Math.round(y * H), w: Math.round(w * W), h: Math.round(h * H) });
  const ta = t.titleAlign || 'l';
  const tColor = t.titleColor || 'tx1';
  const onTitle = t.titleText || tColor;
  const scale = Math.min(W / 12192000, H / 6858000) || 1; // text sizes follow smaller slides
  const sz = (pt) => Math.max(8, Math.round(pt * Math.max(0.75, scale)));
  const titleLv = (size, color = tColor, align = ta) => plainLevels({ font: '+mj', size: sz(size), color, align, lnSpc: 90 });
  const subLv = (size, color = 'tx2', align = ta) => plainLevels({ font: '+mn', size: sz(size), color, align });
  const bodyLv = (size) => bodyLevels(sz(size)).map((l) => ({ ...l, size: l.size }));
  const footers = {
    dt: { ...R(0.07, 0.925, 0.22, 0.05), anchor: 'ctr', lvl: plainLevels({ size: sz(11), color: 'tx2' }) },
    ftr: { ...R(0.35, 0.925, 0.3, 0.05), anchor: 'ctr', lvl: plainLevels({ size: sz(11), color: 'tx2', align: 'ctr' }) },
    sldNum: { ...R(0.71, 0.925, 0.22, 0.05), anchor: 'ctr', lvl: plainLevels({ size: sz(11), color: 'tx2', align: 'r' }) },
  };
  const deco = (k) => ((t.deco && t.deco[k]) || []).map((d, i) => ({
    id: `deco-${k}-${i}`, type: 'shape', geom: d.geom, ...R(d.x, d.y, d.w, d.h), rot: 0, fill: d.fill || null, line: d.line || null, locked: true, decor: true,
  }));
  const titleX = ta === 'ctr' ? 0.1 : 0.07;
  const L = (kind, ph, extra = {}) => ({ id: `layout-${kind}`, kind, name: (LAYOUT_KINDS.find((k) => k.kind === kind) || {}).name || kind, bg: null, decor: kind === 'blank' || kind === 'pictureCaption' ? [] : deco(kind === 'title' ? 'title' : kind === 'section' ? 'section' : 'content'), ph: { ...ph, ...footers }, ...extra });
  const titlePh = { ...R(0.07, 0.055, 0.86, 0.14), anchor: 'b', lvl: titleLv(36), prompt: 'Click to add title' };
  return [
    L('title', {
      ctrTitle: { ...R(titleX, 0.26, ta === 'ctr' ? 0.8 : 0.78, 0.3), anchor: 'b', lvl: titleLv(52, t.titleBg ? onTitle : tColor), prompt: 'Click to add title' },
      subTitle: { ...R(titleX, 0.6, ta === 'ctr' ? 0.8 : 0.7, 0.16), anchor: 't', lvl: subLv(22, t.titleBg ? (t.titleText ? `${t.titleText}@85` : 'tx2') : 'tx2'), prompt: 'Click to add subtitle' },
    }, { bg: t.titleBg || null }),
    L('titleContent', { title: titlePh, body: { ...R(0.07, 0.25, 0.86, 0.64), anchor: 't', lvl: bodyLv(24), prompt: 'Click to add text' } }),
    L('section', {
      title: { ...R(0.07, 0.3, 0.8, 0.28), anchor: 'b', lvl: titleLv(44), prompt: 'Click to add title' },
      body: { ...R(0.07, 0.65, 0.8, 0.14), anchor: 't', lvl: subLv(20, 'tx2', 'l'), prompt: 'Click to add text' },
    }),
    L('twoContent', { title: titlePh, body: { ...R(0.07, 0.25, 0.42, 0.64), anchor: 't', lvl: bodyLv(22), prompt: 'Click to add text' }, body2: { ...R(0.51, 0.25, 0.42, 0.64), anchor: 't', lvl: bodyLv(22), prompt: 'Click to add text' } }),
    L('comparison', {
      title: titlePh,
      body: { ...R(0.07, 0.24, 0.42, 0.09), anchor: 'b', lvl: plainLevels({ size: sz(22), b: true, color: 'accent1' }), prompt: 'Click to add text' },
      body2: { ...R(0.07, 0.35, 0.42, 0.54), anchor: 't', lvl: bodyLv(20), prompt: 'Click to add text' },
      body3: { ...R(0.51, 0.24, 0.42, 0.09), anchor: 'b', lvl: plainLevels({ size: sz(22), b: true, color: 'accent1' }), prompt: 'Click to add text' },
      body4: { ...R(0.51, 0.35, 0.42, 0.54), anchor: 't', lvl: bodyLv(20), prompt: 'Click to add text' },
    }),
    L('titleOnly', { title: titlePh }),
    L('blank', {}),
    L('pictureCaption', {
      title: { ...R(0.07, 0.1, 0.32, 0.2), anchor: 'b', lvl: titleLv(30), prompt: 'Click to add title' },
      pic: { ...R(0.43, 0.08, 0.5, 0.82), anchor: 't', lvl: plainLevels({ size: sz(20), align: 'ctr', color: 'tx2' }), prompt: 'Click the icon to add a picture' },
      body: { ...R(0.07, 0.33, 0.32, 0.56), anchor: 't', lvl: plainLevels({ size: sz(16), color: 'tx2' }), prompt: 'Click to add text' },
    }),
  ];
}

export const PT = EMU_PER_PT;
