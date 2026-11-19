/*
 * Units of the slides model. Everything is stored in EMU (English Metric Units, like OOXML): integers.
 * Rendering uses 1 CSS px = 1 pt in slide space (EMU / 12700), then scales with CSS transforms.
 */
export const EMU_PER_IN = 914400;
export const EMU_PER_CM = 360000;
export const EMU_PER_MM = 36000;
export const EMU_PER_PT = 12700;
export const EMU_PER_PX = 9525; // 96 dpi pixel

export const emuToPt = (e) => (Number(e) || 0) / EMU_PER_PT;
export const ptToEmu = (p) => Math.round((Number(p) || 0) * EMU_PER_PT);
export const emuToIn = (e) => (Number(e) || 0) / EMU_PER_IN;
export const inToEmu = (i) => Math.round((Number(i) || 0) * EMU_PER_IN);
export const emuToCm = (e) => (Number(e) || 0) / EMU_PER_CM;
export const cmToEmu = (c) => Math.round((Number(c) || 0) * EMU_PER_CM);
export const emuToPx = (e) => (Number(e) || 0) / EMU_PER_PX;
export const pxToEmu = (p) => Math.round((Number(p) || 0) * EMU_PER_PX);
/** slide-space CSS px (1 px = 1 pt) */
export const S = (e) => (Number(e) || 0) / EMU_PER_PT;

/**
 * Parse an ODF/CSS length ("2.5cm", "1in", "12pt", "30mm", "96px", "0.5pc") into EMU. Bare numbers are treated as `def` unit.
 */
export function lengthToEmu(v, def = 'cm') {
  if (v == null || v === '') return 0;
  const m = /^\s*(-?[\d.]+(?:e-?\d+)?)\s*([a-z%]*)\s*$/i.exec(String(v));
  if (!m) return 0;
  const n = parseFloat(m[1]);
  const u = (m[2] || def).toLowerCase();
  const f = { cm: EMU_PER_CM, mm: EMU_PER_MM, in: EMU_PER_IN, inch: EMU_PER_IN, pt: EMU_PER_PT, px: EMU_PER_PX, pc: EMU_PER_PT * 12, emu: 1 }[u];
  return Math.round(n * (f || EMU_PER_CM));
}
export const emuToLength = (e, unit = 'cm') => {
  const f = { cm: EMU_PER_CM, mm: EMU_PER_MM, in: EMU_PER_IN, pt: EMU_PER_PT }[unit] || EMU_PER_CM;
  return `${+((Number(e) || 0) / f).toFixed(4)}${unit}`;
};

/** Standard slide sizes (EMU) */
export const SIZES = {
  '16:9': { w: 12192000, h: 6858000, label: 'Widescreen (16:9)' },
  '4:3': { w: 9144000, h: 6858000, label: 'Standard (4:3)' },
  '16:10': { w: 11430000, h: 7143750, label: 'Widescreen (16:10)' },
  A4: { w: 10692000, h: 7560000, label: 'A4 paper' },
  Letter: { w: 10058400, h: 7772400, label: 'Letter paper' },
};

export function sizePreset(w, h) {
  for (const [k, v] of Object.entries(SIZES)) if (Math.abs(v.w - w) < 20000 && Math.abs(v.h - h) < 20000) return k;
  return 'custom';
}

export const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
export const round = (v) => Math.round(v);
/** angle in OOXML is 60000ths of a degree */
export const angToDeg = (a) => (Number(a) || 0) / 60000;
export const degToAng = (d) => Math.round((((Number(d) || 0) % 360) + 360) % 360 * 60000);
