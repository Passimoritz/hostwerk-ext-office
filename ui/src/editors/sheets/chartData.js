/*
 * Chart data helpers (pure, no React) – shared by charts.jsx, the editor and Node tests.
 */

/** Series colours that read well on the white sheet paper (light and dark theme), colour-blind friendly-ish (Okabe-Ito / Tableau mix). */
export const PALETTE = ['#2563eb', '#f59e0b', '#059669', '#dc2626', '#7c3aed', '#0891b2', '#db2777', '#65a30d', '#ea580c', '#475569'];

export const CHART_TYPES = [
  { id: 'column', label: 'Column' },
  { id: 'bar', label: 'Bar' },
  { id: 'line', label: 'Line' },
  { id: 'area', label: 'Area' },
  { id: 'pie', label: 'Pie' },
  { id: 'doughnut', label: 'Doughnut' },
  { id: 'scatter', label: 'Scatter' },
];

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const isText = (v) => typeof v === 'string' && v.trim() !== '';
const isBlank = (v) => v == null || v === '';
const num = (v) => (isNum(v) ? v : null);

/* ---- A1 range refs ("'Sheet 1'!$A$1:$C$5") ---- */
const colNum = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const colName = (c) => { let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
/** "'My sheet'!$A$1:$C$5" → { sheet, r1, c1, r2, c2 } (zero-based) or null */
export function parseRangeRef(ref) {
  const m = /^(?:(?:'((?:[^']|'')+)'|([^!'\s]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(String(ref || '').trim());
  if (!m) return null;
  const sheet = m[1] != null ? m[1].replace(/''/g, "'") : m[2] || null;
  const c1 = colNum(m[3]); const r1 = +m[4] - 1; const c2 = m[5] ? colNum(m[5]) : c1; const r2 = m[6] ? +m[6] - 1 : r1;
  return { sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}
export const quoteSheet = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${String(name).replace(/'/g, "''")}'`);
/** absolute A1 ref with sheet: ({ sheet, r1, c1, r2, c2 }) → "'Sheet 1'!$A$1:$C$5" */
export function formatRangeRef({ sheet, r1, c1, r2, c2 }) {
  const a = `$${colName(c1)}$${r1 + 1}`; const b = `$${colName(c2)}$${r2 + 1}`;
  return `${sheet ? `${quoteSheet(sheet)}!` : ''}${a}${a === b ? '' : `:${b}`}`;
}
const text = (g, r, c) => {
  const t = g.texts && g.texts[r] ? g.texts[r][c] : undefined;
  if (t != null && t !== '') return String(t);
  const v = g.values[r] ? g.values[r][c] : null;
  return v == null ? '' : String(v && v.code ? v.code : v);
};

/** Transpose a { values, texts } grid. */
function transpose(g) {
  const rows = g.values.length; const cols = rows ? Math.max(...g.values.map((r) => r.length)) : 0;
  const tv = []; const tt = [];
  for (let c = 0; c < cols; c++) {
    tv.push([]); tt.push([]);
    for (let r = 0; r < rows; r++) { tv[c].push(g.values[r][c]); tt[c].push(g.texts && g.texts[r] ? g.texts[r][c] : undefined); }
  }
  return { values: tv, texts: tt };
}

const flat = (g) => {
  if (!g) return { vals: [], texts: [] };
  const vals = []; const texts = [];
  g.values.forEach((row, r) => row.forEach((v, c) => { vals.push(v); texts.push(text(g, r, c)); }));
  return { vals, texts };
};

/**
 * chart → { categories: string[], series: [{ name, values: (number|null)[], color }] }
 * getValues(ref) → { values: any[][], texts: string[][] } | null
 */
export function resolveChartData(chart, getValues) {
  const colors = chart.colors && chart.colors.length ? chart.colors : PALETTE;
  const colorOf = (i, s) => (s && s.color) || colors[i % colors.length];
  // explicit series (from xlsx)
  if (chart.series && chart.series.length) {
    let categories = [];
    const series = chart.series.map((s, i) => {
      const vg = s.valRef ? getValues(s.valRef) : null;
      const values = flat(vg).vals.map(num);
      if (!categories.length && s.catRef) {
        const cg = getValues(s.catRef);
        if (cg) {
          const f = flat(cg);
          categories = chart.type === 'scatter' ? f.vals.map((v) => (isNum(v) ? v : null)) : f.texts;
        }
      }
      let name = s.name || '';
      if (s.nameRef) { const ng = getValues(s.nameRef); if (ng) name = flat(ng).texts.join(' ') || name; }
      return { name: name || `Series ${i + 1}`, values, color: colorOf(i, s), refs: { name: s.nameRef || null, cat: s.catRef || null, val: s.valRef || null } };
    });
    const n = Math.max(0, ...series.map((s) => s.values.length));
    if (!categories.length) categories = Array.from({ length: n }, (_, i) => (chart.type === 'scatter' ? i + 1 : String(i + 1)));
    return { categories, series };
  }
  let g = chart.range ? getValues(chart.range) : null;
  if (!g || !g.values || !g.values.length) return { categories: [], series: [] };
  if (chart.byRows) g = transpose(g);
  const V = g.values; const rows = V.length; const cols = Math.max(...V.map((r) => r.length));
  if (!cols) return { categories: [], series: [] };
  const at = (r, c) => (V[r] ? V[r][c] : null);
  // refs back into the sheet (for writing xlsx charts): grid (r, c) → sheet cells, honouring byRows
  const R = parseRangeRef(chart.range);
  const cell = (ra, rb, ca, cb) => (R ? formatRangeRef(chart.byRows ? { sheet: R.sheet, r1: R.r1 + ca, r2: R.r1 + cb, c1: R.c1 + ra, c2: R.c1 + rb } : { sheet: R.sheet, r1: R.r1 + ra, r2: R.r1 + rb, c1: R.c1 + ca, c2: R.c1 + cb }) : null);
  if (chart.type === 'scatter') {
    // header row if first row has text
    let hdr = false;
    for (let c = 0; c < cols; c++) if (isText(at(0, c))) hdr = true;
    const r0 = hdr ? 1 : 0;
    const one = cols === 1;
    const xs = []; for (let r = r0; r < rows; r++) xs.push(one ? r - r0 + 1 : (isNum(at(r, 0)) ? at(r, 0) : null));
    const series = [];
    for (let c = one ? 0 : 1; c < cols; c++) {
      const values = []; for (let r = r0; r < rows; r++) values.push(num(at(r, c)));
      series.push({ name: hdr ? text(g, 0, c) || `Series ${series.length + 1}` : `Series ${series.length + 1}`, values, color: colorOf(series.length),
        refs: { name: hdr ? cell(0, 0, c, c) : null, cat: one || rows <= r0 ? null : cell(r0, rows - 1, 0, 0), val: rows > r0 ? cell(r0, rows - 1, c, c) : null } });
    }
    return { categories: xs, series };
  }
  // header row: text cells beyond the first column in row 0 (or top-left empty and others text)
  let hdr = false;
  for (let c = 1; c < cols; c++) if (isText(at(0, c))) hdr = true;
  if (cols === 1 && isText(at(0, 0)) && rows > 1) hdr = true;
  // category column: first column is text (below the header) or top-left empty
  let catCol = false;
  if (cols > 1) {
    if (hdr && isBlank(at(0, 0))) catCol = true;
    for (let r = hdr ? 1 : 0; r < rows && !catCol; r++) if (isText(at(r, 0))) catCol = true;
  }
  const r0 = hdr ? 1 : 0; const c0 = catCol ? 1 : 0;
  const categories = [];
  for (let r = r0; r < rows; r++) categories.push(catCol ? text(g, r, 0) : String(r - r0 + 1));
  const catRef = catCol && rows > r0 ? cell(r0, rows - 1, 0, 0) : null;
  const series = [];
  for (let c = c0; c < cols; c++) {
    const values = []; for (let r = r0; r < rows; r++) values.push(num(at(r, c)));
    series.push({ name: hdr ? text(g, 0, c) || `Series ${series.length + 1}` : `Series ${series.length + 1}`, values, color: colorOf(series.length),
      refs: { name: hdr ? cell(0, 0, c, c) : null, cat: catRef, val: rows > r0 ? cell(r0, rows - 1, c, c) : null } });
  }
  return { categories, series };
}

/** "Nice" axis ticks (1-2-5 steps) covering [min, max]. */
export function niceTicks(min, max, count = 5) {
  if (!Number.isFinite(min) || !Number.isFinite(max)) { min = 0; max = 1; }
  if (min === max) { if (min === 0) max = 1; else if (min > 0) min = 0; else max = 0; }
  const span = max - min;
  const raw = span / Math.max(1, count);
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) || 10 * mag;
  const lo = Math.floor(min / step + 1e-9) * step; const hi = Math.ceil(max / step - 1e-9) * step;
  const ticks = [];
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Math.abs(v) < step * 1e-9 ? 0 : +v.toPrecision(12));
  return { min: lo, max: hi, step, ticks };
}
