/* Simple SVG charts (bar/column/line/area/pie/doughnut) from the element's small data table. Slide space px = pt. */
import { createElement as h } from 'react';
import { css } from '../themes.js';
import { fontStack } from './text.js';

const PALETTE = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
const fmt = (v) => (Math.abs(v) >= 1000 ? `${+(v / 1000).toFixed(1)}k` : `${+(+v).toFixed(2)}`);

function niceMax(v) {
  if (v <= 0) return 1;
  const p = 10 ** Math.floor(Math.log10(v));
  const n = v / p;
  return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * p;
}

export default function Chart({ el, theme, w, h: H }) {
  const c = el.chart || {};
  const font = fontStack('+mn', theme);
  const txt = css('tx1', theme, '#333');
  const grid = css('tx1@15', theme, '#ddd');
  const series = (c.series || []).filter((s) => s && Array.isArray(s.values));
  const cats = c.categories || [];
  const colorOf = (s, i) => css(s.color || PALETTE[i % 6], theme);
  const fs = Math.max(8, Math.min(14, H / 22));
  const titleH = c.title ? fs * 2 : fs * 0.5;
  const legendH = c.legend !== false && series.length ? fs * 2 : 0;
  const kids = [];
  if (c.title) kids.push(h('text', { key: 't', x: w / 2, y: fs * 1.4, textAnchor: 'middle', fontSize: fs * 1.25, fontWeight: 600, fill: txt }, c.title));

  if (c.kind === 'pie' || c.kind === 'doughnut') {
    const s = series[0] || { values: [] };
    const vals = s.values.map((v) => Math.max(0, +v || 0));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = w / 2; const cy = titleH + (H - titleH - legendH) / 2; const r = Math.max(4, Math.min(w, H - titleH - legendH) / 2 - 4);
    let a = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a2 = a + (v / total) * Math.PI * 2;
      const large = a2 - a > Math.PI ? 1 : 0;
      const p = v >= total ? `M${cx} ${cy - r}A${r} ${r} 0 1 1 ${cx - 0.01} ${cy - r}Z` : `M${cx} ${cy}L${cx + r * Math.cos(a)} ${cy + r * Math.sin(a)}A${r} ${r} 0 ${large} 1 ${cx + r * Math.cos(a2)} ${cy + r * Math.sin(a2)}Z`;
      kids.push(h('path', { key: `p${i}`, d: p, fill: css((s.colors && s.colors[i]) || PALETTE[i % 6], theme), stroke: css('bg1', theme, '#fff'), strokeWidth: 1 }));
      if (c.labels) { const m = (a + a2) / 2; kids.push(h('text', { key: `l${i}`, x: cx + r * 0.68 * Math.cos(m), y: cy + r * 0.68 * Math.sin(m), textAnchor: 'middle', dominantBaseline: 'middle', fontSize: fs, fill: '#fff' }, `${Math.round((v / total) * 100)}%`)); }
      a = a2;
    });
    if (c.kind === 'doughnut') kids.push(h('circle', { key: 'hole', cx, cy, r: r * 0.5, fill: css('bg1', theme, '#fff') }));
    if (legendH) {
      const step = w / Math.max(1, cats.length);
      cats.forEach((name, i) => {
        const x = step * i + step / 2;
        kids.push(h('rect', { key: `lg${i}`, x: x - fs * 3, y: H - legendH / 2 - fs / 2, width: fs * 0.8, height: fs * 0.8, fill: css((s.colors && s.colors[i]) || PALETTE[i % 6], theme) }));
        kids.push(h('text', { key: `lt${i}`, x: x - fs * 2, y: H - legendH / 2 + fs * 0.3, fontSize: fs, fill: txt }, String(name)));
      });
    }
    return h('svg', { width: w, height: H, viewBox: `0 0 ${w} ${H}`, style: { fontFamily: font, overflow: 'visible', display: 'block' } }, kids);
  }

  const horiz = c.kind === 'bar';
  const stacked = !!c.stacked;
  const n = Math.max(cats.length, ...series.map((s) => s.values.length), 1);
  let max = 0; let min = 0;
  for (let i = 0; i < n; i++) {
    if (stacked && c.kind !== 'line') { let pos = 0; let neg = 0; for (const s of series) { const v = +s.values[i] || 0; if (v > 0) pos += v; else neg += v; } max = Math.max(max, pos); min = Math.min(min, neg); } else for (const s of series) { const v = +s.values[i] || 0; max = Math.max(max, v); min = Math.min(min, v); }
  }
  max = niceMax(max); min = min < 0 ? -niceMax(-min) : 0;
  const axisW = fs * 3.2;
  const L = horiz ? axisW * 1.6 : axisW; const T = titleH + fs * 0.5; const Rr = w - fs; const B = H - legendH - fs * 1.8;
  const pw = Math.max(1, Rr - L); const ph = Math.max(1, B - T);
  const ticks = 5;
  const val = (v) => (horiz ? L + ((v - min) / (max - min)) * pw : B - ((v - min) / (max - min)) * ph);
  for (let k = 0; k <= ticks; k++) {
    const v = min + ((max - min) * k) / ticks;
    if (horiz) { const x = val(v); kids.push(h('line', { key: `g${k}`, x1: x, x2: x, y1: T, y2: B, stroke: grid })); kids.push(h('text', { key: `gt${k}`, x, y: B + fs * 1.2, textAnchor: 'middle', fontSize: fs * 0.9, fill: txt }, fmt(v))); } else { const y = val(v); kids.push(h('line', { key: `g${k}`, x1: L, x2: Rr, y1: y, y2: y, stroke: grid })); kids.push(h('text', { key: `gt${k}`, x: L - fs * 0.4, y: y + fs * 0.35, textAnchor: 'end', fontSize: fs * 0.9, fill: txt }, fmt(v))); }
  }
  const band = (horiz ? ph : pw) / n;
  for (let i = 0; i < n; i++) {
    const label = cats[i] != null ? String(cats[i]) : '';
    if (horiz) kids.push(h('text', { key: `c${i}`, x: L - fs * 0.4, y: T + band * i + band / 2 + fs * 0.35, textAnchor: 'end', fontSize: fs * 0.9, fill: txt }, label));
    else kids.push(h('text', { key: `c${i}`, x: L + band * i + band / 2, y: B + fs * 1.2, textAnchor: 'middle', fontSize: fs * 0.9, fill: txt }, label));
  }
  if (c.kind === 'line' || c.kind === 'area') {
    series.forEach((s, si) => {
      const pts = [];
      for (let i = 0; i < n; i++) pts.push([L + band * i + band / 2, val(+s.values[i] || 0)]);
      const d = pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)} ${p[1].toFixed(1)}`).join('');
      if (c.kind === 'area') kids.push(h('path', { key: `a${si}`, d: `${d}L${pts[pts.length - 1][0]} ${val(Math.max(min, 0))}L${pts[0][0]} ${val(Math.max(min, 0))}Z`, fill: colorOf(s, si), fillOpacity: 0.35 }));
      kids.push(h('path', { key: `s${si}`, d, fill: 'none', stroke: colorOf(s, si), strokeWidth: Math.max(1.5, fs / 5), strokeLinejoin: 'round' }));
      if (c.kind === 'line') pts.forEach((p, i) => kids.push(h('circle', { key: `m${si}-${i}`, cx: p[0], cy: p[1], r: Math.max(2, fs / 4), fill: colorOf(s, si) })));
      if (c.labels) pts.forEach((p, i) => kids.push(h('text', { key: `v${si}-${i}`, x: p[0], y: p[1] - fs * 0.6, textAnchor: 'middle', fontSize: fs * 0.85, fill: txt }, fmt(+s.values[i] || 0))));
    });
  } else {
    const gap = band * 0.25;
    const bw = stacked ? band - gap : (band - gap) / Math.max(1, series.length);
    for (let i = 0; i < n; i++) {
      let pos = 0; let neg = 0;
      series.forEach((s, si) => {
        const v = +s.values[i] || 0;
        let a; let b;
        if (stacked) { if (v >= 0) { a = pos; b = pos + v; pos = b; } else { a = neg + v; b = neg; neg = a; } } else { a = Math.min(0, v); b = Math.max(0, v); }
        const off = band * i + gap / 2 + (stacked ? 0 : bw * si);
        const p1 = val(a); const p2 = val(b);
        const r = horiz ? { x: Math.min(p1, p2), y: T + off, width: Math.abs(p2 - p1), height: bw } : { x: L + off, y: Math.min(p1, p2), width: bw, height: Math.abs(p2 - p1) };
        kids.push(h('rect', { key: `b${i}-${si}`, ...r, fill: colorOf(s, si) }));
        if (c.labels) kids.push(h('text', { key: `bl${i}-${si}`, x: horiz ? r.x + r.width + fs * 0.3 : r.x + r.width / 2, y: horiz ? r.y + r.height / 2 + fs * 0.35 : r.y - fs * 0.3, textAnchor: horiz ? 'start' : 'middle', fontSize: fs * 0.85, fill: txt }, fmt(v)));
      });
    }
  }
  kids.push(h('line', { key: 'ax', x1: horiz ? val(0) : L, x2: horiz ? val(0) : Rr, y1: horiz ? T : val(0), y2: horiz ? B : val(0), stroke: css('tx1@45', theme, '#999') }));
  if (legendH) {
    const step = w / Math.max(1, series.length);
    series.forEach((s, i) => {
      const x = step * i + step / 2;
      kids.push(h('rect', { key: `lg${i}`, x: x - fs * 3, y: H - legendH / 2 - fs / 2, width: fs * 0.8, height: fs * 0.8, fill: colorOf(s, i) }));
      kids.push(h('text', { key: `lt${i}`, x: x - fs * 2, y: H - legendH / 2 + fs * 0.3, fontSize: fs, fill: txt }, String(s.name || '')));
    });
  }
  return h('svg', { width: w, height: H, viewBox: `0 0 ${w} ${H}`, style: { fontFamily: font, overflow: 'visible', display: 'block' } }, kids);
}
