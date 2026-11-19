/*
 * SVG chart rendering for the spreadsheet editor. Charts sit on the white sheet "paper" in both themes,
 * so colours are fixed (white background, dark text). Data comes from resolveChartData (chartData.js).
 */
import { PALETTE, CHART_TYPES, resolveChartData, niceTicks } from './chartData.js';
import { Keep } from '../_shared/keep.jsx';

export { PALETTE, CHART_TYPES, resolveChartData, niceTicks };

const TEXT = '#1f2937';
const MUTED = '#6b7280';
const GRID = '#e5e7eb';
const AXIS = '#9ca3af';

const lang = () => (typeof document !== 'undefined' && document.documentElement.lang) || 'en';
const fmtNum = (v, step) => {
  const d = step && step < 1 ? Math.min(6, Math.ceil(-Math.log10(step))) : 0;
  try { return v.toLocaleString(lang(), { maximumFractionDigits: d }); } catch { return String(v); }
};
const pctText = (v) => { try { return (v * 100).toLocaleString(lang(), { maximumFractionDigits: 1 }) + ' %'; } catch { return `${Math.round(v * 100)}%`; } };
/** rough text width (px) for layout */
const tw = (s, fs = 11) => String(s).length * fs * 0.56;
const clip = (s, n) => { s = String(s ?? ''); return s.length > n ? `${s.slice(0, n - 1)}…` : s; };

function Legend({ items, pos, box }) {
  if (pos === 'none' || !items.length) return null;
  const fs = 11; const sw = 10; const gap = 14;
  if (pos === 'right' || pos === 'left') {
    const x = box.x; let y = box.y + Math.max(0, (box.h - items.length * 18) / 2);
    return <g>{items.map((it, i) => { const yy = y + i * 18; return <g key={i}><rect x={x} y={yy} width={sw} height={sw} rx={2} fill={it.color} /><text x={x + sw + 5} y={yy + sw - 1} fontSize={fs} fill={TEXT}><Keep>{clip(it.name, 18)}</Keep></text></g>; })}</g>;
  }
  // top / bottom: one row, centred (wraps by clipping)
  const widths = items.map((it) => sw + 5 + tw(clip(it.name, 18), fs) + gap);
  const total = widths.reduce((a, b) => a + b, 0) - gap;
  let x = box.x + Math.max(0, (box.w - total) / 2);
  return <g>{items.map((it, i) => { const xx = x; x += widths[i]; return <g key={i}><rect x={xx} y={box.y} width={sw} height={sw} rx={2} fill={it.color} /><text x={xx + sw + 5} y={box.y + sw - 1} fontSize={fs} fill={TEXT}><Keep>{clip(it.name, 18)}</Keep></text></g>; })}</g>;
}

function legendSpace(pos, items, width) {
  if (pos === 'none' || !items.length) return { top: 0, right: 0, bottom: 0, left: 0 };
  const w = Math.min(width * 0.35, 20 + Math.max(...items.map((it) => tw(clip(it.name, 18)))));
  return { top: pos === 'top' ? 22 : 0, bottom: pos === 'bottom' ? 22 : 0, right: pos === 'right' ? w + 8 : 0, left: pos === 'left' ? w + 8 : 0 };
}

function Pie({ data, plot, doughnut }) {
  const s = data.series[0];
  const vals = s ? s.values.map((v) => (v > 0 ? v : 0)) : [];
  const sum = vals.reduce((a, b) => a + b, 0);
  if (!sum) return null;
  const cx = plot.x + plot.w / 2; const cy = plot.y + plot.h / 2; const r = Math.max(10, Math.min(plot.w, plot.h) / 2 - 4); const ri = doughnut ? r * 0.55 : 0;
  let a = -Math.PI / 2; const out = [];
  vals.forEach((v, i) => {
    if (!v) return;
    const da = (v / sum) * Math.PI * 2; const a2 = a + da; const color = (data.pointColors && data.pointColors[i]) || PALETTE[i % PALETTE.length];
    const big = da > Math.PI ? 1 : 0;
    const p = (ang, rad) => `${cx + rad * Math.cos(ang)},${cy + rad * Math.sin(ang)}`;
    let d;
    if (da >= Math.PI * 2 - 1e-9) d = ri ? `M${p(0, r)}A${r},${r} 0 1 1 ${p(Math.PI, r)}A${r},${r} 0 1 1 ${p(0, r)}M${p(0, ri)}A${ri},${ri} 0 1 0 ${p(Math.PI, ri)}A${ri},${ri} 0 1 0 ${p(0, ri)}Z` : `M${p(0, r)}A${r},${r} 0 1 1 ${p(Math.PI, r)}A${r},${r} 0 1 1 ${p(0, r)}Z`;
    else d = ri ? `M${p(a, r)}A${r},${r} 0 ${big} 1 ${p(a2, r)}L${p(a2, ri)}A${ri},${ri} 0 ${big} 0 ${p(a, ri)}Z` : `M${cx},${cy}L${p(a, r)}A${r},${r} 0 ${big} 1 ${p(a2, r)}Z`;
    const mid = a + da / 2; const lr = ri ? (r + ri) / 2 : r * 0.65;
    out.push(<g key={i}><path d={d} fill={color} stroke="#fff" strokeWidth={1.5}><title>{`${data.categories[i] ?? ''}: ${fmtNum(vals[i])}`}</title></path>
      {da > 0.25 && <text x={cx + lr * Math.cos(mid)} y={cy + lr * Math.sin(mid) + 4} textAnchor="middle" fontSize={11} fill="#fff" fontWeight={600}>{pctText(v / sum)}</text>}</g>);
    a = a2;
  });
  return <g>{out}</g>;
}

function Cartesian({ chart, data, plot }) {
  const type = chart.type; const horiz = type === 'bar'; const stacked = (type === 'column' || type === 'bar' || type === 'area') && chart.stacked;
  const percent = stacked === 'percent'; const scatter = type === 'scatter';
  const n = scatter ? data.categories.length : Math.max(data.categories.length, ...data.series.map((s) => s.values.length));
  if (!n) return null;
  // value extent
  let lo = Infinity; let hi = -Infinity;
  const posSum = new Array(n).fill(0); const negSum = new Array(n).fill(0);
  data.series.forEach((s) => s.values.forEach((v, i) => { if (v == null) return; if (v >= 0) posSum[i] += v; else negSum[i] += v; lo = Math.min(lo, v); hi = Math.max(hi, v); }));
  if (stacked) { lo = Math.min(0, ...negSum); hi = Math.max(0, ...posSum); }
  if (percent) { lo = negSum.some((v) => v < 0) ? -1 : 0; hi = posSum.some((v) => v > 0) ? 1 : 0; }
  if (lo === Infinity) { lo = 0; hi = 1; }
  if (type !== 'line' && type !== 'scatter') { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
  const vt = niceTicks(lo, hi, Math.max(2, Math.round((horiz ? plot.w : plot.h) / 50)));
  // x for scatter
  let xt = null;
  if (scatter) { const xs = data.categories.filter((v) => typeof v === 'number'); xt = niceTicks(Math.min(...xs, Infinity) === Infinity ? 0 : Math.min(...xs), xs.length ? Math.max(...xs) : 1, Math.max(2, Math.round(plot.w / 70))); }
  const vLabel = (v) => (percent ? pctText(v) : fmtNum(v, vt.step));
  // margins inside plot for tick labels
  const vLabW = Math.max(...vt.ticks.map((v) => tw(vLabel(v)))) + 8;
  const catLabels = data.categories.map((c) => clip(c, 14));
  const catLabW = horiz ? Math.min(plot.w * 0.3, Math.max(...catLabels.map((c) => tw(c))) + 8) : 0;
  const P = { x: plot.x + (horiz ? catLabW : vLabW), y: plot.y + 4, w: 0, h: 0 };
  P.w = Math.max(10, plot.x + plot.w - P.x - 6); P.h = Math.max(10, plot.y + plot.h - P.y - (horiz ? 18 : 18));
  const vScale = (v) => (horiz ? P.x + ((v - vt.min) / (vt.max - vt.min)) * P.w : P.y + P.h - ((v - vt.min) / (vt.max - vt.min)) * P.h);
  const xScale = scatter ? (v) => P.x + ((v - xt.min) / (xt.max - xt.min)) * P.w : null;
  const band = (horiz ? P.h : P.w) / n;
  const catPos = (i) => (horiz ? P.y + band * (i + 0.5) : P.x + band * (i + 0.5));
  const el = [];
  // grid + value ticks
  vt.ticks.forEach((v, i) => {
    const p = vScale(v);
    el.push(horiz ? <line key={`g${i}`} x1={p} x2={p} y1={P.y} y2={P.y + P.h} stroke={v === 0 ? AXIS : GRID} />
      : <line key={`g${i}`} x1={P.x} x2={P.x + P.w} y1={p} y2={p} stroke={v === 0 ? AXIS : GRID} />);
    el.push(horiz ? <text key={`t${i}`} x={p} y={P.y + P.h + 13} fontSize={10} fill={MUTED} textAnchor="middle">{vLabel(v)}</text>
      : <text key={`t${i}`} x={P.x - 5} y={p + 3.5} fontSize={10} fill={MUTED} textAnchor="end">{vLabel(v)}</text>);
  });
  if (scatter) {
    xt.ticks.forEach((v, i) => { const p = xScale(v); el.push(<line key={`xg${i}`} x1={p} x2={p} y1={P.y} y2={P.y + P.h} stroke={GRID} />); el.push(<text key={`xt${i}`} x={p} y={P.y + P.h + 13} fontSize={10} fill={MUTED} textAnchor="middle">{fmtNum(v, xt.step)}</text>); });
  } else {
    // category labels (skip some when crowded)
    const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horiz ? P.h / 14 : P.w / (Math.max(...catLabels.map((c) => tw(c, 10)), 20) + 6))))));
    catLabels.forEach((c, i) => {
      if (i % every) return;
      el.push(horiz ? <text key={`c${i}`} x={P.x - 5} y={catPos(i) + 3.5} fontSize={10} fill={MUTED} textAnchor="end"><Keep>{c}</Keep></text>
        : <text key={`c${i}`} x={catPos(i)} y={P.y + P.h + 13} fontSize={10} fill={MUTED} textAnchor="middle"><Keep>{c}</Keep></text>);
    });
  }
  // axis base lines
  el.push(horiz ? <line key="ax" x1={P.x} x2={P.x} y1={P.y} y2={P.y + P.h} stroke={AXIS} /> : <line key="ax" x1={P.x} x2={P.x + P.w} y1={P.y + P.h} y2={P.y + P.h} stroke={AXIS} />);
  const S = data.series;
  if (type === 'column' || type === 'bar') {
    const inner = band * 0.75; const bw = stacked ? inner : inner / Math.max(1, S.length);
    const base = new Array(n).fill(0).map(() => ({ p: 0, m: 0 }));
    S.forEach((s, si) => s.values.forEach((v0, i) => {
      if (v0 == null) return;
      let v = v0; if (percent) { const tot = posSum[i] - negSum[i]; v = tot ? v0 / tot : 0; }
      let a; let b;
      if (stacked) { if (v >= 0) { a = base[i].p; b = a + v; base[i].p = b; } else { a = base[i].m; b = a + v; base[i].m = b; } } else { a = 0; b = v; }
      const off = catPos(i) - inner / 2 + (stacked ? 0 : si * bw);
      const p1 = vScale(Math.max(vt.min, Math.min(vt.max, a))); const p2 = vScale(Math.max(vt.min, Math.min(vt.max, b)));
      const r = horiz ? { x: Math.min(p1, p2), y: off, width: Math.abs(p2 - p1), height: Math.max(1, bw - 1) } : { x: off, y: Math.min(p1, p2), width: Math.max(1, bw - 1), height: Math.abs(p2 - p1) };
      el.push(<rect key={`b${si}-${i}`} {...r} fill={s.color}><title>{`${s.name} · ${data.categories[i] ?? ''}: ${fmtNum(v0)}`}</title></rect>);
    }));
  } else if (type === 'line' || type === 'area') {
    const base = new Array(n).fill(0);
    const paths = [];
    S.forEach((s, si) => {
      const pts = [];
      for (let i = 0; i < n; i++) {
        let v = s.values[i];
        if (v == null) { if (!stacked) { pts.push(null); continue; } v = 0; }
        if (percent) { const tot = posSum[i] - negSum[i]; v = tot ? v / tot : 0; }
        const b0 = base[i]; if (stacked) base[i] += v;
        pts.push({ x: catPos(i), y: vScale(stacked ? b0 + v : v), y0: vScale(stacked ? b0 : Math.max(vt.min, Math.min(vt.max, 0))), v: s.values[i] });
      }
      paths.push({ s, si, pts });
    });
    const order = type === 'area' ? [...paths].reverse() : paths;
    order.forEach(({ s, si, pts }) => {
      const segs = []; let cur = [];
      pts.forEach((p) => { if (p) cur.push(p); else if (cur.length) { segs.push(cur); cur = []; } });
      if (cur.length) segs.push(cur);
      segs.forEach((seg, k) => {
        const d = seg.map((p, j) => `${j ? 'L' : 'M'}${p.x},${p.y}`).join('');
        if (type === 'area') {
          const back = [...seg].reverse().map((p) => `L${p.x},${p.y0}`).join('');
          el.push(<path key={`a${si}-${k}`} d={`${d}${back}Z`} fill={s.color} fillOpacity={stacked ? 0.85 : 0.45} stroke={s.color} strokeWidth={1.5} />);
        } else el.push(<path key={`l${si}-${k}`} d={d} fill="none" stroke={s.color} strokeWidth={2} strokeLinejoin="round" />);
      });
      if (type === 'line') pts.forEach((p, i) => p && el.push(<circle key={`m${si}-${i}`} cx={p.x} cy={p.y} r={3} fill="#fff" stroke={s.color} strokeWidth={1.5}><title>{`${s.name} · ${data.categories[i] ?? ''}: ${fmtNum(p.v)}`}</title></circle>));
    });
  } else if (scatter) {
    S.forEach((s, si) => s.values.forEach((v, i) => {
      const x = data.categories[i];
      if (v == null || typeof x !== 'number') return;
      el.push(<circle key={`s${si}-${i}`} cx={xScale(x)} cy={vScale(v)} r={3.5} fill={s.color}><title>{`${s.name}: ${fmtNum(x)}, ${fmtNum(v)}`}</title></circle>);
    }));
  }
  return <g>{el}</g>;
}

/** Pure SVG chart. */
export function ChartSVG({ chart, data, width, height }) {
  const W = Math.max(60, width || 400); const H = Math.max(40, height || 260);
  const type = chart.type || 'column';
  const pie = type === 'pie' || type === 'doughnut';
  const label = chart.title || (CHART_TYPES.find((t) => t.id === type) || {}).label || 'Chart';
  const empty = !data || !data.series || !data.series.length || !data.series.some((s) => s.values.some((v) => v != null && (!pie || v > 0)));
  const items = !data ? [] : pie ? (data.categories || []).map((c, i) => ({ name: c, color: (data.pointColors && data.pointColors[i]) || PALETTE[i % PALETTE.length] })) : data.series.map((s) => ({ name: s.name, color: s.color }));
  const legend = chart.legend || 'right';
  const titleH = chart.title ? 26 : 6;
  const ls = legendSpace(legend, items, W);
  const xTitleH = !pie && chart.xTitle ? 16 : 0; const yTitleW = !pie && chart.yTitle ? 16 : 0;
  const plot = { x: 8 + ls.left + yTitleW, y: titleH + ls.top + 4, w: 0, h: 0 };
  plot.w = Math.max(10, W - plot.x - 8 - ls.right); plot.h = Math.max(10, H - plot.y - 8 - ls.bottom - xTitleH);
  const lbox = legend === 'right' ? { x: W - ls.right + 4, y: plot.y, w: ls.right, h: plot.h } : legend === 'left' ? { x: 8, y: plot.y, w: ls.left, h: plot.h }
    : legend === 'top' ? { x: 0, y: titleH, w: W, h: 20 } : { x: 0, y: H - 18, w: W, h: 20 };
  return (
    <svg xmlns="http://www.w3.org/2000/svg" width={W} height={H} viewBox={`0 0 ${W} ${H}`} role="img" aria-label={label} style={{ fontFamily: 'inherit', display: 'block' }}>
      <rect x={0} y={0} width={W} height={H} fill="#fff" />
      {chart.title && <text x={W / 2} y={18} textAnchor="middle" fontSize={14} fontWeight={600} fill={TEXT}><Keep>{clip(chart.title, Math.floor(W / 8))}</Keep></text>}
      {empty ? <text x={W / 2} y={H / 2} textAnchor="middle" fontSize={12} fill={MUTED}>No data</text> : (
        <>
          {pie ? <Pie data={data} plot={plot} doughnut={type === 'doughnut'} /> : <Cartesian chart={chart} data={data} plot={plot} />}
          <Legend items={items} pos={legend} box={lbox} />
          {xTitleH > 0 && <text x={plot.x + plot.w / 2} y={plot.y + plot.h + xTitleH + 2} textAnchor="middle" fontSize={11} fill={TEXT}><Keep>{chart.xTitle}</Keep></text>}
          {yTitleW > 0 && <text transform={`translate(${8 + ls.left + 11},${plot.y + plot.h / 2}) rotate(-90)`} textAnchor="middle" fontSize={11} fill={TEXT}><Keep>{chart.yTitle}</Keep></text>}
        </>
      )}
    </svg>
  );
}
