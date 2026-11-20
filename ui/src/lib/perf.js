/*
 * Open → interactive measurements (Performance API). The click that opens a file calls markOpen(); the editor host
 * calls interactive(kind) when the chunk reports host.perf('interactive') – or, for chunks that do not, after the
 * editor's first painted frame. Numbers: performance.getEntriesByName('office:open:<kind>') and window.__officePerf
 * ({ <kind>: [ms …] }, also kept for this browser in localStorage 'office:perf' – last 20 per kind).
 */
let openAt = 0;
let openName = '';
export function markOpen(name = '') {
  openAt = performance.now(); openName = name;
  try { performance.mark('office:open', { detail: { name } }); } catch { /* old browser */ }
}
const store = () => { try { return JSON.parse(localStorage.getItem('office:perf') || '{}') || {}; } catch { return {}; } };
export function perfStats() { return store(); }
/** → ms from the click (or from `since`) to now, recorded once per open */
export function interactive(kind, { since = 0, name = '' } = {}) {
  const start = openAt || since;
  if (!start) return null;
  const ms = Math.round(performance.now() - start);
  openAt = 0;
  try { performance.measure(`office:open:${kind}`, { start, end: performance.now(), detail: { name: name || openName } }); } catch { /* old browser */ }
  const all = store();
  all[kind] = [...(all[kind] || []), ms].slice(-20);
  try { localStorage.setItem('office:perf', JSON.stringify(all)); } catch { /* private window */ }
  try { window.__officePerf = all; } catch { /* ignore */ }
  return ms;
}
/** median of the recorded numbers per kind → { docs: { n, median, last } } */
export function perfSummary(all = store()) {
  const out = {};
  for (const [k, list] of Object.entries(all || {})) {
    const s = [...list].sort((a, b) => a - b);
    if (s.length) out[k] = { n: s.length, median: s[Math.floor((s.length - 1) / 2)], last: list[list.length - 1] };
  }
  return out;
}
/** two animation frames = the browser painted what React rendered */
export const afterPaint = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
