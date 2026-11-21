/*
 * Adaptive ribbon (like Office's): when a ribbon body is wider than its window, the right-most groups collapse into a
 * drop-down button one after the other; whenever there is room again the left-most collapsed group that fits opens
 * up (so a big group that had to go first does not leave an empty ribbon behind). Below the narrow
 * breakpoint (≤ 700 px, see ARCHITECTURE "Narrow windows") nothing collapses – the ribbon scrolls sideways there.
 *
 *   const collapsed = useRibbonFit(bodyRef, { key: tab });         // Set of group names (data-group) to collapse
 *   const level = useFitLevel(bodyRef, { key: tab, max: 2 });       // 0 … max compaction steps (toolbars without groups)
 *
 * Groups mark themselves with `data-group="<name>"` (direct children of the body); `data-fit="keep"` never collapses.
 * Pure helpers (`nextToCollapse`, `nextToExpand`, `levelForWidth`, `firstIcon`) are unit-tested in tests/slides-ui.test.js.
 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';

export const NARROW = '(max-width: 700px)';

/** the right-most group that may still collapse, or null */
export function nextToCollapse(groups, collapsed) {
  for (let i = groups.length - 1; i >= 0; i--) {
    const g = groups[i];
    if (!g || !g.name || g.keep || collapsed.includes(g.name)) continue;
    return g.name;
  }
  return null;
}

/** how many steps can be undone at `width` – need[i] = width the body needed with i steps applied */
export function levelForWidth(need, level, width) {
  let l = level;
  while (l > 0 && need[l - 1] != null && width >= need[l - 1]) l--;
  return l;
}

function useNarrow() {
  const mq = typeof window !== 'undefined' && window.matchMedia ? window.matchMedia(NARROW) : null;
  const [narrow, setNarrow] = useState(() => !!(mq && mq.matches));
  useEffect(() => {
    if (!mq) return undefined;
    const on = () => setNarrow(mq.matches);
    if (mq.addEventListener) mq.addEventListener('change', on); else mq.addListener(on);
    return () => { if (mq.removeEventListener) mq.removeEventListener('change', on); else mq.removeListener(on); };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  return narrow;
}

/** core: a stack of steps; `step(el, steps)` returns the next step or null; re-measured after every render */
function useSteps(ref, key, enabled, step) {
  const [steps, setSteps] = useState([]);
  const need = useRef([]);
  const keyRef = useRef(key);
  const [, tick] = useState(0);
  useLayoutEffect(() => {
    const el = ref.current;
    if (keyRef.current !== key) { keyRef.current = key; need.current = []; if (steps.length) { setSteps([]); return; } }
    if (!el) return;
    if (!enabled) { if (steps.length) { need.current = []; setSteps([]); } return; }
    if (el.scrollWidth > el.clientWidth + 1) {
      const next = step(el, steps);
      if (next != null) { need.current[steps.length] = el.scrollWidth; setSteps([...steps, next]); }
    }
  });
  useEffect(() => {
    const el = ref.current;
    if (!el || typeof ResizeObserver === 'undefined') return undefined;
    let w = el.clientWidth;
    const ro = new ResizeObserver(() => {
      const nw = el.clientWidth;
      if (nw === w) return;
      const grew = nw > w; w = nw;
      if (grew) setSteps((s) => { const l = levelForWidth(need.current, s.length, nw); need.current.length = l; return l === s.length ? s : s.slice(0, l); });
      else tick((n) => n + 1); // re-measure (may collapse more)
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [ref.current]); // eslint-disable-line react-hooks/exhaustive-deps
  return steps;
}

/** collapsed groups that may open again in `free` px: [{ name, extra }] in ribbon order → name or null (left-most first) */
export function nextToExpand(collapsed, free, blocked = []) {
  for (const g of collapsed) if (!blocked.includes(g.name) && g.extra <= free) return g.name;
  return null;
}

/*
 * Groups: collapse the right-most group while the body overflows; when there is room again (window grew, or a big
 * group left room behind it) the left-most collapsed group whose full width fits opens again. Full widths are
 * measured just before a group collapses; a group that overflowed right after opening stays closed until the width
 * changes (no flicker).
 */
export function useRibbonFit(ref, { key, off = false } = {}) {
  const narrow = useNarrow() || off;
  const [collapsed, setCollapsed] = useState([]);
  const full = useRef(new Map());
  const blocked = useRef({ names: [], width: 0 });
  const last = useRef(null);
  const keyRef = useRef(key);
  const [, tick] = useState(0);
  useLayoutEffect(() => {
    const el = ref.current;
    if (keyRef.current !== key) { keyRef.current = key; full.current = new Map(); blocked.current = { names: [], width: 0 }; last.current = null; if (collapsed.length) { setCollapsed([]); return; } }
    if (!el) return;
    if (narrow) { if (collapsed.length) setCollapsed([]); return; }
    if (blocked.current.width !== el.clientWidth) blocked.current = { names: [], width: el.clientWidth };
    const kids = [...el.querySelectorAll(':scope > [data-group]')];
    const groups = kids.map((g) => ({ el: g, name: g.getAttribute('data-group'), keep: g.getAttribute('data-fit') === 'keep' }));
    if (el.scrollWidth > el.clientWidth + 1) {
      // the group opened last did not fit after all → close it again and leave it closed at this width
      const reopened = last.current && groups.some((x) => x.name === last.current) && !collapsed.includes(last.current) ? last.current : null;
      const name = reopened || nextToCollapse(groups, collapsed);
      if (name == null) return; // everything collapsed – the body scrolls
      const g = groups.find((x) => x.name === name);
      full.current.set(name, g.el.getBoundingClientRect().width);
      if (reopened) blocked.current.names.push(name);
      last.current = null;
      setCollapsed([...collapsed, name]);
      return;
    }
    last.current = null;
    if (!collapsed.length || !kids.length) return;
    const box = el.getBoundingClientRect();
    const pad = parseFloat(getComputedStyle(el).paddingRight) || 0;
    const used = kids[kids.length - 1].getBoundingClientRect().right - box.left + pad;
    const free = el.clientWidth - used - 2;
    const list = groups.filter((x) => collapsed.includes(x.name)).map((x) => ({ name: x.name, extra: (full.current.get(x.name) || Infinity) - x.el.getBoundingClientRect().width }));
    const open = nextToExpand(list, free, blocked.current.names);
    if (open != null) { last.current = open; setCollapsed(collapsed.filter((n) => n !== open)); }
  });
  useEffect(() => {
    const el = ref.current;
    if (!el || typeof ResizeObserver === 'undefined') return undefined;
    let w = el.clientWidth;
    const ro = new ResizeObserver(() => { if (el.clientWidth !== w) { w = el.clientWidth; tick((n) => n + 1); } });
    ro.observe(el);
    return () => ro.disconnect();
  }, [ref.current]); // eslint-disable-line react-hooks/exhaustive-deps
  return new Set(collapsed);
}

export function useFitLevel(ref, { key, max = 1 } = {}) {
  const narrow = useNarrow();
  const steps = useSteps(ref, key, !narrow, (el, done) => (done.length < max ? done.length + 1 : null));
  return steps.length;
}

/** first lucide-like icon component found in a React children tree (the collapsed group's button shows it) */
export function firstIcon(children, depth = 0) {
  if (depth > 6 || children == null || typeof children !== 'object') return null;
  const list = Array.isArray(children) ? children : [children];
  for (const c of list) {
    if (!c || typeof c !== 'object') continue;
    if (Array.isArray(c)) { const r = firstIcon(c, depth + 1); if (r) return r; continue; }
    const p = c.props || {};
    if (p.icon && (typeof p.icon === 'function' || typeof p.icon === 'object')) return p.icon;
    const r = firstIcon(p.children, depth + 1);
    if (r) return r;
  }
  return null;
}
