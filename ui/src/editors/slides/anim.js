/*
 * Animations of the slide show – pure (used by Present.jsx, the pptx/odp writers and tests).
 *   anim = { id, el, cls: 'entr'|'emph'|'exit'|'path', effect, dir?, trigger: 'click'|'with'|'after', duration, delay,
 *            path?, scale?, angle?, opacity?, raw? }
 */

/* ---------------------------------------------------------------- catalog */
/** effects the editor offers, by class (labels are English UI texts) */
export const ANIM_EFFECTS = {
  entr: [['appear', 'Appear'], ['fade', 'Fade'], ['fly', 'Fly in'], ['float', 'Float in'], ['wipe', 'Wipe'], ['split', 'Split'], ['wheel', 'Wheel'], ['zoom', 'Zoom']],
  emph: [['pulse', 'Pulse'], ['spin', 'Spin'], ['grow', 'Grow/Shrink'], ['teeter', 'Teeter'], ['transparency', 'Transparency']],
  exit: [['appear', 'Disappear'], ['fade', 'Fade out'], ['fly', 'Fly out'], ['float', 'Float out'], ['wipe', 'Wipe out'], ['split', 'Split out'], ['zoom', 'Zoom out']],
  path: [['line', 'Lines'], ['arc', 'Arcs'], ['turn', 'Turns'], ['circle', 'Shapes'], ['custom', 'Custom path']],
};
export const CLASS_LABEL = { entr: 'Entrance', emph: 'Emphasis', exit: 'Exit', path: 'Motion path' };
export const effectLabel = (a) => {
  if (a.effect === 'raw') return 'Other effect (kept)';
  const l = (ANIM_EFFECTS[a.cls || 'entr'] || []).find(([id]) => id === a.effect);
  return l ? l[1] : a.cls === 'path' ? 'Motion path' : a.effect;
};

/** PowerPoint preset ids (presetClass + presetID) */
export const PRESETS = [
  { cls: 'entr', id: 1, effect: 'appear' }, { cls: 'entr', id: 2, effect: 'fly' }, { cls: 'entr', id: 10, effect: 'fade' }, { cls: 'entr', id: 16, effect: 'split' },
  { cls: 'entr', id: 21, effect: 'wheel' }, { cls: 'entr', id: 22, effect: 'wipe' }, { cls: 'entr', id: 42, effect: 'float' }, { cls: 'entr', id: 53, effect: 'zoom' },
  { cls: 'entr', id: 23, effect: 'zoom', alias: true },
  { cls: 'exit', id: 1, effect: 'appear' }, { cls: 'exit', id: 2, effect: 'fly' }, { cls: 'exit', id: 10, effect: 'fade' }, { cls: 'exit', id: 16, effect: 'split' },
  { cls: 'exit', id: 22, effect: 'wipe' }, { cls: 'exit', id: 42, effect: 'float' }, { cls: 'exit', id: 53, effect: 'zoom' },
  { cls: 'emph', id: 6, effect: 'grow' }, { cls: 'emph', id: 8, effect: 'spin' }, { cls: 'emph', id: 9, effect: 'transparency' }, { cls: 'emph', id: 26, effect: 'pulse' },
  { cls: 'emph', id: 32, effect: 'teeter' },
  { cls: 'path', id: 0, effect: 'custom' }, { cls: 'path', id: 42, effect: 'line' }, { cls: 'path', id: 37, effect: 'arc' }, { cls: 'path', id: 41, effect: 'turn' }, { cls: 'path', id: 1, effect: 'circle' },
];
export function presetOf(a) {
  const cls = a.cls || 'entr';
  return PRESETS.find((p) => p.cls === cls && p.effect === a.effect && !p.alias) || (cls === 'path' ? PRESETS.find((p) => p.cls === 'path' && p.id === 0) : null);
}
/** default motion paths (fractions of the slide, relative to the shape's position) */
export const DEFAULT_PATHS = {
  line: 'M 0 0 L 0 0.25 E',
  arc: 'M 0 0 C 0.1 -0.2 0.3 -0.2 0.4 0 E',
  turn: 'M 0 0 L 0.2 0 C 0.3 0 0.3 0.1 0.3 0.2 L 0.3 0.3 E',
  circle: 'M 0 0 C 0.06 0 0.11 0.09 0.11 0.2 C 0.11 0.31 0.06 0.4 0 0.4 C -0.06 0.4 -0.11 0.31 -0.11 0.2 C -0.11 0.09 -0.06 0 0 0 Z E',
  custom: 'M 0 0 L 0.25 0 E',
};

/* ---------------------------------------------------------------- motion paths */
/** points of a PowerPoint motion path (fractions), sampled; curves get `n` points each */
export function pathPoints(path, n = 12) {
  const tk = String(path || '').match(/[MLCZEmlcze]|-?\d*\.?\d+(?:e-?\d+)?/g) || [];
  const pts = [];
  let i = 0; let cmd = 'M'; let cur = [0, 0]; let start = [0, 0];
  const num = () => parseFloat(tk[i++]);
  while (i < tk.length) {
    if (/^[A-Za-z]$/.test(tk[i])) { cmd = tk[i++].toUpperCase(); if (cmd === 'E') break; if (cmd === 'Z') { pts.push(start.slice()); cur = start.slice(); continue; } }
    if (i >= tk.length) break;
    if (cmd === 'M') { cur = [num(), num()]; start = cur.slice(); pts.push(cur.slice()); cmd = 'L'; }
    else if (cmd === 'L') { cur = [num(), num()]; pts.push(cur.slice()); }
    else if (cmd === 'C') {
      const p1 = [num(), num()]; const p2 = [num(), num()]; const p3 = [num(), num()];
      for (let k = 1; k <= n; k++) { const t = k / n; const u = 1 - t; pts.push([u * u * u * cur[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t * t * t * p3[0], u * u * u * cur[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t * t * t * p3[1]]); }
      cur = p3;
    } else i++;
  }
  return pts.length ? pts : [[0, 0]];
}
export const pathEnd = (path) => { const p = pathPoints(path, 2); return p[p.length - 1]; };
/** the path as SVG (slide px) starting at x,y – for the editor's path overlay */
export function pathSvg(path, x, y, W, H) {
  return pathPoints(path, 16).map((p, k) => `${k ? 'L' : 'M'}${(x + p[0] * W).toFixed(1)} ${(y + p[1] * H).toFixed(1)}`).join('');
}

/* ---------------------------------------------------------------- steps */
/** click groups of a slide: [{ auto, items: [{ el, cls, effect, dir, duration, delay, … }] }] */
export function buildSteps(slide) {
  const groups = [];
  let t0 = 0; let last = 0;
  const ids = new Set(); const walkIds = (els) => els.forEach((e) => { ids.add(e.id); if (e.children) walkIds(e.children); });
  walkIds(slide.elements || []);
  for (const a of slide.animations || []) {
    if (!ids.has(a.el)) continue;
    const dur = a.duration || 500;
    if (a.trigger === 'click' || !groups.length) { groups.push({ auto: a.trigger !== 'click' && !groups.length, items: [] }); t0 = 0; last = 0; }
    let delay = a.delay || 0;
    if (a.trigger === 'after') delay += last;
    else if (a.trigger === 'with') delay += t0;
    const g = groups[groups.length - 1];
    g.items.push({ ...a, cls: a.cls || 'entr', effect: a.effect || 'fade', dir: a.dir || 'b', duration: dur, delay });
    t0 = delay; last = Math.max(last, delay + dur);
  }
  return groups;
}
export const visibleSlides = (doc) => doc.slides.map((s, i) => ({ s, i })).filter((x) => !x.s.hidden);

const CSS_NAME = { appear: 'appear', fade: 'fade', zoom: 'zoom', wipe: 'wipe', float: 'float', split: 'split', wheel: 'wheel' };
/**
 * element decoration for a slide in a given animation state: step = click groups done (the last one is playing when
 * `playing`); size = { W, H } of the slide in px (motion paths). Returns el → { style } | null.
 */
export function animDecor(groups, step, playing, anim = true, size = { W: 960, H: 540 }) {
  const per = {};
  groups.forEach((g, gi) => g.items.forEach((it) => { (per[it.el] = per[it.el] || []).push({ ...it, gi }); }));
  return (el) => {
    const list = per[el.id];
    if (!list) return null;
    let visible = list[0].cls !== 'entr' || list[0].effect === 'raw' && !/presetClass="entr"/.test(list[0].raw || '');
    if (list[0].effect === 'raw' && /presetClass="entr"/.test(list[0].raw || '')) visible = false;
    let dx = 0; let dy = 0; let scale = 1; let opacity = null; let play = null;
    for (const it of list) {
      if (it.gi >= step) break;
      const now = anim && playing && it.gi === step - 1;
      if (now && !play) play = { it, before: { visible, dx, dy, scale, opacity } };
      if (it.cls === 'entr') visible = true;
      else if (it.cls === 'exit') visible = false;
      else if (it.cls === 'path') { const e = pathEnd(it.path); dx += e[0] * size.W; dy += e[1] * size.H; }
      else if (it.cls === 'emph' && it.effect === 'grow') scale *= it.scale || 1.5;
      else if (it.cls === 'emph' && it.effect === 'transparency') opacity = it.opacity ?? 0.5;
      else if (it.effect === 'raw' && /presetClass="entr"/.test(it.raw || '')) visible = true;
      else if (it.effect === 'raw' && /presetClass="exit"/.test(it.raw || '')) visible = false;
    }
    const st = {};
    const b = play ? play.before : null;
    const ox = b ? b.dx : dx; const oy = b ? b.dy : dy; const os = b ? b.scale : scale;
    if (ox || oy) st.translate = `${ox.toFixed(1)}px ${oy.toFixed(1)}px`;
    if (os !== 1) st.scale = String(os);
    if ((b ? b.opacity : opacity) != null) st.opacity = b ? b.opacity : opacity;
    if (play) {
      const it = play.it;
      const timing = `${it.duration}ms cubic-bezier(.2,.8,.2,1) ${it.delay}ms both`;
      if (it.cls === 'entr') {
        if (it.effect === 'appear') st.animation = `of-slides-appear 1ms linear ${it.delay}ms both`;
        else st.animation = `${it.effect === 'fly' ? `of-slides-fly-${it.dir || 'b'}` : `of-slides-${CSS_NAME[it.effect] || 'fade'}`} ${timing}`;
      } else if (it.cls === 'exit') {
        if (it.effect === 'appear') st.animation = `of-slides-appear 1ms linear ${it.delay}ms reverse both`;
        else st.animation = `${it.effect === 'fly' ? `of-slides-fly-${it.dir || 'b'}` : `of-slides-${CSS_NAME[it.effect] || 'fade'}`} ${it.duration}ms cubic-bezier(.8,.2,.8,.2) ${it.delay}ms reverse both`;
      } else if (it.cls === 'emph') {
        st['--of-a'] = `${it.angle || 360}deg`; st['--of-s'] = String(it.scale || 1.5); st['--of-o'] = String(it.opacity ?? 0.5);
        st.animation = `of-slides-e-${it.effect === "spin" ? "turn" : it.effect} ${it.duration}ms ease-in-out ${it.delay}ms both`;
      } else if (it.cls === 'path') {
        st['--of-x0'] = `${ox}px`; st['--of-y0'] = `${oy}px`; // motion itself: motionJobs() (Web Animations)
      }
      if (it.cls !== 'entr' && !b.visible) st.visibility = 'hidden';
      return { style: st, playing: true };
    }
    if (!visible) st.visibility = 'hidden';
    return Object.keys(st).length ? { style: st } : null;
  };
}

/** Web Animations for the motion paths of the playing group: [{ el, frames: [{ translate }], duration, delay }] */
export function motionJobs(groups, step, size = { W: 960, H: 540 }) {
  const g = groups[step - 1];
  if (!g) return [];
  const before = {};
  groups.slice(0, step - 1).forEach((gg) => gg.items.forEach((it) => { if (it.cls === 'path') { const e = pathEnd(it.path); const b = before[it.el] || [0, 0]; before[it.el] = [b[0] + e[0] * size.W, b[1] + e[1] * size.H]; } }));
  return g.items.filter((it) => it.cls === 'path').map((it) => {
    const o = before[it.el] || [0, 0];
    const frames = pathPoints(it.path).map((p) => ({ translate: `${(o[0] + p[0] * size.W).toFixed(1)}px ${(o[1] + p[1] * size.H).toFixed(1)}px` }));
    return { el: it.el, frames, duration: it.duration, delay: it.delay };
  });
}

/* ---------------------------------------------------------------- morph */
const textOf = (e) => (e.text && e.text.paras ? e.text.paras.map((p) => (p.runs || []).map((r) => r.text || '').join('')).join('\n') : '');
/**
 * pairs of elements a Morph transition moves from the previous slide to this one: same name ("!!name" or any name
 * used once on both slides), else same type + same text, else same picture. [{ from: El, to: El }]
 */
export function morphPairs(prev, cur) {
  const pairs = [];
  const used = new Set();
  const cand = (prev && prev.elements) || [];
  const key = [(e) => (e.name ? `n:${e.name}` : null), (e) => (textOf(e) ? `t:${e.type}:${textOf(e)}` : null), (e) => (e.src ? `s:${e.src.length}:${e.src.slice(-40)}` : null), (e) => (e.type === 'shape' && !textOf(e) ? `g:${e.geom}:${JSON.stringify(e.fill)}` : null)];
  for (const k of key) {
    for (const e of (cur && cur.elements) || []) {
      if (pairs.some((p) => p.to === e)) continue;
      const v = k(e);
      if (!v) continue;
      const same = cand.filter((x) => !used.has(x.id) && k(x) === v);
      if (same.length === 1 || (same.length && k === key[0])) { used.add(same[0].id); pairs.push({ from: same[0], to: e }); }
    }
  }
  return pairs;
}
