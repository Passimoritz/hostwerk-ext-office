/* Slide show steps (click groups of entrance animations) – pure, used by Present.jsx and tests. */
/* ---------------------------------------------------------------- steps */
/** click groups of a slide: [{ auto, items: [{ el, effect, dir, duration, delay }] }] */
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
    if (a.trigger === 'after') { delay += last; }
    else if (a.trigger === 'with') delay += t0;
    const g = groups[groups.length - 1];
    g.items.push({ el: a.el, effect: a.effect || 'fade', dir: a.dir || 'b', duration: dur, delay });
    t0 = delay; last = delay + dur;
  }
  return groups;
}
export const visibleSlides = (doc) => doc.slides.map((s, i) => ({ s, i })).filter((x) => !x.s.hidden);

/** element decoration for a slide in a given animation state */
export function animDecor(groups, step, playing, anim = true) {
  const groupOf = {}; const item = {};
  groups.forEach((g, gi) => g.items.forEach((it) => { if (groupOf[it.el] === undefined) { groupOf[it.el] = gi; item[it.el] = it; } }));
  return (el) => {
    const gi = groupOf[el.id];
    if (gi === undefined) return null;
    if (gi >= step) return { style: { visibility: 'hidden' } };
    if (anim && playing && gi === step - 1) {
      const it = item[el.id];
      if (it.effect === 'appear') return { style: { animation: `of-slides-appear 1ms linear ${it.delay}ms both` } };
      const name = it.effect === 'fly' ? `of-slides-fly-${it.dir || 'b'}` : `of-slides-${it.effect}`;
      return { style: { animation: `${name} ${it.duration}ms cubic-bezier(.2,.8,.2,1) ${it.delay}ms both` } };
    }
    return null;
  };
}

