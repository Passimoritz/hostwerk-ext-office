/*
 * Print layout pagination – real page breaks inside one editing flow:
 *   · blocks that do not fit move to the next page (spacer widget before them)
 *   · long paragraphs, lists and quotes break between two LINES (a block spacer widget at the first text position of
 *     the line that no longer fits – the line starts there anyway, so nothing re-wraps)
 *   · long tables break between ROWS; a header row (tableHeader cells) is repeated at the top of the next page
 *   · page breaks and section breaks start a new page; every section has its own page size, orientation and margins
 *     (blocks of a section with another text width get that width)
 * Speed: only the part after the first change is measured again, and measuring stops as soon as the layout is the
 * same as before (typing in a 300-page document measures a page or two); long first measurements run in slices so
 * the page stays responsive. Margins are cached per element style; nothing touches the DOM while measuring.
 */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

export const pageKey = new PluginKey('ofPages');
const MM = 96 / 25.4;
const SLICE_MS = 12;
const r1 = (v) => Math.round(v * 10) / 10;

/** page geometry of a section (mm page setup → px) */
export function geometry(page, box) {
  const m = page.margins || { top: 25, right: 25, bottom: 20, left: 25 };
  return { pageW: box.width * MM, pageH: box.height * MM, mt: m.top * MM, mb: m.bottom * MM, ml: m.left * MM, mr: m.right * MM, contentW: (box.width - m.left - m.right) * MM, contentH: (box.height - m.top - m.bottom) * MM };
}

const decoCache = new WeakMap();

export const Pagination = Extension.create({
  name: 'ofPages',
  addStorage() { return { enabled: false, sections: [], gap: 24, onPages: null, pages: [], stats: { last: 0, blocks: 0 } }; },
  addCommands() {
    return {
      setPagination: (opts) => ({ tr, dispatch }) => { Object.assign(this.storage, opts); if (dispatch) dispatch(tr.setMeta(pageKey, { reset: true }).setMeta('ofPages', true).setMeta('addToHistory', false)); return true; },
    };
  },
  addProseMirrorPlugins() {
    const ext = this;
    const buildDecos = (state) => {
      const st = pageKey.getState(state);
      const decos = [];
      const size = state.doc.content.size;
      for (const s of st.spacers) {
        if (s.pos < 0 || s.pos > size) continue;
        if (s.kind === 'row') (s.cells || []).forEach((c, i) => { if (c >= 0 && c <= size) decos.push(Decoration.widget(c, () => rowWidget(s, i), { side: -1, key: `pr${c}_${Math.round(s.height)}_${s.band.gapTop}`, ignoreSelection: true })); });
        else decos.push(Decoration.widget(s.pos, () => { const d = document.createElement(s.kind === 'line' ? 'span' : 'div'); d.className = s.kind === 'line' ? 'of-docs-pgsp of-pgsp-line' : 'of-docs-pgsp'; d.style.height = `${s.height}px`; d.contentEditable = 'false'; return d; }, { side: -1, key: `pg${s.kind}${s.pos}_${Math.round(s.height)}`, ignoreSelection: true, marks: [] }));
      }
      // sections with another text width than the last one (the body column is the last section's)
      const secs = ext.storage.sections;
      if (secs.length > 1) {
        const base = secs[secs.length - 1];
        let si = 0;
        state.doc.forEach((node, offset) => {
          const g = secs[Math.min(si, secs.length - 1)];
          if (node.type.name === 'sectionBreak') { si++; return; }
          if (g && (Math.abs(g.contentW - base.contentW) > 1 || Math.abs(g.ml - base.ml) > 1)) decos.push(Decoration.node(offset, offset + node.nodeSize, { style: `position:relative;left:${Math.round(g.ml - base.ml)}px;width:${Math.round(g.contentW)}px;max-width:none;align-self:flex-start` }));
        });
      }
      return decos.length ? DecorationSet.create(state.doc, decos) : null;
    };
    return [new Plugin({
      key: pageKey,
      state: {
        init: () => ({ spacers: [], dirty: 0, layout: [] }),
        apply(tr, value) {
          const meta = tr.getMeta(pageKey);
          if (meta && meta.spacers) return { spacers: meta.spacers, dirty: meta.dirty != null ? meta.dirty : Infinity, layout: meta.layout || value.layout };
          if (meta && meta.reset) return { spacers: value.spacers, dirty: 0, layout: [] };
          if (meta && meta.dirtyAt != null) return { ...value, dirty: Math.min(value.dirty === Infinity ? Number.MAX_SAFE_INTEGER : value.dirty, meta.dirtyAt) };
          if (!tr.docChanged) return value;
          // first changed position (new document coordinates) → measure again from there
          let dirty = value.dirty === Infinity ? Infinity : tr.mapping.map(value.dirty, -1);
          const maps = tr.mapping.maps;
          maps.forEach((map, k) => map.forEach((oldStart, oldEnd, newStart) => {
            let p = newStart; for (let j = k + 1; j < maps.length; j++) p = maps[j].map(p, -1);
            if (p < dirty) dirty = p;
          }));
          const spacers = [];
          for (const s of value.spacers) {
            const r = tr.mapping.mapResult(s.pos, -1);
            if (r.deleted) continue;
            spacers.push(s.cells ? { ...s, pos: r.pos, cells: s.cells.map((c) => tr.mapping.map(c, 1)) } : { ...s, pos: r.pos });
          }
          const layout = value.layout.map((b) => ({ ...b, pos: tr.mapping.map(b.pos, 1) }));
          const next = { spacers, dirty, layout };
          const prevDeco = decoCache.get(value);
          if (prevDeco) decoCache.set(next, prevDeco.map(tr.mapping, tr.doc));
          return next;
        },
      },
      props: {
        decorations(state) {
          if (!ext.storage.enabled) return null;
          const st = pageKey.getState(state);
          // built once per spacer set and then only mapped (typing in a 300-page document stays cheap)
          const cached = decoCache.get(st);
          if (cached !== undefined) return cached;
          const built = buildDecos(state);
          decoCache.set(st, built);
          return built;
        },
      },
      view(view) {
        let timer = null; let job = null;
        let pointerUntil = 0;
        const marginCache = new Map();
        const margins = (el) => {
          const key = `${el.tagName}|${el.className}|${el.getAttribute('style') || ''}|${el.getAttribute('data-style') || ''}`;
          let v = marginCache.get(key);
          if (!v) { const cs = getComputedStyle(el); v = [parseFloat(cs.marginTop) || 0, parseFloat(cs.marginBottom) || 0]; if (marginCache.size > 2000) marginCache.clear(); marginCache.set(key, v); }
          return v;
        };
        const finish = (out, pages, layout, measured, t0, from) => {
          job = null;
          const st = ext.storage;
          const cur = pageKey.getState(view.state);
          const same = out.length === cur.spacers.length && out.every((s, i) => s.pos === cur.spacers[i].pos && s.kind === cur.spacers[i].kind && Math.abs(s.height - cur.spacers[i].height) < 1);
          view.dispatch(view.state.tr.setMeta(pageKey, { spacers: same ? cur.spacers : out, dirty: Infinity, layout }).setMeta('addToHistory', false).setMeta('ofPages', true));
          st.pages = pages;
          st.stats = { last: Math.round((performance.now() - t0) * 10) / 10, blocks: measured, from, total: view.state.doc.childCount, pages: pages.length, at: Date.now() };
          if (st.onPages) st.onPages(pages);
        };
        const measure = () => {
          timer = null;
          const st = ext.storage;
          if (!st.enabled || !view.dom.isConnected || !st.sections.length) { if (st.onPages) st.onPages(null); return; }
          // never re-render under a running mouse selection (a double/triple click would lose its word selection)
          if (performance.now() < pointerUntil) { schedule(pointerUntil - performance.now() + 30); return; }
          const state = view.state;
          const pst = pageKey.getState(state);
          if (pst.dirty === Infinity && pst.layout.length) return;
          const t0 = performance.now();
          const dom = view.dom;
          const rect = dom.getBoundingClientRect();
          const k = dom.offsetWidth ? rect.width / dom.offsetWidth : 1;
          if (!k) return;
          const old = pst.spacers;
          const secs = st.sections; const gap = st.gap;
          const sec = (i) => secs[Math.min(i, secs.length - 1)];
          const pages = [];
          const addPage = (section) => { const prev = pages[pages.length - 1]; const g = sec(section); const top = prev ? prev.top + prev.h + gap : 0; pages.push({ top, section, h: g.pageH, w: g.pageW, mt: g.mt, mb: g.mb, ml: g.ml, mr: g.mr }); return pages.length - 1; };
          const mt0 = secs[0].mt;
          const C = (p) => pages[p].top + pages[p].mt - mt0;
          const E = (p) => C(p) + sec(pages[p].section).contentH;
          // rendered spacers before a position (subtracted → natural layout)
          const prefix = []; let acc = 0;
          for (const s of old) { acc += s.height; prefix.push(acc); }
          const heightBefore = (pos, incl) => { let lo = 0; let hi = old.length; while (lo < hi) { const mid = (lo + hi) >> 1; if (incl ? old[mid].pos <= pos : old[mid].pos < pos) lo = mid + 1; else hi = mid; } return lo ? prefix[lo - 1] : 0; };
          const blocks = []; state.doc.forEach((node, offset) => blocks.push({ node, offset }));
          const prevLayout = pst.layout || [];
          // start at the block before the first change – everything before it keeps its layout
          let startIdx = 0; let at = -1;
          if (pst.dirty !== 0 && prevLayout.length && st.pages && st.pages.length) {
            let lo = 0; let hi = blocks.length - 1;
            while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (blocks[mid].offset <= pst.dirty) lo = mid; else hi = mid - 1; }
            startIdx = Math.max(0, lo - 1);
            at = prevLayout.findIndex((b) => b.pos === blocks[startIdx].offset);
            if (at < 0 || prevLayout[at].page >= st.pages.length) startIdx = 0;
          }
          let out; let layout; let page; let section;
          if (startIdx > 0) {
            const keepTo = blocks[startIdx].offset;
            out = old.filter((s) => s.pos < keepTo);
            layout = prevLayout.slice(0, at);
            ({ section, page } = prevLayout[at]);
            for (let p = 0; p <= page; p++) pages.push({ ...st.pages[p] });
          } else { out = []; layout = []; section = 0; addPage(0); page = 0; }
          let shift = out.reduce((a, s) => a + s.height, 0);
          let forced = null; let i = startIdx; let measured = 0;
          const oldByPos = new Map(prevLayout.map((b) => [b.pos, b]));
          const dirtyPos = pst.dirty;
          const next = (p) => (pages[p + 1] ? p + 1 : addPage(section));
          const env = { k, rect, view, C, E, gap, next, mb: (p) => pages[p].mb, mt: (p) => pages[p].mt };
          const pushBlock = (offset, top) => { const p = next(page); const push = C(p) - top; if (push > 0.5) { out.push({ pos: offset, height: r1(push), kind: 'block' }); shift += push; } page = p; return Math.max(0, push); };
          const step = () => {
            const sliceStart = performance.now();
            for (; i < blocks.length; i++) {
              if (measured > 60 && performance.now() - sliceStart > SLICE_MS) { job = setTimeout(step, 0); return; }
              const { node, offset } = blocks[i];
              const el = view.nodeDOM(offset);
              if (!el || el.nodeType !== 1) continue;
              measured++;
              const r = el.getBoundingClientRect();
              const [mt, mb] = margins(el);
              const before = heightBefore(offset, true);
              const inner = heightBefore(offset + node.nodeSize, false) - before;
              const natTop = (r.top - rect.top) / k - before - mt;
              let top = natTop + shift;
              let bottom = (r.bottom - rect.top) / k - before - inner + mb + shift;
              const name = node.type.name;
              if (forced) {
                if (forced === 'section') section++;
                const p = addPage(section);
                const push = C(p) - top;
                if (push > 0.5) { out.push({ pos: offset, height: r1(push), kind: 'block' }); shift += push; top += push; bottom += push; }
                page = p; forced = null;
              }
              while (top >= E(page) - 1 && pages.length < 5000) {
                const p = next(page); page = p;
                if (top < C(page)) { const push = C(page) - top; out.push({ pos: offset, height: r1(push), kind: 'block' }); shift += push; top += push; bottom += push; }
              }
              const contentH = sec(section).contentH;
              const tall = bottom - top > contentH - 4;
              const startOfPage = top <= C(page) + 2;
              if (bottom > E(page) + 0.5) {
                if (name === 'table') {
                  const res = splitTable(el, node, offset, before, shift, () => page, (p) => { page = p; }, env);
                  if (res.whole && !startOfPage && !tall) { const push = pushBlock(offset, top); top += push; bottom += push; }
                  else { for (const s of res.spacers) { out.push(s); shift += s.height; } bottom += res.added; }
                } else if (['paragraph', 'bulletList', 'orderedList', 'taskList', 'blockquote', 'callout', 'controlBlock', 'codeBlock'].includes(name)) {
                  const res = splitLines(el, before, shift, () => page, (p) => { page = p; }, env, offset);
                  if (res.whole && !startOfPage) {
                    const push = pushBlock(offset, top); top += push; bottom += push;
                    if (tall) { const r2 = splitLines(el, before, shift, () => page, (p) => { page = p; }, env, offset); for (const s of r2.spacers) { out.push(s); shift += s.height; } bottom += r2.added; }
                  } else { for (const s of res.spacers) { out.push(s); shift += s.height; } bottom += res.added; }
                } else if (!startOfPage && !tall) { const push = pushBlock(offset, top); top += push; bottom += push; }
              } else if (name === 'heading' && bottom > E(page) - 30 && !startOfPage && blocks[i + 1]) {
                // keep a heading together with the text after it
                const push = pushBlock(offset, top); top += push; bottom += push;
              }
              while (bottom > E(page) + 1 && pages.length < 5000) { page = next(page); }
              layout.push({ pos: offset, top: Math.round(top), page, section });
              if (name === 'pageBreak') forced = 'page';
              else if (name === 'sectionBreak') { if (node.attrs.type === 'continuous') section++; else forced = 'section'; }
              // same layout as before, after the change → the rest stays
              if (startIdx > 0 && offset > dirtyPos && !forced && blocks[i + 1]) {
                const ob = oldByPos.get(offset);
                const nextOff = blocks[i + 1].offset;
                if (ob && ob.top === Math.round(top) && ob.page === page && ob.section === section) {
                  const mine = out.filter((s) => s.pos >= offset && s.pos < nextOff);
                  const oldMine = old.filter((s) => s.pos >= offset && s.pos < nextOff);
                  if (mine.length === oldMine.length && mine.every((s, j) => s.pos === oldMine[j].pos && Math.abs(s.height - oldMine[j].height) < 1)) {
                    out = [...out, ...old.filter((s) => s.pos >= nextOff)];
                    const a2 = prevLayout.findIndex((b) => b.pos === nextOff);
                    if (a2 >= 0) layout = [...layout, ...prevLayout.slice(a2)];
                    const restPages = st.pages.length > pages.length ? st.pages.slice(pages.length) : [];
                    finish(out, [...pages, ...restPages], layout, measured, t0, startIdx);
                    return;
                  }
                }
              }
            }
            if (forced) addPage(forced === 'section' ? section + 1 : section);
            finish(out, pages, layout, measured, t0, startIdx);
          };
          step();
        };
        const schedule = (ms = 120) => { if (job) { clearTimeout(job); job = null; } if (timer) clearTimeout(timer); timer = setTimeout(measure, ms); };
        const markDirty = (at) => view.dispatch(view.state.tr.setMeta(pageKey, { dirtyAt: at }).setMeta('addToHistory', false).setMeta('ofPages', true));
        let lastW = 0;
        const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => { const w = view.dom.offsetWidth; if (w !== lastW) { lastW = w; markDirty(0); } schedule(); }) : null;
        if (ro) ro.observe(view.dom);
        schedule(30);
        // pictures and formulas load later and change heights
        const onLoad = (e) => { if (e.target && e.target.tagName === 'IMG') { markDirty(Math.max(0, posOf(view, e.target) - 1)); schedule(); } };
        view.dom.addEventListener('load', onLoad, true);
        const down = () => { pointerUntil = performance.now() + 650; };
        view.dom.addEventListener('mousedown', down, true);
        ext.storage.remeasure = (at = 0) => { markDirty(at); schedule(20); };
        return {
          update(v, prev) { if (v.state.doc !== prev.doc) schedule(v.state.doc.content.size > 400000 ? 260 : 140); else if (pageKey.getState(v.state).dirty !== Infinity && pageKey.getState(prev).dirty === Infinity) schedule(40); },
          destroy() { if (ro) ro.disconnect(); view.dom.removeEventListener('load', onLoad, true); view.dom.removeEventListener('mousedown', down, true); clearTimeout(timer); if (job) clearTimeout(job); },
        };
      },
    })];
  },
});

function posOf(view, el) { try { return view.posAtDOM(el, 0); } catch { return 0; } }

/** a row spacer inside one cell: paper to the page end, the gap between the pages, paper, the repeated header cell */
function rowWidget(s, i) {
  const d = document.createElement('div');
  d.className = 'of-docs-pgsp of-pgsp-row';
  d.contentEditable = 'false';
  d.style.height = `${s.height}px`;
  const band = (top, h, cls) => { if (h <= 0) return; const b = document.createElement('div'); b.className = cls; b.style.top = `${top}px`; b.style.height = `${h}px`; b.style.left = `${(s.band.left[i] || 0) - 3}px`; b.style.width = `${s.band.width + 6}px`; d.append(b); };
  band(-s.band.pad, s.band.gapTop + s.band.pad, 'of-pgsp-paper');
  band(s.band.gapTop, s.band.gap, 'of-pgsp-gap');
  band(s.band.gapTop + s.band.gap, s.band.after, 'of-pgsp-paper');
  if (s.head && s.head[i] != null) { const h = document.createElement('div'); h.className = 'of-pgsp-head'; h.innerHTML = s.head[i]; h.style.height = `${s.headH}px`; d.append(h); }
  return d;
}

/** line boxes of an element in natural body coordinates (spacers subtracted): [{ top, bottom, x, y }] */
function lineBoxes(el, k, rect, before) {
  const range = document.createRange();
  range.selectNodeContents(el);
  const rs = [...range.getClientRects()].filter((r) => r.height > 0 && r.width > 0);
  const sp = [...el.querySelectorAll('.of-pgsp-line, .of-pgsp-row')].map((x) => { const b = x.getBoundingClientRect(); return { top: b.top, h: b.height }; });
  rs.sort((a, b) => a.top - b.top || a.left - b.left);
  const lines = [];
  for (const r of rs) {
    if (sp.some((x) => Math.abs(x.top - r.top) < 1 && Math.abs(x.h - r.height) < 1)) continue; // the spacer itself
    const last = lines[lines.length - 1];
    if (last && r.top < last.sb - 2) { last.sb = Math.max(last.sb, r.bottom); if (r.left < last.left) { last.left = r.left; last.sy = r.top + r.height / 2; } continue; }
    lines.push({ st: r.top, sb: r.bottom, left: r.left, sy: r.top + r.height / 2 });
  }
  return lines.map((l) => {
    const above = sp.filter((x) => x.top < l.st - 0.5).reduce((a, x) => a + x.h, 0);
    return { top: (l.st - rect.top) / k - before - above, bottom: (l.sb - rect.top) / k - before - above, x: l.left + 1, y: l.sy };
  }).filter((l) => l.bottom > l.top);
}

/** break a text block between lines (possibly over several pages) → { whole, spacers, added } */
function splitLines(el, before, shift, getPage, setPage, env, offset) {
  const { k, rect, view, C, E, next } = env;
  const lines = lineBoxes(el, k, rect, before);
  if (!lines.length) return { whole: true, spacers: [], added: 0 };
  const spacers = []; let added = 0; let page = getPage();
  for (let j = 0; j < lines.length; j++) {
    const lt = lines[j].top + shift + added; const lb = lines[j].bottom + shift + added;
    if (lb <= E(page) + 0.5) continue;
    if (j === 0 && lt > C(page) + 2) return { whole: true, spacers: [], added: 0 };
    if (j === 0) continue; // a single line taller than the page: let it run
    const p = next(page);
    const push = C(p) - lt;
    if (push <= 0.5) { page = p; continue; }
    let pos = null;
    try { const hit = view.posAtCoords({ left: lines[j].x, top: lines[j].y }); pos = hit ? hit.pos : null; } catch { pos = null; }
    if (pos == null || pos <= offset + 1 || (spacers.length && pos <= spacers[spacers.length - 1].pos)) { if (!spacers.length) return { whole: true, spacers: [], added: 0 }; break; }
    spacers.push({ pos, height: r1(push), kind: 'line' });
    added += push; page = p;
  }
  setPage(page);
  return { whole: false, spacers, added };
}

/** break a table between rows, the header row repeated → { whole, spacers, added } */
function splitTable(el, node, offset, before, shift, getPage, setPage, env) {
  const { k, rect, C, E, next, gap, mb, mt } = env;
  const table = el.querySelector('table') || el;
  const trs = [...table.rows];
  const rowPos = []; let p = offset + 1;
  node.forEach((row) => { rowPos.push({ pos: p, row }); p += row.nodeSize; });
  if (trs.length < 2 || rowPos.length !== trs.length) return { whole: true, spacers: [], added: 0 };
  // natural row tops/bottoms: subtract the row spacers above (one per pushed row) and the row's own spacer
  const spOf = (tr) => { const w = tr.querySelector(':scope > * > .of-pgsp-row'); return w ? parseFloat(w.style.height) || 0 : 0; };
  let pushedAbove = 0;
  const nat = trs.map((tr) => { const b = tr.getBoundingClientRect(); const own = spOf(tr); const top = (b.top - rect.top) / k - before - pushedAbove; const res = { top, bottom: (b.bottom - rect.top) / k - before - pushedAbove - own }; pushedAbove += own; return res; });
  const tableTopNat = nat[0].top;
  let headN = 0;
  while (headN < rowPos.length - 1 && rowPos[headN].row.childCount && Array.from({ length: rowPos[headN].row.childCount }, (_, c) => rowPos[headN].row.child(c)).every((c) => c.type.name === 'tableHeader')) headN++;
  const headH = headN ? nat[headN - 1].bottom - tableTopNat : 0;
  const headHtml = headN === 1 ? [...trs[0].cells].map((c) => c.innerHTML.replace(/contenteditable="true"/g, 'contenteditable="false"')) : null;
  const tRect = table.getBoundingClientRect();
  const spacers = []; let added = 0; let page = getPage();
  for (let j = 0; j < trs.length; j++) {
    const rt = nat[j].top + shift + added; const rb = nat[j].bottom + shift + added;
    if (rb <= E(page) + 0.5) continue;
    if (j <= headN) { if (rt > C(page) + 2) return { whole: true, spacers: [], added: 0 }; continue; }
    const np = next(page);
    const repeat = headHtml && headH < (E(np) - C(np)) / 3 ? headH : 0;
    const push = C(np) - rt + repeat;
    if (push <= 0.5) { page = np; continue; }
    const cells = []; let q = rowPos[j].pos + 1;
    rowPos[j].row.forEach((cell) => { cells.push(q + 1); q += cell.nodeSize; });
    const cellEls = [...trs[j].cells];
    if (cellEls.length !== cells.length) return { whole: true, spacers: [], added: 0 };
    const cs0 = getComputedStyle(cellEls[0]);
    const pad = parseFloat(cs0.paddingTop) || 0;
    const left = cellEls.map((c) => { const b = c.getBoundingClientRect(); const cs = getComputedStyle(c); return -Math.round((b.left - tRect.left) / k + (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.borderLeftWidth) || 0)); });
    // positions inside the widget (its top = the row's top + the cell padding)
    const gapTop = Math.round(E(page) + mb(page) - rt - pad);
    spacers.push({ pos: rowPos[j].pos, kind: 'row', cells, height: r1(push), head: repeat ? headHtml : null, headH: repeat ? Math.round(headH) : 0, band: { gapTop, gap, after: Math.round(mt(np)), pad, left, width: Math.round(tRect.width / k) } });
    added += push; page = np;
  }
  setPage(page);
  return { whole: false, spacers, added };
}
