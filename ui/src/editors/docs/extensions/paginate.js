/*
 * Print layout pagination – real page breaks inside one editing flow:
 *   · blocks that do not fit move to the next page (spacer widget before them)
 *   · long paragraphs, lists and quotes break between two LINES (a block spacer widget at the first text position of
 *     the line that no longer fits – the line starts there anyway, so nothing re-wraps)
 *   · tables break between ROWS, and inside a row between its lines when the row may break (Word's "allow row to
 *     break across pages"; rows marked cantSplit move whole, rows taller than a page always break); header rows are
 *     repeated at the top of every page; nothing of a table ever lies in the gap or the margins of a page
 *   · page breaks and section breaks start a new page; every section has its own page size, orientation and margins
 *     (blocks of a section with another text width get that width)
 * The decisions are made by the pure page model (pagemodel.js) on the NATURAL layout: text blocks subtract their
 * spacers, a table is measured with its spacers hidden for a moment (one class on its wrapper, read synchronously).
 * Speed: only the part after the first change is measured again, and measuring stops as soon as the layout is the
 * same as before (typing in a 300-page document measures a page or two); long first measurements run in slices so
 * the page stays responsive. Heights that change without an edit (fonts, pictures, formulas loading) are noticed by
 * a ResizeObserver and measured again from the first block that moved.
 */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { pager, planLines, planTable } from './pagemodel.js';

export const pageKey = new PluginKey('ofPages');
const MM = 96 / 25.4;
const SLICE_MS = 12;
const r1 = (v) => Math.round(v * 10) / 10;
const growOf = (s) => (s.grow != null ? s.grow : s.height);

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
        if (s.kind === 'row') (s.cells || []).forEach((c, i) => { if (c >= 0 && c <= size) decos.push(Decoration.widget(c, () => rowWidget(s, i), { side: -1, key: `pr${c}_${Math.round(s.height)}_${i ? 0 : bandKey(s.band)}`, ignoreSelection: true })); });
        else if (s.kind === 'cline') decos.push(Decoration.widget(s.pos, () => clineWidget(s), { side: -1, key: `pc${s.pos}_${Math.round(s.height)}_${bandKey(s.band)}`, ignoreSelection: true, marks: [] }));
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
        let settledH = 0; // the editor height after the last layout (an other height without an edit = something loaded)
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
          const same = out.length === cur.spacers.length && out.every((s, i) => s.pos === cur.spacers[i].pos && s.kind === cur.spacers[i].kind && Math.abs(s.height - cur.spacers[i].height) < 1 && bandKey(s.band) === bandKey(cur.spacers[i].band));
          view.dispatch(view.state.tr.setMeta(pageKey, { spacers: same ? cur.spacers : out, dirty: Infinity, layout }).setMeta('addToHistory', false).setMeta('ofPages', true));
          settledH = view.dom.offsetHeight;
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
          // rendered spacers before a position (their growth subtracted → natural layout)
          const prefix = []; let acc = 0;
          for (const s of old) { acc += growOf(s); prefix.push(acc); }
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
          let out; let layout; let page; let section; let P;
          if (startIdx > 0) {
            const keepTo = blocks[startIdx].offset;
            out = old.filter((s) => s.pos < keepTo);
            layout = prevLayout.slice(0, at);
            ({ section, page } = prevLayout[at]);
            P = pager(secs, gap, st.pages.slice(0, page + 1));
          } else { out = []; layout = []; section = 0; P = pager(secs, gap); P.add(0); page = 0; }
          let shift = out.reduce((a, s) => a + growOf(s), 0);
          let forced = null; let i = startIdx; let measured = 0;
          const oldByPos = new Map(prevLayout.map((b) => [b.pos, b]));
          const dirtyPos = pst.dirty;
          const next = (p) => P.next(p, section);
          const env = { k, rect, view, P, gap };
          const pushBlock = (offset, top) => { const p = next(page); const push = P.C(p) - top; if (push > 0.5) { out.push({ pos: offset, height: r1(push), kind: 'block' }); shift += push; } page = p; return Math.max(0, push); };
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
                const p = P.add(section);
                const push = P.C(p) - top;
                if (push > 0.5) { out.push({ pos: offset, height: r1(push), kind: 'block' }); shift += push; top += push; bottom += push; }
                page = p; forced = null;
              }
              while (top >= P.E(page) - 1 && P.pages.length < 5000) {
                page = next(page);
                if (top < P.C(page)) { const push = P.C(page) - top; out.push({ pos: offset, height: r1(push), kind: 'block' }); shift += push; top += push; bottom += push; }
              }
              const contentH = P.height(page);
              const tall = bottom - top > contentH - 4;
              const startOfPage = top <= P.C(page) + 2;
              if (bottom > P.E(page) + 0.5) {
                if (name === 'table') {
                  const natBottom = bottom - shift;
                  const t = tableSpacers(el, node, offset, natTop + mt, env, { shift: () => shift, page: () => page, push: () => { const push = pushBlock(offset, top); top += push; }, startOfPage });
                  for (const s of t.spacers) out.push(s);
                  shift += t.added; page = t.page;
                  bottom = natBottom + shift;
                } else if (['paragraph', 'bulletList', 'orderedList', 'taskList', 'blockquote', 'callout', 'controlBlock', 'codeBlock'].includes(name)) {
                  const lines = lineBoxes(el, k, rect, before);
                  let res = planLines(lines, { shift, page, P });
                  if (res.whole && !startOfPage) {
                    const push = pushBlock(offset, top); top += push; bottom += push;
                    res = tall ? planLines(lines, { shift, page, P }) : { spacers: [], added: 0, page };
                  }
                  for (const sp of res.spacers || []) { const pos = posAtLine(view, el, lines[sp.line], k, rect, before); if (pos == null || pos <= offset + 1 || (out.length && out[out.length - 1].pos >= pos && out[out.length - 1].kind === 'line')) break; out.push({ pos, height: sp.height, kind: 'line' }); shift += sp.height; bottom += sp.height; }
                  if (res.page != null) page = Math.max(page, res.page);
                } else if (!startOfPage && !tall) { const push = pushBlock(offset, top); top += push; bottom += push; }
              } else if (name === 'heading' && bottom > P.E(page) - 30 && !startOfPage && blocks[i + 1]) {
                // keep a heading together with the text after it
                const push = pushBlock(offset, top); top += push; bottom += push;
              }
              while (bottom > P.E(page) + 1 && P.pages.length < 5000) { page = next(page); }
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
                    const restPages = st.pages.length > P.pages.length ? st.pages.slice(P.pages.length) : [];
                    finish(out, [...P.pages, ...restPages], layout, measured, t0, startIdx);
                    return;
                  }
                }
              }
            }
            if (forced) P.add(forced === 'section' ? section + 1 : section);
            finish(out, P.pages, layout, measured, t0, startIdx);
          };
          step();
        };
        const schedule = (ms = 120) => { if (job) { clearTimeout(job); job = null; } if (timer) clearTimeout(timer); timer = setTimeout(measure, ms); };
        const markDirty = (at) => view.dispatch(view.state.tr.setMeta(pageKey, { dirtyAt: at }).setMeta('addToHistory', false).setMeta('ofPages', true));
        /** the first top-level block whose top is not where the last layout put it → measure again from there */
        const firstMoved = () => {
          const lay = pageKey.getState(view.state).layout || [];
          if (!lay.length) return 0;
          const rect = view.dom.getBoundingClientRect();
          const k = view.dom.offsetWidth ? rect.width / view.dom.offsetWidth : 1;
          const off = (b) => { const el = view.nodeDOM(b.pos); if (!el || el.nodeType !== 1) return 0; const r = el.getBoundingClientRect(); return (r.top - rect.top) / k - margins(el)[0]; };
          // the layout keeps each block's top as it is rendered (spacers included)
          const moved = (b) => Math.abs(off(b) - b.top) > 1.5;
          let lo = 0; let hi = lay.length - 1; let found = -1;
          while (lo <= hi) { const mid = (lo + hi) >> 1; if (moved(lay[mid])) { found = mid; hi = mid - 1; } else lo = mid + 1; }
          if (found < 0) return lay[lay.length - 1].pos;
          return Math.max(0, lay[Math.max(0, found - 1)].pos);
        };
        let lastW = 0;
        const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => {
          const w = view.dom.offsetWidth;
          if (w !== lastW) { lastW = w; markDirty(0); schedule(); return; }
          // same width, other height without an edit: a font, picture or formula arrived
          const st = pageKey.getState(view.state);
          if (st.dirty === Infinity && settledH && Math.abs(view.dom.offsetHeight - settledH) > 1 && !timer && !job) { markDirty(firstMoved()); schedule(60); }
        }) : null;
        if (ro) ro.observe(view.dom);
        schedule(30);
        // pictures and formulas load later and change heights
        const onLoad = (e) => { if (e.target && e.target.tagName === 'IMG') { markDirty(Math.max(0, posOf(view, e.target) - 1)); schedule(); } };
        view.dom.addEventListener('load', onLoad, true);
        const onFonts = () => { if (!view.dom.isConnected) return; markDirty(0); schedule(60); };
        const fonts = typeof document !== 'undefined' && document.fonts;
        if (fonts && fonts.addEventListener) fonts.addEventListener('loadingdone', onFonts);
        const down = () => { pointerUntil = performance.now() + 650; };
        view.dom.addEventListener('mousedown', down, true);
        ext.storage.remeasure = (at = 0) => { markDirty(at); schedule(20); };
        return {
          update(v, prev) { if (v.state.doc !== prev.doc) schedule(v.state.doc.content.size > 400000 ? 260 : 140); else if (pageKey.getState(v.state).dirty !== Infinity && pageKey.getState(prev).dirty === Infinity) schedule(40); },
          destroy() { if (ro) ro.disconnect(); view.dom.removeEventListener('load', onLoad, true); view.dom.removeEventListener('mousedown', down, true); if (fonts && fonts.removeEventListener) fonts.removeEventListener('loadingdone', onFonts); clearTimeout(timer); if (job) clearTimeout(job); },
        };
      },
    })];
  },
});

function posOf(view, el) { try { return view.posAtDOM(el, 0); } catch { return 0; } }
const bandKey = (b) => (b ? `${Math.round(b.cutTop || 0)}_${Math.round(b.cut)}_${Math.round(b.C)}_${Math.round(b.head || 0)}_${b.headId || ''}` : '');

/* ---------------- widgets ---------------- */
/**
 * the band drawn across the whole table where it breaks: paper to the page end (the row's bottom border at the
 * break), the gap between the pages, paper of the next page's top margin, the repeated header rows, the row's top
 * border where it goes on. Its numbers are flow coordinates; `origin` = flow y of the widget's own top.
 */
function bandEl(band, origin) {
  const wrap = document.createElement('div');
  wrap.className = 'of-pgsp-band';
  const piece = (top, h, cls) => { if (h <= 0) return null; const b = document.createElement('div'); b.className = cls; b.style.top = `${top - origin}px`; b.style.height = `${h}px`; wrap.append(b); return b; };
  const from = band.cutTop != null ? band.cutTop : band.at + 1;
  const p1 = piece(from, band.cut - from, 'of-pgsp-paper');
  if (p1 && band.cutTop != null) p1.classList.add('of-pgsp-cutline');
  piece(band.cut, band.next - band.cut, 'of-pgsp-gap');
  const p2 = piece(band.next, band.C - band.next, 'of-pgsp-paper');
  if (band.head && band.headHtml) {
    const h = document.createElement('div');
    h.className = 'of-pgsp-head';
    h.style.top = `${band.C - origin}px`;
    h.style.height = `${band.head}px`;
    h.innerHTML = band.headHtml;
    wrap.append(h);
  } else if (band.cutTop != null) {
    const l = piece(band.C - 1, 1, 'of-pgsp-paper of-pgsp-cutline');
    void l;
  }
  void p2;
  // full table width: measured once the widget is in the document (left edge of the table, scale of the zoom)
  requestAnimationFrame(() => {
    const table = wrap.closest('table');
    const pm = wrap.closest('.ProseMirror');
    if (!table || !pm) return;
    const k = pm.offsetWidth ? pm.getBoundingClientRect().width / pm.offsetWidth : 1;
    const tb = table.getBoundingClientRect(); const me = wrap.parentElement.getBoundingClientRect();
    wrap.style.left = `${Math.round((tb.left - me.left) / k) - 2}px`;
    wrap.style.width = `${Math.round(tb.width / k) + 4}px`;
    // a row widget sits below the cell padding (or lower when the cell is aligned to its middle/bottom)
    if (band.rowTopEl) { const tr = wrap.closest('tr'); if (tr) { const d = (me.top - tr.getBoundingClientRect().top) / k - band.off; if (Math.abs(d) > 0.5) wrap.style.transform = `translateY(${-d}px)`; } }
  });
  return wrap;
}
/** a row spacer inside one cell (the row moves whole to the next page); the first cell draws the band */
function rowWidget(s, i) {
  const d = document.createElement('div');
  d.className = 'of-docs-pgsp of-pgsp-row';
  d.contentEditable = 'false';
  d.style.height = `${s.height}px`;
  if (i === 0 && s.band) d.append(bandEl({ ...s.band, rowTopEl: true }, s.band.at + s.band.off));
  return d;
}
/** a line spacer inside a table cell (the cell's text goes on on the next page) */
function clineWidget(s) {
  const d = document.createElement('span');
  d.className = 'of-docs-pgsp of-pgsp-cline';
  d.contentEditable = 'false';
  d.style.height = `${s.height}px`;
  if (s.band) d.append(bandEl(s.band, s.band.at));
  return d;
}

/* ---------------- measuring ---------------- */
/** line boxes of an element in natural body coordinates (spacers subtracted): [{ top, bottom, x, y }] */
function lineBoxes(el, k, rect, before) {
  const rs = inlineRects(el);
  const sp = [...el.querySelectorAll('.of-pgsp-line, .of-pgsp-row, .of-pgsp-cline')].map((x) => { const b = x.getBoundingClientRect(); return { top: b.top, h: b.height }; });
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
    return { top: (l.st - rect.top) / k - before - above, bottom: (l.sb - rect.top) / k - before - above, x: l.left + 1, y: l.sy, st: l.st };
  }).filter((l) => l.bottom > l.top);
}

/**
 * the document position where a line starts – found by walking the text (works off screen, unlike posAtCoords,
 * which needs the point inside the viewport)
 */
function posAtLine(view, el, line, k, rect, before) {
  if (!line) return null;
  const y = line.st != null ? line.st : line.y;
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement && n.parentElement.closest('.of-docs-pgsp, [contenteditable="false"]') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  const range = document.createRange();
  let n;
  while ((n = walker.nextNode())) {
    if (!n.data.length) continue;
    range.selectNodeContents(n);
    const rs = range.getClientRects();
    if (!rs.length || rs[rs.length - 1].bottom <= y + 1) continue;
    // first character whose box starts on (or below) the line
    let lo = 0; let hi = n.data.length - 1; let found = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      range.setStart(n, mid); range.setEnd(n, mid + 1);
      const r = range.getClientRects()[0] || range.getBoundingClientRect();
      if (r.top >= y - 1) { found = mid; hi = mid - 1; } else lo = mid + 1;
    }
    if (found < 0) continue;
    try { return view.posAtDOM(n, found); } catch { return null; }
  }
  // no text on that line (a picture, a formula): fall back to the point
  try { const hit = view.posAtCoords({ left: line.x, top: line.y }); return hit ? hit.pos : null; } catch { return null; }
}

/**
 * a table that does not fit: measured without its spacers (a class on the wrapper hides them while reading), planned by
 * the page model, turned into row / cell-line spacers. Returns { spacers, shift, page, top }.
 */
function tableSpacers(el, node, offset, base, env, at) {
  const { k, view, P } = env;
  const wrap = el; const table = el.querySelector('table') || el;
  let page = at.page();
  const spacers = []; let added = 0;
  wrap.classList.add('of-pg-nat');
  try {
    const wr = wrap.getBoundingClientRect();
    const trs = [...table.rows];
    const rowNodes = []; let p = offset + 1;
    node.forEach((row) => { rowNodes.push({ pos: p, row }); p += row.nodeSize; });
    if (!trs.length || rowNodes.length !== trs.length) return { spacers, added, page };
    // base = natural flow y of the wrapper's top (spacers above subtracted by the caller)
    const ys = (r) => base + (r - wr.top) / k;
    const rows = trs.map((tr, j) => {
      const b = tr.getBoundingClientRect();
      const ra = rowNodes[j].row.attrs || {};
      const cells = [...tr.cells].map((td) => {
        const tb = td.getBoundingClientRect();
        const cs = getComputedStyle(td);
        return {
          el: td,
          bottom: ys(tb.bottom),
          lines: () => cellLines(td).map((l) => ({ top: ys(l.top), bottom: ys(l.bottom), st: l.top, x: l.left + 1, y: (l.top + l.bottom) / 2, el: td })),
          padTop: (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.borderTopWidth) || 0),
          contentBottom: () => { let m = tb.top + (parseFloat(cs.paddingTop) || 0); for (const c of td.children) { if (c.classList.contains('of-docs-pgsp')) continue; m = Math.max(m, c.getBoundingClientRect().bottom + (parseFloat(getComputedStyle(c).marginBottom) || 0)); } return ys(m + (parseFloat(cs.paddingBottom) || 0)); },
        };
      });
      for (const c of cells) c.bottom = c.contentBottom();
      return { top: ys(b.top), bottom: ys(b.bottom), cantSplit: !!ra.cantSplit, exact: ra.heightRule === 'exact', cells };
    });
    let headN = 0;
    while (headN < rowNodes.length - 1 && rowNodes[headN].row.childCount && Array.from({ length: rowNodes[headN].row.childCount }, (_, c) => rowNodes[headN].row.child(c)).every((c) => c.type.name === 'tableHeader')) headN++;
    const headH = headN ? rows[headN - 1].bottom - rows[0].top : 0;
    const headHtml = headN ? headClone(table, trs.slice(0, headN)) : null;
    let res = planTable(rows, { shift: at.shift(), page, P, headN, headH });
    if (res.whole) {
      // the header (or a first row that must stay with it) does not fit here: the table starts on the next page
      if (!at.startOfPage) { at.push(); page = at.page(); }
      res = planTable(rows, { shift: at.shift(), page, P, headN, headH, force: true });
    }
    for (const it of res.items) {
      const row = rowNodes[it.row]; const tr = trs[it.row];
      if (it.kind === 'row') {
        const cells = []; let q = row.pos + 1;
        row.row.forEach((cell) => { cells.push(q + 1); q += cell.nodeSize; });
        if (cells.length !== tr.cells.length) continue;
        const off = rows[it.row].cells[0] ? rows[it.row].cells[0].padTop : 0;
        spacers.push({ pos: row.pos, kind: 'row', cells, height: r1(it.height), grow: r1(it.height), band: { at: r1(it.at), off: r1(off), cut: r1(it.cut), next: r1(it.next), C: r1(it.C), head: Math.round(it.head), headHtml: it.head ? headHtml : null } });
      } else {
        const c = rows[it.row].cells[it.cell];
        const ls = typeof c.lines === 'function' ? c.lines() : c.lines;
        const ln = ls[it.line];
        const pos = ln ? posAtLine(view, c.el, ln, k, null, 0) : null;
        if (pos == null) continue;
        spacers.push({ pos, kind: 'cline', height: r1(it.height), grow: it.grow != null ? r1(it.grow) : 0, band: it.band ? { at: r1(it.at), cutTop: r1(it.band.cutTop), cut: r1(it.band.cut), next: r1(it.band.next), C: r1(it.band.C), head: Math.round(it.band.head), headHtml: it.band.head ? headHtml : null } : null });
      }
    }
    spacers.sort((a, b) => a.pos - b.pos);
    added = res.added;
    page = res.page;
  } finally { wrap.classList.remove('of-pg-nat'); }
  return { spacers, added, page };
}

/**
 * the boxes of the text (and inline pictures/formulas) inside an element – a range over the whole element would also
 * return the boxes of its paragraphs (one rect over all their lines)
 */
function inlineRects(el) {
  const out = [];
  const range = document.createRange();
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
    acceptNode: (n) => {
      if (n.nodeType === 1) return n.classList && n.classList.contains('of-docs-pgsp') ? NodeFilter.FILTER_REJECT : (['IMG', 'svg', 'math', 'BR'].includes(n.tagName) || (n.getAttribute && n.getAttribute('contenteditable') === 'false')) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
      return n.data && n.data.length ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
    },
  });
  let n;
  while ((n = walker.nextNode())) {
    if (n.nodeType === 1) { if (n.tagName !== 'BR') { const r = n.getBoundingClientRect(); if (r.height > 0) out.push(r); } continue; }
    range.selectNodeContents(n);
    for (const r of range.getClientRects()) if (r.height > 0 && r.width > 0) out.push(r);
  }
  // an empty paragraph still has a line
  if (!out.length) for (const p of el.querySelectorAll('p, h1, h2, h3, h4, h5, h6')) { const r = p.getBoundingClientRect(); if (r.height > 0) out.push(r); }
  return out;
}
/** text lines of a table cell (client coordinates), spacers hidden */
function cellLines(td) {
  const rs = inlineRects(td).sort((a, b) => a.top - b.top || a.left - b.left);
  const lines = [];
  for (const r of rs) {
    const last = lines[lines.length - 1];
    if (last && r.top < last.bottom - 2) { last.bottom = Math.max(last.bottom, r.bottom); if (r.left < last.left) last.left = r.left; continue; }
    lines.push({ top: r.top, bottom: r.bottom, left: r.left });
  }
  return lines;
}

/** the header rows as a small read-only table with the same classes, look and column widths */
function headClone(table, trs) {
  const cols = table.querySelector('colgroup');
  const rows = trs.map((tr) => {
    const c = tr.cloneNode(true);
    c.querySelectorAll('.of-docs-pgsp').forEach((x) => x.remove());
    c.querySelectorAll('[contenteditable]').forEach((x) => x.setAttribute('contenteditable', 'false'));
    c.querySelectorAll('.column-resize-handle').forEach((x) => x.remove());
    return c.outerHTML;
  }).join('');
  const style = (table.getAttribute('style') || '').replace(/margin-left:[^;]*;?/g, '');
  return `<table class="${table.className} of-pgsp-headtable" style="${style.replace(/"/g, '&quot;')};width:100%;margin:0">${cols ? cols.outerHTML : ''}<tbody>${rows}</tbody></table>`;
}
