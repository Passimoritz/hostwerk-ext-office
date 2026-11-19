/*
 * View helpers as ProseMirror plugins: find & replace highlighting, formatting marks (¶ · →), print-layout
 * pagination (spacer widgets push blocks to the next page box; blocks taller than a page run through) and the
 * format painter.
 */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

/* ---------------- find & replace ---------------- */
export const searchKey = new PluginKey('ofSearch');
export function findMatches(doc, query, { matchCase = false, wholeWord = false, regex = false } = {}) {
  if (!query) return [];
  let re;
  try {
    const src = regex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    re = new RegExp(wholeWord ? `(?<![\\p{L}\\p{N}_])(?:${src})(?![\\p{L}\\p{N}_])` : src, `gu${matchCase ? '' : 'i'}`);
  } catch { return []; }
  const out = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    // text of the block with a position map (atoms count as one placeholder char)
    let text = ''; const map = [];
    node.forEach((child, offset) => {
      if (child.isText) { for (let k = 0; k < child.text.length; k++) map.push(pos + 1 + offset + k); text += child.text; } else { map.push(pos + 1 + offset); text += '￼'; }
    });
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(text))) {
      if (!m[0].length) { re.lastIndex++; continue; }
      out.push({ from: map[m.index], to: map[m.index + m[0].length - 1] + 1, groups: m });
      if (out.length > 5000) return false;
    }
    return false;
  });
  return out;
}
export const Search = Extension.create({
  name: 'ofSearch',
  addStorage() { return { query: '', options: {}, matches: [], index: -1 }; },
  addCommands() {
    const st = () => this.storage;
    const refresh = (state) => { st().matches = findMatches(state.doc, st().query, st().options); return st().matches; };
    const select = (tr, view, m) => { tr.setSelection(TextSelection.create(tr.doc, m.from, m.to)).scrollIntoView(); };
    return {
      setSearch: (query, options = {}) => ({ state, tr, dispatch }) => {
        st().query = query; st().options = options;
        const ms = refresh(state);
        const from = state.selection.from;
        st().index = ms.length ? Math.max(0, ms.findIndex((m) => m.from >= from)) : -1;
        if (dispatch) dispatch(tr.setMeta(searchKey, { tick: Date.now() }));
        return true;
      },
      findNext: (dir = 1) => ({ state, tr, dispatch, view }) => {
        const ms = refresh(state);
        if (!ms.length) return false;
        const pos = dir > 0 ? state.selection.to : state.selection.from;
        let k = dir > 0 ? ms.findIndex((m) => m.from >= pos) : ms.length - 1 - [...ms].reverse().findIndex((m) => m.to <= pos);
        if (k < 0 || k >= ms.length) k = dir > 0 ? 0 : ms.length - 1;
        st().index = k;
        select(tr, view, ms[k]);
        if (dispatch) dispatch(tr.setMeta(searchKey, { tick: Date.now() }));
        return true;
      },
      replaceCurrent: (replacement) => ({ state, tr, dispatch }) => {
        const ms = refresh(state);
        const { from, to } = state.selection;
        const m = ms.find((x) => x.from === from && x.to === to) || ms[st().index];
        if (!m) return false;
        const text = st().options.regex ? m.groups[0].replace(new RegExp(st().query, st().options.matchCase ? 'u' : 'iu'), replacement) : replacement;
        const marks = state.doc.resolve(m.from + 1).marks();
        if (text) tr.replaceWith(m.from, m.to, state.schema.text(text, marks)); else tr.delete(m.from, m.to);
        const next = findMatches(tr.doc, st().query, st().options).find((x) => x.from >= m.from + text.length);
        if (next) tr.setSelection(TextSelection.create(tr.doc, next.from, next.to)).scrollIntoView();
        if (dispatch) dispatch(tr);
        return true;
      },
      replaceAll: (replacement) => ({ state, tr, dispatch }) => {
        const ms = refresh(state);
        if (!ms.length) return false;
        for (const m of [...ms].reverse()) {
          const text = st().options.regex ? m.groups[0].replace(new RegExp(st().query, st().options.matchCase ? 'u' : 'iu'), replacement) : replacement;
          const marks = state.doc.resolve(Math.min(m.from + 1, state.doc.content.size)).marks();
          if (text) tr.replaceWith(m.from, m.to, state.schema.text(text, marks)); else tr.delete(m.from, m.to);
        }
        this.storage.lastReplaced = ms.length;
        if (dispatch) dispatch(tr);
        return true;
      },
    };
  },
  addProseMirrorPlugins() {
    const ext = this;
    return [new Plugin({
      key: searchKey,
      props: {
        decorations(state) {
          const s = ext.storage;
          if (!s.query) return null;
          const ms = findMatches(state.doc, s.query, s.options);
          s.matches = ms;
          const { from, to } = state.selection;
          return DecorationSet.create(state.doc, ms.map((m) => Decoration.inline(m.from, m.to, { class: m.from === from && m.to === to ? 'of-docs-hit of-docs-hit-cur' : 'of-docs-hit' })));
        },
      },
    })];
  },
});

/* ---------------- formatting marks ---------------- */
export const marksKey = new PluginKey('ofMarks');
export const FormattingMarks = Extension.create({
  name: 'ofMarks',
  addStorage() { return { on: false }; },
  addCommands() { return { setFormattingMarks: (on) => ({ tr, dispatch }) => { this.storage.on = !!on; if (dispatch) dispatch(tr.setMeta(marksKey, on)); return true; } }; },
  addProseMirrorPlugins() {
    const ext = this;
    return [new Plugin({
      key: marksKey,
      props: {
        decorations(state) {
          if (!ext.storage.on) return null;
          const decos = [];
          let budget = 20000;
          state.doc.descendants((node, pos) => {
            if (node.isTextblock) {
              decos.push(Decoration.widget(pos + node.nodeSize - 1, () => { const s = document.createElement('span'); s.className = 'of-docs-pilcrow'; s.textContent = '¶'; return s; }, { side: 1, key: `pil${pos}`, ignoreSelection: true }));
              return true;
            }
            if (node.isText && budget > 0) {
              const t = node.text;
              for (let k = 0; k < t.length && budget > 0; k++) {
                if (t[k] === ' ' || t[k] === ' ') { decos.push(Decoration.inline(pos + k, pos + k + 1, { class: 'of-docs-sp' })); budget--; } else if (t[k] === '\t') { decos.push(Decoration.inline(pos + k, pos + k + 1, { class: 'of-docs-tab' })); budget--; }
              }
            }
            if (node.type.name === 'hardBreak') decos.push(Decoration.widget(pos, () => { const s = document.createElement('span'); s.className = 'of-docs-pilcrow'; s.textContent = '↵'; return s; }, { side: -1, key: `br${pos}` }));
            return true;
          });
          return DecorationSet.create(state.doc, decos);
        },
      },
    })];
  },
});

/* ---------------- pagination (print layout) ---------------- */
export const pageKey = new PluginKey('ofPages');
const MM = 96 / 25.4;
export const Pagination = Extension.create({
  name: 'ofPages',
  addStorage() { return { enabled: false, pageH: 297 * MM, contentH: 252 * MM, gap: 24, onPages: null, pages: 1 }; },
  addCommands() {
    return {
      setPagination: (opts) => ({ tr, dispatch }) => { Object.assign(this.storage, opts); if (dispatch) dispatch(tr.setMeta(pageKey, { reset: true })); return true; },
    };
  },
  addProseMirrorPlugins() {
    const ext = this;
    return [new Plugin({
      key: pageKey,
      state: {
        init: () => ({ spacers: [] }),
        apply(tr, value) {
          const meta = tr.getMeta(pageKey);
          if (meta && meta.spacers) return { spacers: meta.spacers };
          if (meta && meta.reset) return { spacers: [] };
          if (!tr.docChanged) return value;
          return { spacers: value.spacers.map((s) => ({ ...s, pos: tr.mapping.map(s.pos, -1) })) };
        },
      },
      props: {
        decorations(state) {
          if (!ext.storage.enabled) return null;
          const { spacers } = pageKey.getState(state);
          if (!spacers.length) return null;
          return DecorationSet.create(state.doc, spacers.filter((s) => s.pos >= 0 && s.pos <= state.doc.content.size).map((s) => Decoration.widget(s.pos, () => { const d = document.createElement('div'); d.className = 'of-docs-pgsp'; d.style.height = `${s.height}px`; d.contentEditable = 'false'; return d; }, { side: -1, key: `pg${s.pos}_${Math.round(s.height)}`, ignoreSelection: true })));
        },
      },
      view(view) {
        let timer = null;
        const measure = () => {
          timer = null;
          const st = ext.storage;
          if (!st.enabled || !view.dom.isConnected) { if (st.onPages) st.onPages(1); return; }
          const dom = view.dom;
          const rect = dom.getBoundingClientRect();
          const k = dom.offsetWidth ? rect.width / dom.offsetWidth : 1;
          if (!k) return;
          const { spacers } = pageKey.getState(view.state);
          const P = st.pageH + st.gap; const H = st.contentH;
          const out = []; let shift = 0; let forced = false;
          let page = 0;
          view.state.doc.forEach((node, offset) => {
            const el = view.nodeDOM(offset);
            if (!el || el.nodeType !== 1) return;
            const r = el.getBoundingClientRect();
            const before = spacers.filter((s) => s.pos <= offset).reduce((a, s) => a + s.height, 0);
            const cs = getComputedStyle(el);
            const mt = parseFloat(cs.marginTop) || 0; const mb = parseFloat(cs.marginBottom) || 0;
            let top = (r.top - rect.top) / k - before - mt + shift;
            let bottom = (r.bottom - rect.top) / k - before + mb + shift;
            page = Math.max(page, Math.floor(top / P));
            const pageStart = page * P;
            const pageEnd = pageStart + H;
            let push = 0;
            if (forced) { push = (page + 1) * P - top; forced = false; }
            else if (top > pageEnd - 1) push = (page + 1) * P - top;
            else if (bottom > pageEnd + 0.5 && top > pageStart + 2 && bottom - top < H) push = (page + 1) * P - top;
            if (push > 0.5) { out.push({ pos: offset, height: Math.round(push * 10) / 10 }); shift += push; top += push; bottom += push; page = Math.floor(top / P); }
            if (bottom > page * P + H) page = Math.floor(bottom / P);
            if (node.type.name === 'pageBreak') forced = true;
          });
          const pages = Math.max(1, page + 1 + (forced ? 1 : 0));
          const same = out.length === spacers.length && out.every((s, i) => s.pos === spacers[i].pos && Math.abs(s.height - spacers[i].height) < 1);
          if (!same) view.dispatch(view.state.tr.setMeta(pageKey, { spacers: out }).setMeta('addToHistory', false));
          st.pages = pages;
          if (st.onPages) st.onPages(pages);
        };
        const schedule = () => { if (!timer) timer = setTimeout(measure, 120); };
        const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(schedule) : null;
        if (ro) ro.observe(view.dom);
        schedule();
        const later = [setTimeout(schedule, 400), setTimeout(schedule, 1500)];
        // images load later
        const onLoad = (e) => { if (e.target && e.target.tagName === 'IMG') schedule(); };
        view.dom.addEventListener('load', onLoad, true);
        return {
          update(v, prev) { if (v.state.doc !== prev.doc || pageKey.getState(v.state) !== pageKey.getState(prev)) { schedule(); if (pageKey.getState(v.state).spacers.length === 0) later.push(setTimeout(schedule, 400)); } },
          destroy() { later.forEach(clearTimeout); if (ro) ro.disconnect(); view.dom.removeEventListener('load', onLoad, true); clearTimeout(timer); },
        };
      },
    })];
  },
});
