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

