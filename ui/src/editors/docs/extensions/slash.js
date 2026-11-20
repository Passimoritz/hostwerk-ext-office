/*
 * "/" command menu: typing "/" at the start of a paragraph or after a space opens the menu at the caret; the text
 * after it filters ("/tab" → Table), ↑/↓ choose, Enter/Tab inserts, Esc or a space with nothing found closes. The
 * menu itself is React (extras.jsx); this plugin only tracks where it is open and hands over the keys.
 * Works with Track changes on (the "/" may be typed through another plugin – it is detected from the transaction).
 */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';

export const slashKey = new PluginKey('ofSlash');

export const Slash = Extension.create({
  name: 'ofSlash',
  addStorage() { return { onChange: null, onKey: null }; },
  addCommands() {
    return {
      openSlash: (at) => ({ tr, dispatch, state }) => { if (dispatch) dispatch(tr.setMeta(slashKey, { open: true, from: at != null ? at : state.selection.from - 1 })); return true; },
      closeSlash: () => ({ tr, dispatch }) => { if (dispatch) dispatch(tr.setMeta(slashKey, 'close')); return true; },
    };
  },
  addProseMirrorPlugins() {
    const ext = this;
    return [new Plugin({
      key: slashKey,
      state: {
        init: () => null,
        apply(tr, prev, oldState, newState) {
          const meta = tr.getMeta(slashKey);
          if (meta === 'close') return null;
          if (meta && meta.open) return { from: meta.from, query: '' };
          let cur = prev;
          // a "/" just typed at a word start
          if (!cur && tr.docChanged && newState.selection.empty && newState.doc.content.size - oldState.doc.content.size === 1 && !tr.getMeta('uiEvent') && !tr.getMeta('ofPages')) {
            const pos = newState.selection.from;
            const $p = newState.doc.resolve(pos);
            if ($p.parent.isTextblock && !$p.parent.type.spec.code && $p.parentOffset >= 1) {
              const before = $p.parent.textBetween(0, $p.parentOffset, '\0', '\0');
              if (before.endsWith('/') && (before.length === 1 || /\s$/.test(before.slice(0, -1)))) cur = { from: pos - 1, query: '' };
            }
          }
          if (!cur) return null;
          const from = prev ? tr.mapping.map(cur.from) : cur.from;
          const sel = newState.selection;
          if (!sel.empty || sel.from <= from || from < 0 || from >= newState.doc.content.size) return null;
          const $f = newState.doc.resolve(from); const $s = newState.doc.resolve(sel.from);
          if ($f.parent !== $s.parent) return null;
          const text = newState.doc.textBetween(from, sel.from, '\n', '\0');
          if (!text.startsWith('/') || text.length > 32 || /\s\s|\n|\0/.test(text)) return null;
          return { from, query: text.slice(1) };
        },
      },
      props: {
        handleKeyDown(view, e) {
          const st = slashKey.getState(view.state);
          if (!st || !ext.storage.onKey) return false;
          return ext.storage.onKey(e, st) === true;
        },
      },
      view() {
        let last = null;
        return {
          update(view) {
            const st = slashKey.getState(view.state);
            if (st === last) return;
            last = st;
            let coords = null;
            if (st) { try { coords = view.coordsAtPos(st.from); } catch { coords = null; } }
            if (ext.storage.onChange) ext.storage.onChange(st ? { ...st, coords } : null);
          },
          destroy() { if (ext.storage.onChange) ext.storage.onChange(null); },
        };
      },
    })];
  },
});
