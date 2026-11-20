/*
 * Word paragraph and character properties in the editor (io/wordprops.js): indents right, exact / at-least line
 * spacing, keep with next, keep lines together, widow control, outline level, paragraph borders and shading, tab stops
 * with leaders, drop caps; character spacing, scale, position, caps / small caps, double strike, hidden text and the
 * language of a text (spell check follows it). Tab stops are laid out by a small plugin that measures each tab.
 */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { paraCss, runCss } from '../io/wordprops.js';

const keep = () => ({ default: null, parseHTML: () => null, renderHTML: () => ({}), keepOnSplit: true });
const once = (name) => ({ default: null, parseHTML: () => null, keepOnSplit: name !== 'dropCap' && name !== 'pageBreakBefore' });

export const WordFormat = Extension.create({
  name: 'wordFormat',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          // one attribute renders the CSS of all (tiptap hands every renderHTML the whole attribute set)
          rightIndent: { ...once('rightIndent'), keepOnSplit: true, renderHTML: (a) => { const s = paraCss(a); const o = {}; if (s) o.style = s; if (a.dropCap) { o['data-dropcap'] = String(a.dropCap.lines || 3); o.class = 'of-dropcap'; } if (a.tabs && a.tabs.length) o['data-tabs'] = '1'; if (a.outline != null) o['data-outline'] = String(a.outline + 1); return o; } },
          lineRule: keep(), lineExact: keep(), keepNext: keep(), keepLines: keep(), widow: keep(), outline: keep(), pborders: keep(), pshading: keep(), tabs: keep(), dropCap: { ...keep(), keepOnSplit: false },
        },
      },
      {
        types: ['textStyle'],
        attributes: {
          letterSpacing: { default: null, parseHTML: () => null, renderHTML: (a) => { const s = runCss(a); const o = {}; if (s) o.style = s; if (a.hidden) o['data-hidden'] = '1'; if (a.lang) { o.lang = a.lang; } return o; } },
          scale: keep(), position: keep(), caps: keep(), dstrike: keep(), hidden: keep(), lang: keep(),
        },
      },
    ];
  },
  addCommands() {
    return {
      /** character format of the selection (textStyle attrs) – null clears one */
      setRunFormat: (attrs) => ({ chain }) => chain().setMark('textStyle', attrs).removeEmptyTextStyle().run(),
    };
  },
  addProseMirrorPlugins() {
    return [tabStops()];
  },
});

/* ---------------- tab stops ---------------- */
const tabKey = new PluginKey('ofTabs');
/**
 * every tab character of a paragraph with own tab stops gets an inline width up to its stop (left/right/centre/
 * decimal) and the leader drawn into the gap; measured after layout, re-measured when the paragraph changes
 */
function tabStops() {
  return new Plugin({
    key: tabKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, set) { const m = tr.getMeta(tabKey); if (m) return m; return set.map(tr.mapping, tr.doc); },
    },
    props: { decorations(state) { return tabKey.getState(state); } },
    view(view) {
      let raf = 0;
      const run = () => {
        raf = 0;
        const { state } = view;
        const decos = [];
        state.doc.descendants((node, pos) => {
          if (node.type.name !== 'paragraph' && node.type.name !== 'heading') return true;
          const stops = node.attrs.tabs;
          if (!stops || !stops.length || node.textContent.indexOf('\t') < 0) return false;
          const el = view.nodeDOM(pos);
          if (!el || !el.getBoundingClientRect) return false;
          const k = view.dom.offsetWidth ? view.dom.getBoundingClientRect().width / view.dom.offsetWidth : 1;
          const box = el.getBoundingClientRect();
          const cs = getComputedStyle(el);
          const left0 = box.left + (parseFloat(cs.paddingLeft) || 0);
          const sorted = [...stops].sort((a, b) => a.pos - b.pos);
          node.forEach((child, off) => {
            if (!child.isText) return;
            let i = child.text.indexOf('\t');
            while (i >= 0) {
              const at = pos + 1 + off + i;
              try {
                const c = view.coordsAtPos(at);
                const x = (c.left - left0) / k; // px from the paragraph's text start
                const xpt = x * 0.75 + (node.attrs.indent || 0);
                const stop = sorted.find((s) => s.pos > xpt + 0.5);
                if (stop) {
                  // text after the tab up to the next tab or the end (right/centre/decimal stops align it)
                  const rest = child.text.slice(i + 1).split('\t')[0];
                  let w = (stop.pos - xpt) / 0.75;
                  if (stop.align !== 'left' && rest) {
                    const a2 = view.coordsAtPos(at + 1); const b2 = view.coordsAtPos(at + 1 + rest.length);
                    let tw = (b2.left - a2.left) / k;
                    if (stop.align === 'decimal') { const d = rest.search(/[.,]/); if (d >= 0) { const c3 = view.coordsAtPos(at + 1 + d); tw = (c3.left - a2.left) / k; } }
                    w -= stop.align === 'center' ? tw / 2 : tw;
                  }
                  if (w > 1) decos.push(Decoration.inline(at, at + 1, { class: `of-tab${stop.leader ? ` of-tab-${stop.leader}` : ''}`, style: `display:inline-block;width:${Math.round(w)}px` }));
                }
              } catch { /* not laid out */ }
              i = child.text.indexOf('\t', i + 1);
            }
          });
          return false;
        });
        const cur = tabKey.getState(view.state);
        const next = decos.length ? DecorationSet.create(view.state.doc, decos) : DecorationSet.empty;
        if (cur === DecorationSet.empty && next === DecorationSet.empty) return;
        // only when it changed (keys = position + width)
        const sig = (s) => s.find().map((d) => `${d.from}:${d.type.attrs && d.type.attrs.style}`).join('|');
        if (sig(cur) !== sig(next)) view.dispatch(view.state.tr.setMeta(tabKey, next).setMeta('addToHistory', false).setMeta('ofPages', true));
      };
      const schedule = () => { if (!raf) raf = requestAnimationFrame(run); };
      schedule();
      return { update(v, prev) { if (v.state.doc !== prev.doc) { const has = []; v.state.doc.descendants((n) => { if (n.attrs && n.attrs.tabs && n.attrs.tabs.length) { has.push(1); return false; } return !n.isTextblock; }); if (has.length) schedule(); } }, destroy() { if (raf) cancelAnimationFrame(raf); } };
    },
  });
}
