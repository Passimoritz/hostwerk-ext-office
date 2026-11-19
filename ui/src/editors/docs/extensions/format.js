/*
 * Formatting extensions: font size, paragraph format (line height, spacing, indents, paragraph style), list styles,
 * cell background, invisible "keep" attributes for lossless saving (ox/oo), run properties we do not edit (runKeep)
 * and Word-like keyboard shortcuts that never take Hostwerk's global ones (Ctrl+K, Ctrl+J, Ctrl+1…9, Ctrl+, …).
 */
import { Extension, Mark } from '@tiptap/core';

const keepAttr = (name) => ({ [name]: { default: null, parseHTML: () => null, renderHTML: () => ({}), keepOnSplit: false } });
const ptStyle = (v) => (v == null ? null : `${v}pt`);
const parsePt = (s) => { if (!s) return null; const m = /(-?[\d.]+)(pt|px|cm|mm|em)?/.exec(s); if (!m) return null; const n = Number(m[1]); return Math.round(({ pt: n, px: n * 0.75, cm: n * 28.35, mm: n * 2.835, em: n * 11 }[m[2] || 'px']) * 10) / 10; };

export const FontSize = Extension.create({
  name: 'fontSize',
  addGlobalAttributes() {
    return [{ types: ['textStyle'], attributes: { fontSize: { default: null, parseHTML: (el) => el.style.fontSize || null, renderHTML: (a) => (a.fontSize ? { style: `font-size:${a.fontSize}` } : {}) } } }];
  },
  addCommands() {
    return {
      setFontSize: (size) => ({ chain }) => chain().setMark('textStyle', { fontSize: size }).run(),
      unsetFontSize: () => ({ chain }) => chain().setMark('textStyle', { fontSize: null }).removeEmptyTextStyle().run(),
    };
  },
});

export const ParagraphFormat = Extension.create({
  name: 'paragraphFormat',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          lineHeight: { default: null, parseHTML: (el) => { const v = el.style.lineHeight; return v && /^[\d.]+$/.test(v) ? Number(v) : null; }, renderHTML: (a) => (a.lineHeight ? { style: `line-height:${a.lineHeight}` } : {}) },
          spaceBefore: { default: null, parseHTML: (el) => parsePt(el.style.marginTop), renderHTML: (a) => (a.spaceBefore != null ? { style: `margin-top:${ptStyle(a.spaceBefore)}` } : {}) },
          spaceAfter: { default: null, parseHTML: (el) => parsePt(el.style.marginBottom), renderHTML: (a) => (a.spaceAfter != null ? { style: `margin-bottom:${ptStyle(a.spaceAfter)}` } : {}) },
          indent: { default: null, parseHTML: (el) => parsePt(el.style.marginLeft), renderHTML: (a) => (a.indent ? { style: `margin-left:${ptStyle(a.indent)}` } : {}) },
          firstLine: { default: null, parseHTML: (el) => parsePt(el.style.textIndent), renderHTML: (a) => (a.firstLine ? { style: `text-indent:${ptStyle(a.firstLine)}` } : {}) },
          styleId: { default: null, parseHTML: () => null, renderHTML: (a) => (a.styleId ? { 'data-style-id': a.styleId } : {}) },
          ...keepAttr('ox'), ...keepAttr('oo'),
        },
      },
      { types: ['paragraph'], attributes: { pstyle: { default: null, parseHTML: (el) => el.getAttribute('data-style') || null, renderHTML: (a) => (a.pstyle ? { 'data-style': a.pstyle, class: `of-p-${a.pstyle}` } : {}) } } },
      { types: ['bulletList', 'orderedList'], attributes: { listStyle: { default: null, parseHTML: (el) => el.getAttribute('data-list-style') || null, renderHTML: (a) => (a.listStyle ? { 'data-list-style': a.listStyle, class: a.listStyle === 'outline' ? 'of-list-outline' : null, style: a.listStyle !== 'outline' ? `list-style-type:${a.listStyle}` : null } : {}) }, ...keepAttr('ox') } },
      { types: ['table'], attributes: keepAttr('ox') },
      { types: ['tableCell', 'tableHeader'], attributes: { backgroundColor: { default: null, parseHTML: (el) => el.getAttribute('data-background') || el.style.backgroundColor || null, renderHTML: (a) => (a.backgroundColor ? { 'data-background': a.backgroundColor, style: `background-color:${a.backgroundColor}` } : {}) } } },
    ];
  },
  addCommands() {
    const setAttr = (attrs) => ({ state, tr, dispatch }) => {
      const { from, to } = state.selection;
      state.doc.nodesBetween(from, to, (node, pos) => { if (node.type.name === 'paragraph' || node.type.name === 'heading') tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrs }); });
      if (dispatch) dispatch(tr);
      return true;
    };
    return {
      setParagraphFormat: (attrs) => setAttr(attrs),
      changeIndent: (delta) => ({ state, tr, dispatch, editor }) => {
        if (editor.isActive('listItem') || editor.isActive('taskItem')) return delta > 0 ? editor.commands.sinkListItem(editor.isActive('taskItem') ? 'taskItem' : 'listItem') : editor.commands.liftListItem(editor.isActive('taskItem') ? 'taskItem' : 'listItem');
        const { from, to } = state.selection;
        state.doc.nodesBetween(from, to, (node, pos) => { if (node.type.name === 'paragraph' || node.type.name === 'heading') tr.setNodeMarkup(pos, undefined, { ...node.attrs, indent: Math.max(0, (node.attrs.indent || 0) + delta) || null }); });
        if (dispatch) dispatch(tr);
        return true;
      },
      setParagraphStyle: (style) => ({ chain, editor }) => {
        const c = chain().focus();
        if (editor.isActive('codeBlock') && style !== 'code') c.toggleCodeBlock();
        if (editor.isActive('blockquote') && style !== 'quote') c.lift('blockquote');
        switch (style) {
          case 'normal': return c.setParagraph().setParagraphFormat({ pstyle: null, styleId: null }).run();
          case 'title': case 'subtitle': case 'listparagraph': return c.setParagraph().setParagraphFormat({ pstyle: style }).run();
          case 'quote': return (editor.isActive('blockquote') ? c : c.setParagraph().setParagraphFormat({ pstyle: null }).setBlockquote()).run();
          case 'code': return (editor.isActive('codeBlock') ? c : c.setCodeBlock()).run();
          default: { const m = /^h(\d)$/.exec(style); return m ? c.setHeading({ level: Number(m[1]) }).run() : c.run(); }
        }
      },
    };
  },
});

/** run properties from Word we do not edit (kept for saving) */
export const RunKeep = Mark.create({
  name: 'runKeep',
  addAttributes() { return { ox: { default: null, parseHTML: () => null, renderHTML: () => ({}) } }; },
  parseHTML() { return []; },
  renderHTML() { return ['span', { class: 'of-docs-rk' }, 0]; },
  excludes: '',
});

/** Word-like shortcuts, priority over tiptap's defaults */
export const WordKeys = Extension.create({
  name: 'wordKeys',
  priority: 1000,
  addOptions() { return { onLink: null, onPageBreak: null, onTrack: null, onMarks: null, sizes: [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72] }; },
  addKeyboardShortcuts() {
    const ed = () => this.editor;
    const step = (dir) => () => {
      const cur = parseFloat(ed().getAttributes('textStyle').fontSize) || this.editor.storage.wordKeys.base || 11;
      const s = this.options.sizes;
      const next = dir > 0 ? s.find((x) => x > cur) || cur + 2 : [...s].reverse().find((x) => x < cur) || Math.max(1, cur - 1);
      return ed().chain().focus().setFontSize(`${next}pt`).run();
    };
    const inStructure = () => ed().isActive('listItem') || ed().isActive('taskItem') || ed().isActive('tableCell') || ed().isActive('tableHeader') || ed().isActive('codeBlock');
    return {
      'Mod-e': () => ed().chain().focus().setTextAlign('center').run(),
      'Mod-l': () => ed().chain().focus().setTextAlign('left').run(),
      'Mod-r': () => ed().chain().focus().setTextAlign('right').run(),
      'Mod-Shift-j': () => ed().chain().focus().setTextAlign('justify').run(),
      'Mod-Shift-l': () => ed().chain().focus().toggleBulletList().run(),
      'Mod-Shift-7': () => ed().chain().focus().toggleOrderedList().run(),
      'Mod-Enter': () => { if (this.options.onPageBreak) this.options.onPageBreak(); else ed().chain().focus().insertContent({ type: 'pageBreak' }).run(); return true; },
      'Mod-=': () => ed().chain().focus().toggleSubscript().run(),
      'Mod-Shift-=': () => ed().chain().focus().toggleSuperscript().run(),
      'Mod-+': () => ed().chain().focus().toggleSuperscript().run(),
      'Alt-Shift-5': () => ed().chain().focus().toggleStrike().run(),
      'Mod-Shift-x': () => false,
      'Mod-Shift-k': () => { if (this.options.onLink) this.options.onLink(); return true; },
      'Mod-Shift-e': () => { if (this.options.onTrack) this.options.onTrack(); return true; },
      'Mod-Shift-8': () => { if (this.options.onMarks) this.options.onMarks(); return true; },
      'Mod-Shift-v': () => false,
      'Mod-]': step(1), 'Mod-[': step(-1), 'Mod-Shift->': step(1), 'Mod-Shift-<': step(-1),
      'Mod-Space': () => ed().chain().focus().unsetAllMarks().run(),
      'Mod-Alt-0': () => ed().chain().focus().setParagraphStyle('normal').run(),
      Tab: () => { if (inStructure()) return false; return ed().chain().insertContent('\t').run(); },
    };
  },
  addStorage() { return { base: 11 }; },
});
