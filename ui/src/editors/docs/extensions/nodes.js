/*
 * Custom nodes: page break, footnote, table of contents, equation, call-out/text box, preserved (kept 1:1 from the
 * source file), image with resize handles/alignment/wrap/alt text, and the comment mark.
 */
import { Node, Mark, mergeAttributes } from '@tiptap/core';
import Image from '@tiptap/extension-image';
import { t } from '@hostwerk/i18n';
import { equationHtml } from '../io/equation.js';
import { floatStyle, moveAnchor } from '../io/anchor.js';

const json = (s, d) => { try { return JSON.parse(s); } catch { return d; } };
const keep = { default: null, parseHTML: () => null, renderHTML: () => ({}) };

export const PageBreak = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,
  parseHTML() { return [{ tag: 'div[data-type="page-break"]' }, { tag: 'div', getAttrs: (el) => (/page-break-(after|before):\s*always|break-(after|before):\s*page/.test(el.getAttribute('style') || '') && !el.textContent.trim() ? {} : false) }]; },
  renderHTML() { return ['div', { 'data-type': 'page-break', class: 'of-docs-pagebreak', contenteditable: 'false' }, ['span', {}, t('Page break')]]; },
  addCommands() { return { setPageBreak: () => ({ chain }) => chain().insertContent([{ type: 'pageBreak' }, { type: 'paragraph' }]).run() }; },
});

export const Footnote = Node.create({
  name: 'footnote',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addAttributes() { return { text: { default: '', parseHTML: (el) => el.getAttribute('data-text') || '' }, ox: keep }; },
  parseHTML() { return [{ tag: 'sup[data-type="footnote"]', priority: 80 }]; },
  renderHTML({ HTMLAttributes }) { return ['sup', mergeAttributes(HTMLAttributes, { 'data-type': 'footnote', 'data-text': HTMLAttributes.text, class: 'of-docs-fn', title: HTMLAttributes.text })]; },
  addCommands() { return { insertFootnote: (text = '') => ({ chain }) => chain().insertContent({ type: 'footnote', attrs: { text } }).run() }; },
});

export const Toc = Node.create({
  name: 'toc',
  group: 'block',
  atom: true,
  selectable: true,
  addOptions() { return { onJump: null, onUpdate: null }; },
  addAttributes() {
    return {
      title: { default: '', parseHTML: (el) => el.getAttribute('data-title') || '' },
      entries: { default: [], parseHTML: (el) => json(el.getAttribute('data-entries'), []) },
    };
  },
  parseHTML() { return [{ tag: 'nav[data-type="toc"]', priority: 80 }]; },
  renderHTML({ HTMLAttributes }) { return ['nav', { 'data-type': 'toc', 'data-title': HTMLAttributes.title, 'data-entries': JSON.stringify(HTMLAttributes.entries || []), class: 'of-docs-toc' }]; },
  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement('nav');
      dom.className = 'of-docs-toc';
      dom.contentEditable = 'false';
      const draw = (n) => {
        dom.innerHTML = '';
        const head = document.createElement('div');
        head.className = 'of-toc-head';
        const title = document.createElement('p'); title.className = 'of-toc-title'; title.textContent = n.attrs.title || '';
        const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'btn xs ghost of-toc-update'; btn.textContent = t('Update table'); btn.title = t('Update table');
        btn.addEventListener('mousedown', (e) => { e.preventDefault(); if (this.options.onUpdate) this.options.onUpdate(); });
        head.append(title, btn);
        dom.append(head);
        const es = n.attrs.entries || [];
        if (!es.length) { const p = document.createElement('p'); p.className = 'of-toc-empty'; p.textContent = t('No headings yet – use Heading 1–3 and update the table.'); dom.append(p); }
        es.forEach((e, k) => {
          const p = document.createElement('p'); p.className = `of-toc-l${Math.min(6, e.level || 1)}`;
          const a = document.createElement('a'); a.href = '#'; a.textContent = e.text;
          a.addEventListener('mousedown', (ev) => { ev.preventDefault(); if (this.options.onJump) this.options.onJump(e.index != null ? e.index : k, e.text); });
          p.append(a); dom.append(p);
        });
      };
      draw(node);
      return { dom, update: (n) => { if (n.type.name !== 'toc') return false; draw(n); return true; }, ignoreMutation: () => true, stopEvent: (e) => e.target && e.target.closest && !!e.target.closest('button, a') };
    };
  },
});

export const Equation = Node.create({
  name: 'equation',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addOptions() { return { onEdit: null }; },
  addAttributes() { return { latex: { default: '', parseHTML: (el) => el.getAttribute('data-latex') || el.textContent || '' }, omml: keep, ommlLatex: keep }; },
  parseHTML() { return [{ tag: 'span[data-type="equation"]', priority: 80 }]; },
  renderHTML({ HTMLAttributes }) { return ['span', { 'data-type': 'equation', 'data-latex': HTMLAttributes.latex, class: 'of-docs-eq' }]; },
  addNodeView() {
    return ({ node, getPos }) => {
      const dom = document.createElement('span');
      dom.className = 'of-docs-eq';
      dom.contentEditable = 'false';
      dom.title = t('Double-click to edit the equation');
      const draw = (n) => { dom.innerHTML = equationHtml(n.attrs.latex) || '<i>x</i>'; };
      draw(node);
      dom.addEventListener('dblclick', () => { if (this.options.onEdit) this.options.onEdit(typeof getPos === 'function' ? getPos() : null, node.attrs.latex); });
      return { dom, update: (n) => { if (n.type.name !== 'equation') return false; node = n; draw(n); return true; } };
    };
  },
});

export const Callout = Node.create({
  name: 'callout',
  group: 'block',
  content: 'block+',
  defining: true,
  addAttributes() { return { kind: { default: 'info', parseHTML: (el) => el.getAttribute('data-kind') || 'info' } }; },
  parseHTML() { return [{ tag: 'div[data-type="callout"]' }]; },
  renderHTML({ HTMLAttributes }) { return ['div', { 'data-type': 'callout', 'data-kind': HTMLAttributes.kind, class: `of-docs-callout of-callout-${HTMLAttributes.kind}` }, 0]; },
  addCommands() {
    return {
      insertCallout: (kind = 'info') => ({ chain, state }) => {
        const { from, to } = state.selection;
        const text = state.doc.textBetween(from, to, ' ');
        return chain().insertContent({ type: 'callout', attrs: { kind }, content: [{ type: 'paragraph', content: text ? [{ type: 'text', text }] : [] }] }).run();
      },
    };
  },
});

/**
 * A Word text box / shape anchored in the text (wp:anchor): floats at its position – wrap square = the text flows
 * around it, top-and-bottom = on its own line, none = over the text. Drag moves it (written back as wp:posOffset from
 * the column / paragraph – the rest of the drawing stays byte-identical); double-click edits its text.
 */
function floatingView(node, editor, getPos) {
  let a = node.attrs;
  const dom = document.createElement('span');
  dom.contentEditable = 'false';
  const box = document.createElement('span');
  box.className = 'of-kept-floatbox';
  const paint = () => {
    const f = a.float;
    const secs = (editor.storage.ofPages && editor.storage.ofPages.sections) || [];
    const g = secs[secs.length - 1];
    const contentW = editor.view.dom.offsetWidth || (g && g.contentW) || 600;
    const st = floatStyle(f, { contentW, marginLeft: g ? g.ml : 94 });
    dom.className = `of-docs-kept of-kept-${a.kind} of-kept-float is-${st.mode}${f.behind ? ' is-behind' : ''}`;
    dom.title = t('{name} – drag to move it, double-click to edit its text', { name: t(a.label) });
    box.replaceChildren();
    if (a.src) { const img = document.createElement('img'); img.src = a.src; img.alt = a.label; img.draggable = false; box.append(img); } else { const tx = document.createElement('span'); tx.className = 'of-kept-box'; tx.textContent = a.text || ''; box.append(tx); }
    box.style.width = f.w ? `${f.w}px` : ''; box.style.minHeight = f.h ? `${f.h}px` : '';
    dom.style.cssText = ''; box.style.left = ''; box.style.top = '';
    if (st.mode === 'float-left') Object.assign(dom.style, { float: 'left', marginLeft: `${st.marginLeft}px`, marginTop: `${st.top}px` });
    else if (st.mode === 'float-right') Object.assign(dom.style, { float: 'right', marginRight: `${st.marginRight}px`, marginTop: `${st.top}px` });
    else if (st.mode === 'block') Object.assign(dom.style, { display: 'block', marginLeft: `${st.marginLeft}px`, marginTop: `${st.top}px`, width: f.w ? `${f.w}px` : '' });
    else {
      // over the text: a zero-size anchor at the node, the box placed from the start of its paragraph
      Object.assign(box.style, { top: `${st.top}px`, left: `${st.left}px` });
      requestAnimationFrame(() => {
        const p = dom.closest('p, h1, h2, h3, h4, h5, h6, li, td, th') || dom.parentElement;
        if (!p || !dom.isConnected) return;
        const k = editor.view.dom.offsetWidth ? editor.view.dom.getBoundingClientRect().width / editor.view.dom.offsetWidth : 1;
        const dx = (dom.getBoundingClientRect().left - p.getBoundingClientRect().left) / (k || 1);
        const dy = (dom.getBoundingClientRect().top - p.getBoundingClientRect().top) / (k || 1);
        box.style.left = `${Math.round(st.left - dx)}px`; box.style.top = `${Math.round(st.top - dy)}px`;
      });
    }
    return st;
  };
  dom.append(box);
  const tag = document.createElement('span'); tag.className = 'of-kept-tag'; tag.textContent = t(a.label); box.append(tag);
  let st = paint(); box.append(tag);
  // drag to move
  let drag = null;
  const down = (e) => {
    if (e.button !== 0 || !editor.isEditable || e.detail > 1) return;
    e.preventDefault();
    const k = editor.view.dom.offsetWidth ? editor.view.dom.getBoundingClientRect().width / editor.view.dom.offsetWidth : 1;
    drag = { x0: e.clientX, y0: e.clientY, k: k || 1, moved: false };
    try { dom.setPointerCapture(e.pointerId); } catch { /* old browser */ }
  };
  const move = (e) => {
    if (!drag) return;
    const dx = (e.clientX - drag.x0) / drag.k; const dy = (e.clientY - drag.y0) / drag.k;
    if (!drag.moved && Math.hypot(dx, dy) < 3) return;
    drag.moved = true; dom.classList.add('is-dragging');
    box.style.transform = `translate(${dx}px, ${dy}px)`;
  };
  const up = (e) => {
    if (!drag) return;
    const d = drag; drag = null;
    dom.classList.remove('is-dragging'); box.style.transform = '';
    const pos = typeof getPos === 'function' ? getPos() : null;
    if (pos == null) return;
    if (!d.moved) { editor.chain().setNodeSelection(pos).run(); return; }
    const dx = Math.round((e.clientX - d.x0) / d.k); const dy = Math.round((e.clientY - d.y0) / d.k);
    // the shown position (column / paragraph) + the move → the new offsets
    const baseX = st.mode === 'float-right' ? (editor.view.dom.offsetWidth || 600) - (st.marginRight || 0) - (a.float.w || 0) : st.mode === 'absolute' ? st.left : st.marginLeft || 0;
    const nx = baseX + dx; const ny = (st.top || 0) + dy;
    const float = { ...a.float, x: nx, y: ny, relH: 'column', relV: 'paragraph', alignH: null, alignV: null };
    editor.view.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, { ...a, float, xml: a.xml ? moveAnchor(a.xml, nx, ny) : a.xml }));
  };
  dom.addEventListener('pointerdown', down);
  dom.addEventListener('pointermove', move);
  dom.addEventListener('pointerup', up);
  dom.addEventListener('dblclick', (e) => { e.preventDefault(); const pos = typeof getPos === 'function' ? getPos() : null; const fn = editor.storage.preserved && editor.storage.preserved.onEdit; if (pos != null && fn && editor.isEditable) fn(pos, editor.state.doc.nodeAt(pos)); });
  dom.addEventListener('pointercancel', () => { drag = null; dom.classList.remove('is-dragging'); box.style.transform = ''; });
  return {
    dom,
    update: (n) => { if (n.type.name !== 'preserved' || !n.attrs.float) return false; a = n.attrs; st = paint(); box.append(tag); return true; },
    selectNode: () => dom.classList.add('of-sel'), deselectNode: () => dom.classList.remove('of-sel'),
    stopEvent: (e) => /^(pointer|mouse|click|dblclick|dragstart)/.test(e.type),
    ignoreMutation: () => true,
  };
}

/** kept 1:1 from the source file (SmartArt, charts, shapes, OLE objects, fields …): shown with its fallback */
export const Preserved = Node.create({
  name: 'preserved',
  addStorage() { return { onEdit: null }; },
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: false,
  addAttributes() {
    return { format: { default: null }, kind: { default: 'object' }, label: { default: 'Object' }, text: { default: '' }, src: { default: null }, width: { default: null }, xml: keep, float: keep };
  },
  parseHTML() { return [{ tag: 'span[data-type="preserved"]', getAttrs: (el) => ({ kind: el.getAttribute('data-kind') || 'object', label: el.getAttribute('data-label') || 'Object', text: el.textContent || '' }) }]; },
  renderHTML({ HTMLAttributes: a }) { return ['span', { 'data-type': 'preserved', 'data-kind': a.kind, 'data-label': a.label, class: 'of-docs-kept' }, a.text || a.label || '']; },
  addNodeView() {
    return ({ node, editor, getPos }) => {
      if (node.attrs.float && node.attrs.format === 'docx' && (node.attrs.kind === 'textbox' || node.attrs.kind === 'shape')) return floatingView(node, editor, getPos);
      const dom = document.createElement('span');
      dom.contentEditable = 'false';
      const a = node.attrs;
      dom.className = `of-docs-kept of-kept-${a.kind}${a.kind === 'field' ? ' of-kept-inline' : ''}`;
      dom.title = a.kind === 'field' ? t('Field (kept from the original file): {name}', { name: a.label }) : t('{name} – kept as it is; edit it in the desktop app or the full editor', { name: t(a.label) });
      if (a.kind === 'field') { dom.textContent = a.text || a.label; return { dom }; }
      if (a.src) { const img = document.createElement('img'); img.src = a.src; img.alt = a.label; if (a.width) img.style.width = `${Math.min(a.width, 640)}px`; dom.append(img); }
      else { const box = document.createElement('span'); box.className = 'of-kept-box'; box.textContent = a.text || ''; if (a.width) box.style.width = `${Math.min(a.width, 640)}px`; dom.append(box); }
      const tag = document.createElement('span'); tag.className = 'of-kept-tag'; tag.textContent = t(a.label);
      dom.append(tag);
      return { dom };
    };
  },
});

/** image: inline, with width, alignment (left/center/right), wrap (square = text flows around) and alt text */
export const ImageEx = Image.extend({
  inline: true,
  group: 'inline',
  draggable: true,
  addAttributes() {
    return {
      ...this.parent?.(),
      width: { default: null, parseHTML: (el) => { const w = el.getAttribute('width') || (el.style.width || '').replace('px', ''); return w && !/%/.test(w) ? Number(w) || null : null; } },
      height: { default: null, parseHTML: () => null, renderHTML: () => ({}) },
      align: { default: null, parseHTML: (el) => el.getAttribute('data-align') || null },
      wrap: { default: null, parseHTML: (el) => el.getAttribute('data-wrap') || (/float:\s*(left|right)/.test(el.getAttribute('style') || '') ? 'square' : null) },
      ox: keep,
    };
  },
  renderHTML({ HTMLAttributes: a }) {
    const style = [a.width ? `width:${a.width}px` : '', 'max-width:100%', 'height:auto', a.wrap === 'square' && (a.align === 'left' || a.align === 'right') ? `float:${a.align};margin:${a.align === 'left' ? '0 12px 6px 0' : '0 0 6px 12px'}` : '', a.align === 'center' && a.wrap !== 'square' ? 'display:block;margin:0 auto' : ''].filter(Boolean).join(';');
    return ['img', { src: a.src, alt: a.alt || '', title: a.title || null, width: a.width || null, 'data-align': a.align || null, 'data-wrap': a.wrap || null, style }];
  },
  addNodeView() {
    return ({ node, getPos, editor }) => {
      const wrap = document.createElement('span');
      const img = document.createElement('img');
      wrap.append(img);
      const handles = ['nw', 'ne', 'sw', 'se'].map((c) => { const h = document.createElement('span'); h.className = `of-img-h of-img-${c}`; h.dataset.corner = c; wrap.append(h); return h; });
      let cur = node;
      const draw = (n) => {
        const a = n.attrs;
        img.src = a.src; img.alt = a.alt || ''; img.title = a.alt || '';
        img.style.width = a.width ? `${a.width}px` : '';
        wrap.className = `of-docs-img${a.align ? ` of-img-${a.align}` : ''}${a.wrap === 'square' ? ' of-img-wrap' : ''}`;
      };
      draw(node);
      const start = (e) => {
        if (!editor.isEditable) return;
        e.preventDefault(); e.stopPropagation();
        const corner = e.currentTarget.dataset.corner;
        const x0 = e.clientX; const w0 = img.getBoundingClientRect().width;
        const scale = w0 / (img.offsetWidth || w0) || 1;
        const move = (ev) => { const dx = (ev.clientX - x0) * (corner.includes('w') ? -1 : 1); img.style.width = `${Math.max(24, Math.round((w0 + dx) / scale))}px`; };
        const up = () => {
          window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
          const w = Math.round(parseFloat(img.style.width));
          if (typeof getPos === 'function') editor.chain().command(({ tr }) => { tr.setNodeMarkup(getPos(), undefined, { ...cur.attrs, width: w, height: null }); return true; }).run();
        };
        window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
      };
      for (const h of handles) h.addEventListener('pointerdown', start);
      return {
        dom: wrap,
        update: (n) => { if (n.type.name !== 'image') return false; cur = n; draw(n); return true; },
        selectNode: () => wrap.classList.add('of-img-selected'),
        deselectNode: () => wrap.classList.remove('of-img-selected'),
        stopEvent: (e) => !!(e.target && e.target.classList && e.target.classList.contains('of-img-h')),
      };
    };
  },
});

/** comment anchors (the texts live in the comments panel) */
export const Comment = Mark.create({
  name: 'comment',
  inclusive: false,
  excludes: '',
  addAttributes() { return { id: { default: null, parseHTML: (el) => el.getAttribute('data-comment-id') } }; },
  parseHTML() { return [{ tag: 'span[data-comment-id]' }]; },
  renderHTML({ HTMLAttributes }) { return ['span', { 'data-comment-id': HTMLAttributes.id, class: 'of-docs-comment' }, 0]; },
  addCommands() {
    return {
      setComment: (id) => ({ commands }) => commands.setMark('comment', { id }),
      removeComment: (id) => ({ state, tr, dispatch }) => {
        const type = state.schema.marks.comment;
        state.doc.descendants((node, pos) => { for (const m of node.marks) if (m.type === type && m.attrs.id === id) tr.removeMark(pos, pos + node.nodeSize, m); });
        if (dispatch) dispatch(tr);
        return true;
      },
    };
  },
});
