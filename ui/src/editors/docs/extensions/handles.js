/*
 * Block handles: hovering a paragraph, heading, list, table or picture shows a grip on the left. Drag it to move the
 * whole block (ProseMirror moves the node – undo works, track changes records it like a cut/paste); the "+" next to
 * it adds an empty paragraph below and opens the "/" menu. Keyboard users have Alt+Shift+↑/↓ (move block).
 */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, NodeSelection, TextSelection } from '@tiptap/pm/state';
import { t } from '@hostwerk/i18n';
import { slashKey } from './slash.js';

export const handlesKey = new PluginKey('ofHandles');

/** top-level block (index, pos, node) at a DOM element */
function blockAt(view, el) {
  let cur = el;
  while (cur && cur.parentElement !== view.dom) cur = cur.parentElement;
  if (!cur || cur.classList.contains('of-docs-pgsp')) return null;
  let found = null;
  view.state.doc.forEach((node, offset, index) => { if (!found && view.nodeDOM(offset) === cur) found = { node, pos: offset, index, dom: cur }; });
  return found;
}

export const BlockHandles = Extension.create({
  name: 'ofHandles',
  addOptions() { return { enabled: true }; },
  addKeyboardShortcuts() {
    const move = (dir) => () => {
      const { state, view } = this.editor;
      const $p = state.selection.$from;
      if ($p.depth < 1) return false;
      const pos = $p.before(1); const node = state.doc.nodeAt(pos);
      const index = $p.index(0);
      const tr = state.tr;
      if (dir < 0) {
        if (index === 0) return true;
        const prev = state.doc.child(index - 1);
        tr.delete(pos, pos + node.nodeSize).insert(pos - prev.nodeSize, node);
        tr.setSelection(TextSelection.near(tr.doc.resolve(pos - prev.nodeSize + 1)));
      } else {
        if (index >= state.doc.childCount - 1) return true;
        const next = state.doc.child(index + 1);
        tr.delete(pos, pos + node.nodeSize).insert(pos + next.nodeSize, node);
        tr.setSelection(TextSelection.near(tr.doc.resolve(pos + next.nodeSize + 1)));
      }
      view.dispatch(tr.scrollIntoView());
      return true;
    };
    return { 'Alt-Shift-ArrowUp': move(-1), 'Alt-Shift-ArrowDown': move(1) };
  },
  addProseMirrorPlugins() {
    const ext = this;
    return [new Plugin({
      key: handlesKey,
      view(view) {
        const box = document.createElement('div');
        box.className = 'of-docs-handle';
        box.contentEditable = 'false';
        const plus = document.createElement('button'); plus.type = 'button'; plus.className = 'of-hd-plus'; plus.textContent = '+'; plus.title = t('Insert below (/ menu)'); plus.setAttribute('aria-label', plus.title);
        const grip = document.createElement('button'); grip.type = 'button'; grip.className = 'of-hd-grip'; grip.draggable = true; grip.title = t('Drag to move · click to select (Alt+Shift+↑/↓ moves too)'); grip.setAttribute('aria-label', t('Move block'));
        grip.innerHTML = '<svg width="10" height="14" viewBox="0 0 10 14" aria-hidden="true"><g fill="currentColor"><circle cx="2.5" cy="2.5" r="1.4"/><circle cx="7.5" cy="2.5" r="1.4"/><circle cx="2.5" cy="7" r="1.4"/><circle cx="7.5" cy="7" r="1.4"/><circle cx="2.5" cy="11.5" r="1.4"/><circle cx="7.5" cy="11.5" r="1.4"/></g></svg>';
        box.append(plus, grip);
        let cur = null; let hideTimer = null;
        const host = () => view.dom.parentElement;
        const place = (b) => {
          if (!view.editable || !ext.options.enabled) return;
          cur = b;
          const h = host(); if (!h) return;
          if (box.parentElement !== h) h.append(box);
          const top = b.dom.offsetTop + (view.dom.offsetTop || 0);
          const lh = Math.min(28, Math.max(18, b.dom.firstElementChild ? 22 : b.dom.offsetHeight));
          box.style.top = `${Math.round(top + Math.max(0, (Math.min(b.dom.offsetHeight, 28) - lh) / 2))}px`;
          box.style.left = `${Math.round((b.dom.offsetLeft || 0) - 52)}px`;
          box.classList.add('of-on');
        };
        const hide = () => { box.classList.remove('of-on'); cur = null; };
        const move = (e) => {
          if (view.dragging) return;
          if (e.buttons) return;
          const b = blockAt(view, e.target);
          clearTimeout(hideTimer);
          if (b) place(b); else hideTimer = setTimeout(hide, 400);
        };
        const leave = (e) => { if (!box.contains(e.relatedTarget)) hideTimer = setTimeout(hide, 500); };
        view.dom.addEventListener('mousemove', move);
        view.dom.addEventListener('mouseleave', leave);
        box.addEventListener('mouseenter', () => clearTimeout(hideTimer));
        box.addEventListener('mouseleave', leave);
        grip.addEventListener('mousedown', (e) => { e.stopPropagation(); });
        grip.addEventListener('click', () => { if (!cur) return; view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, cur.pos))); view.focus(); });
        grip.addEventListener('dragstart', (e) => {
          if (!cur) return;
          const sel = NodeSelection.create(view.state.doc, cur.pos);
          view.dispatch(view.state.tr.setSelection(sel));
          const slice = sel.content();
          view.dragging = { slice, move: true };
          if (e.dataTransfer) {
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', cur.node.textContent.slice(0, 200));
            try { e.dataTransfer.setDragImage(cur.dom, 0, 0); } catch { /* old browser */ }
          }
          view.dom.classList.add('of-dragging');
        });
        grip.addEventListener('dragend', () => { view.dom.classList.remove('of-dragging'); hide(); });
        plus.addEventListener('mousedown', (e) => e.preventDefault());
        plus.addEventListener('click', () => {
          if (!cur) return;
          const end = cur.pos + cur.node.nodeSize;
          const { schema } = view.state;
          const empty = cur.node.type.name === 'paragraph' && !cur.node.content.size;
          let tr = view.state.tr;
          let at;
          if (empty) { at = cur.pos + 1; tr = tr.insertText('/', at); } else { tr = tr.insert(end, schema.nodes.paragraph.create(null, schema.text('/'))); at = end + 1; }
          tr.setSelection(TextSelection.create(tr.doc, at + 1)).setMeta(slashKey, { open: true, from: at });
          view.dispatch(tr.scrollIntoView());
          view.focus();
          hide();
        });
        return { update() { if (cur && !cur.dom.isConnected) hide(); }, destroy() { view.dom.removeEventListener('mousemove', move); view.dom.removeEventListener('mouseleave', leave); box.remove(); } };
      },
    })];
  },
});
