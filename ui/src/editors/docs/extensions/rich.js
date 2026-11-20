/*
 * 1.1 structures: content controls (inline check box / drop-down / date, the `sdt` mark for text controls, block
 * controls), live fields, section breaks with their own page setup, diagrams from text (Mermaid), formulas rendered
 * as real MathML (temml, lazy) with `$…$` input, and the tracked-change attributes of paragraphs and table rows.
 */
import { Node, Mark, Extension, InputRule, mergeAttributes } from '@tiptap/core';
import { t } from '@hostwerk/i18n';
import { equationHtml } from '../io/equation.js';
import { renderLatex } from '../lazy.js';

const keep = { default: null, parseHTML: () => null, renderHTML: () => ({}) };
const host = () => (typeof document !== 'undefined' ? document.querySelector('.app') || document.body : null);

/** a small menu under an element (plain DOM, used inside node views) */
function popover(anchor, build, onClose) {
  const box = document.createElement('div');
  box.className = 'of-docs-pop of-docs-ccpop';
  box.setAttribute('role', 'menu');
  const r = anchor.getBoundingClientRect();
  box.style.left = `${Math.max(8, Math.min(r.left, window.innerWidth - 260))}px`;
  box.style.top = `${r.bottom + 4}px`;
  const close = () => { box.remove(); window.removeEventListener('mousedown', out, true); window.removeEventListener('keydown', key, true); if (onClose) onClose(); };
  const out = (e) => { if (!box.contains(e.target)) close(); };
  const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(); } };
  build(box, close);
  (host() || document.body).append(box);
  setTimeout(() => { window.addEventListener('mousedown', out, true); window.addEventListener('keydown', key, true); }, 0);
  const first = box.querySelector('button, input');
  if (first) setTimeout(() => first.focus(), 0);
  return close;
}

/** dd.MM.yyyy-style formats of Word's date controls */
export function formatDate(iso, fmt = 'dd.MM.yyyy') {
  if (!iso) return '';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  const p = (n, l = 2) => String(n).padStart(l, '0');
  const lang = (typeof document !== 'undefined' && document.documentElement.lang) || 'de';
  const mon = (style) => d.toLocaleDateString(lang, { month: style, timeZone: 'UTC' });
  return String(fmt || 'dd.MM.yyyy').replace(/yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|mm/g, (k) => ({
    yyyy: d.getUTCFullYear(), yy: p(d.getUTCFullYear() % 100), MMMM: mon('long'), MMM: mon('short'), MM: p(d.getUTCMonth() + 1), M: d.getUTCMonth() + 1,
    dddd: d.toLocaleDateString(lang, { weekday: 'long', timeZone: 'UTC' }), ddd: d.toLocaleDateString(lang, { weekday: 'short', timeZone: 'UTC' }), dd: p(d.getUTCDate()), d: d.getUTCDate(),
    HH: p(d.getUTCHours()), H: d.getUTCHours(), mm: p(d.getUTCMinutes()),
  }[k]));
}

/* ---------------- content controls ---------------- */
export const ContentControl = Node.create({
  name: 'contentControl',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addOptions() { return { onEdit: null }; },
  addAttributes() {
    return {
      kind: { default: 'text', parseHTML: (el) => el.getAttribute('data-kind') || 'text' },
      value: { default: '', parseHTML: (el) => el.textContent || '' },
      checked: { default: false, parseHTML: (el) => el.getAttribute('data-checked') === 'true' },
      items: { default: null, parseHTML: (el) => { try { return JSON.parse(el.getAttribute('data-items') || 'null'); } catch { return null; } } },
      format: { default: null }, fullDate: { default: null }, tag: { default: null }, alias: { default: null }, placeholder: { default: false },
      pr: keep, end: keep,
    };
  },
  parseHTML() { return [{ tag: 'span[data-type="content-control"]', priority: 80 }]; },
  renderHTML({ HTMLAttributes: a }) { return ['span', { 'data-type': 'content-control', 'data-kind': a.kind, 'data-checked': a.checked ? 'true' : 'false', 'data-items': a.items ? JSON.stringify(a.items) : null, class: 'of-docs-cc' }, a.kind === 'checkbox' ? (a.checked ? '☒' : '☐') : a.value || '']; },
  addNodeView() {
    return ({ node, getPos, editor }) => {
      const dom = document.createElement('span');
      dom.contentEditable = 'false';
      let cur = node;
      const set = (patch) => { if (typeof getPos !== 'function' || !editor.isEditable) return; const pos = getPos(); editor.chain().command(({ tr }) => { tr.setNodeMarkup(pos, undefined, { ...cur.attrs, ...patch }); return true; }).run(); };
      const draw = (n) => {
        const a = n.attrs;
        dom.className = `of-docs-cc of-cc-${a.kind}${!a.value && a.kind !== 'checkbox' ? ' of-cc-empty' : ''}`;
        dom.title = `${a.alias ? `${a.alias} · ` : ''}${t({ checkbox: 'Check box', dropdown: 'Drop-down list', combo: 'Combo box', date: 'Date picker', text: 'Text field' }[a.kind] || 'Content control')}`;
        dom.textContent = '';
        if (a.kind === 'checkbox') { dom.textContent = a.checked ? '☒' : '☐'; dom.setAttribute('role', 'checkbox'); dom.setAttribute('aria-checked', a.checked ? 'true' : 'false'); return; }
        const v = document.createElement('span'); v.className = 'of-cc-v';
        v.textContent = a.value || (a.kind === 'date' ? t('Pick a date') : a.kind === 'text' ? t('Click to enter text') : t('Choose an item'));
        dom.append(v);
        if (a.kind !== 'text') { const c = document.createElement('span'); c.className = 'of-cc-caret'; c.textContent = a.kind === 'date' ? '📅' : '▾'; c.setAttribute('aria-hidden', 'true'); dom.append(c); }
      };
      draw(node);
      const open = () => {
        const a = cur.attrs;
        if (!editor.isEditable) return;
        if (a.kind === 'checkbox') { set({ checked: !a.checked }); return; }
        if (a.kind === 'dropdown' || a.kind === 'combo') {
          popover(dom, (box, close) => {
            for (const it of a.items || []) {
              const b = document.createElement('button'); b.type = 'button'; b.className = `of-docs-mi${it.text === a.value ? ' of-on' : ''}`; b.textContent = it.text;
              b.addEventListener('click', () => { set({ value: it.text, placeholder: false }); close(); });
              box.append(b);
            }
            if (a.kind === 'combo') {
              const inp = document.createElement('input'); inp.className = 'input sm'; inp.value = a.value || ''; inp.placeholder = t('Own value');
              inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { set({ value: inp.value, placeholder: false }); close(); } });
              box.append(inp);
            }
          });
          return;
        }
        if (a.kind === 'date') {
          popover(dom, (box, close) => {
            const inp = document.createElement('input'); inp.type = 'date'; inp.className = 'input sm';
            if (a.fullDate) inp.value = String(a.fullDate).slice(0, 10);
            inp.addEventListener('change', () => { if (!inp.value) return; const iso = `${inp.value}T00:00:00Z`; set({ fullDate: iso, value: formatDate(iso, a.format || 'dd.MM.yyyy'), placeholder: false }); close(); });
            const today = document.createElement('button'); today.type = 'button'; today.className = 'of-docs-mi'; today.textContent = t('Today');
            today.addEventListener('click', () => { const iso = `${new Date().toISOString().slice(0, 10)}T00:00:00Z`; set({ fullDate: iso, value: formatDate(iso, a.format || 'dd.MM.yyyy'), placeholder: false }); close(); });
            box.append(inp, today);
          });
          return;
        }
        // text control without text yet: ask for the text (afterwards it is normal text inside the control)
        if (this.options.onEdit) this.options.onEdit(typeof getPos === 'function' ? getPos() : null, cur);
      };
      dom.addEventListener('mousedown', (e) => { e.preventDefault(); });
      dom.addEventListener('click', (e) => { e.preventDefault(); open(); });
      dom.addEventListener('keydown', (e) => { if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); open(); } });
      return { dom, update: (n) => { if (n.type.name !== 'contentControl') return false; cur = n; draw(n); return true; }, stopEvent: (e) => e.type === 'mousedown' || e.type === 'click', ignoreMutation: () => true };
    };
  },
});

/** plain/rich text content controls: the text stays editable text, the mark carries the control (written as w:sdt) */
export const SdtMark = Mark.create({
  name: 'sdt',
  inclusive: false,
  excludes: '',
  addAttributes() { return { id: { default: null }, kind: { default: 'rich' }, alias: { default: null }, tag: { default: null }, pr: keep, end: keep }; },
  parseHTML() { return [{ tag: 'span[data-sdt]' }]; },
  renderHTML({ HTMLAttributes: a }) { return ['span', { 'data-sdt': a.id || '', class: 'of-docs-sdt', title: a.alias || null }, 0]; },
});

/** block content control around paragraphs/tables */
export const ControlBlock = Node.create({
  name: 'controlBlock',
  group: 'block',
  content: 'block+',
  defining: true,
  addAttributes() { return { kind: { default: 'rich' }, alias: { default: null }, tag: { default: null }, pr: keep, end: keep }; },
  parseHTML() { return [{ tag: 'div[data-type="control-block"]' }]; },
  renderHTML({ HTMLAttributes: a }) { return ['div', { 'data-type': 'control-block', class: 'of-docs-ccblock', 'data-alias': a.alias || t('Content control') }, 0]; },
});

/* ---------------- live fields ---------------- */
export const FIELD_LABELS = { page: 'Page number', pages: 'Number of pages', date: 'Date', time: 'Time', author: 'Author', title: 'Title', filename: 'File name', subject: 'Subject', words: 'Number of words', hostwerk: 'Hostwerk value', cell: 'Cell of a spreadsheet' };
export const Field = Node.create({
  name: 'field',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addOptions() { return { onEdit: null }; },
  addAttributes() {
    return {
      kind: { default: 'date', parseHTML: (el) => el.getAttribute('data-kind') || 'date' },
      value: { default: '', parseHTML: (el) => el.textContent || '' },
      instr: { default: '' }, source: { default: null }, params: { default: null }, label: { default: '' }, updatedAt: { default: null },
      ox: keep,
    };
  },
  parseHTML() { return [{ tag: 'span[data-type="field"]', priority: 80 }]; },
  renderHTML({ HTMLAttributes: a }) { return ['span', { 'data-type': 'field', 'data-kind': a.kind, class: 'of-docs-field' }, a.value || '']; },
  addNodeView() {
    return ({ node, getPos }) => {
      const dom = document.createElement('span');
      dom.contentEditable = 'false';
      const draw = (n) => {
        const a = n.attrs;
        dom.className = `of-docs-field of-fld-${a.kind}${a.kind === 'hostwerk' || a.kind === 'cell' ? ' of-fld-live' : ''}`;
        dom.textContent = a.value || (a.kind === 'page' ? '#' : '…');
        const when = a.updatedAt ? ` · ${t('updated {time}', { time: new Date(a.updatedAt).toLocaleString() })}` : '';
        dom.title = `${t(FIELD_LABELS[a.kind] || 'Field')}${a.label ? `: ${a.label}` : ''}${when} – ${t('double-click to change')}`;
      };
      draw(node);
      dom.addEventListener('dblclick', (e) => { e.preventDefault(); if (this.options.onEdit) this.options.onEdit(typeof getPos === 'function' ? getPos() : null, node); });
      return { dom, update: (n) => { if (n.type.name !== 'field') return false; node = n; draw(n); return true; }, ignoreMutation: () => true };
    };
  },
});

/* ---------------- section breaks ---------------- */
export const SectionBreak = Node.create({
  name: 'sectionBreak',
  group: 'sectionBlock',
  atom: true,
  selectable: true,
  addOptions() { return { onEdit: null }; },
  addAttributes() { return { type: { default: 'nextPage' }, page: { default: null }, titlePage: { default: false }, hf: keep, sectPr: keep }; },
  parseHTML() { return [{ tag: 'div[data-type="section-break"]' }]; },
  renderHTML() { return ['div', { 'data-type': 'section-break', class: 'of-docs-secbreak' }]; },
  addNodeView() {
    return ({ node, getPos }) => {
      const dom = document.createElement('div');
      dom.contentEditable = 'false';
      dom.className = 'of-docs-secbreak';
      const draw = (n) => {
        const a = n.attrs; const p = a.page || {};
        dom.innerHTML = '';
        const l = document.createElement('span');
        l.textContent = `${a.type === 'continuous' ? t('Section break (continuous)') : t('Section break (next page)')}${p.size ? ` · ${p.size === 'custom' ? '' : p.size} ${p.orientation === 'landscape' ? t('Landscape') : t('Portrait')}` : ''}`;
        const b = document.createElement('button'); b.type = 'button'; b.className = 'btn xs ghost'; b.textContent = t('Page setup of this section …');
        b.addEventListener('mousedown', (e) => { e.preventDefault(); if (this.options.onEdit) this.options.onEdit(typeof getPos === 'function' ? getPos() : null, node); });
        dom.append(l, b);
      };
      draw(node);
      return { dom, update: (n) => { if (n.type.name !== 'sectionBreak') return false; node = n; draw(n); return true; }, stopEvent: (e) => !!(e.target && e.target.closest && e.target.closest('button')), ignoreMutation: () => true };
    };
  },
});

/* ---------------- diagrams from text ---------------- */
export const Diagram = Node.create({
  name: 'diagram',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: true,
  addOptions() { return { onEdit: null }; },
  addAttributes() { return { kind: { default: 'mermaid' }, source: { default: '' }, src: { default: null }, svg: keep, width: { default: null }, height: { default: null } }; },
  parseHTML() { return [{ tag: 'img[data-type="diagram"]', getAttrs: (el) => ({ src: el.getAttribute('src'), source: el.getAttribute('alt') || '' }) }]; },
  renderHTML({ HTMLAttributes: a }) { return ['img', { 'data-type': 'diagram', src: a.src, alt: a.source }]; },
  addNodeView() {
    return ({ node, getPos }) => {
      const dom = document.createElement('span');
      dom.contentEditable = 'false';
      dom.className = 'of-docs-diagram';
      const draw = (n) => {
        const a = n.attrs;
        dom.innerHTML = '';
        dom.title = t('Diagram – double-click to edit its text');
        if (a.src) { const img = document.createElement('img'); img.src = a.src; img.alt = a.source || ''; img.loading = 'lazy'; img.decoding = 'async'; if (a.width) img.style.width = `${Math.min(a.width, 900)}px`; dom.append(img); }
        else { const pre = document.createElement('pre'); pre.className = 'of-docs-diagram-src'; pre.textContent = a.source || t('Empty diagram'); dom.append(pre); }
        const tag = document.createElement('span'); tag.className = 'of-kept-tag'; tag.textContent = t('Diagram'); dom.append(tag);
      };
      draw(node);
      dom.addEventListener('dblclick', (e) => { e.preventDefault(); if (this.options.onEdit) this.options.onEdit(typeof getPos === 'function' ? getPos() : null, node); });
      return { dom, update: (n) => { if (n.type.name !== 'diagram') return false; node = n; draw(n); return true }, ignoreMutation: () => true, selectNode: () => dom.classList.add('of-sel'), deselectNode: () => dom.classList.remove('of-sel') };
    };
  },
});

/* ---------------- formulas (real MathML, lazy) ---------------- */
const mathCache = new Map();
export function cachedMathml(latex, display) { return mathCache.get(`${display ? 'D' : 'I'}${latex}`); }
async function mathmlFor(latex, display) {
  const k = `${display ? 'D' : 'I'}${latex}`;
  if (mathCache.has(k)) return mathCache.get(k);
  const r = await renderLatex(latex, display);
  mathCache.set(k, r.mathml);
  if (mathCache.size > 500) mathCache.delete(mathCache.keys().next().value);
  return r.mathml;
}
/** render lazily: only when the formula comes near the viewport (long documents with many formulas) */
const io = typeof IntersectionObserver !== 'undefined' ? new IntersectionObserver((entries) => { for (const e of entries) if (e.isIntersecting && e.target.__ofDraw) { io.unobserve(e.target); e.target.__ofDraw(); } }, { rootMargin: '800px 0px' }) : null;

export const EquationEx = Node.create({
  name: 'equation',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addOptions() { return { onEdit: null }; },
  addAttributes() {
    return {
      latex: { default: '', parseHTML: (el) => el.getAttribute('data-latex') || el.textContent || '' },
      display: { default: false, parseHTML: (el) => el.getAttribute('data-display') === '1' },
      mathml: keep, omml: keep, ommlLatex: keep,
    };
  },
  parseHTML() { return [{ tag: 'span[data-type="equation"]', priority: 80 }]; },
  renderHTML({ HTMLAttributes: a }) { return ['span', mergeAttributes({ 'data-type': 'equation', 'data-latex': a.latex, 'data-display': a.display ? '1' : null, class: 'of-docs-eq' })]; },
  addInputRules() {
    // $x^2$ while typing → a formula; "$$ " at the start of an empty paragraph → a display formula (dialog)
    return [
      new InputRule({
        find: /(?:^|[\s(])\$([^$\n]{1,400}?[^\s$\\])\$$/,
        handler: ({ state, range, match }) => {
          const latex = match[1];
          const start = range.from + (match[0].length - latex.length - 2);
          state.tr.replaceWith(start, range.to, this.type.create({ latex }));
        },
      }),
      new InputRule({
        find: /^\$\$\s$/,
        handler: ({ state, range }) => {
          state.tr.delete(range.from, range.to);
          if (this.options.onEdit) setTimeout(() => this.options.onEdit(null, '', true), 0);
        },
      }),
    ];
  },
  addNodeView() {
    return ({ node, getPos }) => {
      const dom = document.createElement('span');
      dom.contentEditable = 'false';
      dom.title = t('Double-click to edit the formula');
      let cur = node; let token = 0;
      const draw = (n) => {
        const a = n.attrs;
        dom.className = `of-docs-eq${a.display ? ' of-eq-block' : ''}`;
        const mm = a.mathml || cachedMathml(a.latex, a.display);
        if (mm) { dom.innerHTML = mm.replace(/^[\s\S]*?(<math\b)/, '$1'); return; }
        dom.innerHTML = equationHtml(a.latex) || '<i>x</i>';
        const my = ++token;
        dom.__ofDraw = () => mathmlFor(a.latex, a.display).then((m) => { if (my === token && m) dom.innerHTML = m; }, () => {});
        if (io) io.observe(dom); else dom.__ofDraw();
      };
      draw(node);
      dom.addEventListener('dblclick', () => { if (this.options.onEdit) this.options.onEdit(typeof getPos === 'function' ? getPos() : null, cur.attrs.latex, cur.attrs.display); });
      return { dom, update: (n) => { if (n.type.name !== 'equation') return false; if (n.attrs.latex !== cur.attrs.latex || n.attrs.display !== cur.attrs.display || n.attrs.mathml !== cur.attrs.mathml) { cur = n; draw(n); } cur = n; return true; }, destroy: () => { if (io) io.unobserve(dom); }, ignoreMutation: () => true };
    };
  },
});

/* ---------------- tracked-change attributes of paragraphs and rows ---------------- */
export const RevisionAttrs = Extension.create({
  name: 'revisionAttrs',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          pmark: { default: null, keepOnSplit: false, parseHTML: () => null, renderHTML: (a) => (a.pmark ? { 'data-pmark': a.pmark.type, title: `${a.pmark.type === 'ins' ? t('New paragraph') : t('Paragraphs joined')} · ${a.pmark.author || ''}` } : {}) },
          pchange: { default: null, keepOnSplit: false, parseHTML: () => null, renderHTML: (a) => (a.pchange ? { 'data-pchange': '1' } : {}) },
        },
      },
      { types: ['tableRow'], attributes: { change: { default: null, parseHTML: () => null, renderHTML: (a) => (a.change ? { 'data-rowchange': a.change.type } : {}) } } },
    ];
  },
});
