/*
 * Track changes: insertion / deletion / format-change marks (author, date), recording while "Track changes" is on
 * (typing, Backspace/Delete incl. word-wise, cut, paste, IME input), accept/reject single or all, next/previous.
 * Paragraph splits/joins and table structure changes are not recorded (Word records them; we keep the text).
 */
import { Extension, Mark } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';

const revAttrs = { id: { default: null, parseHTML: (el) => el.getAttribute('data-id') }, author: { default: '', parseHTML: (el) => el.getAttribute('data-author') || '' }, date: { default: null, parseHTML: (el) => el.getAttribute('data-date') }, move: { default: false, parseHTML: () => false, renderHTML: () => ({}) } };
const title = (a, what) => `${what}: ${a.author || '?'}${a.date ? ` · ${new Date(a.date).toLocaleString()}` : ''}`;

export const Insertion = Mark.create({
  name: 'insertion',
  inclusive: false,
  excludes: 'deletion',
  addAttributes() { return revAttrs; },
  parseHTML() { return [{ tag: 'ins' }]; },
  renderHTML({ HTMLAttributes: a }) { return ['ins', { class: 'of-docs-ins', 'data-id': a.id, 'data-author': a.author, 'data-date': a.date, title: title(a, '+'), style: `--of-rev:${colorOf(a.author)}` }, 0]; },
});
export const Deletion = Mark.create({
  name: 'deletion',
  inclusive: false,
  excludes: 'insertion',
  addAttributes() { return revAttrs; },
  parseHTML() { return [{ tag: 'del' }]; },
  renderHTML({ HTMLAttributes: a }) { return ['del', { class: 'of-docs-del', 'data-id': a.id, 'data-author': a.author, 'data-date': a.date, title: title(a, '−'), style: `--of-rev:${colorOf(a.author)}` }, 0]; },
});
export const FormatChange = Mark.create({
  name: 'formatChange',
  inclusive: false,
  excludes: '',
  addAttributes() { return { ...revAttrs, old: { default: '', parseHTML: () => '', renderHTML: () => ({}) } }; },
  parseHTML() { return []; },
  renderHTML({ HTMLAttributes: a }) { return ['span', { class: 'of-docs-fmt', title: title(a, '≈'), style: `--of-rev:${colorOf(a.author)}` }, 0]; },
});

const PALETTE = ['#1a5fb4', '#c01c28', '#26a269', '#a347ba', '#c64600', '#0e7c86', '#865e3c', '#e01b7a'];
export function colorOf(name) { let h = 0; for (const c of String(name || '')) h = (h * 31 + c.charCodeAt(0)) >>> 0; return PALETTE[h % PALETTE.length]; }

export const trackKey = new PluginKey('ofTrack');

/** all change ranges in document order: [{ from, to, type: 'insertion'|'deletion'|'formatChange', attrs }] (neighbours merged) */
export function changeRanges(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (!node.isInline) return;
    for (const m of node.marks) {
      if (!['insertion', 'deletion', 'formatChange'].includes(m.type.name)) continue;
      const last = out[out.length - 1];
      if (last && last.type === m.type.name && last.to === pos && last.attrs.author === m.attrs.author) last.to = pos + node.nodeSize;
      else out.push({ from: pos, to: pos + node.nodeSize, type: m.type.name, attrs: m.attrs });
    }
  });
  return out;
}

/** parse the "old" Word run properties of a format change into marks (reject) */
function oldMarks(schema, old) {
  const m = [];
  const s = String(old || '');
  if (/<w:b\/>|<w:b w:val="(1|true|on)"/.test(s)) m.push(schema.marks.bold.create());
  if (/<w:i\/>|<w:i w:val="(1|true|on)"/.test(s)) m.push(schema.marks.italic.create());
  if (/<w:u w:val="(?!none)/.test(s) && schema.marks.underline) m.push(schema.marks.underline.create());
  if (/<w:strike\/>/.test(s)) m.push(schema.marks.strike.create());
  const color = /<w:color w:val="([0-9A-Fa-f]{6})"/.exec(s); const sz = /<w:sz w:val="(\d+)"/.exec(s); const font = /<w:rFonts[^>]*w:ascii="([^"]+)"/.exec(s);
  if (color || sz || font) m.push(schema.marks.textStyle.create({ color: color ? `#${color[1]}` : null, fontSize: sz ? `${Number(sz[1]) / 2}pt` : null, fontFamily: font ? font[1] : null }));
  return m;
}

function resolve(tr, r, accept) {
  const schema = tr.doc.type.schema;
  if (r.type === 'insertion') { if (accept) tr.removeMark(r.from, r.to, schema.marks.insertion); else tr.delete(r.from, r.to); }
  else if (r.type === 'deletion') { if (accept) tr.delete(r.from, r.to); else tr.removeMark(r.from, r.to, schema.marks.deletion); }
  else {
    tr.removeMark(r.from, r.to, schema.marks.formatChange);
    if (!accept) {
      for (const n of ['bold', 'italic', 'underline', 'strike', 'textStyle']) if (schema.marks[n]) tr.removeMark(r.from, r.to, schema.marks[n]);
      for (const m of oldMarks(schema, r.attrs.old)) tr.addMark(r.from, r.to, m);
    }
  }
}

export const TrackChanges = Extension.create({
  name: 'trackChanges',
  addStorage() { return { enabled: false, author: '', session: `s${Date.now().toString(36)}` }; },
  addCommands() {
    const ranges = (state, all) => {
      const rs = changeRanges(state.doc);
      if (all) return rs;
      const { from, to } = state.selection;
      return rs.filter((r) => r.to >= from && r.from <= Math.max(to, from));
    };
    const run = (accept, all) => ({ state, tr, dispatch }) => {
      const rs = ranges(state, all);
      if (!rs.length) return false;
      for (const r of [...rs].sort((a, b) => b.from - a.from)) resolve(tr, r, accept);
      tr.setMeta('ofTrackSkip', true);
      if (dispatch) dispatch(tr);
      return true;
    };
    const go = (dir) => ({ state, tr, dispatch, view }) => {
      const rs = changeRanges(state.doc);
      if (!rs.length) return false;
      const pos = state.selection.from;
      const r = dir > 0 ? rs.find((x) => x.from > pos) || rs[0] : [...rs].reverse().find((x) => x.to < pos) || rs[rs.length - 1];
      tr.setSelection(TextSelection.create(state.doc, r.from, r.to)).scrollIntoView();
      if (dispatch) dispatch(tr);
      if (view) view.focus();
      return true;
    };
    return {
      setTrackChanges: (on) => () => { this.storage.enabled = !!on; return true; },
      acceptChange: () => run(true, false), rejectChange: () => run(false, false),
      acceptAllChanges: () => run(true, true), rejectAllChanges: () => run(false, true),
      nextChange: () => go(1), prevChange: () => go(-1),
    };
  },
  addProseMirrorPlugins() {
    const ext = this;
    const attrs = () => ({ id: ext.storage.session, author: ext.storage.author || 'Unknown', date: new Date(Math.floor(Date.now() / 60000) * 60000).toISOString() });
    /** mark [from,to) deleted (text inserted while tracking is really removed) */
    const del = (tr, from, to) => {
      const schema = tr.doc.type.schema;
      const parts = [];
      tr.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isInline) return;
        const s = Math.max(from, pos); const e = Math.min(to, pos + node.nodeSize);
        if (s >= e) return;
        parts.push({ s, e, ins: node.marks.some((m) => m.type.name === 'insertion'), del: node.marks.some((m) => m.type.name === 'deletion') });
      });
      for (const p of parts.reverse()) { if (p.ins) tr.delete(p.s, p.e); else if (!p.del) tr.addMark(p.s, p.e, schema.marks.deletion.create(attrs())); }
    };
    const insertText = (view, from, to, text) => {
      const { state } = view;
      const tr = state.tr;
      if (from !== to) del(tr, from, to);
      const pos = tr.mapping.map(to);
      const marks = (state.storedMarks || state.doc.resolve(from).marks()).filter((m) => m.type.name !== 'deletion' && m.type.name !== 'insertion');
      tr.insert(pos, state.schema.text(text, [...marks, state.schema.marks.insertion.create(attrs())]));
      tr.setSelection(TextSelection.create(tr.doc, pos + text.length));
      tr.setMeta('ofTrackSkip', true);
      view.dispatch(tr.scrollIntoView());
      return true;
    };
    return [new Plugin({
      key: trackKey,
      props: {
        handleTextInput(view, from, to, text) { if (!ext.storage.enabled || view.composing) return false; return insertText(view, from, to, text); },
        handleKeyDown(view, e) {
          if (!ext.storage.enabled) return false;
          if (e.key !== 'Backspace' && e.key !== 'Delete') return false;
          const { state } = view;
          const sel = state.selection;
          let from = sel.from; let to = sel.to;
          if (sel.empty) {
            const $p = sel.$from;
            const textBefore = $p.parent.textBetween(0, $p.parentOffset, '\0', '\0');
            const textAfter = $p.parent.textBetween($p.parentOffset, $p.parent.content.size, '\0', '\0');
            if (e.key === 'Backspace') {
              if ($p.parentOffset === 0) return false;
              const n = e.ctrlKey || e.altKey || e.metaKey ? (/\S*\s*$/.exec(textBefore) || [''])[0].length || 1 : 1;
              from = sel.from - n; to = sel.from;
            } else {
              if ($p.parentOffset >= $p.parent.content.size) return false;
              const n = e.ctrlKey || e.altKey || e.metaKey ? (/^\s*\S*/.exec(textAfter) || [''])[0].length || 1 : 1;
              from = sel.from; to = sel.from + n;
            }
          }
          const tr = state.tr;
          del(tr, from, to);
          const caret = e.key === 'Backspace' ? tr.mapping.map(from, -1) : tr.mapping.map(to, 1);
          tr.setSelection(TextSelection.create(tr.doc, Math.min(caret, tr.doc.content.size)));
          tr.setMeta('ofTrackSkip', true);
          view.dispatch(tr.scrollIntoView());
          e.preventDefault();
          return true;
        },
        handlePaste(view, event, slice) {
          if (!ext.storage.enabled) return false;
          const { state } = view;
          const tr = state.tr;
          const { from, to } = state.selection;
          if (from !== to) del(tr, from, to);
          const pos = tr.mapping.map(to);
          const before = tr.doc.content.size;
          tr.replace(pos, pos, slice);
          const end = pos + (tr.doc.content.size - before);
          tr.addMark(pos, end, state.schema.marks.insertion.create(attrs()));
          tr.setSelection(TextSelection.near(tr.doc.resolve(end)));
          tr.setMeta('ofTrackSkip', true);
          view.dispatch(tr.scrollIntoView());
          return true;
        },
        handleDOMEvents: {
          cut(view, e) {
            if (!ext.storage.enabled || view.state.selection.empty) return false;
            // copy normally, then mark as deleted instead of removing
            setTimeout(() => { const tr = view.state.tr; del(tr, view.state.selection.from, view.state.selection.to); tr.setMeta('ofTrackSkip', true); view.dispatch(tr); }, 0);
            const text = view.state.doc.textBetween(view.state.selection.from, view.state.selection.to, '\n');
            if (e.clipboardData) { e.clipboardData.setData('text/plain', text); e.preventDefault(); }
            return true;
          },
        },
      },
      // composition (IME) and other input that bypasses the handlers above: mark inserted text
      appendTransaction(trs, oldState, newState) {
        if (!ext.storage.enabled) return null;
        const user = trs.filter((tr) => tr.docChanged && !tr.getMeta('ofTrackSkip') && (tr.getMeta('composition') !== undefined || tr.getMeta('uiEvent') === 'drop' || tr.getMeta('inputType')));
        if (!user.length) return null;
        const tr = newState.tr;
        for (const t of user) t.steps.forEach((step, k) => {
          const map = step.getMap();
          map.forEach((oldStart, oldEnd, newStart, newEnd) => {
            if (newEnd <= newStart) return;
            let s = newStart; let e = newEnd;
            for (let j = k + 1; j < t.steps.length; j++) { s = t.steps[j].getMap().map(s, 1); e = t.steps[j].getMap().map(e, -1); }
            if (e > s) tr.addMark(s, Math.min(e, newState.doc.content.size), newState.schema.marks.insertion.create(attrs()));
          });
        });
        tr.setMeta('ofTrackSkip', true);
        return tr.docChanged || tr.steps.length ? tr : null;
      },
    })];
  },
});
