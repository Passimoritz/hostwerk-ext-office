/*
 * Track changes: insertion / deletion / format-change marks (author, date), recording while "Track changes" is on
 * (typing, Backspace/Delete incl. word-wise, cut, paste, IME input), plus – like Word – paragraph splits (Enter) and
 * joins (Backspace/Delete across a paragraph end) as inserted/deleted paragraph marks, character formatting
 * (w:rPrChange), paragraph formatting (w:pPrChange) and inserted/deleted table rows. Accept/reject single or all,
 * next/previous.
 */
import { Extension, Mark } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import { AddMarkStep, RemoveMarkStep, ReplaceStep, ReplaceAroundStep } from '@tiptap/pm/transform';

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
const TEXTBLOCK = new Set(['paragraph', 'heading']);
const FORMAT_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'textStyle', 'highlight', 'subscript', 'superscript', 'code']);
const PARA_KEYS = ['textAlign', 'lineHeight', 'spaceBefore', 'spaceAfter', 'indent', 'firstLine'];

/**
 * all changes in document order: [{ from, to, type, attrs }] – type insertion|deletion|formatChange (text, neighbours
 * merged), pmark (a paragraph mark: attrs.type ins|del), pchange (paragraph formatting), row (attrs.type ins|del)
 */
export function changeRanges(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (TEXTBLOCK.has(node.type.name)) {
      if (node.attrs.pmark) out.push({ from: pos + node.nodeSize - 1, to: pos + node.nodeSize, type: 'pmark', attrs: node.attrs.pmark, node: pos });
      if (node.attrs.pchange) out.push({ from: pos + 1, to: pos + node.nodeSize - 1, type: 'pchange', attrs: node.attrs.pchange, node: pos });
    }
    if (node.type.name === 'tableRow' && node.attrs.change) out.push({ from: pos + 1, to: pos + node.nodeSize - 1, type: 'row', attrs: node.attrs.change, node: pos });
    if (!node.isInline) return;
    for (const m of node.marks) {
      if (!['insertion', 'deletion', 'formatChange'].includes(m.type.name)) continue;
      const last = out[out.length - 1];
      if (last && last.type === m.type.name && last.to === pos && last.attrs.author === m.attrs.author) last.to = pos + node.nodeSize;
      else out.push({ from: pos, to: pos + node.nodeSize, type: m.type.name, attrs: m.attrs });
    }
  });
  return out.sort((a, b) => a.from - b.from);
}

/* ---------- Word run/paragraph properties of the "before" state (rPrChange / pPrChange) ---------- */
const hex = (c) => { const m = /^#?([0-9a-f]{6})$/i.exec(String(c || '')); return m ? m[1].toUpperCase() : null; };
export function rPrXml(marks) {
  let s = '';
  for (const m of marks) {
    const a = m.attrs || {};
    if (m.type.name === 'bold') s += '<w:b/>';
    else if (m.type.name === 'italic') s += '<w:i/>';
    else if (m.type.name === 'underline') s += '<w:u w:val="single"/>';
    else if (m.type.name === 'strike') s += '<w:strike/>';
    else if (m.type.name === 'superscript') s += '<w:vertAlign w:val="superscript"/>';
    else if (m.type.name === 'subscript') s += '<w:vertAlign w:val="subscript"/>';
    else if (m.type.name === 'textStyle') {
      const f = a.fontFamily && String(a.fontFamily).split(',')[0].replace(/["']/g, '').trim();
      if (f) s += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}"/>`;
      const c = hex(a.color); if (c) s += `<w:color w:val="${c}"/>`;
      const pt = parseFloat(a.fontSize); if (pt) s += `<w:sz w:val="${Math.round(pt * 2)}"/>`;
    }
  }
  return `<w:rPr>${s}</w:rPr>`;
}
export function pPrXml(a) {
  let s = '';
  if (a.spaceBefore != null || a.spaceAfter != null || a.lineHeight) s += `<w:spacing${a.spaceBefore != null ? ` w:before="${Math.round(a.spaceBefore * 20)}"` : ''}${a.spaceAfter != null ? ` w:after="${Math.round(a.spaceAfter * 20)}"` : ''}${a.lineHeight ? ` w:line="${Math.round(a.lineHeight * 240)}" w:lineRule="auto"` : ''}/>`;
  if (a.indent || a.firstLine) s += `<w:ind w:left="${Math.round((a.indent || 0) * 20)}"${a.firstLine < 0 ? ` w:hanging="${Math.round(-a.firstLine * 20)}"` : a.firstLine ? ` w:firstLine="${Math.round(a.firstLine * 20)}"` : ''}/>`;
  if (a.textAlign && a.textAlign !== 'left') s += `<w:jc w:val="${a.textAlign === 'justify' ? 'both' : a.textAlign}"/>`;
  return `<w:pPr>${s}</w:pPr>`;
}
/** the paragraph attributes of a stored w:pPr (reject a paragraph formatting change) */
export function attrsOfPPr(xml) {
  const s = String(xml || '');
  const o = { textAlign: null, lineHeight: null, spaceBefore: null, spaceAfter: null, indent: null, firstLine: null };
  const jc = /<w:jc w:val="([^"]+)"/.exec(s); if (jc) o.textAlign = { both: 'justify', distribute: 'justify', center: 'center', right: 'right', end: 'right' }[jc[1]] || null;
  const sp = /<w:spacing\b([^>]*)\/>/.exec(s);
  if (sp) { const g = (n) => { const m = new RegExp(`w:${n}="(-?\\d+)"`).exec(sp[1]); return m ? Number(m[1]) : null; }; if (g('before') != null) o.spaceBefore = g('before') / 20; if (g('after') != null) o.spaceAfter = g('after') / 20; if (g('line') != null) o.lineHeight = Math.round((g('line') / 240) * 100) / 100; }
  const ind = /<w:ind\b([^>]*)\/>/.exec(s);
  if (ind) { const g = (n) => { const m = new RegExp(`w:${n}="(-?\\d+)"`).exec(ind[1]); return m ? Number(m[1]) : null; }; const l = g('left') != null ? g('left') : g('start'); if (l) o.indent = Math.round(l / 20); if (g('hanging')) o.firstLine = -Math.round(g('hanging') / 20); else if (g('firstLine')) o.firstLine = Math.round(g('firstLine') / 20); }
  return o;
}

/** parse the "old" Word run properties of a format change into marks (reject) */
function oldMarks(schema, old) {
  const m = [];
  const s = String(old || '');
  if (/<w:b\/>|<w:b w:val="(1|true|on)"/.test(s)) m.push(schema.marks.bold.create());
  if (/<w:i\/>|<w:i w:val="(1|true|on)"/.test(s)) m.push(schema.marks.italic.create());
  if (/<w:u w:val="(?!none)/.test(s) && schema.marks.underline) m.push(schema.marks.underline.create());
  if (/<w:strike\/>/.test(s)) m.push(schema.marks.strike.create());
  if (/<w:vertAlign w:val="superscript"/.test(s) && schema.marks.superscript) m.push(schema.marks.superscript.create());
  if (/<w:vertAlign w:val="subscript"/.test(s) && schema.marks.subscript) m.push(schema.marks.subscript.create());
  const color = /<w:color w:val="([0-9A-Fa-f]{6})"/.exec(s); const sz = /<w:sz w:val="(\d+)"/.exec(s); const font = /<w:rFonts[^>]*w:ascii="([^"]+)"/.exec(s);
  if (color || sz || font) m.push(schema.marks.textStyle.create({ color: color ? `#${color[1]}` : null, fontSize: sz ? `${Number(sz[1]) / 2}pt` : null, fontFamily: font ? font[1] : null }));
  return m;
}

/** join the textblock at `pos` with the following one (accepting a deleted / rejecting an inserted paragraph mark) */
function joinAfter(tr, pos) {
  const node = tr.doc.nodeAt(pos);
  if (!node) return;
  const end = pos + node.nodeSize;
  const next = tr.doc.nodeAt(end);
  if (next && next.isTextblock) { try { tr.join(end); return; } catch { /* different parents – fall back */ } }
  // the next paragraph is the first one of a list item / cell: move its text up
  const $e = tr.doc.resolve(Math.min(end + 1, tr.doc.content.size));
  let n2 = null; let p2 = -1;
  tr.doc.nodesBetween(end, Math.min(tr.doc.content.size, end + 400), (x, p) => { if (n2) return false; if (x.isTextblock && p >= end) { n2 = x; p2 = p; return false; } return true; });
  void $e;
  if (n2) { tr.insert(end - 1, n2.content); const m = tr.mapping.map(p2); tr.delete(m, m + n2.nodeSize); }
}

function resolve(tr, r, accept) {
  const schema = tr.doc.type.schema;
  const at = r.node != null ? tr.mapping.map(r.node) : null;
  if (r.type === 'insertion') { if (accept) tr.removeMark(tr.mapping.map(r.from), tr.mapping.map(r.to), schema.marks.insertion); else tr.delete(tr.mapping.map(r.from), tr.mapping.map(r.to)); }
  else if (r.type === 'deletion') { if (accept) tr.delete(tr.mapping.map(r.from), tr.mapping.map(r.to)); else tr.removeMark(tr.mapping.map(r.from), tr.mapping.map(r.to), schema.marks.deletion); }
  else if (r.type === 'formatChange') {
    const f = tr.mapping.map(r.from); const t2 = tr.mapping.map(r.to);
    tr.removeMark(f, t2, schema.marks.formatChange);
    if (!accept) {
      for (const n of FORMAT_MARKS) if (schema.marks[n]) tr.removeMark(f, t2, schema.marks[n]);
      for (const m of oldMarks(schema, r.attrs.old)) tr.addMark(f, t2, m);
    }
  } else if (r.type === 'pmark') {
    const node = tr.doc.nodeAt(at); if (!node) return;
    tr.setNodeMarkup(at, undefined, { ...node.attrs, pmark: null });
    if ((r.attrs.type === 'del') === accept) joinAfter(tr, at);
  } else if (r.type === 'pchange') {
    const node = tr.doc.nodeAt(at); if (!node) return;
    tr.setNodeMarkup(at, undefined, { ...node.attrs, ...(accept ? {} : attrsOfPPr(r.attrs.old)), pchange: null });
  } else if (r.type === 'row') {
    const node = tr.doc.nodeAt(at); if (!node) return;
    if ((r.attrs.type === 'del') === accept) {
      // the row goes: if it is the only row, the table goes
      const $r = tr.doc.resolve(at);
      const table = $r.parent;
      if (table.childCount <= 1) tr.delete($r.before(), $r.after()); else tr.delete(at, at + node.nodeSize);
    } else tr.setNodeMarkup(at, undefined, { ...node.attrs, change: null });
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
      // text first (back to front), structure after it – positions stay valid through the mapping
      const text = rs.filter((r) => ['insertion', 'deletion', 'formatChange'].includes(r.type)).sort((a, b) => b.from - a.from);
      const struct = rs.filter((r) => !['insertion', 'deletion', 'formatChange'].includes(r.type)).sort((a, b) => b.from - a.from);
      for (const r of [...text, ...struct]) resolve(tr, r, accept);
      tr.setMeta('ofTrackSkip', true);
      if (dispatch) dispatch(tr);
      return true;
    };
    const go = (dir) => ({ state, tr, dispatch, view }) => {
      const rs = changeRanges(state.doc);
      if (!rs.length) return false;
      const pos = state.selection.from;
      const r = dir > 0 ? rs.find((x) => x.from > pos) || rs[0] : [...rs].reverse().find((x) => x.to < pos) || rs[rs.length - 1];
      tr.setSelection(TextSelection.create(state.doc, Math.max(1, r.from), Math.min(state.doc.content.size, Math.max(r.from, r.to)))).scrollIntoView();
      if (dispatch) dispatch(tr);
      if (view) view.focus();
      return true;
    };
    return {
      setTrackChanges: (on) => () => { this.storage.enabled = !!on; return true; },
      acceptChange: () => run(true, false), rejectChange: () => run(false, false),
      acceptAllChanges: () => run(true, true), rejectAllChanges: () => run(false, true),
      nextChange: () => go(1), prevChange: () => go(-1),
      /** delete the row(s) of the selection – recorded as deleted rows (text marked deleted) while tracking */
      trackedDeleteRow: () => ({ state, tr, dispatch }) => {
        const $p = state.selection.$from;
        let d = $p.depth; while (d > 0 && $p.node(d).type.name !== 'tableRow') d--;
        if (!d) return false;
        const pos = $p.before(d); const row = $p.node(d);
        const a = { id: this.storage.session, author: this.storage.author || 'Unknown', date: new Date().toISOString() };
        if (row.attrs.change && row.attrs.change.type === 'ins') { tr.delete(pos, pos + row.nodeSize); } else {
          tr.setNodeMarkup(pos, undefined, { ...row.attrs, change: { type: 'del', ...a } });
          tr.addMark(pos + 1, pos + row.nodeSize - 1, state.schema.marks.deletion.create(a));
        }
        tr.setMeta('ofTrackSkip', true);
        if (dispatch) dispatch(tr);
        return true;
      },
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
      // paragraph marks inside the range: deleted paragraph marks (Word joins them on accept)
      const marksAt = [];
      tr.doc.nodesBetween(from, to, (node, pos) => { if (TEXTBLOCK.has(node.type.name) && pos + node.nodeSize - 1 < to && pos + node.nodeSize - 1 >= from) marksAt.push(pos); return !node.isTextblock; });
      for (const pos of marksAt) { const n = tr.doc.nodeAt(pos); if (n && !(n.attrs.pmark && n.attrs.pmark.type === 'ins')) tr.setNodeMarkup(pos, undefined, { ...n.attrs, pmark: { type: 'del', ...attrs() } }); }
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
    /** Backspace at the start / Delete at the end of a paragraph: the paragraph mark is marked deleted (no join) */
    const markJoin = (view, dir) => {
      const { state } = view;
      const $p = state.selection.$from;
      if (!TEXTBLOCK.has($p.parent.type.name)) return false;
      let prevPos = null;
      if (dir < 0) {
        const before = $p.before();
        const $b = state.doc.resolve(before);
        const idx = $b.index();
        if (idx === 0) return false;
        const prev = $b.parent.child(idx - 1);
        if (!TEXTBLOCK.has(prev.type.name)) return false;
        prevPos = before - prev.nodeSize;
      } else {
        prevPos = $p.before();
        const after = $p.after();
        const nxt = state.doc.nodeAt(after);
        if (!nxt || !TEXTBLOCK.has(nxt.type.name)) return false;
      }
      const prev = state.doc.nodeAt(prevPos);
      const tr = state.tr;
      if (prev.attrs.pmark && prev.attrs.pmark.type === 'ins') {
        // a paragraph made while tracking: simply join again
        tr.setNodeMarkup(prevPos, undefined, { ...prev.attrs, pmark: null });
        tr.join(prevPos + prev.nodeSize);
      } else {
        tr.setNodeMarkup(prevPos, undefined, { ...prev.attrs, pmark: { type: 'del', ...attrs() } });
        if (dir < 0) tr.setSelection(TextSelection.create(tr.doc, prevPos + prev.nodeSize - 1));
      }
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
              if ($p.parentOffset === 0) { if (markJoin(view, -1)) { e.preventDefault(); return true; } return false; }
              const n = e.ctrlKey || e.altKey || e.metaKey ? (/\S*\s*$/.exec(textBefore) || [''])[0].length || 1 : 1;
              from = sel.from - n; to = sel.from;
            } else {
              if ($p.parentOffset >= $p.parent.content.size) { if (markJoin(view, 1)) { e.preventDefault(); return true; } return false; }
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
      // everything that bypasses the handlers above: IME/drop input, Enter (paragraph splits), formatting, paragraph
      // formatting and inserted table rows
      appendTransaction(trs, oldState, newState) {
        if (!ext.storage.enabled) return null;
        const user = trs.filter((tr) => tr.docChanged && !tr.getMeta('ofTrackSkip') && !tr.getMeta('ofPages') && !tr.getMeta('ofSilent') && !tr.getMeta('y-sync$'));
        if (!user.length) return null;
        const tr = newState.tr;
        const a = attrs();
        const schema = newState.schema;
        // map a position of transaction i / step k to the new document
        const toNew = (ti, k, pos, assoc = 1) => { let p = pos; const t0 = user[ti]; for (let j = k + 1; j < t0.steps.length; j++) p = t0.steps[j].getMap().map(p, assoc); for (let i = trs.indexOf(t0) + 1; i < trs.length; i++) p = trs[i].mapping.map(p, assoc); return p; };
        const textInput = (t0) => t0.getMeta('composition') !== undefined || t0.getMeta('uiEvent') === 'drop' || t0.getMeta('inputType');
        user.forEach((t0, ti) => t0.steps.forEach((step, k) => {
          const before = t0.docs[k];
          if (step instanceof ReplaceStep) {
            const sl = step.slice;
            // Enter: a split = an empty replace with a closed + opened textblock
            if (step.from === step.to && sl.openStart > 0 && sl.openEnd > 0 && sl.content.childCount === 2) {
              const p = toNew(ti, k, step.from, -1);
              const $p = newState.doc.resolve(Math.min(p, newState.doc.content.size));
              for (let d = $p.depth; d > 0; d--) {
                if (TEXTBLOCK.has($p.node(d).type.name)) { const pos = $p.before(d); const n = newState.doc.nodeAt(pos); if (n && !n.attrs.pmark) tr.setNodeMarkup(pos, undefined, { ...n.attrs, pmark: { type: 'ins', ...a } }); break; }
              }
            }
            // inserted table rows
            let rows = 0; sl.content.descendants((n) => { if (n.type.name === 'tableRow') rows++; return n.type.name !== 'tableRow'; });
            if (rows) {
              const s = toNew(ti, k, step.from, 1); const e = toNew(ti, k, step.from + sl.size, -1);
              newState.doc.nodesBetween(s, Math.min(e + sl.size, newState.doc.content.size), (n, pos) => { if (n.type.name === 'tableRow' && pos >= s - 1 && !n.attrs.change) { tr.setNodeMarkup(pos, undefined, { ...n.attrs, change: { type: 'ins', ...a } }); return false; } return n.type.name !== 'tableRow'; });
            }
            // composition / drop: mark the inserted text
            if (textInput(t0) && sl.size) {
              const s = toNew(ti, k, step.from, 1); const e = toNew(ti, k, step.from + sl.size, -1);
              if (e > s) tr.addMark(s, Math.min(e, newState.doc.content.size), schema.marks.insertion.create(a));
            }
          } else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
            // character formatting: remember the formatting before (rPrChange)
            if (!FORMAT_MARKS.has(step.mark.type.name)) return;
            before.nodesBetween(step.from, step.to, (n, pos) => {
              if (!n.isText) return;
              if (n.marks.some((m) => ['insertion', 'formatChange'].includes(m.type.name))) return;
              const s = toNew(ti, k, Math.max(step.from, pos), 1); const e = toNew(ti, k, Math.min(step.to, pos + n.nodeSize), -1);
              if (e > s) tr.addMark(s, e, schema.marks.formatChange.create({ ...a, old: rPrXml(n.marks.filter((m) => FORMAT_MARKS.has(m.type.name))) }));
            });
          } else if (step instanceof ReplaceAroundStep && step.structure && step.insert === 1 && step.gapFrom === step.from + 1) {
            // setNodeMarkup on a paragraph (alignment, spacing, indents): remember the paragraph formatting before
            const old = before.nodeAt(step.from);
            const p = toNew(ti, k, step.from, 1);
            const now = newState.doc.nodeAt(p);
            if (!old || !now || !TEXTBLOCK.has(old.type.name) || now.type !== old.type || now.attrs.pchange) return;
            if (PARA_KEYS.some((key) => (old.attrs[key] || null) !== (now.attrs[key] || null))) tr.setNodeMarkup(p, undefined, { ...now.attrs, pchange: { ...a, old: pPrXml(old.attrs) } });
          }
        }));
        tr.setMeta('ofTrackSkip', true);
        tr.setMeta('addToHistory', true);
        return tr.steps.length ? tr : null;
      },
    })];
  },
});
