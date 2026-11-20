/*
 * Opens one of the lazily loaded Word dialogs (word/dialogs.jsx in the docs-math sub-chunk) for the editor and applies
 * what it returns – paragraph, font, table properties, borders and shading, sort.
 */
import { useEffect, useState } from 'react';
import { toast } from '@hostwerk/ui';
import { wordPack } from '../lazy.js';

function tableCtx(editor) {
  const $p = editor.state.selection.$from;
  let table = null; let row = null; let cell = null; let rowIndex = 0;
  for (let d = $p.depth; d > 0; d--) {
    const n = $p.node(d);
    if (!cell && (n.type.name === 'tableCell' || n.type.name === 'tableHeader')) cell = n;
    else if (!row && n.type.name === 'tableRow') { row = n; rowIndex = $p.index(d - 1); }
    else if (!table && n.type.name === 'table') table = n;
  }
  return { table, row, cell, rowIndex };
}

export function WordDialog({ name, editor, onClose }) {
  const [pack, setPack] = useState(null);
  useEffect(() => { let on = true; wordPack().then((p) => on && setPack(p), (e) => { toast.error(e); onClose(); }); return () => { on = false; }; }, []); // eslint-disable-line react-hooks/exhaustive-deps
  if (!pack) return null;
  const c = () => editor.chain().focus();
  const done = (fn) => (v) => { fn(v); onClose(); };
  if (name === 'paragraph') {
    const attrs = { ...editor.getAttributes('paragraph'), ...editor.getAttributes('heading') };
    return <pack.ParagraphDialog attrs={attrs} onClose={onClose} onSave={done((p) => { const { textAlign, ...rest } = p; let ch = c().setParagraphFormat(rest); ch = textAlign ? ch.setTextAlign(textAlign) : ch.unsetTextAlign(); ch.run(); })} />;
  }
  if (name === 'font') {
    const marks = { strike: editor.isActive('strike'), superscript: editor.isActive('superscript'), subscript: editor.isActive('subscript') };
    return <pack.FontDialog attrs={editor.getAttributes('textStyle')} marks={marks} onClose={onClose} onSave={done(({ marks: m, style }) => {
      let ch = c();
      for (const k of ['strike', 'superscript', 'subscript']) if (m[k] !== marks[k]) ch = m[k] ? ch.setMark(k) : ch.unsetMark(k);
      ch.setRunFormat(style).run();
    })} />;
  }
  const tc = tableCtx(editor);
  if (name === 'tableProps' && tc.table) {
    return <pack.TablePropsDialog table={tc.table.attrs} row={{ ...tc.row.attrs, header: tc.row.childCount && Array.from({ length: tc.row.childCount }, (_, i) => tc.row.child(i)).every((x) => x.type.name === 'tableHeader') }} cell={tc.cell ? tc.cell.attrs : {}} colWidth={tc.cell && tc.cell.attrs.colwidth ? tc.cell.attrs.colwidth[0] : null} onClose={onClose} onSave={done((v) => {
      const ch = c().setTableAttrs(v.table).setRowsAttrs(v.row).setCellsAttrs(v.cell);
      if (v.col) ch.setCellAttribute('colwidth', [Math.round(v.col * 100) / 100]);
      ch.setHeaderRows(v.header ? Math.max(1, tc.rowIndex + 1) : 0).run();
    })} />;
  }
  if (name === 'borders' && tc.table) {
    return <pack.BordersDialog onClose={onClose} onSave={done(({ to, borders, shading }) => {
      if (to === 'table') c().setTableAttrs({ borders: { ...(tc.table.attrs.borders || {}), ...borders } }).run();
      else {
        // cells: inside lines become the cells' own top/bottom/left/right
        const own = {};
        for (const [k, v] of Object.entries(borders)) { if (k === 'insideH') { own.top = v; own.bottom = v; } else if (k === 'insideV') { own.left = v; own.right = v; } else own[k] = v; }
        c().setCellsAttrs((a) => ({ borders: { ...(a.borders || {}), ...own }, ...(shading ? { backgroundColor: shading } : {}) })).run();
      }
      if (shading && to === 'table') c().setCellsAttrs({ backgroundColor: shading }).run();
    })} />;
  }
  if (name === 'sort' && tc.table) {
    const first = tc.table.firstChild; const cols = [];
    if (first) first.forEach((cl) => cols.push(cl.textContent.trim().slice(0, 40)));
    return <pack.SortDialog columns={cols} onClose={onClose} onSave={done((o) => { if (!editor.chain().focus().sortTable(o).run()) toast.error('This table cannot be sorted (merged cells).'); })} />;
  }
  return null;
}
