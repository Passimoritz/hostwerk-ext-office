/*
 * Form controls (xlsx ctrlProps, io/xlsxControls.js) – the editing logic, pure (Node tests):
 * a check box toggles and writes TRUE/FALSE (Mixed = #N/A) into its linked cell, option buttons with the same linked
 * cell form a group (the cell gets the 1-based number of the chosen one), a drop-down / list box writes the 1-based
 * index of the picked item, a spinner / scroll bar its value. Everything is one undoable step.
 */
import { MAXC, parseRange } from './formula/refs.js';
import { errOf } from './formula/errors.js';

const K = (r, c) => r * MAXC + c;
function target(book, sheet, ref) {
  const g = ref && parseRange(String(ref).replace(/^=/, ''));
  if (!g) return null;
  const sh = g.sheet ? book.sheetByName(g.sheet) : sheet;
  return sh ? { sh, r: g.r1, c: g.c1, g } : null;
}
/** the items of a drop-down / list box (texts of its input range) */
export function controlItems(book, sheet, ctl) {
  const t = target(book, sheet, ctl.range);
  if (!t) return [];
  const out = [];
  for (let r = t.g.r1; r <= Math.min(t.g.r2, t.g.r1 + 999); r++) for (let c = t.g.c1; c <= Math.min(t.g.c2, t.g.c1 + 50); c++) out.push(book.text(t.sh, r, c));
  return out;
}
/** current state, taking the linked cell into account (Excel keeps both in sync) */
export function controlState(book, sheet, ctl) {
  const t = target(book, sheet, ctl.link);
  if (!t) return ctl;
  const v = book.value(t.sh, t.r, t.c);
  if (ctl.type === 'checkbox') return { ...ctl, checked: v === true ? 1 : v === false ? 0 : v && v.code === '#N/A' ? 2 : ctl.checked };
  if (ctl.type === 'option') { const group = (sheet.controls || []).filter((x) => x.type === 'option' && (x.group || x.link) === (ctl.group || ctl.link)); return { ...ctl, checked: typeof v === 'number' && group[v - 1] === ctl ? 1 : typeof v === 'number' ? 0 : ctl.checked }; }
  if (ctl.type === 'dropdown' || ctl.type === 'list') return { ...ctl, sel: typeof v === 'number' ? v : ctl.sel };
  if (ctl.type === 'spinner' || ctl.type === 'scrollbar') return { ...ctl, val: typeof v === 'number' ? v : ctl.val };
  return ctl;
}
function writeLink(book, sheet, ctl, value) {
  const t = target(book, sheet, ctl.link);
  if (!t) return;
  const old = t.sh.cells.get(K(t.r, t.c));
  book.setCell(t.sh, t.r, t.c, { ...(old || {}), v: value, f: undefined, _si: undefined });
}
/**
 * action: { toggle: true } | { sel: n } | { val: n } | { step: ±1 }. Returns true when something changed.
 */
export function applyControl(book, sheet, id, action, label = 'Form control') {
  const ctl = (sheet.controls || []).find((x) => x.id === id);
  if (!ctl) return false;
  const cur = controlState(book, sheet, ctl);
  let changed = false;
  book.run(label, () => {
    book.touch(sheet);
    const list = sheet.controls.map((x) => ({ ...x }));
    const me = list.find((x) => x.id === id);
    if (ctl.type === 'checkbox' && action.toggle) {
      me.checked = cur.checked === 1 ? 0 : 1; changed = true;
      writeLink(book, sheet, me, me.checked === 1);
    } else if (ctl.type === 'option' && action.toggle) {
      const group = list.filter((x) => x.type === 'option' && (x.group || x.link || x.id) === (ctl.group || ctl.link || ctl.id));
      group.forEach((x) => { x.checked = x.id === id ? 1 : 0; });
      changed = true;
      writeLink(book, sheet, me, group.indexOf(me) + 1);
    } else if ((ctl.type === 'dropdown' || ctl.type === 'list') && action.sel != null) {
      me.sel = action.sel; changed = true;
      writeLink(book, sheet, me, action.sel);
    } else if ((ctl.type === 'spinner' || ctl.type === 'scrollbar') && (action.val != null || action.step)) {
      const v = action.val != null ? action.val : (cur.val || 0) + action.step * (ctl.inc || 1);
      me.val = Math.max(ctl.min ?? 0, Math.min(ctl.max ?? 30000, v)); changed = true;
      writeLink(book, sheet, me, me.val);
    }
    sheet.controls = list;
  });
  return changed;
}
export const MIXED = errOf('#N/A');
