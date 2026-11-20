/*
 * Form controls of xlsx files (Excel 2010+ "ctrlProps" + the legacy VML drawing): check boxes, option buttons,
 * drop-downs / list boxes, spinners, scroll bars, buttons, group boxes and labels.
 *
 *   readControls(zip, sheetXml, srels, anchorFrom) → [{ id, type, name, text, checked, sel, link, range, min, max,
 *     inc, val, macro, anchor, _orig: { ctrlPath, shapeId, vmlPath }, _sig }]
 *   patchControls(zip, sheet)  – writes changed state (checked / selected item / value) back into ctrlProp + VML
 *
 * The editor shows them over the grid; a check box toggles and a drop-down picks a value – both write the linked
 * cell like Excel (TRUE/FALSE, the 1-based index). Buttons run macros in Excel only (shown, not run).
 */
import { attrs, elements, first, parseRels, relsPathOf, unesc } from './xml.js';

const TYPE = { CheckBox: 'checkbox', Radio: 'option', Drop: 'dropdown', List: 'list', Spin: 'spinner', Scroll: 'scrollbar', Button: 'button', GBox: 'group', Label: 'label', EditBox: 'edit' };
const VML_TYPE = { Checkbox: 'checkbox', Radio: 'option', Drop: 'dropdown', List: 'list', Spin: 'spinner', Scroll: 'scrollbar', Button: 'button', GBox: 'group', Label: 'label', Edit: 'edit' };
export const ctrlSig = (c) => JSON.stringify([c.checked, c.sel, c.val]);
const CTRL_REL = /\/ctrlProp$/;
const VML_REL = /\/vmlDrawing$/;

function vmlShapes(vml) {
  const out = new Map();
  if (!vml) return out;
  for (const m of vml.matchAll(/<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g)) {
    const a = attrs(m[1]); const id = a.id || ''; const num = (/_s(\d+)$/.exec(id) || [])[1];
    if (!num) continue;
    const tb = /<v:textbox\b[^>]*>([\s\S]*?)<\/v:textbox>/.exec(m[2]);
    const text = tb ? unesc(tb[1].replace(/<br\s*\/?>/g, '\n').replace(/<[^>]+>/g, '')).replace(/\s+\n/g, '\n').trim() : '';
    const cd = /<x:ClientData\b[^>]*ObjectType="(\w+)"[^>]*>([\s\S]*?)<\/x:ClientData>/.exec(m[2]);
    const tag = (t) => { const q = cd && new RegExp(`<x:${t}>([\\s\\S]*?)</x:${t}>`).exec(cd[2]); return q ? unesc(q[1]).trim() : null; };
    out.set(num, { text, type: cd ? VML_TYPE[cd[1]] || cd[1] : null, checked: tag('Checked'), sel: tag('Sel'), link: tag('FmlaLink'), range: tag('FmlaRange'), macro: tag('FmlaMacro'), val: tag('Val') });
  }
  return out;
}

export async function readControls(zip, xml, srels, anchorFrom) {
  const out = [];
  const ctl = first(xml, 'controls', true);
  if (!ctl) return out;
  const vRel = srels.find((r) => VML_REL.test(r.type) && r.path);
  const vml = vRel && zip.file(vRel.path) ? await zip.file(vRel.path).async('string') : null;
  const shapes = vmlShapes(vml);
  const seen = new Set();
  for (const e of elements(ctl.inner, 'control', true)) {
    const a = attrs(e.a);
    const sid = a.shapeId; if (!sid || seen.has(sid)) continue;
    const rid = a['r:id'] || Object.entries(a).find(([k]) => /:id$/.test(k))?.[1];
    const rel = rid && srels.find((r) => r.id === rid && CTRL_REL.test(r.type));
    const pr = first(e.inner, 'anchor');
    const pt = (tag) => { const q = pr && first(pr.inner, tag); if (!q) return null; const g = (t) => { const x = first(q.inner, t); return x ? +x.inner : 0; }; return { c: g('col'), r: g('row'), dx: Math.round(g('colOff') / 9525), dy: Math.round(g('rowOff') / 9525) }; };
    const from = pt('from'); const to = pt('to');
    if (!from) continue;
    seen.add(sid);
    let p = {};
    if (rel && zip.file(rel.path)) { const x = await zip.file(rel.path).async('string'); const m = /<(?:\w+:)?formControlPr\b([^>]*?)\/?>/.exec(x); if (m) p = attrs(m[1]); }
    const v = shapes.get(sid) || {};
    const type = TYPE[p.objectType] || v.type || 'other';
    const items = [];
    const c = {
      id: `ctl${sid}`, type, name: a.name || '', text: v.text || a.name || '',
      checked: p.checked ? (p.checked === 'Checked' ? 1 : p.checked === 'Mixed' ? 2 : 0) : v.checked != null ? +v.checked || 0 : 0,
      sel: p.sel != null ? +p.sel : v.sel != null ? +v.sel : 0, val: p.val != null ? +p.val : v.val != null ? +v.val : 0,
      min: p.min != null ? +p.min : 0, max: p.max != null ? +p.max : 100, inc: p.inc != null ? +p.inc : 1,
      link: (p.fmlaLink || v.link || '').replace(/^=/, '') || null, range: (p.fmlaRange || v.range || '').replace(/^=/, '') || null,
      macro: v.macro || null, items, first: p.firstButton === '1',
      anchor: anchorFrom(from, to, to ? null : { w: 100, h: 20 }),
      _orig: { ctrlPath: rel ? rel.path : null, shapeId: sid, vmlPath: vRel ? vRel.path : null },
    };
    c._sig = ctrlSig(c);
    out.push(c);
  }
  // option buttons: a group starts at firstButton (only it carries the linked cell in Excel's files)
  let grp = null;
  for (const c of out) { if (c.type !== 'option') continue; if (c.first || !grp) grp = { id: c.id, link: c.link }; c.group = grp.id; if (!c.link) c.link = grp.link; }
  return out;
}

/** write changed control state back (ctrlProp attribute + VML ClientData) */
export async function patchControls(zip, sheet) {
  const changed = (sheet.controls || []).filter((c) => c._orig && c._sig !== ctrlSig(c));
  if (!changed.length) return;
  const vmlEdits = new Map();
  for (const c of changed) {
    const o = c._orig;
    if (o.ctrlPath && zip.file(o.ctrlPath)) {
      let x = await zip.file(o.ctrlPath).async('string');
      x = x.replace(/<((?:\w+:)?formControlPr)\b([^>]*?)(\/?)>/, (m, tag, at, sl) => {
        let s = at;
        const set = (k, v) => { s = new RegExp(`\\s${k}="[^"]*"`).test(s) ? s.replace(new RegExp(`\\s${k}="[^"]*"`), v == null ? '' : ` ${k}="${v}"`) : v == null ? s : `${s} ${k}="${v}"`; };
        if (c.type === 'checkbox' || c.type === 'option') set('checked', c.checked === 1 ? 'Checked' : c.checked === 2 ? 'Mixed' : null);
        if (c.type === 'dropdown' || c.type === 'list') set('sel', c.sel || 0);
        if (c.type === 'spinner' || c.type === 'scrollbar') set('val', c.val || 0);
        return `<${tag}${s}${sl}>`;
      });
      zip.file(o.ctrlPath, x);
    }
    if (o.vmlPath) { if (!vmlEdits.has(o.vmlPath)) vmlEdits.set(o.vmlPath, []); vmlEdits.get(o.vmlPath).push(c); }
  }
  for (const [path, list] of vmlEdits) {
    const f = zip.file(path); if (!f) continue;
    let vml = await f.async('string');
    for (const c of list) {
      vml = vml.replace(new RegExp(`(<v:shape\\b[^>]*\\bid="_x0000_s${c._orig.shapeId}"[^>]*>[\\s\\S]*?<x:ClientData\\b[^>]*>)([\\s\\S]*?)(</x:ClientData>)`), (m, a, body, z) => {
        const set = (tag, v) => { if (new RegExp(`<x:${tag}>`).test(body)) body = body.replace(new RegExp(`<x:${tag}>[\\s\\S]*?</x:${tag}>`), v == null ? '' : `<x:${tag}>${v}</x:${tag}>`); else if (v != null) body += `<x:${tag}>${v}</x:${tag}>`; };
        if (c.type === 'checkbox' || c.type === 'option') set('Checked', c.checked ? c.checked : null);
        if (c.type === 'dropdown' || c.type === 'list') set('Sel', c.sel || null);
        if (c.type === 'spinner' || c.type === 'scrollbar') set('Val', c.val || 0);
        return a + body + z;
      });
    }
    zip.file(path, vml);
  }
}
export { parseRels, relsPathOf };
