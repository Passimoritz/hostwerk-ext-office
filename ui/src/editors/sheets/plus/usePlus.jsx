/*
 * Office 1.1 features of the sheets editor in one hook (keeps Editor.jsx readable): tables, pivot tables, sparklines,
 * icon sets, pictures and shapes, outline buttons, views, live data from Hostwerk, database tables as sheets, forms,
 * rules, co-editing, compare, AI, calculation options and R1C1. Editor.jsx calls plus.run(cmd, arg) for commands it does
 * not know, plus.dialog(d, close) for dialogs and renders plus.overlay / plus.bars.
 */
import { useEffect, useRef, useState } from 'react';
import { t } from '@hostwerk/i18n';
import { toast, confirmDialog, promptDialog, Modal } from '@hostwerk/ui';
import { Database, RefreshCw, ChevronLeft, ChevronRight, Upload, ClipboardList, Radio, X, Users } from 'lucide-react';
import { Keep } from '../../_shared/keep.jsx';
import { MAXC, rangeA1, colName, parseRange, quoteSheet } from '../formula/refs.js';
import { addTable, setTableProps, toggleTotals, convertToRange, tableLook, autoExpand } from '../tables.js';
import { computePivot, renderPivot, pivotSource, pivotArea } from '../pivot.js';
import { outlineGroups, isCollapsed } from '../paintPlus.js';
import { toInternal, toDisplay } from '../formula/parser.js';
import { toR1C1, fromR1C1 } from '../formula/r1c1.js';
import { PivotPanel, SparklineDialog, FormDialog, AiDialog, CalcOptionsDialog, TableStyleGrid } from './dialogs.jsx';
import { ViewBar, KanbanView, CalendarView, GalleryView, sheetList } from './Views.jsx';
import { ObjBox, ShapeSVG, SlicerBox, ControlBox } from './objects.jsx';
import { controlItems, controlState, applyControl } from '../formcontrols.js';
import { startCollab } from './collab.js';

const K = (r, c) => r * MAXC + c;
const COLORS = ['#e8590c', '#7b2fbe', '#0b7285', '#c2255c', '#2b8a3e', '#1971c2', '#e67700', '#5c940d'];
const errMsg = (e) => (e && (e.message || e.error)) || String(e);
const missing = (e) => e && (e.status === 404 || e.status === 501);

export function usePlus(ctx) {
  const { book, sheet, sel, setSel, host, readOnly, loc, guard, setDialog, settings, activeRange, clipRange, switchSheet, focusGrid } = ctx;
  const S = (settings && settings.sheets) || {};
  const [objSel, setObjSel] = useState(null); // { kind: 'image'|'shape'|'slicer', id }
  const [r1c1, setR1c1] = useState(!!S.r1c1);
  const [peers, setPeers] = useState([]);
  const [collabOn, setCollabOn] = useState(false);
  const [pending, setPending] = useState(null); // form answers waiting { id, rows, sheet }
  const [dbState, setDbState] = useState({});
  const [, force] = useState(0);
  const collab = useRef(null);
  const live = useRef(new Map());
  const H = host || {};

  /* ---------- pivot labels in the UI language (Excel writes them into the file too) ---------- */
  if (book && !book.pivotLabels) book.pivotLabels = { grand: t('Grand Total'), values: t('Values'), blank: t('(blank)'), all: t('(All)'), multiple: t('(Multiple Items)'), caption: (fn, f) => t(`${{ sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' }[fn] || 'Sum'} of {f}`, { f }) };

  const tableAt = (sh, r, c) => (sh && (sh.tables || []).find((x) => r >= x.range.r1 && r <= x.range.r2 && c >= x.range.c1 && c <= x.range.c2)) || null;
  const pivotAt = (sh, r, c) => (sh && (sh.pivots || []).find((p) => p.range && r >= p.range.r1 && r <= p.range.r2 && c >= p.range.c1 && c <= p.range.c2)) || null;
  const curTable = sheet ? tableAt(sheet, sel.r, sel.c) : null;
  const curPivot = sheet ? pivotAt(sheet, sel.r, sel.c) : null;

  /* ---------- pivots ---------- */
  const applyPivot = (pv, { label = 'PivotTable', target = sheet } = {}) => {
    const res = computePivot(book, pv, book.pivotLabels);
    if (res.error) { toast.error(t(res.error)); return false; }
    const area = pivotArea(pv, res);
    const hit = (target.pivots || []).find((p) => p.id !== pv.id && p.range && p.range.r1 <= area.r2 && area.r1 <= p.range.r2 && p.range.c1 <= area.c2 && area.c1 <= p.range.c2);
    if (hit) { toast.error(t('The pivot table would overlap “{name}” – move it first.', { name: hit.name })); return false; }
    book.run(label, () => {
      book.touch(target, true);
      const area = renderPivot(book, target, pv, res);
      const next = { ...pv, range: area, _dirty: true };
      target.pivots = [...(target.pivots || []).filter((p) => p.id !== pv.id), next];
      book.tx.structural = true;
    });
    return true;
  };
  const refreshAll = () => { let n = 0; for (const sh of book.sheets) for (const pv of sh.pivots || []) if (!pv.readOnly && applyPivot(pv, { label: 'Refresh all', target: sh })) n++; if (n) toast.success(t('{n} pivot tables refreshed', { n })); };
  const newPivot = async () => {
    const g = clipRange(activeRange());
    let src = g;
    if (g.r1 === g.r2 && g.c1 === g.c2) { const tb = tableAt(sheet, g.r1, g.c1); src = tb ? tb.range : (book.dataRange(sheet) || g); }
    const source = `${quoteSheet(sheet.name)}!${rangeA1({ r1: src.r1, c1: src.c1, r2: src.r2, c2: src.c2 }).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`;
    const probe = pivotSource(book, { source });
    if (probe.error) { toast.error(t(probe.error)); return; }
    const where = await promptDialog({ title: t('New pivot table'), label: t('Put it on'), value: t('a new sheet'), confirm: t('Create') }).catch(() => null);
    if (where == null) return;
    let target = sheet; let at = { r: src.r1, c: src.c2 + 2 };
    if (where === t('a new sheet') || !parseRange(where)) { target = book.addSheet(book.uniqueSheetName(t('Pivot'))); at = { r: 2, c: 0 }; } else { const p = parseRange(where); if (p.sheet && book.sheetByName(p.sheet)) target = book.sheetByName(p.sheet); at = { r: p.r1, c: p.c1 }; }
    const n = book.allTables ? book.sheets.reduce((x, s) => x + (s.pivots || []).length, 0) + 1 : 1;
    const num = probe.fields.find((f, i) => probe.records.some((r) => typeof r[i] === 'number'));
    const txt = probe.fields.find((f, i) => !probe.records.some((r) => typeof r[i] === 'number'));
    const pv = { id: `pv${Date.now().toString(36)}`, name: `PivotTable${n}`, source, at, rows: txt ? [txt] : [], cols: [], values: num ? [{ field: num, fn: 'sum' }] : [], filters: [], style: 'PivotStyleLight16' };
    if (applyPivot(pv, { label: 'New pivot table', target })) {
      const i = book.sheets.indexOf(target); if (i !== book.active) switchSheet(i);
      setTimeout(() => setDialog({ kind: 'plus:pivot', id: pv.id, sheetId: target.id }), 0);
    }
  };

  /* ---------- pictures / shapes ---------- */
  const pickImage = () => {
    const inp = document.createElement('input'); inp.type = 'file'; inp.accept = 'image/png,image/jpeg,image/gif,image/svg+xml,image/webp';
    inp.onchange = () => { const f = inp.files && inp.files[0]; if (f) addImageFile(f); };
    inp.click();
  };
  const addImageFile = (file, at) => new Promise((resolve) => {
    if (file.size > 15 * 1048576) { toast.error(t('The picture is too big (more than 15 MB).')); resolve(null); return; }
    const fr = new FileReader();
    fr.onload = () => {
      const src = String(fr.result);
      const img = new Image();
      img.onload = () => {
        const s = Math.min(1, 480 / img.width, 360 / img.height);
        const im = { id: `im${Date.now().toString(36)}`, src, name: file.name, anchor: { r: at ? at.r : sel.r, c: at ? at.c : sel.c, dx: 0, dy: 0, w: Math.max(16, Math.round(img.width * s)), h: Math.max(16, Math.round(img.height * s)) } };
        book.run(t('Insert picture'), () => { book.touch(sheet); sheet.images = [...sheet.images, im]; });
        setObjSel({ kind: 'image', id: im.id }); resolve(im);
      };
      img.onerror = () => { toast.error(t('This picture can’t be read.')); resolve(null); };
      img.src = src;
    };
    fr.readAsDataURL(file);
  });
  const addShape = (kind) => {
    const sh = { id: `shp${Date.now().toString(36)}`, kind, anchor: { r: sel.r, c: sel.c, dx: 4, dy: 4, w: kind === 'line' ? 160 : kind === 'textbox' ? 200 : 140, h: kind === 'line' ? 12 : kind === 'textbox' ? 50 : 80 }, fill: kind === 'textbox' || kind === 'line' ? null : '#4472C4', line: kind === 'textbox' ? '#7f7f7f' : '#2F528F', text: kind === 'textbox' ? t('Text') : '' };
    book.run(t('Insert shape'), () => { book.touch(sheet); sheet.shapes = [...(sheet.shapes || []), sh]; });
    setObjSel({ kind: 'shape', id: sh.id });
  };
  const objList = (kind) => (kind === 'image' ? sheet.images : kind === 'shape' ? sheet.shapes || [] : []);
  const updObj = (kind, id, patch, label = 'Move') => book.run(t(label), () => { book.touch(sheet); if (kind === 'image') sheet.images = sheet.images.map((x) => (x.id === id ? { ...x, ...patch } : x)); else sheet.shapes = (sheet.shapes || []).map((x) => (x.id === id ? { ...x, ...patch } : x)); });
  const delObj = (kind, id) => { if (!guard()) return; book.run(t('Delete'), () => { book.touch(sheet); if (kind === 'image') sheet.images = sheet.images.filter((x) => x.id !== id); else sheet.shapes = (sheet.shapes || []).filter((x) => x.id !== id); }); setObjSel(null); };
  const orderObj = (kind, id, front) => book.run(t('Arrange'), () => { book.touch(sheet); const list = objList(kind).slice(); const i = list.findIndex((x) => x.id === id); const [o] = list.splice(i, 1); if (front) list.push(o); else list.unshift(o); if (kind === 'image') sheet.images = list; else sheet.shapes = list; });
  // drop pictures from Files / the desktop on the grid
  useEffect(() => {
    if (!H.insertFiles || readOnly || !book) return undefined;
    return H.insertFiles(async (files) => { for (const f of files) if (/^image\//.test(f.mime || '') || /\.(png|jpe?g|gif|svg|webp)$/i.test(f.name)) { const bytes = await f.bytes(); await addImageFile(new File([bytes], f.name, { type: f.mime || 'image/png' })); } });
  }, [book, readOnly]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- outline buttons ---------- */
  const outline = (btn) => {
    const axis = btn.axis;
    const map = axis === 'row' ? sheet.outlineRows : sheet.outlineCols;
    const levels = Math.max(0, ...map.values());
    if (btn.level) { // show levels up to btn.level
      const hide = []; const show = [];
      for (const [i, l] of map) (l >= btn.level ? hide : show).push(i);
      book.run(t('Outline'), () => { if (show.length) book.setHidden(sheet, axis, show, false); if (hide.length) book.setHidden(sheet, axis, hide, true); });
      return;
    }
    const gr = btn.group; const collapsed = isCollapsed(sheet, axis, gr);
    const idx = []; for (let i = gr.a; i <= gr.b; i++) idx.push(i);
    if (collapsed) { // expand: show this level, keep deeper collapsed groups closed
      const deeper = outlineGroups(map, levels).filter((x) => x.level > gr.level && x.a >= gr.a && x.b <= gr.b && isCollapsed(sheet, axis, x));
      const keep = new Set(); for (const d of deeper) for (let i = d.a; i <= d.b; i++) keep.add(i);
      book.setHidden(sheet, axis, idx.filter((i) => !keep.has(i)), false);
    } else book.setHidden(sheet, axis, idx, true);
  };

  /* ---------- live data from Hostwerk ---------- */
  const writeLive = (sh, def, res) => {
    const cols = res.columns || []; const rows = res.rows || [];
    book.run(t('Live data'), () => {
      book.touch(sh, true);
      const { r, c } = def.at;
      if (def.kind === 'value' || (res.value !== undefined && !cols.length)) { book.setCell(sh, r, c, { ...(sh.cells.get(K(r, c)) || {}), v: res.value ?? null, f: undefined }); return; }
      const old = def.range;
      if (old) for (let y = old.r1; y <= old.r2; y++) for (let x = old.c1; x <= old.c2; x++) book.setCell(sh, y, x, null);
      cols.forEach((col, j) => book.setCell(sh, r, c + j, { v: String(col.label || col.id) }));
      rows.forEach((row, i) => row.forEach((v, j) => { const type = cols[j] && cols[j].type; let val = v; if (type === 'date' && v) { const d = new Date(v); if (!Number.isNaN(+d)) val = (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes()) - Date.UTC(1899, 11, 30)) / 86400000; } book.setCell(sh, r + 1 + i, c + j, { v: val == null ? null : typeof val === 'object' ? JSON.stringify(val) : val, ...(type === 'date' ? { s: book.styles.intern({ nf: loc.dateTime }) } : {}) }); }));
      const range = { r1: r, c1: c, r2: r + Math.max(1, rows.length), c2: c + Math.max(0, cols.length - 1) };
      let tb = (sh.tables || []).find((x) => x.name === def.table);
      if (tb) sh.tables = sh.tables.map((x) => (x === tb ? { ...x, range, _dirty: true } : x));
      else { tb = { id: `t${Date.now().toString(36)}`, name: def.table, range, header: true, totals: false, style: 'TableStyleMedium2', stripes: true, columns: cols.map((x) => String(x.label || x.id)), totalsFn: {} }; sh.tables = [...(sh.tables || []), tb]; }
      sh.live = (sh.live || []).map((x) => (x.id === def.id ? { ...x, range, updatedAt: res.updatedAt || Date.now() } : x));
      book.tx.structural = true;
    });
  };
  const subscribeLive = (sh, def) => {
    if (!H.data || live.current.has(def.id)) return;
    try { const off = H.data.subscribe(def.source, def.params || {}, (res) => { const cur = (sh.live || []).find((x) => x.id === def.id); if (cur && !readOnly) writeLive(sh, cur, res); }); live.current.set(def.id, off); } catch { /* no live data on this server */ }
  };
  const insertLive = async () => {
    if (!guard()) return;
    if (!H.data) { toast(t('Live data needs a newer Hostwerk Office.')); return; }
    let pick;
    try { pick = await H.data.pick({}); } catch (e) { toast.error(missing(e) ? t('This Hostwerk has no live data sources yet.') : errMsg(e)); return; }
    if (!pick) return;
    const def = { id: `ld${Date.now().toString(36)}`, source: pick.source, params: pick.params || {}, label: pick.label || pick.source, at: { r: sel.r, c: sel.c }, table: `Hostwerk_${String(pick.source).replace(/[^\w]/g, '_').slice(0, 30)}_${(sheet.live || []).length + 1}`, refresh: 'live', kind: pick.kind };
    book.run(t('Live data'), () => { book.touch(sheet); sheet.live = [...(sheet.live || []), def]; });
    try { writeLive(sheet, def, await H.data.query(def.source, def.params)); subscribeLive(sheet, def); toast.success(t('Live data “{name}” inserted – it updates while the file is open.', { name: def.label })); } catch (e) { toast.error(errMsg(e)); }
  };
  const refreshLive = async () => { for (const sh of book.sheets) for (const def of sh.live || []) { try { writeLive(sh, def, await H.data.query(def.source, def.params)); } catch (e) { toast.error(errMsg(e)); return; } } };
  useEffect(() => { // subscribe on open
    if (!book || readOnly) return undefined;
    for (const sh of book.sheets) for (const def of sh.live || []) subscribeLive(sh, def);
    const m = live.current;
    return () => { for (const off of m.values()) { try { off(); } catch { /* gone */ } } m.clear(); };
  }, [book]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- database tables as sheets ---------- */
  const loadDb = async (sh, def, { offset = def.offset || 0 } = {}) => {
    setDbState((s) => ({ ...s, [sh.id]: { busy: true } }));
    try {
      const res = await H.db.rows({ server: def.server, db: def.db, table: def.table, offset, limit: def.limit || 200, sort: def.sort || undefined, filter: def.filter || undefined });
      const cols = res.columns || [];
      const names = cols.map((x) => (typeof x === 'string' ? x : x.name || x.id));
      const orig = {};
      book.run(t('Database table'), () => {
        book.touch(sh, true);
        sh.cells = new Map();
        names.forEach((n, j) => sh.cells.set(K(0, j), { v: n, s: book.styles.intern({ b: true, fill: '#E2EFDA', bd: { b: { s: 'thin', c: '#A9D08E' } } }) }));
        (res.rows || []).forEach((row, i) => { row.forEach((v, j) => { if (v != null) sh.cells.set(K(i + 1, j), { v: typeof v === 'object' ? JSON.stringify(v) : v }); }); orig[i + 1] = row; });
        sh.freeze = { rows: 1, cols: 0 };
        sh.db = { ...def, offset, total: res.total, primaryKey: res.primaryKey || null, columns: names };
        sh._dbOrig = orig;
        book.tx.structural = true;
      });
      setDbState((s) => ({ ...s, [sh.id]: { busy: false } }));
    } catch (e) { setDbState((s) => ({ ...s, [sh.id]: { busy: false, error: errMsg(e) } })); toast.error(errMsg(e)); }
  };
  const openDb = async () => {
    if (!H.db) { toast(t('Database tables need a newer Hostwerk Office.')); return; }
    let pick; try { pick = await H.db.pick(); } catch (e) { toast.error(missing(e) ? t('This Hostwerk has no database servers.') : errMsg(e)); return; }
    if (!pick) return;
    const sh = book.addSheet(book.uniqueSheetName(String(pick.table).slice(0, 28)));
    switchSheet(book.sheets.indexOf(sh));
    await loadDb(sh, { server: pick.server, db: pick.db, table: pick.table, offset: 0, limit: 200, sort: null, filter: null });
  };
  const dbChanges = (sh) => {
    const d = sh.db; if (!d || !sh._dbOrig) return [];
    const pk = d.primaryKey; if (!pk) return [];
    const pks = Array.isArray(pk) ? pk : [pk];
    const pi = pks.map((k) => d.columns.indexOf(k));
    const out = [];
    const rows = Object.keys(sh._dbOrig).map(Number);
    const maxR = book.dataRange(sh) ? book.dataRange(sh).r2 : 0;
    for (const r of rows) {
      const o = sh._dbOrig[r];
      const now = d.columns.map((_, j) => { const v = book.value(sh, r, j); return v === '' ? null : v; });
      if (now.every((v) => v == null)) { out.push({ op: 'delete', key: Object.fromEntries(pks.map((k, z) => [k, o[pi[z]]])) }); continue; }
      const values = {}; d.columns.forEach((n, j) => { const a = o[j] == null ? null : o[j]; const b = now[j]; if (String(a ?? '') !== String(b ?? '')) values[n] = b; });
      if (Object.keys(values).length) out.push({ op: 'update', key: Object.fromEntries(pks.map((k, z) => [k, o[pi[z]]])), values });
    }
    for (let r = Math.max(0, ...rows) + 1; r <= maxR; r++) { const values = {}; let any = false; d.columns.forEach((n, j) => { const v = book.value(sh, r, j); if (v != null && v !== '') { values[n] = v; any = true; } }); if (any) out.push({ op: 'insert', values }); }
    return out;
  };
  const writeDb = async (sh) => {
    const d = sh.db; if (!d) return;
    if (!d.primaryKey) { toast(t('This table has no primary key – it is read-only here.')); return; }
    const changes = dbChanges(sh);
    if (!changes.length) { toast(t('No changes to write.')); return; }
    try { const r = await H.db.write({ server: d.server, db: d.db, table: d.table, changes }); if (r) await loadDb(sh, d); } catch (e) { toast.error(errMsg(e)); }
  };

  /* ---------- forms ---------- */
  const formRange = () => { const tb = tableAt(sheet, sel.r, sel.c) || (sheet.tables || [])[0]; if (tb) return tb.range; const g = clipRange(activeRange()); if (g.r1 !== g.r2) return g; const u = sheetList(book, sheet); return u.g; };
  const createForm = async ({ title, fields, public: pub }) => {
    if (!H.forms) { toast(t('Forms need a newer Hostwerk Office.')); return; }
    if (H.save) { try { await H.save(); } catch { /* saved by the shell */ } }
    try {
      const res = await H.forms.create({ sheet: sheet.name, title, columns: fields.map((f) => ({ label: f.label, type: f.type, required: f.required, options: f.options })), timestamp: false, header: false, public: pub });
      book.run(t('Form'), () => { book.touch(sheet); sheet.form = { id: res.id, url: res.url, columns: fields.map((f) => f.c) }; });
      if (H.forms.open) H.forms.open({ id: res.id }); else toast.success(t('Form published: {url}', { url: res.url }));
    } catch (e) { toast.error(errMsg(e)); }
  };
  const checkPending = async () => {
    for (const sh of book.sheets) {
      if (!sh.form || !sh.form.id) continue;
      try {
        const r = await fetch(`/api/office/forms/${encodeURIComponent(sh.form.id)}/pending`, { credentials: 'same-origin' });
        if (!r.ok) continue;
        const j = await r.json();
        if (j.rows && j.rows.length) { setPending({ id: sh.form.id, rows: j.rows, sheet: sh }); return; }
      } catch { /* offline */ }
    }
    setPending(null);
  };
  useEffect(() => {
    if (!book || readOnly || !book.sheets.some((s) => s.form)) return undefined;
    checkPending();
    const iv = setInterval(() => { if (document.visibilityState === 'visible') checkPending(); }, 30000);
    return () => clearInterval(iv);
  }, [book]); // eslint-disable-line react-hooks/exhaustive-deps
  const mergePending = async () => {
    const p = pending; if (!p) return;
    const sh = p.sheet;
    const tb = (sh.tables || [])[0];
    const g = tb ? tb.range : sheetList(book, sh).g || { r1: 0, c1: 0, r2: 0, c2: 0 };
    const cols = (sh.form.columns && sh.form.columns.length ? sh.form.columns : null);
    book.run(t('Form answers'), () => {
      let r = g.r2 + 1;
      for (const row of p.rows) { row.forEach((v, j) => { const c = cols ? cols[j] : g.c1 + j; if (c != null && v != null && v !== '') book.input(sh, r, c, String(v)); }); r++; }
      if (tb) { book.touch(sh); sh.tables = sh.tables.map((x) => (x === tb ? { ...x, range: { ...x.range, r2: r - 1 }, _dirty: true } : x)); book.tx.structural = true; }
    });
    try { await fetch(`/api/office/forms/${encodeURIComponent(p.id)}/pending/ack`, { method: 'POST', credentials: 'same-origin', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ count: p.rows.length }) }); } catch { /* next time */ }
    setPending(null);
    toast.success(t('{n} answers added', { n: p.rows.length }));
  };

  /* ---------- co-editing ---------- */
  const toggleCollab = async () => {
    if (collab.current) { collab.current.stop(); collab.current = null; setCollabOn(false); setPeers([]); return; }
    if (!H.collab) { toast(t('Co-editing needs a newer Hostwerk Office.')); return; }
    try {
      if (!(await H.collab.available())) { toast(t('Co-editing is not available on this server.')); return; }
      collab.current = await startCollab(H, book, { onPeers: (list) => { setPeers(list.map((p, i) => ({ ...p, color: p.color || COLORS[i % COLORS.length] }))); force((n) => n + 1); }, onStatus: () => force((n) => n + 1) });
      setCollabOn(true);
      toast.success(t('Co-editing is on – everyone who opens this file edits together.'));
    } catch (e) { toast.error(errMsg(e)); }
  };
  useEffect(() => { if (collab.current) collab.current.setSelection({ sheet: sheet && sheet.name, r: sel.r, c: sel.c, ranges: sel.ranges.slice(0, 4) }); }, [sel, sheet]);
  useEffect(() => () => { if (collab.current) collab.current.stop(); }, []);
  // join automatically when somebody else edits this file right now
  useEffect(() => {
    if (!book || readOnly || !H.collab || !H.collab.available) return;
    const f = H.file; if (!f || !f.path) return;
    fetch(`/api/office/collab/state?${new URLSearchParams({ path: f.path, ...(f.root ? { root: f.root } : {}) })}`, { credentials: 'same-origin' }).then((r) => (r.ok ? r.json() : null)).then((j) => { if (j && (Array.isArray(j.peers) ? j.peers.length : Number(j.peers) || 0) > 0 && !collab.current) toggleCollab(); }).catch(() => {});
  }, [book]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- AI ---------- */
  const ai = {
    explain: async (f) => { const r = await H.ai.ask({ task: 'explain-formula', prompt: `Explain this spreadsheet formula step by step in ${loc.lang === 'de' ? 'German' : 'English'}, short and friendly:\n=${f}`, text: `=${f}` }); return r.text; },
    write: async (desc) => {
      const head = []; const u = book.dataRange(sheet);
      if (u) for (let c = 0; c <= Math.min(u.c2, 25); c++) head.push(`${colName(c)}: ${book.text(sheet, 0, c)}`);
      const r = await H.ai.ask({ task: 'write-formula', json: true, system: 'You write Excel formulas. Answer JSON {"formula": "..."} with an English-name Excel formula (comma separators, no leading =).', prompt: `Columns of the sheet (row 1 has the headers, data from row 2 to ${u ? u.r2 + 1 : 2}):\n${head.join('\n')}\nThe formula goes into ${colName(sel.c)}${sel.r + 1}. Task: ${desc}` });
      const f = String((r.json && r.json.formula) || (/=?\s*([A-Z][\s\S]*)/.exec(r.text || '') || [])[1] || '').replace(/^=/, '').trim();
      if (!f) throw new Error(t('The AI did not return a formula.'));
      return { formula: f, display: toDisplay(f, loc) };
    },
    clean: async (instr, sample) => {
      const r = await H.ai.ask({ task: 'clean-data', json: true, system: 'You clean spreadsheet data. Answer JSON {"rows": [[...], ...]} – one output row per input value, columns as asked.', prompt: `Instruction: ${instr || 'clean up'}\nValues (one per line):\n${sample.join('\n')}` });
      const rows = r.json && Array.isArray(r.json.rows) ? r.json.rows : null;
      if (!rows) throw new Error(t('The AI did not return rows.'));
      return { values: rows.map((x) => (Array.isArray(x) ? x : [x])) };
    },
  };
  const aiReady = !!(H.ai && (!H.ai.ready || H.ai.ready()));

  /* ---------- commands ---------- */
  const run = (cmd, arg) => {
    if (!book || !sheet) return false;
    const g = clipRange(activeRange());
    const ro = () => { if (readOnly) { guard(); return true; } return false; };
    switch (cmd) {
      case 'table': {
        if (ro()) return true;
        if (curTable) { setTableProps(book, sheet, curTable.id, { style: arg || 'TableStyleMedium2' }, 'Table style'); return true; }
        let rg = g;
        if (g.r1 === g.r2 && g.c1 === g.c2) rg = book.dataRange(sheet) || g;
        const hasHeader = [...Array(rg.c2 - rg.c1 + 1).keys()].every((j) => typeof book.value(sheet, rg.r1, rg.c1 + j) === 'string');
        const tb = addTable(book, sheet, rg, { header: hasHeader, style: arg || 'TableStyleMedium2', base: loc.lang === 'de' ? 'Tabelle' : 'Table' });
        if (!tb) toast.error(t('A table can’t overlap another table.'));
        return true;
      }
      case 'tableStyle': if (!ro() && curTable) setTableProps(book, sheet, curTable.id, { style: arg }, 'Table style'); return true;
      case 'tableOpt': if (!ro() && curTable) { if (arg === 'totals') toggleTotals(book, sheet, curTable.id, !curTable.totals); else setTableProps(book, sheet, curTable.id, { [arg]: !curTable[arg] }, 'Table'); } return true;
      case 'tableToRange': if (!ro() && curTable) convertToRange(book, sheet, curTable.id); return true;
      case 'tableName': if (!ro() && curTable) promptDialog({ title: t('Table name'), value: curTable.name, confirm: t('Rename') }).then((n) => { if (n && /^[A-Za-z_\\À-￿][\w.À-￿]*$/.test(n)) { const old = curTable.name; setTableProps(book, sheet, curTable.id, { name: n }, 'Rename table'); book.run(t('Rename table'), () => book.rewriteAll((f) => f.replace(new RegExp(`\\b${old}\\[`, 'g'), `${n}[`))); } else if (n) toast.error(t('This name can’t be used.')); }); return true;
      case 'tableStylesMenu': setDialog({ kind: 'plus:tstyle' }); return true;
      case 'pivotNew': if (!ro()) newPivot(); return true;
      case 'pivotEdit': if (curPivot) setDialog({ kind: 'plus:pivot', id: curPivot.id, sheetId: sheet.id }); return true;
      case 'pivotRefresh': if (curPivot && !curPivot.readOnly) applyPivot(curPivot, { label: 'Refresh' }); else if (curPivot) toast(t('This pivot table uses features only Excel can refresh.')); return true;
      case 'pivotRefreshAll': refreshAll(); return true;
      case 'sparkline': if (!ro()) setDialog({ kind: 'plus:spark', g }); return true;
      case 'sparkClear': if (!ro()) book.run(t('Clear sparklines'), () => { book.touch(sheet); sheet.sparklines = (sheet.sparklines || []).map((sg) => ({ ...sg, items: sg.items.filter((it) => !(it.r >= g.r1 && it.r <= g.r2 && it.c >= g.c1 && it.c <= g.c2)) })).filter((sg) => sg.items.length); }); return true;
      case 'iconSet': if (!ro()) book.setSheetProps(sheet, { cf: [...sheet.cf, { id: `cf${Date.now().toString(36)}`, ranges: sel.ranges.map(clipRange), type: 'icons', set: arg || '3TrafficLights1', stops: [], priority: sheet.cf.length + 1 }] }, 'Conditional formatting'); return true;
      case 'image': if (!ro()) pickImage(); return true;
      case 'shape': if (!ro()) addShape(arg || 'rect'); return true;
      case 'objDelete': if (objSel) delObj(objSel.kind, objSel.id); return true;
      case 'outline': if (arg) outline(arg); return true;
      case 'liveData': insertLive(); return true;
      case 'liveRefresh': refreshLive(); return true;
      case 'dbOpen': if (!ro()) openDb(); return true;
      case 'dbRefresh': if (sheet.db) loadDb(sheet, sheet.db); return true;
      case 'dbWrite': if (sheet.db) writeDb(sheet); return true;
      case 'view': book.setSheetProps(sheet, { view: arg && arg !== 'grid' ? { ...(sheet.view || {}), type: arg } : null }, 'View'); return true;
      case 'form': { if (ro()) return true; const fg = formRange(); if (!fg) { toast(t('Put a header row with column names at the top first.')); return true; } setDialog({ kind: 'plus:form', g: fg }); return true; }
      case 'formOpen': if (H.forms && H.forms.open) H.forms.open(sheet.form ? { id: sheet.form.id } : {}); return true;
      case 'rules': if (H.rules && H.rules.open) H.rules.open({ sheet: sheet.name, range: rangeA1(g) }).catch((e) => toast.error(errMsg(e))); else toast(t('Rules need a newer Hostwerk Office.')); return true;
      case 'collab': toggleCollab(); return true;
      case 'compare': compareMenu(arg); return true;
      case 'ai': if (!aiReady) { toast(t('Set up an AI provider in Hostwerk first.')); return true; } setDialog({ kind: 'plus:ai', mode: arg }); return true;
      case 'calcOptions': setDialog({ kind: 'plus:calc' }); return true;
      case 'r1c1': setR1c1((x) => !x); return true;
      default: return false;
    }
  };
  const compareMenu = async (at) => {
    if (!H.compare) { toast(t('Compare needs a newer Hostwerk Office.')); return; }
    try {
      const f = H.file || {};
      const r = await fetch(`/api/office/versions?${new URLSearchParams({ path: f.path || '', ...(f.root ? { root: f.root } : {}) })}`, { credentials: 'same-origin' });
      const j = r.ok ? await r.json() : { versions: [] };
      const list = (j.versions || j || []).slice(0, 12);
      if (!list.length) { toast(t('There are no older versions of this file yet.')); return; }
      ctx.showMenu(at || { x: 200, y: 120 }, [{ header: t('Compare with …') }, ...list.map((v) => ({ label: `${new Date(v.time).toLocaleString(loc.lang === 'de' ? 'de-DE' : 'en-US')}${v.label ? ` · ${v.label}` : ''}`, keep: true, onClick: () => H.compare({ id: v.id, time: v.time, label: v.label }, 'current').catch((e) => toast.error(errMsg(e))) }))]);
    } catch (e) { toast.error(errMsg(e)); }
  };

  /* ---------- typing into tables (auto-expand, calculated columns) ---------- */
  const afterInput = (sh, r, c, text) => {
    if (!(sh.tables || []).length) return;
    const t0 = tableAt(sh, r, c);
    if (!t0) { autoExpand(book, sh, r, c); return; }
    if (typeof text === 'string' && text.startsWith('=')) {
      const g = t0.range; const top = g.r1 + (t0.header !== false ? 1 : 0); const bot = g.r2 - (t0.totals ? 1 : 0);
      if (r < top || r > bot) return;
      let empty = true; for (let y = top; y <= bot; y++) if (y !== r) { const x = sh.cells.get(K(y, c)); if (x && (x.v != null || x.f)) { empty = false; break; } }
      const cell = sh.cells.get(K(r, c));
      if (empty && cell && cell.f && /\[/.test(cell.f)) for (let y = top; y <= bot; y++) if (y !== r) book.setCell(sh, y, c, { ...(sh.cells.get(K(y, c)) || {}), f: cell.f, v: null });
    }
  };

  /* ---------- dialogs ---------- */
  const dialog = (d, close) => {
    switch (d.kind) {
      case 'plus:pivot': {
        const sh = book.sheetById(d.sheetId) || sheet; const pv = (sh.pivots || []).find((p) => p.id === d.id); if (!pv) return null;
        return <PivotPanel book={book} pv={pv} readOnly={readOnly || pv.readOnly} onClose={close} onApply={(n) => { if (applyPivot(n, { label: 'PivotTable', target: sh })) close(); }} onRefresh={() => applyPivot(pv, { label: 'Refresh', target: sh })}
          onDelete={async () => { if (!(await confirmDialog({ title: t('Delete this pivot table?'), confirm: t('Delete'), danger: true }))) return; book.run(t('Delete pivot table'), () => { book.touch(sh, true); if (pv.range) for (let r = pv.range.r1; r <= pv.range.r2; r++) for (let c = pv.range.c1; c <= pv.range.c2; c++) book.setCell(sh, r, c, null); sh.pivots = sh.pivots.filter((p) => p.id !== pv.id); book.tx.structural = true; }); close(); }} />;
      }
      case 'plus:spark': return <SparklineDialog sel={d.g} sheetName={sheet.name} onClose={close} onApply={(sg) => book.run(t('Insert sparklines'), () => { book.touch(sheet); sheet.sparklines = [...(sheet.sparklines || []).map((x) => ({ ...x, items: x.items.filter((it) => !sg.items.some((n) => n.r === it.r && n.c === it.c)) })).filter((x) => x.items.length), sg]; })} />;
      case 'plus:form': return <FormDialog book={book} sheet={sheet} g={d.g} onClose={close} onCreate={createForm} />;
      case 'plus:ai': {
        const cell = sheet.cells.get(K(sel.r, sel.c));
        const g = clipRange(activeRange());
        const sample = []; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 199); r++) { const v = book.text(sheet, r, g.c1); if (v !== '') sample.push(v); }
        return <AiDialog mode={d.mode} formula={cell && cell.f ? toDisplay(cell.f, loc) : null} sample={g.r1 === g.r2 && g.c1 === g.c2 ? [] : sample} ai={ai} onClose={close}
          onInsert={(o) => { if (!guard()) return; if (o.formula) book.run(t('AI formula'), () => book.input(sheet, sel.r, sel.c, `=${toDisplay(o.formula, loc)}`)); if (o.values) book.run(t('AI clean-up'), () => o.values.forEach((row, i) => row.forEach((v, j) => book.input(sheet, g.r1 + i, g.c1 + j, String(v ?? ''))))); }} />;
      }
      case 'plus:calc': return <CalcOptionsDialog book={book} r1c1={r1c1} onR1C1={setR1c1} onClose={close} onApply={({ iterate, date1904 }) => {
        book.run(t('Calculation options'), () => {
          book.touchBook(); book.iterate = iterate;
          if (date1904 !== book.date1904) { // keep the same dates: shift every date-formatted number
            const d = date1904 ? -1462 : 1462;
            for (const sh of book.sheets) for (const [k, cell] of sh.cells) if (!cell.f && typeof cell.v === 'number' && cell.s && cell.s.nf && /[dmy]/i.test(cell.s.nf.replace(/"[^"]*"/g, ''))) book.setCell(sh, Math.floor(k / MAXC), k % MAXC, { ...cell, v: cell.v + d });
            book.date1904 = date1904;
          }
          book.tx.structural = true;
        });
      }} />;
      case 'plus:tstyle': return <TableStyleModal value={curTable && curTable.style} look={(id) => tableLook(id, book.pkg && book.pkg.theme)} onPick={(id) => { run('table', id); close(); }} onClose={close} />;
      default: return null;
    }
  };

  /* ---------- overlay: pictures, shapes, slicers ---------- */
  const overlay = (objRect, z) => {
    if (!sheet) return null;
    const slicerItems = (s) => {
      const c = s.cache; if (!c) return [];
      let src = null; let field = c.field;
      const pv = book.sheets.flatMap((x) => x.pivots || []).find((p) => p.cacheId === c.pivotCacheId);
      if (pv) src = pivotSource(book, pv);
      if (!src || src.error) return [];
      const i = src.fields.findIndex((f) => f.toLowerCase() === String(field).toLowerCase());
      if (i < 0) return [];
      const ex = new Set([...(pv.filters || []).filter((f) => f.field === src.fields[i]).flatMap((f) => f.exclude || []), ...((pv.itemFilters && pv.itemFilters[src.fields[i]]) || [])].map(String));
      return [...new Set(src.records.map((r) => r[i]))].sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true })).map((v) => ({ v, on: !ex.has(String(v)) }));
    };
    return <>
      {(sheet.shapes || []).filter((s) => !s._ctl).map((s) => { const rc = objRect(s.anchor); return (
        <ObjBox key={s.id} obj={s} rect={rc} z={z} selected={objSel && objSel.id === s.id} readOnly={readOnly} label={t('Shape')}
          onSelect={() => setObjSel({ kind: 'shape', id: s.id })} onChange={(a) => updObj('shape', s.id, { anchor: ctx.normAnchor(a) })}
          onOpen={async () => { if (readOnly) return; const v = await promptDialog({ title: t('Text of the shape'), value: s.text || '', multiline: true, confirm: t('OK') }).catch(() => null); if (v != null) updObj('shape', s.id, { text: v }, 'Shape text'); }}
          onMenu={(at) => ctx.showMenu(at, [{ label: t('Edit text …'), onClick: async () => { const v = await promptDialog({ title: t('Text of the shape'), value: s.text || '', confirm: t('OK') }).catch(() => null); if (v != null) updObj('shape', s.id, { text: v }, 'Shape text'); } }, { label: t('Fill colour …'), onClick: () => ctx.pickColor && ctx.pickColor(at, (c) => updObj('shape', s.id, { fill: c }, 'Shape fill')) }, { label: t('Line colour …'), onClick: () => ctx.pickColor && ctx.pickColor(at, (c) => updObj('shape', s.id, { line: c }, 'Shape line')) }, '-', { label: t('Bring to front'), onClick: () => orderObj('shape', s.id, true) }, { label: t('Send to back'), onClick: () => orderObj('shape', s.id, false) }, '-', { label: t('Delete'), danger: true, kbd: 'Del', onClick: () => delObj('shape', s.id) }])}>
          <ShapeSVG shape={s} w={rc.width} h={rc.height} z={z} />
        </ObjBox>); })}
      {sheet.images.map((im) => { const rc = objRect(im.anchor); const editable = im._orig !== true; return (
        <ObjBox key={im.id} obj={im} rect={rc} z={z} selected={objSel && objSel.id === im.id} readOnly={readOnly || !editable} label={t('Picture')} className="is-image"
          onSelect={() => setObjSel({ kind: 'image', id: im.id })} onChange={(a) => updObj('image', im.id, { anchor: ctx.normAnchor(a) })}
          onMenu={(at) => ctx.showMenu(at, [{ label: t('Bring to front'), onClick: () => orderObj('image', im.id, true) }, { label: t('Send to back'), onClick: () => orderObj('image', im.id, false) }, { label: t('Reset size'), onClick: () => { const img = new Image(); img.onload = () => updObj('image', im.id, { anchor: { ...im.anchor, w: img.width, h: img.height } }, 'Size'); img.src = im.src; } }, '-', { label: t('Delete'), danger: true, kbd: 'Del', disabled: !editable, onClick: () => delObj('image', im.id) }])}>
          <img src={im.src} alt={im.alt || ''} draggable={false} />
        </ObjBox>); })}
      {(sheet.controls || []).map((c) => { const st = controlState(book, sheet, c); return <ControlBox key={c.id} ctl={st} rect={objRect(c.anchor)} z={z} readOnly={readOnly} items={c.type === 'dropdown' || c.type === 'list' ? controlItems(book, sheet, c) : []} onAct={(a) => { if (guard && !guard()) return; applyControl(book, sheet, c.id, a, t('Form control')); }} />; })}
      {(sheet.slicers || []).map((s) => <div key={s.id} className="of-sheet-obj is-slicer" style={objRect(s.anchor)} title={t('Slicer – filter with it in Excel')}><SlicerBox slicer={s} items={slicerItems(s)} /></div>)}
    </>;
  };

  /* ---------- bars above the grid ---------- */
  const bars = () => {
    if (!sheet) return null;
    const out = [];
    if (pending) out.push(<div key="form" className="of-sheet-strip small is-info"><ClipboardList size={14} /><span className="grow">{t('{n} new form answers for “{sheet}”', { n: pending.rows.length, sheet: pending.sheet.name })}</span><button type="button" className="btn primary xs" onClick={mergePending}>{t('Add them')}</button><button type="button" className="btn ghost xs icon" aria-label={t('Later')} title={t('Later')} onClick={() => setPending(null)}><X size={12} /></button></div>);
    if (sheet.db) {
      const d = sheet.db; const st = dbState[sheet.id] || {};
      const n = dbChanges(sheet).length;
      out.push(<div key="db" className="of-sheet-strip small"><Database size={14} />
        <b><Keep>{`${d.db ? `${d.db}.` : ''}${d.table}`}</Keep></b>
        <span className="muted">{t('Rows {a}–{b} of {n}', { a: (d.offset || 0) + 1, b: Math.min(d.total || 0, (d.offset || 0) + (d.limit || 200)), n: d.total ?? '?' })}</span>
        {!d.primaryKey && <span className="pill small">{t('read-only (no primary key)')}</span>}
        <span className="grow" />
        <button type="button" className="btn ghost xs icon" aria-label={t('Previous page')} title={t('Previous page')} disabled={st.busy || !d.offset} onClick={() => loadDb(sheet, d, { offset: Math.max(0, (d.offset || 0) - (d.limit || 200)) })}><ChevronLeft size={13} /></button>
        <button type="button" className="btn ghost xs icon" aria-label={t('Next page')} title={t('Next page')} disabled={st.busy || (d.total != null && (d.offset || 0) + (d.limit || 200) >= d.total)} onClick={() => loadDb(sheet, d, { offset: (d.offset || 0) + (d.limit || 200) })}><ChevronRight size={13} /></button>
        <select className="input xs" aria-label={t('Sort by')} value={d.sort || ''} onChange={(e) => loadDb(sheet, { ...d, sort: e.target.value || null }, { offset: 0 })}><option value="">{t('Sort: as stored')}</option>{(d.columns || []).flatMap((c) => [<option key={`${c}+`} value={c}>{`${c} ↑`}</option>, <option key={`${c}-`} value={`-${c}`}>{`${c} ↓`}</option>])}</select>
        <input className="input xs" placeholder={t('Filter, e.g. name=Anna')} defaultValue={d.filter || ''} aria-label={t('Filter')} onKeyDown={(e) => { if (e.key === 'Enter') loadDb(sheet, { ...d, filter: e.currentTarget.value || null }, { offset: 0 }); }} />
        <button type="button" className="btn ghost xs" disabled={st.busy} onClick={() => loadDb(sheet, d)}><RefreshCw size={12} /> {t('Reload')}</button>
        {d.primaryKey && <button type="button" className="btn primary xs" disabled={st.busy || !n || readOnly} onClick={() => writeDb(sheet)}><Upload size={12} /> {n ? t('Write {n} changes …', { n }) : t('No changes')}</button>}
      </div>);
    }
    if ((sheet.live || []).length) out.push(<div key="live" className="of-sheet-strip small"><Radio size={14} className="of-sheet-live-dot" /><span className="grow">{t('Live data: {list}', { list: sheet.live.map((x) => x.label).join(' · ') })}</span><button type="button" className="btn ghost xs" onClick={refreshLive}><RefreshCw size={12} /> {t('Refresh')}</button></div>);
    if (collabOn) out.push(<div key="collab" className="of-sheet-strip small"><Users size={14} /><span className="grow">{peers.length ? t('Editing together with {names}', { names: peers.map((p) => p.name).join(', ') }) : t('Co-editing is on – nobody else is here yet.')}</span><button type="button" className="btn ghost xs" onClick={toggleCollab}>{t('Stop co-editing')}</button></div>);
    return out.length ? out : null;
  };

  /* ---------- the non-grid views ---------- */
  const viewType = sheet && sheet.view && sheet.view.type;
  const openRow = (r) => { book.setSheetProps(sheet, { view: { ...(sheet.view || {}), type: 'grid', last: sheet.view && sheet.view.type } }, 'View'); setSel({ r, c: 0, ranges: [{ r1: r, c1: 0, r2: r, c2: 0 }], anchor: { r, c: 0 } }); setTimeout(focusGrid, 0); };
  const viewBar = sheet ? <ViewBar book={book} sheet={sheet} view={sheet.view} readOnly={readOnly} onChange={(v) => book.setSheetProps(sheet, { view: v }, 'View')} /> : null;
  const viewBody = !viewType || viewType === 'grid' ? null : viewType === 'kanban' ? <KanbanView book={book} sheet={sheet} view={sheet.view} readOnly={readOnly} onOpen={openRow} />
    : viewType === 'calendar' ? <CalendarView book={book} sheet={sheet} view={sheet.view} readOnly={readOnly} onOpen={openRow} loc={loc} /> : <GalleryView book={book} sheet={sheet} view={sheet.view} onOpen={openRow} />;

  /* ---------- R1C1 ---------- */
  const editText = (sh, r, c, toDisp) => { const txt = book.editText(sh, r, c, toDisp); return r1c1 && txt.startsWith('=') ? `=${toR1C1(txt.slice(1), r, c)}` : txt; };
  const inputText = (text, r, c) => (r1c1 && typeof text === 'string' && text.startsWith('=') ? `=${fromR1C1(text.slice(1), r, c)}` : text);

  const remote = collabOn ? peers.filter((p) => p.sel).map((p) => ({ ...p.sel, name: p.name, color: p.color })) : null;
  return {
    run, dialog, overlay, bars, viewBar, viewBody, viewType, objSel, setObjSel, delObj, r1c1, editText, inputText, afterInput, remote,
    curTable, curPivot, aiReady, collabOn, peers, outline, addImageFile, follower: () => !!(collab.current && !collab.current.leader),
    flags: { inTable: !!curTable, table: curTable, inPivot: !!curPivot, pivot: curPivot, objSel, r1c1, collabOn, aiReady, hasDb: !!(sheet && sheet.db), hasLive: !!(sheet && (sheet.live || []).length), hasForm: !!(sheet && sheet.form), view: viewType || 'grid' },
  };
}

function TableStyleModal({ value, look, onPick, onClose }) {
  return <Modal title={t('Table styles')} onClose={onClose}><TableStyleGrid value={value} look={look} onPick={onPick} /></Modal>;
}
