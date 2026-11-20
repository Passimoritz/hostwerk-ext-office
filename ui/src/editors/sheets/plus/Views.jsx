/*
 * Views of a sheet (saved per sheet in the Hostwerk sidecar part – Excel just shows the cells):
 * Kanban (lanes = values of a column, dragging a card writes the cell), Calendar (items on the dates of a date column,
 * drag to another day), Gallery (cards with a picture column). Rows = the table at A1 / the data region below the header row.
 */
import { useMemo, useState } from 'react';
import { t } from '@hostwerk/i18n';
import { ChevronLeft, ChevronRight, Plus, Image as ImageIcon, GripVertical, Table2, Columns3, CalendarDays, LayoutGrid } from 'lucide-react';
import { Keep, Raw } from '../../_shared/keep.jsx';
import { MAXC } from '../formula/refs.js';
import { partsFromSerial, serialFromParts } from '../format/dates.js';
import { isDateFormat } from '../format/numfmt.js';

const K = (r, c) => r * MAXC + c;
export const VIEW_TYPES = [['grid', 'Grid', Table2], ['kanban', 'Kanban', Columns3], ['calendar', 'Calendar', CalendarDays], ['gallery', 'Gallery', LayoutGrid]];

/** header + data rows of the sheet's list (the table containing A1, else the first table, else the used range) */
export function sheetList(book, sheet) {
  let g = null;
  const t0 = (sheet.tables || [])[0];
  if (t0) g = { ...t0.range, r2: t0.totals ? t0.range.r2 - 1 : t0.range.r2 };
  else { const u = book.dataRange(sheet); if (u) g = { r1: 0, c1: 0, r2: u.r2, c2: u.c2 }; }
  if (!g) return { g: null, cols: [], rows: [] };
  const cols = [];
  for (let c = g.c1; c <= g.c2; c++) cols.push({ c, name: book.text(sheet, g.r1, c) || t('Column {n}', { n: c - g.c1 + 1 }) });
  const rows = [];
  for (let r = g.r1 + 1; r <= g.r2; r++) {
    let any = false;
    for (let c = g.c1; c <= g.c2; c++) { const v = book.value(sheet, r, c); if (v != null && v !== '') { any = true; break; } }
    if (any) rows.push(r);
  }
  return { g, cols, rows };
}
const isDateCol = (book, sheet, list, c) => list.rows.slice(0, 20).some((r) => { const cell = sheet.cells.get(K(r, c)); return cell && typeof cell.v === 'number' && cell.s && cell.s.nf && isDateFormat(cell.s.nf); });

export function ViewBar({ book, sheet, view, onChange, readOnly }) {
  const list = useMemo(() => sheetList(book, sheet), [book, sheet, book.version]); // eslint-disable-line react-hooks/exhaustive-deps
  const type = (view && view.type) || 'grid';
  const set = (patch) => onChange({ ...(view || {}), ...patch });
  const colSel = (key, label, filter) => (
    <label className="row small of-sheet-vb-field">{label}
      <select className="input sm" value={(view && view[type] && view[type][key]) ?? ''} disabled={readOnly} onChange={(e) => set({ [type]: { ...((view && view[type]) || {}), [key]: e.target.value === '' ? null : +e.target.value } })}>
        <option value="">{t('Choose …')}</option>
        {list.cols.filter((x) => !filter || filter(x)).map((x) => <Raw as="option" key={x.c} value={x.c}>{x.name}</Raw>)}
      </select>
    </label>
  );
  return (
    <div className="of-sheet-viewbar">
      <div className="of-sheet-seg" role="tablist" aria-label={t('View')}>
        {VIEW_TYPES.map(([id, label, I]) => <button key={id} type="button" role="tab" aria-selected={type === id} className={type === id ? 'is-on' : ''} title={t(label)} onClick={() => onChange(id === 'grid' ? null : { ...(view || {}), type: id })}><I size={14} /><span>{t(label)}</span></button>)}
      </div>
      {type === 'kanban' && colSel('col', t('Lanes from'))}
      {type === 'kanban' && colSel('title', t('Card title'))}
      {type === 'calendar' && colSel('col', t('Date column'), (x) => isDateCol(book, sheet, list, x.c))}
      {type === 'calendar' && colSel('title', t('Title'))}
      {type === 'gallery' && colSel('image', t('Picture column'))}
      {type === 'gallery' && colSel('title', t('Title'))}
      <span className="grow" />
      {type !== 'grid' && <span className="small muted">{t('{n} rows', { n: list.rows.length })}</span>}
    </div>
  );
}

function Card({ book, sheet, list, r, title, onOpen, drag, extra }) {
  const tc = title != null ? title : list.cols[0] && list.cols[0].c;
  const head = tc != null ? book.text(sheet, r, tc) : '';
  const fields = list.cols.filter((x) => x.c !== tc).slice(0, 4).map((x) => [x.name, book.text(sheet, r, x.c)]).filter(([, v]) => v);
  return (
    <div className="of-sheet-card" draggable={!!drag} onDragStart={drag ? (e) => { e.dataTransfer.setData('text/x-of-row', String(r)); e.dataTransfer.effectAllowed = 'move'; } : undefined}
      onDoubleClick={() => onOpen(r)} onKeyDown={(e) => { if (e.key === 'Enter') onOpen(r); }} tabIndex={0} role="button" aria-label={t('Row {n}', { n: r + 1 })}>
      {extra}
      <div className="of-sheet-card-head">{drag && <GripVertical size={13} className="muted" />}<b className="grow"><Keep>{head || t('(no title)')}</Keep></b><span className="small muted">{r + 1}</span></div>
      {fields.map(([k, v]) => <div key={k} className="small of-sheet-card-f"><span className="muted"><Keep>{k}</Keep></span> <Keep>{v}</Keep></div>)}
    </div>
  );
}

export function KanbanView({ book, sheet, view, readOnly, onOpen }) {
  const list = useMemo(() => sheetList(book, sheet), [book, sheet, book.version]); // eslint-disable-line react-hooks/exhaustive-deps
  const cfg = (view && view.kanban) || {};
  const gc = cfg.col != null ? cfg.col : list.cols[1] ? list.cols[1].c : null;
  const [over, setOver] = useState(null);
  if (!list.g || gc == null) return <Empty text={t('Put a header row with column names at the top, then pick the column for the lanes.')} />;
  const lanes = new Map();
  for (const r of list.rows) { const k = book.text(sheet, r, gc) || ''; if (!lanes.has(k)) lanes.set(k, []); lanes.get(k).push(r); }
  // values of a list validation on the lane column come first (empty lanes too)
  const dv = sheet.dv.find((d) => d.type === 'list' && d.ranges.some((g) => gc >= g.c1 && gc <= g.c2) && /^".*"$/.test(d.f1 || ''));
  const order = dv ? String(dv.f1).slice(1, -1).split(/[,;]/).map((x) => x.trim()) : [];
  for (const k of order) if (!lanes.has(k)) lanes.set(k, []);
  const keys = [...lanes.keys()].sort((a, b) => { const ia = order.indexOf(a); const ib = order.indexOf(b); if (ia >= 0 || ib >= 0) return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib); if (!a) return 1; if (!b) return -1; return a.localeCompare(b, undefined, { numeric: true }); });
  const drop = (lane, e) => {
    e.preventDefault(); setOver(null);
    const r = +e.dataTransfer.getData('text/x-of-row');
    if (readOnly || !Number.isFinite(r) || book.text(sheet, r, gc) === lane) return;
    book.run(t('Move card'), () => book.input(sheet, r, gc, lane));
  };
  const add = (lane) => {
    const r = list.rows.length ? Math.max(...list.rows) + 1 : list.g.r1 + 1;
    book.run(t('New card'), () => { book.input(sheet, r, gc, lane); if (cfg.title != null && cfg.title !== gc) book.input(sheet, r, cfg.title, t('New')); else if (list.cols[0].c !== gc) book.input(sheet, r, list.cols[0].c, t('New')); });
    onOpen(r);
  };
  return (
    <div className="of-sheet-kanban">
      {keys.map((k) => (
        <section key={k || '\u0000'} className={`of-sheet-lane${over === k ? ' is-over' : ''}`} onDragOver={(e) => { e.preventDefault(); setOver(k); }} onDragLeave={() => setOver(null)} onDrop={(e) => drop(k, e)} aria-label={k || t('(empty)')}>
          <header className="row"><b className="grow">{k ? <Keep>{k}</Keep> : <span className="muted">{t('(empty)')}</span>}</b><span className="pill small">{lanes.get(k).length}</span>
            {!readOnly && <button type="button" className="btn ghost xs icon" aria-label={t('New card')} title={t('New card')} onClick={() => add(k)}><Plus size={13} /></button>}</header>
          <div className="of-sheet-lane-body">{lanes.get(k).map((r) => <Card key={r} book={book} sheet={sheet} list={list} r={r} title={cfg.title} onOpen={onOpen} drag={!readOnly} />)}</div>
        </section>
      ))}
    </div>
  );
}

export function CalendarView({ book, sheet, view, readOnly, onOpen, loc }) {
  const list = useMemo(() => sheetList(book, sheet), [book, sheet, book.version]); // eslint-disable-line react-hooks/exhaustive-deps
  const cfg = (view && view.calendar) || {};
  const dc = cfg.col != null ? cfg.col : (list.cols.find((x) => isDateCol(book, sheet, list, x.c)) || {}).c;
  const items = new Map();
  let first = null;
  if (dc != null) for (const r of list.rows) { const v = book.value(sheet, r, dc); if (typeof v !== 'number') continue; const d = Math.floor(v); if (first == null || d < first) first = d; if (!items.has(d)) items.set(d, []); items.get(d).push(r); }
  const today = Math.floor(serialFromParts(new Date().getFullYear(), new Date().getMonth() + 1, new Date().getDate()));
  const [month, setMonth] = useState(() => { const p = partsFromSerial(items.size ? (items.has(today) || [...items.keys()].some((d) => Math.abs(d - today) < 40) ? today : first) : today); return { y: p.y, m: p.m }; });
  if (dc == null) return <Empty text={t('This sheet has no date column. Format a column as a date, then pick it here.')} />;
  const start = serialFromParts(month.y, month.m, 1);
  const lead = (partsFromSerial(start).dow + 6) % 7; // Monday first
  const days = [];
  for (let i = 0; i < 42; i++) days.push(start - lead + i);
  const move = (d) => setMonth((m) => { let y = m.y; let mm = m.m + d; if (mm < 1) { mm = 12; y--; } if (mm > 12) { mm = 1; y++; } return { y, m: mm }; });
  const drop = (day, e) => {
    e.preventDefault();
    const r = +e.dataTransfer.getData('text/x-of-row'); if (readOnly || !Number.isFinite(r)) return;
    const old = book.value(sheet, r, dc); const frac = typeof old === 'number' ? old - Math.floor(old) : 0;
    const cell = sheet.cells.get(K(r, dc)) || {};
    book.run(t('Move to another day'), () => book.setCell(sheet, r, dc, { ...cell, f: undefined, v: day + frac }));
  };
  const tc = cfg.title != null ? cfg.title : (list.cols.find((x) => x.c !== dc) || {}).c;
  const names = (loc && loc.daysShort) || ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  return (
    <div className="of-sheet-cal">
      <div className="row of-sheet-cal-head">
        <button type="button" className="btn ghost sm icon" aria-label={t('Previous month')} title={t('Previous month')} onClick={() => move(-1)}><ChevronLeft size={15} /></button>
        <b className="of-sheet-cal-title"><Keep>{`${(loc && loc.months ? loc.months[month.m - 1] : month.m)} ${month.y}`}</Keep></b>
        <button type="button" className="btn ghost sm icon" aria-label={t('Next month')} title={t('Next month')} onClick={() => move(1)}><ChevronRight size={15} /></button>
        <button type="button" className="btn ghost sm" onClick={() => { const p = partsFromSerial(today); setMonth({ y: p.y, m: p.m }); }}>{t('Today')}</button>
      </div>
      <div className="of-sheet-cal-grid">
        {[1, 2, 3, 4, 5, 6, 0].map((d) => <div key={d} className="of-sheet-cal-dow small muted"><Keep>{names[d]}</Keep></div>)}
        {days.map((d) => { const p = partsFromSerial(d); const its = items.get(d) || []; return (
          <div key={d} className={`of-sheet-cal-day${p.m !== month.m ? ' is-out' : ''}${d === today ? ' is-today' : ''}`} onDragOver={(e) => e.preventDefault()} onDrop={(e) => drop(d, e)}>
            <span className="small of-sheet-cal-num">{p.d}</span>
            {its.slice(0, 4).map((r) => <button key={r} type="button" className="of-sheet-cal-item small" draggable={!readOnly} onDragStart={(e) => e.dataTransfer.setData('text/x-of-row', String(r))} onClick={() => onOpen(r)} title={t('Row {n}', { n: r + 1 })}><Keep>{tc != null ? book.text(sheet, r, tc) || '·' : '·'}</Keep></button>)}
            {its.length > 4 && <span className="small muted">{t('+{n} more', { n: its.length - 4 })}</span>}
          </div>); })}
      </div>
    </div>
  );
}

const imgOk = (s) => /^(data:image\/|https?:\/\/|\/api\/)/i.test(String(s || ''));
export function GalleryView({ book, sheet, view, onOpen }) {
  const list = useMemo(() => sheetList(book, sheet), [book, sheet, book.version]); // eslint-disable-line react-hooks/exhaustive-deps
  const cfg = (view && view.gallery) || {};
  const ic = cfg.image != null ? cfg.image : (list.cols.find((x) => list.rows.slice(0, 10).some((r) => imgOk(book.value(sheet, r, x.c)))) || {}).c;
  if (!list.g) return <Empty text={t('Put a header row with column names at the top – every row becomes a card.')} />;
  return (
    <div className="of-sheet-gallery">
      {list.rows.map((r) => {
        const src = ic != null ? book.value(sheet, r, ic) : null;
        return <Card key={r} book={book} sheet={sheet} list={{ ...list, cols: list.cols.filter((x) => x.c !== ic) }} r={r} title={cfg.title} onOpen={onOpen}
          extra={<div className="of-sheet-card-img">{imgOk(src) ? <img src={src} alt="" loading="lazy" /> : <ImageIcon size={28} className="muted" />}</div>} />;
      })}
    </div>
  );
}
function Empty({ text }) { return <div className="of-sheet-view-empty muted">{text}</div>; }
