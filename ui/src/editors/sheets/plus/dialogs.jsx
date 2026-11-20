/* Office 1.1 dialogs of the sheets editor: pivot table field list, sparklines, forms, AI, calculation options, table name. */
import { useMemo, useState } from 'react';
import { Modal, Btn, Field, Callout, Seg } from '@hostwerk/ui';
import * as HUI from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { ArrowUp, ArrowDown, X, Sparkles, Rows3, Columns3, Sigma, Filter as FilterIcon, RefreshCw } from 'lucide-react';
import { Keep, Raw } from '../../_shared/keep.jsx';
import { parseRange, rangeA1, colName } from '../formula/refs.js';
import { pivotSource, PIVOT_FNS, computePivot } from '../pivot.js';
import { isDateFormat } from '../format/numfmt.js';
import { MAXC } from '../formula/refs.js';

const Sheet = HUI.SideSheet || Modal;
const K = (r, c) => r * MAXC + c;
const FN_LABEL = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count numbers', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' };

/* ---------------- pivot table field list ---------------- */
export function PivotPanel({ book, pv: pv0, onApply, onClose, onRefresh, onDelete, readOnly }) {
  const [pv, setPv] = useState(() => JSON.parse(JSON.stringify(pv0)));
  const src = useMemo(() => pivotSource(book, pv), [book, pv.source]); // eslint-disable-line react-hooks/exhaustive-deps
  const used = new Set([...(pv.rows || []), ...(pv.cols || []), ...(pv.filters || []).map((f) => f.field), ...(pv.values || []).map((v) => v.field)]);
  const area = (k) => (k === 'filters' ? (pv.filters || []).map((f) => f.field) : k === 'values' ? (pv.values || []).map((v) => v.field) : pv[k] || []);
  const put = (field, k) => setPv((p) => {
    const n = { ...p, rows: (p.rows || []).filter((x) => x !== field), cols: (p.cols || []).filter((x) => x !== field), filters: (p.filters || []).filter((x) => x.field !== field) };
    if (k === 'values') { const num = src.records && src.records.some((r) => typeof r[src.fields.indexOf(field)] === 'number'); n.values = [...(p.values || []), { field, fn: num ? 'sum' : 'count' }]; } else n.values = (p.values || []).filter((v) => v.field !== field || k === 'none');
    if (k === 'rows') n.rows = [...n.rows, field];
    if (k === 'cols') n.cols = [...n.cols, field];
    if (k === 'filters') n.filters = [...n.filters, { field, exclude: [] }];
    if (k === 'none') n.values = (p.values || []).filter((v) => v.field !== field);
    return n;
  });
  const moveIn = (k, i, d) => setPv((p) => { const key = k === 'filters' ? 'filters' : k; const a = (p[key] || []).slice(); const j = i + d; if (j < 0 || j >= a.length) return p; [a[i], a[j]] = [a[j], a[i]]; return { ...p, [key]: a }; });
  const live = useMemo(() => (src.error ? null : computePivot(book, pv)), [book, pv, src]); // preview size
  const AREAS = [['filters', t('Filters'), FilterIcon], ['cols', t('Columns'), Columns3], ['rows', t('Rows'), Rows3], ['values', t('Values'), Sigma]];
  const drop = (k) => (e) => { e.preventDefault(); const f = e.dataTransfer.getData('text/x-of-field'); if (f) put(f, k); };
  return (
    <Sheet title={t('PivotTable fields')} onClose={onClose} footer={<>
      {onDelete && !readOnly && <Btn className="ghost" onClick={onDelete}>{t('Delete')}</Btn>}
      <div className="grow" />
      {onRefresh && <Btn className="ghost" onClick={onRefresh}><RefreshCw size={14} /> {t('Refresh')}</Btn>}
      <Btn className="primary" disabled={readOnly || !!src.error} onClick={() => onApply(pv)}>{t('Apply')}</Btn></>}>
      <div className="stack of-sheet-pvp">
        <Field label={t('Source data')}><input className="input mono" value={pv.source || ''} onChange={(e) => setPv({ ...pv, source: e.target.value })} disabled={readOnly} /></Field>
        {src.error && <Callout kind="warn">{t(src.error)}</Callout>}
        {!src.error && <>
          <div className="small muted">{t('Drag fields into the areas – or tick them (numbers go to Values).')}</div>
          <div className="of-sheet-pvp-fields" role="list">
            {src.fields.map((f) => <label key={f} className="row of-sheet-pvp-field" draggable onDragStart={(e) => e.dataTransfer.setData('text/x-of-field', f)} role="listitem">
              <input type="checkbox" checked={used.has(f)} disabled={readOnly} onChange={(e) => { if (!e.target.checked) { setPv((p) => ({ ...p, rows: (p.rows || []).filter((x) => x !== f), cols: (p.cols || []).filter((x) => x !== f), filters: (p.filters || []).filter((x) => x.field !== f), values: (p.values || []).filter((v) => v.field !== f) })); return; } const i = src.fields.indexOf(f); const num = src.records.some((r) => typeof r[i] === 'number'); put(f, num ? 'values' : 'rows'); }} />
              <Keep>{f}</Keep></label>)}
          </div>
          <div className="of-sheet-pvp-areas">
            {AREAS.map(([k, label, I]) => (
              <div key={k} className="of-sheet-pvp-area" onDragOver={(e) => e.preventDefault()} onDrop={drop(k)} aria-label={label}>
                <div className="small muted row"><I size={13} /> {label}</div>
                {area(k).map((f, i) => (
                  <div key={`${f}${i}`} className="of-sheet-pvp-chip row" draggable onDragStart={(e) => e.dataTransfer.setData('text/x-of-field', f)}>
                    <span className="grow"><Keep>{k === 'values' ? `${t(FN_LABEL[pv.values[i].fn || 'sum'])} · ${f}` : f}</Keep></span>
                    {k === 'values' && <select className="input xs" aria-label={t('Summarize by')} value={pv.values[i].fn || 'sum'} onChange={(e) => setPv((p) => ({ ...p, values: p.values.map((v, j) => (j === i ? { ...v, fn: e.target.value, name: undefined } : v)) }))}>{PIVOT_FNS.map((x) => <option key={x} value={x}>{t(FN_LABEL[x])}</option>)}</select>}
                    {k === 'filters' && <FilterItems src={src} f={pv.filters[i]} onChange={(ex) => setPv((p) => ({ ...p, filters: p.filters.map((x, j) => (j === i ? { ...x, exclude: ex } : x)) }))} />}
                    <button type="button" className="btn ghost xs icon" aria-label={t('Move up')} title={t('Move up')} onClick={() => moveIn(k, i, -1)}><ArrowUp size={12} /></button>
                    <button type="button" className="btn ghost xs icon" aria-label={t('Move down')} title={t('Move down')} onClick={() => moveIn(k, i, 1)}><ArrowDown size={12} /></button>
                    <button type="button" className="btn ghost xs icon" aria-label={t('Remove')} title={t('Remove')} onClick={() => setPv((p) => (k === 'values' ? { ...p, values: p.values.filter((_, j) => j !== i) } : k === 'filters' ? { ...p, filters: p.filters.filter((_, j) => j !== i) } : { ...p, [k]: p[k].filter((_, j) => j !== i) }))}><X size={12} /></button>
                  </div>
                ))}
              </div>
            ))}
          </div>
          <div className="row small">
            <label className="row"><input type="checkbox" checked={pv.grandRows !== false} onChange={(e) => setPv({ ...pv, grandRows: e.target.checked })} /> {t('Grand totals for rows')}</label>
            <label className="row"><input type="checkbox" checked={pv.grandCols !== false} onChange={(e) => setPv({ ...pv, grandCols: e.target.checked })} /> {t('… for columns')}</label>
          </div>
          {live && !live.error && <div className="small muted">{t('Result: {r} rows × {c} columns · {n} source rows', { r: live.loc.h, c: live.loc.w, n: src.records.length })}</div>}
        </>}
      </div>
    </Sheet>
  );
}
function FilterItems({ src, f, onChange }) {
  const i = src.fields.indexOf(f.field);
  const items = [...new Set(src.records.map((r) => r[i]))].sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
  const ex = new Set((f.exclude || []).map(String));
  return (
    <details className="of-sheet-pvp-filter"><summary className="small">{ex.size ? t('{n} hidden', { n: ex.size }) : t('All')}</summary>
      <div className="stack">{items.slice(0, 200).map((v) => <label key={String(v)} className="row small"><input type="checkbox" checked={!ex.has(String(v))} onChange={(e) => { const n = new Set(ex); if (e.target.checked) n.delete(String(v)); else n.add(String(v)); onChange(items.filter((x) => n.has(String(x)))); }} /><Keep>{v == null ? t('(blank)') : String(v)}</Keep></label>)}</div>
    </details>
  );
}

/* ---------------- sparklines ---------------- */
export function SparklineDialog({ sel, sheetName, onApply, onClose }) {
  const g = sel;
  const [type, setType] = useState('line');
  const [data, setData] = useState(rangeA1(g));
  const [loc, setLoc] = useState(() => (g.c2 - g.c1 >= g.r2 - g.r1 ? `${colName(g.c2 + 1)}${g.r1 + 1}:${colName(g.c2 + 1)}${g.r2 + 1}` : `${colName(g.c1)}${g.r2 + 2}:${colName(g.c2)}${g.r2 + 2}`));
  const [color, setColor] = useState('#376092');
  const [opts, setOpts] = useState({ markers: false, high: true, low: false, negative: true });
  const apply = () => {
    const d = parseRange(data); const l = parseRange(loc);
    if (!d || !l) return;
    const byRow = l.c1 === l.c2; // one sparkline per row of the data
    const items = [];
    const n = byRow ? l.r2 - l.r1 + 1 : l.c2 - l.c1 + 1;
    for (let i = 0; i < n; i++) {
      const src = byRow ? { r1: d.r1 + i, r2: d.r1 + i, c1: d.c1, c2: d.c2 } : { r1: d.r1, r2: d.r2, c1: d.c1 + i, c2: d.c1 + i };
      items.push({ ref: `${/^[A-Za-z_][\w.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`}!${rangeA1(src)}`, r: byRow ? l.r1 + i : l.r1, c: byRow ? l.c1 : l.c1 + i });
    }
    onApply({ id: `sp${Date.now().toString(36)}`, type, color, neg: '#D00000', markerColor: '#D00000', highColor: '#D00000', lowColor: '#D00000', weight: 0.75, empty: 'gap', ...opts, items });
    onClose();
  };
  return (
    <Modal title={t('Insert sparklines')} onClose={onClose} footer={<><div className="grow" /><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={apply}>{t('Insert')}</Btn></>}>
      <div className="stack">
        <Raw as={Seg} value={type} onChange={setType} options={[{ id: 'line', label: t('Line') }, { id: 'column', label: t('Column') }, { id: 'stacked', label: t('Win/loss') }]} />
        <Field label={t('Data range')}><input className="input mono" value={data} onChange={(e) => setData(e.target.value)} /></Field>
        <Field label={t('Location range')} hint={t('One cell per row (or column) of the data')}><input className="input mono" value={loc} onChange={(e) => setLoc(e.target.value)} /></Field>
        <div className="row"><label className="row small">{t('Colour')} <input type="color" value={color} onChange={(e) => setColor(e.target.value)} aria-label={t('Colour')} /></label>
          {[['high', 'High point'], ['low', 'Low point'], ['negative', 'Negative points'], ['markers', 'Markers']].map(([k, l]) => <label key={k} className="row small"><input type="checkbox" checked={!!opts[k]} onChange={(e) => setOpts({ ...opts, [k]: e.target.checked })} /> {t(l)}</label>)}</div>
      </div>
    </Modal>
  );
}

/* ---------------- forms ---------------- */
/** field types from the header row: data validation, number formats, names */
export function formFields(book, sheet, g) {
  const out = [];
  for (let c = g.c1; c <= g.c2; c++) {
    const label = book.text(sheet, g.r1, c).trim(); if (!label) continue;
    const dv = sheet.dv.find((d) => d.ranges.some((x) => c >= x.c1 && c <= x.c2 && x.r2 > g.r1));
    let type = 'text'; let options;
    const sample = sheet.cells.get(K(g.r1 + 1, c));
    const nf = sample && sample.s && sample.s.nf;
    if (dv && dv.type === 'list' && /^".*"$/.test(dv.f1 || '')) { type = 'choice'; options = dv.f1.slice(1, -1).split(/[,;]/).map((x) => x.trim()).filter(Boolean); } else if (dv && (dv.type === 'whole' || dv.type === 'decimal')) type = 'number';
    else if (dv && dv.type === 'date') type = 'date';
    else if (nf && isDateFormat(nf)) type = 'date';
    else if (sample && typeof sample.v === 'number') type = 'number';
    else if (sample && typeof sample.v === 'boolean') type = 'checkbox';
    else if (/e-?mail/i.test(label)) type = 'email';
    else if (/phone|telefon|tel\b/i.test(label)) type = 'tel';
    else if (/url|link|web/i.test(label)) type = 'url';
    out.push({ c, label, type, required: false, ...(options ? { options } : {}) });
  }
  return out;
}
export function FormDialog({ book, sheet, g, title0, onCreate, onClose }) {
  const [fields, setFields] = useState(() => formFields(book, sheet, g));
  const [title, setTitle] = useState(title0 || sheet.name);
  const [pub, setPub] = useState({ enabled: true, expires: '', password: '' });
  const [busy, setBusy] = useState(false);
  const TYPES = [['text', 'Text'], ['textarea', 'Long text'], ['number', 'Number'], ['date', 'Date'], ['email', 'E-mail'], ['tel', 'Phone'], ['url', 'Link'], ['choice', 'Choice'], ['checkbox', 'Yes/no']];
  const upd = (i, patch) => setFields((f) => f.map((x, j) => (j === i ? { ...x, ...patch } : x)));
  return (
    <Modal title={t('Form from this sheet')} onClose={onClose} footer={<><div className="grow" /><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn>
      <Btn className="primary" disabled={busy || !fields.length} onClick={async () => { setBusy(true); try { await onCreate({ title, fields, public: { enabled: pub.enabled, expires: pub.expires || null, password: pub.password || null } }); onClose(); } finally { setBusy(false); } }}>{t('Publish form')}</Btn></>}>
      <div className="stack">
        <span className="small muted">{t('Every answer becomes a new row below the list. The fields come from the header row; types from validation and number formats.')}</span>
        <Field label={t('Title')}><input className="input" value={title} onChange={(e) => setTitle(e.target.value)} /></Field>
        {!fields.length && <Callout kind="warn">{t('The first row of the selection needs column names.')}</Callout>}
        <div className="of-sheet-formfields">
          {fields.map((f, i) => (
            <div key={f.c} className="row">
              <span className="mono small muted">{colName(f.c)}</span>
              <input className="input sm grow" value={f.label} aria-label={t('Field name')} onChange={(e) => upd(i, { label: e.target.value })} />
              <select className="input sm" value={f.type} aria-label={t('Type')} onChange={(e) => upd(i, { type: e.target.value })}>{TYPES.map(([v, l]) => <option key={v} value={v}>{t(l)}</option>)}</select>
              {f.type === 'choice' && <input className="input sm" placeholder={t('Options; separated')} value={(f.options || []).join('; ')} onChange={(e) => upd(i, { options: e.target.value.split(';').map((x) => x.trim()).filter(Boolean) })} />}
              <label className="row small"><input type="checkbox" checked={f.required} onChange={(e) => upd(i, { required: e.target.checked })} /> {t('Required')}</label>
              <button type="button" className="btn ghost xs icon" aria-label={t('Remove')} title={t('Remove')} onClick={() => setFields((x) => x.filter((_, j) => j !== i))}><X size={12} /></button>
            </div>
          ))}
        </div>
        <label className="row"><input type="checkbox" checked={pub.enabled} onChange={(e) => setPub({ ...pub, enabled: e.target.checked })} /> {t('Public link (anyone with the link can answer)')}</label>
        {pub.enabled && <div className="row"><Field label={t('Expires')}><input type="date" className="input sm" value={pub.expires} onChange={(e) => setPub({ ...pub, expires: e.target.value })} /></Field>
          <Field label={t('Password (optional)')}><input type="password" className="input sm" autoComplete="new-password" value={pub.password} onChange={(e) => setPub({ ...pub, password: e.target.value })} /></Field></div>}
      </div>
    </Modal>
  );
}

/* ---------------- AI ---------------- */
export function AiDialog({ mode: mode0, formula, sample, onInsert, onClose, ai }) {
  const [mode, setMode] = useState(mode0 || 'write');
  const [text, setText] = useState('');
  const [busy, setBusy] = useState(false);
  const [out, setOut] = useState(null);
  const [err, setErr] = useState(null);
  const go = async () => {
    setBusy(true); setErr(null); setOut(null);
    try {
      if (mode === 'explain') setOut({ text: await ai.explain(formula) });
      else if (mode === 'write') setOut(await ai.write(text));
      else setOut(await ai.clean(text, sample));
    } catch (e) { setErr(e.message || String(e)); } finally { setBusy(false); }
  };
  return (
    <Modal title={t('AI for this sheet')} onClose={onClose} footer={<><div className="grow" /><Btn className="ghost" onClick={onClose}>{t('Close')}</Btn>
      {out && out.formula && <Btn className="primary" onClick={() => { onInsert({ formula: out.formula }); onClose(); }}>{t('Insert formula')}</Btn>}
      {out && out.values && <Btn className="primary" onClick={() => { onInsert({ values: out.values }); onClose(); }}>{t('Write into the cells')}</Btn>}
      {!out && <Btn className="primary" disabled={busy || (mode !== 'explain' && !text.trim() && mode === 'write')} onClick={go}><Sparkles size={14} /> {busy ? t('Thinking …') : t('Go')}</Btn>}</>}>
      <div className="stack">
        <Raw as={Seg} value={mode} onChange={(m) => { setMode(m); setOut(null); }} options={[{ id: 'write', label: t('Write a formula') }, ...(formula ? [{ id: 'explain', label: t('Explain the formula') }] : []), ...(sample && sample.length ? [{ id: 'clean', label: t('Clean / split data') }] : [])]} />
        {mode === 'write' && <Field label={t('What should the formula calculate?')}><textarea className="input" rows={3} autoFocus value={text} placeholder={t('e.g. the sum of column B where column A is “Paid”')} onChange={(e) => setText(e.target.value)} /></Field>}
        {mode === 'explain' && <div className="mono small of-sheet-ai-f"><Keep>{`=${formula || ''}`}</Keep></div>}
        {mode === 'clean' && <><span className="small muted">{t('{n} selected values. Say what to do, e.g. “first and last name into two columns”, “trim and capitalize”, “dates as YYYY-MM-DD”.', { n: (sample || []).length })}</span>
          <textarea className="input" rows={2} value={text} onChange={(e) => setText(e.target.value)} aria-label={t('Instruction')} /></>}
        {err && <Callout kind="err">{err}</Callout>}
        {out && out.text && <div className="of-sheet-ai-out"><Keep>{out.text}</Keep></div>}
        {out && out.formula && <div className="mono of-sheet-ai-f"><Keep>{`=${out.display || out.formula}`}</Keep></div>}
        {out && out.values && <div className="of-sheet-ai-prev small"><table className="table"><tbody>{out.values.slice(0, 12).map((r, i) => <tr key={i}>{r.map((v, j) => <td key={j}><Keep>{String(v ?? '')}</Keep></td>)}</tr>)}</tbody></table>{out.values.length > 12 && <span className="muted">{t('… and {n} more rows', { n: out.values.length - 12 })}</span>}</div>}
        <span className="small muted">{t('Uses the AI provider set up in Hostwerk. Only the formula or the selected values are sent.')}</span>
      </div>
    </Modal>
  );
}

/* ---------------- calculation options ---------------- */
export function CalcOptionsDialog({ book, onApply, onClose, r1c1, onR1C1 }) {
  const [it, setIt] = useState({ ...book.iterate });
  const [d1904, setD] = useState(!!book.date1904);
  return (
    <Modal title={t('Calculation options')} onClose={onClose} footer={<><div className="grow" /><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply({ iterate: it, date1904: d1904 }); onClose(); }}>{t('OK')}</Btn></>}>
      <div className="stack">
        <label className="row"><input type="checkbox" checked={it.on} onChange={(e) => setIt({ ...it, on: e.target.checked })} /> {t('Enable iterative calculation (circular references)')}</label>
        <div className="row">
          <Field label={t('Maximum iterations')}><input type="number" className="input sm" min={1} max={32767} value={it.count} disabled={!it.on} onChange={(e) => setIt({ ...it, count: Math.max(1, Math.min(32767, +e.target.value || 100)) })} /></Field>
          <Field label={t('Maximum change')}><input type="number" className="input sm" step="0.001" min={0} value={it.delta} disabled={!it.on} onChange={(e) => setIt({ ...it, delta: Math.max(0, +e.target.value || 0) })} /></Field>
        </div>
        <label className="row"><input type="checkbox" checked={!!r1c1} onChange={(e) => onR1C1(e.target.checked)} /> {t('R1C1 reference style')}</label>
        <label className="row"><input type="checkbox" checked={d1904} onChange={(e) => setD(e.target.checked)} /> {t('Use 1904 date system')}</label>
        {d1904 !== !!book.date1904 && <Callout kind="warn">{t('Existing dates move by 4 years and 1 day when you switch the date system – like in Excel.')}</Callout>}
      </div>
    </Modal>
  );
}

/* ---------------- table name / style picker ---------------- */
export function TableStyleGrid({ value, onPick, look }) {
  const groups = [['Light', 21], ['Medium', 28], ['Dark', 11]];
  return (
    <div className="of-sheet-tstyles">
      {groups.map(([fam, n]) => <div key={fam}><div className="small muted">{t(fam)}</div><div className="of-sheet-tstyle-row">
        {Array.from({ length: n }, (_, i) => { const id = `TableStyle${fam}${i + 1}`; const L = look(id); return (
          <button key={id} type="button" className={`of-sheet-tstyle${value === id ? ' is-on' : ''}`} title={id} aria-label={id} onClick={() => onPick(id)}>
            <span style={{ background: L.head || '#fff', borderBottom: `2px solid ${L.line || L.head || '#ccc'}` }} />
            <span style={{ background: L.band || '#fff' }} /><span style={{ background: L.bandOdd || '#fff' }} /><span style={{ background: L.band || '#fff' }} />
          </button>); })}
      </div></div>)}
    </div>
  );
}
