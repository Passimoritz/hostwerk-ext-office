/* Data dialogs: sort, AutoFilter menu, text to columns, remove duplicates, data validation, CSV import options. */
import { useMemo, useState } from 'react';
import { Modal, Btn, Field, Switch } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Keep } from '../../_shared/keep.jsx';
import { Plus, Trash2, ArrowDownAZ, ArrowUpZA, Search } from 'lucide-react';
import { colName } from '../formula/refs.js';
import { sniff, parseCSV } from '../io/csv.js';

const colLabel = (book, sheet, g, c, header) => (header ? `${book.text(sheet, g.r1, c) || t('Column {c}', { c: colName(c) })}` : t('Column {c}', { c: colName(c) }));

export function SortDialog({ book, sheet, g, header0, onApply, onClose }) {
  const [header, setHeader] = useState(header0);
  const [levels, setLevels] = useState([{ col: g.c1, desc: false, by: 'value' }]);
  const cols = []; for (let c = g.c1; c <= g.c2; c++) cols.push(c);
  const upd = (i, p) => setLevels((l) => l.map((x, j) => (j === i ? { ...x, ...p } : x)));
  return (
    <Modal title={t('Sort')} subtitle={t('Range {r}', { r: `${colName(g.c1)}${g.r1 + 1}:${colName(g.c2)}${g.r2 + 1}` })} onClose={onClose}
      footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply(levels, header); onClose(); }}>{t('Sort')}</Btn></>}>
      <div className="stack">
        <label className="check small"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />{t('My data has headers')}</label>
        {levels.map((lv, i) => (
          <div key={i} className="row">
            <span className="small muted of-sheet-lvl">{i ? t('Then by') : t('Sort by')}</span>
            <select className="input sm grow" value={lv.col} onChange={(e) => upd(i, { col: +e.target.value })} aria-label={t('Column')}>{cols.map((c) => <option key={c} value={c}>{colLabel(book, sheet, g, c, header)}</option>)}</select>
            <select className="input sm" value={lv.desc ? 'desc' : 'asc'} onChange={(e) => upd(i, { desc: e.target.value === 'desc' })} aria-label={t('Order')}><option value="asc">{t('A to Z / smallest first')}</option><option value="desc">{t('Z to A / largest first')}</option></select>
            <Btn className="ghost sm icon" icon={Trash2} title={t('Delete level')} disabled={levels.length === 1} onClick={() => setLevels((l) => l.filter((_, j) => j !== i))} />
          </div>
        ))}
        <div><Btn className="sm" icon={Plus} onClick={() => setLevels((l) => [...l, { col: cols.find((c) => !l.some((x) => x.col === c)) ?? g.c1, desc: false, by: 'value' }])}>{t('Add level')}</Btn></div>
      </div>
    </Modal>
  );
}

/** popover content for an AutoFilter column */
export function FilterMenu({ book, sheet, col, onSort, onApply, onClose }) {
  const f = sheet.filter;
  const crit = (f && f.cols[col]) || {};
  const values = useMemo(() => {
    const set = new Map();
    const r2 = Math.min(f.range.r2, book.bounds(sheet).rows - 1);
    for (let r = f.range.r1 + 1; r <= r2; r++) {
      // only rows visible through the other columns' filters
      let hiddenByOther = false;
      for (const [c, x] of Object.entries(f.cols)) { if (+c === col) continue; void x; }
      void hiddenByOther;
      const text = book.text(sheet, r, col);
      if (!set.has(text)) set.set(text, book.value(sheet, r, col));
    }
    return [...set.entries()].sort((a, b) => (a[0] === '' ? 1 : b[0] === '' ? -1 : typeof a[1] === 'number' && typeof b[1] === 'number' ? a[1] - b[1] : a[0].localeCompare(b[0])));
  }, [book, sheet, col, f]);
  const [q, setQ] = useState('');
  const [checked, setChecked] = useState(() => new Set(crit.values ? crit.values.map((x) => String(x).toLowerCase()) : values.map(([v]) => v.toLowerCase())));
  const [cond, setCond] = useState(crit.cond || { op: '', v: '' });
  const numeric = values.filter(([, v]) => typeof v === 'number').length > values.length / 2;
  const shown = values.filter(([txt]) => !q || txt.toLowerCase().includes(q.toLowerCase()));
  const allOn = shown.every(([txt]) => checked.has(txt.toLowerCase()));
  const toggle = (txt) => setChecked((s) => { const n = new Set(s); const k = txt.toLowerCase(); if (n.has(k)) n.delete(k); else n.add(k); return n; });
  const apply = () => {
    const all = values.every(([txt]) => checked.has(txt.toLowerCase()));
    const c = cond.op ? { ...cond } : null;
    onApply(all && !c ? null : { values: all ? null : values.filter(([txt]) => checked.has(txt.toLowerCase())).map(([txt]) => txt), cond: c });
    onClose();
  };
  const OPS = numeric ? [['', t('No condition')], ['eq', t('Equals')], ['ne', t('Does not equal')], ['gt', t('Greater than')], ['ge', t('Greater than or equal to')], ['lt', t('Less than')], ['le', t('Less than or equal to')]] : [['', t('No condition')], ['eq', t('Equals')], ['ne', t('Does not equal')], ['contains', t('Contains')], ['notContains', t('Does not contain')], ['begins', t('Begins with')], ['ends', t('Ends with')]];
  return (
    <div className="of-sheet-filter stack">
      <div className="row-tight">
        <button type="button" className="btn ghost sm" onClick={() => { onSort(false); onClose(); }}><ArrowDownAZ size={14} />{t('Sort A to Z')}</button>
        <button type="button" className="btn ghost sm" onClick={() => { onSort(true); onClose(); }}><ArrowUpZA size={14} />{t('Sort Z to A')}</button>
      </div>
      <div className="row-tight">
        <select className="input sm" value={cond.op} onChange={(e) => setCond({ ...cond, op: e.target.value })} aria-label={numeric ? t('Number filter') : t('Text filter')}>{OPS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select>
        {cond.op && <input className="input sm grow" value={cond.v} onChange={(e) => setCond({ ...cond, v: e.target.value })} aria-label={t('Value')} />}
      </div>
      <div className="input-icon"><Search /><input className="input sm" placeholder={t('Search')} value={q} onChange={(e) => setQ(e.target.value)} aria-label={t('Search values')} /></div>
      <div className="of-sheet-filter-list">
        <label className="check small"><input type="checkbox" checked={allOn} onChange={() => setChecked((s) => { const n = new Set(s); for (const [txt] of shown) { if (allOn) n.delete(txt.toLowerCase()); else n.add(txt.toLowerCase()); } return n; })} /><b>{t('(Select all)')}</b></label>
        {shown.slice(0, 1000).map(([txt]) => <label key={txt} className="check small"><input type="checkbox" checked={checked.has(txt.toLowerCase())} onChange={() => toggle(txt)} />{txt === '' ? t('(Blanks)') : <Keep>{txt}</Keep>}</label>)}
        {shown.length > 1000 && <div className="small muted">{t('{n} more values – use the search', { n: shown.length - 1000 })}</div>}
      </div>
      <div className="row-tight of-sheet-filter-foot">
        <button type="button" className="btn ghost sm" onClick={() => { onApply(null); onClose(); }}>{t('Clear filter')}</button>
        <div className="grow" />
        <button type="button" className="btn ghost sm" onClick={onClose}>{t('Cancel')}</button>
        <button type="button" className="btn primary sm" onClick={apply}>{t('OK')}</button>
      </div>
    </div>
  );
}

export function TextToColumnsDialog({ book, sheet, g, onApply, onClose }) {
  const [mode, setMode] = useState('delim');
  const [d, setD] = useState({ tab: true, semicolon: false, comma: true, space: false, other: '' });
  const [cons, setCons] = useState(false);
  const [widths, setWidths] = useState('10, 10');
  const delims = Object.entries({ tab: 'tab', semicolon: ';', comma: ',', space: 'space' }).filter(([k]) => d[k]).map(([, v]) => v);
  const sample = []; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 6); r++) sample.push(String(book.value(sheet, r, g.c1) ?? ''));
  const split = (v) => {
    if (mode === 'fixed') { const ws = widths.split(/[,;\s]+/).map(Number).filter((x) => x > 0); const out = []; let p = 0; for (const w of ws) { out.push(v.slice(p, p + w)); p += w; } out.push(v.slice(p)); return out; }
    const ds = delims.map((x) => (x === 'tab' ? '\t' : x === 'space' ? ' ' : x)).concat(d.other ? [d.other] : []);
    if (!ds.length) return [v];
    const re = new RegExp(`[${ds.map((x) => x.replace(/[\]\\^-]/g, '\\$&')).join('')}]${cons ? '+' : ''}`);
    return v.split(re);
  };
  return (
    <Modal title={t('Text to columns')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply(mode === 'fixed' ? { fixed: widths.split(/[,;\s]+/).map(Number).filter((x) => x > 0) } : { delimiters: delims, other: d.other, treatConsecutive: cons }); onClose(); }}>{t('Split')}</Btn></>}>
      <div className="stack">
        <div className="row"><label className="check small"><input type="radio" checked={mode === 'delim'} onChange={() => setMode('delim')} />{t('Delimited')}</label><label className="check small"><input type="radio" checked={mode === 'fixed'} onChange={() => setMode('fixed')} />{t('Fixed width')}</label></div>
        {mode === 'delim' ? (
          <div className="row">
            {[['tab', t('Tab')], ['semicolon', t('Semicolon')], ['comma', t('Comma')], ['space', t('Space')]].map(([k, l]) => <label key={k} className="check small"><input type="checkbox" checked={d[k]} onChange={(e) => setD({ ...d, [k]: e.target.checked })} />{l}</label>)}
            <label className="row-tight small">{t('Other')}<input className="input sm of-sheet-short" maxLength={1} value={d.other} onChange={(e) => setD({ ...d, other: e.target.value })} aria-label={t('Other delimiter')} /></label>
            <label className="check small"><input type="checkbox" checked={cons} onChange={(e) => setCons(e.target.checked)} />{t('Treat consecutive delimiters as one')}</label>
          </div>
        ) : <Field label={t('Column widths (characters)')}><input className="input sm" value={widths} onChange={(e) => setWidths(e.target.value)} /></Field>}
        <Field label={t('Preview')}><table className="table of-sheet-preview"><tbody>{sample.map((v, i) => <tr key={i}>{split(v).map((p, j) => <td key={j}>{p}</td>)}</tr>)}</tbody></table></Field>
        <p className="small muted">{t('Existing values to the right are overwritten.')}</p>
      </div>
    </Modal>
  );
}

export function RemoveDuplicatesDialog({ book, sheet, g, onApply, onClose }) {
  const [header, setHeader] = useState(() => typeof book.value(sheet, g.r1, g.c1) === 'string' && g.r2 > g.r1);
  const cols = []; for (let c = g.c1; c <= g.c2; c++) cols.push(c);
  const [on, setOn] = useState(() => new Set(cols));
  return (
    <Modal title={t('Remove duplicates')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" disabled={!on.size} onClick={() => { onApply([...on], header); onClose(); }}>{t('Remove duplicates')}</Btn></>}>
      <div className="stack">
        <label className="check small"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />{t('My data has headers')}</label>
        <div className="row-tight"><Btn className="xs ghost" onClick={() => setOn(new Set(cols))}>{t('Select all')}</Btn><Btn className="xs ghost" onClick={() => setOn(new Set())}>{t('Unselect all')}</Btn></div>
        <div className="of-sheet-filter-list">{cols.map((c) => <label key={c} className="check small"><input type="checkbox" checked={on.has(c)} onChange={() => setOn((s) => { const n = new Set(s); if (n.has(c)) n.delete(c); else n.add(c); return n; })} /><Keep>{colLabel(book, sheet, g, c, header)}</Keep></label>)}</div>
      </div>
    </Modal>
  );
}

export function ValidationDialog({ rule, toDisplay, fromDisplay, onApply, onClear, onClose }) {
  const r0 = rule || { type: 'any', op: 'between', f1: '', f2: '', allowBlank: true, dropdown: true, prompt: null, error: { style: 'stop', title: '', text: '', show: true } };
  const [v, setV] = useState({ ...r0, f1: r0.f1 ? toDisplay(r0.f1) : '', f2: r0.f2 ? toDisplay(r0.f2) : '', prompt: r0.prompt || { title: '', text: '' }, error: { style: 'stop', show: true, title: '', text: '', ...(r0.error || {}) } });
  const set = (p) => setV((x) => ({ ...x, ...p }));
  const listText = v.type === 'list' && /^".*"$/.test(v.f1) ? v.f1.slice(1, -1) : null;
  const twoArgs = v.op === 'between' || v.op === 'notBetween';
  const TYPES = [['any', t('Any value')], ['whole', t('Whole number')], ['decimal', t('Decimal')], ['list', t('List')], ['date', t('Date')], ['time', t('Time')], ['textLength', t('Text length')], ['custom', t('Custom')]];
  const OPS = [['between', t('between')], ['notBetween', t('not between')], ['equal', t('equal to')], ['notEqual', t('not equal to')], ['greaterThan', t('greater than')], ['lessThan', t('less than')], ['greaterThanOrEqual', t('greater than or equal to')], ['lessThanOrEqual', t('less than or equal to')]];
  const ok = () => {
    const out = { ...v, f1: v.f1 ? fromDisplay(v.f1) : '', f2: v.f2 ? fromDisplay(v.f2) : '', prompt: v.prompt.title || v.prompt.text ? v.prompt : null };
    if (v.type === 'any') onClear(); else onApply(out);
    onClose();
  };
  return (
    <Modal title={t('Data validation')} onClose={onClose} footer={<><Btn className="ghost" onClick={() => { onClear(); onClose(); }}>{t('Clear all')}</Btn><div className="grow" /><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={ok}>{t('OK')}</Btn></>}>
      <div className="stack">
        <div className="row">
          <Field label={t('Allow')}><select className="input sm" value={v.type} onChange={(e) => set({ type: e.target.value })}>{TYPES.map(([id, l]) => <option key={id} value={id}>{l}</option>)}</select></Field>
          {!['any', 'list', 'custom'].includes(v.type) && <Field label={t('Data')}><select className="input sm" value={v.op} onChange={(e) => set({ op: e.target.value })}>{OPS.map(([id, l]) => <option key={id} value={id}>{l}</option>)}</select></Field>}
        </div>
        {v.type === 'list' && <Field label={t('Source')} hint={t('Items separated by commas (Open, Done) or a range like =$A$1:$A$5')}><input className="input" value={listText != null ? listText : v.f1} onChange={(e) => { const x = e.target.value; set({ f1: x.startsWith('=') ? x.slice(1) : `"${x}"` }); }} /></Field>}
        {v.type === 'custom' && <Field label={t('Formula')} hint={t('The value is valid when the formula returns TRUE, e.g. =A1>0')}><input className="input mono" value={v.f1 ? `=${v.f1}` : ''} onChange={(e) => set({ f1: e.target.value.replace(/^=/, '') })} /></Field>}
        {!['any', 'list', 'custom'].includes(v.type) && <div className="row">
          <Field label={twoArgs ? t('Minimum') : t('Value')}><input className="input sm" value={v.f1} onChange={(e) => set({ f1: e.target.value })} /></Field>
          {twoArgs && <Field label={t('Maximum')}><input className="input sm" value={v.f2} onChange={(e) => set({ f2: e.target.value })} /></Field>}
        </div>}
        <div className="row">
          <label className="check small"><input type="checkbox" checked={!!v.allowBlank} onChange={(e) => set({ allowBlank: e.target.checked })} />{t('Ignore blank')}</label>
          {v.type === 'list' && <label className="check small"><input type="checkbox" checked={v.dropdown !== false} onChange={(e) => set({ dropdown: e.target.checked })} />{t('In-cell dropdown')}</label>}
        </div>
        <Field label={t('Input message (shown when the cell is selected)')}>
          <input className="input sm" placeholder={t('Title')} value={v.prompt.title} onChange={(e) => set({ prompt: { ...v.prompt, title: e.target.value } })} />
          <textarea className="input" rows={2} placeholder={t('Message')} value={v.prompt.text} onChange={(e) => set({ prompt: { ...v.prompt, text: e.target.value } })} />
        </Field>
        <Field label={t('Error alert (after invalid data)')}>
          <div className="row">
            <select className="input sm" value={v.error.style} onChange={(e) => set({ error: { ...v.error, style: e.target.value } })}><option value="stop">{t('Stop')}</option><option value="warning">{t('Warning')}</option><option value="information">{t('Information')}</option></select>
            <Switch checked={v.error.show !== false} onChange={(x) => set({ error: { ...v.error, show: x } })} label={t('Show error alert')} />
          </div>
          <input className="input sm" placeholder={t('Title')} value={v.error.title} onChange={(e) => set({ error: { ...v.error, title: e.target.value } })} />
          <textarea className="input" rows={2} placeholder={t('Message')} value={v.error.text} onChange={(e) => set({ error: { ...v.error, text: e.target.value } })} />
        </Field>
      </div>
    </Modal>
  );
}

const ENCODINGS = [['auto', 'Detect automatically'], ['utf-8', 'UTF-8'], ['windows-1252', 'Windows-1252 (Western)'], ['iso-8859-1', 'ISO-8859-1 (Latin-1)'], ['utf-16le', 'UTF-16 LE'], ['utf-16be', 'UTF-16 BE']];
export function CsvDialog({ bytes, initial, onApply, onClose, title }) {
  const [o, setO] = useState({ delimiter: 'auto', encoding: 'auto', header: false, numberLocale: '', ...initial });
  const sn = useMemo(() => { try { return sniff(bytes, { encoding: o.encoding, delimiter: o.delimiter }); } catch { return null; } }, [bytes, o.encoding, o.delimiter]);
  const rows = sn ? parseCSV(sn.text, { delimiter: sn.delimiter, maxRows: 12 }) : [];
  const D = [['auto', t('Detect automatically')], [';', t('Semicolon ;')], [',', t('Comma ,')], ['\t', t('Tab')], ['|', t('Pipe |')]];
  return (
    <Modal size="wide" title={title || t('Text import')} subtitle={sn ? t('Detected: {enc}, delimiter {d}', { enc: sn.encoding, d: sn.delimiter === '\t' ? t('Tab') : sn.delimiter }) : ''} onClose={onClose}
      footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply({ ...o, delimiter: o.delimiter === 'auto' ? sn && sn.delimiter : o.delimiter, encoding: o.encoding === 'auto' ? sn && sn.encoding : o.encoding }); onClose(); }}>{t('Import')}</Btn></>}>
      <div className="stack">
        <div className="row">
          <Field label={t('Delimiter')}><select className="input sm" value={o.delimiter} onChange={(e) => setO({ ...o, delimiter: e.target.value })}>{D.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
          <Field label={t('Encoding')}><select className="input sm" value={o.encoding} onChange={(e) => setO({ ...o, encoding: e.target.value })}>{ENCODINGS.map(([v, l]) => <option key={v} value={v}>{t(l)}</option>)}</select></Field>
          <Field label={t('Numbers and dates')}><select className="input sm" value={o.numberLocale} onChange={(e) => setO({ ...o, numberLocale: e.target.value })}><option value="">{t('Like my language')}</option><option value=",">{t('German (1.234,5 · 18.11.2026)')}</option><option value=".">{t('English (1,234.5 · 11/18/2026)')}</option></select></Field>
          <label className="check small"><input type="checkbox" checked={o.header} onChange={(e) => setO({ ...o, header: e.target.checked })} />{t('First row is a header')}</label>
        </div>
        <div className="of-sheet-csvprev"><table className="table of-sheet-preview"><tbody>{rows.map((r, i) => <tr key={i}>{r.slice(0, 20).map((c, j) => <td key={j} className={i === 0 && o.header ? 'is-head' : ''}>{c}</td>)}</tr>)}</tbody></table></div>
      </div>
    </Modal>
  );
}
