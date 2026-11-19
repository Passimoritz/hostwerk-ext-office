/* Find & replace, go to, names, insert function, link, note, chart, print, paste special, insert/delete cells, macros, statistics, command search. */
import { useEffect, useMemo, useRef, useState } from 'react';
import { Modal, Btn, Field, Empty, SideSheet, Callout, Seg } from '@hostwerk/ui';
import * as HUI from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Search, Replace, Plus, Trash2, Pencil, Code, Tag } from 'lucide-react';
import { addr, rangeA1, parseRange, quoteSheet, colName } from '../formula/refs.js';
import { FUNC_INFO, CATEGORIES } from '../formula/funcinfo.js';
import { FUNCTION_NAMES } from '../formula/functions.js';
import { localFn, localArgList, isExcelName } from '../formula/i18n.js';
import { Keep } from '../../_shared/keep.jsx';
import { CHART_TYPES } from '../chartData.js';

const Sheet = HUI.SideSheet || Modal;

export function FindPanel({ book, sheet, replace: rep0, onGo, onClose, readOnly }) {
  const [q, setQ] = useState('');
  const [rep, setRep] = useState('');
  const [mode, setMode] = useState(rep0 ? 'replace' : 'find');
  const [o, setO] = useState({ matchCase: false, whole: false, inFormulas: false, allSheets: false });
  const [hits, setHits] = useState(null);
  const [idx, setIdx] = useState(-1);
  const inp = useRef(null);
  useEffect(() => { if (inp.current) inp.current.focus(); }, []);
  const opts = () => ({ query: q, ...o, sheet });
  const findAll = () => { const h = book.findAll(opts()); setHits(h); return h; };
  const next = (dir = 1) => {
    const h = hits && hits.length && hits._q === q ? hits : findAll();
    h._q = q;
    if (!h.length) { setIdx(-1); return; }
    const i = (idx + dir + h.length) % h.length;
    setIdx(i); onGo(h[i]);
  };
  const replaceOne = () => { if (idx < 0 || !hits || !hits[idx]) { next(); return; } book.replaceAll(opts(), rep, [hits[idx]]); const h = findAll(); setIdx(Math.min(idx, h.length - 1) - 1); next(); };
  const replaceAll = () => { const n = book.replaceAll(opts(), rep); HUI.toast.success(t('{n} replacements made', { n })); findAll(); };
  return (
    <Sheet title={t('Find & replace')} onClose={onClose}>
      <div className="stack">
        <Seg value={mode} onChange={setMode} options={[{ id: 'find', label: t('Find') }, { id: 'replace', label: t('Replace') }]} />
        <Field label={t('Find what')}><input ref={inp} className="input" value={q} onChange={(e) => { setQ(e.target.value); setHits(null); setIdx(-1); }} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); next(e.shiftKey ? -1 : 1); } }} /></Field>
        {mode === 'replace' && <Field label={t('Replace with')}><input className="input" value={rep} onChange={(e) => setRep(e.target.value)} /></Field>}
        <div className="row">
          <label className="check small"><input type="checkbox" checked={o.matchCase} onChange={(e) => { setO({ ...o, matchCase: e.target.checked }); setHits(null); }} />{t('Match case')}</label>
          <label className="check small"><input type="checkbox" checked={o.whole} onChange={(e) => { setO({ ...o, whole: e.target.checked }); setHits(null); }} />{t('Match entire cell contents')}</label>
          <label className="check small"><input type="checkbox" checked={o.inFormulas} onChange={(e) => { setO({ ...o, inFormulas: e.target.checked }); setHits(null); }} />{t('Search in formulas')}</label>
          <label className="check small"><input type="checkbox" checked={o.allSheets} onChange={(e) => { setO({ ...o, allSheets: e.target.checked }); setHits(null); }} />{t('All sheets')}</label>
        </div>
        <div className="row">
          <Btn className="sm" icon={Search} disabled={!q} onClick={() => next(1)}>{t('Find next')}</Btn>
          <Btn className="sm ghost" disabled={!q} onClick={() => { findAll(); setIdx(-1); }}>{t('Find all')}</Btn>
          {mode === 'replace' && <Btn className="sm" icon={Replace} disabled={!q || readOnly} onClick={replaceOne}>{t('Replace')}</Btn>}
          {mode === 'replace' && <Btn className="sm primary" disabled={!q || readOnly} onClick={replaceAll}>{t('Replace all')}</Btn>}
        </div>
        {hits && (hits.length ? (
          <div className="of-sheet-hits">
            <div className="small muted">{t('{n} cells found', { n: hits.length })}</div>
            {hits.slice(0, 500).map((h, i) => <button key={i} type="button" className={`of-sheet-hit${i === idx ? ' is-on' : ''}`} onClick={() => { setIdx(i); onGo(h); }}><span className="mono small">{h.sheet.name}!{addr(h.r, h.c)}</span><span className="small grow">{book.text(h.sheet, h.r, h.c)}</span></button>)}
          </div>
        ) : <p className="small muted">{t('Nothing found.')}</p>)}
      </div>
    </Sheet>
  );
}

export function GoToDialog({ book, onGo, onClose, special }) {
  const [v, setV] = useState('');
  const [sp, setSp] = useState('blanks');
  const names = [...book.names.values()].filter((n) => !n.hidden && !n.raw);
  return (
    <Modal title={special ? t('Go to special') : t('Go to')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { if (onGo(special ? { special: sp } : v)) onClose(); }}>{t('OK')}</Btn></>}>
      {special ? (
        <div className="stack">
          {[['blanks', t('Blanks')], ['constants', t('Constants')], ['formulas', t('Formulas')], ['notes', t('Notes')], ['errors', t('Formulas with errors')], ['visible', t('Visible cells only')], ['validation', t('Data validation')], ['cf', t('Conditional formats')]].map(([id, l]) => <label key={id} className="check small"><input type="radio" checked={sp === id} onChange={() => setSp(id)} />{l}</label>)}
        </div>
      ) : (
        <div className="stack">
          <Field label={t('Reference')} hint={t('e.g. B12, A1:D20, Sheet2!C3 or a name')}><input className="input mono" autoFocus value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && onGo(v)) onClose(); }} /></Field>
          {names.length > 0 && <div className="of-sheet-list compact">{names.map((n) => <button key={n.name + (n.scope ? n.scope.id : '')} type="button" onClick={() => setV(n.name)} onDoubleClick={() => { if (onGo(n.name)) onClose(); }}><b><Keep>{n.name}</Keep></b> <span className="mono small muted"><Keep>{`=${n.ref}`}</Keep></span></button>)}</div>}
        </div>
      )}
    </Modal>
  );
}

export function NamesDialog({ book, current, toDisplay, onSave, onDelete, onClose, readOnly, startNew }) {
  const [edit, setEdit] = useState(startNew ? { name: '', ref: current, scope: '', comment: '', key: null } : null);
  const list = [...book.names.entries()].filter(([, n]) => !n.raw);
  const save = () => { if (onSave(edit)) setEdit(null); };
  return (
    <Modal size="wide" title={t('Name manager')} onClose={onClose} footer={<>{!edit && <Btn icon={Plus} disabled={readOnly} onClick={() => setEdit({ name: '', ref: current, scope: '', comment: '', key: null })}>{t('New name …')}</Btn>}<div className="grow" /><Btn className="primary" onClick={onClose}>{t('Close')}</Btn></>}>
      {edit ? (
        <div className="stack">
          <div className="row">
            <Field label={t('Name')}><input className="input" autoFocus value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} /></Field>
            <Field label={t('Scope')}><select className="input sm" value={edit.scope} onChange={(e) => setEdit({ ...edit, scope: e.target.value })}><option value="">{t('Workbook')}</option>{book.sheets.map((s) => <option key={s.id} value={s.id}><Keep>{s.name}</Keep></option>)}</select></Field>
          </div>
          <Field label={t('Refers to')}><input className="input mono" value={edit.ref} onChange={(e) => setEdit({ ...edit, ref: e.target.value })} /></Field>
          <Field label={t('Comment')}><input className="input" value={edit.comment} onChange={(e) => setEdit({ ...edit, comment: e.target.value })} /></Field>
          <div className="row"><Btn className="ghost" onClick={() => setEdit(null)}>{t('Cancel')}</Btn><Btn className="primary" disabled={!edit.name || !edit.ref} onClick={save}>{t('Save')}</Btn></div>
        </div>
      ) : !list.length ? <Empty icon={Tag} title={t('No names yet')} text={t('Names make formulas easier to read, e.g. =SUM(Sales).')} /> : (
        <table className="table">
          <thead><tr><th>{t('Name')}</th><th>{t('Refers to')}</th><th>{t('Scope')}</th><th>{t('Comment')}</th><th /></tr></thead>
          <tbody>{list.map(([k, n]) => (
            <tr key={k}><td><b><Keep>{n.name}</Keep></b></td><td className="mono small"><Keep>{`=${toDisplay(n.ref)}`}</Keep></td><td>{n.scope ? <Keep>{n.scope.name}</Keep> : t('Workbook')}</td><td className="small"><Keep>{n.comment}</Keep></td>
              <td className="row-tight"><Btn className="xs ghost icon" icon={Pencil} title={t('Edit')} disabled={readOnly} onClick={() => setEdit({ name: n.name, ref: `=${toDisplay(n.ref)}`, scope: n.scope ? n.scope.id : '', comment: n.comment || '', key: k })} /><Btn className="xs ghost icon" icon={Trash2} title={t('Delete')} disabled={readOnly} onClick={() => onDelete(k)} /></td></tr>
          ))}</tbody>
        </table>
      )}
    </Modal>
  );
}

export function InsertFunctionDialog({ loc, onPick, onClose }) {
  const [q, setQ] = useState('');
  const [cat, setCat] = useState('');
  const [selName, setSelName] = useState('SUM');
  // shown in the user's language (SUMME …), stored and picked by the English name; the engine's own aliases are left out
  const all = useMemo(() => FUNCTION_NAMES.filter(isExcelName).map((n) => ({ ...(FUNC_INFO.get(n) || { name: n, cat: '', args: '', desc: '' }), local: localFn(n, loc), argList: localArgList((FUNC_INFO.get(n) || {}).args, loc) })).sort((a, b) => a.local.localeCompare(b.local)), [loc]);
  const list = all.filter((f) => (!cat || f.cat === cat) && (!q || f.local.includes(q.toUpperCase()) || f.name.includes(q.toUpperCase()) || t(f.desc).toLowerCase().includes(q.toLowerCase()))).slice(0, 400);
  const cur = all.find((f) => f.name === selName);
  return (
    <Modal size="wide" title={t('Insert function')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" disabled={!cur} onClick={() => { onPick(selName); onClose(); }}>{t('Insert')}</Btn></>}>
      <div className="stack">
        <div className="row">
          <div className="input-icon grow"><Search /><input className="input" autoFocus placeholder={t('Search for a function')} value={q} onChange={(e) => setQ(e.target.value)} /></div>
          <select className="input sm" value={cat} onChange={(e) => setCat(e.target.value)} aria-label={t('Category')}><option value="">{t('All')}</option>{CATEGORIES.map((c) => <option key={c} value={c}>{t(c)}</option>)}</select>
        </div>
        <div className="of-sheet-list tall" role="listbox">{list.map((f) => <button key={f.name} type="button" role="option" aria-selected={f.name === selName} className={f.name === selName ? 'is-on' : ''} onClick={() => setSelName(f.name)} onDoubleClick={() => { onPick(f.name); onClose(); }}><b className="mono"><Keep>{f.local}</Keep></b>{f.desc && <span className="small muted"> – {t(f.desc)}</span>}</button>)}</div>
        {cur && <div className="of-sheet-sample"><b className="mono"><Keep>{`${cur.local}(${cur.argList.join(loc && loc.decimal === ',' ? '; ' : ', ')})`}</Keep></b><span className="small">{cur.desc ? t(cur.desc) : t('Function from the standard library.')}</span></div>}
        <p className="small muted">{t('Files keep the English function names – Excel and LibreOffice show them in their own language. Settings → Spreadsheets → Function names.')}</p>
      </div>
    </Modal>
  );
}

export function LinkDialog({ book, init, onApply, onClose }) {
  const internal = init.link && init.link.startsWith('#');
  const [kind, setKind] = useState(internal ? 'place' : /^mailto:/i.test(init.link || '') ? 'mail' : 'web');
  const [text, setText] = useState(init.text || '');
  const [url, setUrl] = useState(!internal && init.link ? init.link.replace(/^mailto:/i, '') : '');
  const g0 = internal ? init.link.slice(1) : '';
  const [sheetName, setSheetName] = useState(g0.includes('!') ? g0.split('!')[0].replace(/^'|'$/g, '') : book.sheets[0].name);
  const [cell, setCell] = useState(g0.includes('!') ? g0.split('!')[1] : 'A1');
  const ok = () => {
    let link;
    if (kind === 'web') link = /^[a-z][\w+.-]*:/i.test(url) ? url : `https://${url}`;
    else if (kind === 'mail') link = `mailto:${url}`;
    else link = `#${quoteSheet(sheetName)}!${cell}`;
    onApply(link, text); onClose();
  };
  return (
    <Modal title={t('Insert link')} onClose={onClose} footer={<>{init.link && <Btn className="ghost" onClick={() => { onApply(null, text); onClose(); }}>{t('Remove link')}</Btn>}<div className="grow" /><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" disabled={kind !== 'place' && !url} onClick={ok}>{t('OK')}</Btn></>}>
      <div className="stack">
        <Seg value={kind} onChange={setKind} options={[{ id: 'web', label: t('Web page') }, { id: 'mail', label: t('E-mail') }, { id: 'place', label: t('Place in this workbook') }]} />
        <Field label={t('Text to display')}><input className="input" value={text} onChange={(e) => setText(e.target.value)} /></Field>
        {kind !== 'place' ? <Field label={kind === 'web' ? t('Address') : t('E-mail address')}><input className="input" autoFocus value={url} onChange={(e) => setUrl(e.target.value)} placeholder={kind === 'web' ? 'https://' : 'name@example.com'} /></Field> : (
          <div className="row">
            <Field label={t('Sheet')}><select className="input sm" value={sheetName} onChange={(e) => setSheetName(e.target.value)}>{book.sheets.map((s) => <option key={s.id} value={s.name}><Keep>{s.name}</Keep></option>)}</select></Field>
            <Field label={t('Cell reference')}><input className="input sm mono" value={cell} onChange={(e) => setCell(e.target.value)} /></Field>
          </div>
        )}
      </div>
    </Modal>
  );
}

export function NoteDialog({ init, author, onApply, onClose }) {
  const [v, setV] = useState(init || '');
  return (
    <Modal title={init ? t('Edit note') : t('New note')} onClose={onClose} footer={<>{init && <Btn className="ghost" onClick={() => { onApply(null); onClose(); }}>{t('Delete note')}</Btn>}<div className="grow" /><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply(v.trim() ? v : null); onClose(); }}>{t('Save')}</Btn></>}>
      <div className="stack">
        {author && <span className="small muted"><Keep>{author}</Keep></span>}
        <textarea className="input" rows={6} autoFocus value={v} onChange={(e) => setV(e.target.value)} aria-label={t('Note')} />
      </div>
    </Modal>
  );
}

export function ChartDialog({ chart, onApply, onDelete, onClose }) {
  const [c, setC] = useState({ ...chart });
  const set = (p) => setC((x) => ({ ...x, ...p }));
  const usesRange = !c.series || !c.series.length;
  return (
    <Modal title={t('Chart')} onClose={onClose} footer={<><Btn className="ghost" onClick={() => { onDelete(); onClose(); }}>{t('Delete chart')}</Btn><div className="grow" /><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply(c); onClose(); }}>{t('OK')}</Btn></>}>
      <div className="stack">
        <Field label={t('Type')}><div className="row-tight of-sheet-wrap">{CHART_TYPES.map((x) => <button key={x.id} type="button" className={`btn sm${c.type === x.id ? ' primary' : ' ghost'}`} onClick={() => set({ type: x.id })}>{t(x.label)}</button>)}</div></Field>
        {['column', 'bar', 'area'].includes(c.type) && <Field label={t('Grouping')}><Seg value={c.stacked || 'none'} onChange={(v) => set({ stacked: v === 'none' ? false : v })} options={[{ id: 'none', label: t('Clustered') }, { id: 'stacked', label: t('Stacked') }, { id: 'percent', label: t('100 %') }]} /></Field>}
        <Field label={t('Title')}><input className="input" value={c.title || ''} onChange={(e) => set({ title: e.target.value })} /></Field>
        <div className="row">
          <Field label={t('Legend')}><select className="input sm" value={c.legend || 'right'} onChange={(e) => set({ legend: e.target.value })}>{[['right', t('Right')], ['bottom', t('Bottom')], ['top', t('Top')], ['left', t('Left')], ['none', t('None')]].map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
          {!['pie', 'doughnut'].includes(c.type) && <Field label={t('Horizontal axis title')}><input className="input sm" value={c.xTitle || ''} onChange={(e) => set({ xTitle: e.target.value })} /></Field>}
          {!['pie', 'doughnut'].includes(c.type) && <Field label={t('Vertical axis title')}><input className="input sm" value={c.yTitle || ''} onChange={(e) => set({ yTitle: e.target.value })} /></Field>}
        </div>
        {usesRange ? (
          <div className="row">
            <Field label={t('Data range')}><input className="input sm mono" value={c.range || ''} onChange={(e) => set({ range: e.target.value })} /></Field>
            <Field label={t('Series in')}><Seg value={c.byRows ? 'rows' : 'cols'} onChange={(v) => set({ byRows: v === 'rows' })} options={[{ id: 'cols', label: t('Columns') }, { id: 'rows', label: t('Rows') }]} /></Field>
          </div>
        ) : (
          <Field label={t('Series')}>
            {c.series.map((s, i) => (
              <div key={i} className="row-tight">
                <input className="input sm" value={s.name || ''} aria-label={t('Name')} onChange={(e) => set({ series: c.series.map((x, j) => (j === i ? { ...x, name: e.target.value, nameRef: null } : x)) })} />
                <input className="input sm mono grow" value={s.valRef || ''} aria-label={t('Values')} onChange={(e) => set({ series: c.series.map((x, j) => (j === i ? { ...x, valRef: e.target.value } : x)) })} />
                <input className="input sm mono" value={s.catRef || ''} aria-label={t('Categories')} onChange={(e) => set({ series: c.series.map((x, j) => (j === i ? { ...x, catRef: e.target.value } : x)) })} />
              </div>
            ))}
            <Btn className="xs ghost" onClick={() => set({ series: null })}>{t('Use one data range instead')}</Btn>
          </Field>
        )}
      </div>
    </Modal>
  );
}

export function PrintDialog({ sheet, hasSelection, hasConverter, onPrint, onPdf, onClose }) {
  const [o, setO] = useState({ what: sheet.print.area ? 'area' : 'sheet', orientation: sheet.print.orientation || 'portrait', fit: sheet.print.fit || 'width', gridlines: !!sheet.print.gridlines, headings: !!sheet.print.headings });
  return (
    <Modal title={t('Print')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn>{hasConverter && <Btn onClick={() => { onPdf(o); onClose(); }}>{t('Save as PDF')}</Btn>}<Btn className="primary" onClick={() => { onPrint(o); onClose(); }}>{t('Print')}</Btn></>}>
      <div className="stack">
        <Field label={t('Print what')}><select className="input sm" value={o.what} onChange={(e) => setO({ ...o, what: e.target.value })}><option value="sheet">{t('Active sheet')}</option>{hasSelection && <option value="selection">{t('Selection')}</option>}{sheet.print.area && <option value="area">{t('Print area')} ({sheet.print.area})</option>}</select></Field>
        <div className="row">
          <Field label={t('Orientation')}><Seg value={o.orientation} onChange={(v) => setO({ ...o, orientation: v })} options={[{ id: 'portrait', label: t('Portrait') }, { id: 'landscape', label: t('Landscape') }]} /></Field>
          <Field label={t('Scaling')}><select className="input sm" value={o.fit} onChange={(e) => setO({ ...o, fit: e.target.value })}><option value="none">{t('No scaling')}</option><option value="width">{t('Fit all columns on one page')}</option><option value="page">{t('Fit sheet on one page')}</option></select></Field>
        </div>
        <div className="row">
          <label className="check small"><input type="checkbox" checked={o.gridlines} onChange={(e) => setO({ ...o, gridlines: e.target.checked })} />{t('Print gridlines')}</label>
          <label className="check small"><input type="checkbox" checked={o.headings} onChange={(e) => setO({ ...o, headings: e.target.checked })} />{t('Print row and column headings')}</label>
        </div>
        {!hasConverter && <p className="small muted">{t('Tip: choose “Save as PDF” as the printer to get a PDF. Charts are printed by the Office engine only.')}</p>}
      </div>
    </Modal>
  );
}

export function PasteSpecialDialog({ onApply, onClose }) {
  const [what, setWhat] = useState('all');
  const [tr, setTr] = useState(false);
  const [skip, setSkip] = useState(false);
  return (
    <Modal title={t('Paste special')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply({ what, transpose: tr, skipBlanks: skip }); onClose(); }}>{t('Paste')}</Btn></>}>
      <div className="stack">
        {[['all', t('All')], ['formulas', t('Formulas')], ['values', t('Values')], ['formats', t('Formats')], ['noBorders', t('All except borders')], ['colWidths', t('Column widths')]].map(([id, l]) => <label key={id} className="check small"><input type="radio" checked={what === id} onChange={() => setWhat(id)} />{l}</label>)}
        <div className="row"><label className="check small"><input type="checkbox" checked={skip} onChange={(e) => setSkip(e.target.checked)} />{t('Skip blanks')}</label><label className="check small"><input type="checkbox" checked={tr} onChange={(e) => setTr(e.target.checked)} />{t('Transpose')}</label></div>
      </div>
    </Modal>
  );
}

export function CellsDialog({ insert, onApply, onClose }) {
  const [v, setV] = useState(insert ? 'down' : 'up');
  const opts = insert ? [['right', t('Shift cells right')], ['down', t('Shift cells down')], ['row', t('Entire row')], ['col', t('Entire column')]] : [['left', t('Shift cells left')], ['up', t('Shift cells up')], ['row', t('Entire row')], ['col', t('Entire column')]];
  return (
    <Modal size="dialog" title={insert ? t('Insert cells') : t('Delete cells')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply(v); onClose(); }}>{t('OK')}</Btn></>}>
      <div className="stack">{opts.map(([id, l]) => <label key={id} className="check small"><input type="radio" checked={v === id} onChange={() => setV(id)} />{l}</label>)}</div>
    </Modal>
  );
}

export function MacrosSheet({ load, onClose }) {
  const [mods, setMods] = useState(null);
  const [cur, setCur] = useState(0);
  useEffect(() => { load().then((m) => setMods(m || [])).catch(() => setMods([])); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <Sheet title={t('Macros')} subtitle={t('Read only – macros are kept in the file and never run here.')} onClose={onClose} size="wide">
      {!mods ? <HUI.Loading /> : !mods.length ? <Empty icon={Code} title={t('No readable macros')} text={t('The macro project is kept unchanged when you save.')} /> : (
        <div className="stack">
          <Callout kind="info" title={t('Macros are not executed')}>{t('To run or edit them, open the file in Excel or the full office editor.')}</Callout>
          <select className="input sm" value={cur} onChange={(e) => setCur(+e.target.value)} aria-label={t('Module')}>{mods.map((m, i) => <option key={i} value={i}>{m.name} ({t(m.kind === 'class' ? 'class module' : m.kind === 'document' ? 'document module' : 'module')})</option>)}</select>
          <pre className="of-sheet-code">{mods[cur].code || t('(empty)')}</pre>
        </div>
      )}
    </Sheet>
  );
}

export function StatsDialog({ book, onClose }) {
  let cells = 0; let formulas = 0; let notes = 0;
  for (const s of book.sheets) for (const c of s.cells.values()) { if (c.v != null || c.f) cells++; if (c.f) formulas++; if (c.note) notes++; }
  const charts = book.sheets.reduce((n, s) => n + s.charts.length, 0);
  const sh = book.sheets[book.active];
  const u = book.usedRange(sh);
  const rows = [[t('Sheets'), book.sheets.length], [t('Cells with data'), cells], [t('Formulas'), formulas], [t('Notes'), notes], [t('Charts'), charts], [t('Defined names'), [...book.names.values()].filter((n) => !n.raw).length], [t('Used range of this sheet'), u ? `A1:${colName(u.c2)}${u.r2 + 1}` : '–']];
  return (
    <Modal size="dialog" title={t('Workbook statistics')} onClose={onClose} footer={<Btn className="primary" onClick={onClose}>{t('Close')}</Btn>}>
      <table className="table"><tbody>{rows.map(([k, v]) => <tr key={k}><td>{k}</td><td className="num">{typeof v === 'number' ? v.toLocaleString(document.documentElement.lang || 'en') : v}</td></tr>)}</tbody></table>
    </Modal>
  );
}

/** "Search commands and functions" (Alt+Q in Excel) */
export function CommandSearch({ commands, onRun, onClose }) {
  const [q, setQ] = useState('');
  const [i, setI] = useState(0);
  const list = commands.filter((c) => !q || c.label.toLowerCase().includes(q.toLowerCase())).slice(0, 12);
  return (
    <Modal title={t('Search commands and functions')} onClose={onClose}>
      <div className="stack">
        <div className="input-icon"><Search /><input className="input" autoFocus value={q} onChange={(e) => { setQ(e.target.value); setI(0); }} onKeyDown={(e) => { if (e.key === 'ArrowDown') { e.preventDefault(); setI((x) => Math.min(list.length - 1, x + 1)); } if (e.key === 'ArrowUp') { e.preventDefault(); setI((x) => Math.max(0, x - 1)); } if (e.key === 'Enter' && list[i]) { onClose(); onRun(list[i]); } }} placeholder={t('e.g. merge, freeze, chart, VLOOKUP')} /></div>
        <div className="of-sheet-list">{list.map((c, k) => <button key={c.id} type="button" className={k === i ? 'is-on' : ''} onClick={() => { onClose(); onRun(c); }}>{c.label}{c.hint && <span className="small muted"> – {c.hint}</span>}</button>)}</div>
      </div>
    </Modal>
  );
}

export function UnhideSheetDialog({ book, onApply, onClose }) {
  const hidden = book.sheets.filter((s) => s.hidden);
  const [v, setV] = useState(hidden[0] ? hidden[0].id : '');
  return (
    <Modal size="dialog" title={t('Unhide sheet')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" disabled={!v} onClick={() => { onApply(v); onClose(); }}>{t('Unhide')}</Btn></>}>
      {hidden.length ? <div className="of-sheet-list">{hidden.map((s) => <button key={s.id} type="button" className={v === s.id ? 'is-on' : ''} onClick={() => setV(s.id)}><Keep>{s.name}</Keep></button>)}</div> : <p className="small muted">{t('No hidden sheets.')}</p>}
    </Modal>
  );
}

export { rangeA1, parseRange };
