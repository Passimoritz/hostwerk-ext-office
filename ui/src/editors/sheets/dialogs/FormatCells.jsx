/* Format cells (Ctrl+1): Number · Alignment · Font · Border · Fill · Protection */
import { useMemo, useState } from 'react';
import { Modal, Btn, Field, Tabs } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { FORMAT_PRESETS, CUSTOM_FORMATS, fmtKind, formatValue, makeNumberFormat, makeCurrencyFormat } from '../format/numfmt.js';
import { ColorGrid } from '../menus.jsx';

const FONTS = ['Calibri', 'Aptos', 'Arial', 'Cambria', 'Candara', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
const LINES = [['thin', 'Thin'], ['hair', 'Hair'], ['dotted', 'Dotted'], ['dashed', 'Dashed'], ['medium', 'Medium'], ['mediumDashed', 'Medium dashed'], ['thick', 'Thick'], ['double', 'Double']];

export function FormatCellsDialog({ book, style, value, initialTab = 'number', onApply, onClose, readOnly }) {
  const loc = book.loc;
  const def = book.props.defaultFont || { font: 'Calibri', size: 11 };
  const s0 = style || {};
  const [tab, setTab] = useState(initialTab);
  const [s, setS] = useState({ ...s0 });
  const [borderOps, setBorderOps] = useState([]);
  const [line, setLine] = useState({ s: 'thin', c: '#000000' });
  const presets = useMemo(() => FORMAT_PRESETS(loc), [loc]);
  const customs = useMemo(() => CUSTOM_FORMATS(loc), [loc]);
  const nf = s.nf == null ? 'General' : s.nf;
  const kind0 = fmtKind(nf);
  const [cat, setCat] = useState(kind0 === 'date' ? 'shortDate' : kind0 === 'time' ? 'time' : kind0 === 'custom' ? 'custom' : presets.find((p) => fmtKind(p.code) === kind0)?.id || 'general');
  const [opt, setOpt] = useState({ decimals: 2, thousands: true, negRed: false, symbol: loc.currency, after: loc.decimal === ',' });
  const set = (p) => setS((x) => ({ ...x, ...p }));
  const preview = (code) => formatValue(typeof value === 'number' ? value : value == null ? 1234.5678 : value, code, loc);
  const applyCat = (id, o = opt) => {
    setCat(id);
    if (id === 'number') set({ nf: makeNumberFormat({ decimals: o.decimals, thousands: o.thousands, negRed: o.negRed }) });
    else if (id === 'currency') set({ nf: makeCurrencyFormat({ symbol: o.symbol, decimals: o.decimals, after: o.after, negRed: o.negRed }) });
    else if (id === 'percent') set({ nf: o.decimals ? `0.${'0'.repeat(o.decimals)}%` : '0%' });
    else if (id === 'scientific') set({ nf: `0.${'0'.repeat(Math.max(1, o.decimals))}E+00` });
    else if (id !== 'custom') { const p = presets.find((x) => x.id === id); if (p) set({ nf: p.code }); }
  };
  const setOpt2 = (p) => { const o = { ...opt, ...p }; setOpt(o); applyCat(cat, o); };
  const dateCodes = customs.filter((c) => fmtKind(c) === 'date');
  const timeCodes = customs.filter((c) => fmtKind(c) === 'time');
  const ok = () => {
    const patch = {};
    for (const k of ['nf', 'h', 'va', 'wrap', 'shrink', 'rot', 'indent', 'font', 'size', 'b', 'i', 'u', 'st', 'color', 'fill', 'unlocked', 'hiddenF']) if (s[k] !== s0[k]) patch[k] = s[k] === '' ? undefined : s[k];
    onApply(patch, borderOps, s.merge !== undefined ? s.merge : null);
    onClose();
  };
  const edge = (e) => setBorderOps((ops) => [...ops, { preset: e, line: { ...line } }]);
  const tabs = [['number', t('Number')], ['align', t('Alignment')], ['font', t('Font')], ['border', t('Border')], ['fill', t('Fill')], ['protect', t('Protection')]].map(([id, label]) => ({ id, label }));
  return (
    <Modal size="wide" title={t('Format cells')} onClose={onClose}
      footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" disabled={readOnly} onClick={ok}>{t('OK')}</Btn></>}>
      <div className="of-sheet-dlgtabs"><Tabs tabs={tabs} value={tab} onChange={setTab} /></div>
      {tab === 'number' && (
        <div className="of-sheet-fmt">
          <div className="of-sheet-list" role="listbox" aria-label={t('Category')}>
            {[...presets, { id: 'custom', label: 'Custom' }].map((p) => <button key={p.id} type="button" role="option" aria-selected={cat === p.id} className={cat === p.id ? 'is-on' : ''} onClick={() => applyCat(p.id)}>{t(p.label)}</button>)}
          </div>
          <div className="stack grow">
            <div className="of-sheet-sample"><span className="small muted">{t('Sample')}</span><b style={{ color: preview(nf).color || undefined }}>{preview(nf).text || ' '}</b></div>
            {['number', 'currency', 'percent', 'scientific'].includes(cat) && (
              <div className="row">
                <Field label={t('Decimal places')}><input className="input sm" type="number" min={0} max={30} value={opt.decimals} onChange={(e) => setOpt2({ decimals: Math.max(0, Math.min(30, +e.target.value || 0)) })} /></Field>
                {cat === 'number' && <label className="check small"><input type="checkbox" checked={opt.thousands} onChange={(e) => setOpt2({ thousands: e.target.checked })} />{t('Use 1000 separator')}</label>}
                {cat === 'currency' && <Field label={t('Symbol')}><select className="input sm" value={opt.symbol} onChange={(e) => setOpt2({ symbol: e.target.value })}>{['€', '$', '£', '¥', 'CHF', 'Fr.', 'kr', 'zł', '₹'].map((x) => <option key={x}>{x}</option>)}</select></Field>}
                {cat === 'currency' && <label className="check small"><input type="checkbox" checked={opt.after} onChange={(e) => setOpt2({ after: e.target.checked })} />{t('Symbol after the number')}</label>}
                {(cat === 'number' || cat === 'currency') && <label className="check small"><input type="checkbox" checked={opt.negRed} onChange={(e) => setOpt2({ negRed: e.target.checked })} />{t('Negative numbers in red')}</label>}
              </div>
            )}
            {(cat === 'shortDate' || cat === 'longDate') && <Field label={t('Type')}><div className="of-sheet-list compact">{dateCodes.map((c) => <button key={c} type="button" className={nf === c ? 'is-on' : ''} onClick={() => set({ nf: c })}>{formatValue(typeof value === 'number' ? value : 46344.5, c, loc).text}</button>)}</div></Field>}
            {cat === 'time' && <Field label={t('Type')}><div className="of-sheet-list compact">{timeCodes.map((c) => <button key={c} type="button" className={nf === c ? 'is-on' : ''} onClick={() => set({ nf: c })}>{formatValue(typeof value === 'number' ? value : 0.6458, c, loc).text}</button>)}</div></Field>}
            {cat === 'custom' && <Field label={t('Type')}><div className="of-sheet-list compact">{customs.map((c) => <button key={c} type="button" className={nf === c ? 'is-on' : ''} onClick={() => set({ nf: c })}><span className="mono">{c}</span></button>)}</div></Field>}
            <Field label={t('Format code')} hint={t('Excel format codes, e.g. #,##0.00;[Red]-#,##0.00 or dd.mm.yyyy')}>
              <input className="input mono" value={typeof nf === 'number' ? (nf === 14 ? loc.shortDate : loc.dateTime) : nf} onChange={(e) => { setCat('custom'); set({ nf: e.target.value || 'General' }); }} aria-label={t('Format code')} />
            </Field>
          </div>
        </div>
      )}
      {tab === 'align' && (
        <div className="stack">
          <div className="row">
            <Field label={t('Horizontal')}><select className="input sm" value={s.h || 'general'} onChange={(e) => set({ h: e.target.value === 'general' ? undefined : e.target.value })}>{[['general', t('General')], ['left', t('Left')], ['center', t('Center')], ['right', t('Right')], ['fill', t('Fill')], ['justify', t('Justify')], ['centerContinuous', t('Center across selection')], ['distributed', t('Distributed')]].map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
            <Field label={t('Vertical')}><select className="input sm" value={s.va || 'bottom'} onChange={(e) => set({ va: e.target.value === 'bottom' ? undefined : e.target.value })}>{[['top', t('Top')], ['center', t('Center')], ['bottom', t('Bottom')], ['justify', t('Justify')], ['distributed', t('Distributed')]].map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
            <Field label={t('Indent')}><input className="input sm" type="number" min={0} max={15} value={s.indent || 0} onChange={(e) => set({ indent: +e.target.value || undefined })} /></Field>
            <Field label={t('Rotation (degrees)')}><input className="input sm" type="number" min={-90} max={90} value={s.rot ? (s.rot > 90 && s.rot !== 255 ? -(s.rot - 90) : s.rot === 255 ? 0 : s.rot) : 0} onChange={(e) => { const d = Math.max(-90, Math.min(90, +e.target.value || 0)); set({ rot: d ? (d < 0 ? 90 - d : d) : undefined }); }} /></Field>
          </div>
          <label className="check small"><input type="checkbox" checked={!!s.wrap} onChange={(e) => set({ wrap: e.target.checked || undefined })} />{t('Wrap text')}</label>
          <label className="check small"><input type="checkbox" checked={!!s.shrink} onChange={(e) => set({ shrink: e.target.checked || undefined })} />{t('Shrink to fit')}</label>
          <label className="check small"><input type="checkbox" checked={s.rot === 255} onChange={(e) => set({ rot: e.target.checked ? 255 : undefined })} />{t('Vertical text')}</label>
          <label className="check small"><input type="checkbox" checked={!!s.merge} onChange={(e) => set({ merge: e.target.checked })} />{t('Merge cells')}</label>
        </div>
      )}
      {tab === 'font' && (
        <div className="stack">
          <div className="row">
            <Field label={t('Font')}><select className="input sm" value={s.font || def.font} onChange={(e) => set({ font: e.target.value === def.font ? undefined : e.target.value })}>{[...new Set([s.font || def.font, ...FONTS])].map((f) => <option key={f}>{f}</option>)}</select></Field>
            <Field label={t('Size')}><input className="input sm" type="number" min={1} max={409} value={s.size || def.size} onChange={(e) => set({ size: +e.target.value === def.size ? undefined : +e.target.value })} /></Field>
            <Field label={t('Underline')}><select className="input sm" value={s.u === 'double' ? 'double' : s.u ? 'single' : 'none'} onChange={(e) => set({ u: e.target.value === 'none' ? undefined : e.target.value === 'double' ? 'double' : true })}><option value="none">{t('None')}</option><option value="single">{t('Single')}</option><option value="double">{t('Double')}</option></select></Field>
          </div>
          <div className="row">
            <label className="check small"><input type="checkbox" checked={!!s.b} onChange={(e) => set({ b: e.target.checked || undefined })} />{t('Bold')}</label>
            <label className="check small"><input type="checkbox" checked={!!s.i} onChange={(e) => set({ i: e.target.checked || undefined })} />{t('Italic')}</label>
            <label className="check small"><input type="checkbox" checked={!!s.st} onChange={(e) => set({ st: e.target.checked || undefined })} />{t('Strikethrough')}</label>
          </div>
          <Field label={t('Colour')}><ColorGrid value={s.color} onPick={(c) => set({ color: c || undefined })} noneLabel="Automatic" /></Field>
          <div className="of-sheet-sample paper" style={{ fontFamily: s.font || def.font, fontSize: `${s.size || def.size}pt`, fontWeight: s.b ? 700 : 400, fontStyle: s.i ? 'italic' : 'normal', textDecoration: `${s.u ? 'underline' : ''} ${s.st ? 'line-through' : ''}`, color: s.color || '#000' }}>AaBbYyZz 123</div>
        </div>
      )}
      {tab === 'border' && (
        <div className="stack">
          <div className="row">
            {[['none', t('None')], ['outside', t('Outline')], ['inside', t('Inside')], ['all', t('All')]].map(([p, l]) => <Btn key={p} className="sm" onClick={() => edge(p)}>{l}</Btn>)}
          </div>
          <div className="row">
            {[['top', t('Top')], ['bottom', t('Bottom')], ['left', t('Left')], ['right', t('Right')]].map(([p, l]) => <Btn key={p} className="sm ghost" onClick={() => edge(p)}>{l}</Btn>)}
          </div>
          <div className="row">
            <Field label={t('Style')}><select className="input sm" value={line.s} onChange={(e) => setLine({ ...line, s: e.target.value })}>{LINES.map(([v, l]) => <option key={v} value={v}>{t(l)}</option>)}</select></Field>
          </div>
          <Field label={t('Colour')}><ColorGrid value={line.c} onPick={(c) => setLine({ ...line, c: c || '#000000' })} allowNone={false} /></Field>
          {borderOps.length > 0 && <p className="small muted">{t('{n} border changes will be applied.', { n: borderOps.length })}</p>}
        </div>
      )}
      {tab === 'fill' && (
        <div className="stack">
          <Field label={t('Background colour')}><ColorGrid value={s.fill} onPick={(c) => set({ fill: c || undefined })} /></Field>
          <div className="of-sheet-sample paper" style={{ background: s.fill || '#fff' }}>{t('Sample')}</div>
        </div>
      )}
      {tab === 'protect' && (
        <div className="stack">
          <label className="check small"><input type="checkbox" checked={!s.unlocked} onChange={(e) => set({ unlocked: e.target.checked ? undefined : true })} />{t('Locked')}</label>
          <label className="check small"><input type="checkbox" checked={!!s.hiddenF} onChange={(e) => set({ hiddenF: e.target.checked || undefined })} />{t('Hidden')}</label>
          <p className="small muted">{t('Locking cells or hiding formulas has no effect until the sheet is protected (in Excel). The settings are saved in the file.')}</p>
        </div>
      )}
    </Modal>
  );
}
