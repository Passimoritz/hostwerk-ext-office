/* Dialogs: header & footer, custom slide size, chart data. */
import { useState } from 'react';
import { Modal, Btn, Field, Switch } from '@hostwerk/ui';
import { Plus, Trash2 } from 'lucide-react';
import { t } from '@hostwerk/i18n';
import { emuToCm, cmToEmu, emuToIn, inToEmu } from './units.js';
import { Raw } from '../_shared/keep.jsx';

export function FooterDialog({ footer, onClose, onApply }) {
  const [f, setF] = useState({ text: '', showFooter: false, showNumber: false, showDate: false, date: '', hideOnTitle: true, ...(footer || {}) });
  return (
    <Modal title={t('Header & footer')} onClose={onClose} footer={<><Btn onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply(f); onClose(); }}>{t('Apply to all')}</Btn></>}>
      <div className="stack">
        <Switch checked={f.showDate} onChange={(v) => setF({ ...f, showDate: v })} label={t('Date and time')} />
        {f.showDate && <Field label={t('Fixed date (empty = today)')}><input className="input" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} /></Field>}
        <Switch checked={f.showNumber} onChange={(v) => setF({ ...f, showNumber: v })} label={t('Slide number')} />
        <Switch checked={f.showFooter} onChange={(v) => setF({ ...f, showFooter: v })} label={t('Footer')} />
        {f.showFooter && <Field label={t('Footer text')}><input className="input" value={f.text} onChange={(e) => setF({ ...f, text: e.target.value })} /></Field>}
        <Switch checked={f.hideOnTitle} onChange={(v) => setF({ ...f, hideOnTitle: v })} label={t('Don’t show on title slide')} />
      </div>
    </Modal>
  );
}

export function SizeDialog({ size, units = 'cm', onClose, onApply }) {
  const toU = (e) => +(units === 'in' ? emuToIn(e) : emuToCm(e)).toFixed(2);
  const fromU = (v) => (units === 'in' ? inToEmu(v) : cmToEmu(v));
  const [w, setW] = useState(String(toU(size.w)));
  const [h, setH] = useState(String(toU(size.h)));
  const [how, setHow] = useState('fit');
  const nw = parseFloat(w.replace(',', '.')); const nh = parseFloat(h.replace(',', '.'));
  const ok = nw > 0.5 && nh > 0.5 && nw < 150 && nh < 150;
  return (
    <Modal title={t('Custom slide size')} onClose={onClose} footer={<><Btn onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" disabled={!ok} onClick={() => { onApply({ w: fromU(nw), h: fromU(nh) }, how); onClose(); }}>{t('Apply')}</Btn></>}>
      <div className="stack">
        <div className="grid grid-2">
          <Field label={t('Width ({unit})', { unit: units })}><input className="input" inputMode="decimal" value={w} onChange={(e) => setW(e.target.value)} /></Field>
          <Field label={t('Height ({unit})', { unit: units })}><input className="input" inputMode="decimal" value={h} onChange={(e) => setH(e.target.value)} /></Field>
        </div>
        <Field label={t('Content')}>
          <select className="input" value={how} onChange={(e) => setHow(e.target.value)}><option value="fit">{t('Ensure fit (scale down, keep proportions)')}</option><option value="stretch">{t('Maximize (stretch to the new size)')}</option></select>
        </Field>
      </div>
    </Modal>
  );
}

export function ChartDataDialog({ chart, onClose, onApply }) {
  const [c, setC] = useState(() => JSON.parse(JSON.stringify(chart)));
  const n = Math.max(c.categories.length, ...c.series.map((s) => s.values.length));
  const setCat = (i, v) => setC({ ...c, categories: Array.from({ length: n }, (_, k) => (k === i ? v : c.categories[k] ?? '')) });
  const setVal = (si, i, v) => setC({ ...c, series: c.series.map((s, k) => (k !== si ? s : { ...s, values: Array.from({ length: n }, (_, j) => (j === i ? v : s.values[j] ?? 0)) })) });
  const num = (v) => { const x = parseFloat(String(v).replace(',', '.')); return Number.isFinite(x) ? x : 0; };
  return (
    <Modal size="lg" title={t('Chart data')} onClose={onClose} footer={<><Btn onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={() => { onApply({ ...c, series: c.series.map((s) => ({ ...s, values: s.values.map(num) })) }); onClose(); }}>{t('Apply')}</Btn></>}>
      <div className="of-slides-chartdata">
        <table className="table">
          <thead><tr><th>{t('Category')}</th>{c.series.map((s, si) => <th key={si}><div className="row-tight"><input className="input" aria-label={t('Series name')} value={s.name} onChange={(e) => setC({ ...c, series: c.series.map((x, k) => (k === si ? { ...x, name: e.target.value } : x)) })} />{c.series.length > 1 && <button type="button" className="btn xs ghost icon" title={t('Remove series')} aria-label={t('Remove series')} onClick={() => setC({ ...c, series: c.series.filter((_, k) => k !== si) })}><Trash2 size={12} /></button>}</div></th>)}</tr></thead>
          <tbody>
            {Array.from({ length: n }, (_, i) => (
              <tr key={i}>
                <td><div className="row-tight"><input className="input" aria-label={t('Category')} value={c.categories[i] ?? ''} onChange={(e) => setCat(i, e.target.value)} />
                  <button type="button" className="btn xs ghost icon" title={t('Remove row')} aria-label={t('Remove row')} onClick={() => setC({ ...c, categories: c.categories.filter((_, k) => k !== i), series: c.series.map((s) => ({ ...s, values: s.values.filter((_, k) => k !== i) })) })}><Trash2 size={12} /></button></div></td>
                {c.series.map((s, si) => <td key={si} className="num"><Raw as="input" className="input" inputMode="decimal" aria-label={`${s.name} ${c.categories[i] ?? ''}`} value={s.values[i] ?? ''} onChange={(e) => setVal(si, i, e.target.value)} /></td>)}
              </tr>
            ))}
          </tbody>
        </table>
        <div className="row">
          <Btn className="sm" icon={Plus} onClick={() => setC({ ...c, categories: [...c.categories, `${t('Category')} ${n + 1}`], series: c.series.map((s) => ({ ...s, values: [...s.values, 0] })) })}>{t('Add row')}</Btn>
          <Btn className="sm" icon={Plus} onClick={() => setC({ ...c, series: [...c.series, { name: `${t('Series')} ${c.series.length + 1}`, values: Array.from({ length: n }, () => 0), color: `accent${(c.series.length % 6) + 1}` }] })}>{t('Add series')}</Btn>
        </div>
      </div>
    </Modal>
  );
}
