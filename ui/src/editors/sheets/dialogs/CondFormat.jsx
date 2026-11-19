/* Conditional formatting: quick rules, rule editor and the rules manager. */
import { useState } from 'react';
import { Modal, Btn, Field, Empty } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { ArrowUp, ArrowDown, Pencil, Trash2, Plus, ListChecks } from 'lucide-react';
import { rangeA1, parseRange } from '../formula/refs.js';
import { ColorGrid } from '../menus.jsx';

export const CF_STYLES = [
  { id: 'red', label: 'Light red fill with dark red text', style: { fill: '#FFC7CE', color: '#9C0006' } },
  { id: 'yellow', label: 'Yellow fill with dark yellow text', style: { fill: '#FFEB9C', color: '#9C5700' } },
  { id: 'green', label: 'Green fill with dark green text', style: { fill: '#C6EFCE', color: '#006100' } },
  { id: 'redfill', label: 'Light red fill', style: { fill: '#FFC7CE' } },
  { id: 'redtext', label: 'Red text', style: { color: '#C00000' } },
  { id: 'bold', label: 'Bold', style: { b: true } },
];
const TYPES = [
  ['cell', 'Cell value'], ['text', 'Specific text'], ['dup', 'Duplicate values'], ['unique', 'Unique values'], ['top', 'Top values'], ['bottom', 'Bottom values'],
  ['above', 'Above average'], ['below', 'Below average'], ['blank', 'Blanks'], ['notBlank', 'No blanks'], ['error', 'Errors'], ['formula', 'Formula'], ['scale', 'Colour scale'], ['bar', 'Data bar'],
];
const OPS = [['gt', 'greater than'], ['lt', 'less than'], ['between', 'between'], ['notBetween', 'not between'], ['eq', 'equal to'], ['ne', 'not equal to'], ['ge', 'greater than or equal to'], ['le', 'less than or equal to']];
const TOPS = [['contains', 'containing'], ['notContains', 'not containing'], ['begins', 'beginning with'], ['ends', 'ending with']];

export function describeRule(r, toDisplay) {
  const op = (OPS.find((x) => x[0] === r.op) || [])[1];
  switch (r.type) {
    case 'cell': return `${t('Cell value')} ${t(op || '')} ${toDisplay(r.f1 || '')}${/between/i.test(r.op) ? ` ${t('and')} ${toDisplay(r.f2 || '')}` : ''}`;
    case 'text': return `${t('Text')} ${t((TOPS.find((x) => x[0] === r.op) || [, 'containing'])[1])} „${r.text}“`;
    case 'top': case 'bottom': return `${t(r.type === 'top' ? 'Top' : 'Bottom')} ${r.rank}${r.percent ? ' %' : ''}`;
    case 'formula': return `${t('Formula')}: =${toDisplay(r.formula || '')}`;
    case 'raw': return t('Rule from the file ({kind}) – kept', { kind: r.kind });
    default: return t((TYPES.find((x) => x[0] === r.type) || [, r.type])[1]);
  }
}

function StylePick({ value, onChange }) {
  const [custom, setCustom] = useState(false);
  const cur = CF_STYLES.find((s) => JSON.stringify(s.style) === JSON.stringify(value));
  return (
    <div className="stack">
      <select className="input sm" value={custom ? 'custom' : cur ? cur.id : 'custom'} onChange={(e) => { if (e.target.value === 'custom') setCustom(true); else { setCustom(false); onChange(CF_STYLES.find((s) => s.id === e.target.value).style); } }} aria-label={t('Format with')}>
        {CF_STYLES.map((s) => <option key={s.id} value={s.id}>{t(s.label)}</option>)}
        <option value="custom">{t('Custom format …')}</option>
      </select>
      {(custom || !cur) && (
        <div className="row">
          <Field label={t('Fill')}><ColorGrid value={value && value.fill} onPick={(c) => onChange({ ...(value || {}), fill: c || undefined })} /></Field>
          <Field label={t('Font colour')}><ColorGrid value={value && value.color} onPick={(c) => onChange({ ...(value || {}), color: c || undefined })} noneLabel="Automatic" /></Field>
          <div className="stack">
            <label className="check small"><input type="checkbox" checked={!!(value && value.b)} onChange={(e) => onChange({ ...(value || {}), b: e.target.checked || undefined })} />{t('Bold')}</label>
            <label className="check small"><input type="checkbox" checked={!!(value && value.i)} onChange={(e) => onChange({ ...(value || {}), i: e.target.checked || undefined })} />{t('Italic')}</label>
            <label className="check small"><input type="checkbox" checked={!!(value && value.st)} onChange={(e) => onChange({ ...(value || {}), st: e.target.checked || undefined })} />{t('Strikethrough')}</label>
          </div>
        </div>
      )}
      <div className="of-sheet-sample paper" style={{ background: (value && value.fill) || '#fff', color: (value && value.color) || '#000', fontWeight: value && value.b ? 700 : 400, fontStyle: value && value.i ? 'italic' : 'normal' }}>AaBbCc 123</div>
    </div>
  );
}

/** rule editor (also used for the quick rules) */
export function RuleDialog({ rule, ranges, toDisplay, fromDisplay, onSave, onClose, title }) {
  const [r, setR] = useState(() => {
    const base = { type: 'cell', op: 'gt', text: '', rank: 10, percent: false, style: CF_STYLES[0].style, colors: ['#F8696B', '#FFEB84', '#63BE7B'], color: '#638EC6', ...rule };
    return { ...base, f1: rule && rule.f1 ? toDisplay(rule.f1) : '', f2: rule && rule.f2 ? toDisplay(rule.f2) : '', formula: rule && rule.formula ? toDisplay(rule.formula) : '' };
  });
  const [where, setWhere] = useState((rule && rule.ranges ? rule.ranges : ranges).map(rangeA1).join(' '));
  const set = (p) => setR((x) => ({ ...x, ...p }));
  const save = () => {
    const rg = where.split(/[\s,;]+/).map((x) => parseRange(x)).filter(Boolean).map(({ r1, c1, r2, c2 }) => ({ r1, c1, r2, c2 }));
    if (!rg.length) return;
    const out = { ...r, ranges: rg, f1: r.f1 !== '' ? fromDisplay(String(r.f1).replace(/^=/, '')) : '', f2: r.f2 !== '' ? fromDisplay(String(r.f2).replace(/^=/, '')) : '', formula: r.formula ? fromDisplay(String(r.formula).replace(/^=/, '')) : '' };
    delete out._raw; delete out._rawSig;
    onSave(out); onClose();
  };
  const needsStyle = !['scale', 'bar', 'raw'].includes(r.type);
  return (
    <Modal size="wide" title={title || t('Conditional formatting rule')} onClose={onClose} footer={<><Btn className="ghost" onClick={onClose}>{t('Cancel')}</Btn><Btn className="primary" onClick={save}>{t('OK')}</Btn></>}>
      <div className="stack">
        <Field label={t('Applies to')}><input className="input sm mono" value={where} onChange={(e) => setWhere(e.target.value)} /></Field>
        <div className="row">
          <Field label={t('Format cells with')}><select className="input sm" value={r.type} onChange={(e) => set({ type: e.target.value })}>{TYPES.map(([id, l]) => <option key={id} value={id}>{t(l)}</option>)}</select></Field>
          {r.type === 'cell' && <Field label={t('Condition')}><select className="input sm" value={r.op} onChange={(e) => set({ op: e.target.value })}>{OPS.map(([id, l]) => <option key={id} value={id}>{t(l)}</option>)}</select></Field>}
          {r.type === 'cell' && <Field label={t('Value')}><input className="input sm" value={r.f1} onChange={(e) => set({ f1: e.target.value })} autoFocus /></Field>}
          {r.type === 'cell' && /between/i.test(r.op) && <Field label={t('and')}><input className="input sm" value={r.f2} onChange={(e) => set({ f2: e.target.value })} /></Field>}
          {r.type === 'text' && <Field label={t('Condition')}><select className="input sm" value={r.op || 'contains'} onChange={(e) => set({ op: e.target.value })}>{TOPS.map(([id, l]) => <option key={id} value={id}>{t(l)}</option>)}</select></Field>}
          {r.type === 'text' && <Field label={t('Text')}><input className="input sm" value={r.text} onChange={(e) => set({ text: e.target.value })} autoFocus /></Field>}
          {(r.type === 'top' || r.type === 'bottom') && <Field label={t('Count')}><input className="input sm" type="number" min={1} value={r.rank} onChange={(e) => set({ rank: Math.max(1, +e.target.value || 1) })} /></Field>}
          {(r.type === 'top' || r.type === 'bottom') && <label className="check small"><input type="checkbox" checked={!!r.percent} onChange={(e) => set({ percent: e.target.checked })} />{t('% of the selected range')}</label>}
          {r.type === 'formula' && <Field label={t('Formula (TRUE = format)')} hint={t('Write it for the top-left cell of the range, e.g. =$E2="Done"')}><input className="input sm mono" value={r.formula ? `=${String(r.formula).replace(/^=/, '')}` : ''} onChange={(e) => set({ formula: e.target.value.replace(/^=/, '') })} autoFocus /></Field>}
        </div>
        {r.type === 'scale' && <div className="row">{r.colors.map((c, i) => <Field key={i} label={i === 0 ? t('Minimum') : i === r.colors.length - 1 ? t('Maximum') : t('Midpoint')}><ColorGrid value={c} allowNone={false} onPick={(x) => set({ colors: r.colors.map((y, j) => (j === i ? x : y)) })} /></Field>)}
          <label className="check small"><input type="checkbox" checked={r.colors.length === 3} onChange={(e) => set({ colors: e.target.checked ? [r.colors[0], '#FFEB84', r.colors[r.colors.length - 1]] : [r.colors[0], r.colors[r.colors.length - 1]], stops: null })} />{t('3-colour scale')}</label></div>}
        {r.type === 'bar' && <Field label={t('Bar colour')}><ColorGrid value={r.color} allowNone={false} onPick={(x) => set({ color: x })} /></Field>}
        {needsStyle && <Field label={t('Format')}><StylePick value={r.style} onChange={(st) => set({ style: st })} /></Field>}
      </div>
    </Modal>
  );
}

export function RulesManager({ sheet, sel, toDisplay, onChange, onEdit, onNew, onClose }) {
  const [scope, setScope] = useState('sel');
  const inter = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
  const list = sheet.cf.map((r, i) => ({ r, i })).filter(({ r }) => scope === 'sheet' || r.ranges.some((g) => sel.ranges.some((s) => inter(g, s))));
  const move = (i, d) => { const cf = sheet.cf.slice(); const j = i + d; if (j < 0 || j >= cf.length) return; [cf[i], cf[j]] = [cf[j], cf[i]]; onChange(cf.map((x, k) => ({ ...x, priority: k + 1 }))); };
  return (
    <Modal size="wide" title={t('Conditional formatting rules')} onClose={onClose} footer={<><Btn icon={Plus} onClick={onNew}>{t('New rule …')}</Btn><div className="grow" /><Btn className="primary" onClick={onClose}>{t('Close')}</Btn></>}>
      <div className="stack">
        <select className="input sm" value={scope} onChange={(e) => setScope(e.target.value)} aria-label={t('Show rules for')}><option value="sel">{t('Current selection')}</option><option value="sheet">{t('This sheet')}</option></select>
        {!list.length ? <Empty icon={ListChecks} title={t('No rules yet')} text={t('Add a rule to colour cells by their values.')} action={<Btn icon={Plus} onClick={onNew}>{t('New rule …')}</Btn>} /> : (
          <table className="table">
            <thead><tr><th>{t('Rule (in order of priority)')}</th><th>{t('Format')}</th><th>{t('Applies to')}</th><th>{t('Stop if true')}</th><th /></tr></thead>
            <tbody>{list.map(({ r, i }) => (
              <tr key={r.id || i}>
                <td>{describeRule(r, toDisplay)}</td>
                <td><span className="of-sheet-cfprev" style={{ background: r.type === 'scale' ? `linear-gradient(90deg, ${(r.colors || []).join(',')})` : r.type === 'bar' ? `linear-gradient(90deg, ${r.color} 60%, transparent 60%)` : (r.style && r.style.fill) || '#fff', color: (r.style && r.style.color) || '#000', fontWeight: r.style && r.style.b ? 700 : 400 }}>AaBbCc</span></td>
                <td className="mono small">{r.ranges.map(rangeA1).join(' ')}</td>
                <td><input type="checkbox" checked={!!r.stop} aria-label={t('Stop if true')} onChange={(e) => onChange(sheet.cf.map((x, k) => (k === i ? { ...x, stop: e.target.checked } : x)))} /></td>
                <td className="row-tight">
                  <Btn className="xs ghost icon" icon={ArrowUp} title={t('Move up')} onClick={() => move(i, -1)} />
                  <Btn className="xs ghost icon" icon={ArrowDown} title={t('Move down')} onClick={() => move(i, 1)} />
                  <Btn className="xs ghost icon" icon={Pencil} title={t('Edit rule')} disabled={r.type === 'raw'} onClick={() => onEdit(i)} />
                  <Btn className="xs ghost icon" icon={Trash2} title={t('Delete rule')} onClick={() => onChange(sheet.cf.filter((_, k) => k !== i))} />
                </td>
              </tr>
            ))}</tbody>
          </table>
        )}
      </div>
    </Modal>
  );
}
