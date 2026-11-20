/*
 * Word dialogs that load on demand (lazy sub-chunk docs-math.js – see lazy.js wordPack()): paragraph (indents and
 * spacing, line and page breaks, borders and shading, tabs), font (effects, character spacing, scale, position),
 * table properties (table, row, column, cell, cell margins), borders and shading, sort.
 * Every dialog gets the current attrs and returns a patch; the editor applies it (extensions/table.js, wordfmt.js).
 */
import { useState } from 'react';
import { Modal, Btn, Field, Tabs, Switch } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';

const num = (v) => (v === '' || v == null || Number.isNaN(Number(v)) ? null : Number(v));
const dot = (v) => String(v).replace(',', '.');
function NumIn({ label, value, onChange, hint, step }) {
  return <Field label={label} hint={hint}><input className="input" inputMode="decimal" step={step} value={value ?? ''} onChange={(e) => onChange(dot(e.target.value))} aria-label={label} /></Field>;
}
function Sel({ label, value, onChange, options }) {
  return <Field label={label}><select className="input" value={value ?? ''} onChange={(e) => onChange(e.target.value)} aria-label={label}>{options.map(([v, l]) => <option key={v} value={v}>{t(l)}</option>)}</select></Field>;
}
const Check = ({ label, checked, onChange }) => <label className="row-tight small"><Switch checked={!!checked} onChange={onChange} label={t(label)} /><span>{t(label)}</span></label>;
const footer = (onClose, onSave) => <><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" onClick={onSave}>OK</Btn></>;

const BORDER_STYLES = [['single', 'Single line'], ['double', 'Double line'], ['dotted', 'Dotted'], ['dashed', 'Dashed'], ['thick', 'Thick']];
const WIDTHS = [['2', '¼ pt'], ['4', '½ pt'], ['6', '¾ pt'], ['8', '1 pt'], ['12', '1½ pt'], ['18', '2¼ pt'], ['24', '3 pt']];
const LEADERS = [['', 'None'], ['dot', '.......'], ['hyphen', '-------'], ['underscore', '______']];

/* ---------------- paragraph ---------------- */
export function ParagraphDialog({ attrs = {}, onSave, onClose }) {
  const [tab, setTab] = useState('indents');
  const rule = attrs.lineRule || (attrs.lineHeight ? 'multiple' : 'auto');
  const [a, setA] = useState({
    textAlign: attrs.textAlign || 'left', indent: attrs.indent ?? '', rightIndent: attrs.rightIndent ?? '',
    special: attrs.firstLine > 0 ? 'first' : attrs.firstLine < 0 ? 'hanging' : 'none', by: attrs.firstLine ? Math.abs(attrs.firstLine) : '',
    spaceBefore: attrs.spaceBefore ?? '', spaceAfter: attrs.spaceAfter ?? '', rule, at: rule === 'exact' || rule === 'atLeast' ? attrs.lineExact : attrs.lineHeight ?? '',
    keepNext: !!attrs.keepNext, keepLines: !!attrs.keepLines, widow: attrs.widow !== false, outline: attrs.outline != null ? String(attrs.outline) : '',
    border: attrs.pborders ? 'box' : 'none', shading: attrs.pshading || '', tabs: (attrs.tabs || []).map((x) => ({ ...x })),
  });
  const set = (k) => (v) => setA((x) => ({ ...x, [k]: v }));
  const save = () => {
    const fl = a.special === 'none' ? null : (num(a.by) || 0) * (a.special === 'hanging' ? -1 : 1);
    const line = a.rule === 'multiple' ? { lineHeight: num(a.at), lineRule: null, lineExact: null } : a.rule === 'auto' ? { lineHeight: null, lineRule: null, lineExact: null } : { lineHeight: null, lineRule: a.rule, lineExact: num(a.at) };
    const b = 'single:4:auto:4';
    onSave({
      textAlign: a.textAlign === 'left' ? null : a.textAlign, indent: num(a.indent), rightIndent: num(a.rightIndent), firstLine: fl || null,
      spaceBefore: num(a.spaceBefore), spaceAfter: num(a.spaceAfter), ...line,
      keepNext: a.keepNext || null, keepLines: a.keepLines || null, widow: a.widow ? null : false, outline: a.outline === '' ? null : Number(a.outline),
      pborders: a.border === 'none' ? null : a.border === 'box' ? { top: b, left: b, bottom: b, right: b } : a.border === 'bottom' ? { bottom: b } : { top: b, bottom: b },
      pshading: a.shading || null, tabs: a.tabs.filter((x) => num(x.pos) != null).map((x) => ({ pos: num(x.pos), align: x.align || 'left', leader: x.leader || null })),
    });
  };
  return (
    <Modal title="Paragraph" size="lg" onClose={onClose} footer={footer(onClose, save)}>
      <Tabs tabs={[{ id: 'indents', label: t('Indents and spacing') }, { id: 'breaks', label: t('Line and page breaks') }, { id: 'border', label: t('Borders and shading') }, { id: 'tabs', label: t('Tabs') }]} value={tab} onChange={setTab} />
      {tab === 'indents' && (
        <div className="grid grid-2">
          <Sel label="Alignment" value={a.textAlign} onChange={set('textAlign')} options={[['left', 'Left'], ['center', 'Centred'], ['right', 'Right'], ['justify', 'Justified']]} />
          <Sel label="Outline level" value={a.outline} onChange={set('outline')} options={[['', 'Body text'], ['6', 'Level 7'], ['7', 'Level 8'], ['8', 'Level 9']]} />
          <NumIn label={t('Indent left (pt)')} value={a.indent} onChange={set('indent')} />
          <NumIn label={t('Indent right (pt)')} value={a.rightIndent} onChange={set('rightIndent')} />
          <Sel label="Special" value={a.special} onChange={set('special')} options={[['none', '(none)'], ['first', 'First line'], ['hanging', 'Hanging']]} />
          <NumIn label={t('By (pt)')} value={a.by} onChange={set('by')} />
          <NumIn label={t('Spacing before (pt)')} value={a.spaceBefore} onChange={set('spaceBefore')} />
          <NumIn label={t('Spacing after (pt)')} value={a.spaceAfter} onChange={set('spaceAfter')} />
          <Sel label="Line spacing" value={a.rule} onChange={set('rule')} options={[['auto', 'Single'], ['multiple', 'Multiple'], ['atLeast', 'At least'], ['exact', 'Exactly']]} />
          {a.rule !== 'auto' && <NumIn label={a.rule === 'multiple' ? t('At (× lines)') : t('At (pt)')} value={a.at} onChange={set('at')} />}
        </div>
      )}
      {tab === 'breaks' && (
        <div className="stack">
          <Check label="Widow/orphan control" checked={a.widow} onChange={set('widow')} />
          <Check label="Keep with next" checked={a.keepNext} onChange={set('keepNext')} />
          <Check label="Keep lines together" checked={a.keepLines} onChange={set('keepLines')} />
        </div>
      )}
      {tab === 'border' && (
        <div className="grid grid-2">
          <Sel label="Border" value={a.border} onChange={set('border')} options={[['none', 'None'], ['box', 'Box'], ['bottom', 'Bottom line'], ['topbottom', 'Top and bottom']]} />
          <Field label="Shading"><input type="color" className="input" value={a.shading || '#ffffff'} onChange={(e) => set('shading')(e.target.value === '#ffffff' ? '' : e.target.value)} aria-label={t('Shading')} /></Field>
        </div>
      )}
      {tab === 'tabs' && (
        <div className="stack">
          {a.tabs.map((x, i) => (
            <div key={i} className="row-tight">
              <input className="input" style={{ width: 90 }} inputMode="decimal" value={x.pos ?? ''} placeholder="pt" aria-label={t('Tab stop position (pt)')} onChange={(e) => set('tabs')(a.tabs.map((y, j) => (j === i ? { ...y, pos: dot(e.target.value) } : y)))} />
              <select className="input" value={x.align || 'left'} aria-label={t('Alignment')} onChange={(e) => set('tabs')(a.tabs.map((y, j) => (j === i ? { ...y, align: e.target.value } : y)))}>{[['left', 'Left'], ['center', 'Centred'], ['right', 'Right'], ['decimal', 'Decimal']].map(([v, l]) => <option key={v} value={v}>{t(l)}</option>)}</select>
              <select className="input" value={x.leader || ''} aria-label={t('Leader')} onChange={(e) => set('tabs')(a.tabs.map((y, j) => (j === i ? { ...y, leader: e.target.value || null } : y)))}>{LEADERS.map(([v, l]) => <option key={v} value={v}>{t(l)}</option>)}</select>
              <Btn className="sm ghost" onClick={() => set('tabs')(a.tabs.filter((_, j) => j !== i))}>Remove</Btn>
            </div>
          ))}
          <Btn className="sm" onClick={() => set('tabs')([...a.tabs, { pos: '', align: 'left', leader: null }])}>Add tab stop</Btn>
        </div>
      )}
    </Modal>
  );
}

/* ---------------- font (character) ---------------- */
export function FontDialog({ attrs = {}, marks = {}, onSave, onClose }) {
  const [tab, setTab] = useState('font');
  const [a, setA] = useState({ strike: !!marks.strike, dstrike: !!attrs.dstrike, sup: !!marks.superscript, sub: !!marks.subscript, caps: attrs.caps || '', hidden: !!attrs.hidden, letterSpacing: attrs.letterSpacing ?? '', scale: attrs.scale ?? 100, position: attrs.position ?? '', lang: attrs.lang || '' });
  const set = (k) => (v) => setA((x) => ({ ...x, [k]: v }));
  const save = () => onSave({ marks: { strike: a.strike, superscript: a.sup, subscript: a.sub }, style: { dstrike: a.dstrike || null, caps: a.caps || null, hidden: a.hidden || null, letterSpacing: num(a.letterSpacing), scale: num(a.scale) === 100 ? null : num(a.scale), position: num(a.position), lang: a.lang || null } });
  return (
    <Modal title="Font" onClose={onClose} footer={footer(onClose, save)}>
      <Tabs tabs={[{ id: 'font', label: t('Effects') }, { id: 'adv', label: t('Advanced') }]} value={tab} onChange={setTab} />
      {tab === 'font' ? (
        <div className="grid grid-2">
          <Check label="Strikethrough" checked={a.strike} onChange={set('strike')} />
          <Check label="Double strikethrough" checked={a.dstrike} onChange={set('dstrike')} />
          <Check label="Superscript" checked={a.sup} onChange={(v) => setA((x) => ({ ...x, sup: v, sub: v ? false : x.sub }))} />
          <Check label="Subscript" checked={a.sub} onChange={(v) => setA((x) => ({ ...x, sub: v, sup: v ? false : x.sup }))} />
          <Check label="Small caps" checked={a.caps === 'small'} onChange={(v) => set('caps')(v ? 'small' : '')} />
          <Check label="All caps" checked={a.caps === 'all'} onChange={(v) => set('caps')(v ? 'all' : '')} />
          <Check label="Hidden" checked={a.hidden} onChange={set('hidden')} />
        </div>
      ) : (
        <div className="grid grid-2">
          <NumIn label={t('Scale (%)')} value={a.scale} onChange={set('scale')} />
          <NumIn label={t('Spacing (pt, + expanded / − condensed)')} value={a.letterSpacing} onChange={set('letterSpacing')} />
          <NumIn label={t('Position (pt, + raised / − lowered)')} value={a.position} onChange={set('position')} />
          <Sel label="Language" value={a.lang} onChange={set('lang')} options={[['', '(document language)'], ['de-DE', 'German'], ['de-CH', 'German (Switzerland)'], ['de-AT', 'German (Austria)'], ['en-US', 'English (US)'], ['en-GB', 'English (UK)'], ['fr-FR', 'French'], ['es-ES', 'Spanish'], ['it-IT', 'Italian'], ['nl-NL', 'Dutch']]} />
        </div>
      )}
    </Modal>
  );
}

/* ---------------- table properties ---------------- */
const twPt = (tw) => (tw == null ? '' : Math.round(tw / 2) / 10);
const ptTw = (v) => (num(v) == null ? null : Math.round(num(v) * 20));
export function TablePropsDialog({ table = {}, row = {}, cell = {}, colWidth, onSave, onClose }) {
  const [tab, setTab] = useState('table');
  const w = table.width || null;
  const cm = table.cellMargins || {};
  const [a, setA] = useState({
    wType: w ? w.type : 'auto', wVal: w ? (w.type === 'pct' ? w.w / 50 : twPt(w.w)) : '', align: table.align || 'left', indent: twPt(table.indent), layout: table.layout === 'fixed' ? 'fixed' : 'auto',
    rowH: twPt(row.height), rule: row.heightRule || 'atLeast', cantSplit: !!row.cantSplit, header: !!row.header,
    col: colWidth != null ? Math.round(colWidth * 0.75 * 10) / 10 : '', vAlign: cell.vAlign || 'top', textDir: cell.textDir || '', noWrap: !!cell.noWrap,
    mt: twPt(cm.top ?? 0), mb: twPt(cm.bottom ?? 0), ml: twPt(cm.left ?? 108), mr: twPt(cm.right ?? 108),
  });
  const set = (k) => (v) => setA((x) => ({ ...x, [k]: v }));
  const save = () => onSave({
    table: { width: a.wType === 'auto' || num(a.wVal) == null ? null : a.wType === 'pct' ? { type: 'pct', w: Math.round(num(a.wVal) * 50) } : { type: 'dxa', w: ptTw(a.wVal) }, align: a.align === 'left' ? null : a.align, indent: a.align === 'left' ? ptTw(a.indent) : null, layout: a.layout === 'fixed' ? 'fixed' : null, cellMargins: { top: ptTw(a.mt) || 0, bottom: ptTw(a.mb) || 0, left: ptTw(a.ml) || 0, right: ptTw(a.mr) || 0 } },
    row: { height: ptTw(a.rowH), heightRule: num(a.rowH) ? (a.rule === 'exact' ? 'exact' : null) : null, cantSplit: a.cantSplit || null },
    header: a.header,
    col: num(a.col) != null ? num(a.col) / 0.75 : null,
    cell: { vAlign: a.vAlign === 'top' ? null : a.vAlign, textDir: a.textDir || null, noWrap: a.noWrap || null },
  });
  return (
    <Modal title="Table properties" size="lg" onClose={onClose} footer={footer(onClose, save)}>
      <Tabs tabs={[{ id: 'table', label: t('Table') }, { id: 'row', label: t('Row') }, { id: 'col', label: t('Column') }, { id: 'cell', label: t('Cell') }]} value={tab} onChange={setTab} />
      {tab === 'table' && (
        <div className="grid grid-2">
          <Sel label="Preferred width" value={a.wType} onChange={set('wType')} options={[['auto', 'Automatic'], ['pct', 'Percent of the text width'], ['dxa', 'Points']]} />
          {a.wType !== 'auto' && <NumIn label={a.wType === 'pct' ? t('Width (%)') : t('Width (pt)')} value={a.wVal} onChange={set('wVal')} />}
          <Sel label="Alignment" value={a.align} onChange={set('align')} options={[['left', 'Left'], ['center', 'Centred'], ['right', 'Right']]} />
          {a.align === 'left' && <NumIn label={t('Indent from left (pt)')} value={a.indent} onChange={set('indent')} />}
          <Sel label="Column widths" value={a.layout} onChange={set('layout')} options={[['auto', 'Automatically resize to fit contents'], ['fixed', 'Fixed column widths']]} />
          <span />
          <NumIn label={t('Cell margin top (pt)')} value={a.mt} onChange={set('mt')} />
          <NumIn label={t('Cell margin bottom (pt)')} value={a.mb} onChange={set('mb')} />
          <NumIn label={t('Cell margin left (pt)')} value={a.ml} onChange={set('ml')} />
          <NumIn label={t('Cell margin right (pt)')} value={a.mr} onChange={set('mr')} />
        </div>
      )}
      {tab === 'row' && (
        <div className="grid grid-2">
          <NumIn label={t('Row height (pt, empty = automatic)')} value={a.rowH} onChange={set('rowH')} />
          <Sel label="Row height is" value={a.rule} onChange={set('rule')} options={[['atLeast', 'At least'], ['exact', 'Exactly']]} />
          <Check label="Allow row to break across pages" checked={!a.cantSplit} onChange={(v) => set('cantSplit')(!v)} />
          <Check label="Repeat as header row at the top of each page" checked={a.header} onChange={set('header')} />
        </div>
      )}
      {tab === 'col' && <div className="grid grid-2"><NumIn label={t('Column width (pt)')} value={a.col} onChange={set('col')} /></div>}
      {tab === 'cell' && (
        <div className="grid grid-2">
          <Sel label="Vertical alignment" value={a.vAlign} onChange={set('vAlign')} options={[['top', 'Top'], ['center', 'Centre'], ['bottom', 'Bottom']]} />
          <Sel label="Text direction" value={a.textDir} onChange={set('textDir')} options={[['', 'Horizontal'], ['btLr', 'Bottom to top'], ['tbRl', 'Top to bottom']]} />
          <Check label="Do not wrap text" checked={a.noWrap} onChange={set('noWrap')} />
        </div>
      )}
    </Modal>
  );
}

/* ---------------- borders and shading (cells or the whole table) ---------------- */
export function BordersDialog({ onSave, onClose }) {
  const [a, setA] = useState({ preset: 'all', style: 'single', sz: '4', color: '#000000', to: 'cells', shading: '' });
  const set = (k) => (v) => setA((x) => ({ ...x, [k]: v }));
  const save = () => {
    const b = a.preset === 'none' ? 'nil' : `${a.style}:${a.sz}:${a.color.replace('#', '').toUpperCase()}:0`;
    const sides = { all: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'], box: ['top', 'left', 'bottom', 'right'], none: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'], inside: ['insideH', 'insideV'], top: ['top'], bottom: ['bottom'], left: ['left'], right: ['right'] }[a.preset];
    const borders = {}; for (const s of sides) borders[s] = b;
    onSave({ to: a.to, borders, preset: a.preset, shading: a.shading || null });
  };
  return (
    <Modal title="Borders and shading" onClose={onClose} footer={footer(onClose, save)}>
      <div className="grid grid-2">
        <Sel label="Setting" value={a.preset} onChange={set('preset')} options={[['all', 'All borders'], ['box', 'Outside borders'], ['inside', 'Inside borders'], ['none', 'No border'], ['top', 'Top border'], ['bottom', 'Bottom border'], ['left', 'Left border'], ['right', 'Right border']]} />
        <Sel label="Apply to" value={a.to} onChange={set('to')} options={[['cells', 'Selected cells'], ['table', 'Table']]} />
        <Sel label="Style" value={a.style} onChange={set('style')} options={BORDER_STYLES} />
        <Sel label="Width" value={a.sz} onChange={set('sz')} options={WIDTHS} />
        <Field label="Colour"><input type="color" className="input" value={a.color} onChange={(e) => set('color')(e.target.value)} aria-label={t('Colour')} /></Field>
        <Field label="Shading"><input type="color" className="input" value={a.shading || '#ffffff'} onChange={(e) => set('shading')(e.target.value === '#ffffff' ? '' : e.target.value)} aria-label={t('Shading')} /></Field>
      </div>
    </Modal>
  );
}

/* ---------------- sort ---------------- */
export function SortDialog({ columns = [], onSave, onClose }) {
  const [a, setA] = useState({ col: '0', type: 'auto', desc: false });
  const set = (k) => (v) => setA((x) => ({ ...x, [k]: v }));
  return (
    <Modal title="Sort" onClose={onClose} footer={footer(onClose, () => onSave({ col: Number(a.col), desc: a.desc, numeric: a.type === 'auto' ? null : a.type === 'number' }))}>
      <div className="grid grid-2">
        <Sel label="Sort by" value={a.col} onChange={set('col')} options={columns.map((c, i) => [String(i), c || t('Column {n}', { n: i + 1 })])} />
        <Sel label="Type" value={a.type} onChange={set('type')} options={[['auto', 'Automatic'], ['text', 'Text'], ['number', 'Number']]} />
        <Sel label="Order" value={a.desc ? 'desc' : 'asc'} onChange={(v) => set('desc')(v === 'desc')} options={[['asc', 'Ascending'], ['desc', 'Descending']]} />
      </div>
      <p className="small muted">{t('Header rows stay at the top.')}</p>
    </Modal>
  );
}

export default { ParagraphDialog, FontDialog, TablePropsDialog, BordersDialog, SortDialog };
