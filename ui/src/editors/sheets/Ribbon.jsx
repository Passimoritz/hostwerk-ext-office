/* Ribbon with the tabs people know from Excel/Calc: Home, Insert, Page layout, Formulas, Data, Review, View. */
import { useState } from 'react';
import { t } from '@hostwerk/i18n';
import {
  Undo2, Redo2, ClipboardPaste, Scissors, Copy, Paintbrush, Bold, Italic, Underline, Strikethrough, Baseline, PaintBucket, Grid3x3,
  AlignLeft, AlignCenter, AlignRight, AlignVerticalJustifyStart, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd, WrapText, TableCellsMerge,
  IndentIncrease, IndentDecrease, RotateCcw, Percent, Euro, DollarSign, ChartColumn, ChartBar, ChartLine, ChartArea, ChartPie, ChartScatter, Sigma,
  Filter, FilterX, Search, ArrowDownAZ, ArrowUpZA, ArrowDownWideNarrow, Link, StickyNote, MessageSquareX, Printer, Eye, Snowflake, ZoomIn, ZoomOut,
  ListChecks, SquareFunction, Calculator, Plus, Trash2, Palette, DecimalsArrowLeft, DecimalsArrowRight, AArrowUp, AArrowDown, Eraser, Group, Ungroup,
  RefreshCw, CodeXml, PanelTop, BetweenHorizontalStart, BetweenVerticalStart, Table, Columns3, Rows3, FileSpreadsheet, Type, Replace, Lock,
  Image as ImageIcon, Shapes, TableProperties, ChartSpline, Radio, Database, ClipboardList, BellRing, Users, GitCompare, Sparkles, Settings2,
  ChevronUp, Save, CalendarDays, LayoutGrid, Table2, Trash, Pin,
} from 'lucide-react';
import { DropBtn, SplitBtn, pickColor, anchorOf, showMenu } from './menus.jsx';
import { CHART_TYPES } from './chartData.js';
import { CATEGORIES, FUNC_INFO } from './formula/funcinfo.js';
import { localFn } from './formula/i18n.js';
import { withKeys } from '../_shared/keys.js';

const FONTS = ['Calibri', 'Aptos', 'Arial', 'Arial Black', 'Cambria', 'Candara', 'Carlito', 'Comic Sans MS', 'Consolas', 'Courier New', 'DejaVu Sans', 'Georgia', 'Liberation Sans', 'Liberation Serif', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const CHART_ICON = { column: ChartColumn, bar: ChartBar, line: ChartLine, area: ChartArea, pie: ChartPie, doughnut: ChartPie, scatter: ChartScatter };

const Tool = ({ icon: I, label, title, kbd, onClick, active, disabled, wide }) => (
  <button type="button" className={`btn ghost sm${wide ? '' : ' icon'}${active ? ' is-active' : ''}`} title={kbd ? withKeys(title || label, kbd) : title || label} aria-label={title || label} aria-pressed={active === undefined ? undefined : !!active} disabled={disabled} onClick={onClick}>
    {I && <I size={15} />}{wide && label && <span>{label}</span>}
  </button>
);
const Group_ = ({ label, children }) => <div className="of-sheet-rgroup" role="group" aria-label={label}><div className="of-sheet-rgroup-body">{children}</div><div className="of-sheet-rgroup-label">{label}</div></div>;

export function Ribbon({ run, cur, flags }) {
  const [tab0, setTab] = useState('home');
  const [collapsed, setCollapsed] = useState(() => { try { return localStorage.getItem('of-sheet-ribbon') === 'min'; } catch { return false; } });
  const [peek, setPeek] = useState(false);
  const P = flags.plus || {};
  const collapse = (v) => { setCollapsed(v); setPeek(false); try { localStorage.setItem('of-sheet-ribbon', v ? 'min' : ''); } catch { /* private mode */ } };
  const ro = flags.readOnly;
  const s = cur || {};
  const TABS = [['home', t('Home')], ['insert', t('Insert')], ['layout', t('Page layout')], ['formulas', t('Formulas')], ['data', t('Data')], ['review', t('Review')], ['view', t('View')]];
  // contextual tabs (like Excel: they appear while a table, pivot table or object is selected)
  const CTX = [P.inTable && ['ctx-table', t('Table design')], P.inPivot && ['ctx-pivot', t('PivotTable')], P.objSel && ['ctx-obj', P.objSel.kind === 'image' ? t('Picture') : t('Shape format')]].filter(Boolean);
  const tab = tab0.startsWith('ctx-') && !CTX.some((x) => x[0] === tab0) ? 'home' : tab0;
  const shapeItems = () => [['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Oval'], ['triangle', 'Triangle'], ['diamond', 'Diamond'], ['rightArrow', 'Arrow'], ['star5', 'Star'], ['callout', 'Speech bubble'], ['line', 'Line'], '-', ['textbox', 'Text box']].map((x) => (x === '-' ? '-' : { label: t(x[1]), onClick: () => run('shape', x[0]) }));
  const numPresets = flags.formatPresets;
  const borderItems = () => [
    { label: t('Bottom border'), onClick: () => run('border', 'bottom') }, { label: t('Top border'), onClick: () => run('border', 'top') },
    { label: t('Left border'), onClick: () => run('border', 'left') }, { label: t('Right border'), onClick: () => run('border', 'right') }, '-',
    { label: t('No border'), onClick: () => run('border', 'none') }, { label: t('All borders'), onClick: () => run('border', 'all') },
    { label: t('Outside borders'), onClick: () => run('border', 'outside') }, { label: t('Inside borders'), onClick: () => run('border', 'inside') },
    { label: t('Thick box border'), onClick: () => run('border', 'thickBox') }, { label: t('Bottom double border'), onClick: () => run('border', 'bottomDouble') },
    { label: t('Thick bottom border'), onClick: () => run('border', 'thickBottom') }, { label: t('Top and bottom border'), onClick: () => run('border', 'topBottom') }, '-',
    { label: t('Line colour …'), icon: Palette, onClick: () => run('borderColor') },
    { label: t('Line style'), items: [['thin', t('Thin')], ['medium', t('Medium')], ['thick', t('Thick')], ['dashed', t('Dashed')], ['dotted', t('Dotted')], ['double', t('Double')]].map(([id, l]) => ({ label: l, checked: flags.borderLine === id, onClick: () => run('borderStyle', id) })) },
    { label: t('More borders …'), onClick: () => run('formatCells', 'border') },
  ];
  const cfItems = () => [
    { label: t('Highlight cell rules'), items: [
      { label: t('Greater than …'), onClick: () => run('cfQuick', 'gt') }, { label: t('Less than …'), onClick: () => run('cfQuick', 'lt') },
      { label: t('Between …'), onClick: () => run('cfQuick', 'between') }, { label: t('Equal to …'), onClick: () => run('cfQuick', 'eq') },
      { label: t('Text that contains …'), onClick: () => run('cfQuick', 'text') }, { label: t('A date occurring …'), onClick: () => run('cfQuick', 'date') },
      { label: t('Duplicate values …'), onClick: () => run('cfQuick', 'dup') },
    ] },
    { label: t('Top/bottom rules'), items: [
      { label: t('Top 10 items …'), onClick: () => run('cfQuick', 'top') }, { label: t('Top 10 % …'), onClick: () => run('cfQuick', 'topPct') },
      { label: t('Bottom 10 items …'), onClick: () => run('cfQuick', 'bottom') }, { label: t('Above average'), onClick: () => run('cfQuick', 'above') }, { label: t('Below average'), onClick: () => run('cfQuick', 'below') },
    ] },
    { label: t('Data bars'), items: [['#638EC6', t('Blue')], ['#63BE7B', t('Green')], ['#FF555A', t('Red')], ['#FFB628', t('Orange')], ['#008AEF', t('Light blue')], ['#D6007B', t('Purple')]].map(([c, l]) => ({ label: l, onClick: () => run('cfBar', c) })) },
    { label: t('Colour scales'), items: [
      { label: t('Green – yellow – red'), onClick: () => run('cfScale', ['#63BE7B', '#FFEB84', '#F8696B']) }, { label: t('Red – yellow – green'), onClick: () => run('cfScale', ['#F8696B', '#FFEB84', '#63BE7B']) },
      { label: t('White – green'), onClick: () => run('cfScale', ['#FFFFFF', '#63BE7B']) }, { label: t('White – red'), onClick: () => run('cfScale', ['#FFFFFF', '#F8696B']) }, { label: t('Blue – white – red'), onClick: () => run('cfScale', ['#5A8AC6', '#FFFFFF', '#F8696B']) },
    ] },
    { label: t('Icon sets'), items: Object.entries({ '3Arrows': 'Arrows', '3TrafficLights1': 'Traffic lights', '3Flags': 'Flags', '3Symbols': 'Symbols', '3Triangles': 'Triangles', '3Stars': 'Stars', '4Rating': 'Rating', '5Quarters': 'Quarters', '5Arrows': 'Five arrows' }).map(([id, l]) => ({ label: t(l), onClick: () => run('iconSet', id) })) },
    '-',
    { label: t('New rule …'), onClick: () => run('cfNew') },
    { label: t('Clear rules from selected cells'), onClick: () => run('cfClear', 'sel') },
    { label: t('Clear rules from entire sheet'), onClick: () => run('cfClear', 'sheet') },
    { label: t('Manage rules …'), onClick: () => run('cfManage') },
  ];
  const styleItems = () => flags.cellStyles.map((cs) => ({ label: t(cs.label), onClick: () => run('cellStyle', cs.id) }));
  const fnMenu = (cat) => [...FUNC_INFO.values()].filter((f) => f.cat === cat).map((f) => ({ label: localFn(f.name, flags.loc), keep: true, onClick: () => run('insertFn', f.name) }));
  const ins = () => [
    { label: t('Insert cells …'), onClick: () => run('insertCellsDlg') },
    { label: t('Insert sheet rows'), icon: BetweenHorizontalStart, onClick: () => run('insertRows') },
    { label: t('Insert sheet columns'), icon: BetweenVerticalStart, onClick: () => run('insertCols') },
    { label: t('Insert sheet'), icon: FileSpreadsheet, onClick: () => run('insertSheet') },
  ];
  const del = () => [
    { label: t('Delete cells …'), onClick: () => run('deleteCellsDlg') },
    { label: t('Delete sheet rows'), onClick: () => run('deleteRows') },
    { label: t('Delete sheet columns'), onClick: () => run('deleteCols') },
    { label: t('Delete sheet'), danger: true, onClick: () => run('deleteSheet') },
  ];
  const fmtMenu = () => [
    { header: t('Cell size') },
    { label: t('Row height …'), onClick: () => run('rowHeight') }, { label: t('AutoFit row height'), onClick: () => run('autofitRows') },
    { label: t('Column width …'), onClick: () => run('colWidth') }, { label: t('AutoFit column width'), onClick: () => run('autofitCols') },
    { label: t('Default width …'), onClick: () => run('defaultWidth') },
    { header: t('Visibility') },
    { label: t('Hide rows'), onClick: () => run('hide', 'row') }, { label: t('Hide columns'), onClick: () => run('hide', 'col') },
    { label: t('Unhide rows'), onClick: () => run('unhide', 'row') }, { label: t('Unhide columns'), onClick: () => run('unhide', 'col') },
    { label: t('Hide sheet'), onClick: () => run('hideSheet') }, { label: t('Unhide sheet …'), onClick: () => run('unhideSheet') },
    { header: t('Organise sheets') },
    { label: t('Rename sheet'), onClick: () => run('renameSheet') }, { label: t('Duplicate sheet'), onClick: () => run('duplicateSheet') }, { label: t('Tab colour …'), onClick: () => run('tabColor') },
    '-', { label: t('Format cells …'), kbd: 'Ctrl+1', onClick: () => run('formatCells') },
  ];
  const clearMenu = () => [
    { label: t('Clear all'), onClick: () => run('clear', 'all') }, { label: t('Clear formats'), onClick: () => run('clear', 'formats') },
    { label: t('Clear contents'), kbd: 'Del', onClick: () => run('clear', 'contents') }, { label: t('Clear notes'), onClick: () => run('clear', 'notes') }, { label: t('Clear hyperlinks'), onClick: () => run('clear', 'links') },
  ];
  const sortMenu = () => [
    { label: t('Sort A to Z'), icon: ArrowDownAZ, onClick: () => run('sort', 'asc') }, { label: t('Sort Z to A'), icon: ArrowUpZA, onClick: () => run('sort', 'desc') },
    { label: t('Custom sort …'), icon: ArrowDownWideNarrow, onClick: () => run('sortDlg') }, '-',
    { label: t('Filter'), icon: Filter, kbd: 'Ctrl+Shift+L', checked: flags.hasFilter, onClick: () => run('filterToggle') },
    { label: t('Clear filter'), icon: FilterX, disabled: !flags.filtered, onClick: () => run('filterClear') }, { label: t('Reapply'), icon: RefreshCw, disabled: !flags.hasFilter, onClick: () => run('filterReapply') },
  ];
  const findMenu = () => [
    { label: t('Find …'), icon: Search, kbd: 'Ctrl+F', onClick: () => run('find') }, { label: t('Replace …'), icon: Replace, kbd: 'Ctrl+H', onClick: () => run('replace') },
    { label: t('Go to …'), kbd: 'Ctrl+G', onClick: () => run('goto') }, { label: t('Go to special …'), onClick: () => run('gotoSpecial') },
    '-', { label: t('Formulas'), onClick: () => run('selectSpecial', 'formulas') }, { label: t('Notes'), onClick: () => run('selectSpecial', 'notes') }, { label: t('Constants'), onClick: () => run('selectSpecial', 'constants') }, { label: t('Data validation'), onClick: () => run('selectSpecial', 'validation') },
  ];
  const sumMenu = () => [['SUM', t('Sum')], ['AVERAGE', t('Average')], ['COUNT', t('Count numbers')], ['MAX', t('Max')], ['MIN', t('Min')]].map(([f, l]) => ({ label: l, onClick: () => run('autosum', f) })).concat(['-', { label: t('More functions …'), onClick: () => run('insertFunction') }]);
  const fillMenu = () => [{ label: t('Down'), kbd: 'Ctrl+D', onClick: () => run('fillDir', 'down') }, { label: t('Right'), kbd: 'Ctrl+R', onClick: () => run('fillDir', 'right') }, { label: t('Up'), onClick: () => run('fillDir', 'up') }, { label: t('Left'), onClick: () => run('fillDir', 'left') }];
  const mergeMenu = () => [{ label: t('Merge & center'), onClick: () => run('merge', 'center') }, { label: t('Merge across'), onClick: () => run('merge', 'across') }, { label: t('Merge cells'), onClick: () => run('merge', 'plain') }, { label: t('Unmerge cells'), onClick: () => run('unmerge') }];
  const rotMenu = () => [[45, t('Angle counterclockwise')], [-45, t('Angle clockwise')], [255, t('Vertical text')], [90, t('Rotate text up')], [-90, t('Rotate text down')], [0, t('No rotation')]].map(([d, l]) => ({ label: l, onClick: () => run('rotate', d) })).concat(['-', { label: t('Format cell alignment …'), onClick: () => run('formatCells', 'align') }]);
  const pasteMenu = () => [
    { label: t('Paste'), kbd: 'Ctrl+V', onClick: () => run('paste') }, { label: t('Values only'), kbd: 'Ctrl+Shift+V', onClick: () => run('pasteSpecial', { what: 'values' }) },
    { label: t('Formulas only'), onClick: () => run('pasteSpecial', { what: 'formulas' }) }, { label: t('Formats only'), onClick: () => run('pasteSpecial', { what: 'formats' }) },
    { label: t('Transpose'), onClick: () => run('pasteSpecial', { what: 'all', transpose: true }) }, { label: t('Keep column widths'), onClick: () => run('pasteSpecial', { what: 'all', colWidths: true }) },
    { label: t('No borders'), onClick: () => run('pasteSpecial', { what: 'noBorders' }) }, '-', { label: t('Paste special …'), onClick: () => run('pasteSpecialDlg') },
  ];
  const freezeMenu = () => [
    { label: t('Freeze panes'), onClick: () => run('freeze', 'sel') }, { label: t('Freeze top row'), onClick: () => run('freeze', 'row') }, { label: t('Freeze first column'), onClick: () => run('freeze', 'col') },
    { label: t('Unfreeze panes'), disabled: !flags.frozen, onClick: () => run('freeze', 'none') },
  ];
  const numFmtSel = (
    <select className="input sm of-sheet-numsel" aria-label={t('Number format')} title={t('Number format')} disabled={ro} value={flags.fmtKind || ''} onChange={(e) => { if (e.target.value === 'more') run('formatCells', 'number'); else run('numPreset', e.target.value); e.target.blur(); }}>
      {numPresets.map((p) => <option key={p.id} value={p.id}>{t(p.label)}</option>)}
      <option value="custom" disabled>{t('Custom')}</option>
      <option value="more">{t('More number formats …')}</option>
    </select>
  );
  return (
    <div className="of-sheet-ribbon">
      <div className="of-sheet-rtabs" role="tablist" aria-label={t('Ribbon')}>
        <span className="of-sheet-qat" role="group" aria-label={t('Quick access')}>
          <Tool icon={Save} label={t('Save')} kbd="Ctrl+S" disabled={ro} onClick={() => run('save')} />
          <Tool icon={Undo2} label={t('Undo')} kbd="Ctrl+Z" disabled={ro} onClick={() => run('undo')} />
          <Tool icon={Redo2} label={t('Redo')} kbd="Ctrl+Y" disabled={ro} onClick={() => run('redo')} />
        </span>
        {TABS.map(([id, label]) => <button key={id} type="button" role="tab" aria-selected={tab === id && (!collapsed || peek)} className={`of-sheet-rtab${tab === id && (!collapsed || peek) ? ' is-on' : ''}`} onClick={() => { if (collapsed) setPeek(!(peek && tab === id)); setTab(id); }} onDoubleClick={() => collapse(!collapsed)}>{label}</button>)}
        {CTX.map(([id, label]) => <button key={id} type="button" role="tab" aria-selected={tab === id} className={`of-sheet-rtab is-ctx${tab === id ? ' is-on' : ''}`} onClick={() => { if (collapsed) setPeek(true); setTab(id); }}>{label}</button>)}
        <div className="grow" />
        <button type="button" className="btn ghost xs" onClick={() => run('commandSearch')} title={t('Search commands and functions')}><Search size={13} /><span className="muted">{t('Search')}</span></button>
        <button type="button" className="btn ghost xs icon" onClick={() => collapse(!collapsed)} aria-label={collapsed ? t('Pin the ribbon') : t('Collapse the ribbon')} title={withKeys(collapsed ? t('Pin the ribbon') : t('Collapse the ribbon'), 'Ctrl+F1')}>{collapsed ? <Pin size={13} /> : <ChevronUp size={13} />}</button>
      </div>
      {(!collapsed || peek) && <div className={`of-sheet-rbody${collapsed ? ' is-peek' : ''}`} role="toolbar" aria-label={([...TABS, ...CTX].find((x) => x[0] === tab) || TABS[0])[1]}>
        {tab === 'home' && <>
          <Group_ label={t('Undo')}>
            <Tool icon={Undo2} label={t('Undo')} kbd="Ctrl+Z" disabled={ro || !flags.canUndo} onClick={() => run('undo')} />
            <Tool icon={Redo2} label={t('Redo')} kbd="Ctrl+Y" disabled={ro || !flags.canRedo} onClick={() => run('redo')} />
          </Group_>
          <Group_ label={t('Clipboard')}>
            <SplitBtn icon={ClipboardPaste} label={t('Paste')} kbd="Ctrl+V" disabled={ro} onClick={() => run('paste')} items={pasteMenu} />
            <Tool icon={Scissors} label={t('Cut')} kbd="Ctrl+X" disabled={ro} onClick={() => run('cut')} />
            <Tool icon={Copy} label={t('Copy')} kbd="Ctrl+C" onClick={() => run('copy')} />
            <Tool icon={Paintbrush} label={t('Format painter')} active={!!flags.painter} disabled={ro} onClick={() => run('painter')} />
          </Group_>
          <Group_ label={t('Font')}>
            <div className="of-sheet-rrow">
              <select className="input sm of-sheet-font" aria-label={t('Font')} title={t('Font')} disabled={ro} value={s.font || flags.defFont.font} onChange={(e) => run('font', e.target.value)}>
                {[...new Set([s.font || flags.defFont.font, ...FONTS])].map((f) => <option key={f} value={f}>{f}</option>)}
              </select>
              <select className="input sm of-sheet-size" aria-label={t('Font size')} title={t('Font size')} disabled={ro} value={s.size || flags.defFont.size} onChange={(e) => run('size', Number(e.target.value))}>
                {[...new Set([s.size || flags.defFont.size, ...SIZES])].sort((a, b) => a - b).map((n) => <option key={n} value={n}>{n}</option>)}
              </select>
              <Tool icon={AArrowUp} label={t('Increase font size')} disabled={ro} onClick={() => run('grow', 1)} />
              <Tool icon={AArrowDown} label={t('Decrease font size')} disabled={ro} onClick={() => run('grow', -1)} />
            </div>
            <div className="of-sheet-rrow">
              <Tool icon={Bold} label={t('Bold')} kbd="Ctrl+B" active={!!s.b} disabled={ro} onClick={() => run('bold')} />
              <Tool icon={Italic} label={t('Italic')} kbd="Ctrl+I" active={!!s.i} disabled={ro} onClick={() => run('italic')} />
              <Tool icon={Underline} label={t('Underline')} kbd="Ctrl+U" active={!!s.u} disabled={ro} onClick={() => run('underline')} />
              <Tool icon={Strikethrough} label={t('Strikethrough')} kbd="Ctrl+5" active={!!s.st} disabled={ro} onClick={() => run('strike')} />
              <SplitBtn icon={Grid3x3} title={t('Borders')} disabled={ro} onClick={() => run('border', flags.lastBorder || 'bottom')} items={borderItems} />
              <SplitBtn icon={PaintBucket} title={t('Fill colour')} swatch={flags.lastFill} disabled={ro} onClick={() => run('fill', flags.lastFill)} onArrow={(at) => run('pickFill', at)} />
              <SplitBtn icon={Baseline} title={t('Font colour')} swatch={flags.lastColor} disabled={ro} onClick={() => run('color', flags.lastColor)} onArrow={(at) => run('pickColor', at)} />
            </div>
          </Group_>
          <Group_ label={t('Alignment')}>
            <div className="of-sheet-rrow">
              <Tool icon={AlignVerticalJustifyStart} label={t('Top align')} active={s.va === 'top'} disabled={ro} onClick={() => run('valign', 'top')} />
              <Tool icon={AlignVerticalJustifyCenter} label={t('Middle align')} active={s.va === 'center'} disabled={ro} onClick={() => run('valign', 'center')} />
              <Tool icon={AlignVerticalJustifyEnd} label={t('Bottom align')} active={!s.va || s.va === 'bottom'} disabled={ro} onClick={() => run('valign', 'bottom')} />
              <DropBtn icon={RotateCcw} title={t('Orientation')} disabled={ro} items={rotMenu} />
              <Tool icon={WrapText} label={t('Wrap text')} active={!!s.wrap} disabled={ro} onClick={() => run('wrap')} />
            </div>
            <div className="of-sheet-rrow">
              <Tool icon={AlignLeft} label={t('Align left')} active={s.h === 'left'} disabled={ro} onClick={() => run('align', 'left')} />
              <Tool icon={AlignCenter} label={t('Center')} active={s.h === 'center'} disabled={ro} onClick={() => run('align', 'center')} />
              <Tool icon={AlignRight} label={t('Align right')} active={s.h === 'right'} disabled={ro} onClick={() => run('align', 'right')} />
              <Tool icon={IndentDecrease} label={t('Decrease indent')} disabled={ro} onClick={() => run('indent', -1)} />
              <Tool icon={IndentIncrease} label={t('Increase indent')} disabled={ro} onClick={() => run('indent', 1)} />
              <SplitBtn icon={TableCellsMerge} title={t('Merge & center')} disabled={ro} onClick={() => run(flags.merged ? 'unmerge' : 'merge', 'center')} items={mergeMenu} />
            </div>
          </Group_>
          <Group_ label={t('Number')}>
            <div className="of-sheet-rrow">{numFmtSel}</div>
            <div className="of-sheet-rrow">
              <DropBtn icon={flags.currency === '€' ? Euro : DollarSign} title={t('Currency')} disabled={ro} items={() => [['€', t('€ Euro')], ['$', t('$ Dollar')], ['£', t('£ Pound')], ['CHF', t('CHF Franc')], ['acc', t('Accounting')]].map(([c, l]) => ({ label: l, onClick: () => run('currency', c) }))} />
              <Tool icon={Percent} label={t('Percent style')} kbd="Ctrl+Shift+%" disabled={ro} onClick={() => run('numPreset', 'percent')} />
              <Tool label={t('Thousands separator')} disabled={ro} onClick={() => run('comma')} wide />
              <Tool icon={DecimalsArrowRight} label={t('Increase decimal')} disabled={ro} onClick={() => run('decimals', 1)} />
              <Tool icon={DecimalsArrowLeft} label={t('Decrease decimal')} disabled={ro} onClick={() => run('decimals', -1)} />
            </div>
          </Group_>
          <Group_ label={t('Styles')}>
            <DropBtn icon={ListChecks} label={t('Conditional formatting')} disabled={ro} items={cfItems} />
            <Tool icon={Table} label={t('Format as table')} wide disabled={ro} onClick={() => run('tableStylesMenu')} />
            <DropBtn icon={Palette} label={t('Cell styles')} disabled={ro} items={styleItems} />
          </Group_>
          <Group_ label={t('Cells')}>
            <DropBtn icon={Plus} label={t('Insert')} disabled={ro} items={ins} />
            <DropBtn icon={Trash2} label={t('Delete')} disabled={ro} items={del} />
            <DropBtn icon={Table} label={t('Format')} disabled={ro} items={fmtMenu} />
          </Group_>
          <Group_ label={t('Editing')}>
            <SplitBtn icon={Sigma} title={t('AutoSum')} kbd="Alt+=" disabled={ro} onClick={() => run('autosum', 'SUM')} items={sumMenu} />
            <DropBtn icon={ArrowDownWideNarrow} title={t('Fill')} disabled={ro} items={fillMenu} />
            <DropBtn icon={Eraser} title={t('Clear')} disabled={ro} items={clearMenu} />
            <DropBtn icon={Filter} title={t('Sort & filter')} items={sortMenu} />
            <DropBtn icon={Search} title={t('Find & select')} items={findMenu} />
          </Group_>
        </>}
        {tab === 'insert' && <>
          <Group_ label={t('Charts')}>
            {CHART_TYPES.map((c) => { const I = CHART_ICON[c.id] || ChartColumn; return <Tool key={c.id} icon={I} label={t(c.label)} title={t('Insert {type} chart', { type: t(c.label) })} disabled={ro} wide onClick={() => run('chart', c.id)} />; })}
          </Group_>
          <Group_ label={t('Tables')}>
            <Tool icon={TableCellsMerge} label={t('PivotTable')} wide disabled={ro} onClick={() => run('pivotNew')} />
            <Tool icon={Table} label={t('Table')} kbd="Ctrl+T" wide disabled={ro} onClick={() => run('table')} />
            <Tool icon={Filter} label={t('Filter')} wide onClick={() => run('filterToggle')} />
          </Group_>
          <Group_ label={t('Illustrations')}>
            <Tool icon={ImageIcon} label={t('Pictures')} wide disabled={ro} onClick={() => run('image')} />
            <DropBtn icon={Shapes} label={t('Shapes')} disabled={ro} items={shapeItems} />
          </Group_>
          <Group_ label={t('Sparklines')}>
            <Tool icon={ChartSpline} label={t('Sparklines …')} wide disabled={ro} onClick={() => run('sparkline')} />
            <Tool label={t('Clear')} wide disabled={ro} onClick={() => run('sparkClear')} />
          </Group_>
          <Group_ label={t('Hostwerk')}>
            <Tool icon={Radio} label={t('Live data …')} wide disabled={ro} onClick={() => run('liveData')} />
            <Tool icon={ClipboardList} label={t('Form …')} wide disabled={ro} onClick={() => run('form')} />
          </Group_>
          <Group_ label={t('Links')}>
            <Tool icon={Link} label={t('Link …')} wide disabled={ro} onClick={() => run('link')} />
          </Group_>
          <Group_ label={t('Notes')}>
            <Tool icon={StickyNote} label={t('Note …')} wide disabled={ro} onClick={() => run('note')} />
          </Group_>
          <Group_ label={t('Functions')}>
            <Tool icon={SquareFunction} label={t('Function …')} wide disabled={ro} onClick={() => run('insertFunction')} />
          </Group_>
        </>}
        {tab === 'layout' && <>
          <Group_ label={t('Page setup')}>
            <DropBtn label={t('Orientation')} items={() => [{ label: t('Portrait'), checked: flags.orientation !== 'landscape', onClick: () => run('orientation', 'portrait') }, { label: t('Landscape'), checked: flags.orientation === 'landscape', onClick: () => run('orientation', 'landscape') }]} />
            <DropBtn label={t('Scaling')} items={() => [['none', t('No scaling')], ['width', t('Fit all columns on one page')], ['page', t('Fit sheet on one page')]].map(([id, l]) => ({ label: l, checked: flags.fit === id, onClick: () => run('fit', id) }))} />
            <DropBtn label={t('Print area')} items={() => [{ label: t('Set print area'), onClick: () => run('printArea', 'set') }, { label: t('Clear print area'), disabled: !flags.printArea, onClick: () => run('printArea', 'clear') }]} />
          </Group_>
          <Group_ label={t('Sheet options')}>
            <label className="check small"><input type="checkbox" checked={!!flags.printGrid} onChange={() => run('printGrid')} />{t('Print gridlines')}</label>
            <label className="check small"><input type="checkbox" checked={!!flags.printHead} onChange={() => run('printHead')} />{t('Print headings')}</label>
          </Group_>
          <Group_ label={t('Print')}>
            <Tool icon={Printer} label={t('Print …')} kbd="Ctrl+P" wide onClick={() => run('print')} />
          </Group_>
        </>}
        {tab === 'formulas' && <>
          <Group_ label={t('Function library')}>
            <Tool icon={SquareFunction} label={t('Insert function')} wide disabled={ro} onClick={() => run('insertFunction')} />
            <SplitBtn icon={Sigma} label={t('AutoSum')} disabled={ro} onClick={() => run('autosum', 'SUM')} items={sumMenu} />
            {CATEGORIES.slice(0, 7).map((c) => <DropBtn key={c} label={t(c)} disabled={ro} items={() => fnMenu(c)} />)}
          </Group_>
          <Group_ label={t('Defined names')}>
            <Tool label={t('Name manager')} wide onClick={() => run('names')} />
            <Tool label={t('Define name …')} wide disabled={ro} onClick={() => run('defineName')} />
          </Group_>
          <Group_ label={t('Formula auditing')}>
            <Tool icon={CodeXml} label={t('Show formulas')} kbd="Ctrl+`" active={flags.showFormulas} wide onClick={() => run('showFormulas')} />
          </Group_>
          <Group_ label={t('Calculation')}>
            <DropBtn icon={Calculator} label={t('Calculation options')} items={() => [{ label: t('Automatic'), checked: flags.autoCalc, onClick: () => run('calcMode', true) }, { label: t('Manual'), checked: !flags.autoCalc, onClick: () => run('calcMode', false) }]} />
            <Tool icon={RefreshCw} label={t('Calculate now')} kbd="F9" wide onClick={() => run('calcNow')} />
            <Tool icon={Settings2} label={t('Options …')} title={t('Iterative calculation, R1C1, 1904 dates')} wide onClick={() => run('calcOptions')} />
          </Group_>
          {P.aiReady && <Group_ label={t('AI')}>
            <Tool icon={Sparkles} label={t('Write a formula …')} wide disabled={ro} onClick={() => run('ai', 'write')} />
            <Tool label={t('Explain formula')} wide onClick={() => run('ai', 'explain')} />
          </Group_>}
        </>}
        {tab === 'data' && <>
          <Group_ label={t('Sort & filter')}>
            <Tool icon={ArrowDownAZ} label={t('Sort A to Z')} onClick={() => run('sort', 'asc')} disabled={ro} />
            <Tool icon={ArrowUpZA} label={t('Sort Z to A')} onClick={() => run('sort', 'desc')} disabled={ro} />
            <Tool icon={ArrowDownWideNarrow} label={t('Sort …')} wide onClick={() => run('sortDlg')} disabled={ro} />
            <Tool icon={Filter} label={t('Filter')} kbd="Ctrl+Shift+L" active={flags.hasFilter} wide onClick={() => run('filterToggle')} />
            <Tool icon={FilterX} label={t('Clear')} wide disabled={!flags.filtered} onClick={() => run('filterClear')} />
            <Tool icon={RefreshCw} label={t('Reapply')} disabled={!flags.hasFilter} onClick={() => run('filterReapply')} />
          </Group_>
          <Group_ label={t('Data tools')}>
            <Tool icon={Columns3} label={t('Text to columns …')} wide disabled={ro} onClick={() => run('textToColumns')} />
            <Tool icon={Rows3} label={t('Remove duplicates …')} wide disabled={ro} onClick={() => run('removeDuplicates')} />
            <Tool icon={ListChecks} label={t('Data validation …')} wide disabled={ro} onClick={() => run('validation')} />
            {flags.isCsv && <Tool icon={Type} label={t('Text import options …')} wide onClick={() => run('csvOptions')} />}
            <Tool label={t('Import CSV …')} wide disabled={ro} onClick={() => run('importCsv')} />
          </Group_>
          <Group_ label={t('Hostwerk data')}>
            <Tool icon={Radio} label={t('Live data …')} wide disabled={ro} onClick={() => run('liveData')} />
            <Tool icon={Database} label={t('Database table …')} wide disabled={ro} onClick={() => run('dbOpen')} />
            <Tool icon={RefreshCw} label={t('Refresh all')} kbd="Ctrl+Alt+F5" wide onClick={() => { run('pivotRefreshAll'); run('liveRefresh'); }} />
          </Group_>
          <Group_ label={t('Automate')}>
            <Tool icon={ClipboardList} label={P.hasForm ? t('Form …') : t('Create form …')} wide disabled={ro} onClick={() => run(P.hasForm ? 'formOpen' : 'form')} />
            <Tool icon={BellRing} label={t('Rules & notifications …')} wide onClick={() => run('rules')} />
            {P.aiReady && <Tool icon={Sparkles} label={t('Clean up data …')} wide disabled={ro} onClick={() => run('ai', 'clean')} />}
          </Group_>
          <Group_ label={t('Outline')}>
            <Tool icon={Group} label={t('Group')} disabled={ro} onClick={() => run('group', 1)} />
            <Tool icon={Ungroup} label={t('Ungroup')} disabled={ro} onClick={() => run('group', -1)} />
            <Tool label={t('Hide detail')} wide disabled={ro} onClick={() => run('detail', false)} />
            <Tool label={t('Show detail')} wide disabled={ro} onClick={() => run('detail', true)} />
          </Group_>
        </>}
        {tab === 'review' && <>
          <Group_ label={t('Notes')}>
            <Tool icon={StickyNote} label={flags.hasNote ? t('Edit note') : t('New note')} kbd="Shift+F2" wide disabled={ro} onClick={() => run('note')} />
            <Tool icon={MessageSquareX} label={t('Delete note')} wide disabled={ro || !flags.hasNote} onClick={() => run('clear', 'notes')} />
            <Tool icon={Eye} label={t('Show all notes')} active={flags.showNotes} wide onClick={() => run('showNotes')} />
          </Group_>
          <Group_ label={t('Find')}>
            <Tool icon={Search} label={t('Find & replace')} kbd="Ctrl+H" wide onClick={() => run('replace')} />
          </Group_>
          <Group_ label={t('Together')}>
            <Tool icon={Users} label={P.collabOn ? t('Stop co-editing') : t('Edit together')} active={P.collabOn} wide disabled={ro} onClick={() => run('collab')} />
            <Tool icon={GitCompare} label={t('Compare versions …')} wide onClick={(e) => { const r = e && e.currentTarget ? e.currentTarget.getBoundingClientRect() : null; run('compare', r ? { left: r.left, below: r.bottom + 2 } : null); }} />
          </Group_>
          {flags.hasMacros && <Group_ label={t('Macros')}><Tool icon={Lock} label={t('View macros')} wide onClick={() => run('macros')} /></Group_>}
          <Group_ label={t('Workbook')}>
            <Tool label={t('Workbook statistics')} wide onClick={() => run('stats')} />
          </Group_>
        </>}
        {tab === 'view' && <>
          <Group_ label={t('Window')}>
            <DropBtn icon={Snowflake} label={t('Freeze panes')} items={freezeMenu} />
          </Group_>
          <Group_ label={t('Show')}>
            <label className="check small"><input type="checkbox" checked={flags.gridlines} onChange={() => run('gridlines')} />{t('Gridlines')}</label>
            <label className="check small"><input type="checkbox" checked={flags.headings} onChange={() => run('headings')} />{t('Headings')}</label>
            <label className="check small"><input type="checkbox" checked={flags.formulaBar} onChange={() => run('formulaBar')} />{t('Formula bar')}</label>
          </Group_>
          <Group_ label={t('Zoom')}>
            <Tool icon={ZoomOut} label={t('Zoom out')} onClick={() => run('zoom', -10)} />
            <button type="button" className="btn ghost sm" onClick={() => run('zoomSet', 100)} title={t('Zoom to 100 %')}>100 %</button>
            <Tool icon={ZoomIn} label={t('Zoom in')} onClick={() => run('zoom', 10)} />
            <Tool label={t('Zoom to selection')} wide onClick={() => run('zoomSel')} />
          </Group_>
          <Group_ label={t('Formulas')}>
            <Tool icon={CodeXml} label={t('Show formulas')} active={flags.showFormulas} wide onClick={() => run('showFormulas')} />
            <Tool label={t('R1C1 reference style')} active={P.r1c1} wide onClick={() => run('r1c1')} />
            <Tool icon={PanelTop} label={t('Full screen grid')} active={flags.compact} wide onClick={() => run('compact')} />
          </Group_>
          <Group_ label={t('Views')}>
            <Tool icon={Table2} label={t('Grid')} active={P.view === 'grid'} wide onClick={() => run('view', 'grid')} />
            <Tool icon={Columns3} label={t('Kanban')} active={P.view === 'kanban'} wide onClick={() => run('view', 'kanban')} />
            <Tool icon={CalendarDays} label={t('Calendar')} active={P.view === 'calendar'} wide onClick={() => run('view', 'calendar')} />
            <Tool icon={LayoutGrid} label={t('Gallery')} active={P.view === 'gallery'} wide onClick={() => run('view', 'gallery')} />
          </Group_>
        </>}
        {tab === 'ctx-table' && P.table && <>
          <Group_ label={t('Properties')}>
            <Tool icon={TableProperties} label={P.table.name} title={t('Rename table …')} wide onClick={() => run('tableName')} />
            <Tool label={t('Convert to range')} wide disabled={ro} onClick={() => run('tableToRange')} />
            <Tool icon={TableCellsMerge} label={t('Summarize with pivot table')} wide disabled={ro} onClick={() => run('pivotNew')} />
          </Group_>
          <Group_ label={t('Table style options')}>
            <label className="check small"><input type="checkbox" checked={!!P.table.totals} disabled={ro} onChange={() => run('tableOpt', 'totals')} />{t('Total row')}</label>
            <label className="check small"><input type="checkbox" checked={P.table.stripes !== false} disabled={ro} onChange={() => run('tableOpt', 'stripes')} />{t('Banded rows')}</label>
            <label className="check small"><input type="checkbox" checked={!!P.table.colStripes} disabled={ro} onChange={() => run('tableOpt', 'colStripes')} />{t('Banded columns')}</label>
            <label className="check small"><input type="checkbox" checked={!!P.table.firstCol} disabled={ro} onChange={() => run('tableOpt', 'firstCol')} />{t('First column')}</label>
            <label className="check small"><input type="checkbox" checked={!!P.table.lastCol} disabled={ro} onChange={() => run('tableOpt', 'lastCol')} />{t('Last column')}</label>
          </Group_>
          <Group_ label={t('Table styles')}>
            <Tool icon={Palette} label={t('Table styles …')} wide disabled={ro} onClick={() => run('tableStylesMenu')} />
          </Group_>
        </>}
        {tab === 'ctx-pivot' && P.pivot && <>
          <Group_ label={t('PivotTable')}>
            <Tool icon={TableCellsMerge} label={t('Field list')} wide onClick={() => run('pivotEdit')} />
            <Tool icon={RefreshCw} label={t('Refresh')} kbd="Alt+F5" wide disabled={ro || P.pivot.readOnly} onClick={() => run('pivotRefresh')} />
            <Tool icon={RefreshCw} label={t('Refresh all')} kbd="Ctrl+Alt+F5" wide disabled={ro} onClick={() => run('pivotRefreshAll')} />
          </Group_>
        </>}
        {tab === 'ctx-obj' && P.objSel && <>
          <Group_ label={t('Arrange')}>
            <Tool icon={Trash} label={t('Delete')} kbd="Del" wide disabled={ro} onClick={() => run('objDelete')} />
          </Group_>
        </>}
      </div>}
    </div>
  );
}

export { pickColor, anchorOf, showMenu };
