/*
 * Ribbon of the PDF editor: View, Annotate, Pages, Fill & sign, Tools – in the PDF red of Office's colour scheme,
 * collapsible (the tabs stay), with contextual options: colour/width/size follow the active tool or the selected
 * annotation. Every button has a label or title with its shortcut. `st` = state, `act` = actions (see index.jsx).
 */
import {
  PanelLeft, ChevronLeft, ChevronRight, ZoomIn, ZoomOut, MoveHorizontal, Maximize, RotateCw, RotateCcw, Search, Moon, Presentation,
  Printer, Info, FileText, MousePointer2, Hand, Highlighter, Underline, Strikethrough, PenLine, Square, Circle, Minus, ArrowUpRight,
  Type, StickyNote, Stamp, ImagePlus, Signature, Undo2, Redo2, Trash2, FilePlus2, FileInput, Files, Scissors, FileOutput, ArrowUp, ArrowDown,
  TextSelect, ScanText, Layers, ListChecks, ChevronsUp, ChevronsDown, EyeOff, GitCompare, Gauge, Lock, LockOpen, KeyRound,
} from 'lucide-react';
import { Seg, Switch } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { COLORS, STAMPS } from './annotations.js';
import { withKeys } from '../_shared/keys.js';
import { colorName } from '../_shared/colors.js';

const ZOOMS = [0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4];

function B({ icon: I, label, title, on, onClick, disabled, wide, danger }) {
  return (
    <button type="button" className={`btn sm ghost${wide ? ' of-pdf-wide' : ' of-pdf-tb'}${on ? ' on' : ''}${danger ? ' of-pdf-dangerbtn' : ''}`} onClick={onClick} disabled={disabled} title={title || label} aria-label={title || label} aria-pressed={on === undefined ? undefined : !!on}>
      <I size={16} />{wide && <span>{label}</span>}
    </button>
  );
}
/** a big ribbon button: icon over the label (Tools tab) */
function Big({ icon: I, label, title, onClick, on, disabled }) {
  return (
    <button type="button" className={`of-pdf-big${on ? ' on' : ''}`} onClick={onClick} disabled={disabled} title={title || label} aria-label={title || label} aria-pressed={on === undefined ? undefined : !!on}>
      <I size={20} /><span>{label}</span>
    </button>
  );
}
const Sep = () => <span className="of-pdf-sep" aria-hidden="true" />;
const Group = ({ label, children }) => <div className="of-pdf-group" role="group" aria-label={label}><div className="of-pdf-group-items">{children}</div><span className="of-pdf-group-label">{label}</span></div>;

function Colors({ value, onChange }) {
  return (
    <span className="of-pdf-colors" role="radiogroup" aria-label="Colour">
      {COLORS.map((c) => <button key={c} type="button" role="radio" aria-checked={value === c} className={`of-pdf-swatch${value === c ? ' on' : ''}`} style={{ background: c }} title={t(colorName(c))} aria-label={t(colorName(c))} onClick={() => onChange(c)} />)}
      <input type="color" className="of-pdf-colorpick" value={value} onChange={(e) => onChange(e.target.value)} title="More colours" aria-label="More colours" />
    </span>
  );
}

const SHAPES = ['ink', 'rect', 'ellipse', 'line', 'arrow'];
const TOOL_LABEL = { highlight: 'Highlight', underline: 'Underline', strike: 'Strike through', ink: 'Pen', rect: 'Rectangle', ellipse: 'Ellipse', line: 'Line', arrow: 'Arrow', text: 'Text box', note: 'Sticky note', stamp: 'Stamp' };

export default function Toolbar({ st, act }) {
  const { tab, editable } = st;
  const tabs = [{ id: 'home', label: t('View') }];
  if (editable) tabs.push({ id: 'annotate', label: t('Annotate') }, { id: 'pages', label: t('Pages') }, { id: 'forms', label: t('Fill & sign') });
  if (!st.readOnly) tabs.push({ id: 'tools', label: t('Tools') });
  const shown = tabs.some((x) => x.id === tab) ? tab : 'home';
  const zoomLabel = typeof st.zoom === 'number' ? `${Math.round(st.zoom * 100)} %` : st.zoom === 'page-fit' ? t('Whole page') : st.zoom === 'page-width' ? t('Page width') : t('Automatic');
  const tool = (id, I, label, key) => <B key={id} icon={I} label={label} title={withKeys(label, key)} on={st.tool === id} onClick={() => act.setTool(id)} />;
  // contextual options: the active tool, or the selected annotation
  const ctx = st.selectedType && st.tool === 'select' ? st.selectedType : st.tool;
  const showColor = editable && (['highlight', 'underline', 'strike', 'text', ...SHAPES].includes(ctx));
  const showWidth = editable && SHAPES.includes(ctx);
  const showSize = editable && ctx === 'text';

  return (
    <div className={`of-pdf-ribbon${st.ribbonOpen ? '' : ' collapsed'}`}>
      <div className="of-pdf-ribbon-tabs" role="tablist">
        {tabs.map((x) => <button key={x.id} type="button" role="tab" aria-selected={shown === x.id} className={`of-pdf-rtab${shown === x.id ? ' active' : ''}`} onClick={() => { if (!st.ribbonOpen) act.toggleRibbon(); act.setTab(x.id); }} onDoubleClick={act.toggleRibbon}>{x.label}</button>)}
        <span className="grow" />
        {editable && <><B icon={Undo2} label={t('Undo')} title={withKeys(t('Undo'), 'Ctrl+Z')} onClick={act.undo} disabled={!st.canUndo} /><B icon={Redo2} label={t('Redo')} title={withKeys(t('Redo'), 'Ctrl+Y')} onClick={act.redo} disabled={!st.canRedo} /></>}
        <B icon={Search} label={t('Search')} title={withKeys(t('Search'), 'Ctrl+F')} on={st.searchOpen} onClick={act.toggleSearch} />
        <B icon={st.ribbonOpen ? ChevronsUp : ChevronsDown} label={st.ribbonOpen ? t('Collapse the ribbon') : t('Show the ribbon')} title={st.ribbonOpen ? t('Collapse the ribbon (double-click a tab)') : t('Show the ribbon')} onClick={act.toggleRibbon} />
      </div>
      {st.ribbonOpen && (
        <div className="of-pdf-ribbon-body" key={shown}>
          {shown === 'home' && (
            <>
              <B icon={PanelLeft} label={t('Sidebar')} on={st.sidebar} onClick={act.toggleSidebar} />
              <Sep />
              <B icon={ChevronLeft} label={t('Previous page')} title={`${t('Previous page')} (Page Up)`} onClick={act.prev} disabled={st.pageNo <= 1} />
              <span className="of-pdf-pageno">
                <input className="input" value={st.pageInput} onChange={(e) => act.setPageInput(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') act.goInput(); }} onBlur={act.goInput} aria-label="Page number" inputMode="numeric" />
                <span className="small muted">/ {st.pageCount}</span>
              </span>
              <B icon={ChevronRight} label={t('Next page')} title={`${t('Next page')} (Page Down)`} onClick={act.next} disabled={st.pageNo >= st.pageCount} />
              <Sep />
              <B icon={ZoomOut} label={t('Zoom out')} title={withKeys(t('Zoom out'), 'Ctrl+-')} onClick={() => act.zoomStep(-1)} />
              <select className="input of-pdf-zoomsel" value={typeof st.zoom === 'number' ? String(st.zoom) : st.zoom} onChange={(e) => act.setZoom(/^[\d.]+$/.test(e.target.value) ? Number(e.target.value) : e.target.value)} aria-label="Zoom" title={zoomLabel}>
                <option value="auto">{t('Automatic')}</option>
                <option value="page-width">{t('Page width')}</option>
                <option value="page-fit">{t('Whole page')}</option>
                {ZOOMS.map((z) => <option key={z} value={String(z)}>{Math.round(z * 100)} %</option>)}
                {typeof st.zoom === 'number' && !ZOOMS.includes(st.zoom) && <option value={String(st.zoom)}>{Math.round(st.zoom * 100)} %</option>}
              </select>
              <B icon={ZoomIn} label={t('Zoom in')} title={withKeys(t('Zoom in'), 'Ctrl++')} onClick={() => act.zoomStep(1)} />
              <B icon={MoveHorizontal} label={t('Page width')} on={st.zoom === 'page-width'} onClick={() => act.setZoom('page-width')} />
              <B icon={Maximize} label={t('Whole page')} title={withKeys(t('Whole page'), 'Ctrl+0')} on={st.zoom === 'page-fit'} onClick={() => act.setZoom('page-fit')} />
              <B icon={RotateCw} label={t('Rotate view')} title={`${t('Rotate view')} (R)`} onClick={act.rotateView} />
              <Sep />
              <Seg value={st.scroll} onChange={act.setScroll} options={[{ id: 'vertical', label: t('Scroll'), title: t('Continuous vertical scroll') }, { id: 'horizontal', label: t('Side by side'), title: t('Horizontal scroll') }, { id: 'page', label: t('Single page'), title: t('One page at a time') }]} />
              <Sep />
              <B icon={st.tool === 'hand' ? Hand : TextSelect} label={st.tool === 'hand' ? t('Hand tool') : t('Select text')} title={t('Select text / hand tool (H) – the middle mouse button always pans')} on={st.tool === 'hand'} onClick={() => act.setTool(st.tool === 'hand' ? 'text-select' : 'hand')} />
              <B icon={Moon} label={t('Night reading')} title={t('Night reading – shows the pages dark (the file stays as it is)')} on={st.night} onClick={act.toggleNight} />
              <B icon={Presentation} label={t('Presentation')} title={`${t('Presentation')} (F5)`} onClick={act.present} />
              <span className="grow" />
              <B icon={Printer} label={t('Print …')} title={withKeys(t('Print …'), 'Ctrl+P')} onClick={act.print} wide />
              <B icon={Info} label={t('Properties')} title={t('Document properties')} onClick={act.properties} />
            </>
          )}
          {shown === 'annotate' && (
            <>
              {tool('text-select', TextSelect, t('Select text'))}
              {tool('select', MousePointer2, t('Select and move annotations'), 'V')}
              {tool('hand', Hand, t('Hand tool'), 'H')}
              <Sep />
              {tool('highlight', Highlighter, t('Highlight text'))}
              {tool('underline', Underline, t('Underline text'))}
              {tool('strike', Strikethrough, t('Strike through text'))}
              <Sep />
              {tool('ink', PenLine, t('Pen'), 'P')}
              {tool('rect', Square, t('Rectangle'))}
              {tool('ellipse', Circle, t('Ellipse'))}
              {tool('line', Minus, t('Line'))}
              {tool('arrow', ArrowUpRight, t('Arrow'))}
              <Sep />
              {tool('text', Type, t('Text box'), 'T')}
              {tool('note', StickyNote, t('Sticky note'))}
              <span className="of-pdf-stampsel">
                <B icon={Stamp} label={t('Stamp')} on={st.tool === 'stamp'} onClick={() => act.setTool('stamp')} />
                <select className="input" value={st.stamp} onChange={(e) => { act.setStamp(e.target.value); act.setTool('stamp'); }} aria-label="Stamp" title="Stamp">
                  {STAMPS.map((s) => <option key={s.text} value={s.text}>{t(s.text)}</option>)}
                </select>
              </span>
              <B icon={ImagePlus} label={t('Image …')} onClick={act.pickImage} />
              {(showColor || showWidth || showSize || st.selectedAnnot) && (
                <span className="of-pdf-context" aria-label="Options">
                  <span className="of-pdf-context-label small">{st.selectedAnnot && st.tool === 'select' ? t('Selected: {what}', { what: t(TOOL_LABEL[st.selectedType] || 'Annotation') }) : t(TOOL_LABEL[ctx] || '')}</span>
                  {showColor && <Colors value={st.color} onChange={act.setColor} />}
                  {showWidth && (
                    <select className="input of-pdf-small" value={st.width} onChange={(e) => act.setWidth(Number(e.target.value))} aria-label="Line width" title="Line width">
                      {[1, 2, 3, 4, 6, 8, 12].map((w) => <option key={w} value={w}>{w} pt</option>)}
                    </select>
                  )}
                  {showSize && (
                    <select className="input of-pdf-small" value={st.fontSize} onChange={(e) => act.setFontSize(Number(e.target.value))} aria-label="Font size" title="Font size">
                      {[8, 9, 10, 11, 12, 14, 16, 18, 24, 32, 48].map((w) => <option key={w} value={w}>{w} pt</option>)}
                    </select>
                  )}
                  {st.selectedAnnot && <B icon={Trash2} label={t('Delete annotation')} title={withKeys(t('Delete annotation'), 'Del')} onClick={act.deleteAnnot} />}
                </span>
              )}
              <span className="grow" />
              <span className="of-pdf-mode" title={t('Editable: stay editable here and in other PDF apps. Burned in: become part of the page.')}>
                <Layers size={14} />
                <Seg value={st.annotMode} onChange={act.setAnnotMode} options={[{ id: 'annot', label: t('Editable') }, { id: 'burn', label: t('Burned in') }]} />
              </span>
            </>
          )}
          {shown === 'pages' && (
            <>
              <B icon={FilePlus2} label={t('Blank page')} title={t('Insert a blank page after the current one')} onClick={act.insertBlank} wide />
              <B icon={FileInput} label={t('Insert file …')} title={t('Insert pages from PDFs or images after the current page – with their form fields and bookmarks')} onClick={act.insertFile} wide />
              <B icon={Files} label={t('Merge PDFs …')} title={t('Append several PDFs at the end')} onClick={act.mergeFiles} wide />
              <Sep />
              <B icon={RotateCcw} label={t('Rotate left')} onClick={() => act.rotatePages(-90)} />
              <B icon={RotateCw} label={t('Rotate right')} onClick={() => act.rotatePages(90)} />
              <B icon={ArrowUp} label={t('Move up')} onClick={() => act.movePages(-1)} />
              <B icon={ArrowDown} label={t('Move down')} onClick={() => act.movePages(1)} />
              <B icon={Trash2} label={t('Delete pages')} onClick={act.deletePages} disabled={st.pageCount <= 1} />
              <Sep />
              <B icon={FileOutput} label={t('Extract …')} title={t('Save the selected pages as a new PDF')} onClick={act.extract} wide />
              <B icon={Scissors} label={t('Split …')} title={t('Split into several PDFs')} onClick={act.split} wide />
              <span className="grow" />
              <span className="small muted">{st.selCount > 1 ? t('{n} pages selected', { n: st.selCount }) : t('Ctrl/Shift-click thumbnails to select several pages – drag them to reorder')}</span>
            </>
          )}
          {shown === 'forms' && (
            <>
              <B icon={Signature} label={t('Signature …')} onClick={act.signature} wide />
              <B icon={Type} label={t('Text box')} on={st.tool === 'text'} onClick={() => act.setTool('text')} wide />
              <B icon={Stamp} label={t('Stamp')} on={st.tool === 'stamp'} onClick={() => act.setTool('stamp')} wide />
              <Sep />
              {st.formInfo && st.formInfo.fields > 0 ? (
                <>
                  <span className="small"><ListChecks size={14} /> {t('{n} form fields – click into the page, Tab moves to the next field', { n: st.formInfo.fields })}</span>
                  <Sep />
                  <Switch checked={st.flatten} onChange={act.setFlatten} label={t('Flatten form when saving')} />
                </>
              ) : <span className="small muted"><ScanText size={14} /> {t('This PDF has no form fields – use text boxes to fill it in.')}</span>}
            </>
          )}
          {shown === 'tools' && (
            <>
              <Group label={t('Recognize')}>
                <Big icon={ScanText} label={t('OCR')} title={t('Recognize the text of scanned pages – searchable and copyable everywhere')} onClick={act.ocr} disabled={!editable} />
              </Group>
              <Group label={t('Protect')}>
                <Big icon={EyeOff} label={t('Redact')} title={t('Remove personal data for good – areas, found e-mail addresses, IBANs, phone numbers …')} on={st.redactOpen} onClick={act.redact} disabled={!editable} />
                {st.locked ? <Big icon={KeyRound} label={t('Unlock')} title={t('Enter the permissions password')} onClick={act.unlock} />
                  : <Big icon={st.protectedNow ? Lock : LockOpen} label={st.protectedNow ? t('Protected') : t('Password')} title={t('Protect with a password (AES-256) or remove the protection')} on={st.protectedNow} onClick={act.protect} disabled={!editable} />}
              </Group>
              <Group label={t('Check & optimize')}>
                <Big icon={GitCompare} label={t('Compare')} title={t('Compare with another PDF – text and visual differences')} onClick={act.compare} />
                <Big icon={Gauge} label={t('Compress')} title={t('Make the file smaller')} onClick={act.compress} disabled={!editable} />
              </Group>
              {st.hasConverter && (
                <Group label={t('Convert')}>
                  <Big icon={FileText} label={t('To Word')} title={t('Convert to Word …')} onClick={act.toWord} />
                </Group>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
}
