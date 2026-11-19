/*
 * Ribbon of the PDF editor: Home (view), Annotate, Pages, Fill & sign. Every button has a label or title with the
 * shortcut. `st` = state, `act` = actions (see index.jsx).
 */
import {
  PanelLeft, ChevronLeft, ChevronRight, ZoomIn, ZoomOut, MoveHorizontal, Maximize, RotateCw, RotateCcw, Search, Moon, Presentation,
  Printer, Info, FileText, MousePointer2, Hand, Highlighter, Underline, Strikethrough, PenLine, Square, Circle, Minus, ArrowUpRight,
  Type, StickyNote, Stamp, ImagePlus, Signature, Undo2, Redo2, Trash2, FilePlus2, FileInput, Files, Scissors, FileOutput, ArrowUp, ArrowDown,
  TextSelect, ScanText, Layers, ListChecks,
} from 'lucide-react';
import { Seg, Switch } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { COLORS, STAMPS } from './annotations.js';
import { withKeys } from '../_shared/keys.js';
import { colorName } from '../_shared/colors.js';

const ZOOMS = [0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4];

function B({ icon: I, label, title, on, onClick, disabled, wide }) {
  return (
    <button type="button" className={`btn sm ghost${wide ? '' : ' of-pdf-tb'}${on ? ' on' : ''}`} onClick={onClick} disabled={disabled} title={title || label} aria-label={title || label} aria-pressed={on === undefined ? undefined : !!on}>
      <I size={16} />{wide && <span>{label}</span>}
    </button>
  );
}
const Sep = () => <span className="of-pdf-sep" aria-hidden="true" />;

function Colors({ value, onChange }) {
  return (
    <span className="of-pdf-colors" role="radiogroup" aria-label="Colour">
      {COLORS.map((c) => <button key={c} type="button" role="radio" aria-checked={value === c} className={`of-pdf-swatch${value === c ? ' on' : ''}`} style={{ background: c }} title={t(colorName(c))} aria-label={t(colorName(c))} onClick={() => onChange(c)} />)}
      <input type="color" className="of-pdf-colorpick" value={value} onChange={(e) => onChange(e.target.value)} title="More colours" aria-label="More colours" />
    </span>
  );
}

export default function Toolbar({ st, act }) {
  const { tab, editable } = st;
  const tabs = [{ id: 'home', label: t('View') }];
  if (editable) tabs.push({ id: 'annotate', label: t('Annotate') }, { id: 'pages', label: t('Pages') }, { id: 'forms', label: t('Fill & sign') });
  const zoomLabel = typeof st.zoom === 'number' ? `${Math.round(st.zoom * 100)} %` : st.zoom === 'page-fit' ? t('Whole page') : st.zoom === 'page-width' ? t('Page width') : t('Automatic');
  const tool = (id, I, label, key) => <B key={id} icon={I} label={label} title={withKeys(label, key)} on={st.tool === id} onClick={() => act.setTool(id)} />;

  return (
    <div className="of-pdf-ribbon">
      <div className="of-pdf-ribbon-tabs" role="tablist">
        {tabs.map((x) => <button key={x.id} type="button" role="tab" aria-selected={tab === x.id} className={`of-pdf-rtab${tab === x.id ? ' active' : ''}`} onClick={() => act.setTab(x.id)}>{x.label}</button>)}
        <span className="grow" />
        {editable && <><B icon={Undo2} label={t('Undo')} title={withKeys(t('Undo'), 'Ctrl+Z')} onClick={act.undo} disabled={!st.canUndo} /><B icon={Redo2} label={t('Redo')} title={withKeys(t('Redo'), 'Ctrl+Y')} onClick={act.redo} disabled={!st.canRedo} /></>}
        <B icon={Search} label={t('Search')} title={withKeys(t('Search'), 'Ctrl+F')} on={st.searchOpen} onClick={act.toggleSearch} />
      </div>
      <div className="of-pdf-ribbon-body">
        {tab === 'home' && (
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
            <B icon={st.tool === 'hand' ? Hand : TextSelect} label={st.tool === 'hand' ? t('Hand tool') : t('Select text')} title={t('Select text / hand tool (H)')} on={st.tool === 'hand'} onClick={() => act.setTool(st.tool === 'hand' ? 'text-select' : 'hand')} />
            <B icon={Moon} label={t('Night reading')} title={t('Night reading – shows the pages dark (the file stays as it is)')} on={st.night} onClick={act.toggleNight} />
            <B icon={Presentation} label={t('Presentation')} title={`${t('Presentation')} (F5)`} onClick={act.present} />
            <span className="grow" />
            <B icon={Printer} label={t('Print …')} title={withKeys(t('Print …'), 'Ctrl+P')} onClick={act.print} wide />
            <B icon={Info} label={t('Properties')} title={t('Document properties')} onClick={act.properties} />
            {st.hasConverter && <B icon={FileText} label={t('Convert to Word …')} onClick={act.toWord} wide />}
          </>
        )}
        {tab === 'annotate' && (
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
            <Sep />
            <Colors value={st.color} onChange={act.setColor} />
            {['ink', 'rect', 'ellipse', 'line', 'arrow'].includes(st.tool) && (
              <select className="input of-pdf-small" value={st.width} onChange={(e) => act.setWidth(Number(e.target.value))} aria-label="Line width" title="Line width">
                {[1, 2, 3, 4, 6, 8, 12].map((w) => <option key={w} value={w}>{w} pt</option>)}
              </select>
            )}
            {st.tool === 'text' && (
              <select className="input of-pdf-small" value={st.fontSize} onChange={(e) => act.setFontSize(Number(e.target.value))} aria-label="Font size" title="Font size">
                {[8, 9, 10, 11, 12, 14, 16, 18, 24, 32, 48].map((w) => <option key={w} value={w}>{w} pt</option>)}
              </select>
            )}
            <B icon={Trash2} label={t('Delete annotation')} title={withKeys(t('Delete annotation'), 'Del')} onClick={act.deleteAnnot} disabled={!st.selectedAnnot} />
            <span className="grow" />
            <span className="of-pdf-mode" title={t('Editable: other PDF apps can still move or delete them. Burned in: become part of the page.')}>
              <Layers size={14} />
              <Seg value={st.annotMode} onChange={act.setAnnotMode} options={[{ id: 'annot', label: t('Editable') }, { id: 'burn', label: t('Burned in') }]} />
            </span>
          </>
        )}
        {tab === 'pages' && (
          <>
            <B icon={FilePlus2} label={t('Blank page')} title={t('Insert a blank page after the current one')} onClick={act.insertBlank} wide />
            <B icon={FileInput} label={t('Insert file …')} title={t('Insert pages from PDFs or images after the current page')} onClick={act.insertFile} wide />
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
            <span className="small muted">{st.selCount > 1 ? t('{n} pages selected', { n: st.selCount }) : t('Ctrl/Shift-click thumbnails to select several pages')}</span>
          </>
        )}
        {tab === 'forms' && (
          <>
            <B icon={Signature} label={t('Signature …')} onClick={act.signature} wide />
            <B icon={Type} label={t('Text box')} on={st.tool === 'text'} onClick={() => act.setTool('text')} wide />
            <B icon={Stamp} label={t('Stamp')} on={st.tool === 'stamp'} onClick={() => act.setTool('stamp')} wide />
            <Sep />
            {st.formInfo && st.formInfo.fields > 0 ? (
              <>
                <span className="small"><ListChecks size={14} /> {t('{n} form fields – click into the page to fill them', { n: st.formInfo.fields })}</span>
                <Sep />
                <Switch checked={st.flatten} onChange={act.setFlatten} label={t('Flatten form when saving')} />
              </>
            ) : <span className="small muted"><ScanText size={14} /> {t('This PDF has no form fields – use text boxes to fill it in.')}</span>}
          </>
        )}
      </div>
    </div>
  );
}
