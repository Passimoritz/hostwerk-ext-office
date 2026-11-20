/*
 * The ribbon: Home · Insert · Layout · Review · View, plus contextual Table and Picture tabs (like Word).
 * Every button keeps the editor selection and names its shortcut in the tooltip. Hostwerk's own shortcuts stay free:
 * link = Ctrl+Shift+K (Ctrl+K is Hostwerk's search), justify = Ctrl+Shift+J (Ctrl+J is Hostwerk's AI panel).
 */
import {
  Undo2, Redo2, Scissors, Copy, ClipboardPaste, ClipboardType, Paintbrush, Bold, Italic, Underline, Strikethrough, Subscript, Superscript,
  Baseline, Highlighter, RemoveFormatting, CaseSensitive, List, ListOrdered, ListTodo, ListTree, IndentIncrease, IndentDecrease,
  AlignLeft, AlignCenter, AlignRight, AlignJustify, Pilcrow, Table, Image, Link, Minus, SeparatorHorizontal, Omega, Smile, CalendarClock,
  TableOfContents, PanelTop, MessageSquarePlus, Sigma, SquareDashed, FileText, RectangleHorizontal, RectangleVertical, Columns2, PaintBucket,
  SpellCheck, Search, Replace, Check, X, ChevronLeft, ChevronRight, Ruler, PanelLeft, PanelRight, ZoomIn, ZoomOut, Focus, Printer,
  TableCellsMerge, TableCellsSplit, Trash2, BetweenHorizontalStart, BetweenHorizontalEnd, BetweenVerticalStart, BetweenVerticalEnd,
  AArrowUp, AArrowDown, FilePen, FileDigit, Languages, Code, Unlink, ImageUp, Type, Maximize, SquareCode,
  Workflow, Braces, RefreshCw, CheckSquare, Sparkles, GitCompare, Users, LayoutGrid, ChevronUp, Save, SquareSplitVertical, ChevronsUpDown, Wand2,
} from 'lucide-react';
import { t } from '@hostwerk/i18n';
import { RB, Drop, MenuItem, ColorGrid, Group, HIGHLIGHTS } from './ui.jsx';
import { TableGrid, EmojiGrid, dateFormats } from './dialogs.jsx';
import { withKeys } from '../_shared/keys.js';

export const FONTS = ['Calibri', 'Aptos', 'Arial', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Helvetica', 'Impact', 'Liberation Sans', 'Liberation Serif', 'Lucida Sans', 'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'system-ui'];
export const FONT_STACK = { Calibri: 'Calibri, Carlito, "Segoe UI", sans-serif', Cambria: 'Cambria, Caladea, Georgia, serif', Arial: 'Arial, "Liberation Sans", Helvetica, sans-serif', 'Times New Roman': '"Times New Roman", "Liberation Serif", Times, serif', 'Courier New': '"Courier New", "Liberation Mono", monospace', Aptos: 'Aptos, Calibri, Carlito, sans-serif', Helvetica: 'Helvetica, Arial, sans-serif', Georgia: 'Georgia, serif', Consolas: 'Consolas, "Cascadia Mono", monospace', 'system-ui': 'system-ui, sans-serif' };
export const fontCss = (f) => FONT_STACK[f] || `"${f}", sans-serif`;
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const STYLES = [
  { id: 'normal', label: 'Normal' }, { id: 'title', label: 'Title' }, { id: 'subtitle', label: 'Subtitle' },
  { id: 'h1', label: 'Heading 1', keys: 'Ctrl+Alt+1' }, { id: 'h2', label: 'Heading 2', keys: 'Ctrl+Alt+2' }, { id: 'h3', label: 'Heading 3', keys: 'Ctrl+Alt+3' },
  { id: 'h4', label: 'Heading 4' }, { id: 'h5', label: 'Heading 5' }, { id: 'h6', label: 'Heading 6' },
  { id: 'quote', label: 'Quote' }, { id: 'code', label: 'Code' }, { id: 'listparagraph', label: 'List paragraph' },
];
const MARGINS = [
  { id: 'normal', label: 'Normal', m: { top: 25, right: 25, bottom: 25, left: 25 } },
  { id: 'din', label: 'DIN 5008', m: { top: 20, right: 20, bottom: 20, left: 25 } },
  { id: 'narrow', label: 'Narrow', m: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 } },
  { id: 'moderate', label: 'Moderate', m: { top: 25.4, right: 19.1, bottom: 25.4, left: 19.1 } },
  { id: 'wide', label: 'Wide', m: { top: 25.4, right: 50.8, bottom: 25.4, left: 50.8 } },
];
const LANGS = [['de-DE', 'German'], ['de-CH', 'German (Switzerland)'], ['de-AT', 'German (Austria)'], ['en-US', 'English (US)'], ['en-GB', 'English (UK)'], ['fr-FR', 'French'], ['es-ES', 'Spanish'], ['it-IT', 'Italian'], ['nl-NL', 'Dutch'], ['pl-PL', 'Polish'], ['pt-PT', 'Portuguese'], ['tr-TR', 'Turkish']];

export function currentStyle(editor) {
  if (editor.isActive('codeBlock')) return 'code';
  if (editor.isActive('blockquote')) return 'quote';
  for (let l = 1; l <= 6; l++) if (editor.isActive('heading', { level: l })) return `h${l}`;
  const ps = editor.getAttributes('paragraph').pstyle;
  return ps || 'normal';
}

export function Ribbon({ editor, readOnly, tab, setTab, s, act, host = {} }) {
  const e = editor;
  const c = () => e.chain().focus();
  const inTable = e.isActive('table');
  const img = e.isActive('image') ? e.getAttributes('image') : null;
  const ts = e.getAttributes('textStyle');
  const fontNow = (ts.fontFamily || s.defaultFont || '').split(',')[0].replace(/"/g, '').trim();
  const sizeNow = ts.fontSize ? parseFloat(ts.fontSize) : s.defaultSize;
  const tabs = [
    { id: 'home', label: 'Home' },
    ...(readOnly ? [] : [{ id: 'insert', label: 'Insert' }, { id: 'layout', label: 'Layout' }]),
    { id: 'review', label: 'Review' },
    { id: 'view', label: 'View' },
    ...(!readOnly && inTable ? [{ id: 'table', label: 'Table', ctx: true }] : []),
    ...(!readOnly && img ? [{ id: 'picture', label: 'Picture', ctx: true }] : []),
  ];
  const cur = tabs.some((x) => x.id === tab) ? tab : 'home';
  const style = currentStyle(e);
  const align = (a) => e.isActive({ textAlign: a });

  const home = (
    <>
      <Group label="Clipboard">
        {!readOnly && <RB icon={Scissors} title="Cut" keys="Ctrl+X" onClick={() => act.clipboard('cut')} />}
        <RB icon={Copy} title="Copy" keys="Ctrl+C" onClick={() => act.clipboard('copy')} />
        {!readOnly && <RB icon={ClipboardPaste} title="Paste" keys="Ctrl+V" onClick={() => act.paste(false)} />}
        {!readOnly && <RB icon={ClipboardType} title="Paste as plain text" keys="Ctrl+Shift+V" onClick={() => act.paste(true)} />}
        {!readOnly && <RB icon={Paintbrush} title="Format painter – copies the formatting to the next text you select" active={!!s.painter} onClick={act.painter} />}
      </Group>
      {!readOnly && (
        <Group label="Font">
          <div className="of-docs-row">
            <select className="input sm of-docs-font" value={FONTS.includes(fontNow) ? fontNow : fontNow || ''} onChange={(ev) => (ev.target.value ? c().setFontFamily(fontCss(ev.target.value)).run() : c().unsetFontFamily().run())} aria-label={t('Font')} title={t('Font')} style={{ fontFamily: fontCss(fontNow) }}>
              {!FONTS.includes(fontNow) && fontNow && <option value={fontNow}>{fontNow}</option>}
              {FONTS.map((f) => <option key={f} value={f} style={{ fontFamily: fontCss(f) }}>{f === 'system-ui' ? t('System font') : f}</option>)}
            </select>
            <input className="input sm of-docs-size" list="of-docs-sizes" value={sizeNow || ''} aria-label={t('Font size')} title={t('Font size')}
              onChange={(ev) => { const v = parseFloat(String(ev.target.value).replace(',', '.')); if (v > 0 && v < 1000) c().setFontSize(`${v}pt`).run(); }} />
            <datalist id="of-docs-sizes">{SIZES.map((x) => <option key={x} value={x} />)}</datalist>
            <RB icon={AArrowUp} title="Grow font" keys="Ctrl+]" onClick={() => act.fontStep(1)} />
            <RB icon={AArrowDown} title="Shrink font" keys="Ctrl+[" onClick={() => act.fontStep(-1)} />
          </div>
          <div className="of-docs-row">
            <RB icon={Bold} title="Bold" keys="Ctrl+B" active={e.isActive('bold')} onClick={() => c().toggleBold().run()} />
            <RB icon={Italic} title="Italic" keys="Ctrl+I" active={e.isActive('italic')} onClick={() => c().toggleItalic().run()} />
            <RB icon={Underline} title="Underline" keys="Ctrl+U" active={e.isActive('underline')} onClick={() => c().toggleUnderline().run()} />
            <RB icon={Strikethrough} title="Strikethrough" keys="Alt+Shift+5" active={e.isActive('strike')} onClick={() => c().toggleStrike().run()} />
            <RB icon={Subscript} title="Subscript" keys="Ctrl+=" active={e.isActive('subscript')} onClick={() => c().toggleSubscript().run()} />
            <RB icon={Superscript} title="Superscript" keys="Ctrl+Shift+=" active={e.isActive('superscript')} onClick={() => c().toggleSuperscript().run()} />
            <Drop icon={Baseline} title="Text colour" panelClass="of-pop-colors">{(close) => <ColorGrid close={close} onPick={(col) => (col ? c().setColor(col).run() : c().unsetColor().run())} />}</Drop>
            <Drop icon={Highlighter} title="Highlight" panelClass="of-pop-colors">{(close) => <ColorGrid close={close} colors={HIGHLIGHTS} noneLabel="No highlight" onPick={(col) => (col ? c().setHighlight({ color: col }).run() : c().unsetHighlight().run())} />}</Drop>
            <RB icon={RemoveFormatting} title="Clear formatting" keys="Ctrl+Space" onClick={() => c().unsetAllMarks().setParagraphFormat({ lineHeight: null, spaceBefore: null, spaceAfter: null, indent: null, firstLine: null }).run()} />
            <Drop icon={CaseSensitive} title="Change case">{(close) => (
              <>
                {[['sentence', 'Sentence case.'], ['lower', 'lowercase'], ['upper', 'UPPERCASE'], ['title', 'Capitalize Each Word'], ['toggle', 'tOGGLE cASE']].map(([id, label]) => <MenuItem key={id} onClick={() => { act.changeCase(id); close(); }}>{label}</MenuItem>)}
              </>
            )}</Drop>
          </div>
        </Group>
      )}
      {!readOnly && (
        <Group label="Paragraph">
          <div className="of-docs-row">
            <RB icon={List} title="Bullets" keys="Ctrl+Shift+L" active={e.isActive('bulletList')} onClick={() => c().toggleBulletList().run()} />
            <Drop icon={ListOrdered} title="Numbering" active={e.isActive('orderedList')} split={<RB icon={ListOrdered} title="Numbering" keys="Ctrl+Shift+7" active={e.isActive('orderedList')} onClick={() => c().toggleOrderedList().run()} />}>{(close) => (
              <>
                {[['decimal', '1. 2. 3.'], ['lower-alpha', 'a. b. c.'], ['upper-alpha', 'A. B. C.'], ['lower-roman', 'i. ii. iii.'], ['upper-roman', 'I. II. III.']].map(([id, label]) => <MenuItem key={id} onClick={() => { act.listStyle(id); close(); }}>{label}</MenuItem>)}
              </>
            )}</Drop>
            <RB icon={ListTree} title="Multilevel list (1. 1.1. 1.1.1.) – Tab and Shift+Tab change the level" active={e.isActive('orderedList', { listStyle: 'outline' })} onClick={() => act.listStyle('outline')} />
            <RB icon={ListTodo} title="Checklist" active={e.isActive('taskList')} onClick={() => c().toggleTaskList().run()} />
            <RB icon={IndentDecrease} title="Decrease indent" keys="Shift+Tab" onClick={() => c().changeIndent(-36).run()} />
            <RB icon={IndentIncrease} title="Increase indent" keys="Tab" onClick={() => c().changeIndent(36).run()} />
          </div>
          <div className="of-docs-row">
            <RB icon={AlignLeft} title="Align left" keys="Ctrl+L" active={align('left')} onClick={() => c().setTextAlign('left').run()} />
            <RB icon={AlignCenter} title="Centre" keys="Ctrl+E" active={align('center')} onClick={() => c().setTextAlign('center').run()} />
            <RB icon={AlignRight} title="Align right" keys="Ctrl+R" active={align('right')} onClick={() => c().setTextAlign('right').run()} />
            <RB icon={AlignJustify} title="Justify – Ctrl+J opens Hostwerk's AI panel" keys="Ctrl+Shift+J" active={align('justify')} onClick={() => c().setTextAlign('justify').run()} />
            <Drop icon={SeparatorHorizontal} title="Line and paragraph spacing">{(close) => (
              <>
                {[1, 1.15, 1.5, 2, 2.5, 3].map((v) => <MenuItem key={v} active={(e.getAttributes('paragraph').lineHeight || e.getAttributes('heading').lineHeight) === v} onClick={() => { c().setParagraphFormat({ lineHeight: v }).run(); close(); }}>{String(v).replace('.', document.documentElement.lang === 'de' ? ',' : '.')}</MenuItem>)}
                <hr />
                <MenuItem onClick={() => { c().setParagraphFormat({ spaceBefore: 12 }).run(); close(); }}>Add space before paragraph</MenuItem>
                <MenuItem onClick={() => { c().setParagraphFormat({ spaceBefore: 0 }).run(); close(); }}>Remove space before paragraph</MenuItem>
                <MenuItem onClick={() => { c().setParagraphFormat({ spaceAfter: 8 }).run(); close(); }}>Add space after paragraph</MenuItem>
                <MenuItem onClick={() => { c().setParagraphFormat({ spaceAfter: 0 }).run(); close(); }}>Remove space after paragraph</MenuItem>
                <hr />
                <MenuItem onClick={() => { close(); act.paragraphDialog(); }}>Line spacing options …</MenuItem>
              </>
            )}</Drop>
            <RB icon={Pilcrow} title="Show formatting marks" keys="Ctrl+Shift+8" active={s.marks} onClick={act.toggleMarks} />
          </div>
        </Group>
      )}
      {!readOnly && (
        <Group label="Styles" className="of-docs-styles-group">
          <div className="of-docs-styles" role="listbox" aria-label={t('Styles')}>
            {STYLES.map((st) => (
              <button type="button" role="option" aria-selected={style === st.id} key={st.id} className={`of-docs-style of-st-${st.id} ${style === st.id ? 'of-on' : ''}`} title={withKeys(t(st.label), st.keys)} onMouseDown={(ev) => ev.preventDefault()} onClick={() => e.chain().focus().setParagraphStyle(st.id).run()}>{st.label}</button>
            ))}
          </div>
        </Group>
      )}
      <Group label="Editing">
        <RB icon={Search} title="Find" keys="Ctrl+F" wide label="Find" onClick={() => act.find('find')} />
        {!readOnly && <RB icon={Replace} title="Replace" keys="Ctrl+H" wide label="Replace" onClick={() => act.find('replace')} />}
        <RB icon={Printer} title="Print" keys="Ctrl+P" onClick={act.print} />
      </Group>
      {!readOnly && s.ai && (
        <Group label="AI">
          <RB icon={Sparkles} title="Write with AI – or type “/ai” in the text" wide label="Write" onClick={() => act.ai('write')} />
          <RB icon={Wand2} title="Rewrite the selection with AI" wide label="Rewrite" disabled={e.state.selection.empty} onClick={() => act.ai('rewrite')} />
        </Group>
      )}
    </>
  );

  const insert = (
    <>
      <Group label="Pages">
        <RB icon={SeparatorHorizontal} title="Page break" keys="Ctrl+Enter" wide label="Page break" onClick={() => c().setPageBreak().run()} />
        <Drop icon={SquareSplitVertical} title="Section break – pages with their own size, orientation and margins" wide label="Section">{(close) => (
          <>
            <MenuItem icon={SquareSplitVertical} onClick={() => { close(); act.sectionBreak('nextPage'); }}>Section break, next page</MenuItem>
            <MenuItem icon={ChevronsUpDown} onClick={() => { close(); act.sectionBreak('continuous'); }}>Section break, continuous</MenuItem>
          </>
        )}</Drop>
      </Group>
      <Group label="Tables">
        <Drop icon={Table} title="Table" wide label="Table" panelClass="of-pop-table">{(close) => <TableGrid onPick={(r, col) => { c().insertTable({ rows: r, cols: col, withHeaderRow: true }).run(); close(); }} onCustom={() => { close(); act.customTable(); }} />}</Drop>
      </Group>
      <Group label="Illustrations">
        <Drop icon={Image} title="Picture" wide label="Picture">{(close) => (
          <>
            <MenuItem icon={ImageUp} onClick={() => { close(); act.imageFile(); }}>From this device …</MenuItem>
            <MenuItem icon={Link} onClick={() => { close(); act.imageUrl(); }}>From a web address …</MenuItem>
            <MenuItem icon={ClipboardPaste} onClick={() => { close(); act.imageClipboard(); }}>From the clipboard</MenuItem>
          </>
        )}</Drop>
        <Drop icon={SquareDashed} title="Text box / call-out" wide label="Text box">{(close) => (
          <>
            {[['info', 'Info box'], ['note', 'Neutral box'], ['ok', 'Success box'], ['warn', 'Warning box']].map(([k, l]) => <MenuItem key={k} onClick={() => { c().insertCallout(k).run(); close(); }}>{l}</MenuItem>)}
          </>
        )}</Drop>
      </Group>
      <Group label="Links">
        <RB icon={Link} title="Link – Ctrl+K is Hostwerk's search" keys="Ctrl+Shift+K" wide label="Link" onClick={act.link} />
        <RB icon={MessageSquarePlus} title="New comment" keys="Ctrl+Alt+M" wide label="Comment" onClick={act.newComment} disabled={e.state.selection.empty} />
      </Group>
      <Group label="References">
        <RB icon={Superscript} title="Footnote" keys="Ctrl+Alt+F" wide label="Footnote" onClick={act.footnote} />
        <RB icon={TableOfContents} title="Table of contents" wide label="Contents" onClick={act.toc} />
      </Group>
      <Group label="Header & footer">
        <RB icon={PanelTop} title="Header, footer and page numbers" wide label="Header & footer" onClick={act.headerFooter} />
      </Group>
      <Group label="Symbols">
        <RB icon={Minus} title="Horizontal line" onClick={() => c().setHorizontalRule().run()} />
        <RB icon={Omega} title="Symbol" onClick={act.symbols} />
        <Drop icon={Smile} title="Emoji" panelClass="of-pop-emoji">{(close) => <EmojiGrid onPick={(em) => { c().insertContent(em).run(); close(); }} />}</Drop>
        <Drop icon={CalendarClock} title="Date and time">{(close) => (
          <>
            <MenuItem icon={RefreshCw} onClick={() => { close(); act.quickField('date'); }}>Date that updates itself</MenuItem>
            <hr />
            {dateFormats(document.documentElement.lang).map((d) => <MenuItem key={d} onClick={() => { c().insertContent(d).run(); close(); }}>{d}</MenuItem>)}
          </>
        )}</Drop>
      </Group>
      <Group label="Formulas & diagrams">
        <RB icon={Sigma} title="Formula (LaTeX) – or type $…$ in the text" wide label="Formula" onClick={() => act.equation()} />
        <RB icon={Workflow} title="Diagram from text (flowchart, sequence, Gantt …)" wide label="Diagram" onClick={() => act.diagram()} />
      </Group>
      <Group label="Fields">
        <RB icon={Braces} title="Live field: date, page number, a Hostwerk value or a cell of a spreadsheet" wide label="Field" onClick={() => act.field()} />
        <RB icon={RefreshCw} title="Update all fields" keys="F9" onClick={act.refreshFields} />
        <Drop icon={CheckSquare} title="Form controls (Word content controls)" wide label="Form">{(close) => (
          <>
            <MenuItem icon={CheckSquare} onClick={() => { close(); act.control('checkbox'); }}>Check box</MenuItem>
            <MenuItem icon={ListOrdered} onClick={() => { close(); act.control('dropdown'); }}>Drop-down list</MenuItem>
            <MenuItem icon={CalendarClock} onClick={() => { close(); act.control('date'); }}>Date picker</MenuItem>
            <MenuItem icon={Type} onClick={() => { close(); act.control('text'); }}>Text field</MenuItem>
          </>
        )}</Drop>
      </Group>
      {s.ai && <Group label="AI"><RB icon={Sparkles} title="Write with AI" wide label="AI" onClick={() => act.ai('write')} /></Group>}
    </>
  );

  const p = s.page;
  const layout = (
    <>
      <Group label="Page setup">
        <Drop icon={FileText} title="Size" wide label="Size">{(close) => ['A4', 'A5', 'A3', 'Letter', 'Legal'].map((k) => <MenuItem key={k} active={p.size === k} onClick={() => { act.pageSize(k); close(); }}>{k}</MenuItem>)}</Drop>
        <Drop icon={p.orientation === 'landscape' ? RectangleHorizontal : RectangleVertical} title="Orientation" wide label="Orientation">{(close) => (
          <>
            <MenuItem icon={RectangleVertical} active={p.orientation !== 'landscape'} onClick={() => { act.page({ orientation: 'portrait' }); close(); }}>Portrait</MenuItem>
            <MenuItem icon={RectangleHorizontal} active={p.orientation === 'landscape'} onClick={() => { act.page({ orientation: 'landscape' }); close(); }}>Landscape</MenuItem>
          </>
        )}</Drop>
        <Drop icon={Maximize} title="Margins" wide label="Margins">{(close) => (
          <>
            {MARGINS.map((m) => <MenuItem key={m.id} active={['top', 'right', 'bottom', 'left'].every((k) => Math.abs(p.margins[k] - m.m[k]) < 0.6)} onClick={() => { act.page({ margins: m.m }); close(); }}>{t(m.label)} <span className="muted small">{`${m.m.top / 10} · ${m.m.left / 10} cm`}</span></MenuItem>)}
            <hr />
            <MenuItem onClick={() => { close(); act.marginsDialog(); }}>Custom margins …</MenuItem>
          </>
        )}</Drop>
        <Drop icon={SquareSplitVertical} title="Breaks" wide label="Breaks">{(close) => (
          <>
            <MenuItem icon={SeparatorHorizontal} keys="Ctrl+Enter" onClick={() => { close(); act.pageBreak(); }}>Page break</MenuItem>
            <MenuItem icon={SquareSplitVertical} onClick={() => { close(); act.sectionBreak('nextPage'); }}>Section break, next page</MenuItem>
            <MenuItem icon={ChevronsUpDown} onClick={() => { close(); act.sectionBreak('continuous'); }}>Section break, continuous</MenuItem>
          </>
        )}</Drop>
        <Drop icon={Columns2} title="Columns – the text flows in columns (pages break in one column here, exports are exact)" wide label="Columns">{(close) => [1, 2, 3].map((n) => <MenuItem key={n} active={(p.columns || 1) === n} onClick={() => { act.page({ columns: n }); close(); }}>{n === 1 ? t('One') : n === 2 ? t('Two') : t('Three')}</MenuItem>)}</Drop>
      </Group>
      <Group label="Page background">
        <Drop icon={PaintBucket} title="Page colour" wide label="Page colour" panelClass="of-pop-colors">{(close) => <ColorGrid close={close} noneLabel="No colour" onPick={(col) => act.page({ color: col })} />}</Drop>
      </Group>
      <Group label="Paragraph">
        <RB icon={IndentDecrease} title="Decrease indent" onClick={() => c().changeIndent(-36).run()} />
        <RB icon={IndentIncrease} title="Increase indent" onClick={() => c().changeIndent(36).run()} />
        <RB icon={SeparatorHorizontal} title="Spacing and indents …" wide label="Spacing …" onClick={act.paragraphDialog} />
      </Group>
    </>
  );

  const review = (
    <>
      <Group label="Proofing">
        <RB icon={SpellCheck} title="Spell check (browser)" active={s.spell} onClick={act.toggleSpell} wide label="Spelling" />
        <Drop icon={Languages} title="Language for spell check" wide label={s.lang}>{(close) => LANGS.map(([id, label]) => <MenuItem key={id} active={s.lang === id} onClick={() => { act.setLang(id); close(); }}>{label}</MenuItem>)}</Drop>
        <RB icon={FileDigit} title="Word count" wide label="Word count" onClick={act.wordCount} />
      </Group>
      <Group label="Comments">
        {!readOnly && <RB icon={MessageSquarePlus} title="New comment" keys="Ctrl+Alt+M" wide label="New comment" onClick={act.newComment} disabled={e.state.selection.empty} />}
        <RB icon={PanelRight} title="Show comments" active={s.comments} onClick={() => act.side('comments')} wide label="Show comments" />
      </Group>
      <Group label="Tracking">
        {!readOnly && <RB icon={FilePen} title="Track changes – record insertions and deletions" keys="Ctrl+Shift+E" active={s.track} wide label="Track changes" onClick={act.toggleTrack} />}
        <RB icon={ChevronLeft} title="Previous change" onClick={() => c().prevChange().run()} disabled={!s.changes} />
        <RB icon={ChevronRight} title="Next change" onClick={() => c().nextChange().run()} disabled={!s.changes} />
        {!readOnly && <RB icon={Check} title="Accept change" onClick={() => c().acceptChange().run()} disabled={!s.changes} />}
        {!readOnly && <RB icon={X} title="Reject change" onClick={() => c().rejectChange().run()} disabled={!s.changes} />}
        {!readOnly && <Drop title="All changes" wide label="All" disabled={!s.changes}>{(close) => (
          <>
            <MenuItem icon={Check} onClick={() => { c().acceptAllChanges().run(); close(); }}>Accept all changes</MenuItem>
            <MenuItem icon={X} onClick={() => { c().rejectAllChanges().run(); close(); }}>Reject all changes</MenuItem>
          </>
        )}</Drop>}
        <RB icon={PanelRight} title="Show changes" onClick={() => act.side('changes')} />
      </Group>
      <Group label="Compare & share">
        <RB icon={GitCompare} title="Compare with an older version or another file" wide label="Compare" onClick={act.compare} />
        {s.canCollab && <RB icon={Users} title={s.collab ? 'Stop editing together' : 'Edit together – everyone who opens the file edits live with you'} active={s.collab} disabled={s.joining} wide label={s.collab ? 'Together' : 'Edit together'} onClick={act.collab} />}
      </Group>
      {s.hasMacros && <Group label="Macros"><RB icon={SquareCode} title="Show the macros of this file (read only)" wide label="Macros" onClick={act.macros} /></Group>}
    </>
  );

  const view = (
    <>
      <Group label="Views">
        <RB icon={FileText} title="Print layout" active={s.layout === 'print'} wide label="Print layout" onClick={() => act.setLayout('print')} />
        <RB icon={Type} title="Web layout" active={s.layout === 'web'} wide label="Web layout" onClick={() => act.setLayout('web')} />
        <RB icon={Focus} title="Focus mode (Esc to leave)" wide label="Focus" onClick={act.focusMode} />
      </Group>
      <Group label="Show">
        <RB icon={Ruler} title="Ruler" active={s.ruler} wide label="Ruler" onClick={act.toggleRuler} />
        <RB icon={PanelLeft} title="Navigation pane" active={s.nav} wide label="Navigation" onClick={act.toggleNav} />
        <RB icon={PanelRight} title="Comments and changes" active={s.comments} wide label="Comments" onClick={() => act.side('comments')} />
        <RB icon={LayoutGrid} title="Page thumbnails" active={s.thumbs} wide label="Pages" onClick={act.thumbs} />
        <RB icon={Pilcrow} title="Formatting marks" active={s.marks} onClick={act.toggleMarks} />
        <RB icon={ChevronUp} title="Small ribbon (one line)" active={s.mini} onClick={act.toggleMini} />
      </Group>
      <Group label="Zoom">
        <RB icon={ZoomOut} title="Zoom out" onClick={() => act.zoom(s.zoom - 10)} />
        <RB title="100 %" wide label="100 %" onClick={() => act.zoom(100)} />
        <RB icon={ZoomIn} title="Zoom in" onClick={() => act.zoom(s.zoom + 10)} />
        <RB title="Page width" wide label="Page width" onClick={act.fitWidth} />
      </Group>
    </>
  );

  const table = inTable && (
    <>
      <Group label="Rows & columns">
        <RB icon={BetweenHorizontalStart} title="Insert row above" onClick={() => c().addRowBefore().run()} />
        <RB icon={BetweenHorizontalEnd} title="Insert row below" onClick={() => c().addRowAfter().run()} />
        <RB icon={BetweenVerticalStart} title="Insert column left" onClick={() => c().addColumnBefore().run()} />
        <RB icon={BetweenVerticalEnd} title="Insert column right" onClick={() => c().addColumnAfter().run()} />
        <Drop icon={Trash2} title="Delete">{(close) => (
          <>
            <MenuItem onClick={() => { act.deleteRow(); close(); }}>Delete row</MenuItem>
            <MenuItem onClick={() => { c().deleteColumn().run(); close(); }}>Delete column</MenuItem>
            <MenuItem onClick={() => { c().deleteTable().run(); close(); }}>Delete table</MenuItem>
          </>
        )}</Drop>
      </Group>
      <Group label="Merge">
        <RB icon={TableCellsMerge} title="Merge cells" disabled={!e.can().mergeCells()} onClick={() => c().mergeCells().run()} />
        <RB icon={TableCellsSplit} title="Split cell" disabled={!e.can().splitCell()} onClick={() => c().splitCell().run()} />
      </Group>
      <Group label="Format">
        <RB title="Header row" wide label="Header row" onClick={() => c().toggleHeaderRow().run()} />
        <Drop icon={PaintBucket} title="Cell background" panelClass="of-pop-colors">{(close) => <ColorGrid close={close} noneLabel="No colour" onPick={(col) => c().setCellAttribute('backgroundColor', col).run()} />}</Drop>
        <RB title="Distribute columns evenly" wide label="Even columns" onClick={act.evenColumns} />
        <RB title="Column width …" wide label="Width …" onClick={act.columnWidth} />
      </Group>
    </>
  );

  const picture = img && (
    <>
      <Group label="Position">
        <RB icon={AlignLeft} title="In line with text" active={!img.align && !img.wrap} onClick={() => act.image({ align: null, wrap: null })} />
        <RB icon={AlignCenter} title="Centred on its own line" active={img.align === 'center'} onClick={() => act.image({ align: 'center', wrap: null })} />
        <RB title="Left, text flows around" wide label="Left, wrapped" active={img.align === 'left' && img.wrap === 'square'} onClick={() => act.image({ align: 'left', wrap: 'square' })} />
        <RB title="Right, text flows around" wide label="Right, wrapped" active={img.align === 'right' && img.wrap === 'square'} onClick={() => act.image({ align: 'right', wrap: 'square' })} />
      </Group>
      <Group label="Picture size">
        {[25, 50, 75, 100].map((pc) => <RB key={pc} title={`${pc} % of the line width`} wide label={`${pc} %`} onClick={() => act.image({ width: Math.round((s.contentWidth * pc) / 100) })} />)}
        <RB title="Exact width …" wide label="Width …" onClick={act.imageWidth} />
      </Group>
      <Group label="Picture">
        <RB title="Alternative text for screen readers" wide label="Alt text …" onClick={act.altText} />
        <RB icon={ImageUp} title="Replace picture …" onClick={act.imageFile} />
        <RB icon={Trash2} title="Delete picture" onClick={() => c().deleteSelection().run()} />
      </Group>
    </>
  );

  const link = !readOnly && e.isActive('link');
  const onTabKey = (ev) => {
    if (ev.key !== 'ArrowRight' && ev.key !== 'ArrowLeft') return;
    const k = tabs.findIndex((x) => x.id === cur);
    const n = tabs[(k + (ev.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length];
    setTab(n.id);
    const el = ev.currentTarget.querySelector(`[data-tab="${n.id}"]`); if (el) el.focus();
  };
  return (
    <div className={`of-docs-ribbon ${s.mini ? 'of-mini' : ''}`}>
      <div className="of-docs-tabsrow">
        {!readOnly && (
          <div className="of-docs-qat" role="toolbar" aria-label={t('Quick access')}>
            <RB icon={Undo2} title="Undo" keys="Ctrl+Z" disabled={!e.can().undo()} onClick={() => c().undo().run()} />
            <RB icon={Redo2} title="Redo" keys="Ctrl+Y" disabled={!e.can().redo()} onClick={() => c().redo().run()} />
            {host.save && <RB icon={Save} title="Save" keys="Ctrl+S" onClick={() => host.save()} />}
            <RB icon={Printer} title="Print" keys="Ctrl+P" onClick={act.print} />
          </div>
        )}
        <div className="of-docs-tabs" role="tablist" aria-label={t('Ribbon')} onKeyDown={onTabKey}>
          {tabs.map((x) => (
            <button key={x.id} type="button" role="tab" data-tab={x.id} aria-selected={cur === x.id} tabIndex={cur === x.id ? 0 : -1} className={`of-docs-tab ${cur === x.id ? 'of-on' : ''} ${x.ctx ? 'of-ctx-tab' : ''}`}
              onMouseDown={(ev) => ev.preventDefault()} onClick={() => setTab(x.id)} onDoubleClick={act.toggleMini}>{x.label}</button>
          ))}
        </div>
        <span className="grow" />
        {link && <div className="of-docs-linkbar small"><Link size={13} /><span className="mono of-docs-linkurl">{e.getAttributes('link').href}</span><RB icon={Link} title="Edit link" onClick={act.link} /><RB icon={Unlink} title="Remove link" onClick={() => c().extendMarkRange('link').unsetLink().run()} /></div>}
        <RB icon={ChevronUp} className={`of-docs-minbtn ${s.mini ? 'of-flip' : ''}`} title={s.mini ? 'Show the full ribbon' : 'Small ribbon (one line) – double-click a tab does the same'} onClick={act.toggleMini} />
      </div>
      <div className="of-docs-groups" key={cur}>
        {cur === 'home' && home}
        {cur === 'insert' && insert}
        {cur === 'layout' && layout}
        {cur === 'review' && review}
        {cur === 'view' && view}
        {cur === 'table' && table}
        {cur === 'picture' && picture}
      </div>
    </div>
  );
}
void Code;
