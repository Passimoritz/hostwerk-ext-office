/* The ribbon: Home, Insert, Design, Transitions, Animations, Slide show, Review, View (+ contextual Format tab). */
import { useState } from 'react';
import {
  Plus, LayoutTemplate, RotateCcw, FolderPlus, ClipboardPaste, Scissors, Copy, Paintbrush, Bold, Italic, Underline, Strikethrough, Superscript, Subscript,
  RemoveFormatting, Baseline, Highlighter, CaseSensitive, AArrowUp, AArrowDown, List, ListOrdered, IndentDecrease, IndentIncrease, AlignLeft, AlignCenter,
  AlignRight, AlignJustify, AlignVerticalJustifyStart, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd, BringToFront, SendToBack, Group, Ungroup,
  AlignStartVertical, AlignCenterVertical, AlignEndVertical, AlignStartHorizontal, AlignCenterHorizontal, AlignEndHorizontal, AlignHorizontalDistributeCenter,
  AlignVerticalDistributeCenter, RotateCw, FlipHorizontal2, FlipVertical2, PaintBucket, PenLine, Shapes, Type, Image, ImagePlus, Clipboard, Link2, Table,
  ChartColumn, ChartBar, ChartLine, ChartPie, Smile, Omega, PanelBottom, Film, Palette, Monitor, Presentation, Timer, Play, Repeat, EyeOff, SpellCheck,
  MessageSquare, MessageSquarePlus, ZoomIn, ZoomOut, Maximize, Grid3x3, Ruler, Magnet, StickyNote, PanelLeft, PanelRight, LayoutList, Sparkles, Layers,
  Square, Trash2, Lock, Grid2x2, TableCellsMerge, TableCellsSplit, BetweenHorizontalStart, BetweenHorizontalEnd, BetweenVerticalStart, BetweenVerticalEnd,
  Text, MonitorPlay, ListTree, Undo2, Redo2, ArrowUp, ArrowDown, Volume2, SquareDashed, Wand2,
} from 'lucide-react';
import { createElement as h } from 'react';
import { t } from '@hostwerk/i18n';
import { RB, Group as G, Pop, Menu, ColorPop, ColorGrid, Num, ShapeGallery, ShapeIcon } from './controls.jsx';
import SlideView from './render/SlideView.jsx';
import { THEMES, LAYOUT_KINDS, themeData, buildLayouts, css } from './themes.js';
import { SIZES, S } from './units.js';
import { EFFECTS } from './Panels.jsx';
import { newDoc, newSlide } from './model.js';
import { Raw } from '../_shared/keep.jsx';

/** built-in layouts (id "layout-<kind>") carry English names to translate; layouts from a file keep their own name */
const layoutName = (l) => (String(l.id || '').startsWith('layout-') ? t(l.name || '') : String(l.name || ''));

const TABS = [
  { id: 'home', label: 'Home' }, { id: 'insert', label: 'Insert' }, { id: 'design', label: 'Design' }, { id: 'transitions', label: 'Transitions' },
  { id: 'animations', label: 'Animations' }, { id: 'show', label: 'Slide show' }, { id: 'review', label: 'Review' }, { id: 'view', label: 'View' },
];
export const FONTS = ['Aptos', 'Aptos Display', 'Arial', 'Arial Black', 'Bahnschrift', 'Calibri', 'Calibri Light', 'Cambria', 'Candara', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Gabriola', 'Garamond', 'Georgia', 'Impact', 'Lucida Sans Unicode', 'Palatino Linotype', 'Segoe Print', 'Segoe Script', 'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold', 'Sitka Text', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Liberation Sans', 'Liberation Serif', 'DejaVu Sans', 'Noto Sans', 'Noto Serif', 'Open Sans', 'Roboto', 'Helvetica', 'system-ui'];
export const SIZES_PT = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
const BULLETS = ['•', '◦', '▪', '–', '➢', '✓', '★', '◆', '→'];
const NUMBERS = [['arabicPeriod', '1. 2. 3.'], ['arabicParenR', '1) 2) 3)'], ['alphaUcPeriod', 'A. B. C.'], ['alphaLcPeriod', 'a. b. c.'], ['alphaLcParenR', 'a) b) c)'], ['romanUcPeriod', 'I. II. III.'], ['romanLcPeriod', 'i. ii. iii.']];
const SPACING = [100, 115, 150, 200, 250, 300];
export const TRANSITIONS = [{ id: 'none', label: 'None' }, { id: 'fade', label: 'Fade' }, { id: 'push', label: 'Push' }, { id: 'wipe', label: 'Wipe' }, { id: 'cover', label: 'Cover' }, { id: 'split', label: 'Split' }, { id: 'zoom', label: 'Zoom' }];
const DIRS = [{ id: 'l', label: 'From right' }, { id: 'r', label: 'From left' }, { id: 'u', label: 'From bottom' }, { id: 'd', label: 'From top' }];
const FLY_DIRS = [{ id: 'b', label: 'From bottom' }, { id: 't', label: 'From top' }, { id: 'l', label: 'From left' }, { id: 'r', label: 'From right' }];
const EMOJI = '😀 😃 😄 😁 😊 😍 🤩 😎 🤔 😮 😢 😡 👍 👎 👏 🙌 🙏 💪 👋 ✌️ 🤝 ❤️ 🧡 💛 💚 💙 💜 ⭐ 🌟 ✨ 🔥 💡 🎯 🚀 🏆 🎉 🎁 📈 📉 📊 📌 📎 📝 📅 ⏰ ✅ ❌ ⚠️ ❓ ❗ ➕ ➖ ➡️ ⬅️ ⬆️ ⬇️ 🔒 🔑 💬 📣 📷 🎵 🌍 🏠 🏢 🚗 ✈️ 🍀 ☀️ 🌙 ☁️ ⚡ 💧 🌱 🍎 ☕ 💰 💶 💳 🛒 🧩 ⚙️ 🔧 🧪 📚 🎓 💻 📱 🖨️ 🔔'.split(' ');
const SYMBOLS = '→ ← ↑ ↓ ↔ ⇒ ⇔ • · … – — ° ± × ÷ ≠ ≈ ≤ ≥ ∞ √ ∑ ∏ ∫ ∆ π Ω α β γ δ ε λ µ σ φ € £ ¥ ¢ $ © ® ™ § ¶ † ‡ ‰ ½ ¼ ¾ ² ³ ✓ ✗ ★ ☆ ♥ ♦ ♣ ♠ ■ □ ● ○ ▲ △ ◆ ◇ ☐ ☑ ☒ « » „ “ ” ‚ ‘ ’'.split(' ');
const QUICK = [
  { fill: 'accent1', line: null, text: 'lt1' }, { fill: 'accent2', line: null, text: 'lt1' }, { fill: 'accent3', line: null, text: 'lt1' }, { fill: 'accent4', line: null, text: 'tx1' },
  { fill: 'accent1+80', line: 'accent1', text: 'accent1-50' }, { fill: 'accent2+80', line: 'accent2', text: 'accent2-50' }, { fill: 'accent3+80', line: 'accent3', text: 'accent3-50' }, { fill: 'bg1', line: 'accent1', text: 'accent1' },
  { fill: null, line: 'accent1', text: 'tx1' }, { fill: 'tx1', line: null, text: 'bg1' }, { fill: 'grad1', line: null, text: 'lt1' }, { fill: 'grad2', line: null, text: 'lt1' },
];

function LayoutPicker({ doc, onPick, close }) {
  const W = S(doc.size.w);
  const scale = 112 / W;
  return (
    <div className="of-slides-layouts">
      {(doc.layouts || []).map((l) => {
        const s = newSlide(doc, l.kind);
        s.layoutId = l.id;
        return (
          <Raw as="button" key={l.id} type="button" className="of-slides-layoutbtn" onClick={() => { close(); onPick(l.kind, l.id); }} title={layoutName(l)}>
            <SlideView doc={doc} slide={s} scale={scale} mode="edit" prompt={(e) => t(e.prompt || '')} />
            {h('span', { className: 'small' }, layoutName(l))}
          </Raw>
        );
      })}
    </div>
  );
}

function ThemeGallery({ doc, onPick, close }) {
  return (
    <div className="of-slides-themes">
      {THEMES.map((th) => {
        const td = { ...newDoc({ size: doc.size, theme: th.id }) };
        const s = newSlide(td, 'title', { ctrTitle: th.name, subTitle: 'Aa' });
        return (
          <Raw as="button" key={th.id} type="button" className={`of-slides-themebtn ${doc.theme.id === th.id ? 'on' : ''}`} title={th.name} onClick={() => { close(); onPick(th.id); }}>
            <SlideView doc={td} slide={s} scale={150 / S(doc.size.w)} mode="thumb" />
            {h('span', { className: 'small' }, th.name)}
          </Raw>
        );
      })}
    </div>
  );
}

function TablePicker({ onPick, close }) {
  const [hv, setHv] = useState({ r: 0, c: 0 });
  return (
    <div className="of-slides-tablepick">
      <div className="small">{hv.r ? t('{c} × {r} table', { c: hv.c, r: hv.r }) : t('Insert table')}</div>
      <div className="of-slides-tablegrid" onMouseLeave={() => setHv({ r: 0, c: 0 })}>
        {Array.from({ length: 8 }, (_, r) => Array.from({ length: 10 }, (__, c) => (
          <button key={`${r}-${c}`} type="button" aria-label={`${c + 1} × ${r + 1}`} className={r < hv.r && c < hv.c ? 'on' : ''} onMouseEnter={() => setHv({ r: r + 1, c: c + 1 })} onClick={() => { close(); onPick(r + 1, c + 1); }} />
        )))}
      </div>
    </div>
  );
}

export default function Ribbon({ c, tab, setTab }) {
  const { doc, slide, selEls, fmt, readOnly, state } = c;
  const has = selEls.length > 0;
  const hasText = has && selEls.some((e) => e.text || e.type === 'table');
  const one = selEls.length === 1 ? selEls[0] : null;
  const ro = readOnly;
  const theme = doc.theme;
  const ctxTab = one && ['table', 'chart', 'image'].includes(one.type) ? one.type : has ? 'shape' : null;
  const tabs = [...TABS, ...(ctxTab && !ro ? [{ id: 'format', label: ctxTab === 'table' ? 'Table' : ctxTab === 'chart' ? 'Chart' : ctxTab === 'image' ? 'Picture' : 'Shape format', ctx: true }] : [])];
  const cur = tabs.some((x) => x.id === tab) ? tab : 'home';
  const tr = (slide && slide.transition) || { type: 'none' };
  const selAnim = one && slide ? (slide.animations || []).find((a) => a.el === one.id) : null;

  const fontGroup = (
    <G label={t('Font')}>
      <div className="of-slides-rrow">
        <select className="input of-slides-fontsel" aria-label={t('Font')} title={t('Font')} disabled={ro || !hasText} value={fmt.fontRaw || ''} onChange={(e) => c.fontFamily(e.target.value)}>
          <option value="" disabled>{fmt.font || ''}</option>
          <optgroup label={t('Theme fonts')}><option value="+mj">{`${theme.fonts.major} (${t('Headings')})`}</option><option value="+mn">{`${theme.fonts.minor} (${t('Body')})`}</option></optgroup>
          <optgroup label={t('All fonts')}>{FONTS.map((f) => <option key={f} value={f}>{f}</option>)}</optgroup>
        </select>
        <FontSize value={fmt.size} disabled={ro || !hasText} onChange={c.fontSize} />
        <RB icon={AArrowUp} title={t('Increase font size')} kbd="Ctrl+]" disabled={ro || !hasText} onClick={() => c.growFont(1)} />
        <RB icon={AArrowDown} title={t('Decrease font size')} kbd="Ctrl+[" disabled={ro || !hasText} onClick={() => c.growFont(-1)} />
        <RB icon={RemoveFormatting} title={t('Clear formatting')} disabled={ro || !hasText} onClick={c.clearFormat} />
      </div>
      <div className="of-slides-rrow">
        <RB icon={Bold} title={t('Bold')} kbd="Ctrl+B" active={fmt.b} disabled={ro || !hasText} onClick={() => c.toggleRun('b')} />
        <RB icon={Italic} title={t('Italic')} kbd="Ctrl+I" active={fmt.i} disabled={ro || !hasText} onClick={() => c.toggleRun('i')} />
        <RB icon={Underline} title={t('Underline')} kbd="Ctrl+U" active={fmt.u} disabled={ro || !hasText} onClick={() => c.toggleRun('u')} />
        <RB icon={Strikethrough} title={t('Strikethrough')} active={fmt.s} disabled={ro || !hasText} onClick={() => c.toggleRun('s')} />
        <RB icon={Superscript} title={t('Superscript')} active={fmt.baseline > 0} disabled={ro || !hasText} onClick={() => c.baseline(30)} />
        <RB icon={Subscript} title={t('Subscript')} active={fmt.baseline < 0} disabled={ro || !hasText} onClick={() => c.baseline(-25)} />
        <Pop icon={CaseSensitive} title={t('Change case')} disabled={ro || !hasText} content={(close) => <Menu close={close} items={[{ label: t('Sentence case.'), onClick: () => c.changeCase('sentence') }, { label: t('lowercase'), onClick: () => c.changeCase('lower') }, { label: t('UPPERCASE'), onClick: () => c.changeCase('upper') }, { label: t('Capitalize Each Word'), onClick: () => c.changeCase('title') }]} />} />
        <ColorPop icon={Baseline} title={t('Font colour')} theme={theme} value={fmt.color} allowNone={false} disabled={ro || !hasText} onPick={c.fontColor} />
        <ColorPop icon={Highlighter} title={t('Text highlight colour')} theme={theme} value={fmt.highlight} noneLabel="No colour" disabled={ro || !hasText} onPick={c.highlight} />
      </div>
    </G>
  );
  const paraGroup = (
    <G label={t('Paragraph')}>
      <div className="of-slides-rrow">
        <Pop icon={List} title={t('Bullets')} split active={fmt.bullet === 'char'} onMain={() => c.bullets(fmt.bullet === 'char' ? false : '•')} disabled={ro || !hasText} content={(close) => (
          <div className="of-slides-bulpick"><button type="button" className="of-slides-menu-item" onClick={() => { close(); c.bullets(false); }}>{t('None')}</button>
            <div className="of-slides-bulgrid">{BULLETS.map((b) => <button key={b} type="button" className="of-slides-bulbtn" onClick={() => { close(); c.bullets(b); }}>{b}</button>)}</div></div>)} />
        <Pop icon={ListOrdered} title={t('Numbering')} split active={fmt.bullet === 'num'} onMain={() => c.numbering(fmt.bullet === 'num' ? false : 'arabicPeriod')} disabled={ro || !hasText} content={(close) => <Menu close={close} items={[{ label: t('None'), onClick: () => c.numbering(false) }, ...NUMBERS.map(([k, l]) => ({ label: l, onClick: () => c.numbering(k) }))]} />} />
        <RB icon={IndentDecrease} title={t('Decrease list level')} kbd="Shift+Tab" disabled={ro || !hasText} onClick={() => c.level(-1)} />
        <RB icon={IndentIncrease} title={t('Increase list level')} kbd="Tab" disabled={ro || !hasText} onClick={() => c.level(1)} />
        <Pop icon={Text} title={t('Line spacing')} disabled={ro || !hasText} content={(close) => <Menu close={close} items={SPACING.map((v) => ({ label: (v / 100).toFixed(v % 100 ? 2 : 1).replace(/0$/, ''), active: fmt.lnSpc === v, onClick: () => c.lineSpacing(v) }))} />} />
      </div>
      <div className="of-slides-rrow">
        <RB icon={AlignLeft} title={t('Align left')} kbd="Ctrl+L" active={fmt.align === 'l'} disabled={ro || !hasText} onClick={() => c.align('l')} />
        <RB icon={AlignCenter} title={t('Center')} kbd="Ctrl+E" active={fmt.align === 'ctr'} disabled={ro || !hasText} onClick={() => c.align('ctr')} />
        <RB icon={AlignRight} title={t('Align right')} kbd="Ctrl+R" active={fmt.align === 'r'} disabled={ro || !hasText} onClick={() => c.align('r')} />
        <RB icon={AlignJustify} title={t('Justify')} kbd="Ctrl+J" active={fmt.align === 'just'} disabled={ro || !hasText} onClick={() => c.align('just')} />
        <RB icon={AlignVerticalJustifyStart} title={t('Align top')} active={fmt.anchor === 't'} disabled={ro || !hasText} onClick={() => c.anchor('t')} />
        <RB icon={AlignVerticalJustifyCenter} title={t('Align middle')} active={fmt.anchor === 'ctr'} disabled={ro || !hasText} onClick={() => c.anchor('ctr')} />
        <RB icon={AlignVerticalJustifyEnd} title={t('Align bottom')} active={fmt.anchor === 'b'} disabled={ro || !hasText} onClick={() => c.anchor('b')} />
      </div>
    </G>
  );
  const arrangeMenu = (close) => <Menu close={close} items={[
    { icon: BringToFront, label: t('Bring to front'), hint: 'Ctrl+Shift+]', onClick: () => c.arrange('front') }, { icon: ArrowUp, label: t('Bring forward'), hint: 'Ctrl+]', onClick: () => c.arrange('forward') },
    { icon: ArrowDown, label: t('Send backward'), hint: 'Ctrl+[', onClick: () => c.arrange('backward') }, { icon: SendToBack, label: t('Send to back'), hint: 'Ctrl+Shift+[', onClick: () => c.arrange('back') }, '-',
    { icon: Group, label: t('Group'), hint: 'Ctrl+G', disabled: selEls.length < 2, onClick: c.group }, { icon: Ungroup, label: t('Ungroup'), hint: 'Ctrl+Shift+G', disabled: !selEls.some((e) => e.type === 'group'), onClick: c.ungroup }, '-',
    { icon: AlignStartVertical, label: t('Align left'), onClick: () => c.alignEls('l') }, { icon: AlignCenterVertical, label: t('Align center'), onClick: () => c.alignEls('c') }, { icon: AlignEndVertical, label: t('Align right'), onClick: () => c.alignEls('r') },
    { icon: AlignStartHorizontal, label: t('Align top'), onClick: () => c.alignEls('t') }, { icon: AlignCenterHorizontal, label: t('Align middle'), onClick: () => c.alignEls('m') }, { icon: AlignEndHorizontal, label: t('Align bottom'), onClick: () => c.alignEls('b') },
    { icon: AlignHorizontalDistributeCenter, label: t('Distribute horizontally'), disabled: selEls.length < 3, onClick: () => c.distribute('h') }, { icon: AlignVerticalDistributeCenter, label: t('Distribute vertically'), disabled: selEls.length < 3, onClick: () => c.distribute('v') }, '-',
    { icon: RotateCw, label: t('Rotate right 90°'), onClick: () => c.rotate(90) }, { icon: RotateCcw, label: t('Rotate left 90°'), onClick: () => c.rotate(-90) },
    { icon: FlipHorizontal2, label: t('Flip horizontal'), onClick: () => c.flip('h') }, { icon: FlipVertical2, label: t('Flip vertical'), onClick: () => c.flip('v') }, '-',
    { icon: Lock, label: selEls.some((e) => e.locked) ? t('Unlock') : t('Lock'), onClick: c.toggleLock },
  ]} />;
  const drawingGroup = (
    <G label={t('Drawing')}>
      <div className="of-slides-quickshapes">
        {['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow', 'star5', 'wedgeRoundRectCallout', 'line', 'arrow'].map((g) => <RB key={g} title={t(g === 'line' ? 'Line' : g === 'arrow' ? 'Arrow' : 'Shape')} onClick={() => c.setTool(g)} active={state.tool === g} disabled={ro}><ShapeIcon geom={g} size={16} /></RB>)}
        <Pop icon={Shapes} title={t('More shapes')} disabled={ro} width={300} content={(close) => <ShapeGallery close={close} onPick={(g) => c.setTool(g)} />} />
      </div>
      <div className="of-slides-rrow">
        <Pop icon={Layers} label={t('Arrange')} disabled={ro || !has} content={arrangeMenu} />
        <Pop icon={Wand2} label={t('Quick styles')} disabled={ro || !has} width={220} content={(close) => (
          <div className="of-slides-quickstyles">{QUICK.map((q, i) => <button key={i} type="button" className="of-slides-qs" title={t('Style {n}', { n: i + 1 })} aria-label={t('Style {n}', { n: i + 1 })} onClick={() => { close(); c.quickStyle(q); }}
            style={{ background: q.fill === 'grad1' ? `linear-gradient(135deg, ${css('accent1', theme)}, ${css('accent2', theme)})` : q.fill === 'grad2' ? `linear-gradient(135deg, ${css('accent3', theme)}, ${css('accent5', theme)})` : q.fill ? css(q.fill, theme) : 'transparent', border: `2px solid ${q.line ? css(q.line, theme) : 'transparent'}`, color: css(q.text, theme) }}>Abc</button>)}</div>)} />
      </div>
      <div className="of-slides-rrow">
        <ColorPop icon={PaintBucket} title={t('Shape fill')} theme={theme} value={fmt.fill} noneLabel="No fill" disabled={ro || !has} onPick={c.shapeFill} />
        <ColorPop icon={PenLine} title={t('Shape outline')} theme={theme} value={fmt.line} noneLabel="No outline" disabled={ro || !has} onPick={c.shapeLine} />
        <Pop icon={SquareDashed} title={t('Outline weight and dashes')} disabled={ro || !has} content={(close) => <Menu close={close} items={[...[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => ({ label: `${w} pt`, onClick: () => c.lineWidth(w) })), '-', ...[['solid', 'Solid'], ['dash', 'Dashed'], ['dot', 'Dotted'], ['dashDot', 'Dash dot'], ['lgDash', 'Long dash']].map(([d, l]) => ({ label: t(l), onClick: () => c.lineDash(d) }))]} />} />
        <RB icon={Square} title={t('Shadow')} active={fmt.shadow} disabled={ro || !has} onClick={c.toggleShadow} />
        <RB icon={PanelRight} title={t('Format pane')} active={state.pane === 'format'} onClick={() => c.togglePane('format')} />
      </div>
    </G>
  );

  return (
    <div className="of-slides-ribbon">
      <div className="of-slides-tabs" role="tablist">
        {tabs.map((x) => <button key={x.id} type="button" role="tab" aria-selected={cur === x.id} className={`of-slides-tab ${cur === x.id ? 'on' : ''} ${x.ctx ? 'ctx' : ''}`} onClick={() => setTab(x.id)}>{t(x.label)}</button>)}
        <span className="grow" />
        <RB icon={Undo2} title={t('Undo')} kbd="Ctrl+Z" disabled={ro || !state.canUndo} onClick={c.undo} />
        <RB icon={Redo2} title={t('Redo')} kbd="Ctrl+Y" disabled={ro || !state.canRedo} onClick={c.redo} />
        <RB icon={Presentation} label={t('Present')} title={t('Present from the current slide')} kbd="Shift+F5" onClick={() => c.present('current')} />
      </div>
      <div className="of-slides-rbody" role="tabpanel">
        {cur === 'home' && <>
          <G label={t('Slides')}>
            <Pop icon={Plus} label={t('New slide')} big split onMain={() => c.newSlide()} kbd="Ctrl+M" disabled={ro} width={500} content={(close) => <LayoutPicker doc={doc} close={close} onPick={(k) => c.newSlide(k)} />} />
            <div className="of-slides-rcol">
              <Pop icon={LayoutTemplate} label={t('Layout')} disabled={ro || !slide} width={500} content={(close) => <LayoutPicker doc={doc} close={close} onPick={(k) => c.setLayout(k)} />} />
              <RB icon={RotateCcw} label={t('Reset')} title={t('Reset slide')} disabled={ro || !slide} onClick={c.resetSlide} />
              <Pop icon={FolderPlus} label={t('Section')} disabled={ro || !slide} content={(close) => <Menu close={close} items={[{ label: t('Add section …'), onClick: c.addSection }, { label: t('Rename section …'), disabled: !c.curSection, onClick: c.renameSection }, { label: t('Remove section'), disabled: !c.curSection, onClick: c.removeSection }]} />} />
            </div>
          </G>
          <G label={t('Clipboard')}>
            <RB icon={ClipboardPaste} label={t('Paste')} big kbd="Ctrl+V" disabled={ro} onClick={c.paste} />
            <div className="of-slides-rcol">
              <RB icon={Scissors} title={t('Cut')} kbd="Ctrl+X" disabled={ro || !has} onClick={c.cut} />
              <RB icon={Copy} title={t('Copy')} kbd="Ctrl+C" disabled={!has} onClick={c.copy} />
              <RB icon={Paintbrush} title={t('Format painter')} active={!!state.painter} disabled={ro || !has} onClick={c.formatPainter} />
            </div>
          </G>
          {fontGroup}
          {paraGroup}
          {drawingGroup}
        </>}
        {cur === 'insert' && <>
          <G label={t('Slides')}><Pop icon={Plus} label={t('New slide')} big split onMain={() => c.newSlide()} disabled={ro} width={500} content={(close) => <LayoutPicker doc={doc} close={close} onPick={(k) => c.newSlide(k)} />} /></G>
          <G label={t('Tables')}><Pop icon={Table} label={t('Table')} big disabled={ro} content={(close) => <TablePicker close={close} onPick={c.insertTable} />} /></G>
          <G label={t('Images')}>
            <Pop icon={Image} label={t('Pictures')} big disabled={ro} content={(close) => <Menu close={close} items={[{ icon: ImagePlus, label: t('From this device …'), onClick: c.pictureFromFile }, { icon: Clipboard, label: t('From clipboard'), onClick: c.pictureFromClipboard }, { icon: Link2, label: t('From a web address …'), onClick: c.pictureFromUrl }]} />} />
          </G>
          <G label={t('Illustrations')}>
            <Pop icon={Shapes} label={t('Shapes')} big disabled={ro} width={300} content={(close) => <ShapeGallery close={close} onPick={(g) => c.setTool(g)} />} />
            <Pop icon={Smile} label={t('Icons')} big disabled={ro} width={300} content={(close) => <div className="of-slides-emoji">{EMOJI.map((e) => <button key={e} type="button" onClick={() => { close(); c.insertEmoji(e); }}>{e}</button>)}</div>} />
            <Pop icon={ChartColumn} label={t('Chart')} big disabled={ro} content={(close) => <Menu close={close} items={[{ icon: ChartColumn, label: t('Column'), onClick: () => c.insertChart('col') }, { icon: ChartBar, label: t('Bar'), onClick: () => c.insertChart('bar') }, { icon: ChartLine, label: t('Line'), onClick: () => c.insertChart('line') }, { icon: ChartLine, label: t('Area'), onClick: () => c.insertChart('area') }, { icon: ChartPie, label: t('Pie'), onClick: () => c.insertChart('pie') }, { icon: ChartPie, label: t('Doughnut'), onClick: () => c.insertChart('doughnut') }]} />} />
          </G>
          <G label={t('Links')}><RB icon={Link2} label={t('Link')} big kbd="Ctrl+K" disabled={ro || !has} onClick={c.insertLink} /></G>
          <G label={t('Comments')}><RB icon={MessageSquarePlus} label={t('Comment')} big disabled={ro || !slide} onClick={() => c.togglePane('comments', true)} /></G>
          <G label={t('Text')}>
            <RB icon={Type} label={t('Text box')} big active={state.tool === 'textbox'} disabled={ro} onClick={() => c.setTool('textbox')} />
            <div className="of-slides-rcol">
              <RB icon={PanelBottom} label={t('Header & footer')} disabled={ro} onClick={c.footerDialog} />
              <Pop icon={Omega} label={t('Symbol')} disabled={ro} width={300} content={(close) => <div className="of-slides-emoji sym">{SYMBOLS.map((e) => <button key={e} type="button" onClick={() => { close(); c.insertSymbol(e); }}>{e}</button>)}</div>} />
            </div>
          </G>
          <G label={t('Media')}>
            <RB icon={Film} label={t('Video')} big disabled={ro} onClick={() => c.insertMedia('video')} />
            <RB icon={Volume2} label={t('Audio')} big disabled={ro} onClick={() => c.insertMedia('audio')} />
          </G>
        </>}
        {cur === 'design' && <>
          <G label={t('Themes')}>
            <div className="of-slides-themestrip">
              {THEMES.slice(0, 5).map((th) => <ThemeChip key={th.id} th={th} doc={doc} on={doc.theme.id === th.id} onClick={() => c.applyTheme(th.id)} disabled={ro} />)}
              <Pop icon={Palette} title={t('All themes')} disabled={ro} width={500} content={(close) => <ThemeGallery doc={doc} close={close} onPick={c.applyTheme} />} />
            </div>
          </G>
          <G label={t('Variants')}>
            <div className="of-slides-rcol">
              <Pop icon={Palette} label={t('Colours')} disabled={ro} width={240} content={(close) => <Menu close={close} items={THEMES.map((th) => ({ label: th.name, onClick: () => c.themeColors(th.colors) }))} />} />
              <Pop icon={Type} label={t('Fonts')} disabled={ro} width={280} content={(close) => <Menu close={close} items={[['Segoe UI Semibold', 'Segoe UI'], ['Aptos Display', 'Aptos'], ['Calibri Light', 'Calibri'], ['Georgia', 'Segoe UI'], ['Cambria', 'Calibri'], ['Arial Black', 'Arial'], ['Trebuchet MS', 'Trebuchet MS'], ['Times New Roman', 'Arial'], ['Verdana', 'Verdana'], ['Garamond', 'Garamond']].map(([a, b]) => ({ label: `${a} / ${b}`, onClick: () => c.themeFonts({ major: a, minor: b }) }))} />} />
            </div>
          </G>
          <G label={t('Customize')}>
            <Pop icon={Monitor} label={t('Slide size')} big disabled={ro} content={(close) => <Menu close={close} items={[...['16:9', '4:3', '16:10', 'A4', 'Letter'].map((k) => ({ label: t(SIZES[k].label), active: doc.size.w === SIZES[k].w && doc.size.h === SIZES[k].h, onClick: () => c.slideSize(SIZES[k]) })), '-', { label: t('Custom slide size …'), onClick: c.customSize }]} />} />
            <Pop icon={PaintBucket} label={t('Format background')} big disabled={ro || !slide} width={270} content={(close) => <BackgroundPop c={c} close={close} />} />
          </G>
        </>}
        {cur === 'transitions' && <>
          <G label={t('Transition to this slide')}>
            <div className="of-slides-trgallery">{TRANSITIONS.map((x) => <button key={x.id} type="button" disabled={ro || !slide} className={`of-slides-trbtn ${tr.type === x.id ? 'on' : ''}`} onClick={() => c.setTransition({ type: x.id })}><TransitionIcon type={x.id} /><span className="small">{t(x.label)}</span></button>)}</div>
          </G>
          <G label={t('Effect options')}>
            <Pop icon={Sparkles} label={t('Direction')} disabled={ro || !['push', 'wipe', 'cover'].includes(tr.type)} content={(close) => <Menu close={close} items={DIRS.map((d) => ({ label: t(d.label), active: tr.dir === d.id, onClick: () => c.setTransition({ dir: d.id }) }))} />} />
            <RB icon={Play} label={t('Preview')} disabled={!slide || tr.type === 'none'} onClick={c.previewTransition} />
          </G>
          <G label={t('Timing')}>
            <div className="of-slides-rcol">
              <div className="row-tight small"><Timer size={14} />{t('Duration')}<Num label={t('Duration (seconds)')} unit="s" step={0.25} min={0.1} max={30} value={(tr.duration || 500) / 1000} disabled={ro || !slide} onChange={(v) => c.setTransition({ duration: Math.round(v * 1000) })} /></div>
              <label className="row-tight small"><input type="checkbox" disabled={ro || !slide} checked={tr.advanceAfter != null} onChange={(e) => c.setTransition({ advanceAfter: e.target.checked ? 5000 : null })} />{t('After')}<Num label={t('Advance after (seconds)')} unit="s" min={0} max={3600} value={tr.advanceAfter != null ? tr.advanceAfter / 1000 : 0} disabled={ro || tr.advanceAfter == null} onChange={(v) => c.setTransition({ advanceAfter: Math.round(v * 1000) })} /></label>
            </div>
            <RB icon={Layers} label={t('Apply to all')} disabled={ro || !slide} onClick={c.transitionToAll} />
          </G>
        </>}
        {cur === 'animations' && <>
          <G label={t('Animation')}>
            <div className="of-slides-trgallery">
              <button type="button" className={`of-slides-trbtn ${one && !selAnim ? 'on' : ''}`} disabled={ro || !one} onClick={c.removeAnim}><TransitionIcon type="none" /><span className="small">{t('None')}</span></button>
              {EFFECTS.map((x) => <button key={x.id} type="button" disabled={ro || !has} className={`of-slides-trbtn ${selAnim && selAnim.effect === x.id ? 'on' : ''}`} onClick={() => c.setAnim({ effect: x.id })}><TransitionIcon type={x.id === 'appear' ? 'appear' : x.id === 'fly' ? 'push' : x.id} /><span className="small">{t(x.label)}</span></button>)}
            </div>
          </G>
          <G label={t('Effect options')}>
            <Pop icon={Sparkles} label={t('Direction')} disabled={ro || !selAnim || selAnim.effect !== 'fly'} content={(close) => <Menu close={close} items={FLY_DIRS.map((d) => ({ label: t(d.label), active: selAnim && selAnim.dir === d.id, onClick: () => c.setAnim({ dir: d.id }) }))} />} />
            <RB icon={Play} label={t('Preview')} disabled={!slide || !(slide.animations || []).length} onClick={c.previewAnims} />
          </G>
          <G label={t('Timing')}>
            <div className="of-slides-rcol">
              <select className="input" aria-label={t('Start')} disabled={ro || !selAnim} value={selAnim ? selAnim.trigger : 'click'} onChange={(e) => c.setAnim({ trigger: e.target.value })}>
                <option value="click">{t('On click')}</option><option value="with">{t('With previous')}</option><option value="after">{t('After previous')}</option>
              </select>
              <div className="row-tight small">{t('Duration')}<Num label={t('Duration (seconds)')} unit="s" step={0.25} min={0.05} max={30} disabled={ro || !selAnim} value={selAnim ? (selAnim.duration || 500) / 1000 : 0.5} onChange={(v) => c.setAnim({ duration: Math.round(v * 1000) })} /></div>
              <div className="row-tight small">{t('Delay')}<Num label={t('Delay (seconds)')} unit="s" step={0.25} min={0} max={60} disabled={ro || !selAnim} value={selAnim ? (selAnim.delay || 0) / 1000 : 0} onChange={(v) => c.setAnim({ delay: Math.round(v * 1000) })} /></div>
            </div>
            <div className="of-slides-rcol">
              <RB icon={ArrowUp} label={t('Move earlier')} disabled={ro || !selAnim} onClick={() => c.moveAnim(-1)} />
              <RB icon={ArrowDown} label={t('Move later')} disabled={ro || !selAnim} onClick={() => c.moveAnim(1)} />
            </div>
          </G>
          <G label={t('Advanced')}><RB icon={ListTree} label={t('Animation pane')} big active={state.pane === 'anim'} onClick={() => c.togglePane('anim')} /></G>
        </>}
        {cur === 'show' && <>
          <G label={t('Start slide show')}>
            <RB icon={Play} label={t('From beginning')} big kbd="F5" disabled={!doc.slides.length} onClick={() => c.present('start')} />
            <RB icon={MonitorPlay} label={t('From current slide')} big kbd="Shift+F5" disabled={!slide} onClick={() => c.present('current')} />
            <RB icon={Monitor} label={t('Presenter view')} big disabled={!slide} onClick={() => c.present('presenter')} />
          </G>
          <G label={t('Set up')}>
            <div className="of-slides-rcol">
              <RB icon={EyeOff} label={t('Hide slide')} active={slide && slide.hidden} disabled={ro || !slide} onClick={c.toggleHidden} />
              <RB icon={Timer} label={t('Rehearse timings')} disabled={ro || !doc.slides.length} onClick={() => c.present('rehearse')} />
            </div>
            <div className="of-slides-rcol">
              <label className="row-tight small"><input type="checkbox" disabled={ro} checked={!!(doc.show && doc.show.loop)} onChange={(e) => c.setShow({ loop: e.target.checked })} /><Repeat size={14} />{t('Loop until Esc')}</label>
              <label className="row-tight small"><input type="checkbox" disabled={ro} checked={!(doc.show && doc.show.useTimings === false)} onChange={(e) => c.setShow({ useTimings: e.target.checked })} />{t('Use timings')}</label>
              <label className="row-tight small"><input type="checkbox" checked={!!state.presenterView} onChange={(e) => c.setPresenterView(e.target.checked)} />{t('Use presenter view')}</label>
            </div>
          </G>
        </>}
        {cur === 'review' && <>
          <G label={t('Proofing')}><RB icon={SpellCheck} label={t('Spelling')} big active={state.spellcheck} onClick={c.toggleSpell} /></G>
          <G label={t('Comments')}>
            <RB icon={MessageSquarePlus} label={t('New comment')} big disabled={ro || !slide} onClick={() => c.togglePane('comments', true)} />
            <RB icon={MessageSquare} label={t('Show comments')} big active={state.pane === 'comments'} onClick={() => c.togglePane('comments')} />
          </G>
        </>}
        {cur === 'view' && <>
          <G label={t('Presentation views')}>
            <RB icon={PanelLeft} label={t('Normal')} big active={state.view === 'normal'} onClick={() => c.setView('normal')} />
            <RB icon={LayoutList} label={t('Outline view')} big active={state.view === 'outline'} onClick={() => c.setView('outline')} />
            <RB icon={Grid2x2} label={t('Slide sorter')} big active={state.view === 'sorter'} onClick={() => c.setView('sorter')} />
            <RB icon={StickyNote} label={t('Notes page')} big active={state.view === 'notes'} onClick={() => c.setView('notes')} />
          </G>
          <G label={t('Show')}>
            <div className="of-slides-rcol">
              <label className="row-tight small"><input type="checkbox" checked={state.guides} onChange={c.toggleGuides} /><Ruler size={14} />{t('Guides')}</label>
              <label className="row-tight small"><input type="checkbox" checked={state.grid} onChange={c.toggleGrid} /><Grid3x3 size={14} />{t('Gridlines')}</label>
              <label className="row-tight small"><input type="checkbox" checked={state.snap} onChange={c.toggleSnap} /><Magnet size={14} />{t('Snap to objects')}</label>
            </div>
            <div className="of-slides-rcol">
              <label className="row-tight small"><input type="checkbox" checked={state.notesPane} onChange={c.toggleNotes} /><StickyNote size={14} />{t('Notes')}</label>
              <label className="row-tight small"><input type="checkbox" checked={state.thumbs} onChange={c.toggleThumbs} /><PanelLeft size={14} />{t('Slide panel')}</label>
            </div>
          </G>
          <G label={t('Zoom')}>
            <RB icon={ZoomOut} title={t('Zoom out')} onClick={() => c.zoomBy(-1)} />
            <RB icon={Maximize} label={t('Fit to window')} active={state.fit} onClick={c.zoomFit} />
            <RB icon={ZoomIn} title={t('Zoom in')} onClick={() => c.zoomBy(1)} />
          </G>
        </>}
        {cur === 'format' && one && one.type === 'table' && <>
          <G label={t('Rows & columns')}>
            <div className="of-slides-rcol"><RB icon={BetweenHorizontalStart} label={t('Insert above')} disabled={ro} onClick={() => c.tableOp('rowAbove')} /><RB icon={BetweenHorizontalEnd} label={t('Insert below')} disabled={ro} onClick={() => c.tableOp('rowBelow')} /></div>
            <div className="of-slides-rcol"><RB icon={BetweenVerticalStart} label={t('Insert left')} disabled={ro} onClick={() => c.tableOp('colLeft')} /><RB icon={BetweenVerticalEnd} label={t('Insert right')} disabled={ro} onClick={() => c.tableOp('colRight')} /></div>
            <div className="of-slides-rcol"><RB icon={Trash2} label={t('Delete row')} disabled={ro} onClick={() => c.tableOp('delRow')} /><RB icon={Trash2} label={t('Delete column')} disabled={ro} onClick={() => c.tableOp('delCol')} /></div>
          </G>
          <G label={t('Merge')}>
            <div className="of-slides-rcol"><RB icon={TableCellsMerge} label={t('Merge right')} disabled={ro} onClick={() => c.tableOp('mergeRight')} /><RB icon={TableCellsMerge} label={t('Merge down')} disabled={ro} onClick={() => c.tableOp('mergeDown')} /></div>
            <RB icon={TableCellsSplit} label={t('Split cells')} big disabled={ro} onClick={() => c.tableOp('split')} />
          </G>
          <G label={t('Table style options')}>
            <div className="of-slides-rcol">
              <label className="row-tight small"><input type="checkbox" checked={!!(one.style && one.style.firstRow)} onChange={(e) => c.tableStyle({ firstRow: e.target.checked })} />{t('Header row')}</label>
              <label className="row-tight small"><input type="checkbox" checked={!!(one.style && one.style.bandRows)} onChange={(e) => c.tableStyle({ bandRows: e.target.checked })} />{t('Banded rows')}</label>
              <label className="row-tight small"><input type="checkbox" checked={!!(one.style && one.style.firstCol)} onChange={(e) => c.tableStyle({ firstCol: e.target.checked })} />{t('First column')}</label>
            </div>
            <Pop icon={Palette} label={t('Table colour')} disabled={ro} content={(close) => <Menu close={close} items={['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'tx1'].map((a) => ({ label: a, active: one.style && one.style.accent === a, onClick: () => c.tableStyle({ accent: a }) }))} />} />
            <ColorPop icon={PaintBucket} title={t('Cell shading')} theme={theme} value={null} noneLabel="No fill" disabled={ro || !state.cell} onPick={(col) => c.tableOp('fill', col)} />
          </G>
        </>}
        {cur === 'format' && one && one.type === 'chart' && <>
          <G label={t('Chart type')}>
            <select className="input" aria-label={t('Chart type')} disabled={ro} value={one.chart.kind} onChange={(e) => c.chartSet({ kind: e.target.value })}>
              {[['col', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie'], ['doughnut', 'Doughnut']].map(([k, l]) => <option key={k} value={k}>{t(l)}</option>)}
            </select>
          </G>
          <G label={t('Data')}><RB icon={Table} label={t('Edit data …')} big disabled={ro} onClick={c.chartData} /></G>
          <G label={t('Chart elements')}>
            <div className="of-slides-rcol">
              <label className="row-tight small"><input type="checkbox" checked={one.chart.legend !== false} onChange={(e) => c.chartSet({ legend: e.target.checked })} />{t('Legend')}</label>
              <label className="row-tight small"><input type="checkbox" checked={!!one.chart.labels} onChange={(e) => c.chartSet({ labels: e.target.checked })} />{t('Data labels')}</label>
              <label className="row-tight small"><input type="checkbox" checked={!!one.chart.stacked} onChange={(e) => c.chartSet({ stacked: e.target.checked })} />{t('Stacked')}</label>
            </div>
            <input className="input" placeholder={t('Chart title')} aria-label={t('Chart title')} value={one.chart.title || ''} onKeyDown={(e) => e.stopPropagation()} onChange={(e) => c.chartSet({ title: e.target.value }, `charttitle:${one.id}`)} />
          </G>
        </>}
        {cur === 'format' && one && one.type === 'image' && <>
          <G label={t('Adjust')}>
            <RB icon={ImagePlus} label={t('Change picture …')} big disabled={ro} onClick={c.replacePicture} />
            <RB icon={RotateCcw} label={t('Reset picture')} big disabled={ro} onClick={c.resetPicture} />
          </G>
          <G label={t('Picture styles')}>
            <ColorPop icon={PenLine} title={t('Picture border')} theme={theme} value={one.line && one.line.color} noneLabel="No outline" disabled={ro} onPick={c.shapeLine} />
            <RB icon={Square} title={t('Shadow')} active={!!one.shadow} disabled={ro} onClick={c.toggleShadow} />
            <Pop icon={Shapes} label={t('Crop to shape')} disabled={ro} width={300} content={(close) => <ShapeGallery close={close} onPick={(g) => c.cropToShape(g)} />} />
          </G>
          <G label={t('Size')}><RB icon={PanelRight} label={t('Crop …')} big onClick={() => c.togglePane('format', true)} /></G>
          {fontGroup && null}
        </>}
        {cur === 'format' && ctxTab === 'shape' && <>
          {drawingGroup}
          <G label={t('Size')}><RB icon={PanelRight} label={t('Format pane')} big active={state.pane === 'format'} onClick={() => c.togglePane('format')} /></G>
        </>}
      </div>
    </div>
  );
}

function ThemeChip({ th, doc, on, onClick, disabled }) {
  const td = newDoc({ size: doc.size, theme: th.id });
  const s = newSlide(td, 'title', { ctrTitle: 'Aa', subTitle: th.name });
  return <Raw as="button" type="button" disabled={disabled} className={`of-slides-themechip ${on ? 'on' : ''}`} title={th.name} aria-label={th.name} onMouseDown={(e) => e.preventDefault()} onClick={onClick}><SlideView doc={td} slide={s} scale={88 / S(doc.size.w)} mode="thumb" /></Raw>;
}

function TransitionIcon({ type }) {
  const p = { width: 26, height: 20, viewBox: '0 0 26 20', 'aria-hidden': true };
  const box = <rect x="1" y="1" width="24" height="18" rx="2" fill="none" stroke="currentColor" strokeWidth="1.3" />;
  if (type === 'none') return <svg {...p}>{box}<path d="M5 15L21 5" stroke="currentColor" strokeWidth="1.3" /></svg>;
  if (type === 'fade' || type === 'appear') return <svg {...p}>{box}<rect x="4" y="4" width="18" height="12" fill="currentColor" opacity={type === 'appear' ? 0.9 : 0.35} /></svg>;
  if (type === 'push') return <svg {...p}>{box}<path d="M7 10h11M14 6l4 4-4 4" fill="none" stroke="currentColor" strokeWidth="1.5" /></svg>;
  if (type === 'wipe') return <svg {...p}>{box}<rect x="4" y="4" width="9" height="12" fill="currentColor" opacity=".5" /><path d="M13 3v14" stroke="currentColor" strokeWidth="1.5" /></svg>;
  if (type === 'cover') return <svg {...p}>{box}<rect x="9" y="4" width="13" height="12" fill="currentColor" opacity=".5" /><path d="M9 4v12" stroke="currentColor" /></svg>;
  if (type === 'split') return <svg {...p}>{box}<path d="M13 3v14M9 10H4M17 10h5" stroke="currentColor" strokeWidth="1.5" /></svg>;
  return <svg {...p}>{box}<rect x="9" y="7" width="8" height="6" fill="currentColor" opacity=".5" /><rect x="5" y="4" width="16" height="12" fill="none" stroke="currentColor" strokeDasharray="2 2" /></svg>;
}

function FontSize({ value, onChange, disabled }) {
  const [v, setV] = useState('');
  const shown = v !== '' ? v : value ? String(value) : '';
  return (
    <span className="of-slides-sizebox">
      <input className="input" list="of-slides-sizes" aria-label={t('Font size')} title={t('Font size')} disabled={disabled} value={shown}
        onChange={(e) => setV(e.target.value)} onBlur={() => { const n = parseFloat(String(v).replace(',', '.')); if (v !== '' && n > 0 && n <= 4000) onChange(n); setV(''); }}
        onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter') { const n = parseFloat(String(v).replace(',', '.')); if (n > 0 && n <= 4000) onChange(n); setV(''); e.target.blur(); } }} />
      <datalist id="of-slides-sizes">{SIZES_PT.map((s) => <option key={s} value={s} />)}</datalist>
    </span>
  );
}

function BackgroundPop({ c, close }) {
  const { doc, slide } = c;
  const bg = slide.background;
  const [kind, setKind] = useState(bg ? bg.type : 'solid');
  return (
    <div className="stack-tight of-slides-bgpop">
      <div className="seg-like row-tight">
        {[['solid', 'Solid'], ['grad', 'Gradient'], ['image', 'Picture']].map(([k, l]) => <button key={k} type="button" className={`btn xs ${kind === k ? 'primary' : ''}`} onClick={() => setKind(k)}>{t(l)}</button>)}
      </div>
      {kind === 'solid' && <ColorGrid theme={doc.theme} value={bg && bg.color} allowNone={false} onPick={(col) => c.setBackground({ type: 'solid', color: col })} />}
      {kind === 'grad' && (
        <div className="stack-tight">
          {[['accent1', 'accent2'], ['accent1', 'accent1-50'], ['bg1', 'bg2'], ['#0F172A', '#1E293B'], ['accent3', 'accent5'], ['#FFB347', '#F25F5C'], ['tx1', 'tx2'], ['#FFFFFF', 'accent1+80']].map(([a, b], i) => (
            <button key={i} type="button" className="of-slides-gradbtn" style={{ background: `linear-gradient(135deg, ${css(a, doc.theme)}, ${css(b, doc.theme)})` }} aria-label={t('Gradient {n}', { n: i + 1 })} onClick={() => c.setBackground({ type: 'grad', angle: 45, stops: [{ pos: 0, color: a }, { pos: 100, color: b }] })} />
          ))}
        </div>
      )}
      {kind === 'image' && <button type="button" className="btn sm" onClick={() => { close(); c.backgroundPicture(); }}><ImagePlus size={14} />{t('Choose picture …')}</button>}
      <label className="row-tight small"><input type="checkbox" checked={slide.showMaster === false} onChange={(e) => c.hideBackgroundGraphics(e.target.checked)} />{t('Hide background graphics')}</label>
      <div className="row-tight">
        <button type="button" className="btn xs" onClick={() => { close(); c.backgroundToAll(); }}>{t('Apply to all')}</button>
        <button type="button" className="btn xs ghost" onClick={() => { close(); c.setBackground(null); }}>{t('Reset background')}</button>
      </div>
    </div>
  );
}

export { themeData, buildLayouts, LAYOUT_KINDS };
