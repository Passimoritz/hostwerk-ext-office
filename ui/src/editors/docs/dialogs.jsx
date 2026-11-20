/*
 * Dialogs of the docs editor (Hostwerk Modal / SideSheet): symbols, emoji, link, header & footer, page setup,
 * paragraph spacing, word count, equation, macros, table grid picker, date & time.
 */
import { useMemo, useState } from 'react';
import * as HUI from '@hostwerk/ui';
import { Btn, Modal, Field, Seg, Switch, Callout } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Raw } from '../_shared/keep.jsx';
import { equationHtml } from './io/equation.js';

const Sheet = HUI.SideSheet || Modal;

const SYMBOLS = {
  Common: '©®™§¶†‡•·…–—‰′″€£¥¢¤°±×÷≈≠≤≥∞√∑∏∫∂∆∇∈∉∩∪⊂⊃∀∃¬∧∨⇒⇔←→↑↓↔↵✓✗★☆♥♦♣♠☐☑☒○●□■▲▼◆◇',
  Quotes: '„“”‚‘’«»‹›"\'',
  Latin: 'ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿŒœŠšŽžŸĆćČčĐđĘęŁłŃńŚśŹźŻżĞğİıŞş',
  Greek: 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψω',
  Maths: '½⅓⅔¼¾⅛⅜⅝⅞¹²³⁴⁵⁶⁷⁸⁹⁰₀₁₂₃₄₅₆₇₈₉∠⊥∥∝∼≡≅⌀∅ℕℤℚℝℂ',
  Arrows: '←↑→↓↔↕↖↗↘↙⇐⇑⇒⇓⇔➔➜➤⟵⟶⟷',
  Currency: '€$£¥¢₹₽₩₺₿₴₦₱₫',
};
const EMOJI = '😀😃😄😁😆😅😂🙂😉😊😇🥰😍🤩😘😋😎🤓🧐🤔🤨😐😑😶🙄😏😴😌😛🤗🤭🤫😬😮😯😲😳🥺😢😭😤😡🤯😱👍👎👌✌️🤞👏🙌🙏💪👋✋☝️👉👈👆👇🤝❤️🧡💛💚💙💜🖤🤍💯✅❌⚠️❗❓💡🔥⭐🌟✨🎉🎊🎁🎂🏆🥇📌📎📝📄📅📆📊📈📉🗂️📁🔒🔓🔑⏰⌛☕🍕🍰🌍🌞🌧️❄️🌈🚀✈️🚗🏠💼🖥️💻📱☎️📧✉️🔔🎵📷🎬🐶🐱🌻🍀';

export function SymbolDialog({ onInsert, onClose }) {
  const [cat, setCat] = useState('Common');
  const [code, setCode] = useState('');
  const [recent, setRecent] = useState(() => { try { return JSON.parse(localStorage.getItem('of-docs-recent-symbols') || '[]'); } catch { return []; } });
  const pick = (ch) => {
    onInsert(ch);
    const r = [ch, ...recent.filter((x) => x !== ch)].slice(0, 20);
    setRecent(r);
    try { localStorage.setItem('of-docs-recent-symbols', JSON.stringify(r)); } catch { /* fine */ }
  };
  const fromCode = () => { const n = parseInt(code.replace(/^u\+|^0x/i, ''), 16); if (n > 0 && n < 0x110000) pick(String.fromCodePoint(n)); };
  return (
    <Modal title="Insert symbol" onClose={onClose} footer={<Btn onClick={onClose}>Close</Btn>}>
      <div className="stack">
        <Seg value={cat} onChange={setCat} options={Object.keys(SYMBOLS).map((k) => ({ id: k, label: k }))} />
        <div className="of-docs-symbols">{[...SYMBOLS[cat]].map((ch, k) => <button type="button" key={`${ch}${k}`} className="of-docs-sym" title={`U+${ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`} onClick={() => pick(ch)}>{ch}</button>)}</div>
        {recent.length > 0 && <Field label="Recently used"><div className="of-docs-symbols">{recent.map((ch) => <button type="button" key={ch} className="of-docs-sym" onClick={() => pick(ch)}>{ch}</button>)}</div></Field>}
        <Field label="Unicode code point" hint="For example 2603 for ☃"><div className="row-tight"><input className="input mono" value={code} onChange={(e) => setCode(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && fromCode()} placeholder="U+2603" aria-label="Unicode code point" /><Btn className="sm" onClick={fromCode}>Insert</Btn></div></Field>
      </div>
    </Modal>
  );
}

export function EmojiGrid({ onPick }) {
  return <div className="of-docs-emoji">{[...new Intl.Segmenter().segment(EMOJI)].map((s) => s.segment).map((e, k) => <button type="button" key={`${e}${k}`} className="of-docs-sym" onClick={() => onPick(e)}>{e}</button>)}</div>;
}

export function LinkDialog({ initial, onSave, onRemove, onClose }) {
  const [href, setHref] = useState(initial.href || '');
  const [text, setText] = useState(initial.text || '');
  const valid = /^(https?:|mailto:|tel:|#|ftp:|file:)/i.test(href.trim()) || /^[\w.-]+\.[a-z]{2,}(\/.*)?$/i.test(href.trim());
  const save = () => { let h = href.trim(); if (!/^[a-z]+:|^#/i.test(h)) h = /@/.test(h) ? `mailto:${h}` : `https://${h}`; onSave(h, text); };
  return (
    <Modal title={initial.href ? 'Edit link' : 'Insert link'} onClose={onClose} footer={<>{initial.href && <Btn className="ghost" onClick={onRemove}>Remove link</Btn>}<Btn onClick={onClose}>Cancel</Btn><Btn className="primary" disabled={!valid} onClick={save}>Save</Btn></>}>
      <div className="stack">
        <Field label="Address" hint="Web address, e-mail address or #heading. Ctrl+click opens a link in the document."><input className="input" autoFocus value={href} onChange={(e) => setHref(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && valid && save()} placeholder="https://example.org" aria-label="Address" /></Field>
        <Field label="Text to show"><input className="input" value={text} onChange={(e) => setText(e.target.value)} placeholder={t('Leave empty to use the address')} aria-label="Text to show" /></Field>
      </div>
    </Modal>
  );
}

export function HeaderFooterDialog({ header, footer, titlePage, onSave, onClose }) {
  const [h, setH] = useState({ ...header });
  const [f, setF] = useState({ ...footer });
  const [first, setFirst] = useState(!!titlePage);
  const alignOpts = [{ id: 'left', label: 'Left' }, { id: 'center', label: 'Centre' }, { id: 'right', label: 'Right' }];
  const tokens = (set, v) => (
    <div className="row-tight of-docs-tokens">
      <Btn className="xs ghost" onClick={() => set({ ...v, text: `${v.text || ''}{page}` })}>Page number</Btn>
      <Btn className="xs ghost" onClick={() => set({ ...v, text: `${v.text || ''}{pages}` })}>Number of pages</Btn>
      <Btn className="xs ghost" onClick={() => set({ ...v, text: `${v.text || ''}{date}` })}>Date</Btn>
      <Btn className="xs ghost" onClick={() => set({ ...v, text: `${v.text || ''}{title}` })}>Title</Btn>
      <Btn className="xs ghost" onClick={() => set({ ...v, text: t('Page {page} of {pages}').replace('{page}', '{page}') })}>“Page X of Y”</Btn>
    </div>
  );
  return (
    <Modal title="Header and footer" subtitle={t('{page} becomes the page number, {pages} the number of pages.')} onClose={onClose} footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" onClick={() => onSave(h, f, first)}>Save</Btn></>}>
      <div className="stack">
        <Field label="Header"><input className="input" value={h.text || ''} onChange={(e) => setH({ ...h, text: e.target.value })} placeholder={t('e.g. company name')} aria-label="Header" /></Field>
        <Seg value={h.align || 'center'} onChange={(v) => setH({ ...h, align: v })} options={alignOpts} />
        {tokens(setH, h)}
        <Field label="Footer"><input className="input" value={f.text || ''} onChange={(e) => setF({ ...f, text: e.target.value })} placeholder={t('e.g. Page {page} of {pages}')} aria-label="Footer" /></Field>
        <Seg value={f.align || 'center'} onChange={(v) => setF({ ...f, align: v })} options={alignOpts} />
        {tokens(setF, f)}
        <label className="check"><Switch checked={first} onChange={setFirst} /> {t('Different first page (no header/footer on the title page)')}</label>
      </div>
    </Modal>
  );
}

export function MarginsDialog({ page, units, onSave, onClose }) {
  const inch = units === 'in';
  const conv = (mm) => (inch ? Math.round((mm / 25.4) * 100) / 100 : Math.round(mm / 10 * 100) / 100);
  const back = (v) => (inch ? Number(v) * 25.4 : Number(v) * 10);
  const [m, setM] = useState({ top: conv(page.margins.top), right: conv(page.margins.right), bottom: conv(page.margins.bottom), left: conv(page.margins.left) });
  const [w, setW] = useState(conv(page.width)); const [h, setH] = useState(conv(page.height));
  const u = inch ? 'in' : 'cm';
  const num = (k) => <Field label={`${t({ top: 'Top', right: 'Right', bottom: 'Bottom', left: 'Left' }[k])} (${u})`}><input className="input" inputMode="decimal" value={m[k]} onChange={(e) => setM({ ...m, [k]: e.target.value.replace(',', '.') })} aria-label={k} /></Field>;
  const ok = Object.values(m).every((v) => Number(v) >= 0 && Number(v) < (inch ? 10 : 25)) && Number(w) > 2 && Number(h) > 2;
  return (
    <Modal title="Page setup" onClose={onClose} footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" disabled={!ok} onClick={() => onSave({ margins: { top: back(m.top), right: back(m.right), bottom: back(m.bottom), left: back(m.left) }, width: back(w), height: back(h) })}>Save</Btn></>}>
      <div className="stack">
        <div className="grid grid-2">{num('top')}{num('bottom')}{num('left')}{num('right')}</div>
        <div className="grid grid-2">
          <Field label={`${t('Paper width')} (${u})`}><input className="input" inputMode="decimal" value={w} onChange={(e) => setW(e.target.value.replace(',', '.'))} aria-label="Paper width" /></Field>
          <Field label={`${t('Paper height')} (${u})`}><input className="input" inputMode="decimal" value={h} onChange={(e) => setH(e.target.value.replace(',', '.'))} aria-label="Paper height" /></Field>
        </div>
      </div>
    </Modal>
  );
}

export function ParagraphDialog({ attrs, onSave, onClose }) {
  const [a, setA] = useState({ spaceBefore: attrs.spaceBefore ?? '', spaceAfter: attrs.spaceAfter ?? '', lineHeight: attrs.lineHeight ?? '', indent: attrs.indent ?? '', firstLine: attrs.firstLine ?? '' });
  const f = (k, label, hint) => <Field label={label} hint={hint}><input className="input" inputMode="decimal" value={a[k]} onChange={(e) => setA({ ...a, [k]: e.target.value.replace(',', '.') })} aria-label={label} /></Field>;
  const num = (v) => (v === '' || v == null ? null : Number(v));
  return (
    <Modal title="Paragraph" onClose={onClose} footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" onClick={() => onSave({ spaceBefore: num(a.spaceBefore), spaceAfter: num(a.spaceAfter), lineHeight: num(a.lineHeight), indent: num(a.indent), firstLine: num(a.firstLine) })}>Save</Btn></>}>
      <div className="grid grid-2">
        {f('spaceBefore', 'Spacing before (pt)')}{f('spaceAfter', 'Spacing after (pt)')}
        {f('lineHeight', 'Line spacing (multiple)', 'e.g. 1, 1.15, 1.5, 2')}{f('indent', 'Indent left (pt)')}
        {f('firstLine', 'First line (pt)', 'Negative = hanging indent')}
      </div>
    </Modal>
  );
}

export function WordCountDialog({ stats, selection, pages, onClose, goal = 0, onGoal }) {
  const [g, setG] = useState(goal ? String(goal) : '');
  const rows = [['Pages', pages], ['Words', stats.words], ['Characters (no spaces)', stats.charsNoSpaces], ['Characters (with spaces)', stats.chars], ['Paragraphs', stats.paragraphs]];
  return (
    <Modal title="Word count" onClose={onClose} footer={<Btn className="primary" onClick={onClose}>Close</Btn>}>
      <table className="table"><tbody>{rows.map(([k, v]) => <tr key={k}><td>{k}</td><td className="num">{Number(v).toLocaleString()}</td></tr>)}</tbody></table>
      {selection && <p className="small muted">{t('Selection: {w} words, {c} characters', { w: selection.words, c: selection.chars })}</p>}
      {onGoal && <Field label="Word goal" hint="The status bar shows how far you are. Empty = no goal."><div className="row-tight"><input className="input" inputMode="numeric" value={g} onChange={(e) => setG(e.target.value.replace(/D/g, ''))} placeholder={t('e.g. 1500')} aria-label="Word goal" /><Btn onClick={() => { onGoal(Number(g) || 0); onClose(); }}>Set goal</Btn></div></Field>}
    </Modal>
  );
}

export function EquationDialog({ latex, onSave, onClose }) {
  const [v, setV] = useState(latex || '');
  const html = useMemo(() => equationHtml(v), [v]);
  const samples = ['x^2 + y^2 = r^2', '\\frac{a}{b}', '\\sqrt{x}', 'a_1 + a_2', '\\sum_{i=1}^{n} i', 'E = mc^2', '\\alpha \\cdot \\beta', 'x \\leq y'];
  return (
    <Modal title="Equation" onClose={onClose} footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" disabled={!v.trim()} onClick={() => onSave(v.trim())}>{latex ? 'Save' : 'Insert'}</Btn></>}>
      <div className="stack">
        <Field label="Formula" hint="Simple LaTeX: ^ superscript, _ subscript, \\frac{a}{b}, \\sqrt{x}, Greek letters like \\alpha"><input className="input mono" autoFocus value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && v.trim() && onSave(v.trim())} aria-label="Formula" /></Field>
        <div className="of-docs-eq-preview of-docs-paper"><span className="of-docs-eq" dangerouslySetInnerHTML={{ __html: html || '&nbsp;' }} /></div>
        <div className="row-tight of-docs-tokens">{samples.map((s) => <Btn key={s} className="xs ghost mono" onClick={() => setV(s)}>{s}</Btn>)}</div>
      </div>
    </Modal>
  );
}

export function MacrosSheet({ modules, error, onClose }) {
  const [open, setOpen] = useState(modules && modules[0] ? modules[0].name : null);
  const cur = (modules || []).find((m) => m.name === open);
  return (
    <Sheet title="Macros" subtitle={t('Kept in the file and never run here – read only.')} onClose={onClose} size="lg">
      <div className="stack">
        <Callout kind="info" title="Macros stay in the file">{t('When you save, the macros are written back unchanged. To run or edit them, open the file in the desktop app.')}</Callout>
        {error && <Callout kind="warn">{error}</Callout>}
        {modules && modules.length > 0 ? (
          <>
            <Raw as={Seg} value={open} onChange={setOpen} options={modules.map((m) => ({ id: m.name, label: m.name }))} />
            {cur && <pre className="of-docs-vba mono">{cur.source || t('(no source code – compiled only)')}</pre>}
          </>
        ) : !error && <p className="muted">{t('No macro modules found.')}</p>}
      </div>
    </Sheet>
  );
}

export function TableGrid({ onPick, onCustom }) {
  const [hover, setHover] = useState([0, 0]);
  return (
    <div className="of-docs-tgrid-wrap">
      <div className="of-docs-tgrid" onMouseLeave={() => setHover([0, 0])}>
        {Array.from({ length: 8 }, (_, r) => Array.from({ length: 10 }, (__, c) => (
          <button type="button" key={`${r}-${c}`} aria-label={`${r + 1} × ${c + 1}`} className={`of-tcell ${r < hover[0] && c < hover[1] ? 'of-on' : ''}`} onMouseEnter={() => setHover([r + 1, c + 1])} onClick={() => onPick(r + 1, c + 1)} />
        )))}
      </div>
      <div className="small muted of-tgrid-label">{hover[0] ? t('{r} × {c} table', { r: hover[0], c: hover[1] }) : t('Insert table')}</div>
      <button type="button" className="of-docs-mi" onClick={onCustom}>{t('Custom size …')}</button>
    </div>
  );
}

export function dateFormats(lang) {
  const loc = lang === 'de' ? 'de-DE' : 'en-GB';
  const d = new Date();
  return [
    d.toLocaleDateString(loc),
    d.toLocaleDateString(loc, { day: 'numeric', month: 'long', year: 'numeric' }),
    d.toLocaleDateString(loc, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }),
    d.toISOString().slice(0, 10),
    d.toLocaleTimeString(loc, { hour: '2-digit', minute: '2-digit' }),
    d.toLocaleString(loc, { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
  ];
}
