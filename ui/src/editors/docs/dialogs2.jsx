/*
 * 1.1 dialogs of the docs editor: formula editor (LaTeX with live MathML preview and a palette), live fields,
 * diagrams from text (Mermaid, live preview), AI (write / rewrite / summarise / translate / continue), headers and
 * footers with rich content (first / even pages), page setup of one section, compare two versions (also the chunk's
 * `Compare` export for the shell's compare view).
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import Underline from '@tiptap/extension-underline';
import TextAlign from '@tiptap/extension-text-align';
import TextStyle from '@tiptap/extension-text-style';
import Color from '@tiptap/extension-color';
import FontFamily from '@tiptap/extension-font-family';
import Table from '@tiptap/extension-table';
import TableRow from '@tiptap/extension-table-row';
import TableCell from '@tiptap/extension-table-cell';
import TableHeader from '@tiptap/extension-table-header';
import { Bold, Italic, Underline as UIcon, AlignLeft, AlignCenter, AlignRight, Hash, Sigma, CalendarDays, Image as ImageIcon, Table as TableIcon, FileText, Sparkles, Columns2, Rows3, Check, Copy, Replace, Plus, GitCompare, ArrowRight, Wand2 } from 'lucide-react';
import { Btn, Modal, Field, Seg, Switch, Callout, Spinner, toast } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Keep, Raw } from '../_shared/keep.jsx';
import { renderLatex, renderMermaid, svgToPng, compareLib } from './lazy.js';
import { FontSize, ParagraphFormat } from './extensions/format.js';
import { ImageEx } from './extensions/nodes.js';
import { Field as FieldNode, FIELD_LABELS, formatDate } from './extensions/rich.js';
import { equationHtml } from './io/equation.js';
import { readDocument } from './io/index.js';
import { PAGE_SIZES, inlineText, walk } from './io/pm.js';
import { docToHtml } from './io/render-html.js';
import { mdToHtml } from './io/md.js';

/* ---------------- formulas ---------------- */
const PALETTE = [
  ['Structures', [['\\frac{a}{b}', 'a/b'], ['x^{2}', 'xⁿ'], ['x_{i}', 'xᵢ'], ['\\sqrt{x}', '√x'], ['\\sqrt[n]{x}', 'ⁿ√x'], ['\\sum_{i=1}^{n}', '∑'], ['\\int_{a}^{b}', '∫'], ['\\prod_{i=1}^{n}', '∏'], ['\\lim_{x \\to \\infty}', 'lim'], ['\\left( x \\right)', '( )'], ['\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}', '( ⁂ )'], ['\\begin{cases} a & x > 0 \\\\ b & \\text{sonst} \\end{cases}', '{ cases'], ['\\binom{n}{k}', '(ⁿₖ)'], ['\\overline{x}', 'x̄'], ['\\vec{v}', 'v⃗'], ['\\hat{x}', 'x̂']]],
  ['Greek', [['\\alpha', 'α'], ['\\beta', 'β'], ['\\gamma', 'γ'], ['\\delta', 'δ'], ['\\epsilon', 'ε'], ['\\theta', 'θ'], ['\\lambda', 'λ'], ['\\mu', 'μ'], ['\\pi', 'π'], ['\\rho', 'ρ'], ['\\sigma', 'σ'], ['\\tau', 'τ'], ['\\phi', 'φ'], ['\\omega', 'ω'], ['\\Delta', 'Δ'], ['\\Sigma', 'Σ'], ['\\Omega', 'Ω']]],
  ['Operators', [['\\cdot', '·'], ['\\times', '×'], ['\\div', '÷'], ['\\pm', '±'], ['\\leq', '≤'], ['\\geq', '≥'], ['\\neq', '≠'], ['\\approx', '≈'], ['\\equiv', '≡'], ['\\infty', '∞'], ['\\partial', '∂'], ['\\nabla', '∇'], ['\\in', '∈'], ['\\subset', '⊂'], ['\\cup', '∪'], ['\\cap', '∩'], ['\\forall', '∀'], ['\\exists', '∃']]],
  ['Arrows', [['\\to', '→'], ['\\leftarrow', '←'], ['\\Rightarrow', '⇒'], ['\\Leftrightarrow', '⇔'], ['\\mapsto', '↦'], ['\\uparrow', '↑'], ['\\downarrow', '↓']]],
];
const SAMPLES = ['E = mc^2', 'a^2 + b^2 = c^2', 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}', '\\int_0^1 x^2 \\, dx = \\frac{1}{3}', '\\sum_{k=1}^{n} k = \\frac{n(n+1)}{2}', 'e^{i\\pi} + 1 = 0'];

export function EquationEditor({ latex: initial = '', display: initialDisplay = false, onSave, onClose }) {
  const [v, setV] = useState(initial);
  const [display, setDisplay] = useState(!!initialDisplay);
  const [res, setRes] = useState({ mathml: null, error: null, loading: true });
  const [group, setGroup] = useState('Structures');
  const ta = useRef(null);
  useEffect(() => {
    let alive = true;
    const h = setTimeout(() => {
      if (!v.trim()) { setRes({ mathml: null, error: null, loading: false }); return; }
      renderLatex(v, display).then((r) => alive && setRes({ ...r, loading: false }), () => alive && setRes({ mathml: null, error: null, loading: false, offline: true }));
    }, 90);
    return () => { alive = false; clearTimeout(h); };
  }, [v, display]);
  const insert = (snippet) => {
    const el = ta.current;
    const s = el ? el.selectionStart : v.length; const e = el ? el.selectionEnd : v.length;
    const sel = v.slice(s, e);
    // the selection goes into the first {…} of the snippet
    let ins = snippet;
    if (sel) ins = snippet.replace(/\{[^{}]*\}/, `{${sel}}`);
    const next = v.slice(0, s) + ins + v.slice(e);
    setV(next);
    setTimeout(() => { if (!el) return; el.focus(); const first = next.indexOf('{', s); const p = sel || first < 0 || first > s + ins.length ? s + ins.length : first + 1; const q = !sel && first >= 0 && first < s + ins.length ? next.indexOf('}', p) : p; el.setSelectionRange(p, q); }, 0);
  };
  const save = () => { if (v.trim() && !res.error) onSave(v.trim(), display, res.mathml); };
  return (
    <Modal size="wide" title="Formula" subtitle={t('Type LaTeX or click the symbols – the preview is how it is saved (Word: real Office maths).')} onClose={onClose}
      footer={<><Seg value={display ? 'block' : 'inline'} onChange={(x) => setDisplay(x === 'block')} options={[{ id: 'inline', label: 'In the line' }, { id: 'block', label: 'On its own line' }]} /><span className="grow" /><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" disabled={!v.trim() || !!res.error} onClick={save}>{initial ? 'Save' : 'Insert'}</Btn></>}>
      <div className="of-docs-eqed">
        <div className="of-docs-eq-preview of-docs-paper" aria-live="polite">
          {res.mathml ? <span className={`of-docs-eq ${display ? 'of-eq-block' : ''}`} dangerouslySetInnerHTML={{ __html: res.mathml }} /> : v.trim() ? <span className="of-docs-eq" dangerouslySetInnerHTML={{ __html: equationHtml(v) }} /> : <span className="muted small">{t('The preview appears here.')}</span>}
        </div>
        {res.error && <p className="small of-docs-eq-err">{t('Not valid LaTeX yet: {m}', { m: res.error })}</p>}
        <textarea ref={ta} className="input mono of-docs-eq-src" rows={3} autoFocus value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); save(); } }} aria-label="LaTeX" spellCheck={false} placeholder="\frac{a}{b} + x^2" />
        <div className="of-docs-eq-pal">
          <Seg value={group} onChange={setGroup} options={PALETTE.map(([g]) => ({ id: g, label: g }))} />
          <div className="of-docs-eq-keys">{PALETTE.find(([g]) => g === group)[1].map(([tex, label]) => <Raw as="button" type="button" key={tex} className="btn ghost sm of-docs-eq-key" title={tex} onClick={() => insert(tex)}>{label}</Raw>)}</div>
        </div>
        <div className="row-tight of-docs-tokens">{SAMPLES.map((s) => <Raw as="button" type="button" key={s} className="btn xs ghost mono" title={s} onClick={() => setV(s)}>{s.length > 28 ? `${s.slice(0, 26)}…` : s}</Raw>)}</div>
      </div>
    </Modal>
  );
}

/* ---------------- live fields ---------------- */
const DATE_FORMATS = ['dd.MM.yyyy', 'd. MMMM yyyy', 'dddd, d. MMMM yyyy', 'yyyy-MM-dd', 'MM/dd/yyyy', 'MMMM yyyy'];
export function fieldValue(kind, { format, page, pages, title, author, filename, words } = {}) {
  const now = new Date();
  switch (kind) {
    case 'date': return formatDate(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())).toISOString(), format || 'dd.MM.yyyy');
    case 'time': return now.toLocaleTimeString(document.documentElement.lang || 'de', { hour: '2-digit', minute: '2-digit' });
    case 'page': return String(page || 1);
    case 'pages': return String(pages || 1);
    case 'title': return title || '';
    case 'author': return author || '';
    case 'filename': return filename || '';
    case 'words': return String(words || 0);
    default: return '';
  }
}
export function FieldDialog({ host, initial, ctx, onSave, onClose }) {
  const [kind, setKind] = useState((initial && initial.kind) || 'date');
  const [format, setFormat] = useState((initial && /\\@ "([^"]+)"/.exec(initial.instr || '') ? /\\@ "([^"]+)"/.exec(initial.instr)[1] : null) || 'dd.MM.yyyy');
  const [cell, setCell] = useState(() => ({ path: '', sheet: '', cell: 'A1', ...((initial && initial.kind === 'cell' && initial.params) || {}) }));
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState(null);
  const local = ['date', 'time', 'page', 'pages', 'title', 'author', 'filename', 'words'];
  const preview = local.includes(kind) ? fieldValue(kind, { ...ctx, format }) : '';
  const fieldsOk = host && host.fields && typeof host.fields.value === 'function';
  const pickHostwerk = async () => {
    if (!host.fields || typeof host.fields.pick !== 'function') { setErr(t('Hostwerk values need a newer Office server.')); return; }
    try {
      const p = await host.fields.pick();
      if (!p) return;
      setBusy(true);
      const v = await host.fields.value(p.id, p.params);
      onSave({ kind: 'hostwerk', source: p.id, params: p.params || null, label: p.label || p.id, value: String((v && (v.text != null ? v.text : v.value)) ?? ''), updatedAt: (v && v.updatedAt) || new Date().toISOString(), instr: '' });
    } catch (e) { setErr(e.message || String(e)); } finally { setBusy(false); }
  };
  const saveCell = async () => {
    setBusy(true); setErr(null);
    try {
      const params = { path: cell.path.trim(), sheet: cell.sheet.trim() || null, cell: cell.cell.trim().toUpperCase() };
      const v = await host.fields.value('cell', params);
      onSave({ kind: 'cell', source: 'cell', params, label: `${params.path.split('/').pop()}${params.sheet ? `!${params.sheet}` : ''}!${params.cell}`, value: String((v && (v.text != null ? v.text : v.value)) ?? ''), updatedAt: (v && v.updatedAt) || new Date().toISOString(), instr: '' });
    } catch (e) { setErr(e.message || String(e)); } finally { setBusy(false); }
  };
  const save = () => {
    if (kind === 'hostwerk') { pickHostwerk(); return; }
    if (kind === 'cell') { saveCell(); return; }
    const instr = kind === 'date' ? `DATE \\@ "${format}"` : kind === 'time' ? 'TIME \\@ "HH:mm"' : '';
    onSave({ kind, value: preview, instr, label: t(FIELD_LABELS[kind]), updatedAt: new Date().toISOString(), source: null, params: null });
  };
  const KINDS = [...local, 'hostwerk', 'cell'];
  return (
    <Modal title="Insert field" subtitle={t('Fields keep their last value in the file and update when you open it (Word and Google show the last value).')} onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" disabled={busy || (kind === 'cell' && (!cell.path.trim() || !/^[A-Z]{1,3}\d{1,7}$/i.test(cell.cell.trim())))} onClick={save}>{busy ? <Spinner size={14} /> : kind === 'hostwerk' ? t('Choose a value …') : initial ? t('Save') : t('Insert')}</Btn></>}>
      <div className="stack">
        <div className="of-docs-fieldkinds" role="listbox" aria-label={t('Field')}>
          {KINDS.map((k) => <button type="button" key={k} role="option" aria-selected={kind === k} className={`of-docs-fieldkind ${kind === k ? 'of-on' : ''}`} onClick={() => setKind(k)}>{t(FIELD_LABELS[k])}</button>)}
        </div>
        {kind === 'date' && <Field label="Format"><select className="input" value={format} onChange={(e) => setFormat(e.target.value)} aria-label="Format">{DATE_FORMATS.map((f) => <option key={f} value={f}>{`${fieldValue('date', { format: f })}  (${f})`}</option>)}</select></Field>}
        {kind === 'hostwerk' && <p className="small muted">{t('A value from Hostwerk – for example the CPU load, the players on a game server or the state of a backup. Hostwerk shows its catalogue next.')}</p>}
        {kind === 'cell' && (
          <div className="grid grid-3">
            <Field label="Spreadsheet" hint="Path in your Office documents"><input className="input" value={cell.path} onChange={(e) => setCell({ ...cell, path: e.target.value })} placeholder="Finanzen/Budget.xlsx" aria-label="Spreadsheet" /></Field>
            <Field label="Sheet"><input className="input" value={cell.sheet} onChange={(e) => setCell({ ...cell, sheet: e.target.value })} placeholder={t('first sheet')} aria-label="Sheet" /></Field>
            <Field label="Cell"><input className="input mono" value={cell.cell} onChange={(e) => setCell({ ...cell, cell: e.target.value })} placeholder="B4" aria-label="Cell" /></Field>
          </div>
        )}
        {(kind === 'hostwerk' || kind === 'cell') && !fieldsOk && <Callout kind="warn" dismissible={false}>{t('This needs a newer Office server.')}</Callout>}
        {preview && <div className="of-docs-fieldprev"><span className="muted small">{t('Shows now:')}</span> <span className="of-docs-field"><Keep>{preview}</Keep></span></div>}
        {err && <Callout kind="err" dismissible={false}><Keep>{err}</Keep></Callout>}
      </div>
    </Modal>
  );
}

/* ---------------- diagrams ---------------- */
const DIAGRAMS = [
  ['Flowchart', 'flowchart TD\n  A[Start] --> B{Alles klar?}\n  B -- Ja --> C[Weiter]\n  B -- Nein --> D[Nachbessern]\n  D --> B'],
  ['Sequence', 'sequenceDiagram\n  participant Kunde\n  participant Shop\n  Kunde->>Shop: Bestellung\n  Shop-->>Kunde: Bestätigung'],
  ['Gantt', 'gantt\n  title Projektplan\n  dateFormat YYYY-MM-DD\n  section Planung\n  Konzept :a1, 2026-01-05, 10d\n  section Umsetzung\n  Bau :after a1, 20d'],
  ['Pie', 'pie title Verteilung\n  "A" : 45\n  "B" : 30\n  "C" : 25'],
  ['Mind map', 'mindmap\n  root((Idee))\n    Ziele\n    Risiken\n    Team'],
  ['Class', 'classDiagram\n  class Kunde {\n    +name\n    +bestellen()\n  }\n  Kunde --> Bestellung'],
];
export function DiagramDialog({ host, initial, onSave, onClose }) {
  const [src, setSrc] = useState((initial && initial.source) || DIAGRAMS[0][1]);
  const [out, setOut] = useState(null);
  const [state, setState] = useState('loading');
  const [err, setErr] = useState(null);
  const [saving, setSaving] = useState(false);
  useEffect(() => {
    let alive = true;
    const h = setTimeout(() => {
      renderMermaid(host, src, { ask: true }).then((r) => { if (!alive) return; setOut(r); setErr(null); setState('ok'); }, (e) => { if (!alive) return; if (e && (e.code === 'NO_COMPONENTS' || e.code === 'NOT_INSTALLED')) setState('missing'); else { setErr(String((e && e.message) || e).split('\n')[0]); setState('error'); } });
    }, state === 'loading' ? 0 : 450);
    return () => { alive = false; clearTimeout(h); };
  }, [src]); // eslint-disable-line react-hooks/exhaustive-deps
  const save = async () => {
    if (!out) { onSave({ source: src, src: (initial && initial.src) || null, width: (initial && initial.width) || null, height: (initial && initial.height) || null, kind: 'mermaid' }); return; }
    setSaving(true);
    try {
      const w = Math.min(out.width, 640); const h = Math.round(out.height * (w / out.width));
      const png = await svgToPng(out.svg, out.width, out.height, 2);
      onSave({ source: src, src: png, width: w, height: h, kind: 'mermaid' });
    } catch (e) { toast.error(e); setSaving(false); }
  };
  return (
    <Modal size="wide" title="Diagram from text" subtitle={t('Write the diagram as text (Mermaid) – it is saved as a picture that Word and Google show, with the text kept for editing here.')} onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" disabled={saving || state === 'error' || !src.trim()} onClick={save}>{saving ? <Spinner size={14} /> : initial ? 'Save' : 'Insert'}</Btn></>}>
      <div className="of-docs-dgm">
        <div className="stack">
          <div className="row-tight of-docs-tokens">{DIAGRAMS.map(([n, s]) => <Btn key={n} className="xs ghost" onClick={() => setSrc(s)}>{n}</Btn>)}</div>
          <textarea className="input mono of-docs-dgm-src" value={src} onChange={(e) => setSrc(e.target.value)} spellCheck={false} aria-label="Mermaid" rows={14} />
          <a className="small" href="#" onClick={(e) => { e.preventDefault(); if (host.openLink) host.openLink('https://mermaid.js.org/intro/syntax-reference.html'); }}>{t('Mermaid syntax')}</a>
        </div>
        <div className="of-docs-dgm-prev of-docs-paper">
          {state === 'loading' && <Spinner />}
          {state === 'ok' && out && <div dangerouslySetInnerHTML={{ __html: out.svg }} />}
          {state === 'error' && <p className="small of-docs-eq-err"><Keep>{err}</Keep></p>}
          {state === 'missing' && <Callout kind="info" dismissible={false} title="Diagram component">{t('Diagrams are drawn by a component Office downloads once. Without it the text is kept and the last picture is shown.')}</Callout>}
        </div>
      </div>
    </Modal>
  );
}

/* ---------------- AI ---------------- */
const AI_MODES = [
  { id: 'write', label: 'Write', icon: Sparkles, needs: null },
  { id: 'rewrite', label: 'Rewrite', icon: Wand2, needs: 'selection' },
  { id: 'summarize', label: 'Summarise', icon: FileText, needs: null },
  { id: 'translate', label: 'Translate', icon: ArrowRight, needs: null },
  { id: 'continue', label: 'Continue', icon: Plus, needs: null },
];
const TONES = ['clearer', 'shorter', 'friendlier', 'more formal', 'simpler', 'fix spelling and grammar'];
export function AiDialog({ host, mode: m0 = 'write', selection = '', context = '', onInsert, onClose }) {
  const [mode, setMode] = useState(m0 === 'rewrite' && !selection ? 'write' : m0);
  const [prompt, setPrompt] = useState('');
  const [tone, setTone] = useState(TONES[0]);
  const [lang, setLang] = useState(document.documentElement.lang === 'de' ? 'English' : 'Deutsch');
  const [busy, setBusy] = useState(false);
  const [out, setOut] = useState('');
  const [err, setErr] = useState(null);
  const ready = host && host.ai && (typeof host.ai.ready !== 'function' || host.ai.ready());
  const run = async () => {
    setBusy(true); setErr(null); setOut('');
    const base = selection || context;
    const req = {
      write: { prompt: prompt || t('Write a short paragraph.'), text: context.slice(-3000), system: 'You write text for a document. Answer with the text only (Markdown allowed for headings, lists, bold), in the language of the request.' },
      rewrite: { prompt: `Rewrite this text: ${tone}. ${prompt}`.trim(), text: selection, system: 'You rewrite document text. Answer with the rewritten text only, same language, keep the meaning.' },
      summarize: { prompt: `Summarise this text. ${prompt}`.trim(), text: base.slice(0, 20000), system: 'You summarise document text. Answer with the summary only (short paragraphs or a list), in the language of the text.' },
      translate: { prompt: `Translate this text into ${lang}. ${prompt}`.trim(), text: base.slice(0, 20000), system: 'You translate document text. Answer with the translation only, keep the structure (paragraphs, lists).' },
      continue: { prompt: `Continue this text naturally by one or two paragraphs. ${prompt}`.trim(), text: context.slice(-4000), system: 'You continue document text. Answer with the new text only, same language and style.' },
    }[mode];
    try {
      const r = await host.ai.ask(req);
      setOut(String((r && (r.text != null ? r.text : r)) || '').trim());
    } catch (e) { setErr(e.message || String(e)); } finally { setBusy(false); }
  };
  return (
    <Modal size="wide" title="AI" subtitle={t('Uses the AI you set up in Hostwerk. The result is inserted as normal text – Ctrl+Z takes it back.')} onClose={onClose}
      footer={<><Btn onClick={onClose}>Close</Btn><span className="grow" />{out && <Btn icon={Copy} onClick={() => { navigator.clipboard.writeText(out).then(() => toast(t('Copied')), () => {}); }}>Copy</Btn>}{out && selection && mode !== 'continue' && <Btn icon={Replace} onClick={() => onInsert(out, true)}>Replace selection</Btn>}{out && <Btn className="primary" icon={Check} onClick={() => onInsert(out, false)}>Insert</Btn>}{!out && <Btn className="primary" icon={Sparkles} disabled={busy || !ready || (mode === 'write' && !prompt.trim())} onClick={run}>{busy ? <Spinner size={14} /> : 'Go'}</Btn>}</>}>
      <div className="stack">
        {!ready && <Callout kind="info" dismissible={false}>{t('Set up an AI provider in Hostwerk’s settings first.')}</Callout>}
        <div className="of-docs-aimodes">{AI_MODES.map((x) => { const I = x.icon; const off = x.needs === 'selection' && !selection; return <button type="button" key={x.id} disabled={off} className={`of-docs-aimode ${mode === x.id ? 'of-on' : ''}`} onClick={() => { setMode(x.id); setOut(''); }} title={off ? t('Select text first') : t(x.label)}><I size={16} aria-hidden="true" />{t(x.label)}</button>; })}</div>
        {mode === 'rewrite' && <div className="row-tight of-docs-tokens">{TONES.map((x) => <Btn key={x} className={`xs ${tone === x ? '' : 'ghost'}`} onClick={() => setTone(x)}>{x}</Btn>)}</div>}
        {mode === 'translate' && <Field label="Into"><select className="input" value={lang} onChange={(e) => setLang(e.target.value)} aria-label="Into">{['Deutsch', 'English', 'Français', 'Español', 'Italiano', 'Nederlands', 'Polski', 'Türkçe', 'Українська', '中文', '日本語'].map((l) => <Raw as="option" key={l} value={l}>{l}</Raw>)}</select></Field>}
        <textarea className="input" rows={mode === 'write' ? 3 : 2} autoFocus value={prompt} onChange={(e) => setPrompt(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) run(); }} placeholder={mode === 'write' ? t('What should be written? e.g. “An invitation to the summer party on 12 July”') : t('Optional: what to pay attention to')} aria-label="Prompt" />
        {(mode === 'rewrite' || mode === 'summarize' || mode === 'translate') && (selection || context) && <div className="of-docs-aisrc small muted"><Keep>{(selection || context).slice(0, 400)}{(selection || context).length > 400 ? '…' : ''}</Keep></div>}
        {busy && <div className="row-tight"><Spinner size={14} /><span className="small muted">{t('The AI is writing …')}</span></div>}
        {out && <div className="of-docs-aiout of-docs-paper" dangerouslySetInnerHTML={{ __html: mdToHtml(out) }} />}
        {out && <Btn className="xs ghost" onClick={run}>Try again</Btn>}
        {err && <Callout kind="err" dismissible={false}><Keep>{err}</Keep></Callout>}
      </div>
    </Modal>
  );
}

/* ---------------- headers and footers (rich) ---------------- */
/** plain text of a header/footer document with {page}/{pages}/{date} tokens (RTF/ODT writers, summaries) */
export function hfText(doc) {
  const out = [];
  walk(doc, (n) => {
    if (n.type === 'paragraph' || n.type === 'heading') {
      out.push((n.content || []).map((c) => (c.type === 'text' ? c.text : c.type === 'field' ? ({ page: '{page}', pages: '{pages}', date: '{date}', title: '{title}' }[c.attrs.kind] || c.attrs.value || '') : '')).join(''));
      return false;
    }
    return undefined;
  });
  return out.filter((x) => x.trim()).join(' ').replace(/\s+/g, ' ').trim();
}
/** an old plain header ({ text, align }) as a document */
function hfDoc(x) {
  if (x && x.doc) return x.doc;
  const text = (x && x.text) || '';
  const content = [];
  for (const part of text.split(/(\{page\}|\{pages\}|\{date\}|\{title\})/)) {
    if (!part) continue;
    const m = /^\{(page|pages|date|title)\}$/.exec(part);
    content.push(m ? { type: 'field', attrs: { kind: m[1], value: m[1] === 'page' || m[1] === 'pages' ? '1' : '' } } : { type: 'text', text: part });
  }
  return { type: 'doc', content: [{ type: 'paragraph', attrs: { textAlign: (x && x.align) || 'center' }, content: content.length ? content : undefined }] };
}
function HfEditor({ value, onChange, label }) {
  const editor = useEditor({
    extensions: [StarterKit.configure({ heading: false, codeBlock: false, blockquote: false, horizontalRule: false }), Underline, TextStyle, Color, FontFamily, FontSize, ParagraphFormat, TextAlign.configure({ types: ['paragraph'] }), ImageEx.configure({ allowBase64: true, inline: true }), Table.configure({ resizable: false }), TableRow, TableCell, TableHeader, FieldNode],
    content: value,
    editorProps: { attributes: { class: 'of-docs-pm of-docs-hfpm', 'aria-label': label } },
    onUpdate: ({ editor: ed }) => onChange(ed.getJSON()),
  }, []);
  if (!editor) return null;
  const c = () => editor.chain().focus();
  const field = (kind) => c().insertContent({ type: 'field', attrs: { kind, value: kind === 'page' || kind === 'pages' ? '1' : fieldValue(kind, {}), instr: kind === 'date' ? 'DATE \\@ "dd.MM.yyyy"' : '' } }).run();
  const picture = () => { const i = document.createElement('input'); i.type = 'file'; i.accept = 'image/*'; i.onchange = () => { const f = i.files && i.files[0]; if (!f) return; const fr = new FileReader(); fr.onload = () => c().setImage({ src: fr.result, width: 120 }).run(); fr.readAsDataURL(f); }; i.click(); };
  const B = ({ icon: I, title, on, onClick }) => <button type="button" className={`btn ghost sm icon ${on ? 'of-on' : ''}`} title={t(title)} aria-label={t(title)} onMouseDown={(e) => e.preventDefault()} onClick={onClick}><I size={15} /></button>;
  return (
    <div className="of-docs-hfed">
      <div className="of-docs-hfbar">
        <B icon={Bold} title="Bold" on={editor.isActive('bold')} onClick={() => c().toggleBold().run()} />
        <B icon={Italic} title="Italic" on={editor.isActive('italic')} onClick={() => c().toggleItalic().run()} />
        <B icon={UIcon} title="Underline" on={editor.isActive('underline')} onClick={() => c().toggleUnderline().run()} />
        <span className="of-docs-hfsep" />
        <B icon={AlignLeft} title="Align left" on={editor.isActive({ textAlign: 'left' })} onClick={() => c().setTextAlign('left').run()} />
        <B icon={AlignCenter} title="Centre" on={editor.isActive({ textAlign: 'center' })} onClick={() => c().setTextAlign('center').run()} />
        <B icon={AlignRight} title="Align right" on={editor.isActive({ textAlign: 'right' })} onClick={() => c().setTextAlign('right').run()} />
        <span className="of-docs-hfsep" />
        <B icon={Hash} title="Page number" onClick={() => field('page')} />
        <B icon={Sigma} title="Number of pages" onClick={() => field('pages')} />
        <B icon={CalendarDays} title="Date" onClick={() => field('date')} />
        <B icon={FileText} title="Title" onClick={() => field('title')} />
        <span className="of-docs-hfsep" />
        <B icon={ImageIcon} title="Picture (logo)" onClick={picture} />
        <B icon={TableIcon} title="Table with three columns (left · centre · right)" onClick={() => c().insertTable({ rows: 1, cols: 3, withHeaderRow: false }).run()} />
        <button type="button" className="btn ghost sm" onMouseDown={(e) => e.preventDefault()} onClick={() => c().insertContent([{ type: 'text', text: t('Page') + ' ' }, { type: 'field', attrs: { kind: 'page', value: '1' } }, { type: 'text', text: ` ${t('of')} ` }, { type: 'field', attrs: { kind: 'pages', value: '1' } }]).run()}>{t('“Page X of Y”')}</button>
      </div>
      <div className="of-docs-hfpaper of-docs-paper"><EditorContent editor={editor} /></div>
    </div>
  );
}
export function HeaderFooterEditor({ model, onSave, onClose }) {
  const [first, setFirst] = useState(!!model.titlePage);
  const [evenOdd, setEvenOdd] = useState(!!model.evenOdd);
  const [docs, setDocs] = useState(() => Object.fromEntries(['header', 'footer', 'headerFirst', 'footerFirst', 'headerEven', 'footerEven'].map((k) => [k, hfDoc(model[k])])));
  const [part, setPart] = useState('header');
  const [which, setWhich] = useState('');
  const key = `${part}${which}`;
  const pages = [{ id: '', label: evenOdd ? 'Odd pages' : 'All pages' }, ...(first ? [{ id: 'First', label: 'First page' }] : []), ...(evenOdd ? [{ id: 'Even', label: 'Even pages' }] : [])];
  useEffect(() => { if (!pages.some((p) => p.id === which)) setWhich(''); }, [first, evenOdd]); // eslint-disable-line react-hooks/exhaustive-deps
  const save = () => {
    const patch = { titlePage: first, evenOdd };
    for (const [k, d] of Object.entries(docs)) {
      if ((k.endsWith('First') && !first) || (k.endsWith('Even') && !evenOdd)) { patch[k] = null; continue; }
      const p = (d.content || []).find((n) => n.type === 'paragraph');
      patch[k] = { text: hfText(d), align: (p && p.attrs && p.attrs.textAlign) || 'left', doc: d };
    }
    onSave(patch);
  };
  return (
    <Modal size="wide" title="Header and footer" subtitle={t('Format them like normal text – with pictures, tables and page numbers.')} onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" onClick={save}>Save</Btn></>}>
      <div className="stack">
        <div className="row of-docs-hfopts">
          <Seg value={part} onChange={setPart} options={[{ id: 'header', label: 'Header' }, { id: 'footer', label: 'Footer' }]} />
          <Seg value={which} onChange={setWhich} options={pages} />
        </div>
        <HfEditor key={key} label={t(part === 'header' ? 'Header' : 'Footer')} value={docs[key]} onChange={(d) => setDocs((x) => ({ ...x, [key]: d }))} />
        <div className="row of-docs-hfopts">
          <label className="check"><Switch checked={first} onChange={setFirst} /> {t('Different first page')}</label>
          <label className="check"><Switch checked={evenOdd} onChange={setEvenOdd} /> {t('Different odd and even pages')}</label>
        </div>
      </div>
    </Modal>
  );
}

/* ---------------- page setup of one section ---------------- */
export function SectionDialog({ attrs, units, onSave, onClose }) {
  const p0 = attrs.page || {};
  const inch = units === 'in';
  const conv = (mm) => (inch ? Math.round((mm / 25.4) * 100) / 100 : Math.round(mm / 10 * 100) / 100);
  const back = (v) => (inch ? Number(v) * 25.4 : Number(v) * 10);
  const [size, setSize] = useState(PAGE_SIZES[p0.size] ? p0.size : 'A4');
  const [orient, setOrient] = useState(p0.orientation || 'portrait');
  const [m, setM] = useState(() => { const mm = p0.margins || { top: 25, right: 25, bottom: 20, left: 25 }; return { top: conv(mm.top), right: conv(mm.right), bottom: conv(mm.bottom), left: conv(mm.left) }; });
  const [type, setType] = useState(attrs.type || 'nextPage');
  const [first, setFirst] = useState(!!attrs.titlePage);
  const u = inch ? 'in' : 'cm';
  const num = (k) => <Field label={`${t({ top: 'Top', right: 'Right', bottom: 'Bottom', left: 'Left' }[k])} (${u})`}><input className="input" inputMode="decimal" value={m[k]} onChange={(e) => setM({ ...m, [k]: e.target.value.replace(',', '.') })} aria-label={k} /></Field>;
  return (
    <Modal title="Page setup of this section" subtitle={t('Applies to the pages before this section break.')} onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancel</Btn><Btn className="primary" onClick={() => onSave({ type, titlePage: first, page: { ...p0, size, ...PAGE_SIZES[size], orientation: orient, margins: { top: back(m.top), right: back(m.right), bottom: back(m.bottom), left: back(m.left) }, columns: p0.columns || 1 } })}>Save</Btn></>}>
      <div className="stack">
        <div className="row">
          <Field label="Size"><select className="input" value={size} onChange={(e) => setSize(e.target.value)} aria-label="Size">{Object.keys(PAGE_SIZES).map((k) => <option key={k} value={k}>{k}</option>)}</select></Field>
          <Field label="Orientation"><Seg value={orient} onChange={setOrient} options={[{ id: 'portrait', label: 'Portrait' }, { id: 'landscape', label: 'Landscape' }]} /></Field>
        </div>
        <div className="grid grid-2">{num('top')}{num('bottom')}{num('left')}{num('right')}</div>
        <Field label="Section starts"><select className="input" value={type} onChange={(e) => setType(e.target.value)} aria-label="Section starts"><option value="nextPage">{t('On the next page')}</option><option value="continuous">{t('On the same page (continuous)')}</option><option value="evenPage">{t('On the next even page')}</option><option value="oddPage">{t('On the next odd page')}</option></select></Field>
        <label className="check"><Switch checked={first} onChange={setFirst} /> {t('Different first page')}</label>
      </div>
    </Modal>
  );
}

/* ---------------- compare ---------------- */
let active = null;
/** the open editor registers "apply the comparison as tracked changes" here (the shell's compare view calls it) */
export function registerCompareTarget(fn) { active = fn; return () => { if (active === fn) active = null; }; }

// objects in the text (diagram, field, formula, picture) are one U+FFFC in the diff – shown as a small diamond, not as a box
const objMark = (s) => String(s || '').replace(/￼/g, '◆');
function blockLabel(n) { if (!n) return ''; if (['paragraph', 'heading', 'codeBlock'].includes(n.type)) return objMark(inlineText(n)); return docToHtml({ type: 'doc', content: [n] }).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); }

/** side-by-side and inline view of two document versions (the chunk's Compare export) */
export function Compare({ a, b, ext, host }) {
  const [res, setRes] = useState(null);
  const [err, setErr] = useState(null);
  const [view, setView] = useState('inline');
  const [onlyChanges, setOnlyChanges] = useState(true);
  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const e = String(ext || 'docx').toLowerCase();
        const [ma, mb, lib] = await Promise.all([readDocument(a.bytes, e), readDocument(b.bytes, e), compareLib()]);
        if (!alive) return;
        setRes({ diff: lib.compareDocs(ma.model.doc, mb.model.doc), docA: ma.model.doc, docB: mb.model.doc });
      } catch (x) { if (alive) setErr(x.message || String(x)); }
    })();
    return () => { alive = false; };
  }, [a, b, ext]);
  if (err) return <Callout kind="err" dismissible={false}><Keep>{err}</Keep></Callout>;
  if (!res) return <div className="of-docs-center"><Spinner /></div>;
  const { blocks, stats } = res.diff;
  const shown = onlyChanges ? blocks.filter((x, k) => x.kind !== 'same' || (blocks[k - 1] && blocks[k - 1].kind !== 'same') || (blocks[k + 1] && blocks[k + 1].kind !== 'same')) : blocks;
  const canApply = !!active;
  return (
    <div className="of-docs-cmp">
      <div className="of-docs-cmp-head">
        <span className="pill">{t('{n} words added', { n: stats.insertedWords })}</span>
        <span className="pill">{t('{n} words removed', { n: stats.deletedWords })}</span>
        <span className="pill">{t('{n} paragraphs changed', { n: stats.changed + stats.added + stats.removed })}</span>
        <span className="grow" />
        <label className="check small"><input type="checkbox" checked={onlyChanges} onChange={(e) => setOnlyChanges(e.target.checked)} /> {t('Only changes')}</label>
        <Seg value={view} onChange={setView} options={[{ id: 'inline', label: 'Inline' }, { id: 'side', label: 'Side by side' }]} />
        {canApply && <Btn className="primary sm" icon={GitCompare} onClick={() => { Promise.resolve(active(res.docA, res.docB, a.label)).then(() => toast(t('The differences are now tracked changes – accept or reject them in Review.')), (x) => toast.error(x)); }}>Apply as tracked changes</Btn>}
      </div>
      {view === 'side' && <div className="of-docs-cmp-cols small muted"><span><Keep>{a.label || t('Before')}</Keep></span><span><Keep>{b.label || t('Now')}</Keep></span></div>}
      <div className={`of-docs-cmp-body of-docs-paper ${view === 'side' ? 'of-side' : ''}`}>
        {!shown.some((x) => x.kind !== 'same') && <p className="muted">{t('Both versions have the same text.')}</p>}
        {shown.map((x, k) => {
          if (view === 'side') {
            return (
              <div key={k} className={`of-docs-cmp-row of-k-${x.kind}`}>
                <div>{x.kind === 'changed' ? x.parts.filter((p) => p.op <= 0).map((p, j) => <span key={j} className={p.op < 0 ? 'of-docs-del' : ''}><Keep>{objMark(p.text)}</Keep></span>) : <Keep>{blockLabel(x.a)}</Keep>}</div>
                <div>{x.kind === 'changed' ? x.parts.filter((p) => p.op >= 0).map((p, j) => <span key={j} className={p.op > 0 ? 'of-docs-ins' : ''}><Keep>{objMark(p.text)}</Keep></span>) : <Keep>{blockLabel(x.b)}</Keep>}</div>
              </div>
            );
          }
          return (
            <p key={k} className={`of-docs-cmp-p of-k-${x.kind}`}>
              {x.kind === 'changed' ? x.parts.map((p, j) => <span key={j} className={p.op < 0 ? 'of-docs-del' : p.op > 0 ? 'of-docs-ins' : ''}><Keep>{objMark(p.text)}</Keep></span>)
                : x.kind === 'added' ? <span className="of-docs-ins"><Keep>{blockLabel(x.b)}</Keep></span>
                  : x.kind === 'removed' ? <span className="of-docs-del"><Keep>{blockLabel(x.a)}</Keep></span> : <Keep>{blockLabel(x.b)}</Keep>}
            </p>
          );
        })}
      </div>
    </div>
  );
}
void Columns2; void Rows3; void useMemo;
