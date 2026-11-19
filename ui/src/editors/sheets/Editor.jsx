/*
 * Spreadsheet editor: loads the workbook, owns selection/editing state and every command (ribbon, keyboard, context
 * menus), registers saver/exporter with the shell. Rendering of the cells lives in Grid.jsx / render.js.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Loading, Empty, Callout, Btn, toast, confirmDialog, promptDialog } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Keep, Raw } from '../_shared/keep.jsx';
import { Check, X, SquareFunction, Plus, ChevronLeft, ChevronRight, AlertTriangle, ExternalLink, ChevronDown, Info, Minus, CloudOff } from 'lucide-react';
import { Book } from './model.js';
import { Grid } from './Grid.jsx';
import { Ribbon } from './Ribbon.jsx';
import { CellInput } from './CellInput.jsx';
import { MenuHost, showMenu, showPopover, closeMenu, pickColor, MenuList } from './menus.jsx';
import { makeLocale } from './format/locale.js';
import { FORMAT_PRESETS, fmtKind, changeDecimals, makeCurrencyFormat, ACCOUNTING, formatValue, isDateFormat } from './format/numfmt.js';
import { todaySerial, nowSerial } from './format/dates.js';
import { MAXR, MAXC, addr, colName, parseRange, rangeText, rangeA1, quoteSheet } from './formula/refs.js';
import { refsIn, toDisplay as fToDisplay, toInternal, shiftFormula } from './formula/parser.js';
import { FUNCTION_NAMES } from './formula/functions.js';
import { FUNC_INFO, callAt } from './formula/funcinfo.js';
import { localFn, internalFn, localArgList, isExcelName } from './formula/i18n.js';
import { isErr } from './formula/errors.js';
import { REF_COLORS } from './render.js';
import { axesOf } from './layout.js';
import { resolveChartData, PALETTE } from './chartData.js';
import { openBytes, saveBytes, bookToHTML, csvToBook, OOXML, CSV } from './io/formats.js';
import { writeXlsx, vbaProjectBytes } from './io/xlsx.js';
import { readVba } from './io/vba.js';
import { toTSV, toHTML, parseTSV, parseHTML } from './clipboard.js';
import { FormatCellsDialog } from './dialogs/FormatCells.jsx';
import { SortDialog, FilterMenu, TextToColumnsDialog, RemoveDuplicatesDialog, ValidationDialog, CsvDialog } from './dialogs/Data.jsx';
import { RuleDialog, RulesManager, CF_STYLES } from './dialogs/CondFormat.jsx';
import { FindPanel, GoToDialog, NamesDialog, InsertFunctionDialog, LinkDialog, NoteDialog, ChartDialog, PrintDialog, PasteSpecialDialog, CellsDialog, MacrosSheet, StatsDialog, CommandSearch, UnhideSheetDialog } from './dialogs/Misc.jsx';

const K = (r, c) => r * MAXC + c;
const norm = (a, b) => ({ r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) });
const cellSel = (r, c) => ({ r, c, ranges: [{ r1: r, c1: c, r2: r, c2: c }], anchor: { r, c }, end: { r, c } });
const WRITABLE = new Set(['xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'xlt', 'xlsb', 'ods', 'ots', 'fods', 'csv', 'tsv', 'tab', 'dif', 'slk', 'dbf']);
const CELL_STYLES = [
  { id: 'normal', label: 'Normal', s: null },
  { id: 'good', label: 'Good', s: { fill: '#C6EFCE', color: '#006100' } }, { id: 'bad', label: 'Bad', s: { fill: '#FFC7CE', color: '#9C0006' } }, { id: 'neutral', label: 'Neutral', s: { fill: '#FFEB9C', color: '#9C5700' } },
  { id: 'title', label: 'Title', s: { size: 18, color: '#44546A', b: true } }, { id: 'h1', label: 'Heading 1', s: { size: 15, b: true, color: '#44546A', bd: { b: { s: 'thick', c: '#4472C4' } } } },
  { id: 'h2', label: 'Heading 2', s: { size: 13, b: true, color: '#44546A', bd: { b: { s: 'thick', c: '#A2B8E1' } } } }, { id: 'h3', label: 'Heading 3', s: { b: true, color: '#44546A', bd: { b: { s: 'medium', c: '#8EA9DB' } } } },
  { id: 'total', label: 'Total', s: { b: true, bd: { t: { s: 'thin', c: '#4472C4' }, b: { s: 'double', c: '#4472C4' } } } },
  { id: 'input', label: 'Input', s: { fill: '#FFCC99', color: '#3F3F76', bd: { t: { s: 'thin', c: '#7F7F7F' }, b: { s: 'thin', c: '#7F7F7F' }, l: { s: 'thin', c: '#7F7F7F' }, r: { s: 'thin', c: '#7F7F7F' } } } },
  { id: 'output', label: 'Output', s: { fill: '#F2F2F2', color: '#3F3F3F', b: true, bd: { t: { s: 'thin', c: '#3F3F3F' }, b: { s: 'thin', c: '#3F3F3F' }, l: { s: 'thin', c: '#3F3F3F' }, r: { s: 'thin', c: '#3F3F3F' } } } },
  { id: 'calc', label: 'Calculation', s: { fill: '#F2F2F2', color: '#FA7D00', b: true, bd: { t: { s: 'thin', c: '#7F7F7F' }, b: { s: 'thin', c: '#7F7F7F' }, l: { s: 'thin', c: '#7F7F7F' }, r: { s: 'thin', c: '#7F7F7F' } } } },
  { id: 'note', label: 'Note', s: { fill: '#FFFFCC', bd: { t: { s: 'thin', c: '#B2B2B2' }, b: { s: 'thin', c: '#B2B2B2' }, l: { s: 'thin', c: '#B2B2B2' }, r: { s: 'thin', c: '#B2B2B2' } } } },
  { id: 'warn', label: 'Warning text', s: { color: '#FF0000' } }, { id: 'explan', label: 'Explanatory text', s: { color: '#7F7F7F', i: true } },
  { id: 'acc1', label: 'Accent 1', s: { fill: '#4472C4', color: '#FFFFFF' } }, { id: 'acc2', label: 'Accent 2', s: { fill: '#ED7D31', color: '#FFFFFF' } }, { id: 'acc6', label: 'Accent 6', s: { fill: '#70AD47', color: '#FFFFFF' } },
];

function uiLang(host) { const l = (host && host.lang) || document.documentElement.lang || 'en'; return String(l).toLowerCase().startsWith('de') ? 'de' : 'en'; }

export function Editor({ file, bytes, readOnly = false, settings = {}, host = {} }) {
  const S = settings.sheets || {};
  const lang = uiLang(host);
  const loc = useMemo(() => makeLocale({ lang, decimal: S.decimal || 'auto', dateFormat: S.dateFormat || 'auto', functionNames: S.functionNames || 'auto' }), [lang, S.decimal, S.dateFormat, S.functionNames]);
  const workExt = String((file.converted && file.converted.to) || file.ext || 'xlsx').toLowerCase();
  const saveExt = WRITABLE.has(workExt) ? workExt : 'xlsx';
  const [book, setBook] = useState(null);
  const [err, setErr] = useState(null);
  const [ver, setVer] = useState(0);
  const [sel, setSelState] = useState(cellSel(0, 0));
  const [edit, setEdit] = useState(null);
  const [zoomMap, setZoomMap] = useState({});
  const [dialog, setDialog] = useState(null);
  const [copy, setCopy] = useState(null);
  const [chartSel, setChartSel] = useState(null);
  const [painter, setPainter] = useState(null);
  const [showFormulas, setShowFormulas] = useState(false);
  const [formulaBar, setFormulaBar] = useState(S.formulaBar !== false);
  const [compact, setCompact] = useState(false);
  const [ac, setAc] = useState(null); // autocomplete { items, i }
  const [fid, setFid] = useState([]);
  const [csvNote, setCsvNote] = useState(true);
  const [showNotes, setShowNotes] = useState(false);
  const [last, setLast] = useState({ fill: '#FFFF00', color: '#C00000', border: 'bottom', line: { s: 'thin', c: '#000000' } });
  const gridApi = useRef(null);
  const input = useRef(null);
  const barInput = useRef(null);
  const loading = useRef(true);
  const pasteMode = useRef(null);
  const tabStart = useRef(null);
  const stRef = useRef({});
  const clipId = useRef(`c${Math.random().toString(36).slice(2)}`);
  const csvOpts = useRef(null);

  /* ---------- load ---------- */
  const load = useCallback(async (opts) => {
    try {
      loading.current = true;
      let b;
      if (!bytes || !bytes.length) b = Book.blank({ loc, settings: S, sheetName: lang === 'de' ? 'Tabelle1' : 'Sheet1' });
      else if (CSV.includes(workExt) || (workExt === 'txt')) b = csvToBook(bytes, { loc, settings: S, ext: workExt, ...(opts || {}) });
      else b = await openBytes(bytes, workExt, { loc, settings: S });
      if (!b.sheets.length) b = Book.blank({ loc, settings: S });
      if (S.autoCalc === false) b.autoCalc = false;
      if (b.fileManualCalc) b.autoCalc = false;
      b.undoStack = []; b.redoStack = [];
      setBook(b);
      const sh = b.sheets[b.active] || b.sheets[0];
      const s0 = sh.sel ? cellSel(sh.sel.r, sh.sel.c) : cellSel(0, 0);
      setSelState(s0);
      setVer((v) => v + 1);
      setFid(b.fidelity || []);
      setErr(null);
    } catch (e) { setErr(e); }
    finally { setTimeout(() => { loading.current = false; }, 0); }
  }, [bytes, workExt, loc]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { load(); }, [load]);
  useEffect(() => {
    if (!book) return undefined;
    return book.on((info) => {
      setVer((v) => v + 1);
      if (!loading.current && !readOnly && !(info && info.viewOnly) && host.dirty) host.dirty(true);
    });
  }, [book]); // eslint-disable-line react-hooks/exhaustive-deps
  // fidelity: let the shell show it (fallback: own callout)
  useEffect(() => { if (book && host.fidelity) host.fidelity({ items: fid, kept: true }); }, [book, fid]); // eslint-disable-line react-hooks/exhaustive-deps

  const sheet = book ? book.sheets[book.active] || book.sheets[0] : null;
  if (typeof window !== 'undefined' && window.__sheetsDebug) Object.assign(window.__sheetsDebug, { book, sel, edit });
  const zoom = sheet ? zoomMap[sheet.id] || sheet.zoom || S.zoom || 100 : 100;
  const defFont = (book && book.props.defaultFont) || { font: S.font || 'Calibri', size: S.fontSize || 11 };
  const toDisp = useCallback((f) => fToDisplay(f, loc), [loc]);
  const fromDisp = useCallback((f) => toInternal(String(f).replace(/^=/, ''), loc), [loc]);

  /* ---------- helpers ---------- */
  const guard = () => { if (readOnly) { toast(t('This file is open read-only.')); return false; } return true; };
  const activeRange = () => sel.ranges[sel.ranges.length - 1];
  const expandMerges = (g0) => {
    let g = { ...g0 }; let changed = true; let n = 0;
    while (changed && n++ < 20) { changed = false; for (const m of sheet.merges) if (m.r1 <= g.r2 && g.r1 <= m.r2 && m.c1 <= g.c2 && g.c1 <= m.c2 && (m.r1 < g.r1 || m.c1 < g.c1 || m.r2 > g.r2 || m.c2 > g.c2)) { g = { r1: Math.min(g.r1, m.r1), c1: Math.min(g.c1, m.c1), r2: Math.max(g.r2, m.r2), c2: Math.max(g.c2, m.c2) }; changed = true; } }
    return g;
  };
  const setSel = (s, { scroll = true } = {}) => {
    const n = { ...s, end: s.end || s.anchor || { r: s.r, c: s.c } };
    setSelState(n);
    if (sheet) sheet.sel = { r: n.r, c: n.c };
    if (scroll && gridApi.current) requestAnimationFrame(() => gridApi.current && gridApi.current.scrollTo(n.end ? n.end.r : n.r, n.end ? n.end.c : n.c));
    if (painter) applyPainter(n);
  };
  const go = (r, c) => { r = Math.max(0, Math.min(MAXR - 1, r)); c = Math.max(0, Math.min(MAXC - 1, c)); const m = book.mergeAt(sheet, r, c); setSel(m ? { r: m.r1, c: m.c1, ranges: [{ ...m }], anchor: { r: m.r1, c: m.c1 }, end: { r: m.r1, c: m.c1 } } : cellSel(r, c)); };
  const clipRange = (g) => { const b = book.bounds(sheet); return { ...g, r2: g.r2 >= MAXR - 1 ? Math.max(g.r1, b.rows - 1) : g.r2, c2: g.c2 >= MAXC - 1 ? Math.max(g.c1, b.cols - 1) : g.c2 }; };
  const filled = (r, c) => { const v = book.value(sheet, r, c); return v != null && v !== ''; };
  const region = (r, c) => {
    let g = { r1: r, c1: c, r2: r, c2: c };
    const any = (r1, c1, r2, c2) => { for (let y = Math.max(0, r1); y <= r2; y++) for (let x = Math.max(0, c1); x <= c2; x++) if (filled(y, x)) return true; return false; };
    for (let i = 0; i < 200; i++) {
      let ch = false;
      if (g.r1 > 0 && any(g.r1 - 1, g.c1 - 1, g.r1 - 1, g.c2 + 1)) { g.r1--; ch = true; }
      if (any(g.r2 + 1, g.c1 - 1, g.r2 + 1, g.c2 + 1) && g.r2 < MAXR - 1) { g.r2++; ch = true; }
      if (g.c1 > 0 && any(g.r1 - 1, g.c1 - 1, g.r2 + 1, g.c1 - 1)) { g.c1--; ch = true; }
      if (any(g.r1 - 1, g.c2 + 1, g.r2 + 1, g.c2 + 1) && g.c2 < MAXC - 1) { g.c2++; ch = true; }
      if (!ch) break;
    }
    return g;
  };
  /** the range a data command works on: the selection, or the current region around a single cell */
  const dataRange = () => { const g = activeRange(); if (g.r1 === g.r2 && g.c1 === g.c2) return region(sel.r, sel.c); return clipRange(g); };
  const getValues = useCallback((ref) => {
    const g = parseRange(ref); if (!g) return null;
    const sh = g.sheet ? book.sheetByName(g.sheet) : sheet; if (!sh) return null;
    const cg = { ...g, r2: Math.min(g.r2, g.r1 + 5000), c2: Math.min(g.c2, g.c1 + 200) };
    const values = []; const texts = [];
    for (let r = cg.r1; r <= cg.r2; r++) { if (sh.hiddenRows.has(r) || sh.filterHidden.has(r)) continue; const a = []; const tx = []; for (let c = cg.c1; c <= cg.c2; c++) { a.push(book.value(sh, r, c)); tx.push(book.text(sh, r, c)); } values.push(a); texts.push(tx); }
    return { values, texts };
  }, [book, sheet]);
  const chartCache = useRef(new Map());
  const chartData = (ch) => { const k = `${ver}|${ch.id}|${JSON.stringify([ch.range, ch.series, ch.byRows, ch.type, ch.colors])}`; let d = chartCache.current.get(k); if (!d) { d = resolveChartData(ch, getValues); if (chartCache.current.size > 200) chartCache.current.clear(); chartCache.current.set(k, d); } return d; };

  /* ---------- saver / exporter ---------- */
  stRef.current = { book, edit, sheet, sel, zoomMap };
  useEffect(() => {
    if (!book || !host.setSaver) return;
    const prep = () => { const st = stRef.current; if (st.edit) commit(null); for (const sh of st.book.sheets) { const z = st.zoomMap[sh.id]; if (z) sh.zoom = z; } };
    const opts = { getValues: (ref) => { const g = parseRange(ref); if (!g) return null; const sh = g.sheet ? book.sheetByName(g.sheet) : book.sheets[0]; if (!sh) return null; const values = []; const texts = []; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) { const a = []; const tx = []; for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 200); c++) { a.push(book.value(sh, r, c)); tx.push(book.text(sh, r, c)); } values.push(a); texts.push(tx); } return { values, texts }; }, settings: S, convert: host.hasConverter && host.convert ? (b, f, to) => host.convert(b, f, to) : null };
    if (!readOnly) host.setSaver(async () => {
      prep();
      const { bytes: out, note } = await saveBytes(book, saveExt, opts);
      if (note) setFid((f) => (f.includes(t(note)) ? f : [...f, t(note)]));
      return saveExt === workExt && !file.converted ? out : { bytes: out, ext: saveExt };
    }, { ext: saveExt });
    if (host.setExporter) {
      const formats = [['xlsx', 'Excel workbook'], ['xlsm', 'Excel workbook with macros'], ['xls', 'Excel 97–2003 workbook'], ['ods', 'OpenDocument spreadsheet'], ['fods', 'Flat OpenDocument spreadsheet'], ['csv', 'CSV table'], ['tsv', 'Tab-separated table'], ['html', 'Web page'], ['dif', 'Data Interchange Format'], ['slk', 'SYLK table']]
        .filter(([e]) => e !== 'xlsm' || ['xlsm', 'xltm'].includes(workExt) || (book.fidelity || []).some((x) => /Macros/.test(x)))
        .map(([ext, label]) => ({ ext, label: t(label) }));
      if (host.hasConverter && host.convert) formats.push({ ext: 'pdf', label: t('PDF document') });
      host.setExporter({ formats, run: async (ext) => {
        prep();
        if (ext === 'pdf') { const x = await writeXlsx(book, { ext: 'xlsx', getValues: opts.getValues }); return host.convert(x, 'xlsx', 'pdf'); }
        const { bytes: out } = await saveBytes(book, ext, { ...opts, sheet: book.active });
        return out;
      } });
    }
  }, [book, readOnly]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!book || !host.status || !sheet) return;
    host.status(book.sheets.length > 1 ? t('Sheet {n} of {total}', { n: book.active + 1, total: book.sheets.length }) : sheet.name);
  }, [book, ver]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- editing ---------- */
  const editSheet = edit ? book.sheetById(edit.sheetId) : sheet;
  const startEdit = (text, mode = 'enter', where = 'cell', at = null) => {
    if (!guard()) return;
    const r = at ? at.r : sel.r; const c = at ? at.c : sel.c;
    const sp = sheet.spill.get(K(r, c));
    if (sp && text === undefined) { go(Math.floor(sp.anchor / MAXC), sp.anchor % MAXC); }
    const full = text === undefined ? book.editText(sheet, r, c, toDisp) : text;
    setEdit({ r, c, sheetId: sheet.id, text: full, caret: full.length, mode, where, orig: book.editText(sheet, r, c, toDisp), point: null, pc: null });
    setAc(null);
    if (gridApi.current) gridApi.current.scrollTo(r, c);
    if (where === 'bar') setTimeout(() => barInput.current && barInput.current.focus(), 0);
    else setTimeout(() => { const el = input.current; if (el) { el.focus(); if (text === undefined) el.setSelectionRange(full.length, full.length); } }, 0);
  };
  const cancelEdit = () => { if (edit && edit.sheetId !== sheet.id) switchSheet(book.sheets.findIndex((s) => s.id === edit.sheetId)); setEdit(null); setAc(null); focusGrid(); };
  const focusGrid = () => setTimeout(() => input.current && input.current.focus(), 0);
  const dvAt = (sh, r, c) => sh.dv.find((d) => d.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)) || null;
  const evalRuleF = (sh, rule, f, r, c) => { if (f == null || f === '') return null; const g = rule.ranges[0]; return book.calc.evalText(shiftFormula(String(f), r - g.r1, c - g.c1), sh, r, c); };
  const dvList = (sh, rule, r, c) => {
    if (!rule.f1) return [];
    if (/^".*"$/.test(rule.f1)) return rule.f1.slice(1, -1).split(/[,;]/).map((x) => x.trim()).filter((x) => x !== '');
    const v = evalRuleF(sh, rule, rule.f1, r, c);
    const out = [];
    const walk = (x) => { if (Array.isArray(x)) x.forEach(walk); else if (x && x.sheet && x.r1 != null) { for (let y = x.r1; y <= Math.min(x.r2, x.r1 + 2000); y++) for (let z = x.c1; z <= x.c2; z++) { const t2 = book.text(x.sheet, y, z); if (t2 !== '') out.push(t2); } } else if (x != null && x !== '') out.push(String(x)); };
    // resolve ranges directly (evalText returns values for single cells only)
    const g = parseRange(String(rule.f1).replace(/\$/g, ''));
    if (g) { const sh2 = g.sheet ? book.sheetByName(g.sheet) : sh; if (sh2) walk({ sheet: sh2, ...g }); return [...new Set(out)]; }
    const nm = book.nameDef(String(rule.f1), sh);
    if (nm) { const g2 = parseRange(nm.ref.replace(/\$/g, '')); if (g2) { const sh2 = g2.sheet ? book.sheetByName(g2.sheet) : sh; if (sh2) walk({ sheet: sh2, ...g2 }); return [...new Set(out)]; } }
    walk(v);
    return [...new Set(out)];
  };
  const validValue = (sh, rule, r, c) => {
    const v = book.value(sh, r, c);
    if (v == null || v === '') return rule.allowBlank !== false;
    const cmp = (x) => {
      const a = evalRuleF(sh, rule, rule.f1, r, c); const b = evalRuleF(sh, rule, rule.f2, r, c);
      const an = typeof a === 'number' ? a : Number(a); const bn = typeof b === 'number' ? b : Number(b);
      switch (rule.op || 'between') {
        case 'between': return x >= Math.min(an, bn) && x <= Math.max(an, bn);
        case 'notBetween': return x < Math.min(an, bn) || x > Math.max(an, bn);
        case 'equal': return x === an; case 'notEqual': return x !== an;
        case 'greaterThan': return x > an; case 'lessThan': return x < an;
        case 'greaterThanOrEqual': return x >= an; case 'lessThanOrEqual': return x <= an;
        default: return true;
      }
    };
    switch (rule.type) {
      case 'list': { const list = dvList(sh, rule, r, c); return list.some((x) => x.toLowerCase() === String(book.text(sh, r, c)).toLowerCase() || x.toLowerCase() === String(v).toLowerCase()); }
      case 'whole': return typeof v === 'number' && Number.isInteger(v) && cmp(v);
      case 'decimal': case 'date': return typeof v === 'number' && cmp(v);
      case 'time': return typeof v === 'number' && cmp(v - Math.floor(v));
      case 'textLength': return cmp(String(v).length);
      case 'custom': { const x = evalRuleF(sh, rule, rule.f1, r, c); return x === true || x === 1; }
      default: return true;
    }
  };
  /** commit the editor; move = { dr, dc } | 'stay' | null; ctrlEnter fills the whole selection */
  const commit = async (move, { ctrlEnter = false } = {}) => {
    const e = edit; if (!e) return true;
    const sh = book.sheetById(e.sheetId);
    if (!sh) { setEdit(null); return true; }
    const text = e.text;
    if (text === e.orig && !ctrlEnter) { setEdit(null); setAc(null); if (sh !== sheet) switchSheet(book.sheets.indexOf(sh)); afterCommitMove(move); return true; }
    book.begin('Typing');
    try {
      book.input(sh, e.r, e.c, text);
      if (!text.startsWith('=') && text.includes('\n')) { const cell = sh.cells.get(K(e.r, e.c)); if (cell && !(cell.s && cell.s.wrap)) book.setCell(sh, e.r, e.c, { ...cell, s: book.styles.merge(cell.s, { wrap: true }) }); }
      if (ctrlEnter) {
        const src = sh.cells.get(K(e.r, e.c));
        for (const g of sel.ranges) for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 500); c++) {
          if (r === e.r && c === e.c) continue;
          const old = sh.cells.get(K(r, c));
          book.setCell(sh, r, c, src ? { ...(old || {}), v: src.v, f: src.f ? shiftFormula(src.f, r - e.r, c - e.c) : undefined, s: old && old.s ? old.s : src.s } : null);
        }
      }
    } finally { book.end(); }
    const rule = dvAt(sh, e.r, e.c);
    if (rule && rule.error && rule.error.show !== false && !validValue(sh, rule, e.r, e.c)) {
      const style = rule.error.style || 'stop';
      const msg = rule.error.text || t('This value doesn’t match the data validation rules for this cell.');
      if (style === 'stop') {
        book.undo(); book.redoStack.pop();
        toast.error(rule.error.title ? `${rule.error.title}: ${msg}` : msg);
        setTimeout(() => input.current && input.current.focus(), 0);
        return false;
      }
      if (style === 'warning') {
        const ok = await confirmDialog({ title: rule.error.title || t('Keep this value?'), text: msg, confirm: t('Keep') });
        if (!ok) { book.undo(); book.redoStack.pop(); setTimeout(() => input.current && input.current.focus(), 0); return false; }
      } else toast(msg);
    }
    setEdit(null); setAc(null);
    if (sh !== sheet) switchSheet(book.sheets.indexOf(sh));
    afterCommitMove(move);
    return true;
  };
  const afterCommitMove = (move) => { if (move && move !== 'stay') moveActive(move.dr, move.dc, true); focusGrid(); };

  /* ---------- point mode (click/arrow to insert references) ---------- */
  const formulaText = edit && edit.text.startsWith('=');
  const pointable = (e) => e && e.text.startsWith('=') && (e.point ? e.caret === e.point.end : /[=(,;+\-*/^&<>:]\s*$/.test(e.text.slice(0, e.caret)));
  const insertRef = (g, phase) => {
    if (!edit) return;
    const onOther = sheet.id !== edit.sheetId;
    const whole = g.r1 === 0 && g.r2 >= MAXR - 1 ? 'cols' : g.c1 === 0 && g.c2 >= MAXC - 1 ? 'rows' : null;
    let ref = rangeText({ ...g, sheet: onOther ? sheet.name : null, kind: whole || (g.r1 === g.r2 && g.c1 === g.c2 ? 'cell' : 'range') });
    if (loc.decimal === ',' && false) ref = ref.replace(',', ';');
    let { text, caret } = edit;
    let start = caret;
    if (edit.point && (phase === 'move' || phase === 'start' || phase === 'extend') && edit.point.end === caret) { start = edit.point.start; text = text.slice(0, start) + text.slice(edit.point.end); }
    else if (phase === 'add') { const sep = loc.decimal === ',' ? ';' : ','; text = `${text.slice(0, caret)}${sep}${text.slice(caret)}`; start = caret + 1; }
    const nt = text.slice(0, start) + ref + text.slice(start);
    const nc = start + ref.length;
    setEdit({ ...edit, text: nt, caret: nc, point: { start, end: nc }, pc: { r: g.r2, c: g.c2, a: { r: g.r1, c: g.c1 } } });
    setTimeout(() => { const el = edit.where === 'bar' || sheet.id !== edit.sheetId ? barInput.current : input.current; if (el) { el.focus(); el.setSelectionRange(nc, nc); } }, 0);
  };
  const refs = useMemo(() => {
    if (!edit || !edit.text.startsWith('=')) return null;
    const es = book.sheetById(edit.sheetId);
    try {
      return refsIn(edit.text.slice(1), { decimal: loc.decimal }).map((x, i) => ({ ...x, sheet: x.sheet ? book.sheetByName(x.sheet) : es, color: REF_COLORS[i % REF_COLORS.length], r1: Math.min(x.r1, x.r2), r2: Math.max(x.r1, x.r2), c1: Math.min(x.c1, x.c2), c2: Math.max(x.c1, x.c2) }));
    } catch { return null; }
  }, [edit && edit.text, book]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- autocomplete & hints ---------- */
  const updateAc = (text, caret) => {
    if (!text.startsWith('=')) { setAc(null); return; }
    const m = /(?:^=|[=(,;+\-*/^&<>:\s])([A-Za-z_À-￿][\wÀ-￿.]*)$/.exec(text.slice(0, caret));
    if (!m || m[1].length < 1) { setAc(null); return; }
    const q = m[1].toUpperCase();
    const names = [...book.names.values()].filter((n) => !n.raw && n.name.toUpperCase().startsWith(q)).map((n) => ({ name: n.name, kind: 'name' }));
    // function names in the user's language (German: SUMME, SVERWEIS …); the engine's own aliases are not offered
    const fns = FUNCTION_NAMES.filter(isExcelName).map((en) => ({ name: localFn(en, loc), en, kind: 'fn' })).filter((f) => f.name.startsWith(q))
      .sort((a, b) => (FUNC_INFO.has(b.en) ? 1 : 0) - (FUNC_INFO.has(a.en) ? 1 : 0) || a.name.length - b.name.length || a.name.localeCompare(b.name)).slice(0, 12);
    const items = [...names, ...fns].slice(0, 12);
    setAc(items.length && !(items.length === 1 && items[0].name === q && items[0].kind === 'name') ? { items, i: 0, token: m[1] } : null);
  };
  const acceptAc = (item) => {
    const e = edit; if (!e || !ac) return;
    const before = e.text.slice(0, e.caret - ac.token.length);
    const ins = item.kind === 'fn' ? `${item.name}(` : item.name;
    const nt = before + ins + e.text.slice(e.caret);
    const nc = before.length + ins.length;
    setEdit({ ...e, text: nt, caret: nc, point: null });
    setAc(null);
    setTimeout(() => { const el = e.where === 'bar' ? barInput.current : input.current; if (el) { el.focus(); el.setSelectionRange(nc, nc); } }, 0);
  };
  const hint = useMemo(() => {
    if (!edit || !edit.text.startsWith('=')) return null;
    const call = callAt(edit.text, edit.caret, loc.decimal === ',' ? ';' : ',;');
    if (!call) return null;
    const en = internalFn(call.name, loc);
    const info = FUNC_INFO.get(en);
    if (!info && !FUNCTION_NAMES.includes(en)) return null;
    const args = info ? localArgList(info.args, loc) : [];
    return { name: localFn(en, loc), args, arg: call.arg, desc: info ? info.desc : '' };
  }, [edit && edit.text, edit && edit.caret]); // eslint-disable-line react-hooks/exhaustive-deps

  /* ---------- movement ---------- */
  const moveActive = (dr, dc, inSel = false) => {
    const ax = axesOf(sheet);
    const g = activeRange();
    const multi = inSel && (sel.ranges.length > 1 || g.r1 !== g.r2 || g.c1 !== g.c2);
    if (multi && !(g.r2 >= MAXR - 1 || g.c2 >= MAXC - 1)) { // Enter/Tab move inside the selection
      let { r, c } = sel;
      if (dr) { r += dr; if (r > g.r2) { r = g.r1; c = c + 1 > g.c2 ? g.c1 : c + 1; } else if (r < g.r1) { r = g.r2; c = c - 1 < g.c1 ? g.c2 : c - 1; } } else if (dc) { c += dc; if (c > g.c2) { c = g.c1; r = r + 1 > g.r2 ? g.r1 : r + 1; } else if (c < g.c1) { c = g.c2; r = r - 1 < g.r1 ? g.r2 : r - 1; } }
      setSelState({ ...sel, r, c }); gridApi.current && gridApi.current.scrollTo(r, c);
      return;
    }
    const m = book.mergeAt(sheet, sel.r, sel.c);
    let r = dr > 0 && m ? m.r2 : dr < 0 && m ? m.r1 : sel.r;
    let c = dc > 0 && m ? m.c2 : dc < 0 && m ? m.c1 : sel.c;
    if (dr) for (let i = 0; i < Math.abs(dr); i++) r = ax.rows.next(r, Math.sign(dr));
    if (dc) for (let i = 0; i < Math.abs(dc); i++) c = ax.cols.next(c, Math.sign(dc));
    if (inSel && dc === 0 && dr !== 0 && tabStart.current != null) { c = tabStart.current; tabStart.current = null; }
    go(r, c);
  };
  const edgeOf = (r, c, dr, dc) => {
    const b = book.bounds(sheet);
    const ax = axesOf(sheet);
    const stepR = (x) => ax.rows.next(x, dr); const stepC = (x) => ax.cols.next(x, dc);
    let nr = r; let nc = c;
    const nxt = () => { const a = dr ? stepR(nr) : nr; const b2 = dc ? stepC(nc) : nc; const same = a === nr && b2 === nc; nr = a; nc = b2; return !same; };
    const here = filled(r, c);
    const peekR = dr ? stepR(r) : r; const peekC = dc ? stepC(c) : c;
    const nextFilled = (peekR !== r || peekC !== c) && filled(peekR, peekC);
    if (here && nextFilled) { while (nxt()) { const ar = dr ? stepR(nr) : nr; const ac2 = dc ? stepC(nc) : nc; if ((ar === nr && ac2 === nc) || !filled(ar, ac2)) break; } return { r: nr, c: nc }; }
    // find the next filled cell
    for (let i = 0; i < 1100000; i++) {
      if (!nxt()) break;
      if ((dr > 0 && nr >= b.rows) || (dc > 0 && nc >= b.cols)) return { r: dr > 0 ? MAXR - 1 : nr, c: dc > 0 ? MAXC - 1 : nc };
      if (filled(nr, nc)) return { r: nr, c: nc };
    }
    return { r: dr < 0 ? 0 : dr > 0 ? MAXR - 1 : nr, c: dc < 0 ? 0 : dc > 0 ? MAXC - 1 : nc };
  };
  const extendTo = (r, c) => {
    const a = { r: sel.r, c: sel.c };
    const g = expandMerges(norm(a, { r, c }));
    setSel({ r: sel.r, c: sel.c, ranges: [...sel.ranges.slice(0, -1), g], anchor: a, end: { r, c } });
  };
  const extendBy = (dr, dc, ctrl) => {
    const end = sel.end || { r: sel.r, c: sel.c };
    let p;
    if (ctrl) p = edgeOf(end.r, end.c, dr, dc);
    else { const ax = axesOf(sheet); p = { r: dr ? ax.rows.next(end.r, dr) : end.r, c: dc ? ax.cols.next(end.c, dc) : end.c }; }
    extendTo(p.r, p.c);
  };

  /* ---------- sheets ---------- */
  const switchSheet = (i) => {
    if (i < 0 || i >= book.sheets.length) return;
    if (sheet) sheet.sel = { r: sel.r, c: sel.c };
    book.active = i;
    const sh = book.sheets[i];
    if (!edit || !formulaText) setSelState(sh.sel ? cellSel(sh.sel.r, sh.sel.c) : cellSel(0, 0));
    else { setEdit((x) => x && { ...x, where: 'bar', mode: 'enter' }); setTimeout(() => { const el = barInput.current; if (el) { el.focus(); const n = el.value.length; el.setSelectionRange(n, n); } }, 0); }
    setChartSel(null);
    setVer((v) => v + 1);
  };

  /* ---------- clipboard ---------- */
  const doCopy = (cut = false, ev = null) => {
    if (sel.ranges.length > 1) { toast(t('This doesn’t work with several selections.')); return; }
    const g = clipRange(activeRange());
    if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > 2000000) { toast(t('The selection is too large to copy.')); return; }
    const blk = book.copyBlock(sheet, g);
    const id = `${clipId.current}-${Date.now()}`;
    const tsv = toTSV(book, sheet, g);
    const html = toHTML(book, sheet, g, id);
    setCopy({ blk, cut, g, sheetId: sheet.id, id, tsv });
    if (ev && ev.clipboardData) { ev.clipboardData.setData('text/plain', tsv); ev.clipboardData.setData('text/html', html); ev.preventDefault(); return; }
    try {
      if (window.ClipboardItem && navigator.clipboard && navigator.clipboard.write) navigator.clipboard.write([new window.ClipboardItem({ 'text/plain': new Blob([tsv], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }) })]).catch(() => navigator.clipboard.writeText(tsv).catch(() => {}));
      else if (navigator.clipboard) navigator.clipboard.writeText(tsv).catch(() => {});
    } catch { /* the internal clipboard still works */ }
  };
  const pasteData = (html, text, mode = {}) => {
    if (!guard()) return;
    const target = activeRange();
    const parsed = html ? parseHTML(html) : null;
    let res;
    if (copy && ((parsed && parsed.clipId === copy.id) || (!parsed && text && text === copy.tsv) || (!html && !text))) {
      res = book.paste(sheet, target, copy.blk, { ...mode, cut: copy.cut && !mode.what });
      if (copy.cut) setCopy(null);
    } else if (parsed && parsed.rows.length) {
      const grid = parsed.rows.map((r) => r.map((x) => (x ? x.text : '')));
      res = book.pasteText(sheet, target.r1, target.c1, grid, { transpose: mode.transpose });
      if (mode.what !== 'values') book.run('Paste formats', () => { parsed.rows.forEach((row, i) => row.forEach((x, j) => { if (!x || !Object.keys(x.style).length) return; const r = target.r1 + (mode.transpose ? j : i); const c = target.c1 + (mode.transpose ? i : j); const cell = sheet.cells.get(K(r, c)); book.setCell(sheet, r, c, { ...(cell || { v: null }), s: book.styles.merge(cell && cell.s, x.style) }); })); });
    } else if (text) {
      const grid = parseTSV(text);
      if (!grid.length) return;
      if (grid.length === 1 && grid[0].length === 1 && (target.r1 !== target.r2 || target.c1 !== target.c2)) { book.run('Paste', () => { for (let r = target.r1; r <= Math.min(target.r2, target.r1 + 10000); r++) for (let c = target.c1; c <= Math.min(target.c2, target.c1 + 500); c++) book.input(sheet, r, c, grid[0][0]); }); res = target; } else res = book.pasteText(sheet, target.r1, target.c1, grid, { transpose: mode.transpose });
    }
    if (res) setSel({ r: res.r1, c: res.c1, ranges: [res], anchor: { r: res.r1, c: res.c1 }, end: { r: res.r2, c: res.c2 } });
  };
  const pasteFromSystem = async (mode = {}) => {
    try {
      if (navigator.clipboard && navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        let html = null; let text = null;
        for (const it of items) { if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text(); if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text(); }
        pasteData(html, text, mode); return;
      }
      if (navigator.clipboard && navigator.clipboard.readText) { pasteData(null, await navigator.clipboard.readText(), mode); return; }
    } catch { /* permission denied */ }
    if (copy) pasteData(null, null, mode); else toast(t('Press Ctrl+V to paste from the clipboard.'));
  };
  const applyPainter = (n) => {
    const p = painter; if (!p) return;
    const g = n.ranges[n.ranges.length - 1];
    const blk = p.blk;
    book.paste(sheet, { r1: g.r1, c1: g.c1, r2: Math.max(g.r2, g.r1 + blk.rows.length - 1), c2: Math.max(g.c2, g.c1 + blk.rows[0].length - 1) }, blk, { what: 'formats' });
    if (!p.keep) setPainter(null);
  };

  /* ---------- formatting helpers ---------- */
  const curCell = sheet && sheet.cells.get(K(sel.r, sel.c));
  const curStyle = (curCell && curCell.s) || null;
  const fmt = (patch, label) => { if (!guard()) return; book.format(sheet, sel.ranges.map(clipRangeForFormat), patch, label); };
  const clipRangeForFormat = (g) => g;
  const toggle = (k) => { const on = !(curStyle && curStyle[k]); fmt({ [k]: on || undefined }); };

  /* ---------- the command dispatcher ---------- */
  const run = async (cmd, arg) => {
    if (!book) return;
    const MUT = !['copy', 'find', 'replace', 'goto', 'gotoSpecial', 'selectSpecial', 'print', 'printDo', 'zoom', 'zoomSet', 'zoomSel', 'gridlinesView', 'showFormulas', 'formulaBar', 'compact', 'macros', 'stats', 'names', 'commandSearch', 'filterMenu', 'calcNow', 'freezeView', 'showNotes', 'csvOptions', 'filterToggleView'].includes(cmd);
    if (MUT && readOnly && !['filterToggle', 'filterClear', 'filterReapply', 'gridlines', 'headings', 'freeze'].includes(cmd)) { guard(); return; }
    if (edit && MUT && !['insertFn', 'autosum'].includes(cmd)) { if (!(await commit('stay'))) return; }
    const g = activeRange();
    switch (cmd) {
      case 'undo': { if (edit) { cancelEdit(); return; } const tx = book.undo(); if (tx && tx.meta && tx.meta.sel) setSel(tx.meta.sel, { scroll: true }); break; }
      case 'redo': book.redo(); break;
      case 'copy': doCopy(false); break;
      case 'cut': doCopy(true); break;
      case 'paste': await pasteFromSystem(); break;
      case 'pasteSpecial': if (copy) pasteData(null, null, arg); else await pasteFromSystem(arg); break;
      case 'pasteSpecialDlg': setDialog({ kind: 'pasteSpecial' }); break;
      case 'painter': if (painter) setPainter(null); else setPainter({ blk: book.copyBlock(sheet, clipRange(g)), keep: !!arg }); break;
      case 'font': fmt({ font: arg === defFont.font ? undefined : arg }); break;
      case 'size': fmt({ size: arg === defFont.size ? undefined : arg }); break;
      case 'grow': { const SZ = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]; fmt((s) => { const cur = (s && s.size) || defFont.size; const i = SZ.findIndex((x) => x >= cur); const n = arg > 0 ? SZ[Math.min(SZ.length - 1, SZ[i] === cur ? i + 1 : i)] : SZ[Math.max(0, i - 1)]; return { size: n === defFont.size ? undefined : n }; }); break; }
      case 'bold': toggle('b'); break;
      case 'italic': toggle('i'); break;
      case 'underline': toggle('u'); break;
      case 'strike': toggle('st'); break;
      case 'color': fmt({ color: arg || undefined }, 'Font colour'); break;
      case 'fill': fmt({ fill: arg || undefined }, 'Fill colour'); break;
      case 'pickFill': pickColor(arg, curStyle && curStyle.fill, (c) => { setLast((l) => ({ ...l, fill: c })); run('fill', c); }); break;
      case 'pickColor': pickColor(arg, curStyle && curStyle.color, (c) => { setLast((l) => ({ ...l, color: c })); run('color', c); }, { noneLabel: 'Automatic' }); break;
      case 'border': setLast((l) => ({ ...l, border: arg })); for (const x of sel.ranges) book.borders(sheet, clipRange(x), arg, last.line); break;
      case 'borderColor': pickColor({ left: window.innerWidth / 2 - 120, below: 160 }, last.line.c, (c) => setLast((l) => ({ ...l, line: { ...l.line, c: c || '#000000' } })), { allowNone: false }); break;
      case 'borderStyle': setLast((l) => ({ ...l, line: { ...l.line, s: arg } })); break;
      case 'align': fmt({ h: curStyle && curStyle.h === arg ? undefined : arg }); break;
      case 'valign': fmt({ va: arg === 'bottom' ? undefined : arg }); break;
      case 'wrap': toggle('wrap'); break;
      case 'indent': fmt((s) => { const n = Math.max(0, Math.min(15, ((s && s.indent) || 0) + arg)); return { indent: n || undefined, h: n && !(s && s.h) ? 'left' : s && s.h }; }); break;
      case 'rotate': fmt({ rot: arg === 0 ? undefined : arg < 0 ? 90 - arg : arg }); break;
      case 'merge': for (const x of sel.ranges) book.merge(sheet, x, { center: arg === 'center', across: arg === 'across' }); break;
      case 'unmerge': for (const x of sel.ranges) book.unmerge(sheet, x); break;
      case 'numPreset': { const p = FORMAT_PRESETS(loc).find((x) => x.id === arg); if (p) fmt({ nf: p.code === 'General' ? undefined : p.code }, 'Number format'); break; }
      case 'currency': fmt({ nf: arg === 'acc' ? ACCOUNTING(loc) : makeCurrencyFormat({ symbol: arg, decimals: 2, after: loc.decimal === ',' || arg === 'CHF' }) }, 'Number format'); break;
      case 'comma': fmt({ nf: '#,##0.00' }, 'Number format'); break;
      case 'decimals': fmt((s, r, c) => ({ nf: changeDecimals((s && s.nf) || 'General', arg, book.value(sheet, r, c)) }), arg > 0 ? 'Increase decimal' : 'Decrease decimal'); break;
      case 'cellStyle': { const cs = CELL_STYLES.find((x) => x.id === arg); if (!cs) break; fmt((s) => (cs.s ? { ...cs.s } : { font: undefined, size: undefined, b: undefined, i: undefined, u: undefined, st: undefined, color: undefined, fill: undefined, bd: undefined, nf: s && s.nf }), t(cs.label)); break; }
      case 'formatCells': setDialog({ kind: 'formatCells', tab: arg || 'number' }); break;
      case 'clear': for (const x of sel.ranges) book.clear(sheet, [clipRange(x)], arg); break;
      case 'insertRows': { const x = g.r1; const n = g.r2 >= MAXR - 1 ? 1 : g.r2 - g.r1 + 1; book.insert(sheet, 'row', x, n); break; }
      case 'insertCols': { const x = g.c1; const n = g.c2 >= MAXC - 1 ? 1 : g.c2 - g.c1 + 1; book.insert(sheet, 'col', x, n); break; }
      case 'deleteRows': { const n = g.r2 >= MAXR - 1 ? 1 : g.r2 - g.r1 + 1; book.remove(sheet, 'row', g.r1, n); break; }
      case 'deleteCols': { const n = g.c2 >= MAXC - 1 ? 1 : g.c2 - g.c1 + 1; book.remove(sheet, 'col', g.c1, n); break; }
      case 'insertCellsDlg': if (g.r1 === 0 && g.r2 >= MAXR - 1) run('insertCols'); else if (g.c1 === 0 && g.c2 >= MAXC - 1) run('insertRows'); else setDialog({ kind: 'cells', insert: true }); break;
      case 'deleteCellsDlg': if (g.r1 === 0 && g.r2 >= MAXR - 1) run('deleteCols'); else if (g.c1 === 0 && g.c2 >= MAXC - 1) run('deleteRows'); else setDialog({ kind: 'cells', insert: false }); break;
      case 'cells': if (arg === 'row') run(dialog && dialog.insert ? 'insertRows' : 'deleteRows'); else if (arg === 'col') run(dialog && dialog.insert ? 'insertCols' : 'deleteCols'); else book.shiftCells(sheet, g, arg, arg === 'down' || arg === 'right'); break;
      case 'insertSheet': { book.addSheet(lang === 'de' ? book.nextSheetName('Tabelle') : book.nextSheetName('Sheet'), book.active + 1); switchSheet(book.active); break; }
      case 'deleteSheet': { if (book.sheets.filter((s) => !s.hidden).length <= 1) { toast(t('A workbook needs at least one visible sheet.')); break; } const used = book.usedRange(sheet); if (used) { const ok = await confirmDialog({ title: t('Delete sheet “{name}”?', { name: sheet.name }), text: t('The sheet and everything on it is removed. You can undo this with Ctrl+Z.'), confirm: t('Delete'), danger: true }); if (!ok) break; } book.deleteSheet(sheet); switchSheet(book.active); break; }
      case 'renameSheet': { const name = await promptDialog({ title: t('Rename sheet'), label: t('Name'), value: sheet.name, confirm: t('Rename'), validate: (v) => (!v.trim() ? t('Enter a name') : /[\\/?*[\]:]/.test(v) ? t('These characters are not allowed: \\ / ? * [ ] :') : v.length > 31 ? t('At most 31 characters') : book.sheetByName(v.trim()) && book.sheetByName(v.trim()) !== sheet ? t('This name is already used') : null) }); if (name) book.renameSheet(sheet, name); break; }
      case 'duplicateSheet': book.duplicateSheet(sheet); switchSheet(book.active); break;
      case 'tabColor': pickColor(arg || { left: 200, below: window.innerHeight - 330 }, sheet.color, (c) => book.setSheetProps(sheet, { color: c }, 'Tab colour')); break;
      case 'hideSheet': if (book.sheets.filter((s) => !s.hidden).length <= 1) { toast(t('A workbook needs at least one visible sheet.')); break; } book.setSheetProps(sheet, { hidden: 'hidden' }, 'Hide sheet'); switchSheet(book.sheets.findIndex((s) => !s.hidden)); break;
      case 'unhideSheet': setDialog({ kind: 'unhideSheet' }); break;
      case 'rowHeight': case 'colWidth': case 'defaultWidth': {
        const isRow = cmd === 'rowHeight';
        const cur = cmd === 'defaultWidth' ? sheet.defW : isRow ? axesOf(sheet).rows.size(sel.r) : axesOf(sheet).cols.size(sel.c);
        const v = await promptDialog({ title: cmd === 'defaultWidth' ? t('Default column width') : isRow ? t('Row height') : t('Column width'), label: t('Pixels'), value: String(cur), confirm: t('OK'), validate: (x) => (Number(x) > 0 && Number(x) < 2000 ? null : t('Enter a number between 1 and 2000')) });
        if (v == null) break;
        if (cmd === 'defaultWidth') { book.setSheetProps(sheet, { defW: Math.round(+v) }, 'Default width'); break; }
        const idx = []; for (const x of sel.ranges) for (let i = isRow ? x.r1 : x.c1; i <= Math.min(isRow ? x.r2 : x.c2, (isRow ? x.r1 : x.c1) + 20000); i++) idx.push(i);
        book.setSizes(sheet, isRow ? 'row' : 'col', idx, +v);
        break;
      }
      case 'autofitRows': case 'autofitCols': { const isRow = cmd === 'autofitRows'; const idx = []; for (const x of sel.ranges.map(clipRange)) for (let i = isRow ? x.r1 : x.c1; i <= (isRow ? x.r2 : x.c2) && idx.length < 2000; i++) idx.push(i); gridApi.current.autoFit(isRow ? 'row' : 'col', idx); break; }
      case 'hide': case 'unhide': {
        const axis = arg; const idx = [];
        for (const x of sel.ranges) { const a = axis === 'row' ? x.r1 : x.c1; const b = axis === 'row' ? (x.r2 >= MAXR - 1 && cmd === 'unhide' ? book.bounds(sheet).rows + 100 : x.r2) : x.c2 >= MAXC - 1 && cmd === 'unhide' ? book.bounds(sheet).cols + 50 : x.c2; for (let i = Math.max(0, cmd === 'unhide' ? a - 1 : a); i <= Math.min(b + (cmd === 'unhide' ? 1 : 0), a + 20000); i++) idx.push(i); }
        book.setHidden(sheet, axis, idx, cmd === 'hide');
        break;
      }
      case 'autosum': autosum(arg || 'SUM'); break;
      case 'fillDir': fillDir(arg); break;
      case 'sort': {
        const dg = dataRange();
        const header = guessHeader(dg);
        const col = Math.max(dg.c1, Math.min(dg.c2, sel.c));
        book.sort(sheet, dg, [{ col, desc: arg === 'desc' }], { header });
        break;
      }
      case 'sortDlg': { const dg = dataRange(); setDialog({ kind: 'sort', g: dg, header: guessHeader(dg) }); break; }
      case 'filterToggle': {
        if (sheet.filter) { book.setFilter(sheet, null); break; }
        const dg = dataRange();
        book.setFilter(sheet, dg);
        break;
      }
      case 'filterClear': if (sheet.filter) book.run('Clear filter', () => { book.touch(sheet); sheet.filter = { ...sheet.filter, cols: {} }; sheet.filterHidden = new Set(); }); break;
      case 'filterReapply': if (sheet.filter) book.run('Reapply', () => { book.touch(sheet); book.applyFilter(sheet); }); break;
      case 'find': setDialog({ kind: 'find' }); break;
      case 'replace': setDialog({ kind: 'find', replace: true }); break;
      case 'goto': setDialog({ kind: 'goto' }); break;
      case 'gotoSpecial': setDialog({ kind: 'goto', special: true }); break;
      case 'selectSpecial': selectSpecial(arg); break;
      case 'chart': insertChart(arg); break;
      case 'insertFunction': setDialog({ kind: 'insertFn' }); break;
      case 'insertFn': { const name = localFn(arg, loc); if (edit && edit.text.startsWith('=')) { const nt = `${edit.text.slice(0, edit.caret)}${name}(${edit.text.slice(edit.caret)}`; setEdit({ ...edit, text: nt, caret: edit.caret + name.length + 1 }); } else startEdit(`=${name}(`, 'edit'); break; }
      case 'link': setDialog({ kind: 'link' }); break;
      case 'note': setDialog({ kind: 'note' }); break;
      case 'showNotes': setShowNotes((x) => !x); break;
      case 'print': setDialog({ kind: 'print' }); break;
      case 'orientation': book.setSheetProps(sheet, { print: { ...sheet.print, orientation: arg } }, 'Orientation'); break;
      case 'fit': book.setSheetProps(sheet, { print: { ...sheet.print, fit: arg } }, 'Scaling'); break;
      case 'printArea': book.setSheetProps(sheet, { print: { ...sheet.print, area: arg === 'set' ? rangeA1(clipRange(g)) : null } }, 'Print area'); break;
      case 'printGrid': book.setSheetProps(sheet, { print: { ...sheet.print, gridlines: !sheet.print.gridlines } }, 'Print gridlines'); break;
      case 'printHead': book.setSheetProps(sheet, { print: { ...sheet.print, headings: !sheet.print.headings } }, 'Print headings'); break;
      case 'names': setDialog({ kind: 'names' }); break;
      case 'defineName': setDialog({ kind: 'names', startNew: true }); break;
      case 'showFormulas': setShowFormulas((x) => !x); break;
      case 'calcMode': book.autoCalc = !!arg; if (arg) book.calc.recalc(null, { force: true }); book.emit({ viewOnly: false }); break;
      case 'calcNow': book.calc.rebuild(); book.calc.recalc(null, { force: true }); book.emit({ viewOnly: true }); break;
      case 'textToColumns': setDialog({ kind: 'ttc', g: clipRange({ ...g, c2: g.c1 }) }); break;
      case 'removeDuplicates': setDialog({ kind: 'dups', g: dataRange() }); break;
      case 'validation': setDialog({ kind: 'dv' }); break;
      case 'group': for (const x of sel.ranges) { if (x.c1 === 0 && x.c2 >= MAXC - 1) book.setOutline(sheet, 'row', x, arg); else if (x.r1 === 0 && x.r2 >= MAXR - 1) book.setOutline(sheet, 'col', x, arg); else book.setOutline(sheet, 'row', clipRange(x), arg); } break;
      case 'detail': { // collapse/expand grouped rows around the selection
        const rows = []; const lvl = sheet.outlineRows;
        let a = g.r1; let b = g.r2 >= MAXR - 1 ? book.bounds(sheet).rows : g.r2;
        while (a > 0 && lvl.get(a - 1)) a--; while (lvl.get(b + 1)) b++;
        for (let i = a; i <= b; i++) if (lvl.get(i)) rows.push(i);
        if (!rows.length) { toast(t('Group rows first (Data → Group).')); break; }
        book.setHidden(sheet, 'row', rows, !arg);
        break;
      }
      case 'csvOptions': setDialog({ kind: 'csv', bytes, reimport: true }); break;
      case 'importCsv': pickFile('.csv,.tsv,.txt,.tab', (b2) => setDialog({ kind: 'csv', bytes: b2 })); break;
      case 'freeze': {
        const fr = arg === 'none' ? { rows: 0, cols: 0 } : arg === 'row' ? { rows: 1, cols: 0 } : arg === 'col' ? { rows: 0, cols: 1 } : { rows: sel.r, cols: sel.c };
        book.setSheetProps(sheet, { freeze: fr }, 'Freeze panes');
        break;
      }
      case 'gridlines': book.setSheetProps(sheet, { gridlines: sheet.gridlines === false }, 'Gridlines'); break;
      case 'headings': book.setSheetProps(sheet, { headings: sheet.headings === false }, 'Headings'); break;
      case 'formulaBar': setFormulaBar((x) => !x); break;
      case 'compact': setCompact((x) => !x); break;
      case 'zoom': setZoom(zoom + arg); break;
      case 'zoomSet': setZoom(arg); break;
      case 'zoomSel': { const api = gridApi.current; const geo = api && api.geo(); if (!geo) break; const cg = clipRange(g); const w = axesOf(sheet).cols.pos(cg.c2 + 1) - axesOf(sheet).cols.pos(cg.c1); const h = axesOf(sheet).rows.pos(cg.r2 + 1) - axesOf(sheet).rows.pos(cg.r1); setZoom(Math.floor(Math.min((geo.W - geo.hw - 20) / w, (geo.H - geo.hh - 20) / h) * 100)); break; }
      case 'cfQuick': cfQuick(arg); break;
      case 'cfBar': addCF({ type: 'bar', color: arg }); break;
      case 'cfScale': addCF({ type: 'scale', colors: arg }); break;
      case 'cfNew': setDialog({ kind: 'cfRule', rule: { type: 'cell', op: 'gt' } }); break;
      case 'cfManage': setDialog({ kind: 'cfManage' }); break;
      case 'cfClear': book.setSheetProps(sheet, { cf: arg === 'sheet' ? [] : sheet.cf.map((r) => ({ ...r, ranges: r.ranges.filter((x) => !sel.ranges.some((s) => s.r1 <= x.r1 && s.r2 >= x.r2 && s.c1 <= x.c1 && s.c2 >= x.c2)) })).filter((r) => r.ranges.length) }, 'Clear rules'); break;
      case 'macros': setDialog({ kind: 'macros' }); break;
      case 'stats': setDialog({ kind: 'stats' }); break;
      case 'commandSearch': setDialog({ kind: 'cmdSearch' }); break;
      default: break;
    }
    if (!edit) focusGrid();
  };
  const setZoom = (z) => { const v = Math.max(10, Math.min(400, Math.round(z))); setZoomMap((m) => ({ ...m, [sheet.id]: v })); };
  const guessHeader = (dg) => {
    if (dg.r2 <= dg.r1) return false;
    let txt = 0; let other = 0; let bold = false;
    for (let c = dg.c1; c <= dg.c2; c++) { const v = book.value(sheet, dg.r1, c); const v2 = book.value(sheet, dg.r1 + 1, c); if (typeof v === 'string') txt++; if (typeof v2 !== 'string' && v2 != null) other++; const s = book.styleOf(sheet, dg.r1, c); if (s && (s.b || s.fill)) bold = true; }
    return txt === dg.c2 - dg.c1 + 1 && (other > 0 || bold);
  };
  const pickFile = (accept, cb) => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.onchange = async () => { const f = i.files && i.files[0]; if (f) cb(new Uint8Array(await f.arrayBuffer())); }; i.click(); };
  const autosum = (fn) => {
    const g = activeRange();
    const sep = loc.decimal === ',' ? ';' : ',';
    void sep;
    if (g.r1 === g.r2 && g.c1 === g.c2) {
      const r = sel.r; const c = sel.c;
      let r0 = r - 1; while (r0 >= 0 && typeof book.value(sheet, r0, c) === 'number') r0--;
      let range = null;
      if (r0 < r - 1) range = `${addr(r0 + 1, c)}:${addr(r - 1, c)}`;
      else { let c0 = c - 1; while (c0 >= 0 && typeof book.value(sheet, r, c0) === 'number') c0--; if (c0 < c - 1) range = `${addr(r, c0 + 1)}:${addr(r, c - 1)}`; }
      startEdit(`=${fn}(${range || ''})`, 'edit');
      return;
    }
    // a block: totals below each column (or to the right of a single row)
    const cg = clipRange(g);
    book.run('AutoSum', () => {
      if (cg.r1 === cg.r2) { book.input(sheet, cg.r1, cg.c2 + 1, `=${fn}(${addr(cg.r1, cg.c1)}:${addr(cg.r1, cg.c2)})`); return; }
      for (let c = cg.c1; c <= cg.c2; c++) book.setCell(sheet, cg.r2 + 1, c, { ...(sheet.cells.get(K(cg.r2 + 1, c)) || {}), v: null, f: `${fn}(${addr(cg.r1, c)}:${addr(cg.r2, c)})` });
    });
  };
  const fillDir = (dir) => {
    const g = clipRange(activeRange());
    if (dir === 'down') { if (g.r1 === g.r2) { if (g.r1 === 0) return; book.fill(sheet, { ...g, r1: g.r1 - 1, r2: g.r1 - 1 }, g, { copy: true }); } else book.fill(sheet, { ...g, r2: g.r1 }, g, { copy: true }); }
    if (dir === 'right') { if (g.c1 === g.c2) { if (g.c1 === 0) return; book.fill(sheet, { ...g, c1: g.c1 - 1, c2: g.c1 - 1 }, g, { copy: true }); } else book.fill(sheet, { ...g, c2: g.c1 }, g, { copy: true }); }
    if (dir === 'up' && g.r2 > g.r1) book.fill(sheet, { ...g, r1: g.r2 }, g, { copy: true });
    if (dir === 'left' && g.c2 > g.c1) book.fill(sheet, { ...g, c1: g.c2 }, g, { copy: true });
  };
  const selectSpecial = (kind) => {
    const keys = [];
    const u = book.usedRange(sheet) || { r1: 0, c1: 0, r2: 0, c2: 0 };
    const g0 = activeRange();
    const g = g0.r1 === g0.r2 && g0.c1 === g0.c2 ? u : clipRange(g0);
    const inG = (r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;
    if (kind === 'blanks') { for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) for (let c = g.c1; c <= g.c2; c++) if (!filled(r, c)) keys.push([r, c]); } else if (kind === 'validation' || kind === 'cf') {
      const list = kind === 'validation' ? sheet.dv : sheet.cf;
      const rs = list.flatMap((x) => x.ranges);
      if (rs.length) { setSel({ r: rs[0].r1, c: rs[0].c1, ranges: rs.map((x) => ({ ...x })), anchor: { r: rs[0].r1, c: rs[0].c1 } }); return; }
    } else {
      for (const [k, cell] of sheet.cells) {
        const r = Math.floor(k / MAXC); const c = k % MAXC;
        if (!inG(r, c)) continue;
        if ((kind === 'formulas' && cell.f) || (kind === 'constants' && !cell.f && cell.v != null) || (kind === 'notes' && cell.note) || (kind === 'errors' && cell.f && isErr(cell.v))) keys.push([r, c]);
      }
    }
    if (!keys.length) { toast(t('No cells were found.')); return; }
    keys.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    const ranges = keys.slice(0, 5000).map(([r, c]) => ({ r1: r, c1: c, r2: r, c2: c }));
    setSel({ r: keys[0][0], c: keys[0][1], ranges, anchor: { r: keys[0][0], c: keys[0][1] } });
  };
  const insertChart = (type) => {
    if (!guard()) return;
    const dg = dataRange();
    if (!filled(dg.r1, dg.c1) && !filled(dg.r2, dg.c2) && dg.r1 === dg.r2) { toast(t('Select the data for the chart first.')); return; }
    const ax = axesOf(sheet);
    const c = Math.min(MAXC - 1, dg.c2 + 2);
    const ch = { id: `c${Date.now().toString(36)}`, type, stacked: false, title: '', legend: type === 'pie' || type === 'doughnut' ? 'right' : 'bottom', xTitle: '', yTitle: '', range: rangeText({ sheet: sheet.name, ...dg, ar1: true, ac1: true, ar2: true, ac2: true, kind: 'range' }), byRows: false, series: null, anchor: { r: dg.r1, c, dx: 0, dy: 0, w: 480, h: 288 } };
    // a nice default title: the header above a single series
    if (dg.c2 - dg.c1 === 1 && typeof book.value(sheet, dg.r1, dg.c2) === 'string') ch.title = book.value(sheet, dg.r1, dg.c2);
    void ax;
    book.setSheetProps(sheet, { charts: [...sheet.charts, ch] }, 'Insert chart');
    setChartSel(ch.id);
  };
  const normAnchor = (a) => {
    const ax = axesOf(sheet);
    const x = ax.cols.pos(a.c) + (a.dx || 0); const y = ax.rows.pos(a.r) + (a.dy || 0);
    const c = ax.cols.index(Math.max(0, x)); const r = ax.rows.index(Math.max(0, y));
    return { ...a, c, r, dx: Math.max(0, Math.round(x - ax.cols.pos(c))), dy: Math.max(0, Math.round(y - ax.rows.pos(r))), _px: undefined };
  };
  const updateChart = (id, patch) => book.setSheetProps(sheet, { charts: sheet.charts.map((c) => (c.id === id ? { ...c, ...patch } : c)) }, 'Chart');
  const deleteChart = (id) => { book.setSheetProps(sheet, { charts: sheet.charts.filter((c) => c.id !== id) }, 'Delete chart'); setChartSel(null); };
  const addCF = (rule) => { if (!guard()) return; const ranges = sel.ranges.map(clipRange); book.setSheetProps(sheet, { cf: [...sheet.cf, { id: `cf${Date.now().toString(36)}`, priority: 0, ...rule, ranges }].map((x, i) => ({ ...x, priority: i + 1 })) }, 'Conditional formatting'); };
  const cfQuick = (kind) => {
    const map = { gt: { type: 'cell', op: 'gt' }, lt: { type: 'cell', op: 'lt' }, between: { type: 'cell', op: 'between' }, eq: { type: 'cell', op: 'eq' }, text: { type: 'text', op: 'contains' }, dup: { type: 'dup' }, top: { type: 'top', rank: 10 }, topPct: { type: 'top', rank: 10, percent: true }, bottom: { type: 'bottom', rank: 10 }, above: { type: 'above' }, below: { type: 'below' }, date: { type: 'formula', formula: `AND(${addr(activeRange().r1, activeRange().c1)}>=TODAY()-7,${addr(activeRange().r1, activeRange().c1)}<=TODAY())` } };
    const rule = map[kind];
    if (kind === 'dup' || kind === 'above' || kind === 'below') { addCF({ ...rule, style: CF_STYLES[0].style }); return; }
    setDialog({ kind: 'cfRule', rule: { ...rule, style: CF_STYLES[0].style } });
  };

  /* ---------- keyboard ---------- */
  const onKey = (e) => { onKeyInner(e); if (e.defaultPrevented) e.stopPropagation(); };
  const onKeyInner = (e) => {
    if (!book) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    const lower = k.length === 1 ? k.toLowerCase() : k;
    // autocomplete navigation
    if (edit && ac) {
      if (k === 'ArrowDown' || k === 'ArrowUp') { e.preventDefault(); setAc({ ...ac, i: (ac.i + (k === 'ArrowDown' ? 1 : -1) + ac.items.length) % ac.items.length }); return; }
      if (k === 'Tab' || (k === 'Enter' && !e.altKey && !ctrl)) { e.preventDefault(); acceptAc(ac.items[ac.i]); return; }
      if (k === 'Escape') { e.preventDefault(); setAc(null); return; }
    }
    if (edit) {
      if (k === 'Enter' && e.altKey) { e.preventDefault(); const el = e.target; const s = el.selectionStart; const nt = `${edit.text.slice(0, s)}\n${edit.text.slice(el.selectionEnd)}`; setEdit({ ...edit, text: nt, caret: s + 1, mode: 'edit' }); setTimeout(() => el.setSelectionRange(s + 1, s + 1), 0); return; }
      if (k === 'Enter') { e.preventDefault(); commit(ctrl ? 'stay' : { dr: e.shiftKey ? -1 : 1, dc: 0 }, { ctrlEnter: ctrl }); return; }
      if (k === 'Tab') { e.preventDefault(); if (tabStart.current == null) tabStart.current = edit.c; commit({ dr: 0, dc: e.shiftKey ? -1 : 1 }); return; }
      if (k === 'Escape') { e.preventDefault(); cancelEdit(); return; }
      if (k === 'F2') { e.preventDefault(); setEdit({ ...edit, mode: edit.mode === 'edit' ? 'enter' : 'edit', point: null }); return; }
      if (k === 'F4' && edit.text.startsWith('=')) { e.preventDefault(); cycleAbs(e.target); return; }
      if (ctrl && (lower === ';' || lower === '.' || lower === ':' || lower === ',')) { e.preventDefault(); const time = e.shiftKey || lower === ':'; const v = formatValue(time ? nowSerial() : todaySerial(), time ? loc.time : loc.shortDate, loc).text; const el = e.target; const s = el.selectionStart; const nt = edit.text.slice(0, s) + v + edit.text.slice(el.selectionEnd); setEdit({ ...edit, text: nt, caret: s + v.length }); return; }
      const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (arrows[k] && edit.mode === 'enter') {
        if (pointable(edit)) { // point mode: arrows pick references
          e.preventDefault();
          const [dr, dc] = arrows[k];
          const ax = axesOf(sheet);
          const pc = edit.pc || { r: edit.r, c: edit.c, a: null };
          const nr = dr ? ax.rows.next(pc.r, dr) : pc.r; const nc = dc ? ax.cols.next(pc.c, dc) : pc.c;
          const a = e.shiftKey ? pc.a || { r: pc.r, c: pc.c } : { r: nr, c: nc };
          insertRef(norm(a, { r: nr, c: nc }), 'move');
          setEdit((x) => x && { ...x, pc: { r: nr, c: nc, a } });
          gridApi.current && gridApi.current.scrollTo(nr, nc);
          return;
        }
        if (!edit.text.startsWith('=') || !edit.point) { e.preventDefault(); commit({ dr: arrows[k][0], dc: arrows[k][1] }); return; }
      }
      if (ctrl && lower === 'a' && edit.text.startsWith('=')) return; // select text inside the editor
      if (ctrl && lower === 'z') { return; }
      if (ctrl && lower === 's') { e.preventDefault(); commit('stay').then((ok) => ok && host.save && host.save()); return; }
      return; // typing
    }
    // ---- not editing ----
    if (chartSel) {
      if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); if (guard()) deleteChart(chartSel); return; }
      if (k === 'Escape') { setChartSel(null); return; }
    }
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) {
      e.preventDefault();
      const [dr, dc] = arrows[k];
      if (e.shiftKey) { extendBy(dr, dc, ctrl); return; }
      if (ctrl) { const p = edgeOf(sel.r, sel.c, dr, dc); go(p.r, p.c); return; }
      tabStart.current = null;
      moveActive(dr, dc);
      return;
    }
    switch (true) {
      case k === 'Enter': e.preventDefault(); moveActive(e.shiftKey ? -1 : 1, 0, true); return;
      case k === 'Tab': e.preventDefault(); if (!e.shiftKey && tabStart.current == null) tabStart.current = sel.c; moveActive(0, e.shiftKey ? -1 : 1, true); return;
      case k === 'Home' && ctrl: e.preventDefault(); if (e.shiftKey) extendTo(sheet.freeze.rows || 0, sheet.freeze.cols || 0); else go(sheet.freeze.rows || 0, sheet.freeze.cols || 0); return;
      case k === 'End' && ctrl: { e.preventDefault(); const u = book.dataRange(sheet) || { r2: 0, c2: 0 }; if (e.shiftKey) extendTo(u.r2, u.c2); else go(u.r2, u.c2); return; }
      case k === 'Home': e.preventDefault(); if (e.shiftKey) extendTo(sel.r, 0); else go(sel.r, 0); return;
      case k === 'End': { e.preventDefault(); let c = 0; for (const key of sheet.cells.keys()) if (Math.floor(key / MAXC) === sel.r && sheet.cells.get(key).v != null) c = Math.max(c, key % MAXC); if (e.shiftKey) extendTo(sel.r, c); else go(sel.r, c); return; }
      case k === 'PageDown' || k === 'PageUp': {
        e.preventDefault();
        if (ctrl) { let i = book.active; const d = k === 'PageDown' ? 1 : -1; do { i += d; } while (i >= 0 && i < book.sheets.length && book.sheets[i].hidden); if (i >= 0 && i < book.sheets.length) switchSheet(i); return; }
        const n = gridApi.current ? gridApi.current.pageRows() : 20;
        if (e.altKey) go(sel.r, Math.max(0, sel.c + (k === 'PageDown' ? 10 : -10)));
        else if (e.shiftKey) extendTo(Math.max(0, Math.min(MAXR - 1, (sel.end || sel).r + (k === 'PageDown' ? n : -n))), (sel.end || sel).c);
        else go(Math.max(0, sel.r + (k === 'PageDown' ? n : -n)), sel.c);
        return;
      }
      case k === 'F2' && e.shiftKey: e.preventDefault(); run('note'); return;
      case k === 'F2': e.preventDefault(); startEdit(undefined, 'edit'); return;
      case k === 'F9': e.preventDefault(); run('calcNow'); return;
      case k === 'Delete': e.preventDefault(); if (guard()) for (const x of sel.ranges) book.clear(sheet, [clipRange(x)], 'contents'); return;
      case k === 'Backspace': e.preventDefault(); if (guard()) startEdit('', 'enter'); return;
      case k === 'Escape': if (copy) setCopy(null); if (painter) setPainter(null); return;
      case k === ' ' && ctrl && !e.shiftKey: e.preventDefault(); setSel({ ...sel, ranges: [{ r1: 0, r2: MAXR - 1, c1: activeRange().c1, c2: activeRange().c2 }] }, { scroll: false }); return;
      case k === ' ' && e.shiftKey && !ctrl: e.preventDefault(); setSel({ ...sel, ranges: [{ r1: activeRange().r1, r2: activeRange().r2, c1: 0, c2: MAXC - 1 }] }, { scroll: false }); return;
      case k === ' ' && e.shiftKey && ctrl: e.preventDefault(); setSel({ ...sel, ranges: [{ r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 }] }, { scroll: false }); return;
      case e.altKey && (k === '=' || k === '0' && false): e.preventDefault(); run('autosum', 'SUM'); return;
      default: break;
    }
    if (ctrl) {
      const map = {
        z: 'undo', y: 'redo', b: 'bold', i: 'italic', u: 'underline', 5: 'strike', f: 'find', h: 'replace', g: 'goto', p: 'print', d: ['fillDir', 'down'], r: ['fillDir', 'right'], '`': 'showFormulas',
      };
      if (lower === 'z' && e.shiftKey) { e.preventDefault(); run('redo'); return; }
      if (lower === 's') { e.preventDefault(); if (host.save) host.save(); return; }
      if (lower === 'a') { e.preventDefault(); const rg = region(sel.r, sel.c); const cur = activeRange(); const isRegion = cur.r1 === rg.r1 && cur.c1 === rg.c1 && cur.r2 === rg.r2 && cur.c2 === rg.c2; setSel({ ...sel, ranges: [isRegion || (rg.r1 === rg.r2 && rg.c1 === rg.c2) ? { r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 } : rg] }, { scroll: false }); return; }
      if ((k === '1' && !e.shiftKey) || (lower === 'f' && e.shiftKey)) { e.preventDefault(); e.stopPropagation(); run('formatCells'); return; }
      if (lower === 'l' && e.shiftKey) { e.preventDefault(); run('filterToggle'); return; }
      if (lower === 'v' && e.shiftKey) { pasteMode.current = { what: 'values' }; return; } // the paste event follows
      if (k === '%' || (e.shiftKey && k === '5')) { e.preventDefault(); run('numPreset', 'percent'); return; }
      if (k === '$') { e.preventDefault(); run('numPreset', 'currency'); return; }
      if (k === '#') { e.preventDefault(); run('numPreset', 'shortDate'); return; }
      if (k === '!') { e.preventDefault(); run('comma'); return; }
      if (k === '-' && !e.shiftKey) { e.preventDefault(); run('deleteCellsDlg'); return; }
      if (k === '+' || (k === '=' && e.shiftKey)) { e.preventDefault(); run('insertCellsDlg'); return; }
      if (lower === ';' || lower === '.' || lower === ':' || lower === ',') {
        e.preventDefault();
        if (!guard()) return;
        const time = e.shiftKey || lower === ':';
        const v = formatValue(time ? nowSerial() : todaySerial(), time ? loc.time : loc.shortDate, loc).text;
        book.run(time ? 'Time' : 'Date', () => book.input(sheet, sel.r, sel.c, v));
        return;
      }
      const c = map[lower] || map[k];
      if (c) { e.preventDefault(); if (Array.isArray(c)) run(c[0], c[1]); else run(c); return; }
      return;
    }
    // printable → start typing (the character lands in the textarea)
    if (k.length === 1 && !e.altKey) {
      if (readOnly) { e.preventDefault(); guard(); return; }
      const sp = sheet.spill.get(K(sel.r, sel.c));
      void sp;
      setEdit({ r: sel.r, c: sel.c, sheetId: sheet.id, text: '', caret: 0, mode: 'enter', where: 'cell', orig: book.editText(sheet, sel.r, sel.c, toDisp), point: null, pc: null });
      if (gridApi.current) gridApi.current.scrollTo(sel.r, sel.c);
    }
  };
  const cycleAbs = (el) => {
    // F4: cycle $A$1 → A$1 → $A1 → A1 for the reference at the caret
    const s = el.selectionStart;
    const text = edit.text;
    const body = text.slice(1);
    const list = refsIn(body, { decimal: loc.decimal });
    const hit = list.find((x) => s - 1 >= x.pos && s - 1 <= x.pos + x.len);
    if (!hit) return;
    const seg = body.slice(hit.pos, hit.pos + hit.len);
    const m = /^((?:'[^']*'|[^!]+)!)?(.*)$/.exec(seg);
    const pre = m[1] || ''; const ref = m[2];
    const cyc = (p) => { const q = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(p); if (!q) return p; const state = (q[1] ? 2 : 0) + (q[3] ? 1 : 0); const next = { 3: 1, 1: 2, 2: 0, 0: 3 }[state]; return `${next & 2 ? '$' : ''}${q[2]}${next & 1 ? '$' : ''}${q[4]}`; };
    const nref = ref.split(':').map(cyc).join(':');
    const nt = `=${body.slice(0, hit.pos)}${pre}${nref}${body.slice(hit.pos + hit.len)}`;
    const caret = 1 + hit.pos + pre.length + nref.length;
    setEdit({ ...edit, text: nt, caret });
    setTimeout(() => el.setSelectionRange(caret, caret), 0);
  };

  const onPaste = (e) => {
    if (edit) return; // normal text paste inside the editor
    e.preventDefault();
    if (!guard()) return;
    const mode = pasteMode.current || {};
    pasteMode.current = null;
    const html = e.clipboardData.getData('text/html');
    const text = e.clipboardData.getData('text/plain');
    pasteData(html || null, text || null, mode);
  };
  const onCopyEv = (e) => { if (edit) return; doCopy(false, e); };
  const onCutEv = (e) => { if (edit) return; if (!guard()) { e.preventDefault(); return; } doCopy(true, e); };

  /* ---------- context menus ---------- */
  const cellMenu = (at, h) => {
    const items = [
      { label: t('Cut'), kbd: 'Ctrl+X', disabled: readOnly, onClick: () => run('cut') }, { label: t('Copy'), kbd: 'Ctrl+C', onClick: () => run('copy') },
      { label: t('Paste'), kbd: 'Ctrl+V', disabled: readOnly, onClick: () => run('paste') },
      { label: t('Paste special'), disabled: readOnly, items: [{ label: t('Values only'), onClick: () => run('pasteSpecial', { what: 'values' }) }, { label: t('Formulas only'), onClick: () => run('pasteSpecial', { what: 'formulas' }) }, { label: t('Formats only'), onClick: () => run('pasteSpecial', { what: 'formats' }) }, { label: t('Transpose'), onClick: () => run('pasteSpecial', { what: 'all', transpose: true }) }, '-', { label: t('Paste special …'), onClick: () => run('pasteSpecialDlg') }] },
      '-',
    ];
    if (h.area === 'colhead') items.push({ label: t('Insert columns'), disabled: readOnly, onClick: () => run('insertCols') }, { label: t('Delete columns'), disabled: readOnly, onClick: () => run('deleteCols') }, { label: t('Clear contents'), disabled: readOnly, onClick: () => run('clear', 'contents') }, '-', { label: t('Column width …'), disabled: readOnly, onClick: () => run('colWidth') }, { label: t('AutoFit column width'), disabled: readOnly, onClick: () => run('autofitCols') }, { label: t('Hide'), disabled: readOnly, onClick: () => run('hide', 'col') }, { label: t('Unhide'), disabled: readOnly, onClick: () => run('unhide', 'col') });
    else if (h.area === 'rowhead') items.push({ label: t('Insert rows'), disabled: readOnly, onClick: () => run('insertRows') }, { label: t('Delete rows'), disabled: readOnly, onClick: () => run('deleteRows') }, { label: t('Clear contents'), disabled: readOnly, onClick: () => run('clear', 'contents') }, '-', { label: t('Row height …'), disabled: readOnly, onClick: () => run('rowHeight') }, { label: t('AutoFit row height'), disabled: readOnly, onClick: () => run('autofitRows') }, { label: t('Hide'), disabled: readOnly, onClick: () => run('hide', 'row') }, { label: t('Unhide'), disabled: readOnly, onClick: () => run('unhide', 'row') }, '-', { label: t('Group'), disabled: readOnly, onClick: () => run('group', 1) });
    else items.push(
      { label: t('Insert …'), disabled: readOnly, onClick: () => run('insertCellsDlg') }, { label: t('Delete …'), disabled: readOnly, onClick: () => run('deleteCellsDlg') }, { label: t('Clear contents'), kbd: 'Del', disabled: readOnly, onClick: () => run('clear', 'contents') }, '-',
      { label: t('Filter'), items: [{ label: sheet.filter ? t('Remove filter') : t('Filter'), onClick: () => run('filterToggle') }, { label: t('Filter by selected cell’s value'), disabled: !sheet.filter, onClick: () => { const c = sel.c; book.setFilterColumn(sheet, c, { values: [book.text(sheet, sel.r, c)] }); } }] },
      { label: t('Sort'), items: [{ label: t('Sort A to Z'), onClick: () => run('sort', 'asc') }, { label: t('Sort Z to A'), onClick: () => run('sort', 'desc') }, { label: t('Custom sort …'), onClick: () => run('sortDlg') }] }, '-',
      { label: curCell && curCell.note ? t('Edit note') : t('New note'), disabled: readOnly, onClick: () => run('note') }, ...(curCell && curCell.note ? [{ label: t('Delete note'), disabled: readOnly, onClick: () => run('clear', 'notes') }] : []),
      { label: curCell && curCell.link ? t('Edit link …') : t('Link …'), disabled: readOnly, onClick: () => run('link') },
    );
    items.push('-', { label: t('Format cells …'), kbd: 'Ctrl+1', disabled: readOnly, onClick: () => run('formatCells') }, { label: t('Define name …'), disabled: readOnly, onClick: () => run('defineName') });
    showMenu(at, items);
  };
  const tabMenu = (at, i) => {
    const sh = book.sheets[i];
    showMenu(at, [
      { label: t('Insert sheet'), disabled: readOnly, onClick: () => run('insertSheet') }, { label: t('Delete'), disabled: readOnly, danger: true, onClick: () => run('deleteSheet') },
      { label: t('Rename'), disabled: readOnly, onClick: () => run('renameSheet') }, { label: t('Duplicate'), disabled: readOnly, onClick: () => run('duplicateSheet') },
      { label: t('Move left'), disabled: readOnly || i === 0, onClick: () => book.moveSheet(i, i - 1) }, { label: t('Move right'), disabled: readOnly || i === book.sheets.length - 1, onClick: () => book.moveSheet(i, i + 1) },
      { label: t('Tab colour …'), disabled: readOnly, onClick: () => run('tabColor', at) },
      { label: t('Hide'), disabled: readOnly, onClick: () => run('hideSheet') }, { label: t('Unhide …'), disabled: readOnly || !book.sheets.some((s) => s.hidden), onClick: () => run('unhideSheet') },
      ...(sh.hasPivot ? ['-', { header: t('This sheet has a pivot table (stored values, kept)') }] : []),
    ]);
  };

  /* ---------- render ---------- */
  if (err) return <div className="of-sheet-root"><Empty icon={CloudOff} title={t('This spreadsheet could not be opened')} text={String(err.message || err)} action={<Btn onClick={() => load()}>{t('Try again')}</Btn>} /></div>;
  if (!book || !sheet) return <div className="of-sheet-root"><Loading text={t('Opening spreadsheet …')} /></div>;

  const g = activeRange();
  const nameBox = edit ? addr(edit.r, edit.c) : (() => {
    for (const n of book.names.values()) { if (n.raw) continue; const rg = parseRange(n.ref); if (rg && (!rg.sheet || rg.sheet.toLowerCase() === sheet.name.toLowerCase()) && rg.r1 === g.r1 && rg.c1 === g.c1 && rg.r2 === g.r2 && rg.c2 === g.c2) return n.name; }
    if (g.r1 === g.r2 && g.c1 === g.c2) return addr(sel.r, sel.c);
    if (g.r1 === 0 && g.r2 >= MAXR - 1) return g.c1 === 0 && g.c2 >= MAXC - 1 ? t('All') : `${colName(g.c1)}:${colName(g.c2)}`;
    if (g.c1 === 0 && g.c2 >= MAXC - 1) return `${g.r1 + 1}:${g.r2 + 1}`;
    return `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`;
  })();
  const barText = edit ? edit.text : (sheet.cells.get(K(sel.r, sel.c)) ? book.editText(sheet, sel.r, sel.c, toDisp) : (() => { const sp = sheet.spill.get(K(sel.r, sel.c)); if (!sp) return ''; const a = sp.anchor; const ac2 = sheet.cells.get(a); return ac2 && ac2.f ? `=${toDisp(ac2.f)}` : ''; })());
  const spilledHere = !edit && !sheet.cells.has(K(sel.r, sel.c)) && sheet.spill.has(K(sel.r, sel.c));
  const nameGo = (v) => {
    const s = String(v).trim(); if (!s) return false;
    let gg = parseRange(s);
    let sh = sheet;
    if (!gg) { const nm = book.nameDef(s, sheet); if (nm) gg = parseRange(nm.ref); }
    if (!gg) {
      if (/^[A-Za-z_À-￿][\w.À-￿]*$/.test(s) && !readOnly) { if (book.setName(s, rangeText({ sheet: sheet.name, ...clipRange(g), ar1: true, ac1: true, ar2: true, ac2: true }), null)) { toast.success(t('Name “{name}” defined', { name: s })); return true; } }
      toast.error(t('“{ref}” is not a valid reference or name.', { ref: s })); return false;
    }
    if (gg.sheet) { const i = book.sheets.findIndex((x) => x.name.toLowerCase() === gg.sheet.toLowerCase()); if (i < 0) { toast.error(t('Sheet “{name}” doesn’t exist.', { name: gg.sheet })); return false; } sh = book.sheets[i]; if (sh !== sheet) switchSheet(i); }
    setTimeout(() => { setSel({ r: gg.r1, c: gg.c1, ranges: [{ r1: gg.r1, c1: gg.c1, r2: gg.r2, c2: gg.c2 }], anchor: { r: gg.r1, c: gg.c1 }, end: { r: gg.r2, c: gg.c2 } }); focusGrid(); }, 0);
    return true;
  };
  // status bar numbers
  const stats = (() => {
    let sum = 0; let n = 0; let cnt = 0; let min = Infinity; let max = -Infinity;
    const area = sel.ranges.reduce((a, x) => a + (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1), 0);
    if (area < 2) return null;
    const add = (v) => { if (v == null || v === '') return; cnt++; if (typeof v === 'number') { sum += v; n++; if (v < min) min = v; if (v > max) max = v; } };
    if (area > 100000) { for (const [k, c] of sheet.cells) { const r = Math.floor(k / MAXC); const cc = k % MAXC; if (sel.ranges.some((x) => r >= x.r1 && r <= x.r2 && cc >= x.c1 && cc <= x.c2) && !sheet.filterHidden.has(r)) add(c.v); } } else for (const x of sel.ranges) for (let r = x.r1; r <= x.r2; r++) { if (sheet.filterHidden.has(r) || sheet.hiddenRows.has(r)) continue; for (let c = x.c1; c <= x.c2; c++) add(book.value(sheet, r, c)); }
    if (!cnt) return null;
    const nf = (curStyle && curStyle.nf && !isDateFormat(curStyle.nf)) ? curStyle.nf : 'General';
    const f = (v) => formatValue(Number(v.toPrecision(15)), nf, loc).text;
    return { cnt, n, sum: n ? f(sum) : null, avg: n ? f(sum / n) : null, min: n ? f(min) : null, max: n ? f(max) : null };
  })();
  const mode = edit ? (pointable(edit) && edit.mode === 'enter' ? t('Point') : edit.mode === 'edit' ? t('Edit') : t('Enter')) : copy ? t('Select destination and press Enter or Paste') : t('Ready');
  const dvRule = !edit && dvAt(sheet, sel.r, sel.c);
  const presets = FORMAT_PRESETS(loc);
  const kindNow = curStyle && curStyle.nf ? fmtKind(curStyle.nf) : 'general';
  const flags = {
    loc,
    readOnly, canUndo: book.canUndo(), canRedo: book.canRedo(), defFont, formatPresets: presets,
    fmtKind: presets.find((p) => p.code === (curStyle && curStyle.nf))?.id || (kindNow === 'general' ? 'general' : kindNow === 'date' ? (presets.find((p) => p.code === (curStyle && curStyle.nf)) ? 'shortDate' : 'custom') : presets.find((p) => fmtKind(p.code) === kindNow)?.id || 'custom'),
    currency: loc.currency, painter, hasFilter: !!sheet.filter, filtered: !!(sheet.filter && Object.keys(sheet.filter.cols).length), merged: !!book.mergeAt(sheet, sel.r, sel.c),
    lastFill: last.fill, lastColor: last.color, lastBorder: last.border, borderLine: last.line.s, cellStyles: CELL_STYLES,
    orientation: sheet.print.orientation, fit: sheet.print.fit, printArea: sheet.print.area, printGrid: sheet.print.gridlines, printHead: sheet.print.headings,
    showFormulas, autoCalc: book.autoCalc, hasNote: !!(curCell && curCell.note), showNotes, hasMacros: (book.fidelity || []).some((x) => /Macros/.test(x)) || ['xlsm', 'xltm'].includes(workExt),
    frozen: !!(sheet.freeze.rows || sheet.freeze.cols), gridlines: sheet.gridlines !== false, headings: sheet.headings !== false, formulaBar, compact, isCsv: CSV.includes(workExt),
  };
  const editorFont = (() => { const s = (edit && editSheet && editSheet.cells.get(K(edit.r, edit.c)) || {}).s || curStyle || {}; return { fontFamily: `"${s.font || defFont.font}", Calibri, Carlito, sans-serif`, fontSize: `${((s.size || defFont.size) * 4 / 3) * (zoom / 100)}px`, fontWeight: s.b ? 700 : 400, fontStyle: s.i ? 'italic' : 'normal', color: s.color || '#000', textAlign: s.h === 'right' || s.h === 'center' ? s.h : 'left', background: s.fill || '#fff' }; })();
  const marks = refs ? refs.map((x) => ({ pos: x.pos, len: x.len, color: x.color })) : null;
  const linkHere = !edit && curCell && (curCell.link || curCell.link_);
  const openLink = (u) => {
    if (u.startsWith('#')) { nameGo(u.slice(1)); return; }
    if (host.openLink) host.openLink(u); else toast(t('Links open in Hostwerk’s own window – this Hostwerk version can’t do that yet.'));
  };
  const cmdList = [
    ...['Merge & center', 'Freeze panes', 'Filter', 'Sort A to Z', 'Sort Z to A', 'Format cells', 'Conditional formatting', 'Data validation', 'Remove duplicates', 'Text to columns', 'Insert chart', 'Name manager', 'Find & replace', 'Go to', 'Print', 'Wrap text', 'Show formulas', 'Insert rows', 'Insert columns', 'Delete rows', 'Delete columns', 'Rename sheet', 'Insert sheet', 'Paste special'].map((l) => ({ id: l, label: t(l) })),
    ...FUNCTION_NAMES.filter(isExcelName).map((n) => ({ id: `fn:${n}`, label: localFn(n, loc), hint: FUNC_INFO.get(n) ? t(FUNC_INFO.get(n).desc) : t('Function') })),
  ];
  const runCmd = (c) => {
    if (c.id.startsWith('fn:')) { run('insertFn', c.id.slice(3)); return; }
    const m = { 'Merge & center': ['merge', 'center'], 'Freeze panes': ['freeze', 'sel'], Filter: ['filterToggle'], 'Sort A to Z': ['sort', 'asc'], 'Sort Z to A': ['sort', 'desc'], 'Format cells': ['formatCells'], 'Conditional formatting': ['cfManage'], 'Data validation': ['validation'], 'Remove duplicates': ['removeDuplicates'], 'Text to columns': ['textToColumns'], 'Insert chart': ['chart', 'column'], 'Name manager': ['names'], 'Find & replace': ['replace'], 'Go to': ['goto'], Print: ['print'], 'Wrap text': ['wrap'], 'Show formulas': ['showFormulas'], 'Insert rows': ['insertRows'], 'Insert columns': ['insertCols'], 'Delete rows': ['deleteRows'], 'Delete columns': ['deleteCols'], 'Rename sheet': ['renameSheet'], 'Insert sheet': ['insertSheet'], 'Paste special': ['pasteSpecialDlg'] }[c.id];
    if (m) run(m[0], m[1]);
  };

  return (
    <div className={`of-sheet-root${compact ? ' is-compact' : ''}`}>
      <MenuHost />
      {!compact && <Ribbon run={run} cur={curStyle} flags={flags} />}
      {!host.fidelity && fid.length > 0 && <div className="of-sheet-callout"><Callout kind="info" title={t('Some parts of this file can’t be edited here')}>{t('They are kept unchanged when you save: {list}.', { list: fid.map((x) => t(x)).join(' · ') })}</Callout></div>}
      {CSV.includes(workExt) && csvNote && book.csv && <div className="of-sheet-strip small"><Info size={14} /><span className="grow">{t('Text file · delimiter {d} · {enc}. Formatting, formulas’ results and further sheets are not stored in this format – save a copy as .xlsx to keep them.', { d: book.csv.delimiter === '\t' ? t('Tab') : book.csv.delimiter, enc: book.csv.encoding })}</span><Btn className="xs ghost" onClick={() => run('csvOptions')}>{t('Import options …')}</Btn><button type="button" className="btn ghost xs icon" aria-label={t('Hide')} title={t('Hide')} onClick={() => setCsvNote(false)}><X size={12} /></button></div>}
      {formulaBar && (
        <div className="of-sheet-fbar">
          <input className="input sm mono of-sheet-namebox" aria-label={t('Name box')} title={t('Name box – type a cell, range or name and press Enter')} key={nameBox} defaultValue={nameBox}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); nameGo(e.currentTarget.value); } if (e.key === 'Escape') { e.currentTarget.value = nameBox; focusGrid(); } }} onFocus={(e) => e.target.select()} />
          <button type="button" className="btn ghost sm icon" aria-label={t('Names')} title={t('Names')} onClick={(e) => { const list = [...book.names.values()].filter((n) => !n.raw && !n.hidden); showMenu({ left: e.currentTarget.getBoundingClientRect().left, below: e.currentTarget.getBoundingClientRect().bottom + 2 }, list.length ? list.map((n) => ({ label: n.name, keep: true, onClick: () => nameGo(n.name) })) : [{ header: t('No names yet') }, { label: t('Define name …'), onClick: () => run('defineName') }]); }}><ChevronDown size={13} /></button>
          {edit ? <>
            <button type="button" className="btn ghost sm icon" aria-label={t('Cancel')} title={`${t('Cancel')} (Esc)`} onMouseDown={(e) => e.preventDefault()} onClick={cancelEdit}><X size={15} /></button>
            <button type="button" className="btn ghost sm icon" aria-label={t('Enter')} title={`${t('Enter')} (Enter)`} onMouseDown={(e) => e.preventDefault()} onClick={() => commit('stay')}><Check size={15} /></button>
          </> : null}
          <button type="button" className="btn ghost sm icon" aria-label={t('Insert function')} title={t('Insert function')} disabled={readOnly} onClick={() => run('insertFunction')}><SquareFunction size={15} /></button>
          <CellInput className="of-sheet-barinput" inputRef={barInput} value={barText} readOnly={readOnly || spilledHere} ariaLabel={t('Formula bar')} marks={edit ? marks : null}
            onFocus={() => { if (!edit && !readOnly && !spilledHere) startEdit(undefined, 'edit', 'bar'); else if (edit && edit.where !== 'bar') setEdit({ ...edit, where: 'bar', mode: 'edit' }); }}
            onChange={(v, caret) => { if (!edit) return; setEdit({ ...edit, text: v, caret, point: null }); updateAc(v, caret); }}
            onSelect={(s) => { if (edit && s !== edit.caret) setEdit((x) => x && { ...x, caret: s, point: x.point && x.point.end === s ? x.point : null }); }}
            onKeyDown={onKey} />
        </div>
      )}
      <div className="of-sheet-main">
        <Grid ref={gridApi} book={book} sheet={sheet} version={ver} sel={sel} zoom={zoom} readOnly={readOnly} showFormulas={showFormulas} loc={loc} defFont={defFont}
          edit={edit && edit.sheetId === sheet.id ? edit : edit ? { other: true } : null} refs={refs} pointing={!!(edit && pointable(edit))} copy={copy && { sheetId: copy.sheetId, g: copy.g }}
          onSel={(s) => { const single = s.ranges.length === 1 && s.ranges[0].r1 === s.ranges[0].r2 && s.ranges[0].c1 === s.ranges[0].c2; if (edit && !pointable(edit)) { commit(null).then((ok) => { if (ok) setSel(s, { scroll: single }); }); return; } setSel(s, { scroll: single }); setChartSel(null); }}
          expandMerges={expandMerges}
          onPointRef={(gg, phase) => insertRef(gg, phase)}
          onEditStart={(r, c) => { if (readOnly) return; const cur = sheet.cells.get(K(r, c)); if (cur && (cur.link || cur.link_) && false) return; setSelState(cellSel(r, c)); setTimeout(() => startEdit(undefined, 'edit', 'cell', { r, c }), 0); }}
          onFill={(src, dst, ctrlCopy) => { if (!guard()) return; book.fill(sheet, src, dst, { copy: ctrlCopy }); setSel({ r: sel.r, c: sel.c, ranges: [dst], anchor: sel.anchor, end: { r: dst.r2, c: dst.c2 } }, { scroll: false }); }}
          onMove={(src, r, c, cp) => { if (!guard()) return; const res = book.moveRange(sheet, src, r, c, { copy: cp }); setSel({ r: res.r1, c: res.c1, ranges: [res], anchor: { r: res.r1, c: res.c1 }, end: { r: res.r2, c: res.c2 } }, { scroll: false }); }}
          onResize={(axis, idxs, px) => { if (!guard()) return; book.setSizes(sheet, axis, idxs, px); }}
          onContextMenu={(at, h) => { if (edit) commit(null); cellMenu(at, h); }}
          onFilterButton={(col, at) => showPopover(at, <FilterMenu book={book} sheet={sheet} col={col} onClose={closeMenu} onSort={(desc) => { const f = sheet.filter.range; book.sort(sheet, clipRange(f), [{ col, desc }], { header: true }); }} onApply={(crit) => book.setFilterColumn(sheet, col, crit)} />)}
          onZoom={(d) => setZoom(zoom + d)}
          focus={() => { if (!edit) input.current && input.current.focus(); }}
          chartSel={chartSel} chartData={chartData} onChartSel={(id) => setChartSel(id)} onChartChange={(id, a) => { if (!guard()) return; updateChart(id, { anchor: a._px ? normAnchor(a) : a }); }}
          onChartOpen={(id) => setDialog({ kind: 'chart', id })} onChartMenu={(id, at) => showMenu(at, [{ label: t('Edit chart …'), onClick: () => setDialog({ kind: 'chart', id }) }, { label: t('Delete chart'), danger: true, disabled: readOnly, onClick: () => deleteChart(id) }])}
          showNotes={showNotes}>
          {(geo, cellRect) => {
            const er = edit && edit.sheetId === sheet.id ? cellRect(edit.r, edit.c) : cellRect(sel.r, sel.c);
            const visible = !!(edit && edit.sheetId === sheet.id && edit.where === 'cell');
            const maxW = geo.W - er.x - 4;
            const lines = edit ? edit.text.split('\n').length : 1;
            const estW = edit ? Math.min(maxW, Math.max(er.w, Math.max(...edit.text.split('\n').map((l) => l.length)) * 8 * (zoom / 100) + 24)) : er.w;
            return <>
              <CellInput inputRef={input} className={`of-sheet-celledit${visible ? ' is-on' : ''}`} value={visible ? edit.text : ''} marks={visible ? marks : null} ariaLabel={t('Cell {ref}', { ref: addr(sel.r, sel.c) })}
                style={visible ? { left: er.x, top: er.y, minWidth: er.w, width: estW, minHeight: er.h, height: Math.max(er.h, lines * parseFloat(editorFont.fontSize) * 1.25 + 6), ...editorFont } : { left: Math.max(0, er.x), top: Math.max(0, er.y), width: 1, height: 1 }}
                onChange={(v, caret) => { if (!edit) { if (readOnly) return; setEdit({ r: sel.r, c: sel.c, sheetId: sheet.id, text: v, caret, mode: 'enter', where: 'cell', orig: book.editText(sheet, sel.r, sel.c, toDisp), point: null, pc: null }); updateAc(v, caret); return; } setEdit({ ...edit, text: v, caret, point: edit.point && caret === edit.point.end ? edit.point : null }); updateAc(v, caret); }}
                onSelect={(s) => { if (edit && s !== edit.caret) setEdit((x) => x && { ...x, caret: s, point: x.point && x.point.end === s ? x.point : null }); }}
                onCompositionStart={() => { if (!edit && !readOnly) setEdit({ r: sel.r, c: sel.c, sheetId: sheet.id, text: '', caret: 0, mode: 'enter', where: 'cell', orig: book.editText(sheet, sel.r, sel.c, toDisp), point: null, pc: null }); }}
                onKeyDown={onKey} onPaste={onPaste} onCopy={onCopyEv} onCut={onCutEv} />
              {edit && ac && (edit.where === 'cell' ? visible : true) && <div className="of-sheet-ac" role="listbox" style={edit.where === 'bar' ? { left: 200, top: 2 } : { left: er.x, top: er.y + Math.max(er.h, 20) + 2 }}>
                {ac.items.map((it, i) => <button key={it.name + it.kind} type="button" role="option" aria-selected={i === ac.i} className={i === ac.i ? 'is-on' : ''} onMouseDown={(e) => { e.preventDefault(); acceptAc(it); }}><b className="mono"><Keep>{it.name}</Keep></b><span className="small muted">{it.kind === 'name' ? t('Name') : FUNC_INFO.get(it.en) ? t(FUNC_INFO.get(it.en).desc) : ''}</span></button>)}
              </div>}
              {edit && hint && !ac && <div className="of-sheet-hint small" style={edit.where === 'bar' ? { left: 200, top: 2 } : { left: er.x, top: er.y + Math.max(er.h, 20) + 2 }}>
                <span className="mono"><b><Keep>{hint.name}</Keep></b>({hint.args.map((a, i) => <span key={i}>{i ? (loc.decimal === ',' ? '; ' : ', ') : ''}<span className={i === Math.min(hint.arg, hint.args.length - 1) || (hint.args[hint.args.length - 1] === '…' && i === hint.args.length - 2 && hint.arg >= i) ? 'is-arg' : ''}><Keep>{a}</Keep></span></span>)})</span>
                {hint.desc && <span className="muted"> – {t(hint.desc)}</span>}
              </div>}
              {dvRule && dvRule.type === 'list' && dvRule.dropdown !== false && !readOnly && (() => { const rc = cellRect(sel.r, sel.c); return <button type="button" className="of-sheet-dvbtn" style={{ left: rc.x + rc.w + 1, top: rc.y + rc.h - 18 }} aria-label={t('Choose from the list')} title={t('Choose from the list')} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); const list = dvList(sheet, dvRule, sel.r, sel.c); showPopover({ left: r.left - rc.w, below: r.bottom + 2 }, <div className="of-sheet-dvlist">{list.length ? list.map((x) => <button key={x} type="button" onClick={() => { closeMenu(); book.run('Choose', () => book.input(sheet, sel.r, sel.c, x)); focusGrid(); }}><Keep>{x}</Keep></button>) : <span className="small muted">{t('The list is empty.')}</span>}</div>); }}><ChevronDown size={12} /></button>; })()}
              {dvRule && dvRule.prompt && (dvRule.prompt.title || dvRule.prompt.text) && (() => { const rc = cellRect(sel.r, sel.c); return <div className="of-sheet-prompt small" style={{ left: rc.x + 8, top: rc.y + rc.h + 6 }}>{dvRule.prompt.title && <b><Keep>{dvRule.prompt.title}</Keep></b>}<div><Keep>{dvRule.prompt.text}</Keep></div></div>; })()}
              {linkHere && (() => { const rc = cellRect(sel.r, sel.c); const u = curCell.link || curCell.link_; return <div className="of-sheet-linkchip small" style={{ left: rc.x, top: rc.y + rc.h + 4 }}><ExternalLink size={12} /><button type="button" className="btn ghost xs" onClick={() => openLink(u)}>{u.startsWith('#') ? t('Go to {ref}', { ref: u.slice(1) }) : <Keep>{u.length > 50 ? `${u.slice(0, 50)}…` : u}</Keep>}</button>{!readOnly && curCell.link && <button type="button" className="btn ghost xs" onClick={() => run('link')}>{t('Edit')}</button>}</div>; })()}
              {showNotes && [...sheet.cells].filter(([, c]) => c.note).slice(0, 200).map(([k, c]) => { const rc = cellRect(Math.floor(k / MAXC), k % MAXC); if (rc.x > geo.W || rc.y > geo.H || rc.x + rc.w < 0 || rc.y < 0) return null; return <div key={k} className="of-sheet-note is-static" style={{ left: rc.x + rc.w + 6, top: rc.y }}>{c.noteAuthor && <b className="small"><Keep>{c.noteAuthor}</Keep></b>}<div className="small"><Keep>{c.note}</Keep></div></div>; })}
            </>;
          }}
        </Grid>
      </div>
      <div className="of-sheet-bottom">
        <div className="of-sheet-tabs" role="tablist" aria-label={t('Sheets')}>
          <button type="button" className="btn ghost xs icon" aria-label={t('Previous sheet')} title={t('Previous sheet (Ctrl+PageUp)')} onClick={() => { let i = book.active - 1; while (i >= 0 && book.sheets[i].hidden) i--; if (i >= 0) switchSheet(i); }}><ChevronLeft size={14} /></button>
          <button type="button" className="btn ghost xs icon" aria-label={t('Next sheet')} title={t('Next sheet (Ctrl+PageDown)')} onClick={() => { let i = book.active + 1; while (i < book.sheets.length && book.sheets[i].hidden) i++; if (i < book.sheets.length) switchSheet(i); }}><ChevronRight size={14} /></button>
          <div className="of-sheet-tablist">
            {book.sheets.map((sh, i) => sh.hidden ? null : (
              <SheetTab key={sh.id} sheet={sh} active={i === book.active} readOnly={readOnly}
                onClick={() => { if (edit && !formulaText) commit(null).then((ok) => ok && switchSheet(i)); else switchSheet(i); }}
                onRename={(name) => { if (!book.renameSheet(sh, name)) toast.error(t('This name can’t be used.')); }}
                onMenu={(at) => { switchSheet(i); tabMenu(at, i); }}
                onDrop={(from) => { if (!readOnly) book.moveSheet(from, i); }} index={i} />
            ))}
          </div>
          {!readOnly && <button type="button" className="btn ghost xs icon" aria-label={t('Insert sheet')} title={t('Insert sheet')} onClick={() => run('insertSheet')}><Plus size={14} /></button>}
        </div>
        <div className="of-sheet-status small" role="status">
          <span className="muted">{mode}</span>
          {book.calc.circular && <span className="of-sheet-warn"><AlertTriangle size={12} />{t('Circular reference: {ref}', { ref: `${book.sheets.find((s) => s.id === book.calc.circular.sheet.id)?.name || ''}!${addr(Math.floor(book.calc.circular.key / MAXC), book.calc.circular.key % MAXC)}` })}</span>}
          {!book.autoCalc && <button type="button" className="btn ghost xs" onClick={() => run('calcNow')} title={t('Calculation is set to manual – press F9 to calculate')}>{t('Calculate (F9)')}</button>}
          {sheet.filter && sheet.filterHidden.size > 0 && <span className="muted">{t('Filter mode')}</span>}
          <span className="grow" />
          {stats && <>
            {stats.avg != null && <span>{t('Average')}: <b>{stats.avg}</b></span>}
            <span>{t('Count')}: <b>{stats.cnt}</b></span>
            {stats.sum != null && <span>{t('Sum')}: <b>{stats.sum}</b></span>}
            {stats.min != null && <span className="of-sheet-hide-sm">{t('Min')}: <b>{stats.min}</b></span>}
            {stats.max != null && <span className="of-sheet-hide-sm">{t('Max')}: <b>{stats.max}</b></span>}
          </>}
          <span className="of-sheet-zoom">
            <button type="button" className="btn ghost xs icon" aria-label={t('Zoom out')} title={t('Zoom out')} onClick={() => setZoom(zoom - 10)}><Minus size={12} /></button>
            <input type="range" min={10} max={400} step={5} value={zoom} onChange={(e) => setZoom(+e.target.value)} aria-label={t('Zoom')} />
            <button type="button" className="btn ghost xs icon" aria-label={t('Zoom in')} title={t('Zoom in')} onClick={() => setZoom(zoom + 10)}><Plus size={12} /></button>
            <button type="button" className="btn ghost xs" onClick={(e) => showMenu({ left: e.currentTarget.getBoundingClientRect().left - 40, above: e.currentTarget.getBoundingClientRect().top - 2, below: e.currentTarget.getBoundingClientRect().top - 260 }, [25, 50, 75, 100, 125, 150, 200, 300, 400].map((z) => ({ label: `${z} %`, checked: z === zoom, onClick: () => setZoom(z) })))} title={t('Zoom')}>{zoom} %</button>
          </span>
        </div>
      </div>
      {renderDialog()}
    </div>
  );

  function renderDialog() {
    if (!dialog) return null;
    const close = () => { setDialog(null); focusGrid(); };
    const g2 = activeRange();
    switch (dialog.kind) {
      case 'formatCells': return <FormatCellsDialog book={book} style={curStyle} value={book.value(sheet, sel.r, sel.c)} initialTab={dialog.tab} readOnly={readOnly} onClose={close}
        onApply={(patch, ops, merge) => { book.begin('Format cells'); try { if (Object.keys(patch).length) book.format(sheet, sel.ranges, patch); for (const op of ops) for (const x of sel.ranges) book.borders(sheet, clipRange(x), op.preset, op.line); if (merge === true) for (const x of sel.ranges) book.merge(sheet, x, {}); if (merge === false && book.mergeAt(sheet, sel.r, sel.c)) for (const x of sel.ranges) book.unmerge(sheet, x); } finally { book.end(); } }} />;
      case 'sort': return <SortDialog book={book} sheet={sheet} g={dialog.g} header0={dialog.header} onClose={close} onApply={(levels, header) => book.sort(sheet, dialog.g, levels, { header })} />;
      case 'ttc': return <TextToColumnsDialog book={book} sheet={sheet} g={dialog.g} onClose={close} onApply={(o) => book.textToColumns(sheet, dialog.g, o)} />;
      case 'dups': return <RemoveDuplicatesDialog book={book} sheet={sheet} g={dialog.g} onClose={close} onApply={(cols, header) => { const n = book.removeDuplicates(sheet, dialog.g, cols, header); toast.success(n ? t('{n} duplicate rows removed', { n }) : t('No duplicates found.')); }} />;
      case 'dv': { const rule = dvAt(sheet, sel.r, sel.c); return <ValidationDialog rule={rule} toDisplay={toDisp} fromDisplay={fromDisp} onClose={close}
        onApply={(r) => { const ranges = sel.ranges.map(clipRange); const others = sheet.dv.map((d) => ({ ...d, ranges: d.ranges.filter((x) => !ranges.some((y) => y.r1 <= x.r1 && y.r2 >= x.r2 && y.c1 <= x.c1 && y.c2 >= x.c2)) })).filter((d) => d.ranges.length && d !== rule); book.setSheetProps(sheet, { dv: [...others.filter((d) => d.id !== (rule && rule.id)), { ...r, id: (rule && rule.id) || `dv${Date.now().toString(36)}`, ranges: rule ? rule.ranges : ranges }] }, 'Data validation'); }}
        onClear={() => { const ranges = sel.ranges; book.setSheetProps(sheet, { dv: sheet.dv.map((d) => ({ ...d, ranges: d.ranges.filter((x) => !ranges.some((y) => x.r1 <= y.r2 && y.r1 <= x.r2 && x.c1 <= y.c2 && y.c1 <= x.c2)) })).filter((d) => d.ranges.length) }, 'Data validation'); }} />; }
      case 'csv': return <CsvDialog bytes={dialog.bytes} title={dialog.reimport ? t('Text import options') : t('Import CSV into this sheet')} initial={{ delimiter: book.csv ? book.csv.delimiter : 'auto', encoding: book.csv ? book.csv.encoding : 'auto' }} onClose={close}
        onApply={async (o) => {
          if (dialog.reimport) { if (book.canUndo() && !(await confirmDialog({ title: t('Import the file again?'), text: t('Changes made since opening are discarded.'), confirm: t('Import again'), danger: true }))) return; csvOpts.current = o; load(o); return; }
          const nb = csvToBook(dialog.bytes, { loc, settings: S, delimiter: o.delimiter, encoding: o.encoding, numberLocale: o.numberLocale, header: o.header });
          const src = nb.sheets[0]; const u = nb.usedRange(src); if (!u) return;
          const blk = nb.copyBlock(src, u); blk.sheetId = null;
          book.paste(sheet, { r1: sel.r, c1: sel.c, r2: sel.r + u.r2, c2: sel.c + u.c2 }, blk, {});
        }} />;
      case 'cfRule': return <RuleDialog rule={dialog.rule} ranges={sel.ranges.map(clipRange)} toDisplay={toDisp} fromDisplay={fromDisp} onClose={() => (dialog.back ? setDialog({ kind: 'cfManage' }) : close())}
        onSave={(r) => { const list = dialog.index != null ? sheet.cf.map((x, i) => (i === dialog.index ? { ...x, ...r } : x)) : [...sheet.cf, { ...r, id: `cf${Date.now().toString(36)}` }]; book.setSheetProps(sheet, { cf: list.map((x, i) => ({ ...x, priority: i + 1 })) }, 'Conditional formatting'); if (dialog.back) setTimeout(() => setDialog({ kind: 'cfManage' }), 0); }} />;
      case 'cfManage': return <RulesManager sheet={sheet} sel={sel} toDisplay={toDisp} onClose={close} onChange={(cf) => book.setSheetProps(sheet, { cf }, 'Conditional formatting')} onEdit={(i) => setDialog({ kind: 'cfRule', rule: sheet.cf[i], index: i, back: true })} onNew={() => setDialog({ kind: 'cfRule', rule: { type: 'cell', op: 'gt', style: CF_STYLES[0].style }, back: true })} />;
      case 'find': return <FindPanel book={book} sheet={sheet} replace={dialog.replace} readOnly={readOnly} onClose={close} onGo={(h) => { const i = book.sheets.indexOf(h.sheet); if (i !== book.active) switchSheet(i); setTimeout(() => go(h.r, h.c), 0); }} />;
      case 'goto': return <GoToDialog book={book} special={dialog.special} onClose={close} onGo={(v) => { if (v && v.special) { selectSpecial(v.special); return true; } return nameGo(v); }} />;
      case 'names': return <NamesDialog book={book} current={`=${rangeText({ sheet: sheet.name, ...clipRange(g2), ar1: true, ac1: true, ar2: true, ac2: true })}`} startNew={dialog.startNew} readOnly={readOnly} toDisplay={toDisp} onClose={close} onDelete={(k) => book.deleteName(k)}
        onSave={(e) => { const scope = e.scope ? book.sheetById(e.scope) : null; const ok = book.setName(e.name.trim(), e.ref, scope, e.comment, e.key); if (!ok) toast.error(t('“{name}” can’t be used as a name (it must start with a letter and must not look like a cell address).', { name: e.name })); return ok; }} />;
      case 'insertFn': return <InsertFunctionDialog loc={loc} onClose={close} onPick={(n) => run('insertFn', n)} />;
      case 'link': return <LinkDialog book={book} init={{ link: curCell && curCell.link, text: book.text(sheet, sel.r, sel.c) }} onClose={close}
        onApply={(link, text) => book.run('Link', () => { const cell = sheet.cells.get(K(sel.r, sel.c)); if (text && text !== book.text(sheet, sel.r, sel.c)) book.input(sheet, sel.r, sel.c, text); else if (!cell || cell.v == null) book.input(sheet, sel.r, sel.c, text || (link ? link.replace(/^mailto:|^#/, '') : '')); const c2 = sheet.cells.get(K(sel.r, sel.c)) || { v: null }; book.setCell(sheet, sel.r, sel.c, { ...c2, link: link || undefined, s: link ? book.styles.merge(c2.s, { color: (c2.s && c2.s.color) || '#0563C1', u: true }) : c2.s }); })} />;
      case 'note': return <NoteDialog init={curCell && curCell.note} author={(settings.general && settings.general.author) || ''} onClose={close} onApply={(v) => book.run('Note', () => book.patchCell(sheet, sel.r, sel.c, { note: v || undefined, noteAuthor: v ? (settings.general && settings.general.author) || (curCell && curCell.noteAuthor) || undefined : undefined }))} />;
      case 'chart': { const ch = sheet.charts.find((c) => c.id === dialog.id); if (!ch) return null; return <ChartDialog chart={ch} onClose={close} onDelete={() => deleteChart(ch.id)} onApply={(c) => updateChart(ch.id, c)} />; }
      case 'print': return <PrintDialog sheet={sheet} hasSelection={!(g2.r1 === g2.r2 && g2.c1 === g2.c2)} hasConverter={!!(host.hasConverter && host.convert)} onClose={close}
        onPrint={(o) => { const range = o.what === 'selection' ? clipRange(g2) : null; const html = bookToHTML(book, { sheet: book.active, range, gridlines: o.gridlines, headings: o.headings, orientation: o.orientation, fit: o.fit }); if (!readOnly && (o.orientation !== sheet.print.orientation || o.fit !== sheet.print.fit || o.gridlines !== !!sheet.print.gridlines || o.headings !== !!sheet.print.headings)) book.setSheetProps(sheet, { print: { ...sheet.print, orientation: o.orientation, fit: o.fit, gridlines: o.gridlines, headings: o.headings } }, 'Page setup'); if (host.print) host.print(html); else toast(t('Printing needs a newer Hostwerk version.')); }}
        onPdf={async (o) => { try { const tmp = { ...sheet.print }; Object.assign(sheet.print, { orientation: o.orientation, fit: o.fit, gridlines: o.gridlines, headings: o.headings }); const x = await writeXlsx(book, { ext: 'xlsx', getValues }); Object.assign(sheet.print, tmp); const pdf = await host.convert(x, 'xlsx', 'pdf'); host.download(pdf, `${String(file.name || 'sheet').replace(/\.[^.]+$/, '')}.pdf`, 'application/pdf'); } catch (e) { toast.error(e.message || String(e)); } }} />;
      case 'pasteSpecial': return <PasteSpecialDialog onClose={close} onApply={(m) => run('pasteSpecial', m)} />;
      case 'cells': return <CellsDialog insert={dialog.insert} onClose={close} onApply={(v) => { if (v === 'row') run(dialog.insert ? 'insertRows' : 'deleteRows'); else if (v === 'col') run(dialog.insert ? 'insertCols' : 'deleteCols'); else book.shiftCells(sheet, clipRange(g2), v, dialog.insert); }} />;
      case 'macros': return <MacrosSheet onClose={close} load={async () => { const b = book.pkg ? await vbaProjectBytes(book.pkg.bytes) : null; return b ? readVba(b) : []; }} />;
      case 'stats': return <StatsDialog book={book} onClose={close} />;
      case 'cmdSearch': return <CommandSearch commands={cmdList} onClose={close} onRun={runCmd} />;
      case 'unhideSheet': return <UnhideSheetDialog book={book} onClose={close} onApply={(id) => { const sh = book.sheetById(id); book.setSheetProps(sh, { hidden: false }, 'Unhide sheet'); switchSheet(book.sheets.indexOf(sh)); }} />;
      default: return null;
    }
  }
}

function SheetTab({ sheet, active, readOnly, onClick, onRename, onMenu, onDrop, index }) {
  const [renaming, setRenaming] = useState(false);
  const [v, setV] = useState(sheet.name);
  useEffect(() => setV(sheet.name), [sheet.name]);
  if (renaming) return <input className="input xs of-sheet-tabedit" autoFocus value={v} onChange={(e) => setV(e.target.value)} aria-label={t('Sheet name')} onBlur={() => { setRenaming(false); if (v.trim() && v !== sheet.name) onRename(v.trim()); }} onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); if (e.key === 'Escape') { setV(sheet.name); setRenaming(false); } }} />;
  return (
    <Raw as="button" type="button" role="tab" aria-selected={active} className={`of-sheet-tab${active ? ' is-on' : ''}`} style={sheet.color ? { '--tab': sheet.color } : undefined}
      draggable={!readOnly} onDragStart={(e) => { e.dataTransfer.setData('text/x-sheet-index', String(index)); e.dataTransfer.effectAllowed = 'move'; }}
      onDragOver={(e) => { if ([...e.dataTransfer.types].includes('text/x-sheet-index')) e.preventDefault(); }}
      onDrop={(e) => { const from = Number(e.dataTransfer.getData('text/x-sheet-index')); if (!Number.isNaN(from)) onDrop(from); }}
      onClick={onClick} onDoubleClick={() => { if (!readOnly) setRenaming(true); }} onContextMenu={(e) => { e.preventDefault(); onMenu({ x: e.clientX, y: e.clientY }); }} title={sheet.name}>
      {sheet.color && <span className="of-sheet-tabcolor" />}<Keep>{sheet.name}</Keep>
    </Raw>
  );
}

export { MenuList, PALETTE };
