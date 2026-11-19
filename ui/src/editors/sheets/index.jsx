/*
 * Spreadsheet editor chunk (ARCHITECTURE.md "Editor chunk contract"): Editor, create (templates), toText, Preview.
 */
import { useEffect, useState } from 'react';
import { t } from '@hostwerk/i18n';
import { Editor } from './Editor.jsx';
import { buildTemplate } from './templates.js';
import { makeLocale } from './format/locale.js';
import { saveBytes, csvToBook, CSV } from './io/formats.js';
import { sheetjsText, sheetjsGrid } from './io/sheetjs.js';
import { decodeText, detectDelimiter, parseCSV } from './io/csv.js';
import './style.css';

const docLang = () => (String((typeof document !== 'undefined' && document.documentElement.lang) || 'en').toLowerCase().startsWith('de') ? 'de' : 'en');

async function create({ ext = 'xlsx', template = 'blank', settings = {} } = {}) {
  const S = settings.sheets || {};
  const lang = docLang();
  const loc = makeLocale({ lang, decimal: S.decimal || 'auto', dateFormat: S.dateFormat || 'auto', functionNames: S.functionNames || 'auto' });
  const book = buildTemplate(template || 'blank', { loc, settings: S, lang });
  const e = String(ext || 'xlsx').toLowerCase();
  const { bytes } = await saveBytes(book, ['numbers', 'sxc', 'wk1', 'wks', 'qpw', 'xlsb'].includes(e) ? 'xlsx' : e, { getValues: null, settings: S });
  return bytes;
}

async function toText(bytes, ext) {
  const e = String(ext || '').toLowerCase();
  if (CSV.includes(e) || e === 'txt') { const { text } = decodeText(bytes); return text; }
  try { return sheetjsText(bytes, 3); } catch { return ''; }
}

/** small read-only preview: first sheet, first ~50×15 cells */
function Preview({ bytes, ext, name }) {
  const [data, setData] = useState(null);
  const [failed, setFailed] = useState(false);
  useEffect(() => {
    let off = false;
    (async () => {
      try {
        const e = String(ext || '').toLowerCase();
        let d;
        if (CSV.includes(e)) {
          const { text } = decodeText(bytes);
          const rows = parseCSV(text, { delimiter: e === 'csv' ? detectDelimiter(text) : '\t', maxRows: 50 }).map((r) => r.slice(0, 15).map((x) => (x === '' ? null : { t: x, n: /^-?[\d.,]+$/.test(x) })));
          d = { name: name || '', sheets: [], rows };
        } else d = sheetjsGrid(bytes, 50, 15, makeLocale({ lang: docLang() }));
        if (!off) setData(d);
      } catch { if (!off) setFailed(true); }
    })();
    return () => { off = true; };
  }, [bytes, ext, name]);
  if (failed) return <div className="of-sheet-pv small muted">{t('No preview for this file.')}</div>;
  if (!data) return <div className="of-sheet-pv small muted">{t('Loading …')}</div>;
  const cols = Math.max(1, ...data.rows.map((r) => r.length));
  const L = (i) => { let s = ''; let n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  return (
    <div className="of-sheet-pv">
      {data.sheets && data.sheets.length > 1 && <div className="small muted of-sheet-pv-head">{t('{n} sheets · showing “{name}”', { n: data.sheets.length, name: data.name })}</div>}
      <div className="of-sheet-pv-scroll">
        <table>
          <thead><tr><th />{Array.from({ length: cols }, (_, i) => <th key={i}>{L(i)}</th>)}</tr></thead>
          <tbody>{data.rows.map((r, i) => <tr key={i}><th>{i + 1}</th>{Array.from({ length: cols }, (_, j) => <td key={j} className={r[j] && r[j].n ? 'num' : ''}>{r[j] ? r[j].t : ''}</td>)}</tr>)}</tbody>
        </table>
      </div>
    </div>
  );
}

export default { kind: 'sheets', Editor, create, toText, Preview };
export { csvToBook };
