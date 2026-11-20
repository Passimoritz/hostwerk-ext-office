/*
 * Compare two versions of a workbook cell by cell (host.compare → the chunk's Compare export): summary per sheet,
 * a grid with added (green) / removed (red) / changed (yellow, old → new) cells and a list of all changes.
 * hostwerk-allow: tables – the diff grid shows the sheet itself (A/B/C, 1/2/3) like the editor's paper, not a Hostwerk list
 */
import { useEffect, useMemo, useState } from 'react';
import { t } from '@hostwerk/i18n';
import { Loading, Callout } from '@hostwerk/ui';
import { Keep, Raw } from '../../_shared/keep.jsx';
import { openBytes } from '../io/formats.js';
import { makeLocale } from '../format/locale.js';
import { colName, MAXC } from '../formula/refs.js';
import { toDisplay } from '../formula/parser.js';
import { diffBooks } from './diff.js';

export function Compare({ a, b, ext, host }) {
  const [res, setRes] = useState(null);
  const [err, setErr] = useState(null);
  const [cur, setCur] = useState(0);
  const loc = useMemo(() => makeLocale({ lang: (host && host.lang) || 'en' }), [host]);
  useEffect(() => {
    let off = false;
    (async () => {
      try {
        const e = String(ext || 'xlsx').toLowerCase();
        const [A, B] = await Promise.all([openBytes(a.bytes, e, { loc }), openBytes(b.bytes, e, { loc })]);
        const d = diffBooks(A, B);
        if (!off) { setRes({ A, B, d }); const i = d.sheets.findIndex((s) => s.status !== 'same'); setCur(i < 0 ? 0 : i); }
      } catch (x) { if (!off) setErr(x); }
    })();
    return () => { off = true; };
  }, [a, b, ext, loc]);
  if (err) return <Callout kind="err">{String(err.message || err)}</Callout>;
  if (!res) return <Loading />;
  const sh = res.d.sheets[cur];
  const total = res.d.sheets.reduce((n, s) => n + s.cells.length, 0);
  const byKey = new Map(sh.cells.map((x) => [x.r * MAXC + x.c, x]));
  const r2 = Math.min(200, Math.max(10, ...sh.cells.map((x) => x.r + 2))); const c2 = Math.min(40, Math.max(6, ...sh.cells.map((x) => x.c + 2)));
  const sheetA = res.A.sheetByName(sh.name); const sheetB = res.B.sheetByName(sh.name);
  const text = (r, c) => { const x = byKey.get(r * MAXC + c); if (x) return x.b ? x.b.text : x.a.text; return sheetB ? res.B.text(sheetB, r, c) : sheetA ? res.A.text(sheetA, r, c) : ''; };
  return (
    <div className="of-sheet-cmp">
      <div className="row of-sheet-cmp-head">
        <b><Keep>{a.label}</Keep></b><span className="muted">→</span><b><Keep>{b.label}</Keep></b>
        <span className="grow" />
        <span className="pill small">{total ? t('{n} changed cells', { n: total }) : t('No differences')}</span>
      </div>
      <div className="of-sheet-cmp-tabs" role="tablist">
        {res.d.sheets.map((s, i) => <Raw as="button" key={s.name} type="button" role="tab" aria-selected={i === cur} className={`btn ghost sm${i === cur ? ' is-active' : ''} is-${s.status}`} title={s.name} onClick={() => setCur(i)}>{`${s.name}${s.cells.length ? ` · ${s.cells.length}` : s.status === 'added' ? ' +' : s.status === 'removed' ? ' −' : ''}`}</Raw>)}
      </div>
      {sh.status === 'added' || sh.status === 'removed' ? <Callout kind="info">{sh.status === 'added' ? t('This sheet is new.') : t('This sheet was deleted.')}</Callout> : (
        <div className="of-sheet-cmp-body">
          <div className="of-sheet-cmp-grid"><table><thead><tr><th />{Array.from({ length: c2 }, (_, c) => <th key={c}>{colName(c)}</th>)}</tr></thead>
            <tbody>{Array.from({ length: r2 }, (_, r) => <tr key={r}><th>{r + 1}</th>{Array.from({ length: c2 }, (_, c) => { const x = byKey.get(r * MAXC + c); return <td key={c} className={x ? `is-${x.kind}` : ''} title={x ? `${x.a ? x.a.text : '∅'} → ${x.b ? x.b.text : '∅'}` : undefined}><Keep>{text(r, c)}</Keep></td>; })}</tr>)}</tbody></table></div>
          <ul className="of-sheet-cmp-list small">
            {sh.cells.slice(0, 500).map((x) => <li key={`${x.r}:${x.c}`} className={`is-${x.kind}`}><b className="mono">{colName(x.c)}{x.r + 1}</b> <span className="muted"><Keep>{x.a ? (x.a.f ? `=${toDisplay(x.a.f, loc)}` : x.a.text) : '∅'}</Keep></span> → <Keep>{x.b ? (x.b.f ? `=${toDisplay(x.b.f, loc)}` : x.b.text) : '∅'}</Keep></li>)}
          </ul>
        </div>
      )}
    </div>
  );
}
