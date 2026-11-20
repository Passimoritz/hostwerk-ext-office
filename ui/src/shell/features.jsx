/*
 * Office 1.1 – dialogs of the new server features, for the editor host (host.data/db/fields/forms/rules/components/
 * remote/compare) and the file actions in Files (encrypt, clean before sharing, publish, scheduled reports, rules,
 * forms, older states from Hostwerk's backups). Chunks never render these – they call host.* (EditorHost) and get a
 * Promise of the result (null = cancelled).
 *
 *   askFeature(type, props) → Promise<value|null>     one dialog at a time, shown by the FeatureHost mounted last
 *   <FeatureHost />                                    in the dialog host (sidebar.footer) and in every editor host
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  Database, Table2, Braces, ClipboardList, Bell, Package, Smartphone, Lock, Unlock, Eraser, Globe, CalendarClock, History, GitCompare,
  Check, Copy, Trash2, Plus, RotateCcw, Download, Search, Play, ShieldCheck,
} from 'lucide-react';
import { Modal, Btn, Field, Switch, Seg, Loading, Empty, Callout, Progress, toast, confirmDialog } from '@hostwerk/ui';
import { apiUrl } from '@hostwerk/api';
import { t, pl } from '@hostwerk/i18n';
import { office, isMissing, url as apiPath } from '../lib/api';
import { FORMATS, extOf } from '../lib/formats';
import { FileIcon, Keep, Raw, fmtDateTime, fmtWhen, copyText, useDebounced } from './common';
import { useTopHost } from './engine';
import { changed } from './store';
import { fmtBytes, splitExt } from './util';
import { FilePreviewBox } from './preview';
import { lineDiff } from './diff';

/* ---------------- bus ---------------- */
let cur = null;
let seq = 0;
let mounted = 0;
const subs = new Set();
const emit = () => subs.forEach((f) => f(cur));
/** Show one feature dialog → Promise of its result (null when cancelled or when no host is mounted) */
export function askFeature(type, props = {}) {
  if (!mounted) { toast(t('Open the file in Office for this.')); return Promise.resolve(null); }
  return new Promise((resolve) => {
    if (cur) cur.resolve(null);
    cur = { type, props, resolve, id: ++seq };
    emit();
  });
}
function closeFeature(v) { const c = cur; cur = null; emit(); if (c) c.resolve(v === undefined ? null : v); }
export function FeatureHost() {
  const [d, setD] = useState(cur);
  const top = useTopHost('features');
  useEffect(() => { mounted++; subs.add(setD); return () => { mounted--; subs.delete(setD); }; }, []);
  if (!d || !top) return null;
  const C = DIALOGS[d.type];
  if (!C) return null;
  return <C key={d.id} {...d.props} onDone={closeFeature} />;
}
const fail = (e) => (isMissing(e) ? t('This needs a newer Office server part – update Office (or Hostwerk) and try again.') : (e && e.message) || String(e));
/** QR code of a link: the server's SVG (x.qr) or GET /office/qr?text= */
export function QrImage({ value, svg, size = 180 }) {
  const src = svg ? (String(svg).startsWith('data:') ? svg : `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`) : apiUrl(apiPath('/qr', { text: value }));
  const [bad, setBad] = useState(false);
  if (bad) return null;
  return <Raw as="img" className="of-shell-qr" src={src} width={size} height={size} alt="" onError={() => setBad(true)} />;
}
function LinkBox({ link }) {
  return <div className="of-shell-link"><Check size={16} className="c-ok" aria-hidden="true" /><Raw as="input" className="input mono grow" readOnly value={link} onFocus={(e) => e.target.select()} aria-label={t('Link')} /><Btn className="sm" icon={Copy} onClick={() => copyText(link)}>Copy</Btn></div>;
}

/* ---------------- databases ---------------- */
function DbConfirm({ table, changes = [], onDone }) {
  const n = { update: 0, insert: 0, delete: 0 };
  for (const c of changes) n[c.op] = (n[c.op] || 0) + 1;
  const danger = n.delete > 0;
  return (
    <Modal size="dialog" title={t('Write {n} changes to “{table}”?', { n: changes.length, table })} subtitle="The database is changed right away – in one transaction" onClose={() => onDone(null)}
      footer={<><Btn onClick={() => onDone(null)}>Cancel</Btn><Btn className={danger ? 'danger' : 'primary'} icon={Database} onClick={() => onDone(true)}>{t('Write {n} changes', { n: changes.length })}</Btn></>}>
      <div className="stack">
        <div className="row-tight of-shell-wrap">
          {n.update > 0 && <span className="badge accent">{pl(n.update, '{n} changed row', '{n} changed rows')}</span>}
          {n.insert > 0 && <span className="badge ok">{pl(n.insert, '{n} new row', '{n} new rows')}</span>}
          {n.delete > 0 && <span className="badge err">{pl(n.delete, '{n} deleted row', '{n} deleted rows')}</span>}
        </div>
        <ul className="of-shell-rows of-shell-scroll">
          {changes.slice(0, 12).map((c, i) => (
            <li key={i} className="of-shell-row"><span className={`badge ${c.op === 'delete' ? 'err' : c.op === 'insert' ? 'ok' : ''}`}>{t(c.op === 'delete' ? 'delete' : c.op === 'insert' ? 'new' : 'change')}</span>
              <span className="small mono ellipsis grow"><Keep>{[c.key != null ? JSON.stringify(c.key) : '', c.values ? Object.entries(c.values).map(([k, v]) => `${k} = ${v}`).join(', ') : ''].filter(Boolean).join(' · ')}</Keep></span></li>
          ))}
        </ul>
        {changes.length > 12 && <p className="tiny muted of-shell-p">{t('… and {n} more', { n: changes.length - 12 })}</p>}
        {danger && <Callout kind="warn" dismissible={false}>Deleted rows are gone for good unless the database has its own backup.</Callout>}
      </div>
    </Modal>
  );
}
function DbPick({ onDone }) {
  const [servers, setServers] = useState(null);
  const [server, setServer] = useState('');
  const [tables, setTables] = useState(null);
  const [err, setErr] = useState('');
  const [q, setQ] = useState('');
  useEffect(() => { office.db.servers().then((l) => { setServers(l); if (l[0]) setServer(l[0].id || l[0].name); }, (e) => { setErr(fail(e)); setServers([]); }); }, []);
  useEffect(() => { if (!server) return; setTables(null); office.db.tables(server).then(setTables, (e) => { setErr(fail(e)); setTables([]); }); }, [server]);
  const shown = (tables || []).filter((x) => !q || `${x.db || ''}.${x.table || x.name}`.toLowerCase().includes(q.toLowerCase()));
  return (
    <Modal size="wide" title="Open a database table" subtitle="Rows of a Hostwerk database as a spreadsheet – changes are written back after you confirm" onClose={() => onDone(null)} footer={<Btn onClick={() => onDone(null)}>Cancel</Btn>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !servers ? <Loading inline /> : !servers.length ? <Empty icon={Database} title="No database yet" text="Create one under Databases – then its tables show up here." action={<a className="btn" href="#/databases">Open Databases</a>} /> : (
        <div className="of-shell-split">
          <ul className="of-shell-rows of-shell-scroll" role="listbox" aria-label={t('Servers')}>
            {servers.map((s) => { const id = s.id || s.name; return <li key={id} role="option" aria-selected={id === server} tabIndex={0} className={`of-shell-row is-click ${id === server ? 'active' : ''}`} onClick={() => setServer(id)} onKeyDown={(e) => { if (e.key === 'Enter') setServer(id); }}><Database size={14} aria-hidden="true" /><span className="small ellipsis grow"><Keep>{s.label || s.name || id}</Keep></span>{s.engine && <span className="tiny muted"><Keep>{s.engine}</Keep></span>}</li>; })}
          </ul>
          <div className="stack of-shell-gap-2">
            <div className="of-shell-search"><Search size={14} aria-hidden="true" /><input className="input sm" type="search" value={q} onChange={(e) => setQ(e.target.value)} placeholder={t('Filter tables …')} aria-label={t('Filter tables')} /></div>
            {!tables ? <Loading inline /> : !shown.length ? <p className="small muted of-shell-p">No tables.</p> : (
              <ul className="of-shell-rows of-shell-scroll">
                {shown.map((x) => <li key={`${x.db}.${x.table || x.name}`} className="of-shell-row is-click" tabIndex={0} onClick={() => onDone({ server, db: x.db, table: x.table || x.name })} onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.click(); }}>
                  <Table2 size={14} className="c-ok" aria-hidden="true" /><span className="small grow ellipsis"><b><Keep>{x.table || x.name}</Keep></b> <span className="muted"><Keep>{x.db || ''}</Keep></span></span>
                  {x.rows != null && <span className="tiny muted">{pl(x.rows, '{n} row', '{n} rows')}</span>}{x.primaryKey === null && <span className="badge" title={t('No primary key – read-only')}>{t('read-only')}</span>}
                </li>)}
              </ul>
            )}
          </div>
        </div>
      )}
    </Modal>
  );
}

/* ---------------- live data and fields ---------------- */
function ParamInputs({ params = [], value, onChange }) {
  if (!params.length) return null;
  return (
    <div className="of-shell-new-grid">
      {params.map((p) => (
        <Field key={p.id} label={p.label || p.id}>
          {Array.isArray(p.options) ? (
            <select className="input select" value={value[p.id] ?? ''} onChange={(e) => onChange({ ...value, [p.id]: e.target.value })}>
              <option value="">{t('– choose –')}</option>
              {p.options.map((o) => { const id = typeof o === 'object' ? o.id ?? o.value : o; return <Raw as="option" key={id} value={id}>{typeof o === 'object' ? o.label || id : o}</Raw>; })}
            </select>
          ) : <Raw as="input" className="input" type={p.type === 'number' ? 'number' : 'text'} value={value[p.id] ?? ''} onChange={(e) => onChange({ ...value, [p.id]: e.target.value })} />}
        </Field>
      ))}
    </div>
  );
}
function ResultPreview({ r }) {
  if (!r) return <Loading inline />;
  if (r.error) return <Callout kind="err" dismissible={false}>{r.error}</Callout>;
  if (r.value != null && !(r.rows || []).length) return <div className="of-shell-value"><Keep>{String(r.text ?? r.value)}</Keep></div>;
  const cols = r.columns || [];
  return (
    <div className="of-shell-scroll of-shell-tablebox">
      <table className="table small">
        <thead><tr>{cols.map((c) => <th key={c.id}><Keep>{c.label || c.id}</Keep></th>)}</tr></thead>
        <tbody>{(r.rows || []).slice(0, 6).map((row, i) => <tr key={i}>{row.map((v, j) => <td key={j}><Keep>{v == null ? '' : String(v)}</Keep></td>)}</tr>)}</tbody>
      </table>
      {(r.rows || []).length > 6 && <p className="tiny muted of-shell-p">{pl(r.rows.length, '{n} row in total', '{n} rows in total')}</p>}
    </div>
  );
}
function DataPick({ kind: want, onDone }) {
  const [list, setList] = useState(null);
  const [err, setErr] = useState('');
  const [sel, setSel] = useState(null);
  const [params, setParams] = useState({});
  const [res, setRes] = useState(null);
  const [q, setQ] = useState('');
  useEffect(() => { office.data.sources().then((l) => { const f = want ? l.filter((x) => x.kind === want) : l; setList(f); if (f[0]) setSel(f[0]); }, (e) => { setErr(fail(e)); setList([]); }); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const dp = useDebounced(JSON.stringify(params), 400);
  useEffect(() => { if (!sel) return; setRes(null); office.data.query(sel.id, params).then(setRes, (e) => setRes({ error: fail(e) })); }, [sel && sel.id, dp]); // eslint-disable-line react-hooks/exhaustive-deps
  const shown = (list || []).filter((x) => !q || String(x.label || x.id).toLowerCase().includes(q.toLowerCase()));
  return (
    <Modal size="wide" title="Insert live data" subtitle="Values from Hostwerk that update by themselves – servers, backups, tasks, websites …" onClose={() => onDone(null)}
      footer={<><Btn onClick={() => onDone(null)}>Cancel</Btn><Btn className="primary" icon={Plus} disabled={!sel} onClick={() => onDone({ source: sel.id, params, label: sel.label || sel.id, kind: sel.kind })}>Insert</Btn></>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !list ? <Loading inline /> : !list.length ? <Empty icon={Braces} title="No live data" text="Nothing to offer yet." /> : (
        <div className="of-shell-split">
          <div className="stack of-shell-gap-2">
            <div className="of-shell-search"><Search size={14} aria-hidden="true" /><input className="input sm" type="search" value={q} onChange={(e) => setQ(e.target.value)} placeholder={t('Search …')} aria-label={t('Search')} /></div>
            <ul className="of-shell-rows of-shell-scroll" role="listbox" aria-label={t('Sources')}>
              {shown.map((x) => <li key={x.id} role="option" aria-selected={sel && sel.id === x.id} tabIndex={0} className={`of-shell-row is-click ${sel && sel.id === x.id ? 'active' : ''}`} onClick={() => { setSel(x); setParams({}); }} onKeyDown={(e) => { if (e.key === 'Enter') { setSel(x); setParams({}); } }}>
                {x.kind === 'table' ? <Table2 size={14} aria-hidden="true" /> : <Braces size={14} aria-hidden="true" />}<span className="small ellipsis grow"><Keep>{x.label || x.id}</Keep></span></li>)}
            </ul>
          </div>
          <div className="stack of-shell-gap-2">
            {sel && <ParamInputs params={sel.params} value={params} onChange={setParams} />}
            <div className="of-shell-sec-title">Preview</div>
            <ResultPreview r={res} />
            {res && res.updatedAt && <span className="tiny muted">{t('as of {time}', { time: fmtDateTime(res.updatedAt) })}</span>}
          </div>
        </div>
      )}
    </Modal>
  );
}
function FieldPick({ onDone }) {
  const [list, setList] = useState(null);
  const [err, setErr] = useState('');
  const [q, setQ] = useState('');
  const [sel, setSel] = useState(null);
  const [val, setVal] = useState(null);
  useEffect(() => { office.fields.list().then(setList, (e) => { setErr(fail(e)); setList([]); }); }, []);
  useEffect(() => { if (!sel) return; setVal(null); office.fields.value(sel.id).then(setVal, (e) => setVal({ error: fail(e) })); }, [sel && sel.id]); // eslint-disable-line react-hooks/exhaustive-deps
  const groups = useMemo(() => { const m = new Map(); for (const f of (list || []).filter((x) => !q || `${x.label} ${x.group || ''}`.toLowerCase().includes(q.toLowerCase()))) { const g = f.group || ''; if (!m.has(g)) m.set(g, []); m.get(g).push(f); } return [...m]; }, [list, q]);
  return (
    <Modal size="wide" title="Insert field" subtitle="A field shows a live value – date, page number, server status, last backup …" onClose={() => onDone(null)}
      footer={<><Btn onClick={() => onDone(null)}>Cancel</Btn><Btn className="primary" icon={Plus} disabled={!sel} onClick={() => onDone({ id: sel.id, params: {}, label: sel.label })}>Insert</Btn></>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !list ? <Loading inline /> : (
        <div className="of-shell-split">
          <div className="stack of-shell-gap-2">
            <div className="of-shell-search"><Search size={14} aria-hidden="true" /><input className="input sm" type="search" autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder={t('Search fields …')} aria-label={t('Search fields')} /></div>
            <div className="of-shell-scroll stack of-shell-gap-2">
              {groups.map(([g, fs]) => (
                <div key={g || '-'}>
                  {g && <div className="of-shell-sec-title"><Keep>{t(g)}</Keep></div>}
                  <ul className="of-shell-rows">{fs.map((f) => <li key={f.id} tabIndex={0} className={`of-shell-row is-click ${sel && sel.id === f.id ? 'active' : ''}`} onClick={() => setSel(f)} onDoubleClick={() => onDone({ id: f.id, params: {}, label: f.label })} onKeyDown={(e) => { if (e.key === 'Enter') onDone({ id: f.id, params: {}, label: f.label }); }}><Braces size={14} aria-hidden="true" /><span className="small grow ellipsis">{t(f.label || f.id)}</span></li>)}</ul>
                </div>
              ))}
            </div>
          </div>
          <div className="stack of-shell-gap-2"><div className="of-shell-sec-title">Value now</div>{!sel ? <p className="small muted of-shell-p">Pick a field on the left.</p> : <ResultPreview r={val} />}</div>
        </div>
      )}
    </Modal>
  );
}

/* ---------------- forms ---------------- */
function FormsDialog({ file, sheets = [], columns = [], create = false, onDone }) {
  const [list, setList] = useState(null);
  const [err, setErr] = useState('');
  const [mode, setMode] = useState(create ? 'new' : 'list');
  const [title, setTitle] = useState(splitExt(file.name || '').base);
  const [sheet, setSheet] = useState(sheets[0] || '');
  const [cols, setCols] = useState(columns.map((c) => (typeof c === 'string' ? { id: c, label: c, on: true } : { ...c, on: c.on !== false })));
  const [pub, setPub] = useState(true);
  const [days, setDays] = useState('30');
  const [made, setMade] = useState(null);
  const load = () => office.forms.list(file).then(setList, (e) => { setErr(fail(e)); setList([]); });
  useEffect(() => { load(); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const make = async () => {
    try {
      const r = await office.forms.create(file, { sheet, title, columns: cols.filter((c) => c.on).map(({ on, ...c }) => c), public: { enabled: pub, expires: days === '0' ? null : Date.now() + Number(days) * 86400000 } });
      setMade(r); setMode('list'); load();
      toast.success('Form created');
    } catch (e) { toast.error(fail(e)); }
  };
  const remove = async (x) => { if (!(await confirmDialog({ title: 'Delete this form?', text: 'The link stops working. Answers already in the sheet stay.', confirm: 'Delete form', danger: true }))) return; await office.forms.remove(x.id).catch((e) => toast.error(fail(e))); load(); };
  return (
    <Modal size="wide" title={t('Forms of “{name}”', { name: file.name })} subtitle="People fill in a form – every answer becomes a new row in the sheet" onClose={() => onDone(made)}
      footer={mode === 'new' ? <><Btn onClick={() => setMode('list')}>Cancel</Btn><Btn className="primary" icon={ClipboardList} disabled={!title.trim() || !cols.some((c) => c.on)} onClick={make}>Create form</Btn></> : <><Btn onClick={() => onDone(made)}>Close</Btn><Btn className="primary" icon={Plus} onClick={() => setMode('new')}>New form</Btn></>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : mode === 'new' ? (
        <div className="stack">
          <div className="of-shell-new-grid">
            <Field label="Title"><Raw as="input" className="input" value={title} onChange={(e) => setTitle(e.target.value)} /></Field>
            {sheets.length > 0 && <Field label="Sheet"><select className="input select" value={sheet} onChange={(e) => setSheet(e.target.value)}>{sheets.map((s) => <Raw as="option" key={s} value={s}>{s}</Raw>)}</select></Field>}
          </div>
          {cols.length > 0 && <Field label="Questions" hint="One question per column – untick what people should not fill in.">
            <div className="of-shell-chips">{cols.map((c, i) => <label key={c.id} className={`chip-toggle ${c.on ? 'on' : ''}`}><input type="checkbox" checked={c.on} onChange={() => setCols(cols.map((x, j) => (j === i ? { ...x, on: !x.on } : x)))} /><Keep>{c.label || c.id}</Keep></label>)}</div>
          </Field>}
          <div className="setting-row"><div className="txt"><b>Public link</b><span>Anyone with the link can answer – no Hostwerk login needed.</span></div><Switch checked={pub} onChange={setPub} /></div>
          {pub && <Field label="Open for"><Seg value={days} onChange={setDays} options={[{ id: '7', label: '7 days' }, { id: '30', label: '30 days' }, { id: '0', label: 'No end' }]} /></Field>}
        </div>
      ) : (
        <div className="stack">
          {made && made.url && <div className="of-shell-qrbox hw-pop"><QrImage value={made.url} svg={made.qr} /><div className="stack of-shell-gap-2 grow"><b className="small">Form is ready – share the link or the QR code</b><LinkBox link={made.url} /></div></div>}
          {!list ? <Loading inline /> : !list.length ? <Empty icon={ClipboardList} title="No form yet" text="A form turns this sheet into a sign-up list, survey or order form." /> : (
            <ul className="of-shell-rows">{list.map((x) => (
              <li key={x.id} className="of-shell-row"><ClipboardList size={14} className="c-ok" aria-hidden="true" />
                <div className="grow of-shell-min0"><b className="small ellipsis"><Keep>{x.title || x.id}</Keep></b><div className="tiny muted">{[x.sheet && t('sheet “{s}”', { s: x.sheet }), x.responses != null && pl(x.responses, '{n} answer', '{n} answers'), x.public && x.public.enabled ? t('public') : t('Hostwerk only')].filter(Boolean).join(' · ')}</div></div>
                {x.url && <Btn className="xs ghost icon" icon={Copy} title="Copy link" aria-label="Copy link" onClick={() => copyText(x.url)} />}
                {x.url && <Btn className="xs ghost icon" icon={Smartphone} title="QR code" aria-label="QR code" onClick={() => setMade(x)} />}
                <Btn className="xs ghost icon" icon={Trash2} title="Delete form" aria-label="Delete form" onClick={() => remove(x)} />
              </li>))}</ul>
          )}
        </div>
      )}
    </Modal>
  );
}

/* ---------------- rules & notifications ---------------- */
const OPS = [['>', 'greater than'], ['>=', 'at least'], ['<', 'less than'], ['<=', 'at most'], ['=', 'equals'], ['!=', 'is not'], ['changed', 'changes'], ['contains', 'contains'], ['empty', 'is empty']];
const CHANNELS = [['hostwerk', 'Hostwerk notification'], ['mail', 'E-mail'], ['discord', 'Discord'], ['ntfy', 'ntfy (phone)']];
const newRule = (sheet = '', range = '') => ({ id: `r${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`, when: { sheet, range, op: '>', value: '' }, then: { notify: { channel: 'hostwerk', text: '' } } });
function RulesDialog({ file, sheet = '', range = '', onDone }) {
  const [rules, setRules] = useState(null);
  const [err, setErr] = useState('');
  const [dirty, setDirty] = useState(false);
  useEffect(() => { office.rules.get(file).then((l) => setRules(l.length || !range ? l : [newRule(sheet, range)]), (e) => { setErr(fail(e)); setRules([]); }); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const upd = (i, path, v) => { setDirty(true); setRules((rs) => rs.map((r, j) => { if (j !== i) return r; const c = JSON.parse(JSON.stringify(r)); let o = c; const ks = path.split('.'); for (const k of ks.slice(0, -1)) o = o[k] || (o[k] = {}); o[ks[ks.length - 1]] = v; return c; })); };
  const save = async () => { try { await office.rules.set(file, rules.map(({ when, then, id }) => ({ id, when: { ...when, cell: when.range && !when.range.includes(':') ? when.range : undefined }, then }))); toast.success('Rules saved'); onDone(rules); } catch (e) { toast.error(fail(e)); } };
  return (
    <Modal size="wide" dirty={dirty} title={t('Rules of “{name}”', { name: file.name })} subtitle="When a value meets a condition, Hostwerk notifies you – checked on every save and every few minutes" onClose={() => onDone(null)}
      footer={<><Btn className="ghost" icon={Plus} onClick={() => { setDirty(true); setRules((r) => [...(r || []), newRule(sheet, range)]); }}>Add rule</Btn><span className="grow" /><Btn onClick={() => onDone(null)}>Cancel</Btn><Btn className="primary" disabled={!rules} onClick={save}>Save</Btn></>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !rules ? <Loading inline /> : !rules.length ? <Empty icon={Bell} title="No rules yet" text="Example: when B2 (stock) is less than 5, send me a notification." /> : (
        <div className="stack">
          {rules.map((r, i) => (
            <div key={r.id} className="of-shell-rule">
              <span className="of-shell-rule-n" aria-hidden="true">{i + 1}</span>
              <div className="of-shell-rule-body">
                <div className="of-shell-rule-line"><b className="small">When</b>
                  <Raw as="input" className="input sm of-shell-w-sm" value={r.when.sheet || ''} placeholder={t('Sheet')} aria-label={t('Sheet')} onChange={(e) => upd(i, 'when.sheet', e.target.value)} />
                  <Raw as="input" className="input sm mono of-shell-w-xs" value={r.when.range || r.when.cell || ''} placeholder="B2" aria-label={t('Cell or range')} onChange={(e) => upd(i, 'when.range', e.target.value.toUpperCase())} />
                  <select className="input select sm" value={r.when.op} onChange={(e) => upd(i, 'when.op', e.target.value)} aria-label={t('Condition')}>{OPS.map(([id, l]) => <option key={id} value={id}>{t(l)}</option>)}</select>
                  {!['changed', 'empty'].includes(r.when.op) && <Raw as="input" className="input sm of-shell-w-sm" value={r.when.value ?? ''} placeholder={t('Value')} aria-label={t('Value')} onChange={(e) => upd(i, 'when.value', e.target.value)} />}
                </div>
                <div className="of-shell-rule-line"><b className="small">then</b>
                  <select className="input select sm" value={r.then.notify.channel} onChange={(e) => upd(i, 'then.notify.channel', e.target.value)} aria-label={t('Channel')}>{CHANNELS.map(([id, l]) => <option key={id} value={id}>{t(l)}</option>)}</select>
                  <Raw as="input" className="input sm grow" value={r.then.notify.text || ''} placeholder={t('Message, e.g. “Stock of {cell} is low: {value}”')} aria-label={t('Message')} onChange={(e) => upd(i, 'then.notify.text', e.target.value)} />
                </div>
              </div>
              <Btn className="xs ghost icon" icon={Trash2} title="Remove rule" aria-label="Remove rule" onClick={() => { setDirty(true); setRules(rules.filter((_, j) => j !== i)); }} />
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

/* ---------------- components on demand ---------------- */
function ComponentDialog({ id, label, reason, onDone }) {
  const [st, setSt] = useState(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState('');
  const poll = useRef(null);
  const load = () => office.components.list().then((l) => {
    const c = l.find((x) => x.id === id || (Array.isArray(x.aliases) && x.aliases.includes(id))) || { id, label };
    setSt(c);
    if (c.state === 'ready' || c.installed === true) { clearInterval(poll.current); onDone(c); }
    else if (c.state === 'error' && poll.current) { clearInterval(poll.current); poll.current = null; setBusy(false); setErr(c.error ? t(c.error) : t('The add-on could not be installed.')); }
    return c;
  }, (e) => { setErr(fail(e)); });
  useEffect(() => { load(); return () => clearInterval(poll.current); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const install = async () => { setBusy(true); setErr(''); try { await office.components.install(id); poll.current = setInterval(load, 1000); } catch (e) { setBusy(false); setErr(fail(e)); } };
  const pct = st && typeof st.progress === 'number' ? st.progress : st && st.progress && typeof st.progress.percent === 'number' ? st.progress.percent : null;
  const dl = st && (st.download || st.size);
  return (
    <Modal size="dialog" title={t('Download “{name}” once?', { name: t((st && st.label) || label || id) })} onClose={() => onDone(null)}
      footer={<><Btn onClick={() => onDone(null)}>{busy ? 'In the background' : 'Not now'}</Btn>{!busy && <Btn className="primary" icon={Download} disabled={!st} onClick={install}>{dl ? t('Download ({size})', { size: fmtBytes(dl) }) : t('Download')}</Btn>}</>}>
      <div className="stack">
        {reason && <p className="small of-shell-p">{reason}</p>}
        <p className="small muted of-shell-p">{t('It is too big for the Office package, so Office fetches it once from a fixed address and checks its fingerprint. It stays on this PC and works offline afterwards.')}</p>
        {busy && <div className="stack of-shell-gap-2"><Progress value={pct || 0} indeterminate={pct == null} /><span className="small muted">{(st && st.progress && st.progress.text) || t('Downloading …')}{pct != null ? ` · ${Math.round(pct)} %` : ''}</span></div>}
        {err && <Callout kind="err" dismissible={false}>{err}</Callout>}
      </div>
    </Modal>
  );
}

/* ---------------- remote control (slide show) ---------------- */
function RemoteDialog({ session, onDone }) {
  const [peers, setPeers] = useState(session.peers || 0);
  useEffect(() => (session.on ? session.on('peers', (n) => setPeers(typeof n === 'number' ? n : (n && n.count) || 0)) : undefined), [session]);
  return (
    <Modal size="dialog" title="Phone as remote control" subtitle="Scan the code with the phone's camera – next/previous, notes, timer and laser pointer" onClose={() => onDone(null)}
      footer={<><Btn className="ghost" onClick={() => { if (session.stop) session.stop(); onDone(false); }}>Stop remote control</Btn><Btn className="primary" onClick={() => onDone(true)}>Done</Btn></>}>
      <div className="of-shell-qrbox">
        <QrImage value={session.url} svg={session.qr} size={200} />
        <div className="stack of-shell-gap-2 grow">
          <span className={`badge ${peers ? 'ok' : ''}`}>{peers ? pl(peers, '{n} phone connected', '{n} phones connected') : t('Waiting for the phone …')}</span>
          {session.url && <LinkBox link={session.url} />}
          <span className="tiny muted">The phone must reach this PC (same network, or Hostwerk's remote access).</span>
        </div>
      </div>
    </Modal>
  );
}

/* ---------------- encrypt / decrypt with Hostwerk's vault ---------------- */
function VaultDialog({ entry, mode = 'encrypt', onDone }) {
  const [st, setSt] = useState(null);
  const [busy, setBusy] = useState(false);
  useEffect(() => { office.vault.state().then(setSt, (e) => setSt({ error: fail(e) })); }, []);
  const go = async () => {
    setBusy(true);
    try { const r = await (mode === 'encrypt' ? office.vault.encrypt(entry) : office.vault.decrypt(entry)); changed(); toast.success(mode === 'encrypt' ? t('“{name}” is encrypted', { name: entry.name }) : t('“{name}” is decrypted', { name: entry.name })); onDone(r && (r.entry || r)); } catch (e) { setBusy(false); toast.error(fail(e)); }
  };
  const enc = mode === 'encrypt';
  return (
    <Modal size="dialog" title={t(enc ? 'Encrypt “{name}”?' : 'Decrypt “{name}”?', { name: entry.name })} onClose={() => onDone(null)}
      footer={<><Btn onClick={() => onDone(null)}>Cancel</Btn><Btn className="primary" icon={enc ? Lock : Unlock} busy={busy} disabled={!st || st.error || st.available === false} onClick={go}>{enc ? 'Encrypt' : 'Decrypt'}</Btn></>}>
      <div className="stack">
        <div className="of-shell-hero-mini"><ShieldCheck size={22} aria-hidden="true" /><span className="small">{enc ? t('The file is locked with the key of your Hostwerk vault (AES-256). Only this Hostwerk – unlocked – can open it; it is useless for anyone who copies it.') : t('The file is stored readable again – everyone with access to the folder can open it.')}</span></div>
        {!st ? <Loading inline /> : st.error ? <Callout kind="err" dismissible={false}>{st.error}</Callout> : st.available === false ? <Callout kind="warn" dismissible={false}>Set up the vault first (Passwords) – then files can be encrypted with it.</Callout> : st.unlocked === false ? <Callout kind="info" dismissible={false}>The vault is locked – Hostwerk asks you to unlock it.</Callout> : null}
        {enc && <p className="tiny muted of-shell-p">Versions and the search index of this file are removed, thumbnails too. Opening it in Office decrypts it in memory only.</p>}
      </div>
    </Modal>
  );
}

/* ---------------- clean before sharing ---------------- */
function CleanDialog({ entry, onDone }) {
  const [items, setItems] = useState(null);
  const [off, setOff] = useState(new Set());
  const [err, setErr] = useState('');
  const [busy, setBusy] = useState(false);
  const [copy, setCopy] = useState(true);
  useEffect(() => { office.clean.preview(entry).then((r) => setItems((r && r.items) || []), (e) => { setErr(fail(e)); setItems([]); }); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const go = async () => {
    setBusy(true);
    try { const r = await office.clean.run(entry, items.filter((x) => !off.has(x.id)).map((x) => x.id), { copy }); changed(); toast.success(copy ? t('Clean copy saved as “{name}”', { name: (r && r.name) || '' }) : t('“{name}” is cleaned', { name: entry.name })); onDone(r); } catch (e) { setBusy(false); toast.error(fail(e)); }
  };
  const chosen = (items || []).filter((x) => !off.has(x.id));
  return (
    <Modal size="wide" title={t('Clean “{name}” before sharing', { name: entry.name })} subtitle="Hidden things a recipient should not see – Office shows what it found" onClose={() => onDone(null)}
      footer={<><Btn onClick={() => onDone(null)}>Cancel</Btn><Btn className="primary" icon={Eraser} busy={busy} disabled={!items || !chosen.length} onClick={go}>{copy ? t('Save a clean copy') : t('Remove {n} things', { n: chosen.length })}</Btn></>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !items ? <Loading inline text="Looking into the file …" /> : !items.length ? <Empty icon={ShieldCheck} title="Nothing hidden found" text="No author names, comments, tracked changes, hidden sheets or metadata worth removing." /> : (
        <div className="stack">
          <ul className="of-shell-rows">
            {items.map((x) => (
              <li key={x.id} className="of-shell-row is-click" onClick={() => setOff((s) => { const n = new Set(s); if (n.has(x.id)) n.delete(x.id); else n.add(x.id); return n; })}>
                <input type="checkbox" checked={!off.has(x.id)} readOnly tabIndex={-1} aria-label={t(x.label)} />
                <div className="grow of-shell-min0"><b className="small">{t(x.label)}{x.count != null ? ` (${x.count})` : ''}</b>{x.detail && <div className="tiny muted ellipsis"><Keep>{Array.isArray(x.detail) ? x.detail.join(', ') : String(x.detail)}</Keep></div>}</div>
              </li>
            ))}
          </ul>
          <div className="setting-row"><div className="txt"><b>Keep the original</b><span>Recommended – the clean file is saved next to it as “… (clean)”.</span></div><Switch checked={copy} onChange={setCopy} /></div>
        </div>
      )}
    </Modal>
  );
}

/* ---------------- publish as a web page ---------------- */
function PublishDialog({ entry, onDone }) {
  const [targets, setTargets] = useState(null);
  const [pubs, setPubs] = useState([]);
  const [target, setTarget] = useState('');
  const [title, setTitle] = useState(splitExt(entry.name).base);
  const [path, setPath] = useState(splitExt(entry.name).base.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'page');
  const [err, setErr] = useState('');
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(null);
  const load = () => office.publish.list(entry).then(setPubs, () => setPubs([]));
  useEffect(() => { office.publish.targets().then((l) => { setTargets(l); if (l[0]) setTarget(l[0].id); }, (e) => { setErr(fail(e)); setTargets([]); }); load(); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const go = async () => { setBusy(true); try { const r = await office.publish.run(entry, { target, title, path }); setDone(r); load(); toast.success('Published'); } catch (e) { toast.error(fail(e)); } setBusy(false); };
  return (
    <Modal size="wide" title={t('Publish “{name}” as a web page', { name: entry.name })} subtitle="A clean, read-only page on one of your websites – updated whenever you publish again" onClose={() => onDone(done)}
      footer={<><Btn onClick={() => onDone(done)}>Close</Btn><Btn className="primary" icon={Globe} busy={busy} disabled={!target || !path.trim()} onClick={go}>{pubs.some((p) => p.target === target) ? 'Publish again' : 'Publish'}</Btn></>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !targets ? <Loading inline /> : !targets.length ? <Empty icon={Globe} title="No website yet" text="Create a website in Hostwerk (Web server) – then documents can be published on it." action={<a className="btn" href="#/web">Open Web server</a>} /> : (
        <div className="stack">
          {done && done.url && <LinkBox link={done.url} />}
          <Field label="Website">
            <div className="of-shell-pickgrid" role="radiogroup" aria-label={t('Website')}>
              {targets.map((x) => <button key={x.id} type="button" role="radio" aria-checked={target === x.id} className={`of-shell-pickcard ${target === x.id ? 'on' : ''}`} onClick={() => setTarget(x.id)}><Globe size={16} aria-hidden="true" /><span className="stack of-shell-gap-0 of-shell-min0"><b className="small ellipsis"><Keep>{x.label || x.id}</Keep></b>{x.url && <span className="tiny muted ellipsis"><Keep>{x.url}</Keep></span>}</span></button>)}
            </div>
          </Field>
          <div className="of-shell-new-grid">
            <Field label="Page title"><Raw as="input" className="input" value={title} onChange={(e) => setTitle(e.target.value)} /></Field>
            <Field label="Address" hint="Part of the address after the website, e.g. “price-list”."><Raw as="input" className="input mono" value={path} onChange={(e) => setPath(e.target.value.replace(/[^\w-/]/g, '-'))} /></Field>
          </div>
          {pubs.length > 0 && <div><div className="of-shell-sec-title">Published</div><ul className="of-shell-rows">{pubs.map((p) => (
            <li key={p.id} className="of-shell-row"><Globe size={14} className="muted" aria-hidden="true" /><span className="small grow ellipsis"><Keep>{p.url || p.target}</Keep></span><span className="tiny muted">{fmtWhen(p.updatedAt)}</span>
              {p.url && <Btn className="xs ghost icon" icon={Copy} title="Copy link" aria-label="Copy link" onClick={() => copyText(p.url)} />}
              <Btn className="xs ghost icon" icon={Trash2} title="Take the page offline" aria-label="Take the page offline" onClick={async () => { if (await confirmDialog({ title: 'Take the page offline?', text: 'The web page is removed. The document stays.', confirm: 'Take offline', danger: true })) { await office.publish.remove(p.id).catch((e) => toast.error(fail(e))); load(); } }} />
            </li>))}</ul></div>}
        </div>
      )}
    </Modal>
  );
}

/* ---------------- scheduled reports ---------------- */
const WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
function reportFormats(ext) {
  const kind = (FORMATS[ext] || {}).kind;
  if (kind === 'sheets') return ['pdf', 'xlsx', 'csv'];
  if (kind === 'slides') return ['pdf', 'pptx'];
  return ['pdf', 'docx'];
}
// report recipients: Hostwerk's notification, every alert channel, e-mail addresses (Hostwerk's mail relay), a Discord
// webhook, an ntfy address or topic – an empty address uses the one set up under Live → Alerts
const REPORT_CHANNELS = [['hostwerk', 'Hostwerk notification'], ['alerts', 'All alert channels'], ['mail', 'E-mail'], ['discord', 'Discord'], ['ntfy', 'ntfy (phone)']];
const EVERY_UI = { daily: 'day', weekly: 'week', monthly: 'month', hourly: 'day' };
/** a stored report → the dialog's shape (the server keeps every: daily|weekly|monthly and time) */
export function reportForEdit(r) {
  const s = r.schedule || {};
  return { ...r, schedule: { ...s, every: EVERY_UI[s.every] || s.every || 'week', at: s.at || s.time || '08:00', weekday: s.weekday != null ? s.weekday : 1, day: s.day || 1 }, recipients: (r.recipients && r.recipients.length ? r.recipients : [{ channel: r.channel === 'all' ? 'alerts' : r.channel || 'mail', to: '' }]).map((x) => ({ channel: x.channel === 'all' ? 'alerts' : x.channel, to: x.to || '' })) };
}
function ReportDialog({ entry, onDone }) {
  const fmts = reportFormats(extOf(entry.name));
  const blank = { schedule: { every: 'week', at: '08:00', weekday: 1, day: 1 }, format: fmts[0], recipients: [{ channel: 'mail', to: '' }] };
  const [list, setList] = useState(null);
  const [edit, setEdit] = useState(null);
  const [err, setErr] = useState('');
  const load = () => office.reports.list(entry).then((l) => { setList(l.map(reportForEdit)); if (!l.length) setEdit(blank); }, (e) => { setErr(fail(e)); setList([]); });
  const runNow = async (r) => {
    try {
      const res = await office.reports.run(r.id);
      if (res && res.ok === false && !res.running) toast.error(res.error || t('The report could not be sent.'));
      else if (res && res.error) toast(t('Sent – but not to everyone: {error}', { error: res.error }));
      else toast.success('Report sent');
      load();
    } catch (e) { toast.error(fail(e)); }
  };
  useEffect(() => { load(); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const save = async () => { try { await office.reports.save(entry, edit); toast.success('Report scheduled'); setEdit(null); load(); } catch (e) { toast.error(fail(e)); } };
  const set = (k, v) => setEdit((x) => ({ ...x, [k]: v }));
  const setS = (k, v) => setEdit((x) => ({ ...x, schedule: { ...x.schedule, [k]: v } }));
  const when = (r) => { const s = r.schedule || {}; return s.every === 'day' ? t('every day at {at}', { at: s.at }) : s.every === 'month' ? t('on day {d} of every month at {at}', { d: s.day || 1, at: s.at }) : t('every {day} at {at}', { day: t(WEEKDAYS[((s.weekday || 1) + 6) % 7]), at: s.at }); };
  return (
    <Modal size="wide" title={t('Scheduled report – {name}', { name: entry.name })} subtitle="Office sends the file on a schedule – as PDF or in its own format – to everyone you list here" onClose={() => onDone(null)}
      footer={edit ? <><Btn onClick={() => (list && list.length ? setEdit(null) : onDone(null))}>Cancel</Btn><Btn className="primary" icon={CalendarClock} disabled={!edit.recipients.length} onClick={save}>Save schedule</Btn></> : <><Btn onClick={() => onDone(null)}>Close</Btn><Btn className="primary" icon={Plus} onClick={() => setEdit(blank)}>New schedule</Btn></>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !list ? <Loading inline /> : edit ? (
        <div className="stack">
          <Field label="How often"><Seg value={edit.schedule.every} onChange={(v) => setS('every', v)} options={[{ id: 'day', label: 'Daily' }, { id: 'week', label: 'Weekly' }, { id: 'month', label: 'Monthly' }]} /></Field>
          <div className="of-shell-new-grid">
            {edit.schedule.every === 'week' && <Field label="Day"><select className="input select" value={edit.schedule.weekday} onChange={(e) => setS('weekday', Number(e.target.value))}>{WEEKDAYS.map((d, i) => <option key={d} value={(i + 1) % 7}>{t(d)}</option>)}</select></Field>}
            {edit.schedule.every === 'month' && <Field label="Day of the month"><input className="input of-shell-num-in" type="number" min="1" max="28" value={edit.schedule.day} onChange={(e) => setS('day', Math.max(1, Math.min(28, Number(e.target.value) || 1)))} /></Field>}
            <Field label="Time"><input className="input of-shell-num-in" type="time" value={edit.schedule.at} onChange={(e) => setS('at', e.target.value)} /></Field>
            <Field label="Format"><Seg value={edit.format} onChange={(v) => set('format', v)} options={fmts.map((f) => ({ id: f, label: f.toUpperCase() }))} /></Field>
          </div>
          <Field label="Recipients">
            <div className="stack of-shell-gap-2">
              {edit.recipients.map((r, i) => (
                <div key={i} className="of-shell-rcpt">
                  <select className="input select sm" value={r.channel} onChange={(e) => set('recipients', edit.recipients.map((x, j) => (j === i ? { ...x, channel: e.target.value } : x)))} aria-label={t('Channel')}>{REPORT_CHANNELS.map(([id, l]) => <option key={id} value={id}>{t(l)}</option>)}</select>
                  {['mail', 'discord', 'ntfy'].includes(r.channel) && <Raw as="input" className="input sm grow" value={r.to} type={r.channel === 'mail' ? 'text' : 'url'} placeholder={r.channel === 'mail' ? t('name@example.org, … – empty: the alert address') : r.channel === 'ntfy' ? t('Topic or https://ntfy.sh/topic – empty: the alert address') : t('Webhook address – empty: the alert webhook')} aria-label={t('Recipient')} title={t('Recipient')} onChange={(e) => set('recipients', edit.recipients.map((x, j) => (j === i ? { ...x, to: e.target.value } : x)))} />}
                  {['hostwerk', 'alerts'].includes(r.channel) && <span className="small muted grow">{r.channel === 'alerts' ? t('Everything set up under Live → Alerts') : t('A notification in Hostwerk')}</span>}
                  {edit.recipients.length > 1 && <Btn className="xs ghost icon" icon={Trash2} title="Remove" aria-label="Remove" onClick={() => set('recipients', edit.recipients.filter((_, j) => j !== i))} />}
                </div>
              ))}
              <div><Btn className="sm ghost" icon={Plus} onClick={() => set('recipients', [...edit.recipients, { channel: 'mail', to: '' }])}>Add recipient</Btn></div>
            </div>
          </Field>
          <p className="tiny muted of-shell-p">{t('E-mails go out through Hostwerk’s mail relay (E-mail → Forwarding) – without one they land in Hostwerk’s mail catcher. Formats other than the file’s own need a converter (Office engine or Microsoft Office).')}</p>
        </div>
      ) : !list.length ? <Empty icon={CalendarClock} title="No schedule yet" /> : (
        <ul className="of-shell-rows">{list.map((r) => (
          <li key={r.id} className="of-shell-row"><CalendarClock size={14} className="muted" aria-hidden="true" />
            <div className="grow of-shell-min0"><b className="small">{when(r)} · {String(r.format || 'pdf').toUpperCase()}</b><div className="tiny muted ellipsis">{(r.recipients || []).map((x, k) => <span key={k}>{k ? ', ' : ''}{x.to ? <Keep>{x.to}</Keep> : t((REPORT_CHANNELS.find((c) => c[0] === x.channel) || [0, x.channel])[1])}</span>)}{r.nextRun ? ` · ${t('next: {when}', { when: fmtDateTime(r.nextRun) })}` : ''}</div>{r.lastError && <Raw as="div" className="tiny of-shell-err ellipsis" title={r.lastError}><Keep>{r.lastError}</Keep></Raw>}</div>
            <Btn className="xs ghost icon" icon={Play} title="Send now" aria-label="Send now" onClick={() => runNow(r)} />
            <Btn className="xs ghost" onClick={() => setEdit(r)}>Edit</Btn>
            <Btn className="xs ghost icon" icon={Trash2} title="Delete schedule" aria-label="Delete schedule" onClick={async () => { await office.reports.remove(r.id).catch((e) => toast.error(fail(e))); load(); }} />
          </li>))}</ul>
      )}
    </Modal>
  );
}

/* ---------------- older states from Hostwerk's backups (time machine) ---------------- */
function BackupsDialog({ entry, onCompare, onDone }) {
  const [list, setList] = useState(null);
  const [err, setErr] = useState('');
  const [cur, setCur] = useState(null);
  const [bytes, setBytes] = useState(null);
  useEffect(() => { office.backups.list(entry).then((l) => { const s = [...l].sort((a, b) => b.time - a.time); setList(s); if (s[0]) setCur(s[0]); }, (e) => { setErr(fail(e)); setList([]); }); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { setBytes(null); if (cur) office.backups.bytes(entry, cur.id).then(setBytes, (e) => toast.error(fail(e))); }, [cur && cur.id]); // eslint-disable-line react-hooks/exhaustive-deps
  const restore = async () => {
    if (!(await confirmDialog({ title: t('Restore the state of {when}?', { when: fmtDateTime(cur.time) }), text: 'The current content is kept as a version, so you can go back.', confirm: 'Restore' }))) return;
    try { await office.backups.restore(entry, cur.id); changed(); toast.success('Restored from the backup'); onDone(true); } catch (e) { toast.error(fail(e)); }
  };
  const asCopy = async () => {
    const { base, ext } = splitExt(entry.name);
    const name = `${base} (${fmtDateTime(cur.time).replace(/[/:.]/g, '-')}).${ext}`;
    const dirOf = entry.path.split('/').slice(0, -1).join('/');
    const e = await office.create(entry.root ? { root: entry.root, dir: dirOf } : dirOf, name, bytes);
    changed(); toast.success(t('Saved as “{name}”', { name: e.name || name }));
  };
  return (
    <Modal size="xwide" title={t('Older states of “{name}”', { name: entry.name })} subtitle="From Hostwerk's backups – also states older than the versions" onClose={() => onDone(null)}
      footer={<><a className="btn ghost" href="#/backups" onClick={() => onDone(null)}><History size={14} aria-hidden="true" /> {t('Backup center')}</a><span className="grow" /><Btn onClick={() => onDone(null)}>Close</Btn>{cur && bytes && onCompare && <Btn icon={GitCompare} onClick={() => { onDone(null); onCompare({ bytes, label: fmtDateTime(cur.time) }); }}>Compare with now</Btn>}{cur && bytes && <Btn icon={Copy} onClick={asCopy}>Save as a copy</Btn>}{cur && <Btn className="primary" icon={RotateCcw} onClick={restore}>Restore</Btn>}</>}>
      {err ? <Callout kind="err" dismissible={false}>{err}</Callout> : !list ? <Loading inline /> : !list.length ? <Empty icon={History} title="No backup contains this file" text="Back up the Office documents (Backups) – then every backup is a state you can go back to." action={<a className="btn" href="#/backups">Set up backups</a>} /> : (
        <div className="of-shell-versions">
          <ul className="of-shell-timeline" role="listbox" aria-label={t('Backups')}>
            {list.map((v) => <li key={v.id} role="option" aria-selected={cur && cur.id === v.id} tabIndex={0} className={cur && cur.id === v.id ? 'on' : ''} onClick={() => setCur(v)} onKeyDown={(e) => { if (e.key === 'Enter') setCur(v); }}>
              <b className="small">{fmtWhen(v.time)}</b><span className="tiny muted">{[fmtDateTime(v.time), v.size != null ? fmtBytes(v.size) : '', v.label].filter(Boolean).join(' · ')}</span></li>)}
          </ul>
          <div className="of-shell-vpreview">{!cur ? null : bytes ? <FilePreviewBox file={{ ...entry, mtime: cur.time }} bytes={bytes} editorFallback /> : <div className="of-shell-skel-page" aria-busy="true" />}</div>
        </div>
      )}
    </Modal>
  );
}

/* ---------------- compare ---------------- */
function TextCompare({ a, b, chunk, ext }) {
  const [d, setD] = useState(null);
  useEffect(() => {
    let alive = true;
    (async () => {
      const ta = chunk && chunk.toText ? await chunk.toText(a.bytes, ext).catch(() => '') : '';
      const tb = chunk && chunk.toText ? await chunk.toText(b.bytes, ext).catch(() => '') : '';
      if (alive) setD(lineDiff(ta, tb));
    })();
    return () => { alive = false; };
  }, [a, b]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!d) return <Loading inline />;
  const n = d.filter((x) => x.op !== '=').length;
  return (
    <div className="stack of-shell-gap-2">
      <span className="small muted">{n ? pl(n, '{n} changed line', '{n} changed lines') : t('No difference in the text.')}</span>
      <div className="of-shell-diff">{d.map((x, i) => <div key={i} className={`of-shell-diff-l is-${x.op === '+' ? 'add' : x.op === '-' ? 'del' : 'eq'}`}><span aria-hidden="true">{x.op === '=' ? ' ' : x.op}</span><Keep>{x.text || ' '}</Keep></div>)}</div>
    </div>
  );
}
function CompareDialog({ a, b, chunk, ext, host, onDone }) {
  const C = chunk && chunk.Compare;
  return (
    <Modal size="xwide" title="Compare" subtitle={<Keep>{`${a.label || ''}  ↔  ${b.label || ''}`}</Keep>} onClose={() => onDone(null)} footer={<Btn onClick={() => onDone(null)}>Close</Btn>}>
      <div className="of-shell-compare">{C ? <C a={a} b={b} ext={ext} host={host} /> : <TextCompare a={a} b={b} chunk={chunk} ext={ext} />}</div>
    </Modal>
  );
}

const DIALOGS = {
  'db-confirm': DbConfirm, 'db-pick': DbPick, 'data-pick': DataPick, 'field-pick': FieldPick, forms: FormsDialog, rules: RulesDialog,
  component: ComponentDialog, remote: RemoteDialog, vault: VaultDialog, clean: CleanDialog, publish: PublishDialog, report: ReportDialog,
  backups: BackupsDialog, compare: CompareDialog,
};
/** more dialogs from other shell modules (templates.jsx: 'save-template') */
export function registerFeature(type, C) { DIALOGS[type] = C; }
/** is a FeatureHost mounted (dialog host in Files, an editor host)? */
export const featureHostMounted = () => mounted > 0;
