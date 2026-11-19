/* JSON (collapsible tree + raw), XML (pretty printed, collapsible) and any other text (encoding detection, monospace). */
import { useMemo, useState, memo } from 'react';
import { ChevronRight, ChevronDown, Braces, Code2, FileText, WrapText, Copy } from 'lucide-react';
import { Seg, Callout, toast } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { parseXml, isEl } from '../_shared/xmltree.js';
import { ActionRow, decodeText } from './common.jsx';
import { Keep } from '../_shared/keep.jsx';

const MAX_TEXT = 3 * 1024 * 1024;

const JsonNode = memo(function JsonNode({ k, v, depth }) {
  const [open, setOpen] = useState(depth < 2);
  const isObj = v && typeof v === 'object';
  const label = k != null && <span className="of-view-jk"><Keep>{typeof k === 'number' ? k : JSON.stringify(k)}</Keep>: </span>;
  if (!isObj) {
    const cls = v === null ? 'null' : typeof v;
    return <div className="of-view-jrow" style={{ paddingLeft: depth * 16 + 18 }}>{label}<span className={`of-view-j-${cls}`}><Keep>{JSON.stringify(v)}</Keep></span></div>;
  }
  const entries = Array.isArray(v) ? v.map((x, i) => [i, x]) : Object.entries(v);
  const brackets = Array.isArray(v) ? ['[', ']'] : ['{', '}'];
  return (
    <>
      <div className="of-view-jrow" style={{ paddingLeft: depth * 16 }}>
        <button type="button" className="of-view-jtog" aria-expanded={open} aria-label={open ? 'Collapse' : 'Expand'} onClick={() => setOpen(!open)}>{open ? <ChevronDown size={12} /> : <ChevronRight size={12} />}</button>
        {label}<span className="muted">{brackets[0]}{!open && ` ${t('{n} entries', { n: entries.length })} ${brackets[1]}`}</span>
      </div>
      {open && entries.slice(0, 2000).map(([ck, cv]) => <JsonNode key={ck} k={ck} v={cv} depth={depth + 1} />)}
      {open && entries.length > 2000 && <div className="of-view-jrow muted" style={{ paddingLeft: (depth + 1) * 16 + 18 }}>{t('… {n} more', { n: entries.length - 2000 })}</div>}
      {open && <div className="of-view-jrow muted" style={{ paddingLeft: depth * 16 + 18 }}>{brackets[1]}</div>}
    </>
  );
});

const XmlNode = memo(function XmlNode({ el, depth }) {
  const [open, setOpen] = useState(depth < 3);
  if (typeof el === 'string') { const s = el.trim(); return s ? <div className="of-view-jrow of-view-xtext" style={{ paddingLeft: depth * 16 + 18 }}><Keep>{s}</Keep></div> : null; }
  const attrs = Object.entries(el.a || {});
  const children = (el.c || []).filter((c) => isEl(c) || String(c).trim());
  const onlyText = children.length === 1 && typeof children[0] === 'string' && children[0].length < 120;
  const head = <><span className="of-view-xtag">&lt;<Keep>{el.n}</Keep></span>{attrs.map(([k, v]) => <span key={k}> <span className="of-view-xattr"><Keep>{k}</Keep></span>=<span className="of-view-j-string">&quot;<Keep>{v}</Keep>&quot;</span></span>)}<span className="of-view-xtag">{children.length ? '>' : ' />'}</span></>;
  if (!children.length || onlyText) return <div className="of-view-jrow" style={{ paddingLeft: depth * 16 + 18 }}>{head}{onlyText && <><span className="of-view-xtext"><Keep>{children[0].trim()}</Keep></span><span className="of-view-xtag">&lt;/<Keep>{el.n}</Keep>&gt;</span></>}</div>;
  return (
    <>
      <div className="of-view-jrow" style={{ paddingLeft: depth * 16 }}>
        <button type="button" className="of-view-jtog" aria-expanded={open} aria-label={open ? 'Collapse' : 'Expand'} onClick={() => setOpen(!open)}>{open ? <ChevronDown size={12} /> : <ChevronRight size={12} />}</button>
        {head}{!open && <span className="muted"> … <span className="of-view-xtag">&lt;/<Keep>{el.n}</Keep>&gt;</span></span>}
      </div>
      {open && children.slice(0, 3000).map((c, i) => <XmlNode key={i} el={c} depth={depth + 1} />)}
      {open && <div className="of-view-jrow" style={{ paddingLeft: depth * 16 + 18 }}><span className="of-view-xtag">&lt;/<Keep>{el.n}</Keep>&gt;</span></div>}
    </>
  );
});

/** XML pretty printer (pure, on the xmltree) */
export function prettyXml(el, ind = '') {
  if (typeof el === 'string') return el.trim() ? `${ind}${el.trim()}\n` : '';
  const at = Object.entries(el.a || {}).map(([k, v]) => ` ${k}="${String(v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}"`).join('');
  const kids = (el.c || []).filter((c) => isEl(c) || String(c).trim());
  if (!kids.length) return `${ind}<${el.n}${at}/>\n`;
  if (kids.length === 1 && typeof kids[0] === 'string') return `${ind}<${el.n}${at}>${kids[0].trim().replace(/&/g, '&amp;').replace(/</g, '&lt;')}</${el.n}>\n`;
  return `${ind}<${el.n}${at}>\n${kids.map((c) => (typeof c === 'string' ? `${ind}  ${c.trim().replace(/&/g, '&amp;').replace(/</g, '&lt;')}\n` : prettyXml(c, `${ind}  `))).join('')}${ind}</${el.n}>\n`;
}

export default function DataView({ file, bytes, host }) {
  const { text, encoding } = useMemo(() => decodeText(bytes.length > MAX_TEXT ? bytes.subarray(0, MAX_TEXT) : bytes), [bytes]);
  const kind = file.ext === 'json' ? 'json' : file.ext === 'xml' ? 'xml' : 'text';
  const parsed = useMemo(() => {
    if (kind === 'json') { try { return { ok: true, v: JSON.parse(text.replace(/^﻿/, '')) }; } catch (e) { return { ok: false, err: e.message }; } }
    if (kind === 'xml') { try { const root = parseXml(text); return root ? { ok: true, v: root, pretty: prettyXml(root) } : { ok: false, err: t('No XML element found.') }; } catch (e) { return { ok: false, err: e.message }; } }
    return { ok: false };
  }, [text, kind]);
  const [mode, setMode] = useState(kind === 'text' ? 'raw' : 'tree');
  const [wrap, setWrap] = useState(true);
  const raw = mode === 'pretty' && parsed.ok ? (kind === 'json' ? JSON.stringify(parsed.v, null, 2) : parsed.pretty) : text;
  const lines = useMemo(() => raw.split('\n').length, [raw]);
  const copy = async () => { try { await navigator.clipboard.writeText(raw); toast.success(t('Copied')); } catch { toast.error(t('Could not copy')); } };
  return (
    <div className="of-view-data">
      <div className="of-view-toolbar">
        {kind !== 'text' && parsed.ok && <Seg value={mode} onChange={setMode} options={[{ id: 'tree', label: t('Tree'), icon: Braces }, { id: 'pretty', label: t('Formatted'), icon: Code2 }, { id: 'raw', label: t('Raw'), icon: FileText }]} />}
        {mode !== 'tree' && <button type="button" className={`btn sm ghost${wrap ? ' on' : ''}`} aria-pressed={wrap} onClick={() => setWrap(!wrap)} title="Wrap lines"><WrapText size={16} /> {t('Wrap lines')}</button>}
        <button type="button" className="btn sm ghost" onClick={copy} title="Copy all"><Copy size={16} /> {t('Copy')}</button>
        <span className="grow" />
        <span className="small muted">{encoding} · {t('{n} lines', { n: lines })}</span>
      </div>
      {kind !== 'text' && !parsed.ok && <div className="of-view-pad0"><Callout kind="warn" title={kind === 'json' ? t('This JSON is not valid') : t('This XML is not well-formed')}>{parsed.err}</Callout></div>}
      {bytes.length > MAX_TEXT && <div className="of-view-pad0"><Callout kind="info">{t('Only the first 3 MB are shown. Download the file to see everything.')}</Callout></div>}
      <div className="of-view-code">
        {mode === 'tree' && parsed.ok ? (kind === 'json' ? <JsonNode k={null} v={parsed.v} depth={0} /> : <XmlNode el={parsed.v} depth={0} />) : <pre className={`mono${wrap ? ' wrap' : ''}`}><Keep>{raw}</Keep></pre>}
      </div>
      <ActionRow file={file} bytes={bytes} host={host} />
    </div>
  );
}
