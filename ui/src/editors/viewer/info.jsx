/*
 * Formats we cannot render fully ourselves. Never a dead end: a built-in best effort (preview image, drawing text,
 * formula as MathML, database objects, text found in a binary file) plus "Open fully", which renders the file
 * exactly through the Office engine (host.convert → PDF; the shell sets the engine up when it is missing).
 *
 *   ConvertView – vsd, pub, odg, otg, fodg, odf (and any other ext with a PDF route)
 *   IworkView   – Apple Pages, Keynote, Numbers (QuickLook preview + IWA text; "Edit" = convert to docx/pptx/xlsx)
 *   InfoView    – one, mpp, accdb, mdb, odb
 */
import { useEffect, useMemo, useState } from 'react';
import JSZip from 'jszip';
import DOMPurify from 'dompurify';
import { Sparkles, FileText, Database, NotebookTabs, GanttChart, Shapes, Sigma, Newspaper, PenLine, Table2, Presentation, ChevronDown, ChevronRight } from 'lucide-react';
import { Btn, Callout, Loading } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { parseXml, kids, text as xtext, local } from '../_shared/xmltree.js';
import { ActionRow, InfoCard, PdfPages, fmtSize, fmtDate, baseName, useBlobUrl } from './common.jsx';
import { binaryStrings, iworkText } from './extract.js';
import { FORMATS } from '../../lib/formats.js';
import { Keep } from '../_shared/keep.jsx';
import { FileTypeIcon } from '../../lib/fileIcons.jsx';

const isZip = (b) => b[0] === 0x50 && b[1] === 0x4b;
function descend(el, name, out = []) { for (const x of (el && el.c) || []) { if (typeof x === 'string') continue; if (local(x.n) === name) out.push(x); descend(x, name, out); } return out; }

/** ODF text paragraphs (content.xml or flat XML) */
export function odfParagraphs(xml) {
  const root = parseXml(xml);
  const body = descend(root, 'body')[0] || root;
  return [...descend(body, 'p'), ...descend(body, 'h')].map((p) => xtext(p).replace(/\s+/g, ' ').trim()).filter(Boolean);
}

/** what the built-in view can show for this file: { image?: {bytes, mime}, mathml?, paragraphs: [], note } */
async function builtIn(ext, bytes) {
  const res = { paragraphs: [], image: null, mathml: '', pdf: null };
  if (['odg', 'otg', 'odf'].includes(ext) && isZip(bytes)) {
    const zip = await JSZip.loadAsync(bytes);
    const th = zip.file('Thumbnails/thumbnail.png');
    if (th) res.image = { bytes: await th.async('uint8array'), mime: 'image/png' };
    const content = zip.file('content.xml') && await zip.file('content.xml').async('string');
    if (content && ext === 'odf') res.mathml = mathFrom(content);
    else if (content) res.paragraphs = odfParagraphs(content);
    return res;
  }
  if (ext === 'fodg') { res.paragraphs = odfParagraphs(new TextDecoder().decode(bytes)); return res; }
  res.paragraphs = binaryStrings(bytes, { min: 6, max: 1500 }).strings.filter((s) => /\s/.test(s) || s.length > 12).slice(0, 800);
  return res;
}
function mathFrom(xml) {
  const hit = /<(math:)?math[\s>][\s\S]*?<\/(math:)?math>/.exec(xml);
  if (!hit) return '';
  const m = hit[0].replace(/<(\/?)math:/g, '<$1').replace(/<annotation[\s\S]*?<\/annotation>/g, '');
  return DOMPurify.sanitize(m, { USE_PROFILES: { mathMl: true } });
}

function TextFound({ paragraphs, title }) {
  const [open, setOpen] = useState(paragraphs.length < 60);
  if (!paragraphs.length) return null;
  return (
    <div className="card">
      <button type="button" className="card-head of-view-acc" aria-expanded={open} onClick={() => setOpen(!open)}>
        {open ? <ChevronDown size={16} /> : <ChevronRight size={16} />}<h3>{title || t('Text in this file')}</h3><span className="badge">{paragraphs.length}</span>
      </button>
      {open && <div className="card-body of-view-found">{paragraphs.map((p, i) => <p key={i}><Keep>{p}</Keep></p>)}</div>}
    </div>
  );
}

const ICON = { vsd: Shapes, pub: Newspaper, odg: PenLine, otg: PenLine, fodg: PenLine, odf: Sigma, one: NotebookTabs, mpp: GanttChart, accdb: Database, mdb: Database, odb: Database, pages: FileText, key: Presentation, numbers: Table2 };

export function ConvertView({ file, bytes, host }) {
  const [bi, setBi] = useState(null);
  const [exact, setExact] = useState(null);
  const [busyErr, setBusyErr] = useState(null);
  useEffect(() => { let gone = false; builtIn(file.ext, bytes).then((r) => !gone && setBi(r)).catch(() => !gone && setBi({ paragraphs: [] })); return () => { gone = true; }; }, [bytes, file.ext]);
  const img = useBlobUrl(bi && bi.image && bi.image.bytes, bi && bi.image && bi.image.mime);
  const fmt = FORMATS[file.ext] || {};
  const Icon = ICON[file.ext] || FileText;
  const open = async () => { setBusyErr(null); try { setExact(await host.convert(bytes, file.ext, 'pdf')); } catch (e) { setBusyErr(e); throw e; } };
  return (
    <div className="of-view-pad stack">
      <div className="of-view-toolbar of-view-bare">
        <h2 className="of-view-title"><FileTypeIcon name={file.name} size={20} /> <Keep>{file.name}</Keep></h2>
        <span className="small muted">{t(fmt.label || 'File')} · {fmtSize(bytes.length)}</span>
        <span className="grow" />
        {exact ? <Btn className="sm" onClick={() => setExact(null)}>Built-in view</Btn> : <Btn className="sm primary" icon={Sparkles} onClick={open}>Open fully</Btn>}
      </div>
      {exact ? <PdfPages bytes={exact} maxWidth={1200} /> : (
        <>
          <Callout kind="info" title={t('Simple view')} dismissible={false}>{t('This shows what can be read from the file directly. “Open fully” renders it exactly with the Office engine.')}</Callout>
          {busyErr && <Callout kind="warn">{busyErr.message}</Callout>}
          {!bi && <Loading inline />}
          {img && <div className="of-view-thumb card"><img src={img} alt={t('Preview stored in the file')} /><span className="small muted">{t('Preview stored in the file')}</span></div>}
          {bi && bi.mathml && <div className="card of-view-math"><div className="card-body" dangerouslySetInnerHTML={{ __html: bi.mathml }} /></div>}
          {bi && <TextFound paragraphs={bi.paragraphs} />}
          {bi && !img && !bi.mathml && !bi.paragraphs.length && <Callout kind="info">{t('No readable text found – “Open fully” shows the file.')}</Callout>}
        </>
      )}
      <ActionRow file={file} bytes={bytes} host={host} convertTo={fmt.via === 'pdf' || ['vsd', 'pub'].includes(file.ext) ? 'pdf' : null} convertLabel={t('Save as PDF')} />
    </div>
  );
}

/* ---------------- Apple iWork ---------------- */
const IWORK_TARGET = { pages: 'docx', key: 'pptx', numbers: 'xlsx' };
export async function readIwork(bytes) {
  const out = { pdf: null, image: null, text: '' };
  if (!isZip(bytes)) return out;
  const zip = await JSZip.loadAsync(bytes);
  const find = (re) => Object.keys(zip.files).find((p) => re.test(p));
  const pdf = find(/(^|\/)QuickLook\/Preview\.pdf$/i);
  if (pdf) out.pdf = await zip.file(pdf).async('uint8array');
  const img = find(/(^|\/)preview\.jpg$/i) || find(/(^|\/)preview-web\.jpg$/i) || find(/(^|\/)QuickLook\/Thumbnail\.jpg$/i) || find(/(^|\/)preview-micro\.jpg$/i);
  if (img) out.image = await zip.file(img).async('uint8array');
  const iwa = Object.keys(zip.files).filter((p) => /\.iwa$/i.test(p)).sort((a, b) => (/Document\.iwa$/i.test(b) ? 1 : 0) - (/Document\.iwa$/i.test(a) ? 1 : 0));
  if (iwa.length) out.text = iworkText(await Promise.all(iwa.slice(0, 40).map((p) => zip.file(p).async('uint8array'))));
  else {
    const xml = find(/(^|\/)index\.(xml|apxl)$/i);
    if (xml) out.text = odfParagraphs(await zip.file(xml).async('string')).join('\n');
  }
  return out;
}

export function IworkView({ file, bytes, host }) {
  const [r, setR] = useState(null);
  const [err, setErr] = useState(null);
  useEffect(() => { let gone = false; readIwork(bytes).then((x) => !gone && setR(x)).catch((e) => !gone && setErr(e)); return () => { gone = true; }; }, [bytes]);
  const img = useBlobUrl(r && r.image, 'image/jpeg');
  const target = IWORK_TARGET[file.ext] || 'pdf';
  const edit = async () => {
    const out = await host.convert(bytes, file.ext, target);
    if (host.openAs) host.openAs({ bytes: out, ext: target, name: `${baseName(file.name)}.${target}` });
    else host.download(out, `${baseName(file.name)}.${target}`, (FORMATS[target] || {}).mime);
  };
  const paragraphs = useMemo(() => (r && r.text ? r.text.split('\n').filter(Boolean) : []), [r]);
  return (
    <div className="of-view-pad stack">
      <div className="of-view-toolbar of-view-bare">
        <h2 className="of-view-title"><FileTypeIcon name={file.name} size={20} /> <Keep>{file.name}</Keep></h2>
        <span className="small muted">{t((FORMATS[file.ext] || {}).label || 'Apple iWork file')} · {fmtSize(bytes.length)}</span>
        <span className="grow" />
        <Btn className="sm primary" icon={Sparkles} onClick={edit}>{t('Edit as {ext}', { ext: target.toUpperCase() })}</Btn>
      </div>
      {err && <Callout kind="warn">{err.message}</Callout>}
      {!r && !err && <Loading inline />}
      {r && r.pdf && <PdfPages bytes={r.pdf} />}
      {r && !r.pdf && img && <div className="of-view-thumb card"><img src={img} alt={t('Preview stored in the file')} /><span className="small muted">{t('Preview of the first page, stored in the file')}</span></div>}
      {r && <TextFound paragraphs={paragraphs} title={t('Text found in the document')} />}
      {r && !r.pdf && !img && !paragraphs.length && <Callout kind="info">{t('This file has no stored preview. “Edit” converts it with the Office engine.')}</Callout>}
      <ActionRow file={file} bytes={bytes} host={host} convertTo="pdf" convertLabel={t('Save as PDF')} />
    </div>
  );
}

/* ---------------- desktop-only formats ---------------- */
const ABOUT = {
  one: ['OneNote section', 'A section of a OneNote notebook with pages, ink and attachments. Its layout can only be shown by OneNote (Windows, Mac, web).'],
  mpp: ['Microsoft Project plan', 'A project plan with tasks, dates, resources and dependencies. Microsoft Project or ProjectLibre can open it.'],
  accdb: ['Access database', 'A Microsoft Access database with tables, queries, forms and reports. Microsoft Access (or LibreOffice Base with a driver) opens it.'],
  mdb: ['Access database (97–2003)', 'An older Microsoft Access database. Microsoft Access opens it; LibreOffice Base can read the tables with a driver.'],
  odb: ['OpenDocument database', 'A LibreOffice Base database: connection, queries, forms and reports – often with an embedded HSQLDB or Firebird database.'],
};

export async function readOdb(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const xml = zip.file('content.xml') && await zip.file('content.xml').async('string');
  const res = { tables: [], queries: [], forms: [], reports: [], connection: '', embedded: '' };
  if (!xml) return res;
  const root = parseXml(xml);
  const names = (n) => descend(root, n).map((e) => e.a['db:name']).filter(Boolean);
  res.tables = names('table-representation');
  res.queries = names('query');
  const comps = (n) => { const c = descend(root, n)[0]; return c ? descend(c, 'component').map((e) => e.a['db:name']).filter(Boolean) : []; };
  res.forms = comps('forms'); res.reports = comps('reports');
  const conn = descend(root, 'connection-resource')[0];
  res.connection = conn ? conn.a['xlink:href'] || '' : '';
  if (zip.file('database/script') || zip.file('database/properties')) res.embedded = 'HSQLDB';
  else if (Object.keys(zip.files).some((p) => /^database\/firebird/i.test(p))) res.embedded = 'Firebird';
  void kids;
  return res;
}

export function InfoView({ file, bytes, host }) {
  const [about, text] = ABOUT[file.ext] || [FORMATS[file.ext] ? FORMATS[file.ext].label : 'File', ''];
  const [odb, setOdb] = useState(null);
  const strings = useMemo(() => (file.ext === 'odb' ? [] : binaryStrings(bytes, { min: 6, max: 1200 }).strings.filter((s) => /\s/.test(s)).slice(0, 600)), [bytes, file.ext]);
  useEffect(() => { if (file.ext === 'odb') readOdb(bytes).then(setOdb).catch(() => setOdb({ tables: [], queries: [], forms: [], reports: [] })); }, [bytes, file.ext]);
  const Icon = ICON[file.ext] || FileText;
  const list = (label, items) => (items && items.length ? <div><b className="small">{label}</b><div className="of-view-chips">{items.map((x) => <span key={x} className="pill"><Keep>{x}</Keep></span>)}</div></div> : null);
  return (
    <div className="of-view-pad stack">
      <InfoCard icon={Icon} title={t(about)} text={t(text)} facts={[[t('File'), file.name], [t('Size'), fmtSize(file.size || bytes.length)], [t('Modified'), file.mtime ? fmtDate(file.mtime) : '']]}>
        <p className="small muted">{t('Hostwerk keeps the file safe and unchanged. Download it or open it with its desktop app (top bar → Open in desktop app).')}</p>
        {odb && (
          <div className="stack">
            {odb.embedded && <div className="small">{t('Embedded database: {db}', { db: odb.embedded })}</div>}
            {odb.connection && <div className="small mono of-view-ellipsis" title={odb.connection}>{odb.connection}</div>}
            {list(t('Tables'), odb.tables)}{list(t('Queries'), odb.queries)}{list(t('Forms'), odb.forms)}{list(t('Reports'), odb.reports)}
            {!odb.tables.length && !odb.queries.length && !odb.forms.length && !odb.reports.length && <span className="small muted">{t('No tables, queries, forms or reports are listed in this file.')}</span>}
          </div>
        )}
      </InfoCard>
      <TextFound paragraphs={strings} title={t('Text found in the file')} />
      <ActionRow file={file} bytes={bytes} host={host} />
    </div>
  );
}

export async function infoText(bytes, ext) {
  if (ext === 'odb') { const o = await readOdb(bytes); return [...o.tables, ...o.queries, ...o.forms, ...o.reports].join('\n'); }
  if (['pages', 'key', 'numbers'].includes(ext)) return (await readIwork(bytes)).text;
  const b = await builtIn(ext, bytes);
  return b.paragraphs.join('\n');
}
