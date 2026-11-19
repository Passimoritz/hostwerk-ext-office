/*
 * Google Drive shortcut files. Everything stays inside Hostwerk: when Google Drive is connected the shell opens the
 * exported file in the matching editor itself; here we show the card, a read-only PDF view of the export (if the
 * user asks) and "Connect Google Drive" otherwise (sign-in happens in a Hostwerk window).
 */
import { useEffect, useMemo, useState } from 'react';
import { FileText, Sheet, Presentation, PenTool, ClipboardList, Globe, Map as MapIcon, StickyNote, Link as LinkIcon, PlugZap, Eye, Download } from 'lucide-react';
import { get, post } from '@hostwerk/api';
import { Btn, Callout, Loading } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { GOOGLE, parseGoogleFile, googleId, googleUrl, googleExport } from './google.js';
import { ActionRow, PdfPages, baseName } from './common.jsx';
import { Keep, Raw } from '../_shared/keep.jsx';

const ICONS = { FileText, Sheet, Presentation, PenTool, ClipboardList, Globe, Map: MapIcon, StickyNote, Link: LinkIcon };
const MIMES = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', pdf: 'application/pdf', odt: 'application/vnd.oasis.opendocument.text', ods: 'application/vnd.oasis.opendocument.spreadsheet', odp: 'application/vnd.oasis.opendocument.presentation', svg: 'image/svg+xml', png: 'image/png' };

async function fetchExport(id, as) {
  const r = await fetch(`/api/office/google/raw?id=${encodeURIComponent(id)}&as=${encodeURIComponent(as)}`, { credentials: 'same-origin' });
  if (!r.ok) { let m = `${r.status}`; try { const j = await r.json(); m = j.error || m; } catch { /* raw */ } throw new Error(m); }
  return new Uint8Array(await r.arrayBuffer());
}

export default function GoogleView({ file, bytes, host }) {
  const info = useMemo(() => parseGoogleFile(bytes), [bytes]);
  const kind = GOOGLE[file.ext] || GOOGLE.glink;
  const Icon = ICONS[kind.icon] || LinkIcon;
  const id = googleId(info);
  const url = googleUrl(file.ext, info);
  const exp = googleExport(file.ext);
  const [state, setState] = useState(null); // { connected, email? } | { error }
  const [pdf, setPdf] = useState(null);
  const load = () => get('/office/google/state').then(setState).catch((e) => setState({ connected: false, unavailable: e && e.status === 404, error: e.message }));
  useEffect(() => { load(); }, []);
  const connect = async () => { await post('/office/google/connect', {}); await load(); };
  const preview = async () => setPdf(await fetchExport(id, 'pdf'));
  const save = async (as) => { const out = await fetchExport(id, as); host.download(out, `${baseName(file.name)}.${as}`, MIMES[as]); };
  const openHere = async () => {
    const as = exp ? exp.as : 'pdf';
    const out = await fetchExport(id, as);
    if (host.openAs) host.openAs({ bytes: out, ext: as, name: `${baseName(file.name)}.${as}`, google: { id } });
    else setPdf(as === 'pdf' ? out : await fetchExport(id, 'pdf'));
  };
  return (
    <div className="of-view-google of-view-pad stack">
      <div className="card of-view-gcard">
        <div className="of-view-gicon" aria-hidden="true"><Icon size={36} /></div>
        <div className="grow stack">
          <div>
            <h2><Keep>{baseName(file.name)}</Keep></h2>
            <div className="muted">{t(kind.label)}<Keep>{info.email ? ` · ${info.email}` : ''}</Keep></div>
          </div>
          {!id && !url && <Callout kind="warn">{t('This shortcut does not say which Google file it points to.')}</Callout>}
          {state == null && <Loading inline />}
          {state && state.connected && id && (
            <div className="row wrap">
              {exp && <Btn className="primary" icon={Eye} onClick={openHere}>{exp.kind === 'pdf' ? t('Open here') : t('Open in Office')}</Btn>}
              {(exp || kind.also.includes('pdf')) && <Btn icon={Eye} onClick={preview}>Show as PDF</Btn>}
              {[exp && exp.as, ...kind.also].filter(Boolean).filter((x, i, a) => a.indexOf(x) === i).map((as) => <Btn key={as} className="ghost" icon={Download} onClick={() => save(as)}>{t('Download as {ext}', { ext: as.toUpperCase() })}</Btn>)}
              {!exp && <span className="small muted">{t('Google does not offer an export for this kind of file.')}</span>}
            </div>
          )}
          {state && !state.connected && (
            <div className="stack">
              {state.unavailable
                ? <Callout kind="info" title={t('Google Drive is not available yet')}>{t('Update the Office extension to connect Google Drive.')}</Callout>
                : <div className="row wrap"><Btn className="primary" icon={PlugZap} onClick={connect}>Connect Google Drive …</Btn><span className="small muted">{t('Sign in once in a Hostwerk window – then Google files open right here.')}</span></div>}
            </div>
          )}
          <details className="small">
            <summary>{t('Without connecting: how to edit it here')}</summary>
            <p className="muted">{t('In Google, choose File → Download → Microsoft Word / Excel / PowerPoint (or OpenDocument) and put the file into your Office library – it then opens here like any other document. Google Takeout exports are already Office files.')}</p>
            {url && <Raw as="p" className="mono small of-view-ellipsis" title={url}><Keep>{url}</Keep></Raw>}
          </details>
        </div>
      </div>
      {pdf && <PdfPages bytes={pdf} />}
      <ActionRow file={file} bytes={bytes} host={host} />
    </div>
  );
}

export function googleText(bytes, ext, name) { const j = parseGoogleFile(bytes); return [name, t((GOOGLE[ext] || GOOGLE.glink).label), j.email || '', googleUrl(ext, j)].filter(Boolean).join('\n'); }
