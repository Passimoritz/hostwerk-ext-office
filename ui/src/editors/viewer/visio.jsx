/* Visio drawings (vsdx/vsdm/vstx): built-in SVG rendering with page tabs and zoom; "Exact view" through the Office engine. */
import { useEffect, useMemo, useState } from 'react';
import JSZip from 'jszip';
import DOMPurify from 'dompurify';
import { Shapes, ZoomIn, ZoomOut, Maximize, Sparkles } from 'lucide-react';
import { Btn, Loading, Callout, Tabs } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { readVsdx } from './vsdx.js';
import { metafileToDataUrl } from './emf.js';
import { ActionRow, PdfPages } from './common.jsx';
import { Keep, Raw } from '../_shared/keep.jsx';
import { FileTypeIcon } from '../../lib/fileIcons.jsx';

const MEDIA = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', emf: 'image/emf', wmf: 'image/wmf', svg: 'image/svg+xml', tif: 'image/tiff', tiff: 'image/tiff' };

export async function loadVisio(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const files = { media: {} };
  for (const [path, f] of Object.entries(zip.files)) {
    if (f.dir) continue;
    if (/\.(xml|rels)$/i.test(path)) files[path] = await f.async('string');
    else if (/^visio\/media\//i.test(path)) {
      const ext = path.split('.').pop().toLowerCase();
      if (ext === 'emf' || ext === 'wmf') { const svg = metafileToDataUrl(await f.async('uint8array')); if (svg) { files.media[path] = svg; continue; } }
      const b = await f.async('base64');
      files.media[path] = `data:${MEDIA[ext] || 'application/octet-stream'};base64,${b}`;
    }
  }
  return readVsdx(files);
}

export default function VisioView({ file, bytes, host }) {
  const [doc, setDoc] = useState(null);
  const [err, setErr] = useState(null);
  const [page, setPage] = useState(0);
  const [zoom, setZoom] = useState('fit');
  const [exact, setExact] = useState(null); // PDF from the engine
  useEffect(() => {
    let gone = false;
    loadVisio(bytes).then((d) => !gone && setDoc(d)).catch((e) => !gone && setErr(e));
    return () => { gone = true; };
  }, [bytes]);
  const p = doc && doc.pages[page];
  const svg = useMemo(() => (p ? DOMPurify.sanitize(p.svg, { USE_PROFILES: { svg: true, svgFilters: true }, ADD_TAGS: ['use'], ADD_ATTR: ['marker-start', 'marker-end', 'orient', 'refX', 'refY', 'markerUnits'] }) : ''), [p]);
  useEffect(() => { if (doc) host.status(t('Page {n} of {count}', { n: page + 1, count: doc.pages.length })); }, [page, doc]); // eslint-disable-line react-hooks/exhaustive-deps
  const openExact = async () => setExact(await host.convert(bytes, file.ext, 'pdf'));

  if (err) return <div className="of-view-pad"><Callout kind="err" title={t('This drawing cannot be read')}>{err.message}</Callout><ActionRow file={file} bytes={bytes} host={host} convertTo="pdf" /></div>;
  if (!doc) return <Loading text={t('Drawing the diagram …')} />;
  const scale = zoom === 'fit' ? null : zoom;
  return (
    <div className="of-view-visio">
      <div className="of-view-toolbar">
        <h2 className="of-view-title"><FileTypeIcon name={file.name} size={20} /> <Keep>{file.name}</Keep></h2>
        <span className="grow" />
        {!exact && <>
          <button type="button" className="btn sm ghost icon" title="Zoom out" aria-label="Zoom out" onClick={() => setZoom(Math.max(0.25, (scale || 1) / 1.25))}><ZoomOut size={16} /></button>
          <span className="small muted">{scale ? `${Math.round(scale * 100)} %` : t('Fit')}</span>
          <button type="button" className="btn sm ghost icon" title="Zoom in" aria-label="Zoom in" onClick={() => setZoom(Math.min(6, (scale || 1) * 1.25))}><ZoomIn size={16} /></button>
          <button type="button" className={`btn sm ghost icon${zoom === 'fit' ? ' on' : ''}`} title="Fit to window" aria-label="Fit to window" onClick={() => setZoom('fit')}><Maximize size={16} /></button>
        </>}
        {exact ? <Btn className="sm" onClick={() => setExact(null)}>Built-in view</Btn> : <Btn className="sm" icon={Sparkles} onClick={openExact} title={t('Renders the drawing exactly with the Office engine')}>Open fully</Btn>}
      </div>
      {!exact && doc.pages.length > 1 && <Raw as={Tabs} tabs={doc.pages.map((x, i) => ({ id: String(i), label: x.name || t('Page {n}', { n: i + 1 }) }))} value={String(page)} onChange={(v) => setPage(Number(v))} />}
      <div className="of-view-canvas">
        {exact ? <PdfPages bytes={exact} maxWidth={1400} /> : p ? (
          <div className={`of-view-paper${zoom === 'fit' ? ' fit' : ''}`} style={scale ? { width: p.w * 96 * scale } : undefined} dangerouslySetInnerHTML={{ __html: svg }} />
        ) : <Callout kind="info">{t('This drawing has no pages.')}</Callout>}
      </div>
      {!exact && <div className="of-view-pad0 small muted">{t('Built-in view: shapes, text, lines and fills are drawn from the file; themes, effects and some special shapes look simpler than in Visio. “Open fully” shows it exactly.')}</div>}
      <ActionRow file={file} bytes={bytes} host={host} convertTo="pdf" convertLabel={t('Save as PDF')} />
    </div>
  );
}

export async function visioText(bytes) { const d = await loadVisio(bytes); return d.pages.map((p) => [p.name, ...p.texts].join('\n')).join('\n\n'); }
