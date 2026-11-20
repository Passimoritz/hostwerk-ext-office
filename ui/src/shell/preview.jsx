/*
 * Read-only previews: the editor chunk's Preview component (or a short text from toText) for the details panel, the
 * versions sheet and Hostwerk's file browser. A missing chunk is no error here – the icon and "Open" still work.
 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { t } from '@hostwerk/i18n';
import { AlertTriangle, RotateCcw } from 'lucide-react';
import { Loading, Btn } from '@hostwerk/ui';
import { ErrorBoundary } from '@hostwerk/kit';
import { office } from '../lib/api';
import { loadEditor } from '../lib/loader';
import { formatOf, extOf } from '../lib/formats';
import { settingsNow, hasConverterIn, statusNow } from './store';
import { FileIcon, uiLang } from './common';
import { download } from '../lib/api';

const MAX_PREVIEW = 25 * 1048576; // larger files: no automatic preview (would load the whole file)
// kinds the preview pane shows with the editor itself (read-only) instead of the small first-page Preview
const PANE_KINDS = new Set(['docs', 'sheets', 'slides', 'pdf']);
const A4_PX = 794; // 210 mm at 96 dpi
/** the user's settings, adjusted for the preview pane: page fitted to the width, no ruler/formula bar/side bar */
export function paneSettings(kind, width) {
  const s = settingsNow();
  const fit = Math.max(30, Math.min(120, Math.floor(((width - 56) / A4_PX) * 100)));
  if (kind === 'docs') return { ...s, docs: { ...s.docs, zoom: fit, ruler: false, showFormatting: false } };
  if (kind === 'sheets') return { ...s, sheets: { ...s.sheets, formulaBar: false, zoom: width < 520 ? 80 : 100 } };
  if (kind === 'pdf') return { ...s, pdf: { ...s.pdf, zoom: 'page-width', sidebar: false } };
  if (kind === 'slides') return { ...s, slides: { ...s.slides, notes: false } };
  return s;
}

/** Minimal host for a read-only Editor (used when a chunk has no Preview) */
export function readOnlyHost(extra = {}) {
  const noop = () => {};
  return {
    setSaver: noop, setExporter: noop, dirty: noop, save: async () => {}, status: noop, close: noop,
    convert: (bytes, from, to) => office.convert(bytes, from, to),
    hasConverter: hasConverterIn(statusNow()),
    download: (bytes, name, mime) => download(bytes, name, mime),
    print: noop,
    lang: uiLang(),
    ...extra,
  };
}

/**
 * Loads bytes (from `bytes` or the file) and the chunk, then shows chunk.Preview → read-only Editor (when `editorFallback`)
 * → toText → icon. Props: file { name, path, root?, size }, bytes? (Uint8Array, e.g. a version), compact (details panel).
 */
export function FilePreviewBox({ file, bytes: given = null, compact = false, editorFallback = false, pane = false, onText }) {
  const ext = extOf(file.name);
  const fmt = formatOf(file.name);
  const [st, setSt] = useState({ phase: 'loading' });
  const [tries, setTries] = useState(0);
  const boxRef = useRef(null);
  const paneRef = useRef(null);
  const [width, setWidth] = useState(0);
  // pane (Files' preview): the editor read-only, sized to the pane – measure it first (fit the page to the width)
  useLayoutEffect(() => {
    if (!pane || !boxRef.current) return undefined;
    const el = boxRef.current;
    const m = () => setWidth((w) => w || el.clientWidth || 0); // the first real width – the page is fitted once
    m();
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(m) : null;
    if (ro) ro.observe(el);
    return () => ro && ro.disconnect();
  }, [pane]);
  useEffect(() => {
    let alive = true;
    setSt({ phase: 'loading' });
    if (!fmt) { setSt({ phase: 'none' }); return undefined; }
    if (!given && file.size > MAX_PREVIEW) { setSt({ phase: 'big' }); return undefined; }
    (async () => {
      const bytes = given || (await office.read(file)).bytes;
      let chunk = null;
      try { chunk = await loadEditor(fmt.kind); } catch { chunk = null; }
      if (!alive) return;
      // the pane shows the real thing: all pages (docs, print layout), the grid with its sheet tabs, the slide list, the PDF
      if (pane && chunk && chunk.Editor && PANE_KINDS.has(fmt.kind)) { setSt({ phase: 'editor', chunk, bytes }); return; }
      if (chunk && chunk.Preview) { setSt({ phase: 'preview', chunk, bytes }); return; }
      if (chunk && editorFallback && chunk.Editor) { setSt({ phase: 'editor', chunk, bytes }); return; }
      let text = '';
      if (chunk && chunk.toText) { try { text = await chunk.toText(bytes, ext); } catch { text = ''; } }
      if (!text && /^(txt|md|markdown|csv|tsv|html?|xml|json)$/.test(ext)) text = new TextDecoder().decode(bytes.slice(0, 20000));
      if (!alive) return;
      if (onText) onText(text);
      setSt(text ? { phase: 'text', text } : { phase: 'none', missing: !chunk });
    })().catch((e) => alive && setSt({ phase: 'error', error: e.message || String(e) }));
    return () => { alive = false; };
  }, [file.path, file.root, file.mtime, given, tries]); // eslint-disable-line react-hooks/exhaustive-deps

  if (pane) {
    // Files' preview pane: fills the pane, the editor read-only (ribbon hidden by CSS), the page fitted to the width
    let inner;
    if (st.phase === 'editor' && width > 0) {
      // settings and host once per opened file (a new object on every resize would re-render the document)
      if (!paneRef.current || paneRef.current.bytes !== st.bytes) paneRef.current = { bytes: st.bytes, settings: paneSettings(fmt.kind, width), host: readOnlyHost(), file: { name: file.name, path: file.path, root: file.root, ext, size: st.bytes.length, mtime: file.mtime } };
      const E = st.chunk.Editor;
      const p = paneRef.current;
      inner = <ErrorBoundary compact where="office preview"><E file={p.file} bytes={st.bytes} readOnly settings={p.settings} host={p.host} /></ErrorBoundary>;
    } else if (st.phase === 'loading' || st.phase === 'editor') inner = <div className="of-shell-pv-load"><Loading inline text="Loading preview …" /></div>;
    else if (st.phase === 'error') inner = <div className="of-shell-pv-msg"><AlertTriangle size={20} aria-hidden="true" /><span className="small">{st.error}</span><Btn className="sm" icon={RotateCcw} onClick={() => setTries((n) => n + 1)}>Try again</Btn></div>;
    else if (st.phase === 'preview') { const P = st.chunk.Preview; inner = <div className="of-shell-pv-scroll"><ErrorBoundary compact where="office preview"><P bytes={st.bytes} ext={ext} name={file.name} settings={settingsNow()} /></ErrorBoundary></div>; }
    else if (st.phase === 'text') inner = <div className="of-shell-pv-scroll"><pre className="of-shell-preview-text">{st.text.slice(0, 60000)}</pre></div>;
    else inner = <div className="of-shell-pv-msg"><FileIcon entry={file} size="lg" /><span className="small muted">{st.phase === 'big' ? t('Too large for a quick preview – open the file to see it.') : st.missing ? t('No preview yet – the editor for this kind of file is not available.') : t('No preview for this file.')}</span></div>;
    return <div ref={boxRef} className={`of-shell-pv is-${fmt ? fmt.kind : 'none'}`}>{inner}</div>;
  }
  if (st.phase === 'loading') return <div className={`of-shell-preview ${compact ? 'is-compact' : ''}`}><Loading inline text="Loading preview …" /></div>;
  if (st.phase === 'error') return <div className="of-shell-preview is-msg"><AlertTriangle size={16} aria-hidden="true" /><span className="small">{st.error}</span><Btn className="xs ghost" icon={RotateCcw} onClick={() => setTries((n) => n + 1)}>Try again</Btn></div>;
  if (st.phase === 'preview') {
    const P = st.chunk.Preview;
    return <div className={`of-shell-preview ${compact ? 'is-compact' : ''}`}><ErrorBoundary compact where="office preview"><P bytes={st.bytes} ext={ext} name={file.name} settings={settingsNow()} /></ErrorBoundary></div>;
  }
  if (st.phase === 'editor') {
    const E = st.chunk.Editor;
    return <div className={`of-shell-preview is-editor ${compact ? 'is-compact' : ''}`}><ErrorBoundary compact where="office preview"><E file={{ name: file.name, path: file.path, root: file.root, ext, size: st.bytes.length, mtime: file.mtime }} bytes={st.bytes} readOnly settings={settingsNow()} host={readOnlyHost()} /></ErrorBoundary></div>;
  }
  if (st.phase === 'text') return <div className={`of-shell-preview is-text ${compact ? 'is-compact' : ''}`}><pre className="of-shell-preview-text">{st.text.slice(0, compact ? 1200 : 20000)}</pre></div>;
  return (
    <div className={`of-shell-preview is-msg ${compact ? 'is-compact' : ''}`}>
      <FileIcon entry={file} size="lg" />
      <span className="small muted">{st.phase === 'big' ? 'Too large for a quick preview – open the file to see it.' : st.missing ? 'No preview yet – the editor for this kind of file is not available.' : 'No preview for this file.'}</span>
    </div>
  );
}
