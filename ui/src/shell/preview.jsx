/*
 * Read-only previews: the editor chunk's Preview component (or a short text from toText) for the details panel, the
 * versions sheet and Hostwerk's file browser. A missing chunk is no error here – the icon and "Open" still work.
 */
import { useEffect, useState } from 'react';
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
export function FilePreviewBox({ file, bytes: given = null, compact = false, editorFallback = false, onText }) {
  const ext = extOf(file.name);
  const fmt = formatOf(file.name);
  const [st, setSt] = useState({ phase: 'loading' });
  const [tries, setTries] = useState(0);
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
