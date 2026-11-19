/*
 * Docs editor chunk (ARCHITECTURE.md "Editor chunk contract"): Word (docx/docm/dotx/dotm, doc/dot built in),
 * OpenDocument (odt/ott/fodt), RTF, Markdown, HTML, plain text.
 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { DocsEditor } from './Editor.jsx';
import { templateModel } from './templates.js';
import { readDocument, writeDocument, documentText } from './io/index.js';
import { docToHtml } from './io/render-html.js';
import { pageBox, MM_PX } from './io/pm.js';
import { fontCss } from './Ribbon.jsx';
import './style.css';

async function create({ ext = 'docx', template = 'blank', settings } = {}) {
  const lang = (typeof document !== 'undefined' && document.documentElement.lang) || 'en';
  const model = templateModel(template || 'blank', settings || {}, lang === 'de' ? 'de' : 'en');
  return writeDocument(model, String(ext).toLowerCase());
}

async function toText(bytes, ext) {
  return documentText(bytes, String(ext || 'txt').toLowerCase());
}

/** first page, read-only, paper look (Hostwerk's file browser preview) */
function Preview({ bytes, ext, name }) {
  const [state, setState] = useState(null);
  const ref = useRef(null);
  const [scale, setScale] = useState(0.3);
  useEffect(() => {
    let alive = true;
    const e = String(ext || (name || '').split('.').pop() || 'docx').toLowerCase();
    readDocument(bytes, e).then((r) => {
      if (!alive) return;
      const html = r.html != null ? r.html : docToHtml(r.model.doc, { exportMode: false });
      setState({ html, model: r.model });
    }, (err) => alive && setState({ error: err.message }));
    return () => { alive = false; };
  }, [bytes, ext, name]);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el || !state || !state.model) return undefined;
    const w = pageBox(state.model.page).width * MM_PX;
    const fit = () => setScale(el.clientWidth / w);
    fit();
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(fit) : null;
    if (ro) ro.observe(el);
    return () => ro && ro.disconnect();
  }, [state]);
  if (!state) return <div className="of-docs-preview of-docs-preview-empty" />;
  if (state.error) return <div className="of-docs-preview of-docs-preview-empty small muted">{state.error}</div>;
  const box = pageBox(state.model.page);
  const m = state.model.page.margins;
  const d = state.model.defaults || {};
  return (
    <div ref={ref} className="of-docs-preview" style={{ aspectRatio: `${box.width} / ${box.height}` }}>
      <div className="of-docs-preview-page of-docs-paper" style={{ width: box.width * MM_PX, height: box.height * MM_PX, padding: `${m.top * MM_PX}px ${m.right * MM_PX}px ${m.bottom * MM_PX}px ${m.left * MM_PX}px`, '--of-font': fontCss(d.font || 'Calibri'), '--of-size': `${d.fontSize || 11}pt`, '--of-lh': d.lineHeight || 1.15, '--of-after': `${d.spaceAfter != null ? d.spaceAfter : 8}pt`, background: state.model.page.color || '#fff', transform: `scale(${scale})` }}>
        <div className="of-docs-pm" dangerouslySetInnerHTML={{ __html: state.html }} />
      </div>
    </div>
  );
}

export default { kind: 'docs', Editor: DocsEditor, create, toText, Preview };
