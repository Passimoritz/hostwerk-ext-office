/*
 * Building blocks shared by the viewer's parts: the action row (no dead ends), info card, PDF pages (pdf.js),
 * blob URLs, sanitised HTML in a sandboxed frame, sizes and dates.
 */
import { createElement, useEffect, useMemo, useRef, useState } from 'react';
import { Download, Wand2, FileWarning } from 'lucide-react';
import { Btn, Loading, Callout } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { openPdf, renderPageToCanvas, pdfText } from '../_shared/pdfjs.js';
import { FORMATS } from '../../lib/formats.js';
import { Keep } from '../_shared/keep.jsx';

export const fmtSize = (n) => (n == null ? '' : n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : n < 1073741824 ? `${(n / 1048576).toFixed(1)} MB` : `${(n / 1073741824).toFixed(2)} GB`);
export const locale = () => (document.documentElement.lang === 'en' ? 'en-US' : document.documentElement.lang || 'de-DE');
export const fmtDate = (d, withTime = true) => (d ? new Date(d).toLocaleString(locale(), withTime ? { dateStyle: 'medium', timeStyle: 'short' } : { dateStyle: 'medium' }) : '');
export const baseName = (n) => String(n || 'file').replace(/\.[^.]+$/, '');

export function useBlobUrl(bytes, mime) {
  const url = useMemo(() => (bytes ? URL.createObjectURL(new Blob([bytes], { type: mime || 'application/octet-stream' })) : null), [bytes, mime]);
  useEffect(() => () => { if (url) URL.revokeObjectURL(url); }, [url]);
  return url;
}

/** row of next steps under every view: download, convert to an editable copy / PDF, open fully with the engine */
export function ActionRow({ file, bytes, host, children, convertTo, convertLabel, extra }) {
  const fmt = FORMATS[file.ext] || {};
  const download = () => host.download(bytes, file.name, fmt.mime || 'application/octet-stream');
  const convert = async () => {
    const out = await host.convert(bytes, file.ext, convertTo);
    host.download(out, `${baseName(file.name)}.${convertTo}`, (FORMATS[convertTo] || {}).mime || 'application/octet-stream');
  };
  return (
    <div className="of-view-actions">
      {children}
      <span className="grow" />
      {extra}
      {convertTo && <Btn className="sm" icon={Wand2} onClick={convert} title={host.hasConverter ? undefined : t('Uses the Office engine – Hostwerk sets it up when needed')}>{convertLabel || t('Save as {ext}', { ext: convertTo.toUpperCase() })}</Btn>}
      <Btn className="sm" icon={Download} onClick={download}>Download</Btn>
    </div>
  );
}

/** honest card for files that need their desktop app */
export function InfoCard({ icon: Icon, title, text, facts = [], children }) {
  return (
    <div className="card of-view-info">
      <div className="card-head">{Icon && <Icon size={16} />}<h3>{title}</h3></div>
      <div className="card-body stack">
        {text && <p>{text}</p>}
        {facts.length > 0 && <table className="table of-view-facts"><tbody>{facts.filter(([, v]) => v != null && v !== '').map(([k, v]) => <tr key={k}><th>{k}</th><td><Keep>{v}</Keep></td></tr>)}</tbody></table>}
        {children}
      </div>
    </div>
  );
}

/** all pages of a PDF, rendered lazily with pdf.js (used for previews, converted files and Google exports) */
export function PdfPages({ bytes, maxWidth = 900, onText }) {
  const [doc, setDoc] = useState(null);
  const [err, setErr] = useState('');
  useEffect(() => {
    let d = null; let gone = false;
    setDoc(null); setErr('');
    openPdf(bytes).then((x) => { d = x; if (gone) x.destroy(); else setDoc(x); }).catch((e) => !gone && setErr(e.message || String(e)));
    return () => { gone = true; if (d) d.destroy(); };
  }, [bytes]);
  useEffect(() => { if (doc && onText) pdfText(doc, { maxPages: 50 }).then(onText).catch(() => {}); }, [doc]); // eslint-disable-line react-hooks/exhaustive-deps
  if (err) return <Callout kind="err" title={t('The PDF cannot be shown')}>{err}</Callout>;
  if (!doc) return <Loading inline />;
  return <div className="of-view-pdf">{[...Array(Math.min(doc.numPages, 500)).keys()].map((i) => <PdfPage key={i} doc={doc} n={i + 1} maxWidth={maxWidth} />)}</div>;
}
function PdfPage({ doc, n, maxWidth }) {
  const c = useRef(null); const box = useRef(null);
  const [seen, setSeen] = useState(n <= 2);
  const [size, setSize] = useState(null);
  useEffect(() => { doc.getPage(n).then((p) => { const v = p.getViewport({ scale: 1 }); setSize({ w: v.width, h: v.height }); }); }, [doc, n]);
  useEffect(() => {
    const io = new IntersectionObserver((es) => { if (es.some((e) => e.isIntersecting)) setSeen(true); }, { rootMargin: '600px' });
    if (box.current) io.observe(box.current);
    return () => io.disconnect();
  }, []);
  useEffect(() => {
    if (!seen || !size) return;
    const w = Math.min(maxWidth, (box.current && box.current.parentElement.clientWidth - 16) || maxWidth);
    doc.getPage(n).then((p) => renderPageToCanvas(p, c.current, { scale: (w / size.w) * Math.min(2, window.devicePixelRatio || 1) })).catch(() => {});
  }, [seen, size, doc, n, maxWidth]);
  return <div ref={box} className="of-view-pdfpage" style={size ? { aspectRatio: `${size.w} / ${size.h}`, maxWidth } : { maxWidth }}><canvas ref={c} aria-label={t('Page {n}', { n })} /></div>;
}

/** HTML (already sanitised) in a sandboxed frame without scripts; links go to host.openLink; height follows content */
export function SafeFrame({ html, title, name, onLink, minHeight = 200, className = '' }) {
  const ref = useRef(null);
  const [h, setH] = useState(minHeight);
  useEffect(() => {
    const f = ref.current; if (!f) return undefined;
    let ro = null;
    const load = () => {
      const d = f.contentDocument; if (!d) return;
      const fit = () => setH(Math.max(minHeight, d.documentElement.scrollHeight + 4));
      fit();
      try { ro = new ResizeObserver(fit); ro.observe(d.documentElement); } catch { /* old engine */ }
      d.addEventListener('click', (e) => {
        const a = e.target.closest && e.target.closest('a[href]');
        if (!a) return;
        const href = a.getAttribute('href') || '';
        e.preventDefault();
        if (href.startsWith('#')) { const el = d.getElementById(href.slice(1)) || d.getElementsByName(href.slice(1))[0]; if (el) el.scrollIntoView(); return; }
        onLink && onLink(href);
      });
    };
    f.addEventListener('load', load);
    return () => { f.removeEventListener('load', load); if (ro) ro.disconnect(); };
  }, [html]); // eslint-disable-line react-hooks/exhaustive-deps
  // allow-same-origin only (no scripts): needed to measure the height and catch link clicks; content is sanitised
  // createElement: the title (a mail subject …) must not be auto-translated
  return createElement('iframe', { ref, className: `of-view-frame ${className}`, title: name ?? title, sandbox: 'allow-same-origin', srcDoc: html, style: { height: h } });
}

export function ErrorCard({ error, file, bytes, host }) {
  return (
    <div className="of-view-pad stack">
      <Callout kind="err" icon={FileWarning} title={t('This file cannot be shown')}>{String((error && error.message) || error)}</Callout>
      <ActionRow file={file} bytes={bytes} host={host} />
    </div>
  );
}

/** text decoding with BOM / UTF-8 validity check → { text, encoding } */
export function decodeText(u8) {
  if (u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) return { text: new TextDecoder('utf-8').decode(u8.subarray(3)), encoding: 'UTF-8 (BOM)' };
  if (u8[0] === 0xff && u8[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(u8.subarray(2)), encoding: 'UTF-16 LE' };
  if (u8[0] === 0xfe && u8[1] === 0xff) return { text: new TextDecoder('utf-16be').decode(u8.subarray(2)), encoding: 'UTF-16 BE' };
  // UTF-16 without BOM: many zero bytes at odd/even positions
  const n = Math.min(u8.length, 4000); let z0 = 0; let z1 = 0;
  for (let i = 0; i < n; i++) if (!u8[i]) { if (i % 2) z1++; else z0++; }
  if (n > 8 && z1 > n / 4 && z0 < n / 50) return { text: new TextDecoder('utf-16le').decode(u8), encoding: 'UTF-16 LE' };
  if (n > 8 && z0 > n / 4 && z1 < n / 50) return { text: new TextDecoder('utf-16be').decode(u8), encoding: 'UTF-16 BE' };
  try { return { text: new TextDecoder('utf-8', { fatal: true }).decode(u8), encoding: 'UTF-8' }; } catch { return { text: new TextDecoder('windows-1252').decode(u8), encoding: 'Windows-1252' }; }
}
export const looksBinary = (u8) => { const n = Math.min(u8.length, 4000); let bad = 0; for (let i = 0; i < n; i++) { const c = u8[i]; if (c < 9 || (c > 13 && c < 32)) bad++; } return n > 0 && bad / n > 0.1; };
