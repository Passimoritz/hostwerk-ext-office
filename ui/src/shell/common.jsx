/*
 * Small building blocks of the shell: kind-coloured file icons, thumbnails, highlighted search text, dates.
 */
import { createElement, useEffect, useRef, useState } from 'react';
import { FileText, Sheet, Presentation, FileType, File, Folder } from 'lucide-react';
import { apiUrl } from '@hostwerk/api';
import * as HAPI from '@hostwerk/api';
import { toast, confirmDialog } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Keep } from '@hostwerk/kit';
import { FileTypeIcon } from '../lib/fileIcons';
import { thumbSrc, requestThumb, canThumb, forgetThumb } from '../lib/thumbs';
import { highlightParts } from './util';

export const KIND_ICON = { docs: FileText, sheets: Sheet, slides: Presentation, pdf: FileType, viewer: File, folder: Folder };
export const KIND_LABEL = { docs: 'Document', sheets: 'Spreadsheet', slides: 'Presentation', pdf: 'PDF', viewer: 'Other file', folder: 'Folder' };

const PX = { xs: 16, sm: 16, md: 32, lg: 64, xl: 64 };
/** Coloured file icon (lib/fileIcons.jsx) – size 'sm' (list rows, 16 px), 'md' (32 px) or 'lg' (64 px with the extension) or a number */
export function FileIcon({ entry, name, size = 'sm', className = '' }) {
  const px = typeof size === 'number' ? size : PX[size] || 16;
  return <FileTypeIcon entry={entry || { name }} size={px} showExt={px >= 48} className={`of-shell-ficon ${className}`} />;
}

/**
 * First-page thumbnail (embedded or generated, GET /office/thumb) – otherwise the icon. Files without a thumbnail are
 * queued for the generator (lib/thumbs.js: idle time, one at a time) when `generate` is set; the picture fades in.
 */
export function Thumb({ entry, generate = true, size = 'lg' }) {
  const [src, setSrc] = useState(() => thumbSrc(entry));
  const [loaded, setLoaded] = useState(false);
  const ref = useRef(null);
  useEffect(() => { setSrc(thumbSrc(entry)); setLoaded(false); }, [entry.path, entry.root, entry.mtime, entry.thumb]); // eslint-disable-line react-hooks/exhaustive-deps
  // not there yet: render it when the tile comes into view (and the browser is idle)
  useEffect(() => {
    if (src || !generate || entry.dir || !canThumb(entry)) return undefined;
    const el = ref.current;
    let off = null;
    const go = () => { off = requestThumb(entry, (u) => setSrc(u)); };
    if (!el || typeof IntersectionObserver === 'undefined') { go(); return () => off && off(); }
    const io = new IntersectionObserver((xs) => { if (xs.some((x) => x.isIntersecting)) { io.disconnect(); go(); } }, { rootMargin: '200px' });
    io.observe(el);
    return () => { io.disconnect(); if (off) off(); };
  }, [src, entry.path, entry.root, entry.mtime]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <span ref={ref} className={`of-shell-thumb ${src && loaded ? 'is-img' : ''}`}>
      {(!src || !loaded) && <FileIcon entry={entry} size={size} />}
      {src && <img className={`of-shell-thumb-img ${loaded ? 'on' : ''}`} alt="" loading="lazy" decoding="async" src={src} onLoad={() => setLoaded(true)} onError={() => { setSrc(null); forgetThumb(entry); }} />}
    </span>
  );
}

/*
 * User data (file and folder names, tags, notes, metadata) must not go through Hostwerk's automatic JSX translation –
 * a folder called "Templates" would otherwise show as "Vorlagen". Every such text is wrapped in <Keep>. The runtime
 * also translates the props title/label/text/hint/placeholder/aria-label … of every element, so user data never goes
 * into those props: attributes go through <Raw> (an element whose props stay as they are), Highlight takes `value`.
 */
export { Keep };
export function Raw({ as = 'span', ...props }) { return createElement(as, props); }
Raw[Symbol.for('hostwerk.keep')] = true;
/** Text with the search words marked (rendered as text – never as HTML) */
export function Highlight({ value, q }) {
  return <>{highlightParts(value, q).map((p, i) => (p.hit ? <mark key={i} className="of-shell-mark"><Keep>{p.text}</Keep></mark> : <span key={i}><Keep>{p.text}</Keep></span>))}</>;
}

const locale = () => (document.documentElement.lang === 'en' ? 'en-US' : document.documentElement.lang || 'de-DE');
/** "Today 12:03" · "Yesterday 08:10" · "12.03.2026" */
export function fmtWhen(ts) {
  if (!ts) return '';
  const d = new Date(ts); const now = new Date();
  const time = d.toLocaleTimeString(locale(), { hour: '2-digit', minute: '2-digit' });
  const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  if (sameDay(d, now)) return t('Today {time}', { time });
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (sameDay(d, y)) return t('Yesterday {time}', { time });
  return d.toLocaleDateString(locale(), { day: '2-digit', month: '2-digit', year: 'numeric' });
}
export const fmtDateTime = (ts) => (ts ? new Date(ts).toLocaleString(locale(), { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '');
export const fmtTime = (ts) => (ts ? new Date(ts).toLocaleTimeString(locale(), { hour: '2-digit', minute: '2-digit' }) : '');
export const fmtNum = (n) => (n == null ? '' : Number(n).toLocaleString(locale()));

export async function copyText(text) {
  try { await navigator.clipboard.writeText(text); toast.success('Copied'); } catch { toast.error('Could not copy – select the text and press Ctrl+C.'); }
}
/** Download through the browser from a /api/office/… address (works for other machines too – apiUrl) */
export function downloadFrom(href, name = '') {
  const a = document.createElement('a');
  a.href = apiUrl(href);
  if (name) a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
}
/** Debounced value for search fields */
export function useDebounced(value, ms = 300) {
  const [v, setV] = useState(value);
  useEffect(() => { const h = setTimeout(() => setV(value), ms); return () => clearTimeout(h); }, [value, ms]);
  return v;
}
/** Current interface language of Hostwerk */
export const uiLang = () => (document.documentElement.lang === 'en' ? 'en' : 'de');

/**
 * Opens a link INSIDE Hostwerk – never the external browser: own routes (#/…) navigate, other addresses use Hostwerk's
 * in-app web window (openInApp from @hostwerk/api, feature-detected); without it the link is offered to copy.
 */
export async function openLink(href) {
  const u = String(href || '').trim();
  if (!u) return;
  if (u.startsWith('#/')) { location.hash = u.slice(1); return; }
  if (u.startsWith('/') && !u.startsWith('//')) { location.hash = u; return; }
  if (/^mailto:/i.test(u)) { await copyText(u.replace(/^mailto:/i, '')); return; }
  // only with a real Hostwerk window (desktop app) – openInApp falls back to window.open, i.e. a browser tab
  const inApp = typeof HAPI.canOpenInApp === 'function' ? HAPI.canOpenInApp() : !!(window.hostwerk && window.hostwerk.openWebWindow);
  if (inApp && typeof HAPI.openInApp === 'function') { try { await HAPI.openInApp(u); return; } catch (e) { toast.error(e); } }
  const ok = await confirmDialog({ title: 'Open this link?', text: t('{url}\n\nHostwerk cannot show it in its own window here. Copy the link instead?', { url: u }), confirm: 'Copy link' });
  if (ok) await copyText(u);
}
/** Detached Hostwerk window (Hostwerk's main.js opens own-origin addresses with window=1 as Hostwerk windows) */
export function openHostwerkWindow(route) {
  const r = String(route || '').replace(/^#?\/?/, '/');
  const sep = r.includes('?') ? '&' : '?';
  const w = window.open(`${location.pathname}${location.search}#${r}${sep}window=1`, '_blank');
  if (!w) location.hash = r; // pop-ups blocked (plain browser): stay in this window
}
