/*
 * Small building blocks of the shell: kind-coloured file icons, thumbnails, highlighted search text, dates.
 */
import { createElement, useEffect, useState } from 'react';
import { FileText, Sheet, Presentation, FileType, File, Folder, FileArchive, Mail, Image as ImageIcon, BookOpen, Contact, CalendarDays, Link2 } from 'lucide-react';
import { apiUrl } from '@hostwerk/api';
import * as HAPI from '@hostwerk/api';
import { toast, confirmDialog } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { Keep } from '@hostwerk/kit';
import { FORMATS, extOf } from '../lib/formats';
import { url } from '../lib/api';
import { highlightParts, kindOfEntry } from './util';

export const KIND_ICON = { docs: FileText, sheets: Sheet, slides: Presentation, pdf: FileType, viewer: File, folder: Folder };
const VIEWER_ICON = { eml: Mail, msg: Mail, vcf: Contact, ics: CalendarDays, epub: BookOpen, zip: FileArchive, png: ImageIcon, jpg: ImageIcon, jpeg: ImageIcon, gif: ImageIcon, webp: ImageIcon, bmp: ImageIcon, svg: ImageIcon };
export const KIND_LABEL = { docs: 'Document', sheets: 'Spreadsheet', slides: 'Presentation', pdf: 'PDF', viewer: 'Other file', folder: 'Folder' };

/** Kind-coloured icon with the extension underneath – size 'sm' (list rows), 'md' (details) or 'lg' (grid tiles) */
export function FileIcon({ entry, name, size = 'sm', className = '' }) {
  const e = entry || { name };
  const kind = kindOfEntry(e);
  const ext = e.dir ? '' : extOf(e.name);
  const google = (FORMATS[ext] || {}).family === 'google';
  const I = kind === 'viewer' ? (google ? Link2 : VIEWER_ICON[ext] || File) : KIND_ICON[kind] || File;
  const px = size === 'lg' ? 30 : size === 'md' ? 22 : 16;
  return (
    <span className={`of-shell-icon of-shell-k-${kind} of-shell-icon-${size} ${className}`} aria-hidden="true">
      <I size={px} />
      {size !== 'sm' && ext && <span className="of-shell-icon-ext">{ext.slice(0, 5)}</span>}
    </span>
  );
}

/** Embedded thumbnail (GET /office/thumb) when the file has one, otherwise the icon */
export function Thumb({ entry }) {
  const [bad, setBad] = useState(false);
  useEffect(() => setBad(false), [entry.path, entry.mtime]);
  if (!entry.thumb || bad || entry.dir || entry.root) return <FileIcon entry={entry} size="lg" />;
  return <img className="of-shell-thumb-img" alt="" loading="lazy" src={apiUrl(url('/thumb', { path: entry.path, v: entry.mtime || '' }))} onError={() => setBad(true)} />;
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
