/*
 * E-book reader (.epub): container.xml → OPF (manifest, spine) → navigation (EPUB 3 nav or NCX). Chapters are
 * sanitised and shown in a sandboxed frame; images, fonts and CSS from the archive become blob URLs. Font size and
 * theme are adjustable; the position is remembered per file on this device.
 */
import { createElement, useEffect, useMemo, useRef, useState } from 'react';
import JSZip from 'jszip';
import DOMPurify from 'dompurify';
import { BookOpen, ChevronLeft, ChevronRight, List, AArrowDown, AArrowUp } from 'lucide-react';
import { Btn, Seg, Loading, Callout } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { parseXml, kids, text as xtext, local } from '../_shared/xmltree.js';
import { ActionRow } from './common.jsx';
import { Keep, Raw } from '../_shared/keep.jsx';
import { FileTypeIcon } from '../../lib/fileIcons.jsx';

const dirOf = (p) => p.slice(0, p.lastIndexOf('/') + 1);
export function resolvePath(base, href) {
  const clean = decodeURIComponent(String(href).split('#')[0].split('?')[0]);
  if (!clean) return base;
  const parts = (clean.startsWith('/') ? clean.slice(1) : dirOf(base) + clean).split('/');
  const out = [];
  for (const p of parts) { if (p === '..') out.pop(); else if (p !== '.' && p !== '') out.push(p); }
  return out.join('/');
}
function byLocal(el, name, out = []) {
  for (const x of (el && el.c) || []) { if (typeof x === 'string') continue; if (local(x.n) === name) out.push(x); byLocal(x, name, out); }
  return out;
}
const firstLocal = (el, name) => byLocal(el, name)[0] || null;

/** OPF + navigation → { title, author, spine: [path], toc: [{ label, path, hash, children }], cover } (pure apart from zip reads) */
export async function readEpub(zip) {
  const container = await zip.file('META-INF/container.xml')?.async('string');
  if (!container) throw new Error(t('This e-book has no META-INF/container.xml.'));
  const rf = firstLocal(parseXml(container), 'rootfile');
  const opfPath = rf && rf.a['full-path'];
  const opfXml = opfPath && await zip.file(opfPath)?.async('string');
  if (!opfXml) throw new Error(t('The package file of this e-book is missing.'));
  const opf = parseXml(opfXml);
  const meta = firstLocal(opf, 'metadata');
  const title = meta ? xtext(firstLocal(meta, 'title') || { c: [] }).trim() : '';
  const author = meta ? byLocal(meta, 'creator').map((x) => xtext(x).trim()).join(', ') : '';
  const manifest = new Map();
  for (const it of byLocal(opf, 'item')) manifest.set(it.a.id, { href: resolvePath(opfPath, it.a.href), type: it.a['media-type'], props: it.a.properties || '' });
  const spineEl = firstLocal(opf, 'spine');
  const spine = spineEl ? byLocal(spineEl, 'itemref').filter((r) => r.a.linear !== 'no').map((r) => manifest.get(r.a.idref)).filter(Boolean).map((m) => m.href) : [];
  let toc = [];
  const navItem = [...manifest.values()].find((m) => /\bnav\b/.test(m.props));
  if (navItem) {
    const nav = parseXml(await zip.file(navItem.href).async('string'));
    const navEl = byLocal(nav, 'nav').find((n) => /toc/.test(n.a['epub:type'] || n.a.type || '')) || firstLocal(nav, 'nav');
    const walk = (ol) => kids(ol).filter((x) => local(x.n) === 'li').map((li) => {
      const a = kids(li).find((x) => local(x.n) === 'a' || local(x.n) === 'span');
      const href = a && a.a.href;
      const sub = kids(li).find((x) => local(x.n) === 'ol');
      return { label: a ? xtext(a).replace(/\s+/g, ' ').trim() : '', path: href ? resolvePath(navItem.href, href) : '', hash: href && href.includes('#') ? href.split('#')[1] : '', children: sub ? walk(sub) : [] };
    });
    const ol = navEl && kids(navEl).find((x) => local(x.n) === 'ol');
    if (ol) toc = walk(ol);
  }
  if (!toc.length) {
    const ncxId = spineEl && spineEl.a.toc;
    const ncx = (ncxId && manifest.get(ncxId)) || [...manifest.values()].find((m) => m.type === 'application/x-dtbncx+xml');
    if (ncx && zip.file(ncx.href)) {
      const doc = parseXml(await zip.file(ncx.href).async('string'));
      const walk = (el) => kids(el).filter((x) => local(x.n) === 'navPoint').map((np) => {
        const src = firstLocal(np, 'content'); const href = src ? src.a.src : '';
        const lbl = firstLocal(np, 'navLabel');
        return { label: lbl ? xtext(lbl).replace(/\s+/g, ' ').trim() : '', path: href ? resolvePath(ncx.href, href) : '', hash: href.includes('#') ? href.split('#')[1] : '', children: walk(np) };
      });
      const map = firstLocal(doc, 'navMap');
      if (map) toc = walk(map);
    }
  }
  const coverMeta = meta && byLocal(meta, 'meta').find((m) => m.a.name === 'cover');
  const coverItem = [...manifest.values()].find((m) => /cover-image/.test(m.props)) || (coverMeta && manifest.get(coverMeta.a.content));
  return { title, author, spine, toc, cover: coverItem ? coverItem.href : null, manifest };
}

const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp', css: 'text/css', otf: 'font/otf', ttf: 'font/ttf', woff: 'font/woff', woff2: 'font/woff2' };
const THEMES = { light: { bg: '#ffffff', fg: '#1b1b1b', link: '#1f5fbf' }, sepia: { bg: '#f4ecd8', fg: '#3b2f1e', link: '#7a4b12' }, dark: { bg: '#1c1c1e', fg: '#e6e6e6', link: '#8ab4ff' } };

export default function EpubView({ file, bytes, host }) {
  const [book, setBook] = useState(null);
  const [err, setErr] = useState(null);
  const zipRef = useRef(null);
  const urls = useRef(new Map());
  const key = `office-epub:${file.root || ''}:${file.path || file.name}`;
  const saved = useMemo(() => { try { return JSON.parse(localStorage.getItem(key) || 'null'); } catch { return null; } }, [key]);
  const [idx, setIdx] = useState(saved ? saved.idx || 0 : 0);
  const [size, setSize] = useState(() => { try { return Number(localStorage.getItem('office-epub-size')) || 110; } catch { return 110; } });
  const [theme, setTheme] = useState(() => { try { return localStorage.getItem('office-epub-theme') || 'light'; } catch { return 'light'; } });
  const [toc, setToc] = useState(false);
  const [html, setHtml] = useState('');
  const [hash, setHash] = useState('');
  const frame = useRef(null);
  const restore = useRef(saved ? saved.y || 0 : 0);

  useEffect(() => {
    let gone = false;
    (async () => {
      try {
        const zip = await JSZip.loadAsync(bytes);
        zipRef.current = zip;
        const b = await readEpub(zip);
        if (!gone) { setBook(b); if (idx >= b.spine.length) setIdx(0); }
      } catch (e) { if (!gone) setErr(e); }
    })();
    return () => { gone = true; for (const u of urls.current.values()) URL.revokeObjectURL(u); urls.current.clear(); };
  }, [bytes]); // eslint-disable-line react-hooks/exhaustive-deps

  const blobUrl = async (path) => {
    if (urls.current.has(path)) return urls.current.get(path);
    const f = zipRef.current.file(path); if (!f) return '';
    const ext = path.split('.').pop().toLowerCase();
    let data = await f.async('uint8array');
    if (ext === 'css') data = new TextEncoder().encode(await rewriteCss(new TextDecoder().decode(data), path));
    const u = URL.createObjectURL(new Blob([data], { type: MIME[ext] || 'application/octet-stream' }));
    urls.current.set(path, u);
    return u;
  };
  const rewriteCss = async (css, base) => {
    const out = []; let last = 0;
    for (const m of css.matchAll(/url\(\s*(['"]?)([^'")]+)\1\s*\)/g)) {
      out.push(css.slice(last, m.index));
      const ref = m[2];
      out.push(/^(data:|blob:)/.test(ref) ? m[0] : /^[a-z]+:/i.test(ref) ? 'none' : `url("${await blobUrl(resolvePath(base, ref))}")`);
      last = m.index + m[0].length;
    }
    out.push(css.slice(last));
    return out.join('').replace(/@import[^;]+;/g, '');
  };

  // build the chapter document
  useEffect(() => {
    if (!book || !book.spine.length) return;
    let gone = false;
    (async () => {
      const path = book.spine[idx];
      const src = await zipRef.current.file(path)?.async('string');
      if (src == null) { setHtml('<p></p>'); return; }
      const doc = new DOMParser().parseFromString(src, /\.x?html?$/i.test(path) && /<\?xml|xmlns=/.test(src) ? 'application/xhtml+xml' : 'text/html');
      const bad = doc.querySelector('parsererror');
      const d = bad ? new DOMParser().parseFromString(src, 'text/html') : doc;
      const styles = [];
      for (const l of d.querySelectorAll('link[rel~="stylesheet"][href]')) { const f = zipRef.current.file(resolvePath(path, l.getAttribute('href'))); if (f) styles.push(await rewriteCss(await f.async('string'), resolvePath(path, l.getAttribute('href')))); }
      for (const s of d.querySelectorAll('style')) styles.push(await rewriteCss(s.textContent, path));
      for (const el of d.querySelectorAll('img[src], image, source[src]')) {
        const attr = el.hasAttribute('src') ? 'src' : el.hasAttribute('href') ? 'href' : 'xlink:href';
        const v = el.getAttribute(attr) || el.getAttributeNS('http://www.w3.org/1999/xlink', 'href');
        if (!v || /^(data:|blob:)/.test(v)) continue;
        const u = /^[a-z]+:/i.test(v) ? '' : await blobUrl(resolvePath(path, v));
        if (el.tagName.toLowerCase() === 'image') { el.setAttribute('href', u); el.removeAttributeNS('http://www.w3.org/1999/xlink', 'href'); } else el.setAttribute(attr, u);
      }
      // internal links → data-of-path
      for (const a of d.querySelectorAll('a[href]')) {
        const h = a.getAttribute('href');
        if (/^[a-z]+:/i.test(h)) continue;
        if (h.startsWith('#')) continue;
        a.setAttribute('href', `#of:${resolvePath(path, h)}${h.includes('#') ? `#${h.split('#')[1]}` : ''}`);
      }
      const body = DOMPurify.sanitize(d.body ? d.body.innerHTML : '', { FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'form', 'input', 'base', 'link', 'meta', 'style'], ADD_URI_SAFE_ATTR: ['href'], ALLOWED_URI_REGEXP: /^(?:(?:https?|mailto|blob|data):|#|[^a-z]|[a-z+.-]+(?:[^a-z+.\-:]|$))/i });
      const th = THEMES[theme] || THEMES.light;
      const css = DOMPurify.sanitize(`<style>${styles.join('\n')}</style>`, { FORCE_BODY: true, ALLOWED_TAGS: ['style'] });
      const page = `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src blob: data:; style-src 'unsafe-inline' blob:; font-src blob: data:;">${css}<style>html{background:${th.bg};color:${th.fg}}body{max-width:42em;margin:0 auto;padding:24px 28px 64px;font:${size}%/1.6 Georgia,'Times New Roman',serif;background:${th.bg} !important;color:${th.fg} !important;overflow-wrap:break-word}img,svg{max-width:100%;height:auto}a{color:${th.link}}*{color:inherit}</style></head><body>${body}</body></html>`;
      if (!gone) setHtml(page);
    })();
    return () => { gone = true; };
  }, [book, idx, size, theme]); // eslint-disable-line react-hooks/exhaustive-deps

  // frame wiring: links, position
  useEffect(() => {
    const f = frame.current; if (!f) return undefined;
    const onLoad = () => {
      const d = f.contentDocument; if (!d) return;
      const win = f.contentWindow;
      if (hash) { const el = d.getElementById(hash); if (el) el.scrollIntoView(); setHash(''); } else if (restore.current) { const y = restore.current; restore.current = 0; win.scrollTo(0, y * (d.documentElement.scrollHeight - win.innerHeight)); }
      d.addEventListener('click', (e) => {
        const a = e.target.closest && e.target.closest('a[href]'); if (!a) return;
        const h = a.getAttribute('href');
        e.preventDefault();
        if (h.startsWith('#of:')) { const [p, hh] = h.slice(4).split('#'); go(p, hh); } else if (h.startsWith('#')) { const el = d.getElementById(h.slice(1)); if (el) el.scrollIntoView(); } else if (/^https?:|^mailto:/i.test(h) && host.openLink) host.openLink(h);
      });
      d.addEventListener('keydown', onKey);
      let tm = 0;
      win.addEventListener('scroll', () => { clearTimeout(tm); tm = setTimeout(() => { try { localStorage.setItem(key, JSON.stringify({ idx, y: win.scrollY / Math.max(1, d.documentElement.scrollHeight - win.innerHeight) })); } catch { /* off */ } }, 300); });
    };
    f.addEventListener('load', onLoad);
    return () => f.removeEventListener('load', onLoad);
  }, [html]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { try { localStorage.setItem(key, JSON.stringify({ idx, y: 0 })); } catch { /* off */ } }, [idx]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (book) host.status(t('Chapter {n} of {count}', { n: idx + 1, count: book.spine.length })); }, [idx, book]); // eslint-disable-line react-hooks/exhaustive-deps

  const go = (path, h) => { const i = book.spine.indexOf(path); if (i >= 0) { setIdx(i); setHash(h || ''); setToc(false); } };
  const onKey = (e) => { if (e.key === 'ArrowRight' || e.key === 'PageDown' && e.altKey) setIdx((i) => Math.min(book.spine.length - 1, i + 1)); else if (e.key === 'ArrowLeft') setIdx((i) => Math.max(0, i - 1)); };
  const setSz = (v) => { setSize(v); try { localStorage.setItem('office-epub-size', String(v)); } catch { /* off */ } };
  const setTh = (v) => { setTheme(v); try { localStorage.setItem('office-epub-theme', v); } catch { /* off */ } };

  if (err) return <div className="of-view-pad"><Callout kind="err" title={t('This e-book cannot be opened')}>{err.message}</Callout><ActionRow file={file} bytes={bytes} host={host} /></div>;
  if (!book) return <Loading text={t('Opening e-book …')} />;
  const TocList = ({ items, depth = 0 }) => <ul className="of-view-toc">{items.map((it, i) => <li key={i}><button type="button" style={{ paddingLeft: 8 + depth * 14 }} className={`of-view-tocitem${book.spine[idx] === it.path ? ' active' : ''}`} onClick={() => go(it.path, it.hash)}><Keep>{it.label || it.path}</Keep></button>{it.children.length > 0 && <TocList items={it.children} depth={depth + 1} />}</li>)}</ul>;

  return (
    <div className="of-view-epub" onKeyDown={onKey} tabIndex={-1}>
      <div className="of-view-toolbar">
        <button type="button" className={`btn sm ghost${toc ? ' on' : ''}`} onClick={() => setToc(!toc)} aria-pressed={toc} title="Table of contents"><List size={16} /> {t('Contents')}</button>
        <Raw as="span" className="of-view-ellipsis" title={[book.title || file.name, book.author].filter(Boolean).join(' · ')}><FileTypeIcon name={file.name} size={20} /> <b><Keep>{book.title || file.name}</Keep></b>{book.author && <span className="muted"> · <Keep>{book.author}</Keep></span>}</Raw>
        <span className="grow" />
        <button type="button" className="btn sm ghost icon" title="Smaller text" aria-label="Smaller text" onClick={() => setSz(Math.max(70, size - 10))}><AArrowDown size={16} /></button>
        <span className="small muted">{size} %</span>
        <button type="button" className="btn sm ghost icon" title="Larger text" aria-label="Larger text" onClick={() => setSz(Math.min(220, size + 10))}><AArrowUp size={16} /></button>
        <Seg value={theme} onChange={setTh} options={[{ id: 'light', label: t('Light') }, { id: 'sepia', label: t('Sepia') }, { id: 'dark', label: t('Dark') }]} />
      </div>
      <div className="of-view-epubbody">
        {toc && <aside className="of-view-list of-view-epubtoc">{book.toc.length ? <TocList items={book.toc} /> : <TocList items={book.spine.map((p, i) => ({ label: t('Chapter {n}', { n: i + 1 }), path: p, hash: '', children: [] }))} />}</aside>}
        {createElement('iframe', { ref: frame, className: 'of-view-epubframe', title: book.title || t('E-book'), sandbox: 'allow-same-origin', srcDoc: html })}
      </div>
      <ActionRow file={file} bytes={bytes} host={host}>
        <Btn className="sm" icon={ChevronLeft} disabled={idx <= 0} onClick={() => setIdx(idx - 1)}>Previous chapter</Btn>
        <span className="small muted">{t('Chapter {n} of {count}', { n: idx + 1, count: book.spine.length })}</span>
        <Btn className="sm" icon={ChevronRight} disabled={idx >= book.spine.length - 1} onClick={() => setIdx(idx + 1)}>Next chapter</Btn>
      </ActionRow>
    </div>
  );
}

export async function epubText(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const b = await readEpub(zip);
  const out = [b.title, b.author].filter(Boolean);
  for (const p of b.spine) { const s = await zip.file(p)?.async('string'); if (s) out.push(s.replace(/<(script|style|head)[\s\S]*?<\/\1>/gi, '').replace(/<br\s*\/?>|<\/(p|div|h\d|li)>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\n{3,}/g, '\n\n').trim()); }
  return out.join('\n\n');
}
export async function epubCover(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const b = await readEpub(zip);
  const f = b.cover && zip.file(b.cover);
  return { title: b.title, author: b.author, cover: f ? await f.async('uint8array') : null, mime: b.cover ? MIME[b.cover.split('.').pop().toLowerCase()] : null };
}
