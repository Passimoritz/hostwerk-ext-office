/*
 * Sidebar: page thumbnails (select, drag to reorder, rotate/delete), outline (bookmarks), attachments, digital
 * signatures (information only).
 */
import { useEffect, useRef, useState, memo } from 'react';
import { RotateCcw, RotateCw, Trash2, ChevronRight, ChevronDown, Paperclip, Download, Bookmark, FileSignature, Image as ImageIcon } from 'lucide-react';
import { Seg, TileEmpty } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { pageRotation } from './Viewer.jsx';
import { Keep, Raw } from '../_shared/keep.jsx';

const Thumb = memo(function Thumb({ mp, n, proxy, active, selected, editable, onClick, onRotate, onDelete, dragProps }) {
  const canvas = useRef(null);
  const box = useRef(null);
  const [seen, setSeen] = useState(false);
  const rot = pageRotation(proxy, mp, 0);
  useEffect(() => {
    const io = new IntersectionObserver((es) => { if (es.some((e) => e.isIntersecting)) setSeen(true); }, { rootMargin: '200px' });
    io.observe(box.current);
    return () => io.disconnect();
  }, []);
  useEffect(() => {
    if (!seen || !proxy) return undefined;
    let task = null; let gone = false;
    const base = proxy.getViewport({ scale: 1, rotation: rot });
    const vp = proxy.getViewport({ scale: (110 * Math.min(2, window.devicePixelRatio || 1)) / base.width, rotation: rot });
    const c = canvas.current;
    c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
    task = proxy.render({ canvasContext: c.getContext('2d'), viewport: vp, background: '#ffffff' });
    task.promise.catch(() => {});
    return () => { gone = true; try { task.cancel(); } catch { /* done */ } return gone; };
  }, [seen, proxy, rot]);
  const base = proxy ? proxy.getViewport({ scale: 1, rotation: rot }) : { width: 595, height: 842 };
  return (
    <div ref={box} className={`of-pdf-thumb${active ? ' active' : ''}${selected ? ' selected' : ''}`} {...dragProps}>
      <button type="button" className="of-pdf-thumb-img" onClick={onClick} title={t('Page {n}', { n })} aria-label={t('Page {n}', { n })} aria-current={active || undefined}>
        <canvas ref={canvas} style={{ aspectRatio: `${base.width} / ${base.height}` }} />
      </button>
      <div className="of-pdf-thumb-bar">
        <span className="small muted">{n}</span>
        {editable && (
          <span className="of-pdf-thumb-tools">
            <button type="button" className="btn xs ghost icon" title="Rotate left" aria-label="Rotate left" onClick={() => onRotate(-90)}><RotateCcw size={12} /></button>
            <button type="button" className="btn xs ghost icon" title="Rotate right" aria-label="Rotate right" onClick={() => onRotate(90)}><RotateCw size={12} /></button>
            <button type="button" className="btn xs ghost icon" title="Delete page" aria-label="Delete page" onClick={onDelete}><Trash2 size={12} /></button>
          </span>
        )}
      </div>
    </div>
  );
});

function OutlineItem({ item, onGo, depth = 0 }) {
  const [open, setOpen] = useState(depth < 1);
  const has = item.items && item.items.length > 0;
  return (
    <li>
      <div className="of-pdf-ol-row" style={{ paddingLeft: depth * 12 }}>
        {has ? <button type="button" className="btn xs ghost icon" aria-label={open ? 'Collapse' : 'Expand'} title={open ? 'Collapse' : 'Expand'} onClick={() => setOpen(!open)}>{open ? <ChevronDown size={12} /> : <ChevronRight size={12} />}</button> : <span className="of-pdf-ol-pad" />}
        <button type="button" className={`of-pdf-ol-link${item.bold ? ' b' : ''}${item.italic ? ' i' : ''}`} onClick={() => onGo(item)}><Keep>{item.title}</Keep></button>
      </div>
      {has && open && <ul>{item.items.map((c, i) => <OutlineItem key={i} item={c} onGo={onGo} depth={depth + 1} />)}</ul>}
    </li>
  );
}

export default function Sidebar({
  tab, onTab, pages, proxies, current, selection, onSelect, editable, onMove, onRotate, onDelete, outline, onOutline,
  attachments, onAttachment, signatures,
}) {
  const [drag, setDrag] = useState(null);
  const [over, setOver] = useState(null);
  const tabs = [{ id: 'pages', label: t('Pages'), icon: ImageIcon }, { id: 'outline', label: t('Bookmarks'), icon: Bookmark }];
  if (attachments && attachments.length) tabs.push({ id: 'files', label: t('Attachments'), icon: Paperclip });
  if (signatures && signatures.length) tabs.push({ id: 'sign', label: t('Signatures'), icon: FileSignature });
  const click = (mp, e) => {
    if (e.shiftKey && selection.size) {
      const ids = pages.map((p) => p.id); const last = [...selection].pop();
      const [a, b] = [ids.indexOf(last), ids.indexOf(mp.id)].sort((x, y) => x - y);
      onSelect(new Set(ids.slice(a, b + 1)), mp.id);
    } else if (e.ctrlKey || e.metaKey) {
      const s = new Set(selection); if (s.has(mp.id)) s.delete(mp.id); else s.add(mp.id); onSelect(s, mp.id);
    } else onSelect(new Set([mp.id]), mp.id);
  };
  return (
    <aside className="of-pdf-side" aria-label="Sidebar">
      <div className="of-pdf-side-head"><Seg options={tabs.map((x) => ({ id: x.id, label: x.label, title: x.label }))} value={tabs.some((x) => x.id === tab) ? tab : 'pages'} onChange={onTab} /></div>
      <div className="of-pdf-side-body">
        {(tab === 'pages' || !tabs.some((x) => x.id === tab)) && (
          <div className="of-pdf-thumbs" role="list">
            {pages.map((mp, i) => (
              <Thumb key={mp.id} mp={mp} n={i + 1} proxy={proxies.get(`${mp.src}:${mp.index}`)} active={mp.id === current} selected={selection.has(mp.id) && selection.size > 1}
                editable={editable} onClick={(e) => click(mp, e)} onRotate={(d) => onRotate([mp.id], d)} onDelete={() => onDelete([mp.id])}
                dragProps={editable ? {
                  draggable: true, role: 'listitem',
                  onDragStart: (e) => { setDrag(mp.id); e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', mp.id); } catch { /* ok */ } },
                  onDragOver: (e) => { if (!drag) return; e.preventDefault(); const r = e.currentTarget.getBoundingClientRect(); setOver({ id: mp.id, after: e.clientY > r.top + r.height / 2 }); },
                  onDragLeave: () => setOver(null),
                  onDrop: (e) => { e.preventDefault(); if (drag && over) onMove(selection.has(drag) ? [...selection] : [drag], over.id, over.after); setDrag(null); setOver(null); },
                  onDragEnd: () => { setDrag(null); setOver(null); },
                  'data-drop': over && over.id === mp.id ? (over.after ? 'after' : 'before') : undefined,
                } : { role: 'listitem' }} />
            ))}
          </div>
        )}
        {tab === 'outline' && (outline && outline.length
          ? <ul className="of-pdf-outline">{outline.map((it, i) => <OutlineItem key={i} item={it} onGo={onOutline} />)}</ul>
          : <TileEmpty icon={Bookmark} text={t('This PDF has no bookmarks.')} />)}
        {tab === 'files' && (
          <ul className="of-pdf-files">
            {attachments.map((a, i) => (
              <li key={i}>
                <Paperclip size={14} />
                <Raw className="grow of-pdf-ellipsis" title={a.filename}><Keep>{a.filename}</Keep><span className="small muted"> · {fmtSize(a.content.length)}</span></Raw>
                <button type="button" className="btn xs ghost icon" title="Download" aria-label="Download" onClick={() => onAttachment(a)}><Download size={12} /></button>
              </li>
            ))}
          </ul>
        )}
        {tab === 'sign' && (
          <div className="stack of-pdf-sigs">
            <p className="small muted">{t('Digital signatures are shown for information – their validity is not checked here. Saving changes makes them invalid.')}</p>
            {signatures.map((s, i) => (
              <div key={i} className="card of-pdf-sig">
                <div className="row"><FileSignature size={14} /><b className="grow"><Keep>{s.name || s.field}</Keep></b>{s.signed ? <span className="badge ok">{t('Signed')}</span> : <span className="badge">{t('Empty')}</span>}</div>
                {s.date && <div className="small muted">{fmtPdfDate(s.date)}</div>}
                {s.reason && <div className="small">{t('Reason')}: <Keep>{s.reason}</Keep></div>}
                {s.location && <div className="small">{t('Location')}: <Keep>{s.location}</Keep></div>}
                {s.filter && <div className="small muted mono">{s.filter}</div>}
              </div>
            ))}
          </div>
        )}
      </div>
    </aside>
  );
}

export const fmtSize = (n) => (n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(1)} MB`);
export function fmtPdfDate(s) {
  const m = /^D?:?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?/.exec(String(s || ''));
  if (!m) return String(s || '');
  const d = new Date(Date.UTC(+m[1], (+m[2] || 1) - 1, +m[3] || 1, +m[4] || 0, +m[5] || 0, +m[6] || 0));
  return d.toLocaleString(document.documentElement.lang === 'en' ? 'en-US' : 'de-DE');
}
