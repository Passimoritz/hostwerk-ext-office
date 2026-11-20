/* Images (png, jpg, gif, webp, bmp, svg, emf/wmf via emf.js – svg sanitised and shown as <img>): fit, 100 %, zoom (wheel/buttons), pan, rotate, info. */
import { useEffect, useMemo, useRef, useState } from 'react';
import DOMPurify from 'dompurify';
import { ZoomIn, ZoomOut, Maximize, RotateCw, RotateCcw, Scan } from 'lucide-react';
import { t } from '@hostwerk/i18n';
import { ActionRow, useBlobUrl, fmtSize } from './common.jsx';
import { FORMATS } from '../../lib/formats.js';
import { Raw, Keep } from '../_shared/keep.jsx';
import { metafileToSvg } from './emf.js';
import { FileTypeIcon } from '../../lib/fileIcons.jsx';

export function imageBytes(bytes, ext) {
  if (ext === 'emf' || ext === 'wmf' || ext === 'emz') {
    // Windows metafiles: drawn as SVG by our small EMF/WMF reader (common records)
    const r = metafileToSvg(bytes);
    if (r) return { bytes: new TextEncoder().encode(r.svg), mime: 'image/svg+xml', converted: true };
  }
  if (ext !== 'svg') return { bytes, mime: (FORMATS[ext] || {}).mime || 'image/png' };
  const clean = DOMPurify.sanitize(new TextDecoder().decode(bytes), { USE_PROFILES: { svg: true, svgFilters: true } });
  return { bytes: new TextEncoder().encode(clean), mime: 'image/svg+xml' };
}

export default function ImageView({ file, bytes, host }) {
  const img = useMemo(() => imageBytes(bytes, file.ext), [bytes, file.ext]);
  const url = useBlobUrl(img.bytes, img.mime);
  const box = useRef(null);
  const [nat, setNat] = useState(null);
  const [zoom, setZoom] = useState('fit');
  const [rot, setRot] = useState(0);
  const [pan, setPan] = useState({ x: 0, y: 0 });
  const [boxSize, setBoxSize] = useState({ w: 800, h: 600 });
  useEffect(() => { const el = box.current; if (!el) return undefined; const ro = new ResizeObserver(() => setBoxSize({ w: el.clientWidth, h: el.clientHeight })); ro.observe(el); return () => ro.disconnect(); }, []);
  const turned = rot % 180 !== 0;
  const fit = nat ? Math.min(1, (boxSize.w - 32) / (turned ? nat.h : nat.w), (boxSize.h - 32) / (turned ? nat.w : nat.h)) : 1;
  const s = zoom === 'fit' ? fit : zoom;
  useEffect(() => { if (nat) host.status(`${nat.w} × ${nat.h} px · ${Math.round(s * 100)} %`); }, [nat, s]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    const el = box.current; if (!el) return undefined;
    const wheel = (e) => { if (!e.ctrlKey && !e.metaKey && zoom === 'fit') return; e.preventDefault(); setZoom((z) => Math.max(0.05, Math.min(32, (z === 'fit' ? fit : z) * (e.deltaY < 0 ? 1.15 : 1 / 1.15)))); };
    el.addEventListener('wheel', wheel, { passive: false });
    return () => el.removeEventListener('wheel', wheel);
  }, [zoom, fit]);
  const drag = (e) => {
    if (zoom === 'fit' || e.button !== 0) return;
    const sx = e.clientX; const sy = e.clientY; const p0 = pan;
    const mv = (ev) => setPan({ x: p0.x + ev.clientX - sx, y: p0.y + ev.clientY - sy });
    const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); };
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
  };
  const setZ = (z) => { setZoom(z); if (z === 'fit') setPan({ x: 0, y: 0 }); };
  const key = (e) => {
    if ((e.ctrlKey || e.metaKey) && (e.key === '+' || e.key === '=')) { e.preventDefault(); setZ(s * 1.25); } else if ((e.ctrlKey || e.metaKey) && e.key === '-') { e.preventDefault(); setZ(Math.max(0.05, s / 1.25)); } else if ((e.ctrlKey || e.metaKey) && e.key === '0') { e.preventDefault(); setZ('fit'); } else if (e.key.toLowerCase() === 'r') setRot((r) => (r + (e.shiftKey ? 270 : 90)) % 360);
  };
  return (
    <div className="of-view-image" onKeyDown={key} tabIndex={-1}>
      <div className="of-view-toolbar">
        <span className="of-view-tname"><FileTypeIcon name={file.name} size={20} /><b className="of-view-ellipsis"><Keep>{file.name}</Keep></b></span>
        <button type="button" className="btn sm ghost icon" title="Zoom out (Ctrl+-)" aria-label="Zoom out" onClick={() => setZ(Math.max(0.05, s / 1.25))}><ZoomOut size={16} /></button>
        <span className="small muted of-view-zoomv">{Math.round(s * 100)} %</span>
        <button type="button" className="btn sm ghost icon" title="Zoom in (Ctrl++)" aria-label="Zoom in" onClick={() => setZ(s * 1.25)}><ZoomIn size={16} /></button>
        <button type="button" className={`btn sm ghost icon${zoom === 'fit' ? ' on' : ''}`} title="Fit to window (Ctrl+0)" aria-label="Fit to window" onClick={() => setZ('fit')}><Maximize size={16} /></button>
        <button type="button" className={`btn sm ghost icon${zoom === 1 ? ' on' : ''}`} title="Actual size" aria-label="Actual size" onClick={() => setZ(1)}><Scan size={16} /></button>
        <button type="button" className="btn sm ghost icon" title="Rotate left (Shift+R)" aria-label="Rotate left" onClick={() => setRot((r) => (r + 270) % 360)}><RotateCcw size={16} /></button>
        <button type="button" className="btn sm ghost icon" title="Rotate right (R)" aria-label="Rotate right" onClick={() => setRot((r) => (r + 90) % 360)}><RotateCw size={16} /></button>
        <span className="grow" />
        <span className="small muted">{nat ? `${nat.w} × ${nat.h} px · ` : ''}{(FORMATS[file.ext] || {}).label ? t((FORMATS[file.ext] || {}).label) : file.ext.toUpperCase()} · {fmtSize(bytes.length)}</span>
      </div>
      <div ref={box} className={`of-view-imgbox${zoom !== 'fit' ? ' pan' : ''}`} onPointerDown={drag}>
        {url && <Raw as="img" src={url} alt={file.name} draggable={false} onLoad={(e) => setNat({ w: e.target.naturalWidth || 300, h: e.target.naturalHeight || 150 })}
          style={nat ? { width: nat.w * s, height: nat.h * s, transform: `translate(${pan.x}px, ${pan.y}px) rotate(${rot}deg)` } : { maxWidth: '100%' }} />}
      </div>
      <ActionRow file={file} bytes={bytes} host={host} />
    </div>
  );
}
