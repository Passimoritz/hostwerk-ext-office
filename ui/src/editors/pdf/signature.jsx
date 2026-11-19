/*
 * Signature dialog: draw (mouse/touch/pen), type (script-like system font, name from settings.pdf.signatureName) or
 * an image. The result is a transparent PNG placed on the page like an image. Signatures can be remembered on this
 * device (localStorage of the browser/Hostwerk window, at most 6, deletable) – they never leave this computer.
 */
import { useEffect, useRef, useState } from 'react';
import { Trash2, Eraser, Upload } from 'lucide-react';
import { Modal, Btn, Tabs, Field, Switch } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { colorName } from '../_shared/colors.js';

const KEY = 'office-pdf-signatures';
const SCRIPT = "'Segoe Script', 'Brush Script MT', 'Lucida Handwriting', 'Apple Chancery', 'Snell Roundhand', 'URW Chancery L', cursive";

export function savedSignatures() { try { return JSON.parse(localStorage.getItem(KEY) || '[]'); } catch { return []; } }
function storeSignatures(list) { try { localStorage.setItem(KEY, JSON.stringify(list.slice(0, 6))); } catch { /* storage full or off */ } }

export const dataUrlToBytes = (u) => { const b = atob(u.split(',')[1]); const a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; };

/** crop a canvas to its non-transparent pixels → { url, w, h } */
function trim(c) {
  const ctx = c.getContext('2d');
  const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height);
  let x0 = width; let y0 = height; let x1 = -1; let y1 = -1;
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) if (data[(y * width + x) * 4 + 3] > 8) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  if (x1 < 0) return null;
  const pad = 6; x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(width - 1, x1 + pad); y1 = Math.min(height - 1, y1 + pad);
  const o = document.createElement('canvas'); o.width = x1 - x0 + 1; o.height = y1 - y0 + 1;
  o.getContext('2d').drawImage(c, x0, y0, o.width, o.height, 0, 0, o.width, o.height);
  return { url: o.toDataURL('image/png'), w: o.width, h: o.height };
}

function DrawPad({ color, onChange }) {
  const c = useRef(null);
  const last = useRef(null);
  const [empty, setEmpty] = useState(true);
  useEffect(() => { const cv = c.current; const r = window.devicePixelRatio || 1; cv.width = 560 * r; cv.height = 200 * r; cv.getContext('2d').scale(r, r); }, []);
  const pos = (e) => { const r = c.current.getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top, e.pressure || 0.5]; };
  const down = (e) => { c.current.setPointerCapture(e.pointerId); last.current = pos(e); };
  const move = (e) => {
    if (!last.current) return;
    const p = pos(e); const ctx = c.current.getContext('2d');
    ctx.strokeStyle = color; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    ctx.lineWidth = 1.6 + (e.pointerType === 'pen' ? p[2] * 3 : 1.4);
    ctx.beginPath(); ctx.moveTo(last.current[0], last.current[1]); ctx.lineTo(p[0], p[1]); ctx.stroke();
    last.current = p;
    if (empty) setEmpty(false);
  };
  const up = () => { if (!last.current) return; last.current = null; onChange(trim(c.current)); };
  const clear = () => { const cv = c.current; cv.getContext('2d').clearRect(0, 0, cv.width, cv.height); setEmpty(true); onChange(null); };
  return (
    <div className="stack">
      <canvas ref={c} className="of-pdf-sigpad" style={{ width: 560, height: 200 }} onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerCancel={up} aria-label="Signature pad" />
      <div className="row"><span className="small muted grow">{empty ? t('Sign with the mouse, a pen or your finger.') : ''}</span><Btn className="sm ghost" icon={Eraser} onClick={clear}>Clear</Btn></div>
    </div>
  );
}

function typed(name, color) {
  const c = document.createElement('canvas'); const r = 2;
  c.width = 900 * r; c.height = 220 * r;
  const ctx = c.getContext('2d'); ctx.scale(r, r);
  ctx.fillStyle = color; ctx.font = `64px ${SCRIPT}`; ctx.textBaseline = 'middle';
  ctx.fillText(name, 20, 110);
  return trim(c);
}

export default function SignatureDialog({ defaultName, onClose, onPick }) {
  const [tab, setTab] = useState(savedSignatures().length ? 'saved' : 'draw');
  const [color, setColor] = useState('#1a237e');
  const [drawn, setDrawn] = useState(null);
  const [name, setName] = useState(defaultName || '');
  const [img, setImg] = useState(null);
  const [remember, setRemember] = useState(true);
  const [saved, setSaved] = useState(savedSignatures);
  const result = tab === 'draw' ? drawn : tab === 'type' ? (name.trim() ? typed(name.trim(), color) : null) : tab === 'image' ? img : null;
  const use = (sig) => {
    if (remember && tab !== 'saved') { const list = [sig, ...saved.filter((s) => s.url !== sig.url)]; storeSignatures(list); }
    onPick({ png: dataUrlToBytes(sig.url), w: sig.w, h: sig.h });
    onClose();
  };
  const pickImage = (f) => {
    if (!f) return;
    const u = URL.createObjectURL(f); const im = new Image();
    im.onload = () => { const c = document.createElement('canvas'); const s = Math.min(1, 900 / im.width); c.width = im.width * s; c.height = im.height * s; c.getContext('2d').drawImage(im, 0, 0, c.width, c.height); URL.revokeObjectURL(u); setImg(trim(c) || { url: c.toDataURL('image/png'), w: c.width, h: c.height }); };
    im.src = u;
  };
  const tabs = [{ id: 'draw', label: t('Draw') }, { id: 'type', label: t('Type in') }, { id: 'image', label: t('Image') }];
  if (saved.length) tabs.unshift({ id: 'saved', label: t('Saved') });
  return (
    <Modal title="Add signature" subtitle="Click on the page afterwards to place it – drag the corner to resize." onClose={onClose} size="lg"
      footer={<><Btn onClick={onClose}>Cancel</Btn><span className="grow" />{tab !== 'saved' && <Switch checked={remember} onChange={setRemember} label="Remember on this device" />}<Btn className="primary" disabled={!result} onClick={() => use(result)}>Use signature</Btn></>}>
      <Tabs tabs={tabs} value={tab} onChange={setTab} />
      <div className="stack of-pdf-sigbody">
        {tab !== 'saved' && tab !== 'image' && (
          <div className="row">{['#1a237e', '#111111', '#0b5d1e', '#8b0000'].map((c) => <button key={c} type="button" className={`of-pdf-swatch${color === c ? ' on' : ''}`} style={{ background: c }} aria-label={t(colorName(c))} title={t(colorName(c))} onClick={() => setColor(c)} />)}</div>
        )}
        {tab === 'draw' && <DrawPad color={color} onChange={setDrawn} />}
        {tab === 'type' && (
          <>
            <Field label="Your name" hint="Default under Settings → PDF → Signature name"><input className="input" value={name} onChange={(e) => setName(e.target.value)} autoFocus aria-label="Your name" /></Field>
            <div className="of-pdf-sigprev" style={{ fontFamily: SCRIPT, color }}>{name || ' '}</div>
          </>
        )}
        {tab === 'image' && (
          <>
            <label className="btn"><Upload size={14} />{t('Choose image …')}<input type="file" accept="image/png,image/jpeg,image/webp,image/gif" hidden onChange={(e) => pickImage(e.target.files[0])} /></label>
            {img && <img className="of-pdf-sigimg" src={img.url} alt="" />}
            <p className="small muted">{t('A scan on white paper works best – white stays visible, so a PNG with a transparent background looks most natural.')}</p>
          </>
        )}
        {tab === 'saved' && (
          <div className="of-pdf-sigsaved">
            {saved.map((s, i) => (
              <div key={i} className="of-pdf-sigcard">
                <button type="button" className="of-pdf-sigpick" onClick={() => use(s)} title="Use signature" aria-label="Use signature"><img src={s.url} alt="" /></button>
                <button type="button" className="btn xs ghost icon" title="Forget signature" aria-label="Forget signature" onClick={() => { const l = saved.filter((_, k) => k !== i); setSaved(l); storeSignatures(l); }}><Trash2 size={12} /></button>
              </div>
            ))}
          </div>
        )}
      </div>
    </Modal>
  );
}
