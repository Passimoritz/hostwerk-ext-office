/*
 * Signature dialog: draw (mouse/touch/pen, pressure, smoothing, undo stroke), type (several handwriting styles of the
 * system, name from settings.pdf.signatureName) or an image. The result is a transparent PNG placed on the page like
 * an image. Saved signatures follow the user in Hostwerk (GET/PUT /api/office/signatures – the server route; until it
 * exists, or offline, they stay in this browser's localStorage and move to Hostwerk once the route answers).
 */
import { useEffect, useRef, useState } from 'react';
import { Trash2, Eraser, Upload, Undo2, PenLine, Keyboard, Image as ImageIcon, Bookmark, Cloud, HardDrive } from 'lucide-react';
import { Modal, Btn, Field } from '@hostwerk/ui';
import { SwitchRow } from '../_shared/switchrow.jsx';
import { t } from '@hostwerk/i18n';
import { colorName } from '../_shared/colors.js';
import { Keep } from '../_shared/keep.jsx';

const KEY = 'office-pdf-signatures';
const API = '/api/office/signatures';
const STYLES = [
  { id: 'script', css: "'Segoe Script', 'Brush Script MT', 'Snell Roundhand', 'URW Chancery L', cursive" },
  { id: 'hand', css: "'Lucida Handwriting', 'Bradley Hand', 'Apple Chancery', 'Comic Sans MS', cursive" },
  { id: 'ink', css: "'Ink Free', 'Segoe Print', 'Marker Felt', 'Chalkboard', cursive" },
  { id: 'formal', css: "'Monotype Corsiva', 'Edwardian Script ITC', 'Zapfino', 'Palatino Linotype', serif" },
];

const local = { get: () => { try { return JSON.parse(localStorage.getItem(KEY) || '[]'); } catch { return []; } }, set: (l) => { try { localStorage.setItem(KEY, JSON.stringify(l.slice(0, 8))); } catch { /* storage full or off */ } } };

/** saved signatures → { list, where: 'hostwerk' | 'device' } */
export async function loadSignatures(host) {
  try {
    if (host && host.signatures && typeof host.signatures.list === 'function') return { list: await host.signatures.list(), where: 'hostwerk' };
    const r = await fetch(API, { credentials: 'same-origin' });
    if (r.ok) {
      const j = await r.json(); let list = Array.isArray(j) ? j : j.signatures || [];
      // signatures kept in this browser before move to Hostwerk once
      const old = local.get();
      if (old.length) { const urls = new Set(list.map((s) => s.url)); list = [...list, ...old.filter((s) => !urls.has(s.url))].slice(0, 8); await saveSignatures(host, list, 'hostwerk'); local.set([]); }
      return { list, where: 'hostwerk' };
    }
  } catch { /* offline */ }
  return { list: local.get(), where: 'device' };
}
async function saveSignatures(host, list, where) {
  const l = list.slice(0, 8);
  if (where === 'hostwerk') {
    try {
      if (host && host.signatures && typeof host.signatures.set === 'function') { await host.signatures.set(l); return; }
      const r = await fetch(API, { method: 'PUT', credentials: 'same-origin', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ signatures: l }) });
      if (r.ok) return;
    } catch { /* fall back */ }
  }
  local.set(l);
}

export const dataUrlToBytes = (u) => { const b = atob(u.split(',')[1]); const a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; };

/** crop a canvas to its non-transparent pixels → { url, w, h } */
function trim(c) {
  const ctx = c.getContext('2d', { willReadFrequently: true });
  const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height);
  let x0 = width; let y0 = height; let x1 = -1; let y1 = -1;
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) if (data[(y * width + x) * 4 + 3] > 8) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  if (x1 < 0) return null;
  const pad = 6; x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(width - 1, x1 + pad); y1 = Math.min(height - 1, y1 + pad);
  const o = document.createElement('canvas'); o.width = x1 - x0 + 1; o.height = y1 - y0 + 1;
  o.getContext('2d').drawImage(c, x0, y0, o.width, o.height, 0, 0, o.width, o.height);
  return { url: o.toDataURL('image/png'), w: o.width, h: o.height };
}

function DrawPad({ color, weight, onChange }) {
  const c = useRef(null);
  const strokes = useRef([]); const cur = useRef(null);
  const [count, setCount] = useState(0);
  const W = 560; const H = 200;
  const redraw = () => {
    const cv = c.current; const ctx = cv.getContext('2d');
    ctx.clearRect(0, 0, W, H);
    ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    for (const s of [...strokes.current, ...(cur.current ? [cur.current] : [])]) {
      ctx.strokeStyle = s.color;
      for (let i = 1; i < s.pts.length; i++) {
        const [x0, y0, p0] = s.pts[i - 1]; const [x1, y1, p1] = s.pts[i];
        // smooth: quadratic through the midpoints, width from pressure / speed
        const mx = (x0 + x1) / 2; const my = (y0 + y1) / 2;
        ctx.lineWidth = s.weight * (0.6 + ((p0 + p1) / 2) * 0.9);
        ctx.beginPath();
        if (i === 1) ctx.moveTo(x0, y0); else { const [px, py] = s.pts[i - 2]; ctx.moveTo((px + x0) / 2, (py + y0) / 2); }
        ctx.quadraticCurveTo(x0, y0, mx, my);
        ctx.stroke();
      }
    }
  };
  useEffect(() => { const cv = c.current; const r = window.devicePixelRatio || 1; cv.width = W * r; cv.height = H * r; cv.getContext('2d').scale(r, r); }, []);
  const pos = (e) => { const r = c.current.getBoundingClientRect(); const k = W / r.width; return [(e.clientX - r.left) * k, (e.clientY - r.top) * k, e.pointerType === 'pen' ? (e.pressure || 0.5) : 0.55]; };
  const down = (e) => { c.current.setPointerCapture(e.pointerId); cur.current = { color, weight, pts: [pos(e)] }; };
  const move = (e) => { if (!cur.current) return; const p = pos(e); const l = cur.current.pts[cur.current.pts.length - 1]; if (Math.hypot(p[0] - l[0], p[1] - l[1]) < 1.2) return; cur.current.pts.push(p); redraw(); };
  const up = () => { if (!cur.current) return; if (cur.current.pts.length > 1) strokes.current.push(cur.current); cur.current = null; redraw(); setCount(strokes.current.length); onChange(strokes.current.length ? trim(c.current) : null); };
  const undo = () => { strokes.current.pop(); redraw(); setCount(strokes.current.length); onChange(strokes.current.length ? trim(c.current) : null); };
  const clear = () => { strokes.current = []; redraw(); setCount(0); onChange(null); };
  return (
    <div className="stack">
      <div className="of-pdf-sigpad-wrap">
        <canvas ref={c} className="of-pdf-sigpad" style={{ width: '100%', maxWidth: W, aspectRatio: `${W} / ${H}` }} onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerCancel={up} aria-label="Signature pad" />
        <span className="of-pdf-sigline" aria-hidden="true">✕</span>
      </div>
      <div className="row of-pdf-sigrow"><span className="small muted grow">{count ? t('{n} strokes', { n: count }) : t('Sign with the mouse, a pen or your finger.')}</span>
        <Btn className="sm ghost" icon={Undo2} disabled={!count} onClick={undo}>Undo stroke</Btn><Btn className="sm ghost" icon={Eraser} disabled={!count} onClick={clear}>Clear</Btn></div>
    </div>
  );
}

function typed(name, color, css) {
  const c = document.createElement('canvas'); const r = 2;
  c.width = 1000 * r; c.height = 240 * r;
  const ctx = c.getContext('2d'); ctx.scale(r, r);
  ctx.fillStyle = color; ctx.font = `68px ${css}`; ctx.textBaseline = 'middle';
  ctx.fillText(name, 20, 120);
  return trim(c);
}

export default function SignatureDialog({ host, defaultName, onClose, onPick }) {
  const [saved, setSaved] = useState(() => local.get());
  const [where, setWhere] = useState('device');
  const [tab, setTab] = useState(saved.length ? 'saved' : 'draw');
  const [color, setColor] = useState('#1a237e');
  const [weight, setWeight] = useState(2.4);
  const [drawn, setDrawn] = useState(null);
  const [name, setName] = useState(defaultName || '');
  const [style, setStyle] = useState('script');
  const [img, setImg] = useState(null);
  const [remember, setRemember] = useState(true);
  useEffect(() => { loadSignatures(host).then(({ list, where: w }) => { setSaved(list); setWhere(w); if (list.length && tab === 'draw' && !drawn) setTab('saved'); }); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const css = (STYLES.find((s) => s.id === style) || STYLES[0]).css;
  const result = tab === 'draw' ? drawn : tab === 'type' ? (name.trim() ? typed(name.trim(), color, css) : null) : tab === 'image' ? img : null;
  const store = (list) => { setSaved(list); saveSignatures(host, list, where); };
  const use = (sig) => {
    if (remember && tab !== 'saved') store([{ url: sig.url, w: sig.w, h: sig.h, at: Date.now() }, ...saved.filter((s) => s.url !== sig.url)]);
    onPick({ png: dataUrlToBytes(sig.url), w: sig.w, h: sig.h });
    onClose();
  };
  const pickImage = (f) => {
    if (!f) return;
    const u = URL.createObjectURL(f); const im = new Image();
    im.onload = () => {
      const c = document.createElement('canvas'); const s = Math.min(1, 900 / im.width); c.width = im.width * s; c.height = im.height * s;
      const ctx = c.getContext('2d', { willReadFrequently: true }); ctx.drawImage(im, 0, 0, c.width, c.height); URL.revokeObjectURL(u);
      // a scan on white paper: make the white transparent so it sits on the page like ink
      const d = ctx.getImageData(0, 0, c.width, c.height); const px = d.data;
      for (let i = 0; i < px.length; i += 4) { const l = (px[i] + px[i + 1] + px[i + 2]) / 3; if (l > 225) px[i + 3] = 0; else if (l > 180) px[i + 3] = Math.round(px[i + 3] * (225 - l) / 45); }
      ctx.putImageData(d, 0, 0);
      setImg(trim(c) || { url: c.toDataURL('image/png'), w: c.width, h: c.height });
    };
    im.src = u;
  };
  const tabs = [{ id: 'draw', label: t('Draw'), icon: PenLine }, { id: 'type', label: t('Type in'), icon: Keyboard }, { id: 'image', label: t('Image'), icon: ImageIcon }];
  if (saved.length) tabs.unshift({ id: 'saved', label: t('Saved'), icon: Bookmark });
  return (
    <Modal title="Add signature" subtitle="Click on the page afterwards to place it – drag a corner to resize." onClose={onClose} size="lg"
      footer={<><Btn onClick={onClose}>Cancel</Btn><span className="grow" />{tab !== 'saved' && <SwitchRow className="of-pdf-sigremember" small checked={remember} onChange={setRemember} label={where === 'hostwerk' ? t('Keep in Hostwerk') : t('Remember on this device')} />}<Btn className="primary" disabled={!result} onClick={() => use(result)}>Use signature</Btn></>}>
      <div className="of-pdf-sigtabs" role="tablist">
        {tabs.map((x) => <button key={x.id} type="button" role="tab" aria-selected={tab === x.id} className={`of-pdf-sigtab${tab === x.id ? ' on' : ''}`} onClick={() => setTab(x.id)}><x.icon size={15} /> {x.label}</button>)}
      </div>
      <div className="stack of-pdf-sigbody">
        {tab !== 'saved' && tab !== 'image' && (
          <div className="row">
            {['#1a237e', '#111111', '#0b5d1e', '#8b0000'].map((c) => <button key={c} type="button" className={`of-pdf-swatch lg${color === c ? ' on' : ''}`} style={{ background: c }} aria-label={t(colorName(c))} title={t(colorName(c))} onClick={() => setColor(c)} />)}
            {tab === 'draw' && <><span className="of-pdf-sep" /><span className="small muted">{t('Pen')}</span>{[1.6, 2.4, 3.6].map((w, i) => <button key={w} type="button" className={`of-pdf-weight${weight === w ? ' on' : ''}`} onClick={() => setWeight(w)} title={[t('Fine'), t('Medium'), t('Thick')][i]} aria-label={[t('Fine'), t('Medium'), t('Thick')][i]}><span style={{ height: w * 1.6 }} /></button>)}</>}
          </div>
        )}
        {tab === 'draw' && <DrawPad color={color} weight={weight} onChange={setDrawn} />}
        {tab === 'type' && (
          <>
            <Field label="Your name" hint="Default under Settings → PDF → Signature name"><input className="input" value={name} onChange={(e) => setName(e.target.value)} autoFocus aria-label="Your name" /></Field>
            <div className="of-pdf-sigstyles">
              {STYLES.map((s) => <button key={s.id} type="button" className={`of-pdf-sigstyle${style === s.id ? ' on' : ''}`} onClick={() => setStyle(s.id)} style={{ fontFamily: s.css, color }}><Keep>{name || t('Your name')}</Keep></button>)}
            </div>
          </>
        )}
        {tab === 'image' && (
          <>
            <label className="btn"><Upload size={14} />{t('Choose image …')}<input type="file" accept="image/png,image/jpeg,image/webp,image/gif" hidden onChange={(e) => pickImage(e.target.files[0])} /></label>
            {img && <img className="of-pdf-sigimg" src={img.url} alt="" />}
            <p className="small muted">{t('A scan on white paper works fine – the white background is removed automatically.')}</p>
          </>
        )}
        {tab === 'saved' && (
          <>
            <div className="of-pdf-sigsaved">
              {saved.map((s, i) => (
                <div key={i} className="of-pdf-sigcard">
                  <button type="button" className="of-pdf-sigpick" onClick={() => use(s)} title="Use signature" aria-label="Use signature"><img src={s.url} alt="" /></button>
                  <button type="button" className="btn xs ghost icon of-pdf-sigdel" title="Forget signature" aria-label="Forget signature" onClick={() => store(saved.filter((_, k) => k !== i))}><Trash2 size={12} /></button>
                </div>
              ))}
            </div>
            <span className="small muted row">{where === 'hostwerk' ? <><Cloud size={13} /> {t('Kept in Hostwerk – available on every device you sign in to.')}</> : <><HardDrive size={13} /> {t('Kept in this browser only.')}</>}</span>
          </>
        )}
      </div>
    </Modal>
  );
}
