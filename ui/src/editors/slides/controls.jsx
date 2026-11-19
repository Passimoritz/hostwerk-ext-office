/* Small UI pieces of the slides editor: ribbon buttons, popovers, colour picker, number field, shape gallery. */
import { useEffect, useLayoutEffect, useRef, useState, createElement as h } from 'react';
import { createPortal } from 'react-dom';
import { ChevronDown, Ban, Pipette } from 'lucide-react';
import { t } from '@hostwerk/i18n';
import { css, SCHEME_KEYS } from './themes.js';
import { SHAPES, SHAPE_GROUPS, shapePath } from './render/geometry.js';
import { keyLabel } from '../_shared/keys.js';
import { colorName } from '../_shared/colors.js';

/** ribbon button: keeps the text selection (mousedown prevented) */
export function RB({ icon: Icon, label, title, onClick, active, disabled, big, children, className = '', kbd }) {
  const tip = [title || label, keyLabel(kbd)].filter(Boolean).join(' · ');
  return (
    <button type="button" className={`of-slides-rb ${big ? 'big' : ''} ${active ? 'on' : ''} ${className}`} title={tip} aria-label={title || label} aria-pressed={active === undefined ? undefined : !!active}
      disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
      {Icon && <Icon size={big ? 22 : 16} aria-hidden="true" />}
      {label && <span className="of-slides-rb-l">{label}</span>}
      {children}
    </button>
  );
}
export function Group({ label, children }) {
  return <div className="of-slides-rgroup" role="group" aria-label={label}><div className="of-slides-rgroup-body">{children}</div><div className="of-slides-rgroup-label">{label}</div></div>;
}

/** button that opens a popover; content(close) */
export function Pop({ icon, label, title, disabled, content, big, className = '', width, active, split, onMain, kbd }) {
  const [open, setOpen] = useState(false);
  const btn = useRef(null);
  const box = useRef(null);
  const [pos, setPos] = useState({ left: 0, top: 0 });
  useLayoutEffect(() => {
    if (!open || !btn.current) return;
    const r = btn.current.getBoundingClientRect();
    const bw = (box.current && box.current.offsetWidth) || width || 240;
    const bh = (box.current && box.current.offsetHeight) || 200;
    let left = r.left; let top = r.bottom + 4;
    if (left + bw > window.innerWidth - 8) left = Math.max(8, window.innerWidth - bw - 8);
    if (top + bh > window.innerHeight - 8) top = Math.max(8, r.top - bh - 4);
    setPos({ left, top });
  }, [open, width]);
  useEffect(() => {
    if (!open) return undefined;
    const down = (e) => { if (box.current && !box.current.contains(e.target) && btn.current && !btn.current.contains(e.target)) setOpen(false); };
    const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); setOpen(false); } };
    window.addEventListener('mousedown', down, true); window.addEventListener('keydown', key, true);
    return () => { window.removeEventListener('mousedown', down, true); window.removeEventListener('keydown', key, true); };
  }, [open]);
  const host = typeof document !== 'undefined' ? document.querySelector('.app') || document.body : null;
  const Icon = icon;
  return (
    <>
      <span className={`of-slides-popwrap ${split ? 'split' : ''}`}>
        {split && <RB icon={icon} title={title || label} onClick={onMain} disabled={disabled} active={active} big={big} kbd={kbd} />}
        <button type="button" ref={btn} className={`of-slides-rb ${big ? 'big' : ''} ${open ? 'on' : ''} ${active && !split ? 'on' : ''} ${className}`} disabled={disabled} title={title || label} aria-label={title || label} aria-haspopup="true" aria-expanded={open}
          onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen((o) => !o)}>
          {!split && Icon && <Icon size={big ? 22 : 16} aria-hidden="true" />}
          {!split && label && <span className="of-slides-rb-l">{label}</span>}
          <ChevronDown size={12} aria-hidden="true" className="of-slides-chev" />
        </button>
      </span>
      {open && host && createPortal(
        <div ref={box} className="of-slides-pop" style={{ left: pos.left, top: pos.top, width }} onMouseDown={(e) => { if (!/INPUT|SELECT|TEXTAREA/.test(e.target.tagName)) e.preventDefault(); }}>
          {content(() => setOpen(false))}
        </div>, host)}
    </>
  );
}

export function Menu({ items, close }) {
  return (
    <div className="of-slides-menu" role="menu">
      {items.filter(Boolean).map((it, i) => (it === '-' ? <div key={i} className="of-slides-menu-sep" /> : (
        <button key={i} type="button" role="menuitem" className={`of-slides-menu-item ${it.active ? 'on' : ''}`} disabled={it.disabled} onClick={() => { close(); it.onClick(); }}>
          {it.icon ? <it.icon size={14} aria-hidden="true" /> : <span className="of-slides-menu-ic" />}
          <span className="grow">{it.label}</span>
          {it.hint && <span className="muted small">{keyLabel(it.hint)}</span>}
        </button>
      )))}
    </div>
  );
}

const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
const COLS = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
const MODS = ['', '+80', '+60', '+40', '-25', '-50'];
let recent = [];
const REF_NAME = { bg1: 'Background 1', tx1: 'Text 1', bg2: 'Background 2', tx2: 'Text 2' };
/** "accent1+80" → "Accent 1, lighter 80 %", "#FF0000" → "Red · #FF0000" */
function swatchName(ref) {
  const s = String(ref || '');
  if (s.startsWith('#')) return `${t(colorName(s))} · ${s}`;
  const m = /^([a-z]+\d?)([+-]\d+)?$/.exec(s);
  if (!m) return s;
  const base = REF_NAME[m[1]] ? t(REF_NAME[m[1]]) : /^accent(\d)$/.test(m[1]) ? t('Accent {n}', { n: m[1].slice(6) }) : m[1];
  if (!m[2]) return base;
  const p = Math.abs(Number(m[2]));
  return m[2][0] === '+' ? t('{name}, lighter {p} %', { name: base, p }) : t('{name}, darker {p} %', { name: base, p });
}

/** colour picker: theme colours with tints/shades, standard colours, recent, custom, none */
export function ColorGrid({ theme, value, onPick, allowNone = true, noneLabel = 'No colour' }) {
  const pick = (c) => { if (c && !recent.includes(c)) recent = [c, ...recent].slice(0, 10); onPick(c); };
  return (
    <div className="of-slides-colors">
      {allowNone && <button type="button" className="of-slides-menu-item" onClick={() => pick(null)}><Ban size={14} aria-hidden="true" /><span>{t(noneLabel)}</span></button>}
      <div className="small muted">{t('Theme colours')}</div>
      <div className="of-slides-swatches" role="listbox" aria-label={t('Theme colours')}>
        {MODS.map((m) => COLS.map((c) => {
          const ref = `${c}${m}`;
          return <button key={ref} type="button" className={`of-slides-sw ${value === ref ? 'on' : ''}`} style={{ background: css(ref, theme) }} title={swatchName(ref)} aria-label={swatchName(ref)} onClick={() => pick(ref)} />;
        }))}
      </div>
      <div className="small muted">{t('Standard colours')}</div>
      <div className="of-slides-swatches">{STANDARD.map((c) => <button key={c} type="button" className={`of-slides-sw ${value === c ? 'on' : ''}`} style={{ background: c }} title={swatchName(c)} aria-label={swatchName(c)} onClick={() => pick(c)} />)}</div>
      {recent.length > 0 && <><div className="small muted">{t('Recent colours')}</div><div className="of-slides-swatches">{recent.map((c) => <button key={c} type="button" className="of-slides-sw" style={{ background: css(c, theme) }} title={swatchName(c)} aria-label={swatchName(c)} onClick={() => pick(c)} />)}</div></>}
      <label className="of-slides-menu-item"><Pipette size={14} aria-hidden="true" /><span className="grow">{t('More colours …')}</span>
        <input type="color" aria-label={t('More colours …')} defaultValue={/^#/.test(value || '') ? value.slice(0, 7) : '#4472C4'} onChange={(e) => pick(e.target.value.toUpperCase())} /></label>
    </div>
  );
}
export function ColorPop({ icon, label, title, theme, value, onPick, allowNone, noneLabel, disabled, bar }) {
  const Icon = icon;
  const BarIcon = (p) => <span className="of-slides-colicon"><Icon {...p} /><span className="of-slides-colbar" style={{ background: value ? css(value, theme) : 'transparent' }} /></span>;
  return <Pop icon={bar !== false ? BarIcon : icon} label={label} title={title} disabled={disabled} width={252} content={(close) => <ColorGrid theme={theme} value={value} allowNone={allowNone} noneLabel={noneLabel} onPick={(c) => { close(); onPick(c); }} />} />;
}

/** number field that commits on Enter/blur; unit suffix */
export function Num({ value, onChange, step = 1, min, max, unit, label, width = 64, disabled }) {
  const [v, setV] = useState(value == null ? '' : String(value));
  useEffect(() => { setV(value == null ? '' : String(value)); }, [value]);
  const commit = () => { const n = parseFloat(String(v).replace(',', '.')); if (Number.isFinite(n)) onChange(Math.min(max ?? Infinity, Math.max(min ?? -Infinity, n))); else setV(value == null ? '' : String(value)); };
  return (
    <label className="of-slides-num" title={label}>
      <input className="input" style={{ width }} value={v} disabled={disabled} aria-label={label} inputMode="decimal"
        onChange={(e) => setV(e.target.value)} onBlur={commit}
        onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); commit(); } if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); const n = (parseFloat(String(v).replace(',', '.')) || 0) + (e.key === 'ArrowUp' ? step : -step); const r = Math.round(n * 100) / 100; setV(String(r)); onChange(Math.min(max ?? Infinity, Math.max(min ?? -Infinity, r))); } }} />
      {unit && <span className="small muted">{unit}</span>}
    </label>
  );
}

export function ShapeIcon({ geom, size = 22 }) {
  if (geom === 'line' || geom === 'arrow' || geom === 'doubleArrow' || geom === 'elbow' || geom === 'curve') {
    const d = geom === 'elbow' ? 'M3 4L11 4L11 18L19 18' : geom === 'curve' ? 'M3 18C10 18 12 4 19 4' : 'M4 18L18 4';
    return h('svg', { width: size, height: size, viewBox: '0 0 22 22', 'aria-hidden': true }, h('path', { d, fill: 'none', stroke: 'currentColor', strokeWidth: 1.6 }),
      geom !== 'line' && geom !== 'elbow' && geom !== 'curve' ? h('path', { d: 'M18 4L12 5.5L16.5 10Z', fill: 'currentColor' }) : null, geom === 'doubleArrow' ? h('path', { d: 'M4 18L10 16.5L5.5 12Z', fill: 'currentColor' }) : null);
  }
  const p = shapePath(geom, 18, 14);
  return h('svg', { width: size, height: size, viewBox: '-2 -4 22 22', 'aria-hidden': true }, h('path', { d: p.d, fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, fillRule: p.evenodd ? 'evenodd' : undefined }));
}

export const LINE_TOOLS = [{ id: 'line', label: 'Line' }, { id: 'arrow', label: 'Arrow' }, { id: 'doubleArrow', label: 'Double arrow' }, { id: 'elbow', label: 'Elbow connector' }, { id: 'curve', label: 'Curved connector' }];
export function ShapeGallery({ onPick, close }) {
  return (
    <div className="of-slides-gallery">
      <div className="small muted">{t('Lines')}</div>
      <div className="of-slides-shapes">{LINE_TOOLS.map((l) => <button type="button" key={l.id} className="of-slides-shape" title={t(l.label)} aria-label={t(l.label)} onClick={() => { close(); onPick(l.id); }}><ShapeIcon geom={l.id} /></button>)}</div>
      {SHAPE_GROUPS.map((g) => (
        <div key={g.id}>
          <div className="small muted">{t(g.label)}</div>
          <div className="of-slides-shapes">
            {Object.entries(SHAPES).filter(([, s]) => s.group === g.id).map(([k, s]) => <button type="button" key={k} className="of-slides-shape" title={t(s.label)} aria-label={t(s.label)} onClick={() => { close(); onPick(k); }}><ShapeIcon geom={k} /></button>)}
          </div>
        </div>
      ))}
    </div>
  );
}

export const themeSwatch = (theme) => [...SCHEME_KEYS.slice(4, 10)].map((k) => css(k, theme));
