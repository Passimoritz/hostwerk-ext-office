/*
 * Small building blocks of the docs ribbon: icon buttons that keep the editor selection, drop-down popovers,
 * colour grid, ribbon groups.
 */
import { useEffect, useRef, useState, useLayoutEffect } from 'react';
import { createPortal } from 'react-dom';
import { ChevronDown } from 'lucide-react';
import { t } from '@hostwerk/i18n';
import { withKeys, keyLabel } from '../_shared/keys.js';
import { colorName } from '../_shared/colors.js';

/** ribbon button: keeps the editor focus (mousedown is prevented), shows its shortcut in the tooltip */
export function RB({ icon: Icon, label, title, keys, active, disabled, onClick, wide, className = '', children }) {
  const tip = withKeys(t(title || label || ''), keys);
  return (
    <button type="button" className={`btn ghost sm ${wide || children ? '' : 'icon'} of-docs-rb ${active ? 'of-on' : ''} ${className}`} title={tip} aria-label={tip} aria-pressed={active === undefined ? undefined : !!active} disabled={disabled}
      onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
      {Icon && <Icon />}{wide && label ? <span>{label}</span> : null}{children}
    </button>
  );
}

/** a drop-down: trigger button + panel (portal, positioned under the trigger) */
export function Drop({ icon: Icon, label, title, keys, disabled, active, children, wide, panelClass = '', split = null, className = '' }) {
  const [open, setOpen] = useState(false);
  const btn = useRef(null);
  const panel = useRef(null);
  const [pos, setPos] = useState(null);
  useLayoutEffect(() => {
    if (!open || !btn.current) return;
    const r = btn.current.getBoundingClientRect();
    const w = panel.current ? panel.current.offsetWidth : 240;
    setPos({ left: Math.max(8, Math.min(r.left, window.innerWidth - w - 8)), top: r.bottom + 4 });
  }, [open]);
  useEffect(() => {
    if (!open) return undefined;
    const down = (e) => { if (panel.current && !panel.current.contains(e.target) && btn.current && !btn.current.contains(e.target)) setOpen(false); };
    const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); setOpen(false); } };
    window.addEventListener('mousedown', down, true); window.addEventListener('keydown', key, true);
    return () => { window.removeEventListener('mousedown', down, true); window.removeEventListener('keydown', key, true); };
  }, [open]);
  const tip = withKeys(t(title || label || ''), keys);
  const host = typeof document !== 'undefined' ? document.querySelector('.app') || document.body : null;
  return (
    <span className={`of-docs-drop ${split ? 'of-split' : ''} ${className}`}>
      {split}
      <button ref={btn} type="button" className={`btn ghost sm ${wide ? '' : 'icon'} of-docs-rb ${active ? 'of-on' : ''} ${split ? 'of-caret' : ''}`} title={tip} aria-label={tip} aria-haspopup="true" aria-expanded={open} disabled={disabled}
        onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen((o) => !o)}>
        {!split && Icon && <Icon />}{wide && label ? <span>{label}</span> : null}<ChevronDown className="of-chev" />
      </button>
      {open && host && createPortal(
        <div ref={panel} className={`of-docs-pop ${panelClass}`} role="menu" style={pos ? { left: pos.left, top: pos.top } : { visibility: 'hidden' }} onMouseDown={(e) => { if (!/INPUT|SELECT|TEXTAREA/.test(e.target.tagName)) e.preventDefault(); }}>
          {typeof children === 'function' ? children(() => setOpen(false)) : children}
        </div>, host)}
    </span>
  );
}

export function MenuItem({ icon: Icon, children, onClick, active, keys, disabled }) {
  return (
    <button type="button" role="menuitem" className={`of-docs-mi ${active ? 'of-on' : ''}`} disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
      {Icon ? <Icon size={14} /> : <span className="of-mi-ic" />}<span className="grow">{children}</span>{keys && <span className="of-mi-keys">{keyLabel(keys)}</span>}
    </button>
  );
}

export const THEME_COLORS = [
  ['#000000', '#FFFFFF', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#C00000'],
  ['#7F7F7F', '#F2F2F2', '#D6DCE4', '#D9E2F3', '#FBE5D5', '#EDEDED', '#FFF2CC', '#DEEBF6', '#E2EFD9', '#FF0000'],
  ['#595959', '#D8D8D8', '#ADB9CA', '#B4C6E7', '#F7CAAC', '#DBDBDB', '#FEE599', '#BDD7EE', '#C5E0B3', '#FFC000'],
  ['#3F3F3F', '#BFBFBF', '#8496B0', '#8EAADB', '#F4B183', '#C9C9C9', '#FFD965', '#9CC3E5', '#A8D08D', '#00B050'],
  ['#262626', '#A5A5A5', '#323F4F', '#2F5496', '#C55A11', '#7B7B7B', '#BF9000', '#2E75B5', '#538135', '#0070C0'],
];
export const HIGHLIGHTS = ['#FFFF00', '#00FF00', '#00FFFF', '#FF00FF', '#FFC000', '#FF0000', '#9CC3E5', '#C5E0B3', '#D9D9D9', '#F7CAAC'];

export function ColorGrid({ colors = THEME_COLORS, onPick, noneLabel = 'Automatic', close }) {
  const rows = Array.isArray(colors[0]) ? colors : [colors];
  return (
    <div className="of-docs-colors">
      <button type="button" className="of-docs-mi" onClick={() => { onPick(null); if (close) close(); }}><span className="of-swatch of-none" /> {t(noneLabel)}</button>
      {rows.map((r, k) => (
        <div className="of-swatch-row" key={k}>
          {r.map((c) => <button type="button" key={c} className="of-swatch" style={{ background: c }} title={`${t(colorName(c))} · ${c}`} aria-label={t(colorName(c))} onClick={() => { onPick(c); if (close) close(); }} />)}
        </div>
      ))}
      <label className="of-docs-mi of-more-colors">{t('More colours …')}<input type="color" onChange={(e) => onPick(e.target.value)} onBlur={() => close && close()} /></label>
    </div>
  );
}

export function Group({ label, children, className = '' }) {
  return (
    <div className={`of-docs-group ${className}`} role="group" aria-label={t(label)}>
      <div className="of-docs-group-body">{children}</div>
      <div className="of-docs-group-label">{label}</div>
    </div>
  );
}
