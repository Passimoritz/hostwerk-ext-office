/* Small UI pieces the kit doesn't have: popover menus, context menu, colour picker, split buttons. */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Check, ChevronDown, ChevronRight } from 'lucide-react';
import { t } from '@hostwerk/i18n';
import { Keep } from '../_shared/keep.jsx';
import { keyLabel, withKeys } from '../_shared/keys.js';
import { colorName } from '../_shared/colors.js';

let openPop = null;
const host = () => (typeof document !== 'undefined' ? document.querySelector('.app') || document.body : null);

/** floating box at an anchor rect or point; closes on outside click / Esc */
export function Floating({ at, onClose, children, className = '', role = 'menu', label }) {
  const ref = useRef(null);
  const [pos, setPos] = useState({ left: -9999, top: -9999 });
  useLayoutEffect(() => {
    const el = ref.current; if (!el) return;
    const w = el.offsetWidth; const h = el.offsetHeight;
    const vw = window.innerWidth; const vh = window.innerHeight;
    let left = at.x; let top = at.y;
    if (at.below !== undefined) { left = at.left; top = at.below; if (top + h > vh - 8 && at.above !== undefined) top = Math.max(8, at.above - h); }
    left = Math.max(8, Math.min(left, vw - w - 8)); top = Math.max(8, Math.min(top, vh - h - 8));
    setPos({ left, top });
  }, [at]);
  useEffect(() => {
    const down = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); };
    const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); onClose(); } };
    const t1 = setTimeout(() => { document.addEventListener('mousedown', down, true); }, 0);
    window.addEventListener('keydown', key, true);
    window.addEventListener('blur', onClose);
    return () => { clearTimeout(t1); document.removeEventListener('mousedown', down, true); window.removeEventListener('keydown', key, true); window.removeEventListener('blur', onClose); };
  }, [onClose]);
  const el = <div ref={ref} className={`of-sheet-pop ${className}`} role={role} aria-label={label} style={pos} onMouseDown={(e) => e.stopPropagation()}>{children}</div>;
  return createPortal(el, host());
}
export const anchorOf = (el) => { const r = el.getBoundingClientRect(); return { left: r.left, below: r.bottom + 2, above: r.top - 2 }; };

/** items: [{ label, icon, onClick, kbd, disabled, checked, items (submenu), danger } | '-' | { header }] */
export function MenuList({ items, onClose }) {
  const [sub, setSub] = useState(null);
  const ref = useRef(null);
  useEffect(() => { const b = ref.current && ref.current.querySelector('button:not([disabled])'); if (b) b.focus(); }, []);
  const key = (e) => {
    const list = [...ref.current.querySelectorAll(':scope > button:not([disabled])')];
    const i = list.indexOf(document.activeElement);
    if (e.key === 'ArrowDown') { e.preventDefault(); (list[i + 1] || list[0]).focus(); }
    if (e.key === 'ArrowUp') { e.preventDefault(); (list[i - 1] || list[list.length - 1]).focus(); }
  };
  return (
    <div ref={ref} className="of-sheet-menu" onKeyDown={key}>
      {items.filter(Boolean).map((it, i) => {
        if (it === '-') return <div key={i} className="of-sheet-menu-sep" role="separator" />;
        if (it.header) return <div key={i} className="of-sheet-menu-head small muted">{it.header}</div>;
        if (it.custom) return <div key={i} className="of-sheet-menu-custom">{it.custom}</div>;
        const I = it.icon;
        return (
          <button key={i} type="button" role="menuitem" className={`of-sheet-menu-item${it.danger ? ' is-danger' : ''}`} disabled={it.disabled}
            onMouseEnter={(e) => setSub(it.items ? { i, at: (() => { const r = e.currentTarget.getBoundingClientRect(); return { x: r.right - 2, y: r.top - 4 }; })() } : null)}
            onClick={(e) => { if (it.items) { const r = e.currentTarget.getBoundingClientRect(); setSub({ i, at: { x: r.right - 2, y: r.top - 4 } }); return; } onClose(); if (it.onClick) it.onClick(); }}>
            <span className="of-sheet-menu-ic">{it.checked ? <Check size={14} /> : I ? <I size={14} /> : null}</span>
            <span className="grow">{it.keep ? <Keep>{it.label}</Keep> : it.label}</span>
            {it.kbd && <kbd>{keyLabel(it.kbd)}</kbd>}
            {it.items && <ChevronRight size={13} />}
          </button>
        );
      })}
      {sub && items[sub.i] && items[sub.i].items && <Floating at={sub.at} onClose={() => setSub(null)}><MenuList items={items[sub.i].items} onClose={() => { setSub(null); onClose(); }} /></Floating>}
    </div>
  );
}

/** imperative context/dropdown menu */
let menuSetter = null;
export function MenuHost() {
  const [m, setM] = useState(null);
  useEffect(() => { menuSetter = setM; return () => { menuSetter = null; }; }, []);
  if (!m) return null;
  return <Floating at={m.at} onClose={() => setM(null)}>{m.content || <MenuList items={m.items} onClose={() => setM(null)} />}</Floating>;
}
export function showMenu(at, items) { if (menuSetter) menuSetter({ at, items }); }
export function showPopover(at, content) { if (menuSetter) menuSetter({ at, content }); }
export function closeMenu() { if (menuSetter) menuSetter(null); }

/** dropdown button */
export function DropBtn({ icon: I, label, title, items, className = '', disabled, children, onOpen }) {
  return (
    <button type="button" className={`btn ghost sm of-sheet-drop ${className}`} title={title || label} aria-label={title || label} aria-haspopup="menu" disabled={disabled}
      onClick={(e) => { const at = anchorOf(e.currentTarget); if (onOpen) onOpen(at); else showMenu(at, typeof items === 'function' ? items() : items); }}>
      {I && <I size={15} />}{children}{label && <span>{label}</span>}<ChevronDown size={12} className="of-sheet-chev" />
    </button>
  );
}

/** button with a main action and a menu arrow */
export function SplitBtn({ icon: I, label, title, onClick, items, onArrow, disabled, kbd, swatch }) {
  return (
    <span className="of-sheet-split">
      <button type="button" className="btn ghost sm" title={kbd ? withKeys(title || label, kbd) : title || label} aria-label={title || label} disabled={disabled} onClick={onClick}>
        {I && <I size={15} />}{swatch !== undefined && <span className="of-sheet-swatch" style={{ background: swatch || 'transparent' }} />}{label && <span>{label}</span>}
      </button>
      <button type="button" className="btn ghost sm of-sheet-split-arrow" aria-label={t('More options')} title={t('More options')} aria-haspopup="menu" disabled={disabled}
        onClick={(e) => { const at = anchorOf(e.currentTarget.parentElement); if (onArrow) onArrow(at); else showMenu(at, typeof items === 'function' ? items() : items); }}><ChevronDown size={12} /></button>
    </span>
  );
}

const THEME = [
  ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'],
  ['#F2F2F2', '#7F7F7F', '#D0CECE', '#D6DCE4', '#D9E2F3', '#FBE5D5', '#EDEDED', '#FFF2CC', '#DEEBF6', '#E2EFD9'],
  ['#D8D8D8', '#595959', '#AEABAB', '#ADB9CA', '#B4C6E7', '#F7CBAC', '#DBDBDB', '#FEE599', '#BDD7EE', '#C5E0B3'],
  ['#BFBFBF', '#3F3F3F', '#757070', '#8496B0', '#8EAADB', '#F4B183', '#C9C9C9', '#FFD965', '#9CC3E5', '#A8D08D'],
  ['#A5A5A5', '#262626', '#3A3838', '#323F4F', '#2F5496', '#C55A11', '#7B7B7B', '#BF9000', '#2E75B5', '#538135'],
  ['#7F7F7F', '#0C0C0C', '#171616', '#222A35', '#1F3864', '#833C0B', '#525252', '#7F6000', '#1E4E79', '#375623'],
];
const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];

export function ColorGrid({ value, onPick, noneLabel = 'No colour', allowNone = true }) {
  // named like Excel's swatches ("Gold · #FFC000"), not just the hex code
  const cell = (c) => { const n = `${t(colorName(c))} · ${c}`; return <button key={c} type="button" className={`of-sheet-color${String(value || '').toUpperCase() === c ? ' is-on' : ''}`} style={{ background: c }} title={n} aria-label={n} onClick={() => onPick(c)} />; };
  return (
    <div className="of-sheet-colors">
      {allowNone && <button type="button" className="btn ghost xs of-sheet-color-none" onClick={() => onPick(null)}>{t(noneLabel)}</button>}
      <div className="of-sheet-color-grid">{THEME.flat().map(cell)}</div>
      <div className="small muted">{t('Standard colours')}</div>
      <div className="of-sheet-color-grid">{STANDARD.map(cell)}</div>
      <label className="row-tight small of-sheet-color-custom">{t('Custom …')}<input type="color" value={/^#[0-9a-f]{6}$/i.test(value || '') ? value : '#000000'} onChange={(e) => onPick(e.target.value.toUpperCase())} aria-label={t('Custom colour')} /></label>
    </div>
  );
}
export function pickColor(at, value, onPick, opts = {}) {
  showPopover(at, <ColorGrid value={value} onPick={(c) => { closeMenu(); onPick(c); }} {...opts} />);
}
