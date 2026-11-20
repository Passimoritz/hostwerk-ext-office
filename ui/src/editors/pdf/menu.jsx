/*
 * Right-click menu of the PDF editor: icons, shortcut labels in the interface language (_shared/keys.js), keyboard
 * navigation (↑/↓/Home/End, Enter, Esc), closes on outside click/scroll. Position stays inside the window.
 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { keyLabel } from '../_shared/keys.js';

export default function ContextMenu({ x, y, items, onClose }) {
  const ref = useRef(null);
  const [pos, setPos] = useState({ left: x, top: y });
  const list = items.filter(Boolean);
  useLayoutEffect(() => {
    const el = ref.current; if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({ left: Math.max(4, Math.min(x, window.innerWidth - r.width - 4)), top: Math.max(4, Math.min(y, window.innerHeight - r.height - 4)) });
    const first = el.querySelector('button:not([disabled])');
    if (first) first.focus({ preventScroll: true });
  }, [x, y]);
  useEffect(() => {
    const away = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); };
    const esc = (e) => { if (e.key === 'Escape') { e.stopPropagation(); onClose(); } };
    window.addEventListener('pointerdown', away, true);
    window.addEventListener('keydown', esc, true);
    window.addEventListener('blur', onClose);
    window.addEventListener('wheel', onClose, { passive: true });
    return () => { window.removeEventListener('pointerdown', away, true); window.removeEventListener('keydown', esc, true); window.removeEventListener('blur', onClose); window.removeEventListener('wheel', onClose); };
  }, [onClose]);
  const nav = (e) => {
    const btns = [...ref.current.querySelectorAll('button:not([disabled])')];
    const i = btns.indexOf(document.activeElement);
    if (e.key === 'ArrowDown') { e.preventDefault(); btns[(i + 1) % btns.length].focus(); } else if (e.key === 'ArrowUp') { e.preventDefault(); btns[(i - 1 + btns.length) % btns.length].focus(); } else if (e.key === 'Home') { e.preventDefault(); btns[0].focus(); } else if (e.key === 'End') { e.preventDefault(); btns[btns.length - 1].focus(); }
  };
  return (
    <div ref={ref} className="of-pdf-menu" role="menu" style={pos} onKeyDown={nav} onContextMenu={(e) => e.preventDefault()}>
      {list.map((it, i) => (it === '-' ? <div key={i} className="of-pdf-menu-sep" role="separator" /> : (
        <button key={i} type="button" role="menuitem" className={`of-pdf-menu-item${it.danger ? ' danger' : ''}`} disabled={it.disabled} onClick={() => { onClose(); it.run(); }}>
          <span className="of-pdf-menu-ic">{it.icon ? <it.icon size={15} /> : null}</span>
          <span className="of-pdf-menu-label">{it.label}</span>
          {it.keys && <kbd>{keyLabel(it.keys)}</kbd>}
        </button>
      )))}
    </div>
  );
}
