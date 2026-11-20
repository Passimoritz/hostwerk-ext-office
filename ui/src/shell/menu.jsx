/*
 * Compact drop-down menu of the editor host (More …, Download as …). Hostwerk's contextMenu needs its menu host, which
 * screens (#/office-edit, the editor in front) do not mount – this one renders itself (portal into the page).
 *   popMenu(button, items)  items: { label, icon, hint, onClick, danger } | '-'  (labels already translated)
 *   <PopMenuHost />         once in the editor host
 * Keyboard: ↑/↓/Home/End move, Enter/Space run, Esc closes and gives the focus back to the button.
 */
import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Keep } from './common';

let cur = null;
const subs = new Set();
export function popMenu(btn, items) {
  const r = btn.getBoundingClientRect();
  const list = [];
  for (const x of items.filter(Boolean)) if (x !== '-' || (list.length && list[list.length - 1] !== '-')) list.push(x);
  while (list[list.length - 1] === '-') list.pop();
  cur = { btn, items: list, x: r.right, y: r.bottom + 4, id: Math.random() };
  subs.forEach((f) => f(cur));
}
const close = (focus = true) => { const c = cur; cur = null; subs.forEach((f) => f(null)); if (focus && c && c.btn && c.btn.focus) c.btn.focus(); };

export function PopMenuHost() {
  const [m, setM] = useState(cur);
  const box = useRef(null);
  useEffect(() => { subs.add(setM); return () => { subs.delete(setM); }; }, []);
  useEffect(() => {
    if (!m) return undefined;
    const el = box.current;
    if (el) {
      // keep it inside the window
      const r = el.getBoundingClientRect();
      el.style.left = `${Math.max(8, Math.min(m.x - r.width, window.innerWidth - r.width - 8))}px`;
      el.style.top = `${Math.max(8, Math.min(m.y, window.innerHeight - r.height - 8))}px`;
      const first = el.querySelector('[role=menuitem]');
      if (first) first.focus();
    }
    const down = (e) => { if (el && !el.contains(e.target) && !(m.btn && m.btn.contains(e.target))) close(false); };
    const away = () => close(false);
    document.addEventListener('mousedown', down, true);
    window.addEventListener('resize', away);
    window.addEventListener('blur', away);
    return () => { document.removeEventListener('mousedown', down, true); window.removeEventListener('resize', away); window.removeEventListener('blur', away); };
  }, [m && m.id]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!m) return null;
  const onKey = (e) => {
    const list = [...box.current.querySelectorAll('[role=menuitem]')];
    const i = list.indexOf(document.activeElement);
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); (list[i + 1] || list[0]).focus(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); (list[i - 1] || list[list.length - 1]).focus(); }
    else if (e.key === 'Home') { e.preventDefault(); list[0].focus(); }
    else if (e.key === 'End') { e.preventDefault(); list[list.length - 1].focus(); }
    else if (e.key === 'Tab') close(false);
    e.stopPropagation();
  };
  return createPortal(
    <div ref={box} className="of-shell-pop hw-pop" role="menu" style={{ left: m.x, top: m.y }} onKeyDown={onKey}>
      {m.items.map((x, n) => (x === '-' ? <div key={n} className="of-shell-pop-sep" role="separator" /> : (
        <button key={n} type="button" role="menuitem" className={`of-shell-pop-item ${x.danger ? 'is-danger' : ''}`} onClick={() => { close(false); try { x.onClick && x.onClick(); } catch { /* the action reports itself */ } }}>
          <span className="of-shell-pop-ic" aria-hidden="true">{x.icon ? <x.icon size={16} /> : null}</span>
          <span className="grow">{x.label}</span>
          {x.hint && <kbd className="of-shell-kbd"><Keep>{x.hint}</Keep></kbd>}
        </button>
      )))}
    </div>,
    document.body,
  );
}
