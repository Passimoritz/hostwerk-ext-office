/*
 * Command palette of the editor host (Alt+Q like Office's search box, or Ctrl+Shift+P when the editor did not use that
 * key): every command the chunk registered (host.palette.register / chunk.commands) plus the file actions of the shell.
 * Type to filter (fuzzy, umlauts folded), ↑/↓, Enter runs, Esc closes; recently used commands come first.
 */
import { memo, useEffect, useMemo, useRef, useState } from 'react';
import { Search, CornerDownLeft } from 'lucide-react';
import { t } from '@hostwerk/i18n';
import { keyLabel } from '../editors/_shared/keys.js';
import { rank, pushRecent } from './paletteModel';
import { Keep } from './common';

const RECENT = 'office:palette-recent';
const readRecent = () => { try { return JSON.parse(localStorage.getItem(RECENT) || '[]'); } catch { return []; } };

function PaletteImpl({ commands, onClose }) {
  const [q, setQ] = useState('');
  const [i, setI] = useState(0);
  const [recent, setRecent] = useState(readRecent);
  const list = useRef(null);
  const input = useRef(null);
  const all = useMemo(() => commands.map((c) => ({ ...c, title: t(c.title), group: c.group ? t(c.group) : '' })), [commands]);
  const shown = useMemo(() => rank(all, q, recent), [all, q, recent]);
  useEffect(() => { setI(0); }, [q]);
  useEffect(() => { const el = list.current && list.current.children[i]; if (el && el.scrollIntoView) el.scrollIntoView({ block: 'nearest' }); }, [i]);
  useEffect(() => { const h = setTimeout(() => input.current && input.current.focus(), 0); return () => clearTimeout(h); }, []);
  const runIt = (c) => {
    if (!c) return;
    const r = pushRecent(recent, c.id);
    setRecent(r);
    try { localStorage.setItem(RECENT, JSON.stringify(r)); } catch { /* private window */ }
    onClose();
    // after the palette is gone, so the command's own dialogs/focus land in the editor
    setTimeout(() => { try { const p = c.run(); if (p && p.catch) p.catch(() => {}); } catch { /* the command reports its own error */ } }, 0);
  };
  const onKey = (e) => {
    if (e.key === 'ArrowDown') { e.preventDefault(); setI((x) => Math.min(shown.length - 1, x + 1)); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); setI((x) => Math.max(0, x - 1)); }
    else if (e.key === 'PageDown') { e.preventDefault(); setI((x) => Math.min(shown.length - 1, x + 8)); }
    else if (e.key === 'PageUp') { e.preventDefault(); setI((x) => Math.max(0, x - 8)); }
    else if (e.key === 'Enter') { e.preventDefault(); runIt(shown[i]); }
    else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); onClose(); }
    e.stopPropagation(); // the editor below must not see the keys
  };
  let lastGroup = null;
  return (
    <div className="of-shell-palette-back" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="of-shell-palette hw-pop" role="dialog" aria-modal="true" aria-label={t('Commands')}>
        <div className="of-shell-palette-in">
          <Search size={16} aria-hidden="true" />
          <input ref={input} className="of-shell-palette-q" value={q} onChange={(e) => setQ(e.target.value)} onKeyDown={onKey} placeholder={t('What do you want to do? – e.g. “table”, “pdf”, “share”')}
            role="combobox" aria-expanded="true" aria-controls="of-palette-list" aria-activedescendant={shown[i] ? `of-pal-${i}` : undefined} aria-label={t('Search commands')} />
          <kbd className="of-shell-kbd">Esc</kbd>
        </div>
        <ul ref={list} id="of-palette-list" className="of-shell-palette-list" role="listbox">
          {shown.map((c, n) => {
            const head = !q && c.group && c.group !== lastGroup ? c.group : null;
            lastGroup = c.group;
            const I = c.icon;
            return (
              <li key={c.id} id={`of-pal-${n}`} role="option" aria-selected={n === i} className={`of-shell-palette-item ${n === i ? 'on' : ''}`} data-group={head || undefined}
                onMouseMove={() => { if (n !== i) setI(n); }} onClick={() => runIt(c)}>
                {head && <span className="of-shell-palette-group">{head}</span>}
                <span className="of-shell-palette-ic" aria-hidden="true">{I ? <I size={16} /> : null}</span>
                <span className="grow ellipsis">{c.title}</span>
                {q && c.group && <span className="tiny muted ellipsis of-shell-palette-g">{c.group}</span>}
                {c.keys && <kbd className="of-shell-kbd"><Keep>{keyLabel(c.keys)}</Keep></kbd>}
                {n === i && <CornerDownLeft size={14} className="muted" aria-hidden="true" />}
              </li>
            );
          })}
          {!shown.length && <li className="of-shell-palette-empty small muted">{t('No command found.')}</li>}
        </ul>
      </div>
    </div>
  );
}
export const Palette = memo(PaletteImpl);
