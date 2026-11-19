/*
 * Folder tree of the library (navigation and "Move to …" picker). Children load when a folder is opened; the open
 * folders are remembered. Folders accept drops (files from the desktop and entries of the library).
 */
import { useEffect, useState } from 'react';
import { ChevronRight, ChevronDown, Folder, FolderOpen, Library } from 'lucide-react';
import { useRemembered } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { office } from '../lib/api';
import { crumbsOf, sortEntries } from './util';
import { useChanged } from './store';
import { Keep } from './common';

const cache = new Map(); // dir → [names]

export function FolderTree({ value, onSelect, onDrop, id = 'nav', disabled = null, rootLabel = 'Library', counts = null }) {
  const [open, setOpen] = useRemembered(`office:tree:${id}`, ['']);
  const [kids, setKids] = useState(() => Object.fromEntries(cache));
  const [over, setOver] = useState(null);
  const load = async (dir, force = false) => {
    if (!force && cache.has(dir)) return;
    try {
      const r = await office.list({ dir, view: 'all', sort: 'name' });
      const names = sortEntries((r.entries || []).filter((e) => e.dir), 'name').map((e) => e.name);
      cache.set(dir, names);
    } catch { cache.set(dir, []); }
    setKids(Object.fromEntries(cache));
  };
  // the current folder is always visible: open its parents
  useEffect(() => {
    const need = ['', ...crumbsOf(value).map((c) => c.dir).slice(0, -1)];
    const missing = need.filter((d) => !open.includes(d));
    if (missing.length) setOpen([...open, ...missing]);
  }, [value]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { open.forEach((d) => load(d)); }, [open]); // eslint-disable-line react-hooks/exhaustive-deps
  // reload every open folder – ops.js clears the cache (resetTree) before it signals the change, so the cache keys
  // cannot tell which folders were loaded
  useChanged(() => { cache.clear(); open.forEach((d) => load(d, true)); }, [open]);
  const toggle = (dir) => setOpen(open.includes(dir) ? open.filter((d) => d !== dir) : [...open, dir]);
  const dropProps = (dir) => (onDrop ? {
    onDragOver: (e) => { e.preventDefault(); e.dataTransfer.dropEffect = e.dataTransfer.types.includes('Files') ? 'copy' : 'move'; setOver(dir); },
    onDragLeave: () => setOver((o) => (o === dir ? null : o)),
    onDrop: (e) => { e.preventDefault(); e.stopPropagation(); setOver(null); onDrop(dir, e); },
  } : {});
  return <ul className="of-shell-tree" role="tree" aria-label={t('Folders')}><TreeNode dir="" name="" depth={0} c={{ open, kids, over, value, disabled, toggle, onSelect, dropProps, rootLabel, counts }} /></ul>;
}
function TreeNode({ dir, name, depth, c }) {
  const { open, kids, over, value, disabled, toggle, onSelect, dropProps, rootLabel, counts } = c;
  const isOpen = open.includes(dir);
  const children = kids[dir];
  const off = disabled && disabled(dir);
  return (
    <li role="treeitem" aria-expanded={children && children.length === 0 ? undefined : isOpen} aria-selected={value === dir}>
      <div className={`of-shell-tree-row ${value === dir ? 'active' : ''} ${over === dir ? 'is-over' : ''} ${off ? 'is-off' : ''}`} style={{ paddingLeft: `calc(var(--sp-1) + ${depth} * var(--sp-3))` }} {...dropProps(dir)}>
        <button type="button" className="of-shell-tree-tog" tabIndex={-1} aria-label={isOpen ? t('Collapse') : t('Expand')} title={isOpen ? t('Collapse') : t('Expand')} onClick={() => toggle(dir)} disabled={children && children.length === 0}>
          {children && children.length === 0 ? <span className="of-shell-tree-dot" /> : isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
        </button>
        <button type="button" className="of-shell-tree-name" disabled={off} onClick={() => onSelect(dir)} onDoubleClick={() => toggle(dir)} title={dir ? undefined : t(rootLabel)}>
          {dir === '' ? <Library size={14} aria-hidden="true" /> : isOpen ? <FolderOpen size={14} aria-hidden="true" /> : <Folder size={14} aria-hidden="true" />}
          <span className="ellipsis grow">{dir === '' ? t(rootLabel) : <Keep>{name}</Keep>}</span>
          {dir === '' && counts != null && <span className="of-shell-nav-n">{counts}</span>}
        </button>
      </div>
      {isOpen && children && children.length > 0 && (
        <ul role="group" className="of-shell-tree-list">{children.map((n) => { const d = dir ? `${dir}/${n}` : n; return <TreeNode key={d} dir={d} name={n} depth={depth + 1} c={c} />; })}</ul>
      )}
    </li>
  );
}

/** Forget cached folder lists (after creating/moving folders) */
export const resetTree = () => cache.clear();
