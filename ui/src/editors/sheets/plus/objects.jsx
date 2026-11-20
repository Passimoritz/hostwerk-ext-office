/* Pictures, shapes and slicers floating over the grid: select, move, resize (corner keeps the aspect with Shift), edit text. */
import { useState } from 'react';
import { t } from '@hostwerk/i18n';
import { Keep } from '../../_shared/keep.jsx';

export function ObjBox({ obj, rect, z, selected, readOnly, onSelect, onChange, onMenu, onOpen, className = '', label, children, minW = 16, minH = 16 }) {
  const [live, setLive] = useState(null);
  const r = live || rect;
  const begin = (e, mode) => {
    if (e.button === 2) return;
    e.stopPropagation(); e.preventDefault();
    onSelect();
    if (readOnly) return;
    const sx = e.clientX; const sy = e.clientY; const a0 = { ...obj.anchor }; const ratio = a0.w / Math.max(1, a0.h);
    const calc = (ev) => {
      const dx = (ev.clientX - sx) / z; const dy = (ev.clientY - sy) / z;
      if (mode === 'move') return { dx, dy, w: a0.w, h: a0.h };
      let w = a0.w + (mode.includes('e') ? dx : mode.includes('w') ? -dx : 0);
      let h = a0.h + (mode.includes('s') ? dy : mode.includes('n') ? -dy : 0);
      if (ev.shiftKey && mode.length === 2) h = w / ratio;
      w = Math.max(minW, w); h = Math.max(minH, h);
      return { dx: mode.includes('w') ? a0.w - w : 0, dy: mode.includes('n') ? a0.h - h : 0, w, h };
    };
    const move = (ev) => { const d = calc(ev); move.d = d; setLive({ ...rect, left: rect.left + d.dx * z, top: rect.top + d.dy * z, width: d.w * z, height: d.h * z }); };
    const up = () => {
      window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
      setLive(null);
      const d = move.d; if (!d || (Math.abs(d.dx) < 0.5 && Math.abs(d.dy) < 0.5 && d.w === a0.w && d.h === a0.h)) return;
      onChange({ ...a0, dx: (a0.dx || 0) + d.dx, dy: (a0.dy || 0) + d.dy, w: Math.round(d.w), h: Math.round(d.h), _px: true });
    };
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
  };
  return (
    <div className={`of-sheet-obj ${className}${selected ? ' is-sel' : ''}`} style={r} onPointerDown={(e) => begin(e, 'move')} onDoubleClick={onOpen} tabIndex={-1} aria-label={label}
      onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); onSelect(); if (onMenu) onMenu({ x: e.clientX, y: e.clientY }); }}>
      {children}
      {selected && !readOnly && ['n', 's', 'e', 'w', 'ne', 'nw', 'se', 'sw'].map((m) => <span key={m} className={`of-sheet-handle h-${m}`} onPointerDown={(e) => begin(e, m)} />)}
    </div>
  );
}

/** SVG of a preset shape with its text */
export function ShapeSVG({ shape, w, h, z }) {
  const sw = shape.line ? 1.5 : 0; const p = sw;
  const fill = shape.fill || 'none'; const stroke = shape.line || 'none';
  const W = Math.max(2, w); const H = Math.max(2, h);
  let el;
  switch (shape.kind) {
    case 'ellipse': el = <ellipse cx={W / 2} cy={H / 2} rx={W / 2 - p} ry={H / 2 - p} />; break;
    case 'roundRect': el = <rect x={p} y={p} width={W - 2 * p} height={H - 2 * p} rx={Math.min(W, H) * 0.16} />; break;
    case 'triangle': el = <polygon points={`${W / 2},${p} ${W - p},${H - p} ${p},${H - p}`} />; break;
    case 'diamond': el = <polygon points={`${W / 2},${p} ${W - p},${H / 2} ${W / 2},${H - p} ${p},${H / 2}`} />; break;
    case 'rightArrow': el = <polygon points={`${p},${H * 0.25} ${W * 0.65},${H * 0.25} ${W * 0.65},${p} ${W - p},${H / 2} ${W * 0.65},${H - p} ${W * 0.65},${H * 0.75} ${p},${H * 0.75}`} />; break;
    case 'star5': { const pts = []; for (let i = 0; i < 10; i++) { const rr = i % 2 ? 0.4 : 1; const a = -Math.PI / 2 + (i * Math.PI) / 5; pts.push(`${W / 2 + (W / 2 - p) * rr * Math.cos(a)},${H / 2 + (H / 2 - p) * rr * Math.sin(a)}`); } el = <polygon points={pts.join(' ')} />; break; }
    case 'line': el = <line x1={p} y1={H / 2} x2={W - p} y2={H / 2} />; break;
    case 'wedgeRoundRectCallout': case 'callout': el = <path d={`M${p + 6},${p} H${W - p - 6} Q${W - p},${p} ${W - p},${p + 6} V${H * 0.72} Q${W - p},${H * 0.78} ${W - p - 6},${H * 0.78} H${W * 0.42} L${W * 0.2},${H - p} L${W * 0.28},${H * 0.78} H${p + 6} Q${p},${H * 0.78} ${p},${H * 0.72} V${p + 6} Q${p},${p} ${p + 6},${p} Z`} />; break;
    default: el = <rect x={p / 2} y={p / 2} width={W - p} height={H - p} />;
  }
  const color = shape.textColor || (shape.fill && shape.kind !== 'textbox' ? '#ffffff' : '#000000');
  return (
    <div className="of-sheet-shape" style={{ width: W, height: H }}>
      <svg width={W} height={H} aria-hidden="true"><g fill={shape.kind === 'line' ? 'none' : fill} stroke={stroke} strokeWidth={shape.kind === 'line' ? Math.max(1.5, 2 * z) : sw}>{el}</g></svg>
      {shape.text && <div className={`of-sheet-shape-text${shape.kind === 'textbox' ? ' is-box' : ''}`} style={{ color, fontSize: `${(shape.fontSize || 11) * (4 / 3) * z}px` }}><Keep>{shape.text}</Keep></div>}
    </div>
  );
}

/** a slicer as Excel shows it: caption + item buttons (selected items filled) */
export function SlicerBox({ slicer, items }) {
  return (
    <div className="of-sheet-slicer">
      <div className="of-sheet-slicer-head"><Keep>{slicer.caption || slicer.name}</Keep></div>
      <div className="of-sheet-slicer-items">{(items || []).slice(0, 60).map((it) => <span key={String(it.v)} className={`of-sheet-slicer-item${it.on ? ' is-on' : ''}`}><Keep>{it.v == null || it.v === '' ? t('(blank)') : String(it.v)}</Keep></span>)}</div>
    </div>
  );
}
