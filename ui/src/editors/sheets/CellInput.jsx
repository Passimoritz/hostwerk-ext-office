/* Text input for cells and the formula bar; colours references while a formula is typed (mirror behind a transparent textarea). */
import { useLayoutEffect, useRef } from 'react';

export function CellInput({ value, onChange, onKeyDown, inputRef, style, className = '', marks, ariaLabel, readOnly, onFocus, onBlur, onSelect, onCompositionStart, onPaste, onCopy, onCut, spellCheck = false }) {
  const mirror = useRef(null);
  const own = useRef(null);
  const ta = inputRef || own;
  const colored = !!(marks && marks.length && typeof value === 'string' && value.startsWith('='));
  const sync = () => { if (mirror.current && ta.current) { mirror.current.scrollTop = ta.current.scrollTop; mirror.current.scrollLeft = ta.current.scrollLeft; } };
  useLayoutEffect(sync);
  let parts = null;
  if (colored) {
    parts = [];
    let p = 0;
    const sorted = marks.slice().sort((a, b) => a.pos - b.pos);
    for (const m of sorted) {
      const s = m.pos + 1; // marks are positions in the text after '='
      if (s < p) continue;
      if (s > p) parts.push(<span key={`t${p}`}>{value.slice(p, s)}</span>);
      parts.push(<span key={`m${s}`} style={{ color: m.color }}>{value.slice(s, s + m.len)}</span>);
      p = s + m.len;
    }
    parts.push(<span key="end">{value.slice(p)}{'​'}</span>);
  }
  return (
    <div className={`of-sheet-input ${className}`} style={style}>
      {colored && <div ref={mirror} className="of-sheet-input-mirror" aria-hidden="true">{parts}</div>}
      <textarea ref={ta} value={value} readOnly={readOnly} aria-label={ariaLabel} spellCheck={spellCheck} autoComplete="off" autoCorrect="off" autoCapitalize="off"
        className={colored ? 'is-colored' : ''} rows={1} wrap="soft"
        onChange={(e) => onChange && onChange(e.target.value, e.target.selectionStart)}
        onKeyDown={onKeyDown} onScroll={sync} onFocus={onFocus} onBlur={onBlur} onPaste={onPaste} onCopy={onCopy} onCut={onCut}
        onSelect={(e) => onSelect && onSelect(e.target.selectionStart, e.target.selectionEnd)} onCompositionStart={onCompositionStart} />
    </div>
  );
}
