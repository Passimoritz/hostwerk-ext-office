/* Text input for cells and the formula bar; formulas get syntax colours (functions, texts, numbers, errors, names) and
   coloured references while typed (mirror behind a transparent textarea). */
import { useLayoutEffect, useMemo, useRef } from 'react';
import { tokenize } from './formula/parser.js';

const CLS = { func: 'of-fx-fn', str: 'of-fx-str', num: 'of-fx-num', err: 'of-fx-err', bool: 'of-fx-num', name: 'of-fx-name', sref: 'of-fx-name', lp: 'of-fx-paren', rp: 'of-fx-paren' };

export function CellInput({ value, onChange, onKeyDown, inputRef, style, className = '', marks, ariaLabel, readOnly, onFocus, onBlur, onSelect, onCompositionStart, onPaste, onCopy, onCut, spellCheck = false }) {
  const mirror = useRef(null);
  const own = useRef(null);
  const ta = inputRef || own;
  const isF = typeof value === 'string' && value.startsWith('=') && value.length > 1 && value.length < 4000;
  const colored = isF;
  const sync = () => { if (mirror.current && ta.current) { mirror.current.scrollTop = ta.current.scrollTop; mirror.current.scrollLeft = ta.current.scrollLeft; } };
  useLayoutEffect(sync);
  const parts = useMemo(() => {
    if (!colored) return null;
    const byPos = new Map((marks || []).map((m) => [m.pos, m]));
    let toks;
    try { toks = tokenize(value.slice(1), { decimal: /;/.test(value) ? ',' : '.' }); } catch { toks = [{ t: 'bad', text: value.slice(1), pos: 0 }]; }
    const out = [<span key="eq">=</span>];
    let depth = 0;
    for (const tk of toks) {
      const m = tk.t === 'ref' ? byPos.get(tk.pos) : null;
      if (tk.t === 'rp') depth = Math.max(0, depth - 1);
      const cls = tk.t === 'ref' ? null : CLS[tk.t];
      out.push(<span key={tk.pos} className={cls || undefined} style={m ? { color: m.color, fontWeight: 600 } : tk.t === 'lp' || tk.t === 'rp' ? { color: ['#4b5563', '#2f6fe4', '#d9480f', '#2b8a3e'][depth % 4] } : undefined}>{tk.text}</span>);
      if (tk.t === 'lp') depth++;
    }
    out.push(<span key="end">{'​'}</span>);
    return out;
  }, [value, marks, colored]);
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
