/*
 * Simple equations: a LaTeX-ish linear text ("x^2 + \frac{a}{b} = \sqrt{c}") shown as formatted HTML and written to
 * other formats as Unicode linear text ("x² + a/b = √c"). Not a TeX engine – the common 90 % of school/office maths.
 */
const SYM = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  cdot: '·', times: '×', div: '÷', pm: '±', mp: '∓', leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', approx: '≈', equiv: '≡', sim: '∼', propto: '∝',
  infty: '∞', sum: '∑', prod: '∏', int: '∫', oint: '∮', partial: '∂', nabla: '∇', forall: '∀', exists: '∃', in: '∈', notin: '∉', subset: '⊂', supset: '⊃', cup: '∪', cap: '∩', emptyset: '∅',
  to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftrightarrow: '⇔', leftrightarrow: '↔', degree: '°', circ: '∘', ldots: '…', cdots: '⋯', angle: '∠', perp: '⊥', parallel: '∥',
  lim: 'lim', sin: 'sin', cos: 'cos', tan: 'tan', log: 'log', ln: 'ln', exp: 'exp', max: 'max', min: 'min',
};
const SUP = { 0: '⁰', 1: '¹', 2: '²', 3: '³', 4: '⁴', 5: '⁵', 6: '⁶', 7: '⁷', 8: '⁸', 9: '⁹', '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾', n: 'ⁿ', i: 'ⁱ' };
const SUB = { 0: '₀', 1: '₁', 2: '₂', 3: '₃', 4: '₄', 5: '₅', 6: '₆', 7: '₇', 8: '₈', 9: '₉', '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎' };

/** tokenise into a tiny tree: [string | { cmd, args: [tree…] } | { sup: tree } | { sub: tree }] */
function parse(src) {
  let i = 0;
  const s = String(src || '');
  const group = () => {
    while (s[i] === ' ') i++;
    if (s[i] === '{') { i++; const t = seq('}'); i++; return t; }
    if (s[i] === '\\') { const t = []; atom(t); return t; }
    const ch = s[i++]; return ch === undefined ? [] : [ch];
  };
  const atom = (out) => {
    if (s[i] === '\\') {
      i++;
      const m = /^[a-zA-Z]+/.exec(s.slice(i));
      if (!m) { out.push(s[i++] || ''); return; }
      i += m[0].length;
      const cmd = m[0];
      if (cmd === 'frac') out.push({ cmd, args: [group(), group()] });
      else if (cmd === 'sqrt') out.push({ cmd, args: [group()] });
      else if (cmd === 'text' || cmd === 'mathrm') out.push({ cmd: 'text', args: [group()] });
      else if (cmd === 'left' || cmd === 'right') { /* \left( → ( */ }
      else out.push(SYM[cmd] !== undefined ? SYM[cmd] : `\\${cmd}`);
      return;
    }
    out.push(s[i++]);
  };
  const seq = (end) => {
    const out = [];
    while (i < s.length && s[i] !== end) {
      if (s[i] === '^') { i++; out.push({ sup: group() }); } else if (s[i] === '_') { i++; out.push({ sub: group() }); } else if (s[i] === '{') { i++; out.push(...seq('}')); i++; } else atom(out);
    }
    return out;
  };
  return seq(undefined);
}

const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
function html(tree) {
  return tree.map((x) => {
    if (typeof x === 'string') return /[a-zA-Z]/.test(x) && x.length === 1 ? `<i>${esc(x)}</i>` : esc(x);
    if (x.sup) return `<sup>${html(x.sup)}</sup>`;
    if (x.sub) return `<sub>${html(x.sub)}</sub>`;
    if (x.cmd === 'frac') return `<span class="of-eq-frac"><span>${html(x.args[0])}</span><span>${html(x.args[1])}</span></span>`;
    if (x.cmd === 'sqrt') return `<span class="of-eq-sqrt">√<span>${html(x.args[0])}</span></span>`;
    if (x.cmd === 'text') return esc(linear(x.args[0]));
    return '';
  }).join('');
}
function linear(tree) {
  const wrap = (t) => { const l = linear(t); return l.length > 1 && !/^\(.*\)$/.test(l) ? `(${l})` : l; };
  return tree.map((x) => {
    if (typeof x === 'string') return x;
    if (x.sup) { const l = linear(x.sup); return [...l].every((c) => SUP[c]) ? [...l].map((c) => SUP[c]).join('') : `^${wrap(x.sup)}`; }
    if (x.sub) { const l = linear(x.sub); return [...l].every((c) => SUB[c]) ? [...l].map((c) => SUB[c]).join('') : `_${wrap(x.sub)}`; }
    if (x.cmd === 'frac') return `${wrap(x.args[0])}/${wrap(x.args[1])}`;
    if (x.cmd === 'sqrt') return `√${wrap(x.args[0])}`;
    if (x.cmd === 'text') return linear(x.args[0]);
    return '';
  }).join('');
}

function omml(tree) {
  const run = (t) => (t ? `<m:r><m:t xml:space="preserve">${esc(t)}</m:t></m:r>` : '');
  let s = ''; let buf = '';
  const flush = () => { s += run(buf); buf = ''; };
  for (let k = 0; k < tree.length; k++) {
    const x = tree[k];
    if (typeof x === 'string') { buf += x; continue; }
    if (x.sup || x.sub) {
      // the base is the last character collected
      const base = buf.slice(-1); buf = buf.slice(0, -1); flush();
      const nx = tree[k + 1];
      if (x.sub && nx && nx.sup) { s += `<m:sSubSup><m:e>${run(base)}</m:e><m:sub>${omml(x.sub)}</m:sub><m:sup>${omml(nx.sup)}</m:sup></m:sSubSup>`; k++; continue; }
      s += x.sup ? `<m:sSup><m:e>${run(base)}</m:e><m:sup>${omml(x.sup)}</m:sup></m:sSup>` : `<m:sSub><m:e>${run(base)}</m:e><m:sub>${omml(x.sub)}</m:sub></m:sSub>`;
      continue;
    }
    flush();
    if (x.cmd === 'frac') s += `<m:f><m:num>${omml(x.args[0])}</m:num><m:den>${omml(x.args[1])}</m:den></m:f>`;
    else if (x.cmd === 'sqrt') s += `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${omml(x.args[0])}</m:e></m:rad>`;
    else if (x.cmd === 'text') s += run(linear(x.args[0]));
  }
  flush();
  return s;
}

export const equationHtml = (latex) => html(parse(latex));
/** Office Math (OMML) for docx */
export const equationOmml = (latex) => `<m:oMath>${omml(parse(latex))}</m:oMath>`;
export const equationText = (latex) => linear(parse(latex));
