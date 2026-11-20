/*
 * MathML (as temml writes it from LaTeX) → Office Math (OMML) for docx, and MathML → LaTeX for MathML that comes
 * from other programs (ODT formulas). Pure string/tree code – runs in Node (unit tests) and in the browser.
 *
 * Covered: mi/mn/mo/mtext/ms/mspace, mrow/mstyle/mpadded/semantics, mfrac (incl. no bar), msqrt/mroot, msub/msup/
 * msubsup, munder/mover/munderover (accents, limits), large operators (∑ ∫ ∏ … → m:nary), fenced rows (\left( … \right)
 * → m:d), mtable (→ m:m, inside fences as a matrix), menclose (→ m:borderBox / m:bar), mphantom (dropped).
 */
import { parseXml, kids, local, esc } from '../../_shared/xmltree.js';

const NARY = new Set(['∑', '∏', '∐', '∫', '∬', '∭', '∮', '∯', '∰', '⋃', '⋂', '⋁', '⋀', '⨁', '⨂', '⨀']);
const els = (el) => kids(el, '*');
const txt = (el) => (el.c || []).map((c) => (typeof c === 'string' ? c : txt(c))).join('');
const nameOf = (el) => (el && typeof el === 'object' ? local(el.n) : '');

/** the <math> element of a MathML string (or null) */
export function mathRoot(mathml) {
  if (!mathml) return null;
  const doc = typeof mathml === 'string' ? parseXml(mathml) : mathml;
  const find = (el) => { if (nameOf(el) === 'math') return el; for (const c of els(el)) { const f = find(c); if (f) return f; } return null; };
  return find(doc);
}

/** the LaTeX annotation temml (or others) stored inside the MathML, if any */
export function texAnnotation(mathml) {
  const root = mathRoot(mathml);
  if (!root) return null;
  let found = null;
  const visit = (el) => { if (found) return; if (nameOf(el) === 'annotation' && /tex/i.test(el.a.encoding || '')) { found = txt(el); return; } els(el).forEach(visit); };
  visit(root);
  return found;
}

/* ---------------- MathML → OMML ---------------- */
const run = (text, sty) => {
  const s = String(text || '').replace(/⁡|⁢|⁣/g, '');
  if (!s) return '';
  return `<m:r>${sty ? `<m:rPr>${sty}</m:rPr>` : ''}<m:t xml:space="preserve">${esc(s)}</m:t></m:r>`;
};
const VARIANT = { normal: '<m:sty m:val="p"/>', bold: '<m:sty m:val="b"/>', 'bold-italic': '<m:sty m:val="bi"/>', italic: '<m:sty m:val="i"/>', 'double-struck': '<m:scr m:val="double-struck"/><m:sty m:val="p"/>', script: '<m:scr m:val="script"/>', fraktur: '<m:scr m:val="fraktur"/>', monospace: '<m:scr m:val="monospace"/>', 'sans-serif': '<m:scr m:val="sans-serif"/>' };
const ACCENT_CHARS = new Set(['ˆ', '^', '̂', '¯', '‾', '̄', '→', '⃗', '˜', '~', '̃', '˙', '̇', '¨', '̈', 'ˇ', '̌', '´', '̀', '`', '˘', '̆', '←', '↔', '⏞', '⏟']);

function omml(el) {
  if (typeof el === 'string') return '';
  const n = nameOf(el);
  const c = els(el);
  const seq = (list) => list.map(omml).join('');
  switch (n) {
    case 'math': case 'semantics': return seq(c.filter((x) => !/^annotation/.test(nameOf(x))));
    case 'annotation': case 'annotation-xml': case 'mphantom': case 'none': case 'mprescripts': return '';
    case 'mrow': case 'mstyle': case 'mpadded': case 'merror': case 'maction': return row(c);
    case 'mi': { const v = el.a.mathvariant; const s = txt(el); return run(s, v ? VARIANT[v] || '' : s.length > 1 ? VARIANT.normal : ''); }
    case 'mn': case 'mo': return run(txt(el), n === 'mn' ? '' : '');
    case 'mtext': case 'ms': return run(txt(el), '<m:sty m:val="p"/>');
    case 'mspace': { const w = parseFloat(el.a.width || '0'); return w > 0.05 ? run(w >= 0.9 ? ' ' : w >= 0.25 ? ' ' : ' ') : ''; }
    case 'mfrac': return `<m:f>${/^0/.test(el.a.linethickness || '') ? '<m:fPr><m:type m:val="noBar"/></m:fPr>' : ''}<m:num>${omml(c[0] || '')}</m:num><m:den>${omml(c[1] || '')}</m:den></m:f>`;
    case 'msqrt': return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${row(c)}</m:e></m:rad>`;
    case 'mroot': return `<m:rad><m:deg>${omml(c[1] || '')}</m:deg><m:e>${omml(c[0] || '')}</m:e></m:rad>`;
    case 'msup': return isNary(c[0]) ? nary(c[0], null, c[1], '') : `<m:sSup><m:e>${omml(c[0] || '')}</m:e><m:sup>${omml(c[1] || '')}</m:sup></m:sSup>`;
    case 'msub': return isNary(c[0]) ? nary(c[0], c[1], null, '') : `<m:sSub><m:e>${omml(c[0] || '')}</m:e><m:sub>${omml(c[1] || '')}</m:sub></m:sSub>`;
    case 'msubsup': return isNary(c[0]) ? nary(c[0], c[1], c[2], '') : `<m:sSubSup><m:e>${omml(c[0] || '')}</m:e><m:sub>${omml(c[1] || '')}</m:sub><m:sup>${omml(c[2] || '')}</m:sup></m:sSubSup>`;
    case 'munder': case 'mover': case 'munderover': {
      if (isNary(c[0])) return nary(c[0], n === 'mover' ? null : c[1], n === 'munder' ? null : n === 'mover' ? c[1] : c[2], 'undOvr');
      const acc = n === 'mover' && nameOf(c[1]) === 'mo' && ACCENT_CHARS.has(txt(c[1]).trim());
      if (acc) {
        const ch = txt(c[1]).trim();
        if (ch === '¯' || ch === '‾') return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${omml(c[0])}</m:e></m:bar>`;
        if (ch === '⏞' || ch === '⏟') return `<m:groupChr><m:groupChrPr><m:chr m:val="${esc(ch)}"/><m:pos m:val="${ch === '⏞' ? 'top' : 'bot'}"/></m:groupChrPr><m:e>${omml(c[0])}</m:e></m:groupChr>`;
        const map = { 'ˆ': '̂', '^': '̂', '→': '⃗', '˜': '̃', '~': '̃', '˙': '̇', '¨': '̈', 'ˇ': '̌', '´': '́', '`': '̀', '˘': '̆' };
        return `<m:acc><m:accPr><m:chr m:val="${esc(map[ch] || ch)}"/></m:accPr><m:e>${omml(c[0])}</m:e></m:acc>`;
      }
      if (n === 'munder') return `<m:limLow><m:e>${omml(c[0] || '')}</m:e><m:lim>${omml(c[1] || '')}</m:lim></m:limLow>`;
      if (n === 'mover') return `<m:limUpp><m:e>${omml(c[0] || '')}</m:e><m:lim>${omml(c[1] || '')}</m:lim></m:limUpp>`;
      return `<m:limUpp><m:e><m:limLow><m:e>${omml(c[0] || '')}</m:e><m:lim>${omml(c[1] || '')}</m:lim></m:limLow></m:e><m:lim>${omml(c[2] || '')}</m:lim></m:limUpp>`;
    }
    case 'mtable': return matrix(el);
    case 'mtr': case 'mlabeledtr': return els(el).map((td) => omml(td)).join('');
    case 'mtd': return row(c);
    case 'menclose': {
      const no = el.a.notation || 'box';
      if (/top/.test(no) && !/box/.test(no)) return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${row(c)}</m:e></m:bar>`;
      if (/bottom/.test(no) && !/box/.test(no)) return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${row(c)}</m:e></m:bar>`;
      return `<m:borderBox><m:e>${row(c)}</m:e></m:borderBox>`;
    }
    case 'mfenced': return delim(el.a.open != null ? el.a.open : '(', el.a.close != null ? el.a.close : ')', c);
    default: return row(c);
  }
}
function isNary(el) { return nameOf(el) === 'mo' && NARY.has(txt(el).trim()); }
function nary(op, sub, sup, loc) {
  const ch = txt(op).trim();
  const pr = `<m:naryPr><m:chr m:val="${esc(ch)}"/>${loc ? `<m:limLoc m:val="${loc}"/>` : ch === '∑' || ch === '∏' ? '<m:limLoc m:val="undOvr"/>' : '<m:limLoc m:val="subSup"/>'}${sub ? '' : '<m:subHide m:val="1"/>'}${sup ? '' : '<m:supHide m:val="1"/>'}</m:naryPr>`;
  return `<m:nary>${pr}<m:sub>${sub ? omml(sub) : ''}</m:sub><m:sup>${sup ? omml(sup) : ''}</m:sup><m:e></m:e></m:nary>`;
}
const isFence = (el, form) => nameOf(el) === 'mo' && (el.a.fence === 'true' || el.a.stretchy === 'true') && (!form || el.a.form === form || (!el.a.form && (form === 'prefix' ? /^[([{⟨|‖⌊⌈]$/.test(txt(el).trim()) : /^[)\]}⟩|‖⌋⌉]$/.test(txt(el).trim()))));
function row(list) {
  // \left( … \right) → one delimiter object
  if (list.length >= 2 && isFence(list[0], 'prefix') && isFence(list[list.length - 1], 'postfix')) return delim(txt(list[0]).trim(), txt(list[list.length - 1]).trim(), list.slice(1, -1));
  return list.map(omml).join('');
}
function delim(open, close, inner) {
  const pr = `<m:dPr>${open !== '(' ? `<m:begChr m:val="${esc(open)}"/>` : ''}${close !== ')' ? `<m:endChr m:val="${esc(close)}"/>` : ''}</m:dPr>`;
  return `<m:d>${pr}<m:e>${inner.map(omml).join('')}</m:e></m:d>`;
}
function matrix(el) {
  const rows = els(el).filter((r) => /^m(labeled)?tr$/.test(nameOf(r)));
  const cols = Math.max(1, ...rows.map((r) => els(r).filter((x) => nameOf(x) === 'mtd').length));
  return `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${cols}"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>${rows.map((r) => `<m:mr>${els(r).filter((x) => nameOf(x) === 'mtd').map((td) => `<m:e>${omml(td)}</m:e>`).join('')}</m:mr>`).join('')}</m:m>`;
}

/** MathML string → `<m:oMath>…</m:oMath>` (display = wrapped in `<m:oMathPara>`) */
export function mathmlToOmml(mathml, { display = false } = {}) {
  const root = mathRoot(mathml);
  const inner = root ? omml(root) : '';
  const m = `<m:oMath>${inner}</m:oMath>`;
  return display ? `<m:oMathPara>${m}</m:oMathPara>` : m;
}

/* ---------------- MathML → LaTeX (for MathML without a TeX annotation) ---------------- */
const TEX = { '∑': '\\sum', '∏': '\\prod', '∫': '\\int', '∮': '\\oint', '∞': '\\infty', '≤': '\\leq', '≥': '\\geq', '≠': '\\neq', '±': '\\pm', '·': '\\cdot', '×': '\\times', '→': '\\to', '≈': '\\approx', 'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta', 'π': '\\pi', 'λ': '\\lambda', 'μ': '\\mu', 'σ': '\\sigma', 'θ': '\\theta', 'ω': '\\omega', 'Δ': '\\Delta', 'Σ': '\\Sigma', 'Ω': '\\Omega', '∈': '\\in', '∂': '\\partial', '∇': '\\nabla' };
function tex(el) {
  if (typeof el === 'string') return '';
  const n = nameOf(el); const c = els(el);
  const g = (x) => { const s = tex(x); return s.length === 1 ? s : `{${s}}`; };
  const sym = (s) => [...s].map((ch) => (TEX[ch] ? `${TEX[ch]} ` : ch)).join('').trim();
  switch (n) {
    case 'annotation': case 'annotation-xml': return '';
    case 'mi': { const s = txt(el); return s.length > 1 ? `\\mathrm{${s}}` : sym(s); }
    case 'mn': return txt(el);
    case 'mo': return sym(txt(el).trim());
    case 'mtext': return `\\text{${txt(el)}}`;
    case 'mfrac': return `\\frac{${tex(c[0] || '')}}{${tex(c[1] || '')}}`;
    case 'msqrt': return `\\sqrt{${c.map(tex).join('')}}`;
    case 'mroot': return `\\sqrt[${tex(c[1] || '')}]{${tex(c[0] || '')}}`;
    case 'msup': return `${tex(c[0] || '')}^${g(c[1] || '')}`;
    case 'msub': return `${tex(c[0] || '')}_${g(c[1] || '')}`;
    case 'msubsup': case 'munderover': return `${tex(c[0] || '')}_${g(c[1] || '')}^${g(c[2] || '')}`;
    case 'munder': return `${tex(c[0] || '')}_${g(c[1] || '')}`;
    case 'mover': return `${tex(c[0] || '')}^${g(c[1] || '')}`;
    case 'mtable': return `\\begin{matrix}${c.map((r) => els(r).map(tex).join(' & ')).join(' \\\\ ')}\\end{matrix}`;
    case 'mspace': return '\\,';
    default: return c.map(tex).join(' ').replace(/\s+/g, ' ');
  }
}
/** MathML → LaTeX: the TeX annotation when there is one, otherwise a structural conversion */
export function mathmlToLatex(mathml) {
  const a = texAnnotation(mathml);
  if (a) return a.trim();
  const root = mathRoot(mathml);
  return root ? tex(root).replace(/\s+/g, ' ').trim() : '';
}
/** strip the namespace prefixes ODF uses (math:mi → mi) and give the root the MathML namespace */
export function normalizeMathml(xml) {
  return String(xml || '').replace(/<(\/?)math:/g, '<$1').replace(/\sxmlns:math="[^"]*"/g, '').replace(/<math(?![^>]*xmlns=)/, '<math xmlns="http://www.w3.org/1998/Math/MathML"');
}
