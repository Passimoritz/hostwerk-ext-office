/*
 * Lazy sub-chunk of the docs editor (not an editor kind – like pdfworker, only a build output ui/assets/docs-math.js):
 * temml (LaTeX → MathML, ~210 KB) and the version compare (diff-match-patch) load the first time a document shows or
 * edits a formula / compares two versions.
 */
import temml from 'temml';
import { compareDocs, applyAsTrackedChanges } from '../docs/io/compare.js';

export function latexToMathml(latex, { display = false } = {}) {
  return temml.renderToString(String(latex || ''), { xml: true, annotate: true, displayMode: !!display, throwOnError: true, trust: false });
}
export function tryLatexToMathml(latex, opts) {
  try { return { mathml: latexToMathml(latex, opts), error: null }; } catch (e) { return { mathml: null, error: String((e && e.message) || e).replace(/^ParseError:\s*/, '') }; }
}
export default { latexToMathml, tryLatexToMathml, compare: { compareDocs, applyAsTrackedChanges } };
