/*
 * Compare two PDFs – pure parts, tested in tests/pdf-compare.test.js.
 *
 *   diffWords(a, b)              → [{ op: -1 | 0 | 1, text }]  word-level diff (diff-match-patch on word tokens)
 *   comparePages(pagesA, pagesB) → [{ a, b, ops, added, removed, same }]  pages matched in order (by similarity)
 *   pixelDiff(A, B, w, h)        → { count, ratio, mask }  RGBA arrays of the same size; mask: 1 = differs
 */
import DiffMatchPatch from 'diff-match-patch';

const dmp = new DiffMatchPatch();
dmp.Diff_Timeout = 2;

/** words + whitespace as tokens → one char each (like diff_linesToChars) */
function tokensToChars(a, b) {
  const map = new Map(); const list = [];
  const enc = (s) => {
    let out = '';
    for (const tok of s.match(/\s+|[^\s]+/g) || []) {
      let id = map.get(tok);
      if (id === undefined) { id = list.length; list.push(tok); map.set(tok, id); }
      out += String.fromCharCode(id + 32 < 0xd800 ? id + 32 : id + 0x2000 + 32);
    }
    return out;
  };
  return { ca: enc(a), cb: enc(b), list };
}
const back = (chars, list) => Array.from(chars, (c) => { const n = c.charCodeAt(0); return list[n >= 0xd800 ? n - 0x2000 - 32 : n - 32]; }).join('');

export function diffWords(a, b) {
  const { ca, cb, list } = tokensToChars(String(a || ''), String(b || ''));
  const d = dmp.diff_main(ca, cb, false);
  dmp.diff_cleanupSemantic(d);
  const out = [];
  for (const [op, chars] of d) {
    const text = back(chars, list);
    const last = out[out.length - 1];
    if (last && last.op === op) last.text += text; else out.push({ op, text });
  }
  return out;
}

const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim();
function similarity(a, b) {
  const x = norm(a); const y = norm(b);
  if (!x && !y) return 1;
  if (!x || !y) return 0;
  const d = dmp.diff_main(x, y); const lev = dmp.diff_levenshtein(d);
  return 1 - lev / Math.max(x.length, y.length);
}

/**
 * Pages of two versions, matched in order: a page that only exists on one side (inserted/deleted page) is detected
 * by looking one page ahead. pagesA/pagesB: [text]
 */
export function comparePages(pagesA, pagesB) {
  const out = []; let i = 0; let j = 0;
  while (i < pagesA.length || j < pagesB.length) {
    if (i >= pagesA.length) { out.push({ a: null, b: j, ops: [{ op: 1, text: pagesB[j] }], added: norm(pagesB[j]).length, removed: 0, same: false }); j++; continue; }
    if (j >= pagesB.length) { out.push({ a: i, b: null, ops: [{ op: -1, text: pagesA[i] }], added: 0, removed: norm(pagesA[i]).length, same: false }); i++; continue; }
    const s = similarity(pagesA[i], pagesB[j]);
    if (s < 0.5) {
      const skipB = j + 1 < pagesB.length ? similarity(pagesA[i], pagesB[j + 1]) : 0;
      const skipA = i + 1 < pagesA.length ? similarity(pagesA[i + 1], pagesB[j]) : 0;
      if (skipB > s + 0.25 && skipB >= skipA) { out.push({ a: null, b: j, ops: [{ op: 1, text: pagesB[j] }], added: norm(pagesB[j]).length, removed: 0, same: false }); j++; continue; }
      if (skipA > s + 0.25) { out.push({ a: i, b: null, ops: [{ op: -1, text: pagesA[i] }], added: 0, removed: norm(pagesA[i]).length, same: false }); i++; continue; }
    }
    const ops = diffWords(pagesA[i], pagesB[j]);
    let added = 0; let removed = 0;
    for (const o of ops) { if (o.op > 0) added += norm(o.text).length; else if (o.op < 0) removed += norm(o.text).length; }
    out.push({ a: i, b: j, ops, added, removed, same: !added && !removed });
    i++; j++;
  }
  return out;
}

/** RGBA pixels → which differ noticeably (anti-aliasing tolerance) */
export function pixelDiff(A, B, w, h, tol = 48) {
  const mask = new Uint8Array(w * h); let count = 0;
  for (let p = 0, q = 0; p < w * h; p++, q += 4) {
    const d = Math.abs(A[q] - B[q]) + Math.abs(A[q + 1] - B[q + 1]) + Math.abs(A[q + 2] - B[q + 2]);
    if (d > tol) { mask[p] = 1; count++; }
  }
  return { count, ratio: count / Math.max(1, w * h), mask };
}
