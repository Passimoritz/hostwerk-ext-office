/*
 * Compare two versions of a document (diff-match-patch): blocks are matched first (like lines), changed text blocks
 * are diffed word by word. `applyAsTrackedChanges(old, new)` turns the result into ONE document with real insertion /
 * deletion marks (and inserted/deleted paragraph marks) – saved as w:ins / w:del, accept/reject like any change.
 * Formatting of the new version is kept for unchanged and inserted text, that of the old one for deleted text.
 */
import DMP from 'diff-match-patch';
import { inlineText, TEXTBLOCKS, docText } from './pm.js';

const DiffMatchPatch = DMP.diff_match_patch || DMP;
const OBJ = '￼';
const clone = (x) => JSON.parse(JSON.stringify(x));

/** the comparable text of a top-level block */
function blockText(n) {
  if (TEXTBLOCKS.has(n.type)) return `${n.type}:${n.type === 'heading' ? (n.attrs && n.attrs.level) || 1 : ''}:${inlineFlat(n.content)}`;
  return `${n.type}:${docText({ type: 'doc', content: [n] })}`;
}
/** inline content as one string (objects = one placeholder char) */
function inlineFlat(content) { return (content || []).map((c) => (c.type === 'text' ? c.text : OBJ)).join(''); }

/** diff of two token lists: each distinct token becomes one char (the "lines to chars" trick) */
function diffTokens(ta, tb, semantic = true) {
  const map = new Map(); const list = [];
  const enc = (arr) => arr.map((t) => { if (!map.has(t)) { map.set(t, list.length); list.push(t); } return String.fromCharCode(0x100 + map.get(t)); }).join('');
  const dmp = new DiffMatchPatch();
  dmp.Diff_Timeout = 2;
  const d = dmp.diff_main(enc(ta), enc(tb), false);
  if (semantic) dmp.diff_cleanupSemantic(d);
  return d.map(([op, s]) => ({ op, tokens: [...s].map((ch) => list[ch.charCodeAt(0) - 0x100]) }));
}
const words = (s) => s.match(/\s+|[\p{L}\p{N}_'’-]+|[^\s\p{L}\p{N}_'’-]/gu) || [];

/** → { blocks: [{ kind: same|added|removed|changed, a, b, parts? }], stats } */
export function compareDocs(docA, docB) {
  const A = (docA && docA.content) || []; const B = (docB && docB.content) || [];
  const d = diffTokens(A.map(blockText), B.map(blockText), false);
  const blocks = [];
  let ia = 0; let ib = 0;
  for (let k = 0; k < d.length; k++) {
    const { op, tokens } = d[k];
    if (op === 0) { for (let j = 0; j < tokens.length; j++) blocks.push({ kind: 'same', a: A[ia++], b: B[ib++] }); continue; }
    if (op === -1 && d[k + 1] && d[k + 1].op === 1) {
      // a run of removed blocks followed by added ones: pair text blocks of the same type as "changed"
      const rem = tokens.map(() => A[ia++]); const add = d[k + 1].tokens.map(() => B[ib++]);
      k++;
      const n = Math.max(rem.length, add.length);
      for (let j = 0; j < n; j++) {
        const x = rem[j]; const y = add[j];
        if (x && y && TEXTBLOCKS.has(x.type) && TEXTBLOCKS.has(y.type)) {
          const parts = diffTokens(words(inlineFlat(x.content)), words(inlineFlat(y.content))).map((p) => ({ op: p.op, text: p.tokens.join('') }));
          blocks.push({ kind: 'changed', a: x, b: y, parts });
        } else { if (x) blocks.push({ kind: 'removed', a: x }); if (y) blocks.push({ kind: 'added', b: y }); }
      }
      continue;
    }
    for (let j = 0; j < tokens.length; j++) { if (op === -1) blocks.push({ kind: 'removed', a: A[ia++] }); else blocks.push({ kind: 'added', b: B[ib++] }); }
  }
  const stats = { same: 0, unchanged: 0, changed: 0, added: 0, removed: 0, insertedWords: 0, deletedWords: 0 };
  for (const x of blocks) {
    if (x.kind === 'same') { stats.same++; stats.unchanged++; } else stats[x.kind]++;
    if (x.kind === 'changed') for (const p of x.parts) { const w = words(p.text).filter((t) => /\S/.test(t)).length; if (p.op === 1) stats.insertedWords += w; if (p.op === -1) stats.deletedWords += w; }
    if (x.kind === 'added') stats.insertedWords += words(docText({ type: 'doc', content: [x.b] })).filter((t) => /\S/.test(t)).length;
    if (x.kind === 'removed') stats.deletedWords += words(docText({ type: 'doc', content: [x.a] })).filter((t) => /\S/.test(t)).length;
  }
  return { blocks, stats };
}

/** inline nodes of [start, end) of a flat inline string (text sliced, objects whole) */
function sliceInline(content, start, end) {
  const out = []; let pos = 0;
  for (const c of content || []) {
    const len = c.type === 'text' ? c.text.length : 1;
    const s = Math.max(start, pos); const e = Math.min(end, pos + len);
    if (s < e) out.push(c.type === 'text' ? { ...clone(c), text: c.text.slice(s - pos, e - pos) } : clone(c));
    pos += len;
    if (pos >= end) break;
  }
  return out;
}
const REV = new Set(['insertion', 'deletion', 'formatChange']);
function withMark(nodes, mark, textOnly = false) {
  return nodes.filter((n) => !textOnly || n.type === 'text').map((n) => ({ ...n, marks: [...(n.marks || []).filter((m) => !REV.has(m.type)), mark] }));
}
function markAll(node, mark, pmark) {
  const n = clone(node);
  const visit = (x) => {
    if (TEXTBLOCKS.has(x.type)) { x.content = withMark(x.content || [], mark, mark.type === 'deletion'); x.attrs = { ...(x.attrs || {}), pmark }; return; }
    for (const c of x.content || []) visit(c);
  };
  visit(n);
  return n;
}

/** the new version with the differences to the old one as tracked changes */
export function applyAsTrackedChanges(docA, docB, { author = '', date = new Date().toISOString() } = {}) {
  const { blocks } = compareDocs(docA, docB);
  let rid = 0;
  const rev = (type) => ({ type, attrs: { id: `cmp${++rid}`, author, date } });
  const content = [];
  for (const x of blocks) {
    if (x.kind === 'same') content.push(clone(x.b));
    else if (x.kind === 'added') content.push(markAll(x.b, rev('insertion'), { type: 'ins', id: `cmp${++rid}`, author, date }));
    else if (x.kind === 'removed') content.push(markAll(x.a, rev('deletion'), { type: 'del', id: `cmp${++rid}`, author, date }));
    else {
      const node = clone(x.b);
      const inl = []; let pa = 0; let pb = 0;
      for (const p of x.parts) {
        const len = p.text.length;
        if (p.op === 0) { inl.push(...sliceInline(x.b.content, pb, pb + len)); pa += len; pb += len; }
        else if (p.op === 1) { inl.push(...withMark(sliceInline(x.b.content, pb, pb + len), rev('insertion'))); pb += len; }
        else { inl.push(...withMark(sliceInline(x.a.content, pa, pa + len), rev('deletion'), true)); pa += len; }
      }
      if (inl.length) node.content = inl; else delete node.content;
      content.push(node);
    }
  }
  return { type: 'doc', content: content.length ? content : [{ type: 'paragraph' }] };
}
