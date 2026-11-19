/*
 * Full-text search over pdf.js text items. Pure (testable): the text of a page is the items joined, with a space
 * where pdf.js reports an end of line; a hit is mapped back to the items it covers.
 */
export function pageIndex(items) {
  const strs = []; const offs = []; let text = '';
  for (const it of items) {
    if (it.str === undefined) continue;
    offs.push(text.length); strs.push(it.str);
    text += it.str;
    if (it.hasEOL && !/\s$/.test(it.str)) text += ' ';
  }
  return { text, strs, offs };
}

const fold = (s) => s.normalize('NFKD').replace(/[̀-ͯ]/g, '');

/** hits in one page index: [{ start, end, parts: [{ i, a, b }] }] */
export function findInPage(idx, query, { matchCase = false, wholeWord = false } = {}) {
  if (!query) return [];
  // folding keeps the length for most Latin text; when it would not, search the raw text
  let hay = idx.text; let needle = query;
  const f = fold(hay);
  if (!matchCase) { hay = (f.length === hay.length ? f : hay).toLowerCase(); needle = fold(needle).toLowerCase(); }
  const out = [];
  let from = 0;
  while (needle && from <= hay.length) {
    const at = hay.indexOf(needle, from);
    if (at < 0) break;
    const end = at + needle.length;
    from = at + Math.max(1, needle.length);
    if (wholeWord && ((at > 0 && /\w/.test(hay[at - 1])) || (end < hay.length && /\w/.test(hay[end])))) continue;
    const parts = [];
    for (let i = 0; i < idx.strs.length; i++) {
      const s = idx.offs[i]; const e = s + idx.strs[i].length;
      if (e <= at || s >= end) continue;
      parts.push({ i, a: Math.max(0, at - s), b: Math.min(idx.strs[i].length, end - s) });
    }
    out.push({ start: at, end, parts });
  }
  return out;
}

/** short text around a hit for the result list */
export function snippet(idx, hit, around = 30) {
  const a = Math.max(0, hit.start - around); const b = Math.min(idx.text.length, hit.end + around);
  return `${a > 0 ? '…' : ''}${idx.text.slice(a, b).trim()}${b < idx.text.length ? '…' : ''}`;
}
