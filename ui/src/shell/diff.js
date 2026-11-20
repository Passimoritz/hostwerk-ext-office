/*
 * Line diff (LCS) for the compare view's text fallback (chunks without their own Compare) – pure, unit-tested.
 *   lineDiff(a, b) → [{ op: '=' | '+' | '-', text }]
 */
export function lineDiff(a, b, max = 4000) {
  const A = String(a || '').split('\n').slice(0, max);
  const B = String(b || '').split('\n').slice(0, max);
  // common head and tail first – most edits touch a small part, the table stays small
  let s = 0;
  while (s < A.length && s < B.length && A[s] === B[s]) s++;
  let e = 0;
  while (e < A.length - s && e < B.length - s && A[A.length - 1 - e] === B[B.length - 1 - e]) e++;
  const head = A.slice(0, s).map((text) => ({ op: '=', text }));
  const tail = A.slice(A.length - e).map((text) => ({ op: '=', text }));
  const a2 = A.slice(s, A.length - e); const b2 = B.slice(s, B.length - e);
  const n = a2.length; const m = b2.length;
  if (n * m > 4e6) return [...head, ...a2.map((text) => ({ op: '-', text })), ...b2.map((text) => ({ op: '+', text })), ...tail];
  const L = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = a2[i] === b2[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
  const mid = []; let i = 0; let j = 0;
  while (i < n && j < m) {
    if (a2[i] === b2[j]) { mid.push({ op: '=', text: a2[i] }); i++; j++; } else if (L[i + 1][j] >= L[i][j + 1]) mid.push({ op: '-', text: a2[i++] }); else mid.push({ op: '+', text: b2[j++] });
  }
  while (i < n) mid.push({ op: '-', text: a2[i++] });
  while (j < m) mid.push({ op: '+', text: b2[j++] });
  return [...head, ...mid, ...tail];
}
