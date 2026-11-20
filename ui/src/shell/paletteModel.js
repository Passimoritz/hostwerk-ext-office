/*
 * Ranking of the command palette (pure, unit-tested in tests/shell-palette.test.js).
 *   score('bold', 'Bold text') → number (0 = no match): word starts beat substrings beat scattered letters;
 *   umlauts and case are folded, so "einfg" finds "Einfügen" and "tab" finds "Tabelle einfügen".
 *   rank(commands, query, recent) → the commands to show, best first (recent ones first when the query is empty)
 */
export const fold = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/ß/g, 'ss');

export function score(query, text) {
  const q = fold(query).trim();
  if (!q) return 1;
  const s = fold(text);
  if (!s) return 0;
  if (s === q) return 1000;
  if (s.startsWith(q)) return 800 - s.length;
  const words = s.split(/[\s\-–/·.,:()]+/).filter(Boolean);
  if (words.some((w) => w.startsWith(q))) return 600 - s.length;
  // every word of the query starts a word of the text ("ins tab" → "Insert table")
  const qw = q.split(/\s+/).filter(Boolean);
  if (qw.length > 1 && qw.every((x) => words.some((w) => w.startsWith(x)))) return 500 - s.length;
  const at = s.indexOf(q);
  if (at >= 0) return 400 - at - s.length / 10;
  // initials ("ic" → "Insert comment")
  const ini = words.map((w) => w[0]).join('');
  if (q.length >= 2 && ini.startsWith(q)) return 350 - s.length / 10;
  // scattered letters in order, a gap costs
  let i = 0; let gaps = 0; let last = -1;
  for (let j = 0; j < s.length && i < q.length; j++) {
    if (s[j] === q[i]) { if (last >= 0 && j - last > 1) gaps += j - last - 1; last = j; i++; }
  }
  if (i < q.length || gaps > q.length * 4) return 0;
  return Math.max(1, 200 - gaps * 8 - s.length / 10);
}

/** commands: [{ id, title, group?, keywords?, when?() }] → sorted matches (title translated by the caller) */
export function rank(commands, query, recent = [], max = 60) {
  const rec = new Map(recent.map((id, i) => [id, recent.length - i]));
  const out = [];
  for (const c of commands) {
    if (!c || !c.title) continue;
    try { if (typeof c.when === 'function' && !c.when()) continue; } catch { continue; }
    const s = Math.max(score(query, c.title), (c.keywords || c.words) ? score(query, String(c.keywords || c.words)) * 0.8 : 0, c.group ? score(query, `${c.group} ${c.title}`) * 0.7 : 0);
    if (s <= 0) continue;
    out.push({ c, s: s + (rec.get(c.id) || 0) * (query ? 6 : 100) });
  }
  out.sort((a, b) => b.s - a.s || String(a.c.title).localeCompare(String(b.c.title)));
  return out.slice(0, max).map((x) => x.c);
}

/** remember a used command (most recent first, at most 12) */
export function pushRecent(recent, id, max = 12) { return [id, ...(recent || []).filter((x) => x !== id)].slice(0, max); }
