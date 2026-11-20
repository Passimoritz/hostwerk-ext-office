/*
 * The page model of the print layout – pure numbers, no DOM (unit-tested in tests/docs-pagination.test.js).
 *
 * Coordinates are "flow" pixels: y = 0 is the top of the text area of the first page, the pages follow each other with
 * a gap. paginate.js measures the natural layout (without any spacers), asks these planners where content has to move
 * and turns the answers into spacer widgets.
 *
 *   pager(sections, gap)          page list: C(p) text top, E(p) text bottom, edge(p) paper bottom, top(p) paper top
 *   planLines(lines, …)           a paragraph/list broken between lines
 *   planTable(rows, …)            a table broken between rows – or inside a row between its lines (Word: "allow row
 *                                 to break across pages"), header rows repeated, cantSplit rows moved whole, rows
 *                                 taller than a page always broken, never anything inside the margins or the gap
 */

export function pager(sections, gap, keep = []) {
  const pages = keep.map((p) => ({ ...p }));
  const sec = (i) => sections[Math.min(i, sections.length - 1)];
  const mt0 = sections[0].mt;
  const P = {
    pages,
    gap,
    sec,
    add(section) { const prev = pages[pages.length - 1]; const g = sec(section); const top = prev ? prev.top + prev.h + gap : 0; pages.push({ top, section, h: g.pageH, w: g.pageW, mt: g.mt, mb: g.mb, ml: g.ml, mr: g.mr }); return pages.length - 1; },
    /** text area top / bottom of page p */
    C: (p) => pages[p].top + pages[p].mt - mt0,
    E: (p) => pages[p].top + pages[p].mt - mt0 + sec(pages[p].section).contentH,
    /** paper bottom edge / paper top edge */
    edge: (p) => pages[p].top + pages[p].h - mt0,
    top: (p) => pages[p].top - mt0,
    height: (p) => sec(pages[p].section).contentH,
  };
  P.next = (p, section) => (pages[p + 1] ? p + 1 : P.add(section != null ? section : pages[p].section));
  return P;
}

const r1 = (v) => Math.round(v * 10) / 10;

/**
 * a block of text lines (natural tops/bottoms) that starts on page `page`, already moved by `shift` → spacers
 * { whole, spacers: [{ line, height }], added, page }. whole = not even the first line fits (move the block).
 */
export function planLines(lines, { shift, page, P }) {
  const spacers = []; let added = 0;
  for (let j = 0; j < lines.length; j++) {
    const lt = lines[j].top + shift + added; const lb = lines[j].bottom + shift + added;
    if (lb <= P.E(page) + 0.5) continue;
    if (j === 0 && lt > P.C(page) + 2) return { whole: true, spacers: [], added: 0, page };
    if (lt <= P.C(page) + 2) continue; // a single line taller than the page: let it run
    const np = P.next(page);
    const push = P.C(np) - lt;
    if (push <= 0.5) { page = np; continue; }
    spacers.push({ line: j, height: r1(push) });
    added += push; page = np;
    j--; // the same line again on the new page (it may still not fit – a line taller than a page)
    if (spacers.length > 10000) break;
  }
  return { whole: false, spacers, added, page };
}

/**
 * rows: [{ top, bottom, cantSplit, exact, cells: [{ lines: [{ top, bottom }] | () => [...], bottom }] }] in natural flow
 * coordinates (the table without any spacers), headN = leading header rows, headH = their height.
 * → { whole, page, added, items } where items are
 *     { kind: 'row', row, height, at, cut, next, C, head }            the row moves whole to the next page
 *     { kind: 'cell', row, cell, line, height, at, band? }             the cell's text continues on the next page from
 *                                                                      that line; band = { cutTop, cut, next, C, head }
 * at = where the spacer starts (flow y), cut = paper bottom edge, next = paper top of the new page, C = its text top,
 * head = height of the repeated header rows drawn there.
 */
export function planTable(rows, { shift, page, P, headN = 0, headH = 0, force = false }) {
  const items = []; let added = 0;
  const repeatOn = (p) => (headN && headH < P.height(p) / 3 ? headH : 0);
  const linesOf = (cell) => (typeof cell.lines === 'function' ? (cell.lines = cell.lines()) : cell.lines || []);
  for (let j = 0; j < rows.length; j++) {
    const row = rows[j];
    const base = shift + added;
    let rt = row.top + base; const rb = row.bottom + base;
    // the row starts where the page has ended already (the previous row ended at the page end)
    if (rt >= P.E(page) - 0.5) {
      if (j === 0) return { whole: true, page, added: 0, items: [] };
      const np = P.next(page); const rep = repeatOn(np);
      const push = P.C(np) + rep - rt;
      items.push({ kind: 'row', row: j, height: r1(push), at: rt, cut: P.edge(page), next: P.top(np), C: P.C(np), head: rep });
      added += push; page = np; rt += push;
      if (rb + push <= P.E(page) + 0.5) continue;
      j--; continue; // place it again on the new page (it may need splitting)
    }
    if (rb <= P.E(page) + 0.5) continue;
    if (j < headN) { if (rt > P.C(page) + 2 && !force) return { whole: true, page, added: 0, items: [] }; continue; }
    const np = P.next(page); const rep = repeatOn(np);
    const h = rb - rt;
    const fitsPage = h + rep <= P.height(np) + 0.5;
    const atTop = rt <= P.C(page) + repeatOn(page) + 2;
    const firstBody = j === headN;
    const moveRow = () => {
      const push = P.C(np) + rep - rt;
      items.push({ kind: 'row', row: j, height: r1(push), at: rt, cut: P.edge(page), next: P.top(np), C: P.C(np), head: rep });
      added += push; page = np;
    };
    if ((row.cantSplit || row.exact) && fitsPage && !atTop) {
      if (firstBody && !force && rt > P.C(page) + 2) return { whole: true, page, added: 0, items: [] }; // keep the header with its first row
      moveRow();
      continue;
    }
    // break the row between lines – every cell for itself, all continue at the top of the next page
    const cells = row.cells || [];
    const shiftC = cells.map(() => 0);
    const from = cells.map(() => 0);
    let pg = page; let guard = 0; let first = true; let retry = false;
    const firstItem = items.length;
    for (;;) {
      if (guard++ > 2000) break;
      const E = P.E(pg);
      const cross = cells.map((cell, c) => {
        const ls = linesOf(cell);
        for (let k = from[c]; k < ls.length; k++) if (ls[k].bottom + base + shiftC[c] > E + 0.5) return k;
        return -1;
      });
      if (cross.every((k) => k < 0)) break;
      // nothing of the row fits on this page (only the first lines cross) → the whole row moves (Word does the same)
      // and is broken on the next page if it is still too tall there
      if (first && !atTop && cross.every((k, c) => k <= 0)) {
        if (firstBody && !force && rt > P.C(page) + 2) return { whole: true, page, added: 0, items: [] };
        moveRow();
        retry = true;
        break;
      }
      first = false;
      const n = P.next(pg); const rp = repeatOn(n);
      let cutTop = -Infinity;
      cells.forEach((cell, c) => {
        const ls = linesOf(cell);
        const k = cross[c] < 0 ? ls.length : cross[c];
        const lastKept = k - 1 >= from[c] ? ls[k - 1].bottom + base + shiftC[c] : -Infinity;
        if (lastKept > cutTop) cutTop = lastKept;
      });
      let band = null;
      cells.forEach((cell, c) => {
        const k = cross[c]; if (k < 0) return;
        const ls = linesOf(cell);
        const lt = ls[k].top + base + shiftC[c];
        if (lt <= P.C(pg) + repeatOn(pg) + 2 && ls[k].bottom - ls[k].top > P.height(pg) - 4) { from[c] = k + 1; return; } // one line taller than a page
        const hh = P.C(n) + rp - lt;
        if (hh <= 0.5) { from[c] = k; return; }
        const it = { kind: 'cell', row: j, cell: c, line: k, height: r1(hh), at: lt };
        if (!band) { band = { cutTop: Number.isFinite(cutTop) ? cutTop : lt, cut: P.edge(pg), next: P.top(n), C: P.C(n), head: rp }; it.band = band; }
        items.push(it);
        shiftC[c] += hh; from[c] = k;
      });
      if (!band) { from.forEach((_, c) => { if (cross[c] >= 0) from[c] = Math.max(from[c], cross[c] + 1); }); continue; }
      pg = n;
    }
    if (retry) { j--; continue; }
    // the row's new bottom: the lowest cell after its spacers (never above the natural bottom)
    let nb = rb;
    cells.forEach((cell, c) => { if (shiftC[c]) nb = Math.max(nb, cell.bottom + base + shiftC[c]); });
    // the table grows by this much – carried by the row's first spacer (the layout subtracts it to get back the natural one)
    for (let q = firstItem; q < items.length; q++) items[q].grow = 0;
    if (items[firstItem]) items[firstItem].grow = r1(nb - rb);
    added += nb - rb;
    page = pg;
  }
  return { whole: false, page, added, items };
}
