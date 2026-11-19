/* System clipboard formats: TSV + HTML table (what Excel, LibreOffice and Google Sheets read and write).
   hostwerk-allow: tables – the HTML goes to other programs via the clipboard, Hostwerk's kit classes mean nothing there */
import { parseCSV } from './io/csv.js';
import { bookToHTML } from './io/formats.js';

export function toTSV(book, sheet, g) {
  const rows = [];
  for (let r = g.r1; r <= g.r2; r++) {
    if (sheet.hiddenRows.has(r) || sheet.filterHidden.has(r)) continue;
    const row = [];
    for (let c = g.c1; c <= g.c2; c++) {
      if (sheet.hiddenCols.has(c)) continue;
      const t = book.text(sheet, r, c);
      row.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
    }
    rows.push(row.join('\t'));
  }
  return `${rows.join('\r\n')}\r\n`;
}

export function toHTML(book, sheet, g, clipId) {
  const html = bookToHTML(book, { sheet: book.sheets.indexOf(sheet), range: g, gridlines: false, headings: false, fit: 'none' });
  return html.replace('<table>', `<table data-hw-clip="${clipId}">`);
}

export function parseTSV(text) {
  const t = String(text || '').replace(/\r?\n$/, '');
  if (!t) return [];
  return parseCSV(t, { delimiter: '\t' });
}

const rgbToHex = (s) => {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(s || '');
  if (m) { if (m[4] === '0') return null; return `#${[m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase()}`; }
  const h = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(String(s || '').trim());
  if (!h) return null;
  const x = h[1].length === 3 ? h[1].split('').map((ch) => ch + ch).join('') : h[1];
  return `#${x.toUpperCase()}`;
};

/** HTML clipboard → { clipId, rows: [[{ text, style }]] } or null when there is no table */
export function parseHTML(html) {
  if (typeof DOMParser === 'undefined' || !/<table/i.test(html || '')) return null;
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  const clipId = table.getAttribute('data-hw-clip');
  // Excel puts class styles into <style>; read the simple ones
  const classStyle = {};
  for (const st of doc.querySelectorAll('style')) {
    for (const m of st.textContent.matchAll(/\.([\w-]+)\s*\{([^}]*)\}/g)) classStyle[m[1]] = m[2];
  }
  const grid = [];
  const taken = [];
  [...table.rows].forEach((tr, ri) => {
    let ci = 0;
    grid[ri] = grid[ri] || [];
    for (const td of tr.cells) {
      while (taken[ri] && taken[ri][ci]) ci++;
      const css = `${(td.className && classStyle[td.className]) || ''};${td.getAttribute('style') || ''}`;
      const get = (prop) => { const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([^;]+)`, 'i').exec(css); return m ? m[1].trim() : null; };
      const style = {};
      const fw = get('font-weight'); if (fw && (fw === 'bold' || +fw >= 600)) style.b = true;
      if (td.querySelector('b,strong')) style.b = true;
      if ((get('font-style') || '').includes('italic') || td.querySelector('i,em')) style.i = true;
      if ((get('text-decoration') || '').includes('underline') || td.querySelector('u')) style.u = true;
      if ((get('text-decoration') || '').includes('line-through')) style.st = true;
      const col = rgbToHex(get('color')); if (col && col !== '#000000') style.color = col;
      const bg = rgbToHex(get('background-color') || get('background')); if (bg && bg !== '#FFFFFF') style.fill = bg;
      const ta = get('text-align'); if (ta && ['left', 'center', 'right'].includes(ta)) style.h = ta;
      const text = (td.innerText !== undefined ? td.innerText : td.textContent).replace(/ /g, ' ').replace(/\n$/, '');
      const cs = +td.getAttribute('colspan') || 1; const rs = +td.getAttribute('rowspan') || 1;
      grid[ri][ci] = { text, style, cs, rs };
      for (let y = 0; y < rs; y++) for (let x = 0; x < cs; x++) { taken[ri + y] = taken[ri + y] || []; taken[ri + y][ci + x] = true; }
      ci += cs;
    }
  });
  return { clipId, rows: grid.map((r) => Array.from({ length: Math.max(0, r.length) }, (_, i) => r[i] || null)) };
}
