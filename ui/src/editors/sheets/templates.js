/* Built-in templates (TEMPLATES in lib/formats.js): blank, budget, invoice, schedule, tracker, inventory – German or English. */
import { Book } from './model.js';

const T = {
  en: {
    sheet: 'Sheet1',
    budget: { name: 'Budget', title: 'Household budget', month: 'Month', cat: 'Category', plan: 'Planned', act: 'Actual', diff: 'Difference', share: 'Share', income: 'Income', expenses: 'Expenses', totalIn: 'Total income', totalOut: 'Total expenses', balance: 'Balance', chart: 'Planned vs. actual',
      inc: [['Salary', 3200, 3200], ['Side job', 400, 350], ['Other', 50, 80]], exp: [['Rent', 1100, 1100], ['Utilities', 180, 195], ['Groceries', 450, 510], ['Transport', 160, 140], ['Insurance', 120, 120], ['Phone & internet', 60, 60], ['Leisure', 200, 260], ['Savings', 500, 500], ['Other', 100, 75]] },
    invoice: { name: 'Invoice', title: 'INVOICE', from: ['Your Company Ltd', '1 Main Street', '12345 Springfield', 'hello@example.com'], to: 'Bill to', cust: ['Customer Inc.', 'Attn. Jane Doe', '99 Market Road', '54321 Shelbyville'], no: 'Invoice no.', date: 'Date', due: 'Due date', pos: 'Pos.', desc: 'Description', qty: 'Qty', price: 'Unit price', total: 'Amount', items: [['Consulting (hours)', 8, 95], ['Web design', 1, 1200], ['Hosting (12 months)', 12, 9.9]], sub: 'Subtotal', vat: 'Sales tax', vatRate: 0.08, sum: 'Total due', thanks: 'Thank you for your business. Please pay within 14 days.', invNo: 'INV-2026-001' },
    schedule: { name: 'Schedule', title: 'Weekly schedule', time: 'Time', days: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'], sample: [[1, 1, 'Team meeting'], [3, 3, 'Gym'], [5, 2, 'Project work'], [9, 4, 'Dentist']] },
    tracker: { name: 'Projects', title: 'Project tracker', cols: ['Task', 'Owner', 'Start', 'Due', 'Status', 'Progress', 'Notes'], status: ['Not started', 'In progress', 'Done', 'Blocked'],
      rows: [['Kick-off meeting', 'Anna', 0, 1, 'Done', 1, ''], ['Requirements', 'Ben', 1, 7, 'Done', 1, ''], ['Design', 'Chris', 7, 18, 'In progress', 0.6, 'Review on Friday'], ['Implementation', 'Ben', 14, 40, 'In progress', 0.2, ''], ['Testing', 'Dana', 35, 50, 'Not started', 0, ''], ['Launch', 'Anna', 50, 52, 'Blocked', 0, 'Waiting for approval']], prompt: 'Pick a status from the list' },
    inventory: { name: 'Inventory', title: 'Inventory list', cols: ['Item no.', 'Item', 'Category', 'In stock', 'Min. stock', 'Unit price', 'Stock value', 'Reorder?'], yes: 'Yes', total: 'Total',
      rows: [['A-1001', 'Printer paper A4', 'Office', 40, 20, 4.5], ['A-1002', 'Ballpoint pens (10)', 'Office', 8, 10, 3.2], ['B-2001', 'USB-C cable', 'IT', 15, 5, 7.9], ['B-2002', 'Wireless mouse', 'IT', 3, 4, 19.9], ['C-3001', 'Coffee beans 1 kg', 'Kitchen', 6, 4, 14.5], ['C-3002', 'Dish soap', 'Kitchen', 2, 3, 2.1]] },
  },
  de: {
    sheet: 'Tabelle1',
    budget: { name: 'Budget', title: 'Haushaltsbuch', month: 'Monat', cat: 'Kategorie', plan: 'Geplant', act: 'Tatsächlich', diff: 'Differenz', share: 'Anteil', income: 'Einnahmen', expenses: 'Ausgaben', totalIn: 'Summe Einnahmen', totalOut: 'Summe Ausgaben', balance: 'Saldo', chart: 'Geplant und tatsächlich',
      inc: [['Gehalt', 3200, 3200], ['Nebenjob', 400, 350], ['Sonstiges', 50, 80]], exp: [['Miete', 1100, 1100], ['Nebenkosten', 180, 195], ['Lebensmittel', 450, 510], ['Mobilität', 160, 140], ['Versicherungen', 120, 120], ['Telefon & Internet', 60, 60], ['Freizeit', 200, 260], ['Sparen', 500, 500], ['Sonstiges', 100, 75]] },
    invoice: { name: 'Rechnung', title: 'RECHNUNG', from: ['Deine Firma GmbH', 'Hauptstraße 1', '12345 Musterstadt', 'hallo@example.com'], to: 'Rechnung an', cust: ['Kunde AG', 'z. Hd. Erika Mustermann', 'Marktweg 99', '54321 Beispielstadt'], no: 'Rechnungsnr.', date: 'Datum', due: 'Fällig am', pos: 'Pos.', desc: 'Beschreibung', qty: 'Menge', price: 'Einzelpreis', total: 'Betrag', items: [['Beratung (Stunden)', 8, 95], ['Webdesign', 1, 1200], ['Hosting (12 Monate)', 12, 9.9]], sub: 'Zwischensumme', vat: 'USt. 19 %', vatRate: 0.19, sum: 'Gesamtbetrag', thanks: 'Vielen Dank für deinen Auftrag. Bitte zahle innerhalb von 14 Tagen.', invNo: 'RE-2026-001' },
    schedule: { name: 'Wochenplan', title: 'Wochenplan', time: 'Uhrzeit', days: ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'], sample: [[1, 1, 'Teambesprechung'], [3, 3, 'Sport'], [5, 2, 'Projektarbeit'], [9, 4, 'Zahnarzt']] },
    tracker: { name: 'Projekte', title: 'Projektplan', cols: ['Aufgabe', 'Verantwortlich', 'Beginn', 'Fällig', 'Status', 'Fortschritt', 'Notizen'], status: ['Nicht begonnen', 'In Arbeit', 'Erledigt', 'Blockiert'],
      rows: [['Auftaktbesprechung', 'Anna', 0, 1, 'Erledigt', 1, ''], ['Anforderungen', 'Ben', 1, 7, 'Erledigt', 1, ''], ['Entwurf', 'Chris', 7, 18, 'In Arbeit', 0.6, 'Abnahme am Freitag'], ['Umsetzung', 'Ben', 14, 40, 'In Arbeit', 0.2, ''], ['Test', 'Dana', 35, 50, 'Nicht begonnen', 0, ''], ['Start', 'Anna', 50, 52, 'Blockiert', 0, 'Wartet auf Freigabe']], prompt: 'Wähle einen Status aus der Liste' },
    inventory: { name: 'Lager', title: 'Lagerliste', cols: ['Artikelnr.', 'Artikel', 'Kategorie', 'Bestand', 'Mindestbestand', 'Einzelpreis', 'Lagerwert', 'Nachbestellen?'], yes: 'Ja', total: 'Summe',
      rows: [['A-1001', 'Druckerpapier A4', 'Büro', 40, 20, 4.5], ['A-1002', 'Kugelschreiber (10)', 'Büro', 8, 10, 3.2], ['B-2001', 'USB-C-Kabel', 'IT', 15, 5, 7.9], ['B-2002', 'Funkmaus', 'IT', 3, 4, 19.9], ['C-3001', 'Kaffeebohnen 1 kg', 'Küche', 6, 4, 14.5], ['C-3002', 'Spülmittel', 'Küche', 2, 3, 2.1]] },
  },
};

const ACCENT = '#2F5597';
const HEAD = { b: true, color: '#FFFFFF', fill: ACCENT, va: 'center' };
const TITLE = { b: true, size: 18, color: ACCENT };

export function buildTemplate(id, { loc, settings = {}, lang = 'en' } = {}) {
  const L = T[lang === 'de' ? 'de' : 'en'];
  const book = Book.blank({ loc, settings, sheetName: L.sheet });
  const s = book.sheets[0];
  const cur = loc.decimal === ',' ? '#,##0.00 €' : '$#,##0.00';
  const set = (r, c, v, style) => { if (typeof v === 'string' && v.startsWith('=')) book.setCell(s, r, c, { ...(s.cells.get(r * 16384 + c) || {}), v: null, f: v.slice(1) }); else book.input(s, r, c, typeof v === 'number' ? String(v).replace('.', loc.decimal) : v); if (style) { const cell = s.cells.get(r * 16384 + c) || { v: null }; book.setCell(s, r, c, { ...cell, s: book.styles.merge(cell.s, style) }); } };
  const fmt = (g, style) => book.format(s, [g], style);
  const widths = (arr) => arr.forEach((w, c) => s.colW.set(c, w));
  book.begin('template');
  switch (id) {
    case 'budget': {
      const B = L.budget;
      s.name = B.name;
      set(0, 0, B.title, TITLE); s.rowH.set(0, 34);
      widths([200, 110, 110, 110, 90]);
      const hdr = [B.cat, B.plan, B.act, B.diff, B.share];
      let r = 2;
      set(r, 0, B.income, { b: true, size: 13, color: ACCENT }); r++;
      hdr.forEach((h, c) => set(r, c, h, { ...HEAD, h: c ? 'right' : undefined })); r++;
      const inStart = r;
      for (const [n, p, a] of B.inc) { set(r, 0, n); set(r, 1, p); set(r, 2, a); set(r, 3, `=C${r + 1}-B${r + 1}`); set(r, 4, `=IFERROR(C${r + 1}/$C$${inStart + B.inc.length + 1},0)`); r++; }
      const inTot = r;
      set(r, 0, B.totalIn, { b: true }); set(r, 1, `=SUM(B${inStart + 1}:B${r})`, { b: true }); set(r, 2, `=SUM(C${inStart + 1}:C${r})`, { b: true }); set(r, 3, `=C${r + 1}-B${r + 1}`, { b: true });
      book.borders(s, { r1: r, c1: 0, r2: r, c2: 4 }, 'top');
      r += 2;
      set(r, 0, B.expenses, { b: true, size: 13, color: ACCENT }); r++;
      const hdr2 = r;
      hdr.forEach((h, c) => set(r, c, h, { ...HEAD, h: c ? 'right' : undefined })); r++;
      const exStart = r;
      for (const [n, p, a] of B.exp) { set(r, 0, n); set(r, 1, p); set(r, 2, a); set(r, 3, `=B${r + 1}-C${r + 1}`); set(r, 4, `=IFERROR(C${r + 1}/$C$${exStart + B.exp.length + 1},0)`); r++; }
      const exEnd = r - 1;
      set(r, 0, B.totalOut, { b: true }); set(r, 1, `=SUM(B${exStart + 1}:B${r})`, { b: true }); set(r, 2, `=SUM(C${exStart + 1}:C${r})`, { b: true }); set(r, 3, `=B${r + 1}-C${r + 1}`, { b: true });
      book.borders(s, { r1: r, c1: 0, r2: r, c2: 4 }, 'top');
      const exTot = r;
      r += 2;
      set(r, 0, B.balance, { b: true, size: 13 }); set(r, 1, `=B${inTot + 1}-B${exTot + 1}`, { b: true, size: 13 }); set(r, 2, `=C${inTot + 1}-C${exTot + 1}`, { b: true, size: 13 });
      book.borders(s, { r1: r, c1: 0, r2: r, c2: 2 }, 'bottomDouble');
      fmt({ r1: inStart, c1: 1, r2: r, c2: 3 }, { nf: cur });
      fmt({ r1: inStart, c1: 4, r2: exTot, c2: 4 }, { nf: '0.0%' });
      s.cf.push({ id: 'b1', ranges: [{ r1: inStart, c1: 3, r2: exTot, c2: 3 }], type: 'cell', op: 'lt', f1: '0', style: { color: '#C00000', fill: '#FFC7CE' } });
      s.cf.push({ id: 'b2', ranges: [{ r1: exStart, c1: 4, r2: exEnd, c2: 4 }], type: 'bar', color: '#5B9BD5' });
      s.charts.push({ id: 'c1', type: 'column', stacked: false, title: B.chart, legend: 'bottom', xTitle: '', yTitle: '', byRows: false, series: [{ name: B.plan, nameRef: `${B.name}!$B$${hdr2 + 1}`, catRef: `${B.name}!$A$${exStart + 1}:$A$${exEnd + 1}`, valRef: `${B.name}!$B$${exStart + 1}:$B$${exEnd + 1}` }, { name: B.act, nameRef: `${B.name}!$C$${hdr2 + 1}`, catRef: `${B.name}!$A$${exStart + 1}:$A$${exEnd + 1}`, valRef: `${B.name}!$C$${exStart + 1}:$C$${exEnd + 1}` }], range: null, anchor: { r: 2, c: 6, dx: 0, dy: 0, w: 520, h: 320 } });
      s.freeze = { rows: 1, cols: 0 };
      break;
    }
    case 'invoice': {
      const I = L.invoice;
      s.name = I.name;
      widths([50, 280, 70, 110, 120]);
      s.gridlines = false;
      set(0, 0, I.title, { ...TITLE, size: 22 }); s.rowH.set(0, 40);
      I.from.forEach((t, i) => set(i, 3, t, { h: 'right', color: '#595959' }));
      set(5, 0, I.to, { b: true, color: ACCENT });
      I.cust.forEach((t, i) => set(6 + i, 0, t));
      [[I.no, I.invNo], [I.date, '=TODAY()'], [I.due, '=D8+14']].forEach(([k, v], i) => { set(6 + i, 2, k, { b: true, h: 'right' }); set(6 + i, 3, v, { h: 'right' }); });
      fmt({ r1: 7, c1: 3, r2: 8, c2: 3 }, { nf: loc.shortDate });
      const h = 11;
      [I.pos, I.desc, I.qty, I.price, I.total].forEach((t, c) => set(h, c, t, { ...HEAD, h: c >= 2 ? 'right' : c === 0 ? 'center' : undefined }));
      s.rowH.set(h, 24);
      const first = h + 1;
      I.items.forEach(([d, q, p], i) => { const r = first + i; set(r, 0, i + 1, { h: 'center' }); set(r, 1, d); set(r, 2, q); set(r, 3, p); set(r, 4, `=C${r + 1}*D${r + 1}`); });
      const last = first + 7; // room for more lines
      for (let r = first + I.items.length; r <= last; r++) set(r, 4, `=IF(C${r + 1}="","",C${r + 1}*D${r + 1})`);
      book.borders(s, { r1: h, c1: 0, r2: last, c2: 4 }, 'all', { s: 'thin', c: '#BFBFBF' });
      fmt({ r1: first, c1: 3, r2: last + 3, c2: 4 }, { nf: cur });
      const t = last + 1;
      set(t, 3, I.sub, { h: 'right' }); set(t, 4, `=SUM(E${first + 1}:E${last + 1})`);
      set(t + 1, 3, lang === 'de' ? I.vat : `${I.vat} (${Math.round(I.vatRate * 100)} %)`.replace(' %', '%'), { h: 'right' }); set(t + 1, 4, `=ROUND(E${t + 1}*${String(I.vatRate)},2)`);
      set(t + 2, 3, I.sum, { b: true, h: 'right' }); set(t + 2, 4, `=E${t + 1}+E${t + 2}`, { b: true, fill: '#DEEAF6' });
      book.borders(s, { r1: t + 2, c1: 3, r2: t + 2, c2: 4 }, 'topBottom');
      set(t + 5, 0, I.thanks, { i: true, color: '#595959' });
      s.print = { ...s.print, area: `A1:E${t + 6}`, fit: 'width', orientation: 'portrait' };
      break;
    }
    case 'schedule': {
      const S = L.schedule;
      s.name = S.name;
      set(0, 0, S.title, TITLE); s.rowH.set(0, 34);
      widths([80, 130, 130, 130, 130, 130, 130, 130]);
      set(2, 0, S.time, HEAD);
      S.days.forEach((d, i) => set(2, i + 1, d, { ...HEAD, h: 'center' }));
      for (let i = 0; i <= 11; i++) { const r = 3 + i; set(r, 0, `${String(7 + i).padStart(2, '0')}:00`, { b: true, fill: '#DEEAF6', h: 'center' }); s.rowH.set(r, 30); }
      fmt({ r1: 3, c1: 0, r2: 14, c2: 0 }, { nf: 'hh:mm' });
      for (const [r, c, t] of S.sample) set(3 + r, c, t, { fill: '#FFF2CC', wrap: true, va: 'center' });
      fmt({ r1: 3, c1: 6, r2: 14, c2: 7 }, (st) => (st && st.fill ? null : { fill: '#F2F2F2' }));
      book.borders(s, { r1: 2, c1: 0, r2: 14, c2: 7 }, 'all', { s: 'thin', c: '#BFBFBF' });
      s.freeze = { rows: 3, cols: 1 };
      s.print = { ...s.print, orientation: 'landscape', fit: 'page' };
      break;
    }
    case 'tracker': {
      const P = L.tracker;
      s.name = P.name;
      set(0, 0, P.title, TITLE); s.rowH.set(0, 34);
      widths([220, 120, 100, 100, 120, 100, 260]);
      P.cols.forEach((t, c) => set(2, c, t, HEAD));
      s.rowH.set(2, 24);
      P.rows.forEach(([task, who, st, due, status, prog, note], i) => {
        const r = 3 + i;
        set(r, 0, task); set(r, 1, who); set(r, 2, `=TODAY()-20+${st}`); set(r, 3, `=TODAY()-20+${due}`); set(r, 4, status); set(r, 5, prog); if (note) set(r, 6, note);
      });
      const last = 3 + P.rows.length + 10;
      fmt({ r1: 3, c1: 2, r2: last, c2: 3 }, { nf: loc.shortDate });
      fmt({ r1: 3, c1: 5, r2: last, c2: 5 }, { nf: '0%' });
      const list = `"${P.status.join(',')}"`;
      s.dv.push({ id: 'v1', ranges: [{ r1: 3, c1: 4, r2: last, c2: 4 }], type: 'list', f1: list, allowBlank: true, dropdown: true, prompt: { title: P.cols[4], text: P.prompt }, error: { style: 'stop', title: '', text: '', show: true } });
      s.dv.push({ id: 'v2', ranges: [{ r1: 3, c1: 5, r2: last, c2: 5 }], type: 'decimal', op: 'between', f1: '0', f2: '1', allowBlank: true, error: { style: 'stop', show: true } });
      const g = { r1: 3, c1: 4, r2: last, c2: 4 };
      s.cf.push({ id: 't1', ranges: [g], type: 'cell', op: 'eq', f1: `"${P.status[2]}"`, style: { fill: '#C6EFCE', color: '#006100' } });
      s.cf.push({ id: 't2', ranges: [g], type: 'cell', op: 'eq', f1: `"${P.status[3]}"`, style: { fill: '#FFC7CE', color: '#9C0006' } });
      s.cf.push({ id: 't3', ranges: [g], type: 'cell', op: 'eq', f1: `"${P.status[1]}"`, style: { fill: '#FFEB9C', color: '#9C5700' } });
      s.cf.push({ id: 't4', ranges: [{ r1: 3, c1: 3, r2: last, c2: 3 }], type: 'formula', formula: `AND(D4<>"",D4<TODAY(),E4<>"${P.status[2]}")`, style: { color: '#C00000', b: true } });
      s.cf.push({ id: 't5', ranges: [{ r1: 3, c1: 5, r2: last, c2: 5 }], type: 'bar', color: '#63BE7B', stops: [{ type: 'num', val: '0' }, { type: 'num', val: '1' }] });
      s.filter = { range: { r1: 2, c1: 0, r2: 3 + P.rows.length - 1, c2: 6 }, cols: {} };
      s.freeze = { rows: 3, cols: 1 };
      s.print = { ...s.print, orientation: 'landscape', fit: 'width' };
      break;
    }
    case 'inventory': {
      const V = L.inventory;
      s.name = V.name;
      set(0, 0, V.title, TITLE); s.rowH.set(0, 34);
      widths([90, 200, 110, 80, 110, 100, 110, 120]);
      V.cols.forEach((t, c) => set(2, c, t, { ...HEAD, h: c >= 3 && c <= 6 ? 'right' : undefined }));
      const first = 3;
      V.rows.forEach(([no, name, cat, stock, min, price], i) => {
        const r = first + i;
        set(r, 0, no); set(r, 1, name); set(r, 2, cat); set(r, 3, stock); set(r, 4, min); set(r, 5, price);
        set(r, 6, `=D${r + 1}*F${r + 1}`); set(r, 7, `=IF(D${r + 1}<=E${r + 1},"${V.yes}","")`, { h: 'center' });
      });
      const last = first + V.rows.length - 1;
      fmt({ r1: first, c1: 5, r2: last + 1, c2: 6 }, { nf: cur });
      set(last + 1, 0, V.total, { b: true }); set(last + 1, 3, `=SUBTOTAL(9,D${first + 1}:D${last + 1})`, { b: true }); set(last + 1, 6, `=SUBTOTAL(9,G${first + 1}:G${last + 1})`, { b: true, nf: cur });
      book.borders(s, { r1: last + 1, c1: 0, r2: last + 1, c2: 7 }, 'top');
      s.cf.push({ id: 'i1', ranges: [{ r1: first, c1: 7, r2: last, c2: 7 }], type: 'cell', op: 'eq', f1: `"${V.yes}"`, style: { fill: '#FFC7CE', color: '#9C0006', b: true } });
      s.cf.push({ id: 'i2', ranges: [{ r1: first, c1: 3, r2: last, c2: 3 }], type: 'scale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] });
      s.filter = { range: { r1: 2, c1: 0, r2: last, c2: 7 }, cols: {} };
      s.freeze = { rows: 3, cols: 0 };
      break;
    }
    default: break;
  }
  book.end();
  // a template starts as a fresh file: no undo history
  book.undoStack = []; book.redoStack = [];
  book.calc.rebuild(); book.calc.recalc(null, { force: true });
  return book;
}
