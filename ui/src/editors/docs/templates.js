/*
 * Built-in document templates (ids from formats.js TEMPLATES, kind 'docs'). Sample texts in German or English
 * depending on Hostwerk's language; written in any format the chunk can write.
 */
import { defaultModel, para, heading, bullets, tasks, table, cell, txt, b, i, fixDoc } from './io/pm.js';
import { refreshToc } from './io/index.js';

const L = (de, en, lang) => (lang === 'de' ? de : en);
const today = (lang) => new Date().toLocaleDateString(lang === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: 'long', year: 'numeric' });
const small = (t) => txt(t, [{ type: 'textStyle', attrs: { fontSize: '8pt', color: '#555555', fontFamily: null } }]);
const muted = (t) => txt(t, [{ type: 'textStyle', attrs: { color: '#595959', fontFamily: null, fontSize: null } }]);
const shade = '#DCE6F2';

export function templateModel(id, settings, lang = (typeof document !== 'undefined' && document.documentElement.lang) || 'en') {
  const m = defaultModel(settings);
  const author = (settings && settings.general && settings.general.author) || L('Alex Beispiel', 'Alex Example', lang);
  let content = [para('')];
  switch (id) {
    case 'letter': {
      // DIN 5008 form B: return line + address window starting 45 mm from the top edge
      m.page.margins = { top: 20, right: 20, bottom: 20, left: 25 };
      m.defaults = { ...m.defaults, spaceAfter: 0, lineHeight: 1.15 };
      m.meta.title = L('Brief', 'Letter', lang);
      content = [
        para([b(author)], { textAlign: 'right' }),
        para(L('Musterstraße 1 · 12345 Musterstadt', '1 Sample Street · Sampletown AB1 2CD', lang), { textAlign: 'right' }),
        para(L('Telefon 0123 456789 · alex@example.org', 'Phone 0123 456789 · alex@example.org', lang), { textAlign: 'right' }),
        para([small(L(`${author} · Musterstraße 1 · 12345 Musterstadt`, `${author} · 1 Sample Street · Sampletown`, lang))], { spaceBefore: 36 }),
        para(L('Firma Empfänger GmbH', 'Recipient Ltd', lang), { spaceBefore: 6 }),
        para(L('Frau Erika Mustermann', 'Ms Erika Sample', lang)),
        para(L('Hauptstraße 12', '12 High Street', lang)),
        para(L('54321 Beispielstadt', 'Exampleton EX4 5MP', lang)),
        para(today(lang), { textAlign: 'right', spaceBefore: 60 }),
        para([b(L('Ihre Anfrage vom 3. März – Angebot', 'Your enquiry of 3 March – quotation', lang))], { spaceBefore: 24, spaceAfter: 18 }),
        para(L('Sehr geehrte Frau Mustermann,', 'Dear Ms Sample,', lang), { spaceAfter: 10 }),
        para(L('vielen Dank für Ihre Anfrage. Gern senden wir Ihnen unser Angebot für die gewünschten Leistungen. Alle Einzelheiten finden Sie in der Anlage.', 'Thank you for your enquiry. We are pleased to send you our quotation for the services you asked about. You will find all details in the enclosure.', lang), { spaceAfter: 10 }),
        para(L('Für Rückfragen stehe ich Ihnen jederzeit gern zur Verfügung.', 'If you have any questions, please do not hesitate to contact me.', lang), { spaceAfter: 10 }),
        para(L('Mit freundlichen Grüßen', 'Kind regards', lang), { spaceAfter: 36 }),
        para(author),
        para([small(L('Anlage: Angebot Nr. 2024-017', 'Enclosure: quotation no. 2024-017', lang))], { spaceBefore: 24 }),
      ];
      m.footer = { text: L('Musterbank · IBAN DE00 1234 5678 9012 3456 78 · BIC MUSTDEFF', 'Sample Bank · Account 12345678 · Sort code 12-34-56', lang), align: 'center' };
      break;
    }
    case 'invoice': {
      m.meta.title = L('Rechnung', 'Invoice', lang);
      const items = [
        [L('Beratung (Stunden)', 'Consulting (hours)', lang), '8', '95,00', '760,00'],
        [L('Einrichtung Server', 'Server setup', lang), '1', '450,00', '450,00'],
        [L('Wartung (Monat)', 'Maintenance (month)', lang), '3', '120,00', '360,00'],
      ];
      const right = (t, bold) => cell([para(bold ? [b(t)] : t, { textAlign: 'right', spaceAfter: 0 })]);
      const left = (t, bold) => cell([para(bold ? [b(t)] : t, { spaceAfter: 0 })]);
      const head = [L('Leistung', 'Item', lang), L('Menge', 'Qty', lang), L('Einzelpreis €', 'Unit price €', lang), L('Gesamt €', 'Amount €', lang)].map((t, k) => ({ ...cell([para([b(t)], { textAlign: k ? 'right' : null, spaceAfter: 0 })], { backgroundColor: shade }, true) }));
      content = [
        para([b(L('Ihre Firma GmbH', 'Your Company Ltd', lang))], { textAlign: 'right' }),
        para(L('Musterstraße 1 · 12345 Musterstadt · USt-IdNr. DE123456789', '1 Sample Street · Sampletown · VAT no. GB123456789', lang), { textAlign: 'right' }),
        para(L('Kunde AG\nHerrn Max Kunde\nKundenweg 5\n54321 Beispielstadt', 'Customer plc\nMr Max Customer\n5 Customer Way\nExampleton EX4 5MP', lang).split('\n').flatMap((l, k) => (k ? [{ type: 'hardBreak' }, l] : [l])), { spaceBefore: 30, spaceAfter: 24 }),
        { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: L('Rechnung Nr. 2024-042', 'Invoice no. 2024-042', lang) }] },
        para([b(L('Rechnungsdatum: ', 'Invoice date: ', lang)), today(lang), '   ', b(L('Leistungszeitraum: ', 'Service period: ', lang)), L('März 2024', 'March 2024', lang)]),
        { type: 'table', content: [
          { type: 'tableRow', content: head },
          ...items.map((r) => ({ type: 'tableRow', content: [left(r[0]), right(r[1]), right(r[2]), right(r[3])] })),
          { type: 'tableRow', content: [{ ...left(L('Zwischensumme', 'Subtotal', lang)), attrs: { colspan: 3, rowspan: 1, colwidth: null, backgroundColor: null } }, right('1.570,00')] },
          { type: 'tableRow', content: [{ ...left(L('zzgl. 19 % USt.', 'VAT 20 %', lang)), attrs: { colspan: 3, rowspan: 1, colwidth: null, backgroundColor: null } }, right(L('298,30', '314,00', lang))] },
          { type: 'tableRow', content: [{ ...left(L('Gesamtbetrag', 'Total', lang), true), attrs: { colspan: 3, rowspan: 1, colwidth: null, backgroundColor: shade } }, { ...right(L('1.868,30', '1.884,00', lang), true), attrs: { colspan: 1, rowspan: 1, colwidth: null, backgroundColor: shade } }] },
        ] },
        para(L('Bitte überweisen Sie den Betrag innerhalb von 14 Tagen ohne Abzug auf das unten genannte Konto.', 'Please transfer the amount within 14 days to the account below.', lang), { spaceBefore: 12 }),
        para(L('Vielen Dank für Ihren Auftrag!', 'Thank you for your business.', lang)),
      ];
      m.footer = { text: L('Musterbank · IBAN DE00 1234 5678 9012 3456 78 · Seite {page} von {pages}', 'Sample Bank · Account 12345678 · Page {page} of {pages}', lang), align: 'center' };
      break;
    }
    case 'cv': {
      m.meta.title = L('Lebenslauf', 'CV', lang);
      const row = (when, what, detail) => ({ type: 'tableRow', content: [cell([para([muted(when)], { spaceAfter: 2 })], { colwidth: [150] }), cell([para([b(what)], { spaceAfter: 0 }), para(detail, { spaceAfter: 4 })])] });
      content = [
        para(author, { pstyle: 'title' }),
        para(L('Softwareentwicklerin · Musterstadt · alex@example.org · 0123 456789', 'Software developer · Sampletown · alex@example.org · 0123 456789', lang), { pstyle: 'subtitle' }),
        heading(1, L('Profil', 'Profile', lang)),
        para(L('Erfahrene Entwicklerin mit Schwerpunkt Web und Datenbanken. Arbeitet gern im Team, erklärt Technik verständlich und liefert verlässlich.', 'Experienced developer focusing on web and databases. Enjoys teamwork, explains technology clearly and delivers reliably.', lang)),
        heading(1, L('Berufserfahrung', 'Experience', lang)),
        { type: 'table', content: [row(L('2020 – heute', '2020 – today', lang), L('Senior Entwicklerin, Beispiel AG', 'Senior developer, Example plc', lang), L('Leitung eines Teams von vier Personen, Einführung automatischer Tests.', 'Led a team of four, introduced automated testing.', lang)), row('2016 – 2020', L('Entwicklerin, Muster GmbH', 'Developer, Sample Ltd', lang), L('Web-Anwendungen für Kunden aus dem Handel.', 'Web applications for retail clients.', lang))] },
        heading(1, L('Ausbildung', 'Education', lang)),
        { type: 'table', content: [row('2012 – 2016', L('B.Sc. Informatik, Universität Musterstadt', 'BSc Computer Science, University of Sampletown', lang), L('Abschlussnote 1,7', 'First-class honours', lang))] },
        heading(1, L('Kenntnisse', 'Skills', lang)),
        bullets([L('JavaScript, TypeScript, Node.js, React', 'JavaScript, TypeScript, Node.js, React', lang), L('SQL, PostgreSQL, Datenmodellierung', 'SQL, PostgreSQL, data modelling', lang), L('Englisch fließend, Spanisch Grundkenntnisse', 'German fluent, Spanish basic', lang)]),
      ];
      break;
    }
    case 'report': {
      m.meta.title = L('Bericht', 'Report', lang);
      content = [
        para('', { spaceBefore: 180 }),
        para(L('Jahresbericht 2024', 'Annual report 2024', lang), { pstyle: 'title', textAlign: 'center' }),
        para(L('Projekt „Neue Wege“ – Ergebnisse und Ausblick', 'Project "New Ways" – results and outlook', lang), { pstyle: 'subtitle', textAlign: 'center' }),
        para([i(`${author} · ${today(lang)}`)], { textAlign: 'center', spaceBefore: 120 }),
        { type: 'pageBreak' },
        { type: 'toc', attrs: { title: L('Inhalt', 'Contents', lang), entries: [] } },
        { type: 'pageBreak' },
        heading(1, L('Zusammenfassung', 'Summary', lang)),
        para(L('Dieses Kapitel fasst die wichtigsten Ergebnisse auf einer Seite zusammen. Wer nur wenig Zeit hat, liest hier.', 'This chapter sums up the key results on one page. If you are short of time, read this.', lang)),
        { type: 'callout', attrs: { kind: 'info' }, content: [para([b(L('Kernaussage: ', 'Key point: ', lang)), L('Die Ziele wurden zu 92 % erreicht.', '92 % of the goals were reached.', lang)])] },
        heading(1, L('Ausgangslage', 'Background', lang)),
        para(L('Beschreibe hier, warum das Projekt gestartet wurde und welche Fragen es beantworten soll.', 'Describe why the project started and which questions it should answer.', lang)),
        heading(2, L('Ziele', 'Goals', lang)),
        bullets([L('Kosten um 10 % senken', 'Cut costs by 10 %', lang), L('Zufriedenheit der Kundschaft erhöhen', 'Raise customer satisfaction', lang), L('Abläufe vereinfachen', 'Simplify processes', lang)]),
        heading(1, L('Ergebnisse', 'Results', lang)),
        table([[L('Kennzahl', 'Metric', lang), L('Ziel', 'Target', lang), L('Ist', 'Actual', lang)], [L('Kosten', 'Costs', lang), '−10 %', '−12 %'], [L('Zufriedenheit', 'Satisfaction', lang), '4,5', '4,3']]),
        heading(2, L('Bewertung', 'Assessment', lang)),
        para([L('Die Ergebnisse liegen insgesamt im Plan', 'Overall, the results are on track', lang), { type: 'footnote', attrs: { text: L('Stand der Daten: 31. Dezember 2024.', 'Data as of 31 December 2024.', lang) } }, '.']),
        heading(1, L('Ausblick', 'Outlook', lang)),
        para(L('Was als Nächstes kommt, wer es macht und bis wann.', 'What comes next, who does it and by when.', lang)),
      ];
      m.titlePage = true;
      m.header = { text: L('Jahresbericht 2024', 'Annual report 2024', lang), align: 'right' };
      m.footer = { text: L('Seite {page} von {pages}', 'Page {page} of {pages}', lang), align: 'center' };
      break;
    }
    case 'minutes': {
      m.meta.title = L('Protokoll', 'Minutes', lang);
      const kv = (k, v) => ({ type: 'tableRow', content: [cell([para([b(k)], { spaceAfter: 0 })], { backgroundColor: shade, colwidth: [160] }), cell([para(v, { spaceAfter: 0 })])] });
      content = [
        heading(1, L('Protokoll: Teambesprechung', 'Minutes: team meeting', lang)),
        { type: 'table', content: [kv(L('Datum', 'Date', lang), today(lang)), kv(L('Zeit', 'Time', lang), '10:00 – 11:00'), kv(L('Ort', 'Place', lang), L('Raum 2 / online', 'Room 2 / online', lang)), kv(L('Leitung', 'Chair', lang), author), kv(L('Protokoll', 'Minutes by', lang), L('Sam Muster', 'Sam Sample', lang))] },
        heading(2, L('Teilnehmende', 'Attendees', lang)),
        para(L('Alex, Sam, Kim, Jo (entschuldigt: Lou)', 'Alex, Sam, Kim, Jo (apologies: Lou)', lang)),
        heading(2, L('Tagesordnung', 'Agenda', lang)),
        bullets([L('Stand der Projekte', 'Project status', lang), L('Urlaubsplanung', 'Holiday planning', lang), L('Verschiedenes', 'Any other business', lang)], true),
        heading(2, L('Beschlüsse', 'Decisions', lang)),
        bullets([L('Release 2.0 erscheint am 15. des Monats.', 'Release 2.0 ships on the 15th.', lang), L('Neue Besprechungszeit: dienstags 9:30.', 'New meeting time: Tuesdays 9:30.', lang)]),
        heading(2, L('Aufgaben', 'Action items', lang)),
        table([[L('Aufgabe', 'Task', lang), L('Wer', 'Who', lang), L('Bis', 'Due', lang)], [L('Testplan schreiben', 'Write test plan', lang), 'Kim', L('Freitag', 'Friday', lang)], [L('Kunden informieren', 'Inform customers', lang), 'Jo', L('nächste Woche', 'next week', lang)]]),
        para(L('Nächste Besprechung: in zwei Wochen, gleiche Zeit.', 'Next meeting: in two weeks, same time.', lang), { spaceBefore: 12 }),
      ];
      break;
    }
    case 'notes': {
      m.meta.title = L('Notizen', 'Notes', lang);
      content = [
        heading(1, L('Notizen', 'Notes', lang)),
        para([i(today(lang))]),
        heading(2, L('Zu erledigen', 'To do', lang)),
        tasks([[true, L('Angebot verschicken', 'Send the quotation', lang)], [false, L('Termin mit Kim abstimmen', 'Agree a date with Kim', lang)], [false, L('Präsentation vorbereiten', 'Prepare the presentation', lang)]]),
        heading(2, L('Ideen', 'Ideas', lang)),
        bullets([L('Newsletter einmal im Monat', 'Monthly newsletter', lang), L('Kurze Anleitung als Video', 'Short how-to video', lang)]),
        heading(2, L('Merken', 'Remember', lang)),
        para(L('Tipp: Tippe „[ ] “ am Zeilenanfang für eine neue Aufgabe, „# “ für eine Überschrift.', 'Tip: type "[ ] " at the start of a line for a new task, "# " for a heading.', lang)),
      ];
      break;
    }
    default: content = [para('')];
  }
  m.doc = fixDoc({ type: 'doc', content });
  if (id === 'report') refreshToc(m.doc);
  return m;
}
