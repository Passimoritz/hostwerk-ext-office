/*
 * Built-in templates (formats.js TEMPLATES, kind 'slides'): blank, pitch, lesson, photo. Sample text is content, not UI,
 * so both languages live here (picked by document.documentElement.lang). Pure – runs in Node.
 */
import { newDoc, addSlide, shape, textBox, bodyFromText, body, para, run, chart, table, cell, uid, updateSlide } from './model.js';

const L = (lang, en, de) => (lang === 'en' ? en : de);
export const docLang = () => (typeof document !== 'undefined' && document.documentElement && document.documentElement.lang === 'en' ? 'en' : 'de');

function add(doc, kind, texts) { const r = addSlide(doc, kind, undefined, texts); return { doc: r.doc, id: r.slide.id }; }
const put = (doc, id, els) => updateSlide(doc, id, (s) => ({ ...s, elements: [...s.elements, ...els] }));
const withNotes = (doc, id, notes) => updateSlide(doc, id, { notes });

export function buildTemplate(id = 'blank', { lang = docLang(), size = '16:9', theme = 'hostwerk', font, author = '' } = {}) {
  let doc = newDoc({ size, theme, font, author, title: '' });
  const W = doc.size.w; const H = doc.size.h;
  const X = (f) => Math.round(W * f); const Y = (f) => Math.round(H * f);
  let r;
  if (id === 'pitch') {
    doc.meta.title = L(lang, 'Pitch deck', 'Pitch-Deck');
    r = add(doc, 'title', { ctrTitle: L(lang, 'Your company', 'Dein Unternehmen'), subTitle: L(lang, 'One sentence about what you do – and for whom', 'Ein Satz, was ihr macht – und für wen') });
    doc = withNotes(r.doc, r.id, L(lang, 'Introduce yourself and the idea in 30 seconds.', 'Stell dich und die Idee in 30 Sekunden vor.'));
    r = add(doc, 'titleContent', { title: L(lang, 'The problem', 'Das Problem'), body: L(lang, 'Who has the problem?\nHow do they solve it today?\nWhy is that not good enough?', 'Wer hat das Problem?\nWie wird es heute gelöst?\nWarum reicht das nicht?') });
    doc = r.doc;
    r = add(doc, 'twoContent', { title: L(lang, 'Our solution', 'Unsere Lösung'), body: L(lang, 'What the product does\nThe one thing that makes it different\nHow customers benefit', 'Was das Produkt macht\nDas eine, was es besonders macht\nWas Kunden davon haben'), body2: '' });
    doc = r.doc;
    {
      const s = doc.slides.find((x) => x.id === r.id);
      const b2 = s.elements.find((e) => e.ph === 'body2');
      doc = updateSlide(doc, r.id, (sl) => ({ ...sl, elements: sl.elements.filter((e) => e.id !== b2.id) }));
      doc = put(doc, r.id, [
        shape('roundRect', b2.x, b2.y + Y(0.02), b2.w, b2.h - Y(0.06), { fill: { type: 'grad', angle: 45, stops: [{ pos: 0, color: 'accent1' }, { pos: 100, color: 'accent2' }] }, line: null, text: body([para(L(lang, 'Screenshot or product photo', 'Screenshot oder Produktfoto'), { align: 'ctr' }, { color: 'lt1', size: 20, b: true })], { anchor: 'ctr' }) }),
      ]);
    }
    r = add(doc, 'titleOnly', { title: L(lang, 'Market', 'Markt') });
    doc = put(r.doc, r.id, [
      { ...chart('col', X(0.07), Y(0.25), X(0.55), Y(0.65)), chart: { kind: 'col', title: L(lang, 'Market size (bn €)', 'Marktgröße (Mrd. €)'), legend: false, labels: true, categories: ['2024', '2025', '2026', '2027'], series: [{ name: L(lang, 'Market', 'Markt'), values: [1.2, 1.8, 2.6, 3.9], color: 'accent1' }] } },
      ...[['TAM', L(lang, '€3.9 bn', '3,9 Mrd. €'), 0.27], ['SAM', L(lang, '€850 m', '850 Mio. €'), 0.47], ['SOM', L(lang, '€40 m', '40 Mio. €'), 0.67]].map(([k, v, y]) => textBox(X(0.66), Y(y), X(0.27), Y(0.16), body([para(k, {}, { size: 14, color: 'tx2', b: true }), para(v, {}, { size: 28, color: 'accent1', b: true, font: '+mj' })], { autofit: 'none' }))),
    ]);
    r = add(doc, 'titleOnly', { title: L(lang, 'Business model', 'Geschäftsmodell') });
    {
      const t = table(4, 3, X(0.07), Y(0.27), X(0.86), Y(0.5));
      const rows = [[L(lang, 'Plan', 'Tarif'), L(lang, 'Price', 'Preis'), L(lang, 'For', 'Für')], ['Starter', L(lang, '€9 / month', '9 € / Monat'), L(lang, 'Individuals', 'Einzelne')], ['Team', L(lang, '€29 / month', '29 € / Monat'), L(lang, 'Small teams', 'Kleine Teams')], ['Business', L(lang, 'on request', 'auf Anfrage'), L(lang, 'Companies', 'Firmen')]];
      t.rows = t.rows.map((row, ri) => ({ ...row, cells: row.cells.map((c, ci) => cell(rows[ri][ci], { })) }));
      t.rows.forEach((row) => row.cells.forEach((c) => { c.text.paras[0].runs[0] && (c.text.paras[0].runs[0].size = 18); }));
      doc = put(r.doc, r.id, [t]);
    }
    r = add(doc, 'titleOnly', { title: 'Team' });
    doc = r.doc;
    {
      const people = [['Alex Muster', 'CEO'], ['Sam Beispiel', 'CTO'], ['Kim Vorlage', L(lang, 'Design', 'Design')]];
      const els = [];
      people.forEach(([n, role], i) => {
        const cx = X(0.2 + i * 0.3);
        els.push(shape('ellipse', cx - Y(0.14), Y(0.3), Y(0.28), Y(0.28), { fill: { type: 'solid', color: `accent${i + 1}+40` }, line: null, text: body([para(n.split(' ').map((w) => w[0]).join(''), { align: 'ctr' }, { size: 32, color: 'lt1', b: true })], { anchor: 'ctr' }) }));
        els.push(textBox(cx - X(0.13), Y(0.62), X(0.26), Y(0.14), body([para(n, { align: 'ctr' }, { size: 20, b: true }), para(role, { align: 'ctr' }, { size: 16, color: 'tx2' })])));
      });
      doc = put(doc, r.id, els);
    }
    r = add(doc, 'titleContent', { title: L(lang, 'Traction & roadmap', 'Zugkraft & Fahrplan'), body: L(lang, '1,200 active users after 6 months\nFirst paying customers\nNext: launch in 3 countries', '1.200 aktive Nutzer nach 6 Monaten\nErste zahlende Kunden\nAls Nächstes: Start in 3 Ländern') });
    doc = r.doc;
    r = add(doc, 'section', { title: L(lang, 'Thank you', 'Danke'), body: L(lang, 'hello@example.com · example.com', 'hallo@beispiel.de · beispiel.de') });
    doc = r.doc;
    doc = { ...doc, sections: [{ id: uid('sec'), name: L(lang, 'Intro', 'Einstieg'), start: doc.slides[0].id }, { id: uid('sec'), name: L(lang, 'Business', 'Geschäft'), start: doc.slides[3].id }] };
  } else if (id === 'lesson') {
    doc.meta.title = L(lang, 'Lesson', 'Unterricht');
    r = add(doc, 'title', { ctrTitle: L(lang, 'Topic of the lesson', 'Thema der Stunde'), subTitle: L(lang, 'Name · Course · Date', 'Name · Kurs · Datum') });
    doc = r.doc;
    r = add(doc, 'titleContent', { title: L(lang, 'Agenda', 'Ablauf') });
    {
      const items = L(lang, 'Warm-up question\nNew content\nPractice in pairs\nSummary', 'Einstiegsfrage\nNeuer Inhalt\nÜbung zu zweit\nZusammenfassung').split('\n');
      doc = updateSlide(r.doc, r.id, (s) => ({ ...s, elements: s.elements.map((e) => (e.ph === 'body' ? { ...e, text: { ...e.text, paras: items.map((t) => ({ ...para(t), bullet: { num: 'arabicPeriod' } })) } } : e)) }));
    }
    r = add(doc, 'section', { title: L(lang, 'New content', 'Neuer Inhalt'), body: L(lang, 'What we learn today', 'Was wir heute lernen') });
    doc = r.doc;
    r = add(doc, 'comparison', { title: L(lang, 'Compare', 'Vergleich'), body: L(lang, 'Before', 'Vorher'), body2: L(lang, 'How it used to be\nTypical problems', 'Wie es früher war\nTypische Probleme'), body3: L(lang, 'After', 'Nachher'), body4: L(lang, 'How it works now\nWhat improves', 'Wie es jetzt geht\nWas besser wird') });
    doc = r.doc;
    r = add(doc, 'titleOnly', { title: L(lang, 'Three key facts', 'Drei Kernaussagen') });
    {
      const els = [];
      const facts = L(lang, 'First idea\nSecond idea\nThird idea', 'Erster Gedanke\nZweiter Gedanke\nDritter Gedanke').split('\n');
      facts.forEach((f, i) => {
        const x = X(0.07 + i * 0.3);
        els.push(shape('roundRect', x, Y(0.3), X(0.26), Y(0.5), { fill: { type: 'solid', color: `accent${i + 1}+80` }, line: { color: `accent${i + 1}`, width: 19050, dash: 'solid' }, text: body([para(String(i + 1), { align: 'ctr' }, { size: 44, b: true, color: `accent${i + 1}`, font: '+mj' }), para(f, { align: 'ctr' }, { size: 20, color: 'tx1' })], { anchor: 'ctr' }) }));
      });
      doc = put(r.doc, r.id, els);
      doc = updateSlide(doc, r.id, (s) => ({ ...s, animations: s.elements.filter((e) => !e.ph).map((e, k) => ({ id: uid('a'), el: e.id, effect: 'fade', trigger: k ? 'click' : 'click', duration: 500, delay: 0 })) }));
    }
    r = add(doc, 'titleContent', { title: L(lang, 'Summary & questions', 'Zusammenfassung & Fragen'), body: L(lang, 'What was new for you?\nWhat is still unclear?\nHomework: …', 'Was war neu für dich?\nWas ist noch unklar?\nHausaufgabe: …') });
    doc = withNotes(r.doc, r.id, L(lang, 'Leave 5 minutes for questions.', 'Plane 5 Minuten für Fragen ein.'));
  } else if (id === 'photo') {
    doc.meta.title = L(lang, 'Photo album', 'Fotoalbum');
    r = add(doc, 'title', { ctrTitle: L(lang, 'Photo album', 'Fotoalbum'), subTitle: L(lang, 'Summer 2026 · by you', 'Sommer 2026 · von dir') });
    doc = r.doc;
    for (const cap of [L(lang, 'The beginning', 'Der Anfang'), L(lang, 'On the way', 'Unterwegs'), L(lang, 'The best moment', 'Der schönste Moment')]) {
      r = add(doc, 'pictureCaption', { title: cap, body: L(lang, 'A few words about this picture.', 'Ein paar Worte zu diesem Bild.') });
      doc = r.doc;
    }
    r = add(doc, 'blank');
    {
      const els = [];
      for (let i = 0; i < 4; i++) {
        const x = X(0.06 + (i % 2) * 0.45); const y = Y(0.07 + Math.floor(i / 2) * 0.45);
        els.push(shape('rect', x, y, X(0.43), Y(0.41), { name: `Photo ${i + 1}`, fill: { type: 'solid', color: 'bg2' }, line: null, text: body([para(L(lang, 'Double-click to add a photo', 'Doppelklick, um ein Foto einzufügen'), { align: 'ctr' }, { size: 14, color: 'tx2' })], { anchor: 'ctr' }), photoSlot: true }));
      }
      doc = put(r.doc, r.id, els);
    }
    r = add(doc, 'section', { title: L(lang, 'The end', 'Ende'), body: L(lang, 'Thanks for watching', 'Danke fürs Anschauen') });
    doc = r.doc;
    doc = { ...doc, slides: doc.slides.map((s) => ({ ...s, transition: { type: 'fade', dir: 'l', duration: 700, advanceAfter: null } })) };
  } else {
    r = add(doc, 'title', {});
    doc = r.doc;
  }
  return doc;
}

export { run, bodyFromText };
