/*
 * Office 1.1 helpers of the slides chunk – pure functions (Node tests):
 *   outline of a document (docx / odt / md / txt) → slides ("presentation from a document")
 *   Designer: alternative arrangements of a slide's content (offline)
 *   compare two versions slide by slide
 *   live fields in runs
 */
import { parseXml } from '../_shared/xmltree.js';
import { newSlide, para, body, run, bodyText, walk, uid, layoutByKind, image as imageEl, textBox, shape as shapeEl, slideTitle } from './model.js';

/* ---------------------------------------------------------------- outline of a document */
const ln = (n) => n.slice(n.indexOf(':') + 1);
const KS = (el, name) => (el && el.c ? el.c.filter((x) => typeof x !== 'string' && (!name || ln(x.n) === name)) : []);
const textIn = (el) => { let s = ''; const w = (n) => { for (const c of n.c || []) { if (typeof c === 'string') s += c; else if (ln(c.n) === 't' || ln(c.n) === 's' || ln(c.n) === 'span' || ln(c.n) === 'a' || ln(c.n) === 'r') w(c); else if (ln(c.n) === 'tab') s += ' '; else if (ln(c.n) === 'br' || ln(c.n) === 'line-break') s += ' '; else if (!['rPr', 'pPr', 'instrText', 'delText'].includes(ln(c.n))) w(c); } }; w(el); return s.replace(/\s+/g, ' ').trim(); };

/** outline items: [{ kind: 'title'|'h'|'p'|'li', level, text }] from Markdown / plain text */
export function outlineFromMd(md) {
  const out = [];
  for (const raw of String(md || '').split(/\r?\n/)) {
    const l = raw.replace(/\s+$/, '');
    if (!l.trim()) continue;
    let m;
    if ((m = /^(#{1,6})\s+(.*)$/.exec(l))) out.push({ kind: 'h', level: m[1].length, text: m[2].replace(/[*_`]/g, '').trim() });
    else if ((m = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/.exec(l))) out.push({ kind: 'li', level: Math.floor(m[1].replace(/\t/g, '  ').length / 2), text: m[3].replace(/\*\*|__|[*_`]/g, '').trim() });
    else if (/^(=+|-+)$/.test(l.trim()) && out.length && out[out.length - 1].kind === 'p') { const last = out.pop(); out.push({ kind: 'h', level: l.trim()[0] === '=' ? 1 : 2, text: last.text }); }
    else out.push({ kind: 'p', level: 0, text: l.replace(/\*\*|__|[*_`]/g, '').trim() });
  }
  return out;
}
/** from word/document.xml (+ styles.xml for heading names) */
export function outlineFromDocxXml(docXml, stylesXml = '') {
  const doc = parseXml(docXml);
  const headingOf = {};
  if (stylesXml) {
    for (const s of KS(parseXml(stylesXml), 'style')) {
      const id = s.a['w:styleId']; const nm = KS(s, 'name')[0];
      const name = (nm && nm.a['w:val']) || '';
      const ol = KS(KS(s, 'pPr')[0], 'outlineLvl')[0];
      if (/^title$/i.test(name)) headingOf[id] = 0;
      else if (ol) headingOf[id] = Number(ol.a['w:val']) + 1;
      else { const m = /heading\s*(\d)/i.exec(name); if (m) headingOf[id] = Number(m[1]); }
    }
  }
  const out = [];
  const body0 = KS(doc, 'body')[0] || doc;
  const paras = []; const w = (n) => { for (const c of KS(n)) { if (ln(c.n) === 'p') paras.push(c); else if (ln(c.n) !== 'sectPr') w(c); } }; w(body0);
  for (const p of paras) {
    const text = textIn(p);
    if (!text) continue;
    const pPr = KS(p, 'pPr')[0];
    const sty = pPr && KS(pPr, 'pStyle')[0] ? KS(pPr, 'pStyle')[0].a['w:val'] : '';
    const ol = pPr && KS(pPr, 'outlineLvl')[0];
    const num = pPr && KS(pPr, 'numPr')[0];
    let lvl = headingOf[sty];
    if (lvl === undefined) { const m = /^(?:heading|berschrift|titre)\s*(\d)$/i.exec(sty); if (m) lvl = Number(m[1]); else if (/^title$/i.test(sty)) lvl = 0; }
    if (lvl === undefined && ol) lvl = Number(ol.a['w:val']) + 1;
    if (lvl === 0) out.push({ kind: 'title', level: 0, text });
    else if (lvl) out.push({ kind: 'h', level: lvl, text });
    else if (num) { const il = KS(num, 'ilvl')[0]; out.push({ kind: 'li', level: il ? Number(il.a['w:val']) || 0 : 0, text }); }
    else out.push({ kind: 'p', level: 0, text });
  }
  return out;
}
/** from ODF content.xml */
export function outlineFromOdtXml(xml) {
  const doc = parseXml(xml);
  const out = [];
  const w = (n, depth) => {
    for (const c of KS(n)) {
      const name = ln(c.n);
      if (name === 'h') { const t = textIn(c); if (t) out.push({ kind: 'h', level: Number(c.a['text:outline-level'] || 1), text: t }); }
      else if (name === 'p') { const t = textIn(c); if (t) out.push(depth >= 0 ? { kind: 'li', level: depth, text: t } : /title/i.test(c.a['text:style-name'] || '') ? { kind: 'title', level: 0, text: t } : { kind: 'p', level: 0, text: t }); }
      else if (name === 'list') w(c, depth + 1);
      else if (!['tracked-changes', 'sequence-decls', 'variable-decls', 'forms'].includes(name)) w(c, depth);
    }
  };
  w(doc, -1);
  return out;
}

const sentences = (s) => (String(s).match(/[^.!?]+[.!?]+(\s|$)|[^.!?]+$/g) || [s]).map((x) => x.trim()).filter(Boolean);
/** shorten a paragraph to a bullet (first sentence, ≤ 120 characters) */
export function bulletOf(text, max = 120) {
  const first = sentences(text)[0] || text;
  return first.length > max ? `${first.slice(0, max - 1).replace(/\s+\S*$/, '')} …` : first;
}

/**
 * outline → slide plans: [{ kind: 'title'|'section'|'titleContent'|'twoContent', title, subtitle?, bullets: [{ text, level }] }]
 * level-1 headings become section slides when level-2 headings follow, else content slides; more than 7 bullets
 * split into "(continued)" slides.
 */
export function planSlides(outline, { maxBullets = 7, title: docTitle } = {}) {
  const items = outline.filter((x) => x.text);
  const plans = [];
  const hasH2 = items.some((x) => x.kind === 'h' && x.level >= 2);
  let titleItem = items.find((x) => x.kind === 'title') || (items[0] && items[0].kind === 'h' && items[0].level === 1 && items.filter((x) => x.kind === 'h' && x.level === 1).length === 1 ? items[0] : null);
  if (!titleItem && docTitle) titleItem = { kind: 'title', text: docTitle };
  if (titleItem) {
    const i = items.indexOf(titleItem);
    const sub = items.slice(i + 1).find((x) => x.kind === 'p');
    plans.push({ kind: 'title', title: titleItem.text, subtitle: sub && items.indexOf(sub) < (items.findIndex((x, k) => k > i && x.kind === 'h') + 1 || Infinity) ? bulletOf(sub.text, 160) : '', bullets: [] });
  }
  let cur = null;
  const flush = () => { if (cur && (cur.bullets.length || cur.kind === 'section')) plans.push(cur); cur = null; };
  for (const it of items) {
    if (it === titleItem) continue;
    if (it.kind === 'h' || it.kind === 'title') {
      flush();
      const section = hasH2 && it.level === 1;
      cur = { kind: section ? 'section' : 'titleContent', title: it.text, bullets: [] };
      if (section) { flush(); }
      continue;
    }
    if (!cur) cur = { kind: 'titleContent', title: plans.length ? '' : (docTitle || ''), bullets: [] };
    if (plans.length && plans[0].kind === 'title' && it.kind === 'p' && plans[0].subtitle && bulletOf(it.text, 160) === plans[0].subtitle && !cur.bullets.length && !cur.title) continue;
    cur.bullets.push({ text: it.kind === 'p' ? bulletOf(it.text) : it.text, level: it.kind === 'li' ? Math.min(4, it.level) : 0 });
    if (cur.bullets.length >= maxBullets) { const t0 = cur.title; flush(); cur = { kind: 'titleContent', title: t0 ? `${t0} (${'…'})` : '', bullets: [], cont: true }; }
  }
  flush();
  return plans.filter((p) => p.kind !== 'titleContent' || p.bullets.length);
}
/** plans → slides of `doc` (layouts of its theme) */
export function slidesFromPlans(doc, plans) {
  return plans.map((p) => {
    if (p.kind === 'title') return newSlide(doc, 'title', { ctrTitle: p.title, subTitle: p.subtitle || '', notes: p.notes || '' });
    if (p.kind === 'section') return newSlide(doc, layoutByKind(doc, 'section') && layoutByKind(doc, 'section').kind === 'section' ? 'section' : 'title', { title: p.title, ctrTitle: p.title, notes: p.notes || '' });
    const s = newSlide(doc, 'titleContent', { title: p.title, notes: p.notes || '' });
    const b = s.elements.find((e) => e.ph && /^body/.test(e.ph));
    if (b) b.text = { ...b.text, paras: p.bullets.map((x) => ({ ...para(x.text), level: x.level })) };
    return s;
  });
}
/** a compact text of the outline for the AI (≤ maxChars) */
export const outlineText = (outline, maxChars = 12000) => outline.map((x) => (x.kind === 'h' || x.kind === 'title' ? `${'#'.repeat(Math.max(1, x.level || 1))} ${x.text}` : x.kind === 'li' ? `${'  '.repeat(x.level)}- ${x.text}` : x.text)).join('\n').slice(0, maxChars);
/** parse the AI's answer: JSON [{ title, bullets: [..], notes }] (tolerates prose around it) */
export function plansFromAi(answer) {
  const s = String(answer || '');
  const i = s.indexOf('['); const j = s.lastIndexOf(']');
  if (i < 0 || j <= i) return null;
  try {
    const arr = JSON.parse(s.slice(i, j + 1));
    if (!Array.isArray(arr) || !arr.length) return null;
    return arr.map((x, k) => ({ kind: k === 0 && x.kind !== 'content' && !(x.bullets || []).length ? 'title' : x.kind === 'section' ? 'section' : 'titleContent', title: String(x.title || ''), subtitle: String(x.subtitle || ''), notes: String(x.notes || ''), bullets: (x.bullets || []).map((b) => (typeof b === 'string' ? { text: b, level: 0 } : { text: String(b.text || ''), level: Math.min(4, b.level || 0) })) }));
  } catch { return null; }
}

/* ---------------------------------------------------------------- Designer */
/**
 * alternative arrangements of a slide's content (title, text, pictures, other shapes) – offline, using the theme
 * colours. Returns [{ id, name, elements, background? }] (elements = full new element list for the slide).
 */
export function designIdeas(doc, slide) {
  const W = doc.size.w; const H = doc.size.h;
  const els = slide.elements || [];
  const title = els.find((e) => e.ph && /title/i.test(e.ph) && !/sub/i.test(e.ph)) || els.find((e) => e.text && bodyText(e.text).trim() && e.h < H * 0.25);
  const texts = els.filter((e) => e !== title && e.text && bodyText(e.text).trim() && e.type === 'shape');
  const pics = els.filter((e) => e.type === 'image');
  const others = els.filter((e) => e !== title && !texts.includes(e) && !pics.includes(e) && !(e.type === 'shape' && e.text && !bodyText(e.text).trim() && !e.fill));
  const m = Math.round(W * 0.06); const gap = Math.round(W * 0.03);
  const place = (e, x, y, w, h, extra = {}) => ({ ...e, x: Math.round(x), y: Math.round(y), w: Math.max(12700, Math.round(w)), h: Math.max(12700, Math.round(h)), rot: 0, ...extra });
  const cover = (img, x, y, w, h) => {
    // crop the picture to fill the box (keeps its proportions)
    const ar = (img.w * (1 - (img.crop?.l || 0) - (img.crop?.r || 0))) / Math.max(1, img.h * (1 - (img.crop?.t || 0) - (img.crop?.b || 0)));
    const br = w / h; const crop = ar > br ? { l: (1 - br / ar) / 2, r: (1 - br / ar) / 2, t: 0, b: 0 } : { t: (1 - ar / br) / 2, b: (1 - ar / br) / 2, l: 0, r: 0 };
    return place(img, x, y, w, h, { crop });
  };
  const withText = (e, patch) => (e && e.text ? { ...e, text: { ...e.text, ...patch } } : e);
  const ideas = [];
  const titleH = Math.round(H * 0.16);
  const tt = title ? place(withText(title, { anchor: 'b' }), m, m * 0.7, W - 2 * m, titleH) : null;
  const body0 = texts[0];
  const stackTexts = (x, y, w, h) => texts.map((e, i) => place(e, x, y + (h / Math.max(1, texts.length)) * i, w, h / Math.max(1, texts.length) - (texts.length > 1 ? gap / 2 : 0)));
  if (pics.length === 1) {
    const p = pics[0];
    const y0 = m * 0.7 + titleH + gap;
    ideas.push({ id: 'pic-right', name: 'Picture right', elements: [...others, tt, ...stackTexts(m, y0, W * 0.46 - m, H - y0 - m), cover(p, W * 0.52, 0, W * 0.48, H)].filter(Boolean) });
    ideas.push({ id: 'pic-left', name: 'Picture left', elements: [...others, cover(p, 0, 0, W * 0.45, H), title ? place(withText(title, { anchor: 'b' }), W * 0.5, m, W * 0.5 - m, titleH) : null, ...stackTexts(W * 0.5, m + titleH + gap, W * 0.5 - m, H - (m + titleH + gap) - m)].filter(Boolean) });
    const band = { ...shapeEl('rect', 0, H * 0.62, W, H * 0.38), id: uid('d'), name: 'Designer band', fill: { type: 'solid', color: 'tx1@70' }, line: null, text: body(), textbox: false };
    ideas.push({ id: 'pic-full', name: 'Full picture', elements: [cover(p, 0, 0, W, H), band, ...others, title ? place(withText(title, { anchor: 'ctr', lvl: (title.text.lvl || [{}]).map((l) => ({ ...l, color: 'bg1' })) }), m, H * 0.64, W - 2 * m, H * 0.17) : null, body0 ? place(withText(body0, { lvl: (body0.text.lvl || [{}]).map((l) => ({ ...l, color: 'bg1' })) }), m, H * 0.8, W - 2 * m, H * 0.16) : null, ...texts.slice(1)].filter(Boolean) });
    ideas.push({ id: 'pic-top', name: 'Picture on top', elements: [...others, cover(p, m, m, W - 2 * m, H * 0.5), title ? place(withText(title, { anchor: 't' }), m, H * 0.5 + m + gap, W - 2 * m, H * 0.14) : null, ...stackTexts(m, H * 0.66 + m, W - 2 * m, H * 0.34 - 2 * m)].filter(Boolean) });
  } else if (pics.length > 1) {
    const n = Math.min(pics.length, 6); const cols = n <= 3 ? n : Math.ceil(n / 2); const rows = Math.ceil(n / cols);
    const y0 = m * 0.7 + titleH + gap; const gw = (W - 2 * m - (cols - 1) * gap) / cols; const gh = (H - y0 - m - (rows - 1) * gap - (texts.length ? H * 0.12 : 0)) / rows;
    const grid = pics.slice(0, n).map((p, i) => cover(p, m + (i % cols) * (gw + gap), y0 + Math.floor(i / cols) * (gh + gap), gw, gh));
    ideas.push({ id: 'grid', name: 'Picture grid', elements: [...others, tt, ...grid, ...pics.slice(n), ...(texts.length ? stackTexts(m, H - m - H * 0.12, W - 2 * m, H * 0.12) : [])].filter(Boolean) });
    const sw = (W - 2 * m - (n - 1) * gap / 2) / n;
    ideas.push({ id: 'strip', name: 'Filmstrip', elements: [...others, tt, ...pics.slice(0, n).map((p, i) => cover(p, m + i * (sw + gap / 2), H * 0.32, sw, H * 0.42)), ...pics.slice(n), ...(texts.length ? stackTexts(m, H * 0.78, W - 2 * m, H * 0.16) : [])].filter(Boolean) });
  }
  if (!pics.length && body0) {
    const paras = body0.text.paras.filter((p) => (p.runs || []).some((r) => r.text));
    if (paras.length >= 2 && paras.length <= 6) {
      // one card per point
      const n = paras.length; const cols = n <= 3 ? n : Math.ceil(n / 2); const rows = Math.ceil(n / cols);
      const y0 = m * 0.7 + titleH + gap; const cw = (W - 2 * m - (cols - 1) * gap) / cols; const ch = (H - y0 - m - (rows - 1) * gap) / rows;
      const cards = paras.map((p, i) => ({ ...shapeEl('roundRect', m + (i % cols) * (cw + gap), y0 + Math.floor(i / cols) * (ch + gap), cw, ch), id: uid('d'), name: `Card ${i + 1}`, adj: [8000], fill: { type: 'solid', color: `accent${(i % 6) + 1}` }, line: null, text: { ...body([{ ...p, level: 0, bullet: false, align: 'ctr' }], { anchor: 'ctr' }), lvl: Array.from({ length: 9 }, () => ({ color: 'lt1', align: 'ctr', size: n > 3 ? 18 : 22, b: true })) } }));
      ideas.push({ id: 'cards', name: 'Cards', elements: [...others, tt, ...cards, ...texts.slice(1)].filter(Boolean) });
      // numbered process
      const pw = (W - 2 * m - (n - 1) * gap) / n;
      const steps = paras.flatMap((p, i) => {
        const cx = m + i * (pw + gap);
        const dot = { ...shapeEl('ellipse', cx + pw / 2 - H * 0.07, H * 0.36, H * 0.14, H * 0.14), id: uid('d'), name: `Step ${i + 1}`, fill: { type: 'solid', color: 'accent1' }, line: null, text: { ...body([para(String(i + 1), { align: 'ctr' })], { anchor: 'ctr' }), lvl: Array.from({ length: 9 }, () => ({ color: 'lt1', align: 'ctr', size: 24, b: true })) } };
        const txt = textBox(cx, H * 0.54, pw, H * 0.3, { ...body([{ ...p, level: 0, bullet: false, align: 'ctr' }], { anchor: 't' }), lvl: Array.from({ length: 9 }, () => ({ align: 'ctr', size: n > 4 ? 14 : 18 })) });
        return [dot, { ...txt, id: uid('d') }];
      });
      const lineY = H * 0.43;
      const conn = { ...shapeEl('rect', m + pw / 2, lineY - 6350, (n - 1) * (pw + gap), 12700), id: uid('d'), name: 'Process line', fill: { type: 'solid', color: 'accent1+60' }, line: null, text: body() };
      ideas.push({ id: 'process', name: 'Process', elements: [...others, tt, conn, ...steps, ...texts.slice(1)].filter(Boolean) });
    }
    // two columns / big statement
    if (paras.length > 3) {
      const half = Math.ceil(paras.length / 2);
      const y0 = m * 0.7 + titleH + gap; const cw = (W - 2 * m - gap) / 2;
      const a = { ...body0, ...place(body0, m, y0, cw, H - y0 - m), text: { ...body0.text, paras: body0.text.paras.slice(0, half), autofit: 'shrink' } };
      const b = { ...body0, id: uid('d'), ph: undefined, ...place(body0, m + cw + gap, y0, cw, H - y0 - m), text: { ...body0.text, paras: body0.text.paras.slice(half), autofit: 'shrink' }, textbox: true };
      ideas.push({ id: 'columns', name: 'Two columns', elements: [...others, tt, a, b, ...texts.slice(1)].filter(Boolean) });
    }
    const bar = { ...shapeEl('rect', 0, 0, W * 0.34, H), id: uid('d'), name: 'Designer panel', fill: { type: 'solid', color: 'accent1' }, line: null, text: body() };
    ideas.push({ id: 'panel', name: 'Side panel', elements: [bar, ...others, title ? place(withText(title, { anchor: 'ctr', lvl: (title.text.lvl || [{}]).map((l) => ({ ...l, color: 'lt1' })) }), m * 0.6, m, W * 0.34 - m * 1.2, H - 2 * m) : null, ...stackTexts(W * 0.34 + m, m, W * 0.66 - 2 * m, H - 2 * m)].filter(Boolean) });
  }
  // a chart or table: give it room next to or under the text
  const vis = others.filter((e) => e.type === 'chart' || e.type === 'table');
  if (!pics.length && vis.length === 1) {
    const v = vis[0]; const rest = others.filter((e) => e !== v);
    const y0 = m * 0.7 + titleH + gap;
    if (body0) ideas.push({ id: 'vis-right', name: 'Chart right', elements: [...rest, tt, ...stackTexts(m, y0, W * 0.36 - m, H - y0 - m), place(v, W * 0.4, y0, W * 0.6 - m, H - y0 - m)].filter(Boolean) });
    ideas.push({ id: 'vis-wide', name: 'Chart wide', elements: [...rest, tt, ...(body0 ? stackTexts(m, y0, W - 2 * m, H * 0.12) : []), place(v, m, body0 ? y0 + H * 0.14 : y0, W - 2 * m, H - (body0 ? y0 + H * 0.14 : y0) - m)].filter(Boolean) });
    const bar = { ...shapeEl('rect', 0, 0, W, H * 0.24), id: uid('d'), name: 'Designer band', fill: { type: 'solid', color: 'accent1' }, line: null, text: body() };
    ideas.push({ id: 'vis-band', name: 'Title band', elements: [bar, ...rest, title ? place(withText(title, { anchor: 'ctr', lvl: (title.text.lvl || [{}]).map((l) => ({ ...l, color: 'lt1' })) }), m, H * 0.03, W - 2 * m, H * 0.18) : null, place(v, m, H * 0.3, W - 2 * m, H * 0.64)].filter(Boolean) });
  }
  if (!pics.length && !body0 && !vis.length && title && bodyText(title.text).trim()) {
    ideas.push({ id: 'statement', name: 'Big statement', elements: [...others, place(withText(title, { anchor: 'ctr', lvl: (title.text.lvl || [{}]).map((l) => ({ ...l, align: 'ctr', size: 54 })) }), m, H * 0.25, W - 2 * m, H * 0.5)], background: { type: 'solid', color: 'accent1' } });
  }
  return ideas;
}

/* ---------------------------------------------------------------- compare versions */
const elKey = (e) => (e.ph ? `ph:${e.ph}` : e.name && !/^(shape|Shape|TextBox|Textfeld) \d+$/.test(e.name) ? `n:${e.name}` : e.text ? `t:${bodyText(e.text).slice(0, 40)}` : `${e.type}:${e.src ? e.src.length : e.geom || ''}`);
const elSig = (e) => JSON.stringify({ t: e.type, x: Math.round(e.x / 12700), y: Math.round(e.y / 12700), w: Math.round(e.w / 12700), h: Math.round(e.h / 12700), r: Math.round(e.rot || 0), text: e.text ? bodyText(e.text) : null, fill: e.fill || null, line: e.line || null, src: e.src ? e.src.length : 0, ch: e.chart || null, rows: e.rows ? e.rows.map((r) => r.cells.map((c) => bodyText(c.text))) : null, kids: e.children ? e.children.map(elSig) : null });
const slideText = (s) => { const out = []; walk(s.elements, (e) => { if (e.text) out.push(bodyText(e.text)); }); return out.join('\n'); };
function sim(a, b) {
  const ta = slideTitle(a); const tb = slideTitle(b);
  if (ta && ta === tb) return 1;
  const A = new Set(slideText(a).toLowerCase().split(/\W+/).filter(Boolean)); const B = new Set(slideText(b).toLowerCase().split(/\W+/).filter(Boolean));
  if (!A.size && !B.size) return 0.5;
  let n = 0; for (const x of A) if (B.has(x)) n++;
  return n / Math.max(A.size, B.size);
}
/**
 * slide-by-slide comparison of an older doc `a` with the current `b`:
 * [{ status: 'same'|'changed'|'added'|'removed'|'moved', a?: index, b?: index, changes: { added, removed, changed: [element ids of b / a] , notes } }]
 */
export function compareDecks(a, b) {
  const A = a.slides; const B = b.slides;
  // LCS on similarity ≥ 0.5 keeps order; leftovers matched by best similarity (moved)
  const n = A.length; const m = B.length;
  const ok = (i, j) => sim(A[i], B[j]) >= 0.5;
  const L = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = ok(i, j) ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
  const pairs = []; let i = 0; let j = 0;
  while (i < n && j < m) { if (ok(i, j) && L[i][j] === L[i + 1][j + 1] + 1) { pairs.push([i, j]); i++; j++; } else if (L[i + 1][j] >= L[i][j + 1]) i++; else j++; }
  const usedA = new Set(pairs.map((p) => p[0])); const usedB = new Set(pairs.map((p) => p[1]));
  const moved = [];
  for (let x = 0; x < n; x++) {
    if (usedA.has(x)) continue;
    let best = -1; let bs = 0.5;
    for (let y = 0; y < m; y++) if (!usedB.has(y)) { const s = sim(A[x], B[y]); if (s >= bs) { bs = s; best = y; } }
    if (best >= 0) { usedA.add(x); usedB.add(best); moved.push([x, best]); }
  }
  const diffEls = (sa, sb) => {
    const ea = sa.elements || []; const eb = sb.elements || [];
    const mapA = new Map(); for (const e of ea) mapA.set(elKey(e), e);
    const added = []; const changed = []; const seen = new Set();
    for (const e of eb) { const o = mapA.get(elKey(e)); if (!o) added.push(e.id); else { seen.add(o.id); if (elSig(o) !== elSig(e)) changed.push(e.id); } }
    const removed = ea.filter((e) => !seen.has(e.id)).map((e) => e.id);
    return { added, removed, changed, notes: (sa.notes || '') !== (sb.notes || ''), background: JSON.stringify(sa.background || null) !== JSON.stringify(sb.background || null), layout: false };
  };
  const rows = [];
  const all = [...pairs.map((p) => [...p, false]), ...moved.map((p) => [...p, true])];
  for (const [x, y, mv] of all) {
    const ch = diffEls(A[x], B[y]);
    const same = !ch.added.length && !ch.removed.length && !ch.changed.length && !ch.notes && !ch.background;
    rows.push({ status: mv ? 'moved' : same ? 'same' : 'changed', a: x, b: y, changes: ch });
  }
  for (let y = 0; y < m; y++) if (!usedB.has(y)) rows.push({ status: 'added', b: y, changes: null });
  for (let x = 0; x < n; x++) if (!usedA.has(x)) rows.push({ status: 'removed', a: x, changes: null });
  rows.sort((p, q) => (p.b ?? (p.a + 0.5)) - (q.b ?? (q.a + 0.5)));
  return rows;
}

/* ---------------------------------------------------------------- live fields */
/** every live run of a doc: [{ slideId, elId, key: `${id}|${params}`, live }] */
export function liveRuns(doc) {
  const out = [];
  for (const s of doc.slides) walk(s.elements, (e) => { if (e.text) for (const p of e.text.paras) for (const r of p.runs || []) if (r.live && r.live.id) out.push({ slideId: s.id, elId: e.id, key: `${r.live.id}|${JSON.stringify(r.live.params || {})}`, live: r.live }); });
  return out;
}
/** doc with the live runs' text replaced by values[key] (unchanged doc when nothing changed) */
export function applyLiveValues(doc, values) {
  let changedAny = false;
  const slides = doc.slides.map((s) => {
    let ch = false;
    const mapEl = (e) => {
      let n = e;
      if (e.text && e.text.paras.some((p) => (p.runs || []).some((r) => r.live))) {
        const paras = e.text.paras.map((p) => ({ ...p, runs: (p.runs || []).map((r) => { if (!r.live) return r; const v = values[`${r.live.id}|${JSON.stringify(r.live.params || {})}`]; if (v == null || String(v) === r.text) return r; ch = true; return { ...r, text: String(v) }; }) }));
        n = { ...e, text: { ...e.text, paras } };
      }
      if (e.children) n = { ...n, children: e.children.map(mapEl) };
      return n;
    };
    const elements = s.elements.map(mapEl);
    if (!ch) return s;
    changedAny = true;
    return { ...s, elements };
  });
  return changedAny ? { ...doc, slides } : doc;
}
/** text box with one live field */
export function liveFieldBox(doc, field, value) {
  const w = Math.round(doc.size.w * 0.3); const hh = Math.round(doc.size.h * 0.12);
  return textBox(Math.round((doc.size.w - w) / 2), Math.round((doc.size.h - hh) / 2), w, hh, body([{ ...para(''), runs: [run(value == null ? '…' : String(value), { live: { id: field.id, ...(field.params ? { params: field.params } : {}), label: field.label || field.id }, size: 28, b: true })] }], { anchor: 'ctr', autofit: 'none' }), { name: `Live: ${field.label || field.id}` });
}
export { imageEl };
