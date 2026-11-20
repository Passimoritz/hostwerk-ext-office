/*
 * docx / docm / dotx / dotm reader (own WordprocessingML parser – mammoth drops alignment, fonts, colours, comment
 * ranges and tracked changes, which we need for editing and for lossless saving).
 *
 * Keeps for lossless saving (model.keep.docx): the original package bytes, the root tag, the final sectPr, original
 * header/footer texts and footnote XML. Unknown paragraph/run properties travel along as attributes (`ox`) and are
 * written back; things we cannot edit (SmartArt, charts, shapes/text boxes, OLE objects, unknown fields, mc:Alternate-
 * Content) become `preserved` nodes holding their original XML plus a fallback (image or text) for display.
 */
import JSZip from 'jszip';
import { parseXml, kids, kid, all, text as xtext, serialize, local } from '../../_shared/xmltree.js';
import { defaultModel, sizeName, twipToMm, normalizeInline, fixDoc, dataUrl, mimeOf, uid, imageSize, hfKey } from './pm.js';
import { anchorOf } from './anchor.js';
import { paraPropsFromPPr, runPropsFromRPr, RUN_PROP_TAGS } from './wordprops.js';
import { specFromStyle, bordersFromEl, marginsFromEl, lookFromEl, builtinTableStyles } from './tables.js';

const W = (el, name) => kid(el, `w:${name}`);
const val = (el) => (el ? el.a['w:val'] : undefined);
const onOff = (el) => !!el && !['0', 'false', 'off'].includes(String(val(el) == null ? 'true' : val(el)));
const HL = { yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF' };
const PPR_MODELED = new Set(['w:pStyle', 'w:numPr', 'w:jc', 'w:spacing', 'w:ind', 'w:pageBreakBefore', 'w:outlineLvl', 'w:keepNext', 'w:keepLines', 'w:widowControl', 'w:pBdr', 'w:shd', 'w:tabs']);
const RPR_MODELED = new Set(['w:spacing', 'w:w', 'w:position', 'w:caps', 'w:smallCaps', 'w:vanish', 'w:lang', 'w:rStyle', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:u', 'w:strike', 'w:dstrike', 'w:vertAlign', 'w:sz', 'w:szCs', 'w:rFonts', 'w:color', 'w:highlight', 'w:shd', 'w:rPrChange', 'w:ins', 'w:del']);

/** Word fields we show as live fields (the rest stay preserved with their result) */
const FIELD_KINDS = { PAGE: 'page', NUMPAGES: 'pages', SECTIONPAGES: 'pages', DATE: 'date', TIME: 'time', CREATEDATE: 'date', SAVEDATE: 'date', PRINTDATE: 'date', AUTHOR: 'author', TITLE: 'title', FILENAME: 'filename', SUBJECT: 'subject', NUMWORDS: 'words' };
export function fieldKind(instr) { const s = String(instr || '').trim(); if (/^=\s*(SUM|AVERAGE|MIN|MAX|COUNT|PRODUCT)\s*\(/i.test(s)) return 'formula'; const w = s.split(/\s+/)[0].toUpperCase(); return FIELD_KINDS[w] || null; }
const firstMarks = (nodes) => { const n = (nodes || []).find((x) => x.type === 'text' && x.marks && x.marks.length); return n ? n.marks.filter((m) => !['comment', 'insertion', 'deletion', 'sdt'].includes(m.type)) : undefined; };
/** what kind of content control a w:sdtPr describes */
export function sdtInfo(pr) {
  const o = { kind: 'rich' };
  if (!pr) return o;
  o.tag = val(W(pr, 'tag')) || null; o.alias = val(W(pr, 'alias')) || null; o.placeholder = !!W(pr, 'showingPlcHdr');
  const cb = kid(pr, 'w14:checkbox');
  if (cb) { const c = kid(cb, 'w14:checked'); o.kind = 'checkbox'; o.checked = !!c && ['1', 'true'].includes(String(c.a['w14:val'] != null ? c.a['w14:val'] : val(c))); }
  else if (W(pr, 'dropDownList') || W(pr, 'comboBox')) { const l = W(pr, 'dropDownList') || W(pr, 'comboBox'); o.kind = l.n === 'w:comboBox' ? 'combo' : 'dropdown'; o.items = kids(l, 'w:listItem').map((li) => ({ text: li.a['w:displayText'] != null ? li.a['w:displayText'] : li.a['w:value'] || '', value: li.a['w:value'] != null ? li.a['w:value'] : li.a['w:displayText'] || '' })); }
  else if (W(pr, 'date')) { const d = W(pr, 'date'); o.kind = 'date'; o.fullDate = d.a['w:fullDate'] || null; o.format = val(W(d, 'dateFormat')) || null; }
  else if (W(pr, 'text')) o.kind = 'text';
  else if (W(pr, 'picture')) o.kind = 'picture';
  else if (W(pr, 'docPartObj') || W(pr, 'docPartList')) o.kind = 'docpart';
  else if (W(pr, 'group')) o.kind = 'group';
  if (o.tag && o.tag.startsWith('hostwerk-field:')) { try { const f = JSON.parse(o.tag.slice(15)); o.field = { kind: f.k === 'cell' ? 'cell' : 'hostwerk', instr: '', source: f.s || null, params: f.p || null, label: o.alias || f.l || '' }; } catch { /* a normal control */ } }
  return o;
}

async function readText(zip, path) { const f = zip.file(path); return f ? f.async('string') : null; }
function relsOf(xml, base) {
  const map = {};
  if (!xml) return map;
  const root = parseXml(xml);
  for (const r of kids(root, 'Relationship')) {
    const mode = r.a.TargetMode;
    let target = r.a.Target || '';
    if (mode !== 'External') target = resolvePath(base, target);
    map[r.a.Id] = { type: (r.a.Type || '').split('/').pop(), target, external: mode === 'External' };
  }
  return map;
}
export function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const p of target.split('/')) { if (p === '..') parts.pop(); else if (p && p !== '.') parts.push(p); }
  return parts.join('/');
}

export async function readDocx(bytes, settings, { ext = 'docx' } = {}) {
  const zip = await JSZip.loadAsync(bytes);
  const model = defaultModel(settings);
  model.fidelity = [];
  const fid = new Set();
  const rootRels = relsOf(await readText(zip, '_rels/.rels'), '');
  const docPath = (Object.values(rootRels).find((r) => r.type === 'officeDocument') || { target: 'word/document.xml' }).target;
  const docXml = await readText(zip, docPath);
  if (!docXml) throw new Error('This file has no Word document inside.');
  const relsPath = docPath.replace(/([^/]+)$/, '_rels/$1.rels');
  let rels = relsOf(await readText(zip, relsPath), docPath);
  const docRels = rels;
  const relOfType = (t) => Object.values(rels).find((r) => r.type === t);

  // theme fonts
  const themeRel = relOfType('theme');
  const theme = { minor: 'Calibri', major: 'Calibri Light' };
  if (themeRel) {
    const tx = await readText(zip, themeRel.target);
    if (tx) { const mi = /<a:minorFont>\s*<a:latin typeface="([^"]*)"/.exec(tx); const ma = /<a:majorFont>\s*<a:latin typeface="([^"]*)"/.exec(tx); if (mi) theme.minor = mi[1]; if (ma) theme.major = ma[1]; }
  }
  const fontOf = (rf) => {
    if (!rf) return null;
    const f = rf.a['w:ascii'] || rf.a['w:hAnsi'];
    if (f) return f;
    const t = rf.a['w:asciiTheme'] || rf.a['w:hAnsiTheme'];
    if (t) return /major/i.test(t) ? theme.major : theme.minor;
    return null;
  };

  // styles
  const styles = {}; const styleByName = {};
  let docLang = null; // the document's language (runs in it carry no language of their own here)
  const stylesRel = relOfType('styles');
  const stylesXml = stylesRel ? await readText(zip, stylesRel.target) : null;
  if (stylesXml) {
    const sr = parseXml(stylesXml);
    const dd = W(sr, 'docDefaults');
    const drpr = dd && kid(kid(dd, 'w:rPrDefault'), 'w:rPr');
    const dppr = dd && kid(kid(dd, 'w:pPrDefault'), 'w:pPr');
    if (drpr && W(drpr, 'lang')) docLang = W(drpr, 'lang').a['w:val'] || null;
    if (drpr) { const f = fontOf(W(drpr, 'rFonts')); if (f) model.defaults.font = f; const sz = W(drpr, 'sz'); if (sz) model.defaults.fontSize = Number(val(sz)) / 2; }
    if (dppr) { const sp = W(dppr, 'spacing'); if (sp) { if (sp.a['w:after'] != null) model.defaults.spaceAfter = Number(sp.a['w:after']) / 20; if (sp.a['w:line'] && (sp.a['w:lineRule'] || 'auto') === 'auto') model.defaults.lineHeight = Math.round((Number(sp.a['w:line']) / 240) * 100) / 100; } } else { model.defaults.spaceAfter = 0; model.defaults.lineHeight = 1; }
    for (const s of kids(sr, 'w:style')) {
      const id = s.a['w:styleId'];
      const name = (val(W(s, 'name')) || id || '').toLowerCase();
      const st = { id, name, type: s.a['w:type'], basedOn: val(W(s, 'basedOn')), pPr: W(s, 'pPr'), rPr: W(s, 'rPr'), def: s.a['w:default'] === '1', el: s };
      styles[id] = st; styleByName[name] = id;
      if (st.def && st.type === 'paragraph') {
        // "Normal" may carry the real default font/size/spacing
        const rp = st.rPr; if (rp) { const f = fontOf(W(rp, 'rFonts')); if (f) model.defaults.font = f; const sz = W(rp, 'sz'); if (sz) model.defaults.fontSize = Number(val(sz)) / 2; }
        const pp = st.pPr; const sp = pp && W(pp, 'spacing'); if (sp) { if (sp.a['w:after'] != null) model.defaults.spaceAfter = Number(sp.a['w:after']) / 20; if (sp.a['w:line'] && (sp.a['w:lineRule'] || 'auto') === 'auto') model.defaults.lineHeight = Math.round((Number(sp.a['w:line']) / 240) * 100) / 100; }
      }
    }
  }
  // table styles → display spec (borders, cell margins, header row / banding colours) – resolved along basedOn
  const tspecs = {};
  const tableSpec = (id) => {
    if (!id) return null;
    if (tspecs[id] !== undefined) return tspecs[id];
    tspecs[id] = null;
    const st = styles[id];
    if (st && st.type === 'table') tspecs[id] = specFromStyle(st.el, { kid, kids }, st.basedOn ? tableSpec(st.basedOn) : null);
    else { const b = builtinTableStyles()[id]; if (b) tspecs[id] = b.spec; }
    return tspecs[id];
  };
  const defaultTableStyle = Object.values(styles).find((x) => x.type === 'table' && x.def);
  const styleChain = (id) => { const out = []; let s = styles[id]; let guard = 0; while (s && guard++ < 20) { out.unshift(s); s = styles[s.basedOn]; } return out; };
  const headingOfStyle = (id) => {
    for (const s of styleChain(id).reverse()) {
      const m = /^(?:heading|überschrift|titre|titolo|título|kop)\s*(\d)$/i.exec(s.name);
      if (m) return Number(m[1]);
      const ol = s.pPr && W(s.pPr, 'outlineLvl');
      if (ol && Number(val(ol)) < 6 && !/toc|verzeichnis/i.test(s.name)) return Number(val(ol)) + 1;
    }
    return 0;
  };
  const roleOfStyle = (id) => {
    const s = styles[id]; if (!s) return null;
    const n = s.name;
    if (n === 'title' || n === 'titel') return 'title';
    if (n === 'subtitle' || n === 'untertitel') return 'subtitle';
    if (n === 'quote' || n === 'intense quote' || n === 'zitat' || n === 'intensives zitat') return 'quote';
    if (n === 'list paragraph' || n === 'listenabsatz') return 'listparagraph';
    if (/^(html preformatted|code|source code|html vorformatiert|quelltext)$/.test(n)) return 'code';
    if (/^toc \d|^verzeichnis \d|^inhaltsverzeichnis \d/.test(n)) return 'toc';
    if (n === 'normal' || n === 'standard' || s.def) return 'normal';
    return null;
  };

  // numbering
  const nums = {}; const abstracts = {};
  const numRel = relOfType('numbering');
  const numXml = numRel ? await readText(zip, numRel.target) : null;
  if (numXml) {
    const nr = parseXml(numXml);
    for (const a of kids(nr, 'w:abstractNum')) {
      const levels = {};
      for (const l of kids(a, 'w:lvl')) levels[l.a['w:ilvl']] = { fmt: val(W(l, 'numFmt')) || 'decimal', text: val(W(l, 'lvlText')) || '', start: Number(val(W(l, 'start')) || 1) };
      abstracts[a.a['w:abstractNumId']] = levels;
    }
    for (const n of kids(nr, 'w:num')) {
      const over = {};
      for (const o of kids(n, 'w:lvlOverride')) { const so = W(o, 'startOverride'); if (so) over[o.a['w:ilvl']] = Number(val(so)); }
      nums[n.a['w:numId']] = { abs: val(W(n, 'abstractNumId')), over };
    }
  }
  const listKind = (numId, ilvl) => {
    const n = nums[numId]; if (!n) return null;
    const l = (abstracts[n.abs] || {})[ilvl] || (abstracts[n.abs] || {})[0];
    if (!l || l.fmt === 'none') return null;
    if (l.fmt === 'bullet') return { kind: /[☐☑☒□■]/.test(l.text) ? 'task' : 'bullet', start: 1 };
    const ls = { lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman' }[l.fmt] || (/%\d\.%\d/.test(l.text) ? 'outline' : null);
    return { kind: 'ordered', listStyle: ls, start: n.over[ilvl] != null ? n.over[ilvl] : l.start };
  };

  // footnotes / endnotes
  const notes = {};
  for (const [type, tagName] of [['footnotes', 'w:footnote'], ['endnotes', 'w:endnote']]) {
    const r = relOfType(type);
    const x = r ? await readText(zip, r.target) : null;
    if (!x) continue;
    for (const f of kids(parseXml(x), tagName)) {
      if (f.a['w:type'] && f.a['w:type'] !== 'normal') continue;
      const textOf = kids(f, 'w:p').map((p) => paraPlain(p)).join(' ').replace(/\s+/g, ' ').trim();
      notes[`${type}:${f.a['w:id']}`] = { text: textOf, xml: kids(f, '*').map(serialize).join('') };
      if (type === 'endnotes') fid.add('endnotes');
    }
  }

  // comments (+ replies and "done" from commentsExtended)
  const comments = {}; const paraToComment = {};
  const cRel = relOfType('comments');
  const cXml = cRel ? await readText(zip, cRel.target) : null;
  if (cXml) {
    for (const c of kids(parseXml(cXml), 'w:comment')) {
      const ps = kids(c, 'w:p');
      const id = c.a['w:id'];
      comments[id] = { id: `c${id}`, author: c.a['w:author'] || '', initials: c.a['w:initials'] || '', date: c.a['w:date'] || null, text: ps.map(paraPlain).join('\n').trim(), resolved: false, replies: [], parent: null };
      const last = ps[ps.length - 1];
      if (last && last.a['w14:paraId']) paraToComment[last.a['w14:paraId']] = id;
    }
    const exRel = relOfType('commentsExtended');
    const ex = exRel ? await readText(zip, exRel.target) : null;
    if (ex) {
      for (const e of all(parseXml(ex), 'w15:commentEx')) {
        const id = paraToComment[e.a['w15:paraId']];
        if (id == null) continue;
        if (e.a['w15:done'] === '1') comments[id].resolved = true;
        const parent = e.a['w15:paraIdParent'] && paraToComment[e.a['w15:paraIdParent']];
        if (parent != null && comments[parent]) comments[id].parent = parent;
      }
    }
    // durable ids (Word 2016+): kept so a thread stays the same thread after saving here
    const idRel = Object.values(rels).find((r) => r.type === 'commentsIds');
    const idXml = idRel ? await readText(zip, idRel.target) : null;
    if (idXml) for (const e of all(parseXml(idXml), 'w16cid:commentId')) { const id = paraToComment[e.a['w16cid:paraId']]; if (id != null && comments[id]) comments[id].durableId = e.a['w16cid:durableId'] || null; }
    for (const c of Object.values(comments)) if (c.parent != null) comments[c.parent].replies.push({ author: c.author, date: c.date, text: c.text, durableId: c.durableId || null });
    model.comments = Object.entries(comments).filter(([, c]) => c.parent == null).map(([, c]) => ({ id: c.id, author: c.author, date: c.date, text: c.text, resolved: c.resolved, replies: c.replies, durableId: c.durableId || null }));
  }

  // media cache
  const media = {};
  const mediaUrl = async (rid) => {
    const r = rels[rid];
    if (!r) return null;
    if (r.external) return r.target;
    if (media[r.target] !== undefined) return media[r.target];
    const f = zip.file(r.target);
    if (!f) { media[r.target] = null; return null; }
    const mime = mimeOf(r.target);
    const u8 = await f.async('uint8array');
    media[r.target] = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/svg+xml', 'image/webp'].includes(mime) ? dataUrl(u8, mime) : null;
    return media[r.target];
  };

  /* ---------------- body ---------------- */
  const root = parseXml(docXml);
  const body = W(root, 'body');
  const bg = W(root, 'background');
  if (bg && bg.a['w:color'] && bg.a['w:color'] !== 'auto') model.page.color = `#${bg.a['w:color']}`;
  const rootTag = /<w:document\b[^>]*>/.exec(docXml);
  let openComments = new Set();
  const commentIds = new Set(Object.keys(comments).filter((k) => comments[k].parent == null));

  const runMarks = (rPr, extra = []) => {
    const m = [];
    if (!rPr) return extra.slice();
    const rs = val(W(rPr, 'rStyle'));
    let base = {};
    if (rs && styles[rs]) for (const s of styleChain(rs)) if (s.rPr) base = { ...base, ...plainRun(s.rPr) };
    const own = plainRun(rPr);
    const r = { ...base, ...own };
    if (r.b) m.push({ type: 'bold' });
    if (r.i) m.push({ type: 'italic' });
    if (r.u) m.push({ type: 'underline' });
    if (r.strike) m.push({ type: 'strike' });
    if (r.sup) m.push({ type: 'superscript' });
    if (r.sub) m.push({ type: 'subscript' });
    const xp = { ...(rs && styles[rs] ? styleChain(rs).reduce((o, st) => ({ ...o, ...runPropsFromRPr(st.rPr, { kid, kids }) }), {}) : {}), ...runPropsFromRPr(rPr, { kid, kids }) };
    if (r.font || r.size || r.color || Object.keys(xp).length) m.push({ type: 'textStyle', attrs: { fontFamily: r.font || null, fontSize: r.size ? `${r.size}pt` : null, color: r.color || null, ...xp } });
    if (r.hl) m.push({ type: 'highlight', attrs: { color: r.hl } });
    const ch = W(rPr, 'rPrChange');
    if (ch) m.push({ type: 'formatChange', attrs: { id: ch.a['w:id'] || uid('f'), author: ch.a['w:author'] || '', date: ch.a['w:date'] || null, old: kid(ch, 'w:rPr') ? serialize(kid(ch, 'w:rPr')) : '<w:rPr/>' } });
    const keep = kids(rPr, '*').filter((x) => !RPR_MODELED.has(x.n)).map((x) => [x.n, serialize(x)]);
    if (keep.length) m.push({ type: 'runKeep', attrs: { ox: keep } });
    if (rs && !/hyperlink|footnote|endnote|comment/i.test(rs)) m.push({ type: 'runKeep', attrs: { ox: [...keep, ['w:rStyle', serialize(W(rPr, 'rStyle'))]] } });
    // only one runKeep mark
    const rk = m.filter((x) => x.type === 'runKeep');
    if (rk.length > 1) { const lastRk = rk[rk.length - 1]; return [...m.filter((x) => x.type !== 'runKeep'), lastRk, ...extra]; }
    return [...m, ...extra];
  };
  function plainRun(rPr) {
    const r = {};
    if (W(rPr, 'b')) r.b = onOff(W(rPr, 'b'));
    if (W(rPr, 'i')) r.i = onOff(W(rPr, 'i'));
    const u = W(rPr, 'u'); if (u) r.u = val(u) !== 'none' && val(u) !== '0';
    if (W(rPr, 'strike') || W(rPr, 'dstrike')) r.strike = onOff(W(rPr, 'strike') || W(rPr, 'dstrike'));
    const va = val(W(rPr, 'vertAlign')); if (va) { r.sup = va === 'superscript'; r.sub = va === 'subscript'; }
    const sz = W(rPr, 'sz'); if (sz) r.size = Number(val(sz)) / 2;
    const f = fontOf(W(rPr, 'rFonts')); if (f) r.font = f;
    const c = val(W(rPr, 'color')); if (c && c !== 'auto') r.color = `#${c}`;
    const hl = val(W(rPr, 'highlight')); if (hl && hl !== 'none') r.hl = HL[hl] || '#FFFF00';
    const shd = W(rPr, 'shd'); if (shd && shd.a['w:fill'] && !['auto', 'FFFFFF'].includes(shd.a['w:fill']) && !r.hl) r.hl = `#${shd.a['w:fill']}`;
    return r;
  }
  const cleanMarks = (marks) => {
    // drop default-ish values (font equal to document default) so the editor stays clean
    return marks.map((m) => {
      if (m.type !== 'textStyle') return m;
      const a = { ...m.attrs };
      if (a.fontFamily && a.fontFamily === model.defaults.font) a.fontFamily = null;
      if (a.fontSize && parseFloat(a.fontSize) === model.defaults.fontSize) a.fontSize = null;
      if (a.lang && a.lang === docLang) delete a.lang;
      return Object.values(a).some((v) => v != null && v !== false) ? { type: 'textStyle', attrs: a } : null;
    }).filter(Boolean);
  };

  const commentMarks = () => [...openComments].filter((id) => commentIds.has(id)).map((id) => ({ type: 'comment', attrs: { id: `c${id}` } }));

  /** inline content of a paragraph-like element → [inline nodes] ; `ctx.marks` = marks from wrappers (ins/del/link) */
  async function inlines(el, ctx, out) {
    for (const c of el.c || []) {
      if (typeof c === 'string') continue;
      switch (c.n) {
        case 'w:r': await run(c, ctx, out); break;
        case 'w:hyperlink': {
          const rid = c.a['r:id'];
          const href = rid && rels[rid] ? rels[rid].target : c.a['w:anchor'] ? `#${c.a['w:anchor']}` : null;
          await inlines(c, { ...ctx, marks: [...ctx.marks, ...(href ? [{ type: 'link', attrs: { href } }] : [])] }, out);
          break;
        }
        case 'w:ins': case 'w:moveTo':
          await inlines(c, { ...ctx, marks: [...ctx.marks, { type: 'insertion', attrs: { id: c.a['w:id'] || uid('i'), author: c.a['w:author'] || '', date: c.a['w:date'] || null, move: c.n === 'w:moveTo' } }] }, out);
          break;
        case 'w:del': case 'w:moveFrom':
          await inlines(c, { ...ctx, marks: [...ctx.marks, { type: 'deletion', attrs: { id: c.a['w:id'] || uid('d'), author: c.a['w:author'] || '', date: c.a['w:date'] || null, move: c.n === 'w:moveFrom' } }] }, out);
          break;
        case 'w:commentRangeStart': openComments.add(c.a['w:id']); break;
        case 'w:commentRangeEnd': openComments.delete(c.a['w:id']); break;
        case 'w:smartTag': case 'w:customXml': case 'w:dir': case 'w:bdo': await inlines(c, ctx, out); break;
        case 'w:sdt': {
          const pr = W(c, 'sdtPr');
          const info = sdtInfo(pr);
          const content = W(c, 'sdtContent') || { c: [] };
          const prXml = pr ? serialize(pr) : '';
          const endXml = W(c, 'sdtEndPr') ? serialize(W(c, 'sdtEndPr')) : '';
          // a check box at the start of a paragraph = a check list item (what our writer and Word's lists do)
          if (info.kind === 'checkbox' && !out.some((x) => x.type !== 'text' || /\S/.test(x.text || '')) && !ctx.checkbox && !ctx.noTask) { ctx.checkbox = { checked: info.checked }; break; }
          if (info.field) {
            const inner = []; await inlines(content, ctx, inner);
            out.push({ type: 'field', attrs: { ...info.field, value: inner.map((x) => x.text || '').join(''), ox: serialize(c) }, marks: firstMarks(inner) });
            break;
          }
          if (['checkbox', 'dropdown', 'combo', 'date'].includes(info.kind)) {
            const inner = []; await inlines(content, ctx, inner);
            const value = inner.map((x) => x.text || '').join('');
            out.push({ type: 'contentControl', attrs: { kind: info.kind, value, checked: !!info.checked, items: info.items || null, format: info.format || null, fullDate: info.fullDate || null, tag: info.tag || null, alias: info.alias || null, placeholder: !!info.placeholder, pr: prXml, end: endXml }, marks: firstMarks(inner) });
            break;
          }
          // plain/rich text controls: the text stays normal, editable text with an `sdt` mark – written back as the same control
          const sid = uid('s');
          const before = out.length;
          await inlines(content, { ...ctx, marks: [...ctx.marks, { type: 'sdt', attrs: { id: sid, kind: info.kind, alias: info.alias || null, tag: info.tag || null, pr: prXml, end: endXml } }] }, out);
          if (out.length === before || !out.slice(before).some((x) => x.type === 'text')) {
            if (!out.slice(before).length) out.push({ type: 'contentControl', attrs: { kind: 'text', value: '', checked: false, items: null, format: null, fullDate: null, tag: info.tag || null, alias: info.alias || null, placeholder: true, pr: prXml, end: endXml } });
            else fid.add('contentControls');
          }
          break;
        }
        case 'w:fldSimple': {
          const instr = (c.a['w:instr'] || '').trim();
          const hm = /^HYPERLINK\s+"([^"]+)"/i.exec(instr);
          if (hm) { await inlines(c, { ...ctx, marks: [...ctx.marks, { type: 'link', attrs: { href: hm[1] } }] }, out); break; }
          const inner = []; await inlines(c, ctx, inner);
          const resultText = inner.map((x) => x.text || '').join('') || paraPlain(c);
          const fk = fieldKind(instr);
          if (fk) { out.push({ type: 'field', attrs: { kind: fk, instr, value: resultText, ox: serialize(c) }, marks: firstMarks(inner) }); break; }
          out.push({ type: 'preserved', attrs: { format: 'docx', kind: 'field', label: instr.split(/\s+/)[0] || 'Field', text: resultText, src: null, xml: serialize(c) } });
          break;
        }
        case 'm:oMathPara': case 'm:oMath': { const latex = ommlToLatex(c).replace(/\s+/g, ' ').trim(); out.push({ type: 'equation', attrs: { latex, display: c.n === 'm:oMathPara', omml: serialize(c), ommlLatex: latex } }); break; }
        case 'w:bookmarkStart': case 'w:bookmarkEnd': case 'w:proofErr': case 'w:permStart': case 'w:permEnd': case 'w:pPr': break;
        case 'mc:AlternateContent': out.push(await preservedOf({ n: 'w:r', a: {}, c: [c] }, c)); break;
        default: break;
      }
    }
  }

  /** a run: text, tabs, breaks, notes, drawings, fields */
  async function run(r, ctx, out) {
    const rPr = W(r, 'rPr');
    // complex fields are collected per paragraph in ctx.field
    const fc = W(r, 'fldChar');
    if (fc) {
      const t = fc.a['w:fldCharType'];
      if (t === 'begin') { ctx.fields.push({ instr: '', phase: 'instr', xml: [serialize(r)], start: out.length }); return; }
      const f = ctx.fields[ctx.fields.length - 1];
      if (f) {
        f.xml.push(serialize(r));
        if (t === 'separate') { f.phase = 'result'; f.resultStart = out.length; if (/^TOC\b/i.test(f.instr.trim())) ctx.toc = true; return; }
        if (t === 'end') {
          ctx.fields.pop();
          const instr = f.instr.trim();
          const hm = /^HYPERLINK\s+(?:\\l\s+)?"([^"]+)"/i.exec(instr);
          const result = out.splice(f.resultStart != null ? f.resultStart : out.length);
          out.splice(f.start);
          if (hm) { for (const n of result) if (n.type === 'text') n.marks = [...(n.marks || []).filter((m) => m.type !== 'link'), { type: 'link', attrs: { href: /^\\l/.test(instr.slice(9).trim()) ? `#${hm[1]}` : hm[1] } }]; out.push(...result); return; }
          if (/^TOC\b/i.test(instr)) { ctx.toc = true; out.push(...result); return; }
          if (ctx.fields.length) { out.push(...result); return; }
          const textOf = result.map((n) => n.text || '').join('');
          const fk = fieldKind(instr);
          if (fk && result.every((n) => n.type === 'text')) { out.push({ type: 'field', attrs: { kind: fk, instr, value: textOf, ox: f.xml.join('') }, marks: firstMarks(result) }); return; }
          out.push({ type: 'preserved', attrs: { format: 'docx', kind: 'field', label: instr.split(/\s+/)[0] || 'Field', text: textOf, src: null, xml: f.xml.join('') } });
          return;
        }
      }
      return;
    }
    const field = ctx.fields[ctx.fields.length - 1];
    if (field && field.phase === 'instr') { field.xml.push(serialize(r)); for (const it of kids(r, 'w:instrText')) field.instr += xtext(it); return; }
    if (field) field.xml.push(serialize(r));
    const marks = cleanMarks(runMarks(rPr, [...ctx.marks, ...commentMarks()]));
    const push = (t) => { if (t) out.push({ type: 'text', text: t, marks }); };
    for (const c of r.c || []) {
      if (typeof c === 'string') continue;
      switch (c.n) {
        case 'w:t': case 'w:delText': push(xtext(c)); break;
        case 'w:tab': case 'w:ptab': push('\t'); break;
        case 'w:br': {
          const t = c.a['w:type'];
          if (t === 'page') out.push({ type: '__pagebreak' }); else if (t === 'column') out.push({ type: 'hardBreak' }); else out.push({ type: 'hardBreak' });
          break;
        }
        case 'w:cr': out.push({ type: 'hardBreak' }); break;
        case 'w:noBreakHyphen': push('‑'); break;
        case 'w:softHyphen': push('­'); break;
        case 'w:sym': { const code = parseInt(c.a['w:char'] || '0', 16); push(String.fromCharCode(code >= 0xf000 ? code - 0xf000 : code)); break; }
        case 'w:footnoteReference': case 'w:endnoteReference': {
          const n = notes[`${c.n === 'w:footnoteReference' ? 'footnotes' : 'endnotes'}:${c.a['w:id']}`];
          out.push({ type: 'footnote', attrs: { text: n ? n.text : '', ox: n && c.n === 'w:footnoteReference' ? { xml: n.xml, text: n.text } : null } });
          break;
        }
        case 'w:drawing': out.push(await drawing(c, r)); break;
        case 'w:pict': out.push(await vml(c, r)); break;
        case 'w:object': out.push(await preservedOf(r, c)); break;
        case 'mc:AlternateContent': out.push(await preservedOf(r, c)); break;
        case 'w:fldChar': case 'w:instrText': case 'w:rPr': case 'w:lastRenderedPageBreak': case 'w:commentReference': case 'w:annotationRef': case 'w:separator': case 'w:continuationSeparator': break;
        default: break;
      }
    }
  }

  async function drawing(dr, r) {
    const holder = kid(dr, 'wp:inline') || kid(dr, 'wp:anchor');
    const gd = holder && all(holder, 'a:graphicData')[0];
    const uri = (gd && gd.a.uri) || '';
    if (holder && /picture/.test(uri)) {
      const blip = all(gd, 'a:blip')[0];
      const rid = blip && (blip.a['r:embed'] || blip.a['r:link']);
      const src = rid ? await mediaUrl(rid) : null;
      const ext = kid(holder, 'wp:extent');
      const width = ext ? Math.round(Number(ext.a.cx) / 9525) : null;
      const height = ext ? Math.round(Number(ext.a.cy) / 9525) : null;
      const pr = kid(holder, 'wp:docPr');
      let align = null; let wrap = null;
      if (holder.n === 'wp:anchor') {
        const ph = kid(holder, 'wp:positionH');
        const al = ph && kid(ph, 'wp:align');
        align = al ? xtext(al) : 'left';
        if (kid(holder, 'wp:wrapSquare') || kid(holder, 'wp:wrapTight') || kid(holder, 'wp:wrapThrough')) wrap = 'square';
        else if (kid(holder, 'wp:wrapTopAndBottom')) { wrap = null; align = align === 'right' ? 'right' : 'center'; }
        if (align !== 'left' && align !== 'right' && align !== 'center') align = 'left';
      }
      // a diagram made here: the picture is its rendering, the alt text its source (round trip through Word/Google)
      if (src && pr && /^hostwerk-diagram/.test(pr.a.title || '')) return { type: 'diagram', attrs: { kind: (pr.a.title.split(':')[1] || 'mermaid'), source: pr.a.descr || '', src, svg: null, width, height } };
      if (src) {
        return { type: 'image', attrs: { src, alt: (pr && pr.a.descr) || '', title: (pr && pr.a.title) || null, width, height, align, wrap, ox: { xml: serialize(r), width, src: src.length } } };
      }
      if (rid) fid.add('images');
    }
    return preservedOf(r, dr, uri);
  }
  async function vml(pict, r) {
    const im = all(pict, 'v:imagedata')[0];
    if (im && !all(pict, 'v:textbox').length) {
      const src = await mediaUrl(im.a['r:id']);
      if (src) return { type: 'image', attrs: { src, alt: im.a['o:title'] || '', title: null, width: null, height: null, align: null, wrap: null, ox: { xml: serialize(r), width: null, src: src.length } } };
    }
    return preservedOf(r, pict);
  }
  /** something we keep 1:1: find a label, a fallback picture and its text */
  async function preservedOf(r, el, uri = '') {
    let kind = 'object'; let label = 'Object';
    const s = serialize(el);
    if (/diagram/.test(uri) || /dgm:relIds/.test(s)) { kind = 'smartart'; label = 'SmartArt'; } else if (/chart/.test(uri) || /c:chart/.test(s)) { kind = 'chart'; label = 'Chart'; } else if (/wps:txbx|v:textbox|w:txbxContent/.test(s)) { kind = 'textbox'; label = 'Text box'; } else if (/wpg:wgp|wps:wsp|v:shape|v:rect|v:oval|wpc:/.test(s)) { kind = 'shape'; label = 'Shape'; } else if (el.n === 'w:object' || /o:OLEObject/.test(s)) { kind = 'ole'; label = 'Embedded object'; }
    if (/o:OLEObject[^>]*ProgID="([^"]+)"/.test(s)) label = `${label} (${/ProgID="([^"]+)"/.exec(s)[1]})`;
    fid.add(kind);
    // fallback picture: blip in the fallback, VML imagedata, or the chart/diagram title text
    let src = null;
    const blips = all(el, 'a:blip');
    const imd = all(el, 'v:imagedata');
    for (const b of [...blips].reverse()) { const id = b.a['r:embed']; if (id) { src = await mediaUrl(id); if (src) break; } }
    if (!src) for (const im of imd) { src = await mediaUrl(im.a['r:id']); if (src) break; }
    let txt = '';
    for (const tb of all(el, 'w:txbxContent')) { txt = kids(tb, 'w:p').map(paraPlain).join('\n'); if (txt) break; }
    if (!txt && kind === 'smartart') {
      const relIds = all(el, 'dgm:relIds')[0];
      const dm = relIds && rels[relIds.a['r:dm']];
      const dx = dm ? await readText(zip, dm.target) : null;
      if (dx) txt = all(parseXml(dx), 'a:t').map(xtext).filter(Boolean).join(' · ');
    }
    if (!txt && kind === 'chart') {
      const c = all(el, 'c:chart')[0];
      const cr = c && rels[c.a['r:id']];
      const cx = cr ? await readText(zip, cr.target) : null;
      if (cx) { const t = all(parseXml(cx), 'c:title')[0]; txt = t ? all(t, 'a:t').map(xtext).join('') : ''; }
    }
    const ext = all(el, 'wp:extent')[0];
    const width = ext ? Math.round(Number(ext.a.cx) / 9525) : null;
    const xml = serialize(r);
    // text boxes and shapes anchored in the text (wp:anchor): shown floating at their position, movable (io/anchor.js)
    const float = kind === 'textbox' || kind === 'shape' ? anchorOf(xml) : null;
    return { type: 'preserved', attrs: { format: 'docx', kind, label, text: txt, src, width, xml, float } };
  }

  /** paragraph → block nodes (usually one; page breaks inside split it) */
  async function paragraph(p, inherited = {}) {
    const pPr = W(p, 'pPr');
    const sid = val(pPr && W(pPr, 'pStyle'));
    const role = roleOfStyle(sid);
    let level = sid ? headingOfStyle(sid) : 0;
    const ol = pPr && W(pPr, 'outlineLvl');
    if (!level && ol && Number(val(ol)) < 6 && role !== 'toc') level = Number(val(ol)) + 1;
    const attrs = {};
    // style formatting for display (custom styles)
    let sp = {};
    for (const s of styleChain(sid)) if (s.pPr) sp = { ...sp, ...plainPara(s.pPr) };
    const own = pPr ? plainPara(pPr) : {};
    const pa = { ...(level || role === 'title' || role === 'subtitle' ? {} : sp), ...own };
    if (pa.align) attrs.textAlign = pa.align;
    if (pa.lineHeight) attrs.lineHeight = pa.lineHeight;
    if (pa.before != null) attrs.spaceBefore = pa.before;
    if (pa.after != null) attrs.spaceAfter = pa.after;
    if (pa.indent) attrs.indent = pa.indent;
    if (pa.firstLine) attrs.firstLine = pa.firstLine;
    // keep with next, keep lines, widow control, borders/shading, tabs, right indent, exact line spacing, drop cap
    const xp = { ...(level ? {} : styleChain(sid).reduce((o, st) => ({ ...o, ...(st.pPr ? paraPropsFromPPr(st.pPr, { kid, kids }) : {}) }), {})), ...paraPropsFromPPr(pPr, { kid, kids }) };
    delete xp.outline; if (!level && pPr && W(pPr, 'outlineLvl') && Number(val(W(pPr, 'outlineLvl'))) >= 6) xp.outline = Number(val(W(pPr, 'outlineLvl')));
    for (const [k, v] of Object.entries(xp)) if (v != null) attrs[k] = v;
    if (role === 'title' || role === 'subtitle' || role === 'listparagraph') attrs.pstyle = role;
    if (sid && !level && !['title', 'subtitle', 'listparagraph', 'normal', 'quote', 'code', 'toc'].includes(role)) attrs.styleId = sid;
    const keep = [];
    let section = null;
    for (const x of pPr ? kids(pPr, '*') : []) {
      if (PPR_MODELED.has(x.n)) continue;
      if (x.n === 'w:framePr' && attrs.dropCap) continue;
      if (x.n === 'w:sectPr') { section = x; continue; }
      if (x.n === 'w:pPrChange') { const op = kid(x, 'w:pPr'); attrs.pchange = { id: x.a['w:id'] || uid('p'), author: x.a['w:author'] || '', date: x.a['w:date'] || null, old: op ? serialize(op) : '<w:pPr/>' }; continue; }
      if (x.n === 'w:rPr') {
        // the paragraph mark itself was inserted / deleted (Enter / joining paragraphs with track changes on)
        const mi = W(x, 'ins') || W(x, 'moveTo'); const md = W(x, 'del') || W(x, 'moveFrom');
        const rev = mi || md;
        if (rev) attrs.pmark = { type: mi ? 'ins' : 'del', id: rev.a['w:id'] || uid('m'), author: rev.a['w:author'] || '', date: rev.a['w:date'] || null };
        const rest = kids(x, '*').filter((y) => !['w:ins', 'w:del', 'w:moveTo', 'w:moveFrom'].includes(y.n));
        if (rest.length) keep.push(['w:rPr', `<w:rPr>${rest.map(serialize).join('')}</w:rPr>`]);
        continue;
      }
      keep.push([x.n, serialize(x)]);
    }
    if (keep.length) attrs.ox = { ppr: keep };
    // style-level run formatting of custom styles shows as direct formatting
    let styleRun = {};
    if (sid && !level && role !== 'title' && role !== 'subtitle') for (const s of styleChain(sid)) if (s.rPr && !s.def) styleRun = { ...styleRun, ...plainRun(s.rPr) };
    const ctx = { marks: [], fields: inherited.fields || [], checkbox: null, toc: false };
    const out = [];
    await inlines(p, ctx, out);
    if (Object.keys(styleRun).length) {
      const sm = [];
      if (styleRun.b) sm.push({ type: 'bold' }); if (styleRun.i) sm.push({ type: 'italic' }); if (styleRun.u) sm.push({ type: 'underline' });
      if (styleRun.color || (styleRun.size && styleRun.size !== model.defaults.fontSize) || (styleRun.font && styleRun.font !== model.defaults.font)) sm.push({ type: 'textStyle', attrs: { color: styleRun.color || null, fontSize: styleRun.size && styleRun.size !== model.defaults.fontSize ? `${styleRun.size}pt` : null, fontFamily: styleRun.font && styleRun.font !== model.defaults.font ? styleRun.font : null } });
      for (const n of out) if (n.type === 'text') { const have = new Set((n.marks || []).map((m) => m.type)); n.marks = [...sm.filter((m) => !have.has(m.type)), ...(n.marks || [])]; }
    }
    const numPr = pPr && W(pPr, 'numPr');
    let list = null;
    let numId = numPr && val(W(numPr, 'numId'));
    let ilvl = numPr ? Number(val(W(numPr, 'ilvl')) || 0) : 0;
    if (!numPr && sid) { for (const s of styleChain(sid).reverse()) { const np = s.pPr && W(s.pPr, 'numPr'); if (np) { numId = val(W(np, 'numId')); ilvl = Number(val(W(np, 'ilvl')) || 0); break; } } }
    if (numId && numId !== '0') list = listKind(numId, ilvl);
    if (ctx.checkbox) { list = { kind: 'task', checked: ctx.checkbox.checked }; if (out[0] && out[0].type === 'text') out[0].text = out[0].text.replace(/^\s+/, ''); }
    else if (list && list.kind === 'task' && out[0] && out[0].type === 'text') { const m = /^\s*([☐☑☒])\s*/.exec(out[0].text); if (m) { list.checked = m[1] !== '☐'; out[0].text = out[0].text.slice(m[0].length); } }
    if (list) { delete attrs.indent; delete attrs.firstLine; if (attrs.pstyle === 'listparagraph') delete attrs.pstyle; }
    // numbering switched off (numId 0) or not shown as a list: keep it as it is
    else if (numPr) attrs.ox = { ppr: [...((attrs.ox && attrs.ox.ppr) || []), ['w:numPr', serialize(numPr)]] };
    // split at page breaks
    const blocks = [];
    let cur = [];
    const emit = (content) => {
      const node = level ? { type: 'heading', attrs: { level, ...attrs }, content: normalizeInline(content) } : role === 'code' ? { type: 'codeBlock', attrs: {}, content: [{ type: 'text', text: content.map((x) => x.text || (x.type === 'hardBreak' ? '\n' : '')).join('') }].filter((x) => x.text) } : { type: 'paragraph', attrs: { ...attrs }, content: normalizeInline(content) };
      if (!node.content.length) delete node.content;
      blocks.push(node);
    };
    if (pa.pageBreakBefore) blocks.push({ type: 'pageBreak' });
    for (const n of out) { if (n.type === '__pagebreak') { if (cur.length) emit(cur); blocks.push({ type: 'pageBreak' }); cur = []; } else cur.push(n); }
    if (cur.length || !blocks.length) emit(cur);
    return { blocks, list: list ? { ...list, numId, level: ilvl } : null, role, toc: ctx.toc, fieldOpen: ctx.fields.length > 0, fields: ctx.fields, section };
  }
  function plainPara(pPr) {
    const o = {};
    const jc = val(W(pPr, 'jc'));
    if (jc) o.align = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', left: null, start: null }[jc] || null;
    const sp = W(pPr, 'spacing');
    if (sp) {
      if (sp.a['w:before'] != null) o.before = Number(sp.a['w:before']) / 20;
      if (sp.a['w:after'] != null) o.after = Number(sp.a['w:after']) / 20;
      if (sp.a['w:line'] && (sp.a['w:lineRule'] || 'auto') === 'auto') o.lineHeight = Math.round((Number(sp.a['w:line']) / 240) * 100) / 100;
    }
    const ind = W(pPr, 'ind');
    if (ind) {
      const l = ind.a['w:left'] || ind.a['w:start'];
      if (l != null) o.indent = Math.round(Number(l) / 20);
      if (ind.a['w:hanging']) o.firstLine = -Math.round(Number(ind.a['w:hanging']) / 20);
      else if (ind.a['w:firstLine']) o.firstLine = Math.round(Number(ind.a['w:firstLine']) / 20);
    }
    if (W(pPr, 'pageBreakBefore') && onOff(W(pPr, 'pageBreakBefore'))) o.pageBreakBefore = true;
    return o;
  }

  async function tableNode(tbl) {
    const tblPr = W(tbl, 'tblPr');
    const caption = tblPr && val(W(tblPr, 'tblCaption'));
    // grid widths in px (Word stores its last layout there); a grid without widths → the first row's cell widths
    const grid = kids(W(tbl, 'tblGrid'), 'w:gridCol').map((g) => Math.round((Number(g.a['w:w'] || 0) / 15) * 100) / 100);
    const rows = [];
    const vOrigin = {}; // column → cell node of the merge start
    const firstTcW = [];
    for (const tr of kids(tbl, 'w:tr')) {
      const trPr = W(tr, 'trPr');
      const header = !!(trPr && W(trPr, 'tblHeader') && onOff(W(trPr, 'tblHeader')));
      const row = { type: 'tableRow', content: [] };
      const ra = {};
      const rIns = trPr && W(trPr, 'ins'); const rDel = trPr && W(trPr, 'del');
      if (rIns || rDel) { const rv = rIns || rDel; ra.change = { type: rIns ? 'ins' : 'del', id: rv.a['w:id'] || uid('r'), author: rv.a['w:author'] || '', date: rv.a['w:date'] || null }; }
      if (trPr) {
        const h = W(trPr, 'trHeight');
        if (h && Number(h.a['w:val']) > 0) { ra.height = Number(h.a['w:val']); if (h.a['w:hRule'] === 'exact') ra.heightRule = 'exact'; }
        if (W(trPr, 'cantSplit') && onOff(W(trPr, 'cantSplit'))) ra.cantSplit = true;
        const rest = kids(trPr, '*').filter((x) => !['w:trHeight', 'w:cantSplit', 'w:tblHeader', 'w:ins', 'w:del'].includes(x.n)).map((x) => [x.n, serialize(x)]);
        if (rest.length) ra.ox = rest;
      }
      if (Object.keys(ra).length) row.attrs = ra;
      let col = 0;
      const cellsOf = (el) => kids(el, '*').flatMap((x) => (x.n === 'w:tc' ? [x] : x.n === 'w:sdt' ? kids(W(x, 'sdtContent'), 'w:tc') : x.n === 'w:customXml' ? kids(x, 'w:tc') : []));
      const gb = trPr && W(trPr, 'gridBefore');
      if (gb) col += Number(val(gb)) || 0;
      for (const tc of cellsOf(tr)) {
        const tcPr = W(tc, 'tcPr');
        const span = Number(val(tcPr && W(tcPr, 'gridSpan')) || 1);
        const vm = tcPr && W(tcPr, 'vMerge');
        if (vm && val(vm) !== 'restart') { const o = vOrigin[col]; if (o) o.attrs.rowspan += 1; col += span; continue; }
        const shd = tcPr && W(tcPr, 'shd');
        const fill = shd && shd.a['w:fill'] && shd.a['w:fill'] !== 'auto' ? `#${shd.a['w:fill']}` : null;
        const tcW = tcPr && W(tcPr, 'tcW');
        if (!rows.length && tcW && (tcW.a['w:type'] || 'dxa') === 'dxa') for (let k = 0; k < span; k++) firstTcW[col + k] = Number(tcW.a['w:w']) / span;
        const widths = grid.slice(col, col + span).filter(Boolean);
        const content = await blocksOf(tc);
        const attrs = { colspan: span, rowspan: 1, colwidth: widths.length === span ? widths : null, backgroundColor: fill };
        if (tcPr) {
          const va = val(W(tcPr, 'vAlign')); if (va && va !== 'top') attrs.vAlign = va === 'both' ? 'center' : va;
          const bd = bordersFromEl(W(tcPr, 'tcBorders'), kids); if (bd) attrs.borders = bd;
          const mg = marginsFromEl(W(tcPr, 'tcMar'), kids); if (mg) attrs.margins = mg;
          const td = val(W(tcPr, 'textDirection')); if (td && td !== 'lrTb' && td !== 'tb') attrs.textDir = td === 'tbRlV' ? 'tbRl' : td === 'btLrV' ? 'btLr' : td;
          if (W(tcPr, 'noWrap') && onOff(W(tcPr, 'noWrap'))) attrs.noWrap = true;
          const rest = kids(tcPr, '*').filter((x) => !['w:tcW', 'w:gridSpan', 'w:vMerge', 'w:hMerge', 'w:shd', 'w:vAlign', 'w:tcBorders', 'w:tcMar', 'w:textDirection', 'w:noWrap'].includes(x.n)).map((x) => [x.n, serialize(x)]);
          if (tcW) rest.unshift(['w:tcW', serialize(tcW)]);
          if (rest.length) attrs.ox = rest;
        }
        const cell = { type: header ? 'tableHeader' : 'tableCell', attrs, content: content.length ? content : [{ type: 'paragraph' }] };
        if (vm) vOrigin[col] = cell; else delete vOrigin[col];
        row.content.push(cell);
        col += span;
      }
      if (row.content.length) rows.push(row);
    }
    if (!rows.length) return [];
    if (caption && /^of-callout:/.test(caption) && rows.length === 1 && rows[0].content.length === 1) return [{ type: 'callout', attrs: { kind: caption.split(':')[1] || 'info' }, content: rows[0].content[0].content }];
    if (!grid.some(Boolean) && firstTcW.length) {
      let c = 0;
      for (const cell of rows[0].content) { const sp = cell.attrs.colspan || 1; const ws = firstTcW.slice(c, c + sp).map((w) => Math.round(((w || 0) / 15) * 100) / 100); if (ws.length === sp && ws.every(Boolean)) cell.attrs.colwidth = ws; c += sp; }
    }
    const ta = { ox: tblPr ? { tblPr: serialize(tblPr) } : null };
    if (tblPr) {
      const lay = W(tblPr, 'tblLayout'); if (lay && lay.a['w:type'] === 'fixed') ta.layout = 'fixed';
      const tw = W(tblPr, 'tblW');
      if (tw) { const type = tw.a['w:type'] || 'dxa'; const w = String(tw.a['w:w'] || '0'); if (type === 'pct' || (type === 'dxa' && Number(w) > 0)) ta.width = { type, w: w.endsWith('%') ? Math.round(parseFloat(w) * 50) : Number(w) || 0 }; }
      const ind = W(tblPr, 'tblInd'); if (ind && (ind.a['w:type'] || 'dxa') === 'dxa' && Number(ind.a['w:w'])) ta.indent = Number(ind.a['w:w']);
      const jc = val(W(tblPr, 'jc')); if (jc === 'center' || jc === 'right' || jc === 'end') ta.align = jc === 'end' ? 'right' : jc;
      const bd = bordersFromEl(W(tblPr, 'tblBorders'), kids); if (bd) ta.borders = bd;
      const cm = marginsFromEl(W(tblPr, 'tblCellMar'), kids); if (cm) ta.cellMargins = cm;
      const look = lookFromEl(W(tblPr, 'tblLook')); if (look) ta.look = look;
      if (val(W(tblPr, 'tblStyle'))) ta.styleId = val(W(tblPr, 'tblStyle'));
    }
    // Word's look: the style's borders/margins/colours (the default table style when none is named – no borders)
    ta.tstyle = tableSpec(ta.styleId || (defaultTableStyle && defaultTableStyle.id)) || { cond: {} };
    return [{ type: 'table', attrs: ta, content: rows }];
  }

  /** block container (body, cell, sdtContent …) → blocks, with lists grouped and TOC fields collapsed */
  async function blocksOf(container) {
    const flat = []; // { node, list }
    const items = kids(container, '*');
    let fields = [];
    for (let k = 0; k < items.length; k++) {
      const el = items[k];
      if (el.n === 'w:p') {
        const r = await paragraph(el, { fields });
        fields = r.fields;
        if (r.toc || r.role === 'toc') {
          // a TOC field (or TOC-styled paragraphs): collect until the field ends → one toc node
          const entries = [];
          const addEntry = (p, st) => { const t = paraPlain(p).replace(/\t+\s*\d+\s*$/, '').replace(/\t/g, ' ').trim(); const m = /(\d)\s*$/.exec((styles[st] && styles[st].name) || ''); if (t) entries.push({ level: m ? Number(m[1]) : 1, text: t }); };
          addEntry(el, val(W(W(el, 'pPr') || { c: [] }, 'pStyle')));
          let depth = r.fieldOpen ? 1 : 0;
          while (k + 1 < items.length && items[k + 1].n === 'w:p' && (depth > 0 || roleOfStyle(val(W(W(items[k + 1], 'pPr') || { c: [] }, 'pStyle'))) === 'toc')) {
            k++;
            const p = items[k];
            for (const fc of all(p, 'w:fldChar')) { if (fc.a['w:fldCharType'] === 'begin') depth++; if (fc.a['w:fldCharType'] === 'end') depth--; }
            addEntry(p, val(W(W(p, 'pPr') || { c: [] }, 'pStyle')));
          }
          fields = [];
          flat.push({ node: { type: 'toc', attrs: { title: '', entries } } });
          continue;
        }
        for (const b of r.blocks) flat.push({ node: b, list: r.list && b.type !== 'pageBreak' ? r.list : null, quote: r.role === 'quote' });
        // a paragraph with a w:sectPr ends a section – the section break (own page setup, headers) follows it
        if (r.section) flat.push({ node: { type: 'sectionBreak', attrs: await sectionAttrs(r.section) } });
      } else if (el.n === 'w:tbl') for (const t of await tableNode(el)) flat.push({ node: t });
      else if (el.n === 'w:sdt') {
        const pr = W(el, 'sdtPr');
        const gal = pr && all(pr, 'w:docPartGallery')[0];
        const content = W(el, 'sdtContent') || { c: [] };
        if (gal && /Table of Contents/i.test(val(gal) || '')) {
          const ps = kids(content, 'w:p');
          const entries = []; let title = '';
          for (const p of ps) {
            const st = val(W(W(p, 'pPr') || { c: [] }, 'pStyle'));
            const role = roleOfStyle(st);
            const t = paraPlain(p).replace(/\t+\s*\d+\s*$/, '').replace(/\t/g, ' ').trim();
            if (!t) continue;
            if (role === 'toc') { const m = /(\d)\s*$/.exec(styles[st].name); entries.push({ level: m ? Number(m[1]) : 1, text: t }); } else if (!entries.length) title = t;
          }
          flat.push({ node: { type: 'toc', attrs: { title, entries } } });
        } else {
          // block content control (rich text around paragraphs/tables): kept as a frame around editable blocks
          const info = sdtInfo(pr);
          const inner = await blocksOf(content);
          flat.push({ node: { type: 'controlBlock', attrs: { kind: info.kind, alias: info.alias || null, tag: info.tag || null, pr: pr ? serialize(pr) : '', end: W(el, 'sdtEndPr') ? serialize(W(el, 'sdtEndPr')) : '' }, content: inner.length ? inner : [{ type: 'paragraph' }] } });
        }
      } else if (el.n === 'w:customXml') { for (const b of await blocksOf(el)) flat.push({ node: b }); }
      else if (el.n === 'mc:AlternateContent') { const pv = await preservedOf({ n: 'w:p', a: {}, c: [el] }, el); flat.push({ node: { type: 'paragraph', content: [pv] } }); }
      else if (el.n === 'w:sectPr' || el.n === 'w:bookmarkStart' || el.n === 'w:bookmarkEnd' || el.n === 'w:proofErr') { /* handled elsewhere / no content */ }
      else if (el.n === 'w:commentRangeStart') openComments.add(el.a['w:id']);
      else if (el.n === 'w:commentRangeEnd') openComments.delete(el.a['w:id']);
      else if (el.n === 'w:altChunk') fid.add('altChunk');
    }
    return groupBlocks(flat);
  }

  /** one header/footer part → { text, align, doc } (rich content with fields, pictures, tables) */
  const hfCache = {};
  async function readHf(rid) {
    const rr = docRels[rid];
    if (!rr) return null;
    if (hfCache[rr.target]) return hfCache[rr.target];
    const hx = await readText(zip, rr.target);
    if (!hx) return null;
    const hr = parseXml(hx);
    const ps = all(hr, 'w:p');
    const t = ps.map((p) => paraPlain(p, true)).filter((x) => x.trim()).join(' ').replace(/\s+/g, ' ').trim();
    const firstFilled = ps.find((p) => paraPlain(p, true).trim());
    const jc = firstFilled && val(W(W(firstFilled, 'pPr') || { c: [] }, 'jc'));
    const tabbed = firstFilled && /\t/.test(paraPlain(firstFilled, true).trim());
    // the part's own relationships (pictures, links) while reading its content
    const saved = rels; const savedComments = openComments;
    rels = relsOf(await readText(zip, rr.target.replace(/([^/]+)$/, '_rels/$1.rels')), rr.target);
    openComments = new Set();
    let blocks = [];
    try { blocks = await blocksOf(hr); } finally { rels = saved; openComments = savedComments; }
    const doc = fixDoc({ type: 'doc', content: blocks });
    const res = { text: t, align: jc === 'center' || tabbed ? 'center' : jc === 'right' || jc === 'end' ? 'right' : 'left', doc, part: rr.target };
    hfCache[rr.target] = res;
    return res;
  }
  /** a w:sectPr → page setup, break type, header/footer of that section */
  async function sectionAttrs(sect) {
    const page = { ...model.page, margins: { ...model.page.margins } };
    const pg = W(sect, 'pgSz');
    if (pg) {
      const w = twipToMm(pg.a['w:w']); const h = twipToMm(pg.a['w:h']);
      const land = pg.a['w:orient'] === 'landscape' || w > h;
      page.orientation = land ? 'landscape' : 'portrait';
      page.width = Math.min(w, h); page.height = Math.max(w, h);
      page.size = sizeName(w, h);
    }
    const mar = W(sect, 'pgMar');
    if (mar) page.margins = { top: twipToMm(Math.abs(Number(mar.a['w:top'] || 1440))), right: twipToMm(mar.a['w:right'] || 1440), bottom: twipToMm(Math.abs(Number(mar.a['w:bottom'] || 1440))), left: twipToMm(mar.a['w:left'] || 1440) };
    const cols = W(sect, 'cols');
    page.columns = cols && Number(cols.a['w:num']) > 1 ? Math.min(3, Number(cols.a['w:num'])) : 1;
    const hf = {};
    for (const [tagName, key] of [['w:headerReference', 'header'], ['w:footerReference', 'footer']]) {
      for (const ref of kids(sect, tagName)) {
        const type = ref.a['w:type'] || 'default';
        const h = await readHf(ref.a['r:id']);
        if (h) hf[`${key}${type === 'default' ? '' : type === 'first' ? 'First' : 'Even'}`] = h;
      }
    }
    return { type: val(W(sect, 'type')) || 'nextPage', page, titlePage: !!(W(sect, 'titlePg') && onOff(W(sect, 'titlePg'))), hf, sectPr: serialize(sect) };
  }

  const content = await blocksOf(body);
  model.doc = fixDoc({ type: 'doc', content });

  // final section: page setup, header/footer
  const sect = W(body, 'sectPr');
  if (sect) {
    const s = await sectionAttrs(sect);
    model.page = s.page;
    if (s.titlePage) model.titlePage = true;
    for (const k of ['header', 'footer', 'headerFirst', 'footerFirst', 'headerEven', 'footerEven']) if (s.hf[k]) model[k] = s.hf[k];
  }
  // sections before inherit their headers/footers from the one before (Word's rule) – resolved when shown

  // settings, core properties
  const setRel = relOfType('settings');
  const setXml = setRel ? await readText(zip, setRel.target) : null;
  if (setXml && /<w:trackRevisions(?:\s*\/>|\s+w:val="(?:1|true|on)")/.test(setXml)) model.trackChanges = true;
  if (setXml && /<w:evenAndOddHeaders(?:\s*\/>|\s+w:val="(?:1|true|on)")/.test(setXml)) model.evenOdd = true;
  const coreRel = Object.values(rootRels).find((r) => r.type === 'core-properties');
  const core = coreRel ? await readText(zip, coreRel.target) : null;
  if (core) {
    const g = (tagName) => { const m = new RegExp(`<${tagName}[^>]*>([^<]*)</${tagName}>`).exec(core); return m ? m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>') : ''; };
    model.meta = { ...model.meta, title: g('dc:title'), author: g('dc:creator') || model.meta.author, subject: g('dc:subject'), keywords: g('cp:keywords'), created: g('dcterms:created'), modified: g('dcterms:modified') };
  }
  if (zip.file('word/vbaProject.bin') || Object.values(rels).some((r) => r.type === 'vbaProject')) fid.add('macros');
  if (Object.keys(zip.files).some((f) => f.startsWith('_xmlsignatures/'))) fid.add('signature');
  if (Object.values(rels).some((r) => r.type === 'glossaryDocument')) { /* building blocks: kept as part */ }

  const hasChanges = /<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b/.test(docXml);
  if (hasChanges) fid.add('trackedChanges');
  if (/<w:pPrChange\b/.test(docXml)) fid.add('paragraphChanges');
  model.fidelity = [...fid];
  model.keep.docx = {
    zip: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), docPath, rootTag: rootTag ? rootTag[0] : null,
    sectPr: sect ? serialize(sect) : null, header: { ...model.header }, footer: { ...model.footer }, ext,
    hf: Object.fromEntries(['header', 'footer', 'headerFirst', 'footerFirst', 'headerEven', 'footerEven'].map((k) => [k, model[k] ? hfKey(model[k]) : ''])),
  };
  return model;
}

/** plain text of a paragraph element (fields: result text; with tokens=true PAGE/NUMPAGES become {page}/{pages}) */
export function paraPlain(p, tokens = false) {
  let s = '';
  let inInstr = false; let instr = ''; let skipResult = false;
  const visit = (el) => {
    for (const c of el.c || []) {
      if (typeof c === 'string') continue;
      if (c.n === 'w:fldChar') {
        const t = c.a['w:fldCharType'];
        if (t === 'begin') { inInstr = true; instr = ''; } else if (t === 'separate') { inInstr = false; const f = instr.trim().split(/\s+/)[0].toUpperCase(); if (tokens && (f === 'PAGE' || f === 'NUMPAGES')) { s += f === 'PAGE' ? '{page}' : '{pages}'; skipResult = true; } } else if (t === 'end') { if (inInstr && tokens) { const f = instr.trim().split(/\s+/)[0].toUpperCase(); if (f === 'PAGE') s += '{page}'; if (f === 'NUMPAGES') s += '{pages}'; } inInstr = false; skipResult = false; }
        continue;
      }
      if (c.n === 'w:instrText') { instr += xtext(c); continue; }
      if (c.n === 'w:fldSimple' && tokens) { const f = (c.a['w:instr'] || '').trim().split(/\s+/)[0].toUpperCase(); if (f === 'PAGE') { s += '{page}'; continue; } if (f === 'NUMPAGES') { s += '{pages}'; continue; } }
      if (skipResult && c.n === 'w:r' && !kid(c, 'w:fldChar')) continue;
      if (c.n === 'w:t' || c.n === 'w:delText') { if (c.n === 'w:t' || !tokens) s += xtext(c); continue; }
      if (c.n === 'w:tab' || c.n === 'w:ptab') { s += '\t'; continue; }
      if (c.n === 'w:br' || c.n === 'w:cr') { s += ' '; continue; }
      if (c.n === 'w:del' && tokens) continue;
      visit(c);
    }
  };
  visit(p);
  return s;
}

/** group flat blocks: consecutive list paragraphs → nested lists, quote paragraphs → blockquote */
export function groupBlocks(flat) {
  const out = [];
  let k = 0;
  while (k < flat.length) {
    const f = flat[k];
    if (f.list) {
      const run = [];
      while (k < flat.length && flat[k].list) { run.push(flat[k]); k++; }
      out.push(...nestList(run));
      continue;
    }
    if (f.quote) {
      const q = { type: 'blockquote', content: [] };
      while (k < flat.length && flat[k].quote && !flat[k].list) { q.content.push(flat[k].node); k++; }
      out.push(q);
      continue;
    }
    out.push(f.node);
    k++;
  }
  return out;
}
function nestList(run) {
  const minLevel = Math.min(...run.map((r) => r.list.level || 0));
  const build = (items, level) => {
    const lists = []; let cur = null; let idx = 0;
    while (idx < items.length) {
      const it = items[idx];
      const l = it.list;
      if ((l.level || 0) > level && cur && cur.content.length) {
        const sub = [];
        while (idx < items.length && (items[idx].list.level || 0) > level) sub.push(items[idx++]);
        cur.content[cur.content.length - 1].content.push(...build(sub, level + 1));
        continue;
      }
      const type = l.kind === 'ordered' ? 'orderedList' : l.kind === 'task' ? 'taskList' : 'bulletList';
      if (!cur || cur.type !== type || (cur.__num !== l.numId && level === minLevel && type === 'orderedList')) {
        cur = { type, attrs: type === 'orderedList' ? { start: l.start || 1, listStyle: l.listStyle || null, ox: l.numId ? { numId: l.numId } : null } : type === 'bulletList' ? { listStyle: null, ox: l.numId ? { numId: l.numId } : null } : {}, content: [], __num: l.numId };
        lists.push(cur);
      }
      const p = it.node.type === 'heading' ? { ...it.node, type: 'paragraph', attrs: { ...it.node.attrs } } : it.node;
      if (p.attrs) delete p.attrs.level;
      cur.content.push(type === 'taskList' ? { type: 'taskItem', attrs: { checked: !!l.checked }, content: [p] } : { type: 'listItem', content: [p] });
      idx++;
    }
    for (const x of lists) delete x.__num;
    return lists;
  };
  return build(run, minLevel);
}

/* ---------------- OMML → LaTeX-ish ---------------- */
export function ommlToLatex(el) {
  if (!el || typeof el === 'string') return '';
  const ch = (name) => kid(el, `m:${name}`);
  const inner = (x) => (x ? (x.c || []).map((c) => (typeof c === 'string' ? '' : ommlToLatex(c))).join('') : '');
  const g = (x) => { const s = inner(x); return s.length === 1 ? s : `{${s}}`; };
  switch (local(el.n)) {
    case 'r': return all(el, 'm:t').map(xtext).join('') || all(el, 'w:t').map(xtext).join('');
    case 'f': return `\\frac{${inner(ch('num'))}}{${inner(ch('den'))}}`;
    case 'sSup': return `${inner(ch('e'))}^${g(ch('sup'))}`;
    case 'sSub': return `${inner(ch('e'))}_${g(ch('sub'))}`;
    case 'sSubSup': return `${inner(ch('e'))}_${g(ch('sub'))}^${g(ch('sup'))}`;
    case 'd': {
      const pr = ch('dPr'); const b = pr && kid(pr, 'm:begChr'); const e = pr && kid(pr, 'm:endChr');
      const fence = (v, def) => { const s = v == null ? def : v; return s === '' ? '.' : s === '{' ? '\\{' : s === '}' ? '\\}' : s === '|' ? '|' : s; };
      return `\\left${fence(b && b.a['m:val'], '(')} ${kids(el, 'm:e').map(inner).join(',')} \\right${fence(e && e.a['m:val'], ')')}`;
    }
    case 'nary': {
      const pr = ch('naryPr'); const c = pr && kid(pr, 'm:chr'); const sym = c ? c.a['m:val'] : '∫';
      const cmd = { '∑': '\\sum', '∏': '\\prod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint', '⋃': '\\bigcup', '⋂': '\\bigcap' }[sym] || sym;
      return `${cmd}${ch('sub') && inner(ch('sub')) ? `_${g(ch('sub'))}` : ''}${ch('sup') && inner(ch('sup')) ? `^${g(ch('sup'))}` : ''} ${inner(ch('e'))}`;
    }
    case 'func': return `${inner(ch('fName'))} ${inner(ch('e'))}`;
    case 'eqArr': return `\\begin{aligned}${kids(el, 'm:e').map(inner).join(' \\\\ ')}\\end{aligned}`;
    case 'm': return `\\begin{matrix}${kids(el, 'm:mr').map((r) => kids(r, 'm:e').map(inner).join(' & ')).join(' \\\\ ')}\\end{matrix}`;
    case 'acc': { const pr = ch('accPr'); const c = pr && kid(pr, 'm:chr'); const a = c ? c.a['m:val'] : '̂'; const cmd = { '̂': 'hat', '⃗': 'vec', '̃': 'tilde', '̇': 'dot', '̈': 'ddot', '̄': 'bar', '̌': 'check' }[a] || 'hat'; return `\\${cmd}{${inner(ch('e'))}}`; }
    case 'bar': return `\\overline{${inner(ch('e'))}}`;
    case 'limLow': return `${inner(ch('e'))}_${g(ch('lim'))}`;
    case 'limUpp': return `${inner(ch('e'))}^${g(ch('lim'))}`;
    case 'borderBox': return `\\boxed{${inner(ch('e'))}}`;
    case 'rad': { const d = ch('deg'); const dv = d ? inner(d) : ''; return dv ? `\\sqrt[${dv}]{${inner(ch('e'))}}` : `\\sqrt{${inner(ch('e'))}}`; }
    case 'rPr': case 'ctrlPr': case 'fPr': case 'radPr': case 'sSupPr': case 'sSubPr': case 'dPr': case 'naryPr': case 'oMathParaPr': case 'deg': case 'accPr': case 'barPr': case 'mPr': case 'limLowPr': case 'limUppPr': case 'eqArrPr': case 'funcPr': case 'sSubSupPr': case 'borderBoxPr': return '';
    default: return inner(el);
  }
}

/** plain text of a docx (search, preview, AI) */
export async function docxToText(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const x = await readText(zip, 'word/document.xml');
  if (!x) return '';
  const body = W(parseXml(x), 'body');
  return all(body, 'w:p').map((p) => paraPlain(p)).join('\n');
}
void imageSize;

