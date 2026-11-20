/*
 * RTF reader and writer (own code, no library).
 *
 * Reader: text, bold/italic/underline/strike/super/sub, fonts, sizes, colours, highlight, paragraphs (alignment,
 * indents, spacing, line spacing), headings (\outlinelevel or "heading n" styles), lists (from \listtext/\pntext –
 * bullets, numbers, ☐/☑ check lists, levels), tables (\trowd … \cell \row, horizontal merges, cell shading), links
 * (HYPERLINK fields), images (\pngblip, \jpegblip), footnotes, comments (annotations), tracked changes (\revised,
 * \deleted), page breaks, page size/margins/orientation/columns, header/footer, document title/author.
 * Writer: the same features back (legacy \pn numbering, which Word, LibreOffice and WordPad all read).
 */
import { defaultModel, sizeName, twipToMm, mmToTwip, firstFont, ptOf, hexColor, normalizeInline, fixDoc, dataUrl, parseDataUrl, imageSize, walk, inlineText, pageBox } from './pm.js';
import { equationText } from './equation.js';

const CHARSET_CP = { 0: 1252, 1: 1252, 2: 1252, 77: 10000, 128: 932, 129: 949, 134: 936, 136: 950, 161: 1253, 162: 1254, 163: 1258, 177: 1255, 178: 1256, 186: 1257, 204: 1251, 222: 874, 238: 1250 };
function decoderFor(cp) {
  const names = { 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 10000: 'macintosh', 874: 'windows-874' };
  const name = names[cp] || `windows-${cp}`;
  try { return new TextDecoder(name); } catch { return new TextDecoder('windows-1252'); }
}
const SKIP_DEST = new Set(['fonttbl', 'colortbl', 'stylesheet', 'info', 'listtable', 'listoverridetable', 'revtbl', 'rsidtbl', 'generator', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'xmlnstbl', 'mmathPr', 'pgdsctbl', 'filetbl', 'bkmkstart', 'bkmkend', 'nonshppict', 'fldtype', 'xe', 'tc', 'pnseclvl', 'listpicture', 'wgrffmtfilter', 'docvar', 'userprops', 'passwordhash', 'defchp', 'defpap', 'ftnsep', 'ftnsepc', 'aftnsep', 'aftnsepc', 'template', 'object', 'objdata', 'shpinst', 'atnid', 'atnauthor', 'atndate', 'atnref', 'atnparent', 'atnicn', 'atntime', 'datafield', 'formfield', 'ffdeftext', 'blipuid', 'panose', 'falt', 'fname', 'background']);
const HEADER_DEST = { header: 'header', headerr: 'header', headerl: 'header', footer: 'footer', footerr: 'footer', footerl: 'footer', headerf: 'skip', footerf: 'skip' };

export function parseRtf(input, settings) {
  const src = typeof input === 'string' ? input : new TextDecoder('latin1').decode(input);
  const model = defaultModel(settings);
  model.defaults.spaceAfter = 0;
  const fonts = {}; const fontCp = {}; const colors = []; const styles = {}; const revAuthors = [];
  let docCp = 1252;
  let deff = 0;
  const blocks = [];
  let inl = [];
  // paragraph and character state (scoped by groups)
  const P0 = () => ({ align: null, li: 0, fi: 0, sb: null, sa: null, sl: 0, slmult: 0, intbl: false, ilvl: 0, list: false, outline: null, style: 0, pagebb: false });
  const C0 = () => ({ b: false, i: false, u: false, strike: false, sup: false, sub: false, fs: null, f: null, cf: 0, hl: 0, hidden: false, ins: 0, del: 0, link: null });
  let st = { dest: 'body', uc: 1, p: P0(), c: C0() };
  const stack = [];
  let skipChars = 0;
  let hexBuf = [];
  let cur = null; // collector for special destinations { kind, text, … }
  let fontEntry = null; let colorEntry = null; let styleEntry = null;
  let pict = null;
  let field = null; // { inst: '', inResult }
  const fieldStack = [];
  let listText = null;
  let pendingFootnote = null;
  let annotation = null;
  const openAnn = {}; // atrfstart id → true
  const annRanges = []; // [{ id, start: {block,inline}, end }]
  let annCounter = 0;
  // table builder
  let row = null; let rowDef = { cellx: [], bg: [], merge: [] }; let nextCell = { bg: 0, merge: null }; let tableRows = null; let cellBlocks = [];
  let lastListMeta = null;
  const listParas = []; // [{ blockIndex, level, kind }]

  const flushHex = () => {
    if (!hexBuf.length) return;
    const cp = (st.c.f != null && fontCp[st.c.f]) || docCp;
    addText(decoderFor(cp).decode(new Uint8Array(hexBuf)));
    hexBuf = [];
  };
  const marks = () => {
    const c = st.c; const m = [];
    if (c.b) m.push({ type: 'bold' });
    if (c.i) m.push({ type: 'italic' });
    if (c.u) m.push({ type: 'underline' });
    if (c.strike) m.push({ type: 'strike' });
    if (c.sup) m.push({ type: 'superscript' });
    if (c.sub) m.push({ type: 'subscript' });
    const ts = {};
    if (c.f != null && fonts[c.f] && c.f !== deff) ts.fontFamily = fonts[c.f];
    if (c.fs) ts.fontSize = `${c.fs / 2}pt`;
    if (c.cf && colors[c.cf]) ts.color = `#${colors[c.cf]}`;
    if (Object.keys(ts).length) m.push({ type: 'textStyle', attrs: { color: ts.color || null, fontFamily: ts.fontFamily || null, fontSize: ts.fontSize || null } });
    if (c.hl && colors[c.hl]) m.push({ type: 'highlight', attrs: { color: `#${colors[c.hl]}` } });
    if (c.link) m.push({ type: 'link', attrs: { href: c.link } });
    if (c.ins) m.push({ type: 'insertion', attrs: { id: `r${c.ins}`, author: revAuthors[c.ins] || '', date: null } });
    if (c.del) m.push({ type: 'deletion', attrs: { id: `r${c.del}`, author: revAuthors[c.del] || '', date: null } });
    return m;
  };
  function addText(t) {
    if (!t) return;
    const d = st.dest;
    if (d === 'skip') return;
    if (d === 'fonttbl') { if (fontEntry) fontEntry.name += t; return; }
    if (d === 'stylesheet') { if (styleEntry) styleEntry.name += t; return; }
    if (d === 'fldinst') { field.inst += t; return; }
    if (d === 'listtext') { listText += t; return; }
    if (d === 'title' || d === 'author' || d === 'subject' || d === 'keywords') { model.meta[d] = (model.meta[d] || '') + t; return; }
    if (d === 'revauthor') { revAuthors[revAuthors.length - 1] += t; return; }
    if (d === 'pict') { pict.hex += t; return; }
    if (st.c.hidden) return;
    if (d === 'footnote' || d === 'annotation' || d === 'header' || d === 'footer') { cur.text += t; return; }
    if (d !== 'body') return;
    inl.push({ type: 'text', text: t, marks: marks() });
  }
  const paraAttrs = () => {
    const p = st.p; const a = {};
    if (p.align) a.textAlign = p.align;
    if (p.li && !p.list) a.indent = Math.round(p.li / 20);
    if (p.fi && !p.list) a.firstLine = Math.round(p.fi / 20);
    if (p.sb != null) a.spaceBefore = Math.round(p.sb / 20);
    if (p.sa != null) a.spaceAfter = Math.round(p.sa / 20);
    if (p.sl && p.slmult) a.lineHeight = Math.round((Math.abs(p.sl) / 240) * 100) / 100;
    return a;
  };
  const headingLevel = () => {
    if (st.p.outline != null && st.p.outline < 6) return st.p.outline + 1;
    const s = styles[st.p.style];
    if (s && s.level) return s.level;
    return 0;
  };
  function endPara(kind = 'par') {
    flushHex();
    const content = normalizeInline(inl);
    inl = [];
    let node;
    const lvl = headingLevel();
    if (lvl) node = { type: 'heading', attrs: { level: lvl, ...paraAttrs() }, content };
    else node = { type: 'paragraph', attrs: paraAttrs(), content };
    if (styles[st.p.style] && styles[st.p.style].pstyle && !lvl) node.attrs.pstyle = styles[st.p.style].pstyle;
    if (!node.content.length) delete node.content;
    if (st.p.pagebb) blocks.push({ type: 'pageBreak' });
    if (st.p.intbl) {
      cellBlocks.push(node);
      if (kind === 'cell') {
        if (!row) row = [];
        row.push({ blocks: cellBlocks });
        cellBlocks = [];
      }
      return;
    }
    if (tableRows) closeTable();
    const li = listText != null || st.p.list ? (lastListMeta || { kind: 'bullet' }) : null;
    blocks.push(node);
    if (li) listParas.push({ index: blocks.length - 1, level: st.p.ilvl || 0, kind: li.kind, checked: li.checked });
    listText = null; lastListMeta = null;
  }
  function endRow() {
    if (cellBlocks.length || inl.length) { endPara('cell'); }
    if (!row) return;
    const cells = row.map((c, k) => {
      const bg = rowDef.bg[k] && colors[rowDef.bg[k]] ? `#${colors[rowDef.bg[k]]}` : null;
      const w = rowDef.cellx[k] != null ? Math.round(((rowDef.cellx[k] - (k ? rowDef.cellx[k - 1] : 0)) / 15)) : null;
      return { merge: rowDef.merge[k], node: { type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: w > 0 ? [w] : null, backgroundColor: bg }, content: c.blocks.length ? c.blocks : [{ type: 'paragraph' }] } };
    });
    // horizontal merges: \clmgf starts, \clmrg continues
    const merged = [];
    for (const c of cells) {
      if (c.merge === 'cont' && merged.length) { const prev = merged[merged.length - 1]; prev.attrs.colspan += 1; if (prev.attrs.colwidth && c.node.attrs.colwidth) prev.attrs.colwidth = [...prev.attrs.colwidth, ...c.node.attrs.colwidth]; continue; }
      merged.push(c.node);
    }
    if (!tableRows) tableRows = [];
    tableRows.push({ type: 'tableRow', content: merged });
    row = null;
  }
  function closeTable() {
    if (!tableRows) return;
    blocks.push({ type: 'table', content: tableRows });
    tableRows = null;
  }

  const pushState = () => { stack.push(st); st = { dest: st.dest, uc: st.uc, p: { ...st.p }, c: { ...st.c } }; };
  const popState = () => {
    flushHex();
    const ending = st.dest;
    const prev = stack.pop() || st;
    // destination ends
    if (ending === 'fonttbl' && fontEntry && prev.dest !== 'fonttbl') fontEntry = null;
    if (ending === 'pict' && prev.dest !== 'pict' && pict) {
      const hex = pict.hex.replace(/[^0-9a-fA-F]/g, '');
      if (pict.type && hex.length > 8) {
        const u8 = new Uint8Array(hex.length >> 1);
        for (let k = 0; k < u8.length; k++) u8[k] = parseInt(hex.substr(k * 2, 2), 16);
        const sz = imageSize(u8);
        let width = pict.wgoal ? Math.round((pict.wgoal / 15) * (pict.scalex || 100) / 100) : (sz ? sz.width : null);
        if (width > 650) width = 650;
        const node = { type: 'image', attrs: { src: dataUrl(u8, pict.type), alt: '', title: null, width: width || null, align: null, wrap: null } };
        if (prev.dest === 'body') inl.push(node);
      } else if (!pict.type) model.fidelity.push('images');
      pict = null;
    }
    if (ending === 'fldinst' && prev.dest !== 'fldinst' && field) {
      const m = /HYPERLINK\s+(?:\\l\s+)?"([^"]+)"/i.exec(field.inst) || /HYPERLINK\s+(\S+)/i.exec(field.inst);
      if (m) field.href = m[1];
      if (/^\s*PAGE\b/i.test(field.inst)) field.token = '{page}';
      if (/^\s*NUMPAGES\b/i.test(field.inst)) field.token = '{pages}';
      if (/^\s*TOC\b/i.test(field.inst)) field.toc = true;
    }
    if ((ending === 'footnote') && prev.dest !== 'footnote' && cur) {
      const node = { type: 'footnote', attrs: { text: cur.text.replace(/\s+/g, ' ').trim() } };
      if (prev.dest === 'body') inl.push(node);
      cur = null;
    }
    if (ending === 'annotation' && prev.dest !== 'annotation' && cur) {
      annotation = cur; cur = null;
      const id = annotation.ref || `a${++annCounter}`;
      model.comments.push({ id, author: annotation.author || '', date: null, text: annotation.text.trim(), resolved: false, replies: [] });
      // no range given → anchor to the last word before the reference
      if (!annRanges.find((r) => r.id === id)) annRanges.push({ id, lastWord: true, block: blocks.length, inline: inl.length });
    }
    if ((ending === 'header' || ending === 'footer') && prev.dest !== ending && cur) {
      const text = cur.text.replace(/\s+/g, ' ').trim();
      if (text && !model[ending].text) model[ending] = { text, align: cur.align || 'center' };
      cur = null;
    }
    if (ending === 'listtext' && prev.dest !== 'listtext') {
      const t = (listText || '').trim();
      lastListMeta = /[☐☑☒]/.test(t) ? { kind: 'task', checked: /[☑☒]/.test(t) } : /[0-9a-zA-Z]/.test(t) ? { kind: 'ordered' } : { kind: 'bullet' };
    }
    if (ending === 'stylesheet' && styleEntry && prev.dest !== 'stylesheet') styleEntry = null;
    if (st.fieldEnd && field) { // the field group closes
      if (field.inResult && field.href) { /* link already applied */ }
      field = fieldStack.pop() || null;
    }
    st = prev;
  };

  const ctrl = (word, param, hasParam) => {
    const on = !hasParam || param !== 0;
    const d = st.dest;
    // destinations
    if (d === 'skip') { if (word === 'bin' && hasParam) i += param; return; }
    switch (word) {
      case 'ansicpg': docCp = param; return;
      case 'deff': deff = param; return;
      case 'fonttbl': st.dest = 'fonttbl'; return;
      case 'colortbl': st.dest = 'colortbl'; colorEntry = { r: 0, g: 0, b: 0, set: false }; return;
      case 'stylesheet': st.dest = 'stylesheet'; return;
      case 'revtbl': st.dest = 'revtbl'; return;
      case 'info': st.dest = 'info'; return;
      case 'title': case 'author': case 'subject': case 'keywords': if (d === 'info') { st.dest = word; model.meta[word] = ''; } return;
      case 'pict': st.dest = 'pict'; pict = { hex: '', type: null, wgoal: 0, scalex: 100 }; return;
      case 'pngblip': if (pict) pict.type = 'image/png'; return;
      case 'jpegblip': if (pict) pict.type = 'image/jpeg'; return;
      case 'emfblip': case 'wmetafile': case 'macpict': case 'dibitmap': case 'wbitmap': if (pict) pict.type = null; return;
      case 'picwgoal': if (pict) pict.wgoal = param; return;
      case 'picscalex': if (pict) pict.scalex = param; return;
      case 'shppict': return; // keep: contains the picture
      case 'field': fieldStack.push(field); field = { inst: '', href: null, inResult: false }; st.fieldEnd = true; return;
      case 'fldinst': st.dest = 'fldinst'; return;
      case 'fldrslt':
        if (field) {
          field.inResult = true;
          if (field.href) st.c.link = field.href;
          if (field.token && (d === 'header' || d === 'footer')) { cur.text += field.token; st.dest = 'skip'; }
        }
        return;
      case 'footnote': if (d === 'body') { flushHex(); st.dest = 'footnote'; cur = { text: '' }; } else st.dest = 'skip'; return;
      case 'annotation': st.dest = 'annotation'; cur = { text: '', author: pendingAuthor, ref: pendingRef }; pendingAuthor = ''; pendingRef = null; return;
      case 'atnauthor': st.dest = 'atnauthor'; return;
      case 'atrfstart': st.dest = 'atrfstart'; return;
      case 'atrfend': st.dest = 'atrfend'; return;
      case 'listtext': case 'pntext': st.dest = 'listtext'; listText = ''; return;
      case 'pn': st.dest = 'skip'; st.p.list = true; return;
      case 'header': case 'headerr': case 'headerl': case 'footer': case 'footerr': case 'footerl': case 'headerf': case 'footerf':
        if (HEADER_DEST[word] === 'skip' || model[HEADER_DEST[word]].text) st.dest = 'skip'; else { st.dest = HEADER_DEST[word]; cur = { text: '', align: null }; }
        return;
      case 'chftn': case 'chatn': return;
      default:
    }
    if (SKIP_DEST.has(word) && !['fonttbl', 'colortbl', 'stylesheet', 'info', 'revtbl'].includes(word)) { st.dest = word === 'atnid' ? 'skip' : 'skip'; return; }
    // font table entries
    if (d === 'fonttbl') { if (word === 'f') { fontEntry = { num: param, name: '' }; fonts[param] = ''; } else if (word === 'fcharset' && fontEntry) fontCp[fontEntry.num] = CHARSET_CP[param] || 1252; return; }
    if (d === 'colortbl') { if (word === 'red') { colorEntry.r = param; colorEntry.set = true; } else if (word === 'green') { colorEntry.g = param; colorEntry.set = true; } else if (word === 'blue') { colorEntry.b = param; colorEntry.set = true; } return; }
    if (d === 'stylesheet') { if (word === 's' || word === 'cs' || word === 'ds' || word === 'ts') { styleEntry = { num: word === 's' ? param : -1, name: '' }; } else if (word === 'outlinelevel' && styleEntry && styleEntry.num >= 0) styleEntry.outline = param; return; }
    // page setup
    switch (word) {
      case 'paperw': model.page.width = twipToMm(param); return;
      case 'paperh': model.page.height = twipToMm(param); return;
      case 'margl': model.page.margins.left = twipToMm(param); return;
      case 'margr': model.page.margins.right = twipToMm(param); return;
      case 'margt': model.page.margins.top = twipToMm(param); return;
      case 'margb': model.page.margins.bottom = twipToMm(param); return;
      case 'landscape': landscape = true; return;
      case 'cols': model.page.columns = Math.min(3, Math.max(1, param)); return;
      case 'titlepg': model.titlePage = true; return;
      default:
    }
    // paragraph
    switch (word) {
      case 'par': if (d === 'body') endPara(); else if (cur) cur.text += ' '; return;
      case 'pard': st.p = { ...P0(), intbl: false }; return;
      case 'plain': st.c = { ...C0(), link: st.c.link }; return;
      case 'ql': st.p.align = null; if (cur && d !== 'body') cur.align = 'left'; return;
      case 'qc': st.p.align = 'center'; if (cur && d !== 'body') cur.align = 'center'; return;
      case 'qr': st.p.align = 'right'; if (cur && d !== 'body') cur.align = 'right'; return;
      case 'qj': st.p.align = 'justify'; return;
      case 'li': st.p.li = param; return;
      case 'fi': st.p.fi = param; return;
      case 'sb': st.p.sb = param; return;
      case 'sa': st.p.sa = param; return;
      case 'sl': st.p.sl = param; return;
      case 'slmult': st.p.slmult = param; return;
      case 'outlinelevel': st.p.outline = param; return;
      case 's': st.p.style = param; return;
      case 'ls': st.p.list = true; return;
      case 'ilvl': st.p.ilvl = param; return;
      case 'pnlvl': st.p.ilvl = Math.max(0, param - 1); return;
      case 'intbl': st.p.intbl = true; return;
      case 'pagebb': st.p.pagebb = true; return;
      case 'page': if (d === 'body') { if (inl.length) endPara(); if (tableRows) closeTable(); blocks.push({ type: 'pageBreak' }); } return;
      case 'line': if (d === 'body') inl.push({ type: 'hardBreak' }); else if (cur) cur.text += ' '; return;
      case 'tab': addText('\t'); return;
      case 'cell': if (d === 'body') { st.p.intbl = true; endPara('cell'); } return;
      case 'nestcell': if (d === 'body') { st.p.intbl = true; endPara('cell'); } return;
      case 'row': if (d === 'body') endRow(); return;
      case 'trowd': rowDef = { cellx: [], bg: [], merge: [] }; nextCell = { bg: 0, merge: null }; return;
      case 'cellx': rowDef.cellx.push(param); rowDef.bg.push(nextCell.bg); rowDef.merge.push(nextCell.merge); nextCell = { bg: 0, merge: null }; return;
      case 'clcbpat': nextCell.bg = param; return;
      case 'clmgf': nextCell.merge = 'first'; return;
      case 'clmrg': nextCell.merge = 'cont'; return;
      default:
    }
    // character
    switch (word) {
      case 'b': st.c.b = on; return;
      case 'i': st.c.i = on; return;
      case 'ul': case 'uld': case 'uldb': case 'ulw': case 'ulth': case 'uldash': case 'ulwave': st.c.u = on; return;
      case 'ulnone': st.c.u = false; return;
      case 'strike': case 'striked': st.c.strike = on; return;
      case 'super': case 'up': st.c.sup = on; st.c.sub = false; return;
      case 'sub': case 'dn': st.c.sub = on; st.c.sup = false; return;
      case 'nosupersub': st.c.sup = false; st.c.sub = false; return;
      case 'fs': st.c.fs = param; return;
      case 'f': st.c.f = param; return;
      case 'cf': st.c.cf = param; return;
      case 'highlight': case 'cb': case 'chcbpat': st.c.hl = param; return;
      case 'v': st.c.hidden = on; return;
      case 'revised': st.c.ins = on ? (st.c.revauth || 1) : 0; return;
      case 'deleted': st.c.del = on ? (st.c.revauth || 1) : 0; return;
      case 'revauth': st.c.revauth = param; if (st.c.ins) st.c.ins = param; return;
      case 'revauthdel': st.c.revauth = param; if (st.c.del) st.c.del = param; return;
      case 'uc': st.uc = param; return;
      case 'u': { flushHex(); let cp = param; if (cp < 0) cp += 65536; addText(String.fromCharCode(cp)); skipChars = st.uc; return; }
      case 'emdash': addText('—'); return;
      case 'endash': addText('–'); return;
      case 'bullet': addText('•'); return;
      case 'lquote': addText('‘'); return;
      case 'rquote': addText('’'); return;
      case 'ldblquote': addText('“'); return;
      case 'rdblquote': addText('”'); return;
      case 'emspace': case 'enspace': case 'qmspace': addText(' '); return;
      default:
    }
  };
  let landscape = false;
  let pendingAuthor = ''; let pendingRef = null;

  let i = 0;
  const n = src.length;
  let ignorableNext = false;
  while (i < n) {
    const ch = src[i];
    if (ch === '{') { flushHex(); pushState(); i++; continue; }
    if (ch === '}') {
      // capture small destination texts before the state ends
      if (st.dest === 'atnauthor') pendingAuthor = (st.text || '').trim();
      if (st.dest === 'atrfstart') { const id = (st.text || '').trim(); openAnn[id] = { block: blocks.length, inline: inl.length }; }
      if (st.dest === 'atrfend') { const id = (st.text || '').trim(); if (openAnn[id]) { annRanges.push({ id: `a${id}`, start: openAnn[id], end: { block: blocks.length, inline: inl.length } }); pendingRef = `a${id}`; } }
      popState(); i++; continue;
    }
    if (ch === '\\') {
      const nx = src[i + 1];
      if (/[a-zA-Z]/.test(nx || '')) {
        let j = i + 1;
        while (j < n && /[a-zA-Z]/.test(src[j])) j++;
        const word = src.slice(i + 1, j);
        let k = j; let hasParam = false; let param = 0;
        if (src[k] === '-' || /[0-9]/.test(src[k] || '')) { const m = /^-?\d+/.exec(src.slice(k, k + 12)); if (m) { hasParam = true; param = parseInt(m[0], 10); k += m[0].length; } }
        if (src[k] === ' ') k++;
        i = k;
        if (word !== 'u' || !hexBuf.length) { if (word !== 'u') flushHex(); }
        if (ignorableNext) { ignorableNext = false; if (!['footnote', 'annotation', 'atnauthor', 'atrfstart', 'atrfend', 'atnid', 'shppict', 'fldinst', 'pn', 'listtext', 'pntext', 'header', 'footer', 'headerr', 'footerr', 'revtbl', 'bkmkstart', 'bkmkend'].includes(word)) { st.dest = 'skip'; continue; } }
        if (word === 'bin' && hasParam) { i += param; continue; }
        ctrl(word, param, hasParam);
        if (word === 'u') continue;
        continue;
      }
      // control symbols
      i += 2;
      if (nx === '\'') {
        const hex = src.substr(i, 2); i += 2;
        if (skipChars > 0) { skipChars--; continue; }
        if (st.dest === 'colortbl' || st.dest === 'skip') continue;
        hexBuf.push(parseInt(hex, 16));
        continue;
      }
      flushHex();
      if (skipChars > 0) { skipChars--; continue; }
      if (nx === '*') { ignorableNext = true; continue; }
      if (nx === '~') { addText(' '); continue; }
      if (nx === '_') { addText('‑'); continue; }
      if (nx === '-') continue;
      if (nx === '\n' || nx === '\r') { if (st.dest === 'body') endPara(); continue; }
      if (nx === '\\' || nx === '{' || nx === '}') { addText(nx); continue; }
      continue;
    }
    if (ch === '\r' || ch === '\n') { i++; continue; }
    // plain text run
    let j = i;
    while (j < n && !'{}\\\r\n'.includes(src[j])) j++;
    let t = src.slice(i, j);
    i = j;
    if (skipChars > 0) { const s = Math.min(skipChars, t.length); t = t.slice(s); skipChars -= s; }
    if (!t) continue;
    if (st.dest === 'colortbl') {
      for (const c of t) if (c === ';') { colors.push(colorEntry.set ? [colorEntry.r, colorEntry.g, colorEntry.b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase() : null); colorEntry = { r: 0, g: 0, b: 0, set: false }; }
      if (colors.length === 1 && colors[0] === null && t.trim().startsWith(';')) { /* auto colour */ }
      continue;
    }
    if (st.dest === 'fonttbl') { if (fontEntry) { fontEntry.name += t; if (fontEntry.name.includes(';')) { fonts[fontEntry.num] = fontEntry.name.split(';')[0].trim(); } } continue; }
    if (st.dest === 'stylesheet') {
      if (styleEntry) {
        styleEntry.name += t;
        if (styleEntry.name.includes(';') && styleEntry.num >= 0) {
          const nm = styleEntry.name.split(';')[0].trim().toLowerCase();
          const hm = /(?:heading|überschrift|titre|titolo|encabezado)\s*(\d)/.exec(nm);
          styles[styleEntry.num] = { name: nm, level: hm ? Number(hm[1]) : styleEntry.outline != null && styleEntry.outline < 6 ? styleEntry.outline + 1 : 0, pstyle: nm === 'title' || nm === 'titel' ? 'title' : nm === 'subtitle' || nm === 'untertitel' ? 'subtitle' : null };
        }
      }
      continue;
    }
    if (st.dest === 'revtbl') { for (const part of t.split(';')) if (part.trim()) revAuthors.push(part.trim()); continue; }
    if (st.dest === 'atnauthor' || st.dest === 'atrfstart' || st.dest === 'atrfend') { st.text = (st.text || '') + t; continue; }
    flushHex();
    addText(t);
  }
  flushHex();
  if (inl.length) endPara();
  if (row) endRow();
  closeTable();

  // revision table: first entry is "Unknown" in Word files
  // colours: the first ";" produced the auto entry already – shift if needed
  if (landscape) model.page.orientation = 'landscape';
  const w = model.page.width; const h = model.page.height;
  model.page.size = sizeName(w, h);
  if (landscape && w > h) { model.page.width = h; model.page.height = w; }
  if (fonts[deff]) model.defaults.font = fonts[deff];

  // comments with ranges → comment marks
  for (const r of annRanges) applyCommentRange(blocks, r);
  // lists
  const out = buildLists(blocks, listParas);
  model.doc = fixDoc({ type: 'doc', content: out });
  if (!model.meta.title) delete model.meta.titleFromFile;
  return model;
}

function applyCommentRange(blocks, r) {
  const mark = { type: 'comment', attrs: { id: r.id } };
  const add = (node, from = 0, to = Infinity) => {
    (node.content || []).forEach((c, k) => { if (k >= from && k < to && c.type === 'text') c.marks = [...(c.marks || []).filter((m) => m.type !== 'comment'), mark]; });
  };
  if (r.lastWord) {
    // anchor to the last text node before the annotation
    const b = blocks[r.block] || blocks[blocks.length - 1];
    if (b && b.content && b.content.length) { const last = [...b.content].reverse().find((c) => c.type === 'text'); if (last) last.marks = [...(last.marks || []), mark]; }
    return;
  }
  // start/end are positions in blocks list + inline index at the time (inline index refers to the unfinished paragraph)
  for (let k = r.start.block; k <= Math.min(r.end.block, blocks.length - 1); k++) {
    const from = k === r.start.block ? r.start.inline : 0;
    const to = k === r.end.block ? r.end.inline : Infinity;
    if (blocks[k]) add(blocks[k], from, to);
  }
}

/** turn flat paragraphs marked as list paragraphs into nested bullet/ordered/task lists */
function buildLists(blocks, listParas) {
  if (!listParas.length) return blocks;
  const byIndex = new Map(listParas.map((l) => [l.index, l]));
  const out = [];
  let k = 0;
  while (k < blocks.length) {
    if (!byIndex.has(k)) { out.push(blocks[k]); k++; continue; }
    const run = [];
    while (k < blocks.length && byIndex.has(k)) { run.push({ node: blocks[k], ...byIndex.get(k) }); k++; }
    out.push(...nest(run, 0));
  }
  return out;
  function listNode(kind) { return kind === 'ordered' ? { type: 'orderedList', attrs: { start: 1 }, content: [] } : kind === 'task' ? { type: 'taskList', content: [] } : { type: 'bulletList', content: [] }; }
  function nest(items, level) {
    const lists = [];
    let cur = null;
    let idx = 0;
    while (idx < items.length) {
      const it = items[idx];
      if (it.level > level && cur && cur.content.length) {
        const sub = [];
        while (idx < items.length && items[idx].level > level) sub.push(items[idx++]);
        const last = cur.content[cur.content.length - 1];
        last.content.push(...nest(sub, level + 1));
        continue;
      }
      if (!cur || kindOf(cur) !== it.kind) { cur = listNode(it.kind); lists.push(cur); }
      const p = it.node.type === 'heading' ? { ...it.node, type: 'paragraph', attrs: { ...it.node.attrs } } : it.node;
      if (p.attrs) { delete p.attrs.level; delete p.attrs.indent; delete p.attrs.firstLine; }
      cur.content.push(it.kind === 'task' ? { type: 'taskItem', attrs: { checked: !!it.checked }, content: [p] } : { type: 'listItem', content: [p] });
      idx++;
    }
    return lists;
  }
  function kindOf(l) { return l.type === 'orderedList' ? 'ordered' : l.type === 'taskList' ? 'task' : 'bullet'; }
}

/* =====================================================================================================
 * Writer
 * ===================================================================================================== */
const rtfEsc = (s) => {
  let o = '';
  for (const ch of String(s)) {
    const c = ch.codePointAt(0);
    if (ch === '\\' || ch === '{' || ch === '}') o += `\\${ch}`;
    else if (ch === '\t') o += '\\tab ';
    else if (ch === '\n') o += '\\line ';
    else if (c < 128) o += ch;
    else if (c < 0x10000) o += `\\u${c > 32767 ? c - 65536 : c}?`;
    else { const hi = Math.floor((c - 0x10000) / 0x400) + 0xd800; const lo = ((c - 0x10000) % 0x400) + 0xdc00; o += `\\u${hi - 65536}?\\u${lo - 65536}?`; }
  }
  return o;
};
const HEAD_SIZE = { 1: 32, 2: 26, 3: 24, 4: 22, 5: 22, 6: 22 }; // half-points

export function writeRtf(model) {
  const fonts = [firstFont((model.defaults && model.defaults.font) || 'Calibri'), 'Symbol', 'Consolas', 'Cambria Math'];
  const colors = ['000000'];
  const authors = ['Unknown'];
  const fontIdx = (f) => { const n = firstFont(f); let k = fonts.indexOf(n); if (k < 0) { fonts.push(n); k = fonts.length - 1; } return k; };
  const colorIdx = (c) => { const h = hexColor(c); if (!h) return 0; let k = colors.indexOf(h); if (k < 0) { colors.push(h); k = colors.length - 1; } return k + 1; };
  const authorIdx = (a) => { const n = a || 'Unknown'; let k = authors.indexOf(n); if (k < 0) { authors.push(n); k = authors.length - 1; } return k; };
  const baseFs = Math.round(((model.defaults && model.defaults.fontSize) || 11) * 2);
  const spaceAfter = Math.round(((model.defaults && model.defaults.spaceAfter) != null ? model.defaults.spaceAfter : 8) * 20);
  const lineMult = Math.round(((model.defaults && model.defaults.lineHeight) || 1.15) * 240);
  const blue = colorIdx('#0563C1');
  const comments = new Map((model.comments || []).map((c) => [String(c.id), c]));
  let annNum = 0;

  const runs = (content) => {
    let s = '';
    let openComments = [];
    const closeComment = (id) => {
      const c = comments.get(id);
      s += `{\\*\\atrfend ${annNum}}`;
      if (c) s += `{\\*\\atnid ${rtfEsc(initials(c.author))}}{\\*\\atnauthor ${rtfEsc(c.author || '')}}\\chatn{\\*\\annotation{\\*\\atnref ${annNum}}\\pard\\plain ${rtfEsc(c.text || '')}}`;
    };
    const list = content || [];
    list.forEach((n, idx) => {
      const ids = n.type === 'text' ? (n.marks || []).filter((m) => m.type === 'comment').map((m) => String(m.attrs.id)) : openComments;
      for (const id of ids) if (!openComments.includes(id)) { annNum++; s += `{\\*\\atrfstart ${annNum}}`; openComments.push(id); }
      if (n.type === 'text') {
        let f = '';
        let link = null;
        for (const m of n.marks || []) {
          const a = m.attrs || {};
          if (m.type === 'bold') f += '\\b'; else if (m.type === 'italic') f += '\\i'; else if (m.type === 'underline') f += '\\ul'; else if (m.type === 'strike') f += '\\strike';
          else if (m.type === 'superscript') f += '\\super'; else if (m.type === 'subscript') f += '\\sub'; else if (m.type === 'code') f += `\\f${fontIdx('Consolas')}`;
          else if (m.type === 'textStyle') { if (a.fontFamily) f += `\\f${fontIdx(a.fontFamily)}`; const pt = ptOf(a.fontSize); if (pt) f += `\\fs${Math.round(pt * 2)}`; if (a.color) f += `\\cf${colorIdx(a.color)}`; }
          else if (m.type === 'highlight') f += `\\highlight${colorIdx(a.color || '#FFFF00')}`;
          else if (m.type === 'link') link = a.href;
          else if (m.type === 'insertion') f += `\\revised\\revauth${authorIdx(a.author)}`;
          else if (m.type === 'deletion') f += `\\deleted\\revauthdel${authorIdx(a.author)}`;
        }
        const body = `{${f}${f ? ' ' : ''}${rtfEsc(n.text)}}`;
        s += link ? `{\\field{\\*\\fldinst{HYPERLINK "${String(link).replace(/"/g, '')}"}}{\\fldrslt{\\ul\\cf${blue} ${body}}}}` : body;
      } else if (n.type === 'hardBreak') s += '\\line ';
      else if (n.type === 'image') s += pictOf(n.attrs);
      else if (n.type === 'footnote') s += `{\\super\\chftn}{\\footnote\\pard\\plain\\fs20{\\super\\chftn} ${rtfEsc(n.attrs.text || '')}}`;
      else if (n.type === 'equation') s += `{\\i\\f${fontIdx('Cambria Math')} ${rtfEsc(equationText(n.attrs.latex || ''))}}`;
      else if (n.type === 'field') { const a = n.attrs || {}; const ins = { page: 'PAGE', pages: 'NUMPAGES', date: 'DATE', time: 'TIME', author: 'AUTHOR', title: 'TITLE' }[a.kind]; s += ins ? `{\\field{\\*\\fldinst ${ins}}{\\fldrslt ${rtfEsc(a.value || '')}}}` : rtfEsc(a.value || ''); }
      else if (n.type === 'contentControl') { const a = n.attrs || {}; s += rtfEsc(a.kind === 'checkbox' ? (a.checked ? '☒' : '☐') : a.value || ''); }
      else if (n.type === 'diagram') { const a = n.attrs || {}; if (a.src) s += pictOf({ src: a.src, width: a.width, height: a.height, alt: a.source }); }
      // close comments that do not continue in the next node
      const next = list[idx + 1];
      const nextIds = next && next.type === 'text' ? (next.marks || []).filter((m) => m.type === 'comment').map((m) => String(m.attrs.id)) : [];
      for (const id of [...openComments]) if (!nextIds.includes(id)) { closeComment(id); openComments = openComments.filter((x) => x !== id); }
    });
    return s;
  };
  const paraFmt = (a = {}, { extraLi = 0 } = {}) => {
    let f = '';
    if (a.textAlign === 'center') f += '\\qc'; else if (a.textAlign === 'right') f += '\\qr'; else if (a.textAlign === 'justify') f += '\\qj';
    const li = Math.round((a.indent || 0) * 20) + extraLi;
    if (li) f += `\\li${li}`;
    if (a.firstLine) f += `\\fi${Math.round(a.firstLine * 20)}`;
    f += `\\sb${Math.round((a.spaceBefore || 0) * 20)}\\sa${a.spaceAfter != null ? Math.round(a.spaceAfter * 20) : spaceAfter}`;
    f += `\\sl${a.lineHeight ? Math.round(a.lineHeight * 240) : lineMult}\\slmult1`;
    return f;
  };
  const base = `\\f0\\fs${baseFs}`;
  const out = [];
  const block = (n, ctx = {}) => {
    const a = n.attrs || {};
    const intbl = ctx.intbl ? '\\intbl' : '';
    switch (n.type) {
      case 'paragraph': {
        const ps = a.pstyle === 'title' ? `\\s15\\fs56` : a.pstyle === 'subtitle' ? `\\s16\\fs30\\cf${colorIdx('#5A5A5A')}` : '';
        out.push(`\\pard\\plain${intbl}${a.pstyle === 'title' ? '\\s15' : a.pstyle === 'subtitle' ? '\\s16' : ''}${paraFmt(a, ctx)}${base}${ps.replace(/\\s1[56]/, '')} ${runs(n.content)}${ctx.noPar ? '' : '\\par'}`);
        break;
      }
      case 'heading': { const l = a.level || 1; out.push(`\\pard\\plain${intbl}\\s${l}\\outlinelevel${l - 1}\\keepn${paraFmt({ spaceBefore: 12, spaceAfter: 4, ...a })}${base}\\b\\fs${HEAD_SIZE[l] || 24}\\cf${colorIdx(l <= 2 ? '#2F5496' : '#1F3763')} ${runs(n.content)}\\par`); break; }
      case 'blockquote': for (const c of n.content || []) block(c, { ...ctx, extraLi: (ctx.extraLi || 0) + 720 }); break;
      case 'callout': table([[{ content: n.content, attrs: { backgroundColor: { warn: '#FDF6E6', ok: '#ECF8F2', note: '#F6F6F6' }[a.kind] || '#EEF4FC' } }]]); break;
      case 'codeBlock': for (const line of (n.content || []).map((c) => c.text).join('').split('\n')) out.push(`\\pard\\plain${intbl}\\li${ctx.extraLi || 0}\\sa0${base}\\f${fontIdx('Consolas')}\\fs20 ${rtfEsc(line)}\\par`); break;
      case 'bulletList': case 'orderedList': case 'taskList': list(n, ctx, 0); break;
      case 'horizontalRule': out.push(`\\pard\\plain${intbl}\\brdrb\\brdrs\\brdrw10\\brsp20\\sa120${base}\\par`); break;
      case 'pageBreak': out.push('\\pard\\plain\\page'); break;
      case 'sectionBreak': if (!a.type || a.type !== 'continuous') out.push('\\pard\\plain\\page'); break;
      case 'toc':
        if (a.title) out.push(`\\pard\\plain\\sa120${base}\\b\\fs28\\cf${colorIdx('#2F5496')} ${rtfEsc(a.title)}\\par`);
        for (const e of a.entries || []) out.push(`\\pard\\plain\\li${((e.level || 1) - 1) * 360}\\sa60${base} ${rtfEsc(e.text)}\\par`);
        break;
      case 'table': table((n.content || []).map((r) => (r.content || []).map((c) => ({ content: c.content, attrs: c.attrs || {}, header: c.type === 'tableHeader' })))); break;
      default: for (const c of n.content || []) block(c, ctx);
    }
  };
  const list = (n, ctx, level) => {
    let num = (n.attrs && n.attrs.start) || 1;
    const ls = n.attrs && n.attrs.listStyle;
    for (const li of n.content || []) {
      (li.content || []).forEach((c, k) => {
        if (k === 0 && (c.type === 'paragraph' || c.type === 'heading')) {
          const lvl = (ctx.listLevel || 0) + level;
          const indent = 720 + lvl * 360 + (ctx.extraLi || 0);
          let pn;
          if (n.type === 'orderedList') {
            const label = numLabel(num, ls, lvl);
            pn = `{\\listtext\\pard\\plain ${rtfEsc(label)}\\tab}{\\*\\pn\\pnlvlbody\\pnf0\\pnindent0\\pnstart${num}${{ 'lower-alpha': '\\pnlcltr', 'upper-alpha': '\\pnucltr', 'lower-roman': '\\pnlcrm', 'upper-roman': '\\pnucrm' }[ls] || '\\pndec'}{\\pntxta .}}`;
            num++;
          } else if (n.type === 'taskList') {
            const box = li.attrs && li.attrs.checked ? '☑' : '☐';
            pn = `{\\listtext\\pard\\plain ${rtfEsc(box)}\\tab}{\\*\\pn\\pnlvlblt\\pnf0\\pnindent0{\\pntxtb ${rtfEsc(box)}}}`;
          } else {
            const bullet = ['•', '◦', '▪'][lvl % 3];
            pn = `{\\listtext\\pard\\plain\\f${fontIdx('Symbol')} \\'b7\\tab}{\\*\\pn\\pnlvlblt\\pnf${fontIdx('Symbol')}\\pnindent0{\\pntxtb\\'b7}}`;
            void bullet;
          }
          out.push(`\\pard\\plain${ctx.intbl ? '\\intbl' : ''}\\ilvl${lvl}${pn}\\fi-360\\li${indent}${paraFmt({ ...c.attrs, indent: 0, firstLine: 0, spaceAfter: c.attrs && c.attrs.spaceAfter != null ? c.attrs.spaceAfter : 2 })}${base} ${runs(c.content)}\\par`);
        } else if (['bulletList', 'orderedList', 'taskList'].includes(c.type)) list(c, ctx, level + 1);
        else block(c, { ...ctx, extraLi: (ctx.extraLi || 0) + 720 + level * 360 });
      });
    }
  };
  const table = (rows) => {
    // grid with spans so merged cells become \clmgf/\clmrg (horizontal) and \clvmgf/\clvmrg (vertical)
    const box = pageBox(model.page);
    const usable = mmToTwip(box.width - model.page.margins.left - model.page.margins.right);
    const grid = [];
    rows.forEach((r, ri) => {
      grid[ri] = grid[ri] || [];
      let ci = 0;
      for (const c of r) {
        while (grid[ri][ci]) ci++;
        const cs = (c.attrs && c.attrs.colspan) || 1; const rs = (c.attrs && c.attrs.rowspan) || 1;
        for (let y = 0; y < rs; y++) for (let x = 0; x < cs; x++) { grid[ri + y] = grid[ri + y] || []; grid[ri + y][ci + x] = { cell: c, first: y === 0 && x === 0, h: x > 0, v: y > 0, vStart: y === 0 && rs > 1 }; }
        ci += cs;
      }
    });
    const cols = Math.max(...grid.map((r) => r.length));
    const widths = [];
    for (let k = 0; k < cols; k++) {
      const g = grid[0] && grid[0][k];
      const cw = g && g.cell.attrs && g.cell.attrs.colwidth;
      widths.push(cw && cw[0] && !g.h ? Math.round(cw[0] * 15) : null);
    }
    const known = widths.filter(Boolean).reduce((s, w) => s + w, 0);
    const free = widths.filter((w) => !w).length;
    const each = free ? Math.max(600, Math.round((usable - known) / free)) : 0;
    grid.forEach((r) => {
      let x = 0;
      let def = '\\trowd\\trgaph108\\trleft0';
      for (let k = 0; k < cols; k++) {
        const g = r[k];
        x += widths[k] || each;
        const bg = g && g.cell.attrs && g.cell.attrs.backgroundColor ? `\\clcbpat${colorIdx(g.cell.attrs.backgroundColor)}` : '';
        const merge = g ? (g.h ? '\\clmrg' : g.cell.attrs && g.cell.attrs.colspan > 1 ? '\\clmgf' : '') + (g.v ? '\\clvmrg' : g.vStart ? '\\clvmgf' : '') : '';
        def += `${merge}\\clvertalt${bg}\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10\\cellx${x}`;
      }
      out.push(def);
      for (let k = 0; k < cols; k++) {
        const g = r[k];
        if (!g || !g.first) { out.push('\\pard\\plain\\intbl\\cell'); continue; }
        const content = g.cell.content && g.cell.content.length ? g.cell.content : [{ type: 'paragraph' }];
        content.forEach((c, idx) => {
          const before = out.length;
          block(c.type === 'paragraph' && g.cell.header ? { ...c, content: (c.content || []).map((t) => (t.type === 'text' ? { ...t, marks: [...(t.marks || []), { type: 'bold' }] } : t)) } : c, { intbl: true });
          if (idx === content.length - 1) {
            // last paragraph of the cell ends with \cell instead of \par
            const last = out.length - 1;
            if (last >= before && out[last].endsWith('\\par')) out[last] = `${out[last].slice(0, -4)}\\cell`; else out.push('\\pard\\plain\\intbl\\cell');
          }
        });
      }
      out.push('\\row');
    });
  };
  const pictOf = (a) => {
    const d = parseDataUrl(a.src);
    if (!d || !['image/png', 'image/jpeg'].includes(d.mime)) return '';
    const sz = imageSize(d.bytes) || { width: 300, height: 200 };
    const w = a.width || Math.min(sz.width, 600);
    const h = Math.round(w * (sz.height / sz.width));
    let hex = '';
    for (let k = 0; k < d.bytes.length; k++) { hex += d.bytes[k].toString(16).padStart(2, '0'); if (k % 64 === 63) hex += '\n'; }
    return `{\\pict${d.mime === 'image/png' ? '\\pngblip' : '\\jpegblip'}\\picw${sz.width}\\pich${sz.height}\\picwgoal${Math.round(w * 15)}\\pichgoal${Math.round(h * 15)}\n${hex}}`;
  };

  for (const n of (model.doc && model.doc.content) || []) block(n);
  // header/footer
  const hf = (x, kind) => {
    if (!x || !x.text) return '';
    const parts = String(x.text).split(/(\{page\}|\{pages\})/).map((t) => (t === '{page}' ? '{\\field{\\*\\fldinst PAGE}{\\fldrslt 1}}' : t === '{pages}' ? '{\\field{\\*\\fldinst NUMPAGES}{\\fldrslt 1}}' : rtfEsc(t.replace(/\{title\}/g, (model.meta && model.meta.title) || '').replace(/\{date\}/g, new Date().toLocaleDateString())))).join('');
    return `{\\${kind}\\pard\\plain${x.align === 'left' ? '' : x.align === 'right' ? '\\qr' : '\\qc'}\\f0\\fs18 ${parts}\\par}`;
  };
  const p = model.page || {};
  const box = pageBox(p);
  const m = p.margins || {};
  const meta = model.meta || {};
  const head = [
    '{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1031',
    `{\\fonttbl${fonts.map((f, k) => `{\\f${k}\\fnil\\fcharset${f === 'Symbol' ? 2 : 0} ${rtfEsc(f)};}`).join('')}}`,
    `{\\colortbl ;${colors.map((c) => `\\red${parseInt(c.slice(0, 2), 16)}\\green${parseInt(c.slice(2, 4), 16)}\\blue${parseInt(c.slice(4, 6), 16)};`).join('')}}`,
    `{\\stylesheet{\\s0\\f0\\fs${baseFs} Normal;}${[1, 2, 3, 4, 5, 6].map((l) => `{\\s${l}\\outlinelevel${l - 1}\\b\\fs${HEAD_SIZE[l]} heading ${l};}`).join('')}{\\s15\\fs56 Title;}{\\s16\\fs30 Subtitle;}}`,
    authors.length > 1 ? `{\\*\\revtbl ${authors.map((x) => `{${rtfEsc(x)};}`).join('')}}` : '',
    `{\\info${meta.title ? `{\\title ${rtfEsc(meta.title)}}` : ''}${meta.author ? `{\\author ${rtfEsc(meta.author)}}` : ''}${meta.subject ? `{\\subject ${rtfEsc(meta.subject)}}` : ''}}`,
    `\\paperw${mmToTwip(box.width)}\\paperh${mmToTwip(box.height)}\\margl${mmToTwip(m.left)}\\margr${mmToTwip(m.right)}\\margt${mmToTwip(m.top)}\\margb${mmToTwip(m.bottom)}${p.orientation === 'landscape' ? '\\landscape' : ''}${p.columns > 1 ? `\\cols${p.columns}` : ''}${model.titlePage ? '\\titlepg' : ''}\\ftnbj\\fet0\\viewkind4\\uc1`, // \ftnbj: without it Word drops every \footnote (verified in Word 16)
    hf(model.header, 'header'), hf(model.footer, 'footer'),
  ].filter(Boolean);
  return `${head.join('\n')}\n${out.join('\n')}\n}`;
}

function numLabel(n, ls, lvl) {
  if (ls === 'lower-alpha') return `${alpha(n)}.`;
  if (ls === 'upper-alpha') return `${alpha(n).toUpperCase()}.`;
  if (ls === 'lower-roman') return `${roman(n).toLowerCase()}.`;
  if (ls === 'upper-roman') return `${roman(n)}.`;
  void lvl;
  return `${n}.`;
}
export const alpha = (n) => { let s = ''; let x = n; while (x > 0) { x--; s = String.fromCharCode(97 + (x % 26)) + s; x = Math.floor(x / 26); } return s; };
export const roman = (n) => { const r = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let s = ''; let x = n; for (const [v, c] of r) while (x >= v) { s += c; x -= v; } return s; };
const initials = (name) => String(name || '?').split(/\s+/).filter(Boolean).map((w) => w[0]).join('').slice(0, 3).toUpperCase() || '?';

/** plain text of an RTF file (search/preview) */
export function rtfToText(input) {
  const m = parseRtf(input);
  const lines = [];
  walk(m.doc, (n) => { if (n.type === 'paragraph' || n.type === 'heading') { lines.push(inlineText(n)); return false; } return undefined; });
  return lines.join('\n');
}
