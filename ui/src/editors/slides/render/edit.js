/*
 * In-place text editing (browser only): model body → contentEditable HTML with data attributes, and back.
 * Formatting while editing uses document.execCommand (styleWithCSS) plus small fix-ups; the parser compares each
 * span's inline style with the style its data-r run would render, so only real changes override the run props.
 */
import { paraProps, runProps, paraCss, runCss, bulletLabels, runText } from './text.js';
import { css, resolveHex, resolveFont } from '../themes.js';
import { S } from '../units.js';

const escH = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const kebab = (k) => k.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`);
const styleStr = (o) => Object.entries(o).filter(([, v]) => v !== undefined && v !== null && v !== '').map(([k, v]) => `${kebab(k)}:${String(v).replace(/"/g, "'")}`).join(';');
const PARA_KEYS = ['level', 'align', 'bullet', 'lnSpc', 'spcBef', 'spcAft', 'marL', 'indent', 'endProps', 'rtl'];

function paraAttrs(p, body, theme, label) {
  const pp = paraProps(p, body, theme);
  const data = {}; for (const k of PARA_KEYS) if (p[k] !== undefined) data[k] = p[k];
  const pc = paraCss(pp);
  const first = runProps((p.runs || [])[0] || p.endProps || {}, p, body);
  const st = { ...pc, margin: 0, marginTop: pc.marginTop, marginBottom: pc.marginBottom, '--bi': `${Math.max(0, -S(pp.indent))}px`, '--bc': css((label && label.color) || first.color || 'tx1', theme), '--bs': `${first.size}px` };
  return `data-p="${escH(JSON.stringify(data))}"${label ? ` data-bl="${escH(label.text)}"` : ''} style="${escH(styleStr(st))}"`;
}

export function bodyToHtml(body, theme, ctx = {}) {
  const labels = bulletLabels({ ...body, paras: body.paras.map((p) => ({ ...p, forceBullet: true })) }, theme);
  return body.paras.map((p, i) => {
    const runs = (p.runs || []).filter((r) => r.text || r.field);
    const inner = runs.map((r) => {
      const { text, ...props } = r;
      const rp = runProps(r, p, body);
      return `<span data-r="${escH(JSON.stringify(props))}" style="${escH(styleStr(runCss(rp, theme)))}">${escH(runText(r, ctx)).replace(/\n/g, '<br>')}</span>`;
    }).join('');
    const end = runProps(p.endProps || (p.runs || [])[0] || {}, p, body);
    return `<div ${paraAttrs(p, body, theme, labels[i])}>${inner || `<span data-r="${escH(JSON.stringify(p.endProps || {}))}" style="${escH(styleStr(runCss(end, theme)))}"><br></span>`}</div>`;
  }).join('');
}

/* ---------------------------------------------------------------- parse back */
let probe = null;
const norm = (prop, value) => {
  if (!probe) probe = document.createElement('span');
  probe.style.cssText = '';
  probe.style[prop] = value;
  return probe.style[prop];
};
const rgbToHex = (s) => {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(s || '');
  if (!m) return /^#/.test(s || '') ? s.toUpperCase() : null;
  if (m[4] !== undefined && +m[4] === 0) return 'transparent';
  return `#${[m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('')}`.toUpperCase();
};
const firstFamily = (f) => String(f || '').split(',')[0].trim().replace(/^['"]|['"]$/g, '');

function applyElement(el, props, theme, pctx) {
  const tag = el.tagName;
  if (tag === 'B' || tag === 'STRONG') props.b = true;
  if (tag === 'I' || tag === 'EM') props.i = true;
  if (tag === 'U') props.u = true;
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') props.s = true;
  if (tag === 'SUP') props.baseline = 30;
  if (tag === 'SUB') props.baseline = -25;
  if (tag === 'FONT') { if (el.getAttribute('color')) props.color = rgbToHex(el.getAttribute('color')) || el.getAttribute('color'); if (el.getAttribute('face')) props.font = el.getAttribute('face'); }
  if (tag === 'A' && el.getAttribute('href')) props.link = el.getAttribute('href');
  let expected = {};
  if (el.dataset && el.dataset.r !== undefined) {
    try { Object.assign(props, JSON.parse(el.dataset.r)); } catch { /* ignore */ }
    expected = runCss(runProps(props, pctx && pctx.p, pctx && pctx.body), theme);
  }
  const st = el.style;
  if (!st || !st.length) { if (el.dataset && el.dataset.color) props.color = el.dataset.color; return; }
  const differs = (prop) => st[prop] && st[prop] !== (expected[prop] !== undefined ? norm(prop, String(expected[prop])) : '');
  if (differs('fontWeight')) props.b = st.fontWeight === 'bold' || +st.fontWeight >= 600;
  if (differs('fontStyle')) props.i = st.fontStyle === 'italic';
  const deco = st.textDecorationLine || st.textDecoration;
  if (deco && deco !== norm('textDecorationLine', expected.textDecoration || 'none')) { props.u = /underline/.test(deco); props.s = /line-through/.test(deco); }
  if (el.dataset && el.dataset.color) props.color = el.dataset.color;
  else if (differs('color')) { const hx = rgbToHex(st.color); if (hx && hx !== 'transparent' && hx !== resolveHex(props.color || 'tx1', theme)) props.color = hx; }
  if (el.dataset && el.dataset.hl) props.highlight = el.dataset.hl === 'none' ? undefined : el.dataset.hl;
  else if (differs('backgroundColor')) { const hx = rgbToHex(st.backgroundColor); props.highlight = hx && hx !== 'transparent' ? hx : undefined; }
  if (differs('fontFamily')) { const f = firstFamily(st.fontFamily); if (f && f !== resolveFont(props.font, theme)) props.font = f === resolveFont('+mj', theme) ? '+mj' : f === resolveFont('+mn', theme) ? '+mn' : f; }
  if (differs('fontSize') && /px$/.test(st.fontSize)) props.size = Math.round(parseFloat(st.fontSize) * 2) / 2;
  if (differs('verticalAlign')) props.baseline = st.verticalAlign === 'super' ? 30 : st.verticalAlign === 'sub' ? -25 : undefined;
}

function collectRuns(pdiv, theme, pctx) {
  const runs = [];
  const push = (text, props) => {
    const last = runs[runs.length - 1];
    const clean = {}; for (const [k, v] of Object.entries(props)) if (v !== undefined && v !== false && v !== null && k !== 'text') clean[k] = v;
    if (clean.field) { runs.push({ ...clean, text }); return; }
    if (last && !last.field && JSON.stringify({ ...last, text: '' }) === JSON.stringify({ ...clean, text: '' })) last.text += text;
    else runs.push({ ...clean, text });
  };
  const walkNode = (node, props) => {
    for (const ch of node.childNodes) {
      if (ch.nodeType === 3) { const tx = ch.nodeValue.replace(/​/g, ''); if (tx) push(tx, props); continue; }
      if (ch.nodeType !== 1) continue;
      if (ch.tagName === 'BR') {
        // trailing <br> is the browser's placeholder of an empty line
        const isLast = !nextMeaningful(ch, pdiv);
        if (!isLast) push('\n', props);
        continue;
      }
      if (ch.tagName === 'DIV' || ch.tagName === 'P') { walkNode(ch, props); continue; }
      const np = { ...props };
      applyElement(ch, np, theme, pctx);
      walkNode(ch, np);
    }
  };
  walkNode(pdiv, {});
  return runs;
}
function nextMeaningful(node, stop) {
  let n = node;
  while (n && n !== stop) {
    let s = n.nextSibling;
    while (s) { if ((s.nodeType === 3 && s.nodeValue.replace(/​/g, '')) || (s.nodeType === 1 && (s.tagName === 'BR' || s.textContent))) return s; s = s.nextSibling; }
    n = n.parentNode;
  }
  return null;
}

/** contentEditable root → body (keeps body-level props of `base`) */
export function htmlToBody(root, base, theme) {
  const paras = [];
  let fallback = {};
  const blocks = [];
  let loose = null;
  for (const ch of root.childNodes) {
    if (ch.nodeType === 1 && (ch.tagName === 'DIV' || ch.tagName === 'P')) { loose = null; blocks.push(ch); } else {
      if (!loose) { loose = document.createElement('div'); blocks.push(loose); }
      loose.appendChild(ch.cloneNode(true));
    }
  }
  for (const b of blocks) {
    let props = {};
    try { props = b.dataset && b.dataset.p ? JSON.parse(b.dataset.p) : { ...fallback }; } catch { props = { ...fallback }; }
    fallback = props;
    const pctx = { p: { ...props, runs: [] }, body: base };
    const runs = collectRuns(b, theme, pctx);
    const p = { ...props, runs, level: props.level || 0 };
    if (!runs.length) {
      // keep formatting of an empty paragraph (typed text will use it)
      const sp = b.querySelector('[data-r]');
      if (sp) { const np = {}; applyElement(sp, np, theme, pctx); p.endProps = Object.fromEntries(Object.entries(np).filter(([k, v]) => k !== 'text' && v !== undefined)); }
    }
    paras.push(p);
  }
  if (!paras.length) paras.push({ runs: [], level: 0 });
  return { ...base, paras };
}

/** refresh bullet labels + paragraph styles after a change (Enter copies data-p, numbering needs recounting) */
export function refreshParas(root, base, theme) {
  const body = htmlToBody(root, base, theme);
  const labels = bulletLabels({ ...body, paras: body.paras.map((p) => ({ ...p, forceBullet: true })) }, theme);
  const divs = [...root.children].filter((d) => d.tagName === 'DIV' || d.tagName === 'P');
  divs.forEach((d, i) => {
    const l = labels[i];
    if (l) { if (d.dataset.bl !== l.text) d.dataset.bl = l.text; } else if (d.dataset.bl !== undefined) delete d.dataset.bl;
  });
  return body;
}

/* ---------------------------------------------------------------- commands while editing */
export function saveRange(root) {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return null;
  const r = sel.getRangeAt(0);
  return root.contains(r.commonAncestorContainer) ? r.cloneRange() : null;
}
export function restoreRange(root, range) {
  if (!range) return;
  root.focus({ preventScroll: true });
  const sel = window.getSelection();
  sel.removeAllRanges(); sel.addRange(range);
}
/** paragraph divs touched by the selection (all when nothing is selected inside) */
export function selectedParas(root) {
  const divs = [...root.children].filter((d) => d.tagName === 'DIV' || d.tagName === 'P');
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return divs;
  const r = sel.getRangeAt(0);
  if (!root.contains(r.commonAncestorContainer)) return divs;
  const hit = divs.filter((d) => r.intersectsNode(d));
  return hit.length ? hit : divs;
}
export function setParaProps(root, patchFn, theme, base) {
  for (const d of selectedParas(root)) {
    let p = {}; try { p = JSON.parse(d.dataset.p || '{}'); } catch { /* ignore */ }
    const np = { ...p, ...patchFn(p) };
    for (const k of Object.keys(np)) if (np[k] === undefined) delete np[k];
    d.dataset.p = JSON.stringify(np);
    const pp = paraProps({ ...np, runs: [] }, base, theme);
    const pc = paraCss(pp);
    d.style.textAlign = pc.textAlign; d.style.lineHeight = pc.lineHeight; d.style.marginTop = pc.marginTop; d.style.marginBottom = pc.marginBottom;
    d.style.paddingLeft = pc.paddingLeft; d.style.textIndent = pc.textIndent; d.style.setProperty('--bi', `${Math.max(0, -S(pp.indent))}px`);
  }
}

const SENTINEL = { color: 'rgb(254, 1, 2)', hex: '#FE0102' };
/** run-level formatting of the current selection */
export function execRun(root, cmd, value, theme) {
  try { document.execCommand('styleWithCSS', false, true); } catch { /* old engines */ }
  const sel = window.getSelection();
  const collapsed = !sel || !sel.rangeCount || sel.getRangeAt(0).collapsed;
  if (collapsed && cmd !== 'removeFormat') {
    // no selection: apply to the word under the caret
    try { sel.modify('move', 'backward', 'word'); sel.modify('extend', 'forward', 'word'); } catch { /* ignore */ }
  }
  if (cmd === 'b') document.execCommand('bold');
  else if (cmd === 'i') document.execCommand('italic');
  else if (cmd === 'u') document.execCommand('underline');
  else if (cmd === 's') document.execCommand('strikeThrough');
  else if (cmd === 'sup') document.execCommand('superscript');
  else if (cmd === 'sub') document.execCommand('subscript');
  else if (cmd === 'removeFormat') document.execCommand('removeFormat');
  else if (cmd === 'font') document.execCommand('fontName', false, resolveFont(value, theme));
  else if (cmd === 'color' || cmd === 'highlight') {
    document.execCommand(cmd === 'color' ? 'foreColor' : 'hiliteColor', false, SENTINEL.hex);
    const prop = cmd === 'color' ? 'color' : 'backgroundColor';
    root.querySelectorAll('*').forEach((el) => {
      if (el.style && norm(prop, el.style[prop]) === SENTINEL.color) {
        if (!value) { el.style[prop] = cmd === 'highlight' ? 'transparent' : ''; if (cmd === 'highlight') el.dataset.hl = 'none'; } else { el.style[prop] = css(value, theme); el.dataset[cmd === 'color' ? 'color' : 'hl'] = value; }
      }
      if (el.tagName === 'FONT' && cmd === 'color' && /fe0102/i.test(el.getAttribute('color') || '')) { el.removeAttribute('color'); el.style.color = css(value, theme); el.dataset.color = value; }
    });
  } else if (cmd === 'size') {
    document.execCommand('fontSize', false, '7');
    root.querySelectorAll('font[size="7"]').forEach((f) => { const sp = document.createElement('span'); sp.style.fontSize = `${value}px`; while (f.firstChild) sp.appendChild(f.firstChild); f.replaceWith(sp); });
    root.querySelectorAll('*').forEach((el) => { if (el.style && el.style.fontSize === 'xxx-large') el.style.fontSize = `${value}px`; });
  } else if (cmd === 'link') {
    if (value) document.execCommand('createLink', false, value); else document.execCommand('unlink');
  }
}
/** current run props at the caret (for the ribbon state) */
export function caretRunProps(root, base, theme) {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return null;
  let n = sel.getRangeAt(0).startContainer;
  if (n.nodeType === 3) n = n.parentNode;
  if (!root.contains(n)) return null;
  const chain = [];
  while (n && n !== root) { chain.unshift(n); n = n.parentNode; }
  const props = {};
  let pdiv = null;
  for (const el of chain) if (el.dataset && el.dataset.p) pdiv = el;
  let p = {}; try { p = pdiv ? JSON.parse(pdiv.dataset.p) : {}; } catch { /* ignore */ }
  for (const el of chain) if (el !== pdiv) applyElement(el, props, theme, { p: { ...p, runs: [] }, body: base });
  return { run: runProps(props, { ...p, runs: [] }, base), para: paraProps({ ...p, runs: [] }, base, theme), level: p.level || 0 };
}
