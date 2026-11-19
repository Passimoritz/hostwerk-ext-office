/*
 * Tiny, dependency-free XML reader/writer that works in the browser AND in Node (unit tests) – used by the slides
 * chunk (pptx/odp) and free for anyone. Not validating; namespaces are kept as prefixes ("p:sp").
 *
 *   parseXml(str)               → root element { n: 'p:sld', a: { …attrs }, c: [ element | 'text' ] }
 *   kids(el, name?)             → child elements (optionally only those named `name`, '*' = all)
 *   kid(el, name)               → first child element with that name (or null)
 *   path(el, 'a/b/c')           → walk first matching children
 *   all(el, name)               → every descendant element with that name (depth first)
 *   text(el)                    → concatenated text content
 *   esc(s) / escAttr(s)         → escaping for writing
 *   tag(name, attrs, inner)     → '<name a="…">inner</name>' (inner: string | array | null → self-closing)
 */

const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeEntities(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, e) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(cp); } catch { return m; }
    }
    return ENT[e] !== undefined ? ENT[e] : m;
  });
}

export function parseXml(src) {
  const s = typeof src === 'string' ? src : new TextDecoder('utf-8').decode(src);
  const root = { n: '#doc', a: {}, c: [] };
  const stack = [root];
  let i = 0;
  const len = s.length;
  const attrRe = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  while (i < len) {
    const lt = s.indexOf('<', i);
    if (lt < 0) { pushText(stack[stack.length - 1], s.slice(i)); break; }
    if (lt > i) pushText(stack[stack.length - 1], s.slice(i, lt));
    const c1 = s[lt + 1];
    if (c1 === '?') { const e = s.indexOf('?>', lt); i = e < 0 ? len : e + 2; continue; }
    if (c1 === '!') {
      if (s.startsWith('<!--', lt)) { const e = s.indexOf('-->', lt); i = e < 0 ? len : e + 3; continue; }
      if (s.startsWith('<![CDATA[', lt)) {
        const e = s.indexOf(']]>', lt);
        stack[stack.length - 1].c.push(s.slice(lt + 9, e < 0 ? len : e));
        i = e < 0 ? len : e + 3; continue;
      }
      // DOCTYPE (possibly with internal subset)
      let depth = 0; let j = lt;
      for (; j < len; j++) { if (s[j] === '[') depth++; else if (s[j] === ']') depth--; else if (s[j] === '>' && depth <= 0) break; }
      i = j + 1; continue;
    }
    if (c1 === '/') {
      const e = s.indexOf('>', lt);
      const name = s.slice(lt + 2, e).trim();
      // pop to the matching element (tolerant to broken nesting)
      for (let k = stack.length - 1; k > 0; k--) { if (stack[k].n === name) { stack.length = k; break; } }
      i = e + 1; continue;
    }
    // start tag – find the closing '>' outside of quotes
    let j = lt + 1; let q = null;
    for (; j < len; j++) { const ch = s[j]; if (q) { if (ch === q) q = null; } else if (ch === '"' || ch === "'") q = ch; else if (ch === '>') break; }
    let body = s.slice(lt + 1, j);
    const self = body.endsWith('/');
    if (self) body = body.slice(0, -1);
    const sp = body.search(/[\s]/);
    const name = sp < 0 ? body : body.slice(0, sp);
    const el = { n: name, a: {}, c: [] };
    if (sp >= 0) {
      const rest = body.slice(sp);
      attrRe.lastIndex = 0;
      let m;
      while ((m = attrRe.exec(rest))) el.a[m[1]] = decodeEntities(m[3] !== undefined ? m[3] : m[4]);
    }
    stack[stack.length - 1].c.push(el);
    if (!self) stack.push(el);
    i = j + 1;
  }
  return root.c.find((x) => typeof x !== 'string') || null;
}

function pushText(parent, raw) {
  if (!raw) return;
  parent.c.push(decodeEntities(raw));
}

export const isEl = (x) => x && typeof x === 'object';
export function kids(el, name) {
  if (!el || !el.c) return [];
  const out = [];
  for (const x of el.c) if (typeof x !== 'string' && (!name || name === '*' || x.n === name)) out.push(x);
  return out;
}
export function kid(el, name) {
  if (!el || !el.c) return null;
  for (const x of el.c) if (typeof x !== 'string' && x.n === name) return x;
  return null;
}
export function path(el, p) {
  let cur = el;
  for (const part of p.split('/')) { if (!cur) return null; cur = kid(cur, part); }
  return cur;
}
export function all(el, name, out = []) {
  if (!el || !el.c) return out;
  for (const x of el.c) {
    if (typeof x === 'string') continue;
    if (x.n === name) out.push(x);
    all(x, name, out);
  }
  return out;
}
export function text(el) {
  if (el == null) return '';
  if (typeof el === 'string') return el;
  let s = '';
  for (const x of el.c || []) s += typeof x === 'string' ? x : text(x);
  return s;
}
/** local name without prefix */
export const local = (n) => { const i = n.indexOf(':'); return i < 0 ? n : n.slice(i + 1); };

export const esc = (s) => String(s == null ? '' : s).replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'))
  // strip characters XML 1.0 does not allow
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
export const escAttr = (s) => esc(s).replace(/"/g, '&quot;').replace(/\n/g, '&#10;');

export function attrs(a) {
  let s = '';
  if (!a) return s;
  for (const k of Object.keys(a)) { const v = a[k]; if (v === undefined || v === null || v === false) continue; s += ` ${k}="${escAttr(v === true ? 'true' : v)}"`; }
  return s;
}
export function tag(name, a, inner) {
  const body = Array.isArray(inner) ? inner.filter((x) => x != null && x !== false).join('') : inner;
  if (body === null || body === undefined || body === '') return `<${name}${attrs(a)}/>`;
  return `<${name}${attrs(a)}>${body}</${name}>`;
}

/** serialize a parsed element back to XML (used to round-trip unknown fragments) */
export function serialize(el) {
  if (typeof el === 'string') return esc(el);
  return tag(el.n, el.a, el.c && el.c.length ? el.c.map(serialize).join('') : null);
}

export const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
