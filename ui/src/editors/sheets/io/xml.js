/* Small string-based XML helpers (fast, no DOM – also run in Node tests). */
export const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
export const escText = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  // characters XML 1.0 can't hold → Excel's _xHHHH_ escape
  .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/g, (ch) => `_x${ch.charCodeAt(0).toString(16).padStart(4, '0').toUpperCase()}_`);
export const unesc = (s) => (s.indexOf('&') < 0 ? s : s.replace(/&(lt|gt|quot|apos|amp|#x[0-9a-fA-F]+|#\d+);/g, (m, e) => (e === 'lt' ? '<' : e === 'gt' ? '>' : e === 'quot' ? '"' : e === 'apos' ? "'" : e === 'amp' ? '&' : e[1] === 'x' ? String.fromCodePoint(parseInt(e.slice(2), 16)) : String.fromCodePoint(+e.slice(1)))));
export const unx = (s) => (s.indexOf('_x') < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16))));

const ATTR = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
export function attrs(tag) {
  const o = {};
  ATTR.lastIndex = 0;
  let m;
  while ((m = ATTR.exec(tag))) { const k = m[1]; const v = unesc(m[2] ?? m[3]); o[k] = v; }
  return o;
}

/**
 * All elements <tag …>…</tag> / <tag …/> (tag without prefix; any prefix matches). Not for self-nesting tags
 * unless `nested` (slower depth counting).
 */
export function elements(xml, tag, nested = false) {
  const out = [];
  const re = new RegExp(`<((?:[\\w.-]+:)?${tag})(?=[\\s/>])([^>]*?)(/?)>`, 'g');
  let m;
  while ((m = re.exec(xml))) {
    const name = m[1];
    if (m[3] === '/') { out.push({ name, a: m[2], inner: '', start: m.index, end: re.lastIndex }); continue; }
    const close = `</${name}>`;
    let end;
    if (!nested) end = xml.indexOf(close, re.lastIndex);
    else {
      let depth = 1; let i = re.lastIndex;
      const open = new RegExp(`<${name.replace('.', '\\.')}(?=[\\s/>])[^>]*?(/?)>`, 'g');
      for (;;) {
        const ci = xml.indexOf(close, i);
        if (ci < 0) { end = -1; break; }
        open.lastIndex = i;
        let om;
        while ((om = open.exec(xml)) && om.index < ci) if (om[1] !== '/') depth++;
        depth--;
        if (!depth) { end = ci; break; }
        i = ci + close.length;
      }
    }
    if (end < 0) break;
    out.push({ name, a: m[2], inner: xml.slice(re.lastIndex, end), start: m.index, end: end + close.length });
    re.lastIndex = end + close.length;
  }
  return out;
}
export const first = (xml, tag, nested) => { const e = elements(xml, tag, nested); return e.length ? e[0] : null; };
export const attrOf = (el) => (el ? attrs(el.a) : {});

/** direct children of the root element: [{ name (local), xml }] plus { head, tail } */
export function topChildren(xml) {
  const rootM = /<([\w:.-]+)(?=[\s>])[^>]*>/.exec(xml.replace(/<\?[\s\S]*?\?>/g, (s) => ' '.repeat(s.length)));
  if (!rootM) return null;
  const startBody = rootM.index + rootM[0].length;
  const rootName = rootM[1];
  const endBody = xml.lastIndexOf(`</${rootName}>`);
  const kids = [];
  const re = /<(\/?)([\w:.-]+)([^>]*?)(\/?)>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g;
  re.lastIndex = startBody;
  let depth = 0; let cur = -1; let curName = null;
  let m;
  while ((m = re.exec(xml)) && m.index < endBody) {
    if (!m[2]) continue;
    if (m[1] === '/') { depth--; if (depth === 0) { kids.push({ name: curName.replace(/^[\w.-]+:/, ''), xml: xml.slice(cur, re.lastIndex) }); cur = -1; } continue; }
    if (depth === 0) { cur = m.index; curName = m[2]; }
    if (m[4] === '/') { if (depth === 0) { kids.push({ name: curName.replace(/^[\w.-]+:/, ''), xml: xml.slice(cur, re.lastIndex) }); cur = -1; } continue; }
    depth++;
  }
  return { head: xml.slice(0, startBody), tail: xml.slice(endBody), kids, rootName };
}

/** concatenated text of all <t> runs (shared strings, comments); skips phonetic runs */
export function textOf(inner) {
  const x = inner.replace(/<(\w+:)?rPh\b[\s\S]*?<\/(\w+:)?rPh>/g, '');
  let s = '';
  const re = /<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>|<(?:\w+:)?t\s*\/>/g;
  let m;
  while ((m = re.exec(x))) if (m[1]) s += m[1];
  return unx(unesc(s));
}

export const dirOf = (p) => p.slice(0, p.lastIndexOf('/') + 1);
export function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(base) + target).split('/'); const out = [];
  for (const x of parts) { if (x === '..') out.pop(); else if (x !== '.') out.push(x); }
  return out.join('/');
}
export const relsPathOf = (p) => `${dirOf(p)}_rels/${p.slice(p.lastIndexOf('/') + 1)}.rels`;
export function relTarget(from, to) {
  const a = dirOf(from).split('/').filter(Boolean); const b = to.split('/');
  let i = 0; while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return `${'../'.repeat(a.length - i)}${b.slice(i).join('/')}`;
}
export function parseRels(xml, part) {
  const out = [];
  if (!xml) return out;
  for (const m of xml.match(/<Relationship\b[^>]*>/g) || []) {
    const a = attrs(m);
    out.push({ id: a.Id, type: a.Type, target: a.Target, mode: a.TargetMode, path: a.TargetMode === 'External' ? null : resolvePath(part, a.Target), tag: m });
  }
  return out;
}
export const REL = {
  sheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  sst: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  calcChain: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  vml: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  metadata: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata',
  officeDoc: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  vba: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
  pivot: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
  table: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
};
export function addRel(relsXml, id, type, target, external = false) {
  const x = relsXml || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const tag = `<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"${external ? ' TargetMode="External"' : ''}/>`;
  if (/<Relationships[^>]*\/>\s*$/.test(x)) return x.replace(/<Relationships([^>]*)\/>\s*$/, `<Relationships$1>${tag}</Relationships>`);
  return x.replace(/<\/Relationships>\s*$/, `${tag}</Relationships>`);
}
export function nextRid(relsXml) { let n = 0; for (const m of (relsXml || '').matchAll(/Id="rId(\d+)"/g)) n = Math.max(n, +m[1]); return `rId${n + 1}`; }
export function removeRels(relsXml, pred) {
  if (!relsXml) return relsXml;
  return relsXml.replace(/<Relationship\b[^>]*>/g, (tag) => (pred(attrs(tag)) ? '' : tag));
}
export function ctOverride(ct, part, type) {
  if (ct.includes(`PartName="/${part}"`)) return ct.replace(new RegExp(`<Override PartName="/${part.replace(/[.[\]]/g, '\\$&')}"[^>]*/>`), `<Override PartName="/${part}" ContentType="${type}"/>`);
  return ct.replace(/<\/Types>\s*$/, `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
}
export function ctDefault(ct, ext, type) {
  if (new RegExp(`<Default Extension="${ext}"`, 'i').test(ct)) return ct;
  return ct.replace(/<Types([^>]*)>/, `<Types$1><Default Extension="${ext}" ContentType="${type}"/>`);
}
export const ctRemove = (ct, part) => ct.replace(new RegExp(`<Override PartName="/${part.replace(/[.[\]]/g, '\\$&')}"[^>]*/>`), '');
