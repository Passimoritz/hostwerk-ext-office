/*
 * Flat ODF (fodp): one XML file instead of the zip. toFlat(odpZipBytes) → Uint8Array – takes the package our ODP writer
 * made and merges meta/settings/styles/content into <office:document>; pictures become <office:binary-data>.
 * Automatic styles of styles.xml are renamed (prefix "Ms") so they can share one office:automatic-styles with content.xml.
 */
import JSZip from 'jszip';

const sect = (xml, name) => {
  const open = new RegExp(`<${name}(\\s[^>]*)?(/>|>)`).exec(xml);
  if (!open) return null;
  if (open[2] === '/>') return { attrs: open[1] || '', inner: '' };
  const end = xml.indexOf(`</${name}>`, open.index);
  return end < 0 ? null : { attrs: open[1] || '', inner: xml.slice(open.index + open[0].length, end) };
};
const rootAttrs = (xml) => { const m = /<office:document-[a-z]+(\s[^>]*)>/.exec(xml); return m ? m[1] : ''; };
const nsOf = (attrs, out) => { const re = /\s(xmlns:[\w-]+)="([^"]*)"/g; let m; while ((m = re.exec(attrs))) if (!out[m[1]]) out[m[1]] = m[2]; };
function b64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return typeof btoa === 'function' ? btoa(s) : Buffer.from(bytes).toString('base64');
}

export async function toFlat(zipBytes) {
  const z = await JSZip.loadAsync(zipBytes);
  const read = async (p) => (z.file(p) ? z.file(p).async('string') : '');
  const [content, styles, meta, settings] = await Promise.all(['content.xml', 'styles.xml', 'meta.xml', 'settings.xml'].map(read));
  const ns = {};
  for (const x of [content, styles, meta, settings]) nsOf(rootAttrs(x), ns);
  for (const [k, v] of Object.entries({ 'xmlns:office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', 'xmlns:xlink': 'http://www.w3.org/1999/xlink', 'xmlns:style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0' })) if (!ns[k]) ns[k] = v;
  // rename automatic styles of styles.xml
  const sAuto = sect(styles, 'office:automatic-styles');
  const map = {};
  let sAutoInner = sAuto ? sAuto.inner : '';
  sAutoInner = sAutoInner.replace(/(<style:(?:style|page-layout)\b[^>]*\sstyle:name=")([^"]+)"/g, (m, a, n) => { map[n] = `Ms${n}`; return `${a}Ms${n}"`; });
  sAutoInner = sAutoInner.replace(/(<text:list-style\b[^>]*\sstyle:name=")([^"]+)"/g, (m, a, n) => { map[n] = `Ms${n}`; return `${a}Ms${n}"`; });
  const fixRefs = (s) => s.replace(/([\w-]+:(?:[\w-]*style-name|page-layout-name))="([^"]+)"/g, (m, a, n) => (map[n] ? `${a}="${map[n]}"` : m));
  const master = sect(styles, 'office:master-styles');
  const masterInner = fixRefs(master ? master.inner : '');
  sAutoInner = fixRefs(sAutoInner);
  // font faces (unique by name)
  const faces = {};
  for (const x of [styles, content]) { const f = sect(x, 'office:font-face-decls'); if (f) { const re = /<style:font-face\b[^>]*?style:name="([^"]+)"[^>]*?(?:\/>|>[\s\S]*?<\/style:font-face>)/g; let m; while ((m = re.exec(f.inner))) if (!faces[m[1]]) faces[m[1]] = m[0]; } }
  const cAuto = sect(content, 'office:automatic-styles');
  const body = sect(content, 'office:body');
  const common = sect(styles, 'office:styles');
  const metaS = sect(meta, 'office:meta');
  const settingsS = sect(settings, 'office:settings');
  // pictures → binary data
  const files = {};
  for (const p of Object.keys(z.files)) if (/^(Pictures|media)\//.test(p) && !z.files[p].dir) files[p] = await z.file(p).async('uint8array');
  const inline = (s) => s.replace(/<(draw:image|draw:fill-image|style:background-image)\b([^>]*?)\sxlink:href="([^"]+)"([^>]*?)(\/>|>)/g, (m, tg, a1, href, a2, close) => {
    const bytes = files[href];
    if (!bytes) return m;
    const rest = `${a1}${a2}`.replace(/\sxlink:(type|show|actuate)="[^"]*"/g, '');
    const bin = `<office:binary-data>${b64(bytes)}</office:binary-data>`;
    return close === '/>' ? `<${tg}${rest}>${bin}</${tg}>` : `<${tg}${rest}>${bin}`;
  });
  const out = '<?xml version="1.0" encoding="UTF-8"?>\n'
    + `<office:document ${Object.entries(ns).map(([k, v]) => `${k}="${v}"`).join(' ')} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation">`
    + (metaS ? `<office:meta>${metaS.inner}</office:meta>` : '')
    + (settingsS ? `<office:settings>${settingsS.inner}</office:settings>` : '')
    + `<office:font-face-decls>${Object.values(faces).join('')}</office:font-face-decls>`
    + `<office:styles>${inline(common ? common.inner : '')}</office:styles>`
    + `<office:automatic-styles>${inline(sAutoInner)}${inline(cAuto ? cAuto.inner : '')}</office:automatic-styles>`
    + `<office:master-styles>${inline(masterInner)}</office:master-styles>`
    + `<office:body>${inline(body ? body.inner : '')}</office:body>`
    + '</office:document>';
  return new TextEncoder().encode(out);
}
