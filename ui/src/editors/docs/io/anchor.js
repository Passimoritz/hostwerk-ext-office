/*
 * Floating Word drawings (wp:anchor) – text boxes and shapes kept 1:1 as `preserved` nodes:
 *   anchorOf(xml)              → { x, y, w, h, relH, relV, alignH, alignV, wrap: 'square'|'none'|'topBottom', behind } | null
 *                                (px at 96 dpi; x/y from wp:posOffset, null when Word aligns instead)
 *   moveAnchor(xml, x, y)      → the run XML with the box at x/y (px): positionH relativeFrom="column", positionV
 *                                relativeFrom="paragraph", wp:posOffset in EMU (wp:align replaced) – the rest stays as it is
 *   floatStyle(f, { contentW, marginLeft }) → how the editor shows it: { mode: 'float-left'|'float-right'|'block'|'absolute', left, top, marginLeft, marginRight }
 * Pure (strings in, strings out) – used by the docx reader and the editor's node view, tested in tests/docs-anchor.test.js.
 */
const EMU = 9525;
const num = (s) => { const v = Number(s); return Number.isFinite(v) ? v : null; };

function part(xml, tag) {
  const m = new RegExp(`<wp:${tag}\\b([^>]*)>([\\s\\S]*?)<\\/wp:${tag}>`).exec(xml);
  if (!m) return null;
  const rel = (/relativeFrom="([^"]*)"/.exec(m[1]) || [])[1] || null;
  const off = /<wp:posOffset>\s*(-?\d+)\s*<\/wp:posOffset>/.exec(m[2]);
  const al = /<wp:align>\s*([\w-]+)\s*<\/wp:align>/.exec(m[2]);
  return { rel, off: off ? num(off[1]) : null, align: al ? al[1] : null };
}

export function anchorOf(xml) {
  const s = String(xml || '');
  const a = /<wp:anchor\b([^>]*)>([\s\S]*?)<\/wp:anchor>/.exec(s);
  if (!a) return null;
  const body = a[2];
  const h = part(body, 'positionH'); const v = part(body, 'positionV');
  const ext = /<wp:extent\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(body);
  const wrap = /<wp:wrapNone\b/.test(body) ? 'none' : /<wp:wrapTopAndBottom\b/.test(body) ? 'topBottom' : /<wp:wrap(Square|Tight|Through)\b/.test(body) ? 'square' : 'none';
  return {
    x: h && h.off != null ? Math.round(h.off / EMU) : null,
    y: v && v.off != null ? Math.round(v.off / EMU) : null,
    w: ext ? Math.round(Number(ext[1]) / EMU) : null,
    h: ext ? Math.round(Number(ext[2]) / EMU) : null,
    relH: (h && h.rel) || 'column', relV: (v && v.rel) || 'paragraph',
    alignH: h && h.align, alignV: v && v.align,
    wrap, behind: /\bbehindDoc="(1|true)"/.test(a[1]),
  };
}

export function moveAnchor(xml, x, y) {
  const s = String(xml || '');
  const set = (str, tag, rel, px) => str.replace(new RegExp(`<wp:${tag}\\b[^>]*>[\\s\\S]*?<\\/wp:${tag}>`), `<wp:${tag} relativeFrom="${rel}"><wp:posOffset>${Math.round(px * EMU)}</wp:posOffset></wp:${tag}>`);
  return s.replace(/<wp:anchor\b[^>]*>[\s\S]*?<\/wp:anchor>/, (m) => set(set(m, 'positionH', 'column', x), 'positionV', 'paragraph', y));
}

/**
 * where the editor shows a floating box. contentW: text width (px), marginLeft: the page's left margin (px, for
 * positions measured from the page edge). Vertical offsets count only from the anchor's paragraph/line – positions from
 * the page or margin top keep the box at its paragraph (the editor flows, it has no fixed page grid while typing).
 */
export function floatStyle(f, { contentW = 600, marginLeft = 94 } = {}) {
  if (!f) return null;
  const w = f.w || 160;
  let x;
  if (f.alignH) x = f.alignH === 'right' || f.alignH === 'outside' ? contentW - w : f.alignH === 'center' ? (contentW - w) / 2 : 0;
  else x = (f.x || 0) - (f.relH === 'page' ? marginLeft : 0);
  x = Math.max(-marginLeft, Math.min(contentW - 8, Math.round(x)));
  const top = ['paragraph', 'line'].includes(f.relV) && !f.alignV ? Math.round(f.y || 0) : 0;
  if (f.wrap === 'square') {
    const right = f.alignH === 'right' || f.alignH === 'outside' || (!f.alignH && x + w / 2 > contentW / 2);
    return right ? { mode: 'float-right', marginRight: Math.max(0, Math.round(contentW - x - w)), top: Math.max(0, top) } : { mode: 'float-left', marginLeft: Math.max(0, x), top: Math.max(0, top) };
  }
  if (f.wrap === 'topBottom') return { mode: 'block', marginLeft: Math.max(0, x), top: Math.max(0, top) };
  return { mode: 'absolute', left: x, top };
}
