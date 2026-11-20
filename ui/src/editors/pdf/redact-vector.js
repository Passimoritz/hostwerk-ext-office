/*
 * Redaction that keeps the page as vectors (no picture): the page's content stream is rewritten so that
 *  - every glyph whose box centre lies inside a mark is taken out of its text operator (the following glyphs keep
 *    their place through a TJ offset),
 *  - every path whose bounds lie inside a mark is not painted (its painting operator becomes `n`),
 *  - the rest is drawn through a clip that leaves the marks out, and the marks are painted black on top.
 * Only for pages we can follow completely: no images (XObject or inline), no form XObjects, no Type3 fonts, fonts with
 * known widths (simple fonts with /Widths, the standard 14 with WinAnsi, Type0 with Identity-H/V). Everything else →
 * false and the caller uses the picture method (redact.js). The caller verifies the result in any case.
 */
import { PDFName, PDFDict, PDFArray, PDFRef, PDFNumber, PDFRawStream, PDFStream, decodePDFRawStream } from 'pdf-lib';
import { Font, FontNames, Encodings } from '@pdf-lib/standard-fonts';

const look = (ctx, o) => (o instanceof PDFRef ? ctx.lookup(o) : o);
const num = (ctx, o, d = 0) => { const v = look(ctx, o); return v instanceof PDFNumber ? v.asNumber() : d; };

/* ---------------- lexer ---------------- */
const WS = new Set([0, 9, 10, 12, 13, 32]);
const DELIM = new Set('()<>[]{}/%'.split('').map((c) => c.charCodeAt(0)));
export function lex(b) {
  const out = []; let i = 0; const n = b.length;
  const str = (a, z) => { let s = ''; for (let k = a; k < z; k++) s += String.fromCharCode(b[k]); return s; };
  while (i < n) {
    const c = b[i];
    if (WS.has(c)) { i++; continue; }
    if (c === 37) { while (i < n && b[i] !== 10 && b[i] !== 13) i++; continue; } // %
    if (c === 40) { // ( literal string )
      const bytes = []; let depth = 1; i++;
      while (i < n && depth) {
        let ch = b[i++];
        if (ch === 92) { // backslash
          const e = b[i++];
          const map = { 110: 10, 114: 13, 116: 9, 98: 8, 102: 12, 40: 40, 41: 41, 92: 92 };
          if (map[e] != null) bytes.push(map[e]);
          else if (e >= 48 && e <= 55) { let v = e - 48; for (let k = 0; k < 2 && b[i] >= 48 && b[i] <= 55; k++) v = v * 8 + (b[i++] - 48); bytes.push(v & 255); } else if (e === 13) { if (b[i] === 10) i++; } else if (e !== 10) bytes.push(e);
          continue;
        }
        if (ch === 40) depth++; else if (ch === 41) { depth--; if (!depth) break; }
        bytes.push(ch);
      }
      out.push({ t: 'str', v: Uint8Array.from(bytes) });
      continue;
    }
    if (c === 60 && b[i + 1] === 60) { out.push({ t: 'op', v: '<<' }); i += 2; continue; }
    if (c === 62 && b[i + 1] === 62) { out.push({ t: 'op', v: '>>' }); i += 2; continue; }
    if (c === 60) { // <hex>
      let j = i + 1; let h = '';
      while (j < n && b[j] !== 62) { if (!WS.has(b[j])) h += String.fromCharCode(b[j]); j++; }
      if (h.length % 2) h += '0';
      const bytes = new Uint8Array(h.length / 2); for (let k = 0; k < bytes.length; k++) bytes[k] = parseInt(h.slice(k * 2, k * 2 + 2), 16);
      out.push({ t: 'str', v: bytes }); i = j + 1; continue;
    }
    if (c === 91) { out.push({ t: '[' }); i++; continue; }
    if (c === 93) { out.push({ t: ']' }); i++; continue; }
    if (c === 47) { let j = i + 1; while (j < n && !WS.has(b[j]) && !DELIM.has(b[j])) j++; out.push({ t: 'name', v: str(i, j) }); i = j; continue; }
    let j = i; while (j < n && !WS.has(b[j]) && !DELIM.has(b[j])) j++;
    if (j === i) { i++; continue; }
    const w = str(i, j);
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(w)) out.push({ t: 'num', v: parseFloat(w), raw: w });
    else if (w === 'BI') return null; // inline image
    else out.push({ t: 'op', v: w });
    i = j;
  }
  return out;
}
const tokText = (t) => {
  if (t.t === 'num') return t.raw != null ? t.raw : fmt(t.v);
  if (t.t === 'name') return t.v;
  if (t.t === 'str') return `<${Array.from(t.v, (x) => x.toString(16).padStart(2, '0')).join('')}>`;
  if (t.t === '[' || t.t === ']') return t.t;
  return t.v;
};
const fmt = (v) => (Math.abs(v) < 1e-9 ? '0' : String(Math.round(v * 10000) / 10000));

/* ---------------- fonts ---------------- */
let winAnsiNames = null;
function winAnsi() {
  if (winAnsiNames) return winAnsiNames;
  winAnsiNames = new Map();
  const m = Encodings.WinAnsi.unicodeMappings;
  for (const k of Object.keys(m)) { const [code, name] = m[k]; winAnsiNames.set(code, name); }
  return winAnsiNames;
}
const STD = { Helvetica: FontNames.Helvetica, 'Helvetica-Bold': FontNames.HelveticaBold, 'Helvetica-Oblique': FontNames.HelveticaOblique, 'Helvetica-BoldOblique': FontNames.HelveticaBoldOblique, 'Times-Roman': FontNames.TimesRoman, 'Times-Bold': FontNames.TimesRomanBold, 'Times-Italic': FontNames.TimesRomanItalic, 'Times-BoldItalic': FontNames.TimesRomanBoldItalic, Courier: FontNames.Courier, 'Courier-Bold': FontNames.CourierBold, 'Courier-Oblique': FontNames.CourierOblique, 'Courier-BoldOblique': FontNames.CourierBoldOblique, Arial: FontNames.Helvetica, 'Arial,Bold': FontNames.HelveticaBold, TimesNewRoman: FontNames.TimesRoman };

/** font dict → { bytes: 1|2, width(code) → glyph space units (1/1000), vertical } or null when unsupported */
function fontInfo(ctx, fd) {
  if (!(fd instanceof PDFDict)) return null;
  const sub = String(look(ctx, fd.get(PDFName.of('Subtype'))) || '');
  if (sub === '/Type3') return null;
  if (sub === '/Type0') {
    const enc = String(look(ctx, fd.get(PDFName.of('Encoding'))) || '');
    if (enc !== '/Identity-H' && enc !== '/Identity-V') return null;
    const desc = look(ctx, fd.get(PDFName.of('DescendantFonts')));
    const cid = desc instanceof PDFArray ? look(ctx, desc.get(0)) : null;
    if (!(cid instanceof PDFDict)) return null;
    const dw = num(ctx, cid.get(PDFName.of('DW')), 1000);
    const W = look(ctx, cid.get(PDFName.of('W')));
    const widths = new Map();
    if (W instanceof PDFArray) {
      const a = W.asArray().map((x) => look(ctx, x));
      for (let i = 0; i < a.length;) {
        const first = a[i] instanceof PDFNumber ? a[i].asNumber() : 0;
        if (a[i + 1] instanceof PDFArray) { a[i + 1].asArray().forEach((w, k) => widths.set(first + k, num(ctx, w, dw))); i += 2; } else { const last = num(ctx, a[i + 1]); const w = num(ctx, a[i + 2], dw); for (let c = first; c <= last && c - first < 70000; c++) widths.set(c, w); i += 3; }
      }
    }
    return { bytes: 2, width: (c) => (widths.has(c) ? widths.get(c) : dw), vertical: enc === '/Identity-V' };
  }
  const W = look(ctx, fd.get(PDFName.of('Widths')));
  if (W instanceof PDFArray) {
    const fc = num(ctx, fd.get(PDFName.of('FirstChar')), 0);
    const desc = look(ctx, fd.get(PDFName.of('FontDescriptor')));
    const missing = desc instanceof PDFDict ? num(ctx, desc.get(PDFName.of('MissingWidth')), 0) : 0;
    const ws = W.asArray().map((x) => num(ctx, x, missing));
    return { bytes: 1, width: (c) => (c >= fc && c - fc < ws.length ? ws[c - fc] : missing) };
  }
  const base = String(look(ctx, fd.get(PDFName.of('BaseFont'))) || '').replace(/^\//, '').replace(/^[A-Z]{6}\+/, '');
  const std = STD[base];
  if (!std) return null;
  const enc = look(ctx, fd.get(PDFName.of('Encoding')));
  if (enc && String(enc) !== '/WinAnsiEncoding' && !/Symbol|Zapf/.test(base)) return null;
  const f = Font.load(std); const names = winAnsi();
  return { bytes: 1, width: (c) => { const g = names.get(c); try { return g ? f.getWidthOfGlyph(g) : 0; } catch { return 0; } } };
}

/* ---------------- geometry ---------------- */
const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const inside = (x, y, r) => x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h;

/**
 * Rewrites the page in place. Returns true when done, false when the page needs the picture method.
 */
export function redactPageVector(doc, page, rects) {
  const ctx = doc.context;
  const node = page.node;
  const res = look(ctx, node.Resources ? node.Resources() : node.get(PDFName.of('Resources')));
  const xo = res instanceof PDFDict ? look(ctx, res.get(PDFName.of('XObject'))) : null;
  if (xo instanceof PDFDict && xo.keys().length) return false; // pictures or forms on the page
  const fontsDict = res instanceof PDFDict ? look(ctx, res.get(PDFName.of('Font'))) : null;
  const fontCache = new Map();
  const fontOf = (name) => {
    if (fontCache.has(name)) return fontCache.get(name);
    const fd = fontsDict instanceof PDFDict ? look(ctx, fontsDict.get(PDFName.of(name.replace(/^\//, '')))) : null;
    const fi = fontInfo(ctx, fd); fontCache.set(name, fi); return fi;
  };
  // content
  const cont = look(ctx, node.get(PDFName.of('Contents')));
  const parts = cont instanceof PDFArray ? cont.asArray().map((x) => look(ctx, x)) : cont ? [cont] : [];
  const chunks = [];
  for (const s of parts) {
    if (!(s instanceof PDFStream)) return false;
    let data;
    try { data = s instanceof PDFRawStream ? decodePDFRawStream(s).decode() : s.getContents(); } catch { return false; }
    chunks.push(data, new Uint8Array([10]));
  }
  const all = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0)); let p = 0; for (const c of chunks) { all.set(c, p); p += c.length; }
  const toks = lex(all);
  if (!toks) return false;
  // walk
  const out = [];
  let ops = [];
  let gs = { ctm: [1, 0, 0, 1, 0, 0] }; const stack = [];
  let ts = { tm: [1, 0, 0, 1, 0, 0], tlm: [1, 0, 0, 1, 0, 0], font: null, size: 0, tc: 0, tw: 0, th: 1, tl: 0, rise: 0 };
  let path = []; // device-space points of the current path
  let pathStart = 0;
  let unsupported = false;
  // marked-content properties may repeat the text (/ActualText, /Alt, /E) – they go
  const emit = (arr) => { const keep = []; for (let k = 0; k < arr.length; k++) { const x = arr[k]; if (x.t === 'name' && /^\/(ActualText|Alt|E)$/.test(x.v) && arr[k + 1] && arr[k + 1].t === 'str') { k++; continue; } keep.push(x); } out.push(keep.map(tokText).join(' ')); };
  const showRun = (items) => {
    // items: [{ str } | { num }] → kept TJ array
    const fi = ts.font && fontOf(ts.font);
    if (!fi) { unsupported = true; return null; }
    const res2 = [];
    const trmBase = () => mul([ts.size * ts.th, 0, 0, ts.size, 0, ts.rise], mul(ts.tm, gs.ctm));
    for (const it of items) {
      if (it.num != null) { const tx = (-it.num / 1000) * ts.size * ts.th; if (fi.vertical) ts.tm = mul([1, 0, 0, 1, 0, (-it.num / 1000) * ts.size], ts.tm); else ts.tm = mul([1, 0, 0, 1, tx, 0], ts.tm); res2.push({ num: it.num }); continue; }
      const b = it.str; let kept = []; let skip = 0;
      const flush = () => { if (kept.length) { res2.push({ str: Uint8Array.from(kept) }); kept = []; } if (skip) { res2.push({ num: -skip }); skip = 0; } };
      for (let i = 0; i + fi.bytes <= b.length; i += fi.bytes) {
        const code = fi.bytes === 2 ? (b[i] << 8) | b[i + 1] : b[i];
        const w0 = fi.width(code);
        const trm = trmBase();
        // glyph box centre in text space units (w0/2, ~0.35 em) → page
        const [cx, cy] = fi.vertical ? apply(trm, 0, -w0 / 2000) : apply(trm, w0 / 2000, 0.35);
        // hit: the glyph's centre is in a mark, or a mark covers at least 30 % of its box (marks from text search are
        // estimated from average character widths and may end a little early)
        let hit = rects.some((r) => inside(cx, cy, r));
        if (!hit && w0 > 0) {
          const q = (fi.vertical ? [[-w0 / 2000, 0], [w0 / 2000, 0], [-w0 / 2000, -w0 / 1000], [w0 / 2000, -w0 / 1000]] : [[0, -0.2], [w0 / 1000, -0.2], [0, 0.8], [w0 / 1000, 0.8]]).map(([x, y]) => apply(trm, x, y));
          const gx = Math.min(...q.map((v) => v[0])); const gX = Math.max(...q.map((v) => v[0])); const gy = Math.min(...q.map((v) => v[1])); const gY = Math.max(...q.map((v) => v[1]));
          const area = Math.max(1e-6, (gX - gx) * (gY - gy));
          hit = rects.some((r) => { const ox = Math.min(gX, r.x + r.w) - Math.max(gx, r.x); const oy = Math.min(gY, r.y + r.h) - Math.max(gy, r.y); return ox > 0 && oy > 0 && (ox * oy) / area >= 0.3; });
        }
        const wsp = fi.bytes === 1 && code === 32 ? ts.tw : 0;
        const adv = fi.vertical ? (w0 / 1000) * ts.size + ts.tc + wsp : ((w0 / 1000) * ts.size + ts.tc + wsp) * ts.th;
        if (hit) { if (kept.length) { res2.push({ str: Uint8Array.from(kept) }); kept = []; } skip += ts.size ? (adv / (fi.vertical ? 1 : ts.th) / ts.size) * 1000 : 0; } else { if (skip) { res2.push({ num: -skip }); skip = 0; } for (let k = 0; k < fi.bytes; k++) kept.push(b[i + k]); }
        ts.tm = fi.vertical ? mul([1, 0, 0, 1, 0, -adv], ts.tm) : mul([1, 0, 0, 1, adv, 0], ts.tm);
      }
      flush();
    }
    return res2;
  };
  const tjOut = (arr) => {
    const items = arr.filter((x) => x.str ? x.str.length : x.num);
    if (!items.length) return '';
    return `[${items.map((x) => (x.str ? tokText({ t: 'str', v: x.str }) : fmt(x.num))).join(' ')}] TJ`;
  };
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t !== 'op') { ops.push(t); continue; }
    const op = t.v; const a = ops; ops = [];
    const N = (k) => (a[k] && a[k].t === 'num' ? a[k].v : 0);
    switch (op) {
      case 'q': stack.push({ ...gs }); emit([...a, t]); break;
      case 'Q': gs = stack.pop() || gs; emit([...a, t]); break;
      case 'cm': gs = { ...gs, ctm: mul([N(0), N(1), N(2), N(3), N(4), N(5)], gs.ctm) }; emit([...a, t]); break;
      case 'BT': ts.tm = [1, 0, 0, 1, 0, 0]; ts.tlm = [1, 0, 0, 1, 0, 0]; emit([...a, t]); break;
      case 'Tf': ts.font = a[0] && a[0].v; ts.size = N(1); emit([...a, t]); break;
      case 'Tc': ts.tc = N(0); emit([...a, t]); break;
      case 'Tw': ts.tw = N(0); emit([...a, t]); break;
      case 'Tz': ts.th = N(0) / 100; emit([...a, t]); break;
      case 'TL': ts.tl = N(0); emit([...a, t]); break;
      case 'Ts': ts.rise = N(0); emit([...a, t]); break;
      case 'Td': ts.tlm = mul([1, 0, 0, 1, N(0), N(1)], ts.tlm); ts.tm = ts.tlm.slice(); emit([...a, t]); break;
      case 'TD': ts.tl = -N(1); ts.tlm = mul([1, 0, 0, 1, N(0), N(1)], ts.tlm); ts.tm = ts.tlm.slice(); emit([...a, t]); break;
      case 'Tm': ts.tlm = [N(0), N(1), N(2), N(3), N(4), N(5)]; ts.tm = ts.tlm.slice(); emit([...a, t]); break;
      case 'T*': ts.tlm = mul([1, 0, 0, 1, 0, -ts.tl], ts.tlm); ts.tm = ts.tlm.slice(); emit([...a, t]); break;
      case 'Tj': { const r = showRun([{ str: a[0] && a[0].v || new Uint8Array() }]); if (!r) break; out.push(tjOut(r)); break; }
      case "'": { ts.tlm = mul([1, 0, 0, 1, 0, -ts.tl], ts.tlm); ts.tm = ts.tlm.slice(); out.push('T*'); const r = showRun([{ str: a[0] && a[0].v || new Uint8Array() }]); if (!r) break; out.push(tjOut(r)); break; }
      case '"': { ts.tw = N(0); ts.tc = N(1); out.push(`${fmt(ts.tw)} Tw ${fmt(ts.tc)} Tc`); ts.tlm = mul([1, 0, 0, 1, 0, -ts.tl], ts.tlm); ts.tm = ts.tlm.slice(); out.push('T*'); const r = showRun([{ str: a[2] && a[2].v || new Uint8Array() }]); if (!r) break; out.push(tjOut(r)); break; }
      case 'TJ': {
        const items = []; for (const x of a) { if (x.t === 'str') items.push({ str: x.v }); else if (x.t === 'num') items.push({ num: x.v }); }
        const r = showRun(items); if (!r) break; out.push(tjOut(r)); break;
      }
      // paths: remember the device-space points; at the painting operator decide
      case 'm': case 'l': if (op === 'm' && !path.length) pathStart = out.length; path.push(apply(gs.ctm, N(0), N(1))); emit([...a, t]); break;
      case 'c': path.push(apply(gs.ctm, N(0), N(1)), apply(gs.ctm, N(2), N(3)), apply(gs.ctm, N(4), N(5))); emit([...a, t]); break;
      case 'v': case 'y': path.push(apply(gs.ctm, N(0), N(1)), apply(gs.ctm, N(2), N(3))); emit([...a, t]); break;
      case 're': { if (!path.length) pathStart = out.length; const x = N(0); const y = N(1); const w = N(2); const h = N(3); path.push(apply(gs.ctm, x, y), apply(gs.ctm, x + w, y), apply(gs.ctm, x, y + h), apply(gs.ctm, x + w, y + h)); emit([...a, t]); break; }
      case 'h': emit([...a, t]); break;
      case 'S': case 's': case 'f': case 'F': case 'f*': case 'B': case 'B*': case 'b': case 'b*': case 'n': {
        const inMark = path.length && rects.some((r) => path.every(([x, y]) => inside(x, y, r)));
        if (inMark && op !== 'n') { out.length = pathStart; out.push('n'); } else emit([...a, t]);
        path = []; break;
      }
      case 'W': case 'W*': emit([...a, t]); break;
      case 'Do': case 'BI': case 'ID': case 'EI': case 'd0': case 'd1': unsupported = true; break;
      default: emit([...a, t]);
    }
    if (unsupported) return false;
  }
  if (unsupported) return false;
  // clip the marks out of everything, then paint them black
  const box = page.getMediaBox();
  const clip = `${fmt(box.x)} ${fmt(box.y)} ${fmt(box.width)} ${fmt(box.height)} re ${rects.map((r) => `${fmt(r.x)} ${fmt(r.y)} ${fmt(r.w)} ${fmt(r.h)} re`).join(' ')} W* n`;
  const black = rects.map((r) => `${fmt(r.x)} ${fmt(r.y)} ${fmt(r.w)} ${fmt(r.h)} re`).join(' ');
  const text = `q ${clip}\n${out.filter(Boolean).join('\n')}\nQ\nq 0 g ${black} f Q\n`;
  const stream = ctx.flateStream(new TextEncoder().encode(text));
  node.set(PDFName.of('Contents'), ctx.register(stream));
  return true;
}
