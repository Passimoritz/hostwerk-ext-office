/*
 * Small MIME (RFC 5322 / 2045–2047 / 2231) parser for .eml files – pure, works in Node and the browser.
 *
 *   parseMime(bytes)  → part tree { headers: [[name, value]], type, params, disposition, dparams, body: Uint8Array, parts: [] }
 *   readMail(bytes)   → { subject, from, to, cc, bcc, replyTo, date, messageId, html, text, attachments: [{ name, mime, bytes, cid, inline, size }] }
 *   decodeWords(str)  → RFC 2047 encoded-words decoded
 *
 * Bytes are handled as a "binary string" (one char per byte) so offsets stay exact; text is decoded with
 * TextDecoder by charset at the very end.
 */

const latin1 = (u8) => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return s; };
const toBytes = (bin) => { const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i) & 0xff; return u; };

const ALIAS = { 'utf8': 'utf-8', 'latin1': 'iso-8859-1', 'latin-1': 'iso-8859-1', 'us-ascii': 'utf-8', 'ascii': 'utf-8', 'cp1252': 'windows-1252', 'x-unknown': 'utf-8', 'unknown-8bit': 'windows-1252', 'ks_c_5601-1987': 'euc-kr', 'gb2312': 'gbk' };
export function decodeCharset(u8, charset) {
  let cs = String(charset || 'utf-8').trim().toLowerCase().replace(/^"|"$/g, '');
  cs = ALIAS[cs] || cs;
  try { return new TextDecoder(cs, { fatal: cs === 'utf-8' }).decode(u8); } catch { /* wrong label or invalid utf-8 */ }
  try { return new TextDecoder('windows-1252').decode(u8); } catch { return latin1(u8); }
}

export function decodeBase64(s) {
  const clean = String(s).replace(/[^A-Za-z0-9+/]/g, '');
  const cut = clean.length % 4 === 1 ? clean.slice(0, -1) : clean;
  const padded = cut + '='.repeat((4 - (cut.length % 4)) % 4);
  const bin = typeof atob === 'function' ? atob(padded) : Buffer.from(padded, 'base64').toString('latin1');
  return toBytes(bin);
}

/** quoted-printable → bytes. header=true: "_" is a space (RFC 2047 Q encoding) */
export function decodeQP(s, header = false) {
  let t = String(s);
  if (header) t = t.replace(/_/g, ' ');
  else t = t.replace(/=\r?\n/g, '');
  const out = [];
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (c === '=' && /^[0-9A-Fa-f]{2}$/.test(t.substr(i + 1, 2))) { out.push(parseInt(t.substr(i + 1, 2), 16)); i += 2; } else out.push(t.charCodeAt(i) & 0xff);
  }
  return Uint8Array.from(out);
}

/** RFC 2047: =?charset?B|Q?text?= – whitespace between adjacent encoded words is dropped */
export function decodeWords(str) {
  if (!str || str.indexOf('=?') < 0) return str || '';
  const re = /=\?([^?\s]+)\?([BbQq])\?([^?]*)\?=/g;
  // join adjacent words first: same charset words may split multi-byte characters
  const tokens = []; let last = 0; let m;
  while ((m = re.exec(str))) {
    if (m.index > last) tokens.push({ text: str.slice(last, m.index) });
    tokens.push({ cs: m[1].split('*')[0], enc: m[2].toUpperCase(), data: m[3] });
    last = re.lastIndex;
  }
  if (last < str.length) tokens.push({ text: str.slice(last) });
  const out = [];
  for (let i = 0; i < tokens.length; i++) {
    const tk = tokens[i];
    if (tk.text != null) {
      if (/^\s+$/.test(tk.text) && tokens[i - 1] && tokens[i - 1].cs && tokens[i + 1] && tokens[i + 1].cs) continue;
      out.push(tk.text); continue;
    }
    // collect a run of words with the same charset
    const bytes = [];
    let j = i;
    for (; j < tokens.length; j++) {
      const w = tokens[j];
      if (w.text != null) { if (/^\s+$/.test(w.text) && tokens[j + 1] && tokens[j + 1].cs && tokens[j + 1].cs.toLowerCase() === tk.cs.toLowerCase()) continue; break; }
      if (w.cs.toLowerCase() !== tk.cs.toLowerCase()) break;
      bytes.push(...(w.enc === 'B' ? decodeBase64(w.data) : decodeQP(w.data, true)));
    }
    out.push(decodeCharset(Uint8Array.from(bytes), tk.cs));
    i = j - 1;
  }
  return out.join('');
}

/** header block → [[lowercased name, raw value]] (folded lines joined) */
export function parseHeaders(block) {
  const out = [];
  for (const line of block.split(/\r?\n/)) {
    if (/^[ \t]/.test(line) && out.length) out[out.length - 1][1] += ` ${line.trim()}`;
    else { const i = line.indexOf(':'); if (i > 0) out.push([line.slice(0, i).trim().toLowerCase(), line.slice(i + 1).trim()]); }
  }
  for (const h of out) h[1] = raw8(h[1]);
  return out;
}
/** raw 8-bit header text (RFC 6532 – mailers write UTF-8 straight into the headers) → text; invalid UTF-8 → Windows-1252 */
function raw8(v) { return /[\x80-\xff]/.test(v) ? decodeCharset(toBytes(v), 'utf-8') : v; }

/** "text/html; charset=utf-8; name*=…" → { value, params } (RFC 2231 continuations and charsets) */
export function parseParams(v) {
  const parts = []; let cur = ''; let q = false;
  for (let i = 0; i < (v || '').length; i++) {
    const c = v[i];
    if (c === '"' && v[i - 1] !== '\\') q = !q;
    if (c === ';' && !q) { parts.push(cur); cur = ''; } else cur += c;
  }
  parts.push(cur);
  const value = (parts.shift() || '').trim().toLowerCase();
  const raw = {};
  for (const p of parts) {
    const i = p.indexOf('='); if (i < 0) continue;
    const k = p.slice(0, i).trim().toLowerCase(); let val = p.slice(i + 1).trim();
    if (val.startsWith('"')) val = val.slice(1, val.endsWith('"') ? -1 : undefined).replace(/\\(.)/g, '$1');
    raw[k] = val;
  }
  const params = {}; const cont = {};
  for (const [k, val] of Object.entries(raw)) {
    const m = /^([^*]+)\*(\d+)?(\*)?$/.exec(k);
    if (!m) { params[k] = decodeWords(val); continue; }
    const [, name, n, star] = m;
    if (n === undefined) { params[name] = decode2231(val, true); continue; }
    (cont[name] = cont[name] || []).push({ n: Number(n), val, star: !!star });
  }
  for (const [name, list] of Object.entries(cont)) {
    list.sort((a, b) => a.n - b.n);
    const enc = list[0].star;
    if (enc) params[name] = decode2231(list.map((x) => x.val).join(''), true);
    else params[name] = list.map((x) => x.val).join('');
  }
  return { value, params };
}
function decode2231(val, extended) {
  if (!extended) return val;
  const m = /^([^']*)'[^']*'(.*)$/.exec(val);
  const cs = m ? m[1] || 'utf-8' : 'utf-8'; const data = m ? m[2] : val;
  const bytes = []; for (let i = 0; i < data.length; i++) { if (data[i] === '%' && /^[0-9A-Fa-f]{2}$/.test(data.substr(i + 1, 2))) { bytes.push(parseInt(data.substr(i + 1, 2), 16)); i += 2; } else bytes.push(data.charCodeAt(i) & 0xff); }
  return decodeCharset(Uint8Array.from(bytes), cs);
}

const header = (h, name) => { const e = h.find(([k]) => k === name); return e ? e[1] : ''; };

function parsePart(bin, depth = 0) {
  const sep = /\r?\n\r?\n/.exec(bin);
  const headBlock = sep ? bin.slice(0, sep.index) : bin;
  const bodyBin = sep ? bin.slice(sep.index + sep[0].length) : '';
  const headers = parseHeaders(headBlock);
  const ct = parseParams(header(headers, 'content-type') || 'text/plain');
  const cd = parseParams(header(headers, 'content-disposition'));
  const cte = (header(headers, 'content-transfer-encoding') || '7bit').trim().toLowerCase();
  const part = { headers, type: ct.value || 'text/plain', params: ct.params, disposition: cd.value, dparams: cd.params, encoding: cte, parts: [], body: null };
  if (part.type.startsWith('multipart/') && ct.params.boundary && depth < 30) {
    const b = `--${ct.params.boundary}`;
    const chunks = [];
    let pos = bin.indexOf(b, sep ? sep.index : 0);
    while (pos >= 0) {
      let start = pos + b.length;
      if (bin.startsWith('--', start)) break; // closing boundary
      const nl = bin.indexOf('\n', start); if (nl < 0) break;
      start = nl + 1;
      const next = bin.indexOf(`\n${b}`, start);
      const end = next < 0 ? bin.length : next;
      chunks.push(bin.slice(start, end).replace(/\r$/, ''));
      pos = next < 0 ? -1 : next + 1;
    }
    part.parts = chunks.map((c) => parsePart(c, depth + 1));
    return part;
  }
  part.body = cte === 'base64' ? decodeBase64(bodyBin) : cte === 'quoted-printable' ? decodeQP(bodyBin) : toBytes(bodyBin);
  return part;
}

export function parseMime(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return parsePart(latin1(u8));
}

/** "Lena <lena@x.de>, "Doe, John" <j@d.e>" → [{ name, email }] */
export function parseAddresses(v) {
  const s = decodeWords(v || '');
  const out = []; let cur = ''; let q = false; let angle = 0;
  for (const c of s) {
    if (c === '"') q = !q;
    if (!q && c === '<') angle++;
    if (!q && c === '>') angle--;
    if (c === ',' && !q && angle <= 0) { out.push(cur); cur = ''; } else cur += c;
  }
  out.push(cur);
  return out.map((x) => x.trim()).filter(Boolean).map((x) => {
    const m = /^(.*)<([^>]*)>\s*$/.exec(x);
    if (m) return { name: m[1].trim().replace(/^"|"$/g, '').replace(/\\"/g, '"'), email: m[2].trim() };
    return { name: '', email: x.replace(/^"|"$/g, '') };
  });
}

export const partText = (p) => decodeCharset(p.body || new Uint8Array(), p.params.charset);

/** walk the tree → body (html preferred, text as fallback) + attachments */
export function readMail(bytes) {
  const root = parseMime(bytes);
  const h = root.headers;
  const res = {
    subject: decodeWords(header(h, 'subject')), from: parseAddresses(header(h, 'from')), to: parseAddresses(header(h, 'to')),
    cc: parseAddresses(header(h, 'cc')), bcc: parseAddresses(header(h, 'bcc')), replyTo: parseAddresses(header(h, 'reply-to')),
    date: header(h, 'date') ? new Date(header(h, 'date')) : null, messageId: header(h, 'message-id'), html: '', text: '', attachments: [], headers: h,
  };
  if (res.date && Number.isNaN(res.date.getTime())) res.date = null;
  const walk = (p, inAlt) => {
    if (p.type.startsWith('multipart/')) {
      if (p.type === 'multipart/alternative') {
        // the last part we can show is the best one
        const html = p.parts.filter((x) => x.type === 'text/html' || x.type.startsWith('multipart/'));
        const text = p.parts.find((x) => x.type === 'text/plain');
        if (text && !res.text) res.text = partText(text);
        if (html.length) html.forEach((x) => walk(x, true)); else if (text && !res.html && !res.text) res.text = partText(text);
        return;
      }
      p.parts.forEach((x) => walk(x, inAlt));
      return;
    }
    const name = p.dparams.filename || p.params.name || '';
    const isAtt = p.disposition === 'attachment' || (name && !(p.type === 'text/html' || p.type === 'text/plain'));
    if (!isAtt && p.type === 'text/html' && !res.html) { res.html = partText(p); return; }
    if (!isAtt && p.type === 'text/plain' && !res.text && !inAlt) { res.text = partText(p); return; }
    if (!isAtt && p.type === 'text/plain' && res.text && !inAlt && p.disposition !== 'attachment') { res.text += `\n\n${partText(p)}`; return; }
    const cid = (header(p.headers, 'content-id') || '').replace(/^<|>$/g, '');
    const ext = p.type === 'message/rfc822' ? '.eml' : '';
    res.attachments.push({ name: name || (cid ? cid.replace(/@.*/, '') : `attachment${res.attachments.length + 1}${ext}`), mime: p.type, bytes: p.body, cid, inline: p.disposition === 'inline' || (!!cid && !name), size: p.body ? p.body.length : 0 });
  };
  walk(root, false);
  return res;
}

/** plain text out of HTML (for toText and the text view) – no DOM needed */
export function htmlToText(html) {
  return String(html || '')
    .replace(/<(script|style|head)[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|tr|li|h\d|blockquote)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n))).replace(/&amp;/g, '&')
    .replace(/\n{3,}/g, '\n\n').trim();
}
