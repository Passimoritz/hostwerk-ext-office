/*
 * Outlook .msg reader on top of the shared CFB reader – pure. Returns the same shape as mime.js readMail().
 * (@kenjiuno/msgreader would add ~580 KB because of iconv-lite; MAPI properties are simple enough to read here.)
 *
 * Streams are named __substg1.0_<PROPID><TYPE>: 001F = UTF-16LE string, 001E = 8-bit string (code page), 0102 = binary.
 * Fixed-size values (times, recipient type) live in __properties_version1.0.
 */
import { readCfb } from '../_shared/cfb.js';
import { decodeCharset, readMail, htmlToText } from './mime.js';

const CP = { 1252: 'windows-1252', 1250: 'windows-1250', 1251: 'windows-1251', 1253: 'windows-1253', 1254: 'windows-1254', 1257: 'windows-1257', 28591: 'iso-8859-1', 28592: 'iso-8859-2', 28605: 'iso-8859-15', 65001: 'utf-8', 20127: 'utf-8', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 1200: 'utf-16le', 50220: 'iso-2022-jp', 51932: 'euc-jp' };

function props(streams, prefix, headerSize) {
  const out = new Map();
  const p = streams.get(`${prefix}__properties_version1.0`);
  if (!p) return out;
  const dv = new DataView(p.buffer, p.byteOffset, p.byteLength);
  for (let o = headerSize; o + 16 <= p.length; o += 16) {
    const type = dv.getUint16(o, true); const id = dv.getUint16(o + 2, true);
    if (type === 0x0003) out.set(id, dv.getInt32(o + 8, true));
    else if (type === 0x0040) { const lo = dv.getUint32(o + 8, true); const hi = dv.getUint32(o + 12, true); const ms = (hi * 4294967296 + lo) / 10000 - 11644473600000; out.set(id, new Date(ms)); } else if (type === 0x000b) out.set(id, dv.getUint16(o + 8, true) !== 0);
  }
  return out;
}

function reader(streams, prefix, codepage) {
  const get = (id, kinds = ['001F', '001E', '0102']) => {
    const hex = id.toString(16).toUpperCase().padStart(4, '0');
    for (const k of kinds) { const s = streams.get(`${prefix}__substg1.0_${hex}${k}`); if (s) return { k, s }; }
    return null;
  };
  const str = (id) => {
    const r = get(id, ['001F', '001E']);
    if (!r) return '';
    const t = r.k === '001F' ? new TextDecoder('utf-16le').decode(r.s) : decodeCharset(r.s, codepage);
    return t.replace(/\u0000+$/, '');
  };
  const bin = (id) => { const r = get(id, ['0102']); return r ? r.s : null; };
  return { get, str, bin };
}

export function isMsg(bytes) { return bytes && bytes[0] === 0xd0 && bytes[1] === 0xcf; }

export function readMsg(bytes, depth = 0) {
  const cfb = readCfb(bytes);
  const { streams } = cfb;
  const top = props(streams, '', depth ? 24 : 32);
  const cpId = top.get(0x3ffd) || top.get(0x3fde) || 1252;
  const codepage = CP[cpId] || 'windows-1252';
  const r = reader(streams, '', codepage);
  const headers = r.str(0x007d);
  // with the transport headers we get exact From/To/Cc/Date with encodings – reuse the MIME parser for them
  const hdr = headers ? readMail(new TextEncoder().encode(`${headers.trim()}\r\n\r\n`)) : null;
  const res = {
    subject: r.str(0x0037) || (hdr && hdr.subject) || '',
    from: [{ name: r.str(0x0c1a) || r.str(0x0042), email: r.str(0x5d01) || r.str(0x0c1f) || r.str(0x0065) }].filter((x) => x.name || x.email),
    to: [], cc: [], bcc: [], replyTo: hdr ? hdr.replyTo : [],
    date: top.get(0x0e06) || top.get(0x0039) || (hdr && hdr.date) || null,
    html: '', text: r.str(0x1000), attachments: [], headers: hdr ? hdr.headers : [], rtfOnly: false,
  };
  if (hdr && hdr.from.length && !res.from.some((x) => /@/.test(x.email))) res.from = hdr.from;
  // HTML body: string or binary in the internet code page
  const html = r.get(0x1013, ['001F', '001E', '0102']);
  if (html) res.html = html.k === '001F' ? new TextDecoder('utf-16le').decode(html.s) : decodeCharset(html.s, CP[top.get(0x3fde)] || codepage);
  if (!res.html && !res.text && r.bin(0x1009)) res.rtfOnly = true;

  // recipients
  const recips = new Set(); const atts = new Set();
  for (const e of cfb.entries) {
    const m = /^(__recip_version1\.0_#[0-9A-Fa-f]{8}|__attach_version1\.0_#[0-9A-Fa-f]{8})$/.exec(e.path);
    if (m && e.type === 'storage') (e.path.startsWith('__recip') ? recips : atts).add(e.path);
  }
  for (const p of [...recips].sort()) {
    const pr = props(streams, `${p}/`, 8); const rr = reader(streams, `${p}/`, codepage);
    const who = { name: rr.str(0x3001), email: rr.str(0x39fe) || rr.str(0x3003) };
    const type = pr.get(0x0c15) || 1;
    (type === 2 ? res.cc : type === 3 ? res.bcc : res.to).push(who);
  }
  if (!res.to.length && hdr) { res.to = hdr.to; res.cc = hdr.cc; }
  // attachments
  for (const p of [...atts].sort()) {
    const pr = props(streams, `${p}/`, 8); const ar = reader(streams, `${p}/`, codepage);
    const name = ar.str(0x3707) || ar.str(0x3704) || ar.str(0x3001) || 'attachment';
    const data = ar.bin(0x3701);
    const embedded = cfb.entries.some((e) => e.path === `${p}/__substg1.0_3701000D`);
    const cid = ar.str(0x3712);
    const hidden = pr.get(0x7ffe) === true;
    if (!data && !embedded) continue;
    res.attachments.push({ name: embedded && !/\.msg$/i.test(name) ? `${name}.msg` : name, mime: ar.str(0x370e) || (embedded ? 'application/vnd.ms-outlook' : 'application/octet-stream'), bytes: data || null, cid, inline: !!cid || hidden, size: data ? data.length : 0, embedded });
  }
  if (!res.text && res.html) res.text = htmlToText(res.html);
  return res;
}
