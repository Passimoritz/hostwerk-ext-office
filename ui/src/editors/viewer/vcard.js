/*
 * vCard 2.1 / 3.0 / 4.0 parser (pure). parseVCards(text) → [{ version, fn, n: { family, given, middle, prefix, suffix },
 * org, title, role, nickname, bday, note, url: [], tel: [{ value, types }], email: [...], adr: [{ types, street, city,
 * region, code, country, pobox, ext, label }], photo: { mime, data (base64) | url }, categories: [], raw }].
 */
import { decodeQP, decodeCharset } from './mime.js';

/** unfold (RFC 6350 folding + vCard 2.1 quoted-printable soft breaks) → logical lines */
export function unfold(text) {
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
  const out = [];
  for (const l of lines) {
    const prev = out.length ? out[out.length - 1] : null;
    if (prev != null && /^[ \t]/.test(l)) { out[out.length - 1] = prev + l.slice(1); continue; }
    if (prev != null && /ENCODING=QUOTED-PRINTABLE/i.test(prev.split(':')[0]) && /=$/.test(prev)) { out[out.length - 1] = prev.slice(0, -1) + l; continue; }
    out.push(l);
  }
  return out.filter((l) => l.trim() !== '');
}

/** "item1.TEL;TYPE=work,voice;PREF=1:+49 …" → { group, name, params: { type: ['work','voice'], … }, value } */
export function parseLine(line) {
  let q = false; let i = 0;
  for (; i < line.length; i++) { const c = line[i]; if (c === '"') q = !q; else if (c === ':' && !q) break; }
  const head = line.slice(0, i); const value = line.slice(i + 1);
  const segs = []; let cur = ''; q = false;
  for (const c of head) { if (c === '"') q = !q; if (c === ';' && !q) { segs.push(cur); cur = ''; } else cur += c; }
  segs.push(cur);
  let name = segs.shift() || ''; let group = '';
  const dot = name.indexOf('.'); if (dot > 0) { group = name.slice(0, dot); name = name.slice(dot + 1); }
  const params = {};
  for (const s of segs) {
    const e = s.indexOf('=');
    if (e < 0) { (params.type = params.type || []).push(s.toLowerCase()); continue; } // vCard 2.1 bare types
    const k = s.slice(0, e).toLowerCase(); const v = s.slice(e + 1).replace(/^"|"$/g, '');
    if (k === 'type') (params.type = params.type || []).push(...v.toLowerCase().split(','));
    else params[k] = v;
  }
  return { group, name: name.toUpperCase(), params, value };
}

const unescape = (s) => String(s).replace(/\\n/gi, '\n').replace(/\\([,;:\\])/g, '$1');
/** split by unescaped `sep` */
export function splitEsc(s, sep = ';') {
  const out = []; let cur = '';
  for (let i = 0; i < s.length; i++) { if (s[i] === '\\' && i + 1 < s.length) { cur += s[i] + s[i + 1]; i++; } else if (s[i] === sep) { out.push(cur); cur = ''; } else cur += s[i]; }
  out.push(cur);
  return out.map(unescape);
}

function decodeValue(p) {
  const enc = (p.params.encoding || '').toLowerCase();
  if (enc === 'quoted-printable') return decodeCharset(decodeQP(p.value), p.params.charset || 'utf-8');
  return p.value;
}

const TYPE_WORDS = new Set(['work', 'home', 'cell', 'mobile', 'voice', 'fax', 'pager', 'pref', 'internet', 'main', 'other', 'iphone', 'text', 'video', 'x400', 'dom', 'intl', 'postal', 'parcel', 'msg', 'bbs', 'modem', 'car', 'isdn', 'pcs']);

export function parseVCards(text) {
  const cards = [];
  let cur = null; let rawLines = [];
  for (const line of unfold(text)) {
    const p = parseLine(line);
    if (p.name === 'BEGIN' && /^vcard$/i.test(p.value.trim())) { cur = { version: '', fn: '', n: null, org: '', title: '', role: '', nickname: '', bday: '', note: '', url: [], tel: [], email: [], adr: [], impp: [], photo: null, categories: [], anniversary: '', gender: '' }; rawLines = [line]; continue; }
    if (!cur) continue;
    rawLines.push(line);
    if (p.name === 'END') { cur.raw = `${rawLines.join('\r\n')}\r\n`; if (!cur.fn && cur.n) cur.fn = [cur.n.prefix, cur.n.given, cur.n.middle, cur.n.family, cur.n.suffix].filter(Boolean).join(' '); if (!cur.fn) cur.fn = cur.org || (cur.email[0] && cur.email[0].value) || ''; cards.push(cur); cur = null; continue; }
    const v = decodeValue(p);
    const types = (p.params.type || []).filter((x) => TYPE_WORDS.has(x) || !/^(encoding|charset)/.test(x));
    switch (p.name) {
      case 'VERSION': cur.version = v.trim(); break;
      case 'FN': cur.fn = unescape(v); break;
      case 'N': { const [family, given, middle, prefix, suffix] = splitEsc(v); cur.n = { family: family || '', given: given || '', middle: middle || '', prefix: prefix || '', suffix: suffix || '' }; break; }
      case 'ORG': cur.org = splitEsc(v).filter(Boolean).join(' · '); break;
      case 'TITLE': cur.title = unescape(v); break;
      case 'ROLE': cur.role = unescape(v); break;
      case 'NICKNAME': cur.nickname = unescape(v); break;
      case 'BDAY': cur.bday = v.trim(); break;
      case 'ANNIVERSARY': cur.anniversary = v.trim(); break;
      case 'GENDER': cur.gender = v.trim(); break;
      case 'NOTE': cur.note = unescape(v); break;
      case 'URL': cur.url.push(unescape(v).trim()); break;
      case 'TEL': cur.tel.push({ value: unescape(v).replace(/^tel:/i, '').trim(), types }); break;
      case 'EMAIL': cur.email.push({ value: unescape(v).replace(/^mailto:/i, '').trim(), types: types.filter((x) => x !== 'internet') }); break;
      case 'IMPP': case 'X-SKYPE': case 'X-JABBER': cur.impp.push(unescape(v).trim()); break;
      case 'CATEGORIES': cur.categories.push(...splitEsc(v, ',').map((x) => x.trim()).filter(Boolean)); break;
      case 'ADR': {
        const [pobox, ext, street, city, region, code, country] = splitEsc(v);
        cur.adr.push({ types, pobox: pobox || '', ext: ext || '', street: street || '', city: city || '', region: region || '', code: code || '', country: country || '', label: p.params.label ? unescape(p.params.label) : '' });
        break;
      }
      case 'PHOTO': case 'LOGO': {
        if (cur.photo && p.name === 'LOGO') break;
        const val = p.value.trim();
        const enc = (p.params.encoding || '').toLowerCase();
        const m = /^data:([^;,]+)?(;base64)?,(.*)$/i.exec(val);
        const t = (p.params.type && p.params.type.find((x) => /^(jpe?g|png|gif|image\/)/.test(x))) || p.params.mediatype || '';
        if (m) cur.photo = { mime: m[1] || 'image/jpeg', data: m[2] ? m[3] : '' , url: m[2] ? '' : val };
        else if (enc === 'b' || enc === 'base64') cur.photo = { mime: t.startsWith('image/') ? t : `image/${(t || 'jpeg').replace('jpg', 'jpeg')}`, data: val.replace(/\s+/g, '') };
        else if (/^https?:/i.test(val)) cur.photo = { mime: '', data: '', url: val };
        break;
      }
      default: break;
    }
  }
  return cards;
}

export function cardText(c) {
  return [c.fn, c.org, c.title, ...c.tel.map((x) => x.value), ...c.email.map((x) => x.value), ...c.adr.map(formatAdr), c.note].filter(Boolean).join('\n');
}
export const formatAdr = (a) => a.label || [a.street, a.ext, a.pobox, [a.code, a.city].filter(Boolean).join(' '), a.region, a.country].filter(Boolean).join(', ');
