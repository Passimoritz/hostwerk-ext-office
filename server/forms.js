'use strict';
/*
 * Forms from sheets: a public page (token link, optional expiry and password, honeypot, rate limit) whose answers are
 * appended as rows to a spreadsheet (xlsx/ods/csv via sheet.appendRow – the rest of the file stays byte-identical).
 * While the file is open in an editor (soft lock) answers wait in a queue; they are written as soon as the lock is gone
 * (or the editor takes them: GET /office/forms/:id/pending + POST …/pending/ack).
 *
 *   validateForm(body) → { title, description, sheet, columns, timestamp, public: { enabled, expires, password? } }
 *   new Forms({ file }) · list() · get(id) · byToken(token) · add(def) · update(id, patch) · remove(id)
 *   checkAnswer(form, body) → values (throws 400 with field errors)
 *   formPage(lang, form, { base, errors, values, done, locked, wrongPassword, nonce })
 */
const crypto = require('crypto');
const { readJson, writeJson } = require('./store');

const TYPES = ['text', 'textarea', 'number', 'email', 'date', 'choice', 'checkbox', 'tel', 'url'];
const fail = (msg, vars, extra) => Object.assign(new Error(msg), { status: 400, vars }, extra || {});
const TOKEN = /^[A-Za-z0-9_-]{32}$/;

function validateForm(b = {}) {
  const cols = Array.isArray(b.columns) ? b.columns : [];
  if (!cols.length || cols.length > 60) throw fail('A form needs 1 to 60 fields.');
  const columns = cols.map((c0, i) => {
    const c = typeof c0 === 'string' ? { label: c0 } : c0 || {};
    const type = TYPES.includes(c.type) ? c.type : 'text';
    const label = String(c.label || c.name || '').trim().slice(0, 200);
    if (!label) throw fail('Field {n} needs a name.', { n: i + 1 });
    const options = type === 'choice' ? (Array.isArray(c.options) ? c.options : String(c.options || '').split(/\r?\n|;/)).map((x) => String(x).trim().slice(0, 200)).filter(Boolean).slice(0, 100) : undefined;
    if (type === 'choice' && !options.length) throw fail('The choice field {name} needs options.', { name: label });
    return { id: `f${i}`, label, type, required: !!c.required, help: String(c.help || '').slice(0, 300), ...(options ? { options } : {}), ...(type === 'number' && c.min != null ? { min: Number(c.min) } : {}), ...(type === 'number' && c.max != null ? { max: Number(c.max) } : {}) };
  });
  const p = b.public || {};
  return { title: String(b.title || '').trim().slice(0, 200) || 'Form', description: String(b.description || '').slice(0, 2000), sheet: String(b.sheet || '').slice(0, 120), columns, timestamp: b.timestamp !== false, thanks: String(b.thanks || '').slice(0, 500), notify: ['hostwerk', 'discord', 'ntfy', 'mail', 'all'].includes(b.notify) ? b.notify : null, public: { enabled: p.enabled !== false, expires: p.expires, password: p.password ? String(p.password) : '' } };
}

function checkAnswer(form, body = {}) {
  const errors = {};
  const values = form.columns.map((c) => {
    let v = body[c.id];
    if (Array.isArray(v)) v = v[v.length - 1];
    v = v == null ? '' : String(v);
    if (c.type === 'checkbox') return v === 'on' || v === 'true' || v === '1';
    v = v.replace(/\r\n/g, '\n').trim().slice(0, c.type === 'textarea' ? 5000 : 500);
    if (!v) { if (c.required) errors[c.id] = 'required'; return ''; }
    if (c.type === 'number') { const n = Number(v.replace(',', '.')); if (!Number.isFinite(n)) { errors[c.id] = 'number'; return v; } if ((c.min != null && n < c.min) || (c.max != null && n > c.max)) errors[c.id] = 'range'; return n; }
    if (c.type === 'email' && !/^[^\s@]{1,64}@[^\s@]{1,255}\.[^\s@]{2,}$/.test(v)) errors[c.id] = 'email';
    if (c.type === 'url' && !/^https?:\/\/\S+$/i.test(v)) errors[c.id] = 'url';
    if (c.type === 'date' && !/^\d{4}-\d{2}-\d{2}$/.test(v)) errors[c.id] = 'date';
    if (c.type === 'tel' && !/^[+\d ()/.-]{3,40}$/.test(v)) errors[c.id] = 'tel';
    if (c.type === 'choice' && !c.options.includes(v)) errors[c.id] = 'choice';
    return v;
  });
  if (Object.keys(errors).length) throw fail('Please check the marked fields.', null, { fields: errors });
  return values;
}

class Forms {
  constructor({ file }) { this.file = file; this.list0 = readJson(file, []); }
  save() { writeJson(this.file, this.list0); }
  list() { return this.list0; }
  get(id) { return this.list0.find((f) => f.id === id) || null; }
  byToken(token) { if (!TOKEN.test(String(token || ''))) return null; return this.list0.find((f) => f.token === token) || null; }
  add(f) { if (this.list0.length >= 200) throw fail('At most 200 forms – remove old ones first.'); this.list0.push(f); this.save(); return f; }
  update(id, patch) { const f = this.get(id); if (!f) throw Object.assign(new Error('Form not found'), { status: 404 }); Object.assign(f, patch); this.save(); return f; }
  remove(id) { const n = this.list0.length; this.list0 = this.list0.filter((f) => f.id !== id); this.save(); return n !== this.list0.length; }
  rekey(a, b) { let ch = false; for (const f of this.list0) if (f.key === a || f.key.startsWith(`${a}/`)) { f.key = b + f.key.slice(a.length); if (!f.root) f.path = f.key; ch = true; } if (ch) this.save(); }
}
const newToken = () => crypto.randomBytes(24).toString('base64url');
const open = (f, now = Date.now()) => !!f && !f.closed && f.public && f.public.enabled !== false && !(f.public.expires && f.public.expires <= now);

const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const T = {
  de: { send: 'Absenden', required: 'Pflichtfeld', thanks: 'Danke! Deine Antwort ist angekommen.', again: 'Noch eine Antwort senden', closed: 'Dieses Formular nimmt keine Antworten mehr an.', password: 'Passwort', open: 'Öffnen', locked: 'Dieses Formular ist mit einem Passwort geschützt.', wrong: 'Das Passwort stimmt nicht.', check: 'Bitte prüfe die markierten Felder.', footer: 'Formular mit Hostwerk', choose: 'Bitte wählen …', err: { required: 'Bitte ausfüllen.', number: 'Bitte eine Zahl eingeben.', range: 'Die Zahl liegt außerhalb des erlaubten Bereichs.', email: 'Bitte eine gültige E-Mail-Adresse eingeben.', url: 'Bitte eine Adresse mit https:// eingeben.', date: 'Bitte ein Datum wählen.', tel: 'Bitte eine Telefonnummer eingeben.', choice: 'Bitte eine der Möglichkeiten wählen.' } },
  en: { send: 'Send', required: 'Required', thanks: 'Thank you! Your answer has arrived.', again: 'Send another answer', closed: 'This form no longer accepts answers.', password: 'Password', open: 'Open', locked: 'This form is protected with a password.', wrong: 'The password is not correct.', check: 'Please check the marked fields.', footer: 'Form made with Hostwerk', choose: 'Please choose …', err: { required: 'Please fill this in.', number: 'Please enter a number.', range: 'The number is outside the allowed range.', email: 'Please enter a valid e-mail address.', url: 'Please enter an address with https://.', date: 'Please choose a date.', tel: 'Please enter a phone number.', choice: 'Please choose one of the options.' } },
};
const CSS = `:root{color-scheme:light dark;--bg:#f6f7f9;--panel:#fff;--text:#14161a;--muted:#5d6470;--line:#d9dde3;--accent:#16a34a;--err:#dc2626}
@media (prefers-color-scheme:dark){:root{--bg:#0f1115;--panel:#171a21;--text:#e8eaf0;--muted:#9aa1ad;--line:#2f3542;--accent:#22c55e;--err:#f87171}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:16px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:640px;margin:0 auto;padding:24px 16px 48px}.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:24px;display:flex;flex-direction:column;gap:18px}
h1{font-size:22px;margin:0}.desc{color:var(--muted);white-space:pre-wrap;margin:0}label{display:flex;flex-direction:column;gap:6px;font-weight:600}.help{font-weight:400;color:var(--muted);font-size:14px}
input,textarea,select{font:inherit;padding:10px 12px;border-radius:9px;border:1px solid var(--line);background:var(--bg);color:var(--text);width:100%}textarea{min-height:110px;resize:vertical}
input:focus-visible,textarea:focus-visible,select:focus-visible,button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}.check{flex-direction:row;align-items:center;gap:10px}.check input{width:auto}
.req{color:var(--err);font-weight:400}.e{color:var(--err);font-weight:400;font-size:14px}.bad input,.bad textarea,.bad select{border-color:var(--err)}
button{font:inherit;font-weight:600;padding:11px 18px;border-radius:9px;border:0;background:var(--accent);color:#fff;cursor:pointer;align-self:flex-start}
.hp{position:absolute;left:-10000px;width:1px;height:1px;overflow:hidden}.alert{color:var(--err)}footer{text-align:center;color:var(--muted);font-size:13px;margin-top:16px}a{color:var(--accent)}`;
const shell = (lang, title, body) => `<!doctype html><html lang="${lang}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><meta name="robots" content="noindex"><title>${esc(title)}</title><style>${CSS}</style></head><body><main>${body}<footer>${(T[lang] || T.en).footer}</footer></main></body></html>`;

/** o: { base, errors: { fid: code }, values: { fid }, done, closed, locked, wrongPassword, stamp } */
function formPage(lang, f, o = {}) {
  const t = T[lang] || T.en;
  if (o.closed) return shell(lang, f.title, `<div class="card"><h1>${esc(f.title)}</h1><p class="desc">${t.closed}</p></div>`);
  if (o.locked) return shell(lang, f.title, `<form class="card" method="post" action="${esc(o.base)}/unlock"><h1>${esc(f.title)}</h1><p class="desc">${t.locked}</p><label>${t.password}<input type="password" name="password" autocomplete="current-password" required autofocus></label>${o.wrongPassword ? `<p class="alert" role="alert">${t.wrong}</p>` : ''}<button type="submit">${t.open}</button></form>`);
  if (o.done) return shell(lang, f.title, `<div class="card"><h1>${esc(f.title)}</h1><p role="status">${esc(f.thanks || t.thanks)}</p><a href="${esc(o.base)}">${t.again}</a></div>`);
  const errs = o.errors || {}; const vals = o.values || {};
  const fields = f.columns.map((c) => {
    const id = `i_${c.id}`; const bad = errs[c.id] ? ' bad' : ''; const v = vals[c.id] == null ? '' : String(vals[c.id]);
    const req = c.required ? ' required' : ''; const star = c.required ? ` <span class="req" aria-hidden="true">*</span>` : '';
    const help = c.help ? `<span class="help">${esc(c.help)}</span>` : '';
    const msg = errs[c.id] ? `<span class="e" id="${id}_e">${t.err[errs[c.id]] || t.check}</span>` : '';
    const aria = errs[c.id] ? ` aria-invalid="true" aria-describedby="${id}_e"` : '';
    if (c.type === 'checkbox') return `<label class="check${bad}" for="${id}"><input type="checkbox" id="${id}" name="${c.id}"${vals[c.id] ? ' checked' : ''}${aria}> ${esc(c.label)}${star}</label>${msg}`;
    let input;
    if (c.type === 'textarea') input = `<textarea id="${id}" name="${c.id}" maxlength="5000"${req}${aria}>${esc(v)}</textarea>`;
    else if (c.type === 'choice') input = `<select id="${id}" name="${c.id}"${req}${aria}><option value="">${t.choose}</option>${c.options.map((x) => `<option${x === v ? ' selected' : ''}>${esc(x)}</option>`).join('')}</select>`;
    else input = `<input id="${id}" name="${c.id}" type="${c.type === 'number' ? 'text" inputmode="decimal' : c.type}" value="${esc(v)}" maxlength="500"${req}${aria}${c.type === 'email' ? ' autocomplete="email"' : ''}>`;
    return `<label class="${bad.trim()}" for="${id}">${esc(c.label)}${star}${help}${input}${msg}</label>`;
  }).join('');
  return shell(lang, f.title, `<form class="card" method="post" action="${esc(o.base)}" novalidate><h1>${esc(f.title)}</h1>${f.description ? `<p class="desc">${esc(f.description)}</p>` : ''}${Object.keys(errs).length ? `<p class="alert" role="alert">${t.check}</p>` : ''}${fields}
<div class="hp" aria-hidden="true"><label>Website<input type="text" name="website" tabindex="-1" autocomplete="off"></label></div><input type="hidden" name="_t" value="${esc(o.stamp || '')}"><button type="submit">${t.send}</button></form>`);
}

/** time stamp proof: the page was loaded at least 2 s before sending (bots post at once) and less than 1 day ago */
const stamp = (secret, token, now = Date.now()) => { const t = String(now); return `${t}.${crypto.createHmac('sha256', secret).update(`${token}\n${t}`).digest('base64url').slice(0, 22)}`; };
function stampOk(secret, token, v, now = Date.now()) {
  const [t, sig] = String(v || '').split('.');
  if (!t || !sig) return false;
  const want = crypto.createHmac('sha256', secret).update(`${token}\n${t}`).digest('base64url').slice(0, 22);
  if (want.length !== sig.length || !crypto.timingSafeEqual(Buffer.from(want), Buffer.from(sig))) return false;
  const age = now - Number(t);
  return age >= 2000 && age < 86400000;
}

module.exports = { Forms, validateForm, checkAnswer, formPage, newToken, open, stamp, stampOk, TYPES };
