import test from 'node:test';
import assert from 'node:assert/strict';
import { readMail, decodeWords, decodeQP, parseParams, parseAddresses, htmlToText } from '../ui/src/editors/viewer/mime.js';

const enc = (s) => new TextEncoder().encode(s);
const latin1 = (s) => Uint8Array.from([...s].map((c) => c.charCodeAt(0)));

test('RFC 2047 encoded words (B and Q, adjacent words joined)', () => {
  assert.equal(decodeWords('=?UTF-8?B?w5xiZXJ3ZWlzdW5n?='), 'Überweisung');
  assert.equal(decodeWords('=?ISO-8859-1?Q?Gr=FC=DFe_aus_K=F6ln?='), 'Grüße aus Köln');
  assert.equal(decodeWords('=?UTF-8?Q?a?= =?UTF-8?Q?b?= c'), 'ab c');
  assert.equal(decodeWords('Re: =?utf-8?q?Caf=C3=A9?= heute'), 'Re: Café heute');
  // a multi-byte character split across two words
  assert.equal(decodeWords('=?UTF-8?B?w6Q=?= =?UTF-8?B?w7Y=?='), 'äö');
});
test('quoted-printable body', () => {
  assert.equal(new TextDecoder().decode(decodeQP('Hallo=20Welt=\r\nweiter =C3=A4')), 'Hallo Weltweiter ä');
});
test('parameters with quotes and RFC 2231', () => {
  const p = parseParams("attachment; filename*=UTF-8''Rechnung%20M%C3%A4rz.pdf");
  assert.equal(p.value, 'attachment');
  assert.equal(p.params.filename, 'Rechnung März.pdf');
  const q = parseParams('text/plain; charset="iso-8859-1"; format=flowed');
  assert.equal(q.params.charset, 'iso-8859-1');
  const c = parseParams('application/pdf; name*0="long"; name*1="name.pdf"');
  assert.equal(c.params.name, 'longname.pdf');
});
test('addresses', () => {
  assert.deepEqual(parseAddresses('"Doe, John" <j@d.e>, lena@x.de, =?UTF-8?Q?J=C3=BCrgen?= <jw@x.de>'), [
    { name: 'Doe, John', email: 'j@d.e' }, { name: '', email: 'lena@x.de' }, { name: 'Jürgen', email: 'jw@x.de' }]);
});
test('multipart/mixed with alternative, related inline image and base64 attachment', () => {
  const raw = [
    'From: =?UTF-8?Q?J=C3=BCrgen?= <jw@example.com>', 'To: lena@example.com', 'Cc: a@b.c', 'Subject: =?UTF-8?B?VGVzdCDDpA==?=',
    'Date: Tue, 05 Mar 2024 10:15:00 +0100', 'MIME-Version: 1.0', 'Content-Type: multipart/mixed; boundary="MIX"', '',
    'preamble', '--MIX', 'Content-Type: multipart/alternative; boundary=ALT', '',
    '--ALT', 'Content-Type: text/plain; charset=iso-8859-1', 'Content-Transfer-Encoding: quoted-printable', '', 'Gr=FC=DFe', '',
    '--ALT', 'Content-Type: multipart/related; boundary=REL', '',
    '--REL', 'Content-Type: text/html; charset=utf-8', '', '<p>GrÃ¼Ã\u009fe <img src="cid:logo@x"></p>',
    '--REL', 'Content-Type: image/png', 'Content-ID: <logo@x>', 'Content-Transfer-Encoding: base64', '', 'iVBORw0KGgo=',
    '--REL--', '--ALT--',
    '--MIX', 'Content-Type: application/pdf; name="a.pdf"', 'Content-Disposition: attachment; filename="Rechnung.pdf"', 'Content-Transfer-Encoding: base64', '', 'JVBERi0xLjQ=', '--MIX--', '',
  ].join('\r\n');
  const m = readMail(latin1(raw));
  assert.equal(m.subject, 'Test ä');
  assert.deepEqual(m.from, [{ name: 'Jürgen', email: 'jw@example.com' }]);
  assert.equal(m.to[0].email, 'lena@example.com');
  assert.equal(m.cc[0].email, 'a@b.c');
  assert.equal(m.date.toISOString(), '2024-03-05T09:15:00.000Z');
  assert.equal(m.text.trim(), 'Grüße');
  assert.match(m.html, /Grüße <img src="cid:logo@x">/);
  const img = m.attachments.find((a) => a.cid === 'logo@x');
  assert.ok(img && img.inline);
  assert.deepEqual([...img.bytes.slice(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
  const pdf = m.attachments.find((a) => a.name === 'Rechnung.pdf');
  assert.equal(new TextDecoder().decode(pdf.bytes), '%PDF-1.4');
});
test('single part plain text with windows-1252 and folded subject', () => {
  const m = readMail(latin1('Subject: Hallo\r\n Welt\r\nContent-Type: text/plain; charset=windows-1252\r\n\r\nPreis: 5 \x80\r\n'));
  assert.equal(m.subject, 'Hallo Welt');
  assert.equal(m.text.trim(), 'Preis: 5 €');
});
test('utf-8 body without charset and html to text', () => {
  const m = readMail(enc('Subject: x\n\nÄrger\n'));
  assert.equal(m.text.trim(), 'Ärger');
  assert.equal(htmlToText('<p>a<br>b</p><style>x{}</style>&amp;'), 'a\nb\n&');
});
test('raw UTF-8 headers (RFC 6532) and raw Windows-1252 headers', () => {
  const m = readMail(enc('Subject: Projektstand und nächste Schritte für Q4\r\nFrom: Jürgen <j@x.de>\r\n\r\nHi\r\n'));
  assert.equal(m.subject, 'Projektstand und nächste Schritte für Q4');
  assert.equal(m.from[0].name, 'Jürgen');
  const w = readMail(latin1('Subject: Gr\xfc\xdfe\r\n\r\nx\r\n'));
  assert.equal(w.subject, 'Grüße');
});
