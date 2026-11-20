import test from 'node:test';
import assert from 'node:assert/strict';
import nodeCrypto from 'node:crypto';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { md5, rc4, sha256, sha384, sha512, aesEncryptCbc, aesDecryptCbc, aesEncryptEcbBlock } from '../ui/src/editors/pdf/crypto.js';
import { encryptPdf, openProtected, isEncryptedBytes, permissionsToP, pToPermissions, createProtection } from '../ui/src/editors/pdf/security.js';

const hex = (b) => Buffer.from(b).toString('hex');
const H = (s) => Uint8Array.from(Buffer.from(s, 'hex'));
const rnd = (n) => Uint8Array.from(nodeCrypto.randomBytes(n));

test('hashes match node:crypto', () => {
  for (const len of [0, 1, 3, 55, 56, 63, 64, 65, 111, 112, 127, 128, 129, 1000, 5000]) {
    const d = rnd(len);
    assert.equal(hex(md5(d)), nodeCrypto.createHash('md5').update(d).digest('hex'), `md5 ${len}`);
    assert.equal(hex(sha256(d)), nodeCrypto.createHash('sha256').update(d).digest('hex'), `sha256 ${len}`);
    assert.equal(hex(sha384(d)), nodeCrypto.createHash('sha384').update(d).digest('hex'), `sha384 ${len}`);
    assert.equal(hex(sha512(d)), nodeCrypto.createHash('sha512').update(d).digest('hex'), `sha512 ${len}`);
  }
});

test('AES FIPS-197 vectors and CBC against node:crypto', () => {
  assert.equal(hex(aesEncryptEcbBlock(H('000102030405060708090a0b0c0d0e0f'), H('00112233445566778899aabbccddeeff'))), '69c4e0d86a7b0430d8cdb78070b4c55a');
  assert.equal(hex(aesEncryptEcbBlock(H('000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f'), H('00112233445566778899aabbccddeeff'))), '8ea2b7ca516745bfeafc49904b496089');
  for (const kl of [16, 32]) {
    for (const len of [0, 5, 16, 31, 100]) {
      const key = rnd(kl); const iv = rnd(16); const d = rnd(len);
      const ours = aesEncryptCbc(key, iv, d);
      const c = nodeCrypto.createCipheriv(kl === 16 ? 'aes-128-cbc' : 'aes-256-cbc', key, iv);
      assert.equal(hex(ours), hex(Buffer.concat([c.update(d), c.final()])));
      assert.equal(hex(aesDecryptCbc(key, iv, ours)), hex(d));
    }
  }
});

test('RC4 matches a known vector', () => {
  // "Key" / "Plaintext" → BBF316E8D940AF0AD3
  assert.equal(hex(rc4(new TextEncoder().encode('Key'), new TextEncoder().encode('Plaintext'))), 'bbf316e8d940af0ad3');
});

test('permissions round trip', () => {
  assert.equal(permissionsToP({}), -4);
  assert.deepEqual(pToPermissions(permissionsToP({ print: false, copy: false })), { print: false, modify: true, copy: false, annotate: true, fill: true, assemble: true });
});

async function sample() {
  const d = await PDFDocument.create();
  const f = await d.embedFont(StandardFonts.Helvetica);
  d.addPage([300, 200]).drawText('Geheimer Text 123', { x: 20, y: 100, size: 14, font: f });
  d.addPage([300, 200]).drawText('Seite zwei', { x: 20, y: 100, size: 14, font: f });
  d.setTitle('Vertraulich');
  return d.save();
}
async function pdfjsText(bytes, password) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: bytes.slice(), password, isEvalSupported: false, verbosity: 0 }).promise;
  let s = '';
  for (let i = 1; i <= doc.numPages; i++) s += (await (await doc.getPage(i)).getTextContent()).items.map((x) => x.str).join('') + '\n';
  const meta = await doc.getMetadata();
  await doc.destroy();
  return { text: s, title: meta.info.Title };
}

test('AES-256 (R6): encrypt → pdf.js opens with the password, our reader decrypts again', async () => {
  const plain = await sample();
  const enc = await encryptPdf(plain, { userPassword: 'geheim', ownerPassword: 'chef', permissions: { copy: false } });
  assert.ok(isEncryptedBytes(enc));
  assert.ok(!isEncryptedBytes(plain));
  assert.ok(!Buffer.from(enc).includes('Geheimer'), 'text must not be readable');
  const r = await pdfjsText(enc, 'geheim');
  assert.match(r.text, /Geheimer Text 123/);
  assert.equal(r.title, 'Vertraulich');
  await assert.rejects(() => pdfjsText(enc, 'falsch'));
  await assert.rejects(() => openProtected(enc, 'falsch'), (e) => e.code === 'PASSWORD');
  const u = await openProtected(enc, 'geheim');
  assert.equal(u.protection.owner, false);
  assert.equal(u.protection.permissions.copy, false);
  const o = await openProtected(enc, 'chef');
  assert.equal(o.protection.owner, true);
  assert.ok(!isEncryptedBytes(o.plain));
  assert.match((await pdfjsText(o.plain)).text, /Seite zwei/);
  // save again with the same protection: both passwords still work
  const again = await encryptPdf(o.plain, { keep: o.protection });
  assert.match((await pdfjsText(again, 'geheim')).text, /Geheimer/);
  assert.match((await pdfjsText(again, 'chef')).text, /Geheimer/);
});

test('empty open password (restrictions only) opens without asking', async () => {
  const enc = await encryptPdf(await sample(), { userPassword: '', ownerPassword: 'x', permissions: { modify: false } });
  const r = await openProtected(enc, '');
  assert.equal(r.protection.permissions.modify, false);
  assert.match((await pdfjsText(r.plain)).text, /Geheimer/);
});

test('RC4 and AES-128 files (made with the keep path of other revisions) decrypt', async () => {
  // build R3/RC4-128 and R4/AESV2 protections by hand via the handler of a fake Encrypt dict, encrypt, decrypt
  const { makeHandler } = await import('../ui/src/editors/pdf/security.js');
  for (const variant of [{ V: 2, R: 3, Length: 128 }, { V: 4, R: 4, Length: 128, CF: { StdCF: { CFM: 'AESV2', Length: 16 } }, StmF: 'StdCF', StrF: 'StdCF' }]) {
    const id = rnd(16);
    // compute O and U like Acrobat (algorithms 3, 5) for user "u", owner "o"
    const PAD = H('28bf4e5e4e758a4164004e56fffa01082e2e00b6d0683e802f0ca9fe6453697a');
    const pad = (s) => { const b = Buffer.from(s, 'latin1'); return Uint8Array.from(Buffer.concat([b, Buffer.from(PAD.slice(0, 32 - b.length))])); };
    let ok = md5(pad('o')); for (let i = 0; i < 50; i++) ok = md5(ok);
    let O = rc4(ok.slice(0, 16), pad('u')); for (let i = 1; i <= 19; i++) O = rc4(ok.slice(0, 16).map((b) => b ^ i), O);
    const P = -4;
    const le = Uint8Array.from([P & 255, (P >>> 8) & 255, (P >>> 16) & 255, (P >>> 24) & 255]);
    let k = md5(Uint8Array.from([...pad('u'), ...O, ...le, ...id])); for (let i = 0; i < 50; i++) k = md5(k.slice(0, 16)); k = k.slice(0, 16);
    let U = rc4(k, md5(Uint8Array.from([...PAD, ...id]))); for (let i = 1; i <= 19; i++) U = rc4(k.map((b) => b ^ i), U);
    U = Uint8Array.from([...U, ...new Uint8Array(16)]);
    const enc = { ...variant, P, O, U, EncryptMetadata: true };
    const prot = { handler: makeHandler(enc, id), key: k, ids: [id, id], encrypt: enc };
    const file = await encryptPdf(await sample(), { keep: prot });
    assert.match((await pdfjsText(file, 'u')).text, /Geheimer/, `pdf.js reads V${variant.V}`);
    const back = await openProtected(file, 'u');
    assert.match((await pdfjsText(back.plain)).text, /Geheimer Text 123/);
    assert.equal((await openProtected(file, 'o')).protection.owner, true);
  }
});

test('createProtection gives a fresh key every time', () => {
  assert.notEqual(hex(createProtection({ userPassword: 'a' }).key), hex(createProtection({ userPassword: 'a' }).key));
});
