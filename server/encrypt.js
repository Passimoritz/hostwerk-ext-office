'use strict';
/*
 * Password encryption of office files – pure JS with Node's crypto, nothing executed.
 *
 * OOXML (docx/xlsx/pptx/…, MS-OFFCRYPTO / ECMA-376 "agile" encryption, what Office 2010+ writes):
 *   encryptOoxml(zipBytes, password, { spinCount }) → CFB bytes (EncryptionInfo 4.4 XML + EncryptedPackage + \x06DataSpaces)
 *   decryptOoxml(cfbBytes, password) → zip bytes   (agile AES/SHA-1…512 and "standard" AES-128 ECB of Office 2007)
 *   isEncryptedOoxml(bytes)
 * ODF (odt/ods/odp/…, ODF 1.2 manifest encryption-data: AES-256-CBC, PBKDF2-HMAC-SHA1 over SHA-256(password)):
 *   encryptOdf(zipBytes, password, { iterations }) · decryptOdf(zipBytes, password) · isEncryptedOdf(bytes)
 * Errors: code 'EPASSWORD' (wrong password), 'EFORMAT' (not such a file), 'EALREADY' (already encrypted).
 */
const crypto = require('crypto');
const zlib = require('zlib');
const { readCfb, isCfb } = require('./cfb');
const { writeCfb } = require('./cfbwrite');
const ZP = require('./zippatch');
const { isZip } = require('./zip');

const fail = (code, msg, status = 400) => Object.assign(new Error(msg), { code, status });
const u32 = (n) => { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0); return b; };
const HASHES = { SHA1: ['sha1', 20], SHA256: ['sha256', 32], SHA384: ['sha384', 48], SHA512: ['sha512', 64], MD5: ['md5', 16] };
const hashOf = (alg, ...parts) => { const h = crypto.createHash(alg); for (const p of parts) h.update(p); return h.digest(); };
const BLK = { verifierInput: Buffer.from('fea7d2763b4b9e79', 'hex'), verifierValue: Buffer.from('d7aa0f6d3061344e', 'hex'), key: Buffer.from('146e0be7abacd0d6', 'hex'), hmacKey: Buffer.from('5fb2ad010cb9e1f6', 'hex'), hmacValue: Buffer.from('a0677f02b22c8433', 'hex') };
const fit = (b, n, padByte) => (b.length >= n ? b.subarray(0, n) : Buffer.concat([b, Buffer.alloc(n - b.length, padByte)]));
const aes = (mode, key, iv, data, enc) => { const name = `aes-${key.length * 8}-${mode}`; const c = enc ? crypto.createCipheriv(name, key, mode === 'ecb' ? null : iv) : crypto.createDecipheriv(name, key, mode === 'ecb' ? null : iv); c.setAutoPadding(false); return Buffer.concat([c.update(data), c.final()]); };
const pad16 = (b) => (b.length % 16 ? Buffer.concat([b, Buffer.alloc(16 - (b.length % 16))]) : b);

/* ---------------- OOXML agile ---------------- */
/** Iterated password hash H_n (MS-OFFCRYPTO 2.3.4.11) – spinCount ≤ 10 million */
function passwordHash(alg, salt, password, spin) {
  let h = hashOf(alg, salt, Buffer.from(String(password), 'utf16le'));
  const it = Buffer.alloc(4);
  for (let i = 0; i < spin; i++) { it.writeUInt32LE(i); h = crypto.createHash(alg).update(it).update(h).digest(); }
  return h;
}
const keyFrom = (alg, h, block, bytes) => fit(hashOf(alg, h, block), bytes, 0x36);

/** UNICODE-LP-P4: uint32 byte length + UTF-16LE, padded to 4 bytes */
const lpp4 = (s) => { const b = Buffer.from(s, 'utf16le'); return Buffer.concat([u32(b.length), b, Buffer.alloc((4 - (b.length % 4)) % 4)]); };
const ver = () => Buffer.from([1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0]);
function dataSpaces() {
  const version = Buffer.concat([lpp4('Microsoft.Container.DataSpaces'), ver()]);
  const ref = Buffer.concat([u32(1), u32(0), lpp4('EncryptedPackage'), lpp4('StrongEncryptionDataSpace')]);
  const map = Buffer.concat([u32(8), u32(1), u32(ref.length + 4), ref]);
  const info = Buffer.concat([u32(8), u32(1), lpp4('StrongEncryptionTransform')]);
  const tid = lpp4('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  const primary = Buffer.concat([u32(8 + tid.length), u32(1), tid, lpp4('Microsoft.Container.EncryptionTransform'), ver(), u32(0), u32(0), u32(0), u32(4)]);
  return {
    '\u0006DataSpaces/Version': version,
    '\u0006DataSpaces/DataSpaceMap': map,
    '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace': info,
    '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary': primary,
  };
}

function encryptOoxml(zipBytes, password, { spinCount = 100000 } = {}) {
  const pkg = Buffer.isBuffer(zipBytes) ? zipBytes : Buffer.from(zipBytes);
  if (isCfb(pkg)) throw fail('EALREADY', 'This file is already encrypted.', 409);
  if (!isZip(pkg)) throw fail('EFORMAT', 'This is not an Office Open XML file.');
  if (!password) throw fail('EPASSWORD', 'Please enter a password.');
  const alg = 'sha512'; const hs = 64; const keyBytes = 32;
  const keySalt = crypto.randomBytes(16); const pwSalt = crypto.randomBytes(16);
  const secret = crypto.randomBytes(keyBytes);
  // package: size + 4096-byte segments, IV = H(keySalt + segment index)
  const segs = [];
  for (let i = 0, n = 0; i < pkg.length; i += 4096, n++) {
    const iv = fit(hashOf(alg, keySalt, u32(n)), 16, 0x36);
    segs.push(aes('cbc', secret, iv, pad16(pkg.subarray(i, i + 4096)), true));
  }
  const size = Buffer.alloc(8); size.writeBigUInt64LE(BigInt(pkg.length));
  const encPkg = Buffer.concat([size, ...segs]);
  // data integrity
  const hmacKey = crypto.randomBytes(hs);
  const hmacVal = crypto.createHmac(alg, hmacKey).update(encPkg).digest();
  const encHmacKey = aes('cbc', secret, fit(hashOf(alg, keySalt, BLK.hmacKey), 16, 0x36), hmacKey, true);
  const encHmacVal = aes('cbc', secret, fit(hashOf(alg, keySalt, BLK.hmacValue), 16, 0x36), hmacVal, true);
  // password key encryptor
  const h = passwordHash(alg, pwSalt, password, spinCount);
  const verifier = crypto.randomBytes(16);
  const encVerIn = aes('cbc', keyFrom(alg, h, BLK.verifierInput, keyBytes), pwSalt, verifier, true);
  const encVerVal = aes('cbc', keyFrom(alg, h, BLK.verifierValue, keyBytes), pwSalt, hashOf(alg, verifier), true);
  const encKey = aes('cbc', keyFrom(alg, h, BLK.key, keyBytes), pwSalt, secret, true);
  const b64 = (b) => b.toString('base64');
  const common = `saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"`;
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><keyData ${common} saltValue="${b64(keySalt)}"/><dataIntegrity encryptedHmacKey="${b64(encHmacKey)}" encryptedHmacValue="${b64(encHmacVal)}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="${spinCount}" ${common} saltValue="${b64(pwSalt)}" encryptedVerifierHashInput="${b64(encVerIn)}" encryptedVerifierHashValue="${b64(encVerVal)}" encryptedKeyValue="${b64(encKey)}"/></keyEncryptor></keyEncryptors></encryption>`;
  const info = Buffer.concat([Buffer.from([4, 0, 4, 0]), u32(0x40), Buffer.from(xml, 'utf8')]);
  return writeCfb({ EncryptionInfo: info, EncryptedPackage: encPkg, ...dataSpaces() });
}

const attr = (tag, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag); return m ? m[1] : null; };

function decryptOoxml(bytes, password) {
  const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
  if (!isCfb(buf)) throw fail('EFORMAT', 'This file is not encrypted.');
  let s; try { s = readCfb(buf); } catch { throw fail('EFORMAT', 'This file is not encrypted.'); }
  const info = s.get('EncryptionInfo'); const pkg = s.get('EncryptedPackage');
  if (!info || !pkg) throw fail('EFORMAT', 'This is not an encrypted Office file (maybe an old binary .doc/.xls/.ppt).');
  const major = info.readUInt16LE(0); const minor = info.readUInt16LE(2);
  const size = Number(pkg.readBigUInt64LE(0));
  if (size > 1024 * 1048576) throw fail('EFORMAT', 'The file is too large.');
  if (major === 4 && minor === 4) {
    const xml = info.subarray(8).toString('utf8');
    const kd = /<keyData\b[^>]*>/.exec(xml); const ek = /<(?:\w+:)?encryptedKey\b[^>]*>/.exec(xml);
    if (!kd || !ek) throw fail('EFORMAT', 'Only password encryption is supported (no certificates).');
    const hk = (t) => { const a = String(attr(t, 'hashAlgorithm') || 'SHA1').toUpperCase(); if (!HASHES[a]) throw fail('EFORMAT', `Unsupported hash ${a}`); return HASHES[a][0]; };
    if (attr(kd[0], 'cipherAlgorithm') !== 'AES' || attr(ek[0], 'cipherAlgorithm') !== 'AES') throw fail('EFORMAT', 'Only AES encryption is supported.');
    const e = ek[0]; const alg = hk(e); const kbytes = Number(attr(e, 'keyBits')) / 8; const blk = Number(attr(e, 'blockSize') || 16);
    const spin = Number(attr(e, 'spinCount'));
    if (!(spin >= 0 && spin <= 10000000)) throw fail('EFORMAT', 'Invalid spin count');
    const salt = Buffer.from(attr(e, 'saltValue'), 'base64');
    const h = passwordHash(alg, salt, password, spin);
    const iv = fit(salt, blk, 0x36);
    const verIn = aes('cbc', keyFrom(alg, h, BLK.verifierInput, kbytes), iv, Buffer.from(attr(e, 'encryptedVerifierHashInput'), 'base64'), false);
    const verVal = aes('cbc', keyFrom(alg, h, BLK.verifierValue, kbytes), iv, Buffer.from(attr(e, 'encryptedVerifierHashValue'), 'base64'), false);
    const want = hashOf(alg, verIn.subarray(0, Number(attr(e, 'saltSize') || 16)));
    if (!crypto.timingSafeEqual(want, verVal.subarray(0, want.length))) throw fail('EPASSWORD', 'The password is not correct.', 403);
    const k0 = kd[0]; const kalg = hk(k0); const kb = Number(attr(k0, 'keyBits')) / 8;
    const secret = aes('cbc', keyFrom(alg, h, BLK.key, kbytes), iv, Buffer.from(attr(e, 'encryptedKeyValue'), 'base64'), false).subarray(0, kb);
    const keySalt = Buffer.from(attr(k0, 'saltValue'), 'base64'); const kblk = Number(attr(k0, 'blockSize') || 16);
    // integrity (if present): HMAC over the whole EncryptedPackage stream
    const di = /<dataIntegrity\b[^>]*>/.exec(xml);
    if (di) {
      try {
        const hkey = aes('cbc', secret, fit(hashOf(kalg, keySalt, BLK.hmacKey), kblk, 0x36), Buffer.from(attr(di[0], 'encryptedHmacKey'), 'base64'), false).subarray(0, HASHES[String(attr(k0, 'hashAlgorithm')).toUpperCase()][1]);
        const hval = aes('cbc', secret, fit(hashOf(kalg, keySalt, BLK.hmacValue), kblk, 0x36), Buffer.from(attr(di[0], 'encryptedHmacValue'), 'base64'), false);
        const calc = crypto.createHmac(kalg, hkey).update(pkg).digest();
        if (!crypto.timingSafeEqual(calc, hval.subarray(0, calc.length))) throw fail('EFORMAT', 'The encrypted file is damaged (integrity check failed).', 422);
      } catch (x) { if (x.code === 'EFORMAT') throw x; /* unusual integrity block – content is still checked by the zip reader */ }
    }
    const parts = [];
    const body = pkg.subarray(8);
    for (let i = 0, n = 0; i < body.length; i += 4096, n++) {
      const chunk = body.subarray(i, i + 4096);
      if (chunk.length % kblk) break;
      parts.push(aes('cbc', secret, fit(hashOf(kalg, keySalt, u32(n)), kblk, 0x36), chunk, false));
    }
    const out = Buffer.concat(parts).subarray(0, size);
    if (!isZip(out)) throw fail('EPASSWORD', 'The password is not correct.', 403);
    return out;
  }
  if ((major === 3 || major === 4) && minor === 2) return decryptStandard(info, pkg, password, size);
  throw fail('EFORMAT', `Unsupported encryption (version ${major}.${minor}).`);
}

/** Office 2007 "standard" encryption: AES-128/192/256 ECB, SHA-1, 50 000 rounds */
function decryptStandard(info, pkg, password, size) {
  const hsize = info.readUInt32LE(8);
  const hdr = info.subarray(12, 12 + hsize);
  const alg = hdr.readUInt32LE(8); const keyBits = hdr.readUInt32LE(16) || 128;
  if (![0x660e, 0x660f, 0x6610].includes(alg)) throw fail('EFORMAT', 'Only AES encryption is supported (RC4 files: open and save them in Office first).');
  const v = info.subarray(12 + hsize);
  const saltSize = v.readUInt32LE(0); const salt = v.subarray(4, 4 + saltSize);
  const encVer = v.subarray(4 + saltSize, 20 + saltSize);
  const encHash = v.subarray(24 + saltSize, 56 + saltSize);
  let h = hashOf('sha1', salt, Buffer.from(String(password), 'utf16le'));
  const it = Buffer.alloc(4);
  for (let i = 0; i < 50000; i++) { it.writeUInt32LE(i); h = hashOf('sha1', it, h); }
  const hf = hashOf('sha1', h, u32(0));
  const x = (c) => { const b = Buffer.alloc(64, c); for (let i = 0; i < hf.length; i++) b[i] ^= hf[i]; return hashOf('sha1', b); };
  const key = Buffer.concat([x(0x36), x(0x5c)]).subarray(0, keyBits / 8);
  const ver = aes('ecb', key, null, encVer, false);
  const vh = aes('ecb', key, null, encHash, false);
  if (!hashOf('sha1', ver).equals(vh.subarray(0, 20))) throw fail('EPASSWORD', 'The password is not correct.', 403);
  const body = pkg.subarray(8); const n = body.length - (body.length % 16);
  return aes('ecb', key, null, body.subarray(0, n), false).subarray(0, size);
}

const isEncryptedOoxml = (b) => { if (!isCfb(b)) return false; try { const s = readCfb(Buffer.from(b)); return s.has('EncryptionInfo') && s.has('EncryptedPackage'); } catch { return false; } };

/* ---------------- ODF ---------------- */
const M = 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0';
const xmlEsc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const keep = (name) => name === 'mimetype' || name.startsWith('META-INF/') || name.endsWith('/');

function encryptOdf(zipBytes, password, { iterations = 100000 } = {}) {
  const buf = Buffer.isBuffer(zipBytes) ? zipBytes : Buffer.from(zipBytes);
  if (!isZip(buf)) throw fail('EFORMAT', 'This is not an OpenDocument file.');
  if (!password) throw fail('EPASSWORD', 'Please enter a password.');
  const list = ZP.entries(buf);
  const man = list.find((e) => e.name === 'META-INF/manifest.xml');
  if (!man || !list.find((e) => e.name === 'mimetype')) throw fail('EFORMAT', 'This is not an OpenDocument file.');
  let xml = ZP.inflate(man).toString('utf8');
  if (/encryption-data/.test(xml)) throw fail('EALREADY', 'This file is already encrypted.', 409);
  const startKey = hashOf('sha256', Buffer.from(String(password), 'utf8'));
  const changes = {};
  for (const e of list) {
    if (keep(e.name)) continue;
    // LibreOffice drops the preview of encrypted files – a thumbnail would show the content
    if (/^Thumbnails\//.test(e.name)) { changes[e.name] = null; xml = xml.replace(new RegExp(`<manifest:file-entry\\b[^>]*manifest:full-path="${e.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), ''); continue; }
    const plain = ZP.inflate(e);
    const comp = zlib.deflateRawSync(plain, { level: 6 });
    const salt = crypto.randomBytes(16); const iv = crypto.randomBytes(16);
    const key = crypto.pbkdf2Sync(startKey, salt, iterations, 32, 'sha1');
    const c = crypto.createCipheriv('aes-256-cbc', key, iv); // PKCS#7 padding = valid W3C padding
    const enc = Buffer.concat([c.update(comp), c.final()]);
    const checksum = hashOf('sha256', comp.subarray(0, 1024)).toString('base64');
    changes[e.name] = { raw: enc, method: 0, crc: require('./zip').crc32(enc), size: enc.length };
    const data = `<manifest:encryption-data manifest:checksum-type="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k" manifest:checksum="${checksum}"><manifest:algorithm manifest:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc" manifest:initialisation-vector="${iv.toString('base64')}"/><manifest:start-key-generation manifest:start-key-generation-name="http://www.w3.org/2000/09/xmldsig#sha256" manifest:key-size="32"/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:key-size="32" manifest:iteration-count="${iterations}" manifest:salt="${salt.toString('base64')}"/></manifest:encryption-data>`;
    const re = new RegExp(`<manifest:file-entry\\b([^>]*?)manifest:full-path="${xmlEsc(e.name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"([^>]*?)(/>|>)`);
    const m = re.exec(xml);
    if (m) {
      let open = m[0].replace(/\s*\/?>$/, '').replace(/\smanifest:size="\d*"/, '');
      open += ` manifest:size="${plain.length}">`;
      if (m[3] === '/>') xml = xml.replace(m[0], `${open}${data}</manifest:file-entry>`);
      else xml = xml.replace(m[0], `${open}${data}`);
    } else xml = xml.replace(/<\/manifest:manifest>\s*$/, `<manifest:file-entry manifest:full-path="${xmlEsc(e.name)}" manifest:media-type="" manifest:size="${plain.length}">${data}</manifest:file-entry></manifest:manifest>`);
  }
  changes['META-INF/manifest.xml'] = xml;
  return ZP.patchZip(buf, changes, { first: ['mimetype'] });
}

function decryptOdf(zipBytes, password) {
  const buf = Buffer.isBuffer(zipBytes) ? zipBytes : Buffer.from(zipBytes);
  if (!isZip(buf)) throw fail('EFORMAT', 'This is not an OpenDocument file.');
  const list = ZP.entries(buf);
  const man = list.find((e) => e.name === 'META-INF/manifest.xml');
  if (!man) throw fail('EFORMAT', 'This is not an OpenDocument file.');
  let xml = ZP.inflate(man).toString('utf8');
  if (!/encryption-data/.test(xml)) throw fail('EFORMAT', 'This file is not encrypted.');
  const pw = Buffer.from(String(password), 'utf8');
  const changes = {};
  let checked = false;
  const entryRe = /<manifest:file-entry\b[^>]*[^/]>[\s\S]*?<\/manifest:file-entry>/g;
  for (const m of xml.match(entryRe) || []) {
    if (!/encryption-data/.test(m)) continue;
    const name = (/manifest:full-path="([^"]*)"/.exec(m) || [])[1];
    const e = list.find((x) => x.name === name.replace(/&amp;/g, '&'));
    if (!e) continue;
    const algo = (/manifest:algorithm-name="([^"]*)"/.exec(m) || [])[1] || '';
    const iv = Buffer.from((/manifest:initialisation-vector="([^"]*)"/.exec(m) || [])[1] || '', 'base64');
    const salt = Buffer.from((/manifest:salt="([^"]*)"/.exec(m) || [])[1] || '', 'base64');
    const iters = Number((/manifest:iteration-count="(\d+)"/.exec(m) || [])[1] || 1024);
    const ksize = Number((/manifest:key-derivation\b[^>]*manifest:key-size="(\d+)"/.exec(m) || [])[1] || 16);
    const skg = (/start-key-generation-name="([^"]*)"/.exec(m) || [])[1] || '';
    const startKey = /sha256/.test(skg) ? hashOf('sha256', pw) : hashOf('sha1', pw);
    if (!/(pbkdf2|PBKDF2)/.test(m) && /argon/i.test(m)) throw fail('EFORMAT', 'This ODF encryption (Argon2) is not supported.');
    if (iters > 10000000) throw fail('EFORMAT', 'Invalid iteration count');
    const key = crypto.pbkdf2Sync(startKey, salt, iters, ksize, 'sha1');
    let comp;
    try {
      if (/aes256-cbc|aes128-cbc|aes192-cbc/.test(algo)) {
        const d = crypto.createDecipheriv(`aes-${ksize * 8}-cbc`, key, iv); d.setAutoPadding(false);
        const cipher = e.method === 8 ? zlib.inflateRawSync(e.raw, { maxOutputLength: 512 * 1048576 }) : e.raw;
        if (!cipher.length || cipher.length % 16) throw fail('EPASSWORD', 'The password is not correct.', 403);
        const p = Buffer.concat([d.update(cipher), d.final()]);
        const padLen = p[p.length - 1];
        comp = padLen >= 1 && padLen <= 16 ? p.subarray(0, p.length - padLen) : p;
      } else if (/blowfish/i.test(algo)) {
        throw fail('EFORMAT', 'Old Blowfish-encrypted ODF files are not supported – open and save them in LibreOffice first.');
      } else throw fail('EFORMAT', `Unsupported ODF encryption ${algo}`);
    } catch (x) { if (x.code) throw x; throw fail('EPASSWORD', 'The password is not correct.', 403); }
    // a wrong key gives garbage that does not inflate; the 1K checksum is an extra check (LibreOffice variants differ)
    let plain;
    try { plain = zlib.inflateRawSync(comp, { maxOutputLength: 512 * 1048576 }); } catch { throw fail('EPASSWORD', 'The password is not correct.', 403); }
    if (!checked) {
      const want = (/manifest:checksum="([^"]*)"/.exec(m) || [])[1];
      const type = (/manifest:checksum-type="([^"]*)"/.exec(m) || [])[1] || '';
      if (want && !plain.length) {
        const alg = /sha256/i.test(type) ? 'sha256' : 'sha1';
        if (hashOf(alg, /1k/i.test(type) ? comp.subarray(0, 1024) : comp).toString('base64') !== want) throw fail('EPASSWORD', 'The password is not correct.', 403);
      }
      checked = true;
    }
    changes[e.name] = plain;
  }
  xml = xml.replace(/<manifest:encryption-data\b[\s\S]*?<\/manifest:encryption-data>/g, '');
  changes['META-INF/manifest.xml'] = xml;
  return ZP.patchZip(buf, changes, { first: ['mimetype'] });
}
const isEncryptedOdf = (b) => { try { if (!isZip(b)) return false; const man = ZP.entries(Buffer.from(b)).find((e) => e.name === 'META-INF/manifest.xml'); return !!man && /encryption-data/.test(ZP.inflate(man).toString('utf8')); } catch { return false; } };

module.exports = { encryptOoxml, decryptOoxml, isEncryptedOoxml, encryptOdf, decryptOdf, isEncryptedOdf, passwordHash };
