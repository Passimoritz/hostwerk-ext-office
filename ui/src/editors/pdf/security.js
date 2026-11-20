/*
 * PDF standard security handler (ISO 32000-2 §7.6.4) on top of pdf-lib – pure, no DOM, tested in
 * tests/pdf-crypto.test.js (round trips + pdf.js opening our encrypted files).
 *
 *   isEncryptedBytes(bytes)                       → quick byte scan for an /Encrypt entry
 *   openProtected(bytes, password)                → { plain: Uint8Array, protection } | throws { code: 'PASSWORD' | 'UNSUPPORTED' }
 *        protection = { handler, key, id, encrypt (raw Encrypt dict values), owner: bool, permissions }
 *   encryptPdf(plainBytes, spec)                  → encrypted bytes
 *        spec = { keep: protection }               – same passwords, same file key (both passwords keep working)
 *             | { userPassword, ownerPassword, permissions: { print, modify, copy, annotate, fill, assemble } } – new AES-256 (R6)
 *
 * Revisions 2–4 (RC4 40/128, AES-128) and 5/6 (AES-256) are read; new protection is always AES-256 (R6).
 */
import {
  PDFDocument, PDFParser, PDFObjectStreamParser, PDFXRefStreamParser, PDFWriter, PDFName, PDFDict, PDFArray,
  PDFString, PDFHexString, PDFRawStream, PDFNumber, PDFBool, PDFRef,
} from 'pdf-lib';
import { md5, rc4, sha256, sha384, sha512, aesEncryptCbc, aesDecryptCbc, aesEncryptEcbBlock, concat, randomBytes } from './crypto.js';

const PAD = Uint8Array.from([0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08, 0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a]);
const SALT = Uint8Array.from([0x73, 0x41, 0x6c, 0x54]); // "sAlT"
const ENDOBJ = [101, 110, 100, 111, 98, 106];

const eq = (a, b, n = a.length) => { if (a.length < n || b.length < n) return false; for (let i = 0; i < n; i++) if (a[i] !== b[i]) return false; return true; };
const le32 = (v) => Uint8Array.from([v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]);
const latin1 = (s) => Uint8Array.from(String(s || ''), (c) => { const n = c.charCodeAt(0); return n < 256 ? n : 63; }).slice(0, 32);
const utf8pw = (s) => new TextEncoder().encode(String(s || '').normalize('NFKC')).slice(0, 127);
const pad32 = (pw) => concat(pw.slice(0, 32), PAD.slice(0, 32 - Math.min(32, pw.length)));
const toHex = (b) => Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('');

/** quick check without parsing: the trailer (or xref stream dictionary) is never compressed */
export function isEncryptedBytes(bytes) {
  const pat = [47, 69, 110, 99, 114, 121, 112, 116]; // "/Encrypt"
  outer: for (let i = 0, n = bytes.length - pat.length; i <= n; i++) {
    if (bytes[i] !== 47) continue;
    for (let k = 1; k < pat.length; k++) if (bytes[i + k] !== pat[k]) continue outer;
    const nx = bytes[i + pat.length];
    if (nx === 32 || nx === 10 || nx === 13 || nx === 60 || nx === 9 || (nx >= 48 && nx <= 57)) return true;
  }
  return false;
}

/* ---------------- handler ---------------- */

/** permissions → P (int32) */
export function permissionsToP(p = {}) {
  const on = (k) => p[k] !== false;
  let v = 0xfffff0c0;
  if (on('print')) v |= 0x4 | 0x800;
  if (on('modify')) v |= 0x8;
  if (on('copy')) v |= 0x10 | 0x200;
  if (on('annotate')) v |= 0x20;
  if (on('fill')) v |= 0x100;
  if (on('assemble')) v |= 0x400;
  return v | 0;
}
export function pToPermissions(P) {
  return { print: !!(P & 0x4), modify: !!(P & 0x8), copy: !!(P & 0x10), annotate: !!(P & 0x20), fill: !!(P & 0x100), assemble: !!(P & 0x400) };
}

/** enc: { V, R, Length, O, U, OE?, UE?, Perms?, P, EncryptMetadata, StmF, StrF, CF: { name: { CFM, Length } } } */
export function makeHandler(enc, id0) {
  const V = enc.V || 0; const R = enc.R || 2;
  if (![1, 2, 4, 5].includes(V) || R < 2 || R > 6) throw Object.assign(new Error(`Unsupported PDF encryption (V ${V}, R ${R})`), { code: 'UNSUPPORTED' });
  const method = (name) => {
    if (V < 4) return 'rc4';
    if (!name || name === 'Identity') return 'none';
    const cf = (enc.CF || {})[name];
    const cfm = cf && cf.CFM;
    if (cfm === 'AESV3') return 'aes256';
    if (cfm === 'AESV2') return 'aes128';
    if (cfm === 'V2') return 'rc4';
    if (cfm === 'None') return 'none';
    return V === 5 ? 'aes256' : 'rc4';
  };
  let n = V === 1 ? 5 : Math.round((enc.Length || 40) / 8);
  if (V === 4) n = 16;
  if (V === 5) n = 32;
  if (R === 2) n = 5;
  return { V, R, n, enc, id0: id0 || new Uint8Array(0), stm: method(enc.StmF), str: method(enc.StrF), encryptMetadata: enc.EncryptMetadata !== false };
}

function key234(h, pw) {
  const { enc, R, n, id0 } = h;
  let k = md5(concat(pad32(pw), enc.O.slice(0, 32), le32(enc.P), id0, R >= 4 && !h.encryptMetadata ? Uint8Array.of(255, 255, 255, 255) : new Uint8Array(0)));
  if (R >= 3) for (let i = 0; i < 50; i++) k = md5(k.slice(0, n));
  return k.slice(0, n);
}
function userCheck234(h, key) {
  if (h.R === 2) return eq(rc4(key, PAD), h.enc.U, 32);
  let x = rc4(key, md5(concat(PAD, h.id0)));
  for (let i = 1; i <= 19; i++) x = rc4(key.map((b) => b ^ i), x);
  return eq(x, h.enc.U, 16);
}
function ownerKey234(h, ownerPw) {
  let k = md5(pad32(ownerPw));
  if (h.R >= 3) for (let i = 0; i < 50; i++) k = md5(k);
  return k.slice(0, h.n);
}
function ownerToUser234(h, ownerPw) {
  const k = ownerKey234(h, ownerPw);
  if (h.R === 2) return rc4(k, h.enc.O.slice(0, 32));
  let x = h.enc.O.slice(0, 32);
  for (let i = 19; i >= 0; i--) x = rc4(k.map((b) => b ^ i), x);
  return x;
}

/** Algorithm 2.B (R6) / plain SHA-256 (R5) */
export function hash2B(pw, salt, udata, rev) {
  let K = sha256(concat(pw, salt, udata));
  if (rev < 6) return K;
  let i = 0; let E;
  for (;;) {
    const one = concat(pw, K, udata);
    const K1 = new Uint8Array(one.length * 64);
    for (let r = 0; r < 64; r++) K1.set(one, r * one.length);
    E = aesEncryptCbc(K.slice(0, 16), K.slice(16, 32), K1, false);
    let sum = 0; for (let j = 0; j < 16; j++) sum += E[j];
    const m = sum % 3;
    K = m === 0 ? sha256(E) : m === 1 ? sha384(E) : sha512(E);
    i++;
    if (i >= 64 && E[E.length - 1] <= i - 32) break;
  }
  return K.slice(0, 32);
}
const ZERO16 = new Uint8Array(16);

/** → { key, owner } or null */
export function authenticate(h, password) {
  if (h.R >= 5) {
    const pw = utf8pw(password); const { O, U, OE, UE } = h.enc;
    if (eq(hash2B(pw, O.slice(32, 40), U.slice(0, 48), h.R), O, 32)) return { key: aesDecryptCbc(hash2B(pw, O.slice(40, 48), U.slice(0, 48), h.R), ZERO16, OE, false), owner: true };
    if (eq(hash2B(pw, U.slice(32, 40), new Uint8Array(0), h.R), U, 32)) return { key: aesDecryptCbc(hash2B(pw, U.slice(40, 48), new Uint8Array(0), h.R), ZERO16, UE, false), owner: false };
    return null;
  }
  const pw = latin1(password);
  // owner first – it unlocks everything
  const asUser = ownerToUser234(h, pw);
  const ko = key234(h, asUser);
  if (userCheck234(h, ko)) return { key: ko, owner: true };
  const ku = key234(h, pw);
  if (userCheck234(h, ku)) return { key: ku, owner: false };
  return null;
}

function objKey(h, key, num, gen, aes) {
  if (h.V === 5) return key;
  const k = md5(concat(key, Uint8Array.of(num & 255, (num >>> 8) & 255, (num >>> 16) & 255, gen & 255, (gen >>> 8) & 255), aes ? SALT : new Uint8Array(0)));
  return k.slice(0, Math.min(h.n + 5, 16));
}
export function decryptData(h, key, num, gen, data, kind) {
  const m = kind === 'string' ? h.str : h.stm;
  if (m === 'none') return data;
  if (m === 'rc4') return rc4(objKey(h, key, num, gen, false), data);
  if (data.length < 16) return new Uint8Array(0);
  const k = m === 'aes256' ? key : objKey(h, key, num, gen, true);
  return aesDecryptCbc(k, data.subarray(0, 16), data.subarray(16));
}
export function encryptData(h, key, num, gen, data, kind) {
  const m = kind === 'string' ? h.str : h.stm;
  if (m === 'none') return data;
  if (m === 'rc4') return rc4(objKey(h, key, num, gen, false), data);
  const k = m === 'aes256' ? key : objKey(h, key, num, gen, true);
  const iv = randomBytes(16);
  return concat(iv, aesEncryptCbc(k, iv, data));
}

/* ---------------- pdf-lib glue ---------------- */

const N = (s) => PDFName.of(s);
const isStr = (o) => o instanceof PDFString || o instanceof PDFHexString;
function bytesOf(o) { return o instanceof PDFString || o instanceof PDFHexString ? o.asBytes() : new Uint8Array(0); }

/** Encrypt dictionary (pdf-lib) → plain values */
function readEncrypt(dict, ctx) {
  const get = (k) => { const v = dict.get(N(k)); return v instanceof PDFRef ? ctx.lookup(v) : v; };
  const num = (k, d) => { const v = get(k); return v instanceof PDFNumber ? v.asNumber() : d; };
  const name = (k) => { const v = get(k); return v instanceof PDFName ? v.decodeText() : undefined; };
  const filter = name('Filter');
  if (filter && filter !== 'Standard') throw Object.assign(new Error(`This PDF uses the “${filter}” protection (certificates) – it can only be opened in the app that made it.`), { code: 'UNSUPPORTED' });
  const CF = {};
  const cf = get('CF');
  if (cf instanceof PDFDict) {
    for (const [k, v] of cf.entries()) {
      const d = v instanceof PDFRef ? ctx.lookup(v) : v;
      if (!(d instanceof PDFDict)) continue;
      const cfm = d.get(N('CFM')); const len = d.get(N('Length'));
      CF[k.decodeText()] = { CFM: cfm instanceof PDFName ? cfm.decodeText() : 'None', Length: len instanceof PDFNumber ? len.asNumber() : undefined };
    }
  }
  const em = get('EncryptMetadata');
  return {
    V: num('V', 0), R: num('R', 2), Length: num('Length', 40), P: num('P', -4) | 0,
    O: bytesOf(get('O')), U: bytesOf(get('U')), OE: bytesOf(get('OE')), UE: bytesOf(get('UE')), Perms: bytesOf(get('Perms')),
    EncryptMetadata: !(em instanceof PDFBool) || em.asBoolean(), StmF: name('StmF'), StrF: name('StrF'), CF,
  };
}

/** walk an object and replace strings/stream data through fn(bytes, kind) */
function transform(obj, fn, skipContents = false) {
  if (isStr(obj)) return PDFHexString.of(toHex(fn(obj.asBytes(), 'string')));
  if (obj instanceof PDFDict) {
    const sig = obj.get(N('ByteRange')) instanceof PDFArray || obj.get(N('Type')) === N('Sig');
    for (const [k, v] of obj.entries()) {
      if (sig && k === N('Contents')) continue;
      if (isStr(v) || v instanceof PDFDict || v instanceof PDFArray) obj.set(k, transform(v, fn));
    }
    return obj;
  }
  if (obj instanceof PDFArray) {
    for (let i = 0; i < obj.size(); i++) { const v = obj.get(i); if (isStr(v) || v instanceof PDFDict || v instanceof PDFArray) obj.set(i, transform(v, fn)); }
    return obj;
  }
  if (obj instanceof PDFRawStream) {
    transform(obj.dict, fn);
    if (!skipContents) obj.contents = fn(obj.contents, 'stream');
    return obj;
  }
  return obj;
}
const typeOf = (o) => (o instanceof PDFRawStream ? o.dict.get(N('Type')) : o instanceof PDFDict ? o.get(N('Type')) : null);

// pdf-lib parses objects in file order and unpacks object streams at once – decrypt each object right after it was
// read, before an object stream is unpacked. Only parsers with __officeDecrypt set are affected.
let patched = false;
function patchParser() {
  if (patched) return;
  patched = true;
  const orig = PDFParser.prototype.parseIndirectObject;
  PDFParser.prototype.parseIndirectObject = async function parseIndirectObjectDecrypting() {
    const dec = this.__officeDecrypt;
    if (!dec) return orig.call(this);
    const ref = this.parseIndirectObjectHeader();
    this.skipWhitespaceAndComments();
    const object = this.parseObject();
    this.skipWhitespaceAndComments();
    this.matchKeyword(ENDOBJ);
    const type = typeOf(object);
    if (object instanceof PDFRawStream && type === N('XRef')) { PDFXRefStreamParser.forStream(object).parseIntoContext(); return ref; }
    dec(object, ref);
    if (object instanceof PDFRawStream && type === N('ObjStm')) await PDFObjectStreamParser.forStream(object, this.shouldWaitForTick).parseIntoContext();
    else this.context.assign(ref, object);
    return ref;
  };
}

async function parse(bytes, decrypt) {
  const parser = PDFParser.forBytesWithOptions(bytes, Infinity, false, false);
  if (decrypt) { patchParser(); parser.__officeDecrypt = decrypt; }
  const context = await parser.parseDocument();
  return context;
}

/** decrypt with the password ('' first – many PDFs only restrict printing/copying) */
export async function openProtected(bytes, password = '') {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const first = await parse(b, null);
  const encRef = first.trailerInfo.Encrypt;
  const encObj = encRef instanceof PDFRef ? first.lookup(encRef) : encRef;
  if (!(encObj instanceof PDFDict)) throw Object.assign(new Error('Not encrypted'), { code: 'PLAIN' });
  const enc = readEncrypt(encObj, first);
  const idArr = first.trailerInfo.ID;
  const ids = idArr instanceof PDFArray ? [0, 1].map((i) => bytesOf(idArr.get(i))) : [new Uint8Array(0), new Uint8Array(0)];
  const h = makeHandler(enc, ids[0]);
  const auth = authenticate(h, password);
  if (!auth) throw Object.assign(new Error('Wrong password'), { code: 'PASSWORD' });
  const { key } = auth;
  const encKey = encRef instanceof PDFRef ? encRef.toString() : null;
  const ctx = await parse(b, (obj, ref) => {
    if (ref.toString() === encKey) return;
    const t = typeOf(obj);
    const skip = t === N('Metadata') && !h.encryptMetadata;
    transform(obj, (data, kind) => decryptData(h, key, ref.objectNumber, ref.generationNumber, data, kind), skip);
  });
  if (encRef instanceof PDFRef) ctx.delete(encRef);
  ctx.trailerInfo.Encrypt = undefined;
  const doc = new PDFDocument(ctx, false, false);
  const plain = await doc.save({ useObjectStreams: true, addDefaultPage: false, updateFieldAppearances: false });
  return {
    plain,
    protection: { handler: h, key, ids, encrypt: enc, owner: auth.owner, permissions: pToPermissions(enc.P), revision: h.R, method: h.stm },
  };
}

/** new AES-256 (R6) protection */
export function createProtection({ userPassword = '', ownerPassword = '', permissions } = {}) {
  const key = randomBytes(32);
  const upw = utf8pw(userPassword); const opw = utf8pw(ownerPassword || userPassword || toHex(randomBytes(16)));
  const uvs = randomBytes(8); const uks = randomBytes(8);
  const U = concat(hash2B(upw, uvs, new Uint8Array(0), 6), uvs, uks);
  const UE = aesEncryptCbc(hash2B(upw, uks, new Uint8Array(0), 6), ZERO16, key, false);
  const ovs = randomBytes(8); const oks = randomBytes(8);
  const O = concat(hash2B(opw, ovs, U, 6), ovs, oks);
  const OE = aesEncryptCbc(hash2B(opw, oks, U, 6), ZERO16, key, false);
  const P = permissionsToP(permissions);
  const perms = concat(le32(P), Uint8Array.of(255, 255, 255, 255, 84, 97, 100, 98), randomBytes(4)); // 'T' 'adb'
  const Perms = aesEncryptEcbBlock(key, perms);
  const enc = { V: 5, R: 6, Length: 256, P, O, U, OE, UE, Perms, EncryptMetadata: true, StmF: 'StdCF', StrF: 'StdCF', CF: { StdCF: { CFM: 'AESV3', Length: 32 } } };
  const id = randomBytes(16);
  return { handler: makeHandler(enc, id), key, ids: [id, id], encrypt: enc, owner: true, permissions: pToPermissions(P), revision: 6, method: 'aes256' };
}

function encryptDict(ctx, enc) {
  const hx = (b) => PDFHexString.of(toHex(b));
  const d = { Filter: 'Standard', V: enc.V, R: enc.R, P: enc.P };
  if (enc.V >= 2) d.Length = enc.V === 5 ? 256 : enc.Length;
  if (enc.V >= 4) {
    d.CF = {}; for (const [k, v] of Object.entries(enc.CF || {})) d.CF[k] = { Type: 'CryptFilter', CFM: v.CFM, AuthEvent: 'DocOpen', Length: v.Length || (v.CFM === 'AESV3' ? 32 : 16) };
    d.StmF = enc.StmF || 'Identity'; d.StrF = enc.StrF || 'Identity';
  }
  const dict = ctx.obj(d);
  dict.set(N('O'), hx(enc.O)); dict.set(N('U'), hx(enc.U));
  if (enc.V === 5) { dict.set(N('OE'), hx(enc.OE)); dict.set(N('UE'), hx(enc.UE)); dict.set(N('Perms'), hx(enc.Perms)); }
  if (enc.V >= 4 && enc.EncryptMetadata === false) dict.set(N('EncryptMetadata'), PDFBool.False);
  return dict;
}

/** plain PDF bytes → encrypted bytes (classic xref table, no object streams – those would have to be encrypted as a whole) */
export async function encryptPdf(plainBytes, spec) {
  const prot = spec.keep || createProtection(spec);
  const doc = await PDFDocument.load(plainBytes, { updateMetadata: false, throwOnInvalidObject: false });
  const ctx = doc.context;
  const h = prot.handler; const { key } = prot;
  for (const [ref, obj] of ctx.enumerateIndirectObjects()) {
    let o = obj;
    // pdf-lib's own streams (PDFContentStream/PDFFlateStream …) → raw streams first
    if (!(o instanceof PDFRawStream) && o && typeof o.getContents === 'function' && o.dict) {
      const raw = PDFRawStream.of(o.dict, o.getContents());
      ctx.assign(ref, raw); o = raw;
    }
    const t = typeOf(o);
    if (t === N('XRef') || t === N('ObjStm')) continue;
    const skip = t === N('Metadata') && !h.encryptMetadata;
    transform(o, (data, kind) => encryptData(h, key, ref.objectNumber, ref.generationNumber, data, kind), skip);
  }
  const encRef = ctx.register(encryptDict(ctx, prot.encrypt));
  ctx.trailerInfo.Encrypt = encRef;
  const hx = (b) => PDFHexString.of(toHex(b));
  ctx.trailerInfo.ID = ctx.obj([hx(prot.ids[0]), hx(prot.ids[1] && prot.ids[1].length ? prot.ids[1] : prot.ids[0])]);
  return PDFWriter.forContext(ctx, Infinity).serializeToBuffer();
}
