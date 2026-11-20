/*
 * Small synchronous crypto primitives for the PDF standard security handler – pure JS, no WebCrypto (crypto.subtle
 * only exists in secure contexts, Hostwerk is often reached over plain http in the LAN). Tested against node:crypto
 * in tests/pdf-crypto.test.js.
 *
 *   md5(bytes) · sha256(bytes) · sha384(bytes) · sha512(bytes)  → Uint8Array
 *   rc4(key, data)                                              → Uint8Array
 *   aesEncryptCbc(key, iv, data, pad = true) · aesDecryptCbc(key, iv, data, pad = true)
 *   aesEncryptEcbBlock(key, block)
 *   concat(...arrays) · randomBytes(n)
 */

export function concat(...parts) {
  let n = 0; for (const p of parts) n += p.length;
  const out = new Uint8Array(n); let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
export function randomBytes(n) {
  const b = new Uint8Array(n);
  if (globalThis.crypto && globalThis.crypto.getRandomValues) globalThis.crypto.getRandomValues(b);
  else for (let i = 0; i < n; i++) b[i] = Math.floor(Math.random() * 256);
  return b;
}

/* ---------------- MD5 ---------------- */
const MD5_S = [7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21];
const MD5_K = new Int32Array(64).map((_, i) => Math.floor(Math.abs(Math.sin(i + 1)) * 4294967296) | 0);
export function md5(data) {
  const len = data.length;
  const n = ((len + 72) >>> 6) << 6;
  const buf = new Uint8Array(n);
  buf.set(data); buf[len] = 0x80;
  const bits = len * 8;
  buf[n - 8] = bits & 255; buf[n - 7] = (bits >>> 8) & 255; buf[n - 6] = (bits >>> 16) & 255; buf[n - 5] = (bits >>> 24) & 255;
  buf[n - 4] = Math.floor(len / 536870912) & 255;
  let a0 = 0x67452301; let b0 = 0xefcdab89 | 0; let c0 = 0x98badcfe | 0; let d0 = 0x10325476;
  const M = new Int32Array(16);
  for (let off = 0; off < n; off += 64) {
    for (let j = 0; j < 16; j++) { const o = off + j * 4; M[j] = buf[o] | (buf[o + 1] << 8) | (buf[o + 2] << 16) | (buf[o + 3] << 24); }
    let A = a0; let B = b0; let C = c0; let D = d0;
    for (let i = 0; i < 64; i++) {
      let F; let g;
      if (i < 16) { F = (B & C) | (~B & D); g = i; } else if (i < 32) { F = (D & B) | (~D & C); g = (5 * i + 1) & 15; } else if (i < 48) { F = B ^ C ^ D; g = (3 * i + 5) & 15; } else { F = C ^ (B | ~D); g = (7 * i) & 15; }
      const tmp = D; D = C; C = B;
      const x = (A + F + MD5_K[i] + M[g]) | 0;
      B = (B + ((x << MD5_S[i]) | (x >>> (32 - MD5_S[i])))) | 0;
      A = tmp;
    }
    a0 = (a0 + A) | 0; b0 = (b0 + B) | 0; c0 = (c0 + C) | 0; d0 = (d0 + D) | 0;
  }
  const out = new Uint8Array(16);
  [a0, b0, c0, d0].forEach((v, i) => { out[i * 4] = v & 255; out[i * 4 + 1] = (v >>> 8) & 255; out[i * 4 + 2] = (v >>> 16) & 255; out[i * 4 + 3] = (v >>> 24) & 255; });
  return out;
}

/* ---------------- RC4 ---------------- */
export function rc4(key, data) {
  const s = new Uint8Array(256);
  for (let i = 0; i < 256; i++) s[i] = i;
  for (let i = 0, j = 0; i < 256; i++) { j = (j + s[i] + key[i % key.length]) & 255; const t = s[i]; s[i] = s[j]; s[j] = t; }
  const out = new Uint8Array(data.length);
  for (let k = 0, i = 0, j = 0; k < data.length; k++) {
    i = (i + 1) & 255; j = (j + s[i]) & 255;
    const t = s[i]; s[i] = s[j]; s[j] = t;
    out[k] = data[k] ^ s[(s[i] + s[j]) & 255];
  }
  return out;
}

/* ---------------- SHA-256 ---------------- */
const K256 = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);
function padBE(data, block, lenBytes) {
  const len = data.length;
  const n = Math.ceil((len + 1 + lenBytes) / block) * block;
  const buf = new Uint8Array(n);
  buf.set(data); buf[len] = 0x80;
  const bits = len * 8;
  buf[n - 1] = bits & 255; buf[n - 2] = (bits >>> 8) & 255; buf[n - 3] = (bits >>> 16) & 255; buf[n - 4] = (bits >>> 24) & 255;
  buf[n - 5] = Math.floor(len / 536870912) & 255;
  return buf;
}
export function sha256(data) {
  const buf = padBE(data, 64, 8);
  const H = new Uint32Array([0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19]);
  const W = new Uint32Array(64);
  for (let off = 0; off < buf.length; off += 64) {
    for (let i = 0; i < 16; i++) { const o = off + i * 4; W[i] = (buf[o] << 24) | (buf[o + 1] << 16) | (buf[o + 2] << 8) | buf[o + 3]; }
    for (let i = 16; i < 64; i++) {
      const x = W[i - 15]; const y = W[i - 2];
      const s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3);
      const s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
      W[i] = (W[i - 16] + s0 + W[i - 7] + s1) | 0;
    }
    let [a, b, c, d, e, f, g, h] = H;
    for (let i = 0; i < 64; i++) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const ch = (e & f) ^ (~e & g);
      const t1 = (h + S1 + ch + K256[i] + W[i]) | 0;
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const mj = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (S0 + mj) | 0;
      h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
    }
    H[0] += a; H[1] += b; H[2] += c; H[3] += d; H[4] += e; H[5] += f; H[6] += g; H[7] += h;
  }
  const out = new Uint8Array(32);
  for (let i = 0; i < 8; i++) { out[i * 4] = H[i] >>> 24; out[i * 4 + 1] = (H[i] >>> 16) & 255; out[i * 4 + 2] = (H[i] >>> 8) & 255; out[i * 4 + 3] = H[i] & 255; }
  return out;
}

/* ---------------- SHA-512 / SHA-384 (64-bit words as hi/lo pairs) ---------------- */
const K512 = [
  '428a2f98d728ae22', '7137449123ef65cd', 'b5c0fbcfec4d3b2f', 'e9b5dba58189dbbc', '3956c25bf348b538', '59f111f1b605d019', '923f82a4af194f9b', 'ab1c5ed5da6d8118',
  'd807aa98a3030242', '12835b0145706fbe', '243185be4ee4b28c', '550c7dc3d5ffb4e2', '72be5d74f27b896f', '80deb1fe3b1696b1', '9bdc06a725c71235', 'c19bf174cf692694',
  'e49b69c19ef14ad2', 'efbe4786384f25e3', '0fc19dc68b8cd5b5', '240ca1cc77ac9c65', '2de92c6f592b0275', '4a7484aa6ea6e483', '5cb0a9dcbd41fbd4', '76f988da831153b5',
  '983e5152ee66dfab', 'a831c66d2db43210', 'b00327c898fb213f', 'bf597fc7beef0ee4', 'c6e00bf33da88fc2', 'd5a79147930aa725', '06ca6351e003826f', '142929670a0e6e70',
  '27b70a8546d22ffc', '2e1b21385c26c926', '4d2c6dfc5ac42aed', '53380d139d95b3df', '650a73548baf63de', '766a0abb3c77b2a8', '81c2c92e47edaee6', '92722c851482353b',
  'a2bfe8a14cf10364', 'a81a664bbc423001', 'c24b8b70d0f89791', 'c76c51a30654be30', 'd192e819d6ef5218', 'd69906245565a910', 'f40e35855771202a', '106aa07032bbd1b8',
  '19a4c116b8d2d0c8', '1e376c085141ab53', '2748774cdf8eeb99', '34b0bcb5e19b48a8', '391c0cb3c5c95a63', '4ed8aa4ae3418acb', '5b9cca4f7763e373', '682e6ff3d6b2b8a3',
  '748f82ee5defb2fc', '78a5636f43172f60', '84c87814a1f0ab72', '8cc702081a6439ec', '90befffa23631e28', 'a4506cebde82bde9', 'bef9a3f7b2c67915', 'c67178f2e372532b',
  'ca273eceea26619c', 'd186b8c721c0c207', 'eada7dd6cde0eb1e', 'f57d4f7fee6ed178', '06f067aa72176fba', '0a637dc5a2c898a6', '113f9804bef90dae', '1b710b35131c471b',
  '28db77f523047d84', '32caab7b40c72493', '3c9ebe0a15c9bebc', '431d67c49c100d4c', '4cc5d4becb3e42b6', '597f299cfc657e2a', '5fcb6fab3ad6faec', '6c44198c4a475817',
];
const K512H = new Uint32Array(80); const K512L = new Uint32Array(80);
K512.forEach((s, i) => { K512H[i] = parseInt(s.slice(0, 8), 16); K512L[i] = parseInt(s.slice(8), 16); });
const IV512 = ['6a09e667f3bcc908', 'bb67ae8584caa73b', '3c6ef372fe94f82b', 'a54ff53a5f1d36f1', '510e527fade682d1', '9b05688c2b3e6c1f', '1f83d9abfb41bd6b', '5be0cd19137e2179'];
const IV384 = ['cbbb9d5dc1059ed8', '629a292a367cd507', '9159015a3070dd17', '152fecd8f70e5939', '67332667ffc00b31', '8eb44a8768581511', 'db0c2e0d64f98fa7', '47b5481dbefa4fa4'];

function sha512core(data, iv, outLen) {
  const buf = padBE(data, 128, 16);
  const hh = new Uint32Array(8); const hl = new Uint32Array(8);
  iv.forEach((s, i) => { hh[i] = parseInt(s.slice(0, 8), 16); hl[i] = parseInt(s.slice(8), 16); });
  const WH = new Uint32Array(80); const WL = new Uint32Array(80);
  // 64-bit add of several (hi, lo) pairs
  const add = (pairs) => { let lo = 0; let hi = 0; for (let k = 0; k < pairs.length; k += 2) { lo += pairs[k + 1]; hi += pairs[k] + Math.floor(lo / 4294967296); lo >>>= 0; } return [hi >>> 0, lo >>> 0]; };
  for (let off = 0; off < buf.length; off += 128) {
    for (let i = 0; i < 16; i++) {
      const o = off + i * 8;
      WH[i] = ((buf[o] << 24) | (buf[o + 1] << 16) | (buf[o + 2] << 8) | buf[o + 3]) >>> 0;
      WL[i] = ((buf[o + 4] << 24) | (buf[o + 5] << 16) | (buf[o + 6] << 8) | buf[o + 7]) >>> 0;
    }
    for (let i = 16; i < 80; i++) {
      // s0 = rotr1 ^ rotr8 ^ shr7 of W[i-15]; s1 = rotr19 ^ rotr61 ^ shr6 of W[i-2]
      let xh = WH[i - 15]; let xl = WL[i - 15];
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      xh = WH[i - 2]; xl = WL[i - 2];
      const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
      const s1l = ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26));
      const [h, l] = add([WH[i - 16], WL[i - 16], s0h >>> 0, s0l >>> 0, WH[i - 7], WL[i - 7], s1h >>> 0, s1l >>> 0]);
      WH[i] = h; WL[i] = l;
    }
    let ah = hh[0]; let al = hl[0]; let bh = hh[1]; let bl = hl[1]; let ch = hh[2]; let cl = hl[2]; let dh = hh[3]; let dl = hl[3];
    let eh = hh[4]; let el = hl[4]; let fh = hh[5]; let fl = hl[5]; let gh = hh[6]; let gl = hl[6]; let hh7 = hh[7]; let hl7 = hl[7];
    for (let i = 0; i < 80; i++) {
      // S1 = rotr14 ^ rotr18 ^ rotr41 of e
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh); const chl = (el & fl) ^ (~el & gl);
      const [t1h, t1l] = add([hh7, hl7, S1h >>> 0, S1l >>> 0, chh >>> 0, chl >>> 0, K512H[i], K512L[i], WH[i], WL[i]]);
      // S0 = rotr28 ^ rotr34 ^ rotr39 of a
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch); const mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
      const [t2h, t2l] = add([S0h >>> 0, S0l >>> 0, mjh >>> 0, mjl >>> 0]);
      hh7 = gh; hl7 = gl; gh = fh; gl = fl; fh = eh; fl = el;
      [eh, el] = add([dh, dl, t1h, t1l]);
      dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
      [ah, al] = add([t1h, t1l, t2h, t2l]);
    }
    const vals = [[ah, al], [bh, bl], [ch, cl], [dh, dl], [eh, el], [fh, fl], [gh, gl], [hh7, hl7]];
    for (let i = 0; i < 8; i++) { const [h, l] = add([hh[i], hl[i], vals[i][0], vals[i][1]]); hh[i] = h; hl[i] = l; }
  }
  const out = new Uint8Array(64);
  for (let i = 0; i < 8; i++) {
    out[i * 8] = hh[i] >>> 24; out[i * 8 + 1] = (hh[i] >>> 16) & 255; out[i * 8 + 2] = (hh[i] >>> 8) & 255; out[i * 8 + 3] = hh[i] & 255;
    out[i * 8 + 4] = hl[i] >>> 24; out[i * 8 + 5] = (hl[i] >>> 16) & 255; out[i * 8 + 6] = (hl[i] >>> 8) & 255; out[i * 8 + 7] = hl[i] & 255;
  }
  return out.slice(0, outLen);
}
export const sha512 = (d) => sha512core(d, IV512, 64);
export const sha384 = (d) => sha512core(d, IV384, 48);

/* ---------------- AES (T-tables) ---------------- */
const SBOX = new Uint8Array(256); const INV = new Uint8Array(256);
const TE = [new Uint32Array(256), new Uint32Array(256), new Uint32Array(256), new Uint32Array(256)];
const TD = [new Uint32Array(256), new Uint32Array(256), new Uint32Array(256), new Uint32Array(256)];
(() => {
  const xt = (a) => ((a << 1) ^ (a & 0x80 ? 0x1b : 0)) & 255;
  const mul = (a, b) => { let r = 0; while (b) { if (b & 1) r ^= a; a = xt(a); b >>= 1; } return r; };
  // S-box from the multiplicative inverse + affine transform
  for (let x = 0; x < 256; x++) {
    let inv = 0;
    if (x) for (let y = 1; y < 256; y++) if (mul(x, y) === 1) { inv = y; break; }
    let s = inv; let r = inv;
    for (let k = 0; k < 4; k++) { r = ((r << 1) | (r >>> 7)) & 255; s ^= r; }
    s ^= 0x63;
    SBOX[x] = s; INV[s] = x;
  }
  const rot = (w, n) => ((w >>> n) | (w << (32 - n))) >>> 0;
  for (let x = 0; x < 256; x++) {
    const s = SBOX[x];
    const e = ((mul(s, 2) << 24) | (s << 16) | (s << 8) | mul(s, 3)) >>> 0;
    const i = INV[x];
    const d = ((mul(i, 14) << 24) | (mul(i, 9) << 16) | (mul(i, 13) << 8) | mul(i, 11)) >>> 0;
    TE[0][x] = e; TE[1][x] = rot(e, 8); TE[2][x] = rot(e, 16); TE[3][x] = rot(e, 24);
    TD[0][x] = d; TD[1][x] = rot(d, 8); TD[2][x] = rot(d, 16); TD[3][x] = rot(d, 24);
  }
})();

const keyCache = new Map();
function expand(key) {
  const id = Array.from(key).join(',');
  if (keyCache.has(id)) return keyCache.get(id);
  const nk = key.length / 4;
  if (![4, 6, 8].includes(nk)) throw new Error('AES key must be 16, 24 or 32 bytes');
  const nr = nk + 6;
  const w = new Uint32Array(4 * (nr + 1));
  for (let i = 0; i < nk; i++) w[i] = ((key[4 * i] << 24) | (key[4 * i + 1] << 16) | (key[4 * i + 2] << 8) | key[4 * i + 3]) >>> 0;
  let rcon = 1;
  for (let i = nk; i < w.length; i++) {
    let t = w[i - 1];
    if (i % nk === 0) {
      t = ((t << 8) | (t >>> 24)) >>> 0;
      t = ((SBOX[t >>> 24] << 24) | (SBOX[(t >>> 16) & 255] << 16) | (SBOX[(t >>> 8) & 255] << 8) | SBOX[t & 255]) >>> 0;
      t = (t ^ (rcon << 24)) >>> 0;
      rcon = ((rcon << 1) ^ (rcon & 0x80 ? 0x1b : 0)) & 255;
    } else if (nk > 6 && i % nk === 4) {
      t = ((SBOX[t >>> 24] << 24) | (SBOX[(t >>> 16) & 255] << 16) | (SBOX[(t >>> 8) & 255] << 8) | SBOX[t & 255]) >>> 0;
    }
    w[i] = (w[i - nk] ^ t) >>> 0;
  }
  // decryption keys (equivalent inverse cipher): reversed rounds, InvMixColumns on the middle ones
  const dw = new Uint32Array(w.length);
  for (let r = 0; r <= nr; r++) for (let c = 0; c < 4; c++) {
    const v = w[(nr - r) * 4 + c];
    dw[r * 4 + c] = r === 0 || r === nr ? v : (TD[0][SBOX[v >>> 24]] ^ TD[1][SBOX[(v >>> 16) & 255]] ^ TD[2][SBOX[(v >>> 8) & 255]] ^ TD[3][SBOX[v & 255]]) >>> 0;
  }
  const k = { w, dw, nr };
  if (keyCache.size > 64) keyCache.clear();
  keyCache.set(id, k);
  return k;
}

function encBlock(k, inp, io, out, oo) {
  const { w, nr } = k; const [T0, T1, T2, T3] = TE;
  let s0 = ((inp[io] << 24) | (inp[io + 1] << 16) | (inp[io + 2] << 8) | inp[io + 3]) ^ w[0];
  let s1 = ((inp[io + 4] << 24) | (inp[io + 5] << 16) | (inp[io + 6] << 8) | inp[io + 7]) ^ w[1];
  let s2 = ((inp[io + 8] << 24) | (inp[io + 9] << 16) | (inp[io + 10] << 8) | inp[io + 11]) ^ w[2];
  let s3 = ((inp[io + 12] << 24) | (inp[io + 13] << 16) | (inp[io + 14] << 8) | inp[io + 15]) ^ w[3];
  for (let r = 1; r < nr; r++) {
    const t0 = T0[s0 >>> 24] ^ T1[(s1 >>> 16) & 255] ^ T2[(s2 >>> 8) & 255] ^ T3[s3 & 255] ^ w[4 * r];
    const t1 = T0[s1 >>> 24] ^ T1[(s2 >>> 16) & 255] ^ T2[(s3 >>> 8) & 255] ^ T3[s0 & 255] ^ w[4 * r + 1];
    const t2 = T0[s2 >>> 24] ^ T1[(s3 >>> 16) & 255] ^ T2[(s0 >>> 8) & 255] ^ T3[s1 & 255] ^ w[4 * r + 2];
    const t3 = T0[s3 >>> 24] ^ T1[(s0 >>> 16) & 255] ^ T2[(s1 >>> 8) & 255] ^ T3[s2 & 255] ^ w[4 * r + 3];
    s0 = t0; s1 = t1; s2 = t2; s3 = t3;
  }
  const S = SBOX; const b = 4 * nr;
  const o = [
    (S[s0 >>> 24] << 24) | (S[(s1 >>> 16) & 255] << 16) | (S[(s2 >>> 8) & 255] << 8) | S[s3 & 255],
    (S[s1 >>> 24] << 24) | (S[(s2 >>> 16) & 255] << 16) | (S[(s3 >>> 8) & 255] << 8) | S[s0 & 255],
    (S[s2 >>> 24] << 24) | (S[(s3 >>> 16) & 255] << 16) | (S[(s0 >>> 8) & 255] << 8) | S[s1 & 255],
    (S[s3 >>> 24] << 24) | (S[(s0 >>> 16) & 255] << 16) | (S[(s1 >>> 8) & 255] << 8) | S[s2 & 255],
  ];
  for (let c = 0; c < 4; c++) { const v = (o[c] ^ w[b + c]) >>> 0; out[oo + 4 * c] = v >>> 24; out[oo + 4 * c + 1] = (v >>> 16) & 255; out[oo + 4 * c + 2] = (v >>> 8) & 255; out[oo + 4 * c + 3] = v & 255; }
}
function decBlock(k, inp, io, out, oo) {
  const { dw: w, nr } = k; const [T0, T1, T2, T3] = TD;
  let s0 = ((inp[io] << 24) | (inp[io + 1] << 16) | (inp[io + 2] << 8) | inp[io + 3]) ^ w[0];
  let s1 = ((inp[io + 4] << 24) | (inp[io + 5] << 16) | (inp[io + 6] << 8) | inp[io + 7]) ^ w[1];
  let s2 = ((inp[io + 8] << 24) | (inp[io + 9] << 16) | (inp[io + 10] << 8) | inp[io + 11]) ^ w[2];
  let s3 = ((inp[io + 12] << 24) | (inp[io + 13] << 16) | (inp[io + 14] << 8) | inp[io + 15]) ^ w[3];
  for (let r = 1; r < nr; r++) {
    const t0 = T0[s0 >>> 24] ^ T1[(s3 >>> 16) & 255] ^ T2[(s2 >>> 8) & 255] ^ T3[s1 & 255] ^ w[4 * r];
    const t1 = T0[s1 >>> 24] ^ T1[(s0 >>> 16) & 255] ^ T2[(s3 >>> 8) & 255] ^ T3[s2 & 255] ^ w[4 * r + 1];
    const t2 = T0[s2 >>> 24] ^ T1[(s1 >>> 16) & 255] ^ T2[(s0 >>> 8) & 255] ^ T3[s3 & 255] ^ w[4 * r + 2];
    const t3 = T0[s3 >>> 24] ^ T1[(s2 >>> 16) & 255] ^ T2[(s1 >>> 8) & 255] ^ T3[s0 & 255] ^ w[4 * r + 3];
    s0 = t0; s1 = t1; s2 = t2; s3 = t3;
  }
  const S = INV; const b = 4 * nr;
  const o = [
    (S[s0 >>> 24] << 24) | (S[(s3 >>> 16) & 255] << 16) | (S[(s2 >>> 8) & 255] << 8) | S[s1 & 255],
    (S[s1 >>> 24] << 24) | (S[(s0 >>> 16) & 255] << 16) | (S[(s3 >>> 8) & 255] << 8) | S[s2 & 255],
    (S[s2 >>> 24] << 24) | (S[(s1 >>> 16) & 255] << 16) | (S[(s0 >>> 8) & 255] << 8) | S[s3 & 255],
    (S[s3 >>> 24] << 24) | (S[(s2 >>> 16) & 255] << 16) | (S[(s1 >>> 8) & 255] << 8) | S[s0 & 255],
  ];
  for (let c = 0; c < 4; c++) { const v = (o[c] ^ w[b + c]) >>> 0; out[oo + 4 * c] = v >>> 24; out[oo + 4 * c + 1] = (v >>> 16) & 255; out[oo + 4 * c + 2] = (v >>> 8) & 255; out[oo + 4 * c + 3] = v & 255; }
}

export function aesEncryptEcbBlock(key, block) { const out = new Uint8Array(16); encBlock(expand(key), block, 0, out, 0); return out; }

/** CBC encryption; pad = PKCS#7 (PDF streams/strings), false = data must be a multiple of 16 */
export function aesEncryptCbc(key, iv, data, pad = true) {
  const k = expand(key);
  let src = data;
  if (pad) { const p = 16 - (data.length % 16); src = new Uint8Array(data.length + p); src.set(data); src.fill(p, data.length); }
  if (src.length % 16) throw new Error('AES data is not a multiple of 16 bytes');
  const out = new Uint8Array(src.length);
  const prev = Uint8Array.from(iv); const blk = new Uint8Array(16);
  for (let o = 0; o < src.length; o += 16) {
    for (let i = 0; i < 16; i++) blk[i] = src[o + i] ^ prev[i];
    encBlock(k, blk, 0, out, o);
    prev.set(out.subarray(o, o + 16));
  }
  return out;
}
export function aesDecryptCbc(key, iv, data, pad = true) {
  const k = expand(key);
  const n = data.length - (data.length % 16);
  const out = new Uint8Array(n);
  let prev = iv;
  for (let o = 0; o < n; o += 16) {
    decBlock(k, data, o, out, o);
    for (let i = 0; i < 16; i++) out[o + i] ^= prev[i];
    prev = data.subarray(o, o + 16);
  }
  if (pad && n) { const p = out[n - 1]; if (p >= 1 && p <= 16) return out.subarray(0, n - p); }
  return out;
}
