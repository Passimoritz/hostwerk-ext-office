/* Shared helpers of the slides readers/writers (browser + Node). */

export const MIME = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml',
  tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', ico: 'image/x-icon',
  mp4: 'video/mp4', m4v: 'video/mp4', webm: 'video/webm', mov: 'video/quicktime', avi: 'video/x-msvideo', wmv: 'video/x-ms-wmv', ogv: 'video/ogg',
  mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', ogg: 'audio/ogg', oga: 'audio/ogg', wma: 'audio/x-ms-wma', aac: 'audio/aac',
};
export const EXT_OF_MIME = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/tiff': 'tiff', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf', 'video/mp4': 'mp4', 'video/webm': 'webm', 'video/quicktime': 'mov', 'audio/mpeg': 'mp3', 'audio/mp4': 'm4a', 'audio/wav': 'wav', 'audio/ogg': 'ogg', 'audio/aac': 'aac' };
export const extOfPath = (p) => { const m = /\.([a-z0-9]+)$/i.exec(p || ''); return m ? m[1].toLowerCase() : ''; };
export const mimeOfPath = (p) => MIME[extOfPath(p)] || 'application/octet-stream';

export function bytesToBase64(bytes) {
  let s = '';
  const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
  return btoa(s);
}
export function base64ToBytes(b64) {
  const s = atob(b64);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
export const bytesToDataUrl = (bytes, mime) => `data:${mime};base64,${bytesToBase64(bytes)}`;
/** { mime, bytes } of a data URL (null if not one) */
export function dataUrlToBytes(url) {
  const m = /^data:([^;,]*)(;base64)?,(.*)$/s.exec(url || '');
  if (!m) return null;
  const mime = m[1] || 'application/octet-stream';
  if (m[2]) return { mime, bytes: base64ToBytes(m[3]) };
  return { mime, bytes: new TextEncoder().encode(decodeURIComponent(m[3])) };
}
export const mimeOfDataUrl = (url) => { const m = /^data:([^;,]*)/.exec(url || ''); return m ? m[1] : ''; };

/** small stable hash for de-duplicating media */
export function hashStr(s) {
  let h = 2166136261 >>> 0;
  const step = Math.max(1, Math.floor(s.length / 4096));
  for (let i = 0; i < s.length; i += step) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
  return `${h.toString(36)}${s.length.toString(36)}`;
}

/**
 * Fidelity keys the readers report ({ doc, fidelity: [key] }); the editor shows the labels in its Callout.
 */
export const FIDELITY = {
  macros: 'Macros',
  smartart: 'SmartArt graphics',
  ole: 'Embedded objects',
  charts: 'Charts (simplified)',
  animations: 'Complex animations',
  transitions: 'Some transitions',
  math: 'Equations',
  model3d: '3D models',
  ink: 'Ink drawings',
  geometry: 'Custom shapes (approximated)',
  effects: 'Some effects (3D, glow, reflection)',
  media: 'Linked media',
  comments: 'Comments',
  fields: 'Some fields',
  signature: 'Digital signature (removed when saving)',
};

export const isZip = (bytes) => bytes && bytes.length > 3 && bytes[0] === 0x50 && bytes[1] === 0x4b;
