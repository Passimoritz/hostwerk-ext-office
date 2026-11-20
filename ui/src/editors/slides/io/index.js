/* One entry point for reading/writing every presentation format the chunk handles. */
import { readPptx } from './pptx-read.js';
import { writePptx } from './pptx-write.js';
import { readOdp } from './odp-read.js';
import { writeOdp } from './odp-write.js';
import { readPpt } from './ppt-read.js';
import { isZip } from './util.js';
import { isCfb } from '../../_shared/cfb.js';
import { toFlat } from './fodp.js';

export const OOXML = ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'];
export const ODF = ['odp', 'otp', 'fodp'];
export const LEGACY = ['ppt', 'pps', 'pot'];
export const MACRO = ['pptm', 'ppsm', 'potm'];

/** detect the real format from the bytes (a .ppt can be a renamed pptx and vice versa) */
export function sniff(bytes, ext) {
  const e = String(ext || '').toLowerCase();
  if (isZip(bytes)) {
    const head = new TextDecoder('latin1').decode(bytes.subarray(0, Math.min(bytes.length, 400)));
    if (/mimetypeapplication\/vnd\.oasis\.opendocument/.test(head)) return ODF.includes(e) ? e : 'odp';
    return OOXML.includes(e) ? e : 'pptx';
  }
  if (isCfb(bytes)) return LEGACY.includes(e) ? e : 'ppt';
  const s = new TextDecoder('utf-8').decode(bytes.subarray(0, Math.min(bytes.length, 2000)));
  if (/<office:document/.test(s)) return 'fodp';
  return e;
}

export async function readAny(bytes, ext) {
  const fmt = sniff(bytes, ext);
  let r;
  if (ODF.includes(fmt)) r = await readOdp(bytes, { ext: fmt });
  else if (LEGACY.includes(fmt)) r = await readPpt(bytes, { ext: fmt });
  else r = await readPptx(bytes, { ext: fmt });
  return { doc: r.doc, fidelity: r.fidelity || [], preserved: r.preserved || [], fmt };
}

/** native format a file is saved in: import-only formats write the nearest native one (macros keep pptm) */
export function saveFormat(ext) {
  const e = String(ext || '').toLowerCase();
  if (e === 'pptx' || e === 'odp' || e === 'fodp') return e;
  if (MACRO.includes(e)) return 'pptm';
  if (ODF.includes(e)) return 'odp';
  return 'pptx';
}

/**
 * write: fmt = 'pptx' | 'pptm' | 'ppsx' | 'potx' | 'odp' | 'otp'
 * original = bytes of the opened file (patched losslessly when the format family matches)
 */
export async function writeAny(doc, fmt, { original, originalFmt, chartImage } = {}) {
  const f = String(fmt || 'pptx').toLowerCase();
  if (f === 'fodp') return toFlat(await writeOdp(doc, { ext: 'odp', chartImage }));
  if (ODF.includes(f)) return writeOdp(doc, { ext: f, original: original && ['odp', 'otp'].includes(originalFmt) ? original : undefined, chartImage });
  return writePptx(doc, { ext: f, original: original && OOXML.includes(originalFmt) ? original : undefined });
}
