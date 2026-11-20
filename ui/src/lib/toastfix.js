/*
 * Hostwerk's translating toast for extensions (core ui/src/extNative.js forExt) copies an error with
 * Object.assign(Object.create(proto), e, { message }) – for errors whose `message` is a getter only on the prototype
 * (DOMException: AbortError, NotAllowedError, QuotaExceededError … and similar host objects) that assignment throws
 * "Cannot set property message of #<DOMException> which has only a getter", so the toast never showed and the error
 * surfaced as an uncaught exception. plainError() turns such errors into an ordinary Error (same message, name, status,
 * code) before the core sees them. Installed once (main module first; the chunks share the same toast object).
 */
import { toast } from '@hostwerk/ui';

function messageIsGetterOnly(e) {
  if (Object.prototype.hasOwnProperty.call(e, 'message')) return false;
  for (let p = Object.getPrototypeOf(e); p; p = Object.getPrototypeOf(p)) {
    const d = Object.getOwnPropertyDescriptor(p, 'message');
    if (d) return !!(d.get && !d.set);
  }
  return false;
}
export function plainError(e) {
  if (!e || typeof e !== 'object' || e.message == null) return e;
  if (!messageIsGetterOnly(e) && !Object.isFrozen(e)) return e;
  const out = new Error(String(e.message));
  for (const k of ['name', 'status', 'code', 'hint', 'detail', 'extra']) if (e[k] !== undefined) { try { out[k] = e[k]; } catch { /* read-only */ } }
  return out;
}
if (toast && typeof toast.error === 'function' && !toast.error.ofPlain) {
  const orig = toast.error;
  const wrapped = (e, ...rest) => orig(plainError(e), ...rest);
  wrapped.ofPlain = true;
  try { toast.error = wrapped; } catch { /* frozen toast – nothing to do */ }
}
export default plainError;
