/*
 * Hostwerk's toast as extensions get it. The translating wrapper Hostwerk hands to native extensions
 * (ui/src/extNative.js forExt) only carries toast(), .success and .error – .warn, .dismiss and .done of the core
 * toast are missing (Hostwerk 3.2.1). Calling them threw "toast.dismiss is not a function" after a successful
 * conversion. This module fills the gaps and uses the real ones as soon as Hostwerk provides them.
 */
import { toast as base } from '@hostwerk/ui';

const has = (k) => typeof base[k] === 'function';
export const toast = (...args) => base(...args);
toast.success = (m) => base.success(m);
toast.error = (e) => base.error(e);
toast.warn = (m) => (has('warn') ? base.warn(m) : base(m, 'warn', 6000));
toast.done = (m, o) => (has('done') ? base.done(m, o) : base.success(m));
/** true when a toast can be closed early – otherwise "progress" toasts should get a time limit instead of 0 */
toast.canDismiss = has('dismiss');
toast.dismiss = (id) => { if (has('dismiss') && id != null) base.dismiss(id); };
/** a "working on it …" toast: stays until dismissed where possible, else disappears on its own after `fallbackMs` */
toast.progress = (m, fallbackMs = 30000) => base(m, 'info', has('dismiss') ? 0 : fallbackMs);
