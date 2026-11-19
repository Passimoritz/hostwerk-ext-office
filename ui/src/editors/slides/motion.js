/* Is motion wanted? (system "reduce motion" or Hostwerk's own setting) */
import * as UI from '@hostwerk/ui';

export function motionOk() {
  if (typeof UI.motionOk === 'function') { try { return UI.motionOk(); } catch { /* fall through */ } }
  try { return !window.matchMedia('(prefers-reduced-motion: reduce)').matches && !document.querySelector('.app.reduce-motion'); } catch { return true; }
}
