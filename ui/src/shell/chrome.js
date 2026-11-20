/*
 * Window chrome for Office's full views (screen #/office-edit – Hostwerk's "app-bare" layout without its title bar).
 * Hostwerk's desktop window is frameless (electron/main.js: titleBarStyle 'hidden' + titleBarOverlay 40 px – Windows
 * draws minimise/maximise/close at the top right over the page); its own title bar is the drag region
 * (-webkit-app-region: drag, buttons no-drag, 150 px kept free on the right). A screen hides that bar, so Office's
 * top bar takes the job: it is the drag region (double-click maximises – Windows does that for drag regions), keeps the
 * window controls free (navigator.windowControlsOverlay / env(titlebar-area-*), else 150 px like Hostwerk) and offers
 * full screen. Windows of their own (#…&window=1) have a normal frame (Electron's default for child windows) – nothing to do.
 */
import { useEffect, useState } from 'react';

/** Hostwerk's desktop app (preload exposes window.hostwerk) */
export const isDesktopApp = () => typeof window !== 'undefined' && !!window.hostwerk && typeof window.hostwerk.setTheme === 'function';
/** a window of its own (#/…?window=1) – framed by the system */
export const isOwnWindow = (hash = typeof location !== 'undefined' ? location.hash : '') => /[?&]window=1(&|$)/.test(String(hash));
const zoomOf = () => { try { return parseFloat(document.documentElement.style.zoom) || 1; } catch { return 1; } };

/**
 * Space the native window controls cover → { on, right, height } in CSS px of the page (Hostwerk's interface zoom
 * taken out). on = false in the browser, in own windows and while in full screen.
 */
export function overlayGeometry({ wco = typeof navigator !== 'undefined' ? navigator.windowControlsOverlay : null, innerWidth = typeof window !== 'undefined' ? window.innerWidth : 0, zoom = zoomOf(), desktop = isDesktopApp(), own = isOwnWindow(), full = typeof document !== 'undefined' && !!document.fullscreenElement } = {}) {
  if (!desktop || own || full) return { on: false, right: 0, height: 0 };
  if (wco && wco.visible && typeof wco.getTitlebarAreaRect === 'function') {
    const r = wco.getTitlebarAreaRect();
    const right = Math.max(0, innerWidth - (r.x + r.width));
    return { on: true, right: Math.round(right / zoom), height: Math.round((r.height || 40) / zoom) };
  }
  if (wco && wco.visible === false) return { on: false, right: 0, height: 0 };
  // no WCO API: Hostwerk's own title bar keeps 150 px free for the controls
  return { on: true, right: Math.round(150 / zoom), height: Math.round(40 / zoom) };
}

/** live overlay geometry + full screen state for a top bar */
export function useWindowChrome() {
  const [g, setG] = useState(() => overlayGeometry());
  const [full, setFull] = useState(() => typeof document !== 'undefined' && !!document.fullscreenElement);
  useEffect(() => {
    const upd = () => { setG(overlayGeometry()); setFull(!!document.fullscreenElement); };
    const wco = navigator.windowControlsOverlay;
    if (wco && wco.addEventListener) wco.addEventListener('geometrychange', upd);
    window.addEventListener('resize', upd);
    document.addEventListener('fullscreenchange', upd);
    return () => { if (wco && wco.removeEventListener) wco.removeEventListener('geometrychange', upd); window.removeEventListener('resize', upd); document.removeEventListener('fullscreenchange', upd); };
  }, []);
  return { ...g, full, canFull: typeof document !== 'undefined' && !!document.fullscreenEnabled };
}

/** full screen of the whole window on/off (Esc or the same button leaves it) */
export async function toggleFullscreen() {
  try {
    if (document.fullscreenElement) await document.exitFullscreen();
    else await document.documentElement.requestFullscreen({ navigationUI: 'hide' });
  } catch { /* not allowed here */ }
}
