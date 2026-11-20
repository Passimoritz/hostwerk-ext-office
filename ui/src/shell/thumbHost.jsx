/*
 * Offscreen renderer of the thumbnail queue (lib/thumbs.js): mounts ONE chunk Preview at a time into a hidden container
 * (React portal, never visible, no pointer events), waits until it settled and rasterises it. Mounted by Office's dialog
 * host (sidebar.footer) and the editor screen; the last mounted renderer does the work.
 */
import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { ErrorBoundary } from '@hostwerk/kit';
import { registerRenderer, rasterize, settled } from '../lib/thumbs';
import { settingsNow } from './store';

let box = null;
function holder() {
  if (box && document.body.contains(box)) return box;
  box = document.createElement('div');
  box.className = 'of-shell-thumbstage';
  box.setAttribute('aria-hidden', 'true');
  box.setAttribute('inert', '');
  document.body.appendChild(box);
  return box;
}

export default function ThumbRenderer() {
  const [job, setJob] = useState(null);
  const ref = useRef(null);
  const pending = useRef(null);
  useEffect(() => registerRenderer((j) => new Promise((resolve) => {
    pending.current = resolve;
    setJob({ ...j, id: Math.random() });
    // safety net: a preview that never settles
    setTimeout(() => { if (pending.current === resolve) { pending.current = null; setJob(null); resolve(null); } }, 12000);
  })), []);
  useEffect(() => {
    if (!job || !ref.current) return undefined;
    let alive = true;
    (async () => {
      await settled(ref.current, { quiet: 350, max: 6000 });
      let blob = null;
      try { if (alive && ref.current) blob = await rasterize(ref.current); } catch { blob = null; }
      const r = pending.current; pending.current = null;
      setJob(null);
      if (r) r(blob);
    })();
    return () => { alive = false; };
  }, [job && job.id]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!job) return null;
  const P = job.chunk.Preview;
  return createPortal(
    <div ref={ref} className={`of-shell-thumbpage is-${job.chunk.kind || 'x'}`}>
      <ErrorBoundary compact where="office thumbnail"><P bytes={job.bytes} ext={job.ext} name={job.name} settings={settingsNow()} thumbnail /></ErrorBoundary>
    </div>,
    holder(),
  );
}
