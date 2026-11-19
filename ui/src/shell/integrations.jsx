/*
 * The full editor (ONLYOFFICE Document Server) embedded INSIDE the Hostwerk page (never a browser tab). Hostwerk's
 * window CSP allows no external scripts, so the backend serves a small host page with its own CSP:
 *   GET /api/office/integrations/frame?[root=…&]path=…&app=onlyoffice&lang=de|en
 * It creates the session (ONLYOFFICE config + JWT) and loads ONLYOFFICE's api.js – all inside this iframe. Optional
 * messages from the frame to the shell (window.parent.postMessage):
 *   { source: 'office-frame', event: 'ready' | 'saved' | 'error', message? }
 */
import { useEffect, useState } from 'react';
import { AlertTriangle, RotateCcw } from 'lucide-react';
import { apiUrl } from '@hostwerk/api';
import { Loading, Empty, Btn } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { url } from '../lib/api';
import { uiLang } from './common';

export function IntegrationEditor({ file, app, onEvent }) {
  const [st, setSt] = useState({ phase: 'loading' });
  const [tries, setTries] = useState(0);
  const src = apiUrl(url('/integrations/frame', { ...(file.root ? { root: file.root } : {}), path: file.path, app, lang: uiLang(), n: tries || undefined }));
  useEffect(() => {
    const h = (e) => {
      const d = e.data;
      if (!d || d.source !== 'office-frame') return;
      if (d.event === 'ready') setSt({ phase: 'ready' });
      if (d.event === 'error') setSt({ phase: 'error', error: d.message || t('The full editor reported an error.') });
      if (onEvent) onEvent(d);
    };
    window.addEventListener('message', h);
    return () => window.removeEventListener('message', h);
  }, [onEvent]);
  useEffect(() => { setSt({ phase: 'loading' }); }, [src]);
  const name = 'ONLYOFFICE';
  if (st.phase === 'error') return <Empty icon={AlertTriangle} title="The full editor could not be opened" text={st.error} action={<Btn icon={RotateCcw} onClick={() => setTries((n) => n + 1)}>Try again</Btn>} />;
  return (
    <div className="of-shell-integration">
      {st.phase === 'loading' && <div className="of-shell-integration-wait"><Loading inline text={t('Starting {app} …', { app: name })} /></div>}
      <iframe key={src} src={src} title={`${name} – ${file.name}`} className="of-shell-integration-frame" allow="clipboard-read; clipboard-write; fullscreen" onLoad={() => setSt((s) => (s.phase === 'loading' ? { phase: 'ready' } : s))} />
    </div>
  );
}
