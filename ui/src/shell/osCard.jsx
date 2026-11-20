/*
 * "Hostwerk Office as the default app" – the hero card (Office settings → Default app) and its compact form (Office
 * folder in Files). State from GET /office/os (osreg.js: host.fileTypes + Windows' real choice incl. UserChoiceLatest).
 *
 * Windows 10/11 never let an app make itself the default – the user confirms it in Windows' own UI. Per kind the card
 * opens Windows' "How do you want to open this file?" dialog for that type (POST /office/os/openwith – a sample file of
 * the type, created from the editor chunk); the user picks "Hostwerk Office" → "Always". Then the card polls for a
 * minute and shows ✓. Settings → Default apps is the secondary path. Linux: xdg-mime after asking. macOS / server: a notice.
 */
import { useEffect, useRef, useState } from 'react';
import { MonitorCheck, CheckCircle2, CircleDashed, MinusCircle, ExternalLink, X, RefreshCw, MousePointerClick, ListChecks, Loader2, Settings2 } from 'lucide-react';
import { get, post } from '@hostwerk/api';
import { Btn, Callout, Loading, confirmDialog, Switch } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { toast } from '../lib/toast';
import { office } from '../lib/api';
import { saveSettings } from './store';
import { sampleBytes } from './ops';
import { FileTypeIcon } from '../lib/fileIcons';

/** [kind, label, icon ext, the types it covers (shown small)] */
export const KIND_ROWS = [
  ['docs', 'Documents', 'docx', 'Word, OpenDocument, RTF, Markdown'],
  ['sheets', 'Spreadsheets', 'xlsx', 'Excel, OpenDocument, CSV'],
  ['slides', 'Presentations', 'pptx', 'PowerPoint, OpenDocument'],
  ['pdf', 'PDF', 'pdf', 'PDF documents'],
];
const STATE = {
  default: ['Default app', CheckCircle2, 'ok'],
  partial: ['Partly', CheckCircle2, 'warn'],
  openwith: ['Only under “Open with”', CircleDashed, ''],
  none: ['Not registered', MinusCircle, ''],
};

let cache = null;
const subs = new Set();
const emit = (s) => { cache = s; subs.forEach((f) => f(s)); };
export const loadOs = () => get('/office/os').then((s) => { emit(s); return s; });
/** { os, reload } – os = null while loading, { error } when the backend is older */
export function useOs() {
  const [s, setS] = useState(cache);
  useEffect(() => { subs.add(setS); loadOs().catch((e) => emit({ error: e.message || String(e) })); return () => subs.delete(setS); }, []);
  return { os: s, reload: () => loadOs().catch(() => {}) };
}
const kindState = (os, k) => (os && os.kinds && os.kinds[k] && os.kinds[k].state) || 'none';
export const allDefault = (os) => !!os && os.registered && KIND_ROWS.every(([k]) => kindState(os, k) === 'default');
const able = (os) => !!os && !os.error && os.supported && !os.code;

/** Settings → Office → Default app (Hostwerk's settings tab, the Office group remembered by useRemembered) */
export function gotoOsSettings() {
  try { localStorage.setItem('hw:mem:office:settings-group', JSON.stringify('os')); } catch { /* opens the last group */ }
  location.hash = '#/settings/office';
}

/** Watch one kind (or all) for a minute after Windows' dialog/settings opened → done(kind) once it is the default */
function useWatch() {
  const timer = useRef(null);
  const [watching, setWatching] = useState(null);
  useEffect(() => () => clearInterval(timer.current), []);
  const start = (kind, label) => {
    clearInterval(timer.current);
    setWatching(kind);
    const until = Date.now() + 60000;
    timer.current = setInterval(() => {
      if (Date.now() > until) { clearInterval(timer.current); setWatching(null); return; }
      loadOs().then((s) => {
        const ok = kind === 'all' ? allDefault(s) : kindState(s, kind) === 'default';
        if (!ok) return;
        clearInterval(timer.current); setWatching(null);
        toast.success(kind === 'all' ? t('Hostwerk Office is now the default app for office files') : t('{kind} now open in Hostwerk Office', { kind: t(label) }));
      }).catch(() => {});
    }, 2000);
  };
  return [watching, start];
}

/** Windows: open the "Open with" dialog for one kind */
async function chooseFor(kind, watch) {
  const row = KIND_ROWS.find((r) => r[0] === kind);
  const ext = (cache && cache.kindExt && cache.kindExt[kind]) || row[2];
  const bytes = await sampleBytes(kind, ext).catch(() => null);
  const r = await office.osOpenWith(kind, bytes);
  if (r && r.state) emit({ ...cache, ...r.state });
  toast(t('Windows asks now: pick “Hostwerk Office” and click “Always”.'), 'info', 9000);
  watch(kind, row[1]);
}
/** Windows: Settings → Default apps (secondary path) · Linux: set the default after asking */
async function openDefaults(os, watch) {
  if (os.platform === 'linux') {
    const ok = await confirmDialog({ title: 'Make Hostwerk the default app?', text: 'Documents, spreadsheets, presentations and PDFs then open in Hostwerk Office by double-click. Removing it later restores your previous apps.', confirm: 'Make default' });
    if (!ok) return;
    const r = await post('/office/os/defaults');
    if (r && r.state) emit(r.state);
    toast.success('Hostwerk Office is now the default app for office files');
    return;
  }
  const r = await post('/office/os/defaults');
  if (r && r.state) emit(r.state);
  toast(t('Windows’ settings are open. If Hostwerk Office is not listed there, type a file type such as “.docx” into the search field at the top and pick Hostwerk Office.'), 'info', 12000);
  watch('all', '');
}
/** The setting integrations.osOpenWith decides – saving it registers / removes it in the backend */
export async function setOpenWith(on, onSetting) {
  await saveSettings({ integrations: { osOpenWith: on } });
  if (onSetting) onSetting(on);
  if (on) { const s = await post('/office/os/register'); emit(s); toast.success('Hostwerk now appears under “Open with” for office files'); } else { const s = await post('/office/os/unregister'); emit(s); toast('Removed from “Open with”'); }
}

/** why the card cannot act here (JSX) – or null */
function whyOf(os) {
  if (os.error) return <Callout kind="info" dismissible={false}>{t('This Hostwerk does not offer “Open with” yet – it comes with the next Hostwerk update.')}</Callout>;
  if (os.platform === 'darwin') return <Callout kind="info" dismissible={false} title="macOS">{t('On a Mac, apps register their file types themselves when they are installed – Hostwerk cannot change that while it runs. Open office files from Files here, or use “Open with → Other …” in Finder and pick Hostwerk.')}</Callout>;
  if (os.code === 'EDEVBUILD') return <Callout kind="info" dismissible={false} title="Developer start">{t('Hostwerk runs from its source code here. “Open with” is set up by the installed Hostwerk app – so no developer path ends up in Windows.')}</Callout>;
  if (os.code === 'ESTORE') return <Callout kind="info" dismissible={false}>{t('The Microsoft Store version registers its file types through Windows itself.')}</Callout>;
  if (!os.supported || os.code === 'ENODESKTOP' || os.code === 'ENOTSUPPORTED') return <Callout kind="info" dismissible={false}>{t('Only in the Hostwerk desktop app on Windows or Linux – this Hostwerk runs as a server without a screen.')}</Callout>;
  return null;
}

function Badge({ os }) {
  if (!os || !able(os)) return null;
  if (allDefault(os)) return <span className="badge ok">{t('Default app')}</span>;
  if (KIND_ROWS.some(([k]) => ['default', 'partial'].includes(kindState(os, k)))) return <span className="badge warn">{t('Partly')}</span>;
  if (os.registered) return <span className="badge">{t('Under “Open with”')}</span>;
  return <span className="badge">{t('Not set up')}</span>;
}

/** The hero – top of Settings → Office → Default app. onSetting(on) keeps the settings draft in step. */
export function DefaultAppCard({ onSetting }) {
  const { os, reload } = useOs();
  const [watching, watch] = useWatch();
  const [busy, setBusy] = useState(null);
  const act = (id, fn) => async () => { setBusy(id); try { await fn(); } catch (e) { toast.error(e); } setBusy(null); };
  const why = os ? whyOf(os) : null;
  const ok = able(os);
  const win = os && os.platform === 'win32';
  const dialog = ok && win && os.openWithDialog;
  const firstOpen = os && KIND_ROWS.find(([k]) => kindState(os, k) !== 'default');
  return (
    <div className="card of-os-hero">
      <div className="card-head">
        <span className="of-os-badge" aria-hidden="true"><MonitorCheck size={20} /></span>
        <div className="of-os-title"><h3>Hostwerk Office as the default app</h3><span className="small muted">Open office files from Explorer and the desktop by double-click.</span></div>
        <Badge os={os} />
        <div className="right"><Btn className="sm ghost icon" icon={RefreshCw} title="Check again" aria-label="Check again" onClick={reload} /></div>
      </div>
      <div className="card-body of-os-body">
        {!os ? <Loading inline /> : why || <>
          {dialog && (
            <ol className="of-os-steps" aria-label={t('How it works')}>
              <li><span className="of-os-step-n" aria-hidden="true">1</span><MousePointerClick size={16} aria-hidden="true" /><span>{t('Click “Choose …” next to a file type – Windows asks “How do you want to open this file?”.')}</span></li>
              <li><span className="of-os-step-n" aria-hidden="true">2</span><ListChecks size={16} aria-hidden="true" /><span>{t('Pick “Hostwerk Office” and click “Always”. Windows only lets you decide this yourself.')}</span></li>
            </ol>
          )}
          <div className="of-os-kinds" role="list">
            {KIND_ROWS.map(([k, label, icon, covers]) => {
              const st = kindState(os, k);
              const [text, SI, tone] = STATE[st] || STATE.none;
              const isDef = st === 'default';
              return (
                <div key={k} className={`of-os-kind${isDef ? ' is-default' : ''}`} role="listitem">
                  <FileTypeIcon ext={icon} size={28} />
                  <div className="of-os-kind-txt"><b>{t(label)}</b><span className="tiny muted">{t(covers)}</span></div>
                  <span className={`of-os-chip${tone ? ` ${tone}` : ''}`}>{watching === k ? <Loader2 size={13} className="spin" aria-hidden="true" /> : <SI size={13} aria-hidden="true" />}{watching === k ? t('Waiting for Windows …') : t(text)}</span>
                  {dialog && <Btn className={`sm ${isDef ? 'ghost' : firstOpen && firstOpen[0] === k ? 'primary' : ''}`} busy={busy === k} disabled={!!busy} title={t('Opens Windows’ “Open with” dialog for this file type')} onClick={act(k, () => chooseFor(k, watch))}>{isDef ? t('Change …') : t('Choose …')}</Btn>}
                </div>
              );
            })}
          </div>
          <div className="of-os-foot">
            {win && <Btn className={`sm ${dialog ? 'ghost' : 'primary'}`} icon={ExternalLink} busy={busy === 'settings'} disabled={!!busy} onClick={act('settings', () => openDefaults(os, watch))}>Windows settings …</Btn>}
            {os.platform === 'linux' && !allDefault(os) && <Btn className="sm primary" icon={CheckCircle2} busy={busy === 'settings'} disabled={!!busy} onClick={act('settings', () => openDefaults(os, watch))}>Make it the default app</Btn>}
            <span className="tiny muted of-os-foot-txt">{win ? t('Your previous apps stay under “Open with” – you can switch back at any time.') : t('Hostwerk asks before it changes the default – removing it later restores your previous apps.')}</span>
          </div>
          {os.stale && <Callout kind="warn" dismissible={false}>{t('Hostwerk was moved – switch “Show under “Open with”” off and on once so Windows finds it again.')}</Callout>}
        </>}
      </div>
    </div>
  );
}

/** "Show Hostwerk under Open with" – a setting row, saved at once (registers / removes in the backend) */
export function OpenWithRow({ value, onSetting }) {
  const { os } = useOs();
  const [busy, setBusy] = useState(false);
  const disabled = busy || !able(os);
  return (
    <Switch checked={!!value && !!(os && (os.registered || !able(os)))} disabled={disabled} onChange={async (on) => { setBusy(true); try { await setOpenWith(on, onSetting); } catch (e) { toast.error(e); } setBusy(false); }} />
  );
}

const HINT_KEY = 'office:os-hint-hidden';
const hintHidden = () => { try { return localStorage.getItem(HINT_KEY) === '1'; } catch { return false; } };
/** Compact form for the Office folder in Files – only while Hostwerk is not the default yet; can be hidden */
export function DefaultAppHint() {
  const { os } = useOs();
  const [hidden, setHidden] = useState(hintHidden);
  if (hidden || !os || !able(os) || allDefault(os)) return null;
  const hide = () => { try { localStorage.setItem(HINT_KEY, '1'); } catch { /* this window only */ } setHidden(true); };
  return (
    <div className="of-shell-os-hint">
      <MonitorCheck size={18} aria-hidden="true" />
      <div className="grow of-shell-min0">
        <b className="small">{t('Open office files from Explorer in Hostwerk')}</b>
        <div className="tiny muted">{os.platform === 'win32' ? t('Choose Hostwerk Office per file type – Windows asks you once.') : t('Show Hostwerk under “Open with” and make it the default app.')}</div>
      </div>
      <Btn className="primary sm" icon={Settings2} onClick={gotoOsSettings}>Set up …</Btn>
      <button type="button" className="btn ghost xs icon" title={t('Hide')} aria-label={t('Hide')} onClick={hide}><X size={14} /></button>
    </div>
  );
}

let noticeChecked = false;
/** Once after install: "Hostwerk now appears under Open with" – shown by the dialog host (always mounted in Hostwerk) */
export function checkOsNotice() {
  if (noticeChecked) return;
  noticeChecked = true;
  setTimeout(() => {
    loadOs().then((s) => {
      if (!s || !s.notice) return;
      post('/office/os/seen').catch(() => {});
      try { localStorage.setItem('hw:mem:office:settings-group', JSON.stringify('os')); } catch { /* the last group */ }
      const text = s.platform === 'win32' ? t('Hostwerk Office now appears under “Open with” for office files. Make it the default app under Settings → Office.') : t('Hostwerk Office now appears under “Open with” for office files.');
      if (typeof toast.done === 'function') toast.done(text, { label: t('Settings'), href: '#/settings/office' }); else toast.success(text);
    }).catch(() => {});
  }, 6000);
}
