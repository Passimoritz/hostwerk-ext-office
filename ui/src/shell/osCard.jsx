/*
 * "Hostwerk as the default app for office files" – the big card (Office settings, top) and its compact form (Office
 * folder in Files, first-run hint). State from GET /office/os (backend osreg.js → host.fileTypes.state('office')):
 * per kind ✓ default / only under "Open with" / not registered. Windows does not let any app make itself the default
 * (UserChoice) – the big button opens Windows' "Default apps" page for Hostwerk Office, where the user confirms once.
 * Linux: .desktop file + xdg-mime default after asking. macOS / no desktop app: an honest notice.
 */
import { useEffect, useRef, useState } from 'react';
import { MonitorCheck, CheckCircle2, CircleDashed, MinusCircle, ExternalLink, X, Plus, Trash2, RefreshCw } from 'lucide-react';
import { get, post } from '@hostwerk/api';
import { Btn, Callout, Loading, confirmDialog } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { toast } from '../lib/toast';
import { saveSettings } from './store';
import { FileTypeIcon } from '../lib/fileIcons';

const KIND_ROWS = [['docs', 'Documents', 'docx'], ['sheets', 'Spreadsheets', 'xlsx'], ['slides', 'Presentations', 'pptx'], ['pdf', 'PDF', 'pdf']];
const STATE_TEXT = {
  default: ['Hostwerk opens them', CheckCircle2, 'c-ok'],
  partial: ['Some open in Hostwerk', CheckCircle2, 'c-warn'],
  openwith: ['Only under “Open with”', CircleDashed, 'muted'],
  none: ['Not registered', MinusCircle, 'muted'],
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
const allDefault = (os) => !!os && os.registered && KIND_ROWS.every(([k]) => os.kinds && os.kinds[k] && os.kinds[k].state === 'default');

/** After "Make default": Windows' settings are open – look again every few seconds for two minutes */
function useWatch() {
  const timer = useRef(null);
  useEffect(() => () => clearInterval(timer.current), []);
  return () => {
    clearInterval(timer.current);
    const until = Date.now() + 120000;
    timer.current = setInterval(() => { if (Date.now() > until) { clearInterval(timer.current); return; } loadOs().then((s) => { if (allDefault(s)) clearInterval(timer.current); }).catch(() => {}); }, 3000);
  };
}

async function makeDefault(os, watch) {
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
  toast(t('Windows’ settings are open – choose “Hostwerk Office” there and confirm once.'), 'info', 9000);
  watch();
}
/** setOn(true|false): the setting integrations.osOpenWith decides – saving it registers / removes it in the backend */
async function setOpenWith(on, onSetting) {
  await saveSettings({ integrations: { osOpenWith: on } });
  if (onSetting) onSetting(on);
  // switched on again although the setting was already on (removed by hand under Hostwerk → Settings): register now
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

/** The big card – Office settings (top). onSetting(on) keeps the settings draft in step. */
export function DefaultAppCard({ onSetting }) {
  const { os, reload } = useOs();
  const watch = useWatch();
  const [busy, setBusy] = useState(false);
  const act = (fn) => async () => { setBusy(true); try { await fn(); } catch (e) { toast.error(e); } setBusy(false); };
  const why = os ? whyOf(os) : null;
  const able = os && !os.error && os.supported && !os.code;
  const win = os && os.platform === 'win32';
  return (
    <div className="card of-shell-os is-hero">
      <div className="card-head"><span className="of-shell-os-badge" aria-hidden="true"><MonitorCheck size={22} /></span><h3>Hostwerk as the default app for office files</h3>
        {os && allDefault(os) && <span className="badge ok">Default</span>}
        {os && !allDefault(os) && os.registered && <span className="badge">Under “Open with”</span>}
        <span className="right"><Btn className="sm ghost icon" icon={RefreshCw} title="Check again" aria-label="Check again" onClick={reload} /></span>
      </div>
      <div className="card-body stack">
        <p className="small muted of-shell-p">Open Word, Excel, PowerPoint, OpenDocument and PDF files from Explorer or the desktop straight in Hostwerk – by double-click, or with right-click → “Open with”.</p>
        {!os ? <Loading inline /> : why || <>
          <div className="of-shell-os-kinds">
            {KIND_ROWS.map(([k, label, I]) => {
              const st = (os.kinds && os.kinds[k]) || { state: 'none' };
              const [text, SI, cls] = STATE_TEXT[st.state] || STATE_TEXT.none;
              return (
                <div key={k} className="of-shell-os-kind">
                  <FileTypeIcon ext={I} size={20} /><b className="small">{t(label)}</b>
                  <span className={`small ${cls} of-shell-os-state`}><SI size={14} aria-hidden="true" />{t(text)}</span>
                </div>
              );
            })}
          </div>
          {able && <div className="row of-shell-wrap">
            <Btn className="primary lg of-shell-os-main" icon={win ? ExternalLink : CheckCircle2} busy={busy} disabled={busy} onClick={act(() => makeDefault(os, watch))}>Make it the default app</Btn>
            {os.registered
              ? <Btn className="ghost sm" icon={Trash2} disabled={busy} onClick={act(() => setOpenWith(false, onSetting))}>Remove from “Open with”</Btn>
              : <Btn className="sm" icon={Plus} disabled={busy} onClick={act(() => setOpenWith(true, onSetting))}>Show under “Open with”</Btn>}
          </div>}
          {able && win && (
            <ol className="small of-shell-os-steps">
              <li>{t('Windows opens “Default apps” for Hostwerk Office.')}</li>
              <li>{t('Click “Set default” (or pick Hostwerk Office for single file types) – Windows asks you to confirm once. That is Windows’ rule: no app may make itself the default.')}</li>
            </ol>
          )}
          {able && os.platform === 'linux' && <p className="tiny muted of-shell-p">{t('Hostwerk asks before it changes the default – removing it later restores your previous apps.')}</p>}
          {os.stale && <Callout kind="warn" dismissible={false}>{t('Hostwerk was moved – press “Show under “Open with”” once so Windows finds it again.')}</Callout>}
        </>}
      </div>
    </div>
  );
}

const HINT_KEY = 'office:os-hint-hidden';
const hintHidden = () => { try { return localStorage.getItem(HINT_KEY) === '1'; } catch { return false; } };
/** Compact form for the Office folder in Files – only while Hostwerk is not the default yet; can be hidden */
export function DefaultAppHint() {
  const { os } = useOs();
  const watch = useWatch();
  const [hidden, setHidden] = useState(hintHidden);
  if (hidden || !os || os.error || !os.supported || os.code || allDefault(os)) return null;
  const hide = () => { try { localStorage.setItem(HINT_KEY, '1'); } catch { /* this window only */ } setHidden(true); };
  return (
    <div className="of-shell-os-hint">
      <MonitorCheck size={18} aria-hidden="true" />
      <div className="grow of-shell-min0">
        <b className="small">{t('Open office files from Explorer in Hostwerk')}</b>
        <div className="tiny muted">{os.registered ? t('Hostwerk is under “Open with” – make it the default app with one click (Windows asks once).') : t('Show Hostwerk under “Open with” and make it the default app.')}</div>
      </div>
      <Btn className="primary sm" icon={os.platform === 'win32' ? ExternalLink : CheckCircle2} onClick={() => makeDefault(os, watch).catch((e) => toast.error(e))}>Make it the default app</Btn>
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
      const text = s.platform === 'win32' ? t('Hostwerk Office now appears under “Open with” for office files. Make it the default app under Settings → Office.') : t('Hostwerk Office now appears under “Open with” for office files.');
      if (typeof toast.done === 'function') toast.done(text, { label: t('Settings'), href: '#/settings/office' }); else toast.success(text);
    }).catch(() => {});
  }, 6000);
}
