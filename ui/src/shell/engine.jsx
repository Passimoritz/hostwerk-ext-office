/*
 * The Office engine (Hostwerk's own private LibreOffice, ARCHITECTURE.md "Full fidelity") and the one-click full
 * editor (ONLYOFFICE Document Server in Docker – the only full editor). Nothing needs a manual install:
 *   GET /office/engine · POST /office/engine/install · POST /office/engine/update · DELETE /office/engine
 *   GET /office/integrations/state · POST /office/integrations/setup { app: 'onlyoffice' }
 * settings.converters.autoEngine: 'ask' (one dialog the first time a file needs it) · 'auto' (silently) · 'off'.
 * The dialog lives in <EngineDialogHost/> – mounted by Office's dialog host (Hostwerk's sidebar), the editor host, the
 * settings and the previews; only the one mounted last shows it (useTopHost).
 */
import { useEffect, useRef, useState } from 'react';
import { Cpu, Download, CheckCircle2, AlertTriangle, Sparkles, Play } from 'lucide-react';
import { get, post, del } from '@hostwerk/api';
import { Modal, Btn, Progress, Callout, toast } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { settingsNow, loadStatus, saveSettings } from './store';
import { fmtBytes } from './util';

/* ---------------- Engine state (polled while installing) ---------------- */
let engine = null;
const subs = new Set();
let poll = null;
const emit = () => subs.forEach((f) => f(engine));
export async function loadEngine() {
  try { engine = await get('/office/engine'); } catch (e) { engine = { state: 'unknown', error: e.message }; }
  emit();
  const busy = engine && ['installing', 'updating', 'downloading', 'removing'].includes(engine.state);
  if (busy && !poll) poll = setInterval(async () => {
    try { engine = await get('/office/engine'); emit(); } catch { /* keep polling */ }
    if (!['installing', 'updating', 'downloading', 'removing'].includes(engine && engine.state)) {
      clearInterval(poll); poll = null; loadStatus(true).catch(() => {});
      if (engine && engine.state === 'ready') toast.success('The Office engine is ready');
      if (engine && engine.state === 'error') toast.error(t('The Office engine could not be set up: {error}', { error: engine.error || '' }));
    }
  }, 1500);
  return engine;
}
export const engineReady = (e) => !!e && ['ready', 'system'].includes(e.state);
export const engineBusy = (e) => !!e && ['installing', 'updating', 'downloading', 'removing'].includes(e.state);
export function useEngine() {
  const [e, setE] = useState(engine);
  useEffect(() => { subs.add(setE); loadEngine(); return () => subs.delete(setE); }, []);
  return e;
}
export async function installEngine() { await post('/office/engine/install'); return loadEngine(); }
export async function updateEngine() { await post('/office/engine/update'); return loadEngine(); }
export async function removeEngine() { await del('/office/engine'); return loadEngine(); }

/* ---------------- "Needs the engine" – ask once, install, never block ---------------- */
let ask = null; // { reason, resolve }
const askSubs = new Set();
let askedThisSession = false;
/**
 * Call when a file needs the engine (legacy formats, server conversion). Returns true when it is ready now.
 * 'auto' starts the setup silently, 'ask' shows the dialog once per session, 'off' does nothing.
 * The caller always continues with its built-in fallback meanwhile.
 */
export async function ensureEngine(reason = '') {
  const st = await loadStatus().catch(() => null);
  if (st && st.converters && (st.converters.libreoffice || st.converters.msoffice)) return true;
  const e = await loadEngine();
  if (engineReady(e)) return true;
  if (engineBusy(e)) return false;
  const mode = (settingsNow().converters || {}).autoEngine || 'ask';
  if (mode === 'off') return false;
  if (mode === 'auto') { installEngine().catch((er) => toast.error(er)); return false; }
  if (askedThisSession) return false;
  askedThisSession = true;
  return new Promise((resolve) => { ask = { reason, resolve }; askSubs.forEach((f) => f(ask)); });
}
export function openEngineDialog(reason = '') { return new Promise((resolve) => { ask = { reason, resolve }; askSubs.forEach((f) => f(ask)); }); }

export function EngineProgress({ e }) {
  if (!e) return null;
  const pct = e.progress && typeof e.progress.percent === 'number' ? e.progress.percent : null;
  return (
    <div className="stack of-shell-gap-2">
      <Progress value={pct || 0} indeterminate={pct == null} />
      <span className="small muted">{(e.progress && e.progress.text) || t('Setting up the Office engine …')}{pct != null ? ` · ${Math.round(pct)} %` : ''}</span>
    </div>
  );
}

/**
 * Several places mount a dialog host (Office's host in Hostwerk's sidebar, the editor, the settings, the preview in Files)
 * – only the one mounted last shows the dialog, so it never appears twice.
 */
const stacks = {};
const stackSubs = new Set();
export function useTopHost(group) {
  const me = useRef(null);
  if (!me.current) me.current = {};
  const [, force] = useState(0);
  useEffect(() => {
    const st = stacks[group] || (stacks[group] = []);
    st.push(me.current);
    const f = () => force((n) => n + 1);
    stackSubs.add(f); stackSubs.forEach((x) => x());
    return () => { const i = st.indexOf(me.current); if (i >= 0) st.splice(i, 1); stackSubs.delete(f); stackSubs.forEach((x) => x()); };
  }, [group]);
  const st = stacks[group] || [];
  return st[st.length - 1] === me.current;
}

export function EngineDialogHost() {
  const [a, setA] = useState(ask);
  const e = useEngine();
  const top = useTopHost('engine');
  useEffect(() => { askSubs.add(setA); return () => askSubs.delete(setA); }, []);
  if (!a || !top) return null;
  const close = (v) => { a.resolve(v); ask = null; askSubs.forEach((f) => f(null)); };
  const busy = engineBusy(e);
  const ready = engineReady(e);
  const size = e && e.size ? fmtBytes(e.size) : '≈ 350 MB';
  return (
    <Modal title="Set up the Office engine once" subtitle={t('About {size} – downloaded once, kept in Hostwerk’s data folder', { size })} onClose={() => close(ready)}
      footer={ready ? <Btn className="primary" onClick={() => close(true)}>Done</Btn> : busy ? <Btn onClick={() => close(false)}>Continue in the background</Btn> : <>
        <Btn className="ghost" onClick={async () => { await saveSettings({ converters: { autoEngine: 'off' } }).catch(() => {}); close(false); }}>Never ask</Btn>
        <Btn onClick={() => close(false)}>Not now</Btn>
        <Btn className="primary" icon={Download} onClick={() => installEngine().catch((er) => toast.error(er))}>Set up</Btn>
      </>}>
      <div className="stack">
        {a.reason && <Callout kind="info" dismissible={false}>{a.reason}</Callout>}
        <p className="small of-shell-p">The Office engine opens and saves old and rare formats (Word 97–2003, PowerPoint 97–2003, Pages, Keynote, WordPerfect …) exactly and converts between all formats. It is a private LibreOffice that only Office uses – no system installation, no admin rights, and it gets security updates automatically.</p>
        <p className="small muted of-shell-p">Until it is ready, the file opens with the built-in reader – you can keep working.</p>
        {busy && <EngineProgress e={e} />}
        {ready && <div className="row-tight c-ok"><CheckCircle2 size={16} aria-hidden="true" /><b>Ready – the file opens with full fidelity next time.</b></div>}
        {e && e.state === 'error' && <Callout kind="err" dismissible={false} title="Setup failed">{e.error || ''}</Callout>}
      </div>
    </Modal>
  );
}

/* ---------------- Full editor (ONLYOFFICE Document Server) ---------------- */
export const FULL_APP = 'onlyoffice';
const BUSY = ['installing', 'starting', 'pulling'];
const fullBusy = (st) => !!(st && st[FULL_APP] && BUSY.includes(st[FULL_APP].state));
let integ = null;
const integSubs = new Set();
let integPoll = null;
export async function loadIntegrations() {
  try { integ = await get('/office/integrations/state'); } catch (e) { integ = { error: e.message }; }
  integSubs.forEach((f) => f(integ));
  if (fullBusy(integ) && !integPoll) integPoll = setInterval(async () => {
    try { integ = await get('/office/integrations/state'); integSubs.forEach((f) => f(integ)); } catch { /* keep */ }
    if (!fullBusy(integ)) { clearInterval(integPoll); integPoll = null; }
  }, 2000);
  return integ;
}
export function useIntegrations() {
  const [s, setS] = useState(integ);
  useEffect(() => { integSubs.add(setS); loadIntegrations(); return () => integSubs.delete(setS); }, []);
  return s;
}
/** Is the full editor ready (settings URL or a running container)? → 'onlyoffice' | null */
export function fullEditorApp(s = settingsNow(), st = integ) {
  const ok = (st && st[FULL_APP] && st[FULL_APP].state === 'ready') || !!String(((s.integrations || {})[FULL_APP] || {}).url || '').trim();
  return ok ? FULL_APP : null;
}
export async function setupFullEditor() {
  await post('/office/integrations/setup', { app: FULL_APP });
  toast(t('Setting up the full editor – this takes a few minutes the first time'));
  return loadIntegrations();
}
/** The container exists but is stopped (state 'stopped') → start it again */
export async function startFullEditor() {
  await post('/office/integrations/start', { app: FULL_APP });
  return loadIntegrations();
}
/** Offer of the full editor in a banner/menu: sets it up first if needed. Returns the app id when it can open now. */
export async function ensureFullEditor() {
  let st = await loadIntegrations();
  if (st && st[FULL_APP] && st[FULL_APP].state === 'stopped') {
    toast(t('Starting the full editor …'));
    st = await startFullEditor().catch((e) => { toast.error(e); return st; });
  }
  const app = fullEditorApp(settingsNow(), st);
  if (app) return app;
  if (fullBusy(st)) { toast(t('The full editor is still being set up – a moment, please.')); return null; }
  return new Promise((resolve) => { fullAsk = { resolve }; fullSubs.forEach((f) => f(fullAsk)); });
}

/** Setup state of the full editor: progress while it is set up, error, or the Set up button (shared by dialog + settings) */
export function FullEditorSetup({ st, docker = true, children }) {
  const s = (st && st[FULL_APP]) || {};
  const busy = BUSY.includes(s.state);
  return (
    <div className="of-shell-choice">
      <Sparkles size={18} aria-hidden="true" />
      <div className="grow stack of-shell-gap-1">
        <b>ONLYOFFICE Docs</b>
        <span className="small muted">{t('Looks like Microsoft Office – very good with docx, xlsx and pptx.')}</span>
        {busy && <><Progress value={(s.progress && s.progress.percent) || 0} indeterminate={!(s.progress && s.progress.percent)} /><span className="tiny muted">{(s.progress && s.progress.text) || t('Setting up …')}</span></>}
        {s.state === 'ready' && children}
        {s.state === 'error' && <span className="small c-err"><AlertTriangle size={13} aria-hidden="true" /> {s.error}</span>}
      </div>
      {s.state === 'ready' ? <span className="badge ok">Ready</span> : s.state === 'stopped' ? <Btn className="sm primary" icon={Play} disabled={!docker} title={t('The full editor is set up but stopped')} onClick={() => startFullEditor().catch((er) => toast.error(er))}>{t('Start')}</Btn> : <Btn className="sm primary" icon={Cpu} disabled={busy || !docker} onClick={() => setupFullEditor().catch((er) => toast.error(er))}>{busy ? t('Setting up …') : t('Set up full editor')}</Btn>}
    </div>
  );
}

let fullAsk = null;
const fullSubs = new Set();
export function FullEditorDialogHost() {
  const [a, setA] = useState(fullAsk);
  const st = useIntegrations();
  const top = useTopHost('full');
  useEffect(() => { fullSubs.add(setA); return () => fullSubs.delete(setA); }, []);
  if (!a || !top) return null;
  const close = (v) => { a.resolve(v); fullAsk = null; fullSubs.forEach((f) => f(null)); };
  const docker = !st || st.docker !== false;
  const ready = st && st[FULL_APP] && st[FULL_APP].state === 'ready' ? FULL_APP : null;
  return (
    <Modal title="Set up the full editor" subtitle="100 % compatible editing in the browser – track changes, SmartArt, macros kept, pivot tables" onClose={() => close(ready)}
      footer={ready ? <Btn className="primary" onClick={() => close(ready)}>Open now</Btn> : <Btn onClick={() => close(null)}>Later</Btn>}>
      <div className="stack">
        {!docker && <Callout kind="warn" dismissible={false} title="Docker is needed">{st && st.dockerHint ? st.dockerHint : t('Hostwerk sets up Docker for you under Projects & apps → Docker. Then come back here.')}</Callout>}
        <FullEditorSetup st={st} docker={docker} />
      </div>
    </Modal>
  );
}
