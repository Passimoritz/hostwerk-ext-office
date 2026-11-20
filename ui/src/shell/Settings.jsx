/*
 * Office settings – shown as the tab "Office" in Hostwerk's settings (slot settings.tabs, prop `embedded`) and as
 * #/ext/office/settings. Built like Hostwerk's own settings: a list of sections on the left (a row above the content when
 * the space is narrow), per section a short head and cards of SettingRows (title + one line, control on the right,
 * longer explanations as ⓘ). Saving works like Hostwerk's: switches, segments and lists save at once; typed values save
 * a moment after typing (paths, addresses and secrets when the field is left) – a small status in the head says so.
 */
import { createContext, useContext, useEffect, useMemo, useRef, useState } from 'react';
import {
  Settings2, FileText, Sheet, Presentation, FileType, Cpu, Plug, Share2, ArrowLeft, CheckCircle2, XCircle, Download, RotateCcw, Trash2, FlaskConical, Sparkles, Cloud, LogOut, RefreshCw, MonitorCheck,
  LayoutTemplate, FileSpreadsheet, Calculator, PenLine, FolderOpen, Search, Folders, FilePlus2, MonitorPlay, Save, Link2, History, Type, Package, HardDrive, Loader2, Server, MousePointerClick, AlertTriangle,
} from 'lucide-react';
import { get, post, del } from '@hostwerk/api';
import { SettingsLayout, SettingRow, Switch, Seg, PathInput, Callout, Btn, Loading, Empty, Progress, toast, confirmDialog, useRemembered, InfoTip } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { office, download } from '../lib/api';
import { DEFAULT_SETTINGS, getIn, setIn, diffSettings, validateSettings, groupOfPath, SETTING_GROUPS, LAZY_PATHS, GROUP_RESET } from './settingsModel';
import { useOfficeSettings, useOfficeStatus, saveSettings, loadStatus } from './store';
import { parseQuery, fmtBytes } from './util';
import { useEngine, engineReady, engineBusy, installEngine, updateEngine, removeEngine, EngineProgress, EngineDialogHost, useIntegrations, FullEditorSetup, FullEditorDialogHost } from './engine';
import { LIB_ROOT, filesRoute } from './ops';
import { DefaultAppCard, OpenWithRow } from './osCard';
import { TemplatesPane } from './templates';
import { openLink, Keep } from './common';

const ICONS = { general: Settings2, os: MonitorCheck, docs: FileText, sheets: Sheet, slides: Presentation, pdf: FileType, templates: LayoutTemplate, converters: Cpu, fulleditor: Sparkles, google: Cloud, sharing: Share2, storage: HardDrive, components: Package };
const FONTS = ['Calibri', 'Arial', 'Aptos', 'Cambria', 'Carlito', 'Georgia', 'Helvetica', 'Liberation Sans', 'Liberation Serif', 'Open Sans', 'Roboto', 'Segoe UI', 'Times New Roman', 'Verdana'];
const LANGS = [['auto', 'Like Hostwerk'], ['de-DE', 'German (Germany)'], ['de-AT', 'German (Austria)'], ['de-CH', 'German (Switzerland)'], ['en-US', 'English (US)'], ['en-GB', 'English (UK)'], ['fr-FR', 'French'], ['es-ES', 'Spanish'], ['it-IT', 'Italian'], ['nl-NL', 'Dutch'], ['pl-PL', 'Polish'], ['pt-PT', 'Portuguese']];
const ZOOMS = [50, 75, 90, 100, 110, 125, 150, 175, 200];
const RowCtx = createContext({ v: () => undefined, set: () => () => {}, err: () => null });

/* ---- rows (module level, so inputs keep their focus while typing) ---- */
function Toggle({ p, title, text, tip, disabled }) {
  const { v, set } = useContext(RowCtx);
  return <SettingRow title={<>{title}{tip && <InfoTip pop title={title}>{tip}</InfoTip>}</>} text={text}><Switch checked={!!v(p)} disabled={disabled} onChange={set(p)} /></SettingRow>;
}
function Num({ p, title, text, unit, step = 1, disabled }) {
  const { v, set, err } = useContext(RowCtx);
  return (
    <SettingRow title={title} text={text}>
      <div className="of-set-ctl"><div className="of-set-num"><input className="input" type="number" inputMode="decimal" step={step} disabled={disabled} value={v(p) ?? ''} onChange={(e) => set(p)(e.target.value === '' ? '' : Number(e.target.value))} />{unit && <span className="small muted">{t(unit)}</span>}</div>{err(p)}</div>
    </SettingRow>
  );
}
function Pick({ p, title, text, options }) {
  const { v, set } = useContext(RowCtx);
  return <SettingRow title={title} text={text}><div className="of-set-ctl"><Seg value={String(v(p))} onChange={(x) => set(p)(typeof getIn(DEFAULT_SETTINGS, p) === 'number' ? Number(x) : x)} options={options.map(([id, label]) => ({ id: String(id), label: t(label) }))} /></div></SettingRow>;
}
function Select({ p, title, text, options, num = false }) {
  const { v, set } = useContext(RowCtx);
  return (
    <SettingRow title={title} text={text}>
      <select className="input select of-set-select" value={String(v(p))} onChange={(e) => set(p)(num && e.target.value !== 'auto' && !Number.isNaN(Number(e.target.value)) ? Number(e.target.value) : e.target.value)}>
        {options.map(([id, label]) => <option key={id} value={String(id)}>{t(label)}</option>)}
      </select>
    </SettingRow>
  );
}
function Text({ p, title, text, placeholder, type = 'text', list, mono = false }) {
  const { v, set, err } = useContext(RowCtx);
  return (
    <SettingRow title={title} text={text}><div className="of-set-ctl"><input className={`input of-set-text${mono ? ' mono' : ''}`} type={type} value={v(p) || ''} placeholder={placeholder ? t(placeholder) : ''} list={list} autoComplete={type === 'password' ? 'new-password' : 'off'} spellCheck={false} onChange={(e) => set(p)(e.target.value)} />{err(p)}</div></SettingRow>
  );
}
const Font = ({ p, title = 'Font' }) => <Text p={p} title={title} list="of-set-fonts" />;
/** A card of rows: icon + title (+ ⓘ, + right side) */
function Card({ icon: I, title, tip, right, children, className = '' }) {
  return (
    <div className={`card of-set-card ${className}`}>
      <div className="card-head">{I && <I size={16} aria-hidden="true" />}<h3>{title}</h3>{tip && <InfoTip title={title}>{tip}</InfoTip>}{right && <div className="right">{right}</div>}</div>
      <div className="card-body">{children}</div>
    </div>
  );
}

/** Head of a section: what it is about + saving status */
function PaneHead({ group, state }) {
  const g = SETTING_GROUPS.find((x) => x.id === group);
  const I = ICONS[group] || Settings2;
  return (
    <div className="of-set-head">
      <span className="of-set-head-ico" aria-hidden="true"><I size={18} /></span>
      <div className="of-set-head-txt"><h2>{t(g ? g.label : '')}</h2><p>{t(g ? g.text : '')}</p></div>
      <span className={`of-set-status ${state}`} role="status" aria-live="polite">
        {state === 'saving' ? <><Loader2 size={13} className="spin" aria-hidden="true" />{t('Saving …')}</> : state === 'saved' ? <><CheckCircle2 size={13} aria-hidden="true" />{t('Saved')}</> : state === 'invalid' ? <><AlertTriangle size={13} aria-hidden="true" />{t('Not saved – check the marked field')}</> : state === 'error' ? <><XCircle size={13} aria-hidden="true" />{t('Not saved')}</> : null}
      </span>
    </div>
  );
}

export default function OfficeSettings({ embedded = false, query = '' }) {
  const { settings, error, reload } = useOfficeSettings();
  const [group0, setGroup] = useRemembered('office:settings-group', 'general');
  const group = SETTING_GROUPS.some((x) => x.id === group0) ? group0 : ({ library: 'storage', integrations: 'fulleditor' }[group0] || 'general');
  const [draft, setDraft] = useState(null);
  const [saveState, setSaveState] = useState('');
  const ref = useRef({});
  ref.current.settings = settings; ref.current.draft = draft;
  useEffect(() => { const g = parseQuery(query).group; if (g && SETTING_GROUPS.some((x) => x.id === g)) setGroup(g); }, [query]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (settings && !draft) setDraft(settings); }, [settings]); // eslint-disable-line react-hooks/exhaustive-deps
  const errors = useMemo(() => (draft ? validateSettings(draft) : {}), [draft]);

  /** Save what changed (all = also paths/addresses/secrets, i.e. after leaving a field) */
  const flush = async (all = true) => {
    const { settings: base, draft: d } = ref.current;
    if (!base || !d) return;
    const errs = validateSettings(d);
    let paths = diffSettings(base, d);
    if (!all) paths = paths.filter((p) => !LAZY_PATHS.some((l) => p === l || p.startsWith(`${l}.`)));
    const bad = paths.filter((p) => errs[p]);
    paths = paths.filter((p) => !errs[p]);
    if (!paths.length) { if (bad.length) setSaveState('invalid'); return; }
    let patch = {};
    for (const p of paths) patch = setIn(patch, p, getIn(d, p));
    setSaveState('saving');
    try { await saveSettings(patch); setSaveState(bad.length ? 'invalid' : 'saved'); loadStatus(true).catch(() => {}); } catch (e) { setSaveState('error'); toast.error(e); }
  };
  ref.current.flush = flush;
  useEffect(() => {
    if (!draft || !settings) return undefined;
    const h = setTimeout(() => ref.current.flush(false), 700);
    return () => clearTimeout(h);
  }, [draft]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => () => { if (ref.current.flush) ref.current.flush(true); }, []);
  useEffect(() => { if (saveState !== 'saved') return undefined; const h = setTimeout(() => setSaveState(''), 2500); return () => clearTimeout(h); }, [saveState]);

  if (error && !settings) return <Empty icon={XCircle} title="The settings could not be loaded" text={error} action={<Btn icon={RotateCcw} onClick={reload}>Try again</Btn>} />;
  if (!settings || !draft) return <Loading inline />;
  const v = (p) => getIn(draft, p);
  const set = (p) => (val) => setDraft((d) => setIn(d, p, val));
  const err = (p) => (errors[p] ? <span className="hint c-err">{t(errors[p].text, errors[p].vars)}</span> : null);
  const warn = {};
  for (const p of Object.keys(errors)) { const g = groupOfPath(p); warn[g] = (warn[g] || 0) + 1; }
  const groups = SETTING_GROUPS.map((g) => ({ id: g.id, label: t(g.label), icon: ICONS[g.id], title: t(g.text), warn: warn[g.id] || 0 }));
  const reset = async () => {
    if (!(await confirmDialog({ title: 'Reset this section?', text: 'The values of this section go back to the defaults.', confirm: 'Reset' }))) return;
    setDraft((d) => (GROUP_RESET[group] || []).reduce((acc, k) => setIn(acc, k, getIn(DEFAULT_SETTINGS, k)), d));
  };

  let pane = null;
  if (group === 'general') pane = <>
    <Card icon={Save} title="Saving">
      <Toggle p="general.autosave" title="Save automatically" text="A few seconds after your last change – every save keeps the previous content as a version." />
      {v('general.autosave') && <Num p="general.autosaveSeconds" title="Wait before saving" unit="seconds" />}
      <Toggle p="general.confirmDelete" title="Ask before moving to the trash" text="Deleted files can be restored from the trash either way." />
    </Card>
    <Card icon={FilePlus2} title="New files">
      <Pick p="general.newFormat" title="Format of new files" text="You can pick another one in the “New” dialog at any time." options={[['ms', 'Microsoft Office'], ['odf', 'OpenDocument']]} />
      <Text p="general.author" title="Author name" text="Written into new files and comments." placeholder="e.g. Anna Berger" />
    </Card>
    <Card icon={Type} title="Language and units">
      <Select p="general.language" title="Language of documents" text="For spelling, hyphenation and number formats." options={LANGS} />
      <Toggle p="general.spellcheck" title="Check spelling" text="Marks unknown words while you type." />
      <Pick p="general.units" title="Units" options={[['cm', 'Centimetres'], ['in', 'Inches']]} />
    </Card>
    <Card icon={Folders} title="In Files">
      <Toggle p="general.openFromFiles" title="Open office files in Office" text="A click in Files opens documents, spreadsheets, presentations and PDFs here." tip="In every folder of Files. Off: Files shows its preview – right-click → “Open in Office” still works." />
      <Select p="integrations.defaultOpen" title="A click opens files in" options={[['builtin', 'Built-in editor'], ['onlyoffice', 'Full editor (ONLYOFFICE)'], ['desktop', 'Desktop app (Word, Excel …)']]} />
      <Num p="general.recentCount" title="Recent files" text="How many the Office section lists." unit="files" />
    </Card>
  </>;
  else if (group === 'os') pane = <>
    <DefaultAppCard onSetting={(on) => setDraft((d) => setIn(d, 'integrations.osOpenWith', on))} />
    <Card icon={MousePointerClick} title="Explorer and desktop" tip="Hostwerk registers itself per Windows user, without admin rights, for Word, Excel, PowerPoint, OpenDocument, PDF and the other formats Office opens. Your current apps stay as they are until you choose otherwise. Linux: a .desktop file.">
      <SettingRow title="Show Hostwerk under “Open with”" text="Right-click a file → “Open with” → Hostwerk Office. Off removes all entries again."><OpenWithRow value={v('integrations.osOpenWith')} onSetting={(on) => setDraft((d) => setIn(d, 'integrations.osOpenWith', on))} /></SettingRow>
    </Card>
  </>;
  else if (group === 'docs') pane = <>
    <Card icon={FilePlus2} title="New documents">
      <Font p="docs.font" />
      <Num p="docs.fontSize" title="Font size" unit="pt" step={0.5} />
      <Select p="docs.lineHeight" title="Line spacing" num options={[[1, 'Single (1.0)'], [1.08, '1.08'], [1.15, '1.15'], [1.5, '1.5'], [2, 'Double (2.0)']]} />
      <Num p="docs.paragraphSpacing" title="Space after paragraphs" unit="pt" />
      <Pick p="docs.pageSize" title="Paper size" options={[['A4', 'A4'], ['A5', 'A5'], ['Letter', 'Letter'], ['Legal', 'Legal']]} />
      <Pick p="docs.orientation" title="Orientation" options={[['portrait', 'Portrait'], ['landscape', 'Landscape']]} />
      <SettingRow title="Margins" text="In millimetres.">
        <div className="of-set-ctl"><div className="of-set-margins">{['top', 'right', 'bottom', 'left'].map((k) => <label key={k} className="of-set-margin"><span className="tiny muted">{t(k[0].toUpperCase() + k.slice(1))}</span><input className="input" type="number" value={v(`docs.margins.${k}`) ?? ''} onChange={(e) => set(`docs.margins.${k}`)(e.target.value === '' ? '' : Number(e.target.value))} aria-label={t(`Margin ${k}`)} /></label>)}</div>{['top', 'right', 'bottom', 'left'].map((k) => <span key={k}>{err(`docs.margins.${k}`)}</span>)}</div>
      </SettingRow>
    </Card>
    <Card icon={PenLine} title="Editing and view">
      <Select p="docs.zoom" title="Zoom" num options={ZOOMS.map((z) => [z, `${z} %`])} />
      <Toggle p="docs.ruler" title="Show the ruler" />
      <Toggle p="docs.showFormatting" title="Show formatting marks" text="Paragraph marks, spaces and tabs." />
      <Toggle p="docs.smartQuotes" title="Typographic quotes" text="Turns straight quotes into „…“ or “…” while typing." />
      <Toggle p="docs.autoLinks" title="Recognise links" text="Web and e-mail addresses become links." />
    </Card>
  </>;
  else if (group === 'sheets') pane = <>
    <Card icon={Sheet} title="New spreadsheets and view">
      <Font p="sheets.font" />
      <Num p="sheets.fontSize" title="Font size" unit="pt" step={0.5} />
      <Select p="sheets.zoom" title="Zoom" num options={ZOOMS.map((z) => [z, `${z} %`])} />
      <Toggle p="sheets.gridlines" title="Show gridlines" />
      <Toggle p="sheets.formulaBar" title="Show the formula bar" />
      <Toggle p="sheets.headings" title="Show row and column headings" />
    </Card>
    <Card icon={Calculator} title="Calculation and formats">
      <Toggle p="sheets.autoCalc" title="Calculate automatically" text="Off: F9 calculates." />
      <Select p="sheets.functionNames" title="Function names" text="Files always keep the English names." options={[['auto', 'Like the language'], ['en', 'English (SUM, IF)'], ['de', 'German (SUMME, WENN)']]} />
      <Select p="sheets.decimal" title="Decimal separator" options={[['auto', 'Like the language'], ['.', 'Point (1.5)'], [',', 'Comma (1,5)']]} />
      <Select p="sheets.dateFormat" title="Date format" options={[['auto', 'Like the language'], ['DD.MM.YYYY', '31.12.2026'], ['YYYY-MM-DD', '2026-12-31'], ['MM/DD/YYYY', '12/31/2026'], ['DD/MM/YYYY', '31/12/2026']]} />
      <Toggle p="sheets.r1c1" title="R1C1 references" text="Rows and columns as numbers instead of A1." />
    </Card>
    <Card icon={FileSpreadsheet} title="CSV files">
      <Select p="sheets.csvDelimiter" title="Separator" options={[['auto', 'Detect automatically'], [',', 'Comma (,)'], [';', 'Semicolon (;)'], ['\t', 'Tab'], ['|', 'Vertical bar (|)']]} />
      <Select p="sheets.csvEncoding" title="Text encoding" options={[['utf-8', 'UTF-8 (recommended)'], ['windows-1252', 'Windows-1252 (old Excel)'], ['iso-8859-1', 'ISO-8859-1'], ['utf-16le', 'UTF-16']]} />
    </Card>
  </>;
  else if (group === 'slides') pane = <>
    <Card icon={Presentation} title="New presentations">
      <Pick p="slides.size" title="Slide size" options={[['16:9', '16:9'], ['4:3', '4:3'], ['A4', 'A4']]} />
      <Select p="slides.theme" title="Design" options={[['hostwerk', 'Hostwerk'], ['light', 'Light'], ['dark', 'Dark'], ['minimal', 'Minimal'], ['classic', 'Classic']]} />
      <Font p="slides.font" />
    </Card>
    <Card icon={MonitorPlay} title="Presenting">
      <Toggle p="slides.notes" title="Show speaker notes while editing" />
      <Toggle p="slides.presenterView" title="Presenter view" text="Notes, next slide and timer in a second Hostwerk window." />
      <Select p="slides.transition" title="Transition" options={[['none', 'None'], ['fade', 'Fade'], ['slide', 'Slide in']]} />
      <Num p="slides.autoAdvance" title="Next slide automatically" text="0 = only by click or key." unit="seconds" />
      <Toggle p="slides.loop" title="Start again after the last slide" />
    </Card>
  </>;
  else if (group === 'pdf') pane = <>
    <Card icon={FileType} title="Reading">
      <Select p="pdf.zoom" title="Zoom when opening" options={[['auto', 'Automatic'], ['page-fit', 'Whole page'], ['page-width', 'Page width'], ...ZOOMS.map((z) => [String(z), `${z} %`])]} />
      <Pick p="pdf.scroll" title="Scrolling" options={[['vertical', 'Vertical'], ['horizontal', 'Horizontal'], ['page', 'Page by page']]} />
      <Toggle p="pdf.sidebar" title="Show page thumbnails" />
    </Card>
    <Card icon={PenLine} title="Signing">
      <Text p="pdf.signatureName" title="Name for signatures" text="Suggested when you sign a PDF." placeholder="e.g. Anna Berger" />
    </Card>
  </>;
  else if (group === 'templates') pane = <TemplatesPane v={v} set={set} />;
  else if (group === 'converters') pane = <ConvertersPane v={v} set={set} />;
  else if (group === 'fulleditor') pane = <FullEditorPane v={v} set={set} err={err} flush={flush} />;
  else if (group === 'google') pane = <GooglePane v={v} set={set} err={err} flush={flush} />;
  else if (group === 'sharing') pane = (
    <Card icon={Link2} title="Share links" right={<a className="btn sm ghost" href={filesRoute({ root: LIB_ROOT, dir: true })} onClick={() => { try { localStorage.setItem('office:section-view', 'shared'); } catch { /* ignore */ } }}><Share2 size={14} />{t('Shared files')}</a>}>
      <Toggle p="sharing.enabled" title="Allow share links" text="Read-only links for single files – PDF and pictures directly, office files as PDF." />
      <Num p="sharing.defaultDays" title="Links are valid for" text="0 = no end." unit="days" disabled={!v('sharing.enabled')} />
      <Toggle p="sharing.allowDownload" title="Allow download by default" disabled={!v('sharing.enabled')} />
    </Card>
  );
  else if (group === 'storage') pane = <>
    <Card icon={FolderOpen} title="Folder “Office documents”" tip="Files already there stay where they are; Files shows the new folder.">
      <div className="of-set-field">
        <label className="of-set-label">{t('Folder in Files')}</label>
        <PathInput mode="folder" title={t('Folder of “Office documents” in Files')} value={v('libraryDir')} onChange={set('libraryDir')} placeholder={t('Hostwerk data folder (default)')} />
        {err('libraryDir') || <span className="tiny muted">{t('Empty = Hostwerk’s data folder.')}</span>}
      </div>
    </Card>
    <Card icon={History} title="Trash and versions">
      <Num p="general.trashDays" title="Keep deleted files" text="Afterwards they are removed for good." unit="days" />
      <Num p="general.keepVersions" title="Versions per file" text="Older ones beyond this number are removed." unit="versions" />
      <Num p="general.versionDays" title="Keep versions" unit="days" />
    </Card>
    <Card icon={HardDrive} title="Backups">
      <SettingRow title="Back up the Office documents" text="Hostwerk’s backups include the folder – every backup is a state you can go back to."><a className="btn sm" href="#/backups">{t('Open backups')}</a></SettingRow>
    </Card>
  </>;
  else if (group === 'components') pane = <ComponentsPane />;

  const body = (
    <>
      <datalist id="of-set-fonts">{FONTS.map((f) => <option key={f} value={f} />)}</datalist>
      <RowCtx.Provider value={{ v, set, err }}>
        <div className="of-set-wrap">
          <SettingsLayout className="of-set" groups={groups} value={group} onChange={setGroup} label={t('Office settings')}>
            <div className="of-set-pane" onBlur={() => setTimeout(() => ref.current.flush(true), 0)}>
              <PaneHead group={group} state={saveState} />
              {pane}
              {(GROUP_RESET[group] || []).length > 0 && <div className="of-set-reset"><Btn className="ghost sm" icon={RotateCcw} onClick={reset}>Reset section to defaults</Btn></div>}
            </div>
          </SettingsLayout>
        </div>
      </RowCtx.Provider>
      <EngineDialogHost />
      <FullEditorDialogHost />
    </>
  );
  if (embedded) return <div className="of-shell-settings-embedded">{body}</div>;
  return (
    <div className="page fade-in">
      <div className="page-head">
        <div><h1><FileText size={22} /> Office settings</h1><p>How Office creates, saves, converts and shares your files.</p></div>
        <div className="actions"><a className="btn ghost" href={filesRoute({ root: LIB_ROOT, dir: true })}><ArrowLeft size={16} />{t('Office documents')}</a></div>
      </div>
      {body}
    </div>
  );
}

/* ---------------- Converting & engine ---------------- */
function ConvertersPane({ v, set }) {
  const { status, reload } = useOfficeStatus();
  const e = useEngine();
  const [test, setTest] = useState(null);
  const lo = status && status.converters && status.converters.libreoffice;
  const ms = status && status.converters && status.converters.msoffice;
  const runTest = async () => {
    const t0 = performance.now();
    try {
      const html = new TextEncoder().encode(`<!doctype html><html><body><h1>Office</h1><p>${t('Test conversion')} – ${new Date().toLocaleString()}</p></body></html>`);
      const out = await office.convert(html, 'html', 'pdf');
      setTest({ ok: true, ms: Math.round(performance.now() - t0), size: out.length, out });
    } catch (er) { setTest({ ok: false, error: er.message || String(er) }); }
  };
  const msApps = ms ? ['word', 'excel', 'powerpoint', 'visio'].filter((k) => ms[k]) : [];
  const engBadge = engineReady(e) ? <span className="badge ok">{t('Ready')}</span> : engineBusy(e) ? <span className="badge accent">{t('Setting up …')}</span> : <span className="badge">{t('Not set up')}</span>;
  return <>
    <Card icon={Cpu} title="Office engine" tip="A private LibreOffice that only Office uses – opens old and rare formats exactly and converts between all of them. No system installation, security updates come automatically." right={engBadge}>
      {e && engineBusy(e) && <div className="of-set-block"><EngineProgress e={e} /></div>}
      {e && e.state === 'error' && <div className="of-set-block"><Callout kind="err" dismissible={false} title="Setup failed">{e.error}</Callout></div>}
      <SettingRow title={!engineReady(e) ? t('Old and rare formats, exact conversion') : e.state === 'system' ? t('Uses the LibreOffice installed on this PC') : t('Version {v}', { v: e.version || '?' })} text={!engineReady(e) ? t('About 350 MB, set up once.') : e.state === 'system' ? t('No own engine needed.') : (e.size ? fmtBytes(e.size) : '')}>
        <div className="of-set-btns">
          {!engineReady(e) && <Btn className="primary sm" icon={Download} disabled={engineBusy(e)} onClick={() => installEngine()}>Set up</Btn>}
          {e && e.state === 'ready' && <Btn className="sm" icon={RefreshCw} onClick={() => updateEngine()}>Look for updates</Btn>}
          {e && e.state === 'ready' && <Btn className="sm ghost" icon={Trash2} onClick={async () => { if (await confirmDialog({ title: 'Remove the Office engine?', text: 'Old formats then open with the built-in readers only. You can set it up again at any time.', confirm: 'Remove', danger: true })) await removeEngine(); }}>Remove</Btn>}
        </div>
      </SettingRow>
      <SettingRow title="When a file needs the engine">
        <div className="of-set-ctl"><Seg value={v('converters.autoEngine')} onChange={set('converters.autoEngine')} options={[{ id: 'ask', label: t('Ask once') }, { id: 'auto', label: t('Set up by itself') }, { id: 'off', label: t('Never') }]} /></div>
      </SettingRow>
    </Card>
    <Card icon={Search} title="Found on this PC" right={<Btn className="sm ghost icon" icon={RefreshCw} title="Check again" aria-label="Check again" onClick={reload} />}>
      {!status ? <Loading inline /> : <>
        <SettingRow title="LibreOffice" text={lo ? <Keep>{`${lo.version || ''}${lo.path ? ` · ${lo.path}` : ''}`}</Keep> : t('not found')}>{lo ? <CheckCircle2 size={18} className="c-ok" aria-label={t('found')} /> : <XCircle size={18} className="muted" aria-label={t('not found')} />}</SettingRow>
        <SettingRow title="Microsoft Office" text={msApps.length ? <span title={msApps.map((k) => (typeof ms[k] === 'string' ? ms[k] : k)).join('\n')}>{msApps.map((k) => ({ word: 'Word', excel: 'Excel', powerpoint: 'PowerPoint', visio: 'Visio' }[k])).join(' · ')}</span> : t('not found')}>{msApps.length ? <CheckCircle2 size={18} className="c-ok" aria-label={t('found')} /> : <XCircle size={18} className="muted" aria-label={t('not found')} />}</SettingRow>
        <SettingRow title="Test conversion" text={test ? (test.ok ? t('Works – PDF with {size} in {ms} ms', { size: fmtBytes(test.size), ms: test.ms }) : test.error) : t('Turns a small page into a PDF.')}>
          <div className="of-set-btns">{test && test.ok && <Btn className="sm ghost" icon={Download} onClick={() => download(test.out, 'office-test.pdf', 'application/pdf')}>Download</Btn>}<Btn className="sm" icon={FlaskConical} onClick={runTest}>Test</Btn></div>
        </SettingRow>
      </>}
    </Card>
    <Card icon={Cpu} title="Which converter">
      <SettingRow title="Prefer" text="Microsoft Office converts its own formats most faithfully; LibreOffice knows more formats.">
        <div className="of-set-ctl"><Seg value={v('converters.preferred')} onChange={set('converters.preferred')} options={[{ id: 'auto', label: t('Automatic') }, { id: 'libreoffice', label: 'LibreOffice' }, { id: 'msoffice', label: 'Microsoft Office' }]} /></div>
      </SettingRow>
      <Toggle p="converters.msoffice" title="Use Microsoft Office when installed" text="Windows only – Office runs invisibly in the background." />
      <Num p="converters.timeout" title="Give up after" text="For very large files." unit="seconds" />
      <div className="of-set-field">
        <label className="of-set-label">{t('Path to LibreOffice')}</label>
        <PathInput mode="file" value={v('converters.libreoffice')} onChange={set('converters.libreoffice')} placeholder={t('automatic')} />
        <span className="tiny muted">{t('Empty = find it automatically (or use the Office engine).')}</span>
      </div>
    </Card>
  </>;
}

/* ---------------- Full editor (ONLYOFFICE) ---------------- */
function FullEditorPane({ v, set, err, flush }) {
  const st = useIntegrations();
  const { reload } = useOfficeStatus();
  const test = async () => {
    await flush(true);
    const s = await loadStatus(true);
    const r = s && s.integrations && s.integrations.onlyoffice;
    if (r && (r.ok || r.version)) toast.success(t('{app} answers{v}', { app: 'ONLYOFFICE', v: r.version ? ` – ${t('version {v}', { v: r.version })}` : '' }));
    else toast.error(t('{app} does not answer: {error}', { app: 'ONLYOFFICE', error: (r && r.error) || t('no answer') }));
    reload();
  };
  return <>
    <Card icon={Sparkles} title="Set up with one click" tip="Hostwerk starts ONLYOFFICE Docs in Docker and connects it. “Open in the full editor” then edits every file 100 % compatible inside Hostwerk – track changes, SmartArt, pivot tables and macros are kept.">
      {st && st.docker === false && <div className="of-set-block"><Callout kind="warn" dismissible={false} title="Docker is needed">{st.dockerHint || t('Set up Docker under Projects & apps → Docker, then come back here.')}</Callout></div>}
      <div className="of-set-block"><FullEditorSetup st={st} docker={!(st && st.docker === false)}>{st && st.onlyoffice && <span className="tiny muted">{t('Runs in Docker{url}', { url: st.onlyoffice.url ? ` · ${st.onlyoffice.url}` : '' })}</span>}</FullEditorSetup></div>
      <SettingRow title="A click opens files in" text="Also under General."><select className="input select of-set-select" value={v('integrations.defaultOpen')} onChange={(e) => set('integrations.defaultOpen')(e.target.value)}>
        {[['builtin', 'Built-in editor'], ['onlyoffice', 'Full editor (ONLYOFFICE)'], ['desktop', 'Desktop app (Word, Excel …)']].map(([id, l]) => <option key={id} value={id}>{t(l)}</option>)}
      </select></SettingRow>
    </Card>
    <Card icon={Server} title="Own ONLYOFFICE server" tip="Only if you already run a Document Server – the one-click setup above needs none of this.">
      <SettingRow title="Address" text="Saved when you leave the field.">
        <div className="of-set-ctl"><div className="of-set-inline"><input className="input of-set-text mono" value={v('integrations.onlyoffice.url') || ''} placeholder="https://…" spellCheck={false} onChange={(e) => set('integrations.onlyoffice.url')(e.target.value)} /><Btn className="sm" disabled={!v('integrations.onlyoffice.url')} onClick={test}>Test</Btn></div>{err('integrations.onlyoffice.url')}</div>
      </SettingRow>
      <Text p="integrations.onlyoffice.secret" title="Secret (JWT)" text="Stays on this PC." type="password" />
    </Card>
  </>;
}

/* ---------------- Google Drive ---------------- */
function GooglePane({ v, set, err, flush }) {
  const [g, setG] = useState(null);
  const loadG = () => get('/office/google/state').then(setG, () => setG({ unavailable: true }));
  useEffect(() => { loadG(); }, []);
  const connect = async () => {
    await flush(true); // a client id typed just now must reach the backend first
    const r = await post('/office/google/connect');
    if (r && r.connected) { setG(r); toast.success('Google Drive connected'); return; }
    if (r && r.url) await openLink(r.url);
    const h = setInterval(async () => { const s = await get('/office/google/state').catch(() => null); if (s && s.connected) { clearInterval(h); setG(s); toast.success('Google Drive connected'); } }, 3000);
    setTimeout(() => clearInterval(h), 300000);
  };
  const account = g && g.account ? (typeof g.account === 'object' ? g.account.email || g.account.name : g.account) : '?';
  return <>
    <Card icon={Cloud} title="Google Drive" tip="Opens Google Docs, Sheets and Slides (.gdoc, .gsheet, .gslides) right in Office’s editors – saving goes back to Google." right={g && g.connected ? <span className="badge ok">{t('Connected')}</span> : null}>
      {!g ? <Loading inline /> : g.unavailable ? <p className="small muted of-set-block">{t('Not available in this version of the Office backend yet.')}</p> : g.connected ? (
        <SettingRow title={<Keep>{account}</Keep>} text="Signed in – Google files open in Office."><Btn className="sm ghost" icon={LogOut} onClick={async () => { await post('/office/google/disconnect'); loadG(); }}>Disconnect</Btn></SettingRow>
      ) : (
        <SettingRow title="Connect Google Drive" text="You sign in in a Hostwerk window.">
          <Btn className="sm primary" icon={Cloud} disabled={!g.configured && !String(v('integrations.google.clientId') || '').trim()} title={!g.configured ? t('Enter the client ID of your own Google OAuth client first') : undefined} onClick={() => connect().catch((er) => toast.error(er))}>Connect …</Btn>
        </SettingRow>
      )}
    </Card>
    {g && !g.unavailable && !g.connected && (
      <Card icon={Plug} title="Own OAuth client" tip="Google needs an OAuth client of the type “Desktop app” from your own Google Cloud project – Hostwerk ships none.">
        <Text p="integrations.google.clientId" title="Client ID" placeholder="…apps.googleusercontent.com" mono />
        <Text p="integrations.google.clientSecret" title="Client secret" type="password" />
        {Array.isArray(g.guide) && g.guide.length > 0 && <details className="small of-set-details"><summary>{t('How to create the OAuth client')}</summary><ol>{g.guide.map((x, i) => <li key={i}>{x}</li>)}</ol>{g.redirectUri && <p className="tiny muted">{t('Redirect address')}: <span className="mono"><Keep>{g.redirectUri}</Keep></span></p>}</details>}
      </Card>
    )}
  </>;
}

/* ---------------- Components on demand ---------------- */
function compLabel(l) {
  const m = /^Text recognition: (.+)$/.exec(String(l || ''));
  return m ? t('Text recognition: {lang}', { lang: t(m[1]) }) : t(l);
}
function ComponentsPane() {
  const [list, setList] = useState(null);
  const [error, setError] = useState(null);
  const load = () => get('/office/components').then((r) => { setList((r && r.components) || []); setError(null); return r; }, (e) => setError(e.message || String(e)));
  useEffect(() => { load(); }, []);
  const busy = (list || []).some((c) => c.state === 'installing');
  useEffect(() => { if (!busy) return undefined; const h = setInterval(load, 1500); return () => clearInterval(h); }, [busy]);
  const install = async (id) => { try { await post('/office/components/install', { id }); load(); } catch (e) { toast.error(e); } };
  const remove = async (c) => { if (!(await confirmDialog({ title: 'Remove this component?', text: 'It is downloaded again when a file needs it.', confirm: 'Remove' }))) return; try { await del(`/office/components/${encodeURIComponent(c.id)}`); load(); } catch (e) { toast.error(e); } };
  const main = (list || []).filter((c) => !/^ocr-/.test(c.id));
  const langs = (list || []).filter((c) => /^ocr-/.test(c.id));
  const row = (c) => (
    <SettingRow key={c.id} title={compLabel(c.label)} text={<>{c.state === 'ready' ? `${t('Installed')}${c.size ? ` · ${fmtBytes(c.size)}` : ''}` : t('Download {size}', { size: fmtBytes(c.download) })}{c.error ? <span className="c-err"> · {c.error}</span> : null}</>}>
      <div className="of-set-btns">
        {c.state === 'installing' ? <div className="of-set-progress"><Progress value={c.progress || 0} /><span className="tiny muted">{`${c.progress || 0} %`}</span></div>
          : c.state === 'ready' ? <><span className="badge ok">{t('Ready')}</span><Btn className="sm ghost icon" icon={Trash2} title={t('Remove')} aria-label={t('Remove')} onClick={() => remove(c)} /></>
            : <Btn className="sm" icon={Download} onClick={() => install(c.id)}>Install</Btn>}
      </div>
    </SettingRow>
  );
  if (error) return <Callout kind="err" dismissible={false}>{error}</Callout>;
  if (!list) return <Loading inline />;
  return <>
    <Card icon={Package} title="Downloaded when needed" tip="Large parts that do not fit into the extension. Office downloads them the first time a file needs them – from fixed addresses, every file checked against a fixed checksum. You can also install them ahead, e.g. before going offline.">
      {main.map(row)}
    </Card>
    {langs.length > 0 && <Card icon={Type} title="More languages for text recognition">{langs.map(row)}</Card>}
  </>;
}
