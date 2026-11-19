/*
 * Office settings – every field of the settings object (ARCHITECTURE.md), grouped like Hostwerk's own settings
 * (SettingsLayout left, SettingRow + Switch/Seg/fields right, one SaveBar). Shown as #/ext/office/settings and as the
 * tab "Office" in Hostwerk's settings (slot settings.tabs, prop `embedded`).
 */
import { createContext, useContext, useEffect, useMemo, useState } from 'react';
import {
  Settings2, FileText, Sheet, Presentation, FileType, Cpu, Plug, Share2, Library, ArrowLeft, CheckCircle2, XCircle, Download, RotateCcw, Trash2, FlaskConical, Sparkles, Cloud, LogOut, RefreshCw, MonitorCheck,
} from 'lucide-react';
import { get, post } from '@hostwerk/api';
import { SettingsLayout, SettingRow, Switch, Seg, Field, PathInput, SaveBar, Callout, Btn, Loading, Empty, toast, confirmDialog, useRemembered, InfoTip } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { office, download } from '../lib/api';
import { DEFAULT_SETTINGS, getIn, setIn, patchOf, diffSettings, validateSettings, withDefaults } from './settingsModel';
import { useOfficeSettings, useOfficeStatus, saveSettings, loadStatus } from './store';
import { parseQuery, fmtBytes } from './util';
import { useEngine, engineReady, engineBusy, installEngine, updateEngine, removeEngine, EngineProgress, EngineDialogHost, useIntegrations, FullEditorSetup, loadIntegrations, FullEditorDialogHost } from './engine';
import { LIB_ROOT, filesRoute } from './ops';
import { DefaultAppCard } from './osCard';
import { openLink, fmtDateTime, Keep } from './common';

const GROUPS = [
  { id: 'general', label: 'General', icon: Settings2 },
  { id: 'library', label: 'Folder & storage', icon: Library },
  { id: 'docs', label: 'Documents', icon: FileText },
  { id: 'sheets', label: 'Spreadsheets', icon: Sheet },
  { id: 'slides', label: 'Presentations', icon: Presentation },
  { id: 'pdf', label: 'PDF', icon: FileType },
  { id: 'converters', label: 'Converters & apps', icon: Cpu },
  { id: 'integrations', label: 'Integrations', icon: Plug },
  { id: 'sharing', label: 'Sharing', icon: Share2 },
];
const FONTS = ['Calibri', 'Arial', 'Aptos', 'Cambria', 'Carlito', 'Georgia', 'Helvetica', 'Liberation Sans', 'Liberation Serif', 'Open Sans', 'Roboto', 'Segoe UI', 'Times New Roman', 'Verdana'];
const LANGS = [['auto', 'Like Hostwerk'], ['de-DE', 'German (Germany)'], ['de-AT', 'German (Austria)'], ['de-CH', 'German (Switzerland)'], ['en-US', 'English (US)'], ['en-GB', 'English (UK)'], ['fr-FR', 'French'], ['es-ES', 'Spanish'], ['it-IT', 'Italian'], ['nl-NL', 'Dutch'], ['pl-PL', 'Polish'], ['pt-PT', 'Portuguese']];
const ZOOMS = [50, 75, 90, 100, 110, 125, 150, 175, 200];
const RowCtx = createContext({ v: () => undefined, set: () => () => {}, err: () => null });
/* ---- rows (module level, so inputs keep their focus while typing) ---- */
function Toggle({ p, title, text, disabled }) {
  const { v, set, err } = useContext(RowCtx);
  return <SettingRow title={title} text={text}><Switch checked={!!v(p)} disabled={disabled} onChange={set(p)} /></SettingRow>;
}
function Num({ p, title, text, unit, step = 1, disabled }) {
  const { v, set, err } = useContext(RowCtx);
  return (
    <SettingRow title={title} text={text}>
      <div className="stack of-shell-gap-0 of-shell-num"><div className="row-tight"><input className="input of-shell-num-in" type="number" step={step} disabled={disabled} value={v(p) ?? ''} onChange={(e) => set(p)(e.target.value === '' ? '' : Number(e.target.value))} />{unit && <span className="small muted">{t(unit)}</span>}</div>{err(p)}</div>
    </SettingRow>
  );
}
function Pick({ p, title, text, options }) {
  const { v, set, err } = useContext(RowCtx);
  return <SettingRow title={title} text={text}><Seg value={String(v(p))} onChange={(x) => set(p)(typeof getIn(DEFAULT_SETTINGS, p) === 'number' ? Number(x) : x)} options={options.map(([id, label]) => ({ id: String(id), label }))} /></SettingRow>;
}
function Select({ p, title, text, options, num = false }) {
  const { v, set, err } = useContext(RowCtx);
  return (
    <SettingRow title={title} text={text}>
      <select className="input select of-shell-select" value={String(v(p))} onChange={(e) => set(p)(num && e.target.value !== 'auto' && !Number.isNaN(Number(e.target.value)) ? Number(e.target.value) : e.target.value)}>
        {options.map(([id, label]) => <option key={id} value={String(id)}>{t(label)}</option>)}
      </select>
    </SettingRow>
  );
}
function Text({ p, title, text, placeholder, type = 'text', list }) {
  const { v, set, err } = useContext(RowCtx);
  return (
    <SettingRow title={title} text={text}><div className="stack of-shell-gap-0"><input className="input of-shell-text" type={type} value={v(p) || ''} placeholder={placeholder ? t(placeholder) : ''} list={list} autoComplete="off" onChange={(e) => set(p)(e.target.value)} />{err(p)}</div></SettingRow>
  );
}
function Font({ p, title }) {
  const { v, set, err } = useContext(RowCtx);
  return <><Text p={p} title={title} list="of-shell-fonts" /></>;
}



export default function OfficeSettings({ embedded = false, query = '' }) {
  const { settings, error, reload } = useOfficeSettings();
  const [group, setGroup] = useRemembered('office:settings-group', 'general');
  const [draft, setDraft] = useState(null);
  const [busy, setBusy] = useState(false);
  useEffect(() => { const g = parseQuery(query).group; if (g && GROUPS.some((x) => x.id === g)) setGroup(g); }, [query]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (settings && !draft) setDraft(settings); }, [settings]); // eslint-disable-line react-hooks/exhaustive-deps
  const errors = useMemo(() => (draft ? validateSettings(draft) : {}), [draft]);
  if (error && !settings) return <Empty icon={XCircle} title="The settings could not be loaded" text={error} action={<Btn icon={RotateCcw} onClick={reload}>Try again</Btn>} />;
  if (!settings || !draft) return <Loading inline />;
  const dirty = diffSettings(settings, draft).length > 0;
  const v = (p) => getIn(draft, p);
  const set = (p) => (val) => setDraft((d) => setIn(d, p, val));
  const err = (p) => (errors[p] ? <span className="hint c-err">{t(errors[p].text, errors[p].vars)}</span> : null);
  const save = async () => {
    if (Object.keys(errors).length) { toast.error(t('Please check the marked fields.')); return; }
    setBusy(true);
    try { const s = await saveSettings(patchOf(settings, draft)); setDraft(s); toast.success('Saved'); loadStatus(true).catch(() => {}); } catch (e) { toast.error(e); }
    setBusy(false);
  };
  const badge = (id) => { const n = Object.keys(errors).filter((p) => p.startsWith(`${id}.`) || (id === 'library' && p === 'libraryDir')).length; return n ? { warn: n } : {}; };
  const groups = GROUPS.map((g) => ({ ...g, label: t(g.label), ...badge(g.id === 'general' || g.id === 'library' ? 'general' : g.id) }));

  let pane = null;
  if (group === 'general') pane = <>
    <div className="card"><div className="card-head"><h3>Saving</h3></div><div className="card-body">
      <Toggle p="general.autosave" title="Save automatically" text="Saves a few seconds after your last change. Every save keeps the previous content as a version." />
      <Num p="general.autosaveSeconds" title="Wait before saving" unit="seconds" disabled={!v('general.autosave')} />
      <Pick p="general.newFormat" title="Format of new files" text="Microsoft Office (docx, xlsx, pptx) or OpenDocument (odt, ods, odp) – you can always pick another one in the dialog." options={[['ms', 'Microsoft Office'], ['odf', 'OpenDocument']]} />
      <Toggle p="general.confirmDelete" title="Ask before moving to the trash" text="Deleted files can be restored from the trash either way." />
    </div></div>
    <div className="card"><div className="card-head"><h3>Writing</h3></div><div className="card-body">
      <Text p="general.author" title="Author name" text="Written into new files and comments." placeholder="e.g. Anna Berger" />
      <Toggle p="general.spellcheck" title="Check spelling" text="Marks unknown words while you type (uses the spell checker of this window)." />
      <Select p="general.language" title="Language of documents" text="For spelling, hyphenation and number formats." options={LANGS} />
      <Pick p="general.units" title="Units" options={[['cm', 'Centimetres'], ['in', 'Inches']]} />
    </div></div>
    <div className="card"><div className="card-head"><h3>In Files</h3></div><div className="card-body">
      <Toggle p="general.openFromFiles" title="Open office files in Office" text="A click on a document, spreadsheet, presentation or PDF in Files opens it in Office – in every folder. Off: Files shows its preview; right-click → “Open in Office” still works." />
      <Num p="general.recentCount" title="Recent files in the Office section" unit="files" />
    </div></div>
  </>;
  else if (group === 'library') pane = <>
    <div className="card"><div className="card-head"><h3>Folder “Office documents”</h3></div><div className="card-body stack">
      <Field label="Folder of “Office documents” in Files" hint="Empty = Hostwerk’s data folder. Files already there stay where they are; Files shows the new folder.">
        <PathInput mode="folder" title={t('Folder of “Office documents” in Files')} value={v('libraryDir')} onChange={set('libraryDir')} placeholder={t('Hostwerk data folder (default)')} />
        {err('libraryDir')}
      </Field>
    </div></div>
    <div className="card"><div className="card-head"><h3>Trash and versions</h3></div><div className="card-body">
      <Num p="general.trashDays" title="Keep deleted files" unit="days" text="Afterwards they are removed for good." />
      <Num p="general.keepVersions" title="Versions per file" unit="versions" text="Older versions beyond this number are removed." />
      <Num p="general.versionDays" title="Keep versions" unit="days" />
    </div></div>
  </>;
  else if (group === 'docs') pane = <>
    <div className="card"><div className="card-head"><h3>New documents</h3></div><div className="card-body">
      <Font p="docs.font" title="Font" />
      <Num p="docs.fontSize" title="Font size" unit="pt" step={0.5} />
      <Select p="docs.lineHeight" title="Line spacing" num options={[[1, 'Single (1.0)'], [1.08, '1.08'], [1.15, '1.15'], [1.5, '1.5'], [2, 'Double (2.0)']]} />
      <Num p="docs.paragraphSpacing" title="Space after paragraphs" unit="pt" />
      <Pick p="docs.pageSize" title="Paper size" options={[['A4', 'A4'], ['A5', 'A5'], ['Letter', 'Letter'], ['Legal', 'Legal']]} />
      <Pick p="docs.orientation" title="Orientation" options={[['portrait', 'Portrait'], ['landscape', 'Landscape']]} />
      <SettingRow title="Margins" text="In millimetres – top, right, bottom, left.">
        <div className="of-shell-margins">{['top', 'right', 'bottom', 'left'].map((k) => <label key={k} className="stack of-shell-gap-0"><span className="tiny muted">{t(k[0].toUpperCase() + k.slice(1))}</span><input className="input of-shell-num-in" type="number" value={v(`docs.margins.${k}`) ?? ''} onChange={(e) => set(`docs.margins.${k}`)(e.target.value === '' ? '' : Number(e.target.value))} aria-label={t(`Margin ${k}`)} />{err(`docs.margins.${k}`)}</label>)}</div>
      </SettingRow>
    </div></div>
    <div className="card"><div className="card-head"><h3>Editing and view</h3></div><div className="card-body">
      <Toggle p="docs.ruler" title="Show the ruler" />
      <Select p="docs.zoom" title="Zoom" num options={ZOOMS.map((z) => [z, `${z} %`])} />
      <Toggle p="docs.smartQuotes" title="Typographic quotes" text="Turns straight quotes into „…“ or “…” while typing." />
      <Toggle p="docs.autoLinks" title="Recognise links" text="Web and e-mail addresses become links." />
      <Toggle p="docs.showFormatting" title="Show formatting marks" text="Paragraph marks, spaces and tabs." />
    </div></div>
  </>;
  else if (group === 'sheets') pane = <>
    <div className="card"><div className="card-head"><h3>New spreadsheets and view</h3></div><div className="card-body">
      <Font p="sheets.font" title="Font" />
      <Num p="sheets.fontSize" title="Font size" unit="pt" step={0.5} />
      <Toggle p="sheets.gridlines" title="Show gridlines" />
      <Toggle p="sheets.formulaBar" title="Show the formula bar" />
      <Toggle p="sheets.headings" title="Show row and column headings" />
      <Select p="sheets.zoom" title="Zoom" num options={ZOOMS.map((z) => [z, `${z} %`])} />
    </div></div>
    <div className="card"><div className="card-head"><h3>Calculation and formats</h3></div><div className="card-body">
      <Toggle p="sheets.autoCalc" title="Calculate automatically" text="Off: F9 calculates." />
      <Toggle p="sheets.r1c1" title="R1C1 references" text="Rows and columns as numbers instead of A1." />
      <Pick p="sheets.functionNames" title="Function names" text="Files always keep the English names." options={[['auto', 'Like the language'], ['en', 'English (SUM, IF)'], ['de', 'German (SUMME, WENN)']]} />
      <Pick p="sheets.decimal" title="Decimal separator" options={[['auto', 'Like the language'], ['.', 'Point (1.5)'], [',', 'Comma (1,5)']]} />
      <Select p="sheets.dateFormat" title="Date format" options={[['auto', 'Like the language'], ['DD.MM.YYYY', '31.12.2026'], ['YYYY-MM-DD', '2026-12-31'], ['MM/DD/YYYY', '12/31/2026'], ['DD/MM/YYYY', '31/12/2026']]} />
    </div></div>
    <div className="card"><div className="card-head"><h3>CSV files</h3></div><div className="card-body">
      <Select p="sheets.csvDelimiter" title="Separator" options={[['auto', 'Detect automatically'], [',', 'Comma (,)'], [';', 'Semicolon (;)'], ['\t', 'Tab'], ['|', 'Vertical bar (|)']]} />
      <Select p="sheets.csvEncoding" title="Text encoding" options={[['utf-8', 'UTF-8 (recommended)'], ['windows-1252', 'Windows-1252 (old Excel)'], ['iso-8859-1', 'ISO-8859-1'], ['utf-16le', 'UTF-16']]} />
    </div></div>
  </>;
  else if (group === 'slides') pane = <>
    <div className="card"><div className="card-head"><h3>New presentations</h3></div><div className="card-body">
      <Pick p="slides.size" title="Slide size" options={[['16:9', '16:9'], ['4:3', '4:3'], ['A4', 'A4']]} />
      <Select p="slides.theme" title="Design" options={[['hostwerk', 'Hostwerk'], ['light', 'Light'], ['dark', 'Dark'], ['minimal', 'Minimal'], ['classic', 'Classic']]} />
      <Font p="slides.font" title="Font" />
    </div></div>
    <div className="card"><div className="card-head"><h3>Presenting</h3></div><div className="card-body">
      <Toggle p="slides.notes" title="Show speaker notes while editing" />
      <Toggle p="slides.presenterView" title="Presenter view" text="Notes, next slide and timer in a second Hostwerk window." />
      <Pick p="slides.transition" title="Transition" options={[['none', 'None'], ['fade', 'Fade'], ['slide', 'Slide in']]} />
      <Toggle p="slides.loop" title="Start again after the last slide" />
      <Num p="slides.autoAdvance" title="Next slide automatically" unit="seconds (0 = off)" />
    </div></div>
  </>;
  else if (group === 'pdf') pane = (
    <div className="card"><div className="card-head"><h3>PDF</h3></div><div className="card-body">
      <Select p="pdf.zoom" title="Zoom when opening" options={[['auto', 'Automatic'], ['page-fit', 'Whole page'], ['page-width', 'Page width'], ...ZOOMS.map((z) => [String(z), `${z} %`])]} />
      <Toggle p="pdf.sidebar" title="Show page thumbnails" />
      <Pick p="pdf.scroll" title="Scrolling" options={[['vertical', 'Vertical'], ['horizontal', 'Horizontal'], ['page', 'Page by page']]} />
      <Text p="pdf.signatureName" title="Name for signatures" text="Suggested when you sign a PDF." placeholder="e.g. Anna Berger" />
    </div></div>
  );
  else if (group === 'converters') pane = <ConvertersPane v={v} set={set} err={err} Toggle={Toggle} Num={Num} Pick={Pick} dirty={dirty} />;
  else if (group === 'integrations') pane = <IntegrationsPane v={v} set={set} err={err} Pick={Pick} dirty={dirty} save={save} />;
  else if (group === 'sharing') pane = (
    <div className="card"><div className="card-head"><h3>Share links</h3></div><div className="card-body">
      <Toggle p="sharing.enabled" title="Allow share links" text="Read-only links for single files. Visitors see the file – PDF and pictures directly, office files as PDF." />
      <Num p="sharing.defaultDays" title="Links are valid for" unit="days (0 = no end)" disabled={!v('sharing.enabled')} />
      <Toggle p="sharing.allowDownload" title="Allow download by default" disabled={!v('sharing.enabled')} />
      <div className="row-tight of-shell-pad"><a className="btn sm" href={filesRoute({ root: LIB_ROOT, dir: true })} onClick={() => { try { localStorage.setItem('office:section-view', 'shared'); } catch { /* ignore */ } }}><Share2 size={14} />{t('Show shared files')}</a></div>
    </div></div>
  );

  const layout = (
    <>
      <datalist id="of-shell-fonts">{FONTS.map((f) => <option key={f} value={f} />)}</datalist>
      <RowCtx.Provider value={{ v, set, err }}>
      <DefaultAppCard onSetting={(on) => setDraft((d) => setIn(d, 'integrations.osOpenWith', on))} />
      <SettingsLayout groups={groups} value={group} onChange={setGroup} label={t('Office settings')}>
        <div className="stack of-shell-settings">{pane}
          <div className="row-tight"><Btn className="ghost sm" icon={RotateCcw} onClick={async () => { if (await confirmDialog({ title: 'Reset this group?', text: 'The values of this group go back to the defaults. Nothing is saved until you press “Save”.', confirm: 'Reset' })) { const keys = group === 'library' ? ['libraryDir', 'general.trashDays', 'general.keepVersions', 'general.versionDays'] : [group]; setDraft((d) => keys.reduce((acc, k) => setIn(acc, k, getIn(DEFAULT_SETTINGS, k)), d)); } }}>Reset to defaults</Btn></div>
        </div>
      </SettingsLayout>
      </RowCtx.Provider>
      <SaveBar dirty={dirty} busy={busy} onSave={save} onReset={() => setDraft(settings)} />
      <EngineDialogHost />
      <FullEditorDialogHost />
    </>
  );
  if (embedded) return <div className="of-shell-settings-embedded">{layout}</div>;
  return (
    <div className="page fade-in">
      <div className="page-head">
        <div><h1><FileText size={22} /> Office settings</h1><p>How Office creates, saves, converts and shares your files.</p></div>
        <div className="actions"><a className="btn ghost" href={filesRoute({ root: LIB_ROOT, dir: true })}><ArrowLeft size={16} />{t('Office documents')}</a></div>
      </div>
      {layout}
    </div>
  );
}

/* ---------------- Converters & apps ---------------- */
function ConvertersPane({ v, set, err, Toggle, Num, Pick, dirty }) {
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
  return <>
    <div className="card"><div className="card-head"><Cpu size={16} /><h3>Office engine</h3>{engineReady(e) ? <span className="badge ok">Ready</span> : engineBusy(e) ? <span className="badge accent">Setting up …</span> : <span className="badge">Not set up</span>}</div>
      <div className="card-body stack">
        <p className="small muted of-shell-p">A private LibreOffice that only Office uses – opens old and rare formats exactly and converts between all of them. No system installation, security updates come automatically.</p>
        {e && engineBusy(e) && <EngineProgress e={e} />}
        {e && engineReady(e) && <div className="small">{t('Version {v}', { v: e.version || '?' })}{e.size ? ` · ${fmtBytes(e.size)}` : ''}{e.path ? <> · <span className="mono tiny"><Keep>{e.path}</Keep></span></> : null}{e.state === 'system' ? ` · ${t('uses the LibreOffice installed on this PC')}` : ''}</div>}
        {e && e.state === 'error' && <Callout kind="err" dismissible={false} title="Setup failed">{e.error}</Callout>}
        <div className="row-tight of-shell-wrap">
          {!engineReady(e) && <Btn className="primary sm" icon={Download} disabled={engineBusy(e)} onClick={() => installEngine()}>{t('Set up the Office engine (≈ 350 MB)')}</Btn>}
          {e && e.state === 'ready' && <Btn className="sm" icon={RefreshCw} onClick={() => updateEngine()}>Look for updates</Btn>}
          {e && e.state === 'ready' && <Btn className="sm ghost" icon={Trash2} onClick={async () => { if (await confirmDialog({ title: 'Remove the Office engine?', text: 'Old formats then open with the built-in readers only. You can set it up again at any time.', confirm: 'Remove', danger: true })) await removeEngine(); }}>Remove</Btn>}
        </div>
        <Pick p="converters.autoEngine" title="When a file needs the engine" options={[['ask', 'Ask once'], ['auto', 'Set it up automatically'], ['off', 'Never']]} />
      </div>
    </div>
    <div className="card"><div className="card-head"><h3>Found on this PC</h3><div className="right"><Btn className="sm ghost" icon={RefreshCw} onClick={reload}>Check again</Btn></div></div>
      <div className="card-body stack">
        {!status ? <Loading inline /> : <>
          <div className="row-tight">{lo ? <CheckCircle2 size={16} className="c-ok" aria-hidden="true" /> : <XCircle size={16} className="muted" aria-hidden="true" />}<b className="small">LibreOffice</b><span className="small muted ellipsis">{lo ? `${lo.version || ''} · ${lo.path || ''}` : t('not found')}</span></div>
          <div className="row-tight">{msApps.length ? <CheckCircle2 size={16} className="c-ok" aria-hidden="true" /> : <XCircle size={16} className="muted" aria-hidden="true" />}<b className="small">Microsoft Office</b><span className="small muted">{msApps.length ? msApps.map((k) => (typeof ms[k] === 'string' ? `${k[0].toUpperCase()}${k.slice(1)} ${ms[k]}` : k[0].toUpperCase() + k.slice(1))).join(' · ') : t('not found')}</span></div>
        </>}
        <div className="row-tight of-shell-wrap"><Btn className="sm" icon={FlaskConical} disabled={dirty} title={dirty ? t('Save first') : undefined} onClick={runTest}>Test conversion</Btn>
          {test && (test.ok ? <span className="small c-ok"><CheckCircle2 size={13} aria-hidden="true" /> {t('Works – PDF with {size} in {ms} ms', { size: fmtBytes(test.size), ms: test.ms })} <button type="button" className="link-btn small" onClick={() => download(test.out, 'office-test.pdf', 'application/pdf')}>{t('Download')}</button></span> : <span className="small c-err"><XCircle size={13} aria-hidden="true" /> {test.error}</span>)}
        </div>
      </div>
    </div>
    <div className="card"><div className="card-head"><h3>Which converter</h3></div><div className="card-body">
      <Pick p="converters.preferred" title="Prefer" text="Microsoft Office converts its own formats most faithfully; LibreOffice (or the Office engine) knows more formats." options={[['auto', 'Automatic'], ['libreoffice', 'LibreOffice'], ['msoffice', 'Microsoft Office']]} />
      <Toggle p="converters.msoffice" title="Use Microsoft Office when installed" text="Windows only – Office runs invisibly in the background." />
      <SettingRow title="Path to LibreOffice" text="Empty = find it automatically (or use the Office engine).">
        <div className="of-shell-path-in"><PathInput mode="file" value={v('converters.libreoffice')} onChange={set('converters.libreoffice')} placeholder={t('automatic')} /></div>
      </SettingRow>
      <Num p="converters.timeout" title="Give up after" unit="seconds" text="For very large files." />
    </div></div>
  </>;
}

/* ---------------- Integrations ---------------- */
function IntegrationsPane({ v, set, err, Pick, dirty, save }) {
  const st = useIntegrations();
  const { status, reload } = useOfficeStatus();
  const [g, setG] = useState(null);
  const loadG = () => get('/office/google/state').then(setG, () => setG({ unavailable: true }));
  useEffect(() => { loadG(); }, []);
  const test = async () => {
    if (dirty) await save();
    const s = await loadStatus(true);
    const r = s && s.integrations && s.integrations.onlyoffice;
    if (r && (r.ok || r.version)) toast.success(t('{app} answers{v}', { app: 'ONLYOFFICE', v: r.version ? ` – ${t('version {v}', { v: r.version })}` : '' }));
    else toast.error(t('{app} does not answer: {error}', { app: 'ONLYOFFICE', error: (r && r.error) || t('no answer') }));
    reload();
  };
  const connectGoogle = async () => {
    if (dirty) await save(); // a client id typed just now must reach the backend first
    // desktop app: the backend runs the sign-in in a Hostwerk window and answers with the new state; else { url }
    const r = await post('/office/google/connect');
    if (r && r.connected) { setG(r); toast.success('Google Drive connected'); return; }
    if (r && r.url) await openLink(r.url); // sign-in in a Hostwerk window
    const h = setInterval(async () => { const s = await get('/office/google/state').catch(() => null); if (s && s.connected) { clearInterval(h); setG(s); toast.success('Google Drive connected'); } }, 3000);
    setTimeout(() => clearInterval(h), 300000);
  };
  return <>
    <div className="card"><div className="card-head"><Sparkles size={16} /><h3>Full editor – one click</h3><InfoTip title="Full editor">Hostwerk starts ONLYOFFICE Docs in Docker and connects it. Then “Open in the full editor” edits every file 100 % compatible inside Hostwerk – track changes, SmartArt, pivot tables, macros kept.</InfoTip></div>
      <div className="card-body stack">
        {st && st.docker === false && <Callout kind="warn" dismissible={false} title="Docker is needed">{st.dockerHint || t('Set up Docker under Projects & apps → Docker, then come back here.')}</Callout>}
        <FullEditorSetup st={st} docker={!(st && st.docker === false)}>{st && st.onlyoffice && <span className="tiny muted">{t('Runs in Docker{url}', { url: st.onlyoffice.url ? ` · ${st.onlyoffice.url}` : '' })}</span>}</FullEditorSetup>
        <Pick p="integrations.defaultOpen" title="A click opens files in" options={[['builtin', 'Built-in editor'], ['onlyoffice', 'Full editor (ONLYOFFICE)'], ['desktop', 'Desktop app']]} />
      </div>
    </div>
    <div className="card"><div className="card-head"><MonitorCheck size={16} /><h3>Windows and Linux</h3></div><div className="card-body">
      <Toggle p="integrations.osOpenWith" title="Show Hostwerk under “Open with”" text="Hostwerk registers itself for Word, Excel, PowerPoint, OpenDocument, PDF and the other formats Office opens – right-click a file in Explorer → “Open with” → Hostwerk Office. Your current apps stay the default until you choose otherwise (the card at the top). Off: Hostwerk removes the entries again. Only in the desktop app, per Windows user, no admin rights." />
    </div></div>
    <div className="card"><div className="card-head"><Cloud size={16} /><h3>Google Drive</h3>{g && g.connected && <span className="badge ok">Connected</span>}</div>
      <div className="card-body stack">
        <p className="small muted of-shell-p">Open Google Docs, Sheets and Slides (.gdoc, .gsheet, .gslides) right in Office’s editors – saving goes back to Google.</p>
        {!g ? <Loading inline /> : g.unavailable ? <span className="small muted">{t('Not available in this version of the Office backend yet.')}</span> : g.connected ? (
          <div className="row-tight of-shell-wrap"><span className="small">{t('Signed in as {account}', { account: (g.account && typeof g.account === 'object' ? g.account.email || g.account.name : g.account) || '?' })}</span><Btn className="sm ghost" icon={LogOut} onClick={async () => { await post('/office/google/disconnect'); loadG(); }}>Disconnect</Btn></div>
        ) : <div className="row-tight"><Btn className="sm primary" icon={Cloud} disabled={!g.configured && !String(v('integrations.google.clientId') || '').trim()} title={!g.configured ? t('Enter the client ID of your own Google OAuth client first') : undefined} onClick={() => connectGoogle().catch((er) => toast.error(er))}>Connect Google Drive …</Btn><span className="tiny muted">{t('You sign in in a Hostwerk window.')}</span></div>}
        {g && !g.unavailable && !g.connected && <>
          <SettingRow title="Google OAuth client ID" text="Your own OAuth client of the type “Desktop app” – Hostwerk ships none."><div className="stack of-shell-gap-0"><input className="input of-shell-text" value={v('integrations.google.clientId') || ''} placeholder="…apps.googleusercontent.com" onChange={(e) => set('integrations.google.clientId')(e.target.value.trim())} />{err('integrations.google.clientId')}</div></SettingRow>
          <SettingRow title="Google OAuth client secret"><input className="input of-shell-text" type="password" autoComplete="new-password" value={v('integrations.google.clientSecret') || ''} onChange={(e) => set('integrations.google.clientSecret')(e.target.value.trim())} /></SettingRow>
          {Array.isArray(g.guide) && g.guide.length > 0 && <details className="small"><summary>{t('How to create the OAuth client')}</summary><ol className="of-shell-p">{g.guide.map((x, i) => <li key={i}>{x}</li>)}</ol>{g.redirectUri && <p className="tiny muted">{t('Redirect address')}: <span className="mono">{g.redirectUri}</span></p>}</details>}
        </>}
      </div>
    </div>
    <div className="card"><div className="card-head"><Plug size={16} /><h3>Own ONLYOFFICE server</h3></div><div className="card-body">
      <SettingRow title="ONLYOFFICE Document Server address"><div className="stack of-shell-gap-0"><div className="row-tight"><input className="input of-shell-text" value={v('integrations.onlyoffice.url') || ''} placeholder="https://…" onChange={(e) => set('integrations.onlyoffice.url')(e.target.value)} /><Btn className="sm ghost" disabled={!v('integrations.onlyoffice.url')} onClick={test}>Test</Btn></div>{err('integrations.onlyoffice.url')}</div></SettingRow>
      <SettingRow title="ONLYOFFICE secret" text="The JWT secret of the Document Server (stays on this PC)."><input className="input of-shell-text" type="password" autoComplete="new-password" value={v('integrations.onlyoffice.secret') || ''} onChange={(e) => set('integrations.onlyoffice.secret')(e.target.value)} /></SettingRow>
    </div></div>
  </>;
}
