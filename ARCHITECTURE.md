# Office extension – architecture contract

Native Hostwerk extension (official package, SDK 3.x). Several people build it in parallel – **this file is the contract
between the parts. Stick to it; the decisions taken while putting the parts together are under "Integration" at the end.**

Reference material (read before you start):
- SDK guide: `../../hostwerk/docs/extensions/README.de.md` (sections 4, 5, 8, 15, 16, 21, 23–27, 32), UI rules:
  `../../hostwerk/docs/ui-guidelines.md`, Hostwerk UI components: `../../hostwerk/ui/src/components/ui.jsx`
- Native host API: `../../hostwerk/server/services/extnative.js` (makeHost), core files: `../../hostwerk/server/core/files.js`
- A good native example: `../hostwerk-ext-file-share` (server/ + ui/src/), bigger ones: `../hostwerk-ext-vms`, `../hostwerk-ext-passwords`

## Layout and ownership

| Path | Owner | What |
|---|---|---|
| `manifest.json`, `locales/`, `README*`, `CHANGELOG.md`, `tools/`, `ARCHITECTURE.md` | integrator | merged at the end |
| `server/**`, `tests/server-*.test.js` | backend | API, library, versions, trash, search, converters, sharing, ONLYOFFICE, AI tools |
| `ui/src/index.jsx`, `ui/src/shell/**`, `ui/src/style.css` | shell | Files integration (files.jsx, osCard.jsx, slots.jsx), editor host, settings, dialogs |
| `ui/src/lib/**` | integrator (read-only for others) | `api.js`, `formats.js`, `loader.js` – add, never break |
| `ui/src/editors/docs/**` | docs | text documents |
| `ui/src/editors/sheets/**` | sheets | spreadsheets |
| `ui/src/editors/slides/**` | slides | presentations |
| `ui/src/editors/pdf/**` | pdf | PDF viewer/editor |
| `ui/src/editors/viewer/**` | viewer | everything read-only (mail, vcf, ics, epub, Google links, Visio/… via converter, images, zip …) |
| `ui/src/editors/_shared/**` | anyone may ADD a new file; never edit someone else's | shared helpers (zip/xml utilities …) |
| `tests/<area>-*.test.js` | each area | `node --test` unit tests of pure logic (no browser) |
| `locales/de.json` | everyone (add, keep the wording consistent) | `{ "English text": "Deutscher Text" }` – the parts' files are merged into it |

Never touch files of another area. Do not run `npm install` (all libraries are installed – see `package.json`
devDependencies: tiptap 2, mammoth, docx, exceljs, xlsx (SheetJS CE 0.20.3), pptxgenjs, pdfjs-dist 4, pdf-lib, jszip,
marked, turndown(+gfm), dompurify, @formulajs/formulajs, @kenjiuno/msgreader, lucide-react). If something is truly
missing, write it under "Known limits".

## Build

`node tools/build.js` → `ui/index.js` (main module from `ui/src/index.jsx`, via the SDK's build-native) and one lazy
chunk per editor: `ui/src/editors/<kind>/index.jsx` → `ui/assets/<kind>.js` (+ `.css` if the chunk imports CSS).
`node tools/build.js <kind>` builds one chunk. Chunks can import `react`, `@hostwerk/api|ui|kit|i18n` (shared with
Hostwerk, not bundled) and any installed npm package (bundled). **Rules:** no `eval(` anywhere in the output
(catalog check fails – if a library has it, find another way), the **shipped package** ≤ 10 MB (Office budget 9.8 MB –
the release zip's content, see "Integration 1.1": ui/src, tests and tools are not shipped), no remote scripts/fonts,
everything works offline (heavy optional parts are components on demand). `node tools/build.js --sizes` prints the chunk
sizes and the shipped size, `--analyze` the packages bundled into more than one chunk; `node tools/package.js --dir|--zip`
builds the release package exactly like the CI. Prefer dynamic-free code; if a library needs a Web Worker (pdf.js), create
the worker from a Blob of a bundled string or run it without a worker.

## Design rules (all UI)

Hostwerk look: use Hostwerk's classes (`card`, `btn`, `btn primary|ghost|sm|icon`, `input`, `stack`, `row`, `grid`,
`muted`, `small`, `mono`, `pill` …) and components from `@hostwerk/ui` (`Btn, Modal, SideSheet, Tabs, Seg, Field,
Switch, Callout, Empty, Loading, Spinner, Progress, DropZone, InfoTip, SettingRow, SettingsLayout, toast,
confirmDialog, promptDialog, MasterDetail, Accordion …`), icons from `lucide-react`. Own CSS only for what the kit
lacks, **prefixed with the area** (`.of-docs-…`, `.of-sheet-…`, `.of-slides-…`, `.of-pdf-…`, `.of-view-…`,
`.of-shell-…`) and built from Hostwerk's tokens (`var(--sp-1..8)`, `var(--fs-xs..)`, `var(--r-sm|md|lg|pill)`,
`var(--line)`, `var(--muted)`, `var(--text)`, `var(--bg)`, `var(--panel)`, `var(--accent)`, `var(--ok|warn|err)`) – check the
real token names in `../../hostwerk/ui/src/styles.css`. Light and dark theme must both work; the *page* of a document
(paper) stays white with black text in both themes. Animations stop under `prefers-reduced-motion`. Every button has a
text or `aria-label` + `title`. Keyboard: all common shortcuts of the office apps (Ctrl+S/Z/Y/B/I/U/F/H/P …) inside
the editor area; never steal Hostwerk's global ones (Ctrl+K, Ctrl+J, Ctrl+1…9, F1).

**Texts:** write English in the code (JSX text is translated automatically, strings outside JSX through
`t('…', { vars })` from `@hostwerk/i18n`); put every German translation into `locales/de.json`. German
wording: short, friendly, "du"-form, like Hostwerk ("Speichern", "Als PDF herunterladen", "Rückgängig").

Comfort is a requirement: sensible defaults, everything reachable in ≤ 2 clicks, tooltips with shortcuts, empty
states with a next step, no dead ends (a format we cannot edit always offers: view, convert to an editable copy,
open in the desktop app, download).

## Editor chunk contract (`ui/src/editors/<kind>/index.jsx`, default export)

```js
export default {
  kind: 'docs',                                  // docs | sheets | slides | pdf | viewer
  Editor,                                        // React component, props below
  create: async ({ ext, template, settings }) => Uint8Array,   // a new file (blank or a template id from formats.js TEMPLATES)
  toText: async (bytes, ext) => string,          // optional: plain text (preview, AI, search fallback)
  Preview,                                       // optional: small read-only component ({ bytes, ext, name }) for Hostwerk's file browser
};
```

Editor props (given by the shell):

| prop | meaning |
|---|---|
| `file` | `{ name, path, root?, ext, size, mtime }` – `root` set = file from Hostwerk's file browser |
| `bytes` | `Uint8Array` of the file as stored (for `edit: 'convert'` formats the shell already converted to `via` and sets `file.converted = { from, to }`) |
| `readOnly` | bool – show everything, allow nothing that changes the file (still: zoom, search, copy, print, export) |
| `settings` | full settings object (see below); read your own section (`settings.docs` …) and `settings.general` |
| `host.setSaver(fn)` | register `async () => Uint8Array` returning the file in **`file.ext`** format (or the `via` format for converted files – the shell then converts back or asks) |
| `host.setExporter({ formats: [{ ext, label }], run: async (ext) => Uint8Array })` | everything the chunk can write itself (see `EXPORTS` in formats.js); the shell adds server-converter formats |
| `host.dirty(bool)` | tell the shell about unsaved changes (shell shows "Saving…/Saved", autosave, leave guard) |
| `host.save()` | ask the shell to save now (your Ctrl+S calls this) |
| `host.convert(bytes, fromExt, toExt)` | backend converter (LibreOffice / MS Office); throws with a clear message if none is installed |
| `host.hasConverter` | bool |
| `host.download(bytes, name, mime)` | save a file to the user's computer |
| `host.print(html | Blob)` | shell prints an HTML string or a PDF blob (hidden iframe) |
| `host.status(text)` | short text for the shell's status line (e.g. "Page 2 of 5 · 312 words") |
| `host.close()` | leave the editor |
| `host.lang` | `'de'` or `'en'` |
| `host.setSaver(fn, { ext })` | `ext` = format the saver returns (default `file.ext`); a saver may also return `{ bytes, ext }` |
| `host.status({ text, trackChanges, changes })` | plain string still works; shows a "Track changes · n" badge |
| `host.fidelity({ items, kept })` | ONE calm shell banner with "Open in the full editor" (instead of an own Callout) |
| `host.openLink(url)` | opens links inside Hostwerk (`openInApp`), never the external browser |
| `host.openAs({ bytes, ext, name, google? })` | show bytes in the matching editor as a new, unsaved file (see Integration) |
| `host.saveCopy(bytes, ext, name)` | save bytes as a NEW file next to the original (first free name) and go on editing that copy in the same editor → entry |
| `host.ensureEngine(reason)` | the Office engine dialog; when it is ready the file opens again through the engine → ready |

The shell renders the top bar (back, file name with rename, save state, versions, share, download as …, open in
desktop app, info, close). **The chunk renders everything below:** its ribbon/toolbar (tabs like the office apps:
File-free – Home, Insert, Layout, Review, View … as fits), the content and a status bar. The chunk must fill its
parent (`height: 100%`, own scrolling inside), never the whole window.

Fidelity: when a file contains things the chunk cannot keep when writing it back (macros, tracked changes, embedded
objects, complex layout …), show a calm, dismissible `Callout` at the top ("Some parts of this file can't be edited
here – they are kept as they are / will be lost. Open in the desktop app or a full office server for those.") and
make the saver keep the original parts where possible (round-trip unknown XML parts of the zip unchanged).

## Settings (backend stores, shell edits, chunks read)

```js
{
  libraryDir: '<abs path>',              // default: <Hostwerk data>/office
  general: { autosave: true, autosaveSeconds: 20, newFormat: 'ms' | 'odf', author: '', spellcheck: true, language: 'auto',
             confirmDelete: true, trashDays: 30, keepVersions: 30, versionDays: 90, view: 'grid' | 'list', sort: 'modified',
             openFromFiles: true, recentCount: 20, units: 'cm' | 'in' },
  docs:   { font: 'Calibri', fontSize: 11, lineHeight: 1.15, paragraphSpacing: 8, pageSize: 'A4' | 'Letter' | 'A5' | 'Legal',
            orientation: 'portrait' | 'landscape', margins: { top: 25, right: 25, bottom: 20, left: 25 } /* mm */,
            ruler: true, zoom: 100, smartQuotes: true, autoLinks: true, showFormatting: false },
  sheets: { font: 'Calibri', fontSize: 11, gridlines: true, formulaBar: true, headings: true, autoCalc: true, zoom: 100,
            decimal: 'auto' | '.' | ',', dateFormat: 'auto', csvDelimiter: 'auto', csvEncoding: 'utf-8', r1c1: false,
            functionNames: 'auto' | 'en' | 'de' /* input/display of function names, WAHR/FALSCH, #NV – files stay English */ },
  slides: { size: '16:9' | '4:3' | 'A4', theme: 'hostwerk', font: 'Calibri', notes: true, presenterView: true,
            transition: 'none' | 'fade' | 'slide', loop: false, autoAdvance: 0 },
  pdf:    { zoom: 'auto', sidebar: true, scroll: 'vertical' | 'horizontal' | 'page', signatureName: '' },
  converters: { libreoffice: '' /* path, empty = auto-detect */, msoffice: true, preferred: 'auto' | 'libreoffice' | 'msoffice', timeout: 120,
                autoEngine: 'ask' | 'auto' | 'off' },
  integrations: { onlyoffice: { url: '', secret: '' }, defaultOpen: 'builtin' | 'onlyoffice' | 'desktop', hostUrl: '' /* how ONLYOFFICE reaches Hostwerk */,
                  google: { clientId: '', clientSecret: '' } /* own OAuth client (Desktop app) */,
                  osOpenWith: true /* Hostwerk under the system's "Open with" (SDK 3.15 host.fileTypes) */ },
  sharing: { enabled: true, defaultDays: 7, allowDownload: true }
}
```

## Backend API (`host.router`, Hostwerk prefixes `/api`; JSON unless marked raw)

Files are addressed by `path` relative to the library (`/` separators, `''` = library root) **or** by `root` + `path`
of Hostwerk's file browser (resolve through Hostwerk's files core; same safety rules – never leave the root).

| Route | Meaning |
|---|---|
| `GET /office/status` | `{ version, libraryDir, converters: { libreoffice: { path, version } \| null, msoffice: { word, excel, powerpoint, visio } \| null }, integrations: { onlyoffice }, platform }` |
| `GET/PUT /office/settings` | settings above (PUT merges, validates, returns full) |
| `GET /office/files?dir=&view=all\|recent\|favorites\|shared\|trash\|templates&kind=&family=&tag=&sort=&q=` | `{ dir, entries: [Entry], crumbs }` |
| `GET /office/search?q=&dir=&kind=` | full text: `{ hits: [{ ...Entry, snippet }] }` (indexed text of docx/xlsx/pptx/odf/rtf/txt/md/html/pdf/csv) |
| `POST /office/folder {dir,name}` · `POST /office/rename {path\|root+path,name}` · `POST /office/move {paths,to}` · `POST /office/copy {paths,to}` | file ops, return Entry / entries (rename also in any root – metadata, versions, shares follow) |
| `POST /office/trash {paths}` · `POST /office/trash/restore {ids}` · `POST /office/trash/purge {ids?}` | trash (auto-purge after `trashDays`) |
| `PUT /office/meta {path, favorite?, tags?, color?, note?}` · `GET /office/tags` | metadata (kept in the data folder, follows renames/moves) |
| `GET /office/raw?path=\|root=&path=[&download=1]` | **raw** bytes; header `x-mtime` |
| `PUT /office/raw?path=…&mtime=&force=&label=` | **raw** body; 409 `{ error, mtime }` if the file changed on disk since `mtime` (unless force); keeps the previous content as a version; returns `{ ok, mtime, size, version }` |
| `POST /office/create?dir=&name=&keep=[&root=]` | **raw** body = initial bytes; unique name (never overwrites); `root` = a folder of Hostwerk's Files; returns Entry |
| `GET /office/marks?root=&path=` | marks of the office files in a folder of Files: `{ names: { name: { favorite, shared, lock, tags } } }` (files.marks) |
| `GET /office/os` · `POST /office/os/register` · `POST /office/os/unregister` · `POST /office/os/defaults` · `POST /office/os/seen` | "Open with" / default app (osreg.js over `host.fileTypes`): state per kind, register, remove, Windows' Default apps page (Linux: xdg-mime), first-run notice seen |
| `GET /office/info?path=` | Entry + `{ pages?, words?, sheets?, slides?, author?, title?, created?, modified? }` from the file's metadata |
| `GET /office/thumb?path=` | embedded thumbnail (docProps/thumbnail.*, ODF Thumbnails/thumbnail.png) or 404 |
| `GET /office/versions?path=` · `GET /office/versions/raw?path=&id=` · `POST /office/versions/restore {path,id}` · `PUT /office/versions/label {path,id,label}` | versions `{ id, time, size, label, auto }` |
| `GET /office/templates` · `POST /office/templates {path,name,category}` · `GET /office/templates/:id/raw` · `DELETE /office/templates/:id` | own templates (built-ins live in the chunks) |
| `POST /office/convert?from=&to=` | **raw** in/out via LibreOffice (`soffice --headless --convert-to`) or Microsoft Office (COM via PowerShell, Windows); 501 with a clear German/English message if no converter |
| `GET /office/zip?paths=[json]` | zip download of several files/folders |
| `POST /office/open-desktop {path\|root+path}` · `POST /office/reveal {…}` | open with the system's default app / show in Explorer |
| `POST /office/lock` · `POST /office/unlock` | soft lock while a file is open in an editor (`{ by, since }` in Entry.lock) |
| `GET /office/shares` · `POST /office/shares {path, expires, password?, allowDownload}` · `DELETE /office/shares/:id` | read-only share links (public page via `host.publicRouter`: view PDF/images/text inline, office files as PDF if a converter exists, else download) |
| `GET /office/assets/:file` | serves `ui/assets/*` (js/css, long cache, `?v=` busting) |
| `GET /office/wopi/oo/:id/file` · `POST /office/wopi/oo/:id/callback` (public, signed token) + `POST /office/integrations/session {path, app: 'onlyoffice'}` | file + save callback for the ONLYOFFICE Document Server (config + JWT; the path name `wopi` is historical) |

`Entry = { path, name, dir: bool, ext, kind, family, size, mtime, ctime, favorite, tags: [], color, note, lock, shared: bool, thumb: bool }`
(trash entries add `{ id, deletedAt, from }`; `view=recent` is ordered by last open – the backend records opens on GET /office/raw).

Body size: raw routes accept up to 200 MB (`express.raw({ type: '*/*', limit })` from `host.lib('express')`).
Libraries on the server: only `host.lib(name)` (Hostwerk's own dependencies: yauzl, unpdf, multer, yaml, ws, …) and
Node built-ins (zlib for writing zips). Native official packages run inside Hostwerk – keep them safe anyway:
validate every path, never execute anything from file contents, spawn converters with argument arrays (no shell),
time-limit them, clean temp files.

## Full fidelity – binding requirement (added by the user: "vollumfänglich")

Nothing may depend on the user installing something by hand, and nothing in a file may get lost.

1. **Office engine, provisioned automatically (backend + shell).** If neither LibreOffice nor MS Office is found, the
   backend sets up its own private LibreOffice in Hostwerk's data folder (`<data>/office-engine/`) – no system install,
   no admin rights: Windows = official MSI from download.documentfoundation.org unpacked with `msiexec /a <msi> /qn
   TARGETDIR=<dir>` (administrative install = portable), checksum-verified; Linux = official AppImage (or deb extract);
   macOS = dmg copy. Setting `converters.autoEngine: 'ask' | 'auto' | 'off'` (default `ask`: the first time a file
   needs it, the shell shows one dialog "Set up the Office engine once (≈ 350 MB)?" with progress; `auto` does it
   silently in the background after install). Routes: `GET /office/engine` (state, version, size, progress),
   `POST /office/engine/install`, `POST /office/engine/update`, `DELETE /office/engine`. Register it with
   `host.ctx.patchUpdates` (SDK 2.7 section) so security updates come automatically. While the engine is missing or
   installing, every format still opens with the built-in fallbacks below – never a dead end. The private engine is the
   only way Office provisions LibreOffice (there is no system-wide install); a LibreOffice or Microsoft Office that is
   already installed on the PC is detected and used.
2. **Full office server, one click (backend + shell).** Under Settings → Integrations: "Set up full editor" creates an
   ONLYOFFICE Document Server container (the only full editor) through Hostwerk's Docker support
   (`host.service('docker')` – look how Hostwerk / other extensions create containers), configures URL + secret, and
   from then on "Open with → Full editor" gives 100 % MS-Office-compatible editing in the browser (track changes,
   SmartArt, macros kept, pivot tables, everything). Offered prominently whenever a file contains things the built-in
   editor cannot edit.
3. **Built-in binary/legacy readers (no engine needed):** doc/dot (Word 97–2003, CFB + FIB + piece table: text,
   character/paragraph formatting, tables, lists, images where feasible), ppt/pps/pot (PowerPoint 97–2003: slides, text,
   basic shapes, pictures), xls/xlt (SheetJS already), Pages / Keynote / Numbers (Numbers via SheetJS; Pages/Keynote:
   show the embedded `QuickLook/Preview.pdf` / `preview.jpg` instantly, extract text from the IWA protobuf if feasible;
   full editing via the engine), wps/wpd/wri/abw/lwp/hwp/sxw/sxi/et/dps via the engine, with text extraction fallback.
   Shared CFB (OLE compound file) reader in `ui/src/editors/_shared/cfb.js` (docs agent writes it, slides reuses it).
4. **Nothing gets lost when saving (all editors):** save by patching the original package, not by regenerating it:
   unknown parts (vbaProject.bin, customXml, embeddings, SmartArt `diagrams/*`, charts, pivot caches, ActiveX, digital
   signatures are invalidated – warn) are written back byte-identical; elements the editor cannot edit become
   *preserved nodes* – shown with their stored fallback rendering (SmartArt drawing part, chart cache, OLE preview
   image, mc:AlternateContent fallback) and written back unchanged. **Track changes** (w:ins/w:del/w:moveFrom/To,
   formatting changes; ODF text:tracked-changes) are read, shown, recorded while editing when "Track changes" is on,
   accept/reject single/all, and written back. **Comments** incl. replies round-trip. **Macros** are kept (docm/xlsm/
   pptm stay macro-enabled; they are not executed – shown in a read-only "Macros" panel with the VBA source decoded from
   vbaProject.bin where feasible).

## Everything inside Hostwerk – binding requirement (user)

Never send the user to an external web browser. Extra windows are Hostwerk windows (`window.open('#/…?window=1')`),
external links open through `host.openLink(url)` / `openInApp(url)` (in-app web window, added to the core), office
servers are embedded, Google Drive is connected with sign-in in a Hostwerk window and Google Docs/Sheets/Slides open in
our own editors (export → edit → upload back). The office engine and full editor are set up by Hostwerk itself.

## Integration (integrator's decisions)

The parts were built in parallel against this contract; their deviations are now part of it.

**General**
- **Removed on purpose (user):** the system-wide LibreOffice install via winget (`/office/install/*`) and the Collabora
  Online integration. The private engine is the only way Office provisions LibreOffice (an installed LibreOffice or
  Microsoft Office is detected and used); ONLYOFFICE is the only full editor; a stored `defaultOpen` of the removed
  editor is migrated to `'onlyoffice'` on load.
- **Manifest:** `sdkVersion: "3.15"` (Files: files.open/files.new, host.files.addRoot, host.fileTypes; before: 3.14 openInApp /
  host.webWindow), `page.nav: false`, `hostwerk: ">=3.2.2"` – Hostwerk 3.2.1 is the current release (SDK 3.13); SDK 3.14/3.15
  are in the core's working tree and ship with the next release, which is 3.2.2 at the earliest (Hostwerk bumps the SDK in patch releases too – SDK 3.13 came with 3.2.1). The other extensions name
  their minimum the same way (`>=3.0.0`, `>=3.1.0`). Should the core release as 3.3.0 instead, `>=3.2.2` is still
  right, and `sdkVersion` refuses the install on any Hostwerk without SDK 3.15 anyway. `contributes.new` = `new-docs`
  (icon file), `new-sheets` (table), `new-slides` (tv) – so the backend does not call `host.setNewEntries`.
  `data.folders` = `data/office-data`, `data/office-engine`, `data/office`, `data/office-components` (1.1) (only "Remove completely" offers them,
  each listed with its size). Synergies: file-share, team, passwords, ai-media. Category `other`, permission `native`.
  Settings are the native `settings.tabs` slot; shortcuts come from the ui module (`ext:office:open` = Ctrl+Alt+O,
  new document/spreadsheet/presentation/search without default keys) – no manifest shortcuts needed.
- **Translations:** native extensions use English source texts and one `locales/de.json` (`{ "English": "Deutsch" }`,
  `$patterns` for template strings). There is **no `locales/en.json`** – like file-share and vms; `sdk/check.js` treats
  it as optional (wording overrides only). `locales/parts/*.json` were merged and deleted (nothing read them; the CI
  package step would have shipped them). Duplicate keys with different German were resolved to one wording:
  Font size → Schriftgröße, Whole words → Nur ganze Wörter, Clear → Löschen, Outline → Gliederung, Size → Größe,
  Type → Typ, Header → Kopfzeile, Enter → Eingabe, General → Allgemein, Horizontal/Vertical → Horizontal/Vertikal,
  Search → Suchen, Tab → Tabulator, Fade → Überblenden, Fill → Füllung, None → Keine, Normal → Normal,
  Series → Datenreihe, Slide → Folie, Split → Aufteilen, Home → Start, Design → Design, Larger/Smaller text →
  Größere/Kleinere Schrift, Inventory list → Inventarliste, plus a few sentences. Where the meaning really differed, the
  key was changed in the code: PDF signature tab "Type in" (Tippen), slide-show transition "Slide in" (Hereinschieben),
  vCard "Private" (Privat, was "Home"), PDF sidebar "Bookmarks" (Lesezeichen, was "Outline").
  `node ../tools/i18n-scan.js .` still lists ~380 "texts" – all code, XML, CSS, font names, format codes, shortcut
  labels, time-zone ids and the language-specific built-in templates (they carry their own German), no UI text.
- **Tests:** `tests/server-core.test.js` cross-checks the zip writer with Hostwerk's yauzl, resolved from
  `HOSTWERK_DIR` or the Hostwerk checkout next to the extensions; without it (GitHub CI) only that part is skipped with
  a diagnostic note – nothing is vendored. All other test libraries come from `npm ci` (devDependencies).

**Shell ↔ chunks / backend**
- `host.openAs({ bytes, ext, name, google? })` (viewer: Google "Open in Office", iWork "Edit as docx/pptx/xlsx") shows
  the bytes in the matching editor as a file that is not saved yet. `google = { id }`: saving uploads to Google Drive
  (`PUT /office/google/raw?id=&ext=`); for the pointer file itself (.gdoc …) with a fresh connection the host simply
  reopens it in Google mode. Otherwise the first explicit **Save** creates `<name>.<ext>` next to the original (library:
  `POST /office/create`; file-browser folder: `PUT /office/raw` with the first free name, never overwriting) and the
  editor switches to it; autosave never creates it. Versions, share, rename and "open in …" are hidden meanwhile.
- `?action=open&target=<path>[&root=<root>]` opens that file in its editor (address replaced, so Back does not reopen
  it). The backend now sends `root` as its own parameter (root ids contain ":" – `own:ab12`); the old
  `<root>:<path>` form is still understood (`openTarget()` in `shell/util.js`, unit-tested).
- `pdfworker` is only a build output (`ui/assets/pdfworker.js`); the loader only loads kinds from `formats.js`, so it is
  never treated as an editor.
- An exporter `run(ext)` returning `null` = handled by the chunk (e.g. its print dialog): no error, no download, no
  print of the shell, nothing stored by "then: convert".
- Dirty state: chunks only call `host.dirty(true)`; the shell clears it after a successful save when no edit came in
  while saving (edit counter), otherwise it stays dirty and autosaves again.
- Presenter window (`#/office-edit?…&present=<token>&window=1`): no top bar and no banners – the slides chunk fills the
  window.
- Response shapes fixed on the shell side: tags `{ tags: [{ tag, count }] }` (were dropped), Google state
  `account: { name, email }`, `POST /office/google/connect` answering the new state (desktop app) or `{ url }`,
  integrations state `stopped` (button "Start" → `POST /office/integrations/start`; "Open in the full editor" starts
  it first), frame events `modified` (shows "Unsaved changes"), `saved`, `close` (leaves the editor). A Google pointer
  without connection is read with `?pointer=1` (the backend answers 409 otherwise). Settings got the Google OAuth
  client fields (`integrations.google.clientId/clientSecret`, with the backend's guide and redirect address) and
  `integrations.hostUrl` in the defaults.

**Shortcuts (Ctrl+R, Ctrl+1 … 9) – decision**
- How Hostwerk handles them: `ui/src/shortcuts.js` binds Ctrl+R (reload the interface), Ctrl+Shift+R (restart, asks)
  and Ctrl+1 … 9 (jump to pages) in the **renderer** – one `keydown` listener on `window` (bubbling phase, `App.jsx`).
  `electron/main.js` sets no application menu of its own; Electron's default menu (View → Reload = Ctrl+R) only
  reacts to keys the page did not handle on Windows/Linux, and Hostwerk itself relies on that (its own Ctrl+R /
  Ctrl+Shift+R call `preventDefault()`; otherwise Ctrl+Shift+R would force-reload without asking). There is no menu
  accelerator for Ctrl+1 … 9.
- Therefore **Office keeps Word's/Excel's keys while the focus is inside the editor**: Docs Ctrl+R = align right,
  Sheets Ctrl+R = fill right, Ctrl+D = fill down, Ctrl+1 = Format cells (Ctrl+Shift+F also works). The editors call
  `preventDefault()` + `stopPropagation()`, so neither Hostwerk's listener nor Electron's menu sees them; outside the
  editor area Hostwerk's keys work unchanged. Safety net in the editor host: while there are unsaved changes, a Ctrl+R
  that no editor used (e.g. typing in a cell) is swallowed instead of reloading the interface.
- Other keys stay off Hostwerk's globals: link Ctrl+Shift+K (Ctrl+K = search), justify Ctrl+Shift+J (Ctrl+J = AI),
  strikethrough Alt+Shift+5 (Ctrl+Shift+S = start all), sub/superscript Ctrl+= / Ctrl+Shift+= (Ctrl+, = settings),
  track changes Ctrl+Shift+E, comment Ctrl+Alt+M, footnote Ctrl+Alt+F; Ctrl+N and Ctrl+F only while the focus is on the
  Office page.

**Live test in a throwaway Hostwerk 3.2.1 (2026-11-19) – decisions from it**
- **User data is never translated.** Hostwerk's extension JSX runtime translates every string child and the props
  `title/placeholder/label/text/hint/subtitle/confirm/alt/aria-label/description` (plus `options[].label`, `tabs[].label`) of
  *every* element – also of our own components (`<NavButton label={tag}>` was translated before NavButton ran). Rules:
  user data as a child goes into `<Keep>`; as one of those props through `<Raw as="span|button|…" title={name}>` or
  `<Raw as={Tabs} tabs={…}>` (props of a Raw element stay as they are); own components get user data under other
  prop names (`caption`, `value`, `name`). Helpers: `ui/src/shell/common.jsx` (shell) and `ui/src/editors/_shared/keep.jsx`
  (editors, no dependency on the kit). Texts with user data: `t('… {name} …', { name })` – vars are not translated.
- **Shortcut labels** are written in English in the code and shown through `ui/src/editors/_shared/keys.js`
  (`keyLabel('Ctrl+Shift+Del')` → "Strg+Umschalt+Entf", `withKeys(label, keys)`); colour swatches are named by
  `_shared/colors.js` (`colorName(hex)` → "Yellow" → "Gelb", tested).
- **German function names (Sheets):** `sheets/formula/i18n.js` maps every Excel function of the engine to Microsoft's
  German name (all 446 checked against a German Excel 16 via COM, Range.Formula → Range.FormulaLocal, 2026-11-19: 4 fixed –
  FORECAST = SCHÄTZER, DOLLAR = DM, WRAPCOLS/WRAPROWS = SPALTENUMBRUCH/ZEILENUMBRUCH; error values and WAHR/FALSCH match,
  #ÜBERLAUF!/#KALK! cannot be typed as constants in Excel) (WENN, SVERWEIS, ZÄHLENWENN …), WAHR/FALSCH, German error codes (#NV, #WERT!, #BEZUG!, #ZAHL!, #ÜBERLAUF!,
  #KALK!) and the argument names of the hints. Input/display only: `toInternal`/`toDisplay` (and cell display, editing text,
  autocomplete, hints, Insert function, "Show formulas") follow `loc.fnLang`; files and the model keep English names.
  English names are always accepted. Setting `sheets.functionNames: 'auto' | 'en' | 'de'` (auto = Hostwerk's language).
  The engine's own undotted aliases (CEILINGMATH …) are not offered.
- **Narrow windows:** ≤ 700 px the ribbons drop their group labels and scroll sideways, slides become a strip above the
  slide, side panels (docs navigation/comments, slides format, PDF sidebar) are sheets over the content, the viewer puts
  its list above the item; every content area keeps a minimum height. ≤ 600 px Hostwerk shows its own navigation above
  the page, so the editor host sets its height to Hostwerk's scroll area and scrolls itself into view; the PDF sidebar
  starts closed and Docs starts fitted to the width there.
- **Soft locks** live 5 minutes and the open editor renews them every 2 minutes; leaving the editor, closing the tab or
  reloading releases them (`pagehide` + keepalive request). Only a crashed/killed browser leaves a lock (≤ 5 min).
- **Links** in the plain browser (no Hostwerk window) show the "copy link" dialog – never a browser tab.
- **Toast:** the core's extension toast has `warn`/`dismiss`/`done` since SDK 3.15 (`ui/src/extNative.js`); `ui/src/lib/toast.js` uses
  them and stays the fallback for older Hostwerk versions.
- **Screenshots:** `docs/screenshot.png` (en) and `docs/screenshot.de.png` (Office inside Files – PNG, the shared README tools
  expect these names) and `docs/screenshot-docs|sheets|slides|pdf.webp` (en, the editor in front, WebP q 0.9), 1440×900,
  light theme, listed in `manifest.screenshots` (max. 6). Made with demo data in a throwaway Hostwerk (folder path masked).

**Office inside Files (SDK 3.15) – decisions**
- **No own page in the sidebar** (`page.nav: false`). `#/ext/office` (Hostwerk's "+ New" menu, shortcuts, quick search, backups)
  leads into Files: `?action=new|new-docs…|template` → Files' Office root + the New dialog, `open&target=` → the editor,
  everything else → the Office root. `#/ext/office/settings` and `#/ext/office/edit?…` (older links) stay.
- **The library is a root of Files:** `host.files.addRoot({ id: 'library', label: host.t('Office documents'), path: libraryDir,
  icon: 'file', create: true })` → `ext:office:library`; re-added when `libraryDir` or the language changes, removed on
  deactivate/uninstall. A path of this root (or of any root that contains the library) IS the library path (`library.loc`
  canonicalises it), so favourites, versions, shares, locks and the search index stay one; the UI does the same (`canon()`).
- **files.open** (`priority 20`): every format Office opens except plain text/web/code (txt, html, xml, json, images, zip …);
  Markdown and CSV only outside the folders of web servers, projects and apps (there Hostwerk's code editor stays) – and always
  from Explorer. It respects `general.openFromFiles` and `integrations.defaultOpen` (built-in / full editor / desktop app).
  It opens the **screen** `#/office-edit?…&back=<#/files?root=…&path=…&select=…>` (the editor in front, no sidebar); Back:
  history when we came from Files, `location.replace(back)` for files from Explorer (`&os=1`). The right-click "Open in
  Office" only appears when a click does not already open Office (Files lists the opener as *Open with "Office"*).
- **files.new**: Document / Spreadsheet / Presentation / PDF create `<Document>.<default ext>` at once in the open folder
  (first free name, `POST /office/create?root=`) and open it – rename in the editor's title bar; "From template or other
  format …" opens the New dialog (folder fixed to the open folder outside the library).
- **Dialogs everywhere:** Office's dialog host is mounted once through the slot `sidebar.footer` (it renders nothing; Hostwerk's
  Modal portals to `.app`), so Share, Versions, Info, Tags, Convert and New open right in Files. Without a mounted host the New
  dialog falls back to the screen `#/office-edit?dialog=new&…`. Engine/full-editor dialog hosts are mounted in several places –
  only the one mounted last shows the dialog (`useTopHost`).
- **The Office section** (`sections['files.folder']`, only in the Office root): default-app hint, new buttons, full-text search,
  views Recent · Favourites · Shared · Tags · Templates · Trash (Office's own trash – files deleted in Files go to Hostwerk's
  trash), each row with the same right-click actions; link to the settings. The former library page (Library.jsx) is gone.
- **"Open with"** (`server/osreg.js`): on activate (4 s later) `host.fileTypes.register({ id: 'office', appName: 'Hostwerk Office',
  types })` for 32 types (docx … ics, each with mime + kind) once per install/upgrade (marker `office-data/os.json` = version +
  type list); a registration the user removed himself is not forced back before the next update. `integrations.osOpenWith`
  (default on) off → removed. Developer start → `EDEVBUILD`, silent, shown in the card. `exports.uninstall` unregisters (Hostwerk
  does it too). First registration → one toast after start ("now under Open with", button to the settings). The card
  (`osCard.jsx`) shows per kind ✓ default / some / only under "Open with" / not registered and polls after "Make it the
  default app" (Windows: `openDefaultsSettings` → `ms-settings:defaultapps?registeredAppUser=Hostwerk%20Office`; Linux: register
  with `setDefault` after a confirm; macOS / server / store: an honest notice). Unit test against the core's FileTypes with a
  simulated registry: `tests/server-files.test.js` (registry exactly as before after removal).
- **Shared library chunks:** jszip and pdf.js are built once (`ui/assets/lib-jszip.js`, `lib-pdfjs.js`); the editor chunks load
  them first (top-level await, same `?v=`) into `globalThis.__officeLib` and import a stub. The pdf.js worker keeps its copy.
- **Known gaps:** rename/move/delete done by Files itself (not by Office) does not move favourites/tags/versions of that file
  (Files has no event for it); the Office trash only holds files deleted by Office itself (Files uses Hostwerk's trash).

## Office format check (real Microsoft Office 16, German, Windows)

Every built-in template (docs ×7, sheets ×6, slides ×4; en in every writable format, de in the main one), the source
files made by Office (docx/docm/xlsx/xls/xlsm/pptx/ppt/pptm, three real .doc files shipped with Office) and their round
trips through our readers and writers were opened via COM, one file per fresh Office process (75 s watchdog, leftover
processes killed, none left running). Every Word file was also opened again with alerts on (`--strict`), so a repair
prompt would have shown. PDFs of all files were exported. Harness: `node scripts/gen.mjs`, then `node scripts/drive.mjs
results/jobs.json results/main.jsonl - --pdf --wordpdf` and `… jobs-strict.json strict.jsonl - --strict`, then
`node scripts/report.mjs`. A copy lives in `../../office-check/` (re-run with PDFs and results: `rerun-2026-11-19/`). Word here is unlicensed (it opens read-only and exports
PDF, SaveAs hangs) and the Trust Center blocks VBA project access, so macros are checked by package structure only.

Findings of the first run and their fixes (each has a regression test in `tests/docs-io`, `formats-*`, `slides-ppt`):

| # | Finding | Fix |
|---|---|---|
| 1 | docx "damaged": settings name separator footnotes without `footnotes.xml` | footnotePr/endnotePr ids only when the part exists |
| 2 | docm patched to docx/dotx kept the macro main type / VBA parts | main type from the target ext; vbaProject.bin, its rels, vbaData.xml and overrides dropped for macro-free targets |
| 3 | pptx "damaged": master body placeholder without type | masters use title/body/dt/ftr/sldNum (idx 1–4); layouts get their own ph definitions that match the slides |
| 4 | pptm/ppsm/potm fresh packages had the pptx main type | main content type per ext |
| 5 | slide background `bgRef` lost `n` → patch threw → silent regenerate | fixed; a failed patch now throws `PPTX_PATCH_FAILED` with `regenerate()` – the slides editor offers "Save as a copy" (`host.saveCopy`: the copy next to the original, the editor goes on with it), never a silent loss; exports fall back to the regenerated package with a warning |
| 6 | odt with `style:header-first` (ODF 1.3) rejected by Word | title page = first-page master `OfFirstPage` started by the first paragraph; the reader knows it and header-first |
| 7 | ods/fods cross-sheet formulas written as `of:=Data![.D7]` (#NAME?) | own OpenFormula writer (`sheets/io/odf-formula.js`: `[$Data.D7]`, `[$Data.A3:.A6]`, quoted names, `$`, whole rows/cols, arrays, `COM.MICROSOFT.` functions, strings untouched) placed into SheetJS' output; the reader takes the raw `table:formula` from content.xml; ods zip now has `mimetype` first and stored |
| 8 | RTF footnotes dropped by Word | `tnbj` in the document header |
| 9 | Word 6/95 .doc gave "not a Word 97–2003 document" | wIdent 0xa59b/0xa59c/0xa5db/0xa5dc or nFib < 0xc0 → `DOC_TOO_OLD` "too old for the built-in reader – the Office engine opens it" (the shell converts .doc with the engine first and offers to set it up) |
| 10 | ppt notes contained "*", every 2007+ slide had the blank layout | notes = notes-body placeholder only; layout from RoundTripContentMasterInfo12 or the slide's placeholders |
| 11 | TOC marked dirty + updateFields → "update fields?" prompt | TOC written filled, linked to `_Toc` bookmarks, no dirty flag, no updateFields |

Result of the re-run after the fixes (155 files, 62 of them also strict):

| File group | Result |
|---|---|
| Office-made sources (docx, docm, xlsx, xls, xlsm, pptx, ppt, pptm, 2 × Word 97 .doc) | all open, facts match |
| Word 6/95 .doc (WBTRANS.DOC) | blocked by Word's own Trust Center file block (not our file); our reader now gives the engine message |
| Docs templates: docx ×14, dotx, docm, odt, rtf, html ×7 each | all open, no repair, strict open clean (was: every docx/dotx/docm "damaged", report.odt "damaged", html/report repair prompts) |
| Sheets templates: xlsx ×12, xltx, xls, csv, html ×6, ods ×12 | all open, every formula value matches Excel |
| Slides templates: pptx ×10 (incl. 4:3/A4 + fade), ppsx, potx, pptm, odp ×4 | all open, slides/shapes/notes match (was: every pptx/ppsx/potx/pptm unreadable) |
| Round trips Word (docx, dotx, docm, rtf, odt; docx from .doc) | all open, facts match (rtf footnotes now kept) |
| Round trips Excel (xlsx, xls, xlsm, ods) | all open, values match (ods cross-sheet formulas now correct) |
| Round trips PowerPoint (pptx from pptx/ppt/pptm, pptm, odp) – real reader | all open, facts match |

Remaining differences are expected, not defects: the report template's callout box is a one-cell table in Word (tables 1 vs
2); Word's HTML import has no footnotes and counts the callout icon as a picture; Word's ODT import ignores ODF tracked
changes (`rt-word-src-docx.odt`: our insertion/deletion are valid ODF 1.2 – not verified against LibreOffice here).

## Office 1.1 – expansion contract (binding for all agents of the 1.1 round – done; kept as the contract of the parts)

User's order: close every gap in "Known limits" below where it is technically possible, make everything visibly nicer,
smoother, faster and easier to use (coloured file icons, thumbnails, motion), and build every "Tipp" idea of the user's
roadmap: live fields, "/" commands, LaTeX formulas, diagrams from text, compare versions, live data from Hostwerk,
database tables as sheets, Kanban/calendar/gallery views, forms from sheets, rules & notifications, presentation from
a document, phone as remote control, OCR, real redaction, co-editing, search by meaning, encryption with the Hostwerk
vault, clean before sharing, publish as web page, scheduled reports, backup with time machine. Version → **1.1.0**.

**Ownership (6 parallel agents – never edit another agent's files; ask the integrator via your final report):**

| Agent | Owns |
|---|---|
| core | `../../hostwerk/**` (never `server/core/download.js`, `server/core/proc.js`, `server/services/db/engines.js` – foreign uncommitted work) |
| shell | `server/**`, `ui/src/index.jsx`, `ui/src/shell/**`, `ui/src/style.css`, `ui/src/lib/**`, `tests/server-*`, `tests/shell-*`, `tools/**`, `manifest.json` |
| docs | `ui/src/editors/docs/**`, `tests/docs-*` |
| sheets | `ui/src/editors/sheets/**`, `tests/sheets-*` |
| slides | `ui/src/editors/slides/**`, `tests/slides-*` |
| pdfview | `ui/src/editors/pdf/**`, `ui/src/editors/viewer/**`, `ui/src/editors/pdfworker/**`, `tests/pdf-*`, `tests/viewer-*` |

`ui/src/editors/_shared/**`: add NEW files freely; change an existing shared file only if you created it this round.
**Texts:** new German translations go ONLY into `locales/parts/<agent>.json` (`{ "English": "Deutsch" }`); the
integrator merges them into `locales/de.json` at the end – never edit `locales/de.json` this round.

**Shared server contracts (shell implements, editors call through `ui/src/lib/api.js` helpers the shell adds – until
they exist, call `fetch('/api/office/…')` directly with the same shapes):**
- Live data: `GET /office/data/sources` → `[{ id, label, kind: 'table'|'value', params: [...] }]` (Hostwerk services,
  resources, Minecraft players, backups, tasks, websites, databases …); `GET /office/data/query?source=&…` →
  `{ columns: [{ id, label, type }], rows: [[…]], value?, updatedAt }`; live updates via Hostwerk's websocket broadcast
  `{ type: 'office-data', source, … }`.
- Databases: `GET /office/db/servers` → Hostwerk DB servers; `GET /office/db/tables?server=&db=`;
  `GET /office/db/rows?server=&db=&table=&offset=&limit=&sort=&filter=` → `{ columns, rows, total, primaryKey }`;
  `POST /office/db/rows {server, db, table, changes: [{ op: 'update'|'insert'|'delete', key, values }]}` (transaction,
  confirm dialog in the UI before writing; read-only when no primary key).
- Live fields in documents/slides: `GET /office/fields` (catalog of fields), `GET /office/fields/value?id=&params=`.
- Forms: `POST /office/forms {path|root+path, sheet, columns, title, public: { enabled, expires, password }}` →
  `{ id, url }`; public submit page via `host.publicRouter` appends a row to the sheet (server-side xlsx append with the
  sheets' own writer logic exposed as a pure module, or a queue the editor applies when the file is open).
- Rules & notifications: `GET/PUT /office/rules?path=` – `[{ id, when: { sheet, cell|range, op, value }, then: { notify:
  { channel: 'hostwerk'|'discord'|'ntfy'|'mail', text } } }]`, evaluated by the server on save and on schedule; uses
  Hostwerk's notification API (`host.notifications`).
- Co-editing: websocket `/api/office/collab?path=…` (via `host.on('upgrade')`) relaying Yjs updates + awareness
  (cursor, name, colour); the server keeps a Y.Doc per open file in memory and writes the file through the normal
  saver of the first editor ("leader") – docs (y-prosemirror), sheets (cell map), slides (element map). Login: Hostwerk
  session or a share link with edit right.
- Remote control: `/api/office/remote` websocket + a QR code in the slide show; the phone page (public, token) shows
  next/previous, notes, timer, laser.
- AI (search by meaning, presentation from a document, "/ai" command): through `host.ai` / Hostwerk's AI with the user's
  configured provider; never a fixed cloud.
- Components on demand (heavy things that don't fit into the 10 MB package: OCR engine + language data, Mermaid,
  CJK PDF fonts/CMaps): `GET /office/components` (state), `POST /office/components/install {id}` – downloaded once from
  pinned URLs with **pinned SHA-256** into `<data>/office-components/<id>/`, served by `GET /office/components/<id>/<file>`;
  the UI asks once (size shown) like the Office engine. Never execute a download without matching hash.

**Design direction:** each app has its colour like in Microsoft Office / LibreOffice, from Hostwerk tokens where
possible: documents blue, spreadsheets green, presentations orange, PDF red, other files neutral. Use it for file
icons (everywhere: Dateien list/grid, Office section, recent, search, tabs), the active ribbon tab underline, the
primary editor accents – but keep Hostwerk's look (cards, spacing, typography, light/dark). Motion: short (120–200 ms),
purposeful, `prefers-reduced-motion` respected. Performance budget: opening a 1 MB docx/xlsx/pptx < 1.5 s, typing/
scrolling at 60 fps, no long tasks > 100 ms on interaction; chunk size growth ≤ 200 KB per editor (lazy-load sub-features).
**Package budget: ≤ 9.8 MB total** – check with `node tools/build.js`; anything heavier is a component on demand.

Every agent: live-check its work in a throwaway Hostwerk (never the user's real data; see `../tools/shots.js` and
`C:\Users\Pasca\Documents\Hostwerk\office-check\live-test\hw-run.js`), keep `node --test tests/*.test.js` and
`node ../../hostwerk/server/extensions/sdk/check.js . --strict` green, don't npm install without listing the package in
your report (the integrator installs), don't commit.

### Host object 1.1 (shell → editor chunks; written by the shell agent – code against this)

Every 1.0 member of `host` stays. New members are **always present** (feature-detect only the server behind them: a
call throws `Error` with `status` 404/501 and a German/English message when the server route is missing). Client helpers
for the same routes are in `ui/src/lib/api.js` (`office.data.*`, `office.db.*`, `office.fields.*` …) – chunks may import
them, but prefer `host.*` (the host adds dialogs, confirmations and the file location).

| member | meaning |
|---|---|
| `host.file` | `{ name, path, root?, ext }` of the open file (same object as the `file` prop, kept current after rename/save-as) |
| `host.data.sources()` | → `[{ id, label, kind: 'table'\|'value', params: [{ id, label, type, options? }] }]` |
| `host.data.query(source, params?)` | → `{ columns: [{ id, label, type }], rows: [[…]], value?, updatedAt }` |
| `host.data.subscribe(source, params, cb)` | `cb(result)` at once and on every websocket broadcast `{ type: 'office-data', source }` (fallback: poll every 30 s while visible) → `unsubscribe()` |
| `host.data.pick({ kind? })` | shell dialog "Insert live data" (source + params + preview) → `{ source, params, label }` or `null` |
| `host.db.servers()` · `host.db.tables(server, db?)` | Hostwerk's database servers / `[{ db, table, rows?, primaryKey? }]` |
| `host.db.rows({ server, db, table, offset, limit, sort, filter })` | → `{ columns, rows, total, primaryKey }` (`primaryKey` null = read-only) |
| `host.db.write({ server, db, table, changes })` | shell **confirm dialog** (n changes, first rows listed) then `POST /office/db/rows` → result, or `null` when the user cancels |
| `host.db.pick()` | shell dialog "Open a database table" → `{ server, db, table }` or `null` |
| `host.fields.list()` · `host.fields.value(id, params?)` | live-field catalog `[{ id, label, group, params? }]` / `{ value, text, updatedAt }` |
| `host.fields.pick()` | shell dialog "Insert field" (catalog, search, preview of the value) → `{ id, params, label }` or `null` |
| `host.forms.list()` · `host.forms.create({ sheet, columns, title, public })` · `host.forms.open(id?)` | forms of THIS file / new form → `{ id, url }` / shell dialog with link, QR code, responses, delete |
| `host.rules.get()` · `host.rules.set(rules)` · `host.rules.open({ sheet?, range? })` | rules of THIS file (contract shape) / shell rules editor dialog (prefilled with the selection) |
| `host.collab.connect({ doc, awareness?, user? })` | needs `yjs` (see "Yjs" below). Connects `doc` (a `Y.Doc` of the shared `yjs`) to `/api/office/collab?path=…` → provider `{ doc, awareness, synced: Promise, status: 'connecting'\|'connected'\|'offline', on(evt, fn), off(evt, fn), destroy() }` (events `status`, `synced`, `peers`). Awareness local state is set to `{ user: { name, color, id } }`; the top bar shows the co-editors' avatars. `host.collab.available()` → bool (server route answers) |
| `host.remote.start({ notes?, title })` | starts remote control for the slide show → `{ token, url, qr: dataUrl, on(evt, fn), send(msg), stop() }` (events `next`, `prev`, `goto`, `laser`, `peers`); `host.remote.dialog()` shows the QR dialog |
| `host.ai.ready()` · `host.ai.quick(task, text, { context })` · `host.ai.ask({ prompt, text?, system?, json? })` · `host.ai.panel(question, context)` | Hostwerk's AI with the user's provider (`ready` false = hide AI entries). `quick` = core `/ai/quick/<translate\|explain\|summarize\|describe>`; `ask` = `POST /office/ai` → `{ text, json? }`; `panel` opens Hostwerk's AI side panel (Ctrl+J) |
| `host.components.ensure(id, { reason })` | component on demand (ids `ocr`, `mermaid`, `cjk-fonts`; aliases `ocr-deu`, `pdf-cmaps` … work too): installed (`state: 'ready'`) → `{ id, base, files: [names] }` at once; else ONE dialog (size, progress) like the engine dialog → `{ id, base }` or `null` (declined). `base` = `/api/office/components/<id>/` (absolute URL for `fetch`/`import`). `host.components.state()` → list |
| `host.compare(a, b?)` | opens the compare view. `a`/`b` = `{ id }` (version id), `'current'` (unsaved editor state via the saver) or `{ bytes, label }`. Calls the chunk's `Compare` component (see below) in a full-size shell sheet; `b` defaults to `'current'` |
| `host.palette.register(commands)` | commands for the command palette (**Alt+Q** like Office's "Tell me"/search box, and Ctrl+Shift+P when the chunk did not use that key itself – both are free in Hostwerk): `[{ id, title, group?, keys?, icon?, run, when?() }]`, replaces the chunk's previous list → `unregister()` |
| `host.insertFiles(cb)` | `cb([{ name, ext, mime, size, bytes(): Promise<Uint8Array>, path?, root? }], { x, y })` when files from Dateien or the desktop are dropped on the editor (the shell shows the drop overlay) → `unregister()`; a chunk without it gets nothing (the shell only opens dropped office files as tabs/new editors) |
| `host.toast(text, { kind, action: { label, run }, undo })` | toast with an action (undo toasts) |
| `host.perf(mark)` | `'interactive'` – call once when the document is usable (the shell measures open → interactive per kind) |

**Optional chunk exports (1.1):** `reportsInteractive: true` (the chunk calls `host.perf('interactive')` itself – otherwise
the first painted frame after mounting counts), `Compare({ a: { bytes, label }, b: { bytes, label }, ext, host })` – side-by-side/inline
diff view (docs/sheets/slides each their own); `commands` (static palette entries, same shape as above);
`thumbnail({ bytes, ext }) → Promise<Blob|HTMLCanvasElement|null>` – if a chunk can paint its first page faster/better
than the generic DOM rasteriser (pdf/slides), the thumbnail queue uses it.

**Yjs:** `yjs`, `y-protocols/sync`, `y-protocols/awareness` are a **shared library chunk** (`ui/assets/lib-yjs.js`, like
lib-jszip) – `import * as Y from 'yjs'` in a chunk (also inside `y-prosemirror`) resolves to that one instance, the
same the shell's provider uses (two Yjs copies would break `instanceof`). The provider class is `ui/src/lib/collab.js`.

**Thumbnails (shell):** `PUT /office/thumb?path=|root=&path=&mtime=` raw `image/webp` (≤ 256 KB, 320 px wide) stores a
generated first-page thumbnail in the data folder (keyed like metadata, dropped when the file's mtime changes);
`GET /office/thumb?…&v=` serves it, else the embedded one, else 404; `Entry.thumb` = either exists. The shell renders
missing ones lazily in idle time (`ui/src/lib/thumbs.js`, one at a time, offscreen) with the chunk's `thumbnail()` or
its `Preview` rasterised, then uploads them.

**File-level routes the shell UI calls (server agent: same shapes, or tell the shell agent):**
`GET /office/vault` → `{ available, unlocked }` · `POST /office/vault/encrypt {loc}` / `POST /office/vault/decrypt {loc}` → Entry ·
`POST /office/clean/preview {loc}` → `{ items: [{ id, label, count?, detail? }] }` · `POST /office/clean {loc, remove: [ids], copy}` → Entry ·
`GET /office/publish/targets` → `[{ id, label, url? }]` (`host.sites`) · `GET /office/publish?path=` → `[{ id, target, url, updatedAt }]` ·
`POST /office/publish {loc, target, title, path?}` → `{ id, url }` · `DELETE /office/publish/:id` ·
`GET /office/reports?path=` → `[{ id, schedule: { every: 'day'|'week'|'month', at: 'HH:MM', weekday?, day? }, format, recipients: [{ channel, to }], lastRun?, nextRun? }]` ·
`POST /office/reports {loc, …}` (with `id` = update) · `DELETE /office/reports/:id` · `POST /office/reports/:id/run` ·
`GET /office/forms?path=` → `[{ id, url, title, sheet, columns, public, responses? }]` · `DELETE /office/forms/:id` ·
`GET /office/search?q=&semantic=1` → hits with `score` · `GET /office/backups?path=` → `[{ id, time, size, label? }]` (Hostwerk's
backup center) · `GET /office/backups/raw?path=&id=` raw · `POST /office/backups/restore {loc, id}` ·
`GET /office/components` → `[{ id, label, size, installed, progress? }]` · `POST /office/components/install {id}` ·
`GET /office/collab/state?path=` → `{ peers }` (200 = co-editing available) – websocket `/api/office/collab?path=|root=&path=`,
binary frames exactly like y-websocket (0 = sync step 1/2/update, 1 = awareness, 3 = query awareness) ·
`POST /office/remote {loc, title}` → `{ token, url, qr }` · websocket `/api/office/remote?token=&role=presenter|phone`, JSON
`{ type: 'next'|'prev'|'goto'|'laser'|'notes'|'slide'|'peers', value }` relayed between both roles · `DELETE /office/remote/:token` ·
`POST /office/ai {task?, prompt, text?, system?, json?, lang, loc}` → `{ text, json? }` · `GET /office/qr?text=` → `image/svg+xml`
(Hostwerk's `qrcode` via `host.lib`); forms/remote responses may carry `qr` (SVG markup) directly.

**Shell 1.1 – decisions (shell agent)**
- Icons: `ui/src/lib/fileIcons.jsx` (`<FileTypeIcon name|entry|ext size caption showExt />`, `iconSvg()`), families/colours
  in `ui/src/lib/fileTypes.js` (pure, `tests/shell-icons.test.js`). ≤ 20 px = coloured tile with a white glyph, ≥ 24 px =
  paper (Hostwerk's `--surface`) with the tile. Files slots (SDK 3.16, ignored by older Hostwerk): `files.icon`
  `{ id, match, Component({ name, size, entry }), info(name) → { color, family, label } }` and `files.thumb`
  `{ id, match, thumb({ root, path, name, mtime, size }) → Promise<url|null> }` (also prefetches the editor chunk).
- Thumbnails: `ui/src/lib/thumbs.js` queue (idle time, one at a time), renderer `shell/thumbHost.jsx` (React portal into a
  hidden 820 px stage; mounted by the dialog host, the editor host and the New-dialog screen). The DOM is rasterised
  with inlined computed styles through an SVG `foreignObject` (Chromium/Electron: canvas stays clean); pdf.js canvases are
  copied directly. Built-in template thumbnails are cached as data URLs in localStorage (`office:tpl-thumbs`, per chunk
  version and language). Without `PUT /office/thumb` the pictures live for the session only.
- Speed: `loader.js` `prefetchEditor` (modulepreload of chunk + shared libs on hover/focus), `warmEditor` (load in idle
  time: click on a file, Office root opened); chunks stay loaded. `lib/perf.js` measures `office:open:<kind>`
  (click → interactive) – numbers in `performance` and `localStorage['office:perf']`; `office:perf-log = 1` logs them.
  The chunk's editor is memoised (`ChunkView`), so status/save-state updates never re-render the document.
- Editor menus are the shell's own drop-down (`shell/menu.jsx`): Hostwerk's `contextMenu` has no host on screens
  (`#/office-edit`), so the 1.0 "More" and "Download as" menus did nothing there. (Core could mount `ContextMenuHost` in
  `app-bare` screens too.)
- Command palette (`shell/palette.jsx`, ranking `shell/paletteModel.js`, tested): Alt+Q, or Ctrl+Shift+P when the editor
  did not use it; the chunk's commands + every file action; recently used first (`localStorage['office:palette-recent']`).
- Feature dialogs (`shell/features.jsx`, `askFeature(type, props)`) – one bus, shown by the `FeatureHost` mounted last
  (dialog host in `sidebar.footer` and every editor host); the file actions themselves are one list
  (`shell/fileFeatures.js`) for Files' right-click, the Office section, the editor's menu and the palette.
- View transitions: `ops.viewTransition(update, 'office-open'|'office-close')` (types, 140–180 ms, off with reduced
  motion); the editor host fires `office:view` when mounted.
- The Office section hides Files' own "No file opened" below it (`.fb-folder-pane:has(.of-shell-fsec) > .empty`) – remove
  once the core does it itself.

### Server 1.1 (server agent – the routes as built; JSON, lists wrapped `{ <key>: [...] }` – `list()` in api.js unwraps)

Modules: `server/v11.js` (wiring) · `routes11.js` (registered before the error handler) · `datasources.js` · `dbsheets.js` ·
`fields.js` · `forms.js` · `rules.js` · `notify.js` · `collab.js` · `remote.js` · `components.js` · `semantic.js` · `encrypt.js` +
`cfbwrite.js` · `clean.js` · `sheet.js` (xlsx/ods/csv read + row append) · `zippatch.js` (raw-copy zip patcher) · `extras.js`
(thumbs, publish store, reports, library snapshots). Tests: `tests/server-v11.test.js`. Every library write fires `onWrite`
(rules, publish "stale", collab "saved"), renames `onRekey` (rules/forms/publish/reports/rooms follow). `loc` = top-level
`{ path }` or `{ root, path }` (the vault routes also take `{ loc: {…} }`).

**Websockets – IMPORTANT for the shell:** Hostwerk's core hands only `/ws/*` upgrades of logged-in users to extensions and
destroys every other upgrade *before* extensions see it (`hostwerk/server/index.js` `onUpgrade`). Connect to
**`/ws/office/collab?path=|root=&path=`** and **`/ws/office/remote?token=&role=presenter|phone`** (`/api/office/…` is matched
too but never arrives with today's core – `office.wsUrl()` maps `/api/office/` → `/ws/office/`). A phone without a
Hostwerk login cannot use a websocket at all: it uses the public token page (SSE + POST) below.

| Route | Shape |
|---|---|
| `GET /office/data/sources` | `{ sources: [{ id, label, kind, params: [{ id, label, options? }] }] }` – services, service, resources, system, disks, minecraft, players, playerCount, backups, tasks, websites, databases, certificates, uptime, hostwerk |
| `GET /office/data/query?source=&<param>=…` (or `params=<json>`) | `{ columns: [{ id, label, type }], rows, value?, updatedAt }`; broadcast `{ type: 'office-data', source, updatedAt }` after matching core events (only for sources queried in the last 15 min, ≤ 1 per 3 s and source) |
| `GET /office/db/servers` · `/office/db/databases?server=` · `/office/db/tables?server=&db=` | `{ servers: [{ id, name, type, kind: sql\|mongo, engine, running, readOnly }] }` · `{ databases: [{ name, system }] }` · `{ tables: [{ name, view, rows }] }` |
| `GET /office/db/rows?server=&db=&table=&offset=&limit=(≤1000)&sort=&dir=&filter=<json [{ col, op, value }]>` | `{ columns: [{ id, label, type, key, nullable }], rows, total, primaryKey: [cols], readOnly, offset, limit }` (ops `= != < > <= >= contains LIKE NOT LIKE IS NULL IS NOT NULL`, through the core's own drivers) |
| `POST /office/db/rows {server, db, table, changes: [{ op: update\|insert\|delete, key: { pk: v }, values }]}` | one transaction (MySQL/MariaDB/PostgreSQL/SQLite; MongoDB one by one, `atomic: false`); each update/delete must hit exactly one row, else rollback + 409 → `{ ok, applied, results }`; 403 read-only connection, 409 no primary key / view |
| `GET /office/fields` · `GET /office/fields/value?id=&params=<json>` | `{ fields: [{ id, label, group, params }] }` – date, time, datetime, user, service, players, system, hostwerk, cell (`{ path\|root+path, sheet, cell }`), docprop (`{ path, prop }`) · `{ value, text, updatedAt, ttl }` |
| `GET /office/forms[?loc]` · `POST /office/forms {loc, sheet, title, description, columns: [{ label, type: text\|textarea\|number\|email\|date\|choice\|checkbox\|tel\|url, required, options, help, min, max }], timestamp (default true), thanks, notify: channel?, public: { enabled, expires (days or ms), password }}` | `{ forms: [Form] }` / Form `{ id, title, path, root, sheet, columns, public: { enabled, expires, password: bool }, closed, responses, lastResponse, pending, url, urls, open }`; an empty sheet gets a header row |
| (1.2) answers | the "Submitted at" stamp (local time of the Hostwerk) and date fields are real date cells: xlsx = serial number with the built-in date(-time) format 22/14 (the style is added to styles.xml once – a minimal styles part when there is none; 1904 workbooks respected; the column gets a width when it has none), ods = `office:value-type="date"` + a date style in the automatic styles, csv = ISO `2026-11-20 14:03:05`; queued answers keep `{ $date }`, `GET …/pending` hands them out as ISO text |
| `PUT /office/forms/:id` · `POST …/:id/close` · `DELETE …/:id` · `GET …/:id/pending` · `POST …/:id/pending/ack {count}` | answers arriving while the file is open (soft lock or co-editing room) wait in `pending` – written ≤ 30 s after the lock is gone, or the editor applies them and acks; broadcast `{ type: 'office-form', id, path, pending }` |
| public `GET/POST /api/office/pub/form/:token` · `POST …/unlock {password}` | form page (own CSP, honeypot `website`, signed time stamp ≥ 2 s, 10 answers/10 min/address, 300/h/form, password cookie) → row appended through `Library.write` (version + mtime check like a save) |
| `GET /office/rules?loc` · `PUT /office/rules?loc` (`{ rules }` or `[...]`) · `POST /office/rules/test {loc, rules?}` · `GET /office/rules/channels` | `{ rules: [rule + last: { met, at, value }] }`, rule `{ id, name, on, when: { sheet, cell \| range, op: > < >= <= = != contains notContains empty notEmpty changed, value, agg: any\|all\|sum\|count\|avg\|min\|max }, then: { notify: { channel: hostwerk\|discord\|ntfy\|mail\|all, text ({value} {cell} {file} {sheet} {rule}) } } }`; fires on "not met → met" (every change for `changed`) after saves/appends and every 2 min (stat first) · `{ channels: [{ id, label, ready }] }` |
| `GET /office/collab/state?loc` · ws `/ws/office/collab?loc` | `{ key, peers: [{ id, user, leader }], ws, yjs }` · y-websocket framing (0 sync – step 1 is always answered with a step 2, 1 awareness, 3 query) + **type 100 = control JSON (varString)**: server → `{ type: 'hello', id, leader, peers, user, fresh, mtime }`, `{ type: 'leader', leader: true }` (hand-over), `{ type: 'peers', peers }`, `{ type: 'saved', mtime, by }` (any write of the file), `{ type: 'reset', reason, leader }` (log too big: leader saves, all reload); client → `{ type: 'saved', mtime }` (leader). The leader saves with the normal `PUT /office/raw`. The server keeps the update log, compacted (1.2) with the small Yjs build shipped in `server/vendor/yjs.cjs` (70 KB, MIT, licence next to it): after 300 updates or when the log doubled (+ 1 MB) all updates go into a fresh `Y.Doc` (gc on – deleted text keeps no content) and come out as ONE update; only a compacted state over the cap still resets – ≤ 24 clients/room, 200 rooms, 48 MB log, rooms dropped 2 min after the last client; team members over remote access need the right for `PUT /office/raw` |
| `POST /office/remote {loc?, title?}` · `DELETE /office/remote/:token` · ws `/ws/office/remote?token=&role=presenter\|phone` | `{ token, url, urls, qr (SVG markup), ws, expires (12 h), lan: { reachable, remote } }` (`reachable` false = only localhost URLs → ask to switch on remote access). Presenter → `{ type: 'slide', value: n \| { index, total, title, notes } }`, `{ type: 'notes', value }`, `{ type: 'timer', value: { running, elapsed, startedAt } }`, `{ type: 'blank', value }`, `{ type: 'state', … }`, `{ type: 'end' }`; server → presenter `{ type: next\|prev\|first\|last\|goto\|blank\|laser, value, cmd, n?, x?, y?, on? }`, `{ type: 'peers', value }`; phone (ws) sends `{ type: 'next' }` …, gets `{ type: 'state', slide, total, title, notes, blank, timer, presenter }` |
| public `GET /api/office/pub/remote/:token` (+ `/app.js`, `/events` SSE, `POST /cmd {cmd, n, x, y, on}`) | the phone page behind the QR code (20 commands/s) |
| `GET /office/components` · `POST /office/components/install {id}` · `DELETE /office/components/:id` · `GET /office/components/:id/<file>` | `{ components: [{ id, label, version, license, download, size, installed, state: missing\|installing\|ready\|error, progress, error, files, base, aliases }] }`, broadcast `{ type: 'office-component', id, state, progress }`; aliases `pdf-cmaps`/`cjk` → `cjk-fonts`, `ocr-deu`/`ocr-eng` → `ocr`. Files: ocr `tesseract.min.js`, `worker.min.js`, `core/tesseract-core-{simd-,}lstm{.wasm.js,.js,.wasm}`, `lang/{deu,eng}.traineddata` (plain, use `gzip: false`); mermaid `mermaid.min.js`; cjk-fonts `cmaps/*.bcmap`, `standard_fonts/*` (pdf.js `cMapUrl`/`standardFontDataUrl` = `<base>cmaps/` / `<base>standard_fonts/`). ETag + `immutable` |
| `GET /office/search/meaning?q=&dir=` · `GET /office/search?q=&semantic=1` | `{ hits: [Entry + snippet, score], mode: embeddings\|ai-terms\|fulltext, … }` |
| `GET /office/encrypt/state?loc` · `POST /office/encrypt {loc, password \| vault: true, saveToVault?, mode: inplace\|copy}` · `POST /office/decrypt {loc, password \| fromVault: true, mode: bytes (default – raw response, file stays encrypted) \| inplace \| copy}` · `PUT /office/encrypt/raw?loc&mtime=` (raw plain bytes; header `x-office-password` or vault) | `{ encrypted: ooxml\|odf\|null, supported, vault: { available, locked, hasEntry } }` · `{ ok, mtime, size, version, vaultSaved, generated, versionsWithPlainText }` (in place keeps a "Before encrypting" version – offer deleting old plain versions) · 403 wrong password – with `soft: true` (1.2, sent by the editor host) 200 `{ ok: false, reason: 'password' \| 'vault', error }` instead, so a wrong password is no red console error · PDF → 501 |
| `GET /office/vault` · `POST /office/vault/encrypt {loc}` · `POST /office/vault/decrypt {loc}` | `{ available, unlocked }` · Entry + `result` (generated password in the passwords extension's `passwordVault`, source `office`; 409 no vault, 423 locked) |
| `POST /office/clean/preview {loc}` · `POST /office/clean {loc, remove: [ids], copy}` or `{loc, options: { id: false … }, mode: copy\|inplace\|download}` | `{ items: [{ id, label, count }], warnings: [{ id, label }] }` (ids author, comments, revisions, hidden, custom, thumbnail, rsid, exif, templates) · Entry + `report` (shell shape) / `{ ok, report, entry \| mtime }` |
| `GET /office/publish/targets` · `GET /office/publish?loc` · `POST /office/publish {loc, target: id \| { web, vhost }, sub \| path, title, files?: { 'index.html': string, … ≤ 12, ≤ 2 MB, html/css/js/svg/json/txt }, auto}` · `PUT /office/publish/:id {files, auto}` · `DELETE /office/publish/:id` | `{ targets }` (`host.sites.targets()`) · `{ pages, publications }` with `url`, `site`, `stale` · without `files` a plain text page is published; after each save a page with `auto` is marked stale and `{ type: 'office-publish', id, path }` is broadcast – the open editor re-renders and PUTs (images as inline data: URIs – `host.sites` only takes text files) |
| `GET /office/reports?loc` · `POST /office/reports {loc, schedule: { every: hourly\|daily\|weekly\|monthly (day/week/month too), time \| at: 'HH:MM', weekday, day }, format: pdf\|docx\|odt\|xlsx\|ods\|csv\|pptx\|odp\|same, recipients: [{ channel: hostwerk\|alerts\|mail\|discord\|ntfy, to }], saveCopy, title, on}` (`id` = update; `channel` alone still works) · `PUT/DELETE /office/reports/:id` · `POST …/:id/run` | (1.2) the file in its format (as it is, else through the converter), sent to EVERY recipient on its own: `hostwerk` notification, `alerts` = all alert channels, `mail` = up to 20 addresses (`to` empty = the alert address; with the attachment through Hostwerk's mail relay, without one into Hostwerk's mail catcher), `discord` = a webhook address, `ntfy` = an address or a topic (on the alert ntfy server, else ntfy.sh) – invalid targets → 400; `run` → `{ ok, results: [{ channel, to, ok, error }], error }` (one failure does not stop the others; `lastResults`/`lastError` kept and shown); PDF copy `<name> <date>.pdf` next to the file; own minute timer, missed runs run once |
| `GET /office/library-backups` · `POST /office/library-backups` · `POST …/restore {id, paths?}` · `GET /office/backups?loc` · `GET /office/backups/raw?loc&id=` · `POST /office/backups/restore {loc, id}` | daily zip snapshot of the library (≤ 2 GB, keep 7) in Hostwerk's backup center ("Office library"); restore never deletes, overwritten files keep their content as a version |
| `PUT /office/thumb?loc&mtime=` (raw PNG or WebP ≤ 2 MB) · `GET /office/thumb?loc` | per file key + mtime; GET: embedded first, then the rendered one (`x-thumb: rendered`), ETag/304; `Entry.thumb` true for both |
| `GET /office/qr?text=` · `POST /office/ai {task?, prompt, text?, system?, json?, lang}` | `image/svg+xml` · `{ text, json? }` (ctx.ai: quick tasks or `chats.oneShot` with the user's provider) |
| notifications | `host.notifications` event types `rule`, `form`, `report` (each with its switch in Hostwerk's settings) |

**Server 1.1 limits:** merged into "Known limits" below.

### Integration 1.1 (integrator's decisions – the round's parts put together)

**Package, size, CI**
- **What counts is the shipped package.** The catalog builds its entry from the release zip (`hostwerk-public/scripts/
  build-extension-catalog.js`: download → unzip → `sdk/check.js --strict` on the unpacked folder) and Hostwerk runs
  `checkPackage` on the unpacked package on install. The zip is made by the CI package step (`git ls-files` minus `.github/`,
  `tests/`, `tools/`, `scripts/`, `ui/src/`, `node_modules/`, `package*.json`, `docs/*.md`, `ARCHITECTURE.md`). So the
  repo CI now copies exactly those files into `$RUNNER_TEMP/pkg` and runs `check.js --strict` there (same filter as the
  package step – keep both in sync with `tools/package.js`). 1.1.0: **7.47 MB unpacked, 2.83 MB zip, 73 files** (budget
  9.8 MB, limit 10 MB); the full checkout is 11.75 MB (ui/src 3.8 MB, tests 0.4 MB) and fails `check.js .` on size only.
- **Consequence:** installing from the *source folder* (developer mode "Ordner", HW_EXT_SOURCE, Hostwerk's
  `scripts/ext-native-test.js`) runs the same size check on the whole checkout and refuses it. Developers install the
  package folder (`node tools/package.js --dir <folder>`). The clean fix is in the core (size check without the paths the
  release package excludes, or ext-native-test on the packaged files) – left to the core/user (not touched here).
- **Dedupe check** (`node tools/build.js --analyze`, esbuild metafile): only small duplicates are left – lucide-react
  icons spread over the chunks (≈ 100 KB in total), dompurify in docs + viewer (30 KB), diff-match-patch in pdf + docs-math
  (19 KB); jszip, pdf.js and yjs are shared chunks, formula/numfmt code lives only in sheets. Not worth another shared chunk
  at 76 % of the budget. CSS is minified by esbuild; the two PNG screenshots stay PNG (README tooling expects them).
- **Manifest:** version 1.1.0, `data/office-components` added, **`sdkVersion` stays "3.15"** – nothing on the server needs
  3.16 (`host.*` members used are all in 3.15), the SDK 3.16 parts are UI slots only (`files.icon`, `files.thumb`), which
  older Hostwerk versions ignore. Released Hostwerk 3.2.2 installs 1.1.0; Files' coloured icons and thumbnails appear with
  Hostwerk 3.2.3 / SDK 3.16. `files.icon` matches only the families Office opens (documents, spreadsheets, presentations,
  PDF, mail, contacts, calendars, e-books, diagrams, OneNote, Publisher, Project, Google, iWork) – text, code, images,
  archives, databases and folders keep Hostwerk's own icons.
- **Update safety:** an update activates the new module while the old one shuts down; the old `deactivate()` used to stop
  the *new* activation's state and remove the new "Office documents" root (seen live). Now every activation keeps its
  own state (WeakMap by host), shutdown runs once, and 3 s after activation the root is added again (covers updates from
  1.0.x, whose old code still removes it).

**Translations:** `locales/parts/*.json` (server, shell, docs, sheets, slides, pdfview – 1102 new keys) merged into
`locales/de.json` (3,928 keys) and deleted. Duplicates with different German resolved to the Office/Hostwerk wording: After
→ Nachher, Collapse/Show the ribbon → Menüband reduzieren/anzeigen, Word count → Wörter zählen, Flowchart →
Flussdiagramm, Try again → Erneut versuchen, Comment → Kommentar, Area → Fläche, Show → Anzeigen, Unlock → Entsperren,
Bold → Fett, Stop → Stopp, Location → Ort, Symbols → Symbole, Table style options → Tabellenformatoptionen, Shape fill →
Fülleffekt, Add to/Remove from favourites → Zu den Favoriten / Aus den Favoriten entfernen, Message → Meldung, Not now →
Jetzt nicht, Fly in → Hineinfliegen, Process → Prozess, Icons → Piktogramme, Click to add text → Text durch Klicken
hinzufügen, the msg/RTF sentence as before; "Kennwort" → "Passwort" everywhere (Hostwerk's word), "Backup" → "Sicherung"
(glossary). Where the meaning differed the key was changed in the code: PDF permission "Add comments" (Kommentieren),
PDF redaction mark "Marked area" (Markierter Bereich), signature pen "Thick" (Kräftig), sheets "Stop co-editing",
sparkline "Location range" (Positionsbereich), docs field "Number of words" (Anzahl der Wörter). `i18n-scan` lists ≈ 850
"texts" – code, XML, CSS, font/theme names, AI prompts (sent to the AI, not shown), German function-argument names and the
built-in templates' own German; the real UI texts it found without German were added.

**Cross-part fixes**
- **Co-editing:** the provider (`lib/collab.js`) now reads the server's control channel (type 100): `hello` (leader,
  `fresh`, the server-side user name), `leader`, `peers`, `saved`, `reset`; `reportSaved()`. The editor host learns every
  save of the room (`onRemoteSaved` → its `mtime`), so a co-editor that later becomes the saver gets no false "changed on
  disk" conflict. Docs seed the shared document only in a fresh room (server `hello.fresh`), otherwise wait for the content
  (no duplicated text when two open at once). Co-editors without a browser name get the server's user name instead of
  "Guest 23". Sheets' auto-join read `peers` as a number (the server sends a list) – fixed. Live: two isolated browser
  contexts converge for docs (y-prosemirror, both edits in the saved docx, the second editor stays clean), sheets (cells +
  formulas both ways) and slides (title text). Login-free co-editing via a share link needs a core decision (public `/ws`
  upgrade) – co-editing is for logged-in users.
- **Remote control:** the slides chunk sent `{ idx, pos, … }` as `slide` value (the server reads `index/slide`) – the
  phone always showed slide 1. It now sends the server's `state` shape (position among shown slides, total, title,
  notes, black, timer with elapsed time); `blank`/`first` commands are mapped. Live: phone page (no login, SSE + POST) ↔
  presenter, prev/next both ways, the phone mirrors the slide.
- **Components:** `host.components.ensure` matches ids *and* aliases (`ocr-deu` → `ocr`, `pdf-cmaps` → `cjk-fonts`),
  treats `state: 'ready'` as installed and returns `{ id, base, files }`; the component dialog reads the server's numeric
  progress, shows the download size and the install error. The server lists `files` as names (`fileCount` = number).
  `cjk-fonts` now also brings static TrueType **Noto Sans SC/JP/KR** (weight 400, Google Fonts' versioned files – the
  google/fonts repo only has variable fonts whose default instance is Thin) + OFL, all SHA-256-pinned (download ≈ 33 MB);
  `tests/pdf-cjk-font.test.js` (runs with `OFFICE_CJK_FONTS=<folder>`) embeds Chinese, Japanese and Korean annotation text
  as a subset (copyable). Live: all three components installed and verified in the throwaway instance.
- **Signatures:** `GET/PUT /office/signatures` `{ signatures: [{ url, w, h, at }] }` – per Hostwerk user in
  `office-data/signatures.json`, at most 8, PNG/JPEG/WebP data URLs ≤ 512 KB each, 2 MB together.
- **Encrypted Office files:** the editor host checks after reading (OOXML ext + compound file, or any ODF file) with
  `GET /office/encrypt/state`; encrypted → vault first (when it holds this file's password), else a password dialog (wrong
  password asks again), `POST /office/decrypt` (mode bytes – the file stays encrypted), a calm banner; saving goes through
  `PUT /office/encrypt/raw` with the same password (or the vault), "keep my version" in a conflict writes the copy
  encrypted too, and no thumbnail is made of an encrypted file. Live: a docx encrypted by the server and an xlsx encrypted
  by real Excel 16 open, edit, save, stay encrypted (and open in Word/Excel with the password – see the Office check).
- **Toast error:** Hostwerk's translating `toast.error` (core `extNative.js`) copies errors with `Object.assign(Object.create
  (proto), e, { message })`, which throws "Cannot set property message … which has only a getter" for DOMExceptions
  (AbortError, NotAllowedError …). `ui/src/lib/toastfix.js` (imported first by the main module; chunks share the same toast
  object) turns such errors into plain `Error`s first. (A core fix would be `new Error(tt(e.message))`.)
- **Thumbnails:** the queue asks `GET /office/thumb?…&probe=1` – the server answers 204 when there is none (was a 404 in the
  console for every new file). Text-only documents render with visible text (checked by pixels).
- **Palette:** ranking also uses the chunks' `words` (docs' "/" commands).
- Compare view of documents shows objects (diagram, field, formula, picture) as ◆ instead of a replacement box.
- Hostwerk UI lint: the xlsx table part's XML root is no HTML table (built via a constant), slicer corners use `--r-sm`.
- Obsolete `ui/src/editors/docs-collab/` deleted.

**Live test (throwaway Hostwerk 3.2.2 working tree, package installed from the release zip, 2026-11-20, Chrome
headless, two isolated contexts where needed):** create/edit/save docs, sheets, slides, PDF · command palette · thumbnails
(grid in Files) · versions + compare · clean (copy) · encrypt/decrypt (server + Excel-made) · publish as web page (test web
folder) · scheduled report "run now" (PDF via Microsoft Office COM, copy next to the file) · rules (test → met) · forms
(public page in a second client, row arrives in the sheet) · live data table in a sheet · SQLite table as sheet (read,
write-back with confirm via the route) · live field, LaTeX formula and Mermaid diagram in a doc (saved as field, OMML,
picture) · co-editing docs/sheets/slides · remote control · OCR (scanned page → searchable text) · redaction (e-mail gone
from the file, other text kept) · light/dark, de/en, 390/1440 px without horizontal scrolling · console free of errors
except expected ones (403 for a wrong file password, 404 of Hostwerk's own `/api/mailservers`).

### Office check 1.1 (real Microsoft Office 16, German, 2026-11-20 – combined final run)

`../../office-check/final-1.1-2026-11-20/` (copy of the docs11 harness, `check.ps1` + password support): sheets files
regenerated with the current code (`sheets-1.1/gen.mjs`), the docs and slides agents' 1.1 files (`docs11-2026-11-20/src`,
`c11/files`) and files saved in the live test (co-edited doc, doc with diagram + field + formula, form rows, live data +
database sheet, three-slide deck, and **encrypted** files: docx/xlsx/pptx encrypted by our server, an xlsx encrypted by
real Excel and re-saved by our editor). One file per fresh Office process, no Office running before (the driver refuses
otherwise) and none left after (only processes it started are killed – Word here stays open after Quit). `node
scripts/drive.mjs results/jobs.json results/main.jsonl - --pdf --wordpdf` and `… jobs-strict.json strict.jsonl - --strict`.

| Group | Files | Open (no repair) | Strict open (alerts on) | Facts |
|---|---|---|---|---|
| Sheets 1.1 (tables + structured refs, pivot, objects/sparklines/shapes, 1904 dates, chart + outline, LAMBDA family) + round trips | 12 | 12/12 | 12/12 | 46/46 checked values match; `lambda*.xlsx` 0/17 – this Excel build has no LAMBDA/LET/MAP (`#NAME?` also for formulas typed into it, sheets-1.1 probe) – formulas are kept as written |
| Docs 1.1 (sections, headers, fields, OMML equations, content controls, compare result) docx/dotx/odt + round trips | 6 | 6/6 | 6/6 | 3 fields, 3 equations, 1 picture, 6 content controls (docx) as written |
| Slides 1.1 (chart with workbook, morph, animations, custom geometry) pptx/pptm/odp + round trip | 4 | 4/4 | 4/4 | 4 slides, titles match |
| Live-test files (co-edit, diagram/field/formula, forms, live data + DB sheet, decks) | 6 | 6/6 | 6/6 | text of both co-editors; 1 field, 1 equation, 1 picture; sheets Tabelle1 + kunden |
| Encrypted (server: docx, xlsx, pptx; Excel-made xlsx re-saved by our editor) – opened with the password | 4 | 4/4 | 4/4 | content incl. the edit made in our editor ("VERSCHLUESSELT") |

All 32 files: opened, nothing repaired, every PDF export worked (32/32).


### 1.2 round – shell / server / docs (for the integrator)

- **Co-editing (docs):** the saver is the SERVER's leader (`provider.leader` from `hello`/`leader`) – the same for everyone at
  once; the old "smallest client id" rule only remains for servers without the control channel. A save of the room
  (`saved`) makes a co-editor clean (`host.dirty(false)`); becoming the leader with changes the room has not saved yet makes
  that editor dirty (it saves them). Live: follower stays "Keine Änderungen" while typing, the leader saves, hand-over saves.
- **Docs palette:** `ui/src/editors/docs/palette.js` (formatting, styles, paragraph, insert, layout, review, view/zoom – English
  keys, shown localised by the palette) + the "/" commands; `words` go to the palette ranking (`tests/docs-palette.test.js`).
- **Floating text boxes:** `ui/src/editors/docs/io/anchor.js` (read/move `wp:anchor`, pure) + the `preserved` node view
  (`floatingView` in `extensions/nodes.js`); `Preserved` storage `onEdit` opens the text dialog on double-click.
- **Templates view:** a click on an own template asks for the file name and creates the file in the open folder
  (`newFromOwnTemplate` in `shell/ops.js`); "More templates and formats …" in its menu opens the New dialog.
- **Server:** `sheet.js` date values (`{ $date }`), `cfb.xlsText` (BIFF8), ZIP64 in `zip.js` (`centralDir`, record writers
  shared with `zippatch.js`), `server/vendor/yjs.cjs` (+ `LICENSE-yjs.txt`; loaded with the Yjs "already imported" marker set
  aside), report recipients (`extras.js` `normRecipients`, `notify.js` `send(…, { to })`), `clean.js` `joinDeletedMarks`,
  `POST /office/decrypt {soft}`. The search index reads `.xls` files of older indexes once more (entry `v: 2`).
- **Texts:** new German in `locales/parts/a.json` (36 keys) – to be merged into `de.json` (the package step would ship the part file).

## Known limits

**Not verified for real** (no instance / software on the build machine): Electron-only paths (in-app web window,
presenter view as a Hostwerk window, Google sign-in, the packaged app's `--open-file` start), the ONLYOFFICE container
(Docker not available; `host.docker.internal` must reach Hostwerk – with Hostwerk's own WSL engine set
`integrations.hostUrl` to the LAN address if not), Visio conversion, engine installs on Linux (deb extract) and macOS (dmg),
real-world .doc/.ppt files (only synthetic ones in the tests), .msg (no sample), ODF encryption against LibreOffice (Microsoft
Office opens no password-protected ODF at all), the PDF "Compare" button (file picker – the diff itself is unit-tested),
MariaDB/PostgreSQL/MongoDB as sheets (only SQLite was live-tested), the vault as password source (passwords extension not
installed in the throwaway instance), real notification channels (Discord, ntfy, mail).

**Package / core:** installing from the source folder (developer mode, `HW_EXT_SOURCE`, Hostwerk's
`scripts/ext-native-test.js`) fails Hostwerk's 10 MB size check because the checkout carries ui/src and tests (11.75 MB) –
install `node tools/package.js --dir` instead; a core change is needed for the source folder (see "Integration 1.1").
Co-editing needs a Hostwerk login (no public websocket upgrade in the core; the phone remote works without login).
Hostwerk's translating `toast.error` breaks on DOMExceptions (worked around in `lib/toastfix.js`).

**Shell:** "Download as" uses `<a download>` – Hostwerk offers extensions no save dialog for binary files (`pickPath` only opens
files/folders, `Hostwerk.saveFile` takes text ≤ 5 MB); in the Electron app such a download shows the system's "Save as" dialog
anyway (Electron's default – Hostwerk sets no `will-download` path; not verified live, no Electron here). Dashboard tile only; a
lock left by a crashed browser shows "open" for up to 5 minutes; thumbnails of a one-line document look nearly empty (they are
true to the page).

**Backend:** wps/wpd/wri/abw/lwp/hwp text only via the engine; iWork text is a heuristic; .xls text = BIFF8/BIFF5 cell texts,
numbers and cached formula results (encrypted workbooks and BIFF2–4 give none); office files pass the raw routes up to 200 MB –
ZIP64 is read and written (zip downloads, library snapshots, writers that use ZIP64 records on small files), but a > 4 GB office
file never gets that far; the private engine needs ~1.5 GB on disk (download ≈ 360 MB); Google Drive needs the user's own OAuth
client. PDF passwords are set by the PDF editor itself (AES-256) – the server's `/office/encrypt` answers 501 for PDF; PDF
cleaning blanks Info/XMP in place (uncompressed parts only – object streams are reported); hidden text = Word runs with
`w:vanish` and ODF hidden-text fields; accepting a deleted paragraph mark joins direct siblings only (a mark inside a text box,
or one whose `w:pPrChange` holds a full `w:pPr`, just loses the mark); a co-editing room whose COMPACTED state is still over
48 MB asks the leader to save and everyone to reload; the time machine only snapshots Hostwerk's settings – the library goes to
the backup center (own daily zip, ≤ 2 GB); MongoDB writes are not atomic; report formats other than the file's own need a
converter, report e-mails without Hostwerk's mail relay land in Hostwerk's mail catcher; ODS form answers' date columns get no
width (Excel shows ##### until the column is widened – xlsx gets one).

**Docs:** print layout breaks long paragraphs/lists between lines and long tables between rows (header row repeated) – only a
single line, a single table row or a picture taller than a page runs over the gap; columns: print layout shows the text in
columns on ONE sheet as long as the text (CSS columns cannot break into pages inside one editing flow – exports, print and PDF
are exact); floating text boxes/shapes (`wp:anchor`, docx) float at their horizontal position (wrap square / top-and-bottom /
none) and move by drag (written back as offsets from column/paragraph), but positions measured from the page or margin TOP keep
the box at its paragraph, odt frames do not move, and charts/SmartArt/OLE stay inline; Word 6/95 and encrypted *.doc* (binary)
are not read (the engine opens 6/95); PDF export without the engine uses the print dialog; diagrams are pictures with their
source in the alt text (Word shows the picture); seen once in the 1.0 live test, not reproducible: after inserting a table a
mouse double-click selection was not taken over.

**Sheets (checked against code + tests for 1.2):** form controls of xlsx files are shown; check boxes, option groups,
drop-downs, list boxes and spinners work and write their linked cell (ctrlProp + VML patched on save, checked in Excel);
buttons are shown but run their macros in Excel only, scroll bars behave like spinners, ActiveX controls are kept, not
shown. What-if data tables keep their `TABLE()` formula and their results are recalculated (one- and two-variable;
not for tables over 5,000 cells). External links and Power Query are kept but not editable. Edited Excel charts are
patched into the original chart part (title, axis titles, legend, series ranges/names/colours/count, stacking, type
between column/bar/line/area – Excel's styles, layout and extensions stay); other edits (labels or trendline switched,
pie/doughnut/scatter ↔ others, a range-based chart) still rewrite the chart in the simple form, which keeps labels and
trendlines but not Excel's formatting. ODS/OTS/FODS: own reader and writer with cell styles, number formats as data
styles, merges, sizes, outline groups, frozen panes, conditional formats (cell/formula rules as `style:map`, all rules
as LibreOffice `calcext`), validation, comments, links, rich text, names, autofilter, print setup, images and charts
(chart objects); unknown parts and shapes are kept. Excel ignores frozen panes in ODS (it reads no settings.xml) and
opens no flat .fods at all; colour scales, data bars and icon sets in ODS are LibreOffice-only (calcext). XLS: cell
formatting (fonts, fills in the file's palette colours, borders, alignment, rotation, number formats, protection) is read
by our own BIFF8 reader; charts, pictures and conditional formats of .xls are not read without the engine, and writing
.xls without the engine gives values only (SheetJS CE). XLSB: values, formulas and number formats only (no styles)
without the engine. Function names exist in English and German only – French/Spanish/Italian were not added because
only a German Excel is available here to verify Microsoft's names. The settings gridlines/headings/font are the
defaults of new sheets (a file keeps its own).

**Slides:** the outline view edits titles and every text placeholder/text box of a slide, Tab/Shift+Tab change levels,
drag reorders (verified in Extras.jsx). ppt: tables arrive as real tables when the group carries the table properties
(PowerPoint 2000+), old build effects become entrance animations (fly with direction, wipe, fade, zoom, split, wheel,
appear; order and "after previous") – effects that only exist in the PowerPoint 2002+ timing records come out as the
old fallback the file stores (a Fade is stored as "Cut" → Appear); no comments, sounds or encrypted ppt. odp: no
form-control bindings, no change tracking. Printing sends the vector PDF (no pictures, no `img-src data:`); PNG export
uses the SVG picture and falls back to the vector PDF painted by pdf.js; objects without a vector form (charts, media)
still need the SVG picture inside it. Text uses the system fonts.

**PDF + viewer:** OCR: German + English in the engine, French, Spanish, Italian, Dutch, Polish and Turkish as optional
language packs (`ocr-fra` … `ocr-tur` – the server registry has to list them). Redaction keeps a page as vectors when it
has no pictures, form XObjects, inline images or Type3/unknown-width fonts (marked glyphs and paths are removed from the
content stream, the marks are painted black); any other page, or a page whose check fails, becomes a picture with an
invisible text layer. Own annotations stay editable after another app re-saved the file as long as it kept our
`/HWOffice` key (also inside compressed object streams, the catalog marker may be gone); moved ones follow, a changed
comment text is taken over, resized ones become ordinary annotations, and apps that drop unknown keys leave ordinary
annotations. Access .mdb (Jet 4) and .accdb (ACE) show their tables and rows (memo, currency, dates, Yes/No; "Open as
spreadsheet"); encrypted databases, query SQL, OLE/attachment contents are not shown; the Jet 3 (Access 97) path is
implemented but untested (no such file could be made here). vsd/pub/mpp need the engine (text found in the file only);
OneNote shows the text found in the file (no revision-store parser). vsdx: connectors saved by Visio use their stored
route; connectors without own geometry (other tools) are routed at right angles around the shapes (no line jumps).

**Bundles:** jszip's `setImmediate` polyfill and pdf.js' feature test contain `new Function` (never reached with our
options; no `eval(` anywhere – the catalog check passes). 1.1.0 chunks (js + css): main 320 KB, docs 1041 KB (+ docs-math
225 KB lazy), sheets 1302 KB, slides 620 KB (+ slides-extra 63 KB lazy), pdf 725 KB + worker 1293 KB, viewer 212 KB,
shared lib-jszip 96 KB, lib-pdfjs 324 KB, lib-yjs 98 KB; shipped package 7.47 MB of 10 MB (six screenshots 0.58 MB).
