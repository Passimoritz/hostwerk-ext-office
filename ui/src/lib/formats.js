/*
 * Every file type the Office extension knows – one list for the library, the file browser slots, the "New" menu,
 * the converter and the editors. Keep it in sync with server/formats.js (same ext → kind mapping).
 *
 *   kind     which editor chunk opens it: docs | sheets | slides | pdf | viewer
 *   family   ms | odf | google | apple | open | other    (filter chips and icons)
 *   edit     'native'  – the chunk reads AND writes this format itself
 *            'import'  – the chunk reads it; saving writes the nearest native format (asks first) or converts back
 *            'convert' – needs LibreOffice / Microsoft Office on this PC to open (converted to `via` first)
 *            'view'    – read only
 *   via      intermediate format for 'convert' (and the target of "Save as editable copy")
 */
const F = (kind, family, edit, label, mime, via) => ({ kind, family, edit, label, mime, via });

export const FORMATS = {
  // ---------------- Documents ----------------
  docx: F('docs', 'ms', 'native', 'Word document', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
  docm: F('docs', 'ms', 'import', 'Word document with macros', 'application/vnd.ms-word.document.macroEnabled.12'),
  dotx: F('docs', 'ms', 'import', 'Word template', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template'),
  dotm: F('docs', 'ms', 'import', 'Word template with macros', 'application/vnd.ms-word.template.macroEnabled.12'),
  doc: F('docs', 'ms', 'convert', 'Word 97–2003 document', 'application/msword', 'docx'),
  dot: F('docs', 'ms', 'convert', 'Word 97–2003 template', 'application/msword', 'docx'),
  odt: F('docs', 'odf', 'native', 'OpenDocument text', 'application/vnd.oasis.opendocument.text'),
  ott: F('docs', 'odf', 'import', 'OpenDocument text template', 'application/vnd.oasis.opendocument.text-template'),
  fodt: F('docs', 'odf', 'import', 'Flat OpenDocument text', 'application/vnd.oasis.opendocument.text-flat-xml'),
  sxw: F('docs', 'odf', 'convert', 'OpenOffice.org 1 text', 'application/vnd.sun.xml.writer', 'odt'),
  rtf: F('docs', 'open', 'native', 'Rich Text', 'application/rtf'),
  txt: F('docs', 'open', 'native', 'Plain text', 'text/plain'),
  md: F('docs', 'open', 'native', 'Markdown', 'text/markdown'),
  markdown: F('docs', 'open', 'native', 'Markdown', 'text/markdown'),
  html: F('docs', 'open', 'native', 'Web page', 'text/html'),
  htm: F('docs', 'open', 'native', 'Web page', 'text/html'),
  xhtml: F('docs', 'open', 'import', 'XHTML page', 'application/xhtml+xml'),
  wpd: F('docs', 'other', 'convert', 'WordPerfect document', 'application/vnd.wordperfect', 'odt'),
  wps: F('docs', 'other', 'convert', 'Works / WPS document', 'application/vnd.ms-works', 'docx'),
  wri: F('docs', 'other', 'convert', 'Windows Write document', 'application/x-mswrite', 'docx'),
  pages: F('docs', 'apple', 'convert', 'Apple Pages document', 'application/vnd.apple.pages', 'docx'),
  abw: F('docs', 'open', 'convert', 'AbiWord document', 'application/x-abiword', 'odt'),
  lwp: F('docs', 'other', 'convert', 'Lotus Word Pro document', 'application/vnd.lotus-wordpro', 'odt'),
  hwp: F('docs', 'other', 'convert', 'Hangul document', 'application/x-hwp', 'odt'),
  gdoc: F('viewer', 'google', 'view', 'Google Docs link', 'application/vnd.google-apps.document'),

  // ---------------- Spreadsheets ----------------
  xlsx: F('sheets', 'ms', 'native', 'Excel workbook', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
  xlsm: F('sheets', 'ms', 'import', 'Excel workbook with macros', 'application/vnd.ms-excel.sheet.macroEnabled.12'),
  xlsb: F('sheets', 'ms', 'import', 'Excel binary workbook', 'application/vnd.ms-excel.sheet.binary.macroEnabled.12'),
  xltx: F('sheets', 'ms', 'import', 'Excel template', 'application/vnd.openxmlformats-officedocument.spreadsheetml.template'),
  xltm: F('sheets', 'ms', 'import', 'Excel template with macros', 'application/vnd.ms-excel.template.macroEnabled.12'),
  xls: F('sheets', 'ms', 'native', 'Excel 97–2003 workbook', 'application/vnd.ms-excel'),
  xlt: F('sheets', 'ms', 'import', 'Excel 97–2003 template', 'application/vnd.ms-excel'),
  ods: F('sheets', 'odf', 'native', 'OpenDocument spreadsheet', 'application/vnd.oasis.opendocument.spreadsheet'),
  ots: F('sheets', 'odf', 'import', 'OpenDocument spreadsheet template', 'application/vnd.oasis.opendocument.spreadsheet-template'),
  fods: F('sheets', 'odf', 'native', 'Flat OpenDocument spreadsheet', 'application/vnd.oasis.opendocument.spreadsheet-flat-xml'),
  sxc: F('sheets', 'odf', 'import', 'OpenOffice.org 1 spreadsheet', 'application/vnd.sun.xml.calc'),
  csv: F('sheets', 'open', 'native', 'CSV table', 'text/csv'),
  tsv: F('sheets', 'open', 'native', 'Tab-separated table', 'text/tab-separated-values'),
  tab: F('sheets', 'open', 'import', 'Tab-separated table', 'text/tab-separated-values'),
  dif: F('sheets', 'open', 'native', 'Data Interchange Format', 'text/x-dif'),
  slk: F('sheets', 'open', 'native', 'SYLK table', 'text/x-sylk'),
  prn: F('sheets', 'open', 'import', 'Lotus formatted text', 'text/plain'),
  dbf: F('sheets', 'open', 'native', 'dBASE table', 'application/x-dbf'),
  numbers: F('sheets', 'apple', 'import', 'Apple Numbers spreadsheet', 'application/vnd.apple.numbers'),
  wk1: F('sheets', 'other', 'import', 'Lotus 1-2-3 sheet', 'application/vnd.lotus-1-2-3'),
  wks: F('sheets', 'other', 'import', 'Lotus / Works sheet', 'application/vnd.lotus-1-2-3'),
  qpw: F('sheets', 'other', 'import', 'Quattro Pro sheet', 'application/x-quattropro'),
  et: F('sheets', 'other', 'convert', 'WPS spreadsheet', 'application/vnd.ms-excel', 'xlsx'),
  gsheet: F('viewer', 'google', 'view', 'Google Sheets link', 'application/vnd.google-apps.spreadsheet'),

  // ---------------- Presentations ----------------
  pptx: F('slides', 'ms', 'native', 'PowerPoint presentation', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
  pptm: F('slides', 'ms', 'import', 'PowerPoint presentation with macros', 'application/vnd.ms-powerpoint.presentation.macroEnabled.12'),
  ppsx: F('slides', 'ms', 'import', 'PowerPoint show', 'application/vnd.openxmlformats-officedocument.presentationml.slideshow'),
  ppsm: F('slides', 'ms', 'import', 'PowerPoint show with macros', 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12'),
  potx: F('slides', 'ms', 'import', 'PowerPoint template', 'application/vnd.openxmlformats-officedocument.presentationml.template'),
  potm: F('slides', 'ms', 'import', 'PowerPoint template with macros', 'application/vnd.ms-powerpoint.template.macroEnabled.12'),
  ppt: F('slides', 'ms', 'convert', 'PowerPoint 97–2003 presentation', 'application/vnd.ms-powerpoint', 'pptx'),
  pps: F('slides', 'ms', 'convert', 'PowerPoint 97–2003 show', 'application/vnd.ms-powerpoint', 'pptx'),
  pot: F('slides', 'ms', 'convert', 'PowerPoint 97–2003 template', 'application/vnd.ms-powerpoint', 'pptx'),
  odp: F('slides', 'odf', 'native', 'OpenDocument presentation', 'application/vnd.oasis.opendocument.presentation'),
  otp: F('slides', 'odf', 'import', 'OpenDocument presentation template', 'application/vnd.oasis.opendocument.presentation-template'),
  fodp: F('slides', 'odf', 'import', 'Flat OpenDocument presentation', 'application/vnd.oasis.opendocument.presentation-flat-xml'),
  sxi: F('slides', 'odf', 'convert', 'OpenOffice.org 1 presentation', 'application/vnd.sun.xml.impress', 'odp'),
  key: F('slides', 'apple', 'convert', 'Apple Keynote presentation', 'application/vnd.apple.keynote', 'pptx'),
  dps: F('slides', 'other', 'convert', 'WPS presentation', 'application/vnd.ms-powerpoint', 'pptx'),
  gslides: F('viewer', 'google', 'view', 'Google Slides link', 'application/vnd.google-apps.presentation'),

  // ---------------- PDF and fixed layout ----------------
  pdf: F('pdf', 'open', 'native', 'PDF document', 'application/pdf'),
  xps: F('pdf', 'ms', 'convert', 'XPS document', 'application/vnd.ms-xpsdocument', 'pdf'),
  oxps: F('pdf', 'ms', 'convert', 'OpenXPS document', 'application/oxps', 'pdf'),

  // ---------------- Drawings, diagrams, publishing (viewer; edit via converter / desktop app) ----------------
  odg: F('viewer', 'odf', 'view', 'OpenDocument drawing', 'application/vnd.oasis.opendocument.graphics', 'pdf'),
  otg: F('viewer', 'odf', 'view', 'OpenDocument drawing template', 'application/vnd.oasis.opendocument.graphics-template', 'pdf'),
  fodg: F('viewer', 'odf', 'view', 'Flat OpenDocument drawing', 'application/vnd.oasis.opendocument.graphics-flat-xml', 'pdf'),
  odf: F('viewer', 'odf', 'view', 'OpenDocument formula', 'application/vnd.oasis.opendocument.formula', 'pdf'),
  odb: F('viewer', 'odf', 'view', 'OpenDocument database', 'application/vnd.oasis.opendocument.database'),
  vsdx: F('viewer', 'ms', 'view', 'Visio drawing', 'application/vnd.ms-visio.drawing', 'pdf'),
  vsdm: F('viewer', 'ms', 'view', 'Visio drawing with macros', 'application/vnd.ms-visio.drawing.macroEnabled.12', 'pdf'),
  vsd: F('viewer', 'ms', 'view', 'Visio 2003–2010 drawing', 'application/vnd.visio', 'pdf'),
  vstx: F('viewer', 'ms', 'view', 'Visio template', 'application/vnd.ms-visio.template', 'pdf'),
  pub: F('viewer', 'ms', 'view', 'Publisher document', 'application/x-mspublisher', 'pdf'),
  one: F('viewer', 'ms', 'view', 'OneNote section', 'application/onenote'),
  mpp: F('viewer', 'ms', 'view', 'Project plan', 'application/vnd.ms-project'),
  accdb: F('viewer', 'ms', 'view', 'Access database', 'application/msaccess'),
  mdb: F('viewer', 'ms', 'view', 'Access 97–2003 database', 'application/msaccess'),
  gdraw: F('viewer', 'google', 'view', 'Google Drawings link', 'application/vnd.google-apps.drawing'),
  gform: F('viewer', 'google', 'view', 'Google Forms link', 'application/vnd.google-apps.form'),
  gsite: F('viewer', 'google', 'view', 'Google Sites link', 'application/vnd.google-apps.site'),
  gmap: F('viewer', 'google', 'view', 'Google My Maps link', 'application/vnd.google-apps.map'),
  gjam: F('viewer', 'google', 'view', 'Google Jamboard link', 'application/vnd.google-apps.jam'),
  glink: F('viewer', 'google', 'view', 'Google Drive link', 'application/vnd.google-apps.shortcut'),

  // ---------------- Mail, contacts, calendar, e-books ----------------
  eml: F('viewer', 'open', 'view', 'E-mail', 'message/rfc822'),
  msg: F('viewer', 'ms', 'view', 'Outlook message', 'application/vnd.ms-outlook'),
  vcf: F('viewer', 'open', 'view', 'Contact card', 'text/vcard'),
  ics: F('viewer', 'open', 'view', 'Calendar', 'text/calendar'),
  epub: F('viewer', 'open', 'view', 'E-book', 'application/epub+zip'),

  // ---------------- Plain data that people open next to office files ----------------
  json: F('viewer', 'open', 'view', 'JSON data', 'application/json'),
  xml: F('viewer', 'open', 'view', 'XML data', 'application/xml'),
  svg: F('viewer', 'open', 'view', 'SVG image', 'image/svg+xml'),
  png: F('viewer', 'open', 'view', 'PNG image', 'image/png'),
  jpg: F('viewer', 'open', 'view', 'JPEG image', 'image/jpeg'),
  jpeg: F('viewer', 'open', 'view', 'JPEG image', 'image/jpeg'),
  gif: F('viewer', 'open', 'view', 'GIF image', 'image/gif'),
  webp: F('viewer', 'open', 'view', 'WebP image', 'image/webp'),
  bmp: F('viewer', 'open', 'view', 'Bitmap image', 'image/bmp'),
  zip: F('viewer', 'open', 'view', 'ZIP archive', 'application/zip'),
};

export const KINDS = {
  docs: { label: 'Document', plural: 'Documents', icon: 'FileText', newExt: { ms: 'docx', odf: 'odt' } },
  sheets: { label: 'Spreadsheet', plural: 'Spreadsheets', icon: 'Sheet', newExt: { ms: 'xlsx', odf: 'ods' } },
  slides: { label: 'Presentation', plural: 'Presentations', icon: 'Presentation', newExt: { ms: 'pptx', odf: 'odp' } },
  pdf: { label: 'PDF', plural: 'PDFs', icon: 'FileType', newExt: { ms: 'pdf', odf: 'pdf' } },
  viewer: { label: 'Other file', plural: 'Other files', icon: 'File' },
};

export const FAMILIES = { ms: 'Microsoft Office', odf: 'OpenDocument', google: 'Google', apple: 'Apple iWork', open: 'Open formats', other: 'Other' };

export const extOf = (name) => { const m = /\.([a-z0-9]+)$/i.exec(String(name || '')); return m ? m[1].toLowerCase() : ''; };
export const formatOf = (name) => FORMATS[extOf(name)] || null;
export const kindOf = (name) => (formatOf(name) || {}).kind || null;
export const isOffice = (name) => !!formatOf(name);
/** RegExp for Hostwerk's file browser slots (files.preview / files.actions) */
export const OFFICE_RE = new RegExp(`\\.(${Object.keys(FORMATS).filter((e) => !['json', 'xml', 'svg', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'zip'].includes(e)).join('|')})$`, 'i');

/** "Save as" / "Download as" targets per kind – which the chunk writes itself (client) and which need the converter (server). */
export const EXPORTS = {
  docs: { client: ['docx', 'odt', 'rtf', 'html', 'md', 'txt', 'pdf'], server: ['doc', 'pdf', 'epub', 'wps'] },
  sheets: { client: ['xlsx', 'xls', 'ods', 'fods', 'csv', 'tsv', 'html', 'dif', 'slk', 'pdf'], server: ['xlsb', 'pdf'] },
  slides: { client: ['pptx', 'odp', 'pdf', 'png'], server: ['ppt', 'pdf'] },
  pdf: { client: ['pdf', 'png', 'txt'], server: ['docx', 'odt'] },
};

/*
 * Built-in templates (gallery under "New"). The chunk of `kind` must implement create({ template: id }).
 * Texts are English – German comes from locales/de.json.
 */
export const TEMPLATES = [
  { id: 'blank', kind: 'docs', name: 'Blank document', category: 'Basics' },
  { id: 'letter', kind: 'docs', name: 'Letter (DIN 5008)', category: 'Letters' },
  { id: 'invoice', kind: 'docs', name: 'Invoice', category: 'Business' },
  { id: 'cv', kind: 'docs', name: 'CV / Résumé', category: 'Personal' },
  { id: 'report', kind: 'docs', name: 'Report with table of contents', category: 'Business' },
  { id: 'minutes', kind: 'docs', name: 'Meeting minutes', category: 'Business' },
  { id: 'notes', kind: 'docs', name: 'Notes with checklist', category: 'Personal' },
  { id: 'blank', kind: 'sheets', name: 'Blank spreadsheet', category: 'Basics' },
  { id: 'budget', kind: 'sheets', name: 'Household budget', category: 'Personal' },
  { id: 'invoice', kind: 'sheets', name: 'Invoice with totals', category: 'Business' },
  { id: 'schedule', kind: 'sheets', name: 'Weekly schedule', category: 'Personal' },
  { id: 'tracker', kind: 'sheets', name: 'Project tracker', category: 'Business' },
  { id: 'inventory', kind: 'sheets', name: 'Inventory list', category: 'Business' },
  { id: 'blank', kind: 'slides', name: 'Blank presentation', category: 'Basics' },
  { id: 'pitch', kind: 'slides', name: 'Pitch deck', category: 'Business' },
  { id: 'lesson', kind: 'slides', name: 'Lesson / talk', category: 'Education' },
  { id: 'photo', kind: 'slides', name: 'Photo album', category: 'Personal' },
  { id: 'blank', kind: 'pdf', name: 'Blank PDF', category: 'Basics' },
];
