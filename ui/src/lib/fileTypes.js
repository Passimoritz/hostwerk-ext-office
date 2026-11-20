/*
 * File type families for the coloured file icons (ui/src/lib/fileIcons.jsx) – pure, no React, unit-tested
 * (tests/shell-icons.test.js). One colour per app like Microsoft Office / LibreOffice: documents blue, spreadsheets
 * green, presentations orange, PDF red, everything else its own calm colour or neutral.
 *
 *   familyOf('Report.docx') → 'word'   ·   FAMILIES.word → { color, glyph, label }
 *   iconInfo('x.gsheet')    → { family: 'google', color, glyph: 'grid', badge: 'cloud', label }
 */
import { FORMATS, extOf } from './formats.js';

// colours: mid tones that read on white and on Hostwerk's dark surfaces (contrast of the white glyph ≥ 3:1)
export const FAMILIES = {
  word: { color: '#2b6fd6', glyph: 'lines', label: 'Document' },
  sheet: { color: '#14995a', glyph: 'grid', label: 'Spreadsheet' },
  slides: { color: '#e0662b', glyph: 'screen', label: 'Presentation' },
  pdf: { color: '#dc3545', glyph: 'sign', label: 'PDF' },
  text: { color: '#64708a', glyph: 'lines', label: 'Text' },
  code: { color: '#64708a', glyph: 'code', label: 'Web page' },
  mail: { color: '#1f7fc4', glyph: 'mail', label: 'E-mail' },
  contact: { color: '#0e8f86', glyph: 'person', label: 'Contact' },
  calendar: { color: '#7b5cd6', glyph: 'calendar', label: 'Calendar' },
  ebook: { color: '#b0761a', glyph: 'book', label: 'E-book' },
  diagram: { color: '#3d5bc4', glyph: 'flow', label: 'Diagram' },
  image: { color: '#a24bc0', glyph: 'image', label: 'Image' },
  archive: { color: '#8a6d55', glyph: 'zip', label: 'Archive' },
  database: { color: '#b2384a', glyph: 'db', label: 'Database' },
  note: { color: '#8a3fb8', glyph: 'note', label: 'Notebook' },
  publisher: { color: '#0d8a7a', glyph: 'layout', label: 'Publication' },
  project: { color: '#2f8a3a', glyph: 'gantt', label: 'Project' },
  google: { color: '#4b7be5', glyph: 'lines', label: 'Google file' },
  iwork: { color: '#ef8a1f', glyph: 'lines', label: 'Apple iWork' },
  folder: { color: '#d9a033', glyph: 'folder', label: 'Folder' },
  file: { color: '#7d8597', glyph: 'blank', label: 'File' },
};

const BY_EXT = {
  word: 'doc docx docm dot dotx dotm odt ott fodt sxw rtf wpd wps wri abw lwp hwp',
  sheet: 'xls xlsx xlsm xlsb xlt xltx xltm ods ots fods sxc csv tsv tab prn dbf et dif slk',
  slides: 'ppt pptx pptm pps ppsx ppsm pot potx potm odp otp fodp sxi dps',
  pdf: 'pdf',
  text: 'txt md markdown log ini',
  code: 'html htm xhtml xml json svg',
  mail: 'eml msg mbox emlx',
  contact: 'vcf vcard',
  calendar: 'ics ical ifb',
  ebook: 'epub mobi azw azw3 fb2 djvu',
  diagram: 'vsd vsdx vsdm vdx vss vssx vst vstx drawio odg otg fodg sxd',
  image: 'png jpg jpeg gif webp bmp tif tiff ico heic avif emf wmf',
  archive: 'zip 7z rar tar gz tgz bz2 xz',
  database: 'accdb mdb odb sqlite db',
  note: 'one onetoc2',
  publisher: 'pub',
  project: 'mpp mpt',
  google: 'gdoc gsheet gslides gdraw gform gsite gmap',
  iwork: 'pages numbers key',
};
const EXT_FAMILY = new Map();
for (const [fam, list] of Object.entries(BY_EXT)) for (const x of list.split(' ')) EXT_FAMILY.set(x, fam);

// the glyph of a Google or iWork file follows what it is, the colour follows its app
const KIND_GLYPH = { docs: 'lines', sheets: 'grid', slides: 'screen', pdf: 'sign' };
const GOOGLE = { gdoc: ['lines', '#3f7ee8'], gsheet: ['grid', '#1e9e5a'], gslides: ['screen', '#e8a417'], gdraw: ['flow', '#d9483b'], gform: ['lines', '#7248b9'], gsite: ['layout', '#3f51b5'], gmap: ['image', '#d9483b'] };
const IWORK = { pages: ['lines', '#ef8a1f'], numbers: ['grid', '#25a94a'], key: ['screen', '#2f8ae8'] };

/** family id of a file name, an extension or an entry ({ name, dir }) */
export function familyOf(x) {
  if (x && typeof x === 'object') { if (x.dir) return 'folder'; x = x.name || x.path || ''; }
  const s = String(x || '');
  const ext = s.includes('.') ? extOf(s) : s.toLowerCase();
  if (EXT_FAMILY.has(ext)) return EXT_FAMILY.get(ext);
  const k = (FORMATS[ext] || {}).kind;
  if (k === 'docs') return 'word';
  if (k === 'sheets') return 'sheet';
  if (k === 'slides') return 'slides';
  if (k === 'pdf') return 'pdf';
  return 'file';
}

/** everything an icon needs: { family, color, glyph, badge?, label, ext } */
export function iconInfo(x) {
  const fam = familyOf(x);
  const name = x && typeof x === 'object' ? x.name || x.path || '' : String(x || '');
  const ext = fam === 'folder' ? '' : (name.includes('.') ? extOf(name) : name.toLowerCase());
  const f = FAMILIES[fam] || FAMILIES.file;
  if (fam === 'google' && GOOGLE[ext]) return { family: fam, color: GOOGLE[ext][1], glyph: GOOGLE[ext][0], badge: 'cloud', label: f.label, ext };
  if (fam === 'iwork' && IWORK[ext]) return { family: fam, color: IWORK[ext][1], glyph: IWORK[ext][0], label: f.label, ext };
  return { family: fam, color: f.color, glyph: f.glyph || KIND_GLYPH[(FORMATS[ext] || {}).kind] || 'blank', label: f.label, ext };
}

/** colour of a kind (docs | sheets | slides | pdf | viewer) – app accents in the shell (top bar, tiles) */
export const KIND_COLOR = { docs: FAMILIES.word.color, sheets: FAMILIES.sheet.color, slides: FAMILIES.slides.color, pdf: FAMILIES.pdf.color, viewer: FAMILIES.file.color };
