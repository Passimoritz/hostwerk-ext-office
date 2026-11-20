/*
 * Viewer chunk (kind 'viewer'): read-only views for everything that is not a document, sheet, presentation or PDF –
 * Google shortcuts, e-mail (eml/msg), contacts (vcf), calendars (ics), e-books (epub), Visio, drawings and other
 * formats through the Office engine, Apple iWork previews, images, JSON/XML/text, ZIP, and honest info cards.
 * Every view ends with next steps (download, save as PDF / editable copy, open fully).
 */
import { Component, useEffect, useMemo, useState } from 'react';
import { Mail, Contact, CalendarDays, BookOpen, Shapes, Image as ImageIcon, Braces, Archive, File as FileIcon, Link as LinkIcon } from 'lucide-react';
import { Loading } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { FORMATS } from '../../lib/formats.js';
import MailView, { mailText, mailPreview } from './mail.jsx';
import ContactsView, { parseContacts, Avatar } from './contacts.jsx';
import CalendarView, { parseCalendar } from './calendar.jsx';
import EpubView, { epubText, epubCover } from './epub.jsx';
import VisioView, { visioText } from './visio.jsx';
import ImageView, { imageBytes } from './image.jsx';
import DataView from './data.jsx';
import ZipView, { listZip } from './zip.jsx';
import GoogleView, { googleText } from './google.jsx';
import { ConvertView, IworkView, InfoView, infoText } from './info.jsx';
import { ErrorCard, decodeText, looksBinary, useBlobUrl, fmtSize } from './common.jsx';
import { cardText } from './vcard.js';
import { htmlToText } from './mime.js';
import { parseGoogleFile, GOOGLE } from './google.js';
import { binaryStrings } from './extract.js';
import './style.css';
import { Keep, Raw } from '../_shared/keep.jsx';
import { FileTypeIcon } from '../../lib/fileIcons.jsx';

const IMAGES = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'ico', 'avif', 'emf', 'wmf'];
const VISIO = ['vsdx', 'vsdm', 'vstx', 'vssx', 'vstm'];
const CONVERT = ['vsd', 'vss', 'vst', 'pub', 'odg', 'otg', 'fodg', 'odf', 'wpd', 'wps', 'wri', 'abw', 'lwp', 'hwp', 'sxw', 'sxi', 'sxc', 'et', 'dps', 'xps', 'oxps'];
const IWORK = ['pages', 'key', 'numbers'];
const INFO = ['one', 'mpp', 'accdb', 'mdb', 'odb'];

/** which part shows this extension */
export function viewFor(ext, bytes) {
  if (GOOGLE[ext]) return 'google';
  if (ext === 'eml' || ext === 'msg' || ext === 'mht' || ext === 'mhtml') return 'mail';
  if (ext === 'vcf' || ext === 'vcard') return 'contacts';
  if (ext === 'ics' || ext === 'ical' || ext === 'ifb') return 'calendar';
  if (ext === 'epub') return 'epub';
  if (VISIO.includes(ext)) return 'visio';
  if (IWORK.includes(ext)) return 'iwork';
  if (INFO.includes(ext)) return 'info';
  if (CONVERT.includes(ext)) return 'convert';
  if (IMAGES.includes(ext)) return 'image';
  if (ext === 'zip' || (bytes && bytes[0] === 0x50 && bytes[1] === 0x4b && !FORMATS[ext])) return 'zip';
  if (bytes && looksBinary(bytes)) return 'convert';
  return 'data';
}

class Boundary extends Component {
  constructor(p) { super(p); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  render() { return this.state.error ? <ErrorCard error={this.state.error} file={this.props.file} bytes={this.props.bytes} host={this.props.host} /> : this.props.children; }
}

function Editor({ file, bytes, host }) {
  const view = viewFor(file.ext, bytes);
  useEffect(() => {
    const items = [];
    if (view === 'convert' || view === 'iwork') items.push({ level: 'info', text: t('This format is shown in a simple view. “Open fully” renders it exactly with the Office engine.') });
    if (host.fidelity) host.fidelity({ items, kept: true });
    host.setSaver && host.setSaver(async () => bytes); // read-only: the file stays exactly as it is
  }, [file.ext]); // eslint-disable-line react-hooks/exhaustive-deps
  const props = { file, bytes, host };
  const V = { google: GoogleView, mail: MailView, contacts: ContactsView, calendar: CalendarView, epub: EpubView, visio: VisioView, iwork: IworkView, info: InfoView, convert: ConvertView, image: ImageView, zip: ZipView, data: DataView }[view];
  return <div className="of-view-root"><Boundary {...props}><V {...props} /></Boundary></div>;
}

/* ---------------- small preview for Hostwerk's file browser ---------------- */
const KIND_ICON = { mail: Mail, contacts: Contact, calendar: CalendarDays, epub: BookOpen, visio: Shapes, image: ImageIcon, data: Braces, zip: Archive, google: LinkIcon };
function Preview({ bytes, ext, name }) {
  const view = viewFor(ext, bytes);
  const [info, setInfo] = useState(null);
  const img = useMemo(() => (view === 'image' ? imageBytes(bytes, ext) : null), [view, bytes, ext]);
  const imgUrl = useBlobUrl(img && img.bytes, img && img.mime);
  const [cover, setCover] = useState(null);
  const coverUrl = useBlobUrl(cover && cover.cover, cover && cover.mime);
  useEffect(() => {
    let gone = false;
    (async () => {
      try {
        if (view === 'mail') setInfo(mailPreview(bytes, ext));
        else if (view === 'contacts') { const cs = parseContacts(bytes); setInfo({ title: cs[0] ? cs[0].fn : name, sub: cs.length > 1 ? t('{n} contacts', { n: cs.length }) : cs[0] && (cs[0].org || (cs[0].email[0] || {}).value), card: cs[0] }); } else if (view === 'calendar') { const c = parseCalendar(bytes); const next = c.events.find((e) => e.start && e.start > new Date()) || c.events[0]; setInfo({ title: c.name || name, sub: t('{n} appointments', { n: c.events.length }), line: next ? `${next.start ? next.start.toLocaleDateString() : ''} ${next.summary}` : '' }); } else if (view === 'epub') { const c = await epubCover(bytes); if (!gone) { setCover(c); setInfo({ title: c.title || name, sub: c.author }); } } else if (view === 'zip') { const z = await listZip(bytes); if (!gone) setInfo({ title: name, sub: t('{n} files', { n: z.entries.filter((e) => !e.dir).length }), line: z.entries.slice(0, 4).map((e) => e.path).join(' · ') }); } else if (view === 'google') { const j = parseGoogleFile(bytes); setInfo({ title: name, sub: `${t((GOOGLE[ext] || GOOGLE.glink).label)}${j.email ? ` · ${j.email}` : ''}` }); } else if (view === 'data') setInfo({ title: name, line: decodeText(bytes.subarray(0, 600)).text.slice(0, 300), mono: true });
        else setInfo({ title: name, sub: `${t((FORMATS[ext] || {}).label || 'File')} · ${fmtSize(bytes.length)}` });
      } catch { if (!gone) setInfo({ title: name, sub: fmtSize(bytes.length) }); }
    })();
    return () => { gone = true; };
  }, [bytes, ext, name, view]);
  if (view === 'image') return <div className="of-view-preview img">{imgUrl && <Raw as="img" src={imgUrl} alt={name} />}</div>;
  if (!info) return <div className="of-view-preview"><Loading inline /></div>;
  const I = KIND_ICON[view] || FileIcon;
  return (
    <div className="of-view-preview">
      {coverUrl ? <img className="of-view-cover" src={coverUrl} alt="" /> : info.card ? <Avatar c={info.card} size={40} /> : <span className="of-view-picon"><FileTypeIcon name={name} size={40} /></span>}
      <div className="of-view-ptext">
        <b className="of-view-ellipsis"><Keep>{info.title}</Keep></b>
        {info.sub && <span className="small muted of-view-ellipsis"><Keep>{info.sub}</Keep></span>}
        {info.line && <span className={`small of-view-pline${info.mono ? ' mono' : ''}`}><Keep>{info.line}</Keep></span>}
      </div>
    </div>
  );
}

async function toText(bytes, ext) {
  const view = viewFor(ext, bytes);
  switch (view) {
    case 'mail': return mailText(bytes, ext);
    case 'contacts': return parseContacts(bytes).map(cardText).join('\n\n');
    case 'calendar': { const c = parseCalendar(bytes); return [c.name, ...c.events.map((e) => [e.summary, e.start ? e.start.toISOString() : '', e.location, e.description].filter(Boolean).join(' · ')), ...c.todos.map((x) => `☐ ${x.summary}`)].filter(Boolean).join('\n'); }
    case 'epub': return epubText(bytes);
    case 'visio': return visioText(bytes);
    case 'google': return googleText(bytes, ext, '');
    case 'zip': return (await listZip(bytes)).entries.map((e) => e.path).join('\n');
    case 'data': { const { text } = decodeText(bytes); return ext === 'html' || ext === 'htm' ? htmlToText(text) : text; }
    case 'iwork': case 'info': case 'convert': return CONVERT.includes(ext) || INFO.includes(ext) || IWORK.includes(ext) ? infoText(bytes, ext) : binaryStrings(bytes).strings.join('\n');
    default: return '';
  }
}

export default {
  kind: 'viewer',
  Editor,
  Preview,
  toText,
  create: async () => { throw new Error(t('This kind of file cannot be created here.')); },
};
