/* Contacts (.vcf): list of cards on the left, details on the right, export of a single card. */
import { useMemo, useState } from 'react';
import { User, Phone, Mail, MapPin, Building2, Globe, Cake, StickyNote, Download, Search, Users } from 'lucide-react';
import { Btn, TileEmpty, Empty } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { parseVCards, formatAdr } from './vcard.js';
import { ActionRow } from './common.jsx';
import { decodeText } from './common.jsx';
import { Keep, Raw } from '../_shared/keep.jsx';

const TYPE = { work: 'Work', home: 'Private', cell: 'Mobile', mobile: 'Mobile', fax: 'Fax', pager: 'Pager', main: 'Main', other: 'Other', voice: '', pref: '', internet: '', text: 'Text', iphone: 'iPhone' };
const typeLabel = (types) => (types || []).map((x) => (TYPE[x] === undefined ? x : TYPE[x])).filter(Boolean).map((x) => t(x)).join(', ');
const initials = (s) => String(s || '?').split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0]).join('').toUpperCase();
function bday(v) {
  const m = /^(\d{4}|--)-?(\d{2})-?(\d{2})/.exec(v || '');
  if (!m) return v;
  const d = new Date(m[1] === '--' ? 2000 : +m[1], +m[2] - 1, +m[3]);
  return d.toLocaleDateString(document.documentElement.lang === 'en' ? 'en-US' : 'de-DE', m[1] === '--' ? { day: 'numeric', month: 'long' } : { dateStyle: 'long' });
}

export function Avatar({ c, size = 40 }) {
  const src = c.photo && (c.photo.data ? `data:${c.photo.mime || 'image/jpeg'};base64,${c.photo.data}` : '');
  return src ? <img className="of-view-avatar" src={src} alt="" style={{ width: size, height: size }} /> : <span className="of-view-avatar" style={{ width: size, height: size }} aria-hidden="true"><Keep>{initials(c.fn)}</Keep></span>;
}

export const parseContacts = (bytes) => parseVCards(decodeText(bytes).text);

export default function ContactsView({ file, bytes, host }) {
  const cards = useMemo(() => parseContacts(bytes), [bytes]);
  const [sel, setSel] = useState(0);
  const [q, setQ] = useState('');
  const list = cards.map((c, i) => ({ c, i })).filter(({ c }) => !q || `${c.fn} ${c.org} ${c.email.map((e) => e.value).join(' ')} ${c.tel.map((x) => x.value).join(' ')}`.toLowerCase().includes(q.toLowerCase()));
  const c = cards[sel];
  if (!cards.length) return <div className="of-view-pad"><Empty icon={Users} title={t('No contacts in this file')} text={t('The file does not contain a readable vCard.')} /><ActionRow file={file} bytes={bytes} host={host} /></div>;
  const exportOne = () => host.download(new TextEncoder().encode(c.raw), `${(c.fn || 'contact').replace(/[\\/:*?"<>|]+/g, '_')}.vcf`, 'text/vcard');
  return (
    <div className="of-view-split">
      {cards.length > 1 && (
        <aside className="of-view-list">
          <div className="of-view-listhead"><Search size={14} /><input className="input" value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search contacts" aria-label="Search contacts" /></div>
          <div className="of-view-listbody" role="listbox" aria-label="Contacts">
            {list.map(({ c: x, i }) => (
              <Raw as="button" key={i} type="button" role="option" aria-selected={i === sel} className={`of-view-item${i === sel ? ' active' : ''}`} onClick={() => setSel(i)} title={[x.fn, x.org].filter(Boolean).join(' · ') || undefined}>
                <Avatar c={x} size={28} /><span className="of-view-ellipsis"><b>{x.fn ? <Keep>{x.fn}</Keep> : t('(no name)')}</b>{x.org && <span className="small muted"> · <Keep>{x.org}</Keep></span>}</span>
              </Raw>
            ))}
            {!list.length && <TileEmpty icon={Search} text={t('Nothing found.')} />}
          </div>
          <div className="small muted of-view-listfoot">{t('{n} contacts', { n: cards.length })}</div>
        </aside>
      )}
      <div className="of-view-detail">
        {c && (
          <div className="card of-view-contact">
            <div className="of-view-contacthead">
              <Avatar c={c} size={64} />
              <div>
                <h2>{c.fn ? <Keep>{c.fn}</Keep> : t('(no name)')}</h2>
                {(c.title || c.org) && <div className="muted"><Keep>{[c.title, c.org].filter(Boolean).join(' · ')}</Keep></div>}
                {c.nickname && <div className="small muted">„<Keep>{c.nickname}</Keep>“</div>}
              </div>
              <span className="grow" />
              {cards.length > 1 && <Btn className="sm" icon={Download} onClick={exportOne}>Export this contact</Btn>}
            </div>
            <div className="card-body stack">
              {c.tel.map((x, i) => <Row key={`t${i}`} icon={Phone} label={typeLabel(x.types) || t('Phone')}><a href={`tel:${x.value.replace(/\s+/g, '')}`} onClick={(e) => e.preventDefault()}><Keep>{x.value}</Keep></a></Row>)}
              {c.email.map((x, i) => <Row key={`e${i}`} icon={Mail} label={typeLabel(x.types) || t('E-mail')}><span className="of-view-email"><Keep>{x.value}</Keep></span></Row>)}
              {c.adr.map((x, i) => <Row key={`a${i}`} icon={MapPin} label={typeLabel(x.types) || t('Address')}><span className="of-view-pre"><Keep>{formatAdr(x).replace(/, /g, '\n')}</Keep></span></Row>)}
              {c.org && <Row icon={Building2} label={t('Organisation')}><Keep>{c.org}</Keep></Row>}
              {c.url.map((u, i) => <Row key={`u${i}`} icon={Globe} label={t('Website')}><button type="button" className="of-view-link" onClick={() => host.openLink && /^https?:/i.test(u) && host.openLink(u)}><Keep>{u}</Keep></button></Row>)}
              {c.bday && <Row icon={Cake} label={t('Birthday')}>{bday(c.bday)}</Row>}
              {c.impp.map((x, i) => <Row key={`i${i}`} icon={User} label={t('Messenger')}><Keep>{x}</Keep></Row>)}
              {c.categories.length > 0 && <Row icon={User} label={t('Groups')}>{c.categories.map((x) => <span key={x} className="pill"><Keep>{x}</Keep></span>)}</Row>}
              {c.note && <Row icon={StickyNote} label={t('Note')}><span className="of-view-pre"><Keep>{c.note}</Keep></span></Row>}
            </div>
          </div>
        )}
        <ActionRow file={file} bytes={bytes} host={host} />
      </div>
    </div>
  );
}

function Row({ icon: I, label, children }) {
  return <div className="of-view-row"><I size={16} className="muted" /><span className="of-view-rowlabel small muted">{label}</span><span className="of-view-rowval">{children}</span></div>;
}
