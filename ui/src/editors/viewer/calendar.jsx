/* Calendar (.ics): agenda list and month view, to-dos, details with recurrence as text. */
import { useMemo, useState } from 'react';
import { CalendarDays, ChevronLeft, ChevronRight, MapPin, Repeat, Clock, Users, CheckSquare, Square, Bell, List } from 'lucide-react';
import { Seg, Empty, Modal, Btn } from '@hostwerk/ui';
import { t } from '@hostwerk/i18n';
import { parseICS, occurrences, describeRRule } from './ical.js';
import { ActionRow, decodeText, locale } from './common.jsx';
import { Keep, Raw } from '../_shared/keep.jsx';
import { FileTypeIcon } from '../../lib/fileIcons.jsx';

const tr = (s, v) => t(s, v);
tr.locale = undefined;
const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
const fmtRange = (e) => {
  if (!e.start) return '';
  const L = locale();
  if (e.allDay) {
    const last = e.end ? new Date(e.end.getTime() - 1) : e.start;
    return sameDay(e.start, last) ? e.start.toLocaleDateString(L, { weekday: 'short', dateStyle: undefined, day: 'numeric', month: 'short', year: 'numeric' }) : `${e.start.toLocaleDateString(L, { day: 'numeric', month: 'short' })} – ${last.toLocaleDateString(L, { day: 'numeric', month: 'short', year: 'numeric' })}`;
  }
  const d = e.start.toLocaleDateString(L, { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
  const tm = (x) => x.toLocaleTimeString(L, { hour: '2-digit', minute: '2-digit' });
  if (!e.end) return `${d}, ${tm(e.start)}`;
  return sameDay(e.start, e.end) ? `${d}, ${tm(e.start)}–${tm(e.end)}` : `${d}, ${tm(e.start)} – ${e.end.toLocaleDateString(L, { day: 'numeric', month: 'short' })} ${tm(e.end)}`;
};

export const parseCalendar = (bytes) => parseICS(decodeText(bytes).text);

export default function CalendarView({ file, bytes, host }) {
  const cal = useMemo(() => parseCalendar(bytes), [bytes]);
  const [view, setView] = useState('agenda');
  const first = cal.events.find((e) => e.start && e.start >= new Date(Date.now() - 86400000)) || cal.events[0];
  const [month, setMonth] = useState(() => { const d = (first && first.start) || new Date(); return new Date(d.getFullYear(), d.getMonth(), 1); });
  const [open, setOpen] = useState(null);
  const L = locale();

  if (!cal.events.length && !cal.todos.length) return <div className="of-view-pad"><Empty icon={CalendarDays} title={t('No appointments in this file')} text={t('The calendar is empty or cannot be read.')} /><ActionRow file={file} bytes={bytes} host={host} /></div>;

  const monthEnd = new Date(month.getFullYear(), month.getMonth() + 1, 1);
  const gridStart = new Date(month); gridStart.setDate(1 - ((month.getDay() + 6) % 7));
  const days = [...Array(42).keys()].map((i) => new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + i));
  const occ = view === 'month' ? cal.events.flatMap((e) => occurrences(e, days[0], new Date(days[41].getTime() + 86400000)).map((o) => ({ ...o, e }))) : [];

  return (
    <div className="of-view-cal">
      <div className="of-view-toolbar">
        <h2 className="of-view-title"><FileTypeIcon name={file.name} size={20} /> <Keep>{cal.name || file.name}</Keep></h2>
        <span className="small muted">{t('{n} appointments', { n: cal.events.length })}{cal.todos.length ? ` · ${t('{n} to-dos', { n: cal.todos.length })}` : ''}</span>
        <span className="grow" />
        {view === 'month' && (
          <span className="row">
            <button type="button" className="btn sm ghost icon" title="Previous month" aria-label="Previous month" onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))}><ChevronLeft size={16} /></button>
            <b className="of-view-month">{month.toLocaleDateString(L, { month: 'long', year: 'numeric' })}</b>
            <button type="button" className="btn sm ghost icon" title="Next month" aria-label="Next month" onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}><ChevronRight size={16} /></button>
            <Btn className="sm ghost" onClick={() => { const d = new Date(); setMonth(new Date(d.getFullYear(), d.getMonth(), 1)); }}>Today</Btn>
          </span>
        )}
        <Seg value={view} onChange={setView} options={[{ id: 'agenda', label: t('List'), icon: List }, { id: 'month', label: t('Month'), icon: CalendarDays }]} />
      </div>
      <div className="of-view-calbody">
        {view === 'agenda' && (
          <div className="of-view-agenda">
            {cal.events.map((e, i) => (
              <button key={i} type="button" className="of-view-event" onClick={() => setOpen(e)}>
                <span className="of-view-evdate small">{fmtRange(e)}</span>
                <span className="of-view-evmain"><b>{e.summary ? <Keep>{e.summary}</Keep> : t('(no title)')}</b>{e.location && <span className="small muted"> · <Keep>{e.location}</Keep></span>}</span>
                {e.rrule && <Repeat size={14} className="muted" aria-label={t('Repeats')} />}
              </button>
            ))}
            {cal.todos.length > 0 && <h3 className="of-view-h3">{t('To-dos')}</h3>}
            {cal.todos.map((x, i) => (
              <button key={`t${i}`} type="button" className="of-view-event" onClick={() => setOpen(x)}>
                {x.status === 'COMPLETED' || x.completed ? <CheckSquare size={16} /> : <Square size={16} />}
                <span className="of-view-evmain"><b>{x.summary ? <Keep>{x.summary}</Keep> : t('(no title)')}</b>{x.due && <span className="small muted"> · {t('due {d}', { d: x.due.toLocaleString(L, { dateStyle: 'medium', timeStyle: 'short' }) })}</span>}</span>
              </button>
            ))}
          </div>
        )}
        {view === 'month' && (
          <div className="of-view-month-grid" role="grid">
            {days.slice(0, 7).map((d) => <div key={`h${d}`} className="of-view-dow small muted">{d.toLocaleDateString(L, { weekday: 'short' })}</div>)}
            {days.map((d) => {
              const items = occ.filter((o) => (o.e.allDay ? d >= new Date(o.start.getFullYear(), o.start.getMonth(), o.start.getDate()) && d < o.end : sameDay(o.start, d)));
              return (
                <div key={d.toISOString()} className={`of-view-day${d.getMonth() !== month.getMonth() ? ' out' : ''}${sameDay(d, new Date()) ? ' today' : ''}`} role="gridcell">
                  <span className="small">{d.getDate()}</span>
                  {items.slice(0, 4).map((o, i) => <Raw as="button" key={i} type="button" className={`of-view-chip${o.e.allDay ? ' allday' : ''}`} onClick={() => setOpen({ ...o.e, start: o.start, end: o.end })} title={o.e.summary}>{!o.e.allDay && <span className="mono">{o.start.toLocaleTimeString(L, { hour: '2-digit', minute: '2-digit' })} </span>}{o.e.summary ? <Keep>{o.e.summary}</Keep> : t('(no title)')}</Raw>)}
                  {items.length > 4 && <span className="small muted">{t('+{n} more', { n: items.length - 4 })}</span>}
                </div>
              );
            })}
          </div>
        )}
        {view === 'month' && monthEnd && null}
      </div>
      <ActionRow file={file} bytes={bytes} host={host} />
      {open && (
        <Modal title={open.summary ? <Keep>{open.summary}</Keep> : t('(no title)')} subtitle={fmtRange(open)} onClose={() => setOpen(null)} footer={<Btn onClick={() => setOpen(null)}>Close</Btn>}>
          <div className="stack">
            {open.rrule && <div className="row"><Repeat size={16} /> {describeRRule(open.rrule, tr)}</div>}
            {open.location && <div className="row"><MapPin size={16} /> <Keep>{open.location}</Keep></div>}
            {open.tzid && <div className="row small muted"><Clock size={14} /> {t('Time zone: {tz}', { tz: open.tzid })}</div>}
            {open.alarms > 0 && <div className="row small muted"><Bell size={14} /> {t('{n} reminders', { n: open.alarms })}</div>}
            {open.organizer && <div className="row"><Users size={16} /> {t('Organiser')}: <Keep>{open.organizer.name || open.organizer.email}</Keep></div>}
            {open.attendees.length > 0 && <div className="small">{t('Attendees')}: <Keep>{open.attendees.map((a) => a.name || a.email).join(', ')}</Keep></div>}
            {open.description && <div className="of-view-pre of-view-desc"><Keep>{open.description}</Keep></div>}
            {open.status && <span className="pill">{open.status}</span>}
          </div>
        </Modal>
      )}
    </div>
  );
}
