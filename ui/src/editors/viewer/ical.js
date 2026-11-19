/*
 * iCalendar (RFC 5545) reader – pure. parseICS(text) → { name, tz, events: [Item], todos: [Item] }
 * Item = { uid, summary, description, location, start: Date, end: Date|null, allDay, rrule: {…}|null, rruleText,
 *          tzid, status, organizer, attendees: [], categories: [], due, completed, priority, url, exdates: [] }
 * Time zones: IANA TZIDs via Intl, common Windows names mapped, otherwise the time is taken as local ("floating").
 */
import { unfold, parseLine } from './vcard.js';

const WIN_TZ = {
  'W. Europe Standard Time': 'Europe/Berlin', 'Central Europe Standard Time': 'Europe/Budapest', 'Romance Standard Time': 'Europe/Paris',
  'Central European Standard Time': 'Europe/Warsaw', 'GMT Standard Time': 'Europe/London', 'Greenwich Standard Time': 'Atlantic/Reykjavik',
  'E. Europe Standard Time': 'Europe/Chisinau', 'FLE Standard Time': 'Europe/Kiev', 'GTB Standard Time': 'Europe/Bucharest', 'Russian Standard Time': 'Europe/Moscow',
  'Eastern Standard Time': 'America/New_York', 'Central Standard Time': 'America/Chicago', 'Mountain Standard Time': 'America/Denver', 'Pacific Standard Time': 'America/Los_Angeles',
  'UTC': 'UTC', 'Coordinated Universal Time': 'UTC', 'Tokyo Standard Time': 'Asia/Tokyo', 'China Standard Time': 'Asia/Shanghai', 'India Standard Time': 'Asia/Kolkata',
  'AUS Eastern Standard Time': 'Australia/Sydney', 'Turkey Standard Time': 'Europe/Istanbul', 'Israel Standard Time': 'Asia/Jerusalem',
};

export function ianaZone(tzid) {
  if (!tzid) return null;
  const id = String(tzid).replace(/^"|"$/g, '').replace(/^\/[^/]+\/[^/]+\//, ''); // "/mozilla.org/20050126_1/Europe/Berlin"
  const cand = WIN_TZ[id] || id;
  try { new Intl.DateTimeFormat('en-US', { timeZone: cand }); return cand; } catch { return null; }
}

/** offset in minutes of `zone` at the UTC instant `ms` */
function zoneOffset(zone, ms) {
  const f = new Intl.DateTimeFormat('en-US', { timeZone: zone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const p = Object.fromEntries(f.formatToParts(new Date(ms)).map((x) => [x.type, x.value]));
  const asUtc = Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour % 24, +p.minute, +p.second);
  return (asUtc - ms) / 60000;
}
/** wall time in a zone → Date */
export function zonedTime(y, mo, d, h, mi, s, zone) {
  const guess = Date.UTC(y, mo, d, h, mi, s);
  let off = zoneOffset(zone, guess);
  let t = guess - off * 60000;
  const off2 = zoneOffset(zone, t);
  if (off2 !== off) { off = off2; t = guess - off * 60000; }
  return new Date(t);
}

/** "20240131T090000Z" / "20240131" (+ params) → { date, allDay } */
export function parseDate(value, params = {}, calTz = null) {
  const v = String(value || '').trim();
  const m = /^(\d{4})-?(\d{2})-?(\d{2})(?:T(\d{2}):?(\d{2}):?(\d{2})?(Z)?)?$/.exec(v);
  if (!m) { const d = new Date(v); return { date: Number.isNaN(d.getTime()) ? null : d, allDay: false }; }
  const [, y, mo, d, h, mi, s, z] = m;
  if (h === undefined || params.value === 'DATE') return { date: new Date(+y, +mo - 1, +d), allDay: true };
  if (z) return { date: new Date(Date.UTC(+y, +mo - 1, +d, +h, +mi, +(s || 0))), allDay: false };
  const zone = ianaZone(params.tzid) || (params.tzid ? null : ianaZone(calTz));
  if (zone) return { date: zonedTime(+y, +mo - 1, +d, +h, +mi, +(s || 0), zone), allDay: false, tz: zone };
  return { date: new Date(+y, +mo - 1, +d, +h, +mi, +(s || 0)), allDay: false };
}

/** ISO-8601 duration "P1DT2H" → ms */
export function parseDuration(v) {
  const m = /^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(String(v || '').trim());
  if (!m) return 0;
  const [, sign, w, d, h, mi, s] = m;
  const ms = ((+w || 0) * 7 * 86400 + (+d || 0) * 86400 + (+h || 0) * 3600 + (+mi || 0) * 60 + (+s || 0)) * 1000;
  return sign === '-' ? -ms : ms;
}

export function parseRRule(v) {
  const r = {};
  for (const part of String(v || '').split(';')) { const [k, val] = part.split('='); if (k && val) r[k.toUpperCase()] = val; }
  return r;
}

const DAY_EN = { MO: 'Monday', TU: 'Tuesday', WE: 'Wednesday', TH: 'Thursday', FR: 'Friday', SA: 'Saturday', SU: 'Sunday' };
const ORD = { 1: 'first', 2: 'second', 3: 'third', 4: 'fourth', 5: 'fifth', '-1': 'last' };
/**
 * Recurrence as an English sentence (the UI translates the pieces): "Every 2 weeks on Monday, Wednesday until 31.12.2024".
 * tr(text, vars) lets the caller translate; default is identity with {vars}.
 */
export function describeRRule(r, tr = (s, v) => s.replace(/\{(\w+)\}/g, (_, k) => (v && v[k] != null ? v[k] : ''))) {
  if (!r || !r.FREQ) return '';
  const n = Number(r.INTERVAL || 1);
  const unit = { DAILY: ['day', 'days'], WEEKLY: ['week', 'weeks'], MONTHLY: ['month', 'months'], YEARLY: ['year', 'years'], HOURLY: ['hour', 'hours'], MINUTELY: ['minute', 'minutes'] }[r.FREQ] || ['time', 'times'];
  let s = n === 1 ? tr(`Every ${unit[0]}`) : tr(`Every {n} ${unit[1]}`, { n });
  if (r.BYDAY) {
    const days = r.BYDAY.split(',').map((d) => {
      const m = /^([+-]?\d)?(MO|TU|WE|TH|FR|SA|SU)$/.exec(d);
      if (!m) return d;
      const name = tr(DAY_EN[m[2]]);
      return m[1] ? `${tr(`the ${ORD[String(Number(m[1]))] || m[1]}`)} ${name}` : name;
    });
    s += ` ${tr('on')} ${days.join(', ')}`;
  }
  if (r.BYMONTHDAY) s += ` ${tr('on day {d}', { d: r.BYMONTHDAY })}`;
  if (r.COUNT) s += `, ${tr('{n} times', { n: r.COUNT })}`;
  if (r.UNTIL) { const u = parseDate(r.UNTIL).date; if (u) s += ` ${tr('until {d}', { d: u.toLocaleDateString(tr.locale || undefined) })}`; }
  return s;
}

const unesc = (s) => String(s || '').replace(/\\n/gi, '\n').replace(/\\([,;:\\])/g, '$1');
const person = (p) => ({ name: (p.params.cn || '').replace(/^"|"$/g, ''), email: String(p.value || '').replace(/^mailto:/i, ''), role: p.params.role || '', status: p.params.partstat || '' });

export function parseICS(text) {
  const cal = { name: '', tz: '', events: [], todos: [], journals: 0 };
  const stack = [];
  let cur = null;
  for (const line of unfold(text)) {
    const p = parseLine(line);
    if (p.name === 'BEGIN') {
      const what = p.value.trim().toUpperCase();
      stack.push(what);
      if (what === 'VEVENT' || what === 'VTODO') cur = { kind: what === 'VEVENT' ? 'event' : 'todo', uid: '', summary: '', description: '', location: '', start: null, end: null, allDay: false, rrule: null, rruleText: '', tzid: '', status: '', organizer: null, attendees: [], categories: [], due: null, completed: null, priority: 0, percent: 0, url: '', exdates: [], alarms: 0, _dur: 0 };
      else if (what === 'VALARM' && cur) cur.alarms++;
      else if (what === 'VJOURNAL') cal.journals++;
      continue;
    }
    if (p.name === 'END') {
      const what = stack.pop();
      if ((what === 'VEVENT' || what === 'VTODO') && cur) {
        if (!cur.end && cur.start) cur.end = cur._dur ? new Date(cur.start.getTime() + cur._dur) : cur.allDay ? new Date(cur.start.getTime() + 86400000) : null;
        delete cur._dur;
        (cur.kind === 'event' ? cal.events : cal.todos).push(cur);
        cur = null;
      }
      continue;
    }
    const inside = stack[stack.length - 1];
    if (!cur) {
      if (p.name === 'X-WR-CALNAME' || (p.name === 'NAME' && inside === 'VCALENDAR')) cal.name = unesc(p.value);
      if (p.name === 'X-WR-TIMEZONE') cal.tz = p.value.trim();
      continue;
    }
    if (inside !== 'VEVENT' && inside !== 'VTODO') continue; // inside VALARM
    switch (p.name) {
      case 'UID': cur.uid = p.value; break;
      case 'SUMMARY': cur.summary = unesc(p.value); break;
      case 'DESCRIPTION': cur.description = unesc(p.value); break;
      case 'LOCATION': cur.location = unesc(p.value); break;
      case 'URL': cur.url = p.value.trim(); break;
      case 'STATUS': cur.status = p.value.trim().toUpperCase(); break;
      case 'PRIORITY': cur.priority = Number(p.value) || 0; break;
      case 'PERCENT-COMPLETE': cur.percent = Number(p.value) || 0; break;
      case 'CATEGORIES': cur.categories.push(...p.value.split(',').map((x) => unesc(x).trim()).filter(Boolean)); break;
      case 'DTSTART': { const d = parseDate(p.value, p.params, cal.tz); cur.start = d.date; cur.allDay = d.allDay; cur.tzid = p.params.tzid || ''; break; }
      case 'DTEND': cur.end = parseDate(p.value, p.params, cal.tz).date; break;
      case 'DUE': cur.due = parseDate(p.value, p.params, cal.tz).date; break;
      case 'COMPLETED': cur.completed = parseDate(p.value, p.params, cal.tz).date; break;
      case 'DURATION': cur._dur = parseDuration(p.value); break;
      case 'RRULE': cur.rrule = parseRRule(p.value); cur.rruleText = describeRRule(cur.rrule); break;
      case 'EXDATE': for (const v of p.value.split(',')) { const d = parseDate(v, p.params, cal.tz).date; if (d) cur.exdates.push(d); } break;
      case 'ORGANIZER': cur.organizer = person(p); break;
      case 'ATTENDEE': cur.attendees.push(person(p)); break;
      default: break;
    }
  }
  const t = (x) => (x.start || x.due || new Date(8.64e15)).getTime();
  cal.events.sort((a, b) => t(a) - t(b));
  cal.todos.sort((a, b) => t(a) - t(b));
  return cal;
}

/**
 * Occurrences of an item inside [from, to) – expands simple RRULEs (DAILY/WEEKLY/MONTHLY/YEARLY with INTERVAL, COUNT,
 * UNTIL, BYDAY for weekly) for the month view. Everything else shows only the first occurrence.
 */
export function occurrences(ev, from, to, max = 400) {
  if (!ev.start) return [];
  const dur = ev.end ? ev.end - ev.start : 0;
  const out = [];
  const push = (d) => { if (d < to && d.getTime() + Math.max(dur, 1) > from.getTime() && !ev.exdates.some((x) => Math.abs(x - d) < 60000)) out.push({ start: d, end: new Date(d.getTime() + dur) }); };
  const r = ev.rrule;
  if (!r || !r.FREQ) { push(ev.start); return out; }
  const interval = Number(r.INTERVAL || 1);
  const until = r.UNTIL ? parseDate(r.UNTIL).date : null;
  const count = r.COUNT ? Number(r.COUNT) : Infinity;
  const DAYS = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
  const byday = r.FREQ === 'WEEKLY' && r.BYDAY ? r.BYDAY.split(',').map((d) => DAYS.indexOf(d.slice(-2))).filter((x) => x >= 0) : null;
  let n = 0; let i = 0;
  const s = ev.start;
  while (n < count && i < 5000 && out.length < max) {
    let d;
    if (r.FREQ === 'DAILY') d = new Date(s.getFullYear(), s.getMonth(), s.getDate() + i * interval, s.getHours(), s.getMinutes());
    else if (r.FREQ === 'WEEKLY') {
      const weekStart = new Date(s.getFullYear(), s.getMonth(), s.getDate() - s.getDay() + i * 7 * interval, s.getHours(), s.getMinutes());
      const days = byday || [s.getDay()];
      for (const wd of days.sort()) {
        const dd = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + wd, s.getHours(), s.getMinutes());
        if (dd < s) continue;
        if ((until && dd > until) || n >= count) break;
        n++; push(dd);
      }
      if ((until && weekStart > until) || weekStart > to) break;
      i++; continue;
    } else if (r.FREQ === 'MONTHLY') d = new Date(s.getFullYear(), s.getMonth() + i * interval, s.getDate(), s.getHours(), s.getMinutes());
    else if (r.FREQ === 'YEARLY') d = new Date(s.getFullYear() + i * interval, s.getMonth(), s.getDate(), s.getHours(), s.getMinutes());
    else { push(s); break; }
    if ((until && d > until) || d > to) break;
    n++; push(d); i++;
  }
  return out;
}
