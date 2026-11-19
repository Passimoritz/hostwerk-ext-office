import test from 'node:test';
import assert from 'node:assert/strict';
import { parseVCards, formatAdr } from '../ui/src/editors/viewer/vcard.js';
import { parseICS, describeRRule, parseRRule, parseDuration, occurrences, parseDate } from '../ui/src/editors/viewer/ical.js';

test('vCard 3.0 and 4.0, several cards, folding, escapes', () => {
  const v = [
    'BEGIN:VCARD', 'VERSION:3.0', 'N:Muster;Lena;;Dr.;', 'FN:Dr. Lena Muster', 'ORG:Hostwerk GmbH;Entwicklung', 'TITLE:Leiterin',
    'TEL;TYPE=work,voice:+49 30 123', 'TEL;TYPE=cell:+49 170 456', 'EMAIL;TYPE=internet,work:lena@example.com',
    'ADR;TYPE=work:;;Hauptstr. 1;Berlin;;10115;Deutschland', 'NOTE:Zeile 1\\nZeile 2\\, mit Komma', 'PHOTO;ENCODING=b;TYPE=JPEG:/9j/4AAQ', ' SkZJRg==',
    'END:VCARD',
    'BEGIN:VCARD', 'VERSION:4.0', 'FN:Max', 'EMAIL;TYPE=home:max@example.org', 'PHOTO:data:image/png;base64,iVBORw0KGgo=', 'END:VCARD',
  ].join('\r\n');
  const cs = parseVCards(v);
  assert.equal(cs.length, 2);
  const c = cs[0];
  assert.equal(c.fn, 'Dr. Lena Muster');
  assert.equal(c.n.family, 'Muster');
  assert.equal(c.org, 'Hostwerk GmbH · Entwicklung');
  assert.deepEqual(c.tel.map((x) => [x.value, x.types]), [['+49 30 123', ['work', 'voice']], ['+49 170 456', ['cell']]]);
  assert.deepEqual(c.email[0], { value: 'lena@example.com', types: ['work'] });
  assert.equal(formatAdr(c.adr[0]), 'Hauptstr. 1, 10115 Berlin, Deutschland');
  assert.equal(c.note, 'Zeile 1\nZeile 2, mit Komma');
  assert.equal(c.photo.data, '/9j/4AAQSkZJRg==');
  assert.equal(c.photo.mime, 'image/jpeg');
  assert.equal(cs[1].photo.mime, 'image/png');
  assert.match(c.raw, /^BEGIN:VCARD[\s\S]*END:VCARD\r\n$/);
});
test('vCard 2.1 with quoted-printable and bare types', () => {
  const v = 'BEGIN:VCARD\r\nVERSION:2.1\r\nN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:M=C3=BCller;J=C3=B6rg\r\nTEL;WORK;VOICE:0211 1\r\nADR;HOME;ENCODING=QUOTED-PRINTABLE:;;Stra=C3=9Fe 5=\r\na;K=C3=B6ln;;;\r\nEND:VCARD\r\n';
  const [c] = parseVCards(v);
  assert.equal(c.n.family, 'Müller');
  assert.equal(c.fn, 'Jörg Müller');
  assert.deepEqual(c.tel[0].types, ['work', 'voice']);
  assert.equal(c.adr[0].street, 'Straße 5a');
  assert.equal(c.adr[0].city, 'Köln');
});
test('iCalendar: events, todos, all-day, UTC, TZID, duration, rrule', () => {
  const ics = [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'X-WR-CALNAME:Team', 'BEGIN:VTIMEZONE', 'TZID:Europe/Berlin', 'END:VTIMEZONE',
    'BEGIN:VEVENT', 'UID:1', 'SUMMARY:Daily\\, kurz', 'DTSTART;TZID=Europe/Berlin:20240305T090000', 'DURATION:PT15M', 'RRULE:FREQ=WEEKLY;BYDAY=MO,WE;COUNT=4', 'LOCATION:Raum 1',
    'ATTENDEE;CN="Lena";PARTSTAT=ACCEPTED:mailto:lena@x.de', 'BEGIN:VALARM', 'TRIGGER:-PT5M', 'DESCRIPTION:alarm', 'END:VALARM', 'END:VEVENT',
    'BEGIN:VEVENT', 'UID:2', 'SUMMARY:Feiertag', 'DTSTART;VALUE=DATE:20240501', 'END:VEVENT',
    'BEGIN:VEVENT', 'UID:3', 'SUMMARY:Call', 'DTSTART:20240301T120000Z', 'DTEND:20240301T130000Z', 'DESCRIPTION:Zeile\\nzwei', 'END:VEVENT',
    'BEGIN:VTODO', 'UID:4', 'SUMMARY:Bericht', 'DUE:20240310T170000Z', 'STATUS:NEEDS-ACTION', 'END:VTODO', 'END:VCALENDAR',
  ].join('\r\n');
  const cal = parseICS(ics);
  assert.equal(cal.name, 'Team');
  assert.deepEqual(cal.events.map((e) => e.uid), ['3', '1', '2']);
  const daily = cal.events[1];
  assert.equal(daily.summary, 'Daily, kurz');
  assert.equal(daily.start.toISOString(), '2024-03-05T08:00:00.000Z'); // CET = UTC+1
  assert.equal(daily.end - daily.start, 15 * 60000);
  assert.equal(daily.alarms, 1);
  assert.equal(daily.description, '');
  assert.equal(daily.attendees[0].name, 'Lena');
  assert.equal(daily.rruleText, 'Every week on Monday, Wednesday, 4 times');
  assert.ok(cal.events[2].allDay);
  assert.equal(cal.events[0].description, 'Zeile\nzwei');
  assert.equal(cal.todos[0].summary, 'Bericht');
  const occ = occurrences(daily, new Date('2024-03-01'), new Date('2024-04-01'));
  assert.equal(occ.length, 4);
  // summer time: 9:00 Berlin = 07:00 UTC
  assert.equal(parseDate('20240701T090000', { tzid: 'Europe/Berlin' }).date.toISOString(), '2024-07-01T07:00:00.000Z');
  assert.equal(parseDate('20240701T090000', { tzid: 'W. Europe Standard Time' }).date.toISOString(), '2024-07-01T07:00:00.000Z');
  assert.equal(parseDuration('P1DT2H'), 26 * 3600000);
  assert.equal(describeRRule(parseRRule('FREQ=MONTHLY;INTERVAL=2;BYDAY=-1FR')), 'Every 2 months on the last Friday');
});
