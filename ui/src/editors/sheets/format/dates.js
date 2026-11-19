/*
 * Excel serial dates (1900 system with the historic fake 1900-02-29 = serial 60; optional 1904 system).
 * Serial 1 = 1900-01-01, serial 61 = 1900-03-01, the fraction is the time of day.
 */
const DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);
const D1904 = 1462;

function fromParts(y, m, d, h = 0, mi = 0, s = 0, ms = 0, date1904 = false) {
  if (y === 1900 && m === 2 && d === 29) return 60 + (h * 3600000 + mi * 60000 + s * 1000 + ms) / DAY - (date1904 ? D1904 : 0);
  const dt = new Date(0);
  dt.setUTCFullYear(y, m - 1, d);
  dt.setUTCHours(h, mi, s, ms);
  let serial = (dt.getTime() - EPOCH) / DAY;
  if (serial < 61) serial -= 1; // before 1900-03-01 Excel counts the fake leap day
  return date1904 ? serial - D1904 : serial;
}

/** Like Excel DATE(): years 0–1899 get 1900 added, overflowing months/days roll over. m is 1-based. */
export function serialFromParts(y, m, d, h = 0, mi = 0, s = 0, ms = 0, date1904 = false) {
  y = Math.trunc(y);
  if (y >= 0 && y < 1900) y += 1900;
  return fromParts(y, Math.trunc(m), Math.trunc(d), h, mi, s, ms, date1904);
}

/** serial → { y, m, d, h, mi, s, ms, dow } (dow 0 = Sunday, Excel-compatible incl. serial 0 = 1900-01-00) */
export function partsFromSerial(serial, date1904 = false) {
  if (date1904) serial += D1904;
  const total = Math.round(serial * DAY);
  const days = Math.floor(total / DAY);
  const msDay = total - days * DAY;
  const h = Math.floor(msDay / 3600000), mi = Math.floor(msDay / 60000) % 60, s = Math.floor(msDay / 1000) % 60, ms = msDay % 1000;
  const dow = (((days + 6) % 7) + 7) % 7;
  let y, m, d;
  if (days === 60) { y = 1900; m = 2; d = 29; }
  else if (days === 0) { y = 1900; m = 1; d = 0; }
  else {
    const dt = new Date(EPOCH + (days < 60 ? days + 1 : days) * DAY);
    y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate();
  }
  return { y, m, d, h, mi, s, ms, dow };
}

/** JS Date (local components) → serial */
export function dateToSerial(date, date1904 = false) {
  return fromParts(date.getFullYear(), date.getMonth() + 1, date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds(), date1904);
}

/** serial → local JS Date */
export function serialToDate(serial, date1904 = false) {
  const p = partsFromSerial(serial, date1904);
  const dt = new Date(2000, 0, 1, p.h, p.mi, p.s, p.ms);
  dt.setFullYear(p.y, p.m - 1, p.d);
  return dt;
}

export function daysInMonth(y, m) {
  if (y === 1900 && m === 2) return 29; // Excel's fake leap day
  const dt = new Date(0);
  dt.setUTCFullYear(y, m, 0);
  return dt.getUTCDate();
}

export function todaySerial() {
  const n = new Date();
  return fromParts(n.getFullYear(), n.getMonth() + 1, n.getDate());
}

export const nowSerial = () => dateToSerial(new Date());
