/*
 * Number/date conventions for display and input. `lang` picks names and defaults, `decimal` ('.' | ',') overrides
 * the separators, `dateFormat` (a format code like 'yyyy-mm-dd') overrides the short date, `functionNames` ('auto' | 'en' | 'de')
 * the language of function names, WAHR/FALSCH and error codes (formula/i18n.js).
 */
const NAMES = {
  en: {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    monthsShort: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    daysShort: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  },
  de: {
    months: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
    monthsShort: ['Jan', 'Feb', 'Mrz', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'],
    days: ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'],
    daysShort: ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'],
  },
};
export const MONTH_NAMES = NAMES;

function orderOf(code) {
  const s = String(code).toLowerCase().replace(/"[^"]*"|\[[^\]]*\]/g, '');
  const pos = { Y: s.indexOf('y'), M: s.indexOf('m'), D: s.indexOf('d') };
  const keys = Object.keys(pos).filter((k) => pos[k] >= 0).sort((a, b) => pos[a] - pos[b]);
  const o = keys.join('');
  return o === 'DMY' || o === 'MDY' || o === 'YMD' ? o : null;
}

export function makeLocale({ lang = 'en', decimal = 'auto', dateFormat = 'auto', functionNames = 'auto' } = {}) {
  const de = String(lang || 'en').toLowerCase().startsWith('de');
  const L = de ? 'de' : 'en';
  // function names, WAHR/FALSCH and error codes: like the language, or fixed by settings.sheets.functionNames
  const fnLang = functionNames === 'en' || functionNames === 'de' ? functionNames : L;
  const dec = decimal === '.' || decimal === ',' ? decimal : de ? ',' : '.';
  const shortDate = dateFormat && dateFormat !== 'auto' ? dateFormat : de ? 'dd.mm.yyyy' : 'm/d/yyyy';
  const argSep = dec === ',' ? ';' : ',';
  return {
    lang: L,
    decimal: dec,
    group: dec === ',' ? '.' : ',',
    argSep,
    listSep: argSep,
    dateOrder: orderOf(shortDate) || (de ? 'DMY' : 'MDY'),
    shortDate,
    longDate: de ? 'dddd, d. mmmm yyyy' : 'dddd, mmmm d, yyyy',
    time: de ? 'hh:mm' : 'h:mm AM/PM',
    timeSec: de ? 'hh:mm:ss' : 'h:mm:ss AM/PM',
    dateTime: `${shortDate} ${de ? 'hh:mm' : 'h:mm'}`,
    currency: de ? '€' : '$',
    currencyFmt: de ? '#,##0.00 €' : '$#,##0.00',
    ...NAMES[L],
    fnLang,
    trueStr: fnLang === 'de' ? 'WAHR' : 'TRUE',
    falseStr: fnLang === 'de' ? 'FALSCH' : 'FALSE',
  };
}
