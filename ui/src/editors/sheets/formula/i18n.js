/*
 * Localized function names, logical values and error values (German Excel), for input and display only.
 * Files and the model always store the English names (that is the file format: OOXML, ODF and CSV formulas are English).
 *
 *   localFn('VLOOKUP', 'de')      → 'SVERWEIS'         internalFn('SVERWEIS', 'de') → 'VLOOKUP'
 *   localErr('#N/A', 'de')        → '#NV'              internalErr('#NV')           → '#N/A'
 *   localArgs('number1, [number2], …', 'de') → 'Zahl1, [Zahl2], …'
 *
 * Names are Microsoft's official German function names – checked against a German Excel 16 via COM (Range.Formula →
 * Range.FormulaLocal for all 446 names, 2026-11-19: 4 corrected – SCHÄTZER, DM, SPALTENUMBRUCH, ZEILENUMBRUCH). Functions without a German name in
 * Excel (or the engine's own undotted aliases like CEILINGMATH) stay English. English names are always accepted on input.
 */

export const FN_DE = {
  // Math & trig
  ABS: 'ABS', ACOS: 'ARCCOS', ACOSH: 'ARCCOSHYP', ACOT: 'ARCCOT', ACOTH: 'ARCCOTHYP', AGGREGATE: 'AGGREGAT', ARABIC: 'ARABISCH',
  ASIN: 'ARCSIN', ASINH: 'ARCSINHYP', ATAN: 'ARCTAN', ATAN2: 'ARCTAN2', ATANH: 'ARCTANHYP', BASE: 'BASIS', CEILING: 'OBERGRENZE',
  'CEILING.MATH': 'OBERGRENZE.MATHEMATIK', 'CEILING.PRECISE': 'OBERGRENZE.GENAU', COMBIN: 'KOMBINATIONEN', COMBINA: 'KOMBINATIONEN2',
  COS: 'COS', COSH: 'COSHYP', COT: 'COT', COTH: 'COTHYP', CSC: 'COSEC', CSCH: 'COSECHYP', DECIMAL: 'DEZIMAL', DEGREES: 'GRAD',
  EVEN: 'GERADE', EXP: 'EXP', FACT: 'FAKULTÄT', FACTDOUBLE: 'ZWEIFAKULTÄT', FLOOR: 'UNTERGRENZE', 'FLOOR.MATH': 'UNTERGRENZE.MATHEMATIK',
  'FLOOR.PRECISE': 'UNTERGRENZE.GENAU', GCD: 'GGT', INT: 'GANZZAHL', 'ISO.CEILING': 'ISO.OBERGRENZE', LCM: 'KGV', LN: 'LN', LOG: 'LOG',
  LOG10: 'LOG10', MDETERM: 'MDET', MINVERSE: 'MINV', MMULT: 'MMULT', MOD: 'REST', MROUND: 'VRUNDEN', MULTINOMIAL: 'POLYNOMIAL',
  MUNIT: 'MEINHEIT', ODD: 'UNGERADE', PI: 'PI', POWER: 'POTENZ', PRODUCT: 'PRODUKT', QUOTIENT: 'QUOTIENT', RADIANS: 'BOGENMASS',
  RAND: 'ZUFALLSZAHL', RANDARRAY: 'ZUFALLSMATRIX', RANDBETWEEN: 'ZUFALLSBEREICH', ROMAN: 'RÖMISCH', ROUND: 'RUNDEN', ROUNDDOWN: 'ABRUNDEN',
  ROUNDUP: 'AUFRUNDEN', SEC: 'SEC', SECH: 'SECHYP', SEQUENCE: 'SEQUENZ', SERIESSUM: 'POTENZREIHE', SIGN: 'VORZEICHEN', SIN: 'SIN',
  SINH: 'SINHYP', SQRT: 'WURZEL', SQRTPI: 'WURZELPI', SUBTOTAL: 'TEILERGEBNIS', SUM: 'SUMME', SUMIF: 'SUMMEWENN', SUMIFS: 'SUMMEWENNS',
  SUMPRODUCT: 'SUMMENPRODUKT', SUMSQ: 'QUADRATESUMME', SUMX2MY2: 'SUMMEX2MY2', SUMX2PY2: 'SUMMEX2PY2', SUMXMY2: 'SUMMEXMY2', TAN: 'TAN',
  TANH: 'TANHYP', TRUNC: 'KÜRZEN',
  // Statistical
  AVEDEV: 'MITTELABW', AVERAGE: 'MITTELWERT', AVERAGEA: 'MITTELWERTA', AVERAGEIF: 'MITTELWERTWENN', AVERAGEIFS: 'MITTELWERTWENNS',
  'BETA.DIST': 'BETA.VERT', 'BETA.INV': 'BETA.INV', BETADIST: 'BETAVERT', BETAINV: 'BETAINV', 'BINOM.DIST': 'BINOM.VERT',
  'BINOM.DIST.RANGE': 'BINOM.VERT.BEREICH', 'BINOM.INV': 'BINOM.INV', BINOMDIST: 'BINOMVERT', CHIDIST: 'CHIVERT', CHIINV: 'CHIINV',
  'CHISQ.DIST': 'CHIQU.VERT', 'CHISQ.DIST.RT': 'CHIQU.VERT.RE', 'CHISQ.INV': 'CHIQU.INV', 'CHISQ.INV.RT': 'CHIQU.INV.RE',
  'CHISQ.TEST': 'CHIQU.TEST', CHITEST: 'CHITEST', CONFIDENCE: 'KONFIDENZ', 'CONFIDENCE.NORM': 'KONFIDENZ.NORM', 'CONFIDENCE.T': 'KONFIDENZ.T',
  CORREL: 'KORREL', COUNT: 'ANZAHL', COUNTA: 'ANZAHL2', COUNTBLANK: 'ANZAHLLEEREZELLEN', COUNTIF: 'ZÄHLENWENN', COUNTIFS: 'ZÄHLENWENNS',
  COVAR: 'KOVAR', 'COVARIANCE.P': 'KOVARIANZ.P', 'COVARIANCE.S': 'KOVARIANZ.S', CRITBINOM: 'KRITBINOM', DEVSQ: 'SUMQUADABW',
  'EXPON.DIST': 'EXPON.VERT', EXPONDIST: 'EXPONVERT', 'F.DIST': 'F.VERT', 'F.DIST.RT': 'F.VERT.RE', 'F.INV': 'F.INV', 'F.INV.RT': 'F.INV.RE',
  'F.TEST': 'F.TEST', FDIST: 'FVERT', FINV: 'FINV', FISHER: 'FISHER', FISHERINV: 'FISHERINV', FORECAST: 'SCHÄTZER',
  'FORECAST.LINEAR': 'PROGNOSE.LINEAR', FREQUENCY: 'HÄUFIGKEIT', FTEST: 'FTEST', GAMMA: 'GAMMA', 'GAMMA.DIST': 'GAMMA.VERT',
  'GAMMA.INV': 'GAMMA.INV', GAMMADIST: 'GAMMAVERT', GAMMAINV: 'GAMMAINV', GAMMALN: 'GAMMALN', 'GAMMALN.PRECISE': 'GAMMALN.GENAU',
  GAUSS: 'GAUSS', GEOMEAN: 'GEOMITTEL', GROWTH: 'VARIATION', HARMEAN: 'HARMITTEL', 'HYPGEOM.DIST': 'HYPGEOM.VERT', HYPGEOMDIST: 'HYPGEOMVERT',
  INTERCEPT: 'ACHSENABSCHNITT', KURT: 'KURT', LARGE: 'KGRÖSSTE', LINEST: 'RGP', LOGEST: 'RKP', LOGINV: 'LOGINV',
  'LOGNORM.DIST': 'LOGNORM.VERT', 'LOGNORM.INV': 'LOGNORM.INV', LOGNORMDIST: 'LOGNORMVERT', MAX: 'MAX', MAXA: 'MAXA', MAXIFS: 'MAXWENNS',
  MEDIAN: 'MEDIAN', MIN: 'MIN', MINA: 'MINA', MINIFS: 'MINWENNS', MODE: 'MODALWERT', 'MODE.MULT': 'MODUS.VIELF', 'MODE.SNGL': 'MODUS.EINF',
  'NEGBINOM.DIST': 'NEGBINOM.VERT', NEGBINOMDIST: 'NEGBINOMVERT', 'NORM.DIST': 'NORM.VERT', 'NORM.INV': 'NORM.INV', 'NORM.S.DIST': 'NORM.S.VERT',
  'NORM.S.INV': 'NORM.S.INV', NORMDIST: 'NORMVERT', NORMINV: 'NORMINV', NORMSDIST: 'STANDNORMVERT', NORMSINV: 'STANDNORMINV',
  PEARSON: 'PEARSON', PERCENTILE: 'QUANTIL', 'PERCENTILE.EXC': 'QUANTIL.EXKL', 'PERCENTILE.INC': 'QUANTIL.INKL', PERCENTRANK: 'QUANTILSRANG',
  'PERCENTRANK.EXC': 'QUANTILSRANG.EXKL', 'PERCENTRANK.INC': 'QUANTILSRANG.INKL', PERMUT: 'VARIATIONEN', PERMUTATIONA: 'VARIATIONEN2',
  PHI: 'PHI', POISSON: 'POISSON', 'POISSON.DIST': 'POISSON.VERT', PROB: 'WAHRSCHBEREICH', QUARTILE: 'QUARTILE', 'QUARTILE.EXC': 'QUARTILE.EXKL',
  'QUARTILE.INC': 'QUARTILE.INKL', RANK: 'RANG', 'RANK.AVG': 'RANG.MITTELW', 'RANK.EQ': 'RANG.GLEICH', RSQ: 'BESTIMMTHEITSMASS',
  SKEW: 'SCHIEFE', 'SKEW.P': 'SCHIEFE.P', SLOPE: 'STEIGUNG', SMALL: 'KKLEINSTE', STANDARDIZE: 'STANDARDISIERUNG', STDEV: 'STABW',
  'STDEV.P': 'STABW.N', 'STDEV.S': 'STABW.S', STDEVA: 'STABWA', STDEVP: 'STABWN', STDEVPA: 'STABWNA', STEYX: 'STFEHLERYX',
  'T.DIST': 'T.VERT', 'T.DIST.2T': 'T.VERT.2S', 'T.DIST.RT': 'T.VERT.RE', 'T.INV': 'T.INV', 'T.INV.2T': 'T.INV.2S', 'T.TEST': 'T.TEST',
  TDIST: 'TVERT', TINV: 'TINV', TREND: 'TREND', TRIMMEAN: 'GESTUTZTMITTEL', TTEST: 'TTEST', VAR: 'VARIANZ', 'VAR.P': 'VAR.P', 'VAR.S': 'VAR.S',
  VARA: 'VARIANZA', VARP: 'VARIANZEN', VARPA: 'VARIANZENA', WEIBULL: 'WEIBULL', 'WEIBULL.DIST': 'WEIBULL.VERT', 'Z.TEST': 'G.TEST', ZTEST: 'GTEST',
  // Logical
  AND: 'UND', FALSE: 'FALSCH', IF: 'WENN', IFERROR: 'WENNFEHLER', IFNA: 'WENNNV', IFS: 'WENNS', NOT: 'NICHT', OR: 'ODER',
  SWITCH: 'ERSTERWERT', TRUE: 'WAHR', XOR: 'XODER', LET: 'LET', LAMBDA: 'LAMBDA',
  // Text
  CHAR: 'ZEICHEN', CLEAN: 'SÄUBERN', CODE: 'CODE', CONCAT: 'TEXTKETTE', CONCATENATE: 'VERKETTEN', DOLLAR: 'DM', EXACT: 'IDENTISCH',
  FIND: 'FINDEN', FIXED: 'FEST', LEFT: 'LINKS', LEN: 'LÄNGE', LOWER: 'KLEIN', MID: 'TEIL', NUMBERVALUE: 'ZAHLENWERT', PROPER: 'GROSS2',
  REPLACE: 'ERSETZEN', REPT: 'WIEDERHOLEN', RIGHT: 'RECHTS', SEARCH: 'SUCHEN', SUBSTITUTE: 'WECHSELN', T: 'T', TEXT: 'TEXT',
  TEXTAFTER: 'TEXTNACH', TEXTBEFORE: 'TEXTVOR', TEXTJOIN: 'TEXTVERKETTEN', TEXTSPLIT: 'TEXTTEILEN', TRIM: 'GLÄTTEN', UNICHAR: 'UNIZEICHEN',
  UNICODE: 'UNICODE', UPPER: 'GROSS', VALUE: 'WERT',
  // Date & time
  DATE: 'DATUM', DATEDIF: 'DATEDIF', DATEVALUE: 'DATWERT', DAY: 'TAG', DAYS: 'TAGE', DAYS360: 'TAGE360', EDATE: 'EDATUM', EOMONTH: 'MONATSENDE',
  HOUR: 'STUNDE', ISOWEEKNUM: 'ISOKALENDERWOCHE', MINUTE: 'MINUTE', MONTH: 'MONAT', NETWORKDAYS: 'NETTOARBEITSTAGE',
  'NETWORKDAYS.INTL': 'NETTOARBEITSTAGE.INTL', NOW: 'JETZT', SECOND: 'SEKUNDE', TIME: 'ZEIT', TIMEVALUE: 'ZEITWERT', TODAY: 'HEUTE',
  WEEKDAY: 'WOCHENTAG', WEEKNUM: 'KALENDERWOCHE', WORKDAY: 'ARBEITSTAG', 'WORKDAY.INTL': 'ARBEITSTAG.INTL', YEAR: 'JAHR', YEARFRAC: 'BRTEILJAHRE',
  // Lookup & reference
  ADDRESS: 'ADRESSE', AREAS: 'BEREICHE', CHOOSE: 'WAHL', CHOOSECOLS: 'SPALTENWAHL', CHOOSEROWS: 'ZEILENWAHL', COLUMN: 'SPALTE',
  COLUMNS: 'SPALTEN', DROP: 'WEGLASSEN', EXPAND: 'ERWEITERN', FILTER: 'FILTER', FORMULATEXT: 'FORMELTEXT', HLOOKUP: 'WVERWEIS',
  HSTACK: 'HSTAPELN', HYPERLINK: 'HYPERLINK', INDEX: 'INDEX', INDIRECT: 'INDIREKT', LOOKUP: 'VERWEIS', MATCH: 'VERGLEICH',
  OFFSET: 'BEREICH.VERSCHIEBEN', ROW: 'ZEILE', ROWS: 'ZEILEN', SORT: 'SORTIEREN', SORTBY: 'SORTIERENNACH', TAKE: 'ÜBERNEHMEN',
  TOCOL: 'ZUSPALTE', TOROW: 'ZUZEILE', TRANSPOSE: 'MTRANS', UNIQUE: 'EINDEUTIG', VLOOKUP: 'SVERWEIS', VSTACK: 'VSTAPELN',
  WRAPCOLS: 'SPALTENUMBRUCH', WRAPROWS: 'ZEILENUMBRUCH', XLOOKUP: 'XVERWEIS', XMATCH: 'XVERGLEICH',
  // Information
  CELL: 'ZELLE', 'ERROR.TYPE': 'FEHLER.TYP', INFO: 'INFO', ISBLANK: 'ISTLEER', ISERR: 'ISTFEHL', ISERROR: 'ISTFEHLER', ISEVEN: 'ISTGERADE',
  ISFORMULA: 'ISTFORMEL', ISLOGICAL: 'ISTLOG', ISNA: 'ISTNV', ISNONTEXT: 'ISTKTEXT', ISNUMBER: 'ISTZAHL', ISODD: 'ISTUNGERADE',
  ISREF: 'ISTBEZUG', ISTEXT: 'ISTTEXT', N: 'N', NA: 'NV', SHEET: 'BLATT', SHEETS: 'BLÄTTER', TYPE: 'TYP',
  // Financial
  ACCRINT: 'AUFGELZINS', COUPDAYS: 'ZINSTERMTAGE', CUMIPMT: 'KUMZINSZ', CUMPRINC: 'KUMKAPITAL', DB: 'GDA2', DDB: 'GDA', DISC: 'DISAGIO',
  DOLLARDE: 'NOTIERUNGDEZ', DOLLARFR: 'NOTIERUNGBRU', EFFECT: 'EFFEKTIV', FV: 'ZW', FVSCHEDULE: 'ZW2', IPMT: 'ZINSZ', IRR: 'IKV',
  ISPMT: 'ISPMT', MIRR: 'QIKV', NOMINAL: 'NOMINAL', NPER: 'ZZR', NPV: 'NBW', PDURATION: 'PDURATION', PMT: 'RMZ', PPMT: 'KAPZ',
  PRICEDISC: 'KURSDISAGIO', PV: 'BW', RATE: 'ZINS', RRI: 'ZSATZINVEST', SLN: 'LIA', SYD: 'DIA', TBILLEQ: 'TBILLÄQUIV',
  TBILLPRICE: 'TBILLKURS', TBILLYIELD: 'TBILLRENDITE', XIRR: 'XINTZINSFUSS', XNPV: 'XKAPITALWERT',
  // Engineering
  BESSELI: 'BESSELI', BESSELJ: 'BESSELJ', BESSELK: 'BESSELK', BESSELY: 'BESSELY', BIN2DEC: 'BININDEZ', BIN2HEX: 'BININHEX', BIN2OCT: 'BININOKT',
  BITAND: 'BITUND', BITLSHIFT: 'BITLVERSCHIEB', BITOR: 'BITODER', BITRSHIFT: 'BITRVERSCHIEB', BITXOR: 'BITXODER', COMPLEX: 'KOMPLEXE',
  CONVERT: 'UMWANDELN', DEC2BIN: 'DEZINBIN', DEC2HEX: 'DEZINHEX', DEC2OCT: 'DEZINOKT', DELTA: 'DELTA', ERF: 'GAUSSFEHLER', ERFC: 'GAUSSFKOMPL',
  GESTEP: 'GGANZZAHL', HEX2BIN: 'HEXINBIN', HEX2DEC: 'HEXINDEZ', HEX2OCT: 'HEXINOKT', IMABS: 'IMABS', IMAGINARY: 'IMAGINÄRTEIL',
  IMARGUMENT: 'IMARGUMENT', IMCONJUGATE: 'IMKONJUGIERTE', IMCOS: 'IMCOS', IMCOSH: 'IMCOSHYP', IMCOT: 'IMCOT', IMCSC: 'IMCOSEC',
  IMCSCH: 'IMCOSECHYP', IMDIV: 'IMDIV', IMEXP: 'IMEXP', IMLN: 'IMLN', IMLOG10: 'IMLOG10', IMLOG2: 'IMLOG2', IMPOWER: 'IMAPOTENZ',
  IMPRODUCT: 'IMPRODUKT', IMREAL: 'IMREALTEIL', IMSEC: 'IMSEC', IMSECH: 'IMSECHYP', IMSIN: 'IMSIN', IMSINH: 'IMSINHYP', IMSQRT: 'IMWURZEL',
  IMSUB: 'IMSUB', IMSUM: 'IMSUMME', IMTAN: 'IMTAN', OCT2BIN: 'OKTINBIN', OCT2DEC: 'OKTINDEZ', OCT2HEX: 'OKTINHEX',
  // Database
  DAVERAGE: 'DBMITTELWERT', DCOUNT: 'DBANZAHL', DCOUNTA: 'DBANZAHL2', DGET: 'DBAUSZUG', DMAX: 'DBMAX', DMIN: 'DBMIN', DPRODUCT: 'DBPRODUKT',
  DSTDEV: 'DBSTDABW', DSTDEVP: 'DBSTDABWN', DSUM: 'DBSUMME', DVAR: 'DBVARIANZ', DVARP: 'DBVARIANZEN',
};
/** Older or alternative German spellings that are accepted on input */
const DE_ALIASES = { 'WÄHRUNG': 'DOLLAR', 'KGRÖßTE': 'LARGE', BOGENMAß: 'RADIANS', BESTIMMTHEITSMAß: 'RSQ', XINTZINSFUß: 'XIRR', 'GROß': 'UPPER', 'GROß2': 'PROPER' };

const DE_TO_EN = new Map([...Object.entries(FN_DE).map(([en, de]) => [de, en]), ...Object.entries(DE_ALIASES)]);

/** 'de' when function names, logical values and errors are shown in German (loc from format/locale.js) */
export const fnLangOf = (loc) => (loc && loc.fnLang === 'de' ? 'de' : 'en');
const langOf = (x) => (typeof x === 'string' ? (x === 'de' ? 'de' : 'en') : fnLangOf(x));

/** English (stored) function name → name shown to the user */
export function localFn(en, lang) {
  const up = String(en || '').toUpperCase();
  return langOf(lang) === 'de' ? FN_DE[up] || up : up;
}
/** Name typed by the user (any case, German or English) → English (stored) name */
export function internalFn(name, lang) {
  const up = String(name || '').toUpperCase();
  if (langOf(lang) !== 'de') return up;
  return DE_TO_EN.get(up) || up;
}

/* ---------- logical and error values ---------- */
export const BOOL_DE = { TRUE: 'WAHR', FALSE: 'FALSCH' };
export const localBool = (v, lang) => (langOf(lang) === 'de' ? (v ? 'WAHR' : 'FALSCH') : v ? 'TRUE' : 'FALSE');

export const ERR_DE = {
  '#NULL!': '#NULL!', '#DIV/0!': '#DIV/0!', '#VALUE!': '#WERT!', '#REF!': '#BEZUG!', '#NAME?': '#NAME?', '#NUM!': '#ZAHL!', '#N/A': '#NV',
  '#SPILL!': '#ÜBERLAUF!', '#CALC!': '#KALK!', '#GETTING_DATA': '#DATEN_ABRUFEN',
};
const ERR_EN = new Map(Object.entries(ERR_DE).map(([en, de]) => [de, en]));
/** German error codes, longest first (for the tokenizer) */
export const ERROR_CODES_DE = [...new Set(Object.values(ERR_DE))].filter((c) => !(c in ERR_DE)).sort((a, b) => b.length - a.length);
/** '#N/A' → '#NV' in German */
export const localErr = (code, lang) => (langOf(lang) === 'de' ? ERR_DE[String(code).toUpperCase()] || code : code);
/** '#NV' / '#nv' / '#N/A' → '#N/A' (unknown codes upper-cased) */
export function internalErr(code) {
  const up = String(code || '').toUpperCase();
  return ERR_EN.get(up) || up;
}

/* ---------- argument names (Microsoft's German argument names) ---------- */
const ARG_DE = {
  number: 'Zahl', value: 'Wert', values: 'Werte', text: 'Text', logical: 'Wahrheitswert', logical_test: 'Wahrheitstest',
  value_if_true: 'Wert_wenn_wahr', value_if_false: 'Wert_wenn_falsch', value_if_error: 'Wert_falls_Fehler', value_if_na: 'Wert_bei_NV',
  range: 'Bereich', criteria: 'Suchkriterien', sum_range: 'Summe_Bereich', criteria_range: 'Kriterien_Bereich', average_range: 'Mittelwert_Bereich',
  max_range: 'Max_Bereich', min_range: 'Min_Bereich', array: 'Matrix', num_digits: 'Anzahl_Stellen', divisor: 'Divisor', power: 'Potenz',
  bottom: 'Untere_Zahl', top: 'Obere_Zahl', significance: 'Schritt', mode: 'Modus', multiple: 'Vielfaches', function_num: 'Funktion',
  ref: 'Bezug', options: 'Optionen', rows: 'Zeilen', columns: 'Spalten', start: 'Anfang', step: 'Schritt', min: 'Min', max: 'Max',
  integer: 'ganze_Zahl', k: 'k', order: 'Reihenfolge', quart: 'Quartil', x: 'x', known_ys: 'Y_Werte', known_xs: 'X_Werte',
  expression: 'Ausdruck', result: 'Ergebnis', default: 'Standard', name: 'Name', calculation: 'Berechnung', delimiter: 'Trennzeichen',
  ignore_empty: 'Leer_ignorieren', num_chars: 'Anzahl_Zeichen', start_num: 'Erstes_Zeichen', find_text: 'Suchtext', within_text: 'Text',
  old_text: 'Alter_Text', new_text: 'Neuer_Text', instance_num: 'ntes_Auftreten', format_text: 'Textformat', col_delimiter: 'Spaltentrennzeichen',
  row_delimiter: 'Zeilentrennzeichen', number_times: 'Multiplikator', decimals: 'Dezimalstellen', no_commas: 'Keine_Punkte', year: 'Jahr',
  month: 'Monat', day: 'Tag', hour: 'Stunde', minute: 'Minute', second: 'Sekunde', serial_number: 'Zahl', return_type: 'Typ', date: 'Datum',
  start_date: 'Ausgangsdatum', end_date: 'Enddatum', months: 'Monate', unit: 'Einheit', holidays: 'Freie_Tage', days: 'Tage',
  date_text: 'Datumstext', basis: 'Basis', lookup_value: 'Suchkriterium', table_array: 'Matrix', col_index_num: 'Spaltenindex',
  range_lookup: 'Bereich_Verweis', row_index_num: 'Zeilenindex', lookup_array: 'Suchmatrix', return_array: 'Rückgabematrix',
  if_not_found: 'wenn_nicht_gefunden', match_mode: 'Vergleichsmodus', search_mode: 'Suchmodus', lookup_vector: 'Suchvektor',
  result_vector: 'Ergebnisvektor', match_type: 'Vergleichstyp', row_num: 'Zeile', column_num: 'Spalte', reference: 'Bezug', cols: 'Spalten',
  height: 'Höhe', width: 'Breite', ref_text: 'Bezug', a1: 'A1', abs_num: 'Abs', sheet_text: 'Tabellenname', index_num: 'Index',
  sort_index: 'Sortierindex', sort_order: 'Sortierreihenfolge', by_col: 'nach_Spalte', by_array: 'nach_Matrix', include: 'einschließen',
  if_empty: 'wenn_leer', link_location: 'Hyperlink_Adresse', friendly_name: 'Freundlicher_Name', exactly_once: 'genau_einmal', rate: 'Zins',
  nper: 'Zzr', pv: 'Bw', fv: 'Zw', pmt: 'Rmz', type: 'F', per: 'Zr', guess: 'Schätzwert', cost: 'Ansch_Wert', salvage: 'Restwert',
  life: 'Nutzungsdauer', from_unit: 'Von_Maßeinheit', to_unit: 'In_Maßeinheit', database: 'Datenbank', field: 'Datenbankfeld',
};
/** One argument placeholder ('[number2]', 'criteria_range1', '…') in German */
function argDe(a) {
  const m = /^(\[?)([a-z_]+?)(\d*)(\]?)$/i.exec(a.trim());
  if (!m) return a;
  const base = ARG_DE[m[2].toLowerCase()];
  return base ? `${m[1]}${base}${m[3]}${m[4]}` : a;
}
/** Argument list of funcinfo.js ('number1, [number2], …') in the given language, as an array */
export function localArgList(args, lang) {
  const list = String(args || '').split(/,\s*/).filter((x) => x !== '');
  return langOf(lang) === 'de' ? list.map(argDe) : list;
}

/** true for real Excel function names (the engine also knows undotted aliases like CEILINGMATH – not listed for users) */
export const isExcelName = (n) => Object.prototype.hasOwnProperty.call(FN_DE, String(n || '').toUpperCase());
