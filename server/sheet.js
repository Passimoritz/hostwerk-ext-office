'use strict';
/*
 * Server-side spreadsheet helpers (pure, no dependencies): read cell values and append rows to xlsx/xlsm, ods and csv
 * while keeping everything else of the file byte-identical (zippatch). Used by forms (append), rules (evaluate),
 * live fields (cell of another file) and the data of scheduled reports.
 *
 *   readBook(buf, ext, { maxCells })  → { sheets: [{ name, cells: { A1: value } , rows }] }   (cached formula results)
 *   cellValue(book, sheet, 'B2')      → value | null
 *   rangeValues(book, sheet, 'A1:B3') → [values]
 *   appendRow(buf, ext, { sheet, values: [..] }) → { data: Buffer, row }   (values: string | number | boolean | null |
 *                                     { $date: 'YYYY-MM-DD[THH:MM:SS]' } – a real date/time cell: xlsx serial + date format,
 *                                     ODS office:date-value + date style, CSV ISO text)
 *   dateValue(v) · localStamp(date) · dateText(v)
 *   colName(0) → 'A' · parseRef('AB12') → { c: 27, r: 11 }
 */
const ZP = require('./zippatch');

const MAX_CELLS = 2e6;
const colName = (i) => { let s = ''; let n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
function parseRef(ref) { const m = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/i.exec(String(ref || '').trim()); if (!m || Number(m[2]) < 1) return null; let c = 0; for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64); return { c: c - 1, r: Number(m[2]) - 1 }; }
const unxml = (s) => String(s).replace(/&(lt|gt|amp|quot|apos|#x[0-9a-f]+|#\d+);/gi, (m, e) => ({ lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }[e.toLowerCase()] || (e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : m)));
// eslint-disable-next-line no-control-regex
const xmlText = (s) => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const textOf = (xml) => unxml(String(xml).replace(/<(\w+:)?(tab)\b[^>]*\/>/g, '\t').replace(/<(\w+:)?(br|line-break)\b[^>]*\/>/g, '\n').replace(/<[^>]+>/g, ''));

function zipMap(buf) { const m = new Map(); for (const e of ZP.entries(buf)) m.set(e.name, e); return m; }
const part = (z, n) => { const e = z.get(n); return e ? ZP.inflate(e).toString('utf8') : null; };
const resolveTarget = (base, t) => { if (t.startsWith('/')) return t.slice(1); const parts = base.split('/').slice(0, -1); for (const p of t.split('/')) { if (p === '..') parts.pop(); else if (p && p !== '.') parts.push(p); } return parts.join('/'); };
function rels(z, partName) {
  const i = partName.lastIndexOf('/');
  const rp = `${partName.slice(0, i + 1)}_rels/${partName.slice(i + 1)}.rels`;
  const x = part(z, rp) || '';
  const out = new Map();
  for (const m of x.matchAll(/<Relationship\b[^>]*>/g)) { const id = (/\bId="([^"]*)"/.exec(m[0]) || [])[1]; const t = (/\bTarget="([^"]*)"/.exec(m[0]) || [])[1]; const ty = (/\bType="([^"]*)"/.exec(m[0]) || [])[1] || ''; const mode = (/\bTargetMode="([^"]*)"/.exec(m[0]) || [])[1]; if (id && t && mode !== 'External') out.set(id, { target: resolveTarget(partName, t), type: ty }); }
  return out;
}

/* ---------------- xlsx ---------------- */
function xlsxSheets(z) {
  const wb = part(z, 'xl/workbook.xml');
  if (!wb) throw Object.assign(new Error('This is not an Excel workbook.'), { status: 422 });
  const r = rels(z, 'xl/workbook.xml');
  const out = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unxml((/\bname="([^"]*)"/.exec(m[0]) || [])[1] || '');
    const rid = (/\br:id="([^"]*)"/.exec(m[0]) || /\b\w+:id="([^"]*)"/.exec(m[0]) || [])[1];
    const t = rid && r.get(rid);
    if (t && /worksheet$/.test(t.type)) out.push({ name, part: t.target });
  }
  return out;
}
function sharedStrings(z) {
  const x = part(z, 'xl/sharedStrings.xml'); if (!x) return [];
  const out = [];
  for (const m of x.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) out.push(unxml(m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').replace(/<[^>]+>/g, '')));
  return out;
}
function xlsxCells(xml, sst, max) {
  const cells = {}; let n = 0; let rows = 0;
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    if (++n > max) break;
    const a = m[1]; const body = m[2] || '';
    const ref = (/\br="([A-Z]+\d+)"/.exec(a) || [])[1]; if (!ref) continue;
    const t = (/\bt="(\w+)"/.exec(a) || [])[1] || 'n';
    const v = (/<v>([\s\S]*?)<\/v>/.exec(body) || [])[1];
    let val = null;
    if (t === 's') val = v != null ? sst[Number(v)] ?? null : null;
    else if (t === 'inlineStr') val = textOf((/<is>([\s\S]*?)<\/is>/.exec(body) || [])[1] || '');
    else if (t === 'str' || t === 'e') val = v != null ? unxml(v) : null;
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') { const num = Number(v); val = Number.isFinite(num) ? num : unxml(v); }
    if (val !== null) { cells[ref] = val; const r = parseRef(ref).r + 1; if (r > rows) rows = r; }
  }
  return { cells, rows };
}

/* ---------------- ods ---------------- */
function odsTables(content, max) {
  const out = [];
  for (const t of content.matchAll(/<table:table\b([^>]*)>([\s\S]*?)<\/table:table>/g)) {
    const name = unxml((/table:name="([^"]*)"/.exec(t[1]) || [])[1] || `Sheet${out.length + 1}`);
    const cells = {}; let r = 0; let n = 0; let rows = 0;
    for (const row of t[2].matchAll(/<table:table-row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/table:table-row>)/g)) {
      const rep = Math.min(Number((/table:number-rows-repeated="(\d+)"/.exec(row[1]) || [])[1] || 1), 1e6);
      const body = row[2] || '';
      if (!/<text:p|office:value=/.test(body)) { r += rep; continue; }
      for (let k = 0; k < Math.min(rep, 1000); k++) {
        let c = 0;
        for (const cell of body.matchAll(/<table:(?:covered-)?table-cell\b([^>]*?)(?:\/>|>([\s\S]*?)<\/table:(?:covered-)?table-cell>)/g)) {
          const crep = Math.min(Number((/table:number-columns-repeated="(\d+)"/.exec(cell[1]) || [])[1] || 1), 16384);
          const a = cell[1]; const cb = cell[2] || '';
          const type = (/office:value-type="(\w+)"/.exec(a) || [])[1];
          let val = null;
          if (type === 'float' || type === 'percentage' || type === 'currency') val = Number((/office:value="([^"]*)"/.exec(a) || [])[1]);
          else if (type === 'boolean') val = (/office:boolean-value="(\w+)"/.exec(a) || [])[1] === 'true';
          else if (type === 'date') val = (/office:date-value="([^"]*)"/.exec(a) || [])[1];
          else if (type === 'time') val = (/office:time-value="([^"]*)"/.exec(a) || [])[1];
          else if (/<text:p/.test(cb)) val = [...cb.matchAll(/<text:p\b[^>]*>([\s\S]*?)<\/text:p>|<text:p\b[^>]*\/>/g)].map((p) => textOf(p[1] || '')).join('\n');
          if (val !== null && val !== '') for (let j = 0; j < Math.min(crep, 100); j++) { if (++n > max) break; cells[colName(c + j) + (r + 1)] = val; rows = r + 1; }
          c += crep;
        }
        r++;
      }
      r += Math.max(0, rep - Math.min(rep, 1000));
    }
    out.push({ name, cells, rows });
  }
  return out;
}

/* ---------------- csv ---------------- */
function sniffDelimiter(line) { const cand = [';', ',', '\t', '|']; let best = ','; let n = -1; for (const d of cand) { const k = line.split(d).length; if (k > n) { n = k; best = d; } } return best; }
function parseCsv(text, delim) {
  const rows = []; let row = []; let cur = ''; let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
    if (ch === '"') q = true; else if (ch === delim) { row.push(cur); cur = ''; } else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; } else cur += ch;
    if (rows.length > 200000) break;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}
const decodeText = (buf) => (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf ? buf.subarray(3).toString('utf8') : buf[0] === 0xff && buf[1] === 0xfe ? buf.subarray(2).toString('utf16le') : buf.toString('utf8'));

function readBook(buf, ext, { maxCells = MAX_CELLS } = {}) {
  const e = String(ext || '').toLowerCase();
  if (/^(xlsx|xlsm|xltx|xltm)$/.test(e)) {
    const z = zipMap(buf); const sst = sharedStrings(z);
    return { sheets: xlsxSheets(z).map((s) => { const x = part(z, s.part) || ''; return { name: s.name, ...xlsxCells(x, sst, maxCells) }; }) };
  }
  if (/^(ods|ots)$/.test(e)) { const z = zipMap(buf); return { sheets: odsTables(part(z, 'content.xml') || '', maxCells) }; }
  if (/^(csv|tsv|txt)$/.test(e)) {
    const text = decodeText(buf); const d = e === 'tsv' ? '\t' : sniffDelimiter(text.split(/\r?\n/, 1)[0] || '');
    const cells = {}; const rows = parseCsv(text, d);
    rows.forEach((r, ri) => r.forEach((v, ci) => { if (v !== '') { const num = /^-?\d+(\.\d+)?$/.test(v) ? Number(v) : null; cells[colName(ci) + (ri + 1)] = num != null ? num : v; } }));
    return { sheets: [{ name: 'Sheet1', cells, rows: rows.length }] };
  }
  throw Object.assign(new Error('Only xlsx, ods and csv spreadsheets can be read here.'), { status: 415 });
}
const sheetOf = (book, sheet) => (sheet == null || sheet === '' ? book.sheets[0] : book.sheets.find((s) => s.name === sheet) || (Number.isInteger(Number(sheet)) ? book.sheets[Number(sheet)] : null));
function cellValue(book, sheet, ref) { const s = sheetOf(book, sheet); const p = parseRef(ref); if (!s || !p) return null; const v = s.cells[colName(p.c) + (p.r + 1)]; return v === undefined ? null : v; }
function rangeValues(book, sheet, range) {
  const s = sheetOf(book, sheet); if (!s) return [];
  const [a, b] = String(range || '').split(':'); const p = parseRef(a); const q = parseRef(b || a); if (!p || !q) return [];
  const out = [];
  for (let r = Math.min(p.r, q.r); r <= Math.max(p.r, q.r) && out.length < 100000; r++) for (let c = Math.min(p.c, q.c); c <= Math.max(p.c, q.c); c++) { const v = s.cells[colName(c) + (r + 1)]; if (v !== undefined) out.push(v); }
  return out;
}

/* ---------------- append ---------------- */
/** { $date: 'YYYY-MM-DD' | 'YYYY-MM-DDTHH:MM[:SS]' } (local time, no zone) → { y, mo, d, h, mi, s, time } or null */
function dateValue(v) {
  if (!v || typeof v !== 'object' || typeof v.$date !== 'string') return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(v.$date.trim());
  if (!m) return null;
  const [y, mo, d, h, mi, s] = [m[1], m[2], m[3], m[4] || 0, m[5] || 0, m[6] || 0].map(Number);
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || h > 23 || mi > 59 || s > 59) return null;
  return { y, mo, d, h, mi, s, time: m[4] != null };
}
const p2 = (n) => String(n).padStart(2, '0');
/** the local time of a Date as { $date } (forms: "Submitted at") */
const localStamp = (t = new Date()) => ({ $date: `${t.getFullYear()}-${p2(t.getMonth() + 1)}-${p2(t.getDate())}T${p2(t.getHours())}:${p2(t.getMinutes())}:${p2(t.getSeconds())}` });
/** ISO text of a date value (CSV, notifications, pending rows): 2026-11-20 14:03:05 / 2026-11-20 – other values as text */
const dateText = (v) => { const d = dateValue(v); return d ? `${d.y}-${p2(d.mo)}-${p2(d.d)}${d.time ? ` ${p2(d.h)}:${p2(d.mi)}:${p2(d.s)}` : ''}` : v == null ? '' : typeof v === 'object' ? '' : String(v); };
/** Excel serial number (1900 system, or 1904 when the workbook says so) */
function excelSerial(d, date1904 = false) {
  const days = (Date.UTC(d.y, d.mo - 1, d.d) - (date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30))) / 86400000;
  return days + (d.h * 3600 + d.mi * 60 + d.s) / 86400;
}
const clean = (v) => (v == null ? '' : typeof v === 'number' || typeof v === 'boolean' ? v : dateValue(v) ? v : typeof v === 'object' ? '' : String(v).slice(0, 32000));

/** the cellXfs index of a date (numFmt 14) / date-time (numFmt 22) style – reused or added to styles.xml */
function xlsxDateStyle(styles, time) {
  const fmt = time ? 22 : 14;
  const m = /<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/.exec(styles);
  if (!m) return null;
  const xfs = [...m[2].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((x) => x[0]);
  const found = xfs.findIndex((x) => new RegExp(`\\bnumFmtId="${fmt}"`).test(x) && /\bfontId="0"/.test(x) && /\bfillId="0"/.test(x) && /\bborderId="0"/.test(x) && !/<alignment|<protection/.test(x));
  if (found >= 0) return { index: found, styles };
  const xf = `<xf numFmtId="${fmt}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`;
  const open = /\bcount="\d+"/.test(m[1]) ? m[1].replace(/\bcount="\d+"/, `count="${xfs.length + 1}"`) : `${m[1]} count="${xfs.length + 1}"`;
  return { index: xfs.length, styles: `${styles.slice(0, m.index)}<cellXfs${open}>${m[2]}${xf}</cellXfs>${styles.slice(m.index + m[0].length)}` };
}
const MIN_STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

function appendXlsx(buf, sheetName, values) {
  const z = zipMap(buf);
  const sheets = xlsxSheets(z);
  const s = sheetName ? sheets.find((x) => x.name === sheetName) : sheets[0];
  if (!s) throw Object.assign(new Error('The sheet {name} does not exist (any more).'), { status: 404, vars: { name: sheetName } });
  let xml = part(z, s.part);
  let last = 0;
  for (const m of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*?(\/>|>([\s\S]*?)<\/row>)/g)) if (m[2] !== '/>' && /<c\b/.test(m[3] || '')) last = Math.max(last, Number(m[1]));
  const r = last + 1;
  const changes = {};
  // date/time values: a real serial number with a date format (styles.xml gets the style once)
  const styleOf = {};
  if (values.some((v) => dateValue(v))) {
    const wbx = part(z, 'xl/workbook.xml') || '';
    styleOf.d1904 = /<workbookPr\b[^>]*\bdate1904="(1|true)"/.test(wbx);
    const sp = [...rels(z, 'xl/workbook.xml').values()].find((t) => /\/styles$/.test(t.type));
    let stylesName = sp ? sp.target : null;
    let styles = stylesName ? part(z, stylesName) : null;
    const before = styles;
    if (!styles) {
      // no styles part at all (rare, hand-made files): add a minimal one
      stylesName = 'xl/styles.xml'; styles = MIN_STYLES;
      const rp = 'xl/_rels/workbook.xml.rels'; const rx = part(z, rp);
      if (rx) changes[rp] = rx.replace(/<\/Relationships>/, '<Relationship Id="rIdOfStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
      const ct = part(z, '[Content_Types].xml');
      if (ct && !/PartName="\/xl\/styles\.xml"/.test(ct)) changes['[Content_Types].xml'] = ct.replace(/<\/Types>/, '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
    }
    for (const time of [true, false]) {
      if (!values.some((v) => { const d = dateValue(v); return d && d.time === time; })) continue;
      const res = xlsxDateStyle(styles, time);
      if (res) { styles = res.styles; styleOf[time] = res.index; }
    }
    if (styles !== before) changes[stylesName] = styles;
  }
  const cells = values.map((v0, i) => {
    const v = clean(v0); const ref = `${colName(i)}${r}`;
    if (v === '') return '';
    const dv = dateValue(v);
    if (dv) { const sIdx = styleOf[dv.time]; return `<c r="${ref}"${sIdx != null ? ` s="${sIdx}"` : ''}><v>${Math.round(excelSerial(dv, styleOf.d1904) * 1e10) / 1e10}</v></c>`; }
    if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}"><v>${v}</v></c>`;
    if (typeof v === 'boolean') return `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`;
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlText(v)}</t></is></c>`;
  }).join('');
  const rowXml = `<row r="${r}">${cells}</row>`;
  // an existing (empty) row element with this number is replaced; otherwise insert in order
  const existing = new RegExp(`<row\\b[^>]*\\br="${r}"[^>]*?(\\/>|>[\\s\\S]*?<\\/row>)`).exec(xml);
  if (existing) xml = xml.replace(existing[0], rowXml);
  else {
    const after = [...xml.matchAll(/<row\b[^>]*\br="(\d+)"/g)].find((m) => Number(m[1]) > r);
    if (after) xml = xml.slice(0, after.index) + rowXml + xml.slice(after.index);
    else if (/<sheetData\s*\/>/.test(xml)) xml = xml.replace(/<sheetData\s*\/>/, `<sheetData>${rowXml}</sheetData>`);
    else xml = xml.replace(/<\/sheetData>/, `${rowXml}</sheetData>`);
  }
  // date columns without an own width: wide enough for "20.11.2026 14:03" (Excel shows ##### otherwise)
  const dateCols = values.map((v, i) => { const d = dateValue(v); return d ? { i, w: d.time ? 17 : 11 } : null; }).filter(Boolean);
  if (dateCols.length) {
    const cols = /<cols>([\s\S]*?)<\/cols>/.exec(xml);
    const ranges = cols ? [...cols[1].matchAll(/<col\b[^>]*\bmin="(\d+)"[^>]*\bmax="(\d+)"/g)].map((m) => [Number(m[1]), Number(m[2])]) : [];
    const add = dateCols.filter(({ i }) => !ranges.some(([a, b]) => i + 1 >= a && i + 1 <= b)).map(({ i, w }) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`);
    if (add.length) {
      if (cols) {
        const all = [...cols[1].matchAll(/<col\b[^>]*\/>/g)].map((m) => m[0]).concat(add).sort((a, b) => Number(/min="(\d+)"/.exec(a)[1]) - Number(/min="(\d+)"/.exec(b)[1]));
        xml = xml.replace(cols[0], `<cols>${all.join('')}</cols>`);
      } else xml = xml.replace(/<sheetData\b/, `<cols>${add.join('')}</cols><sheetData`);
    }
  }
  // dimension
  xml = xml.replace(/<dimension\b[^>]*\bref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"/, (m, c1, r1, c2, r2) => { const endC = c2 || c1; const endR = Math.max(Number(r2 || r1), r); const wide = colName(Math.max(values.length - 1, parseRef(`${endC}1`).c)); return m.replace(/ref="[^"]*"/, `ref="${c1}${r1}:${wide}${endR}"`); });
  changes[s.part] = xml;
  // tables (ListObjects) that end on the row above grow by one row
  for (const [, t] of rels(z, s.part)) {
    if (!/\/table$/.test(t.type)) continue;
    let tx = part(z, t.target); if (!tx) continue;
    const m = /<table\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/.exec(tx);
    if (m && Number(m[4]) === last) {
      tx = tx.replace(`ref="${m[1]}${m[2]}:${m[3]}${m[4]}"`, `ref="${m[1]}${m[2]}:${m[3]}${r}"`).replace(/(<autoFilter\b[^>]*\bref=")[A-Z]+\d+:[A-Z]+\d+"/, `$1${m[1]}${m[2]}:${m[3]}${r}"`);
      changes[t.target] = tx;
    }
  }
  return { data: ZP.patchZip(buf, changes), row: r };
}

function appendOds(buf, sheetName, values) {
  const z = zipMap(buf);
  let content = part(z, 'content.xml');
  if (!content) throw Object.assign(new Error('This is not an OpenDocument spreadsheet.'), { status: 422 });
  const tables = [...content.matchAll(/<table:table\b([^>]*)>([\s\S]*?)<\/table:table>/g)];
  const t = sheetName ? tables.find((m) => unxml((/table:name="([^"]*)"/.exec(m[1]) || [])[1] || '') === sheetName) : tables[0];
  if (!t) throw Object.assign(new Error('The sheet {name} does not exist (any more).'), { status: 404, vars: { name: sheetName } });
  const body = t[2];
  let insertAt = -1; let rowNo = 0; let lastRow = 0;
  for (const row of body.matchAll(/<table:table-row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/table:table-row>)/g)) {
    const rep = Number((/table:number-rows-repeated="(\d+)"/.exec(row[1]) || [])[1] || 1);
    rowNo += rep;
    if (/<text:p|office:value=/.test(row[2] || '')) { insertAt = row.index + row[0].length; lastRow = rowNo; }
  }
  if (insertAt < 0) { const h = /<table:table-row\b/.exec(body); insertAt = h ? h.index : body.length; }
  // date/time values: office:date-value + a date style in the automatic styles (added once)
  const dstyle = {};
  for (const time of [true, false]) {
    if (!values.some((v) => { const d = dateValue(v); return d && d.time === time; })) continue;
    const ce = time ? 'ceOfDateTime' : 'ceOfDate'; const ns = time ? 'NOfDateTime' : 'NOfDate';
    if (!content.includes(`style:name="${ce}"`)) {
      const sep = (x) => `<number:text>${x}</number:text>`;
      const num = `<number:date-style style:name="${ns}"><number:day number:style="long"/>${sep('.')}<number:month number:style="long"/>${sep('.')}<number:year number:style="long"/>${time ? `${sep(' ')}<number:hours number:style="long"/>${sep(':')}<number:minutes number:style="long"/>${sep(':')}<number:seconds number:style="long"/>` : ''}</number:date-style>`;
      const st = `${num}<style:style style:name="${ce}" style:family="table-cell" style:data-style-name="${ns}"/>`;
      if (/<office:automatic-styles\s*\/>/.test(content)) content = content.replace(/<office:automatic-styles\s*\/>/, `<office:automatic-styles>${st}</office:automatic-styles>`);
      else if (/<\/office:automatic-styles>/.test(content)) content = content.replace(/<\/office:automatic-styles>/, `${st}</office:automatic-styles>`);
      else content = content.replace(/<office:body>/, `<office:automatic-styles>${st}</office:automatic-styles><office:body>`);
      if (!/xmlns:number=/.test(content)) content = content.replace(/<office:document-content\b/, '<office:document-content xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"');
      if (!/xmlns:style=/.test(content)) content = content.replace(/<office:document-content\b/, '<office:document-content xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"');
    }
    dstyle[time] = ce;
  }
  // the table again (the styles above moved it)
  const t2 = [...content.matchAll(/<table:table\b([^>]*)>([\s\S]*?)<\/table:table>/g)][tables.indexOf(t)] || t;
  const cells = values.map((v0) => {
    const v = clean(v0);
    if (v === '') return '<table:table-cell/>';
    const dv = dateValue(v);
    if (dv) {
      const iso = `${dv.y}-${p2(dv.mo)}-${p2(dv.d)}${dv.time ? `T${p2(dv.h)}:${p2(dv.mi)}:${p2(dv.s)}` : ''}`;
      const shown = `${p2(dv.d)}.${p2(dv.mo)}.${dv.y}${dv.time ? ` ${p2(dv.h)}:${p2(dv.mi)}:${p2(dv.s)}` : ''}`;
      return `<table:table-cell table:style-name="${dstyle[dv.time]}" office:value-type="date" office:date-value="${iso}"><text:p>${shown}</text:p></table:table-cell>`;
    }
    if (typeof v === 'number' && Number.isFinite(v)) return `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
    if (typeof v === 'boolean') return `<table:table-cell office:value-type="boolean" office:boolean-value="${v}"><text:p>${v ? 'TRUE' : 'FALSE'}</text:p></table:table-cell>`;
    return `<table:table-cell office:value-type="string"><text:p>${xmlText(v)}</text:p></table:table-cell>`;
  }).join('');
  let nb = `${body.slice(0, insertAt)}<table:table-row>${cells}</table:table-row>${body.slice(insertAt)}`;
  // one repeated empty row after the data shrinks by one, so the sheet keeps its size
  nb = nb.replace(/(<table:table-row>[\s\S]*?<\/table:table-row>)(<table:table-row\b[^>]*table:number-rows-repeated=")(\d+)"/, (m, a, b, n) => (Number(n) > 1 ? `${a}${b}${Number(n) - 1}"` : m));
  content = content.slice(0, t2.index) + t2[0].replace(body, () => nb) + content.slice(t2.index + t2[0].length);
  return { data: ZP.patchZip(buf, { 'content.xml': content }, { first: ['mimetype'] }), row: lastRow + 1 };
}

function appendCsv(buf, values, ext) {
  const text = decodeText(buf);
  const bom = buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf;
  const d = ext === 'tsv' ? '\t' : sniffDelimiter(text.split(/\r?\n/, 1)[0] || '');
  const nl = /\r\n/.test(text) ? '\r\n' : '\n';
  const q = (v0) => { let v = dateValue(v0) ? dateText(v0) : String(clean(v0)); if (/^[=+\-@\t\r]/.test(v) && typeof v0 !== 'number') v = `'${v}`; return /["\r\n]/.test(v) || v.includes(d) ? `"${v.replace(/"/g, '""')}"` : v; };
  const line = values.map(q).join(d);
  const out = `${text}${text && !/\r?\n$/.test(text) ? nl : ''}${line}${nl}`;
  const rows = parseCsv(text, d).filter((r) => r.some((x) => x !== '')).length;
  return { data: Buffer.concat([bom ? Buffer.from([0xef, 0xbb, 0xbf]) : Buffer.alloc(0), Buffer.from(out, 'utf8')]), row: rows + 1 };
}

function appendRow(buf, ext, { sheet = '', values = [] } = {}) {
  if (!Array.isArray(values) || !values.length || values.length > 200) throw Object.assign(new Error('Invalid row'), { status: 400 });
  const e = String(ext || '').toLowerCase();
  if (/^(xlsx|xlsm)$/.test(e)) return appendXlsx(buf, sheet, values);
  if (e === 'ods') return appendOds(buf, sheet, values);
  if (e === 'csv' || e === 'tsv') return appendCsv(buf, values, e);
  throw Object.assign(new Error('Forms work with xlsx, ods and csv files.'), { status: 415 });
}

module.exports = { dateValue, localStamp, dateText, excelSerial, readBook, cellValue, rangeValues, appendRow, colName, parseRef, parseCsv, sheetOf, xmlText, unxml, textOf, rels, zipMap, part };
