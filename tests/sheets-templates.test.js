import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { buildTemplate } from '../ui/src/editors/sheets/templates.js';
import { makeLocale } from '../ui/src/editors/sheets/format/locale.js';
import { saveBytes, openBytes, csvToBook, bookToHTML } from '../ui/src/editors/sheets/io/formats.js';
import { isErr } from '../ui/src/editors/sheets/formula/errors.js';
import { encodeText } from '../ui/src/editors/sheets/io/csv.js';

for (const lang of ['en', 'de']) {
  const loc = makeLocale({ lang });
  for (const id of ['blank', 'budget', 'invoice', 'schedule', 'tracker', 'inventory']) {
    test(`template ${id} (${lang}) builds, has no formula errors and saves as xlsx/ods/csv`, async () => {
      const b = buildTemplate(id, { loc, lang });
      for (const sh of b.sheets) for (const [k, c] of sh.cells) if (c.f) assert.ok(!isErr(c.v), `${id} ${k} ${c.f} → ${c.v && c.v.code}`);
      const { bytes } = await saveBytes(b, 'xlsx', {});
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(Buffer.from(bytes));
      assert.equal(wb.worksheets.length, 1);
      const again = await openBytes(bytes, 'xlsx', { loc });
      assert.equal(again.sheets[0].name, b.sheets[0].name);
      assert.ok((await saveBytes(b, 'ods', {})).bytes.length > 100);
      assert.ok((await saveBytes(b, 'csv', {})).bytes.length >= 0);
    });
  }
}

test('csv import with German numbers, header, then html export', () => {
  const loc = makeLocale({ lang: 'de' });
  const bytes = encodeText('Name;Betrag;Datum\r\n"Müller; Hans";1.234,50;18.11.2026\r\nSchmidt;-7,5;01.02.2025\r\n', 'windows-1252');
  const b = csvToBook(bytes, { loc, header: true });
  const s = b.sheets[0];
  assert.equal(b.value(s, 1, 0), 'Müller; Hans');
  assert.equal(b.value(s, 1, 1), 1234.5);
  assert.equal(b.text(s, 1, 2), '18.11.2026');
  assert.equal(b.csv.delimiter, ';');
  const html = bookToHTML(b, {});
  assert.match(html, /Müller; Hans/);
  assert.doesNotMatch(html, /<script/i);
});
