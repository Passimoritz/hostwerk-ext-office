import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { readJet, jetText } from '../ui/src/editors/viewer/jet.js';

const fx = (n) => fs.readFileSync(new URL(`./fixtures/${n}`, import.meta.url));

for (const [file, format] of [['jet4.mdb', /Jet 4/], ['ace.accdb', /ACE/]]) {
  test(`${file}: tables, columns and rows without Access`, () => {
    const db = readJet(fx(file));
    assert.match(db.format, format);
    assert.deepEqual(db.tables.map((t) => t.name), ['Kunden', 'Produkte']);
    const k = db.tables[0];
    assert.deepEqual(k.columns.map((c) => c.name), ['ID', 'Name', 'Stadt', 'Umsatz', 'Seit', 'Aktiv', 'Punkte', 'Notiz']);
    assert.equal(k.rowCount, 3);
    assert.deepEqual(k.rows.map((r) => r.map(jetText)), [
      ['1', 'Müller GmbH', 'Köln', '1234.5', '2021-03-15', '✓', '42', 'Erster Kunde – sehr zuverlässig'],
      ['2', 'Smith Ltd', 'London', '99.99', '2023-11-01', '–', '7', ''],
      ['3', 'Ålesund AS', '', '0', '1999-12-31', '✓', '-5', ''],
    ]);
    assert.equal(k.rows[2][2], null);
    assert.equal(db.tables[1].rows[1][2], 0.04);
  });
}

test('accdb: rows over many data pages and a memo over several long-value pages', () => {
  const db = readJet(fx('big.accdb'));
  const t = db.tables[0];
  assert.equal(t.rowCount, 401);
  assert.equal(t.rows[0][2].length, 8100);
  assert.ok(t.rows[0][2].startsWith('Lorem ipsum dolor sit amet Lorem'));
  assert.equal(t.rows[400][1], 'row 400');
  assert.equal(readJet(fx('big.accdb'), { maxRows: 10 }).tables[0].truncated, true);
});

test('not a database', () => {
  assert.throws(() => readJet(new Uint8Array(4096)), /Not an Access database/);
});
