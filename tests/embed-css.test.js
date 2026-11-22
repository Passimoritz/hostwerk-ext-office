'use strict';
// Native ui CSS lives in Hostwerk's own page: no rule may restyle Hostwerk's kit (.btn.on) without an own prefix,
// and the per-file exceptions of the extension check (hostwerk-allow) must survive the build into the package.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const UI = path.join(__dirname, '..', 'ui');

test('no bare .btn.on rule in the shell stylesheet', () => {
  const css = fs.readFileSync(path.join(UI, 'src', 'style.css'), 'utf8').replace(/\/\*[\s\S]*?\*\//g, '');
  assert.ok(!/(^|[},])\s*\.btn\.on\s*[{,]/m.test(css), '.btn.on would recolour every active button in Hostwerk – use .btn.of-on');
});

test('built stylesheets keep the hostwerk-allow notes of their sources', { skip: !fs.existsSync(path.join(UI, 'index.css')) }, () => {
  const main = fs.readFileSync(path.join(UI, 'index.css'), 'utf8');
  assert.match(main.split('\n')[0], /built with/); // build head stays first (the check recognises built files by it)
  assert.match(main, /hostwerk-allow: layout – /);
  assert.match(main, /hostwerk-allow: colors – /);
  assert.ok(!/(^|[},])\.btn\.on[{,]/.test(main));
});
