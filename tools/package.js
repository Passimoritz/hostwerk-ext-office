#!/usr/bin/env node
'use strict';
/*
 * The release package exactly as the CI builds it (.github/workflows/check.yml "Build package"): every file of the
 * repository (git ls-files – here also new files that are not committed yet, never ignored ones) except the repository
 * tooling (tests/, tools/, ui/src/, docs/*.md, ARCHITECTURE.md, package*.json …). The Hostwerk catalog validates the
 * unpacked release zip (hostwerk-public scripts/build-extension-catalog.js → sdk/check.js --strict) and Hostwerk checks
 * the unpacked package on install – so this is the content the 10 MB limit applies to.
 *
 *   node tools/package.js                 → list + size
 *   node tools/package.js --dir <folder>  → copy the package content into <folder> (e.g. to run check.js on it)
 *   node tools/package.js --zip <file>    → write the release zip
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const DIR = path.join(__dirname, '..');
// keep in sync with .github/workflows/check.yml (the grep -vE of the package step)
const EXCLUDE = /^(\.github\/|\.gitignore$|\.gitattributes$|tests\/|tools\/|scripts\/|ui\/src\/|node_modules\/|package(-lock)?\.json$|docs\/.*\.md$|ARCHITECTURE\.md$|.*\.todo\.json$)/;

function packageFiles(dir = DIR, { untracked = true } = {}) {
  let list;
  try {
    list = execFileSync('git', ['ls-files', '-z', ...(untracked ? ['-c', '-o', '--exclude-standard'] : [])], { cwd: dir, encoding: 'utf8', maxBuffer: 64 * 1048576 }).split('\0').filter(Boolean);
  } catch {
    // no git (unpacked copy): every file except the excluded ones
    list = [];
    const walk = (d, rel) => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { if (['.git', 'node_modules'].includes(e.name)) continue; const r = rel ? `${rel}/${e.name}` : e.name; if (e.isDirectory()) walk(path.join(d, e.name), r); else list.push(r); } };
    walk(dir, '');
  }
  return [...new Set(list)].filter((f) => !EXCLUDE.test(f) && fs.existsSync(path.join(dir, f))).sort();
}
const sizeOf = (files, dir = DIR) => files.reduce((n, f) => n + fs.statSync(path.join(dir, f)).size, 0);

module.exports = { packageFiles, sizeOf, EXCLUDE };

if (require.main === module) {
  const arg = (k) => { const i = process.argv.indexOf(k); return i > 0 ? process.argv[i + 1] : null; };
  const files = packageFiles();
  const size = sizeOf(files);
  const toDir = arg('--dir');
  const toZip = arg('--zip');
  if (toDir) {
    fs.rmSync(toDir, { recursive: true, force: true });
    for (const f of files) { const d = path.join(toDir, f); fs.mkdirSync(path.dirname(d), { recursive: true }); fs.copyFileSync(path.join(DIR, f), d); }
  }
  if (toZip) {
    const { makeZip } = require('../server/zip');
    const buf = makeZip(files.map((f) => ({ name: f, data: fs.readFileSync(path.join(DIR, f)) })));
    fs.mkdirSync(path.dirname(path.resolve(toZip)), { recursive: true });
    fs.writeFileSync(toZip, buf);
    console.log(`zip: ${toZip} (${(buf.length / 1048576).toFixed(2)} MB)`);
  }
  if (process.argv.includes('--list')) for (const f of files) console.log(f);
  console.log(`package: ${files.length} files, ${(size / 1048576).toFixed(2)} MB unpacked (limit 10 MB, Office budget 9.8 MB)${toDir ? ` → ${toDir}` : ''}`);
}
