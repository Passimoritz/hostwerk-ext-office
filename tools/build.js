#!/usr/bin/env node
'use strict';
/*
 * Build of the Office extension.
 *
 *   node tools/build.js            → ui/index.js (light main module, via sdk/build-native.js)
 *                                    + ui/assets/<name>.js (+ .css) for every ui/src/editors/<name>/index.jsx
 *   node tools/build.js docs       → only the "docs" editor chunk (faster while working on one editor)
 *
 * Editor chunks are heavy (format libraries) and load lazily: ui/src/lib/loader.js imports
 * /api/office/assets/<name>.js on first use. They share React and Hostwerk's modules exactly like the main module
 * (globalThis.__hostwerk), so JSX text is translated the same way.
 */
const fs = require('fs');
const path = require('path');

const DIR = path.join(__dirname, '..');
// SDK build script: HW_BUILD_NATIVE (CI downloads sdk/build-native.js from hostwerk-public), else the local Hostwerk checkout
const SDK = process.env.HW_BUILD_NATIVE || path.join(DIR, '..', '..', 'hostwerk', 'public-repo', 'sdk', 'build-native.js');
const { build: buildMain, HOST_MODULES } = require(SDK);
const esbuild = require(require.resolve('esbuild', { paths: [DIR, path.join(DIR, '..')] }));
const HEAVY = ['@hostwerk/kit', '@hostwerk/editor', '@hostwerk/files', '@hostwerk/three'];
const ID = JSON.parse(fs.readFileSync(path.join(DIR, 'manifest.json'), 'utf8')).id;

function hostPlugin(used) {
  return {
    name: 'hostwerk-host',
    setup(b) {
      const filter = /^(react|react-dom|react\/jsx-runtime|react\/jsx-dev-runtime|@hostwerk\/[\w-]+(\/jsx-runtime|\/jsx-dev-runtime)?)$/;
      b.onResolve({ filter }, (a) => ({ path: a.path, namespace: 'hostwerk' }));
      b.onLoad({ filter: /.*/, namespace: 'hostwerk' }, (a) => {
        const p = a.path;
        const ext = `globalThis.__hostwerk.forExt(${JSON.stringify(ID)})`;
        if (/jsx-(dev-)?runtime$/.test(p)) return { contents: `module.exports = ${ext}.runtime;`, loader: 'js' };
        if (p === '@hostwerk/ui') return { contents: `module.exports = ${ext}.ui;`, loader: 'js' };
        if (p === '@hostwerk/i18n') return { contents: `module.exports = { t: ${ext}.t, pl: ${ext}.pl };`, loader: 'js' };
        if (!HOST_MODULES.includes(p)) throw new Error(`${p} is not offered by Hostwerk`);
        used.add(p);
        return { contents: `module.exports = globalThis.__hostwerk.mods[${JSON.stringify(p)}];`, loader: 'js' };
      });
    },
  };
}

/*
 * Libraries that several editor chunks use are built ONCE as ui/assets/lib-<name>.js (jszip: docs, sheets, slides,
 * viewer · pdf.js: pdf, viewer) – before, every chunk carried its own copy (≈ 0.6 MB more in the package). A chunk that
 * needs one loads it first (top-level await, same ?v= as the chunk) into globalThis.__officeLib; the import inside the
 * chunk is a stub that reads from there. pdf.js finds its worker next to lib-pdfjs.js (same folder, same ?v=).
 */
// yjs + y-protocols (Office 1.1, co-editing): ONE instance for the shell's provider (ui/src/lib/collab.js) and every
// chunk – also y-prosemirror's own `import 'yjs'` resolves to it (two Yjs copies break instanceof checks). One file
// lib-yjs.js serves three module names; `pick` = the member of its default export.
const SHARED = {
  jszip: { key: 'jszip', file: 'lib-jszip', contents: "import JSZip from 'jszip'; export default JSZip;" },
  'pdfjs-dist': { key: 'pdfjs', file: 'lib-pdfjs', contents: "import * as pdfjs from 'pdfjs-dist'; export default pdfjs;" },
  yjs: { key: 'yjs', file: 'lib-yjs', pick: 'Y', contents: "export { default } from './ui/src/lib/collab.js';" },
  'y-protocols/sync': { key: 'ysync', file: 'lib-yjs', pick: 'syncProtocol' },
  'y-protocols/awareness': { key: 'yawareness', file: 'lib-yjs', pick: 'awarenessProtocol' },
};
const keyOf = (m) => SHARED[m].key;
function sharedPlugin(usedLibs) {
  return {
    name: 'office-shared-libs',
    setup(b) {
      b.onResolve({ filter: /^(jszip|pdfjs-dist|yjs|y-protocols\/sync|y-protocols\/awareness)$/ }, (a) => ({ path: a.path, namespace: 'office-lib' }));
      b.onLoad({ filter: /.*/, namespace: 'office-lib' }, (a) => { usedLibs.add(a.path); return { contents: `module.exports = globalThis.__officeLib[${JSON.stringify(keyOf(a.path))}];`, loader: 'js' }; });
    },
  };
}
const COMMON = {
  bundle: true, format: 'esm', platform: 'browser', target: ['chrome120'], write: false, minify: true, legalComments: 'none', charset: 'utf8',
  define: { 'process.env.NODE_ENV': '"production"', global: 'globalThis' },
  nodePaths: [path.join(DIR, 'node_modules'), path.join(DIR, '..', 'node_modules')], logLevel: 'warning',
};
async function buildShared(mod) {
  const def = SHARED[mod];
  if (!def.contents) return; // another name of a file built under its first name (lib-yjs)
  const out = path.join(DIR, 'ui', 'assets', `${def.file}.js`);
  const res = await esbuild.build({ ...COMMON, stdin: { contents: def.contents, resolveDir: DIR, loader: 'js' }, outfile: out });
  const js = res.outputFiles.find((f) => /\.js$/.test(f.path));
  fs.mkdirSync(path.dirname(out), { recursive: true });
  const EV = new RegExp(`\\b${'ev'}${'al'}\\s*\\(`);
  if (EV.test(js.text)) console.warn(`! ${def.file}: bundle contains code evaluation – the catalog check will reject it`);
  fs.writeFileSync(out, `/* ${ID} shared library "${mod}" of the editor chunks – built with tools/build.js */\n${js.text}`);
  console.log(`✓ ui/assets/${def.file}.js (${Math.round(js.text.length / 1024)} KB, shared)`);
}

/** --analyze: bytes per input module and chunk (esbuild metafile) → packages bundled into more than one chunk */
const ANALYZE = process.argv.includes('--analyze');
const META = new Map(); // package or source folder → Map(chunk → bytes)
function collect(name, res) {
  if (!res.metafile) return;
  for (const o of Object.values(res.metafile.outputs)) {
    for (const [inp, v] of Object.entries(o.inputs || {})) {
      const p = inp.replace(/\\/g, '/');
      const nm = p.lastIndexOf('node_modules/');
      let key;
      if (nm >= 0) { const rest = p.slice(nm + 13).split('/'); key = `npm:${rest[0].startsWith('@') ? `${rest[0]}/${rest[1]}` : rest[0]}`; } else key = process.env.ANALYZE_FILES ? p : p.split('/').slice(0, 4).join('/');
      if (!META.has(key)) META.set(key, new Map());
      META.get(key).set(name, (META.get(key).get(name) || 0) + v.bytesInOutput);
    }
  }
}
function analyzeReport() {
  const K = (n) => Math.round(n / 1024);
  const rows = [...META].map(([k, m]) => { const l = [...m].sort((a, b) => b[1] - a[1]); const all = l.reduce((s, x) => s + x[1], 0); return { k, l, all, wasted: all - (l[0] ? l[0][1] : 0) }; });
  console.log('\nCode bundled into more than one chunk (wasted = every copy but the largest):');
  for (const r of rows.filter((x) => x.wasted > 2048).sort((a, b) => b.wasted - a.wasted)) console.log(`  ${r.k.padEnd(40)} wasted ${K(r.wasted)} KB  ${r.l.map(([c, b]) => `${c} ${K(b)}`).join(' · ')}`);
  console.log('\nLargest packages / source folders (all chunks):');
  for (const r of rows.sort((a, b) => b.all - a.all).slice(0, 25)) console.log(`  ${r.k.padEnd(40)} ${K(r.all)} KB  ${r.l.slice(0, 4).map(([c, b]) => `${c} ${K(b)}`).join(' · ')}`);
}

// Hostwerk's extension check honours per-file exceptions (a CSS comment "hostwerk-allow: colors|layout|… – reason").
// The minifier drops comments, so a package (no ui/src) would lose them: the allow notes of every source stylesheet
// are carried into the head of the built .css they end up in.
function allowNotes(cssFiles) {
  const out = new Set();
  for (const f of cssFiles) {
    let src = '';
    try { src = fs.readFileSync(f, 'utf8'); } catch { continue; }
    for (const c of src.match(/\/\*[\s\S]*?\*\//g) || []) if (/hostwerk-allow:/.test(c)) out.add(c.replace(/\s+/g, ' '));
  }
  return out.size ? `${[...out].join('\n')}\n` : '';
}
function addAllowNotes(cssOut, cssFiles) {
  if (!fs.existsSync(cssOut)) return;
  const notes = allowNotes(cssFiles);
  if (!notes) return;
  const text = fs.readFileSync(cssOut, 'utf8');
  const nl = text.indexOf('\n') + 1; // keep the build head as the first line (the check recognises built files by it)
  fs.writeFileSync(cssOut, text.slice(0, nl) + notes + text.slice(nl));
}

async function buildChunk(name) {
  const entry = ['index.jsx', 'index.js'].map((f) => path.join(DIR, 'ui', 'src', 'editors', name, f)).find((f) => fs.existsSync(f));
  if (!entry) return;
  const out = path.join(DIR, 'ui', 'assets', `${name}.js`);
  const used = new Set();
  const libs = new Set();
  const res = await esbuild.build({
    ...COMMON, entryPoints: [entry], outfile: out, jsx: 'automatic', jsxImportSource: '@hostwerk', metafile: true,
    loader: { '.woff': 'dataurl', '.woff2': 'dataurl', '.ttf': 'dataurl', '.png': 'dataurl', '.svg': 'text' },
    // the pdf.js worker runs on its own (no window, no globalThis.__officeLib) – it keeps its copy
    plugins: [hostPlugin(used), ...(name === 'pdfworker' ? [] : [sharedPlugin(libs)])],
  });
  collect(name, res);
  const js = res.outputFiles.find((f) => /\.js$/.test(f.path));
  const css = res.outputFiles.find((f) => /\.css$/.test(f.path));
  fs.mkdirSync(path.dirname(out), { recursive: true });
  const need = [...used].filter((x) => HEAVY.includes(x));
  // shared libraries first (lib-jszip.js / lib-pdfjs.js next to this chunk, same ?v=) – the stubs read them at once
  const shared = [...libs].map((m) => [keyOf(m), SHARED[m].file, SHARED[m].pick || '']);
  const loadLibs = shared.length ? `{ const L = globalThis.__officeLib || (globalThis.__officeLib = {}); const me = new URL(import.meta.url); await Promise.all(${JSON.stringify(shared)}.filter(([k]) => !L[k]).map(async ([k, f, p]) => { const u = new URL(f + '.js', me); u.search = me.search; const d = (await import(u.href)).default; L[k] = p ? d[p] : d; })); }\n` : '';
  const head = `/* ${ID} editor chunk "${name}" – built with tools/build.js. Source: ui/src/editors/${name} */\n${need.length ? `await globalThis.__hostwerk.need(${JSON.stringify(need)});\n` : ''}${loadLibs}`;
  // The catalog check rejects code evaluation in bundles – warn early (pattern built at runtime so this file itself passes the check)
  const EV = new RegExp(`\\b${'ev'}${'al'}\\s*\\(`);
  if (EV.test(js.text)) console.warn(`! ${name}: bundle contains code evaluation – the catalog check will reject it`);
  fs.writeFileSync(out, head + js.text);
  const cssOut = out.replace(/\.js$/, '.css');
  // same build head as the .js (Hostwerk's extension check and ui-consistency-audit skip built files by it – their source is ui/src)
  if (css) fs.writeFileSync(cssOut, `/* ${ID} editor chunk "${name}" styles – built with tools/build.js. Source: ui/src/editors/${name} */\n${css.text}`); else fs.rmSync(cssOut, { force: true });
  if (css) addAllowNotes(cssOut, Object.keys(res.metafile.inputs).filter((f) => /\.css$/.test(f) && !/node_modules/.test(f)).map((f) => path.resolve(DIR, f)));
  console.log(`✓ ui/assets/${name}.js (${Math.round((head.length + js.text.length) / 1024)} KB${css ? `, css ${Math.round(css.text.length / 1024)} KB` : ''})`);
}

(async () => {
  await null; // the size report below is defined after this block
  const only = process.argv.slice(2).find((a) => !a.startsWith('--')) || (process.argv.includes('--sizes') ? '--sizes' : undefined);
  if (only === '--sizes') { report(); return; }
  const editors = fs.existsSync(path.join(DIR, 'ui', 'src', 'editors')) ? fs.readdirSync(path.join(DIR, 'ui', 'src', 'editors')).filter((n) => !n.startsWith('_')) : [];
  // Version of the editor chunks for the loader (?v=…): changes with every build, so a new package never uses old chunks
  const ver = `${JSON.parse(fs.readFileSync(path.join(DIR, 'manifest.json'), 'utf8')).version}-${Date.now().toString(36)}`;
  if (!only || only === 'main') {
    fs.writeFileSync(path.join(DIR, 'ui', 'src', 'lib', 'version.js'), `// generated by tools/build.js – do not edit\nexport const CHUNK_VERSION = ${JSON.stringify(ver)};\n`);
    await buildMain(DIR);
    addAllowNotes(path.join(DIR, 'ui', 'index.css'), [path.join(DIR, 'ui', 'src', 'style.css')]);
  }
  if (!only || only === 'shared') for (const m of Object.keys(SHARED)) await buildShared(m);
  for (const n of editors) if (!only || only === n) await buildChunk(n);
  report();
  if (ANALYZE) analyzeReport();
})().catch((e) => { console.error(e.message); process.exit(1); });

/*
 * Size report: every chunk with its growth against the 1.0 sizes (budget: ≤ 200 KB more per editor), the shared
 * libraries, the main module and the whole package against the 9.8 MB budget of Office 1.1 (the catalog allows 10 MB).
 * `node tools/build.js --sizes` prints it without building.
 */
const BUDGET = 9.8 * 1048576;
const CHUNK_GROWTH = 200 * 1024;
// sizes of Office 1.0.1 (ARCHITECTURE.md "Bundles") – the base of the growth column
const BASE_KB = { docs: 820, pdf: 525, pdfworker: 1293, sheets: 1111, slides: 459, viewer: 172, 'lib-jszip': 96, 'lib-pdfjs': 324, main: 169 };
function report() {
  const kb = (n) => `${Math.round(n / 1024)} KB`;
  const assets = path.join(DIR, 'ui', 'assets');
  const rows = [];
  if (fs.existsSync(assets)) {
    for (const f of fs.readdirSync(assets).filter((x) => x.endsWith('.js')).sort()) {
      const name = f.replace(/\.js$/, '');
      const js = fs.statSync(path.join(assets, f)).size;
      const cssF = path.join(assets, `${name}.css`);
      const css = fs.existsSync(cssF) ? fs.statSync(cssF).size : 0;
      const base = BASE_KB[name] != null ? BASE_KB[name] * 1024 : null;
      const grow = base == null ? null : js - base; // the 1.0.1 numbers are the .js alone
      rows.push({ name, js, css, grow });
    }
  }
  const main = ['index.js', 'index.css'].map((f) => path.join(DIR, 'ui', f)).filter((f) => fs.existsSync(f)).reduce((s, f) => s + fs.statSync(f).size, 0);
  // what ships = the release zip's content (tools/package.js, same rule as the CI package step) – the catalog and
  // Hostwerk's install check measure that; the full checkout (with ui/src, tests …) is shown for information
  const { packageFiles, sizeOf } = require('./package');
  const total = sizeOf(packageFiles());
  let checkout = 0;
  const walk = (d) => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { if (['node_modules', '.git', '.github'].includes(e.name)) continue; const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else checkout += fs.statSync(p).size; } };
  walk(DIR);
  console.log('\nChunk sizes (js + css, growth against 1.0.1):');
  const mainJs = fs.existsSync(path.join(DIR, 'ui', 'index.js')) ? fs.statSync(path.join(DIR, 'ui', 'index.js')).size : 0;
  console.log(`  ${'main (ui/index.*)'.padEnd(18)} ${kb(main).padStart(8)}  +${kb(mainJs - BASE_KB.main * 1024)}`);
  for (const r of rows) {
    const over = r.grow != null && r.grow > CHUNK_GROWTH && !r.name.startsWith('lib-') && r.name !== 'pdfworker';
    console.log(`  ${r.name.padEnd(18)} ${kb(r.js + r.css).padStart(8)}${r.css ? `  (css ${kb(r.css)})` : ''}${r.grow != null ? `  ${r.grow >= 0 ? '+' : ''}${kb(r.grow)}` : '  new'}${over ? '  ! more than +200 KB' : ''}`);
  }
  const pct = Math.round((total / BUDGET) * 100);
  console.log(`Package size (shipped, as the release zip): ${(total / 1048576).toFixed(2)} MB of 9.8 MB budget (${pct} %; catalog limit 10 MB)${total > BUDGET ? '  ! OVER BUDGET – move something heavy into a component on demand' : ''}`);
  console.log(`Full checkout (ui/src, tests, tools … – not shipped): ${(checkout / 1048576).toFixed(2)} MB`);
  if (total > BUDGET && process.argv.includes('--strict')) process.exitCode = 2; // CI: node tools/build.js --sizes --strict
}
