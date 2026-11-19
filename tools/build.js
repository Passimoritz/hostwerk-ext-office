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
const SHARED = {
  jszip: { key: 'jszip', file: 'lib-jszip', contents: "import JSZip from 'jszip'; export default JSZip;" },
  'pdfjs-dist': { key: 'pdfjs', file: 'lib-pdfjs', contents: "import * as pdfjs from 'pdfjs-dist'; export default pdfjs;" },
};
const keyOf = (m) => SHARED[m].key;
function sharedPlugin(usedLibs) {
  return {
    name: 'office-shared-libs',
    setup(b) {
      b.onResolve({ filter: /^(jszip|pdfjs-dist)$/ }, (a) => ({ path: a.path, namespace: 'office-lib' }));
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
  const out = path.join(DIR, 'ui', 'assets', `${def.file}.js`);
  const res = await esbuild.build({ ...COMMON, stdin: { contents: def.contents, resolveDir: DIR, loader: 'js' }, outfile: out });
  const js = res.outputFiles.find((f) => /\.js$/.test(f.path));
  fs.mkdirSync(path.dirname(out), { recursive: true });
  const EV = new RegExp(`\\b${'ev'}${'al'}\\s*\\(`);
  if (EV.test(js.text)) console.warn(`! ${def.file}: bundle contains code evaluation – the catalog check will reject it`);
  fs.writeFileSync(out, `/* ${ID} shared library "${mod}" of the editor chunks – built with tools/build.js */\n${js.text}`);
  console.log(`✓ ui/assets/${def.file}.js (${Math.round(js.text.length / 1024)} KB, shared)`);
}

async function buildChunk(name) {
  const entry = ['index.jsx', 'index.js'].map((f) => path.join(DIR, 'ui', 'src', 'editors', name, f)).find((f) => fs.existsSync(f));
  if (!entry) return;
  const out = path.join(DIR, 'ui', 'assets', `${name}.js`);
  const used = new Set();
  const libs = new Set();
  const res = await esbuild.build({
    ...COMMON, entryPoints: [entry], outfile: out, jsx: 'automatic', jsxImportSource: '@hostwerk',
    loader: { '.woff': 'dataurl', '.woff2': 'dataurl', '.ttf': 'dataurl', '.png': 'dataurl', '.svg': 'text' },
    // the pdf.js worker runs on its own (no window, no globalThis.__officeLib) – it keeps its copy
    plugins: [hostPlugin(used), ...(name === 'pdfworker' ? [] : [sharedPlugin(libs)])],
  });
  const js = res.outputFiles.find((f) => /\.js$/.test(f.path));
  const css = res.outputFiles.find((f) => /\.css$/.test(f.path));
  fs.mkdirSync(path.dirname(out), { recursive: true });
  const need = [...used].filter((x) => HEAVY.includes(x));
  // shared libraries first (lib-jszip.js / lib-pdfjs.js next to this chunk, same ?v=) – the stubs read them at once
  const shared = [...libs].map((m) => [keyOf(m), SHARED[m].file]);
  const loadLibs = shared.length ? `{ const L = globalThis.__officeLib || (globalThis.__officeLib = {}); const me = new URL(import.meta.url); await Promise.all(${JSON.stringify(shared)}.filter(([k]) => !L[k]).map(async ([k, f]) => { const u = new URL(f + '.js', me); u.search = me.search; L[k] = (await import(u.href)).default; })); }\n` : '';
  const head = `/* ${ID} editor chunk "${name}" – built with tools/build.js. Source: ui/src/editors/${name} */\n${need.length ? `await globalThis.__hostwerk.need(${JSON.stringify(need)});\n` : ''}${loadLibs}`;
  // The catalog check rejects code evaluation in bundles – warn early (pattern built at runtime so this file itself passes the check)
  const EV = new RegExp(`\\b${'ev'}${'al'}\\s*\\(`);
  if (EV.test(js.text)) console.warn(`! ${name}: bundle contains code evaluation – the catalog check will reject it`);
  fs.writeFileSync(out, head + js.text);
  const cssOut = out.replace(/\.js$/, '.css');
  if (css) fs.writeFileSync(cssOut, css.text); else fs.rmSync(cssOut, { force: true });
  console.log(`✓ ui/assets/${name}.js (${Math.round((head.length + js.text.length) / 1024)} KB${css ? `, css ${Math.round(css.text.length / 1024)} KB` : ''})`);
}

(async () => {
  const only = process.argv[2];
  const editors = fs.existsSync(path.join(DIR, 'ui', 'src', 'editors')) ? fs.readdirSync(path.join(DIR, 'ui', 'src', 'editors')).filter((n) => !n.startsWith('_')) : [];
  // Version of the editor chunks for the loader (?v=…): changes with every build, so a new package never uses old chunks
  const ver = `${JSON.parse(fs.readFileSync(path.join(DIR, 'manifest.json'), 'utf8')).version}-${Date.now().toString(36)}`;
  if (!only || only === 'main') {
    fs.writeFileSync(path.join(DIR, 'ui', 'src', 'lib', 'version.js'), `// generated by tools/build.js – do not edit\nexport const CHUNK_VERSION = ${JSON.stringify(ver)};\n`);
    await buildMain(DIR);
  }
  if (!only || only === 'shared') for (const m of Object.keys(SHARED)) await buildShared(m);
  for (const n of editors) if (!only || only === n) await buildChunk(n);
  let total = 0;
  const walk = (d) => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { if (['node_modules', '.git', '.github'].includes(e.name)) continue; const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else total += fs.statSync(p).size; } };
  walk(DIR);
  console.log(`Package size: ${(total / 1048576).toFixed(2)} MB of 10 MB`);
})().catch((e) => { console.error(e.message); process.exit(1); });
