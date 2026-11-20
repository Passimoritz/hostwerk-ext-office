/*
 * Opening PDFs for the editor: pdf.js with CMaps/standard fonts from the "cjk-fonts" component when it is installed,
 * fonts for Unicode annotation text (bundled Latin/Greek/Cyrillic font + the component's CJK fonts), and a quick
 * byte check whether a PDF needs predefined CJK CMaps.
 */
import { pdfjsLib } from '../_shared/pdfjs.js';
import { componentState, componentBase, ensureComponent, componentBytes, isReady } from './components.js';
import { woffToSfnt, dataUrlBytes } from './ttf.js';
import LATIN from './fonts/latin-sans.woff';

let cjkCache;
/** { cMapUrl, standardFontDataUrl, fonts: [url] } of the installed cjk-fonts component, or null (never asks) */
export async function cjkAssets({ wait = 250 } = {}) {
  if (cjkCache !== undefined) return cjkCache;
  if (!pending) pending = componentState('cjk-fonts').then((st) => { cjkCache = isReady(st) ? assetsOf({ base: componentBase('cjk-fonts'), files: st.files }) : null; pending = null; return cjkCache; });
  const st = await Promise.race([pending, new Promise((r) => setTimeout(() => r('slow'), wait))]);
  if (st === 'slow') return null; // do not slow down opening – the state arrives for the next file
  return cjkCache;
}
let pending = null;
function assetsOf(comp) {
  const abs = (u) => new URL(u, location.href).href;
  const files = (Array.isArray(comp.files) ? comp.files : []).map((f) => (typeof f === 'string' ? f : f.name || f.path)).filter(Boolean);
  const dirOf = (re, def) => { const f = files.find((x) => re.test(x)); return f ? f.slice(0, f.lastIndexOf('/') + 1) : def; };
  const cmaps = dirOf(/\.bcmap$/i, 'cmaps/');
  const std = dirOf(/(Foxit|Liberation).*\.(pfb|ttf)$/i, 'standard_fonts/');
  const fonts = files.filter((x) => /\.(ttf|otf)$/i.test(x) && !/^(standard_fonts|cmaps)\//i.test(x) && !/Liberation|Foxit/i.test(x));
  return { cMapUrl: abs(comp.base + cmaps), standardFontDataUrl: abs(comp.base + std), fonts: (fonts.length ? fonts : ['NotoSansSC-Regular.ttf', 'NotoSansJP-Regular.ttf', 'NotoSansKR-Regular.ttf']).map((f) => abs(comp.base + f)) };
}
export function forgetCjk() { cjkCache = undefined; }

export function openDoc(bytes, { password, onPassword, cjk } = {}) {
  const lib = pdfjsLib();
  const task = lib.getDocument({
    data: bytes.slice(), password, isEvalSupported: false, enableXfa: false, useSystemFonts: true, fontExtraProperties: false,
    ...(cjk ? { cMapUrl: cjk.cMapUrl, cMapPacked: true, standardFontDataUrl: cjk.standardFontDataUrl } : {}),
  });
  if (onPassword) {
    task.onPassword = (update, reason) => {
      Promise.resolve(onPassword(reason === lib.PasswordResponses.INCORRECT_PASSWORD)).then((pw) => { if (pw == null) task.destroy(); else update(pw); }, () => task.destroy());
    };
  }
  return task.promise;
}

/** does the file use predefined CJK CMaps (pdf.js cannot show its text without them)? – byte scan of uncompressed parts */
export function needsCmaps(bytes) {
  const n = Math.min(bytes.length, 4e6);
  let s = '';
  for (let i = 0; i < n; i += 65536) s += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(n, i + 65536)));
  return /\/Encoding\s*\/(?!Identity-[HV])(?:Uni(?:JIS|GB|CNS|KS)[\w-]*|90ms[\w-]*|83pv[\w-]*|GBK[\w-]*|GB-EUC[\w-]*|KSC[\w-]*|ETen[\w-]*|B5pc[\w-]*|HKscs[\w-]*|Add-RKSJ[\w-]*|EUC-[HV]|RKSJ-[HV])/.test(s);
}

let latinFont = null;
export async function latinFontBytes() {
  if (!latinFont) latinFont = await woffToSfnt(dataUrlBytes(LATIN));
  return latinFont;
}

const cjkFontCache = new Map();
/** fonts of the component (downloaded once per session) – asks to install it when needed */
export async function cjkFontBytes(host, reason) {
  let assets = await cjkAssets({ wait: 3000 });
  if (!assets) {
    const comp = await ensureComponent(host, 'cjk-fonts', reason);
    if (!comp) return [];
    forgetCjk();
    assets = assetsOf(comp);
  }
  const out = [];
  for (const url of assets.fonts) {
    if (!cjkFontCache.has(url)) cjkFontCache.set(url, componentBytes(url).catch(() => null));
    const b = await cjkFontCache.get(url);
    if (b) out.push(b);
  }
  return out;
}
