/* Heavy, rarely used parts of the slides editor load on demand from ui/assets/slides-extra.js (same folder, same ?v=). */
let p = null;
function sibling(file) {
  const me = new URL(import.meta.url);
  const u = new URL(file, me);
  u.search = me.search;
  return u.href;
}
/** { ICONS, createVectorPdf, qrSvg } */
export function extra() {
  if (!p) p = import(/* @vite-ignore */ sibling('slides-extra.js')).catch((e) => { p = null; throw e; });
  return p;
}
