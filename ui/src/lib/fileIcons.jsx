/*
 * Coloured file icons of Office â€“ one SVG set for the shell, the editors and Hostwerk's Files (slot files.icon).
 * Only React, no Hostwerk kit, so every editor chunk may import it:
 *
 *   <FileTypeIcon name="Report.docx" size={16|20|32|64} />   â€“ size â‰¤ 20: a coloured tile with a white glyph (crisp in lists);
 *                                                              â‰¥ 24: a sheet of paper with a folded corner and the tile on it
 *   <FileTypeIcon entry={{ name, dir }} size={32} ext />     â€“ `ext` shows the extension under large icons
 *   iconSvg('x.xlsx', 32) â†’ SVG markup string (for canvas, favicons, places outside React)
 *
 * Colours come from fileTypes.js (documents blue, spreadsheets green, presentations orange, PDF red â€¦). The paper uses
 * Hostwerk's surface token, so it is light in the light theme and dark in the dark theme; the tile keeps its colour.
 */
import { memo } from 'react';
import { iconInfo } from './fileTypes.js';

// glyphs on a 16Ã—16 tile, drawn in white (stroke 1.5, round)
const G = {
  lines: <path d="M4 5h8M4 8h8M4 11h5" />,
  grid: <><rect x="3.5" y="4" width="9" height="8" rx="1" /><path d="M3.5 8h9M8 4v8" /></>,
  screen: <><rect x="3" y="4" width="10" height="6.5" rx="1" /><path d="M8 10.5v2M5.75 12.75h4.5" /></>,
  sign: <path d="M3.5 11.5c1.4-3.6 2.6-6 3.6-6 1.3 0-.6 5.4.8 5.4 1 0 1.5-2 2.6-2 .7 0 .9 1 2 1" />,
  code: <path d="M6 5 3.5 8 6 11M10 5l2.5 3L10 11" />,
  mail: <><rect x="3" y="4.5" width="10" height="7" rx="1" /><path d="m3.5 5.2 4.5 3.4 4.5-3.4" /></>,
  person: <><circle cx="8" cy="6.2" r="2" /><path d="M4.5 12.3c.6-1.9 2-2.9 3.5-2.9s2.9 1 3.5 2.9" /></>,
  calendar: <><rect x="3.5" y="4.5" width="9" height="8" rx="1" /><path d="M3.5 7.5h9M6 3.3v2.2M10 3.3v2.2" /></>,
  book: <path d="M8 5.4C6.7 4.4 5 4.1 3.5 4.4v7.5c1.5-.3 3.2 0 4.5 1 1.3-1 3-1.3 4.5-1V4.4C11 4.1 9.3 4.4 8 5.4Zm0 0v7.5" />,
  flow: <><rect x="3" y="3.5" width="4.5" height="3" rx=".6" /><rect x="8.5" y="9.5" width="4.5" height="3" rx=".6" /><path d="M5.25 6.5V11H8.5" /></>,
  image: <><rect x="3" y="4" width="10" height="8" rx="1" /><path d="m3.6 11.2 3-3 2.4 2.4 1.4-1.4 2 2" /><circle cx="10.4" cy="6.4" r=".9" fill="#fff" stroke="none" /></>,
  zip: <><rect x="4.5" y="3" width="7" height="10" rx="1" /><path d="M8 3v1.4M8 5.8v1.4" /><rect x="7" y="8.6" width="2" height="2" rx=".4" /></>,
  db: <><ellipse cx="8" cy="4.8" rx="4.3" ry="1.7" /><path d="M3.7 4.8v6.4c0 .95 1.9 1.7 4.3 1.7s4.3-.75 4.3-1.7V4.8M3.7 8c0 .95 1.9 1.7 4.3 1.7s4.3-.75 4.3-1.7" /></>,
  note: <><rect x="4" y="3" width="8" height="10" rx="1" /><path d="M6.2 3v10M8.4 6h2M8.4 8.5h2" /></>,
  layout: <><rect x="3.5" y="4" width="9" height="8" rx="1" /><path d="M3.5 8H8M8 4v8" /></>,
  gantt: <path d="M4 5h5M6 8h5.5M5 11h4" strokeWidth="2" />,
  folder: <path d="M3 5.2c0-.6.4-1 1-1h2.4l1.2 1.3H12c.6 0 1 .4 1 1V11c0 .6-.4 1-1 1H4c-.6 0-1-.4-1-1V5.2Z" />,
  blank: <path d="M5.5 3.5h3.3l2.2 2.2v6.8h-5.5z" />,
};
const CLOUD = <path d="M5.2 10.4h5.4a1.9 1.9 0 0 0 .2-3.8 2.7 2.7 0 0 0-5.1-.6 2.2 2.2 0 0 0-.5 4.4Z" />;

function Glyph({ glyph }) {
  return <g fill="none" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">{G[glyph] || G.blank}</g>;
}

/**
 * props: name | entry ({ name, dir }) | ext, size (px, default 16), showExt (label under large icons), caption (adds role=img),
 * className. Never translated by Hostwerk's JSX runtime (no text children except the extension).
 */
function FileTypeIconImpl({ name, entry, ext: forceExt, size = 16, showExt = false, caption, className = '' }) {
  const info = iconInfo(entry || (forceExt ? `x.${forceExt}` : name) || '');
  const a11y = caption ? { role: 'img', 'aria-label': caption } : { 'aria-hidden': 'true' };
  const cls = `of-fi of-fi-${info.family} ${className}`;
  if (size <= 20) {
    return (
      <svg className={cls} width={size} height={size} viewBox="0 0 16 16" focusable="false" {...a11y}>
        <rect x=".5" y=".5" width="15" height="15" rx="3.5" fill={info.color} />
        <Glyph glyph={info.glyph} />
      </svg>
    );
  }
  const big = size >= 48;
  return (
    <svg className={cls} width={size} height={size} viewBox="0 0 32 32" focusable="false" {...a11y}>
      {/* paper with a folded corner (Hostwerk's surface: light or dark), outlined in the app colour */}
      <path d="M9.5 1.5h11l7 7v20a2 2 0 0 1-2 2h-16a2 2 0 0 1-2-2v-25a2 2 0 0 1 2-2Z" fill="var(--surface, #fff)" stroke={info.color} strokeOpacity=".55" />
      <path d="M20.5 1.5v5a2 2 0 0 0 2 2h5" fill={info.color} fillOpacity=".22" stroke={info.color} strokeOpacity=".55" strokeLinejoin="round" />
      {big && <path d="M17 13.5h7M17 17h7M17 20.5h5" stroke={info.color} strokeOpacity=".35" strokeWidth="1.2" strokeLinecap="round" />}
      {/* the app tile */}
      <g transform="translate(2 12)">
        <rect x="0" y="0" width="15" height="15" rx="2.8" fill={info.color} />
        <g transform="translate(-.5 -.5)"><Glyph glyph={info.glyph} /></g>
      </g>
      {info.badge === 'cloud' && <g transform="translate(18 20) scale(.75)"><circle cx="8" cy="8" r="7.2" fill={info.color} stroke="var(--surface, #fff)" strokeWidth="1.4" /><g fill="#fff">{CLOUD}</g></g>}
      {showExt && info.ext && size >= 40 && <text x="18.5" y="29" textAnchor="middle" fontSize="5.4" fontWeight="700" fill={info.color} fontFamily="system-ui, sans-serif">{info.ext.slice(0, 4).toUpperCase()}</text>}
    </svg>
  );
}
FileTypeIconImpl[Symbol.for('hostwerk.keep')] = true; // the extension label is data, never a text to translate
export const FileTypeIcon = memo(FileTypeIconImpl);
FileTypeIcon[Symbol.for('hostwerk.keep')] = true;

/* ---------------- the same as an SVG string (canvas, favicons, non-React places) ---------------- */
const GS = {
  lines: '<path d="M4 5h8M4 8h8M4 11h5"/>',
  grid: '<rect x="3.5" y="4" width="9" height="8" rx="1"/><path d="M3.5 8h9M8 4v8"/>',
  screen: '<rect x="3" y="4" width="10" height="6.5" rx="1"/><path d="M8 10.5v2M5.75 12.75h4.5"/>',
  sign: '<path d="M3.5 11.5c1.4-3.6 2.6-6 3.6-6 1.3 0-.6 5.4.8 5.4 1 0 1.5-2 2.6-2 .7 0 .9 1 2 1"/>',
};
export function iconSvg(name, size = 32) {
  const i = iconInfo(name);
  const g = `<g fill="none" stroke="#fff" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${GS[i.glyph] || GS.lines}</g>`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 16 16"><rect x=".5" y=".5" width="15" height="15" rx="3.5" fill="${i.color}"/>${g}</svg>`;
}
export { iconInfo };
