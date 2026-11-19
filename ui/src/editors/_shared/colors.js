/*
 * Human colour names for swatch buttons ("Yellow", "Dark blue" …) instead of "#ffd400".
 *   colorName('#ffd400') → 'Yellow'  (English – translate with t() / de.json)
 * Nearest named colour by a perceptual-ish distance (weighted RGB); light/dark variants get "Light"/"Dark" names.
 */
const NAMED = [
  ['Black', '#000000'], ['White', '#ffffff'], ['Grey', '#808080'], ['Light grey', '#d3d3d3'], ['Dark grey', '#404040'],
  ['Red', '#ff0000'], ['Dark red', '#c00000'], ['Light red', '#ff8080'], ['Orange', '#ff8c00'], ['Light orange', '#f7caac'],
  ['Dark orange', '#c55a11'], ['Gold', '#ffc000'], ['Yellow', '#ffff00'], ['Light yellow', '#fff2cc'], ['Dark yellow', '#bf9000'],
  ['Lime', '#92d050'], ['Green', '#00b050'], ['Light green', '#c5e0b3'], ['Dark green', '#375623'], ['Turquoise', '#00ffff'],
  ['Teal', '#008080'], ['Light blue', '#9cc3e5'], ['Blue', '#0070c0'], ['Sky blue', '#00b0f0'], ['Dark blue', '#002060'],
  ['Navy', '#1f3864'], ['Blue grey', '#44546a'], ['Purple', '#7030a0'], ['Violet', '#a46bff'], ['Light purple', '#d9c3e9'],
  ['Pink', '#ff6b9a'], ['Magenta', '#ff00ff'], ['Light pink', '#ffc7ce'], ['Brown', '#7f4f24'], ['Beige', '#f5e6c8'],
];

function rgb(hex) {
  let h = String(hex || '').trim().replace(/^#/, '');
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  if (!/^[0-9a-f]{6}/i.test(h)) return null;
  return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
}
const P = NAMED.map(([n, h]) => [n, rgb(h)]);

/* palette colours used by the editors, named exactly (the nearest match would sometimes say "Gold" for a yellow) */
const EXACT = {
  ffd400: 'Yellow', '5fd068': 'Green', '4ea8ff': 'Blue', ff6b9a: 'Pink', ff5d3a: 'Red', a46bff: 'Purple', '111111': 'Black',
  '4472c4': 'Blue', ed7d31: 'Orange', '70ad47': 'Green', '5b9bd5': 'Light blue', a5a5a5: 'Grey', ffc000: 'Gold', '00ff00': 'Bright green',
  e7e6e6: 'Light grey', '44546a': 'Blue grey',
};

/** English name of the nearest named colour ('' for an unreadable value) */
export function colorName(hex) {
  const key = String(hex || '').trim().replace(/^#/, '').toLowerCase();
  if (EXACT[key]) return EXACT[key];
  const c = rgb(hex);
  if (!c) return '';
  let best = ''; let bd = Infinity;
  for (const [n, p] of P) {
    const rm = (c[0] + p[0]) / 2;
    const dr = c[0] - p[0]; const dg = c[1] - p[1]; const db = c[2] - p[2];
    const d = (2 + rm / 256) * dr * dr + 4 * dg * dg + (2 + (255 - rm) / 256) * db * db;
    if (d < bd) { bd = d; best = n; }
  }
  return best;
}

/** All English names (for the translation check) */
export const COLOR_NAMES = [...new Set([...NAMED.map(([n]) => n), ...Object.values(EXACT)])];
