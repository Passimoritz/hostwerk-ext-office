/*
 * Shortcut labels in the interface language: "Ctrl+Shift+S" → "Strg+Umschalt+S" in German.
 * Code always writes the English key names; only the label is localized (never the key handling).
 *   keyLabel('Ctrl+Shift+Del')            → 'Strg+Umschalt+Entf' (de) / 'Ctrl+Shift+Del' (en)
 *   withKeys(t('Save'), 'Ctrl+S')          → 'Speichern (Strg+S)'
 *   keysInText('Press Ctrl+V or Del')      → 'Press Strg+V or Entf' – only for texts that are nothing but key hints
 */
const DE = {
  Ctrl: 'Strg', Control: 'Strg', Shift: 'Umschalt', Alt: 'Alt', AltGr: 'AltGr', Del: 'Entf', Delete: 'Entf', Ins: 'Einfg', Insert: 'Einfg',
  Enter: 'Eingabe', Return: 'Eingabe', Space: 'Leertaste', Spacebar: 'Leertaste', Backspace: 'Rücktaste', Home: 'Pos1', End: 'Ende',
  PageUp: 'Bild auf', PageDown: 'Bild ab', PgUp: 'Bild auf', PgDn: 'Bild ab', Esc: 'Esc', Escape: 'Esc', Tab: 'Tab',
  Left: '←', Right: '→', Up: '↑', Down: '↓', ArrowLeft: '←', ArrowRight: '→', ArrowUp: '↑', ArrowDown: '↓', Click: 'Klick',
};
const EN = { Control: 'Ctrl', Delete: 'Del', Escape: 'Esc', ArrowLeft: '←', ArrowRight: '→', ArrowUp: '↑', ArrowDown: '↓', Spacebar: 'Space' };

/** Interface language of Hostwerk ('de' | 'en') */
export function uiLang() {
  try { return String((typeof document !== 'undefined' && document.documentElement.lang) || 'de').toLowerCase().startsWith('en') ? 'en' : 'de'; } catch { return 'de'; }
}

const one = (k, map) => (Object.prototype.hasOwnProperty.call(map, k) ? map[k] : k);
/** "Ctrl+Shift+S" → ['Ctrl', 'Shift', 'S'], "Ctrl++" → ['Ctrl', '+'], "Ctrl+-" → ['Ctrl', '-'] */
function splitKeys(p) {
  const out = []; let cur = '';
  for (const ch of p) {
    if (ch === '+' && cur !== '') { out.push(cur); cur = ''; } else cur += ch;
  }
  if (cur !== '') out.push(cur);
  return out;
}

/** One shortcut ("Ctrl+Shift+S", "Ctrl+Alt+1", "Shift+F5", "Ctrl++", "Ctrl+-") in the given language */
export function keyLabel(keys, lang = uiLang()) {
  if (keys == null || keys === '') return '';
  const map = lang === 'en' ? EN : DE;
  const s = String(keys);
  // several alternatives: "Ctrl+Y / Ctrl+Shift+Z", "Ctrl+C, Ctrl+X"
  return s.split(/(\s*(?:\/|,| or | oder |·|–)\s*)/).map((part, i) => {
    if (i % 2) return part;
    const keysOf = splitKeys(part.trim());
    if (keysOf.length > 1 || DE[part.trim()]) return keysOf.map((k) => one(k, map)).join('+');
    return part;
  }).join('');
}

/** "Label (shortcut)" – the label is expected to be translated already */
export function withKeys(label, keys, lang = uiLang()) {
  return keys ? `${label} (${keyLabel(keys, lang)})` : String(label);
}

/** Replaces key names inside a short text (only modifier/special key words joined with "+", plus "Del"/"Esc" on their own) */
export function keysInText(text, lang = uiLang()) {
  if (lang === 'en' || text == null) return text;
  return String(text)
    .replace(/\b((?:Ctrl|Shift|Alt|AltGr)(?:\s*\+\s*[^\s+]+|\+\+|\+-)+)/g, (m) => keyLabel(m.replace(/\s*\+\s*/g, '+'), lang))
    .replace(/\b(Del|PageUp|PageDown|Backspace)\b/g, (m) => DE[m]);
}
