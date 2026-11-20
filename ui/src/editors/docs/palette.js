/*
 * Command palette entries of the docs editor (Alt+Q / Ctrl+Shift+P – the shell's palette): everything of the ribbon a
 * keyboard user looks for – formatting, styles, alignment, lists, insert, page setup, review, view and zoom – plus the
 * "/" commands (extras.jsx). Keys are written in English ("Ctrl+B"); the palette shows them localised ("Strg+B").
 * Pure data: run(act, editor) uses the editor's act object (Editor.jsx) and the tiptap editor.
 *   paletteCommands() → [{ id, title, group, keys?, words, run(act, editor) }]
 */
const c = (e) => e.chain().focus();

export function paletteCommands() {
  return [
    // text formatting
    { id: 'bold', title: 'Bold', group: 'Format', keys: 'Ctrl+B', words: 'fett bold strong', run: (a, e) => c(e).toggleBold().run() },
    { id: 'italic', title: 'Italic', group: 'Format', keys: 'Ctrl+I', words: 'kursiv italic', run: (a, e) => c(e).toggleItalic().run() },
    { id: 'underline', title: 'Underline', group: 'Format', keys: 'Ctrl+U', words: 'unterstrichen underline', run: (a, e) => c(e).toggleUnderline().run() },
    { id: 'strike', title: 'Strikethrough', group: 'Format', keys: 'Alt+Shift+5', words: 'durchgestrichen strike', run: (a, e) => c(e).toggleStrike().run() },
    { id: 'sub', title: 'Subscript', group: 'Format', keys: 'Ctrl+=', words: 'tiefgestellt subscript index', run: (a, e) => c(e).toggleSubscript().run() },
    { id: 'sup', title: 'Superscript', group: 'Format', keys: 'Ctrl+Shift+=', words: 'hochgestellt superscript exponent', run: (a, e) => c(e).toggleSuperscript().run() },
    { id: 'highlight', title: 'Highlight in yellow', group: 'Format', words: 'texthervorhebung markieren highlight marker gelb', run: (a, e) => (e.isActive('highlight') ? c(e).unsetHighlight().run() : c(e).setHighlight({ color: '#ffff00' }).run()) },
    { id: 'clear', title: 'Clear formatting', group: 'Format', keys: 'Ctrl+Space', words: 'formatierung löschen entfernen clear remove formatting', run: (a, e) => c(e).unsetAllMarks().run() },
    { id: 'grow', title: 'Larger text', group: 'Format', keys: 'Ctrl+Shift+>', words: 'schrift größer vergrößern font size grow', run: (a) => a.fontStep(1) },
    { id: 'shrink', title: 'Smaller text', group: 'Format', keys: 'Ctrl+Shift+<', words: 'schrift kleiner verkleinern font size shrink', run: (a) => a.fontStep(-1) },
    { id: 'upper', title: 'UPPERCASE', group: 'Format', words: 'großbuchstaben groß/kleinschreibung case upper', run: (a) => a.changeCase('upper') },
    { id: 'lower', title: 'lowercase', group: 'Format', words: 'kleinbuchstaben groß/kleinschreibung case lower', run: (a) => a.changeCase('lower') },
    { id: 'painter', title: 'Format painter', group: 'Format', words: 'format übertragen pinsel painter copy formatting', run: (a) => a.painter() },
    // styles
    { id: 'normal', title: 'Style: Normal', group: 'Styles', keys: 'Ctrl+Alt+0', words: 'standard normal absatz formatvorlage style', run: (a, e) => c(e).setParagraphStyle('normal').run() },
    { id: 'title', title: 'Style: Title', group: 'Styles', words: 'titel title formatvorlage', run: (a, e) => c(e).setParagraphStyle('title').run() },
    { id: 'subtitle', title: 'Style: Subtitle', group: 'Styles', words: 'untertitel subtitle formatvorlage', run: (a, e) => c(e).setParagraphStyle('subtitle').run() },
    { id: 'h4', title: 'Heading 4', group: 'Styles', words: 'überschrift h4 heading', run: (a, e) => c(e).setParagraphStyle('h4').run() },
    // paragraph
    { id: 'left', title: 'Align left', group: 'Paragraph', keys: 'Ctrl+L', words: 'linksbündig links ausrichten align left', run: (a, e) => c(e).setTextAlign('left').run() },
    { id: 'center', title: 'Center', group: 'Paragraph', keys: 'Ctrl+E', words: 'zentriert zentrieren mitte center', run: (a, e) => c(e).setTextAlign('center').run() },
    { id: 'right', title: 'Align right', group: 'Paragraph', keys: 'Ctrl+R', words: 'rechtsbündig rechts ausrichten align right', run: (a, e) => c(e).setTextAlign('right').run() },
    { id: 'justify', title: 'Justify', group: 'Paragraph', keys: 'Ctrl+Shift+J', words: 'blocksatz justify', run: (a, e) => c(e).setTextAlign('justify').run() },
    { id: 'indent', title: 'Increase indent', group: 'Paragraph', keys: 'Tab', words: 'einzug vergrößern einrücken indent list level', run: (a, e) => c(e).changeIndent(36).run() },
    { id: 'outdent', title: 'Decrease indent', group: 'Paragraph', keys: 'Shift+Tab', words: 'einzug verkleinern ausrücken outdent', run: (a, e) => c(e).changeIndent(-36).run() },
    { id: 'line1', title: 'Line spacing 1.0', group: 'Paragraph', words: 'zeilenabstand einfach line spacing single', run: (a, e) => c(e).setParagraphFormat({ lineHeight: 1 }).run() },
    { id: 'line15', title: 'Line spacing 1.5', group: 'Paragraph', words: 'zeilenabstand anderthalb line spacing', run: (a, e) => c(e).setParagraphFormat({ lineHeight: 1.5 }).run() },
    { id: 'line2', title: 'Line spacing 2.0', group: 'Paragraph', words: 'zeilenabstand doppelt line spacing double', run: (a, e) => c(e).setParagraphFormat({ lineHeight: 2 }).run() },
    { id: 'para', title: 'Paragraph …', group: 'Paragraph', words: 'absatz einzüge abstände paragraph spacing indents', run: (a) => a.paragraphDialog() },
    // insert
    { id: 'tablesize', title: 'Insert table …', group: 'Insert', words: 'tabelle einfügen spalten zeilen table rows columns', run: (a) => a.customTable() },
    { id: 'imagefile', title: 'Insert picture …', group: 'Insert', words: 'bild grafik foto einfügen picture image photo', run: (a) => a.imageFile() },
    { id: 'imageurl', title: 'Picture from a web address …', group: 'Insert', words: 'bild online url adresse picture web', run: (a) => a.imageUrl() },
    { id: 'link', title: 'Link …', group: 'Insert', keys: 'Ctrl+Shift+K', words: 'link hyperlink verknüpfung', run: (a) => a.link() },
    { id: 'hf', title: 'Header and footer …', group: 'Insert', words: 'kopfzeile fußzeile header footer', run: (a) => a.headerFooter() },
    // layout
    { id: 'pagesetup', title: 'Page setup …', group: 'Layout', words: 'seite einrichten seitenränder ränder format page setup margins size', run: (a) => a.marginsDialog() },
    { id: 'portrait', title: 'Portrait', group: 'Layout', words: 'hochformat ausrichtung orientation portrait', run: (a) => a.page({ orientation: 'portrait' }) },
    { id: 'landscape', title: 'Landscape', group: 'Layout', words: 'querformat ausrichtung orientation landscape', run: (a) => a.page({ orientation: 'landscape' }) },
    { id: 'a4', title: 'Page size A4', group: 'Layout', words: 'papierformat größe page size a4', run: (a) => a.pageSize('A4') },
    { id: 'letter', title: 'Page size Letter', group: 'Layout', words: 'papierformat größe page size letter', run: (a) => a.pageSize('Letter') },
    { id: 'cols1', title: 'One column', group: 'Layout', words: 'spalten eine columns one', run: (a) => a.page({ columns: 1 }) },
    { id: 'cols2', title: 'Two columns', group: 'Layout', words: 'spalten zwei columns two', run: (a) => a.page({ columns: 2 }) },
    // review
    { id: 'track', title: 'Track changes on/off', group: 'Review', keys: 'Ctrl+Shift+E', words: 'änderungen nachverfolgen überarbeiten track changes revisions', run: (a) => a.toggleTrack() },
    { id: 'acceptall', title: 'Accept all changes', group: 'Review', words: 'alle änderungen annehmen accept all changes', run: (a, e) => c(e).acceptAllChanges().run() },
    { id: 'rejectall', title: 'Reject all changes', group: 'Review', words: 'alle änderungen ablehnen reject all changes', run: (a, e) => c(e).rejectAllChanges().run() },
    { id: 'nextchange', title: 'Next change', group: 'Review', words: 'nächste änderung next change', run: (a, e) => e.commands.nextChange() },
    { id: 'count', title: 'Word count', group: 'Review', words: 'wörter zählen wortanzahl word count statistics', run: (a) => a.wordCount() },
    { id: 'spell', title: 'Spelling on/off', group: 'Review', words: 'rechtschreibung prüfung spelling spellcheck', run: (a) => a.toggleSpell() },
    { id: 'compare', title: 'Compare with …', group: 'Review', words: 'vergleichen versionen compare versions', run: (a) => a.compare() },
    // view
    { id: 'zoomin', title: 'Zoom in', group: 'View', words: 'vergrößern zoom in größer', run: (a) => a.zoomStep(1) },
    { id: 'zoomout', title: 'Zoom out', group: 'View', words: 'verkleinern zoom out kleiner', run: (a) => a.zoomStep(-1) },
    { id: 'zoom100', title: 'Zoom 100 %', group: 'View', keys: 'Ctrl+0', words: 'zoom 100 originalgröße', run: (a) => a.zoom(100) },
    { id: 'fit', title: 'Page width', group: 'View', words: 'seitenbreite zoom fit width', run: (a) => a.fitWidth() },
    { id: 'print', title: 'Print layout', group: 'View', words: 'drucklayout seitenansicht print layout pages', run: (a) => a.setLayout('print') },
    { id: 'web', title: 'Web layout', group: 'View', words: 'weblayout web layout', run: (a) => a.setLayout('web') },
    { id: 'marks', title: 'Show formatting marks', group: 'View', keys: 'Ctrl+Shift+8', words: 'formatierungszeichen absatzmarken ¶ formatting marks', run: (a) => a.toggleMarks() },
    { id: 'ruler', title: 'Ruler', group: 'View', words: 'lineal ruler', run: (a) => a.toggleRuler() },
    { id: 'nav', title: 'Navigation pane', group: 'View', words: 'navigation gliederung überschriften navigation pane outline', run: (a) => a.toggleNav() },
    { id: 'focus', title: 'Focus mode', group: 'View', words: 'fokus ablenkungsfrei focus mode', run: (a) => a.focusMode() },
    { id: 'find', title: 'Find', group: 'View', keys: 'Ctrl+F', words: 'suchen find search', run: (a) => a.find('find') },
    { id: 'replace', title: 'Replace', group: 'View', keys: 'Ctrl+H', words: 'ersetzen replace', run: (a) => a.find('replace') },
    { id: 'printdoc', title: 'Print …', group: 'View', keys: 'Ctrl+P', words: 'drucken print', run: (a) => a.print() },
  ];
}

/** the palette's list: the "/" commands first (their own groups), then the ribbon commands; ids unique */
export function allPaletteCommands(slash) {
  const seen = new Set();
  return [...slash, ...paletteCommands()].filter((x) => (seen.has(x.id) ? false : (seen.add(x.id), true)));
}
