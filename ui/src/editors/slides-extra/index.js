/*
 * Lazy sub-chunk of the slides editor (ui/assets/slides-extra.js, loaded by slides/lazy.js on first use):
 * the icon library, the vector PDF writer and the QR encoder – not needed to open a presentation.
 */
export { ICONS } from '../slides/icons-data.js';
export { createVectorPdf } from '../slides/io/pdf-vector.js';
export { qrSvg } from '../slides/qr.js';
