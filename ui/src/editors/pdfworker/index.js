/*
 * pdf.js worker as its own asset (ui/assets/pdfworker.js) – not an editor. The pdf and viewer chunks start it with
 * new Worker(<this file>, { type: 'module' }) (see ui/src/editors/_shared/pdfjs.js), so PDFs are parsed off the main
 * thread and everything works offline. If a worker cannot be started, pdf.js imports this same module on the main
 * thread ("fake worker") – that is why it also exports WorkerMessageHandler.
 */
export * from 'pdfjs-dist/build/pdf.worker.mjs';
