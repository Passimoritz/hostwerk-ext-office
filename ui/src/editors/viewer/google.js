/*
 * Google Drive shortcut files (.gdoc, .gsheet, .gslides, .gdraw, .gform, .gsite, .gmap, .gjam, .glink) – small JSON
 * written by Google Drive for desktop / Backup and Sync: { url?, doc_id?, resource_id? ("document:<id>"), email? }.
 * Pure helpers (unit-tested).
 */

export const GOOGLE = {
  gdoc: { label: 'Google Docs', app: 'document', path: 'document/d/{id}/edit', export: 'docx', also: ['pdf', 'odt'], icon: 'FileText' },
  gsheet: { label: 'Google Sheets', app: 'spreadsheet', path: 'spreadsheets/d/{id}/edit', export: 'xlsx', also: ['pdf', 'ods'], icon: 'Sheet' },
  gslides: { label: 'Google Slides', app: 'presentation', path: 'presentation/d/{id}/edit', export: 'pptx', also: ['pdf', 'odp'], icon: 'Presentation' },
  gdraw: { label: 'Google Drawings', app: 'drawing', path: 'drawings/d/{id}/edit', export: 'pdf', also: ['svg', 'png'], icon: 'PenTool' },
  gform: { label: 'Google Forms', app: 'form', path: 'forms/d/{id}/edit', export: null, also: [], icon: 'ClipboardList' },
  gsite: { label: 'Google Sites', app: 'site', path: null, export: null, also: [], icon: 'Globe' },
  gmap: { label: 'Google My Maps', app: 'map', path: null, export: null, also: [], icon: 'Map' },
  gjam: { label: 'Google Jamboard', app: 'jam', path: null, export: 'pdf', also: [], icon: 'StickyNote' },
  glink: { label: 'Google Drive', app: 'shortcut', path: null, export: null, also: [], icon: 'Link' },
};

export function parseGoogleFile(bytes) {
  let txt = typeof bytes === 'string' ? bytes : new TextDecoder('utf-8').decode(bytes);
  txt = txt.replace(/^﻿/, '').trim();
  try { const j = JSON.parse(txt); return j && typeof j === 'object' ? j : {}; } catch { /* old .glink files are sometimes plain URLs */ }
  const m = /https?:\/\/\S+/.exec(txt);
  return m ? { url: m[0] } : {};
}

/** the Drive file id from doc_id, resource_id ("document:ID") or the url */
export function googleId(j) {
  if (!j) return '';
  if (j.doc_id) return String(j.doc_id);
  if (j.resource_id) return String(j.resource_id).replace(/^[a-z_]+:/i, '');
  const u = String(j.url || '');
  const m = /\/d\/([A-Za-z0-9_-]{10,})/.exec(u) || /[?&](?:id|mid)=([A-Za-z0-9_-]{10,})/.exec(u);
  return m ? m[1] : '';
}

/** web address of the file (only https Google hosts are returned – anything else gives '') */
export function googleUrl(ext, j) {
  const kind = GOOGLE[ext] || GOOGLE.glink;
  const id = googleId(j);
  const given = String((j && j.url) || '');
  if (/^https:\/\/([a-z0-9-]+\.)*google\.com\//i.test(given)) return given;
  if (id && kind.path) return `https://docs.google.com/${kind.path.replace('{id}', id)}`;
  if (id && ext === 'gmap') return `https://www.google.com/maps/d/edit?mid=${id}`;
  if (id && ext === 'gjam') return `https://jamboard.google.com/d/${id}/edit`;
  if (id) return `https://drive.google.com/open?id=${id}`;
  return '';
}

/** which export the backend should fetch (GET /office/google/raw?id=&as=) and which editor kind gets it */
export function googleExport(ext) {
  const k = GOOGLE[ext];
  if (!k || !k.export) return null;
  const kind = { docx: 'docs', xlsx: 'sheets', pptx: 'slides', pdf: 'pdf' }[k.export];
  return { as: k.export, kind };
}
