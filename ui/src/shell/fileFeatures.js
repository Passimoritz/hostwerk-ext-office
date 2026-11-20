/*
 * File-level actions of Office 1.1 – the same list in Files' right-click menu (files.actions), the Office section and
 * the editor's overflow menu / command palette: encrypt or decrypt (Hostwerk vault), clean before sharing, publish as a
 * web page, scheduled report, rules and forms (spreadsheets), older states from Hostwerk's backups.
 *   featureItems(entry, { inEditor, onCompare }) → [{ id, title, icon, run }] (order = menu order)
 */
import { Lock, Unlock, Eraser, Globe, CalendarClock, Bell, ClipboardList, History } from 'lucide-react';
import { FORMATS, extOf } from '../lib/formats';
import { askFeature } from './features';

const kindOf = (e) => (FORMATS[extOf((e && (e.name || e.path)) || '')] || {}).kind;
export const isEncrypted = (e) => !!(e && (e.encrypted || /\.hwvault$/i.test(e.name || e.path || '')));

export function featureItems(e, { onCompare = null } = {}) {
  if (!e || e.dir) return [];
  const kind = kindOf(e);
  const office = !!kind;
  const enc = isEncrypted(e);
  const loc = { root: e.root, path: e.path, name: e.name };
  return [
    office && !enc && { id: 'clean', title: 'Clean before sharing …', icon: Eraser, run: () => askFeature('clean', { entry: loc }) },
    office && !enc && kind !== 'viewer' && { id: 'publish', title: 'Publish as a web page …', icon: Globe, run: () => askFeature('publish', { entry: loc }) },
    office && !enc && kind !== 'viewer' && { id: 'report', title: 'Scheduled report …', icon: CalendarClock, run: () => askFeature('report', { entry: loc }) },
    kind === 'sheets' && !enc && { id: 'rules', title: 'Rules and notifications …', icon: Bell, run: () => askFeature('rules', { file: loc }) },
    kind === 'sheets' && !enc && { id: 'forms', title: 'Forms …', icon: ClipboardList, run: () => askFeature('forms', { file: loc }) },
    (office || enc) && { id: enc ? 'decrypt' : 'encrypt', title: enc ? 'Decrypt …' : 'Encrypt with the vault …', icon: enc ? Unlock : Lock, run: () => askFeature('vault', { entry: loc, mode: enc ? 'decrypt' : 'encrypt' }) },
    office && { id: 'backups', title: 'Older states (backups) …', icon: History, run: () => askFeature('backups', { entry: loc, onCompare }) },
  ].filter(Boolean);
}
