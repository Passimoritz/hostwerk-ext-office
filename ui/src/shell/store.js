/*
 * Shared state of the shell: settings and converter status (cached once for the page, the editor host and the
 * file-browser slots), plus a tiny "something changed" signal so every list refreshes after an action.
 */
import { useEffect, useState } from 'react';
import { on } from '@hostwerk/api';
import { office } from '../lib/api';
import { withDefaults } from './settingsModel';

/* ---------------- Settings ---------------- */
let settings = null;
let settingsLoad = null;
const settingsSubs = new Set();
const emitSettings = () => settingsSubs.forEach((f) => f(settings));

export function loadSettings(force = false) {
  if (!force && settings) return Promise.resolve(settings);
  if (!force && settingsLoad) return settingsLoad;
  settingsLoad = office.settings().then((s) => { settings = withDefaults(s); emitSettings(); return settings; })
    .catch((e) => { settingsLoad = null; if (!settings) throw e; return settings; });
  return settingsLoad;
}
/** Settings right now (defaults until the backend answered) – for code outside React (slots, shortcuts). */
export const settingsNow = () => settings || withDefaults(null);
export async function saveSettings(patch) {
  const s = await office.saveSettings(patch);
  settings = withDefaults(s);
  emitSettings();
  return settings;
}
/** { settings, error, reload } – settings stay null until loaded (show a skeleton) */
export function useOfficeSettings() {
  const [s, setS] = useState(settings);
  const [error, setError] = useState(null);
  useEffect(() => {
    settingsSubs.add(setS);
    // settings may have arrived between the first render and this effect (loaded by a slot or shortcut) – the emit
    // then happened before we subscribed, so take them now instead of waiting for the next change
    if (settings) setS(settings);
    else loadSettings().then(setS, (e) => setError(e.message || String(e)));
    return () => settingsSubs.delete(setS);
  }, []);
  return { settings: s, error, reload: () => { setError(null); return loadSettings(true).then(setS, (e) => setError(e.message || String(e))); } };
}

/* ---------------- Status (converters, integrations) ---------------- */
let status = null;
let statusAt = 0;
let statusLoad = null;
const statusSubs = new Set();
export function loadStatus(force = false) {
  if (!force && status && Date.now() - statusAt < 60000) return Promise.resolve(status);
  if (!force && statusLoad) return statusLoad;
  statusLoad = office.status().then((s) => { status = s || {}; statusAt = Date.now(); statusLoad = null; statusSubs.forEach((f) => f(status)); return status; })
    .catch((e) => { statusLoad = null; if (status) return status; throw e; });
  return statusLoad;
}
export const statusNow = () => status;
export const hasConverterIn = (st) => !!(st && st.converters && (st.converters.libreoffice || st.converters.msoffice));
export function useOfficeStatus() {
  const [s, setS] = useState(status);
  const [error, setError] = useState(null);
  useEffect(() => {
    statusSubs.add(setS);
    loadStatus().then(setS, (e) => setError(e.message || String(e)));
    return () => statusSubs.delete(setS);
  }, []);
  return { status: s, error, hasConverter: hasConverterIn(s), reload: () => loadStatus(true).then(setS, (e) => setError(e.message || String(e))) };
}

/* ---------------- Change signal ---------------- */
const EVT = 'office:changed';
/** Tell every open list that files changed (after create, rename, move, trash …). */
export const changed = (detail = {}) => { try { window.dispatchEvent(new CustomEvent(EVT, { detail })); } catch { /* old browser */ } };
/** Run fn when files changed – here, or on the server (backend broadcasts { type: 'office' }), or when the tab comes back. */
export function useChanged(fn, deps = []) {
  useEffect(() => {
    let t = null;
    const run = () => { clearTimeout(t); t = setTimeout(() => fn(), 120); };
    window.addEventListener(EVT, run);
    const off = on('office', run);
    const vis = () => { if (!document.hidden) run(); };
    document.addEventListener('visibilitychange', vis);
    return () => { clearTimeout(t); window.removeEventListener(EVT, run); if (typeof off === 'function') off(); document.removeEventListener('visibilitychange', vis); };
  }, deps); // eslint-disable-line react-hooks/exhaustive-deps
}
