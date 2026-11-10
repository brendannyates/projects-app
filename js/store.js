// Storage on this PC: data in localStorage, the linked workbook's file handle in IndexedDB.
import { uid, addDays, todayIso } from './core.js';

export const DEMO = new URLSearchParams(location.search).has('demo');
const KEY = DEMO ? 'pp.demo.v1' : 'pp.data.v1';
const UNDO = DEMO ? 'pp.demo.undo' : 'pp.data.undo';
const SETTINGS = 'pp.settings.v1';

const empty = () => ({ version: 1, projects: [], tasks: [], meta: { lastWorkbookSave: '', workbookName: '' } });

export function load() {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) return normalise(JSON.parse(raw));
  } catch (e) { console.error('load failed', e); }
  return DEMO ? demoData() : empty();
}

export function save(data) {
  localStorage.setItem(KEY, JSON.stringify(data));
}

export function normalise(d) {
  const out = empty();
  out.meta = { ...out.meta, ...(d.meta || {}) };
  out.projects = (d.projects || []).map(p => ({
    id: p.id || uid(), name: String(p.name || 'Untitled project'), status: p.status || 'Active',
    start: p.start || '', due: p.due || '', notes: p.notes || ''
  }));
  out.tasks = (d.tasks || []).map(t => ({
    id: t.id || uid(), projectId: t.projectId || '', title: String(t.title || 'Untitled task'),
    status: t.status || 'To do', priority: t.priority || 'Normal',
    start: t.start || '', due: t.due || '', doneOn: t.doneOn || '', notes: t.notes || ''
  }));
  return out;
}

// Keep one snapshot so an import can be undone.
export function snapshot(data) { localStorage.setItem(UNDO, JSON.stringify(data)); }
export function hasSnapshot() { return !!localStorage.getItem(UNDO); }
export function restoreSnapshot() {
  const raw = localStorage.getItem(UNDO);
  if (!raw) return null;
  localStorage.removeItem(UNDO);
  return normalise(JSON.parse(raw));
}

export function loadSettings() {
  const def = { autoFullscreen: true, autoSave: false, timelineZoom: 'week', showDone: false };
  try { return { ...def, ...JSON.parse(localStorage.getItem(SETTINGS) || '{}') }; } catch { return def; }
}
export function saveSettings(s) { localStorage.setItem(SETTINGS, JSON.stringify(s)); }

// Ask the browser not to clear our data under storage pressure.
export async function persist() {
  if (!navigator.storage?.persist) return 'unknown';
  if (await navigator.storage.persisted()) return 'persistent';
  return (await navigator.storage.persist()) ? 'persistent' : 'best-effort';
}

// ---- workbook file handle (File System Access API, Edge/Chrome) ----
function idb() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('pp-files', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('h');
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
export async function getHandle() {
  if (DEMO) return null;
  try {
    const db = await idb();
    return await new Promise(res => {
      const q = db.transaction('h').objectStore('h').get('workbook');
      q.onsuccess = () => res(q.result || null); q.onerror = () => res(null);
    });
  } catch { return null; }
}
export async function setHandle(h) {
  if (DEMO) return;
  const db = await idb();
  await new Promise((res, rej) => {
    const tx = db.transaction('h', 'readwrite');
    h ? tx.objectStore('h').put(h, 'workbook') : tx.objectStore('h').delete('workbook');
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
}

// ---- invented demo data (?demo=1). Kept in its own storage key. ----
export function demoData() {
  const T = todayIso(), d = n => addDays(T, n);
  const P = [
    { id: 'p1', name: 'Clinic intake redesign', status: 'Active', start: d(-20), due: d(25), notes: '' },
    { id: 'p2', name: 'Quarterly vendor review', status: 'Active', start: '', due: d(2), notes: '' },
    { id: 'p3', name: 'Team handbook refresh', status: 'Active', start: '', due: '', notes: '' },
    { id: 'p4', name: 'Home office move', status: 'On hold', start: d(30), due: d(60), notes: '' },
    { id: 'p5', name: 'Phone tree cleanup', status: 'Done', start: d(-40), due: d(-12), notes: '' }
  ];
  const t = (projectId, title, status, start, due, priority = 'Normal', doneOn = '') =>
    ({ id: uid(), projectId, title, status, priority, start, due, doneOn, notes: '' });
  const tasks = [
    t('p1', 'Map current intake steps', 'Done', d(-20), d(-14), 'Normal', d(-15)),
    t('p1', 'Interview front desk leads', 'Doing', d(-10), d(-2), 'High'),
    t('p1', 'Draft new intake flow', 'Doing', d(-3), d(0), 'High'),
    t('p1', 'Review flow with ops', 'To do', d(3), d(6)),
    t('p1', 'Pilot at one clinic', 'To do', d(8), d(22)),
    t('p2', 'Pull spend by vendor', 'Done', '', d(-4), 'Normal', d(-4)),
    t('p2', 'Write review memo', 'To do', '', d(1), 'High'),
    t('p2', 'Send memo to leadership', 'To do', '', d(2)),
    t('p3', 'List outdated sections', 'To do', '', d(-1)),
    t('p3', 'Rewrite onboarding chapter', 'To do', d(4), d(12), 'Low'),
    t('p3', 'Add phone system how-to', 'To do', '', d(16), 'Low'),
    t('p4', 'Measure desk space', 'To do', '', d(32)),
    t('p5', 'Retire old menu options', 'Done', d(-40), d(-20), 'Normal', d(-21)),
    t('', 'Renew parking permit', 'To do', '', d(4)),
    t('', 'Book dentist', 'To do', '', '', 'Low')
  ];
  return { version: 1, projects: P, tasks, meta: { lastWorkbookSave: '', workbookName: '' } };
}
