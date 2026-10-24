(() => {
'use strict';

/* =========================================================
   Outils
   ========================================================= */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
const clone = o => JSON.parse(JSON.stringify(o));

const APP_VERSION = '1.14.1';
const LIB = {
  hf: 'lib/hyperformula.full.min.js',
  xlsx: 'lib/xlsx-style.bundle.js',
  jszip: 'lib/jszip.min.js',
  pdfjs: 'lib/pdfjs/pdf.min.mjs', pdfWorker: 'lib/pdfjs/pdf.worker.min.mjs', pdflib: 'lib/pdf-lib.min.js',
  mammoth: 'lib/mammoth.browser.min.js',
  pdfmake: 'lib/pdfmake.min.js', pdfmakeFonts: 'lib/vfs_fonts.js'
};
const loaded = {};
function loadScript(src) {
  return loaded[src] ||= new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = src; s.onload = res;
    s.onerror = () => { delete loaded[src]; rej(new Error('Bibliothèque injoignable')); };
    document.head.appendChild(s);
  });
}

const IC = {
  tick: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
  bell: '<path d="M6 16V11a6 6 0 0 1 12 0v5l1.5 2h-15z"/><path d="M10 20.5a2 2 0 0 0 4 0"/>',
  cal: '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 10h16M9 3v4M15 3v4"/>',
  left: '<path d="M15 6l-6 6 6 6"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="M20 20l-3.6-3.6"/>',
  doc: '<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4M9 12h6M9 16h5"/>',
  grid: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16"/>',
  upload: '<path d="M12 16V4M7 9l5-5 5 5M4 20h16"/>',
  download: '<path d="M12 4v11M7 10l5 5 5-5M4 20h16"/>',
  ul: '<path d="M9 6h11M9 12h11M9 18h11"/><circle cx="4.5" cy="6" r=".8"/><circle cx="4.5" cy="12" r=".8"/><circle cx="4.5" cy="18" r=".8"/>',
  ol: '<path d="M10 6h10M10 12h10M10 18h10M4 4.5l1.5-1V8.5M3.6 13.2c.4-1 2.4-1.1 2.4.2 0 1-2.4 2-2.4 3.1H6"/>',
  alignL: '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>',
  alignC: '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>',
  alignR: '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>',
  alignJ: '<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>',
  table: '<rect x="3" y="4" width="18" height="16" rx="1.5"/><path d="M3 10h18M3 15h18M9 10v10M15 10v10"/>',
  hr: '<path d="M4 12h16M8 7h8M8 17h8" opacity=".9"/>',
  undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 010 11H11"/>',
  redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 000 11H13"/>',
  outline: '<path d="M4 6h16M8 11h12M8 16h12"/><path d="M4 11h.01M4 16h.01"/>',
  focus: '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>',
  feather: '<path d="M20 4C12 4 7 9 7 16v4"/><path d="M7 16c6 0 11-4 13-12"/><path d="M11 12h5"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>',
  moon: '<path d="M20 14.5A8 8 0 019.5 4a8 8 0 1010.5 10.5z"/>',
  sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
  sigma: '<path d="M18 5H6l6 7-6 7h12"/>',
  rowIns: '<rect x="3" y="11" width="18" height="9" rx="1.5"/><path d="M12 3v6M9 6h6"/>',
  rowDel: '<rect x="3" y="11" width="18" height="9" rx="1.5"/><path d="M9 6h6"/>',
  colIns: '<rect x="11" y="3" width="9" height="18" rx="1.5"/><path d="M3 12h6M6 9v6"/>',
  colDel: '<rect x="11" y="3" width="9" height="18" rx="1.5"/><path d="M3 12h6"/>',
  eraser: '<path d="M4 15l9-9 6 6-7 7H8z"/><path d="M12 19h8"/>',
  close: '<path d="M6 6l12 12M18 6L6 18"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  command: '<path d="M9 6a3 3 0 10-3 3h12a3 3 0 10-3-3v12a3 3 0 103-3H6a3 3 0 103 3z"/>',
  home: '<path d="M4 11l8-7 8 7v9H4z"/><path d="M10 20v-6h4v6"/>',
  pen: '<path d="M4 20h4L19 9l-4-4L4 16z"/>',
  type: '<path d="M5 6V4h14v2M12 4v16M9 20h6"/>',
  fillDown: '<path d="M12 4v12M7 11l5 5 5-5M5 20h14"/>',
  wrap: '<path d="M4 6h16M4 12h13a3 3 0 010 6h-4M15 16l-2 2 2 2M4 18h5"/>',
  bucket: '<path d="M5 11l7-7 7 7-7 7z"/><path d="M5 11h14"/><path d="M20 15s1.8 2.1 1.8 3.3a1.8 1.8 0 01-3.6 0c0-1.2 1.8-3.3 1.8-3.3z"/>',
  textColor: '<path d="M7 15l5-11 5 11M9 11h6"/><path d="M5 20h14" stroke-width="3"/>',
  borders: '<rect x="4" y="4" width="16" height="16"/><path d="M4 12h16M12 4v16" opacity=".4"/>',
  freeze: '<rect x="3" y="4" width="18" height="16" rx="1.5"/><path d="M3 9.5h18" stroke-width="2.6"/><path d="M8.5 4v16" opacity=".4"/>',
  right: '<path d="M9 6l6 6-6 6"/>',
  highlighter: '<path d="M14.5 4.5l5 5L11 18H6v-5z"/><path d="M4 21h9"/>',
  link: '<path d="M10 14a4 4 0 005.66 0l3-3a4 4 0 00-5.66-5.66l-1 1"/><path d="M14 10a4 4 0 00-5.66 0l-3 3a4 4 0 005.66 5.66l1-1"/>',
  clearFmt: '<path d="M5 6V4h12v2M11 4L8 20M15 14l5 5M20 14l-5 5"/>',
  check: '<rect x="3" y="4" width="6" height="6" rx="1"/><path d="M4.6 7l1.1 1.1L7.8 6M12 7h9M12 17h9"/><rect x="3" y="14" width="6" height="6" rx="1"/>',
  indent: '<path d="M4 5h16M11 10h9M11 14h9M4 19h16M4 9l3 3-3 3"/>',
  outdent: '<path d="M4 5h16M11 10h9M11 14h9M4 19h16M7 9l-3 3 3 3"/>',
  spacing: '<path d="M11 6h9M11 12h9M11 18h9M5 4v16M3 6l2-2 2 2M3 18l2 2 2-2"/>',
  pagebreak: '<path d="M6 3v5h12V3M6 21v-5h12v5M3 12h3M9 12h2M13 12h2M18 12h3"/>',
  toc: '<path d="M4 6h2M9 6h11M4 12h2M11 12h9M4 18h2M11 18h9"/>',
  layout: '<rect x="5" y="3" width="14" height="18" rx="1.5"/><path d="M8 6.5h8M8 17.5h8" opacity=".55"/><path d="M8 10h8v4H8z"/>',
  sortAsc: '<path d="M7 4v16M4 7l3-3 3 3M13 6h7M13 11h5M13 16h3"/>',
  sortDesc: '<path d="M7 4v16M4 17l3 3 3-3M13 6h3M13 11h5M13 16h7"/>',
  filter: '<path d="M4 5h16l-6 7.5V19l-4 1.5v-8z"/>',
  copy: '<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a1 1 0 00-1-1H5a1 1 0 00-1 1v10a1 1 0 001 1h3"/>',
  phone: '<rect x="7" y="2.5" width="10" height="19" rx="2.2"/><path d="M11 18.5h2"/>',
  monitor: '<rect x="3" y="4" width="18" height="12" rx="1.5"/><path d="M8 20h8M12 16v4"/>',
  sliders: '<path d="M4 6h9M17 6h3M4 12h3M11 12h9M4 18h11M19 18h1"/><circle cx="15" cy="6" r="2"/><circle cx="9" cy="12" r="2"/><circle cx="17" cy="18" r="2"/>',
  image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-9 9"/>',
  chart: '<path d="M4 20V11M10 20V5M16 20v-7M21 20H3"/>',
  merge: '<path d="M3 5v14M21 5v14"/><path d="M7 12h10"/><path d="M10 9l-3 3 3 3M14 9l3 3-3 3"/>',
  eyeOff: '<path d="M3 3l18 18"/><path d="M10.6 5.1A10 10 0 0112 5c5 0 9 5 9 7a12 12 0 01-2.6 3.4M6.6 6.6C4.4 8 3 10.3 3 12c0 2 4 7 9 7 1.6 0 3-.4 4.3-1.1"/><path d="M9.9 9.9a3 3 0 004.2 4.2"/>',
  eye: '<path d="M3 12c0-2 4-7 9-7s9 5 9 7-4 7-9 7-9-5-9-7z"/><circle cx="12" cy="12" r="3"/>',
  note: '<path d="M4 5h16v11H9l-5 4z"/><path d="M8 9h8M8 12h5"/>',
  list: '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M14 11l2 2 2-2"/><path d="M6 12h5"/>',
  cond: '<rect x="3" y="3" width="8" height="8" rx="1.5"/><rect x="13" y="13" width="8" height="8" rx="1.5"/><path d="M13 7h8M3 17h8"/>',
  pivot: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/><path d="M13 14l3 3 3-3"/>',
  tag: '<path d="M3 12V4h8l10 10-8 8z"/><circle cx="7.5" cy="7.5" r="1.5"/>',
  insert: '<rect x="3" y="3" width="18" height="18" rx="3"/><path d="M12 8v8M8 12h8"/>',
  database: '<ellipse cx="12" cy="5.5" rx="8" ry="2.5"/><path d="M4 5.5v13c0 1.4 3.6 2.5 8 2.5s8-1.1 8-2.5v-13"/><path d="M4 12c0 1.4 3.6 2.5 8 2.5s8-1.1 8-2.5"/>',
  pdf: '<path d="M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8z"/><path d="M14 3v5h5"/><path d="M8.5 16v-4h1.6a1.3 1.3 0 010 2.6H8.5M13 12v4h1a2 2 0 000-4z"/>',
  cursor: '<path d="M5 3l14 7-6 2-2 6z"/>',
  square: '<rect x="4" y="5" width="16" height="14" rx="1.5"/>',
  whiteout: '<rect x="3" y="7" width="18" height="10" rx="2"/><path d="M7 12h10" stroke-dasharray="2 2"/>',
  sign: '<path d="M3 17c3-1 4-9 6-9s0 8 2 8 3-4 4-4 1 3 3 3 2-1 3-1"/><path d="M3 21h18"/>',
  rotL: '<path d="M3 4v6h6"/><path d="M3.5 10A8.5 8.5 0 1 1 6 16.5"/>',
  rotR: '<path d="M21 4v6h-6"/><path d="M20.5 10A8.5 8.5 0 1 0 18 16.5"/>',
  scissors: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M8.6 7.6L20 18M8.6 16.4L20 6"/>',
  extract: '<rect x="4" y="3" width="11" height="14" rx="1.5"/><path d="M13 13l6 6M19 14v5h-5"/>',
  filePlus: '<path d="M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8z"/><path d="M14 3v5h5M12 11v6M9 14h6"/>',
  pagePlus: '<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M12 9v6M9 12h6"/>',
  move: '<path d="M12 3v18M3 12h18M8 7l4-4 4 4M8 17l4 4 4-4"/>',
  selAll: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M8 12l3 3 5-6"/>',
  comment: '<path d="M4 5h16v11h-9l-5 4v-4H4z"/><path d="M8 9h8M8 12h5"/>',
  cols2: '<path d="M4 6h7M4 10h7M4 14h7M4 18h5M13 6h7M13 10h7M13 14h7M13 18h4"/>',
  cols3: '<path d="M3 6h4.5M3 10h4.5M3 14h4.5M9.75 6h4.5M9.75 10h4.5M9.75 14h4.5M16.5 6h4.5M16.5 10h4.5M16.5 14h3"/>',
  save: '<path d="M5 4h11l3 3v13H5z"/><path d="M8 4v5h7V4M8 20v-6h8v6"/>',
  saveAs: '<path d="M11 20H5V4h11l3 3v3"/><path d="M8 4v5h7V4"/><path d="M14 20v-2.2l5.4-5.4 2.2 2.2-5.4 5.4z"/>',
  printer: '<path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>'
};
const ic = n => `<svg class="ic" viewBox="0 0 24 24" aria-hidden="true">${IC[n] || ''}</svg>`;
$$('[data-ic]').forEach(el => el.insertAdjacentHTML('afterbegin', ic(el.dataset.ic)));


let toastT;
function toast(msg, err) {
  const t = $('#toast'); t.textContent = msg; t.className = 'toast' + (err ? ' err' : ''); t.hidden = false;
  clearTimeout(toastT); toastT = setTimeout(() => { t.hidden = true; }, err ? 5200 : 2600);
}
const ls = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } }
};

/* =========================================================
   Thème
   ========================================================= */
function currentDark() {
  const t = document.documentElement.dataset.theme;
  if (t) return t === 'dark';
  return matchMedia('(prefers-color-scheme: dark)').matches;
}
function setTheme(t) {
  if (t === 'light' || t === 'dark') document.documentElement.dataset.theme = t; else delete document.documentElement.dataset.theme;
  ls.set('grim.theme', t || 'auto');
  const m = document.querySelector('meta[name="theme-color"]'); if (m) m.content = '#1B1512';
  if (typeof state !== 'undefined' && state.view === 'settings') renderSettings();
}
function toggleTheme() { setTheme(currentDark() ? 'light' : 'dark'); }
(() => { const t = ls.get('grim.theme', 'auto'); if (t === 'light' || t === 'dark') document.documentElement.dataset.theme = t; })();

/* =========================================================
   Mode PC / Téléphone
   ========================================================= */
const COARSE = matchMedia('(pointer: coarse)').matches;
let MODE = ls.get('grim.mode', 'auto');
const isPhone = () => MODE === 'phone' || (MODE === 'auto' && (COARSE || innerWidth < 720));
function applyMode() {
  const ph = isPhone();
  document.body.classList.toggle('m-phone', ph);
  document.body.classList.toggle('m-pc', !ph);
  const b = $('#modeBtn');
  b.innerHTML = ic(ph ? 'phone' : 'monitor') + `<span class="lbl-t">${ph ? 'Téléphone' : 'PC'}</span>`;
  b.title = ph ? 'Mode téléphone actif. Passer en mode PC' : 'Mode PC actif. Passer en mode téléphone';
  b.setAttribute('aria-label', b.title);
  if (typeof state !== 'undefined' && state.view === 'sheet') { if (SH.editing) commitEdit(false); requestAnimationFrame(drawSel); }
}
function setMode(m) { MODE = m; ls.set('grim.mode', m); applyMode(); if (state.view === 'settings') renderSettings(); }
$('#modeBtn').addEventListener('click', () => { setMode(isPhone() ? 'pc' : 'phone'); toast(isPhone() ? 'Mode téléphone' : 'Mode PC'); });
window.addEventListener('resize', debounce(() => { if (MODE === 'auto') applyMode(); }, 200));

/* =========================================================
   État et stockage
   ========================================================= */
const state = { docs: new Map(), loaded: false, view: 'home', cur: null, filter: 'tous', q: '' };

const Store = {
  mode: 'idb', db: null,
  req(r) { return new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
  tx(m) { return this.db.transaction('docs', m).objectStore('docs'); },
  async open() {
    this.db = await new Promise((res, rej) => {
      const r = indexedDB.open('grimoire', 2);
      r.onupgradeneeded = () => {
        const db = r.result;
        if (!db.objectStoreNames.contains('docs')) db.createObjectStore('docs', { keyPath: 'id' });
        if (!db.objectStoreNames.contains('blobs')) db.createObjectStore('blobs', { keyPath: 'id' });
      };
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
  },
  writeLocal() {
    const all = {}; state.docs.forEach((v, k) => { const { _saved, ...d } = v; all[k] = d; });
    if (!ls.set('grim.docs', all)) throw { code: 'local_full' };
  },
  async init() {
    try { await this.open(); const all = await this.req(this.tx('readonly').getAll()); state.docs = new Map(all.map(d => [d.id, d])); }
    catch { this.mode = 'local'; const all = ls.get('grim.docs', {}); state.docs = new Map(Object.entries(all).map(([k, v]) => [k, { ...v, id: k }])); }
    if (!state.docs.size && !ls.get('grim.seeded', false)) {
      ls.set('grim.seeded', true);
      const w = WELCOME(); try { await this.save(w); } catch { state.docs.set(w.id, w); }
    }
    state.loaded = true; renderHome();
    try { if (navigator.storage && navigator.storage.persist) await navigator.storage.persist(); } catch {}
  },
  async save(doc) {
    doc.updatedAt = Date.now();
    state.docs.set(doc.id, doc);
    const { _saved, ...data } = doc;
    if (this.mode === 'idb') await this.req(this.tx('readwrite').put(clone(data)));
    else this.writeLocal();
    doc._saved = true;
    if (typeof scheduleSync === 'function') scheduleSync();
  },
  async putRaw(doc) {
    state.docs.set(doc.id, doc);
    if (this.mode === 'idb') await this.req(this.tx('readwrite').put(clone(doc)));
    else this.writeLocal();
  },
  // Fichiers lourds (PDF) : rangés à part, pour que les documents restent légers à enregistrer
  async putBlob(id, bytes, rev) {
    const m = ls.get('grim.blobrev', {}); m[id] = rev || 'r' + Date.now(); ls.set('grim.blobrev', m);
    if (this.mode === 'idb') await this.req(this.db.transaction('blobs', 'readwrite').objectStore('blobs').put({ id, bytes }));
    else if (!ls.set('grim.blob.' + id, bytesToB64(bytes))) throw { code: 'local_full' };
  },
  async getBlob(id) {
    if (!id) return null;
    if (this.mode === 'idb') { const r = await this.req(this.db.transaction('blobs', 'readonly').objectStore('blobs').get(id)); return r ? new Uint8Array(r.bytes) : null; }
    const b = ls.get('grim.blob.' + id, null); return b ? b64ToBytes(b) : null;
  },
  async delBlob(id) {
    if (!id) return;
    try { if (this.mode === 'idb') await this.req(this.db.transaction('blobs', 'readwrite').objectStore('blobs').delete(id)); else localStorage.removeItem('grim.blob.' + id); } catch {}
  },
  async remove(id, fromSync) {
    const d = state.docs.get(id); if (d && d.pdfId) await this.delBlob(d.pdfId);
    state.docs.delete(id);
    if (this.mode === 'idb') await this.req(this.tx('readwrite').delete(id));
    else this.writeLocal();
    if (!fromSync) { const t = ls.get('grim.tomb', {}); t[id] = Date.now(); ls.set('grim.tomb', t); if (typeof scheduleSync === 'function') scheduleSync(); }
  }
};

let saveTimer = null, saving = false, pendingSave = false;
function setSaveState(s) {
  const el = $('#saveSt'); el.dataset.s = s;
  el.textContent = s === 'ok' ? 'Enregistré' : s === 'dirty' ? 'Modifications…' : 'Non enregistré';
}
function scheduleSave() {
  if (!state.cur) return;
  setSaveState('dirty'); if (typeof markFileDirty === 'function') markFileDirty();
  clearTimeout(saveTimer); saveTimer = setTimeout(flushSave, 900);
}
async function flushSave() {
  clearTimeout(saveTimer); saveTimer = null;
  if (!state.cur) return;
  if (saving) { pendingSave = true; return; }
  saving = true;
  try { await Store.save(state.cur); setSaveState('ok'); }
  catch (e) {
    setSaveState('err');
    const code = e && e.code;
    if (e && e.name === 'QuotaExceededError') toast('Espace de stockage plein. Supprimez quelques documents ou images pour continuer à enregistrer.', true);
    else if (code === 'local_full') toast('Stockage du navigateur plein : ce document n\'a pas pu être enregistré.', true);
    else toast('Enregistrement impossible pour le moment. Nouvel essai à la prochaine modification.', true);
  }
  saving = false;
  if (pendingSave) { pendingSave = false; flushSave(); }
}
window.addEventListener('beforeunload', () => { if (saveTimer) flushSave(); });

/* =========================================================
   Modèles
   ========================================================= */
const WELCOME = () => ({
  id: 'bienvenue', type: 'texte', title: 'Bienvenue dans Grimoire', createdAt: Date.now(), updatedAt: Date.now(),
  html: `<h1>Bienvenue dans Grimoire</h1>
<p>Votre suite bureautique : des documents comme dans Word, des classeurs comme dans Excel, sur téléphone comme sur ordinateur. Ce document est un aide-mémoire, modifiez-le ou supprimez-le librement.</p>
<h2>L'essentiel</h2>
<ul><li><strong>Ctrl K</strong> (ou le bouton ⌘) ouvre la palette : tapez « titre », « tableau », « image », « export » ou le nom d'un document.</li><li><strong>Ctrl E</strong> exporte en Word (.docx) ou en Excel (.xlsx). <strong>Ctrl P</strong> imprime ou crée un PDF.</li><li>Le bouton <strong>PC / Téléphone</strong> en haut bascule l'affichage à tout moment.</li></ul>
<h2>Le traitement de texte</h2>
<p>Police, taille, couleur, surlignage, liens, listes à plusieurs niveaux (Tab), cases à cocher, interligne. Cliquez dans un tableau pour faire apparaître ses outils : ajouter des lignes ou des colonnes, fusionner, colorier. <strong>Ctrl F</strong> recherche, <strong>Ctrl H</strong> remplace, <strong>Ctrl Entrée</strong> insère un saut de page. Le bouton Mise en page règle l'orientation, les marges, l'en-tête, le pied de page et les numéros de page.</p>
<h2>Le tableur</h2>
<p>Les formules s'écrivent comme dans Excel en français : <strong>=SOMME(D4:D20)</strong>, <strong>=SI(A1&gt;100;"oui";"non")</strong>, <strong>=SOMME.SI(C:C;"Courses";D:D)</strong>. Pendant une formule, touchez ou cliquez une cellule pour insérer sa référence. Le bouton <strong>ƒx</strong> liste les fonctions avec leur explication.</p>
<p><strong>Recopier :</strong> faites glisser le petit carré en bas à droite de la sélection pour recopier une formule ou prolonger une série (1, 2, 3… ou des dates). <strong>Ctrl D</strong> recopie vers le bas. <strong>Dates :</strong> tapez 22/10/2026 ou 22/10, et =B1-A1 donne le nombre de jours. <strong>Trier et filtrer :</strong> boutons flèches et entonnoir, dans la colonne voulue.</p>
<p><strong>Sur téléphone :</strong> touchez une cellule pour la sélectionner, puis touchez-la à nouveau pour écrire. Une rangée de touches (=, SOMME, ;, parenthèses…) apparaît sous la barre de formule.</p>
<h2>Les PDF</h2>
<p>La carte <strong>PDF</strong> ouvre un PDF, ou fusionne plusieurs PDF et photos en un seul. <strong>Lire</strong> : recherche, texte copiable, champs de formulaire à remplir. <strong>Annoter</strong> : texte, surlignage, dessin, cadre, bloc blanc pour cacher, signature, image. <strong>Pages</strong> : faire pivoter, supprimer, dupliquer, déplacer (glisser, ou appui long au doigt), extraire, diviser, ajouter des pages. <strong>Exporter</strong> enregistre le PDF modifié.</p>
<h2>Vos données</h2>
<p>Tout est enregistré automatiquement sur cet appareil, et Grimoire fonctionne hors connexion. Pour passer vos documents sur un autre appareil : <strong>Réglages → Sauvegarder tout</strong>, puis <strong>Restaurer</strong> de l'autre côté.</p>
<h2>Plume, l'assistant (facultatif)</h2>
<p>Ajoutez une clé Mistral ou OpenRouter dans les Réglages pour reformuler, corriger ou écrire une formule à partir d'une phrase. Sans clé, tout le reste fonctionne.</p>`
});
const sheetTpl = (title, rows, colW) => {
  const cells = {};
  rows.forEach(([a, v, f]) => { cells[a] = { v: String(v), f: f ? { ...f } : {} }; });
  return { type: 'feuille', title, cells, rows: 60, cols: 10, colW };
};
const E = { n: 'eur' }, EB = { n: 'eur', b: 1 }, B = { b: 1 };
const TPL = {
  lettre: () => ({
    type: 'texte', title: 'Lettre — [objet]',
    html: `<p><strong>[Prénom Nom]</strong><br>[Adresse]<br>[Code postal, Ville]<br>[Téléphone · e-mail]</p>
<p style="text-align:right">[Destinataire]<br>[Adresse]<br>[Code postal, Ville]</p>
<p style="text-align:right">[Ville], le [date]</p>
<p><strong>Objet :</strong> [objet du courrier]</p>
<p>Madame, Monsieur,</p>
<p>[Votre texte.]</p>
<p>Je vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.</p>
<p style="text-align:right">[Signature]</p>`
  }),
  budget: () => sheetTpl('Budget du mois', [
    ['A1', 'Budget du mois', B], ['A2', 'Exemple : remplacez les montants par les vôtres'],
    ['A3', 'Poste', B], ['B3', 'Prévu', { b: 1, a: 'ar' }], ['C3', 'Réel', { b: 1, a: 'ar' }], ['D3', 'Écart', { b: 1, a: 'ar' }],
    ['A4', 'Revenus', B],
    ['A5', 'Salaire'], ['B5', 2800, E], ['C5', 2800, E], ['D5', '=C5-B5', E],
    ['A6', 'Autres revenus'], ['B6', 150, E], ['C6', 90, E], ['D6', '=C6-B6', E],
    ['A7', 'Total revenus', B], ['B7', '=SOMME(B5:B6)', EB], ['C7', '=SOMME(C5:C6)', EB], ['D7', '=C7-B7', EB],
    ['A9', 'Dépenses', B],
    ['A10', 'Logement'], ['B10', 900, E], ['C10', 900, E], ['D10', '=B10-C10', E],
    ['A11', 'Courses'], ['B11', 450, E], ['C11', 512, E], ['D11', '=B11-C11', E],
    ['A12', 'Transport'], ['B12', 180, E], ['C12', 140, E], ['D12', '=B12-C12', E],
    ['A13', 'Loisirs'], ['B13', 200, E], ['C13', 260, E], ['D13', '=B13-C13', E],
    ['A14', 'Abonnements'], ['B14', 60, E], ['C14', 60, E], ['D14', '=B14-C14', E],
    ['A15', 'Total dépenses', B], ['B15', '=SOMME(B10:B14)', EB], ['C15', '=SOMME(C10:C14)', EB], ['D15', '=B15-C15', EB],
    ['A17', 'Reste à la fin du mois', B], ['B17', '=B7-B15', EB], ['C17', '=C7-C15', EB], ['D17', '=C17-B17', EB],
    ['A18', 'Taux d\'épargne'], ['B18', '=SI(B7>0;B17/B7;0)', { n: 'pct' }], ['C18', '=SI(C7>0;C17/C7;0)', { n: 'pct' }],
    ['A20', 'Écart positif = mieux que prévu.']
  ], { 0: 190, 1: 110, 2: 110, 3: 110 }),
  depenses: () => {
    const lines = [['02/10', 'Supermarché', 'Courses', 86.4], ['04/10', 'Cinéma', 'Loisirs', 24], ['05/10', 'Plein d\'essence', 'Transport', 72.15], ['09/10', 'Marché', 'Courses', 31.8], ['12/10', 'Ampoules et vis', 'Maison', 18.9], ['15/10', 'Restaurant', 'Loisirs', 58], ['18/10', 'Supermarché', 'Courses', 112.35], ['21/10', 'Péage', 'Transport', 9.6]];
    const rows = [['A1', 'Suivi des dépenses', B], ['A2', 'Exemple : ajoutez vos lignes sous le tableau, les totaux suivent'],
      ['A3', 'Date', B], ['B3', 'Libellé', B], ['C3', 'Catégorie', B], ['D3', 'Montant', { b: 1, a: 'ar' }],
      ['F3', 'Catégorie', B], ['G3', 'Total', { b: 1, a: 'ar' }]];
    lines.forEach((l, i) => { const r = i + 4; rows.push(['A' + r, l[0]], ['B' + r, l[1]], ['C' + r, l[2]], ['D' + r, String(l[3]).replace('.', ','), E]); });
    ['Courses', 'Loisirs', 'Transport', 'Maison'].forEach((c, i) => { const r = i + 4; rows.push(['F' + r, c], ['G' + r, `=SOMME.SI(C$4:C$50;F${r};D$4:D$50)`, E]); });
    rows.push(['F8', 'Total', B], ['G8', '=SOMME(G4:G7)', EB], ['F10', 'Plus grosse dépense'], ['G10', '=MAX(D4:D50)', E], ['F11', 'Nombre de dépenses'], ['G11', '=NB(D4:D50)'], ['F12', 'Dépense moyenne'], ['G12', '=SIERREUR(MOYENNE(D4:D50);0)', E]);
    return sheetTpl('Suivi des dépenses', rows, { 0: 100, 1: 170, 2: 110, 3: 100, 4: 24, 5: 160, 6: 110 });
  }
};

/* =========================================================
   Navigation
   ========================================================= */
function show(view) {
  state.view = view;
  ['home', 'writer', 'sheet', 'settings', 'pdfv'].forEach(v => { $('#' + v).hidden = v !== view; });
  $('#pageCss').textContent = '';
  const inDoc = view === 'writer' || view === 'sheet' || view === 'pdfv';
  $('#app').classList.toggle('in-doc', inDoc);
  $('#crumb').hidden = !inDoc;
  $('#barSpacer').hidden = inDoc;
  $('#exportBtn').hidden = !inDoc;
  $('#printBtn').hidden = !inDoc;
  $('#saveBtn').hidden = $('#saveAsBtn').hidden = !inDoc;
  $('#fileSt').hidden = true;
  if (inDoc) { const id = state.cur.id; FH.get(id).then(() => { if (state.cur && state.cur.id === id) updFileSt(); }); }
  $('#plumeBtn').hidden = !inDoc || view === 'pdfv';
  closePlume(); closeExport();
  document.body.classList.remove('focus', 'show-outline');
  if (inDoc) { $('#titleIn').value = state.cur.title || ''; setSaveState('ok'); }
}
async function goHome() {
  if (state.view === 'sheet') commitEdit(false);
  if (saveTimer || state.cur && !state.cur._saved) await flushSave();
  if (state.view === 'pdfv') { closePdfFind(); if (PDFV.src) { try { PDFV.src.destroy(); } catch {} } PDFV.src = PDFV.doc = null; PDFV.pages = []; PDFV.srcPages = []; PDFV.bytes = null; PDFV.thumbC.forEach(u => { if (u) URL.revokeObjectURL(u); }); PDFV.thumbC = new Map(); PDFV.hist = []; PDFV.fut = []; PDFV.gen++; RQ.list = []; TQ.hi = []; TQ.lo = []; }
  state.cur = null; show('home'); renderHome(); applyDocStyles();
}
function openDoc(id) {
  const d = state.docs.get(id); if (!d) return;
  state.cur = clone(d); state.cur.id = id; state.cur._saved = true;
  if (d.type === 'feuille') openSheet(); else if (d.type === 'pdf') openPdf(); else openWriter();
}
async function createDoc(type, base) {
  const now = Date.now();
  const doc = Object.assign(
    type === 'feuille' ? { type, title: 'Classeur sans titre', cells: {}, rows: 60, cols: 12, colW: {} }
      : type === 'pdf' ? { type, title: 'PDF', ann: [], form: {} }
      : Object.assign({ type, title: 'Document sans titre', html: '<p><br></p>', autoTitle: true }, ls.get('grim.styles', null) ? { styles: clone(ls.get('grim.styles', null)) } : {}),
    base || {}, { id: uid(), createdAt: now, updatedAt: now });
  state.cur = doc; state.cur._saved = false;
  try { if (doc.type === 'feuille') openSheet(); else if (doc.type === 'pdf') openPdf(); else openWriter(); }
  catch (e) { toast('Ouverture impossible : ' + (e.message || e), true); return; }
  try { await Store.save(doc); setSaveState('ok'); }
  catch (e) { setSaveState('err'); toast('Document créé, mais pas encore enregistré' + (e && e.code ? ' (' + e.code + ')' : '') + '. Nouvel essai à la prochaine modification.', true); }
  // On écrit tout de suite dans la page (ou la cellule A1), comme dans Word ou Excel.
  if (!base && doc.type !== 'pdf') setTimeout(() => {
    if (doc.type === 'feuille') { $('#gridWrap').focus(); return; }
    ed.focus(); const r = document.createRange(), first = ed.firstElementChild || ed;
    r.selectNodeContents(first); r.collapse(true); const sl = getSelection(); sl.removeAllRanges(); sl.addRange(r);
  }, 60);
}
const fromTemplate = k => createDoc(TPL[k]().type, TPL[k]());

$('#brand').addEventListener('click', goHome);
$('#backBtn').addEventListener('click', goHome);
$$('[data-new]').forEach(b => b.addEventListener('click', () => createDoc(b.dataset.new)));
$$('[data-tpl]').forEach(b => b.addEventListener('click', () => fromTemplate(b.dataset.tpl)));
$('#importCard').addEventListener('click', () => openFromComputer());
$('#titleIn').addEventListener('input', e => { if (!state.cur) return; state.cur.title = e.target.value; delete state.cur.autoTitle; scheduleSave(); });
$('#titleIn').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); state.view === 'writer' ? $('#ed').focus() : $('#gridWrap').focus(); } });

/* =========================================================
   Accueil
   ========================================================= */
(() => {
  const h = new Date().getHours();
  $('#greet').textContent = (h >= 18 || h < 5 ? 'Bonsoir' : 'Bonjour') + ', Cyril.';
  const d = new Intl.DateTimeFormat('fr-FR', { weekday: 'long', day: 'numeric', month: 'long' }).format(new Date());
  $('#today').textContent = d.charAt(0).toUpperCase() + d.slice(1);
})();

function docText(d) {
  if (d.type === 'pdf') return d.text || '';
  if (d.type === 'feuille') return (d.sheets || [d]).map(sh => Object.values(sh.cells || {}).map(c => c.v).filter(Boolean).join(' · ')).join(' · ');
  const div = document.createElement('div'); div.innerHTML = d.html || '';
  $$('h1,h2,h3,p,li,td,th,blockquote', div).forEach(el => el.append(' '));
  return (div.textContent || '').replace(/\s+/g, ' ').trim();
}
function relDate(t) {
  if (!t) return '';
  const diff = (Date.now() - t) / 1000, d = new Date(t), now = new Date();
  if (diff < 60) return 'à l\'instant';
  if (diff < 3600) return `il y a ${Math.floor(diff / 60)} min`;
  const hm = d.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
  if (d.toDateString() === now.toDateString()) return 'aujourd\'hui, ' + hm;
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (d.toDateString() === y.toDateString()) return 'hier, ' + hm;
  return d.toLocaleDateString('fr-FR', { day: 'numeric', month: 'short', year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric' });
}
function highlight(text, q) {
  if (!q) return esc(text);
  const n = norm(text), i = n.indexOf(q);
  if (i < 0) return esc(text);
  return esc(text.slice(0, i)) + '<mark>' + esc(text.slice(i, i + q.length)) + '</mark>' + esc(text.slice(i + q.length));
}
function snippet(text, q) {
  if (!q) return text.slice(0, 160);
  const i = norm(text).indexOf(q);
  if (i < 0) return text.slice(0, 160);
  const s = Math.max(0, i - 50);
  return (s > 0 ? '…' : '') + text.slice(s, s + 170);
}
function renderHome() {
  if (state.view !== 'home') return;
  const list = $('#docs');
  renderNotes();
  const all = [...state.docs.values()].filter(d => d.type !== 'note').sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  $('#nAll').textContent = all.length || '';
  $('#nTxt').textContent = all.filter(d => d.type === 'texte').length || '';
  $('#nSh').textContent = all.filter(d => d.type === 'feuille').length || '';
  $('#nPdf').textContent = all.filter(d => d.type === 'pdf').length || '';
  $('#storeFoot').textContent = state.loaded ? 'Enregistré automatiquement sur cet appareil. Pour copier vos documents ailleurs : Réglages, puis « Sauvegarder tout ».' : '';
  if (!state.loaded) { list.innerHTML = '<div class="skel"></div><div class="skel"></div><div class="skel"></div>'; return; }
  const q = norm(state.q.trim());
  $('#listTitle').textContent = q ? 'Résultats' : 'Récents';
  let items = all.filter(d => state.filter === 'tous' || d.type === state.filter);
  const texts = new Map(items.map(d => [d.id, docText(d)]));
  if (q) items = items.filter(d => norm(d.title).includes(q) || norm(texts.get(d.id)).includes(q));
  if (!items.length) {
    list.innerHTML = q
      ? `<p class="empty">Aucun document ne contient « <strong>${esc(state.q.trim())}</strong> ».</p>`
      : `<p class="empty"><strong>Aucun document pour l'instant.</strong><br>Créez-en un ci-dessus, partez d'un modèle, ou importez un fichier Word ou Excel.</p>`;
    return;
  }
  list.innerHTML = items.map(d => {
    const t = texts.get(d.id);
    const meta = d.type === 'pdf' ? `PDF${d.pagesN ? ' · ' + d.pagesN + ' page' + (d.pagesN > 1 ? 's' : '') : ''}${(d.ann || []).length ? ' · annoté' : ''}`
      : d.type === 'feuille'
      ? `Classeur · ${(d.sheets || [d]).reduce((n, sh) => n + Object.values(sh.cells || {}).filter(c => c.v !== '').length, 0)} cellules${(d.sheets || []).length > 1 ? ' · ' + d.sheets.length + ' feuilles' : ''}`
      : `${(t.match(/\S+/g) || []).length.toLocaleString('fr-FR')} mots`;
    return `<button class="drow" data-id="${esc(d.id)}" data-type="${d.type}">
      <span class="dico">${ic(d.type === 'feuille' ? 'grid' : d.type === 'pdf' ? 'pdf' : 'doc')}</span>
      <span style="min-width:0"><div class="dtitle">${highlight(d.title || 'Sans titre', q)}</div><div class="dsnip">${highlight(snippet(t, q), q) || '<em>Vide</em>'}</div></span>
      <span class="dmeta">${relDate(d.updatedAt)}<br>${meta}</span></button>`;
  }).join('');
}
/* =========================================================
   Notes rapides et rappels
   ========================================================= */
const NT = { showDone: false, open: null, timer: 0 };
const notesAll = () => [...state.docs.values()].filter(d => d.type === 'note');
const DAYF = new Intl.DateTimeFormat('fr-FR', { weekday: 'short', day: 'numeric', month: 'short' });
const HF2 = new Intl.DateTimeFormat('fr-FR', { hour: '2-digit', minute: '2-digit' });
function whenTxt(t) {
  const d = new Date(t), now = new Date(), day = x => new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime();
  const diff = Math.round((day(d) - day(now)) / 864e5);
  return (diff === 0 ? "aujourd'hui" : diff === 1 ? 'demain' : diff === -1 ? 'hier' : DAYF.format(d)) + ' ' + HF2.format(d);
}
function remChoices() {
  const now = new Date(), at = (dd, h, m = 0) => { const x = new Date(now); x.setDate(x.getDate() + dd); x.setHours(h, m, 0, 0); return x.getTime(); };
  const out = [['Dans 1 h', Date.now() + 36e5]];
  if (now.getHours() < 18) out.push(['Ce soir 19 h', at(0, 19)]);
  out.push(['Demain 9 h', at(1, 9)]);
  const toMon = ((8 - now.getDay()) % 7) || 7; out.push(['Lundi 9 h', at(toMon, 9)]);
  const pad = n => String(n).padStart(2, '0'), d0 = new Date(Date.now() + 36e5);
  const local = `${d0.getFullYear()}-${pad(d0.getMonth() + 1)}-${pad(d0.getDate())}T${pad(d0.getHours())}:00`;
  return `<div class="rem-ch">${out.map(([l, t]) => `<button type="button" class="chip" data-rt="${t}">${l}</button>`).join('')}
    <label class="chip rem-pick">${ic('cal')}<input type="datetime-local" data-rpick value="${local}" aria-label="Choisir la date et l'heure"></label><button type="button" class="chip" data-rok>OK</button></div>`;
}
function renderNotes() {
  const box = $('#notes'); if (!box || !state.loaded) return;
  const q = norm((state.q || '').trim());
  let all = notesAll();
  const done = all.filter(n => n.done).length;
  all = all.filter(n => !!n.done === NT.showDone && (!q || norm(n.text).includes(q)));
  const now = Date.now();
  // rappels proches en premier, puis les plus récentes
  all.sort((a, b) => (!!b.remind && !b.done) - (!!a.remind && !a.done) || (a.remind && b.remind ? a.remind - b.remind : (b.updatedAt || 0) - (a.updatedAt || 0)));
  box.innerHTML = all.map(n => {
    const late = n.remind && !n.done && n.remind <= now;
    return `<div class="note${late ? ' late' : ''}${n.done ? ' done' : ''}" data-nid="${esc(n.id)}">
      <div class="note-t" contenteditable="${n.done ? 'false' : 'true'}" spellcheck="true" lang="fr">${esc(n.text)}</div>
      ${n.remind ? `<button type="button" class="note-rem" data-na="bell">${ic('bell')}${late ? 'En retard · ' : ''}${whenTxt(n.remind)}</button>` : ''}
      <div class="note-a">
        <button type="button" data-na="done" title="${n.done ? 'Remettre en cours' : 'Fait'}" aria-label="${n.done ? 'Remettre en cours' : 'Marquer comme fait'}">${ic(n.done ? 'undo' : 'tick')}</button>
        ${n.done ? '' : `<button type="button" data-na="bell" title="Rappel" aria-label="Rappel">${ic('bell')}</button>`}
        ${n.remind && !n.done ? `<button type="button" data-na="cal" title="Ajouter au calendrier du téléphone (alarme garantie)" aria-label="Ajouter au calendrier du téléphone">${ic('cal')}</button>` : ''}
        <button type="button" data-na="doc" title="Transformer en document" aria-label="Transformer en document">${ic('doc')}</button>
        <button type="button" data-na="del" title="Supprimer" aria-label="Supprimer">${ic('trash')}</button>
      </div>
      ${NT.open === n.id ? remChoices() + (n.remind ? '<button type="button" class="lnk" data-rclear>Retirer le rappel</button>' : '') : ''}
    </div>`;
  }).join('') || (q || NT.showDone ? '' : '');
  $('#notesDone').hidden = !done && !NT.showDone;
  $('#notesDone').textContent = NT.showDone ? 'Notes en cours' : `${done} faite${done > 1 ? 's' : ''}`;
}
async function saveNote(n) { try { await Store.save(n); } catch { toast('Note non enregistrée pour le moment.', true); } scheduleReminders(); }
async function addNote(text, remind) {
  text = text.trim(); if (!text) return;
  const now = Date.now();
  const n = { id: uid(), type: 'note', title: text.split('\n')[0].slice(0, 80), text, remind: remind || null, notified: false, done: false, createdAt: now, updatedAt: now };
  await saveNote(n); renderNotes();
  if (remind) { askNotify(); toast('Note ajoutée, rappel ' + whenTxt(remind)); }
}
// Saisie rapide
const QN = { remind: null };
function qnRefresh() {
  $('#qnWhen').innerHTML = QN.open ? remChoices() : '';
  $('#qnWhen').hidden = !QN.open;
  $('#qnBell').classList.toggle('on', !!QN.remind);
  $('#qnBell').setAttribute('aria-pressed', String(!!QN.remind));
  $('#qnRemTxt').textContent = QN.remind ? 'Rappel ' + whenTxt(QN.remind) : '';
}
$('#qnForm').addEventListener('submit', async e => {
  e.preventDefault();
  const t = $('#qnText');
  await addNote(t.value, QN.remind);
  t.value = ''; t.style.height = ''; QN.remind = null; QN.open = false; qnRefresh();
});
$('#qnText').addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); $('#qnForm').requestSubmit(); } });
$('#qnText').addEventListener('input', e => { const t = e.target; t.style.height = ''; t.style.height = Math.min(160, t.scrollHeight) + 'px'; });
$('#qnBell').addEventListener('click', () => { QN.open = !QN.open; qnRefresh(); });
function pickFrom(root, e) {
  const b = e.target.closest('[data-rt]'); if (b) return +b.dataset.rt;
  if (e.target.closest('[data-rok]')) { const v = root.querySelector('[data-rpick]').value; const t = v ? new Date(v).getTime() : NaN; if (!(t > Date.now())) { toast('Choisissez une date à venir.', true); return null; } return t; }
  return null;
}
$('#qnWhen').addEventListener('click', e => { const t = pickFrom($('#qnWhen'), e); if (t) { QN.remind = t; QN.open = false; qnRefresh(); askNotify(); $('#qnText').focus(); } });
// Actions sur les notes
$('#notes').addEventListener('click', async e => {
  const card = e.target.closest('.note'); if (!card) return;
  const n = state.docs.get(card.dataset.nid); if (!n) return;
  const t = pickFrom(card, e);
  if (t) { n.remind = t; n.notified = false; NT.open = null; await saveNote(n); renderNotes(); askNotify(); toast('Rappel ' + whenTxt(t)); return; }
  if (e.target.closest('[data-rclear]')) { n.remind = null; NT.open = null; await saveNote(n); renderNotes(); return; }
  const a = e.target.closest('[data-na]'); if (!a) return;
  const k = a.dataset.na;
  if (k === 'done') { n.done = !n.done; NT.open = null; await saveNote(n); renderNotes(); if (n.done) toast('Note faite.'); }
  else if (k === 'bell') { NT.open = NT.open === n.id ? null : n.id; renderNotes(); }
  else if (k === 'cal') noteToCalendar(n);
  else if (k === 'doc') await createDoc('texte', { title: n.title || 'Note', html: n.text.split(/\n/).map(l => `<p>${esc(l) || '<br>'}</p>`).join('') });
  else if (k === 'del') {
    if (!card.classList.contains('armed')) { card.classList.add('armed'); a.title = 'Toucher encore pour supprimer'; toast('Toucher encore la corbeille pour supprimer la note.'); setTimeout(() => card.classList.remove('armed'), 3000); return; }
    try { await Store.remove(n.id); } catch {} renderNotes(); scheduleReminders(); toast('Note supprimée.');
  }
});
$('#notes').addEventListener('focusout', async e => {
  const el = e.target.closest('.note-t'); if (!el) return;
  const n = state.docs.get(el.closest('.note').dataset.nid); if (!n) return;
  const v = el.innerText.replace(/\n$/, '').trim();
  if (!v) { el.textContent = n.text; return; }
  if (v !== n.text) { n.text = v; n.title = v.split('\n')[0].slice(0, 80); await saveNote(n); }
});
$('#notesDone').addEventListener('click', () => { NT.showDone = !NT.showDone; NT.open = null; renderNotes(); });

// Calendrier du téléphone : fichier .ics avec alarme (fonctionne même appli fermée)
function noteToCalendar(n) {
  const z = t => new Date(t).toISOString().replace(/[-:]/g, '').replace(/\.\d+/, '');
  const e = s => String(s).replace(/\\/g, '\\\\').replace(/;/g, '\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
  const fold = l => { const out = []; let cur = ''; for (const ch of l) { if (new TextEncoder().encode(cur + ch).length > 73) { out.push(cur); cur = ' ' + ch; } else cur += ch; } out.push(cur); return out.join('\r\n'); };
  const ics = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Grimoire//Notes//FR', 'CALSCALE:GREGORIAN', 'BEGIN:VEVENT', `UID:${n.id}@grimoire`, `DTSTAMP:${z(Date.now())}`, `DTSTART:${z(n.remind)}`, `DTEND:${z(n.remind + 9e5)}`,
    `SUMMARY:${e(n.title || 'Rappel')}`, `DESCRIPTION:${e(n.text)}`, 'BEGIN:VALARM', 'ACTION:DISPLAY', `DESCRIPTION:${e(n.title || 'Rappel')}`, 'TRIGGER:PT0M', 'END:VALARM', 'END:VEVENT', 'END:VCALENDAR'].map(fold).join('\r\n');
  if (/iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1)) {
    // iPhone : le fichier s'ouvre dans Calendrier, qui propose « Ajouter »
    const a = document.createElement('a'); a.href = 'data:text/calendar;charset=utf-8,' + encodeURIComponent(ics); a.target = '_blank'; a.rel = 'noopener';
    document.body.appendChild(a); a.click(); a.remove();
  } else saveFile('Rappel ' + safeName(n.title).slice(0, 40) + '.ics', ics);
}

// Rappels : notification quand l'heure arrive (appli ouverte ou en arrière-plan récent), pastille sur l'icône
async function askNotify() {
  if (!('Notification' in window) || Notification.permission !== 'default') return;
  try { await Notification.requestPermission(); } catch {}
}
async function fireReminder(n) {
  const body = n.text.length > 180 ? n.text.slice(0, 177) + '…' : n.text;
  let shown = false;
  if ('Notification' in window && Notification.permission === 'granted') {
    try { const reg = await navigator.serviceWorker.ready; await reg.showNotification('Rappel Grimoire', { body, tag: 'note-' + n.id, icon: 'icons/icon-192.png', badge: 'icons/icon-192.png', data: { note: n.id }, requireInteraction: true }); shown = true; }
    catch { try { new Notification('Rappel Grimoire', { body, tag: 'note-' + n.id, icon: 'icons/icon-192.png' }); shown = true; } catch {} }
  }
  if (!shown || document.visibilityState === 'visible') toast('🔔 ' + body);
}
async function checkReminders() {
  if (!state.loaded) return;
  const now = Date.now(); let changed = false;
  for (const n of notesAll()) if (n.remind && !n.done && n.remind <= now && !n.notified) { n.notified = true; changed = true; await fireReminder(n); try { await Store.save(n); } catch {} }
  const due = notesAll().filter(n => n.remind && !n.done && n.remind <= now).length;
  try { if (navigator.setAppBadge) due ? navigator.setAppBadge(due) : navigator.clearAppBadge(); } catch {}
  if (changed || due) renderNotes();
}
function scheduleReminders() {
  clearTimeout(NT.timer);
  const next = notesAll().filter(n => n.remind && !n.done && !n.notified).map(n => n.remind).sort((a, b) => a - b)[0];
  checkReminders();
  if (next) NT.timer = setTimeout(scheduleReminders, Math.max(500, Math.min(next - Date.now() + 200, 6e4)));
}
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') scheduleReminders(); });
// clic sur une notification : afficher la note
async function focusNote(id) {
  if (state.view !== 'home') await goHome();
  NT.showDone = false; renderNotes();
  const el = $(`.note[data-nid="${CSS.escape(id)}"]`); if (el) { el.scrollIntoView({ block: 'center', behavior: 'smooth' }); el.classList.add('flash'); setTimeout(() => el.classList.remove('flash'), 1600); }
}
/* =========================================================
   Sauvegarde en ligne : dépôt GitHub privé, synchronisé entre appareils
   ========================================================= */
const SY = { cfg: ls.get('grim.sync', null), busy: false, again: false, timer: 0, last: ls.get('grim.sync.last', 0), err: '' };
const utf8B64 = s => bytesToB64(new TextEncoder().encode(s));
const b64Utf8 = b => new TextDecoder().decode(b64ToBytes(b.replace(/\s/g, '')));
async function gh(path, opt = {}, cfg = SY.cfg) {
  const h = { Authorization: 'Bearer ' + cfg.token, Accept: opt.raw ? 'application/vnd.github.raw+json' : 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
  if (opt.body) h['Content-Type'] = 'application/json';
  return fetch('https://api.github.com' + (path.startsWith('/user') ? '' : '/repos/' + cfg.repo) + path, { method: opt.method || 'GET', headers: h, body: opt.body ? JSON.stringify(opt.body) : undefined, cache: 'no-store', keepalive: !!opt.keepalive });
}
const ghErr = async r => { let m = ''; try { m = (await r.json()).message; } catch {} return Object.assign(new Error(m || 'HTTP ' + r.status), { status: r.status }); };
async function ghShaOf(path) { const r = await gh('/contents/' + path); return r.ok ? (await r.json()).sha : undefined; }
async function ghPut(path, b64, sha) {
  for (let k = 0; k < 2; k++) {
    const r = await gh('/contents/' + path, { method: 'PUT', body: { message: 'Grimoire : ' + path.split('/').pop(), content: b64, ...(sha ? { sha } : {}) } });
    if (r.ok) return (await r.json()).content.sha;
    if (k === 0 && (r.status === 409 || r.status === 422)) { sha = await ghShaOf(path); continue; }
    throw await ghErr(r);
  }
}
async function ghDel(path, sha) {
  sha ||= await ghShaOf(path); if (!sha) return;
  const r = await gh('/contents/' + path, { method: 'DELETE', body: { message: 'Grimoire : suppression', sha } });
  if (!r.ok && r.status !== 404 && r.status !== 409 && r.status !== 422) throw await ghErr(r);
}
const blobRevs = () => ls.get('grim.blobrev', {});
function blobRevOf(id) { const m = blobRevs(); if (!m[id]) { m[id] = 'r' + Date.now(); ls.set('grim.blobrev', m); } return m[id]; }
async function syncNow() {
  if (!SY.cfg || !navigator.onLine) return;
  if (SY.busy) { SY.again = true; return; }
  SY.busy = true; SY.again = false; syncUI('run');
  let down = 0;
  try {
    if (state.cur && !state.cur._saved) { try { await flushSave(); } catch {} }
    let idx = { v: 1, docs: {} }, idxSha;
    const r = await gh('/contents/grimoire/index.json');
    if (r.ok) { const j = await r.json(); idxSha = j.sha; idx = JSON.parse(b64Utf8(j.content)); idx.docs ||= {}; }
    else if (r.status !== 404) throw await ghErr(r);
    const tomb = ls.get('grim.tomb', {});
    let changed = false;
    const upload = async (L, R) => {
      const { _saved, ...data } = L, e = { u: L.updatedAt || 0, t: L.type };
      if (L.type === 'pdf' && L.pdfId) {
        const rev = blobRevOf(L.pdfId);
        if (R && R.pid === L.pdfId && R.b === rev) Object.assign(e, { pid: R.pid, b: R.b, ps: R.ps });
        else { const bytes = await Store.getBlob(L.pdfId); if (bytes) Object.assign(e, { pid: L.pdfId, b: rev, ps: await ghPut('grimoire/pdf/' + L.pdfId + '.pdf', bytesToB64(bytes), R && R.pid === L.pdfId ? R.ps : undefined) }); }
      }
      e.s = await ghPut('grimoire/docs/' + L.id + '.json', utf8B64(JSON.stringify(data)), R && R.s);
      idx.docs[L.id] = e; changed = true;
    };
    const download = async (id, R) => {
      const rd = await gh('/contents/grimoire/docs/' + id + '.json', { raw: true }); if (!rd.ok) throw await ghErr(rd);
      const d = JSON.parse(await rd.text()); d.id = id; d.updatedAt = R.u;
      if (d.type === 'pdf' && d.pdfId && (blobRevs()[d.pdfId] !== R.b || !(await Store.getBlob(d.pdfId)))) {
        const rb = await gh('/contents/grimoire/pdf/' + d.pdfId + '.pdf', { raw: true }); if (!rb.ok) throw await ghErr(rb);
        await Store.putBlob(d.pdfId, new Uint8Array(await rb.arrayBuffer()), R.b);
      }
      await Store.putRaw(d); down++;
    };
    const ids = new Set([...state.docs.keys(), ...Object.keys(idx.docs), ...Object.keys(tomb)]);
    for (const id of ids) {
      const L = state.docs.get(id), R = idx.docs[id], T = tomb[id];
      if (state.cur && state.cur.id === id && R && L && R.u > (L.updatedAt || 0)) continue; // document ouvert : à la prochaine fois
      if (R && R.d) {
        if (L && (L.updatedAt || 0) > R.d) await upload(L, null);
        else if (L) { await Store.remove(id, true); down++; }
        continue;
      }
      if (!L) {
        if (T) { if (R && T >= R.u) { await ghDel('grimoire/docs/' + id + '.json', R.s); if (R.pid) await ghDel('grimoire/pdf/' + R.pid + '.pdf', R.ps); idx.docs[id] = { d: T }; changed = true; } else if (R) await download(id, R); }
        else if (R) await download(id, R);
        continue;
      }
      if (!R || (L.updatedAt || 0) > R.u || (L.type === 'pdf' && L.pdfId && R.b && blobRevOf(L.pdfId) !== R.b && (L.updatedAt || 0) >= R.u)) await upload(L, R);
      else if (R.u > (L.updatedAt || 0)) await download(id, R);
    }
    if (changed) {
      idx.at = Date.now();
      try { await ghPut('grimoire/index.json', utf8B64(JSON.stringify(idx)), idxSha); }
      catch (e) { if (e.status === 409 || e.status === 422) SY.again = true; else throw e; }
    }
    // suppressions appliquées : les marques locales ne servent plus
    Object.keys(tomb).forEach(id => { if (!idx.docs[id] || idx.docs[id].d) delete tomb[id]; });
    ls.set('grim.tomb', tomb);
    SY.last = Date.now(); ls.set('grim.sync.last', SY.last); SY.err = '';
    syncUI('ok');
  } catch (e) {
    console.warn('sync', e);
    SY.err = e.status === 401 ? 'Jeton refusé ou expiré : collez un nouveau jeton dans Réglages.' : e.status === 404 ? 'Dépôt introuvable ou inaccessible avec ce jeton.' : e.status === 403 ? 'Accès refusé : le jeton doit avoir « Contents : Read and write » sur ce dépôt.' : 'Sauvegarde en ligne impossible pour le moment' + (navigator.onLine ? '' : ' (hors connexion)') + '.';
    syncUI('err');
  } finally {
    SY.busy = false;
    if (down) { if (state.view === 'home') renderHome(); scheduleReminders(); }
    if (SY.again) setTimeout(syncNow, 1500);
  }
}
function scheduleSync(ms = 4000) { if (!SY.cfg) return; clearTimeout(SY.timer); SY.timer = setTimeout(syncNow, ms); }
function syncUI(s) {
  const st = $('#syncSt'), when = SY.last ? HF2.format(new Date(SY.last)) : '';
  const txt = !SY.cfg ? '' : s === 'run' ? 'Sauvegarde en ligne…' : s === 'err' ? SY.err : when ? 'Sauvegardé en ligne à ' + when + '.' : '';
  if (st) { st.textContent = txt; st.className = 'set-res' + (s === 'err' ? ' err' : s === 'ok' ? ' ok' : ''); }
  const f = $('#syncFoot'); if (f) { f.textContent = SY.cfg ? txt : ''; f.classList.toggle('err', s === 'err'); }
  const w = $('#syncWarn'); if (w) w.hidden = !!SY.cfg || ss.get('grim.syncLater');
  $('#syncOn') && ($('#syncOn').hidden = !!SY.cfg); $('#syncNow') && ($('#syncNow').hidden = !SY.cfg); $('#syncOff') && ($('#syncOff').hidden = !SY.cfg);
}
const ss = { get: k => { try { return sessionStorage.getItem(k); } catch { return null; } }, set: (k, v) => { try { sessionStorage.setItem(k, v); } catch {} } };
async function syncEnable() {
  const repo = $('#syncRepo').value.trim().replace(/^https?:\/\/github\.com\//, '').replace(/\.git$/, '').replace(/\/$/, ''), token = $('#syncTok').value.trim();
  const res = $('#syncSt');
  if (!/^[\w.-]+\/[\w.-]+$/.test(repo) || !token) { res.textContent = 'Indiquez le dépôt (nom/grimoire-data) et le jeton.'; res.className = 'set-res err'; return; }
  const cfg = { repo, token };
  res.textContent = 'Vérification…'; res.className = 'set-res';
  try {
    let r = await gh('', {}, cfg);
    if (r.status === 404) { // dépôt absent : création si le jeton le permet
      const c = await gh('/user/repos', { method: 'POST', body: { name: repo.split('/')[1], private: true, description: 'Données Grimoire (privé)' } }, cfg);
      if (!c.ok) throw Object.assign(new Error(), { status: 404 });
      r = await gh('', {}, cfg);
    }
    if (!r.ok) throw await ghErr(r);
    const j = await r.json();
    if (!j.private) { res.textContent = 'Ce dépôt est public : vos documents seraient visibles de tous. Utilisez un dépôt privé.'; res.className = 'set-res err'; return; }
    if (j.permissions && !j.permissions.push) throw Object.assign(new Error(), { status: 403 });
    SY.cfg = cfg; ls.set('grim.sync', cfg); $('#syncTok').value = '';
    await syncNow();
    if (!SY.err) toast('Sauvegarde en ligne activée.');
  } catch (e) {
    res.textContent = e.status === 401 ? 'Jeton refusé : vérifiez qu\'il est bien copié en entier.' : e.status === 404 ? 'Dépôt introuvable : créez le dépôt privé (étape 1) et donnez-y accès au jeton.' : e.status === 403 ? 'Le jeton n\'a pas le droit d\'écrire : « Contents : Read and write » est nécessaire.' : 'Connexion à GitHub impossible.';
    res.className = 'set-res err';
  }
}
$('#syncOn').addEventListener('click', syncEnable);
$('#syncNow').addEventListener('click', () => syncNow());
$('#syncOff').addEventListener('click', () => { SY.cfg = null; ls.set('grim.sync', null); try { localStorage.removeItem('grim.sync'); } catch {} syncUI(); toast('Sauvegarde en ligne désactivée sur cet appareil (les données en ligne sont conservées).'); });
$('#syncLater').addEventListener('click', () => { ss.set('grim.syncLater', '1'); syncUI(); });
$('#syncGo').addEventListener('click', () => openSettings().then(() => setTimeout(() => $('#syncCard').scrollIntoView({ block: 'start', behavior: 'smooth' }), 80)));
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') { if (SY.cfg && !SY.busy) syncNow(); } else scheduleSync(500); });
addEventListener('online', () => scheduleSync(500));
setInterval(() => { if (document.visibilityState === 'visible') scheduleSync(0); }, 120000);

if (navigator.serviceWorker) navigator.serviceWorker.addEventListener('message', e => { if (e.data && e.data.note) focusNote(e.data.note); });

$('#docs').addEventListener('click', e => { const b = e.target.closest('.drow'); if (b) openDoc(b.dataset.id); });
$('#q').addEventListener('input', debounce(e => { state.q = e.target.value; renderHome(); }, 120));
$('#chips').addEventListener('click', e => {
  const b = e.target.closest('.chip'); if (!b) return;
  state.filter = b.dataset.f; $$('.chip').forEach(c => c.setAttribute('aria-pressed', String(c === b))); renderHome();
});

/* =========================================================
   Nettoyage HTML (collage, import)
   ========================================================= */
const KEEP = { P: 1, H1: 1, H2: 1, H3: 1, STRONG: 1, EM: 1, U: 1, S: 1, SUP: 1, SUB: 1, UL: 1, OL: 1, LI: 1, TABLE: 1, THEAD: 1, TBODY: 1, TR: 1, TD: 1, TH: 1, COLGROUP: 1, COL: 1, BR: 1, BLOCKQUOTE: 1, HR: 1, A: 1 };
const MAP = { B: 'strong', I: 'em', H4: 'h3', H5: 'h3', H6: 'h3', DIV: 'p', TFOOT: 'tbody', STRIKE: 's', DEL: 's', INS: 'u' };
const DROP = new Set(['SCRIPT', 'STYLE', 'META', 'LINK', 'TITLE', 'IFRAME', 'OBJECT', 'EMBED', 'SVG', 'VIDEO', 'AUDIO', 'CANVAS', 'FORM', 'INPUT', 'BUTTON', 'SELECT', 'TEXTAREA', 'NOSCRIPT', 'TEMPLATE', 'HEAD']);
const FONT_OK = ['Source Serif 4', 'Instrument Sans', 'Arial', 'Calibri', 'Georgia', 'Times New Roman', 'Verdana', 'Courier New'];
const FONT_STACK = { 'Source Serif 4': "'Source Serif 4', Georgia, serif", 'Instrument Sans': "'Instrument Sans', Arial, sans-serif", 'Arial': 'Arial, Helvetica, sans-serif', 'Calibri': 'Calibri, Carlito, Arial, sans-serif', 'Georgia': 'Georgia, serif', 'Times New Roman': "'Times New Roman', Times, serif", 'Verdana': 'Verdana, Geneva, sans-serif', 'Courier New': "'Courier New', Courier, monospace" };
const DARK_TEXT = /^#(000000|111111|1E1814|1B1512|222222|202124|333333)$/;
let droppedImages = 0;
function cssColor(v) {
  if (!v) return '';
  v = String(v).trim();
  let m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(v);
  if (m) { let h = m[1]; if (h.length === 3) h = h.split('').map(c => c + c).join(''); return '#' + h.toUpperCase(); }
  m = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+)(%?))?\s*\)$/i.exec(v);
  if (m) { if (m[4] !== undefined && parseFloat(m[4]) === 0) return ''; return '#' + [m[1], m[2], m[3]].map(x => Math.min(255, +x).toString(16).padStart(2, '0')).join('').toUpperCase(); }
  return '';
}
function cssFont(v) {
  if (!v) return '';
  const first = String(v).split(',')[0].replace(/["']/g, '').trim().toLowerCase();
  return FONT_OK.find(x => x.toLowerCase() === first) || '';
}
function cssSizePt(v) {
  const m = /^([\d.]+)(px|pt)$/.exec(String(v || '').trim()); if (!m) return 0;
  const pt = Math.round((m[2] === 'px' ? +m[1] * 0.75 : +m[1]) * 2) / 2;
  return pt >= 6 && pt <= 72 ? pt : 0;
}
const cssLen = v => {
  const m = /^(-?[\d.]+)(px|pt|em|rem|cm|mm|in)?$/.exec(String(v || '').trim()); if (!m) return NaN;
  const n = +m[1]; return Math.round(n * ({ px: 1, pt: 4 / 3, em: 16, rem: 16, cm: 37.8, mm: 3.78, in: 96 }[m[2] || 'px']));
};
// Les cases ☐ / ☒ exportées vers Word redeviennent une liste de cases à cocher à l'import.
function restoreChecklists(html) {
  const box = document.createElement('div'); box.innerHTML = html;
  let ul = null;
  [...box.children].forEach(el => {
    const m = el.tagName === 'P' && /^\s*([☐☒☑])\s?/.exec(el.textContent);
    if (!m) { ul = null; return; }
    if (!ul) { ul = document.createElement('ul'); ul.setAttribute('data-check', '1'); el.before(ul); }
    const li = document.createElement('li');
    if (m[1] !== '☐') li.setAttribute('data-done', '1');
    li.innerHTML = el.innerHTML.replace(/^\s*[☐☒☑]\s?/, '');
    ul.appendChild(li); el.remove();
  });
  return box.innerHTML;
}
function sanitize(html) {
  const doc = new DOMParser().parseFromString('<body>' + (html || ''), 'text/html');
  clean(doc.body, doc);
  return doc.body.innerHTML.trim() || '<p><br></p>';
}
function clean(node, doc) {
  for (const ch of [...node.childNodes]) {
    if (ch.nodeType === 3) continue;
    if (ch.nodeType !== 1) { ch.remove(); continue; }
    const tag = ch.tagName;
    if (DROP.has(tag)) { ch.remove(); continue; }
    if (tag === 'IMG') {
      const src = ch.getAttribute('src') || '';
      if (!/^data:image\/(png|jpeg|gif);base64,[A-Za-z0-9+/=\s]+$/.test(src)) { droppedImages++; ch.remove(); continue; }
      const w = ch.style && ch.style.width, wrap = ch.getAttribute('data-wrap');
      for (const at of [...ch.attributes]) ch.removeAttribute(at.name);
      ch.setAttribute('src', src.replace(/\s+/g, '')); ch.setAttribute('alt', '');
      if (w && /^\d{1,3}(\.\d+)?%$/.test(w)) ch.style.width = w;
      if (wrap === 'left' || wrap === 'right') ch.setAttribute('data-wrap', wrap);
      continue;
    }
    if (tag === 'DIV' && /^[23]$/.test(ch.getAttribute('data-cols') || '')) {
      const d = doc.createElement('div'); d.setAttribute('data-cols', ch.getAttribute('data-cols'));
      d.append(...ch.childNodes); clean(d, doc);
      [...d.childNodes].forEach(n => { if (n.nodeType === 3 && n.nodeValue.trim()) { const pp = doc.createElement('p'); n.replaceWith(pp); pp.appendChild(n); } });
      if (!d.firstChild) d.innerHTML = '<p><br></p>';
      ch.replaceWith(d); continue;
    }
    if (tag === 'DIV' && ch.hasAttribute('data-toc')) {
      const d = doc.createElement('div'); d.setAttribute('data-toc', '1'); d.setAttribute('contenteditable', 'false'); ch.replaceWith(d); continue;
    }
    const st = ch.style || {};
    const deco = (st.textDecorationLine || st.textDecoration || '');
    const S = {
      align: st.textAlign, lh: st.lineHeight, ml: st.marginLeft, mb: st.marginBottom,
      color: cssColor(st.color || (tag === 'FONT' ? ch.getAttribute('color') : '')),
      bg: cssColor(st.backgroundColor),
      ff: cssFont(st.fontFamily || (tag === 'FONT' ? ch.getAttribute('face') : '')),
      fs: tag === 'FONT' && ch.getAttribute('size') ? ({ 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 })[ch.getAttribute('size')] || 0 : cssSizePt(st.fontSize),
      bold: /^(bold|[6-9]00)$/.test(st.fontWeight), ital: st.fontStyle === 'italic',
      under: /underline/.test(deco), strike: /line-through/.test(deco), va: st.verticalAlign
    };
    if (tag === 'B' && st.fontWeight === 'normal') { clean(ch, doc); ch.replaceWith(...ch.childNodes); continue; }
    if ((tag === 'INS' || tag === 'DEL') && ch.hasAttribute('data-a')) {
      const el = doc.createElement(tag.toLowerCase());
      el.setAttribute('data-a', String(ch.getAttribute('data-a')).slice(0, 60));
      const t = ch.getAttribute('data-t'); if (/^\d{1,15}$/.test(t || '')) el.setAttribute('data-t', t);
      el.append(...ch.childNodes); clean(el, doc); ch.replaceWith(el); continue;
    }
    if (tag === 'SPAN' && /^[\w-]{1,48}$/.test(ch.getAttribute('data-c') || '')) {
      // ancre de commentaire, éventuellement avec une mise en forme : la mise en forme passe à l'intérieur
      const id = ch.getAttribute('data-c'), sp = doc.createElement('span'); sp.setAttribute('data-c', id);
      ch.removeAttribute('data-c'); const inner = doc.createElement('span'); inner.setAttribute('style', ch.getAttribute('style') || '');
      inner.append(...ch.childNodes); sp.appendChild(inner); clean(sp, doc); ch.replaceWith(sp); continue;
    }
    // Éléments de mise en forme en ligne : on ne garde que les styles connus
    if (tag === 'SPAN' || tag === 'FONT' || tag === 'MARK' || !(KEEP[tag] || MAP[tag])) {
      clean(ch, doc);
      const wraps = [];
      const col = S.color && !DARK_TEXT.test(S.color) ? S.color : '';
      const bg = S.bg && S.bg !== '#FFFFFF' ? S.bg : (tag === 'MARK' ? '#FFF1A8' : '');
      if (col || bg || S.ff || S.fs) {
        const sp = doc.createElement('span');
        if (col) sp.style.color = col;
        if (bg) sp.style.backgroundColor = bg;
        if (S.ff) sp.style.fontFamily = FONT_STACK[S.ff];
        if (S.fs) sp.style.fontSize = S.fs + 'pt';
        wraps.push(sp);
      }
      if (S.bold) wraps.push(doc.createElement('strong'));
      if (S.ital) wraps.push(doc.createElement('em'));
      if (S.under) wraps.push(doc.createElement('u'));
      if (S.strike) wraps.push(doc.createElement('s'));
      if (S.va === 'super') wraps.push(doc.createElement('sup')); else if (S.va === 'sub') wraps.push(doc.createElement('sub'));
      const inner = [...ch.childNodes];
      if (!wraps.length) { ch.replaceWith(...inner); continue; }
      for (let i = 0; i < wraps.length - 1; i++) wraps[i].appendChild(wraps[i + 1]);
      inner.forEach(n => wraps[wraps.length - 1].appendChild(n));
      ch.replaceWith(wraps[0]);
      continue;
    }
    const href = tag === 'A' ? ch.getAttribute('href') : null;
    const keep = {};
    if (tag === 'UL' && ch.hasAttribute('data-check')) keep['data-check'] = '1';
    if (tag === 'LI' && ch.getAttribute('data-done') === '1') keep['data-done'] = '1';
    if (tag === 'HR' && ch.hasAttribute('data-pb')) keep['data-pb'] = '1';
    if (tag === 'TD' || tag === 'TH') { const cs = parseInt(ch.getAttribute('colspan'), 10); if (cs > 1 && cs <= 30) keep.colspan = String(cs); }
    const colW = tag === 'COL' ? (st.width || ch.getAttribute('width') || '') : '';
    let el = ch;
    if (MAP[tag]) { el = doc.createElement(MAP[tag]); while (ch.firstChild) el.appendChild(ch.firstChild); ch.replaceWith(el); }
    for (const a of [...el.attributes]) el.removeAttribute(a.name);
    Object.entries(keep).forEach(([k, v]) => el.setAttribute(k, v));
    const T = el.tagName;
    if (/^(P|H1|H2|H3|LI|TD|TH|BLOCKQUOTE)$/.test(T) && S.align && /^(left|center|right|justify)$/.test(S.align)) el.style.textAlign = S.align;
    if (/^(P|H1|H2|H3|LI|BLOCKQUOTE)$/.test(T)) {
      if (S.lh && /^[\d.]+$/.test(S.lh)) { const lh = parseFloat(S.lh); if (lh >= 1 && lh <= 3) el.style.lineHeight = String(lh); }
      const ml = cssLen(S.ml); if (ml > 0 && ml <= 480 && T !== 'LI') el.style.marginLeft = ml + 'px';
      if (S.mb) { const mb = cssLen(S.mb); if (mb === 0) el.style.marginBottom = '0px'; else if (mb >= 18) el.style.marginBottom = '1.5em'; }
    }
    if ((T === 'TD' || T === 'TH') && S.bg && S.bg !== '#FFFFFF') el.style.backgroundColor = S.bg;
    if (T === 'COL') { const w = parseFloat(colW); if (/%$/.test(colW) && w > 0 && w <= 100) el.style.width = w + '%'; }
    if (T === 'A') {
      if (href && /^(https?:|mailto:)/i.test(href)) { el.setAttribute('href', href); el.setAttribute('target', '_blank'); el.setAttribute('rel', 'noopener'); }
      else { clean(el, doc); el.replaceWith(...el.childNodes); continue; }
    }
    clean(el, doc);
    // listes imbriquées directement dans une liste : rattachées à l'élément précédent
    if (T === 'UL' || T === 'OL') {
      [...el.children].forEach(c => {
        if (c.tagName === 'UL' || c.tagName === 'OL') {
          const prev = c.previousElementSibling;
          if (prev && prev.tagName === 'LI') prev.appendChild(c);
          else { const li = doc.createElement('li'); c.replaceWith(li); li.appendChild(c); }
        } else if (c.tagName !== 'LI') { const li = doc.createElement('li'); c.replaceWith(li); li.appendChild(c); }
      });
    }
    if (T === 'P' && el.querySelector('p,h1,h2,h3,ul,ol,table,blockquote,hr,div')) el.replaceWith(...el.childNodes);
  }
}

/* =========================================================
   Traitement de texte
   ========================================================= */
const ed = $('#ed');
let savedRange = null;
try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch {}

function openWriter() {
  show('writer');
  ed.innerHTML = sanitize(state.cur.html);
  applyPageLayout(); closeFind(); hideTblBar(); hideLinkBar(); closeWPops();
  $('#wScroll').scrollTop = 0;
  savedRange = null;
  if (ls.get('grim.outline', true) === false) document.body.classList.add('no-outline');
  CM.active = CM.edit = null; CM.panel = false; cmPurgeOrphans();
  applyDocStyles(); refreshWriter(); renderComments(); renderTrackBar();
}
const refreshWriter = () => { updateOutline(); updateStats(); updateToolState(); };
const refreshSoon = debounce(() => { updateOutline(); updateStats(); }, 350);

ed.addEventListener('input', () => {
  if (!state.cur) return;
  if (!ed.firstElementChild && !ed.textContent) ed.innerHTML = '<p><br></p>';
  normalizeBlocks();
  state.cur.html = edHtml(); autoTitle(); scheduleSave(); refreshSoon();
});
// Tant qu'il n'a pas été renommé, un nouveau document prend pour titre sa première ligne.
function autoTitle() {
  if (!state.cur || !state.cur.autoTitle) return;
  const line = [...ed.children].map(el => (el.textContent || '').replace(/\s+/g, ' ').trim()).find(Boolean) || '';
  let t = line.length > 60 ? line.slice(0, 60).replace(/\s+\S*$/, '') + '…' : line;
  t = t.replace(/[.:;,]+$/, '') || 'Document sans titre';
  if (t !== state.cur.title) { state.cur.title = t; $('#titleIn').value = t; }
}
ed.addEventListener('paste', e => {
  const cd = e.clipboardData; if (!cd) return;
  e.preventDefault();
  if (typeof trOn === 'function' && trOn()) { const t = cd.getData('text/plain'); if (t) { trInsert(t.replace(/\r?\n+/g, ' ')); trAfter(); } return; }
  const html = cd.getData('text/html');
  const imgFile = [...(cd.files || [])].find(f => /^image\//.test(f.type));
  if (imgFile && !html) { insertImageFile(imgFile); return; }
  if (html) { droppedImages = 0; exec('insertHTML', sanitize(html)); if (droppedImages) toast('Les images venant d\'un site ne sont pas reprises : enregistrez-les puis insérez-les avec le bouton Image.'); }
  else exec('insertText', cd.getData('text/plain'));
});
document.addEventListener('selectionchange', () => {
  if (state.view !== 'writer') return;
  const s = getSelection(); if (!s.rangeCount) return;
  const r = s.getRangeAt(0);
  if (ed.contains(r.commonAncestorContainer)) { savedRange = r.cloneRange(); updateToolState(); if (closestEl('TD') || closestEl('TH')) showTblBar(); else hideTblBar(); if (!$('#plume').hidden && Plume.mode === 'texte') Plume.ctx(); }
});
function restoreRange() {
  ed.focus({ preventScroll: true });
  if (savedRange) { const s = getSelection(); s.removeAllRanges(); s.addRange(savedRange); }
}
// Le navigateur imbrique parfois listes et tableaux dans un <p> : on les remet à plat sans perdre le curseur.
function normalizeBlocks() {
  const bad = ed.querySelectorAll('p:has(> ul, > ol, > p, > table, > h1, > h2, > h3, > blockquote, > hr, > div)');
  if (!bad.length) return;
  const s = getSelection(), an = s.anchorNode, ao = s.anchorOffset, fn = s.focusNode, fo = s.focusOffset;
  bad.forEach(p => { if (p.isConnected) p.replaceWith(...p.childNodes); });
  if (an && ed.contains(an)) { try { s.setBaseAndExtent(an, ao, fn, fo); } catch {} }
}
function exec(cmd, val) {
  if (document.activeElement !== ed) restoreRange();
  document.execCommand(cmd, false, val);
  normalizeBlocks();
  state.cur.html = edHtml(); scheduleSave(); refreshSoon(); updateToolState();
}
function currentBlock() {
  const s = getSelection(); if (!s.rangeCount) return null;
  let n = s.anchorNode; if (!n || !ed.contains(n)) return null;
  while (n && n !== ed) { if (n.nodeType === 1 && /^(P|H1|H2|H3|BLOCKQUOTE|LI|TD|TH)$/.test(n.tagName)) return n; n = n.parentNode; }
  return null;
}
function updateToolState() {
  if (state.view !== 'writer') return;
  const q = c => { try { return document.queryCommandState(c); } catch { return false; } };
  const set = (k, v) => { const b = $(`[data-w="${k}"]`); if (b) b.setAttribute('aria-pressed', String(!!v)); };
  set('bold', q('bold')); set('italic', q('italic')); set('underline', q('underline'));
  set('ul', q('insertUnorderedList')); set('ol', q('insertOrderedList'));
  const b = currentBlock();
  const al = b ? (b.style.textAlign || getComputedStyle(b).textAlign) : 'left';
  set('left', al === 'left' || al === 'start'); set('center', al === 'center'); set('right', al === 'right'); set('justify', al === 'justify');
  const tag = b ? b.tagName.toLowerCase() : 'p';
  $('#blockSel').value = ['h1', 'h2', 'h3', 'blockquote'].includes(tag) ? tag : 'p';
  set('strike', q('strikeThrough'));
  set('link', !!closestEl('A'));
  const li0 = closestEl('LI'); set('check', !!(li0 && li0.parentElement && li0.parentElement.hasAttribute('data-check')));
  let fam = ''; try { fam = cssFont(document.queryCommandValue('fontName')); } catch {}
  $('#fontSel').value = fam && fam !== 'Source Serif 4' ? fam : '';
  set('outline', !document.body.classList.contains('no-outline'));
  set('focus', document.body.classList.contains('focus'));
}
/* --- outils du traitement de texte */
const edHtml = () => ed.innerHTML.replace(/ class="sel"/g, '');
function afterEdit() {
  if (!state.cur) return;
  normalizeBlocks();
  state.cur.html = edHtml(); scheduleSave(); refreshSoon(); updateToolState();
  if (!$('#findBar').hidden) findSoon();
}
function execCss(cmd, val) {
  if (document.activeElement !== ed) restoreRange();
  try { document.execCommand('styleWithCSS', false, true); } catch {}
  document.execCommand(cmd, false, val);
  try { document.execCommand('styleWithCSS', false, false); } catch {}
  afterEdit();
}
const closestEl = (tag, from) => {
  let n = from || (getSelection().rangeCount ? getSelection().anchorNode : null);
  if (!n && savedRange) n = savedRange.startContainer;
  while (n && n !== ed) { if (n.nodeType === 1 && n.tagName === tag) return n; n = n.parentNode; }
  return null;
};
function selectedBlocks() {
  const r = savedRange; if (!r) return [];
  const all = $$('p,h1,h2,h3,blockquote,li', ed).filter(b => !b.closest('table') && !b.closest('[data-toc]'));
  const hit = all.filter(b => r.intersectsNode(b));
  return hit.length ? hit.filter(b => !hit.some(o => o !== b && o.contains(b) && o.tagName !== 'LI')) : (currentBlock() ? [currentBlock()] : []);
}
function setFontSize(pt) {
  if (document.activeElement !== ed) restoreRange();
  const s = getSelection();
  if (!s.rangeCount || s.isCollapsed) { toast('Sélectionnez d\'abord le texte à agrandir ou réduire.'); return; }
  try { document.execCommand('styleWithCSS', false, true); } catch {}
  document.execCommand('fontSize', false, '7');
  try { document.execCommand('styleWithCSS', false, false); } catch {}
  ed.querySelectorAll('span[style*="xxx-large"], font[size="7"]').forEach(el => {
    if (el.tagName === 'FONT') { const sp = document.createElement('span'); sp.style.fontSize = pt + 'pt'; while (el.firstChild) sp.appendChild(el.firstChild); el.replaceWith(sp); }
    else el.style.fontSize = pt + 'pt';
  });
  afterEdit();
}
function insertTable(rows = 3, cols = 3) {
  const th = Array.from({ length: cols }, (_, i) => `<th>Colonne ${i + 1}</th>`).join('');
  const tr = `<tr>${'<td><br></td>'.repeat(cols)}</tr>`.repeat(rows - 1);
  exec('insertHTML', `<table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table><p><br></p>`);
}
function indent(dir) {
  if (document.activeElement !== ed) restoreRange();
  if (closestEl('LI')) { document.execCommand(dir > 0 ? 'indent' : 'outdent'); return afterEdit(); }
  selectedBlocks().forEach(b => {
    const v = Math.max(0, Math.min(480, (parseInt(b.style.marginLeft, 10) || 0) + dir * 40));
    b.style.marginLeft = v ? v + 'px' : '';
  });
  afterEdit();
}
function toggleChecklist() {
  if (document.activeElement !== ed) restoreRange();
  const li = closestEl('LI'), list = li && li.parentElement;
  if (list && list.tagName === 'UL' && list.hasAttribute('data-check')) { list.removeAttribute('data-check'); return afterEdit(); }
  if (!list || list.tagName !== 'UL') document.execCommand('insertUnorderedList');
  const li2 = closestEl('LI'); if (li2) li2.parentElement.setAttribute('data-check', '1');
  afterEdit();
}
function insertPageBreak() { exec('insertHTML', '<hr data-pb="1"><p><br></p>'); }
function insertToc() {
  if (!$$('h1,h2,h3', ed).length) toast('Le sommaire se remplira avec les titres du document (Titre 1 à Titre 3).');
  exec('insertHTML', '<div data-toc="1" contenteditable="false"></div><p><br></p>');
  renderToc();
}
function renderToc() {
  const hs = $$('h1,h2,h3', ed).filter(h => h.textContent.trim());
  $$('[data-toc]', ed).forEach(box => {
    box.setAttribute('contenteditable', 'false');
    box.innerHTML = `<p class="toc-t">Sommaire</p>` + (hs.length ? hs.map((h, i) => `<p class="toc-l${h.tagName[1]}" data-h="${i}">${esc(h.textContent.trim())}</p>`).join('') : '<p class="toc-e">Ajoutez des titres pour remplir le sommaire.</p>');
  });
}
ed.addEventListener('mousedown', e => {
  const li = e.target.closest && e.target.closest('ul[data-check] > li');
  if (li && e.target === li && e.offsetX < 26) e.preventDefault();
  if (e.target.closest && e.target.closest('[data-toc]')) e.preventDefault();
});
ed.addEventListener('click', e => {
  const tocLine = e.target.closest('[data-toc] [data-h]');
  if (tocLine) { const h = $$('h1,h2,h3', ed).filter(x => x.textContent.trim())[+tocLine.dataset.h]; if (h) h.scrollIntoView({ behavior: 'smooth', block: 'start' }); return; }
  const li = e.target.closest('ul[data-check] > li');
  if (li && e.target === li && e.offsetX < 26) {
    e.preventDefault();
    if (li.getAttribute('data-done') === '1') li.removeAttribute('data-done'); else li.setAttribute('data-done', '1');
    afterEdit(); return;
  }
  const a = e.target.closest('a');
  if (a && ed.contains(a)) showLinkBar(a); else hideLinkBar();
  const cell = e.target.closest('td,th');
  if (cell && ed.contains(cell)) showTblBar(); else hideTblBar();
});

/* --- liens */
let linkEl = null;
function showLinkBar(a) {
  linkEl = a;
  const bar = $('#linkBar'), host = $('.w-body').getBoundingClientRect(), r = a.getBoundingClientRect();
  $('#lbUrl').textContent = a.getAttribute('href');
  $('#lbOpen').href = a.getAttribute('href');
  bar.hidden = false;
  bar.style.left = Math.max(8, Math.min(r.left - host.left, host.width - bar.offsetWidth - 8)) + 'px';
  bar.style.top = (r.bottom - host.top + 6) + 'px';
}
function hideLinkBar() { linkEl = null; $('#linkBar').hidden = true; }
function openLinkPop(existing) {
  if (document.activeElement === ed && getSelection().rangeCount) savedRange = getSelection().getRangeAt(0).cloneRange();
  const a = existing || closestEl('A');
  const p = $('#linkPop');
  $('#lpUrl').value = a ? a.getAttribute('href') : '';
  const collapsed = !savedRange || savedRange.collapsed;
  $('#lpTextRow').hidden = !!a || !collapsed;
  $('#lpText').value = '';
  $('#lpDel').hidden = !a;
  p.dataset.edit = a ? '1' : '';
  linkEl = a || null;
  placeWPop(p, $('[data-w="link"]'));
  setTimeout(() => $('#lpUrl').focus(), 30);
}
function normUrl(u) {
  u = u.trim(); if (!u) return '';
  if (/^(https?:|mailto:)/i.test(u)) return u;
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(u)) return 'mailto:' + u;
  return 'https://' + u.replace(/^\/+/, '');
}
function applyLink() {
  const url = normUrl($('#lpUrl').value);
  closeWPops();
  if (!url) return;
  if (linkEl && ed.contains(linkEl)) { linkEl.setAttribute('href', url); return afterEdit(); }
  restoreRange();
  const s = getSelection();
  if (!s.rangeCount || s.isCollapsed) {
    const t = $('#lpText').value.trim() || url.replace(/^(https?:\/\/|mailto:)/, '');
    document.execCommand('insertHTML', false, `<a href="${esc(url)}" target="_blank" rel="noopener">${esc(t)}</a>&nbsp;`);
  } else {
    document.execCommand('createLink', false, url);
    $$('a', ed).forEach(a => { if (a.getAttribute('href') === url) { a.target = '_blank'; a.rel = 'noopener'; } });
  }
  afterEdit();
}
function removeLink(a) {
  a = a || linkEl; if (!a) return;
  a.replaceWith(...a.childNodes); hideLinkBar(); closeWPops(); afterEdit();
}

/* --- tableaux */
const rowGrid = tr => { let c = 0; return [...tr.children].filter(x => x.tagName === 'TD' || x.tagName === 'TH').map(x => { const span = Math.max(1, x.colSpan || 1), o = { cell: x, start: c, span }; c += span; return o; }); };
const gridCols = t => Math.max(1, ...[...t.rows].map(tr => rowGrid(tr).reduce((s, o) => s + o.span, 0)));
function tblCtx() {
  let n = savedRange ? savedRange.startContainer : null;
  while (n && n !== ed && !(n.nodeType === 1 && (n.tagName === 'TD' || n.tagName === 'TH'))) n = n.parentNode;
  if (!n || n === ed) return null;
  const table = n.closest('table'), tr = n.parentElement;
  const o = rowGrid(tr).find(x => x.cell === n);
  return { cell: n, table, tr, start: o.start, span: o.span };
}
function newCell(th) { const c = document.createElement(th ? 'th' : 'td'); c.innerHTML = '<br>'; return c; }
function tblDone(ctx) { const cg = ctx && ctx.table.querySelector('colgroup'); if (cg && ctx.structural) cg.remove(); afterEdit(); showTblBar(); }
const TB = {
  rowBelow() { const c = tblCtx(); if (!c) return; const n = gridCols(c.table), inHead = c.tr.parentElement.tagName === 'THEAD'; const tr = document.createElement('tr'); for (let i = 0; i < n; i++) tr.appendChild(newCell(false)); if (inHead) { let tb = c.table.tBodies[0]; if (!tb) { tb = document.createElement('tbody'); c.table.appendChild(tb); } tb.insertBefore(tr, tb.firstChild); } else c.tr.after(tr); tblDone(c); },
  rowAbove() { const c = tblCtx(); if (!c) return; const n = gridCols(c.table), inHead = c.tr.parentElement.tagName === 'THEAD'; const tr = document.createElement('tr'); for (let i = 0; i < n; i++) tr.appendChild(newCell(inHead)); c.tr.before(tr); tblDone(c); },
  rowDel() { const c = tblCtx(); if (!c) return; if (c.table.rows.length <= 1) return TB.tableDel(); const sec = c.tr.parentElement; c.tr.remove(); if (!sec.rows.length) sec.remove(); tblDone(c); },
  colAt(c, b) {
    [...c.table.rows].forEach(tr => {
      const g = rowGrid(tr), th = tr.parentElement.tagName === 'THEAD';
      const across = g.find(o => o.start < b && b < o.start + o.span);
      if (across) { across.cell.colSpan = across.span + 1; return; }
      const at = g.find(o => o.start === b);
      if (at) at.cell.before(newCell(th)); else tr.appendChild(newCell(th));
    });
    c.structural = true; tblDone(c);
  },
  colRight() { const c = tblCtx(); if (c) TB.colAt(c, c.start + c.span); },
  colLeft() { const c = tblCtx(); if (c) TB.colAt(c, c.start); },
  colDel() {
    const c = tblCtx(); if (!c) return;
    if (gridCols(c.table) <= 1) return TB.tableDel();
    [...c.table.rows].forEach(tr => { const o = rowGrid(tr).find(x => x.start <= c.start && c.start < x.start + x.span); if (!o) return; if (o.span > 1) o.cell.colSpan = o.span - 1; else o.cell.remove(); });
    c.structural = true; tblDone(c);
  },
  merge() {
    const c = tblCtx(); if (!c) return;
    const nx = c.cell.nextElementSibling; if (!nx) return toast('Aucune cellule à droite à fusionner.');
    if (nx.textContent.trim()) { c.cell.appendChild(document.createTextNode(' ')); while (nx.firstChild) c.cell.appendChild(nx.firstChild); }
    c.cell.colSpan = c.span + Math.max(1, nx.colSpan || 1); nx.remove(); tblDone(c);
  },
  split() {
    const c = tblCtx(); if (!c) return;
    if (c.span < 2) return toast('Cette cellule n\'est pas fusionnée.');
    const th = c.cell.tagName === 'TH';
    c.cell.removeAttribute('colspan');
    for (let i = 1; i < c.span; i++) c.cell.after(newCell(th));
    tblDone(c);
  },
  header() {
    const c = tblCtx(); if (!c) return;
    const t = c.table;
    if (t.tHead) {
      const tb = t.tBodies[0] || t.appendChild(document.createElement('tbody'));
      [...t.tHead.rows].reverse().forEach(tr => { [...tr.children].forEach(th => { const td = document.createElement('td'); td.innerHTML = th.innerHTML; if (th.colSpan > 1) td.colSpan = th.colSpan; td.style.cssText = th.style.cssText; th.replaceWith(td); }); tb.insertBefore(tr, tb.firstChild); });
      t.tHead.remove();
    } else {
      const first = t.rows[0]; const th = t.createTHead();
      [...first.children].forEach(td => { const h = document.createElement('th'); h.innerHTML = td.innerHTML; if (td.colSpan > 1) h.colSpan = td.colSpan; h.style.cssText = td.style.cssText; td.replaceWith(h); });
      th.appendChild(first);
    }
    tblDone(c);
  },
  fill() { const c = tblCtx(); if (c) openCellColor(); },
  tableDel() { const c = tblCtx(); if (!c) return; c.table.remove(); hideTblBar(); afterEdit(); }
};
function showTblBar() {
  const c = tblCtx(); const bar = $('#tblBar');
  if (!c || state.view !== 'writer') { bar.hidden = true; return; }
  const host = $('.w-body').getBoundingClientRect(), r = c.table.getBoundingClientRect();
  bar.hidden = false;
  bar.style.left = Math.max(8, Math.min(r.left - host.left, host.width - bar.offsetWidth - 8)) + 'px';
  bar.style.top = Math.max(4, r.top - host.top - bar.offsetHeight - 6) + 'px';
}
function hideTblBar() { $('#tblBar').hidden = true; }
$('#tblBar').addEventListener('mousedown', e => e.preventDefault());
$('#tblBar').addEventListener('click', e => { const b = e.target.closest('[data-tb]'); if (b) TB[b.dataset.tb](); });
// largeur des colonnes à la souris (PC)
let colDrag = null;
ed.addEventListener('mousemove', e => {
  if (colDrag) return;
  const cell = e.target.closest && e.target.closest('td,th');
  const edge = cell && ed.contains(cell) && cell.nextElementSibling && e.offsetX >= cell.clientWidth - 6;
  ed.style.cursor = edge ? 'col-resize' : '';
});
ed.addEventListener('mousedown', e => {
  const cell = e.target.closest && e.target.closest('td,th');
  if (!cell || !ed.contains(cell) || !cell.nextElementSibling || e.offsetX < cell.clientWidth - 6) return;
  e.preventDefault();
  const table = cell.closest('table'), n = gridCols(table);
  let cg = table.querySelector('colgroup');
  if (!cg || cg.children.length !== n) {
    if (cg) cg.remove();
    const row = [...table.rows].find(tr => rowGrid(tr).length === n) || table.rows[0];
    const tw = table.getBoundingClientRect().width;
    cg = document.createElement('colgroup');
    const g = rowGrid(row);
    for (let i = 0; i < n; i++) { const col = document.createElement('col'); const o = g[i]; col.style.width = (o && o.span === 1 ? o.cell.getBoundingClientRect().width / tw * 100 : 100 / n).toFixed(2) + '%'; cg.appendChild(col); }
    table.insertBefore(cg, table.firstChild);
  }
  const o = rowGrid(cell.parentElement).find(x => x.cell === cell), i = o.start + o.span - 1;
  colDrag = { table, cols: [...cg.children], i, x: e.clientX, w: table.getBoundingClientRect().width, a: parseFloat(cg.children[i].style.width), b: parseFloat(cg.children[i + 1].style.width) };
});
document.addEventListener('mousemove', e => {
  if (!colDrag) return;
  const d = (e.clientX - colDrag.x) / colDrag.w * 100, tot = colDrag.a + colDrag.b;
  const a = Math.max(5, Math.min(tot - 5, colDrag.a + d));
  colDrag.cols[colDrag.i].style.width = a.toFixed(2) + '%'; colDrag.cols[colDrag.i + 1].style.width = (tot - a).toFixed(2) + '%';
});
document.addEventListener('mouseup', () => { if (colDrag) { colDrag = null; afterEdit(); } });

/* --- couleurs, polices, interligne : petites fenêtres */
const TXT_COLORS = [['Encre', '#1E1814'], ['Gris', '#6B6B6B'], ['Rouge', '#C0392B'], ['Orange', '#D35400'], ['Vert', '#1E7B45'], ['Bleu', '#1F5FAD'], ['Violet', '#6C3FA8'], ['Cuivre', '#A35E2C']];
const HI_COLORS = [['Jaune', '#FFF1A8'], ['Vert', '#CFF0D2'], ['Bleu', '#D3E6FA'], ['Rose', '#FAD4DD'], ['Orange', '#FDDFC3'], ['Gris', '#E6E6E6']];
function placeWPop(p, anchor) {
  closeWPops(p);
  p.hidden = false;
  const host = $('#writer').getBoundingClientRect(), a = (anchor || $('#wTools')).getBoundingClientRect();
  p.style.left = Math.max(8, Math.min(a.left - host.left, host.width - p.offsetWidth - 8)) + 'px';
  if (a.top - host.top > host.height / 2) { p.style.top = 'auto'; p.style.bottom = (host.bottom - a.top + 6) + 'px'; }
  else { p.style.bottom = 'auto'; p.style.top = (a.bottom - host.top + 6) + 'px'; }
}
function closeWPops(except) { ['#wPop', '#linkPop', '#pagePop', '#stylePop'].forEach(s => { if ($(s) !== except) $(s).hidden = true; }); }
function openColorPop(kind) {
  if (document.activeElement === ed && getSelection().rangeCount) savedRange = getSelection().getRangeAt(0).cloneRange();
  const p = $('#wPop'), list = kind === 'fc' ? TXT_COLORS : HI_COLORS;
  p.dataset.k = kind;
  p.innerHTML = `<div class="mp-t">${kind === 'fc' ? 'Couleur du texte' : 'Surlignage'}</div><div class="swatches">`
    + list.map(([n, h]) => `<button class="sw" data-hex="${h}" title="${n}" aria-label="${n}" style="${kind === 'fc' ? `color:${h};background:#FFFDF9` : `background:${h}`}">${kind === 'fc' ? 'A' : ''}</button>`).join('')
    + `<button class="sw none" data-hex="" aria-label="Aucune">Aucune</button></div>`;
  placeWPop(p, $(`[data-w="${kind === 'fc' ? 'color' : 'hilite'}"]`));
}
function openCellColor() {
  const p = $('#wPop'); p.dataset.k = 'cell';
  p.innerHTML = `<div class="mp-t">Fond de la cellule</div><div class="swatches">` + HI_COLORS.concat([['Sable', '#EAD9B8']]).map(([n, h]) => `<button class="sw" data-hex="${h}" title="${n}" aria-label="${n}" style="background:${h}"></button>`).join('') + `<button class="sw none" data-hex="" aria-label="Aucune">Aucune</button></div>`;
  placeWPop(p, $('#tblBar'));
}
function openSpacingPop() {
  if (document.activeElement === ed && getSelection().rangeCount) savedRange = getSelection().getRangeAt(0).cloneRange();
  const p = $('#wPop'); p.dataset.k = 'sp';
  const b = selectedBlocks()[0], lh = b ? b.style.lineHeight : '', mb = b ? b.style.marginBottom : '';
  p.innerHTML = `<div class="mp-t">Interligne</div>` + [['', 'Normal'], ['1', 'Simple'], ['1.15', '1,15'], ['1.5', '1,5'], ['2', 'Double']].map(([v, l]) => `<button data-lh="${v}" aria-pressed="${lh === v}">${l}</button>`).join('')
    + `<hr><div class="mp-t">Espace après les paragraphes</div>` + [['0px', 'Aucun'], ['', 'Normal'], ['1.5em', 'Grand']].map(([v, l]) => `<button data-mb="${v}" aria-pressed="${mb === v}">${l}</button>`).join('');
  placeWPop(p, $('[data-w="spacing"]'));
}
$('#wPop').addEventListener('mousedown', e => e.preventDefault());
$('#wPop').addEventListener('click', e => {
  const p = $('#wPop'), k = p.dataset.k;
  const sw = e.target.closest('[data-hex]');
  if (sw) {
    closeWPops(); const h = sw.dataset.hex;
    if (k === 'cell') { const c = tblCtx(); if (c) { if (h) c.cell.style.backgroundColor = h; else c.cell.style.backgroundColor = ''; afterEdit(); } return; }
    if (k === 'fc') execCss('foreColor', h || '#1E1814');
    else execCss('hiliteColor', h || 'transparent');
    if (!h) ed.querySelectorAll('span[style*="background-color: transparent"]').forEach(s => { s.style.backgroundColor = ''; if (!s.getAttribute('style')) s.replaceWith(...s.childNodes); });
    return;
  }
  const lh = e.target.closest('[data-lh]'); if (lh) { selectedBlocks().forEach(b => { b.style.lineHeight = lh.dataset.lh; }); closeWPops(); return afterEdit(); }
  const mb = e.target.closest('[data-mb]'); if (mb) { selectedBlocks().forEach(b => { b.style.marginBottom = mb.dataset.mb; }); closeWPops(); return afterEdit(); }
});
$('#fontSel').addEventListener('change', e => { const v = e.target.value; if (v) execCss('fontName', FONT_STACK[v]); });
$('#sizeSel').addEventListener('change', e => { const v = +e.target.value; e.target.value = ''; if (v) setFontSize(v); });
$('#linkPop').addEventListener('submit', e => { e.preventDefault(); applyLink(); });
$('#lpDel').addEventListener('click', () => removeLink(linkEl));
$('#lpX').addEventListener('click', () => closeWPops());
$('#linkBar').addEventListener('mousedown', e => { if (!e.target.closest('#lbOpen')) e.preventDefault(); });
$('#lbEdit').addEventListener('click', () => { const a = linkEl; hideLinkBar(); openLinkPop(a); });
$('#lbDel').addEventListener('click', () => removeLink(linkEl));
document.addEventListener('mousedown', e => {
  if (e.target.closest('#wPop, #linkPop, #pagePop, #stylePop, #wTools, #tblBar, #linkBar')) return;
  closeWPops();
});

/* --- rechercher et remplacer */
const FIND = { hits: [], i: 0 };
const HL = !!(window.CSS && CSS.highlights && window.Highlight);
function findRun(keep) {
  const q = $('#fIn').value, cs = $('#fCase').checked;
  FIND.hits = [];
  if (q) {
    const qq = cs ? q : q.toLowerCase();
    const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = w.nextNode())) {
      if (n.parentElement && n.parentElement.closest('[data-toc]')) continue;
      const t = cs ? n.nodeValue : n.nodeValue.toLowerCase();
      for (let k = t.indexOf(qq); k >= 0; k = t.indexOf(qq, k + qq.length)) { const r = document.createRange(); r.setStart(n, k); r.setEnd(n, k + q.length); FIND.hits.push(r); }
    }
  }
  if (!keep) FIND.i = 0;
  if (FIND.i >= FIND.hits.length) FIND.i = 0;
  paintFind();
}
const findSoon = debounce(() => findRun(true), 250);
function paintFind(scroll) {
  const n = FIND.hits.length;
  $('#fCount').textContent = $('#fIn').value ? (n ? `${FIND.i + 1} / ${n}` : 'Aucun résultat') : '';
  if (HL) {
    CSS.highlights.set('g-find', new Highlight(...FIND.hits));
    CSS.highlights.set('g-cur', new Highlight(...(n ? [FIND.hits[FIND.i]] : [])));
  }
  if (scroll && n) {
    const r = FIND.hits[FIND.i].getBoundingClientRect(), sc = $('#wScroll'), box = sc.getBoundingClientRect();
    if (r.top < box.top + 40 || r.bottom > box.bottom - 40) sc.scrollTop += r.top - box.top - box.height / 3;
    if (!HL) { const s = getSelection(); s.removeAllRanges(); s.addRange(FIND.hits[FIND.i]); }
  }
}
function findGo(d) { const n = FIND.hits.length; if (!n) return; FIND.i = (FIND.i + d + n) % n; paintFind(true); }
function replaceOne() {
  if (!FIND.hits.length) return;
  const r = FIND.hits[FIND.i];
  ed.focus({ preventScroll: true });
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  document.execCommand('insertText', false, $('#rIn').value);
  state.cur.html = edHtml(); scheduleSave(); refreshSoon();
  findRun(true); paintFind(true);
  $('#rIn').focus();
}
function replaceAll() {
  const n = FIND.hits.length; if (!n) return;
  const rep = $('#rIn').value;
  [...FIND.hits].reverse().forEach(r => { r.deleteContents(); if (rep) r.insertNode(document.createTextNode(rep)); });
  ed.normalize();
  state.cur.html = edHtml(); scheduleSave(); refreshSoon();
  findRun();
  toast(`${n} remplacement${n > 1 ? 's' : ''} effectué${n > 1 ? 's' : ''}.`);
}
function openFind(withReplace) {
  const bar = $('#findBar');
  bar.hidden = false; bar.classList.toggle('with-rep', !!withReplace || bar.classList.contains('with-rep'));
  const sel = getSelection().toString().trim();
  if (sel && sel.length < 80 && !sel.includes('\n')) $('#fIn').value = sel;
  $('#fIn').focus(); $('#fIn').select();
  findRun(); paintFind(true);
}
function closeFind() {
  $('#findBar').hidden = true;
  if (HL) { CSS.highlights.delete('g-find'); CSS.highlights.delete('g-cur'); }
  FIND.hits = [];
}
$('#fIn').addEventListener('input', () => { findRun(); paintFind(true); });
$('#fCase').addEventListener('change', () => { findRun(); paintFind(true); });
$('#fIn').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); findGo(e.shiftKey ? -1 : 1); } else if (e.key === 'Escape') { e.preventDefault(); closeFind(); restoreRange(); } });
$('#rIn').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); replaceOne(); } else if (e.key === 'Escape') { e.preventDefault(); closeFind(); } });
$('#findBar').addEventListener('click', e => {
  const b = e.target.closest('[data-f]'); if (!b) return;
  ({ prev: () => findGo(-1), next: () => findGo(1), one: replaceOne, all: replaceAll, rep: () => { $('#findBar').classList.toggle('with-rep'); $('#rIn').focus(); }, close: closeFind })[b.dataset.f]();
});

/* --- mise en page */
const PAGE_DEF = { orient: 'portrait', margin: 'normal', header: '', footer: '', num: 'none' };
const MARGIN_PX = { normal: 94, narrow: 48, wide: 121 };
const MARGIN_TW = { normal: 1418, narrow: 720, wide: 1814 };
const pageOf = () => Object.assign({}, PAGE_DEF, (state.cur && state.cur.page) || {});
function applyPageLayout() {
  const P = pageOf();
  ed.classList.toggle('land', P.orient === 'landscape');
  ed.style.setProperty('--pm', MARGIN_PX[P.margin] + 'px');
  ed.dataset.header = P.header || '';
  ed.dataset.footer = [P.footer, P.num === 'none' ? '' : P.num === 'xofy' ? 'Page 1 sur N' : 'Page 1'].filter(Boolean).join('   ·   ');
  const mm = { normal: 25, narrow: 12.7, wide: 32 }[P.margin];
  const q = s => '"' + String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
  const numTxt = P.num === 'xofy' ? '"Page " counter(page) " sur " counter(pages)' : 'counter(page)';
  let css = `@page{size:A4 ${P.orient};margin:${mm}mm}`;
  const boxes = [];
  if (P.header) boxes.push(`@top-center{content:${q(P.header)};font:9pt sans-serif;color:#555}`);
  if (P.footer) boxes.push(`@bottom-left{content:${q(P.footer)};font:9pt sans-serif;color:#555}`);
  if (P.num !== 'none') boxes.push(`@bottom-${P.num === 'right' ? 'right' : 'center'}{content:${numTxt};font:9pt sans-serif;color:#555}`);
  if (boxes.length) css += `@page{${boxes.join('')}}`;
  $('#pageCss').textContent = state.view === 'writer' ? css : '';
}
function openPagePop() {
  const P = pageOf(), p = $('#pagePop');
  const seg = (k, opts) => `<div class="seg" data-pk="${k}">` + opts.map(([v, l]) => `<button type="button" data-v="${v}" aria-pressed="${P[k] === v}">${l}</button>`).join('') + '</div>';
  p.innerHTML = `<div class="pp-h"><b>Mise en page</b><button type="button" class="xbtn" data-pclose aria-label="Fermer">${ic('close')}</button></div>
    <div class="pp-row"><span>Orientation</span>${seg('orient', [['portrait', 'Portrait'], ['landscape', 'Paysage']])}</div>
    <div class="pp-row"><span>Marges</span>${seg('margin', [['narrow', 'Étroites'], ['normal', 'Normales'], ['wide', 'Larges']])}</div>
    <label class="field" for="ppHead"><span>En-tête</span><input id="ppHead" value="${esc(P.header)}" placeholder="Texte en haut de chaque page" autocomplete="off"></label>
    <label class="field" for="ppFoot"><span>Pied de page</span><input id="ppFoot" value="${esc(P.footer)}" placeholder="Texte en bas de chaque page" autocomplete="off"></label>
    <div class="pp-row"><span>Numéros de page</span>${seg('num', [['none', 'Aucun'], ['center', 'Centre'], ['right', 'Droite'], ['xofy', 'Page X sur Y']])}</div>
    <p class="set-note">S'appliquent à l'impression, au PDF et à l'export Word.</p>`;
  placeWPop(p, $('[data-w="layout"]'));
}
function setPage(k, v) { state.cur.page = Object.assign(pageOf(), { [k]: v }); applyPageLayout(); scheduleSave(); refreshSoon(); }
$('#pagePop').addEventListener('click', e => {
  if (e.target.closest('[data-pclose]')) return closeWPops();
  const b = e.target.closest('.seg [data-v]'); if (!b) return;
  setPage(b.parentElement.dataset.pk, b.dataset.v);
  $$('button', b.parentElement).forEach(x => x.setAttribute('aria-pressed', String(x === b)));
});
$('#pagePop').addEventListener('input', e => { if (e.target.id === 'ppHead') setPage('header', e.target.value); if (e.target.id === 'ppFoot') setPage('footer', e.target.value); });
// repères de pages (aperçu) et nombre de pages
function updatePageMarks() {
  const box = $('#pgMarks');
  if (state.view !== 'writer' || isPhone()) { box.innerHTML = ''; return 0; }
  const P = pageOf(), land = P.orient === 'landscape';
  const m = MARGIN_PX[P.margin], pageH = (land ? 794 : 1123) - 2 * m;
  let start = ed.firstElementChild ? ed.firstElementChild.offsetTop : m, marks = [];
  for (const b of ed.children) {
    const top = b.offsetTop, bottom = top + b.offsetHeight;
    if (b.tagName === 'HR' && b.hasAttribute('data-pb')) { marks.push({ y: bottom, manual: true }); start = bottom; continue; }
    if (bottom - start > pageH) {
      if (top > start + 4) { marks.push({ y: top }); start = top; }
      while (bottom - start > pageH) { start += pageH; marks.push({ y: start }); }
    }
  }
  box.innerHTML = marks.map((k, i) => `<div class="pgm${k.manual ? ' man' : ''}" style="top:${ed.offsetTop + k.y - 1}px"><span>Page ${i + 2}</span></div>`).join('');
  return marks.length + 1;
}

/* --- barre d'outils */
const W = {
  comment: () => addComment(), cols: () => openColsPop(), styles: () => openStylePop(), track: () => toggleTrack(),
  undo: () => exec('undo'), redo: () => exec('redo'),
  bold: () => exec('bold'), italic: () => exec('italic'), underline: () => exec('underline'),
  strike: () => exec('strikeThrough'), sup: () => exec('superscript'), sub: () => exec('subscript'),
  clear: () => { exec('removeFormat'); },
  color: () => openColorPop('fc'), hilite: () => openColorPop('bg'),
  link: () => openLinkPop(),
  ul: () => exec('insertUnorderedList'), ol: () => exec('insertOrderedList'), check: toggleChecklist,
  indent: () => indent(1), outdent: () => indent(-1),
  left: () => exec('justifyLeft'), center: () => exec('justifyCenter'), right: () => exec('justifyRight'), justify: () => exec('justifyFull'),
  spacing: openSpacingPop,
  table: () => insertTable(), hr: () => exec('insertHorizontalRule'), pagebreak: insertPageBreak, toc: insertToc,
  image: () => { if (document.activeElement !== ed) restoreRange(); $('#imgIn').click(); },
  find: () => openFind(false), layout: openPagePop,
  block: v => exec('formatBlock', '<' + v + '>'),
  outline: () => {
    if (isPhone() || innerWidth <= 900) { document.body.classList.toggle('show-outline'); }
    else { const off = document.body.classList.toggle('no-outline'); ls.set('grim.outline', !off); }
    updateToolState();
  },
  focus: () => { document.body.classList.toggle('focus'); updateToolState(); if (document.body.classList.contains('focus')) toast('Mode concentration. Ctrl+K puis « concentration » pour en sortir.'); }
};
$('#wTools').addEventListener('mousedown', e => { if (e.target.closest('button')) e.preventDefault(); });
$('#wTools').addEventListener('click', e => { const b = e.target.closest('[data-w]'); if (b) W[b.dataset.w](); });
$('#blockSel').addEventListener('change', e => { if (e.target.value === '__sty') { updateToolState(); const b = currentBlock(); STY.cur = b && STY_TAGS.some(([t]) => t === b.tagName.toLowerCase()) ? b.tagName.toLowerCase() : 'p'; openStylePop(); return; } W.block(e.target.value); });
ed.addEventListener('keydown', e => {
  const mod = e.ctrlKey || e.metaKey;
  if (e.key === 'Tab' && !e.altKey && !mod) {
    const cell = closestEl('TD', getSelection().anchorNode) || closestEl('TH', getSelection().anchorNode);
    e.preventDefault();
    if (cell && !closestEl('LI')) {
      const cells = [...cell.closest('table').querySelectorAll('td,th')];
      const nx = cells[cells.indexOf(cell) + (e.shiftKey ? -1 : 1)];
      if (nx) { const r = document.createRange(); r.selectNodeContents(nx); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
      else if (!e.shiftKey) { savedRange = getSelection().getRangeAt(0).cloneRange(); TB.rowBelow(); const last = [...cell.closest('table').rows].pop(); const r = document.createRange(); r.selectNodeContents(last.cells[0]); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); }
      return;
    }
    savedRange = getSelection().rangeCount ? getSelection().getRangeAt(0).cloneRange() : savedRange;
    indent(e.shiftKey ? -1 : 1);
    return;
  }
  if (mod && e.key === 'Enter') { e.preventDefault(); insertPageBreak(); return; }
  if (e.key === 'Enter' && !e.shiftKey) setTimeout(() => { const li = closestEl('LI'); if (li && li.getAttribute('data-done') === '1' && !li.textContent.trim()) li.removeAttribute('data-done'); }, 0);
});


const fileToDataURL = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
const loadImg = src => new Promise(res => { const i = new Image(); i.onload = () => res(i); i.onerror = () => res(null); i.src = src; });
async function insertImageFile(f) {
  if (!/^image\//.test(f.type)) return toast('Ce fichier n\'est pas une image.', true);
  const url = await fileToDataURL(f);
  const img = await loadImg(url);
  if (!img) return toast('Image illisible.', true);
  let out = url;
  const w = img.naturalWidth, h = img.naturalHeight, max = 1600;
  if (w > max || h > max || f.size > 700000 || !/^image\/(png|jpeg|gif)$/.test(f.type)) {
    const k = Math.min(1, max / Math.max(w, h));
    const c = document.createElement('canvas'); c.width = Math.round(w * k); c.height = Math.round(h * k);
    const g = c.getContext('2d');
    const keepPng = f.type === 'image/png' && f.size < 1500000;
    if (!keepPng) { g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); }
    g.drawImage(img, 0, 0, c.width, c.height);
    out = keepPng ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.85);
  }
  restoreRange();
  exec('insertHTML', `<img src="${out}" alt=""><p><br></p>`);
}
$('#imgIn').addEventListener('change', async e => { const f = e.target.files[0]; e.target.value = ''; if (f) { try { await insertImageFile(f); } catch { toast('Insertion de l\'image impossible.', true); } } });
let selImg = null;
function hideImgBar() { if (selImg) selImg.classList.remove('sel'); selImg = null; $('#imgBar').hidden = true; }
ed.addEventListener('click', e => {
  if (e.target.tagName !== 'IMG') return hideImgBar();
  hideImgBar(); selImg = e.target; selImg.classList.add('sel');
  const bar = $('#imgBar'), host = $('.w-body').getBoundingClientRect(), r = selImg.getBoundingClientRect();
  const wr = selImg.getAttribute('data-wrap') || 'none';
  $$('[data-wrap-set]', bar).forEach(b => b.setAttribute('aria-pressed', String(b.dataset.wrapSet === wr)));
  const pc = Math.round(parseFloat(selImg.style.width) || 100);
  $$('[data-iw]', bar).forEach(b => b.setAttribute('aria-pressed', String(b.dataset.iw === String(pc))));
  bar.hidden = false;
  bar.style.left = Math.max(8, Math.min(r.left - host.left + r.width / 2 - bar.offsetWidth / 2, host.width - bar.offsetWidth - 8)) + 'px';
  bar.style.top = Math.max(8, r.top - host.top - bar.offsetHeight - 8) + 'px';
});
$('#wScroll').addEventListener('scroll', () => { if (selImg) hideImgBar(); if (!$('#linkBar').hidden) hideLinkBar(); if (!$('#tblBar').hidden) showTblBar(); }, { passive: true });
window.addEventListener('resize', debounce(() => { if (state.view === 'writer') { updatePageMarks(); if (!$('#tblBar').hidden) showTblBar(); } }, 200));
$('#imgBar').addEventListener('mousedown', e => e.preventDefault());
$('#imgBar').addEventListener('click', e => {
  const w = e.target.closest('[data-wrap-set]');
  if (w && selImg) {
    const v = w.dataset.wrapSet;
    if (v === 'none') selImg.removeAttribute('data-wrap');
    else {
      selImg.setAttribute('data-wrap', v);
      // une image flottante pleine largeur ne laisserait aucune place au texte
      if (!selImg.style.width || parseFloat(selImg.style.width) > 60) selImg.style.width = '40%';
      // l'image se place au début du paragraphe suivant pour que le texte l'entoure
      const p = selImg.parentElement;
      if (p && p !== ed && p.tagName === 'P' && !p.textContent.trim() && p.querySelectorAll('img').length === 1) {
        const nx = p.nextElementSibling;
        if (nx && /^(P|H1|H2|H3)$/.test(nx.tagName) && nx.textContent.trim()) { nx.prepend(selImg); p.remove(); }
      } else if (p === ed) {
        const nx = selImg.nextElementSibling;
        if (nx && nx.tagName === 'P') nx.prepend(selImg);
      }
    }
    const img = selImg; hideImgBar(); updatePageMarks();
    state.cur.html = ed.innerHTML.replace(/ class="sel"/g, ''); scheduleSave();
    img.click();
    return;
  }
  const b = e.target.closest('[data-iw]'); if (!b || !selImg) return;
  if (b.dataset.iw === 'del') selImg.remove(); else { selImg.style.width = b.dataset.iw + '%'; selImg.classList.remove('sel'); }
  hideImgBar(); updatePageMarks();
  state.cur.html = ed.innerHTML.replace(/ class="sel"/g, ''); scheduleSave();
});

function updateOutline() {
  renderToc();
  const hs = $$('h1,h2,h3', ed).filter(h => h.textContent.trim());
  $('#outlineList').innerHTML = hs.length
    ? hs.map((h, i) => `<li><button class="l${h.tagName[1]}" data-i="${i}">${esc(h.textContent.trim())}</button></li>`).join('')
    : '<li class="none">Les titres de votre document apparaîtront ici. Utilisez « Titre 1 » à « Titre 3 ».</li>';
  $('#outlineList').onclick = e => {
    const b = e.target.closest('button'); if (!b) return;
    hs[+b.dataset.i].scrollIntoView({ behavior: 'smooth', block: 'start' });
    document.body.classList.remove('show-outline');
  };
}
function updateStats() {
  const txt = ed.innerText || '';
  const words = (txt.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu) || []).length;
  const chars = txt.replace(/\s/g, '').length;
  const mins = Math.max(1, Math.round(words / 230));
  const pg = updatePageMarks();
  const pages = pg || Math.max(1, Math.ceil(words / 480));
  const P = pageOf();
  const nCm = (state.cur && state.cur.comments || []).filter(c => !c.done).length;
  $('#wStatus').innerHTML = `<span><b>${words.toLocaleString('fr-FR')}</b> mots · ${chars.toLocaleString('fr-FR')} caractères · ${mins} min de lecture${nCm ? ` · <button type="button" class="w-cmt" id="wCmtBtn">${ic('comment')}${nCm} commentaire${nCm > 1 ? 's' : ''}</button>` : ''}</span><span>${pg ? '' : '≈ '}${pages} page${pages > 1 ? 's' : ''} A4 ${P.orient === 'landscape' ? 'paysage' : ''}</span>`;
}

/* =========================================================
   Tableur
   ========================================================= */
const FR2EN = {
  // maths
  'SOMME': 'SUM', 'ARRONDI': 'ROUND', 'ARRONDI.SUP': 'ROUNDUP', 'ARRONDI.INF': 'ROUNDDOWN', 'ARRONDI.AU.MULTIPLE': 'MROUND', 'TRONQUE': 'TRUNC',
  'ENT': 'INT', 'PLAFOND': 'CEILING', 'PLAFOND.MATH': 'CEILING.MATH', 'PLAFOND.PRECIS': 'CEILING.PRECISE', 'PLANCHER': 'FLOOR', 'PLANCHER.MATH': 'FLOOR.MATH',
  'PLANCHER.PRECIS': 'FLOOR.PRECISE', 'PAIR': 'EVEN', 'IMPAIR': 'ODD', 'PGCD': 'GCD', 'PPCM': 'LCM', 'PUISSANCE': 'POWER', 'RACINE': 'SQRT', 'PRODUIT': 'PRODUCT',
  'DEGRES': 'DEGREES', 'ALEA': 'RAND', 'ALEA.ENTRE.BORNES': 'RANDBETWEEN', 'ROMAIN': 'ROMAN', 'CHIFFRE.ARABE': 'ARABIC', 'SIGNE': 'SIGN',
  'SOMME.CARRES': 'SUMSQ', 'SOMMEPROD': 'SUMPRODUCT', 'SOUS.TOTAL': 'SUBTOTAL', 'MULTINOMIALE': 'MULTINOMIAL', 'PRODUITMAT': 'MMULT', 'COMBINA': 'COMBINA',
  // statistiques
  'MOYENNE': 'AVERAGE', 'SOMME.SI': 'SUMIF', 'SOMME.SI.ENS': 'SUMIFS', 'NB': 'COUNT', 'NBVAL': 'COUNTA', 'NB.VIDE': 'COUNTBLANK', 'NB.SI': 'COUNTIF',
  'NB.SI.ENS': 'COUNTIFS', 'MOYENNE.SI': 'AVERAGEIF', 'MOYENNE.SI.ENS': 'AVERAGEIFS', 'MAX.SI.ENS': 'MAXIFS', 'MIN.SI.ENS': 'MINIFS', 'MEDIANE': 'MEDIAN',
  'GRANDE.VALEUR': 'LARGE', 'PETITE.VALEUR': 'SMALL', 'RANG': 'RANK', 'EQUATION.RANG': 'RANK', 'ECARTYPE': 'STDEV', 'ECARTYPEP': 'STDEVP',
  'ECARTYPE.STANDARD': 'STDEV.S', 'ECARTYPE.PEARSON': 'STDEV.P', 'VAR.P': 'VARP', 'VAR.P.N': 'VAR.P', 'COEFFICIENT.CORRELATION': 'CORREL',
  'PENTE': 'SLOPE', 'COEFFICIENT.DETERMINATION': 'RSQ', 'ORDONNEE.ORIGINE': 'INTERCEPT', 'PREVISION': 'FORECAST', 'PREVISION.LINEAIRE': 'FORECAST',
  'ECART.MOYEN': 'AVEDEV', 'MOYENNE.GEOMETRIQUE': 'GEOMEAN', 'MOYENNE.HARMONIQUE': 'HARMEAN', 'CENTILE': 'PERCENTILE', 'CENTILE.INCLURE': 'PERCENTILE',
  'QUARTILE.INCLURE': 'QUARTILE', 'MODE.SIMPLE': 'MODE', 'CENTREE.REDUITE': 'STANDARDIZE', 'COVARIANCE': 'COVAR', 'LOI.NORMALE': 'NORMDIST',
  'LOI.NORMALE.N': 'NORM.DIST', 'LOI.NORMALE.INVERSE.N': 'NORM.INV', 'NB.VALEURS.UNIQUES': 'COUNTUNIQUE', 'MAXA': 'MAXA', 'MINA': 'MINA',
  // logique
  'SI': 'IF', 'ET': 'AND', 'OU': 'OR', 'NON': 'NOT', 'OUX': 'XOR', 'SI.CONDITIONS': 'IFS', 'SI.MULTIPLE': 'SWITCH', 'SIERREUR': 'IFERROR',
  'SI.NON.DISP': 'IFNA', 'VRAI': 'TRUE', 'FAUX': 'FALSE',
  // texte
  'CONCATENER': 'CONCATENATE', 'JOINDRE.TEXTE': 'TEXTJOIN', 'GAUCHE': 'LEFT', 'DROITE': 'RIGHT', 'STXT': 'MID', 'NBCAR': 'LEN', 'MAJUSCULE': 'UPPER',
  'MINUSCULE': 'LOWER', 'NOMPROPRE': 'PROPER', 'SUPPRESPACE': 'TRIM', 'EPURAGE': 'CLEAN', 'SUBSTITUE': 'SUBSTITUTE', 'REMPLACER': 'REPLACE',
  'CHERCHE': 'SEARCH', 'TROUVE': 'FIND', 'TEXTE': 'TEXT', 'CNUM': 'VALUE', 'CAR': 'CHAR', 'UNICAR': 'UNICHAR', 'FORMULETEXTE': 'FORMULATEXT',
  'DIVISER': 'SPLIT',
  // dates
  'AUJOURDHUI': 'TODAY', 'MAINTENANT': 'NOW', 'ANNEE': 'YEAR', 'MOIS': 'MONTH', 'JOUR': 'DAY', 'HEURE': 'HOUR', 'SECONDE': 'SECOND', 'TEMPS': 'TIME',
  'TEMPSVAL': 'TIMEVALUE', 'DATEVAL': 'DATEVALUE', 'JOURSEM': 'WEEKDAY', 'NO.SEMAINE': 'WEEKNUM', 'NO.SEMAINE.ISO': 'ISOWEEKNUM', 'MOIS.DECALER': 'EDATE',
  'FIN.MOIS': 'EOMONTH', 'NB.JOURS.OUVRES': 'NETWORKDAYS', 'NB.JOURS.OUVRES.INTL': 'NETWORKDAYS.INTL', 'SERIE.JOUR.OUVRE': 'WORKDAY',
  'SERIE.JOUR.OUVRE.INTL': 'WORKDAY.INTL', 'JOURS': 'DAYS', 'JOURS360': 'DAYS360', 'FRACTION.ANNEE': 'YEARFRAC',
  // recherche
  'RECHERCHEV': 'VLOOKUP', 'RECHERCHEH': 'HLOOKUP', 'RECHERCHE': 'LOOKUP', 'RECHERCHEX': 'XLOOKUP', 'EQUIV': 'MATCH', 'EQUIVX': 'XMATCH',
  'DECALER': 'OFFSET', 'CHOISIR': 'CHOOSE', 'COLONNE': 'COLUMN', 'COLONNES': 'COLUMNS', 'LIGNE': 'ROW', 'LIGNES': 'ROWS', 'ADRESSE': 'ADDRESS',
  'LIEN_HYPERTEXTE': 'HYPERLINK', 'FILTRE': 'FILTER', 'TRIER': 'SORT', 'FEUILLE': 'SHEET', 'FEUILLES': 'SHEETS',
  // informations
  'ESTVIDE': 'ISBLANK', 'ESTNUM': 'ISNUMBER', 'ESTTEXTE': 'ISTEXT', 'ESTERREUR': 'ISERROR', 'ESTERR': 'ISERR', 'ESTNA': 'ISNA', 'ESTLOGIQUE': 'ISLOGICAL',
  'ESTPAIR': 'ISEVEN', 'ESTIMPAIR': 'ISODD', 'ESTNONTEXTE': 'ISNONTEXT', 'ESTREF': 'ISREF', 'ESTFORMULE': 'ISFORMULA',
  // finances
  'VPM': 'PMT', 'VA': 'PV', 'VC': 'FV', 'TAUX': 'RATE', 'NPM': 'NPER', 'VAN': 'NPV', 'PRINCPER': 'PPMT', 'INTPER': 'IPMT', 'AMORLIN': 'SLN',
  'TAUX.EFFECTIF': 'EFFECT', 'TAUX.NOMINAL': 'NOMINAL', 'CUMUL.INTER': 'CUMIPMT', 'CUMUL.PRINCPER': 'CUMPRINC', 'VAN.PAIEMENTS': 'XNPV'
};
// Noms secondaires : acceptés à la saisie, mais la formule s'affiche avec le nom principal.
const FR_ALIASES = new Set(['EQUATION.RANG', 'PREVISION.LINEAIRE', 'CENTILE.INCLURE', 'QUARTILE.INCLURE', 'MODE.SIMPLE']);
const EN2FR = Object.fromEntries(Object.entries(FR2EN).filter(([f]) => !FR_ALIASES.has(f)).map(([f, e]) => [e, f]));

function outsideStrings(f, fn) { return f.split('"').map((p, i) => i % 2 ? p : fn(p)).join('"'); }
const names = (f, dict) => outsideStrings(f.replace(/AUJOURD['’]HUI\s*\(/gi, 'AUJOURDHUI('), p =>
  p.replace(/([A-Za-zÀ-ÿ][A-Za-z0-9À-ÿ._]*)(?=\s*\()/g, m => dict[norm(m).toUpperCase()] || m));
const bools = f => outsideStrings(f, p => p.replace(/\b(VRAI|FAUX|TRUE|FALSE)\b(?!\s*\()/gi, m => (/^(VRAI|TRUE)$/i.test(m) ? 'TRUE' : 'FALSE') + '()'));
const frNamesToEn = f => names(bools(f), FR2EN);
const enNamesToFr = f => outsideStrings(names(f, EN2FR), p => p.replace(/\b(VRAI|FAUX)\(\)/g, '$1'));
const frToExcel = f => outsideStrings(frNamesToEn(f), p => p.replace(/(\d),(\d)/g, '$1.$2').replace(/;/g, ','));
const xlF = f => outsideStrings(frToExcel(f), p => p.replace(/\b(TRUE|FALSE)\(\)/g, '$1'));
const excelToFr = f => enNamesToFr(outsideStrings(f, p => p.replace(/,/g, ';').replace(/(\d)\.(\d)/g, '$1,$2')));

const COLN = c => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
const addr = (r, c) => COLN(c) + (r + 1);
const parseAddr = a => { const m = /^([A-Z]+)(\d+)$/.exec(a); if (!m) return null; let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { r: +m[2] - 1, c: c - 1 }; };
function parseNum(s) {
  let t = String(s).trim().replace(/[\s  ]/g, '');
  let pct = false;
  if (t.endsWith('%')) { pct = true; t = t.slice(0, -1); }
  if (t.endsWith('€')) t = t.slice(0, -1);
  if (!/^-?\d+([.,]\d+)?$/.test(t)) return null;
  let n = parseFloat(t.replace(',', '.'));
  return pct ? n / 100 : n;
}
const EPOCH = Date.UTC(1899, 11, 30);
const toSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / 86400000;
function parseDate(s) {
  const t = String(s).trim();
  let m = /^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2}|\d{4}))?$/.exec(t);
  if (m) {
    const d = +m[1], mo = +m[2]; let y = m[3] ? +m[3] : new Date().getFullYear();
    if (y < 100) y += y < 70 ? 2000 : 1900;
    if (mo < 1 || mo > 12 || d < 1) return null;
    if (new Date(Date.UTC(y, mo - 1, d)).getUTCDate() !== d) return null;
    return toSerial(y, mo, d);
  }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  return m ? toSerial(+m[1], +m[2], +m[3]) : null;
}
function fmtDate(n) { const dt = new Date(EPOCH + Math.floor(n) * 86400000), p = x => String(x).padStart(2, '0'); return `${p(dt.getUTCDate())}/${p(dt.getUTCMonth() + 1)}/${dt.getUTCFullYear()}`; }
const DATE_FN = /(AUJOURD'?HUI|MAINTENANT|\bDATE|FIN\.MOIS|MOIS\.DECALER|SERIE\.JOUR\.OUVRE|DATEVAL)\s*\(/i;
const numToRaw = n => String(+n.toPrecision(15)).replace('.', ',');
const NF = {
  '': new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 6 }),
  dec2: new Intl.NumberFormat('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
  eur: new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }),
  pct: new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 2 })
};

const SH = { keyTap: 0, caret: null, tapAgain: false, hf: null, sid: 0, fr: true, vals: [], cell: [], colTh: [], rowTh: [], sel: { r: 0, c: 0 }, anc: { r: 0, c: 0 }, editing: null, hist: [], fut: [], drag: null, refIns: null, hl: { r: [], c: [] } };
const gw = $('#gridWrap'), cellEd = $('#cellEd'), fx = $('#fx');

// ---------- Fonctions récentes d'Excel ajoutées au moteur de calcul ----------
function registerExtraFunctions(M) {
  if (!M || !M.FunctionPlugin || M._grimoireFns) return;
  M._grimoireFns = true;
  const { FunctionPlugin, FunctionArgumentType: T, ArraySize, SimpleRangeValue, CellError, EmptyValue } = M;
  const NA = () => new CellError(M.ErrorType.NA);
  const VAL = () => new CellError(M.ErrorType.VALUE);
  const isEmpty = v => v === EmptyValue || v === null || v === undefined || v === '';
  const isErr = v => v instanceof CellError;
  const grid = v => v instanceof SimpleRangeValue ? v.data : [[v]];
  const flat = v => grid(v).flat();
  const rawNum = v => (v && typeof v === 'object' && 'val' in v) ? v.val : v; // nombres enrichis (dates, %, €)
  const cmp = (a, b) => {
    a = rawNum(a); b = rawNum(b);
    const ra = isEmpty(a) ? 3 : typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : 2;
    const rb = isEmpty(b) ? 3 : typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : 2;
    if (ra !== rb) return ra - rb;
    if (ra === 0) return a - b;
    if (ra === 1) return a.localeCompare(b, 'fr', { sensitivity: 'base', numeric: true });
    if (ra === 2) return (a === true) - (b === true);
    return 0;
  };
  const same = (a, b) => {
    a = rawNum(a); b = rawNum(b);
    if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
    return a === b || (isEmpty(a) && isEmpty(b));
  };
  const wild = p => new RegExp('^' + String(p).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i');
  // critère « >5 », « <>Courses », « Co* »…
  const criterion = c => {
    c = rawNum(c);
    if (typeof c === 'number' || typeof c === 'boolean') return v => same(v, c);
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(String(c));
    const op = m[1] || '=', rest = m[2];
    const n = rest.trim() !== '' && !isNaN(+rest.replace(',', '.')) ? +rest.replace(',', '.') : null;
    if (op === '=' || op === '<>') {
      const test = n != null ? (v => rawNum(v) === n || (typeof rawNum(v) === 'string' && rawNum(v) === rest))
        : rest === '' ? (v => isEmpty(v)) : /[*?]/.test(rest) ? (re => v => typeof rawNum(v) === 'string' && re.test(rawNum(v)))(wild(rest)) : (v => typeof rawNum(v) === 'string' && rawNum(v).toLowerCase() === rest.toLowerCase());
      return op === '=' ? test : v => !test(v);
    }
    return v => {
      v = rawNum(v);
      if (n != null) { if (typeof v !== 'number') return false; return op === '<' ? v < n : op === '>' ? v > n : op === '<=' ? v <= n : v >= n; }
      if (typeof v !== 'string') return false;
      const d = v.localeCompare(rest, 'fr', { sensitivity: 'base' });
      return op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0;
    };
  };
  const sizeOf = (plugin, ast, state, i) => ast.args[i] ? plugin.arraySizeForAst(ast.args[i], state) : ArraySize.scalar();
  const numArg = (plugin, ast, i, dflt) => { const a = ast.args[i]; if (!a) return dflt; if (a.type === 'NUMBER') return a.value; if (a.type === 'MINUS_UNARY_OP' && a.value && a.value.type === 'NUMBER') return -a.value.value; return null; };

  class GrimoireFns extends FunctionPlugin {
    // RECHERCHEX(valeur; tableau_recherche; tableau_renvoyé; [si_non_trouvé]; [mode_correspondance]; [mode_recherche])
    xlookup(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('XLOOKUP'), (key, look, ret, notFound, mode = 0, search = 1) => {
        const L = look.data, R = ret.data;
        const vertical = L[0].length === 1;
        const list = vertical ? L.map(r => r[0]) : L[0];
        const i = findIndex(key, list, mode, search);
        if (i < 0) return notFound === undefined ? NA() : notFound;
        if (vertical) { if (R.length !== list.length) return VAL(); return R[i].length === 1 ? R[i][0] : SimpleRangeValue.onlyValues([R[i]]); }
        if (R[0].length !== list.length) return VAL();
        return R.length === 1 ? R[0][i] : SimpleRangeValue.onlyValues(R.map(r => [r[i]]));
      });
    }
    xlookupSize(ast, state) {
      const L = sizeOf(this, ast, state, 1), R = sizeOf(this, ast, state, 2);
      return L.width === 1 ? new ArraySize(R.width, 1) : new ArraySize(1, R.height);
    }
    // EQUIVX(valeur; tableau; [mode_correspondance]; [mode_recherche])
    xmatch(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('XMATCH'), (key, look, mode = 0, search = 1) => {
        const L = look.data; const list = L[0].length === 1 ? L.map(r => r[0]) : L[0];
        const i = findIndex(key, list, mode, search);
        return i < 0 ? NA() : i + 1;
      });
    }
    // UNIQUE(plage; [par_colonne]; [une_seule_fois])
    unique(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('UNIQUE'), (rng, byCol = false, once = false) => {
        let rows = rng.data;
        if (byCol) rows = rows[0].map((_, c) => rows.map(r => r[c]));
        const key = r => JSON.stringify(r.map(v => { v = rawNum(v); return typeof v === 'string' ? v.toLowerCase() : isEmpty(v) ? '' : v; }));
        const count = new Map(); rows.forEach(r => { const k = key(r); count.set(k, (count.get(k) || 0) + 1); });
        const seen = new Set(), out = [];
        rows.forEach(r => { const k = key(r); if (seen.has(k)) return; seen.add(k); if (once && count.get(k) > 1) return; out.push(r.map(v => isEmpty(v) ? 0 : v)); });
        if (!out.length) return new CellError(M.ErrorType.NA, 'Aucune valeur');
        const res = byCol ? out[0].map((_, i) => out.map(r => r[i])) : out;
        return SimpleRangeValue.onlyValues(res);
      });
    }
    uniqueSize(ast, state) { const z = sizeOf(this, ast, state, 0); return new ArraySize(z.width, z.height); }
    // TRIER(plage; [index_tri]; [ordre]; [par_colonne])
    sort(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('SORT'), (rng, idx = 1, order = 1, byCol = false) => {
        let rows = rng.data.map(r => r.slice());
        if (byCol) rows = rows[0].map((_, c) => rows.map(r => r[c]));
        const k = Math.trunc(idx) - 1;
        if (k < 0 || k >= rows[0].length) return VAL();
        rows.sort((a, b) => order < 0 ? -cmp(a[k], b[k]) : cmp(a[k], b[k]));
        const res = byCol ? rows[0].map((_, i) => rows.map(r => r[i])) : rows;
        return SimpleRangeValue.onlyValues(res.map(r => r.map(v => isEmpty(v) ? 0 : v)));
      });
    }
    sortSize(ast, state) { const z = sizeOf(this, ast, state, 0); return new ArraySize(z.width, z.height); }
    // TRIERPAR n'est pas couvert ; SEQUENCE(lignes; [colonnes]; [début]; [pas])
    sequence(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('SEQUENCE'), (rows, cols = 1, start = 1, step = 1) => {
        rows = Math.trunc(rows); cols = Math.trunc(cols);
        if (rows < 1 || cols < 1) return VAL();
        const out = [];
        for (let r = 0; r < rows; r++) { const row = []; for (let c = 0; c < cols; c++) row.push(start + step * (r * cols + c)); out.push(row); }
        return SimpleRangeValue.onlyValues(out);
      });
    }
    sequenceSize(ast) {
      const r = numArg(this, ast, 0, 1), c = numArg(this, ast, 1, 1);
      if (r == null || c == null) return ArraySize.error();
      return new ArraySize(Math.max(1, Math.trunc(c)), Math.max(1, Math.trunc(r)));
    }
    // JOINDRE.TEXTE(séparateur; ignorer_vides; texte1; …)
    textjoin(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('TEXTJOIN'), (sep, ignore, ...vals) => {
        const parts = [];
        for (const v of vals) for (const x of flat(v)) {
          if (isErr(x)) return x;
          if (ignore && isEmpty(x)) continue;
          parts.push(isEmpty(x) ? '' : txt(x));
        }
        return parts.join(sep);
      });
    }
    // CONCAT(texte1; …) : accepte des plages
    concat(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('CONCAT'), (...vals) => {
        let s = '';
        for (const v of vals) for (const x of flat(v)) { if (isErr(x)) return x; s += isEmpty(x) ? '' : txt(x); }
        return s;
      });
    }
    // MOYENNE.SI.ENS(plage_moyenne; plage_critère1; critère1; …)
    averageifs(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('AVERAGEIFS'), (avg, ...pairs) => {
        const A = avg.data; let sum = 0, n = 0;
        const tests = [];
        for (let i = 0; i < pairs.length; i += 2) {
          const rr = pairs[i]; if (!(rr instanceof SimpleRangeValue) || rr.height() !== A.length || rr.width() !== A[0].length) return VAL();
          tests.push([rr.data, criterion(pairs[i + 1])]);
        }
        for (let r = 0; r < A.length; r++) for (let c = 0; c < A[0].length; c++) {
          if (!tests.every(([d, t]) => t(d[r][c]))) continue;
          const v = rawNum(A[r][c]); if (typeof v === 'number') { sum += v; n++; }
        }
        return n ? sum / n : new CellError(M.ErrorType.DIV_BY_ZERO);
      });
    }
    // CNUM(texte) : « 1 234,50 », « 12 % », « 3 € », « 22/10/2026 »
    value(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('VALUE'), v => {
        v = rawNum(v);
        if (typeof v === 'number') return v;
        if (typeof v === 'boolean') return VAL();
        let t = String(v).trim().replace(/[\s  ]/g, '');
        if (t === '') return 0;
        let pct = false;
        if (t.endsWith('%')) { pct = true; t = t.slice(0, -1); }
        t = t.replace(/€$/, '');
        if (/^-?\d+([.,]\d+)?$/.test(t)) { const n = parseFloat(t.replace(',', '.')); return pct ? n / 100 : n; }
        const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(t);
        if (m) { let y = +m[3]; if (y < 100) y += y < 70 ? 2000 : 1900; return (Date.UTC(y, +m[2] - 1, +m[1]) - Date.UTC(1899, 11, 30)) / 86400000; }
        return VAL();
      });
    }
    // RANG(nombre; référence; [ordre])
    rank(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('RANK'), (n, ref, order = 0) => {
        const nums = nums_(ref); n = rawNum(n);
        if (!nums.includes(n)) return NA();
        return 1 + nums.filter(x => order ? x < n : x > n).length;
      });
    }
    // RECHERCHE(valeur; vecteur_recherche; [vecteur_résultat]) : plus grande valeur inférieure ou égale
    lookup(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('LOOKUP'), (key, look, res) => {
        const L = look.data, list = L[0].length === 1 ? L.map(r => r[0]) : L[0];
        let best = -1;
        for (let i = 0; i < list.length; i++) { if (isEmpty(list[i])) continue; const d = cmp(list[i], key); if (d <= 0 && typeof rawNum(list[i]) === typeof rawNum(key)) best = i; }
        if (best < 0) return NA();
        if (!res) return list[best];
        const R = res.data, out = R[0].length === 1 ? R.map(r => r[0]) : R[0];
        return best < out.length ? out[best] : NA();
      });
    }
    // CENTILE(plage; k) et QUARTILE(plage; quart)
    percentile(ast, state) { return this.runFunction(ast.args, state, this.metadata('PERCENTILE'), (rng, k) => pct_(nums_(rng), k)); }
    quartile(ast, state) { return this.runFunction(ast.args, state, this.metadata('QUARTILE'), (rng, q) => { q = Math.trunc(q); return q < 0 || q > 4 ? new CellError(M.ErrorType.NUM) : pct_(nums_(rng), q / 4); }); }
    // MODE(plage) : valeur la plus fréquente
    mode(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('MODE'), (...rs) => {
        const nums = rs.flatMap(nums_), cnt = new Map(); let best = null, bc = 1;
        nums.forEach(x => { const c = (cnt.get(x) || 0) + 1; cnt.set(x, c); if (c > bc) { bc = c; best = x; } });
        return best == null ? NA() : best;
      });
    }
    // PREVISION(x; y_connus; x_connus) et ORDONNEE.ORIGINE(y_connus; x_connus)
    forecast(ast, state) { return this.runFunction(ast.args, state, this.metadata('FORECAST'), (x, ys, xs) => { const f = lin_(ys, xs); return isErr(f) ? f : f.a + f.b * rawNum(x); }); }
    intercept(ast, state) { return this.runFunction(ast.args, state, this.metadata('INTERCEPT'), (ys, xs) => { const f = lin_(ys, xs); return isErr(f) ? f : f.a; }); }
  }
  const nums_ = v => flat(v).map(rawNum).filter(x => typeof x === 'number');
  const pct_ = (a, k) => {
    if (!a.length || k < 0 || k > 1) return new CellError(M.ErrorType.NUM);
    a = a.slice().sort((x, y) => x - y);
    const h = (a.length - 1) * k, lo = Math.floor(h);
    return lo + 1 < a.length ? a[lo] + (h - lo) * (a[lo + 1] - a[lo]) : a[lo];
  };
  const lin_ = (ys, xs) => {
    const Y = flat(ys).map(rawNum), X = flat(xs).map(rawNum);
    if (Y.length !== X.length) return NA();
    const P = Y.map((y, i) => [X[i], y]).filter(([x, y]) => typeof x === 'number' && typeof y === 'number');
    if (P.length < 2) return new CellError(M.ErrorType.DIV_BY_ZERO);
    const mx = P.reduce((s, p) => s + p[0], 0) / P.length, my = P.reduce((s, p) => s + p[1], 0) / P.length;
    let sxy = 0, sxx = 0; P.forEach(([x, y]) => { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; });
    if (!sxx) return new CellError(M.ErrorType.DIV_BY_ZERO);
    const b = sxy / sxx; return { a: my - b * mx, b };
  };
  const txt = x => { x = rawNum(x); return typeof x === 'boolean' ? (x ? 'VRAI' : 'FAUX') : typeof x === 'number' ? String(+x.toPrecision(15)).replace('.', ',') : String(x); };
  function findIndex(key, list, mode, search) {
    mode = Math.trunc(mode); search = Math.trunc(search);
    const idx = search < 0 ? list.map((_, i) => list.length - 1 - i) : list.map((_, i) => i);
    if (mode === 2) { const re = wild(txt(key)); for (const i of idx) if (typeof rawNum(list[i]) === 'string' && re.test(rawNum(list[i]))) return i; return -1; }
    for (const i of idx) if (same(list[i], key)) return i;
    if (mode === -1 || mode === 1) {
      let best = -1;
      for (const i of idx) {
        const d = cmp(list[i], key);
        if (rawNum(list[i]) == null || isEmpty(list[i]) || typeof rawNum(list[i]) !== typeof rawNum(key)) continue;
        if (mode === -1 && d < 0 && (best < 0 || cmp(list[i], list[best]) > 0)) best = i;
        if (mode === 1 && d > 0 && (best < 0 || cmp(list[i], list[best]) < 0)) best = i;
      }
      return best;
    }
    return -1;
  }
  const R = { argumentType: T.RANGE }, A = { argumentType: T.ANY }, N = (d) => ({ argumentType: T.NUMBER, ...(d !== undefined ? { defaultValue: d } : {}) }), B = (d) => ({ argumentType: T.BOOLEAN, defaultValue: d }), S = { argumentType: T.STRING };
  GrimoireFns.implementedFunctions = {
    XLOOKUP: { method: 'xlookup', arraySizeMethod: 'xlookupSize', parameters: [A, R, R, { argumentType: T.ANY, optionalArg: true }, N(0), N(1)], vectorizationForbidden: true },
    XMATCH: { method: 'xmatch', parameters: [A, R, N(0), N(1)], vectorizationForbidden: true },
    UNIQUE: { method: 'unique', arraySizeMethod: 'uniqueSize', arrayFunction: true, parameters: [R, B(false), B(false)], vectorizationForbidden: true },
    SORT: { method: 'sort', arraySizeMethod: 'sortSize', arrayFunction: true, parameters: [R, N(1), N(1), B(false)], vectorizationForbidden: true },
    SEQUENCE: { method: 'sequence', arraySizeMethod: 'sequenceSize', parameters: [N(), N(1), N(1), N(1)], vectorizationForbidden: true },
    TEXTJOIN: { method: 'textjoin', parameters: [S, B(true), A], repeatLastArgs: 1, expandRanges: false, vectorizationForbidden: true },
    CONCAT: { method: 'concat', parameters: [A], repeatLastArgs: 1, vectorizationForbidden: true },
    AVERAGEIFS: { method: 'averageifs', parameters: [R, R, A], repeatLastArgs: 2, vectorizationForbidden: true },
    VALUE: { method: 'value', parameters: [A] },
    RANK: { method: 'rank', parameters: [N(), R, N(0)], vectorizationForbidden: true },
    LOOKUP: { method: 'lookup', parameters: [A, R, { argumentType: T.RANGE, optionalArg: true }], vectorizationForbidden: true },
    PERCENTILE: { method: 'percentile', parameters: [R, N()], vectorizationForbidden: true },
    QUARTILE: { method: 'quartile', parameters: [R, N()], vectorizationForbidden: true },
    MODE: { method: 'mode', parameters: [R], repeatLastArgs: 1, vectorizationForbidden: true },
    FORECAST: { method: 'forecast', parameters: [N(), R, R], vectorizationForbidden: true },
    INTERCEPT: { method: 'intercept', parameters: [R, R], vectorizationForbidden: true }
  };
  const names = Object.fromEntries(Object.keys(GrimoireFns.implementedFunctions).map(k => [k, k]));
  try { M.HyperFormula.registerFunctionPlugin(GrimoireFns, { enGB: names, enUS: names }); }
  catch (e) { try { M.HyperFormula.registerFunctionPlugin(GrimoireFns, { enGB: names }); } catch (e2) { console.warn('Fonctions supplémentaires indisponibles', e2); } }
}

async function ensureHF() {
  if (SH.hf) return true;
  await loadScript(LIB.hf);
  const HFm = window.HyperFormula;
  registerExtraFunctions(HFm);
  try {
    SH.hf = HFm.buildEmpty({ licenseKey: 'gpl-v3', decimalSeparator: ',', thousandSeparator: '', functionArgSeparator: ';', arrayColumnSeparator: ';', arrayRowSeparator: '|' });
    SH.fr = true;
  } catch { SH.hf = HFm.buildEmpty({ licenseKey: 'gpl-v3' }); SH.fr = false; }
  SH.sid = SH.hf.getSheetId(SH.hf.addSheet('Feuille1'));
  return true;
}
function toHF(raw) {
  if (raw == null || raw === '') return null;
  const s = String(raw);
  if (s[0] === '=') return '=' + (SH.fr ? frNamesToEn(s.slice(1)) : frToExcel(s.slice(1)));
  if (s[0] === "'") return s.slice(1);
  if (/^(VRAI|FAUX)$/i.test(s.trim())) return /^VRAI$/i.test(s.trim());
  const n = parseNum(s);
  if (n != null) return n;
  const d = parseDate(s);
  return d == null ? s : d;
}
const cellOf = (r, c) => state.cur.cells[addr(r, c)];
const rawOf = (r, c) => { const x = cellOf(r, c); return x ? x.v : ''; };

// Un classeur contient plusieurs feuilles ; cells, rows, cols… désignent toujours la feuille active.
const SHEET_KEYS = ['cells', 'rows', 'cols', 'colW', 'filter', 'frz', 'hid', 'cf', 'charts', 'pivot'];
function ensureSheets(doc) {
  if (!Array.isArray(doc.sheets) || !doc.sheets.length) {
    doc.sheets = [{ id: uid(), name: 'Feuille1', cells: doc.cells || {}, rows: doc.rows || 60, cols: doc.cols || 12, colW: doc.colW || {}, filter: doc.filter, frz: doc.frz }];
    doc.active = 0;
  }
  if (!(doc.active >= 0 && doc.active < doc.sheets.length)) doc.active = 0;
  doc.sheets.forEach((sh, i) => { sh.cells ||= {}; sh.rows ||= 60; sh.cols ||= 12; sh.colW ||= {}; sh.id ||= uid(); sh.name ||= 'Feuille' + (i + 1); sh.hid ||= { r: [], c: [] }; sh.hid.r ||= []; sh.hid.c ||= []; sh.cf ||= []; sh.charts ||= []; });
  doc.names ||= {};
  SHEET_KEYS.forEach(k => {
    const d = Object.getOwnPropertyDescriptor(doc, k);
    if (d && d.get) return;
    delete doc[k];
    Object.defineProperty(doc, k, { configurable: true, enumerable: false, get() { return this.sheets[this.active][k]; }, set(v) { this.sheets[this.active][k] = v; } });
  });
}
const refName = n => /^[A-Za-zÀ-ÿ_][A-Za-z0-9À-ÿ_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;
function sheetData(sh) {
  const data = [];
  for (let r = 0; r < sh.rows; r++) { const row = []; for (let c = 0; c < sh.cols; c++) { const x = sh.cells[addr(r, c)]; row.push(x ? toHF(x.v) : null); } data.push(row); }
  return data;
}
function serToRaw(v) {
  let raw = typeof v === 'number' ? numToRaw(v) : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v);
  if (raw[0] === '=') raw = '=' + (SH.fr ? enNamesToFr(raw.slice(1)) : excelToFr(raw.slice(1)));
  return raw;
}
// Une date reste écrite en date (22/10/2026) même quand le moteur la renvoie en nombre.
function fixDates(cells) {
  Object.values(cells).forEach(x => { if (x.f && x.f.n === 'date' && /^-?\d+(,\d+)?$/.test(String(x.v))) x.v = fmtDate(parseNum(x.v)); });
  return cells;
}
function markDates(sh) {
  Object.values(sh.cells).forEach(x => { const v = String(x.v || ''); if (v && v[0] !== '=' && !(x.f && x.f.n) && parseNum(v) == null && parseDate(v) != null) (x.f ||= {}).n = 'date'; });
}
function cellsFromHF(id, old, fmts) {
  const cells = {};
  (SH.hf.getSheetSerialized(id) || []).forEach((row, ri) => (row || []).forEach((v, ci) => {
    if (v == null || v === '') return;
    const k = addr(ri, ci);
    cells[k] = { v: serToRaw(v), f: fmts ? {} : clone((old[k] && old[k].f) || {}) };
  }));
  if (!fmts) Object.entries(old).forEach(([k, x]) => { if (!cells[k] && x.f && Object.keys(x.f).length) cells[k] = { v: '', f: clone(x.f) }; });
  return fmts ? cells : fixDates(cells);
}
async function openSheet() {
  const cur = state.cur;
  ensureSheets(cur);
  cur.sheets.forEach(markDates);
  SH.hist = []; SH.fut = []; SH.sel = { r: 0, c: 0 }; SH.anc = { r: 0, c: 0 }; SH.editing = null;
  show('sheet');
  buildGrid(); renderTabs();
  $('#sStatus').textContent = 'Chargement du moteur de calcul…';
  try { await ensureHF(); }
  catch { toast('Moteur de calcul injoignable : les formules s\'afficheront sans être calculées. Vérifiez la connexion.', true); }
  if (state.cur !== cur) return;
  loadHF(); renderValues(); SH.hiddenRows = 0; applyFilters(); applyFreeze(); renderTabs(); select(0, 0);
  gw.focus({ preventScroll: true });
}
function loadHF() {
  if (!SH.hf) return;
  const cur = state.cur, hf = SH.hf;
  hf.getSheetNames().forEach(n => hf.removeSheet(hf.getSheetId(n)));
  SH.ids = cur.sheets.map(sh => hf.getSheetId(hf.addSheet(sh.name)));
  loadNames();
  hf.batch(() => cur.sheets.forEach((sh, i) => hf.setSheetContent(SH.ids[i], sheetData(sh))));
  SH.sid = SH.ids[cur.active];
}
function buildGrid() {
  const cur = state.cur;
  let h = '<colgroup><col class="rh">';
  let total = 46;
  const hc = new Set((cur.hid && cur.hid.c) || []);
  for (let c = 0; c < cur.cols; c++) { const w = hc.has(c) ? 6 : cur.colW[c] || 104; total += w; h += `<col style="width:${w}px">`; }
  h += '</colgroup><thead><tr><th class="corner" scope="col"><span class="sr">Lignes et colonnes</span></th>';
  for (let c = 0; c < cur.cols; c++) h += hc.has(c) ? `<th class="ch hc" scope="col" data-c="${c}" title="Colonne ${COLN(c)} masquée : double-cliquez pour l'afficher"></th>` : `<th class="ch" scope="col" data-c="${c}">${COLN(c)}<span class="rz" data-c="${c}"></span></th>`;
  h += '</tr></thead><tbody>';
  for (let r = 0; r < cur.rows; r++) {
    h += `<tr><th scope="row" data-r="${r}">${r + 1}</th>`;
    for (let c = 0; c < cur.cols; c++) h += hc.has(c) ? `<td data-r="${r}" data-c="${c}" class="hc"></td>` : `<td data-r="${r}" data-c="${c}"></td>`;
    h += '</tr>';
  }
  const t = $('#grid'); t.innerHTML = h + '</tbody>'; t.style.width = total + 'px';
  SH.cell = []; $$('tbody tr', t).forEach((tr, r) => { SH.cell[r] = [...tr.children].slice(1); });
  SH.colTh = $$('thead th.ch', t); SH.rowTh = $$('tbody th', t); SH.hl = { r: [], c: [] };
  SH.mgSig = null; applyMerges();
}
const ERR_FR = { '#NAME?': '#NOM?', '#VALUE!': '#VALEUR!', '#NUM!': '#NOMBRE!', '#SPILL!': '#PROPAGATION!', '#ERROR!': '#ERREUR!', '#DIV/0!': '#DIV/0!', '#N/A': '#N/A', '#REF!': '#REF!', '#CYCLE!': '#CYCLE!' };
const ERR_WHY = { NAME: 'nom de fonction ou de plage inconnu', VALUE: 'type de valeur incorrect (texte à la place d\'un nombre ?)', NUM: 'nombre impossible', SPILL: 'le résultat déborde sur des cellules déjà remplies', DIV_BY_ZERO: 'division par zéro', NA: 'valeur introuvable', REF: 'référence vers une cellule supprimée', CYCLE: 'la formule fait référence à elle-même', ERROR: 'formule incorrecte' };
const NF_CACHE = {};
function numFmt(n, d) {
  const k = (n || '') + '|' + (d ?? '');
  if (NF_CACHE[k]) return NF_CACHE[k];
  const o = n === 'eur' ? { style: 'currency', currency: 'EUR' } : n === 'pct' ? { style: 'percent', maximumFractionDigits: 2 } : n === 'dec2' ? { minimumFractionDigits: 2, maximumFractionDigits: 2 } : { maximumFractionDigits: 6 };
  if (d != null) { o.minimumFractionDigits = d; o.maximumFractionDigits = d; }
  return (NF_CACHE[k] = new Intl.NumberFormat('fr-FR', o));
}
const GRID_FONTS = ['Instrument Sans', 'Arial', 'Calibri', 'Georgia', 'Times New Roman', 'Verdana', 'Courier New'];
const GRID_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
function display(r, c) {
  const x = cellOf(r, c), f = (x && x.f) || {};
  let v = SH.hf ? (SH.vals[r] || [])[c] : (x ? (parseNum(x.v) ?? x.v) : null);
  if (!SH.hf && x && String(x.v)[0] === '=') v = x.v;
  let txt = '', cls = '';
  if (f.chk) { cls = 'chk' + (v === true || v === 1 ? ' on' : ''); }
  else if (v == null || v === '') txt = '';
  else if (typeof v === 'object') { txt = ERR_FR[v.value] || v.value || '#ERREUR'; cls = 'err'; }
  else if (typeof v === 'number') { txt = f.n === 'date' ? fmtDate(v) : numFmt(f.n, f.d).format(v); cls = 'num'; }
  else if (typeof v === 'boolean') { txt = v ? 'VRAI' : 'FAUX'; cls = 'ac'; }
  else txt = String(v);
  const cf = SH.cfMap && SH.cfMap[r + ',' + c];
  if (f.b || (cf && cf.b)) cls += ' b';
  if (f.i) cls += ' i';
  if (f.u) cls += ' u';
  if (f.w) cls += ' wr';
  if (f.a) cls += ' ' + f.a;
  if (f.dd) cls += ' dd';
  if (f.note) cls += ' nt';
  if (f.mg) cls += ' mg';
  const sh = [];
  if (f.bt) sh.push('inset 0 1px 0 var(--bd)'); if (f.bb) sh.push('inset 0 -1px 0 var(--bd)');
  if (f.bl) sh.push('inset 1px 0 0 var(--bd)'); if (f.br) sh.push('inset -1px 0 0 var(--bd)');
  const bg = cf && cf.bg ? cf.bg : f.bg ? `var(--bg-${f.bg})` : '';
  const fc = cf && cf.fc ? cf.fc : f.fc ? `var(--fc-${f.fc})` : '';
  return { txt, cls: cls.trim(), v, bg, fc, bd: sh.join(','), ff: f.ff && FONT_STACK[f.ff] ? FONT_STACK[f.ff] : '', fs: f.fs ? f.fs + 'pt' : '' };
}
function renderValues() {
  if (!state.cur || state.view !== 'sheet') return;
  SH.vals = SH.hf ? SH.hf.getSheetValues(SH.sid) : [];
  const cur = state.cur;
  computeCF();
  applyMerges();
  for (let r = 0; r < cur.rows; r++) for (let c = 0; c < cur.cols; c++) {
    const td = SH.cell[r][c], d = display(r, c);
    if (td.textContent !== d.txt) td.textContent = d.txt;
    const keep = [...td.classList].filter(k => /^(fz|hc$|mgh$)/.test(k)).join(' ');
    const want = (d.cls + (keep ? ' ' + keep : '')).trim();
    if (td.className !== want) td.className = want;
    const st = d.bg + '|' + d.fc + '|' + d.bd + '|' + d.ff + '|' + d.fs;
    if (td._st !== st) { td._st = st; td.style.backgroundColor = d.bg; td.style.color = d.fc; td.style.boxShadow = d.bd; td.style.fontFamily = d.ff; td.style.fontSize = d.fs; }
    const x = cellOf(r, c), note = x && x.f && x.f.note;
    const tip = note ? 'Note : ' + note + (d.txt.length > 12 ? '\n' + d.txt : '') : d.txt.length > 12 ? d.txt : '';
    if (tip) { if (td.title !== tip) td.title = tip; } else if (td.title) td.removeAttribute('title');
  }
  if (cur.frz) applyFreeze();
  renderCharts();
}
function applyFreeze() {
  if (!state.cur || state.view !== 'sheet') return;
  const F = state.cur.frz || { r: 0, c: 0 }, t = $('#grid');
  $$('#grid .fz, #grid .fzc').forEach(el => { el.classList.remove('fz', 'fzc', 'fzl', 'fzr'); el.style.top = ''; el.style.left = ''; });
  let top = t.tHead.offsetHeight;
  for (let r = 0; r < Math.min(F.r || 0, state.cur.rows); r++) {
    const tr = SH.cell[r][0].parentElement, th = tr.firstElementChild;
    th.classList.add('fz'); th.style.top = top + 'px';
    SH.cell[r].forEach(td => { td.classList.add('fz'); td.style.top = top + 'px'; if (r === F.r - 1) td.classList.add('fzl'); });
    top += tr.offsetHeight;
  }
  let left = 46;
  for (let c = 0; c < Math.min(F.c || 0, state.cur.cols); c++) {
    const w = state.cur.colW[c] || 104;
    SH.cell.forEach(row => { const td = row[c]; td.classList.add('fzc'); td.style.left = left + 'px'; if (c === F.c - 1) td.classList.add('fzr'); });
    SH.colTh[c].classList.add('fzc'); SH.colTh[c].style.left = left + 'px';
    left += w;
  }
}
const rng = () => ({ r1: Math.min(SH.anc.r, SH.sel.r), r2: Math.max(SH.anc.r, SH.sel.r), c1: Math.min(SH.anc.c, SH.sel.c), c2: Math.max(SH.anc.c, SH.sel.c) });
const rngRef = R => R.r1 === R.r2 && R.c1 === R.c2 ? addr(R.r1, R.c1) : addr(R.r1, R.c1) + ':' + addr(R.r2, R.c2);
function box(el, a, b) {
  const t = $('#grid');
  const l = t.offsetLeft + a.offsetLeft, tp = t.offsetTop + a.offsetTop;
  el.style.left = (l - 1) + 'px'; el.style.top = (tp - 1) + 'px';
  el.style.width = (b.offsetLeft + b.offsetWidth - a.offsetLeft + 1) + 'px';
  el.style.height = (b.offsetTop + b.offsetHeight - a.offsetTop + 1) + 'px';
}
function drawSel() {
  const R = expandMerged(rng());
  const one = SH.merges.find(m => m.r1 === R.r1 && m.c1 === R.c1 && m.r2 === R.r2 && m.c2 === R.c2);
  box($('#selBox'), tdOf(R.r1, R.c1), tdOf(R.r2, R.c2));
  $('#selBox').style.display = (R.r1 === R.r2 && R.c1 === R.c2) || one ? 'none' : '';
  box($('#actBox'), tdOf(SH.sel.r, SH.sel.c), tdOf(SH.sel.r, SH.sel.c));
  SH.hl.r.forEach(e => e.classList.remove('hl')); SH.hl.c.forEach(e => e.classList.remove('hl'));
  SH.hl.r = SH.rowTh.slice(R.r1, R.r2 + 1); SH.hl.c = SH.colTh.slice(R.c1, R.c2 + 1);
  SH.hl.r.forEach(e => e.classList.add('hl')); SH.hl.c.forEach(e => e.classList.add('hl'));
  const brc = tdOf(R.r2, R.c2), tb = $('#grid'), half = isPhone() ? 10 : 5;
  fillH.style.left = (tb.offsetLeft + brc.offsetLeft + brc.offsetWidth - half) + 'px';
  fillH.style.top = (tb.offsetTop + brc.offsetTop + brc.offsetHeight - half) + 'px';
}
function scrollToCell(r, c) {
  const td = tdOf(r, c), t = $('#grid');
  const top = t.offsetTop + td.offsetTop, left = t.offsetLeft + td.offsetLeft;
  const hh = t.tHead.offsetHeight, rw = 46;
  if (top - hh < gw.scrollTop) gw.scrollTop = top - hh;
  else if (top + td.offsetHeight > gw.scrollTop + gw.clientHeight) gw.scrollTop = top + td.offsetHeight - gw.clientHeight;
  if (left - rw < gw.scrollLeft) gw.scrollLeft = left - rw;
  else if (left + td.offsetWidth > gw.scrollLeft + gw.clientWidth) gw.scrollLeft = left + td.offsetWidth - gw.clientWidth;
}
function select(r, c, extend) {
  const cur = state.cur;
  r = Math.max(0, Math.min(cur.rows - 1, r)); c = Math.max(0, Math.min(cur.cols - 1, c));
  if (rowHidden(r)) { const dir = r >= SH.sel.r ? 1 : -1; let rr = r; while (rr >= 0 && rr < cur.rows && rowHidden(rr)) rr += dir; r = rr >= 0 && rr < cur.rows ? rr : SH.sel.r; }
  if (colHidden(c)) { const dir = c >= SH.sel.c ? 1 : -1; let cc = c; while (cc >= 0 && cc < cur.cols && colHidden(cc)) cc += dir; c = cc >= 0 && cc < cur.cols ? cc : SH.sel.c; }
  const cov = SH.cover && SH.cover[r + ',' + c];
  if (cov) { r = cov.r; c = cov.c; }
  SH.sel = { r, c }; if (!extend) SH.anc = { r, c };
  drawSel(); scrollToCell(r, c); updateFx(); updateSelStatus(); updateSheetTools();
  if (!$('#plume').hidden && Plume.mode === 'feuille') Plume.ctx();
}
function updateFx() {
  if (SH.editing) return;
  if (document.activeElement !== $('#fxRef')) $('#fxRef').value = rngRef(rng());
  fx.value = rawOf(SH.sel.r, SH.sel.c);
}
function updateSelStatus() {
  const R = rng(); let n = 0, sum = 0, cnt = 0;
  for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) {
    const v = (SH.vals[r] || [])[c];
    if (v != null && v !== '') cnt++;
    if (typeof v === 'number') { n++; sum += v; }
  }
  const multi = R.r1 !== R.r2 || R.c1 !== R.c2;
  const raw = rawOf(SH.sel.r, SH.sel.c);
  const left = multi
    ? (n ? `Somme <b>${NF.dec2.format(sum)}</b> · Moyenne <b>${NF.dec2.format(sum / n)}</b> · Nombre <b>${cnt}</b>` : `<b>${cnt}</b> cellule${cnt > 1 ? 's' : ''} remplie${cnt > 1 ? 's' : ''}`)
    : (String(raw)[0] === '=' ? `Formule en <b>${addr(SH.sel.r, SH.sel.c)}</b>` : `Cellule <b>${addr(SH.sel.r, SH.sel.c)}</b>`);
  const hint = !SH.hf ? 'Calcul indisponible' : isPhone() ? 'Touchez à nouveau une cellule sélectionnée pour écrire' : 'Tapez directement · Entrée valide · F2 modifie';
  const nx = cellOf(SH.sel.r, SH.sel.c), note = !multi && nx && nx.f && nx.f.note;
  const errV = !multi && (SH.vals[SH.sel.r] || [])[SH.sel.c];
  const why = errV && typeof errV === 'object' && ERR_WHY[errV.type] ? ` · <span class="err-why">${esc(ERR_FR[errV.value] || errV.value)} : ${ERR_WHY[errV.type]}</span>` : '';
  const flt = SH.hiddenRows ? ` <button class="linkbtn" data-act="clearFilter">Filtre : ${SH.hiddenRows} ligne${SH.hiddenRows > 1 ? 's' : ''} masquée${SH.hiddenRows > 1 ? 's' : ''} · tout afficher</button>` : '';
  $('#sStatus').innerHTML = `<span>${left}${why}${note ? ` · <button class="linkbtn" data-act="note">Note : ${esc(note.length > 60 ? note.slice(0, 59) + '…' : note)}</button>` : ''}${flt}</span><span>${hint}</span>`;
}
function updateSheetTools() {
  const f = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {};
  $$('[data-s="bold"]').forEach(b => b.setAttribute('aria-pressed', String(!!f.b)));
  [['ital', 'i'], ['under', 'u'], ['wrap', 'w']].forEach(([s, k]) => { const b = $(`[data-s="${s}"]`); if (b) b.setAttribute('aria-pressed', String(!!f[k])); });
  const fz = state.cur.frz, fzb = $('[data-s="freeze"]'); if (fzb) fzb.setAttribute('aria-pressed', String(!!(fz && (fz.r || fz.c))));
  ['al', 'ac', 'ar'].forEach(k => $(`[data-s="${k}"]`).setAttribute('aria-pressed', String(f.a === k)));
  $$('.seg [data-n]').forEach(b => b.setAttribute('aria-pressed', String((f.n || '') === b.dataset.n)));
  $('#sFont').value = f.ff || ''; $('#sSize').value = f.fs ? String(f.fs) : '';
  $('#sNum').value = f.n || '';
  const mb = $('[data-s="merge"]'); if (mb) mb.setAttribute('aria-pressed', String(!!f.mg));
}

/* --- fusion, lignes et colonnes masquées, mise en forme conditionnelle */
const colHidden = c => !!(state.cur.hid && state.cur.hid.c.includes(c));
SH.merges = []; SH.cover = {};
function tdOf(r, c) { const cv = SH.cover[r + ',' + c]; return cv ? SH.cell[cv.r][cv.c] : SH.cell[r][c]; }
function applyMerges() {
  const cur = state.cur; if (!cur || !SH.cell.length) return;
  const list = [];
  Object.entries(cur.cells).forEach(([k, x]) => {
    const m = x.f && x.f.mg; if (!m) return;
    const p = parseAddr(k); if (!p || p.r >= cur.rows || p.c >= cur.cols) return;
    const r2 = Math.min(cur.rows - 1, p.r + (m.r || 1) - 1), c2 = Math.min(cur.cols - 1, p.c + (m.c || 1) - 1);
    if (r2 === p.r && c2 === p.c) return;
    list.push({ r1: p.r, c1: p.c, r2, c2 });
  });
  const sig = JSON.stringify(list);
  if (sig === SH.mgSig) return;
  SH.mgSig = sig; SH.merges = list; SH.cover = {};
  $$('#grid td.mgh').forEach(td => td.classList.remove('mgh'));
  $$('#grid td[colspan], #grid td[rowspan]').forEach(td => { td.removeAttribute('colspan'); td.removeAttribute('rowspan'); });
  list.forEach(m => {
    const a = SH.cell[m.r1][m.c1];
    a.colSpan = m.c2 - m.c1 + 1; a.rowSpan = m.r2 - m.r1 + 1;
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) { SH.cover[r + ',' + c] = { r: m.r1, c: m.c1 }; SH.cell[r][c].classList.add('mgh'); }
  });
}
function expandMerged(R) {
  R = { ...R };
  for (let changed = true; changed;) {
    changed = false;
    for (const m of SH.merges) {
      if (m.r2 < R.r1 || m.r1 > R.r2 || m.c2 < R.c1 || m.c1 > R.c2) continue;
      if (m.r1 < R.r1) { R.r1 = m.r1; changed = true; } if (m.r2 > R.r2) { R.r2 = m.r2; changed = true; }
      if (m.c1 < R.c1) { R.c1 = m.c1; changed = true; } if (m.c2 > R.c2) { R.c2 = m.c2; changed = true; }
    }
  }
  return R;
}
function toggleMerge() {
  commitEdit(false);
  const R = expandMerged(rng());
  const one = SH.merges.find(m => m.r1 === R.r1 && m.c1 === R.c1 && m.r2 === R.r2 && m.c2 === R.c2);
  pushHist();
  const cells = state.cur.cells;
  if (one) { const x = cells[addr(R.r1, R.c1)]; if (x && x.f) delete x.f.mg; afterChange(); drawSel(); return toast('Cellules séparées'); }
  if (R.r1 === R.r2 && R.c1 === R.c2) { SH.hist.pop(); return toast('Sélectionnez plusieurs cellules à fusionner.'); }
  let lost = false;
  batch(() => { for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) {
    const x = cells[addr(r, c)]; if (!x) continue;
    if (x.f) delete x.f.mg;
    if ((r !== R.r1 || c !== R.c1) && x.v !== '' && x.v != null) { lost = true; applyRaw(r, c, ''); }
  } });
  const k = addr(R.r1, R.c1), a = cells[k] || (cells[k] = { v: '', f: {} }); a.f ||= {};
  a.f.mg = { r: R.r2 - R.r1 + 1, c: R.c2 - R.c1 + 1 };
  if (!a.f.a && R.c2 > R.c1) a.f.a = 'ac';
  SH.anc = { r: R.r1, c: R.c1 }; SH.sel = { r: R.r1, c: R.c1 };
  afterChange(); drawSel();
  toast(lost ? 'Cellules fusionnées : seule la valeur en haut à gauche est gardée.' : 'Cellules fusionnées');
}
function setHidden(kind, on) {
  commitEdit(false);
  const cur = state.cur, R = rng(); pushHist();
  cur.hid ||= { r: [], c: [] };
  if (kind === 'r') {
    if (on) { const set = new Set(cur.hid.r); for (let r = R.r1; r <= R.r2; r++) set.add(r); if (set.size >= cur.rows) { SH.hist.pop(); return toast('Impossible de masquer toutes les lignes.'); } cur.hid.r = [...set].sort((a, b) => a - b); }
    else cur.hid.r = [];
    applyFilters();
    if (on) select(Math.min(R.r2 + 1, cur.rows - 1), SH.sel.c);
  } else {
    if (on) { const set = new Set(cur.hid.c); for (let c = R.c1; c <= R.c2; c++) set.add(c); if (set.size >= cur.cols) { SH.hist.pop(); return toast('Impossible de masquer toutes les colonnes.'); } cur.hid.c = [...set].sort((a, b) => a - b); }
    else cur.hid.c = [];
    buildGrid(); renderValues(); applyFilters(); applyFreeze();
    if (on) select(SH.sel.r, Math.min(R.c2 + 1, cur.cols - 1)); else drawSel();
  }
  scheduleSave();
  toast(on ? (kind === 'r' ? 'Lignes masquées' : 'Colonnes masquées') : (kind === 'r' ? 'Toutes les lignes sont affichées' : 'Toutes les colonnes sont affichées'));
}
function unhideCol(c) { const cur = state.cur; pushHist(); cur.hid.c = cur.hid.c.filter(x => x !== c); buildGrid(); renderValues(); applyFilters(); applyFreeze(); select(SH.sel.r, c); scheduleSave(); }
function stepDecimals(delta) {
  commitEdit(false); pushHist();
  const first = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {};
  const v = (SH.vals[SH.sel.r] || [])[SH.sel.c];
  let base = first.d;
  if (base == null) {
    if (first.n === 'dec2' || first.n === 'eur') base = 2;
    else if (typeof v === 'number') { const t = String(+v.toPrecision(12)); base = first.n === 'pct' ? Math.max(0, (String(+(v * 100).toPrecision(12)).split('.')[1] || '').length) : (t.split('.')[1] || '').length; }
    else base = 0;
  }
  const d = Math.max(0, Math.min(10, base + delta));
  eachInRange((r, c) => { const k = addr(r, c), x = state.cur.cells[k] || (state.cur.cells[k] = { v: '', f: {} }); x.f ||= {}; x.f.d = d; });
  afterChange();
}

// Mise en forme conditionnelle : règles évaluées à chaque affichage
const CF_TYPES = [
  ['gt', 'Supérieur à', 1], ['ge', 'Supérieur ou égal à', 1], ['lt', 'Inférieur à', 1], ['le', 'Inférieur ou égal à', 1],
  ['eq', 'Égal à', 1], ['ne', 'Différent de', 1], ['between', 'Entre', 2], ['contains', 'Le texte contient', 1], ['notContains', 'Le texte ne contient pas', 1],
  ['begins', 'Le texte commence par', 1], ['empty', 'Est vide', 0], ['notEmpty', 'N\'est pas vide', 0], ['dup', 'Valeur en double', 0],
  ['above', 'Au-dessus de la moyenne', 0], ['below', 'En dessous de la moyenne', 0], ['formula', 'La formule est vraie', 1], ['scale', 'Échelle de couleurs', 0]
];
const CF_LABEL = Object.fromEntries(CF_TYPES.map(t => [t[0], t[1]]));
function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim() || '#ffffff'; }
function hexRgb(h) { h = h.replace('#', ''); if (h.length === 3) h = h.split('').map(x => x + x).join(''); return [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16)); }
function mix(a, b, t) { const A = hexRgb(a), B = hexRgb(b); return `rgb(${A.map((x, i) => Math.round(x + (B[i] - x) * t)).join(',')})`; }
const cfNum = s => { if (s == null || s === '') return null; const n = parseNum(String(s)); if (n != null) return n; const d = parseDate(String(s)); return d; };
function cfMatch(rule, v, txt, ctx, r, c) {
  const t = rule.t, isNum = typeof v === 'number';
  const a = cfNum(rule.a), b = cfNum(rule.b);
  const cmpTxt = () => String(txt).toLowerCase(), A = String(rule.a ?? '').toLowerCase();
  switch (t) {
    case 'gt': return isNum && a != null ? v > a : false;
    case 'ge': return isNum && a != null ? v >= a : false;
    case 'lt': return isNum && a != null ? v < a : false;
    case 'le': return isNum && a != null ? v <= a : false;
    case 'eq': return a != null && isNum ? v === a : txt !== '' && cmpTxt() === A;
    case 'ne': return a != null && isNum ? v !== a : cmpTxt() !== A;
    case 'between': return isNum && a != null && b != null ? v >= Math.min(a, b) && v <= Math.max(a, b) : false;
    case 'contains': return !!A && cmpTxt().includes(A);
    case 'notContains': return !A || !cmpTxt().includes(A);
    case 'begins': return !!A && cmpTxt().startsWith(A);
    case 'empty': return v == null || v === '';
    case 'notEmpty': return !(v == null || v === '');
    case 'dup': return txt !== '' && ctx.count.get(cmpTxt()) > 1;
    case 'above': return isNum && ctx.avg != null && v > ctx.avg;
    case 'below': return isNum && ctx.avg != null && v < ctx.avg;
    case 'formula': {
      if (!SH.hf || !rule.a) return false;
      try {
        const f = String(rule.a).replace(/^=/, '');
        const sh = shiftFormula(f, r - rule.r1, c - rule.c1);
        const res = SH.hf.calculateFormula('=' + (SH.fr ? frNamesToEn(sh) : frToExcel(sh)), SH.sid);
        return res === true || (typeof res === 'number' && res !== 0);
      } catch { return false; }
    }
  }
  return false;
}
function computeCF() {
  const cur = state.cur; SH.cfMap = {};
  const rules = cur.cf || []; if (!rules.length) return;
  rules.forEach(rule => {
    const r2 = Math.min(rule.r2, cur.rows - 1), c2 = Math.min(rule.c2, cur.cols - 1);
    if ((r2 - rule.r1 + 1) * (c2 - rule.c1 + 1) > 20000) return;
    const ctx = {};
    const cells = [];
    for (let r = rule.r1; r <= r2; r++) for (let c = rule.c1; c <= c2; c++) {
      const v = (SH.vals[r] || [])[c];
      const x = cellOf(r, c), f = (x && x.f) || {};
      const txt = v == null || v === '' ? '' : typeof v === 'number' ? (f.n === 'date' ? fmtDate(v) : String(v)) : typeof v === 'object' ? '' : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v);
      cells.push({ r, c, v, txt });
    }
    if (rule.t === 'dup') { ctx.count = new Map(); cells.forEach(q => { if (q.txt !== '') { const k = q.txt.toLowerCase(); ctx.count.set(k, (ctx.count.get(k) || 0) + 1); } }); }
    if (rule.t === 'above' || rule.t === 'below' || rule.t === 'scale') {
      const nums = cells.map(q => q.v).filter(v => typeof v === 'number');
      ctx.avg = nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : null;
      ctx.min = Math.min(...nums); ctx.max = Math.max(...nums);
    }
    if (rule.t === 'scale') {
      const paper = cssVar('--paper'), lo = rule.lo ? cssVar('--bg-' + rule.lo) : paper, hi = cssVar('--bg-' + (rule.hi || 'vert'));
      const mid = rule.mid ? cssVar('--bg-' + rule.mid) : null;
      cells.forEach(q => {
        if (typeof q.v !== 'number' || ctx.max === ctx.min) return;
        const k = q.r + ',' + q.c; if (SH.cfMap[k] && SH.cfMap[k].bg) return;
        const t = (q.v - ctx.min) / (ctx.max - ctx.min);
        const bg = mid ? (t < .5 ? mix(lo, mid, t * 2) : mix(mid, hi, (t - .5) * 2)) : mix(lo, hi, t);
        (SH.cfMap[k] ||= {}).bg = bg;
      });
      return;
    }
    const st = rule.st || {};
    cells.forEach(q => {
      if (!cfMatch(rule, q.v, q.txt, ctx, q.r, q.c)) return;
      const k = q.r + ',' + q.c, m = SH.cfMap[k] ||= {};
      if (st.bg && !m.bg) m.bg = `var(--bg-${st.bg})`;
      if (st.fc && !m.fc) m.fc = `var(--fc-${st.fc})`;
      if (st.b) m.b = 1;
    });
  });
}

/* --- historique */
const snap = () => JSON.stringify({ cells: state.cur.cells, rows: state.cur.rows, cols: state.cur.cols, colW: state.cur.colW, hid: state.cur.hid, cf: state.cur.cf, charts: state.cur.charts });
function pushHist() { SH.hist.push(snap()); if (SH.hist.length > 80) SH.hist.shift(); SH.fut = []; }
function restore(s) {
  const o = JSON.parse(s); const dims = o.rows !== state.cur.rows || o.cols !== state.cur.cols || JSON.stringify(o.colW) !== JSON.stringify(state.cur.colW) || JSON.stringify(o.hid) !== JSON.stringify(state.cur.hid);
  if (!o.hid) delete o.hid; if (!o.cf) delete o.cf; if (!o.charts) delete o.charts;
  Object.assign(state.cur, o);
  if (dims) buildGrid();
  loadHF(); afterChange(); applyFilters(); applyFreeze(); select(Math.min(SH.sel.r, state.cur.rows - 1), Math.min(SH.sel.c, state.cur.cols - 1));
}
function undo() { if (!SH.hist.length) return; SH.fut.push(snap()); restore(SH.hist.pop()); }
function redo() { if (!SH.fut.length) return; SH.hist.push(snap()); restore(SH.fut.pop()); }

/* --- écriture des cellules */
function ensureSize(r, c) {
  const cur = state.cur; let grow = false;
  if (r >= cur.rows) { cur.rows = Math.min(r + 20, 2000); grow = true; }
  if (c >= cur.cols) { cur.cols = Math.min(c + 2, 60); grow = true; }
  if (grow) { buildGrid(); }
  return grow;
}
// Saisie à la manière d'Excel : « +64*2 » ou « -A1*2 » deviennent des formules,
// et les guillemets typographiques du clavier (iPhone) redeviennent des guillemets droits.
function normRaw(raw) {
  let s = raw == null ? '' : String(raw);
  if (/^[+\-]/.test(s) && parseNum(s) == null && /[A-Za-z0-9(]/.test(s.slice(1))) s = '=' + (s[0] === '+' ? s.slice(1) : s);
  if (s[0] === '=') s = s.replace(/[“”„«»]/g, '"').replace(/[‘’]/g, "'").replace(/\u00a0/g, ' ');
  return s;
}
function applyRaw(r, c, raw) {
  const k = addr(r, c), cells = state.cur.cells;
  raw = normRaw(raw);
  const old = cells[k];
  if (raw === '') { if (old) { if (old.f && Object.keys(old.f).length) old.v = ''; else delete cells[k]; } }
  else {
    const cell = old || (cells[k] = { v: '', f: {} }); cell.f ||= {};
    cell.v = raw;
    if (raw[0] !== '=' && !cell.f.n) { const t = raw.trim(); if (t.endsWith('%') && parseNum(t) != null) cell.f.n = 'pct'; else if (t.endsWith('€') && parseNum(t) != null) cell.f.n = 'eur'; }
    if (!cell.f.n && (raw[0] === '=' ? DATE_FN.test(raw) : parseNum(raw) == null && parseDate(raw) != null)) cell.f.n = 'date';
  }
  if (SH.hf) SH.hf.setCellContents({ sheet: SH.sid, row: r, col: c }, [[toHF(raw)]]);
}
function afterChange() { renderValues(); updateFx(); updateSelStatus(); updateSheetTools(); scheduleSave(); }
function setRaw(r, c, raw) { pushHist(); applyRaw(r, c, raw); afterChange(); }
function batch(fn) { if (SH.hf) SH.hf.batch(fn); else fn(); }
function eachInRange(fn) { const R = rng(); for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) fn(r, c); }

function clearRange() { pushHist(); batch(() => eachInRange((r, c) => applyRaw(r, c, ''))); afterChange(); }
function setFmt(key, val) {
  commitEdit(false); pushHist();
  const first = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {};
  if (['b', 'i', 'u', 'w'].includes(key)) val = first[key] ? 0 : 1;
  if (key === 'a' && first.a === val) val = '';
  eachInRange((r, c) => {
    const k = addr(r, c), x = state.cur.cells[k] || (state.cur.cells[k] = { v: '', f: {} }); x.f ||= {};
    if (val) x.f[key] = val; else delete x.f[key];
    if (!x.v && !Object.keys(x.f).length) delete state.cur.cells[k];
  });
  afterChange();
}

/* --- édition */
const isFormulaMode = () => {
  const inp = SH.editing === 'fx' ? fx : cellEd;
  if (!SH.editing || inp.value[0] !== '=') return false;
  const before = inp.value.slice(0, inp.selectionStart);
  return /[=(;+\-*/&<>:^ ]$/.test(before) || !!SH.refIns;
};
function placeEditor() {
  const td = tdOf(SH.sel.r, SH.sel.c), t = $('#grid');
  cellEd.style.left = (t.offsetLeft + td.offsetLeft - 1) + 'px';
  cellEd.style.top = (t.offsetTop + td.offsetTop - 1) + 'px';
  cellEd.style.minWidth = (td.offsetWidth + 2) + 'px';
  cellEd.style.height = (td.offsetHeight + 2) + 'px';
  growEditor();
}
function growEditor() {
  const m = document.createElement('span'); m.style.cssText = 'position:absolute;visibility:hidden;white-space:pre;font:' + getComputedStyle(cellEd).font;
  m.textContent = cellEd.value + '  '; document.body.appendChild(m);
  cellEd.style.width = Math.min(Math.max(m.offsetWidth + 18, parseFloat(cellEd.style.minWidth) || 0), 520) + 'px'; m.remove();
  cellEd.classList.toggle('fm', cellEd.value[0] === '=');
}
function startEdit(initial, keep) {
  SH.refIns = null;
  const raw = keep ? String(rawOf(SH.sel.r, SH.sel.c)) : initial;
  if (isPhone()) {
    SH.editing = 'fx'; fx.value = raw; showKeys(true);
    fx.focus({ preventScroll: true }); fx.setSelectionRange(raw.length, raw.length); SH.caret = raw.length;
    return;
  }
  SH.editing = 'cell';
  cellEd.value = raw; cellEd.hidden = false; placeEditor();
  fx.value = raw; cellEd.focus({ preventScroll: true });
  cellEd.setSelectionRange(raw.length, raw.length);
}
function endEdit() { SH.editing = null; SH.refIns = null; cellEd.hidden = true; showKeys(false); hideSug(); }
function commitEdit(move) {
  if (!SH.editing) return;
  const v = SH.editing === 'fx' ? fx.value : cellEd.value;
  const changed = v !== String(rawOf(SH.sel.r, SH.sel.c));
  if (changed && ddReject(SH.sel.r, SH.sel.c, v)) { endEdit(); updateFx(); if (move !== false) gw.focus({ preventScroll: true }); return; }
  endEdit();
  if (changed) setRaw(SH.sel.r, SH.sel.c, v); else updateFx();
  if (move === 'down') select(SH.sel.r + 1, SH.sel.c);
  else if (move === 'right') select(SH.sel.r, SH.sel.c + 1);
  else if (move === 'up') select(SH.sel.r - 1, SH.sel.c);
  else if (move === 'left') select(SH.sel.r, SH.sel.c - 1);
  if (move !== false) gw.focus({ preventScroll: true });
}
function cancelEdit() { endEdit(); updateFx(); gw.focus({ preventScroll: true }); }
function insertRef(a, b) {
  const inp = SH.editing === 'fx' ? fx : cellEd;
  if (!SH.refIns) SH.refIns = { start: inp.selectionStart, end: inp.selectionEnd };
  const txt = rngRef({ r1: Math.min(a.r, b.r), r2: Math.max(a.r, b.r), c1: Math.min(a.c, b.c), c2: Math.max(a.c, b.c) });
  inp.value = inp.value.slice(0, SH.refIns.start) + txt + inp.value.slice(SH.refIns.end);
  SH.refIns.end = SH.refIns.start + txt.length;
  inp.setSelectionRange(SH.refIns.end, SH.refIns.end); SH.caret = SH.refIns.end;
  if (inp === cellEd) { fx.value = cellEd.value; growEditor(); } else if (!cellEd.hidden) cellEd.value = fx.value;
}

cellEd.addEventListener('input', () => { SH.refIns = null; fx.value = cellEd.value; growEditor(); });
cellEd.addEventListener('keydown', e => {
  if (e.key === 'Enter' && !e.altKey && !e.shiftKey) { e.preventDefault(); commitEdit('down'); }
  else if (e.key === 'Enter' && e.shiftKey) { e.preventDefault(); commitEdit('up'); }
  else if (e.key === 'Tab') { e.preventDefault(); commitEdit(e.shiftKey ? 'left' : 'right'); }
  else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
});
cellEd.addEventListener('blur', () => { if (SH.editing === 'cell' && !SH.drag) setTimeout(() => { if (SH.editing === 'cell' && document.activeElement !== cellEd && document.activeElement !== fx) commitEdit(false); }, 0); });
fx.addEventListener('focus', () => { if (SH.editing === 'cell') { SH.editing = 'fx'; return; } if (!SH.editing) { SH.editing = 'fx'; SH.refIns = null; if (isPhone()) showKeys(true); } });
['keyup', 'click', 'select', 'input'].forEach(ev => fx.addEventListener(ev, () => { SH.caret = fx.selectionStart; }));
fx.addEventListener('input', () => { SH.refIns = null; if (!cellEd.hidden) { cellEd.value = fx.value; growEditor(); } });
fx.addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); commitEdit('down'); }
  else if (e.key === 'Tab') { e.preventDefault(); commitEdit(e.shiftKey ? 'left' : 'right'); }
  else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
});
fx.addEventListener('blur', () => { setTimeout(() => { if (Date.now() - SH.keyTap < 700) return; if (SH.editing === 'fx' && document.activeElement !== fx && !SH.drag && document.activeElement !== cellEd) commitEdit(false); }, 0); });

/* --- clavier de formules et fonctions */
// Fonctions proposées dans la liste ƒx et l'aide à la saisie : [nom, syntaxe, description, catégorie]
const FNS = [
  ['SOMME', 'SOMME(plage)', 'Additionne les nombres', 'Essentielles'],
  ['MOYENNE', 'MOYENNE(plage)', 'Moyenne des nombres', 'Essentielles'],
  ['SI', 'SI(test;si_vrai;si_faux)', 'Choisit une valeur selon une condition', 'Essentielles'],
  ['NB', 'NB(plage)', 'Compte les cellules qui contiennent un nombre', 'Essentielles'],
  ['NBVAL', 'NBVAL(plage)', 'Compte les cellules non vides', 'Essentielles'],
  ['MAX', 'MAX(plage)', 'Plus grande valeur', 'Essentielles'],
  ['MIN', 'MIN(plage)', 'Plus petite valeur', 'Essentielles'],
  ['ARRONDI', 'ARRONDI(nombre;décimales)', 'Arrondit à n décimales', 'Essentielles'],
  ['SOMME.SI', 'SOMME.SI(plage;critère;[plage_somme])', 'Additionne selon un critère', 'Conditions'],
  ['SOMME.SI.ENS', 'SOMME.SI.ENS(plage_somme;plage1;critère1;…)', 'Additionne selon plusieurs critères', 'Conditions'],
  ['NB.SI', 'NB.SI(plage;critère)', 'Compte les cellules qui remplissent un critère', 'Conditions'],
  ['NB.SI.ENS', 'NB.SI.ENS(plage1;critère1;…)', 'Compte selon plusieurs critères', 'Conditions'],
  ['MOYENNE.SI', 'MOYENNE.SI(plage;critère;[plage_moyenne])', 'Moyenne selon un critère', 'Conditions'],
  ['MOYENNE.SI.ENS', 'MOYENNE.SI.ENS(plage_moyenne;plage1;critère1;…)', 'Moyenne selon plusieurs critères', 'Conditions'],
  ['MAX.SI.ENS', 'MAX.SI.ENS(plage_max;plage1;critère1;…)', 'Maximum selon des critères', 'Conditions'],
  ['MIN.SI.ENS', 'MIN.SI.ENS(plage_min;plage1;critère1;…)', 'Minimum selon des critères', 'Conditions'],
  ['SIERREUR', 'SIERREUR(valeur;sinon)', 'Remplace une erreur par une autre valeur', 'Conditions'],
  ['SI.CONDITIONS', 'SI.CONDITIONS(test1;valeur1;test2;valeur2;…)', 'Plusieurs SI enchaînés', 'Conditions'],
  ['SI.MULTIPLE', 'SI.MULTIPLE(expression;valeur1;résultat1;…;[sinon])', 'Choisit selon la valeur d\'une expression', 'Conditions'],
  ['ET', 'ET(test1;test2;…)', 'Vrai si tous les tests sont vrais', 'Conditions'],
  ['OU', 'OU(test1;test2;…)', 'Vrai si au moins un test est vrai', 'Conditions'],
  ['NON', 'NON(test)', 'Inverse vrai et faux', 'Conditions'],
  ['RECHERCHEX', 'RECHERCHEX(valeur;plage_recherche;plage_renvoyée;[si_absent])', 'Cherche une valeur et renvoie la valeur correspondante', 'Recherche'],
  ['RECHERCHEV', 'RECHERCHEV(valeur;table;colonne;FAUX)', 'Cherche dans la 1re colonne d\'un tableau', 'Recherche'],
  ['RECHERCHEH', 'RECHERCHEH(valeur;table;ligne;FAUX)', 'Cherche dans la 1re ligne d\'un tableau', 'Recherche'],
  ['INDEX', 'INDEX(plage;ligne;[colonne])', 'Valeur à une position donnée', 'Recherche'],
  ['EQUIV', 'EQUIV(valeur;plage;0)', 'Position d\'une valeur dans une plage', 'Recherche'],
  ['EQUIVX', 'EQUIVX(valeur;plage)', 'Position d\'une valeur (version moderne)', 'Recherche'],
  ['FILTRE', 'FILTRE(plage;condition)', 'Garde les lignes qui remplissent une condition', 'Recherche'],
  ['UNIQUE', 'UNIQUE(plage)', 'Liste sans doublons', 'Recherche'],
  ['TRIER', 'TRIER(plage;[colonne];[1 ou -1])', 'Trie une plage', 'Recherche'],
  ['SEQUENCE', 'SEQUENCE(lignes;[colonnes];[début];[pas])', 'Suite de nombres', 'Recherche'],
  ['CHOISIR', 'CHOISIR(numéro;valeur1;valeur2;…)', 'Choisit dans une liste', 'Recherche'],
  ['DECALER', 'DECALER(réf;lignes;colonnes;[hauteur];[largeur])', 'Plage décalée', 'Recherche'],
  ['TRANSPOSE', 'TRANSPOSE(plage)', 'Échange lignes et colonnes', 'Recherche'],
  ['AUJOURDHUI', 'AUJOURDHUI()', 'Date du jour', 'Dates'],
  ['MAINTENANT', 'MAINTENANT()', 'Date et heure actuelles', 'Dates'],
  ['DATE', 'DATE(année;mois;jour)', 'Construit une date', 'Dates'],
  ['ANNEE', 'ANNEE(date)', 'Année d\'une date', 'Dates'],
  ['MOIS', 'MOIS(date)', 'Mois d\'une date (1 à 12)', 'Dates'],
  ['JOUR', 'JOUR(date)', 'Jour du mois', 'Dates'],
  ['JOURSEM', 'JOURSEM(date;2)', 'Jour de la semaine (lundi = 1 avec 2)', 'Dates'],
  ['NO.SEMAINE.ISO', 'NO.SEMAINE.ISO(date)', 'Numéro de semaine', 'Dates'],
  ['DATEDIF', 'DATEDIF(début;fin;"Y")', 'Écart entre deux dates en années ("Y"), mois ("M") ou jours ("D")', 'Dates'],
  ['JOURS', 'JOURS(fin;début)', 'Nombre de jours entre deux dates', 'Dates'],
  ['NB.JOURS.OUVRES', 'NB.JOURS.OUVRES(début;fin;[fériés])', 'Jours ouvrés entre deux dates', 'Dates'],
  ['SERIE.JOUR.OUVRE', 'SERIE.JOUR.OUVRE(début;nb_jours;[fériés])', 'Date après n jours ouvrés', 'Dates'],
  ['MOIS.DECALER', 'MOIS.DECALER(date;mois)', 'Même jour, n mois plus tard', 'Dates'],
  ['FIN.MOIS', 'FIN.MOIS(date;mois)', 'Dernier jour du mois', 'Dates'],
  ['CONCAT', 'CONCAT(texte1;texte2;…)', 'Assemble des textes ou des plages', 'Texte'],
  ['JOINDRE.TEXTE', 'JOINDRE.TEXTE(séparateur;VRAI;plage)', 'Assemble avec un séparateur', 'Texte'],
  ['GAUCHE', 'GAUCHE(texte;n)', 'Premiers caractères', 'Texte'],
  ['DROITE', 'DROITE(texte;n)', 'Derniers caractères', 'Texte'],
  ['STXT', 'STXT(texte;début;n)', 'Caractères au milieu d\'un texte', 'Texte'],
  ['NBCAR', 'NBCAR(texte)', 'Longueur d\'un texte', 'Texte'],
  ['MAJUSCULE', 'MAJUSCULE(texte)', 'Met en majuscules', 'Texte'],
  ['MINUSCULE', 'MINUSCULE(texte)', 'Met en minuscules', 'Texte'],
  ['NOMPROPRE', 'NOMPROPRE(texte)', 'Première lettre de chaque mot en majuscule', 'Texte'],
  ['SUPPRESPACE', 'SUPPRESPACE(texte)', 'Retire les espaces en trop', 'Texte'],
  ['SUBSTITUE', 'SUBSTITUE(texte;ancien;nouveau)', 'Remplace un texte par un autre', 'Texte'],
  ['CHERCHE', 'CHERCHE(cherché;texte)', 'Position d\'un texte (sans tenir compte des majuscules)', 'Texte'],
  ['TEXTE', 'TEXTE(valeur;"0,00")', 'Met un nombre en forme de texte', 'Texte'],
  ['CNUM', 'CNUM(texte)', 'Convertit un texte en nombre', 'Texte'],
  ['ABS', 'ABS(nombre)', 'Valeur absolue', 'Maths'],
  ['ENT', 'ENT(nombre)', 'Partie entière', 'Maths'],
  ['ARRONDI.SUP', 'ARRONDI.SUP(nombre;décimales)', 'Arrondit vers le haut', 'Maths'],
  ['ARRONDI.INF', 'ARRONDI.INF(nombre;décimales)', 'Arrondit vers le bas', 'Maths'],
  ['MOD', 'MOD(nombre;diviseur)', 'Reste d\'une division', 'Maths'],
  ['PUISSANCE', 'PUISSANCE(nombre;puissance)', 'Élève à une puissance', 'Maths'],
  ['RACINE', 'RACINE(nombre)', 'Racine carrée', 'Maths'],
  ['PRODUIT', 'PRODUIT(plage)', 'Multiplie les nombres', 'Maths'],
  ['SOMMEPROD', 'SOMMEPROD(plage1;plage2)', 'Somme des produits', 'Maths'],
  ['ALEA.ENTRE.BORNES', 'ALEA.ENTRE.BORNES(min;max)', 'Nombre entier au hasard', 'Maths'],
  ['MEDIANE', 'MEDIANE(plage)', 'Valeur du milieu', 'Statistiques'],
  ['RANG', 'RANG(nombre;plage;[0 ou 1])', 'Classement d\'un nombre', 'Statistiques'],
  ['GRANDE.VALEUR', 'GRANDE.VALEUR(plage;k)', 'k-ième plus grande valeur', 'Statistiques'],
  ['PETITE.VALEUR', 'PETITE.VALEUR(plage;k)', 'k-ième plus petite valeur', 'Statistiques'],
  ['ECARTYPE', 'ECARTYPE(plage)', 'Écart type', 'Statistiques'],
  ['CENTILE', 'CENTILE(plage;k)', 'Valeur au k-ième centile (k entre 0 et 1)', 'Statistiques'],
  ['QUARTILE', 'QUARTILE(plage;quart)', 'Quartile (0 à 4)', 'Statistiques'],
  ['MODE', 'MODE(plage)', 'Valeur la plus fréquente', 'Statistiques'],
  ['PREVISION', 'PREVISION(x;y_connus;x_connus)', 'Prévision par tendance linéaire', 'Statistiques'],
  ['VPM', 'VPM(taux;nb_périodes;montant)', 'Mensualité d\'un emprunt', 'Finances'],
  ['VA', 'VA(taux;nb_périodes;paiement)', 'Valeur actuelle', 'Finances'],
  ['VC', 'VC(taux;nb_périodes;paiement)', 'Valeur future d\'une épargne', 'Finances'],
  ['TAUX', 'TAUX(nb_périodes;paiement;va)', 'Taux d\'intérêt', 'Finances'],
  ['ESTVIDE', 'ESTVIDE(cellule)', 'Vrai si la cellule est vide', 'Informations'],
  ['ESTNUM', 'ESTNUM(valeur)', 'Vrai si c\'est un nombre', 'Informations'],
  ['ESTERREUR', 'ESTERREUR(valeur)', 'Vrai si c\'est une erreur', 'Informations']
];
const KEYS = [['=', '='], ['SOMME(', 'SOMME', 'fn'], ['SI(', 'SI', 'fn'], ['(', '('], [')', ')'], [';', ';'], [':', ':'], ['+', '+'], ['-', '−'], ['*', '×'], ['/', '÷'], ['"', '" "'], ['%', '%'], ['$', '$'], ['>', '>'], ['<', '<'], ['__fn', 'ƒ Fonctions', 'fn'], ['__no', 'Annuler', 'no'], ['__ok', 'Valider', 'ok']];
$('#fxkRow').innerHTML = KEYS.map(([v, l, c]) => `<button class="fxk ${c || ''}" data-k="${esc(v)}">${esc(l)}</button>`).join('');
const FN_CATS = ['Toutes', ...new Set(FNS.map(f => f[3]))];
function renderFnList(cat = 'Essentielles') {
  $('#fnList').innerHTML = `<div class="fn-cats">${FN_CATS.map(c => `<button class="fn-cat" data-cat="${c}" aria-pressed="${c === cat}">${c}</button>`).join('')}</div>`
    + FNS.filter(f => cat === 'Toutes' || f[3] === cat).map(([n, sig, d]) => `<button class="fn-it" data-fn="${n}"><b>${esc(sig)}</b><small>${esc(d)}</small></button>`).join('');
}
renderFnList();

/* --- aide à la saisie des formules : suggestions et syntaxe */
const fnSug = document.createElement('div'); fnSug.className = 'fn-sug'; fnSug.hidden = true; document.body.appendChild(fnSug);
const SUG = { items: [], i: 0, start: 0, end: 0 };
const FN_BY = Object.fromEntries(FNS.map(f => [f[0], f]));
function edInput() { return SH.editing === 'cell' ? cellEd : SH.editing === 'fx' ? fx : null; }
function hideSug() { fnSug.hidden = true; SUG.items = []; }
function updateSug() {
  const inp = edInput();
  if (!inp || inp.value[0] !== '=') return hideSug();
  const pos = inp.selectionStart, before = inp.value.slice(0, pos);
  // pas d'aide dans un texte entre guillemets
  if ((before.match(/"/g) || []).length % 2) return hideSug();
  const m = /(^|[=(;+\-*/&<>:^ ])([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ0-9.]*)$/.exec(before);
  let html = '';
  SUG.items = [];
  if (m && !/^[A-Za-z]{1,3}\d+$/.test(m[2])) {
    const q = norm(m[2]).toUpperCase();
    const all = [...FNS.map(f => f[0]), ...Object.keys(FR2EN).filter(k => !FR_ALIASES.has(k) && !FN_BY[k])];
    SUG.items = [...new Set(all)].filter(n => norm(n).toUpperCase().startsWith(q)).slice(0, 7);
    const names = Object.keys(state.cur.names || {}).filter(n => norm(n).toUpperCase().startsWith(q)).slice(0, 3);
    SUG.names = names;
    if (SUG.items.length || names.length) {
      SUG.start = pos - m[2].length; SUG.end = pos; SUG.i = Math.min(SUG.i, SUG.items.length + names.length - 1); if (SUG.i < 0) SUG.i = 0;
      html = SUG.items.map((n, k) => `<button class="sug-it" data-k="${k}" aria-selected="${k === SUG.i}"><b>${esc(n)}</b>${FN_BY[n] ? `<small>${esc(FN_BY[n][2])}</small>` : ''}</button>`).join('')
        + names.map((n, k) => `<button class="sug-it" data-name="${esc(n)}" aria-selected="${SUG.items.length + k === SUG.i}"><b>${esc(n)}</b><small>Plage nommée</small></button>`).join('');
    }
  }
  if (!html) {
    // syntaxe de la fonction dans laquelle se trouve le curseur
    let depth = 0, fnName = null, arg = 0;
    for (let k = before.length - 1; k >= 0; k--) {
      const ch = before[k];
      if (ch === ')') depth++;
      else if (ch === '(') { if (depth === 0) { const mm = /([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ0-9._]*)$/.exec(before.slice(0, k)); fnName = mm ? norm(mm[1]).toUpperCase() : null; break; } depth--; }
      else if (ch === ';' && depth === 0) arg++;
    }
    const f = fnName && FN_BY[fnName];
    if (f) {
      const inner = f[1].slice(f[1].indexOf('(') + 1, -1).split(';');
      const sig = f[0] + '(' + inner.map((a, k) => k === Math.min(arg, inner.length - 1) ? `<u>${esc(a)}</u>` : esc(a)).join(';') + ')';
      html = `<div class="sug-hint"><b>${sig}</b><small>${esc(f[2])}</small></div>`;
    }
  }
  if (!html) return hideSug();
  fnSug.innerHTML = html; fnSug.hidden = false;
  const r = inp.getBoundingClientRect();
  const below = r.bottom + 4, h = fnSug.offsetHeight;
  fnSug.style.left = Math.max(8, Math.min(r.left, innerWidth - fnSug.offsetWidth - 8)) + 'px';
  fnSug.style.top = (below + h > innerHeight - 8 ? Math.max(8, r.top - h - 4) : below) + 'px';
}
function acceptSug(k) {
  const inp = edInput(); if (!inp) return;
  const n = k < SUG.items.length ? SUG.items[k] : (SUG.names || [])[k - SUG.items.length];
  if (!n) return;
  const ins = k < SUG.items.length ? n + '(' : n;
  inp.value = inp.value.slice(0, SUG.start) + ins + inp.value.slice(SUG.end);
  const np = SUG.start + ins.length; inp.focus({ preventScroll: true }); inp.setSelectionRange(np, np); SH.caret = np; SH.refIns = null;
  if (inp === cellEd) { fx.value = cellEd.value; growEditor(); } else if (!cellEd.hidden) { cellEd.value = fx.value; growEditor(); }
  SUG.i = 0; updateSug();
}
fnSug.addEventListener('mousedown', e => { e.preventDefault(); markTap(); });
fnSug.addEventListener('click', e => { const b = e.target.closest('.sug-it'); if (!b) return; markTap(); acceptSug(b.dataset.name ? SUG.items.length + (SUG.names || []).indexOf(b.dataset.name) : +b.dataset.k); });
function sugKey(e) {
  if (fnSug.hidden) return false;
  const n = SUG.items.length + (SUG.names || []).length;
  if (!n) return false;
  if (e.key === 'ArrowDown') { SUG.i = (SUG.i + 1) % n; updateSug(); return true; }
  if (e.key === 'ArrowUp') { SUG.i = (SUG.i - 1 + n) % n; updateSug(); return true; }
  if (e.key === 'Tab' || (e.key === 'Enter' && !e.shiftKey)) { acceptSug(SUG.i); return true; }
  if (e.key === 'Escape') { hideSug(); return true; }
  return false;
}
['input', 'click', 'keyup'].forEach(ev => { cellEd.addEventListener(ev, e => { if (ev === 'keyup' && /^(ArrowUp|ArrowDown|Enter|Tab|Escape)$/.test(e.key)) return; SUG.i = ev === 'input' ? 0 : SUG.i; updateSug(); }); fx.addEventListener(ev, e => { if (ev === 'keyup' && /^(ArrowUp|ArrowDown|Enter|Tab|Escape)$/.test(e.key)) return; SUG.i = ev === 'input' ? 0 : SUG.i; updateSug(); }); });
cellEd.addEventListener('keydown', e => { if (sugKey(e)) { e.preventDefault(); e.stopImmediatePropagation(); } }, true);
fx.addEventListener('keydown', e => { if (sugKey(e)) { e.preventDefault(); e.stopImmediatePropagation(); } }, true);
function showKeys(on) {
  $('#fxKeys').hidden = !on; if (!on) $('#fnList').hidden = true;
  $('#fxBtn').setAttribute('aria-pressed', String(on && !$('#fnList').hidden));
  if (state.view === 'sheet') requestAnimationFrame(() => { drawSel(); });
}
function insertAtCaret(txt) {
  if (!SH.editing) { SH.editing = 'fx'; fx.value = String(rawOf(SH.sel.r, SH.sel.c)); SH.caret = fx.value.length; }
  const inp = SH.editing === 'cell' ? cellEd : fx;
  const pos = inp === fx ? (SH.caret ?? fx.value.length) : inp.selectionStart;
  let v = inp.value, ins = txt;
  if (v === '' && txt !== '=' ) ins = '=' + txt;
  inp.value = v.slice(0, pos) + ins + v.slice(pos);
  const np = pos + ins.length;
  inp.focus({ preventScroll: true }); inp.setSelectionRange(np, np); SH.caret = np; SH.refIns = null;
  if (inp === cellEd) { fx.value = cellEd.value; growEditor(); } else if (!cellEd.hidden) { cellEd.value = fx.value; growEditor(); }
}
const markTap = () => { SH.keyTap = Date.now(); };
['pointerdown', 'mousedown'].forEach(ev => $('#fxKeys').addEventListener(ev, e => { if (e.target.closest('button')) { e.preventDefault(); markTap(); } }));
$('#fxKeys').addEventListener('touchstart', markTap, { passive: true });
$('#fxKeys').addEventListener('click', e => {
  markTap();
  const cat = e.target.closest('[data-cat]');
  if (cat) { renderFnList(cat.dataset.cat); return; }
  const fnb = e.target.closest('[data-fn]');
  if (fnb) { insertAtCaret(fnb.dataset.fn + '('); $('#fnList').hidden = true; $('#fxBtn').setAttribute('aria-pressed', 'false'); return; }
  const b = e.target.closest('[data-k]'); if (!b) return;
  const k = b.dataset.k;
  if (k === '__ok') { commitEdit(isPhone() ? false : 'down'); if (isPhone()) fx.blur(); return; }
  if (k === '__no') return cancelEdit();
  if (k === '__fn') { const l = $('#fnList'); l.hidden = !l.hidden; $('#fxBtn').setAttribute('aria-pressed', String(!l.hidden)); (SH.editing === 'cell' ? cellEd : fx).focus({ preventScroll: true }); return; }
  if (k === '"') { insertAtCaret('""'); const inp = SH.editing === 'cell' ? cellEd : fx; const p = inp.selectionStart - 1; inp.setSelectionRange(p, p); SH.caret = p; return; }
  insertAtCaret(k);
});
$('#fxBtn').addEventListener('mousedown', e => { e.preventDefault(); markTap(); });
$('#fxBtn').addEventListener('click', () => {
  markTap();
  if (!SH.editing) { if (isPhone()) startEdit('', true); else { SH.editing = 'fx'; fx.value = String(rawOf(SH.sel.r, SH.sel.c)); SH.caret = fx.value.length; fx.focus(); } }
  $('#fxKeys').hidden = false;
  const l = $('#fnList'); l.hidden = !l.hidden; $('#fxBtn').setAttribute('aria-pressed', String(!l.hidden));
  requestAnimationFrame(drawSel);
});

/* --- souris */
const tdAt = e => { const td = e.target.closest && e.target.closest('td'); return td && gw.contains(td) ? { r: +td.dataset.r, c: +td.dataset.c } : null; };
gw.addEventListener('mousedown', e => {
  if (e.target === fillH) return;
  if (e.target.classList.contains('rz')) return startResize(e);
  const th = e.target.closest('th');
  if (th && gw.contains(th)) {
    e.preventDefault(); commitEdit(false);
    const cur = state.cur;
    if (th.dataset.c != null) { SH.anc = { r: cur.rows - 1, c: +th.dataset.c }; select(0, +th.dataset.c, true); }
    else if (th.dataset.r != null) { SH.anc = { r: +th.dataset.r, c: cur.cols - 1 }; select(+th.dataset.r, 0, true); }
    else { SH.anc = { r: cur.rows - 1, c: cur.cols - 1 }; select(0, 0, true); }
    gw.focus({ preventScroll: true }); return;
  }
  const p = tdAt(e); if (!p || e.target === cellEd) return;
  e.preventDefault();
  if (SH.editing && isFormulaMode()) { SH.drag = { mode: 'ref', a: p }; insertRef(p, p); return; }
  if (SH.editing) commitEdit(false);
  const pf = (cellOf(p.r, p.c) || {}).f || {};
  if (pf.chk && !e.shiftKey) { select(p.r, p.c); toggleCheck(p.r, p.c); gw.focus({ preventScroll: true }); return; }
  if (pf.dd && !e.shiftKey) { const td = tdOf(p.r, p.c), rc = td.getBoundingClientRect(); if (e.clientX > rc.right - 22 || (isPhone() && SH.sel.r === p.r && SH.sel.c === p.c)) { select(p.r, p.c); openDdPick(p.r, p.c); return; } }
  const R0 = rng();
  SH.tapAgain = isPhone() && R0.r1 === R0.r2 && R0.c1 === R0.c2 && SH.sel.r === p.r && SH.sel.c === p.c;
  select(p.r, p.c, e.shiftKey); SH.drag = { mode: 'sel' };
  gw.focus({ preventScroll: true });
});
gw.addEventListener('mouseover', e => {
  if (!SH.drag || !(e.buttons & 1)) return;
  const p = tdAt(e); if (!p) return;
  if (SH.drag.mode === 'sel') { SH.sel = p; drawSel(); updateFx(); updateSelStatus(); }
  else if (SH.drag.mode === 'ref') insertRef(SH.drag.a, p);
});
document.addEventListener('mouseup', () => {
  if (!SH.drag) return;
  const m = SH.drag.mode; SH.drag = null;
  if (m === 'ref') (SH.editing === 'fx' ? fx : cellEd).focus({ preventScroll: true });
});
gw.addEventListener('click', e => { if (SH.tapAgain && !SH.editing && tdAt(e)) { SH.tapAgain = false; startEdit('', true); } });
gw.addEventListener('dblclick', e => {
  const th = e.target.closest('th.hc'); if (th) { unhideCol(+th.dataset.c); return; }
  const p = tdAt(e); if (!p || SH.editing) return;
  const pf = (cellOf(p.r, p.c) || {}).f || {}; if (pf.chk) return;
  select(p.r, p.c); startEdit('', true);
});

function startResize(e) {
  e.preventDefault(); e.stopPropagation();
  const c = +e.target.dataset.c, startX = e.clientX, col = $$('#grid col')[c + 1], w0 = col.getBoundingClientRect().width || 104;
  const before = snap();
  const mv = ev => { const w = Math.max(40, Math.round(w0 + ev.clientX - startX)); col.style.width = w + 'px'; state.cur.colW[c] = w; };
  const up = () => {
    document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up);
    SH.hist.push(before); SH.fut = []; buildGrid(); renderValues(); drawSel(); scheduleSave();
  };
  document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
}

/* --- clavier */
gw.addEventListener('keydown', e => {
  if (SH.editing || e.target !== gw) return;
  const mod = e.ctrlKey || e.metaKey, k = e.key;
  const mv = e.altKey ? null : { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
  if (mv) {
    e.preventDefault();
    let { r, c } = SH.sel;
    if (mod) { r = mv[0] > 0 ? state.cur.rows - 1 : mv[0] < 0 ? 0 : r; c = mv[1] > 0 ? state.cur.cols - 1 : mv[1] < 0 ? 0 : c; }
    else {
      const m = SH.merges.find(q => q.r1 <= r && r <= q.r2 && q.c1 <= c && c <= q.c2);
      if (m && mv[0] > 0) r = m.r2; if (m && mv[1] > 0) c = m.c2;
      r += mv[0]; c += mv[1];
    }
    return select(r, c, e.shiftKey);
  }
  if (k === 'Enter') { e.preventDefault(); return select(SH.sel.r + (e.shiftKey ? -1 : 1), SH.sel.c); }
  if (k === 'Tab') { e.preventDefault(); return select(SH.sel.r, SH.sel.c + (e.shiftKey ? -1 : 1)); }
  if (k === 'F2') { e.preventDefault(); return startEdit('', true); }
  if (k === ' ' && !mod) { const f = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {}; if (f.chk) { e.preventDefault(); return toggleCheck(SH.sel.r, SH.sel.c); } }
  if (k === 'ArrowDown' && e.altKey) { const f = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {}; if (f.dd) { e.preventDefault(); return openDdPick(SH.sel.r, SH.sel.c); } }
  if (mod && (k === 'f' || k === 'F')) { e.preventDefault(); return openSFind(false); }
  if (mod && (k === 'h' || k === 'H')) { e.preventDefault(); return openSFind(true); }
  if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); return clearRange(); }
  if (k === 'Home') { e.preventDefault(); return select(mod ? 0 : SH.sel.r, 0, e.shiftKey); }
  if (mod && (k === 'z' || k === 'Z')) { e.preventDefault(); return e.shiftKey ? redo() : undo(); }
  if (mod && (k === 'y' || k === 'Y')) { e.preventDefault(); return redo(); }
  if (mod && (k === 'b' || k === 'B')) { e.preventDefault(); return setFmt('b'); }
  if (mod && (k === 'i' || k === 'I')) { e.preventDefault(); return setFmt('i'); }
  if (mod && (k === 'u' || k === 'U')) { e.preventDefault(); return setFmt('u'); }
  if (mod && k === 'PageDown') { e.preventDefault(); return switchSheet(Math.min(state.cur.active + 1, state.cur.sheets.length - 1)); }
  if (mod && k === 'PageUp') { e.preventDefault(); return switchSheet(Math.max(state.cur.active - 1, 0)); }
  if (mod && (k === 'd' || k === 'D')) { e.preventDefault(); return fillDown(); }
  if (mod && (k === 'r' || k === 'R')) { e.preventDefault(); return fillRight(); }
  if (mod && (k === 'a' || k === 'A')) { e.preventDefault(); SH.anc = { r: state.cur.rows - 1, c: state.cur.cols - 1 }; return select(0, 0, true); }
  if (!mod && !e.altKey && k.length === 1) { e.preventDefault(); startEdit(k); }
});

/* --- copier / coller */
function rangeTSV() {
  const R = rng(), lines = [];
  for (let r = R.r1; r <= R.r2; r++) {
    const row = [];
    for (let c = R.c1; c <= R.c2; c++) { const v = (SH.vals[r] || [])[c]; row.push(v == null ? '' : typeof v === 'number' ? (((cellOf(r, c) || {}).f || {}).n === 'date' ? fmtDate(v) : numToRaw(v)) : typeof v === 'object' ? (v.value || '') : String(v)); }
    lines.push(row.join('\t'));
  }
  return lines.join('\n');
}
function keepClip(cut) {
  const R = rng(), cells = [];
  for (let r = R.r1; r <= R.r2; r++) { const row = []; for (let c = R.c1; c <= R.c2; c++) row.push(cellCopy(r, c)); cells.push(row); }
  SH.clip = { R, cells, tsv: rangeTSV(), cut };
  return SH.clip.tsv;
}
document.addEventListener('copy', e => { if (state.view !== 'sheet' || SH.editing || document.activeElement !== gw) return; e.preventDefault(); e.clipboardData.setData('text/plain', keepClip(false)); toast('Copié'); });
document.addEventListener('cut', e => { if (state.view !== 'sheet' || SH.editing || document.activeElement !== gw) return; e.preventDefault(); e.clipboardData.setData('text/plain', keepClip(true)); clearRange(); });
document.addEventListener('paste', e => {
  if (state.view !== 'sheet' || SH.editing || document.activeElement !== gw) return;
  e.preventDefault();
  const txt = (e.clipboardData.getData('text/plain') || '').replace(/\r/g, '');
  const rows = txt.split('\n'); if (rows.length > 1 && rows[rows.length - 1] === '') rows.pop();
  const R = rng();
  pushHist();
  const grew = ensureSize(R.r1 + rows.length - 1, R.c1 + Math.max(...rows.map(l => l.split('\t').length)) - 1);
  if (grew) loadHF();
  const clip = SH.clip && SH.clip.tsv === txt ? SH.clip : null;
  if (clip) {
    const dr = clip.cut ? 0 : R.r1 - clip.R.r1, dc = clip.cut ? 0 : R.c1 - clip.R.c1;
    batch(() => clip.cells.forEach((row, i) => row.forEach((x, j) => putCell(R.r1 + i, R.c1 + j, x, dr, dc))));
    if (clip.cut) SH.clip = null;
  } else batch(() => rows.forEach((line, i) => line.split('\t').forEach((v, j) => applyRaw(R.r1 + i, R.c1 + j, v))));
  afterChange();
  SH.anc = { r: R.r1, c: R.c1 };
  select(R.r1 + rows.length - 1, R.c1 + rows[0].split('\t').length - 1, true);
});

/* --- recopie des formules et séries */
const fillH = $('#fillH'), fillBox = $('#fillBox');
const cellCopy = (r, c) => { const x = cellOf(r, c); return x ? clone(x) : null; };
function shiftFormula(f, dr, dc) {
  return outsideStrings(f, p => p.replace(/(^|[^A-Za-z0-9_.$!'])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_.])/g, (m, pre, ac, col, ar, row) => {
    let c = 0; for (const ch of col.toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64; c--;
    let r = +row - 1;
    if (!ac) c += dc; if (!ar) r += dr;
    if (c < 0 || r < 0) return pre + '#REF!';
    return pre + ac + COLN(c) + ar + (r + 1);
  }));
}
function putCell(r, c, src, dr, dc) {
  const k = addr(r, c), cells = state.cur.cells;
  if (!src) { delete cells[k]; applyRaw(r, c, ''); return; }
  let raw = String(src.v == null ? '' : src.v);
  if (raw[0] === '=') raw = '=' + shiftFormula(raw.slice(1), dr, dc);
  cells[k] = { v: '', f: clone(src.f || {}) };
  applyRaw(r, c, raw);
  if (!cells[k].v && !Object.keys(cells[k].f).length) delete cells[k];
}
function seriesOf(src) {
  const nums = src.map(x => x && x.v !== '' && String(x.v)[0] !== '=' ? (parseNum(x.v) ?? parseDate(x.v)) : null);
  if (!nums.length || nums.some(n => n == null)) return null;
  const isDate = ((src[0].f || {}).n === 'date');
  if (nums.length === 1) return isDate ? { a: nums[0], step: 1, date: true } : null;
  const step = (nums[nums.length - 1] - nums[0]) / (nums.length - 1);
  for (let i = 1; i < nums.length; i++) if (Math.abs(nums[i] - nums[i - 1] - step) > 1e-9) return null;
  return { a: nums[0], step, date: isDate };
}
function doFill(S, T, plain) {
  ensureSize(T.r2, T.c2);
  pushHist();
  const vertical = T.c1 === S.c1 && T.c2 === S.c2;
  batch(() => {
    if (vertical) {
      const len = S.r2 - S.r1 + 1;
      for (let c = S.c1; c <= S.c2; c++) {
        const src = []; for (let r = S.r1; r <= S.r2; r++) src.push(cellCopy(r, c));
        const ser = plain ? null : seriesOf(src);
        for (let r = T.r1; r <= T.r2; r++) {
          if (r >= S.r1 && r <= S.r2) continue;
          const i = ((r - S.r1) % len + len) % len;
          if (ser) { const v = ser.a + ser.step * (r - S.r1); state.cur.cells[addr(r, c)] = { v: '', f: clone(src[i].f || {}) }; applyRaw(r, c, ser.date ? fmtDate(v) : numToRaw(v)); }
          else putCell(r, c, src[i], r - (S.r1 + i), 0);
        }
      }
    } else {
      const len = S.c2 - S.c1 + 1;
      for (let r = S.r1; r <= S.r2; r++) {
        const src = []; for (let c = S.c1; c <= S.c2; c++) src.push(cellCopy(r, c));
        const ser = plain ? null : seriesOf(src);
        for (let c = T.c1; c <= T.c2; c++) {
          if (c >= S.c1 && c <= S.c2) continue;
          const i = ((c - S.c1) % len + len) % len;
          if (ser) { const v = ser.a + ser.step * (c - S.c1); state.cur.cells[addr(r, c)] = { v: '', f: clone(src[i].f || {}) }; applyRaw(r, c, ser.date ? fmtDate(v) : numToRaw(v)); }
          else putCell(r, c, src[i], 0, c - (S.c1 + i));
        }
      }
    }
  });
  afterChange();
}
function fillDown() {
  commitEdit(false);
  const R = rng();
  if (R.r1 === R.r2) { if (R.r1 === 0) return; doFill({ ...R, r1: R.r1 - 1, r2: R.r1 - 1 }, { ...R, r1: R.r1 - 1 }, true); }
  else doFill({ ...R, r2: R.r1 }, R, true);
  toast('Recopié vers le bas');
}
function fillRight() {
  commitEdit(false);
  const R = rng();
  if (R.c1 === R.c2) { if (R.c1 === 0) return; doFill({ ...R, c1: R.c1 - 1, c2: R.c1 - 1 }, { ...R, c1: R.c1 - 1 }, true); }
  else doFill({ ...R, c2: R.c1 }, R, true);
  toast('Recopié vers la droite');
}
fillH.addEventListener('pointerdown', e => {
  e.preventDefault(); e.stopPropagation();
  commitEdit(false);
  try { fillH.setPointerCapture(e.pointerId); } catch {}
  SH.fillDrag = { S: rng(), T: null };
});
fillH.addEventListener('pointermove', e => {
  const fd = SH.fillDrag; if (!fd) return;
  const el = document.elementFromPoint(e.clientX, e.clientY), td = el && el.closest ? el.closest('td') : null;
  const rect = gw.getBoundingClientRect();
  if (e.clientY > rect.bottom - 24) gw.scrollTop += 24; else if (e.clientY < rect.top + 40) gw.scrollTop -= 24;
  if (e.clientX > rect.right - 24) gw.scrollLeft += 24;
  if (!td || !gw.contains(td)) return;
  const r = +td.dataset.r, c = +td.dataset.c, S = fd.S;
  const down = r - S.r2, up = S.r1 - r, right = c - S.c2, left = S.c1 - c;
  const v = Math.max(down, up), h = Math.max(right, left);
  let T = null;
  if (v > 0 || h > 0) T = v >= h ? (down > 0 ? { ...S, r2: r } : { ...S, r1: r }) : (right > 0 ? { ...S, c2: c } : { ...S, c1: c });
  fd.T = T;
  if (T) { fillBox.hidden = false; box(fillBox, SH.cell[T.r1][T.c1], SH.cell[T.r2][T.c2]); } else fillBox.hidden = true;
});
function endFill() {
  const fd = SH.fillDrag; SH.fillDrag = null; fillBox.hidden = true;
  if (!fd || !fd.T) return;
  doFill(fd.S, fd.T, false);
  SH.anc = { r: fd.T.r2, c: fd.T.c2 }; select(fd.T.r1, fd.T.c1, true);
  gw.focus({ preventScroll: true });
}
fillH.addEventListener('pointerup', endFill);
fillH.addEventListener('pointercancel', () => { SH.fillDrag = null; fillBox.hidden = true; });

/* --- tri et filtre */
const rowHidden = r => !!(SH.cell[r] && SH.cell[r][0] && SH.cell[r][0].parentElement.hidden);
const hasVal = (r, c) => r >= 0 && c >= 0 && r < state.cur.rows && c < state.cur.cols && String(rawOf(r, c)) !== '';
function currentRegion(r, c) {
  const R = { r1: r, r2: r, c1: c, c2: c };
  const rowHas = rr => { for (let cc = R.c1 - 1; cc <= R.c2 + 1; cc++) if (hasVal(rr, cc)) return true; return false; };
  const colHas = cc => { for (let rr = R.r1 - 1; rr <= R.r2 + 1; rr++) if (hasVal(rr, cc)) return true; return false; };
  for (let changed = true; changed;) {
    changed = false;
    if (R.r1 > 0 && rowHas(R.r1 - 1)) { R.r1--; changed = true; }
    if (R.r2 < state.cur.rows - 1 && rowHas(R.r2 + 1)) { R.r2++; changed = true; }
    if (R.c1 > 0 && colHas(R.c1 - 1)) { R.c1--; changed = true; }
    if (R.c2 < state.cur.cols - 1 && colHas(R.c2 + 1)) { R.c2++; changed = true; }
  }
  return R;
}
function isHeaderRow(rr, R, col, first) {
  let any = false, bold = false;
  for (let c = R.c1; c <= R.c2; c++) {
    const v = (SH.vals[rr] || [])[c];
    if (v == null || v === '') continue;
    if (typeof v !== 'string') return false;
    any = true; if (((cellOf(rr, c) || {}).f || {}).b) bold = true;
  }
  if (!any) return false;
  if (bold) return true;
  if (!first) return false;
  for (let r = rr + 1; r <= R.r2; r++) if (typeof (SH.vals[r] || [])[col] === 'number') return true;
  return false;
}
// Zone de données autour de la cellule active : ligne d'en-tête exclue.
function dataZone() {
  const { r, c } = SH.sel, sel = rng(), base = currentRegion(r, c);
  if (sel.r2 > sel.r1) return { r1: sel.r1, r2: sel.r2, c1: Math.min(base.c1, sel.c1), c2: Math.max(base.c2, sel.c2), hdr: -1 };
  let hdr = -1;
  for (let rr = r - 1; rr >= base.r1; rr--) if (isHeaderRow(rr, base, c, false)) { hdr = rr; break; }
  if (hdr < 0 && r > base.r1 && isHeaderRow(base.r1, base, c, true)) hdr = base.r1;
  return { r1: hdr >= 0 ? hdr + 1 : base.r1, r2: base.r2, c1: base.c1, c2: base.c2, hdr };
}
function sortZone(desc) {
  commitEdit(false);
  const Z = dataZone(), col = SH.sel.c;
  if (Z.r2 - Z.r1 < 1) return toast('Rien à trier ici : placez-vous dans un tableau d\'au moins deux lignes.');
  if (SH.merges.some(m => m.r2 >= Z.r1 && m.r1 <= Z.r2 && m.c2 >= Z.c1 && m.c1 <= Z.c2)) return toast('Séparez d\'abord les cellules fusionnées de ce tableau pour le trier.', true);
  const rows = [];
  for (let r = Z.r1; r <= Z.r2; r++) { const cells = []; for (let c = Z.c1; c <= Z.c2; c++) cells.push(cellCopy(r, c)); rows.push({ from: r, cells, v: (SH.vals[r] || [])[col] }); }
  const rank = v => v == null || v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
  rows.sort((a, b) => {
    const ra = rank(a.v), rb = rank(b.v);
    if (ra === 3 || rb === 3) return ra - rb;
    let d = ra - rb;
    if (!d) d = ra === 0 ? a.v - b.v : ra === 1 ? String(a.v).localeCompare(String(b.v), 'fr', { numeric: true, sensitivity: 'base' }) : 0;
    return desc ? -d : d;
  });
  pushHist();
  batch(() => rows.forEach((row, i) => { const to = Z.r1 + i; row.cells.forEach((x, j) => putCell(to, Z.c1 + j, x, to - row.from, 0)); }));
  afterChange(); applyFilters();
  toast(`Trié par la colonne ${COLN(col)}, ${desc ? 'de Z à A' : 'de A à Z'}` + (Z.hdr >= 0 ? ' (ligne d\'en-tête conservée)' : ''));
}
const fkey = (r, c) => display(r, c).txt || '(vides)';
function applyFilters() {
  if (state.view !== 'sheet' || !state.cur) return;
  const F = state.cur.filter;
  let hidden = 0;
  const hr = new Set((state.cur.hid && state.cur.hid.r) || []);
  SH.cell.forEach((row, r) => { if (row[0]) row[0].parentElement.hidden = hr.has(r); });
  SH.rowTh.forEach((th, r) => th.classList.toggle('hrb', r > 0 && hr.has(r - 1) && !hr.has(r)));
  SH.colTh.forEach(th => th.classList.remove('flt'));
  if (F) {
    const cols = Object.keys(F.cols || {}).map(Number);
    cols.forEach(c => SH.colTh[c] && SH.colTh[c].classList.add('flt'));
    for (let r = F.r1; r <= Math.min(F.r2, state.cur.rows - 1); r++) {
      if (hr.has(r)) continue;
      const hide = cols.some(c => !F.cols[c].includes(fkey(r, c)));
      if (hide) { SH.cell[r][0].parentElement.hidden = true; hidden++; }
    }
  }
  SH.hiddenRows = hidden;
  if (rowHidden(SH.sel.r)) { SH.anc = { ...SH.sel }; select(SH.sel.r, SH.sel.c); } else drawSel();
  updateSelStatus();
}
function openFilter() {
  commitEdit(false);
  const cur = state.cur, c = SH.sel.c;
  let F = cur.filter;
  if (!F) { const Z = dataZone(); F = { r1: Z.r1, r2: Z.r2, cols: {} }; }
  if (F.r2 < F.r1) return toast('Aucune donnée à filtrer dans cette colonne.');
  const counts = new Map();
  for (let r = F.r1; r <= F.r2; r++) { const k = fkey(r, c); counts.set(k, (counts.get(k) || 0) + 1); }
  const vals = [...counts.keys()].sort((a, b) => a === '(vides)' ? 1 : b === '(vides)' ? -1 : a.localeCompare(b, 'fr', { numeric: true, sensitivity: 'base' }));
  const allowed = F.cols && F.cols[c];
  SH.fdraft = { F, c };
  $('#fpTitle').textContent = `Filtrer la colonne ${COLN(c)} (lignes ${F.r1 + 1} à ${F.r2 + 1})`;
  $('#fpList').innerHTML = vals.map(v => `<label class="fp-it"><input type="checkbox" value="${esc(v)}"${!allowed || allowed.includes(v) ? ' checked' : ''}><span>${esc(v)}</span><small>${counts.get(v)}</small></label>`).join('');
  $('#fpSearch').value = ''; $('#filterPop').hidden = false;
  if (!isPhone()) $('#fpSearch').focus();
}
function closeFilter() { $('#filterPop').hidden = true; SH.fdraft = null; gw.focus({ preventScroll: true }); }
function clearFilters() { if (!state.cur.filter) return; state.cur.filter = undefined; scheduleSave(); applyFilters(); toast('Toutes les lignes sont affichées.'); }
$('#fpSearch').addEventListener('input', e => { const q = norm(e.target.value); $$('#fpList .fp-it').forEach(l => { l.hidden = !!q && !norm(l.textContent).includes(q); }); });
$('#fpAll').addEventListener('click', () => $$('#fpList .fp-it:not([hidden]) input').forEach(i => { i.checked = true; }));
$('#fpNone').addEventListener('click', () => $$('#fpList .fp-it:not([hidden]) input').forEach(i => { i.checked = false; }));
$('#fpX').addEventListener('click', closeFilter);
$('#fpApply').addEventListener('click', () => {
  const d = SH.fdraft; if (!d) return closeFilter();
  const boxes = $$('#fpList input'), on = boxes.filter(i => i.checked).map(i => i.value);
  const F = { r1: d.F.r1, r2: d.F.r2, cols: { ...(d.F.cols || {}) } };
  if (on.length === boxes.length) delete F.cols[d.c]; else F.cols[d.c] = on;
  state.cur.filter = Object.keys(F.cols).length ? F : undefined;
  scheduleSave(); closeFilter(); applyFilters();
});
$('#fpClear').addEventListener('click', () => {
  const d = SH.fdraft; closeFilter(); if (!d || !state.cur.filter) return;
  delete state.cur.filter.cols[d.c];
  if (!Object.keys(state.cur.filter.cols).length) state.cur.filter = undefined;
  scheduleSave(); applyFilters();
});
$('#sStatus').addEventListener('click', e => { if (e.target.closest('[data-act="clearFilter"]')) clearFilters(); if (e.target.closest('[data-act="note"]')) openNoteEditor(); });
document.addEventListener('mousedown', e => { if (!$('#filterPop').hidden && !e.target.closest('#filterPop') && !e.target.closest('[data-s="filter"]')) closeFilter(); });

/* --- onglets */
function refreshSheetView() {
  const cur = state.cur; CH.sel = null; closeDlg(); closeDdPick();
  if (SH.ids) SH.sid = SH.ids[cur.active];
  SH.hist = []; SH.fut = []; SH.hiddenRows = 0;
  buildGrid(); renderValues(); applyFilters(); applyFreeze(); select(0, 0); renderTabs();
}
function renderTabs() {
  const cur = state.cur;
  $('#tabList').innerHTML = cur.sheets.map((sh, i) => `<button class="tab" role="tab" data-i="${i}" aria-selected="${i === cur.active}">${esc(sh.name)}${i === cur.active ? '<span class="tab-more" aria-hidden="true">▾</span>' : ''}</button>`).join('');
  const a = $('#tabList .tab[aria-selected="true"]'); if (a) a.scrollIntoView({ inline: 'nearest', block: 'nearest' });
}
function switchSheet(i) {
  const cur = state.cur;
  if (i === cur.active || !cur.sheets[i]) return;
  commitEdit(false); closeFilter(); closeMini();
  cur.active = i; refreshSheetView(); scheduleSave();
}
function uniqueName(base) {
  const has = n => state.cur.sheets.some(s => s.name.toLowerCase() === n.toLowerCase());
  if (!has(base)) return base;
  const root = base.replace(/\s*\(\d+\)$/, '');
  for (let n = 2; ; n++) { const t = `${root} (${n})`; if (!has(t)) return t; }
}
function addSheet(from) {
  commitEdit(false); closeMini();
  const cur = state.cur;
  let n = cur.sheets.length + 1; while (cur.sheets.some(s => s.name.toLowerCase() === 'feuille' + n)) n++;
  const sh = from ? { ...clone(from), id: uid(), name: uniqueName(from.name) } : { id: uid(), name: 'Feuille' + n, cells: {}, rows: 60, cols: 12, colW: {}, hid: { r: [], c: [] }, cf: [], charts: [] };
  (sh.charts || []).forEach(c => { c.id = uid(); });
  delete sh.filter;
  cur.sheets.splice(cur.active + 1, 0, sh);
  cur.active += 1;
  loadHF(); refreshSheetView(); scheduleSave();
  toast(from ? `Feuille dupliquée : ${sh.name}` : `Nouvelle feuille : ${sh.name}`);
}
function deleteSheet(i) {
  const cur = state.cur;
  if (cur.sheets.length < 2) return toast('Un classeur garde au moins une feuille.');
  commitEdit(false); closeMini();
  const name = cur.sheets[i].name;
  if (SH.hf && SH.ids) {
    SH.hf.removeSheet(SH.ids[i]);
    const ids = SH.ids.filter((_, k) => k !== i);
    cur.sheets.splice(i, 1);
    cur.sheets.forEach((sh, k) => { sh.cells = cellsFromHF(ids[k], sh.cells); });
  } else cur.sheets.splice(i, 1);
  if (cur.active > i || cur.active >= cur.sheets.length) cur.active = Math.max(0, cur.active - 1);
  loadHF(); refreshSheetView(); scheduleSave();
  toast(`Feuille « ${name} » supprimée. Les formules qui la citaient affichent #REF!.`);
}
function moveSheet(i, d) {
  const cur = state.cur, j = i + d;
  if (j < 0 || j >= cur.sheets.length) return;
  const act = cur.sheets[cur.active];
  [cur.sheets[i], cur.sheets[j]] = [cur.sheets[j], cur.sheets[i]];
  cur.active = cur.sheets.indexOf(act);
  loadHF(); renderTabs(); scheduleSave(); closeMini();
}
function renameSheet(i, name) {
  const cur = state.cur, old = cur.sheets[i].name;
  name = String(name || '').replace(/[\[\]:*?\/\\']/g, '').replace(/\s+/g, ' ').trim().slice(0, 31);
  if (!name || name === old) return renderTabs();
  if (cur.sheets.some((s, k) => k !== i && s.name.toLowerCase() === name.toLowerCase())) { toast('Une autre feuille porte déjà ce nom.', true); return renderTabs(); }
  const e = old.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(^|[^A-Za-z0-9À-ÿ_.])(?:'${e.replace(/'/g, "''")}'|${e})!`, 'gi');
  cur.sheets.forEach(sh => Object.values(sh.cells).forEach(x => { if (String(x.v)[0] === '=') x.v = outsideStrings(x.v, p => p.replace(re, (m, pre) => pre + refName(name) + '!')); }));
  cur.sheets[i].name = name;
  loadHF(); renderValues(); renderTabs(); updateFx(); scheduleSave();
}
function startRename(i) {
  closeMini(); renderTabs();
  const t = $(`#tabList .tab[data-i="${i}"]`); if (!t) return;
  t.innerHTML = `<input class="tab-in" id="tabIn" value="${esc(state.cur.sheets[i].name)}" aria-label="Nom de la feuille" autocomplete="off" autocorrect="off">`;
  const inp = $('#tabIn'); inp.focus(); inp.select();
  let done = false;
  const fin = ok => { if (done) return; done = true; ok ? renameSheet(i, inp.value) : renderTabs(); };
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); fin(true); } else if (e.key === 'Escape') { e.preventDefault(); fin(false); } e.stopPropagation(); });
  inp.addEventListener('blur', () => fin(true));
}
function placePop(pop, anchor) {
  pop.hidden = false;
  const host = $('#sheet').getBoundingClientRect(), a = anchor.getBoundingClientRect();
  pop.style.left = Math.max(8, Math.min(a.left - host.left, host.width - pop.offsetWidth - 8)) + 'px';
  if (a.top - host.top > host.height / 2) { pop.style.top = 'auto'; pop.style.bottom = (host.bottom - a.top + 6) + 'px'; }
  else { pop.style.bottom = 'auto'; pop.style.top = (a.bottom - host.top + 6) + 'px'; }
}
function closeMini() { $('#tabMenu').hidden = true; $('#fmtPop').hidden = true; SH.tabDel = false; }
function openTabMenu(el, i) {
  const m = $('#tabMenu'), n = state.cur.sheets.length;
  m.dataset.i = i; SH.tabDel = false;
  m.innerHTML = `<button data-tm="rename">${ic('pen')}Renommer</button><button data-tm="dup">${ic('copy')}Dupliquer</button>`
    + (i > 0 ? `<button data-tm="left">${ic('left')}Déplacer à gauche</button>` : '')
    + (i < n - 1 ? `<button data-tm="right">${ic('right')}Déplacer à droite</button>` : '')
    + (state.cur.sheets[i].pivot ? `<button data-tm="pivot">${ic('pivot')}Actualiser le tableau croisé</button>` : '')
    + (n > 1 ? `<hr><button data-tm="del" class="danger">${ic('trash')}<span>Supprimer la feuille</span></button>` : '');
  placePop(m, el);
}
$('#tabList').addEventListener('click', e => {
  const t = e.target.closest('.tab'); if (!t || t.querySelector('input')) return;
  const i = +t.dataset.i;
  if (i !== state.cur.active) switchSheet(i); else if ($('#tabMenu').hidden) openTabMenu(t, i); else closeMini();
});
$('#tabList').addEventListener('dblclick', e => { const t = e.target.closest('.tab'); if (t && !t.querySelector('input')) startRename(+t.dataset.i); });
$('#tabAdd').addEventListener('click', () => addSheet());
$('#tabMenu').addEventListener('click', e => {
  const b = e.target.closest('[data-tm]'); if (!b) return;
  const i = +$('#tabMenu').dataset.i, k = b.dataset.tm;
  if (k === 'rename') startRename(i);
  else if (k === 'dup') addSheet(state.cur.sheets[i]);
  else if (k === 'left') moveSheet(i, -1);
  else if (k === 'right') moveSheet(i, 1);
  else if (k === 'pivot') { closeMini(); refreshPivot(i); }
  else if (k === 'del') { if (!SH.tabDel) { SH.tabDel = true; b.querySelector('span').textContent = 'Confirmer la suppression'; } else deleteSheet(i); }
});

/* --- couleurs, bordures, figer */
const COLORS = ['jaune', 'vert', 'bleu', 'rouge', 'violet', 'orange', 'gris', 'cuivre'];
const COLOR_FR = { jaune: 'Jaune', vert: 'Vert', bleu: 'Bleu', rouge: 'Rouge', violet: 'Violet', orange: 'Orange', gris: 'Gris', cuivre: 'Cuivre' };
function openFmt(kind) {
  commitEdit(false);
  const p = $('#fmtPop'), anchor = $(`[data-s="${{ bg: 'fill', fc: 'color', bd: 'borders', frz: 'freeze' }[kind]}"]`);
  if (!p.hidden && p.dataset.k === kind) return closeMini();
  closeMini(); p.dataset.k = kind;
  if (kind === 'bg' || kind === 'fc') {
    p.innerHTML = `<div class="mp-t">${kind === 'bg' ? 'Couleur de fond' : 'Couleur du texte'}</div><div class="swatches">`
      + COLORS.map(c => `<button class="sw" data-col="${c}" title="${COLOR_FR[c]}" aria-label="${COLOR_FR[c]}" style="${kind === 'bg' ? `background:var(--bg-${c})` : `color:var(--fc-${c})`}">${kind === 'fc' ? 'A' : ''}</button>`).join('')
      + `<button class="sw none" data-col="" title="Aucune" aria-label="Aucune couleur">Aucune</button></div>`;
  } else if (kind === 'bd') {
    p.innerHTML = `<div class="mp-t">Bordures</div><button data-bd="all">${ic('grid')}Toutes les bordures</button><button data-bd="outer">${ic('borders')}Contour</button><button data-bd="bottom">${ic('hr')}Bordure en bas</button><button data-bd="none">${ic('close')}Aucune bordure</button>`;
  } else {
    const F = state.cur.frz || {};
    p.innerHTML = `<div class="mp-t">Figer</div><button data-fz="1,${F.c || 0}" aria-pressed="${F.r === 1}">${ic('freeze')}La première ligne</button><button data-fz="2,${F.c || 0}" aria-pressed="${F.r === 2}">${ic('freeze')}Les 2 premières lignes</button><button data-fz="${F.r || 0},1" aria-pressed="${F.c === 1}">${ic('freeze')}La première colonne</button><hr><button data-fz="0,0">${ic('close')}Ne rien figer</button>`;
  }
  placePop(p, anchor);
}
function setBorders(mode) {
  commitEdit(false); pushHist();
  const R = rng();
  eachInRange((r, c) => {
    const k = addr(r, c), x = state.cur.cells[k] || (state.cur.cells[k] = { v: '', f: {} }); x.f ||= {};
    const set = (side, on) => { if (on) x.f[side] = 1; else delete x.f[side]; };
    if (mode === 'none') ['bt', 'bb', 'bl', 'br'].forEach(sd => set(sd, false));
    else if (mode === 'all') { set('bt', true); set('bl', true); if (r === R.r2) set('bb', true); if (c === R.c2) set('br', true); }
    else if (mode === 'outer') { if (r === R.r1) set('bt', true); if (r === R.r2) set('bb', true); if (c === R.c1) set('bl', true); if (c === R.c2) set('br', true); }
    else if (mode === 'bottom' && r === R.r2) set('bb', true);
    if (!x.v && !Object.keys(x.f).length) delete state.cur.cells[k];
  });
  afterChange();
}
function setFreeze(r, c) {
  state.cur.frz = r || c ? { r: r || 0, c: c || 0 } : undefined;
  applyFreeze(); drawSel(); updateSheetTools(); scheduleSave();
  toast(r || c ? 'Volets figés' : 'Plus rien n\'est figé');
}
$('#fmtPop').addEventListener('mousedown', e => e.preventDefault());
$('#fmtPop').addEventListener('click', e => {
  const p = $('#fmtPop'), k = p.dataset.k;
  const sw = e.target.closest('[data-col]'); if (sw) { closeMini(); setFmt(k, sw.dataset.col); gw.focus({ preventScroll: true }); return; }
  const bd = e.target.closest('[data-bd]'); if (bd) { closeMini(); setBorders(bd.dataset.bd); gw.focus({ preventScroll: true }); return; }
  const fz = e.target.closest('[data-fz]'); if (fz) { closeMini(); const [a, b] = fz.dataset.fz.split(',').map(Number); setFreeze(a, b); gw.focus({ preventScroll: true }); }
});
document.addEventListener('mousedown', e => {
  if ($('#tabMenu').hidden && $('#fmtPop').hidden) return;
  if (e.target.closest('#tabMenu, #fmtPop, .tab, [data-s="fill"], [data-s="color"], [data-s="borders"], [data-s="freeze"], [data-s="insMenu"], [data-s="dataMenu"], [data-s="structMenu"]')) return;
  closeMini();
});

/* --- fenêtres du classeur : listes, cases, notes, mise en forme conditionnelle, plages nommées, recherche */
const sDlg = $('#sDlg');
function rangeObj(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(String(ref || '').trim());
  if (!m) return null;
  const a = parseAddr(m[1].toUpperCase() + m[2]), b = m[3] ? parseAddr(m[3].toUpperCase() + m[4]) : a;
  if (!a || !b) return null;
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}
const absRef = R => '$' + COLN(R.c1) + '$' + (R.r1 + 1) + (R.r1 === R.r2 && R.c1 === R.c2 ? '' : ':$' + COLN(R.c2) + '$' + (R.r2 + 1));
function openDlg(html, cls) {
  closeMini();
  sDlg.className = 'fpop sdlg' + (cls ? ' ' + cls : '');
  sDlg.innerHTML = `<div class="dlg-in">${html}</div>`; sDlg.hidden = false;
  const f = sDlg.querySelector('[autofocus]'); if (f && !isPhone()) setTimeout(() => f.focus(), 30);
  return sDlg.firstElementChild;
}
function closeDlg() { if (sDlg.hidden) return; sDlg.hidden = true; sDlg.innerHTML = ''; SH.dlg = null; if (state.view === 'sheet') gw.focus({ preventScroll: true }); }
const dlgHead = t => `<div class="fp-h"><b>${t}</b><button class="xbtn" data-dx aria-label="Fermer">${ic('close')}</button></div>`;
sDlg.addEventListener('click', e => { if (e.target.closest('[data-dx]')) closeDlg(); });
sDlg.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeDlg(); } });

/* listes déroulantes et cases à cocher */
function ddValues(f) {
  if (!f || !f.dd) return [];
  if (Array.isArray(f.dd)) return f.dd;
  const src = String(f.dd);
  if (src[0] === '=' && SH.hf) {
    try {
      const res = SH.hf.calculateFormula('=' + (SH.fr ? frNamesToEn(src.slice(1)) : frToExcel(src.slice(1))), SH.sid);
      const flat = Array.isArray(res) ? res.flat() : [res];
      return [...new Set(flat.filter(v => v != null && v !== '' && typeof v !== 'object').map(v => typeof v === 'number' ? numToRaw(v) : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v)))];
    } catch { return []; }
  }
  return [];
}
function openDdEditor() {
  commitEdit(false);
  const f = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {};
  const R = rng();
  const isRef = f.dd && !Array.isArray(f.dd);
  SH.dlg = 'dd';
  openDlg(dlgHead('Liste déroulante') + `
    <p class="dlg-note">Pour ${rngRef(R)} : on choisit la valeur dans une liste, et une saisie hors liste est refusée.</p>
    <div class="seg dd-src" role="group"><button type="button" data-src="list" aria-pressed="${!isRef}">Éléments</button><button type="button" data-src="ref" aria-pressed="${!!isRef}">Plage de cellules</button></div>
    <label class="dlg-f" ${isRef ? 'hidden' : ''} id="ddListF"><span>Un élément par ligne</span><textarea id="ddList" rows="6" autofocus>${esc(Array.isArray(f.dd) ? f.dd.join('\n') : '')}</textarea></label>
    <label class="dlg-f" ${isRef ? '' : 'hidden'} id="ddRefF"><span>Plage (ex. A2:A10 ou Listes!A1:A20)</span><input id="ddRef" value="${esc(isRef ? String(f.dd).slice(1) : '')}" autocomplete="off" autocapitalize="off" spellcheck="false"></label>
    <div class="fp-acts"><button class="pbtn" id="ddOk">Appliquer</button>${f.dd ? '<button class="pbtn ghost" id="ddDel">Retirer la liste</button>' : ''}</div>`);
  sDlg.querySelector('.dd-src').addEventListener('click', e => {
    const b = e.target.closest('[data-src]'); if (!b) return;
    sDlg.querySelectorAll('[data-src]').forEach(x => x.setAttribute('aria-pressed', String(x === b)));
    $('#ddListF').hidden = b.dataset.src !== 'list'; $('#ddRefF').hidden = b.dataset.src !== 'ref';
  });
  $('#ddOk').addEventListener('click', () => {
    const ref = !$('#ddRefF').hidden;
    let val;
    if (ref) { const t = $('#ddRef').value.trim().replace(/^=/, ''); if (!t) return toast('Indiquez une plage.', true); val = '=' + t; }
    else { val = $('#ddList').value.split('\n').map(x => x.trim()).filter(Boolean); if (!val.length) return toast('Ajoutez au moins un élément.', true); val = [...new Set(val)]; }
    applyToRange(f2 => { f2.dd = val; delete f2.chk; });
    closeDlg(); toast('Liste déroulante ajoutée');
  });
  const del = $('#ddDel'); if (del) del.addEventListener('click', () => { applyToRange(f2 => { delete f2.dd; }); closeDlg(); toast('Liste retirée'); });
}
function applyToRange(fn) {
  pushHist();
  eachInRange((r, c) => { const k = addr(r, c), x = state.cur.cells[k] || (state.cur.cells[k] = { v: '', f: {} }); x.f ||= {}; fn(x.f, r, c, x); if (!x.v && !Object.keys(x.f).length) delete state.cur.cells[k]; });
  afterChange();
}
function toggleCheckboxFmt() {
  commitEdit(false);
  const f = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {};
  const on = !f.chk;
  pushHist();
  batch(() => eachInRange((r, c) => {
    const k = addr(r, c), x = state.cur.cells[k] || (state.cur.cells[k] = { v: '', f: {} }); x.f ||= {};
    if (on) { x.f.chk = 1; delete x.f.dd; const v = String(x.v || '').trim(); if (!/^(VRAI|FAUX)$/i.test(v)) applyRaw(r, c, 'FAUX'); }
    else { delete x.f.chk; }
    if (!x.v && !Object.keys(x.f).length) delete state.cur.cells[k];
  }));
  afterChange();
  toast(on ? 'Cases à cocher ajoutées : touchez une case pour la cocher.' : 'Cases à cocher retirées');
}
function toggleCheck(r, c) {
  const x = cellOf(r, c); if (!x || !x.f || !x.f.chk) return false;
  const v = (SH.vals[r] || [])[c];
  setRaw(r, c, v === true ? 'FAUX' : 'VRAI');
  return true;
}
function openDdPick(r, c) {
  const f = (cellOf(r, c) || {}).f || {};
  const vals = ddValues(f); if (!f.dd) return;
  const cur = String(rawOf(r, c));
  const p = $('#ddPop');
  p.innerHTML = (vals.length ? vals.map(v => `<button data-v="${esc(v)}" aria-pressed="${v === cur}">${esc(v)}</button>`).join('') : '<div class="mp-t">Liste vide</div>') + (cur ? `<hr><button data-v="" class="muted">Effacer</button>` : '');
  p.hidden = false;
  const td = tdOf(r, c), host = $('#sheet').getBoundingClientRect(), a = td.getBoundingClientRect();
  p.style.minWidth = Math.max(160, a.width) + 'px';
  p.style.left = Math.max(8, Math.min(a.left - host.left, host.width - p.offsetWidth - 8)) + 'px';
  const below = a.bottom - host.top + 2;
  if (below + p.offsetHeight > host.height - 8) { p.style.top = 'auto'; p.style.bottom = (host.bottom - a.top + 2) + 'px'; } else { p.style.bottom = 'auto'; p.style.top = below + 'px'; }
  p.dataset.r = r; p.dataset.c = c;
}
function closeDdPick() { $('#ddPop').hidden = true; }
$('#ddPop').addEventListener('mousedown', e => e.preventDefault());
$('#ddPop').addEventListener('click', e => {
  const b = e.target.closest('[data-v]'); if (!b) return;
  const p = $('#ddPop'), r = +p.dataset.r, c = +p.dataset.c;
  closeDdPick(); setRaw(r, c, b.dataset.v); gw.focus({ preventScroll: true });
});
document.addEventListener('mousedown', e => { if (!$('#ddPop').hidden && !e.target.closest('#ddPop')) closeDdPick(); }, true);
// Saisie refusée hors de la liste
function ddReject(r, c, raw) {
  const f = (cellOf(r, c) || {}).f || {};
  if (!f.dd || raw === '' || String(raw)[0] === '=') return false;
  const vals = ddValues(f);
  if (!vals.length || vals.some(v => v.toLowerCase() === String(raw).trim().toLowerCase())) return false;
  toast('Cette valeur n\'est pas dans la liste : choisissez-la avec ▾.', true);
  return true;
}

/* notes */
function openNoteEditor() {
  commitEdit(false);
  const r = SH.sel.r, c = SH.sel.c, f = (cellOf(r, c) || {}).f || {};
  SH.dlg = 'note';
  openDlg(dlgHead('Note sur ' + addr(r, c)) + `<label class="dlg-f"><span>Texte de la note</span><textarea id="noteTx" rows="5" autofocus placeholder="Une remarque, une source, un rappel…">${esc(f.note || '')}</textarea></label>
    <div class="fp-acts"><button class="pbtn" id="noteOk">Enregistrer</button>${f.note ? '<button class="pbtn ghost" id="noteDel">Supprimer la note</button>' : ''}</div>`);
  $('#noteOk').addEventListener('click', () => {
    const t = $('#noteTx').value.trim();
    pushHist();
    const k = addr(r, c), x = state.cur.cells[k] || (state.cur.cells[k] = { v: '', f: {} }); x.f ||= {};
    if (t) x.f.note = t.slice(0, 2000); else delete x.f.note;
    if (!x.v && !Object.keys(x.f).length) delete state.cur.cells[k];
    afterChange(); closeDlg(); toast(t ? 'Note enregistrée' : 'Note supprimée');
  });
  const d = $('#noteDel'); if (d) d.addEventListener('click', () => { $('#noteTx').value = ''; $('#noteOk').click(); });
}

/* mise en forme conditionnelle */
const cfStyleLabel = st => [st.bg ? 'fond ' + COLOR_FR[st.bg].toLowerCase() : '', st.fc ? 'texte ' + COLOR_FR[st.fc].toLowerCase() : '', st.b ? 'gras' : ''].filter(Boolean).join(', ');
function cfDescr(rule) {
  const R = rngRef(rule);
  if (rule.t === 'scale') return `${R} · échelle ${rule.mid ? COLOR_FR[rule.lo || 'rouge'] + ' → ' + COLOR_FR[rule.mid] + ' → ' : (rule.lo ? COLOR_FR[rule.lo] : 'Blanc') + ' → '}${COLOR_FR[rule.hi || 'vert']}`;
  const n = (CF_TYPES.find(t => t[0] === rule.t) || [])[2];
  const val = n === 2 ? ` ${rule.a} et ${rule.b}` : n === 1 ? ` « ${rule.a} »` : '';
  return `${R} · ${CF_LABEL[rule.t]}${val} → ${cfStyleLabel(rule.st || {}) || 'sans style'}`;
}
function openCF(editId) {
  commitEdit(false);
  const cur = state.cur; cur.cf ||= [];
  SH.dlg = 'cf';
  const rule = editId ? cur.cf.find(x => x.id === editId) : null;
  const R = rule || rng();
  const list = cur.cf.map((x, i) => `<div class="cf-it"><span class="cf-sw" style="${x.t === 'scale' ? `background:linear-gradient(90deg,var(--bg-${x.lo || 'gris'}),${x.mid ? `var(--bg-${x.mid}),` : ''}var(--bg-${x.hi || 'vert'}))` : `background:${x.st && x.st.bg ? `var(--bg-${x.st.bg})` : 'var(--paper)'};color:${x.st && x.st.fc ? `var(--fc-${x.st.fc})` : 'inherit'}`}">${x.t === 'scale' ? '' : 'Ab'}</span><span class="cf-d">${esc(cfDescr(x))}</span><button class="xbtn" data-cfe="${x.id}" title="Modifier" aria-label="Modifier">${ic('pen')}</button><button class="xbtn" data-cfd="${x.id}" title="Supprimer" aria-label="Supprimer">${ic('trash')}</button></div>`).join('');
  const st = (rule && rule.st) || { bg: 'rouge', fc: '', b: 0 };
  const sw = (k, sel) => `<div class="swatches sm" data-k="${k}">${COLORS.map(c => `<button type="button" class="sw" data-col="${c}" aria-pressed="${sel === c}" title="${COLOR_FR[c]}" style="${k === 'fc' ? `color:var(--fc-${c})` : `background:var(--bg-${c})`}">${k === 'fc' ? 'A' : ''}</button>`).join('')}<button type="button" class="sw none" data-col="" aria-pressed="${!sel}" title="Aucune" aria-label="Aucune">∅</button></div>`;
  const box = openDlg(dlgHead('Mise en forme conditionnelle') + (list ? `<div class="cf-list">${list}</div><hr class="dlg-hr">` : '') + `
    <b class="dlg-sub">${rule ? 'Modifier la règle' : 'Nouvelle règle'}</b>
    <label class="dlg-f"><span>Plage</span><input id="cfRange" value="${esc(rngRef(R))}" autocomplete="off" autocapitalize="off" spellcheck="false"></label>
    <label class="dlg-f"><span>Mettre en forme si…</span><select id="cfType">${CF_TYPES.map(t => `<option value="${t[0]}"${(rule ? rule.t : 'gt') === t[0] ? ' selected' : ''}>${t[1]}</option>`).join('')}</select></label>
    <div class="dlg-row" id="cfVals"><input id="cfA" placeholder="Valeur" value="${esc(rule && rule.a != null ? rule.a : '')}" autocomplete="off"><span id="cfAnd">et</span><input id="cfB" placeholder="Valeur" value="${esc(rule && rule.b != null ? rule.b : '')}" autocomplete="off"></div>
    <p class="dlg-note" id="cfFHelp" hidden>Formule écrite pour la première cellule de la plage, par exemple <code>=$C2&gt;100</code> pour colorer toute la ligne.</p>
    <div id="cfStyle"><span class="dlg-lbl">Couleur de fond</span>${sw('bg', st.bg)}<span class="dlg-lbl">Couleur du texte</span>${sw('fc', st.fc)}<label class="fp-it"><input type="checkbox" id="cfBold"${st.b ? ' checked' : ''}><span>Gras</span></label></div>
    <div id="cfScale" hidden><span class="dlg-lbl">Échelle</span><div class="seg" id="cfScaleSeg">
      <button type="button" data-sc=",,vert">Blanc → vert</button><button type="button" data-sc="rouge,jaune,vert">Rouge → vert</button><button type="button" data-sc="vert,jaune,rouge">Vert → rouge</button><button type="button" data-sc=",,bleu">Blanc → bleu</button></div></div>
    <div class="cf-prev" id="cfPrev">Aperçu</div>
    <div class="fp-acts"><button class="pbtn" id="cfOk">${rule ? 'Enregistrer' : 'Ajouter la règle'}</button>${rule ? '<button class="pbtn ghost" id="cfNew">Nouvelle règle</button>' : ''}</div>`, 'wide');
  const S2 = { bg: st.bg || '', fc: st.fc || '', sc: rule && rule.t === 'scale' ? [rule.lo || '', rule.mid || '', rule.hi || 'vert'].join(',') : 'rouge,jaune,vert' };
  const refresh = () => {
    const t = $('#cfType').value, n = (CF_TYPES.find(x => x[0] === t) || [])[2];
    $('#cfVals').hidden = !n || t === 'scale'; $('#cfB').hidden = n !== 2; $('#cfAnd').hidden = n !== 2;
    $('#cfA').placeholder = t === 'formula' ? '=$C2>100' : 'Valeur';
    $('#cfFHelp').hidden = t !== 'formula';
    $('#cfStyle').hidden = t === 'scale'; $('#cfScale').hidden = t !== 'scale';
    sDlg.querySelectorAll('#cfScaleSeg [data-sc]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.sc === S2.sc)));
    sDlg.querySelectorAll('.swatches[data-k] .sw').forEach(b => b.setAttribute('aria-pressed', String((S2[b.parentElement.dataset.k] || '') === b.dataset.col)));
    const pv = $('#cfPrev');
    if (t === 'scale') { const [lo, mid, hi] = S2.sc.split(','); pv.style.cssText = `background:linear-gradient(90deg,${lo ? `var(--bg-${lo})` : 'var(--paper)'},${mid ? `var(--bg-${mid}),` : ''}var(--bg-${hi}))`; pv.textContent = 'Plus petite valeur → plus grande'; }
    else { pv.style.cssText = `background:${S2.bg ? `var(--bg-${S2.bg})` : 'var(--paper)'};color:${S2.fc ? `var(--fc-${S2.fc})` : 'inherit'};font-weight:${$('#cfBold').checked ? 600 : 400}`; pv.textContent = 'Aperçu 1 234,00'; }
  };
  box.addEventListener('change', refresh);
  box.addEventListener('click', e => {
    const s1 = e.target.closest('.swatches[data-k] .sw'); if (s1) { S2[s1.parentElement.dataset.k] = s1.dataset.col; refresh(); return; }
    const sc = e.target.closest('[data-sc]'); if (sc) { S2.sc = sc.dataset.sc; refresh(); return; }
    const ed = e.target.closest('[data-cfe]'); if (ed) { openCF(ed.dataset.cfe); return; }
    const dl = e.target.closest('[data-cfd]'); if (dl) { pushHist(); cur.cf = cur.cf.filter(x => x.id !== dl.dataset.cfd); afterChange(); openCF(); return; }
  });
  refresh();
  const nw = $('#cfNew'); if (nw) nw.addEventListener('click', () => openCF());
  $('#cfOk').addEventListener('click', () => {
    const o = rangeObj($('#cfRange').value.replace(/^=/, ''));
    if (!o) return toast('Plage incorrecte (exemple : B2:B20).', true);
    const t = $('#cfType').value, n = (CF_TYPES.find(x => x[0] === t) || [])[2];
    const nr = { id: rule ? rule.id : uid(), ...o, t };
    if (t === 'scale') { const [lo, mid, hi] = S2.sc.split(','); if (lo) nr.lo = lo; if (mid) nr.mid = mid; nr.hi = hi || 'vert'; }
    else {
      if (n >= 1) { nr.a = $('#cfA').value.trim(); if (nr.a === '') return toast('Indiquez une valeur.', true); }
      if (n === 2) { nr.b = $('#cfB').value.trim(); if (nr.b === '') return toast('Indiquez la seconde valeur.', true); }
      if (t === 'formula' && nr.a[0] !== '=') nr.a = '=' + nr.a;
      nr.st = {}; if (S2.bg) nr.st.bg = S2.bg; if (S2.fc) nr.st.fc = S2.fc; if ($('#cfBold').checked) nr.st.b = 1;
      if (!Object.keys(nr.st).length) return toast('Choisissez au moins une couleur ou le gras.', true);
    }
    pushHist();
    if (rule) cur.cf[cur.cf.findIndex(x => x.id === rule.id)] = nr; else cur.cf.push(nr);
    afterChange(); toast(rule ? 'Règle modifiée' : 'Règle ajoutée'); openCF();
  });
}

/* plages nommées */
const NAME_RE = /^[A-Za-zÀ-ÿ_][A-Za-z0-9À-ÿ_.]*$/;
function nameOk(n) { return NAME_RE.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^(VRAI|FAUX|TRUE|FALSE)$/i.test(n) && !/^R\d*C\d*$/i.test(n) && n.length <= 60; }
function nameFormula(nm) { return `${refName(nm.s)}!${absRef(rangeObj(nm.ref))}`; }
function loadNames() {
  if (!SH.hf) return;
  const hf = SH.hf;
  try { hf.listNamedExpressions().forEach(n => { try { hf.removeNamedExpression(n); } catch {} }); } catch {}
  Object.entries(state.cur.names || {}).forEach(([n, nm]) => {
    if (!state.cur.sheets.some(s => s.name === nm.s) || !rangeObj(nm.ref)) return;
    try { hf.addNamedExpression(n, '=' + nameFormula(nm)); } catch (e) { console.warn('Nom ignoré', n, e); }
  });
}
function defineName(n, sheetName, ref) {
  n = n.trim();
  if (!nameOk(n)) { toast('Nom refusé : lettres, chiffres, _ et point, sans ressembler à une cellule (ex. Total_ventes).', true); return false; }
  const ex = Object.keys(state.cur.names).find(k => k.toLowerCase() === n.toLowerCase());
  if (ex && ex !== n) delete state.cur.names[ex];
  state.cur.names[n] = { s: sheetName, ref };
  loadHF(); renderValues(); scheduleSave();
  return true;
}
function openNames() {
  commitEdit(false);
  SH.dlg = 'names';
  const cur = state.cur, R = rng(), sname = cur.sheets[cur.active].name;
  const rows = Object.entries(cur.names).sort((a, b) => a[0].localeCompare(b[0], 'fr')).map(([n, nm]) => `<div class="cf-it"><b class="nm-n">${esc(n)}</b><button class="linkbtn nm-go" data-go="${esc(n)}">${esc(nm.s)}!${esc(nm.ref)}</button><button class="xbtn" data-nd="${esc(n)}" aria-label="Supprimer">${ic('trash')}</button></div>`).join('');
  openDlg(dlgHead('Plages nommées') + `<p class="dlg-note">Donnez un nom à une plage pour l'utiliser dans les formules : <code>=SOMME(Ventes)</code>.</p>`
    + (rows ? `<div class="cf-list">${rows}</div><hr class="dlg-hr">` : '')
    + `<label class="dlg-f"><span>Nom</span><input id="nmName" placeholder="Ventes" autocomplete="off" autocapitalize="off" spellcheck="false" autofocus></label>
       <label class="dlg-f"><span>Plage sur « ${esc(sname)} »</span><input id="nmRef" value="${esc(rngRef(R))}" autocomplete="off" autocapitalize="off" spellcheck="false"></label>
       <div class="fp-acts"><button class="pbtn" id="nmOk">Nommer la plage</button></div>`);
  $('#nmOk').addEventListener('click', () => {
    const o = rangeObj($('#nmRef').value);
    if (!o) return toast('Plage incorrecte (exemple : B2:B20).', true);
    if (defineName($('#nmName').value, sname, rngRef(o))) { toast('Nom créé'); openNames(); }
  });
  sDlg.querySelectorAll('[data-nd]').forEach(b => b.addEventListener('click', () => { delete cur.names[b.dataset.nd]; loadHF(); renderValues(); scheduleSave(); openNames(); }));
  sDlg.querySelectorAll('[data-go]').forEach(b => b.addEventListener('click', () => { closeDlg(); goName(b.dataset.go); }));
}
function goName(n) {
  const cur = state.cur, key = Object.keys(cur.names).find(k => k.toLowerCase() === n.toLowerCase()); if (!key) return false;
  const nm = cur.names[key], i = cur.sheets.findIndex(s => s.name === nm.s), o = rangeObj(nm.ref);
  if (i < 0 || !o) return false;
  if (i !== cur.active) switchSheet(i);
  SH.anc = { r: o.r2, c: o.c2 }; select(o.r1, o.c1, true);
  return true;
}
// Zone de nom (à gauche de la barre de formule) : aller à une cellule, une plage, un nom, ou nommer la sélection
const fxRef = $('#fxRef');
let fxRefJust = false;
fxRef.addEventListener('focus', () => { fxRef.select(); fxRefJust = true; });
fxRef.addEventListener('mouseup', e => { if (fxRefJust) { e.preventDefault(); fxRefJust = false; } });
fxRef.addEventListener('input', () => { fxRefJust = false; });
fxRef.addEventListener('keydown', e => {
  if (e.key === 'Escape') { e.preventDefault(); updateFx(); gw.focus({ preventScroll: true }); return; }
  if (e.key !== 'Enter') return;
  e.preventDefault();
  const t = fxRef.value.trim(); if (!t) return;
  const m = /^(?:(.+)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(t);
  if (m) {
    const cur = state.cur;
    if (m[1]) { const i = cur.sheets.findIndex(s => s.name.toLowerCase() === m[1].replace(/^'|'$/g, '').toLowerCase()); if (i < 0) return toast('Feuille introuvable.', true); if (i !== cur.active) switchSheet(i); }
    const o = rangeObj(m[2]); ensureSize(o.r2, o.c2);
    SH.anc = { r: o.r2, c: o.c2 }; select(o.r1, o.c1, true); gw.focus({ preventScroll: true }); return;
  }
  if (goName(t)) { gw.focus({ preventScroll: true }); return; }
  if (defineName(t, state.cur.sheets[state.cur.active].name, rngRef(rng()))) { toast(`« ${t} » désigne maintenant ${rngRef(rng())}`); gw.focus({ preventScroll: true }); }
});

/* rechercher et remplacer dans le classeur */
const SF = { hits: [], i: -1 };
function openSFind(withRep) {
  commitEdit(false); closeDlg();
  const bar = $('#sFind'); bar.hidden = false;
  $('#sfRepRow').hidden = !withRep;
  const q = $('#sfQ'); q.focus(); q.select();
  sfRun();
}
function closeSFind() { $('#sFind').hidden = true; SF.hits = []; SF.i = -1; $('#sfCount').textContent = ''; gw.focus({ preventScroll: true }); }
function sfText(x, v, inF) {
  if (inF || (x && String(x.v)[0] !== '=')) return String(x ? x.v : '');
  return v == null ? '' : typeof v === 'object' ? (v.value || '') : typeof v === 'number' ? String(v).replace('.', ',') : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v);
}
function sfRun(keepPos) {
  const q = $('#sfQ').value, all = $('#sfAll').checked, mc = $('#sfCase').checked, whole = $('#sfWhole').checked, inF = $('#sfForm').checked;
  SF.hits = [];
  if (!q) { $('#sfCount').textContent = ''; return; }
  const cur = state.cur, needle = mc ? q : q.toLowerCase();
  const sheets = all ? cur.sheets.map((s, i) => i) : [cur.active];
  sheets.forEach(si => {
    const sh = cur.sheets[si];
    const vals = SH.hf && SH.ids ? SH.hf.getSheetValues(SH.ids[si]) : [];
    Object.entries(sh.cells).forEach(([k, x]) => {
      if (x.v === '' || x.v == null) return;
      const p = parseAddr(k); if (!p) return;
      let t = sfText(x, (vals[p.r] || [])[p.c], inF); if (!mc) t = t.toLowerCase();
      if (whole ? t === needle : t.includes(needle)) SF.hits.push({ si, r: p.r, c: p.c });
    });
  });
  SF.hits.sort((a, b) => a.si - b.si || a.r - b.r || a.c - b.c);
  if (!keepPos) { const cur2 = { si: state.cur.active, r: SH.sel.r, c: SH.sel.c }; SF.i = SF.hits.findIndex(h => h.si > cur2.si || (h.si === cur2.si && (h.r > cur2.r || (h.r === cur2.r && h.c >= cur2.c)))); if (SF.i < 0) SF.i = 0; }
  SF.i = Math.min(SF.i, SF.hits.length - 1);
  sfShow();
}
function sfShow() {
  $('#sfCount').textContent = SF.hits.length ? `${SF.i + 1} sur ${SF.hits.length}` : 'Aucun résultat';
  const h = SF.hits[SF.i]; if (!h) return;
  if (h.si !== state.cur.active) switchSheet(h.si);
  SH.anc = { r: h.r, c: h.c }; select(h.r, h.c);
}
function sfStep(d) { if (!SF.hits.length) return sfRun(); SF.i = (SF.i + d + SF.hits.length) % SF.hits.length; sfShow(); }
function sfReplaceIn(raw, q, rep, mc, whole) {
  if (whole) return rep;
  const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), mc ? 'g' : 'gi');
  return String(raw).replace(re, () => rep);
}
function sfReplace(all) {
  const q = $('#sfQ').value; if (!q) return;
  const rep = $('#sfR').value, mc = $('#sfCase').checked, whole = $('#sfWhole').checked;
  if (!SF.hits.length) sfRun();
  const todo = all ? SF.hits.slice() : SF.hits[SF.i] ? [SF.hits[SF.i]] : [];
  if (!todo.length) return;
  pushHist();
  let n = 0;
  const cur = state.cur;
  todo.forEach(h => {
    const sh = cur.sheets[h.si], k = addr(h.r, h.c), x = sh.cells[k]; if (!x) return;
    const isF = String(x.v)[0] === '=';
    if (isF && !$('#sfForm').checked) return;
    const nv = sfReplaceIn(x.v, q, rep, mc, whole); if (nv === x.v) return;
    n++;
    if (h.si === cur.active) applyRaw(h.r, h.c, nv); else { x.v = normRaw(nv); }
  });
  if (todo.some(h => h.si !== cur.active)) loadHF();
  afterChange();
  toast(n ? `${n} remplacement${n > 1 ? 's' : ''}` : 'Rien à remplacer (les formules ne sont modifiées que si « Dans les formules » est coché).');
  sfRun(true);
  if (!all && SF.hits.length) sfShow();
}
$('#sfQ').addEventListener('input', () => sfRun());
['sfAll', 'sfCase', 'sfWhole', 'sfForm'].forEach(id => $('#' + id).addEventListener('change', () => sfRun()));
$('#sfQ').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); sfStep(e.shiftKey ? -1 : 1); } else if (e.key === 'Escape') { e.preventDefault(); closeSFind(); } });
$('#sfR').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); sfReplace(false); } else if (e.key === 'Escape') { e.preventDefault(); closeSFind(); } });
$('#sFind').addEventListener('click', e => {
  const b = e.target.closest('[data-sf]'); if (!b) return;
  ({ prev: () => sfStep(-1), next: () => sfStep(1), one: () => sfReplace(false), all: () => sfReplace(true), close: closeSFind, rep: () => { $('#sfRepRow').hidden = !$('#sfRepRow').hidden; } })[b.dataset.sf]();
});

/* menus Insérer et Données */
function openSheetMenu(kind) {
  commitEdit(false);
  const p = $('#fmtPop'), anchor = $(`[data-s="${{ ins: 'insMenu', data: 'dataMenu', struct: 'structMenu' }[kind]}"]`);
  if (!p.hidden && p.dataset.k === kind) return closeMini();
  closeMini(); p.dataset.k = kind;
  const f = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {};
  const cur = state.cur;
  const F = cur.frz || {};
  p.innerHTML = kind === 'struct'
    ? `<div class="mp-t">Lignes et colonnes</div><button data-m="rowIns">${ic('rowIns')}Insérer une ligne au-dessus</button><button data-m="rowDel">${ic('rowDel')}Supprimer la ligne</button><button data-m="colIns">${ic('colIns')}Insérer une colonne à gauche</button><button data-m="colDel">${ic('colDel')}Supprimer la colonne</button><hr>`
      + `<button data-m="hideR">${ic('eyeOff')}Masquer les lignes</button><button data-m="hideC">${ic('eyeOff')}Masquer les colonnes</button>`
      + (cur.hid && cur.hid.r.length ? `<button data-m="showR">${ic('eye')}Afficher les ${cur.hid.r.length} ligne${cur.hid.r.length > 1 ? 's' : ''} masquée${cur.hid.r.length > 1 ? 's' : ''}</button>` : '')
      + (cur.hid && cur.hid.c.length ? `<button data-m="showC">${ic('eye')}Afficher les ${cur.hid.c.length} colonne${cur.hid.c.length > 1 ? 's' : ''} masquée${cur.hid.c.length > 1 ? 's' : ''}</button>` : '')
      + `<hr><button data-m="frz1" aria-pressed="${F.r === 1}">${ic('freeze')}Figer la première ligne</button><button data-m="frzC" aria-pressed="${F.c === 1}">${ic('freeze')}Figer la première colonne</button>${F.r || F.c ? `<button data-m="frz0">${ic('close')}Ne rien figer</button>` : ''}`
      + `<hr><button data-m="fillDown">${ic('fillDown')}Recopier vers le bas<small>Ctrl D</small></button><button data-m="clear">${ic('eraser')}Effacer le contenu<small>Suppr</small></button>`
    : kind === 'ins'
    ? `<div class="mp-t">Insérer</div><button data-m="chart">${ic('chart')}Graphique</button><button data-m="pivot">${ic('pivot')}Tableau croisé dynamique</button><hr>`
      + `<button data-m="dd">${ic('list')}${f.dd ? 'Modifier la liste déroulante' : 'Liste déroulante'}</button><button data-m="chk">${ic('check')}${f.chk ? 'Retirer les cases à cocher' : 'Case à cocher'}</button><button data-m="note">${ic('note')}${f.note ? 'Modifier la note' : 'Note'}</button><hr><button data-m="fn">${ic('sigma')}Fonction…</button>`
    : `<div class="mp-t">Données</div><button data-m="cf">${ic('cond')}Mise en forme conditionnelle</button><button data-m="find">${ic('search')}Rechercher et remplacer<small>Ctrl H</small></button><button data-m="names">${ic('tag')}Plages nommées</button>`
      + `<hr><button data-m="sortAsc">${ic('sortAsc')}Trier de A à Z</button><button data-m="sortDesc">${ic('sortDesc')}Trier de Z à A</button><button data-m="filter">${ic('filter')}Filtrer la colonne</button>`
      + (cur.filter ? `<button data-m="unfilter">${ic('filter')}Retirer tous les filtres</button>` : '');
  placePop(p, anchor);
}
const SMENU = {
  chart: () => openChartEditor(), pivot: () => openPivot(), dd: openDdEditor, chk: toggleCheckboxFmt, note: openNoteEditor,
  fn: () => { if (!SH.editing) startEdit('=', false); $('#fxKeys').hidden = false; $('#fnList').hidden = false; $('#fxBtn').setAttribute('aria-pressed', 'true'); requestAnimationFrame(drawSel); },
  cf: () => openCF(), find: () => openSFind(true), names: openNames,
  hideR: () => setHidden('r', true), hideC: () => setHidden('c', true), showR: () => setHidden('r', false), showC: () => setHidden('c', false), unfilter: clearFilters,
  rowIns: () => structural('rowIns'), rowDel: () => structural('rowDel'), colIns: () => structural('colIns'), colDel: () => structural('colDel'),
  frz1: () => setFreeze(1, (state.cur.frz || {}).c || 0), frzC: () => setFreeze((state.cur.frz || {}).r || 0, 1), frz0: () => setFreeze(0, 0),
  fillDown: () => fillDown(), clear: () => { commitEdit(false); clearRange(); },
  sortAsc: () => sortZone(false), sortDesc: () => sortZone(true), filter: () => openFilter()
};
$('#fmtPop').addEventListener('click', e => { const b = e.target.closest('[data-m]'); if (!b) return; closeMini(); SMENU[b.dataset.m](); });

/* police et taille des cellules */
$('#sFont').innerHTML = '<option value="">Police</option>' + GRID_FONTS.map(f => `<option value="${f}" style="font-family:${FONT_STACK[f]}">${f}</option>`).join('');
$('#sSize').innerHTML = '<option value="">Taille</option>' + GRID_SIZES.map(z => `<option value="${z}">${z}</option>`).join('');
$('#sNum').addEventListener('change', e => { setFmt('n', e.target.value); gw.focus({ preventScroll: true }); });
$('#sFont').addEventListener('change', e => { setFmt('ff', e.target.value === 'Instrument Sans' ? '' : e.target.value); gw.focus({ preventScroll: true }); });
$('#sSize').addEventListener('change', e => { setFmt('fs', e.target.value === '10' ? '' : +e.target.value); gw.focus({ preventScroll: true }); });

/* --- graphiques */
const CHART_TYPES = [['col', 'Colonnes'], ['bar', 'Barres'], ['line', 'Courbes'], ['stack', 'Colonnes empilées'], ['pie', 'Secteurs']];
const chartLayer = $('#chartLayer'), chartTip = $('#chartTip');
const CH = { sel: null, drag: null };
const fmtAxis = new Intl.NumberFormat('fr-FR', { notation: 'compact', maximumFractionDigits: 1 });
const fmtVal = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 2 });
function niceScale(lo, hi, n = 5) {
  if (lo === hi) { hi = lo + 1; lo = Math.min(0, lo); }
  const span = hi - lo, raw = span / n, mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map(m => m * mag).find(s => span / s <= n) || 10 * mag;
  const a = Math.floor(lo / step) * step, b = Math.ceil(hi / step) * step;
  const ticks = []; for (let v = a; v <= b + step / 2; v += step) ticks.push(+v.toPrecision(12));
  return { lo: a, hi: b, ticks };
}
// Lecture des données d'un graphique : 1re ligne = noms des séries, 1re colonne = étiquettes, si elles contiennent du texte.
function chartData(ch, vals, cellsOf) {
  const M = [], T = [];
  for (let r = ch.r1; r <= ch.r2; r++) {
    const row = [], trow = [];
    for (let c = ch.c1; c <= ch.c2; c++) {
      const v = (vals[r] || [])[c], x = cellsOf(r, c), f = (x && x.f) || {};
      row.push(typeof v === 'number' && f.n !== 'date' ? v : typeof v === 'number' ? { d: v } : v);
      trow.push(v == null || v === '' ? '' : typeof v === 'number' ? (f.n === 'date' ? fmtDate(v) : numFmt(f.n, f.d).format(v)) : typeof v === 'object' ? '' : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v));
    }
    M.push(row); T.push(trow);
  }
  let A = M, AT = T;
  if (ch.byRow) { A = M[0].map((_, j) => M.map(r => r[j])); AT = T[0].map((_, j) => T.map(r => r[j])); }
  const isNum = v => typeof v === 'number';
  const hasHdr = A.length > 1 && A[0].some(v => typeof v === 'string' && v !== '') && A.slice(1).some(r => r.some(isNum));
  const body = hasHdr ? A.slice(1) : A, bodyT = hasHdr ? AT.slice(1) : AT;
  const hasLbl = A[0].length > 1 && body.length > 0 && body.every(r => !isNum(r[0]));
  const sCols = []; for (let j = hasLbl ? 1 : 0; j < A[0].length; j++) sCols.push(j);
  const cats = body.map((r, i) => hasLbl ? bodyT[i][0] : String(i + 1));
  const series = sCols.map((j, k) => ({ name: hasHdr ? (AT[0][j] || `Série ${k + 1}`) : `Série ${k + 1}`, vals: body.map(r => isNum(r[j]) ? r[j] : null) })).filter(s => s.vals.some(v => v != null));
  return { cats, series, hasHdr, hasLbl };
}
const esc2 = s => esc(String(s));
function chartSVG(ch, D, W, H) {
  const ser = D.series.slice(0, 8), cats = D.cats;
  if (!ser.length || !cats.length) return `<div class="ch-empty">Aucun nombre dans ${esc(rngRef(ch))}</div>`;
  if (ch.type === 'pie') {
    const s = ser[0], items = cats.map((c, i) => ({ c, v: Math.max(0, s.vals[i] || 0), i })).filter(o => o.v > 0).slice(0, 8);
    const tot = items.reduce((a, o) => a + o.v, 0) || 1;
    const R = Math.max(20, Math.min(W, H) / 2 - 6), r0 = R * 0.56, cx = W / 2, cy = H / 2;
    let a0 = -Math.PI / 2, out = '';
    items.forEach((o, k) => {
      const a1 = a0 + o.v / tot * Math.PI * 2, large = a1 - a0 > Math.PI ? 1 : 0;
      const p = (a, rr) => `${(cx + rr * Math.cos(a)).toFixed(1)},${(cy + rr * Math.sin(a)).toFixed(1)}`;
      const d = items.length === 1 ? `M${cx - R},${cy}a${R},${R} 0 1,0 ${2 * R},0a${R},${R} 0 1,0 ${-2 * R},0M${cx - r0},${cy}a${r0},${r0} 0 1,1 ${2 * r0},0a${r0},${r0} 0 1,1 ${-2 * r0},0Z`
        : `M${p(a0, R)}A${R},${R} 0 ${large} 1 ${p(a1, R)}L${p(a1, r0)}A${r0},${r0} 0 ${large} 0 ${p(a0, r0)}Z`;
      out += `<path d="${d}" fill="var(--s${k + 1})" stroke="var(--ch-bg)" stroke-width="2" fill-rule="evenodd" data-tip="${esc2(o.c)} : ${esc2(fmtVal.format(o.v))} (${Math.round(o.v / tot * 100)} %)"/>`;
      a0 = a1;
    });
    out += `<text x="${cx}" y="${cy - 2}" text-anchor="middle" class="ch-tot">${esc(fmtAxis.format(tot))}</text><text x="${cx}" y="${cy + 14}" text-anchor="middle" class="ch-lbl">Total</text>`;
    return `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="${esc2(ch.title || 'Graphique')}">${out}</svg>`;
  }
  const horiz = ch.type === 'bar', stack = ch.type === 'stack';
  let lo = 0, hi = 0;
  if (stack) cats.forEach((_, i) => { let p = 0, n = 0; ser.forEach(s => { const v = s.vals[i] || 0; if (v >= 0) p += v; else n += v; }); hi = Math.max(hi, p); lo = Math.min(lo, n); });
  else ser.forEach(s => s.vals.forEach(v => { if (v != null) { hi = Math.max(hi, v); lo = Math.min(lo, v); } }));
  if (ch.type === 'line' && lo > 0) { const all = ser.flatMap(s => s.vals.filter(v => v != null)); const mn = Math.min(...all); if (mn > (hi - mn) * 2) lo = mn; }
  const sc = niceScale(lo, hi);
  const labW = horiz ? Math.min(120, Math.max(...cats.map(c => String(c).length)) * 6.5 + 10) : Math.max(...sc.ticks.map(t => fmtAxis.format(t).length)) * 7 + 10;
  const padL = labW, padR = 10, padT = 8, padB = horiz ? 22 : 26;
  const pw = Math.max(20, W - padL - padR), ph = Math.max(20, H - padT - padB);
  const val2 = v => horiz ? padL + (v - sc.lo) / (sc.hi - sc.lo) * pw : padT + ph - (v - sc.lo) / (sc.hi - sc.lo) * ph;
  let g = '';
  sc.ticks.forEach(t => {
    const p = val2(t).toFixed(1);
    g += horiz ? `<line x1="${p}" y1="${padT}" x2="${p}" y2="${padT + ph}" class="ch-grid"/><text x="${p}" y="${H - 6}" text-anchor="middle" class="ch-ax">${esc(fmtAxis.format(t))}</text>`
      : `<line x1="${padL}" y1="${p}" x2="${padL + pw}" y2="${p}" class="ch-grid"/><text x="${padL - 6}" y="${+p + 4}" text-anchor="end" class="ch-ax">${esc(fmtAxis.format(t))}</text>`;
  });
  const z = val2(Math.max(sc.lo, Math.min(0, sc.hi))).toFixed(1);
  g += horiz ? `<line x1="${z}" y1="${padT}" x2="${z}" y2="${padT + ph}" class="ch-base"/>` : `<line x1="${padL}" y1="${z}" x2="${padL + pw}" y2="${z}" class="ch-base"/>`;
  const n = cats.length, band = (horiz ? ph : pw) / n;
  const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horiz ? ph / 16 : pw / 64)))));
  cats.forEach((c, i) => {
    if (i % every) return;
    const m = (horiz ? padT : padL) + band * (i + .5);
    const t = String(c).length > 14 ? String(c).slice(0, 13) + '…' : String(c);
    g += horiz ? `<text x="${padL - 6}" y="${(m + 4).toFixed(1)}" text-anchor="end" class="ch-ax">${esc(t)}</text>` : `<text x="${m.toFixed(1)}" y="${H - 8}" text-anchor="middle" class="ch-ax">${esc(t)}</text>`;
  });
  const rbar = (x, y, w, h, up, k, tip) => {
    // extrémité arrondie (4 px) du côté de la valeur, base droite contre l'axe
    const rr = Math.min(4, Math.abs(w) / 2, Math.abs(h) / 2);
    let d;
    if (!horiz) d = up ? `M${x},${y + h}V${y + rr}Q${x},${y} ${x + rr},${y}H${x + w - rr}Q${x + w},${y} ${x + w},${y + rr}V${y + h}Z` : `M${x},${y}V${y + h - rr}Q${x},${y + h} ${x + rr},${y + h}H${x + w - rr}Q${x + w},${y + h} ${x + w},${y + h - rr}V${y}Z`;
    else d = up ? `M${x},${y}H${x + w - rr}Q${x + w},${y} ${x + w},${y + rr}V${y + h - rr}Q${x + w},${y + h} ${x + w - rr},${y + h}H${x}Z` : `M${x + w},${y}H${x + rr}Q${x},${y} ${x},${y + rr}V${y + h - rr}Q${x},${y + h} ${x + rr},${y + h}H${x + w}Z`;
    return `<path d="${d}" fill="var(--s${k + 1})" data-tip="${esc2(tip)}"/>`;
  };
  let m = '';
  if (ch.type === 'line') {
    ser.forEach((s, k) => {
      let d = '', pen = false;
      s.vals.forEach((v, i) => { if (v == null) { pen = false; return; } const x = (padL + band * (i + .5)).toFixed(1), y = val2(v).toFixed(1); d += (pen ? 'L' : 'M') + x + ',' + y; pen = true; });
      m += `<path d="${d}" fill="none" stroke="var(--s${k + 1})" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
      s.vals.forEach((v, i) => {
        if (v == null) return;
        const x = (padL + band * (i + .5)).toFixed(1), y = val2(v).toFixed(1);
        if (n <= 24) m += `<circle cx="${x}" cy="${y}" r="4" fill="var(--s${k + 1})" stroke="var(--ch-bg)" stroke-width="2"/>`;
        m += `<circle cx="${x}" cy="${y}" r="11" fill="transparent" data-tip="${esc2(s.name)} · ${esc2(cats[i])} : ${esc2(fmtVal.format(v))}"/>`;
      });
    });
  } else if (stack) {
    const bw = Math.max(2, band * .62);
    cats.forEach((c, i) => {
      let pos = 0, neg = 0;
      ser.forEach((s, k) => {
        const v = s.vals[i]; if (!v) return;
        const a = v >= 0 ? pos : neg, b = a + v; if (v >= 0) pos = b; else neg = b;
        const top = !ser.slice(k + 1).some(q => q.vals[i] && Math.sign(q.vals[i]) === Math.sign(v));
        const p1 = val2(a), p2 = val2(b);
        const x = padL + band * i + (band - bw) / 2, y = Math.min(p1, p2), h = Math.max(0, Math.abs(p2 - p1) - (top ? 0 : 2));
        const tip = `${s.name} · ${c} : ${fmtVal.format(v)}`;
        m += top ? rbar(x, v >= 0 ? y : y + 0, bw, h, v >= 0, k, tip) : `<rect x="${x.toFixed(1)}" y="${(v >= 0 ? y + 2 : y).toFixed(1)}" width="${bw.toFixed(1)}" height="${h.toFixed(1)}" fill="var(--s${k + 1})" data-tip="${esc2(tip)}"/>`;
      });
    });
  } else {
    const k = ser.length, inner = band * .74, gap = k > 1 ? 2 : 0, bw = Math.max(1.5, (inner - gap * (k - 1)) / k);
    cats.forEach((c, i) => ser.forEach((s, j) => {
      const v = s.vals[i]; if (v == null) return;
      const start = band * i + (band - inner) / 2 + j * (bw + gap);
      const p0 = val2(Math.max(sc.lo, Math.min(0, sc.hi))), p1 = val2(v), tip = `${k > 1 ? s.name + ' · ' : ''}${c} : ${fmtVal.format(v)}`;
      if (horiz) { const x = Math.min(p0, p1), w = Math.abs(p1 - p0); m += rbar(x, padT + start, w, bw, v >= 0, j, tip); }
      else { const y = Math.min(p0, p1), h = Math.abs(p1 - p0); m += rbar(padL + start, y, bw, h, v >= 0, j, tip); }
    }));
  }
  return `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="${esc2(ch.title || 'Graphique')}">${g}${m}</svg>`;
}
function chartHTML(ch, vals, cellsOf, forPrint) {
  const D = chartData(ch, vals, cellsOf);
  const legendItems = ch.type === 'pie' ? D.cats.slice(0, 8).map((c, i) => [c, i]).filter(([, i]) => (D.series[0] && D.series[0].vals[i]) > 0) : D.series.length > 1 ? D.series.slice(0, 8).map((s, i) => [s.name, i]) : [];
  const head = (ch.title ? `<div class="ch-title">${esc(ch.title)}</div>` : '') + (legendItems.length ? `<div class="ch-leg">${legendItems.map(([t, i]) => `<span><i style="background:var(--s${i + 1})"></i>${esc(String(t))}</span>`).join('')}</div>` : '');
  const hh = (ch.title ? 24 : 0) + (legendItems.length ? 20 * Math.ceil(legendItems.length / Math.max(1, Math.floor((ch.w - 24) / 120))) + 4 : 0);
  const svg = chartSVG(ch, D, Math.max(60, ch.w - 24), Math.max(60, ch.h - 24 - hh));
  return `<div class="ch-in">${head}${svg}</div>`;
}
function chartPos(ch) {
  const r = Math.min(ch.at.r, state.cur.rows - 1), c = Math.min(ch.at.c, state.cur.cols - 1);
  const td = tdOf(r, c), t = $('#grid');
  return { x: t.offsetLeft + td.offsetLeft + (ch.dx || 0), y: t.offsetTop + td.offsetTop + (ch.dy || 0) };
}
function renderCharts() {
  if (!chartLayer) return;
  const cur = state.cur;
  if (!cur || state.view !== 'sheet' || !(cur.charts || []).length) { chartLayer.innerHTML = ''; return; }
  const keep = new Set();
  cur.charts.forEach(ch => {
    keep.add(ch.id);
    let el = chartLayer.querySelector(`[data-id="${ch.id}"]`);
    if (!el) { el = document.createElement('div'); el.className = 'chart'; el.dataset.id = ch.id; chartLayer.appendChild(el); }
    const p = chartPos(ch);
    el.style.left = p.x + 'px'; el.style.top = p.y + 'px'; el.style.width = ch.w + 'px'; el.style.height = ch.h + 'px';
    el.classList.toggle('sel', CH.sel === ch.id);
    const html = chartHTML(ch, SH.vals, cellOf) + `<span class="ch-rz" aria-hidden="true"></span>` + (CH.sel === ch.id ? `<div class="ch-bar"><button data-ca="edit">${ic('pen')}Modifier</button><button data-ca="del" class="danger">${ic('trash')}Supprimer</button></div>` : '');
    if (el._h !== html) { el._h = html; el.innerHTML = html; }
  });
  [...chartLayer.children].forEach(el => { if (!keep.has(el.dataset.id)) el.remove(); });
}
function selectChart(id) { CH.sel = id; renderCharts(); }
function findChart(id) { return (state.cur.charts || []).find(c => c.id === id); }
function anchorAt(x, y) {
  const t = $('#grid'), cur = state.cur;
  let c = 0; for (let j = 0; j < cur.cols; j++) { const td = SH.cell[0][j]; if (t.offsetLeft + td.offsetLeft <= x) c = j; else break; }
  let r = 0; for (let i = 0; i < cur.rows; i++) { const tr = SH.cell[i][0].parentElement; if (tr.hidden) continue; if (t.offsetTop + tr.offsetTop <= y) r = i; else break; }
  const td = SH.cell[r][c];
  return { at: { r, c }, dx: Math.round(x - t.offsetLeft - td.offsetLeft), dy: Math.round(y - t.offsetTop - td.offsetTop) };
}
chartLayer.addEventListener('pointerdown', e => {
  const el = e.target.closest('.chart'); if (!el) return;
  if (e.target.closest('.ch-bar')) return;
  e.preventDefault(); e.stopPropagation();
  commitEdit(false);
  const ch = findChart(el.dataset.id); if (!ch) return;
  if (CH.sel !== ch.id) selectChart(ch.id);
  const p = chartPos(ch);
  CH.drag = { id: ch.id, mode: e.target.classList.contains('ch-rz') ? 'size' : 'move', sx: e.clientX, sy: e.clientY, x: p.x, y: p.y, w: ch.w, h: ch.h, moved: false, before: snap() };
  try { el.setPointerCapture(e.pointerId); } catch {}
});
chartLayer.addEventListener('pointermove', e => {
  const d = CH.drag; if (!d) { const t = e.target.closest('[data-tip]'); if (t) { chartTip.textContent = t.dataset.tip; chartTip.hidden = false; const host = $('#sheet').getBoundingClientRect(); chartTip.style.left = Math.min(e.clientX - host.left + 12, host.width - chartTip.offsetWidth - 8) + 'px'; chartTip.style.top = (e.clientY - host.top - chartTip.offsetHeight - 10) + 'px'; } else chartTip.hidden = true; return; }
  const dx = e.clientX - d.sx, dy = e.clientY - d.sy;
  if (Math.abs(dx) + Math.abs(dy) > 3) d.moved = true;
  const el = chartLayer.querySelector(`[data-id="${d.id}"]`), ch = findChart(d.id); if (!el || !ch) return;
  if (d.mode === 'move') { el.style.left = Math.max(46, d.x + dx) + 'px'; el.style.top = Math.max(26, d.y + dy) + 'px'; }
  else { ch.w = Math.max(220, Math.round(d.w + dx)); ch.h = Math.max(160, Math.round(d.h + dy)); el._h = null; renderCharts(); }
});
chartLayer.addEventListener('pointerleave', () => { if (!CH.drag) chartTip.hidden = true; });
function endChartDrag(e) {
  const d = CH.drag; CH.drag = null; if (!d) return;
  const ch = findChart(d.id); if (!ch) return;
  if (d.moved) {
    if (d.mode === 'move') { const el = chartLayer.querySelector(`[data-id="${d.id}"]`); Object.assign(ch, anchorAt(parseFloat(el.style.left), parseFloat(el.style.top))); }
    SH.hist.push(d.before); SH.fut = [];
    scheduleSave(); renderCharts();
  }
}
chartLayer.addEventListener('pointerup', endChartDrag);
chartLayer.addEventListener('pointercancel', endChartDrag);
chartLayer.addEventListener('dblclick', e => { const el = e.target.closest('.chart'); if (el) openChartEditor(el.dataset.id); });
chartLayer.addEventListener('click', e => {
  const b = e.target.closest('[data-ca]'); if (!b) return;
  const id = CH.sel;
  if (b.dataset.ca === 'edit') openChartEditor(id);
  else deleteChart(id);
});
function deleteChart(id) { pushHist(); state.cur.charts = state.cur.charts.filter(c => c.id !== id); CH.sel = null; renderCharts(); scheduleSave(); toast('Graphique supprimé'); }
gw.addEventListener('mousedown', e => { if (CH.sel && !e.target.closest('.chart')) { CH.sel = null; renderCharts(); } }, true);
document.addEventListener('keydown', e => {
  if (!CH.sel || state.view !== 'sheet' || SH.editing || e.target.closest('input,textarea,select')) return;
  if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); e.stopPropagation(); deleteChart(CH.sel); }
  else if (e.key === 'Escape') { CH.sel = null; renderCharts(); }
}, true);
function openChartEditor(id) {
  commitEdit(false);
  const cur = state.cur; cur.charts ||= [];
  const ex = id ? findChart(id) : null;
  let R = ex ? { r1: ex.r1, c1: ex.c1, r2: ex.r2, c2: ex.c2 } : rng();
  if (!ex && R.r1 === R.r2 && R.c1 === R.c2) R = currentRegion(R.r1, R.c1);
  const draft = ex ? { ...ex } : { type: 'col', title: '', byRow: false };
  SH.dlg = 'chart';
  openDlg(dlgHead(ex ? 'Modifier le graphique' : 'Insérer un graphique') + `
    <span class="dlg-lbl">Type</span><div class="ch-types">${CHART_TYPES.map(([k, l]) => `<button type="button" data-ct="${k}" aria-pressed="${draft.type === k}"><svg viewBox="0 0 24 24" aria-hidden="true">${CT_ICON[k]}</svg>${l}</button>`).join('')}</div>
    <label class="dlg-f"><span>Données</span><input id="chRange" value="${esc(rngRef(R))}" autocomplete="off" autocapitalize="off" spellcheck="false"></label>
    <label class="dlg-f"><span>Titre</span><input id="chTitle" value="${esc(draft.title || '')}" placeholder="Dépenses par mois" autocomplete="off"></label>
    <label class="fp-it"><input type="checkbox" id="chByRow"${draft.byRow ? ' checked' : ''}><span>Séries en lignes (au lieu des colonnes)</span></label>
    <div class="ch-prev" id="chPrev"></div>
    <div class="fp-acts"><button class="pbtn" id="chOk">${ex ? 'Enregistrer' : 'Insérer le graphique'}</button>${ex ? '<button class="pbtn ghost" id="chDel">Supprimer</button>' : ''}</div>`, 'wide');
  const prev = () => {
    const o = rangeObj($('#chRange').value.replace(/^=/, ''));
    const box = $('#chPrev');
    if (!o) { box.innerHTML = '<div class="ch-empty">Plage incorrecte</div>'; return; }
    const w = Math.min(360, sDlg.clientWidth - 30), hh = isPhone() ? 170 : 220;
    box.innerHTML = `<div class="chart static" style="width:${w}px;height:${hh}px">${chartHTML({ ...draft, ...o, title: $('#chTitle').value.trim(), byRow: $('#chByRow').checked, w, h: hh }, SH.vals, cellOf)}</div>`;
  };
  sDlg.querySelector('.ch-types').addEventListener('click', e => { const b = e.target.closest('[data-ct]'); if (!b) return; draft.type = b.dataset.ct; sDlg.querySelectorAll('[data-ct]').forEach(x => x.setAttribute('aria-pressed', String(x === b))); prev(); });
  ['chRange', 'chTitle'].forEach(i => $('#' + i).addEventListener('input', prev));
  $('#chByRow').addEventListener('change', prev);
  prev();
  $('#chOk').addEventListener('click', () => {
    const o = rangeObj($('#chRange').value.replace(/^=/, ''));
    if (!o) return toast('Plage incorrecte (exemple : A1:C12).', true);
    const D = chartData({ ...o, byRow: $('#chByRow').checked }, SH.vals, cellOf);
    if (!D.series.length) return toast('Cette plage ne contient pas de nombres à représenter.', true);
    pushHist();
    const data = { ...o, type: draft.type, title: $('#chTitle').value.trim(), byRow: $('#chByRow').checked };
    if (ex) Object.assign(ex, data);
    else {
      const w = isPhone() ? Math.min(360, gw.clientWidth - 70) : 480, h = isPhone() ? 260 : 300;
      const at = { r: o.r1, c: Math.min(cur.cols - 1, o.c2 + 1) };
      if (isPhone()) { at.r = Math.min(cur.rows - 1, o.r2 + 2); at.c = 0; }
      const ch = { id: uid(), ...data, at, dx: 12, dy: 0, w, h };
      cur.charts.push(ch); CH.sel = ch.id;
    }
    closeDlg(); renderCharts(); scheduleSave();
    const el = chartLayer.querySelector(`[data-id="${CH.sel || (ex && ex.id)}"]`); if (el) el.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'smooth' });
    toast(ex ? 'Graphique modifié' : 'Graphique inséré : faites-le glisser pour le déplacer.');
  });
  const del = $('#chDel'); if (del) del.addEventListener('click', () => { closeDlg(); deleteChart(ex.id); });
}
const CT_ICON = {
  col: '<path d="M5 20V11M10 20V5M15 20v-6M20 20V8" stroke-width="3"/>',
  bar: '<path d="M4 5h9M4 10h15M4 15h6M4 20h11" stroke-width="3"/>',
  line: '<path d="M3 17l5-6 4 4 8-9"/>',
  stack: '<path d="M6 20v-5M6 13V9M12 20v-8M12 10V4M18 20v-4M18 14v-3" stroke-width="3"/>',
  pie: '<circle cx="12" cy="12" r="8"/><path d="M12 4v8l6 5"/>'
};

/* --- tableau croisé dynamique (formules SOMME.SI.ENS sur une nouvelle feuille) */
const AGG = { sum: ['Somme', 'SOMME.SI.ENS'], count: ['Nombre', 'NB.SI.ENS'], avg: ['Moyenne', 'MOYENNE.SI.ENS'], max: ['Maximum', 'MAX.SI.ENS'], min: ['Minimum', 'MIN.SI.ENS'] };
function openPivot() {
  commitEdit(false);
  const cur = state.cur, S0 = rng();
  let R = S0.r2 > S0.r1 ? S0 : currentRegion(S0.r1, S0.c1);
  if (R.r2 - R.r1 < 1) return toast('Placez-vous dans un tableau avec une ligne de titres et des données.', true);
  // ligne de titres : première ligne entièrement remplie de textes
  for (let r = R.r1; r < R.r2; r++) {
    let ok = true; for (let c = R.c1; c <= R.c2; c++) { const v = (SH.vals[r] || [])[c]; if (typeof v !== 'string' || v === '') { ok = false; break; } }
    if (ok) { R = { ...R, r1: r }; break; }
  }
  const hdr = []; for (let c = R.c1; c <= R.c2; c++) hdr.push(display(R.r1, c).txt || COLN(c));
  const numCol = hdr.findIndex((_, j) => { for (let r = R.r1 + 1; r <= R.r2; r++) if (typeof (SH.vals[r] || [])[R.c1 + j] === 'number' && ((cellOf(r, R.c1 + j) || {}).f || {}).n !== 'date') return true; return false; });
  const opt = (sel, none) => (none ? `<option value="-1">(aucune)</option>` : '') + hdr.map((h, j) => `<option value="${j}"${j === sel ? ' selected' : ''}>${esc(h)}</option>`).join('');
  SH.dlg = 'pivot';
  openDlg(dlgHead('Tableau croisé dynamique') + `<p class="dlg-note">Source : ${esc(cur.sheets[cur.active].name)}!${rngRef(R)}, titres en ligne ${R.r1 + 1}. Le tableau est créé sur une nouvelle feuille, avec des formules qui se mettent à jour.</p>
    <label class="dlg-f"><span>Données</span><input id="pvRange" value="${esc(rngRef(R))}" autocomplete="off" autocapitalize="off" spellcheck="false"></label>
    <label class="dlg-f"><span>Lignes</span><select id="pvRows">${opt(0)}</select></label>
    <label class="dlg-f"><span>Colonnes</span><select id="pvCols">${opt(-1, true)}</select></label>
    <div class="dlg-row"><label class="dlg-f"><span>Valeurs</span><select id="pvVal">${opt(numCol < 0 ? hdr.length - 1 : numCol)}</select></label>
    <label class="dlg-f"><span>Calcul</span><select id="pvAgg">${Object.entries(AGG).map(([k, a]) => `<option value="${k}">${a[0]}</option>`).join('')}</select></label></div>
    <div class="fp-acts"><button class="pbtn" id="pvOk">Créer le tableau</button></div>`);
  $('#pvRange').addEventListener('change', () => { const o = rangeObj($('#pvRange').value); if (o) { closeDlg(); SH.anc = { r: o.r2, c: o.c2 }; select(o.r1, o.c1, true); openPivot(); } });
  $('#pvOk').addEventListener('click', () => {
    const def = { src: cur.sheets[cur.active].id, ...R, rows: +$('#pvRows').value, cols: +$('#pvCols').value, val: +$('#pvVal').value, agg: $('#pvAgg').value };
    if (def.cols === def.rows) def.cols = -1;
    closeDlg(); createPivot(def);
  });
}
function pivotCells(def) {
  const cur = state.cur, si = cur.sheets.findIndex(s => s.id === def.src);
  if (si < 0) return null;
  const src = cur.sheets[si], vals = SH.hf && SH.ids ? SH.hf.getSheetValues(SH.ids[si]) : [];
  const keyOf = (r, c) => { const v = (vals[r] || [])[c]; return v == null || v === '' || typeof v === 'object' ? null : v; };
  const uniq = c => { const m = new Map(); for (let r = def.r1 + 1; r <= def.r2; r++) { const v = keyOf(r, c); if (v == null) continue; const k = typeof v === 'string' ? v.toLowerCase() : v; if (!m.has(k)) m.set(k, v); } return [...m.values()].sort((a, b) => typeof a === 'number' && typeof b === 'number' ? a - b : String(a).localeCompare(String(b), 'fr', { numeric: true })); };
  const rowsK = uniq(def.c1 + def.rows), colsK = def.cols >= 0 ? uniq(def.c1 + def.cols) : [];
  const S = refName(src.name) + '!';
  const col = j => `${S}$${COLN(def.c1 + j)}$${def.r1 + 2}:$${COLN(def.c1 + j)}$${def.r2 + 1}`;
  const V = col(def.val), RR = col(def.rows), CC = def.cols >= 0 ? col(def.cols) : null;
  const [aggName, fn] = AGG[def.agg];
  const hdr = c => { const v = keyOf(def.r1, def.c1 + c); return v == null ? COLN(def.c1 + c) : String(v); };
  const crit = (rowRef, colRef) => (rowRef ? `;${RR};${rowRef}` : '') + (colRef ? `;${CC};${colRef}` : '');
  const f = (rowRef, colRef) => {
    if (!rowRef && !colRef) return def.agg === 'sum' ? `=SOMME(${V})` : def.agg === 'count' ? `=NB.SI.ENS(${RR};"<>")` : def.agg === 'avg' ? `=SIERREUR(MOYENNE(${V});"")` : def.agg === 'max' ? `=MAX(${V})` : `=MIN(${V})`;
    if (def.agg === 'count') return `=NB.SI.ENS(${(crit(rowRef, colRef)).slice(1)})`;
    const core = `${fn}(${V}${crit(rowRef, colRef)})`;
    return def.agg === 'avg' ? `=SIERREUR(${core};"")` : `=${core}`;
  };
  const srcFmt = ((src.cells[addr(def.r1 + 1, def.c1 + def.val)] || {}).f || {}).n;
  const nf = def.agg === 'count' ? '' : srcFmt === 'date' ? '' : srcFmt || '';
  const cells = {};
  const put = (r, c, v, fm) => { const t = typeof v === 'number' ? numToRaw(v) : String(v); cells[addr(r, c)] = { v: typeof v === 'string' && v[0] !== '=' && (parseNum(v) != null || parseDate(v) != null || /^(VRAI|FAUX)$/i.test(v)) ? "'" + v : t, f: fm || {} }; };
  const H = { b: 1, bg: 'cuivre', bb: 1 }, T = { b: 1, bt: 1 };
  put(0, 0, `${aggName} de ${hdr(def.val)}`, { ...H });
  if (def.cols >= 0) { colsK.forEach((k, j) => put(0, j + 1, k, { ...H, a: 'ar' })); put(0, colsK.length + 1, 'Total', { ...H, a: 'ar' }); }
  else put(0, 1, aggName, { ...H, a: 'ar' });
  rowsK.forEach((k, i) => {
    const r = i + 1;
    put(r, 0, k, typeof k === 'number' ? { a: 'al' } : {});
    if (def.cols >= 0) {
      colsK.forEach((_, j) => put(r, j + 1, f(`$A${r + 1}`, `${COLN(j + 1)}$1`), nf ? { n: nf } : {}));
      put(r, colsK.length + 1, f(`$A${r + 1}`, null), { b: 1, ...(nf ? { n: nf } : {}) });
    } else put(r, 1, f(`$A${r + 1}`, null), nf ? { n: nf } : {});
  });
  const tr = rowsK.length + 1;
  put(tr, 0, 'Total', { ...T });
  if (def.cols >= 0) { colsK.forEach((_, j) => put(tr, j + 1, f(null, `${COLN(j + 1)}$1`), { ...T, ...(nf ? { n: nf } : {}) })); put(tr, colsK.length + 1, f(null, null), { ...T, ...(nf ? { n: nf } : {}) }); }
  else put(tr, 1, f(null, null), { ...T, ...(nf ? { n: nf } : {}) });
  return { cells, rows: Math.max(60, tr + 20), cols: Math.max(12, (def.cols >= 0 ? colsK.length + 2 : 2) + 2), name: `TCD ${hdr(def.val)}`, colW: { 0: 180 } };
}
function createPivot(def) {
  const P = pivotCells(def); if (!P) return;
  const cur = state.cur;
  const sh = { id: uid(), name: uniqueName(P.name.replace(/[\[\]:*?\/\\']/g, '').slice(0, 31)), cells: P.cells, rows: P.rows, cols: P.cols, colW: P.colW, pivot: def, hid: { r: [], c: [] }, cf: [], charts: [] };
  cur.sheets.splice(cur.active + 1, 0, sh); cur.active += 1;
  cur.sheets.forEach(markDates);
  loadHF(); refreshSheetView(); scheduleSave();
  toast('Tableau croisé créé. Pour l\'actualiser : menu de l\'onglet → Actualiser.');
}
function refreshPivot(i) {
  const cur = state.cur, sh = cur.sheets[i]; if (!sh || !sh.pivot) return;
  const P = pivotCells(sh.pivot); if (!P) return toast('La feuille source a été supprimée.', true);
  sh.cells = P.cells; sh.rows = Math.max(sh.rows, P.rows); sh.cols = Math.max(sh.cols, P.cols);
  loadHF(); if (i === cur.active) refreshSheetView(); scheduleSave(); toast('Tableau croisé actualisé');
}

/* --- somme auto et structure */
function autoSum() {
  commitEdit(false);
  const R = rng();
  if (R.r2 > R.r1 && R.c1 === R.c2) {
    const target = R.r2 + 1; ensureSize(target, R.c1);
    setRaw(target, R.c1, `=SOMME(${addr(R.r1, R.c1)}:${addr(R.r2, R.c1)})`); select(target, R.c1); return;
  }
  const { r, c } = SH.sel; let top = r - 1;
  while (top >= 0 && typeof (SH.vals[top] || [])[c] === 'number') top--;
  const ref = top + 1 <= r - 1 ? `${addr(top + 1, c)}:${addr(r - 1, c)}` : '';
  startEdit(`=SOMME(${ref})`);
  if (!ref) cellEd.setSelectionRange(7, 7);
}
// Insertion ou suppression de ligne / colonne : plages des règles, graphiques, lignes masquées et plages nommées suivent.
function shiftSheetObjects(op, at) {
  const cur = state.cur, row = op[0] === 'r', ins = /Ins$/.test(op);
  const mv = i => ins ? (i >= at ? i + 1 : i) : (i > at ? i - 1 : i === at ? null : i);
  const span = (a, b) => {
    if (ins) return [a >= at ? a + 1 : a, b >= at ? b + 1 : b];
    if (a === at && b === at) return null;
    return [a > at ? a - 1 : a, b >= at ? b - 1 : b];
  };
  const fixR = o => { const k1 = row ? 'r1' : 'c1', k2 = row ? 'r2' : 'c2'; const sp = span(o[k1], o[k2]); if (!sp) return false; o[k1] = sp[0]; o[k2] = sp[1]; return true; };
  if (cur.hid) { const k = row ? 'r' : 'c'; cur.hid[k] = cur.hid[k].map(mv).filter(x => x != null); }
  cur.cf = (cur.cf || []).filter(fixR);
  (cur.charts || []).forEach(ch => { fixR(ch); const k = row ? 'r' : 'c'; const n = mv(ch.at[k]); ch.at[k] = n == null ? at : n; });
  cur.charts = (cur.charts || []).filter(ch => ch.r2 >= ch.r1 && ch.c2 >= ch.c1);
  const sname = cur.sheets[cur.active].name;
  Object.values(state.cur.names || {}).forEach(nm => { if (nm.s !== sname) return; const o = rangeObj(nm.ref); if (!o || !fixR(o)) return; nm.ref = rngRef(o); });
}
function structural(op) {
  commitEdit(false);
  if (!SH.hf) return toast('Le moteur de calcul est nécessaire pour insérer ou supprimer des lignes.', true);
  const cur = state.cur, { r, c } = SH.sel;
  if (cur.filter) { cur.filter = undefined; SH.hiddenRows = 0; toast('Filtre retiré pour modifier la structure.'); }
  if (op === 'rowDel' && cur.rows <= 1 || op === 'colDel' && cur.cols <= 1) return;
  pushHist();
  const fm = {}; Object.entries(cur.cells).forEach(([k, x]) => { if (x.f && Object.keys(x.f).length) fm[k] = x.f; });
  const shift = (k, dr, dc) => { const p = parseAddr(k); return addr(p.r + dr, p.c + dc); };
  const nf = {};
  Object.entries(fm).forEach(([k, f]) => {
    const p = parseAddr(k);
    if (op === 'rowIns') nf[p.r >= r ? shift(k, 1, 0) : k] = f;
    else if (op === 'rowDel') { if (p.r !== r) nf[p.r > r ? shift(k, -1, 0) : k] = f; }
    else if (op === 'colIns') nf[p.c >= c ? shift(k, 0, 1) : k] = f;
    else if (op === 'colDel') { if (p.c !== c) nf[p.c > c ? shift(k, 0, -1) : k] = f; }
  });
  if (op === 'rowIns') { SH.hf.addRows(SH.sid, [r, 1]); cur.rows++; }
  if (op === 'rowDel') { SH.hf.removeRows(SH.sid, [r, 1]); }
  if (op === 'colIns') {
    SH.hf.addColumns(SH.sid, [c, 1]); cur.cols++;
    const w = {}; Object.entries(cur.colW).forEach(([i, v]) => { w[+i >= c ? +i + 1 : i] = v; }); cur.colW = w;
  }
  if (op === 'colDel') {
    SH.hf.removeColumns(SH.sid, [c, 1]);
    const w = {}; Object.entries(cur.colW).forEach(([i, v]) => { if (+i !== c) w[+i > c ? +i - 1 : i] = v; }); cur.colW = w;
  }
  shiftSheetObjects(op, op[0] === 'r' ? r : c);
  const cells = cellsFromHF(SH.sid, {}, true);
  Object.entries(nf).forEach(([k, f]) => { (cells[k] ||= { v: '', f: {} }).f = f; });
  cur.cells = fixDates(cells);
  cur.sheets.forEach((sh, i) => { if (i !== cur.active) sh.cells = cellsFromHF(SH.ids[i], sh.cells); });
  buildGrid(); afterChange(); select(Math.min(r, cur.rows - 1), Math.min(c, cur.cols - 1));
}
const S = {
  undo: () => { commitEdit(false); undo(); }, redo: () => { commitEdit(false); redo(); },
  bold: () => setFmt('b'), al: () => setFmt('a', 'al'), ac: () => setFmt('a', 'ac'), ar: () => setFmt('a', 'ar'),
  sum: autoSum, clear: () => { commitEdit(false); clearRange(); },
  ital: () => setFmt('i'), under: () => setFmt('u'), wrap: () => setFmt('w'),
  fill: () => openFmt('bg'), color: () => openFmt('fc'), borders: () => openFmt('bd'), freeze: () => openFmt('frz'),
  fillDown, sortAsc: () => sortZone(false), sortDesc: () => sortZone(true), filter: () => { $('#filterPop').hidden ? openFilter() : closeFilter(); },
  rowIns: () => structural('rowIns'), rowDel: () => structural('rowDel'), colIns: () => structural('colIns'), colDel: () => structural('colDel'),
  merge: toggleMerge, decLess: () => stepDecimals(-1), decMore: () => stepDecimals(1), structMenu: () => openSheetMenu('struct'),
  insMenu: () => openSheetMenu('ins'), dataMenu: () => openSheetMenu('data')
};
$('#sheet .toolbar').addEventListener('mousedown', e => { if (e.target.closest('button')) e.preventDefault(); });
$('#sheet .toolbar').addEventListener('click', e => {
  const b = e.target.closest('[data-s]'); if (b) { S[b.dataset.s](); if (!SH.editing) gw.focus({ preventScroll: true }); return; }
  const n = e.target.closest('[data-n]'); if (n) { setFmt('n', n.dataset.n); gw.focus({ preventScroll: true }); }
});
window.addEventListener('resize', debounce(() => { if (state.view === 'sheet') { drawSel(); if (SH.editing === 'cell') placeEditor(); } }, 120));

/* =========================================================
   PDF : lire, remplir, annoter, organiser les pages, fusionner, diviser
   ========================================================= */
const PDFV = { doc: null, src: null, srcPages: [], bytes: null, thumbC: new Map(), srcText: [], pages: [], scale: 1, fit: true, mode: 'read', tool: 'select', sel: null, hist: [], fut: [], thumbs: {}, texts: null, obs: null, pageSel: new Set(), fields: [], gen: 0, color: { text: '#1E1814', hl: '#FFE066', rect: '#C0392B', ink: '#1F5FAD' }, size: 14, lw: 2 };
const PDF_COLORS = [['Noir', '#1E1814'], ['Bleu', '#1F5FAD'], ['Rouge', '#C0392B'], ['Vert', '#2F7D46'], ['Orange', '#E67E22'], ['Violet', '#7D3C98']];
const HL_COLORS2 = [['Jaune', '#FFE066'], ['Vert', '#A8E6A1'], ['Bleu', '#9FD3FF'], ['Rose', '#FFB3C7'], ['Orange', '#FFC78A']];
const PT = 96 / 72;
const b64ToBytes = b64 => { const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };
const bytesToB64 = u8 => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };
let pdfjsMod = null;
async function ensurePdfJs() {
  if (!pdfjsMod) pdfjsMod = import(new URL(LIB.pdfjs, location.href).href).catch(e => { pdfjsMod = null; throw e; });
  const L = await pdfjsMod;
  if (!L.GlobalWorkerOptions.workerSrc) L.GlobalWorkerOptions.workerSrc = new URL(LIB.pdfWorker, location.href).href;
  window.pdfjsLib = L;
  return L;
}
async function ensurePdfLib() { await loadScript(LIB.pdflib); return window.PDFLib; }
async function pdfjsOpen(bytes) {
  const L = await ensurePdfJs();
  // isImageDecoderSupported : les photos (JPEG) sont décodées par le navigateur lui-même, bien plus vite
  const task = L.getDocument({ data: bytes.slice(0), cMapUrl: 'lib/pdfjs/cmaps/', cMapPacked: true, standardFontDataUrl: 'lib/pdfjs/standard_fonts/', isEvalSupported: false, isImageDecoderSupported: true });
  task.onPassword = (cb, reason) => {
    const p = window.prompt(reason === 2 ? 'Mot de passe incorrect. Réessayez :' : 'Ce PDF est protégé par un mot de passe :');
    if (p == null) task.destroy(); else cb(p);
  };
  return task.promise;
}
// Texte de chaque page (pour la recherche et la conversion en document)
async function pageText(doc, i, withLines) {
  const tc = await (await doc.getPage(i + 1)).getTextContent();
  let s = '', lastY = null;
  tc.items.forEach(it => {
    if (!('str' in it)) return;
    const y = it.transform ? Math.round(it.transform[5]) : null;
    if (withLines && lastY != null && y != null && Math.abs(y - lastY) > 2 && !/\n$/.test(s)) s += '\n';
    s += it.str + (it.hasEOL ? '\n' : '');
    lastY = y;
  });
  return s;
}
async function pdfAllText(doc, max = 60) {
  const out = [];
  for (let i = 0; i < Math.min(doc.numPages, max); i++) { try { out.push(await pageText(doc, i)); } catch { out.push(''); } }
  return out.join('\n').replace(/[ \t]+/g, ' ').slice(0, 30000);
}

/* --- ouverture
   Le fichier d'origine (cur.pdfId, dans le magasin « blobs ») ne change pas quand on réorganise les pages :
   cur.order décrit les pages affichées ({ s: page source, r: rotation ajoutée }). Le PDF n'est réécrit qu'à l'export. */
async function openPdf() {
  show('pdfv');
  const cur = state.cur;
  cur.ann ||= []; cur.form ||= {};
  PDFV.hist = []; PDFV.fut = []; PDFV.sel = null; PDFV.pageSel = new Set(); PDFV.texts = null; PDFV.fit = true; PDFV.thumbC = new Map(); PDFV.srcText = [];
  setPdfMode('read', true);
  $('#pdfPages').innerHTML = '<div class="pdf-msg">Ouverture du PDF…</div>';
  try {
    if (cur.pdf && !cur.pdfId) { // anciens documents : PDF rangé dans le document lui-même
      const b = b64ToBytes(cur.pdf); cur.pdfId = uid(); await Store.putBlob(cur.pdfId, b); delete cur.pdf; await Store.save(cur);
    }
    const bytes = await Store.getBlob(cur.pdfId);
    if (!bytes) throw new Error('missing');
    if (state.cur !== cur) return;
    PDFV.bytes = bytes;
    await loadPdfDoc();
  } catch (e) { $('#pdfPages').innerHTML = `<div class="pdf-msg">Ce PDF ne peut pas être ouvert${e && e.name === 'PasswordException' ? ' sans son mot de passe' : e && e.message === 'missing' ? ' : le fichier est introuvable sur cet appareil' : ''}.</div>`; return; }
  if (state.cur !== cur) return;
  if (!cur.textDone) setTimeout(async () => {
    try { const t = await pdfAllText(PDFV.src); if (state.cur === cur) { cur.text = t; cur.textDone = true; scheduleSave(); } } catch {}
  }, 1500);
}
async function loadPdfDoc(keepPage) {
  const cur = state.cur, gen = ++PDFV.gen;
  const doc = await pdfjsOpen(PDFV.bytes);
  if (gen !== PDFV.gen) { doc.destroy(); return; }
  if (PDFV.src) try { PDFV.src.destroy(); } catch {}
  PDFV.src = PDFV.doc = doc; PDFV.texts = null;
  PDFV.srcPages = [];
  for (let i = 0; i < doc.numPages; i++) {
    const page = await doc.getPage(i + 1);
    PDFV.srcPages.push({ page, fields: [] });
  }
  if (!Array.isArray(cur.order) || cur.order.some(o => !(o.s >= 0 && o.s < doc.numPages))) cur.order = PDFV.srcPages.map((_, s) => ({ s, r: 0 }));
  // champs de formulaire
  await Promise.all(PDFV.srcPages.map(async S => {
    try {
      const an = await S.page.getAnnotations();
      S.fields = an.filter(a => a.subtype === 'Widget' && a.fieldName && !a.readOnly && ['Tx', 'Btn', 'Ch'].includes(a.fieldType) && !a.pushButton);
    } catch {}
  }));
  if (gen !== PDFV.gen) return;
  rebuildPages();
  $('[data-pm="form"]').hidden = !PDFV.pages.some(P => P.fields.length);
  renderPdf(keepPage);
  setTimeout(() => { if (gen === PDFV.gen) prefetchThumbs(); }, 1200);
}
function rebuildPages() {
  const cur = state.cur;
  PDFV.pages = cur.order.map(o => {
    const S = PDFV.srcPages[o.s], rot = ((S.page.rotate + o.r) % 360 + 360) % 360;
    return { s: o.s, page: S.page, rot, vp1: S.page.getViewport({ scale: 1, rotation: rot }), fields: S.fields };
  });
  cur.pagesN = PDFV.pages.length;
  PDFV.texts = null;
}
const vpOf = (P, scale) => P.page.getViewport({ scale, rotation: P.rot });
function pdfFitScale() {
  const wrap = $('#pdfScroll'), w = Math.max(...PDFV.pages.map(P => P.vp1.width), 100);
  // sur un grand écran, la page ne s'étale pas au-delà d'environ 120 % (lisible, et bien plus rapide à dessiner)
  return Math.max(0.3, Math.min(isPhone() ? 3 : 1.6, (wrap.clientWidth - (isPhone() ? 16 : 48)) / w));
}
function renderPdf(keepPage) {
  const box = $('#pdfPages');
  if (PDFV.mode === 'pages') { renderPdfGrid(); updatePdfBar(); return; }
  const top = keepPage != null ? keepPage : pdfCurrentPage();
  if (PDFV.fit) PDFV.scale = pdfFitScale();
  const S = PDFV.scale;
  box.className = 'pdf-pages';
  box.innerHTML = PDFV.pages.map((P, i) => `<div class="pp" data-p="${i}" style="width:${Math.round(P.vp1.width * S)}px;height:${Math.round(P.vp1.height * S)}px"><canvas></canvas><div class="textLayer"></div><div class="pa" data-p="${i}"></div><span class="pp-n">${i + 1}</span></div>`).join('');
  if (PDFV.obs) PDFV.obs.disconnect();
  // les pages proches de l'écran sont dessinées, les pages éloignées libèrent leur mémoire
  PDFV.obs = new IntersectionObserver(es => es.forEach(e => { const i = +e.target.dataset.p; if (e.isIntersecting) renderPdfPage(i); else releasePdfPage(e.target); }), { root: $('#pdfScroll'), rootMargin: '1200px 0px' });
  $$('.pp', box).forEach(el => PDFV.obs.observe(el));
  PDFV.pages.forEach((_, i) => drawAnns(i));
  if (top > 0) { const el = box.children[top]; if (el) $('#pdfScroll').scrollTop = el.offsetTop - 8; }
  updatePdfBar();
}
function releasePdfPage(el) {
  if (!el._done) return;
  if (el._task) try { el._task.cancel(); } catch {}
  const cv = el.querySelector('canvas'); cv.width = 0; cv.height = 0;
  const tl = el.querySelector('.textLayer'); tl.innerHTML = '';
  el._done = null; el._tl = null; el._drawn = null;
}
// file de rendu : une page à la fois, les plus proches d'abord
const RQ = { list: [], busy: false };
function renderPdfPage(i) {
  const el = $(`#pdfPages .pp[data-p="${i}"]`); if (!el) return Promise.resolve();
  const key = PDFV.scale + '|' + PDFV.gen + '|' + PDFV.mode + '|' + PDFV.pages[i].s + '|' + PDFV.pages[i].rot;
  if (el._done === key) return el._prom || Promise.resolve();
  el._done = key; el._tl = null;
  pagePlaceholder(i, el);
  el._prom = new Promise(res => { RQ.list.push({ i, el, key, res }); pumpRender(); });
  return el._prom;
}
function pageVisible(el) { const sc = $('#pdfScroll'); return el.offsetTop < sc.scrollTop + sc.clientHeight && el.offsetTop + el.offsetHeight > sc.scrollTop; }
// Ordre de travail : pages visibles, puis leur texte sélectionnable, puis les pages voisines, puis le reste du texte.
async function pumpRender() {
  if (RQ.busy) return; RQ.busy = true;
  for (;;) {
    RQ.list = RQ.list.filter(j => { if (j.el._done !== j.key || !j.el.isConnected) { j.res(); return false; } return true; });
    for (let k = TLQ.length - 1; k >= 0; k--) { const j = TLQ[k]; if (j.el._done !== j.key || !j.el.isConnected || j.el._tl === j.key) TLQ.splice(k, 1); }
    if (!RQ.list.length && !TLQ.length) break;
    const here = pdfCurrentPage(), near = (a, b) => Math.abs(a.i - here) - Math.abs(b.i - here);
    RQ.list.sort(near); TLQ.sort(near);
    let job = RQ.list.find(j => pageVisible(j.el)), kind = 'page';
    if (!job) { job = TLQ.find(j => pageVisible(j.el)); kind = 'text'; }
    if (!job && RQ.list.length) { job = RQ.list[0]; kind = 'page'; }
    if (!job) { job = TLQ[0]; kind = 'text'; }
    if (kind === 'page') { RQ.list.splice(RQ.list.indexOf(job), 1); try { await drawPdfPage(job.i, job.el); } catch {} job.res(); }
    else { TLQ.splice(TLQ.indexOf(job), 1); await drawTextLayer(job); }
  }
  RQ.busy = false;
}
async function drawPdfPage(i, el) {
  const P = PDFV.pages[i], S = PDFV.scale, dpr = Math.min(window.devicePixelRatio || 1, 2);
  const vp = vpOf(P, S);
  const cv = el.querySelector('canvas');
  // limite de taille du dessin (mémoire des téléphones)
  const k = Math.min(dpr, Math.sqrt(6e6 / (vp.width * vp.height)));
  cv.width = Math.floor(vp.width * k); cv.height = Math.floor(vp.height * k);
  cv.style.width = vp.width + 'px'; cv.style.height = vp.height + 'px';
  try {
    el._task = P.page.render({ canvasContext: cv.getContext('2d', { alpha: false }), viewport: vp, transform: k !== 1 ? [k, 0, 0, k, 0, 0] : null });
    await el._task.promise;
  } catch (e) { if (e && e.name === 'RenderingCancelledException') return; }
  // la page affichée sert aussi de miniature
  const tk = P.s + ':' + P.rot;
  if (!PDFV.thumbC.has(tk)) { PDFV.thumbC.set(tk, null); canvasToThumb(shrinkCanvas(cv)).then(u => { if (u) { PDFV.thumbC.set(tk, u); PDFV.pages.forEach((Q, j) => { if (Q.s === P.s && Q.rot === P.rot) paintThumb(j); }); } else PDFV.thumbC.delete(tk); }); }
  el.style.backgroundImage = '';
  drawFields(i);
  // le calque de texte (sélection, recherche) est préparé quand plus aucune page n'attend d'être dessinée
  el._drawn = el._done;
  if (PDFV.mode === 'read') TLQ.push({ i, el, key: el._done, vp, S });
}
const TLQ = [];
async function drawTextLayer(job) {
  job.el._tl = job.key;
  const tl = job.el.querySelector('.textLayer');
  tl.innerHTML = ''; tl.style.setProperty('--scale-factor', job.S);
  try {
    tl.dataset.mainRotation = String(PDFV.pages[job.i].rot);
    await new window.pdfjsLib.TextLayer({ textContentSource: PDFV.pages[job.i].page.streamTextContent(), container: tl, viewport: job.vp }).render();
    if (PF.q) paintPdfFind(job.i);
  } catch {}
}
// aperçu flou immédiat (miniature) en attendant le dessin net de la page
function pagePlaceholder(i, el) {
  const P = PDFV.pages[i], u = P && PDFV.thumbC.get(P.s + ':' + P.rot);
  if (u && !el.style.backgroundImage) el.style.backgroundImage = `url("${u}")`;
}
const THUMB_W = () => Math.round(130 * Math.min(window.devicePixelRatio || 1, 2));
// miniature compressée en JPEG (peu de mémoire, même pour des centaines de pages)
function canvasToThumb(c) { return new Promise(res => { try { c.toBlob(b => res(b ? URL.createObjectURL(b) : null), 'image/jpeg', 0.8); } catch { res(null); } }); }
function shrinkCanvas(src) {
  const w = THUMB_W(), h = Math.round(src.height * w / src.width);
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const g = c.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(src, 0, 0, w, h);
  return c;
}
function pdfCurrentPage() {
  const sc = $('#pdfScroll'), kids = $$('#pdfPages .pp');
  if (!kids.length) return 0;
  const mid = sc.scrollTop + sc.clientHeight / 3;
  let best = 0; kids.forEach((el, i) => { if (el.offsetTop <= mid) best = i; });
  return best;
}
function updatePdfBar() {
  const n = PDFV.pages.length;
  $('#pdfPg').textContent = PDFV.mode === 'pages' ? `${n} page${n > 1 ? 's' : ''}` : `${pdfCurrentPage() + 1} / ${n}`;
  $('#pdfZoom').textContent = Math.round(PDFV.scale / PT * 100) + ' %';
}
$('#pdfScroll').addEventListener('scroll', debounce(updatePdfBar, 80), { passive: true });
function pdfZoom(f) {
  const p = pdfCurrentPage();
  if (f === 'fit') PDFV.fit = true;
  else { PDFV.fit = false; PDFV.scale = Math.max(0.3, Math.min(4, PDFV.scale * f)); }
  renderPdf(p);
}
window.addEventListener('resize', debounce(() => { if (state.view === 'pdfv' && PDFV.fit && PDFV.mode !== 'pages') renderPdf(); }, 200));

/* --- modes : lire / annoter / pages */
function setPdfMode(m, silent) {
  const was = PDFV.mode;
  if (was !== 'pages' && $('#pdfPages .pp')) PDFV.lastPage = pdfCurrentPage();
  PDFV.mode = m; PDFV.sel = null;
  $$('#pdfv [data-pm]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.pm === m)));
  $('#pdfv').dataset.mode = m;
  $('#pdfAnnTools').hidden = m !== 'annot';
  $('#pdfPageTools').hidden = m !== 'pages';
  $('#pdfReadTools').hidden = m === 'pages';
  if (m !== 'read') closePdfFind();
  if (!silent && PDFV.doc) {
    renderPdf(m !== 'pages' && was === 'pages' ? (PDFV.pageSel.size ? Math.min(...PDFV.pageSel) : PDFV.lastPage || 0) : undefined);
    if (m === 'pages') { const c = $(`#pdfPages .pg-card[data-i="${PDFV.lastPage || 0}"]`); $('#pdfScroll').scrollTop = c ? Math.max(0, c.offsetTop - 60) : 0; }
  }
  if (m === 'annot') setPdfTool(PDFV.tool);
  updatePdfToolOpts();
}
function setPdfTool(t) {
  PDFV.tool = t;
  $$('#pdfAnnTools [data-pt]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.pt === t)));
  $('#pdfv').dataset.tool = t;
  if (t !== 'select') { PDFV.sel = null; $$('#pdfPages .an.sel').forEach(x => x.classList.remove('sel')); }
  updatePdfToolOpts();
}
function updatePdfToolOpts() {
  const t = PDFV.sel ? annById(PDFV.sel)?.t : PDFV.tool;
  const box = $('#pdfOpts');
  if (PDFV.mode !== 'annot' || !t || t === 'select' || t === 'white' || t === 'img' || t === 'sign' || t === 'image') { box.innerHTML = PDFV.sel ? `<button class="tb" data-po="del" title="Supprimer (Suppr)">${ic('trash')}</button>` : ''; return; }
  const key = t === 'hl' ? 'hl' : t === 'rect' ? 'rect' : t === 'ink' ? 'ink' : 'text';
  const a = PDFV.sel ? annById(PDFV.sel) : null;
  const curCol = a ? a.color : PDFV.color[key];
  const pal = key === 'hl' ? HL_COLORS2 : PDF_COLORS;
  let h = pal.map(([n, c]) => `<button class="pc-sw" data-pc="${c}" title="${n}" aria-label="${n}" aria-pressed="${curCol === c}" style="background:${c}"></button>`).join('');
  if (key === 'text') { const sz = a ? a.size : PDFV.size; h += `<select class="tb-select" id="pdfSize" aria-label="Taille du texte">${[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48].map(z => `<option${z === sz ? ' selected' : ''}>${z}</option>`).join('')}</select>`; }
  if (key === 'ink' || key === 'rect') { const lw = a ? a.lw : PDFV.lw; h += `<div class="seg">${[[1, 'Fin'], [2, 'Moyen'], [4, 'Épais']].map(([v, l]) => `<button data-plw="${v}" aria-pressed="${lw === v}">${l}</button>`).join('')}</div>`; }
  if (PDFV.sel) h += `<button class="tb" data-po="del" title="Supprimer (Suppr)">${ic('trash')}</button>`;
  box.innerHTML = h;
  const sz = $('#pdfSize'); if (sz) sz.addEventListener('change', () => { const v = +sz.value; if (a) { pdfHist(); a.size = v; drawAnns(a.p); pdfChanged(); } else PDFV.size = v; });
}
$('#pdfOpts').addEventListener('click', e => {
  const sw = e.target.closest('[data-pc]'), lw = e.target.closest('[data-plw]'), po = e.target.closest('[data-po]');
  const a = PDFV.sel ? annById(PDFV.sel) : null;
  if (sw) { const t = a ? a.t : PDFV.tool, key = t === 'hl' ? 'hl' : t === 'rect' ? 'rect' : t === 'ink' ? 'ink' : 'text'; if (a) { pdfHist(); a.color = sw.dataset.pc; drawAnns(a.p); pdfChanged(); } else PDFV.color[key] = sw.dataset.pc; updatePdfToolOpts(); }
  if (lw) { const v = +lw.dataset.plw; if (a) { pdfHist(); a.lw = v; drawAnns(a.p); pdfChanged(); } else PDFV.lw = v; updatePdfToolOpts(); }
  if (po && po.dataset.po === 'del' && a) deleteAnn(a.id);
});

/* --- annotations */
const annById = id => (state.cur.ann || []).find(a => a.id === id);
const pdfSnap = () => ({ bytes: PDFV.bytes, order: JSON.stringify(state.cur.order), ann: JSON.stringify(state.cur.ann), form: JSON.stringify(state.cur.form) });
function pdfHist() { PDFV.hist.push(pdfSnap()); if (PDFV.hist.length > 40) PDFV.hist.shift(); PDFV.fut = []; }
function pdfChanged() { scheduleSave(); }
async function pdfUndo(redo) {
  const from = redo ? PDFV.fut : PDFV.hist, to = redo ? PDFV.hist : PDFV.fut;
  const s = from.pop(); if (!s) return;
  const cur = state.cur;
  to.push(pdfSnap());
  const bytesChanged = s.bytes && s.bytes !== PDFV.bytes, orderChanged = s.order !== JSON.stringify(cur.order);
  cur.order = JSON.parse(s.order); cur.ann = JSON.parse(s.ann); cur.form = JSON.parse(s.form); PDFV.sel = null; PDFV.pageSel = new Set();
  if (bytesChanged) { PDFV.bytes = s.bytes; await Store.putBlob(cur.pdfId, s.bytes); await loadPdfDoc(); }
  else if (orderChanged) { rebuildPages(); renderPdf(); }
  else { PDFV.pages.forEach((_, i) => { drawAnns(i); drawFields(i); }); if (PDFV.mode === 'pages') renderPdfGrid(); }
  scheduleSave();
}
function deleteAnn(id) {
  const a = annById(id); if (!a) return;
  pdfHist(); state.cur.ann = state.cur.ann.filter(x => x.id !== id);
  PDFV.sel = null; drawAnns(a.p); updatePdfToolOpts(); pdfChanged();
}
function annBBox(a) {
  if (a.t === 'ink') { const xs = a.pts.map(p => p[0]), ys = a.pts.map(p => p[1]); const m = (a.lw || 2) / 2 + 2; return { x: Math.min(...xs) - m, y: Math.min(...ys) - m, w: Math.max(...xs) - Math.min(...xs) + 2 * m, h: Math.max(...ys) - Math.min(...ys) + 2 * m }; }
  return { x: a.x, y: a.y, w: a.w, h: a.h };
}
function drawAnns(i) {
  const pa = $(`#pdfPages .pa[data-p="${i}"]`); if (!pa) return;
  const S = PDFV.scale;
  const editing = pa.querySelector('.an.editing');
  if (editing) { PDFV.redrawAfterEdit = true; return; }
  pa.querySelectorAll('.an').forEach(x => { if (x.parentNode) x.remove(); });
  (state.cur.ann || []).filter(a => a.p === i).forEach(a => {
    const b = annBBox(a), el = document.createElement('div');
    el.className = 'an an-' + a.t + (PDFV.sel === a.id ? ' sel' : '');
    el.dataset.id = a.id;
    el.style.cssText = `left:${b.x * S}px;top:${b.y * S}px;width:${b.w * S}px;` + (a.t === 'text' ? '' : `height:${b.h * S}px;`);
    if (a.t === 'text') {
      el.style.fontSize = (a.size * S) + 'px'; el.style.color = a.color;
      const tx = document.createElement('div'); tx.className = 'an-tx'; tx.textContent = a.text; el.appendChild(tx);
    } else if (a.t === 'hl') el.style.background = a.color;
    else if (a.t === 'rect') el.style.border = `${Math.max(1, a.lw * S)}px solid ${a.color}`;
    else if (a.t === 'white') el.style.background = '#fff';
    else if (a.t === 'img') el.innerHTML = `<img src="${a.src}" alt="" draggable="false">`;
    else if (a.t === 'ink') {
      const d = a.pts.map((p, k) => (k ? 'L' : 'M') + ((p[0] - b.x) * S).toFixed(1) + ',' + ((p[1] - b.y) * S).toFixed(1)).join('');
      el.innerHTML = `<svg width="${b.w * S}" height="${b.h * S}"><path d="${d}" fill="none" stroke="${a.color}" stroke-width="${Math.max(1, a.lw * S)}" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
    }
    if (PDFV.sel === a.id && a.t !== 'ink') { const h = document.createElement('span'); h.className = 'an-rz'; el.appendChild(h); }
    pa.appendChild(el);
    if (a.t === 'text') { const hh = el.offsetHeight / S; if (Math.abs((a.h || 0) - hh) > 0.5) a.h = hh; }
  });
}
function selectAnn(id) {
  PDFV.sel = id;
  $$('#pdfPages .an').forEach(x => x.classList.toggle('sel', x.dataset.id === id));
  const a = id && annById(id); if (a) drawAnns(a.p);
  updatePdfToolOpts();
}
function pagePoint(e, pa) { const r = pa.getBoundingClientRect(); return [(e.clientX - r.left) / PDFV.scale, (e.clientY - r.top) / PDFV.scale]; }
function editTextAnn(a) {
  const el = $(`#pdfPages .an[data-id="${a.id}"]`); if (!el) return;
  const tx = el.querySelector('.an-tx');
  try { tx.contentEditable = 'plaintext-only'; } catch { tx.contentEditable = 'true'; }
  if (tx.contentEditable !== 'plaintext-only') tx.contentEditable = 'true';
  el.classList.add('editing');
  tx.focus();
  const r = document.createRange(); r.selectNodeContents(tx); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  const before = a.text;
  const done = () => {
    tx.removeEventListener('blur', done);
    const t = tx.innerText.replace(/\n$/, '');
    el.classList.remove('editing'); tx.contentEditable = 'false';
    if (!t.trim()) { state.cur.ann = state.cur.ann.filter(x => x.id !== a.id); if (PDFV.sel === a.id) PDFV.sel = null; setTimeout(() => drawAnns(a.p), 0); pdfChanged(); return; }
    if (t !== before) { if (before) pdfHist(); a.text = t; pdfChanged(); }
    setTimeout(() => drawAnns(a.p), 0);
  };
  tx.addEventListener('blur', done);
  tx.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); tx.blur(); } e.stopPropagation(); });
}
// Dessin et manipulation à la souris ou au doigt
const PD = { op: null };
$('#pdfPages').addEventListener('pointerdown', e => {
  if (PDFV.mode !== 'annot') return;
  const pa = e.target.closest('.pa'); if (!pa) return;
  if (e.target.closest('.editing')) return;
  // un texte en cours d'écriture : un clic ailleurs termine simplement la saisie
  const ed = $('#pdfPages .an.editing .an-tx'); if (ed) { e.preventDefault(); ed.blur(); return; }
  const i = +pa.dataset.p, pt = pagePoint(e, pa), tool = PDFV.tool;
  const hit = e.target.closest('.an');
  if (hit && (tool === 'select' || tool === 'text')) {
    const a = annById(hit.dataset.id); if (!a) return;
    e.preventDefault();
    if (PDFV.sel !== a.id) selectAnn(a.id);
    PD.op = { kind: e.target.classList.contains('an-rz') ? 'size' : 'move', a, i, start: pt, orig: JSON.parse(JSON.stringify(a)), moved: false, pa };
    try { pa.setPointerCapture(e.pointerId); } catch {}
    return;
  }
  if (tool === 'select') { if (PDFV.sel) selectAnn(null); return; }
  if (tool === 'text') {
    e.preventDefault();
    pdfHist();
    const a = { id: uid(), t: 'text', p: i, x: pt[0], y: pt[1] - PDFV.size * 0.6, w: Math.min(260, PDFV.pages[i].vp1.width - pt[0] - 8), h: PDFV.size * 1.25, text: '', size: PDFV.size, color: PDFV.color.text };
    a.w = Math.max(60, a.w);
    state.cur.ann.push(a); PDFV.sel = a.id; drawAnns(i); updatePdfToolOpts();
    setTimeout(() => editTextAnn(a), 0);
    return;
  }
  if (['hl', 'rect', 'white', 'ink'].includes(tool)) {
    e.preventDefault();
    try { pa.setPointerCapture(e.pointerId); } catch {}
    PD.op = { kind: 'draw', tool, i, start: pt, pts: [pt], pa };
    const g = document.createElement('div'); g.className = 'an-ghost an-' + tool; pa.appendChild(g); PD.op.ghost = g;
    if (tool === 'ink') g.innerHTML = '<svg></svg>';
  }
});
$('#pdfPages').addEventListener('pointermove', e => {
  const op = PD.op; if (!op) return;
  const pt = pagePoint(e, op.pa), S = PDFV.scale;
  if (op.kind === 'draw') {
    const g = op.ghost;
    if (op.tool === 'ink') {
      const last = op.pts[op.pts.length - 1];
      if (Math.hypot(pt[0] - last[0], pt[1] - last[1]) * S < 2) return;
      op.pts.push(pt);
      const W = op.pa.clientWidth, H = op.pa.clientHeight;
      g.style.cssText = 'left:0;top:0;width:100%;height:100%';
      g.innerHTML = `<svg width="${W}" height="${H}"><path d="${op.pts.map((p, k) => (k ? 'L' : 'M') + (p[0] * S).toFixed(1) + ',' + (p[1] * S).toFixed(1)).join('')}" fill="none" stroke="${PDFV.color.ink}" stroke-width="${Math.max(1, PDFV.lw * S)}" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
    } else {
      const x = Math.min(op.start[0], pt[0]), y = Math.min(op.start[1], pt[1]), w = Math.abs(pt[0] - op.start[0]), h = Math.abs(pt[1] - op.start[1]);
      g.style.cssText = `left:${x * S}px;top:${y * S}px;width:${w * S}px;height:${h * S}px;` + (op.tool === 'hl' ? `background:${PDFV.color.hl}` : op.tool === 'rect' ? `border:${Math.max(1, PDFV.lw * S)}px solid ${PDFV.color.rect}` : 'background:#fff;outline:1px dashed var(--selline)');
      op.box = { x, y, w, h };
    }
    return;
  }
  const dx = pt[0] - op.start[0], dy = pt[1] - op.start[1];
  if (!op.moved && Math.hypot(dx, dy) * S < 3) return;
  if (!op.moved) { op.moved = true; pdfHist(); }
  const a = op.a, o = op.orig;
  if (op.kind === 'move') {
    if (a.t === 'ink') a.pts = o.pts.map(p => [p[0] + dx, p[1] + dy]);
    else { a.x = o.x + dx; a.y = o.y + dy; }
  } else {
    a.w = Math.max(16, o.w + dx);
    if (a.t !== 'text') a.h = Math.max(8, a.t === 'img' ? o.h * (a.w / o.w) : o.h + dy);
  }
  drawAnns(a.p);
});
function endPdfPointer() {
  const op = PD.op; PD.op = null; if (!op) return;
  if (op.kind === 'draw') {
    op.ghost.remove();
    let a = null;
    if (op.tool === 'ink') { if (op.pts.length > 1) a = { id: uid(), t: 'ink', p: op.i, pts: op.pts.map(p => [+p[0].toFixed(2), +p[1].toFixed(2)]), color: PDFV.color.ink, lw: PDFV.lw }; }
    else if (op.box && op.box.w > 3 && op.box.h > 3) a = { id: uid(), t: op.tool, p: op.i, ...op.box, color: op.tool === 'hl' ? PDFV.color.hl : PDFV.color.rect, lw: PDFV.lw };
    if (a) { pdfHist(); state.cur.ann.push(a); drawAnns(op.i); pdfChanged(); }
    return;
  }
  if (op.moved) pdfChanged();
  else if (op.a.t === 'text' && (PDFV.tool === 'text' || PDFV.tool === 'select')) editTextAnn(op.a);
}
$('#pdfPages').addEventListener('pointerup', endPdfPointer);
$('#pdfPages').addEventListener('pointercancel', endPdfPointer);
$('#pdfPages').addEventListener('dblclick', e => { const hit = e.target.closest('.an-text'); if (hit && PDFV.mode === 'annot') { const a = annById(hit.dataset.id); if (a) editTextAnn(a); } });
document.addEventListener('keydown', e => {
  if (state.view !== 'pdfv') return;
  const inField = e.target.closest && e.target.closest('input,textarea,select,[contenteditable="true"],[contenteditable="plaintext-only"]');
  const mod = e.ctrlKey || e.metaKey;
  if (!inField && (e.key === 'Delete' || e.key === 'Backspace')) {
    if (PDFV.mode === 'annot' && PDFV.sel) { e.preventDefault(); deleteAnn(PDFV.sel); }
    else if (PDFV.mode === 'pages' && PDFV.pageSel.size) { e.preventDefault(); pagesOp('del'); }
  }
  if (!inField && mod && (e.key === 'z' || e.key === 'Z')) { e.preventDefault(); pdfUndo(e.shiftKey); }
  if (!inField && mod && (e.key === 'y' || e.key === 'Y')) { e.preventDefault(); pdfUndo(true); }
  if (mod && (e.key === 'f' || e.key === 'F')) { e.preventDefault(); openPdfFind(); }
  if (!inField && mod && (e.key === 'a' || e.key === 'A') && PDFV.mode === 'pages') { e.preventDefault(); PDFV.pageSel = new Set(PDFV.pages.map((_, i) => i)); renderPdfGrid(); }
  if (!inField && e.key === 'Escape' && PDFV.sel) selectAnn(null);
});

/* --- signature et images */
function addImageAnn(src, w0, h0) {
  const i = Math.max(0, pdfCurrentPage()), P = PDFV.pages[i];
  const w = Math.min(w0, P.vp1.width * 0.5), h = h0 * (w / w0);
  const sc = $('#pdfScroll'), el = $(`#pdfPages .pp[data-p="${i}"]`);
  const visTop = el ? Math.max(0, (sc.scrollTop + sc.clientHeight / 2 - el.offsetTop) / PDFV.scale - h / 2) : 40;
  pdfHist();
  const a = { id: uid(), t: 'img', p: i, x: (P.vp1.width - w) / 2, y: Math.min(P.vp1.height - h - 10, visTop), w, h, src };
  state.cur.ann.push(a); setPdfTool('select'); selectAnn(a.id); drawAnns(i); pdfChanged();
  toast('Faites glisser pour placer, coin en bas à droite pour la taille.');
}
function imgSize(src) { return new Promise(res => { const im = new Image(); im.onload = () => res([im.naturalWidth, im.naturalHeight, im]); im.onerror = () => res([200, 80, null]); im.src = src; }); }
async function fileToImageData(f, max = 2000) {
  const url = URL.createObjectURL(f);
  try {
    const [w, h, im] = await imgSize(url); if (!im) throw new Error('image');
    const k = Math.min(1, max / Math.max(w, h)), c = document.createElement('canvas');
    c.width = Math.round(w * k); c.height = Math.round(h * k);
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(im, 0, 0, c.width, c.height);
    const png = /png|gif/i.test(f.type);
    return { src: png ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.86), w: c.width, h: c.height, png };
  } finally { URL.revokeObjectURL(url); }
}
function openSignPad() {
  const saved = ls.get('grim.sig', null);
  const p = $('#pdfDlg');
  p.innerHTML = `<div class="dlg-in"><div class="fp-h"><b>Signature</b><button class="xbtn" data-dx aria-label="Fermer">${ic('close')}</button></div>
    ${saved ? `<button class="sig-saved" id="sigUse"><img src="${saved}" alt="Ma signature"><span>Utiliser ma signature enregistrée</span></button><hr class="dlg-hr">` : ''}
    <p class="dlg-note">Signez dans le cadre avec le doigt ou la souris.</p>
    <canvas class="sig-pad" id="sigPad" width="600" height="220"></canvas>
    <label class="fp-it"><input type="checkbox" id="sigKeep" checked><span>Enregistrer sur cet appareil pour la prochaine fois</span></label>
    <div class="fp-acts"><button class="pbtn" id="sigOk">Placer la signature</button><button class="pbtn ghost" id="sigClr">Effacer</button></div></div>`;
  p.hidden = false;
  const cv = $('#sigPad'), g = cv.getContext('2d');
  g.lineWidth = 3.2; g.lineCap = 'round'; g.lineJoin = 'round'; g.strokeStyle = '#1B2A4A';
  let drawing = false, any = false, last = null;
  const pos = e => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) * cv.width / r.width, (e.clientY - r.top) * cv.height / r.height]; };
  cv.addEventListener('pointerdown', e => { e.preventDefault(); drawing = true; last = pos(e); try { cv.setPointerCapture(e.pointerId); } catch {} });
  cv.addEventListener('pointermove', e => { if (!drawing) return; const q = pos(e); g.beginPath(); g.moveTo(last[0], last[1]); g.lineTo(q[0], q[1]); g.stroke(); last = q; any = true; });
  const up = () => { drawing = false; }; cv.addEventListener('pointerup', up); cv.addEventListener('pointercancel', up);
  $('#sigClr').addEventListener('click', () => { g.clearRect(0, 0, cv.width, cv.height); any = false; });
  $('#sigOk').addEventListener('click', async () => {
    if (!any) return toast('Signez d\'abord dans le cadre.', true);
    // recadrage sur le tracé
    const d = g.getImageData(0, 0, cv.width, cv.height).data; let x1 = cv.width, y1 = cv.height, x2 = 0, y2 = 0;
    for (let y = 0; y < cv.height; y++) for (let x = 0; x < cv.width; x++) if (d[(y * cv.width + x) * 4 + 3] > 10) { if (x < x1) x1 = x; if (x > x2) x2 = x; if (y < y1) y1 = y; if (y > y2) y2 = y; }
    const m = 6, c2 = document.createElement('canvas'); c2.width = x2 - x1 + 2 * m; c2.height = y2 - y1 + 2 * m;
    c2.getContext('2d').drawImage(cv, x1 - m, y1 - m, c2.width, c2.height, 0, 0, c2.width, c2.height);
    const src = c2.toDataURL('image/png');
    if ($('#sigKeep').checked) ls.set('grim.sig', src);
    p.hidden = true; addImageAnn(src, Math.min(180, c2.width / 2), c2.height / 2 * Math.min(1, 180 / (c2.width / 2)));
  });
  const use = $('#sigUse'); if (use) use.addEventListener('click', async () => { const [w, h] = await imgSize(saved); p.hidden = true; addImageAnn(saved, Math.min(180, w / 2), h / 2 * Math.min(1, 180 / (w / 2))); });
}
$('#pdfDlg').addEventListener('click', e => { if (e.target.closest('[data-dx]')) $('#pdfDlg').hidden = true; });

/* --- formulaires */
function fieldKey(f) { return f.fieldName; }
function drawFields(i) {
  const pa = $(`#pdfPages .pa[data-p="${i}"]`); if (!pa) return;
  pa.querySelectorAll('.fld').forEach(x => x.remove());
  const P = PDFV.pages[i]; if (!P || !P.fields.length || PDFV.mode !== 'read') return;
  const vp = vpOf(P, PDFV.scale), vals = state.cur.form;
  P.fields.forEach(f => {
    const [x1, y1, x2, y2] = vp.convertToViewportRectangle(f.rect);
    const x = Math.min(x1, x2), y = Math.min(y1, y2), w = Math.abs(x2 - x1), h = Math.abs(y2 - y1);
    const k = fieldKey(f), v = vals[k];
    let el;
    if (f.fieldType === 'Tx') {
      el = document.createElement(f.multiLine ? 'textarea' : 'input');
      el.value = v != null ? v : (f.fieldValue || '');
      if (f.maxLen) el.maxLength = f.maxLen;
      el.style.fontSize = Math.max(9, Math.min(h * 0.62, 16 * PDFV.scale / PT)) + 'px';
    } else if (f.fieldType === 'Btn') {
      el = document.createElement('input');
      el.type = f.radioButton ? 'radio' : 'checkbox';
      if (f.radioButton) { el.name = 'r_' + k; el.checked = v != null ? v === f.buttonValue : f.fieldValue === f.buttonValue; el.value = f.buttonValue; }
      else el.checked = v != null ? !!v : (f.fieldValue && f.fieldValue !== 'Off');
    } else {
      el = document.createElement('select');
      const opts = (f.options || []).map(o => [o.exportValue ?? o.displayValue, o.displayValue ?? o.exportValue]);
      el.innerHTML = '<option value=""></option>' + opts.map(([ev, dv]) => `<option value="${esc(String(ev))}">${esc(String(dv))}</option>`).join('');
      el.value = v != null ? v : Array.isArray(f.fieldValue) ? (f.fieldValue[0] || '') : (f.fieldValue || '');
    }
    el.className = 'fld fld-' + f.fieldType;
    el.dataset.k = k; el.title = f.alternativeText || k;
    el.style.left = x + 'px'; el.style.top = y + 'px'; el.style.width = w + 'px'; el.style.height = h + 'px';
    const save = () => {
      pdfHist();
      if (f.fieldType === 'Btn') {
        if (f.radioButton) { if (el.checked) vals[k] = f.buttonValue; }
        else vals[k] = el.checked;
      } else vals[k] = el.value;
      // même champ affiché plusieurs fois : on met tout à jour
      $$(`#pdfPages .fld[data-k="${CSS.escape(k)}"]`).forEach(o => { if (o !== el && f.fieldType !== 'Btn') o.value = el.value; });
      pdfChanged();
    };
    el.addEventListener('change', save);
    if (f.fieldType === 'Tx') el.addEventListener('input', debounce(save, 400));
    pa.appendChild(el);
  });
}

/* --- vue Pages : organiser, faire pivoter, supprimer, extraire, diviser, insérer */
function renderPdfGrid() {
  const box = $('#pdfPages');
  if (PDFV.obs) PDFV.obs.disconnect();
  box.className = 'pdf-grid';
  const annP = new Set((state.cur.ann || []).map(a => a.p));
  box.innerHTML = PDFV.pages.map((P, i) => `<div class="pg-card${PDFV.pageSel.has(i) ? ' on' : ''}" data-i="${i}"><div class="pg-th" style="aspect-ratio:${P.vp1.width}/${P.vp1.height}"><img class="pg-cv" alt="" draggable="false">${annP.has(i) ? '<span class="pg-ann" title="Page annotée">✎</span>' : ''}</div><span class="pg-n">${i + 1}</span><span class="pg-chk" aria-hidden="true"></span></div>`).join('');
  updatePageSelUI();
  // miniatures déjà prêtes : affichées tout de suite ; les autres sont demandées dans l'ordre d'apparition à l'écran
  if (PDFV.gridObs) PDFV.gridObs.disconnect();
  PDFV.gridObs = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) { PDFV.gridObs.unobserve(e.target); needThumb(+e.target.dataset.i); } }), { root: $('#pdfScroll'), rootMargin: '400px 0px' });
  $$('.pg-card', box).forEach(c => { const i = +c.dataset.i; if (!paintThumb(i)) PDFV.gridObs.observe(c); });
  updatePdfBar();
}
function updatePageSelUI() {
  const n = PDFV.pageSel.size;
  $$('#pdfPages .pg-card').forEach(c => c.classList.toggle('on', PDFV.pageSel.has(+c.dataset.i)));
  $('#pdfSelInfo').textContent = n ? `${n} page${n > 1 ? 's' : ''} sélectionnée${n > 1 ? 's' : ''}` : 'Touchez des pages pour les sélectionner';
  $$('#pdfPageTools [data-pg]').forEach(b => { if (b.dataset.need) b.disabled = !n; });
}
function paintThumb(i) {
  const P = PDFV.pages[i]; if (!P) return false;
  const url = PDFV.thumbC.get(P.s + ':' + P.rot); if (!url) return false;
  const im = $(`#pdfPages .pg-card[data-i="${i}"] .pg-cv`); if (!im) return true;
  if (im.getAttribute('src') !== url) { im.onload = () => im.classList.add('ok'); im.src = url; }
  return true;
}
// Miniatures : celles qu'on voit d'abord (file prioritaire), les autres en arrière-plan quand rien d'autre ne se dessine.
const TQ = { hi: [], lo: [], busy: false };
function needThumb(i) { TQ.hi.push(i); pumpThumbs(); }
function prefetchThumbs() { TQ.lo = PDFV.pages.map((_, i) => i); pumpThumbs(); }
const pause = ms => new Promise(r => setTimeout(r, ms));
async function pumpThumbs() {
  if (TQ.busy) return; TQ.busy = true;
  const gen = PDFV.gen;
  while ((TQ.hi.length || TQ.lo.length) && gen === PDFV.gen && state.view === 'pdfv') {
    let i;
    if (TQ.hi.length) i = TQ.hi.shift();
    else {
      // l'affichage des pages passe avant la préparation des miniatures
      if (RQ.busy || RQ.list.length || PD.op) { await pause(250); continue; }
      i = TQ.lo.shift();
    }
    const P = PDFV.pages[i]; if (!P) continue;
    const tk = P.s + ':' + P.rot;
    if (!PDFV.thumbC.has(tk)) {
      try {
        PDFV.thumbC.set(tk, null);
        const vp = vpOf(P, THUMB_W() / P.vp1.width);
        const c = document.createElement('canvas'); c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
        await P.page.render({ canvasContext: c.getContext('2d', { alpha: false }), viewport: vp }).promise;
        const url = await canvasToThumb(c);
        if (url) PDFV.thumbC.set(tk, url); else PDFV.thumbC.delete(tk);
      } catch { PDFV.thumbC.delete(tk); continue; }
    }
    // toutes les cartes qui montrent cette page (pages dupliquées)
    if (PDFV.mode === 'pages') PDFV.pages.forEach((Q, j) => { if (Q.s === P.s && Q.rot === P.rot) paintThumb(j); });
  }
  TQ.busy = false;
}
// sélection et glisser-déposer pour réordonner
const PG = { drag: null, last: null };
$('#pdfPages').addEventListener('pointerdown', e => {
  if (PDFV.mode !== 'pages') return;
  const card = e.target.closest('.pg-card'); if (!card) return;
  PG.drag = { i: +card.dataset.i, x: e.clientX, y: e.clientY, on: false, card, shift: e.shiftKey, ctrl: e.ctrlKey || e.metaKey };
});
$('#pdfPages').addEventListener('pointermove', e => {
  const d = PG.drag; if (!d || PDFV.mode !== 'pages') return;
  if (!d.on && Math.hypot(e.clientX - d.x, e.clientY - d.y) > 8) {
    if (e.pointerType === 'touch' && !d.hold) { clearTimeout(d.holdT); PG.drag = null; return; }
    d.on = true; d.card.classList.add('dragging');
    try { d.card.setPointerCapture(e.pointerId); } catch {}
  }
  if (!d.on) return;
  e.preventDefault();
  const el = document.elementFromPoint(e.clientX, e.clientY), over = el && el.closest('.pg-card');
  if (d.over && d.over !== over) d.over.classList.remove('drop-b', 'drop-a');
  if (over && over !== d.card) { const r = over.getBoundingClientRect(); const after = e.clientX > r.left + r.width / 2; over.classList.toggle('drop-a', after); over.classList.toggle('drop-b', !after); d.to = +over.dataset.i + (after ? 1 : 0); d.over = over; }
  else { d.to = null; d.over = null; }
  const sc = $('#pdfScroll').getBoundingClientRect();
  if (e.clientY > sc.bottom - 40) $('#pdfScroll').scrollTop += 16; else if (e.clientY < sc.top + 40) $('#pdfScroll').scrollTop -= 16;
});
// appui long au doigt pour déplacer une page
$('#pdfPages').addEventListener('touchstart', () => { const d = PG.drag; if (!d) return; d.holdT = setTimeout(() => { if (PG.drag === d) { d.hold = true; d.card.classList.add('lift'); navigator.vibrate && navigator.vibrate(15); } }, 380); }, { passive: true });
$('#pdfPages').addEventListener('touchmove', e => { const d = PG.drag; if (d && d.hold) e.preventDefault(); else if (d) clearTimeout(d.holdT); }, { passive: false });
function endPgDrag() {
  const d = PG.drag; PG.drag = null; if (!d) return;
  clearTimeout(d.holdT);
  d.card.classList.remove('dragging', 'lift'); if (d.over) d.over.classList.remove('drop-b', 'drop-a');
  if (d.on) {
    if (d.to == null) return;
    const moving = PDFV.pageSel.has(d.i) ? [...PDFV.pageSel].sort((a, b) => a - b) : [d.i];
    const rest = PDFV.pages.map((_, i) => i).filter(i => !moving.includes(i));
    const at = rest.filter(i => i < d.to).length;
    const order = [...rest.slice(0, at), ...moving, ...rest.slice(at)];
    if (order.every((v, k) => v === k)) return;
    applyOrder(order, null, 'Pages déplacées');
    return;
  }
  if (d.hold) return;
  // simple touche : sélection
  const i = d.i;
  if (d.shift && PG.last != null) { const [a, b] = [Math.min(PG.last, i), Math.max(PG.last, i)]; for (let k = a; k <= b; k++) PDFV.pageSel.add(k); }
  else if (PDFV.pageSel.has(i)) PDFV.pageSel.delete(i); else PDFV.pageSel.add(i);
  PG.last = i;
  updatePageSelUI();
}
$('#pdfPages').addEventListener('pointerup', e => { if (PDFV.mode === 'pages') endPgDrag(e); });
$('#pdfPages').addEventListener('pointercancel', () => { if (PDFV.mode === 'pages') { const d = PG.drag; if (d && !d.hold) { clearTimeout(d.holdT); PG.drag = null; } } });
$('#pdfPages').addEventListener('dblclick', e => { if (PDFV.mode !== 'pages') return; const c = e.target.closest('.pg-card'); if (c) { const i = +c.dataset.i; PDFV.lastPage = i; PDFV.pageSel = new Set(); setPdfMode('read'); } });

// Nouvel ordre des pages (doublons possibles) et rotations : instantané, le fichier n'est pas réécrit.
function applyOrder(order, rot, msg) {
  const cur = state.cur;
  pdfHist();
  const old = cur.order, oldPages = PDFV.pages;
  cur.order = order.map((k, j) => ({ s: old[k].s, r: (((old[k].r + ((rot && rot[j]) || 0)) % 360) + 360) % 360 }));
  // annotations : suivent leur page, sont dupliquées avec elle, et tournent avec elle
  const ann = [], firstSeen = new Set(), byP = new Map();
  cur.ann.forEach(a => { if (!byP.has(a.p)) byP.set(a.p, []); byP.get(a.p).push(a); });
  order.forEach((k, j) => {
    (byP.get(k) || []).forEach(a => {
      const c = firstSeen.has(k) ? { ...JSON.parse(JSON.stringify(a)), id: uid() } : a;
      c.p = j;
      if (rot && rot[j]) rotateAnn(c, oldPages[k], rot[j]);
      ann.push(c);
    });
    firstSeen.add(k);
  });
  cur.ann = ann;
  PDFV.pageSel = new Set();
  rebuildPages();
  renderPdf();
  scheduleSave();
  if (msg) toast(msg);
}
// Rotation d'une annotation avec sa page (d = +90 ou -90) : passage par les coordonnées du PDF
function rotateAnn(a, P, d) {
  if (!P) return;
  const vA = P.page.getViewport({ scale: 1, rotation: P.rot }), vB = P.page.getViewport({ scale: 1, rotation: ((P.rot + d) % 360 + 360) % 360 });
  const cv = (x, y) => { const [px, py] = vA.convertToPdfPoint(x, y); return vB.convertToViewportPoint(px, py); };
  if (a.t === 'ink') a.pts = a.pts.map(p => cv(p[0], p[1]));
  else {
    const [x1, y1] = cv(a.x, a.y), [x2, y2] = cv(a.x + a.w, a.y + (a.h || 0));
    if (a.t === 'text') { a.x = Math.min(x1, x2); a.y = Math.min(y1, y2); }
    else { a.x = Math.min(x1, x2); a.y = Math.min(y1, y2); a.w = Math.abs(x2 - x1); a.h = Math.abs(y2 - y1); }
  }
}
// Ajout de pages (fichiers, page blanche) : elles sont ajoutées à la fin du fichier source, puis placées dans l'ordre voulu.
async function appendSourcePages(at, build) {
  const PDFLib = await ensurePdfLib();
  const cur = state.cur;
  try {
    const pdf = await PDFLib.PDFDocument.load(PDFV.bytes, { ignoreEncryption: true });
    const n0 = pdf.getPageCount();
    await build(pdf, PDFLib);
    const added = pdf.getPageCount() - n0;
    if (!added) return 0;
    const bytes = await pdf.save({ useObjectStreams: true });
    pdfHist();
    PDFV.hist[PDFV.hist.length - 1].bytes = PDFV.bytes;
    PDFV.bytes = bytes;
    await Store.putBlob(cur.pdfId, bytes);
    const ins = Array.from({ length: added }, (_, k) => ({ s: n0 + k, r: 0 }));
    cur.order = [...cur.order.slice(0, at), ...ins, ...cur.order.slice(at)];
    cur.ann.forEach(a => { if (a.p >= at) a.p += added; });
    cur.textDone = false;
    PDFV.pageSel = new Set();
    await loadPdfDoc();
    scheduleSave();
    return added;
  } catch (e) { console.warn(e); toast('Ajout impossible : ce PDF est peut-être protégé ou endommagé.', true); return 0; }
}
const selOrAll = () => PDFV.pageSel.size ? [...PDFV.pageSel].sort((a, b) => a - b) : [];
async function pagesOp(k) {
  const n = PDFV.pages.length, sel = selOrAll(), all = PDFV.pages.map((_, i) => i);
  if (k === 'all') { PDFV.pageSel = PDFV.pageSel.size === n ? new Set() : new Set(all); return updatePageSelUI(); }
  if (k === 'rotL' || k === 'rotR') {
    if (!sel.length) return toast('Sélectionnez les pages à faire pivoter.');
    const rot = {}; sel.forEach(i => { rot[i] = k === 'rotR' ? 90 : -90; });
    applyOrder(all, rot, 'Pages pivotées'); PDFV.pageSel = new Set(sel); return updatePageSelUI();
  }
  if (k === 'del') {
    if (!sel.length) return;
    if (sel.length === n) return toast('Un PDF garde au moins une page.', true);
    return applyOrder(all.filter(i => !PDFV.pageSel.has(i)), null, `${sel.length} page${sel.length > 1 ? 's supprimées' : ' supprimée'}`);
  }
  if (k === 'dup') {
    if (!sel.length) return;
    const order = []; all.forEach(i => { order.push(i); if (PDFV.pageSel.has(i)) order.push(i); });
    return applyOrder(order, null, 'Pages dupliquées');
  }
  if (k === 'blank') {
    const at = sel.length ? sel[sel.length - 1] + 1 : n;
    const ref = PDFV.pages[Math.max(0, at - 1)].vp1;
    const ok = await appendSourcePages(at, pdf => { pdf.addPage([ref.width, ref.height]); });
    if (ok) toast('Page blanche ajoutée');
    return;
  }
  if (k === 'insert') { PDFV.insertAt = sel.length ? sel[sel.length - 1] + 1 : n; $('#pdfIns').click(); return; }
  if (k === 'extract') { if (!sel.length) return toast('Sélectionnez les pages à extraire.'); return splitPdf([sel], 'grimoire', 'extrait'); }
  if (k === 'split') return openSplit();
  if (k === 'move') {
    if (!sel.length) return;
    const to = parseInt(window.prompt(`Déplacer ${sel.length > 1 ? 'les pages' : 'la page'} avant la page n° (1 à ${n + 1}) :`, '1'), 10);
    if (!(to >= 1 && to <= n + 1)) return;
    const rest = all.filter(i => !PDFV.pageSel.has(i)), at = rest.filter(i => i < to - 1).length;
    return applyOrder([...rest.slice(0, at), ...sel, ...rest.slice(at)], null, 'Pages déplacées');
  }
}
$('#pdfPageTools').addEventListener('click', e => { const b = e.target.closest('[data-pg]'); if (b && !b.disabled) pagesOp(b.dataset.pg); });
// Ajout de pages : PDF et images
async function pagesFromFiles(pdf, PDFLib, files) {
  const out = [];
  for (const f of files) {
    const isPdf = /pdf$/i.test(f.type) || /\.pdf$/i.test(f.name);
    if (isPdf) {
      try {
        const src = await PDFLib.PDFDocument.load(new Uint8Array(await f.arrayBuffer()), { ignoreEncryption: true });
        const cp = await pdf.copyPages(src, src.getPageIndices());
        cp.forEach(p => out.push({ kind: 'ext', page: p }));
      } catch { toast(`« ${f.name} » n'a pas pu être ajouté (PDF protégé ou endommagé).`, true); }
    } else if (/^image\//.test(f.type) || /\.(jpe?g|png|gif|webp|heic)$/i.test(f.name)) {
      try {
        const im = await fileToImageData(f);
        const bytes = b64ToBytes(im.src.split(',')[1]);
        const emb = im.png ? await pdf.embedPng(bytes) : await pdf.embedJpg(bytes);
        const land = im.w > im.h, W = land ? 841.89 : 595.28, H = land ? 595.28 : 841.89, m = 18;
        const k = Math.min((W - 2 * m) / im.w, (H - 2 * m) / im.h);
        out.push({ kind: 'img', W, H, emb, w: im.w * k, h: im.h * k });
      } catch { toast(`« ${f.name} » : image illisible.`, true); }
    } else if (/\.(docx|xlsx|xls|ods|csv|txt|md)$/i.test(f.name)) {
      try {
        const bytes = await fileToPdfBytes(f);
        const src = await PDFLib.PDFDocument.load(bytes);
        const cp = await pdf.copyPages(src, src.getPageIndices());
        cp.forEach(p => out.push({ kind: 'ext', page: p }));
      } catch (e) { console.warn(e); toast(`« ${f.name} » n'a pas pu être converti en PDF.`, true); }
    } else toast(`« ${f.name} » : format non pris en charge.`, true);
  }
  return out;
}
function addBuiltPages(pdf, pages) {
  pages.forEach(k => {
    if (k.kind === 'ext') pdf.addPage(k.page);
    else { const pg = pdf.addPage([k.W, k.H]); pg.drawImage(k.emb, { x: (k.W - k.w) / 2, y: (k.H - k.h) / 2, width: k.w, height: k.h }); }
  });
}
$('#pdfIns').addEventListener('change', async e => {
  const files = [...e.target.files]; e.target.value = ''; if (!files.length) return;
  const at = PDFV.insertAt ?? PDFV.pages.length;
  toast('Ajout en cours…');
  const added = await appendSourcePages(at, async (pdf, PDFLib) => addBuiltPages(pdf, await pagesFromFiles(pdf, PDFLib, files)));
  if (added) toast(`${added} page${added > 1 ? 's ajoutées' : ' ajoutée'}`);
});

/* --- export : PDF avec annotations et formulaire intégrés */
function winAnsiSafe(font, s) {
  let out = '';
  for (const ch of s) { try { font.encodeText(ch); out += ch; } catch { out += ch === '\t' ? '    ' : '?'; } }
  return out;
}
function wrapLines(font, text, size, maxW) {
  const lines = [];
  text.split('\n').forEach(par => {
    const words = par.split(/(\s+)/); let line = '';
    words.forEach(w => {
      const t = line + w;
      if (line && font.widthOfTextAtSize(t.trimEnd(), size) > maxW) { lines.push(line.trimEnd()); line = w.trimStart(); }
      else line = t;
    });
    lines.push(line.trimEnd());
  });
  return lines;
}
const hexRgbPdf = (PDFLib, hex) => { const [r, g, b] = hexRgb(hex); return PDFLib.rgb(r / 255, g / 255, b / 255); };
async function pdfExportBytes(opts = {}) {
  const PDFLib = await ensurePdfLib();
  const cur = state.cur;
  const pdf = await PDFLib.PDFDocument.load(PDFV.bytes, { ignoreEncryption: true });
  // pages dans l'ordre affiché, avec leurs rotations
  const order = cur.order || [], identity = order.length === pdf.getPageCount() && order.every((o, k) => o.s === k && !o.r);
  if (!identity) {
    const orig = pdf.getPages(), seen = new Set(), list = [];
    for (const o of order) {
      if (!seen.has(o.s)) { seen.add(o.s); list.push(orig[o.s]); }
      else { const [cp] = await pdf.copyPages(pdf, [o.s]); list.push(cp); }
    }
    for (let i = orig.length - 1; i >= 0; i--) pdf.removePage(i);
    list.forEach((pg, j) => { pdf.addPage(pg); pg.setRotation(PDFLib.degrees(PDFV.pages[j].rot)); });
  }
  // formulaire
  const vals = cur.form || {};
  if (Object.keys(vals).length) {
    try {
      const form = pdf.getForm();
      Object.entries(vals).forEach(([name, v]) => {
        try {
          const f = form.getField(name);
          if (f instanceof PDFLib.PDFTextField) f.setText(String(v ?? ''));
          else if (f instanceof PDFLib.PDFCheckBox) v ? f.check() : f.uncheck();
          else if (f instanceof PDFLib.PDFRadioGroup) { if (v) f.select(String(v)); }
          else if (f instanceof PDFLib.PDFDropdown || f instanceof PDFLib.PDFOptionList) { if (v) f.select(String(v)); else f.clear && f.clear(); }
        } catch (e) { console.warn('Champ ignoré', name, e); }
      });
      if (opts.flattenForm) form.flatten();
    } catch (e) { console.warn('Formulaire ignoré', e); }
  }
  // annotations
  const ann = cur.ann || [];
  if (ann.length) {
    const font = await pdf.embedFont(PDFLib.StandardFonts.Helvetica);
    const pages = pdf.getPages(), imgCache = new Map();
    for (const a of ann) {
      const pg = pages[a.p], P = PDFV.pages[a.p]; if (!pg || !P) continue;
      const vp = vpOf(P, 1), rot = P.rot;
      const U = (x, y) => vp.convertToPdfPoint(x, y);
      const box = () => { const [x1, y1] = U(a.x, a.y), [x2, y2] = U(a.x + a.w, a.y + a.h); return { x: Math.min(x1, x2), y: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) }; };
      try {
        if (a.t === 'hl') pg.drawRectangle({ ...box(), color: hexRgbPdf(PDFLib, a.color), opacity: 0.38, blendMode: PDFLib.BlendMode.Multiply });
        else if (a.t === 'white') pg.drawRectangle({ ...box(), color: PDFLib.rgb(1, 1, 1) });
        else if (a.t === 'rect') pg.drawRectangle({ ...box(), borderColor: hexRgbPdf(PDFLib, a.color), borderWidth: a.lw || 2 });
        else if (a.t === 'ink') {
          const c = hexRgbPdf(PDFLib, a.color);
          for (let k = 1; k < a.pts.length; k++) { const [x1, y1] = U(...a.pts[k - 1]), [x2, y2] = U(...a.pts[k]); pg.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: a.lw || 2, color: c, lineCap: PDFLib.LineCapStyle.Round }); }
        } else if (a.t === 'text') {
          const lines = wrapLines(font, winAnsiSafe(font, a.text), a.size, a.w);
          const lh = a.size * 1.2;
          lines.forEach((ln, k) => {
            const [x, y] = U(a.x, a.y + a.size * 0.95 + k * lh + (lh - a.size * 1.15) / 2);
            pg.drawText(ln, { x, y, size: a.size, font, color: hexRgbPdf(PDFLib, a.color), rotate: PDFLib.degrees(rot) });
          });
        } else if (a.t === 'img') {
          let emb = imgCache.get(a.src);
          if (!emb) { const bytes = b64ToBytes(a.src.split(',')[1]); emb = /^data:image\/png/.test(a.src) ? await pdf.embedPng(bytes) : await pdf.embedJpg(bytes); imgCache.set(a.src, emb); }
          const [x, y] = U(a.x, a.y + a.h);
          pg.drawImage(emb, { x, y, width: a.w, height: a.h, rotate: PDFLib.degrees(rot) });
        }
      } catch (e) { console.warn('Annotation ignorée', e); }
    }
  }
  return pdf.save({ useObjectStreams: true });
}
async function exportPdfFile() {
  try { const bytes = await pdfExportBytes(); await saveFile(safeName(state.cur.title) + '.pdf', bytes); }
  catch (e) { console.warn(e); toast('Export impossible : ce PDF est peut-être protégé.', true); }
}
async function printPdf() {
  try {
    const bytes = await pdfExportBytes();
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
    if (isPhone()) { window.open(url, '_blank'); return; }
    let fr = $('#pdfPrintFrame'); if (fr) fr.remove();
    fr = document.createElement('iframe'); fr.id = 'pdfPrintFrame'; fr.style.cssText = 'position:fixed;width:1px;height:1px;opacity:0;pointer-events:none;bottom:0;right:0';
    fr.src = url; document.body.appendChild(fr);
    fr.onload = () => { try { fr.contentWindow.focus(); fr.contentWindow.print(); } catch { window.open(url, '_blank'); } };
  } catch { toast('Impression impossible.', true); }
}
// Diviser / extraire : parts = listes d'indices de pages
async function splitPdf(parts, dest, label) {
  const PDFLib = await ensurePdfLib();
  try {
    const full = await PDFLib.PDFDocument.load(await pdfExportBytes({ flattenForm: false }), { ignoreEncryption: true });
    const outs = [];
    for (const idx of parts) {
      const d = await PDFLib.PDFDocument.create();
      (await d.copyPages(full, idx)).forEach(p => d.addPage(p));
      outs.push({ bytes: await d.save({ useObjectStreams: true }), idx });
    }
    const base = state.cur.title || 'Document';
    const nameOf = o => `${base} — ${label === 'extrait' ? 'extrait' : 'partie'} ${o.idx.length === 1 ? 'p. ' + (o.idx[0] + 1) : 'p. ' + (o.idx[0] + 1) + '-' + (o.idx[o.idx.length - 1] + 1)}`;
    if (dest === 'zip') {
      await loadScript(LIB.jszip);
      const z = new window.JSZip(); outs.forEach(o => z.file(safeName(nameOf(o)) + '.pdf', o.bytes));
      await saveFile(safeName(base) + ' (divisé).zip', await z.generateAsync({ type: 'uint8array' }));
    } else if (dest === 'download' && outs.length === 1) await saveFile(safeName(nameOf(outs[0])) + '.pdf', outs[0].bytes);
    else {
      const now = Date.now();
      for (const [k, o] of outs.entries()) {
        const pdfId = uid(); await Store.putBlob(pdfId, o.bytes);
        const doc = { id: uid(), type: 'pdf', title: nameOf(o), pdfId, ann: [], form: {}, pagesN: o.idx.length, createdAt: now + k, updatedAt: now + k };
        await Store.save(doc);
      }
      toast(outs.length === 1 ? `« ${nameOf(outs[0])} » ajouté à vos documents` : `${outs.length} PDF ajoutés à vos documents`);
    }
  } catch (e) { console.warn(e); toast('Division impossible : ce PDF est peut-être protégé.', true); }
}
function parseRanges(s, n) {
  const out = [];
  for (const part of String(s).split(/[;,]/)) {
    const t = part.trim(); if (!t) continue;
    const m = /^(\d+)\s*(?:-|à)\s*(\d+)$/.exec(t), one = /^\d+$/.test(t) ? +t : null;
    if (m) { const a = +m[1], b = +m[2]; if (a < 1 || b > n || a > b) return null; out.push(Array.from({ length: b - a + 1 }, (_, k) => a - 1 + k)); }
    else if (one) { if (one < 1 || one > n) return null; out.push([one - 1]); }
    else return null;
  }
  return out.length ? out : null;
}
function openSplit() {
  const n = PDFV.pages.length, p = $('#pdfDlg');
  p.innerHTML = `<div class="dlg-in"><div class="fp-h"><b>Diviser le PDF</b><button class="xbtn" data-dx aria-label="Fermer">${ic('close')}</button></div>
    <label class="fp-it"><input type="radio" name="spm" value="each" checked><span>Une page par fichier (${n} fichiers)</span></label>
    <label class="fp-it"><input type="radio" name="spm" value="every"><span>Toutes les <input id="spN" type="number" min="1" max="${n}" value="${Math.min(2, n)}" class="sp-n"> pages</span></label>
    <label class="fp-it"><input type="radio" name="spm" value="ranges"><span>Selon des plages</span></label>
    <input id="spR" class="fp-search" placeholder="ex. 1-3, 4, 5-${n}" autocomplete="off">
    <span class="dlg-lbl">Résultat</span>
    <label class="fp-it"><input type="radio" name="spd" value="grimoire" checked><span>Ajouter les PDF à mes documents Grimoire</span></label>
    <label class="fp-it"><input type="radio" name="spd" value="zip"><span>Télécharger un fichier .zip</span></label>
    <div class="fp-acts"><button class="pbtn" id="spOk">Diviser</button></div></div>`;
  p.hidden = false;
  $('#spR').addEventListener('focus', () => { p.querySelector('[name="spm"][value="ranges"]').checked = true; });
  $('#spN').addEventListener('focus', () => { p.querySelector('[name="spm"][value="every"]').checked = true; });
  $('#spOk').addEventListener('click', () => {
    const m = p.querySelector('[name="spm"]:checked').value, dest = p.querySelector('[name="spd"]:checked').value;
    let parts;
    if (m === 'each') parts = PDFV.pages.map((_, i) => [i]);
    else if (m === 'every') { const k = Math.max(1, +$('#spN').value || 1); parts = []; for (let i = 0; i < n; i += k) parts.push(Array.from({ length: Math.min(k, n - i) }, (_, j) => i + j)); }
    else { parts = parseRanges($('#spR').value, n); if (!parts) return toast(`Plages incorrectes : utilisez des numéros de 1 à ${n}, par exemple 1-3, 4, 5-${n}.`, true); }
    p.hidden = true; toast('Division en cours…'); splitPdf(parts, dest, 'partie');
  });
}

/* --- recherche dans le PDF */
const PF = { q: '', hits: [], i: -1 };
async function ensurePdfTexts() {
  if (PDFV.texts) return PDFV.texts;
  const out = [];
  for (const P of PDFV.pages) {
    if (PDFV.srcText[P.s] == null) { try { const tc = await P.page.getTextContent(); PDFV.srcText[P.s] = norm(tc.items.map(it => it.str || '').join('')); } catch { PDFV.srcText[P.s] = ''; } }
    out.push(PDFV.srcText[P.s]);
  }
  return (PDFV.texts = out);
}
function openPdfFind() { if (PDFV.mode !== 'read') setPdfMode('read'); $('#pdfFind').hidden = false; const q = $('#pfQ'); q.focus(); q.select(); }
function closePdfFind() { $('#pdfFind').hidden = true; PF.q = ''; PF.hits = []; if (window.CSS && CSS.highlights) { CSS.highlights.delete('g-find'); CSS.highlights.delete('g-cur'); } $('#pfCount').textContent = ''; }
async function runPdfFind() {
  const q = norm($('#pfQ').value.trim()); PF.q = q; PF.hits = [];
  if (!q) { $('#pfCount').textContent = ''; if (CSS.highlights) { CSS.highlights.delete('g-find'); CSS.highlights.delete('g-cur'); } return; }
  const T = await ensurePdfTexts();
  T.forEach((t, p) => { let k = 0, at = t.indexOf(q); while (at >= 0) { PF.hits.push({ p, k: k++ }); at = t.indexOf(q, at + q.length); } });
  const here = pdfCurrentPage();
  PF.i = Math.max(0, PF.hits.findIndex(h => h.p >= here));
  showPdfHit();
}
function showPdfHit() {
  $('#pfCount').textContent = PF.hits.length ? `${PF.i + 1} sur ${PF.hits.length}` : 'Aucun résultat';
  const h = PF.hits[PF.i]; if (!h) return;
  const el = $(`#pdfPages .pp[data-p="${h.p}"]`); if (!el) return;
  const sc = $('#pdfScroll');
  if (el.offsetTop > sc.scrollTop + sc.clientHeight - 60 || el.offsetTop + el.offsetHeight < sc.scrollTop + 40) sc.scrollTop = el.offsetTop - 8;
  renderPdfPage(h.p).then(() => paintPdfFind(h.p));
  PDFV.pages.forEach((_, i) => { if (i !== h.p) paintPdfFind(i); });
}
function paintPdfFind(only) {
  if (!(window.CSS && CSS.highlights && window.Highlight) || !PF.q) return;
  const all = [], curR = [];
  $$('#pdfPages .pp').forEach(pp => {
    const i = +pp.dataset.p, tl = pp.querySelector('.textLayer');
    const nodes = []; const w = document.createTreeWalker(tl, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode())) nodes.push(n);
    let txt = '', map = [];
    nodes.forEach(nd => { for (let k = 0; k < nd.nodeValue.length; k++) { map.push([nd, k]); } txt += nd.nodeValue; });
    const t = norm(txt); let at = t.indexOf(PF.q), k = 0;
    while (at >= 0 && map[at] && map[at + PF.q.length - 1]) {
      const r = new Range(); r.setStart(map[at][0], map[at][1]); const e = map[at + PF.q.length - 1]; r.setEnd(e[0], e[1] + 1);
      const h = PF.hits[PF.i];
      if (h && h.p === i && h.k === k) curR.push(r); else all.push(r);
      at = t.indexOf(PF.q, at + PF.q.length); k++;
    }
  });
  CSS.highlights.set('g-find', new Highlight(...all)); CSS.highlights.set('g-cur', new Highlight(...curR));
  void only;
}
$('#pfQ').addEventListener('input', debounce(runPdfFind, 200));
$('#pfQ').addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); if (!PF.hits.length) return; PF.i = (PF.i + (e.shiftKey ? -1 : 1) + PF.hits.length) % PF.hits.length; showPdfHit(); }
  else if (e.key === 'Escape') { e.preventDefault(); closePdfFind(); }
});
$('#pdfFind').addEventListener('click', e => {
  const b = e.target.closest('[data-pf]'); if (!b) return;
  if (b.dataset.pf === 'close') return closePdfFind();
  if (!PF.hits.length) return;
  PF.i = (PF.i + (b.dataset.pf === 'prev' ? -1 : 1) + PF.hits.length) % PF.hits.length; showPdfHit();
});

/* --- conversion en document Grimoire */
async function pdfToDoc() {
  toast('Extraction du texte…');
  const parts = [];
  for (let i = 0; i < PDFV.pages.length; i++) {
    const t = await pageText(PDFV.src, PDFV.pages[i].s, true);
    const paras = t.split(/\n\s*\n|\n(?=[A-ZÀ-Ý0-9•\-–])/).map(s => s.replace(/\s*\n\s*/g, ' ').trim()).filter(Boolean);
    parts.push(paras.map(p => `<p>${esc(p)}</p>`).join(''));
  }
  const html = parts.join('<hr data-pb="1">');
  if (!html.replace(/<[^>]+>/g, '').trim()) return toast('Ce PDF ne contient pas de texte (document scanné ?).', true);
  await createDoc('texte', { title: (state.cur.title || 'PDF') + ' (texte)', html });
  toast('Texte du PDF copié dans un nouveau document.');
}

/* --- barre d'outils */
$('#pdfv').addEventListener('click', e => {
  const m = e.target.closest('[data-pm]'); if (m) { if (m.dataset.pm === 'form') { setPdfMode('read'); toast('Touchez les champs du formulaire pour les remplir.'); return; } setPdfMode(m.dataset.pm); return; }
  const t = e.target.closest('[data-pt]'); if (t) {
    const k = t.dataset.pt;
    if (k === 'sign') return openSignPad();
    if (k === 'image') return $('#pdfImg').click();
    setPdfTool(k); return;
  }
  const z = e.target.closest('[data-pz]'); if (z) { const k = z.dataset.pz; pdfZoom(k === 'in' ? 1.2 : k === 'out' ? 1 / 1.2 : 'fit'); return; }
  const a = e.target.closest('[data-pa]'); if (a) { ({ undo: () => pdfUndo(false), redo: () => pdfUndo(true), find: openPdfFind })[a.dataset.pa](); }
});
$('#pdfImg').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try { const im = await fileToImageData(f, 1600); addImageAnn(im.src, Math.min(220, im.w / 2), im.h / 2 * Math.min(1, 220 / (im.w / 2))); }
  catch { toast('Image illisible.', true); }
});
$('#pdfPg').addEventListener('click', () => {
  if (PDFV.mode === 'pages') return;
  const n = PDFV.pages.length, v = parseInt(window.prompt(`Aller à la page (1 à ${n}) :`, String(pdfCurrentPage() + 1)), 10);
  if (v >= 1 && v <= n) { const el = $(`#pdfPages .pp[data-p="${v - 1}"]`); if (el) $('#pdfScroll').scrollTop = el.offsetTop - 8; }
});

/* --- import : un PDF, ou plusieurs fichiers fusionnés (PDF et images) */
async function importPdfFiles(files) {
  files = [...files];
  const one = files.length === 1 && (/pdf$/i.test(files[0].type) || /\.pdf$/i.test(files[0].name));
  try {
    let bytes, title;
    if (one) { bytes = new Uint8Array(await files[0].arrayBuffer()); title = files[0].name.replace(/\.[^.]+$/, ''); }
    else {
      const PDFLib = await ensurePdfLib();
      const pdf = await PDFLib.PDFDocument.create();
      const pages = await pagesFromFiles(pdf, PDFLib, files);
      if (!pages.length) return toast('Aucun fichier utilisable (PDF, image, Word, Excel, CSV ou texte).', true);
      addBuiltPages(pdf, pages);
      bytes = await pdf.save({ useObjectStreams: true });
      title = files.length > 1 ? `Fusion de ${files.length} fichiers` : files[0].name.replace(/\.[^.]+$/, '');
      if (files.length === 1) toast('Conversion en PDF…');
    }
    // vérification rapide (PDF lisible) : seule la structure est lue, pas les pages
    const doc = await pdfjsOpen(bytes);
    const n = doc.numPages; doc.destroy();
    const pdfId = uid(); await Store.putBlob(pdfId, bytes);
    await createDoc('pdf', { title, pdfId, ann: [], form: {}, pagesN: n });
    toast(one || files.length === 1 ? `PDF ouvert : ${n} page${n > 1 ? 's' : ''}` : `PDF créé : ${n} page${n > 1 ? 's' : ''}. Réorganisez-les dans l'onglet Pages.`);
  } catch (e) {
    console.warn(e);
    toast(e && e.name === 'PasswordException' ? 'PDF protégé : mot de passe nécessaire.' : 'Ce fichier PDF ne peut pas être ouvert.', true);
  }
}
$('#pdfCard').addEventListener('click', () => $('#pdfIn').click());
$('#pdfIn').addEventListener('change', e => { const fs = [...e.target.files]; e.target.value = ''; if (fs.length) importPdfFiles(fs); });

/* =========================================================
   Conversion en PDF : documents (Word, Grimoire), classeurs (Excel, CSV, Grimoire), texte
   ========================================================= */
async function ensurePdfMake() { await loadScript(LIB.pdfmake); await loadScript(LIB.pdfmakeFonts); return window.pdfMake; }
const PT_PX = 0.75; // 1 px (écran) = 0,75 pt
const pdfTxt = t => String(t).replace(/[\u202f\u2009\u2007]/g, '\u00a0').replace(/[\u200b\ufeff]/g, '');
const MARGIN_PT = { normal: 70.9, narrow: 36, wide: 90.7 };
function pdfMakeBytes(dd) {
  return new Promise((res, rej) => { try { window.pdfMake.createPdf(dd).getBuffer(b => res(new Uint8Array(b))); } catch (e) { rej(e); } });
}
// Pied et en-tête de page, numéros « Page X sur Y »
function pdfHeadFoot(P, dd, mar) {
  const small = { fontSize: 8.5, color: '#666666' };
  if (P.header) dd.header = () => ({ text: P.header, alignment: 'center', margin: [mar, 22, mar, 0], ...small });
  if (P.footer || (P.num && P.num !== 'none')) dd.footer = (cur, n) => ({
    columns: [
      { text: P.footer || '', alignment: 'left', ...small },
      { text: P.num === 'xofy' ? `Page ${cur} sur ${n}` : P.num && P.num !== 'none' ? String(cur) : '', alignment: P.num === 'right' ? 'right' : 'center', ...small }
    ], margin: [mar, 10, mar, 0]
  });
}
// --- HTML (le sous-ensemble produit par Grimoire) → contenu pdfmake
function htmlToPdfContent(html, contentW, sty = {}) {
  const root = document.createElement('div'); root.innerHTML = sanitize(html || '');
  root.querySelectorAll('del[data-a]').forEach(el => el.remove()); root.querySelectorAll('ins[data-a]').forEach(el => el.replaceWith(...el.childNodes));
  let hasToc = false;
  const runs = (node, st = {}) => {
    const out = [];
    node.childNodes.forEach(n => {
      if (n.nodeType === 3) { const t = pdfTxt(n.nodeValue).replace(/[ \t\n\r]+/g, ' '); if (t) out.push({ text: t, ...st }); return; }
      if (n.nodeType !== 1) return;
      const tg = n.tagName, s2 = { ...st };
      if (tg === 'BR') { out.push({ text: '\n', ...st }); return; }
      if (tg === 'IMG' || /^(UL|OL|TABLE)$/.test(tg)) return;
      if (tg === 'STRONG' || tg === 'B') s2.bold = true;
      if (tg === 'EM' || tg === 'I') s2.italics = true;
      if (tg === 'U') s2.decoration = 'underline';
      if (tg === 'S') s2.decoration = 'lineThrough';
      if (tg === 'SUP') s2.sup = true;
      if (tg === 'SUB') s2.sub = true;
      if (tg === 'A') { s2.link = n.getAttribute('href'); s2.color = '#1F5FAD'; s2.decoration = 'underline'; }
      if (tg === 'SPAN') {
        const c = cssColor(n.style.color), bg = cssColor(n.style.backgroundColor), fs = cssSizePt(n.style.fontSize);
        if (c) s2.color = c; if (bg) s2.background = bg; if (fs) s2.fontSize = fs;
      }
      out.push(...runs(n, s2));
    });
    // espaces de début et de fin de paragraphe
    if (out.length && typeof out[0].text === 'string') out[0].text = out[0].text.replace(/^\s+/, '');
    if (out.length && typeof out[out.length - 1].text === 'string') out[out.length - 1].text = out[out.length - 1].text.replace(/\s+$/, '');
    return out.filter(r => r.text !== '');
  };
  const align = el => ({ center: 'center', right: 'right', justify: 'justify' }[el.style && el.style.textAlign] || undefined);
  const imgs = el => [...el.querySelectorAll(':scope > img, :scope > span > img')].map(im => ({ image: im.getAttribute('src'), width: Math.max(20, contentW * (parseFloat(im.style.width) || 100) / 100), alignment: im.getAttribute('data-wrap') === 'right' ? 'right' : im.getAttribute('data-wrap') === 'left' ? 'left' : 'center', margin: [0, 4, 0, 8] }));
  const para = (el, extra = {}) => {
    const out = [...imgs(el)];
    const r = runs(el);
    if (r.length) {
      const p = { text: r, margin: [0, 0, 0, el.style.marginBottom === '0px' ? 0 : 7], ...extra };
      const a = align(el); if (a) p.alignment = a;
      const lh = parseFloat(el.style.lineHeight); if (lh) p.lineHeight = Math.max(1, lh * 0.82);
      const ml = parseInt(el.style.marginLeft, 10); if (ml) p.margin = [ml * PT_PX, p.margin[1], 0, p.margin[3]];
      out.push(p);
    } else if (!out.length) out.push({ text: ' ', margin: [0, 0, 0, 4] });
    return out;
  };
  const checkBox = done => ({ width: 13, canvas: [{ type: 'rect', x: 0, y: 2.5, w: 8.5, h: 8.5, r: 1.5, lineWidth: 0.9, lineColor: '#555555' }, ...(done ? [{ type: 'polyline', lineWidth: 1.3, lineColor: '#A35E2C', points: [{ x: 1.8, y: 7 }, { x: 3.8, y: 9.2 }, { x: 7.5, y: 4 }] }] : [])] });
  const list = el => {
    const items = [...el.children].filter(li => li.tagName === 'LI').map(li => {
      const own = document.createElement('div'); [...li.childNodes].forEach(c => { if (!(c.nodeType === 1 && /^(UL|OL)$/.test(c.tagName))) own.appendChild(c.cloneNode(true)); });
      const sub = [...li.children].filter(c => /^(UL|OL)$/.test(c.tagName)).map(list);
      const txt = { text: runs(own), margin: [0, 0, 0, 3] };
      if (el.hasAttribute('data-check')) return { columns: [checkBox(li.getAttribute('data-done') === '1'), { ...txt, width: '*', decoration: li.getAttribute('data-done') === '1' ? 'lineThrough' : undefined, color: li.getAttribute('data-done') === '1' ? '#888888' : undefined }], columnGap: 4 };
      return sub.length ? { stack: [txt, ...sub] } : txt;
    });
    if (el.hasAttribute('data-check')) return { stack: items, margin: [0, 0, 0, 7] };
    return el.tagName === 'OL' ? { ol: items, margin: [0, 0, 0, 7] } : { ul: items, margin: [0, 0, 0, 7] };
  };
  const table = tb => {
    const rows = [...tb.rows]; if (!rows.length) return null;
    const n = gridCols(tb);
    const cols = [...tb.querySelectorAll('col')];
    const widths = cols.length === n ? cols.map(c => (parseFloat(c.style.width) || 100 / n) + '%').map(w => (parseFloat(w) / 100) * contentW) : Array(n).fill('*');
    let headerRows = 0;
    const body = rows.map(tr => {
      if (tr.parentElement.tagName === 'THEAD') headerRows++;
      const out = [];
      rowGrid(tr).forEach(o => {
        while (out.length < o.start) out.push({ text: '' });
        const cell = o.cell, th = cell.tagName === 'TH';
        const inner = [...cell.children].some(c => /^(P|UL|OL|H1|H2|H3|TABLE)$/.test(c.tagName)) ? { stack: blocks(cell) } : { text: runs(cell, th ? { bold: true } : {}) };
        const bg = cssColor(cell.style.backgroundColor) || (th ? '#EAD9B8' : undefined);
        const span = Math.min(o.span, n - o.start);
        out.push({ ...inner, fontSize: 9.5, fillColor: bg, colSpan: span > 1 ? span : undefined, alignment: align(cell) });
        for (let k = 1; k < span; k++) out.push({});
      });
      while (out.length < n) out.push({ text: '' });
      return out.slice(0, n);
    });
    return { table: { headerRows, widths, body, dontBreakRows: true }, layout: { hLineWidth: () => 0.6, vLineWidth: () => 0.6, hLineColor: () => '#C5B8A8', vLineColor: () => '#C5B8A8', paddingLeft: () => 5, paddingRight: () => 5, paddingTop: () => 3, paddingBottom: () => 3 }, margin: [0, 2, 0, 10] };
  };
  const blocks = container => {
    const out = [];
    let inline = null;
    const flush = () => { if (inline) { out.push(...para(inline)); inline = null; } };
    [...container.childNodes].forEach(n => {
      if (n.nodeType === 3) { if (n.nodeValue.trim()) { inline ||= document.createElement('p'); inline.appendChild(n.cloneNode()); } return; }
      if (n.nodeType !== 1) return;
      const tg = n.tagName;
      if (/^(STRONG|EM|U|S|SUP|SUB|A|B|I|SPAN|BR|IMG)$/.test(tg)) { inline ||= document.createElement('p'); inline.appendChild(n.cloneNode(true)); return; }
      flush();
      if (/^H[1-3]$/.test(tg)) {
        const lv = +tg[1], sz = [0, 20, 15, 12.5][lv], o = sty['h' + lv] || {};
        out.push({ text: runs(n), fontSize: o.fs ? o.fs * 0.9 : sz, bold: o.b ?? true, italics: o.i || undefined, color: o.color || (lv === 3 ? '#A35E2C' : '#1B1512'), margin: [0, o.sb ?? (lv === 1 ? 6 : 12), 0, o.sa ?? (lv === 1 ? 8 : 5)], alignment: align(n) || o.al, tocItem: true, tocStyle: lv === 1 ? { bold: true } : {}, tocMargin: [(lv - 1) * 12, 0, 0, 2], headlineLevel: lv });
      } else if (tg === 'P') out.push(...para(n));
      else if (tg === 'UL' || tg === 'OL') out.push(list(n));
      else if (tg === 'TABLE') { const t = table(n); if (t) out.push(t); }
      else if (tg === 'BLOCKQUOTE') {
        const inner = n.querySelector('p') ? blocks(n) : para(n);
        const q = sty.blockquote || {};
        out.push({ table: { widths: ['*'], body: [[{ stack: inner, italics: q.i ?? true, color: q.color || '#6B5F56', ...(q.fs ? { fontSize: q.fs * 0.9 } : {}) }]] }, layout: { hLineWidth: () => 0, vLineWidth: i => i === 0 ? 2.5 : 0, vLineColor: () => '#C9804F', paddingLeft: () => 10 }, margin: [0, 2, 0, 8] });
      } else if (tg === 'HR' && n.hasAttribute('data-pb')) out.push({ text: '', pageBreak: 'after' });
      else if (tg === 'HR') out.push({ canvas: [{ type: 'line', x1: 0, y1: 4, x2: contentW, y2: 4, lineWidth: 0.8, lineColor: '#C9804F' }], margin: [0, 4, 0, 10] });
      else if (tg === 'DIV' && n.hasAttribute('data-cols')) {
        // colonnes : le texte est réparti en colonnes de longueur voisine
        const k = +n.getAttribute('data-cols') || 2, kids = [...n.children], tot = kids.reduce((t, c) => t + c.textContent.length + 40, 0);
        const groups = Array.from({ length: k }, () => document.createElement('div')); let acc = 0;
        kids.forEach(c => { const g = Math.min(k - 1, Math.floor(acc / (tot / k))); groups[g].appendChild(c.cloneNode(true)); acc += c.textContent.length + 40; });
        out.push({ columns: groups.map(g => ({ width: '*', stack: blocks(g) })), columnGap: 16, margin: [0, 0, 0, 6] });
      }
      else if (tg === 'DIV' && n.hasAttribute('data-toc')) { hasToc = true; out.push({ toc: { title: { text: 'Sommaire', fontSize: 14, bold: true, margin: [0, 0, 0, 8] }, numberStyle: { color: '#6B5F56' } }, margin: [0, 0, 0, 12] }); }
      else out.push(...para(n));
    });
    flush();
    return out;
  };
  const content = blocks(root);
  if (!hasToc) content.forEach(c => { delete c.tocItem; });
  return content;
}
async function htmlToPdfBytes(html, page = {}, title = '', sty = null) {
  await ensurePdfMake();
  const P = { ...PAGE_DEF, ...page }, land = P.orient === 'landscape', mar = MARGIN_PT[P.margin] || 70.9;
  const contentW = (land ? 841.89 : 595.28) - 2 * mar;
  const dd = {
    pageSize: 'A4', pageOrientation: land ? 'landscape' : 'portrait', pageMargins: [mar, mar, mar, mar],
    info: { title, creator: 'Grimoire' },
    defaultStyle: Object.assign({ font: 'Roboto', fontSize: 11, lineHeight: 1.18, color: '#1E1814' }, sty && sty.p ? { ...(sty.p.fs ? { fontSize: sty.p.fs * 0.9 } : {}), ...(sty.p.color ? { color: sty.p.color } : {}), ...(sty.p.i ? { italics: true } : {}), ...(sty.p.b ? { bold: true } : {}) } : {}),
    content: htmlToPdfContent(html, contentW, sty || {})
  };
  pdfHeadFoot(P, dd, mar);
  return pdfMakeBytes(dd);
}
// --- grille de cellules (classeur) → tableau pdfmake ; grid = { name, rows: [[{ t, b, i, u, a, bg, fc, num, cs, rs, skip }]], colW: [px], charts: [dataURL] }
function gridToPdfContent(grid, maxW) {
  const n = Math.max(1, ...grid.rows.map(r => r.length));
  let widths = Array.from({ length: n }, (_, c) => (grid.colW[c] || 104) * PT_PX);
  const tot = widths.reduce((s, w) => s + w, 0);
  const room = maxW - n * 6.5; // marges intérieures (3 + 3) et traits de chaque colonne
  const k = tot > room ? room / tot : 1;
  widths = widths.map(w => Math.max(8, w * k));
  const fs = Math.max(6, Math.min(9.5, 9.5 * Math.max(k, 0.65)));
  const body = grid.rows.map(row => {
    const out = [];
    for (let c = 0; c < n; c++) {
      const x = row[c];
      if (!x) { out.push({ text: '' }); continue; }
      if (x.skip) { out.push({}); continue; }
      if (x.chk != null) { out.push({ canvas: [{ type: 'rect', x: 0, y: 1, w: fs * 0.9, h: fs * 0.9, r: 1.2, lineWidth: 0.8, lineColor: '#555555' }, ...(x.chk ? [{ type: 'polyline', lineWidth: 1.1, lineColor: '#A35E2C', points: [{ x: fs * 0.18, y: fs * 0.5 }, { x: fs * 0.38, y: fs * 0.75 }, { x: fs * 0.75, y: fs * 0.22 }] }] : [])], fillColor: x.bg, alignment: 'center' }); continue; }
      const cell = { text: pdfTxt(x.t ?? ''), fontSize: fs, bold: !!x.b, italics: !!x.i, decoration: x.u ? 'underline' : undefined, alignment: x.a || (x.num ? 'right' : 'left'), fillColor: x.bg, color: x.fc, noWrap: !x.w && !x.cs };
      if (x.cs > 1) cell.colSpan = x.cs; if (x.rs > 1) cell.rowSpan = x.rs;
      out.push(cell);
    }
    return out;
  });
  const out = [{ text: grid.name, fontSize: 13, bold: true, margin: [0, 0, 0, 8] }];
  if (body.length) out.push({ table: { widths, body, dontBreakRows: true }, layout: { hLineWidth: () => 0.5, vLineWidth: () => 0.5, hLineColor: () => '#D9D2C8', vLineColor: () => '#D9D2C8', paddingLeft: () => 3, paddingRight: () => 3, paddingTop: () => 2, paddingBottom: () => 2 } });
  else out.push({ text: '(feuille vide)', color: '#888888' });
  (grid.charts || []).forEach(ch => out.push({ image: ch.src, width: Math.min(maxW, ch.w * PT_PX), margin: [0, 12, 0, 0] }));
  return { content: out, wide: tot * Math.min(1, 1) > 595.28 - 2 * 36 };
}
async function gridsToPdfBytes(grids, title = '') {
  await ensurePdfMake();
  const content = [];
  // paysage si un tableau est trop large pour une page en portrait
  const anyWide = grids.some(g => g.colW.slice(0, Math.max(1, ...g.rows.map(r => r.length))).reduce((s, w) => s + (w || 104) * PT_PX, 0) > 595.28 - 72);
  const W = (anyWide ? 841.89 : 595.28) - 72;
  grids.forEach((g, i) => { const r = gridToPdfContent(g, W); if (i) r.content[0].pageBreak = 'before'; content.push(...r.content); });
  return pdfMakeBytes({ pageSize: 'A4', pageOrientation: anyWide ? 'landscape' : 'portrait', pageMargins: [36, 40, 36, 40], info: { title, creator: 'Grimoire' }, defaultStyle: { font: 'Roboto', fontSize: 9.5, color: '#1E1814' }, content, footer: (cur, n) => ({ text: `${cur} / ${n}`, alignment: 'right', fontSize: 8, color: '#888888', margin: [36, 10, 36, 0] }) });
}
// Classeur Grimoire → grilles (valeurs calculées, mise en forme, fusions, graphiques)
async function sheetGrids(cur) {
  const grids = [];
  const hexBg = k => k && XL_RGB[k] ? '#' + XL_RGB[k][1] : undefined, hexFc = k => k && XL_RGB[k] ? '#' + XL_RGB[k][0] : undefined;
  for (let i = 0; i < cur.sheets.length; i++) {
    const sh = cur.sheets[i];
    const vals = SH.hf && SH.ids ? SH.hf.getSheetValues(SH.ids[i]) : [];
    let maxR = -1, maxC = -1;
    Object.entries(sh.cells).forEach(([k, x]) => { const p = parseAddr(k); if (!p) return; const v = (vals[p.r] || [])[p.c]; if ((x.v !== '' && x.v != null) || (v != null && v !== '') || (x.f && (x.f.bg || x.f.mg))) { maxR = Math.max(maxR, p.r + ((x.f && x.f.mg && x.f.mg.r) || 1) - 1); maxC = Math.max(maxC, p.c + ((x.f && x.f.mg && x.f.mg.c) || 1) - 1); } });
    (vals || []).forEach((row, r) => (row || []).forEach((v, c) => { if (v != null && v !== '') { maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); } }));
    const hidR = new Set((sh.hid && sh.hid.r) || []), hidC = new Set((sh.hid && sh.hid.c) || []);
    const colIdx = []; for (let c = 0; c <= maxC; c++) if (!hidC.has(c)) colIdx.push(c);
    const covered = new Set();
    const rows = [];
    for (let r = 0; r <= maxR; r++) {
      if (hidR.has(r)) continue;
      const row = [];
      colIdx.forEach(c => {
        if (covered.has(r + ',' + c)) { row.push({ skip: true }); return; }
        const x = sh.cells[addr(r, c)], f = (x && x.f) || {}, v = (vals[r] || [])[c];
        let t = '', num = false;
        if (f.chk || v == null || v === '') t = '';
        else if (typeof v === 'object') t = ERR_FR[v.value] || v.value || '#ERREUR';
        else if (typeof v === 'number') { t = f.n === 'date' ? fmtDate(v) : numFmt(f.n, f.d).format(v).replace(/ | /g, ' '); num = true; }
        else if (typeof v === 'boolean') t = v ? 'VRAI' : 'FAUX';
        else t = String(v);
        const cell = { t, num, chk: f.chk ? (v === true || v === 1) : undefined, b: f.b, i: f.i, u: f.u, w: f.w, a: { al: 'left', ac: 'center', ar: 'right' }[f.a], bg: hexBg(f.bg), fc: hexFc(f.fc) };
        if (f.mg) { cell.cs = f.mg.c || 1; cell.rs = f.mg.r || 1; for (let rr = r; rr < r + cell.rs; rr++) for (let cc = c; cc < c + cell.cs; cc++) if (rr !== r || cc !== c) covered.add(rr + ',' + cc); }
        row.push(cell);
      });
      rows.push(row);
    }
    const charts = [];
    for (const ch of (sh.charts || [])) { const src = await chartPng(ch, vals, (r, c) => sh.cells[addr(r, c)]); if (src) charts.push({ src, w: ch.w }); }
    grids.push({ name: sh.name, rows, colW: colIdx.map(c => sh.colW[c] || 104), charts });
  }
  return grids;
}
// Graphique → image PNG (couleurs écrites en dur, fond blanc)
async function chartPng(ch, vals, cellsOf) {
  try {
    const W = ch.w * 2, H = ch.h * 2;
    const box = document.createElement('div');
    box.className = 'chart static'; box.style.cssText = `width:${ch.w}px;height:${ch.h}px;position:fixed;left:-9999px;top:0;background:#fff;` + ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948'].map((c, k) => `--s${k + 1}:${c}`).join(';') + ';--ch-bg:#fff';
    box.innerHTML = chartHTML(ch, vals, cellsOf);
    document.body.appendChild(box);
    const svg = box.querySelector('svg');
    const c = document.createElement('canvas'); c.width = W; c.height = H;
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
    // titre et légende (HTML) dessinés à la main
    g.scale(2, 2);
    const B = box.getBoundingClientRect(), rel = el => { const r = el.getBoundingClientRect(); return { x: r.left - B.left, y: r.top - B.top, w: r.width, h: r.height }; };
    const tt = box.querySelector('.ch-title');
    if (tt) { const r = rel(tt); g.font = '600 14px Arial, sans-serif'; g.fillStyle = '#1E1814'; g.textBaseline = 'middle'; g.fillText(ch.title, r.x, r.y + r.h / 2); }
    box.querySelectorAll('.ch-leg span').forEach(sp => {
      const i = sp.querySelector('i'), ri = rel(i), r = rel(sp);
      g.fillStyle = getComputedStyle(i).backgroundColor; g.fillRect(ri.x, ri.y, ri.w, ri.h);
      g.font = '12px Arial, sans-serif'; g.fillStyle = '#6B5F56'; g.textBaseline = 'middle'; g.fillText(sp.textContent, ri.x + ri.w + 5, r.y + r.h / 2);
    });
    if (svg) {
      // couleurs et polices calculées écrites en attributs (le SVG isolé ne voit plus les variables CSS)
      const cl = svg.cloneNode(true), src = [svg, ...svg.querySelectorAll('*')], dst = [cl, ...cl.querySelectorAll('*')];
      src.forEach((el, k) => {
        const cs = getComputedStyle(el), d = dst[k];
        if (el.tagName === 'text') { d.setAttribute('fill', /ch-tot/.test(el.getAttribute('class')) ? '#1E1814' : '#6B5F56'); d.setAttribute('font-family', 'Arial, sans-serif'); d.setAttribute('font-size', cs.fontSize); d.setAttribute('font-weight', cs.fontWeight); return; }
        if (cs.fill && cs.fill !== 'none') d.setAttribute('fill', cs.fill); else if (cs.fill === 'none') d.setAttribute('fill', 'none');
        if (cs.stroke && cs.stroke !== 'none') { d.setAttribute('stroke', /ch-grid/.test(el.getAttribute('class')) ? '#E5E0D8' : /ch-base/.test(el.getAttribute('class')) ? '#C5B8A8' : cs.stroke); d.setAttribute('stroke-width', cs.strokeWidth); }
        d.removeAttribute('class'); d.removeAttribute('style');
      });
      cl.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      const bb = svg.getBoundingClientRect();
      cl.setAttribute('width', bb.width); cl.setAttribute('height', bb.height);
      const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(cl)], { type: 'image/svg+xml' }));
      const im = new Image(); await new Promise((res, rej) => { im.onload = res; im.onerror = rej; im.src = url; });
      g.drawImage(im, bb.left - B.left, bb.top - B.top, bb.width, bb.height);
      URL.revokeObjectURL(url);
    }
    box.remove();
    return c.toDataURL('image/png');
  } catch { return null; }
}
// Fichier Excel / ODS / CSV (sans l'ouvrir dans Grimoire) → grilles, avec les valeurs affichées par Excel
async function sheetFileGrids(f) {
  await loadScript(LIB.xlsx);
  const X = window.XLSX, ext = (f.name.split('.').pop() || '').toLowerCase();
  const wb = ext === 'csv' ? X.read(await f.text(), { type: 'string', FS: guessSep(await f.text()) }) : X.read(await f.arrayBuffer(), { cellStyles: true, cellNF: true });
  let styles = {}, chartsOf = {};
  if (ext === 'xlsx') { try { const ex = await xlsxReadExtras(await f.arrayBuffer(), wb.SheetNames); Object.entries(ex).forEach(([k, v]) => { styles[k] = v.fmt || {}; chartsOf[k] = v.charts || []; }); } catch {} }
  const grids = [];
  for (const name of wb.SheetNames.slice(0, 20)) {
    const ws = wb.Sheets[name], ref = X.utils.decode_range(ws['!ref'] || 'A1');
    const maxR = Math.min(ref.e.r, 1999), maxC = Math.min(ref.e.c, 51);
    const hidC = new Set(), hidR = new Set();
    (ws['!cols'] || []).forEach((cw, i) => { if (cw && cw.hidden) hidC.add(i); });
    (ws['!rows'] || []).forEach((rw, i) => { if (rw && rw.hidden) hidR.add(i); });
    const covered = new Set(), mg = {};
    (ws['!merges'] || []).forEach(m => { mg[m.s.r + ',' + m.s.c] = m; for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) covered.add(r + ',' + c); });
    const colIdx = []; for (let c = 0; c <= maxC; c++) if (!hidC.has(c)) colIdx.push(c);
    const st = styles[name] || {};
    const rows = [];
    for (let r = 0; r <= maxR; r++) {
      if (hidR.has(r)) continue;
      rows.push(colIdx.map(c => {
        if (covered.has(r + ',' + c)) return { skip: true };
        const a = X.utils.encode_cell({ r, c }), cl = ws[a], s = st[a] || {};
        // texte tel qu'affiché, remis à la française (1,234.5 → 1 234,5 ; TRUE → VRAI)
        const t = !cl || cl.v == null ? '' : cl.t === 'b' ? (cl.v ? 'VRAI' : 'FAUX') : cl.t === 'e' ? (ERR_FR[cl.w] || cl.w || '#ERREUR') : cl.w == null ? String(cl.v) : cl.t === 'n' && !/[a-z]{3}/i.test(cl.w) ? String(cl.w).replace(/[,.]/g, m => m === ',' ? '\u00a0' : ',') : String(cl.w);
        const cell = { t, num: !!cl && cl.t === 'n', b: s.b, i: s.i, u: s.u, w: s.w, a: { al: 'left', ac: 'center', ar: 'right' }[s.a], bg: s.bg && XL_RGB[s.bg] ? '#' + XL_RGB[s.bg][1] : undefined, fc: s.fc && XL_RGB[s.fc] ? '#' + XL_RGB[s.fc][0] : undefined };
        const m = mg[r + ',' + c]; if (m) { cell.cs = m.e.c - m.s.c + 1; cell.rs = m.e.r - m.s.r + 1; }
        return cell;
      }));
    }
    // dernières lignes vides retirées
    while (rows.length && rows[rows.length - 1].every(x => !x || x.skip || !x.t)) rows.pop();
    const colW = colIdx.map(c => { const cw = (ws['!cols'] || [])[c]; return cw && cw.wpx ? cw.wpx : cw && cw.wch ? cw.wch * 7.5 : 90; });
    const charts = [];
    if ((chartsOf[name] || []).length) {
      const vals = [];
      for (let r = 0; r <= maxR; r++) { vals[r] = []; for (let c = 0; c <= maxC; c++) { const cl = ws[X.utils.encode_cell({ r, c })]; vals[r][c] = cl ? cl.v : null; } }
      for (const ch of chartsOf[name]) { const src = await chartPng(ch, vals, (r, c) => { const cl = ws[X.utils.encode_cell({ r, c })]; return cl ? { v: cl.v } : null; }); if (src) charts.push({ src, w: ch.w }); }
    }
    grids.push({ name, rows, colW, charts });
  }
  return grids;
}
// Fichier quelconque → octets PDF (null si format non pris en charge)
async function fileToPdfBytes(f) {
  const ext = (f.name.split('.').pop() || '').toLowerCase(), title = f.name.replace(/\.[^.]+$/, '');
  if (ext === 'docx') {
    droppedImages = 0;
    let raw, page;
    try { ({ html: raw, page } = await docxToHtml(await f.arrayBuffer())); }
    catch { await loadScript(LIB.mammoth); raw = (await window.mammoth.convertToHtml({ arrayBuffer: await f.arrayBuffer() })).value; }
    const box = document.createElement('div'); box.innerHTML = restoreChecklists(sanitize(raw));
    return htmlToPdfBytes(attachFloats(box).innerHTML, page || {}, title);
  }
  if (['xlsx', 'xls', 'ods', 'csv'].includes(ext)) return gridsToPdfBytes(await sheetFileGrids(f), title);
  if (ext === 'txt' || ext === 'md') {
    const text = await f.text();
    const html = text.split(/\n{2,}/).map(p => { const t = p.trim(); const h = /^(#{1,3})\s+(.*)$/.exec(t); return h ? `<h${h[1].length}>${esc(h[2])}</h${h[1].length}>` : `<p>${esc(t).replace(/\n/g, '<br>')}</p>`; }).join('');
    return htmlToPdfBytes(html, {}, title);
  }
  return null;
}
// Export PDF direct des documents et classeurs Grimoire
async function exportDocPdf() {
  try {
    toast('Création du PDF…');
    const cur = state.cur;
    const bytes = cur.type === 'feuille' ? (commitEdit(false), await gridsToPdfBytes(await sheetGrids(cur), cur.title)) : await htmlToPdfBytes(cur.html, pageOf(), cur.title, cur.styles);
    await saveFile(safeName(cur.title) + '.pdf', bytes);
  } catch (e) { console.warn(e); toast('Création du PDF impossible.', true); }
}

/* =========================================================
   Import
   ========================================================= */
$('#fileIn').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) importFile(f); });
// ---------- Import Word fidèle (lecture directe du .docx) ----------
const WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const FONT_ALIAS = { 'calibri light': 'Calibri', 'aptos': 'Calibri', 'aptos display': 'Calibri', 'aptos narrow': 'Calibri', 'carlito': 'Calibri', 'candara': 'Calibri', 'corbel': 'Calibri', 'helvetica': 'Arial', 'helvetica neue': 'Arial', 'liberation sans': 'Arial', 'arimo': 'Arial', 'segoe ui': 'Arial', 'roboto': 'Arial', 'open sans': 'Arial', 'tahoma': 'Verdana', 'trebuchet ms': 'Verdana', 'century gothic': 'Verdana', 'cambria': 'Georgia', 'garamond': 'Georgia', 'book antiqua': 'Georgia', 'palatino linotype': 'Georgia', 'constantia': 'Georgia', 'times': 'Times New Roman', 'liberation serif': 'Times New Roman', 'tinos': 'Times New Roman', 'consolas': 'Courier New', 'courier': 'Courier New', 'liberation mono': 'Courier New', 'cousine': 'Courier New', 'lucida console': 'Courier New' };
const HL_HEX = { yellow: '#FFF176', green: '#B9F6CA', cyan: '#B2EBF2', magenta: '#F8BBD0', blue: '#90CAF9', red: '#FF8A80', darkBlue: '#9FA8DA', darkCyan: '#80CBC4', darkGreen: '#A5D6A7', darkMagenta: '#CE93D8', darkRed: '#EF9A9A', darkYellow: '#E6EE9C', darkGray: '#BDBDBD', lightGray: '#E6E6E6' };
const mapFont = n => { if (!n) return ''; const k = n.trim().toLowerCase(); return FONT_OK.find(f => f.toLowerCase() === k) || FONT_ALIAS[k] || ''; };
const wKids = (el, name) => el ? [...el.children].filter(c => c.localName === name) : [];
const wKid = (el, name) => el ? [...el.children].find(c => c.localName === name) || null : null;
const wAt = (el, name) => el ? (el.getAttributeNS(WNS, name) ?? el.getAttribute('w:' + name) ?? el.getAttribute(name)) : null;
const wOn = el => { if (!el) return undefined; const v = wAt(el, 'val'); return !(v === '0' || v === 'false' || v === 'off' || v === 'none'); };
const rAt = (el, name) => el ? (el.getAttribute('r:' + name) || el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', name)) : null;
const desc = (el, name) => el ? [...el.getElementsByTagName('*')].filter(c => c.localName === name) : [];

// une image flottante seule dans son paragraphe est rattachée au paragraphe suivant, que le texte l'entoure
function attachFloats(root) {
  root.querySelectorAll('img[data-wrap]').forEach(im => {
    const p = im.parentElement;
    if (p && p !== root && p.tagName === 'P' && !p.textContent.trim() && p.querySelectorAll('img').length === 1) {
      const nx = p.nextElementSibling;
      if (nx && /^(P|H1|H2|H3)$/.test(nx.tagName) && nx.textContent.trim()) { nx.prepend(im); p.remove(); }
    }
  });
  return root;
}
async function docxToHtml(buf) {
  await loadScript(LIB.jszip);
  const zip = await window.JSZip.loadAsync(buf);
  const read = async p => { const f = zip.file(p); return f ? f.async('string') : null; };
  const xml = s => s ? new DOMParser().parseFromString(s, 'application/xml') : null;
  const docX = xml(await read('word/document.xml'));
  if (!docX || docX.getElementsByTagName('parsererror').length) throw new Error('document.xml illisible');
  const relsOf = async part => {
    const m = {}; const x = xml(await read(part.replace(/([^/]+)$/, '_rels/$1.rels')));
    if (x) [...x.getElementsByTagName('Relationship')].forEach(r => { m[r.getAttribute('Id')] = { target: r.getAttribute('Target'), type: r.getAttribute('Type') || '', ext: r.getAttribute('TargetMode') === 'External' }; });
    return m;
  };
  const RELS = await relsOf('word/document.xml');
  // thème (polices majeures / mineures)
  const theme = {};
  const themeX = xml(await read('word/theme/theme1.xml'));
  if (themeX) {
    const pick = tag => { const f = [...themeX.getElementsByTagName('*')].find(e => e.localName === tag); const l = f && [...f.children].find(c => c.localName === 'latin'); return l ? l.getAttribute('typeface') : ''; };
    theme.major = pick('majorFont'); theme.minor = pick('minorFont');
  }
  // styles
  const stylesX = xml(await read('word/styles.xml'));
  const STY = {}; let defPara = null; const docDef = { r: {}, p: {} };
  const readR = (rPr, o) => {
    if (!rPr) return o;
    for (const c of rPr.children) {
      const n = c.localName;
      if (n === 'b' || n === 'i' || n === 'caps' || n === 'vanish') o[n] = wOn(c);
      else if (n === 'strike' || n === 'dstrike') o.s = wOn(c);
      else if (n === 'u') o.u = wAt(c, 'val') !== 'none';
      else if (n === 'color') { const v = wAt(c, 'val'); if (v) o.color = v; }
      else if (n === 'sz') o.size = (+wAt(c, 'val')) / 2;
      else if (n === 'rFonts') {
        const t = wAt(c, 'asciiTheme') || wAt(c, 'hAnsiTheme');
        const f = wAt(c, 'ascii') || wAt(c, 'hAnsi') || (t ? (/^major/.test(t) ? theme.major : theme.minor) : '');
        if (f) o.font = f;
      }
      else if (n === 'highlight') o.hl = wAt(c, 'val');
      else if (n === 'shd') { const f = wAt(c, 'fill'); if (f && f !== 'auto') o.shd = f; }
      else if (n === 'vertAlign') o.va = wAt(c, 'val');
      else if (n === 'rStyle') o.rStyle = wAt(c, 'val');
    }
    return o;
  };
  const readP = (pPr, o) => {
    if (!pPr) return o;
    for (const c of pPr.children) {
      const n = c.localName;
      if (n === 'jc') o.jc = wAt(c, 'val');
      else if (n === 'spacing') { const a = wAt(c, 'after'), l = wAt(c, 'line'), lr = wAt(c, 'lineRule'); if (a != null) o.after = +a; if (l != null && (!lr || lr === 'auto')) o.line = +l / 240; }
      else if (n === 'ind') { const l = wAt(c, 'left') ?? wAt(c, 'start'); if (l != null) o.ind = +l; }
      else if (n === 'numPr') { const id = wAt(wKid(c, 'numId'), 'val'); const lv = wAt(wKid(c, 'ilvl'), 'val'); if (id != null) o.numId = id; if (lv != null) o.ilvl = +lv; }
      else if (n === 'outlineLvl') o.outline = +wAt(c, 'val');
      else if (n === 'pageBreakBefore') o.pbb = wOn(c);
      else if (n === 'pBdr') o.bdrBottom = !!wKid(c, 'bottom');
    }
    return o;
  };
  if (stylesX) {
    const dd = [...stylesX.documentElement.children].find(c => c.localName === 'docDefaults');
    if (dd) { readR(wKid(wKid(dd, 'rPrDefault'), 'rPr'), docDef.r); readP(wKid(wKid(dd, 'pPrDefault'), 'pPr'), docDef.p); }
    [...stylesX.documentElement.children].filter(c => c.localName === 'style').forEach(s => {
      const id = wAt(s, 'styleId');
      STY[id] = { id, type: wAt(s, 'type'), name: (wAt(wKid(s, 'name'), 'val') || id).toLowerCase(), based: wAt(wKid(s, 'basedOn'), 'val'), rPr: wKid(s, 'rPr'), pPr: wKid(s, 'pPr') };
      if (wAt(s, 'type') === 'paragraph' && /^(1|true|on)$/.test(wAt(s, 'default') || '')) defPara = id;
    });
  }
  const chain = id => { const out = []; const seen = new Set(); while (id && STY[id] && !seen.has(id)) { seen.add(id); out.unshift(STY[id]); id = STY[id].based; } return out; };
  const memoR = {}, memoP = {};
  const styleR = id => memoR[id] || (memoR[id] = chain(id).reduce((o, s) => readR(s.rPr, o), {}));
  const styleP = id => memoP[id] || (memoP[id] = chain(id).reduce((o, s) => readP(s.pPr, o), {}));
  const baseR = Object.assign({}, docDef.r, defPara ? styleR(defPara) : {});
  const baseP = Object.assign({}, docDef.p, defPara ? styleP(defPara) : {});
  delete baseR.rStyle;
  const headLevel = id => {
    for (const s of chain(id).reverse()) {
      let m = /^(?:heading|titre)\s*(\d)$/.exec(s.name) || /^(?:heading|titre)(\d)$/i.exec(s.id);
      if (m) return Math.min(3, +m[1]);
      if (s.name === 'title' || s.id === 'Title' || s.id === 'Titre') return 1;
      if (s.name === 'subtitle' || s.id === 'Subtitle' || s.id === 'Sous-titre') return 2;
    }
    const o = styleP(id).outline; return o != null && o < 9 ? Math.min(3, o + 1) : 0;
  };
  const isKind = (id, re) => chain(id).some(s => re.test(s.name) || re.test(s.id));
  // numérotation
  const numX = xml(await read('word/numbering.xml'));
  const ABS = {}, NUM = {};
  if (numX) {
    [...numX.documentElement.children].forEach(c => {
      if (c.localName === 'abstractNum') { const lv = {}; wKids(c, 'lvl').forEach(l => { lv[wAt(l, 'ilvl')] = wAt(wKid(l, 'numFmt'), 'val') || 'decimal'; }); ABS[wAt(c, 'abstractNumId')] = lv; }
      else if (c.localName === 'num') NUM[wAt(c, 'numId')] = wAt(wKid(c, 'abstractNumId'), 'val');
    });
  }
  const listType = (numId, ilvl) => { const a = ABS[NUM[numId]]; const f = a ? a[ilvl] || a[0] : 'bullet'; return f === 'bullet' || f === 'none' ? 'ul' : 'ol'; };
  // mise en page
  const sect = desc(docX.documentElement, 'sectPr').pop();
  const pgSz = wKid(sect, 'pgSz'), pgMar = wKid(sect, 'pgMar');
  const pw = +(wAt(pgSz, 'w') || 11906), ph = +(wAt(pgSz, 'h') || 16838);
  const ml = +(wAt(pgMar, 'left') || 1418), mr = +(wAt(pgMar, 'right') || 1418);
  const textW = Math.max(2000, pw - ml - mr) * 635; // EMU
  const page = { orient: wAt(pgSz, 'orient') === 'landscape' || pw > ph ? 'landscape' : 'portrait' };
  page.margin = Object.entries(MARGIN_TW).sort((a, b) => Math.abs(a[1] - ml) - Math.abs(b[1] - ml))[0][0];
  const hfText = async (kind) => {
    const ref = wKids(sect, kind + 'Reference').find(r => wAt(r, 'type') === 'default') || wKids(sect, kind + 'Reference')[0];
    const rel = ref && RELS[rAt(ref, 'id')]; if (!rel) return null;
    const x = xml(await read('word/' + rel.target.replace(/^\/?word\//, '')));
    if (!x) return null;
    const res = { text: [], num: null };
    desc(x.documentElement, 'p').forEach(p => {
      const instr = desc(p, 'instrText').map(t => t.textContent).join(' ') + ' ' + desc(p, 'fldSimple').map(f => wAt(f, 'instr')).join(' ');
      if (/\bPAGE\b/.test(instr)) { const jc = wAt(wKid(wKid(p, 'pPr'), 'jc'), 'val'); res.num = /NUMPAGES/.test(instr) ? 'xofy' : (jc === 'right' || jc === 'end') ? 'right' : 'center'; return; }
      const t = desc(p, 't').map(t => t.textContent).join('').trim(); if (t) res.text.push(t);
    });
    return res;
  };
  const H = await hfText('header'), F = await hfText('footer');
  if (H) { page.header = H.text.join(' · ').slice(0, 120); }
  if (F) { page.footer = F.text.join(' · ').slice(0, 120); if (F.num) page.num = F.num; }
  if (H && H.num && !page.num) page.num = H.num;
  // notes de bas de page
  const FN = {}; const fnOrder = [];
  const fnX = xml(await read('word/footnotes.xml'));
  if (fnX) desc(fnX.documentElement, 'footnote').forEach(f => { const id = wAt(f, 'id'); if (+id > 0) FN[id] = f; });
  // commentaires (fils de discussion et état « résolu » de Word 2013+)
  const CMT = [], CMW = {}, CM_ACTIVE = [];
  const cmPath = Object.values(RELS).find(r => /\/comments$/.test(r.type));
  const cmxPath = Object.values(RELS).find(r => /commentsExtended$/.test(r.type));
  const cmX = cmPath ? xml(await read('word/' + cmPath.target.replace(/^\/?word\//, '').replace(/^\//, ''))) : null;
  if (cmX) {
    const ex = {};
    const exX = cmxPath ? xml(await read('word/' + cmxPath.target.replace(/^\/?word\//, '').replace(/^\//, ''))) : null;
    if (exX) desc(exX.documentElement, 'commentEx').forEach(e => { const g = n => e.getAttribute('w15:' + n) || e.getAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', n); ex[g('paraId')] = { parent: g('paraIdParent'), done: g('done') === '1' }; });
    const byPid = {};
    desc(cmX.documentElement, 'comment').forEach(c => {
      const ps = desc(c, 'p');
      const text = ps.map(p => desc(p, 't').map(t => t.textContent).join('')).join('\n').trim();
      const last = ps[ps.length - 1];
      const pid = last && (last.getAttribute('w14:paraId') || last.getAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'paraId'));
      const at = Date.parse(wAt(c, 'date') || '') || Date.now();
      const it = { wid: wAt(c, 'id'), author: wAt(c, 'author') || 'Anonyme', ini: wAt(c, 'initials') || '', at, text, pid };
      const e = pid && ex[pid];
      const par = e && e.parent && byPid[e.parent];
      if (par) { par.replies.push({ id: 'r' + uid(), author: it.author, ini: it.ini, at, text }); CMW[it.wid] = par.id; }
      else { const root = { id: 'c' + uid(), author: it.author, ini: it.ini, at, text, done: !!(e && e.done), replies: [] }; CMT.push(root); CMW[it.wid] = root.id; if (pid) byPid[pid] = root; }
    });
  }
  const TRK = [];
  const cw = h => (h = TRK.length ? `<${TRK[TRK.length - 1].k} data-a="${escA(TRK[TRK.length - 1].a)}" data-t="${TRK[TRK.length - 1].t}">${h}</${TRK[TRK.length - 1].k}>` : h, CM_ACTIVE.length ? `<span data-c="${CM_ACTIVE[CM_ACTIVE.length - 1]}">${h}</span>` : h);
  // images
  const imgCache = {};
  const imgData = async rid => {
    const rel = RELS[rid]; if (!rel || rel.ext) return '';
    if (imgCache[rid] != null) return imgCache[rid];
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : 'word/' + rel.target;
    const f = zip.file(path.replace(/[^/]+\/\.\.\//g, ''));
    const ext = (path.split('.').pop() || '').toLowerCase();
    const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' }[ext];
    if (!f || !mime) { droppedImages++; return (imgCache[rid] = ''); }
    return (imgCache[rid] = `data:${mime};base64,` + await f.async('base64'));
  };
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escA = s => esc(s).replace(/"/g, '&quot;');

  // --- exécution (runs) ---
  const runStyle = (rPr, pBase) => {
    const own = readR(rPr, {});
    const st = own.rStyle && STY[own.rStyle] && STY[own.rStyle].name !== 'hyperlink' ? styleR(own.rStyle) : {};
    const e = Object.assign({}, pBase, st, own);
    const css = [];
    if (e.color && e.color !== 'auto' && e.color !== baseR.color && /^[0-9A-Fa-f]{6}$/.test(e.color)) css.push('color:#' + e.color);
    const bg = e.hl && e.hl !== 'none' ? HL_HEX[e.hl] : e.shd && /^[0-9A-Fa-f]{6}$/.test(e.shd) ? '#' + e.shd : '';
    if (bg) css.push('background-color:' + bg);
    if (e.size && e.size !== baseR.size) css.push('font-size:' + e.size + 'pt');
    if (e.font && mapFont(e.font) && mapFont(e.font) !== mapFont(baseR.font || '')) css.push(`font-family:'${mapFont(e.font)}'`);
    if (e.b) css.push('font-weight:bold');
    if (e.i) css.push('font-style:italic');
    const deco = [e.u && 'underline', e.s && 'line-through'].filter(Boolean);
    if (deco.length) css.push('text-decoration:' + deco.join(' '));
    if (e.va === 'superscript') css.push('vertical-align:super'); else if (e.va === 'subscript') css.push('vertical-align:sub');
    return { css: css.join(';'), caps: e.caps, hidden: e.vanish };
  };

  // Construit le HTML d'un paragraphe ; renvoie { html, breaks, extra }
  async function inlineOf(p, pBase, ctx) {
    const parts = []; // {css, html} ou {raw}
    let cur = null;
    const push = (css, html) => { if (cur && cur.css === css && !cur.raw) cur.html += html; else { cur = { css, html }; parts.push(cur); } };
    const raw = html => { cur = { raw: html }; parts.push(cur); cur = null; };
    let linkHref = null, linkParts = null;
    const flushLink = () => { if (linkParts) { const inner = linkParts.map(x => x.raw || (x.css ? `<span style="${x.css}">${x.html}</span>` : x.html)).join(''); parts.push({ raw: linkHref ? `<a href="${escA(linkHref)}">${inner}</a>` : inner }); linkParts = null; linkHref = null; cur = null; } };
    const fld = []; // pile de champs
    const out = { breaks: [], extra: [] };
    const visit = async (node) => {
      for (const c of node.children) {
        const n = c.localName;
        if (n === 'r') await run(c);
        else if (n === 'commentRangeStart') { const id = CMW[wAt(c, 'id')]; if (id && !CM_ACTIVE.includes(id)) CM_ACTIVE.push(id); }
        else if (n === 'commentRangeEnd') { const id = CMW[wAt(c, 'id')]; const i = CM_ACTIVE.indexOf(id); if (i >= 0) CM_ACTIVE.splice(i, 1); }
        else if (n === 'hyperlink') {
          const rel = RELS[rAt(c, 'id')];
          cur = null; const save = parts.length; await visit(c);
          const sub = parts.splice(save);
          const inner = sub.map(x => x.raw || (x.css ? `<span style="${x.css}">${x.html}</span>` : x.html)).join('');
          parts.push({ raw: rel && rel.ext ? `<a href="${escA(rel.target)}">${inner}</a>` : inner }); cur = null;
        }
        else if (n === 'fldSimple') {
          const ins = wAt(c, 'instr') || ''; const m = /HYPERLINK\s+"([^"]+)"/.exec(ins);
          cur = null; const save = parts.length; await visit(c); const sub = parts.splice(save);
          const inner = sub.map(x => x.raw || (x.css ? `<span style="${x.css}">${x.html}</span>` : x.html)).join('');
          parts.push({ raw: m ? `<a href="${escA(m[1])}">${inner}</a>` : inner }); cur = null;
        }
        else if (n === 'sdt') await visit(wKid(c, 'sdtContent') || c);
        else if (n === 'ins' || n === 'del') { TRK.push({ k: n, a: wAt(c, 'author') || 'Anonyme', t: Date.parse(wAt(c, 'date') || '') || Date.now() }); await visit(c); TRK.pop(); }
        else if (/^(smartTag|customXml|moveTo|sdtContent|AlternateContent|Choice)$/.test(n)) {
          if (n === 'AlternateContent') { const ch = wKid(c, 'Choice'); await visit(ch || wKid(c, 'Fallback') || c); }
          else await visit(c);
        }
      }
    };
    const run = async (r) => {
      const rp = runStyle(wKid(r, 'rPr'), pBase);
      for (const c of r.children) {
        const n = c.localName;
        if (n === 'fldChar') {
          const t = wAt(c, 'fldCharType');
          if (t === 'begin') fld.push({ instr: '', phase: 'instr' });
          else if (t === 'separate' && fld.length) {
            const f = fld[fld.length - 1]; f.phase = 'result';
            const m = /HYPERLINK\s+"([^"]+)"/.exec(f.instr);
            if (/^\s*TOC\b/.test(f.instr)) ctx.toc = true;
            if (m) { f.link = true; flushLink(); linkHref = m[1]; linkParts = []; }
          } else if (t === 'end' && fld.length) { const f = fld.pop(); if (f.link) flushLink(); if (/^\s*TOC\b/.test(f.instr)) ctx.tocEnd = true; }
          continue;
        }
        if (n === 'instrText') { if (fld.length) fld[fld.length - 1].instr += c.textContent; continue; }
        if (fld.length && fld[fld.length - 1].phase === 'instr') continue;
        if (rp.hidden) continue;
        const target = linkParts ? { push: (css, html) => linkParts.push({ css, html: cw(html) }) } : { push: (css, html) => push(css, cw(html)) };
        if (n === 't' || n === 'delText') { let t = c.textContent; if (rp.caps) t = t.toUpperCase(); if (t) target.push(rp.css, esc(t)); }
        else if (n === 'tab' || n === 'ptab') target.push(rp.css, '    ');
        else if (n === 'noBreakHyphen') target.push(rp.css, '‑');
        else if (n === 'br' || n === 'cr') {
          const ty = wAt(c, 'type');
          if (ty === 'page' || ty === 'column') { flushLink(); raw('\u0000PB\u0000'); }
          else raw('<br>');
        }
        else if (n === 'sym') { const ch = parseInt(wAt(c, 'char'), 16); if (ch === 0xF0FE || ch === 0xF078) target.push(rp.css, '☒'); else if (ch === 0xF0A8 || ch === 0xF06F) target.push(rp.css, '☐'); }
        else if (n === 'footnoteReference') { const id = wAt(c, 'id'); if (FN[id]) { fnOrder.push(id); raw(`<sup>${fnOrder.length}</sup>`); } }
        else if (n === 'drawing' || n === 'pict' || n === 'AlternateContent' || n === 'object') {
          const host = n === 'AlternateContent' ? (wKid(c, 'Choice') || c) : c;
          const blip = desc(host, 'blip')[0], vimg = desc(host, 'imagedata')[0];
          const rid = blip ? rAt(blip, 'embed') : vimg ? rAt(vimg, 'id') : null;
          if (rid) {
            const src = await imgData(rid);
            if (src) {
              const ext = desc(host, 'extent')[0];
              const cx = ext ? +ext.getAttribute('cx') : 0;
              const pct = cx ? Math.max(8, Math.min(100, Math.round(cx / textW * 100))) : 0;
              const anc = desc(host, 'anchor')[0];
              let wrap = '';
              if (anc && desc(anc, 'wrapSquare').concat(desc(anc, 'wrapTight'), desc(anc, 'wrapThrough')).length) {
                const ph = desc(anc, 'positionH')[0];
                const al = ph && desc(ph, 'align')[0], off = ph && desc(ph, 'posOffset')[0];
                const a = al ? al.textContent.trim() : '';
                if (a === 'right' || a === 'outside') wrap = 'right';
                else if (a === 'left' || a === 'inside') wrap = 'left';
                else if (a === 'center') wrap = '';
                else wrap = off && (+off.textContent + cx / 2) > textW / 2 ? 'right' : 'left';
              }
              raw(`<img src="${src}" alt=""${pct ? ` style="width:${pct}%"` : ''}${wrap ? ` data-wrap="${wrap}"` : ''}>`);
            }
          }
          // zones de texte : leur contenu est ajouté après le paragraphe
          for (const tx of desc(host, 'txbxContent')) out.extra.push(tx);
        }
      }
    };
    await visit(p);
    flushLink();
    const html = parts.map(x => x.raw != null ? x.raw : x.css ? `<span style="${x.css}">${x.html}</span>` : x.html).join('');
    out.html = html;
    return out;
  }

  // --- blocs ---
  async function blocks(container, ctx) {
    let html = '', secStart = 0;
    const colsOf = sp => { const c = desc(sp, 'cols')[0]; return Math.min(3, +(c && wAt(c, 'num')) || 1); };
    // fin de section : colonnes éventuelles, puis nouvelle page sauf si la section suivante est « continue »
    const sectBreak = (spr, el) => {
      endSection(spr);
      const nxt = [...container.children].slice([...container.children].indexOf(el) + 1).map(x => x.localName === 'sectPr' ? x : desc(x, 'sectPr')[0]).find(Boolean);
      const ty = nxt && desc(nxt, 'type')[0];
      if (el.nextElementSibling && !(ty && wAt(ty, 'val') === 'continuous')) emit('<hr data-pb="1">');
    };
    const endSection = sp => { closeLists(); closeQuote(); const k = colsOf(sp); if (k > 1 && html.length > secStart) html = html.slice(0, secStart) + `<div data-cols="${k}">` + html.slice(secStart) + '</div>'; secStart = html.length; };
    let lists = []; // pile {tag, lvl, el:string-builder}
    let quote = null;
    const closeLists = () => { while (lists.length) closeTop(); };
    const closeTop = () => {
      const t = lists.pop();
      const s = `<${t.tag}>${t.items.map(i => `<li${i.style ? ` style="${i.style}"` : ''}>${i.html}${i.sub}</li>`).join('')}</${t.tag}>`;
      if (lists.length) { const parent = lists[lists.length - 1]; if (!parent.items.length) parent.items.push({ html: '', sub: '' }); parent.items[parent.items.length - 1].sub += s; }
      else html += s;
    };
    const closeQuote = () => { if (quote) { html += `<blockquote>${quote.join('')}</blockquote>`; quote = null; } };
    const emit = s => { closeLists(); closeQuote(); html += s; };
    for (const el of container.children) {
      const n = el.localName;
      if (n === 'sdt') {
        const g = desc(wKid(el, 'sdtPr'), 'docPartGallery')[0];
        if (g && /Table of Contents/i.test(wAt(g, 'val') || '')) { if (!ctx.tocDone) { emit('<div data-toc="1"></div>'); ctx.tocDone = true; } continue; }
        const sub = await blocks(wKid(el, 'sdtContent') || el, ctx); emit(sub); continue;
      }
      if (/^(customXml|ins|moveTo)$/.test(n)) { emit(await blocks(el, ctx)); continue; }
      if (n === 'tbl') { emit(await tableHtml(el, ctx)); ctx.afterTbl = true; continue; }
      if (n !== 'p') continue;
      // Word impose un paragraphe vide après un tableau : on ne le reprend pas
      const afterTbl = ctx.afterTbl; ctx.afterTbl = false;
      if (afterTbl && !el.textContent.trim() && !desc(el, 'drawing').length && !desc(el, 'pict').length) continue;
      // paragraphe vide qui ne sert qu'à porter un saut de section : on ne garde que le saut
      { const sp0 = desc(wKid(el, 'pPr'), 'sectPr')[0]; if (sp0 && !el.textContent.trim() && !desc(el, 'drawing').length && !desc(el, 'pict').length) { sectBreak(sp0, el); continue; } }
      const pPr = wKid(el, 'pPr');
      const own = readP(pPr, {});
      const sid = wAt(wKid(pPr, 'pStyle'), 'val') || defPara;
      const sp = sid ? styleP(sid) : {};
      const eff = Object.assign({}, docDef.p, sp, own);
      const lvl = sid ? headLevel(sid) : 0;
      const isToc = sid && isKind(sid, /^(toc ?\d|toc heading|tocheading|en-tête de table des matières)$/i);
      const structural = lvl || (sid && isKind(sid, /^(quote|intense quote|citation|citation intense|table text|list paragraph|paragraphe de liste|header|footer|en-tête|pied de page|caption|légende)$/i));
      const pBase = Object.assign({}, baseR, structural ? {} : (sid ? styleR(sid) : {}));
      delete pBase.rStyle;
      const wasInToc = ctx.inToc;
      const r = await inlineOf(el, pBase, ctx);
      if (isToc || wasInToc || ctx.toc) {
        if (ctx.toc) { ctx.inToc = true; ctx.toc = false; }
        if (ctx.tocEnd) { ctx.inToc = false; ctx.tocEnd = false; }
        if (!ctx.tocDone) { emit('<div data-toc="1"></div>'); ctx.tocDone = true; }
        continue;
      }
      // style du paragraphe
      const css = [];
      const jc = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[eff.jc];
      if (jc) css.push('text-align:' + jc);
      if (!lvl) {
        // interligne simple de Word (1 à 1,2) = interligne normal de Grimoire
        if (eff.line && Math.abs(eff.line - (baseP.line || 0)) > 0.01 && eff.line >= 1.3 && eff.line <= 3) css.push('line-height:' + (Math.round(eff.line * 100) / 100));
        if (eff.after != null && eff.after !== baseP.after) { if (eff.after === 0) css.push('margin-bottom:0px'); else if (eff.after >= 270) css.push('margin-bottom:' + Math.round(eff.after / 15) + 'px'); }
      }
      const style = css.join(';');
      const numId = own.numId ?? sp.numId;
      const rawLvl = own.ilvl ?? sp.ilvl ?? 0;
      let ilvl = rawLvl;
      const lsn = sid && STY[sid] && /^list (?:bullet|number|continue)\s*(\d)$/.exec(STY[sid].name);
      if (lsn) ilvl = Math.max(ilvl, +lsn[1] - 1);
      const segs = r.html.split('\u0000PB\u0000');
      if (eff.pbb) emit('<hr data-pb="1">');
      const text = r.html.replace(/<[^>]+>/g, '').replace(/&nbsp;| /g, ' ').trim();
      const hasImg = /<img/.test(r.html);
      if (numId && numId !== '0' && NUM[numId] != null && !lvl) {
        closeQuote();
        const tag = listType(numId, rawLvl);
        while (lists.length && lists[lists.length - 1].lvl > ilvl) closeTop();
        if (lists.length && lists[lists.length - 1].lvl === ilvl && lists[lists.length - 1].tag !== tag) closeTop();
        if (!lists.length || lists[lists.length - 1].lvl < ilvl) lists.push({ tag, lvl: ilvl, items: [] });
        lists[lists.length - 1].items.push({ html: segs.join('') || '<br>', sub: '', style: jc ? 'text-align:' + jc : '' });
      } else if (!lvl && sid && isKind(sid, /^(quote|intense quote|citation|citation intense)$/i)) {
        closeLists(); quote ||= []; quote.push(`<p>${segs.join('') || '<br>'}</p>`);
      } else if (!lvl && !text && !hasImg && eff.bdrBottom && segs.length === 1) {
        emit('<hr>');
      } else {
        const ind = !lvl && eff.ind > 0 && eff.ind !== baseP.ind ? `margin-left:${Math.round(eff.ind / 15)}px` : '';
        const st = [style, ind].filter(Boolean).join(';');
        const tag = lvl ? 'h' + lvl : 'p';
        segs.forEach((s, k) => {
          if (k > 0) emit('<hr data-pb="1">');
          if (segs.length > 1 && !s.replace(/<br>/g, '').trim()) return;
          emit(`<${tag}${st ? ` style="${st}"` : ''}>${s || '<br>'}</${tag}>`);
        });
      }
      for (const tx of r.extra) emit(await blocks(tx, ctx));
      const spr = desc(pPr, 'sectPr')[0];
      if (spr) sectBreak(spr, el);
    }
    closeLists(); closeQuote();
    const last = wKid(container, 'sectPr'); if (last) endSection(last);
    return html;
  }

  async function tableHtml(tbl, ctx) {
    const grid = wKids(wKid(tbl, 'tblGrid'), 'gridCol').map(g => +wAt(g, 'w') || 0);
    const tot = grid.reduce((s, x) => s + x, 0);
    let out = '<table>';
    if (grid.length > 1 && tot > 0) out += '<colgroup>' + grid.map(w => `<col style="width:${(Math.round(w / tot * 1000) / 10)}%">`).join('') + '</colgroup>';
    let head = '', body = '';
    for (const tr of wKids(tbl, 'tr')) {
      const isHead = !!wKid(wKid(tr, 'trPr'), 'tblHeader') && !body;
      let row = '';
      for (const tc of wKids(tr, 'tc')) {
        const tcPr = wKid(tc, 'tcPr');
        const span = +(wAt(wKid(tcPr, 'gridSpan'), 'val') || 1);
        const vm = wKid(tcPr, 'vMerge');
        const fill = wAt(wKid(tcPr, 'shd'), 'fill');
        const t = isHead ? 'th' : 'td';
        const css = fill && /^[0-9A-Fa-f]{6}$/.test(fill) && !(isHead && fill.toUpperCase() === 'EAD9B8') ? `background-color:#${fill}` : '';
        let inner = vm && wAt(vm, 'val') !== 'restart' ? '' : await blocks(tc, ctx);
        // une seule ligne : on retire le <p> englobant
        const m = /^<p(?: style="([^"]*)")?>([\s\S]*)<\/p>$/.exec(inner);
        let al = '';
        if (m && !/<\/p>/.test(m[2])) { inner = m[2] === '<br>' ? '' : m[2]; const a = /text-align:(\w+)/.exec(m[1] || ''); if (a) al = 'text-align:' + a[1]; }
        const st = [css, al].filter(Boolean).join(';');
        row += `<${t}${span > 1 ? ` colspan="${span}"` : ''}${st ? ` style="${st}"` : ''}>${inner || '<br>'}</${t}>`;
      }
      if (isHead) head += `<tr>${row}</tr>`; else body += `<tr>${row}</tr>`;
    }
    return out + (head ? `<thead>${head}</thead>` : '') + `<tbody>${body}</tbody></table>`;
  }

  const body = wKid(docX.documentElement, 'body');
  const ctx = {};
  let html = await blocks(body, ctx);
  if (fnOrder.length) {
    let notes = '<hr>';
    for (let k = 0; k < fnOrder.length; k++) {
      const inner = (await blocks(FN[fnOrder[k]], {})).replace(/^<p([^>]*)>/, `<p$1><sup>${k + 1}</sup> `);
      notes += inner;
    }
    html += notes;
  }
  const used = new Set([...html.matchAll(/data-c="([\w-]+)"/g)].map(m => m[1]));
  return { html, page, comments: CMT.filter(c => used.has(c.id)) };
}

async function importFile(f) {
  if (state.cur) { if (state.view === 'sheet') commitEdit(false); await flushSave(); }
  const ext = (f.name.split('.').pop() || '').toLowerCase();
  const title = f.name.replace(/\.[^.]+$/, '');
  toast('Import de « ' + title + ' »…');
  try {
    if (ext === 'pdf') { await importPdfFiles([f]); return; }
    if (ext === 'docx') {
      droppedImages = 0;
      let raw, page;
      let comments = [];
      try { ({ html: raw, page, comments } = await docxToHtml(await f.arrayBuffer())); }
      catch (err) {
        console.warn('Import direct impossible, repli sur mammoth', err);
        droppedImages = 0;
        await loadScript(LIB.mammoth);
        raw = (await window.mammoth.convertToHtml({ arrayBuffer: await f.arrayBuffer() }, { styleMap: ['u => u', 'strike => s'] })).value;
      }
      const box = document.createElement('div'); box.innerHTML = restoreChecklists(sanitize(raw));
      const html = attachFloats(box).innerHTML;
      await createDoc('texte', Object.assign({ title, html }, page ? { page } : {}, comments && comments.length ? { comments } : {}));
      toast(droppedImages ? `Document importé. ${droppedImages} image${droppedImages > 1 ? 's' : ''} dans un format non pris en charge.` : 'Document importé.');
    } else if (['xlsx', 'xlsm', 'xls', 'ods', 'csv'].includes(ext)) {
      await loadScript(LIB.xlsx);
      const X = window.XLSX;
      const wb = ext === 'csv' ? X.read(await f.text(), { type: 'string', FS: guessSep(await f.text()) }) : X.read(await f.arrayBuffer(), { cellFormula: true, cellStyles: true });
      let cut = false;
      const sheets = wb.SheetNames.slice(0, 20).map(name => {
      const ws = wb.Sheets[name];
      const ref = X.utils.decode_range(ws['!ref'] || 'A1');
      if (ref.e.r > 999 || ref.e.c > 51) cut = true;
      const maxR = Math.min(ref.e.r, 999), maxC = Math.min(ref.e.c, 51);
      const cells = {};
      for (let r = 0; r <= maxR; r++) for (let c = 0; c <= maxC; c++) {
        const cl = ws[X.utils.encode_cell({ r, c })]; if (!cl) continue;
        let v, fm = {};
        if (cl.f) v = '=' + excelToFr(cl.f);
        else if (cl.t === 'n' && cl.z && X.SSF.is_date(cl.z)) { v = fmtDate(cl.v); fm.n = 'date'; }
        else if (cl.t === 'n') { v = numToRaw(cl.v); const z = String(cl.z || ''); if (z.includes('%')) fm.n = 'pct'; else if (z.includes('€')) fm.n = 'eur'; else if (/0\.00/.test(z)) fm.n = 'dec2'; }
        else if (cl.t === 'b') v = cl.v ? 'VRAI' : 'FAUX';
        else if (cl.t === 'd') v = cl.w || String(cl.v);
        else v = cl.v == null ? '' : String(cl.v);
        if (cl.f && cl.z) { const z = String(cl.z); if (z.includes('%')) fm.n = 'pct'; else if (z.includes('€')) fm.n = 'eur'; else if (X.SSF.is_date(z)) fm.n = 'date'; }
        if (cl.c && cl.c.length) { const t = cl.c.map(x => String(x.t || '').replace(/^\[Threaded comment\][\s\S]*?Comment:\s*/i, '').trim()).filter(Boolean).join('\n'); if (t) fm.note = t.slice(0, 2000); }
        if (cl.s && cl.s.font) {
          if (cl.s.font.bold) fm.b = 1; if (cl.s.font.italic) fm.i = 1; if (cl.s.font.underline) fm.u = 1;
          const fn = mapFont(cl.s.font.name || ''); if (fn && fn !== 'Calibri' && GRID_FONTS.includes(fn)) fm.ff = fn;
          const sz = +cl.s.font.sz; if (sz && sz !== 11 && sz !== 10 && sz >= 6 && sz <= 72) fm.fs = GRID_SIZES.reduce((a, b) => Math.abs(b - sz) < Math.abs(a - sz) ? b : a);
        }
        const zz = String(cl.z || ''); const dm = /0\.(0+)/.exec(zz);
        if (cl.t === 'n' && fm.n !== 'date' && zz && zz !== 'General') { const nd = dm ? dm[1].length : 0; if (!(nd === 2 && (fm.n === 'eur' || fm.n === 'dec2' || fm.n === 'pct')) && (dm || fm.n)) fm.d = nd; }
        if (cl.c && cl.c.length) fm.note = cl.c.map(x => String(x.t || '').replace(/^[^:\n]{1,40}:\n/, '')).join('\n').trim().slice(0, 2000);
        if (v !== '' || Object.keys(fm).length) cells[addr(r, c)] = { v, f: fm };
      }
      (ws['!merges'] || []).forEach(m => { if (m.s.r > maxR || m.s.c > maxC) return; const k = addr(m.s.r, m.s.c); (cells[k] ||= { v: '', f: {} }).f.mg = { r: m.e.r - m.s.r + 1, c: m.e.c - m.s.c + 1 }; });
      const hid = { r: [], c: [] };
      (ws['!rows'] || []).forEach((rw, i) => { if (rw && rw.hidden && i <= maxR) hid.r.push(i); });
      (ws['!cols'] || []).forEach((cw, i) => { if (cw && cw.hidden && i <= maxC) hid.c.push(i); });
      const colW = {}; (ws['!cols'] || []).slice(0, maxC + 1).forEach((cw, i) => { if (cw && cw.wpx) colW[i] = Math.round(Math.max(40, Math.min(cw.wpx, 420))); else if (cw && cw.wch) colW[i] = Math.round(Math.max(40, Math.min(cw.wch * 7.5, 420))); });
      return { id: uid(), src: name, name: String(name).replace(/[\[\]:*?\/\\']/g, '').slice(0, 31) || 'Feuille', cells, rows: Math.max(60, maxR + 21), cols: Math.max(12, maxC + 2), colW, hid, cf: [], charts: [] };
      });
      let vba = null;
      if (ext === 'xlsm') { try { vba = await readVba(await f.arrayBuffer(), sheets); } catch (err) { console.warn('Macros non lues', err); } }
      if (ext === 'xlsx' || ext === 'xlsm') {
        try {
          const extra = await xlsxReadExtras(await f.arrayBuffer(), sheets.map(s => s.src));
          sheets.forEach(sh => {
            const e = extra[sh.src]; if (!e) return;
            sh.cf = e.cf; sh.charts = e.charts;
            Object.entries(e.fmt).forEach(([k, o]) => { const p = parseAddr(k); if (!p || p.r >= sh.rows || p.c >= sh.cols) return; const x = sh.cells[k] || (sh.cells[k] = { v: '', f: {} }); x.f = { ...o, ...(x.f || {}) }; });
            Object.entries(e.dd).forEach(([k, dd]) => { const x = sh.cells[k] || (sh.cells[k] = { v: '', f: {} }); x.f ||= {}; x.f.dd = dd; });
          });
        } catch (err) { console.warn('Compléments Excel ignorés', err); }
      }
      const names = {};
      ((wb.Workbook && wb.Workbook.Names) || []).forEach(n => {
        if (n.Hidden || n.Sheet != null || !n.Name || /^_/.test(n.Name)) return;
        const m = /^'?(.+?)'?!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(String(n.Ref || '')); if (!m) return;
        const sh = sheets.find(s => s.src === m[1].replace(/''/g, "'")); const o = rangeObj(m[2]);
        if (sh && o && nameOk(n.Name)) names[n.Name] = { s: sh.name, ref: rngRef(o) };
      });
      sheets.forEach(s => { delete s.src; });
      await createDoc('feuille', Object.assign({ title, sheets, active: 0, names }, vba ? { vba } : {}));
      const more = wb.SheetNames.length > 20 ? ' (20 premières feuilles)' : '';
      if (vba) toast('Classeur importé. Ses macros sont conservées telles quelles : Grimoire ne les exécute pas, Excel les retrouvera à l\'enregistrement en .xlsm.');
      else toast(cut ? `Classeur importé${more}, limité à 1 000 lignes et 52 colonnes par feuille.` : `Classeur importé : ${sheets.length} feuille${sheets.length > 1 ? 's' : ''}${more}.`);
    } else if (ext === 'txt' || ext === 'md') {
      const text = await f.text();
      const html = text.split(/\n{2,}/).map(p => { const t = p.trim(); const h = /^(#{1,3})\s+(.*)$/.exec(t); return h ? `<h${h[1].length}>${esc(h[2])}</h${h[1].length}>` : `<p>${esc(t).replace(/\n/g, '<br>')}</p>`; }).join('');
      await createDoc('texte', { title, html });
      toast('Texte importé.');
    } else toast('Format non pris en charge. Importez un fichier .docx, .xlsx, .csv ou .txt.', true);
  } catch (err) {
    toast('Import impossible : le fichier est peut-être protégé ou endommagé.', true);
  }
}
function guessSep(t) { const l = t.split('\n')[0] || ''; return (l.match(/;/g) || []).length > (l.match(/,/g) || []).length ? ';' : (l.includes('\t') ? '\t' : ','); }

/* =========================================================
   Export
   ========================================================= */
const MIME = { pdf: 'application/pdf', zip: 'application/zip', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', csv: 'text/csv', md: 'text/markdown', html: 'text/html', json: 'application/json' };
let CAPTURE = null;
async function saveFile(filename, data) {
  if (CAPTURE) { CAPTURE({ filename, data }); return; }
  const ext = filename.split('.').pop().toLowerCase();
  const blob = new Blob([data], { type: MIME[ext] || 'application/octet-stream' });
  if (window.showSaveFilePicker && !isPhone()) {
    try {
      const h = await window.showSaveFilePicker({ suggestedName: filename, types: [{ description: ext.toUpperCase(), accept: { [MIME[ext] || 'application/octet-stream']: ['.' + ext] } }] });
      const w = await h.createWritable(); await w.write(blob); await w.close();
      toast('Enregistré : ' + h.name); return;
    } catch (e) { if (e && e.name === 'AbortError') return; }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = filename; a.hidden = true;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
  toast('Fichier enregistré dans vos téléchargements : ' + filename);
}
/* =========================================================
   Commentaires (documents) : dans la marge de droite, comme Word et Google Docs
   ========================================================= */
const CM = { active: null, edit: null, showDone: false, panel: false };
const cmAuthor = () => Object.assign({ name: 'Cyril', ini: 'CSC' }, ls.get('grim.author', {}));
const cmList = () => (state.cur.comments ||= []);
const cmFind = id => (state.cur && state.cur.comments || []).find(c => c.id === id);
const cmAnchors = id => [...ed.querySelectorAll('[data-c]')].filter(el => el.dataset.c === id);
const cmRail = $('#cmtRail');
function cmWordRange(sel) {
  const n = sel.anchorNode; if (!n || n.nodeType !== 3) return null;
  const t = n.nodeValue, o = sel.anchorOffset, W1 = /[\p{L}\p{N}'’\-]/u;
  let s = o, e = o; while (s > 0 && W1.test(t[s - 1])) s--; while (e < t.length && W1.test(t[e])) e++;
  if (s === e) return null;
  const r = document.createRange(); r.setStart(n, s); r.setEnd(n, e); return r;
}
function cmWrap(r, id) {
  const root = r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentNode;
  const nodes = [], tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let n; while ((n = tw.nextNode())) if (n.nodeValue.length && r.intersectsNode(n) && ed.contains(n)) nodes.push(n);
  const out = [];
  nodes.forEach(t => {
    let s = t === r.startContainer ? r.startOffset : 0, e = t === r.endContainer ? r.endOffset : t.nodeValue.length;
    if (s >= e || !t.nodeValue.slice(s, e).trim() && nodes.length > 1) return;
    if (t.parentNode.closest('[data-toc],[contenteditable="false"]')) return;
    if (e < t.nodeValue.length) t.splitText(e);
    if (s > 0) t = t.splitText(s);
    const sp = document.createElement('span'); sp.dataset.c = id; t.replaceWith(sp); sp.appendChild(t); out.push(sp);
  });
  return out;
}
function addComment() {
  if (state.view !== 'writer') return;
  if (document.activeElement !== ed) restoreRange();
  const sel = getSelection();
  if (!sel.rangeCount || !ed.contains(sel.anchorNode)) { toast('Sélectionnez le texte à commenter.'); return; }
  let r = sel.getRangeAt(0);
  if (r.collapsed || !r.toString().trim()) r = cmWordRange(sel);
  if (!r) { toast('Sélectionnez le texte à commenter.'); return; }
  const id = 'c' + uid();
  if (!cmWrap(r, id).length) { toast('Ce passage ne peut pas recevoir de commentaire.', true); return; }
  const a = cmAuthor();
  cmList().push({ id, author: a.name, ini: a.ini, at: Date.now(), text: '', done: false, replies: [] });
  sel.removeAllRanges();
  CM.active = id; CM.edit = id; CM.panel = true;
  state.cur.html = edHtml(); scheduleSave();
  renderComments();
  setTimeout(() => { const ta = cmRail.querySelector(`[data-cid="${id}"] textarea`); if (ta) ta.focus(); }, 30);
}
function cmUnwrap(id) { cmAnchors(id).forEach(sp => sp.replaceWith(...sp.childNodes)); ed.normalize(); }
function deleteComment(id) {
  cmUnwrap(id);
  state.cur.comments = cmList().filter(c => c.id !== id);
  if (CM.active === id) CM.active = null; if (CM.edit === id) CM.edit = null;
  state.cur.html = edHtml(); scheduleSave(); renderComments();
}
function cmPurgeOrphans() {
  if (!state.cur || !state.cur.comments) return;
  const ids = new Set([...ed.querySelectorAll('[data-c]')].map(el => el.dataset.c));
  const keep = state.cur.comments.filter(c => ids.has(c.id) && (c.text || (c.replies || []).length));
  if (keep.length !== state.cur.comments.length) state.cur.comments = keep;
  // ancres sans commentaire (copier-coller, annulation) : on les retire
  const known = new Set(keep.map(c => c.id));
  let changed = false;
  ed.querySelectorAll('[data-c]').forEach(sp => { if (!known.has(sp.dataset.c)) { sp.replaceWith(...sp.childNodes); changed = true; } });
  if (changed) { ed.normalize(); state.cur.html = edHtml(); }
}
const cmTxt = t => esc(t || '').replace(/\n/g, '<br>');
function cmItem(c, k) {
  return `<div class="cmt-h"><span class="cmt-av">${esc(c.ini || (c.author || '?').slice(0, 2).toUpperCase())}</span><span class="cmt-who"><b>${esc(c.author || 'Anonyme')}</b><time>${relDate(c.at)}</time></span>${k == null ? `<button type="button" data-ca="done" title="${c.done ? 'Rouvrir' : 'Résoudre'}" aria-label="${c.done ? 'Rouvrir le commentaire' : 'Résoudre le commentaire'}">${ic(c.done ? 'undo' : 'tick')}</button><button type="button" data-ca="del" title="Supprimer le commentaire" aria-label="Supprimer le commentaire">${ic('trash')}</button>` : `<button type="button" data-ca="rdel" data-k="${k}" title="Supprimer la réponse" aria-label="Supprimer la réponse">${ic('close')}</button>`}</div>`
    + `<div class="cmt-t" ${k == null ? 'data-ca="edit"' : ''}>${cmTxt(c.text) || '<em>Sans texte</em>'}</div>`;
}
function renderComments() {
  if (!cmRail) return;
  if (state.view !== 'writer' || !state.cur) { cmRail.innerHTML = ''; $('#cmtCss').textContent = ''; return; }
  const list = state.cur.comments || [];
  const firsts = new Map(); ed.querySelectorAll('[data-c]').forEach(el => { if (!firsts.has(el.dataset.c)) firsts.set(el.dataset.c, el); });
  const live = list.filter(c => firsts.has(c.id));
  const open = live.filter(c => !c.done), done = live.length - open.length;
  const shown = live.filter(c => CM.showDone || !c.done || c.id === CM.active);
  shown.sort((a, b) => { const x = firsts.get(a.id), y = firsts.get(b.id); return x === y ? 0 : x.compareDocumentPosition(y) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1; });
  // surlignage : actif plus marqué, résolus discrets
  const sel = id => `.page [data-c="${CSS.escape(id)}"]`;
  $('#cmtCss').textContent = live.filter(c => c.done && c.id !== CM.active).map(c => sel(c.id)).join(',') + (done ? '{background:none;border-bottom:0}' : '')
    + (CM.active && firsts.has(CM.active) ? sel(CM.active) + '{background:var(--cmt-on);border-bottom-color:var(--copper)}' : '');
  const wrap = ed.parentNode, scroll = $('#wScroll');
  cmRail.innerHTML = (live.length ? `<div class="cmt-top"><span>${open.length} commentaire${open.length > 1 ? 's' : ''}</span>${done ? `<button type="button" class="lnk" data-ca="showdone">${CM.showDone ? 'Masquer' : 'Afficher'} les résolus (${done})</button>` : ''}<button type="button" class="cmt-x" data-ca="close" aria-label="Fermer les commentaires">${ic('close')}</button></div>` : '')
    + shown.map(c => {
      const editing = CM.edit === c.id, act = CM.active === c.id;
      return `<div class="cmt-card${act ? ' on' : ''}${c.done ? ' done' : ''}" data-cid="${esc(c.id)}">
        ${editing ? `<div class="cmt-h"><span class="cmt-av">${esc(c.ini || '')}</span><span class="cmt-who"><b>${esc(c.author || '')}</b></span></div><textarea class="cmt-in" rows="2" placeholder="Votre commentaire…" data-ci="main">${esc(c.text || '')}</textarea><div class="cmt-acts"><button type="button" class="pbtn" data-ca="save">Commenter</button><button type="button" class="lnk" data-ca="cancel">Annuler</button></div>`
          : cmItem(c)}
        ${(c.replies || []).map((r, k) => `<div class="cmt-rep">${cmItem(r, k)}</div>`).join('')}
        ${act && !editing && !c.done ? `<textarea class="cmt-in" rows="1" placeholder="Répondre…" data-ci="reply"></textarea>` : ''}
      </div>`;
    }).join('');
  const compact = !live.length ? false : layoutComments(shown, firsts);
  scroll.classList.toggle('cmt-compact', compact);
  scroll.classList.toggle('cmt-open', compact && CM.panel);
  updateStats();
}
function layoutComments(shown, firsts) {
  const wrap = ed.parentNode, scroll = $('#wScroll');
  wrap.style.transform = '';
  const cs = getComputedStyle(scroll), inner = scroll.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
  const free = (inner - wrap.offsetWidth) / 2, need = 304;
  const compact = isPhone() || free * 2 < need;
  if (compact) { cmRail.querySelectorAll('.cmt-card').forEach(el => { el.style.top = ''; }); return true; }
  if (free < need) wrap.style.transform = `translateX(${-(need - free)}px)`;
  const top0 = wrap.getBoundingClientRect().top;
  let y = cmRail.querySelector('.cmt-top') ? cmRail.querySelector('.cmt-top').offsetHeight + 8 : 0;
  const cards = [...cmRail.querySelectorAll('.cmt-card')];
  const want = cards.map(el => { const a = firsts.get(el.dataset.cid); return a ? a.getBoundingClientRect().top - top0 - 6 : 0; });
  // la carte active est alignée sur son texte, les autres s'écartent autour
  const ai = cards.findIndex(el => el.dataset.cid === CM.active);
  const pos = new Array(cards.length);
  if (ai >= 0) {
    pos[ai] = Math.max(want[ai], y);
    for (let i = ai - 1, lim = pos[ai]; i >= 0; i--) { pos[i] = Math.min(want[i], lim - cards[i].offsetHeight - 8); lim = pos[i]; }
    for (let i = ai + 1, lim = pos[ai] + cards[ai].offsetHeight + 8; i < cards.length; i++) { pos[i] = Math.max(want[i], lim); lim = pos[i] + cards[i].offsetHeight + 8; }
    const under = y - Math.min(...pos.slice(0, ai + 1)); if (under > 0) for (let i = 0; i < cards.length; i++) pos[i] += under;
  } else cards.forEach((el, i) => { pos[i] = Math.max(want[i], y); y = pos[i] + el.offsetHeight + 8; });
  cards.forEach((el, i) => { el.style.top = pos[i] + 'px'; });
  return false;
}
// Repositionne les cartes sans les reconstruire (on ne perd pas la saisie en cours).
const relayoutComments = debounce(() => {
  if (state.view !== 'writer' || !state.cur || !(state.cur.comments || []).length) return;
  const firsts = new Map(); ed.querySelectorAll('[data-c]').forEach(el => { if (!firsts.has(el.dataset.c)) firsts.set(el.dataset.c, el); });
  const cards = [...cmRail.querySelectorAll('.cmt-card')];
  const typing = cmRail.contains(document.activeElement);
  if (!typing && (cards.some(c => !firsts.has(c.dataset.cid)) || state.cur.comments.some(c => firsts.has(c.id) && !c.done && !cards.some(x => x.dataset.cid === c.id)))) { renderComments(); return; }
  const scroll = $('#wScroll'), compact = cards.length ? layoutComments(null, firsts) : false;
  scroll.classList.toggle('cmt-compact', compact); scroll.classList.toggle('cmt-open', compact && CM.panel);
}, 160);
function cmSaveEdit(card) {
  const c = cmFind(card.dataset.cid); if (!c) return;
  const ta = card.querySelector('[data-ci="main"]'); const t = ta ? ta.value.trim() : '';
  if (!t && !c.text) { deleteComment(c.id); return; }
  if (t) { c.text = t; c.at = c.at || Date.now(); }
  CM.edit = null; scheduleSave(); renderComments();
}
function cmReply(card, ta) {
  const c = cmFind(card.dataset.cid); const t = ta.value.trim(); if (!c || !t) return;
  const a = cmAuthor(); (c.replies ||= []).push({ id: 'r' + uid(), author: a.name, ini: a.ini, at: Date.now(), text: t });
  scheduleSave(); renderComments();
  setTimeout(() => { const n = cmRail.querySelector(`[data-cid="${c.id}"] [data-ci="reply"]`); if (n) n.focus(); }, 20);
}
cmRail.addEventListener('click', e => {
  const b = e.target.closest('[data-ca]'), card = e.target.closest('.cmt-card');
  const a = b && b.dataset.ca;
  if (a === 'close') { CM.panel = false; CM.active = null; renderComments(); return; }
  if (a === 'showdone') { CM.showDone = !CM.showDone; renderComments(); return; }
  if (!card) return;
  const id = card.dataset.cid, c = cmFind(id); if (!c) return;
  if (a === 'del') { deleteComment(id); toast('Commentaire supprimé'); return; }
  if (a === 'done') { c.done = !c.done; if (c.done && CM.active === id) CM.active = null; scheduleSave(); renderComments(); toast(c.done ? 'Commentaire résolu' : 'Commentaire rouvert'); return; }
  if (a === 'rdel') { c.replies.splice(+b.dataset.k, 1); scheduleSave(); renderComments(); return; }
  if (a === 'save') { cmSaveEdit(card); return; }
  if (a === 'cancel') { if (!c.text) deleteComment(id); else { CM.edit = null; renderComments(); } return; }
  if (a === 'edit' && CM.active === id && !c.done) { CM.edit = id; renderComments(); const ta = cmRail.querySelector(`[data-cid="${id}"] textarea`); if (ta) { ta.focus(); ta.selectionStart = ta.value.length; } return; }
  if (e.target.closest('textarea')) return;
  if (CM.active !== id) { CM.active = id; renderComments(); const an = cmAnchors(id)[0]; if (an) an.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); }
});
cmRail.addEventListener('keydown', e => {
  const ta = e.target.closest('textarea'); if (!ta) return;
  const card = ta.closest('.cmt-card');
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); ta.dataset.ci === 'main' ? cmSaveEdit(card) : cmReply(card, ta); }
  else if (e.key === 'Escape') { e.preventDefault(); const c = cmFind(card.dataset.cid); if (ta.dataset.ci === 'main' && c && !c.text) deleteComment(c.id); else { CM.edit = null; renderComments(); } ed.focus(); }
  e.stopPropagation();
});
cmRail.addEventListener('input', e => { const ta = e.target.closest('textarea'); if (ta) { ta.style.height = 'auto'; ta.style.height = ta.scrollHeight + 'px'; relayoutComments(); } });
ed.addEventListener('click', e => {
  const sp = e.target.closest('[data-c]');
  const id = sp ? sp.dataset.c : null;
  if (id === CM.active) return;
  if (!id && !CM.active) return;
  if (CM.edit && CM.edit !== id) { const c = cmFind(CM.edit); if (c && !c.text) { deleteComment(CM.edit); return; } CM.edit = null; }
  CM.active = id; if (id) CM.panel = true; renderComments();
});
ed.addEventListener('input', relayoutComments);
$('#wStatus').addEventListener('click', e => { if (e.target.closest('#wCmtBtn')) { CM.panel = !CM.panel || !$('#wScroll').classList.contains('cmt-compact'); if (!$('#wScroll').classList.contains('cmt-compact')) { const c = (state.cur.comments || []).find(x => !x.done); if (c) { CM.active = c.id; const an = cmAnchors(c.id)[0]; if (an) an.scrollIntoView({ block: 'center', behavior: 'smooth' }); } } renderComments(); } });
window.addEventListener('resize', relayoutComments);
if (window.ResizeObserver) new ResizeObserver(relayoutComments).observe(ed);
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.altKey && (e.key === 'm' || e.key === 'M' || e.code === 'KeyM') && state.view === 'writer') { e.preventDefault(); addComment(); }
});

/* =========================================================
   Styles du document (Titre 1, Titre 2, Titre 3, Paragraphe, Citation) et colonnes
   ========================================================= */
const STY_TAGS = [['h1', 'Titre 1'], ['h2', 'Titre 2'], ['h3', 'Titre 3'], ['p', 'Paragraphe'], ['blockquote', 'Citation']];
const STY = { cur: 'h1' };
function styleCss(st) {
  return Object.entries(st || {}).map(([t, o]) => {
    const d = [], blk = [];
    if (o.ff && FONT_STACK[o.ff]) d.push(`font-family:${FONT_STACK[o.ff]}`);
    if (o.fs) d.push(`font-size:${o.fs}pt`);
    if (o.color) d.push(`color:${o.color}`);
    if (o.b != null) d.push(`font-weight:${o.b ? 700 : 400}`);
    if (o.i != null) d.push(`font-style:${o.i ? 'italic' : 'normal'}`);
    if (o.al) blk.push(`text-align:${o.al}`);
    if (o.sb != null) blk.push(`margin-top:${o.sb}pt`);
    if (o.sa != null) blk.push(`margin-bottom:${o.sa}pt`);
    if (t === 'p') return (d.length ? `#ed.page{${d.join(';')}}` : '') + (blk.length ? `#ed.page p{${blk.join(';')}}` : '');
    return d.length || blk.length ? `#ed.page ${t}{${d.concat(blk).join(';')}}` : '';
  }).join('\n');
}
function applyDocStyles() { $('#docStyles').textContent = state.view === 'writer' && state.cur ? styleCss(state.cur.styles) : ''; }
function fontKeyOf(cssFam) { const f = String(cssFam || '').split(',')[0].replace(/["']/g, '').trim(); return Object.keys(FONT_STACK).find(k => k.toLowerCase() === f.toLowerCase()) || ''; }
function openStylePop() {
  const p = $('#stylePop'), st = (state.cur.styles ||= {}), o = st[STY.cur] || {};
  const seg = (k, opts, v) => `<div class="seg" data-sk="${k}">` + opts.map(([x, l]) => `<button type="button" data-v="${x}" aria-pressed="${v === x}">${l}</button>`).join('') + '</div>';
  const tri = (k, l) => `<button type="button" class="sty-tg" data-st="${k}" aria-pressed="${o[k] === true}" title="${l}">${l}</button>`;
  p.innerHTML = `<div class="pp-h"><b>Styles du document</b><button type="button" class="xbtn" data-pclose aria-label="Fermer">${ic('close')}</button></div>
    <div class="sty-tabs">${STY_TAGS.map(([t, l]) => `<button type="button" data-stab="${t}" aria-pressed="${STY.cur === t}">${l}</button>`).join('')}</div>
    <div class="sty-grid">
      <label class="field"><span>Police</span><select data-sf="ff"><option value="">Par défaut</option>${Object.keys(FONT_STACK).map(f => `<option${o.ff === f ? ' selected' : ''}>${f}</option>`).join('')}</select></label>
      <label class="field"><span>Taille (pt)</span><input type="number" min="6" max="96" step="0.5" data-sf="fs" value="${o.fs || ''}" placeholder="Par défaut"></label>
      <label class="field"><span>Couleur</span><span class="sty-col"><input type="color" data-sf="color" value="${o.color || '#1E1814'}"><button type="button" class="lnk" data-sreset="color"${o.color ? '' : ' hidden'}>Par défaut</button></span></label>
      <label class="field"><span>Espace avant / après (pt)</span><span class="sty-sp"><input type="number" min="0" max="72" data-sf="sb" value="${o.sb ?? ''}" placeholder="—"><input type="number" min="0" max="72" data-sf="sa" value="${o.sa ?? ''}" placeholder="—"></span></label>
    </div>
    <div class="pp-row"><span>Forme</span><div class="sty-tgs">${tri('b', 'Gras')}${tri('i', 'Italique')}</div></div>
    <div class="pp-row"><span>Alignement</span>${seg('al', [['', 'Défaut'], ['left', 'Gauche'], ['center', 'Centre'], ['right', 'Droite'], ['justify', 'Justifié']], o.al || '')}</div>
    <div class="sty-acts"><button type="button" class="pbtn ghost" data-sact="pick" title="Reprendre la police, la taille, la couleur et l'alignement du paragraphe où se trouve le curseur">D'après la sélection</button><button type="button" class="lnk" data-sact="reset">Rétablir ce style</button></div>
    <label class="ac-tg"><input type="checkbox" data-sact="def"${ls.get('grim.styles', null) && JSON.stringify(ls.get('grim.styles', null)) === JSON.stringify(st) ? ' checked' : ''}> <span>Utiliser ces styles pour mes nouveaux documents</span></label>
    <p class="set-note">S'appliquent à tout le document, à l'impression, au PDF et à l'export Word.</p>`;
  placeWPop(p, $('#blockSel'));
}
function setStyle(k, v) {
  const st = (state.cur.styles ||= {}), o = (st[STY.cur] ||= {});
  if (v === '' || v == null || Number.isNaN(v)) delete o[k]; else o[k] = v;
  if (!Object.keys(o).length) delete st[STY.cur];
  applyDocStyles(); scheduleSave(); refreshSoon(); relayoutComments();
  if (ls.get('grim.styles', null)) { const cb = $('#stylePop [data-sact="def"]'); if (cb && cb.checked) ls.set('grim.styles', clone(st)); }
}
$('#stylePop').addEventListener('click', e => {
  if (e.target.closest('[data-pclose]')) return closeWPops();
  const tab = e.target.closest('[data-stab]'); if (tab) { STY.cur = tab.dataset.stab; openStylePop(); return; }
  const tg = e.target.closest('[data-st]');
  if (tg) { const o = (state.cur.styles || {})[STY.cur] || {}; setStyle(tg.dataset.st, o[tg.dataset.st] === true ? (STY.cur.startsWith('h') && tg.dataset.st === 'b' ? false : null) : true); openStylePop(); return; }
  const sg = e.target.closest('.seg [data-v]'); if (sg) { setStyle(sg.parentElement.dataset.sk, sg.dataset.v || null); openStylePop(); return; }
  const rs = e.target.closest('[data-sreset]'); if (rs) { setStyle(rs.dataset.sreset, null); openStylePop(); return; }
  const a = e.target.closest('[data-sact]'); if (!a) return;
  if (a.dataset.sact === 'reset') { if (state.cur.styles) delete state.cur.styles[STY.cur]; applyDocStyles(); scheduleSave(); openStylePop(); }
  else if (a.dataset.sact === 'pick') {
    restoreRange(); const b = currentBlock(); if (!b) return toast('Placez d\'abord le curseur dans un paragraphe mis en forme.');
    const probe = b.querySelector('span[style],strong,em') || b, cs = getComputedStyle(probe), bs = getComputedStyle(b);
    const tag = b.tagName.toLowerCase(); STY.cur = STY_TAGS.some(([t]) => t === tag) ? tag : 'p';
    const o = {}; const ff = fontKeyOf(cs.fontFamily); if (ff) o.ff = ff;
    o.fs = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
    const c = cssColor(cs.color); if (c) o.color = c;
    o.b = +cs.fontWeight >= 600; o.i = cs.fontStyle === 'italic';
    if (/^(center|right|justify)$/.test(bs.textAlign)) o.al = bs.textAlign;
    (state.cur.styles ||= {})[STY.cur] = o; applyDocStyles(); scheduleSave(); openStylePop();
    toast(STY_TAGS.find(([t]) => t === STY.cur)[1] + ' reprend la mise en forme de ce paragraphe.');
  }
  else if (a.dataset.sact === 'def') { if (a.checked) { ls.set('grim.styles', clone(state.cur.styles || {})); toast('Vos nouveaux documents utiliseront ces styles.'); } else ls.set('grim.styles', null); }
});
$('#stylePop').addEventListener('change', e => {
  const f = e.target.closest('[data-sf]'); if (!f) return;
  const k = f.dataset.sf;
  setStyle(k, k === 'ff' || k === 'color' ? f.value : f.value === '' ? null : Math.max(0, Math.min(96, parseFloat(f.value))));
  if (k === 'color') { const r = $('#stylePop [data-sreset="color"]'); if (r) r.hidden = false; }
});
$('#stylePop').addEventListener('input', e => { const f = e.target.closest('[data-sf="color"]'); if (f) setStyle('color', f.value); });

/* --- colonnes : un passage en 2 ou 3 colonnes (section continue dans Word) */
function topBlock(n) { while (n && n.parentNode !== ed) { if (n === ed) return null; n = n.parentNode; } return n; }
function setColumns(k) {
  if (document.activeElement !== ed) restoreRange();
  const s = getSelection(); if (!s.rangeCount || !ed.contains(s.anchorNode)) return toast('Placez le curseur dans le passage à mettre en colonnes.');
  const r = s.getRangeAt(0);
  let a = topBlock(r.startContainer), b = topBlock(r.endContainer);
  if (!a || !b) return;
  const inCols = a.matches && a.matches('div[data-cols]') ? a : null;
  if (inCols && a === b) {
    if (k === 1) { inCols.replaceWith(...inCols.childNodes); toast('Colonnes retirées'); }
    else { inCols.dataset.cols = String(k); toast(k + ' colonnes'); }
  } else {
    if (k === 1) { [a, b].forEach(x => { if (x && x.matches && x.matches('div[data-cols]')) x.replaceWith(...x.childNodes); }); toast('Colonnes retirées'); }
    else {
      if (a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_PRECEDING) [a, b] = [b, a];
      const div = document.createElement('div'); div.dataset.cols = String(k);
      a.before(div);
      let n = a; while (n) { const nx = n === b ? null : n.nextSibling; if (n.matches && n.matches('div[data-cols]')) div.append(...n.childNodes), n.remove(); else div.appendChild(n); n = nx; }
      if (!div.nextElementSibling) div.after(Object.assign(document.createElement('p'), { innerHTML: '<br>' }));
      toast(k + ' colonnes : le passage s\'écoule d\'une colonne à l\'autre.');
    }
  }
  state.cur.html = edHtml(); scheduleSave(); refreshSoon(); relayoutComments();
}
function openColsPop() {
  const p = $('#wPop');
  p.innerHTML = `<p class="mp-t">Colonnes</p><button type="button" data-cols="1">${ic('doc')}Une colonne</button><button type="button" data-cols="2">${ic('cols2')}Deux colonnes</button><button type="button" data-cols="3">${ic('cols3')}Trois colonnes</button>`;
  placeWPop(p, $('[data-w="cols"]'));
}
$('#wPop').addEventListener('click', e => { const b = e.target.closest('[data-cols]'); if (!b) return; e.stopPropagation(); closeWPops(); setColumns(+b.dataset.cols); }, true);

/* =========================================================
   Suivi des modifications (documents), comme dans Word
   ========================================================= */
const TR = {};
const trOn = () => !!(state.cur && state.cur.track && state.view === 'writer');
const trMine = el => el && el.dataset.a === cmAuthor().name;
const trMark = (tag) => { const el = document.createElement(tag); el.dataset.a = cmAuthor().name; el.dataset.t = String(Date.now()); return el; };
function trTextNodes(r) {
  const root = r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentNode;
  const out = [], tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let n; while ((n = tw.nextNode())) if (n.nodeValue.length && r.intersectsNode(n) && ed.contains(n)) {
    const s = n === r.startContainer ? r.startOffset : 0, e = n === r.endContainer ? r.endOffset : n.nodeValue.length;
    if (s < e) out.push({ n, s, e });
  }
  return out;
}
const trLive = r => trTextNodes(r).some(x => !x.n.parentNode.closest('del[data-a]'));
function trMerge(el) {
  const same = x => x && x.nodeType === 1 && x.tagName === el.tagName && x.dataset.a === el.dataset.a && x.hasAttribute('data-a');
  if (same(el.previousSibling)) { const p = el.previousSibling; p.append(...el.childNodes); el.remove(); el = p; }
  if (same(el.nextSibling)) { const n = el.nextSibling; el.append(...n.childNodes); n.remove(); }
  return el;
}
// Supprime la plage : le texte des autres devient barré, mes propres ajouts disparaissent vraiment.
function trDelete(r, backward) {
  const parts = trTextNodes(r); let first = null, last = null;
  parts.forEach(({ n, s, e }) => {
    if (n.parentNode.closest('del[data-a]')) return;
    if (e < n.nodeValue.length) n.splitText(e);
    if (s > 0) n = n.splitText(s);
    const ins = n.parentNode.closest('ins[data-a]');
    if (ins && trMine(ins)) { const mark = document.createTextNode(''); n.replaceWith(mark); if (!ins.textContent) { ins.replaceWith(mark); } first ||= mark; last = mark; return; }
    const d = trMark('del'); n.replaceWith(d); d.appendChild(n);
    const m = trMerge(d); first ||= m; last = m;
  });
  const sel = getSelection(), c = document.createRange();
  const at = backward ? first : last;
  if (at && at.isConnected) { if (backward) c.setStartBefore(at); else c.setStartAfter(at); } else { c.setStart(r.startContainer, r.startOffset); }
  c.collapse(true); sel.removeAllRanges(); sel.addRange(c);
}
function trInsert(text) {
  const sel = getSelection(); if (!sel.rangeCount) return;
  let r = sel.getRangeAt(0);
  if (!r.collapsed) { trDelete(r, false); r = sel.getRangeAt(0); }
  const host = n => n.nodeType === 3 ? n.parentNode : n;
  const del = host(r.startContainer).closest('del[data-a]'); if (del && ed.contains(del)) { r = document.createRange(); r.setStartAfter(del); r.collapse(true); }
  const before = r.startContainer.nodeType === 3 ? (r.startOffset === 0 ? r.startContainer.previousSibling : null) : r.startContainer.childNodes[r.startOffset - 1];
  let ins = host(r.startContainer).closest('ins[data-a]');
  if (!(ins && trMine(ins) && ed.contains(ins)) && before && before.nodeType === 1 && before.matches('ins[data-a]') && trMine(before)) { ins = before; r = document.createRange(); r.selectNodeContents(ins); r.collapse(false); }
  const t = document.createTextNode(text);
  if (ins && trMine(ins) && ed.contains(ins)) r.insertNode(t);
  else { const el = trMark('ins'); el.appendChild(t); r.insertNode(el); trMerge(el); }
  const c = document.createRange(); c.setStart(t, t.nodeValue.length); c.collapse(true); sel.removeAllRanges(); sel.addRange(c);
}
function trAfter() { state.cur.html = edHtml(); scheduleSave(); refreshSoon(); relayoutComments(); }
ed.addEventListener('beforeinput', e => {
  if (!trOn() || e.isComposing) return;
  const t = e.inputType, sel = getSelection(); if (!sel.rangeCount) return;
  if (t === 'insertText' || t === 'insertReplacementText') {
    const txt = e.data ?? (e.dataTransfer && e.dataTransfer.getData('text/plain')); if (!txt) return;
    e.preventDefault(); trInsert(txt); trAfter();
  } else if (/^delete(Content|Word|SoftLine|HardLine)(Backward|Forward)$/.test(t) || t === 'deleteByCut') {
    const back = /Backward$/.test(t);
    if (sel.isCollapsed) {
      const a = [sel.anchorNode, sel.anchorOffset], unit = /Word/.test(t) ? 'word' : /Line/.test(t) ? 'lineboundary' : 'character';
      // on passe par-dessus le texte déjà barré
      for (let i = 0; i < 400; i++) { const before = sel.toString().length; sel.modify('extend', back ? 'backward' : 'forward', unit); if (trLive(sel.getRangeAt(0)) || sel.toString().length === before) break; }
      if (!trLive(sel.getRangeAt(0))) { sel.collapse(a[0], a[1]); if (!sel.toString()) return; }
    }
    const r = sel.getRangeAt(0);
    if (!trLive(r)) return; // fusion de paragraphes : laissée au navigateur
    e.preventDefault(); trDelete(r, back); trAfter();
  }
});
function trAll(accept, root = ed) {
  root.querySelectorAll('ins[data-a],del[data-a]').forEach(el => { if ((el.tagName === 'INS') === accept) el.replaceWith(...el.childNodes); else el.remove(); });
  ed.normalize(); trAfter(); renderTrackBar();
}
function trCount() { return ed.querySelectorAll('ins[data-a],del[data-a]').length; }
function toggleTrack() {
  state.cur.track = !state.cur.track; if (!state.cur.track) delete state.cur.track;
  scheduleSave(); renderTrackBar(); updateToolState();
  toast(state.cur.track ? 'Suivi des modifications activé : ajouts soulignés, suppressions barrées.' : 'Suivi des modifications désactivé.');
}
function renderTrackBar() {
  const b = $('#trackBar'); if (!b) return;
  const n = state.view === 'writer' && state.cur ? trCount() : 0;
  b.hidden = !(trOn() || n);
  const tb = $('[data-w="track"]'); if (tb) tb.setAttribute('aria-pressed', String(trOn()));
  if (b.hidden) return;
  b.innerHTML = `<span class="trk-st${trOn() ? ' on' : ''}">${ic('pen')}${trOn() ? 'Suivi des modifications activé' : 'Suivi désactivé'}</span><span class="trk-n">${n ? `${n} modification${n > 1 ? 's' : ''}` : 'Aucune modification'}</span>`
    + (n ? `<button type="button" class="lnk" data-tr="prev">Précédente</button><button type="button" class="lnk" data-tr="next">Suivante</button><button type="button" class="pbtn" data-tr="acc">Tout accepter</button><button type="button" class="pbtn ghost" data-tr="ref">Tout refuser</button>` : '')
    + `<button type="button" class="lnk" data-tr="toggle">${trOn() ? 'Désactiver' : 'Activer'}</button>`;
}
let trIdx = -1;
$('#trackBar').addEventListener('click', e => {
  const b = e.target.closest('[data-tr]'); if (!b) return;
  const k = b.dataset.tr;
  if (k === 'toggle') return toggleTrack();
  if (k === 'acc' || k === 'ref') return trAll(k === 'acc');
  const all = [...ed.querySelectorAll('ins[data-a],del[data-a]')]; if (!all.length) return;
  trIdx = (trIdx + (k === 'next' ? 1 : -1) + all.length) % all.length;
  const el = all[trIdx]; el.scrollIntoView({ block: 'center', behavior: 'smooth' }); trPop(el);
});
function trPop(el) {
  const p = $('#wPop'); p.dataset.k = 'trk';
  const who = esc(el.dataset.a || 'Anonyme'), when = relDate(+el.dataset.t);
  p.innerHTML = `<p class="mp-t">${el.tagName === 'INS' ? 'Ajout' : 'Suppression'} · ${who} · ${when}</p><button type="button" data-tra="1">${ic('tick')}Accepter</button><button type="button" data-tra="0">${ic('close')}Refuser</button>`;
  TR.el = el; placeWPop(p, el);
}
$('#wPop').addEventListener('click', e => {
  const b = e.target.closest('[data-tra]'); if (!b || !TR.el) return; e.stopImmediatePropagation(); closeWPops();
  const el = TR.el, acc = b.dataset.tra === '1'; TR.el = null;
  if (!el.isConnected) return;
  if ((el.tagName === 'INS') === acc) el.replaceWith(...el.childNodes); else el.remove();
  ed.normalize(); trAfter(); renderTrackBar();
}, true);
ed.addEventListener('click', e => { const m = e.target.closest('ins[data-a],del[data-a]'); if (m && !e.target.closest('[data-c]')) trPop(m); });
ed.addEventListener('input', debounce(renderTrackBar, 300));
document.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.shiftKey && (e.key === 'E' || e.key === 'e') && state.view === 'writer') { e.preventDefault(); e.stopImmediatePropagation(); toggleTrack(); } }, true);

/* --- Enregistrer / Enregistrer sous : un document Grimoire lié à un vrai fichier de l'ordinateur */
const FH = {
  db: null, links: new Map(),
  open() { return this.db ||= new Promise((res, rej) => { const r = indexedDB.open('grimoire-fichiers', 1); r.onupgradeneeded = () => r.result.createObjectStore('h'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
  async tx(mode, fn) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('h', mode); const q = fn(t.objectStore('h')); t.oncomplete = () => res(q && q.result); t.onerror = () => rej(t.error); }); },
  async get(id) { if (this.links.has(id)) return this.links.get(id); let h = null; try { h = await this.tx('readonly', st => st.get(id)) || null; } catch {} this.links.set(id, h); return h; },
  async set(id, h) { this.links.set(id, h); try { await this.tx('readwrite', st => st.put(h, id)); } catch {} },
  async del(id) { this.links.delete(id); this.dirty(id, false); try { await this.tx('readwrite', st => st.delete(id)); } catch {} },
  dirty(id, v) { const m = ls.get('grim.fdirty', {}); if (v === undefined) return !!m[id]; if (v) m[id] = 1; else delete m[id]; ls.set('grim.fdirty', m); }
};
const SAVE_FMT = {
  texte: [{ ext: 'docx', k: 'docx', d: 'Document Word' }],
  feuille: [{ ext: 'xlsx', k: 'xlsx', d: 'Classeur Excel' }, { ext: 'xlsm', k: 'xlsm', d: 'Classeur Excel avec macros' }, { ext: 'csv', k: 'csv', d: 'CSV (point-virgule)' }],
  pdf: [{ ext: 'pdf', k: 'pdf', d: 'PDF' }]
};
const fileExt = n => (String(n).split('.').pop() || '').toLowerCase();
const kindFor = (type, name) => { const e = fileExt(name), f = (SAVE_FMT[type] || []).find(x => x.ext === e); return f ? f.k : type === 'texte' && (e === 'md' || e === 'html') ? e : null; };
async function produce(k) {
  let out = null; CAPTURE = o => { out = o; };
  try { await runExport(k); } finally { CAPTURE = null; }
  if (!out) throw new Error('Fichier non créé');
  return new Blob([out.data], { type: MIME[fileExt(out.filename)] || 'application/octet-stream' });
}
function updFileSt() {
  const el = $('#fileSt'), cur = state.cur, h = cur && FH.links.get(cur.id);
  el.hidden = !h; if (!h) return;
  const d = FH.dirty(cur.id);
  el.dataset.dirty = d ? '1' : '';
  el.textContent = h.name;
  el.title = d ? `Modifications pas encore enregistrées dans ${h.name} (Ctrl+S)` : `Enregistré dans ${h.name}`;
}
function markFileDirty() { const c = state.cur; if (c && FH.links.get(c.id) && !FH.dirty(c.id)) { FH.dirty(c.id, true); updFileSt(); } }
async function linkHandle(id, h) { if (h && id) { await FH.set(id, h); FH.dirty(id, false); if (state.cur && state.cur.id === id) updFileSt(); } }
async function saveDoc(asNew) {
  const cur = state.cur; if (!cur) return;
  if (state.view === 'sheet') commitEdit(false);
  await flushSave();
  const fmts = (SAVE_FMT[cur.type] || []).filter(f => f.ext !== 'xlsm' || cur.vba);
  if (cur.vba) fmts.sort((a, b) => (b.ext === 'xlsm') - (a.ext === 'xlsm'));
  let h = asNew ? null : await FH.get(cur.id);
  if (h && cur.vba && fileExt(h.name) === 'xlsx') { toast('Ce classeur contient des macros : choisissez un fichier .xlsm pour les garder.'); h = null; }
  if (h && !kindFor(cur.type, h.name)) { toast(`Grimoire ne sait pas réécrire un fichier .${fileExt(h.name)} : choisissez où l'enregistrer.`); h = null; }
  if (h && h.queryPermission) {
    try { if (await h.queryPermission({ mode: 'readwrite' }) !== 'granted' && await h.requestPermission({ mode: 'readwrite' }) !== 'granted') h = null; }
    catch { h = null; }
  }
  // Premier enregistrement par-dessus un fichier ouvert depuis l'ordinateur : prévenir de ce qui pourrait s'y perdre.
  const imp = ls.get('grim.fimp', {});
  if (h && imp[cur.id]) {
    if (!confirm(`Grimoire va remplacer « ${h.name} » par sa version.\n\nCe que Grimoire ne gère pas dans ce fichier (certaines mises en forme avancées, objets spéciaux) y sera perdu.\n\nOK : remplacer le fichier.\nAnnuler : choisir un autre nom (Enregistrer sous).`)) return saveDoc(true);
    delete imp[cur.id]; ls.set('grim.fimp', imp);
  }
  const fresh = !h;
  if (!h) {
    if (!window.showSaveFilePicker || isPhone()) {
      // Pas d'accès direct aux fichiers (Firefox, téléphone) : le fichier part dans les téléchargements.
      try { await runExport(fmts[0].k); } catch { toast('Enregistrement impossible.', true); }
      return;
    }
    try { h = await window.showSaveFilePicker({ suggestedName: safeName(cur.title) + '.' + fmts[0].ext, types: fmts.map(f => ({ description: f.d, accept: { [MIME[f.ext]]: ['.' + f.ext] } })) }); }
    catch (e) { if (e && e.name !== 'AbortError') toast('Enregistrement impossible : ' + (e.message || e), true); return; }
  }
  const k = kindFor(cur.type, h.name) || fmts[0].k;
  if (fresh) {
    // comme Word : le document prend le nom du fichier choisi
    const t = h.name.replace(/\.[^.]+$/, '');
    if (t && t !== cur.title) { cur.title = t; delete cur.autoTitle; $('#titleIn').value = t; await Store.save(cur).catch(() => {}); }
  }
  try {
    const blob = await produce(k);
    const w = await h.createWritable(); await w.write(blob); await w.close();
  } catch (e) { console.warn(e); toast('Le fichier ' + h.name + ' n\'a pas pu être écrit' + (e && e.name === 'NoModificationAllowedError' ? ' : il est peut-être ouvert ailleurs.' : '.'), true); return; }
  await linkHandle(cur.id, h);
  toast('Enregistré dans ' + h.name);
}
async function openFromComputer() {
  if (!window.showOpenFilePicker || isPhone()) { $('#fileIn').click(); return; }
  let hs;
  try { hs = await window.showOpenFilePicker({ multiple: true, types: [{ description: 'Word, Excel, CSV, PDF, texte', accept: { 'application/octet-stream': ['.docx', '.xlsx', '.xlsm', '.xls', '.ods', '.csv', '.txt', '.md', '.pdf'] } }] }); }
  catch (e) { if (e && e.name !== 'AbortError') $('#fileIn').click(); return; }
  await importHandles(hs);
}
async function importHandles(hs) {
  const items = []; for (const h of hs) { try { items.push({ h, f: await h.getFile() }); } catch {} }
  const pdfs = items.filter(x => fileExt(x.f.name) === 'pdf');
  if (pdfs.length > 1) await importPdfFiles(pdfs.map(x => x.f));
  for (const x of items) {
    if (pdfs.length > 1 && fileExt(x.f.name) === 'pdf') continue;
    const prev = state.cur && state.cur.id;
    await importFile(x.f);
    const c = state.cur;
    if (c && c.id !== prev && kindFor(c.type, x.h.name)) { await linkHandle(c.id, x.h); const imp = ls.get('grim.fimp', {}); imp[c.id] = 1; ls.set('grim.fimp', imp); }
  }
}

const safeName = t => (String(t || 'Sans titre').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() || 'Sans titre').slice(0, 120);

const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
// ---- Export Word : texte riche, liens, listes numérotées, tableaux, mise en page
let LINKS = [], NUMS = [], TW = 9070;
const DOCX_FONT = { 'Source Serif 4': 'Georgia', 'Instrument Sans': 'Calibri' };
function spanFmt(el, fmt) {
  const f = { ...fmt }, st = el.style;
  const tg = el.tagName;
  if (/^(STRONG|B|TH)$/.test(tg)) f.b = true;
  if (/^(EM|I)$/.test(tg)) f.i = true;
  if (tg === 'U') f.u = true;
  if (tg === 'S' || tg === 'STRIKE' || tg === 'DEL') f.s = true;
  if (tg === 'SUP') f.sup = true;
  if (tg === 'SUB') f.sub = true;
  if (st) {
    const c = cssColor(st.color); if (c) f.color = c.slice(1);
    const bg = cssColor(st.backgroundColor); if (bg) f.bg = bg.slice(1);
    const ff = cssFont(st.fontFamily); if (ff) f.font = DOCX_FONT[ff] || ff;
    const fs = cssSizePt(st.fontSize); if (fs) f.size = fs;
  }
  return f;
}
function rPr(f) {
  let x = '';
  if (f.link) x += '<w:rStyle w:val="Hyperlink"/>';
  if (f.font) x += `<w:rFonts w:ascii="${xmlEsc(f.font)}" w:hAnsi="${xmlEsc(f.font)}" w:cs="${xmlEsc(f.font)}"/>`;
  if (f.b) x += '<w:b/>';
  if (f.i) x += '<w:i/>';
  if (f.s) x += '<w:strike/>';
  if (f.color) x += `<w:color w:val="${f.color}"/>`;
  if (f.size) x += `<w:sz w:val="${Math.round(f.size * 2)}"/><w:szCs w:val="${Math.round(f.size * 2)}"/>`;
  if (f.u) x += '<w:u w:val="single"/>';
  if (f.bg) x += `<w:shd w:val="clear" w:color="auto" w:fill="${f.bg}"/>`;
  if (f.sup) x += '<w:vertAlign w:val="superscript"/>'; else if (f.sub) x += '<w:vertAlign w:val="subscript"/>';
  return x ? `<w:rPr>${x}</w:rPr>` : '';
}
let CMX = { ids: new Map(), first: new Map(), last: new Map() }, TRK_ID = 1000;
const wDate = t => new Date(t || Date.now()).toISOString().replace(/\.\d+Z$/, 'Z');
const paraId = () => (1 + Math.floor(Math.random() * 0x7FFFFFFE)).toString(16).toUpperCase().padStart(8, '0');
function commentsXml(list) {
  const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  let cx = '', ex = '';
  list.forEach(c => {
    const ws = CMX.ids.get(c.id);
    [c, ...(c.replies || [])].forEach((it, k) => {
      const pid = paraId(), w = ws[k];
      const runs = String(it.text || '').split('\n').map((l, i) => (i ? '<w:r><w:br/></w:r>' : '') + `<w:r><w:t xml:space="preserve">${xmlEsc(l)}</w:t></w:r>`).join('');
      cx += `<w:comment w:id="${w}" w:author="${xmlEsc(it.author || 'Anonyme')}" w:date="${wDate(it.at)}" w:initials="${xmlEsc(it.ini || '')}"><w:p w14:paraId="${pid}" w14:textId="77777777"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:annotationRef/></w:r>${runs}</w:p></w:comment>`;
      if (!k) it._pid = pid;
      ex += `<w15:commentEx w15:paraId="${pid}"${k ? ` w15:paraIdParent="${c._pid}"` : ''} w15:done="${c.done ? 1 : 0}"/>`;
    });
    delete c._pid;
  });
  return {
    comments: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="${W14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">${cx}</w:comments>`,
    ext: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w15">${ex}</w15:commentsEx>`
  };
}
function runsOf(node, fmt = {}) {
  let out = '';
  node.childNodes.forEach(n => {
    if (n.nodeType === 3) {
      const t = n.nodeValue.replace(/\s+/g, ' ');
      if (!t) return;
      const tt = fmt.del ? 'delText' : 't';
      out += `<w:r>${rPr(fmt)}<w:${tt} xml:space="preserve">${xmlEsc(t)}</w:${tt}></w:r>`;
    } else if (n.nodeType === 1) {
      const tg = n.tagName;
      if (tg === 'BR') out += '<w:r><w:br/></w:r>';
      else if (tg === 'IMG') out += imgRun(n);
      else if (/^(UL|OL|TABLE)$/.test(tg)) return;
      else if ((tg === 'INS' || tg === 'DEL') && n.hasAttribute('data-a') && !fmt.ins && !fmt.del) {
        const k = tg === 'INS' ? 'ins' : 'del';
        out += `<w:${k} w:id="${++TRK_ID}" w:author="${xmlEsc(n.getAttribute('data-a'))}" w:date="${wDate(+n.getAttribute('data-t') || Date.now())}">${runsOf(n, { ...fmt, [k]: true })}</w:${k}>`;
      }
      else if (tg === 'SPAN' && n.hasAttribute('data-c') && CMX.ids.has(n.getAttribute('data-c'))) {
        const cid = n.getAttribute('data-c'), ws = CMX.ids.get(cid);
        if (CMX.first.get(cid) === n) out += ws.map(w => `<w:commentRangeStart w:id="${w}"/>`).join('');
        out += runsOf(n, fmt);
        if (CMX.last.get(cid) === n) out += ws.map(w => `<w:commentRangeEnd w:id="${w}"/><w:r><w:commentReference w:id="${w}"/></w:r>`).join('');
      }
      else if (tg === 'A' && !fmt.link) {
        LINKS.push(n.getAttribute('href') || '');
        out += `<w:hyperlink r:id="rIdL${LINKS.length}" w:history="1">${runsOf(n, { ...fmt, link: true })}</w:hyperlink>`;
      } else out += runsOf(n, spanFmt(n, fmt));
    }
  });
  return out;
}
function pPr({ style, numPr, tabs, el, ind, jc } = {}) {
  let x = '';
  if (style) x += `<w:pStyle w:val="${style}"/>`;
  if (numPr) x += numPr;
  if (tabs) x += tabs;
  const st = el && el.style;
  const lh = st && parseFloat(st.lineHeight), mb = st && st.marginBottom, ml = st && parseInt(st.marginLeft, 10);
  let sp = '';
  if (mb === '0px') sp += ' w:after="0"'; else if (mb === '1.5em') sp += ' w:after="360"';
  if (lh) sp += ` w:line="${Math.round(lh * 240)}" w:lineRule="auto"`;
  if (sp) x += `<w:spacing${sp}/>`;
  if (ind) x += ind; else if (ml) x += `<w:ind w:left="${ml * 15}"/>`;
  const al = (st && st.textAlign) || jc;
  if (al) x += `<w:jc w:val="${{ left: 'left', center: 'center', right: 'right', justify: 'both' }[al] || 'left'}"/>`;
  return x ? `<w:pPr>${x}</w:pPr>` : '';
}
const para = (style, runs, el, o = {}) => `<w:p>${pPr({ style, el, ...o })}${runs}</w:p>`;
function listXml(list, level) {
  let out = '';
  const check = list.tagName === 'UL' && list.hasAttribute('data-check');
  let numId = 1;
  if (list.tagName === 'OL') { NUMS.push(Math.min(level, 8)); numId = NUMS.length + 1; }
  [...list.children].forEach(li => {
    if (li.tagName !== 'LI') return;
    if (check) {
      const box = `<w:r><w:t xml:space="preserve">${li.getAttribute('data-done') === '1' ? '☒' : '☐'} </w:t></w:r>`;
      out += para('ListParagraph', box + runsOf(li), li, { ind: `<w:ind w:left="${360 + level * 360}"/>` });
    } else {
      out += para('ListParagraph', runsOf(li), li, { numPr: `<w:numPr><w:ilvl w:val="${Math.min(level, 8)}"/><w:numId w:val="${numId}"/></w:numPr>` });
    }
    [...li.children].filter(c => /^(UL|OL)$/.test(c.tagName)).forEach(sub => { out += listXml(sub, level + 1); });
  });
  return out;
}
function numberingXml() {
  const BUL = ['•', '◦', '▪'], FMT = ['decimal', 'lowerLetter', 'lowerRoman'];
  const lv = (l, bullet) => `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${bullet ? 'bullet' : FMT[l % 3]}"/><w:lvlText w:val="${bullet ? BUL[l % 3] : '%' + (l + 1) + '.'}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 + l * 360}" w:hanging="360"/></w:pPr>${bullet ? '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' : ''}</w:lvl>`;
  const abs = (id, bullet) => `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, l) => lv(l, bullet)).join('')}</w:abstractNum>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W_NS}>${abs(0, true)}${abs(1, false)}<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`
    + NUMS.map((l, k) => `<w:num w:numId="${k + 2}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('') + '</w:numbering>';
}
function tableXml(tb) {
  const rows = [...tb.rows];
  const n = gridCols(tb);
  const cols = [...tb.querySelectorAll('col')];
  const widths = cols.length === n ? cols.map(c => Math.round((parseFloat(c.style.width) || 100 / n) / 100 * TW)) : Array(n).fill(Math.floor(TW / n));
  const b = s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="C5B8A8"/>`;
  let out = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(b).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${widths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  rows.forEach(tr => {
    const head = tr.parentElement && tr.parentElement.tagName === 'THEAD';
    out += `<w:tr>${head ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}`;
    const g = rowGrid(tr);
    let used = 0;
    g.forEach(o => {
      const span = Math.min(o.span, n - o.start); if (span <= 0) return;
      const w = widths.slice(o.start, o.start + span).reduce((s, x) => s + x, 0);
      const th = o.cell.tagName === 'TH';
      const bg = cssColor(o.cell.style.backgroundColor);
      const fill = bg ? bg.slice(1) : th ? 'EAD9B8' : '';
      const inner = [...o.cell.children].some(c => /^(P|UL|OL|H1|H2|H3)$/.test(c.tagName))
        ? bodyXml(o.cell, { b: th })
        : para('TableText', runsOf(o.cell, th ? { b: true } : {}), o.cell);
      out += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}</w:tcPr>${inner || '<w:p/>'}</w:tc>`;
      used = o.start + span;
    });
    if (used < n) { const w = widths.slice(used).reduce((s, x) => s + x, 0); out += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${n - used > 1 ? `<w:gridSpan w:val="${n - used}"/>` : ''}</w:tcPr><w:p/></w:tc>`; }
    out += '</w:tr>';
  });
  return out + '</w:tbl>';
}
function tocXml(root) {
  const hs = [...root.querySelectorAll('h1,h2,h3')].filter(h => h.textContent.trim() && !h.closest('[data-toc]'));
  const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  let out = para('TOCHeading', '<w:r><w:t>Sommaire</w:t></w:r>');
  if (!hs.length) return out + `<w:p>${begin}<w:r><w:t>Aucun titre</w:t></w:r>${end}</w:p>`;
  hs.forEach((h, i) => {
    const t = `<w:r><w:t xml:space="preserve">${xmlEsc(h.textContent.trim())}</w:t></w:r>`;
    out += `<w:p><w:pPr><w:pStyle w:val="TOC${h.tagName[1]}"/></w:pPr>${i === 0 ? begin : ''}${t}${i === hs.length - 1 ? end : ''}</w:p>`;
  });
  return out;
}
function bodyXml(root, cellFmt) {
  let out = '', inline = null;
  const flush = () => { if (inline) { const lone = !inline.textContent.trim() && inline.querySelector('img:not([data-wrap])'); out += para(cellFmt ? 'TableText' : '', runsOf(inline, cellFmt || {}), null, lone ? { jc: 'center' } : {}); inline = null; } };
  [...root.childNodes].forEach(n => {
    if (n.nodeType === 3) { if (n.nodeValue.trim()) { inline ||= document.createElement('p'); inline.appendChild(n.cloneNode()); } return; }
    if (n.nodeType !== 1) return;
    const tg = n.tagName;
    if (/^(STRONG|EM|U|S|SUP|SUB|A|B|I|SPAN|BR|IMG)$/.test(tg)) { inline ||= document.createElement('p'); inline.appendChild(n.cloneNode(true)); return; }
    flush();
    if (tg === 'H1') out += para('Heading1', runsOf(n), n);
    else if (tg === 'H2') out += para('Heading2', runsOf(n), n);
    else if (tg === 'H3') out += para('Heading3', runsOf(n), n);
    else if (tg === 'BLOCKQUOTE') { if (n.querySelector('p')) [...n.children].forEach(c => { out += para('Quote', runsOf(c), c); }); else out += para('Quote', runsOf(n), n); }
    else if (tg === 'UL' || tg === 'OL') out += listXml(n, 0);
    else if (tg === 'TABLE') out += tableXml(n) + para('', '');
    else if (tg === 'HR' && n.hasAttribute('data-pb')) out += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    else if (tg === 'HR') out += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="C9804F"/></w:pBdr></w:pPr></w:p>';
    else if (tg === 'DIV' && n.hasAttribute('data-toc')) out += tocXml(root);
    else if (tg === 'DIV' && n.hasAttribute('data-cols')) {
      const sp = cols => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:sectPr><w:type w:val="continuous"/>${SECT_PAGE}<w:cols w:num="${cols}" w:space="567"/></w:sectPr></w:pPr></w:p>`;
      out += sp(1) + bodyXml(n) + sp(+n.getAttribute('data-cols') || 2); SECT_COLS = true;
    }
    else if (tg === 'P' && !n.textContent.trim() && n.querySelector('img:not([data-wrap])') && !n.style.textAlign) out += para('', runsOf(n), n, { jc: 'center' });
    else out += para(cellFmt ? 'TableText' : '', runsOf(n, cellFmt || {}), n);
  });
  flush();
  return out || '<w:p/>';
}
let SECT_PAGE = '', SECT_COLS = false;
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const D_NS = W_NS + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
let MEDIA = [];
function imgRun(n) {
  const m = MEDIA[+n.getAttribute('data-m')]; if (!m) return '';
  const wrap = n.getAttribute('data-wrap');
  const graphic = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${m.id}" name="${m.name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${m.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${m.cx}" cy="${m.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  if (wrap === 'left' || wrap === 'right') {
    const gap = 228600; // 0,63 cm entre l'image et le texte
    return `<w:r><w:drawing><wp:anchor distT="45720" distB="91440" distL="${wrap === 'right' ? gap : 0}" distR="${wrap === 'left' ? gap : 0}" simplePos="0" relativeHeight="${251658240 + m.id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>${wrap}</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${m.cx}" cy="${m.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${m.id}" name="Image ${m.id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>${graphic}</wp:anchor></w:drawing></w:r>`;
  }
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${m.cx}" cy="${m.cy}"/><wp:docPr id="${m.id}" name="Image ${m.id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${m.id}" name="${m.name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${m.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${m.cx}" cy="${m.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="140" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="200"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:color w:val="1B1512"/><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="2" w:color="DCD3C8"/></w:pBdr><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="1B1512"/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="A35E2C"/><w:sz w:val="23"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="C9804F"/></w:pBdr><w:ind w:left="340"/></w:pPr><w:rPr><w:i/><w:color w:val="6B5F56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="Table Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:color w:val="6B5F56"/><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:color w:val="6B5F56"/><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/><w:ind w:left="240"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC3"><w:name w:val="toc 3"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/><w:ind w:left="480"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="1F5FAD"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>`;
// Styles Word construits d'après les styles du document
const WSTY_BASE = {
  Normal: { name: 'Normal', p: {}, r: {} },
  Heading1: { name: 'heading 1', tag: 'h1', p: { sb: 12, sa: 10, keep: 1, lvl: 0 }, r: { font: 'Georgia', b: true, color: '1B1512', fs: 20 } },
  Heading2: { name: 'heading 2', tag: 'h2', p: { sb: 18, sa: 6, keep: 1, lvl: 1, bdr: '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="2" w:color="DCD3C8"/></w:pBdr>' }, r: { font: 'Calibri', b: true, color: '1B1512', fs: 14 } },
  Heading3: { name: 'heading 3', tag: 'h3', p: { sb: 12, sa: 4, keep: 1, lvl: 2 }, r: { font: 'Calibri', b: true, color: 'A35E2C', fs: 11.5 } },
  Quote: { name: 'Quote', tag: 'blockquote', p: { ind: 340, bdr: '<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="C9804F"/></w:pBdr>' }, r: { i: true, color: '6B5F56' } }
};
function stylesXmlFor(st) {
  let x = STYLES; if (!st || !Object.keys(st).length) return x;
  Object.entries(WSTY_BASE).forEach(([id, B]) => {
    const o = (st || {})[B.tag || 'p'] || {};
    if (!Object.keys(o).length) return;
    const P = { ...B.p, ...(o.sb != null ? { sb: o.sb } : {}), ...(o.sa != null ? { sa: o.sa } : {}), ...(o.al ? { al: o.al } : {}) };
    const R = { ...B.r, ...(o.ff ? { font: DOCX_FONT[o.ff] || o.ff } : {}), ...(o.fs ? { fs: o.fs } : {}), ...(o.color ? { color: o.color.replace('#', '') } : {}), ...(o.b != null ? { b: o.b } : {}), ...(o.i != null ? { i: o.i } : {}) };
    const JC = { left: 'left', center: 'center', right: 'right', justify: 'both' };
    const pp = (P.keep ? '<w:keepNext/>' : '') + (P.bdr || '') + (P.sb != null || P.sa != null ? `<w:spacing${P.sb != null ? ` w:before="${Math.round(P.sb * 20)}"` : ''}${P.sa != null ? ` w:after="${Math.round(P.sa * 20)}"` : ''}/>` : '') + (P.ind ? `<w:ind w:left="${P.ind}"/>` : '') + (P.al ? `<w:jc w:val="${JC[P.al]}"/>` : '') + (P.lvl != null ? `<w:outlineLvl w:val="${P.lvl}"/>` : '');
    const rr = (R.font ? `<w:rFonts w:ascii="${xmlEsc(R.font)}" w:hAnsi="${xmlEsc(R.font)}" w:cs="${xmlEsc(R.font)}"/>` : '') + (R.b ? '<w:b/>' : R.b === false ? '<w:b w:val="0"/>' : '') + (R.i ? '<w:i/>' : R.i === false ? '<w:i w:val="0"/>' : '') + (R.color ? `<w:color w:val="${R.color}"/>` : '') + (R.fs ? `<w:sz w:val="${Math.round(R.fs * 2)}"/><w:szCs w:val="${Math.round(R.fs * 2)}"/>` : '');
    const el = `<w:style w:type="paragraph"${id === 'Normal' ? ' w:default="1"' : ''} w:styleId="${id}"><w:name w:val="${B.name}"/>${id === 'Normal' ? '' : '<w:basedOn w:val="Normal"/>' + (B.tag && B.tag[0] === 'h' ? '<w:next w:val="Normal"/>' : '')}<w:qFormat/>${pp ? `<w:pPr>${pp}</w:pPr>` : ''}${rr ? `<w:rPr>${rr}</w:rPr>` : ''}</w:style>`;
    x = x.replace(new RegExp(`<w:style w:type="paragraph"(?: w:default="1")? w:styleId="${id}">[\\s\\S]*?</w:style>`), el);
  });
  return x;
}
const FLD = instr => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const wRun = s => `<w:r><w:t xml:space="preserve">${xmlEsc(s)}</w:t></w:r>`;
async function exportDocx() {
  await loadScript(LIB.jszip);
  const P = pageOf(), land = P.orient === 'landscape';
  const pgW = land ? 16838 : 11906, pgH = land ? 11906 : 16838, mar = MARGIN_TW[P.margin];
  TW = pgW - 2 * mar;
  const root = document.createElement('div'); root.innerHTML = sanitize(state.cur.html);
  const zip = new window.JSZip();
  const add = zip.file.bind(zip); zip.file = (n, d, o) => add(n, d, { createFolders: false, ...(o || {}) });
  MEDIA = []; LINKS = []; NUMS = [];
  const MAXW = TW * 635;
  for (const im of [...root.querySelectorAll('img')]) {
    const mm = /^data:image\/(png|jpeg|gif);base64,(.+)$/.exec(im.getAttribute('src') || '');
    if (!mm) { im.remove(); continue; }
    const el = await loadImg(im.getAttribute('src'));
    const nw = el ? el.naturalWidth : 600, nh = el ? el.naturalHeight : 400;
    const pct = parseFloat(im.style.width);
    let cx = Math.round(nw * 9525), cy = Math.round(nh * 9525);
    if (pct) { cx = Math.round(MAXW * pct / 100); cy = Math.round(cx * nh / nw); }
    else if (cx > MAXW) { cy = Math.round(cy * MAXW / cx); cx = MAXW; }
    const k = MEDIA.length + 1, ext = mm[1];
    MEDIA.push({ id: k, rid: 'rIdImg' + k, name: `image${k}.${ext}`, cx, cy });
    zip.file(`word/media/image${k}.${ext}`, mm[2], { base64: true });
    im.setAttribute('data-m', String(k - 1));
  }
  attachFloats(root);
  TRK_ID = 1000;
  SECT_PAGE = `<w:pgSz w:w="${pgW}" w:h="${pgH}"${land ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${mar}" w:right="${mar}" w:bottom="${mar}" w:left="${mar}" w:header="567" w:footer="567" w:gutter="0"/>`; SECT_COLS = false;
  CMX = { ids: new Map(), first: new Map(), last: new Map() };
  root.querySelectorAll('[data-c]').forEach(el => { const id = el.getAttribute('data-c'); if (!CMX.first.has(id)) CMX.first.set(id, el); CMX.last.set(id, el); });
  const cmts = (state.cur.comments || []).filter(c => CMX.first.has(c.id) && (c.text || (c.replies || []).length));
  let wn = 0; cmts.forEach(c => CMX.ids.set(c.id, [c, ...(c.replies || [])].map(() => wn++)));
  const body = bodyXml(root);
  if (cmts.length) { const X = commentsXml(cmts); zip.file('word/comments.xml', X.comments); zip.file('word/commentsExtended.xml', X.ext); }
  CMX = { ids: new Map(), first: new Map(), last: new Map() };
  const hasHdr = !!P.header, hasFtr = !!(P.footer || P.num !== 'none');
  const numRuns = P.num === 'xofy' ? wRun("Page ") + FLD('PAGE') + wRun(" sur ") + FLD('NUMPAGES') : FLD('PAGE');
  const ftrParas = (P.footer ? `<w:p><w:pPr><w:pStyle w:val="Footer"/>${P.num === 'none' ? '' : ''}</w:pPr>${wRun(P.footer)}</w:p>` : '')
    + (P.num !== 'none' ? `<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="${P.num === 'right' ? 'right' : 'center'}"/></w:pPr>${numRuns}</w:p>` : '');
  const ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    hasHdr ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : '',
    hasFtr ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : '',
    cmts.length ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>' : '',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'].join('');
  zip.file('[Content_Types].xml', ct);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId1" Type="${R}styles" Target="styles.xml"/><Relationship Id="rIdNum" Type="${R}numbering" Target="numbering.xml"/>`
    + (hasHdr ? `<Relationship Id="rIdHdr" Type="${R}header" Target="header1.xml"/>` : '')
    + (hasFtr ? `<Relationship Id="rIdFtr" Type="${R}footer" Target="footer1.xml"/>` : '')
    + (cmts.length ? `<Relationship Id="rIdCm" Type="${R}comments" Target="comments.xml"/><Relationship Id="rIdCmx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>` : '')
    + MEDIA.map(m => `<Relationship Id="${m.rid}" Type="${R}image" Target="media/${m.name}"/>`).join('')
    + LINKS.map((u, k) => `<Relationship Id="rIdL${k + 1}" Type="${R}hyperlink" Target="${xmlEsc(u)}" TargetMode="External"/>`).join('')
    + '</Relationships>');
  zip.file('word/styles.xml', stylesXmlFor(state.cur.styles));
  zip.file('word/numbering.xml', numberingXml());
  if (hasHdr) zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${D_NS}><w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>${wRun(P.header)}</w:p></w:hdr>`);
  if (hasFtr) zip.file('word/footer1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${D_NS}>${ftrParas}</w:ftr>`);
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xmlEsc(state.cur.title || '')}</dc:title><dc:creator>CSC</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  const sect = `<w:sectPr>${SECT_COLS ? '<w:type w:val="continuous"/>' : ''}${hasHdr ? '<w:headerReference w:type="default" r:id="rIdHdr"/>' : ''}${hasFtr ? '<w:footerReference w:type="default" r:id="rIdFtr"/>' : ''}<w:pgSz w:w="${pgW}" w:h="${pgH}"${land ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${mar}" w:right="${mar}" w:bottom="${mar}" w:left="${mar}" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>`;
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${D_NS}><w:body>${body}${sect}</w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  await saveFile(safeName(state.cur.title) + '.docx', new Blob([bytes]));
}

function exportHtml() {
  const t = esc(state.cur.title || 'Document');
  const html = `<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>${t}</title><style>body{font-family:Georgia,serif;max-width:720px;margin:48px auto;padding:0 20px;line-height:1.6;color:#1E1814}h1{font-size:2em}h2{border-bottom:1px solid #DCD3C8}h3{color:#A35E2C}table{border-collapse:collapse;width:100%}td,th{border:1px solid #C5B8A8;padding:6px 9px;text-align:left}th{background:#EAD9B8}blockquote{border-left:3px solid #C9804F;margin-left:0;padding-left:1em;color:#6B5F56}</style></head><body>${sanitize(state.cur.html)}</body></html>`;
  return saveFile(safeName(state.cur.title) + '.html', html);
}
function exportMd() {
  const root = document.createElement('div'); root.innerHTML = sanitize(state.cur.html);
  const inl = n => [...n.childNodes].map(c => c.nodeType === 3 ? c.nodeValue : c.nodeType !== 1 ? '' : c.tagName === 'BR' ? '  \n' : /^(STRONG|B)$/.test(c.tagName) ? `**${inl(c)}**` : /^(EM|I)$/.test(c.tagName) ? `*${inl(c)}*` : c.tagName === 'S' ? `~~${inl(c)}~~` : c.tagName === 'A' ? `[${inl(c)}](${c.getAttribute('href')})` : c.tagName === 'IMG' ? '' : /^(UL|OL)$/.test(c.tagName) ? '' : inl(c)).join('');
  const md = [...root.children].map(n => {
    const t = n.tagName;
    if (/^H[1-3]$/.test(t)) return '#'.repeat(+t[1]) + ' ' + inl(n).trim();
    if (t === 'UL' || t === 'OL') return [...n.children].map((li, i) => (t === 'OL' ? `${i + 1}. ` : n.hasAttribute('data-check') ? (li.getAttribute('data-done') === '1' ? '- [x] ' : '- [ ] ') : '- ') + inl(li).trim()).join('\n');
    if (t === 'HR' && n.hasAttribute('data-pb')) return '';
    if (t === 'DIV') return '';
    if (t === 'BLOCKQUOTE') return '> ' + inl(n).trim();
    if (t === 'HR') return '---';
    if (t === 'TABLE') { const rows = [...n.querySelectorAll('tr')].map(tr => '| ' + [...tr.children].map(c => inl(c).trim().replace(/\|/g, '\\|')).join(' | ') + ' |'); if (rows.length) rows.splice(1, 0, '|' + ' --- |'.repeat(n.querySelector('tr').children.length)); return rows.join('\n'); }
    return inl(n).trim();
  }).filter(Boolean).join('\n\n');
  return saveFile(safeName(state.cur.title) + '.md', md + '\n');
}
const XL_RGB = { jaune: ['9A6512', 'FDF0C4'], vert: ['2F7D46', 'DDF0DF'], bleu: ['2563A8', 'DCE9F7'], rouge: ['B23F33', 'F8DEDA'], violet: ['6D4AA8', 'E9E1F5'], orange: ['C2611F', 'FBE3CF'], gris: ['5F6670', 'E7E8EA'], cuivre: ['A35E2C', 'F2E3D4'] };
const XL_Z = { eur: '#,##0.00 "€"', pct: '0.00%', dec2: '#,##0.00', date: 'dd/mm/yyyy' };
/* --- Excel : graphiques, mises en forme conditionnelles, listes et décimales (parties écrites à la main dans le fichier) */
const XL_PAL = ['2A78D6', 'EB6834', '1BAF7A', 'EDA100', 'E87BA4', '008300', '6250D6', 'E34948'];
const xmlE = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
function xlNumFmt(n, d) {
  if (n === 'date') return 'dd/mm/yyyy';
  const dec = d ?? (n === 'dec2' || n === 'eur' || n === 'pct' ? 2 : null);
  const frac = dec ? '.' + '0'.repeat(dec) : '';
  if (n === 'eur') return `#,##0${frac} "€"`;
  if (n === 'pct') return `0${frac}%`;
  if (n === 'dec2') return `#,##0${frac}`;
  if (dec != null) return `#,##0${frac}`;
  return null;
}
// Références d'un graphique : cellule du nom de chaque série, étiquettes et valeurs.
function chartRefs(ch, vals, cellsOf, sheetName) {
  const D = chartData(ch, vals, cellsOf);
  const S = refName(sheetName) + '!';
  const R = (r1, c1, r2, c2) => S + absRef({ r1, c1, r2, c2 });
  const out = { cat: null, series: [] };
  if (!ch.byRow) {
    const r0 = ch.r1 + (D.hasHdr ? 1 : 0);
    if (D.hasLbl) out.cat = R(r0, ch.c1, ch.r2, ch.c1);
    for (let c = ch.c1 + (D.hasLbl ? 1 : 0); c <= ch.c2; c++) out.series.push({ name: D.hasHdr ? R(ch.r1, c, ch.r1, c) : null, val: R(r0, c, ch.r2, c) });
  } else {
    const c0 = ch.c1 + (D.hasHdr ? 1 : 0);
    if (D.hasLbl) out.cat = R(ch.r1, c0, ch.r1, ch.c2);
    for (let r = ch.r1 + (D.hasLbl ? 1 : 0); r <= ch.r2; r++) out.series.push({ name: D.hasHdr ? R(r, ch.c1, r, ch.c1) : null, val: R(r, c0, r, ch.c2) });
  }
  out.series = out.series.slice(0, 8);
  return { refs: out, D };
}
function chartXml(ch, refs, D) {
  const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const fill = k => `<a:solidFill><a:srgbClr val="${XL_PAL[k % 8]}"/></a:solidFill>`;
  const tx = (s, k) => s.name ? `<c:tx><c:strRef><c:f>${xmlE(s.name)}</c:f></c:strRef></c:tx>` : `<c:tx><c:v>Série ${k + 1}</c:v></c:tx>`;
  const cat = refs.cat ? `<c:cat><c:strRef><c:f>${xmlE(refs.cat)}</c:f></c:strRef></c:cat>` : '';
  const val = s => `<c:val><c:numRef><c:f>${xmlE(s.val)}</c:f></c:numRef></c:val>`;
  const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="fr-FR" sz="1400" b="1"/><a:t>${xmlE(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = ch.type === 'pie' || refs.series.length > 1 ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : '';
  const axes = horiz => `<c:catAx><c:axId val="500"/><c:scaling><c:orientation val="${horiz ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr><c:crossAx val="501"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
    + `<c:valAx><c:axId val="501"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E5E5E5"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:crossAx val="500"/><c:crosses val="${horiz ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
  let plot;
  if (ch.type === 'pie') {
    const s = refs.series[0];
    const n = D.cats.length;
    const pts = Array.from({ length: Math.min(n, 64) }, (_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${fill(i)}<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('');
    plot = `<c:doughnutChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${tx(s, 0)}${pts}${cat}${val(s)}</c:ser><c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
  } else if (ch.type === 'line') {
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` + refs.series.map((s, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx(s, k)}<c:spPr><a:ln w="28575" cap="rnd">${fill(k)}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${fill(k)}<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>${cat}${val(s)}<c:smooth val="0"/></c:ser>`).join('')
      + `<c:marker val="1"/><c:axId val="500"/><c:axId val="501"/></c:lineChart>` + axes(false);
  } else {
    const horiz = ch.type === 'bar', stack = ch.type === 'stack';
    plot = `<c:barChart><c:barDir val="${horiz ? 'bar' : 'col'}"/><c:grouping val="${stack ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>`
      + refs.series.map((s, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx(s, k)}<c:spPr>${fill(k)}</c:spPr><c:invertIfNegative val="0"/>${cat}${val(s)}</c:ser>`).join('')
      + `<c:gapWidth val="${stack ? 60 : 80}"/>${stack ? '<c:overlap val="100"/>' : refs.series.length > 1 ? '<c:overlap val="-8"/>' : ''}<c:axId val="500"/><c:axId val="501"/></c:barChart>` + axes(horiz);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace ${C}><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr></c:chartSpace>`;
}
const cfFormulaVal = v => { const n = cfNum(v); return n != null ? String(n) : `"${String(v).replace(/"/g, '""')}"`; };
function cfRuleXml(rule, prio, dxfId) {
  const tl = addr(rule.r1, rule.c1), a = rule.a ?? '', q = `"${String(a).replace(/"/g, '""')}"`;
  const ops = { gt: 'greaterThan', ge: 'greaterThanOrEqual', lt: 'lessThan', le: 'lessThanOrEqual', eq: 'equal', ne: 'notEqual', between: 'between' };
  const d = ` dxfId="${dxfId}" priority="${prio}"`;
  if (ops[rule.t]) return `<cfRule type="cellIs"${d} operator="${ops[rule.t]}"><formula>${xmlE(cfFormulaVal(a))}</formula>${rule.t === 'between' ? `<formula>${xmlE(cfFormulaVal(rule.b))}</formula>` : ''}</cfRule>`;
  if (rule.t === 'contains') return `<cfRule type="containsText"${d} operator="containsText" text="${xmlE(a)}"><formula>${xmlE(`NOT(ISERROR(SEARCH(${q},${tl})))`)}</formula></cfRule>`;
  if (rule.t === 'notContains') return `<cfRule type="notContainsText"${d} operator="notContains" text="${xmlE(a)}"><formula>${xmlE(`ISERROR(SEARCH(${q},${tl}))`)}</formula></cfRule>`;
  if (rule.t === 'begins') return `<cfRule type="beginsWith"${d} operator="beginsWith" text="${xmlE(a)}"><formula>${xmlE(`LEFT(${tl},LEN(${q}))=${q}`)}</formula></cfRule>`;
  if (rule.t === 'empty') return `<cfRule type="containsBlanks"${d}><formula>${xmlE(`LEN(TRIM(${tl}))=0`)}</formula></cfRule>`;
  if (rule.t === 'notEmpty') return `<cfRule type="notContainsBlanks"${d}><formula>${xmlE(`LEN(TRIM(${tl}))>0`)}</formula></cfRule>`;
  if (rule.t === 'dup') return `<cfRule type="duplicateValues"${d}/>`;
  if (rule.t === 'above') return `<cfRule type="aboveAverage"${d}/>`;
  if (rule.t === 'below') return `<cfRule type="aboveAverage"${d} aboveAverage="0"/>`;
  if (rule.t === 'formula') return `<cfRule type="expression"${d}><formula>${xmlE(xlF(String(a).replace(/^=/, '')))}</formula></cfRule>`;
  return '';
}
const scaleHex = k => k ? XL_RGB[k][1] : 'FFFFFF';
function cfScaleXml(rule, prio) {
  const lo = rule.lo ? (rule.mid ? XL_RGB[rule.lo][0] : XL_RGB[rule.lo][1]) : 'FFFFFF';
  const mid = rule.mid ? 'FFE699' : null, hi = rule.mid ? XL_RGB[rule.hi || 'vert'][0] : XL_RGB[rule.hi || 'vert'][1];
  return `<cfRule type="colorScale" priority="${prio}"><colorScale><cfvo type="min"/>${mid ? '<cfvo type="percentile" val="50"/>' : ''}<cfvo type="max"/><color rgb="FF${lo}"/>${mid ? `<color rgb="FF${mid}"/>` : ''}<color rgb="FF${hi}"/></colorScale></cfRule>`;
}
function dxfXml(st) {
  const font = (st.b ? '<b/>' : '') + (st.fc ? `<color rgb="FF${XL_RGB[st.fc][0]}"/>` : '');
  return `<dxf>${font ? `<font>${font}</font>` : ''}${st.bg ? `<fill><patternFill patternType="solid"><bgColor rgb="FF${XL_RGB[st.bg][1]}"/></patternFill></fill>` : ''}</dxf>`;
}
// Complète le fichier écrit par la bibliothèque : graphiques, règles, listes déroulantes.
async function xlsxExtras(buf, cur) {
  const hasExtra = cur.sheets.some(sh => (sh.charts || []).length || (sh.cf || []).length || Object.values(sh.cells).some(x => x.f && x.f.dd));
  if (!hasExtra) return buf;
  await loadScript(LIB.jszip);
  const zip = await window.JSZip.loadAsync(buf);
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  let ct = await zip.file('[Content_Types].xml').async('string');
  let styles = await zip.file('xl/styles.xml').async('string');
  const dxfs = [];
  let chartN = 0, drawN = 0;
  for (let i = 0; i < cur.sheets.length; i++) {
    const sh = cur.sheets[i], path = `xl/worksheets/sheet${i + 1}.xml`;
    const f = zip.file(path); if (!f) continue;
    let xml = await f.async('string');
    const vals = SH.hf && SH.ids ? SH.hf.getSheetValues(SH.ids[i]) : [];
    const cellsOf = (r, c) => sh.cells[addr(r, c)];
    // règles
    let cfx = '', prio = 1;
    (sh.cf || []).forEach(rule => {
      const sq = rngRef(rule);
      if (rule.t === 'scale') { cfx += `<conditionalFormatting sqref="${sq}">${cfScaleXml(rule, prio++)}</conditionalFormatting>`; return; }
      dxfs.push(dxfXml(rule.st || {}));
      const rx = cfRuleXml(rule, prio++, dxfs.length - 1);
      if (rx) cfx += `<conditionalFormatting sqref="${sq}">${rx}</conditionalFormatting>`;
    });
    // listes déroulantes, regroupées par liste identique
    const groups = new Map();
    Object.entries(sh.cells).forEach(([k, x]) => { if (x.f && x.f.dd) { const key = JSON.stringify(x.f.dd); if (!groups.has(key)) groups.set(key, []); groups.get(key).push(k); } });
    let dvx = '';
    groups.forEach((cells, key) => {
      const dd = JSON.parse(key);
      const f1 = Array.isArray(dd) ? `"${dd.map(v => String(v).replace(/,/g, ' ').replace(/"/g, "'")).join(',').slice(0, 250)}"` : xlF(String(dd).slice(1));
      dvx += `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Grimoire" error="Choisissez une valeur de la liste." sqref="${cells.join(' ')}"><formula1>${xmlE(f1)}</formula1></dataValidation>`;
    });
    if (dvx) dvx = `<dataValidations count="${groups.size}">${dvx}</dataValidations>`;
    if (cfx || dvx) {
      const at = xml.indexOf('</mergeCells>') >= 0 ? xml.indexOf('</mergeCells>') + '</mergeCells>'.length : xml.indexOf('</sheetData>') + '</sheetData>'.length;
      xml = xml.slice(0, at) + cfx + dvx + xml.slice(at);
    }
    // graphiques
    if ((sh.charts || []).length) {
      drawN++;
      const dPath = `xl/drawings/drawing${drawN}.xml`;
      let anchors = '', drels = '';
      sh.charts.forEach((ch, k) => {
        chartN++;
        const { refs, D } = chartRefs(ch, vals, cellsOf, sh.name);
        zip.file(`xl/charts/chart${chartN}.xml`, chartXml(ch, refs, D));
        ct = ct.replace('</Types>', `<Override PartName="/xl/charts/chart${chartN}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
        drels += `<Relationship Id="rId${k + 1}" Type="${R}/chart" Target="../charts/chart${chartN}.xml"/>`;
        anchors += `<xdr:oneCellAnchor><xdr:from><xdr:col>${ch.at.c}</xdr:col><xdr:colOff>${Math.max(0, ch.dx || 0) * 9525}</xdr:colOff><xdr:row>${ch.at.r}</xdr:row><xdr:rowOff>${Math.max(0, ch.dy || 0) * 9525}</xdr:rowOff></xdr:from><xdr:ext cx="${ch.w * 9525}" cy="${ch.h * 9525}"/>`
          + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${k + 2}" name="${xmlE(ch.title || 'Graphique ' + (k + 1))}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId${k + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
      });
      zip.file(dPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}">${anchors}</xdr:wsDr>`);
      zip.file(`xl/drawings/_rels/drawing${drawN}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drels}</Relationships>`);
      ct = ct.replace('</Types>', `<Override PartName="/${dPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
      const relPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
      let rels = zip.file(relPath) ? await zip.file(relPath).async('string') : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
      rels = rels.replace('</Relationships>', `<Relationship Id="rIdGrDr1" Type="${R}/drawing" Target="../drawings/drawing${drawN}.xml"/></Relationships>`);
      zip.file(relPath, rels);
      const tag = '<drawing r:id="rIdGrDr1"/>';
      const li = xml.indexOf('<legacyDrawing');
      xml = li >= 0 ? xml.slice(0, li) + tag + xml.slice(li) : xml.replace('</worksheet>', tag + '</worksheet>');
    }
    zip.file(path, xml);
  }
  if (dxfs.length) {
    const block = `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`;
    styles = /<dxfs count="0"\/>/.test(styles) ? styles.replace(/<dxfs count="0"\/>/, block) : /<dxfs[\s\S]*?<\/dxfs>/.test(styles) ? styles.replace(/<dxfs[\s\S]*?<\/dxfs>/, block) : styles.replace('<tableStyles', block + '<tableStyles');
    zip.file('xl/styles.xml', styles);
  }
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/* --- import : lecture des graphiques, règles et listes d'un fichier Excel */
function nearestColor(hex, which) {
  if (!hex) return '';
  hex = hex.replace(/^FF/i, '').slice(-6).toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(hex)) return '';
  const A = hexRgb(hex);
  let best = '', bd = Infinity;
  Object.entries(XL_RGB).forEach(([k, v]) => { const B = hexRgb(v[which]); const d = A.reduce((s, x, i) => s + (x - B[i]) ** 2, 0); if (d < bd) { bd = d; best = k; } });
  return best;
}
async function xlsxReadExtras(buf, sheetNames) {
  await loadScript(LIB.jszip);
  const zip = await window.JSZip.loadAsync(buf);
  const P = s => s ? new DOMParser().parseFromString(s, 'application/xml') : null;
  const get = async p => { const f = zip.file(p); return f ? f.async('string') : null; };
  const kids = (el, n) => el ? [...el.children].filter(c => c.localName === n) : [];
  const all = (el, n) => el ? [...el.getElementsByTagName('*')].filter(c => c.localName === n) : [];
  const relsOf = async p => { const m = {}; const x = P(await get(p.replace(/([^/]+)$/, '_rels/$1.rels'))); if (x) [...x.getElementsByTagName('Relationship')].forEach(r => { m[r.getAttribute('Id')] = r.getAttribute('Target'); }); return m; };
  const resolve = (base, t) => { if (t.startsWith('/')) return t.slice(1); const parts = base.split('/'); parts.pop(); t.split('/').forEach(s => { if (s === '..') parts.pop(); else if (s !== '.') parts.push(s); }); return parts.join('/'); };
  const wb = P(await get('xl/workbook.xml')); if (!wb) return {};
  const wbRels = await relsOf('xl/workbook.xml');
  const st = P(await get('xl/styles.xml'));
  const dxfs = all(st && st.documentElement, 'dxfs').length ? kids(all(st.documentElement, 'dxfs')[0], 'dxf') : [];
  const dxfStyle = id => {
    const d = dxfs[+id]; if (!d) return {};
    const o = {};
    const fnt = all(d, 'font')[0]; if (fnt) { if (all(fnt, 'b').length) o.b = 1; const c = all(fnt, 'color')[0]; if (c && c.getAttribute('rgb')) o.fc = nearestColor(c.getAttribute('rgb'), 0); }
    const bg = all(d, 'bgColor')[0] || all(d, 'fgColor')[0]; if (bg && bg.getAttribute('rgb')) o.bg = nearestColor(bg.getAttribute('rgb'), 1);
    return o;
  };
  // styles des cellules (polices, couleurs, bordures, alignement)
  const XF = [];
  if (st) {
    const sect = n => all(st.documentElement, n)[0];
    const fonts = kids(sect('fonts'), 'font').map(fn => {
      const v = (n, a = 'val') => { const e = kids(fn, n)[0]; return e ? (e.getAttribute(a) ?? '') : null; };
      const col = kids(fn, 'color')[0];
      return { b: v('b') !== null && v('b') !== '0', i: v('i') !== null && v('i') !== '0', u: v('u') !== null && v('u') !== 'none', sz: +v('sz') || 0, name: v('name') || '', rgb: col && col.getAttribute('rgb') };
    });
    const fills = kids(sect('fills'), 'fill').map(fl => { const pf = kids(fl, 'patternFill')[0]; if (!pf || pf.getAttribute('patternType') !== 'solid') return null; const fg = kids(pf, 'fgColor')[0]; return fg && fg.getAttribute('rgb'); });
    const borders = kids(sect('borders'), 'border').map(bd => { const o = {}; [['top', 'bt'], ['bottom', 'bb'], ['left', 'bl'], ['right', 'br']].forEach(([n, k]) => { const e = kids(bd, n)[0]; if (e && e.getAttribute('style') && e.getAttribute('style') !== 'none') o[k] = 1; }); return o; });
    const f0 = fonts[0] || {};
    kids(sect('cellXfs'), 'xf').forEach(xf => {
      const o = {}, fn = fonts[+xf.getAttribute('fontId') || 0] || {}, fill = fills[+xf.getAttribute('fillId') || 0], bd = borders[+xf.getAttribute('borderId') || 0] || {};
      if (fn.b) o.b = 1; if (fn.i) o.i = 1; if (fn.u) o.u = 1;
      const ff = mapFont(fn.name); if (ff && fn.name !== f0.name && GRID_FONTS.includes(ff) && ff !== 'Instrument Sans') o.ff = ff;
      if (fn.sz && fn.sz !== f0.sz && fn.sz !== 10) o.fs = GRID_SIZES.reduce((a, b) => Math.abs(b - fn.sz) < Math.abs(a - fn.sz) ? b : a);
      if (fn.rgb && !/^(FF)?000000$/i.test(fn.rgb)) { const c = nearestColor(fn.rgb, 0); if (c) o.fc = c; }
      if (fill && !/^(FF)?FFFFFF$/i.test(fill)) { const c = nearestColor(fill, 1); if (c) o.bg = c; }
      Object.assign(o, bd);
      const al = kids(xf, 'alignment')[0];
      if (al) { const h = al.getAttribute('horizontal'); if (h === 'center' || h === 'centerContinuous') o.a = 'ac'; else if (h === 'right') o.a = 'ar'; else if (h === 'left') o.a = 'al'; if (al.getAttribute('wrapText') === '1' || al.getAttribute('wrapText') === 'true') o.w = 1; }
      XF.push(o);
    });
  }
  const out = {};
  const sheets = all(wb.documentElement, 'sheet');
  for (const s of sheets) {
    const name = s.getAttribute('name'); if (!sheetNames.includes(name)) continue;
    const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const path = resolve('xl/workbook.xml', wbRels[rid] || ''); const x = P(await get(path)); if (!x) continue;
    const res = { cf: [], dd: {}, charts: [], fmt: {} };
    if (XF.length) all(x.documentElement, 'c').forEach(c => { const si = +c.getAttribute('s') || 0; const o = XF[si]; if (si && o && Object.keys(o).length) res.fmt[c.getAttribute('r')] = o; });
    // règles
    all(x.documentElement, 'conditionalFormatting').forEach(cf => {
      const sq = (cf.getAttribute('sqref') || '').split(/\s+/)[0]; const o = rangeObj(sq); if (!o) return;
      kids(cf, 'cfRule').forEach(rule => {
        const t = rule.getAttribute('type'), op = rule.getAttribute('operator'), fs = kids(rule, 'formula').map(f => f.textContent);
        const unq = v => v == null ? '' : /^".*"$/.test(v) ? v.slice(1, -1).replace(/""/g, '"') : String(v).replace('.', ',');
        const st2 = dxfStyle(rule.getAttribute('dxfId'));
        const add = (tt, extra) => res.cf.push({ id: uid(), ...o, t: tt, st: st2, ...extra });
        const OPS = { greaterThan: 'gt', greaterThanOrEqual: 'ge', lessThan: 'lt', lessThanOrEqual: 'le', equal: 'eq', notEqual: 'ne', between: 'between' };
        if (t === 'cellIs' && OPS[op]) add(OPS[op], { a: unq(fs[0]), ...(op === 'between' ? { b: unq(fs[1]) } : {}) });
        else if (t === 'containsText') add('contains', { a: rule.getAttribute('text') || '' });
        else if (t === 'notContainsText') add('notContains', { a: rule.getAttribute('text') || '' });
        else if (t === 'beginsWith') add('begins', { a: rule.getAttribute('text') || '' });
        else if (t === 'containsBlanks') add('empty');
        else if (t === 'notContainsBlanks') add('notEmpty');
        else if (t === 'duplicateValues') add('dup');
        else if (t === 'aboveAverage') add(rule.getAttribute('aboveAverage') === '0' ? 'below' : 'above');
        else if (t === 'expression' && fs[0]) add('formula', { a: '=' + excelToFr(fs[0]) });
        else if (t === 'colorScale') {
          const cols = all(rule, 'color').map(c => c.getAttribute('rgb'));
          const r2 = { id: uid(), ...o, t: 'scale', hi: nearestColor(cols[cols.length - 1], cols.length === 3 ? 0 : 1) || 'vert' };
          if (cols.length === 3) { r2.lo = nearestColor(cols[0], 0) || 'rouge'; r2.mid = 'jaune'; } else { const lo = nearestColor(cols[0], 1); if (lo && !/^F{6}$/i.test(String(cols[0]).slice(-6))) r2.lo = lo; }
          res.cf.push(r2);
        }
      });
    });
    // listes déroulantes
    all(x.documentElement, 'dataValidation').forEach(dv => {
      if (dv.getAttribute('type') !== 'list') return;
      const f1 = (kids(dv, 'formula1')[0] || {}).textContent || '';
      const dd = /^".*"$/.test(f1) ? f1.slice(1, -1).split(',').map(v => v.trim()).filter(Boolean) : f1 ? '=' + excelToFr(f1) : null;
      if (!dd) return;
      (dv.getAttribute('sqref') || '').split(/\s+/).forEach(ref => { const o = rangeObj(ref); if (!o) return; for (let r = o.r1; r <= Math.min(o.r2, o.r1 + 999); r++) for (let c = o.c1; c <= o.c2; c++) res.dd[addr(r, c)] = dd; });
    });
    // graphiques
    const dr = all(x.documentElement, 'drawing')[0];
    if (dr) {
      const srels = await relsOf(path);
      const dpath = resolve(path, srels[dr.getAttribute('r:id') || dr.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')] || '');
      const dx = P(await get(dpath)); const drels = dpath ? await relsOf(dpath) : {};
      if (dx) for (const an of [...dx.documentElement.children]) {
        const cref = all(an, 'chart')[0]; if (!cref) continue;
        const cid = cref.getAttribute('r:id') || cref.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        const cx = P(await get(resolve(dpath, drels[cid] || ''))); if (!cx) continue;
        const from = all(an, 'from')[0], to = all(an, 'to')[0], ext = kids(an, 'ext')[0];
        const num = (el, n) => +((all(el, n)[0] || {}).textContent || 0);
        const at = { r: num(from, 'row'), c: num(from, 'col') };
        let w = 480, h = 300;
        if (ext) { w = Math.round(+ext.getAttribute('cx') / 9525); h = Math.round(+ext.getAttribute('cy') / 9525); }
        else if (to) { w = Math.max(240, (num(to, 'col') - at.c) * 100); h = Math.max(180, (num(to, 'row') - at.r) * 21); }
        const plot = all(cx.documentElement, 'plotArea')[0]; if (!plot) continue;
        const kind = [...plot.children].find(c => /Chart$/.test(c.localName)); if (!kind) continue;
        let type = 'col';
        if (/^(pie|doughnut|pie3D)Chart$/.test(kind.localName)) type = 'pie';
        else if (/^line/.test(kind.localName) || kind.localName === 'areaChart' || kind.localName === 'scatterChart') type = 'line';
        else if (/^bar/.test(kind.localName)) { const dir = (kids(kind, 'barDir')[0] || { getAttribute: () => 'col' }).getAttribute('val'); const grp = (kids(kind, 'grouping')[0] || { getAttribute: () => '' }).getAttribute('val'); type = /stacked/i.test(grp || '') && dir !== 'bar' ? 'stack' : dir === 'bar' ? 'bar' : 'col'; }
        const refs = all(kind, 'f').map(f => f.textContent);
        let box = null, sheetOk = true;
        refs.forEach(f => {
          const m = /^(?:'?(.+?)'?!)?(.+)$/.exec(f.trim()); if (!m) return;
          if (m[1] && m[1].replace(/''/g, "'") !== name) sheetOk = false;
          const o = rangeObj(m[2]); if (!o) return;
          box = box ? { r1: Math.min(box.r1, o.r1), c1: Math.min(box.c1, o.c1), r2: Math.max(box.r2, o.r2), c2: Math.max(box.c2, o.c2) } : o;
        });
        if (!box || !sheetOk) continue;
        const valRef = (all(kind, 'val')[0] && all(all(kind, 'val')[0], 'f')[0] || {}).textContent || '';
        const vo = rangeObj(valRef.replace(/^.*!/, ''));
        const byRow = !!(vo && vo.r1 === vo.r2 && vo.c2 > vo.c1);
        const tt = all(all(cx.documentElement, 'title')[0], 't').map(t => t.textContent).join('');
        res.charts.push({ id: uid(), type, ...box, title: tt, byRow, at, dx: 0, dy: 0, w: Math.max(220, Math.min(w, 1200)), h: Math.max(160, Math.min(h, 900)) });
      }
    }
    out[name] = res;
  }
  return out;
}

function xlStyle(f) {
  const s = {}, font = {}, al = {}, bd = {};
  if (f.b) font.bold = true; if (f.i) font.italic = true; if (f.u) font.underline = true;
  if (f.ff) font.name = f.ff; if (f.fs) font.sz = f.fs;
  if (f.fc && XL_RGB[f.fc]) font.color = { rgb: XL_RGB[f.fc][0] };
  if (Object.keys(font).length) s.font = font;
  if (f.bg && XL_RGB[f.bg]) s.fill = { patternType: 'solid', fgColor: { rgb: XL_RGB[f.bg][1] } };
  if (f.a) al.horizontal = { al: 'left', ac: 'center', ar: 'right' }[f.a];
  if (f.w) { al.wrapText = true; al.vertical = 'top'; }
  if (Object.keys(al).length) s.alignment = al;
  [['bt', 'top'], ['bb', 'bottom'], ['bl', 'left'], ['br', 'right']].forEach(([k, n]) => { if (f[k]) bd[n] = { style: 'thin', color: { rgb: '1B1512' } }; });
  if (Object.keys(bd).length) s.border = bd;
  const z = xlNumFmt(f.n, f.d); if (z) s.numFmt = z;
  if (f.mg && !f.a) { al.horizontal = 'center'; s.alignment = al; }
  if (f.mg) { al.vertical = 'center'; s.alignment = al; }
  return Object.keys(s).length ? s : null;
}
/* --- Macros VBA : conservées à l'identique (jamais exécutées par Grimoire) */
async function readVba(buf, sheets) {
  await loadScript(LIB.jszip);
  const zip = await window.JSZip.loadAsync(buf);
  const bin = zip.file('xl/vbaProject.bin'); if (!bin) return null;
  const P = t => new DOMParser().parseFromString(t, 'application/xml');
  const wbx = P(await zip.file('xl/workbook.xml').async('string'));
  const byLocal = (el, n) => [...el.getElementsByTagName('*')].filter(c => c.localName === n);
  const wpr = byLocal(wbx.documentElement, 'workbookPr')[0];
  const out = { bin: bytesToB64(await bin.async('uint8array')), wb: (wpr && wpr.getAttribute('codeName')) || 'ThisWorkbook', codes: {} };
  const sig = zip.file('xl/vbaProjectSignature.bin'); if (sig) out.signed = true;
  const relsF = zip.file('xl/_rels/workbook.xml.rels'), rels = {};
  if (relsF) byLocal(P(await relsF.async('string')).documentElement, 'Relationship').forEach(r => { rels[r.getAttribute('Id')] = r.getAttribute('Target'); });
  for (const sh of byLocal(wbx.documentElement, 'sheet')) {
    const rid = sh.getAttribute('r:id') || sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const t = (rels[rid] || '').replace(/^\/?xl\//, '').replace(/^\//, ''); const f = zip.file('xl/' + t); if (!f) continue;
    const m = /<(?:\w+:)?sheetPr\b[^>]*\bcodeName="([^"]+)"/.exec(await f.async('string'));
    const mine = sheets.find(x => x.src === sh.getAttribute('name'));
    if (m && mine) out.codes[mine.id] = m[1];
  }
  return out;
}
async function addVba(buf, cur) {
  await loadScript(LIB.jszip);
  const zip = await window.JSZip.loadAsync(buf), V = cur.vba;
  zip.file('xl/vbaProject.bin', b64ToBytes(V.bin));
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml');
  if (!/PartName="\/xl\/vbaProject\.bin"/.test(ct)) ct = ct.replace('</Types>', '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>');
  zip.file('[Content_Types].xml', ct);
  let rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  if (!/vbaProject/.test(rels)) rels = rels.replace('</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
  zip.file('xl/_rels/workbook.xml.rels', rels);
  const setCode = (xml, tag, code, after) => {
    const re = new RegExp(`<${tag}\\b([^>]*?)(/?)>`);
    if (re.test(xml)) return xml.replace(re, (m0, attrs, sl) => `<${tag}${attrs.replace(/\scodeName="[^"]*"/, '')} codeName="${xmlE(code)}"${sl}>`);
    return xml.replace(after, m0 => m0 + `<${tag} codeName="${xmlE(code)}"/>`);
  };
  let wbx = await zip.file('xl/workbook.xml').async('string');
  wbx = /<workbookPr\b/.test(wbx) ? setCode(wbx, 'workbookPr', V.wb || 'ThisWorkbook') : wbx.replace(/(<bookViews|<sheets)/, `<workbookPr codeName="${xmlE(V.wb || 'ThisWorkbook')}"/>$1`);
  zip.file('xl/workbook.xml', wbx);
  for (let i = 0; i < cur.sheets.length; i++) {
    const code = (V.codes || {})[cur.sheets[i].id]; const f = zip.file(`xl/worksheets/sheet${i + 1}.xml`); if (!code || !f) continue;
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, setCode(await f.async('string'), 'sheetPr', code, /<worksheet\b[^>]*>/));
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
const exportXlsm = () => exportXlsx(true);
async function exportXlsx(macro) {
  commitEdit(false);
  await loadScript(LIB.xlsx);
  const X = window.XLSX, cur = state.cur, wb = X.utils.book_new();
  cur.sheets.forEach((sh, i) => {
    const vals = SH.hf && SH.ids ? SH.hf.getSheetValues(SH.ids[i]) : [];
    const ws = {};
    let maxR = 0, maxC = 0;
    Object.entries(sh.cells).forEach(([k, x]) => {
      const p = parseAddr(k); if (!p) return;
      const f = x.f || {}, st = xlStyle(f);
      if ((x.v === '' || x.v == null) && !st && !f.note) return;
      maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c);
      const v = (vals[p.r] || [])[p.c];
      let cell;
      if (x.v === '' || x.v == null) cell = { t: 's', v: '' };
      else if (String(x.v)[0] === '=') {
        const fv = typeof v === 'number' ? { t: 'n', v } : typeof v === 'boolean' ? { t: 'b', v } : typeof v === 'object' && v ? { t: 'e', v: 0, w: v.value } : { t: 's', v: v == null ? '' : String(v) };
        cell = { ...fv, f: xlF(String(x.v).slice(1)) };
      } else {
        const n = parseNum(x.v), d = n == null ? parseDate(x.v) : null;
        cell = /^(VRAI|FAUX)$/i.test(String(x.v).trim()) ? { t: 'b', v: /^VRAI$/i.test(String(x.v).trim()) } : n != null ? { t: 'n', v: n } : d != null ? { t: 'n', v: d } : { t: 's', v: String(x.v).replace(/^'/, '') };
      }
      const zf = xlNumFmt(f.n, f.d); if (zf && cell.t === 'n') cell.z = zf;
      if (st) cell.s = st;
      if (f.note) { cell.c = [{ a: 'Grimoire', t: f.note }]; cell.c.hidden = true; }
      ws[X.utils.encode_cell({ r: p.r, c: p.c })] = cell;
    });
    ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
    const hidC = new Set((sh.hid && sh.hid.c) || []), hidR = new Set((sh.hid && sh.hid.r) || []);
    hidC.forEach(c => { maxC = Math.max(maxC, c); }); hidR.forEach(r => { maxR = Math.max(maxR, r); });
    ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
    ws['!cols'] = Array.from({ length: maxC + 1 }, (_, c) => hidC.has(c) ? { wpx: sh.colW[c] || 104, hidden: true } : { wpx: sh.colW[c] || 104 });
    if (hidR.size) { ws['!rows'] = []; hidR.forEach(r => { ws['!rows'][r] = { hidden: true }; }); }
    const merges = [];
    Object.entries(sh.cells).forEach(([k, x]) => { const m = x.f && x.f.mg; if (!m) return; const p = parseAddr(k); merges.push({ s: { r: p.r, c: p.c }, e: { r: p.r + (m.r || 1) - 1, c: p.c + (m.c || 1) - 1 } }); });
    if (merges.length) ws['!merges'] = merges;
    X.utils.book_append_sheet(wb, ws, sh.name.slice(0, 31));
  });
  const names = Object.entries(cur.names || {}).filter(([, nm]) => cur.sheets.some(s => s.name === nm.s) && rangeObj(nm.ref)).map(([n, nm]) => ({ Name: n, Ref: `'${nm.s.replace(/'/g, "''")}'!${absRef(rangeObj(nm.ref))}` }));
  if (names.length) wb.Workbook = { Names: names };
  let buf = X.write(wb, { type: 'array', bookType: 'xlsx' });
  try { buf = await xlsxExtras(buf, cur); } catch (e) { console.warn('Compléments Excel ignorés', e); toast('Graphiques ou règles non exportés : le reste du classeur est bien là.', true); }
  if (macro === true && cur.vba) { try { buf = await addVba(buf, cur); } catch (e) { console.warn(e); toast('Macros non réécrites : le classeur est enregistré sans elles.', true); } }
  await saveFile(safeName(cur.title) + (macro === true ? '.xlsm' : '.xlsx'), new Blob([buf]));
}
function exportCsv() {
  commitEdit(false);
  const cur = state.cur; let maxR = -1, maxC = -1;
  Object.entries(cur.cells).forEach(([k, x]) => { if (x.v !== '') { const p = parseAddr(k); maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); } });
  const lines = [];
  for (let r = 0; r <= maxR; r++) {
    const row = [];
    for (let c = 0; c <= maxC; c++) {
      const v = (SH.vals[r] || [])[c];
      let s = v == null ? '' : typeof v === 'number' ? (((cellOf(r, c) || {}).f || {}).n === 'date' ? fmtDate(v) : numToRaw(v)) : typeof v === 'object' ? (v.value || '') : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v);
      if (/[;"\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
      row.push(s);
    }
    lines.push(row.join(';'));
  }
  return saveFile(safeName(cur.title) + '.csv', '﻿' + lines.join('\r\n'));
}

/* --- impression */
function buildPrint() {
  if (state.view !== 'sheet' || !state.cur) { $('#printArea').innerHTML = ''; return; }
  commitEdit(false);
  const cur = state.cur; let maxR = -1, maxC = -1;
  Object.entries(cur.cells).forEach(([k, x]) => { if (x.v !== '') { const p = parseAddr(k); if (p) { maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); } } });
  const hidden = { has: r => typeof rowHidden === 'function' && rowHidden(r) };
  const colPx = c => (cur.colW && cur.colW[c]) || 100;
  let totalW = 0; for (let c = 0; c <= maxC; c++) totalW += colPx(c);
  const sheetName = cur.sheets && cur.sheets.length > 1 && cur.sheets[cur.active] ? ' · ' + cur.sheets[cur.active].name : '';
  let h = `<h1 class="print-title">${esc((cur.title || '') + sheetName)}</h1><table class="print-grid"><colgroup>`;
  for (let c = 0; c <= maxC; c++) h += `<col style="width:${colPx(c)}px">`;
  h += '</colgroup>';
  for (let r = 0; r <= maxR; r++) {
    if (hidden.has(r)) continue;
    h += '<tr>';
    for (let c = 0; c <= maxC; c++) {
      const d = display(r, c);
      const st = [d.bg && 'background:' + d.bg, d.fc && 'color:' + d.fc, d.bd && 'box-shadow:' + d.bd].filter(Boolean).join(';');
      h += `<td class="${d.cls}"${st ? ` style="${st}"` : ''}>${esc(d.txt)}</td>`;
    }
    h += '</tr>';
  }
  const charts = (cur.charts || []).map(ch => `<div class="chart static" style="width:${ch.w}px;height:${ch.h}px">${chartHTML(ch, SH.vals, cellOf)}</div>`).join('');
  $('#printArea').innerHTML = h + '</table>' + (charts ? `<div class="print-charts">${charts}</div>` : '');
  // tableau large : impression en paysage
  $('#pageCss').textContent = `@page{size:A4 ${totalW > 700 ? 'landscape' : 'portrait'};margin:12mm}`;
}
window.addEventListener('beforeprint', buildPrint);
function printDoc() {
  if (state.view === 'pdfv') return printPdf();
  hideImgBar();
  if (state.view === 'writer') applyPageLayout();
  buildPrint();
  try { setTimeout(() => window.print(), 50); }
  catch { toast('Impression indisponible ici : ouvrez Grimoire dans son propre onglet.', true); }
}
$('#printBtn').addEventListener('click', printDoc);
$('#saveBtn').addEventListener('click', () => saveDoc(false));
$('#saveAsBtn').addEventListener('click', () => saveDoc(true));

/* --- menu Exporter */
let delArmed = false;
function openExport() {
  const p = $('#exportPop'); const t = state.cur.type;
  delArmed = false;
  p.innerHTML = t === 'pdf'
    ? `<button data-x="pdf">${ic('pdf')}Copie du PDF modifié (.pdf)</button><button data-x="pdfdoc">${ic('doc')}Convertir en document Grimoire</button><button data-x="print">${ic('printer')}Imprimer<small>Ctrl P</small></button>`
      + `<hr><button data-x="dup">${ic('copy')}Dupliquer</button><button data-x="del" class="danger">${ic('trash')}<span>Supprimer le document</span></button>`
    : (t === 'texte'
    ? `<button data-x="docpdf">${ic('pdf')}PDF (.pdf)</button><button data-x="md">${ic('type')}Markdown (.md)</button><button data-x="html">${ic('download')}Page web (.html)</button><button data-x="docx">${ic('doc')}Copie Word (.docx)</button>`
    : `<button data-x="docpdf">${ic('pdf')}PDF (.pdf)</button><button data-x="csv">${ic('download')}CSV pour Excel (.csv)</button><button data-x="xlsx">${ic('grid')}Copie Excel (.xlsx)</button>`)
    + `<button data-x="print">${ic('printer')}Imprimer<small>Ctrl P</small></button>`
    + `<hr><button data-x="dup">${ic('copy')}Dupliquer</button><button data-x="del" class="danger">${ic('trash')}<span>Supprimer le document</span></button>`;
  p.hidden = false;
}
function closeExport() { $('#exportPop').hidden = true; }
const runExport = k => ({ docx: exportDocx, md: exportMd, html: exportHtml, xlsx: exportXlsx, xlsm: exportXlsm, csv: exportCsv, print: printDoc, pdf: exportPdfFile, pdfdoc: pdfToDoc, docpdf: exportDocPdf }[k] || (() => {}))();
$('#exportBtn').addEventListener('click', e => { e.stopPropagation(); $('#exportPop').hidden ? openExport() : closeExport(); });
$('#exportPop').addEventListener('click', async e => {
  e.stopPropagation();
  const b = e.target.closest('[data-x]'); if (!b) return;
  const k = b.dataset.x;
  if (k === 'del') {
    if (!delArmed) { delArmed = true; b.querySelector('span').textContent = 'Confirmer la suppression'; return; }
    const id = state.cur.id; closeExport(); clearTimeout(saveTimer); saveTimer = null; state.cur = null; FH.del(id);
    try { await Store.remove(id); toast('Document supprimé.'); } catch { toast('Suppression impossible pour le moment.', true); }
    show('home'); renderHome(); return;
  }
  closeExport();
  if (k === 'dup') {
    commitEdit(false); await flushSave(); const c = clone(state.cur); delete c.id; delete c._saved; c.title = (c.title || 'Sans titre') + ' (copie)';
    if (c.type === 'pdf') { const b = PDFV.bytes || await Store.getBlob(c.pdfId); c.pdfId = uid(); await Store.putBlob(c.pdfId, b); }
    return createDoc(c.type, c);
  }
  try { await runExport(k); } catch { toast('Export impossible : la bibliothèque n\'a pas pu être chargée.', true); }
});
document.addEventListener('click', e => { if (!$('#exportPop').hidden && !e.target.closest('#exportPop')) closeExport(); });

/* =========================================================
   Plume (assistant d'écriture)
   ========================================================= */
const AI_DEF = { mistral: 'mistral-small-latest', openrouter: 'mistralai/mistral-small-3.2-24b-instruct' };
const AI_SUGG = { mistral: ['mistral-small-latest', 'mistral-medium-latest', 'mistral-large-latest'], openrouter: ['mistralai/mistral-small-3.2-24b-instruct', 'mistralai/mistral-medium-3', 'openrouter/auto'] };
let AI = Object.assign({ provider: 'mistral', key: '', model: '' }, ls.get('grim.ai', {}));
const aiModel = () => AI.model || AI_DEF[AI.provider];
const provName = () => AI.provider === 'mistral' ? 'Mistral' : 'OpenRouter';
async function aiChat(prompt, { signal, onText } = {}) {
  if (!AI.key) throw { code: 'no_key' };
  const url = AI.provider === 'mistral' ? 'https://api.mistral.ai/v1/chat/completions' : 'https://openrouter.ai/api/v1/chat/completions';
  const headers = { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + AI.key };
  if (AI.provider === 'openrouter') { headers['HTTP-Referer'] = location.origin; headers['X-Title'] = 'Grimoire'; }
  let res;
  try { res = await fetch(url, { method: 'POST', headers, signal, body: JSON.stringify({ model: aiModel(), messages: [{ role: 'user', content: prompt }], temperature: 0.4, stream: true }) }); }
  catch (e) { if (e && e.name === 'AbortError') throw e; throw { code: 'network' }; }
  if (!res.ok) { let m = ''; try { const j = await res.json(); m = (j.error && (j.error.message || j.error)) || j.message || ''; } catch {} throw { code: 'http', status: res.status, message: String(m).slice(0, 200) }; }
  const reader = res.body.getReader(), dec = new TextDecoder();
  let buf = '', text = '';
  for (;;) {
    const { done, value } = await reader.read(); if (done) break;
    buf += dec.decode(value, { stream: true });
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
      if (!line.startsWith('data:')) continue;
      const d = line.slice(5).trim(); if (d === '[DONE]') continue;
      try { const j = JSON.parse(d); const t = j.choices && j.choices[0] && j.choices[0].delta && j.choices[0].delta.content; if (t) { text += t; if (onText) onText({ text }); } } catch {}
    }
  }
  return { text };
}
function aiError(e) {
  if (!e) return 'Plume n\'a pas pu répondre.';
  if (e.code === 'no_key') return 'Ajoutez d\'abord une clé API dans les Réglages.';
  if (e.code === 'network') return navigator.onLine === false ? 'Pas de connexion internet.' : `Connexion à ${provName()} impossible.` + (AI.provider === 'mistral' ? ' Si cela persiste, essayez OpenRouter, qui propose aussi les modèles Mistral.' : ' Vérifiez votre connexion.');
  if (e.status === 401 || e.status === 403) return 'Clé API refusée. Vérifiez-la dans les Réglages.';
  if (e.status === 402) return 'Crédit épuisé chez le fournisseur.';
  if (e.status === 404 || e.status === 400) return `Modèle « ${aiModel()} » introuvable ou refusé. Changez de modèle dans les Réglages.` + (e.message ? ' (' + e.message + ')' : '');
  if (e.status === 429) return 'Trop de demandes rapprochées. Réessayez dans un instant.';
  return 'Plume n\'a pas pu répondre' + (e.status ? ` (erreur ${e.status})` : '') + '.';
}
const Plume = {
  available: true, mode: 'texte', ctl: null, result: '', target: null,
  ctx() {
    const c = $('#plumeCtx');
    $('#plumeNote').textContent = AI.key ? `Plume envoie le passage concerné à ${provName()} (${aiModel()}), uniquement quand vous cliquez. Relisez avant d'insérer.` : '';
    if (!AI.key) { c.innerHTML = 'Plume a besoin d\'une clé API Mistral ou OpenRouter. <button class="pbtn" id="plumeSet" style="margin-top:8px">Ouvrir les réglages</button>'; return; }
    if (this.mode === 'texte') {
      const t = selectedText();
      c.innerHTML = t ? `Sélection : <q>${esc(t.length > 220 ? t.slice(0, 220) + '…' : t)}</q>` : 'Aucune sélection : le texte proposé sera inséré à l\'emplacement du curseur.';
    } else {
      const raw = rawOf(SH.sel.r, SH.sel.c);
      c.innerHTML = `Cellule active : <b>${addr(SH.sel.r, SH.sel.c)}</b>${raw ? ` · contenu <code>${esc(raw)}</code>` : ' (vide)'}`;
    }
  }
};
function selectedText() { if (!savedRange) return ''; return savedRange.toString().trim(); }
function openPlume() {
  if (!Plume.available) return;
  Plume.mode = state.cur.type;
  const chips = Plume.mode === 'texte'
    ? [['reformuler', 'Reformuler'], ['concis', 'Plus concis'], ['formel', 'Ton plus formel'], ['corriger', 'Corriger'], ['developper', 'Développer'], ['resumer', 'Résumer le document']]
    : [['expliquer', 'Expliquer la formule'], ['verifier', 'Vérifier le tableau']];
  $('#plumeChips').innerHTML = chips.map(([k, l]) => `<button class="pchip" data-p="${k}">${l}</button>`).join('');
  $('#plumeAsk').placeholder = Plume.mode === 'texte' ? 'Décrivez ce que vous voulez écrire. Exemple : un mot de remerciement chaleureux pour un voisin.' : 'Décrivez le calcul. Exemple : total de la colonne D quand la colonne C vaut « Courses ».';
  $('#plumeGo').querySelector('span').textContent = Plume.mode === 'texte' ? 'Rédiger' : 'Écrire la formule';
  $('#plumeOut').textContent = ''; $('#plumeOut').classList.toggle('mono', Plume.mode !== 'texte'); $('#plumeActs').hidden = true;
  Plume.ctx();
  $('#plume').hidden = false; $('#plumeBtn').setAttribute('aria-pressed', 'true');
}
function closePlume() { if (Plume.ctl) Plume.ctl.abort(); $('#plume').hidden = true; $('#plumeBtn').setAttribute('aria-pressed', 'false'); }
$('#plumeBtn').addEventListener('mousedown', e => e.preventDefault());
$('#plumeBtn').addEventListener('click', () => { commitEdit(false); $('#plume').hidden ? openPlume() : closePlume(); });
$('#plumeX').addEventListener('click', closePlume);
$('#plumeCtx').addEventListener('click', e => { if (e.target.id === 'plumeSet') { closePlume(); openSettings(); } });

const ASKS = {
  reformuler: 'Reformule ce texte en français, même sens, style clair et professionnel.',
  concis: 'Rends ce texte nettement plus concis (environ 40 % plus court) sans perdre d\'information importante.',
  formel: 'Réécris ce texte dans un registre plus soutenu, adapté à un courrier officiel.',
  corriger: 'Corrige l\'orthographe, la grammaire et la typographie française (espaces avant ; : ! ?, guillemets « »), sans changer le style ni le fond.',
  developper: 'Développe ce texte avec des précisions concrètes et opérationnelles, en gardant le même ton.'
};
function docPlain() { return (ed.innerText || '').replace(/\n{3,}/g, '\n\n').trim(); }
function sheetTSV() {
  const cur = state.cur; let maxR = 0, maxC = 0;
  Object.keys(cur.cells).forEach(k => { const p = parseAddr(k); if (p && cur.cells[k].v !== '') { maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); } });
  const lines = ['\t' + Array.from({ length: maxC + 1 }, (_, c) => COLN(c)).join('\t')];
  for (let r = 0; r <= Math.min(maxR, 80); r++) lines.push((r + 1) + '\t' + Array.from({ length: maxC + 1 }, (_, c) => rawOf(r, c)).join('\t'));
  return lines.join('\n').slice(0, 6000);
}
async function runPlume(kind) {
  if (!AI.key) { toast('Ajoutez une clé API dans les Réglages pour utiliser Plume.'); return; }
  if (Plume.ctl) Plume.ctl.abort();
  const ask = $('#plumeAsk').value.trim();
  const title = state.cur.title || 'Sans titre';
  let prompt;
  Plume.target = null;
  if (Plume.mode === 'texte') {
    const sel = selectedText();
    const fmt = '\n\nRéponds uniquement par le texte final, en français, sans introduction ni commentaire. Sépare les paragraphes par une ligne vide. Pour une liste, commence chaque ligne par « - ».';
    if (kind === 'resumer') { prompt = `Résume le document « ${title} » en 5 points clés, une phrase chacun.${fmt}\n\nDocument :\n"""${docPlain().slice(0, 14000)}"""`; Plume.target = 'insert'; }
    else if (kind && ASKS[kind]) {
      const text = sel || (currentBlock() ? currentBlock().innerText : '');
      if (!text.trim()) { toast('Sélectionnez d\'abord un passage du document.'); return; }
      if (!sel) { const r = document.createRange(); r.selectNodeContents(currentBlock()); savedRange = r; }
      prompt = `Tu aides à rédiger un document intitulé « ${title} ». ${ASKS[kind]}${fmt}\n\nTexte :\n"""${text}"""`; Plume.target = 'replace';
    } else {
      if (!ask) { $('#plumeAsk').focus(); return; }
      prompt = `Tu aides à rédiger un document intitulé « ${title} ». Demande : ${ask}${fmt}\n\nContexte, début du document :\n"""${docPlain().slice(0, 4000)}"""` + (sel ? `\n\nPassage sélectionné à remplacer :\n"""${sel}"""` : '');
      Plume.target = sel ? 'replace' : 'insert';
    }
  } else {
    const a = addr(SH.sel.r, SH.sel.c), raw = String(rawOf(SH.sel.r, SH.sel.c));
    const data = `\n\nContenu du classeur (TSV, première ligne = lettres de colonnes, première colonne = numéros de ligne) :\n${sheetTSV()}`;
    if (kind === 'expliquer') {
      if (raw[0] !== '=') { toast('La cellule active ne contient pas de formule.'); return; }
      prompt = `Explique simplement, en 3 phrases au plus et en français, ce que calcule la formule ${raw} placée en ${a}.${data}`; Plume.target = null;
    } else if (kind === 'verifier') {
      prompt = `Relis ce classeur intitulé « ${title} » et signale en français, sous forme de liste courte (« - »), les incohérences possibles : formules qui oublient des lignes, totaux douteux, cellules vides suspectes. Si tout paraît cohérent, dis-le en une phrase.${data}`; Plume.target = null;
    } else {
      if (!ask) { $('#plumeAsk').focus(); return; }
      prompt = `Écris UNE formule de tableur pour la cellule ${a}. Besoin : ${ask}\nUtilise la syntaxe d'Excel en français : noms de fonctions français (SOMME, SI, RECHERCHEV, NB.SI, SOMME.SI…), point-virgule entre les arguments, virgule décimale. Réponds uniquement par la formule, sur une ligne, commençant par =.${data}`; Plume.target = 'cell';
    }
  }
  const out = $('#plumeOut'); out.textContent = 'Plume réfléchit…'; $('#plumeActs').hidden = true;
  $$('#plume button.pchip, #plumeGo').forEach(b => b.disabled = true);
  Plume.ctl = new AbortController();
  try {
    const res = await aiChat(prompt, { signal: Plume.ctl.signal, onText: ({ text }) => { out.textContent = text; } });
    let text = (res.text || '').trim();
    if (Plume.target === 'cell') { const m = text.match(/=[^\n`]+/); text = m ? m[0].trim() : text; }
    Plume.result = text; out.textContent = text;
    showPlumeActs();
  } catch (e) {
    if (e && e.name === 'AbortError') out.textContent = '';
    else out.textContent = aiError(e);
  }
  $$('#plume button.pchip, #plumeGo').forEach(b => b.disabled = false);
  Plume.ctl = null;
}
function textToHtml(t) {
  const blocks = t.split(/\n{2,}/).map(b => b.trim()).filter(Boolean);
  return blocks.map(b => {
    const lines = b.split('\n');
    if (lines.every(l => /^\s*[-•–]\s+/.test(l))) return '<ul>' + lines.map(l => `<li>${esc(l.replace(/^\s*[-•–]\s+/, ''))}</li>`).join('') + '</ul>';
    if (lines.every(l => /^\s*\d+[.)]\s+/.test(l))) return '<ol>' + lines.map(l => `<li>${esc(l.replace(/^\s*\d+[.)]\s+/, ''))}</li>`).join('') + '</ol>';
    return '<p>' + lines.map(esc).join('<br>') + '</p>';
  }).join('');
}
function showPlumeActs() {
  const a = $('#plumeActs'); let h = '';
  if (Plume.target === 'replace') h += `<button class="pbtn" data-pa="replace">Remplacer la sélection</button><button class="pbtn ghost" data-pa="insert">Insérer après</button>`;
  else if (Plume.target === 'insert') h += `<button class="pbtn" data-pa="insert">Insérer au curseur</button>`;
  else if (Plume.target === 'cell') h += `<button class="pbtn" data-pa="cell">Placer en ${addr(SH.sel.r, SH.sel.c)}</button>`;
  h += `<button class="pbtn ghost" data-pa="copy">Copier</button>`;
  a.innerHTML = h; a.hidden = false;
}
$('#plumeChips').addEventListener('mousedown', e => e.preventDefault());
$('#plumeChips').addEventListener('click', e => { const b = e.target.closest('[data-p]'); if (b) runPlume(b.dataset.p); });
$('#plumeGo').addEventListener('click', () => runPlume(null));
$('#plumeAsk').addEventListener('keydown', e => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); runPlume(null); } });
$('#plumeActs').addEventListener('mousedown', e => e.preventDefault());
$('#plumeActs').addEventListener('click', async e => {
  const b = e.target.closest('[data-pa]'); if (!b) return;
  const t = Plume.result; if (!t) return;
  const k = b.dataset.pa;
  if (k === 'copy') { try { await navigator.clipboard.writeText(t); toast('Copié'); } catch { toast('Copie refusée par le navigateur : sélectionnez le texte à la main.', true); } return; }
  if (k === 'cell') { setRaw(SH.sel.r, SH.sel.c, t); toast('Formule placée en ' + addr(SH.sel.r, SH.sel.c)); return; }
  const single = !/\n/.test(t);
  if (k === 'insert' && savedRange) savedRange.collapse(false);
  restoreRange();
  if (single && !/^\s*[-•]/.test(t)) document.execCommand('insertText', false, (k === 'insert' && Plume.target === 'replace' ? ' ' : '') + t);
  else document.execCommand('insertHTML', false, textToHtml(t));
  state.cur.html = ed.innerHTML; scheduleSave(); refreshSoon();
  toast(k === 'replace' ? 'Passage remplacé. Ctrl+Z pour annuler.' : 'Texte inséré. Ctrl+Z pour annuler.');
});

/* =========================================================
   Palette de commandes
   ========================================================= */
let palItems = [], palIdx = 0;
function commands() {
  const v = state.view, L = [];
  const add = (group, label, icon, run, kbd) => L.push({ group, label, icon, run, kbd });
  if (v === 'writer') {
    add('Mise en forme', 'Titre 1', 'type', () => W.block('h1'));
    add('Mise en forme', 'Titre 2', 'type', () => W.block('h2'));
    add('Mise en forme', 'Titre 3', 'type', () => W.block('h3'));
    add('Mise en forme', 'Paragraphe normal', 'type', () => W.block('p'));
    add('Mise en forme', 'Citation', 'type', () => W.block('blockquote'));
    add('Mise en forme', 'Liste à puces', 'ul', W.ul);
    add('Mise en forme', 'Liste numérotée', 'ol', W.ol);
    add('Mise en forme', 'Barré', 'type', W.strike);
    add('Mise en forme', 'Exposant', 'type', W.sup);
    add('Mise en forme', 'Indice', 'type', W.sub);
    add('Mise en forme', 'Couleur du texte', 'textColor', W.color);
    add('Mise en forme', 'Surligner', 'highlighter', W.hilite);
    add('Mise en forme', 'Effacer la mise en forme', 'clearFmt', W.clear);
    add('Mise en forme', 'Liste de cases à cocher', 'check', toggleChecklist);
    add('Mise en forme', 'Augmenter le retrait', 'indent', W.indent, 'Tab');
    add('Mise en forme', 'Diminuer le retrait', 'outdent', W.outdent, 'Maj Tab');
    add('Mise en forme', 'Interligne et espacement', 'spacing', W.spacing);
    add('Insérer', 'Lien', 'link', W.link);
    add('Insérer', 'Saut de page', 'pagebreak', insertPageBreak, 'Ctrl Entrée');
    add('Insérer', 'Sommaire', 'toc', insertToc);
    add('Outils', 'Rechercher', 'search', () => openFind(false), 'Ctrl F');
    add('Outils', 'Rechercher et remplacer', 'search', () => openFind(true), 'Ctrl H');
    add('Outils', 'Mise en page, en-tête et pied de page', 'layout', openPagePop);
    add('Insérer', 'Tableau 3 × 3', 'table', () => insertTable(3, 3));
    add('Insérer', 'Tableau 4 × 5', 'table', () => insertTable(5, 4));
    add('Insérer', 'Commentaire', 'comment', () => addComment(), 'Ctrl Alt M');
    add('Format', 'Modifier les styles (Titre 1, Paragraphe…)', 'type', () => openStylePop());
    add('Révision', 'Activer ou désactiver le suivi des modifications', 'pen', () => toggleTrack(), 'Ctrl Maj E');
    add('Révision', 'Accepter toutes les modifications', 'tick', () => trAll(true));
    add('Révision', 'Refuser toutes les modifications', 'close', () => trAll(false));
    add('Format', 'Deux colonnes', 'cols2', () => setColumns(2));
    add('Format', 'Trois colonnes', 'cols3', () => setColumns(3));
    add('Format', 'Une seule colonne', 'doc', () => setColumns(1));
    add('Affichage', 'Afficher ou masquer les commentaires', 'comment', () => { CM.panel = !CM.panel; renderComments(); });
    add('Insérer', 'Séparateur', 'hr', W.hr);
    add('Insérer', 'Image', 'image', W.image);
    add('Insérer', 'Date du jour', 'pen', () => exec('insertText', new Date().toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })));
    add('Affichage', 'Mode concentration', 'focus', W.focus);
    add('Affichage', 'Afficher ou masquer le plan', 'outline', W.outline);
    if (Plume.available) { add('Plume', 'Reformuler la sélection', 'feather', () => { openPlume(); runPlume('reformuler'); }); add('Plume', 'Corriger la sélection', 'feather', () => { openPlume(); runPlume('corriger'); }); add('Plume', 'Résumer le document', 'feather', () => { openPlume(); runPlume('resumer'); }); add('Plume', 'Rédiger avec Plume…', 'feather', () => { openPlume(); setTimeout(() => $('#plumeAsk').focus(), 30); }); }
    add('Exporter', 'Exporter en PDF (.pdf)', 'pdf', () => runExport('docpdf'));
    add('Exporter', 'Exporter en Markdown (.md)', 'download', () => runExport('md'));
    add('Exporter', 'Exporter une copie Word (.docx)', 'download', () => runExport('docx').catch(() => toast('Export impossible.', true)));
  }
  if (v === 'sheet') {
    add('Cellules', 'Somme automatique', 'sigma', autoSum);
    add('Cellules', 'Recopier vers le bas', 'fillDown', fillDown, 'Ctrl D');
    add('Cellules', 'Recopier vers la droite', 'fillDown', fillRight, 'Ctrl R');
    add('Cellules', 'Format date', 'grid', () => setFmt('n', 'date'));
    add('Cellules', 'Italique', 'type', () => setFmt('i'), 'Ctrl I');
    add('Cellules', 'Souligné', 'type', () => setFmt('u'), 'Ctrl U');
    add('Cellules', 'Retour à la ligne automatique', 'wrap', () => setFmt('w'));
    add('Cellules', 'Couleur de fond', 'bucket', () => openFmt('bg'));
    add('Cellules', 'Couleur du texte', 'textColor', () => openFmt('fc'));
    add('Cellules', 'Toutes les bordures', 'borders', () => setBorders('all'));
    add('Cellules', 'Bordure de contour', 'borders', () => setBorders('outer'));
    add('Cellules', 'Retirer les bordures', 'borders', () => setBorders('none'));
    add('Affichage', 'Figer la première ligne', 'freeze', () => setFreeze(1, (state.cur.frz || {}).c || 0));
    add('Affichage', 'Figer la première colonne', 'freeze', () => setFreeze((state.cur.frz || {}).r || 0, 1));
    add('Affichage', 'Ne rien figer', 'freeze', () => setFreeze(0, 0));
    add('Feuilles', 'Nouvelle feuille', 'plus', () => addSheet());
    add('Feuilles', 'Dupliquer cette feuille', 'copy', () => addSheet(state.cur.sheets[state.cur.active]));
    add('Feuilles', 'Renommer cette feuille', 'pen', () => startRename(state.cur.active));
    add('Feuilles', 'Feuille suivante', 'right', () => switchSheet(Math.min(state.cur.active + 1, state.cur.sheets.length - 1)), 'Ctrl Pg↓');
    add('Feuilles', 'Feuille précédente', 'left', () => switchSheet(Math.max(state.cur.active - 1, 0)), 'Ctrl Pg↑');
    add('Données', 'Trier de A à Z', 'sortAsc', () => sortZone(false));
    add('Données', 'Trier de Z à A', 'sortDesc', () => sortZone(true));
    add('Données', 'Filtrer la colonne', 'filter', openFilter);
    add('Données', 'Retirer tous les filtres', 'filter', clearFilters);
    add('Cellules', 'Format monétaire (€)', 'grid', () => setFmt('n', 'eur'));
    add('Cellules', 'Format pourcentage', 'grid', () => setFmt('n', 'pct'));
    add('Cellules', 'Format deux décimales', 'grid', () => setFmt('n', 'dec2'));
    add('Cellules', 'Format standard', 'grid', () => setFmt('n', ''));
    add('Cellules', 'Gras', 'type', () => setFmt('b'), 'Ctrl B');
    add('Structure', 'Insérer une ligne au-dessus', 'rowIns', () => structural('rowIns'));
    add('Structure', 'Supprimer la ligne', 'rowDel', () => structural('rowDel'));
    add('Structure', 'Insérer une colonne à gauche', 'colIns', () => structural('colIns'));
    add('Structure', 'Supprimer la colonne', 'colDel', () => structural('colDel'));
    add('Insérer', 'Graphique', 'chart', () => openChartEditor());
    add('Insérer', 'Tableau croisé dynamique', 'pivot', openPivot);
    add('Insérer', 'Liste déroulante', 'list', openDdEditor);
    add('Insérer', 'Case à cocher', 'check', toggleCheckboxFmt);
    add('Insérer', 'Note sur la cellule', 'note', openNoteEditor);
    add('Données', 'Mise en forme conditionnelle', 'cond', () => openCF());
    add('Données', 'Rechercher dans le classeur', 'search', () => openSFind(false), 'Ctrl F');
    add('Données', 'Rechercher et remplacer', 'search', () => openSFind(true), 'Ctrl H');
    add('Données', 'Plages nommées', 'tag', openNames);
    add('Cellules', 'Fusionner ou séparer les cellules', 'merge', toggleMerge);
    add('Cellules', 'Ajouter une décimale', 'grid', () => stepDecimals(1));
    add('Cellules', 'Réduire les décimales', 'grid', () => stepDecimals(-1));
    add('Structure', 'Masquer les lignes sélectionnées', 'eyeOff', () => setHidden('r', true));
    add('Structure', 'Masquer les colonnes sélectionnées', 'eyeOff', () => setHidden('c', true));
    add('Structure', 'Afficher les lignes masquées', 'eye', () => setHidden('r', false));
    add('Structure', 'Afficher les colonnes masquées', 'eye', () => setHidden('c', false));
    if (Plume.available) { add('Plume', 'Écrire une formule avec Plume…', 'feather', () => { openPlume(); setTimeout(() => $('#plumeAsk').focus(), 30); }); add('Plume', 'Vérifier le classeur', 'feather', () => { openPlume(); runPlume('verifier'); }); }
    add('Exporter', 'Exporter en PDF (.pdf)', 'pdf', () => runExport('docpdf'));
    add('Exporter', 'Exporter une copie Excel (.xlsx)', 'download', () => runExport('xlsx').catch(() => toast('Export impossible.', true)));
    add('Exporter', 'Exporter en CSV', 'download', () => runExport('csv'));
  }
  if (v !== 'home') {
    add('Fichier', 'Enregistrer', 'save', () => saveDoc(false), 'Ctrl S');
    add('Fichier', 'Enregistrer sous…', 'saveAs', () => saveDoc(true), 'Ctrl Maj S');
    add('Document', 'Renommer', 'pen', () => { $('#titleIn').focus(); $('#titleIn').select(); });
    add('Document', 'Retour à l\'accueil', 'home', goHome);
  }
  add('Créer', 'Nouveau document', 'doc', () => createDoc('texte'));
  add('Créer', 'Nouveau classeur', 'grid', () => createDoc('feuille'));
  add('Créer', 'Ouvrir un PDF, ou fusionner des PDF et des photos', 'pdf', () => $('#pdfIn').click());
  add('Créer', 'Budget du mois (modèle)', 'grid', () => fromTemplate('budget'));
  add('Créer', 'Suivi des dépenses (modèle)', 'grid', () => fromTemplate('depenses'));
  add('Créer', 'Lettre (modèle)', 'doc', () => fromTemplate('lettre'));
  add('Créer', 'Ouvrir un fichier Word, Excel, CSV ou PDF', 'upload', () => openFromComputer());
  if (v === 'writer' || v === 'sheet') add('Exporter', 'Imprimer ou créer un PDF', 'printer', printDoc, 'Ctrl P');
  add('Préférences', isPhone() ? 'Passer en mode PC' : 'Passer en mode téléphone', isPhone() ? 'monitor' : 'phone', () => setMode(isPhone() ? 'pc' : 'phone'));
  add('Préférences', currentDark() ? 'Passer au thème clair' : 'Passer au thème sombre', currentDark() ? 'sun' : 'moon', toggleTheme);
  add('Préférences', 'Réglages', 'sliders', openSettings);
  add('Préférences', 'Sauvegarder tous mes documents', 'download', backupAll);
  [...state.docs.values()].sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0)).forEach(d => {
    if (state.cur && d.id === state.cur.id) return;
    add('Documents', d.title || 'Sans titre', d.type === 'feuille' ? 'grid' : 'doc', async () => { if (state.cur) { commitEdit(false); await flushSave(); } openDoc(d.id); });
  });
  return L;
}
function openPalette() {
  if (state.view === 'sheet') commitEdit(false);
  $('#palette').hidden = false; $('#palIn').value = ''; renderPalette(); $('#palIn').focus();
}
function closePalette(refocus = true) {
  $('#palette').hidden = true;
  if (!refocus) return;
  if (state.view === 'writer' && !document.body.classList.contains('show-outline')) restoreRange();
  else if (state.view === 'sheet') gw.focus({ preventScroll: true });
}
function renderPalette() {
  const q = norm($('#palIn').value.trim());
  const all = commands();
  palItems = q ? all.filter(c => q.split(/\s+/).every(w => norm(c.label + ' ' + c.group).includes(w))) : all.filter(c => c.group !== 'Documents' || all.indexOf(c) < all.findIndex(x => x.group === 'Documents') + 6);
  if (q) palItems.sort((a, b) => (norm(a.label).startsWith(q) ? 0 : 1) - (norm(b.label).startsWith(q) ? 0 : 1));
  palIdx = 0;
  let g = null, h = '';
  palItems.forEach((c, i) => {
    if (!q && c.group !== g) { g = c.group; h += `<div class="pal-g">${esc(g)}</div>`; }
    h += `<button class="pal-it" role="option" data-i="${i}" aria-selected="${i === 0}">${ic(c.icon)}<span>${q ? highlight(c.label, q) : esc(c.label)}</span>${c.kbd ? `<kbd>${c.kbd}</kbd>` : ''}${q ? `<kbd>${esc(c.group)}</kbd>` : ''}</button>`;
  });
  $('#palList').innerHTML = h || '<p class="empty" style="padding:16px">Aucune commande ne correspond.</p>';
}
function movePal(d) {
  if (!palItems.length) return;
  palIdx = (palIdx + d + palItems.length) % palItems.length;
  $$('.pal-it').forEach(b => b.setAttribute('aria-selected', String(+b.dataset.i === palIdx)));
  const cur = $(`.pal-it[data-i="${palIdx}"]`); if (cur) cur.scrollIntoView({ block: 'nearest' });
}
function runPal(i) { const c = palItems[i]; if (!c) return; closePalette(true); setTimeout(() => c.run(), 0); }
$('#palIn').addEventListener('input', renderPalette);
$('#palIn').addEventListener('keydown', e => {
  if (e.key === 'ArrowDown') { e.preventDefault(); movePal(1); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); movePal(-1); }
  else if (e.key === 'Enter') { e.preventDefault(); runPal(palIdx); }
  else if (e.key === 'Escape') { e.preventDefault(); closePalette(); }
});
$('#palList').addEventListener('click', e => { const b = e.target.closest('.pal-it'); if (b) runPal(+b.dataset.i); });
$('#palette').addEventListener('mousedown', e => { if (e.target === $('#palette')) closePalette(); });
$('#palBtn').addEventListener('click', openPalette);

/* --- raccourcis globaux */
document.addEventListener('keydown', e => {
  const mod = e.ctrlKey || e.metaKey;
  if (mod && (e.key === 'k' || e.key === 'K')) { e.preventDefault(); $('#palette').hidden ? openPalette() : closePalette(); return; }
  if (mod && (e.key === 'p' || e.key === 'P') && state.view === 'pdfv') { e.preventDefault(); printPdf(); return; }
  if (mod && (e.key === 's' || e.key === 'S') && state.cur) { e.preventDefault(); saveDoc(e.shiftKey); return; }
  if (mod && state.view === 'writer' && (e.key === 'f' || e.key === 'F')) { e.preventDefault(); openFind(false); return; }
  if (mod && state.view === 'writer' && (e.key === 'h' || e.key === 'H')) { e.preventDefault(); openFind(true); return; }
  if (mod && state.view === 'sheet' && (e.key === 'f' || e.key === 'F' || e.key === 'h' || e.key === 'H')) { e.preventDefault(); openSFind(/h/i.test(e.key)); return; }
  if (mod && (e.key === 'e' || e.key === 'E') && state.cur) { e.preventDefault(); e.stopPropagation(); commitEdit(false); $('#exportPop').hidden ? openExport() : closeExport(); return; }
  if (e.key === 'Escape') {
    if (selImg) hideImgBar();
    else if (!$('#findBar').hidden && state.view === 'writer') closeFind();
    else if (!$('#wPop').hidden || !$('#linkPop').hidden || !$('#pagePop').hidden || !$('#stylePop').hidden) closeWPops();
    else if (!$('#filterPop').hidden) closeFilter();
    else if (!sDlg.hidden) closeDlg();
    else if (!$('#ddPop').hidden) closeDdPick();
    else if (!$('#sFind').hidden && state.view === 'sheet') closeSFind();
    else if (!$('#exportPop').hidden) closeExport();
    else if (!$('#plume').hidden && $('#palette').hidden && document.activeElement !== cellEd && document.activeElement !== fx) closePlume();
    else if (document.body.classList.contains('show-outline')) document.body.classList.remove('show-outline');
  }
});
$('#q').addEventListener('keydown', e => { if (e.key === 'Enter') { const first = $('.drow'); if (first) first.click(); } });

/* =========================================================
   Démarrage
   ========================================================= */
/* =========================================================
   Réglages et données
   ========================================================= */
async function openSettings() {
  if (state.view === 'sheet') commitEdit(false);
  if (state.cur) await flushSave();
  state.cur = null; show('settings'); renderSettings();
}
function segSet(id, v) { $$('#' + id + ' button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.v === v))); }
async function renderSettings() {
  segSet('segMode', MODE); segSet('segTheme', ls.get('grim.theme', 'auto')); segSet('segProv', AI.provider);
  $('#aiKey').value = AI.key; $('#aiModel').value = AI.model; $('#aiModel').placeholder = AI_DEF[AI.provider];
  $('#aiModels').innerHTML = AI_SUGG[AI.provider].map(m => `<option value="${m}">`).join('');
  $('#appVer').textContent = APP_VERSION;
  const inst = matchMedia('(display-mode: standalone)').matches || navigator.standalone;
  $('#instState').textContent = inst ? 'Grimoire est installé sur cet appareil : il s\'ouvre comme une appli et fonctionne hors connexion.'
    : installEvt ? 'Grimoire n\'est pas encore installé sur cet appareil.'
    : 'Grimoire n\'est pas installé ici. Si le bouton Installer n\'apparaît pas, ouvrez Grimoire dans Chrome, Edge ou Brave (voir ci-dessous).';
  const n = state.docs.size;
  let info = `${n} document${n > 1 ? 's' : ''} sur cet appareil`;
  try { if (navigator.storage && navigator.storage.estimate) { const e = await navigator.storage.estimate(); info += ` · ${(e.usage / 1048576).toFixed(1).replace('.', ',')} Mo utilisés`; } } catch {}
  try { if (navigator.storage && navigator.storage.persisted && await navigator.storage.persisted()) info += ' · stockage protégé contre l\'effacement automatique'; } catch {}
  $('#dataInfo').textContent = info + '.';
}
$('#setBtn').addEventListener('click', () => { state.view === 'settings' ? goHome() : openSettings(); });
$('#segMode').addEventListener('click', e => { const b = e.target.closest('[data-v]'); if (b) setMode(b.dataset.v); });
$('#segTheme').addEventListener('click', e => { const b = e.target.closest('[data-v]'); if (b) { setTheme(b.dataset.v); segSet('segTheme', b.dataset.v); } });
$('#segProv').addEventListener('click', e => { const b = e.target.closest('[data-v]'); if (!b) return; AI.provider = b.dataset.v; AI.model = ''; ls.set('grim.ai', AI); $('#aiRes').textContent = ''; renderSettings(); });
$('#aiKey').addEventListener('input', e => { AI.key = e.target.value.trim(); ls.set('grim.ai', AI); });
$('#aiModel').addEventListener('input', e => { AI.model = e.target.value.trim(); ls.set('grim.ai', AI); });
$('#aiTest').addEventListener('click', async () => {
  const r = $('#aiRes'); r.className = 'set-res'; r.textContent = 'Test en cours…';
  try { const res = await aiChat('Réponds uniquement par le mot : OK'); r.className = 'set-res ok'; r.textContent = `Connexion réussie avec ${provName()} (${aiModel()}) : « ${res.text.trim().slice(0, 30)} »`; }
  catch (e) { r.className = 'set-res err'; r.textContent = aiError(e); }
});
async function backupAll() {
  if (state.cur) { if (state.view === 'sheet') commitEdit(false); await flushSave(); }
  const docs = [];
  for (const { _saved, ...d } of state.docs.values()) {
    if (d.type === 'pdf' && d.pdfId) { const b = await Store.getBlob(d.pdfId); if (b) d.pdfData = bytesToB64(b); }
    docs.push(d);
  }
  const day = new Date().toISOString().slice(0, 10);
  await saveFile(`grimoire-sauvegarde-${day}.json`, JSON.stringify({ app: 'grimoire', format: 1, exportedAt: new Date().toISOString(), docs }));
}
$('#backupBtn').addEventListener('click', backupAll);
$('#restoreBtn').addEventListener('click', () => $('#restoreIn').click());
$('#restoreIn').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    const data = JSON.parse(await f.text());
    if (!data || data.app !== 'grimoire' || !Array.isArray(data.docs)) throw 0;
    let added = 0, kept = 0;
    for (const d of data.docs) {
      if (!d || !d.id || !/^(texte|feuille|pdf|note)$/.test(d.type)) continue;
      const cur = state.docs.get(d.id);
      if (cur && (cur.updatedAt || 0) >= (d.updatedAt || 0)) { kept++; continue; }
      if (d.type === 'pdf') {
        const b64 = d.pdfData || d.pdf; delete d.pdfData; delete d.pdf;
        if (!b64) continue;
        d.pdfId ||= uid(); await Store.putBlob(d.pdfId, b64ToBytes(b64));
      }
      await Store.putRaw(d); added++;
    }
    toast(`${added} document${added > 1 ? 's' : ''} restauré${added > 1 ? 's' : ''}` + (kept ? `, ${kept} déjà à jour.` : '.'));
    renderSettings();
  } catch { toast('Ce fichier n\'est pas une sauvegarde Grimoire valide.', true); }
});

let installEvt = null;
window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); installEvt = e; $$('.install-btn').forEach(b => { b.hidden = false; }); $('#installBanner').hidden = !!ls.get('grim.installDismissed', false); });
document.addEventListener('click', async e => {
  if (!e.target.closest('.install-btn') || !installEvt) return;
  installEvt.prompt(); const r = await installEvt.userChoice.catch(() => null);
  if (r && r.outcome === 'accepted') { $$('.install-btn').forEach(b => { b.hidden = true; }); $('#installBanner').hidden = true; }
  else ls.set('grim.installDismissed', true);
  installEvt = null;
});
window.addEventListener('appinstalled', () => { $('#installBanner').hidden = true; toast('Grimoire est installé.'); });

if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
  const hadCtl = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.register('sw.js').catch(() => {});
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (hadCtl) $('#updBanner').hidden = false; });
}
$('#updBtn').addEventListener('click', () => location.reload());
if ('launchQueue' in window) {
    window.launchQueue.setConsumer(async p => { if (p.files && p.files.length) await importHandles(p.files); });
}

window.addEventListener('error', e => { toast('Un problème est survenu : ' + (e.message || 'erreur inconnue'), true); });
window.addEventListener('unhandledrejection', e => { const m = e.reason && (e.reason.message || e.reason.code); toast('Un problème est survenu' + (m ? ' : ' + m : '.'), true); });
applyMode();
renderHome();
/* =========================================================
   Correction automatique (comme Word) : documents, notes, texte des cellules
   ========================================================= */
const AC_WORDS = (() => {
  const m = {}, add = s => s.trim().split(/\s+/).forEach(p => { const [a, b] = p.split('='); m[a] = b.replace(/_/g, ' '); });
  // apostrophes oubliées
  add(`jai=j’ai cest=c’est nest=n’est quil=qu’il quils=qu’ils quelle=quelle daccord=d’accord aujourdhui=aujourd’hui jusqua=jusqu’à jusquau=jusqu’au lorsquil=lorsqu’il puisquil=puisqu’il presquile=presqu’île quelquun=quelqu’un dabord=d’abord dailleurs=d’ailleurs jen=j’en jy=j’y ny=n’y sil=s’il sils=s’ils cetait=c’était jetais=j’étais javais=j’avais jaurai=j’aurai jaurais=j’aurais jespere=j’espère`);
  // accents oubliés (mots sans ambiguïté)
  add(`deja=déjà voila=voilà tres=très apres=après etre=être meme=même memes=mêmes probleme=problème problemes=problèmes reunion=réunion reunions=réunions numero=numéro numeros=numéros telephone=téléphone general=général generale=générale interessant=intéressant systeme=système premiere=première derniere=dernière deuxieme=deuxième troisieme=troisième quatrieme=quatrième bientot=bientôt plutot=plutôt tot=tôt ete=été evenement=événement evenements=événements reponse=réponse reponses=réponses reseau=réseau securite=sécurité societe=société qualite=qualité realite=réalité activite=activité electricite=électricité propriete=propriété proprietaire=propriétaire proprietaires=propriétaires verifier=vérifier verification=vérification etude=étude etudes=études etat=état equipe=équipe equipes=équipes ecole=école role=rôle hopital=hôpital theatre=théâtre foret=forêt fenetre=fenêtre tete=tête fete=fête noel=Noël aout=août fevrier=février decembre=décembre ca=ça oeuvre=œuvre coeur=cœur soeur=sœur soeurs=sœurs oeil=œil oeuf=œuf oeufs=œufs boeuf=bœuf voeux=vœux noeud=nœud manoeuvre=manœuvre controle=contrôle controles=contrôles diagnostic=diagnostic batiment=bâtiment batiments=bâtiments chateau=château age=âge agee=âgée gateau=gâteau grace=grâce hotel=hôtel hotels=hôtels cote=côté interet=intérêt different=différent differente=différente difference=différence necessaire=nécessaire precedent=précédent prochainement=prochainement repondre=répondre recu=reçu recue=reçue facon=façon garcon=garçon lecon=leçon francais=français francaise=française deplacement=déplacement developpement=développement echeance=échéance echeancier=échéancier etablissement=établissement experience=expérience strategie=stratégie modele=modèle resultat=résultat resultats=résultats reference=référence references=références numerique=numérique methode=méthode periode=période preparer=préparer presentation=présentation priorite=priorité regle=règle regles=règles telecharger=télécharger tache=tâche taches=tâches`);
  // fautes de frappe courantes
  add(`ets=est aps=pas psa=pas puor=pour poru=pour aves=avec acev=avec dnas=dans dasn=dans uen=une qeu=que mias=mais msi=mis cmme=comme comem=comme`);
  ['quelle', 'diagnostic', 'prochainement', 'cote', 'tache', 'taches'].forEach(k => delete m[k]);
  return m;
})();
const AC = { on: ls.get('grim.autocorr', true), mine: ls.get('grim.autocorr.mine', ''), ignore: new Set(), last: null };
function acMineMap() { const m = {}; String(AC.mine).split('\n').forEach(l => { const i = l.indexOf('='); if (i > 0) { const a = l.slice(0, i).trim().toLowerCase(), b = l.slice(i + 1).trim(); if (a && b) m[a] = b; } }); return m; }
let AC_MINE = acMineMap();
const isL = ch => /[\p{L}\p{N}]/u.test(ch || '');
const matchCase = (w, r) => w === w.toUpperCase() && w.length > 1 && w !== w.toLowerCase() ? r.toUpperCase() : w[0] !== w[0].toLowerCase() ? r[0].toUpperCase() + r.slice(1) : r;
// Calcule la correction de la fin du texte (ctx = tout le texte du paragraphe jusqu'au curseur, déclencheur compris)
function acFix(ctx, trig, cell) {
  const body = ctx.slice(0, -1);
  let tail = trig, start = ctx.length - 1; // portion remplacée : [start, fin]
  if (!cell && trig === '"') { const prev = body.slice(-1); tail = (!prev || /[\s(\[« ]/.test(prev)) ? '« ' : ' »'; }
  else if (!cell && trig === "'") tail = '’';
  // mot qui précède (avec élision : l’equipe → l’ + equipe)
  const tm = /([\p{L}’']+)$/u.exec(body), token = tm ? tm[1] : '';
  const before = body.slice(0, body.length - token.length);
  const urlish = /(https?:\/\/|www\.|@)\S*$/i.test(body) || /[\/\\_.\d-]$/.test(before);
  let fixed = token;
  if (token && !urlish && trig !== "'" && !/^(https?|www|ftp|mailto)$/i.test(token) && !AC.ignore.has(token.toLowerCase())) {
    const seg = (/(\p{L}+)$/u.exec(token) || ['', ''])[1], pre = token.slice(0, token.length - seg.length).replace(/'/g, '’');
    let fs = seg;
    const caps = seg.length > 1 && seg === seg.toUpperCase();
    const mine = AC_MINE[token.toLowerCase()] ?? AC_MINE[seg.toLowerCase()];
    if (mine) { fixed = AC_MINE[token.toLowerCase()] ? mine : pre + mine; }
    else {
      const rep = caps ? null : AC_WORDS[seg.toLowerCase()];
      if (rep) fs = (pre && rep.includes('’')) ? seg : matchCase(seg, rep);
      else if (/^\p{Lu}{2}\p{Ll}{2,}$/u.test(seg)) fs = seg[0] + seg.slice(1).toLowerCase(); // DEux MAjuscules
      fixed = pre + fs;
      // majuscule en début de phrase
      if (!cell && /^\p{Ll}/u.test(fixed) && (/^\s*$/.test(before) || /[.!?]\s+$/.test(before)) && !/(^|\s)(etc|cf|ex|p|art|env|M|Mme|n°|vol|chap|tél|av|bd)\.\s+$/i.test(before) && !/\d\.\s+$/.test(before))
        fixed = fixed[0].toUpperCase() + fixed.slice(1);
    }
  }
  // espace insécable avant ; : ! ? (pas pour 14:30 ni http:)
  if (!cell && /[;:!?]/.test(trig) && !/(https?|www\.\S*)$/i.test(body)) {
    const p = body.slice(-1);
    if (p === ' ') { tail = ' ' + trig; start = ctx.length - 2; }
    else if (p && (isL(p) || p === '»' || p === ')') && !(trig === ':' && /\d$/.test(body))) tail = ' ' + trig;
  }
  if (!cell && trig === '.' && body.endsWith('..')) { tail = '…'; start = ctx.length - 3; }
  if (fixed !== token) { const ws = body.length - token.length; return { start: ws, text: fixed + ctx.slice(body.length, start) + tail, word: token }; }
  if (tail !== trig || start !== ctx.length - 1) return { start, text: tail };
  return null;
}
const AC_TRIG = /^[\s .,;:!?)"'…]$/;
function acApply(el, trig) {
  if (!AC.on || AC.busy) return;
  AC.busy = true;
  try { acApply2(el, trig); } finally { AC.busy = false; }
}
function acApply2(el, trig) {
  if (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') {
    const pos = el.selectionStart; if (pos !== el.selectionEnd) return;
    const lineStart = el.value.lastIndexOf('\n', pos - 1) + 1, ctx = el.value.slice(lineStart, pos);
    const cell = el.classList.contains('cell-ed');
    if (cell && /^[=+\-]/.test(el.value)) return; // formule
    const f = acFix(ctx, trig, cell); if (!f) return;
    el.setSelectionRange(lineStart + f.start, pos);
    if (!document.execCommand('insertText', false, f.text)) { el.setRangeText(f.text, lineStart + f.start, pos, 'end'); el.dispatchEvent(new Event('input', { bubbles: true })); }
    AC.last = f.word ? { word: f.word, t: Date.now() } : null;
    return;
  }
  const sel = getSelection(); if (!sel.rangeCount || !sel.isCollapsed) return;
  const node = sel.anchorNode, off = sel.anchorOffset;
  if (!node || node.nodeType !== 3) return;
  const block = node.parentElement.closest('p,li,h1,h2,h3,td,th,blockquote,.note-t,div') || el;
  if (!el.contains(block) || node.parentElement.closest('[contenteditable=false]')) return;
  const r = document.createRange(); r.setStart(block, 0); r.setEnd(node, off);
  const ctx = r.toString(), f = acFix(ctx, trig); if (!f) return;
  const len = ctx.length - f.start; if (len > off) return; // la portion à remplacer déborde du nœud texte (mise en forme au milieu du mot)
  const rr = document.createRange(); rr.setStart(node, off - len); rr.setEnd(node, off);
  sel.removeAllRanges(); sel.addRange(rr);
  document.execCommand('insertText', false, f.text);
  AC.last = f.word ? { word: f.word, t: Date.now() } : null;
}
function acListen(el) {
  el.addEventListener('input', e => {
    if (e.inputType === 'historyUndo' && AC.last && Date.now() - AC.last.t < 8000) { AC.ignore.add(AC.last.word.toLowerCase()); AC.last = null; return; } // Ctrl+Z annule la correction et ne la refait plus
    if (e.inputType === 'insertText' && e.data && AC_TRIG.test(e.data.slice(-1)) && !e.isComposing) acApply(e.target.closest('textarea,input') || el, e.data.slice(-1));
  });
  el.setAttribute('autocorrect', 'on'); if (el.tagName !== 'INPUT') el.setAttribute('autocapitalize', 'sentences');
}
acListen($('#ed')); acListen($('#qnText')); acListen($('#notes'));
// Cellules : correction pour le texte, jamais pour les formules
(() => {
  const ce = $('#cellEd'); acListen(ce);
  const upd = () => { const f = /^[=+\-]/.test(ce.value); ce.setAttribute('autocorrect', f ? 'off' : 'on'); ce.setAttribute('autocapitalize', f ? 'off' : 'sentences'); ce.spellcheck = !f; };
  ce.addEventListener('input', upd); ce.addEventListener('focus', upd);
})();
$('#acOn').checked = AC.on; $('#acMine').value = AC.mine;
$('#acOn').addEventListener('change', e => { AC.on = e.target.checked; ls.set('grim.autocorr', AC.on); toast(AC.on ? 'Correction automatique activée.' : 'Correction automatique désactivée.'); });
$('#acMine').addEventListener('change', e => { AC.mine = e.target.value; ls.set('grim.autocorr.mine', AC.mine); AC_MINE = acMineMap(); toast('Remplacements enregistrés.'); });

Store.init().then(() => { syncUI(); if (SY.cfg) syncNow(); scheduleReminders(); const nid = new URLSearchParams(location.search).get('note'); if (nid) { history.replaceState(null, '', location.pathname); focusNote(nid); } });
})();
