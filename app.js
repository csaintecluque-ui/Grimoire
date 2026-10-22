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

const APP_VERSION = '1.0.1';
const LIB = {
  hf: 'lib/hyperformula.full.min.js',
  xlsx: 'lib/xlsx.full.min.js',
  jszip: 'lib/jszip.min.js',
  mammoth: 'lib/mammoth.browser.min.js'
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
  copy: '<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a1 1 0 00-1-1H5a1 1 0 00-1 1v10a1 1 0 001 1h3"/>',
  phone: '<rect x="7" y="2.5" width="10" height="19" rx="2.2"/><path d="M11 18.5h2"/>',
  monitor: '<rect x="3" y="4" width="18" height="12" rx="1.5"/><path d="M8 20h8M12 16v4"/>',
  sliders: '<path d="M4 6h9M17 6h3M4 12h3M11 12h9M4 18h11M19 18h1"/><circle cx="15" cy="6" r="2"/><circle cx="9" cy="12" r="2"/><circle cx="17" cy="18" r="2"/>',
  image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-9 9"/>',
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
      const r = indexedDB.open('grimoire', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('docs', { keyPath: 'id' });
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
  },
  async putRaw(doc) {
    state.docs.set(doc.id, doc);
    if (this.mode === 'idb') await this.req(this.tx('readwrite').put(clone(doc)));
    else this.writeLocal();
  },
  async remove(id) {
    state.docs.delete(id);
    if (this.mode === 'idb') await this.req(this.tx('readwrite').delete(id));
    else this.writeLocal();
  }
};

let saveTimer = null, saving = false, pendingSave = false;
function setSaveState(s) {
  const el = $('#saveSt'); el.dataset.s = s;
  el.textContent = s === 'ok' ? 'Enregistré' : s === 'dirty' ? 'Modifications…' : 'Non enregistré';
}
function scheduleSave() {
  if (!state.cur) return;
  setSaveState('dirty');
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
<h2>Le tableur</h2>
<p>Les formules s'écrivent comme dans Excel en français : <strong>=SOMME(D4:D20)</strong>, <strong>=SI(A1&gt;100;"oui";"non")</strong>, <strong>=SOMME.SI(C:C;"Courses";D:D)</strong>. Pendant une formule, touchez ou cliquez une cellule pour insérer sa référence. Le bouton <strong>ƒx</strong> liste les fonctions avec leur explication.</p>
<p><strong>Sur téléphone :</strong> touchez une cellule pour la sélectionner, puis touchez-la à nouveau pour écrire. Une rangée de touches (=, SOMME, ;, parenthèses…) apparaît sous la barre de formule.</p>
<h2>Vos données</h2>
<p>Tout est enregistré automatiquement sur cet appareil, et Grimoire fonctionne hors connexion. Pour passer vos documents sur un autre appareil : <strong>Réglages → Sauvegarder tout</strong>, puis <strong>Restaurer</strong> de l'autre côté.</p>
<h2>Plume, l'assistant (facultatif)</h2>
<p>Ajoutez une clé Mistral ou OpenRouter dans les Réglages pour reformuler, corriger ou écrire une formule à partir d'une phrase. Sans clé, tout le reste fonctionne.</p>`
});
const sheetTpl = (title, rows, colW) => {
  const cells = {};
  rows.forEach(([a, v, f]) => { cells[a] = { v: String(v), f: f || {} }; });
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
    return sheetTpl('Suivi des dépenses', rows, { 0: 70, 1: 170, 2: 110, 3: 100, 4: 24, 5: 160, 6: 110 });
  }
};

/* =========================================================
   Navigation
   ========================================================= */
function show(view) {
  state.view = view;
  ['home', 'writer', 'sheet', 'settings'].forEach(v => { $('#' + v).hidden = v !== view; });
  const inDoc = view === 'writer' || view === 'sheet';
  $('#app').classList.toggle('in-doc', inDoc);
  $('#crumb').hidden = !inDoc;
  $('#barSpacer').hidden = inDoc;
  $('#exportBtn').hidden = !inDoc;
  $('#plumeBtn').hidden = !inDoc;
  closePlume(); closeExport();
  document.body.classList.remove('focus', 'show-outline');
  if (inDoc) { $('#titleIn').value = state.cur.title || ''; setSaveState('ok'); }
}
async function goHome() {
  if (state.view === 'sheet') commitEdit(false);
  if (saveTimer || state.cur && !state.cur._saved) await flushSave();
  state.cur = null; show('home'); renderHome();
}
function openDoc(id) {
  const d = state.docs.get(id); if (!d) return;
  state.cur = clone(d); state.cur.id = id; state.cur._saved = true;
  if (d.type === 'feuille') openSheet(); else openWriter();
}
async function createDoc(type, base) {
  const now = Date.now();
  const doc = Object.assign(
    type === 'feuille' ? { type, title: 'Classeur sans titre', cells: {}, rows: 60, cols: 12, colW: {} }
                       : { type, title: 'Document sans titre', html: '<p><br></p>' },
    base || {}, { id: uid(), createdAt: now, updatedAt: now });
  state.cur = doc; state.cur._saved = false;
  try { if (doc.type === 'feuille') openSheet(); else openWriter(); }
  catch (e) { toast('Ouverture impossible : ' + (e.message || e), true); return; }
  try { await Store.save(doc); setSaveState('ok'); }
  catch (e) { setSaveState('err'); toast('Document créé, mais pas encore enregistré' + (e && e.code ? ' (' + e.code + ')' : '') + '. Nouvel essai à la prochaine modification.', true); }
  if (!base) setTimeout(() => { $('#titleIn').focus(); $('#titleIn').select(); }, 60);
}
const fromTemplate = k => createDoc(TPL[k]().type, TPL[k]());

$('#brand').addEventListener('click', goHome);
$('#backBtn').addEventListener('click', goHome);
$$('[data-new]').forEach(b => b.addEventListener('click', () => createDoc(b.dataset.new)));
$$('[data-tpl]').forEach(b => b.addEventListener('click', () => fromTemplate(b.dataset.tpl)));
$('#importCard').addEventListener('click', () => $('#fileIn').click());
$('#titleIn').addEventListener('input', e => { if (!state.cur) return; state.cur.title = e.target.value; scheduleSave(); });
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
  if (d.type === 'feuille') return Object.values(d.cells || {}).map(c => c.v).filter(Boolean).join(' · ');
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
  const all = [...state.docs.values()].sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  $('#nAll').textContent = all.length || '';
  $('#nTxt').textContent = all.filter(d => d.type === 'texte').length || '';
  $('#nSh').textContent = all.filter(d => d.type === 'feuille').length || '';
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
    const meta = d.type === 'feuille'
      ? `Classeur · ${Object.values(d.cells || {}).filter(c => c.v !== '').length} cellules`
      : `${(t.match(/\S+/g) || []).length.toLocaleString('fr-FR')} mots`;
    return `<button class="drow" data-id="${esc(d.id)}" data-type="${d.type}">
      <span class="dico">${ic(d.type === 'feuille' ? 'grid' : 'doc')}</span>
      <span style="min-width:0"><div class="dtitle">${highlight(d.title || 'Sans titre', q)}</div><div class="dsnip">${highlight(snippet(t, q), q) || '<em>Vide</em>'}</div></span>
      <span class="dmeta">${relDate(d.updatedAt)}<br>${meta}</span></button>`;
  }).join('');
}
$('#docs').addEventListener('click', e => { const b = e.target.closest('.drow'); if (b) openDoc(b.dataset.id); });
$('#q').addEventListener('input', debounce(e => { state.q = e.target.value; renderHome(); }, 120));
$('#chips').addEventListener('click', e => {
  const b = e.target.closest('.chip'); if (!b) return;
  state.filter = b.dataset.f; $$('.chip').forEach(c => c.setAttribute('aria-pressed', String(c === b))); renderHome();
});

/* =========================================================
   Nettoyage HTML (collage, import)
   ========================================================= */
const KEEP = { P: 1, H1: 1, H2: 1, H3: 1, STRONG: 1, EM: 1, U: 1, UL: 1, OL: 1, LI: 1, TABLE: 1, THEAD: 1, TBODY: 1, TR: 1, TD: 1, TH: 1, BR: 1, BLOCKQUOTE: 1, HR: 1, A: 1 };
const MAP = { B: 'strong', I: 'em', H4: 'h3', H5: 'h3', H6: 'h3', DIV: 'p', TFOOT: 'tbody' };
const DROP = new Set(['SCRIPT', 'STYLE', 'META', 'LINK', 'TITLE', 'IFRAME', 'OBJECT', 'EMBED', 'SVG', 'VIDEO', 'AUDIO', 'CANVAS', 'FORM', 'INPUT', 'BUTTON', 'SELECT', 'TEXTAREA', 'NOSCRIPT', 'TEMPLATE', 'HEAD']);
let droppedImages = 0;
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
      const w = ch.style && ch.style.width;
      for (const at of [...ch.attributes]) ch.removeAttribute(at.name);
      ch.setAttribute('src', src.replace(/\s+/g, '')); ch.setAttribute('alt', '');
      if (w && /^\d{1,3}(\.\d+)?%$/.test(w)) ch.style.width = w;
      continue;
    }
    const align = ch.style && ch.style.textAlign;
    const bold = ch.style && /^(bold|[6-9]00)$/.test(ch.style.fontWeight);
    const ital = ch.style && ch.style.fontStyle === 'italic';
    const href = tag === 'A' ? ch.getAttribute('href') : null;
    let el = ch;
    if (tag === 'B' && ch.style && ch.style.fontWeight === 'normal') { clean(ch, doc); ch.replaceWith(...ch.childNodes); continue; }
    if (MAP[tag]) { el = doc.createElement(MAP[tag]); while (ch.firstChild) el.appendChild(ch.firstChild); ch.replaceWith(el); }
    else if (!KEEP[tag]) {
      clean(ch, doc);
      if (tag === 'SPAN' && (bold || ital)) {
        let w = doc.createElement(bold ? 'strong' : 'em'); while (ch.firstChild) w.appendChild(ch.firstChild);
        if (bold && ital) { const i2 = doc.createElement('em'); while (w.firstChild) i2.appendChild(w.firstChild); w.appendChild(i2); }
        ch.replaceWith(w);
      } else ch.replaceWith(...ch.childNodes);
      continue;
    }
    for (const a of [...el.attributes]) el.removeAttribute(a.name);
    if (align && /^(left|center|right|justify)$/.test(align) && /^(P|H1|H2|H3|LI|TD|TH|BLOCKQUOTE)$/.test(el.tagName)) el.style.textAlign = align;
    if (el.tagName === 'A') {
      if (href && /^https?:\/\//i.test(href)) { el.setAttribute('href', href); el.setAttribute('target', '_blank'); el.setAttribute('rel', 'noopener'); }
      else { clean(el, doc); el.replaceWith(...el.childNodes); continue; }
    }
    clean(el, doc);
    // un <p> qui ne contient que des blocs est déballé
    if (el.tagName === 'P' && el.querySelector('p,h1,h2,h3,ul,ol,table,blockquote')) el.replaceWith(...el.childNodes);
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
  $('#wScroll').scrollTop = 0;
  savedRange = null;
  if (ls.get('grim.outline', true) === false) document.body.classList.add('no-outline');
  refreshWriter();
}
const refreshWriter = () => { updateOutline(); updateStats(); updateToolState(); };
const refreshSoon = debounce(() => { updateOutline(); updateStats(); }, 350);

ed.addEventListener('input', () => {
  if (!state.cur) return;
  if (!ed.firstElementChild && !ed.textContent) ed.innerHTML = '<p><br></p>';
  state.cur.html = ed.innerHTML.replace(/ class="sel"/g, ''); scheduleSave(); refreshSoon();
});
ed.addEventListener('paste', e => {
  const cd = e.clipboardData; if (!cd) return;
  e.preventDefault();
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
  if (ed.contains(r.commonAncestorContainer)) { savedRange = r.cloneRange(); updateToolState(); if (!$('#plume').hidden && Plume.mode === 'texte') Plume.ctx(); }
});
function restoreRange() {
  ed.focus({ preventScroll: true });
  if (savedRange) { const s = getSelection(); s.removeAllRanges(); s.addRange(savedRange); }
}
function exec(cmd, val) {
  if (document.activeElement !== ed) restoreRange();
  document.execCommand(cmd, false, val);
  state.cur.html = ed.innerHTML; scheduleSave(); refreshSoon(); updateToolState();
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
  set('outline', !document.body.classList.contains('no-outline'));
  set('focus', document.body.classList.contains('focus'));
}
function insertTable(rows = 3, cols = 3) {
  const th = Array.from({ length: cols }, (_, i) => `<th>Colonne ${i + 1}</th>`).join('');
  const tr = `<tr>${'<td><br></td>'.repeat(cols)}</tr>`.repeat(rows - 1);
  exec('insertHTML', `<table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table><p><br></p>`);
}
const W = {
  undo: () => exec('undo'), redo: () => exec('redo'),
  bold: () => exec('bold'), italic: () => exec('italic'), underline: () => exec('underline'),
  ul: () => exec('insertUnorderedList'), ol: () => exec('insertOrderedList'),
  left: () => exec('justifyLeft'), center: () => exec('justifyCenter'), right: () => exec('justifyRight'), justify: () => exec('justifyFull'),
  table: () => insertTable(), hr: () => exec('insertHorizontalRule'),
  image: () => { if (document.activeElement !== ed) restoreRange(); $('#imgIn').click(); },
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
$('#blockSel').addEventListener('change', e => W.block(e.target.value));

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
  bar.hidden = false;
  bar.style.left = Math.max(8, Math.min(r.left - host.left + r.width / 2 - bar.offsetWidth / 2, host.width - bar.offsetWidth - 8)) + 'px';
  bar.style.top = Math.max(8, r.top - host.top - bar.offsetHeight - 8) + 'px';
});
$('#wScroll').addEventListener('scroll', () => { if (selImg) hideImgBar(); }, { passive: true });
$('#imgBar').addEventListener('mousedown', e => e.preventDefault());
$('#imgBar').addEventListener('click', e => {
  const b = e.target.closest('[data-iw]'); if (!b || !selImg) return;
  if (b.dataset.iw === 'del') selImg.remove(); else { selImg.style.width = b.dataset.iw + '%'; selImg.classList.remove('sel'); }
  hideImgBar();
  state.cur.html = ed.innerHTML.replace(/ class="sel"/g, ''); scheduleSave();
});

function updateOutline() {
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
  const pages = Math.max(1, Math.ceil(words / 480));
  $('#wStatus').innerHTML = `<span><b>${words.toLocaleString('fr-FR')}</b> mots · ${chars.toLocaleString('fr-FR')} caractères · ${mins} min de lecture</span><span>≈ ${pages} page${pages > 1 ? 's' : ''} A4</span>`;
}

/* =========================================================
   Tableur
   ========================================================= */
const FR2EN = {
  'SOMME': 'SUM', 'MOYENNE': 'AVERAGE', 'SI': 'IF', 'ARRONDI': 'ROUND', 'ARRONDI.SUP': 'ROUNDUP', 'ARRONDI.INF': 'ROUNDDOWN',
  'NB': 'COUNT', 'NBVAL': 'COUNTA', 'NB.VIDE': 'COUNTBLANK', 'NB.SI': 'COUNTIF', 'NB.SI.ENS': 'COUNTIFS', 'SOMME.SI': 'SUMIF',
  'SOMME.SI.ENS': 'SUMIFS', 'MOYENNE.SI': 'AVERAGEIF', 'RECHERCHEV': 'VLOOKUP', 'RECHERCHEH': 'HLOOKUP', 'EQUIV': 'MATCH',
  'ET': 'AND', 'OU': 'OR', 'NON': 'NOT', 'CONCATENER': 'CONCATENATE', 'CONCAT': 'CONCATENATE', 'AUJOURDHUI': 'TODAY',
  'MAINTENANT': 'NOW', 'ANNEE': 'YEAR', 'MOIS': 'MONTH', 'JOUR': 'DAY', 'SIERREUR': 'IFERROR', 'SOMMEPROD': 'SUMPRODUCT',
  'ENT': 'INT', 'PUISSANCE': 'POWER', 'RACINE': 'SQRT', 'GAUCHE': 'LEFT', 'DROITE': 'RIGHT', 'STXT': 'MID', 'NBCAR': 'LEN',
  'MAJUSCULE': 'UPPER', 'MINUSCULE': 'LOWER', 'SUPPRESPACE': 'TRIM', 'TEXTE': 'TEXT', 'ECARTYPE': 'STDEV', 'MEDIANE': 'MEDIAN',
  'GRANDE.VALEUR': 'LARGE', 'PETITE.VALEUR': 'SMALL', 'RANG': 'RANK', 'ESTVIDE': 'ISBLANK', 'ESTNUM': 'ISNUMBER',
  'ESTTEXTE': 'ISTEXT', 'VRAI': 'TRUE', 'FAUX': 'FALSE', 'JOURSEM': 'WEEKDAY', 'NB.JOURS.OUVRES': 'NETWORKDAYS',
  'SERIE.JOUR.OUVRE': 'WORKDAY', 'MOIS.DECALER': 'EDATE', 'FIN.MOIS': 'EOMONTH', 'ARRONDI.AU.MULTIPLE': 'MROUND', 'SUBSTITUE': 'SUBSTITUTE', 'CHERCHE': 'SEARCH', 'TROUVE': 'FIND'
};
const EN2FR = Object.fromEntries(Object.entries(FR2EN).filter(([f]) => f !== 'CONCAT').map(([f, e]) => [e, f]));
function outsideStrings(f, fn) { return f.split('"').map((p, i) => i % 2 ? p : fn(p)).join('"'); }
const names = (f, dict) => outsideStrings(f.replace(/AUJOURD['’]HUI\s*\(/gi, 'AUJOURDHUI('), p =>
  p.replace(/([A-Za-zÀ-ÿ][A-Za-z0-9À-ÿ._]*)(?=\s*\()/g, m => dict[norm(m).toUpperCase()] || m));
const bools = f => outsideStrings(f, p => p.replace(/\b(VRAI|FAUX|TRUE|FALSE)\b(?!\s*\()/gi, m => (/^(VRAI|TRUE)$/i.test(m) ? 'TRUE' : 'FALSE') + '()'));
const frNamesToEn = f => names(bools(f), FR2EN);
const enNamesToFr = f => outsideStrings(names(f, EN2FR), p => p.replace(/\b(VRAI|FAUX)\(\)/g, '$1'));
const frToExcel = f => outsideStrings(frNamesToEn(f), p => p.replace(/(\d),(\d)/g, '$1.$2').replace(/;/g, ','));
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
const numToRaw = n => String(+n.toPrecision(15)).replace('.', ',');
const NF = {
  '': new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 6 }),
  dec2: new Intl.NumberFormat('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
  eur: new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }),
  pct: new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 2 })
};

const SH = { keyTap: 0, caret: null, tapAgain: false, hf: null, sid: 0, fr: true, vals: [], cell: [], colTh: [], rowTh: [], sel: { r: 0, c: 0 }, anc: { r: 0, c: 0 }, editing: null, hist: [], fut: [], drag: null, refIns: null, hl: { r: [], c: [] } };
const gw = $('#gridWrap'), cellEd = $('#cellEd'), fx = $('#fx');

async function ensureHF() {
  if (SH.hf) return true;
  await loadScript(LIB.hf);
  const HFm = window.HyperFormula;
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
  const n = parseNum(s);
  return n == null ? s : n;
}
const cellOf = (r, c) => state.cur.cells[addr(r, c)];
const rawOf = (r, c) => { const x = cellOf(r, c); return x ? x.v : ''; };

async function openSheet() {
  const cur = state.cur;
  cur.cells ||= {}; cur.rows ||= 60; cur.cols ||= 12; cur.colW ||= {};
  SH.hist = []; SH.fut = []; SH.sel = { r: 0, c: 0 }; SH.anc = { r: 0, c: 0 }; SH.editing = null;
  show('sheet');
  buildGrid();
  $('#sStatus').textContent = 'Chargement du moteur de calcul…';
  try { await ensureHF(); }
  catch { toast('Moteur de calcul injoignable : les formules s\'afficheront sans être calculées. Vérifiez la connexion.', true); }
  if (state.cur !== cur) return;
  loadHF(); renderValues(); select(0, 0);
  gw.focus({ preventScroll: true });
}
function loadHF() {
  if (!SH.hf) return;
  const cur = state.cur, data = [];
  for (let r = 0; r < cur.rows; r++) { const row = []; for (let c = 0; c < cur.cols; c++) { const x = cur.cells[addr(r, c)]; row.push(x ? toHF(x.v) : null); } data.push(row); }
  SH.hf.setSheetContent(SH.sid, data);
}
function buildGrid() {
  const cur = state.cur;
  let h = '<colgroup><col class="rh">';
  let total = 46;
  for (let c = 0; c < cur.cols; c++) { const w = cur.colW[c] || 104; total += w; h += `<col style="width:${w}px">`; }
  h += '</colgroup><thead><tr><th class="corner" scope="col"><span class="sr">Lignes et colonnes</span></th>';
  for (let c = 0; c < cur.cols; c++) h += `<th class="ch" scope="col" data-c="${c}">${COLN(c)}<span class="rz" data-c="${c}"></span></th>`;
  h += '</tr></thead><tbody>';
  for (let r = 0; r < cur.rows; r++) {
    h += `<tr><th scope="row" data-r="${r}">${r + 1}</th>`;
    for (let c = 0; c < cur.cols; c++) h += `<td data-r="${r}" data-c="${c}"></td>`;
    h += '</tr>';
  }
  const t = $('#grid'); t.innerHTML = h + '</tbody>'; t.style.width = total + 'px';
  SH.cell = []; $$('tbody tr', t).forEach((tr, r) => { SH.cell[r] = [...tr.children].slice(1); });
  SH.colTh = $$('thead th.ch', t); SH.rowTh = $$('tbody th', t); SH.hl = { r: [], c: [] };
}
function display(r, c) {
  const x = cellOf(r, c), f = (x && x.f) || {};
  let v = SH.hf ? (SH.vals[r] || [])[c] : (x ? (parseNum(x.v) ?? x.v) : null);
  if (!SH.hf && x && String(x.v)[0] === '=') v = x.v;
  let txt = '', cls = '';
  if (v == null || v === '') txt = '';
  else if (typeof v === 'object') { txt = v.value || '#ERREUR'; cls = 'err'; }
  else if (typeof v === 'number') { txt = (NF[f.n || ''] || NF['']).format(v); cls = 'num'; }
  else if (typeof v === 'boolean') { txt = v ? 'VRAI' : 'FAUX'; cls = 'ac'; }
  else txt = String(v);
  if (f.b) cls += ' b';
  if (f.a) cls += ' ' + f.a;
  return { txt, cls: cls.trim(), v };
}
function renderValues() {
  if (!state.cur || state.view !== 'sheet') return;
  SH.vals = SH.hf ? SH.hf.getSheetValues(SH.sid) : [];
  const cur = state.cur;
  for (let r = 0; r < cur.rows; r++) for (let c = 0; c < cur.cols; c++) {
    const td = SH.cell[r][c], d = display(r, c);
    if (td.textContent !== d.txt) td.textContent = d.txt;
    if (td.className !== d.cls) td.className = d.cls;
    if (d.txt.length > 12) td.title = d.txt; else if (td.title) td.removeAttribute('title');
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
  const R = rng();
  box($('#selBox'), SH.cell[R.r1][R.c1], SH.cell[R.r2][R.c2]);
  $('#selBox').style.display = R.r1 === R.r2 && R.c1 === R.c2 ? 'none' : '';
  box($('#actBox'), SH.cell[SH.sel.r][SH.sel.c], SH.cell[SH.sel.r][SH.sel.c]);
  SH.hl.r.forEach(e => e.classList.remove('hl')); SH.hl.c.forEach(e => e.classList.remove('hl'));
  SH.hl.r = SH.rowTh.slice(R.r1, R.r2 + 1); SH.hl.c = SH.colTh.slice(R.c1, R.c2 + 1);
  SH.hl.r.forEach(e => e.classList.add('hl')); SH.hl.c.forEach(e => e.classList.add('hl'));
}
function scrollToCell(r, c) {
  const td = SH.cell[r][c], t = $('#grid');
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
  SH.sel = { r, c }; if (!extend) SH.anc = { r, c };
  drawSel(); scrollToCell(r, c); updateFx(); updateSelStatus(); updateSheetTools();
  if (!$('#plume').hidden && Plume.mode === 'feuille') Plume.ctx();
}
function updateFx() {
  if (SH.editing) return;
  $('#fxRef').textContent = rngRef(rng());
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
  $('#sStatus').innerHTML = `<span>${left}</span><span>${hint}</span>`;
}
function updateSheetTools() {
  const f = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {};
  $$('[data-s="bold"]').forEach(b => b.setAttribute('aria-pressed', String(!!f.b)));
  ['al', 'ac', 'ar'].forEach(k => $(`[data-s="${k}"]`).setAttribute('aria-pressed', String(f.a === k)));
  $$('.seg [data-n]').forEach(b => b.setAttribute('aria-pressed', String((f.n || '') === b.dataset.n)));
}

/* --- historique */
const snap = () => JSON.stringify({ cells: state.cur.cells, rows: state.cur.rows, cols: state.cur.cols, colW: state.cur.colW });
function pushHist() { SH.hist.push(snap()); if (SH.hist.length > 80) SH.hist.shift(); SH.fut = []; }
function restore(s) {
  const o = JSON.parse(s); const dims = o.rows !== state.cur.rows || o.cols !== state.cur.cols || JSON.stringify(o.colW) !== JSON.stringify(state.cur.colW);
  Object.assign(state.cur, o);
  if (dims) buildGrid();
  loadHF(); afterChange(); select(Math.min(SH.sel.r, state.cur.rows - 1), Math.min(SH.sel.c, state.cur.cols - 1));
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
  if (key === 'b') val = first.b ? 0 : 1;
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
  const td = SH.cell[SH.sel.r][SH.sel.c], t = $('#grid');
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
function endEdit() { SH.editing = null; SH.refIns = null; cellEd.hidden = true; showKeys(false); }
function commitEdit(move) {
  if (!SH.editing) return;
  const v = SH.editing === 'fx' ? fx.value : cellEd.value;
  const changed = v !== String(rawOf(SH.sel.r, SH.sel.c));
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
const FNS = [
  ['SOMME', 'SOMME(plage)', 'Additionne les nombres d\'une plage'],
  ['MOYENNE', 'MOYENNE(plage)', 'Moyenne des nombres'],
  ['SI', 'SI(test;si_vrai;si_faux)', 'Choisit une valeur selon une condition'],
  ['NB', 'NB(plage)', 'Compte les cellules qui contiennent un nombre'],
  ['NB.SI', 'NB.SI(plage;critère)', 'Compte les cellules qui remplissent un critère'],
  ['SOMME.SI', 'SOMME.SI(plage;critère;plage_somme)', 'Additionne selon un critère'],
  ['MAX', 'MAX(plage)', 'Plus grande valeur'],
  ['MIN', 'MIN(plage)', 'Plus petite valeur'],
  ['ARRONDI', 'ARRONDI(nombre;décimales)', 'Arrondit à n décimales'],
  ['RECHERCHEV', 'RECHERCHEV(valeur;table;colonne;FAUX)', 'Cherche une valeur dans la 1re colonne d\'un tableau'],
  ['SIERREUR', 'SIERREUR(valeur;sinon)', 'Remplace une erreur par une autre valeur'],
  ['AUJOURDHUI', 'AUJOURDHUI()', 'Date du jour'],
  ['ET', 'ET(test1;test2)', 'Vrai si tous les tests sont vrais'],
  ['OU', 'OU(test1;test2)', 'Vrai si au moins un test est vrai'],
  ['CONCATENER', 'CONCATENER(texte1;texte2)', 'Assemble des textes'],
  ['ENT', 'ENT(nombre)', 'Partie entière'],
  ['ABS', 'ABS(nombre)', 'Valeur absolue']
];
const KEYS = [['=', '='], ['SOMME(', 'SOMME', 'fn'], ['SI(', 'SI', 'fn'], ['(', '('], [')', ')'], [';', ';'], [':', ':'], ['+', '+'], ['-', '−'], ['*', '×'], ['/', '÷'], ['"', '" "'], ['%', '%'], ['$', '$'], ['>', '>'], ['<', '<'], ['__fn', 'ƒ Fonctions', 'fn'], ['__no', 'Annuler', 'no'], ['__ok', 'Valider', 'ok']];
$('#fxkRow').innerHTML = KEYS.map(([v, l, c]) => `<button class="fxk ${c || ''}" data-k="${esc(v)}">${esc(l)}</button>`).join('');
$('#fnList').innerHTML = FNS.map(([n, sig, d]) => `<button class="fn-it" data-fn="${n}"><b>${esc(sig)}</b><small>${esc(d)}</small></button>`).join('');
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
gw.addEventListener('dblclick', e => { const p = tdAt(e); if (p && !SH.editing) { select(p.r, p.c); startEdit('', true); } });

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
  const mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
  if (mv) {
    e.preventDefault();
    let { r, c } = SH.sel;
    if (mod) { r = mv[0] > 0 ? state.cur.rows - 1 : mv[0] < 0 ? 0 : r; c = mv[1] > 0 ? state.cur.cols - 1 : mv[1] < 0 ? 0 : c; }
    else { r += mv[0]; c += mv[1]; }
    return select(r, c, e.shiftKey);
  }
  if (k === 'Enter') { e.preventDefault(); return select(SH.sel.r + (e.shiftKey ? -1 : 1), SH.sel.c); }
  if (k === 'Tab') { e.preventDefault(); return select(SH.sel.r, SH.sel.c + (e.shiftKey ? -1 : 1)); }
  if (k === 'F2') { e.preventDefault(); return startEdit('', true); }
  if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); return clearRange(); }
  if (k === 'Home') { e.preventDefault(); return select(mod ? 0 : SH.sel.r, 0, e.shiftKey); }
  if (mod && (k === 'z' || k === 'Z')) { e.preventDefault(); return e.shiftKey ? redo() : undo(); }
  if (mod && (k === 'y' || k === 'Y')) { e.preventDefault(); return redo(); }
  if (mod && (k === 'b' || k === 'B')) { e.preventDefault(); return setFmt('b'); }
  if (mod && (k === 'a' || k === 'A')) { e.preventDefault(); SH.anc = { r: state.cur.rows - 1, c: state.cur.cols - 1 }; return select(0, 0, true); }
  if (!mod && !e.altKey && k.length === 1) { e.preventDefault(); startEdit(k); }
});

/* --- copier / coller */
function rangeTSV() {
  const R = rng(), lines = [];
  for (let r = R.r1; r <= R.r2; r++) {
    const row = [];
    for (let c = R.c1; c <= R.c2; c++) { const v = (SH.vals[r] || [])[c]; row.push(v == null ? '' : typeof v === 'number' ? numToRaw(v) : typeof v === 'object' ? (v.value || '') : String(v)); }
    lines.push(row.join('\t'));
  }
  return lines.join('\n');
}
document.addEventListener('copy', e => { if (state.view !== 'sheet' || SH.editing || document.activeElement !== gw) return; e.preventDefault(); e.clipboardData.setData('text/plain', rangeTSV()); toast('Copié'); });
document.addEventListener('cut', e => { if (state.view !== 'sheet' || SH.editing || document.activeElement !== gw) return; e.preventDefault(); e.clipboardData.setData('text/plain', rangeTSV()); clearRange(); });
document.addEventListener('paste', e => {
  if (state.view !== 'sheet' || SH.editing || document.activeElement !== gw) return;
  e.preventDefault();
  const txt = (e.clipboardData.getData('text/plain') || '').replace(/\r/g, '');
  const rows = txt.split('\n'); if (rows.length > 1 && rows[rows.length - 1] === '') rows.pop();
  const R = rng();
  pushHist();
  const grew = ensureSize(R.r1 + rows.length - 1, R.c1 + Math.max(...rows.map(l => l.split('\t').length)) - 1);
  if (grew) loadHF();
  batch(() => rows.forEach((line, i) => line.split('\t').forEach((v, j) => applyRaw(R.r1 + i, R.c1 + j, v))));
  afterChange();
  SH.anc = { r: R.r1, c: R.c1 };
  select(R.r1 + rows.length - 1, R.c1 + rows[0].split('\t').length - 1, true);
});

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
function structural(op) {
  commitEdit(false);
  if (!SH.hf) return toast('Le moteur de calcul est nécessaire pour insérer ou supprimer des lignes.', true);
  const cur = state.cur, { r, c } = SH.sel;
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
  const ser = SH.hf.getSheetSerialized(SH.sid), cells = {};
  ser.forEach((row, ri) => (row || []).forEach((v, ci) => {
    if (v == null || v === '') return;
    let raw = typeof v === 'number' ? numToRaw(v) : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v);
    if (raw[0] === '=') raw = '=' + (SH.fr ? enNamesToFr(raw.slice(1)) : excelToFr(raw.slice(1)));
    cells[addr(ri, ci)] = { v: raw, f: {} };
  }));
  Object.entries(nf).forEach(([k, f]) => { (cells[k] ||= { v: '', f: {} }).f = f; });
  cur.cells = cells;
  buildGrid(); afterChange(); select(Math.min(r, cur.rows - 1), Math.min(c, cur.cols - 1));
}
const S = {
  undo: () => { commitEdit(false); undo(); }, redo: () => { commitEdit(false); redo(); },
  bold: () => setFmt('b'), al: () => setFmt('a', 'al'), ac: () => setFmt('a', 'ac'), ar: () => setFmt('a', 'ar'),
  sum: autoSum, clear: () => { commitEdit(false); clearRange(); },
  rowIns: () => structural('rowIns'), rowDel: () => structural('rowDel'), colIns: () => structural('colIns'), colDel: () => structural('colDel')
};
$('#sheet .toolbar').addEventListener('mousedown', e => { if (e.target.closest('button')) e.preventDefault(); });
$('#sheet .toolbar').addEventListener('click', e => {
  const b = e.target.closest('[data-s]'); if (b) { S[b.dataset.s](); if (!SH.editing) gw.focus({ preventScroll: true }); return; }
  const n = e.target.closest('[data-n]'); if (n) { setFmt('n', n.dataset.n); gw.focus({ preventScroll: true }); }
});
window.addEventListener('resize', debounce(() => { if (state.view === 'sheet') { drawSel(); if (SH.editing === 'cell') placeEditor(); } }, 120));

/* =========================================================
   Import
   ========================================================= */
$('#fileIn').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) importFile(f); });
async function importFile(f) {
  if (state.cur) { if (state.view === 'sheet') commitEdit(false); await flushSave(); }
  const ext = (f.name.split('.').pop() || '').toLowerCase();
  const title = f.name.replace(/\.[^.]+$/, '');
  toast('Import de « ' + title + ' »…');
  try {
    if (ext === 'docx') {
      await loadScript(LIB.mammoth);
      const res = await window.mammoth.convertToHtml({ arrayBuffer: await f.arrayBuffer() });
      droppedImages = 0;
      const html = sanitize(res.value);
      await createDoc('texte', { title, html });
      toast(droppedImages ? `Document importé. ${droppedImages} image${droppedImages > 1 ? 's' : ''} dans un format non pris en charge.` : 'Document importé.');
    } else if (['xlsx', 'xls', 'ods', 'csv'].includes(ext)) {
      await loadScript(LIB.xlsx);
      const X = window.XLSX;
      const wb = ext === 'csv' ? X.read(await f.text(), { type: 'string', FS: guessSep(await f.text()) }) : X.read(await f.arrayBuffer(), { cellFormula: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const ref = X.utils.decode_range(ws['!ref'] || 'A1');
      const maxR = Math.min(ref.e.r, 999), maxC = Math.min(ref.e.c, 51);
      const cells = {};
      for (let r = 0; r <= maxR; r++) for (let c = 0; c <= maxC; c++) {
        const cl = ws[X.utils.encode_cell({ r, c })]; if (!cl) continue;
        let v, fm = {};
        if (cl.f) v = '=' + excelToFr(cl.f);
        else if (cl.t === 'n' && cl.z && X.SSF.is_date(cl.z)) v = cl.w || numToRaw(cl.v);
        else if (cl.t === 'n') { v = numToRaw(cl.v); const z = String(cl.z || ''); if (z.includes('%')) fm.n = 'pct'; else if (z.includes('€')) fm.n = 'eur'; else if (/0\.00/.test(z)) fm.n = 'dec2'; }
        else if (cl.t === 'b') v = cl.v ? 'VRAI' : 'FAUX';
        else if (cl.t === 'd') v = cl.w || String(cl.v);
        else v = cl.v == null ? '' : String(cl.v);
        if (cl.f && cl.z) { const z = String(cl.z); if (z.includes('%')) fm.n = 'pct'; else if (z.includes('€')) fm.n = 'eur'; }
        if (v !== '') cells[addr(r, c)] = { v, f: fm };
      }
      const colW = {}; (ws['!cols'] || []).slice(0, maxC + 1).forEach((cw, i) => { if (cw && cw.wpx) colW[i] = Math.round(Math.max(40, Math.min(cw.wpx, 420))); else if (cw && cw.wch) colW[i] = Math.round(Math.max(40, Math.min(cw.wch * 7.5, 420))); });
      await createDoc('feuille', { title, cells, rows: Math.max(60, maxR + 21), cols: Math.max(12, maxC + 2), colW });
      toast(ref.e.r > 999 || ref.e.c > 51 ? 'Classeur importé (limité aux 1 000 premières lignes et 52 colonnes).' : (wb.SheetNames.length > 1 ? `Classeur importé : première feuille « ${wb.SheetNames[0]} » seulement.` : 'Classeur importé.'));
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
const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', csv: 'text/csv', md: 'text/markdown', html: 'text/html', json: 'application/json' };
async function saveFile(filename, data) {
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
const safeName = t => (String(t || 'Sans titre').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() || 'Sans titre').slice(0, 120);

const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
function runsOf(node, fmt = {}) {
  let out = '';
  node.childNodes.forEach(n => {
    if (n.nodeType === 3) {
      const t = n.nodeValue.replace(/\s+/g, ' ');
      if (!t) return;
      const pr = (fmt.b ? '<w:b/>' : '') + (fmt.i ? '<w:i/>' : '') + (fmt.link ? '<w:color w:val="A35E2C"/>' : '') + (fmt.u ? '<w:u w:val="single"/>' : '');
      out += `<w:r>${pr ? '<w:rPr>' + pr + '</w:rPr>' : ''}<w:t xml:space="preserve">${xmlEsc(t)}</w:t></w:r>`;
    } else if (n.nodeType === 1) {
      const tg = n.tagName;
      if (tg === 'BR') out += '<w:r><w:br/></w:r>';
      else if (tg === 'IMG') out += imgRun(n);
      else if (/^(UL|OL|TABLE)$/.test(tg)) return;
      else out += runsOf(n, { ...fmt, b: fmt.b || /^(STRONG|B|TH)$/.test(tg), i: fmt.i || /^(EM|I)$/.test(tg), u: fmt.u || tg === 'U', link: fmt.link || tg === 'A' });
    }
  });
  return out;
}
function para(style, runs, el, extra = '') {
  const al = el && el.style && el.style.textAlign;
  const jc = al ? `<w:jc w:val="${{ left: 'left', center: 'center', right: 'right', justify: 'both' }[al] || 'left'}"/>` : '';
  const ppr = (style ? `<w:pStyle w:val="${style}"/>` : '') + extra + jc;
  return `<w:p>${ppr ? '<w:pPr>' + ppr + '</w:pPr>' : ''}${runs}</w:p>`;
}
function listXml(list, level) {
  let out = '', i = 0;
  [...list.children].forEach(li => {
    if (li.tagName !== 'LI') return;
    i++;
    const mark = list.tagName === 'OL' ? `${i}.` : (level % 2 ? '–' : '•');
    const ind = `<w:ind w:left="${360 + level * 360 + 360}" w:hanging="360"/>`;
    out += para('ListParagraph', `<w:r><w:t xml:space="preserve">${mark}\t</w:t></w:r>` + runsOf(li), li, `<w:tabs><w:tab w:val="left" w:pos="${360 + level * 360 + 360}"/></w:tabs>` + ind);
    [...li.children].filter(c => /^(UL|OL)$/.test(c.tagName)).forEach(sub => { out += listXml(sub, level + 1); });
  });
  return out;
}
function tableXml(tb) {
  const rows = [...tb.querySelectorAll('tr')];
  const n = Math.max(1, ...rows.map(r => r.children.length));
  const w = Math.floor(9070 / n);
  const b = s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="C5B8A8"/>`;
  let out = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(b).join('')}</w:tblBorders><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${'<w:gridCol w:w="' + w + '"/>'.repeat(n)}</w:tblGrid>`;
  rows.forEach(tr => {
    out += '<w:tr>';
    const cells = [...tr.children];
    for (let i = 0; i < n; i++) {
      const cel = cells[i];
      const th = cel && cel.tagName === 'TH';
      out += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${th ? '<w:shd w:val="clear" w:color="auto" w:fill="EAD9B8"/>' : ''}</w:tcPr>${para('TableText', cel ? runsOf(cel, { b: th }) : '', cel)}</w:tc>`;
    }
    out += '</w:tr>';
  });
  return out + '</w:tbl>';
}
function bodyXml(root) {
  let out = '', inline = null;
  const flush = () => { if (inline) { out += para('', runsOf(inline)); inline = null; } };
  [...root.childNodes].forEach(n => {
    if (n.nodeType === 3) { if (n.nodeValue.trim()) { inline ||= document.createElement('p'); inline.appendChild(n.cloneNode()); } return; }
    if (n.nodeType !== 1) return;
    const tg = n.tagName;
    if (/^(STRONG|EM|U|A|B|I|SPAN|BR|IMG)$/.test(tg)) { inline ||= document.createElement('p'); inline.appendChild(n.cloneNode(true)); return; }
    flush();
    if (tg === 'H1') out += para('Heading1', runsOf(n), n);
    else if (tg === 'H2') out += para('Heading2', runsOf(n), n);
    else if (tg === 'H3') out += para('Heading3', runsOf(n), n);
    else if (tg === 'BLOCKQUOTE') { if (n.querySelector('p')) [...n.children].forEach(c => { out += para('Quote', runsOf(c), c); }); else out += para('Quote', runsOf(n), n); }
    else if (tg === 'UL' || tg === 'OL') out += listXml(n, 0);
    else if (tg === 'TABLE') out += tableXml(n) + para('', '');
    else if (tg === 'HR') out += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="C9804F"/></w:pBdr></w:pPr></w:p>';
    else out += para('', runsOf(n), n);
  });
  flush();
  return out || '<w:p/>';
}
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const D_NS = W_NS + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
let MEDIA = [];
function imgRun(n) {
  const m = MEDIA[+n.getAttribute('data-m')]; if (!m) return '';
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
</w:styles>`;
async function exportDocx() {
  await loadScript(LIB.jszip);
  const root = document.createElement('div'); root.innerHTML = sanitize(state.cur.html);
  const zip = new window.JSZip();
  const add = zip.file.bind(zip); zip.file = (n, d, o) => add(n, d, { createFolders: false, ...(o || {}) });
  MEDIA = [];
  const MAXW = 5760000;
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
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' + MEDIA.map(m => `<Relationship Id="${m.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.name}"/>`).join('') + '</Relationships>');
  zip.file('word/styles.xml', STYLES);
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xmlEsc(state.cur.title || '')}</dc:title><dc:creator>CSC</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${D_NS}><w:body>${bodyXml(root)}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`);
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
  const inl = n => [...n.childNodes].map(c => c.nodeType === 3 ? c.nodeValue : c.nodeType !== 1 ? '' : c.tagName === 'BR' ? '  \n' : /^(STRONG|B)$/.test(c.tagName) ? `**${inl(c)}**` : /^(EM|I)$/.test(c.tagName) ? `*${inl(c)}*` : c.tagName === 'A' ? `[${inl(c)}](${c.getAttribute('href')})` : c.tagName === 'IMG' ? '' : /^(UL|OL)$/.test(c.tagName) ? '' : inl(c)).join('');
  const md = [...root.children].map(n => {
    const t = n.tagName;
    if (/^H[1-3]$/.test(t)) return '#'.repeat(+t[1]) + ' ' + inl(n).trim();
    if (t === 'UL' || t === 'OL') return [...n.children].map((li, i) => (t === 'OL' ? `${i + 1}. ` : '- ') + inl(li).trim()).join('\n');
    if (t === 'BLOCKQUOTE') return '> ' + inl(n).trim();
    if (t === 'HR') return '---';
    if (t === 'TABLE') { const rows = [...n.querySelectorAll('tr')].map(tr => '| ' + [...tr.children].map(c => inl(c).trim().replace(/\|/g, '\\|')).join(' | ') + ' |'); if (rows.length) rows.splice(1, 0, '|' + ' --- |'.repeat(n.querySelector('tr').children.length)); return rows.join('\n'); }
    return inl(n).trim();
  }).filter(Boolean).join('\n\n');
  return saveFile(safeName(state.cur.title) + '.md', md + '\n');
}
async function exportXlsx() {
  commitEdit(false);
  await loadScript(LIB.xlsx);
  const X = window.XLSX, cur = state.cur, ws = {};
  let maxR = 0, maxC = 0;
  const Z = { eur: '#,##0.00 "€"', pct: '0.00%', dec2: '#,##0.00' };
  Object.entries(cur.cells).forEach(([k, x]) => {
    if (x.v === '' || x.v == null) return;
    const p = parseAddr(k); if (!p) return;
    maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c);
    const v = (SH.vals[p.r] || [])[p.c], f = x.f || {};
    let cell;
    if (String(x.v)[0] === '=') {
      const fv = typeof v === 'number' ? { t: 'n', v } : typeof v === 'boolean' ? { t: 'b', v } : typeof v === 'object' && v ? { t: 'e', v: 0, w: v.value } : { t: 's', v: v == null ? '' : String(v) };
      cell = { ...fv, f: frToExcel(String(x.v).slice(1)) };
    } else {
      const n = parseNum(x.v);
      cell = n != null ? { t: 'n', v: n } : { t: 's', v: String(x.v).replace(/^'/, '') };
    }
    if (f.n && Z[f.n] && cell.t === 'n') cell.z = Z[f.n];
    ws[X.utils.encode_cell({ r: p.r, c: p.c })] = cell;
  });
  ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  ws['!cols'] = Array.from({ length: maxC + 1 }, (_, i) => ({ wpx: cur.colW[i] || 104 }));
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, 'Feuille1');
  const buf = X.write(wb, { type: 'array', bookType: 'xlsx' });
  await saveFile(safeName(cur.title) + '.xlsx', new Blob([buf]));
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
      let s = v == null ? '' : typeof v === 'number' ? numToRaw(v) : typeof v === 'object' ? (v.value || '') : typeof v === 'boolean' ? (v ? 'VRAI' : 'FAUX') : String(v);
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
  let h = `<h1 class="print-title">${esc(cur.title || '')}</h1><table class="print-grid">`;
  for (let r = 0; r <= maxR; r++) { h += '<tr>'; for (let c = 0; c <= maxC; c++) { const d = display(r, c); h += `<td class="${d.cls}">${esc(d.txt)}</td>`; } h += '</tr>'; }
  $('#printArea').innerHTML = h + '</table>';
}
window.addEventListener('beforeprint', buildPrint);
function printDoc() { hideImgBar(); buildPrint(); setTimeout(() => window.print(), 50); }

/* --- menu Exporter */
let delArmed = false;
function openExport() {
  const p = $('#exportPop'); const t = state.cur.type;
  delArmed = false;
  p.innerHTML = (t === 'texte'
    ? `<button data-x="docx">${ic('doc')}Word (.docx)<small>Ctrl E</small></button><button data-x="md">${ic('type')}Markdown (.md)</button><button data-x="html">${ic('download')}Page web (.html)</button>`
    : `<button data-x="xlsx">${ic('grid')}Excel (.xlsx)<small>Ctrl E</small></button><button data-x="csv">${ic('download')}CSV pour Excel (.csv)</button>`)
    + `<button data-x="print">${ic('printer')}Imprimer ou PDF<small>Ctrl P</small></button>`
    + `<hr><button data-x="dup">${ic('copy')}Dupliquer</button><button data-x="del" class="danger">${ic('trash')}<span>Supprimer le document</span></button>`;
  p.hidden = false;
}
function closeExport() { $('#exportPop').hidden = true; }
const runExport = k => ({ docx: exportDocx, md: exportMd, html: exportHtml, xlsx: exportXlsx, csv: exportCsv, print: printDoc }[k] || (() => {}))();
$('#exportBtn').addEventListener('click', e => { e.stopPropagation(); $('#exportPop').hidden ? openExport() : closeExport(); });
$('#exportPop').addEventListener('click', async e => {
  e.stopPropagation();
  const b = e.target.closest('[data-x]'); if (!b) return;
  const k = b.dataset.x;
  if (k === 'del') {
    if (!delArmed) { delArmed = true; b.querySelector('span').textContent = 'Confirmer la suppression'; return; }
    const id = state.cur.id; closeExport(); clearTimeout(saveTimer); saveTimer = null; state.cur = null;
    try { await Store.remove(id); toast('Document supprimé.'); } catch { toast('Suppression impossible pour le moment.', true); }
    show('home'); renderHome(); return;
  }
  closeExport();
  if (k === 'dup') { commitEdit(false); await flushSave(); const c = clone(state.cur); delete c.id; delete c._saved; c.title = (c.title || 'Sans titre') + ' (copie)'; return createDoc(c.type, c); }
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
    add('Insérer', 'Tableau 3 × 3', 'table', () => insertTable(3, 3));
    add('Insérer', 'Tableau 4 × 5', 'table', () => insertTable(5, 4));
    add('Insérer', 'Séparateur', 'hr', W.hr);
    add('Insérer', 'Image', 'image', W.image);
    add('Insérer', 'Date du jour', 'pen', () => exec('insertText', new Date().toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })));
    add('Affichage', 'Mode concentration', 'focus', W.focus);
    add('Affichage', 'Afficher ou masquer le plan', 'outline', W.outline);
    if (Plume.available) { add('Plume', 'Reformuler la sélection', 'feather', () => { openPlume(); runPlume('reformuler'); }); add('Plume', 'Corriger la sélection', 'feather', () => { openPlume(); runPlume('corriger'); }); add('Plume', 'Résumer le document', 'feather', () => { openPlume(); runPlume('resumer'); }); add('Plume', 'Rédiger avec Plume…', 'feather', () => { openPlume(); setTimeout(() => $('#plumeAsk').focus(), 30); }); }
    add('Exporter', 'Exporter en Word (.docx)', 'download', () => runExport('docx').catch(() => toast('Export impossible.', true)), 'Ctrl E');
    add('Exporter', 'Exporter en Markdown (.md)', 'download', () => runExport('md'));
  }
  if (v === 'sheet') {
    add('Cellules', 'Somme automatique', 'sigma', autoSum);
    add('Cellules', 'Format monétaire (€)', 'grid', () => setFmt('n', 'eur'));
    add('Cellules', 'Format pourcentage', 'grid', () => setFmt('n', 'pct'));
    add('Cellules', 'Format deux décimales', 'grid', () => setFmt('n', 'dec2'));
    add('Cellules', 'Format standard', 'grid', () => setFmt('n', ''));
    add('Cellules', 'Gras', 'type', () => setFmt('b'), 'Ctrl B');
    add('Structure', 'Insérer une ligne au-dessus', 'rowIns', () => structural('rowIns'));
    add('Structure', 'Supprimer la ligne', 'rowDel', () => structural('rowDel'));
    add('Structure', 'Insérer une colonne à gauche', 'colIns', () => structural('colIns'));
    add('Structure', 'Supprimer la colonne', 'colDel', () => structural('colDel'));
    if (Plume.available) { add('Plume', 'Écrire une formule avec Plume…', 'feather', () => { openPlume(); setTimeout(() => $('#plumeAsk').focus(), 30); }); add('Plume', 'Vérifier le classeur', 'feather', () => { openPlume(); runPlume('verifier'); }); }
    add('Exporter', 'Exporter en Excel (.xlsx)', 'download', () => runExport('xlsx').catch(() => toast('Export impossible.', true)), 'Ctrl E');
    add('Exporter', 'Exporter en CSV', 'download', () => runExport('csv'));
  }
  if (v !== 'home') {
    add('Document', 'Renommer', 'pen', () => { $('#titleIn').focus(); $('#titleIn').select(); });
    add('Document', 'Retour à l\'accueil', 'home', goHome);
  }
  add('Créer', 'Nouveau document', 'doc', () => createDoc('texte'));
  add('Créer', 'Nouveau classeur', 'grid', () => createDoc('feuille'));
  add('Créer', 'Budget du mois (modèle)', 'grid', () => fromTemplate('budget'));
  add('Créer', 'Suivi des dépenses (modèle)', 'grid', () => fromTemplate('depenses'));
  add('Créer', 'Lettre (modèle)', 'doc', () => fromTemplate('lettre'));
  add('Créer', 'Importer un fichier Word, Excel ou CSV', 'upload', () => $('#fileIn').click());
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
  if (mod && (e.key === 's' || e.key === 'S') && state.cur) { e.preventDefault(); commitEdit(false); flushSave().then(() => toast('Enregistré')); return; }
  if (mod && (e.key === 'e' || e.key === 'E') && state.cur) { e.preventDefault(); commitEdit(false); runExport(state.cur.type === 'feuille' ? 'xlsx' : 'docx').catch(() => toast('Export impossible.', true)); return; }
  if (e.key === 'Escape') {
    if (selImg) hideImgBar();
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
  const docs = [...state.docs.values()].map(({ _saved, ...d }) => d);
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
      if (!d || !d.id || !/^(texte|feuille)$/.test(d.type)) continue;
      const cur = state.docs.get(d.id);
      if (cur && (cur.updatedAt || 0) >= (d.updatedAt || 0)) { kept++; continue; }
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
  window.launchQueue.setConsumer(async p => { for (const h of (p.files || [])) { try { importFile(await h.getFile()); } catch {} } });
}

window.addEventListener('error', e => { toast('Un problème est survenu : ' + (e.message || 'erreur inconnue'), true); });
window.addEventListener('unhandledrejection', e => { const m = e.reason && (e.reason.message || e.reason.code); toast('Un problème est survenu' + (m ? ' : ' + m : '.'), true); });
applyMode();
renderHome();
Store.init();
})();
