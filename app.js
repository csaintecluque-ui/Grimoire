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

const APP_VERSION = '1.4.0';
const LIB = {
  hf: 'lib/hyperformula.full.min.js',
  xlsx: 'lib/xlsx-style.bundle.js',
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
<h2>Le traitement de texte</h2>
<p>Police, taille, couleur, surlignage, liens, listes à plusieurs niveaux (Tab), cases à cocher, interligne. Cliquez dans un tableau pour faire apparaître ses outils : ajouter des lignes ou des colonnes, fusionner, colorier. <strong>Ctrl F</strong> recherche, <strong>Ctrl H</strong> remplace, <strong>Ctrl Entrée</strong> insère un saut de page. Le bouton Mise en page règle l'orientation, les marges, l'en-tête, le pied de page et les numéros de page.</p>
<h2>Le tableur</h2>
<p>Les formules s'écrivent comme dans Excel en français : <strong>=SOMME(D4:D20)</strong>, <strong>=SI(A1&gt;100;"oui";"non")</strong>, <strong>=SOMME.SI(C:C;"Courses";D:D)</strong>. Pendant une formule, touchez ou cliquez une cellule pour insérer sa référence. Le bouton <strong>ƒx</strong> liste les fonctions avec leur explication.</p>
<p><strong>Recopier :</strong> faites glisser le petit carré en bas à droite de la sélection pour recopier une formule ou prolonger une série (1, 2, 3… ou des dates). <strong>Ctrl D</strong> recopie vers le bas. <strong>Dates :</strong> tapez 22/10/2026 ou 22/10, et =B1-A1 donne le nombre de jours. <strong>Trier et filtrer :</strong> boutons flèches et entonnoir, dans la colonne voulue.</p>
<p><strong>Sur téléphone :</strong> touchez une cellule pour la sélectionner, puis touchez-la à nouveau pour écrire. Une rangée de touches (=, SOMME, ;, parenthèses…) apparaît sous la barre de formule.</p>
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
  ['home', 'writer', 'sheet', 'settings'].forEach(v => { $('#' + v).hidden = v !== view; });
  $('#pageCss').textContent = '';
  const inDoc = view === 'writer' || view === 'sheet';
  $('#app').classList.toggle('in-doc', inDoc);
  $('#crumb').hidden = !inDoc;
  $('#barSpacer').hidden = inDoc;
  $('#exportBtn').hidden = !inDoc;
  $('#printBtn').hidden = !inDoc;
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
      ? `Classeur · ${(d.sheets || [d]).reduce((n, sh) => n + Object.values(sh.cells || {}).filter(c => c.v !== '').length, 0)} cellules${(d.sheets || []).length > 1 ? ' · ' + d.sheets.length + ' feuilles' : ''}`
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
  refreshWriter();
}
const refreshWriter = () => { updateOutline(); updateStats(); updateToolState(); };
const refreshSoon = debounce(() => { updateOutline(); updateStats(); }, 350);

ed.addEventListener('input', () => {
  if (!state.cur) return;
  if (!ed.firstElementChild && !ed.textContent) ed.innerHTML = '<p><br></p>';
  normalizeBlocks();
  state.cur.html = edHtml(); scheduleSave(); refreshSoon();
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
function closeWPops(except) { ['#wPop', '#linkPop', '#pagePop'].forEach(s => { if ($(s) !== except) $(s).hidden = true; }); }
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
  if (e.target.closest('#wPop, #linkPop, #pagePop, #wTools, #tblBar, #linkBar')) return;
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
$('#blockSel').addEventListener('change', e => W.block(e.target.value));
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
  $('#wStatus').innerHTML = `<span><b>${words.toLocaleString('fr-FR')}</b> mots · ${chars.toLocaleString('fr-FR')} caractères · ${mins} min de lecture</span><span>${pg ? '' : '≈ '}${pages} page${pages > 1 ? 's' : ''} A4 ${P.orient === 'landscape' ? 'paysage' : ''}</span>`;
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
  'SERIE.JOUR.OUVRE': 'WORKDAY', 'MOIS.DECALER': 'EDATE', 'FIN.MOIS': 'EOMONTH', 'ARRONDI.AU.MULTIPLE': 'MROUND', 'SUBSTITUE': 'SUBSTITUTE', 'CHERCHE': 'SEARCH', 'TROUVE': 'FIND', 'JOURS': 'DAYS', 'NO.SEMAINE': 'WEEKNUM', 'HEURE': 'HOUR', 'DATEVAL': 'DATEVALUE'
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
  if (n != null) return n;
  const d = parseDate(s);
  return d == null ? s : d;
}
const cellOf = (r, c) => state.cur.cells[addr(r, c)];
const rawOf = (r, c) => { const x = cellOf(r, c); return x ? x.v : ''; };

// Un classeur contient plusieurs feuilles ; cells, rows, cols… désignent toujours la feuille active.
const SHEET_KEYS = ['cells', 'rows', 'cols', 'colW', 'filter', 'frz'];
function ensureSheets(doc) {
  if (!Array.isArray(doc.sheets) || !doc.sheets.length) {
    doc.sheets = [{ id: uid(), name: 'Feuille1', cells: doc.cells || {}, rows: doc.rows || 60, cols: doc.cols || 12, colW: doc.colW || {}, filter: doc.filter, frz: doc.frz }];
    doc.active = 0;
  }
  if (!(doc.active >= 0 && doc.active < doc.sheets.length)) doc.active = 0;
  doc.sheets.forEach((sh, i) => { sh.cells ||= {}; sh.rows ||= 60; sh.cols ||= 12; sh.colW ||= {}; sh.id ||= uid(); sh.name ||= 'Feuille' + (i + 1); });
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
  hf.batch(() => cur.sheets.forEach((sh, i) => hf.setSheetContent(SH.ids[i], sheetData(sh))));
  SH.sid = SH.ids[cur.active];
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
  else if (typeof v === 'number') { txt = f.n === 'date' ? fmtDate(v) : (NF[f.n || ''] || NF['']).format(v); cls = 'num'; }
  else if (typeof v === 'boolean') { txt = v ? 'VRAI' : 'FAUX'; cls = 'ac'; }
  else txt = String(v);
  if (f.b) cls += ' b';
  if (f.i) cls += ' i';
  if (f.u) cls += ' u';
  if (f.w) cls += ' wr';
  if (f.a) cls += ' ' + f.a;
  const sh = [];
  if (f.bt) sh.push('inset 0 1px 0 var(--bd)'); if (f.bb) sh.push('inset 0 -1px 0 var(--bd)');
  if (f.bl) sh.push('inset 1px 0 0 var(--bd)'); if (f.br) sh.push('inset -1px 0 0 var(--bd)');
  return { txt, cls: cls.trim(), v, bg: f.bg ? `var(--bg-${f.bg})` : '', fc: f.fc ? `var(--fc-${f.fc})` : '', bd: sh.join(',') };
}
function renderValues() {
  if (!state.cur || state.view !== 'sheet') return;
  SH.vals = SH.hf ? SH.hf.getSheetValues(SH.sid) : [];
  const cur = state.cur;
  for (let r = 0; r < cur.rows; r++) for (let c = 0; c < cur.cols; c++) {
    const td = SH.cell[r][c], d = display(r, c);
    if (td.textContent !== d.txt) td.textContent = d.txt;
    const keep = td.classList.contains('fz') || td.classList.contains('fzc') ? ' ' + [...td.classList].filter(k => /^fz/.test(k)).join(' ') : '';
    if (td.className !== d.cls + keep) td.className = d.cls + keep;
    const st = d.bg + '|' + d.fc + '|' + d.bd;
    if (td._st !== st) { td._st = st; td.style.backgroundColor = d.bg; td.style.color = d.fc; td.style.boxShadow = d.bd; }
    if (d.txt.length > 12) td.title = d.txt; else if (td.title) td.removeAttribute('title');
  }
  if (cur.frz) applyFreeze();
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
  const R = rng();
  box($('#selBox'), SH.cell[R.r1][R.c1], SH.cell[R.r2][R.c2]);
  $('#selBox').style.display = R.r1 === R.r2 && R.c1 === R.c2 ? 'none' : '';
  box($('#actBox'), SH.cell[SH.sel.r][SH.sel.c], SH.cell[SH.sel.r][SH.sel.c]);
  SH.hl.r.forEach(e => e.classList.remove('hl')); SH.hl.c.forEach(e => e.classList.remove('hl'));
  SH.hl.r = SH.rowTh.slice(R.r1, R.r2 + 1); SH.hl.c = SH.colTh.slice(R.c1, R.c2 + 1);
  SH.hl.r.forEach(e => e.classList.add('hl')); SH.hl.c.forEach(e => e.classList.add('hl'));
  const brc = SH.cell[R.r2][R.c2], tb = $('#grid'), half = isPhone() ? 10 : 5;
  fillH.style.left = (tb.offsetLeft + brc.offsetLeft + brc.offsetWidth - half) + 'px';
  fillH.style.top = (tb.offsetTop + brc.offsetTop + brc.offsetHeight - half) + 'px';
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
  if (rowHidden(r)) { const dir = r >= SH.sel.r ? 1 : -1; let rr = r; while (rr >= 0 && rr < cur.rows && rowHidden(rr)) rr += dir; r = rr >= 0 && rr < cur.rows ? rr : SH.sel.r; }
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
  const flt = SH.hiddenRows ? ` <button class="linkbtn" data-act="clearFilter">Filtre : ${SH.hiddenRows} ligne${SH.hiddenRows > 1 ? 's' : ''} masquée${SH.hiddenRows > 1 ? 's' : ''} · tout afficher</button>` : '';
  $('#sStatus').innerHTML = `<span>${left}${flt}</span><span>${hint}</span>`;
}
function updateSheetTools() {
  const f = (cellOf(SH.sel.r, SH.sel.c) || {}).f || {};
  $$('[data-s="bold"]').forEach(b => b.setAttribute('aria-pressed', String(!!f.b)));
  [['ital', 'i'], ['under', 'u'], ['wrap', 'w']].forEach(([s, k]) => { const b = $(`[data-s="${s}"]`); if (b) b.setAttribute('aria-pressed', String(!!f[k])); });
  const fz = state.cur.frz; $('[data-s="freeze"]').setAttribute('aria-pressed', String(!!(fz && (fz.r || fz.c))));
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
  SH.cell.forEach(row => { if (row[0]) row[0].parentElement.hidden = false; });
  SH.colTh.forEach(th => th.classList.remove('flt'));
  if (F) {
    const cols = Object.keys(F.cols || {}).map(Number);
    cols.forEach(c => SH.colTh[c] && SH.colTh[c].classList.add('flt'));
    for (let r = F.r1; r <= Math.min(F.r2, state.cur.rows - 1); r++) {
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
$('#sStatus').addEventListener('click', e => { if (e.target.closest('[data-act="clearFilter"]')) clearFilters(); });
document.addEventListener('mousedown', e => { if (!$('#filterPop').hidden && !e.target.closest('#filterPop') && !e.target.closest('[data-s="filter"]')) closeFilter(); });

/* --- onglets */
function refreshSheetView() {
  const cur = state.cur;
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
  const sh = from ? { ...clone(from), id: uid(), name: uniqueName(from.name) } : { id: uid(), name: 'Feuille' + n, cells: {}, rows: 60, cols: 12, colW: {} };
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
  if (e.target.closest('#tabMenu, #fmtPop, .tab, [data-s="fill"], [data-s="color"], [data-s="borders"], [data-s="freeze"]')) return;
  closeMini();
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
        else if (/^(smartTag|customXml|ins|moveTo|sdtContent|AlternateContent|Choice)$/.test(n)) {
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
        const target = linkParts ? { push: (css, html) => linkParts.push({ css, html }) } : { push };
        if (n === 't') { let t = c.textContent; if (rp.caps) t = t.toUpperCase(); if (t) target.push(rp.css, esc(t)); }
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
    let html = '';
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
      if (desc(pPr, 'sectPr').length && el.nextElementSibling) emit('<hr data-pb="1">');
    }
    closeLists(); closeQuote();
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
  return { html, page };
}

async function importFile(f) {
  if (state.cur) { if (state.view === 'sheet') commitEdit(false); await flushSave(); }
  const ext = (f.name.split('.').pop() || '').toLowerCase();
  const title = f.name.replace(/\.[^.]+$/, '');
  toast('Import de « ' + title + ' »…');
  try {
    if (ext === 'docx') {
      droppedImages = 0;
      let raw, page;
      try { ({ html: raw, page } = await docxToHtml(await f.arrayBuffer())); }
      catch (err) {
        console.warn('Import direct impossible, repli sur mammoth', err);
        droppedImages = 0;
        await loadScript(LIB.mammoth);
        raw = (await window.mammoth.convertToHtml({ arrayBuffer: await f.arrayBuffer() }, { styleMap: ['u => u', 'strike => s'] })).value;
      }
      const box = document.createElement('div'); box.innerHTML = restoreChecklists(sanitize(raw));
      const html = attachFloats(box).innerHTML;
      await createDoc('texte', page ? { title, html, page } : { title, html });
      toast(droppedImages ? `Document importé. ${droppedImages} image${droppedImages > 1 ? 's' : ''} dans un format non pris en charge.` : 'Document importé.');
    } else if (['xlsx', 'xls', 'ods', 'csv'].includes(ext)) {
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
        if (cl.s && cl.s.font) { if (cl.s.font.bold) fm.b = 1; if (cl.s.font.italic) fm.i = 1; }
        if (v !== '') cells[addr(r, c)] = { v, f: fm };
      }
      const colW = {}; (ws['!cols'] || []).slice(0, maxC + 1).forEach((cw, i) => { if (cw && cw.wpx) colW[i] = Math.round(Math.max(40, Math.min(cw.wpx, 420))); else if (cw && cw.wch) colW[i] = Math.round(Math.max(40, Math.min(cw.wch * 7.5, 420))); });
      return { id: uid(), name: String(name).replace(/[\[\]:*?\/\\']/g, '').slice(0, 31) || 'Feuille', cells, rows: Math.max(60, maxR + 21), cols: Math.max(12, maxC + 2), colW };
      });
      await createDoc('feuille', { title, sheets, active: 0 });
      const more = wb.SheetNames.length > 20 ? ' (20 premières feuilles)' : '';
      toast(cut ? `Classeur importé${more}, limité à 1 000 lignes et 52 colonnes par feuille.` : `Classeur importé : ${sheets.length} feuille${sheets.length > 1 ? 's' : ''}${more}.`);
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
function runsOf(node, fmt = {}) {
  let out = '';
  node.childNodes.forEach(n => {
    if (n.nodeType === 3) {
      const t = n.nodeValue.replace(/\s+/g, ' ');
      if (!t) return;
      out += `<w:r>${rPr(fmt)}<w:t xml:space="preserve">${xmlEsc(t)}</w:t></w:r>`;
    } else if (n.nodeType === 1) {
      const tg = n.tagName;
      if (tg === 'BR') out += '<w:r><w:br/></w:r>';
      else if (tg === 'IMG') out += imgRun(n);
      else if (/^(UL|OL|TABLE)$/.test(tg)) return;
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
    else if (tg === 'P' && !n.textContent.trim() && n.querySelector('img:not([data-wrap])') && !n.style.textAlign) out += para('', runsOf(n), n, { jc: 'center' });
    else out += para(cellFmt ? 'TableText' : '', runsOf(n, cellFmt || {}), n);
  });
  flush();
  return out || '<w:p/>';
}
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
  const body = bodyXml(root);
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
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'].join('');
  zip.file('[Content_Types].xml', ct);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId1" Type="${R}styles" Target="styles.xml"/><Relationship Id="rIdNum" Type="${R}numbering" Target="numbering.xml"/>`
    + (hasHdr ? `<Relationship Id="rIdHdr" Type="${R}header" Target="header1.xml"/>` : '')
    + (hasFtr ? `<Relationship Id="rIdFtr" Type="${R}footer" Target="footer1.xml"/>` : '')
    + MEDIA.map(m => `<Relationship Id="${m.rid}" Type="${R}image" Target="media/${m.name}"/>`).join('')
    + LINKS.map((u, k) => `<Relationship Id="rIdL${k + 1}" Type="${R}hyperlink" Target="${xmlEsc(u)}" TargetMode="External"/>`).join('')
    + '</Relationships>');
  zip.file('word/styles.xml', STYLES);
  zip.file('word/numbering.xml', numberingXml());
  if (hasHdr) zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${D_NS}><w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>${wRun(P.header)}</w:p></w:hdr>`);
  if (hasFtr) zip.file('word/footer1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${D_NS}>${ftrParas}</w:ftr>`);
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xmlEsc(state.cur.title || '')}</dc:title><dc:creator>CSC</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  const sect = `<w:sectPr>${hasHdr ? '<w:headerReference w:type="default" r:id="rIdHdr"/>' : ''}${hasFtr ? '<w:footerReference w:type="default" r:id="rIdFtr"/>' : ''}<w:pgSz w:w="${pgW}" w:h="${pgH}"${land ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${mar}" w:right="${mar}" w:bottom="${mar}" w:left="${mar}" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>`;
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
function xlStyle(f) {
  const s = {}, font = {}, al = {}, bd = {};
  if (f.b) font.bold = true; if (f.i) font.italic = true; if (f.u) font.underline = true;
  if (f.fc && XL_RGB[f.fc]) font.color = { rgb: XL_RGB[f.fc][0] };
  if (Object.keys(font).length) s.font = font;
  if (f.bg && XL_RGB[f.bg]) s.fill = { patternType: 'solid', fgColor: { rgb: XL_RGB[f.bg][1] } };
  if (f.a) al.horizontal = { al: 'left', ac: 'center', ar: 'right' }[f.a];
  if (f.w) { al.wrapText = true; al.vertical = 'top'; }
  if (Object.keys(al).length) s.alignment = al;
  [['bt', 'top'], ['bb', 'bottom'], ['bl', 'left'], ['br', 'right']].forEach(([k, n]) => { if (f[k]) bd[n] = { style: 'thin', color: { rgb: '1B1512' } }; });
  if (Object.keys(bd).length) s.border = bd;
  if (f.n && XL_Z[f.n]) s.numFmt = XL_Z[f.n];
  return Object.keys(s).length ? s : null;
}
async function exportXlsx() {
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
      if ((x.v === '' || x.v == null) && !st) return;
      maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c);
      const v = (vals[p.r] || [])[p.c];
      let cell;
      if (x.v === '' || x.v == null) cell = { t: 's', v: '' };
      else if (String(x.v)[0] === '=') {
        const fv = typeof v === 'number' ? { t: 'n', v } : typeof v === 'boolean' ? { t: 'b', v } : typeof v === 'object' && v ? { t: 'e', v: 0, w: v.value } : { t: 's', v: v == null ? '' : String(v) };
        cell = { ...fv, f: frToExcel(String(x.v).slice(1)) };
      } else {
        const n = parseNum(x.v), d = n == null ? parseDate(x.v) : null;
        cell = n != null ? { t: 'n', v: n } : d != null ? { t: 'n', v: d } : { t: 's', v: String(x.v).replace(/^'/, '') };
      }
      if (f.n && XL_Z[f.n] && cell.t === 'n') cell.z = XL_Z[f.n];
      if (st) cell.s = st;
      ws[X.utils.encode_cell({ r: p.r, c: p.c })] = cell;
    });
    ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
    ws['!cols'] = Array.from({ length: maxC + 1 }, (_, c) => ({ wpx: sh.colW[c] || 104 }));
    X.utils.book_append_sheet(wb, ws, sh.name.slice(0, 31));
  });
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
  $('#printArea').innerHTML = h + '</table>';
  // tableau large : impression en paysage
  $('#pageCss').textContent = `@page{size:A4 ${totalW > 700 ? 'landscape' : 'portrait'};margin:12mm}`;
}
window.addEventListener('beforeprint', buildPrint);
function printDoc() {
  hideImgBar();
  if (state.view === 'writer') applyPageLayout();
  buildPrint();
  try { setTimeout(() => window.print(), 50); }
  catch { toast('Impression indisponible ici : ouvrez Grimoire dans son propre onglet.', true); }
}
$('#printBtn').addEventListener('click', printDoc);

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
  if (mod && state.view === 'writer' && (e.key === 'f' || e.key === 'F')) { e.preventDefault(); openFind(false); return; }
  if (mod && state.view === 'writer' && (e.key === 'h' || e.key === 'H')) { e.preventDefault(); openFind(true); return; }
  if (mod && (e.key === 'e' || e.key === 'E') && state.cur) { e.preventDefault(); commitEdit(false); runExport(state.cur.type === 'feuille' ? 'xlsx' : 'docx').catch(() => toast('Export impossible.', true)); return; }
  if (e.key === 'Escape') {
    if (selImg) hideImgBar();
    else if (!$('#findBar').hidden && state.view === 'writer') closeFind();
    else if (!$('#wPop').hidden || !$('#linkPop').hidden || !$('#pagePop').hidden) closeWPops();
    else if (!$('#filterPop').hidden) closeFilter();
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
