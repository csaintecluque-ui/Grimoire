// Grimoire — application de bureau (Windows, Linux).
// La même appli que la version web (dossier app/, copié par preparer.js), servie
// depuis l'ordinateur par le protocole app://grimoire/ : hors connexion, et des
// données qui ne dépendent plus d'un navigateur.
'use strict';
const { app, BrowserWindow, protocol, ipcMain, dialog, shell, Menu, net } = require('electron');
const path = require('path');
const fs = require('fs');
const fsp = fs.promises;
const { fileURLToPath } = require('url');
const { execFile } = require('child_process');

const WEB = path.join(__dirname, 'app');
const ORIGINE = 'app://grimoire';
const DEPOT = 'csaintecluque-ui/Grimoire';
const FICHIERS_OK = /\.(docx|xlsx|xlsm|xls|ods|csv|txt|md|pdf)$/i;
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.webmanifest': 'application/manifest+json',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon', '.ttf': 'font/ttf', '.pfb': 'application/octet-stream',
  '.bcmap': 'application/octet-stream', '.wasm': 'application/wasm'
};

// profil séparé (essais) : GRIMOIRE_PROFIL=/chemin
if (process.env.GRIMOIRE_PROFIL) app.setPath('userData', process.env.GRIMOIRE_PROFIL);

protocol.registerSchemesAsPrivileged([{
  scheme: 'app',
  privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true, codeCache: true }
}]);

/* ---------- réglages de l'appli de bureau (taille de fenêtre, premier lancement, fichiers autorisés) */
const REGLAGES = () => path.join(app.getPath('userData'), 'bureau.json');
let reglages = {};
function lireReglages() { try { reglages = JSON.parse(fs.readFileSync(REGLAGES(), 'utf8')) || {}; } catch { reglages = {}; } }
function ecrireReglages() { try { fs.writeFileSync(REGLAGES(), JSON.stringify(reglages)); } catch {} }

// Seuls les fichiers que l'utilisateur a choisis (double-clic, Ouvrir, Enregistrer sous) sont lisibles et modifiables.
const autorises = new Set();
function autoriser(p) {
  autorises.add(p);
  const l = (reglages.autorises || []).filter(x => x !== p); l.push(p);
  reglages.autorises = l.slice(-500); ecrireReglages();
}
const estAutorise = p => typeof p === 'string' && path.isAbsolute(p) && autorises.has(path.normalize(p));

/* ---------- fichiers passés au lancement (double-clic, « Ouvrir avec ») */
function fichiersDe(argv, dossier) {
  const out = [];
  for (let a of argv.slice(1)) {
    if (!a || a.startsWith('-') || a === '.' || /electron(\.exe)?$/i.test(a) || a === __dirname || a === app.getAppPath()) continue;
    if (a.startsWith('file://')) { try { a = fileURLToPath(a); } catch { continue; } }
    const p = path.normalize(path.resolve(dossier || process.cwd(), a));
    try { if (FICHIERS_OK.test(p) && fs.statSync(p).isFile()) { out.push(p); autoriser(p); } } catch {}
  }
  return out;
}
let enAttente = [];
let fenetre = null;

function envoyerFichiers(liste) {
  if (!liste.length) return;
  if (fenetre && !fenetre.webContents.isLoading()) fenetre.webContents.send('fichiers:ouvrir', liste);
  else enAttente.push(...liste);
}

/* ---------- fenêtre */
function creerFenetre() {
  const f = reglages.fenetre || {};
  fenetre = new BrowserWindow({
    width: f.width || 1280, height: f.height || 860, x: f.x, y: f.y, minWidth: 360, minHeight: 480,
    title: 'Grimoire', backgroundColor: '#1B1512', show: false, autoHideMenuBar: true,
    icon: path.join(WEB, 'icons', 'icon-512.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, sandbox: true, spellcheck: true, plugins: true }
  });
  if (f.maximized) fenetre.maximize();
  fenetre.once('ready-to-show', () => fenetre.show());

  const memoriser = () => {
    if (!fenetre || fenetre.isMinimized()) return;
    reglages.fenetre = fenetre.isMaximized() ? { ...(reglages.fenetre || {}), maximized: true } : { ...fenetre.getBounds(), maximized: false };
    ecrireReglages();
  };
  fenetre.on('resize', memoriser); fenetre.on('move', memoriser);

  // Fermeture : Grimoire finit d'enregistrer le document ouvert avant de quitter.
  let pret = false;
  fenetre.on('close', e => {
    if (pret) return;
    e.preventDefault();
    const fin = () => { pret = true; if (fenetre && !fenetre.isDestroyed()) fenetre.close(); };
    const t = setTimeout(fin, 4000);
    fenetre.webContents.executeJavaScript('window.__grimoireFlush ? window.__grimoireFlush() : null', true)
      .catch(() => {}).finally(() => { clearTimeout(t); fin(); });
  });
  fenetre.on('closed', () => { fenetre = null; });

  // Liens : vers le navigateur ; l'appli ne quitte jamais app://grimoire/
  const versLeNavigateur = url => { if (/^(https?:|mailto:)/i.test(url)) shell.openExternal(url); };
  fenetre.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('blob:' + ORIGINE) || url.startsWith(ORIGINE + '/')) return { action: 'allow' };
    versLeNavigateur(url); return { action: 'deny' };
  });
  fenetre.webContents.on('will-navigate', (e, url) => { if (!url.startsWith(ORIGINE + '/')) { e.preventDefault(); versLeNavigateur(url); } });

  // Clic droit : orthographe, couper, copier, coller (l'appli garde ses propres menus, sur les cellules par exemple)
  fenetre.webContents.on('context-menu', (e, p) => {
    const m = [];
    if (p.misspelledWord) {
      p.dictionarySuggestions.slice(0, 6).forEach(s => m.push({ label: s, click: () => fenetre.webContents.replaceMisspelling(s) }));
      if (!p.dictionarySuggestions.length) m.push({ label: 'Aucune suggestion', enabled: false });
      m.push({ label: 'Ajouter au dictionnaire', click: () => fenetre.webContents.session.addWordToSpellCheckerDictionary(p.misspelledWord) }, { type: 'separator' });
    }
    if (p.isEditable) m.push({ role: 'cut', label: 'Couper' }, { role: 'copy', label: 'Copier' }, { role: 'paste', label: 'Coller' }, { type: 'separator' }, { role: 'selectAll', label: 'Tout sélectionner' });
    else if (p.selectionText.trim()) m.push({ role: 'copy', label: 'Copier' });
    if (p.linkURL && /^(https?:|mailto:)/i.test(p.linkURL)) m.push({ type: 'separator' }, { label: 'Ouvrir le lien dans le navigateur', click: () => shell.openExternal(p.linkURL) });
    if (m.length) Menu.buildFromTemplate(m).popup({ window: fenetre });
  });

  // Ctrl+Maj+I / F12 : outils de développement (pour diagnostiquer un problème)
  fenetre.webContents.on('before-input-event', (e, i) => {
    if (i.type === 'keyDown' && (i.key === 'F12' || (i.control && i.shift && i.key.toLowerCase() === 'i'))) { fenetre.webContents.toggleDevTools(); e.preventDefault(); }
  });

  fenetre.webContents.on('did-finish-load', () => {
    if (!reglages.premierLancement && !process.env.GRIMOIRE_SANS_ACCUEIL) { reglages.premierLancement = Date.now(); ecrireReglages(); setTimeout(proposerParDefaut, 1200); }
  });
  fenetre.loadURL(ORIGINE + '/index.html');
}

/* ---------- Grimoire, appli par défaut des fichiers */
async function proposerParDefaut() {
  if (!fenetre) return;
  const win = process.platform === 'win32';
  const r = await dialog.showMessageBox(fenetre, {
    type: 'question', buttons: ['Oui', 'Plus tard'], defaultId: 0, cancelId: 1, title: 'Grimoire', noLink: true,
    message: 'Ouvrir les documents Word, les classeurs Excel, les CSV et les PDF avec Grimoire ?',
    detail: win
      ? 'La page « Applications par défaut » de Windows va s\'ouvrir : cliquez sur « Définir par défaut ». Windows réserve ce clic à l\'utilisateur.\n\nPossible plus tard : Réglages → Application.'
      : 'Grimoire deviendra l\'appli par défaut de ces fichiers, avec leurs icônes. Les applis d\'avant sont notées pour pouvoir y revenir (Réglages → Application).'
  });
  if (r.response === 0) {
    const res = await parDefaut();
    if (!win) dialog.showMessageBox(fenetre, { type: res.ok ? 'info' : 'warning', title: 'Grimoire', message: res.message, buttons: ['OK'] });
  }
}
function scriptLinux(args) {
  const script = app.isPackaged ? path.join(process.resourcesPath, 'grimoire-installation.sh') : path.join(__dirname, '..', 'linux', 'grimoire-installation.sh');
  return new Promise(res => execFile('bash', [script, ...args], { timeout: 120000 }, (err, out, errOut) => res({ ok: !err, sortie: String(out || '') + String(errOut || '') })));
}
async function parDefaut() {
  if (process.platform === 'win32') {
    let build = 0; try { build = parseInt(require('os').release().split('.')[2], 10) || 0; } catch {}
    await shell.openExternal(build >= 22000 ? 'ms-settings:defaultapps?registeredAppUser=Grimoire' : 'ms-settings:defaultapps');
    return { ok: true, message: build >= 22000 ? 'Dans la page Grimoire des Paramètres : « Définir par défaut ».' : 'Dans les Paramètres : choisissez Grimoire pour les .docx, .xlsx, .csv et .pdf.' };
  }
  if (process.platform === 'linux') {
    const r = await scriptLinux(['--bureau=grimoire.desktop']);
    return { ok: r.ok, message: r.ok ? 'C\'est fait : les .docx, .xlsx, .csv et .pdf s\'ouvrent avec Grimoire. Fermez et rouvrez le gestionnaire de fichiers pour voir les icônes.' : 'Ça n\'a pas marché : ' + r.sortie.trim().split('\n').slice(-3).join(' ') };
  }
  return { ok: false, message: 'Pas disponible sur ce système.' };
}
async function retablirAvant() {
  if (process.platform !== 'linux') return { ok: false, message: 'Sous Windows : Paramètres → Applications par défaut.' };
  const r = await scriptLinux(['--retirer']);
  return { ok: r.ok, message: r.ok ? 'Les applis d\'avant ouvrent de nouveau ces fichiers.' : 'Ça n\'a pas marché.' };
}

/* ---------- mises à jour : les nouvelles versions publiées sur GitHub */
let majPrete = null;
const plusRecente = (a, b) => { const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number); for (let i = 0; i < 3; i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); } return false; };
function signalerMaj(info) { majPrete = info; if (fenetre) fenetre.webContents.send('maj:prete', info); }
async function verifierMaj() {
  if (!app.isPackaged) return;
  if (process.platform === 'win32') {
    try {
      const { autoUpdater } = require('electron-updater');
      autoUpdater.autoDownload = true; autoUpdater.autoInstallOnAppQuit = true;
      autoUpdater.on('update-downloaded', i => signalerMaj({ version: i.version, systeme: 'win' }));
      autoUpdater.on('error', () => {});
      await autoUpdater.checkForUpdates();
    } catch {}
    return;
  }
  if (process.platform === 'linux') {
    try {
      const r = await net.fetch(`https://api.github.com/repos/${DEPOT}/releases/latest`, { headers: { Accept: 'application/vnd.github+json' } });
      if (!r.ok) return;
      const rel = await r.json(), v = String(rel.tag_name || '').replace(/^v/, '');
      const deb = (rel.assets || []).find(a => /amd64\.deb$/.test(a.name));
      if (!deb || !plusRecente(v, app.getVersion())) return;
      const cible = path.join(app.getPath('downloads'), deb.name);
      if (!fs.existsSync(cible)) {
        const d = await net.fetch(deb.browser_download_url); if (!d.ok) return;
        await fsp.writeFile(cible + '.part', Buffer.from(await d.arrayBuffer())); await fsp.rename(cible + '.part', cible);
      }
      signalerMaj({ version: v, systeme: 'linux', fichier: cible });
    } catch {}
  }
}
async function installerMaj() {
  if (!majPrete) return;
  if (majPrete.systeme === 'win') {
    if (fenetre) await fenetre.webContents.executeJavaScript('window.__grimoireFlush ? window.__grimoireFlush() : null', true).catch(() => {});
    require('electron-updater').autoUpdater.quitAndInstall();
  } else if (majPrete.fichier) {
    // Ouvre l'installateur de paquets du système (mot de passe demandé), comme un double-clic sur le .deb
    await shell.openPath(majPrete.fichier);
  }
}

/* ---------- échanges avec la page */
function ipc() {
  ipcMain.handle('fichiers:attente', () => enAttente.splice(0));
  ipcMain.handle('fichier:lire', async (_e, p) => {
    if (!estAutorise(p)) throw new Error('EACCES: fichier non autorisé');
    const [bytes, st] = await Promise.all([fsp.readFile(p), fsp.stat(p)]);
    return { bytes: new Uint8Array(bytes), mtime: st.mtimeMs };
  });
  ipcMain.handle('fichier:ecrire', async (_e, p, bytes) => {
    if (!estAutorise(p)) throw new Error('EACCES: fichier non autorisé');
    // écriture sûre : fichier temporaire puis remplacement, l'original n'est jamais à moitié écrit
    const tmp = path.join(path.dirname(p), '.' + path.basename(p) + '.grimoire-' + process.pid);
    await fsp.writeFile(tmp, Buffer.from(bytes));
    try { await fsp.rename(tmp, p); } catch (e) { await fsp.rm(tmp, { force: true }).catch(() => {}); throw e; }
    return true;
  });
  const dossierDepart = () => reglages.dernierDossier && fs.existsSync(reglages.dernierDossier) ? reglages.dernierDossier : app.getPath('documents');
  ipcMain.handle('dialogue:enregistrer', async (_e, o = {}) => {
    const nom = String(o.suggestedName || 'Sans titre').replace(/[\\/:*?"<>|]+/g, ' ');
    const r = await dialog.showSaveDialog(fenetre, { defaultPath: path.join(dossierDepart(), nom), filters: o.filters && o.filters.length ? o.filters : undefined, properties: ['showOverwriteConfirmation', 'createDirectory'] });
    if (r.canceled || !r.filePath) return null;
    const p = path.normalize(r.filePath); autoriser(p); reglages.dernierDossier = path.dirname(p); ecrireReglages();
    return p;
  });
  ipcMain.handle('dialogue:ouvrir', async (_e, o = {}) => {
    const r = await dialog.showOpenDialog(fenetre, { defaultPath: dossierDepart(), filters: o.filters && o.filters.length ? [...o.filters, { name: 'Tous les fichiers', extensions: ['*'] }] : undefined, properties: ['openFile', ...(o.multiple ? ['multiSelections'] : [])] });
    if (r.canceled || !r.filePaths.length) return [];
    const l = r.filePaths.map(p => path.normalize(p)); l.forEach(autoriser); reglages.dernierDossier = path.dirname(l[0]); ecrireReglages();
    return l;
  });
  ipcMain.handle('appli:infos', () => ({ version: app.getVersion(), systeme: process.platform, maj: majPrete }));
  ipcMain.handle('appli:par-defaut', () => parDefaut());
  ipcMain.handle('appli:retablir', () => retablirAvant());
  ipcMain.handle('maj:installer', () => installerMaj());
}

/* ---------- démarrage */
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv, dossier) => {
    // Un fichier ouvert alors que Grimoire tourne déjà : il arrive dans la fenêtre existante
    envoyerFichiers(fichiersDe(argv, dossier));
    if (fenetre) { if (fenetre.isMinimized()) fenetre.restore(); fenetre.focus(); }
  });
  app.whenReady().then(() => {
    lireReglages();
    (reglages.autorises || []).forEach(p => autorises.add(p));
    enAttente = fichiersDe(process.argv);
    if (process.platform === 'win32') app.setAppUserModelId('io.github.csaintecluque-ui.grimoire');
    Menu.setApplicationMenu(null);
    protocol.handle('app', async req => {
      const u = new URL(req.url);
      let p = decodeURIComponent(u.pathname); if (p === '/' || !p) p = '/index.html';
      const f = path.normalize(path.join(WEB, p));
      if (u.host !== 'grimoire' || !f.startsWith(WEB + path.sep)) return new Response('', { status: 403 });
      try { return new Response(await fsp.readFile(f), { headers: { 'Content-Type': TYPES[path.extname(f).toLowerCase()] || 'application/octet-stream' } }); }
      catch { return new Response('', { status: 404 }); }
    });
    ipc();
    creerFenetre();
    setTimeout(verifierMaj, 8000);
    setInterval(verifierMaj, 6 * 3600 * 1000);
  });
  app.on('window-all-closed', () => app.quit());
}
