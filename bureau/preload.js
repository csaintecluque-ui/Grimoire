// Grimoire — pont entre la page et l'ordinateur (appli de bureau uniquement).
// Seules ces fonctions sont accessibles à la page ; shim.js s'en sert pour fournir
// les API de fichiers que Grimoire utilise déjà dans Chrome et Edge.
'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('bureauNatif', {
  lire: p => ipcRenderer.invoke('fichier:lire', p),
  ecrire: (p, bytes) => ipcRenderer.invoke('fichier:ecrire', p, bytes),
  enregistrerSous: o => ipcRenderer.invoke('dialogue:enregistrer', o),
  ouvrir: o => ipcRenderer.invoke('dialogue:ouvrir', o),
  fichiersEnAttente: () => ipcRenderer.invoke('fichiers:attente'),
  surFichiers: cb => { ipcRenderer.on('fichiers:ouvrir', (_e, l) => cb(l)); },
  infos: () => ipcRenderer.invoke('appli:infos'),
  parDefaut: () => ipcRenderer.invoke('appli:par-defaut'),
  retablir: () => ipcRenderer.invoke('appli:retablir'),
  surMaj: cb => { ipcRenderer.on('maj:prete', (_e, i) => cb(i)); },
  installerMaj: () => ipcRenderer.invoke('maj:installer')
});
