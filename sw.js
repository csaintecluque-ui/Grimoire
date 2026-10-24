// Grimoire — service worker : fonctionnement hors connexion.
// À chaque mise en ligne d'une nouvelle version, augmenter VERSION.
const VERSION = 'grimoire-1.17.0';
const SHELL = [
  './', './index.html', './app.css', './app.js', './manifest.webmanifest',
  './icons/favicon.svg', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png', './icons/apple-touch-icon.png',
  './lib/hyperformula.full.min.js', './lib/xlsx-style.bundle.js', './lib/jszip.min.js', './lib/mammoth.browser.min.js',
  './lib/pdfjs/pdf.min.mjs', './lib/pdfjs/pdf.worker.min.mjs', './lib/pdf-lib.min.js', './lib/pdfmake.min.js', './lib/vfs_fonts.js',
  // polices standard des PDF (formulaires en Helvetica, Times… non incorporées) : disponibles hors connexion dès la première ouverture
  './lib/pdfjs/standard_fonts/FoxitDingbats.pfb', './lib/pdfjs/standard_fonts/FoxitFixed.pfb', './lib/pdfjs/standard_fonts/FoxitFixedBold.pfb', './lib/pdfjs/standard_fonts/FoxitFixedBoldItalic.pfb', './lib/pdfjs/standard_fonts/FoxitFixedItalic.pfb', './lib/pdfjs/standard_fonts/FoxitSerif.pfb', './lib/pdfjs/standard_fonts/FoxitSerifBold.pfb', './lib/pdfjs/standard_fonts/FoxitSerifBoldItalic.pfb', './lib/pdfjs/standard_fonts/FoxitSerifItalic.pfb', './lib/pdfjs/standard_fonts/FoxitSymbol.pfb', './lib/pdfjs/standard_fonts/LiberationSans-Bold.ttf', './lib/pdfjs/standard_fonts/LiberationSans-BoldItalic.ttf', './lib/pdfjs/standard_fonts/LiberationSans-Italic.ttf', './lib/pdfjs/standard_fonts/LiberationSans-Regular.ttf'
];
const FONTS = 'grimoire-fonts';

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== VERSION && k !== FONTS).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Polices Google : servies depuis le cache, rafraîchies en arrière-plan.
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    e.respondWith(caches.open(FONTS).then(async c => {
      const hit = await c.match(req);
      const net = fetch(req).then(r => { if (r.ok || r.type === 'opaque') c.put(req, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    }));
    return;
  }
  if (url.origin !== self.location.origin) return;
  // Pages : le réseau d'abord (nouvelle version), le cache si hors connexion.
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).catch(() => caches.match('./index.html')));
    return;
  }
  // Fichiers de l'appli : le cache d'abord.
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
    if (r.ok) caches.open(VERSION).then(c => c.put(req, r.clone()));
    return r;
  })));
});

// Rappels : toucher la notification ouvre Grimoire sur la note
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const id = e.notification.data && e.notification.data.note;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    const c = list[0];
    if (c) { c.postMessage({ note: id }); return c.focus(); }
    return self.clients.openWindow('./' + (id ? '?note=' + encodeURIComponent(id) : ''));
  }));
});
