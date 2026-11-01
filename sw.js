// Grimoire — service worker : fonctionnement hors connexion.
// À chaque mise en ligne d'une nouvelle version, augmenter VERSION.
const VERSION = 'grimoire-1.17.3';
const SHELL = [
  './', './index.html', './app.css', './app.js', './manifest.webmanifest',
  './icons/favicon.svg', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png', './icons/apple-touch-icon.png',
  './lib/hyperformula.full.min.js', './lib/xlsx-style.bundle.js', './lib/jszip.min.js', './lib/mammoth.browser.min.js',
  './lib/pdfjs/pdf.min.mjs', './lib/pdfjs/pdf.worker.min.mjs', './lib/pdf-lib.min.js', './lib/pdfmake.min.js', './lib/vfs_fonts.js',
  // polices standard des PDF (formulaires en Helvetica, Times… non incorporées) : disponibles hors connexion dès la première ouverture
  './lib/pdfjs/standard_fonts/FoxitDingbats.pfb', './lib/pdfjs/standard_fonts/FoxitFixed.pfb', './lib/pdfjs/standard_fonts/FoxitFixedBold.pfb', './lib/pdfjs/standard_fonts/FoxitFixedBoldItalic.pfb', './lib/pdfjs/standard_fonts/FoxitFixedItalic.pfb', './lib/pdfjs/standard_fonts/FoxitSerif.pfb', './lib/pdfjs/standard_fonts/FoxitSerifBold.pfb', './lib/pdfjs/standard_fonts/FoxitSerifBoldItalic.pfb', './lib/pdfjs/standard_fonts/FoxitSerifItalic.pfb', './lib/pdfjs/standard_fonts/FoxitSymbol.pfb', './lib/pdfjs/standard_fonts/LiberationSans-Bold.ttf', './lib/pdfjs/standard_fonts/LiberationSans-BoldItalic.ttf', './lib/pdfjs/standard_fonts/LiberationSans-Italic.ttf', './lib/pdfjs/standard_fonts/LiberationSans-Regular.ttf'
];
const FONTS = 'grimoire-fonts';
const wait = ms => new Promise(res => setTimeout(() => res(null), ms));
const empty = type => new Response('', { status: 200, headers: { 'Content-Type': type } });

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
      const net = fetch(req).then(r => { if (r.ok || r.type === 'opaque') c.put(req, r.clone()); return r; });
      net.catch(() => {});
      if (hit) return hit;
      // pas encore en cache et réseau qui traîne : on n'attend pas plus de 3 s (les polices de secours prennent le relais)
      return Promise.race([net, wait(3000)]).then(r => r || empty(url.hostname === 'fonts.googleapis.com' ? 'text/css' : 'font/woff2')).catch(() => empty('text/css'));
    }));
    return;
  }
  if (url.origin !== self.location.origin) return;
  // Pages : la copie de l'appli d'abord, tout de suite, même avec une connexion lente ou coupée.
  // Les nouvelles versions arrivent par la mise à jour du service worker (bandeau « Recharger »).
  if (req.mode === 'navigate') {
    e.respondWith(caches.match('./index.html').then(hit => {
      const net = fetch(req).then(r => { if (r.ok) caches.open(VERSION).then(c => c.put('./index.html', r.clone())); return r; });
      if (hit) { net.catch(() => {}); return hit; }
      return Promise.race([net, wait(8000)]).catch(() => null).then(r => r || caches.match('./')).then(r => r || net);
    }));
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
