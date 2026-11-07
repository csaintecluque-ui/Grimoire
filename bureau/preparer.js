// Prépare l'appli de bureau à partir de la version web (racine du dépôt) :
// copie la page, le code, les icônes et les bibliothèques dans bureau/app/,
// ajoute shim.js avant app.js, et reprend la version de app.js (APP_VERSION).
'use strict';
const fs = require('fs');
const path = require('path');

const RACINE = path.join(__dirname, '..');
const DEST = path.join(__dirname, 'app');

fs.rmSync(DEST, { recursive: true, force: true });
fs.mkdirSync(DEST, { recursive: true });
for (const f of ['app.js', 'app.css', 'manifest.webmanifest']) fs.copyFileSync(path.join(RACINE, f), path.join(DEST, f));
for (const d of ['icons', 'lib']) fs.cpSync(path.join(RACINE, d), path.join(DEST, d), { recursive: true });
fs.copyFileSync(path.join(__dirname, 'shim.js'), path.join(DEST, 'shim.js'));

const html = fs.readFileSync(path.join(RACINE, 'index.html'), 'utf8');
const balise = '<script src="app.js"></script>';
if (!html.includes(balise)) throw new Error('index.html : ' + balise + ' introuvable');
fs.writeFileSync(path.join(DEST, 'index.html'), html.replace(balise, '<script src="shim.js"></script>' + balise));

const m = /APP_VERSION = '(\d+\.\d+\.\d+)'/.exec(fs.readFileSync(path.join(RACINE, 'app.js'), 'utf8'));
if (!m) throw new Error('APP_VERSION introuvable dans app.js');
const pkgF = path.join(__dirname, 'package.json');
const pkg = JSON.parse(fs.readFileSync(pkgF, 'utf8'));
if (pkg.version !== m[1]) { pkg.version = m[1]; fs.writeFileSync(pkgF, JSON.stringify(pkg, null, 2) + '\n'); }

fs.mkdirSync(path.join(__dirname, 'build'), { recursive: true });
fs.copyFileSync(path.join(RACINE, 'icons', 'icon-512.png'), path.join(__dirname, 'build', 'icon.png'));
console.log('Grimoire ' + m[1] + ' prêt dans bureau/app/');
