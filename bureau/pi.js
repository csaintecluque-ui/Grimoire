// Grimoire — accès au serveur Scriptorium (Raspberry Pi) depuis l'appli de bureau :
// connexion, liste des dossiers, lecture et écriture d'un fichier du coffre.
// Tout passe par ce processus : la page n'échange jamais directement avec le Pi.
'use strict';
const fs = require('fs');
const path = require('path');

class Pi {
  constructor({ dossier, safeStorage }) {
    this.fichier = path.join(dossier, 'pi.json');
    this.safe = safeStorage;
    this.cookie = null;
    this.d = { url: null, user: null, pass: null };
    try { Object.assign(this.d, JSON.parse(fs.readFileSync(this.fichier, 'utf8'))); } catch {}
  }
  noter() { try { fs.writeFileSync(this.fichier, JSON.stringify(this.d), { mode: 0o600 }); } catch {} }
  etat() { return { connecte: !!(this.d.url && this.d.user && this.d.pass), url: this.d.url, user: this.d.user }; }

  // mot de passe chiffré par le système quand un trousseau est disponible
  sceller(t) {
    try { if (this.safe.isEncryptionAvailable()) return { enc: true, v: this.safe.encryptString(t).toString('base64') }; } catch {}
    return { enc: false, v: Buffer.from(t, 'utf8').toString('base64') };
  }
  motDePasse() {
    const b = this.d.pass;
    if (!b) return null;
    try { const buf = Buffer.from(b.v, 'base64'); return b.enc ? this.safe.decryptString(buf) : buf.toString('utf8'); } catch { return null; }
  }

  async login(url, user, password) {
    let res;
    try {
      res = await fetch(url + '/api/login', {
        method: 'POST', headers: { 'content-type': 'application/json', 'x-scriptorium': '1' },
        body: JSON.stringify({ user, password }), signal: AbortSignal.timeout(8000)
      });
    } catch { throw new Error('Le Pi est injoignable à cette adresse (hors de la maison sans VPN ?)'); }
    if (res.status === 401) throw new Error('Identifiant ou mot de passe incorrect');
    if (res.status === 429) throw new Error('Trop d\'essais : patientez un peu avant de réessayer');
    if (!res.ok) throw new Error('Ce serveur ne répond pas comme un Scriptorium (' + res.status + ')');
    const cookie = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map(c => c.split(';')[0]).join('; ');
    if (!cookie) throw new Error('Connexion refusée par le serveur');
    return cookie;
  }
  async connecter({ url, user, password } = {}) {
    url = String(url || '').trim().replace(/\/+$/, '');
    if (!url) throw new Error('Indiquez l\'adresse du Pi');
    if (!/^https?:\/\//i.test(url)) url = 'http://' + url;
    try { new URL(url); } catch { throw new Error('Adresse invalide'); }
    user = String(user || '').trim().toLowerCase();
    if (!user || !password) throw new Error('Indiquez l\'identifiant et le mot de passe');
    this.cookie = await this.login(url, user, password);
    this.d = { url, user, pass: this.sceller(password) };
    this.noter();
    return this.etat();
  }
  deconnecter() { this.cookie = null; this.d = { url: null, user: null, pass: null }; this.noter(); return this.etat(); }

  async session() {
    if (this.cookie) return this.cookie;
    if (!this.etat().connecte) throw new Error('Aucun Scriptorium n\'est connecté');
    const pw = this.motDePasse();
    if (!pw) throw new Error('Mot de passe du Pi illisible : reconnectez-vous');
    this.cookie = await this.login(this.d.url, this.d.user, pw);
    return this.cookie;
  }
  async requete(p, opts = {}, encore = true) {
    const cookie = await this.session();
    let res;
    try { res = await fetch(this.d.url + p, { ...opts, headers: { ...(opts.headers || {}), cookie } }); }
    catch { throw new Error('Le Pi est injoignable (hors de la maison sans VPN ?)'); }
    if (res.status === 401 && encore) { this.cookie = null; return this.requete(p, opts, false); }
    return res;
  }
  static chemin(p) {
    p = path.posix.normalize('/' + String(p || '').replace(/\\/g, '/'));
    if (p.includes('\0')) throw new Error('Chemin invalide');
    return p;
  }

  async lister(dir) {
    const res = await this.requete('/api/list?path=' + encodeURIComponent(Pi.chemin(dir)), { signal: AbortSignal.timeout(20000) });
    if (!res.ok) throw new Error(res.status === 404 ? 'Ce dossier n\'existe plus sur le Pi' : 'Erreur du Pi (' + res.status + ')');
    const d = await res.json();
    return { chemin: d.path, elements: d.items.filter(i => !i.name.startsWith('.')) };
  }
  async infos(p) {
    const res = await this.requete('/api/stat?path=' + encodeURIComponent(p), { signal: AbortSignal.timeout(20000) });
    if (res.status === 404) return null;
    if (!res.ok) throw new Error('Erreur du Pi (' + res.status + ')');
    return res.json();
  }
  async lire(p) {
    p = Pi.chemin(p);
    const st = await this.infos(p);
    if (!st || st.dir) throw new Error('ENOENT: ce fichier n\'existe plus sur le Pi');
    const res = await this.requete('/api/download?path=' + encodeURIComponent(p));
    if (!res.ok) throw new Error('Lecture impossible (' + res.status + ')');
    return { bytes: new Uint8Array(await res.arrayBuffer()), mtime: st.mtime * 1000, version: st.version };
  }
  // Écrit le fichier (création ou remplacement ; le Pi garde l'ancienne version). base : version
  // connue à l'ouverture — si le fichier a changé depuis sur le Pi, on refuse (ECONFLIT).
  async ecrire({ chemin, bytes, base } = {}) {
    const p = Pi.chemin(chemin);
    if (p === '/') throw new Error('Chemin invalide');
    const st = await this.infos(p);
    if (st && st.dir) throw new Error('Un dossier porte déjà ce nom sur le Pi');
    if (st && base && st.version !== base) throw new Error('ECONFLIT: ce fichier a été modifié sur le Pi depuis son ouverture');
    const pw = this.motDePasse();
    if (!pw) throw new Error('Mot de passe du Pi illisible : reconnectez-vous');
    let res;
    try {
      res = await fetch(this.d.url + '/dav' + p.split('/').map(encodeURIComponent).join('/'), {
        method: 'PUT', body: Buffer.from(bytes),
        headers: { authorization: 'Basic ' + Buffer.from(this.d.user + ':' + pw, 'utf8').toString('base64') }
      });
    } catch { throw new Error('Le Pi est injoignable (hors de la maison sans VPN ?)'); }
    if (res.status === 507) throw new Error('Le disque du Pi est plein');
    if (res.status === 409) throw new Error('Le dossier de destination n\'existe plus sur le Pi');
    if (res.status !== 201 && res.status !== 204) throw new Error('Écriture refusée par le Pi (' + res.status + ')');
    const apres = await this.infos(p);
    return apres ? apres.version : null;
  }
}

module.exports = { Pi };
