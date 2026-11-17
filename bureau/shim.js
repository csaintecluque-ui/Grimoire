// Grimoire — appli de bureau : fournit à la page les API de fichiers de Chrome et Edge
// (launchQueue, showSaveFilePicker, showOpenFilePicker, poignées de fichiers) en
// s'appuyant sur l'ordinateur. app.js n'a presque rien à savoir de l'appli de bureau.
// Chargé juste avant app.js, uniquement dans l'appli de bureau (voir preparer.js).
(() => {
  'use strict';
  const N = window.bureauNatif;
  if (!N) return;
  document.documentElement.classList.add('bureau');

  const MIMES = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', xls: 'application/vnd.ms-excel', ods: 'application/vnd.oasis.opendocument.spreadsheet', csv: 'text/csv', txt: 'text/plain', md: 'text/markdown', pdf: 'application/pdf', html: 'text/html', json: 'application/json', ics: 'text/calendar' };
  const nomDe = p => String(p).split(/[\\/]/).pop();
  const typeDe = n => MIMES[(String(n).split('.').pop() || '').toLowerCase()] || '';
  const annule = () => new DOMException('Annulé', 'AbortError');
  const erreur = e => {
    const m = String(e && e.message || e);
    return new DOMException(m, /EBUSY|EPERM|EACCES|ETXTBSY|EROFS/.test(m) ? 'NoModificationAllowedError' : /ENOENT/.test(m) ? 'NotFoundError' : 'NotReadableError');
  };

  // Poignée de fichier, comme FileSystemFileHandle. Ses champs (kind, name, __path) se rangent
  // tels quels dans IndexedDB ; grimoireBureau.poignee() lui rend ses méthodes au retour.
  class Poignee {
    constructor(p) { this.kind = 'file'; this.name = nomDe(p); this.__path = p; }
    async getFile() {
      try { const r = await N.lire(this.__path); return new File([r.bytes], this.name, { type: typeDe(this.name), lastModified: r.mtime }); }
      catch (e) { throw erreur(e); }
    }
    async createWritable() {
      const parts = [], p = this.__path;
      return {
        async write(d) {
          if (d && typeof d === 'object' && !(d instanceof Blob) && !ArrayBuffer.isView(d) && !(d instanceof ArrayBuffer) && 'type' in d) { if (d.type !== 'write') return; d = d.data; }
          parts.push(d);
        },
        async close() {
          const bytes = new Uint8Array(await new Blob(parts).arrayBuffer());
          try { await N.ecrire(p, bytes); } catch (e) { throw erreur(e); }
        },
        async abort() { parts.length = 0; }
      };
    }
    async queryPermission() { return 'granted'; }
    async requestPermission() { return 'granted'; }
    async isSameEntry(o) { return !!o && o.__path === this.__path; }
  }
  const filtres = types => (types || []).map(t => ({
    name: t.description || 'Fichiers',
    extensions: Object.values(t.accept || {}).flat().map(e => String(e).replace(/^\./, '')).filter(Boolean)
  })).filter(f => f.extensions.length);

  const definir = (nom, valeur) => Object.defineProperty(window, nom, { value: valeur, configurable: true, writable: true });
  definir('showSaveFilePicker', async (o = {}) => {
    const p = await N.enregistrerSous({ suggestedName: o.suggestedName || 'Sans titre', filters: filtres(o.types) });
    if (!p) throw annule();
    return new Poignee(p);
  });
  definir('showOpenFilePicker', async (o = {}) => {
    const l = await N.ouvrir({ multiple: !!o.multiple, filters: filtres(o.types) });
    if (!l || !l.length) throw annule();
    return l.map(p => new Poignee(p));
  });

  // Fichiers ouverts depuis l'ordinateur (double-clic, « Ouvrir avec »), comme la File Handling API
  let consommateur = null;
  const file = [];
  const livrer = l => {
    const files = (l || []).map(p => new Poignee(p));
    if (!files.length) return;
    if (consommateur) consommateur({ files }); else file.push(...files);
  };
  definir('launchQueue', {
    setConsumer(cb) {
      consommateur = cb;
      if (file.length) cb({ files: file.splice(0) });
      N.fichiersEnAttente().then(livrer).catch(() => {});
    }
  });
  N.surFichiers(livrer);


  // ---------- Fichiers du Pi (serveur Scriptorium) : mêmes poignées que pour l'ordinateur
  const pi = (nom, arg) => N.pi(nom, arg);
  const clair = e => String(e && e.message || e).replace(/^Error invoking remote method '[^']+': (Error: )?/, '');
  const erreurPi = e => { const m = clair(e); return new DOMException(m.replace(/^E[A-Z]+: /, ''), /ENOENT/.test(m) ? 'NotFoundError' : 'NotReadableError'); };
  class PoigneePi {
    constructor(o) { this.kind = 'file'; this.name = nomDe(o.__pi); this.__pi = o.__pi; this.__v = o.__v || null; }
    async getFile() {
      try { const r = await pi('lire', this.__pi); this.__v = r.version; return new File([r.bytes], this.name, { type: typeDe(this.name), lastModified: r.mtime }); }
      catch (e) { throw erreurPi(e); }
    }
    async createWritable() {
      const parts = [], moi = this;
      return {
        async write(d) {
          if (d && typeof d === 'object' && !(d instanceof Blob) && !ArrayBuffer.isView(d) && !(d instanceof ArrayBuffer) && 'type' in d) { if (d.type !== 'write') return; d = d.data; }
          parts.push(d);
        },
        async close() {
          const bytes = new Uint8Array(await new Blob(parts).arrayBuffer());
          try { moi.__v = await pi('ecrire', { chemin: moi.__pi, bytes, base: moi.__v }); }
          catch (e) {
            if (!/ECONFLIT/.test(clair(e))) { const x = erreurPi(e); x.userMsg = 'Enregistrement sur le Pi impossible : ' + x.message; throw x; }
            if (!confirm(`« ${moi.name} » a été modifié sur le Pi depuis que vous l'avez ouvert.\n\nOK : le remplacer par votre version (l'autre reste dans les versions précédentes du Pi).\nAnnuler : ne rien écrire.`)) { const x = new DOMException('Enregistrement annulé', 'AbortError'); x.userMsg = 'Rien n\'a été écrit sur le Pi.'; throw x; }
            try { moi.__v = await pi('ecrire', { chemin: moi.__pi, bytes, base: null }); } catch (e2) { const x = erreurPi(e2); x.userMsg = 'Enregistrement sur le Pi impossible : ' + x.message; throw x; }
          }
        },
        async abort() { parts.length = 0; }
      };
    }
    async queryPermission() { return 'granted'; }
    async requestPermission() { return 'granted'; }
    async isSameEntry(o) { return !!o && o.__pi === this.__pi; }
  }

  // Fenêtre de choix dans le coffre du Pi (connexion à la première utilisation).
  const EXT_OK = ['docx', 'xlsx', 'xlsm', 'xls', 'ods', 'csv', 'txt', 'md', 'pdf'];
  const el = (tag, attrs, ...kids) => {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) { if (v == null || v === false) continue; if (k === 'text') e.textContent = v; else if (k.startsWith('on')) e.addEventListener(k.slice(2), v); else e.setAttribute(k, v === true ? '' : v); }
    for (const k of kids.flat()) if (k != null && k !== false) e.append(k);
    return e;
  };
  const style = document.createElement('style');
  style.textContent = `
.pi-fond{position:fixed;inset:0;z-index:5000;display:grid;place-items:center;padding:16px;background:rgba(15,10,7,.55)}
.pi-boite{width:min(560px,100%);max-height:calc(100vh - 32px);display:flex;flex-direction:column;gap:12px;padding:20px;border-radius:14px;background:var(--surface,#F6F2EC);color:var(--fg,#1B1512);border:1px solid var(--line,#DCD3C8);box-shadow:0 24px 70px rgba(0,0,0,.4);font-family:var(--f-ui,system-ui);font-size:15px}
.pi-boite h2{margin:0;font:600 26px/1.15 var(--f-display,Georgia)}
.pi-boite label{display:grid;gap:5px;font-size:13.5px}
.pi-boite input{font:inherit;color:inherit;background:var(--paper,#fff);border:1px solid var(--line-strong,#C5B8A8);border-radius:8px;padding:9px 11px;min-height:40px}
.pi-boite input:focus-visible,.pi-boite button:focus-visible{outline:2px solid var(--copper,#C9804F);outline-offset:1px}
.pi-chemin{margin:0;padding:8px 11px;border-radius:8px;background:var(--hover,#E5DED4);font-size:13.5px;overflow-wrap:anywhere}
.pi-liste{display:flex;flex-direction:column;gap:1px;min-height:120px;max-height:300px;overflow-y:auto;border:1px solid var(--line,#DCD3C8);border-radius:10px;padding:5px;background:var(--paper,#fff)}
.pi-liste button{display:flex;align-items:center;gap:10px;min-height:40px;padding:0 10px;border:0;border-radius:7px;background:none;color:inherit;font:inherit;text-align:left;cursor:pointer}
.pi-liste button:hover{background:var(--hover,#E5DED4)}
.pi-liste button[aria-pressed=true]{background:var(--sel,rgba(201,128,79,.13));box-shadow:inset 0 0 0 1px var(--selline,#C9804F)}
.pi-liste .pi-t{flex:none;min-width:44px;font-size:11px;font-weight:700;letter-spacing:.04em;color:var(--copper,#C9804F)}
.pi-liste .pi-n{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pi-liste p,.pi-note{margin:8px 10px;font-size:13.5px;opacity:.75}
.pi-note{margin:0}
.pi-err{margin:0;min-height:1.2em;font-size:13.5px;color:var(--err,#B3261E)}
.pi-actions{display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap}
.pi-actions button{min-height:40px;padding:0 16px;border-radius:9px;border:1px solid var(--line-strong,#C5B8A8);background:none;color:inherit;font:inherit;cursor:pointer}
.pi-actions .pi-ok{background:var(--copper,#C9804F);border-color:var(--copper,#C9804F);color:#1B1512;font-weight:600}
.pi-actions .pi-gauche{margin-right:auto;border-color:transparent;opacity:.8}
.pi-actions button[disabled]{opacity:.45;pointer-events:none}`;
  document.head.append(style);

  function dialoguePi({ mode, nom, extensions }) {
    return new Promise(resolve => {
      const avant = document.activeElement;
      const boite = el('div', { class: 'pi-boite', role: 'dialog', 'aria-modal': 'true', 'aria-label': mode === 'ouvrir' ? 'Ouvrir depuis le Pi' : 'Enregistrer sur le Pi' });
      const fond = el('div', { class: 'pi-fond' }, boite);
      const fermer = v => { fond.remove(); document.removeEventListener('keydown', touche, true); if (avant && avant.focus) avant.focus(); resolve(v); };
      const touche = e => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); fermer(null); } };
      document.addEventListener('keydown', touche, true);
      fond.addEventListener('pointerdown', e => { if (e.target === fond) fermer(null); });
      document.body.append(fond);

      const connexion = etat => {
        const url = el('input', { id: 'piUrl', value: etat.url || 'http://192.168.1.80:8080', spellcheck: 'false' });
        const user = el('input', { id: 'piUser', value: etat.user || '', autocapitalize: 'none', spellcheck: 'false' });
        const pass = el('input', { id: 'piPass', type: 'password' });
        const err = el('p', { class: 'pi-err', role: 'alert' });
        const ok = el('button', { class: 'pi-ok', type: 'submit', text: 'Se connecter' });
        const form = el('form', { class: 'pi-boite', style: 'padding:0;border:0;box-shadow:none;background:none' },
          el('h2', { text: 'Se connecter au Pi' }),
          el('label', {}, 'Adresse de Scriptorium', url), el('label', {}, 'Identifiant', user), el('label', {}, 'Mot de passe', pass), err,
          el('p', { class: 'pi-note', text: 'Le mot de passe est gardé sur cet ordinateur, chiffré par le système quand c\'est possible. Hors de la maison, activez d\'abord votre VPN.' }),
          el('div', { class: 'pi-actions' }, el('button', { type: 'button', text: 'Annuler', onclick: () => fermer(null) }), ok));
        form.addEventListener('submit', async e => {
          e.preventDefault(); err.textContent = ''; ok.disabled = true;
          try { await pi('connecter', { url: url.value, user: user.value, password: pass.value }); parcourir(); }
          catch (x) { err.textContent = clair(x); ok.disabled = false; }
        });
        boite.replaceChildren(form); (etat.user ? pass : url).focus();
      };

      const parcourir = () => {
        const exts = mode === 'ouvrir' ? EXT_OK : (extensions && extensions.length ? extensions : EXT_OK);
        let ici = '/', choisi = null, noms = new Map();
        const chemin = el('p', { class: 'pi-chemin' }), liste = el('div', { class: 'pi-liste' }), err = el('p', { class: 'pi-err', role: 'alert' });
        const champ = mode === 'enregistrer' ? el('input', { id: 'piNom', value: nom || 'Sans titre', spellcheck: 'false' }) : null;
        const ok = el('button', { class: 'pi-ok', type: 'button', id: 'piOk', text: mode === 'ouvrir' ? 'Ouvrir' : 'Enregistrer ici', disabled: mode === 'ouvrir' });
        const valider = () => {
          if (mode === 'ouvrir') { if (choisi) { try { localStorage.setItem('grim.pi.dir', ici); } catch {} fermer([(ici === '/' ? '' : ici) + '/' + choisi]); } return; }
          let n = champ.value.trim().replace(/[\\/]+/g, ' ');
          if (!n) { err.textContent = 'Indiquez un nom de fichier.'; return; }
          if (!exts.includes((n.split('.').pop() || '').toLowerCase()) || !n.includes('.')) n += '.' + exts[0];
          if (noms.get(n) === true) { err.textContent = 'Un dossier porte déjà ce nom.'; return; }
          if (noms.has(n) && !confirm(`« ${n} » existe déjà dans ce dossier du Pi.\n\nOK : le remplacer (l'ancien reste dans les versions précédentes du Pi).`)) return;
          try { localStorage.setItem('grim.pi.dir', ici); } catch {}
          fermer((ici === '/' ? '' : ici) + '/' + n);
        };
        ok.addEventListener('click', valider);
        if (champ) champ.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); valider(); } });
        const montrer = async dir => {
          liste.replaceChildren(el('p', { text: 'Chargement…' })); err.textContent = '';
          let r;
          try { r = await pi('lister', dir); }
          catch (x) { if (dir !== '/') return montrer('/'); liste.replaceChildren(el('p', { text: clair(x) })); return; }
          ici = r.chemin; choisi = null; if (mode === 'ouvrir') ok.disabled = true;
          noms = new Map(r.elements.map(i => [i.name, !!i.dir]));
          chemin.textContent = 'Pi' + (ici === '/' ? '' : ici.replace(/\//g, ' / '));
          const fichiers = r.elements.filter(i => !i.dir && exts.includes((i.name.split('.').pop() || '').toLowerCase()));
          liste.replaceChildren(
            ici !== '/' ? el('button', { type: 'button', onclick: () => montrer(ici.slice(0, ici.lastIndexOf('/')) || '/') }, el('span', { class: 'pi-t', text: '↑' }), el('span', { class: 'pi-n', text: 'Dossier parent' })) : null,
            ...r.elements.filter(i => i.dir).map(i => el('button', { type: 'button', 'data-dossier': i.name, onclick: () => montrer((ici === '/' ? '' : ici) + '/' + i.name) }, el('span', { class: 'pi-t', text: 'DOSSIER' }), el('span', { class: 'pi-n', text: i.name }))),
            ...fichiers.map(i => {
              const b = el('button', { type: 'button', 'data-fichier': i.name, 'aria-pressed': 'false' }, el('span', { class: 'pi-t', text: (i.name.split('.').pop() || '').toUpperCase() }), el('span', { class: 'pi-n', text: i.name }));
              b.addEventListener('click', () => {
                if (mode === 'enregistrer') { champ.value = i.name; return; }
                choisi = i.name; ok.disabled = false; liste.querySelectorAll('[data-fichier]').forEach(x => x.setAttribute('aria-pressed', String(x === b)));
              });
              if (mode === 'ouvrir') b.addEventListener('dblclick', () => { choisi = i.name; valider(); });
              return b;
            }),
            !r.elements.filter(i => i.dir).length && !fichiers.length ? el('p', { text: mode === 'ouvrir' ? 'Aucun document ici.' : 'Dossier vide.' }) : null);
        };
        boite.replaceChildren(
          el('h2', { text: mode === 'ouvrir' ? 'Ouvrir depuis le Pi' : 'Enregistrer sur le Pi' }), chemin, liste,
          champ ? el('label', {}, 'Nom du fichier', champ) : null, err,
          el('div', { class: 'pi-actions' },
            el('button', { class: 'pi-gauche', type: 'button', text: 'Changer de compte', onclick: async () => { connexion(await pi('deconnecter')); } }),
            el('button', { type: 'button', text: 'Annuler', onclick: () => fermer(null) }), ok));
        let depart = '/'; try { depart = localStorage.getItem('grim.pi.dir') || '/'; } catch {}
        montrer(depart);
        (champ || ok).focus();
      };

      pi('etat').then(e => (e.connecte ? parcourir() : connexion(e)), x => { boite.replaceChildren(el('p', { class: 'pi-err', text: clair(x) })); });
    });
  }

  window.grimoirePi = {
    poignee: o => new PoigneePi(o),
    async ouvrir() { const l = await dialoguePi({ mode: 'ouvrir' }); if (!l) throw annule(); return l.map(p => new PoigneePi({ __pi: p })); },
    async enregistrerSous(o = {}) {
      const p = await dialoguePi({ mode: 'enregistrer', nom: o.suggestedName, extensions: filtres(o.types).flatMap(f => f.extensions) });
      if (!p) throw annule();
      return new PoigneePi({ __pi: p });
    }
  };

  window.grimoireBureau = { poignee: p => new Poignee(p), infos: () => N.infos() };

  // Mises à jour : bandeau « Grimoire x.y.z est prêt »
  let maj = null;
  const montrerMaj = i => {
    maj = i;
    const b = document.getElementById('updBanner'); if (!b) return;
    b.querySelector('span').textContent = `Grimoire ${i.version} est prêt.`;
    document.getElementById('updBtn').textContent = i.systeme === 'win' ? 'Redémarrer pour mettre à jour' : 'Installer la mise à jour';
    b.hidden = false;
  };
  N.surMaj(montrerMaj);
  N.infos().then(i => { if (i && i.maj) montrerMaj(i.maj); }).catch(() => {});
  const btn = document.getElementById('updBtn');
  if (btn) btn.addEventListener('click', e => { if (!maj) return; e.stopImmediatePropagation(); N.installerMaj(); }, true);

  // Réglages → Application : appli par défaut des fichiers
  const resultat = (r, el) => { el.textContent = r && r.message || ''; el.className = 'set-res' + (r && r.ok ? ' ok' : ' err'); };
  const parDef = document.getElementById('bureauDefaut'), ret = document.getElementById('bureauRetablir'), res = document.getElementById('bureauRes');
  if (parDef) parDef.addEventListener('click', async () => { res.textContent = '…'; resultat(await N.parDefaut(), res); });
  if (ret) ret.addEventListener('click', async () => { res.textContent = '…'; resultat(await N.retablir(), res); });
  N.infos().then(i => { if (ret) ret.hidden = !i || i.systeme !== 'linux'; }).catch(() => {});
})();
