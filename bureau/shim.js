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
