Grammalecte 2.1.2 (Olivier R., https://grammalecte.net, GPL v3), version JavaScript
reprise du greffon ONLYOFFICE « plugin-grammalecte » (vendor/grammalecte-sdk).

Modifications pour Grimoire (fonctionnement dans un Web Worker, voir worker.js) :
- chemins des données (conj, mfsp, phonet, dictionnaire) lus depuis self.GL_BASE ;
- mfsp.js : le module mfsp s'initialisait sur conj (« var mfsp = window.grammalecte["conj"] ») ;
- helpers.js : plus de trace console à chaque fichier chargé.
Fichiers inutiles ici (tests, formateur de texte, autres dictionnaires) non repris.
