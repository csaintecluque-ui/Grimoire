# Grimoire

Suite bureautique personnelle, installable sur téléphone et ordinateur (PWA) :

- **Documents** façon Word : titres, polices, tailles, couleurs, surlignage, liens, listes à niveaux et cases à cocher, tableaux complets (lignes, colonnes, fusion, couleurs, largeurs), images avec habillage du texte (à gauche, à droite), sauts de page, sommaire, rechercher-remplacer, mise en page (orientation, marges, en-tête, pied de page, numéros de page), plan du document, mode concentration.
- **Classeurs** façon Excel : près de 400 fonctions aux noms français (`=SOMME`, `=SI`, `=SOMME.SI.ENS`, `=RECHERCHEX`, `=UNIQUE`, `=TRIER`, `=JOINDRE.TEXTE`…) avec aide à la saisie, onglets et formules entre feuilles, plages nommées, recopie et séries, dates, tri et filtre, couleurs, bordures, polices et tailles, décimales, fusion de cellules, lignes et colonnes masquées, volets figés, mise en forme conditionnelle, listes déroulantes, cases à cocher, notes, rechercher-remplacer, graphiques (colonnes, barres, courbes, secteurs) et tableaux croisés dynamiques.
  Clic droit sur une cellule : insérer des lignes au-dessus ou en dessous, des colonnes à gauche ou à droite (autant que la sélection en compte), supprimer, choisir le type de cellule (Standard, Nombre, Texte, Euros, Pourcentage, Date, ou une autre devise : dollar, livre, franc suisse, dollar canadien, yen…). Le type Texte garde les zéros (téléphones, codes). Conservé à l'export et à l'import Excel.
- **Commentaires** (documents) : sélectionner un passage, bouton bulle ou Ctrl+Alt+M ; cartes dans la marge de droite, réponses, commentaire résolu ou rouvert. Conservés à l'import et à l'export Word (fils de discussion et état « résolu » compris).
- **Styles du document** : Paragraphe → « Modifier les styles… » pour régler Titre 1, Titre 2, Titre 3, Paragraphe et Citation (police, taille, couleur, gras, italique, alignement, espacements), ou les reprendre d'après un paragraphe ; possibilité d'en faire les styles des nouveaux documents. Repris à l'export Word et PDF.
- **Colonnes** : un passage en 2 ou 3 colonnes (bouton colonnes), exporté en vraie section continue dans Word et relu à l'import.
- **Suivi des modifications** (Ctrl+Maj+E) : ajouts soulignés en bleu, suppressions barrées en rouge, avec auteur et date ; accepter ou refuser une à une ou en bloc. Conservé à l'import et à l'export Word.
- **Macros Excel** : un classeur .xlsm garde ses macros VBA telles quelles (Grimoire ne les exécute pas) et les réécrit à l'enregistrement en .xlsm.
- **Notes rapides** (accueil) : une ligne à taper, Entrée pour l'ajouter ; rappel en un geste (dans 1 h, ce soir, demain, lundi, ou date au choix) avec notification et pastille sur l'icône ; bouton calendrier qui crée l'alarme dans le calendrier du téléphone ; note faite, transformée en document ou supprimée.
- **Correction automatique** (Réglages) : majuscules, accents et apostrophes oubliés, fautes de frappe, typographie française, remplacements personnels.
- **Sauvegarde en ligne** (Réglages) : copie automatique de tout (notes, documents, classeurs, PDF) dans un dépôt GitHub privé, et synchronisation entre appareils.
- **PDF** : lecture (recherche, texte copiable), remplissage des formulaires, annotations (texte, surlignage, dessin, cadre, bloc blanc, signature, image), pages (pivoter, supprimer, dupliquer, réordonner, extraire, diviser, insérer des PDF, photos, fichiers Word, Excel, CSV ou texte), fusion de fichiers de tous ces formats (convertis en pages PDF), conversion en document, export du PDF modifié.
- **Enregistrer / Enregistrer sous** (Ctrl+S / Ctrl+Maj+S) : écrit le document dans un vrai fichier de l'ordinateur (.docx, .xlsx, .csv, .pdf) choisi avec la fenêtre « Enregistrer sous ». Un fichier ouvert depuis l'ordinateur (Importer ou « Ouvrir avec Grimoire ») reste lié : Enregistrer le met à jour directement. Nécessite un navigateur à moteur Chromium sur PC ; ailleurs (Firefox, téléphone), le fichier part dans les téléchargements.
- **Export** (Ctrl+E) : PDF, Markdown, page web, CSV, ou copie Word (.docx), Excel (.xlsx avec graphiques, mises en forme conditionnelles, listes déroulantes, notes, fusions et plages nommées), CSV, Markdown.
- **Impression / PDF** : bouton « Imprimer » (ou Ctrl+P), avec en-tête, pied de page et numéros de page. Exporter → PDF (.pdf) crée directement un PDF du document ou du classeur (tableaux, mise en forme, graphiques, sommaire, en-tête et pied de page).
- **Import** de fichiers .xlsx et .xlsm avec leurs graphiques, règles de couleur, listes déroulantes et commentaires de cellules (repris en notes) ; .docx (couleurs, tailles, polices, surlignage, alignements, listes, tableaux, images, en-tête et pied de page conservés), .xlsx, .csv, .txt.
- **Barre de menus** (mode PC), comme Word ou Google Docs : Fichier (nouveau, modèles, ouvrir, récents, enregistrer, exporter, imprimer), Édition, Affichage, Insertion, Format, Données (classeurs), Outils, Aide. Elle reprend toutes les commandes de la palette (Ctrl+K) et s'adapte au document ouvert. Au clavier : Alt+F (Fichier), Alt+E, Alt+A, Alt+I, Alt+O (Format), Alt+D, Alt+U, Alt+H, ou F10 ; flèches, Entrée, Échap.
- **Mode PC / Téléphone** : bouton en haut de l'écran. Sur grand écran en mode PC, l'accueil passe en trois colonnes : Créer à gauche, documents au centre, notes rapides à droite.
- **Recherche** (accueil) : titre et contenu des documents, classeurs, PDF et notes rapides enregistrés dans Grimoire (pas les fichiers de l'ordinateur).
- **Hors connexion** : tout est enregistré sur l'appareil.
- **Plume** (facultatif) : assistant d'écriture avec une clé Mistral ou OpenRouter, à saisir dans les Réglages.
  Dans un classeur : Plume construit directement ce qu'on lui décrit (tableaux, titres, formules à la française, listes déroulantes, règles de couleur, volets figés, graphiques), vérifie ses formules et corrige celles en erreur ; Ctrl+Z annule toute la construction.
  Sur un PDF : résumer (avec dates limites et choses à faire), répondre à une question, remplir un formulaire à partir de « Mes informations » (Réglages) et de consignes, relever une facture dans le classeur « Suivi des factures » (date, fournisseur, n°, catégorie, HT, TVA, TTC), ou plusieurs factures d'un coup (palette Ctrl+K). Les PDF scannés sans texte ne sont pas lus.

## Mise en ligne (GitHub Pages)

1. Déposer tous les fichiers à la racine du dépôt (en gardant les dossiers `icons/` et `lib/`).
2. *Settings → Pages → Build and deployment* : *Deploy from a branch*, branche `main`, dossier `/ (root)`.
3. Après une minute, l'appli est à l'adresse `https://<utilisateur>.github.io/<dépôt>/`.

## Appli de bureau (Windows, Linux)

Le plus simple sur ordinateur : télécharger l'installateur sur la page [Releases](https://github.com/csaintecluque-ui/Grimoire/releases/latest), comme pour LibreOffice.

- **Windows** : `Grimoire-Installation-x.y.z.exe`, à ouvrir (installation pour l'utilisateur, sans droit administrateur). Au premier lancement, Windows peut afficher « Windows a protégé votre ordinateur » (installateur non signé) : *Informations complémentaires* → *Exécuter quand même*.
- **Linux** (Mint, Ubuntu, Debian) : `grimoire_x.y.z_amd64.deb`, double-clic → *Installer le paquet*.

Au premier lancement, Grimoire propose de devenir l'appli par défaut des .docx, .xlsx, .csv et .pdf : sous Linux c'est fait aussitôt (icônes DOC, XLS, CSV, PDF comprises, applis d'avant notées pour pouvoir y revenir) ; sous Windows, la page Grimoire des *Paramètres → Applications par défaut* s'ouvre, il reste à cliquer *Définir par défaut* (Windows réserve ce clic à l'utilisateur, LibreOffice compris). Réglable ensuite dans *Réglages → Application*.

C'est la même appli que la version web, servie depuis l'ordinateur : hors connexion, et des données qui ne dépendent plus d'un navigateur (ni nettoyage du navigateur, ni désinstallation de la PWA). Enregistrer (Ctrl+S) met à jour le fichier d'origine sous Windows comme sous Linux ; à la fermeture, Grimoire finit d'enregistrer avant de quitter. Les mises à jour arrivent toutes seules (bandeau « Redémarrer pour mettre à jour » sous Windows, « Installer la mise à jour » sous Linux).

Les documents ne passent pas tout seuls de la version navigateur à l'appli : *Réglages → Sauvegarder tout* dans le navigateur, puis *Réglages → Restaurer une sauvegarde* dans l'appli (ou la sauvegarde en ligne GitHub des deux côtés).

**Pour le développement** (dossier `bureau/`) : `npm install` puis `npm start` lance l'appli ; `npm run linux` construit le .deb. Une nouvelle release est construite et publiée automatiquement par GitHub Actions (`.github/workflows/bureau.yml`) dès que `APP_VERSION` de `app.js` change sur `main` ; sur une pull request, les installateurs sont construits pour vérification (onglet *Actions* → *Artifacts*).

## Installer l'appli web

- **Android (Chrome)** : ouvrir l'adresse, menu ⋮ → *Installer l'application* (ou le bandeau « Installer » dans Grimoire).
- **PC (Chrome, Chromium, Edge ou Brave, y compris sous Linux)** : icône d'installation à droite de la barre d'adresse, ou Réglages → Application → Installer Grimoire. Firefox ne sait pas installer d'applis web sous Linux.
- **Ouvrir les fichiers avec Grimoire** (une fois installé) : lancer le fichier d'installation de son système (voir plus bas), ou clic droit sur un .docx, .xlsx, .csv ou .pdf → *Ouvrir avec* → Grimoire. Le document reste lié au fichier : Enregistrer (Ctrl+S) met à jour l'original. Un fichier ouvert alors que Grimoire est déjà lancé arrive dans la fenêtre existante (pas de deuxième fenêtre).
- **Notion** : dans une page, taper `/embed` et coller l'adresse.

### Fichiers de l'ordinateur : appli par défaut et icônes (Windows et Linux)

Une appli web ne peut rien écrire elle-même dans le système : Chrome et Edge enregistrent Grimoire pour les .docx, .xlsx, .csv et .pdf, mais avec une seule icône, celle de l'appli (le manifest déclare une icône par type, que Chromium n'applique pas encore). Un fichier d'installation fait donc, une fois par ordinateur, ce que fait l'installateur de LibreOffice : icônes au sceau selon le type (DOC, XLS, CSV, PDF, voir `windows/icones-apercu.png`) et Grimoire comme appli par défaut. À télécharger aussi depuis Grimoire : *Réglages → Application*. Sans droit administrateur, réversible, à relancer si l'appli est réinstallée. Grimoire doit d'abord être installé depuis un navigateur à moteur Chromium (Edge, Chrome, Chromium, Brave) : Firefox ne sait pas ouvrir de fichiers avec une appli web.

- **Windows** : double-cliquer `windows/grimoire-installation.cmd` (si Windows affiche « Windows a protégé votre ordinateur » : *Informations complémentaires* → *Exécuter quand même*). La page Grimoire de *Paramètres → Applications par défaut* s'ouvre : cliquer **Définir par défaut**. Windows interdit à tout programme, LibreOffice compris, de se choisir lui-même par défaut : ce clic reste à faire. Windows 10 : *Ouvrir avec* → « Grimoire · Document » / « Grimoire · Classeur » → *Toujours*. Pour retirer : `grimoire-installation.cmd -Retirer` dans une invite de commandes.
- **Linux** (Nemo, Nautilus, Dolphin, Thunar…) : `bash ~/Téléchargements/grimoire-installation.sh`, puis rouvrir le gestionnaire de fichiers (`nemo -q`). Grimoire devient l'appli par défaut des cinq types (les applis d'avant sont notées) et les icônes sont copiées dans `hicolor` et dans le thème actif (Mint-Y, Papirus…). Options : `--sans-defaut` (icônes seulement), `--diagnostic` (ce que voit le bureau pour chaque type), `--retirer` (tout retirer et remettre les applis d'avant).
- **Mac, téléphone, tablette** : le navigateur ne permet pas de changer ces icônes.

## Mettre à jour

Après toute modification, augmenter `VERSION` dans `sw.js` : les appareils installés proposent alors « Recharger ».

## Données

Les documents sont enregistrés dans le navigateur de chaque appareil (IndexedDB). Pour passer d'un appareil à l'autre : *Réglages → Sauvegarder tout*, puis *Restaurer une sauvegarde* sur l'autre appareil.

## Licences

Bibliothèques incluses dans `lib/` : HyperFormula (GPL v3), SheetJS et xlsx-js-style (Apache 2.0), JSZip (MIT), Mammoth (BSD-2), PDF.js (Apache 2.0, polices standard sous licences Foxit et Liberation jointes dans `lib/pdfjs/standard_fonts/`), pdf-lib (MIT), pdfmake (MIT, police Roboto sous licence Apache 2.0). HyperFormula étant sous GPL v3, ce dépôt public est distribué sous la même licence.
