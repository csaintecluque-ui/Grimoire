# Grimoire

Suite bureautique personnelle, installable sur téléphone et ordinateur (PWA) :

- **Documents** façon Word : titres, polices, tailles, couleurs, surlignage, liens, listes à niveaux et cases à cocher, tableaux complets (lignes, colonnes, fusion, couleurs, largeurs), images avec habillage du texte (à gauche, à droite), sauts de page, sommaire, rechercher-remplacer, mise en page (orientation, marges, en-tête, pied de page, numéros de page), plan du document, mode concentration.
- **Classeurs** façon Excel : près de 400 fonctions aux noms français (`=SOMME`, `=SI`, `=SOMME.SI.ENS`, `=RECHERCHEX`, `=UNIQUE`, `=TRIER`, `=JOINDRE.TEXTE`…) avec aide à la saisie, onglets et formules entre feuilles, plages nommées, recopie et séries, dates, tri et filtre, couleurs, bordures, polices et tailles, décimales, fusion de cellules, lignes et colonnes masquées, volets figés, mise en forme conditionnelle, listes déroulantes, cases à cocher, notes, rechercher-remplacer, graphiques (colonnes, barres, courbes, secteurs) et tableaux croisés dynamiques.
- **Notes rapides** (accueil) : une ligne à taper, Entrée pour l'ajouter ; rappel en un geste (dans 1 h, ce soir, demain, lundi, ou date au choix) avec notification et pastille sur l'icône ; bouton calendrier qui crée l'alarme dans le calendrier du téléphone ; note faite, transformée en document ou supprimée.
- **Correction automatique** (Réglages) : majuscules, accents et apostrophes oubliés, fautes de frappe, typographie française, remplacements personnels.
- **Sauvegarde en ligne** (Réglages) : copie automatique de tout (notes, documents, classeurs, PDF) dans un dépôt GitHub privé, et synchronisation entre appareils.
- **PDF** : lecture (recherche, texte copiable), remplissage des formulaires, annotations (texte, surlignage, dessin, cadre, bloc blanc, signature, image), pages (pivoter, supprimer, dupliquer, réordonner, extraire, diviser, insérer des PDF, photos, fichiers Word, Excel, CSV ou texte), fusion de fichiers de tous ces formats (convertis en pages PDF), conversion en document, export du PDF modifié.
- **Enregistrer / Enregistrer sous** (Ctrl+S / Ctrl+Maj+S) : écrit le document dans un vrai fichier de l'ordinateur (.docx, .xlsx, .csv, .pdf) choisi avec la fenêtre « Enregistrer sous ». Un fichier ouvert depuis l'ordinateur (Importer ou « Ouvrir avec Grimoire ») reste lié : Enregistrer le met à jour directement. Nécessite un navigateur à moteur Chromium sur PC ; ailleurs (Firefox, téléphone), le fichier part dans les téléchargements.
- **Export** (Ctrl+E) : PDF, Markdown, page web, CSV, ou copie Word (.docx), Excel (.xlsx avec graphiques, mises en forme conditionnelles, listes déroulantes, notes, fusions et plages nommées), CSV, Markdown.
- **Impression / PDF** : bouton « Imprimer » (ou Ctrl+P), avec en-tête, pied de page et numéros de page. Exporter → PDF (.pdf) crée directement un PDF du document ou du classeur (tableaux, mise en forme, graphiques, sommaire, en-tête et pied de page).
- **Import** de fichiers .docx (couleurs, tailles, polices, surlignage, alignements, listes, tableaux, images, en-tête et pied de page conservés), .xlsx, .csv, .txt.
- **Mode PC / Téléphone** : bouton en haut de l'écran. Sur grand écran en mode PC, l'accueil passe en trois colonnes : Créer à gauche, documents au centre, notes rapides à droite.
- **Recherche** (accueil) : titre et contenu des documents, classeurs, PDF et notes rapides enregistrés dans Grimoire (pas les fichiers de l'ordinateur).
- **Hors connexion** : tout est enregistré sur l'appareil.
- **Plume** (facultatif) : assistant d'écriture avec une clé Mistral ou OpenRouter, à saisir dans les Réglages.

## Mise en ligne (GitHub Pages)

1. Déposer tous les fichiers à la racine du dépôt (en gardant les dossiers `icons/` et `lib/`).
2. *Settings → Pages → Build and deployment* : *Deploy from a branch*, branche `main`, dossier `/ (root)`.
3. Après une minute, l'appli est à l'adresse `https://<utilisateur>.github.io/<dépôt>/`.

## Installer l'appli

- **Android (Chrome)** : ouvrir l'adresse, menu ⋮ → *Installer l'application* (ou le bandeau « Installer » dans Grimoire).
- **PC (Chrome, Chromium, Edge ou Brave, y compris sous Linux)** : icône d'installation à droite de la barre d'adresse, ou Réglages → Application → Installer Grimoire. Firefox ne sait pas installer d'applis web sous Linux.
- **Ouvrir les fichiers avec Grimoire** (une fois installé) : clic droit sur un .docx, .xlsx, .csv ou .pdf → *Ouvrir avec* → Grimoire ; *Propriétés → Ouvrir avec → Définir par défaut* pour le double-clic. Le document reste lié au fichier : Enregistrer (Ctrl+S) met à jour l'original.
- **Notion** : dans une page, taper `/embed` et coller l'adresse.

## Mettre à jour

Après toute modification, augmenter `VERSION` dans `sw.js` : les appareils installés proposent alors « Recharger ».

## Données

Les documents sont enregistrés dans le navigateur de chaque appareil (IndexedDB). Pour passer d'un appareil à l'autre : *Réglages → Sauvegarder tout*, puis *Restaurer une sauvegarde* sur l'autre appareil.

## Licences

Bibliothèques incluses dans `lib/` : HyperFormula (GPL v3), SheetJS et xlsx-js-style (Apache 2.0), JSZip (MIT), Mammoth (BSD-2), PDF.js (Apache 2.0, polices standard sous licences Foxit et Liberation jointes dans `lib/pdfjs/standard_fonts/`), pdf-lib (MIT), pdfmake (MIT, police Roboto sous licence Apache 2.0). HyperFormula étant sous GPL v3, ce dépôt public est distribué sous la même licence.
