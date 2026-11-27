# MayaExcel

Fusion de plusieurs fichiers Excel (.xlsx, .xls, .csv, .ods) en un seul tableau propre, trié et prêt à exporter.
Application plein écran sur ordinateur, adaptée au téléphone, utilisable hors connexion. Aucun fichier n'est envoyé sur Internet.

## Poste de travail

- Assistant en trois étapes (Sources, Vérification, Export) et navigation latérale
- **Types de fusion** prédéfinis : personnes et effectifs, chiffres et comptabilité, listes et inventaires, empilement simple, contrôle strict
- **Modèles** : enregistrez vos réglages et corrections de colonnes pour les réutiliser sur un prochain lot ; export / import au format `.mayaexcel.json`
- **Historique** des exports (réglages réutilisables, aucun fichier conservé) et **aide** intégrée
- Indicateurs chiffrés, onglets Tableau / Colonnes / Contrôle, barre d'état

## Méthodes de fusion

**Lecture des sources**
- Ligne des titres détectée automatiquement (même sous un titre ou des lignes vides), imposée par numéro, ou absente
- Fichiers sans titres de colonnes : colonnes reconnues d'après leur contenu (e-mails, téléphones, sexe, noms)
- Fiches individuelles « Nom : … / Téléphone : … » : une fiche devient une ligne
- Titres sur deux lignes (Père / Mère…), lignes « Total » et titres répétés ignorés, valeurs « N/A », « néant », « - » ignorées

**Rapprochement des colonnes**
- Synonymes et abréviations français / anglais (Tel., Tph, Dt naissance, DDN, Fct, Mat…), accents et fautes de frappe tolérés
- Reconnaissance par le contenu, par recoupement de valeurs, trois niveaux de sensibilité
- « Nom et Prénoms » séparé en Nom et Prénom, ordre des colonnes cohérent entre les sources
- Règles personnalisées : colonnes à ignorer, équivalences d'intitulés ; corrections manuelles (séparer / rattacher) dans l'onglet Colonnes

**Reconnaissance des personnes**
- Matricule, e-mail, téléphone, nom (ordre inversé, accents, fautes) ; la date de naissance départage les homonymes
- Regroupement en chaîne (une ligne peut relier deux personnes déjà connues), trois niveaux de rigueur
- Valeurs différentes : garder la première, la dernière, la plus complète ou la plus fréquente ; cellule signalée et reportée dans la feuille « Rapport »
- Chaque regroupement est listé dans l'onglet Contrôle et peut être dissocié

**Harmonisation**
- Dates (formats texte, numéros de série Excel), nombres, téléphones (indicatif), majuscules, sexe, situation matrimoniale, Oui / Non, libellés proches unifiés (poste, service, ville…)

## Mettre en ligne (GitHub + Vercel)

1. Décompressez le zip, puis dans le dossier :
   ```bash
   git init
   git add .
   git commit -m "MayaExcel"
   git branch -M main
   git remote add origin https://github.com/VOTRE-COMPTE/mayaexcel.git
   git push -u origin main
   ```
2. Sur [vercel.com](https://vercel.com) : **Add New → Project**, choisissez le dépôt. Réglages par défaut (*Framework Preset* = **Other**) : `vercel.json` indique déjà `npm install`, la construction (`node scripts/build.mjs`) et le dossier `dist`.

La construction embarque les bibliothèques Excel et les polices, écrit l'adresse du site dans les balises OpenGraph et le nom du dépôt GitHub dans les liens de téléchargement.
Domaine personnalisé : variable `SITE_URL` (Vercel → Settings → Environment Variables). Dépôt GitHub différent : variable `GITHUB_REPO` = `compte/depot`.

## Applications Android (APK) et Windows (EXE)

Construites automatiquement par GitHub (gratuit) à chaque envoi sur `main` et publiées dans la release « latest » :
`MayaExcel.apk`, `MayaExcel-Setup.exe` (installateur), `MayaExcel-Portable.exe`.

Le site propose le téléchargement adapté à l'appareil (bouton « Application »). Sur iPhone, il explique l'ajout à l'écran d'accueil depuis Safari.

Premier lancement : onglet **Actions** de GitHub → **Construire les applications (APK + EXE)** → **Run workflow** (ou un simple `git push`). Comptez 8 à 15 minutes.

- **Android** : APK de test (clé de développement) : installation directe après autorisation, pas de publication sur le Play Store. À l'export, la feuille de partage du téléphone s'ouvre.
- **Windows** : SmartScreen peut signaler un programme non signé : *Informations complémentaires* puis *Exécuter quand même*. La fenêtre s'ouvre en plein écran.

## Réglages et raccourcis

Panneau **Réglages** (icône en haut à droite) : fusion, données, tableau, export (xlsx, xls, ods, csv ; et ,, tsv, json), affichage (densité, taille du texte, animations). Les réglages sont mémorisés sur l'appareil.

`Ctrl` + `O` ajouter des fichiers · `Ctrl` + `S` exporter · `Ctrl` + `,` réglages · `Ctrl` + `V` coller un tableau · flèches et `Ctrl` + `C` dans le tableau.

## Structure

```
index.html            page unique
css/style.css         styles
js/config.js          dépôt GitHub et noms des fichiers de l'application
js/merge.js           moteur de fusion (sans dépendance, testable sous Node)
js/app.js             interface, réglages, lecture et export
sw.js                 service worker (mode hors connexion du site)
scripts/build.mjs     construction (OpenGraph, bibliothèques et polices embarquées)
desktop/              application Windows (Electron)
android-res/          icônes Android · capacitor.config.json : application Android (Capacitor)
.github/workflows/    construction automatique de l'APK et des EXE
```
