# MayaExcel

Fusionnez plusieurs fichiers Excel (.xlsx, .xls, .csv) en **un seul tableau** propre, trié et prêt à télécharger.

- Repère automatiquement la ligne d'en-têtes (même sous un titre ou des lignes vides)
- Regroupe les colonnes qui veulent dire la même chose : `Tel`, `Téléphone`, `Phone` → **Téléphone**
- Reconnaît aussi une colonne **d'après son contenu** (e-mails, numéros, M/F…), même sans titre ou avec un titre inconnu, et les fichiers **sans ligne d'en-têtes**
- Sépare « Nom et Prénoms » quand d'autres fichiers ont Nom et Prénom à part
- Regroupe les lignes d'une **même personne** présente dans plusieurs fichiers (matricule, e-mail, téléphone ou nom) ; en cas de valeurs différentes, la première est gardée et la case est surlignée
- Uniformise les dates (« 12 mars 2019 », « 03/02/2021 »…), les nombres, les majuscules, M/F
- Range les lignes (tri automatique par Nom/Prénom, ou par la colonne de votre choix)
- Chaque regroupement automatique peut être corrigé (✂ pour séparer, « Rattacher à » pour rassembler)
- Thème blanc ou sombre, application Android (APK) et Windows (EXE) téléchargeables depuis le site, utilisable hors connexion
- Tout est traité dans le navigateur : aucun fichier n'est envoyé sur un serveur

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
2. Sur [vercel.com](https://vercel.com) : **Add New → Project**, choisissez le dépôt. Laissez les réglages par défaut (*Framework Preset* = **Other**) :
   `vercel.json` indique déjà `npm install`, la commande de construction (`node scripts/build.mjs`) et le dossier `dist`.

La construction embarque les bibliothèques Excel et les polices, écrit l'adresse réelle du site dans les balises OpenGraph et le nom du dépôt GitHub dans les liens de téléchargement de l'application.
Nom de domaine personnalisé : ajoutez dans Vercel (*Settings → Environment Variables*) `SITE_URL` = `https://votre-domaine.com`.
Dépôt GitHub différent du dépôt Vercel : variable `GITHUB_REPO` = `compte/depot`.

## Applications Android (APK) et Windows (EXE)

Elles sont construites **automatiquement par GitHub** (gratuit) à chaque envoi sur la branche `main`, et publiées dans la release « latest » du dépôt :

- `MayaExcel.apk` (Android)
- `MayaExcel-Setup.exe` (Windows, installateur) et `MayaExcel-Portable.exe` (sans installation)

**Le site télécharge directement ces fichiers** : sur Android, le bouton « App Android » lance le téléchargement de l'APK ; sur Windows, celui de l'EXE ; sur iPhone, il explique comment ajouter le site à l'écran d'accueil (Apple n'autorise pas les APK).

Premier lancement : onglet **Actions** de GitHub → **Construire les applications (APK + EXE)** → **Run workflow** (ou simplement un `git push`). Comptez 8 à 15 minutes ; les liens du site fonctionnent dès que la release est publiée.

- **Android** : l'APK est de type « test » (clé de développement) : il s'installe directement après autorisation de l'installation depuis le navigateur, mais n'est pas publiable sur le Play Store. Au téléchargement du résultat, la feuille de partage du téléphone s'ouvre pour enregistrer ou ouvrir le fichier Excel.
- **Windows** : SmartScreen peut afficher « application non reconnue » (programme non signé) : *Informations complémentaires* puis *Exécuter quand même*.

## Réglages

Le panneau **Réglages** (icône en haut à droite, ou `Ctrl` + `,`) contient : fusion (feuilles, ligne des titres, sensibilité du rapprochement, lecture par contenu, clé personnalisée, valeurs en conflit), données (majuscules, téléphones, dates, cases vides), tableau (tri, numérotation, aperçu), export (7 formats, nom du fichier, mise en forme Excel) et affichage (thème clair/sombre/système, couleur d'accentuation, densité, taille du texte, animations).
Les réglages sont mémorisés sur l'appareil.

## Raccourcis

`Ctrl` + `O` choisir des fichiers · `Ctrl` + `S` télécharger · `Ctrl` + `,` réglages · `Ctrl` + `V` coller un tableau copié depuis Excel.

## Aperçu lors d'un partage (OpenGraph)

La page contient toutes les balises OpenGraph et Twitter (titre, description, image 1200×630) avec les adresses complètes, renseignées automatiquement au déploiement.
Après un changement, WhatsApp/Facebook gardent l'ancien aperçu en cache : utilisez le « Débogueur de partage » de Facebook pour le rafraîchir.

## Structure

```
index.html            page unique
css/style.css         styles (thème blanc, mode sombre, accents)
js/config.js          dépôt GitHub et noms des fichiers de l'application
js/merge.js           moteur de fusion (sans dépendance)
js/app.js             interface, réglages, lecture et export
sw.js                 service worker (mode hors connexion du site)
icons/                icônes
scripts/build.mjs     construction (OpenGraph, bibliothèques et polices embarquées)
vercel.json           configuration Vercel
desktop/              application Windows (Electron)
android-res/          icônes de l'application Android
capacitor.config.json configuration de l'application Android (Capacitor)
.github/workflows/    construction automatique de l'APK et des EXE
```
