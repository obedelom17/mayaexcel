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
- Interface mobile, installable sur l'écran d'accueil (PWA), utilisable hors connexion après la première visite
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
2. Sur [vercel.com](https://vercel.com) : **Add New → Project**, choisissez le dépôt `mayaexcel`.
3. Laissez les réglages par défaut (*Framework Preset* = **Other**) : `vercel.json` indique déjà la commande de construction (`node scripts/build.mjs`) et le dossier `dist`. Cliquez sur **Deploy**.

Aucune installation (`npm install`) n'est nécessaire. La commande de construction écrit l'adresse réelle de votre site dans les balises OpenGraph.
Si vous utilisez un nom de domaine personnalisé, ajoutez dans Vercel (*Settings → Environment Variables*) la variable `SITE_URL` = `https://votre-domaine.com`.

## Installer sur le téléphone

- **Android (Chrome)** : un bouton « Installer » apparaît en haut de la page.
- **iPhone (Safari)** : Partager → *Sur l'écran d'accueil*.

## Aperçu lors d'un partage (OpenGraph)

La page contient toutes les balises OpenGraph et Twitter (titre, description, image 1200×630). Les adresses complètes sont
renseignées automatiquement au déploiement. Après un changement, WhatsApp/Facebook gardent l'ancien aperçu en cache :
utilisez l'outil « Débogueur de partage » de Facebook pour le rafraîchir.

## Structure

```
index.html            page unique
css/style.css         styles
js/merge.js           moteur de fusion (sans dépendance)
scripts/build.mjs     construction pour Vercel (adresse OpenGraph)
js/app.js             interface, lecture et export
sw.js                 service worker (mode hors connexion)
manifest.webmanifest  configuration de l'application installable
icons/                icônes de l'application
vercel.json           en-têtes HTTP pour Vercel
```

Les bibliothèques SheetJS (lecture des fichiers) et ExcelJS (création du fichier mis en forme) sont chargées depuis cdnjs / jsDelivr
puis mises en cache par le service worker.
