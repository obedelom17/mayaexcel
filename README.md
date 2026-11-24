# MayaExcel

Fusionnez plusieurs fichiers Excel (.xlsx, .xls, .csv) en **un seul tableau** propre, trié et prêt à télécharger.

- Repère automatiquement la ligne d'en-têtes (même sous un titre ou des lignes vides)
- Regroupe les colonnes qui veulent dire la même chose : `Tel`, `Téléphone`, `Phone` → **Téléphone**
- Range les lignes (tri automatique par Nom/Prénom, ou par la colonne de votre choix), retire les doublons
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
3. Laissez les réglages par défaut : *Framework Preset* = **Other**, pas de *Build Command*, pas d'*Output Directory*. Cliquez sur **Deploy**.

C'est un site statique : aucune installation (`npm`) n'est nécessaire.

## Installer sur le téléphone

- **Android (Chrome)** : un bouton « Installer » apparaît en haut de la page.
- **iPhone (Safari)** : Partager → *Sur l'écran d'accueil*.

## Aperçu lors d'un partage WhatsApp

Dans `index.html`, remplacez `content="/og-image.png"` par l'adresse complète de votre site, par exemple
`content="https://mayaexcel.vercel.app/og-image.png"`.

## Structure

```
index.html            page unique
css/style.css         styles
js/merge.js           moteur de fusion (sans dépendance)
js/app.js             interface, lecture et export
sw.js                 service worker (mode hors connexion)
manifest.webmanifest  configuration de l'application installable
icons/                icônes de l'application
vercel.json           en-têtes HTTP pour Vercel
```

Les bibliothèques SheetJS (lecture des fichiers) et ExcelJS (création du fichier mis en forme) sont chargées depuis cdnjs / jsDelivr
puis mises en cache par le service worker.
