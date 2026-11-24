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

## Applications Android (APK) et Windows (EXE)

Les deux applications sont construites **automatiquement par GitHub** (gratuit), sans rien installer sur votre ordinateur.
Elles contiennent les bibliothèques Excel : tout fonctionne sans connexion, dès la première ouverture.

1. Envoyez le projet sur GitHub (étape ci-dessus).
2. Sur GitHub : onglet **Actions** → **Construire les applications (APK + EXE)** → **Run workflow**.
3. Après 8 à 15 minutes, ouvrez l'exécution terminée : en bas, section **Artifacts**, téléchargez
   - `MayaExcel-Android-APK` → `MayaExcel.apk`
   - `MayaExcel-Windows-EXE` → l'installateur `…Installateur.exe` et la version `…Portable.exe` (sans installation)

Pour publier des liens de téléchargement permanents, créez une *release* avec un tag commençant par `v` (par exemple `v2.0.0`) :
les fichiers sont alors ajoutés automatiquement à la page **Releases** du dépôt.

- **Android** : copiez `MayaExcel.apk` sur le téléphone, ouvrez-le et autorisez l'installation depuis cette source. Au téléchargement du résultat, la feuille de partage du téléphone s'ouvre pour enregistrer ou ouvrir le fichier Excel.
- **Windows** : Windows SmartScreen peut afficher « application non reconnue » (le programme n'est pas signé) : cliquez sur *Informations complémentaires* puis *Exécuter quand même*.

## Installer sans passer par l'APK

- **Android (Chrome)** : un bouton « Installer » apparaît en haut de la page.
- **iPhone (Safari)** : Partager → *Sur l'écran d'accueil*.
- **PC (Chrome ou Edge)** : icône « Installer » dans la barre d'adresse.

L'application installée de cette façon fonctionne aussi hors connexion.

## Structure

```
index.html            page unique
css/style.css         styles
js/merge.js           moteur de fusion (sans dépendance)
js/app.js             interface, lecture et export
sw.js                 service worker (mode hors connexion)
manifest.webmanifest  configuration de l'application installable
icons/                icônes de l'application
scripts/build.mjs     construction (adresse OpenGraph + bibliothèques embarquées)
vercel.json           configuration Vercel
desktop/              application Windows (Electron)
android-res/          icônes de l'application Android
capacitor.config.json configuration de l'application Android (Capacitor)
.github/workflows/    construction automatique de l'APK et de l'EXE
```
