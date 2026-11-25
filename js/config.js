/* Configuration. "__GITHUB_REPO__" est remplacé automatiquement au déploiement (scripts/build.mjs)
   par le dépôt GitHub du projet. Les liens de téléchargement de l'application pointent vers sa dernière release. */
window.MAYA_CONFIG = {
  repo: '__GITHUB_REPO__',
  version: '3.0.0',
  files: { android: 'MayaExcel.apk', windows: 'MayaExcel-Setup.exe', portable: 'MayaExcel-Portable.exe' },
};
