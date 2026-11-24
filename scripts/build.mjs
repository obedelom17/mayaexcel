// Prépare le site pour Vercel : copie les fichiers dans /dist et écrit l'adresse réelle du site
// dans les balises OpenGraph (les aperçus WhatsApp, Facebook, LinkedIn… exigent une adresse complète).
import { cpSync, mkdirSync, readFileSync, writeFileSync, rmSync, readdirSync, existsSync } from 'node:fs';

const host =
  process.env.SITE_URL ||
  (process.env.VERCEL_PROJECT_PRODUCTION_URL && 'https://' + process.env.VERCEL_PROJECT_PRODUCTION_URL) ||
  (process.env.VERCEL_URL && 'https://' + process.env.VERCEL_URL) ||
  'https://mayaexcel.vercel.app';
const site = host.replace(/\/+$/, '');

rmSync('dist', { recursive: true, force: true });
mkdirSync('dist', { recursive: true });

const SKIP = new Set(['dist', 'scripts', 'node_modules', '.git', '.github', '.vercel', 'README.md', 'package.json', 'package-lock.json', 'vercel.json', '.gitignore', 'desktop', 'android', 'android-res', 'capacitor.config.json']);
for (const name of readdirSync('.')) {
  if (!SKIP.has(name)) cpSync(name, 'dist/' + name, { recursive: true });
}

// Bibliothèques Excel embarquées : l'application fonctionne ainsi sans Internet (même la toute première fois).
mkdirSync('dist/vendor', { recursive: true });
for (const [src, dest] of [
  ['node_modules/xlsx/dist/xlsx.full.min.js', 'dist/vendor/xlsx.full.min.js'],
  ['node_modules/exceljs/dist/exceljs.min.js', 'dist/vendor/exceljs.min.js'],
]) {
  if (existsSync(src)) cpSync(src, dest);
  else console.warn('Absent (npm install non lancé ?) :', src);
}

const html = readFileSync('dist/index.html', 'utf8').replaceAll('__SITE_URL__', site);
writeFileSync('dist/index.html', html);
console.log('MayaExcel construit pour', site);
