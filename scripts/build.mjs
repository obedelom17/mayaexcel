// Prépare le site pour Vercel : copie les fichiers dans /dist et écrit l'adresse réelle du site
// dans les balises OpenGraph (les aperçus WhatsApp, Facebook, LinkedIn… exigent une adresse complète).
import { cpSync, mkdirSync, readFileSync, writeFileSync, rmSync, readdirSync } from 'node:fs';

const host =
  process.env.SITE_URL ||
  (process.env.VERCEL_PROJECT_PRODUCTION_URL && 'https://' + process.env.VERCEL_PROJECT_PRODUCTION_URL) ||
  (process.env.VERCEL_URL && 'https://' + process.env.VERCEL_URL) ||
  'https://mayaexcel.vercel.app';
const site = host.replace(/\/+$/, '');

rmSync('dist', { recursive: true, force: true });
mkdirSync('dist', { recursive: true });

const SKIP = new Set(['dist', 'scripts', 'node_modules', '.git', '.vercel', 'README.md', 'package.json', 'vercel.json', '.gitignore']);
for (const name of readdirSync('.')) {
  if (!SKIP.has(name)) cpSync(name, 'dist/' + name, { recursive: true });
}

const html = readFileSync('dist/index.html', 'utf8').replaceAll('__SITE_URL__', site);
writeFileSync('dist/index.html', html);
console.log('MayaExcel construit pour', site);
