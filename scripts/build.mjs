// Prépare le site : copie les fichiers dans /dist, embarque les bibliothèques et polices (mode hors connexion),
// et écrit l'adresse réelle du site (OpenGraph) et le dépôt GitHub (liens de téléchargement de l'application).
import { cpSync, mkdirSync, readFileSync, writeFileSync, rmSync, readdirSync, existsSync } from 'node:fs';

const env = process.env;
const host =
  env.SITE_URL ||
  (env.VERCEL_PROJECT_PRODUCTION_URL && 'https://' + env.VERCEL_PROJECT_PRODUCTION_URL) ||
  (env.VERCEL_URL && 'https://' + env.VERCEL_URL) ||
  'https://mayaexcel.vercel.app';
const site = host.replace(/\/+$/, '');
const repo =
  env.GITHUB_REPO ||
  env.GITHUB_REPOSITORY ||
  (env.VERCEL_GIT_REPO_OWNER && env.VERCEL_GIT_REPO_SLUG && env.VERCEL_GIT_REPO_OWNER + '/' + env.VERCEL_GIT_REPO_SLUG) ||
  'obedelom17/mayaexcel';

rmSync('dist', { recursive: true, force: true });
mkdirSync('dist', { recursive: true });

const SKIP = new Set(['dist', 'scripts', 'node_modules', '.git', '.github', '.vercel', 'README.md', 'package.json', 'package-lock.json', 'vercel.json', '.gitignore', 'desktop', 'android', 'android-res', 'capacitor.config.json']);
for (const name of readdirSync('.')) {
  if (!SKIP.has(name)) cpSync(name, 'dist/' + name, { recursive: true });
}

// Bibliothèques Excel embarquées
mkdirSync('dist/vendor', { recursive: true });
for (const [src, dest] of [
  ['node_modules/xlsx/dist/xlsx.full.min.js', 'dist/vendor/xlsx.full.min.js'],
  ['node_modules/exceljs/dist/exceljs.min.js', 'dist/vendor/exceljs.min.js'],
]) {
  if (existsSync(src)) cpSync(src, dest);
  else console.warn('Absent (npm install non lancé ?) :', src);
}

// Polices embarquées (Figtree + Bricolage Grotesque, variables)
const LATIN = 'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD';
const LATIN_EXT = 'U+0100-02BA,U+02BD-02C5,U+02C7-02CC,U+02CE-02D7,U+02DD-02FF,U+0304,U+0308,U+0329,U+1D00-1DBF,U+1E00-1E9F,U+1EF2-1EFF,U+2020,U+20A0-20AB,U+20AD-20C0,U+2113,U+2C60-2C7F,U+A720-A7FF';
let css = '/* Polices embarquées (généré au déploiement) */\n';
mkdirSync('dist/vendor/fonts', { recursive: true });
for (const [pkg, family] of [['figtree', 'Figtree'], ['bricolage-grotesque', 'Bricolage Grotesque']]) {
  const dir = 'node_modules/@fontsource-variable/' + pkg + '/files';
  if (!existsSync(dir)) { console.warn('Police absente :', pkg); continue; }
  const files = readdirSync(dir);
  for (const [subset, range] of [['latin', LATIN], ['latin-ext', LATIN_EXT]]) {
    const f = files.find((x) => x === pkg + '-' + subset + '-wght-normal.woff2');
    if (!f) continue;
    cpSync(dir + '/' + f, 'dist/vendor/fonts/' + f);
    css += "@font-face{font-family:'" + family + "';font-style:normal;font-display:swap;font-weight:200 800;src:url(/vendor/fonts/" + f + ") format('woff2');unicode-range:" + range + ";}\n";
  }
}
writeFileSync('dist/css/fonts.css', css);

// Adresse du site et dépôt
const html = readFileSync('dist/index.html', 'utf8').replaceAll('__SITE_URL__', site);
writeFileSync('dist/index.html', html);
const cfg = readFileSync('dist/js/config.js', 'utf8').replaceAll('__GITHUB_REPO__', repo);
writeFileSync('dist/js/config.js', cfg);
console.log('MayaExcel construit pour', site, '· dépôt', repo);
