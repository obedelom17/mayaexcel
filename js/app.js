/* MayaExcel – interface */
(function () {
  'use strict';

  const $ = (sel, el) => (el || document).querySelector(sel);
  const $$ = (sel, el) => Array.from((el || document).querySelectorAll(sel));
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const plural = (n, one, many) => n + ' ' + (n > 1 ? many : one);
  const clamp = (n, a, b) => Math.min(b, Math.max(a, n));
  const tick = () => new Promise((r) => setTimeout(r, 0));

  /* ================================================================== */
  /*  Réglages (mémorisés sur l'appareil)                                */
  /* ================================================================== */

  const KEY = 'maya.settings.v3';
  const DEFAULTS = {
    // affichage
    theme: 'system', accent: 'blue', density: 'comfortable', fontScale: 100, motion: 'system',
    // fusion
    allSheets: true, splitNames: true, contentMatch: true, sensitivity: 'balanced',
    headerMode: 'auto', headerRow: 1, combine: true, keyCol: 'auto', conflict: 'first',
    // données
    clean: true, nomCase: 'upper', prenomCase: 'title', emailLower: true, sexNorm: true, placeCase: true,
    phoneFormat: 'keep', phoneCC: '228', dateFormat: 'dmy', emptyText: '', dropEmptyCols: true,
    // tableau
    numbering: true, source: false, sortBy: 'auto', sortDir: 'asc', sortBy2: 'none', sortDir2: 'asc', previewRows: 100,
    // export
    format: 'xlsx', fileName: 'MayaExcel-fusion-{date}', zebra: true, freeze: true, autofilter: true,
    headerColor: 'dark', reportSheet: true, csvSafe: true,
  };

  function loadSettings() {
    const s = Object.assign({}, DEFAULTS);
    try {
      const o = JSON.parse(localStorage.getItem(KEY) || '{}');
      Object.keys(DEFAULTS).forEach((k) => {
        if (k in o && typeof o[k] === typeof DEFAULTS[k]) s[k] = o[k];
      });
    } catch (e) { /* stockage indisponible : réglages par défaut */ }
    return s;
  }
  let S = loadSettings();
  function saveSettings() {
    try { localStorage.setItem(KEY, JSON.stringify(S)); } catch (e) { /* ignoré */ }
  }

  /* ================================================================== */
  /*  Thème et apparence                                                 */
  /* ================================================================== */

  const darkQuery = window.matchMedia('(prefers-color-scheme: dark)');
  const motionQuery = window.matchMedia('(prefers-reduced-motion: reduce)');
  const isDark = () => S.theme === 'dark' || (S.theme === 'system' && darkQuery.matches);
  const reduceMotion = () => S.motion === 'off' || (S.motion === 'system' && motionQuery.matches);

  function applyUi() {
    const r = document.documentElement;
    r.dataset.theme = isDark() ? 'dark' : 'light';
    r.dataset.accent = S.accent;
    r.dataset.density = S.density;
    r.dataset.motion = S.motion;
    r.style.setProperty('--fs', S.fontScale / 100);
    $('#metaTheme').setAttribute('content', isDark() ? '#0a0e17' : '#ffffff');
    $('#themeIcon use').setAttribute('href', isDark() ? '#i-sun' : '#i-moon');
  }
  const onSystemChange = () => S.theme === 'system' && applyUi();
  if (darkQuery.addEventListener) darkQuery.addEventListener('change', onSystemChange);

  /* ================================================================== */
  /*  Bibliothèques Excel (embarquées, sinon CDN)                        */
  /* ================================================================== */

  const LIBS = {
    xlsx: ['/vendor/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'],
    exceljs: ['/vendor/exceljs.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js', 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'],
  };
  const CHECK = { xlsx: () => window.XLSX, exceljs: () => window.ExcelJS };
  const pending = {};

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.onload = resolve;
      s.onerror = () => { s.remove(); reject(new Error('Chargement impossible : ' + src)); };
      document.head.appendChild(s);
    });
  }

  function loadLib(name) {
    if (CHECK[name]()) return Promise.resolve();
    if (!pending[name]) {
      pending[name] = (async () => {
        for (const url of LIBS[name]) {
          try {
            await loadScript(url);
            if (CHECK[name]()) return;
          } catch (e) { /* source suivante */ }
        }
        throw new Error('lib:' + name);
      })().catch((e) => { pending[name] = null; throw e; });
    }
    return pending[name];
  }

  /* ================================================================== */
  /*  État                                                               */
  /* ================================================================== */

  const state = {
    files: [],      // {id,name,size,status:'reading'|'ok'|'error',error,sheets,include}
    cols: [],       // [{id,key,label,include,into}]
    overrides: {},  // key → {label, include, into}
    separate: [],   // intitulés détachés d'un regroupement
    colOrder: [],   // ordre choisi (clés)
    analysis: null,
    built: null,
    quality: null,
    tab: 'table',
    query: '',
    extraRows: 0,
    staging: false,
  };
  let uid = 0;
  const MAX_DOM_ROWS = 5000;

  /* ================================================================== */
  /*  Petits utilitaires d'affichage                                     */
  /* ================================================================== */

  let toastTimer;
  function toast(msg, isError) {
    const el = $('#toast');
    el.textContent = msg;
    el.classList.toggle('is-error', !!isError);
    el.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove('show'), isError ? 5000 : 2800);
  }

  function formatSize(b) {
    if (b < 1024) return b + ' o';
    if (b < 1048576) return Math.round(b / 1024) + ' Ko';
    return (b / 1048576).toFixed(1).replace('.', ',') + ' Mo';
  }

  function colLetter(i) {
    let s = '';
    for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  }

  const MONTHS_FR = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  function fmtDate(d, fmt) {
    const p = (n) => String(n).padStart(2, '0');
    const D = d.getUTCDate();
    const M = d.getUTCMonth() + 1;
    const Y = d.getUTCFullYear();
    let base;
    if (fmt === 'ymd') base = Y + '-' + p(M) + '-' + p(D);
    else if (fmt === 'mdy') base = p(M) + '/' + p(D) + '/' + Y;
    else if (fmt === 'long') base = D + ' ' + MONTHS_FR[M - 1] + ' ' + Y;
    else base = p(D) + '/' + p(M) + '/' + Y;
    return d.getUTCHours() || d.getUTCMinutes() ? base + ' ' + p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) : base;
  }
  const XL_DATE = { dmy: 'dd/mm/yyyy', ymd: 'yyyy-mm-dd', mdy: 'mm/dd/yyyy', long: 'd mmmm yyyy' };

  function cellText(v) {
    if (v === null || v === undefined || v === '') return S.emptyText || '';
    if (v instanceof Date) return fmtDate(v, S.dateFormat);
    if (typeof v === 'boolean') return v ? 'Oui' : 'Non';
    return String(v);
  }

  /* ================================================================== */
  /*  Lecture des fichiers                                               */
  /* ================================================================== */

  const EXT_OK = /\.(xlsx|xlsm|xls|csv|tsv|ods)$/i;

  /** Réduit la plage d'une feuille à ses cellules réellement utilisées (évite les plages fantômes de millions de lignes). */
  function trimRange(ws) {
    const XLSX = window.XLSX;
    let maxR = -1;
    let maxC = -1;
    for (const k in ws) {
      if (k.charAt(0) === '!') continue;
      const c = XLSX.utils.decode_cell(k);
      if (c.r > maxR) maxR = c.r;
      if (c.c > maxC) maxC = c.c;
    }
    if (maxR < 0) delete ws['!ref'];
    else ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  }

  async function readWorkbook(file) {
    await loadLib('xlsx');
    const XLSX = window.XLSX;
    const isText = /\.(csv|tsv)$/i.test(file.name);
    const buf = await file.arrayBuffer();
    let wb;
    if (isText) {
      let text = new TextDecoder('utf-8').decode(buf);
      if (text.indexOf('\uFFFD') !== -1) text = new TextDecoder('windows-1252').decode(buf);
      wb = XLSX.read(text.replace(/^\uFEFF/, ''), { type: 'string', cellDates: true, raw: true });
    } else {
      wb = XLSX.read(buf, { type: 'array', cellDates: true });
    }
    const hidden = (wb.Workbook && wb.Workbook.Sheets) || [];
    const sheets = [];
    wb.SheetNames.forEach((name, i) => {
      if (hidden[i] && hidden[i].Hidden) return;
      const ws = wb.Sheets[name];
      trimRange(ws);
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
      while (rows.length && !(rows[rows.length - 1] || []).some((c) => c !== null && c !== '')) rows.pop();
      if (rows.length) sheets.push({ name, rows });
    });
    return sheets;
  }

  let chain = Promise.resolve();
  function enqueue(task) {
    chain = chain.then(task).catch((e) => { console.error(e); state.staging = false; renderResult(); });
    return chain;
  }

  function addFiles(fileList) {
    const incoming = Array.from(fileList || []);
    if (!incoming.length) return Promise.resolve();
    return enqueue(() => processFiles(incoming));
  }

  async function processFiles(incoming) {
    const fresh = [];
    let skipped = 0;
    let unsupported = 0;
    incoming.forEach((f) => {
      if (!EXT_OK.test(f.name)) return void unsupported++;
      if (state.files.some((x) => x.name === f.name && x.size === f.size)) return void skipped++;
      const entry = { id: ++uid, name: f.name, size: f.size, status: 'reading', error: null, sheets: [], include: true, file: f };
      state.files.push(entry);
      fresh.push(entry);
    });
    if (unsupported) toast(plural(unsupported, 'fichier ignoré', 'fichiers ignorés') + ' : seuls les fichiers Excel et CSV sont acceptés.', true);
    else if (skipped) toast(plural(skipped, 'fichier déjà ajouté', 'fichiers déjà ajoutés') + '.');
    if (!fresh.length) return;

    renderFiles();
    startStage();
    await tick();

    for (const entry of fresh) {
      try {
        entry.sheets = await readWorkbook(entry.file);
        if (!entry.sheets.length) throw new Error('empty');
        entry.status = 'ok';
      } catch (err) {
        entry.status = 'error';
        entry.include = false;
        entry.error =
          err && err.message && err.message.indexOf('lib:') === 0
            ? 'Moteur Excel indisponible. Vérifiez votre connexion Internet.'
            : err && err.message === 'empty'
            ? 'Ce fichier ne contient aucune donnée.'
            : 'Fichier illisible (protégé par un mot de passe ou abîmé ?).';
      }
      delete entry.file;
      renderFiles();
      await tick();
    }
    await endStage();
    recompute(true);
    scrollToResult();
  }

  function scrollToResult() {
    if (state.analysis && state.built && state.built.rows.length) {
      setTimeout(() => $('#resultBlock').scrollIntoView({ behavior: reduceMotion() ? 'auto' : 'smooth', block: 'start' }), 80);
    }
  }

  /** Tableau copié depuis Excel (texte avec tabulations) → fichier virtuel. */
  function addPastedText(text) {
    if (!/[\t\n]/.test(text)) return false;
    const lines = text.replace(/\r/g, '').split('\n');
    while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
    if (lines.length < 2) return false;
    const sep = text.indexOf('\t') !== -1 ? '\t' : lines[0].split(';').length > 1 ? ';' : ',';
    const rows = lines.map((l) => l.split(sep).map((x) => (x.trim() === '' ? null : x.trim())));
    const n = state.files.filter((f) => /^Tableau collé/.test(f.name)).length + 1;
    enqueue(async () => {
      state.files.push({ id: ++uid, name: 'Tableau collé ' + n, size: text.length, status: 'ok', error: null, include: true, sheets: [{ name: 'Collage', rows }] });
      renderFiles();
      startStage();
      await tick();
      await endStage();
      recompute(true);
      scrollToResult();
    });
    return true;
  }

  async function pasteFromButton() {
    try {
      const text = await navigator.clipboard.readText();
      if (!addPastedText(text)) toast('Copiez d’abord un tableau depuis Excel, puis réessayez.', true);
    } catch (e) {
      toast('Collez avec Ctrl + V (⌘ + V sur Mac) quand cette page est active.');
    }
  }

  /** Point d'entrée pour injecter des fichiers déjà lus (tests, exemple). */
  function addParsedFiles(list) {
    list.forEach((f) => state.files.push({ id: ++uid, name: f.name, size: f.size || 4096, status: 'ok', error: null, include: true, sheets: f.sheets }));
    renderFiles();
    recompute(true);
  }

  /* ---------- Scène « fusion en cours » ---------- */
  let stageTimers = [];
  let stageStart = 0;

  function startStage() {
    state.staging = true;
    stageStart = Date.now();
    renderResult();
    const items = $$('#stageSteps li');
    items.forEach((li) => (li.className = ''));
    stageTimers.forEach(clearTimeout);
    stageTimers = [];
    items.forEach((li, i) => {
      stageTimers.push(setTimeout(() => {
        if (i > 0) items[i - 1].className = 'done';
        li.className = 'on';
      }, i * 320));
    });
  }

  async function endStage() {
    const wait = reduceMotion() ? 0 : Math.max(0, 1400 - (Date.now() - stageStart));
    await new Promise((r) => setTimeout(r, wait));
    stageTimers.forEach(clearTimeout);
    $$('#stageSteps li').forEach((li) => (li.className = 'done'));
    await new Promise((r) => setTimeout(r, reduceMotion() ? 0 : 220));
    state.staging = false;
  }

  /* ---------- Exemple : 3 fichiers volontairement différents ---------- */
  function demoFiles() {
    const D = (y, m, d) => new Date(y, m - 1, d);
    return [
      { name: 'RH-Lomé.xlsx', sheets: [{ name: 'Personnel', rows: [
        ['Liste du personnel – Lomé'], [],
        ['Nom', 'Prénoms', 'Tel', 'Poste', "Date d'embauche", 'Sexe'],
        ['AGBODJAN', 'Kofi', '90123456', 'Comptable', D(2020, 1, 15), 'M'],
        ['ADJO', 'Ama', '91000000', 'Assistante RH', '03/02/2021', 'F'],
        ['Total'],
      ] }] },
      { name: 'Équipe-Atakpamé.xlsx', sheets: [{ name: 'Feuil1', rows: [
        ['Nom et Prénoms', 'Téléphone', 'Fonction', 'E-mail', 'Genre', 'Date entrée'],
        ['ABALO Yao', '98001122', 'Informaticien', 'yao.abalo@exemple.tg', 'Masculin', '12 mars 2019'],
        ['Agbodjan Kofi', '90 12 34 56', 'Comptable senior', 'kofi@exemple.tg', 'Homme', '15/01/2020'],
      ] }] },
      { name: 'Saisie-rapide.csv', sheets: [{ name: 'Feuil1', rows: [
        ['KPETO Sena', '92000000', 'sena.kpeto@exemple.tg', 'F', 'Logisticienne'],
        ['DOSSEH Mawuli', '93001234', 'mawuli@exemple.tg', 'M', 'Chauffeur'],
      ] }] },
    ];
  }

  function loadDemo() {
    enqueue(async () => {
      state.files = [];
      state.overrides = {};
      state.separate = [];
      state.colOrder = [];
      demoFiles().forEach((f) => state.files.push({ id: ++uid, name: f.name, size: 4096, status: 'ok', error: null, include: true, sheets: f.sheets }));
      renderFiles();
      startStage();
      await tick();
      await endStage();
      recompute(true);
      setTimeout(() => $('#filesBlock').scrollIntoView({ behavior: reduceMotion() ? 'auto' : 'smooth', block: 'start' }), 80);
    });
  }

  /* ================================================================== */
  /*  Fusion                                                             */
  /* ================================================================== */

  function recompute(enter) {
    const ok = state.files.filter((f) => f.status === 'ok' && f.include !== false);
    if (!ok.length) {
      state.analysis = null;
      state.built = null;
      state.quality = null;
      state.cols = [];
      renderFiles();
      renderResult();
      return;
    }
    state.analysis = MayaMerge.analyze(
      ok.map((f) => ({ id: f.id, name: f.name, sheets: f.sheets })),
      {
        allSheets: S.allSheets, splitNames: S.splitNames, separate: state.separate, sensitivity: S.sensitivity,
        contentMatch: S.contentMatch, headerMode: S.headerMode, headerRow: clamp(Number(S.headerRow) || 1, 1, 200),
      }
    );
    const keyToId = new Map(state.analysis.columns.map((c) => [c.key, c.id]));
    let cols = state.analysis.columns.map((c) => {
      const o = state.overrides[c.key] || {};
      return {
        id: c.id, key: c.key,
        label: o.label != null ? o.label : c.label,
        include: o.include != null ? o.include : !c.autoHide,
        into: o.into && keyToId.has(o.into) ? keyToId.get(o.into) : null,
      };
    });
    if (state.colOrder.length) {
      const idx = new Map(state.colOrder.map((k, i) => [k, i]));
      cols = cols
        .map((c, i) => ({ c, i }))
        .sort((a, b) => (idx.has(a.c.key) ? idx.get(a.c.key) : 1e6 + a.i) - (idx.has(b.c.key) ? idx.get(b.c.key) : 1e6 + b.i))
        .map((x) => x.c);
    }
    state.cols = cols;
    renderFiles();
    rebuild(true, enter);
  }

  function colIdByKey(key) {
    const c = state.cols.find((x) => x.key === key);
    return c ? c.id : null;
  }

  function buildSettings() {
    const sortBy = S.sortBy === 'auto' || S.sortBy === 'none' ? S.sortBy : colIdByKey(S.sortBy) || 'auto';
    return {
      cols: state.cols, source: S.source, numbering: S.numbering, combine: S.combine,
      keyCol: S.keyCol === 'auto' ? 'auto' : colIdByKey(S.keyCol) || 'auto', conflict: S.conflict,
      clean: S.clean, nomCase: S.nomCase, prenomCase: S.prenomCase, emailLower: S.emailLower, sexNorm: S.sexNorm,
      placeCase: S.placeCase, phoneFormat: S.phoneFormat, phoneCC: S.phoneCC, dropEmptyCols: S.dropEmptyCols,
      sortBy, sortDir: S.sortDir, sortBy2: S.sortBy2 === 'none' ? 'none' : colIdByKey(S.sortBy2) || 'none', sortDir2: S.sortDir2,
    };
  }

  function rebuild(full, enter) {
    if (!state.analysis) return;
    state.built = MayaMerge.build(state.analysis, buildSettings());
    state.quality = computeQuality();
    if (full) renderResult(enter);
    else {
      renderSummary();
      renderTable(true);
      renderCheckBadge();
      if (state.tab === 'check') renderCheck();
      updateActions();
    }
  }

  /* ================================================================== */
  /*  Contrôle qualité                                                   */
  /* ================================================================== */

  function computeQuality() {
    const b = state.built;
    if (!b) return null;
    const n = b.rows.length;
    const issues = [];
    const comp = [];
    const bad = new Set();
    const now = Date.now();
    const YEAR = 31557600000;
    const add = (r, c, t, v) => {
      issues.push({ r, c, t, v });
      bad.add(r + ':' + c);
    };
    let matCol = -1;
    const dup = new Map();

    b.headers.forEach((h, j) => {
      const g = b.headMeta[j].group;
      if (g === 'index' || g === 'source') return;
      if (g === 'matricule') matCol = j;
      let filled = 0;
      for (let i = 0; i < n; i++) {
        const v = b.rows[i][j];
        if (v === null || v === undefined || v === '') continue;
        filled++;
        if (g === 'email' && typeof v === 'string' && !MayaMerge.EMAIL_RE.test(v)) add(i, j, 'E-mail invalide', v);
        else if (g === 'telephone' && !MayaMerge.phoneDigits(v)) add(i, j, 'Numéro de téléphone inhabituel', v);
        else if (g === 'naissance' && v instanceof Date) {
          const age = (now - v.getTime()) / YEAR;
          if (age < 14 || age > 100) add(i, j, 'Date de naissance improbable', fmtDate(v, S.dateFormat));
        } else if (g === 'dateembauche' && v instanceof Date && v.getTime() > now + YEAR) add(i, j, 'Date d’embauche dans le futur', fmtDate(v, S.dateFormat));
        else if (g === 'salaire' && typeof v === 'number' && v <= 0) add(i, j, 'Salaire nul ou négatif', v);
        if (g === 'matricule') {
          const k = String(v).toLowerCase();
          if (!dup.has(k)) dup.set(k, []);
          dup.get(k).push(i);
        }
      }
      comp.push({ label: h, filled, pct: n ? filled / n : 0 });
    });
    dup.forEach((rows, k) => {
      if (rows.length > 1) rows.forEach((i) => add(i, matCol, 'Matricule en double', b.rows[i][matCol]));
    });
    Object.keys(b.conflictMap).forEach((key) => {
      const [r, c] = key.split(':').map(Number);
      issues.push({ r, c, t: 'Valeurs différentes selon les fichiers', v: 'autre valeur : ' + b.conflictMap[key], warn: true });
    });
    issues.sort((a, b2) => a.r - b2.r || a.c - b2.c);
    return { issues, comp, bad };
  }

  /* ================================================================== */
  /*  Rendu : fichiers                                                   */
  /* ================================================================== */

  function renderFiles() {
    document.body.classList.toggle('has-files', state.files.length > 0);
    const block = $('#filesBlock');
    block.hidden = !state.files.length;
    if (!state.files.length) return;
    $('#filesCount').textContent = '(' + state.files.length + ')';
    $('#filesHint').hidden = state.files.length < 2;

    const reports = new Map();
    if (state.analysis) state.analysis.report.forEach((r) => reports.set(r.id, r));
    const last = state.files.length - 1;

    $('#filesList').innerHTML = state.files
      .map((f, i) => {
        const ext = (f.name.split('.').pop() || '').toLowerCase().slice(0, 4);
        let meta;
        if (f.status === 'reading') meta = '<span class="spin"></span>Lecture en cours…';
        else if (f.status === 'error') meta = esc(f.error);
        else if (f.include === false) meta = 'Exclu de la fusion';
        else {
          const r = reports.get(f.id);
          if (r) {
            const used = r.sheets.filter((s) => !s.empty).length;
            meta = r.rows ? plural(r.rows, 'ligne', 'lignes') + ' · ' + plural(used, 'feuille', 'feuilles') : 'Aucune donnée trouvée';
            if (r.sheets.some((s) => s.headerless)) meta += ' · sans en-têtes';
          } else meta = formatSize(f.size);
        }
        const check = f.status === 'ok'
          ? '<input type="checkbox" data-act="toggle" data-id="' + f.id + '" ' + (f.include !== false ? 'checked' : '') + ' aria-label="Inclure ' + esc(f.name) + ' dans la fusion">'
          : '<span style="width:20px"></span>';
        return (
          '<li class="file' + (f.status === 'error' ? ' is-error' : '') + (f.status === 'reading' ? ' is-reading' : '') + (f.include === false ? ' is-off' : '') + '" data-id="' + f.id + '">' +
          check +
          '<span class="file-ext" aria-hidden="true">' + (/^Tableau collé/.test(f.name) ? 'txt' : esc(ext)) + '</span>' +
          '<div class="file-body"><div class="file-name" title="' + esc(f.name) + '">' + esc(f.name) + '</div><div class="file-meta">' + meta + '</div></div>' +
          '<div class="file-tools">' +
          '<button type="button" class="icon-btn" data-act="up" data-id="' + f.id + '" aria-label="Monter ' + esc(f.name) + '" ' + (i === 0 ? 'disabled' : '') + '><svg class="ic"><use href="#i-up"/></svg></button>' +
          '<button type="button" class="icon-btn" data-act="down" data-id="' + f.id + '" aria-label="Descendre ' + esc(f.name) + '" ' + (i === last ? 'disabled' : '') + '><svg class="ic"><use href="#i-down"/></svg></button>' +
          '<button type="button" class="icon-btn" data-act="remove" data-id="' + f.id + '" aria-label="Retirer ' + esc(f.name) + '"><svg class="ic"><use href="#i-x"/></svg></button>' +
          '</div></li>'
        );
      })
      .join('');
  }

  /* ================================================================== */
  /*  Rendu : résultat                                                   */
  /* ================================================================== */

  function renderResult(enter) {
    const block = $('#resultBlock');
    if (state.staging) {
      block.hidden = false;
      $('#stage').hidden = false;
      $('#resultBody').hidden = true;
      return;
    }
    $('#stage').hidden = true;
    block.hidden = !state.analysis;
    $('#resultBody').hidden = !state.analysis;
    if (!state.analysis) return;

    renderToolbar();
    renderSummary();
    renderTable(true);
    renderColumns();
    renderCheck();
    renderCheckBadge();
    renderActions();
    updateActions();
    setTab(state.tab, true);
    if (enter && !reduceMotion()) {
      const body = $('#resultBody');
      body.classList.remove('enter');
      void body.offsetWidth;
      body.classList.add('enter');
    }
  }

  function animateNumber(el, to) {
    if (reduceMotion()) { el.textContent = to; return; }
    const t0 = performance.now();
    (function step(now) {
      const k = Math.min(1, (now - t0) / 700);
      el.textContent = Math.round(to * (1 - Math.pow(1 - k, 3)));
      if (k < 1) requestAnimationFrame(step);
    })(t0);
  }
  const num = (n) => '<span class="cnt" data-to="' + n + '">' + n + '</span>';

  function renderSummary() {
    const a = state.analysis;
    const b = state.built;
    if (!a || !b) return;
    const w = (n, one, many) => (n > 1 ? many : one);
    const notes = [];
    if (b.sortedBy.length) notes.push('trié par ' + b.sortedBy.join(', '));
    $('#summary').innerHTML =
      '<strong>' + num(a.fileCount) + ' ' + w(a.fileCount, 'fichier', 'fichiers') + '</strong> fusionnés en un tableau de <strong>' +
      num(b.rows.length) + ' ' + w(b.rows.length, 'ligne', 'lignes') + '</strong> et <strong>' + num(b.columnCount) + ' ' + w(b.columnCount, 'colonne', 'colonnes') + '</strong>.' +
      (notes.length ? '<small>' + esc(notes.join(' · ')) + '</small>' : '');
    $$('#summary .cnt').forEach((el) => animateNumber(el, Number(el.dataset.to)));

    const chips = [];
    if (a.mergedColumns) chips.push(['', plural(a.mergedColumns, 'colonne regroupée', 'colonnes regroupées') + ' malgré des noms différents']);
    if (a.contentColumns) chips.push(['', plural(a.contentColumns, 'colonne reconnue', 'colonnes reconnues') + ' d’après son contenu']);
    if (a.headerlessSheets) chips.push(['', plural(a.headerlessSheets, 'fichier compris', 'fichiers compris') + ' sans titres de colonnes']);
    if (a.splitCount && S.splitNames) chips.push(['', plural(a.splitCount, 'nom complet séparé', 'noms complets séparés') + ' en Nom et Prénom']);
    if (b.combinedRows) chips.push(['', plural(b.combinedRows, 'personne trouvée', 'personnes trouvées') + ' dans plusieurs fichiers']);
    else if (b.absorbed) chips.push(['', plural(b.absorbed, 'doublon retiré', 'doublons retirés')]);
    if (b.emptyDropped) chips.push(['', plural(b.emptyDropped, 'colonne vide masquée', 'colonnes vides masquées')]);
    if (b.conflictCount) chips.push(['warn', plural(b.conflictCount, 'valeur différente', 'valeurs différentes') + ' à vérifier']);
    $('#insights').innerHTML = chips.map(([c, t]) => '<li class="' + c + '">' + (c === 'warn' ? '<span class="dot"></span>' : '') + esc(t) + '</li>').join('');
  }

  function sortOptionsHtml(current, withAuto) {
    let html = withAuto ? '<option value="auto">Automatique' + (autoSortLabel() ? ' (' + esc(autoSortLabel()) + ')' : '') + '</option>' : '<option value="none">Aucun</option>';
    if (withAuto) html += '<option value="none">Ordre des fichiers</option>';
    state.cols.filter((c) => c.include && !c.into).forEach((c) => {
      html += '<option value="' + esc(c.key) + '">' + esc(c.label || 'Colonne') + '</option>';
    });
    void current;
    return html;
  }

  function autoSortLabel() {
    if (!state.analysis) return '';
    const meta = new Map(state.analysis.columns.map((c) => [c.id, c]));
    const has = (g) => state.cols.find((c) => c.include && !c.into && meta.get(c.id).group === g);
    if (has('nom')) return has('prenom') ? 'Nom, Prénom' : 'Nom';
    if (has('nomcomplet')) return 'Nom complet';
    if (has('matricule')) return 'Matricule';
    return '';
  }

  function renderToolbar() {
    const sel = $('#sortBy');
    sel.innerHTML = sortOptionsHtml(S.sortBy, true);
    if (![...sel.options].some((o) => o.value === S.sortBy)) sel.value = 'auto';
    else sel.value = S.sortBy;
    const dir = $('#sortDir');
    dir.textContent = S.sortDir === 'asc' ? 'A → Z' : 'Z → A';
    dir.disabled = sel.value === 'none';
    $('#searchInput').value = state.query;
  }

  function visibleRows() {
    const b = state.built;
    const q = state.query.trim().toLowerCase();
    const idx = [];
    for (let i = 0; i < b.rows.length; i++) {
      if (!q || b.rows[i].some((v) => cellText(v).toLowerCase().indexOf(q) !== -1)) idx.push(i);
    }
    return idx;
  }

  function renderTable(animate) {
    const b = state.built;
    const box = $('#sheet');
    if (!b || !b.rows.length) {
      box.innerHTML = '<p class="empty-note">Aucune ligne à afficher. Vérifiez que des colonnes sont affichées.</p>';
      $('#sheetNote').textContent = '';
      return;
    }
    const q = state.query.trim().toLowerCase();
    const matches = visibleRows();
    const limit = clamp(Math.max(Number(S.previewRows) || 100, state.extraRows), 10, MAX_DOM_ROWS);
    const shown = matches.slice(0, limit);
    const bad = state.quality ? state.quality.bad : new Set();

    if (!shown.length) {
      box.innerHTML = '<p class="empty-note">Aucun résultat pour « ' + esc(state.query) + ' ».</p>';
      $('#sheetNote').textContent = '';
      return;
    }

    let html = '<table class="grid"><thead><tr class="letters"><th class="rn"></th>';
    b.headers.forEach((_, j) => (html += '<th scope="col">' + colLetter(j) + '</th>'));
    html += '</tr><tr class="heads"><th class="rn"></th>';
    b.headers.forEach((h, j) => {
      const m = b.headMeta[j];
      const dot = m.merged ? '<span class="merged-dot" title="Regroupe : ' + esc(m.variants.join(', ')) + '"></span>' : '';
      html += '<th scope="col">' + esc(h) + dot + '</th>';
    });
    html += '</tr></thead><tbody>';
    shown.forEach((ri, k) => {
      html += '<tr style="--i:' + k + '"><td class="rn">' + (ri + 1) + '</td>';
      b.rows[ri].forEach((v, j) => {
        const t = cellText(v);
        const alt = b.conflictMap[ri + ':' + j];
        const cls = [];
        if (alt) cls.push('conf');
        if (bad.has(ri + ':' + j)) cls.push('bad');
        if (q && t.toLowerCase().indexOf(q) !== -1) cls.push('hit');
        html += '<td' + (cls.length ? ' class="' + cls.join(' ') + '"' : '') + ' data-r="' + ri + '" data-c="' + j + '" data-alt="' + (alt ? esc(alt) : '') + '" title="' + esc(alt ? t + ' — autre valeur : ' + alt : t) + '">' + esc(t) + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    box.innerHTML = html;
    if (animate && !reduceMotion()) {
      box.classList.remove('cascade');
      void box.offsetWidth;
      box.classList.add('cascade');
    }
    $('#fbRef').textContent = 'A1';
    $('#fbVal').textContent = 'Touchez une cellule pour voir son contenu';
    $('#fbVal').classList.remove('has');

    let note;
    if (q) note = plural(matches.length, 'ligne correspond', 'lignes correspondent') + ' à la recherche' + (matches.length > limit ? ' (' + limit + ' affichées)' : '') + '.';
    else if (b.rows.length > limit) note = 'Aperçu des ' + limit + ' premières lignes sur ' + b.rows.length + '. Le fichier téléchargé contient tout (changez la taille de l’aperçu dans les réglages).';
    else note = 'Aperçu complet. Faites défiler pour voir toutes les colonnes.';
    $('#sheetNote').textContent = note;
  }

  function selectCell(td) {
    const prev = $('#sheet td.sel');
    if (prev) prev.classList.remove('sel');
    td.classList.add('sel');
    const ref = colLetter(Number(td.dataset.c)) + (Number(td.dataset.r) + 1);
    $('#fbRef').textContent = ref;
    const fv = $('#fbVal');
    const alt = td.dataset.alt;
    fv.textContent = (td.textContent || '(vide)') + (alt ? '   ·   autre valeur : ' + alt : '');
    fv.classList.add('has');
  }

  function gotoCell(r, c) {
    setTab('table');
    state.query = '';
    $('#searchInput').value = '';
    state.extraRows = Math.max(state.extraRows, r + 1);
    renderTable(false);
    const td = $('#sheet td[data-r="' + r + '"][data-c="' + c + '"]');
    if (td) {
      selectCell(td);
      td.scrollIntoView({ block: 'center', inline: 'center', behavior: reduceMotion() ? 'auto' : 'smooth' });
    } else toast('Cette ligne est au-delà de l’aperçu (' + MAX_DOM_ROWS + ' lignes maximum).');
  }

  /* ---------- Onglet Colonnes ---------- */
  const REASON_TXT = { contenu: 'reconnue d’après son contenu', proche: 'rapprochée (intitulé proche)', 'déduit': 'déduite du nom complet' };

  function renderColumns() {
    const a = state.analysis;
    if (!a) return;
    const meta = new Map(a.columns.map((c) => [c.id, c]));
    const total = a.fileCount;
    $('#colsList').innerHTML = state.cols
      .map((c, i) => {
        const m = meta.get(c.id);
        const chips = m.variants
          .map((v, vi) => {
            const k = m.variantKeys[vi];
            const cut = m.variants.length > 1 && k ? '<button type="button" data-act="detach" data-vk="' + esc(k) + '" title="Séparer « ' + esc(v) + ' » de cette colonne" aria-label="Séparer ' + esc(v) + '">✂</button>' : '';
            return '<span class="chip' + (cut ? '' : ' solo') + '">' + esc(v) + cut + '</span>';
          })
          .join('');
        const badges = m.reasons.filter((r) => REASON_TXT[r]).map((r) => '<span class="badge">' + esc(REASON_TXT[r]) + '</span>').join('');
        const others = state.cols.filter((x) => x.id !== c.id && !x.into);
        const select = '<div class="into"><label for="into-' + c.id + '">Rattacher à</label><select id="into-' + c.id + '" data-act="into"><option value="">— (colonne séparée)</option>' +
          others.map((x) => '<option value="' + x.id + '"' + (c.into === x.id ? ' selected' : '') + '>' + esc(x.label) + '</option>').join('') + '</select></div>';
        return (
          '<li class="col-row' + (c.include ? '' : ' is-off') + '" data-id="' + c.id + '">' +
          '<input type="checkbox" data-act="include" ' + (c.include ? 'checked' : '') + ' aria-label="Afficher la colonne ' + esc(c.label) + '">' +
          '<input type="text" class="col-label" data-act="label" value="' + esc(c.label) + '" aria-label="Nom de la colonne" autocomplete="off">' +
          '<div class="col-move">' +
          '<button type="button" class="icon-btn" data-act="up" aria-label="Monter" ' + (i === 0 ? 'disabled' : '') + '><svg class="ic"><use href="#i-up"/></svg></button>' +
          '<button type="button" class="icon-btn" data-act="down" aria-label="Descendre" ' + (i === state.cols.length - 1 ? 'disabled' : '') + '><svg class="ic"><use href="#i-down"/></svg></button>' +
          '</div>' +
          '<div class="col-info"><div class="chips">' + chips + badges + '</div>' +
          '<span>Présente dans ' + m.fileCount + ' fichier' + (m.fileCount > 1 ? 's' : '') + ' sur ' + total + '</span>' + select + '</div>' +
          '</li>'
        );
      })
      .join('');
  }

  /* ---------- Onglet Contrôle ---------- */
  function renderCheckBadge() {
    const q = state.quality;
    const el = $('#checkCount');
    const n = q ? q.issues.length : 0;
    el.hidden = !n;
    el.textContent = n > 99 ? '99+' : n;
  }

  function renderCheck() {
    const q = state.quality;
    const a = state.analysis;
    const panel = $('#panelCheck');
    if (!q || !a) { panel.innerHTML = ''; return; }

    const bars = q.comp
      .map((c) => {
        const pct = Math.round(c.pct * 100);
        return '<div class="bar-row"><span title="' + esc(c.label) + '">' + esc(c.label) + '</span><div class="bar"><i class="' + (pct < 60 ? 'low' : '') + '" style="width:' + pct + '%"></i></div><span class="pct">' + pct + ' %</span></div>';
      })
      .join('');

    const shown = q.issues.slice(0, 200);
    const issues = q.issues.length
      ? '<ul class="issues">' +
        shown.map((it) => {
          const col = state.built.headers[it.c] || '';
          return '<li><button type="button" class="issue" data-goto="' + it.r + ':' + it.c + '"><svg class="ic"><use href="#i-alert"/></svg><span><b>' + esc(it.t) + '</b><small>' + esc(col) + ' · ligne ' + (it.r + 1) + ' · ' + esc(String(it.v)) + '</small></span></button></li>';
        }).join('') + '</ul>' + (q.issues.length > shown.length ? '<p class="sheet-note">' + (q.issues.length - shown.length) + ' autres anomalies non listées.</p>' : '')
      : '<p class="ok-box"><svg class="ic"><use href="#i-check"/></svg>Aucune anomalie détectée.</p>';

    const rows = a.report
      .map((r) => {
        const used = r.sheets.filter((s) => !s.empty);
        const detail = r.error ? '<span class="tag err">erreur</span>' : used.map((s) => esc(s.name) + (s.headerless ? ' <span class="tag">sans en-têtes</span>' : ' <span class="tag">ligne ' + s.headerRow + '</span>')).join(', ') || '—';
        return '<tr><td>' + esc(r.name) + '</td><td>' + detail + '</td><td>' + r.rows + '</td></tr>';
      })
      .join('');

    panel.innerHTML =
      '<div class="check-grid">' +
      '<div class="card"><h3>Remplissage des colonnes</h3><div class="bars">' + (bars || '<p class="sheet-note">Aucune colonne.</p>') + '</div></div>' +
      '<div class="card"><h3>Anomalies (' + q.issues.length + ')</h3>' + issues + '</div>' +
      '<div class="card check-wide"><h3>Fichiers lus</h3><table class="report-table"><thead><tr><th>Fichier</th><th>Feuilles et ligne d’en-têtes</th><th>Lignes</th></tr></thead><tbody>' + rows + '</tbody></table></div>' +
      '</div>';
  }

  /* ---------- Onglets ---------- */
  function setTab(tab, silent) {
    state.tab = tab;
    $$('.tab').forEach((t) => {
      const on = t.dataset.tab === tab;
      t.setAttribute('aria-selected', on ? 'true' : 'false');
      t.tabIndex = on ? 0 : -1;
    });
    $('#panelTable').hidden = tab !== 'table';
    $('#panelCols').hidden = tab !== 'cols';
    $('#panelCheck').hidden = tab !== 'check';
    $('#toolbar .search').style.visibility = tab === 'table' ? '' : 'hidden';
    if (!silent) {
      if (tab === 'cols') renderColumns();
      if (tab === 'check') renderCheck();
    }
  }

  /* ---------- Actions / export ---------- */
  const FORMATS = {
    xlsx: { label: 'xlsx', ext: 'xlsx', name: 'Excel', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    xls: { label: 'xls', ext: 'xls', name: 'Excel 97-2003', mime: 'application/vnd.ms-excel' },
    ods: { label: 'ods', ext: 'ods', name: 'LibreOffice', mime: 'application/vnd.oasis.opendocument.spreadsheet' },
    csv: { label: 'csv ;', ext: 'csv', name: 'CSV (;)', mime: 'text/csv;charset=utf-8' },
    csvc: { label: 'csv ,', ext: 'csv', name: 'CSV (,)', mime: 'text/csv;charset=utf-8' },
    tsv: { label: 'tsv', ext: 'tsv', name: 'TSV', mime: 'text/tab-separated-values;charset=utf-8' },
    json: { label: 'json', ext: 'json', name: 'JSON', mime: 'application/json' },
  };

  function renderActions() {
    const sel = $('#formatSel');
    sel.innerHTML = Object.keys(FORMATS).map((k) => '<option value="' + k + '">' + esc(FORMATS[k].name) + '</option>').join('');
    if (!FORMATS[S.format]) S.format = 'xlsx';
    sel.value = S.format;
    $('#dlLabel').textContent = 'Télécharger (.' + FORMATS[S.format].ext + ')';
    $('#shareBtn').hidden = !canShareFiles();
  }

  function updateActions() {
    const ok = !!(state.built && state.built.rows.length);
    ['#dlBtn', '#copyBtn', '#shareBtn', '#formatSel'].forEach((s) => { const el = $(s); if (el) el.disabled = !ok; });
  }

  function canShareFiles() {
    try {
      return !!(navigator.canShare && navigator.share && navigator.canShare({ files: [new File(['x'], 'a.xlsx', { type: FORMATS.xlsx.mime })] }));
    } catch (e) { return false; }
  }

  function stamp(kind) {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return kind === 'time' ? p(d.getHours()) + 'h' + p(d.getMinutes()) : d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }

  function baseFileName() {
    const b = state.built;
    let name = String(S.fileName || '')
      .replace(/\{date\}/g, stamp())
      .replace(/\{heure\}/g, stamp('time'))
      .replace(/\{n\}/g, b ? b.rows.length : 0)
      .replace(/\{fichiers\}/g, state.analysis ? state.analysis.fileCount : 0)
      .replace(/[\\/:*?"<>|]+/g, '-')
      .trim();
    return name || 'MayaExcel-fusion-' + stamp();
  }

  function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(String(fr.result).split(',')[1]);
      fr.onerror = reject;
      fr.readAsDataURL(blob);
    });
  }

  async function download(blob, name) {
    // Application Android : pas de téléchargement classique, on passe par le partage du téléphone.
    const cap = window.Capacitor;
    if (cap && cap.isNativePlatform && cap.isNativePlatform()) {
      try {
        const Fs = (cap.Plugins && cap.Plugins.Filesystem) || (cap.registerPlugin && cap.registerPlugin('Filesystem'));
        const Share = (cap.Plugins && cap.Plugins.Share) || (cap.registerPlugin && cap.registerPlugin('Share'));
        const saved = await Fs.writeFile({ path: name, data: await blobToBase64(blob), directory: 'CACHE' });
        await Share.share({ title: name, url: saved.uri, dialogTitle: 'Enregistrer ou ouvrir le fichier' });
        return;
      } catch (e) {
        if (/cancel/i.test(String(e && e.message))) return;
      }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  /* ---- Construction des fichiers ---- */
  function rowsForExport(b, forText) {
    return b.rows.map((r) =>
      r.map((v) => {
        if (v === null || v === undefined || v === '') return S.emptyText || null;
        if (forText && v instanceof Date) return fmtDate(v, S.dateFormat);
        return v;
      })
    );
  }

  function toLocalDate(d) {
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes());
  }

  function accentArgb() {
    return { blue: 'FF3355FF', indigo: 'FF5B3DF5', green: 'FF0A8F5B', orange: 'FFE8590C', rose: 'FFD6336C', slate: 'FF334155' }[S.accent] || 'FF3355FF';
  }

  function reportRows() {
    const a = state.analysis;
    const b = state.built;
    const rows = [['Fichier', 'Feuilles lues', 'Lignes']];
    a.report.forEach((r) => rows.push([r.name, r.error ? 'Erreur' : r.sheets.filter((s) => !s.empty).length, r.rows]));
    rows.push([]);
    rows.push(['Total', '', b.rows.length]);
    return rows;
  }

  async function makeXlsx(b) {
    await loadLib('exceljs');
    const wb = new window.ExcelJS.Workbook();
    wb.creator = 'MayaExcel';
    wb.created = new Date();
    const ws = wb.addWorksheet('Fusion', { views: S.freeze ? [{ state: 'frozen', ySplit: 1 }] : [] });
    const data = rowsForExport(b, false);
    ws.addRow(b.headers);
    data.forEach((r) => ws.addRow(r));

    const heavy = data.length > 20000;
    const line = { style: 'thin', color: { argb: 'FFDDE1E8' } };
    const head = ws.getRow(1);
    head.height = 24;
    const hc = S.headerColor;
    head.eachCell((cell) => {
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      if (hc === 'none') { cell.font = { bold: true }; return; }
      const dark = hc === 'dark' || hc === 'accent';
      cell.font = { bold: true, color: { argb: dark ? 'FFFFFFFF' : 'FF0B1220' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hc === 'dark' ? 'FF0B1220' : hc === 'accent' ? accentArgb() : 'FFE9ECF2' } };
      cell.border = { bottom: line };
    });

    b.headers.forEach((h, j) => {
      let width = String(h).length;
      let hasDate = false;
      const sample = Math.min(data.length, 2000);
      for (let i = 0; i < sample; i++) {
        const v = data[i][j];
        if (v instanceof Date) { hasDate = true; width = Math.max(width, 16); }
        else if (v !== null && v !== undefined) width = Math.max(width, String(v).length);
      }
      ws.getColumn(j + 1).width = clamp(width + 3, 10, 48);
      if (hasDate) {
        const fmt = XL_DATE[S.dateFormat] || XL_DATE.dmy;
        ws.getColumn(j + 1).eachCell({ includeEmpty: false }, (cell, rn) => {
          if (rn > 1 && cell.value instanceof Date) {
            cell.numFmt = cell.value.getUTCHours() || cell.value.getUTCMinutes() ? fmt + ' hh:mm' : fmt;
            cell.alignment = { horizontal: 'left' };
          }
        });
      }
    });

    for (let i = 0; i < data.length; i++) {
      const row = ws.getRow(i + 2);
      for (let j = 0; j < b.headers.length; j++) {
        const cell = row.getCell(j + 1);
        if (!heavy) {
          cell.border = { bottom: line };
          if (S.zebra && i % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F6F9' } };
        }
        const alt = b.conflictMap[i + ':' + j];
        if (alt) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF4DB' } };
          cell.note = 'Autre valeur dans un autre fichier : ' + alt;
        }
      }
    }
    if (S.autofilter && b.headers.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: b.headers.length } };

    if (S.reportSheet) {
      const ws2 = wb.addWorksheet('Rapport');
      reportRows().forEach((r) => ws2.addRow(r));
      ws2.getRow(1).font = { bold: true };
      ws2.getColumn(1).width = 44;
      ws2.getColumn(2).width = 16;
      ws2.getColumn(3).width = 12;
      const keys = Object.keys(b.conflictMap);
      if (keys.length) {
        ws2.addRow([]);
        const t = ws2.addRow(['Valeurs différentes selon les fichiers', 'Colonne', 'Valeur gardée', 'Autre valeur']);
        t.font = { bold: true };
        ws2.getColumn(4).width = 30;
        keys.forEach((k) => {
          const [r, c] = k.split(':').map(Number);
          ws2.addRow(['Ligne ' + (r + 1), b.headers[c], cellText(b.rows[r][c]), b.conflictMap[k]]);
        });
      }
    }
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: FORMATS.xlsx.mime });
  }

  async function makeSheetJs(b, bookType) {
    await loadLib('xlsx');
    const XLSX = window.XLSX;
    const data = [b.headers].concat(rowsForExport(b, false).map((r) => r.map((v) => (v instanceof Date ? toLocalDate(v) : v))));
    const ws = XLSX.utils.aoa_to_sheet(data, { cellDates: true });
    ws['!cols'] = b.headers.map((h, j) => ({ wch: clamp(Math.max(String(h).length, ...data.slice(1, 300).map((r) => (r[j] == null ? 0 : String(r[j]).length))) + 2, 8, 48) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Fusion');
    const out = XLSX.write(wb, { bookType, type: 'array', cellDates: true });
    return new Blob([out], { type: FORMATS[bookType === 'biff8' ? 'xls' : bookType].mime });
  }

  function delimited(b, sep) {
    const safe = (s) => (S.csvSafe && /^[=+\-@]/.test(s) && !/^[-+]?\d+([.,]\d+)?$/.test(s) ? "'" + s : s);
    const q = (v) => {
      const s = safe(cellText(v));
      return s.indexOf(sep) !== -1 || /["\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    return [b.headers.map((h) => q(h)).join(sep)].concat(b.rows.map((r) => r.map(q).join(sep))).join('\r\n');
  }

  function toJson(b) {
    const keys = [];
    b.headers.forEach((h) => {
      let k = h;
      let n = 2;
      while (keys.includes(k)) k = h + ' (' + n++ + ')';
      keys.push(k);
    });
    const out = b.rows.map((r) => {
      const o = {};
      r.forEach((v, j) => { o[keys[j]] = v instanceof Date ? v.toISOString().slice(0, v.getUTCHours() || v.getUTCMinutes() ? 16 : 10) : v === undefined ? null : v; });
      return o;
    });
    return JSON.stringify(out, null, 2);
  }

  async function makeExport() {
    const b = state.built;
    const fmt = FORMATS[S.format] ? S.format : 'xlsx';
    const name = baseFileName() + '.' + FORMATS[fmt].ext;
    let blob;
    if (fmt === 'xlsx') {
      try { blob = await makeXlsx(b); } catch (e) { blob = await makeSheetJs(b, 'xlsx'); }
    } else if (fmt === 'xls') blob = await makeSheetJs(b, 'biff8');
    else if (fmt === 'ods') blob = await makeSheetJs(b, 'ods');
    else if (fmt === 'csv') blob = new Blob(['\uFEFF' + delimited(b, ';')], { type: FORMATS.csv.mime });
    else if (fmt === 'csvc') blob = new Blob(['\uFEFF' + delimited(b, ',')], { type: FORMATS.csvc.mime });
    else if (fmt === 'tsv') blob = new Blob(['\uFEFF' + delimited(b, '\t')], { type: FORMATS.tsv.mime });
    else blob = new Blob([toJson(b)], { type: FORMATS.json.mime });
    return { blob, name };
  }

  function burst(anchor) {
    if (reduceMotion() || !anchor) return;
    const r = anchor.getBoundingClientRect();
    const colors = ['#3355ff', '#ffb020', '#0b1220', '#7c93ff'];
    for (let i = 0; i < 22; i++) {
      const el = document.createElement('i');
      el.className = 'confetti';
      const ang = Math.random() * Math.PI - Math.PI;
      const dist = 60 + Math.random() * 120;
      el.style.left = r.left + r.width / 2 + 'px';
      el.style.top = r.top + r.height / 2 + 'px';
      el.style.background = colors[i % colors.length];
      el.style.setProperty('--dx', Math.cos(ang) * dist + 'px');
      el.style.setProperty('--dy', Math.sin(ang) * dist - 20 + 'px');
      el.style.setProperty('--rot', Math.random() * 360 + 'deg');
      document.body.appendChild(el);
      setTimeout(() => el.remove(), 1000);
    }
  }

  async function doDownload() {
    if (!state.built || !state.built.rows.length) return;
    const btn = $('#dlBtn');
    const label = $('#dlLabel');
    const old = label.textContent;
    btn.disabled = true;
    label.textContent = 'Préparation…';
    try {
      const { blob, name } = await makeExport();
      await download(blob, name);
      toast('Fichier prêt : ' + name);
      burst(btn);
    } catch (e) {
      console.error(e);
      toast('Impossible de créer le fichier. Vérifiez votre connexion Internet et réessayez.', true);
    } finally {
      label.textContent = old;
      updateActions();
    }
  }

  async function doCopy() {
    const b = state.built;
    if (!b) return;
    const text = [b.headers.join('\t')].concat(b.rows.map((r) => r.map((v) => cellText(v).replace(/[\t\r\n]+/g, ' ')).join('\t'))).join('\n');
    try {
      await navigator.clipboard.writeText(text);
    } catch (e) {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); } catch (e2) { /* ignoré */ }
      ta.remove();
    }
    toast('Tableau copié : collez-le dans Excel.');
  }

  async function doShare() {
    try {
      const { blob, name } = await makeExport();
      const file = new File([blob], name, { type: blob.type });
      await navigator.share({ files: [file], title: name });
    } catch (e) {
      if (e && e.name !== 'AbortError') toast('Le partage n’est pas disponible ici.', true);
    }
  }

  /* ================================================================== */
  /*  Panneau de réglages                                                */
  /* ================================================================== */

  const SECTION_ORDER = ['fusion', 'data', 'table', 'export', 'look', 'app'];
  let settingsSection = 'fusion';

  function colOptions(withAuto, noneLabel) {
    const o = [];
    if (withAuto) o.push({ v: 'auto', l: 'Automatique' });
    o.push({ v: 'none', l: noneLabel || 'Aucun' });
    state.cols.filter((c) => c.include && !c.into).forEach((c) => o.push({ v: c.key, l: c.label }));
    return o;
  }

  function schema() {
    return {
      fusion: { title: 'Fusion', items: [
        { t: 'group', l: 'Lecture des fichiers' },
        { k: 'allSheets', t: 'switch', l: 'Lire toutes les feuilles de chaque fichier', h: 'Sinon, seule la première feuille est lue.', a: 'analyze' },
        { k: 'headerMode', t: 'select', l: 'Ligne des titres de colonnes', a: 'analyze', o: [{ v: 'auto', l: 'Détection automatique' }, { v: 'fixed', l: 'Numéro de ligne précis' }, { v: 'none', l: 'Aucune (fichiers sans titres)' }] },
        { k: 'headerRow', t: 'number', l: 'Numéro de la ligne des titres', min: 1, max: 200, a: 'analyze', when: () => S.headerMode === 'fixed' },
        { t: 'group', l: 'Rapprochement des colonnes' },
        { k: 'sensitivity', t: 'seg', l: 'Sensibilité', h: 'Strict : seulement les noms quasi identiques. Large : rapproche aussi les noms plus éloignés.', a: 'analyze', o: [{ v: 'strict', l: 'Strict' }, { v: 'balanced', l: 'Équilibré' }, { v: 'large', l: 'Large' }] },
        { k: 'contentMatch', t: 'switch', l: 'Reconnaître les colonnes d’après leur contenu', h: 'E-mails, téléphones, sexe, listes de valeurs identiques…', a: 'analyze' },
        { k: 'splitNames', t: 'switch', l: 'Séparer « Nom et Prénoms » en deux colonnes', h: 'Quand d’autres fichiers ont Nom et Prénom séparés.', a: 'analyze' },
        { t: 'group', l: 'Lignes' },
        { k: 'combine', t: 'switch', l: 'Regrouper les lignes de la même personne', h: 'Reconnue par matricule, e-mail, téléphone ou nom. Retire aussi les doublons exacts.', a: 'build' },
        { k: 'keyCol', t: 'select', l: 'Reconnaître une ligne d’après', a: 'build', when: () => S.combine, o: () => [{ v: 'auto', l: 'Automatique (personne)' }].concat(state.cols.filter((c) => c.include && !c.into).map((c) => ({ v: c.key, l: 'Colonne « ' + c.label + ' »' }))) },
        { k: 'conflict', t: 'select', l: 'En cas de valeurs différentes', a: 'build', when: () => S.combine, o: [{ v: 'first', l: 'Garder la première (ordre des fichiers)' }, { v: 'last', l: 'Garder la dernière' }, { v: 'longest', l: 'Garder la plus complète' }] },
      ] },
      data: { title: 'Données', items: [
        { k: 'clean', t: 'switch', l: 'Harmoniser les données', h: 'Applique les règles ci-dessous.', a: 'build' },
        { k: 'nomCase', t: 'select', l: 'Noms de famille', a: 'build', when: () => S.clean, o: [{ v: 'upper', l: 'MAJUSCULES' }, { v: 'title', l: 'Première Lettre Majuscule' }, { v: 'keep', l: 'Ne pas modifier' }] },
        { k: 'prenomCase', t: 'select', l: 'Prénoms', a: 'build', when: () => S.clean, o: [{ v: 'title', l: 'Première Lettre Majuscule' }, { v: 'upper', l: 'MAJUSCULES' }, { v: 'keep', l: 'Ne pas modifier' }] },
        { k: 'emailLower', t: 'switch', l: 'E-mails en minuscules', a: 'build', when: () => S.clean },
        { k: 'sexNorm', t: 'switch', l: 'Sexe : M, H, Homme → Masculin ; F, Femme → Féminin', a: 'build', when: () => S.clean },
        { k: 'placeCase', t: 'switch', l: 'Villes et pays avec majuscule', a: 'build', when: () => S.clean },
        { k: 'phoneFormat', t: 'select', l: 'Numéros de téléphone', a: 'build', when: () => S.clean, o: [{ v: 'keep', l: 'Ne pas modifier' }, { v: 'digits', l: 'Chiffres seulement' }, { v: 'spaced', l: 'Par groupes de 2 chiffres' }, { v: 'intl', l: 'Avec indicatif du pays' }] },
        { k: 'phoneCC', t: 'text', l: 'Indicatif du pays', h: 'Par exemple 228 pour le Togo.', a: 'build', when: () => S.clean && (S.phoneFormat === 'intl' || S.phoneFormat === 'spaced') },
        { t: 'group', l: 'Affichage des valeurs' },
        { k: 'dateFormat', t: 'select', l: 'Format des dates', a: 'view', o: [{ v: 'dmy', l: '31/12/2025' }, { v: 'ymd', l: '2025-12-31' }, { v: 'mdy', l: '12/31/2025' }, { v: 'long', l: '31 décembre 2025' }] },
        { k: 'emptyText', t: 'select', l: 'Cases vides', a: 'view', o: [{ v: '', l: 'Laisser vide' }, { v: '—', l: 'Tiret (—)' }, { v: 'N/A', l: 'N/A' }, { v: '-', l: '-' }] },
        { k: 'dropEmptyCols', t: 'switch', l: 'Masquer les colonnes entièrement vides', a: 'build' },
      ] },
      table: { title: 'Tableau', items: [
        { k: 'numbering', t: 'switch', l: 'Numéroter les lignes (colonne N°)', a: 'build' },
        { k: 'source', t: 'switch', l: 'Ajouter une colonne « Fichier source »', a: 'build' },
        { t: 'group', l: 'Tri' },
        { k: 'sortBy', t: 'select', l: 'Trier par', a: 'build', o: () => colOptions(true, 'Ordre des fichiers') },
        { k: 'sortDir', t: 'seg', l: 'Sens', a: 'build', o: [{ v: 'asc', l: 'A → Z' }, { v: 'desc', l: 'Z → A' }], when: () => S.sortBy !== 'auto' && S.sortBy !== 'none' },
        { k: 'sortBy2', t: 'select', l: 'Puis par', a: 'build', when: () => S.sortBy !== 'auto' && S.sortBy !== 'none', o: () => colOptions(false) },
        { k: 'sortDir2', t: 'seg', l: 'Sens (2e tri)', a: 'build', o: [{ v: 'asc', l: 'A → Z' }, { v: 'desc', l: 'Z → A' }], when: () => S.sortBy !== 'auto' && S.sortBy !== 'none' && S.sortBy2 !== 'none' },
        { t: 'group', l: 'Aperçu' },
        { k: 'previewRows', t: 'select', l: 'Lignes affichées dans l’aperçu', a: 'view', num: true, o: [{ v: 50, l: '50' }, { v: 100, l: '100' }, { v: 500, l: '500' }, { v: 1000, l: '1 000' }, { v: 5000, l: '5 000' }] },
      ] },
      export: { title: 'Export', items: [
        { k: 'format', t: 'select', l: 'Format du fichier', a: 'export', o: Object.keys(FORMATS).map((k) => ({ v: k, l: FORMATS[k].name + ' (.' + FORMATS[k].ext + ')' })) },
        { k: 'fileName', t: 'text', l: 'Nom du fichier', h: 'Variables : {date}, {heure}, {n} (lignes), {fichiers}.', a: 'export' },
        { t: 'group', l: 'Fichier Excel' },
        { k: 'headerColor', t: 'select', l: 'Ligne de titres', a: 'export', o: [{ v: 'dark', l: 'Fond foncé' }, { v: 'accent', l: 'Couleur d’accentuation' }, { v: 'light', l: 'Fond clair' }, { v: 'none', l: 'Sans couleur' }] },
        { k: 'zebra', t: 'switch', l: 'Lignes alternées (zébrage)', a: 'export' },
        { k: 'freeze', t: 'switch', l: 'Figer la ligne de titres', a: 'export' },
        { k: 'autofilter', t: 'switch', l: 'Ajouter les filtres', a: 'export' },
        { k: 'reportSheet', t: 'switch', l: 'Ajouter une feuille « Rapport »', h: 'Liste des fichiers lus et des valeurs différentes.', a: 'export' },
        { t: 'group', l: 'Fichiers texte' },
        { k: 'csvSafe', t: 'switch', l: 'Protéger contre les formules', h: 'Évite qu’une cellule commençant par = + - @ soit exécutée par Excel.', a: 'export' },
      ] },
      look: { title: 'Affichage', items: [
        { k: 'theme', t: 'seg', l: 'Thème', a: 'ui', o: [{ v: 'light', l: 'Clair' }, { v: 'dark', l: 'Sombre' }, { v: 'system', l: 'Système' }] },
        { k: 'accent', t: 'accent', l: 'Couleur d’accentuation', a: 'ui' },
        { k: 'density', t: 'seg', l: 'Densité du tableau', a: 'ui', o: [{ v: 'comfortable', l: 'Confortable' }, { v: 'compact', l: 'Compacte' }] },
        { k: 'fontScale', t: 'seg', l: 'Taille du texte', a: 'ui', num: true, o: [{ v: 90, l: 'Petit' }, { v: 100, l: 'Normal' }, { v: 115, l: 'Grand' }, { v: 130, l: 'Très grand' }] },
        { k: 'motion', t: 'seg', l: 'Animations', a: 'ui', o: [{ v: 'system', l: 'Système' }, { v: 'on', l: 'Activées' }, { v: 'off', l: 'Réduites' }] },
      ] },
      app: { title: 'Application', items: [{ t: 'html', html: appSectionHtml }] },
    };
  }

  const ACCENTS = { blue: '#3355ff', indigo: '#5b3df5', green: '#0a8f5b', orange: '#e8590c', rose: '#d6336c', slate: '#334155' };
  const ACCENT_NAMES = { blue: 'Bleu', indigo: 'Indigo', green: 'Vert', orange: 'Orange', rose: 'Rose', slate: 'Ardoise' };

  function controlHtml(it) {
    if (it.t === 'group') return '<p class="group-title">' + esc(it.l) + '</p>';
    if (it.t === 'html') return it.html();
    if (it.when && !it.when()) return '';
    const id = 's-' + it.k;
    const hint = it.h ? '<p class="setting-hint">' + esc(it.h) + '</p>' : '';
    const v = S[it.k];
    if (it.t === 'switch') {
      return '<div class="setting"><div class="setting-row"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label><input type="checkbox" role="switch" class="switch" id="' + id + '" data-k="' + it.k + '" ' + (v ? 'checked' : '') + '></div>' + hint + '</div>';
    }
    const opts = typeof it.o === 'function' ? it.o() : it.o;
    if (it.t === 'select') {
      const cur = opts.some((o) => String(o.v) === String(v)) ? v : opts[0] && opts[0].v;
      return '<div class="setting"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label><select id="' + id + '" data-k="' + it.k + '"' + (it.num ? ' data-num="1"' : '') + '>' +
        opts.map((o) => '<option value="' + esc(o.v) + '"' + (String(o.v) === String(cur) ? ' selected' : '') + '>' + esc(o.l) + '</option>').join('') + '</select>' + hint + '</div>';
    }
    if (it.t === 'seg') {
      return '<div class="setting"><span class="setting-label" id="' + id + '-l">' + esc(it.l) + '</span><div class="seg" role="radiogroup" aria-labelledby="' + id + '-l">' +
        opts.map((o) => '<button type="button" role="radio" data-k="' + it.k + '" data-v="' + esc(o.v) + '"' + (it.num ? ' data-num="1"' : '') + ' aria-checked="' + (String(o.v) === String(v)) + '">' + esc(o.l) + '</button>').join('') + '</div>' + hint + '</div>';
    }
    if (it.t === 'accent') {
      return '<div class="setting"><span class="setting-label" id="' + id + '-l">' + esc(it.l) + '</span><div class="swatches" role="radiogroup" aria-labelledby="' + id + '-l">' +
        Object.keys(ACCENTS).map((k) => '<button type="button" class="swatch" role="radio" data-k="accent" data-v="' + k + '" aria-label="' + ACCENT_NAMES[k] + '" title="' + ACCENT_NAMES[k] + '" style="background:' + ACCENTS[k] + '" aria-checked="' + (k === v) + '"><svg class="ic"><use href="#i-check"/></svg></button>').join('') + '</div></div>';
    }
    if (it.t === 'text') {
      return '<div class="setting"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label><input type="text" id="' + id + '" data-k="' + it.k + '" value="' + esc(v) + '" autocomplete="off">' + hint + '</div>';
    }
    if (it.t === 'number') {
      return '<div class="setting"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label><input type="number" id="' + id + '" data-k="' + it.k + '" min="' + it.min + '" max="' + it.max + '" value="' + esc(v) + '" inputmode="numeric">' + hint + '</div>';
    }
    return '';
  }

  function renderSettings() {
    const sc = schema();
    $('#settingsTabs').innerHTML = SECTION_ORDER.map((k) => '<button type="button" role="tab" data-sec="' + k + '" aria-selected="' + (k === settingsSection) + '">' + esc(sc[k].title) + '</button>').join('');
    $('#settingsBody').innerHTML = sc[settingsSection].items.map(controlHtml).join('');
  }

  function changedCount() {
    return Object.keys(DEFAULTS).filter((k) => !['theme', 'accent', 'density', 'fontScale', 'motion'].includes(k) && S[k] !== DEFAULTS[k]).length;
  }
  function updateBadge() {
    const n = changedCount();
    $('#settingsBadge').hidden = !n;
  }

  function applyChange(kind) {
    saveSettings();
    updateBadge();
    if (kind === 'analyze') recompute();
    else if (kind === 'build') { if (state.analysis) rebuild(true); }
    else if (kind === 'view') {
      if (state.built) { state.quality = computeQuality(); renderTable(false); renderCheck(); }
    } else if (kind === 'ui') applyUi();
    else if (kind === 'export') { if (state.analysis) renderActions(); }
    if (state.analysis) renderActions();
  }

  function findItem(k) {
    const sc = schema();
    for (const sec of SECTION_ORDER) for (const it of sc[sec].items) if (it.k === k) return it;
    return null;
  }

  function setSetting(k, raw) {
    const it = findItem(k);
    if (!it) return;
    let v = raw;
    if (typeof DEFAULTS[k] === 'number') {
      v = Number(raw);
      if (!Number.isFinite(v)) return;
      if (it.min != null) v = clamp(Math.round(v), it.min, it.max);
    } else if (typeof DEFAULTS[k] === 'boolean') v = !!raw;
    else v = String(raw);
    if (S[k] === v) return;
    S[k] = v;
    applyChange(it.a);
    if (k === 'accent' || k === 'theme') updateSwatchState();
  }

  function updateSwatchState() { /* l'état visuel est recalculé par renderSettings */ }

  function openSettings(section) {
    if (section) settingsSection = section;
    renderSettings();
    const d = $('#settingsDlg');
    if (!d.open) d.showModal();
  }

  /* ================================================================== */
  /*  Application Android / Windows                                      */
  /* ================================================================== */

  const UA = navigator.userAgent || '';
  const PLATFORM = /Android/i.test(UA) ? 'android' : /iPhone|iPad|iPod/i.test(UA) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1) ? 'ios' : /Windows/i.test(UA) ? 'windows' : 'other';
  const isShell = () => /Electron/i.test(UA) || !!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform());
  const isStandalone = () => window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;

  function repo() {
    const r = (window.MAYA_CONFIG && window.MAYA_CONFIG.repo) || '';
    return r && r.indexOf('__') !== 0 ? r : 'obedelom17/mayaexcel';
  }
  const assetUrl = (file) => 'https://github.com/' + repo() + '/releases/latest/download/' + file;
  const releasesUrl = () => 'https://github.com/' + repo() + '/releases/latest';
  const FILES = (window.MAYA_CONFIG && window.MAYA_CONFIG.files) || { android: 'MayaExcel.apk', windows: 'MayaExcel-Setup.exe', portable: 'MayaExcel-Portable.exe' };

  function startDownload(file) {
    const a = document.createElement('a');
    a.href = assetUrl(file);
    a.rel = 'noopener';
    a.target = '_blank'; // le site reste ouvert, même si la version n'est pas encore publiée
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  function dlCard(file, title, sub, icon) {
    return '<a class="dl-card" href="' + esc(assetUrl(file)) + '" target="_blank" rel="noopener"><svg class="ic"><use href="#' + icon + '"/></svg><span><b>' + esc(title) + '</b><small>' + esc(sub) + '</small></span></a>';
  }

  function appSectionHtml() {
    if (isShell()) return '<div class="setting"><p class="setting-hint">Vous utilisez déjà l’application MayaExcel.</p></div>' + resetBlockHtml();
    return '<p class="group-title">Télécharger l’application</p><div style="display:grid;gap:.6rem;margin:.5rem 0 1rem">' +
      dlCard(FILES.android, 'Android (.apk)', 'Fonctionne sans connexion', 'i-phone') +
      dlCard(FILES.windows, 'Windows (installateur .exe)', 'Installation classique', 'i-monitor') +
      dlCard(FILES.portable, 'Windows (version portable .exe)', 'Sans installation', 'i-monitor') +
      '</div><p class="setting-hint">iPhone : ouvrez ce site dans Safari, touchez Partager puis « Sur l’écran d’accueil ». Les liens pointent vers la dernière version publiée ; <a href="' + esc(releasesUrl()) + '" target="_blank" rel="noopener">voir toutes les versions</a>.</p>' +
      resetBlockHtml();
  }

  function resetBlockHtml() {
    return '<p class="group-title">Raccourcis clavier</p><div class="setting"><p class="setting-hint"><span class="kbd">Ctrl</span> + <span class="kbd">O</span> choisir des fichiers · <span class="kbd">Ctrl</span> + <span class="kbd">S</span> télécharger · <span class="kbd">Ctrl</span> + <span class="kbd">,</span> réglages · <span class="kbd">Ctrl</span> + <span class="kbd">V</span> coller un tableau</p></div>' +
      '<p class="group-title">À propos</p><div class="setting"><p class="setting-hint">MayaExcel ' + esc((window.MAYA_CONFIG && window.MAYA_CONFIG.version) || '') + ' · vos fichiers ne quittent jamais votre appareil.</p></div>';
  }

  function openAppDialog(autoDownload) {
    const body = $('#appBody');
    let html = '';
    if (PLATFORM === 'android') {
      if (autoDownload) startDownload(FILES.android);
      html = '<p>' + (autoDownload ? 'Le téléchargement de <b>MayaExcel.apk</b> démarre.' : 'Installez MayaExcel sur votre téléphone.') + '</p>' +
        '<ol class="steps"><li>Ouvrez le fichier <b>MayaExcel.apk</b> dans vos téléchargements.</li><li>Si Android le demande, autorisez l’installation depuis ce navigateur.</li><li>Touchez <b>Installer</b>, puis ouvrez MayaExcel.</li></ol>' +
        '<p class="note">Rien ne se passe ? La version est peut-être en cours de publication. Réessayez dans quelques minutes.</p>' +
        dlCard(FILES.android, 'Télécharger à nouveau', 'MayaExcel.apk', 'i-download');
    } else if (PLATFORM === 'windows') {
      if (autoDownload) startDownload(FILES.windows);
      html = '<p>' + (autoDownload ? 'Le téléchargement de <b>MayaExcel-Setup.exe</b> démarre.' : 'Installez MayaExcel sur votre ordinateur.') + '</p>' +
        '<ol class="steps"><li>Ouvrez le fichier téléchargé.</li><li>Si Windows affiche « Windows a protégé votre PC », cliquez sur <b>Informations complémentaires</b>, puis <b>Exécuter quand même</b>.</li><li>Suivez l’installation, puis ouvrez MayaExcel depuis le menu Démarrer.</li></ol>' +
        dlCard(FILES.portable, 'Version portable', 'Sans installation : un seul fichier à ouvrir', 'i-monitor');
    } else if (PLATFORM === 'ios') {
      html = '<p>Sur iPhone et iPad, MayaExcel s’ajoute à l’écran d’accueil depuis Safari.</p>' +
        '<ol class="steps"><li>Ouvrez cette page dans <b>Safari</b>.</li><li>Touchez le bouton <b>Partager</b> (le carré avec une flèche vers le haut).</li><li>Choisissez <b>Sur l’écran d’accueil</b>, puis <b>Ajouter</b>.</li></ol>';
    } else {
      html = '<p>Choisissez la version pour votre appareil.</p><div style="display:grid;gap:.6rem">' +
        dlCard(FILES.android, 'Android (.apk)', 'Téléphones et tablettes Android', 'i-phone') +
        dlCard(FILES.windows, 'Windows (installateur)', 'Ordinateurs Windows', 'i-monitor') +
        dlCard(FILES.portable, 'Windows (portable)', 'Sans installation', 'i-monitor') + '</div>';
    }
    body.innerHTML = html;
    const d = $('#appDlg');
    if (!d.open) d.showModal();
  }

  const STRIP_KEY = 'maya.install.dismissed';
  function stripDismissed() { try { return localStorage.getItem(STRIP_KEY) === '1'; } catch (e) { return false; } }

  function setupAppPromos() {
    if (isShell() || isStandalone()) return;
    $('#footApp').hidden = false;
    const labels = { android: ['App Android', 'Téléchargez l’application Android de MayaExcel.', 'Télécharger'], windows: ['App Windows', 'Téléchargez l’application Windows de MayaExcel.', 'Télécharger'], ios: ['Écran d’accueil', 'Ajoutez MayaExcel à l’écran d’accueil de votre iPhone.', 'Voir comment'] };
    const L = labels[PLATFORM];
    if (!L) return;
    $('#getAppBtn').hidden = false;
    $('#getAppLabel').textContent = L[0];
    $('#installText').textContent = L[1];
    $('#installStripBtn').textContent = L[2];
    if (!stripDismissed()) $('#installStrip').hidden = false;
  }

  /* ================================================================== */
  /*  Événements                                                         */
  /* ================================================================== */

  const fileInput = $('#fileInput');
  const drop = $('#drop');
  const openPicker = () => fileInput.click();

  $('#pickBtn').addEventListener('click', (e) => { e.stopPropagation(); openPicker(); });
  $('#pasteBtn').addEventListener('click', (e) => { e.stopPropagation(); pasteFromButton(); });
  $('#demoBtn').addEventListener('click', (e) => { e.stopPropagation(); loadDemo(); });
  drop.addEventListener('click', openPicker);
  drop.addEventListener('keydown', (e) => {
    if (e.target === drop && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); openPicker(); }
  });
  fileInput.addEventListener('change', () => {
    const list = Array.from(fileInput.files);
    fileInput.value = '';
    addFiles(list);
  });

  function swapTitle(text) {
    const t = $('#dropTitle');
    t.textContent = text;
    t.classList.remove('swap');
    void t.offsetWidth;
    t.classList.add('swap');
  }
  ['dragenter', 'dragover'].forEach((ev) =>
    window.addEventListener(ev, (e) => {
      if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return;
      e.preventDefault();
      if (!drop.classList.contains('is-over')) swapTitle('Relâchez pour fusionner');
      drop.classList.add('is-over');
    })
  );
  window.addEventListener('dragleave', (e) => {
    if (e.relatedTarget) return;
    if (drop.classList.contains('is-over')) swapTitle('Déposez vos fichiers Excel ici');
    drop.classList.remove('is-over');
  });
  window.addEventListener('drop', (e) => {
    if (drop.classList.contains('is-over')) swapTitle('Déposez vos fichiers Excel ici');
    drop.classList.remove('is-over');
    if (!e.dataTransfer || !e.dataTransfer.files.length) return;
    e.preventDefault();
    addFiles(e.dataTransfer.files);
  });
  document.addEventListener('paste', (e) => {
    const t = e.target;
    if (t && (/^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName) || t.isContentEditable)) return;
    const text = (e.clipboardData || window.clipboardData || { getData: () => '' }).getData('text');
    if (text && addPastedText(text)) e.preventDefault();
  });

  $('#clearBtn').addEventListener('click', () => {
    state.files = [];
    state.overrides = {};
    state.separate = [];
    state.colOrder = [];
    state.query = '';
    recompute();
    renderFiles();
    window.scrollTo({ top: 0, behavior: reduceMotion() ? 'auto' : 'smooth' });
  });

  $('#filesList').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act]');
    if (!btn || btn.tagName === 'INPUT') return;
    const id = Number(btn.dataset.id);
    const i = state.files.findIndex((f) => f.id === id);
    if (i < 0) return;
    if (btn.dataset.act === 'remove') {
      const li = btn.closest('.file');
      const done = () => { state.files = state.files.filter((f) => f.id !== id); recompute(); renderFiles(); };
      if (reduceMotion() || !li) return done();
      li.classList.add('leaving');
      setTimeout(done, 230);
    } else if (btn.dataset.act === 'up' || btn.dataset.act === 'down') {
      const j = btn.dataset.act === 'up' ? i - 1 : i + 1;
      if (j < 0 || j >= state.files.length) return;
      [state.files[i], state.files[j]] = [state.files[j], state.files[i]];
      renderFiles();
      recompute();
    }
  });
  $('#filesList').addEventListener('change', (e) => {
    if (e.target.dataset.act !== 'toggle') return;
    const f = state.files.find((x) => x.id === Number(e.target.dataset.id));
    if (!f) return;
    f.include = e.target.checked;
    recompute();
  });

  /* Barre d'outils */
  let searchTimer;
  $('#searchInput').addEventListener('input', (e) => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.query = e.target.value;
      if (state.built) renderTable(false);
    }, 150);
  });
  $('#sortBy').addEventListener('change', (e) => {
    S.sortBy = e.target.value;
    saveSettings();
    $('#sortDir').disabled = S.sortBy === 'none';
    rebuild(false);
  });
  $('#sortDir').addEventListener('click', () => {
    S.sortDir = S.sortDir === 'asc' ? 'desc' : 'asc';
    saveSettings();
    $('#sortDir').textContent = S.sortDir === 'asc' ? 'A → Z' : 'Z → A';
    rebuild(false);
  });
  $('#toolbarSettings').addEventListener('click', () => openSettings('fusion'));

  /* Onglets */
  $('.tabs').addEventListener('click', (e) => {
    const t = e.target.closest('.tab');
    if (t) setTab(t.dataset.tab);
  });
  $('.tabs').addEventListener('keydown', (e) => {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    const tabs = $$('.tab');
    const i = tabs.findIndex((t) => t.getAttribute('aria-selected') === 'true');
    const next = tabs[(i + (e.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length];
    setTab(next.dataset.tab);
    next.focus();
  });

  /* Tableau */
  $('#sheet').addEventListener('click', (e) => {
    const td = e.target.closest('td[data-r]');
    if (td) selectCell(td);
  });
  $('#panelCheck').addEventListener('click', (e) => {
    const b = e.target.closest('[data-goto]');
    if (!b) return;
    const [r, c] = b.dataset.goto.split(':').map(Number);
    gotoCell(r, c);
  });

  /* Colonnes */
  function rememberCol(c) {
    const prev = state.overrides[c.key] || {};
    state.overrides[c.key] = { label: c.label, include: c.include, into: prev.into };
  }
  const colsList = $('#colsList');
  colsList.addEventListener('input', (e) => {
    if (e.target.dataset.act !== 'label') return;
    const c = state.cols.find((x) => x.id === e.target.closest('.col-row').dataset.id);
    c.label = e.target.value;
    rememberCol(c);
    rebuild(false);
  });
  colsList.addEventListener('change', (e) => {
    const act = e.target.dataset.act;
    const row = e.target.closest('.col-row');
    if (!row) return;
    const c = state.cols.find((x) => x.id === row.dataset.id);
    if (act === 'include') {
      c.include = e.target.checked;
      rememberCol(c);
      row.classList.toggle('is-off', !c.include);
      renderToolbar();
      rebuild(false);
    } else if (act === 'label') renderToolbar();
    else if (act === 'into') {
      const target = state.cols.find((x) => x.id === e.target.value);
      const prev = state.overrides[c.key] || { label: c.label, include: c.include };
      state.overrides[c.key] = Object.assign({}, prev, { into: target ? target.key : null });
      c.into = target ? target.id : null;
      renderColumns();
      renderToolbar();
      rebuild(false);
    }
  });
  colsList.addEventListener('click', (e) => {
    const cut = e.target.closest('[data-act="detach"]');
    if (cut) {
      if (!state.separate.includes(cut.dataset.vk)) state.separate.push(cut.dataset.vk);
      toast('Colonne séparée.');
      recompute();
      return;
    }
    const btn = e.target.closest('[data-act="up"],[data-act="down"]');
    if (!btn) return;
    const id = btn.closest('.col-row').dataset.id;
    const i = state.cols.findIndex((x) => x.id === id);
    const j = btn.dataset.act === 'up' ? i - 1 : i + 1;
    if (j < 0 || j >= state.cols.length) return;
    [state.cols[i], state.cols[j]] = [state.cols[j], state.cols[i]];
    state.colOrder = state.cols.map((c) => c.key);
    renderColumns();
    rebuild(false);
    const again = colsList.querySelector('[data-id="' + id + '"] [data-act="' + btn.dataset.act + '"]');
    if (again && !again.disabled) again.focus();
  });

  /* Actions */
  $('#dlBtn').addEventListener('click', doDownload);
  $('#copyBtn').addEventListener('click', doCopy);
  $('#shareBtn').addEventListener('click', doShare);
  $('#formatSel').addEventListener('change', (e) => {
    S.format = e.target.value;
    saveSettings();
    $('#dlLabel').textContent = 'Télécharger (.' + FORMATS[S.format].ext + ')';
  });

  /* Thème */
  $('#themeBtn').addEventListener('click', () => {
    S.theme = isDark() ? 'light' : 'dark';
    saveSettings();
    applyUi();
    if ($('#settingsDlg').open) renderSettings();
  });

  /* Panneau de réglages */
  $('#settingsBtn').addEventListener('click', () => openSettings());
  $('#settingsTabs').addEventListener('click', (e) => {
    const b = e.target.closest('[data-sec]');
    if (!b) return;
    settingsSection = b.dataset.sec;
    renderSettings();
    $('#settingsBody').scrollTop = 0;
  });
  const settingsBody = $('#settingsBody');
  settingsBody.addEventListener('change', (e) => {
    const el = e.target;
    if (!el.dataset.k) return;
    if (el.type === 'checkbox') setSetting(el.dataset.k, el.checked);
    else setSetting(el.dataset.k, el.value);
    renderSettings();
  });
  settingsBody.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-k]');
    if (!b) return;
    setSetting(b.dataset.k, b.dataset.v);
    renderSettings();
  });
  $('#resetSettings').addEventListener('click', () => {
    S = Object.assign({}, DEFAULTS);
    saveSettings();
    applyUi();
    updateBadge();
    renderSettings();
    if (state.analysis) { recompute(); renderActions(); }
    toast('Réglages réinitialisés.');
  });

  /* Fenêtres : fermeture par bouton ou clic à l'extérieur */
  $$('dialog').forEach((d) => {
    d.addEventListener('click', (e) => { if (e.target === d || e.target.closest('[data-close]')) d.close(); });
  });

  /* Application */
  $('#getAppBtn').addEventListener('click', () => openAppDialog(true));
  $('#installStripBtn').addEventListener('click', () => openAppDialog(true));
  $('#footApp').addEventListener('click', () => openAppDialog(false));
  $('#installDismiss').addEventListener('click', () => {
    $('#installStrip').hidden = true;
    try { localStorage.setItem(STRIP_KEY, '1'); } catch (e) { /* ignoré */ }
  });
  window.addEventListener('beforeinstallprompt', (e) => e.preventDefault()); // on propose l'APK / l'EXE, pas l'installation par le navigateur

  /* En-tête qui se détache au défilement */
  const topBar = $('#top');
  const onScroll = () => topBar.classList.toggle('is-stuck', window.scrollY > 6);
  window.addEventListener('scroll', onScroll, { passive: true });

  /* Raccourcis clavier */
  document.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    if (e.key === 'o' || e.key === 'O') { e.preventDefault(); openPicker(); }
    else if (e.key === 's' || e.key === 'S') { e.preventDefault(); doDownload(); }
    else if (e.key === ',') { e.preventDefault(); openSettings(); }
  });

  /* ================================================================== */
  /*  Démarrage                                                          */
  /* ================================================================== */

  applyUi();
  updateBadge();
  onScroll();
  setupAppPromos();
  renderFiles();

  if ('serviceWorker' in navigator && !isShell() && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    window.addEventListener('load', () => navigator.serviceWorker.register('/sw.js').catch(() => {}));
  }
  (window.requestIdleCallback || ((f) => setTimeout(f, 1200)))(() => loadLib('xlsx').catch(() => {}));

  window.MayaExcelApp = { addParsedFiles };
})();
