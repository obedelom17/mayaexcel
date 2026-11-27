/* MayaExcel – interface */
(function () {
  'use strict';

  const $ = (sel, el) => (el || document).querySelector(sel);
  const $$ = (sel, el) => Array.from((el || document).querySelectorAll(sel));
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const plural = (n, one, many) => n + ' ' + (n > 1 ? many : one);
  const clamp = (n, a, b) => Math.min(b, Math.max(a, n));
  const tick = () => new Promise((r) => setTimeout(r, 0));
  const isMobile = () => window.matchMedia('(max-width: 899px)').matches;

  /* ================================================================== */
  /*  Réglages (mémorisés sur l'appareil)                                */
  /* ================================================================== */

  const KEY = 'maya.settings.v4';
  const DEFAULTS = {
    // affichage
    density: 'comfortable', fontScale: 100, motion: 'system',
    // lecture et fusion
    allSheets: true, formDetect: true, multiHeader: true, nullTokens: true, headerMode: 'auto', headerRow: 1,
    sensitivity: 'balanced', contentMatch: true, splitNames: true,
    combine: true, personMatch: 'normal', fuzzyNames: true, useBirth: true, keyCol: 'auto', conflict: 'first',
    ignoreColumns: '', synonyms: '',
    // données
    clean: true, nomCase: 'upper', prenomCase: 'title', emailLower: true, sexNorm: true, normCats: true, placeCase: true,
    unifyLabels: true, phoneFormat: 'keep', phoneCC: '228', dateFormat: 'dmy', emptyText: '', dropEmptyCols: true,
    // tableau
    numbering: true, source: false, sortBy: 'auto', sortDir: 'asc', sortBy2: 'none', sortDir2: 'asc', previewRows: 500,
    // export
    format: 'xlsx', fileName: 'MayaExcel-fusion-{date}', zebra: false, freeze: true, autofilter: true,
    headerColor: 'light', reportSheet: true, csvSafe: true,
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

  const motionQuery = window.matchMedia('(prefers-reduced-motion: reduce)');
  const reduceMotion = () => S.motion === 'off' || (S.motion === 'system' && motionQuery.matches);

  function applyUi() {
    const r = document.documentElement;
    r.dataset.density = S.density;
    r.dataset.motion = S.motion;
    r.style.setProperty('--fs', S.fontScale / 100);
  }

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
    files: [],        // {id,name,size,status:'reading'|'ok'|'error',error,sheets,include}
    cols: [],         // [{id,key,label,include,into}]
    overrides: {},    // key → {label, include, into}
    separate: [],     // intitulés détachés d'un regroupement
    colOrder: [],     // ordre choisi (clés)
    colW: {},         // largeurs de colonnes choisies (par intitulé)
    noMerge: new Set(), // enregistrements dissociés à la main
    analysis: null,
    built: null,
    quality: null,
    tab: 'table',
    query: '',
    extraRows: 0,
    exported: false,
  };
  let uid = 0;
  const MAX_DOM_ROWS = 5000;

  /* ================================================================== */
  /*  Utilitaires d'affichage                                            */
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
    chain = chain.then(task).catch((e) => { console.error(e); renderAll(); });
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

    if (!state.analysis && isMobile()) setView('sources');
    renderAll();
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
    recompute();
    afterResult();
  }

  function afterResult() {
    if (state.analysis && state.built && state.built.rows.length && isMobile()) setView('main');
  }

  /** Tableau copié depuis Excel (texte avec tabulations) → source virtuelle. */
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
      recompute();
      afterResult();
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

  /** Point d'entrée pour injecter des fichiers déjà lus (exemple, tests). */
  function addParsedFiles(list) {
    list.forEach((f) => state.files.push({ id: ++uid, name: f.name, size: f.size || 4096, status: 'ok', error: null, include: true, sheets: f.sheets }));
    recompute();
    afterResult();
  }

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
      { name: 'Fiche-Yawo.xlsx', sheets: [{ name: 'Fiche', rows: [
        ['FICHE INDIVIDUELLE'], [],
        ['Nom :', 'TOGBE'], ['Prénoms :', 'Yawo'], ['Tél.', '99887766'], ['Fonction :', 'Chauffeur'], ['E-mail', 'yawo.togbe@exemple.tg'], ['Sexe', 'M'],
      ] }] },
    ];
  }

  function loadDemo() {
    enqueue(async () => {
      state.files = [];
      state.overrides = {};
      state.separate = [];
      state.colOrder = [];
      state.noMerge = new Set();
      demoFiles().forEach((f) => state.files.push({ id: ++uid, name: f.name, size: 4096, status: 'ok', error: null, include: true, sheets: f.sheets }));
      recompute();
      afterResult();
    });
  }

  /* ================================================================== */
  /*  Fusion                                                             */
  /* ================================================================== */

  function recompute() {
    const ok = state.files.filter((f) => f.status === 'ok' && f.include !== false);
    if (!ok.length) {
      state.analysis = null;
      state.built = null;
      state.quality = null;
      state.cols = [];
      renderAll();
      return;
    }
    state.analysis = MayaMerge.analyze(
      ok.map((f) => ({ id: f.id, name: f.name, sheets: f.sheets })),
      {
        allSheets: S.allSheets, splitNames: S.splitNames, separate: state.separate, sensitivity: S.sensitivity,
        contentMatch: S.contentMatch, headerMode: S.headerMode, headerRow: clamp(Number(S.headerRow) || 1, 1, 200),
        formDetect: S.formDetect, multiHeader: S.multiHeader, nullTokens: S.nullTokens,
        ignoreColumns: S.ignoreColumns, synonyms: S.synonyms,
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
    rebuild(true);
  }

  function colIdByKey(key) {
    const c = state.cols.find((x) => x.key === key);
    return c ? c.id : null;
  }

  function buildSettings() {
    const sortBy = S.sortBy === 'auto' || S.sortBy === 'none' ? S.sortBy : colIdByKey(S.sortBy) || 'auto';
    return {
      cols: state.cols, source: S.source, numbering: S.numbering, combine: S.combine, personMatch: S.personMatch,
      fuzzyNames: S.fuzzyNames, useBirth: S.useBirth, noMerge: Array.from(state.noMerge),
      keyCol: S.keyCol === 'auto' ? 'auto' : colIdByKey(S.keyCol) || 'auto', conflict: S.conflict,
      clean: S.clean, nomCase: S.nomCase, prenomCase: S.prenomCase, emailLower: S.emailLower, sexNorm: S.sexNorm,
      normCats: S.normCats, placeCase: S.placeCase, unifyLabels: S.unifyLabels, phoneFormat: S.phoneFormat, phoneCC: S.phoneCC,
      dropEmptyCols: S.dropEmptyCols, fmtDate: (d) => fmtDate(d, S.dateFormat),
      sortBy, sortDir: S.sortDir, sortBy2: S.sortBy2 === 'none' ? 'none' : colIdByKey(S.sortBy2) || 'none', sortDir2: S.sortDir2,
    };
  }

  function rebuild(full) {
    if (!state.analysis) return;
    state.built = MayaMerge.build(state.analysis, buildSettings());
    state.quality = computeQuality();
    state.exported = false;
    renderStepper();
    if (full) renderAll();
    else {
      renderSummary();
      renderTable();
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
    dup.forEach((rows) => {
      if (rows.length > 1) rows.forEach((i) => add(i, matCol, 'Matricule en double', b.rows[i][matCol]));
    });
    Object.keys(b.conflictMap).forEach((key) => {
      const [r, c] = key.split(':').map(Number);
      issues.push({ r, c, t: 'Valeurs différentes selon les sources', v: 'autre valeur : ' + b.conflictMap[key] });
    });
    issues.sort((a, b2) => a.r - b2.r || a.c - b2.c);
    return { issues, comp, bad };
  }

  /* ================================================================== */
  /*  Rendu                                                              */
  /* ================================================================== */

  function renderAll() {
    renderLayout();
    renderFiles();
    renderWork();
    renderActions();
    renderStepper();
    renderStatus();
    renderHome();
    updateBadge();
  }

  function renderStepper() {
    const step = state.exported ? 3 : state.analysis ? 2 : 1;
    $$('#stepper li').forEach((li) => {
      const n = Number(li.dataset.step);
      li.classList.toggle('done', n < step);
      if (n === step) li.setAttribute('aria-current', 'step');
      else li.removeAttribute('aria-current');
    });
  }

  function renderStatus() {
    const b = state.built;
    $('#stRows').textContent = b ? plural(b.rows.length, 'ligne', 'lignes') : 'Aucune source';
    $('#stCols').textContent = b ? plural(b.columnCount, 'colonne', 'colonnes') : '';
    $('#stSort').textContent = b ? (b.sortedBy.length ? 'Tri : ' + b.sortedBy.join(', ') : 'Ordre des sources') : '';
    $('#stVer').textContent = 'v' + ((window.MAYA_CONFIG && window.MAYA_CONFIG.version) || '');
  }

  function renderLayout() {
    const hasFiles = state.files.length > 0;
    $('#appbody').dataset.state = hasFiles ? 'ready' : 'empty';
    $('#empty').hidden = hasFiles;
    $('#side').hidden = !hasFiles;
    $('#work').hidden = !hasFiles;
    $('#mobilenav').hidden = !hasFiles;
    $$('#mobilenav [data-nav]').forEach((b) => {
      const nav = b.dataset.nav;
      const current = nav === 'sources' ? document.body.dataset.view === 'sources' : document.body.dataset.view === 'main' && state.tab === nav;
      if (current) b.setAttribute('aria-current', 'page');
      else b.removeAttribute('aria-current');
    });
    if (!hasFiles) document.body.dataset.view = 'main';
  }

  function setView(v) {
    document.body.dataset.view = v;
    renderLayout();
  }

  function renderFiles() {
    if (!state.files.length) return;
    $('#filesCount').textContent = '(' + state.files.length + ')';
    $('#filesHint').hidden = state.files.length < 2;

    const reports = new Map();
    if (state.analysis) state.analysis.report.forEach((r) => reports.set(r.id, r));
    const last = state.files.length - 1;

    $('#filesList').innerHTML = state.files
      .map((f, i) => {
        const ext = /^Tableau collé/.test(f.name) ? 'txt' : (f.name.split('.').pop() || '').toLowerCase().slice(0, 4);
        let meta;
        if (f.status === 'reading') meta = '<span class="spin"></span>Lecture en cours';
        else if (f.status === 'error') meta = esc(f.error);
        else if (f.include === false) meta = 'Exclue de la fusion';
        else {
          const r = reports.get(f.id);
          if (r) {
            const used = r.sheets.filter((s) => !s.empty).length;
            meta = r.rows ? plural(r.rows, 'ligne', 'lignes') + ' · ' + plural(used, 'feuille', 'feuilles') : 'Aucune donnée trouvée';
            if (r.sheets.some((s) => s.headerless)) meta += ' · sans en-têtes';
            if (r.sheets.some((s) => s.form)) meta += ' · fiche';
          } else meta = formatSize(f.size);
        }
        const check = f.status === 'ok'
          ? '<input type="checkbox" data-act="toggle" data-id="' + f.id + '" ' + (f.include !== false ? 'checked' : '') + ' aria-label="Inclure ' + esc(f.name) + ' dans la fusion">'
          : '<span style="width:16px"></span>';
        return (
          '<li class="file' + (f.status === 'error' ? ' is-error' : '') + (f.include === false ? ' is-off' : '') + '" data-id="' + f.id + '">' +
          check + '<span class="file-ext" aria-hidden="true">' + esc(ext) + '</span>' +
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

  function renderWork() {
    const has = !!(state.analysis && state.built);
    $('.panels').classList.toggle('no-data', !has);
    $('#workEmpty').hidden = has;
    ['#panelTable', '#panelCols', '#panelCheck'].forEach((s) => { if (!has) $(s).hidden = true; });
    $('#toolbar').hidden = !has;
    $('.tabs').hidden = !has;
    $('.summarybar').hidden = !has;
    if (!has) return;
    renderToolbar();
    renderSummary();
    renderTable();
    renderColumns();
    renderCheck();
    renderCheckBadge();
    setTab(state.tab, true);
  }

  const num = (n) => '<span>' + n + '</span>';

  function kpi(label, value, sub, tone) {
    return '<div class="kpi' + (tone ? ' ' + tone : '') + '"><span class="kpi-label">' + esc(label) + '</span><span class="kpi-value">' + value + '</span><span class="kpi-sub">' + esc(sub) + '</span></div>';
  }

  function renderSummary() {
    const a = state.analysis;
    const b = state.built;
    if (!a || !b) return;
    const read = a.report.reduce((n, r) => n + r.rows, 0);
    const issues = state.quality ? state.quality.issues.length : 0;
    $('#kpis').innerHTML =
      kpi('Sources', a.fileCount, plural(read, 'ligne lue', 'lignes lues')) +
      kpi('Lignes', b.rows.length, b.absorbed ? plural(b.absorbed, 'ligne rapprochée', 'lignes rapprochées') : 'aucun regroupement') +
      kpi('Colonnes', b.columnCount, plural(a.columns.length, 'détectée', 'détectées')) +
      kpi('Personnes regroupées', b.combinedRows, b.combinedRows ? 'présentes dans plusieurs sources' : 'aucune') +
      kpi('À vérifier', issues, plural(b.conflictCount, 'valeur différente', 'valeurs différentes'), issues ? 'warn' : '');

    const tags = [];
    if (a.mergedColumns) tags.push(['', plural(a.mergedColumns, 'colonne regroupée', 'colonnes regroupées')]);
    if (a.contentColumns) tags.push(['', plural(a.contentColumns, 'colonne reconnue par son contenu', 'colonnes reconnues par leur contenu')]);
    if (a.headerlessSheets) tags.push(['', plural(a.headerlessSheets, 'feuille sans en-têtes', 'feuilles sans en-têtes')]);
    const forms = a.report.reduce((n, r) => n + r.sheets.filter((x) => x.form).length, 0);
    if (forms) tags.push(['', plural(forms, 'fiche individuelle', 'fiches individuelles')]);
    if (a.splitCount && S.splitNames) tags.push(['', plural(a.splitCount, 'nom complet séparé', 'noms complets séparés')]);
    if (b.emptyDropped) tags.push(['', plural(b.emptyDropped, 'colonne vide masquée', 'colonnes vides masquées')]);
    $('#insights').innerHTML = tags.map(([c, t]) => '<li class="' + c + '">' + esc(t) + '</li>').join('');
    renderStatus();
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
    let html = '<option value="auto">Automatique' + (autoSortLabel() ? ' (' + esc(autoSortLabel()) + ')' : '') + '</option><option value="none">Ordre des sources</option>';
    state.cols.filter((c) => c.include && !c.into).forEach((c) => { html += '<option value="' + esc(c.key) + '">' + esc(c.label || 'Colonne') + '</option>'; });
    sel.innerHTML = html;
    sel.value = [...sel.options].some((o) => o.value === S.sortBy) ? S.sortBy : 'auto';
    const dir = $('#sortDir');
    dir.textContent = S.sortDir === 'asc' ? 'A → Z' : 'Z → A';
    dir.disabled = sel.value === 'none';
    $('#searchInput').value = state.query;
  }

  /* ---------- Tableau ---------- */
  function visibleRows() {
    const b = state.built;
    const q = state.query.trim().toLowerCase();
    const idx = [];
    for (let i = 0; i < b.rows.length; i++) {
      if (!q || b.rows[i].some((v) => cellText(v).toLowerCase().indexOf(q) !== -1)) idx.push(i);
    }
    return idx;
  }

  function colWidth(h, j, b) {
    if (state.colW[h]) return state.colW[h];
    if (b.headMeta[j].group === 'index') return 56;
    let len = String(h).length + 5;
    const sample = Math.min(b.rows.length, 40);
    for (let i = 0; i < sample; i++) {
      const t = cellText(b.rows[i][j]);
      if (t.length > len) len = t.length;
    }
    return clamp(Math.round(len * 8.2 + 24), 80, 320);
  }

  const RN_W = 48;

  function renderTable() {
    const b = state.built;
    const box = $('#sheet');
    const note = $('#sheetNote');
    if (!b || !b.rows.length) {
      box.innerHTML = '<p class="empty-note-table">Aucune ligne à afficher. Vérifiez que des colonnes sont affichées.</p>';
      note.textContent = '';
      return;
    }
    const q = state.query.trim().toLowerCase();
    const matches = visibleRows();
    const limit = clamp(Math.max(Number(S.previewRows) || 500, state.extraRows), 10, MAX_DOM_ROWS);
    const shown = matches.slice(0, limit);
    const bad = state.quality ? state.quality.bad : new Set();

    if (!shown.length) {
      box.innerHTML = '<p class="empty-note-table">Aucun résultat pour « ' + esc(state.query) + ' ».</p>';
      note.textContent = '';
      return;
    }

    const widths = b.headers.map((h, j) => colWidth(h, j, b));
    const total = RN_W + widths.reduce((a, c) => a + c, 0);
    const sortKey = S.sortBy;

    let html = '<table class="grid" style="width:' + total + 'px"><colgroup><col style="width:' + RN_W + 'px">' + widths.map((w) => '<col style="width:' + w + 'px">').join('') + '</colgroup><thead><tr><th class="rn"></th>';
    b.headers.forEach((h, j) => {
      const m = b.headMeta[j];
      const col = m.id ? state.cols.find((c) => c.id === m.id) : null;
      const sorted = col && sortKey === col.key;
      const ind = sorted ? '<span class="sort-ind" aria-hidden="true">' + (S.sortDir === 'asc' ? '▲' : '▼') + '</span>' : '';
      const link = m.merged ? '<svg class="ic th-link" aria-label="Colonne regroupée"><title>Regroupe : ' + esc(m.variants.join(', ')) + '</title><use href="#i-link"/></svg>' : '';
      const label = col
        ? '<button type="button" class="th-btn" data-sort="' + esc(col.key) + '" title="Trier par « ' + esc(h) + ' »">' + esc(h) + '</button>'
        : '<span class="th-btn" style="cursor:default">' + esc(h) + '</span>';
      html += '<th scope="col" aria-sort="' + (sorted ? (S.sortDir === 'asc' ? 'ascending' : 'descending') : 'none') + '"><div class="th-in">' + label + ind + link + '</div><span class="rs" data-j="' + j + '" role="separator" aria-label="Redimensionner la colonne"></span></th>';
    });
    html += '</tr></thead><tbody>';
    shown.forEach((ri) => {
      html += '<tr><td class="rn">' + (ri + 1) + '</td>';
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
    $('#fbRef').textContent = '—';
    $('#fbVal').textContent = 'Sélectionnez une cellule';
    $('#fbVal').classList.remove('has');

    let text;
    if (q) text = plural(matches.length, 'ligne correspond', 'lignes correspondent') + ' à la recherche' + (matches.length > limit ? ' (' + limit + ' affichées)' : '') + '.';
    else if (b.rows.length > limit) text = limit + ' premières lignes sur ' + b.rows.length + ' affichées. L’export contient tout le tableau (taille de l’aperçu dans les réglages).';
    else text = '';
    note.textContent = text;
    note.hidden = !text;
  }

  function selectCell(td) {
    const prev = $('#sheet td.sel');
    if (prev) prev.classList.remove('sel');
    td.classList.add('sel');
    $('#fbRef').textContent = colLetter(Number(td.dataset.c)) + (Number(td.dataset.r) + 1);
    const fv = $('#fbVal');
    const alt = td.dataset.alt;
    fv.textContent = (td.textContent || '(vide)') + (alt ? '   ·   autre valeur : ' + alt : '');
    fv.classList.add('has');
  }

  function moveSel(dr, dc) {
    let td = $('#sheet td.sel');
    if (!td) {
      td = $('#sheet tbody td[data-r]');
      if (td) selectCell(td);
      return;
    }
    let tr = td.parentElement;
    let idx = td.cellIndex;
    if (dr) {
      tr = dr > 0 ? tr.nextElementSibling : tr.previousElementSibling;
      if (!tr) return;
    } else idx = clamp(idx + dc, 1, tr.cells.length - 1);
    const next = tr.cells[idx];
    if (next && next.dataset.r !== undefined) {
      selectCell(next);
      next.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  }

  function gotoCell(r, c) {
    setTab('table');
    if (isMobile()) setView('main');
    state.query = '';
    $('#searchInput').value = '';
    state.extraRows = Math.max(state.extraRows, r + 1);
    renderTable();
    const td = $('#sheet td[data-r="' + r + '"][data-c="' + c + '"]');
    if (td) {
      selectCell(td);
      td.scrollIntoView({ block: 'center', inline: 'center' });
    } else toast('Cette ligne est au-delà de l’aperçu (' + MAX_DOM_ROWS + ' lignes maximum).');
  }

  /* ---------- Onglet Colonnes ---------- */
  const REASON_TXT = { contenu: 'reconnue par son contenu', proche: 'intitulé proche', 'déduit': 'déduite du nom complet' };

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
            const cut = m.variants.length > 1 && k ? '<button type="button" data-act="detach" data-vk="' + esc(k) + '" title="Séparer « ' + esc(v) + ' » de cette colonne" aria-label="Séparer ' + esc(v) + '">Séparer</button>' : '';
            return '<span class="chip' + (cut ? '' : ' solo') + '">' + esc(v) + cut + '</span>';
          })
          .join('');
        const badges = m.reasons.filter((r) => REASON_TXT[r]).map((r) => '<span class="badge">' + esc(REASON_TXT[r]) + '</span>').join('');
        const others = state.cols.filter((x) => x.id !== c.id && !x.into);
        const select = '<div class="into"><label for="into-' + c.id + '">Rattacher à</label><select id="into-' + c.id + '" data-act="into"><option value="">Aucune (colonne séparée)</option>' +
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
          '<span>Présente dans ' + m.fileCount + ' source' + (m.fileCount > 1 ? 's' : '') + ' sur ' + total + '</span>' + select + '</div>' +
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
    const b = state.built;
    const panel = $('#panelCheck');
    if (!q || !a || !b) { panel.innerHTML = ''; return; }

    const bars = q.comp
      .map((c) => {
        const pct = Math.round(c.pct * 100);
        return '<div class="bar-row"><span title="' + esc(c.label) + '">' + esc(c.label) + '</span><div class="bar"><i class="' + (pct < 60 ? 'low' : '') + '" style="width:' + pct + '%"></i></div><span class="pct">' + pct + ' %</span></div>';
      })
      .join('');

    const shown = q.issues.slice(0, 200);
    const issues = q.issues.length
      ? '<ul class="list">' + shown.map((it) => '<li><button type="button" class="list-btn" data-goto="' + it.r + ':' + it.c + '"><svg class="ic"><use href="#i-alert"/></svg><span><b>' + esc(it.t) + '</b><small>' + esc(b.headers[it.c] || '') + ' · ligne ' + (it.r + 1) + ' · ' + esc(String(it.v)) + '</small></span></button></li>').join('') + '</ul>' +
        (q.issues.length > shown.length ? '<p class="muted-note">' + (q.issues.length - shown.length) + ' autres anomalies non listées.</p>' : '')
      : '<p class="ok-line"><svg class="ic"><use href="#i-check"/></svg>Aucune anomalie détectée.</p>';

    const merges = b.merges.length
      ? '<ul class="list">' + b.merges.slice(0, 300).map((m) => '<li class="list-row"><div><b>' + esc(m.label) + '</b><small>' + esc(m.sources.join(' · ')) + ' — reconnue par ' + esc(m.by.join(', ') || 'ligne identique') + '</small></div><button type="button" class="btn btn-sm" data-split="' + m.idx.slice(1).join(',') + '">Dissocier</button></li>').join('') + '</ul>'
      : '<p class="muted-note">Aucune personne regroupée.</p>';
    const manual = state.noMerge.size ? '<p class="muted-note">' + plural(state.noMerge.size, 'ligne dissociée à la main', 'lignes dissociées à la main') + '. <button type="button" class="btn btn-sm" data-act="resetsplit">Rétablir</button></p>' : '';

    const rows = a.report
      .map((r) => {
        const used = r.sheets.filter((s) => !s.empty);
        const detail = r.error ? '<span class="tag-s err">erreur</span>' : used.map((s) => esc(s.name) + ' ' + (s.form ? '<span class="tag-s">fiche</span>' : s.headerless ? '<span class="tag-s">sans en-têtes</span>' : '<span class="tag-s">ligne ' + s.headerRow + '</span>')).join(', ') || '—';
        return '<tr><td>' + esc(r.name) + '</td><td>' + detail + '</td><td>' + r.rows + '</td></tr>';
      })
      .join('');

    panel.innerHTML =
      '<div class="check-grid">' +
      '<section class="card"><h3>Anomalies (' + q.issues.length + ')</h3>' + issues + '</section>' +
      '<section class="card"><h3>Personnes regroupées (' + b.merges.length + ')</h3>' + merges + manual + '</section>' +
      '<section class="card"><h3>Remplissage des colonnes</h3><div class="card-body"><div class="bars">' + (bars || '<p class="muted-note">Aucune colonne.</p>') + '</div></div></section>' +
      '<section class="card check-wide"><h3>Sources lues</h3><table class="report"><thead><tr><th>Source</th><th>Feuilles et ligne d’en-têtes</th><th>Lignes</th></tr></thead><tbody>' + rows + '</tbody></table></section>' +
      '</div>';
  }

  /* ---------- Onglets ---------- */
  function setTab(tab, silent) {
    state.tab = tab;
    document.body.dataset.tab = tab;
    $$('.tab').forEach((t) => {
      const on = t.dataset.tab === tab;
      t.setAttribute('aria-selected', on ? 'true' : 'false');
      t.tabIndex = on ? 0 : -1;
    });
    const has = !!(state.analysis && state.built);
    $('#panelTable').hidden = !has || tab !== 'table';
    $('#panelCols').hidden = !has || tab !== 'cols';
    $('#panelCheck').hidden = !has || tab !== 'check';
    $('#toolbar .search').style.visibility = tab === 'table' ? '' : 'hidden';
    if (!silent && has) {
      if (tab === 'cols') renderColumns();
      if (tab === 'check') renderCheck();
    }
    renderLayout();
  }

  /* ================================================================== */
  /*  Actions et export                                                  */
  /* ================================================================== */

  const FORMATS = {
    xlsx: { ext: 'xlsx', name: 'Excel', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    xls: { ext: 'xls', name: 'Excel 97-2003', mime: 'application/vnd.ms-excel' },
    ods: { ext: 'ods', name: 'LibreOffice Calc', mime: 'application/vnd.oasis.opendocument.spreadsheet' },
    csv: { ext: 'csv', name: 'CSV, séparateur « ; »', mime: 'text/csv;charset=utf-8' },
    csvc: { ext: 'csv', name: 'CSV, séparateur « , »', mime: 'text/csv;charset=utf-8' },
    tsv: { ext: 'tsv', name: 'TSV (tabulation)', mime: 'text/tab-separated-values;charset=utf-8' },
    json: { ext: 'json', name: 'JSON', mime: 'application/json' },
  };

  function canShareFiles() {
    try {
      return !!(navigator.canShare && navigator.share && navigator.canShare({ files: [new File(['x'], 'a.xlsx', { type: FORMATS.xlsx.mime })] }));
    } catch (e) { return false; }
  }

  function renderActions() {
    if (!FORMATS[S.format]) S.format = 'xlsx';
    $('#dlLabel').textContent = 'Exporter';
    $('#dlBtn').title = 'Exporter en .' + FORMATS[S.format].ext + ' (Ctrl + S)';
    $('#fmtMenu').innerHTML =
      '<div class="menu-label">Exporter en</div>' +
      Object.keys(FORMATS).map((k) => '<button type="button" role="menuitemradio" class="menu-item" data-fmt="' + k + '" aria-checked="' + (k === S.format) + '"><span>' + esc(FORMATS[k].name) + ' <span style="color:var(--faint)">.' + FORMATS[k].ext + '</span></span><svg class="ic tick"><use href="#i-check"/></svg></button>').join('') +
      '<div class="menu-sep"></div>' +
      '<button type="button" role="menuitem" class="menu-item" data-act="copy"><svg class="ic"><use href="#i-copy"/></svg>Copier le tableau</button>' +
      (canShareFiles() ? '<button type="button" role="menuitem" class="menu-item" data-act="share"><svg class="ic"><use href="#i-share"/></svg>Partager le fichier…</button>' : '');
    updateActions();
  }

  function updateActions() {
    const ok = !!(state.built && state.built.rows.length);
    $('#dlBtn').disabled = !ok;
    $('#fmtBtn').disabled = !ok;
  }

  function stamp(kind) {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return kind === 'time' ? p(d.getHours()) + 'h' + p(d.getMinutes()) : d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }

  function baseFileName() {
    const b = state.built;
    const name = String(S.fileName || '')
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

  function rowsForExport(b) {
    return b.rows.map((r) => r.map((v) => (v === null || v === undefined || v === '' ? S.emptyText || null : v)));
  }

  const toLocalDate = (d) => new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes());

  function reportRows() {
    const rows = [['Source', 'Feuilles lues', 'Lignes']];
    state.analysis.report.forEach((r) => rows.push([r.name, r.error ? 'Erreur' : r.sheets.filter((s) => !s.empty).length, r.rows]));
    rows.push([]);
    rows.push(['Total', '', state.built.rows.length]);
    return rows;
  }

  async function makeXlsx(b) {
    await loadLib('exceljs');
    const wb = new window.ExcelJS.Workbook();
    wb.creator = 'MayaExcel';
    wb.created = new Date();
    const ws = wb.addWorksheet('Fusion', { views: S.freeze ? [{ state: 'frozen', ySplit: 1 }] : [] });
    const data = rowsForExport(b);
    ws.addRow(b.headers);
    data.forEach((r) => ws.addRow(r));

    const heavy = data.length > 20000;
    const line = { style: 'thin', color: { argb: 'FFDDE1E8' } };
    const head = ws.getRow(1);
    head.height = 22;
    const hc = S.headerColor;
    head.eachCell((cell) => {
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      if (hc === 'none') { cell.font = { bold: true }; return; }
      const dark = hc === 'dark' || hc === 'accent';
      cell.font = { bold: true, color: { argb: dark ? 'FFFFFFFF' : 'FF101828' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hc === 'dark' ? 'FF101828' : hc === 'accent' ? 'FF1D4ED8' : 'FFEEF1F6' } };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFB7BDC9' } } };
    });

    b.headers.forEach((h, j) => {
      let width = String(h).length;
      let hasDate = false;
      const sample = Math.min(data.length, 2000);
      for (let i = 0; i < sample; i++) {
        const v = data[i][j];
        if (v instanceof Date) { hasDate = true; width = Math.max(width, 14); }
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
          if (S.zebra && i % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F8FA' } };
        }
        const alt = b.conflictMap[i + ':' + j];
        if (alt) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF4D6' } };
          cell.note = 'Autre valeur dans une autre source : ' + alt;
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
        const t = ws2.addRow(['Valeurs différentes selon les sources', 'Colonne', 'Valeur retenue', 'Autre valeur']);
        t.font = { bold: true };
        ws2.getColumn(4).width = 30;
        keys.forEach((k) => {
          const [r, c] = k.split(':').map(Number);
          ws2.addRow(['Ligne ' + (r + 1), b.headers[c], cellText(b.rows[r][c]), b.conflictMap[k]]);
        });
      }
      if (b.merges.length) {
        ws2.addRow([]);
        const t = ws2.addRow(['Personnes regroupées', 'Sources', 'Reconnue par']);
        t.font = { bold: true };
        b.merges.forEach((m) => ws2.addRow([m.label, m.sources.join(' · '), m.by.join(', ')]));
      }
    }
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: FORMATS.xlsx.mime });
  }

  async function makeSheetJs(b, bookType) {
    await loadLib('xlsx');
    const XLSX = window.XLSX;
    const data = [b.headers].concat(rowsForExport(b).map((r) => r.map((v) => (v instanceof Date ? toLocalDate(v) : v))));
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
    return JSON.stringify(b.rows.map((r) => {
      const o = {};
      r.forEach((v, j) => { o[keys[j]] = v instanceof Date ? v.toISOString().slice(0, v.getUTCHours() || v.getUTCMinutes() ? 16 : 10) : v === undefined ? null : v; });
      return o;
    }), null, 2);
  }

  async function makeExport(fmtKey) {
    const b = state.built;
    const fmt = FORMATS[fmtKey] ? fmtKey : 'xlsx';
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

  async function doDownload(fmt) {
    if (!state.built || !state.built.rows.length) return;
    const btn = $('#dlBtn');
    const label = $('#dlLabel');
    const old = label.textContent;
    btn.disabled = true;
    label.textContent = 'Préparation…';
    try {
      const { blob, name } = await makeExport(fmt || S.format);
      await download(blob, name);
      toast('Fichier exporté : ' + name);
      state.exported = true;
      renderStepper();
      addHistory(name, fmt || S.format);
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
      ta.style.cssText = 'position:fixed;opacity:0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); } catch (e2) { /* ignoré */ }
      ta.remove();
    }
    toast('Tableau copié : collez-le dans Excel.');
  }

  async function doShare() {
    try {
      const { blob, name } = await makeExport(S.format);
      await navigator.share({ files: [new File([blob], name, { type: blob.type })], title: name });
    } catch (e) {
      if (e && e.name !== 'AbortError') toast('Le partage n’est pas disponible ici.', true);
    }
  }

  /* ---------- Menu d'export ---------- */
  const fmtMenu = $('#fmtMenu');
  const fmtBtn = $('#fmtBtn');
  function toggleMenu(open) {
    const show = open === undefined ? fmtMenu.hidden : open;
    fmtMenu.hidden = !show;
    fmtBtn.setAttribute('aria-expanded', show ? 'true' : 'false');
    if (show) {
      const sel = fmtMenu.querySelector('[aria-checked="true"]') || fmtMenu.querySelector('.menu-item');
      if (sel) sel.focus();
    }
  }
  fmtBtn.addEventListener('click', (e) => { e.stopPropagation(); toggleMenu(); });
  fmtMenu.addEventListener('click', (e) => {
    const it = e.target.closest('.menu-item');
    if (!it) return;
    toggleMenu(false);
    if (it.dataset.fmt) {
      S.format = it.dataset.fmt;
      saveSettings();
      renderActions();
      doDownload(S.format);
    } else if (it.dataset.act === 'copy') doCopy();
    else if (it.dataset.act === 'share') doShare();
  });
  fmtMenu.addEventListener('keydown', (e) => {
    const items = $$('.menu-item', fmtMenu);
    const i = items.indexOf(document.activeElement);
    if (e.key === 'ArrowDown') { e.preventDefault(); items[(i + 1) % items.length].focus(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); items[(i + items.length - 1) % items.length].focus(); }
    else if (e.key === 'Escape') { toggleMenu(false); fmtBtn.focus(); }
  });
  document.addEventListener('click', (e) => { if (!fmtMenu.hidden && !e.target.closest('#exportSplit')) toggleMenu(false); });

  /* ================================================================== */
  /*  Réglages                                                           */
  /* ================================================================== */

  const SECTION_ORDER = ['fusion', 'data', 'table', 'export', 'look', 'templates', 'history', 'help', 'app'];
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
        { t: 'html', html: presetSelectHtml },
        { t: 'group', l: 'Lecture des sources' },
        { k: 'allSheets', t: 'switch', l: 'Lire toutes les feuilles', h: 'Sinon, seule la première feuille de chaque fichier est lue.', a: 'analyze' },
        { k: 'formDetect', t: 'switch', l: 'Lire les fiches individuelles', h: 'Feuilles de la forme « Nom : … », « Téléphone : … » : une fiche devient une ligne.', a: 'analyze' },
        { k: 'multiHeader', t: 'switch', l: 'Titres répartis sur deux lignes', h: 'Une ligne de regroupement (Père, Mère…) lève les doublons de titres.', a: 'analyze' },
        { k: 'nullTokens', t: 'switch', l: 'Ignorer les valeurs « N/A », « néant », « - »', a: 'analyze' },
        { k: 'headerMode', t: 'select', l: 'Ligne des titres', a: 'analyze', o: [{ v: 'auto', l: 'Détection automatique' }, { v: 'fixed', l: 'Numéro de ligne précis' }, { v: 'none', l: 'Aucune (sans titres)' }] },
        { k: 'headerRow', t: 'number', l: 'Numéro de la ligne des titres', min: 1, max: 200, a: 'analyze', when: () => S.headerMode === 'fixed' },
        { t: 'group', l: 'Rapprochement des colonnes' },
        { k: 'sensitivity', t: 'seg', l: 'Sensibilité', h: 'Strict : noms quasi identiques. Large : accepte des intitulés plus éloignés.', a: 'analyze', o: [{ v: 'strict', l: 'Strict' }, { v: 'balanced', l: 'Équilibré' }, { v: 'large', l: 'Large' }], stack: true },
        { k: 'contentMatch', t: 'switch', l: 'Reconnaître les colonnes par leur contenu', h: 'E-mails, téléphones, sexe, listes de valeurs communes.', a: 'analyze' },
        { k: 'splitNames', t: 'switch', l: 'Séparer « Nom et Prénoms »', h: 'Quand d’autres sources ont Nom et Prénom dans des colonnes distinctes.', a: 'analyze' },
        { t: 'group', l: 'Reconnaissance des personnes' },
        { k: 'combine', t: 'switch', l: 'Regrouper les lignes d’une même personne', h: 'Retire aussi les lignes strictement identiques.', a: 'build' },
        { k: 'personMatch', t: 'seg', l: 'Rigueur du regroupement', h: 'Strict : matricule, e-mail ou téléphone. Normal : ajoute le nom. Large : tolère aussi les fautes de frappe.', a: 'build', o: [{ v: 'strict', l: 'Strict' }, { v: 'normal', l: 'Normal' }, { v: 'large', l: 'Large' }], stack: true, when: () => S.combine },
        { k: 'fuzzyNames', t: 'switch', l: 'Tolérer les fautes dans les noms', h: 'Agbodjan / Agbodjang, accents oubliés.', a: 'build', when: () => S.combine },
        { k: 'useBirth', t: 'switch', l: 'Départager les homonymes par la date de naissance', a: 'build', when: () => S.combine },
        { k: 'keyCol', t: 'select', l: 'Reconnaître une ligne d’après', a: 'build', when: () => S.combine, o: () => [{ v: 'auto', l: 'Automatique (personne)' }].concat(state.cols.filter((c) => c.include && !c.into).map((c) => ({ v: c.key, l: 'Colonne « ' + c.label + ' »' }))) },
        { k: 'conflict', t: 'select', l: 'Valeurs différentes', a: 'build', when: () => S.combine, o: [{ v: 'first', l: 'Garder la première source' }, { v: 'last', l: 'Garder la dernière' }, { v: 'longest', l: 'Garder la plus complète' }, { v: 'majority', l: 'Garder la plus fréquente' }] },
        { t: 'group', l: 'Règles personnalisées' },
        { k: 'ignoreColumns', t: 'textarea', l: 'Colonnes à ignorer', h: 'Intitulés séparés par des virgules. Exemple : Signature, Observations', a: 'analyze' },
        { k: 'synonyms', t: 'textarea', l: 'Équivalences d’intitulés', h: 'Une règle par ligne : intitulé = colonne. Exemple : Contact perso = Téléphone', a: 'analyze' },
      ] },
      data: { title: 'Données', items: [
        { k: 'clean', t: 'switch', l: 'Harmoniser les données', h: 'Applique les règles ci-dessous.', a: 'build' },
        { k: 'nomCase', t: 'select', l: 'Noms de famille', a: 'build', when: () => S.clean, o: [{ v: 'upper', l: 'MAJUSCULES' }, { v: 'title', l: 'Première Lettre Majuscule' }, { v: 'keep', l: 'Ne pas modifier' }] },
        { k: 'prenomCase', t: 'select', l: 'Prénoms', a: 'build', when: () => S.clean, o: [{ v: 'title', l: 'Première Lettre Majuscule' }, { v: 'upper', l: 'MAJUSCULES' }, { v: 'keep', l: 'Ne pas modifier' }] },
        { k: 'emailLower', t: 'switch', l: 'E-mails en minuscules', a: 'build', when: () => S.clean },
        { k: 'sexNorm', t: 'switch', l: 'Sexe : Masculin / Féminin', h: 'M, H, Homme → Masculin ; F, Femme → Féminin.', a: 'build', when: () => S.clean },
        { k: 'normCats', t: 'switch', l: 'Situation matrimoniale et Oui / Non', h: 'Uniformise « célibataire », « Marié(e) », « oui », « yes »…', a: 'build', when: () => S.clean },
        { k: 'placeCase', t: 'switch', l: 'Majuscule aux villes et pays', a: 'build', when: () => S.clean },
        { k: 'unifyLabels', t: 'switch', l: 'Unifier les libellés proches', h: 'Poste, service, ville… : « Informaticien » / « Informaticienne », fautes de frappe.', a: 'build', when: () => S.clean },
        { k: 'phoneFormat', t: 'select', l: 'Numéros de téléphone', a: 'build', when: () => S.clean, o: [{ v: 'keep', l: 'Ne pas modifier' }, { v: 'digits', l: 'Chiffres seulement' }, { v: 'spaced', l: 'Groupes de 2 chiffres' }, { v: 'intl', l: 'Avec indicatif du pays' }] },
        { k: 'phoneCC', t: 'text', l: 'Indicatif du pays', h: 'Exemple : 228 pour le Togo.', a: 'build', when: () => S.clean && (S.phoneFormat === 'intl' || S.phoneFormat === 'spaced') },
        { t: 'group', l: 'Affichage des valeurs' },
        { k: 'dateFormat', t: 'select', l: 'Format des dates', a: 'view', o: [{ v: 'dmy', l: '31/12/2025' }, { v: 'ymd', l: '2025-12-31' }, { v: 'mdy', l: '12/31/2025' }, { v: 'long', l: '31 décembre 2025' }] },
        { k: 'emptyText', t: 'select', l: 'Cases vides', a: 'view', o: [{ v: '', l: 'Laisser vide' }, { v: '—', l: 'Tiret (—)' }, { v: 'N/A', l: 'N/A' }, { v: '-', l: '-' }] },
        { k: 'dropEmptyCols', t: 'switch', l: 'Masquer les colonnes vides', a: 'build' },
      ] },
      table: { title: 'Tableau', items: [
        { k: 'numbering', t: 'switch', l: 'Colonne N° (numérotation)', a: 'build' },
        { k: 'source', t: 'switch', l: 'Colonne « Fichier source »', a: 'build' },
        { t: 'group', l: 'Tri' },
        { k: 'sortBy', t: 'select', l: 'Trier par', a: 'build', o: () => colOptions(true, 'Ordre des sources') },
        { k: 'sortDir', t: 'seg', l: 'Sens', a: 'build', o: [{ v: 'asc', l: 'A → Z' }, { v: 'desc', l: 'Z → A' }], when: () => S.sortBy !== 'auto' && S.sortBy !== 'none' },
        { k: 'sortBy2', t: 'select', l: 'Puis par', a: 'build', when: () => S.sortBy !== 'auto' && S.sortBy !== 'none', o: () => colOptions(false) },
        { k: 'sortDir2', t: 'seg', l: 'Sens (2e tri)', a: 'build', o: [{ v: 'asc', l: 'A → Z' }, { v: 'desc', l: 'Z → A' }], when: () => S.sortBy !== 'auto' && S.sortBy !== 'none' && S.sortBy2 !== 'none' },
        { t: 'group', l: 'Aperçu' },
        { k: 'previewRows', t: 'select', l: 'Lignes affichées', a: 'view', num: true, o: [{ v: 100, l: '100' }, { v: 500, l: '500' }, { v: 1000, l: '1 000' }, { v: 5000, l: '5 000' }] },
      ] },
      export: { title: 'Export', items: [
        { k: 'format', t: 'select', l: 'Format par défaut', a: 'export', o: Object.keys(FORMATS).map((k) => ({ v: k, l: FORMATS[k].name + ' (.' + FORMATS[k].ext + ')' })) },
        { k: 'fileName', t: 'text', l: 'Nom du fichier', h: 'Variables : {date}, {heure}, {n} (lignes), {fichiers}.', a: 'export' },
        { t: 'group', l: 'Fichier Excel' },
        { k: 'headerColor', t: 'select', l: 'Ligne de titres', a: 'export', o: [{ v: 'light', l: 'Fond clair' }, { v: 'dark', l: 'Fond foncé' }, { v: 'accent', l: 'Fond bleu' }, { v: 'none', l: 'Sans couleur' }] },
        { k: 'zebra', t: 'switch', l: 'Lignes alternées', a: 'export' },
        { k: 'freeze', t: 'switch', l: 'Figer la ligne de titres', a: 'export' },
        { k: 'autofilter', t: 'switch', l: 'Ajouter les filtres', a: 'export' },
        { k: 'reportSheet', t: 'switch', l: 'Ajouter une feuille « Rapport »', h: 'Sources lues, valeurs différentes, personnes regroupées.', a: 'export' },
        { t: 'group', l: 'Fichiers texte' },
        { k: 'csvSafe', t: 'switch', l: 'Protéger contre les formules', h: 'Évite qu’une cellule commençant par = + - @ soit exécutée.', a: 'export' },
      ] },
      look: { title: 'Affichage', items: [
        { k: 'density', t: 'seg', l: 'Densité du tableau', a: 'ui', o: [{ v: 'comfortable', l: 'Confortable' }, { v: 'compact', l: 'Compacte' }] },
        { k: 'fontScale', t: 'seg', l: 'Taille du texte', a: 'ui', num: true, o: [{ v: 90, l: 'Petit' }, { v: 100, l: 'Normal' }, { v: 115, l: 'Grand' }, { v: 130, l: 'Très grand' }] },
        { k: 'motion', t: 'seg', l: 'Animations', a: 'ui', o: [{ v: 'system', l: 'Système' }, { v: 'on', l: 'Activées' }, { v: 'off', l: 'Réduites' }] },
      ] },
      templates: { title: 'Modèles', items: [{ t: 'html', html: templatesSectionHtml }] },
      history: { title: 'Historique', items: [{ t: 'html', html: historySectionHtml }] },
      help: { title: 'Aide', items: [{ t: 'html', html: helpSectionHtml }] },
      app: { title: 'Application', items: [{ t: 'html', html: appSectionHtml }] },
    };
  }

  function controlHtml(it) {
    if (it.t === 'group') return '<p class="group-title">' + esc(it.l) + '</p>';
    if (it.t === 'html') return it.html();
    if (it.when && !it.when()) return '';
    const id = 's-' + it.k;
    const hint = it.h ? '<p class="setting-hint">' + esc(it.h) + '</p>' : '';
    const v = S[it.k];
    const cls = 'setting' + (it.stack || it.t === 'textarea' ? ' stack' : '');
    if (it.t === 'switch') return '<div class="' + cls + '"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label><input type="checkbox" role="switch" class="switch" id="' + id + '" data-k="' + it.k + '" ' + (v ? 'checked' : '') + '>' + hint + '</div>';
    const opts = typeof it.o === 'function' ? it.o() : it.o;
    if (it.t === 'select') {
      const cur = opts.some((o) => String(o.v) === String(v)) ? v : opts[0] && opts[0].v;
      return '<div class="' + cls + '"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label><select id="' + id + '" data-k="' + it.k + '">' +
        opts.map((o) => '<option value="' + esc(o.v) + '"' + (String(o.v) === String(cur) ? ' selected' : '') + '>' + esc(o.l) + '</option>').join('') + '</select>' + hint + '</div>';
    }
    if (it.t === 'seg') return '<div class="' + cls + '"><span class="setting-label" id="' + id + '-l">' + esc(it.l) + '</span><div class="seg" role="radiogroup" aria-labelledby="' + id + '-l">' +
      opts.map((o) => '<button type="button" role="radio" data-k="' + it.k + '" data-v="' + esc(o.v) + '" aria-checked="' + (String(o.v) === String(v)) + '">' + esc(o.l) + '</button>').join('') + '</div>' + hint + '</div>';
    if (it.t === 'text') return '<div class="' + cls + '"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label><input type="text" id="' + id + '" data-k="' + it.k + '" value="' + esc(v) + '" autocomplete="off">' + hint + '</div>';
    if (it.t === 'number') return '<div class="' + cls + '"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label><input type="number" id="' + id + '" data-k="' + it.k + '" min="' + it.min + '" max="' + it.max + '" value="' + esc(v) + '" inputmode="numeric">' + hint + '</div>';
    if (it.t === 'textarea') return '<div class="' + cls + '"><label class="setting-label" for="' + id + '">' + esc(it.l) + '</label>' + hint + '<textarea id="' + id + '" data-k="' + it.k + '" rows="3" spellcheck="false">' + esc(v) + '</textarea></div>';
    return '';
  }

  function renderSettings() {
    const sc = schema();
    $('#settingsTabs').innerHTML = SECTION_ORDER.map((k) => '<button type="button" role="tab" data-sec="' + k + '" aria-selected="' + (k === settingsSection) + '">' + esc(sc[k].title) + '</button>').join('');
    $('#settingsBody').innerHTML = sc[settingsSection].items.map(controlHtml).join('');
  }

  const UI_KEYS = ['density', 'fontScale', 'motion'];
  function updateBadge() {
    const n = Object.keys(DEFAULTS).filter((k) => !UI_KEYS.includes(k) && S[k] !== DEFAULTS[k]).length;
    $('#settingsBadge').hidden = !n;
  }

  function applyChange(kind) {
    saveSettings();
    updateBadge();
    if (kind === 'analyze') recompute();
    else if (kind === 'build') { if (state.analysis) rebuild(true); }
    else if (kind === 'view') {
      if (state.built) { state.quality = computeQuality(); renderTable(); renderCheck(); }
    } else if (kind === 'ui') applyUi();
    renderActions();
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
  }

  function openSettings(section) {
    if (section) settingsSection = section;
    renderSettings();
    const d = $('#settingsDlg');
    if (!d.open) d.showModal();
  }

  /* ================================================================== */
  /*  Types de fusion, modèles, historique                               */
  /* ================================================================== */

  const PRESET_KEYS = ['allSheets', 'formDetect', 'multiHeader', 'nullTokens', 'headerMode', 'headerRow', 'sensitivity', 'contentMatch', 'splitNames', 'combine', 'personMatch', 'fuzzyNames', 'useBirth', 'keyCol', 'conflict', 'ignoreColumns', 'synonyms', 'clean', 'nomCase', 'prenomCase', 'emailLower', 'sexNorm', 'normCats', 'placeCase', 'unifyLabels', 'phoneFormat', 'phoneCC', 'dropEmptyCols', 'numbering', 'source', 'sortBy', 'sortDir', 'sortBy2', 'sortDir2'];
  const TEMPLATE_KEYS = PRESET_KEYS.concat(['dateFormat', 'emptyText', 'previewRows', 'format', 'fileName', 'zebra', 'freeze', 'autofilter', 'headerColor', 'reportSheet', 'csvSafe']);
  const PRESETS = [
    { id: 'personnes', name: 'Personnes et effectifs', desc: 'Une ligne par personne : doublons entre sources regroupés, Nom et Prénom séparés, données harmonisées.', set: {} },
    { id: 'chiffres', name: 'Chiffres et comptabilité', desc: 'Toutes les lignes sont empilées telles quelles, avec la source de chaque ligne.', set: { combine: false, clean: false, splitNames: false, sortBy: 'none', source: true } },
    { id: 'listes', name: 'Listes et inventaires', desc: 'Les lignes qui partagent un identifiant sont regroupées ; la dernière source l’emporte en cas de désaccord.', set: { personMatch: 'strict', conflict: 'last', clean: false, splitNames: false } },
    { id: 'empiler', name: 'Empilement simple', desc: 'Aucun regroupement ni nettoyage : toutes les lignes sont conservées dans l’ordre des sources.', set: { combine: false, clean: false, unifyLabels: false, splitNames: false, contentMatch: false, sortBy: 'none', source: true } },
    { id: 'strict', name: 'Contrôle strict', desc: 'Rapprochements prudents : intitulés quasi identiques, personnes reconnues par identifiant uniquement.', set: { sensitivity: 'strict', personMatch: 'strict', fuzzyNames: false, unifyLabels: false } },
  ];
  const presetValues = (p) => Object.fromEntries(PRESET_KEYS.map((k) => [k, k in p.set ? p.set[k] : DEFAULTS[k]]));
  const currentPresetId = () => {
    const p = PRESETS.find((x) => { const v = presetValues(x); return PRESET_KEYS.every((k) => S[k] === v[k]); });
    return p ? p.id : null;
  };

  const TPL_KEY = 'maya.templates.v1';
  const HIST_KEY = 'maya.history.v1';
  function readList(key) {
    try { const v = JSON.parse(localStorage.getItem(key) || '[]'); return Array.isArray(v) ? v : []; } catch (e) { return []; }
  }
  function writeList(key, v) {
    try { localStorage.setItem(key, JSON.stringify(v)); } catch (e) { /* ignoré */ }
  }
  const when = (ts) => new Date(ts).toLocaleString('fr-FR', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });

  function snapshot() {
    const settings = {};
    TEMPLATE_KEYS.forEach((k) => (settings[k] = S[k]));
    return { settings, overrides: state.overrides, separate: state.separate.slice(), colOrder: state.colOrder.slice() };
  }

  function applySnapshot(snap) {
    if (!snap || !snap.settings) return false;
    TEMPLATE_KEYS.forEach((k) => {
      S[k] = k in snap.settings && typeof snap.settings[k] === typeof DEFAULTS[k] ? snap.settings[k] : DEFAULTS[k];
    });
    state.overrides = snap.overrides && typeof snap.overrides === 'object' ? snap.overrides : {};
    state.separate = Array.isArray(snap.separate) ? snap.separate : [];
    state.colOrder = Array.isArray(snap.colOrder) ? snap.colOrder : [];
    saveSettings();
    if (state.analysis) recompute();
    else renderAll();
    return true;
  }

  function applyPreset(id) {
    const p = PRESETS.find((x) => x.id === id);
    if (!p) return;
    const v = presetValues(p);
    PRESET_KEYS.forEach((k) => (S[k] = v[k]));
    saveSettings();
    if (state.analysis) recompute();
    else renderAll();
    toast('Type de fusion : ' + p.name);
  }

  function saveTemplate(name) {
    const n = String(name || '').trim();
    if (!n) { toast('Donnez un nom au modèle.', true); return; }
    const list = readList(TPL_KEY).filter((t) => t.name.toLowerCase() !== n.toLowerCase());
    list.unshift({ id: 't' + Date.now(), name: n, created: Date.now(), snap: snapshot() });
    writeList(TPL_KEY, list.slice(0, 30));
    toast('Modèle « ' + n + ' » enregistré.');
  }

  function exportTemplate(t) {
    const blob = new Blob([JSON.stringify({ app: 'MayaExcel', version: 1, name: t.name, snap: t.snap }, null, 2)], { type: 'application/json' });
    download(blob, 'modele-' + t.name.replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').toLowerCase() + '.mayaexcel.json');
  }

  async function importTemplate(file) {
    try {
      const o = JSON.parse(await file.text());
      if (!o || !o.snap || !o.snap.settings) throw new Error('format');
      const list = readList(TPL_KEY);
      list.unshift({ id: 't' + Date.now(), name: String(o.name || file.name.replace(/\.json$/i, '')).slice(0, 80), created: Date.now(), snap: o.snap });
      writeList(TPL_KEY, list.slice(0, 30));
      toast('Modèle importé.');
    } catch (e) {
      toast('Ce fichier n’est pas un modèle MayaExcel valide.', true);
    }
  }

  function addHistory(name, fmt) {
    const b = state.built;
    if (!b) return;
    const list = readList(HIST_KEY);
    list.unshift({
      id: 'h' + Date.now(), t: Date.now(), name, fmt, rows: b.rows.length, cols: b.columnCount,
      sources: state.analysis ? state.analysis.fileCount : 0, files: state.files.filter((f) => f.include !== false).map((f) => f.name).slice(0, 8), snap: snapshot(),
    });
    writeList(HIST_KEY, list.slice(0, 15));
    renderHome();
  }

  function presetSelectHtml() {
    const cur = currentPresetId();
    return '<div class="setting stack" style="padding-top:14px"><label class="setting-label" for="presetSel">Type de fusion</label><select id="presetSel">' +
      (cur ? '' : '<option value="">Personnalisé</option>') + PRESETS.map((p) => '<option value="' + p.id + '"' + (p.id === cur ? ' selected' : '') + '>' + esc(p.name) + '</option>').join('') +
      '</select><p class="setting-hint">Applique un ensemble de réglages cohérent. Chaque réglage reste modifiable ensuite.</p></div>';
  }

  const EMPTY_LINE = (t) => '<p class="muted-note">' + esc(t) + '</p>';

  function tplRows(compact) {
    const list = readList(TPL_KEY);
    if (!list.length) return EMPTY_LINE('Aucun modèle enregistré. Un modèle garde vos réglages et vos corrections de colonnes pour les réutiliser sur un prochain lot de fichiers.');
    return '<ul class="rows">' + list.slice(0, compact ? 4 : 30).map((t) =>
      '<li class="row-item"><div class="row-main"><b>' + esc(t.name) + '</b><small>Enregistré le ' + esc(when(t.created)) + '</small></div><div class="row-actions">' +
      '<button type="button" class="btn btn-sm" data-tpl="apply" data-id="' + t.id + '">Appliquer</button>' +
      (compact ? '' : '<button type="button" class="icon-btn" data-tpl="export" data-id="' + t.id + '" aria-label="Exporter le modèle ' + esc(t.name) + '" title="Exporter"><svg class="ic"><use href="#i-download"/></svg></button><button type="button" class="icon-btn" data-tpl="delete" data-id="' + t.id + '" aria-label="Supprimer le modèle ' + esc(t.name) + '" title="Supprimer"><svg class="ic"><use href="#i-trash"/></svg></button>') +
      '</div></li>').join('') + '</ul>';
  }

  function histRows(compact) {
    const list = readList(HIST_KEY);
    if (!list.length) return EMPTY_LINE('Aucune fusion exportée pour le moment.');
    return '<ul class="rows">' + list.slice(0, compact ? 4 : 15).map((h) =>
      '<li class="row-item"><div class="row-main"><b title="' + esc(h.name) + '">' + esc(h.name) + '</b><small>' + esc(when(h.t)) + ' · ' + plural(h.rows, 'ligne', 'lignes') + ' · ' + plural(h.sources, 'source', 'sources') + '</small>' +
      (compact ? '' : '<small>' + esc((h.files || []).join(', ')) + '</small>') + '</div><div class="row-actions"><button type="button" class="btn btn-sm" data-hist="' + h.id + '" title="Réappliquer les réglages de cette fusion">Réutiliser</button></div></li>').join('') + '</ul>';
  }

  function renderHome() {
    const cur = currentPresetId();
    $('#presetSub').textContent = cur ? '' : 'Réglages personnalisés';
    $('#homePresets').innerHTML = PRESETS.map((p) =>
      '<button type="button" role="radio" class="preset" data-preset="' + p.id + '" aria-checked="' + (p.id === cur) + '"><span class="preset-name">' + esc(p.name) + '<svg class="ic tick"><use href="#i-check"/></svg></span><span class="preset-desc">' + esc(p.desc) + '</span></button>').join('');
    $('#homeTemplates').innerHTML = tplRows(true);
    $('#homeHistory').innerHTML = histRows(true);
  }

  function templatesSectionHtml() {
    return '<p class="setting-hint-wide" style="padding-top:12px">Un modèle mémorise tous les réglages de fusion et d’export, ainsi que vos corrections de colonnes (renommages, regroupements, ordre). Appliquez-le à un nouveau lot de fichiers pour obtenir le même résultat.</p>' +
      '<p class="group-title">Enregistrer la configuration actuelle</p>' +
      '<div class="inline-form"><input type="text" id="tplName" placeholder="Nom du modèle" maxlength="60" autocomplete="off"><button type="button" class="btn btn-primary" data-tpl="save"><svg class="ic"><use href="#i-save"/></svg>Enregistrer</button></div>' +
      '<p class="group-title">Modèles enregistrés</p>' + tplRows(false) +
      '<div style="margin-top:12px"><button type="button" class="btn" data-tpl="import"><svg class="ic"><use href="#i-upload"/></svg>Importer un modèle</button></div>';
  }

  function historySectionHtml() {
    return '<p class="setting-hint-wide" style="padding-top:12px">Les 15 derniers exports, conservés sur cet appareil. « Réutiliser » réapplique les réglages de la fusion correspondante (les fichiers eux-mêmes ne sont jamais conservés).</p>' + histRows(false) +
      (readList(HIST_KEY).length ? '<div style="margin-top:12px"><button type="button" class="btn btn-ghost" data-hist-clear="1">Vider l’historique</button></div>' : '');
  }

  function helpSectionHtml() {
    return '<div class="help">' +
      '<h3>Fusionner en trois étapes</h3><ol class="steps"><li><b>Sources</b> : ajoutez les fichiers (glisser-déposer, bouton Ajouter, ou collage d’un tableau copié dans Excel).</li><li><b>Vérification</b> : contrôlez le tableau, l’interprétation des colonnes (onglet Colonnes) et les anomalies (onglet Contrôle).</li><li><b>Export</b> : choisissez le format. Un rapport peut être joint au fichier Excel.</li></ol>' +
      '<h3>Comment les colonnes sont rapprochées</h3><p>Par leur intitulé (synonymes, abréviations, accents, fautes de frappe), puis par leur contenu (e-mails, téléphones, sexe, listes de valeurs communes). Quand un rapprochement est incertain, la colonne porte une mention et vous pouvez la séparer ou la rattacher à la main.</p>' +
      '<h3>Comment les personnes sont reconnues</h3><p>Par matricule, e-mail, téléphone puis nom (ordre inversé, accents et fautes tolérés). Deux personnes ayant le même nom mais des dates de naissance différentes restent distinctes. Chaque regroupement figure dans Contrôle et peut être dissocié.</p>' +
      '<h3>Valeurs en conflit</h3><p>Quand deux sources donnent des valeurs différentes pour la même personne, une seule est retenue (réglable : première, dernière, plus complète, plus fréquente). La cellule est marquée d’un coin orange et la valeur écartée est consultable.</p>' +
      '<h3>Fiches individuelles</h3><p>Une feuille de la forme « Nom : … », « Téléphone : … » est lue comme une seule ligne. Utile quand chaque personne envoie sa propre fiche.</p>' +
      '<h3>Raccourcis</h3><p><span class="kbd">Ctrl</span> + <span class="kbd">O</span> ajouter · <span class="kbd">Ctrl</span> + <span class="kbd">S</span> exporter · <span class="kbd">Ctrl</span> + <span class="kbd">,</span> réglages · <span class="kbd">Ctrl</span> + <span class="kbd">V</span> coller · flèches et <span class="kbd">Ctrl</span> + <span class="kbd">C</span> dans le tableau.</p>' +
      '<h3>Confidentialité</h3><p>Tout est traité sur votre appareil. Les fichiers ne sont ni envoyés ni conservés ; seuls vos réglages, modèles et le résumé de l’historique sont mémorisés localement.</p></div>';
  }

  document.addEventListener('click', (e) => {
    const sec = e.target.closest('[data-sec]');
    if (sec && !sec.closest('#settingsTabs')) { openSettings(sec.dataset.sec); return; }
    const home = e.target.closest('[data-rail="home"]');
    if (home) { if (state.files.length) { setTab('table'); setView('main'); } return; }
    const pr = e.target.closest('[data-preset]');
    if (pr) { applyPreset(pr.dataset.preset); return; }
    const hc = e.target.closest('[data-hist-clear]');
    if (hc) { writeList(HIST_KEY, []); renderHome(); if ($('#settingsDlg').open) renderSettings(); return; }
    const h = e.target.closest('[data-hist]');
    if (h) {
      const entry = readList(HIST_KEY).find((x) => x.id === h.dataset.hist);
      if (entry && applySnapshot(entry.snap)) { toast('Réglages de cette fusion réappliqués.'); if ($('#settingsDlg').open) $('#settingsDlg').close(); }
      return;
    }
    const t = e.target.closest('[data-tpl]');
    if (!t) return;
    const act = t.dataset.tpl;
    const list = readList(TPL_KEY);
    const tpl = list.find((x) => x.id === t.dataset.id);
    if (act === 'save') { saveTemplate($('#tplName') ? $('#tplName').value : ''); }
    else if (act === 'apply' && tpl) { if (applySnapshot(tpl.snap)) { toast('Modèle « ' + tpl.name + ' » appliqué.'); if ($('#settingsDlg').open) $('#settingsDlg').close(); } return; }
    else if (act === 'export' && tpl) { exportTemplate(tpl); return; }
    else if (act === 'delete' && tpl) { writeList(TPL_KEY, list.filter((x) => x.id !== tpl.id)); toast('Modèle supprimé.'); }
    else if (act === 'import') { $('#tplFile').click(); return; }
    renderHome();
    if ($('#settingsDlg').open) renderSettings();
  });
  $('#tplFile').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    await importTemplate(f);
    renderHome();
    if ($('#settingsDlg').open) renderSettings();
  });

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
    const about = '<p class="group-title">Raccourcis clavier</p><p class="setting-hint-wide"><span class="kbd">Ctrl</span> + <span class="kbd">O</span> ajouter des fichiers · <span class="kbd">Ctrl</span> + <span class="kbd">S</span> exporter · <span class="kbd">Ctrl</span> + <span class="kbd">,</span> réglages · <span class="kbd">Ctrl</span> + <span class="kbd">V</span> coller un tableau · flèches : se déplacer dans le tableau · <span class="kbd">Ctrl</span> + <span class="kbd">C</span> copier la cellule</p>' +
      '<p class="group-title">À propos</p><p class="setting-hint-wide">MayaExcel ' + esc((window.MAYA_CONFIG && window.MAYA_CONFIG.version) || '') + '. Vos fichiers ne quittent jamais votre appareil.</p>';
    if (isShell()) return '<p class="setting-hint-wide">Vous utilisez déjà l’application MayaExcel.</p>' + about;
    return '<p class="group-title">Télécharger l’application</p><div style="display:grid;gap:8px;margin-top:8px">' +
      dlCard(FILES.android, 'Android (.apk)', 'Fonctionne sans connexion', 'i-phone') +
      dlCard(FILES.windows, 'Windows (installateur .exe)', 'Installation classique', 'i-monitor') +
      dlCard(FILES.portable, 'Windows (version portable .exe)', 'Sans installation', 'i-monitor') +
      '</div><p class="setting-hint-wide">iPhone : ouvrez ce site dans Safari, touchez Partager puis « Sur l’écran d’accueil ». <a href="' + esc(releasesUrl()) + '" target="_blank" rel="noopener">Toutes les versions</a>.</p>' + about;
  }

  function openAppDialog(autoDownload) {
    let html = '';
    if (PLATFORM === 'android') {
      if (autoDownload) startDownload(FILES.android);
      html = '<p>' + (autoDownload ? 'Le téléchargement de <b>MayaExcel.apk</b> démarre.' : 'Installez MayaExcel sur votre téléphone.') + '</p>' +
        '<ol class="steps"><li>Ouvrez le fichier <b>MayaExcel.apk</b> dans vos téléchargements.</li><li>Si Android le demande, autorisez l’installation depuis ce navigateur.</li><li>Touchez <b>Installer</b>, puis ouvrez MayaExcel.</li></ol>' +
        '<p class="note">Rien ne se passe ? La version est peut-être en cours de publication : réessayez dans quelques minutes.</p>' +
        dlCard(FILES.android, 'Télécharger à nouveau', 'MayaExcel.apk', 'i-download');
    } else if (PLATFORM === 'windows') {
      if (autoDownload) startDownload(FILES.windows);
      html = '<p>' + (autoDownload ? 'Le téléchargement de <b>MayaExcel-Setup.exe</b> démarre.' : 'Installez MayaExcel sur votre ordinateur.') + '</p>' +
        '<ol class="steps"><li>Ouvrez le fichier téléchargé.</li><li>Si Windows affiche « Windows a protégé votre PC », cliquez sur <b>Informations complémentaires</b>, puis <b>Exécuter quand même</b>.</li><li>Suivez l’installation, puis ouvrez MayaExcel depuis le menu Démarrer.</li></ol>' +
        dlCard(FILES.portable, 'Version portable', 'Un seul fichier, sans installation', 'i-monitor');
    } else if (PLATFORM === 'ios') {
      html = '<p>Sur iPhone et iPad, MayaExcel s’ajoute à l’écran d’accueil depuis Safari.</p>' +
        '<ol class="steps"><li>Ouvrez cette page dans <b>Safari</b>.</li><li>Touchez le bouton <b>Partager</b> (le carré avec une flèche vers le haut).</li><li>Choisissez <b>Sur l’écran d’accueil</b>, puis <b>Ajouter</b>.</li></ol>';
    } else {
      html = '<p>Choisissez la version pour votre appareil.</p>' +
        dlCard(FILES.android, 'Android (.apk)', 'Téléphones et tablettes Android', 'i-phone') +
        dlCard(FILES.windows, 'Windows (installateur)', 'Ordinateurs Windows', 'i-monitor') +
        dlCard(FILES.portable, 'Windows (portable)', 'Sans installation', 'i-monitor');
    }
    $('#appBody').innerHTML = html;
    const d = $('#appDlg');
    if (!d.open) d.showModal();
  }

  function setupAppButton() {
    if (isShell() || isStandalone()) return;
    const labels = { android: 'Application Android', windows: 'Application Windows', ios: 'Écran d’accueil' };
    if (!labels[PLATFORM]) return;
    $('#getAppBtn').hidden = false;
    $('#getAppLabel').textContent = labels[PLATFORM];
  }

  /* ================================================================== */
  /*  Événements                                                         */
  /* ================================================================== */

  const fileInput = $('#fileInput');
  const openPicker = () => fileInput.click();

  $$('.js-pick').forEach((b) => b.addEventListener('click', openPicker));
  $$('.js-paste').forEach((b) => b.addEventListener('click', pasteFromButton));
  $$('.js-demo').forEach((b) => b.addEventListener('click', loadDemo));
  $('#drop').addEventListener('click', (e) => { if (!e.target.closest('button')) openPicker(); });
  fileInput.addEventListener('change', () => {
    const list = Array.from(fileInput.files);
    fileInput.value = '';
    addFiles(list);
  });

  let dragDepth = 0;
  const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  window.addEventListener('dragenter', (e) => { if (hasFiles(e)) { e.preventDefault(); dragDepth++; document.body.classList.add('dragging'); } });
  window.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('dragleave', (e) => { if (hasFiles(e) && --dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dragging'); } });
  window.addEventListener('drop', (e) => {
    dragDepth = 0;
    document.body.classList.remove('dragging');
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
    state.noMerge = new Set();
    state.query = '';
    recompute();
  });

  $('#filesList').addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-act]');
    if (!btn) return;
    const id = Number(btn.dataset.id);
    const i = state.files.findIndex((f) => f.id === id);
    if (i < 0) return;
    if (btn.dataset.act === 'remove') {
      state.files.splice(i, 1);
      recompute();
    } else if (btn.dataset.act === 'up' || btn.dataset.act === 'down') {
      const j = btn.dataset.act === 'up' ? i - 1 : i + 1;
      if (j < 0 || j >= state.files.length) return;
      [state.files[i], state.files[j]] = [state.files[j], state.files[i]];
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
      if (state.built) renderTable();
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

  /* Onglets et navigation mobile */
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
  $('#mobilenav').addEventListener('click', (e) => {
    const b = e.target.closest('[data-nav]');
    if (!b) return;
    if (b.dataset.nav === 'sources') setView('sources');
    else {
      document.body.dataset.view = 'main';
      setTab(b.dataset.nav);
    }
  });

  /* Tableau : sélection, tri par en-tête, clavier, redimensionnement */
  const sheet = $('#sheet');
  sheet.addEventListener('click', (e) => {
    const sortBtn = e.target.closest('[data-sort]');
    if (sortBtn) {
      const key = sortBtn.dataset.sort;
      if (S.sortBy === key) S.sortDir = S.sortDir === 'asc' ? 'desc' : 'asc';
      else { S.sortBy = key; S.sortDir = 'asc'; }
      saveSettings();
      renderToolbar();
      rebuild(false);
      return;
    }
    const td = e.target.closest('td[data-r]');
    if (td) selectCell(td);
  });
  sheet.addEventListener('keydown', (e) => {
    const k = e.key;
    if (k === 'ArrowDown') { e.preventDefault(); moveSel(1, 0); }
    else if (k === 'ArrowUp') { e.preventDefault(); moveSel(-1, 0); }
    else if (k === 'ArrowRight') { e.preventDefault(); moveSel(0, 1); }
    else if (k === 'ArrowLeft') { e.preventDefault(); moveSel(0, -1); }
    else if ((e.ctrlKey || e.metaKey) && (k === 'c' || k === 'C')) {
      const td = $('#sheet td.sel');
      if (td && navigator.clipboard) { e.preventDefault(); navigator.clipboard.writeText(td.textContent).then(() => toast('Cellule copiée.')).catch(() => {}); }
    }
  });
  sheet.addEventListener('pointerdown', (e) => {
    const h = e.target.closest('.rs');
    if (!h) return;
    e.preventDefault();
    const table = sheet.querySelector('table');
    const j = Number(h.dataset.j);
    const th = h.parentElement;
    const col = table.querySelectorAll('col')[j + 1];
    const startX = e.clientX;
    const startW = th.offsetWidth;
    const label = state.built.headers[j];
    h.setPointerCapture(e.pointerId);
    h.classList.add('active');
    const move = (ev) => {
      const w = clamp(startW + ev.clientX - startX, 50, 700);
      col.style.width = w + 'px';
      table.style.width = Array.from(table.querySelectorAll('col')).reduce((a, c) => a + parseFloat(c.style.width), 0) + 'px';
      state.colW[label] = w;
    };
    const up = () => {
      h.classList.remove('active');
      h.removeEventListener('pointermove', move);
      h.removeEventListener('pointerup', up);
      h.removeEventListener('pointercancel', up);
    };
    h.addEventListener('pointermove', move);
    h.addEventListener('pointerup', up);
    h.addEventListener('pointercancel', up);
  });

  /* Contrôle */
  $('#panelCheck').addEventListener('click', (e) => {
    const g = e.target.closest('[data-goto]');
    if (g) {
      const [r, c] = g.dataset.goto.split(':').map(Number);
      gotoCell(r, c);
      return;
    }
    const sp = e.target.closest('[data-split]');
    if (sp) {
      sp.dataset.split.split(',').filter(Boolean).forEach((i) => state.noMerge.add(Number(i)));
      toast('Lignes dissociées.');
      rebuild(true);
      return;
    }
    if (e.target.closest('[data-act="resetsplit"]')) {
      state.noMerge = new Set();
      rebuild(true);
    }
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

  /* Export */
  $('#dlBtn').addEventListener('click', () => doDownload());

  /* Réglages */
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
    if (el.id === 'presetSel') {
      if (el.value) applyPreset(el.value);
      renderSettings();
      return;
    }
    if (!el.dataset.k) return;
    setSetting(el.dataset.k, el.type === 'checkbox' ? el.checked : el.value);
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
    renderSettings();
    if (state.analysis) recompute();
    else renderAll();
    toast('Réglages réinitialisés.');
  });

  $$('dialog').forEach((d) => {
    d.addEventListener('click', (e) => { if (e.target === d || e.target.closest('[data-close]')) d.close(); });
  });

  /* Application */
  $('#getAppBtn').addEventListener('click', () => openAppDialog(true));
  window.addEventListener('beforeinstallprompt', (e) => e.preventDefault()); // on propose l'APK / l'EXE, pas l'installation par le navigateur

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
  setupAppButton();
  renderAll();

  if ('serviceWorker' in navigator && !isShell() && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    window.addEventListener('load', () => navigator.serviceWorker.register('/sw.js').catch(() => {}));
  }
  (window.requestIdleCallback || ((f) => setTimeout(f, 1200)))(() => loadLib('xlsx').catch(() => {}));

  window.MayaExcelApp = { addParsedFiles };
})();
