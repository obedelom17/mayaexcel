/* MayaExcel – interface */
(function () {
  'use strict';

  const $ = (sel) => document.querySelector(sel);
  const esc = (s) =>
    String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ------------------------------------------------------------------ */
  /*  Chargement des bibliothèques (SheetJS pour lire, ExcelJS pour écrire) */
  /* ------------------------------------------------------------------ */

  const LIBS = {
    xlsx: [
      'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
      'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    ],
    exceljs: [
      'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js',
      'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
    ],
  };
  const CHECK = { xlsx: () => window.XLSX, exceljs: () => window.ExcelJS };
  const pending = {};

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.onload = resolve;
      s.onerror = () => {
        s.remove();
        reject(new Error('Chargement impossible : ' + src));
      };
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
          } catch (e) {
            /* on essaie la source suivante */
          }
        }
        throw new Error('lib:' + name);
      })().catch((e) => {
        pending[name] = null;
        throw e;
      });
    }
    return pending[name];
  }

  /* ------------------------------------------------------------------ */
  /*  État                                                               */
  /* ------------------------------------------------------------------ */

  const state = {
    files: [], // {id,name,size,status:'reading'|'ok'|'error',error,sheets}
    allSheets: true,
    opts: { dedupe: true, numbering: true, source: false, sortBy: 'auto', sortDir: 'asc' },
    cols: [], // [{id,key,label,include}]
    overrides: {}, // key → {label, include}  (mémorise les choix de l'utilisateur)
    analysis: null,
    built: null,
  };
  let uid = 0;
  const PREVIEW_ROWS = 100;

  /* ------------------------------------------------------------------ */
  /*  Utilitaires d'affichage                                            */
  /* ------------------------------------------------------------------ */

  let toastTimer;
  function toast(msg, isError) {
    const el = $('#toast');
    el.textContent = msg;
    el.classList.toggle('is-error', !!isError);
    el.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove('show'), isError ? 5000 : 2800);
  }

  const plural = (n, one, many) => n + ' ' + (n > 1 ? many : one);

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

  function formatCell(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return MayaMerge.dateToText(v);
    if (typeof v === 'boolean') return v ? 'Oui' : 'Non';
    return String(v);
  }

  /* ------------------------------------------------------------------ */
  /*  Lecture des fichiers                                               */
  /* ------------------------------------------------------------------ */

  const EXT_OK = /\.(xlsx|xlsm|xls|csv|tsv|ods)$/i;

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
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
      if (rows.length) sheets.push({ name, rows });
    });
    return sheets;
  }

  const tick = () => new Promise((r) => setTimeout(r, 0));

  async function addFiles(fileList) {
    const incoming = Array.from(fileList || []);
    if (!incoming.length) return;

    const fresh = [];
    let skipped = 0;
    let unsupported = 0;
    incoming.forEach((f) => {
      if (!EXT_OK.test(f.name)) return void unsupported++;
      if (state.files.some((x) => x.name === f.name && x.size === f.size)) return void skipped++;
      const entry = { id: ++uid, name: f.name, size: f.size, status: 'reading', error: null, sheets: [], file: f };
      state.files.push(entry);
      fresh.push(entry);
    });
    if (unsupported) toast(plural(unsupported, 'fichier ignoré', 'fichiers ignorés') + ' : seuls les fichiers Excel et CSV sont acceptés.', true);
    else if (skipped) toast(plural(skipped, 'fichier déjà ajouté', 'fichiers déjà ajoutés') + '.');
    if (!fresh.length) return;

    renderFiles();
    await tick();

    for (const entry of fresh) {
      try {
        entry.sheets = await readWorkbook(entry.file);
        if (!entry.sheets.length) throw new Error('empty');
        entry.status = 'ok';
      } catch (err) {
        entry.status = 'error';
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
    const hadResult = !$('#resultBlock').hidden;
    recompute();
    if (!hadResult && state.analysis && state.built && state.built.rows.length) {
      setTimeout(() => $('#resultBlock').scrollIntoView({ behavior: 'smooth', block: 'start' }), 80);
    }
  }

  /** Point d'entrée pour injecter des fichiers déjà lus (tests). */
  function addParsedFiles(list) {
    list.forEach((f) => state.files.push({ id: ++uid, name: f.name, size: f.size || 0, status: 'ok', error: null, sheets: f.sheets }));
    renderFiles();
    recompute();
  }

  /* ------------------------------------------------------------------ */
  /*  Fusion                                                             */
  /* ------------------------------------------------------------------ */

  function recompute() {
    const ok = state.files.filter((f) => f.status === 'ok');
    if (!ok.length) {
      state.analysis = null;
      state.built = null;
      state.cols = [];
      renderResult();
      return;
    }
    state.analysis = MayaMerge.analyze(
      ok.map((f) => ({ id: f.id, name: f.name, sheets: f.sheets })),
      { allSheets: state.allSheets }
    );
    state.cols = state.analysis.columns.map((c) => {
      const o = state.overrides[c.key] || {};
      return { id: c.id, key: c.key, label: o.label != null ? o.label : c.label, include: o.include != null ? o.include : true };
    });
    renderFiles(); // met à jour le détail de chaque fichier
    rebuild(true);
  }

  function currentSettings() {
    return Object.assign({ cols: state.cols }, { source: state.opts.source, numbering: state.opts.numbering, dedupe: state.opts.dedupe, sortBy: state.opts.sortBy, sortDir: state.opts.sortDir });
  }

  function rebuild(full) {
    state.built = MayaMerge.build(state.analysis, currentSettings());
    if (full) renderResult();
    else {
      renderSummary();
      renderPreview();
      updateActions();
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Rendu : fichiers                                                   */
  /* ------------------------------------------------------------------ */

  function renderFiles() {
    const block = $('#filesBlock');
    block.hidden = !state.files.length;
    if (!state.files.length) return;
    $('#filesCount').textContent = '(' + state.files.length + ')';

    const reports = new Map();
    if (state.analysis) state.analysis.report.forEach((r) => reports.set(r.id, r));

    $('#filesList').innerHTML = state.files
      .map((f) => {
        const ext = (f.name.split('.').pop() || '').toLowerCase().slice(0, 4);
        let meta;
        if (f.status === 'reading') meta = '<span class="spin"></span>Lecture en cours…';
        else if (f.status === 'error') meta = esc(f.error);
        else {
          const r = reports.get(f.id);
          if (r) {
            const used = r.sheets.filter((s) => !s.empty);
            meta = plural(r.rows, 'ligne', 'lignes') + ' · ' + plural(used.length, 'feuille', 'feuilles');
            if (!r.rows) meta = 'Aucune donnée trouvée';
          } else meta = formatSize(f.size);
        }
        return (
          '<li class="file' + (f.status === 'error' ? ' is-error' : '') + '">' +
          '<span class="file-ext" aria-hidden="true">' + esc(ext) + '</span>' +
          '<div class="file-body"><div class="file-name" title="' + esc(f.name) + '">' + esc(f.name) + '</div>' +
          '<div class="file-meta">' + meta + '</div></div>' +
          '<button type="button" class="icon-btn" data-remove="' + f.id + '" aria-label="Retirer ' + esc(f.name) + '">×</button>' +
          '</li>'
        );
      })
      .join('');
  }

  /* ------------------------------------------------------------------ */
  /*  Rendu : résultat                                                   */
  /* ------------------------------------------------------------------ */

  function renderResult() {
    const block = $('#resultBlock');
    block.hidden = !state.analysis;
    if (!state.analysis) return;

    $('#optDedupe').checked = state.opts.dedupe;
    $('#optNumber').checked = state.opts.numbering;
    $('#optSource').checked = state.opts.source;
    $('#optSheets').checked = state.allSheets;

    renderSortOptions();
    renderColumns();
    renderSummary();
    renderPreview();
    updateActions();
  }

  function renderSortOptions() {
    const sel = $('#sortBy');
    const autoLabel = autoSortLabel();
    let html = '<option value="auto">' + esc(autoLabel ? 'Automatique (' + autoLabel + ')' : 'Automatique') + '</option>';
    html += '<option value="none">Ordre des fichiers</option>';
    state.cols.filter((c) => c.include).forEach((c) => {
      html += '<option value="' + c.id + '">' + esc(c.label || 'Colonne') + '</option>';
    });
    sel.innerHTML = html;
    if (![...sel.options].some((o) => o.value === state.opts.sortBy)) state.opts.sortBy = 'auto';
    sel.value = state.opts.sortBy;
    const dir = $('#sortDir');
    dir.textContent = state.opts.sortDir === 'asc' ? 'A → Z' : 'Z → A';
    dir.disabled = state.opts.sortBy === 'none';
  }

  function autoSortLabel() {
    if (!state.analysis) return '';
    const cols = state.cols.filter((c) => c.include);
    const meta = new Map(state.analysis.columns.map((c) => [c.id, c]));
    const has = (g) => cols.find((c) => meta.get(c.id).group === g);
    if (has('nom')) return has('prenom') ? 'Nom, Prénom' : 'Nom';
    if (has('nomcomplet')) return 'Nom complet';
    if (has('matricule')) return 'Matricule';
    return '';
  }

  function renderSummary() {
    const a = state.analysis;
    const b = state.built;
    if (!a || !b) return;
    const files = a.fileCount;
    const notes = [];
    if (a.mergedColumns) notes.push(plural(a.mergedColumns, 'colonne regroupée', 'colonnes regroupées') + ' automatiquement');
    if (b.removed) notes.push(plural(b.removed, 'doublon retiré', 'doublons retirés'));
    if (b.sortedBy.length) notes.push('trié par ' + b.sortedBy.join(', '));
    $('#summary').innerHTML =
      '<strong>' + plural(files, 'fichier', 'fichiers') + '</strong> fusionnés en un tableau de <strong>' +
      plural(b.rows.length, 'ligne', 'lignes') + '</strong> et <strong>' + plural(b.columnCount, 'colonne', 'colonnes') + '</strong>.' +
      (notes.length ? '<small>' + esc(notes.join(' · ')) + '</small>' : '');
  }

  function renderColumns() {
    const a = state.analysis;
    const meta = new Map(a.columns.map((c) => [c.id, c]));
    const total = a.fileCount;
    const shown = state.cols.filter((c) => c.include).length;
    $('#colsSummary').textContent = 'Colonnes : ' + shown + ' sur ' + state.cols.length + ' affichées';

    $('#colsList').innerHTML = state.cols
      .map((c, i) => {
        const m = meta.get(c.id);
        const info =
          m.variants.length > 1
            ? '<span class="dot"></span>Regroupe : ' + esc(m.variants.join(', '))
            : 'Présente dans ' + m.fileCount + ' fichier' + (m.fileCount > 1 ? 's' : '') + ' sur ' + total;
        return (
          '<li class="col-row' + (c.include ? '' : ' is-off') + '" data-id="' + c.id + '">' +
          '<input type="checkbox" data-act="include" ' + (c.include ? 'checked' : '') + ' aria-label="Afficher la colonne ' + esc(c.label) + '">' +
          '<input type="text" class="col-label" data-act="label" value="' + esc(c.label) + '" aria-label="Nom de la colonne" autocomplete="off">' +
          '<div class="col-move">' +
          '<button type="button" class="icon-btn" data-act="up" aria-label="Monter" ' + (i === 0 ? 'disabled' : '') + '>↑</button>' +
          '<button type="button" class="icon-btn" data-act="down" aria-label="Descendre" ' + (i === state.cols.length - 1 ? 'disabled' : '') + '>↓</button>' +
          '</div>' +
          '<div class="col-info">' + info + '</div>' +
          '</li>'
        );
      })
      .join('');
  }

  function renderPreview() {
    const b = state.built;
    const box = $('#sheet');
    if (!b || !b.rows.length) {
      box.innerHTML = '<p class="sheet-note" style="padding:1.5rem;text-align:center">Aucune ligne à afficher. Vérifiez que des colonnes sont affichées.</p>';
      $('#sheetNote').textContent = '';
      return;
    }
    const rows = b.rows.slice(0, PREVIEW_ROWS);

    let html = '<table class="grid"><thead><tr class="letters"><th class="rn"></th>';
    b.headers.forEach((_, j) => (html += '<th scope="col">' + colLetter(j) + '</th>'));
    html += '</tr><tr class="heads"><th class="rn"></th>';
    b.headers.forEach((h, j) => {
      const v = b.variantsList[j];
      const dot = b.mergedFlags[j] ? '<span class="merged-dot" title="Regroupe : ' + esc(v.join(', ')) + '"></span>' : '';
      html += '<th scope="col">' + esc(h) + dot + '</th>';
    });
    html += '</tr></thead><tbody>';
    rows.forEach((r, i) => {
      html += '<tr><td class="rn">' + (i + 1) + '</td>';
      r.forEach((v, j) => {
        const t = formatCell(v);
        html += '<td' + (typeof v === 'number' ? ' class="num"' : '') + ' title="' + esc(t) + '">' + esc(t) + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    box.innerHTML = html;

    $('#sheetNote').textContent =
      b.rows.length > PREVIEW_ROWS
        ? 'Aperçu des ' + PREVIEW_ROWS + ' premières lignes sur ' + b.rows.length + '. Le fichier téléchargé contient tout.'
        : 'Aperçu complet. Faites défiler pour voir toutes les colonnes.';
  }

  function updateActions() {
    const ok = !!(state.built && state.built.rows.length);
    $('#dlXlsx').disabled = !ok;
    $('#dlCsv').disabled = !ok;
  }

  /* ------------------------------------------------------------------ */
  /*  Export                                                             */
  /* ------------------------------------------------------------------ */

  function stamp() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }

  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  function summaryRows() {
    const rows = [['Fichier', 'Feuilles lues', 'Lignes']];
    state.analysis.report.forEach((r) => {
      rows.push([r.name, r.error ? 'Erreur' : r.sheets.filter((s) => !s.empty).length, r.rows]);
    });
    return rows;
  }

  async function exportXlsx() {
    const b = state.built;
    const name = 'MayaExcel-fusion-' + stamp() + '.xlsx';
    const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const btn = $('#dlXlsx');
    const label = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Préparation…';
    try {
      try {
        await loadLib('exceljs');
        const wb = new window.ExcelJS.Workbook();
        wb.creator = 'MayaExcel';
        wb.created = new Date();
        const ws = wb.addWorksheet('Fusion', { views: [{ state: 'frozen', ySplit: 1 }] });

        ws.addRow(b.headers);
        b.rows.forEach((r) => ws.addRow(r.map((v) => (v === null ? null : v))));

        const border = { style: 'thin', color: { argb: 'FFD5DFD9' } };
        const head = ws.getRow(1);
        head.height = 24;
        head.eachCell((cell) => {
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A7F53' } };
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
          cell.border = { bottom: border };
        });

        b.headers.forEach((h, j) => {
          let width = String(h).length;
          const col = ws.getColumn(j + 1);
          const sample = Math.min(b.rows.length, 2000);
          let hasDate = false;
          for (let i = 0; i < sample; i++) {
            const v = b.rows[i][j];
            if (v instanceof Date) {
              hasDate = true;
              width = Math.max(width, 16);
            } else if (v !== null) width = Math.max(width, String(v).length);
          }
          col.width = Math.min(Math.max(width + 3, 10), 48);
          if (hasDate) {
            col.eachCell({ includeEmpty: false }, (cell, rowNumber) => {
              if (rowNumber > 1 && cell.value instanceof Date) {
                const v = cell.value;
                cell.numFmt = v.getUTCHours() || v.getUTCMinutes() ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy';
                cell.alignment = { horizontal: 'left' };
              }
            });
          }
        });

        for (let i = 2; i <= b.rows.length + 1; i++) {
          const row = ws.getRow(i);
          const zebra = i % 2 === 1;
          for (let j = 1; j <= b.headers.length; j++) {
            const cell = row.getCell(j);
            cell.border = { bottom: border };
            if (zebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F8F5' } };
          }
        }
        ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: b.headers.length } };

        const ws2 = wb.addWorksheet('Récapitulatif');
        summaryRows().forEach((r) => ws2.addRow(r));
        ws2.getRow(1).font = { bold: true };
        ws2.getColumn(1).width = 44;
        ws2.getColumn(2).width = 16;
        ws2.getColumn(3).width = 12;

        const buf = await wb.xlsx.writeBuffer();
        download(new Blob([buf], { type: mime }), name);
      } catch (styledErr) {
        // Plan B : fichier simple avec SheetJS
        await loadLib('xlsx');
        const XLSX = window.XLSX;
        const data = [b.headers].concat(b.rows);
        const ws = XLSX.utils.aoa_to_sheet(data, { cellDates: true });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Fusion');
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellDates: true });
        download(new Blob([out], { type: mime }), name);
      }
      toast('Fichier Excel prêt.');
    } catch (e) {
      toast('Impossible de créer le fichier. Vérifiez votre connexion Internet et réessayez.', true);
    } finally {
      btn.textContent = label;
      updateActions();
    }
  }

  function exportCsv() {
    const b = state.built;
    const q = (v) => {
      const s = formatCell(v);
      return /[";\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const lines = [b.headers.map(q).join(';')].concat(b.rows.map((r) => r.map(q).join(';')));
    // BOM + point-virgule : s'ouvre correctement dans Excel en français
    download(new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), 'MayaExcel-fusion-' + stamp() + '.csv');
    toast('Fichier CSV prêt.');
  }

  /* ------------------------------------------------------------------ */
  /*  Événements                                                         */
  /* ------------------------------------------------------------------ */

  const fileInput = $('#fileInput');
  const drop = $('#drop');

  function openPicker() {
    fileInput.click();
  }
  $('#pickBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    openPicker();
  });
  $('#addMoreBtn').addEventListener('click', openPicker);
  drop.addEventListener('click', openPicker);
  drop.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      openPicker();
    }
  });
  fileInput.addEventListener('change', () => {
    const list = Array.from(fileInput.files);
    fileInput.value = '';
    addFiles(list);
  });

  ['dragenter', 'dragover'].forEach((ev) =>
    window.addEventListener(ev, (e) => {
      if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return;
      e.preventDefault();
      drop.classList.add('is-over');
    })
  );
  ['dragleave', 'drop'].forEach((ev) =>
    window.addEventListener(ev, (e) => {
      if (ev === 'dragleave' && e.relatedTarget) return;
      drop.classList.remove('is-over');
    })
  );
  window.addEventListener('drop', (e) => {
    if (!e.dataTransfer || !e.dataTransfer.files.length) return;
    e.preventDefault();
    addFiles(e.dataTransfer.files);
  });

  $('#clearBtn').addEventListener('click', () => {
    state.files = [];
    state.overrides = {};
    state.opts.sortBy = 'auto';
    recompute();
    renderFiles();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  $('#filesList').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-remove]');
    if (!btn) return;
    const id = Number(btn.dataset.remove);
    state.files = state.files.filter((f) => f.id !== id);
    recompute();
    renderFiles();
  });

  $('#sortBy').addEventListener('change', (e) => {
    state.opts.sortBy = e.target.value;
    $('#sortDir').disabled = state.opts.sortBy === 'none';
    rebuild(false);
  });
  $('#sortDir').addEventListener('click', () => {
    state.opts.sortDir = state.opts.sortDir === 'asc' ? 'desc' : 'asc';
    $('#sortDir').textContent = state.opts.sortDir === 'asc' ? 'A → Z' : 'Z → A';
    rebuild(false);
  });
  $('#optDedupe').addEventListener('change', (e) => {
    state.opts.dedupe = e.target.checked;
    rebuild(false);
  });
  $('#optNumber').addEventListener('change', (e) => {
    state.opts.numbering = e.target.checked;
    rebuild(false);
  });
  $('#optSource').addEventListener('change', (e) => {
    state.opts.source = e.target.checked;
    rebuild(false);
  });
  $('#optSheets').addEventListener('change', (e) => {
    state.allSheets = e.target.checked;
    recompute();
  });

  function rememberCol(c) {
    state.overrides[c.key] = { label: c.label, include: c.include };
  }

  const colsList = $('#colsList');
  colsList.addEventListener('input', (e) => {
    if (e.target.dataset.act !== 'label') return;
    const id = e.target.closest('.col-row').dataset.id;
    const c = state.cols.find((x) => x.id === id);
    c.label = e.target.value;
    rememberCol(c);
    rebuild(false);
  });
  colsList.addEventListener('change', (e) => {
    if (e.target.dataset.act === 'include') {
      const row = e.target.closest('.col-row');
      const c = state.cols.find((x) => x.id === row.dataset.id);
      c.include = e.target.checked;
      rememberCol(c);
      row.classList.toggle('is-off', !c.include);
      renderSortOptions();
      const shown = state.cols.filter((x) => x.include).length;
      $('#colsSummary').textContent = 'Colonnes : ' + shown + ' sur ' + state.cols.length + ' affichées';
      rebuild(false);
    } else if (e.target.dataset.act === 'label') {
      renderSortOptions();
    }
  });
  colsList.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act="up"],[data-act="down"]');
    if (!btn) return;
    const id = btn.closest('.col-row').dataset.id;
    const i = state.cols.findIndex((x) => x.id === id);
    const j = btn.dataset.act === 'up' ? i - 1 : i + 1;
    if (j < 0 || j >= state.cols.length) return;
    [state.cols[i], state.cols[j]] = [state.cols[j], state.cols[i]];
    renderColumns();
    rebuild(false);
    const again = colsList.querySelector('[data-id="' + id + '"] [data-act="' + btn.dataset.act + '"]');
    if (again && !again.disabled) again.focus();
  });

  $('#dlXlsx').addEventListener('click', exportXlsx);
  $('#dlCsv').addEventListener('click', exportCsv);

  /* ------------------------------------------------------------------ */
  /*  Installation comme application (PWA)                               */
  /* ------------------------------------------------------------------ */

  let deferredPrompt = null;
  const installBtn = $('#installBtn');
  const strip = $('#installStrip');
  const DISMISS_KEY = 'maya-install-dismissed';

  const isStandalone = () =>
    window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  const isIOS = () =>
    /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

  function dismissed() {
    try {
      return localStorage.getItem(DISMISS_KEY) === '1';
    } catch (e) {
      return false;
    }
  }

  function showInstallUi() {
    if (isStandalone()) return;
    installBtn.hidden = false;
    if (!dismissed()) strip.hidden = false;
  }

  async function install() {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      try {
        await deferredPrompt.userChoice;
      } catch (e) {
        /* ignoré */
      }
      deferredPrompt = null;
      installBtn.hidden = true;
      strip.hidden = true;
    } else if (isIOS()) {
      $('#iosDialog').showModal();
    }
  }

  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    showInstallUi();
  });
  window.addEventListener('appinstalled', () => {
    deferredPrompt = null;
    installBtn.hidden = true;
    strip.hidden = true;
    toast('MayaExcel est installée sur votre appareil.');
  });
  installBtn.addEventListener('click', install);
  $('#installStripBtn').addEventListener('click', install);
  $('#installDismiss').addEventListener('click', () => {
    strip.hidden = true;
    try {
      localStorage.setItem(DISMISS_KEY, '1');
    } catch (e) {
      /* ignoré */
    }
  });

  if (isIOS() && !isStandalone()) showInstallUi();

  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    window.addEventListener('load', () => navigator.serviceWorker.register('/sw.js').catch(() => {}));
  }

  // Charge le moteur Excel en arrière-plan pour que le premier fichier soit lu sans attente.
  (window.requestIdleCallback || ((f) => setTimeout(f, 1200)))(() => loadLib('xlsx').catch(() => {}));

  window.MayaExcelApp = { addParsedFiles };
})();
