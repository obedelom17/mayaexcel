/*
 * MayaExcel – moteur de fusion
 * Aucune dépendance : fonctionne dans le navigateur et sous Node (tests).
 *
 * 1. analyze(files)  → lit des tableaux bruts (lignes/colonnes), repère la ligne d'en-têtes,
 *                      rapproche les colonnes qui veulent dire la même chose ("Tel", "Téléphone",
 *                      "Phone"…) et produit une liste d'enregistrements unifiés.
 * 2. build(analysis, réglages) → applique l'ordre des colonnes, les doublons, le tri, etc.
 *
 * Dates : les objets Date produits ici représentent l'heure "murale" lue dans Excel,
 * exprimée en UTC (on les lit et on les écrit toujours avec les getters/setters UTC).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MayaMerge = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ------------------------------------------------------------------ */
  /*  Normalisation des intitulés                                        */
  /* ------------------------------------------------------------------ */

  const STOP = new Set(['de', 'du', 'des', 'd', 'la', 'le', 'les', 'l', 'the', 'of', 'au', 'aux', 'en']);
  const NUM_WORDS = new Set(['n', 'no', 'num', 'numero', 'nbr']);

  function stripAccents(s) {
    return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  }

  function singular(w) {
    return w.length > 3 && w.endsWith('s') && !w.endsWith('ss') ? w.slice(0, -1) : w;
  }

  function tokens(text) {
    let t = stripAccents(text)
      .replace(/[’'`´]/g, ' ')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim()
      .split(' ')
      .filter(Boolean)
      .filter((w) => !STOP.has(w))
      .map(singular);
    // "Numéro de téléphone" → "telephone"
    if (t.length > 1) {
      const rest = t.filter((w) => !NUM_WORDS.has(w));
      if (rest.length) t = rest;
    }
    return t;
  }

  const keyOf = (t) => t.join(' ');
  const sortedKeyOf = (t) => t.slice().sort().join(' ');

  function levenshtein(a, b) {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    let prev = new Array(b.length + 1);
    for (let j = 0; j <= b.length; j++) prev[j] = j;
    for (let i = 1; i <= a.length; i++) {
      const cur = [i];
      for (let j = 1; j <= b.length; j++) {
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      }
      prev = cur;
    }
    return prev[b.length];
  }

  function similarity(a, b) {
    const max = Math.max(a.length, b.length);
    return max ? 1 - levenshtein(a, b) / max : 1;
  }

  const digitsOf = (s) => (s.match(/\d+/g) || []).join('-');

  /** Deux intitulés (déjà normalisés) sont-ils "quasi identiques" ? */
  function fuzzyMatch(a, b) {
    if (a.length < 5 || b.length < 5) return false;
    if (digitsOf(a) !== digitsOf(b)) return false; // "Adresse 1" ≠ "Adresse 2"
    return similarity(a, b) >= 0.86;
  }

  /* ------------------------------------------------------------------ */
  /*  Dictionnaire de synonymes (FR / EN) – les plus courants en RH      */
  /* ------------------------------------------------------------------ */

  const GROUPS = [
    ['matricule', 'Matricule', ['matricule', 'id', 'identifiant', 'id employe', 'employe id', 'employee id', 'code employe', 'employee number', 'numero employe', 'numero matricule', 'badge']],
    ['nom', 'Nom', ['nom', 'nom de famille', 'last name', 'lastname', 'surname', 'family name']],
    ['prenom', 'Prénom', ['prenom', 'prenoms', 'first name', 'firstname', 'given name']],
    ['nomcomplet', 'Nom complet', ['nom complet', 'nom et prenom', 'nom prenom', 'nom et prenoms', 'noms et prenoms', 'full name', 'fullname', 'name', 'employe', 'employee', 'agent', 'salarie', 'collaborateur', 'nom employe', 'employee name', 'nom agent']],
    ['sexe', 'Sexe', ['sexe', 'genre', 'gender', 'sex']],
    ['naissance', 'Date de naissance', ['date de naissance', 'date naissance', 'ne le', 'nee le', 'birth date', 'birthdate', 'date of birth', 'dob', 'naissance']],
    ['lieunaissance', 'Lieu de naissance', ['lieu de naissance', 'lieu naissance', 'place of birth', 'birthplace']],
    ['telephone', 'Téléphone', ['telephone', 'tel', 'tel mobile', 'telephone mobile', 'phone', 'phone number', 'mobile', 'portable', 'contact', 'gsm', 'cellulaire']],
    ['email', 'Email', ['email', 'e mail', 'mail', 'courriel', 'adresse email', 'adresse mail', 'adresse e mail', 'email address']],
    ['adresse', 'Adresse', ['adresse', 'address', 'adresse postale', 'domicile', 'residence', 'adresse domicile', 'lieu de residence']],
    ['ville', 'Ville', ['ville', 'city', 'commune', 'localite']],
    ['pays', 'Pays', ['pays', 'country']],
    ['nationalite', 'Nationalité', ['nationalite', 'nationality']],
    ['poste', 'Poste', ['poste', 'fonction', 'job title', 'emploi', 'intitule du poste', 'intitule poste', 'metier', 'profession']],
    ['service', 'Service', ['service', 'departement', 'department']],
    ['dateembauche', "Date d'embauche", ['date embauche', 'date d embauche', 'date d entree', 'date entree', 'hire date', 'start date', 'date de recrutement', 'date recrutement', 'date de prise de service', 'date prise de service']],
    ['salaire', 'Salaire', ['salaire', 'salary', 'remuneration', 'salaire de base', 'base salary']],
    ['statut', 'Statut', ['statut', 'status']],
    ['matrimonial', 'Situation matrimoniale', ['situation matrimoniale', 'etat civil', 'situation familiale', 'marital status']],
    ['enfants', "Nombre d'enfants", ['nombre d enfants', 'nombre enfants', 'nb enfants', 'enfants', 'children']],
    ['cnss', 'N° CNSS', ['cnss', 'numero cnss', 'securite sociale', 'numero securite sociale']],
    ['diplome', 'Diplôme', ['diplome', 'niveau d etudes', 'niveau etude', 'education']],
    ['contrat', 'Type de contrat', ['type de contrat', 'type contrat', 'contrat', 'contract type']],
    ['observation', 'Observations', ['observation', 'remarque', 'commentaire', 'notes', 'note', 'comment']],
  ];

  const GROUP_LABEL = {};
  const GROUP_KEYS = []; // [{key, group}] pour la recherche floue
  const GROUP_MAP = new Map();
  GROUPS.forEach(([id, label, names]) => {
    GROUP_LABEL[id] = label;
    names.forEach((n) => {
      const t = tokens(n);
      const k = keyOf(t);
      const sk = sortedKeyOf(t);
      if (!GROUP_MAP.has(k)) GROUP_MAP.set(k, id);
      if (!GROUP_MAP.has('~' + sk)) GROUP_MAP.set('~' + sk, id);
      GROUP_KEYS.push({ key: k, group: id });
    });
  });

  function resolveGroup(t) {
    if (!t.length) return null;
    const k = keyOf(t);
    const sk = sortedKeyOf(t);
    if (GROUP_MAP.has(k)) return GROUP_MAP.get(k);
    if (GROUP_MAP.has('~' + sk)) return GROUP_MAP.get('~' + sk);
    for (const g of GROUP_KEYS) if (fuzzyMatch(k, g.key)) return g.group;
    return null;
  }

  function prettify(text) {
    const s = String(text).replace(/\s+/g, ' ').trim();
    if (!s) return s;
    const hasLetters = s !== s.toLowerCase() || s !== s.toUpperCase();
    if (hasLetters && s === s.toUpperCase() && s.length > 4) return s.charAt(0) + s.slice(1).toLowerCase();
    if (hasLetters && s === s.toLowerCase()) return s.charAt(0).toUpperCase() + s.slice(1);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  Valeurs                                                            */
  /* ------------------------------------------------------------------ */

  const ERROR_RE = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NULL!|NUM!)/i;

  function isEmpty(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }

  function cleanValue(v) {
    if (isEmpty(v)) return null;
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      const r = new Date(Math.round(v.getTime() / 60000) * 60000); // arrondi à la minute
      return new Date(Date.UTC(r.getFullYear(), r.getMonth(), r.getDate(), r.getHours(), r.getMinutes()));
    }
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (typeof v === 'boolean') return v;
    const s = String(v).replace(/[\s\u00a0]+/g, ' ').trim();
    if (!s || ERROR_RE.test(s)) return null;
    return s;
  }

  function dateToText(d) {
    const p = (n) => String(n).padStart(2, '0');
    const base = p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
    return d.getUTCHours() || d.getUTCMinutes() ? base + ' ' + p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) : base;
  }

  function headerText(v) {
    if (v instanceof Date) return dateToText(cleanValue(v) || v);
    return String(v).replace(/[\s\u00a0]+/g, ' ').trim();
  }

  /* ------------------------------------------------------------------ */
  /*  Lecture d'une feuille : où sont les en-têtes ?                     */
  /* ------------------------------------------------------------------ */

  function readSheet(rows) {
    if (!rows || !rows.length) return null;
    const stat = rows.slice(0, 40).map((row) => {
      let filled = 0;
      let strings = 0;
      (row || []).forEach((c) => {
        if (isEmpty(c)) return;
        filled++;
        if (typeof c === 'string' && !/^[\d\s.,%€$-]+$/.test(c)) strings++;
      });
      return { filled, strings };
    });
    const maxFilled = Math.max(0, ...stat.map((s) => s.filled));
    if (!maxFilled) return null;

    const need = Math.min(2, maxFilled) > 1 ? Math.max(2, Math.ceil(maxFilled * 0.6)) : 1;
    let headerRow = stat.findIndex((s, i) => i < 25 && s.filled >= need && s.strings / s.filled >= 0.6);
    if (headerRow < 0) headerRow = stat.findIndex((s) => s.filled > 0);
    if (headerRow < 0) return null;

    // Une feuille sans aucune ligne de données sous l'en-tête est ignorée.
    let hasData = false;
    for (let r = headerRow + 1; r < rows.length && !hasData; r++) {
      hasData = (rows[r] || []).some((c) => !isEmpty(c));
    }
    if (!hasData) return null;

    const head = rows[headerRow] || [];
    let width = head.length;
    for (let r = headerRow + 1; r < rows.length; r++) width = Math.max(width, (rows[r] || []).length);

    const headers = [];
    for (let j = 0; j < width; j++) {
      let text = isEmpty(head[j]) ? '' : headerText(head[j]);
      if (!text) {
        let hasData = false;
        for (let r = headerRow + 1; r < rows.length && !hasData; r++) hasData = !isEmpty((rows[r] || [])[j]);
        if (!hasData) continue;
        text = 'Colonne ' + (j + 1);
      }
      headers.push({ idx: j, text });
    }
    return headers.length ? { headerRow, headers } : null;
  }

  /* ------------------------------------------------------------------ */
  /*  Analyse                                                            */
  /* ------------------------------------------------------------------ */

  function baseName(name) {
    return String(name).replace(/\.[^.]+$/, '');
  }

  /**
   * @param files  [{id?, name, error?, sheets:[{name, rows:any[][]}]}]
   * @param options {allSheets:boolean}
   */
  function analyze(files, options) {
    const opts = Object.assign({ allSheets: true }, options || {});
    const cols = [];
    const byId = new Map();
    const order = [];
    const records = [];
    const report = [];
    let seq = 0;

    function createColumn(text, t, group, used) {
      let key = group || keyOf(t) || stripAccents(text);
      let label = group ? GROUP_LABEL[group] : prettify(text);
      // Deux colonnes du même fichier qui portent le même nom : on les garde séparées.
      const clash = cols.filter((c) => c.baseKey === key).length;
      const baseKey = key;
      if (clash) {
        key = key + '#' + (clash + 1);
        label = label + ' (' + (clash + 1) + ')';
      }
      const col = {
        id: 'c' + ++seq,
        key,
        baseKey,
        label,
        group: group || null,
        norm: keyOf(t),
        sorted: sortedKeyOf(t),
        variants: new Set(),
        fileIds: new Set(),
        filled: 0,
      };
      cols.push(col);
      byId.set(col.id, col);
      return col;
    }

    function matchColumn(text, used) {
      const t = tokens(text);
      const group = resolveGroup(t);
      const k = keyOf(t);
      const sk = sortedKeyOf(t);
      let found = null;
      if (group) found = cols.find((c) => c.group === group && !used.has(c.id));
      if (!found && k) {
        found = cols.find((c) => !used.has(c.id) && !c.group && (c.norm === k || c.sorted === sk));
      }
      if (!found && k) {
        found = cols.find((c) => !used.has(c.id) && !c.group && fuzzyMatch(c.norm, k));
      }
      return found || { create: true, text, t, group };
    }

    files.forEach((file, fi) => {
      const fileId = file.id != null ? file.id : 'f' + fi;
      const fr = { id: fileId, name: file.name, sheets: [], rows: 0, error: file.error || null };
      report.push(fr);
      if (file.error) return;

      const sheets = opts.allSheets ? file.sheets : file.sheets.slice(0, 1);
      const parsed = sheets.map((sh) => ({ sh, info: readSheet(sh.rows) }));
      const usable = parsed.filter((p) => p.info);
      parsed.forEach((p) => {
        if (!p.info) fr.sheets.push({ name: p.sh.name, rows: 0, cols: 0, empty: true });
      });

      usable.forEach(({ sh, info }) => {
        const used = new Set();
        let prev = null;
        const map = [];

        info.headers.forEach(({ idx, text }) => {
          let col = matchColumn(text, used);
          if (col.create) {
            col = createColumn(col.text, col.t, col.group, used);
            const at = prev ? order.indexOf(prev) + 1 : 0;
            order.splice(at, 0, col.id);
          }
          used.add(col.id);
          col.variants.add(prettify(text));
          col.fileIds.add(fileId);
          map.push({ idx, col });
          prev = col.id;
        });

        const headKeys = map.map((m) => keyOf(tokens(info.headers.find((h) => h.idx === m.idx).text)));
        const label = usable.length > 1 ? baseName(file.name) + ' – ' + sh.name : baseName(file.name);
        let count = 0;

        for (let r = info.headerRow + 1; r < sh.rows.length; r++) {
          const row = sh.rows[r];
          if (!row) continue;
          const vals = {};
          let filled = 0;
          let repeatedHeader = 0;
          map.forEach(({ idx, col }, mi) => {
            const v = cleanValue(row[idx]);
            if (v === null) return;
            vals[col.id] = v;
            filled++;
            if (typeof v === 'string' && keyOf(tokens(v)) === headKeys[mi]) repeatedHeader++;
          });
          if (!filled) continue;
          if (map.length >= 2 && repeatedHeader >= Math.max(2, Math.ceil(map.length * 0.6))) continue;
          if (filled === 1 && map.length >= 3) {
            const only = Object.values(vals)[0];
            if (typeof only === 'string' && /^(total|sous[- ]?total|signature|fait [aà])/i.test(only)) continue;
          }
          records.push({ fileId, file: baseName(file.name), source: label, vals });
          count++;
        }
        fr.sheets.push({ name: sh.name, rows: count, cols: map.length, headerRow: info.headerRow + 1 });
        fr.rows += count;
      });
    });

    records.forEach((r) => {
      Object.keys(r.vals).forEach((id) => {
        byId.get(id).filled++;
      });
    });

    const columns = order.map((id) => {
      const c = byId.get(id);
      return {
        id: c.id,
        key: c.key,
        label: c.label,
        group: c.group,
        variants: Array.from(c.variants),
        fileCount: c.fileIds.size,
        filled: c.filled,
      };
    });

    return {
      columns,
      records,
      report,
      fileCount: files.filter((f) => !f.error).length,
      mergedColumns: columns.filter((c) => c.variants.length > 1).length,
    };
  }

  /* ------------------------------------------------------------------ */
  /*  Construction du tableau final                                      */
  /* ------------------------------------------------------------------ */

  const collator = typeof Intl !== 'undefined' ? new Intl.Collator('fr', { numeric: true, sensitivity: 'base' }) : null;

  function compare(a, b) {
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (a instanceof Date && b instanceof Date) return a - b;
    const sa = a instanceof Date ? a.toISOString() : String(a);
    const sb = b instanceof Date ? b.toISOString() : String(b);
    return collator ? collator.compare(sa, sb) : sa.localeCompare(sb);
  }

  function signature(v) {
    if (v === undefined || v === null) return '';
    if (v instanceof Date) return 'd' + v.getTime();
    return String(v).toLowerCase();
  }

  /**
   * @param analysis résultat de analyze()
   * @param st {cols:[{id,label,include}], source, numbering, dedupe, sortBy:'auto'|'none'|id, sortDir:'asc'|'desc'}
   */
  function build(analysis, st) {
    const meta = new Map(analysis.columns.map((c) => [c.id, c]));
    const cols = st.cols
      .filter((c) => c.include && meta.has(c.id))
      .map((c) => ({ id: c.id, label: c.label, group: meta.get(c.id).group, variants: meta.get(c.id).variants }));

    let recs = analysis.records.filter((r) => cols.some((c) => r.vals[c.id] !== undefined));

    let removed = 0;
    if (st.dedupe) {
      const seen = new Set();
      recs = recs.filter((r) => {
        const k = cols.map((c) => signature(r.vals[c.id])).join('\u0001');
        if (seen.has(k)) {
          removed++;
          return false;
        }
        seen.add(k);
        return true;
      });
    }

    let keys = [];
    if (st.sortBy === 'auto') {
      const g = (id) => cols.find((c) => c.group === id);
      if (g('nom')) keys = [g('nom'), g('prenom')].filter(Boolean);
      else if (g('nomcomplet')) keys = [g('nomcomplet')];
      else if (g('matricule')) keys = [g('matricule')];
    } else if (st.sortBy && st.sortBy !== 'none') {
      const c = cols.find((x) => x.id === st.sortBy);
      if (c) keys = [c];
    }

    if (keys.length) {
      const dir = st.sortDir === 'desc' ? -1 : 1;
      recs = recs.slice().sort((x, y) => {
        for (const k of keys) {
          const a = x.vals[k.id];
          const b = y.vals[k.id];
          const an = a === undefined;
          const bn = b === undefined;
          if (an || bn) {
            if (an && bn) continue;
            return an ? 1 : -1; // les cases vides restent toujours en bas
          }
          const c = compare(a, b);
          if (c) return c * dir;
        }
        return 0;
      });
    }

    const headers = [];
    if (st.numbering) headers.push('N°');
    cols.forEach((c) => headers.push(c.label || 'Colonne'));
    if (st.source) headers.push('Fichier source');

    const rows = recs.map((r, i) => {
      const out = [];
      if (st.numbering) out.push(i + 1);
      cols.forEach((c) => out.push(r.vals[c.id] === undefined ? null : r.vals[c.id]));
      if (st.source) out.push(r.source);
      return out;
    });

    return {
      headers,
      rows,
      removed,
      sortedBy: keys.map((k) => k.label),
      columnCount: cols.length,
      mergedFlags: (st.numbering ? [false] : [])
        .concat(cols.map((c) => c.variants.length > 1))
        .concat(st.source ? [false] : []),
      variantsList: (st.numbering ? [[]] : [])
        .concat(cols.map((c) => c.variants))
        .concat(st.source ? [[]] : []),
    };
  }

  return { analyze, build, dateToText, tokens, resolveGroup, readSheet };
});
