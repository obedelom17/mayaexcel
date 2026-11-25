/*
 * MayaExcel – moteur de fusion (v2)
 * Aucune dépendance : fonctionne dans le navigateur et sous Node (tests).
 *
 * analyze(files)  → lit les tableaux bruts, repère les en-têtes (ou leur absence), rapproche les
 *                   colonnes par leur intitulé ET par leur contenu, nettoie les types (dates, nombres),
 *                   sépare "Nom et Prénoms" quand d'autres fichiers ont Nom / Prénom à part.
 * build(analysis) → applique les réglages : rattachements manuels, harmonisation, regroupement des
 *                   lignes d'une même personne, tri, numérotation.
 *
 * Dates : les objets Date représentent l'heure "murale" lue dans Excel, exprimée en UTC
 * (toujours lus et écrits avec les accesseurs UTC).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MayaMerge = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ================================================================== */
  /*  Intitulés                                                          */
  /* ================================================================== */

  const STOP = new Set(['de', 'du', 'des', 'd', 'la', 'le', 'les', 'l', 'the', 'of', 'au', 'aux', 'en']);
  const NUM_WORDS = new Set(['n', 'no', 'num', 'numero', 'nbr']);

  const stripAccents = (s) => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const singular = (w) => (w.length > 3 && w.endsWith('s') && !w.endsWith('ss') ? w.slice(0, -1) : w);

  function tokens(text) {
    let t = stripAccents(text)
      .replace(/[’'`´]/g, ' ')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim()
      .split(' ')
      .filter(Boolean)
      .filter((w) => !STOP.has(w))
      .map(singular);
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

  const similarity = (a, b) => {
    const max = Math.max(a.length, b.length);
    return max ? 1 - levenshtein(a, b) / max : 1;
  };

  const digitsOf = (s) => (s.match(/\d+/g) || []).join('-');

  function fuzzyMatch(a, b, th) {
    if (a.length < 5 || b.length < 5) return false;
    if (digitsOf(a) !== digitsOf(b)) return false; // "Adresse 1" ≠ "Adresse 2"
    return similarity(a, b) >= (th || 0.86);
  }

  function jaccardTokens(ta, tb) {
    const A = new Set(ta);
    const B = new Set(tb);
    let inter = 0;
    A.forEach((x) => B.has(x) && inter++);
    const union = A.size + B.size - inter;
    return { score: union ? inter / union : 0, union };
  }

  const GENERIC_RE = /^(colonne|column|col|champ|field|unnamed|sans titre|untitled)[\s:_-]*\d*$/;
  const isGenericHeader = (text) => GENERIC_RE.test(stripAccents(text).trim());

  /* ================================================================== */
  /*  Dictionnaire de synonymes (FR / EN)                                */
  /* ================================================================== */

  const GROUPS = [
    ['matricule', 'Matricule', ['matricule', 'id', 'identifiant', 'id employe', 'employe id', 'employee id', 'code employe', 'employee number', 'numero employe', 'numero matricule', 'badge']],
    ['nom', 'Nom', ['nom', 'nom de famille', 'last name', 'lastname', 'surname', 'family name']],
    ['prenom', 'Prénom', ['prenom', 'prenoms', 'first name', 'firstname', 'given name']],
    ['nomcomplet', 'Nom complet', ['nom complet', 'nom et prenom', 'nom prenom', 'nom et prenoms', 'noms et prenoms', 'full name', 'fullname', 'name', 'employe', 'employee', 'agent', 'salarie', 'collaborateur', 'nom employe', 'employee name', 'nom agent', 'identite', 'beneficiaire']],
    ['sexe', 'Sexe', ['sexe', 'genre', 'gender', 'sex']],
    ['naissance', 'Date de naissance', ['date de naissance', 'date naissance', 'date naiss', 'ne le', 'nee le', 'birth date', 'birthdate', 'date of birth', 'dob', 'naissance']],
    ['lieunaissance', 'Lieu de naissance', ['lieu de naissance', 'lieu naissance', 'place of birth', 'birthplace']],
    ['telephone', 'Téléphone', ['telephone', 'tel', 'tel mobile', 'telephone mobile', 'tel portable', 'phone', 'phone number', 'mobile', 'portable', 'contact', 'gsm', 'cellulaire', 'whatsapp']],
    ['email', 'Email', ['email', 'e mail', 'mail', 'courriel', 'adresse email', 'adresse mail', 'adresse e mail', 'email address']],
    ['adresse', 'Adresse', ['adresse', 'address', 'adresse postale', 'domicile', 'residence', 'adresse domicile', 'lieu de residence']],
    ['ville', 'Ville', ['ville', 'city', 'commune', 'localite']],
    ['pays', 'Pays', ['pays', 'country']],
    ['nationalite', 'Nationalité', ['nationalite', 'nationality']],
    ['poste', 'Poste', ['poste', 'fonction', 'job title', 'emploi', 'intitule du poste', 'intitule poste', 'metier', 'profession', 'poste occupe', 'fonction occupee']],
    ['service', 'Service', ['service', 'departement', 'department']],
    ['dateembauche', "Date d'embauche", ['date embauche', 'date d embauche', 'date d entree', 'date entree', 'hire date', 'start date', 'date de recrutement', 'date recrutement', 'date de prise de service', 'date prise de service', 'date d engagement', 'date engagement']],
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
  const GROUP_KEYS = [];
  const GROUP_MAP = new Map();
  GROUPS.forEach(([id, label, names]) => {
    GROUP_LABEL[id] = label;
    names.forEach((n) => {
      const t = tokens(n);
      const k = keyOf(t);
      if (!GROUP_MAP.has(k)) GROUP_MAP.set(k, id);
      if (!GROUP_MAP.has('~' + sortedKeyOf(t))) GROUP_MAP.set('~' + sortedKeyOf(t), id);
      GROUP_KEYS.push({ key: k, group: id });
    });
  });

  function resolveGroup(t, th) {
    if (!t.length) return null;
    const k = keyOf(t);
    if (GROUP_MAP.has(k)) return GROUP_MAP.get(k);
    if (GROUP_MAP.has('~' + sortedKeyOf(t))) return GROUP_MAP.get('~' + sortedKeyOf(t));
    for (const g of GROUP_KEYS) if (fuzzyMatch(k, g.key, th)) return g.group;
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

  const PHONE_HINTS = new Set(['tel', 'phone', 'mobile', 'contact', 'portable', 'gsm', 'whatsapp', 'cell', 'cellulaire', 'numero', 'num']);

  /* ================================================================== */
  /*  Valeurs : nettoyage et reconnaissance                              */
  /* ================================================================== */

  const ERROR_RE = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NULL!|NUM!)/i;
  const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  const DATE_DMY = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/;
  const DATE_ISO = /^(\d{4})-(\d{2})-(\d{2})(?:[T ].*)?$/;
  const DATE_TXT = /^(\d{1,2})(?:er)?\s+([a-zA-ZÀ-ÿ]+)\.?\s+(\d{4})$/;
  const MONTHS = {
    janvier: 1, jan: 1, janv: 1, fevrier: 2, fev: 2, fevr: 2, mars: 3, mar: 3, avril: 4, avr: 4, mai: 5, juin: 6,
    juillet: 7, juil: 7, aout: 8, septembre: 9, sept: 9, sep: 9, octobre: 10, oct: 10, novembre: 11, nov: 11,
    decembre: 12, dec: 12, january: 1, february: 2, march: 3, april: 4, may: 5, june: 6, july: 7, august: 8,
    october: 10, december: 12,
  };
  const MALE = new Set(['m', 'h', 'masculin', 'homme', 'male', 'masc', 'garcon']);
  const FEMALE = new Set(['f', 'feminin', 'femme', 'female', 'fem', 'fille']);

  const isEmpty = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');

  function cleanValue(v) {
    if (isEmpty(v)) return null;
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      const r = new Date(Math.round(v.getTime() / 60000) * 60000);
      return new Date(Date.UTC(r.getFullYear(), r.getMonth(), r.getDate(), r.getHours(), r.getMinutes()));
    }
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (typeof v === 'boolean') return v;
    const s = String(v).replace(/[\s\u00a0]+/g, ' ').trim();
    if (!s || ERROR_RE.test(s)) return null;
    return s;
  }

  function mkDate(y, m, d) {
    if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1900 || y > 2100) return null;
    const dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCMonth() === m - 1 ? dt : null;
  }

  function parseDateText(s) {
    if (typeof s !== 'string') return null;
    const t = s.trim();
    let m = DATE_ISO.exec(t);
    if (m) return mkDate(+m[1], +m[2], +m[3]);
    m = DATE_DMY.exec(t);
    if (m) {
      let d = +m[1];
      let mo = +m[2];
      let y = +m[3];
      if (m[3].length === 2) y += y <= 30 ? 2000 : 1900;
      if (mo > 12 && d <= 12) [d, mo] = [mo, d];
      return mkDate(y, mo, d);
    }
    m = DATE_TXT.exec(t);
    if (m) {
      const mo = MONTHS[stripAccents(m[2])];
      return mo ? mkDate(+m[3], mo, +m[1]) : null;
    }
    return null;
  }

  function phoneDigits(v) {
    let s;
    if (typeof v === 'number') {
      if (!Number.isInteger(v) || v < 0) return null;
      s = String(v);
    } else if (typeof v === 'string') {
      const t = v.trim();
      if (DATE_DMY.test(t) || DATE_ISO.test(t)) return null;
      if (!/^\+?[\d\s.\-()\/]+$/.test(t)) return null;
      s = t.replace(/\D/g, '');
    } else return null;
    return s.length >= 8 && s.length <= 15 ? s : null;
  }

  function parseNumberText(v) {
    if (typeof v === 'number') return v;
    if (typeof v !== 'string') return null;
    let s = v.replace(/\b(fcfa|f cfa|cfa|xof|xaf|eur|euros?|usd)\b/gi, '').replace(/[€$]/g, '').replace(/[\s\u00a0\u202f]/g, '');
    if (!s || !/^[-+]?[\d.,]+$/.test(s)) return null;
    if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
    else if (/^\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, '');
    else {
      const lc = s.lastIndexOf(',');
      const ld = s.lastIndexOf('.');
      if (lc > -1 && ld > -1) s = lc > ld ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
      else if (lc > -1) s = s.replace(',', '.');
    }
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  const sexKey = (v) => stripAccents(v).replace(/[.\s]/g, '');
  const isSexValue = (v) => typeof v === 'string' && (MALE.has(sexKey(v)) || FEMALE.has(sexKey(v)));

  function normalizeSex(v) {
    if (typeof v !== 'string') return v;
    const k = sexKey(v);
    if (MALE.has(k)) return 'Masculin';
    if (FEMALE.has(k)) return 'Féminin';
    return v;
  }

  function isNameLike(v) {
    if (typeof v !== 'string' || !/^[A-Za-zÀ-ÿ'’.\- ]{3,60}$/.test(v)) return false;
    const parts = v.split(' ').filter(Boolean);
    return parts.length >= 2 && parts.length <= 5 && parts.filter((p) => p.replace(/[.'’-]/g, '').length >= 2).length >= parts.length / 2;
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

  const signature = (v) => {
    if (v === undefined || v === null) return '';
    if (v instanceof Date) return 'd' + v.getTime();
    return String(v).toLowerCase();
  };

  /** Statistiques sur le contenu d'une colonne. */
  function profile(vals) {
    const n = vals.length;
    const c = { email: 0, phone: 0, date: 0, sex: 0, number: 0, name: 0 };
    const distinct = new Set();
    vals.forEach((v) => {
      if (typeof v === 'string' && EMAIL_RE.test(v)) c.email++;
      if (phoneDigits(v)) c.phone++;
      if (v instanceof Date || parseDateText(v)) c.date++;
      if (typeof v === 'number') c.number++;
      if (isSexValue(v)) c.sex++;
      if (isNameLike(v)) c.name++;
      if (distinct.size < 60) distinct.add(signature(v));
    });
    const r = {};
    Object.keys(c).forEach((k) => (r[k] = n ? c[k] / n : 0));
    r.n = n;
    r.distinct = distinct;
    return r;
  }

  /** "AGBODJAN Kofi Jean" → {nom, prenom}. */
  function splitName(full) {
    const s = String(full).replace(/\s+/g, ' ').trim();
    if (!s) return { nom: '', prenom: '' };
    if (s.includes(',')) {
      const [a, ...b] = s.split(',');
      return { nom: a.trim(), prenom: b.join(' ').trim() };
    }
    const parts = s.split(' ');
    if (parts.length === 1) return { nom: s, prenom: '' };
    const PART = new Set(['de', 'du', 'des', 'van', 'von', 'le', 'la', 'el', 'al', 'ben', 'ibn']);
    const isUp = (w) => {
      const L = w.replace(/[^A-Za-zÀ-ÿ]/g, '');
      return L.length >= 2 && L === L.toUpperCase() && L !== L.toLowerCase();
    };
    const ups = parts.map(isUp);
    if (ups.every(Boolean)) return { nom: parts[0], prenom: parts.slice(1).join(' ') };
    if (ups.some(Boolean)) {
      let i = 0;
      while (i < parts.length && (ups[i] || (PART.has(parts[i].toLowerCase()) && i < parts.length - 1 && ups[i + 1]))) i++;
      if (i > 0 && i < parts.length) return { nom: parts.slice(0, i).join(' '), prenom: parts.slice(i).join(' ') };
      let j = parts.length;
      while (j > 0 && ups[j - 1]) j--;
      if (j > 0 && j < parts.length) return { nom: parts.slice(j).join(' '), prenom: parts.slice(0, j).join(' ') };
    }
    return { nom: parts[0], prenom: parts.slice(1).join(' ') };
  }

  const titleCase = (s) => s.toLowerCase().replace(/(^|[\s\-'’])([a-zà-ÿ])/g, (m, p, c) => p + c.toUpperCase());

  /* ================================================================== */
  /*  Lecture d'une feuille : où sont les en-têtes (s'il y en a) ?       */
  /* ================================================================== */

  function looksLikeData(row) {
    const cells = (row || []).filter((c) => !isEmpty(c));
    if (cells.length < 2) return false;
    let known = false;
    let dataish = false;
    let nums = 0;
    cells.forEach((c) => {
      if (typeof c === 'string' && resolveGroup(tokens(c))) known = true;
      if (c instanceof Date) dataish = true;
      if (typeof c === 'number') nums++;
      if (EMAIL_RE.test(String(c).trim()) || phoneDigits(c)) dataish = true;
    });
    if (known) return false;
    return dataish || nums / cells.length >= 0.5;
  }

  function readSheet(rows, ho) {
    ho = ho || { mode: 'auto' };
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

    let headerless;
    if (ho.mode === 'none') {
      headerRow = stat.findIndex((s) => s.filled > 0);
      headerless = true;
    } else if (ho.mode === 'fixed' && ho.row >= 1 && ho.row <= rows.length && (rows[ho.row - 1] || []).some((c) => !isEmpty(c))) {
      headerRow = ho.row - 1;
      headerless = false;
    } else {
      // La "ligne d'en-têtes" est peut-être déjà une ligne de données (fichier sans titres de colonnes).
      headerless = looksLikeData(rows[headerRow]);
    }
    const start = headerless ? headerRow : headerRow + 1;

    let hasData = false;
    for (let r = start; r < rows.length && !hasData; r++) hasData = (rows[r] || []).some((c) => !isEmpty(c));
    if (!hasData) return null;

    const head = rows[headerRow] || [];
    let width = head.length;
    for (let r = start; r < rows.length; r++) width = Math.max(width, (rows[r] || []).length);

    const headers = [];
    for (let j = 0; j < width; j++) {
      let text = !headerless && !isEmpty(head[j]) ? headerText(head[j]) : '';
      if (!text) {
        let any = false;
        for (let r = start; r < rows.length && !any; r++) any = !isEmpty((rows[r] || [])[j]);
        if (!any) continue;
        text = 'Colonne ' + (j + 1);
      }
      headers.push({ idx: j, text });
    }
    return headers.length ? { headerRow, headerless, start, headers } : null;
  }

  /* ================================================================== */
  /*  Analyse                                                            */
  /* ================================================================== */

  const baseName = (name) => String(name).replace(/\.[^.]+$/, '');

  /**
   * @param files   [{id?, name, error?, sheets:[{name, rows:any[][]}]}]
   * @param options {allSheets, splitNames, separate:[clés d'intitulés à ne jamais regrouper]}
   */
  function analyze(files, options) {
    const opts = Object.assign({ allSheets: true, splitNames: true, separate: [], sensitivity: 'balanced', contentMatch: true, headerMode: 'auto', headerRow: 1 }, options || {});
    const SENS = {
      strict: { fuzzy: 0.93, syn: 0.92, jac: 2, overlapMin: 1, overlap: 2 },
      balanced: { fuzzy: 0.86, syn: 0.86, jac: 0.75, overlapMin: 2, overlap: 0.5 },
      large: { fuzzy: 0.78, syn: 0.8, jac: 0.6, overlapMin: 2, overlap: 0.34 },
    };
    const sens = SENS[opts.sensitivity] || SENS.balanced;
    const useContent = opts.contentMatch !== false;
    const headerOpts = { mode: opts.headerMode, row: Number(opts.headerRow) || 1 };
    const separate = new Set(opts.separate || []);
    const cols = [];
    const byId = new Map();
    const order = [];
    const records = [];
    const report = [];
    let seq = 0;
    let genSeq = 0;
    let headerlessSheets = 0;

    function createColumn(o) {
      let key;
      let label;
      if (o.own) {
        key = 'sep:' + keyOf(o.t);
        label = prettify(o.text);
      } else if (o.group) {
        key = o.group;
        label = GROUP_LABEL[o.group];
      } else if (o.generic) {
        key = 'gen:' + ++genSeq;
        label = o.text;
      } else {
        key = keyOf(o.t) || stripAccents(o.text);
        label = prettify(o.text);
      }
      const baseKey = key;
      const clash = cols.filter((c) => c.baseKey === baseKey).length;
      if (clash) {
        key += '#' + (clash + 1);
        label += ' (' + (clash + 1) + ')';
      }
      const col = {
        id: 'c' + ++seq,
        key,
        baseKey,
        label,
        group: o.group || null,
        generic: !!o.generic && !o.group,
        ownKey: o.own ? keyOf(o.t) : null,
        norm: o.own ? '@' + keyOf(o.t) : keyOf(o.t),
        sorted: o.own ? '@' + keyOf(o.t) : sortedKeyOf(o.t),
        toks: o.t,
        variants: new Set(),
        variantKeys: new Map(),
        fileIds: new Set(),
        filled: 0,
        reasons: new Set(),
        low: false,
        values: new Set(),
        type: 'text',
        autoHide: false,
      };
      cols.push(col);
      byId.set(col.id, col);
      return col;
    }

    /** Trouve la colonne unifiée qui correspond à cet intitulé / ce contenu. */
    function matchColumn(h, used) {
      const k = keyOf(h.t);
      const sk = sortedKeyOf(h.t);

      if (!h.generic && separate.has(k)) {
        const f = cols.find((c) => c.ownKey === k && !used.has(c.id));
        return f ? { col: f } : { create: { text: h.text, t: h.t, own: true } };
      }

      if (h.group) {
        const f = cols.find((c) => c.group === h.group && !used.has(c.id));
        return f ? { col: f, reason: h.groupReason, low: h.groupReason === 'contenu' } : { create: { text: h.text, t: h.t, group: h.group }, reason: h.groupReason };
      }

      const free = (c) => !used.has(c.id) && !c.group && !c.ownKey && !c.generic;
      if (!h.generic && k) {
        let f = cols.find((c) => free(c) && (c.norm === k || c.sorted === sk));
        if (f) return { col: f, reason: 'intitulé' };
        f = cols.find((c) => {
          if (!free(c) || digitsOf(c.norm) !== digitsOf(k)) return false;
          if (fuzzyMatch(c.norm, k, sens.fuzzy)) return true;
          const j = jaccardTokens(c.toks, h.t);
          return j.union >= 3 && j.score >= sens.jac;
        });
        if (f) return { col: f, reason: 'proche', low: true };
      }

      // Même contenu (listes de valeurs qui se recoupent : services, statuts, villes…)
      const p = h.prof;
      if (useContent && sens.overlap <= 1 && p && p.n && p.distinct.size >= 1 && p.distinct.size <= 40 && p.number < 0.5 && p.date < 0.5) {
        let best = null;
        let bestScore = 0;
        cols.forEach((c) => {
          if (!free(c) && !(c.generic && !used.has(c.id))) return;
          if (c.type !== 'text' || !c.values.size) return;
          const vals = Array.from(c.values);
          if (vals.filter((x) => /^-?\d+([.,]\d+)?$/.test(x)).length > vals.length / 2) return; // colonne numérique
          let inter = 0;
          p.distinct.forEach((x) => c.values.has(x) && inter++);
          const score = inter / (p.distinct.size + c.values.size - inter);
          if (inter >= sens.overlapMin && score >= sens.overlap && score > bestScore) {
            best = c;
            bestScore = score;
          }
        });
        if (best) return { col: best, reason: 'contenu', low: true };
      }
      return { create: { text: h.text, t: h.t, group: null, generic: h.generic } };
    }

    files.forEach((file, fi) => {
      const fileId = file.id != null ? file.id : 'f' + fi;
      const fr = { id: fileId, name: file.name, sheets: [], rows: 0, error: file.error || null };
      report.push(fr);
      if (file.error) return;

      const sheets = opts.allSheets ? file.sheets : file.sheets.slice(0, 1);
      const parsed = sheets.map((sh) => ({ sh, info: readSheet(sh.rows, headerOpts) }));
      const usable = parsed.filter((p) => p.info);
      parsed.forEach((p) => {
        if (!p.info) fr.sheets.push({ name: p.sh.name, rows: 0, cols: 0, empty: true });
      });

      usable.forEach(({ sh, info }) => {
        if (info.headerless) headerlessSheets++;

        // 1. Lignes nettoyées + échantillons par colonne
        const heads = info.headers.map((h) => ({
          idx: h.idx,
          text: h.text,
          t: tokens(h.text),
          generic: info.headerless || isGenericHeader(h.text),
        }));
        const samples = heads.map(() => []);
        const rowsClean = [];
        for (let r = info.start; r < sh.rows.length; r++) {
          const row = sh.rows[r];
          if (!row) continue;
          const cells = heads.map((h, i) => {
            const v = cleanValue(row[h.idx]);
            if (v !== null && samples[i].length < 300) samples[i].push(v);
            return v;
          });
          if (cells.some((v) => v !== null)) rowsClean.push(cells);
        }

        // 2. Quel type d'information contient chaque colonne ?
        heads.forEach((h, i) => {
          h.prof = profile(samples[i]);
          h.group = h.generic ? null : resolveGroup(h.t, sens.syn);
          h.groupReason = h.group ? 'synonyme' : null;
        });
        if (useContent) heads.forEach((h) => {
          const p = h.prof;
          const emailish = p.n && p.email >= 0.7;
          const phoneish = p.n && p.phone >= 0.7 && p.email < 0.3;
          if (h.group === 'telephone' && emailish) [h.group, h.groupReason] = ['email', 'contenu'];
          else if (h.group === 'email' && phoneish) [h.group, h.groupReason] = ['telephone', 'contenu'];
          else if (!h.group) {
            if (emailish) [h.group, h.groupReason] = ['email', 'contenu'];
            else if (h.generic && p.n && p.sex >= 0.8 && p.distinct.size <= 6) [h.group, h.groupReason] = ['sexe', 'contenu'];
            else if ((h.generic || h.t.some((w) => PHONE_HINTS.has(w))) && phoneish) [h.group, h.groupReason] = ['telephone', 'contenu'];
          }
        });
        if (useContent && !heads.some((h) => ['nom', 'prenom', 'nomcomplet'].includes(h.group))) {
          const cand = heads.find((h) => (h.generic || !h.group) && h.prof.n && h.prof.name >= 0.8);
          if (cand) [cand.group, cand.groupReason] = ['nomcomplet', 'contenu'];
        }

        // 3. Rapprochement avec les colonnes déjà connues
        const used = new Set();
        let prev = null;
        const map = [];
        heads.forEach((h, i) => {
          const m = matchColumn(h, used);
          let col = m.col;
          if (!col) {
            col = createColumn(m.create);
            const at = prev ? order.indexOf(prev) + 1 : 0;
            order.splice(at, 0, col.id);
          }
          if (m.reason) col.reasons.add(m.reason);
          if (m.low || h.groupReason === 'contenu') col.low = true;
          used.add(col.id);
          col.variants.add(prettify(h.text));
          if (!h.generic) col.variantKeys.set(prettify(h.text), keyOf(h.t));
          col.fileIds.add(fileId);
          h.prof.distinct.forEach((x) => col.values.size < 80 && col.values.add(x));
          map.push({ i, col });
          prev = col.id;
        });

        // 4. Enregistrements
        const headKeys = heads.map((h) => (h.generic ? null : keyOf(h.t)));
        const label = usable.length > 1 ? baseName(file.name) + ' – ' + sh.name : baseName(file.name);
        let count = 0;
        rowsClean.forEach((cells) => {
          const vals = {};
          let filled = 0;
          let repeated = 0;
          map.forEach(({ i, col }) => {
            const v = cells[i];
            if (v === null) return;
            vals[col.id] = v;
            filled++;
            if (headKeys[i] && typeof v === 'string' && keyOf(tokens(v)) === headKeys[i]) repeated++;
          });
          if (!filled) return;
          if (map.length >= 2 && repeated >= Math.max(2, Math.ceil(map.length * 0.6))) return;
          if (filled === 1 && map.length >= 3) {
            const only = Object.values(vals)[0];
            if (typeof only === 'string' && /^(total|sous[- ]?total|signature|fait [aà])/i.test(only)) return;
          }
          records.push({ fileId, file: baseName(file.name), source: label, vals });
          count++;
        });
        fr.sheets.push({ name: sh.name, rows: count, cols: map.length, headerRow: info.headerless ? null : info.headerRow + 1, headerless: info.headerless });
        fr.rows += count;
      });
    });

    /* ---- Typage des colonnes : dates, nombres, téléphones ---------- */
    const stats = new Map(cols.map((c) => [c.id, { n: 0, date: 0, isNum: 0, num: 0 }]));
    records.forEach((r) => {
      for (const id in r.vals) {
        const v = r.vals[id];
        const s = stats.get(id);
        s.n++;
        if (v instanceof Date || parseDateText(v)) s.date++;
        if (typeof v === 'number') s.isNum++;
        else if (parseNumberText(v) !== null) s.num++;
      }
    });
    cols.forEach((c) => {
      const s = stats.get(c.id);
      if (!s.n) return;
      if (['telephone', 'matricule', 'cnss'].includes(c.group)) c.type = 'text';
      else if (s.date / s.n >= 0.7) c.type = 'date';
      else if ((s.isNum / s.n >= 0.7 || ['salaire', 'enfants'].includes(c.group)) && (s.isNum + s.num) / s.n >= 0.7) c.type = 'number';
    });
    records.forEach((r) => {
      for (const id in r.vals) {
        const c = byId.get(id);
        const v = r.vals[id];
        if (c.type === 'date' && typeof v === 'string') {
          const d = parseDateText(v);
          if (d) r.vals[id] = d;
        } else if (c.type === 'number' && typeof v === 'string') {
          const n = parseNumberText(v);
          if (n !== null) r.vals[id] = n;
        } else if (c.group === 'telephone' && typeof v === 'number') r.vals[id] = String(v);
      }
    });

    /* ---- "Nom et Prénoms" → Nom + Prénom quand d'autres fichiers les séparent ---- */
    let splitCount = 0;
    if (opts.splitNames) {
      const full = cols.find((c) => c.group === 'nomcomplet');
      let nomC = cols.find((c) => c.group === 'nom');
      let preC = cols.find((c) => c.group === 'prenom');
      if (full && (nomC || preC)) {
        const make = (group) => {
          const c = createColumn({ text: GROUP_LABEL[group], t: tokens(GROUP_LABEL[group]), group });
          order.splice(order.indexOf(full.id), 0, c.id);
          c.variants.add('déduit de « ' + full.label + ' »');
          full.fileIds.forEach((f) => c.fileIds.add(f));
          return c;
        };
        if (!nomC) nomC = make('nom');
        if (!preC) preC = make('prenom');
        records.forEach((r) => {
          const f = r.vals[full.id];
          if (typeof f !== 'string') return;
          const sp = splitName(f);
          let did = false;
          if (sp.nom && r.vals[nomC.id] === undefined) [r.vals[nomC.id], did] = [sp.nom, true];
          if (sp.prenom && r.vals[preC.id] === undefined) [r.vals[preC.id], did] = [sp.prenom, true];
          if (did) splitCount++;
        });
        if (splitCount) {
          nomC.reasons.add('déduit');
          preC.reasons.add('déduit');
          full.autoHide = true;
        }
      }
    }

    records.forEach((r) => Object.keys(r.vals).forEach((id) => byId.get(id).filled++));

    const columns = order.map((id) => {
      const c = byId.get(id);
      return {
        id: c.id,
        key: c.key,
        label: c.label,
        group: c.group,
        type: c.type,
        variants: Array.from(c.variants),
        variantKeys: Array.from(c.variants).map((v) => c.variantKeys.get(v) || null),
        fileCount: c.fileIds.size,
        filled: c.filled,
        reasons: Array.from(c.reasons),
        low: c.low,
        autoHide: c.autoHide,
      };
    });

    return {
      columns,
      records,
      report,
      fileCount: files.filter((f) => !f.error).length,
      mergedColumns: columns.filter((c) => c.variants.length > 1 && !c.reasons.includes('déduit')).length,
      contentColumns: columns.filter((c) => c.reasons.includes('contenu')).length,
      splitCount,
      headerlessSheets,
    };
  }

  /* ================================================================== */
  /*  Construction du tableau final                                      */
  /* ================================================================== */

  const collator = typeof Intl !== 'undefined' ? new Intl.Collator('fr', { numeric: true, sensitivity: 'base' }) : null;

  function compare(a, b) {
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (a instanceof Date && b instanceof Date) return a - b;
    const sa = a instanceof Date ? a.toISOString() : String(a);
    const sb = b instanceof Date ? b.toISOString() : String(b);
    return collator ? collator.compare(sa, sb) : sa.localeCompare(sb);
  }

  const caseOf = (v, mode) => (mode === 'upper' ? v.toLocaleUpperCase('fr') : mode === 'title' ? titleCase(v) : v);

  /** Formate un numéro de téléphone (espaces, indicatif…). */
  function formatPhone(v, fmt, cc) {
    const raw = String(v).trim();
    let plus = raw.startsWith('+');
    let d = raw.replace(/\D/g, '');
    if (!d) return raw;
    if (d.startsWith('00')) {
      d = d.slice(2);
      plus = true;
    }
    if (fmt === 'digits') return (plus ? '+' : '') + d;
    const code = String(cc || '').replace(/\D/g, '');
    const grp = (x) => x.replace(/(\d{2})(?=\d)/g, '$1 ').trim();
    if (plus && code && d.startsWith(code)) return '+' + code + ' ' + grp(d.slice(code.length));
    if (fmt === 'intl' && code && !plus) {
      if (d.startsWith(code) && d.length > 9) return '+' + code + ' ' + grp(d.slice(code.length));
      return '+' + code + ' ' + grp(d);
    }
    return (plus ? '+' : '') + grp(d);
  }

  function cleanByGroup(group, v, st) {
    if (group === 'telephone') {
      return st.phoneFormat && st.phoneFormat !== 'keep' && (typeof v === 'string' || typeof v === 'number') ? formatPhone(v, st.phoneFormat, st.phoneCC) : v;
    }
    if (typeof v !== 'string') return v;
    switch (group) {
      case 'nom': return caseOf(v, st.nomCase || 'upper');
      case 'prenom': return caseOf(v, st.prenomCase || 'title');
      case 'ville':
      case 'pays':
      case 'nationalite': return st.placeCase === false ? v : titleCase(v);
      case 'email': return st.emailLower === false ? v : v.toLowerCase();
      case 'sexe': return st.sexNorm === false ? v : normalizeSex(v);
      default: return v;
    }
  }

  const nameTokens = (s) => stripAccents(s).split(/[^a-z]+/).filter((w) => w.length >= 2);

  function nameSimilar(a, b) {
    if (a === b) return true;
    const ta = a.split(' ');
    const tb = b.split(' ');
    const [small, big] = ta.length <= tb.length ? [ta, tb] : [tb, ta];
    if (small.every((w) => big.includes(w))) return true;
    return similarity(a, b) >= 0.82;
  }

  /**
   * @param st {cols:[{id,label,include,into?}], source, numbering, combine, keyCol:'auto'|id, conflict:'first'|'last'|'longest',
   *            clean, nomCase, prenomCase, emailLower, sexNorm, placeCase, phoneFormat, phoneCC, dropEmptyCols,
   *            sortBy:'auto'|'none'|id, sortDir, sortBy2:'none'|id, sortDir2}
   */
  function build(analysis, st) {
    const meta = new Map(analysis.columns.map((c) => [c.id, c]));
    const colsAll = st.cols.filter((c) => meta.has(c.id));

    // Rattachements manuels : la colonne B est versée dans la colonne A.
    const into = new Map();
    colsAll.forEach((c) => {
      if (c.into && c.into !== c.id && meta.has(c.into) && !(colsAll.find((x) => x.id === c.into) || {}).into) into.set(c.id, c.into);
    });
    let cols = colsAll
      .filter((c) => c.include && !into.has(c.id))
      .map((c) => ({ id: c.id, label: c.label, group: meta.get(c.id).group, meta: meta.get(c.id) }));

    const recs = analysis.records.map((r) => {
      const vals = {};
      Object.keys(r.vals).forEach((id) => {
        if (!into.has(id)) vals[id] = r.vals[id];
      });
      Object.keys(r.vals).forEach((id) => {
        if (into.has(id) && vals[into.get(id)] === undefined) vals[into.get(id)] = r.vals[id];
      });
      return { source: r.source, vals };
    });

    if (st.clean) {
      recs.forEach((r) => {
        for (const id in r.vals) r.vals[id] = cleanByGroup(meta.get(id).group, r.vals[id], st);
      });
    }

    const ids = (g) => analysis.columns.filter((c) => c.group === g).map((c) => c.id);
    const matIds = ids('matricule');
    const mailIds = ids('email');
    const telIds = ids('telephone');
    const nomIds = ids('nom');
    const preIds = ids('prenom');
    const fullIds = ids('nomcomplet');
    const first = (r, list) => {
      for (const id of list) if (r.vals[id] !== undefined) return r.vals[id];
      return undefined;
    };
    const customKey = st.keyCol && st.keyCol !== 'auto' && meta.has(st.keyCol) ? (into.get(st.keyCol) || st.keyCol) : null;

    function keysOf(r) {
      const k = { mat: null, mail: null, tel: null, name: null, custom: null };
      if (customKey) {
        const v = r.vals[customKey];
        if (v !== undefined) k.custom = stripAccents(v instanceof Date ? v.getTime() : String(v)).replace(/\s+/g, ' ').trim();
        return k;
      }
      const m = first(r, matIds);
      if (m !== undefined) k.mat = String(m).toLowerCase().replace(/\s+/g, '');
      const e = first(r, mailIds);
      if (typeof e === 'string' && EMAIL_RE.test(e)) k.mail = e.toLowerCase();
      const t = first(r, telIds);
      const d = t === undefined ? null : phoneDigits(t);
      if (d) k.tel = d.slice(-8);
      let toks = [];
      const n = first(r, nomIds);
      const p = first(r, preIds);
      if (n !== undefined) toks = toks.concat(nameTokens(String(n)));
      if (p !== undefined) toks = toks.concat(nameTokens(String(p)));
      if (!toks.length) {
        const f = first(r, fullIds);
        if (f !== undefined) toks = nameTokens(String(f));
      }
      if (toks.length) k.name = toks.sort().join(' ');
      return k;
    }

    const KINDS = customKey ? ['custom'] : ['mat', 'mail', 'name', 'tel'];
    const strongKey = (k) => (customKey ? !!k.custom : !!(k.mat || k.mail || k.tel || (k.name && k.name.split(' ').length >= 2)));

    function absorb(target, r) {
      target.n++;
      target.sources.add(r.source);
      for (const id in r.vals) {
        const v = r.vals[id];
        const cur = target.vals[id];
        if (cur === undefined) {
          target.vals[id] = v;
        } else if (signature(cur) !== signature(v)) {
          const alts = (target.conflicts[id] = target.conflicts[id] || new Set());
          const show = (x) => (x instanceof Date ? dateToText(x) : String(x));
          const pick = st.conflict === 'last' || (st.conflict === 'longest' && String(v).length > String(cur).length);
          if (pick) {
            alts.add(show(cur));
            target.vals[id] = v;
          } else alts.add(show(v));
        }
      }
    }

    let groups = [];
    let absorbed = 0;
    if (st.combine) {
      const maps = { mat: new Map(), mail: new Map(), name: new Map(), tel: new Map(), custom: new Map(), sig: new Map() };
      const compatible = (g, k) => {
        if (customKey) return true;
        if (g.k.mat && k.mat && g.k.mat !== k.mat) return false;
        if (g.k.name && k.name && !nameSimilar(g.k.name, k.name)) return false;
        return true;
      };
      recs.forEach((r) => {
        const k = keysOf(r);
        let target = null;
        for (const kind of KINDS) {
          if (!k[kind]) continue;
          if (kind === 'name' && k.name.split(' ').length < 2) continue;
          const g = maps[kind].get(k[kind]);
          if (g && compatible(g, k)) {
            target = g;
            break;
          }
        }
        let sig = null;
        if (!target && !strongKey(k)) {
          sig = cols.map((c) => signature(r.vals[c.id])).join('\u0001');
          target = maps.sig.get(sig) || null;
        }
        if (target) {
          absorbed++;
          absorb(target, r);
          if (!target.k.mat) target.k.mat = k.mat;
          if (!target.k.name) target.k.name = k.name;
          KINDS.forEach((kind) => k[kind] && !maps[kind].has(k[kind]) && maps[kind].set(k[kind], target));
        } else {
          const g = { vals: Object.assign({}, r.vals), sources: new Set([r.source]), conflicts: {}, n: 1, k };
          groups.push(g);
          KINDS.forEach((kind) => k[kind] && !maps[kind].has(k[kind]) && maps[kind].set(k[kind], g));
          if (sig !== null) maps.sig.set(sig, g);
        }
      });
    } else {
      groups = recs.map((r) => ({ vals: r.vals, sources: new Set([r.source]), conflicts: {}, n: 1 }));
    }

    let emptyDropped = 0;
    if (st.dropEmptyCols) {
      const before = cols.length;
      cols = cols.filter((c) => groups.some((g) => g.vals[c.id] !== undefined));
      emptyDropped = before - cols.length;
    }
    groups = groups.filter((g) => cols.some((c) => g.vals[c.id] !== undefined));

    // Tri (principal + secondaire)
    let keys = [];
    const resolve = (id) => (id && id !== 'none' && id !== 'auto' ? cols.find((x) => x.id === (into.get(id) || id)) : null);
    if (st.sortBy === 'auto') {
      const g = (id) => cols.find((c) => c.group === id);
      if (g('nom')) keys = [g('nom'), g('prenom')].filter(Boolean).map((c) => ({ c, dir: st.sortDir === 'desc' ? -1 : 1 }));
      else if (g('nomcomplet')) keys = [{ c: g('nomcomplet'), dir: st.sortDir === 'desc' ? -1 : 1 }];
      else if (g('matricule')) keys = [{ c: g('matricule'), dir: st.sortDir === 'desc' ? -1 : 1 }];
    } else if (st.sortBy && st.sortBy !== 'none') {
      const c = resolve(st.sortBy);
      if (c) keys.push({ c, dir: st.sortDir === 'desc' ? -1 : 1 });
      const c2 = resolve(st.sortBy2);
      if (c2 && (!c || c2.id !== c.id)) keys.push({ c: c2, dir: st.sortDir2 === 'desc' ? -1 : 1 });
    }
    if (keys.length) {
      groups = groups.slice().sort((x, y) => {
        for (const { c, dir } of keys) {
          const a = x.vals[c.id];
          const b = y.vals[c.id];
          if (a === undefined || b === undefined) {
            if (a === undefined && b === undefined) continue;
            return a === undefined ? 1 : -1;
          }
          const r = compare(a, b);
          if (r) return r * dir;
        }
        return 0;
      });
    }

    const off = st.numbering ? 1 : 0;
    const headers = [];
    if (st.numbering) headers.push('N°');
    cols.forEach((c) => headers.push(c.label || 'Colonne'));
    if (st.source) headers.push('Fichier source');

    const conflictMap = {};
    let conflictCount = 0;
    const rows = groups.map((g, ri) => {
      const out = [];
      if (st.numbering) out.push(ri + 1);
      cols.forEach((c, ci) => {
        out.push(g.vals[c.id] === undefined ? null : g.vals[c.id]);
        const alts = g.conflicts[c.id];
        if (alts && alts.size) {
          conflictMap[ri + ':' + (ci + off)] = Array.from(alts).join(' | ');
          conflictCount++;
        }
      });
      if (st.source) out.push(Array.from(g.sources).join(' + '));
      return out;
    });

    const headMeta = (st.numbering ? [{ variants: [], merged: false, group: 'index' }] : [])
      .concat(cols.map((c) => ({ variants: c.meta.variants, merged: c.meta.variants.length > 1 && !c.meta.reasons.includes('déduit'), low: c.meta.low, group: c.group, type: c.meta.type })))
      .concat(st.source ? [{ variants: [], merged: false, group: 'source' }] : []);

    return {
      headers,
      rows,
      headMeta,
      conflictMap,
      conflictCount,
      absorbed,
      emptyDropped,
      combinedRows: groups.filter((g) => g.n > 1).length,
      sortedBy: keys.map((k) => k.c.label),
      columnCount: cols.length,
    };
  }

  return { analyze, build, dateToText, tokens, resolveGroup, readSheet, splitName, parseDateText, formatPhone, EMAIL_RE, phoneDigits };
});
