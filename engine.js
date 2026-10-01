/*
 * Live Timeline — scheduling rules engine.
 * Pure functions, no DOM and no Office.js, so the same code runs in the task pane,
 * in the browser preview and in Node tests.
 *
 * Dates are handled at month precision as a "month index": year * 12 + (month - 1).
 */
(function (root) {
  'use strict';

  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var SCHEMA = 'live-timeline/1';

  var DEFAULT_RULES = {
    reviewPriority: 8,      // months from BLA submission to approval with priority review / PRV (60-day filing + 6-month goal)
    reviewStandard: 12,     // months from submission to approval with standard review (60-day filing + 10-month goal)
    acipMonths: [2, 6, 10], // months in which ACIP is assumed to meet
    acipMinGap: 1,          // ACIP vote at the first meeting at least this many months after approval
    readoutLag: 2,          // months from primary completion (PCD) to topline data
    submissionLag: 1,       // months from topline data to submission
    pointLabels: 'quarter'  // how D and S markers are labelled: 'quarter' or 'month'
  };

  // Scenario input fields, in the order used by the task pane, the team-update table and the change log.
  var FIELDS = [
    { key: 'studyName', label: 'Study name', type: 'text' },
    { key: 'ssd', label: 'Study start (SSD)', type: 'month' },
    { key: 'studyMonths', label: 'Study length (months, SSD to PCD)', type: 'int' },
    { key: 'readoutLag', label: 'Topline after PCD (months)', type: 'int', ruleDefault: 'readoutLag' },
    { key: 'submissionLag', label: 'Submission after topline (months)', type: 'int', ruleDefault: 'submissionLag' },
    { key: 'submission', label: 'Fixed submission date', type: 'month' },
    { key: 'priorityReview', label: 'Priority review / PRV', type: 'bool' },
    { key: 'reviewMonths', label: 'Review length override (months)', type: 'int' },
    { key: 'approval', label: 'Fixed FDA approval date', type: 'month' },
    { key: 'note', label: 'Note shown above the row', type: 'text' }
  ];

  var EVENT_KEYS = ['ssd', 'pcd', 'd', 's', 'fda', 'acip'];
  var EVENT_NAMES = { ssd: 'Study start', pcd: 'Primary completion', d: 'Data readout', s: 'Submission', fda: 'FDA approval', acip: 'ACIP' };

  // ---------- month helpers ----------
  function mi(year, month) { return year * 12 + (month - 1); }
  function yearOf(i) { return Math.floor(i / 12); }
  function monthOf(i) { return ((i % 12) + 12) % 12 + 1; }
  function yy(i) { var y = yearOf(i) % 100; return (y < 10 ? '0' : '') + y; }
  function quarterOf(i) { return Math.floor((monthOf(i) - 1) / 3) + 1; }
  function monthLabel(i) { return MONTHS[monthOf(i) - 1] + " '" + yy(i); }
  function quarterLabel(i) { return quarterOf(i) + "Q '" + yy(i); }
  function longMonth(i) { return MONTHS[monthOf(i) - 1] + ' ' + yearOf(i); }
  function toYM(i) { var m = monthOf(i); return yearOf(i) + '-' + (m < 10 ? '0' : '') + m; }

  function monthIndexOfName(name) {
    var n = String(name).slice(0, 3).toLowerCase();
    for (var k = 0; k < 12; k++) if (MONTHS[k].toLowerCase() === n) return k + 1;
    return 0;
  }
  function fullYear(y) { y = Number(y); return y < 100 ? 2000 + y : y; }

  /**
   * Reads a month from the many ways people type dates:
   * 2027-01, 2027/1, 01/2027, Jan 2027, Jan '27, Jan-27, January 2027, 4Q '26, Q4 2026, 1/15/2027,
   * an Excel serial date, or an existing month index. Quarters map to their middle month.
   * Returns a month index, or null when blank or unreadable.
   */
  function parseMonth(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return isFinite(v) ? Math.round(v) : null;
    var s = String(v).trim().replace(/[‘’`´]/g, "'");
    if (!s) return null;
    var m;
    if ((m = s.match(/^(\d{4})[-\/.](\d{1,2})(?:[-\/.]\d{1,2})?$/))) return okMonth(+m[1], +m[2]);
    if ((m = s.match(/^(\d{1,2})[-\/.](\d{4})$/))) return okMonth(+m[2], +m[1]);
    if ((m = s.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})$/))) {
      // Day/month order is ambiguous; treat as month/day unless the first number cannot be a month.
      var a = +m[1], b = +m[2];
      return okMonth(fullYear(m[3]), a > 12 ? b : a);
    }
    if ((m = s.match(/^([A-Za-z]{3,9})\.?[\s\-,]*'?(\d{2}|\d{4})$/))) {
      var mon = monthIndexOfName(m[1]);
      return mon ? mi(fullYear(m[2]), mon) : null;
    }
    if ((m = s.match(/^\d{1,2}[\s\-]([A-Za-z]{3,9})[\s\-,]*'?(\d{2}|\d{4})$/))) {
      var mon2 = monthIndexOfName(m[1]);
      return mon2 ? mi(fullYear(m[2]), mon2) : null;
    }
    if ((m = s.match(/^([1-4])\s*Q\s*'?(\d{2}|\d{4})$/i)) || (m = s.match(/^Q\s*([1-4])\s*'?(\d{2}|\d{4})$/i))) {
      return mi(fullYear(m[2]), (+m[1] - 1) * 3 + 2);
    }
    if ((m = s.match(/^(\d{5})(?:\.\d+)?$/))) {
      var d = new Date(Date.UTC(1899, 11, 30) + Number(m[1]) * 86400000);
      return mi(d.getUTCFullYear(), d.getUTCMonth() + 1);
    }
    return null;
  }
  function okMonth(y, mo) { return (mo >= 1 && mo <= 12) ? mi(y, mo) : null; }

  /** Normalises a month value to 'YYYY-MM' (or '' when blank). Unreadable text is returned unchanged so it can be flagged. */
  function normMonth(v) {
    if (v === null || v === undefined || String(v).trim() === '') return '';
    var i = parseMonth(v);
    return i === null ? String(v).trim() : toYM(i);
  }

  function num(v, dflt) {
    if (v === '' || v === null || v === undefined) return dflt;
    var n = Number(v);
    return isFinite(n) ? n : dflt;
  }

  function nextAcip(fromIdx, months) {
    var list = (months && months.length ? months : DEFAULT_RULES.acipMonths)
      .map(Number).filter(function (x) { return x >= 1 && x <= 12; });
    if (!list.length) return null;
    for (var i = fromIdx; i < fromIdx + 36; i++) if (list.indexOf(monthOf(i)) >= 0) return i;
    return null;
  }

  function rulesOf(model) {
    var r = {}, k;
    for (k in DEFAULT_RULES) r[k] = DEFAULT_RULES[k];
    if (model && model.rules) for (k in model.rules) if (model.rules[k] !== undefined && model.rules[k] !== '') r[k] = model.rules[k];
    return r;
  }

  /**
   * Applies the rules to one scenario.
   * SSD + study length -> PCD; PCD + topline lag -> D; D + submission lag -> S (unless a fixed submission date is set);
   * S + review length (priority or standard) -> FDA approval (unless fixed); approval -> next ACIP meeting.
   */
  function computeScenario(sc, rules) {
    rules = rules || DEFAULT_RULES;
    var ev = { ssd: null, pcd: null, d: null, s: null, fda: null, acip: null };
    var warnings = [];

    var ssd = parseMonth(sc.ssd);
    if (sc.ssd && ssd === null) warnings.push('Could not read the study start "' + sc.ssd + '"');
    if (ssd !== null) {
      var len = num(sc.studyMonths, null);
      if (len === null || len < 0) { warnings.push('Study length is missing'); len = 0; }
      ev.ssd = ssd;
      ev.pcd = ssd + Math.round(len);
      if (sc.showReadout !== false) ev.d = ev.pcd + Math.round(num(sc.readoutLag, rules.readoutLag));
    }

    var subFixed = parseMonth(sc.submission);
    if (sc.submission && subFixed === null) warnings.push('Could not read the submission date "' + sc.submission + '"');
    if (subFixed !== null) ev.s = subFixed;
    else if (ev.d !== null) ev.s = ev.d + Math.round(num(sc.submissionLag, rules.submissionLag));
    else if (ev.pcd !== null) ev.s = ev.pcd + Math.round(num(sc.readoutLag, rules.readoutLag)) + Math.round(num(sc.submissionLag, rules.submissionLag));

    var review = Math.round(num(sc.reviewMonths, sc.priorityReview ? rules.reviewPriority : rules.reviewStandard));
    var apFixed = parseMonth(sc.approval);
    if (sc.approval && apFixed === null) warnings.push('Could not read the approval date "' + sc.approval + '"');
    if (apFixed !== null) ev.fda = apFixed;
    else if (ev.s !== null) ev.fda = ev.s + review;

    if (ev.fda !== null && sc.showAcip !== false) ev.acip = nextAcip(ev.fda + Math.round(num(rules.acipMinGap, 1)), rules.acipMonths);

    if (ev.d !== null && ev.s !== null && ev.s < ev.d) warnings.push('Submission is earlier than the data readout');
    if (ev.ssd !== null && ev.s !== null && ev.s < ev.ssd) warnings.push('Submission is earlier than the study start');
    if (ev.s !== null && ev.fda !== null && ev.fda <= ev.s) warnings.push('Approval is not after submission');
    if (ev.ssd === null && ev.s === null) warnings.push('Set a study start or a fixed submission date');

    return { id: sc.id, events: ev, review: review, warnings: warnings };
  }

  function computeAll(model) {
    var rules = rulesOf(model);
    return (model.scenarios || []).map(function (sc) { return computeScenario(sc, rules); });
  }

  /** Text labels in the slide's conventions (SSD/PCD/FDA/ACIP by month; D and S by quarter unless set to month). */
  function labels(ev, rules) {
    var pt = (rules && rules.pointLabels === 'month') ? monthLabel : quarterLabel;
    return {
      ssd: ev.ssd === null ? '' : 'SSD ' + monthLabel(ev.ssd),
      pcd: ev.pcd === null ? '' : 'PCD ' + monthLabel(ev.pcd),
      d: ev.d === null ? '' : pt(ev.d),
      s: ev.s === null ? '' : pt(ev.s),
      fda: ev.fda === null ? '' : 'FDA ' + monthLabel(ev.fda),
      acip: ev.acip === null ? '' : 'ACIP ' + monthLabel(ev.acip)
    };
  }

  // ---------- model helpers ----------
  function blankScenario(i) {
    return {
      id: 's' + (i + 1), name: 'Scenario ' + (i + 1), row: i,
      studyName: 'Ph3', ssd: '', studyMonths: 12, readoutLag: '', submissionLag: '',
      submission: '', priorityReview: false, reviewMonths: '', approval: '',
      showAcip: true, note: '', highlightAcip: false
    };
  }

  function normalizeModel(m) {
    m = m ? JSON.parse(JSON.stringify(m)) : {};
    m.schema = SCHEMA;
    if (!Array.isArray(m.scenarios)) m.scenarios = [];
    m.rules = Object.assign({}, DEFAULT_RULES, m.rules || {});
    if (!Array.isArray(m.rules.acipMonths)) m.rules.acipMonths = DEFAULT_RULES.acipMonths.slice();
    m.scenarios = m.scenarios.map(function (sc, i) {
      var b = blankScenario(i), out = Object.assign(b, sc);
      if (!out.id) out.id = 's' + (i + 1);
      if (out.row === undefined || out.row === null || out.row === '') out.row = i;
      ['ssd', 'submission', 'approval'].forEach(function (k) { out[k] = normMonth(out[k]); });
      return out;
    });
    if (m.showToday === undefined) m.showToday = true;
    if (!Array.isArray(m.changelog)) m.changelog = [];
    if (m.baseline === undefined) m.baseline = null;
    if (m.showSlippage === undefined) m.showSlippage = false;
    return m;
  }

  function snapshot(model, label) {
    var events = {};
    computeAll(model).forEach(function (r) {
      var e = {};
      EVENT_KEYS.forEach(function (k) { e[k] = r.events[k] === null ? '' : toYM(r.events[k]); });
      events[r.id] = e;
    });
    return { label: label || 'Baseline', savedAt: new Date().toISOString(), events: events };
  }

  /** Month differences between the current schedule and the saved baseline, per scenario and event. */
  function deltas(model) {
    var out = {};
    if (!model.baseline || !model.baseline.events) return out;
    computeAll(model).forEach(function (r) {
      var b = model.baseline.events[r.id];
      if (!b) return;
      var d = {};
      EVENT_KEYS.forEach(function (k) {
        var bi = parseMonth(b[k]), ci = r.events[k];
        d[k] = (bi === null || ci === null) ? null : ci - bi;
      });
      out[r.id] = { delta: d, base: b };
    });
    return out;
  }

  function fieldValueText(f, v) {
    if (f.type === 'bool') return v ? 'Y' : 'N';
    if (f.type === 'month') { var i = parseMonth(v); return i === null ? (v || '') : longMonth(i); }
    return (v === null || v === undefined) ? '' : String(v);
  }

  function fieldByKey(key) { for (var i = 0; i < FIELDS.length; i++) if (FIELDS[i].key === key) return FIELDS[i]; return null; }

  function coerce(f, raw) {
    var s = (raw === null || raw === undefined) ? '' : String(raw).trim();
    if (f.type === 'bool') return /^(y|yes|true|1|x|prv|priority)$/i.test(s);
    if (f.type === 'int') { if (s === '') return ''; var n = Number(s.replace(/[^\d.\-]/g, '')); return isFinite(n) ? Math.round(n) : ''; }
    if (f.type === 'month') return normMonth(s);
    return s;
  }

  function sameValue(f, a, b) {
    if (f.type === 'bool') return !!a === !!b;
    if (f.type === 'month') return normMonth(a) === normMonth(b);
    return String(a === null || a === undefined ? '' : a) === String(b === null || b === undefined ? '' : b);
  }

  /** Tab-separated table of scenario inputs for the team to edit in Excel, Teams or email. */
  function toTSV(model) {
    var head = ['ID', 'Scenario'].concat(FIELDS.map(function (f) { return f.label; }));
    var rows = [head.join('\t')];
    (model.scenarios || []).forEach(function (sc) {
      rows.push([sc.id, sc.name].concat(FIELDS.map(function (f) { return fieldValueText(f, sc[f.key]).replace(/[\t\r\n]+/g, ' '); })).join('\t'));
    });
    return rows.join('\n');
  }

  /**
   * Reads a pasted table (tab- or comma-separated, header row required) and returns the proposed changes
   * without applying them: [{ id, name, key, label, from, to }], plus problems found while reading.
   */
  function parseUpdates(model, text) {
    var lines = String(text || '').replace(/\r/g, '').split('\n').filter(function (l) { return l.trim() !== ''; });
    var problems = [], changes = [];
    if (lines.length < 2) return { changes: changes, problems: ['Paste the header row and at least one scenario row.'] };
    var sep = lines[0].indexOf('\t') >= 0 ? '\t' : ',';
    var split = function (l) { return sep === '\t' ? l.split('\t') : splitCSV(l); };
    var head = split(lines[0]).map(function (h) { return h.trim().toLowerCase(); });
    var col = {};
    head.forEach(function (h, i) {
      if (h === 'id') col.id = i;
      else if (h === 'scenario') col.name = i;
      else FIELDS.forEach(function (f) { if (h === f.label.toLowerCase() || h === f.key.toLowerCase()) col[f.key] = i; });
    });
    if (col.id === undefined && col.name === undefined) problems.push('No "ID" or "Scenario" column found in the header row.');
    lines.slice(1).forEach(function (l, n) {
      var cells = split(l);
      var id = col.id !== undefined ? String(cells[col.id] || '').trim() : '';
      var name = col.name !== undefined ? String(cells[col.name] || '').trim() : '';
      var sc = null;
      (model.scenarios || []).forEach(function (s) { if ((id && s.id === id) || (!sc && name && s.name.trim().toLowerCase() === name.toLowerCase())) sc = sc || s; });
      if (!sc) { problems.push('Row ' + (n + 2) + ': no scenario matches "' + (id || name) + '"'); return; }
      FIELDS.forEach(function (f) {
        if (col[f.key] === undefined) return;
        var raw = cells[col[f.key]];
        if (raw === undefined) return;
        var v = coerce(f, raw);
        if (f.type === 'month' && v && parseMonth(v) === null) { problems.push(sc.id + ' ' + f.label + ': could not read "' + raw + '"'); return; }
        if (!sameValue(f, sc[f.key], v)) changes.push({ id: sc.id, name: sc.name, key: f.key, label: f.label, from: sc[f.key], to: v });
      });
    });
    return { changes: changes, problems: problems };
  }

  function splitCSV(line) {
    var out = [], cur = '', q = false;
    for (var i = 0; i < line.length; i++) {
      var c = line[i];
      if (q) { if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c; }
      else if (c === '"') q = true; else if (c === ',') { out.push(cur); cur = ''; } else cur += c;
    }
    out.push(cur);
    return out;
  }

  /** Field-level differences between two versions of a scenario (for the change log). */
  function diffScenario(a, b) {
    var out = [];
    FIELDS.forEach(function (f) { if (!sameValue(f, a[f.key], b[f.key])) out.push({ key: f.key, label: f.label, from: a[f.key], to: b[f.key] }); });
    return out;
  }

  var api = {
    SCHEMA: SCHEMA, MONTHS: MONTHS, DEFAULT_RULES: DEFAULT_RULES, FIELDS: FIELDS, EVENT_KEYS: EVENT_KEYS, EVENT_NAMES: EVENT_NAMES,
    mi: mi, yearOf: yearOf, monthOf: monthOf, quarterOf: quarterOf, monthLabel: monthLabel, quarterLabel: quarterLabel,
    longMonth: longMonth, toYM: toYM, parseMonth: parseMonth, normMonth: normMonth, nextAcip: nextAcip,
    rulesOf: rulesOf, computeScenario: computeScenario, computeAll: computeAll, labels: labels,
    blankScenario: blankScenario, normalizeModel: normalizeModel, snapshot: snapshot, deltas: deltas,
    fieldByKey: fieldByKey, fieldValueText: fieldValueText, coerce: coerce, sameValue: sameValue,
    toTSV: toTSV, parseUpdates: parseUpdates, diffScenario: diffScenario
  };

  root.LT = root.LT || {};
  root.LT.engine = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
