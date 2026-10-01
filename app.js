/* global Office */
/*
 * Live Timeline — task pane UI.
 * In PowerPoint it edits the model saved in the selected slide and redraws the slide;
 * opened in a browser it runs the same pane against a preview of the slide.
 */
(function () {
  'use strict';
  var LT = window.LT, E = LT.engine, Lay = LT.layout;

  var S = {
    mode: 'preview', model: null, cal: null, calSource: '', calReason: '',
    slideId: null, slideIndex: null, names: null,
    history: [], live: true, dirty: false, busy: false, again: false, timer: 0, saveTimer: 0,
    backdrop: null, backdropSvg: null, asOf: null, proposed: null, legacy: [], open: { scenarios: true }
  };

  // ---------- small helpers ----------
  /** @param {string} id @returns {any} */
  function $(id) { return document.getElementById(id); }
  function h(tag, attrs, kids) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k === 'value') el.value = v;
      else if (k === 'checked') el.checked = !!v;
      else if (k.slice(0, 2) === 'on') el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? '' : v);
    });
    (kids || []).forEach(function (c) {
      if (c === null || c === undefined || c === false) return;
      el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return el;
  }
  function store(k, v) {
    try { if (v === undefined) return window.localStorage.getItem(k); window.localStorage.setItem(k, v); } catch (e) { /* storage unavailable */ }
    return null;
  }
  function fmtTime(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    return d.toLocaleDateString(undefined, { day: 'numeric', month: 'short' }) + ' ' + d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
  }
  function short(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n - 1) + '…' : s; }
  function findSc(id, m) { var list = (m || S.model).scenarios; for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i]; return null; }
  function scName(id) { if (id === 'rules') return 'Rules'; if (id === 'baseline') return 'Baseline'; var sc = findSc(id); return sc ? sc.name : id; }
  function monthText(v) { var i = E.parseMonth(v); return i === null ? (v || '') : E.longMonth(i); }

  function setStatus(msg, kind) { var el = $('status'); el.textContent = msg || ''; el.className = 'status' + (kind ? ' ' + kind : ''); }
  function showError(e) { if (window.console) console.error(e); setStatus('Problem: ' + ((e && e.message) ? e.message : String(e)), 'error'); }

  // ---------- start-up ----------
  function boot() {
    $('by').value = store('ltl.by') || '';
    $('by').addEventListener('change', function () { store('ltl.by', $('by').value.trim()); });
    $('btnUpdate').addEventListener('click', function () { if (S.mode === 'ppt') drawSlide(); else renderStage(); });
    $('btnUndo').addEventListener('click', undo);
    $('chkLive').addEventListener('change', function () { S.live = $('chkLive').checked; if (S.live && S.dirty) drawSlide(); markDirty(); });
    $('stage').addEventListener('click', function (ev) {
      var t = /** @type {any} */ (ev.target);
      var g = t && t.closest ? t.closest('g[data-key]') : null;
      if (!g) return;
      var id = g.getAttribute('data-key').split('.')[0];
      var card = $('card-' + id);
      if (card) { card.scrollIntoView({ behavior: 'smooth', block: 'start' }); card.classList.remove('flash'); void card.offsetWidth; card.classList.add('flash'); }
    });
    if (window.Office && typeof Office.onReady === 'function') {
      Office.onReady(function (info) {
        if (info && info.host === Office.HostType.PowerPoint) startPowerPoint(); else startPreview();
      });
    } else startPreview();
  }

  function startPreview() {
    S.mode = 'preview';
    document.body.classList.add('mode-preview');
    var P = window.LT_PREVIEW || LT.sample;
    S.model = E.normalizeModel(P.model);
    S.cal = P.cal;
    S.names = P.names || null;
    S.backdrop = P.backdrop || null;
    S.backdropSvg = P.backdrop ? null : LT.svg.tableBackdrop(P.cal, P.names, P.title);
    S.asOf = P.asOf || null;
    if (!S.model.baseline && P.baselineLabel) S.model.baseline = E.snapshot(S.model, P.baselineLabel);
    S.initial = JSON.stringify(S.model);
    renderAll();
    setStatus(P.statusText || 'Preview: the add-in’s engine and task pane, drawing on a picture of the slide.', 'ok');
  }

  function startPowerPoint() {
    S.mode = 'ppt';
    document.body.classList.add('mode-ppt');
    try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged); } catch (e) { /* older hosts */ }
    if (!LT.ppt.supports('1.4')) {
      setStatus('This version of PowerPoint is too old for drawing shapes (needs Microsoft 365, Version 2207 or later).', 'error');
      return;
    }
    loadSlide(null);
  }

  function loadSlide(id) {
    setStatus('Reading the slide…', 'busy');
    return LT.ppt.load(id).then(function (r) {
      S.slideId = r.slideId;
      S.slideIndex = r.index;
      S.model = r.model ? E.normalizeModel(r.model) : null;
      S.cal = r.cal || (r.model && r.model.calibration) || null;
      S.calSource = r.cal ? 'table' : (S.cal ? 'saved' : 'none');
      S.calReason = r.reason || '';
      S.names = r.cal ? r.cal.names : null;
      S.history = [];
      S.legacy = [];
      S.proposed = null;
      renderAll();
      var where = S.slideIndex !== null ? 'Slide ' + (S.slideIndex + 1) : 'This slide';
      if (!S.model) setStatus(where + ' has no live timeline yet.', S.cal ? 'ok' : 'warn');
      else if (!S.cal) setStatus(where + ': ' + (S.calReason || 'cannot find the timeline table.'), 'warn');
      else setStatus(where + ' · ' + S.model.scenarios.length + ' scenarios · ' + (S.calSource === 'table' ? 'measured from the table' : 'using the saved layout'), 'ok');
    }).catch(showError);
  }

  var selTimer = 0;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(function () {
      LT.ppt.selectedSlideId().then(function (id) {
        if (id && id !== S.slideId && !S.busy) loadSlide(id);
      }).catch(function () { /* ignore */ });
    }, 400);
  }

  // ---------- changes, history, drawing ----------
  function entry(scope, what, from, to) {
    return { at: new Date().toISOString(), by: ($('by').value || '').trim(), id: scope, what: what, from: from, to: to, source: ($('src').value || '').trim() };
  }

  function change(mutate, entries, rerender) {
    S.history.push(JSON.stringify(S.model));
    if (S.history.length > 60) S.history.shift();
    mutate(S.model);
    (entries || []).forEach(function (e) {
      // Several clicks on the same stepper within two minutes become one log entry (first "from", latest "to").
      var log = S.model.changelog, last = log[log.length - 1];
      if (last && last.id === e.id && last.what === e.what && last.by === e.by && last.source === e.source && Date.parse(e.at) - Date.parse(last.at) < 120000) {
        last.to = e.to; last.at = e.at;
        if (last.from === last.to) log.pop();
      } else log.push(e);
    });
    if (S.model.changelog.length > 400) S.model.changelog = S.model.changelog.slice(-400);
    if (rerender) renderAll(); else refreshComputed();
    scheduleDraw();
  }

  function undo() {
    if (!S.history.length) return;
    S.model = JSON.parse(S.history.pop());
    renderAll();
    scheduleDraw(true);
    setStatus('Undid the last change.', 'ok');
  }

  function scheduleDraw(force) {
    if (S.mode === 'preview') { renderStage(); return; }
    if (!S.live && !force) {
      S.dirty = true; markDirty();
      clearTimeout(S.saveTimer);
      S.saveTimer = setTimeout(function () { LT.ppt.save(S.slideId, S.model).catch(showError); }, 600);
      return;
    }
    clearTimeout(S.timer);
    S.timer = setTimeout(drawSlide, 250);
  }

  function drawSlide() {
    if (S.mode !== 'ppt' || !S.model) return;
    if (!S.cal) { setStatus(S.calReason || 'Cannot find the timeline table on this slide.', 'warn'); return; }
    if (S.busy) { S.again = true; return; }
    S.busy = true;
    setStatus('Updating the slide…', 'busy');
    LT.ppt.redraw(S.slideId, S.model, S.cal, {}).then(function (r) {
      S.dirty = false; markDirty();
      setStatus('Slide updated · ' + r.count + ' shapes' + (r.warnings.length ? ' · see warnings' : ''), r.warnings.length ? 'warn' : 'ok');
      renderGlobalWarnings(r.warnings);
    }).catch(showError).then(function () {
      S.busy = false;
      if (S.again) { S.again = false; drawSlide(); }
    });
  }

  function markDirty() {
    var b = $('btnUpdate');
    b.classList.toggle('attention', S.mode === 'ppt' && S.dirty);
    b.textContent = (S.mode === 'ppt' && S.dirty) ? 'Update slide •' : 'Update slide';
  }

  function renderStage() {
    var st = $('stage');
    if (S.mode !== 'preview' || !S.model || !S.cal) { st.innerHTML = ''; return; }
    var built = Lay.build(S.model, S.cal, { asOf: S.asOf || undefined });
    st.innerHTML = '<div class="slide">' + LT.svg.render(built.specs, { backdrop: S.backdrop, backdropSvg: S.backdropSvg }) + '</div>';
    renderGlobalWarnings(built.warnings);
  }

  function renderGlobalWarnings(list) {
    var el = $('warnings');
    if (!el) return;
    el.innerHTML = '';
    var outside = (list || []).filter(function (w) { return /outside the timeline|no matching row/.test(w); });
    outside.forEach(function (w) { el.appendChild(h('div', { class: 'note warn', text: w })); });
  }

  // ---------- inputs ----------
  function commitField(id, key, raw, opts) {
    var sc = findSc(id), f = E.fieldByKey(key);
    if (!sc || !f) return;
    var v = E.coerce(f, raw);
    if (f.type === 'month' && v && E.parseMonth(v) === null) { setStatus('Could not read "' + raw + '" as a month. Try Apr 2027 or 2027-04.', 'warn'); return false; }
    if (E.sameValue(f, sc[key], v)) return true;
    var e = entry(id, f.label, E.fieldValueText(f, sc[key]), E.fieldValueText(f, v));
    var restructure = (key === 'ssd' || key === 'submission') && (!sc.ssd !== !(key === 'ssd' ? v : sc.ssd) || !sc.submission !== !(key === 'submission' ? v : sc.submission));
    change(function (m) { findSc(id, m)[key] = v; }, [e], restructure || (opts && opts.rerender));
    return true;
  }

  function textInput(id, key, val, ph) {
    return h('input', { type: 'text', class: 'in', value: val || '', placeholder: ph || '', onchange: function (ev) { commitField(id, key, ev.target.value); } });
  }

  function monthInput(id, key, val, ph) {
    var inp = h('input', { type: 'text', class: 'in month', value: monthText(val), placeholder: ph || 'e.g. Apr 2027', 'aria-label': E.fieldByKey(key).label });
    inp.addEventListener('change', function () {
      var ok = commitField(id, key, inp.value);
      inp.classList.toggle('bad', ok === false);
      if (ok !== false) inp.value = monthText(findSc(id)[key]);
    });
    function step(d) {
      var sc = findSc(id), cur = E.parseMonth(sc[key]);
      var now = new Date();
      var base = cur === null ? E.mi(now.getFullYear(), now.getMonth() + 1) : cur + d;
      commitField(id, key, E.toYM(base));
      inp.value = monthText(findSc(id)[key]);
      inp.classList.remove('bad');
    }
    return h('div', { class: 'stepper' }, [
      h('button', { class: 'step', title: 'One month earlier', 'aria-label': 'One month earlier', onclick: function () { step(-1); }, text: '‹' }),
      inp,
      h('button', { class: 'step', title: 'One month later', 'aria-label': 'One month later', onclick: function () { step(1); }, text: '›' })
    ]);
  }

  function numInput(id, key, val, ph, suffix) {
    var inp = h('input', { type: 'text', inputmode: 'numeric', class: 'in num', value: (val === '' || val === null || val === undefined) ? '' : String(val), placeholder: ph || '', 'aria-label': E.fieldByKey(key).label });
    inp.addEventListener('change', function () { commitField(id, key, inp.value); inp.value = String(findSc(id)[key]); });
    function step(d) {
      var sc = findSc(id), cur = Number(sc[key] === '' ? (ph || 0) : sc[key]) || 0;
      commitField(id, key, String(Math.max(0, cur + d)));
      inp.value = String(findSc(id)[key]);
    }
    return h('div', { class: 'stepper' }, [
      h('button', { class: 'step', 'aria-label': 'Less', onclick: function () { step(-1); }, text: '‹' }),
      inp,
      h('button', { class: 'step', 'aria-label': 'More', onclick: function () { step(1); }, text: '›' }),
      suffix ? h('span', { class: 'suffix', text: suffix }) : null
    ]);
  }

  function checkbox(id, key, val, text) {
    return h('label', { class: 'check' }, [
      h('input', { type: 'checkbox', checked: !!val, onchange: function (ev) { commitField(id, key, ev.target.checked ? 'Y' : 'N'); } }),
      ' ' + text
    ]);
  }

  function row(label, control, hint) {
    return h('div', { class: 'row' }, [h('label', { class: 'lbl', text: label }), h('div', { class: 'ctl' }, [control, hint ? h('div', { class: 'hint', text: hint }) : null])]);
  }

  function shiftScenario(id, d) {
    var sc = findSc(id);
    var key = sc.ssd ? 'ssd' : sc.submission ? 'submission' : sc.approval ? 'approval' : null;
    if (!key) return;
    var from = E.parseMonth(sc[key]);
    change(function (m) {
      var x = findSc(id, m);
      ['ssd', 'submission', 'approval'].forEach(function (k) { var i = E.parseMonth(x[k]); if (i !== null) x[k] = E.toYM(i + d); });
    }, [entry(id, 'Shifted whole scenario', E.longMonth(from), E.longMonth(from + d))], true);
  }

  // ---------- views ----------
  function section(key, title, kids, extraHead) {
    var d = h('details', { class: 'sec', id: 'sec-' + key, open: S.open[key] ? true : null }, [
      h('summary', {}, [h('span', { text: title }), extraHead || null])
    ].concat(kids));
    d.addEventListener('toggle', function () { S.open[key] = d.open; });
    return d;
  }

  function renderAll() {
    var body = $('content');
    body.innerHTML = '';
    $('btnUndo').disabled = !S.history.length;
    $('btnUpdate').disabled = !S.model;
    markDirty();
    if (!S.model) { body.appendChild(setupView()); renderStage(); return; }
    body.appendChild(h('div', { id: 'warnings' }));
    var list = h('div', { id: 'scenarios' });
    S.model.scenarios.forEach(function (sc) { list.appendChild(scenarioCard(sc)); });
    body.appendChild(list);
    body.appendChild(teamView());
    body.appendChild(baselineView());
    body.appendChild(rulesView());
    body.appendChild(logView());
    body.appendChild(slideView());
    refreshComputed();
    renderStage();
  }

  function scenarioCard(sc) {
    var id = sc.id;
    var fixedSub = !!sc.submission && !sc.ssd;
    var main = h('div', { class: 'grid' }, fixedSub ? [
      row('Submission', monthInput(id, 'submission', sc.submission)),
      row('Review', checkbox(id, 'priorityReview', sc.priorityReview, 'Priority review / PRV'))
    ] : [
      row('Study', textInput(id, 'studyName', sc.studyName, 'e.g. Ph3 rVE')),
      row('Start (SSD)', monthInput(id, 'ssd', sc.ssd)),
      row('Length', numInput(id, 'studyMonths', sc.studyMonths, '', 'months to PCD')),
      row('Review', checkbox(id, 'priorityReview', sc.priorityReview, 'Priority review / PRV'))
    ]);
    var r = E.rulesOf(S.model);
    var moreKids = [
      h('div', { class: 'grid' }, [
        fixedSub ? row('Study', textInput(id, 'studyName', sc.studyName, 'optional')) : null,
        fixedSub ? row('Start (SSD)', monthInput(id, 'ssd', sc.ssd, 'none')) : null,
        row('Topline', numInput(id, 'readoutLag', sc.readoutLag, String(r.readoutLag), 'months after PCD')),
        row('Submission', numInput(id, 'submissionLag', sc.submissionLag, String(r.submissionLag), 'months after topline')),
        fixedSub ? null : row('Fixed submission', monthInput(id, 'submission', sc.submission, 'none'), 'Overrides the study-based date'),
        row('Review length', numInput(id, 'reviewMonths', sc.reviewMonths, String(sc.priorityReview ? r.reviewPriority : r.reviewStandard), 'months'), 'Blank uses the rule'),
        row('Fixed approval', monthInput(id, 'approval', sc.approval, 'none')),
        row('Note on slide', textInput(id, 'note', sc.note, 'e.g. Re-file with new data')),
        row('ACIP label', h('label', { class: 'check' }, [
          h('input', { type: 'checkbox', checked: !!sc.highlightAcip, onchange: function (ev) {
            var v = ev.target.checked;
            change(function (m) { findSc(id, m).highlightAcip = v; }, [entry(id, 'Highlight ACIP', v ? 'N' : 'Y', v ? 'Y' : 'N')]);
          } }), ' Highlight in red'
        ]))
      ])
    ];
    var more = h('details', { class: 'more', open: S.open['more-' + id] ? true : null }, [h('summary', { text: 'More inputs' })].concat(moreKids));
    more.addEventListener('toggle', function () { S.open['more-' + id] = more.open; });

    return h('section', { class: 'card', id: 'card-' + id }, [
      h('div', { class: 'card-head' }, [
        h('div', { class: 'card-title', title: sc.name, text: sc.name || id }),
        h('div', { class: 'shift' }, [
          h('button', { class: 'chip', title: 'Move this scenario 3 months earlier', onclick: function () { shiftScenario(id, -3); }, text: '−3 mo' }),
          h('button', { class: 'chip', title: 'Move this scenario 3 months later', onclick: function () { shiftScenario(id, 3); }, text: '+3 mo' })
        ])
      ]),
      h('div', { class: 'chain', id: 'chain-' + id }),
      h('div', { class: 'card-warn', id: 'warn-' + id }),
      main,
      more
    ]);
  }

  function refreshComputed() {
    if (!S.model) return;
    var rules = E.rulesOf(S.model);
    var dls = E.deltas(S.model);
    S.model.scenarios.forEach(function (sc) {
      var r = E.computeScenario(sc, rules);
      var L = E.labels(r.events, rules);
      var chain = $('chain-' + sc.id);
      if (chain) {
        chain.innerHTML = '';
        var dl = dls[sc.id];
        [['ssd', 'p-ssd'], ['pcd', 'p-pcd'], ['d', 'p-d', 'D '], ['s', 'p-s', 'S '], ['fda', 'p-fda'], ['acip', 'p-acip']].forEach(function (k, i, arr) {
          if (!L[k[0]]) return;
          var kids = [h('span', { text: (k[2] || '') + L[k[0]] })];
          var dv = dl ? dl.delta[k[0]] : null;
          if (k[0] === 'fda' && dv) kids.push(h('b', { class: dv > 0 ? 'late' : 'early', text: ' ' + (dv > 0 ? '+' : '−') + Math.abs(dv) + ' mo' }));
          chain.appendChild(h('span', { class: 'pill ' + k[1] }, kids));
        });
      }
      var w = $('warn-' + sc.id);
      if (w) { w.innerHTML = ''; r.warnings.forEach(function (t) { w.appendChild(h('div', { class: 'note warn', text: t })); }); }
    });
    var bt = $('baseTable');
    if (bt) fillBaselineTable(bt);
    var lc = $('logCount');
    if (lc) lc.textContent = S.model.changelog.length ? String(S.model.changelog.length) : '';
    var ll = $('logList');
    if (ll) fillLog(ll);
    $('btnUndo').disabled = !S.history.length;
  }

  // ---- team updates ----
  function teamView() {
    var paste = h('textarea', { id: 'paste', class: 'in', rows: '4', placeholder: 'Paste the team’s edited table here (straight from Excel, Teams or email)' });
    var copyOut = h('textarea', { id: 'copyOut', class: 'in mono hidden', rows: '4', readonly: true });
    var proposals = h('div', { id: 'proposals' });
    return section('team', 'Team updates', [
      h('p', { class: 'help', text: 'Send the inputs as a table, let the team edit dates, lengths or PRV (Y/N), paste it back and apply only the changes you accept. Each change is logged with its source.' }),
      h('div', { class: 'btns' }, [
        h('button', { class: 'btn', onclick: copyTable, text: 'Copy input table' }),
        h('button', { class: 'btn', onclick: copyTemplateNote, text: 'Copy note for the team' })
      ]),
      copyOut,
      paste,
      h('div', { class: 'btns' }, [h('button', { class: 'btn', onclick: function () { reviewPaste(paste.value); }, text: 'Review changes' })]),
      proposals
    ]);
  }

  function copyText(text, fallbackEl, what) {
    function fallback() {
      fallbackEl.classList.remove('hidden');
      fallbackEl.value = text;
      fallbackEl.focus(); fallbackEl.select();
      setStatus('Select-all is done: press Ctrl+C (⌘C on Mac) to copy the ' + what + '.', 'warn');
    }
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () { fallbackEl.classList.add('hidden'); setStatus('Copied the ' + what + ' to the clipboard.', 'ok'); }, fallback);
      } else fallback();
    } catch (e) { fallback(); }
  }
  function copyTable() { copyText(E.toTSV(S.model), $('copyOut'), 'input table'); }
  function copyTemplateNote() {
    var text = 'Timeline check-in: please update any inputs that have changed in the table below (dates like "Apr 2027", lengths in months, PRV as Y/N), add your source, and send it back.\n\n' + E.toTSV(S.model);
    copyText(text, $('copyOut'), 'note and table');
  }

  function reviewPaste(text) {
    var res = E.parseUpdates(S.model, text);
    S.proposed = res;
    var box = $('proposals');
    box.innerHTML = '';
    res.problems.forEach(function (p) { box.appendChild(h('div', { class: 'note warn', text: p })); });
    if (!res.changes.length) { box.appendChild(h('div', { class: 'note', text: res.problems.length ? 'No changes could be read.' : 'No differences from the current inputs.' })); return; }
    var list = h('div', { class: 'proposals' });
    res.changes.forEach(function (c, i) {
      var f = E.fieldByKey(c.key);
      list.appendChild(h('label', { class: 'proposal' }, [
        h('input', { type: 'checkbox', checked: true, 'data-i': String(i) }),
        h('span', {}, [h('b', { text: short(scName(c.id), 26) }), ' · ' + c.label + ': ', h('s', { text: E.fieldValueText(f, c.from) || 'blank' }), ' → ', h('b', { text: E.fieldValueText(f, c.to) || 'blank' })])
      ]));
    });
    box.appendChild(list);
    box.appendChild(h('div', { class: 'btns' }, [h('button', { class: 'btn primary', onclick: applyProposals, text: 'Apply checked changes' })]));
  }

  function applyProposals() {
    if (!S.proposed) return;
    var picked = Array.prototype.slice.call(document.querySelectorAll('#proposals input[type=checkbox]'))
      .filter(function (c) { return c.checked; }).map(function (c) { return S.proposed.changes[Number(c.getAttribute('data-i'))]; });
    if (!picked.length) return;
    var src = ($('src').value || '').trim() || 'Team update (pasted table)';
    var entries = picked.map(function (c) {
      var f = E.fieldByKey(c.key), e = entry(c.id, c.label, E.fieldValueText(f, c.from), E.fieldValueText(f, c.to));
      e.source = src;
      return e;
    });
    change(function (m) { picked.forEach(function (c) { findSc(c.id, m)[c.key] = c.to; }); }, entries, true);
    S.proposed = null;
    setStatus('Applied ' + picked.length + ' change' + (picked.length > 1 ? 's' : '') + ' from the team table.', 'ok');
  }

  // ---- baseline ----
  function baselineView() {
    var b = S.model.baseline;
    var label = h('input', { id: 'baseLabel', class: 'in', type: 'text', placeholder: 'e.g. Oct 2026 steering committee' });
    return section('baseline', 'Baseline and slippage', [
      h('p', { class: 'help', text: b ? 'Baseline “' + b.label + '”, saved ' + fmtTime(b.savedAt) + '.' : 'No baseline saved yet. Save one before a review to see what moved.' }),
      h('label', { class: 'check' }, [
        h('input', { type: 'checkbox', checked: !!S.model.showSlippage, disabled: b ? null : true, onchange: function (ev) {
          var v = ev.target.checked;
          change(function (m) { m.showSlippage = v; }, []);
        } }), ' Show slippage on the slide (hollow star at the baseline approval, ± months next to FDA)'
      ]),
      h('div', { id: 'baseTable' }),
      h('div', { class: 'row-inline' }, [label, h('button', { class: 'btn', onclick: function () {
        var name = label.value.trim() || ('Baseline ' + new Date().toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' }));
        change(function (m) { m.baseline = E.snapshot(m, name); }, [entry('baseline', 'Saved baseline', '', name)], true);
      }, text: 'Save current as baseline' })])
    ]);
  }

  function fillBaselineTable(el) {
    el.innerHTML = '';
    var dls = E.deltas(S.model);
    if (!S.model.baseline) return;
    var rules = E.rulesOf(S.model);
    var t = h('table', { class: 'tbl' }, [h('tr', {}, [h('th', { text: 'Scenario' }), h('th', { text: 'FDA, baseline' }), h('th', { text: 'Now' }), h('th', { text: 'Δ' })])]);
    S.model.scenarios.forEach(function (sc) {
      var r = E.computeScenario(sc, rules), dl = dls[sc.id];
      var base = dl ? E.parseMonth(dl.base.fda) : null, now = r.events.fda;
      var d = (base !== null && now !== null) ? now - base : null;
      t.appendChild(h('tr', {}, [
        h('td', { text: short(sc.name, 20), title: sc.name }),
        h('td', { text: base !== null ? E.monthLabel(base) : '–' }),
        h('td', { text: now !== null ? E.monthLabel(now) : '–' }),
        h('td', { class: d > 0 ? 'late' : d < 0 ? 'early' : '', text: d === null ? '–' : d === 0 ? '0' : (d > 0 ? '+' : '−') + Math.abs(d) + ' mo' })
      ]));
    });
    el.appendChild(t);
  }

  // ---- rules ----
  function rulesView() {
    var r = E.rulesOf(S.model);
    function ruleNum(key, label, hint) {
      var inp = h('input', { type: 'text', inputmode: 'numeric', class: 'in num', value: String(r[key]) });
      inp.addEventListener('change', function () {
        var v = Math.round(Number(inp.value));
        if (!isFinite(v) || v < 0) { inp.value = String(E.rulesOf(S.model)[key]); return; }
        var old = E.rulesOf(S.model)[key];
        if (old === v) return;
        change(function (m) { m.rules[key] = v; }, [entry('rules', label, String(old), String(v))], true);
      });
      return row(label, h('div', { class: 'stepper' }, [inp, h('span', { class: 'suffix', text: hint })]));
    }
    var chips = h('div', { class: 'chips' }, E.MONTHS.map(function (mn, i) {
      var on = r.acipMonths.indexOf(i + 1) >= 0;
      return h('button', { class: 'chip' + (on ? ' on' : ''), 'aria-pressed': on ? 'true' : 'false', text: mn.charAt(0) + mn.slice(1, 3), onclick: function () {
        var cur = E.rulesOf(S.model).acipMonths.slice(), k = cur.indexOf(i + 1);
        if (k >= 0) cur.splice(k, 1); else cur.push(i + 1);
        cur.sort(function (a, b) { return a - b; });
        var label = function (a) { return a.map(function (x) { return E.MONTHS[x - 1]; }).join(', '); };
        change(function (m) { m.rules.acipMonths = cur; }, [entry('rules', 'ACIP meeting months', label(E.rulesOf(S.model).acipMonths), label(cur))], true);
      } });
    }));
    var asOf = h('input', { type: 'date', class: 'in', value: S.model.asOf || '' });
    asOf.addEventListener('change', function () { var v = asOf.value; change(function (m) { m.asOf = v; }, [entry('rules', 'Today line date', S.model.asOf || 'today', v || 'today')]); });
    var pl = h('select', { class: 'in' }, [h('option', { value: 'quarter', text: 'Quarter (4Q ’27)' }), h('option', { value: 'month', text: 'Month (Nov ’27)' })]);
    pl.value = r.pointLabels;
    pl.addEventListener('change', function () { var v = pl.value; change(function (m) { m.rules.pointLabels = v; }, [entry('rules', 'D and S labels', r.pointLabels, v)]); });
    return section('rules', 'Rules', [
      h('p', { class: 'help', text: 'Applied to every scenario. A blank scenario input uses these defaults.' }),
      h('div', { class: 'grid' }, [
        ruleNum('reviewPriority', 'Priority review', 'months, submission → approval'),
        ruleNum('reviewStandard', 'Standard review', 'months, submission → approval'),
        ruleNum('readoutLag', 'Topline', 'months after PCD'),
        ruleNum('submissionLag', 'Submission', 'months after topline'),
        row('ACIP meets in', chips),
        ruleNum('acipMinGap', 'ACIP vote', 'months after approval, at least'),
        row('D and S labels', pl),
        row('Today line', h('div', {}, [
          h('label', { class: 'check' }, [h('input', { type: 'checkbox', checked: S.model.showToday !== false, onchange: function (ev) { var v = ev.target.checked; change(function (m) { m.showToday = v; }, []); } }), ' Show']),
          asOf
        ]), 'Date blank = today')
      ])
    ]);
  }

  // ---- change log ----
  function logView() {
    return section('log', 'Change log', [
      h('div', { class: 'btns' }, [h('button', { class: 'btn', onclick: function () {
        var rows = [['When', 'By', 'Scenario', 'What', 'From', 'To', 'Source'].join('\t')].concat(S.model.changelog.map(function (e) {
          return [e.at, e.by, scName(e.id), e.what, e.from, e.to, e.source].map(function (x) { return String(x || '').replace(/[\t\n]/g, ' '); }).join('\t');
        }));
        copyText(rows.join('\n'), $('logOut'), 'change log');
      }, text: 'Copy log' })]),
      h('textarea', { id: 'logOut', class: 'in mono hidden', rows: '4', readonly: true }),
      h('ol', { id: 'logList', class: 'log' })
    ], h('span', { id: 'logCount', class: 'count' }));
  }

  function fillLog(el) {
    el.innerHTML = '';
    var items = S.model.changelog.slice(-80).reverse();
    if (!items.length) { el.appendChild(h('li', { class: 'empty', text: 'No changes yet. Edits made here are recorded with who made them and the source.' })); return; }
    items.forEach(function (e) {
      el.appendChild(h('li', {}, [
        h('div', { class: 'meta', text: fmtTime(e.at) + (e.by ? ' · ' + e.by : '') + ' · ' + short(scName(e.id), 30) }),
        h('div', {}, [e.what + ': ', e.from ? h('s', { text: e.from }) : null, e.from ? ' → ' : '', h('b', { text: e.to || 'blank' })]),
        e.source ? h('div', { class: 'src', text: e.source }) : null
      ]));
    });
  }

  // ---- slide setup ----
  function calText() {
    if (!S.cal) return S.calReason || 'No timeline table found.';
    var ax = S.cal.axis;
    return E.quarterLabel(ax.startMonth) + ' to ' + E.quarterLabel(ax.endMonth - 1) + ' · ' + S.cal.lanes.length + ' scenario rows · ' +
      (S.calSource === 'table' ? 'measured from the table' : S.calSource === 'saved' ? 'saved layout (table could not be measured)' : 'example layout');
  }

  function slideView() {
    if (S.mode === 'preview') {
      return section('slide', 'About this preview', [
        h('p', { class: 'help', text: 'This page runs the add-in’s engine and task pane against a picture of the slide. Inside PowerPoint the same pane draws native, editable shapes on the slide and saves its inputs in the slide itself.' }),
        h('p', { class: 'help', text: 'Timeline: ' + calText() }),
        h('div', { class: 'btns' }, [h('button', { class: 'btn', onclick: function () {
          S.model = E.normalizeModel(JSON.parse(S.initial)); S.history = []; renderAll(); setStatus('Preview reset to the original inputs.', 'ok');
        }, text: 'Reset preview' })])
      ]);
    }
    var legacyBox = h('div', { id: 'legacy' });
    return section('slide', 'Slide setup', [
      h('p', { class: 'help', text: 'Timeline: ' + calText() }),
      h('div', { class: 'btns' }, [
        h('button', { class: 'btn', onclick: function () { loadSlide(S.slideId); }, text: 'Re-read slide' }),
        h('button', { class: 'btn', onclick: findLegacy, disabled: S.cal ? null : true, text: 'Find hand-drawn shapes' }),
        h('button', { class: 'btn danger', onclick: removeTimeline, text: 'Remove live timeline' })
      ]),
      legacyBox,
      h('p', { class: 'help small', text: 'The add-in only touches shapes whose names start with “LTL|” (see Home › Arrange › Selection Pane). Inputs, baseline and change log are saved inside this slide.' })
    ]);
  }

  function findLegacy() {
    LT.ppt.findLegacy(S.slideId, S.cal).then(function (list) {
      S.legacy = list;
      var box = $('legacy');
      box.innerHTML = '';
      if (!list.length) { box.appendChild(h('div', { class: 'note', text: 'No hand-drawn timeline shapes found inside the table area.' })); return; }
      box.appendChild(h('div', { class: 'note warn', text: list.length + ' shapes are now selected on the slide (bars, markers and labels drawn by hand). Check the selection, then delete them so only the live timeline remains.' }));
      box.appendChild(h('div', { class: 'btns' }, [h('button', { class: 'btn danger', onclick: function () {
        LT.ppt.deleteShapes(S.slideId, S.legacy.map(function (s) { return s.id; })).then(function (n) {
          S.legacy = []; box.innerHTML = '';
          setStatus('Deleted ' + n + ' hand-drawn shapes.', 'ok');
          if (S.model) drawSlide();
        }).catch(showError);
      }, text: 'Delete these ' + list.length + ' shapes' })]));
    }).catch(showError);
  }

  function removeTimeline() {
    LT.ppt.removeAll(S.slideId).then(function (n) {
      setStatus('Removed the live timeline (' + n + ' shape groups) and its saved inputs from this slide.', 'ok');
      return loadSlide(S.slideId);
    }).catch(showError);
  }

  function setupView() {
    var kids = [];
    if (S.mode === 'ppt') {
      if (S.cal) {
        kids.push(h('p', { class: 'help', text: 'This slide has a timeline table (' + calText() + ').' }));
        kids.push(h('ul', { class: 'rows' }, (S.names || []).map(function (n, i) { return h('li', { text: n || ('Row ' + (i + 1)) }); })));
        kids.push(h('p', { class: 'help', text: 'If the timeline was drawn by hand, remove those shapes first so they are not left under the live version.' }));
        kids.push(h('div', { class: 'btns' }, [h('button', { class: 'btn', onclick: findLegacy, text: 'Find hand-drawn shapes' })]));
        kids.push(h('div', { id: 'legacy' }));
        kids.push(h('div', { class: 'btns' }, [
          h('button', { class: 'btn primary', onclick: function () { startModel(false); }, text: 'Start a live timeline for these rows' }),
          h('button', { class: 'btn', onclick: function () { startModel(true); }, text: 'Use example inputs' })
        ]));
      } else {
        kids.push(h('p', { class: 'help', text: S.calReason || 'No timeline table found on this slide.' }));
        kids.push(h('p', { class: 'help', text: 'Select a slide with a table whose header has years and quarters (like the filing-scenarios slide), then press Re-read slide.' }));
        kids.push(h('div', { class: 'btns' }, [h('button', { class: 'btn', onclick: function () { loadSlide(null); }, text: 'Re-read slide' })]));
      }
    }
    return h('section', { class: 'card setup' }, [h('div', { class: 'card-title', text: 'Set up this slide' })].concat(kids));
  }

  function startModel(example) {
    var names = S.names && S.names.length ? S.names : S.cal.lanes.map(function (_, i) { return 'Scenario ' + (i + 1); });
    var scen = names.map(function (n, i) {
      var b = E.blankScenario(i);
      b.name = n || b.name;
      if (example && LT.sample.model.scenarios[i]) { var ex = LT.sample.model.scenarios[i]; Object.keys(ex).forEach(function (k) { if (k !== 'name' && k !== 'id' && k !== 'row') b[k] = ex[k]; }); }
      else { b.studyName = ''; b.studyMonths = ''; }
      return b;
    });
    S.model = E.normalizeModel({ scenarios: scen, rules: LT.sample.model.rules });
    S.model.changelog.push(entry('rules', 'Started live timeline', '', example ? 'example inputs' : 'blank inputs'));
    S.history = [];
    S.open.scenarios = true;
    renderAll();
    drawSlide();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
