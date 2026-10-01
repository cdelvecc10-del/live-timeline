/*
 * Live Timeline — layout engine.
 * Turns the computed schedule into a flat list of shape specs (points, slide coordinates) in the visual
 * language of the reference slide: green study bar, D / S circles, dashed FDA review box, gold approval star,
 * month labels for SSD / PCD / FDA / ACIP, quarter labels under D and S, and a dashed "Today" line.
 * The same specs are drawn by three renderers: PowerPoint (Office.js), SVG (browser preview) and python-pptx (tests).
 */
(function (root) {
  'use strict';
  var E = (root.LT && root.LT.engine) || (typeof require !== 'undefined' ? require('./engine.js') : null);

  // Arial (metric-compatible with Helvetica) advance widths, 1/1000 em, for ASCII 32..126.
  var W_REG = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584];
  var W_BOLD = [278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584];
  var W_EXTRA = { '‘': [222, 278], '’': [222, 278], '“': [333, 500], '”': [333, 500], '–': [556, 556], '—': [1000, 1000], '−': [584, 584], '→': [1000, 1000], 'Δ': [668, 668] };

  function textWidth(str, size, bold) {
    var t = 0, s = String(str || '');
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i), w;
      if (c >= 32 && c <= 126) w = (bold ? W_BOLD : W_REG)[c - 32];
      else if (W_EXTRA[s[i]]) w = W_EXTRA[s[i]][bold ? 1 : 0];
      else w = 556;
      t += w;
    }
    return t * size / 1000;
  }

  var DEFAULT_STYLE = {
    barH: 16, barFill: '#69B445', barText: '#FFFFFF', barTextSize: 9,
    labelColor: '#595959', labelSize: 9, labelH: 12,
    qColor: '#333333', qSize: 8,
    marker: 14, dFill: '#244EA2', sFill: '#E21860', markerLine: '#FFFFFF', markerLineW: 1, markerTextSize: 8,
    reviewH: 10, reviewFill: '#DDEFD3', reviewLine: '#69B445', reviewLineW: 0.75, reviewText: '#595959', reviewTextSize: 8,
    star: 16, starFill: '#FFC709', starLine: '#8A6A00', starLineW: 0.75,
    apColor: '#333333', apSize: 9, apLineH: 10.5, highlight: '#E21860',
    noteColor: '#595959', noteSize: 9,
    todayColor: '#F36633', todayLineW: 0.75, todayTagW: 36, todayTagH: 13, todaySize: 8,
    ghost: '#8C8C8C', later: '#C62828', earlier: '#2E7D32'
  };

  function styleOf(model) {
    var s = {}, k;
    for (k in DEFAULT_STYLE) s[k] = DEFAULT_STYLE[k];
    if (model && model.style) for (k in model.style) s[k] = model.style[k];
    return s;
  }

  function overlaps(a, b, pad) {
    pad = pad || 0;
    return a.x < b.x + b.w + pad && b.x < a.x + a.w + pad && a.y < b.y + b.h && b.y < a.y + a.h;
  }

  function asOfMonth(asOf) {
    var d;
    if (asOf instanceof Date) d = asOf;
    else if (typeof asOf === 'string' && /^\d{4}-\d{2}-\d{2}/.test(asOf)) {
      var p = asOf.split('-');
      d = new Date(+p[0], +p[1] - 1, +p[2].slice(0, 2));
    } else d = new Date();
    var dim = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    return E.mi(d.getFullYear(), d.getMonth() + 1) + (d.getDate() - 1) / dim;
  }

  /**
   * @param {any} model      timeline model (see engine.normalizeModel)
   * @param {any} cal        calibration: { axis:{startMonth,endMonth,left,right}, lanes:[{top,bottom}], gridTop, gridBottom }
   * @param {any} [opts]     { asOf: Date|'YYYY-MM-DD' }
   * @returns {{specs: Array<object>, warnings: Array<string>}}
   */
  function build(model, cal, opts) {
    opts = opts || {};
    var st = styleOf(model), rules = E.rulesOf(model);
    var ax = cal.axis;
    var mw = (ax.right - ax.left) / (ax.endMonth - ax.startMonth);
    var X = function (m) { return ax.left + (m - ax.startMonth) * mw; };
    var inAxis = function (m) { return m !== null && m >= ax.startMonth && m < ax.endMonth; };
    var layers = [[], [], [], [], [], []]; // 0 today line, 1 bars/boxes, 2 baseline ghosts, 3 markers, 4 labels, 5 today tag
    var warnings = [];
    var deltas = (model.showSlippage && model.baseline) ? E.deltas(model) : {};
    var axisText = E.quarterLabel(ax.startMonth) + ' to ' + E.quarterLabel(ax.endMonth - 1);

    function textSpec(key, name, x, y, w, h, runs, align, extra) {
      var s = { key: key, name: name, geom: 'text', x: x, y: y, w: w, h: h, fill: null, line: null, text: { runs: runs, align: align || 'l' } };
      if (extra) for (var k in extra) s[k] = extra[k];
      return s;
    }
    function run(t, size, color, bold) { return { t: t, size: size, color: color, bold: !!bold }; }

    (model.scenarios || []).forEach(function (sc, idx) {
      var lane = cal.lanes[(sc.row === undefined || sc.row === '') ? idx : Number(sc.row)];
      var tag = sc.id;
      var label = (sc.name || sc.id);
      if (!lane) { warnings.push(label + ': no matching row in the table'); return; }
      var cy = (lane.top + lane.bottom) / 2;
      var r = E.computeScenario(sc, rules);
      var ev = r.events;
      var L = E.labels(ev, rules);
      r.warnings.forEach(function (w) { warnings.push(label + ': ' + w); });
      var above = [];   // label boxes above the lane line
      var below = [];   // label boxes below the lane line
      var chainStart = null;
      var nm = function (role) { return 'LTL|' + tag + '|' + role; };

      ['ssd', 's', 'fda'].forEach(function (k) {
        if (ev[k] !== null && !inAxis(ev[k])) warnings.push(label + ': ' + E.EVENT_NAMES[k] + ' (' + E.monthLabel(ev[k]) + ') is outside the timeline (' + axisText + ')');
      });

      // ---- study bar, SSD and PCD labels ----
      if (ev.ssd !== null) {
        var x1 = Math.max(X(ev.ssd), ax.left), x2 = Math.min(X(ev.pcd + 1), ax.right);
        if (x2 > x1 + 0.5) {
          chainStart = x1;
          var bar = { key: tag + '.bar', name: nm('Study bar'), geom: 'roundRect', x: x1, y: cy - st.barH / 2, w: x2 - x1, h: st.barH, fill: st.barFill, line: null, text: null };
          var nameW = textWidth(sc.studyName, st.barTextSize, true);
          if (sc.studyName && nameW + 6 <= bar.w) bar.text = { runs: [run(sc.studyName, st.barTextSize, st.barText, true)], align: 'c' };
          else if (sc.studyName) above.push(textSpec(tag + '.barname', nm('Study name'), x1, cy - st.barH / 2 - 1.3 - st.labelH, nameW + 2, st.labelH, [run(sc.studyName, st.labelSize, '#333333', true)], 'l'));
          layers[1].push(bar);

          var ssdW = textWidth(L.ssd, st.labelSize) + 2;
          if (x1 - 3.3 - ssdW >= ax.left - 1) {
            layers[4].push(textSpec(tag + '.ssd', nm('SSD label'), x1 - 3.3 - ssdW, cy - st.labelH / 2, ssdW, st.labelH, [run(L.ssd, st.labelSize, st.labelColor)], 'r'));
            chainStart = x1 - 3.3 - ssdW;
          } else {
            above.push(textSpec(tag + '.ssd', nm('SSD label'), x1, cy - st.barH / 2 - 1.3 - st.labelH, ssdW, st.labelH, [run(L.ssd, st.labelSize, st.labelColor)], 'l'));
          }
          var pcdW = textWidth(L.pcd, st.labelSize) + 2;
          above.push(textSpec(tag + '.pcd', nm('PCD label'), x2 - pcdW, cy - st.barH / 2 - 1.3 - st.labelH, pcdW, st.labelH, [run(L.pcd, st.labelSize, st.labelColor)], 'r'));
        }
      }

      // ---- D and S markers ----
      var R = st.marker / 2;
      var xd = (ev.d !== null && inAxis(ev.d)) ? X(ev.d + 0.5) : null;
      var xs = (ev.s !== null && inAxis(ev.s)) ? X(ev.s + 0.5) : null;
      if (xd !== null && xs !== null && xs - xd < st.marker + 1) xd = xs - st.marker - 1;   // keep the circles apart
      var xf = (ev.fda !== null && inAxis(ev.fda)) ? X(ev.fda + 0.5) : null;

      // Baseline approval (when showing slippage): a hollow star where approval sat in the saved baseline.
      var dl = deltas[sc.id];
      var ghostX = null;
      if (dl && dl.delta.fda && xf !== null) {
        var bf = E.parseMonth(dl.base.fda);
        if (inAxis(bf)) {
          ghostX = X(bf + 0.5);
          layers[2].push({ key: tag + '.ghost', name: nm('Baseline approval'), geom: 'star5', x: ghostX - st.star / 2, y: cy - st.star / 2, w: st.star, h: st.star, fill: '#FFFFFF', line: { color: st.ghost, w: 1, dash: 'solid' }, text: null });
        }
      }

      // FDA review box from submission to approval; its label sits in the widest free stretch of the box.
      if (xs !== null && xf !== null && xf - xs > 2) {
        layers[1].push({ key: tag + '.review', name: nm('FDA review'), geom: 'rect', x: xs, y: cy - st.reviewH / 2, w: xf - xs, h: st.reviewH, fill: st.reviewFill, line: { color: st.reviewLine, w: st.reviewLineW, dash: 'dash' }, text: null });
        var free = [[xs + R + 1, xf - st.star / 2 - 1]];
        if (ghostX !== null && ghostX > free[0][0] && ghostX < free[0][1]) free = [[free[0][0], ghostX - st.star / 2 - 1], [ghostX + st.star / 2 + 1, free[0][1]]];
        free.sort(function (a, b) { return (b[1] - b[0]) - (a[1] - a[0]); });
        var rw = textWidth('FDA review', st.reviewTextSize) + 4;
        if (free[0][1] - free[0][0] >= rw) {
          var rc = (free[0][0] + free[0][1]) / 2;
          layers[4].push(textSpec(tag + '.reviewlbl', nm('FDA review label'), rc - rw / 2, cy - st.labelH / 2, rw, st.labelH, [run('FDA review', st.reviewTextSize, st.reviewText)], 'c'));
        }
      }

      [['d', xd, st.dFill, 'D', 'Data readout'], ['s', xs, st.sFill, 'S', 'Submission']].forEach(function (m) {
        var k = m[0], x = m[1];
        if (x === null) return;
        if (chainStart === null || x - R < chainStart) chainStart = Math.min(chainStart === null ? Infinity : chainStart, x - R);
        layers[3].push({ key: tag + '.' + k, name: nm(m[4]), geom: 'ellipse', x: x - R, y: cy - R, w: st.marker, h: st.marker, fill: m[2], line: { color: st.markerLine, w: st.markerLineW, dash: 'solid' }, text: { runs: [run(m[3], st.markerTextSize, '#FFFFFF', true)], align: 'c' } });
        var qt = L[k], qw = textWidth(qt, st.qSize) + 4;
        below.push(textSpec(tag + '.' + k + 'q', nm(m[4] + ' label'), x - qw / 2, cy + 9, qw, st.labelH, [run(qt, st.qSize, st.qColor)], 'c', { anchorX: x }));
      });

      // ---- approval star and approval / ACIP labels ----
      if (xf !== null) {
        if (chainStart === null) chainStart = xf - st.star / 2;
        layers[3].push({ key: tag + '.star', name: nm('FDA approval'), geom: 'star5', x: xf - st.star / 2, y: cy - st.star / 2, w: st.star, h: st.star, fill: st.starFill, line: { color: st.starLine, w: st.starLineW, dash: 'solid' }, text: null });
        var fdaRuns = [run(L.fda, st.apSize, st.apColor)];
        if (dl && dl.delta.fda) {
          var d = dl.delta.fda;
          fdaRuns.push(run(' ' + (d > 0 ? '+' : '−') + Math.abs(d) + ' mo', st.apSize, d > 0 ? st.later : st.earlier, true));
        }
        var lines = [fdaRuns];
        if (L.acip) lines.push([run(L.acip, st.apSize, sc.highlightAcip ? st.highlight : st.apColor)]);
        var widths = lines.map(function (rs) { return rs.reduce(function (a, q) { return a + textWidth(q.t, q.size, q.bold); }, 0) + 2; });
        var maxW = Math.max.apply(null, widths);
        var lx = xf + st.star / 2 + 0.5;
        if (ghostX !== null && ghostX > xf && ghostX - st.star / 2 < lx + maxW) lx = ghostX + st.star / 2 + 1;
        var totalH = st.apLineH * lines.length;
        if (lx + maxW <= ax.right + 6) {
          lines.forEach(function (rs, i) {
            layers[4].push(textSpec(tag + '.ap' + i, nm(i === 0 ? 'FDA approval label' : 'ACIP label'), lx, cy - totalH / 2 + i * st.apLineH, widths[i], st.apLineH, rs, 'l'));
          });
        } else {
          // No room before the "Why" column: stack the labels under the star instead.
          lines.forEach(function (rs, i) {
            below.push(textSpec(tag + '.ap' + i, nm(i === 0 ? 'FDA approval label' : 'ACIP label'), xf - widths[i] / 2, cy + 9 + i * st.apLineH, widths[i], st.apLineH, rs, 'c', { anchorX: xf, stackWith: tag + '.ap' }));
          });
        }
      }

      // ---- optional note above the row ----
      if (sc.note && chainStart !== null) {
        var nw = textWidth(sc.note, st.noteSize) + 2;
        var nx = Math.min(Math.max(chainStart, ax.left + 2), ax.right - nw);
        above.push(textSpec(tag + '.note', nm('Note'), nx, cy - 33.5, nw, st.labelH, [run(sc.note, st.noteSize, st.noteColor)], 'l', { prefLevel: 1 }));
      }

      resolveAbove(above, st.labelH);
      resolveBelow(below);
      above.concat(below).forEach(function (s) { delete s.anchorX; delete s.stackWith; delete s.prefLevel; layers[4].push(s); });
    });

    // ---- Today line and tag ----
    if (model.showToday !== false) {
      var tm = asOfMonth(opts.asOf || model.asOf || null);
      if (tm >= ax.startMonth && tm < ax.endMonth) {
        var tx = X(tm);
        var tagTop = cal.gridBottom + 14.2;
        var y0 = cal.gridTop + 7.6, y1 = tagTop + st.todayTagH / 2;
        // Leave a gap wherever a label sits on the line, so the dashes never run through text.
        var gaps = layers[4].filter(function (s) { return s.x - 1.5 <= tx && tx <= s.x + s.w + 1.5; })
          .map(function (s) { return [s.y - 1, s.y + s.h + 1]; })
          .sort(function (a, b) { return a[0] - b[0]; });
        var segs = [], cur = y0;
        gaps.forEach(function (g) { if (g[0] > cur) segs.push([cur, Math.min(g[0], y1)]); cur = Math.max(cur, g[1]); });
        if (cur < y1) segs.push([cur, y1]);
        segs.filter(function (q) { return q[1] - q[0] >= 2; }).forEach(function (q, i) {
          layers[0].push({ key: 'today.line' + (i ? '.' + i : ''), name: 'LTL|today|Today line', geom: 'line', x: tx, y: q[0], w: 0, h: q[1] - q[0], fill: null, line: { color: st.todayColor, w: st.todayLineW, dash: 'sysDash' }, text: null });
        });
        layers[5].push({ key: 'today.tag', name: 'LTL|today|Today tag', geom: 'roundRect', x: tx - st.todayTagW / 2, y: tagTop, w: st.todayTagW, h: st.todayTagH, fill: st.todayColor, line: null, text: { runs: [run('Today', st.todaySize, '#FFFFFF', true)], align: 'c' } });
      }
    }

    var specs = [].concat.apply([], layers);
    specs.forEach(function (s) { ['x', 'y', 'w', 'h'].forEach(function (k) { s[k] = Math.round(s[k] * 100) / 100; }); });
    return { specs: specs, warnings: warnings };
  }

  // Labels above the lane: keep each on its preferred level, lift one level when it would overlap another.
  function resolveAbove(list, lh) {
    list.sort(function (a, b) { return (a.prefLevel || 0) - (b.prefLevel || 0) || a.x - b.x; });
    var placed = [];
    list.forEach(function (s) {
      var guard = 0;
      while (placed.some(function (p) { return overlaps(p, s, 2); }) && guard++ < 4) s.y -= lh;
      placed.push(s);
    });
  }

  // Labels below the lane: merge identical neighbours (e.g. D and S in the same quarter), otherwise push right.
  function resolveBelow(list) {
    var stacks = list.filter(function (s) { return s.stackWith; });
    var singles = list.filter(function (s) { return !s.stackWith; }).sort(function (a, b) { return a.x - b.x; });
    for (var i = 1; i < singles.length; i++) {
      var a = singles[i - 1], b = singles[i];
      if (!overlaps(a, b, 2)) continue;
      if (a.text.runs[0].t === b.text.runs[0].t) {
        var mid = (a.anchorX + b.anchorX) / 2;
        a.x = mid - a.w / 2;
        a.anchorX = mid;
        a.name = a.name.replace(/\|[^|]*$/, '|Readout and submission label');
        singles.splice(i, 1); list.splice(list.indexOf(b), 1); i--;
      } else {
        var push = (a.x + a.w + 2) - b.x;
        a.x -= push / 2; b.x += push / 2;
      }
    }
    // A stacked approval label below the star moves down past any marker label it would hit.
    if (stacks.length) {
      var hit = stacks.some(function (s) { return singles.some(function (q) { return overlaps(q, s, 1); }); });
      if (hit) stacks.forEach(function (s) { s.y += 12; });
    }
  }

  var api = { build: build, textWidth: textWidth, DEFAULT_STYLE: DEFAULT_STYLE, styleOf: styleOf, asOfMonth: asOfMonth };
  root.LT = root.LT || {};
  root.LT.layout = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
