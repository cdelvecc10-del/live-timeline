/*
 * Live Timeline — reading the time axis and scenario rows from a timeline table.
 * Pure functions: the PowerPoint layer passes in the table's text, column widths and row heights (points).
 */
(function (root) {
  'use strict';
  var E = (root.LT && root.LT.engine) || (typeof require !== 'undefined' ? require('./engine.js') : null);

  var RE_Q = /^\s*(?:([1-4])\s*Q|Q\s*([1-4]))\s*$/i;
  var RE_H = /^\s*(?:([12])\s*H|H\s*([12]))\s*$/i;
  var RE_Y = /^\s*(?:FY\s*)?'?(\d{4}|\d{2})\s*$/i;

  function clean(v) { return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim(); }
  function sum(a, from, to) { var t = 0; for (var i = from; i < to; i++) t += a[i]; return t; }

  /** Finds the period columns of a header row. Returns { unit, cols:[{c, n}] } or null. */
  function periodColumns(row) {
    var qs = [], hs = [], ms = [];
    row.forEach(function (v, c) {
      var s = clean(v), m;
      if ((m = s.match(RE_Q))) qs.push({ c: c, n: +(m[1] || m[2]) });
      else if ((m = s.match(RE_H))) hs.push({ c: c, n: +(m[1] || m[2]) });
      else if (s.length >= 3 && /^[A-Za-z]{3,9}$/.test(s) && E.parseMonth(s + ' 2000') !== null) ms.push({ c: c, n: E.monthOf(E.parseMonth(s + ' 2000')) });
    });
    if (qs.length >= 2) return { unit: 3, cols: qs };
    if (ms.length >= 2) return { unit: 1, cols: ms };
    if (hs.length >= 2) return { unit: 6, cols: hs };
    return null;
  }

  /**
   * @param {{left:number, top:number}} geom  table position on the slide (points)
   * @param {string[][]} values               cell text (merged cells may be blank)
   * @param {number[]} colW                   column widths (points)
   * @param {number[]} rowH                   rendered row heights (points)
   * @returns calibration or null when the table does not look like a timeline
   */
  function fromTable(geom, values, colW, rowH) {
    if (!values || !values.length || !colW || !rowH) return null;
    var nR = values.length;
    var yearRow = -1, unitRow = -1, per = null;
    for (var r = 0; r < Math.min(nR, 4); r++) {
      var p = periodColumns(values[r]);
      if (p) { unitRow = r; per = p; break; }
    }
    // Year header: the row (at or above the period row) with at least two year-like cells.
    for (var ry = 0; ry <= Math.max(unitRow, Math.min(nR - 1, 2)); ry++) {
      var years = values[ry].filter(function (v) { var s = clean(v); return RE_Y.test(s) && (s.length >= 4 || /'/.test(s) || /FY/i.test(s)); });
      if (years.length >= 2 || (per && years.length >= 1 && ry < unitRow)) { yearRow = ry; break; }
    }
    if (yearRow < 0) return null;

    // Year for each column: carry the last year seen from the left (merged year cells span their quarters).
    var yearAt = [], last = null;
    values[yearRow].forEach(function (v, c) {
      var m = clean(v).match(RE_Y);
      if (m && (clean(v).length >= 4 || /'/.test(v) || /FY/i.test(v))) last = Number(m[1]) < 100 ? 2000 + Number(m[1]) : Number(m[1]);
      yearAt[c] = last;
    });

    var cols, unit;
    if (per) { cols = per.cols; unit = per.unit; }
    else {
      unit = 12;
      cols = [];
      values[yearRow].forEach(function (v, c) { if (RE_Y.test(clean(v)) && clean(v).length >= 2) cols.push({ c: c, n: 1 }); });
    }
    if (cols.length < 2) return null;
    // Columns must be contiguous.
    for (var i = 1; i < cols.length; i++) if (cols[i].c !== cols[i - 1].c + 1) return null;
    var c0 = cols[0].c, cN = cols[cols.length - 1].c;
    var y0 = yearAt[c0];
    if (!y0) return null;
    var startMonth = E.mi(y0, (cols[0].n - 1) * unit + 1);
    var endMonth = startMonth + unit * cols.length;

    var left = geom.left + sum(colW, 0, c0);
    var right = left + sum(colW, c0, cN + 1);
    var headerRows = Math.max(yearRow, unitRow) + 1;
    var lanes = [], names = [];
    for (var rr = headerRows; rr < nR; rr++) {
      var top = geom.top + sum(rowH, 0, rr);
      lanes.push({ top: round(top), bottom: round(top + rowH[rr]) });
      var name = '';
      for (var cc = 0; cc < c0 && !name; cc++) name = clean(values[rr][cc]);
      names.push(name);
    }
    if (!lanes.length) return null;
    return {
      axis: { startMonth: startMonth, endMonth: endMonth, left: round(left), right: round(right) },
      lanes: lanes, names: names,
      gridTop: lanes[0].top, gridBottom: lanes[lanes.length - 1].bottom,
      source: 'table'
    };
  }

  function round(v) { return Math.round(v * 100) / 100; }

  /**
   * Picks hand-drawn timeline shapes (from before the add-in) so they can be removed once.
   * shapes: [{id, name, type, left, top, width, height}] in points.
   */
  function legacyCandidates(shapes, cal) {
    var ax = cal.axis;
    return shapes.filter(function (s) {
      if (!s || (s.name || '').indexOf('LTL|') === 0) return false;
      if (/^today/i.test(s.name || '')) return true;
      if (['GeometricShape', 'TextBox', 'Line', 'Group', 'Freeform', 'Callout'].indexOf(s.type) < 0) return false;
      var cx = s.left + s.width / 2, cy = s.top + s.height / 2;
      return cx >= ax.left && cx <= ax.right && cy >= cal.gridTop && cy <= cal.gridBottom;
    }).map(function (s) { return s.id; });
  }

  var api = { fromTable: fromTable, legacyCandidates: legacyCandidates, periodColumns: periodColumns };
  root.LT = root.LT || {};
  root.LT.calib = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
