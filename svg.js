/*
 * Live Timeline — SVG renderer for the browser preview (draws the same shape specs PowerPoint gets).
 */
(function (root) {
  'use strict';

  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }
  var DASH = { dash: [4, 3], sysDash: [3, 1], sysDot: [1, 1], lgDash: [8, 3] };
  function strokeAttrs(line) {
    if (!line) return ' stroke="none"';
    var a = ' stroke="' + line.color + '" stroke-width="' + line.w + '"';
    var d = DASH[line.dash];
    if (d) a += ' stroke-dasharray="' + d.map(function (v) { return +(v * line.w).toFixed(2); }).join(' ') + '"';
    return a;
  }
  function f2(v) { return +(+v).toFixed(2); }

  // PowerPoint's star5 preset: outer points touch the box; the centre sits at 55.3% of the height.
  function star5(x, y, w, h) {
    var cx = x + w / 2, cy = y + h * 0.55279, rx = w * 0.52573, ry = h * 0.55279, k = 0.38196, pts = [];
    for (var i = 0; i < 10; i++) {
      var a = -Math.PI / 2 + i * Math.PI / 5, r = (i % 2) ? k : 1;
      pts.push(f2(cx + rx * r * Math.cos(a)) + ',' + f2(cy + ry * r * Math.sin(a)));
    }
    return pts.join(' ');
  }

  function shape(s) {
    var fill = s.fill ? ' fill="' + s.fill + '"' : ' fill="none"';
    var out = '';
    switch (s.geom) {
      case 'rect': out = '<rect x="' + f2(s.x) + '" y="' + f2(s.y) + '" width="' + f2(s.w) + '" height="' + f2(s.h) + '"' + fill + strokeAttrs(s.line) + '/>'; break;
      case 'roundRect': var r = f2(Math.min(s.w, s.h) * 0.16667);
        out = '<rect x="' + f2(s.x) + '" y="' + f2(s.y) + '" width="' + f2(s.w) + '" height="' + f2(s.h) + '" rx="' + r + '" ry="' + r + '"' + fill + strokeAttrs(s.line) + '/>'; break;
      case 'ellipse': out = '<ellipse cx="' + f2(s.x + s.w / 2) + '" cy="' + f2(s.y + s.h / 2) + '" rx="' + f2(s.w / 2) + '" ry="' + f2(s.h / 2) + '"' + fill + strokeAttrs(s.line) + '/>'; break;
      case 'star5': out = '<polygon points="' + star5(s.x, s.y, s.w, s.h) + '"' + fill + strokeAttrs(s.line) + ' stroke-linejoin="miter"/>'; break;
      case 'line': out = '<line x1="' + f2(s.x) + '" y1="' + f2(s.y) + '" x2="' + f2(s.x + s.w) + '" y2="' + f2(s.y + s.h) + '"' + strokeAttrs(s.line) + '/>'; break;
      default: if (s.fill) out = '<rect x="' + f2(s.x) + '" y="' + f2(s.y) + '" width="' + f2(s.w) + '" height="' + f2(s.h) + '"' + fill + '/>';
    }
    if (s.text && s.text.runs.length) {
      var al = s.text.align, anchor = al === 'c' ? 'middle' : al === 'r' ? 'end' : 'start';
      var tx = al === 'c' ? s.x + s.w / 2 : al === 'r' ? s.x + s.w : s.x;
      var size = s.text.runs[0].size;
      var ty = s.y + s.h / 2 + size * 0.355;
      out += '<text x="' + f2(tx) + '" y="' + f2(ty) + '" text-anchor="' + anchor + '" font-family="Arial, Helvetica, \'Liberation Sans\', sans-serif" style="white-space:pre">' +
        s.text.runs.map(function (r) {
          return '<tspan font-size="' + r.size + '" fill="' + r.color + '"' + (r.bold ? ' font-weight="700"' : '') + '>' + esc(r.t) + '</tspan>';
        }).join('') + '</text>';
    }
    return '<g data-key="' + esc(s.key) + '"><title>' + esc(s.name.replace(/^LTL\|/, '').replace(/\|/g, ' · ')) + '</title>' + out + '</g>';
  }

  /**
   * @param {Array} specs
   * @param {{width?:number, height?:number, backdrop?:string, backdropSvg?:string}} [opts]
   */
  function render(specs, opts) {
    opts = opts || {};
    var W = opts.width || 960, H = opts.height || 540;
    var bg = opts.backdrop ? '<image href="' + opts.backdrop + '" x="0" y="0" width="' + W + '" height="' + H + '" preserveAspectRatio="none"/>' : (opts.backdropSvg || '<rect width="' + W + '" height="' + H + '" fill="#fff"/>');
    return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Timeline slide preview">' + bg + '<g class="ltl">' + specs.map(shape).join('') + '</g></svg>';
  }

  /** A plain timeline table (years, quarters, scenario rows) for previews that have no slide picture. */
  function tableBackdrop(cal, names, title) {
    var E = root.LT.engine, ax = cal.axis, out = [];
    var W = 960, mw = (ax.right - ax.left) / (ax.endMonth - ax.startMonth);
    var top = cal.gridTop - 64, yr = top + 28;
    out.push('<rect width="960" height="540" fill="#fff"/>');
    if (title) out.push('<text x="29" y="44" font-family="Arial" font-size="22" font-weight="700" fill="#3C3C3C">' + esc(title) + '</text>');
    out.push('<rect x="29" y="' + top + '" width="' + (ax.left - 29) + '" height="64" fill="#3C3C3C"/>');
    out.push('<text x="40" y="' + (top + 37) + '" font-family="Arial" font-size="12" font-weight="700" fill="#fff">Scenario</text>');
    for (var m = ax.startMonth; m < ax.endMonth; m += 3) {
      var x = ax.left + (m - ax.startMonth) * mw, w = 3 * mw;
      if (E.monthOf(m) <= 3) {
        out.push('<rect x="' + f2(x) + '" y="' + top + '" width="' + f2(Math.min(12 * mw, ax.right - x)) + '" height="28" fill="#3C3C3C" stroke="#fff" stroke-width="1"/>');
        out.push('<text x="' + f2(x + Math.min(12 * mw, ax.right - x) / 2) + '" y="' + (top + 19) + '" text-anchor="middle" font-family="Arial" font-size="13" font-weight="700" fill="#fff">' + E.yearOf(m) + '</text>');
      }
      out.push('<rect x="' + f2(x) + '" y="' + yr + '" width="' + f2(w) + '" height="36" fill="#F2F2F2" stroke="#fff" stroke-width="1"/>');
      out.push('<text x="' + f2(x + w / 2) + '" y="' + (yr + 22) + '" text-anchor="middle" font-family="Arial" font-size="10" font-weight="700" fill="#3C3C3C">' + E.quarterOf(m) + 'Q</text>');
      if (E.yearOf(m) % 2 === 1) out.push('<rect x="' + f2(x) + '" y="' + cal.gridTop + '" width="' + f2(w) + '" height="' + (cal.gridBottom - cal.gridTop) + '" fill="#F5F5F5"/>');
      out.push('<line x1="' + f2(x) + '" y1="' + cal.gridTop + '" x2="' + f2(x) + '" y2="' + cal.gridBottom + '" stroke="#D9D9D9" stroke-width="0.5" stroke-dasharray="2 2"/>');
    }
    cal.lanes.forEach(function (ln, i) {
      out.push('<rect x="29" y="' + ln.top + '" width="' + (ax.left - 29) + '" height="' + (ln.bottom - ln.top) + '" fill="#F4F4F4"/>');
      out.push('<line x1="29" y1="' + ln.bottom + '" x2="' + ax.right + '" y2="' + ln.bottom + '" stroke="#BFBFBF" stroke-width="0.75"/>');
      out.push('<text x="40" y="' + ((ln.top + ln.bottom) / 2 + 4) + '" font-family="Arial" font-size="11" font-weight="700" fill="#3C3C3C">' + esc((names && names[i]) || ('Scenario ' + (i + 1))) + '</text>');
    });
    return out.join('');
  }

  var api = { render: render, shape: shape, tableBackdrop: tableBackdrop, star5: star5 };
  root.LT = root.LT || {};
  root.LT.svg = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
