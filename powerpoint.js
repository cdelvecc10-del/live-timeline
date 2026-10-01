/* global PowerPoint, Office */
/*
 * Live Timeline — PowerPoint layer (Office.js).
 * Reads the timeline table, stores the model in slide tags, and draws the layout engine's shape specs as native shapes.
 * Every shape the add-in owns is named "LTL|..." so it can be found and redrawn; nothing else on the slide is touched.
 *
 * Requirement sets used: PowerPointApi 1.3 (tags), 1.4 (shapes, text), 1.5 (selection);
 * optional 1.8 (grouping, reading tables) and 1.9 (table column widths / row heights).
 */
(function (root) {
  'use strict';
  var LT = root.LT;
  var PREFIX = 'LTL|';
  var GROUP_NAME = 'LTL|Live timeline';
  var TAG_N = 'LTL_MODEL_N';
  var TAG_P = 'LTL_MODEL_';
  var CHUNK = 4000;

  /** @param {string} v */
  function supports(v) {
    try { return Office.context.requirements.isSetSupported('PowerPointApi', v); } catch (e) { return false; }
  }

  /** @param {string} g */
  function geomType(g) {
    switch (g) {
      case 'roundRect': return 'RoundRectangle';
      case 'ellipse': return 'Ellipse';
      case 'star5': return 'Star5';
      default: return 'Rectangle';
    }
  }
  /** @param {string} d */
  function dashType(d) {
    switch (d) {
      case 'dash': return 'Dash';
      case 'sysDash': return 'SystemDash';
      case 'sysDot': return 'SystemDot';
      case 'lgDash': return 'LongDash';
      default: return 'Solid';
    }
  }
  /** @param {string} a */
  function alignType(a) { return a === 'c' ? 'Center' : a === 'r' ? 'Right' : 'Left'; }

  /**
   * @param {PowerPoint.RequestContext} ctx
   * @param {string|null|undefined} slideId
   * @returns {Promise<PowerPoint.Slide>}
   */
  async function getSlide(ctx, slideId) {
    if (slideId) {
      var byId = ctx.presentation.slides.getItemOrNullObject(slideId);
      byId.load('id');
      await ctx.sync();
      if (!byId.isNullObject) return byId;
    }
    var sel = ctx.presentation.getSelectedSlides();
    sel.load('items/id');
    await ctx.sync();
    if (sel.items.length) return sel.items[0];
    var all = ctx.presentation.slides;
    all.load('items/id');
    await ctx.sync();
    if (!all.items.length) throw new Error('This presentation has no slides.');
    return all.items[0];
  }

  /** @param {PowerPoint.RequestContext} ctx @param {PowerPoint.Slide} slide */
  async function readModel(ctx, slide) {
    var tags = slide.tags;
    tags.load('items/key,items/value');
    await ctx.sync();
    /** @type {Object<string, string>} */
    var map = {};
    tags.items.forEach(function (t) { map[t.key.toUpperCase()] = t.value; });
    var n = parseInt(map[TAG_N] || '0', 10);
    if (!n) return null;
    var json = '';
    for (var i = 0; i < n; i++) json += map[TAG_P + i] || '';
    try { return JSON.parse(json); } catch (e) { return null; }
  }

  /** @param {PowerPoint.RequestContext} ctx @param {PowerPoint.Slide} slide @param {object} model */
  async function writeModel(ctx, slide, model) {
    var json = JSON.stringify(model);
    /** @type {string[]} */
    var parts = [];
    for (var i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
    var tags = slide.tags;
    tags.load('items/key');
    await ctx.sync();
    tags.items.forEach(function (t) {
      var k = t.key.toUpperCase();
      if (k.indexOf(TAG_P) === 0 && k !== TAG_N) {
        var idx = parseInt(k.slice(TAG_P.length), 10);
        if (!(idx < parts.length)) tags.delete(t.key);
      }
    });
    parts.forEach(function (p, j) { tags.add(TAG_P + j, p); });
    tags.add(TAG_N, String(parts.length));
    await ctx.sync();
  }

  /** @param {PowerPoint.RequestContext} ctx @param {PowerPoint.Slide} slide */
  async function shapeList(ctx, slide) {
    var shapes = slide.shapes;
    shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
    await ctx.sync();
    return shapes;
  }

  /**
   * Finds the timeline table and measures its period columns and scenario rows.
   * @param {PowerPoint.RequestContext} ctx @param {PowerPoint.Slide} slide
   */
  async function detect(ctx, slide) {
    if (!supports('1.8')) return { cal: null, reason: 'Reading the table needs PowerPoint for Microsoft 365, Version 2504 or later.' };
    var shapes = await shapeList(ctx, slide);
    var tables = shapes.items.filter(function (s) { return s.type === 'Table'; });
    if (!tables.length) return { cal: null, reason: 'There is no table on this slide.' };
    var has19 = supports('1.9');
    var jobs = tables.map(function (s) {
      var t = s.getTable();
      t.load('rowCount,columnCount,values');
      if (has19) { t.columns.load('items/width'); t.rows.load('items/currentHeight'); }
      return { s: s, t: t };
    });
    await ctx.sync();
    if (!has19) return { cal: null, reason: 'Measuring the table needs PowerPoint for Microsoft 365, Version 2508 or later.' };
    for (var i = 0; i < jobs.length; i++) {
      var j = jobs[i];
      var cal = LT.calib.fromTable({ left: j.s.left, top: j.s.top }, j.t.values,
        j.t.columns.items.map(function (c) { return c.width; }),
        j.t.rows.items.map(function (r) { return r.currentHeight; }));
      if (cal) { cal.tableId = j.s.id; return { cal: cal, reason: '' }; }
    }
    return { cal: null, reason: 'No table with year and quarter headers was found on this slide.' };
  }

  /** @param {PowerPoint.RequestContext} ctx @param {PowerPoint.Slide} slide */
  async function clearOwned(ctx, slide) {
    var shapes = await shapeList(ctx, slide);
    var n = 0;
    shapes.items.forEach(function (s) { if ((s.name || '').indexOf(PREFIX) === 0) { s.delete(); n++; } });
    await ctx.sync();
    return n;
  }

  /**
   * @param {PowerPoint.Shape} sh
   * @param {{geom:string, text:{runs:Array<{t:string,size:number,color:string,bold:boolean}>, align:string}}} s
   */
  function applyText(sh, s) {
    var tf = sh.textFrame;
    // Theme defaults for new shapes can add autofit, padding, italics or a larger font; reset all of them.
    tf.autoSizeSetting = 'AutoSizeNone';
    tf.wordWrap = false;
    tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
    tf.verticalAlignment = 'Middle';
    var runs = s.text.runs;
    if (s.geom !== 'text') tf.textRange.text = runs.map(function (r) { return r.t; }).join('');
    var tr = tf.textRange;
    tr.paragraphFormat.horizontalAlignment = alignType(s.text.align);
    tr.paragraphFormat.bulletFormat.visible = false;
    var f = tr.font;
    f.size = runs[0].size; f.bold = !!runs[0].bold; f.italic = false; f.color = runs[0].color;
    var pos = runs[0].t.length;
    for (var i = 1; i < runs.length; i++) {
      var sub = tr.getSubstring(pos, runs[i].t.length).font;
      sub.size = runs[i].size; sub.bold = !!runs[i].bold; sub.italic = false; sub.color = runs[i].color;
      pos += runs[i].t.length;
    }
  }

  /** @param {PowerPoint.ShapeCollection} shapes @param {any} s */
  function addSpec(shapes, s) {
    var opt = { left: s.x, top: s.y, width: s.w, height: s.h };
    /** @type {PowerPoint.Shape} */
    var sh;
    if (s.geom === 'text') sh = shapes.addTextBox(s.text ? s.text.runs.map(function (/** @type {{t:string}} */ r) { return r.t; }).join('') : '', opt);
    else if (s.geom === 'line') sh = shapes.addLine('Straight', opt);
    else sh = shapes.addGeometricShape(geomType(s.geom), opt);
    sh.name = s.name;
    if (s.geom !== 'line') { if (s.fill) sh.fill.setSolidColor(s.fill); else sh.fill.clear(); }
    if (s.line) {
      sh.lineFormat.visible = true;
      sh.lineFormat.color = s.line.color;
      sh.lineFormat.weight = s.line.w;
      sh.lineFormat.dashStyle = dashType(s.line.dash);
    } else sh.lineFormat.visible = false;
    if (s.text) applyText(sh, s);
    // Re-assert the geometry last in case a theme's autofit resized the shape when text was added.
    sh.left = s.x; sh.top = s.y; sh.width = s.w; sh.height = s.h;
    return sh;
  }

  /**
   * Redraws the live timeline on the slide and saves the model into the slide.
   * @param {string|null} slideId @param {any} model @param {any} cal @param {any} [opts]
   */
  function redraw(slideId, model, cal, opts) {
    return PowerPoint.run(async function (ctx) {
      var slide = await getSlide(ctx, slideId);
      var removed = await clearOwned(ctx, slide);
      var built = LT.layout.build(model, cal, opts || {});
      var shapes = slide.shapes;
      var made = built.specs.map(function (/** @type {any} */ s) { return addSpec(shapes, s); });
      await ctx.sync();
      var grouped = false;
      if (made.length > 1 && supports('1.8')) {
        try {
          var g = shapes.addGroup(made);
          g.name = GROUP_NAME;
          await ctx.sync();
          grouped = true;
        } catch (e) { /* grouping is cosmetic; shapes stay individually named */ }
      }
      model.calibration = cal;
      await writeModel(ctx, slide, model);
      return { count: made.length, removed: removed, grouped: grouped, warnings: built.warnings, slideId: slide.id };
    });
  }

  /** Saves the model without redrawing (e.g. when live updates are off). @param {string|null} slideId @param {any} model */
  function save(slideId, model) {
    return PowerPoint.run(async function (ctx) {
      var slide = await getSlide(ctx, slideId);
      await writeModel(ctx, slide, model);
    });
  }

  /** Reads the slide: saved model (if any), timeline table geometry, slide id and position. @param {string|null} [slideId] */
  function load(slideId) {
    return PowerPoint.run(async function (ctx) {
      var slide = await getSlide(ctx, slideId);
      var index = null;
      if (supports('1.8')) { slide.load('index'); await ctx.sync(); index = slide.index; }
      var model = await readModel(ctx, slide);
      var det = await detect(ctx, slide);
      return { slideId: slide.id, index: index, model: model, cal: det.cal, reason: det.reason };
    });
  }

  function selectedSlideId() {
    return PowerPoint.run(async function (ctx) {
      var sel = ctx.presentation.getSelectedSlides();
      sel.load('items/id');
      await ctx.sync();
      return sel.items.length ? sel.items[0].id : null;
    });
  }

  /** Finds shapes that look like a hand-drawn timeline inside the table area, and selects them so they can be checked. @param {string|null} slideId @param {any} cal */
  function findLegacy(slideId, cal) {
    return PowerPoint.run(async function (ctx) {
      var slide = await getSlide(ctx, slideId);
      var shapes = await shapeList(ctx, slide);
      var list = shapes.items.map(function (s) { return { id: s.id, name: s.name, type: String(s.type), left: s.left, top: s.top, width: s.width, height: s.height }; });
      var ids = LT.calib.legacyCandidates(list, cal);
      if (ids.length && supports('1.5')) { slide.setSelectedShapes(ids); await ctx.sync(); }
      return list.filter(function (s) { return ids.indexOf(s.id) >= 0; });
    });
  }

  /** @param {string|null} slideId @param {string[]} ids */
  function deleteShapes(slideId, ids) {
    return PowerPoint.run(async function (ctx) {
      var slide = await getSlide(ctx, slideId);
      ids.forEach(function (id) { slide.shapes.getItem(id).delete(); });
      await ctx.sync();
      return ids.length;
    });
  }

  /** Removes the add-in's shapes and saved data from the slide. @param {string|null} slideId */
  function removeAll(slideId) {
    return PowerPoint.run(async function (ctx) {
      var slide = await getSlide(ctx, slideId);
      var n = await clearOwned(ctx, slide);
      var tags = slide.tags;
      tags.load('items/key');
      await ctx.sync();
      tags.items.forEach(function (t) { if (t.key.toUpperCase().indexOf('LTL_') === 0) tags.delete(t.key); });
      await ctx.sync();
      return n;
    });
  }

  var api = {
    supports: supports, load: load, redraw: redraw, save: save, selectedSlideId: selectedSlideId,
    findLegacy: findLegacy, deleteShapes: deleteShapes, removeAll: removeAll
  };
  root.LT = root.LT || {};
  root.LT.ppt = api;
})(typeof window !== 'undefined' ? window : globalThis);
