// make_manga_text_v2.jsx — Accurate center placement for manga balloons
// - Uses image_size scaling
// - Prefers tight text bbox center; falls back to polygon centroid if provided
// - Rotation-aware (rotation_deg)
// - Snaps layer to visual center after rendering
// - Designed for Persian translations (centered paragraph text)

#target photoshop
app.displayDialogs = DialogModes.NO;
var __origRulerUnits = app.preferences.rulerUnits;
var __origTypeUnits = app.preferences.typeUnits;
app.preferences.rulerUnits = Units.PIXELS;
try { app.preferences.typeUnits = TypeUnits.PIXELS; } catch (e) {}

function __restorePrefs() {
  try { app.preferences.typeUnits = __origTypeUnits; } catch (e) {}
  try { app.preferences.rulerUnits = __origRulerUnits; } catch (e) {}
}

// ===== USER CONFIG =====
var scriptFolder = Folder("C:/Users/abbas/Desktop/psd maker new");  // change if needed
var imageFile   = File(scriptFolder + "/wild/raw 48/0048-011.png");
var jsonFile    = File(scriptFolder + "/wild/json final/0048-011.json");
var outputPSD   = File(scriptFolder + "/manga_output.psd");
var outputJPG   = File(scriptFolder + "/manga_output.jpg");
var EXPORT_PSD_TO_JPG = true; // set to false to skip exporting a JPG copy of the PSD

// Optional batch mode: process an entire chapter
// Put all page images inside `chapterImagesFolder` and a matching JSON per image
// (same base name, .json extension) inside `chapterJsonFolder`.
// Outputs will be written to `chapterOutputFolder` using `<basename>_output.psd`.
var PROCESS_WHOLE_CHAPTER = false;
var chapterImagesFolder  = Folder(scriptFolder + "/wild/raw 48");
var chapterJsonFolder    = Folder(scriptFolder + "/wild/json final");
var chapterOutputFolder  = Folder(scriptFolder + "/chapter_output");

// ===== DEBUG + FILE LOGGING =====
var DEBUG = true;
var LOG_TO_FILE = true;
var DEBUG_CONTENT_AWARE = false; // highlight removal selections for debugging
var CONTENT_AWARE_DEBUG_COLOR = { red: 255, green: 0, blue: 0, opacity: 35 };
var logFile = File(scriptFolder + "/manga_log.txt");

function initLog() {
  if (!LOG_TO_FILE) return;
  try {
    logFile.encoding = "UTF8";
    logFile.open("a");
    var now = new Date();
    logFile.writeln("\n==============================");
    logFile.writeln("Run at: " + now.toString());
    logFile.writeln("==============================");
    logFile.close();
  } catch (e) {}
}

function log(msg){
  if (DEBUG) {
    try { $.writeln(msg); } catch(e) {}
  }
  if (LOG_TO_FILE) {
    try {
      logFile.encoding = "UTF8";
      logFile.open("a");
      logFile.writeln(msg);
      logFile.close();
    } catch(e) {}
  }
}
initLog();

// ===== JSON (legacy-safe) =====
if (typeof JSON === 'undefined') { JSON = {}; JSON.parse = function (s) { return eval('(' + s + ')'); }; }

// ===== String helpers =====
function toStr(v){
  if (v === undefined || v === null) return "";
  try{
    if (typeof v === "object") {
      if (v.valueOf && typeof v.valueOf() !== "object") return String(v.valueOf());
      try { return JSON.stringify(v); } catch(e2){ return String(v); }
    }
    return String(v);
  } catch(e){ return ""; }
}
function safeTrim(s){
  return keepZWNJ(s, function (txt) {
    return toStr(txt)
      .replace(/^[ \t\u00A0\r\n]+/, "")
      .replace(/[ \t\u00A0\r\n]+$/, "");
  });
}
// Do not treat Zero Width Non-Joiner (\u200C) as whitespace so Persian نیم‌فاصله stays intact
// Protect Persian zero-width non-joiner (half-space) from being eaten by whitespace cleanup
var ZWNJ = "\u200C";
// Normalize ZWNJ handling for ExtendScript JSON quirks
function normalizeZWNJInJsonText(rawJson) {
  if (!rawJson) return "";
  var txt = String(rawJson);

  // Normalize any HTML-style ZWNJ encodings to actual U+200C
  txt = txt.replace(/&zwnj;/gi, ZWNJ);
  txt = txt.replace(/&#8204;/gi, ZWNJ);

  // Convert every actual U+200C in the JSON text into a literal "\u200c"
  txt = txt.replace(/\u200C/g, "\\u200c");

  return txt;
}

function normalizeZWNJFromText(val) {
  var str = (val === undefined || val === null) ? "" : String(val);
  if (!str) return "";

  var normalized = str;
  normalized = normalized.replace(/\u200c/gi, ZWNJ);
  normalized = normalized.replace(/\\u200c/gi, ZWNJ);
  normalized = normalized.replace(/&zwnj;/gi, ZWNJ);
  normalized = normalized.replace(/&#8204;/gi, ZWNJ);
  return normalized;
}

function restoreZWNJ(val, placeholder) {
  if (val instanceof Array) {
    var restored = [];
    for (var i = 0; i < val.length; i++) {
      restored.push(restoreZWNJ(val[i], placeholder));
    }
    return restored;
  }
  return toStr(val).split(placeholder).join(ZWNJ);
}

function keepZWNJ(str, transformFn) {
  if (!transformFn) return str;
  var placeholder = "__ZWNJ__SAFE__";
  var guarded = toStr(str).split(ZWNJ).join(placeholder);
  var result = transformFn(guarded);
  return restoreZWNJ(result, placeholder);
}

function collapseWhitespace(s){
  return keepZWNJ(s, function (txt) {
    return txt.replace(/[ \t\u00A0\r\n]+/g, " ");
  });
}

// Keep line breaks, normalize spaces, and convert literal "\n" to real newlines
function normalizeWSKeepBreaks(s){
  var str = toStr(s);
  if (!str) return "";
  return keepZWNJ(str, function (raw) {
    var norm = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    norm = norm.replace(/\\n/g, "\n");
    var parts = norm.split("\n");
    for (var i=0; i<parts.length; i++){
      parts[i] = parts[i].replace(/[ \t\u00A0]+/g, " ");
    }
    return parts.join("\n");
  });
}

function normalizeItemZWNJ(item) {
  if (!item) return item;
  if (item.text !== undefined) item.text = normalizeZWNJFromText(item.text);
  return item;
}

function normalizeItemsZWNJ(items) {
  if (!items || !items.length) return items;
  for (var i = 0; i < items.length; i++) {
    normalizeItemZWNJ(items[i]);
  }
  return items;
}

// ===== Geometry helpers =====
function solidBlack(){ var c=new SolidColor(); c.rgb.red=0;c.rgb.green=0;c.rgb.blue=0; return c; }
function solidWhite(){ var c=new SolidColor(); c.rgb.red=255;c.rgb.green=255;c.rgb.blue=255; return c; }
function buildItemPalette(item) {
  var reversed = !!(item && item.reverse_color === true);
  return {
    textColor: reversed ? solidWhite() : solidBlack(),
    strokeColor: reversed ? solidBlack() : solidWhite(),
    simpleBgColor: reversed ? solidBlack() : solidWhite(),
    reversed: reversed
  };
}
function clampInt(v){ return Math.round(v); }
function layerBoundsPx(lyr){
  var b = lyr.bounds;
  return {
    left:   b[0].as('px'),
    top:    b[1].as('px'),
    right:  b[2].as('px'),
    bottom: b[3].as('px'),
    width:  b[2].as('px') - b[0].as('px'),
    height: b[3].as('px') - b[1].as('px')
  };
}
function layerCenterPx(lyr){
  var bb = layerBoundsPx(lyr);
  return { x:(bb.left+bb.right)/2, y:(bb.top+bb.bottom)/2 };
}
function translateToCenter(lyr, cx, cy){
  var c = layerCenterPx(lyr);
  lyr.translate(cx - c.x, cy - c.y);
}

function scaleLayerToBubbleIfSmallText(lyr, ti, item, scaleX, scaleY, fillRatio, minPointSize, cx, cy) {
  if (!lyr || !ti || !item) return;
  if (!item.bbox_bubble) return;

  var threshold = (typeof minPointSize === "number") ? minPointSize : 18;
  if (ti.size >= threshold) return;

  var maxPointSize = 22;
  if (typeof maxPointSize === "number" && maxPointSize > ti.size) {
    var scaleUp = maxPointSize / ti.size;
    if (scaleUp > 1) {
      lyr.resize(scaleUp * 100, scaleUp * 100, AnchorPosition.MIDDLECENTER);
      translateToCenter(lyr, cx, cy);
    }
  }

  var bubbleBox = normalizeBox(item.bbox_bubble);
  if (!bubbleBox) return;

  var ratio = (typeof fillRatio === "number") ? fillRatio : 0.9;
  var bubbleW = Math.max(1, (bubbleBox.right - bubbleBox.left) * scaleX);
  var bubbleH = Math.max(1, (bubbleBox.bottom - bubbleBox.top) * scaleY);
  var targetW = bubbleW * ratio;
  var targetH = bubbleH * ratio;

  var bounds = layerBoundsPx(lyr);
  if (!bounds || bounds.width <= 0 || bounds.height <= 0) return;

  var scaleW = targetW / bounds.width;
  var scaleH = targetH / bounds.height;
  var scale = Math.min(scaleW, scaleH);
  if (scale >= 1) return;

  lyr.resize(scale * 100, scale * 100, AnchorPosition.MIDDLECENTER);
  translateToCenter(lyr, cx, cy);
}

function applyStrokeColor(layer, sizePx, color) {
  if (!layer) return;
  var strokeSize = (typeof sizePx === 'number' && sizePx > 0) ? sizePx : 2;
  var col = (color && color.rgb) ? color.rgb : solidWhite().rgb;
  try {
    var s2t = stringIDToTypeID, c2t = charIDToTypeID;
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(c2t('Prpr'), s2t('layerEffects'));
    ref.putEnumerated(c2t('Lyr '), c2t('Ordn'), c2t('Trgt'));
    desc.putReference(c2t('null'), ref);

    var effects = new ActionDescriptor();
    var stroke = new ActionDescriptor();

    stroke.putBoolean(s2t('enabled'), true);
    stroke.putBoolean(s2t('present'), true);
    stroke.putBoolean(s2t('showInDialog'), true);
    stroke.putEnumerated(s2t('style'), s2t('frameStyle'), s2t('outsetFrame'));
    stroke.putEnumerated(s2t('paintType'), s2t('paintType'), s2t('solidColor'));
    stroke.putEnumerated(s2t('mode'), s2t('blendMode'), s2t('normal'));
    stroke.putUnitDouble(s2t('opacity'), c2t('#Prc'), 100.0);
    stroke.putUnitDouble(s2t('size'), c2t('#Pxl'), strokeSize);

    var colorDesc = new ActionDescriptor();
    colorDesc.putDouble(c2t('Rd  '), col.red);
    colorDesc.putDouble(c2t('Grn '), col.green);
    colorDesc.putDouble(c2t('Bl  '), col.blue);
    stroke.putObject(c2t('Clr '), c2t('RGBC'), colorDesc);

    effects.putObject(s2t('frameFX'), s2t('frameFX'), stroke);
    desc.putObject(c2t('T   '), s2t('layerEffects'), effects);
    executeAction(c2t('setd'), desc, DialogModes.NO);
  } catch (e) {
    log('  ⚠️ unable to apply stroke: ' + e);
  }
}

function getBasePixelLayer(doc){
  if (!doc) return null;
  try {
    if (doc.backgroundLayer) return doc.backgroundLayer;
  } catch (e) {}
  try {
    if (doc.layers && doc.layers.length) return doc.layers[doc.layers.length - 1];
  } catch (e2) {}
  return null;
}

function ensureCleanupLayer(doc, baseLayer) {
  if (!doc || !baseLayer) return null;
  var CLEANUP_NAME = "ContentAware Cleanup";
  try {
    for (var i = 0; i < doc.layers.length; i++) {
      var layer = doc.layers[i];
      if (layer && layer.name === CLEANUP_NAME) {
        return layer;
      }
    }
  } catch (scanErr) {}

  try {
    var duplicate = baseLayer.duplicate();
    duplicate.name = CLEANUP_NAME;
    log('  created cleanup layer "' + CLEANUP_NAME + '" above background');
    return duplicate;
  } catch (dupErr) {
    log('  ⚠️ unable to duplicate base layer for cleanup: ' + dupErr);
  }
  return null;
}

function buildContentAwareDebugColor() {
  var c = new SolidColor();
  var cfg = CONTENT_AWARE_DEBUG_COLOR || {};
  c.rgb.red   = (typeof cfg.red === 'number')   ? cfg.red   : 255;
  c.rgb.green = (typeof cfg.green === 'number') ? cfg.green : 0;
  c.rgb.blue  = (typeof cfg.blue === 'number')  ? cfg.blue  : 0;
  return c;
}

function createSelectionDebugLayer(doc) {
  if (!doc) return null;
  try {
    var layer = doc.artLayers.add();
    layer.name = "ContentAware Selection Debug";
    layer.blendMode = BlendMode.NORMAL;
    var cfgOpacity = CONTENT_AWARE_DEBUG_COLOR && CONTENT_AWARE_DEBUG_COLOR.opacity;
    layer.opacity = (typeof cfgOpacity === 'number') ? cfgOpacity : 35;
    try {
      if (doc.layers && doc.layers.length) {
        layer.move(doc.layers[0], ElementPlacement.PLACEBEFORE);
      }
    } catch (moveErr) {}
    return layer;
  } catch (e) {
    log('  ⚠️ unable to create selection debug layer: ' + e);
  }
  return null;
}

function highlightSelectionForDebug(doc, cleanupLayer) {
  if (!DEBUG_CONTENT_AWARE) return null;
  if (!doc) return null;
  var debugLayer = createSelectionDebugLayer(doc);
  if (!debugLayer) return null;
  try {
    doc.activeLayer = debugLayer;
    var color = buildContentAwareDebugColor();
    try {
      doc.selection.fill(color);
    } catch (fillErr) {
      log('  ⚠️ unable to fill debug selection: ' + fillErr);
    }
  } catch (layerErr) {
    log('  ⚠️ could not activate debug layer: ' + layerErr);
  }
  try {
    if (cleanupLayer) doc.activeLayer = cleanupLayer;
  } catch (restoreErr) {}
  log('  [debug] highlighted removal selection on "' + debugLayer.name + '"');
  return debugLayer;
}

function contentAwareFillSelection(){
  try {
    var s2t = stringIDToTypeID;
    var c2t = charIDToTypeID;
    var desc = new ActionDescriptor();
    desc.putEnumerated(c2t('Usng'), c2t('FlCn'), s2t('contentAware'));
    desc.putUnitDouble(c2t('Opct'), c2t('#Prc'), 100.0);
    desc.putEnumerated(c2t('Md  '), c2t('BlnM'), c2t('Nrml'));
    desc.putBoolean(c2t('PrsT'), false);
    executeAction(c2t('Fl  '), desc, DialogModes.NO);
    return true;
  } catch (e) {
    log('  ⚠️ content aware fill failed: ' + e);
    return false;
  }
}

function segmentPoints(seg) {
  if (!seg) return null;
  if (seg.points && seg.points.length) return seg.points;
  if (seg.polygon && seg.polygon.length) return seg.polygon;
  if (seg.vertices && seg.vertices.length) return seg.vertices;
  if (seg.coords && seg.coords.length) return seg.coords;
  if (seg.path && seg.path.length) return seg.path;
  if (seg.length && seg[0] && (seg[0].length >= 2 || (typeof seg[0].x === 'number' && typeof seg[0].y === 'number')))
    return seg;
  return null;
}

function buildRegionFromSegment(seg, scaleX, scaleY) {
  if (!seg) return null;
  var pts = segmentPoints(seg);
  if (pts && pts.length >= 3) {
    var regionPts = [];
    for (var i = 0; i < pts.length; i++) {
      var p = pts[i];
      if (!p) continue;
      var px, py;
      if (p.length && p.length >= 2) {
        px = Number(p[0]);
        py = Number(p[1]);
      } else {
        px = Number(p.x);
        py = Number(p.y);
      }
      if (isNaN(px) || isNaN(py)) continue;
      regionPts.push([px * scaleX, py * scaleY]);
    }
    if (regionPts.length >= 3) return regionPts;
  }

  var x1 = Number(seg.x_min);
  var x2 = Number(seg.x_max);
  var y1 = Number(seg.y_min);
  var y2 = Number(seg.y_max);
  if (isNaN(x1) || isNaN(x2) || isNaN(y1) || isNaN(y2)) return null;
  var left = Math.min(x1, x2) * scaleX;
  var right = Math.max(x1, x2) * scaleX;
  var top = Math.min(y1, y2) * scaleY;
  var bottom = Math.max(y1, y2) * scaleY;
  if ((right - left) < 1 || (bottom - top) < 1) return null;
  return [
    [left, top],
    [right, top],
    [right, bottom],
    [left, bottom]
  ];
}

function selectSegments(doc, segments, scaleX, scaleY){
  if (!doc || !segments || !segments.length) return false;
  var selectionMade = false;
  for (var i = 0; i < segments.length; i++) {
    var seg = segments[i];
    var region = buildRegionFromSegment(seg, scaleX, scaleY);
    if (!region) continue;
    try {
      doc.selection.select(region, selectionMade ? SelectionType.EXTEND : SelectionType.REPLACE, 0, false);
      selectionMade = true;
    } catch (e) {}
  }
  return selectionMade;
}

function boxToSegment(box, pad){
  var norm = normalizeBox(box);
  if (!norm) return null;
  var p = (typeof pad === "number") ? pad : 0;
  return {
    x_min: norm.left   - p,
    x_max: norm.right  + p,
    y_min: norm.top    - p,
    y_max: norm.bottom + p
  };
}

function polygonToBox(points){
  if (!points || points.length < 1) return null;
  var minX = points[0][0], maxX = points[0][0];
  var minY = points[0][1], maxY = points[0][1];
  for (var i = 1; i < points.length; i++) {
    var pt = points[i];
    if (!pt || pt.length < 2) continue;
    if (pt[0] < minX) minX = pt[0];
    if (pt[0] > maxX) maxX = pt[0];
    if (pt[1] < minY) minY = pt[1];
    if (pt[1] > maxY) maxY = pt[1];
  }
  return { left:minX, top:minY, right:maxX, bottom:maxY };
}

function normalizeBox(box) {
  if (!box) return null;

  function isNum(v) { return typeof v === "number" && !isNaN(v); }

  var hasLTRB = isNum(box.left) && isNum(box.top) && isNum(box.right) && isNum(box.bottom);
  if (hasLTRB) {
    return {
      left: Number(box.left),
      top: Number(box.top),
      right: Number(box.right),
      bottom: Number(box.bottom)
    };
  }

  var hasMinMax = isNum(box.x_min) && isNum(box.x_max) && isNum(box.y_min) && isNum(box.y_max);
  if (hasMinMax) {
    var x1 = Number(box.x_min), x2 = Number(box.x_max);
    var y1 = Number(box.y_min), y2 = Number(box.y_max);
    return {
      left: Math.min(x1, x2),
      right: Math.max(x1, x2),
      top: Math.min(y1, y2),
      bottom: Math.max(y1, y2)
    };
  }

  return null;
}

function collectRemovalSegments(item){
  var result = { segments: [], source: "none" };
  if (!item) return result;
  if (item.segments && item.segments.length) {
    result.segments = item.segments;
    result.source = "json";
    return result;
  }

  var derived = [];
  var pad = 2;
  var box = null;

  if (item.polygon_text && item.polygon_text.length >= 3) {
    box = polygonToBox(item.polygon_text);
    result.source = "polygon_text";
  } else if (item.bbox_text) {
    box = item.bbox_text;
    result.source = "bbox_text";
  } else if (item.bbox_bubble) {
    box = item.bbox_bubble;
    result.source = "bbox_bubble";
  }

  if (box) {
    var seg = boxToSegment(box, pad);
    if (seg) derived.push(seg);
  }

  result.segments = derived;
  if (!derived.length) result.source = "none";
  return result;
}

function removeOldTextSegments(doc, item, scaleX, scaleY, palette){
  var segmentsInfo = collectRemovalSegments(item);
  if (!segmentsInfo.segments.length) {
    log('  ⚠️ no segments available for content-aware fill');
    return;
  }
  var baseLayer = getBasePixelLayer(doc);
  if (!baseLayer) {
    log('  ⚠️ no base layer found for content-aware fill');
    return;
  }
  var cleanupLayer = ensureCleanupLayer(doc, baseLayer);
  if (!cleanupLayer) {
    log('  ⚠️ could not prepare cleanup layer; skipping content-aware fill');
    return;
  }
  doc.activeLayer = cleanupLayer;
  try { doc.selection.deselect(); } catch (e) {}
  var useContentAware = !(item && item.complex_background === false);
  if (segmentsInfo.source !== 'json') {
    log('  deriving removal segments from ' + segmentsInfo.source);
  }
  var selectionMade = selectSegments(doc, segmentsInfo.segments, scaleX, scaleY);
  if (!selectionMade) {
    log('  ⚠️ segments provided but no valid selection could be made');
    return;
  }
  highlightSelectionForDebug(doc, cleanupLayer);
  if (useContentAware) {
    log('  removing original text via content-aware fill over ' + segmentsInfo.segments.length + ' segments');
    var filled = contentAwareFillSelection();
    if (!filled) {
      log('  ⚠️ content-aware fill skipped due to error');
    }
  } else {
    var fillColor = (palette && palette.simpleBgColor) ? palette.simpleBgColor : solidBlack();
    log('  simple background -> filling selection with color over ' + segmentsInfo.segments.length + ' segments');
    try {
      doc.selection.fill(fillColor);
    } catch (fillErr) {
      log('  ⚠️ unable to fill selection: ' + fillErr);
    }
  }
  try { doc.selection.deselect(); } catch (e2) {}
}

// ===== Centers/boxes =====اه
function polygonCentroid(points){
  var n = points.length;
  if (n < 3) return null;
  var A=0, Cx=0, Cy=0, i, x1,y1,x2,y2,c;
  for (i=0; i<n; i++){
    x1 = points[i][0]; y1 = points[i][1];
    x2 = points[(i+1)%n][0]; y2 = points[(i+1)%n][1];
    c  = (x1*y2 - x2*y1);
    A  += c;
    Cx += (x1 + x2) * c;
    Cy += (y1 + y2) * c;
  }
  A /= 2;
  if (Math.abs(A) < 1e-6){
    var sx=0, sy=0;
    for (i=0;i<n;i++){ sx+=points[i][0]; sy+=points[i][1]; }
    return { x:sx/n, y:sy/n };
  }
  return { x: Cx/(6*A), y: Cy/(6*A) };
}
function deriveCenter(item){
  if (item && item.polygon_text && item.polygon_text.length>=3){
    var c = polygonCentroid(item.polygon_text); if (c) return c;
  }
  if (item && item.center && typeof item.center.x==="number" && typeof item.center.y==="number")
    return { x:item.center.x, y:item.center.y };

  if (item && item.bbox_text){
    var bt=normalizeBox(item.bbox_text);
    if (bt) return { x:(bt.left+bt.right)/2, y:(bt.top+bt.bottom)/2 };
  }
  if (item && item.bbox_bubble){
    var bb=normalizeBox(item.bbox_bubble);
    if (bb) return { x:(bb.left+bb.right)/2, y:(bb.top+bb.bottom)/2 };
  }
  return null;
}
function deriveBox(item){
  if (item && item.bbox_text){
    var bt = normalizeBox(item.bbox_text);
    if (bt) return bt;
  }
  if (item && item.bbox_bubble){
    var bb = normalizeBox(item.bbox_bubble);
    if (bb) return bb;
  }
  return null;
}

// ===== Fonts & RTL helpers =====
var FONT_ALIASES = {
  "IRANSans Black": [
    "IRANSans-Black",
    "IRANSansBlack",
    "IRANSans",
    "IRANSansWeb-Black",
    "IRANSansWeb(FaNum)",
    "IRANSans(FaNum)",
    "IRANSans(Farsi)",
    "IranSans-Black",
    "IranSans Black"
  ],
  "IRKoodak": [
    "IRKoodak-Regular",
    "IR Koodak",
    "IR Koodak Bold",
    "IRKoodak Bold",
    "IRKoodakBold",
    "IRKoodak(Bold)",
    "B Koodak",
    "BKoodak",
    "B Koodak Bold",
    "B Koodak(Farsi)"
  ],
  "B Morvarid": ["B Morvarid-Regular", "B Morvarid Regular"],
  "AFSANEH": ["AFSANEH-Regular", "AFSANEH Regular"],
  "Potk": ["Potk-Black"],
  "Kalameh": ["Kalameh-Regular"],
  "A nic Regular": ["A nic", "A-nic", "A-nic-Regular"],
  "IRFarnaz": [
    "IRFarnaz-Regular",
    "IRFarnaz Regular",
    "IR Farnaz",
    "IRFarnaz(Farsi)",
    "IRFarnaz Farsi",
    "B Farnaz",
    "BFarnaz",
    "Farnaz",
    "Farnaz(Farsi)"
  ],
  "Shabnam-BoldItalic": ["Shabnam Bold Italic", "Shabnam BoldItalic", "ShabnamBI"]
};

function normalizeFontId(name) {
  return collapseWhitespace(toStr(name)).toLowerCase().replace(/[\s_-]+/g, "");
}

function collectFontCandidates(fontName) {
  var seen = {};
  var candidates = [];
  function add(name) {
    if (!name) return;
    var id = normalizeFontId(name);
    if (!id || seen[id]) return;
    seen[id] = true;
    candidates.push(name);
  }

  add(fontName);
  if (FONT_ALIASES[fontName]) {
    for (var i = 0; i < FONT_ALIASES[fontName].length; i++) {
      add(FONT_ALIASES[fontName][i]);
    }
  }

  try {
    var fonts = app.fonts;
    if (fonts && fonts.length) {
      for (var c = 0; c < candidates.length; c++) {
        var target = normalizeFontId(candidates[c]);
        for (var j = 0; j < fonts.length; j++) {
          var f = fonts[j];
          if (!f) continue;
          var psId = normalizeFontId(f.postScriptName);
          var nameId = normalizeFontId(f.name);
          if (psId === target || nameId === target) {
            add(f.postScriptName);
            add(f.name);
          }
        }
      }
    }
  } catch (e) {}

  add("ArialMT");
  add("Arial");
  add("Arial-BoldMT");
  return candidates;
}

function findInstalledFont(preferredNames) {
  try {
    var fonts = app.fonts;
    if (!fonts || !fonts.length) return null;
    for (var i = 0; i < preferredNames.length; i++) {
      var target = normalizeFontId(preferredNames[i]);
      for (var j = 0; j < fonts.length; j++) {
        var f = fonts[j];
        if (!f) continue;
        if (normalizeFontId(f.postScriptName) === target || normalizeFontId(f.name) === target) {
          return f;
        }
      }
    }
  } catch (e) {}
  return null;
}

function resolveFontOrFallback(fontName) {
  var candidates = collectFontCandidates(fontName);

  var found = findInstalledFont(candidates);
  if (found) {
    if (normalizeFontId(found.postScriptName) !== normalizeFontId(fontName)) {
      log('  ⚠️ requested font "' + fontName + '" resolved to "' + found.postScriptName + '"');
    }
    return found.postScriptName;
  }

  log('  ⚠️ requested font "' + fontName + '" not found; falling back to Photoshop default');
  try {
    return app.fonts[0].postScriptName;
  } catch (e) {
    return fontName || "ArialMT";
  }
}

function applyFontToTextItem(textItem, requestedFont) {
  if (!textItem) return requestedFont;

  var candidates = collectFontCandidates(requestedFont);
  var fallback = requestedFont;

  // Prefer assigning via a real font object (textFont) so Photoshop doesn't silently
  // swap the font when the string id is ambiguous.
  var matched = findInstalledFont(candidates);
  if (matched) {
    try {
      textItem.textFont = matched;
      textItem.font = matched.postScriptName;
      if (requestedFont && normalizeFontId(matched.postScriptName) !== normalizeFontId(requestedFont)) {
        log('  ⚠️ requested font "' + requestedFont + '" resolved to "' + matched.postScriptName + '"');
      }
      return matched.postScriptName;
    } catch (assignErr) {
      log('  ⚠️ failed to apply font object "' + matched.postScriptName + '": ' + assignErr);
    }
  }

  for (var i = 0; i < candidates.length; i++) {
    var candidate = candidates[i];
    try {
      textItem.font = candidate;
      var applied = textItem.font;
      if (normalizeFontId(applied) !== normalizeFontId(candidate)) {
        log('  ⚠️ Photoshop applied font "' + applied + '" instead of requested "' + candidate + '"');
        continue;
      }
      if (requestedFont && normalizeFontId(applied) !== normalizeFontId(requestedFont)) {
        log('  ⚠️ requested font "' + requestedFont + '" resolved to "' + applied + '"');
      }
      return applied;
    } catch (e) {
      log('  ⚠️ failed to apply font "' + candidate + '": ' + e);
    }

    if (!fallback) fallback = candidate;
  }

  try {
    var defaultFont = app.fonts && app.fonts.length ? app.fonts[0].postScriptName : (fallback || "ArialMT");
    textItem.font = defaultFont;
    log('  ⚠️ falling back to default font "' + defaultFont + '"');
    return defaultFont;
  } catch (fallbackErr) {
    log('  ⚠️ could not apply fallback font: ' + fallbackErr);
  }

  return fallback || requestedFont;
}

function getFontForType(type){
  var requested;
  switch(type){
    case "Standard": requested = "IRKoodak"; break;
    case "Thought":         requested = "B Morvarid"; break;
    case "Shouting/Emotion":requested = "AFSANEH"; break;
    case "Whisper/Soft":    requested = "A nic Regular"; break;
    case "Electronic":      requested = "Consolas"; break;
    case "Narration":   requested = "IRFarnaz"; break;
    case "Distorted/Custom":requested = "Shabnam-BoldItalic"; break;
    default:                requested = "ArialMT"; break;
  }
  return resolveFontOrFallback(requested);
}

function forceRTL(s){
  var RLE="\u202B", PDF="\u202C", RLM="\u200F";
  var str = toStr(s);
  if (!str) return RLM;
  return keepZWNJ(str, function(raw){
    var norm = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    var parts = norm.split("\n");
    for (var i=0; i<parts.length; i++){
      parts[i] = RLM + RLE + parts[i] + PDF;
    }
    return parts.join("\r");
  });
}

function applyParagraphDirectionRTL(textForDirection) {
  try {
    if (app.activeDocument.activeLayer.kind !== LayerKind.TEXT) return;
    var s2t = stringIDToTypeID, c2t = charIDToTypeID;
    var ti  = app.activeDocument.activeLayer.textItem;
    var txt = forceRTL(textForDirection !== undefined ? textForDirection : ti.contents);
    var len = txt.length;

    var desc = new ActionDescriptor();
    var ref  = new ActionReference();
    ref.putEnumerated(s2t('textLayer'), c2t('Ordn'), c2t('Trgt'));
    desc.putReference(c2t('null'), ref);

    var textDesc = new ActionDescriptor();

    var psList  = new ActionList();
    var psRange = new ActionDescriptor();
    psRange.putInteger(s2t('from'), 0);
    psRange.putInteger(s2t('to'),   len);

    var pStyle = new ActionDescriptor();
    pStyle.putEnumerated(
      s2t('paragraphDirection'),
      s2t('paragraphDirection'),
      s2t('RightToLeftParagraph')
    );
    pStyle.putEnumerated(
      s2t('justification'),
      s2t('justification'),
      s2t('center')
    );

    psRange.putObject(s2t('paragraphStyle'), s2t('paragraphStyle'), pStyle);
    psList.putObject(s2t('paragraphStyleRange'), psRange);
    textDesc.putList(s2t('paragraphStyleRange'), psList);

    desc.putObject(c2t('T   '), s2t('textLayer'), textDesc);
    executeAction(c2t('setd'), desc, DialogModes.NO);

    ti.justification = Justification.CENTER;
  } catch(e) {}
}

function trySetMEEveryLineComposer() {
  try {
    var s2t = stringIDToTypeID, c2t = charIDToTypeID;
    var d = new ActionDescriptor(), r = new ActionReference();
    r.putEnumerated(s2t('textLayer'), c2t('Ordn'), c2t('Trgt'));
    d.putReference(c2t('null'), r);
    var t = new ActionDescriptor();
    t.putEnumerated(
      s2t('textComposer'),
      s2t('textComposer'),
      s2t('adbeMEEveryLineComposer')
    );
    d.putObject(c2t('T   '), s2t('textLayer'), t);
    executeAction(c2t('setd'), d, DialogModes.NO);
  } catch (e) {}
}

function deriveRotationDeg(item){
  if (typeof item.rotation_deg === "number") return item.rotation_deg;
  if (item.writing_mode === "vertical-rl" || item.writing_mode === "vertical-lr") return 90;
  return 0;
}

// ===== Precise wrapping inspired by the Python helper =====
function TextMeasureContext(doc, fontName, sizePx) {
  this.doc = doc;
  this.layer = doc.artLayers.add();
  this.layer.name = "__measure__";
  this.layer.kind = LayerKind.TEXT;
  this.layer.opacity = 0;
  var ti = this.layer.textItem;
  ti.kind = TextType.POINTTEXT;
  ti.contents = ".";
  applyFontToTextItem(ti, fontName);
  ti.size = sizePx;
  ti.justification = Justification.CENTER;
  ti.position = [0, 0];
  this.textItem = ti;
  this.size = sizePx;
}

TextMeasureContext.prototype.dispose = function(){
  try { this.layer.remove(); } catch (e) {}
};

TextMeasureContext.prototype.measureWidth = function(text){
  if (!text) return 0;
  try { app.activeDocument.activeLayer = this.layer; } catch (e) {}
  this.layer.visible = true;
  setTextContentsRTL(this.textItem, text);
  var bounds = layerBoundsPx(this.layer);
  return bounds.width;
};

TextMeasureContext.prototype.lineHeight = function(){
  return Math.max(1, Math.round(this.size * 1.18) + 4);
};

function splitWordsForLayout(text) {
  var s = safeTrim(toStr(text));
  if (!s) return [];
  return keepZWNJ(s, function (guarded) {
    var raw = guarded.split(/[ \t\u00A0\r\n]+/);
    var words = [];
    for (var i = 0; i < raw.length; i++) {
      if (raw[i]) words.push(raw[i]);
    }
    return words;
  });
}

function greedyWrapWordsPixels(words, measureCtx, maxWidth) {
  var lines = [];
  var current = [];
  if (!words || !words.length) return lines;
  for (var i = 0; i < words.length; i++) {
    var word = words[i];
    var candidate = current.concat([word]).join(" ");
    if (!current.length) {
      current = [word];
      continue;
    }
    if (measureCtx.measureWidth(candidate) <= maxWidth) {
      current.push(word);
    } else {
      lines.push(current);
      current = [word];
    }
  }
  if (current.length) lines.push(current);
  return lines;
}

function rebalanceFirstLineStrict(linesWords, measureCtx, maxWidth, maxMoves) {
  if (!linesWords || linesWords.length < 2) return linesWords;
  var moves = 0;
  var limit = maxMoves || 5;
  while (moves < limit && linesWords[0].length > 1) {
    var line1 = linesWords[0].join(" ");
    var line2 = linesWords[1].join(" ");
    var w1 = measureCtx.measureWidth(line1);
    var w2 = measureCtx.measureWidth(line2);
    if (w1 < w2) break;

    var lastWord = linesWords[0][linesWords[0].length - 1];
    var newLine1 = linesWords[0].slice(0, -1);
    var newLine2 = [lastWord].concat(linesWords[1]);
    var newLine1Text = newLine1.length ? newLine1.join(" ") : lastWord;
    var newLine2Text = newLine2.join(" ");
    var newW1 = measureCtx.measureWidth(newLine1Text);
    var newW2 = measureCtx.measureWidth(newLine2Text);
    if (newW2 > maxWidth) break;
    linesWords[0] = newLine1.length ? newLine1 : [lastWord];
    linesWords[1] = newLine2;
    moves++;
  }
  return linesWords;
}

function rebalanceMiddleLines(wordsLines, measureCtx, maxWidth, passes) {
  if (!wordsLines || wordsLines.length < 3) return wordsLines;
  var maxPasses = passes || 4;
  for (var pass = 0; pass < maxPasses; pass++) {
    var changed = false;
    var widths = [];
    for (var i = 0; i < wordsLines.length; i++) {
      widths[i] = measureCtx.measureWidth(wordsLines[i].join(" "));
    }
    for (var j = 1; j < wordsLines.length - 1; j++) {
      if (!wordsLines[j + 1] || !wordsLines[j + 1].length) continue;
      var wCurrent = widths[j];
      var wNext = widths[j + 1];
      if (wCurrent >= wNext) continue;
      var nextWord = wordsLines[j + 1][0];
      var candidate = wordsLines[j].concat([nextWord]).join(" ");
      var candidateWidth = measureCtx.measureWidth(candidate);
      if (candidateWidth <= maxWidth) {
        wordsLines[j].push(wordsLines[j + 1].shift());
        widths[j] = candidateWidth;
        widths[j + 1] = measureCtx.measureWidth(wordsLines[j + 1].join(" "));
        changed = true;
      }
    }
    if (!changed) break;
  }
  return wordsLines;
}

function layoutBubbleLines(text, doc, fontName, sizePx, maxWidth, maxHeight, sharedCtx) {
  var widthLimit = Math.max(1, maxWidth || 1);
  var ownsContext = !sharedCtx;
  var ctx = sharedCtx || new TextMeasureContext(doc, fontName, sizePx);
  try {
    var words = splitWordsForLayout(text);
    if (!words.length) return [];
    var linesWords = greedyWrapWordsPixels(words, ctx, widthLimit);
    if (maxHeight) {
      var lh = ctx.lineHeight();
      var maxLines = Math.max(1, Math.floor(maxHeight / lh));
      if (linesWords.length > maxLines) {
        // if overflow, caller can shrink font later
      }
    }
    linesWords = rebalanceFirstLineStrict(linesWords, ctx, widthLimit, 5);
    linesWords = rebalanceMiddleLines(linesWords, ctx, widthLimit, 4);
    return linesWords;
  } finally {
    if (ownsContext) ctx.dispose();
  }
}

function layoutBubble(text, doc, fontName, sizePx, maxWidth, maxHeight, sharedCtx) {
  var linesWords = layoutBubbleLines(text, doc, fontName, sizePx, maxWidth, maxHeight, sharedCtx);
  if (!linesWords.length) return "";
  var parts = [];
  for (var i = 0; i < linesWords.length; i++) {
    parts.push(linesWords[i].join(" "));
  }
  return parts.join("\r");
}

// === Estimate SAFE starting font size (can grow up to 2x) ===
function estimateSafeFontSizeForWrapped(wrapped, innerW, innerH, baseSize){
  var txt = toStr(wrapped);
  if (!txt) return baseSize;

  var parts = txt.split(/\r|\n/);
  var lineCount = parts.length;
  if (lineCount < 1) return baseSize;

  var maxChars = 0;
  for (var i=0; i<parts.length; i++){
    var line = safeTrim(parts[i]);
    if (!line) continue;
    if (line.length > maxChars) maxChars = line.length;
  }
  if (maxChars <= 0) return baseSize;

  var charWidthFactor = 0.6;
  var leadingFactor   = 1.18;

  var maxSizeW = Math.floor(innerW / (maxChars * charWidthFactor));
  var maxSizeH = Math.floor(innerH / (lineCount * leadingFactor * 1.1));

  var upperBound = Math.max(6, Math.min(maxSizeW, maxSizeH));
  var safeSize;

  if (upperBound <= baseSize) {
    safeSize = upperBound;           // must shrink
  } else {
    var growLimit = Math.floor(baseSize * 2.0); // allow up to 2x
    safeSize = Math.min(upperBound, growLimit);
    if (safeSize < baseSize) safeSize = baseSize;
  }

  if (safeSize < 6) safeSize = 6;
  if (safeSize > 60) safeSize = 60;

  log("  [preSize] base=" + baseSize +
      " maxW=" + maxSizeW + " maxH=" + maxSizeH +
      " upper=" + upperBound + " -> safe=" + safeSize);

  return safeSize;
}

// Create paragraph exactly the inner bubble size
function setTextContentsRTL(ti, text) {
  if (!ti) return;
  ti.contents = forceRTL(text);
}

// ===== Fast point-text fitting (single-pass scale) =====
// Creates a POINTTEXT layer, fits it to the target bubble once, and centers it.
// Requirements:
// - Uses point text (no paragraph box)
// - No textItem.width/height
// - Explicitly sets textItem.kind = TextType.POINTTEXT
// - Uses manual line breaks (\r) from precomputed lines
// - Single measurement to compute scale, one font-size set, optional re-measure, single translate
function createPointTextFastFit(doc, linesArray, fontName, bubbleRect, alignment) {
  if (!doc) return null;

  var lines = linesArray && linesArray.length ? linesArray : [""];
  var text = lines.join("\r");

  var lyr = doc.artLayers.add();
  lyr.kind = LayerKind.TEXT;
  var ti = lyr.textItem;
  ti.kind = TextType.POINTTEXT;
  setTextContentsRTL(ti, text);
  applyFontToTextItem(ti, fontName);

  var initialSize = 48;
  ti.size = initialSize;
  try { ti.leading = Math.max(1, Math.floor(initialSize * 1.18)); } catch (e) {}

  var justification = Justification.CENTER;
  if (alignment === Justification.LEFT || alignment === Justification.RIGHT || alignment === Justification.CENTER) {
    justification = alignment;
  } else if (typeof alignment === "string") {
    var alignLower = alignment.toLowerCase();
    if (alignLower === "left") justification = Justification.LEFT;
    if (alignLower === "right") justification = Justification.RIGHT;
    if (alignLower === "center") justification = Justification.CENTER;
  }
  ti.justification = justification;

  // Use a stable start position; we will center via translate after sizing.
  ti.position = [bubbleRect.x, bubbleRect.y];

  // Measure once at the initial size.
  var initialBounds = layerBoundsPx(lyr);
  var measuredW = Math.max(1, initialBounds.width);
  var measuredH = Math.max(1, initialBounds.height);

  var scaleW = bubbleRect.width / measuredW;
  var scaleH = bubbleRect.height / measuredH;
  var scale = Math.min(scaleW, scaleH) * 0.98;
  var finalSize = Math.max(1, initialSize * scale);

  // Apply the scaled font size once.
  ti.size = finalSize;
  try { ti.leading = Math.max(1, Math.floor(finalSize * 1.18)); } catch (e2) {}

  // Optional re-measure after resizing (no loops).
  var finalBounds = layerBoundsPx(lyr);

  var targetCx = bubbleRect.x + bubbleRect.width / 2;
  var targetCy = bubbleRect.y + bubbleRect.height / 2;
  var currentCx = (finalBounds.left + finalBounds.right) / 2;
  var currentCy = (finalBounds.top + finalBounds.bottom) / 2;

  // Center the text inside the bubble with a single translate call.
  lyr.translate(targetCx - currentCx, targetCy - currentCy);

  return lyr;
}

function createParagraphFullBox(doc, text, fontName, sizePx, cx, cy, innerW, innerH, textColor){
  var left = cx - innerW/2, top = cy - innerH/2;
  var lyr = doc.artLayers.add();
  lyr.kind = LayerKind.TEXT;
  var ti = lyr.textItem;
  ti.kind = TextType.PARAGRAPHTEXT;
  setTextContentsRTL(ti, text);
  applyFontToTextItem(ti, fontName);
  ti.size = sizePx;
  try { ti.leading = Math.max(1, Math.floor(sizePx * 1.18)); } catch(e){}
  ti.justification = Justification.CENTER;
  ti.position = [left, top];
  ti.width  = innerW;
  ti.height = innerH;
  try { ti.color = textColor ? textColor : solidBlack(); } catch(e){ ti.color = solidBlack(); }
  return lyr;
}

// ===== Auto fit (shrink or grow using binary search) =====
function normalizeLinesForAutoFit(text) {
  var str = toStr(text);
  if (!str) return [];
  var norm = str.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  var parts = norm.split("\n");
  var cleaned = [];
  for (var i = 0; i < parts.length; i++) {
    var part = collapseWhitespace(parts[i]);
    if (part.length) cleaned.push(part);
  }
  return cleaned;
}

function buildMeasureLayerFromLine(lyr, lineText) {
  if (!lyr || !lineText) return null;
  try {
    var measureLayer = lyr.duplicate();
    measureLayer.visible = false;
    var measureTi = measureLayer.textItem;
    measureTi.kind = TextType.POINTTEXT;
    measureTi.contents = forceRTL(lineText);
    measureTi.justification = Justification.CENTER;
    measureTi.position = [0, 0];
    return measureLayer;
  } catch (e) {
    return null;
  }
}

function buildMeasureLayerFromText(lyr, text) {
  if (!lyr || !text) return null;
  try {
    var measureLayer = lyr.duplicate();
    measureLayer.visible = false;
    var measureTi = measureLayer.textItem;
    measureTi.kind = TextType.POINTTEXT;
    measureTi.contents = forceRTL(toStr(text).replace(/\r/g, "\n"));
    measureTi.justification = Justification.CENTER;
    measureTi.position = [0, 0];
    return measureLayer;
  } catch (e) {
    return null;
  }
}

function measureLineWidthFromLayer(measureLayer, sizePx) {
  if (!measureLayer) return 0;
  try {
    var ti = measureLayer.textItem;
    ti.size = sizePx;
    var bounds = layerBoundsPx(measureLayer);
    return bounds.width;
  } catch (e) {
    return 0;
  }
}

function measureTextSizeFromLayer(measureLayer, sizePx) {
  if (!measureLayer) return null;
  try {
    var ti = measureLayer.textItem;
    ti.size = sizePx;
    var bounds = layerBoundsPx(measureLayer);
    return { width: bounds.width, height: bounds.height };
  } catch (e) {
    return null;
  }
}

function cleanupMeasureLayer(layer) {
  if (!layer) return;
  try { layer.remove(); } catch (e) {}
}

function overflowSlackPx(innerW, innerH) {
  var slackW = Math.max(2, Math.min(6, Math.round(innerW * 0.03)));
  var slackH = Math.max(2, Math.min(6, Math.round(innerH * 0.03)));
  return { w: slackW, h: slackH };
}

// Final safety pass: if the rendered text still overflows the intended box, shrink slightly
// until it fits. This guards against occasional measurement inaccuracies (e.g., fonts with
// large descenders or extra line spacing) that can leave part of the text clipped.
function clampRenderedTextToBox(lyr, ti, cx, cy, maxWidth, maxHeight, minSize) {
  if (!lyr || !ti) return;

  if (minSize === undefined) minSize = 10;

  var slack = overflowSlackPx(maxWidth, maxHeight);
  var MAX_ADJUSTMENTS = 8;
  translateToCenter(lyr, cx, cy);
  for (var i = 0; i < MAX_ADJUSTMENTS; i++) {
    var bounds = layerBoundsPx(lyr);
    var overflowW = bounds.width - maxWidth;
    var overflowH = bounds.height - maxHeight;

    if (overflowW <= slack.w && overflowH <= slack.h) break;

    // Nudge the font size down based on the worst overflow dimension, but keep a floor.
    var worstOverflow = Math.max(overflowW - slack.w, overflowH - slack.h);
    var shrink = Math.max(1, Math.ceil(worstOverflow / 8));
    var newSize = Math.max(minSize, ti.size - shrink);
    if (newSize === ti.size) newSize = Math.max(minSize, ti.size - 1);
    if (newSize === ti.size) break;

    ti.size = newSize;
    try { ti.leading = Math.max(1, Math.floor(newSize * 1.12)); } catch (e) {}
  }

  translateToCenter(lyr, cx, cy);
}

// After fitting/shrinking, make sure the paragraph box itself fully contains the rendered
// glyph bounds so nothing gets clipped by a too-small text box.
function expandParagraphBoxToContent(lyr, ti, cx, cy, paddingPx) {
  if (!lyr || !ti) return;

  var pad = (paddingPx === undefined) ? 6 : Math.max(0, paddingPx);

  translateToCenter(lyr, cx, cy);

  var bounds = layerBoundsPx(lyr);
  if (!bounds) return;

  var desiredW = Math.ceil(bounds.width + pad * 2);
  var desiredH = Math.ceil(bounds.height + pad * 2);

  var changed = false;

  try {
    if (desiredW > ti.width + 1) { ti.width = desiredW; changed = true; }
    if (desiredH > ti.height + 1) { ti.height = desiredH; changed = true; }
  } catch (e) {}

  if (changed) translateToCenter(lyr, cx, cy);
}

function autoFitTextLayer(lyr, ti, cx, cy, innerW, innerH, minSize, maxSize, rawTextForLines) {
  if (minSize === undefined) minSize = 10;
  if (maxSize === undefined) maxSize = 52;

  var lines = normalizeLinesForAutoFit(rawTextForLines);
  var longestLine = lines.length ? lines[0] : "";
  for (var li = 1; li < lines.length; li++) {
    if (lines[li].length > longestLine.length) longestLine = lines[li];
  }
  var lineCount = Math.max(1, lines.length || 1);
  var measureLayer = longestLine ? buildMeasureLayerFromLine(lyr, longestLine) : null;
  var measureAllTextLayer = rawTextForLines ? buildMeasureLayerFromText(lyr, rawTextForLines) : null;

  var slack = overflowSlackPx(innerW, innerH);

  log("  [autoFit] innerW=" + innerW + " innerH=" + innerH +
      (longestLine ? " longestLine=\"" + longestLine + "\"" : ""));

  function sizeFits(sizePx) {
    ti.size = sizePx;
    try { ti.leading = Math.max(1, Math.floor(sizePx * 1.12)); } catch (e) {}

    var b = layerBoundsPx(lyr);
    var w = b.width;
    var h = b.height;

    var measuredLineWidth = measureLayer ? measureLineWidthFromLayer(measureLayer, sizePx) : w;
    var measuredAllSize = measureTextSizeFromLayer(measureAllTextLayer, sizePx);
    var measuredAllWidth = measuredAllSize ? measuredAllSize.width : w;
    var measuredAllHeight = measuredAllSize ? measuredAllSize.height : h;

    var widthOverflow  = (w > innerW + slack.w) || (measuredLineWidth > innerW + slack.w) || (measuredAllWidth > innerW + slack.w);
    var heightOverflow = (h > innerH + slack.h) || (measuredAllHeight > innerH + slack.h);
    var overflow = widthOverflow || heightOverflow;

    log("    test size=" + sizePx +
        " bounds=(" + w.toFixed(1) + "x" + h.toFixed(1) + ")" +
        (measureLayer ? " longestLineWidth=" + measuredLineWidth.toFixed(1) : "") +
        (measuredAllSize ? " measureAll=(" + measuredAllWidth.toFixed(1) + "x" + measuredAllHeight.toFixed(1) + ")" : "") +
        " overflow=" + overflow);

    return !overflow;
  }

  var lo = minSize;
  var hi = maxSize;
  var best = minSize;

  if (ti.size >= minSize && ti.size <= maxSize) {
    if (sizeFits(ti.size)) best = ti.size;
  }

  var maxSteps = 10;
  var steps = 0;
  while (lo <= hi && steps < maxSteps) {
    var mid = Math.floor((lo + hi) / 2);
    if (sizeFits(mid)) {
      best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
    steps++;
  }

  var finalBounds = layerBoundsPx(lyr);
  var finalHeightEstimate = finalBounds ? finalBounds.height : innerH;
  log("  [autoFit] finalSize=" + best + " estHeight=" + finalHeightEstimate);

  sizeFits(best);

  cleanupMeasureLayer(measureLayer);
  cleanupMeasureLayer(measureAllTextLayer);

  var safeHeight = Math.max(innerH, finalHeightEstimate + 20);
  if (ti.height < safeHeight) {
    ti.height = safeHeight;
    log("  [autoFit] expand text box height to " + safeHeight);
  }

  translateToCenter(lyr, cx, cy);

  return {
    size: best,
    height: ti.height
  };
}

function ensureFolderExists(folder) {
  if (!folder) return null;
  try {
    if (!folder.exists) folder.create();
  } catch (e) {}
  return folder;
}

function processImageWithJson(imageFile, jsonFile, outputPSD, outputJPG) {
  // ===== OPEN IMAGE =====
  if (!imageFile.exists) {
    alert("❌ Image not found:\n" + imageFile.fsName);
    throw new Error("Image file not found");
  }
  log("Opening image: " + imageFile.fsName);
  var doc = app.open(imageFile);

  // ===== READ JSON =====
  if (!jsonFile.exists) {
    alert("❌ JSON not found:\n" + jsonFile.fsName);
    throw new Error("JSON file not found");
  }
  jsonFile.open("r");
  var jsonText = jsonFile.read();
  jsonFile.close();
  jsonText = normalizeZWNJInJsonText(jsonText);
  log("Loaded JSON: " + jsonFile.fsName);
  var data = JSON.parse(jsonText);

  // Accept { image_size, items } OR legacy array
  var items=null, srcW=null, srcH=null;
  if (data && data.items && data.image_size){
    items=normalizeItemsZWNJ(data.items); srcW=data.image_size.width; srcH=data.image_size.height;
  } else if (data instanceof Array){
    items=normalizeItemsZWNJ(data); srcW=doc.width.as('px'); srcH=doc.height.as('px');
  } else {
    throw new Error("JSON must contain { image_size, items } or be an array");
  }

  // ===== SCALE FACTORS =====
  var dstW = doc.width.as('px'), dstH = doc.height.as('px');
  var scaleX = (srcW && srcW>0) ? (dstW/srcW) : 1.0;
  var scaleY = (srcH && srcH>0) ? (dstH/srcH) : 1.0;
  log("Scale factors: src=(" + srcW + "x" + srcH + ") dst=(" + dstW + "x" + dstH + ") scaleX=" + scaleX + " scaleY=" + scaleY);

  // ===== PROCESS =====
  for (var i=0; i<items.length; i++){
    var item = items[i];
    if (!item) continue;

    var palette = buildItemPalette(item);
    var layoutCtx = null;
    function ensureLayoutContext() {
      if (!layoutCtx) layoutCtx = new TextMeasureContext(doc, fontName, baseSize);
      return layoutCtx;
    }

    var raw = toStr(item.text);
    raw = normalizeWSKeepBreaks(raw);
    raw = safeTrim(raw);
    if (!raw) continue;

    var preview = raw.length > 40 ? raw.substr(0,40) + "..." : raw;
    log("Item " + i + " bubble_id=" + (item.bubble_id || "N/A") + " text=\"" + preview + "\"");

    var hasManualBreaks = /\n/.test(raw);

    var c = deriveCenter(item);
    if (!c) {
      log("  -> no center, skipping");
      continue;
    }
    var cx = clampInt(c.x * scaleX), cy = clampInt(c.y * scaleY);
    log("  center=(" + cx + "," + cy + ")");

    removeOldTextSegments(doc, item, scaleX, scaleY, palette);

    var box = deriveBox(item);
    var PAD_FRAC = 0.10, MIN_W = 70, MIN_H = 60;

    var phys = box
      ? { left:box.left*scaleX, top:box.top*scaleY, right:box.right*scaleX, bottom:box.bottom*scaleY }
      : { left:cx-120, top:cy-100, right:cx+120, bottom:cy+100 };

    var bw = Math.max(MIN_W, phys.right - phys.left);
    var bh = Math.max(MIN_H, phys.bottom - phys.top);
    var pad = Math.round(Math.min(bw, bh) * PAD_FRAC);

    var innerW = Math.max(20, bw - 2*pad);
    var innerH = Math.max(20, bh - 2*pad);
    log("  bubbleSize=(" + bw + "x" + bh + ") pad=" + pad + " inner=(" + innerW + "x" + innerH + ")");

    var baseSize = (typeof item.size === "number" && item.size > 0) ? item.size : 28;
    var fontName  = getFontForType(item.bubble_type || "Standard");

    // build text seed
    var baseSeedText;
    var wrappedForced;
    if (hasManualBreaks) {
      var manualParts = raw.split(/\n+/);
      baseSeedText = collapseWhitespace(manualParts.join(" "));
      var manualLines = [];
      for (var m = 0; m < manualParts.length; m++) {
        manualLines.push(collapseWhitespace(manualParts[m]));
      }
      wrappedForced = manualLines.join("\r");
    } else {
      baseSeedText = collapseWhitespace(raw);
      wrappedForced = layoutBubble(baseSeedText, doc, fontName, baseSize, innerW, innerH, ensureLayoutContext());
      if (!wrappedForced) wrappedForced = baseSeedText;
    }

    // --- first pass: respect manual breaks (if any) ---
    var sizeForced = estimateSafeFontSizeForWrapped(wrappedForced, innerW, innerH, baseSize);

    // --- second pass: if manual breaks made font too small, ignore them ---
    var wrapped = wrappedForced;

    if (hasManualBreaks && sizeForced < baseSize * 0.7) {
      log("  manual breaks cause tiny size ("+sizeForced+") -> try reflow without forced lines");
      var wrappedFree = layoutBubble(baseSeedText, doc, fontName, baseSize, innerW, innerH, ensureLayoutContext());
      var sizeFree    = estimateSafeFontSizeForWrapped(wrappedFree, innerW, innerH, baseSize);
      log("  alt reflow size=" + sizeFree);
      if (sizeFree > sizeForced) {
        wrapped   = wrappedFree;
      }
    }

    log("  font=" + fontName);

    var linesArray = wrapped.replace(/\r\n/g, "\r").replace(/\n/g, "\r").split("\r");
    var bubbleRect = {
      x: cx - innerW / 2,
      y: cy - innerH / 2,
      width: innerW,
      height: innerH
    };
    var lyr = createPointTextFastFit(doc, linesArray, fontName, bubbleRect, "center");
    var ti = lyr.textItem;
    try { ti.color = palette.textColor; } catch (colorErr) { ti.color = solidBlack(); }
    log("  point-text fitted size=" + ti.size);

    var rot = deriveRotationDeg(item);
    if (rot && Math.abs(rot) > 0.001) {
      log("  rotation=" + rot);
      lyr.rotate(rot, AnchorPosition.MIDDLECENTER);
    }
    translateToCenter(lyr, cx, cy);
    scaleLayerToBubbleIfSmallText(lyr, ti, item, scaleX, scaleY, 0.9, 18, cx, cy);

    var needsStroke = item && !item.bbox_bubble && item.complex_background === true;
    if (needsStroke) {
      log('  complex background without bbox -> applying 3px stroke');
      try { doc.activeLayer = lyr; } catch (strokeErr) {}
      applyStrokeColor(lyr, 3, palette.strokeColor);
    }

    if (layoutCtx) {
      layoutCtx.dispose();
    }
  }

  // ===== SAVE =====
  var psdSaveOptions = new PhotoshopSaveOptions();
  psdSaveOptions.embedColorProfile = true;
  psdSaveOptions.layers = true;
  psdSaveOptions.maximizeCompatibility = true;
  psdSaveOptions.alphaChannels = true;
  psdSaveOptions.annotations = true;
  psdSaveOptions.spotColors = true;

  doc.saveAs(outputPSD, psdSaveOptions, true);
  log("✅ Done! Saved PSD: " + outputPSD.fsName);

  var finalAlert = "✅ Done! Saved PSD: " + outputPSD.fsName;

  if (EXPORT_PSD_TO_JPG) {
    var jpgTarget = outputJPG;
    if (!jpgTarget) {
      jpgTarget = File(outputPSD.fsName.replace(/\.psd$/i, ".jpg"));
    }

    var jpgOptions = new JPEGSaveOptions();
    jpgOptions.embedColorProfile = true;
    jpgOptions.quality = 12; // highest quality
    jpgOptions.formatOptions = FormatOptions.STANDARDBASELINE;
    jpgOptions.matte = MatteType.NONE;

    doc.saveAs(jpgTarget, jpgOptions, true);
    log("✅ Also exported JPG: " + jpgTarget.fsName);
    finalAlert += "\nJPG: " + jpgTarget.fsName;
  } else {
    log("✅ JPG export disabled by EXPORT_PSD_TO_JPG");
  }

  if (!PROCESS_WHOLE_CHAPTER) {
    alert(finalAlert);
  }

  doc.close(SaveOptions.DONOTSAVECHANGES);
}

function listImageFiles(folder) {
  if (!folder || !folder.exists) return [];
  return folder.getFiles(function (f) {
    return f instanceof File && /\.(png|jpe?g|tif?f?|psd)$/i.test(f.name);
  });
}

function runChapterBatch() {
  if (!chapterImagesFolder.exists) {
    alert("❌ Chapter images folder not found:\n" + chapterImagesFolder.fsName);
    throw new Error("Chapter images folder not found");
  }
  if (!chapterJsonFolder.exists) {
    alert("❌ Chapter JSON folder not found:\n" + chapterJsonFolder.fsName);
    throw new Error("Chapter JSON folder not found");
  }

  ensureFolderExists(chapterOutputFolder);

  var images = listImageFiles(chapterImagesFolder);
  if (!images.length) {
    alert("❌ No images found in:\n" + chapterImagesFolder.fsName);
    throw new Error("No images to process");
  }

  log("Starting chapter batch: " + images.length + " page(s)");

  for (var i = 0; i < images.length; i++) {
    var img = images[i];
    var base = img.name.replace(/\.[^.]+$/, "");
    var jsonPath = File(chapterJsonFolder + "/" + base + ".json");

    if (!jsonPath.exists) {
      log("⚠️ Skipping " + img.name + " (missing JSON: " + jsonPath.fsName + ")");
      continue;
    }

    var outPath = File(chapterOutputFolder + "/" + base + "_output.psd");
    var outJpg  = File(chapterOutputFolder + "/" + base + "_output.jpg");
    processImageWithJson(img, jsonPath, outPath, outJpg);
  }

  alert("✅ Chapter batch complete! Saved to: " + chapterOutputFolder.fsName);
}

// ===== ENTRY POINT =====
try {
  if (PROCESS_WHOLE_CHAPTER) {
    runChapterBatch();
  } else {
    processImageWithJson(imageFile, jsonFile, outputPSD, outputJPG);
  }
} finally {
  __restorePrefs();
}
