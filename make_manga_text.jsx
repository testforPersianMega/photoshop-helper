// make_manga_text_v2.jsx — Accurate center placement for manga balloons
// - Uses image_size scaling
// - Prefers tight text bbox center; falls back to polygon centroid if provided
// - Rotation-aware (rotation_deg)
// - Snaps layer to visual center after rendering
// - Designed for Persian translations (centered paragraph text)

#target photoshop
app.displayDialogs = DialogModes.NO;
app.preferences.rulerUnits = Units.PIXELS;

// ===== USER CONFIG =====
var scriptFolder = Folder("C:/Users/abbas/Desktop/psd maker");  // change if needed
var imageFile   = File(scriptFolder + "/94107d26-f530-4994-8a94-e48a6e70777c.png");
var jsonFile    = File(scriptFolder + "/positions.json");
var outputPSD   = File(scriptFolder + "/manga_output.psd");

// Optional batch mode: process an entire chapter
// Put all page images inside `chapterImagesFolder` and a matching JSON per image
// (same base name, .json extension) inside `chapterJsonFolder`.
// Outputs will be written to `chapterOutputFolder` using `<basename>_output.psd`.
var PROCESS_WHOLE_CHAPTER = false;
var chapterImagesFolder  = Folder(scriptFolder + "/chapter_images");
var chapterJsonFolder    = Folder(scriptFolder + "/chapter_json");
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
function safeTrim(s){ return s.replace(/^[\s\u00A0]+/, "").replace(/[\s\u00A0]+$/, ""); }
function collapseWhitespace(s){ return s.replace(/\s+/g, " "); }

// Keep line breaks, normalize spaces, and convert literal "\n" to real newlines
function normalizeWSKeepBreaks(s){
  var str = toStr(s);
  if (!str) return "";
  var norm = str.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  norm = norm.replace(/\\n/g, "\n");
  var parts = norm.split("\n");
  for (var i=0; i<parts.length; i++){
    parts[i] = parts[i].replace(/[ \t\u00A0]+/g, " ");
  }
  return parts.join("\n");
}

// ===== Geometry helpers =====
function solidBlack(){ var c=new SolidColor(); c.rgb.red=0;c.rgb.green=0;c.rgb.blue=0; return c; }
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
  if (!box) return null;
  var p = (typeof pad === "number") ? pad : 0;
  return {
    x_min: Number(box.left)   - p,
    x_max: Number(box.right)  + p,
    y_min: Number(box.top)    - p,
    y_max: Number(box.bottom) + p
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

function removeOldTextSegments(doc, item, scaleX, scaleY){
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
  if (segmentsInfo.source !== 'json') {
    log('  deriving removal segments from ' + segmentsInfo.source);
  }
  var selectionMade = selectSegments(doc, segmentsInfo.segments, scaleX, scaleY);
  if (!selectionMade) {
    log('  ⚠️ segments provided but no valid selection could be made');
    return;
  }
  highlightSelectionForDebug(doc, cleanupLayer);
  log('  removing original text via content-aware fill over ' + segmentsInfo.segments.length + ' segments');
  var filled = contentAwareFillSelection();
  try { doc.selection.deselect(); } catch (e2) {}
  if (!filled) {
    log('  ⚠️ content-aware fill skipped due to error');
  }
}

// ===== Centers/boxes =====
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
    var bt=item.bbox_text;
    return { x:(bt.left+bt.right)/2, y:(bt.top+bt.bottom)/2 };
  }
  if (item && item.bbox_bubble){
    var bb=item.bbox_bubble;
    return { x:(bb.left+bb.right)/2, y:(bb.top+bb.bottom)/2 };
  }
  return null;
}
function deriveBox(item){
  if (item && item.bbox_text)
    return { left:item.bbox_text.left, top:item.bbox_text.top, right:item.bbox_text.right, bottom:item.bbox_text.bottom };
  if (item && item.bbox_bubble)
    return { left:item.bbox_bubble.left, top:item.bbox_bubble.top, right:item.bbox_bubble.right, bottom:item.bbox_bubble.bottom };
  return null;
}

// ===== Fonts & RTL helpers =====
function getFontForType(type){
  switch(type){
    case "Standard Speech": return "Potk-Black";
    case "Thought":         return "B Titr";
    case "Shouting/Emotion":return "Impact";
    case "Whisper/Soft":    return "B Nazanin Light";
    case "Electronic":      return "Consolas";
    case "Narration Box":   return "B Mitra";
    case "Distorted/Custom":return "Shabnam-BoldItalic";
    default:                return "ArialMT";
  }
}

function forceRTL(s){
  var RLE="\u202B", PDF="\u202C", RLM="\u200F";
  var str = toStr(s);
  if (!str) return RLM;
  var norm = str.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  var parts = norm.split("\n");
  for (var i=0; i<parts.length; i++){
    parts[i] = RLM + RLE + parts[i] + PDF;
  }
  return parts.join("\r");
}

function applyParagraphDirectionRTL() {
  try {
    if (app.activeDocument.activeLayer.kind !== LayerKind.TEXT) return;
    var s2t = stringIDToTypeID, c2t = charIDToTypeID;
    var ti  = app.activeDocument.activeLayer.textItem;
    var txt = ti.contents;
    var len = txt.length;

    var desc = new ActionDescriptor();
    var ref  = new ActionReference();
    ref.putEnumerated(s2t('textLayer'), c2t('Ordn'), c2t('Trgt'));
    desc.putReference(c2t('null'), ref);

    var textDesc = new ActionDescriptor();
    textDesc.putString(s2t('textKey'), txt);

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
  ti.font = fontName;
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
  this.textItem.contents = text;
  var bounds = layerBoundsPx(this.layer);
  return bounds.width;
};

TextMeasureContext.prototype.lineHeight = function(){
  return Math.max(1, Math.round(this.size * 1.18) + 4);
};

function splitWordsForLayout(text) {
  var s = safeTrim(toStr(text));
  if (!s) return [];
  var raw = s.split(/\s+/);
  var words = [];
  for (var i = 0; i < raw.length; i++) {
    if (raw[i]) words.push(raw[i]);
  }
  return words;
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

function layoutBubbleLines(text, doc, fontName, sizePx, maxWidth, maxHeight) {
  var widthLimit = Math.max(1, maxWidth || 1);
  var ctx = new TextMeasureContext(doc, fontName, sizePx);
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
    ctx.dispose();
  }
}

function layoutBubble(text, doc, fontName, sizePx, maxWidth, maxHeight) {
  var linesWords = layoutBubbleLines(text, doc, fontName, sizePx, maxWidth, maxHeight);
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
function createParagraphFullBox(doc, text, fontName, sizePx, cx, cy, innerW, innerH){
  var left = cx - innerW/2, top = cy - innerH/2;
  var lyr = doc.artLayers.add();
  lyr.kind = LayerKind.TEXT;
  var ti = lyr.textItem;
  ti.kind = TextType.PARAGRAPHTEXT;
  ti.contents = forceRTL(text);
  ti.font = fontName;
  ti.size = sizePx;
  try { ti.leading = Math.max(1, Math.floor(sizePx * 1.18)); } catch(e){}
  ti.justification = Justification.CENTER;
  ti.position = [left, top];
  ti.width  = innerW;
  ti.height = innerH;
  try { ti.color = app.foregroundColor; } catch(e){ ti.color = solidBlack(); }
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

function cleanupMeasureLayer(layer) {
  if (!layer) return;
  try { layer.remove(); } catch (e) {}
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

  log("  [autoFit] innerW=" + innerW + " innerH=" + innerH +
      (longestLine ? " longestLine=\"" + longestLine + "\"" : ""));

  function estimateContentHeight(sizePx) {
    var lineHeight = Math.max(1, Math.floor(sizePx * 1.18));
    return lineHeight * lineCount;
  }

  function sizeFits(sizePx) {
    ti.size = sizePx;
    try { ti.leading = Math.max(1, Math.floor(sizePx * 1.12)); } catch (e) {}
    translateToCenter(lyr, cx, cy);

    var b = layerBoundsPx(lyr);
    var w = b.width;
    var h = b.height;

    var measuredLineWidth = measureLayer ? measureLineWidthFromLayer(measureLayer, sizePx) : w;
    var widthOverflow  = (measuredLineWidth > innerW + 1);
    var estHeight = estimateContentHeight(sizePx);
    var heightOverflow = (estHeight > innerH + 1);
    var overflow = widthOverflow || heightOverflow;

    log("    test size=" + sizePx +
        " bounds=(" + w.toFixed(1) + "x" + h.toFixed(1) + ")" +
        (measureLayer ? " longestLineWidth=" + measuredLineWidth.toFixed(1) : "") +
        " estHeight=" + estHeight.toFixed(1) +
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

  var finalHeightEstimate = estimateContentHeight(best);
  log("  [autoFit] finalSize=" + best + " estHeight=" + finalHeightEstimate);

  sizeFits(best);

  cleanupMeasureLayer(measureLayer);

  var safeHeight = Math.max(innerH, finalHeightEstimate + 20);
  if (ti.height < safeHeight) {
    ti.height = safeHeight;
    log("  [autoFit] expand text box height to " + safeHeight);
  }

  translateToCenter(lyr, cx, cy);
}

function ensureFolderExists(folder) {
  if (!folder) return null;
  try {
    if (!folder.exists) folder.create();
  } catch (e) {}
  return folder;
}

function processImageWithJson(imageFile, jsonFile, outputPSD) {
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
  log("Loaded JSON: " + jsonFile.fsName);
  var data = JSON.parse(jsonText);

  // Accept { image_size, items } OR legacy array
  var items=null, srcW=null, srcH=null;
  if (data && data.items && data.image_size){
    items=data.items; srcW=data.image_size.width; srcH=data.image_size.height;
  } else if (data instanceof Array){
    items=data; srcW=doc.width.as('px'); srcH=doc.height.as('px');
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

    removeOldTextSegments(doc, item, scaleX, scaleY);

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
    var fontName  = getFontForType(item.bubble_type || "Standard Speech");

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
      wrappedForced = layoutBubble(baseSeedText, doc, fontName, baseSize, innerW, innerH);
      if (!wrappedForced) wrappedForced = baseSeedText;
    }

    // --- first pass: respect manual breaks (if any) ---
    var sizeForced = estimateSafeFontSizeForWrapped(wrappedForced, innerW, innerH, baseSize);

    // --- second pass: if manual breaks made font too small, ignore them ---
    var wrapped = wrappedForced;

    if (hasManualBreaks && sizeForced < baseSize * 0.7) {
      log("  manual breaks cause tiny size ("+sizeForced+") -> try reflow without forced lines");
      var wrappedFree = layoutBubble(baseSeedText, doc, fontName, baseSize, innerW, innerH);
      var sizeFree    = estimateSafeFontSizeForWrapped(wrappedFree, innerW, innerH, baseSize);
      log("  alt reflow size=" + sizeFree);
      if (sizeFree > sizeForced) {
        wrapped   = wrappedFree;
      }
    }

    var finalSize = baseSize;
    log("  startSize(final)=" + finalSize + " font=" + fontName);

    var lyr = createParagraphFullBox(doc, wrapped, fontName, finalSize, cx, cy, innerW, innerH);
    var ti  = lyr.textItem;

    app.activeDocument.activeLayer = lyr;
    applyParagraphDirectionRTL();
    trySetMEEveryLineComposer();

    ti = lyr.textItem;
    ti.justification = Justification.CENTER;

    ti.position = [cx - ti.width/2, cy - ti.height/2];
    translateToCenter(lyr, cx, cy);

    var MIN_SIZE = 12;
    var MAX_SIZE = 60; // cap for binary search auto-fit
    autoFitTextLayer(lyr, ti, cx, cy, innerW, innerH, MIN_SIZE, MAX_SIZE, wrapped);

    var rot = deriveRotationDeg(item);
    if (rot && Math.abs(rot) > 0.001) {
      log("  rotation=" + rot);
      lyr.rotate(rot, AnchorPosition.MIDDLECENTER);
    }
    translateToCenter(lyr, cx, cy);
  }

  // ===== SAVE =====
  var psdSaveOptions = new PhotoshopSaveOptions();
  doc.saveAs(outputPSD, psdSaveOptions, true);
  doc.close(SaveOptions.DONOTSAVECHANGES);
  log("✅ Done! Saved: " + outputPSD.fsName);
  alert("✅ Done! Saved: " + outputPSD.fsName);
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
    processImageWithJson(img, jsonPath, outPath);
  }

  alert("✅ Chapter batch complete! Saved to: " + chapterOutputFolder.fsName);
}

// ===== ENTRY POINT =====
if (PROCESS_WHOLE_CHAPTER) {
  runChapterBatch();
} else {
  processImageWithJson(imageFile, jsonFile, outputPSD);
}
