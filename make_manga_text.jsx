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
var imageFile   = File(scriptFolder + "/RCO018_1462962423.jpg");
var jsonFile    = File(scriptFolder + "/positions.json");
var outputPSD   = File(scriptFolder + "/manga_output.psd");

// ===== DEBUG + FILE LOGGING =====
var DEBUG = true;
var LOG_TO_FILE = true;
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

// ===== Auto fit (shrink only) =====
function autoFitTextLayer(lyr, ti, cx, cy, innerW, innerH, minSize, maxSize) {
  if (minSize === undefined) minSize = 10;
  if (maxSize === undefined) maxSize = 52;

  log("  [autoFit] start size=" + ti.size + " innerW=" + innerW + " innerH=" + innerH);

  for (var step = 0; step < 15; step++) {
    var b = layerBoundsPx(lyr);
    var w = b.width;
    var h = b.height;
    var overflow = (w > innerW + 1 || h > innerH + 1);

    log("    step " + step + " size=" + ti.size +
        " bounds=(" + w.toFixed(1) + "x" + h.toFixed(1) + ")" +
        " overflow=" + overflow);

    if (!overflow) {
      log("    no overflow, done.");
      break;
    }

    if (overflow && ti.size > minSize) {
      var wRatio = innerW / Math.max(1, w);
      var hRatio = innerH / Math.max(1, h);
      var scaleDown = Math.min(wRatio, hRatio) * 0.95;
      var newSizeDown = Math.max(minSize, Math.floor(ti.size * scaleDown));

      if (newSizeDown < ti.size) {
        log("      shrink -> " + newSizeDown + " (scaleDown=" + scaleDown.toFixed(3) + ")");
        ti.size = newSizeDown;
        try { ti.leading = Math.max(1, Math.floor(newSizeDown * 1.12)); } catch (e) {}
        translateToCenter(lyr, cx, cy);
        continue;
      }
    }
    log("    cannot shrink further or no change; break.");
    break;
  }

  var bb = layerBoundsPx(lyr);
  if (bb.height >= innerH - 2) {
    var safeH = Math.max(innerH + 30, bb.height + 10);
    log("  [autoFit] safety height increase to " + safeH);
    ti.height = safeH;
  }
  translateToCenter(lyr, cx, cy);
}

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
  var sizeForced    = estimateSafeFontSizeForWrapped(wrappedForced, innerW, innerH, baseSize);

  // --- second pass: if manual breaks made font too small, ignore them ---
  var wrapped = wrappedForced;
  var finalSize = sizeForced;

  if (hasManualBreaks && sizeForced < baseSize * 0.7) {
    log("  manual breaks cause tiny size ("+sizeForced+") -> try reflow without forced lines");
    var wrappedFree = layoutBubble(baseSeedText, doc, fontName, baseSize, innerW, innerH);
    var sizeFree    = estimateSafeFontSizeForWrapped(wrappedFree, innerW, innerH, baseSize);
    log("  alt reflow size=" + sizeFree);
    if (sizeFree > sizeForced) {
      wrapped   = wrappedFree;
      finalSize = sizeFree;
    }
  }

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
  var MAX_SIZE = 60; // just a cap; we only shrink in autoFit
  autoFitTextLayer(lyr, ti, cx, cy, innerW, innerH, MIN_SIZE, MAX_SIZE);

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
