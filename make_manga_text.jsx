// make_manga_text_v2.jsx — Accurate center placement for manga balloons
// - Uses image_size scaling
// - Prefers tight text bbox center; falls back to polygon centroid if provided
// - Rotation-aware (rotation_deg)
// - Snaps layer to visual center after rendering
// - Designed for Persian translations (centered point text works well)
//
// Expected JSON shape:
// {
//   "image_size": { "width": <int>, "height": <int> },
//   "items": [
//     {
//       "bubble_id": "p1-b01",
//       "cluster_id": "c1",
//       "order": 1,
//       "text_original": "...",
//       "text": "<Persian translation>",
//       "center": { "x": <int>, "y": <int> },
//       "bbox_text": { "left": <int>, "top": <int>, "right": <int>, "bottom": <int> },
//       "bbox_bubble": { "left": <int>, "top": <int>, "right": <int>, "bottom": <int> }, // optional
//       "polygon_text": [[x0,y0],[x1,y1],...], // optional, clockwise
//       "rotation_deg": <number>,               // optional
//       "writing_mode": "<horizontal|vertical-rl|vertical-lr>", // optional
//       "bubble_type": "<Standard Speech|Thought|...>",         // optional
//       "confidence": <number>,                 // optional
//       "size": <int>                           // optional
//     }
//   ]
// }

// make_manga_text_RTL_fullbox.jsx
// Multi-line first (via line breaks), keep full bubble-sized text box,
// shrink font only if needed, RTL paragraph direction, centered text.

#target photoshop
app.displayDialogs = DialogModes.NO;
app.preferences.rulerUnits = Units.PIXELS;

// ===== USER CONFIG =====
var scriptFolder = Folder("C:/Users/abbas/Desktop/psd maker");
var imageFile   = File(scriptFolder + "/94107d26-f530-4994-8a94-e48a6e70777c.png");
var jsonFile    = File(scriptFolder + "/positions.json");
var outputPSD   = File(scriptFolder + "/manga_output.psd");

// ===== Legacy-safe JSON =====
if (typeof JSON === 'undefined') { JSON = {}; JSON.parse = function (s) { return eval('(' + s + ')'); }; }

// ===== Utils =====
function solidBlack(){ var c=new SolidColor(); c.rgb.red=0;c.rgb.green=0;c.rgb.blue=0; return c; }
function clampInt(v){ return Math.round(v); }
function toStr(v){ if (v===undefined||v===null) return ""; try{
  if (typeof v==="object"){ if (v.valueOf&&typeof v.valueOf()!=="object") return String(v.valueOf());
    try { return JSON.stringify(v); } catch(e2){ return String(v); } }
  return String(v);
} catch(e){ return ""; } }
function safeTrim(s){ return s.replace(/^[\s\u00A0]+/,"").replace(/[\s\u00A0]+$/,""); }
function normalizeWS(s){ return s.replace(/\s+/g," "); }

function layerBoundsPx(lyr){ var b=lyr.bounds; return {
  left:b[0].as('px'), top:b[1].as('px'), right:b[2].as('px'), bottom:b[3].as('px'),
  width:b[2].as('px')-b[0].as('px'), height:b[3].as('px')-b[1].as('px')
};}
function layerCenterPx(lyr){ var bb=layerBoundsPx(lyr); return { x:(bb.left+bb.right)/2, y:(bb.top+bb.bottom)/2 }; }
function translateToCenter(lyr,cx,cy){ var c=layerCenterPx(lyr); lyr.translate(cx-c.x, cy-c.y); }
function fitsWithinRect(lyr,L,T,R,B){ var b=layerBoundsPx(lyr); return (b.left>=L-0.5&&b.top>=T-0.5&&b.right<=R+0.5&&b.bottom<=B+0.5); }

// Centers/boxes
function polygonCentroid(points){ var n=points.length;if(n<3)return null; var A=0,Cx=0,Cy=0, i,x1,y1,x2,y2,c;
  for(i=0;i<n;i++){ x1=points[i][0]; y1=points[i][1]; x2=points[(i+1)%n][0]; y2=points[(i+1)%n][1];
    c=(x1*y2 - x2*y1); A+=c; Cx+=(x1+x2)*c; Cy+=(y1+y2)*c; } A/=2;
  if (Math.abs(A)<1e-6){ var sx=0,sy=0; for(i=0;i<n;i++){ sx+=points[i][0]; sy+=points[i][1]; } return {x:sx/n,y:sy/n}; }
  return { x: Cx/(6*A), y: Cy/(6*A) };
}
function deriveCenter(item){
  if (item && item.polygon_text && item.polygon_text.length>=3){ var c=polygonCentroid(item.polygon_text); if(c) return c; }
  if (item && item.center && typeof item.center.x==="number" && typeof item.center.y==="number") return {x:item.center.x,y:item.center.y};
  if (item && item.bbox_text){ var bt=item.bbox_text; return { x:(bt.left+bt.right)/2, y:(bt.top+bt.bottom)/2 }; }
  if (item && item.bbox_bubble){ var bb=item.bbox_bubble; return { x:(bb.left+bb.right)/2, y:(bb.top+bb.bottom)/2 }; }
  return null;
}
function deriveBox(item){
  if (item && item.bbox_text)   return {left:item.bbox_text.left, top:item.bbox_text.top, right:item.bbox_text.right, bottom:item.bbox_text.bottom};
  if (item && item.bbox_bubble) return {left:item.bbox_bubble.left, top:item.bbox_bubble.top, right:item.bbox_bubble.right, bottom:item.bbox_bubble.bottom};
  return null;
}

// Fonts & RTL
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
// Force RTL direction using Unicode marks (fallback) while we also try to set Photoshop paragraphDirection to RTL.
function forceRTL(s){ var RLE="\u202B", PDF="\u202C", RLM="\u200F"; return RLM + RLE + s + PDF; }

// Try to set paragraph direction to RTL using Action Manager (works on ME-enabled PS).
function setParagraphDirectionRTL(){
  try{
    var idsetd = charIDToTypeID("setd");
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putEnumerated(stringIDToTypeID("textLayer"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
    desc.putReference(charIDToTypeID("null"), ref);

    var textDesc = new ActionDescriptor();
    var paraStyleList = new ActionList();
    var paraStyle = new ActionDescriptor();
    paraStyle.putEnumerated(stringIDToTypeID("paragraphDirection"),
                            stringIDToTypeID("paragraphDirection"),
                            stringIDToTypeID("RightToLeftParagraph"));
    paraStyleList.putObject(stringIDToTypeID("paragraphStyle"), paraStyle);
    textDesc.putList(stringIDToTypeID("paragraphStyle"), paraStyleList);

    desc.putObject(charIDToTypeID("T   "), stringIDToTypeID("textLayer"), textDesc);
    executeAction(idsetd, desc, DialogModes.NO);
  }catch(e){ /* silently ignore; Unicode RTL marks still enforce RTL */ }
}

function deriveRotationDeg(item){
  if (typeof item.rotation_deg === "number") return item.rotation_deg;
  if (item.writing_mode === "vertical-rl" || item.writing_mode === "vertical-lr") return 90;
  return 0;
}

// ==== Multi-line first: diamond breaks (short → long → short when many lines) ====
function estMaxLines(innerH, sizePx){ var leading=Math.max(1, Math.floor(sizePx*1.18)); return Math.max(1, Math.floor(innerH/leading)); }
function estCharsPerLine(innerW, sizePx){ var avg=Math.max(3, sizePx*0.6); return Math.max(8, Math.floor(innerW/avg)); }
function chooseLineCount(totalChars, maxCPL, maxLines){ var need=Math.max(1, Math.ceil(totalChars/Math.max(1,maxCPL))); return Math.min(Math.max(need,1), Math.max(maxLines,1)); }
function capsDiamond(L, baseCap){ var caps=[], mid=(L-1)/2.0; for (var i=0;i<L;i++){ var dist=Math.abs(i-mid)/Math.max(1,mid); var factor=1.2 - 0.5*dist; caps.push(Math.max(6, Math.floor(baseCap*factor))); } return caps; }
function diamondWrap(text, innerW, innerH, sizePx){
  var s = toStr(text); s = normalizeWS(s); s = safeTrim(s); if (!s) return "";
  var words = s.split(" "); if (words.length <= 3) return s;
  var maxLines=estMaxLines(innerH,sizePx), maxCPL=estCharsPerLine(innerW,sizePx), totalChars=s.length;
  var L = chooseLineCount(totalChars,maxCPL,maxLines); if (L<=1) return s;
  var caps = (L>=4)? capsDiamond(L,maxCPL) : (function(){ var a=[]; for (var i=0;i<L;i++) a[i]=Math.max(6, Math.floor(maxCPL*0.9)); return a; })();
  var lines=[], cur=[], curLen=0, idx=0, cap=caps[idx];
  for (var w=0; w<words.length; w++){
    var token=words[w], extra=(curLen===0?token.length:token.length+1);
    if (curLen+extra<=cap){ cur.push(token); curLen+=extra; }
    else { lines.push(cur.join(" ")); cur=[token]; curLen=token.length; idx=Math.min(idx+1,caps.length-1); cap=caps[idx]; }
  }
  if (cur.length) lines.push(cur.join(" "));
  while (lines.length > L){ var last = lines.pop(); lines[lines.length-1] = lines[lines.length-1] + " " + last; }
  return lines.join("\r"); // Photoshop newline
}

// Create paragraph text layer exactly the inner bubble size (no shrinking of the layer)
function createParagraphFullBox(doc, text, fontName, sizePx, cx, cy, innerW, innerH){
  var left = cx - innerW/2, top = cy - innerH/2;
  var lyr = doc.artLayers.add();
  lyr.kind = LayerKind.TEXT;
  var ti = lyr.textItem;
  ti.kind = TextType.PARAGRAPHTEXT;
  ti.contents = forceRTL(text);                 // fallback RTL marks
  ti.font = fontName;
  ti.size = sizePx;
  try { ti.leading = Math.max(1, Math.floor(sizePx * 1.18)); } catch(e){}
  ti.justification = Justification.CENTER;      // centered lines
  ti.position = [left, top];
  ti.width = innerW; ti.height = innerH;       // keep full bubble size
  try { ti.color = app.foregroundColor; } catch(e){ ti.color = solidBlack(); }

  // Try to set real RTL paragraph direction (ME builds). If it fails, marks still do RTL.
  try { doc.activeLayer = lyr; } catch(e){}
  setParagraphDirectionRTL();
  return lyr;
}

// Fast size fallback (ratio shrink, ≤ 2 passes)
function quickShrinkToFit(lyr, ti, innerW, innerH){
  var b = layerBoundsPx(lyr);
  var wRatio = innerW / Math.max(1, b.width);
  var hRatio = innerH / Math.max(1, b.height);
  var scale = Math.min(wRatio, hRatio) * 0.98; // tiny safety
  if (scale < 1.0){
    var newSize = Math.max(8, Math.floor(ti.size * scale));
    ti.size = newSize; try { ti.leading = Math.max(1, Math.floor(newSize * 1.12)); } catch(e){}
    // micro pass if still a hair over
    b = layerBoundsPx(lyr);
    if (b.width > innerW || b.height > innerH){
      var w2 = innerW / Math.max(1, b.width), h2 = innerH / Math.max(1, b.height);
      var s2 = Math.min(w2, h2) * 0.995;
      if (s2 < 1.0){
        newSize = Math.max(8, Math.floor(ti.size * s2));
        ti.size = newSize; try { ti.leading = Math.max(1, Math.floor(newSize * 1.12)); } catch(e){}
      }
    }
    return true;
  }
  return false;
}

// ===== OPEN IMAGE =====
if (!imageFile.exists) { alert("❌ Image not found:\n"+imageFile.fsName); throw new Error("Image file not found"); }
var doc = app.open(imageFile);

// ===== READ JSON =====
if (!jsonFile.exists) { alert("❌ JSON not found:\n"+jsonFile.fsName); throw new Error("JSON file not found"); }
jsonFile.open("r"); var jsonText = jsonFile.read(); jsonFile.close();
var data = JSON.parse(jsonText);

// Accept { image_size, items } OR legacy array
var items=null, srcW=null, srcH=null;
if (data && data.items && data.image_size){ items=data.items; srcW=data.image_size.width; srcH=data.image_size.height; }
else if (data instanceof Array){ items=data; srcW=doc.width.as('px'); srcH=doc.height.as('px'); }
else { throw new Error("JSON must contain { image_size, items } or be an array"); }

// ===== SCALE FACTORS =====
var dstW = doc.width.as('px'), dstH = doc.height.as('px');
var scaleX = (srcW && srcW>0) ? (dstW/srcW) : 1.0;
var scaleY = (srcH && srcH>0) ? (dstH/srcH) : 1.0;

// ===== PROCESS =====
for (var i=0; i<items.length; i++){
  var item = items[i]; if (!item) continue;

  var raw = toStr(item.text); raw = normalizeWS(raw); raw = safeTrim(raw);
  if (!raw) continue;

  var c = deriveCenter(item); if(!c) continue;
  var cx = clampInt(c.x * scaleX), cy = clampInt(c.y * scaleY);

  var box = deriveBox(item);
  var PAD_FRAC = 0.10, MIN_W = 70, MIN_H = 60;

  var phys = box ? { left:box.left*scaleX, top:box.top*scaleY, right:box.right*scaleX, bottom:box.bottom*scaleY }
                 : { left: cx-120, top: cy-100, right: cx+120, bottom: cy+100 };

  var bw = Math.max(MIN_W, phys.right - phys.left);
  var bh = Math.max(MIN_H, phys.bottom - phys.top);
  var pad = Math.round(Math.min(bw, bh) * PAD_FRAC);

  // Inner paragraph box = FULL bubble size minus padding (never shrinks)
  var innerW = Math.max(20, bw - 2*pad);
  var innerH = Math.max(20, bh - 2*pad);

  // 1) MULTI-LINE FIRST (insert diamond line breaks), fixed full-size box
  var startSize = (typeof item.size === "number" && item.size > 0) ? item.size : 28;
  var wrapped = diamondWrap(raw, innerW, innerH, startSize);
  var fontName = getFontForType(item.bubble_type || "Standard Speech");

  var lyr = createParagraphFullBox(doc, wrapped, fontName, startSize, cx, cy, innerW, innerH);
  var ti = lyr.textItem;

  // 2) SIZE FALLBACK (only if still overflowing)
  var L = cx - innerW/2, T = cy - innerH/2, R = cx + innerW/2, B = cy + innerH/2;
  if (!fitsWithinRect(lyr, L, T, R, B)){
    quickShrinkToFit(lyr, ti, innerW, innerH);
  }

  // Rotation + final snap
  var rot = deriveRotationDeg(item);
  if (rot && Math.abs(rot)>0.001) { lyr.rotate(rot, AnchorPosition.MIDDLECENTER); }
  translateToCenter(lyr, cx, cy);
}

// ===== SAVE =====
var psdSaveOptions = new PhotoshopSaveOptions();
doc.saveAs(outputPSD, psdSaveOptions, true);
doc.close(SaveOptions.DONOTSAVECHANGES);
alert("✅ Done! Saved: " + outputPSD.fsName);
