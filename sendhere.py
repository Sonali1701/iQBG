/***********************
 * PAPER GENERATOR (FACULTY PDF) + DTP FINALIZER (DOCX/PPT/ZIP)
 * -------------------------------------------------------------
 * CHUNKED MERGE DRAFT
 *  - keeps your selection / replace / preview business logic
 *  - adds real preview images for Docs-only sources
 *  - adds PPT output
 *  - adds ZIP bundle output
 *
 * NOTE:
 *  Preview images here are generated from Google Docs blocks,
 *  exported to PDF, then rendered via Drive thumbnail.
 *  That is the correct approach for Docs / Word-only sources.
 ***********************/

/***********************
 * CONFIG
 ***********************/
var AITS_SHEET_NAME = "AITS+Real Test";
var TOPICS_SHEET_NAME = "Topics";
var OUTPUT_FOLDER_ID = "1m2Cm73tDmaZQ-rMJRmaXU74NGuHTH-Ig";
var TEMPLATE_DOC_ID = "1lvHgZ0-403gbOxzyDjR6Ljg3opXHGMe8rN5kyvSysIg";
var DTP_TEMPLATE_DOC_ID = "1lvHgZ0-403gbOxzyDjR6Ljg3opXHGMe8rN5kyvSysIg";
var PPT_TEMPLATE_FILE_ID = "";
var PPT_ANCHOR_SHAPE_NAME = "SCREENSHOT_BOX";
var PPT_TARGET_FRACTION = 0.325;
var FORCE_TEMPLATE_REFRESH_IF_DOCX = false;
var SAFE_SPACE = " ";
var PREVIEW_CACHE_FOLDER_NAME = "Preview Cache";

var PROJECT_OPTIONS = [
  "Milestone","VPRP","NSAT","FTP","FRT","AIR","AITS","Real Test",
  "Backlog killer Series","PW Universe","Mission 100 Test Series"
];

var ACTIVE_OUTPUT_FOLDER_ID = OUTPUT_FOLDER_ID;

/***********************
 * NORMALIZATION HELPERS
 ***********************/
function normStr_(v){ return String(v==null?"":v).trim().toLowerCase(); }
function collapseSpaces_(s){
  return String(s==null?"":s).replace(/\u00A0/g," ").replace(/\s+/g," ").trim();
}
function normClass_(v){
  var s = collapseSpaces_(v);
  var m = s.match(/\d+/);
  return m ? m[0] : s.trim();
}
function isClass11plus12_(className){
  var s = String(className||"").replace(/\s+/g,"");
  return s === "11+12" || s === "12+11" || s === "11&12" || s === "11-12";
}
function classSetFromReq_(className){
  if (isClass11plus12_(className)) return {"11":true,"12":true};
  var c = normClass_(className);
  if (c === "11" || c === "12") { var o={}; o[c]=true; return o; }
  var o2={}; o2[c]=true; return o2;
}

/***********************
 * CHAPTER FUZZY MATCH HELPERS
 ***********************/
function singularizeToken_(tok){
  if (tok.length > 4 && tok.endsWith("s")) return tok.slice(0,-1);
  return tok;
}
function normChapterKey_(s){
  s = collapseSpaces_(s).toLowerCase();
  s = s.replace(/&/g," and ").replace(/[^a-z0-9 ]+/g," ");
  s = collapseSpaces_(s);
  var toks = s.split(" ").map(singularizeToken_).filter(Boolean);
  return toks.join(" ");
}
function tokenSet_(s){
  var key = normChapterKey_(s);
  var toks = key.split(" ").filter(Boolean);
  var set = {};
  for (var i=0;i<toks.length;i++) set[toks[i]] = true;
  return set;
}
function jaccard_(a,b){
  var inter=0, uni=0, seen={};
  for (var k in a) if (a.hasOwnProperty(k)){
    uni++; seen[k]=true;
    if (b[k]) inter++;
  }
  for (k in b) if (b.hasOwnProperty(k) && !seen[k]) uni++;
  return uni ? inter/uni : 0;
}

/***********************
 * FILE NAMING
 ***********************/
function qpFileName_(paperName){ return paperName + " - Question Paper"; }
function akhsFileName_(paperName){ return paperName + " - Answer Key & Hints and Solution"; }
function akFileName_(paperName){ return paperName + " - Answer Key"; }
function tgDocFileName_(paperName){ return paperName + " - Tagging Sheet"; }
function pptFileName_(paperName){ return paperName + " - Questions"; }
function zipFileName_(paperName){ return paperName + " - Bundle.zip"; }

/***********************
 * SUBJECT START NUMBER / FOLDER ROUTING
 ***********************/
function getQuestionStartNo_(subject){
  var s = normStr_(subject);
  if (s==="physics") return 1;
  if (s==="chemistry") return 46;
  if (s==="botany") return 91;
  if (s==="zoology") return 136;
  return 1;
}
function getSubjectFolderName_(subject){
  var s = normStr_(subject);
  if (s==="physics") return "Physics";
  if (s==="chemistry") return "Chemistry";
  if (s==="botany") return "Botany";
  if (s==="zoology") return "Zoology";
  return "";
}
function getSubjectOutputFolderId_(subject){
  var name = getSubjectFolderName_(subject);
  if (!name) return OUTPUT_FOLDER_ID;
  var parent = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next().getId();
  return parent.createFolder(name).getId();
}

/***********************
 * WEB APP
 ***********************/
function doGet(e){
  try {
    e = e || {};
    var page = (e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : "faculty";

    if (page === "dtp") {
      var token = (e.parameter && e.parameter.token) ? String(e.parameter.token).trim() : "";
      var expected = getDtpToken_();
      if (!expected) {
        return HtmlService.createHtmlOutput(
          "<pre>DTP is not enabled.\nSet Script Property: DTP_TOKEN</pre>"
        ).setTitle("DTP Disabled");
      }
      if (token !== expected) {
        return HtmlService.createHtmlOutput(
          "<pre>Unauthorized.\nDTP page requires a valid token.</pre>"
        ).setTitle("Unauthorized");
      }

      return HtmlService.createHtmlOutputFromFile("DTP")
        .setTitle("DTP – Final Paper")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    return HtmlService.createHtmlOutputFromFile("Index")
      .setTitle("Paper Generator (Faculty)")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      "<pre>doGet ERROR:\n" + (err && err.stack ? err.stack : err) + "</pre>"
    ).setTitle("doGet Error");
  }
}
function getDtpToken_(){
  return String(PropertiesService.getScriptProperties().getProperty("DTP_TOKEN") || "").trim();
}

/***********************
 * AVAILABLE SOURCE SHEETS
 ***********************/
function getAvailableSourceSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var required = ["Class","Subject","Chapter Name","Question File","Answer Key","Hints and Solutions","Metadata Sheet"];
  var out = [];
  for (var i=0;i<sheets.length;i++){
    var sh = sheets[i];
    var name = sh.getName();
    if (name === TOPICS_SHEET_NAME) continue;
    var lastCol = sh.getLastColumn();
    var lastRow = sh.getLastRow();
    if (lastCol < 1 || lastRow < 1) continue;
    var headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(x){ return String(x).trim(); });
    var ok = true;
    for (var r=0;r<required.length;r++){
      if (headers.indexOf(required[r]) === -1) { ok=false; break; }
    }
    if (ok) out.push(name);
  }
  out.sort();
  return out;
}

/***********************
 * TOPICS API
 ***********************/
function getTopicsForChapters(params){
  params = params || {};
  var className = String(params.className||"").trim();
  var subject = String(params.subject||"").trim();
  var chapters = Array.isArray(params.chapters) ? params.chapters : [];
  if (!className || !subject) throw new Error("getTopicsForChapters: className and subject are required.");

  var classSet = classSetFromReq_(className);
  var wantSubject = normStr_(subject);
  var chapterSet = {};
  for (var i=0;i<chapters.length;i++){
    var k = normStr_(chapters[i]);
    if (k) chapterSet[k]=true;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TOPICS_SHEET_NAME);
  if (!sh) throw new Error('Topics sheet not found: "'+TOPICS_SHEET_NAME+'"');

  var values = sh.getDataRange().getValues();
  if (values.length<2) return { chapters: [] };
  var headers = values[0].map(function(x){ return String(x).trim(); });
  var hl = headers.map(function(x){ return x.toLowerCase(); });
  function idxAny_(names){
    for (var j=0;j<names.length;j++){
      var idx = hl.indexOf(String(names[j]).toLowerCase());
      if (idx!==-1) return idx;
    }
    return -1;
  }

  var classIdx   = idxAny_(["Class"]);
  var subjectIdx = idxAny_(["Subject"]);
  var chapterIdx = idxAny_(["Chapter"]);
  var chapterIdIdx = idxAny_(["Chapter Id","ChapterId"]);
  var topicIdx   = idxAny_(["Topic"]);
  var topicIdIdx = idxAny_(["Topic Id","TopicId"]);

  if (classIdx===-1 || subjectIdx===-1 || chapterIdx===-1 || topicIdx===-1 || topicIdIdx===-1){
    throw new Error('Topics sheet must contain headers: Class, Subject, Chapter, Topic, Topic Id (and optionally Chapter Id).');
  }

  var out = {};
  for (var r=1;r<values.length;r++){
    var row = values[r];
    var rowClass = normClass_(row[classIdx]);
    if (!classSet[rowClass]) continue;
    if (normStr_(row[subjectIdx])!==wantSubject) continue;

    var chName = String(row[chapterIdx]||"").trim();
    if (!chName) continue;
    if (chapters.length && !chapterSet[normStr_(chName)]) continue;

    var chId = (chapterIdIdx!==-1) ? String(row[chapterIdIdx]||"").trim() : "";
    var tName = String(row[topicIdx]||"").trim();
    var tId = String(row[topicIdIdx]||"").trim();
    if (!tName || !tId) continue;

    if (!out[chName]) out[chName] = { chapterName: chName, chapterId: chId, topicsMap: {} };
    out[chName].topicsMap[tId] = tName;
    if (!out[chName].chapterId && chId) out[chName].chapterId = chId;
  }

  var chaptersOut=[];
  for (var ch in out) if (out.hasOwnProperty(ch)){
    var obj = out[ch];
    var topicsArr=[];
    for (var tid in obj.topicsMap) if (obj.topicsMap.hasOwnProperty(tid)){
      topicsArr.push({ id: tid, name: obj.topicsMap[tid] });
    }
    topicsArr.sort(function(a,b){ return a.name.localeCompare(b.name); });
    chaptersOut.push({ chapterName: obj.chapterName, chapterId: obj.chapterId, topics: topicsArr });
  }
  chaptersOut.sort(function(a,b){ return a.chapterName.localeCompare(b.chapterName); });
  return { chapters: chaptersOut };
}

/***********************
 * CHAPTER CONFIG / COUNTS / SYLLABUS
 ***********************/
function parseChaptersConfigFromReq_(req){
  var out=[];
  if (Array.isArray(req.chaptersConfig) && req.chaptersConfig.length){
    for (var i=0;i<req.chaptersConfig.length;i++){
      var c=req.chaptersConfig[i]||{};
      var name=String(c.chapterName||"").trim();
      if (!name) continue;
      var w=Number(c.weight);
      if (!isFinite(w)) w=0;
      w=Math.max(0, Math.min(100, w));
      var topicIds = Array.isArray(c.topicIds)
        ? c.topicIds.map(function(x){ return String(x||"").trim(); }).filter(Boolean)
        : [];
      out.push({ chapterName:name, weight:w, topicIds:topicIds });
    }
  }
  return out;
}
function assertWeightsSum100_(chaptersConfig){
  var sum = 0;
  for (var i=0;i<chaptersConfig.length;i++) sum += Number(chaptersConfig[i].weight||0) || 0;
  if (sum !== 100) throw new Error("Total chapter weightage must be exactly 100. Current total = " + sum + ".");
}
function allocateCountsByWeights_(names, weights, total){
  total=Math.max(0, Math.floor(Number(total||0)));
  var n=names.length;
  if(!n) return {};
  var ws=[], sumW=0;
  for (var i=0;i<n;i++){
    var w=Number(weights[i]);
    if(!isFinite(w) || w<0) w=0;
    ws[i]=w; sumW+=w;
  }
  if(sumW<=0){ sumW=n; for(i=0;i<n;i++) ws[i]=1; }
  var raw=[], base=[], frac=[], used=0;
  for(i=0;i<n;i++){
    raw[i]=total*(ws[i]/sumW);
    base[i]=Math.floor(raw[i]);
    frac[i]=raw[i]-base[i];
    used+=base[i];
  }
  var left=total-used;
  var order=[];
  for(i=0;i<n;i++) order.push({i:i, f:frac[i]});
  order.sort(function(a,b){ return b.f-a.f; });
  for(i=0;i<order.length && left>0;i++){ base[order[i].i]++; left--; }
  var out={};
  for(i=0;i<n;i++) out[names[i]]=base[i];
  return out;
}
function normalizeTypeCounts_(obj){
  var outMap={}, total=0;
  if (!obj || typeof obj!=="object") return {enabled:false,total:0,map:{}};
  for (var k in obj) if (obj.hasOwnProperty(k)){
    var raw=String(k||"").trim();
    var key=normStr_(raw);
    var n=Math.floor(Number(obj[k]||0));
    if (!key || !n || n<=0) continue;
    outMap[key] = (outMap[key]||0) + n;
    total += n;
  }
  return { enabled: total>0, total: total, map: outMap };
}
function formatSyllabusText_(chapters) {
  var parts = (chapters || []).map(function (x) { return String(x || "").trim(); }).filter(Boolean);
  var seen = {}, uniq = [];
  for (var i=0;i<parts.length;i++){
    var k = parts[i].toLowerCase();
    if (!seen[k]) { seen[k]=true; uniq.push(parts[i]); }
  }
  return uniq.join(", ");
}
function formatSyllabusFromChaptersConfig_(chaptersConfig){
  chaptersConfig = Array.isArray(chaptersConfig) ? chaptersConfig : [];
  var parts = [];
  var seen = {};
  for (var i=0;i<chaptersConfig.length;i++){
    var c = chaptersConfig[i] || {};
    var name = String(c.chapterName||"").trim();
    if (!name) continue;
    var topicIds = Array.isArray(c.topicIds) ? c.topicIds.filter(Boolean) : [];
    var mode = topicIds.length ? "P" : "C";
    var item = name + " (" + mode + ")";
    var key = item.toLowerCase();
    if (!seen[key]) { seen[key]=true; parts.push(item); }
  }
  return parts.join(", ");
}

/***********************
 * DOCS API / EXPORTS / PREVIEW ASSET HELPERS
 ***********************/
function docsApiReplacePaperName_(docId, paperName){
  var display = formatPaperNameForHeader_(paperName);
  var variants = ["{{PaperName}}","{{ PaperName }}","{{PaperName }}","{{ PaperName}}"];
  var requests = variants.map(function(v){
    return { replaceAllText: { containsText:{ text:v, matchCase:true }, replaceText:String(display) } };
  });
  var url="https://docs.googleapis.com/v1/documents/"+encodeURIComponent(docId)+":batchUpdate";
  UrlFetchApp.fetch(url,{
    method:"post",
    contentType:"application/json",
    headers:{ Authorization:"Bearer "+ScriptApp.getOAuthToken() },
    payload:JSON.stringify({ requests: requests })
  });
}
function formatPaperNameForHeader_(paperName){
  var s=collapseSpaces_(paperName);
  var maxChars=65;
  if (s.length<=maxChars) return s;
  return s.slice(0,maxChars-1).trimEnd()+"…";
}
function docsApiReplaceSyllabus_(docId, syllabusText) {
  var text = String(syllabusText || "").replace(/\s*\n\s*/g, " ").trim();
  var variants = [
    "{{Syllabus}}", "{{ Syllabus }}", "{{Syllabus }}", "{{ Syllabus}}",
    "{{$Syllabus}}", "{{ $Syllabus }}", "{{$Syllabus }}", "{{ $Syllabus}}"
  ];
  var requests = variants.map(function (v) {
    return {
      replaceAllText: {
        containsText: { text: v, matchCase: true },
        replaceText: text
      }
    };
  });
  var url = "https://docs.googleapis.com/v1/documents/" + encodeURIComponent(docId) + ":batchUpdate";
  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ requests: requests })
  });
}
function docsApiTightenPreviewDocLayout_(docId, isRichLayout){
  var requests = [{
    updateDocumentStyle: {
      documentStyle: {
        marginTop: { magnitude: isRichLayout ? 18 : 12, unit: "PT" },
        marginBottom: { magnitude: isRichLayout ? 18 : 12, unit: "PT" },
        marginLeft: { magnitude: isRichLayout ? 18 : 12, unit: "PT" },
        marginRight: { magnitude: isRichLayout ? 18 : 12, unit: "PT" },
        pageSize: {
          width: { magnitude: isRichLayout ? 520 : 430, unit: "PT" },
          height: { magnitude: isRichLayout ? 560 : 320, unit: "PT" }
        }
      },
      fields: "marginTop,marginBottom,marginLeft,marginRight,pageSize"
    }
  }];
  var url = "https://docs.googleapis.com/v1/documents/" + encodeURIComponent(docId) + ":batchUpdate";
  UrlFetchApp.fetch(url, {
    method:"post",
    contentType:"application/json",
    headers:{ Authorization:"Bearer " + ScriptApp.getOAuthToken() },
    payload:JSON.stringify({ requests: requests }),
    muteHttpExceptions:true
  });
}
function exportGoogleDocToPdfFile_(googleDocId, outName, folderId){
  var url = "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(googleDocId) +
    "/export?mimeType=" + encodeURIComponent("application/pdf") + "&alt=media";
  var resp = UrlFetchApp.fetch(url, {
    method:"get",
    headers:{ Authorization:"Bearer " + ScriptApp.getOAuthToken() }
  });
  var blob = resp.getBlob().setName(outName);
  return uploadBlobToDrive_(blob, outName, "application/pdf", folderId);
}
function exportGoogleDocToPdfBlob_(googleDocId, outName){
  var url = "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(googleDocId) +
    "/export?mimeType=" + encodeURIComponent("application/pdf") + "&alt=media";
  var resp = UrlFetchApp.fetch(url, {
    method:"get",
    headers:{ Authorization:"Bearer " + ScriptApp.getOAuthToken() }
  });
  return resp.getBlob().setName(outName || "preview.pdf");
}
function exportGoogleDocToDocxFile_(googleDocId, outName, folderId){
  var url = "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(googleDocId) +
    "/export?mimeType=" + encodeURIComponent("application/vnd.openxmlformats-officedocument.wordprocessingml.document") +
    "&alt=media";
  var resp = UrlFetchApp.fetch(url, {
    method:"get",
    headers:{ Authorization:"Bearer " + ScriptApp.getOAuthToken() }
  });
  var blob = resp.getBlob().setName(outName);
  return uploadBlobToDrive_(blob, outName, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", folderId);
}
function exportGoogleSlidesToPptxFile_(presentationId, outName, folderId){
  var url = "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(presentationId) +
    "/export?mimeType=" + encodeURIComponent("application/vnd.openxmlformats-officedocument.presentationml.presentation") +
    "&alt=media";
  var resp = UrlFetchApp.fetch(url, {
    method:"get",
    headers:{ Authorization:"Bearer " + ScriptApp.getOAuthToken() }
  });
  var blob = resp.getBlob().setName(outName);
  return uploadBlobToDrive_(blob, outName, "application/vnd.openxmlformats-officedocument.presentationml.presentation", folderId);
}
function getOrCreateSubFolder_(parentId, name){
  var parent = DriveApp.getFolderById(parentId);
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function getPreviewCacheFolderId_(){
  return getOrCreateSubFolder_(ACTIVE_OUTPUT_FOLDER_ID, PREVIEW_CACHE_FOLDER_NAME).getId();
}
function uploadBlobSimple_(blob, folderId){
  var file = DriveApp.getFolderById(folderId).createFile(blob);
  return { id:file.getId(), url:"https://drive.google.com/file/d/" + file.getId() + "/view" };
}
function getPreviewKindForItem_(item){
  var text = String((item && item.previewText) || "");
  var qType = normStr_(item && item.questionType || "");
  if (qType.indexOf("match") !== -1) return "rich";
  if (text.indexOf("[Table]") !== -1 || text.indexOf("[Image/Diagram]") !== -1) return "rich";
  return "text";
}
function getPreviewKindFromItemOnly_(item){
  var qType = normStr_(item && item.questionType || "");
  if (qType.indexOf("match") !== -1) return "rich";
  return "text";
}
function blockHasRenderableMedia_(block){
  if (!block || !block.length) return false;
  for (var i=0;i<block.length;i++){
    var el = block[i];
    var type = el.getType();
    if (type === DocumentApp.ElementType.TABLE ||
        type === DocumentApp.ElementType.INLINE_IMAGE ||
        type === DocumentApp.ElementType.POSITIONED_IMAGE){
      return true;
    }
    if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM){
      try {
        var n = el.getNumChildren ? el.getNumChildren() : 0;
        for (var j=0;j<n;j++){
          var childType = el.getChild(j).getType();
          if (childType === DocumentApp.ElementType.INLINE_IMAGE ||
              childType === DocumentApp.ElementType.POSITIONED_IMAGE){
            return true;
          }
        }
      } catch (e) {}
    }
  }
  return false;
}
function isUsablePreviewBase64_(b64){
  return !!b64 && String(b64).length > 2500;
}
function fetchDriveThumbnailBase64_(fileId){
  var meta = Drive.Files.get(fileId, { supportsAllDrives:true, fields:"thumbnailLink" });
  if (!meta.thumbnailLink) return "";
  var thumbUrl = String(meta.thumbnailLink).replace(/=s\d+/, "=s1600");
  var resp = UrlFetchApp.fetch(thumbUrl, {
    headers:{ Authorization:"Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions:true
  });
  if (resp.getResponseCode() >= 300) return "";
  return Utilities.base64Encode(resp.getBlob().getBytes());
}
function waitForDriveThumbnailBase64_(fileId, attempts, sleepMs){
  var tries = Math.max(1, Number(attempts || 6));
  var delay = Math.max(250, Number(sleepMs || 1200));
  for (var i=0; i<tries; i++){
    var base64 = "";
    try { base64 = fetchDriveThumbnailBase64_(fileId); } catch (e) {}
    if (base64) return base64;
    Utilities.sleep(delay + (i * 400));
  }
  return "";
}
function extractDriveIdFromUrl_(url){
  var s = String(url||"").trim();
  var m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  return "";
}

function createQuestionPreviewDoc_(item, finalQNo, folderId){
  var doc = DocumentApp.create("Preview Q" + finalQNo);
  var docId = doc.getId();
  moveDriveFileToFolder_(docId, folderId);

  var body = doc.getBody();
  body.clear();

  var blocks = extractQuestionBlocks_(openDocWithRetry_(item.sourceDocId, 10));
  var block = blocks[item.origQNo];
  docsApiTightenPreviewDocLayout_(docId, blockHasRenderableMedia_(block));
  if (!block || !block.length){
    body.appendParagraph("Q" + finalQNo + ". Preview unavailable.");
    doc.saveAndClose();
    return docId;
  }

  for (var i = 0; i < block.length; i++){
    appendPreviewElementSafe_(body, block[i], i === 0);
  }

  removeTrailingBlankPagesSafe_(doc);
  doc.saveAndClose();
  Utilities.sleep(1200);
  return docId;
}

function buildPreviewImageForItem_(item, finalQNo, folderId){
  var blocks = extractQuestionBlocks_(openDocWithRetry_(item.sourceDocId, 10));
  var block = blocks[item.origQNo];
  var previewKind = blockHasRenderableMedia_(block) ? "rich" : getPreviewKindFromItemOnly_(item);
  if (previewKind !== "rich"){
    return {
      previewDocId: "",
      previewPdfFileId: "",
      previewPdfUrl: "",
      previewImageBase64: "",
      previewImageDataUrl: "",
      previewImageWidth: 0,
      previewImageHeight: 0,
      previewImageMode: ""
    };
  }
  var previewDocId = createQuestionPreviewDoc_(item, finalQNo, folderId);

  // Try both the doc thumbnail and exported PDF thumbnail; keep the richer one.
  var docThumb = waitForDriveThumbnailBase64_(previewDocId, 5, 1100);
  var pdfFile = null;

  try{
    var pdfBlob = exportGoogleDocToPdfBlob_(previewDocId, "Q" + finalQNo + ".pdf");
    pdfFile = uploadBlobSimple_(pdfBlob, folderId);
  } catch (ePdf) {}

  var pdfThumb = "";
  if (pdfFile && pdfFile.id){
    pdfThumb = waitForDriveThumbnailBase64_(pdfFile.id, 7, 1300);
  }

  var imageBase64 = "";
  if (previewKind === "rich"){
    if (isUsablePreviewBase64_(pdfThumb) && pdfThumb.length >= docThumb.length){
      imageBase64 = pdfThumb;
    } else if (isUsablePreviewBase64_(docThumb)){
      imageBase64 = docThumb;
    } else if (isUsablePreviewBase64_(pdfThumb)){
      imageBase64 = pdfThumb;
    }
  }

  return {
    previewDocId: previewDocId,
    previewPdfFileId: pdfFile ? pdfFile.id : "",
    previewPdfUrl: pdfFile ? pdfFile.url : "",
    previewImageBase64: imageBase64,
    previewImageDataUrl: imageBase64 ? ("data:image/png;base64," + imageBase64) : "",
    previewImageWidth: 0,
    previewImageHeight: 0,
    previewImageMode: imageBase64 ? "local" : ""
  };
}

/***********************
 * PREVIEW / MANUAL SELECTION FLOW
 ***********************/
function previewPaper(req){
  var ctx = prepareGenerationContext_(req);
  var previewItems = hydratePreviewItems_(ctx.selectedItems, ctx.startNo);
  return {
    startNo: ctx.startNo,
    requestedTotalQuestions: ctx.requestedTotal,
    actualTotalQuestions: previewItems.length,
    syllabusText: ctx.syllabusText,
    notes: ctx.notes || [],
    items: previewItems,
    summary: summarizePreviewItems_(previewItems),
    output_folder_id: ctx.outputFolderId
  };
}

function searchQuestionsForPreview(payload){
  payload = payload || {};
  var baseRequest = payload.baseRequest || {};
  var filters = payload.filters || {};
  var limit = Math.max(1, Math.min(50, Math.floor(Number(filters.limit || 20))));
  var previewFolderId = getPreviewCacheFolderId_();
  var excludeKeysArr = Array.isArray(payload.excludeKeys) ? payload.excludeKeys : [];
  var excludeKeys = {};
  for (var i=0;i<excludeKeysArr.length;i++) excludeKeys[String(excludeKeysArr[i])] = true;

  var ctx = prepareGenerationContext_(baseRequest);
  var usablePool = ctx.usablePool || [];
  var wantedChapter = String(filters.chapter || "").trim();
  var wantedDiff = Number(filters.difficulty || 0) || 0;
  var wantedType = normStr_(filters.questionType || "");
  var wantedTopicId = String(filters.topicId || "").trim();
  var containsNeedle = normStr_(filters.containsText || "");
  var out = [];
  var blockCache = {};
  var metaCache = {};

  for (var j=0; j<usablePool.length && out.length < limit; j++){
    var item = usablePool[j];
    var key = itemKey_(item);
    if (excludeKeys[key]) continue;
    if (wantedChapter && item.chapter !== wantedChapter) continue;
    if (wantedDiff && Number(item.difficulty || 0) !== wantedDiff) continue;
    if (wantedType && normStr_(item.questionType || item.qTypeNorm || "") !== wantedType) continue;
    if (wantedTopicId && String(item.topicId || "") !== wantedTopicId) continue;
    var previewText = getPreviewTextForItem_(item, blockCache);
    if (containsNeedle && normStr_(previewText).indexOf(containsNeedle) === -1) continue;
    var metaObj = getPreviewMetaForItem_(item, metaCache);
    var previewAsset = buildPreviewImageForItem_(item, item.origQNo, previewFolderId);
    out.push(makePreviewDto_(item, "", metaObj, previewText, previewAsset));
  }

  return {
    items: out,
    notes: out.length ? [] : ["No matching questions found for the selected criteria."]
  };
}

function prepareGenerationContext_(req){
  var notes = [];
  if (!req || !req.paperName || !req.className || !req.subject){
    throw new Error("Missing required fields: paperName, className, subject.");
  }

  ACTIVE_OUTPUT_FOLDER_ID = getSubjectOutputFolderId_(req.subject);
  var sourceSheets = Array.isArray(req.sourceSheets) && req.sourceSheets.length
    ? req.sourceSheets.map(function(x){ return String(x||"").trim(); }).filter(Boolean)
    : [AITS_SHEET_NAME];

  var chaptersConfig = parseChaptersConfigFromReq_(req);
  if (!chaptersConfig.length) throw new Error("Select at least one chapter.");
  assertWeightsSum100_(chaptersConfig);

  var includeHints = (req.includeHints !== false);
  var includeVideoLinks = (req.includeVideoLinks !== false);
  var excludeProjects = Array.isArray(req.excludeProjects) ? req.excludeProjects : [];
  var excludeSet = buildLowerSet_(excludeProjects);

  var needE = Number(req.easy || 0);
  var needM = Number(req.medium || 0);
  var needH = Number(req.hard || 0);
  var requestedTotal = needE + needM + needH;
  var typeNorm = normalizeTypeCounts_(req.questionTypeCounts);
  if (typeNorm.enabled && typeNorm.total !== requestedTotal){
    notes.push("Type split total (" + typeNorm.total + ") != total (" + requestedTotal + "). Type split treated as best-effort preference.");
  }

  var startNo = getQuestionStartNo_(req.subject);
  var topicNameMap = buildTopicNameMap_(req.className, req.subject);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classSet = classSetFromReq_(req.className);
  var subjectNorm = normStr_(req.subject);
  var topicSetByChapter = {};
  var anyTopicSelected = false;
  for (var i=0;i<chaptersConfig.length;i++){
    var cfg = chaptersConfig[i];
    var set = {};
    (cfg.topicIds || []).forEach(function(t){
      if (t) set[String(t).trim()] = true;
    });
    topicSetByChapter[cfg.chapterName] = set;
    if (Object.keys(set).length) anyTopicSelected = true;
  }

  var pool = [];
  var metaRefs = {};
  var topicMissingNotes = [];

  for (var ci=0; ci<chaptersConfig.length; ci++){
    var chName = chaptersConfig[ci].chapterName;
    var selectedTopicIds = Object.keys(topicSetByChapter[chName] || {});
    var chapterHadAnySelectedTopicRows = false;

    for (var si=0; si<sourceSheets.length; si++){
      var sheetName = sourceSheets[si];
      var srcSheet = ss.getSheetByName(sheetName);
      if (!srcSheet){
        notes.push("Source sheet not found: " + sheetName);
        continue;
      }

      var data = srcSheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var headers = data[0].map(function(h){ return String(h).trim(); });
      var rows = findAllSourceRows_(data, headers, classSet, subjectNorm, chName);
      if (!rows.length) continue;

      for (var rr=0; rr<rows.length; rr++){
        var rowObj = rows[rr];
        var questionFileId = extractId_(rowObj["Question File"]);
        var answerKeyFileId = extractId_(rowObj["Answer Key"]);
        var hintsFileId = extractId_(rowObj["Hints and Solutions"]);
        var metaSheetId = extractId_(rowObj["Metadata Sheet"]);

        if (!questionFileId || !answerKeyFileId || !hintsFileId || !metaSheetId){
          notes.push("Chapter '" + chName + "' in " + sheetName + ": missing file IDs → skipped.");
          continue;
        }

        var sourceDocId = ensureGoogleDocId_(questionFileId);
        var answerKeyDocId = ensureGoogleDocId_(answerKeyFileId);
        var hintsDocId = ensureGoogleDocId_(hintsFileId);
        var metaSS = SpreadsheetApp.openById(metaSheetId);
        var metaSheet = findMetaSheet_(metaSS);
        var meta;
        try{
          meta = readMeta_(metaSheet, { requireTopic: anyTopicSelected, requireQuestionType: typeNorm.enabled });
        } catch(eMeta){
          notes.push("Metadata issue for chapter '" + chName + "' (" + sheetName + "): " + (eMeta.message || eMeta));
          continue;
        }

        if (anyTopicSelected && selectedTopicIds.length){
          for (var rmi=0; rmi<meta.rows.length; rmi++){
            var tId = String(meta.rows[rmi].TopicId||"").trim();
            if (tId && topicSetByChapter[chName][tId]) chapterHadAnySelectedTopicRows = true;
          }
        }

        var metaKey = metaSS.getId() + "::" + metaSheet.getSheetId();
        metaRefs[metaKey] = { sheet: metaSheet, usedColIndex: meta.usedColIndex, projectUsedColIndex: meta.projectUsedColIndex };
        for (var ri=0; ri<meta.rows.length; ri++){
          var r = meta.rows[ri];
          pool.push({
            chapter: chName,
            sourceSheet: sheetName,
            difficulty: Number(r.Difficulty||0) || 0,
            used: Number(r.Used||0) || 0,
            projectUsed: String(r.ProjectUsed||""),
            questionType: String(r.QuestionType||"").trim(),
            qTypeNorm: normStr_(String(r.QuestionType||"").trim()),
            topicId: String(r.TopicId||"").trim(),
            sourceDocId: sourceDocId,
            answerKeyDocId: answerKeyDocId,
            hintsDocId: hintsDocId,
            questionFileId: questionFileId,
            answerKeyFileId: answerKeyFileId,
            hintsFileId: hintsFileId,
            origQNo: r.__rowNum - 1,
            metaKey: metaKey,
            metaRowNum: r.__rowNum
          });
        }
      }
    }

    if (anyTopicSelected && selectedTopicIds.length && !chapterHadAnySelectedTopicRows) {
      var names = selectedTopicIds.map(function(id){
        return topicNameMap[chName+"::"+id] || ("TopicId:"+id);
      });
      topicMissingNotes.push("Chapter '" + chName + "': no questions found for selected topic(s): " + names.join(", ") + ". Used whole chapter instead.");
    }
  }

  if (!pool.length) throw new Error("No questions found across selected source sheets with given filters.");
  var selRes = selectFromPoolBestEffort_(pool, req, excludeSet, chaptersConfig, topicSetByChapter, typeNorm, topicMissingNotes);
  notes = notes.concat(selRes.notes || []);
  var usablePool = selRes.usablePool || [];
  var selectedItems = [];

  if (Array.isArray(req.manualSelectedItems) && req.manualSelectedItems.length){
    selectedItems = normalizeManualSelectedItems_(req.manualSelectedItems);
    if (!selectedItems.length) throw new Error("No valid questions were received from preview selection.");
    notes.push("Using preview-edited selection of " + selectedItems.length + " question(s).");
    if (req.shuffleOrder) notes.push("Preview order preserved. shuffleOrder ignored because manual preview order is being used.");
  } else {
    selectedItems = selRes.selectedItems || [];
    var sanRes = sanitizeSelectionSkipGarbledBestEffort_(selectedItems, usablePool, {
      requireSameType: typeNorm.enabled,
      requireSameChapter: true
    });
    selectedItems = sanRes.items;
    notes = notes.concat(sanRes.notes || []);
    if (req.shuffleOrder) selectedItems = shuffle_(selectedItems);
  }

  var syllabusText = formatSyllabusFromChaptersConfig_(chaptersConfig);
  return {
    notes: notes,
    selectedItems: selectedItems,
    usablePool: usablePool,
    metaRefs: metaRefs,
    startNo: startNo,
    includeHints: includeHints,
    includeVideoLinks: includeVideoLinks,
    chaptersConfig: chaptersConfig,
    requestedTotal: requestedTotal,
    sourceSheets: sourceSheets,
    syllabusText: syllabusText,
    outputFolderId: ACTIVE_OUTPUT_FOLDER_ID
  };
}

function normalizeManualSelectedItems_(items){
  var out = [];
  for (var i=0;i<(items || []).length;i++){
    var it = items[i] || {};
    var questionFileId = String(it.questionFileId || "").trim();
    var answerKeyFileId = String(it.answerKeyFileId || "").trim();
    var hintsFileId = String(it.hintsFileId || "").trim();
    var sourceDocId = String(it.sourceDocId || "").trim();
    var answerKeyDocId = String(it.answerKeyDocId || "").trim();
    var hintsDocId = String(it.hintsDocId || "").trim();
    if (!sourceDocId && questionFileId) sourceDocId = ensureGoogleDocId_(questionFileId);
    if (!answerKeyDocId && answerKeyFileId) answerKeyDocId = ensureGoogleDocId_(answerKeyFileId);
    if (!hintsDocId && hintsFileId) hintsDocId = ensureGoogleDocId_(hintsFileId);

    var norm = {
      chapter: String(it.chapter || "").trim(),
      sourceSheet: String(it.sourceSheet || "").trim(),
      difficulty: Number(it.difficulty || 0) || 0,
      used: Number(it.used || 0) || 0,
      projectUsed: String(it.projectUsed || ""),
      questionType: String(it.questionType || "").trim(),
      qTypeNorm: normStr_(String(it.questionType || it.qTypeNorm || "").trim()),
      topicId: String(it.topicId || "").trim(),
      sourceDocId: sourceDocId,
      answerKeyDocId: answerKeyDocId,
      hintsDocId: hintsDocId,
      questionFileId: questionFileId,
      answerKeyFileId: answerKeyFileId,
      hintsFileId: hintsFileId,
      origQNo: Number(it.origQNo || 0) || 0,
      metaKey: String(it.metaKey || "").trim(),
      metaRowNum: Number(it.metaRowNum || 0) || 0
    };

    if (!norm.sourceDocId || !norm.answerKeyDocId || !norm.hintsDocId) continue;
    if (!norm.metaKey || !norm.metaRowNum) continue;
    if (!norm.origQNo) continue;
    out.push(norm);
  }
  return out;
}

function hydratePreviewItems_(items, startNo){
  var out = [];
  var metaCache = {};
  var blockCache = {};
  var previewFolderId = getPreviewCacheFolderId_();
  for (var i=0;i<(items || []).length;i++){
    var item = items[i];
    var finalQNo = startNo + i;
    var metaObj = getPreviewMetaForItem_(item, metaCache);
    var previewText = getPreviewTextForItem_(item, blockCache);
    var previewAsset = buildPreviewImageForItem_(item, finalQNo, previewFolderId);
    out.push(makePreviewDto_(item, finalQNo, metaObj, previewText, previewAsset));
  }
  return out;
}

function makePreviewDto_(item, finalQNo, metaObj, previewText, previewAsset){
  metaObj = metaObj || {};
  previewAsset = previewAsset || {};
  return {
    itemKey: itemKey_(item),
    finalQNo: finalQNo,
    qid: String(metaObj.QID || ""),
    tagSubjectCode: String(metaObj.Tag_SubjectCode || ""),
    chapter: String(item.chapter || metaObj.Chapter || ""),
    topicName: String(metaObj.Topic || ""),
    subtopic: String(metaObj.Subtopic || ""),
    topicId: String(item.topicId || ""),
    difficulty: Number(item.difficulty || 0) || 0,
    difficultyLabel: difficultyLabel_(item.difficulty),
    questionType: String(item.questionType || metaObj["Question type"] || ""),
    used: Number(item.used || 0) || 0,
    projectUsed: String(item.projectUsed || ""),
    sourceSheet: String(item.sourceSheet || ""),
    origQNo: Number(item.origQNo || 0) || 0,
    selected: true,
    sourceDocId: item.sourceDocId,
    answerKeyDocId: item.answerKeyDocId,
    hintsDocId: item.hintsDocId,
    questionFileId: item.questionFileId,
    answerKeyFileId: item.answerKeyFileId,
    hintsFileId: item.hintsFileId,
    metaKey: item.metaKey,
    metaRowNum: item.metaRowNum,
    previewText: String(previewText || ""),
    previewImageDataUrl: String(previewAsset.previewImageDataUrl || ""),
    previewPdfUrl: String(previewAsset.previewPdfUrl || ""),
    previewImageWidth: Number(previewAsset.previewImageWidth || 0) || 0,
    previewImageHeight: Number(previewAsset.previewImageHeight || 0) || 0,
    previewImageMode: String(previewAsset.previewImageMode || ""),
    previewKind: getPreviewKindForItem_({
      questionType: item.questionType || metaObj["Question type"] || "",
      previewText: previewText
    })
  };
}

function getPreviewMetaForItem_(item, metaCache){
  metaCache = metaCache || {};
  if (!metaCache[item.metaKey]) metaCache[item.metaKey] = buildMetaIndex_(item);
  return getMetaRowObject_(metaCache[item.metaKey], item.metaRowNum);
}

function getPreviewTextForItem_(item, blockCache){
  blockCache = blockCache || {};
  if (!blockCache[item.sourceDocId]){
    blockCache[item.sourceDocId] = extractQuestionBlocks_(openDocWithRetry_(item.sourceDocId, 10));
  }
  var block = blockCache[item.sourceDocId][item.origQNo];
  return blockToPreviewText_(block, 1800);
}

function paragraphChildToPreviewChunk_(child){
  if (!child) return "";
  var type = child.getType();
  if (type === DocumentApp.ElementType.TEXT){
    return collapseSpaces_(child.asText().getText());
  }
  if (type === DocumentApp.ElementType.INLINE_IMAGE || type === DocumentApp.ElementType.POSITIONED_IMAGE){
    return "[Image/Diagram]";
  }
  try {
    if (child.getText) return collapseSpaces_(child.getText());
  } catch (e) {}
  return "";
}

function paragraphLikeToPreviewText_(el){
  if (!el) return "";
  var parts = [];
  try {
    var n = el.getNumChildren ? el.getNumChildren() : 0;
    for (var i=0;i<n;i++){
      var chunk = paragraphChildToPreviewChunk_(el.getChild(i));
      if (chunk) parts.push(chunk);
    }
  } catch (e) {}
  if (!parts.length){
    try { return collapseSpaces_(el.getText ? el.getText() : ""); } catch (e2) {}
  }
  return collapseSpaces_(parts.join(" "));
}

function blockToPreviewText_(block, maxChars){
  if (!block || !block.length) return "[Question preview unavailable]";
  var parts = [];
  for (var i=0;i<block.length;i++){
    var el = block[i];
    var type = el.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM){
      var txt = paragraphLikeToPreviewText_(el);
      if (txt) parts.push(txt);
      continue;
    }
    if (type === DocumentApp.ElementType.TABLE){
      parts.push(tableToPreviewText_(el.asTable()));
      continue;
    }
    if (type === DocumentApp.ElementType.INLINE_IMAGE || type === DocumentApp.ElementType.POSITIONED_IMAGE){
      parts.push("[Image/Diagram]");
      continue;
    }
  }
  var s = parts.join("\n");
  s = s.replace(/\n{3,}/g, "\n\n").trim();
  if (!s) s = "[Question preview unavailable]";
  if (maxChars && s.length > maxChars) s = s.slice(0, maxChars - 1).trimEnd() + "…";
  return s;
}

function tableToPreviewText_(table){
  try {
    var rows = [];
    var maxRows = Math.min(table.getNumRows(), 3);
    for (var r=0;r<maxRows;r++){
      var row = table.getRow(r);
      var cells = [];
      var maxCells = Math.min(row.getNumCells(), 4);
      for (var c=0;c<maxCells;c++){
        var txt = "";
        var cell = row.getCell(c);
        var paraParts = [];
        for (var p=0; p<cell.getNumChildren(); p++){
          var child = cell.getChild(p);
          var childType = child.getType();
          if (childType === DocumentApp.ElementType.PARAGRAPH || childType === DocumentApp.ElementType.LIST_ITEM){
            var childTxt = paragraphLikeToPreviewText_(child);
            if (childTxt) paraParts.push(childTxt);
          } else if (childType === DocumentApp.ElementType.INLINE_IMAGE || childType === DocumentApp.ElementType.POSITIONED_IMAGE){
            paraParts.push("[Image/Diagram]");
          }
        }
        txt = collapseSpaces_(paraParts.join(" "));
        if (txt) cells.push(txt);
      }
      if (cells.length) rows.push(cells.join(" | "));
    }
    return rows.length ? "[Table] " + rows.join(" || ") : "[Table]";
  } catch (e){
    return "[Table]";
  }
}

function difficultyLabel_(d){
  d = Number(d || 0) || 0;
  if (d === 1) return "Easy";
  if (d === 2) return "Medium";
  if (d === 3) return "Hard";
  return "Unknown";
}

function summarizePreviewItems_(items){
  var out = { total: 0, easy: 0, medium: 0, hard: 0, byType: {} };
  for (var i=0;i<(items || []).length;i++){
    var it = items[i];
    out.total++;
    var d = Number(it.difficulty || 0) || 0;
    if (d === 1) out.easy++;
    else if (d === 2) out.medium++;
    else if (d === 3) out.hard++;
    var t = String(it.questionType || "").trim() || "Unknown";
    out.byType[t] = (out.byType[t] || 0) + 1;
  }
  return out;
}

/***********************
 * TOPIC NAME MAP / SOURCE SHEET ROW FIND
 ***********************/
function buildTopicNameMap_(className, subject){
  var out = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TOPICS_SHEET_NAME);
  if (!sh) return out;
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return out;
  var headers = values[0].map(function(x){ return String(x).trim(); });
  var hl = headers.map(function(x){ return x.toLowerCase(); });
  function idxAny_(names){
    for (var j=0;j<names.length;j++){
      var idx = hl.indexOf(String(names[j]).toLowerCase());
      if (idx!==-1) return idx;
    }
    return -1;
  }
  var classIdx = idxAny_(["Class"]);
  var subjectIdx = idxAny_(["Subject"]);
  var chapterIdx = idxAny_(["Chapter"]);
  var topicIdx = idxAny_(["Topic"]);
  var topicIdIdx = idxAny_(["Topic Id","TopicId"]);
  if (classIdx===-1 || subjectIdx===-1 || chapterIdx===-1 || topicIdx===-1 || topicIdIdx===-1) return out;
  var classSet = classSetFromReq_(className);
  var subj = normStr_(subject);
  for (var r=1;r<values.length;r++){
    var row = values[r];
    var rc = normClass_(row[classIdx]);
    if (!classSet[rc]) continue;
    if (normStr_(row[subjectIdx]) !== subj) continue;
    var ch = String(row[chapterIdx]||"").trim();
    var tid = String(row[topicIdIdx]||"").trim();
    var tname = String(row[topicIdx]||"").trim();
    if (!ch || !tid || !tname) continue;
    out[ch+"::"+tid] = tname;
  }
  return out;
}

function findAllSourceRows_(data, headers, classSet, subjectNorm, chapter){
  var col={};
  for (var i=0;i<headers.length;i++) col[headers[i]]=i;
  var required=["Class","Subject","Chapter Name","Question File","Answer Key","Hints and Solutions","Metadata Sheet"];
  for (var r=0;r<required.length;r++) if (!(required[r] in col)) return [];
  var wantChapRaw=collapseSpaces_(chapter);
  var wantChapNorm=normChapterKey_(wantChapRaw);
  var wantSet=tokenSet_(wantChapRaw);
  var candidates=[];
  for (var rr=1; rr<data.length; rr++){
    var row=data[rr];
    var rc = normClass_(row[col["Class"]]);
    if (!classSet[rc]) continue;
    if (normStr_(row[col["Subject"]]) !== subjectNorm) continue;
    var chCell=String(row[col["Chapter Name"]]||"");
    var chClean=collapseSpaces_(chCell);
    if(!chClean) continue;
    candidates.push({ rr: rr, ch: chClean });
  }

  var out=[];
  for (i=0;i<candidates.length;i++){
    if (normStr_(candidates[i].ch) === normStr_(wantChapRaw)) out.push(rowToObj_(data[candidates[i].rr], headers));
  }
  if (out.length) return out;
  for (i=0;i<candidates.length;i++){
    if (normChapterKey_(candidates[i].ch) === wantChapNorm) out.push(rowToObj_(data[candidates[i].rr], headers));
  }
  if (out.length) return out;

  var scored=[];
  for (i=0;i<candidates.length;i++){
    var candSet=tokenSet_(candidates[i].ch);
    var score=jaccard_(wantSet,candSet);
    scored.push({ rr:candidates[i].rr, score:score });
  }
  scored.sort(function(a,b){ return b.score-a.score; });
  if (!scored.length || scored[0].score < 0.75) return [];
  var best = scored[0].score;
  for (i=0;i<scored.length;i++){
    if (scored[i].score >= best - 0.02) out.push(rowToObj_(data[scored[i].rr], headers));
    else break;
  }
  return out;
}

function rowToObj_(row, headers){
  var obj={};
  for (var h=0; h<headers.length; h++) obj[headers[h]] = row[h];
  return obj;
}

function extractId_(urlOrId){
  if (!urlOrId) return "";
  var s=String(urlOrId).trim();
  if (/^[a-zA-Z0-9-_]{20,}$/.test(s)) return s;
  var m=s.match(/\/d\/([a-zA-Z0-9-_]+)/); if (m) return m[1];
  m=s.match(/[?&]id=([a-zA-Z0-9-_]+)/); if (m) return m[1];
  return "";
}

/***********************
 * PROJECT DETECTION / FILTER / UPDATE
 ***********************/
function normalizeProjectMatchKey_(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, "");
}
function detectProjectsFromPaperName_(paperName) {
  var nameKey = normalizeProjectMatchKey_(paperName);
  var found = [];
  for (var i = 0; i < PROJECT_OPTIONS.length; i++) {
    var p = PROJECT_OPTIONS[i];
    var pKey = normalizeProjectMatchKey_(p);
    if (pKey && nameKey.indexOf(pKey) !== -1) found.push(p);
  }
  var seen = {}, out = [];
  for (var j = 0; j < found.length; j++) {
    var k = found[j].toLowerCase();
    if (!seen[k]) { seen[k] = true; out.push(found[j]); }
  }
  return out;
}
function getProjectsForThisPaper_(req){
  if (Array.isArray(req.projectsForThisPaper) && req.projectsForThisPaper.length){
    return req.projectsForThisPaper.map(function(x){ return String(x||"").trim(); }).filter(Boolean);
  }
  return detectProjectsFromPaperName_(req.paperName);
}
function buildLowerSet_(arr){
  var set={};
  for (var i=0;i<(arr||[]).length;i++){
    var k=normStr_(arr[i]);
    if (k) set[k]=true;
  }
  return set;
}
function splitCsvToLowerSet_(s){
  var out={};
  var parts=String(s||"").split(",");
  for (var i=0;i<parts.length;i++){
    var k=normStr_(parts[i]);
    if (k) out[k]=true;
  }
  return out;
}
function hasAnyExcludedProject_(projectUsedCell, excludeSet){
  if (!excludeSet) return false;
  var usedSet=splitCsvToLowerSet_(projectUsedCell);
  for (var k in excludeSet){
    if (excludeSet.hasOwnProperty(k) && usedSet[k]) return true;
  }
  return false;
}
function normalizeProjectsToCanonical_(arr){
  var canon={};
  for (var i=0;i<PROJECT_OPTIONS.length;i++) canon[normStr_(PROJECT_OPTIONS[i])] = PROJECT_OPTIONS[i];
  var out=[], seen={};
  for (var j=0;j<(arr||[]).length;j++){
    var raw=String(arr[j]||"").trim();
    if(!raw) continue;
    var key=normStr_(raw);
    if(seen[key]) continue;
    seen[key]=true;
    out.push(canon[key] || raw);
  }
  return out;
}
function incrementUsedAndProjectUsedMulti_(selectedItems, metaRefs, projectsForThisPaper){
  var projects = Array.isArray(projectsForThisPaper) ? projectsForThisPaper : [];
  var grouped={};
  for (var i=0;i<selectedItems.length;i++){
    var it=selectedItems[i];
    if (!it || !it.metaKey || !it.metaRowNum) continue;
    if (!grouped[it.metaKey]) grouped[it.metaKey]=[];
    grouped[it.metaKey].push(it);
  }
  for (var metaKey in grouped) if (grouped.hasOwnProperty(metaKey)){
    var ref=metaRefs[metaKey];
    if (!ref || !ref.sheet) continue;
    var sheet=ref.sheet;
    var usedCol=ref.usedColIndex+1;
    var projCol=(ref.projectUsedColIndex>=0) ? (ref.projectUsedColIndex+1) : -1;
    var items=grouped[metaKey];
    for (var j=0;j<items.length;j++){
      var rn=items[j].metaRowNum;
      var usedCell=sheet.getRange(rn, usedCol);
      var val=Number(usedCell.getValue()||0)||0;
      usedCell.setValue(val+1);
      if (projCol!==-1 && projects.length){
        var pc=sheet.getRange(rn, projCol);
        var existing=String(pc.getValue()||"");
        var set=splitCsvToLowerSet_(existing);
        for (var k=0;k<projects.length;k++){
          var p=String(projects[k]).trim();
          if(!p) continue;
          var lk=normStr_(p);
          if(!set[lk]) set[lk]=p;
        }
        var outArr=[];
        for (var key in set) if (set.hasOwnProperty(key)) outArr.push(set[key]===true ? key : set[key]);
        outArr = normalizeProjectsToCanonical_(outArr);
        pc.setValue(outArr.join(", "));
      }
    }
  }
}

/***********************
 * STYLE SYLLABUS
 ***********************/
function styleSyllabusInDoc_(doc, syllabusText) {
  var body = doc.getBody();
  var full = String(syllabusText || "").trim();
  if (!full) return;
  var L = full.length;
  var fontSize = 10;
  if (L > 140) fontSize = 9;
  if (L > 200) fontSize = 8;
  if (L > 260) fontSize = 7;
  var key = full.slice(0, Math.min(25, full.length));
  var tables = body.getTables();
  var maxTables = Math.min(3, tables.length);
  for (var ti = 0; ti < maxTables; ti++) {
    var t = tables[ti];
    for (var r = 0; r < t.getNumRows(); r++) {
      var row = t.getRow(r);
      for (var c = 0; c < row.getNumCells(); c++) {
        var cell = row.getCell(c);
        for (var k = 0; k < cell.getNumChildren(); k++) {
          var child = cell.getChild(k);
          if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
          var p = child.asParagraph();
          var txt = p.getText() || "";
          if (txt.indexOf(key) === -1) continue;
          p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
          try { p.setLineSpacing(0.5); }
          catch (e1) {
            try { p.setLineSpacing(0.8); }
            catch (e2) { p.setLineSpacing(1.0); }
          }
          p.setSpacingBefore(0);
          p.setSpacingAfter(0);
          try { p.editAsText().setFontSize(fontSize); } catch (e3) {}
          return;
        }
      }
    }
  }
}

/***********************
 * BEST-EFFORT SELECTION
 ***********************/
function selectFromPoolBestEffort_(pool, req, excludeSet, chaptersConfig, topicSetByChapter, typeNorm, topicMissingNotes){
  var notes=[];
  if (topicMissingNotes && topicMissingNotes.length) notes = notes.concat(topicMissingNotes);
  var needE=Number(req.easy||0), needM=Number(req.medium||0), needH=Number(req.hard||0);
  var totalNeed=needE+needM+needH;
  if (!totalNeed) return { selectedItems:[], usablePool:[], notes:["Requested total is 0."] };
  if (!pool || !pool.length) return { selectedItems:[], usablePool:[], notes:["No questions available."] };
  var uniqueOnly = !!req.uniqueOnly;
  var chapterNames = chaptersConfig.map(function(c){ return c.chapterName; });
  var weights = chaptersConfig.map(function(c){ return Number(c.weight)||0; });
  var usable = [];

  for (var ci=0; ci<chapterNames.length; ci++){
    var ch = chapterNames[ci];
    var tset = topicSetByChapter && topicSetByChapter[ch] ? topicSetByChapter[ch] : {};
    var hasTopicFilter = Object.keys(tset).length > 0;
    var chunk = [];
    for (var i=0;i<pool.length;i++){
      var x=pool[i];
      if (x.chapter !== ch) continue;
      var d=Number(x.difficulty||0);
      if (d!==1 && d!==2 && d!==3) continue;
      if (uniqueOnly && Number(x.used||0)!==0) continue;
      if (excludeSet && hasAnyExcludedProject_(x.projectUsed, excludeSet)) continue;
      if (hasTopicFilter){
        if (x.topicId && tset[x.topicId]) chunk.push(x);
      } else {
        chunk.push(x);
      }
    }
    if (hasTopicFilter && chunk.length === 0){
      notes.push("Chapter '"+ch+"': selected topic filter had 0 usable questions after filters → used whole chapter instead.");
      for (i=0;i<pool.length;i++){
        x=pool[i];
        if (x.chapter !== ch) continue;
        d=Number(x.difficulty||0);
        if (d!==1 && d!==2 && d!==3) continue;
        if (uniqueOnly && Number(x.used||0)!==0) continue;
        if (excludeSet && hasAnyExcludedProject_(x.projectUsed, excludeSet)) continue;
        chunk.push(x);
      }
    }
    usable = usable.concat(chunk);
  }

  if (!usable.length) return { selectedItems:[], usablePool:[], notes:notes.concat(["No questions remain after filters."]) };
  usable = shuffle_(usable);
  var byCh={};
  for (ci=0; ci<chapterNames.length; ci++) byCh[chapterNames[ci]]={1:[],2:[],3:[]};
  for (var ui=0;ui<usable.length;ui++){
    var q=usable[ui];
    if (byCh[q.chapter]) byCh[q.chapter][q.difficulty].push(q);
  }
  for (var cName in byCh) if (byCh.hasOwnProperty(cName)){
    byCh[cName][1]=shuffle_(byCh[cName][1]);
    byCh[cName][2]=shuffle_(byCh[cName][2]);
    byCh[cName][3]=shuffle_(byCh[cName][3]);
  }

  var needEByCh=allocateCountsByWeights_(chapterNames, weights, needE);
  var needMByCh=allocateCountsByWeights_(chapterNames, weights, needM);
  var needHByCh=allocateCountsByWeights_(chapterNames, weights, needH);
  var useType = !!(typeNorm && typeNorm.enabled && typeNorm.total>0);
  var typeRemain = {};
  if (useType){
    for (var tk in typeNorm.map) if (typeNorm.map.hasOwnProperty(tk)) typeRemain[tk] = typeNorm.map[tk];
  }

  var picked=[], pickedKeys={};
  function pickKey_(it){ return String(it.sourceDocId||"")+"::"+String(it.origQNo||""); }
  function pickAny_(list){
    while(list && list.length){
      var cand=list.pop();
      var key=pickKey_(cand);
      if (pickedKeys[key]) continue;
      pickedKeys[key]=true;
      return cand;
    }
    return null;
  }
  function pickPreferredType_(list){
    if (!useType) return pickAny_(list);
    for (var k=list.length-1;k>=0;k--){
      var cand=list[k];
      var t=cand.qTypeNorm||"";
      if (typeRemain[t] > 0){
        list.splice(k,1);
        var key=pickKey_(cand);
        if (pickedKeys[key]) continue;
        pickedKeys[key]=true;
        typeRemain[t]--;
        return cand;
      }
    }
    return pickAny_(list);
  }

  var gotByDiff={1:0,2:0,3:0};
  function track_(it, diff){
    picked.push(it);
    gotByDiff[diff]++;
  }

  for (ci=0; ci<chapterNames.length; ci++){
    var ch2 = chapterNames[ci];
    var needCh={1:needEByCh[ch2]||0,2:needMByCh[ch2]||0,3:needHByCh[ch2]||0};
    for (var d2=1; d2<=3; d2++){
      var needD=needCh[d2];
      while(needD>0){
        var q1=pickPreferredType_(byCh[ch2][d2]);
        if (!q1) break;
        track_(q1, d2);
        needD--;
      }
      if (needD>0) notes.push("Chapter '"+ch2+"' shortage: "+(d2===1?"Easy":d2===2?"Medium":"Hard")+" short by "+needD+".");
    }
  }

  function pullFromAnyChapterDiff_(diff,count){
    var lists=chapterNames.map(function(cn){ return byCh[cn]?byCh[cn][diff]:[]; });
    var pulled=0;
    while(count>0){
      var progressed=false;
      for (var li=0; li<lists.length && count>0; li++){
        var q2=pickPreferredType_(lists[li]);
        if(!q2) continue;
        track_(q2, diff);
        pulled++; count--; progressed=true;
      }
      if(!progressed) break;
    }
    return pulled;
  }

  var remE=Math.max(0,needE-gotByDiff[1]);
  var remM=Math.max(0,needM-gotByDiff[2]);
  var remH=Math.max(0,needH-gotByDiff[3]);
  if (remE>0){ var gE=pullFromAnyChapterDiff_(1,remE); if (gE<remE) notes.push("Overall Easy shortage: requested "+needE+", got "+gotByDiff[1]+"."); }
  if (remM>0){ var gM=pullFromAnyChapterDiff_(2,remM); if (gM<remM) notes.push("Overall Medium shortage: requested "+needM+", got "+gotByDiff[2]+"."); }
  if (remH>0){ var gH=pullFromAnyChapterDiff_(3,remH); if (gH<remH) notes.push("Overall Hard shortage: requested "+needH+", got "+gotByDiff[3]+"."); }

  if (picked.length < totalNeed){
    var needAny=totalNeed-picked.length;
    var filled=0;
    outer:
    while(needAny>0){
      var progressed2=false;
      for (ci=0; ci<chapterNames.length; ci++){
        var ch3 = chapterNames[ci];
        for (d2=1; d2<=3; d2++){
          if (needAny<=0) break outer;
          var q3=pickPreferredType_(byCh[ch3][d2]);
          if (!q3) continue;
          track_(q3, d2);
          needAny--; filled++; progressed2=true;
        }
      }
      if(!progressed2) break;
    }
    if (filled>0) notes.push("Filled "+filled+" question(s) outside difficulty split to reach total (availability constraint).");
    if (picked.length < totalNeed) notes.push("Total shortage: requested "+totalNeed+", got "+picked.length+".");
  }

  if (useType){
    for (var tk2 in typeRemain) if (typeRemain.hasOwnProperty(tk2) && typeRemain[tk2] > 0){
      notes.push("Type shortage: '"+tk2+"' short by "+typeRemain[tk2]+".");
    }
  }
  return { selectedItems:picked, usablePool:usable, notes:notes };
}

/***********************
 * GARBLED DETECTION + QUESTION BLOCK EXTRACTION
 ***********************/
function looksLikeLegacyHindiGarbage_(text){
  var t=String(text||"");
  if (!t.trim()) return false;
  var patterns=[/fn\[/i,/\bgSA\b/i,/\bdks\b/i,/\bvks\b/i,/\bLoFj\b/i,/\bHkkot\b/i,/;k/,/¼/,/½/,/¿/];
  var hits=0;
  for (var i=0;i<patterns.length;i++) if (patterns[i].test(t)) hits++;
  return hits>=2;
}
function previewFromBlocks_(blocks, qNo){
  var block=blocks[qNo];
  if (!block || !block.length) return "";
  for (var i=0;i<block.length;i++){
    var el=block[i];
    var type=el.getType();
    if (type===DocumentApp.ElementType.PARAGRAPH || type===DocumentApp.ElementType.LIST_ITEM){
      var txt=(el.getText ? el.getText() : "");
      txt=String(txt||"").trim();
      if (txt) return txt;
    }
  }
  return "";
}
function extractQuestionBlocks_(doc){
  var body=doc.getBody();
  var n=body.getNumChildren();
  var starts=[];
  for (var i=0;i<n;i++){
    var el=body.getChild(i);
    var text="";
    if (el.getType()===DocumentApp.ElementType.PARAGRAPH) text=el.asParagraph().getText().trim();
    else if (el.getType()===DocumentApp.ElementType.LIST_ITEM) text=el.asListItem().getText().trim();
    var m=text.match(/^Q(\d+)\./);
    if (m) starts.push({qNo:Number(m[1]), idx:i});
  }
  var blocks={};
  for (var s=0;s<starts.length;s++){
    var start=starts[s];
    var endIdx=(s+1<starts.length) ? (starts[s+1].idx-1) : (n-1);
    var elements=[];
    for (var j=start.idx;j<=endIdx;j++) elements.push(body.getChild(j));
    blocks[start.qNo]=elements;
  }
  return blocks;
}
function itemKey_(it){
  return String(it.sourceDocId||"")+"::"+String(it.origQNo||"");
}

function sanitizeSelectionSkipGarbledBestEffort_(selectedItems, pool, opts){
  opts=opts||{};
  var requireSameType=!!opts.requireSameType;
  var requireSameChapter=!!opts.requireSameChapter;
  var notes=[];
  var replacedCount=0;
  if (!selectedItems || !selectedItems.length){
    return { items:[], notes:notes, replacedCount:0 };
  }

  var poolByDiff={1:[],2:[],3:[]};
  var poolByDiffType={1:{},2:{},3:{}};
  for (var i=0;i<(pool||[]).length;i++){
    var it=pool[i];
    var d=Number(it.difficulty||0);
    if (!poolByDiff[d]) continue;
    poolByDiff[d].push(it);
    var tkey=it.qTypeNorm||"";
    if (!poolByDiffType[d][tkey]) poolByDiffType[d][tkey]=[];
    poolByDiffType[d][tkey].push(it);
  }
  poolByDiff[1]=shuffle_(poolByDiff[1]);
  poolByDiff[2]=shuffle_(poolByDiff[2]);
  poolByDiff[3]=shuffle_(poolByDiff[3]);
  function shuffleMap_(diff){
    var m=poolByDiffType[diff];
    for (var k in m) if (m.hasOwnProperty(k)) m[k]=shuffle_(m[k]);
  }
  shuffleMap_(1); shuffleMap_(2); shuffleMap_(3);

  var usedKeys={};
  for (i=0;i<selectedItems.length;i++) usedKeys[itemKey_(selectedItems[i])] = true;
  var blocksCache={}, previewCache={};
  function getPreview_(it2){
    var k=itemKey_(it2);
    if (previewCache.hasOwnProperty(k)) return previewCache[k];
    if (!blocksCache[it2.sourceDocId]) blocksCache[it2.sourceDocId]=extractQuestionBlocks_(openDocWithRetry_(it2.sourceDocId,10));
    var prev=previewFromBlocks_(blocksCache[it2.sourceDocId], it2.origQNo);
    previewCache[k]=prev;
    return prev;
  }
  function chapterOk_(cand, chapter){
    return !requireSameChapter || String(cand.chapter||"")===String(chapter||"");
  }

  var ptr={1:0,2:0,3:0};
  var ptrType={1:{},2:{},3:{}};
  function findReplacement_(diff, typeKey, chapter){
    if (requireSameType){
      var arrT=(poolByDiffType[diff] && poolByDiffType[diff][typeKey]) ? poolByDiffType[diff][typeKey] : [];
      if (ptrType[diff][typeKey]==null) ptrType[diff][typeKey]=0;
      while(ptrType[diff][typeKey] < arrT.length){
        var cand=arrT[ptrType[diff][typeKey]++];
        if (!chapterOk_(cand, chapter)) continue;
        var key=itemKey_(cand);
        if (usedKeys[key]) continue;
        var prev=getPreview_(cand);
        if (looksLikeLegacyHindiGarbage_(prev)) continue;
        usedKeys[key]=true;
        return cand;
      }
      return null;
    }
    var arr=poolByDiff[diff]||[];
    while(ptr[diff] < arr.length){
      var cand2=arr[ptr[diff]++];
      if (!chapterOk_(cand2, chapter)) continue;
      var key2=itemKey_(cand2);
      if (usedKeys[key2]) continue;
      var prev2=getPreview_(cand2);
      if (looksLikeLegacyHindiGarbage_(prev2)) continue;
      usedKeys[key2]=true;
      return cand2;
    }
    return null;
  }

  var out=[];
  for (i=0;i<selectedItems.length;i++){
    var cur=selectedItems[i];
    var prevCur=getPreview_(cur);
    if (!looksLikeLegacyHindiGarbage_(prevCur)){
      out.push(cur);
      continue;
    }
    var diff=Number(cur.difficulty||0);
    var tkeyCur=cur.qTypeNorm||"";
    var rep=findReplacement_(diff, tkeyCur, cur.chapter);
    if (rep){
      replacedCount++;
      out.push(rep);
      notes.push("Replaced a garbled question (chapter '"+(cur.chapter||"")+"', difficulty "+diff+").");
    } else {
      out.push(cur);
      notes.push("Could not replace a garbled question (chapter '"+(cur.chapter||"")+"', difficulty "+diff+").");
    }
  }
  return { items: out, notes: notes, replacedCount: replacedCount };
}

/***********************
 * IMAGE-SAFE EMPTY CHECKS / SAFE INSERT-APPEND
 ***********************/
function isTrulyEmptyParagraph_(p) {
  var txt = "";
  try { txt = (p.getText() || "").replace(/\u00A0/g, " ").trim(); } catch (e) {}
  if (txt) return false;
  try {
    for (var i = 0; i < p.getNumChildren(); i++) {
      var t = p.getChild(i).getType();
      if (t === DocumentApp.ElementType.INLINE_IMAGE) return false;
      if (t === DocumentApp.ElementType.POSITIONED_IMAGE) return false;
    }
  } catch (e2) {}
  return true;
}
function isTrulyEmptyListItem_(li) {
  var txt = "";
  try { txt = (li.getText() || "").replace(/\u00A0/g, " ").trim(); } catch (e) {}
  if (txt) return false;
  try {
    for (var i = 0; i < li.getNumChildren(); i++) {
      var t = li.getChild(i).getType();
      if (t === DocumentApp.ElementType.INLINE_IMAGE) return false;
      if (t === DocumentApp.ElementType.POSITIONED_IMAGE) return false;
    }
  } catch (e2) {}
  return true;
}
function insertInlineImageElement_(container, cursor, imgEl) {
  try {
    var blob = imgEl.getBlob();
    if (typeof container.insertImage === "function") {
      container.insertImage(cursor, blob);
      return cursor + 1;
    }
  } catch (e) {}
  container.insertParagraph(cursor, "[Image]");
  return cursor + 1;
}
function insertPositionedImageElement_(container, cursor, imgEl) {
  try {
    var blob = imgEl.getBlob();
    if (typeof container.insertImage === "function") {
      container.insertImage(cursor, blob);
      return cursor + 1;
    }
  } catch (e) {}
  container.insertParagraph(cursor, "[Image]");
  return cursor + 1;
}
function insertElementSafe_(container, cursor, el, outQNo, isFirstLine) {
  var type = el.getType();
  if (type === DocumentApp.ElementType.PARAGRAPH) {
    var srcP = el.asParagraph();
    if (isTrulyEmptyParagraph_(srcP)) {
      container.insertParagraph(cursor, SAFE_SPACE);
      return cursor + 1;
    }
    var inserted = container.insertParagraph(cursor, srcP.copy());
    if (isFirstLine) {
      try { inserted.editAsText().replaceText("^Q\\s*\\d+\\s*\\.", "Q" + outQNo + "."); } catch (e) {}
    }
    return cursor + 1;
  }
  if (type === DocumentApp.ElementType.LIST_ITEM) {
    var srcLi = el.asListItem();
    if (isTrulyEmptyListItem_(srcLi)) {
      container.insertParagraph(cursor, SAFE_SPACE);
      return cursor + 1;
    }
    if (typeof container.insertListItem === "function") container.insertListItem(cursor, srcLi.copy());
    else container.insertParagraph(cursor, srcLi.getText());
    return cursor + 1;
  }
  if (type === DocumentApp.ElementType.TABLE) {
    if (typeof container.insertTable === "function") container.insertTable(cursor, el.asTable().copy());
    else container.insertParagraph(cursor, "[Table omitted]");
    return cursor + 1;
  }
  if (type === DocumentApp.ElementType.INLINE_IMAGE) return insertInlineImageElement_(container, cursor, el.asInlineImage());
  if (type === DocumentApp.ElementType.POSITIONED_IMAGE) return insertPositionedImageElement_(container, cursor, el.asPositionedImage());
  container.insertParagraph(cursor, (el.getText ? String(el.getText()).trim() : "") || SAFE_SPACE);
  return cursor + 1;
}
function appendElementSafe_(body, el, outQNo, isFirstLine) {
  var type = el.getType();
  if (type === DocumentApp.ElementType.PARAGRAPH) {
    var srcP = el.asParagraph();
    if (isTrulyEmptyParagraph_(srcP)) { body.appendParagraph(SAFE_SPACE); return; }
    var p = body.appendParagraph(srcP.copy());
    if (isFirstLine) {
      try { p.editAsText().replaceText("^Q\\s*\\d+\\s*\\.", "Q" + outQNo + "."); } catch (e) {}
    }
    return;
  }
  if (type === DocumentApp.ElementType.LIST_ITEM) {
    var srcLi = el.asListItem();
    if (isTrulyEmptyListItem_(srcLi)) { body.appendParagraph(SAFE_SPACE); return; }
    body.appendListItem(srcLi.copy());
    return;
  }
  if (type === DocumentApp.ElementType.TABLE) { body.appendTable(el.asTable().copy()); return; }
  if (type === DocumentApp.ElementType.INLINE_IMAGE) {
    try { body.appendImage(el.asInlineImage().getBlob()); return; } catch (e) {}
    body.appendParagraph("[Image]");
    return;
  }
  if (type === DocumentApp.ElementType.POSITIONED_IMAGE) {
    try { body.appendImage(el.asPositionedImage().getBlob()); return; } catch (e) {}
    body.appendParagraph("[Image]");
    return;
  }
  body.appendParagraph((el.getText ? String(el.getText()).trim() : "") || SAFE_SPACE);
}
function appendPreviewElementSafe_(body, el, isFirstLine) {
  var type = el.getType();
  if (type === DocumentApp.ElementType.PARAGRAPH) {
    var srcP = el.asParagraph();
    if (isTrulyEmptyParagraph_(srcP)) { body.appendParagraph(SAFE_SPACE); return; }
    var p = body.appendParagraph(srcP.copy());
    if (isFirstLine) {
      try { p.editAsText().replaceText("^Q\\s*\\d+\\s*\\.\\s*", ""); } catch (e) {}
    }
    return;
  }
  if (type === DocumentApp.ElementType.LIST_ITEM) {
    var srcLi = el.asListItem();
    if (isTrulyEmptyListItem_(srcLi)) { body.appendParagraph(SAFE_SPACE); return; }
    var li = body.appendListItem(srcLi.copy());
    if (isFirstLine) {
      try { li.editAsText().replaceText("^Q\\s*\\d+\\s*\\.\\s*", ""); } catch (e2) {}
    }
    return;
  }
  if (type === DocumentApp.ElementType.TABLE) { body.appendTable(el.asTable().copy()); return; }
  if (type === DocumentApp.ElementType.INLINE_IMAGE) {
    try { body.appendImage(el.asInlineImage().getBlob()); return; } catch (e3) {}
    body.appendParagraph("[Image]");
    return;
  }
  if (type === DocumentApp.ElementType.POSITIONED_IMAGE) {
    try { body.appendImage(el.asPositionedImage().getBlob()); return; } catch (e4) {}
    body.appendParagraph("[Image]");
    return;
  }
  body.appendParagraph((el.getText ? String(el.getText()).trim() : "") || SAFE_SPACE);
}
function safeClearBodyPreserveOneParagraph_(doc) {
  var body = doc.getBody();
  body.clear();
  if (body.getNumChildren() === 0) { body.appendParagraph(SAFE_SPACE); return; }
  var first = body.getChild(0);
  if (first.getType() === DocumentApp.ElementType.PARAGRAPH) first.asParagraph().setText(SAFE_SPACE);
  else body.appendParagraph(SAFE_SPACE);
}
function removeTrailingBlankPagesSafe_(doc) {
  var body = doc.getBody();
  for (var i = body.getNumChildren() - 1; i >= 0; i--) {
    var el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) break;
    var p = el.asParagraph();
    if (!isTrulyEmptyParagraph_(p)) break;
    if (body.getNumChildren() <= 1) { p.setText(SAFE_SPACE); break; }
    if (i === body.getNumChildren() - 1) { p.setText(SAFE_SPACE); break; }
    body.removeChild(el);
  }
}

/***********************
 * ANSWER KEY / TAGGING / PLAN HELPERS
 ***********************/
function parseAnswerKeyDoc_(answerKeyDocId){
  var doc = openDocWithRetry_(answerKeyDocId, 10);
  var body = doc.getBody();
  var map = {};
  var re = /^\s*Q\s*0*([0-9]+)\s*[\.\)\: -]*\s*([A-Da-d]|[1-4])\s*$/;
  for (var i=0;i<body.getNumChildren();i++){
    var el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH && el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;
    var text = (el.getText ? el.getText() : "").trim();
    if (!text) continue;
    var m = text.match(re);
    if (!m) continue;
    map[Number(m[1])] = String(m[2]).toUpperCase();
  }
  return map;
}
function createTaggingSheetForPaper_(paperName, selectedItems, startNo, includeVideoLinks, folderId){
  var outSS = SpreadsheetApp.create(paperName + " - Tagging");
  var sh = outSS.getSheets()[0];
  sh.setName("Tagging");
  var headers = includeVideoLinks
    ? ["FinalQNo","QID","Tag_SubjectCode","Chapter","Topic","Subtopic","Difficulty","Question type","Video link"]
    : ["FinalQNo","QID","Tag_SubjectCode","Chapter","Topic","Subtopic","Difficulty","Question type"];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  var cache={}, rowsOut=[];
  for (var i=0;i<(selectedItems||[]).length;i++){
    var item=selectedItems[i];
    var finalQNo=startNo+i;
    if (!cache[item.metaKey]) cache[item.metaKey]=buildMetaIndex_(item);
    var metaObj=getMetaRowObject_(cache[item.metaKey], item.metaRowNum);
    var qType = (item.questionType && String(item.questionType).trim())
      ? String(item.questionType).trim()
      : (metaObj["Question type"] || "");
    var row=[
      finalQNo,
      metaObj.QID || "",
      metaObj.Tag_SubjectCode || "",
      metaObj.Chapter || "",
      metaObj.Topic || "",
      metaObj.Subtopic || "",
      metaObj.Difficulty || "",
      qType
    ];
    if (includeVideoLinks) row.push(metaObj["Video link"] || "");
    rowsOut.push(row);
  }
  if (rowsOut.length) sh.getRange(2,1,rowsOut.length,headers.length).setValues(rowsOut);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  moveDriveFileToFolder_(outSS.getId(), folderId || OUTPUT_FOLDER_ID);
  return outSS.getUrl();
}
function buildMetaIndex_(item){
  var parts=String(item.metaKey).split("::");
  var ssId=parts[0], sheetId=parts[1];
  var ss=SpreadsheetApp.openById(ssId);
  var sheets=ss.getSheets();
  var sheet=null;
  for (var i=0;i<sheets.length;i++){
    if (String(sheets[i].getSheetId())===String(sheetId)){ sheet=sheets[i]; break; }
  }
  if (!sheet) throw new Error("Tagging: metadata sheet not found for metaKey=" + item.metaKey);
  var lastCol=sheet.getLastColumn();
  var headerVals=sheet.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var map={};
  for (var c=0;c<headerVals.length;c++) map[headerVals[c]]=c+1;
  if (!map["Difficulty"] && map["difficuly"]) map["Difficulty"]=map["difficuly"];
  if (!map["Video link"] && map["Video Link"]) map["Video link"]=map["Video Link"];
  if (!map["Tag_SubjectCode"] && map["Tag_SubjectCod"]) map["Tag_SubjectCode"]=map["Tag_SubjectCod"];
  if (!map["Question type"] && map["Question Type"]) map["Question type"]=map["Question Type"];
  return { sheet: sheet, colMap: map };
}
function getMetaRowObject_(metaIndex, rowNum){
  var sh=metaIndex.sheet;
  var m=metaIndex.colMap;
  function read(colName){
    var col=m[colName];
    if(!col) return "";
    return sh.getRange(rowNum,col).getValue();
  }
  return {
    QID: read("QID"),
    Tag_SubjectCode: read("Tag_SubjectCode"),
    Chapter: read("Chapter"),
    Topic: read("Topic"),
    Subtopic: read("Subtopic"),
    Difficulty: read("Difficulty"),
    "Question type": read("Question type"),
    "Video link": read("Video link")
  };
}
function writePaperPlanJson_(req, selectedItems, startNo, includeHints, includeVideoLinks, excludeProjects, folderId, chaptersConfig, requestedTotal, sourceSheets){
  var plan = {
    version: 6,
    generatedAt: new Date().toISOString(),
    paperName: req.paperName,
    className: req.className,
    subject: req.subject,
    startNo: startNo,
    shuffleOrder: !!req.shuffleOrder,
    uniqueOnly: !!req.uniqueOnly,
    includeHints: includeHints,
    includeVideoLinks: includeVideoLinks,
    excludeProjects: excludeProjects || [],
    requestedTotalQuestions: Number(requestedTotal || 0),
    actualTotalQuestions: (selectedItems||[]).length,
    questionTypeCountsRequested: req.questionTypeCounts || {},
    chaptersConfig: chaptersConfig || [],
    sourceSheets: sourceSheets || [],
    outputFolderId: folderId || OUTPUT_FOLDER_ID,
    qpTemplateFileId: extractId_(TEMPLATE_DOC_ID),
    selectedQuestions: []
  };
  for (var i=0;i<(selectedItems||[]).length;i++){
    var it=selectedItems[i];
    plan.selectedQuestions.push({
      finalQNo: startNo+i,
      chapter: it.chapter||"",
      sourceSheet: it.sourceSheet || "",
      difficulty: Number(it.difficulty||0)||0,
      topicId: it.topicId||"",
      questionType: it.questionType||"",
      origQNo: Number(it.origQNo||0)||0,
      questionFileId: it.questionFileId||"",
      answerKeyFileId: it.answerKeyFileId||"",
      hintsFileId: it.hintsFileId||"",
      metaKey: it.metaKey||"",
      metaRowNum: it.metaRowNum||0
    });
  }
  var json=JSON.stringify(plan,null,2);
  var filename=req.paperName + " - paper_plan.json";
  var blob=Utilities.newBlob(json,"application/json",filename);
  var file=DriveApp.createFile(blob);
  moveDriveFileToFolder_(file.getId(), folderId || OUTPUT_FOLDER_ID);
  return { id:file.getId(), url:"https://drive.google.com/file/d/"+file.getId()+"/view" };
}

/***********************
 * METADATA DISCOVERY + READ
 ***********************/
function findMetaSheet_(metaSS){
  var sheets=metaSS.getSheets();
  for (var i=0;i<sheets.length;i++){
    var sh=sheets[i];
    var lastCol=sh.getLastColumn();
    var lastRow=sh.getLastRow();
    if (lastCol<1 || lastRow<1) continue;
    var headers=sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){
      return String(h).trim().toLowerCase();
    });
    var hasQid=headers.indexOf("qid")!==-1 || headers.indexOf("questionid")!==-1;
    var hasDiff=headers.indexOf("difficulty")!==-1 || headers.indexOf("difficuly")!==-1;
    if (hasQid && hasDiff) return sh;
  }
  throw new Error("Could not find metadata tab (needs QID + Difficulty/difficuly).");
}
function readMeta_(sheet, opts){
  opts=opts||{};
  var requireTopic=!!opts.requireTopic;
  var requireQuestionType=!!opts.requireQuestionType;
  var lastRow=sheet.getLastRow();
  var lastCol=sheet.getLastColumn();
  if (lastRow<2) throw new Error("Metadata sheet has no data rows.");
  var values=sheet.getRange(1,1,lastRow,lastCol).getValues();
  var headersLower=values[0].map(function(h){ return String(h).trim().toLowerCase(); });
  function idxAny_(names){
    for (var i=0;i<names.length;i++){
      var idx=headersLower.indexOf(names[i]);
      if (idx!==-1) return idx;
    }
    return -1;
  }
  var diffIdx=idxAny_(["difficulty","difficuly"]);
  if (diffIdx===-1) throw new Error("Metadata missing Difficulty column.");
  var topicIdx=idxAny_(["topic"]);
  if (requireTopic && topicIdx===-1) throw new Error('Metadata missing "Topic" column.');
  var qTypeIdx=idxAny_(["question type","questiontype"]);
  if (requireQuestionType && qTypeIdx===-1) throw new Error('Metadata missing "Question type" column.');
  var usedColIndex=headersLower.indexOf("used");
  if (usedColIndex===-1){
    sheet.getRange(1,lastCol+1).setValue("Used");
    sheet.getRange(2,lastCol+1,lastRow-1,1).setValue(0);
    lastCol = sheet.getLastColumn();
  }
  var projectUsedColIndex=headersLower.indexOf("project used");
  if (projectUsedColIndex===-1){
    sheet.getRange(1,lastCol+1).setValue("Project Used");
    sheet.getRange(2,lastCol+1,lastRow-1,1).setValue("");
    lastCol = sheet.getLastColumn();
  }
  values=sheet.getRange(1,1,lastRow,lastCol).getValues();
  headersLower=values[0].map(function(h){ return String(h).trim().toLowerCase(); });
  usedColIndex=headersLower.indexOf("used");
  projectUsedColIndex=headersLower.indexOf("project used");
  diffIdx=-1; topicIdx=-1; qTypeIdx=-1;
  for (var h=0;h<headersLower.length;h++){
    var nm=headersLower[h];
    if (diffIdx===-1 && (nm==="difficulty" || nm==="difficuly")) diffIdx=h;
    if (topicIdx===-1 && nm==="topic") topicIdx=h;
    if (qTypeIdx===-1 && (nm==="question type" || nm==="questiontype")) qTypeIdx=h;
  }
  var rows=[];
  for (var r=1;r<values.length;r++){
    var row=values[r];
    if (row.join("").trim()==="") continue;
    rows.push({
      __rowNum: r+1,
      Difficulty: Number(row[diffIdx]||0)||0,
      Used: Number(row[usedColIndex]||0)||0,
      ProjectUsed: String(row[projectUsedColIndex]||""),
      TopicId: (topicIdx!==-1) ? String(row[topicIdx]||"") : "",
      QuestionType: (qTypeIdx!==-1) ? String(row[qTypeIdx]||"") : ""
    });
  }
  return { rows: rows, usedColIndex: usedColIndex, projectUsedColIndex: projectUsedColIndex };
}

/***********************
 * DRIVE / DOC UTILITIES
 ***********************/
function resolveTemplateDocId_(templateIdOrLink){
  var templateFileId = extractId_(templateIdOrLink);
  if (!templateFileId) throw new Error("Invalid template id/link.");
  var meta = Drive.Files.get(templateFileId, { supportsAllDrives:true });
  if (meta.mimeType === "application/vnd.google-apps.document") return templateFileId;
  if (meta.mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"){
    if (FORCE_TEMPLATE_REFRESH_IF_DOCX){
      PropertiesService.getScriptProperties().deleteProperty("converted_google_doc_"+templateFileId);
    }
    return ensureGoogleDocId_(templateFileId);
  }
  throw new Error("Template must be Google Doc or DOCX. Found: " + meta.mimeType);
}
function resolvePptTemplatePresentationId_(templateIdOrLink){
  var templateFileId = extractId_(templateIdOrLink);
  if (!templateFileId) throw new Error("Invalid PPT template id/link.");
  var meta = Drive.Files.get(templateFileId, { supportsAllDrives:true });
  if (meta.mimeType === "application/vnd.google-apps.presentation") return templateFileId;
  if (meta.mimeType === "application/vnd.openxmlformats-officedocument.presentationml.presentation"){
    var props = PropertiesService.getScriptProperties();
    var cacheKey = "converted_google_slides_" + templateFileId;
    var cached = props.getProperty(cacheKey);
    if (cached){
      try {
        var cachedMeta = Drive.Files.get(cached, { supportsAllDrives:true });
        if (cachedMeta.mimeType === "application/vnd.google-apps.presentation") return cached;
      } catch (e) {}
    }
    var converted = driveCopyWithRetry_(
      { title:(meta.title || "PPT Template") + " (Converted)", mimeType:"application/vnd.google-apps.presentation", parents:[{id:OUTPUT_FOLDER_ID}] },
      templateFileId,
      { convert:true, supportsAllDrives:true }
    );
    props.setProperty(cacheKey, converted.id);
    return converted.id;
  }
  throw new Error("PPT template must be Google Slides or PPTX. Found: " + meta.mimeType);
}
function openDocWithRetry_(docId, attempts){
  var tries=attempts||8, lastErr=null;
  for (var i=0;i<tries;i++){
    try{ return DocumentApp.openById(docId); }
    catch(e){ lastErr=e; Utilities.sleep(800 + i*450); }
  }
  throw lastErr || new Error("Failed to open document after retries.");
}
function ensureGoogleDocId_(fileId){
  var props=PropertiesService.getScriptProperties();
  var cacheKey="converted_google_doc_"+fileId;
  var cached=props.getProperty(cacheKey);
  if (cached){
    try{
      var f2=Drive.Files.get(cached,{supportsAllDrives:true});
      if (f2.mimeType==="application/vnd.google-apps.document"){
        openDocWithRetry_(cached,3);
        return cached;
      }
    } catch(e){}
  }
  var f=Drive.Files.get(fileId,{supportsAllDrives:true});
  if (f.mimeType==="application/vnd.google-apps.document"){
    openDocWithRetry_(fileId,3);
    return fileId;
  }
  if (f.mimeType!=="application/vnd.openxmlformats-officedocument.wordprocessingml.document"){
    throw new Error("File must be Google Doc or DOCX. Found mimeType="+f.mimeType+" fileId="+fileId);
  }
  var converted = driveCopyWithRetry_(
    { title:(f.title||"Converted")+" (Converted)", mimeType:"application/vnd.google-apps.document", parents:[{id:OUTPUT_FOLDER_ID}] },
    fileId,
    { convert:true, supportsAllDrives:true }
  );
  Utilities.sleep(900);
  openDocWithRetry_(converted.id,8);
  props.setProperty(cacheKey, converted.id);
  return converted.id;
}
function uploadBlobToDrive_(blob, name, mimeType, parentFolderId){
  var boundary="-------314159265358979323846";
  var delimiter="\r\n--"+boundary+"\r\n";
  var closeDelim="\r\n--"+boundary+"--";
  var metadata = { name:name, mimeType:mimeType, parents:[parentFolderId] };
  var multipartBody =
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) +
    delimiter +
    "Content-Type: " + mimeType + "\r\n" +
    "Content-Transfer-Encoding: base64\r\n\r\n" +
    Utilities.base64Encode(blob.getBytes()) +
    closeDelim;
  var resp = UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true",
    {
      method:"post",
      contentType:"multipart/related; boundary="+boundary,
      payload:multipartBody,
      headers:{ Authorization:"Bearer " + ScriptApp.getOAuthToken() }
    }
  );
  return JSON.parse(resp.getContentText());
}
function copyTemplateIntoFolder_(templateId, newTitle, folderId){
  try{
    var folder=DriveApp.getFolderById(folderId);
    var file=DriveApp.getFileById(templateId).makeCopy(newTitle, folder);
    return { id:file.getId() };
  } catch(e){
    return driveCopyWithRetry_(
      { title:newTitle, parents:[{id:folderId}] },
      templateId,
      { supportsAllDrives:true }
    );
  }
}
function copyPptTemplateIntoFolder_(templateId, newTitle, folderId){
  try{
    var folder = DriveApp.getFolderById(folderId);
    var file = DriveApp.getFileById(templateId).makeCopy(newTitle, folder);
    return { id:file.getId() };
  } catch (e){
    return driveCopyWithRetry_(
      { title:newTitle, parents:[{id:folderId}] },
      templateId,
      { supportsAllDrives:true }
    );
  }
}
function moveDriveFileToFolder_(fileId, folderId){
  var file=Drive.Files.get(fileId,{supportsAllDrives:true});
  var prevParents=(file.parents||[]).map(function(p){ return p.id; }).join(",");
  Drive.Files.update({}, fileId, null, { addParents: folderId, removeParents: prevParents, supportsAllDrives:true });
}
function isRateLimitError_(e){
  var msg=String((e&&e.message)?e.message:e);
  return msg.indexOf("User rate limit exceeded")!==-1 ||
         msg.indexOf("Rate Limit Exceeded")!==-1 ||
         msg.indexOf("Quota exceeded")!==-1 ||
         msg.indexOf("RESOURCE_EXHAUSTED")!==-1 ||
         msg.indexOf("429")!==-1;
}
function withBackoff_(fn){
  var delayMs=800, maxDelay=30000, maxRetries=8;
  for (var i=0;i<maxRetries;i++){
    try{ return fn(); }
    catch(e){
      if(!isRateLimitError_(e)) throw e;
      Utilities.sleep(delayMs + Math.floor(Math.random()*250));
      delayMs=Math.min(delayMs*2, maxDelay);
    }
  }
  return fn();
}
function driveCopyWithRetry_(resource, fileId, params){
  return withBackoff_(function(){ return Drive.Files.copy(resource, fileId, params||{}); });
}
function shuffle_(arr){
  var a=arr.slice();
  for (var i=a.length-1;i>0;i--){
    var j=Math.floor(Math.random()*(i+1));
    var tmp=a[i]; a[i]=a[j]; a[j]=tmp;
  }
  return a;
}

/***********************
 * DOC BUILDERS / EXPORT PACK
 ***********************/
function buildQuestionPaperDoc_(opts){
  var folderId = opts.folderId || OUTPUT_FOLDER_ID;
  var copied = copyTemplateIntoFolder_(opts.templateId, qpFileName_(opts.paperName), folderId);
  var docId = copied.id;
  docsApiReplacePaperName_(docId, opts.paperName);
  docsApiReplaceSyllabus_(docId, opts.syllabusText || "");
  Utilities.sleep(700);
  var doc = openDocWithRetry_(docId, 10);
  styleSyllabusInDoc_(doc, opts.syllabusText || "");
  var body = doc.getBody();
  var idx = 0;
  while (idx < body.getNumChildren() && body.getChild(idx).getType() === DocumentApp.ElementType.TABLE) idx++;
  if (idx < body.getNumChildren() && body.getChild(idx).getType() === DocumentApp.ElementType.PARAGRAPH) {
    var t = (body.getChild(idx).asParagraph().getText() || "").trim();
    if (t === "" || t === SAFE_SPACE) idx++;
  }
  var cursor = idx;
  var blocksCache = {};
  var items = opts.selectedItems || [];
  for (var i=0;i<items.length;i++){
    var item = items[i];
    if (!blocksCache[item.sourceDocId]){
      blocksCache[item.sourceDocId] = extractQuestionBlocks_(openDocWithRetry_(item.sourceDocId, 10));
    }
    var block = blocksCache[item.sourceDocId][item.origQNo];
    if (!block) continue;
    var outQNo = opts.startNo + i;
    for (var b=0;b<block.length;b++){
      cursor = insertElementSafe_(body, cursor, block[b], outQNo, b===0);
    }
    if (i !== items.length-1){
      body.insertParagraph(cursor, SAFE_SPACE);
      cursor++;
    }
  }
  removeTrailingBlankPagesSafe_(doc);
  doc.saveAndClose();
  return docId;
}

function buildAnswerHintsDoc_(opts){
  var folderId = opts.folderId || OUTPUT_FOLDER_ID;
  var copied = copyTemplateIntoFolder_(opts.baseHintsTemplateId, akhsFileName_(opts.paperName), folderId);
  var docId = copied.id;
  Utilities.sleep(700);
  docsApiReplacePaperName_(docId, opts.paperName);
  Utilities.sleep(250);
  var doc = openDocWithRetry_(docId, 10);
  safeClearBodyPreserveOneParagraph_(doc);
  var body = doc.getBody();
  var title = body.appendParagraph("Answer Key");
  try { title.setBold(true); title.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e){}
  body.appendParagraph(SAFE_SPACE);
  var answerCache = {};
  function getAnswerMap_(docId2){
    if (!answerCache[docId2]) answerCache[docId2] = parseAnswerKeyDoc_(docId2);
    return answerCache[docId2];
  }
  for (var i=0;i<opts.selectedItems.length;i++){
    var item=opts.selectedItems[i];
    var outQNo=opts.startNo+i;
    var amap=getAnswerMap_(item.answerKeyDocId);
    var ans=(amap[item.origQNo]!=null) ? String(amap[item.origQNo]) : "N/A";
    body.appendParagraph("Q"+outQNo+". "+ans);
  }
  if (opts.includeHints === false){
    removeTrailingBlankPagesSafe_(doc);
    doc.saveAndClose();
    return docId;
  }
  body.appendPageBreak();
  var ht=body.appendParagraph("Hints and Solution");
  try { ht.setBold(true); ht.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e){}
  body.appendParagraph(SAFE_SPACE);
  var hintsBlocksCache = {};
  for (var j=0;j<opts.selectedItems.length;j++){
    var it=opts.selectedItems[j];
    var qno=opts.startNo+j;
    if (!hintsBlocksCache[it.hintsDocId]){
      hintsBlocksCache[it.hintsDocId]=extractQuestionBlocks_(openDocWithRetry_(it.hintsDocId, 10));
    }
    var blk=hintsBlocksCache[it.hintsDocId][it.origQNo];
    if(!blk || !blk.length){
      body.appendParagraph("Q"+qno+". N/A");
      if (j !== opts.selectedItems.length-1) body.appendParagraph(SAFE_SPACE);
      continue;
    }
    for (var k=0;k<blk.length;k++){
      var el=blk[k];
      try{
        appendElementSafe_(body, el, qno, k===0);
      } catch(e2){
        body.appendParagraph("Q"+qno+". N/A");
        break;
      }
    }
    if (j !== opts.selectedItems.length-1) body.appendParagraph(SAFE_SPACE);
  }
  removeTrailingBlankPagesSafe_(doc);
  doc.saveAndClose();
  return docId;
}

function createEmptyAnswerDoc_(paperName, folderId){
  var doc = DocumentApp.create(akhsFileName_(paperName));
  var id = doc.getId();
  moveDriveFileToFolder_(id, folderId || OUTPUT_FOLDER_ID);
  var body = doc.getBody();
  body.clear();
  body.appendParagraph("Answer Key").setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(SAFE_SPACE);
  body.appendParagraph("No questions were generated based on available data.");
  doc.saveAndClose();
  return id;
}

function shouldUseRenderedPreviewForItem_(item){
  var text = String((item && item.previewText) || "");
  if (!item || !item.previewImageDataUrl) return false;
  if (String(item.previewKind || "") === "rich") return true;
  return text.indexOf("[Image/Diagram]") !== -1 || text.indexOf("[Table]") !== -1;
}

function buildRenderedPreviewImageDataUrlForExport_(item, folderId){
  if (item && item.previewImageDataUrl) return String(item.previewImageDataUrl);
  var finalQNo = Number(item && item.finalQNo || 0) || Number(item && item.origQNo || 0) || 1;
  var previewDocId = createQuestionPreviewDoc_(item, finalQNo, folderId);
  var docThumb = waitForDriveThumbnailBase64_(previewDocId, 5, 1100);
  var pdfFile = null;
  try {
    var pdfBlob = exportGoogleDocToPdfBlob_(previewDocId, "Q" + finalQNo + ".pdf");
    pdfFile = uploadBlobSimple_(pdfBlob, folderId);
  } catch (ePdf) {}
  var pdfThumb = "";
  if (pdfFile && pdfFile.id){
    pdfThumb = waitForDriveThumbnailBase64_(pdfFile.id, 7, 1300);
  }
  var imageBase64 = "";
  if (isUsablePreviewBase64_(pdfThumb) && pdfThumb.length >= docThumb.length){
    imageBase64 = pdfThumb;
  } else if (isUsablePreviewBase64_(docThumb)){
    imageBase64 = docThumb;
  } else if (isUsablePreviewBase64_(pdfThumb)){
    imageBase64 = pdfThumb;
  }
  return imageBase64 ? ("data:image/png;base64," + imageBase64) : "";
}

function buildPreviewExportAssets_(previewItems, folderId){
  var assets = [];
  for (var i=0;i<(previewItems||[]).length;i++){
    var item = previewItems[i];
    var dataUrl = buildRenderedPreviewImageDataUrlForExport_(item, folderId);
    if (!dataUrl) continue;
    var raw = String(dataUrl).replace(/^data:image\/png;base64,/, "");
    var fileName = "QUES_ENG_Q" + item.finalQNo + ".png";
    var blob = Utilities.newBlob(Utilities.base64Decode(raw), "image/png", fileName);
    assets.push({
      finalQNo: item.finalQNo,
      chapter: item.chapter || "",
      blob: blob,
      zipPath: "final_output/images/" + fileName
    });
  }
  return assets;
}

function getPptAnchorBox_(slide){
  var elements = slide.getPageElements();
  for (var i=0;i<elements.length;i++){
    var el = elements[i];
    try {
      if (String(el.getTitle && el.getTitle() || "").trim() === PPT_ANCHOR_SHAPE_NAME){
        return {
          left: el.getLeft(),
          top: el.getTop(),
          width: el.getWidth(),
          height: el.getHeight(),
          element: el
        };
      }
    } catch (e) {}
  }
  return null;
}

function computePptSizeForAreaFraction_(slideW, slideH, left, top, aspect, targetFraction){
  var slideArea = slideW * slideH;
  var desiredArea = slideArea * targetFraction;
  var w = Math.sqrt(desiredArea * aspect);
  var h = Math.sqrt(desiredArea / aspect);
  var maxW = Math.max(1, slideW - left);
  var maxH = Math.max(1, slideH - top);
  if (w > maxW || h > maxH){
    var scale = Math.min(maxW / w, maxH / h);
    w *= scale;
    h *= scale;
  }
  return { width:w, height:h };
}

function placeExportImageOnSlide_(slide, blob, pres){
  var slideW = pres.getPageWidth();
  var slideH = pres.getPageHeight();
  var anchor = getPptAnchorBox_(slide);
  var left = anchor ? anchor.left : 20;
  var top = anchor ? anchor.top : 20;
  var img = slide.insertImage(blob);
  var aspect = Math.max(0.01, img.getWidth() / Math.max(1, img.getHeight()));
  var size = computePptSizeForAreaFraction_(slideW, slideH, left, top, aspect, PPT_TARGET_FRACTION);
  img.setLeft(left);
  img.setTop(top);
  img.setWidth(size.width);
  img.setHeight(size.height);
}

function createQuestionPptForPaper_(paperName, previewItems, folderId){
  var exportAssets = buildPreviewExportAssets_(previewItems, folderId || ACTIVE_OUTPUT_FOLDER_ID);
  var presId = "";
  var pres = null;
  var useTemplate = !!String(PPT_TEMPLATE_FILE_ID || "").trim();
  if (useTemplate){
    var templateId = resolvePptTemplatePresentationId_(PPT_TEMPLATE_FILE_ID);
    var copied = copyPptTemplateIntoFolder_(templateId, pptFileName_(paperName), folderId || ACTIVE_OUTPUT_FOLDER_ID);
    presId = copied.id;
    pres = SlidesApp.openById(presId);
  } else {
    pres = SlidesApp.create(pptFileName_(paperName));
    presId = pres.getId();
    moveDriveFileToFolder_(presId, folderId || ACTIVE_OUTPUT_FOLDER_ID);
  }

  var slides = pres.getSlides();
  if (!slides.length){
    throw new Error("PPT template has no slides.");
  }
  while (slides.length > 1){
    try { slides[slides.length - 1].remove(); } catch (eRm) { break; }
    slides = pres.getSlides();
  }

  var baseSlide = pres.getSlides()[0];
  if (!exportAssets.length){
    try { baseSlide.insertTextBox("No preview questions available.", 20, 20, 400, 40); } catch (eTxt) {}
  }

  for (var i=0;i<exportAssets.length;i++){
    var item = exportAssets[i];
    var slide = baseSlide.duplicate();
    placeExportImageOnSlide_(slide, item.blob, pres);
  }
  if (exportAssets.length){
    try { baseSlide.remove(); } catch (eBase) {}
  }
  pres.saveAndClose();
  var pptx = exportGoogleSlidesToPptxFile_(presId, pptFileName_(paperName) + ".pptx", folderId || ACTIVE_OUTPUT_FOLDER_ID);
  return {
    presentationId: presId,
    pptxFileId: pptx.id,
    pptxUrl: "https://drive.google.com/file/d/" + pptx.id + "/view",
    exportAssets: exportAssets
  };
}

function createZipBundleForPaper_(paperName, fileIds, folderId, extraBlobs){
  var blobs = [];
  for (var i=0;i<(fileIds||[]).length;i++){
    var id = String(fileIds[i]||"").trim();
    if (!id) continue;
    blobs.push(DriveApp.getFileById(id).getBlob());
  }
  for (var j=0;j<(extraBlobs||[]).length;j++){
    if (extraBlobs[j]) blobs.push(extraBlobs[j]);
  }
  if (!blobs.length) throw new Error("ZIP: no files to bundle.");
  var zipBlob = Utilities.zip(blobs, zipFileName_(paperName));
  var zipFile = DriveApp.getFolderById(folderId || ACTIVE_OUTPUT_FOLDER_ID).createFile(zipBlob);
  return { id:zipFile.getId(), url:"https://drive.google.com/file/d/" + zipFile.getId() + "/view" };
}

/***********************
 * MAIN GENERATOR
 ***********************/
function generatePaper(req){
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    var ctx = prepareGenerationContext_(req);
    var notes = (ctx.notes || []).slice();
    var selectedItems = ctx.selectedItems || [];
    var startNo = ctx.startNo;
    var includeHints = ctx.includeHints;
    var includeVideoLinks = ctx.includeVideoLinks;
    var chaptersConfig = ctx.chaptersConfig;
    var requestedTotal = ctx.requestedTotal;
    var sourceSheets = ctx.sourceSheets;
    var syllabusText = ctx.syllabusText;
    var metaRefs = ctx.metaRefs || {};

    var qpTemplateId = resolveTemplateDocId_(TEMPLATE_DOC_ID);
    var qpDocId = buildQuestionPaperDoc_({
      templateId: qpTemplateId,
      paperName: req.paperName,
      syllabusText: syllabusText,
      selectedItems: selectedItems,
      startNo: startNo,
      folderId: ACTIVE_OUTPUT_FOLDER_ID
    });

    var akhsDocId;
    if (selectedItems.length){
      akhsDocId = buildAnswerHintsDoc_({
        baseHintsTemplateId: selectedItems[0].hintsDocId,
        paperName: req.paperName,
        selectedItems: selectedItems,
        startNo: startNo,
        includeHints: includeHints,
        folderId: ACTIVE_OUTPUT_FOLDER_ID
      });
    } else {
      akhsDocId = createEmptyAnswerDoc_(req.paperName, ACTIVE_OUTPUT_FOLDER_ID);
      notes.push("No questions selected → answer doc contains placeholder.");
    }

    var qpPdf = exportGoogleDocToPdfFile_(qpDocId, qpFileName_(req.paperName) + ".pdf", ACTIVE_OUTPUT_FOLDER_ID);
    var akhsPdf = exportGoogleDocToPdfFile_(akhsDocId, akhsFileName_(req.paperName) + ".pdf", ACTIVE_OUTPUT_FOLDER_ID);
    var qpDocx = exportGoogleDocToDocxFile_(qpDocId, qpFileName_(req.paperName) + ".docx", ACTIVE_OUTPUT_FOLDER_ID);
    var akhsDocx = exportGoogleDocToDocxFile_(akhsDocId, akhsFileName_(req.paperName) + ".docx", ACTIVE_OUTPUT_FOLDER_ID);
    var updateUsage = !!req.updateUsage;
    if (updateUsage && selectedItems.length){
      var projectsForThisPaper = getProjectsForThisPaper_(req);
      if (!projectsForThisPaper.length){
        notes.push("Project Used not updated: no project keyword detected. Add a project keyword in Paper Name (e.g., AITS, Milestone, VPRP...).");
      }
      incrementUsedAndProjectUsedMulti_(selectedItems, metaRefs, projectsForThisPaper);
    } else if (!updateUsage) {
      notes.push("NOTE: Used / Project Used NOT updated (Faculty draft mode).");
    }

    var taggingSheetUrl = createTaggingSheetForPaper_(req.paperName, selectedItems, startNo, includeVideoLinks, ACTIVE_OUTPUT_FOLDER_ID);
    var taggingSheetId = extractDriveIdFromUrl_(taggingSheetUrl);
    var taggingDocId = createTaggingDocForPaper_(req.paperName, selectedItems, startNo, includeVideoLinks, ACTIVE_OUTPUT_FOLDER_ID);
    var taggingDocx = exportGoogleDocToDocxFile_(taggingDocId, tgDocFileName_(req.paperName) + ".docx", ACTIVE_OUTPUT_FOLDER_ID);
    return {
      question_paper_pdf_url: "https://drive.google.com/file/d/" + qpPdf.id + "/view",
      answer_hints_pdf_url: "https://drive.google.com/file/d/" + akhsPdf.id + "/view",
      question_paper_docx_url: "https://drive.google.com/file/d/" + qpDocx.id + "/view",
      answer_hints_docx_url: "https://drive.google.com/file/d/" + akhsDocx.id + "/view",
      tagging_sheet_url: taggingSheetUrl,
      tagging_docx_url: "https://drive.google.com/file/d/" + taggingDocx.id + "/view",
      output_folder_id: ACTIVE_OUTPUT_FOLDER_ID,
      skipped_replaced_count: 0,
      notes: notes
    };
  } finally {
    lock.releaseLock();
  }
}

/***********************
 * DTP SUPPORT
 ***********************/
function createInfoTxtForDtp(req){
  req = req || {};
  var paperName = String(req.paperName||"").trim();
  var className = String(req.className||"").trim();
  var subject   = String(req.subject||"").trim();
  var planFileId = String(req.planFileId||"").trim();
  var outputFolderId = String(req.outputFolderId||"").trim();
  var qidCsv = String(req.qidCsv||"").trim();
  if (!paperName || !className || !subject) throw new Error("info.txt: paperName/className/subject required.");
  if (!planFileId) throw new Error("info.txt: missing planFileId (generate paper first).");
  if (!qidCsv) throw new Error("info.txt: provide QIDs (comma-separated).");
  var chaptersConfig = Array.isArray(req.chaptersConfig) ? req.chaptersConfig : [];
  if (!chaptersConfig.length) throw new Error("info.txt: chaptersConfig missing.");
  var projects = detectProjectsFromPaperName_(paperName);
  var projectLine = projects.length ? projects.join(", ") : "";
  var chapterLines = chaptersConfig.map(function(c){
    var ch = String(c.chapterName||"").trim();
    var topicIds = Array.isArray(c.topicIds) ? c.topicIds.filter(Boolean) : [];
    var mode = topicIds.length ? "P" : "C";
    return ch + " (" + mode + ")";
  });
  var text =
    "Paper Name: " + paperName + "\n" +
    "Class: " + className + "\n" +
    "Subject: " + subject + "\n" +
    "Project: " + projectLine + "\n" +
    "PlanFileId: " + planFileId + "\n" +
    "OutputFolderId: " + outputFolderId + "\n" +
    "QIDs: " + qidCsv + "\n" +
    "Chapters:\n" +
    chapterLines.map(function(x){ return "- " + x; }).join("\n") + "\n";
  var fileName = "info_" + collapseSpaces_(paperName).replace(/[^a-zA-Z0-9 _-]+/g,"").trim() + ".txt";
  var file = DriveApp.createFile(fileName, text, MimeType.PLAIN_TEXT);
  if (outputFolderId) moveDriveFileToFolder_(file.getId(), outputFolderId);
  else moveDriveFileToFolder_(file.getId(), OUTPUT_FOLDER_ID);
  return {
    name: file.getName(),
    url: "https://drive.google.com/file/d/" + file.getId() + "/view",
    fileId: file.getId()
  };
}

function parseInfoTxt_(text){
  var t = String(text||"");
  var lines = t.split(/\r?\n/);
  function getVal_(re){
    for (var i=0;i<lines.length;i++){
      var m = String(lines[i]||"").match(re);
      if (m) return String(m[1]||"").trim();
    }
    return "";
  }
  return {
    paperName: getVal_(/^\s*Paper Name:\s*(.+)\s*$/i),
    className: getVal_(/^\s*Class:\s*(.+)\s*$/i),
    subject: getVal_(/^\s*Subject:\s*(.+)\s*$/i),
    project: getVal_(/^\s*Project:\s*(.*)\s*$/i),
    planFileId: getVal_(/^\s*PlanFileId:\s*(.+)\s*$/i),
    outputFolderId: getVal_(/^\s*OutputFolderId:\s*(.*)\s*$/i)
  };
}

function readPlanJsonByFileId_(fileId){
  var blob = DriveApp.getFileById(fileId).getBlob();
  var txt = blob.getDataAsString("UTF-8");
  return JSON.parse(txt);
}

function parseQidCsv_(s){
  return String(s || "").split(",").map(function(x){ return String(x || "").trim(); }).filter(function(x){ return x; });
}

function makeQidReader_(){
  var cache = {};
  return function(metaKey, rowNum){
    if (!metaKey || !rowNum) return "";
    if (!cache[metaKey]){
      var parts = String(metaKey).split("::");
      var ssId = parts[0], sheetId = parts[1];
      var ss = SpreadsheetApp.openById(ssId);
      var sh = null;
      var sheets = ss.getSheets();
      for (var i=0;i<sheets.length;i++){
        if (String(sheets[i].getSheetId()) === String(sheetId)){ sh = sheets[i]; break; }
      }
      if (!sh) throw new Error("QID lookup: sheet not found for metaKey=" + metaKey);
      var lastCol = sh.getLastColumn();
      var headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(x){ return String(x).trim().toLowerCase(); });
      var qidCol = headers.indexOf("qid");
      if (qidCol === -1) qidCol = headers.indexOf("questionid");
      if (qidCol === -1) throw new Error("QID lookup: metadata sheet missing QID column.");
      cache[metaKey] = { sheet: sh, qidCol: qidCol + 1 };
    }
    var ref = cache[metaKey];
    return String(ref.sheet.getRange(rowNum, ref.qidCol).getValue() || "").trim();
  };
}

function buildSelectedItemsFromPlanByQids_(plan, qids){
  var entries = (plan && plan.selectedQuestions) ? plan.selectedQuestions : [];
  if (!entries.length) throw new Error("Plan JSON has no selectedQuestions.");
  var readQid = makeQidReader_();
  var qidToEntry = {};
  for (var i=0;i<entries.length;i++){
    var e = entries[i];
    var qid = readQid(e.metaKey, e.metaRowNum);
    if (qid && !qidToEntry[qid]) qidToEntry[qid] = e;
  }
  var out = [];
  for (var j=0;j<qids.length;j++){
    var q = qids[j];
    var pe = qidToEntry[q];
    if (!pe) throw new Error("QID not found in draft plan (cannot finalize): " + q);
    var sourceDocId   = ensureGoogleDocId_(pe.questionFileId);
    var answerKeyDocId= ensureGoogleDocId_(pe.answerKeyFileId);
    var hintsDocId    = ensureGoogleDocId_(pe.hintsFileId);
    out.push({
      chapter: pe.chapter,
      sourceSheet: pe.sourceSheet,
      difficulty: Number(pe.difficulty||0)||0,
      topicId: pe.topicId||"",
      questionType: pe.questionType||"",
      qTypeNorm: normStr_(pe.questionType||""),
      sourceDocId: sourceDocId,
      answerKeyDocId: answerKeyDocId,
      hintsDocId: hintsDocId,
      questionFileId: pe.questionFileId,
      answerKeyFileId: pe.answerKeyFileId,
      hintsFileId: pe.hintsFileId,
      origQNo: Number(pe.origQNo||0)||0,
      metaKey: pe.metaKey,
      metaRowNum: pe.metaRowNum
    });
  }
  return out;
}

function buildMetaRefsForItems_(items){
  var refs = {};
  for (var i=0;i<items.length;i++){
    var mk = items[i].metaKey;
    if (!mk || refs[mk]) continue;
    var parts = String(mk).split("::");
    var ssId = parts[0], sheetId = parts[1];
    var ss = SpreadsheetApp.openById(ssId);
    var sh = null;
    var sheets = ss.getSheets();
    for (var s=0;s<sheets.length;s++){
      if (String(sheets[s].getSheetId()) === String(sheetId)){ sh = sheets[s]; break; }
    }
    if (!sh) throw new Error("MetaRefs: sheet not found for metaKey=" + mk);
    var meta = readMeta_(sh, { requireTopic:false, requireQuestionType:false });
    refs[mk] = { sheet: sh, usedColIndex: meta.usedColIndex, projectUsedColIndex: meta.projectUsedColIndex };
  }
  return refs;
}

function createTaggingDocForPaper_(paperName, selectedItems, startNo, includeVideoLinks, folderId){
  var doc = DocumentApp.create(tgDocFileName_(paperName));
  var docId = doc.getId();
  moveDriveFileToFolder_(docId, folderId || OUTPUT_FOLDER_ID);
  var body = doc.getBody();
  body.clear();
  var paras = body.getParagraphs();
  var titleP = paras && paras.length ? paras[0] : body.appendParagraph("");
  titleP.setText("Tagging Sheet");
  try { titleP.setBold(true); titleP.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
  body.appendParagraph(SAFE_SPACE);
  var headers = includeVideoLinks
    ? ["FinalQNo","QID","Tag_SubjectCode","Chapter","Topic","Subtopic","Difficulty","Question type","Video link"]
    : ["FinalQNo","QID","Tag_SubjectCode","Chapter","Topic","Subtopic","Difficulty","Question type"];
  var table = body.appendTable([headers]);
  var cache = {};
  for (var i=0;i<selectedItems.length;i++){
    var item = selectedItems[i];
    var finalQNo = startNo + i;
    if (!cache[item.metaKey]) cache[item.metaKey] = buildMetaIndex_(item);
    var metaObj = getMetaRowObject_(cache[item.metaKey], item.metaRowNum);
    var qType = (item.questionType && String(item.questionType).trim())
      ? String(item.questionType).trim()
      : (metaObj["Question type"] || "");
    var row = [
      String(finalQNo),
      String(metaObj.QID || ""),
      String(metaObj.Tag_SubjectCode || ""),
      String(metaObj.Chapter || ""),
      String(metaObj.Topic || ""),
      String(metaObj.Subtopic || ""),
      String(metaObj.Difficulty || ""),
      String(qType || "")
    ];
    if (includeVideoLinks) row.push(String(metaObj["Video link"] || ""));
    var tr = table.appendTableRow();
    for (var c=0;c<row.length;c++) tr.appendTableCell(row[c]);
  }
  doc.saveAndClose();
  return docId;
}

function dtpFinalizePaper(payload){
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try{
    payload = payload || {};
    var infoText = String(payload.infoText||"");
    var qidCsv = String(payload.qidCsv||"").trim();
    if (!infoText) throw new Error("DTP: upload info.txt.");
    if (!qidCsv) throw new Error("DTP: provide QIDs (comma-separated).");
    var info = parseInfoTxt_(infoText);
    if (!info.planFileId) throw new Error("DTP: info.txt missing PlanFileId.");
    var plan = readPlanJsonByFileId_(info.planFileId);
    var paperName = info.paperName || plan.paperName || "Final Paper";
    var subject   = info.subject   || plan.subject   || "";
    if (!subject) throw new Error("DTP: subject not found in info.txt or plan JSON.");
    var overrideFolder = String(payload.outputFolderId||"").trim();
    ACTIVE_OUTPUT_FOLDER_ID = overrideFolder || info.outputFolderId || getSubjectOutputFolderId_(subject);
    var qids = parseQidCsv_(qidCsv);
    var selectedItems = buildSelectedItemsFromPlanByQids_(plan, qids);
    if (!selectedItems.length) throw new Error("DTP: no questions after applying QID order.");
    var startNo = getQuestionStartNo_(subject);
    var syllabusText = "";
    if (plan && plan.chaptersConfig && plan.chaptersConfig.length){
      syllabusText = formatSyllabusFromChaptersConfig_(plan.chaptersConfig);
    } else {
      syllabusText = formatSyllabusText_(selectedItems.map(function(x){ return x.chapter; }));
    }
    var qpTemplateId = resolveTemplateDocId_((DTP_TEMPLATE_DOC_ID && String(DTP_TEMPLATE_DOC_ID).trim()) ? DTP_TEMPLATE_DOC_ID : TEMPLATE_DOC_ID);
    var qpDocId = buildQuestionPaperDoc_({
      templateId: qpTemplateId,
      paperName: paperName,
      syllabusText: syllabusText,
      selectedItems: selectedItems,
      startNo: startNo,
      folderId: ACTIVE_OUTPUT_FOLDER_ID
    });
    var akDocId = buildAnswerHintsDoc_({
      baseHintsTemplateId: selectedItems[0].hintsDocId,
      paperName: paperName,
      selectedItems: selectedItems,
      startNo: startNo,
      includeHints: true,
      folderId: ACTIVE_OUTPUT_FOLDER_ID
    });
    var includeVideoLinks = (plan && plan.includeVideoLinks !== false);
    var tgDocId = createTaggingDocForPaper_(paperName, selectedItems, startNo, includeVideoLinks, ACTIVE_OUTPUT_FOLDER_ID);
    var qpDocx = exportGoogleDocToDocxFile_(qpDocId, qpFileName_(paperName) + ".docx", ACTIVE_OUTPUT_FOLDER_ID);
    var akDocx = exportGoogleDocToDocxFile_(akDocId, akhsFileName_(paperName) + ".docx", ACTIVE_OUTPUT_FOLDER_ID);
    var tgDocx = exportGoogleDocToDocxFile_(tgDocId, tgDocFileName_(paperName) + ".docx", ACTIVE_OUTPUT_FOLDER_ID);
    var previewItems = hydratePreviewItems_(selectedItems, startNo);
    var pptInfo = null;
    var zipInfo = null;
    try {
      pptInfo = createQuestionPptForPaper_(paperName, previewItems, ACTIVE_OUTPUT_FOLDER_ID);
    } catch (ePpt) {}
    try {
      var zipIds = [qpDocx.id, akDocx.id, tgDocx.id];
      if (pptInfo && pptInfo.pptxFileId) zipIds.push(pptInfo.pptxFileId);
      var zipExtraBlobs = (pptInfo && pptInfo.exportAssets)
        ? pptInfo.exportAssets.map(function(x){ return Utilities.newBlob(x.blob.getBytes(), x.blob.getContentType(), x.zipPath); })
        : [];
      zipInfo = createZipBundleForPaper_(paperName, zipIds, ACTIVE_OUTPUT_FOLDER_ID, zipExtraBlobs);
    } catch (eZip) {}
    var metaRefs = buildMetaRefsForItems_(selectedItems);
    var projectsForThisPaper = [];
    if (String(info.project||"").trim()){
      projectsForThisPaper = String(info.project).split(",").map(function(x){ return String(x||"").trim(); }).filter(Boolean);
    } else {
      projectsForThisPaper = detectProjectsFromPaperName_(paperName);
    }
    incrementUsedAndProjectUsedMulti_(selectedItems, metaRefs, projectsForThisPaper);
    return {
      question_paper_docx_url: "https://drive.google.com/file/d/" + qpDocx.id + "/view",
      answer_key_docx_url:     "https://drive.google.com/file/d/" + akDocx.id + "/view",
      tagging_docx_url:        "https://drive.google.com/file/d/" + tgDocx.id + "/view",
      ppt_url:                 pptInfo ? pptInfo.pptxUrl : "",
      zip_url:                 zipInfo ? zipInfo.url : "",
      output_folder_id:        ACTIVE_OUTPUT_FOLDER_ID
    };
  } finally {
    lock.releaseLock();
  }
}

/***********************
 * HELPER: manual DTP token set
 ***********************/
function setDtpTokenManual() {
  var token = "dtp12345";
  PropertiesService.getScriptProperties().setProperty("DTP_TOKEN", token);
  Logger.log("DTP_TOKEN set to: " + token);
}
