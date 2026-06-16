// RecruitPro - Google Apps Script Backend
// Sheet ID     : 1DoOjLQd2Fx5M06p2dA_CyvpmVADD0ZKXsE68GhNzoJk
// Drive Folder : 1Lz_JuXjGFALvWxNW7s8qs8cCjWjLhiiJ

var SHEET_ID   = "1DoOjLQd2Fx5M06p2dA_CyvpmVADD0ZKXsE68GhNzoJk";
var FOLDER_ID  = "1Lz_JuXjGFALvWxNW7s8qs8cCjWjLhiiJ";
var STORE_TAB  = "AppData";
var CANDS_TAB  = "Candidates";

// ─── Serve the app ─────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("RecruitPro - Ocrolus")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Required for <?!= include("filename") ?> in templates
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── Handle all POST requests ──────────────────────────────────────────────────
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === "saveData")        return ok(saveData(body.key, body.data));
    if (action === "loadData")        return ok(loadData(body.key));
    if (action === "saveCandidate")   return ok(syncToSheet(body.candidate));
    if (action === "deleteCandidate") return ok(deleteFromSheet(body.id));
    if (action === "uploadResume")    return ok(uploadResume(body.pan, body.dateStr, body.base64Data));

    return ok({ error: "Unknown action: " + action });
  } catch(ex) {
    return ok({ error: String(ex) });
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Key-Value storage (AppData tab) ─────────────────────────────────────────
function getStoreSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STORE_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(STORE_TAB);
    sheet.appendRow(["key","data","updated_at"]);
    sheet.getRange(1,1,1,3).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#fff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveData(key, data) {
  try {
    var sheet = getStoreSheet();
    var vals  = sheet.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === key) {
        sheet.getRange(i+1,2).setValue(data);
        sheet.getRange(i+1,3).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    sheet.appendRow([key, data, new Date().toISOString()]);
    return { success: true };
  } catch(ex) { return { error: String(ex) }; }
}

function loadData(key) {
  try {
    var sheet = getStoreSheet();
    var vals  = sheet.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === key) return { data: String(vals[i][1]) };
    }
    return { data: null };
  } catch(ex) { return { error: String(ex) }; }
}

// ─── Candidates sheet (human-readable rows) ───────────────────────────────────
function getCandSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CANDS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CANDS_TAB);
    var h = ["ID","Registered At","Name","Mobile","DOB","Gender","Graduation",
             "Interview Location","Refer By","Emp ID","Is Rejoiner",
             "PAN","Aadhaar","Resume","Resume Drive URL",
             "Addr1","Addr2","Landmark","City","State","Pincode",
             "Apt Score","Apt Passed","Typing WPM","Typing Acc","Typing Passed",
             "Panel Status","Panel Remark","Assigned Panelist","Duplicate Flagged"];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#fff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function buildRow(c, driveUrl) {
  var aptPassed = "", tWpm = "", tAcc = "", tPass = "";
  if (c.aptScore !== null && c.aptScore !== undefined && c.aptScore !== "")
    aptPassed = (Number(c.aptScore) >= 10) ? "Yes" : "No";
  if (c.typingScore) {
    tWpm = c.typingScore.wpm||""; tAcc = c.typingScore.accuracy||"";
    tPass = c.typingScore.passed ? "Yes" : "No";
  }
  return [
    String(c.id||""),c.registeredAt||"",c.name||"",c.mobile||"",c.dob||"",
    c.gender||"",c.graduationDone||"",c.interviewLocation||"",
    c.referBy||"",c.empId||"",c.isRejoiner||"",c.pan||"",c.aadhaar||"",
    c.resume||"",driveUrl||"",c.addr1||"",c.addr2||"",c.landmark||"",
    c.city||"",c.state||"",c.pincode||"",
    (c.aptScore!==null&&c.aptScore!==undefined&&c.aptScore!=="") ? c.aptScore : "",
    aptPassed,tWpm,tAcc,tPass,
    c.panelStatus||"pending",c.panelRemark||"",c.assignedPanelistId||"",
    c.duplicateFlagged ? "Yes" : "No"
  ];
}

function syncToSheet(c) {
  try {
    var sheet = getCandSheet();
    var vals  = sheet.getDataRange().getValues();
    var cid   = String(c.id);
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === cid) {
        sheet.getRange(i+1,1,1,30).setValues([buildRow(c, vals[i][14])]);
        return { success: true };
      }
    }
    var row = buildRow(c,"");
    sheet.appendRow(row);
    var last = sheet.getLastRow();
    if (last%2===0) sheet.getRange(last,1,1,row.length).setBackground("#f0f4f8");
    return { success: true };
  } catch(ex) { return { error: String(ex) }; }
}

function deleteFromSheet(id) {
  try {
    var sheet = getCandSheet();
    var vals  = sheet.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; }
    }
    return { success: true };
  } catch(ex) { return { error: String(ex) }; }
}

// ─── Resume upload ────────────────────────────────────────────────────────────
function uploadResume(pan, dateStr, base64Data) {
  try {
    var filename = pan.toUpperCase() + "_" + dateStr + ".pdf";
    var b64 = base64Data;
    var ci  = base64Data.indexOf(",");
    if (ci !== -1) b64 = base64Data.substring(ci+1);
    var bytes  = Utilities.base64Decode(b64);
    var blob   = Utilities.newBlob(bytes,"application/pdf",filename);
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var old = folder.getFilesByName(filename);
    while (old.hasNext()) old.next().setTrashed(true);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = "https://drive.google.com/file/d/" + file.getId() + "/view";
    var sheet = getCandSheet();
    var vals  = sheet.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][11]).toUpperCase() === pan.toUpperCase()) {
        sheet.getRange(i+1,14).setValue(filename);
        sheet.getRange(i+1,15).setValue(url);
      }
    }
    return { success: true, url: url, filename: filename };
  } catch(ex) { return { error: String(ex) }; }
}
