// ╔══════════════════════════════════════════════════════════╗
// ║   UFFIZIO KRA BRIDGE — FINAL                            ║
// ║   Sheet ID: 1XZ32kMV32ASmgxXTF26sIPIoqRtB8ADOQ4j1__9sabM ║
// ║                                                          ║
// ║   HOW TO DEPLOY:                                         ║
// ║   1. Open script.google.com                              ║
// ║   2. Paste this entire file                              ║
// ║   3. Save (Ctrl+S)                                       ║
// ║   4. Deploy → New deployment                             ║
// ║   5. Type: Web app                                       ║
// ║   6. Execute as: Me                                      ║
// ║   7. Who has access: Anyone                              ║
// ║   8. Click Deploy → Copy the URL                         ║
// ╚══════════════════════════════════════════════════════════╝

var SPREADSHEET_ID = "1XZ32kMV32ASmgxXTF26sIPIoqRtB8ADOQ4j1__9sabM";
var SHEET_NAME     = "All Tasks";
var HEADER_ROW     = 3; // Row number where headers are

// ── GET: supports both regular JSON and JSONP (no CORS issues) ──
function doGet(e) {
  var action   = (e.parameter && e.parameter.action) || "getTasks";
  var callback = (e.parameter && e.parameter.callback) || null;
  var result;
  try {
    if (action === "getTasks") result = getTasks();
    else result = { error: "Unknown action: " + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  var json = JSON.stringify(result);
  if (callback) {
    // JSONP — bypasses CORS completely, works from Vercel
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST: save ratings and add tasks ──────────────────────────
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) {}
  var result;
  try {
    if (body.action === "saveRatings") result = saveRatings(body);
    else if (body.action === "addTask")  result = addTask(body);
    else result = { error: "Unknown action" };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── READ ALL TASKS from your sheet ────────────────────────────
function getTasks() {
  var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh   = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { error: "Sheet '" + SHEET_NAME + "' not found" };

  var data = sh.getDataRange().getValues();
  var hIdx = HEADER_ROW - 1;
  var cols = data[hIdx];

  function col(name) { return cols.indexOf(name); }

  var members = ["Harshil","Hinesh","Mansi","Vishal","Mayur","Kinjal"];
  var tasks   = [];

  for (var i = hIdx + 1; i < data.length; i++) {
    var row      = data[i];
    var assignee = String(row[col("Assignee")] || "").trim();
    var task     = String(row[col("Task")]     || "").trim();
    if (!task || !members.includes(assignee)) continue;

    var sd         = row[col("Start Date")];
    var month      = sd instanceof Date ? Utilities.formatDate(sd, "Asia/Kolkata", "yyyy-MM")      : "";
    var monthLabel = sd instanceof Date ? Utilities.formatDate(sd, "Asia/Kolkata", "MMMM yyyy")    : "Unassigned";

    tasks.push({
      id:         i,
      task:       task,
      assignee:   assignee,
      priority:   String(row[col("Priority")]  || "Medium").trim(),
      start:      fmtDate(row[col("Start Date")]),
      end:        fmtDate(row[col("End Date")]),
      status:     String(row[col("Status")]    || "Todo").trim(),
      estHrs:     row[col("Est. Hrs")]  || null,
      actHrs:     row[col("Act. Hrs")]  || null,
      jira:       String(row[col("Jira")]      || "").trim(),
      notes:      String(row[col("Notes")]     || "").trim(),
      month:      month,
      monthLabel: monthLabel,
      ratings:    getSavedRatings(i)
    });
  }
  return { tasks: tasks, total: tasks.length };
}

// ── GET saved ratings for one task row ───────────────────────
function getSavedRatings(rowId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName("KRA_Ratings");
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == rowId) {
      var r = {};
      for (var k = 1; k <= 10; k++) r[k] = data[i][k] || "";
      return r;
    }
  }
  return {};
}

// ── SAVE ratings back to KRA_Ratings tab ─────────────────────
function saveRatings(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName("KRA_Ratings");
  if (!sh) {
    sh = ss.insertSheet("KRA_Ratings");
    sh.appendRow([
      "RowID","KPI1","KPI2","KPI3","KPI4","KPI5",
      "KPI6","KPI7","KPI8","KPI9","KPI10",
      "Score","Note","Assignee","Task","Month","Saved_At"
    ]);
    sh.getRange(1,1,1,17)
      .setFontWeight("bold")
      .setBackground("#534AB7")
      .setFontColor("#ffffff");
  }

  var data     = sh.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == body.rowId) { foundRow = i + 1; break; }
  }

  var row = [
    body.rowId,
    body.r1||"", body.r2||"", body.r3||"", body.r4||"",  body.r5||"",
    body.r6||"", body.r7||"", body.r8||"", body.r9||"",  body.r10||"",
    body.score   || 0,
    body.note    || "",
    body.assignee|| "",
    body.task    || "",
    body.month   || "",
    new Date().toISOString()
  ];

  if (foundRow > 0) sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
  else              sh.appendRow(row);

  updateSummaryTab(body.assignee, body.month);
  return { saved: true, rowId: body.rowId };
}

// ── AUTO-UPDATE KRA_Summary tab ───────────────────────────────
function updateSummaryTab(assignee, month) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName("KRA_Summary");
    if (!sh) {
      sh = ss.insertSheet("KRA_Summary");
      sh.appendRow([
        "Member","Month",
        "KPI1_Avg","KPI2_Avg","KPI3_Avg","KPI4_Avg","KPI5_Avg",
        "KPI6_Avg","KPI7_Avg","KPI8_Avg","KPI9_Avg","KPI10_Avg",
        "KRA_Score","Grade","Tasks_Rated","Updated_At"
      ]);
      sh.getRange(1,1,1,16)
        .setFontWeight("bold")
        .setBackground("#534AB7")
        .setFontColor("#ffffff");
    }

    var weights = [12,12,16,12,8,8,12,10,5,5];

    // Get all task row IDs for this member+month
    var tSh  = ss.getSheetByName(SHEET_NAME);
    var tData = tSh.getDataRange().getValues();
    var hIdx = HEADER_ROW - 1;
    var cols = tData[hIdx];
    function col(n){ return cols.indexOf(n); }

    var rowIds = [];
    for (var i = hIdx+1; i < tData.length; i++) {
      var a  = String(tData[i][col("Assignee")]||"").trim();
      var sd = tData[i][col("Start Date")];
      var m  = sd instanceof Date ? Utilities.formatDate(sd,"Asia/Kolkata","yyyy-MM") : "";
      if (a === assignee && m === month) rowIds.push(i);
    }

    // Sum KPI scores from KRA_Ratings
    var rSh   = ss.getSheetByName("KRA_Ratings");
    if (!rSh) return;
    var rData = rSh.getDataRange().getValues();
    var kpiSums = [0,0,0,0,0,0,0,0,0,0];
    var ratedCount = 0;

    rowIds.forEach(function(rid) {
      for (var i = 1; i < rData.length; i++) {
        if (rData[i][0] == rid) {
          var allRated = true;
          for (var k = 1; k <= 10; k++) if (!rData[i][k]) { allRated = false; break; }
          if (allRated) {
            ratedCount++;
            for (var k = 0; k < 10; k++) {
              var band  = String(rData[i][k+1]||"");
              var score = band==="100" ? 100 : (parseInt((band.split("-")[1])||"0")||0);
              kpiSums[k] += score;
            }
          }
          break;
        }
      }
    });

    if (ratedCount === 0) return;

    var kpiAvgs  = kpiSums.map(function(s){ return Math.round(s/ratedCount*10)/10; });
    var kraScore = 0;
    kpiAvgs.forEach(function(avg,i){ kraScore += (avg/100)*weights[i]; });
    kraScore = Math.round(kraScore*10)/10;
    var grade = kraScore>=85 ? "A - Exceeds"
              : kraScore>=70 ? "B - Meets"
              : kraScore>=50 ? "C - Needs Improvement"
              :                "D - Poor";

    var sData   = sh.getDataRange().getValues();
    var sRow    = -1;
    for (var i = 1; i < sData.length; i++) {
      if (sData[i][0]===assignee && sData[i][1]===month) { sRow=i+1; break; }
    }
    var summaryRow = [assignee, month].concat(kpiAvgs).concat([kraScore, grade, ratedCount, new Date().toISOString()]);
    if (sRow > 0) sh.getRange(sRow,1,1,summaryRow.length).setValues([summaryRow]);
    else          sh.appendRow(summaryRow);
  } catch(e) { /* best-effort */ }
}

// ── ADD a new task to the sheet ───────────────────────────────
function addTask(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { error: "Sheet not found" };
  sh.appendRow([
    "", body.task||"", body.start||"", body.end||"", "",
    body.assignee||"", body.priority||"Medium",
    body.estHrs||"", "", "",
    body.status||"Todo", body.jira||"", body.notes||""
  ]);
  return { added: true };
}

// ── Date formatter ────────────────────────────────────────────
function fmtDate(v) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd");
  return String(v).substring(0,10);
}
