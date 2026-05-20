// ╔══════════════════════════════════════════════════════════════╗
// ║  UFFIZIO KRA BRIDGE — Date-Fixed Version                    ║
// ║  Sheet ID: 1XZ32kMV32ASmgxXTF26sIPIoqRtB8ADOQ4j1__9sabM   ║
// ╚══════════════════════════════════════════════════════════════╝

var SPREADSHEET_ID = "1XZ32kMV32ASmgxXTF26sIPIoqRtB8ADOQ4j1__9sabM";
var SHEET_NAME     = "All Tasks";
var HEADER_ROW     = 3;

function doGet(e) {
  var action   = (e.parameter && e.parameter.action) || "getTasks";
  var callback = (e.parameter && e.parameter.callback) || null;
  var result;
  try {
    if (action === "getTasks")      result = getTasks();
    else if (action === "saveRatings") result = saveRatingsFromGet(e.parameter);
    else if (action === "addTask")     result = addTaskFromGet(e.parameter);
    else result = { error: "Unknown action" };
  } catch(err) {
    result = { error: err.toString() };
  }
  var json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}


// ── SAVE RATINGS via GET (JSONP) ──────────────────────────────
function saveRatingsFromGet(params) {
  var body = {
    rowId:    params.rowId,
    score:    parseFloat(params.score) || 0,
    note:     decodeURIComponent(params.note || ""),
    assignee: decodeURIComponent(params.assignee || ""),
    task:     decodeURIComponent(params.task || ""),
    month:    params.month || ""
  };
  for (var k = 1; k <= 10; k++) {
    body["r"+k] = decodeURIComponent(params["r"+k] || "");
  }
  return saveRatings(body);
}

// ── ADD TASK via GET (JSONP) ───────────────────────────────────
function addTaskFromGet(params) {
  var body = {
    task:     decodeURIComponent(params.task || ""),
    assignee: decodeURIComponent(params.assignee || ""),
    priority: params.priority || "Medium",
    start:    params.start || "",
    end:      params.end || "",
    status:   params.status || "Todo",
    jira:     params.jira || ""
  };
  return addTask(body);
}

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

// ── SMART DATE FORMATTER ──────────────────────────────────────
// Handles all cases: Date objects, serial numbers, strings
function fmtDate(v) {
  if (!v && v !== 0) return "";
  
  // Case 1: Already a proper Date object
  if (v instanceof Date) {
    var y = v.getFullYear();
    var m = v.getMonth() + 1;
    var d = v.getDate();
    // If year looks wrong, try UTC values
    if (y < 2020 || y > 2035) {
      y = v.getUTCFullYear();
      m = v.getUTCMonth() + 1;
      d = v.getUTCDate();
    }
    if (y >= 2020 && y <= 2035) {
      return y + "-" + pad(m) + "-" + pad(d);
    }
  }
  
  // Case 2: Numeric serial number (Excel date)
  if (typeof v === 'number') {
    // Convert Excel serial to date
    // Excel epoch: Dec 30, 1899 (accounting for the Lotus 1-2-3 bug)
    var excelEpoch = new Date(1899, 11, 30);
    var dateFromSerial = new Date(excelEpoch.getTime() + v * 86400000);
    var y2 = dateFromSerial.getFullYear();
    var m2 = dateFromSerial.getMonth() + 1;
    var d2 = dateFromSerial.getDate();
    if (y2 >= 2020 && y2 <= 2035) {
      return y2 + "-" + pad(m2) + "-" + pad(d2);
    }
    // Try 1904 date system (Mac Excel)
    var epoch1904 = new Date(1904, 0, 1);
    var dateFrom1904 = new Date(epoch1904.getTime() + v * 86400000);
    var y3 = dateFrom1904.getFullYear();
    if (y3 >= 2020 && y3 <= 2035) {
      return y3 + "-" + pad(dateFrom1904.getMonth()+1) + "-" + pad(dateFrom1904.getDate());
    }
  }
  
  // Case 3: String date - try to parse
  var s = String(v).trim();
  // Already formatted correctly: "2026-04-13"
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    var yr = parseInt(s.substring(0,4));
    if (yr >= 2020 && yr <= 2035) return s.substring(0,10);
  }
  // DD/MM/YYYY or MM/DD/YYYY
  if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(s)) {
    var parts = s.split(/[\/\-]/);
    var yr2 = parseInt(parts[2]);
    if (yr2 >= 2020 && yr2 <= 2035) {
      // Assume DD/MM/YYYY (Indian format)
      return yr2 + "-" + pad(parseInt(parts[1])) + "-" + pad(parseInt(parts[0]));
    }
  }
  return "";
}

function pad(n) { return n < 10 ? "0" + n : "" + n; }

// ── GET ALL TASKS ─────────────────────────────────────────────
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

    var startStr = fmtDate(row[col("Start Date")]);
    var endStr   = fmtDate(row[col("End Date")]);
    var month    = startStr ? startStr.substring(0, 7) : "";
    var monthLabel = "";
    if (month) {
      try {
        var mp = month.split("-");
        var months_arr = ["January","February","March","April","May","June",
                          "July","August","September","October","November","December"];
        monthLabel = months_arr[parseInt(mp[1])-1] + " " + mp[0];
      } catch(e) { monthLabel = month; }
    }

    tasks.push({
      id:         i,
      task:       task,
      assignee:   assignee,
      priority:   String(row[col("Priority")]  || "Medium").trim(),
      start:      startStr,
      end:        endStr,
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

// ── GET RATINGS ───────────────────────────────────────────────
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

// ── SAVE RATINGS ──────────────────────────────────────────────
function saveRatings(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName("KRA_Ratings");
  if (!sh) {
    sh = ss.insertSheet("KRA_Ratings");
    sh.appendRow(["RowID","KPI1","KPI2","KPI3","KPI4","KPI5",
                  "KPI6","KPI7","KPI8","KPI9","KPI10",
                  "Score","Note","Assignee","Task","Month","Saved_At"]);
    sh.getRange(1,1,1,17).setFontWeight("bold")
      .setBackground("#534AB7").setFontColor("#ffffff");
  }
  var data = sh.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == body.rowId) { foundRow = i + 1; break; }
  }
  var row = [
    body.rowId,
    body.r1||"",body.r2||"",body.r3||"",body.r4||"",body.r5||"",
    body.r6||"",body.r7||"",body.r8||"",body.r9||"",body.r10||"",
    body.score||0, body.note||"",
    body.assignee||"", body.task||"", body.month||"",
    new Date().toISOString()
  ];
  if (foundRow > 0) sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { saved: true };
}

// ── ADD TASK ──────────────────────────────────────────────────
function addTask(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { error: "Sheet not found" };
  sh.appendRow([
    "", body.task||"", body.start||"", body.end||"", "",
    body.assignee||"", body.priority||"Medium",
    "", "", "", body.status||"Todo", body.jira||"", ""
  ]);
  return { added: true };
}
