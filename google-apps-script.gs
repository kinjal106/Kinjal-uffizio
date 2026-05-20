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
// Handles ALL formats: Date objects, serial numbers, DD-MMM-YY, YYYY-MM-DD etc
var MONTH_MAP = {
  jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",
  jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"
};

function fmtDate(v) {
  if (!v && v !== 0) return "";
  
  // Case 1: Proper Date object
  if (v instanceof Date) {
    var y = v.getFullYear();
    var m = v.getMonth() + 1;
    var d = v.getDate();
    if (y < 2020 || y > 2035) {
      y = v.getUTCFullYear();
      m = v.getUTCMonth() + 1;
      d = v.getUTCDate();
    }
    if (y >= 2020 && y <= 2035) return y + "-" + pad(m) + "-" + pad(d);
  }
  
  // Case 2: Excel serial number
  if (typeof v === 'number') {
    var excelEpoch = new Date(1899, 11, 30);
    var dt = new Date(excelEpoch.getTime() + v * 86400000);
    var y2 = dt.getFullYear();
    if (y2 >= 2020 && y2 <= 2035)
      return y2 + "-" + pad(dt.getMonth()+1) + "-" + pad(dt.getDate());
    // Try 1904 system
    var dt2 = new Date(new Date(1904,0,1).getTime() + v * 86400000);
    var y3 = dt2.getFullYear();
    if (y3 >= 2020 && y3 <= 2035)
      return y3 + "-" + pad(dt2.getMonth()+1) + "-" + pad(dt2.getDate());
  }
  
  var s = String(v).trim();
  
  // Case 3: "DD-MMM-YY" or "DD-MMM-YYYY" e.g. "19-May-26" or "19-May-2026"
  var m1 = s.match(/^(\d{1,2})[\-\/]([A-Za-z]{3})[\-\/](\d{2,4})$/);
  if (m1) {
    var mon = MONTH_MAP[m1[2].toLowerCase()];
    if (mon) {
      var yr = m1[3].length === 2 ? "20" + m1[3] : m1[3];
      if (parseInt(yr) >= 2020 && parseInt(yr) <= 2035)
        return yr + "-" + mon + "-" + pad(parseInt(m1[1]));
    }
  }
  
  // Case 4: "MMM-DD-YY" or "MMM DD, YYYY"
  var m2 = s.match(/^([A-Za-z]{3})[\-\/\s](\d{1,2})[,\s]+(\d{2,4})$/);
  if (m2) {
    var mon2 = MONTH_MAP[m2[1].toLowerCase()];
    if (mon2) {
      var yr2 = m2[3].length === 2 ? "20" + m2[3] : m2[3];
      if (parseInt(yr2) >= 2020 && parseInt(yr2) <= 2035)
        return yr2 + "-" + mon2 + "-" + pad(parseInt(m2[2]));
    }
  }

  // Case 5: "YYYY-MM-DD" already correct
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    var yr3 = parseInt(s.substring(0,4));
    if (yr3 >= 2020 && yr3 <= 2035) return s.substring(0,10);
  }
  
  // Case 6: "DD/MM/YYYY" or "MM/DD/YYYY"
  var m3 = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m3) {
    var yr4 = parseInt(m3[3]);
    if (yr4 >= 2020 && yr4 <= 2035)
      return yr4 + "-" + pad(parseInt(m3[2])) + "-" + pad(parseInt(m3[1]));
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
