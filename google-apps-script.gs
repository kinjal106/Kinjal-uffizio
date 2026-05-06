// ── UFFIZIO KRA BRIDGE ──────────────────────────────────────
// Paste this entire file into Google Apps Script
// Extensions → Apps Script → Delete everything → Paste → Save → Deploy
// ────────────────────────────────────────────────────────────

const SHEET_NAME = "All Tasks";
const HEADER_ROW = 3; // Row number where headers are (Task, Assignee...)

function doGet(e) {
  const action = e.parameter.action || "getTasks";
  if (action === "getTasks") return getTasks();
  if (action === "saveRatings") return saveRatings(e);
  return ok({error:"unknown action"});
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.action === "saveRatings") return saveRatingsPost(body);
  if (body.action === "addTask") return addTask(body);
  return ok({error:"unknown action"});
}

function getTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  const hIdx = HEADER_ROW - 1;
  const cols = data[hIdx];
  const col = (n) => cols.indexOf(n);
  const tasks = [];
  const members = ["Harshil","Hinesh","Mansi","Vishal","Mayur","Kinjal"];
  for (let i = hIdx+1; i < data.length; i++) {
    const row = data[i];
    const assignee = String(row[col("Assignee")]||"").trim();
    const task = String(row[col("Task")]||"").trim();
    if (!task || !members.includes(assignee)) continue;
    tasks.push({
      id: i,
      task: task,
      assignee: assignee,
      priority: String(row[col("Priority")]||"Medium").trim(),
      start: formatDate(row[col("Start Date")]),
      end: formatDate(row[col("End Date")]),
      status: String(row[col("Status")]||"Todo").trim(),
      estHrs: row[col("Est. Hrs")] || null,
      actHrs: row[col("Act. Hrs")] || null,
      jira: String(row[col("Jira")]||"").trim(),
      notes: String(row[col("Notes")]||"").trim(),
      ratings: getRatings(i)
    });
  }
  return ok({tasks});
}

function getRatings(rowId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("KRA_Ratings");
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  const ratings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == rowId) {
      for (let k = 1; k <= 7; k++) ratings[k] = data[i][k] || 0;
      return ratings;
    }
  }
  return {};
}

function saveRatingsPost(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("KRA_Ratings");
  if (!sh) {
    sh = ss.insertSheet("KRA_Ratings");
    sh.appendRow(["RowID","KPI1","KPI2","KPI3","KPI4","KPI5","KPI6","KPI7",
                  "Score","Note","Assignee","Task","Timestamp"]);
  }
  const data = sh.getDataRange().getValues();
  let found = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == body.rowId) { found = i+1; break; }
  }
  const row = [
    body.rowId,
    body.r1||0, body.r2||0, body.r3||0, body.r4||0,
    body.r5||0, body.r6||0, body.r7||0,
    body.score||0, body.note||"",
    body.assignee||"", body.task||"",
    new Date().toISOString()
  ];
  if (found > 0) sh.getRange(found, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return ok({saved:true});
}

function addTask(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  sh.appendRow([
    "", body.task, body.start||"", body.end||"", "",
    body.assignee, body.priority||"Medium",
    body.estHrs||"", "", "",
    body.status||"Todo", body.jira||"", body.notes||""
  ]);
  return ok({added:true});
}

function formatDate(v) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd");
  return String(v).substring(0,10);
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
