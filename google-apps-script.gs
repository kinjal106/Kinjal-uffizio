// ── UFFIZIO KRA BRIDGE — Final Version with JSONP ────────────
// IMPORTANT: After pasting this, do a NEW deployment (not update)
// Deploy → New deployment → Web app → Anyone → Deploy
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = "All Tasks";
const HEADER_ROW = 3;

// ── GET handler — supports both regular fetch and JSONP ───────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || "getTasks";
  const callback = e.parameter && e.parameter.callback; // for JSONP

  let result;
  try {
    if (action === "getTasks") result = getTasks();
    else result = { error: "unknown action" };
  } catch (err) {
    result = { error: err.toString() };
  }

  const json = JSON.stringify(result);

  // If callback param present → JSONP (bypasses CORS completely)
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Otherwise regular JSON with CORS headers
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST handler ──────────────────────────────────────────────
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}

  let result;
  try {
    if (body.action === "saveRatings") result = saveRatings(body);
    else if (body.action === "addTask") result = addTask(body);
    else result = { error: "unknown action" };
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET ALL TASKS ─────────────────────────────────────────────
function getTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { error: "Sheet '" + SHEET_NAME + "' not found" };

  const data = sh.getDataRange().getValues();
  const hIdx = HEADER_ROW - 1;
  const cols = data[hIdx];
  const col = (n) => cols.indexOf(n);
  const members = ["Harshil","Hinesh","Mansi","Vishal","Mayur","Kinjal"];
  const tasks = [];

  for (let i = hIdx + 1; i < data.length; i++) {
    const row = data[i];
    const assignee = String(row[col("Assignee")] || "").trim();
    const task = String(row[col("Task")] || "").trim();
    if (!task || !members.includes(assignee)) continue;

    const sd = row[col("Start Date")];
    const month = sd instanceof Date
      ? Utilities.formatDate(sd, "Asia/Kolkata", "yyyy-MM") : "";
    const monthLabel = sd instanceof Date
      ? Utilities.formatDate(sd, "Asia/Kolkata", "MMMM yyyy") : "Unassigned";

    tasks.push({
      id: i, task, assignee,
      priority: String(row[col("Priority")] || "Medium").trim(),
      start: formatDate(row[col("Start Date")]),
      end: formatDate(row[col("End Date")]),
      status: String(row[col("Status")] || "Todo").trim(),
      jira: String(row[col("Jira")] || "").trim(),
      notes: String(row[col("Notes")] || "").trim(),
      month, monthLabel,
      ratings: getRatings(i)
    });
  }
  return { tasks, total: tasks.length };
}

// ── GET SAVED RATINGS ─────────────────────────────────────────
function getRatings(rowId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("KRA_Ratings");
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == rowId) {
      const r = {};
      for (let k = 1; k <= 10; k++) r[k] = data[i][k] || "";
      return r;
    }
  }
  return {};
}

// ── SAVE RATINGS ──────────────────────────────────────────────
function saveRatings(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("KRA_Ratings");
  if (!sh) {
    sh = ss.insertSheet("KRA_Ratings");
    sh.appendRow([
      "RowID","KPI1","KPI2","KPI3","KPI4","KPI5",
      "KPI6","KPI7","KPI8","KPI9","KPI10",
      "Score","Note","Assignee","Task","Month","Timestamp"
    ]);
    sh.getRange(1,1,1,17).setFontWeight("bold")
      .setBackground("#4a4e7a").setFontColor("#ffffff");
  }

  const data = sh.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == body.rowId) { foundRow = i + 1; break; }
  }

  const row = [
    body.rowId,
    body.r1||"", body.r2||"", body.r3||"", body.r4||"", body.r5||"",
    body.r6||"", body.r7||"", body.r8||"", body.r9||"", body.r10||"",
    body.score||0, body.note||"",
    body.assignee||"", body.task||"", body.month||"",
    new Date().toISOString()
  ];

  if (foundRow > 0) sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);

  // Auto-update KRA_Summary tab
  updateSummary(body.assignee, body.month);
  return { saved: true };
}

// ── UPDATE KRA SUMMARY TAB ────────────────────────────────────
function updateSummary(assignee, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName("KRA_Summary");
    if (!sh) {
      sh = ss.insertSheet("KRA_Summary");
      sh.appendRow([
        "Member","Month",
        "KPI1_Avg","KPI2_Avg","KPI3_Avg","KPI4_Avg","KPI5_Avg",
        "KPI6_Avg","KPI7_Avg","KPI8_Avg","KPI9_Avg","KPI10_Avg",
        "KRA_Score","Grade","Tasks_Rated","Updated"
      ]);
      sh.getRange(1,1,1,16).setFontWeight("bold")
        .setBackground("#4a4e7a").setFontColor("#ffffff");
    }

    const weights = [12,12,16,12,8,8,12,10,5,5];
    const ratingSh = ss.getSheetByName("KRA_Ratings");
    if (!ratingSh) return;
    const rData = ratingSh.getDataRange().getValues();

    // Find all row IDs for this member+month from the task sheet
    const tSh = ss.getSheetByName(SHEET_NAME);
    const tData = tSh.getDataRange().getValues();
    const hIdx = HEADER_ROW - 1;
    const cols = tData[hIdx];
    const col = (n) => cols.indexOf(n);

    const rowIds = [];
    for (let i = hIdx+1; i < tData.length; i++) {
      const a = String(tData[i][col("Assignee")]||"").trim();
      const sd = tData[i][col("Start Date")];
      const m = sd instanceof Date ? Utilities.formatDate(sd,"Asia/Kolkata","yyyy-MM") : "";
      if (a === assignee && m === month) rowIds.push(i);
    }

    // Average KPI scores from all rated tasks
    const kpiSums = new Array(10).fill(0);
    let ratedCount = 0;

    rowIds.forEach(rid => {
      for (let i = 1; i < rData.length; i++) {
        if (rData[i][0] == rid) {
          const allRated = [1,2,3,4,5,6,7,8,9,10].every(k => rData[i][k] !== "" && rData[i][k] !== null);
          if (allRated) {
            ratedCount++;
            for (let k = 0; k < 10; k++) {
              const band = String(rData[i][k+1] || "");
              const score = band === "100" ? 100 : (parseInt((band.split("-")[1])) || 0);
              kpiSums[k] += score;
            }
          }
          break;
        }
      }
    });

    if (ratedCount === 0) return;

    const kpiAvgs = kpiSums.map(s => Math.round(s / ratedCount * 10) / 10);
    let kraScore = 0;
    kpiAvgs.forEach((avg, i) => { kraScore += (avg / 100) * weights[i]; });
    kraScore = Math.round(kraScore * 10) / 10;
    const grade = kraScore >= 85 ? "A - Exceeds" : kraScore >= 70 ? "B - Meets" : kraScore >= 50 ? "C - Needs Improvement" : "D - Poor";

    const sData = sh.getDataRange().getValues();
    let sRow = -1;
    for (let i = 1; i < sData.length; i++) {
      if (sData[i][0] === assignee && sData[i][1] === month) { sRow = i+1; break; }
    }

    const summaryRow = [assignee, month, ...kpiAvgs, kraScore, grade, ratedCount, new Date().toISOString()];
    if (sRow > 0) sh.getRange(sRow, 1, 1, summaryRow.length).setValues([summaryRow]);
    else sh.appendRow(summaryRow);
  } catch(e) {}
}

// ── ADD TASK ──────────────────────────────────────────────────
function addTask(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { error: "Sheet not found" };
  sh.appendRow([
    "", body.task||"", body.start||"", body.end||"", "",
    body.assignee||"", body.priority||"Medium",
    "", "", "", body.status||"Todo", body.jira||"", body.notes||""
  ]);
  return { added: true };
}

// ── DATE FORMATTER ────────────────────────────────────────────
function formatDate(v) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd");
  return String(v).substring(0, 10);
}
