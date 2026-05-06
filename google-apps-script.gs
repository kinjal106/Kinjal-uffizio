// ── UFFIZIO KRA BRIDGE v2 — CORS Fixed ──────────────────────
// Replace your existing Apps Script with this entire file
// Then re-deploy as a NEW deployment (not update existing)
// ────────────────────────────────────────────────────────────

const SHEET_NAME = "All Tasks";
const HEADER_ROW = 3;

// ── CORS HEADERS ─────────────────────────────────────────────
function setCorsHeaders(output) {
  return output
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// ── HANDLE GET ────────────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || "getTasks";
  let result;
  try {
    if (action === "getTasks") result = getTasks();
    else result = { error: "unknown action" };
  } catch (err) {
    result = { error: err.toString() };
  }
  const output = ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
  return setCorsHeaders(output);
}

// ── HANDLE POST ───────────────────────────────────────────────
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
  const output = ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
  return setCorsHeaders(output);
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

  const members = ["Harshil", "Hinesh", "Mansi", "Vishal", "Mayur", "Kinjal"];
  const tasks = [];

  for (let i = hIdx + 1; i < data.length; i++) {
    const row = data[i];
    const assignee = String(row[col("Assignee")] || "").trim();
    const task = String(row[col("Task")] || "").trim();
    if (!task || !members.includes(assignee)) continue;

    const sd = row[col("Start Date")];
    const month = sd instanceof Date
      ? Utilities.formatDate(sd, "Asia/Kolkata", "yyyy-MM")
      : "";
    const monthLabel = sd instanceof Date
      ? Utilities.formatDate(sd, "Asia/Kolkata", "MMMM yyyy")
      : "Unassigned";

    tasks.push({
      id: i,
      task: task,
      assignee: assignee,
      priority: String(row[col("Priority")] || "Medium").trim(),
      start: formatDate(row[col("Start Date")]),
      end: formatDate(row[col("End Date")]),
      status: String(row[col("Status")] || "Todo").trim(),
      estHrs: row[col("Est. Hrs")] || null,
      actHrs: row[col("Act. Hrs")] || null,
      jira: String(row[col("Jira")] || "").trim(),
      notes: String(row[col("Notes")] || "").trim(),
      month: month,
      monthLabel: monthLabel,
      ratings: getRatings(i)
    });
  }
  return { tasks: tasks, total: tasks.length };
}

// ── GET SAVED RATINGS FOR A ROW ───────────────────────────────
function getRatings(rowId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("KRA_Ratings");
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == rowId) {
      const r = {};
      for (let k = 1; k <= 7; k++) r[k] = data[i][k] || 0;
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
      "RowID","KPI1","KPI2","KPI3","KPI4","KPI5","KPI6","KPI7",
      "Score","Note","Assignee","Task","Month","Timestamp"
    ]);
    // Format header row
    sh.getRange(1, 1, 1, 14).setFontWeight("bold")
      .setBackground("#4a4e7a").setFontColor("#ffffff");
  }

  const data = sh.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == body.rowId) { foundRow = i + 1; break; }
  }

  const row = [
    body.rowId,
    body.r1 || 0, body.r2 || 0, body.r3 || 0, body.r4 || 0,
    body.r5 || 0, body.r6 || 0, body.r7 || 0,
    body.score || 0,
    body.note || "",
    body.assignee || "",
    body.task || "",
    body.month || "",
    new Date().toISOString()
  ];

  if (foundRow > 0) {
    sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }

  // Also update KRA_Summary sheet
  updateSummary(body.assignee, body.month);

  return { saved: true, rowId: body.rowId };
}

// ── UPDATE SUMMARY SHEET ──────────────────────────────────────
function updateSummary(assignee, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName("KRA_Summary");
    if (!sh) {
      sh = ss.insertSheet("KRA_Summary");
      sh.appendRow([
        "Member","Month","KPI1_Avg","KPI2_Avg","KPI3_Avg",
        "KPI4_Avg","KPI5_Avg","KPI6_Avg","KPI7_Avg",
        "KRA_Score","Grade","Tasks_Rated","Updated"
      ]);
      sh.getRange(1,1,1,13).setFontWeight("bold")
        .setBackground("#4a4e7a").setFontColor("#ffffff");
    }

    // Recalculate this member+month KRA
    const ratingSh = ss.getSheetByName("KRA_Ratings");
    if (!ratingSh) return;

    const rData = ratingSh.getDataRange().getValues();
    const taskSh = ss.getSheetByName(SHEET_NAME);
    const tData = taskSh.getDataRange().getValues();
    const hIdx = HEADER_ROW - 1;
    const cols = tData[hIdx];
    const col = (n) => cols.indexOf(n);

    // Get all row IDs for this member+month
    const rowIds = [];
    for (let i = hIdx + 1; i < tData.length; i++) {
      const a = String(tData[i][col("Assignee")] || "").trim();
      const sd = tData[i][col("Start Date")];
      const m = sd instanceof Date ? Utilities.formatDate(sd, "Asia/Kolkata", "yyyy-MM") : "";
      if (a === assignee && m === month) rowIds.push(i);
    }

    // Gather ratings for those rows
    const weights = [15, 15, 20, 15, 10, 10, 15];
    const kpiSums = [0,0,0,0,0,0,0];
    let ratedCount = 0;

    rowIds.forEach(rid => {
      for (let i = 1; i < rData.length; i++) {
        if (rData[i][0] == rid) {
          let allRated = true;
          for (let k = 1; k <= 7; k++) { if (!rData[i][k]) allRated = false; }
          if (allRated) {
            ratedCount++;
            for (let k = 0; k < 7; k++) kpiSums[k] += rData[i][k+1] || 0;
          }
          break;
        }
      }
    });

    if (ratedCount === 0) return;

    const kpiAvgs = kpiSums.map(s => Math.round((s / ratedCount) * 10) / 10);
    let kraScore = 0;
    kpiAvgs.forEach((avg, i) => { kraScore += (avg / 5) * weights[i]; });
    kraScore = Math.round(kraScore * 10) / 10;
    const grade = kraScore >= 85 ? "A" : kraScore >= 70 ? "B" : kraScore >= 55 ? "C" : "D";

    // Find existing row or append
    const sData = sh.getDataRange().getValues();
    let sRow = -1;
    for (let i = 1; i < sData.length; i++) {
      if (sData[i][0] === assignee && sData[i][1] === month) { sRow = i + 1; break; }
    }

    const summaryRow = [
      assignee, month,
      kpiAvgs[0], kpiAvgs[1], kpiAvgs[2], kpiAvgs[3],
      kpiAvgs[4], kpiAvgs[5], kpiAvgs[6],
      kraScore, grade, ratedCount,
      new Date().toISOString()
    ];

    if (sRow > 0) sh.getRange(sRow, 1, 1, summaryRow.length).setValues([summaryRow]);
    else sh.appendRow(summaryRow);
  } catch(e) {
    // Summary update is best-effort
  }
}

// ── ADD NEW TASK ──────────────────────────────────────────────
function addTask(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { error: "Sheet not found" };
  sh.appendRow([
    "", body.task || "",
    body.start || "", body.end || "", "",
    body.assignee || "", body.priority || "Medium",
    body.estHrs || "", "", "",
    body.status || "Todo",
    body.jira || "", body.notes || ""
  ]);
  return { added: true };
}

// ── DATE FORMATTER ────────────────────────────────────────────
function formatDate(v) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd");
  return String(v).substring(0, 10);
}
