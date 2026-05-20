// ╔══════════════════════════════════════════════════════════════╗
// ║  UFFIZIO KRA BRIDGE — Final v4                              ║
// ║  1. Replace YOUR_SHEET_ID_HERE with your Sheet ID           ║
// ║  2. Save → Deploy → New deployment → Web app → Anyone      ║
// ╚══════════════════════════════════════════════════════════════╝

var SPREADSHEET_ID = "1XZ32kMV32ASmgxXTF26sIPIoqRtB8ADOQ4j1__9sabM";
var SHEET_NAME     = "All Tasks";
var HEADER_ROW     = 3;

// ── Helpers ───────────────────────────────────────────────────
var MMAP = {jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",
            jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};
var MLONG = ["January","February","March","April","May","June",
             "July","August","September","October","November","December"];

function pad(n){ return n<10?"0"+n:""+n; }

function fmtDate(v){
  if(!v && v!==0) return "";
  // Date object
  if(v instanceof Date){
    var y=v.getFullYear(), mo=v.getMonth()+1, d=v.getDate();
    if(y<2020||y>2035){ y=v.getUTCFullYear(); mo=v.getUTCMonth()+1; d=v.getUTCDate(); }
    if(y>=2020&&y<=2035) return y+"-"+pad(mo)+"-"+pad(d);
  }
  // Excel serial
  if(typeof v==="number"){
    var dt=new Date(new Date(1899,11,30).getTime()+v*86400000);
    if(dt.getFullYear()>=2020&&dt.getFullYear()<=2035)
      return dt.getFullYear()+"-"+pad(dt.getMonth()+1)+"-"+pad(dt.getDate());
  }
  var s=String(v).trim();
  // YYYY-MM-DD
  if(/^\d{4}-\d{2}-\d{2}/.test(s)){
    var y2=parseInt(s.substring(0,4));
    if(y2>=2020&&y2<=2035) return s.substring(0,10);
    if(y2<2020) return "2026"+s.substring(4,10);
  }
  // DD-MMM-YY or DD-MMM-YYYY  e.g. "19-May-26"
  var m1=s.match(/^(\d{1,2})[\-\/]([A-Za-z]{3})[\-\/](\d{2,4})$/);
  if(m1){
    var mn=MMAP[m1[2].toLowerCase()];
    if(mn){ var y3=m1[3].length===2?"20"+m1[3]:m1[3]; if(parseInt(y3)>=2020) return y3+"-"+mn+"-"+pad(parseInt(m1[1])); }
  }
  // DD/MM/YYYY
  var m2=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if(m2&&parseInt(m2[3])>=2020) return m2[3]+"-"+pad(parseInt(m2[2]))+"-"+pad(parseInt(m2[1]));
  return "";
}

function monthLabel(mo){
  try{ var p=mo.split("-"); return MLONG[parseInt(p[1])-1]+" "+p[0]; }catch(e){ return mo; }
}

// ── GET handler ───────────────────────────────────────────────
function doGet(e){
  var p=e.parameter||{}, action=p.action||"getTasks", cb=p.callback||null;
  var result;
  try{
    if(action==="getTasks")       result=getTasks();
    else if(action==="saveRatings") result=saveRatingsGET(p);
    else if(action==="addTask")     result=addTaskGET(p);
    else result={error:"Unknown action"};
  }catch(err){ result={error:err.toString()}; }
  var json=JSON.stringify(result);
  if(cb) return ContentService.createTextOutput(cb+"("+json+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  var body={};
  try{ body=JSON.parse(e.postData.contents); }catch(err){}
  var result;
  try{
    if(body.action==="saveRatings") result=saveRatings(body);
    else if(body.action==="addTask") result=addTask(body);
    else result={error:"Unknown action"};
  }catch(err){ result={error:err.toString()}; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ── GET ALL TASKS ─────────────────────────────────────────────
function getTasks(){
  var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh=ss.getSheetByName(SHEET_NAME);
  if(!sh) return {error:"Sheet '"+SHEET_NAME+"' not found"};

  var data=sh.getDataRange().getValues();
  var hIdx=HEADER_ROW-1;
  var cols=data[hIdx];

  // Find column index by name
  function col(name){
    for(var i=0;i<cols.length;i++) if(String(cols[i]).trim()===name) return i;
    return -1;
  }

  var members=["Harshil","Hinesh","Mansi","Vishal","Mayur","Kinjal"];
  var tasks=[];

  for(var i=hIdx+1; i<data.length; i++){
    var row=data[i];
    var assignee=col("Assignee")>=0 ? String(row[col("Assignee")]||"").trim() : "";
    var task=col("Task")>=0 ? String(row[col("Task")]||"").trim() : "";

    // Skip if no assignee or assignee not in team
    if(!assignee || !members.includes(assignee)) continue;
    // Skip if no task name
    if(!task || task==="") continue;

    var startRaw=col("Start Date")>=0 ? row[col("Start Date")] : "";
    var endRaw=col("End Date")>=0 ? row[col("End Date")] : "";
    var startStr=fmtDate(startRaw);
    var endStr=fmtDate(endRaw);

    // Use end date as fallback if start date is missing
    var dateForMonth=startStr||endStr;
    var month=dateForMonth ? dateForMonth.substring(0,7) : "";

    // Last resort: if still no month, use current month
    if(!month){
      var now=new Date();
      month=now.getFullYear()+"-"+pad(now.getMonth()+1);
    }

    tasks.push({
      id:         i,
      task:       task,
      assignee:   assignee,
      priority:   col("Priority")>=0 ? (String(row[col("Priority")]||"").trim()||"Medium") : "Medium",
      start:      startStr || endStr, // use end as fallback
      end:        endStr,
      status:     col("Status")>=0 ? (String(row[col("Status")]||"").trim()||"Todo") : "Todo",
      jira:       col("Jira")>=0 ? String(row[col("Jira")]||"").trim() : "",
      notes:      col("Notes")>=0 ? String(row[col("Notes")]||"").trim() : "",
      month:      month,
      monthLabel: monthLabel(month),
      ratings:    {}
    });
  }
  return {tasks:tasks, total:tasks.length};
}

// ── SAVE RATINGS ──────────────────────────────────────────────
function saveRatingsGET(p){
  return saveRatings({
    rowId:p.rowId, score:parseFloat(p.score)||0,
    note:decodeURIComponent(p.note||""),
    assignee:decodeURIComponent(p.assignee||""),
    task:decodeURIComponent(p.task||""),
    month:p.month||"",
    r1:decodeURIComponent(p.r1||""),r2:decodeURIComponent(p.r2||""),
    r3:decodeURIComponent(p.r3||""),r4:decodeURIComponent(p.r4||""),
    r5:decodeURIComponent(p.r5||""),r6:decodeURIComponent(p.r6||""),
    r7:decodeURIComponent(p.r7||""),r8:decodeURIComponent(p.r8||""),
    r9:decodeURIComponent(p.r9||""),r10:decodeURIComponent(p.r10||"")
  });
}

function saveRatings(body){
  var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh=ss.getSheetByName("KRA_Ratings");
  if(!sh){
    sh=ss.insertSheet("KRA_Ratings");
    var hdr=["Row ID","KPI1","KPI2","KPI3","KPI4","KPI5",
             "KPI6","KPI7","KPI8","KPI9","KPI10",
             "Total Score","Note","Assignee","Task","Month","Saved At"];
    sh.appendRow(hdr);
    sh.getRange(1,1,1,hdr.length).setFontWeight("bold").setBackground("#534AB7").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }
  var data=sh.getDataRange().getValues();
  var foundRow=-1;
  for(var i=1;i<data.length;i++) if(String(data[i][0])===String(body.rowId)){foundRow=i+1;break;}
  var row=[body.rowId,
    body.r1||"",body.r2||"",body.r3||"",body.r4||"",body.r5||"",
    body.r6||"",body.r7||"",body.r8||"",body.r9||"",body.r10||"",
    body.score||0,body.note||"",body.assignee||"",body.task||"",body.month||"",
    new Date().toISOString()];
  if(foundRow>0) sh.getRange(foundRow,1,1,row.length).setValues([row]);
  else sh.appendRow(row);
  return {saved:true};
}

// ── ADD TASK ──────────────────────────────────────────────────
function addTaskGET(p){
  return addTask({
    task:decodeURIComponent(p.task||""),assignee:decodeURIComponent(p.assignee||""),
    priority:p.priority||"Medium",start:p.start||"",end:p.end||"",
    status:p.status||"Todo",jira:decodeURIComponent(p.jira||"")
  });
}

function addTask(body){
  var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh=ss.getSheetByName(SHEET_NAME);
  if(!sh) return {error:"Sheet not found"};
  sh.appendRow(["",body.task||"",body.start||"",body.end||"","",
    body.assignee||"",body.priority||"Medium","","","",body.status||"Todo",body.jira||",",""]);
  return {added:true};
}
