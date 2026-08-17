/**
 * Backend สำหรับ Web ติดตามงานภายใน — วางโค้ดนี้ใน Google Apps Script
 * (Extensions > Apps Script ในไฟล์ Google Sheet ของคุณ) แล้ว Deploy เป็น Web App
 *
 * วิธี Deploy:
 * 1. เปิด Google Sheet ที่ใช้เป็นฐานข้อมูล
 * 2. Extensions > Apps Script
 * 3. วางโค้ดนี้ทับของเดิม แล้วกด Save
 * 4. Deploy > Manage deployments > แก้ deployment เดิม (หรือ New deployment ถ้ายังไม่มี)
 *    - Execute as: Me
 *    - Who has access: Anyone (หรือ Anyone within organization ถ้าใช้ Google Workspace)
 * 5. คัดลอก URL ที่ได้ (ลงท้ายด้วย /exec) ไปใส่ในหน้าเว็บ (ตัวแปร API_URL ใน src/config.js)
 *
 * Sheet "Log" จะถูกสร้างขึ้นอัตโนมัติครั้งแรกที่มีการแก้ไขข้อมูลผ่านเว็บ ไม่ต้องสร้างมือ
 */

const SHEET_OVERVIEW = "Overview";
const SHEET_LOG = "Log";
const TASK_HEADERS = ["Type", "Start", "Due", "Status", "Progress", "Task Name", "Link", "Task Detail"];

function doGet(e) {
  const action = e.parameter.action || "overview";
  if (action === "overview") return respond(getOverview());
  if (action === "tasks") return respond(getTasks(e.parameter.projectId));
  if (action === "alltasks") return respond(getAllTasks());
  if (action === "log") return respond(getLog(Number(e.parameter.limit) || 200));
  return respond({ error: "unknown action" });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const user = body.user || "ไม่ระบุ";
  let result;

  switch (body.action) {
    case "updateTaskStatus":
      result = updateTaskStatus(body.projectId, body.rowIndex, body.status, body.progress);
      if (!result.error) logAction(user, "เปลี่ยนสถานะงาน", body.projectId, "แถว " + body.rowIndex + " → " + body.status);
      break;
    case "addTask":
      result = addTask(body.projectId, body.task);
      if (!result.error) logAction(user, "เพิ่มงาน", body.projectId, body.task["Task Name"] || "");
      break;
    case "editTask":
      result = editTask(body.projectId, body.rowIndex, body.fields);
      if (!result.error) logAction(user, "แก้ไขงาน", body.projectId, body.fields["Task Name"] || ("แถว " + body.rowIndex));
      break;
    case "deleteTask":
      result = deleteTask(body.projectId, body.rowIndex);
      if (!result.error) logAction(user, "ลบงาน", body.projectId, "แถว " + body.rowIndex);
      break;
    case "addProject":
      result = addProject(body.project);
      if (!result.error) logAction(user, "เพิ่มโปรเจกต์", result.projectId, body.project.name || "");
      break;
    case "editProjectDetail":
      result = editProjectDetail(body.projectId, body.fields);
      if (!result.error) logAction(user, "แก้ไขรายละเอียดโปรเจกต์", body.projectId, Object.keys(body.fields || {}).join(", "));
      break;
    default:
      return respond({ error: "unknown action" });
  }
  return respond(result);
}

// ---------------------------------------------------------------------------
// Overview / Projects
// ---------------------------------------------------------------------------

function getOverview() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_OVERVIEW);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map((row) => rowToObject(headers, row));
}

/**
 * project: { name, pm, sales, start, due, platform }
 * สร้างรหัสโปรเจกต์อัตโนมัติ รูปแบบ PM-YYMM## เช่น PM-260801
 * (YY = 2 ตัวท้ายปี ค.ศ., MM = เดือน 2 หลัก, ## = ลำดับโปรเจกต์ในเดือนนั้น เริ่ม 01)
 * แล้วสร้างชีตใหม่สำหรับเก็บ task ของโปรเจกต์นี้ด้วย
 */
function addProject(project) {
  const overview = SpreadsheetApp.getActive().getSheetByName(SHEET_OVERVIEW);
  const headers = overview.getRange(1, 1, 1, overview.getLastColumn()).getValues()[0];
  const projectId = generateProjectId(overview, headers);

  const startDate = project.start ? new Date(project.start) : "";
  const dueDate = project.due ? new Date(project.due) : "";

  const fieldMap = {
    "Project ID": projectId,
    "Project name": project.name || "",
    Type: project.type || "",
    PM: project.pm || "",
    "Sales Owner": project.sales || "",
    "Start Date": startDate,
    "Due Date": dueDate,
    Platform: project.platform || "",
    "Task progress": 0,
  };
  const row = headers.map((h) => (fieldMap[h] !== undefined ? fieldMap[h] : ""));
  overview.appendRow(row);

  const newRowIndex = overview.getLastRow();
  const numCols = headers.length;

  // 1.2 ตีกรอบล้อมทั้งแถวใหม่ (คอลัมน์ A ถึงคอลัมน์สุดท้าย = Task progress = X)
  overview.getRange(newRowIndex, 1, 1, numCols).setBorder(true, true, true, true, true, true);

  // 1.3 คอลัมน์ A (Project ID) จัดกึ่งกลาง
  const idCol = headers.indexOf("Project ID") + 1;
  if (idCol > 0) overview.getRange(newRowIndex, idCol).setHorizontalAlignment("center");

  // 1.4 คอลัมน์ B/C (Start Date, Due Date) ฟอร์แมตวันที่ d/M/yyyy เช่น 31/7/2026 และจัดกึ่งกลาง
  const startCol = headers.indexOf("Start Date") + 1;
  const dueCol = headers.indexOf("Due Date") + 1;
  if (startCol > 0) overview.getRange(newRowIndex, startCol).setNumberFormat("d/M/yyyy").setHorizontalAlignment("center");
  if (dueCol > 0) overview.getRange(newRowIndex, dueCol).setNumberFormat("d/M/yyyy").setHorizontalAlignment("center");

  // 1.1 สร้างชีตสำหรับเก็บ task ของโปรเจกต์นี้ ให้หน้าตาตรงกับชีตโปรเจกต์เดิมๆ ทุกอย่าง
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(projectId)) {
    const taskSheet = ss.insertSheet(projectId);
    taskSheet.appendRow(TASK_HEADERS);
    taskSheet.getRange(1, 1, 1, TASK_HEADERS.length).setFontWeight("bold");
    taskSheet.setFrozenRows(1);

    // ความกว้างคอลัมน์ให้ตรงกับ template เดิม (Type, Start, Due, Status, Progress, Task Name, Link, Task Detail)
    taskSheet.setColumnWidth(1, 70); // Type
    taskSheet.setColumnWidth(2, 170); // Start
    taskSheet.setColumnWidth(4, 80); // Status
    taskSheet.setColumnWidth(5, 170); // Progress
    taskSheet.setColumnWidth(6, 348); // Task Name
    taskSheet.setColumnWidth(7, 301); // Link
    taskSheet.setColumnWidth(8, 356); // Task Detail

    // ตั้งฟอร์แมตล่วงหน้าให้ตรงกับชีตเดิม: วันที่แบบ d-mmm, Progress แบบ 0%
    const PRE_FORMAT_ROWS = 200;
    taskSheet.getRange(2, 2, PRE_FORMAT_ROWS, 1).setNumberFormat("d-mmm"); // Start
    taskSheet.getRange(2, 3, PRE_FORMAT_ROWS, 1).setNumberFormat("d-mmm"); // Due
    taskSheet.getRange(2, 5, PRE_FORMAT_ROWS, 1).setNumberFormat("0%"); // Progress

    // ย้ายชีตใหม่ไปไว้ท้ายสุด เพื่อความเป็นระเบียบ
    ss.setActiveSheet(taskSheet);
    ss.moveActiveSheet(ss.getNumSheets());
  }

  return { ok: true, projectId: projectId };
}

/**
 * fields: partial object keyed by exact sheet column names to update, e.g.
 * { "Project name": "New name", "Project Value": 100000, "Sale brief": true, "Start Date": "2026-08-01" }
 * ใช้แก้ไขข้อมูลในชีต Overview ของโปรเจกต์ที่มี Project ID ตรงกับ projectId
 */
function editProjectDetail(projectId, fields) {
  const overview = SpreadsheetApp.getActive().getSheetByName(SHEET_OVERVIEW);
  const headers = overview.getRange(1, 1, 1, overview.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf("Project ID") + 1;
  if (idCol === 0) return { error: "ไม่พบคอลัมน์ Project ID ใน Overview" };

  const lastRow = overview.getLastRow();
  if (lastRow < 2) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const ids = overview.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
  const offset = ids.indexOf(projectId);
  if (offset === -1) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const rowIndex = offset + 2;

  Object.keys(fields).forEach((key) => {
    const col = headers.indexOf(key) + 1;
    if (col === 0) return;
    let value = fields[key];
    const isDateField = key === "Start Date" || key === "Due Date";
    if (isDateField && value) value = new Date(value);
    overview.getRange(rowIndex, col).setValue(value);
    if (isDateField) {
      overview.getRange(rowIndex, col).setNumberFormat("d/M/yyyy").setHorizontalAlignment("center");
    }
  });

  return { ok: true };
}

function generateProjectId(overview, headers) {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(-2);
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const prefix = "PM-" + yy + mm;

  const idCol = headers.indexOf("Project ID");
  const lastRow = overview.getLastRow();
  let maxSeq = 0;
  if (lastRow > 1 && idCol > -1) {
    const values = overview.getRange(2, idCol + 1, lastRow - 1, 1).getValues().flat();
    values.forEach((v) => {
      const s = String(v);
      if (s.indexOf(prefix) === 0) {
        const seq = parseInt(s.slice(prefix.length), 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    });
  }
  return prefix + String(maxSeq + 1).padStart(2, "0");
}

// ---------------------------------------------------------------------------
// Tasks
// ---------------------------------------------------------------------------

function getTasks(projectId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(projectId);
  if (!sheet) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map((row, i) => Object.assign(rowToObject(headers, row), { rowIndex: i + 2 }));
}

/**
 * ดึง task ของทุกโปรเจกต์รวดเดียว (ใช้กับฟีเจอร์ export PowerPoint)
 * คืนค่าเป็น object: { "PM-260601": [ {Type,...}, ... ], ... }
 */
function getAllTasks() {
  const overview = SpreadsheetApp.getActive().getSheetByName(SHEET_OVERVIEW);
  const headers = overview.getRange(1, 1, 1, overview.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf("Project ID") + 1;
  const lastRow = overview.getLastRow();
  const result = {};
  if (lastRow < 2 || idCol === 0) return result;

  const ids = overview.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
  const ss = SpreadsheetApp.getActive();
  ids.forEach((id) => {
    if (!id) return;
    const sheet = ss.getSheetByName(id);
    if (!sheet) {
      result[id] = [];
      return;
    }
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      result[id] = [];
      return;
    }
    const taskHeaders = values[0];
    result[id] = values
      .slice(1)
      .filter((row) => row.some((c) => c !== "" && c !== null))
      .map((row) => rowToObject(taskHeaders, row));
  });
  return result;
}

function addTask(projectId, task) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(projectId);
  if (!sheet) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map((h) => (task[h] !== undefined ? task[h] : ""));
  sheet.appendRow(row);
  return { ok: true, rowIndex: sheet.getLastRow() };
}

function editTask(projectId, rowIndex, fields) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(projectId);
  if (!sheet) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Object.keys(fields).forEach((key) => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(rowIndex, col).setValue(fields[key]);
  });
  return { ok: true };
}

function deleteTask(projectId, rowIndex) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(projectId);
  if (!sheet) return { error: "ไม่พบโปรเจกต์ " + projectId };
  sheet.deleteRow(rowIndex);
  return { ok: true };
}

function updateTaskStatus(projectId, rowIndex, status, progress) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(projectId);
  if (!sheet) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf("Status") + 1;
  const progressCol = headers.indexOf("Progress") + 1;
  if (statusCol > 0) sheet.getRange(rowIndex, statusCol).setValue(status);
  if (progressCol > 0) sheet.getRange(rowIndex, progressCol).setValue(progress);
  return { ok: true };
}

// ---------------------------------------------------------------------------
// Log — เนื่องจากระบบไม่มี login จริง เก็บแค่ เวลา/ผู้ใช้ที่เลือก/การกระทำ
// ---------------------------------------------------------------------------

function logAction(user, action, projectId, detail) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.appendRow(["Timestamp", "User", "Action", "Project", "Detail"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  sheet.appendRow([new Date(), user, action, projectId || "", detail || ""]);
}

function getLog(limit) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1).map((row) => rowToObject(headers, row));
  rows.reverse();
  return limit ? rows.slice(0, limit) : rows;
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, i) => (obj[h] = row[i]));
  return obj;
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Repair utilities — รันครั้งเดียวจาก Apps Script editor (ไม่ได้เชื่อมกับเว็บ)
// ใช้แก้ชีตที่สร้างไว้ก่อนโค้ดเวอร์ชันที่มีการจัดฟอร์แมตอัตโนมัติ
// วิธีรัน: เปิดไฟล์นี้ใน Apps Script editor > เลือกชื่อฟังก์ชันจาก dropdown ด้านบน > กด Run
// ---------------------------------------------------------------------------

function repairAllTaskSheetFormatting() {
  const ss = SpreadsheetApp.getActive();
  const overview = ss.getSheetByName(SHEET_OVERVIEW);
  const headers = overview.getRange(1, 1, 1, overview.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf("Project ID") + 1;
  const lastRow = overview.getLastRow();
  if (lastRow < 2) return;
  const ids = overview.getRange(2, idCol, lastRow - 1, 1).getValues().flat();

  let fixedCount = 0;
  ids.forEach((id) => {
    if (!id) return;
    const sheet = ss.getSheetByName(id);
    if (!sheet) return;

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 70); // Type
    sheet.setColumnWidth(2, 170); // Start
    sheet.setColumnWidth(4, 80); // Status
    sheet.setColumnWidth(5, 170); // Progress
    sheet.setColumnWidth(6, 348); // Task Name
    sheet.setColumnWidth(7, 301); // Link
    sheet.setColumnWidth(8, 356); // Task Detail

    const rowsToFormat = Math.max(sheet.getLastRow() + 20, 200);
    sheet.getRange(2, 2, rowsToFormat, 1).setNumberFormat("d-mmm"); // Start
    sheet.getRange(2, 3, rowsToFormat, 1).setNumberFormat("d-mmm"); // Due
    sheet.getRange(2, 5, rowsToFormat, 1).setNumberFormat("0%"); // Progress
    fixedCount++;
  });

  Logger.log("ซ่อมฟอร์แมตชีตงานแล้ว " + fixedCount + " ชีต");
}

function repairOverviewFormatting() {
  const overview = SpreadsheetApp.getActive().getSheetByName(SHEET_OVERVIEW);
  const headers = overview.getRange(1, 1, 1, overview.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf("Project ID") + 1;
  const startCol = headers.indexOf("Start Date") + 1;
  const dueCol = headers.indexOf("Due Date") + 1;
  const lastRow = overview.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  if (idCol > 0) overview.getRange(2, idCol, numRows, 1).setHorizontalAlignment("center");
  if (startCol > 0) overview.getRange(2, startCol, numRows, 1).setNumberFormat("d/M/yyyy").setHorizontalAlignment("center");
  if (dueCol > 0) overview.getRange(2, dueCol, numRows, 1).setNumberFormat("d/M/yyyy").setHorizontalAlignment("center");

  Logger.log("ซ่อมฟอร์แมต Overview แล้ว " + numRows + " แถว");
}
