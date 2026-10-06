// พอร์ตมาจาก apps_script_backend.gs เดิม — logic เหมือนเดิมทุกจุด เปลี่ยนแค่วิธีคุยกับ
// Google Sheets จาก SpreadsheetApp (Apps Script) มาเป็น Sheets API v4 ตรงๆ (sheets.js)

import sheets from "./sheets.js";

const SHEET_OVERVIEW = "Overview";
const SHEET_LOG = "Log";
const TASK_HEADERS = ["Type", "Start", "Due", "Status", "Progress", "Task Name", "Link", "Task Detail"];
const DATE_HEADERS_OVERVIEW = ["Start Date", "Due Date"];
const DATE_HEADERS_TASK = ["Start", "Due"];

// ---------------------------------------------------------------------------
// วันที่: Sheets API แบบ UNFORMATTED_VALUE คืนวันที่เป็นเลข serial (นับวันจาก 30 ธ.ค. 1899)
// ไม่ใช่ Date object แบบที่ Apps Script คืนให้ ต้องแปลงเองเฉพาะคอลัมน์ที่รู้ว่าเป็นวันที่
// ---------------------------------------------------------------------------
const SERIAL_EPOCH_MS = Date.UTC(1899, 11, 30);

function serialToISO(serial) {
  const n = Number(serial);
  if (isNaN(n)) return "";
  return new Date(SERIAL_EPOCH_MS + Math.round(n * 86400000)).toISOString();
}

function rowToObject(headers, row, dateHeaders) {
  const dh = dateHeaders || [];
  const obj = {};
  headers.forEach((h, i) => {
    let v = row[i];
    if (v === undefined || v === null) v = "";
    if (dh.includes(h) && v !== "" && typeof v === "number") v = serialToISO(v);
    obj[h] = v;
  });
  return obj;
}

// ---------------------------------------------------------------------------
// Overview / Projects
// ---------------------------------------------------------------------------

async function getHeaders(sheetTitle) {
  const rows = await sheets.getValues(`${sheets.quoteSheetName(sheetTitle)}!1:1`);
  return rows[0] || [];
}

/**
 * หาแถวสุดท้ายที่มี "Project ID" จริงๆ ไม่ใช่แค่แถวสุดท้ายที่มีอะไรก็ได้ (เช่นสูตร Task
 * progress ที่ลากไว้ล่วงหน้าเป็นร้อยแถวโดยไม่มีข้อมูลโปรเจกต์จริง)
 */
async function getLastRealDataRow(sheetTitle, idColIndex0) {
  const colL = sheets.colLetter(idColIndex0 + 1);
  const values = await sheets.getValues(`${sheets.quoteSheetName(sheetTitle)}!${colL}2:${colL}`);
  let lastReal = 1; // แถวหัวตาราง
  values.forEach((r, i) => {
    const v = r[0];
    if (v !== undefined && v !== "" && v !== null) lastReal = i + 2;
  });
  return lastReal;
}

async function getOverview() {
  const headers = await getHeaders(SHEET_OVERVIEW);
  const idColIndex0 = headers.indexOf("Project ID");
  if (idColIndex0 === -1) return [];
  const lastRow = await getLastRealDataRow(SHEET_OVERVIEW, idColIndex0);
  if (lastRow < 2) return [];
  const dataRange = `${sheets.quoteSheetName(SHEET_OVERVIEW)}!A2:${sheets.colLetter(headers.length)}${lastRow}`;
  const rows = await sheets.getValues(dataRange);
  return rows.map((row) => rowToObject(headers, row, DATE_HEADERS_OVERVIEW));
}

async function generateProjectId(headers) {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(-2);
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const prefix = "PM-" + yy + mm;

  const idColIndex0 = headers.indexOf("Project ID");
  let maxSeq = 0;
  if (idColIndex0 > -1) {
    const colL = sheets.colLetter(idColIndex0 + 1);
    const values = await sheets.getValues(`${sheets.quoteSheetName(SHEET_OVERVIEW)}!${colL}2:${colL}`);
    values.forEach((r) => {
      const s = String(r[0] || "");
      if (s.indexOf(prefix) === 0) {
        const seq = parseInt(s.slice(prefix.length), 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    });
  }
  return prefix + String(maxSeq + 1).padStart(2, "0");
}

async function createTaskSheet(projectId) {
  const meta = await sheets.getSpreadsheetMeta();
  const totalSheets = (meta.sheets || []).length;

  const addResult = await sheets.batchUpdate([{ addSheet: { properties: { title: projectId } } }]);
  const newSheetId = addResult.replies[0].addSheet.properties.sheetId;

  await sheets.updateValues(`${sheets.quoteSheetName(projectId)}!A1:H1`, [TASK_HEADERS]);

  const PRE_FORMAT_ROWS = 200;
  const colWidths = [
    [0, 70], // Type
    [1, 170], // Start
    [3, 80], // Status
    [4, 170], // Progress
    [5, 348], // Task Name
    [6, 301], // Link
    [7, 356], // Task Detail
  ];

  await sheets.batchUpdate([
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: TASK_HEADERS.length },
        cell: { userEnteredFormat: { textFormat: { bold: true } } },
        fields: "userEnteredFormat.textFormat.bold",
      },
    },
    {
      updateSheetProperties: {
        properties: { sheetId: newSheetId, gridProperties: { frozenRowCount: 1 } },
        fields: "gridProperties.frozenRowCount",
      },
    },
    ...colWidths.map(([col0, px]) => ({
      updateDimensionProperties: {
        range: { sheetId: newSheetId, dimension: "COLUMNS", startIndex: col0, endIndex: col0 + 1 },
        properties: { pixelSize: px },
        fields: "pixelSize",
      },
    })),
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: 1, endRowIndex: 1 + PRE_FORMAT_ROWS, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { numberFormat: { type: "DATE", pattern: "d-mmm" } } },
        fields: "userEnteredFormat.numberFormat",
      },
    },
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: 1, endRowIndex: 1 + PRE_FORMAT_ROWS, startColumnIndex: 2, endColumnIndex: 3 },
        cell: { userEnteredFormat: { numberFormat: { type: "DATE", pattern: "d-mmm" } } },
        fields: "userEnteredFormat.numberFormat",
      },
    },
    {
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: 1, endRowIndex: 1 + PRE_FORMAT_ROWS, startColumnIndex: 4, endColumnIndex: 5 },
        cell: { userEnteredFormat: { numberFormat: { type: "PERCENT", pattern: "0%" } } },
        fields: "userEnteredFormat.numberFormat",
      },
    },
    {
      updateSheetProperties: {
        properties: { sheetId: newSheetId, index: totalSheets },
        fields: "index",
      },
    },
  ]);

  return newSheetId;
}

async function addProject(project) {
  const headers = await getHeaders(SHEET_OVERVIEW);
  const idColIndex0 = headers.indexOf("Project ID");
  const projectId = await generateProjectId(headers);

  const fieldMap = {
    "Project ID": projectId,
    "Project name": project.name || "",
    Type: project.type || "",
    PM: project.pm || "",
    "Sales Owner": project.sales || "",
    "Start Date": project.start || "",
    "Due Date": project.due || "",
    Platform: project.platform || "",
    "Task progress": 0,
  };
  const row = headers.map((h) => (fieldMap[h] !== undefined ? fieldMap[h] : ""));

  const lastDataRow = await getLastRealDataRow(SHEET_OVERVIEW, idColIndex0);
  const overviewSheetId = await sheets.getSheetIdByTitle(SHEET_OVERVIEW);
  const insertAt0 = lastDataRow; // 0-indexed ตำแหน่งแทรก = 1-indexed lastDataRow (แทรกต่อจากแถวนั้น)
  const newRowIndex1 = lastDataRow + 1;

  // 1) แทรกแถวว่าง โดยสืบทอดฟอร์แมต/สูตรจากแถวด้านบน (รวมสูตร Task progress ด้วย)
  await sheets.batchUpdate([
    {
      insertDimension: {
        range: { sheetId: overviewSheetId, dimension: "ROWS", startIndex: insertAt0, endIndex: insertAt0 + 1 },
        inheritFromBefore: true,
      },
    },
  ]);

  // 2) เขียนค่าลงแถวใหม่
  const dataRange = `${sheets.quoteSheetName(SHEET_OVERVIEW)}!A${newRowIndex1}:${sheets.colLetter(headers.length)}${newRowIndex1}`;
  await sheets.updateValues(dataRange, [row]);

  // 3) จัดฟอร์แมต: กรอบรอบแถว, จัดกึ่งกลาง Project ID, ฟอร์แมตวันที่ Start/Due Date
  const startCol0 = headers.indexOf("Start Date");
  const dueCol0 = headers.indexOf("Due Date");
  const formatRequests = [
    {
      updateBorders: {
        range: { sheetId: overviewSheetId, startRowIndex: newRowIndex1 - 1, endRowIndex: newRowIndex1, startColumnIndex: 0, endColumnIndex: headers.length },
        top: { style: "SOLID" },
        bottom: { style: "SOLID" },
        left: { style: "SOLID" },
        right: { style: "SOLID" },
        innerHorizontal: { style: "SOLID" },
        innerVertical: { style: "SOLID" },
      },
    },
  ];
  if (idColIndex0 > -1) {
    formatRequests.push({
      repeatCell: {
        range: { sheetId: overviewSheetId, startRowIndex: newRowIndex1 - 1, endRowIndex: newRowIndex1, startColumnIndex: idColIndex0, endColumnIndex: idColIndex0 + 1 },
        cell: { userEnteredFormat: { horizontalAlignment: "CENTER" } },
        fields: "userEnteredFormat.horizontalAlignment",
      },
    });
  }
  [startCol0, dueCol0].forEach((col0) => {
    if (col0 > -1) {
      formatRequests.push({
        repeatCell: {
          range: { sheetId: overviewSheetId, startRowIndex: newRowIndex1 - 1, endRowIndex: newRowIndex1, startColumnIndex: col0, endColumnIndex: col0 + 1 },
          cell: { userEnteredFormat: { numberFormat: { type: "DATE", pattern: "d/M/yyyy" }, horizontalAlignment: "CENTER" } },
          fields: "userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment",
        },
      });
    }
  });
  await sheets.batchUpdate(formatRequests);

  // 4) สร้างชีตสำหรับเก็บ task ของโปรเจกต์นี้ (ถ้ายังไม่มี)
  const existingSheetId = await sheets.getSheetIdByTitle(projectId);
  if (existingSheetId === null) {
    await createTaskSheet(projectId);
  }

  return { ok: true, projectId };
}

async function findProjectRow(projectId) {
  const headers = await getHeaders(SHEET_OVERVIEW);
  const idColIndex0 = headers.indexOf("Project ID");
  if (idColIndex0 === -1) return { error: "ไม่พบคอลัมน์ Project ID ใน Overview" };
  const colL = sheets.colLetter(idColIndex0 + 1);
  const idValues = await sheets.getValues(`${sheets.quoteSheetName(SHEET_OVERVIEW)}!${colL}2:${colL}`);
  const offset = idValues.findIndex((r) => r[0] === projectId);
  if (offset === -1) return { error: "ไม่พบโปรเจกต์ " + projectId };
  return { headers, rowIndex1: offset + 2 };
}

async function editProjectDetail(projectId, fields) {
  const found = await findProjectRow(projectId);
  if (found.error) return { error: found.error };
  const { headers, rowIndex1 } = found;

  const overviewSheetId = await sheets.getSheetIdByTitle(SHEET_OVERVIEW);
  const dateFields = ["Start Date", "Due Date"];
  const data = [];
  const formatRequests = [];

  Object.keys(fields).forEach((key) => {
    const col0 = headers.indexOf(key);
    if (col0 === -1) return;
    const colL = sheets.colLetter(col0 + 1);
    data.push({ range: `${sheets.quoteSheetName(SHEET_OVERVIEW)}!${colL}${rowIndex1}`, values: [[fields[key]]] });
    if (dateFields.includes(key)) {
      formatRequests.push({
        repeatCell: {
          range: { sheetId: overviewSheetId, startRowIndex: rowIndex1 - 1, endRowIndex: rowIndex1, startColumnIndex: col0, endColumnIndex: col0 + 1 },
          cell: { userEnteredFormat: { numberFormat: { type: "DATE", pattern: "d/M/yyyy" }, horizontalAlignment: "CENTER" } },
          fields: "userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment",
        },
      });
    }
  });

  if (data.length) await sheets.batchUpdateValues(data);
  if (formatRequests.length) await sheets.batchUpdate(formatRequests);

  return { ok: true };
}

/**
 * ลบโปรเจกต์ทั้งอัน: ลบแถวใน Overview + ลบชีตเก็บ task ของโปรเจกต์นั้นทิ้งไปด้วย
 * หมายเหตุ: ถ้ามีแถวซ้ำ Project ID เดียวกันมากกว่า 1 แถว จะลบแค่แถวแรกที่เจอ
 */
async function deleteProject(projectId) {
  const found = await findProjectRow(projectId);
  if (found.error) return { error: found.error };
  const { rowIndex1 } = found;

  const overviewSheetId = await sheets.getSheetIdByTitle(SHEET_OVERVIEW);
  const requests = [
    { deleteDimension: { range: { sheetId: overviewSheetId, dimension: "ROWS", startIndex: rowIndex1 - 1, endIndex: rowIndex1 } } },
  ];
  const taskSheetId = await sheets.getSheetIdByTitle(projectId);
  if (taskSheetId !== null) requests.push({ deleteSheet: { sheetId: taskSheetId } });

  await sheets.batchUpdate(requests);
  return { ok: true };
}

// ---------------------------------------------------------------------------
// Tasks
// ---------------------------------------------------------------------------

async function getTasks(projectId) {
  const sheetId = await sheets.getSheetIdByTitle(projectId);
  if (sheetId === null) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const rows = await sheets.getValues(`${sheets.quoteSheetName(projectId)}`);
  if (!rows.length) return [];
  const headers = rows[0];
  return rows.slice(1).map((row, i) => Object.assign(rowToObject(headers, row, DATE_HEADERS_TASK), { rowIndex: i + 2 }));
}

/**
 * ดึง task ของทุกโปรเจกต์รวดเดียว (ใช้กับฟีเจอร์ export PowerPoint) — ยิงคำขอเดียว
 * ด้วย values:batchGet แทนการอ่านทีละชีต เพื่อประหยัดโควต้าและเร็วกว่ามาก
 * คืนค่าเป็น object: { "PM-260601": [ {Type,...}, ... ], ... }
 */
async function getAllTasks() {
  const headers = await getHeaders(SHEET_OVERVIEW);
  const idColIndex0 = headers.indexOf("Project ID");
  const result = {};
  if (idColIndex0 === -1) return result;

  const colL = sheets.colLetter(idColIndex0 + 1);
  const idValues = await sheets.getValues(`${sheets.quoteSheetName(SHEET_OVERVIEW)}!${colL}2:${colL}`);
  // ตัด ID ที่ซ้ำ/ว่างทิ้ง
  const ids = [...new Set(idValues.map((r) => r[0]).filter((id) => id).map(String))];
  if (!ids.length) return result;

  // โปรเจกต์ไหนที่ยังไม่มีชีตงาน ให้ได้รายการว่าง (เหมือน Apps Script เดิม) — ต้องกรองก่อนเรียก
  // values:batchGet เพราะ Google จะปฏิเสธ "ทั้งคำขอ" ถ้ามีชื่อชีตแม้แต่ชื่อเดียวที่ไม่มีอยู่จริง
  ids.forEach((id) => {
    result[id] = [];
  });
  const meta = await sheets.getSpreadsheetMeta();
  const existingTitles = new Set((meta.sheets || []).map((s) => s.properties.title));
  const idsWithSheet = ids.filter((id) => existingTitles.has(id));

  // แบ่งเป็นชุดละ 40 ชีต กัน URL ยาวเกินเมื่อโปรเจกต์เยอะขึ้น (ยังเป็นการเรียกเพียงไม่กี่ครั้ง)
  const CHUNK = 40;
  for (let i = 0; i < idsWithSheet.length; i += CHUNK) {
    const chunk = idsWithSheet.slice(i, i + CHUNK);
    const valueRanges = await sheets.batchGetValues(chunk.map((id) => sheets.quoteSheetName(id)));
    valueRanges.forEach((vr, j) => {
      const id = chunk[j];
      const rows = vr.values || [];
      if (rows.length < 2) return;
      const taskHeaders = rows[0];
      result[id] = rows
        .slice(1)
        .filter((row) => row.some((c) => c !== "" && c !== undefined && c !== null))
        .map((row) => rowToObject(taskHeaders, row, DATE_HEADERS_TASK));
    });
  }
  return result;
}

async function addTask(projectId, task) {
  const sheetId = await sheets.getSheetIdByTitle(projectId);
  if (sheetId === null) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const headers = (await getHeaders(projectId)) || TASK_HEADERS;
  const row = headers.map((h) => (task[h] !== undefined ? task[h] : ""));
  await sheets.appendValues(`${sheets.quoteSheetName(projectId)}!A:${sheets.colLetter(headers.length)}`, [row]);
  return { ok: true };
}

async function editTask(projectId, rowIndex, fields) {
  const sheetId = await sheets.getSheetIdByTitle(projectId);
  if (sheetId === null) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const headers = (await getHeaders(projectId)) || TASK_HEADERS;
  const data = [];
  Object.keys(fields).forEach((key) => {
    const col0 = headers.indexOf(key);
    if (col0 === -1) return;
    const colL = sheets.colLetter(col0 + 1);
    data.push({ range: `${sheets.quoteSheetName(projectId)}!${colL}${rowIndex}`, values: [[fields[key]]] });
  });
  if (data.length) await sheets.batchUpdateValues(data);
  return { ok: true };
}

async function deleteTask(projectId, rowIndex) {
  const sheetId = await sheets.getSheetIdByTitle(projectId);
  if (sheetId === null) return { error: "ไม่พบโปรเจกต์ " + projectId };
  await sheets.batchUpdate([
    { deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: rowIndex - 1, endIndex: rowIndex } } },
  ]);
  return { ok: true };
}

async function updateTaskStatus(projectId, rowIndex, status, progress) {
  const sheetId = await sheets.getSheetIdByTitle(projectId);
  if (sheetId === null) return { error: "ไม่พบโปรเจกต์ " + projectId };
  const headers = (await getHeaders(projectId)) || TASK_HEADERS;
  const statusCol0 = headers.indexOf("Status");
  const progressCol0 = headers.indexOf("Progress");
  const data = [];
  if (statusCol0 > -1) data.push({ range: `${sheets.quoteSheetName(projectId)}!${sheets.colLetter(statusCol0 + 1)}${rowIndex}`, values: [[status]] });
  if (progressCol0 > -1) data.push({ range: `${sheets.quoteSheetName(projectId)}!${sheets.colLetter(progressCol0 + 1)}${rowIndex}`, values: [[progress]] });
  if (data.length) await sheets.batchUpdateValues(data);
  return { ok: true };
}

// ---------------------------------------------------------------------------
// Log — เนื่องจากระบบไม่มี login จริง เก็บแค่ เวลา/ผู้ใช้ที่เลือก/การกระทำ
// ---------------------------------------------------------------------------

async function logAction(user, action, projectId, detail) {
  let sheetId = await sheets.getSheetIdByTitle(SHEET_LOG);
  if (sheetId === null) {
    const addResult = await sheets.batchUpdate([{ addSheet: { properties: { title: SHEET_LOG } } }]);
    sheetId = addResult.replies[0].addSheet.properties.sheetId;
    await sheets.updateValues(`${sheets.quoteSheetName(SHEET_LOG)}!A1:E1`, [["Timestamp", "User", "Action", "Project", "Detail"]]);
    await sheets.batchUpdate([
      {
        repeatCell: {
          range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 5 },
          cell: { userEnteredFormat: { textFormat: { bold: true } } },
          fields: "userEnteredFormat.textFormat.bold",
        },
      },
    ]);
  }
  await sheets.appendValues(`${sheets.quoteSheetName(SHEET_LOG)}!A:E`, [
    [new Date().toISOString(), user, action, projectId || "", detail || ""],
  ]);
}

async function getLog(limit) {
  const sheetId = await sheets.getSheetIdByTitle(SHEET_LOG);
  if (sheetId === null) return [];
  const rows = await sheets.getValues(`${sheets.quoteSheetName(SHEET_LOG)}`);
  if (rows.length < 2) return [];
  const headers = rows[0];
  const entries = rows.slice(1).map((row) => rowToObject(headers, row, ["Timestamp"]));
  entries.reverse();
  return limit ? entries.slice(0, limit) : entries;
}

export default {
  getOverview,
  addProject,
  editProjectDetail,
  deleteProject,
  getTasks,
  getAllTasks,
  addTask,
  editTask,
  deleteTask,
  updateTaskStatus,
  logAction,
  getLog,
};
