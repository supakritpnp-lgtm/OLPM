// API หลักของเว็บ — แทนที่ apps_script_backend.gs ทั้งหมด
// ทำงานเป็น Netlify Function ธรรมดา คุยกับ Google Sheets ตรงๆ ผ่าน Sheets API v4
// (ดู netlify/functions/lib/sheets.js และ lib/logic.js)
//
// เรียกใช้จากหน้าเว็บที่ path เดียวกันเสมอ: /.netlify/functions/api
//   GET  /.netlify/functions/api?action=overview
//   POST /.netlify/functions/api   (body เป็น JSON เหมือนเดิมทุกอย่าง)
// ไม่มีปัญหา CORS เลย เพราะอยู่โดเมนเดียวกับตัวเว็บทั้งตอน dev (netlify dev) และตอน deploy จริง

import logic from "./lib/logic.js";

export const handler = async (event) => {
  try {
    if (event.httpMethod === "GET") {
      const params = event.queryStringParameters || {};
      const action = params.action || "overview";
      let data;
      if (action === "overview") data = await logic.getOverview();
      else if (action === "tasks") data = await logic.getTasks(params.projectId);
      else if (action === "alltasks") data = await logic.getAllTasks();
      else if (action === "log") data = await logic.getLog(Number(params.limit) || 200);
      else data = { error: "unknown action" };
      return respond(200, data);
    }

    if (event.httpMethod === "POST") {
      const body = JSON.parse(event.body || "{}");
      const user = body.user || "ไม่ระบุ";
      let result;

      switch (body.action) {
        case "updateTaskStatus":
          result = await logic.updateTaskStatus(body.projectId, body.rowIndex, body.status, body.progress);
          if (!result.error) await logic.logAction(user, "เปลี่ยนสถานะงาน", body.projectId, "แถว " + body.rowIndex + " → " + body.status);
          break;
        case "addTask":
          result = await logic.addTask(body.projectId, body.task);
          if (!result.error) await logic.logAction(user, "เพิ่มงาน", body.projectId, (body.task && body.task["Task Name"]) || "");
          break;
        case "editTask":
          result = await logic.editTask(body.projectId, body.rowIndex, body.fields);
          if (!result.error) await logic.logAction(user, "แก้ไขงาน", body.projectId, (body.fields && body.fields["Task Name"]) || "แถว " + body.rowIndex);
          break;
        case "deleteTask":
          result = await logic.deleteTask(body.projectId, body.rowIndex);
          if (!result.error) await logic.logAction(user, "ลบงาน", body.projectId, "แถว " + body.rowIndex);
          break;
        case "addProject":
          result = await logic.addProject(body.project);
          if (!result.error) await logic.logAction(user, "เพิ่มโปรเจกต์", result.projectId, (body.project && body.project.name) || "");
          break;
        case "editProjectDetail":
          result = await logic.editProjectDetail(body.projectId, body.fields);
          if (!result.error) await logic.logAction(user, "แก้ไขรายละเอียดโปรเจกต์", body.projectId, Object.keys(body.fields || {}).join(", "));
          break;
        case "deleteProject":
          result = await logic.deleteProject(body.projectId);
          if (!result.error) await logic.logAction(user, "ลบโปรเจกต์", body.projectId, "");
          break;
        default:
          return respond(200, { error: "unknown action" });
      }
      return respond(200, result);
    }

    return respond(405, { error: "method not allowed" });
  } catch (err) {
    console.error(err);
    return respond(500, { error: err.message || "internal error" });
  }
};

function respond(statusCode, data) {
  return {
    statusCode,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
    body: JSON.stringify(data),
  };
}
