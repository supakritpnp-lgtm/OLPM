// สร้างไฟล์ PowerPoint สรุปโปรเจกต์ทั้งหมด — รันฝั่งเบราว์เซอร์ทั้งหมด ไม่ต้องมี backend แยก
//
// หลักการสำคัญ: ทุกกล่อง/ตารางมีตำแหน่งและขนาดคงที่ตายตัว ไม่มีการใช้ autoPage ที่ไหนเลย
// เพื่อไม่ให้เกิดปัญหาข้อมูลล้นไปสร้างสไลด์ที่ 2 แบบไม่ได้ตั้งใจ — ถ้าข้อมูลยาวเกินจะถูกตัด (…)
// หรือใส่ข้อความ "+N รายการ" แทน ไม่ปล่อยให้ล้นกรอบ
// Dashboard แยกเป็น 3 สไลด์: Project เต็ม, Solution เต็ม, Work Load
import pptxgen from "pptxgenjs";

const COLOR = {
  ink: "24262B",
  inkMuted: "6B6D70",
  inkFaint: "9A9A94",
  border: "D9D8D2",
  headerBg: "EFEFEA",
  accent: "33538F",
  accentSoft: "E8EDF5",
  green: "3E8264",
  greenSoft: "E5F0EA",
  amber: "B9822E",
  violet: "6E5CA6",
  violetSoft: "EFEBF6",
  white: "FFFFFF",
  bg: "FBFBFA",
};

function num(n) {
  if (n === "" || n === null || n === undefined) return "–";
  const v = Number(n);
  if (isNaN(v)) return String(n);
  return v.toLocaleString("en-US", { maximumFractionDigits: 0 });
}

function bool(v) {
  return v ? "✓" : "–";
}

function truncate(s, max) {
  if (!s) return "–";
  const str = String(s).replace(/\n/g, " ").trim();
  if (!str) return "–";
  return str.length > max ? str.slice(0, max - 1).trimEnd() + "…" : str;
}

function parsePMString(s) {
  return s ? String(s).split(",").map((x) => x.trim()).filter(Boolean) : [];
}

function hasPM(project, pmName) {
  return parsePMString(project.pm).includes(pmName);
}

// ---------------------------------------------------------------------------
// Dashboard — 1 สไลด์ต่อประเภทงาน (Project, Solution, ...) + Work Load
// อยู่ใต้รายการเสมอในสไลด์เดียวกัน ใช้ผังคงที่ 3 คอลัมน์ ไม่มีทางล้นหน้า
// ---------------------------------------------------------------------------

function addCategorySlide(pres, title, items, color, projectItems, solutionItems, pmUsers) {
  const sumValue = items.reduce((s, p) => s + (Number(p.value) || 0), 0);
  const sumValueOf = (arr) => arr.reduce((s, p) => s + (Number(p.value) || 0), 0);

  const s = pres.addSlide();
  s.background = { color: COLOR.white };
  s.addText(`${title} Dashboard`, { x: 0.5, y: 0.25, w: 8, h: 0.55, fontSize: 36, bold: true, color: COLOR.ink });
  s.addText(new Date().toLocaleDateString("th-TH", { day: "2-digit", month: "short", year: "numeric" }), {
    x: 9.0, y: 0.42, w: 3.8, h: 0.32, fontSize: 16, color: COLOR.inkFaint, align: "right",
  });
  s.addText([{ text: `${title} : `, options: { bold: true, color: COLOR.ink } }, { text: String(items.length), options: { bold: true, color } }], {
    x: 0.5, y: 1.0, w: 6, h: 0.34, fontSize: 20,
  });
  s.addText(`Value : ${num(sumValue)}`, { x: 0.5, y: 1.36, w: 6, h: 0.26, fontSize: 12, color: COLOR.inkMuted });

  // เผื่อพื้นที่ด้านล่างไว้ให้ Work Load เสมอ
  const WORKLOAD_TOP = 5.55;
  const listTop = 1.75;
  const ROW_H = 0.5;
  const NUM_COLS = 3;
  const colW = 4.0;
  const colGap = 0.17;
  const MAX_ROWS = Math.floor((WORKLOAD_TOP - 0.15 - listTop) / ROW_H);
  const maxShown = Math.min(items.length, MAX_ROWS * NUM_COLS);
  const shown = items.slice(0, maxShown);

  shown.forEach((p, i) => {
    const row = Math.floor(i / NUM_COLS);
    const col = i % NUM_COLS;
    const ix = 0.5 + col * (colW + colGap);
    const iy = listTop + row * ROW_H;
    const pct = Math.round((p.progress || 0) * 100);
    const pctColor = pct >= 100 ? COLOR.green : COLOR.accent;
    s.addText(String(i + 1).padStart(2, "0"), { x: ix, y: iy, w: 0.4, h: 0.26, fontSize: 12, color: COLOR.inkFaint, wrap: false });
    s.addText(truncate(p.name, 20), { x: ix + 0.4, y: iy, w: 2.95, h: 0.26, fontSize: 14, bold: true, color: COLOR.ink, wrap: false });
    s.addText(`${pct}%`, { x: ix + colW - 0.65, y: iy, w: 0.65, h: 0.26, fontSize: 14, bold: true, color: pctColor, align: "right", wrap: false });
    s.addText(num(p.value), { x: ix + 0.4, y: iy + 0.25, w: colW - 0.4, h: 0.2, fontSize: 10.5, color: COLOR.inkFaint });
  });

  if (items.length > maxShown) {
    s.addText(`+ อีก ${items.length - maxShown} โปรเจกต์`, {
      x: 0.5, y: listTop + Math.ceil(maxShown / NUM_COLS) * ROW_H + 0.02, w: 12, h: 0.24, fontSize: 11, italic: true, color: COLOR.inkFaint,
    });
  }

  // Work Load — อยู่ใต้รายการเสมอ
  s.addText("Workload", { x: 0.5, y: WORKLOAD_TOP, w: 6, h: 0.4, fontSize: 24, bold: true, color: COLOR.ink });

  const cardW = (12.33 - 0.3 * (pmUsers.length - 1)) / pmUsers.length;
  const cardY = WORKLOAD_TOP + 0.55;
  const cardH = 1.0;
  pmUsers.forEach((pm, i) => {
    const x = 0.5 + i * (cardW + 0.3);
    const pProj = projectItems.filter((p) => hasPM(p, pm));
    const pSol = solutionItems.filter((p) => hasPM(p, pm));
    s.addShape(pres.ShapeType.roundRect, { x, y: cardY, w: cardW, h: cardH, rectRadius: 0.06, fill: { color: COLOR.bg }, line: { color: COLOR.border, width: 0.75 } });
    s.addText(pm, { x: x + 0.16, y: cardY + 0.08, w: cardW - 0.32, h: 0.3, fontSize: 18, bold: true, color: COLOR.ink });
    s.addText([{ text: "Project ", options: { color: COLOR.inkMuted } }, { text: String(pProj.length), options: { bold: true, color: COLOR.accent } }, { text: `   ${num(sumValueOf(pProj))}`, options: { color: COLOR.inkFaint, fontSize: 9 } }], {
      x: x + 0.16, y: cardY + 0.46, w: cardW - 0.32, h: 0.24, fontSize: 11.5, wrap: false,
    });
    s.addText([{ text: "Solution ", options: { color: COLOR.inkMuted } }, { text: String(pSol.length), options: { bold: true, color: COLOR.violet } }, { text: `   ${num(sumValueOf(pSol))}`, options: { color: COLOR.inkFaint, fontSize: 9 } }], {
      x: x + 0.16, y: cardY + 0.72, w: cardW - 0.32, h: 0.24, fontSize: 11.5, wrap: false,
    });
  });
}

// ---------------------------------------------------------------------------
// Project detail — 2 โปรเจกต์ต่อสไลด์ ตำแหน่งคงที่ ไม่มีทางล้นหน้า
// ---------------------------------------------------------------------------

const CARD_Y = 1.15;
const CARD_H = 6.05;
const FIELD_ROW_H = 0.27;
const MAX_TASKS_SHOWN = 5;

function drawProjectCard(pres, s, p, tasks, x, w) {
  s.addShape(pres.ShapeType.roundRect, { x, y: CARD_Y, w, h: CARD_H, rectRadius: 0.06, fill: { color: COLOR.white }, line: { color: COLOR.border, width: 1 } });

  const pad = 0.25;
  const innerX = x + pad;
  const innerW = w - pad * 2;

  // Title + progress pill
  const pct = Math.round((p.progress || 0) * 100);
  const pillColor = pct >= 100 ? COLOR.green : COLOR.accent;
  const pillSoft = pct >= 100 ? COLOR.greenSoft : COLOR.accentSoft;
  s.addText(truncate(p.name, 28), { x: innerX, y: CARD_Y + 0.15, w: innerW - 1.2, h: 0.44, fontSize: 20, bold: true, color: COLOR.ink, wrap: false });
  s.addText(p.id || "", { x: innerX, y: CARD_Y + 0.58, w: innerW - 1.2, h: 0.24, fontSize: 11, color: COLOR.inkFaint, fontFace: "Courier New" });
  s.addShape(pres.ShapeType.roundRect, { x: x + w - pad - 1.0, y: CARD_Y + 0.18, w: 1.0, h: 0.4, rectRadius: 0.2, fill: { color: pillSoft } });
  s.addText(`${pct}%`, { x: x + w - pad - 1.0, y: CARD_Y + 0.18, w: 1.0, h: 0.4, fontSize: 16, bold: true, color: pillColor, align: "center", valign: "middle" });

  // Field grid — 2 คอลัมน์
  const fieldsTop = CARD_Y + 1.0;
  const colGap = 0.25;
  const colW = (innerW - colGap) / 2;
  const col2X = innerX + colW + colGap;

  const pairs = [
    ["PM", p.pm || "–", "Sales Owner", p.sales || "–"],
    ["Value", num(p.value), "Volume", p.volume ?? "–"],
    ["Cost", num(p.cost), "Benefit", typeof p.benefit === "number" ? num(p.benefit) : p.benefit || "–"],
    ["Sale brief", bool(p.saleBrief), "Contract/PO", bool(p.contractPO)],
    ["GPS", truncate(p.gps, 12), "GPS Option", truncate(p.gpsOption, 20)],
    ["MDVR", truncate(p.mdvr, 12), "MDVR Option", truncate(p.mdvrOption, 20)],
    ["Platform", truncate(p.platform, 12), "Platform Customize", bool(p.platformCustomize)],
    ["Platform Dev Detail", truncate(p.platformDevDetail, 20), "", ""],
  ];

  const fieldCell = (label, value, cx, cy, cw) => {
    if (!label) return;
    s.addText(
      [
        { text: `${label}  `, options: { bold: true, color: COLOR.inkMuted } },
        { text: value, options: { color: COLOR.ink } },
      ],
      { x: cx, y: cy, w: cw, h: FIELD_ROW_H, fontSize: 12, valign: "middle", wrap: false }
    );
  };

  pairs.forEach((row, i) => {
    const y = fieldsTop + i * FIELD_ROW_H;
    fieldCell(row[0], row[1], innerX, y, colW);
    fieldCell(row[2], row[3], col2X, y, colW);
  });

  const noteY = fieldsTop + pairs.length * FIELD_ROW_H + 0.05;
  s.addText(
    [
      { text: "Project Note  ", options: { bold: true, color: COLOR.inkMuted } },
      { text: truncate(p.note, 90), options: { color: COLOR.ink } },
    ],
    { x: innerX, y: noteY, w: innerW, h: 0.44, fontSize: 12, valign: "top", wrap: true }
  );

  const dividerY = noteY + 0.48;
  s.addShape(pres.ShapeType.line, { x: innerX, y: dividerY, w: innerW, h: 0, line: { color: COLOR.border, width: 0.75 } });

  // Tasks
  const tasksTop = dividerY + 0.15;
  if (!tasks || tasks.length === 0) {
    s.addText("ยังไม่มี Task ในชีทของโปรเจกต์นี้", { x: innerX, y: tasksTop, w: innerW, h: 0.32, fontSize: 12, italic: true, color: COLOR.inkFaint });
    return;
  }

  // วาดตาราง Task เองด้วย addText/addShape แทน addTable เพราะ addTable ของ pptxgenjs
  // ไม่รองรับ wrap:false ในเซลล์จริง (ค่านี้จะถูกละเว้นเงียบๆ) ทำให้ข้อความยาวยังขึ้นบรรทัดใหม่
  // ได้ในโปรแกรม PowerPoint จริง จนไปทับข้อความ "+ อีก N งาน" ด้านล่าง — addText เท่านั้นที่
  // เคารพ wrap:false จริง จึงการันตีว่าแต่ละแถวสูงเท่ากันเป๊ะ ไม่มีทางล้น
  const shown = tasks.slice(0, MAX_TASKS_SHOWN);
  const TASK_ROW_H = 0.3;
  const nameColW = innerW - 1.5;
  const statusColX = innerX + nameColW;
  const pctColX = innerX + nameColW + 0.85;

  s.addShape(pres.ShapeType.rect, { x: innerX, y: tasksTop, w: innerW, h: TASK_ROW_H, fill: { color: COLOR.accent }, line: { type: "none" } });
  s.addText("Task", { x: innerX + 0.08, y: tasksTop, w: nameColW - 0.08, h: TASK_ROW_H, fontSize: 11, bold: true, color: COLOR.white, valign: "middle", wrap: false });
  s.addText("Status", { x: statusColX, y: tasksTop, w: 0.85, h: TASK_ROW_H, fontSize: 11, bold: true, color: COLOR.white, valign: "middle", wrap: false });
  s.addText("%", { x: pctColX, y: tasksTop, w: innerW - nameColW - 0.85 - 0.08, h: TASK_ROW_H, fontSize: 11, bold: true, color: COLOR.white, align: "right", valign: "middle", wrap: false });

  shown.forEach((t, i) => {
    const rowY = tasksTop + (i + 1) * TASK_ROW_H;
    s.addShape(pres.ShapeType.rect, { x: innerX, y: rowY, w: innerW, h: TASK_ROW_H, fill: { color: COLOR.white }, line: { color: COLOR.border, width: 0.5 } });
    s.addText(truncate(t["Task Name"], 26), { x: innerX + 0.08, y: rowY, w: nameColW - 0.08, h: TASK_ROW_H, fontSize: 11, color: COLOR.ink, valign: "middle", wrap: false });
    s.addText(t["Status"] || "–", { x: statusColX, y: rowY, w: 0.85, h: TASK_ROW_H, fontSize: 10.5, color: COLOR.inkMuted, valign: "middle", wrap: false });
    s.addText(
      t["Progress"] || t["Progress"] === 0 ? `${Math.round(Number(t["Progress"]) * 100)}%` : "–",
      { x: pctColX, y: rowY, w: innerW - nameColW - 0.85 - 0.08, h: TASK_ROW_H, fontSize: 10.5, color: COLOR.accent, align: "right", valign: "middle", wrap: false }
    );
  });

  if (tasks.length > MAX_TASKS_SHOWN) {
    const moreY = tasksTop + (MAX_TASKS_SHOWN + 1) * TASK_ROW_H + 0.15;
    s.addText(`+ อีก ${tasks.length - MAX_TASKS_SHOWN} งาน`, {
      x: innerX, y: moreY, w: innerW, h: 0.26, fontSize: 11, italic: true, color: COLOR.inkFaint,
    });
  }
}

const OVERVIEW_TYPE_ORDER = ["Project", "Solution", "Issue", "CR", "IT Request"];

function addProjectPairSlides(pres, projects, tasksByProjectId) {
  const byType = new Map();
  projects.forEach((p) => {
    const key = OVERVIEW_TYPE_ORDER.includes(p.type) ? p.type : p.type || "อื่นๆ";
    if (!byType.has(key)) byType.set(key, []);
    byType.get(key).push(p);
  });

  const orderedKeys = [...OVERVIEW_TYPE_ORDER.filter((t) => byType.has(t)), ...[...byType.keys()].filter((k) => !OVERVIEW_TYPE_ORDER.includes(k))];

  orderedKeys.forEach((type) => {
    const items = byType.get(type);
    const pairs = [];
    for (let i = 0; i < items.length; i += 2) pairs.push([items[i], items[i + 1]]);

    pairs.forEach(([a, b], idx) => {
      const s = pres.addSlide();
      s.background = { color: COLOR.white };
      s.addText(`${type} Overview`, { x: 0.5, y: 0.25, w: 8, h: 0.5, fontSize: 24, bold: true, color: COLOR.ink });
      s.addText(`${idx + 1} / ${pairs.length}`, { x: 11.0, y: 0.38, w: 1.8, h: 0.32, fontSize: 14, color: COLOR.inkFaint, align: "right" });

      const gap = 0.3;
      const cardW = (12.33 - gap) / 2;
      drawProjectCard(pres, s, a, tasksByProjectId?.[a.id], 0.5, cardW);
      if (b) drawProjectCard(pres, s, b, tasksByProjectId?.[b.id], 0.5 + cardW + gap, cardW);
    });
  });
}

export async function exportPptx({ projects, tasksByProjectId, pmUsers }) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.33" x 7.5"
  pres.author = "Project Desk";
  pres.title = "Project Overview";

  const projectItems = projects.filter((p) => p.type === "Project");
  const solutionItems = projects.filter((p) => p.type === "Solution");

  addCategorySlide(pres, "Project", projectItems, COLOR.accent, projectItems, solutionItems, pmUsers);
  addCategorySlide(pres, "Solution", solutionItems, COLOR.violet, projectItems, solutionItems, pmUsers);
  addProjectPairSlides(pres, projects, tasksByProjectId);

  const dateStr = new Date().toISOString().slice(0, 10);
  await pres.writeFile({ fileName: `Project-Overview-${dateStr}.pptx` });
}
