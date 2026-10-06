// เชื่อมต่อ Google Sheets โดยตรงผ่าน Sheets API v4 ด้วย Service Account
// (ไม่ผ่าน Apps Script เลย — ตัดปัญหา deployment หลุด/URL หมดอายุ/ต้อง reauthorize ที่เจอบ่อยไปทั้งหมด)
//
// ต้องตั้งค่า Environment variables ใน Netlify (Site settings > Environment variables)
// หรือไฟล์ .env ในเครื่องตอนรัน `netlify dev`:
//   GOOGLE_SERVICE_ACCOUNT_EMAIL  อีเมลของ Service Account (ลงท้าย ...iam.gserviceaccount.com)
//   GOOGLE_PRIVATE_KEY            private key จากไฟล์ JSON ที่ดาวน์โหลดมา (เก็บ \n ไว้ตามเดิม)
//   GOOGLE_SPREADSHEET_ID         ID ของ Google Sheet (ส่วนที่อยู่ระหว่าง /d/ กับ /edit ใน URL)
//
// อย่าลืม "แชร์" Google Sheet ให้อีเมลของ Service Account เป็น Editor ด้วย ไม่งั้นจะเจอ
// permission denied ตอนอ่าน/เขียนข้อมูล

import { JWT } from "google-auth-library";

const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];

// รับได้ทั้งรหัสเพียวๆ และกรณีเผลอใส่ URL ทั้งเส้น / มีช่องว่าง / มีเครื่องหมายคำพูดติดมา
// เช่น "https://docs.google.com/spreadsheets/d/1AbC.../edit#gid=0" หรือ "1AbC.../edit?gid=0#gid=0" -> "1AbC..."
function normalizeSpreadsheetId(raw) {
  const s = String(raw || "").trim().replace(/^["']+|["']+$/g, "").trim();
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  // กรณีคัดลอกมาเฉพาะส่วนท้าย เช่น "1AbC.../edit?pli=1&gid=0#gid=0" -> ตัดตั้งแต่ / ? # เป็นต้นไปทิ้ง
  return s.split(/[\/?#]/)[0];
}

function getSpreadsheetId() {
  const id = normalizeSpreadsheetId(process.env.GOOGLE_SPREADSHEET_ID);
  if (!id) throw new Error("ยังไม่ได้ตั้งค่า Environment variable: GOOGLE_SPREADSHEET_ID");
  if (!/^[a-zA-Z0-9-_]+$/.test(id)) {
    throw new Error(
      "GOOGLE_SPREADSHEET_ID รูปแบบไม่ถูกต้อง (มีอักขระแปลกปลอม) — ต้องเป็นเฉพาะรหัสที่อยู่ระหว่าง /d/ กับ /edit ใน URL ของ Google Sheet"
    );
  }
  return id;
}

function getBase() {
  return `https://sheets.googleapis.com/v4/spreadsheets/${getSpreadsheetId()}`;
}

let authClient = null;
function getAuthClient() {
  if (authClient) return authClient;
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const rawKey = process.env.GOOGLE_PRIVATE_KEY;
  if (!email || !rawKey) {
    throw new Error(
      "ยังไม่ได้ตั้งค่า Environment variables: GOOGLE_SERVICE_ACCOUNT_EMAIL และ/หรือ GOOGLE_PRIVATE_KEY"
    );
  }
  // ตอนเก็บ private key ใน env var ทั่วไปมักถูกเก็บเป็นบรรทัดเดียวโดยแทนขึ้นบรรทัดใหม่ด้วย \n
  // (ตัวอักษร backslash-n) ต้องแปลงกลับเป็นตัวขึ้นบรรทัดใหม่จริงก่อนใช้เซ็น JWT
  const key = rawKey.replace(/\\n/g, "\n");
  authClient = new JWT({ email, key, scopes: SCOPES });
  return authClient;
}

async function getAccessToken() {
  const client = getAuthClient();
  const token = await client.authorize();
  return token.access_token;
}

async function sheetsFetch(pathAndQuery, options = {}) {
  const token = await getAccessToken();
  const res = await fetch(`${getBase()}${pathAndQuery}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  const text = await res.text();
  return parseSheetsResponse(res.status, res.ok, text);
}

// แปลงคำตอบจาก Sheets API — ถ้า Google ตอบกลับมาเป็น HTML (ไม่ใช่ JSON) จะอธิบายสาเหตุให้ชัด
// แทนที่จะขึ้นข้อความงงๆ ว่า "Unexpected token '<'"
function parseSheetsResponse(status, ok, text) {
  let data = {};
  if (text) {
    try {
      data = JSON.parse(text);
    } catch (e) {
      const title = (text.match(/<title>([^<]*)<\/title>/i) || [])[1] || text.slice(0, 80).replace(/\s+/g, " ");
      throw new Error(
        `Google Sheets API ตอบกลับไม่ใช่ JSON (HTTP ${status}: ${title.trim()}) — ` +
          `มักเกิดจาก GOOGLE_SPREADSHEET_ID ผิดรูปแบบ หรือมี proxy/ไฟร์วอลล์ขององค์กรแทรกแซงการเชื่อมต่อ`
      );
    }
  }
  if (!ok) {
    const message = (data.error && data.error.message) || `Sheets API error ${status}`;
    throw new Error(message);
  }
  return data;
}

// ใส่เครื่องหมายคำพูดเดี่ยวล้อมชื่อชีตเสมอ (จำเป็นถ้าชื่อมีขีด/เว้นวรรค เช่น Project ID "PM-260801")
function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

// แปลงเลขคอลัมน์ (1 = A) เป็นตัวอักษรคอลัมน์แบบ A1 notation
function colLetter(n) {
  let s = "";
  let x = n;
  while (x > 0) {
    const rem = (x - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

// อ่านค่าดิบ (ไม่ format) — ตัวเลข/บูลีนจะได้ชนิดข้อมูลจริงเหมือน Apps Script's getValues()
// ยกเว้นวันที่ ซึ่งจะได้เป็นเลข serial แทน Date object ต้องแปลงเองที่ชั้น logic.js
async function getValues(rangeA1) {
  const data = await sheetsFetch(
    `/values/${encodeURIComponent(rangeA1)}?valueRenderOption=UNFORMATTED_VALUE&majorDimension=ROWS`
  );
  return data.values || [];
}

// อ่านหลายช่วงพร้อมกันในคำขอเดียว (ใช้แทนการยิงทีละชีต ~40-50 ครั้งตอน export/alltasks)
async function batchGetValues(rangesA1) {
  const query = rangesA1.map((r) => `ranges=${encodeURIComponent(r)}`).join("&");
  const data = await sheetsFetch(`/values:batchGet?valueRenderOption=UNFORMATTED_VALUE&${query}`);
  return data.valueRanges || []; // [{ range, values }], เรียงลำดับตรงกับ rangesA1 ที่ส่งไป
}

async function updateValues(rangeA1, values) {
  return sheetsFetch(`/values/${encodeURIComponent(rangeA1)}?valueInputOption=USER_ENTERED`, {
    method: "PUT",
    body: JSON.stringify({ range: rangeA1, values }),
  });
}

// เขียนหลายช่วง/เซลล์พร้อมกันในคำขอเดียว — ใช้แทนการ update ทีละเซลล์เวลามีหลาย field เปลี่ยน
async function batchUpdateValues(data) {
  // data: [{ range: "Sheet!A1", values: [[value]] }, ...]
  return sheetsFetch(`/values:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({ valueInputOption: "USER_ENTERED", data }),
  });
}

async function appendValues(rangeA1, values) {
  return sheetsFetch(
    `/values/${encodeURIComponent(rangeA1)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    { method: "POST", body: JSON.stringify({ values }) }
  );
}

// batchUpdate สำหรับคำสั่งเชิงโครงสร้าง (แทรก/ลบแถว, สร้าง/ลบชีต, จัดฟอร์แมต ฯลฯ)
async function batchUpdate(requests) {
  return sheetsFetch(`:batchUpdate`, { method: "POST", body: JSON.stringify({ requests }) });
}

async function getSpreadsheetMeta() {
  return sheetsFetch(`?fields=sheets.properties`);
}

async function getSheetIdByTitle(title) {
  const meta = await getSpreadsheetMeta();
  const sheet = (meta.sheets || []).find((s) => s.properties.title === title);
  return sheet ? sheet.properties.sheetId : null;
}

export default {
  normalizeSpreadsheetId,
  parseSheetsResponse,
  getValues,
  batchGetValues,
  updateValues,
  batchUpdateValues,
  appendValues,
  batchUpdate,
  getSpreadsheetMeta,
  getSheetIdByTitle,
  quoteSheetName,
  colLetter,
};
