// URL ของ API — ตอนนี้เป็น Netlify Function (คุยกับ Google Sheets ตรงๆ ผ่าน Service Account)
// ไม่ใช่ Google Apps Script Web App อีกต่อไป — ไม่มีปัญหา deployment หลุด/URL หมดอายุอีกแล้ว
// path นี้ใช้ได้ทั้งตอน deploy จริงบน Netlify และตอนรัน `netlify dev` ในเครื่อง (โดเมนเดียวกันเสมอ)
export const API_URL = "/.netlify/functions/api";
