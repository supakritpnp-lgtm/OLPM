# ระบบติดตามงานภายใน

เว็บติดตามโปรเจกต์/งาน อ่าน-เขียนข้อมูลตรงกับ Google Sheet ผ่าน Google Apps Script
(URL ตั้งไว้แล้วใน `src/config.js`)

## รันดูบนเครื่องตัวเอง

```bash
npm install
npm run dev
```
เปิด http://localhost:5173

**ก่อนรัน** ให้เช็คว่า Apps Script deploy แบบ "Who has access: Anyone" แล้ว
(ไม่งั้นเว็บจะดึงข้อมูลไม่ได้ เพราะเบราว์เซอร์ยิง request แบบไม่ล็อกอิน)

## Deploy ขึ้น Netlify

1. สร้าง repo ใหม่บน GitHub แล้ว push โฟลเดอร์นี้ขึ้นไปทั้งหมด
   ```bash
   git init
   git add .
   git commit -m "init project tracker"
   git branch -M main
   git remote add origin <URL ของ repo คุณ>
   git push -u origin main
   ```
2. เข้า https://app.netlify.com > "Add new site" > "Import an existing project"
3. เลือก repo ที่เพิ่ง push ไป — Netlify จะอ่านค่า build จาก `netlify.toml` ให้อัตโนมัติ
   (build command: `npm run build`, publish directory: `dist`)
4. กด Deploy — เสร็จแล้วจะได้ลิงก์เว็บ เช่น `your-site.netlify.app`
5. ทุกครั้งที่ push โค้ดใหม่ขึ้น GitHub, Netlify จะ build และอัปเดตเว็บให้อัตโนมัติ

## ถ้าต้องการเปลี่ยน URL ของ Apps Script ภายหลัง

แก้ค่าเดียวใน `src/config.js` แล้ว push ขึ้น GitHub ใหม่ Netlify จะ build ให้เอง

## โครงสร้างข้อมูลที่ระบบคาดหวัง

**Sheet "Overview"** ต้องมีคอลัมน์: `Project ID`, `Project name`, `PM`, `Sales Owner`,
`Start Date`, `Due Date`, `Project Value`, `Task progress` (0–1), `Platform`

**Sheet ต่อโปรเจกต์** (ชื่อ sheet = Project ID เช่น `PM-260601`) ต้องมีคอลัมน์:
`Type`, `Status`, `Progress` (0–1), `Task Name`, `Link`, `Task Detail`

## ปัญหาที่พบบ่อย

- **ขึ้น error "โหลดข้อมูลไม่สำเร็จ"** → เช็คว่า Apps Script deploy เป็น "Web app" ที่
  "Who has access: Anyone" ไม่ใช่ "Only myself"
- **กดเปลี่ยนสถานะแล้วไม่เซฟ** → เช็คว่าคอลัมน์ `Status` และ `Progress` สะกดตรงกับที่
  backend คาดหวัง (`apps_script_backend.gs` อ่านจากชื่อหัวคอลัมน์ตรงๆ)
- **แก้โค้ด backend แล้วเว็บไม่เปลี่ยน** → ต้องกด Deploy > Manage deployments > New deployment
  ทุกครั้งที่แก้โค้ดใน Apps Script (แก้แล้ว Save เฉยๆ ไม่พอ)
