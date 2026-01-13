# Excel/CSV Validate → SQL (NodeJS Web App)

เว็บแอปนี้ใช้สำหรับ:
1) อัปโหลดไฟล์ Excel/CSV  
2) ทำ **PHASE 1 — EXTRACT (C→AI)** แสดงเป็นตาราง (ยังไม่สร้าง SQL)  
3) ถ้าผู้ใช้พิมพ์ยืนยัน **"Extract ตรง Excel 100%"** จึงทำ **PHASE 2 — GENERATE SQL** และ **ต่อท้ายด้วย Validation SQL** ตามที่กำหนดเสมอ

> ข้อสำคัญ: โปรแกรม “ไม่เดา/ไม่เติม/ไม่แก้/ไม่ normalize ค่า”  
> (ยกเว้นข้อจำกัดของ CSV ที่ค่าว่างมักถูกมองเป็น NULL)

---

## สิ่งที่ต้องมี (Prerequisites)
- เครื่อง Windows / macOS / Linux
- Internet สำหรับดาวน์โหลด NodeJS และ packages
- Browser (Chrome/Edge/Safari)

---

## Step 1) ติดตั้ง NodeJS (สำหรับคนไม่เคยติดตั้ง)

### Windows / macOS (แนะนำ)
1. ไปที่เว็บไซต์ NodeJS: https://nodejs.org  
2. ดาวน์โหลด **LTS** (แนะนำที่สุด)  
3. ติดตั้งตามขั้นตอน (กด Next ไปเรื่อยๆ ได้)

### ตรวจสอบว่าติดตั้งสำเร็จ
เปิด Terminal/Command Prompt แล้วพิมพ์:

**Windows**
- เปิด `Command Prompt` หรือ `PowerShell`

**macOS**
- เปิด `Terminal`

แล้วรัน:
```bash
node -v
npm -v
```

ถ้าขึ้นเวอร์ชันประมาณ `v18.x` หรือ `v20.x` แปลว่าพร้อมใช้งานแล้ว

---

## Step 2) เตรียมโปรเจกต์

ให้คุณมีโฟลเดอร์โปรเจกต์หน้าตาประมาณนี้:

```
excel-validate-app/
  package.json
  server.js
  public/
    index.html
    app.js
```

> ถ้ายังไม่มีไฟล์ ให้เอาโค้ดที่คุยกันวางตามไฟล์ชื่อเดียวกัน

---

## Step 3) ติดตั้ง dependencies (ครั้งแรกครั้งเดียว)

1) เปิด Terminal/Command Prompt  
2) เข้าไปที่โฟลเดอร์โปรเจกต์ เช่น:

```bash
cd excel-validate-app
```

3) ติดตั้งแพ็กเกจ:
```bash
npm install
```

> ขั้นตอนนี้จะสร้างโฟลเดอร์ `node_modules/` อัตโนมัติ (อาจใช้เวลานิดหน่อย)

---

## Step 4) รันเว็บแอป

ในโฟลเดอร์โปรเจกต์เดียวกัน รัน:

```bash
npm start
```

ถ้าเห็นข้อความประมาณนี้ แปลว่าเซิร์ฟเวอร์ทำงานแล้ว:
- `Server running: http://localhost:3000`

---

## Step 5) เปิดหน้าเว็บ

เปิด Browser แล้วเข้า:
- http://localhost:3000

คุณจะเห็นหน้าเว็บสำหรับ:
- ดาวน์โหลด Template
- อัปโหลดไฟล์
- ดูผล Extract
- Generate SQL

---

## Step 6) ดาวน์โหลด Template (Header Only)

ในหน้าเว็บจะมีลิงก์:
- **Download Template .xlsx**
- **Download Template .csv**

Template จะเป็น “เฉพาะ header” ไม่มี sample data

---

## วิธีใช้งานตาม Process (2 Phase)

### PHASE 1 — EXTRACT
1) เลือกไฟล์ `.xlsx/.xls/.csv`  
2) กดปุ่ม **PHASE 1: Extract**  
3) ระบบจะแสดง “ตาราง Extract” ตามคอลัมน์ **C → AI**
   - `excel_row` จะเริ่มที่ 2 สำหรับแถวข้อมูลแรก
   - แถวที่คอลัมน์ C→AI ว่างทั้งหมด จะไม่ถูกแสดง

> ตารางในหน้านี้ **แก้ไขได้** (ถือว่าเป็น override ของผู้ใช้ = source of truth)

### PHASE 2 — GENERATE SQL
1) พิมพ์คำยืนยันให้ตรง 100%:
   - `Extract ตรง Excel 100%`
2) กด **PHASE 2: Generate SQL**
3) ระบบจะแสดง SQL ที่ประกอบด้วย:
   - `WITH excel AS (...)` ที่แปลงจากไฟล์/override แบบ 1:1
   - ต่อท้ายด้วย `SQL_VALIDATION_SUFFIX` (Validation SELECT + LEFT JOIN ...) เสมอ

---

## หมายเหตุเรื่อง CSV (สำคัญ)
- CSV โดยทั่วไป “ค่าว่าง” มักแยกไม่ชัดว่าเป็น `NULL` หรือ `''`
- แอปนี้จะตีความช่องว่างใน CSV เป็น `NULL` (ตามหมายเหตุใน requirement)
- ถ้าคุณต้องการให้ `''` (blank string) แยกจาก NULL ได้ชัด แนะนำใช้ `.xlsx`

---

## Troubleshooting

### 1) เปิด http://localhost:3000 ไม่ได้
- ตรวจว่า `npm start` ยังรันอยู่ใน Terminal
- ตรวจว่าไม่มีโปรแกรมอื่นใช้ port 3000

ถ้าจะเปลี่ยน port:

**macOS/Linux**
```bash
PORT=4000 npm start
```

**Windows (PowerShell)**
```powershell
$env:PORT=4000; npm start
```

แล้วเปิด:
- http://localhost:4000

### 2) `node -v` / `npm -v` ไม่เจอคำสั่ง
- แปลว่ายังไม่ได้ติดตั้ง NodeJS หรือ PATH ยังไม่ถูกตั้ง
- แนะนำติดตั้ง NodeJS LTS ใหม่จาก https://nodejs.org แล้ว “ปิด-เปิด Terminal ใหม่” อีกครั้ง

### 3) `npm install` error เรื่อง permission (macOS)
- ลองใช้:
```bash
sudo npm install
```
หรือแนะนำให้ติดตั้ง NodeJS ใหม่แบบ LTS (ส่วนใหญ่จะแก้ได้)

---

## ปิดโปรแกรม
กลับไปที่ Terminal ที่รัน `npm start` แล้วกด:
- `Ctrl + C`

---

## ความปลอดภัย
- แอปรันในเครื่องคุณเอง (localhost)
- ไฟล์ที่อัปโหลดถูกอ่านใน memory เพื่อ extract/generate SQL

---

## License
Internal / for validation usage
