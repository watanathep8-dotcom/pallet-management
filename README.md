# Pallet Hub — ระบบบริหารพาเลท RM + PK + FG (ราชบุรี)

เวอร์ชัน GitHub Pages + Google Apps Script (ย้ายมาจากเวอร์ชัน PHP/MySQL บน XAMPP)

- **หน้าเว็บ** (`docs/`) เป็นไฟล์ static เปิดบน GitHub Pages ได้เลย หน้าตาและการใช้งานเหมือนเดิมทุกอย่าง
- **ฐานข้อมูล** เป็น Google Sheet (1 ชีตต่อ 1 ตาราง: `pallet_types`, `departments`, `repairs`, `movements`, `audit_logs`)
- **API** เป็น Google Apps Script Web App (`apps-script/Code.gs`) กฎการทำงานเหมือน `api.php` เดิมทุกข้อ

## โครงสร้างไฟล์

```
docs/index.html          หน้าเว็บ (เดิมคือ index.php)
docs/config.js           ตั้งค่า apiUrl (URL ของ Apps Script /exec)
docs/assets/app.js       โค้ดหน้าเว็บ (เปลี่ยนเฉพาะส่วนเรียก API)
docs/assets/style.css    สไตล์ (เหมือนเดิม)
apps-script/Code.gs      โค้ดฝั่งเซิร์ฟเวอร์ (Google Apps Script)
appsscript.json          ไฟล์ manifest ของ Apps Script (เขตเวลา Asia/Bangkok, V8)
test/gas-mock.js         ชุดทดสอบ (Node.js) จำลอง Apps Script ในหน่วยความจำ — ไม่ได้เผยแพร่
```

## วิธีติดตั้ง (ทำครั้งเดียว)

### 1. สร้างโปรเจกต์ Apps Script
1. ไปที่ <https://script.google.com> → **New project** ตั้งชื่อ เช่น `Pallet Hub API`
2. ลบโค้ดเดิมในไฟล์ `Code.gs` แล้ว **คัดลอกเนื้อหาทั้งหมดจาก `apps-script/Code.gs`** ไปวาง
3. ไปที่ **Project Settings** (รูปเฟือง) → ติ๊ก **Show "appsscript.json" manifest file in editor**
4. กลับไปที่ Editor เปิดไฟล์ `appsscript.json` แล้ววางเนื้อหาจากไฟล์ `appsscript.json` ในโปรเจกต์นี้ (เขตเวลา `Asia/Bangkok`, runtime `V8`)
5. กด **Save**

### 2. สร้างฐานข้อมูล (Google Sheet)
1. เลือกฟังก์ชัน **`setupSystem`** ในแถบด้านบน แล้วกด **Run**
2. อนุญาตสิทธิ์ (Review permissions → เลือกบัญชี → Allow)
3. ระบบจะสร้างสเปรดชีต **"Pallet Hub Database"** ใน Google Drive พร้อมชีตทุกตาราง ประเภทพาเลท 4 แบบ และรายชื่อฝ่าย 8 ฝ่าย
   แล้วบันทึก ID ไว้ใน Script Properties (`PALLET_SPREADSHEET_ID`) — ดูลิงก์ได้ใน Execution log
4. รัน `setupSystem` ซ้ำได้ ระบบจะไม่สร้างสเปรดชีตใหม่ (ใช้ของเดิม และเติมชีต/ข้อมูลตั้งต้นที่ขาดเท่านั้น)

### 3. ตั้งรหัสผ่านสำหรับบันทึกข้อมูล
1. **Project Settings** → เลื่อนลงไปที่ **Script properties** → **Add script property**
2. Property: `PALLET_ACTION_PASSWORD`  Value: รหัสผ่านที่ต้องการ → **Save script properties**
3. รหัสผ่านนี้เก็บใน Script Properties เท่านั้น **ไม่อยู่ในซอร์สโค้ด** ถ้ายังไม่ตั้ง ระบบจะไม่ยอมให้บันทึกข้อมูลใด ๆ

### 4. Deploy เป็น Web App
1. กด **Deploy → New deployment** → เลือกประเภท **Web app**
2. **Execute as:** `Me`  **Who has access:** `Anyone`
3. กด **Deploy** แล้วคัดลอก **Web app URL** (ลงท้ายด้วย `/exec`)
4. เมื่อแก้ `Code.gs` ภายหลัง ให้ **Deploy → Manage deployments → แก้ไข (ดินสอ) → Version: New version → Deploy** เพื่อให้ URL เดิมใช้โค้ดใหม่

### 5. ตั้งค่าหน้าเว็บ
1. เปิดไฟล์ `docs/config.js` แล้วใส่ URL:
   ```js
   window.PALLET_CONFIG = { apiUrl: "https://script.google.com/macros/s/xxxxxxxx/exec" };
   ```
2. อัปโหลดขึ้น GitHub แล้วเปิด GitHub Pages: **Settings → Pages → Branch: `main` / Folder: `/docs`**
3. ถ้ายังไม่ได้ใส่ `apiUrl` หน้าเว็บจะแสดงแถบแจ้งเตือน "ยังไม่ได้ตั้งค่า apiUrl"

## การใช้งาน

- **ดูข้อมูล** (แดชบอร์ด คงเหลือ ประวัติ Log งานซ่อม) ทุกคนเปิดดูได้โดยไม่ต้องใส่รหัส
- **บันทึก/แก้ไข** (รับเข้า เบิกจ่าย รับคืน แจ้งชำรุด ส่งซ่อม ซ่อมเสร็จ ตัดจำหน่าย เพิ่ม/ลบฝ่าย) ระบบจะถามรหัสผ่าน **ครั้งเดียวต่อการเปิดหน้าเว็บ**
  (เก็บในหน่วยความจำเท่านั้น ปิด/รีเฟรชหน้าแล้วต้องใส่ใหม่) ถ้ารหัสผิดหรือถูกเปลี่ยน ระบบจะถามใหม่ในครั้งถัดไป
- ใส่รหัสผิดเกิน 10 ครั้งภายใน 15 นาที ระบบจะล็อกการบันทึกชั่วคราว 15 นาที
- **ชื่อผู้ใช้**: กดมุมขวาบนเพื่อตั้งชื่อ (เก็บในเบราว์เซอร์) ระบบจะใส่ชื่อในทุกรายการและในบันทึกประวัติ (Log)
- **ส่งออก Excel**: ปุ่ม Excel ในหน้า ประวัติเคลื่อนไหว และ บันทึกประวัติ (Log) จะดาวน์โหลดไฟล์ CSV (UTF-8 BOM เปิดใน Excel ได้ภาษาไทยถูกต้อง)
- ช่อง IP ใน Log จะแสดงเป็น `web` เพราะ Apps Script ไม่สามารถเห็น IP ของผู้ใช้

## สถานะพาเลท

```
ภายนอก --รับเข้า--> พร้อมใช้ --เบิกจ่าย(ฝ่าย)--> เบิกไปใช้งาน --รับคืน--> พร้อมใช้
พร้อมใช้ / รับคืนสภาพชำรุด --> ชำรุด --ส่งซ่อม--> กำลังซ่อม --ซ่อมเสร็จ--> พร้อมใช้
ชำรุด / กำลังซ่อม --ซ่อมไม่ได้--> ตัดจำหน่าย
```

## API (สำหรับนักพัฒนา)

- อ่านข้อมูล: `GET <apiUrl>?action=<action>&...` → `{ ok: true, data }` หรือ `{ ok: false, error }`
  - `bootstrap`, `dashboard` (`days`=7/14/30), `repairs`, `history` / `export` (`from`, `to`, `type`, `act`, `dept`, `q`), `logs` / `logs_export` (`from`, `to`, `cat`, `q`)
- บันทึกข้อมูล: `POST <apiUrl>` header `Content-Type: text/plain;charset=UTF-8` body JSON `{ action, password, actor, ... }`
  - `receive`, `issue`, `return`, `damage`, `repair_start`, `repair_done`, `scrap`, `dept_save`, `dept_delete`, `verifyPassword`
- ทุกการบันทึกทำภายใต้ `LockService` (ป้องกันเลขที่เอกสารซ้ำเมื่อบันทึกพร้อมกัน)

## ทดสอบ

```
node test/gas-mock.js
```
ชุดทดสอบจำลอง SpreadsheetApp / PropertiesService / CacheService / LockService / Utilities / ContentService
แล้วรัน `setupSystem()` และทดสอบทุก action ผ่าน `doGet` / `doPost` (รวมรหัสผ่านผิด/ถูก การล็อก และการเปลี่ยนสถานะทุกแบบ)

## หมายเหตุ

- ฟอนต์ ไอคอน และกราฟโหลดจาก CDN (Google Fonts, cdnjs) ต้องต่ออินเทอร์เน็ต
- Google Sheets ช้ากว่า MySQL แต่ละคำสั่งอาจใช้เวลา 1–3 วินาที
- ห้ามแก้ไขหัวตาราง (แถวที่ 1) ในชีต — แก้ข้อมูลในแถวอื่นได้ แต่ควรทำผ่านหน้าเว็บเพื่อให้มี Log
