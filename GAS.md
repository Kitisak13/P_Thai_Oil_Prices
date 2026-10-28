# คู่มือการติดตั้งระบบ Auto-Sync บน Google Apps Script (ร่วมกับ ScraperAPI)

ระบบนี้ทำหน้าที่ดึงข้อมูลราคาน้ำมันรายวันจากเว็บไซต์ EPPO เข้าสู่ Google Sheet แบบ Wide Format อัตโนมัติวันละ 2 รอบ (08:00 และ 16:00 น. เวลาประเทศไทย) โดยทำงานบน Google Cloud 100% (ไม่ต้องเปิดคอมพิวเตอร์ทิ้งไว้)

---

## 1. การตั้งค่า Script Properties (ตัวแปรความปลอดภัย)

ให้เข้าไปที่เมนู **Project Settings (รูปเฟือง) > Script Properties > Edit script properties** ในหน้าต่าง Apps Script:

| Property Name | จำเป็นหรือไม่ | ตัวอย่างค่า | คำอธิบาย |
| :--- | :--- | :--- | :--- |
| `SCRAPERAPI_KEY` | **จำเป็น** | `a1b2c3d4e5f6...` | API Key ที่ได้จาก [scraperapi.com](https://www.scraperapi.com) |
| `TARGET_SHEET_NAME` | แนะนำ | `Sheet1` หรือ `Combined_Oil_Prices` | ชื่อแท็บใน Google Sheet ที่เก็บข้อมูล |
| `TIMEZONE` | ไม่บังคับ | `Asia/Bangkok` | ไทม์โซน (ค่าเริ่มต้น: `Asia/Bangkok`) |
| `EPPO_URL` | ไม่บังคับ | URL ของ EPPO | URL ต้นทาง (มีค่าเริ่มต้นในโค้ดแล้ว) |
| `LINE_NOTIFY_TOKEN` | ทางเลือก | `...` | Token แจ้งเตือนเข้า LINE เมื่ออัปเดตข้อมูลสำเร็จ |

---

## 2. ขั้นตอนการติดตั้งบน Google Sheet

### ขั้นตอนที่ 1: เปิด Apps Script
1. เปิดไฟล์ Google Sheet ฐานข้อมูลราคาน้ำมันของคุณ
2. ไปที่เมนู **ส่วนขยาย (Extensions) > Apps Script**

### ขั้นตอนที่ 2: วางโค้ด [Code.gs](Code.gs)
1. คัดลอกโค้ดทั้งหมดจากไฟล์ `Code.gs` ไปวางแทนที่โค้ดเดิมใน Apps Script แล้วกด **บันทึก (Ctrl + S)**

### ขั้นตอนที่ 3: เปิดใช้งานบริการ Drive API (จำเป็น)
1. ที่แถบเมนูด้านซ้าย ในหัวข้อ **บริการ (Services)** ให้คลิกเครื่องหมาย **`+`**
2. เลือก **Drive API**
3. กด **เพิ่ม (Add)**

### ขั้นตอนที่ 4: ใส่ `SCRAPERAPI_KEY` ใน Script Properties
1. คลิกที่ **การตั้งค่าโครงการ (Project Settings - รูปเฟือง)** ด้านซ้าย
2. เลื่อนลงมาที่หัวข้อ **คุณสมบัติของสคริปต์ (Script Properties)**
3. คลิก **แก้ไขคุณสมบัติของสคริปต์ (Edit script properties)**
4. เพิ่ม Property:
   - **Property:** `SCRAPERAPI_KEY`
   - **Value:** *(นำ API Key จาก scraperapi.com มาวาง)*
   - *(หากชื่อแท็บไม่ใช่แท็บแรก ให้เพิ่ม `TARGET_SHEET_NAME` ด้วย)*
5. กด **บันทึกคุณสมบัติของสคริปต์ (Save script properties)**

### ขั้นตอนที่ 5: ทดสอบการเชื่อมต่อ (Diagnostic Test)
1. กลับมาที่หน้า **โปรแกรมแก้ไข (Editor - รูป `< >`)**
2. เลือกฟังก์ชัน **`diagnoseEppoConnection`** แล้วกด **เรียกใช้ (Run)**
3. ดูที่แถบ **บันทึกการประมวลผล (Execution log)**:
   - หากขึ้น `✅ SUCCESS! Bypassed Incapsula successfully...` แสดงว่าผ่านฉลุย!
4. จากนั้นเลือกฟังก์ชัน **`testRunToday`** แล้วกด **เรียกใช้ (Run)** เพื่อทดสอบดึงและบันทึกข้อมูลของวันนี้ลง Google Sheet

### ขั้นตอนที่ 6: เปิดระบบตั้งเวลาอัตโนมัติ (Trigger Setup)
1. เลือกฟังก์ชัน **`setupDailyTriggers`** แล้วกด **เรียกใช้ (Run)** 1 ครั้ง
2. ระบบจะสร้าง Trigger รันอัตโนมัติวันละ 2 รอบ (08:00 และ 16:00 น. เวลาประเทศไทย) ให้อัตโนมัติ
