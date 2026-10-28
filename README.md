# Thai Oil Price Scraper & Pipeline

ระบบดาวน์โหลดและประมวลผลข้อมูลราคาน้ำมันขายปลีกในเขตพื้นที่กรุงเทพมหานครและปริมณฑลแบบรายวัน จากเว็บไซต์ของสำนักงานนโยบายและแผนพลังงาน (EPPO - Energy Policy and Planning Office)

ระบบได้รับการออกแบบใหม่ตามแนวทาง Data Engineering และ DevOps Production Best Practices:
- **Resilience & Idempotency:** รองรับ Dynamic ASP.NET ViewState Token, Request Timeout, Auto-retry, และข้ามไฟล์ที่มีอยู่แล้ว
- **Incremental Sync:** ตรวจหาวันที่ล่าสุดที่มีในระบบโดยอัตโนมัติ และดึงเฉพาะข้อมูลวันใหม่จนถึงปัจจุบัน
- **Dynamic Header Parsing:** จับคู่ชื่อบริษัทน้ำมันจากตารางหัวเรื่องจริงอย่างยืดหยุ่น ป้องกันปัญหาการจัดลำดับคอลัมน์สลับตำแหน่ง
- **Dual Data Modeling (Wide & Long Format):** ส่งออกทั้งตารางสรุปแบบกว้าง (Wide Format) และ Tidy Data (Long Format) ที่พร้อมนำเข้าสู่ Power BI, Looker Studio หรือฐานข้อมูล SQL
- **High-Performance Storage:** รองรับการจัดเก็บทั้งไฟล์ CSV (`utf-8-sig`) และ Apache Parquet (`.parquet`)

---

## สถาปัตยกรรมและโครงสร้างไฟล์

```text
Thai-Oil-Price/
├── data/
│   ├── raw/                  # แหล่งจัดเก็บไฟล์ Excel ดิบรายวัน (.xls)
│   └── processed/            # ผลลัพธ์ข้อมูลที่รวมและ Clean แล้ว (.csv, .parquet)
│       ├── Combined_Oil_Prices_Wide.csv
│       ├── Combined_Oil_Prices_Wide.parquet
│       ├── Combined_Oil_Prices_Long.csv
│       └── Combined_Oil_Prices_Long.parquet
├── dev_reviews/              # เอกสาร Code Review & Architecture Review
├── main.py                   # Data Scraper (ดึงข้อมูลดิบจากเว็บ EPPO)
├── combine_data.py           # Data Transformer (Clean, Parse, จัดรูปแบบ Wide/Long)
├── requirements.txt          # Library Dependencies
├── Combined_Oil_Prices.csv   # ไฟล์ Wide Format ฉบับเข้ากันได้กับโค้ดเวอร์ชันเดิม
├── PROCESS.md                # คู่มือขั้นตอนการทำงานและคำสั่ง CLI
└── README.md                 # เอกสารแนะนำโครงการ
```

---

## การติดตั้ง (Installation)

แนะนำให้ติดตั้ง Dependencies ผ่าน `requirements.txt`:

```bash
python -m pip install -r requirements.txt
```

**รายการ Libraries หลัก:**
- `requests` - การเชื่อมต่อ HTTP Client
- `beautifulsoup4` - การแกะและดึง Form Tokens (__VIEWSTATE, __EVENTVALIDATION)
- `pandas` - การจัดการและแปลงโครงสร้างข้อมูล
- `xlrd` - เครื่องมืออ่านไฟล์ Excel 97-2003 (.xls)
- `lxml` - HTML Parser ความเร็วสูง
- `pyarrow` - สำหรับส่งออกข้อมูลเป็น Apache Parquet

---

## รูปแบบข้อมูลที่ส่งออก (Data Formats)

### 1. Wide Format (`Combined_Oil_Prices_Wide.csv` / `.parquet`)
ตารางแนวนอนที่นำชื่อบริษัทน้ำมันเป็นหัวคอลัมน์ เหมาะสำหรับการเปิดดูด้วยโปรแกรม Microsoft Excel หรือทำตารางสรุป
- **คอลัมน์:** `Date`, `Product`, `PTT`, `BCP`, `Shell`, `Esso`, `Chevron`, `IRPC`, `PT`, `Susco`, `Pure`, `SUSCO Dealers`

### 2. Long Format / Tidy Data (`Combined_Oil_Prices_Long.csv` / `.parquet`)
โครงสร้างแบบ Normalization (Unpivoted) ที่เหมาะสมที่สุดสำหรับงาน Data Warehouse, การเขียน Query SQL, และการทำ Visualizations (เช่น Power BI, Tableau, Looker Studio):
- **คอลัมน์:**
  - `Date`: วันที่ประกาศราคา (รูปแบบ `YYYY-MM-DD`)
  - `Product`: ชื่อผลิตภัณฑ์น้ำมันมาตรฐานสากล (เช่น `Gasohol 95-E10`, `ULG 95 RON`, `HSD B7`, `Gasohol 95 Premium`)
  - `Company`: ชื่อผู้ให้บริการปั๊มน้ำมัน (เช่น `PTT`, `Shell`, `BCP`, ฯลฯ)
  - `Price`: ราคาน้ำมันขายปลีก (บาท/ลิตร)

---

## วิธีใช้งานเบื้องต้น (Quickstart)

### 1. ดึงข้อมูลรายวันแบบ Incremental (รันอัตโนมัติหาวันล่าสุด)
```bash
python main.py
```
> ระบบจะสแกนหาวันที่ล่าสุดที่มีอยู่แล้ว และดาวน์โหลดต่อตั้งแต่วันถัดไปจนถึงวันนี้โดยอัตโนมัติ

### 2. รวมและแปลงข้อมูล
```bash
python combine_data.py
```
> ระบบจะอ่านไฟล์ Excel ทั้งหมด คลีนชื่อผลิตภัณฑ์ สกัดราคา และส่งออกไฟล์ Wide/Long Formats ทั้ง CSV และ Parquet ไปยังโฟลเดอร์ `data/processed/`

### 3. ตัวเลือกเพิ่มเติมผ่าน Command Line (CLI)
- **ดึงข้อมูลระบุช่วงวันที่:**
  ```bash
  python main.py --start-date 2026-01-01 --end-date 2026-10-28
  ```
- **ส่งออกเฉพาะ Long Format:**
  ```bash
  python combine_data.py --format long
  ```

ดูรายละเอียดคำสั่งและตัวเลือกทั้งหมดได้ในเอกสาร [PROCESS.md](PROCESS.md)
