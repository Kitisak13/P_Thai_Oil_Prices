# กระบวนการทำงานและการใช้งานระบบ (Process & Workflow)

กระบวนการทำงานของระบบแบ่งออกเป็น 2 ส่วนหลัก ได้แก่ การดึงข้อมูลเข้าสู่ระบบ (Ingestion) และการประมวลผลจัดรูปแบบข้อมูล (Transformation)

---

## 1. การดึงข้อมูลราคาน้ำมัน (Data Ingestion)

สคริปต์ [main.py](main.py) ทำหน้าที่เชื่อมต่อไปยังเว็บไซต์ EPPO และดาวน์โหลดไฟล์ Excel รายวัน

### 1.1 การใช้งานแบบ Incremental (แนะนำสำหรับการใช้งานประจำวัน)
ดึงเฉพาะข้อมูลของวันใหม่ที่ยังไม่มีในเครื่อง:
```bash
python main.py
```
> ระบบจะค้นหาไฟล์ล่าสุดใน `data/raw/` หรือ `Data/` แล้วเริ่มดึงต่อจากวันล่าสุดจนถึงวันปัจจุบันโดยอัตโนมัติ

### 1.2 การดึงข้อมูลย้อนหลังตามช่วงเวลาที่กำหนด
```bash
python main.py --start-date 2026-01-01 --end-date 2026-10-28
```

### 1.3 ตัวเลือกทั้งหมดของ `main.py`
| พารามิเตอร์ | ค่าเริ่มต้น | คำอธิบาย |
| :--- | :--- | :--- |
| `--start-date` | Auto-detect | วันที่เริ่มต้น (รูปแบบ `YYYY-MM-DD` หรือ `YYYYMMDD`) |
| `--end-date` | วันปัจจุบัน | วันที่สิ้นสุด (รูปแบบ `YYYY-MM-DD` หรือ `YYYYMMDD`) |
| `--output-dir` | `data/raw` หรือ `Data` | โฟลเดอร์ปลายทางสำหรับบันทึกไฟล์ `.xls` |
| `--incremental`, `-i` | False | บังคับให้เริ่มดึงต่อจากวันล่าสุด + 1 วัน |
| `--timeout` | 20 | ระยะเวลา Timeout ของ HTTP Request (วินาที) |
| `--delay` | 1.0 | หน่วงเวลาระหว่าง Request เพื่อไม่ให้เซิร์ฟเวอร์โหลดเกินไป (วินาที) |

---

## 2. การรวมและจัดโครงสร้างข้อมูล (Data Transformation)

สคริปต์ [combine_data.py](combine_data.py) ทำหน้าที่อ่านไฟล์ `.xls` ดิบทั้งหมด คลีนชื่อผลิตภัณฑ์ ตรวจสอบ Header คอลัมน์ และบันทึกผลลัพธ์

### 2.1 รันการประมวลผลข้อมูลทั้งหมด
```bash
python combine_data.py
```
เมื่อคำสั่งทำงานเสร็จ จะสร้างไฟล์ผลลัพธ์ในโฟลเดอร์ `data/processed/`:
1. `Combined_Oil_Prices_Wide.csv` - ตารางแบบกว้าง (บริษัทเป็นคอลัมน์)
2. `Combined_Oil_Prices_Wide.parquet` - ตารางแบบกว้างแบบ Parquet
3. `Combined_Oil_Prices_Long.csv` - ตาราง Tidy Data แบบยาว (`Date`, `Product`, `Company`, `Price`)
4. `Combined_Oil_Prices_Long.parquet` - ตาราง Tidy Data แบบ Parquet
5. `Combined_Oil_Prices.csv` (ที่ Root Directory เพื่อรองรับ Workflow เดิม)

### 2.2 ตัวเลือกทั้งหมดของ `combine_data.py`
| พารามิเตอร์ | ค่าเริ่มต้น | คำอธิบาย |
| :--- | :--- | :--- |
| `--input-dir` | Auto | โฟลเดอร์ที่เก็บไฟล์ `.xls` ดิบ (ค้นหา `data/raw` แล้วตามด้วย `Data`) |
| `--output-dir` | `data/processed` | โฟลเดอร์ปลายทางสำหรับไฟล์ผลลัพธ์ |
| `--format` | `both` | รูปแบบข้อมูลที่ต้องการส่งออก (`wide`, `long`, หรือ `both`) |
| `--no-parquet` | False | ปิดการบันทึกไฟล์ `.parquet` |
| `--no-legacy-csv` | False | ปิดการบันทึกไฟล์ `Combined_Oil_Prices.csv` ที่ Root Directory |

---

## 3. ตัวอย่างการนำข้อมูลไปใช้งานต่อ (Data Consumption)

### ตัวอย่างการอ่านไฟล์ด้วย Pandas
```python
import pandas as pd

# อ่าน Long Format (เหมาะกับงานวิเคราะห์/Power BI)
df_long = pd.read_parquet('data/processed/Combined_Oil_Prices_Long.parquet')
print(df_long.head())

# คำนวณราคาน้ำมันเฉลี่ยรายเดือนของแต่ละค่าย
df_long['Date'] = pd.to_datetime(df_long['Date'])
monthly_avg = df_long.groupby([df_long['Date'].dt.to_period('M'), 'Product', 'Company'])['Price'].mean()
```
