โค้ดชุดนี้วางโครงสร้างแยกส่วนระหว่างการดึงข้อมูล (Ingestion/Scraping) ใน `main.py` และการแปลงข้อมูล (Transformation) ใน `combine_data.py` ได้ชัดเจน มีการจัดการเรื่อง ASP.NET ViewState, การทำ Idempotency เบื้องต้น (เช็คไฟล์ซ้ำก่อนดาวน์โหลด), และการเขียนเอกสารประกอบ (`README.md`, `PROCESS.md`) ที่เข้าใจง่าย

ในมุมมองของ Developer และ Data Engineer มีจุดที่สามารถปรับปรุงเพื่อให้โค้ดมีความยืดหยุ่น ทนทานต่อข้อผิดพลาด (Resilience) และนำไปต่อยอดกับเครื่องมือวิเคราะห์ข้อมูลได้ดียิ่งขึ้น ดังนี้ครับ

---

### 1. รีวิวไฟล์ `main.py` (Data Scraper)

- **ความเสี่ยงเรื่อง ASP.NET Token หมดอายุ หรือไม่มีอยู่จริง:**
- โค้ดใช้ `soup.find(id='__VIEWSTATE')['value']` โดยตรง หากวันไหนเว็บ EPPO ปิดปรับปรุงหรือเปลี่ยนโครงสร้าง HTML จนหา ID ไม่เจอ จะเกิด `TypeError: 'NoneType' object is not subscriptable` ทันที

- การใช้ค่า `__VIEWSTATE` เดิมซ้ำไปเรื่อยๆ ในลูป หากดึงข้อมูลย้อนหลังหลายร้อยวัน เซิร์ฟเวอร์อาจตัด Session หรือเมื่อเว็บโหลดหน้าซ้ำกรณีไม่มีข้อมูลของวันนั้น ค่า ViewState อาจจะเปลี่ยนไป แนะนำให้อัปเดตค่า ViewState ใหม่ทุกครั้งที่ได้ Response เป็นหน้า HTML กลับมา

- **ขาดการกำหนด `timeout` ใน HTTP Request:**
- ทั้ง `session.get(url)` และ `session.post(url, data=data)` ไม่ได้ระบุ `timeout` หากเซิร์ฟเวอร์ปลายทางค้าง สคริปต์จะค้างไปตลอดกาล ควรเพิ่ม `timeout=15` หรือ `timeout=30` เสมอ

- **Hardcoded Path และ Date:**
- การกำหนด `output_directory = r'D:\Project\Thai-Oil-Price\Data'` ทำให้ย้ายโค้ดไปรันบนเครื่องอื่นหรือบน Linux/WSL ไม่ได้ แนะนำให้ใช้ `pathlib.Path` อ้างอิงจากตำแหน่งไฟล์สคริปต์ (Relative Path) และใช้ `argparse` เพื่อรับค่า `start_date` และ `end_date` ผ่าน Command Line แทนการแก้โค้ดทุกครั้ง

---

### 2. รีวิวไฟล์ `combine_data.py` (Data Parser & Transformer)

- **การใช้ Bare `except:`:**
- ในบรรทัดที่แปลงค่าเป็นทศนิยม มีการใช้ `except:` เปล่าๆ:

```python
try:
    row_dict[comp] = float(val)
except:
    row_dict[comp] = None

```

ควรระบุ Exception ให้ชัดเจนเป็น `except (ValueError, TypeError):` เพื่อไม่ให้ไปดักจับ System Exit หรือ KeyboardInterrupt โดยไม่ตั้งใจ

- **ความเปราะบางของ `clean_product_name`:**
- การใช้ `if ... in raw_name` เรียงต่อกันหลายบรรทัด ทำงานได้ถูกต้องแต่ดูแลรักษายาก หากในอนาคตมีชื่อน้ำมันใหม่เพิ่มเข้ามา โค้ดส่วนนี้จะยาวขึ้นเรื่อยๆ สามารถเปลี่ยนไปใช้ Dictionary Mapping เพื่อให้อ่านง่ายและแก้ไขสะดวกขึ้น

- **การพึ่งพาตำแหน่งคอลัมน์แบบตายตัว (Positional Indexing):**
- โค้ดกำหนดรายชื่อบริษัทไว้ตายตัว `companies = ['PTT', 'BCP', 'Shell', ...]` แล้ววนลูปดึงตาม `col_idx = i + 1` หากวันไหนไฟล์ Excel ของ EPPO มีการสลับคอลัมน์ หรือเพิ่ม/ลดบริษัท (เช่น กรณีควบรวมกิจการปั๊มน้ำมัน) ข้อมูลราคาจะถูกใส่ผิดบริษัททันทีโดยไม่แจ้งเตือน หากเป็นไปได้ควรเพิ่ม Logic ตรวจสอบแถวที่เป็น Header ของตารางก่อนดึงข้อมูล

- **รูปแบบไฟล์ `.xls` ของเว็บราชการไทย:**
- โค้ดใช้ `pd.read_excel(file_path, engine='xlrd')` ซึ่งเป็นวิธีที่ถูกต้องสำหรับไฟล์ `.xls` รุ่นเก่า (BIFF8) แต่หลายครั้งระบบเว็บ ASP.NET มักจะส่งออกไฟล์ที่เป็น HTML Table แต่ตั้งนามสกุลหลอกไว้เป็น `.xls` หากเจอไฟล์ลักษณะนั้น `xlrd` จะโยน Error ทันที สามารถเขียน Fallback ให้ลองใช้ `pd.read_html()` หาก `read_excel()` ล้มเหลวได้

- **ปัญหา Encoding เมื่อเปิดไฟล์ CSV ด้วย Excel:**
- การบันทึกด้วย `encoding='utf-8'` หากยังมีชื่อผลิตภัณฑ์บางตัวที่ไม่เข้าเงื่อนไขภาษาอังกฤษและหลุดออกมาเป็นภาษาไทย เวลาเปิดไฟล์ CSV ในโปรแกรม Microsoft Excel บน Windows ตัวอักษรภาษาไทยจะเพี้ยน (Mojibake) แนะนำให้ใช้ `encoding='utf-8-sig'` แทน

---

### 3. รีวิวไฟล์เอกสาร (`README.md` และ `PROCESS.md`)

- **ความคลาดเคลื่อนของคำศัพท์ Data Modeling (Wide vs. Long Format):**
- ใน `README.md` และ `PROCESS.md` ระบุว่าข้อมูลถูกจัดให้อยู่ในรูปแบบ **Database Long Form** โดยมีชื่อบริษัท (PTT, BCP, Shell ฯลฯ) เป็นหัวคอลัมน์

- ในทางเทคนิค โครงสร้างที่นำชื่อบริษัทเรียงเป็นคอลัมน์แนวนอนแบบนี้เรียกว่า **Wide Format (ตารางแบบกว้าง)** ครับ
- หากต้องการทำเป็น **Long Format (หรือ Unpivoted / Tidy Data)** จริงๆ เพื่อนำไปเข้าฐานข้อมูลหรือทำ Data Visualization (เช่น Power BI หรือ Looker Studio) ได้ยืดหยุ่นที่สุด โครงสร้างคอลัมน์ควรมีแค่ 4 คอลัมน์ คือ: `Date`, `Product`, `Company`, และ `Price` ซึ่งสามารถทำได้ง่ายๆ ก่อน Export ด้วยคำสั่ง `df.melt()` ของ Pandas:

```python
long_df = final_df.melt(
    id_vars=['Date', 'Product'],
    var_name='Company',
    value_name='Price'
)

```

- **การจัดการ Dependencies:**
- แทนที่จะให้รัน `pip install ...` แบบพิมพ์รายชื่อใน `README.md` แนะนำให้สร้างไฟล์ `requirements.txt` ไว้ในโปรเจกต์เพื่อให้ติดตั้งผ่าน `pip install -r requirements.txt` และล็อกเวอร์ชันของไลบรารีได้แม่นยำขึ้น

---

### 4. ตัวอย่างการ Refactor โค้ดให้เป็นมาตรฐาน Production

#### ตัวอย่างการปรับปรุง `main.py`

```python
import os
import time
import logging
from pathlib import Path
from datetime import date, timedelta
import requests
from bs4 import BeautifulSoup

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = BASE_DIR / 'Data'
EPPO_URL = 'https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx'

def extract_tokens(html_text):
    soup = BeautifulSoup(html_text, 'html.parser')
    try:
        return {
            '__VIEWSTATE': soup.find(id='__VIEWSTATE')['value'],
            '__VIEWSTATEGENERATOR': soup.find(id='__VIEWSTATEGENERATOR')['value'],
            '__EVENTVALIDATION': soup.find(id='__EVENTVALIDATION')['value'],
        }
    except (TypeError, KeyError) as e:
        raise ValueError(f"ไม่สามารถดึง ASP.NET tokens จากหน้าเว็บได้: {e}")

def download_eppo_oil_prices(start_date: date, end_date: date, output_dir: Path = DEFAULT_OUTPUT_DIR):
    output_dir.mkdir(parents=True, exist_ok=True)

    with requests.Session() as session:
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

        logging.info("Fetching initial page to get ViewState...")
        try:
            response = session.get(EPPO_URL, timeout=15)
            response.raise_for_status()
            tokens = extract_tokens(response.text)
        except Exception as e:
            logging.error(f"Failed to initialize session: {e}")
            return

        current_date = start_date
        while current_date <= end_date:
            date_str = current_date.strftime('%d/%m/%Y')
            filename = f"EPPO_RetailOilPrice_on_{current_date.strftime('%Y%m%d')}.xls"
            filepath = output_dir / filename

            if filepath.exists():
                logging.info(f"File {filename} already exists. Skipping...")
                current_date += timedelta(days=1)
                continue

            logging.info(f"Downloading data for {date_str}...")
            payload = {
                **tokens,
                'TbxToDate': date_str,
                'BtnGenerate.x': '10',
                'BtnGenerate.y': '10'
            }

            try:
                res = session.post(EPPO_URL, data=payload, timeout=20)
                res.raise_for_status()

                content_type = res.headers.get('Content-Type', '')
                if 'application/vnd.ms-excel' in content_type or res.headers.get('Content-Disposition'):
                    filepath.write_bytes(res.content)
                    logging.info(f"Successfully downloaded {filename}")
                else:
                    logging.warning(f"No Excel file returned for {date_str}.")
                    # อัปเดต token ใหม่เผื่อเซิร์ฟเวอร์สร้าง ViewState ใหม่ในหน้า HTML ที่ส่งกลับมา
                    tokens = extract_tokens(res.text)

            except requests.exceptions.RequestException as e:
                logging.error(f"Error downloading for {date_str}: {e}")
            except ValueError as e:
                logging.error(f"Token parsing error on {date_str}: {e}")
                break

            time.sleep(1)
            current_date += timedelta(days=1)

if __name__ == "__main__":
    start = date(2026, 7, 12)
    end = date(2026, 10, 27)
    download_eppo_oil_prices(start, end)

```

#### ตัวอย่างการปรับปรุงฟังก์ชันใน `combine_data.py`

```python
from pathlib import Path
import re
import pandas as pd

PRODUCT_MAPPING = {
    'Gasohol 95-E10': 'Gasohol 95-E10',
    'Gasohol 95-E20': 'Gasohol 95-E20',
    'Gasohol 95-E85': 'Gasohol 95-E85',
    'Gasohol 91-E10': 'Gasohol 91-E10',
    'ULG 95 RON': 'ULG 95 RON',
    'HSD B7': 'HSD B7',
    'HSD B10': 'HSD B10',
    'HSD B20': 'HSD B20',
}

def clean_product_name(raw_name: str) -> str:
    cleaned = str(raw_name).replace('\xa0', ' ').strip()

    for key, val in PRODUCT_MAPPING.items():
        if key in cleaned:
            return val

    if 'พรีเมียม' in cleaned or 'Premium' in cleaned:
        if any(k in cleaned for k in ('แก๊สโซฮอล์ 95', 'แก๊สโซฮอล ออกเทน 95', 'Gasohol 95')):
            return 'Gasohol 95 Premium'
        if any(k in cleaned for k in ('แก๊สโซฮอล ออกเทน 99', 'Gasohol 99')):
            return 'Gasohol 99 Premium'
        if any(k in cleaned for k in ('เบนซิน', 'ULG')):
            return 'ULG 95 Premium'
        if any(k in cleaned for k in ('ดีเซล', 'HSD')):
            return 'HSD Premium'

    match = re.search(r'\((.*?)\)', cleaned)
    return match.group(1).strip() if match else cleaned

```
