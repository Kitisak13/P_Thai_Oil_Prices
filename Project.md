# ผมต้องการดึงไฟล์ราคาน้ำมันของกรมพลังงาน

ไฟล์ราคาน้ำมันจะอยู่ที่หน้าเว็บไซต์นี้ https://www.eppo.go.th/epposite/index.php/th/petroleum/price/oil-price?orders[publishUp]=publishUp&issearch=1

โดยตำแหน่งที่สามารถโหลดไฟล์ได้จะอยู่ในส่วนนี้ของเว็บไซต์
![alt text](image.png)

ผมพยายามเช็คตำแหน่งของจุดที่ดาวน์โหลดไฟล์ได้ดังนี้ (คุณอาจช่วยเช็คเพื่อความมั่นใจอีกครั้ง)

## ตำแหน่งของวันที่

//\*[@id="TbxToDate"]

<input name="TbxToDate" type="text" value="22/04/2026" id="TbxToDate" style="height:16px;width:80px;">

## ตำแหน่งของปุ่มโหลดไฟล์

//\*[@id="BtnGenerate"]

<input type="image" name="BtnGenerate" id="BtnGenerate" src="../picture/PicForm/buttonGenerate.gif" align="absmiddle" style="width:89px;border-width:0px;">

# ผมต้องการดาวน์โหลดไฟล์ราคาน้ำมันของทุกๆวัน ตั้งแต่วันที่ 1 มกราคม 2569 เป็นต้นมา

โดยทำซ้ำไปเรื่อยๆ ทุกๆ วัน โดยไฟล์จะเป็นไฟล์ Excel vesion 97-2003
โดยตัวโปรแกรมจะบันทึกไฟล์ไว้ที่ D:\Project\Thai-Oil-Price\Data

# ตัวอย่างไฟล์ราคาน้ำมันที่ดึงมา

D:\Project\Thai-Oil-Price\EPPO_RetailOilPrice_on_20260421.xls
