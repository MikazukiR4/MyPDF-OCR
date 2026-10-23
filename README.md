# MyPDF-style Thai OCR

Streamlit OCR สำหรับ PDF สแกน โดยทำงานทีละหน้าแบบเดียวกับ workflow ของ MyPDF:

`PDF → render high-DPI image → Thai OCR → normalize Thai text → TXT`

ค่าเริ่มต้นใช้ Tesseract ภาษาไทยในเครื่อง พร้อมปรับภาพความละเอียดสูงและทำความสะอาดข้อความที่เกิดจากสระ/วรรณยุกต์ลอยของเอกสารสแกน

หากตั้งค่า `TYPHOON_OCR_API_KEY` ใน `.env` หรือ `.env.typhoon` จะเปิดตัวเลือก **Typhoon OCR อ่านเอกสารราชการไทยทั้งหน้า (แนะนำ)** อัตโนมัติ ซึ่งเหมาะกับเอกสารภาษาไทย ลายมือ และเลขไทยมากกว่า Tesseract แต่ภาพเอกสารจะถูกส่งไปยัง Typhoon OCR API

## Local

ต้องติดตั้ง Tesseract และภาษาไทยก่อน จากนั้นรัน:

```powershell
pip install -r requirements.txt
streamlit run app.py
```

หากติดตั้ง Tesseract ไว้คนละตำแหน่ง ให้ตั้ง `TESSERACT_CMD` ใน `.env`

ตัวอย่างตัวแปรแวดล้อม:

```dotenv
TESSERACT_CMD=C:\\Program Files\\Tesseract-OCR\\tesseract.exe
TYPHOON_OCR_API_KEY=your-key
```
