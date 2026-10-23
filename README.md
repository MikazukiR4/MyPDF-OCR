# MyPDF-style Thai OCR

Streamlit OCR สำหรับ PDF สแกน โดยทำงานทีละหน้าแบบเดียวกับ workflow ของ MyPDF:

`PDF → render high-DPI image → Thai OCR → normalize Thai text → TXT`

ค่าเริ่มต้นใช้ Tesseract ภาษาไทยในเครื่อง พร้อมปรับภาพความละเอียดสูงและทำความสะอาดข้อความที่เกิดจากสระ/วรรณยุกต์ลอยของเอกสารสแกน

หากตั้งค่า `TYPHOON_OCR_API_KEY` ใน `.env` หรือ `.env.typhoon` จะเปิดตัวเลือก **Typhoon OCR อ่านเอกสารราชการไทยทั้งหน้า (แนะนำ)** อัตโนมัติ ซึ่งเหมาะกับเอกสารภาษาไทย ลายมือ และเลขไทยมากกว่า Tesseract แต่ภาพเอกสารจะถูกส่งไปยัง Typhoon OCR API หากไม่มี key ระบบจะเลือก Tesseract ให้อัตโนมัติ

## Local

ต้องติดตั้ง Tesseract และภาษาไทยก่อน จากนั้นรัน:

```powershell
pip install -r requirements.txt
streamlit run app.py
```

หากติดตั้ง Tesseract ไว้คนละตำแหน่ง ให้ตั้ง `TESSERACT_CMD` ใน `.env`

คัดลอก `.env.example` เป็น `.env` สำหรับการรันในเครื่อง ตัวอย่างตัวแปรแวดล้อม:

```dotenv
TESSERACT_CMD=C:\\Program Files\\Tesseract-OCR\\tesseract.exe
TYPHOON_OCR_API_KEY=your-key
GEMINI_API_KEY=your-key
```

## Render

โปรเจกต์ใช้ Docker runtime โดยติดตั้ง Python และ Tesseract จาก `Dockerfile` แล้ว ตั้งค่าเฉพาะ secret ที่ต้องใช้ใน Render Dashboard → **Environment Variables**:

- `TYPHOON_OCR_API_KEY` (ไม่บังคับ)
- `GEMINI_API_KEY` (ไม่บังคับ)
- `GEMINI_MODEL` (ไม่บังคับ)

ห้าม commit API key จริงลง Git หาก key เคยถูกแสดงใน terminal, log หรือ source control ให้ revoke/rotate key เดิม แล้วตั้งค่า key ใหม่เฉพาะใน Render Environment Variables
