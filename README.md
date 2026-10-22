# THSummarizer

ระบบ Streamlit สำหรับอ่านเอกสารราชการไทยด้วยข้อความจาก PDF, EasyOCR และ ThaiTrOCR แล้วให้ Gemini เปรียบเทียบหลักฐานและสรุปผล

## Run locally

```powershell
pip install -r requirements.txt
$env:GEMINI_API_KEY = "your-api-key"
streamlit run app.py
```

## Deploy on Render

1. เปิด Render แล้วเลือก **New → Blueprint**
2. เลือก repository `MikazukiR4/THSummarizer`
3. Render จะอ่านค่าใน `render.yaml`
4. ตั้งค่า `GEMINI_API_KEY` ใน Environment ของ service
5. กด deploy

`GEMINI_API_KEY` เป็น secret ห้ามใส่ลงใน repository หรือ commit ลง Git
