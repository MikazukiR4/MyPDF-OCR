import os
import sys
import base64
import urllib.request
import urllib.error
import json
from pathlib import Path

# ฟังก์ชันทดสอบการอ่านลายมือด้วย Gemini แยกเดี่ยว
def test_gemini_handwriting(image_path: str, api_key: str, model: str = "gemini-2.5-flash") -> str:
    img_bytes = Path(image_path).read_bytes()
    encoded_image = base64.b64encode(img_bytes).decode("ascii")
    
    prompt = (
        "คุณคือผู้เชี่ยวชาญอ่านลายมือภาษาไทยจากเอกสารสแกน\n"
        "อ่านจากภาพโดยตรง ทีละบรรทัดจากบนลงล่าง ถอดข้อความตามที่เขียนจริง\n"
        "- เลือกตัวอักษรที่สอดคล้องกับบริบทของคำและประโยค แต่ห้ามเติมข้อความที่ไม่มีในภาพ\n"
        "- ตรวจสระบน/ล่าง วรรณยุกต์ ไม้ไต่คู้ การันต์ ให้ครบ เพราะลายมือมักเขียนเล็กและชิดกัน\n"
        "- ชื่อคน นามสกุล ที่อยู่ เลขที่ ตัวเลข ให้คัดลอกตามที่เห็นทีละตัว ห้ามแก้ให้เป็นคำที่คุ้นเคย\n"
        "- คงเลขไทยหรือเลขอารบิกตามที่เขียน\n"
        "- ส่วนที่อ่านไม่ออกจริง ๆ ให้ใส่ [?] แทนคำนั้น\n"
        "- ส่งคืนเฉพาะข้อความ ไม่อธิบายลายเซ็น/รูป ไม่ใช้ Markdown\n"
    )
    
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/png", "data": encoded_image}}
            ]
        }],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 8192
        }
    }
    
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST"
    )
    
    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            res = json.loads(response.read().decode("utf-8"))
            return "".join(p["text"] for p in res["candidates"][0]["content"]["parts"] if p.get("text")).strip()
    except Exception as e:
        return f"ERROR: {e}"

if __name__ == "__main__":
    print("สคริปต์ทดสอบ Gemini Handwriting พร้อมใช้งานแล้ว")
