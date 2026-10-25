"""
สคริปต์อ่าน PDF ด้วย Pymupdf (รองรับภาษาไทย)
"""
import pymupdf  # PyMuPDF v2.0+
from pathlib import Path
import json


from core.utils.cleaners import clean_thai_ocr_text


def read_pdf_with_path(pdf_path: str):
    """
    อ่านข้อมูลจากไฟล์ PDF โดยใช้ path

    Args:
        pdf_path: เส้นทางเต็มของไฟล์ PDF (เช่น C:\\...\\file.pdf)

    Returns:
        ข้อความที่อ่านได้จากทุกหน้าเป็น dictionary
    """
    try:
        print(f"กำลังเปิดไฟล์: {pdf_path}")

        # เปิดไฟล์ด้วย pymupdf
        doc = pymupdf.open(pdf_path)
        extracted_data = {}

        for i, page in enumerate(doc):
            # Extract ข้อความธรรมดา
            text = clean_thai_ocr_text(page.get_text("text") or "")

            # แปลงตารางถ้ามี
            tables = page.find_tables()

            # เก็บข้อมูลทุกหน้า
            extracted_data[f'Page {i+1}'] = {
                'raw_text': text.strip(),
                'length': len(text),
                'has_table': len(tables) > 0
            }

        doc.close()

        print(f"✅ อ่านเสร็จแล้ว! พบ {len(extracted_data)} หน้า")
        return extracted_data

    except FileNotFoundError as e:
        print(f"❌ ไม่พบไฟล์: {pdf_path}")
        return {}
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาด: {e}")
        import traceback
        traceback.print_exc()
        return {}


def save_to_json(pdf_path: str, json_file: str = "extracted_data.json"):
    """
    อ่าน PDF และบันทึกเป็นไฟล์ JSON
    """
    data = read_pdf_with_path(pdf_path)
    if data:
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"✅ บันทึกข้อมูลลงใน {json_file} เรียบร้อย!")


def print_first_page(pdf_path: str):
    """พิมพ์เฉพาะหน้าแรกให้ดู"""
    data = read_pdf_with_path(pdf_path)
    if data:
        key, value = next(iter(data.items()))  # หน้าแรก
        print(f"\n--- {key} ---")
        raw_text = value.get("raw_text", "")
        print(raw_text[:1000] + "..." if len(raw_text) > 1000 else raw_text)


# ==================== ตัวอย่างการใช้งาน ====================
if __name__ == "__main__":
    # ตั้งชื่อไฟล์ PDF (พิมพ์ path เต็มเข้ามาเลย!)
    pdf_path = input("พิมพ์ Path ไฟล์ PDF เต็ม: ").strip().strip('"').strip("'")  # ลบเครื่องหมายคำพูด

    print("\n" + "="*80)

    if not pdf_path:
        print("❌ ต้องระบุ path ไฟล์!")
    else:
        try:
            # อ่านและแสดงหน้าแรกก่อน (เพื่อให้แน่ใจว่าอ่านได้ถูกต้อง)
            print_first_page(pdf_path)

            # ถามว่าต้องการบันทึก JSON หรือไม่
            choice = input("\nต้องการบันทึกเป็น JSON หรือไม่? (y/n): ").strip().lower()
            if choice in ['y', 'yes']:
                json_file = "output.json"
                save_to_json(pdf_path, json_file)

        except Exception as e:
            print(f"\n❌ เกิดข้อผิดพลาดขณะประมวลผล: {e}")