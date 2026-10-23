"""
สคริปต์สำหรับ Extract ข้อความจากไฟล์ PDF อย่างแม่นยำ
ปรับปรุงการดึงข้อมูลตามลำดับการอ่าน (Reading Order) และรองรับภาษาไทย
"""
import fitz  # PyMuPDF
import re
from pathlib import Path
from datetime import datetime
import json

def clean_text(text: str) -> str:
    """
    ทำความสะอาดข้อความและจัดการช่องว่างให้เป็นระเบียบ
    """
    if not text:
        return ""
    # ลบช่องว่างที่ซ้ำซ้อน
    text = re.sub(r'[ \t]+', ' ', text)
    # จัดการการขึ้นบรรทัดใหม่
    text = re.sub(r'\n\s*\n', '\n\n', text)
    return text.strip()

def extract_text_from_pdf(pdf_path: str, output_format: str = "dict") -> dict | str | None:
    """
    อ่านข้อความจาก PDF โดยเน้นความแม่นยำของลำดับเนื้อหา (Sorting blocks)
    """
    if not Path(pdf_path).exists():
        print(f"❌ ไม่พบไฟล์: {pdf_path}")
        return None

    try:
        doc = fitz.open(pdf_path)
        extracted_data = {}

        for i, page in enumerate(doc):
            # ใช้ get_text("blocks") และตั้งค่า sort=True
            # เพื่อให้ไล่ลำดับข้อความจากบนลงล่าง ซ้ายไปขวา ตามโครงสร้างจริง
            blocks = page.get_text("blocks", sort=True)

            page_text_parts = []
            for block in blocks:
                # block[4] คือเนื้อหาข้อความในบล็อกนั้นๆ
                block_text = block[4].strip()
                if block_text:
                    page_text_parts.append(block_text)

            full_page_text = "\n".join(page_text_parts)

            extracted_data[f'page_{i+1}'] = {
                'raw_text': clean_text(full_page_text),
                'metadata': {
                    'page_number': i + 1,
                    'dimensions': {
                        'width': round(page.rect.width, 2),
                        'height': round(page.rect.height, 2)
                    },
                    'block_count': len(blocks)
                },
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }

        doc.close()

        if output_format == "dict":
            return extracted_data
        elif output_format == "json":
            return json.dumps(extracted_data, ensure_ascii=False, indent=2)
        elif output_format == "text":
            text_output = []
            for page_id, content in extracted_data.items():
                text_output.append(f"--- {page_id} ---\n{content['raw_text']}")
            return "\n\n".join(text_output)
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดขณะประมวลผล: {e}")
        return None

def save_refined_output(pdf_path: str, output_name: str = "refined_output.txt"):
    """
    ดึงข้อมูลและบันทึกลงไฟล์แยกตามโครงสร้างที่อ่านได้แม่นยำขึ้น
    """
    data = extract_text_from_pdf(pdf_path, output_format="dict")
    if not data:
        return False

    try:
        with open(output_name, 'w', encoding='utf-8') as f:
            for page_key, content in data.items():
                f.write(f"=== {page_key.upper()} ===\n")
                f.write(f"Size: {content['metadata']['dimensions']}\n")
                f.write("-" * 30 + "\n")
                f.write(content['raw_text'])
                f.write("\n\n")
        print(f"✅ บันทึกข้อมูลที่ปรับปรุงแล้วลงใน {output_name}")
        return True
    except Exception as e:
        print(f"❌ ไม่สามารถบันทึกไฟล์ได้: {e}")
        return False

def generate_visual_html(pdf_path: str):
    """
    สร้าง HTML Preview ที่แสดงโครงสร้าง Blocks เพื่อตรวจสอบความแม่นยำ
    """
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0: return

        page = doc[0]
        blocks = page.get_text("blocks", sort=True)

        html_blocks = ""
        for b in blocks:
            # สร้าง div จำลองตำแหน่ง block เพื่อดูการเรียงลำดับ
            html_blocks += f"<div style='border:1px solid #ccc; margin:5px; padding:10px; border-left: 5px solid #007bff;'>{b[4]}</div>"

        html_content = f"""
        <!DOCTYPE html>
        <html lang="th">
        <head>
            <meta charset="UTF-8">
            <title>PDF Accuracy Preview</title>
            <style>
                body {{ font-family: 'Sarabun', sans-serif; line-height: 1.6; padding: 40px; background: #f4f4f9; }}
                .container {{ max-width: 800px; margin: auto; background: white; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                h1 {{ color: #333; border-bottom: 2px solid #007bff; padding-bottom: 10px; }}
                .meta {{ color: #666; font-size: 0.9em; margin-bottom: 20px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>โครงสร้างข้อความหน้าแรก (Sorted Blocks)</h1>
                <div class="meta">ขนาดหน้า: {page.rect.width} x {page.rect.height} pt | จำนวน Blocks: {len(blocks)}</div>
                {html_blocks}
            </div>
        </body>
        </html>
        """

        output_html = "pdf_accuracy_check.html"
        with open(output_html, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"✅ สร้างไฟล์พรีวิวความแม่นยำที่: {output_html}")
    except Exception as e:
        print(f"❌ ไม่สามารถสร้าง HTML ได้: {e}")

if __name__ == "__main__":
    # ระบุชื่อไฟล์ PDF ที่ต้องการทดสอบ
    target_pdf = "input.pdf"

    print("🚀 เริ่มการ Extract ข้อความแบบแม่นยำสูง...")

    # 1. ทดสอบอ่านค่าและแสดงผลเบื้องต้น
    results = extract_text_from_pdf(target_pdf, output_format="dict")

    if results:
        # แสดงผลลัพธ์ของหน้าแรกเพื่อตรวจสอบ
        first_page = results.get('page_1', {})
        print("\n--- ตัวอย่างข้อความจากหน้าแรก (500 ตัวอักษร) ---")
        print(first_page.get('raw_text', '')[:500])

        # 2. บันทึกลงไฟล์ text แบบจัดระเบียบ
        save_refined_output(target_pdf)

        # 3. สร้าง HTML เพื่อตรวจเช็ค Reading Order
        generate_visual_html(target_pdf)
    else:
        print("💡 คำแนะนำ: ตรวจสอบว่ามีไฟล์ชื่อ 'input.pdf' อยู่ในโฟลเดอร์เดียวกับสคริปต์หรือไม่")
