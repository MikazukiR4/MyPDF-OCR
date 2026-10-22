import io
import os
from typing import Any

import streamlit as st
from dotenv import load_dotenv
from google import genai


load_dotenv()

GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

SYSTEM_INSTRUCTION = """
คุณเป็นผู้ช่วยสรุปเอกสารราชการไทยแบบตรวจสอบหลักฐาน
ใช้เฉพาะข้อมูลที่อยู่ในหลักฐานที่ให้มา ห้ามเดาข้อมูลที่ไม่มี
เมื่อ OCR หลายตัวให้ผลไม่ตรงกัน ให้เปรียบเทียบโดยดูความสอดคล้องกับเอกสารและระบุจุดที่ไม่แน่ใจ
แยกข้อความที่อ่านได้ชัดเจนออกจากข้อความที่คาดเดาไม่ได้
ตอบเป็นภาษาไทยอย่างเป็นทางการ กระชับ และจัดหัวข้อ 5W1H กับ Action Items
"""


def normalize_thai(text: str) -> str:
    """Normalize Thai text without inserting spaces that could alter meaning."""
    try:
        from pythainlp.util import normalize

        return normalize(text)
    except ImportError:
        return text


def extract_pdf_text(pdf_bytes: bytes) -> str:
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            pages.append(f"[หน้า {page_number}]\n{text.strip()}")
    return "\n\n".join(pages)


def render_pdf_pages(pdf_bytes: bytes) -> list[Any]:
    """Render scanned PDF pages for image-based OCR."""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    from PIL import Image

    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    for page in document:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        pages.append(Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples))
    document.close()
    return pages


@st.cache_resource
def load_easyocr_reader() -> Any:
    import easyocr

    return easyocr.Reader(["th", "en"], gpu=False)


def run_easyocr(pages: list[Any]) -> str:
    reader = load_easyocr_reader()
    results = []
    for page_number, image in enumerate(pages, start=1):
        detections = reader.readtext(image, detail=1, paragraph=True)
        lines = []
        for detection in detections:
            if len(detection) >= 3:
                text, confidence = detection[1], float(detection[2])
                lines.append(f"{text} (ความมั่นใจ {confidence:.0%})")
        if lines:
            results.append(f"[หน้า {page_number}]\n" + "\n".join(lines))
    return "\n\n".join(results)


@st.cache_resource
def load_thaitrocr() -> tuple[Any, Any, Any]:
    import torch
    from transformers import AutoProcessor, VisionEncoderDecoderModel

    model_name = "openthaigpt/thai-trocr"
    processor = AutoProcessor.from_pretrained(model_name)
    model = VisionEncoderDecoderModel.from_pretrained(model_name)
    model.eval()
    return processor, model, torch


def run_thaitrocr(pages: list[Any]) -> str:
    processor, model, torch = load_thaitrocr()
    results = []
    for page_number, image in enumerate(pages, start=1):
        pixel_values = processor(images=image, return_tensors="pt").pixel_values
        with torch.no_grad():
            generated_ids = model.generate(pixel_values, max_new_tokens=256)
        text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0].strip()
        if text:
            results.append(f"[หน้า {page_number}]\n{text}")
    return "\n\n".join(results)


def build_evidence(pdf_text: str, easyocr_text: str, thaitrocr_text: str) -> str:
    sections = []
    if pdf_text.strip():
        sections.append("[PDF_TEXT — ข้อความที่ฝังอยู่ใน PDF]\n" + normalize_thai(pdf_text))
    if easyocr_text.strip():
        sections.append("[EASYOCR — OCR ตัวพิมพ์/ข้อความที่ตรวจจับจากภาพ]\n" + normalize_thai(easyocr_text))
    if thaitrocr_text.strip():
        sections.append("[THAITROCR — OCR ลายมือไทย]\n" + normalize_thai(thaitrocr_text))
    return "\n\n".join(sections)


def summarize_with_gemini(evidence: str) -> str:
    prompt = f"""
ต่อไปนี้คือหลักฐานจากเอกสารเดียวกันที่ได้จากหลายตัวอ่านข้อมูล
ให้ตรวจเทียบหลักฐานก่อนสรุป หากข้อมูลขัดแย้งกันให้เลือกข้อความที่มีหลักฐานสนับสนุนมากกว่า
และใส่หัวข้อ 'จุดที่ควรตรวจสอบกับต้นฉบับ' สำหรับข้อความที่ OCR อ่านไม่แน่ชัด

{evidence}
"""
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config={
            "system_instruction": SYSTEM_INSTRUCTION,
            "temperature": 0.1,
        },
    )
    return response.text or "Gemini ไม่ได้ส่งผลลัพธ์กลับมา"


st.set_page_config(
    page_title="ระบบสรุปหนังสือราชการ e-Office",
    page_icon="📄",
    layout="centered",
)

st.title("📄 ระบบสรุปหนังสือราชการ e-Office")
st.write("อ่าน PDF ด้วยหลายวิธี แล้วให้ Gemini เปรียบเทียบหลักฐานและสรุปผล")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF หนังสือราชการ", type=["pdf"])
run_easy = st.checkbox("ใช้ EasyOCR อ่านข้อความจากภาพ", value=True)
run_handwriting = st.checkbox("ใช้ ThaiTrOCR อ่านลายมือไทย (ดาวน์โหลดโมเดลครั้งแรก)", value=True)

if uploaded_file is not None and st.button("🚀 อ่านและสรุปด้วย Gemini", type="primary"):
    pdf_bytes = uploaded_file.getvalue()
    errors = []

    with st.spinner("กำลังเตรียมหลักฐานจากเอกสาร..."):
        try:
            pdf_text = extract_pdf_text(pdf_bytes)
        except Exception as exc:
            pdf_text = ""
            errors.append(f"อ่านข้อความฝังใน PDF ไม่สำเร็จ: {exc}")

        pages = []
        if run_easy or run_handwriting:
            try:
                pages = render_pdf_pages(pdf_bytes)
            except Exception as exc:
                errors.append(f"แปลง PDF เป็นภาพไม่สำเร็จ: {exc}")

        easyocr_text = ""
        if run_easy and pages:
            try:
                easyocr_text = run_easyocr(pages)
            except Exception as exc:
                errors.append(f"EasyOCR ใช้งานไม่ได้ — ติดตั้ง easyocr และ torch ก่อน: {exc}")

        thaitrocr_text = ""
        if run_handwriting and pages:
            try:
                thaitrocr_text = run_thaitrocr(pages)
            except Exception as exc:
                errors.append(f"ThaiTrOCR ใช้งานไม่ได้ — ติดตั้ง transformers, torch และ pillow ก่อน: {exc}")

        evidence = build_evidence(pdf_text, easyocr_text, thaitrocr_text)

    if not evidence.strip():
        st.error("ไม่พบข้อความหรือหลักฐานจากไฟล์นี้")
    else:
        with st.spinner("กำลังให้ Gemini เปรียบเทียบผล OCR และสรุป..."):
            try:
                result = summarize_with_gemini(evidence)
                st.success("วิเคราะห์สำเร็จ")
                st.markdown(result)
                with st.expander("ดูหลักฐานที่ส่งให้ Gemini"):
                    st.text(evidence)
            except Exception as exc:
                st.error(f"เรียก Gemini ไม่สำเร็จ: {exc}")

    if errors:
        with st.expander("ตัวอ่านที่ทำงานไม่สำเร็จ"):
            for error in errors:
                st.warning(error)
