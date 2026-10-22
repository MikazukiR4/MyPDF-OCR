import hashlib
import io
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import streamlit as st
from dotenv import load_dotenv


load_dotenv()
load_dotenv(".env.typhoon", override=False)

MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_PAGES = 50
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
DOCX_DEFAULT_FONT = os.getenv("DOCX_DEFAULT_FONT", "TH Sarabun PSK")
FONT_DIR = Path(__file__).resolve().parent / "fonts"
NOTO_THAI_ZIP_URL = "https://github.com/notofonts/thai/releases/download/NotoSansThai-v2.002/NotoSansThai-v2.002.zip"


def font_label(path: Path) -> str:
    name = path.stem.lower()
    if "thsarabun" in name:
        return "TH Sarabun PSK"
    if "angsa" in name or "angsana" in name:
        return "Angsana New"
    if "notosans" in name:
        return "Noto Sans Thai"
    return path.stem


def available_fonts() -> dict[str, Path]:
    FONT_DIR.mkdir(parents=True, exist_ok=True)
    result = {}
    for path in sorted(FONT_DIR.iterdir()):
        if path.suffix.lower() in {".ttf", ".otf"}:
            result.setdefault(font_label(path), path)
    return result


def download_fallback_font() -> tuple[bool, str]:
    import requests

    try:
        response = requests.get(NOTO_THAI_ZIP_URL, timeout=30)
        response.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(response.content)) as archive:
            extracted = 0
            for member in archive.infolist():
                filename = Path(member.filename).name
                if not filename or Path(filename).suffix.lower() not in {".ttf", ".otf", ".txt"}:
                    continue
                target = FONT_DIR / filename
                target.write_bytes(archive.read(member))
                extracted += 1
        if extracted == 0:
            return False, "ไม่พบไฟล์ฟอนต์ในแพ็กเกจ Noto Sans Thai"
        return True, f"ติดตั้ง Noto Sans Thai แล้ว ({extracted} ไฟล์)"
    except Exception as exc:
        return False, f"ดาวน์โหลดฟอนต์ไม่สำเร็จ: {exc}"


def set_docx_font(run: Any, font_name: str | None = None, size: float | None = None) -> None:
    """Set Latin and East Asian font slots so Thai renders consistently in Word/WordCraft."""
    from docx.oxml.ns import qn

    selected = font_name or DOCX_DEFAULT_FONT
    run.font.name = selected
    if size is not None:
        from docx.shared import Pt

        run.font.size = Pt(size)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        from docx.oxml import OxmlElement

        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.set(qn(f"w:{slot}"), selected)


def choose_docx_font(pdf_font: str | None) -> str:
    name = (pdf_font or "").lower()
    if "angsana" in name or "angsa" in name:
        return "Angsana New"
    if "sarabun" in name or "thsarabun" in name:
        return "TH Sarabun PSK"
    return DOCX_DEFAULT_FONT


def set_docx_table_borders(table: Any) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    properties = table._tbl.tblPr
    borders = properties.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        properties.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "808080")


def add_editable_table(word: Any, table_data: list[list[Any]], width_points: float) -> None:
    from docx.shared import Inches, Pt

    row_count = max(len(table_data), 1)
    column_count = max((len(row) for row in table_data), default=0)
    column_count = max(column_count, 1)
    word_table = word.add_table(rows=row_count, cols=column_count)
    word_table.autofit = False
    set_docx_table_borders(word_table)
    column_width = Inches(max(width_points, 72) / 72 / column_count)
    for row_index in range(row_count):
        row_data = table_data[row_index] if row_index < len(table_data) else []
        for column_index in range(column_count):
            cell = word_table.cell(row_index, column_index)
            cell.width = column_width
            cell.text = ""
            value = str(row_data[column_index] or "") if column_index < len(row_data) else ""
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1
            set_docx_font(paragraph.add_run(value), size=12)


def extract_pdf_pages(pdf_bytes: bytes) -> list[tuple[int, str]]:
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(pdf_bytes))
    return [
        (number, (page.extract_text() or "").strip())
        for number, page in enumerate(reader.pages, start=1)
    ]


def render_page_bytes(pdf_bytes: bytes, page_number: int) -> bytes:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = document[page_number - 1]
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        return pixmap.tobytes("jpeg", jpg_quality=90)
    finally:
        document.close()


def export_pdf_as_docx(pdf_bytes: bytes) -> bytes:
    """Create a visually faithful DOCX by placing each PDF page as a full-page image."""
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.shared import Inches

    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    word = Document()
    try:
        for index, page in enumerate(document):
            if index:
                section = word.add_section(WD_SECTION.NEW_PAGE)
            else:
                section = word.sections[0]

            rect = page.rect
            section.page_width = Inches(rect.width / 72)
            section.page_height = Inches(rect.height / 72)
            section.top_margin = Inches(0)
            section.bottom_margin = Inches(0)
            section.left_margin = Inches(0)
            section.right_margin = Inches(0)
            section.header_distance = Inches(0)
            section.footer_distance = Inches(0)

            image_bytes = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False).tobytes(
                "png"
            )
            paragraph = word.add_paragraph()
            paragraph.paragraph_format.space_before = Inches(0)
            paragraph.paragraph_format.space_after = Inches(0)
            paragraph.paragraph_format.line_spacing = 1
            run = paragraph.add_run()
            run.add_picture(io.BytesIO(image_bytes), width=section.page_width)

        output = io.BytesIO()
        word.save(output)
        return output.getvalue()
    finally:
        document.close()


def export_editable_docx(pdf_bytes: bytes, extracted_text: str) -> bytes:
    """Create an editable DOCX using PDF text blocks and OCR text for scanned pages."""
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.shared import Inches, Pt

    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    page_texts: dict[int, str] = {}
    for match in re.finditer(r"\[หน้า\s+(\d+)\s+\|[^\]]+\]\s*\n(.*?)(?=\n\n\[หน้า\s+\d+\s+\||\Z)", extracted_text, re.S):
        page_texts[int(match.group(1))] = match.group(2).strip()

    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    word = Document()
    try:
        for page_index, page in enumerate(pdf):
            section = word.sections[0] if page_index == 0 else word.add_section(WD_SECTION.NEW_PAGE)
            section.page_width = Inches(page.rect.width / 72)
            section.page_height = Inches(page.rect.height / 72)
            section.top_margin = Inches(0.35)
            section.bottom_margin = Inches(0.35)
            section.left_margin = Inches(0.45)
            section.right_margin = Inches(0.45)

            blocks = [block for block in page.get_text("dict").get("blocks", []) if "lines" in block]
            tables = []
            try:
                tables = list(page.find_tables().tables)
            except Exception:
                tables = []

            def overlaps_table(block: Any) -> bool:
                bx0, by0, bx1, by1 = block["bbox"]
                for table in tables:
                    tx0, ty0, tx1, ty1 = table.bbox
                    overlap_width = max(0, min(bx1, tx1) - max(bx0, tx0))
                    overlap_height = max(0, min(by1, ty1) - max(by0, ty0))
                    if overlap_width * overlap_height > 0.35 * max(1, (bx1 - bx0) * (by1 - by0)):
                        return True
                return False

            elements = [(block["bbox"][1], 1, "block", block) for block in blocks if not overlaps_table(block)]
            elements.extend((table.bbox[1], 0, "table", table) for table in tables)
            elements.sort(key=lambda item: (item[0], item[1]))
            if elements:
                previous_bottom = 0.0
                for _, _, element_type, element in elements:
                    if element_type == "table":
                        try:
                            add_editable_table(word, element.extract(), element.bbox[2] - element.bbox[0])
                            previous_bottom = element.bbox[3]
                        except Exception:
                            # Keep the page text if a malformed table cannot be reconstructed.
                            pass
                        continue
                    block = element
                    x0, y0, _, y1 = block["bbox"]
                    paragraph = word.add_paragraph()
                    paragraph.paragraph_format.left_indent = Inches(max(0, x0 - 32) / 72)
                    paragraph.paragraph_format.space_before = Pt(max(0, y0 - previous_bottom))
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1
                    for line_index, line in enumerate(block["lines"]):
                        for span in line.get("spans", []):
                            run = paragraph.add_run(span.get("text", ""))
                            set_docx_font(
                                run,
                                choose_docx_font(span.get("font")),
                                max(6, min(72, span.get("size", 12))),
                            )
                        if line_index < len(block["lines"]) - 1:
                            paragraph.add_run().add_break()
                    previous_bottom = y1
            else:
                text = page_texts.get(page_index + 1, "")
                for line in text.splitlines() or [""]:
                    paragraph = word.add_paragraph(line)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1
                    for run in paragraph.runs:
                        set_docx_font(run, size=12)

        output = io.BytesIO()
        word.save(output)
        return output.getvalue()
    finally:
        pdf.close()


@st.cache_resource
def load_paddle_ocr() -> Any:
    from paddleocr import PaddleOCR

    return PaddleOCR(
        lang="th",
        use_doc_orientation_classify=True,
        use_doc_unwarping=True,
        use_textline_orientation=True,
    )


def extract_paddle_page(image_bytes: bytes) -> str:
    import numpy as np
    from PIL import Image

    image = np.array(Image.open(io.BytesIO(image_bytes)).convert("RGB"))
    ocr = load_paddle_ocr()
    results = ocr.predict(image)
    texts = []
    for result in results:
        data = result.json if hasattr(result, "json") else result
        if callable(data):
            data = data()
        if isinstance(data, str):
            data = json.loads(data)
        if isinstance(data, dict):
            data = data.get("res", data)
            texts.extend(str(text).strip() for text in data.get("rec_texts", []) if str(text).strip())
    return "\n".join(texts)


def extract_tesseract_page(image_bytes: bytes) -> str:
    """Read Thai national-document text with an optional THND Tesseract model."""
    import pytesseract
    from PIL import Image

    tesseract_cmd = os.getenv("TESSERACT_CMD")
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    languages = os.getenv("THND_TESSERACT_LANG", "tha+eng")
    config = os.getenv("THND_TESSERACT_CONFIG", "--psm 6")
    text = pytesseract.image_to_string(
        Image.open(io.BytesIO(image_bytes)).convert("RGB"),
        lang=languages,
        config=config,
    )
    return text.strip()


@st.cache_resource
def load_gemini_client() -> Any:
    from google import genai

    return genai.Client(api_key=os.environ["GEMINI_API_KEY"])


def extract_gemini_page(image_bytes: bytes) -> str:
    if not os.getenv("GEMINI_API_KEY"):
        raise RuntimeError("ยังไม่ได้ตั้ง GEMINI_API_KEY")

    from google.genai import types

    response = load_gemini_client().models.generate_content(
        model=GEMINI_MODEL,
        contents=[
            "ถอดข้อความจากภาพเอกสารนี้เท่านั้น ห้ามสรุปและห้ามเดา "
            "รักษาลำดับบรรทัด ถ้าอ่านไม่ได้ให้ใส่ [อ่านไม่ออก]",
            types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
        ],
        config={"temperature": 0.0, "max_output_tokens": 1500},
    )
    return (response.text or "").strip()


def extract_document(
    pdf_bytes: bytes,
    progress_callback: Any | None = None,
) -> tuple[str, list[str], dict[int, bytes]]:
    if len(pdf_bytes) > MAX_FILE_BYTES:
        raise ValueError("ไฟล์ใหญ่เกิน 20 MB กรุณาลดขนาดหรือแบ่งเอกสารก่อน")

    pages = extract_pdf_pages(pdf_bytes)
    if len(pages) > MAX_PAGES:
        raise ValueError(f"เอกสารมี {len(pages)} หน้า เกินขีดจำกัด {MAX_PAGES} หน้า")

    output = []
    errors = []
    failed_images = {}
    for index, (page_number, embedded_text) in enumerate(pages, start=1):
        if progress_callback:
            progress_callback((index - 1) / len(pages), f"กำลังอ่านหน้า {page_number}/{len(pages)}")

        if embedded_text:
            output.append(f"[หน้า {page_number} | pypdf]\n{embedded_text}")
            continue

        image_bytes = render_page_bytes(pdf_bytes, page_number)
        candidates = []
        stage_errors = []

        try:
            gemini_text = extract_gemini_page(image_bytes)
            if gemini_text:
                candidates.append(("Gemini Vision", gemini_text))
        except Exception as exc:
            stage_errors.append(f"Gemini Vision: {exc}")

        if not candidates:
            try:
                paddle_text = extract_paddle_page(image_bytes)
                if paddle_text:
                    candidates.append(("PaddleOCR", paddle_text))
            except Exception as exc:
                stage_errors.append(f"PaddleOCR: {exc}")

        if not candidates:
            try:
                tesseract_text = extract_tesseract_page(image_bytes)
                if tesseract_text:
                    candidates.append(("THND Tesseract", tesseract_text))
            except Exception as exc:
                stage_errors.append(f"THND Tesseract: {exc}")

        if candidates:
            method, text = candidates[0]
            output.append(f"[หน้า {page_number} | {method}]\n{text}")
        else:
            output.append(f"[หน้า {page_number} | อ่านไม่สำเร็จ]\n")
            failed_images[page_number] = image_bytes
            errors.append(f"หน้า {page_number}: " + " | ".join(stage_errors))

        if progress_callback:
            progress_callback(index / len(pages), f"อ่านหน้า {page_number}/{len(pages)} เสร็จแล้ว")

    return "\n\n".join(output), errors, failed_images


st.set_page_config(page_title="ระบบ Extract ข้อความจาก PDF", page_icon="📄", layout="centered")
st.title("📄 Extract ข้อความจาก PDF")
st.write("pypdf → Gemini Vision → PaddleOCR → THND Tesseract")
st.info("ระบบจะเลื่อนไปใช้เครื่องมือถัดไปเมื่อเครื่องมือก่อนหน้าอ่านไม่สำเร็จ และไม่มีการสรุปเอกสาร")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type=["pdf"])

if uploaded_file is not None and st.button("📥 Extract ข้อความ", type="primary"):
    pdf_bytes = uploaded_file.getvalue()
    font_catalog = available_fonts()
    if not font_catalog and not st.session_state.get("font_auto_attempted"):
        st.session_state["font_auto_attempted"] = True
        download_fallback_font()
        font_catalog = available_fonts()
    with st.sidebar:
        st.subheader("จัดการฟอนต์ Word")
        if font_catalog:
            selected_font = st.selectbox("ฟอนต์สำหรับ Word แบบแก้ไขได้", list(font_catalog))
            DOCX_DEFAULT_FONT = selected_font
            st.caption(f"ฟอนต์ที่พร้อมใช้: {len(font_catalog)} แบบ")
        else:
            st.warning("ยังไม่มีฟอนต์ในระบบ")
        if st.button("ดาวน์โหลด Noto Sans Thai", use_container_width=True):
            ok, message = download_fallback_font()
            (st.success if ok else st.error)(message)
            if ok:
                st.rerun()
        uploaded_font = st.file_uploader(
            "ติดตั้งฟอนต์ด้วยตัวเอง (.ttf/.otf/.woff)",
            type=["ttf", "otf", "woff", "woff2"],
            key="font_upload",
        )
        if uploaded_font is not None:
            filename = Path(uploaded_font.name).name
            target = FONT_DIR / filename
            if target.suffix.lower() in {".ttf", ".otf", ".woff", ".woff2"}:
                if not target.exists():
                    target.write_bytes(uploaded_font.getvalue())
                    st.success(f"ติดตั้ง {filename} แล้ว กดรีเฟรชเพื่อใช้งาน")
            else:
                st.error("รองรับเฉพาะไฟล์ฟอนต์ .ttf, .otf, .woff และ .woff2")
    file_hash = hashlib.sha256(pdf_bytes).hexdigest()[:12]
    progress_bar = st.progress(0, text="เริ่มต้นการอ่าน... (0%)")

    def show_progress(ratio: float, message: str) -> None:
        value = max(0, min(100, round(ratio * 100)))
        progress_bar.progress(value, text=f"{message} ({value}%)")

    try:
        original_docx = b""
        try:
            original_docx = export_pdf_as_docx(pdf_bytes)
        except Exception as export_error:
            st.warning(f"สร้าง Word แบบคงรูปแบบไม่สำเร็จ: {type(export_error).__name__}: {export_error}")

        extracted_text, errors, failed_images = extract_document(pdf_bytes, show_progress)
        progress_bar.progress(100, text="Extract เสร็จสมบูรณ์ (100%)")
        st.success(f"Extract สำเร็จ | รหัสไฟล์ {file_hash}")
        st.text_area("ข้อความที่ Extract ได้", extracted_text, height=500)
        editable_docx = b""
        try:
            editable_docx = export_editable_docx(pdf_bytes, extracted_text)
        except Exception as export_error:
            st.warning(f"สร้าง Word แบบแก้ไขได้ไม่สำเร็จ: {type(export_error).__name__}: {export_error}")
        st.download_button(
            "⬇️ ดาวน์โหลดข้อความเป็น .txt",
            data=extracted_text,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_extracted.txt",
            mime="text/plain",
        )
        st.download_button(
            "📄 ดาวน์โหลด Word (คงหน้าตา PDF เดิม)",
            data=original_docx,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_original-layout.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            help="แต่ละหน้า PDF จะถูกวางเป็นภาพเต็มหน้า จึงรักษารูปแบบต้นฉบับได้ดีที่สุด แต่ข้อความใน Word จะไม่สามารถแก้ไขเป็นตัวอักษรได้",
        )
        st.download_button(
            "✏️ ดาวน์โหลด Word (แก้ไขข้อความได้)",
            data=editable_docx,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_editable.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            help="ข้อความในไฟล์แก้ไขได้ โดยจัดตำแหน่งและขนาดตัวอักษรจาก PDF เท่าที่ทำได้",
        )

        if errors:
            with st.expander("หน้าที่อ่านไม่สำเร็จ"):
                for error in errors:
                    st.error(error)
                for page_number, image_bytes in failed_images.items():
                    st.download_button(
                        f"⬇️ ดาวน์โหลดภาพหน้า {page_number}",
                        data=image_bytes,
                        file_name=f"page_{page_number}_failed.jpg",
                        mime="image/jpeg",
                        key=f"failed-page-{page_number}",
                    )
    except Exception as exc:
        progress_bar.progress(100, text="หยุดการทำงาน (100%)")
        st.error(f"Extract ไม่สำเร็จ: {exc}")
