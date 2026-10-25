import io
import html
import hashlib
import csv
import base64
import json
from concurrent.futures import ThreadPoolExecutor
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv

# Import OCR cleaners and exceptions
from core.utils.cleaners import (
    clean_thai_ocr_text,
    clean_simple_text,
    thai_text_quality_score,
    normalize_thai_unicode,
)
from core.exceptions import with_api_retry, OCRError, TransientOCRError, PermanentOCRError

load_dotenv()
load_dotenv(".env.typhoon", override=False)

MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_PAGES = 100
GEMINI_MAX_WORKERS = 3
OCR_MAX_WORKERS = max(1, int(os.getenv("OCR_MAX_WORKERS", "3")))
TESSDATA_DIR = Path(__file__).resolve().parent / "tessdata"

# pytesseract configuration
OEM_LSTM_ONLY = "1"
PSM_VALUES = ["3", "6", "4"]
PDF_VIEWER_HEIGHT = 820
OCR_EDITOR_HEIGHT = PDF_VIEWER_HEIGHT


TYPHOON_BASE_URL = os.getenv("TYPHOON_BASE_URL", "https://api.opentyphoon.ai/v1").rstrip("/")
TYPHOON_MODEL = os.getenv("TYPHOON_MODEL", "typhoon-ocr")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
def render_page(document: object, page_number: int, mode: str = "best") -> bytes:
    """Render one page from an open PyMuPDF document to PNG."""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    page = document[page_number - 1]
    if mode == "fast":
        # 150 DPI is sufficient for the fast OCR path and avoids rendering
        # unnecessarily large images before preprocessing.
        pixmap = page.get_pixmap(dpi=150, alpha=False)
        return pixmap.tobytes("png")
    width = max(1, page.rect.width)
    target_width = 2400
    minimum_scale = 2.5
    scale = min(4.2, max(minimum_scale, target_width / width))
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return pixmap.tobytes("png")


@st.cache_data(show_spinner=False)
def export_pdf_visual_docx(pdf_bytes: bytes) -> bytes:
    """Create a DOCX that preserves the PDF's visual page layout.

    Each PDF page is rendered as a full-page image in a matching DOCX section.
    This keeps the original PDF as the source of truth so scanned text, tables,
    columns, images, stamps, and rotated content remain visually faithful.
    The resulting page images are not editable Word text.
    """
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.shared import Inches

    if not pdf_bytes:
        raise ValueError("ไม่พบข้อมูล PDF สำหรับสร้าง DOCX")

    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        if pdf_document.page_count == 0:
            raise ValueError("ไม่พบหน้าเอกสารใน PDF")

        output = io.BytesIO()
        word_document = Document()
        first_section = word_document.sections[0]

        for page_index in range(pdf_document.page_count):
            page = pdf_document[page_index]
            page_width = max(1.0, float(page.rect.width))
            page_height = max(1.0, float(page.rect.height))
            section = first_section if page_index == 0 else word_document.add_section(WD_SECTION.NEW_PAGE)
            section.page_width = Inches(page_width / 72.0)
            section.page_height = Inches(page_height / 72.0)
            section.top_margin = Inches(0)
            section.bottom_margin = Inches(0)
            section.left_margin = Inches(0)
            section.right_margin = Inches(0)
            section.header_distance = Inches(0)
            section.footer_distance = Inches(0)

            # Render at approximately 200 DPI: readable Thai text and table
            # lines without producing unnecessarily large DOCX files.
            scale = min(4.2, max(1.5, 200.0 / 72.0))
            pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            image_stream = io.BytesIO(pixmap.tobytes("png"))
            paragraph = word_document.add_paragraph()
            paragraph.paragraph_format.space_before = Inches(0)
            paragraph.paragraph_format.space_after = Inches(0)
            paragraph.paragraph_format.line_spacing = 1
            paragraph.add_run().add_picture(
                image_stream,
                width=Inches(page_width / 72.0),
                height=Inches(page_height / 72.0),
            )

        word_document.core_properties.title = "MyOCR PDF visual PDF export"
        word_document.core_properties.subject = "PDF page-preserving DOCX export"
        word_document.core_properties.comments = (
            "Each page is embedded as an image to preserve the original PDF layout."
        )
        word_document.save(output)
        return output.getvalue()
    finally:
        pdf_document.close()


@st.cache_data(show_spinner=False)
def export_pdf_editable_docx(
    pdf_bytes: bytes,
    ocr_text: str = "",
    page_numbers: tuple[int, ...] | None = None,
) -> tuple[bytes, dict]:
    """Layered, editable DOCX (see docx_port.py): page picture behind real Word text boxes.

    ocr_text is the final text shown in the app (Typhoon / Gemini / Tesseract) with its
    "----- หน้า N -----" markers; docx_port aligns it onto each page line geometry, so
    corrected text reaches the DOCX instead of a second, independent Tesseract pass.
    """
    from docx_port import build_editable_docx

    return build_editable_docx(
        pdf_bytes,
        ocr_text,
        page_numbers=list(page_numbers) if page_numbers else None,
        tesseract_cmd=tesseract_path(),
        tessdata_dir=TESSDATA_DIR,
        font_dir=Path(__file__).resolve().parent / "fonts",
    )


@st.cache_data(show_spinner=False)
def export_pdf_flowing_docx(ocr_text: str = "") -> bytes:
    """Clean, flowing Thai DOCX with proper typography, tables, and page breaks."""
    from docx_generator import create_flowing_docx

    return create_flowing_docx(ocr_text)



def preprocess_image(image_bytes: bytes, mode: str) -> bytes:
    """Convert to clean grayscale, upscale only if small, and suppress dotted-line/scan noise."""
    from PIL import Image, ImageFilter, ImageOps

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    target_width = 1500 if mode == "fast" else 2200
    if image.width < target_width:
        scale = min(2.5, target_width / max(1, image.width))
        image = image.resize(
            (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
            Image.Resampling.LANCZOS,
        )
    image = ImageOps.grayscale(image)
    if mode == "best":
        image = ImageOps.autocontrast(image, cutoff=0.5)
        image = image.filter(ImageFilter.MedianFilter(size=3))
        image = image.point(lambda p: 255 if p > 235 else (0 if p < 45 else p))
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def embedded_pdf_text(page: object) -> str:
    """Read an existing PDF text layer before falling back to image OCR."""
    try:
        raw = page.get_text("text") or ""
    except Exception:
        return ""
    raw = re.sub(r"[ \t]+", " ", raw)
    raw = re.sub(r"\n{3,}", "\n\n", raw)
    return clean_thai_ocr_text(raw.strip())


def typhoon_api_key() -> str:
    return os.getenv("TYPHOON_OCR_API_KEY") or os.getenv("TYPHOON_API_KEY") or ""


def available_ocr_engines() -> list[str]:
    """Return only OCR engines that can actually run in this deployment."""
    engines = ["tesseract"] if tesseract_path() else []
    if typhoon_api_key():
        engines.insert(0, "typhoon")
    return engines


def usable_ocr_engine(engine: str) -> str:
    """Prevent stale UI/session state from selecting Typhoon without a key."""
    if engine == "typhoon" and not typhoon_api_key():
        return "tesseract"
    return engine if engine in available_ocr_engines() else "tesseract"


@with_api_retry()
def typhoon_transcribe(image_bytes: bytes) -> str:
    """Use Typhoon OCR for Thai government documents and handwriting."""
    api_key = typhoon_api_key()
    if not api_key:
        raise RuntimeError("ยังไม่ได้ตั้งค่า TYPHOON_OCR_API_KEY")
    from PIL import Image

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    if max(image.size) > 1800:
        scale = 1800 / max(image.size)
        image = image.resize((round(image.width * scale), round(image.height * scale)), Image.Resampling.LANCZOS)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    payload = {
        "model": TYPHOON_MODEL,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": "Extract all text exactly as written. Preserve Thai script, Thai numerals, and line breaks. Return only document text; do not describe images or signatures."},
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode('ascii')}"}},
        ]}],
        "max_tokens": 8192,
        "temperature": 0.1,
        "top_p": 0.6,
        "repetition_penalty": 1.1,
    }
    request = urllib.request.Request(
        f"{TYPHOON_BASE_URL}/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            result = json.loads(response.read().decode("utf-8"))
        return clean_thai_ocr_text(result["choices"][0]["message"]["content"] or "")
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"Typhoon OCR HTTP {exc.code}: {exc.read().decode('utf-8', errors='replace')[:300]}") from exc
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("Typhoon OCR ไม่ส่งข้อความกลับมา") from exc


def gemini_api_key() -> str:
    return os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or ""


@st.cache_data(ttl=300, show_spinner=False)
def list_gemini_models() -> list[dict[str, object]]:
    """Return model metadata available to this API key for generateContent."""
    api_key = gemini_api_key()
    if not api_key:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GEMINI_API_KEY")
    models: list[dict[str, object]] = []
    page_token = ""
    try:
        while True:
            query = "?pageSize=1000"
            if page_token:
                query += f"&pageToken={urllib.parse.quote(page_token)}"
            request = urllib.request.Request(
                f"https://generativelanguage.googleapis.com/v1beta/models{query}",
                headers={"x-goog-api-key": api_key},
                method="GET",
            )
            with urllib.request.urlopen(request, timeout=30) as response:
                result = json.loads(response.read().decode("utf-8"))
            for model in result.get("models", []):
                if "generateContent" in model.get("supportedGenerationMethods", []):
                    name = model.get("name", "")
                    if name.startswith("models/"):
                        name = name[7:]
                    if name:
                        models.append({
                            "name": name,
                            "display_name": model.get("displayName") or name,
                            "description": model.get("description") or "ไม่มีคำอธิบายจาก API",
                            "input_token_limit": model.get("inputTokenLimit"),
                            "output_token_limit": model.get("outputTokenLimit"),
                            "supported_methods": model.get("supportedGenerationMethods", []),
                        })
            page_token = result.get("nextPageToken", "")
            if not page_token:
                break
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Gemini ListModels HTTP {exc.code}: {detail[:300]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"เชื่อมต่อ Gemini ListModels ไม่สำเร็จ: {exc.reason}") from exc
    unique_models = {str(model["name"]): model for model in models}
    return [unique_models[name] for name in sorted(unique_models)]


def gemini_model_kind(model_name: str) -> str:
    name = model_name.lower()
    if "pro" in name:
        return "Pro — เน้นคุณภาพและการให้เหตุผล"
    if "flash-lite" in name or "lite" in name:
        return "Flash-Lite — เร็วและประหยัด"
    if "flash" in name:
        return "Flash — สมดุลความเร็วและคุณภาพ"
    return "รุ่น Gemini อื่น ๆ"


def format_token_limit(value: object) -> str:
    if isinstance(value, (int, float)):
        return f"{int(value):,} tokens"
    return "ไม่ระบุ"


def gemini_model_profile(model_name: str) -> dict[str, object]:
    """Provide an easy-to-read, approximate profile based on the model family."""
    name = model_name.lower()
    if "pro" in name:
        return {
            "quality": 5,
            "speed": 3,
            "vision": 5,
            "economy": 2,
            "best_for": "แก้เอกสารซับซ้อน ต้องการความละเอียดและการให้เหตุผลสูง",
        }
    if "flash-lite" in name or "lite" in name:
        return {
            "quality": 3,
            "speed": 5,
            "vision": 4,
            "economy": 5,
            "best_for": "งานจำนวนมากที่ต้องการความเร็วและประหยัดค่าใช้จ่าย",
        }
    if "flash" in name:
        return {
            "quality": 4,
            "speed": 4,
            "vision": 5,
            "economy": 4,
            "best_for": "แก้ OCR ภาษาไทยทั่วไป สมดุลระหว่างคุณภาพ ความเร็ว และค่าใช้จ่าย",
        }
    return {
        "quality": 3,
        "speed": 3,
        "vision": 3,
        "economy": 3,
        "best_for": "ควรทดสอบกับเอกสารจริงก่อนใช้งานจำนวนมาก",
    }


def star_scale(score: object) -> str:
    if not isinstance(score, int):
        return "ไม่ระบุ"
    score = max(0, min(5, score))
    return "★" * score + "☆" * (5 - score)


def gemini_model_option_label(model_name: str) -> str:
    """Show only the model name in the dropdown options."""
    return model_name


def gemini_model_is_recommended(model_name: str) -> bool:
    """Keep the main picker focused on text/image models suitable for PDF correction."""
    name = model_name.lower()
    excluded_terms = (
        "tts", "image", "audio", "transcribe", "robotics", "computer-use",
        "lyria", "antigravity", "deep-research", "nano-banana",
    )
    return ("flash" in name or "pro" in name) and not any(term in name for term in excluded_terms)


def gemini_rating_html(label: str, score: object, description: str) -> str:
    """Render a rating row with a hover tooltip explaining the 1-5 scale."""
    return (
        '<div class="gemini-rating-row" title="'
        f'{description}'
        '">'
        f'<span class="gemini-rating-label">{label}</span>'
        f'<span class="gemini-rating-stars">{star_scale(score)}</span>'
        '<span class="gemini-rating-help">ⓘ</span>'
        '</div>'
    )


def validate_gemini_guard(source_text: str, corrected: str) -> None:
    """Ensure Gemini did not hallucinate, drop critical numbers, or alter text length wildly."""
    if not corrected or len(corrected) < max(10, round(len(source_text) * 0.40)):
        raise RuntimeError("Gemini ส่งข้อความแก้ไขกลับมาสั้นผิดปกติ จึงคงข้อความเดิมไว้")
    if len(corrected) > max(100, round(len(source_text) * 2.0)):
        raise RuntimeError("Gemini ส่งข้อความแก้ไขกลับมายาวเกินต้นฉบับมากผิดปกติ (อาจ hallucinate)")

    source_nums = set(re.findall(r"[0-9๐-๙]{2,}", source_text))
    if source_nums:
        corrected_nums = set(re.findall(r"[0-9๐-๙]{2,}", corrected))
        th_to_ar = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")
        source_normalized = {n.translate(th_to_ar) for n in source_nums}
        corrected_normalized = {n.translate(th_to_ar) for n in corrected_nums}
        preserved = source_normalized.intersection(corrected_normalized)
        if len(source_normalized) >= 3 and len(preserved) / len(source_normalized) < 0.50:
            raise RuntimeError("Gemini แก้ไขแล้วทำให้ตัวเลข/วันที่สำคัญในเอกสารสูญหายเกิน 50% จึงคงข้อความเดิมไว้")


@with_api_retry()
def gemini_correct_text(image_bytes: bytes, source_text: str, model: str) -> str:
    """Correct Thai PDF text using the page image as the source of truth."""
    api_key = gemini_api_key()
    if not api_key:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GEMINI_API_KEY")
    prompt = (
        "คุณเป็นระบบตรวจแก้ข้อความ OCR ภาษาไทยจากเอกสารราชการ\n"
        "อ่านข้อความจากภาพต้นฉบับเป็นหลัก และใช้ข้อความ OCR ด้านล่างเป็นเพียงตัวช่วย\n"
        "แก้เฉพาะสระ วรรณยุกต์ การแยกคำ ช่องว่าง และอักขระที่อ่านผิดจาก PDF font mapping\n"
        "ห้ามสรุป ห้ามตัดข้อความ ห้ามเพิ่มข้อความ และห้ามเปลี่ยนเลขไทย วันที่ เลขที่หนังสือ หรือสาระสำคัญ\n"
        "รักษาการขึ้นบรรทัดใหม่ให้ใกล้เคียงต้นฉบับ หากไม่แน่ใจให้คงข้อความเดิมไว้\n"
        "ส่งคืนเฉพาะข้อความที่แก้แล้ว ห้ามใส่ Markdown ห้ามใส่คำอธิบาย\n\n"
        "ข้อความ OCR เดิม:\n"
        f"{source_text}"
    )
    payload = {
        "contents": [{"parts": [
            {"text": prompt},
            {"inline_data": {"mime_type": "image/png", "data": base64.b64encode(image_bytes).decode("ascii")}},
        ]}],
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 8192},
    }
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            result = json.loads(response.read().decode("utf-8"))
        corrected = "\n".join(
            part["text"]
            for part in result["candidates"][0]["content"]["parts"]
            if part.get("text")
        ).strip()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        if exc.code == 404:
            raise RuntimeError(
                f"Gemini HTTP 404: ไม่พบหรือไม่เปิดใช้โมเดล '{model}' "
                "กรุณาเลือกโมเดลจากรายการที่ตรวจพบใน Sidebar"
            ) from exc
        raise RuntimeError(f"Gemini HTTP {exc.code}: {detail[:300]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"เชื่อมต่อ Gemini ไม่สำเร็จ: {exc.reason}") from exc
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("Gemini ไม่ส่งข้อความกลับมา") from exc

    validate_gemini_guard(source_text, corrected)
    return clean_thai_ocr_text(corrected)


@with_api_retry()
def gemini_correct_document(pdf_bytes: bytes, source_text: str, model: str) -> str:
    """Correct a PDF using Gemini's native document understanding.

    Sending the original PDF lets Gemini use its embedded text and page
    layout instead of requiring one rendered PNG request per page. The local
    OCR text is retained as a reference for characters that are difficult to
    read from the document alone.
    """
    api_key = gemini_api_key()
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY is not configured")

    prompt = (
        "You are correcting Thai OCR from an official document. Read the PDF "
        "itself as the primary source, including its native text, layout, and "
        "visual content. Use the supplied OCR only as a secondary reference. "
        "Correct spelling, Thai vowels, tone marks, spacing, and PDF font "
        "mapping errors without summarizing or adding text. Preserve every "
        "page marker and return the corrected text in the same page order. "
        "Return only the corrected document text.\n\n"
        "LOCAL OCR REFERENCE:\n"
        f"{source_text}"
    )
    payload = {
        "contents": [{"parts": [
            {"text": prompt},
            {"inline_data": {
                "mime_type": "application/pdf",
                "data": base64.b64encode(pdf_bytes).decode("ascii"),
            }},
        ]}],
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 32768},
    }
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            result = json.loads(response.read().decode("utf-8"))
        corrected = "\n".join(
            part["text"]
            for part in result["candidates"][0]["content"]["parts"]
            if part.get("text")
        ).strip()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Gemini document HTTP {exc.code}: {detail[:300]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Gemini document connection failed: {exc.reason}") from exc
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("Gemini returned no document text") from exc

    validate_gemini_guard(source_text, corrected)
    validate_document_markers(source_text, corrected)
    return clean_thai_ocr_text(corrected)


PAGE_MARKER_RE = re.compile(r"^----- หน้า ([0-9]+)(?: \| .*?)? -----$", re.MULTILINE)


def document_page_markers(text: str) -> list[int]:
    """Extract ordered page markers from a document OCR result."""
    return [int(match.group(1)) for match in PAGE_MARKER_RE.finditer(text or "")]


def validate_document_markers(source_text: str, corrected_text: str) -> None:
    """Reject Gemini output that loses, duplicates, or reorders page markers."""
    expected = document_page_markers(source_text)
    actual = document_page_markers(corrected_text)
    if not expected:
        raise RuntimeError("Local OCR output has no page markers to validate")
    if actual != expected:
        missing = [page for page in expected if page not in actual]
        unexpected = [page for page in actual if page not in expected]
        details = []
        if missing:
            details.append(f"หน้าที่หาย: {', '.join(map(str, missing[:10]))}")
        if unexpected:
            details.append(f"หน้าที่เกิน/ซ้ำ: {', '.join(map(str, unexpected[:10]))}")
        if not details:
            details.append("ลำดับ marker ไม่ตรงกับเอกสารต้นฉบับ")
        raise RuntimeError("Gemini ส่ง marker หน้าไม่ครบหรือเรียงลำดับผิด: " + "; ".join(details))




def tesseract_path() -> str | None:
    configured = os.getenv("TESSERACT_CMD")
    if configured:
        return configured
    found = shutil.which("tesseract")
    if found:
        return found
    for candidate in (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return None


def _extract_page_text_once(
    image_bytes: bytes,
    language: str,
    mode: str,
    psm: str,
    oem: str = OEM_LSTM_ONLY,
) -> str:
    command = tesseract_path()
    if not command:
        raise RuntimeError(
            "ไม่พบ Tesseract ในระบบ กรุณาติดตั้ง tesseract-ocr และ tesseract-ocr-tha"
        )
    prepared = preprocess_image(image_bytes, mode)
    # Call the Tesseract executable directly. This avoids pytesseract's
    # temporary-file decoding path, which can raise UnicodeDecodeError on
    # Windows when Thai output contains mixed encodings.
    with tempfile.TemporaryDirectory(prefix="th_ocr_") as temp_dir:
        image_path = Path(temp_dir) / "page.png"
        image_path.write_bytes(prepared)
        args = [
            command,
            str(image_path),
            "stdout",
            "-l",
            language,
            "--psm",
            str(psm),
            "--oem",
            str(oem),
            "--dpi",
            "300",
            "-c",
            "user_defined_dpi=300",
        ]
        if "tha" not in language:
            args.extend(["-c", "preserve_interword_spaces=1"])
        if (TESSDATA_DIR / "tha.traineddata").exists() and (TESSDATA_DIR / "eng.traineddata").exists():
            args.extend(["--tessdata-dir", str(TESSDATA_DIR)])
        result = subprocess.run(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
            check=False,
        )
    if result.returncode != 0:
        error = result.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(f"Tesseract ทำงานไม่สำเร็จ (รหัส {result.returncode}): {error}")

    raw = result.stdout
    # Decode defensively so one page cannot fail only because of a Windows
    # code-page mismatch.
    candidates = []
    for encoding in ("utf-8", "cp874", "tis-620"):
        decoded = raw.decode(encoding, errors="replace")
        replacement_count = decoded.count("\ufffd")
        control_count = sum(1 for char in decoded if ord(char) < 32 and char not in "\n\r\t")
        candidates.append((replacement_count, control_count, decoded))
    return clean_thai_ocr_text(min(candidates, key=lambda item: (item[0], item[1]))[2])


class TesseractWorker:
    """Reusable OCR worker with a Tesseract.js-like lifecycle."""

    def __init__(self, language: str, mode: str, psm: str, oem: str = OEM_LSTM_ONLY) -> None:
        self.command = tesseract_path()
        if not self.command:
            raise RuntimeError("ไม่พบ Tesseract ในระบบ กรุณาติดตั้ง tesseract-ocr และ tesseract-ocr-tha")
        self.language = language
        self.mode = mode
        self.psm = psm
        self.oem = oem
        self.closed = False

    def recognize(self, image_bytes: bytes) -> str:
        if self.closed:
            raise RuntimeError("OCR worker ถูกปิดไปแล้ว")
        return _extract_page_text_once(image_bytes, self.language, self.mode, self.psm, self.oem)

    def terminate(self) -> None:
        self.closed = True


def parse_page_range(range_str: str, max_pages: int) -> list[int]:
    """Parse string like '1-3, 5, 8-10' into sorted list of unique 1-indexed page numbers."""
    if not range_str or not range_str.strip():
        return list(range(1, max_pages + 1))
    pages = set()
    parts = [p.strip() for p in range_str.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            sub = part.split("-")
            if len(sub) == 2 and sub[0].strip().isdigit() and sub[1].strip().isdigit():
                start = max(1, int(sub[0].strip()))
                end = min(max_pages, int(sub[1].strip()))
                if start <= end:
                    pages.update(range(start, end + 1))
        elif part.isdigit():
            val = int(part)
            if 1 <= val <= max_pages:
                pages.add(val)
    res = sorted(pages)
    return res if res else list(range(1, max_pages + 1))


def extract_pdf(
    pdf_bytes: bytes,
    language: str,
    mode: str,
    psm: str,
    oem: str,
    engine: str,
    use_gemini_correction: bool,
    gemini_model: str,
    progress: object,
    page_callback: object | None = None,
    pages_to_process: list[int] | None = None,
    page_cache: dict[int, tuple[str, str]] | None = None,
) -> tuple[str, list[tuple[int | None, str]], dict[int, bytes]]:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    engine = usable_ocr_engine(engine)
    if len(pdf_bytes) > MAX_FILE_BYTES:
        raise ValueError("ไฟล์ใหญ่เกิน 50 MB")

    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        page_count = document.page_count
        if page_count == 0:
            raise ValueError("ไม่พบหน้าเอกสารใน PDF")
        if page_count > MAX_PAGES:
            raise ValueError(f"เอกสารมี {page_count} หน้า เกินขีดจำกัด {MAX_PAGES} หน้า")

        if pages_to_process:
            target_pages = [p for p in pages_to_process if 1 <= p <= page_count]
        else:
            target_pages = list(range(1, page_count + 1))

        if not target_pages:
            raise ValueError("ไม่พบหน้าที่ระบุในเอกสาร PDF")

        total_target = len(target_pages)
        output: list[str] = []
        errors: list[tuple[int | None, str]] = []
        failed_images: dict[int, bytes] = {}
        worker = None if engine == "typhoon" else TesseractWorker(language, mode, psm, oem)
        gemini_document_mode = use_gemini_correction  # Removed redundant file size check; upstream already raises if too large (line 890-891)

        def publish_page() -> None:
            if page_callback is not None:
                page_callback("\n\n".join(output))

        try:
            batch_size = max(1, OCR_MAX_WORKERS)
            for batch_start_idx in range(0, total_target, batch_size):
                current_batch = target_pages[batch_start_idx : batch_start_idx + batch_size]
                page_images: dict[int, bytes] = {}
                for idx_in_batch, page_number in enumerate(current_batch):
                    render_ratio = (batch_start_idx + idx_in_batch) / total_target
                    progress.progress(
                        0.1 + render_ratio * 0.45,
                        text=f"กำลังเตรียมหน้า {page_number} ({batch_start_idx + idx_in_batch + 1}/{total_target})",
                    )
                    page_images[page_number] = render_page(document, page_number, mode)

                def recognize_page(page_number: int) -> tuple[int, str, str, bytes]:
                    image_bytes = page_images[page_number]
                    # Step 0: Check per-page cache
                    if page_cache and page_number in page_cache:
                        cached_text, cached_method = page_cache[page_number]
                        if cached_text:
                            return page_number, cached_text, f"{cached_method} (cached)", image_bytes

                    # Step 1: Check embedded text layer
                    page_obj = document[page_number - 1]
                    embedded = embedded_pdf_text(page_obj)
                    if embedded and thai_text_quality_score(embedded) >= 0.70:
                        if page_cache is not None:
                            page_cache[page_number] = (embedded, "embedded text")
                        return page_number, embedded, "embedded text", image_bytes

                    # Step 2: Run primary OCR engine
                    method = engine
                    if engine == "typhoon":
                        text = typhoon_transcribe(image_bytes)
                    else:
                        text = worker.recognize(image_bytes)
                        # Step 3: Fallback chain to Typhoon if Tesseract produced poor quality and Typhoon API key is available
                        if thai_text_quality_score(text) < 0.55 and typhoon_api_key():
                            try:
                                typhoon_text = typhoon_transcribe(image_bytes)
                                if thai_text_quality_score(typhoon_text) > thai_text_quality_score(text):
                                    text = typhoon_text
                                    method = "typhoon (fallback from low-quality tesseract)"
                            except Exception:
                                pass

                    if not text:
                        if embedded:
                            if page_cache is not None:
                                page_cache[page_number] = (embedded, "embedded text (fallback)")
                            return page_number, embedded, "embedded text (fallback)", image_bytes
                        raise RuntimeError("OCR engine ส่งข้อความว่างกลับมา")

                    if page_cache is not None:
                        page_cache[page_number] = (text, method)
                    return page_number, text, method, image_bytes

                page_data: dict[int, tuple[str, str, bytes | None]] = {}
                with ThreadPoolExecutor(max_workers=min(OCR_MAX_WORKERS, len(page_images))) as executor:
                    futures = {
                        page_number: executor.submit(recognize_page, page_number)
                        for page_number in page_images
                    }
                    for page_number in current_batch:
                        try:
                            _, text, method, image_bytes = futures[page_number].result()
                            page_data[page_number] = (text, method, image_bytes)
                        except Exception as exc:
                            errors.append((page_number, f"{type(exc).__name__}: {exc}"))
                            failed_images[page_number] = page_images[page_number]

                gemini_jobs = {
                    page_number: (image_bytes, text)
                    for page_number, (text, method, image_bytes) in page_data.items()
                    if use_gemini_correction
                    and not gemini_document_mode
                    and not method.startswith("embedded")
                    and image_bytes is not None
                    and text
                    and not text.startswith("[")
                }
                gemini_futures = {}
                if gemini_jobs:
                    with ThreadPoolExecutor(max_workers=min(GEMINI_MAX_WORKERS, len(gemini_jobs))) as executor:
                        gemini_futures = {
                            page_number: executor.submit(
                                gemini_correct_text, image_bytes, source_text, gemini_model
                            )
                            for page_number, (image_bytes, source_text) in gemini_jobs.items()
                        }
                        for page_number in current_batch:
                            if page_number in gemini_futures:
                                try:
                                    corrected_text = gemini_futures[page_number].result()
                                    _, _, image_bytes = page_data[page_number]
                                    page_data[page_number] = (corrected_text, "Gemini corrected", image_bytes)
                                    if page_cache is not None:
                                        page_cache[page_number] = (corrected_text, "Gemini corrected")
                                except Exception as correction_error:
                                    errors.append((page_number, f"Gemini ตรวจแก้ไม่สำเร็จ: {correction_error}"))

                for idx_in_batch, page_number in enumerate(current_batch):
                    if page_number not in page_data:
                        output.append(f"----- หน้า {page_number} -----\n[อ่านหน้านี้ไม่สำเร็จ]")
                    else:
                        text, method, _ = page_data[page_number]
                        output.append(f"----- หน้า {page_number} | {method} -----\n{text}")
                    done_count = batch_start_idx + idx_in_batch + 1
                    progress.progress(
                        0.1 + (done_count / total_target) * 0.55,
                        text=f"อ่านหน้า {page_number} เสร็จแล้ว ({done_count}/{total_target})",
                    )
                    publish_page()
        finally:
            if worker is not None:
                worker.terminate()

        local_output = "\n\n".join(output)
        if gemini_document_mode and local_output:
            progress.progress(0.75, text="กำลังส่ง PDF ให้ Gemini ตรวจแก้...")
            try:
                if len(target_pages) < page_count:
                    sub_doc = fitz.open()
                    for p in target_pages:
                        sub_doc.insert_pdf(document, from_page=p - 1, to_page=p - 1)
                    target_pdf_bytes = sub_doc.tobytes()
                    sub_doc.close()
                else:
                    target_pdf_bytes = pdf_bytes

                corrected_document = gemini_correct_document(target_pdf_bytes, local_output, gemini_model)
                progress.progress(0.98, text="ตรวจสอบลำดับและ marker หน้าแล้ว")
                return corrected_document, errors, failed_images
            except Exception as correction_error:
                errors.append((None, f"Gemini ตรวจแก้ทั้งเอกสารไม่สำเร็จ: {correction_error}"))
        return local_output, errors, failed_images
    finally:
        document.close()


def render_pdf_viewer(pdf_bytes: bytes, filename: str, height: int = PDF_VIEWER_HEIGHT) -> None:
    """Show a PDF viewer with built-in fit, zoom-in, and zoom-out controls."""
    if not pdf_bytes:
        st.error("ไม่พบข้อมูล PDF สำหรับแสดงตัวอย่าง")
        return
    display_name = filename or "document.pdf"
    encoded_pdf = base64.b64encode(pdf_bytes).decode("ascii")
    safe_name = json.dumps(display_name, ensure_ascii=False)
    viewer_html = f"""
    <!doctype html>
    <html lang="th">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          html, body {{ width:100%; height:100%; min-height:100%; margin:0; }}
          body {{ overflow:hidden; background:#172554; }}
          #viewer-shell {{ position:relative; width:100%; height:100%; min-height:0; overflow-x:hidden; overflow-y:auto; overscroll-behavior:contain; }}
          #pdf-viewer {{ width:100%; height:100%; min-height:0; }}
          #zoom-controls {{ position:absolute; top:72px; right:12px; z-index:20; display:flex; align-items:center; gap:4px; padding:4px; border-radius:10px; background:rgba(15,23,42,.88); box-shadow:0 3px 12px rgba(0,0,0,.28); }}
          #zoom-controls button {{ width:34px; height:32px; border:0; border-radius:7px; background:#334155; color:#fff; font-size:20px; line-height:1; cursor:pointer; }}
          #zoom-controls button:hover {{ background:#475569; }}
          #zoom-level {{ min-width:52px; color:#fff; font:600 12px system-ui,sans-serif; text-align:center; }}
          #viewer-error {{ padding:1rem; color:#fecaca; font:14px sans-serif; }}
        </style>
      </head>
      <body>
        <div id="viewer-shell">
          <div id="pdf-viewer" aria-label="ตัวอย่าง PDF: {safe_name}"></div>
          <div id="zoom-controls" aria-label="ควบคุมการซูม PDF">
            <button id="zoom-out" type="button" aria-label="ซูมออก" title="ซูมออก">−</button>
            <span id="zoom-level" aria-live="polite">100%</span>
            <button id="zoom-in" type="button" aria-label="ซูมเข้า" title="ซูมเข้า">+</button>
            <button id="zoom-fit" type="button" aria-label="พอดีหน้าต่าง" title="พอดีหน้าต่าง">Fit</button>
          </div>
        </div>
        <div id="viewer-error" role="alert" hidden></div>
        <script type="module">
          import EmbedPDF from "https://cdn.jsdelivr.net/npm/@embedpdf/snippet@2/dist/embedpdf.js";
          const payload = {{ data: "{encoded_pdf}" }};
          const target = document.getElementById("pdf-viewer");
          const errorTarget = document.getElementById("viewer-error");
          const zoomLevel = document.getElementById("zoom-level");
          let zoom = 1;
          function applyZoom(nextZoom) {{
            zoom = Math.max(0.75, Math.min(2.5, nextZoom));
            // CSS zoom changes the layout width of EmbedPDF and makes the page
            // shrink when zooming in. Transform keeps the viewer's layout box
            // stable while enlarging the rendered document predictably.
            target.style.zoom = "1";
            target.style.transform = `scale(${{zoom}})`;
            target.style.transformOrigin = "top left";
            target.style.width = "100%";
            target.style.height = "100%";
            zoomLevel.textContent = `${{Math.round(zoom * 100)}}%`;
          }}
          document.getElementById("zoom-out").addEventListener("click", () => applyZoom(zoom - 0.1));
          document.getElementById("zoom-in").addEventListener("click", () => applyZoom(zoom + 0.1));
          document.getElementById("zoom-fit").addEventListener("click", () => applyZoom(1));
          function showError(message) {{
            errorTarget.textContent = message;
            errorTarget.hidden = false;
          }}
          try {{
            const binary = atob(payload.data);
            const bytes = Uint8Array.from(binary, character => character.charCodeAt(0));
            const blob = new Blob([bytes], {{type:"application/pdf"}});
            const pdfUrl = URL.createObjectURL(blob);
            EmbedPDF.init({{
              type: "container",
              target,
              src: pdfUrl,
              theme: {{ preference: "system" }}
            }});
          }} catch (error) {{
            showError(`เปิดตัวอย่าง PDF ไม่สำเร็จ: ${{error.message}}`);
          }}
        </script>
      </body>
    </html>
    """
    components.html(viewer_html, height=height, scrolling=False)
    st.download_button(
        "ดาวน์โหลด PDF ต้นฉบับ",
        data=pdf_bytes,
        file_name=display_name,
        mime="application/pdf",
        key=f"download-pdf-{hashlib.sha256(pdf_bytes).hexdigest()[:16]}",
        use_container_width=True,
    )

def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root { --ink:#0f172a; --muted:#475569; --line:#cbd5e1; --accent:#dc2626; --accent-dark:#991b1b; --navy:#172554; }
        html, body, button, input, textarea, select { font-family:"Segoe UI", "Noto Sans Thai", Tahoma, sans-serif !important; }
        .stApp, [data-testid="stAppViewContainer"] { background: #f8fafc !important; color: var(--ink) !important; }
        header[data-testid="stHeader"], [data-testid="stHeader"] { background: #0b1220 !important; height: 2.5rem !important; }
        [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"] { background: transparent !important; }
        .block-container { width:100%; max-width:none; padding:3.5rem clamp(1rem, 3vw, 3rem) 2rem; }
        :root { --panel-height: min(78vh, 820px); }
        div[data-testid="stIFrame"], div[data-testid="stCustomComponentV1"] { height:var(--panel-height) !important; min-height:0 !important; }
        div[data-testid="stIFrame"] iframe, div[data-testid="stCustomComponentV1"] iframe { height:100% !important; min-height:0 !important; }
        .mypdf-header { display:flex; align-items:center; gap:12px; margin-bottom:2rem; }
        .mypdf-icon { background:var(--accent); color:#fff; border-radius:9px; padding:9px 12px; font-size:24px; box-shadow:0 4px 10px rgba(220,38,38,.2); }
        .mypdf-name { font-size:30px; font-weight:800; color:var(--ink); letter-spacing:-.3px; }
        .mypdf-name span { color:var(--accent); }
        .ocr-card { background:#fff; border:1px solid #e5e7eb; border-radius:16px; padding:24px; box-shadow:0 5px 20px rgba(15,23,42,.06); }
        .subtitle { color:var(--muted) !important; margin-top:-12px; margin-bottom:24px; font-size:15px; }
        .workflow-steps { display:flex; flex-wrap:wrap; align-items:center; gap:8px; margin:-8px 0 24px; color:#64748b; font-size:14px; }
        .workflow-step { padding:6px 10px; border:1px solid #cbd5e1; border-radius:999px; background:#fff; }
        .workflow-step.active { color:#991b1b; border-color:#fca5a5; background:#fef2f2; font-weight:700; }
        div[data-testid="stFileUploader"] { background:#ffffff !important; border:2px dashed #94a3b8; border-radius:14px; padding:10px; }
        div[data-testid="stFileUploader"] section, div[data-testid="stFileUploader"] section > div { background:#ffffff !important; color:#172033 !important; }
        div[data-testid="stFileUploader"] small, div[data-testid="stFileUploader"] span, div[data-testid="stFileUploader"] label { color:#334155 !important; }
        div[data-testid="stProgressBar"] > div { background:#e2e8f0 !important; }
        div[data-testid="stButton"] > button { background:#dc2626 !important; color:#ffffff !important; border:1px solid #991b1b !important; border-radius:8px !important; font-weight:700 !important; }
        div[data-testid="stButton"] > button:hover { background:#b91c1c !important; color:#ffffff !important; border-color:#991b1b !important; }
        div[data-testid="stButton"] > button[kind="primary"], button[kind="primary"] { background:var(--accent) !important; color:#ffffff !important; border-color:var(--accent-dark) !important; box-shadow:0 4px 10px rgba(220,38,38,.18); }
        div[data-testid="stButton"] > button[kind="primary"]:hover, button[kind="primary"]:hover { background:var(--accent-dark) !important; color:#ffffff !important; }
        div[data-testid="stFileUploader"] button { background:#334155 !important; color:#ffffff !important; border:1px solid #1e293b !important; border-radius:7px !important; }
        div[data-testid="stFileUploader"] button:hover { background:#475569 !important; color:#ffffff !important; }
        div[data-testid="stDownloadButton"] > button { background:#0f766e !important; color:#ffffff !important; border:1px solid #115e59 !important; font-weight:700 !important; min-height:44px !important; }
        div[data-testid="stDownloadButton"] > button:hover { background:#115e59 !important; color:#ffffff !important; }
        div[data-testid="stTextArea"] label, div[data-testid="stTextArea"] label p { color:#172033 !important; font-weight:700 !important; }
        div[data-testid="stTextArea"] textarea { background:#020617 !important; color:#ffffff !important; border:2px solid #64748b !important; border-radius:10px !important; font-size:16px !important; line-height:1.7 !important; opacity:1 !important; -webkit-text-fill-color:#ffffff !important; caret-color:#ffffff !important; }
        div[data-testid="stTextArea"] textarea { height:var(--panel-height) !important; min-height:var(--panel-height) !important; max-height:var(--panel-height) !important; overflow-x:hidden !important; overflow-y:auto !important; white-space:pre-wrap !important; overflow-wrap:anywhere !important; word-break:break-word !important; resize:none !important; box-sizing:border-box !important; }
        div[data-testid="stCodeBlock"] pre { max-height:var(--panel-height) !important; overflow-x:hidden !important; overflow-y:auto !important; white-space:pre-wrap !important; overflow-wrap:anywhere !important; word-break:break-word !important; }
        div[data-testid="stTextArea"] textarea::placeholder { color:#64748b !important; opacity:1 !important; }
        div[data-testid="stAlert"] { color:#172033 !important; border-radius:10px !important; }
        div[data-testid="stAlert"] p { color:inherit !important; }
        .stCaption, .stMarkdown, .stMarkdown p, label, label p { color:#334155 !important; }
        aside[data-testid="stSidebar"] { background:var(--navy) !important; border-right:1px solid #1e3a8a !important; }
        aside[data-testid="stSidebar"] * { color:#f8fafc !important; }
        aside[data-testid="stSidebar"] [data-testid="stWidgetLabel"] *, aside[data-testid="stSidebar"] [data-testid="stCaptionContainer"] *, aside[data-testid="stSidebar"] .stMarkdown * { color:#f8fafc !important; }
        aside[data-testid="stSidebar"] label, aside[data-testid="stSidebar"] p, aside[data-testid="stSidebar"] h1, aside[data-testid="stSidebar"] h2, aside[data-testid="stSidebar"] h3 { color:#f8fafc !important; }
        aside[data-testid="stSidebar"] [data-baseweb="select"] { background:#ffffff !important; }
        aside[data-testid="stSidebar"] [data-baseweb="select"] * { color:#172033 !important; }
        aside[data-testid="stSidebar"] [data-baseweb="select"] input, aside[data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] { color:#172033 !important; -webkit-text-fill-color:#172033 !important; }
        .small-note { color:#64748b; font-size:13px; }
        .pdf-empty { height:var(--panel-height); box-sizing:border-box; display:flex; align-items:center; justify-content:center; padding:2rem; border:1px dashed #64748b; border-radius:12px; background:#020617; color:#94a3b8; text-align:center; line-height:1.7; }
        .ocr-live-preview { height:var(--panel-height); box-sizing:border-box; padding:16px; border:2px solid #64748b; border-radius:10px; background:#020617; color:#ffffff; font:16px/1.7 "Segoe UI", "Noto Sans Thai", Tahoma, sans-serif; white-space:pre-wrap; overflow-wrap:anywhere; word-break:break-word; overflow-x:hidden; overflow-y:auto; }
        /* Dark mode: keep every surface and text role readable together. */
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] { background:#0b1220 !important; color:#f8fafc !important; }
        .block-container { color:#f8fafc !important; }
        .ocr-card { background:#111827 !important; border-color:#334155 !important; box-shadow:0 8px 24px rgba(0,0,0,.28) !important; }
        .mypdf-name { color:#f8fafc !important; }
        .subtitle, .small-note, .stCaption, .stMarkdown, .stMarkdown p, label, label p { color:#cbd5e1 !important; }
        div[data-testid="stFileUploader"], div[data-testid="stFileUploader"] section, div[data-testid="stFileUploader"] section > div { background:#111827 !important; border-color:#64748b !important; color:#f8fafc !important; }
        div[data-testid="stFileUploader"] small, div[data-testid="stFileUploader"] span, div[data-testid="stFileUploader"] label { color:#cbd5e1 !important; }
        div[data-testid="stTextArea"] textarea { background:#020617 !important; color:#ffffff !important; border:2px solid #64748b !important; -webkit-text-fill-color:#ffffff !important; caret-color:#ffffff !important; font-size:16px !important; line-height:1.7 !important; height:var(--panel-height) !important; min-height:var(--panel-height) !important; max-height:var(--panel-height) !important; overflow-x:hidden !important; overflow-y:auto !important; white-space:pre-wrap !important; overflow-wrap:anywhere !important; word-break:break-word !important; resize:none !important; }
        div[data-testid="stTextArea"] label, div[data-testid="stTextArea"] label p { color:#f8fafc !important; }
        div[data-testid="stAlert"] { background:#172033 !important; color:#f8fafc !important; border-color:#475569 !important; }
        div[data-testid="stAlert"] p { color:#f8fafc !important; }
        div[data-testid="stProgressBar"] > div { background:#1e293b !important; }
        [data-baseweb="select"], [data-baseweb="select"] > div { background:#0f172a !important; border-color:#475569 !important; color:#f8fafc !important; }
        [data-baseweb="select"] * { color:#f8fafc !important; -webkit-text-fill-color:#f8fafc !important; }
        [role="listbox"] { background:#111827 !important; color:#f8fafc !important; min-width:520px !important; width:max-content !important; max-width:90vw !important; }
        [role="option"] { background:#111827 !important; color:#f8fafc !important; white-space:normal !important; overflow:visible !important; text-overflow:clip !important; min-height:42px !important; line-height:1.35 !important; padding:8px 12px !important; }
        [role="option"]:hover { background:#1e3a5f !important; }
        aside[data-testid="stSidebar"] [data-baseweb="select"] { min-width:100% !important; }
        aside[data-testid="stSidebar"] [data-baseweb="select"] > div { min-width:100% !important; }
        div[data-testid="stDownloadButton"] > button { background:#0f766e !important; color:#ffffff !important; border:1px solid #115e59 !important; min-height:44px !important; }
        aside[data-testid="stSidebar"] { background:#111827 !important; }
        .gemini-rating-row { display:flex; align-items:center; gap:8px; padding:4px 6px; border-radius:5px; cursor:help; }
        .gemini-rating-row:hover { background:#1e293b; }
        .gemini-rating-label { min-width:145px; color:#e5e7eb; font-size:0.88rem; }
        .gemini-rating-stars { color:#fbbf24; letter-spacing:1px; font-size:1rem; }
        .gemini-rating-help { color:#94a3b8; font-size:0.78rem; }
        @media (max-width: 700px) {
            :root { --panel-height: min(62vh, 620px); }
            .block-container { padding-left:1rem; padding-right:1rem; padding-top:2.5rem; }
            .mypdf-header { margin-bottom:1.25rem; }
            .mypdf-name { font-size:25px; }
            .workflow-steps { gap:5px; font-size:12px; }
            .workflow-steps > span:not(.workflow-step) { display:none; }
            .workflow-step { padding:5px 7px; }
            .ocr-card { padding:16px; }
            div[data-testid="stTextArea"] textarea { font-size:15px !important; }
        }
        @media (max-width: 900px) {
            [data-testid="stHorizontalBlock"] { flex-direction:column !important; gap:1rem !important; }
            [data-testid="stColumn"] { width:100% !important; min-width:100% !important; flex:1 1 100% !important; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title="MyOCR PDF — Thai", page_icon="📄", layout="wide")
inject_css()

st.markdown(
    '<div class="mypdf-header"><div class="mypdf-icon">PDF</div>'
    '<div class="mypdf-name"><span>MyOCR</span> PDF</div></div>',
    unsafe_allow_html=True,
)
st.markdown('<div class="subtitle">อ่านข้อความจาก PDF สแกน — ภาษาไทยและภาษาอังกฤษ</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="workflow-steps" aria-label="ขั้นตอนการใช้งาน">'
    '<span class="workflow-step active">① เลือกไฟล์</span><span>→</span>'
    '<span class="workflow-step">② อ่านเอกสาร</span><span>→</span>'
    '<span class="workflow-step">③ ตรวจผล</span><span>→</span>'
    '<span class="workflow-step">④ ดาวน์โหลด</span></div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("ตั้งค่าการอ่านเอกสาร")
    st.caption("ขั้นที่ 1: เลือกเครื่องมืออ่านข้อความจาก PDF")
    available_engines = available_ocr_engines()
    if not available_engines:
        st.error("ยังไม่พบ OCR engine: ติดตั้ง Tesseract พร้อมภาษาไทย หรือกำหนด TYPHOON_OCR_API_KEY")
        st.stop()
        available_engines = ["tesseract"]
    default_engine = available_engines[0]
    engine = st.radio(
        "เครื่องมืออ่านเอกสาร",
        available_engines,
        index=available_engines.index(default_engine),
        key="ocr_engine",
        format_func=lambda value: {
            "typhoon": "Typhoon OCR — แนะนำสำหรับเอกสารไทย",
            "tesseract": "Tesseract — อ่านในเครื่อง / Offline",
        }[value],
    )
    engine = usable_ocr_engine(engine)
    if not typhoon_api_key():
        st.info("ไม่มี Typhoon API key — ระบบใช้ Tesseract อัตโนมัติ")
    language = st.selectbox(
        "ภาษา",
        ["tha+eng", "tha", "eng"],
        index=0,
        format_func=lambda value: {"tha": "ไทย", "tha+eng": "ไทย + อังกฤษ", "eng": "อังกฤษ"}[value],
    )
    mode = st.selectbox(
        "การประมวลผลภาพสำหรับ OCR",
        ["best", "fast"],
        index=1,
        format_func=lambda value: "คมชัดสูง" if value == "best" else "เร็ว",
        help="มีผลกับ Tesseract เท่านั้น ระบบจะอ่านข้อความจากภาพเพื่อหลีกเลี่ยง text layer ของ PDF ที่อาจถอดภาษาไทยผิด",
    )
    if engine == "tesseract":
        psm = st.selectbox(
            "รูปแบบหน้าเอกสาร",
            PSM_VALUES,
            index=0,
            format_func=lambda value: {"3": "เอกสารทั่วไป", "6": "บล็อกข้อความ / ตาราง", "4": "คอลัมน์เดียว"}[value],
            help="เลือกตามลักษณะการจัดวางข้อความในหน้าเอกสาร มีผลเฉพาะกับ Tesseract",
        )
    else:
        psm = "6"
    oem = OEM_LSTM_ONLY
    gemini_key_available = bool(gemini_api_key())
    if not gemini_key_available:
        # Reset a stale browser/session value so a previous checked state
        # cannot enable a provider that is unavailable in this deployment.
        st.session_state["use_gemini_correction"] = False
    elif "use_gemini_correction" not in st.session_state:
        st.session_state["use_gemini_correction"] = True
    st.caption("ขั้นที่ 2: เลือกให้ Gemini ตรวจแก้ข้อความไทยหลังอ่านเอกสาร")
    use_gemini_correction = st.checkbox(
        "ใช้ Gemini ตรวจแก้สระและวรรณยุกต์",
        key="use_gemini_correction",
        disabled=not gemini_key_available,
        help="Gemini จะใช้ภาพ PDF ร่วมกับข้อความที่อ่านได้ เพื่อแก้ปัญหา เช่น กำร/นำยกรัฐมนตรี/ว่ำ",
    )
    gemini_model = GEMINI_MODEL
    if use_gemini_correction and gemini_key_available:
        try:
            available_gemini_models = list_gemini_models()
            if available_gemini_models:
                recommended_models = [
                    model for model in available_gemini_models
                    if gemini_model_is_recommended(str(model["name"]))
                ]
                show_advanced_models = st.checkbox(
                    "แสดง Gemini รุ่นขั้นสูงทั้งหมด",
                    value=False,
                    help="เปิดเพื่อดูรุ่นเฉพาะทาง เช่น TTS, Image, Robotics และรุ่นทดลอง",
                )
                selectable_models = available_gemini_models if show_advanced_models else recommended_models
                model_names = [str(model["name"]) for model in selectable_models]
                if not model_names:
                    st.warning("ไม่พบรุ่นที่แนะนำ จะแสดงโมเดลทั้งหมดแทน")
                    selectable_models = available_gemini_models
                    model_names = [str(model["name"]) for model in selectable_models]
                default_index = (
                    model_names.index(GEMINI_MODEL)
                    if GEMINI_MODEL in model_names else 0
                )
                gemini_model = st.selectbox(
                    "รุ่น Gemini สำหรับตรวจแก้",
                    model_names,
                    index=default_index,
                    format_func=gemini_model_option_label,
                    help="ค่าเริ่มต้นจะแสดงเฉพาะรุ่นที่เหมาะกับการตรวจแก้ข้อความและภาพ PDF",
                )
                profile = gemini_model_profile(gemini_model)
                st.info(f"เหมาะกับ: {profile['best_for']}")
            else:
                st.error("ไม่พบ Gemini model ที่รองรับ generateContent")
        except Exception as exc:
            st.error(f"ตรวจสอบ Gemini model ไม่สำเร็จ: {exc}")
            gemini_model = st.text_input("ชื่อ Gemini model สำรอง", value=GEMINI_MODEL)
    if engine == "typhoon":
        if typhoon_api_key():
            st.success("พร้อมใช้ Typhoon OCR")
            st.caption("ภาพแต่ละหน้าจะถูกส่งไปยัง Typhoon OCR API")
        else:
            st.warning("ยังไม่ได้ตั้งค่า TYPHOON_OCR_API_KEY — กรุณาเลือก Tesseract หรือเพิ่ม API key ใน Render Environment Variables")
    elif tesseract_path():
        st.success("พร้อมใช้ Tesseract ในเครื่อง")
    else:
        st.error("ไม่พบ Tesseract")
    with st.expander("รายละเอียดการประมวลผล", expanded=False):
        if use_gemini_correction:
            st.warning("Gemini จะส่ง PDF และข้อความ OCR ไปยัง Google Gemini อาจใช้โควตา API และไม่ควรใช้กับเอกสารลับหากไม่ได้รับอนุญาต")
        elif not gemini_key_available:
            st.caption("ยังไม่ได้ตั้งค่า GEMINI_API_KEY — ระบบจะอ่านเอกสารด้วย OCR ตามปกติ")
        st.caption("รองรับ PDF สูงสุด 50 MB / 100 หน้า")

uploaded = st.file_uploader(
    "เลือกไฟล์ PDF เพื่อเริ่มต้น",
    type=["pdf"],
    accept_multiple_files=False,
    help="รองรับเฉพาะไฟล์ .pdf เท่านั้น หากมีไฟล์ Word ให้บันทึกเป็น PDF ก่อน",
)
uploaded_bytes = uploaded.getvalue() if uploaded is not None else b""
uploaded_fingerprint = hashlib.sha256(uploaded_bytes).hexdigest() if uploaded_bytes else ""
previous_fingerprint = st.session_state.get("ocr_file_fingerprint", "")
if uploaded_fingerprint and previous_fingerprint and uploaded_fingerprint != previous_fingerprint:
    st.session_state.pop("ocr_text", None)
    st.session_state.pop("ocr_status", None)
    st.session_state.pop("docx_bytes", None)
    st.session_state.pop("docx_result_fingerprint", None)
    st.session_state.pop("docx_export_mode", None)
    st.session_state.pop("target_pages_processed", None)
    st.session_state.pop("ocr_page_cache", None)
    st.session_state["ocr_file_fingerprint"] = uploaded_fingerprint
elif uploaded_fingerprint and not previous_fingerprint:
    st.session_state["ocr_file_fingerprint"] = uploaded_fingerprint

total_doc_pages = 0
if uploaded_bytes:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    try:
        _temp_doc = fitz.open(stream=uploaded_bytes, filetype="pdf")
        total_doc_pages = _temp_doc.page_count
        _temp_doc.close()
    except Exception:
        total_doc_pages = 0

target_pages_to_process = None
effective_use_gemini = use_gemini_correction
effective_engine = engine

if uploaded is not None and total_doc_pages > 0:
    st.markdown("##### ⚙️ ตั้งค่าก่อนเริ่มอ่านเอกสาร (ควบคุมและประหยัด Token)")
    cfg_col1, cfg_col2 = st.columns([1, 1], gap="medium")
    with cfg_col1:
        st.markdown("**1. เป้าหมายการแปลงไฟล์**")
        processing_goal = st.radio(
            "เลือกเป้าหมายการประมวลผล:",
            options=[
                "🟢 เอกสารคุณภาพสูง (ใช้ AI ตรวจแก้ / จัดรูปเล่มสวยงาม)",
                "🟡 ประหยัด Token 100% (ใช้ Tesseract OCR ในเครื่อง - ฟรี)",
            ],
            index=0 if gemini_key_available else 1,
            key="processing_goal_radio",
            help="หากเลือกแบบประหยัด Token ระบบจะปิดการเรียก Gemini/Typhoon ชั่วคราวและอ่านด้วย OCR ออฟไลน์ในเครื่อง 100%",
        )
        if processing_goal.startswith("🟡"):
            effective_use_gemini = False
            effective_engine = "tesseract"
        else:
            effective_use_gemini = use_gemini_correction
            effective_engine = engine

    with cfg_col2:
        st.markdown(f"**2. เลือกช่วงหน้า (พบทั้งหมด {total_doc_pages} หน้า)**")
        page_selection_mode = st.radio(
            "หน้าที่ต้องการอ่าน:",
            options=["ทุกหน้าในเอกสาร", "ระบุช่วงหน้าเอง (Custom Range)"],
            index=0,
            horizontal=True,
            key="page_selection_mode_radio",
        )
        if page_selection_mode == "ระบุช่วงหน้าเอง (Custom Range)":
            custom_range_str = st.text_input(
                "ระบุหน้า (เช่น 1-3, 5):",
                value=f"1-{min(total_doc_pages, 3)}",
                key="custom_page_range_input",
                help=f"กรอกหมายเลขหน้าระหว่าง 1 ถึง {total_doc_pages} คั่นด้วยเครื่องหมายจุลภาค (,) หรือยัติภังค์ (-)",
            )
            target_pages_to_process = parse_page_range(custom_range_str, total_doc_pages)
            saved_percent = round((1.0 - len(target_pages_to_process) / total_doc_pages) * 100)
            if saved_percent > 0:
                st.caption(f"⚡ จะประมวลผล {len(target_pages_to_process)} จาก {total_doc_pages} หน้า (ประหยัด Token ไปได้ ~{saved_percent}%)")
            else:
                st.caption(f"จะประมวลผลทั้งหมด {len(target_pages_to_process)} หน้า")
        else:
            target_pages_to_process = list(range(1, total_doc_pages + 1))
            st.caption(f"จะประมวลผลครบทั้ง {total_doc_pages} หน้า")

run = st.button("🔍 เริ่มอ่านเอกสาร", type="primary", use_container_width=True)

preview_column, result_column = st.columns(2, gap="medium")
with preview_column:
    st.markdown("#### PDF ต้นฉบับ")
    st.caption("เปิดค้างไว้เพื่อเทียบกับผล OCR")
    if uploaded is not None:
        render_pdf_viewer(uploaded_bytes, uploaded.name, height=PDF_VIEWER_HEIGHT)
    else:
        st.markdown(
            '<div class="pdf-empty">อัปโหลดไฟล์ PDF เพื่อแสดงตัวอย่างเอกสารที่นี่</div>',
            unsafe_allow_html=True,
        )

with result_column:
    st.markdown("#### ผล OCR")
    ocr_preview = st.empty()

if run:
    if uploaded is None:
        st.warning("กรุณาเลือกไฟล์ PDF ก่อนเริ่ม OCR")
        st.stop()
    progress = st.progress(0, text="กำลังเตรียมไฟล์... (0%)")

    def publish_ocr_preview(current_text: str) -> None:
        if current_text:
            ocr_preview.markdown(
                f'<div class="ocr-live-preview">{html.escape(current_text)}</div>',
                unsafe_allow_html=True,
            )

    try:
        page_cache = st.session_state.setdefault("ocr_page_cache", {})
        text, errors, failed_images = extract_pdf(
            uploaded_bytes, language, mode, psm, oem, effective_engine,
            effective_use_gemini, gemini_model, progress,
            page_callback=publish_ocr_preview,
            pages_to_process=target_pages_to_process,
            page_cache=page_cache,
        )
        progress.progress(1.0, text="อ่านเอกสารเสร็จสมบูรณ์ (100%)")
        st.session_state["ocr_text"] = text
        st.session_state["ocr_result_fingerprint"] = uploaded_fingerprint
        st.session_state["target_pages_processed"] = target_pages_to_process
        ocr_preview.text_area("ผล OCR", text, height=OCR_EDITOR_HEIGHT, key="ocr_result_editor")
        document_errors = [(page_number, error) for page_number, error in errors if page_number is None]
        page_errors = [(page_number, error) for page_number, error in errors if page_number is not None]
        if effective_use_gemini and document_errors:
            result_status = "fallback"
            st.warning("อ่านเอกสารเสร็จแล้ว แต่ใช้ผล OCR ในเครื่องแทน Gemini เนื่องจากการตรวจแก้ทั้งเอกสารไม่สำเร็จ")
        elif page_errors:
            result_status = "partial"
            st.warning(f"อ่านเอกสารเสร็จแล้ว แต่มี {len(page_errors)} หน้าที่ควรตรวจสอบ")
        elif effective_use_gemini:
            result_status = "gemini"
            st.success("อ่านเอกสารและตรวจแก้ด้วย Gemini สำเร็จ")
        else:
            result_status = "local"
            st.success("อ่านเอกสารสำเร็จด้วย OCR ในเครื่อง (0 Token)")
        st.session_state["ocr_status"] = result_status
        stem = Path(uploaded.name).stem
        if document_errors:
            with st.expander(f"ปัญหาระดับเอกสาร ({len(document_errors)})"):
                for _, error in document_errors:
                    st.error(error)
        if page_errors:
            with st.expander(f"หน้าที่ควรตรวจสอบ ({len(page_errors)})"):
                for page_number, error in page_errors:
                    st.error(f"หน้า {page_number}: {error}")
                    failed_image = failed_images.get(page_number)
                    if failed_image is not None:
                        st.download_button(
                            f"ดาวน์โหลดภาพหน้า {page_number}",
                            data=failed_image,
                            file_name=f"{stem}-page-{page_number}-failed.png",
                            mime="image/png",
                            key=f"failed-{page_number}",
                        )
    except Exception as exc:
        progress.progress(1.0, text="OCR ไม่สำเร็จ (100%)")
        st.session_state.pop("ocr_text", None)
        st.session_state.pop("ocr_result_fingerprint", None)
        st.session_state["ocr_status"] = "failed"
        with st.expander("รายละเอียดข้อผิดพลาด", expanded=False):
            st.error(f"อ่านเอกสารไม่สำเร็จ: {type(exc).__name__}: {exc}")

# Keep the OCR editor visible by default, including before the first OCR run.
# This gives users a stable target for the result and keeps the two panels from
# jumping when OCR starts or when the selected file changes.
if not run:
    ocr_preview.text_area(
        "ผล OCR",
        st.session_state.get("ocr_text", ""),
        height=OCR_EDITOR_HEIGHT,
        key="ocr_result_editor",
    )

if (
    uploaded is not None
    and st.session_state.get("ocr_result_fingerprint") == uploaded_fingerprint
    and st.session_state.get("ocr_text")
):
    st.markdown("---")
    st.markdown("### 📥 ดาวน์โหลดผลลัพธ์ (DOCX)")

    current_ocr_text = st.session_state.get("ocr_result_editor") or st.session_state.get("ocr_text", "")
    stem = Path(uploaded.name).stem

    col_mode, col_btn = st.columns([3, 2], gap="medium")
    with col_mode:
        selected_mode = st.radio(
            "เลือกรูปแบบไฟล์ DOCX:",
            options=[
                "📄 จัดรูปเล่มอัตโนมัติ (Flowing — สระไม่ตก ไม่ตัด สวยงาม แนะนำสำหรับภาษาไทย)",
                "📐 จัดวางตามพิกัดหน้าเดิม (Layered Layout — วางกล่องข้อความตามพิกัด PDF)",
            ],
            index=0,
            key="docx_download_mode_radio",
            help="สามารถแก้ไขข้อความในกล่อง 'ผล OCR' ด้านบน แล้วดาวน์โหลดตามข้อความที่แก้ไขได้ทันที",
        )

    with col_btn:
        st.write("")
        st.write("")
        target_docx_bytes = None
        target_filename = ""
        target_help = ""

        with st.spinner("กำลังเตรียมไฟล์ DOCX..."):
            try:
                if selected_mode.startswith("📄"):
                    target_docx_bytes = export_pdf_flowing_docx(current_ocr_text)
                    target_filename = f"{stem}-flowing.docx"
                    target_help = "ไฟล์ Word จัดรูปเล่มอัตโนมัติ ใช้ฟอนต์ TH Sarabun New สระและวรรณยุกต์ครบถ้วน ไม่มั่ว แก้ไขง่าย"
                else:
                    processed_pages = st.session_state.get("target_pages_processed")
                    target_docx_bytes, report = export_pdf_editable_docx(
                        uploaded_bytes,
                        current_ocr_text,
                        page_numbers=tuple(processed_pages) if processed_pages else None,
                    )
                    target_filename = f"{stem}-layered.docx"
                    target_help = "ไฟล์ Word อิงพิกัดเดิม มีภาพพื้นหลังและกล่องข้อความลอยตามตำแหน่งเดิม"
            except Exception as export_err:
                st.error(f"เกิดข้อผิดพลาดในการสร้าง DOCX: {type(export_err).__name__}: {export_err}")

        if target_docx_bytes:
            st.download_button(
                "⬇️ ดาวน์โหลดไฟล์ DOCX",
                data=target_docx_bytes,
                file_name=target_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                help=target_help,
                use_container_width=True,
                type="primary",
            )
            st.caption(f"ขนาดไฟล์: {len(target_docx_bytes) / 1024:.1f} KB (อัปเดตตามข้อความล่าสุด)")
