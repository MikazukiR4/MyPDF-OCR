import io
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


load_dotenv()
load_dotenv(".env.typhoon", override=False)

MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_PAGES = 100
GEMINI_MAX_WORKERS = 3
TESSDATA_DIR = Path(__file__).resolve().parent / "tessdata"

OEM_LSTM_ONLY = "1"

PSM_VALUES = ["3", "6", "4"]

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


def _normalize_pdf_text(text: str) -> str:
    """Normalize PDF glyph artifacts without rewriting document content."""
    text = text.translate(str.maketrans({
        "\uf70a": "\u0e48",
        "\uf70b": "\u0e49",
        "\uf70e": "\u0e4c",
    }))
    text = unicodedata.normalize("NFC", text)
    text = re.sub(r"\u0e4d\s*\u0e32", "\u0e33", text)
    text = re.sub(r"([\u0e48-\u0e4b])\u0e33", "ำ\\1", text)
    text = re.sub(r"([\u0e48-\u0e4b])([\u0e38-\u0e3a])", r"\2\1", text)
    return text.strip()


def _docx_text_box(paragraph: object, text: str, left: float, top: float, width: float, height: float, font_size: float) -> None:
    """Add an editable, absolutely positioned Word text box to a paragraph."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import nsmap, qn

    # python-docx does not register the legacy VML namespace by default.
    # Register it before creating the absolutely positioned text box.
    nsmap.setdefault("v", "urn:schemas-microsoft-com:vml")
    nsmap.setdefault("o", "urn:schemas-microsoft-com:office:office")

    text = _normalize_pdf_text(text)
    if not text:
        return
    run = paragraph.add_run()
    pict = OxmlElement("w:pict")
    shape = OxmlElement("v:shape")
    shape.set("id", f"EditablePdfText{len(paragraph._p) + 1}")
    shape.set(
        "style",
        "position:absolute;"
        f"margin-left:{left:.2f}pt;margin-top:{top:.2f}pt;"
        f"width:{max(2.0, width):.2f}pt;height:{max(8.0, height):.2f}pt;"
        "mso-wrap-style:none;mso-fit-shape-to-text:false;v-text-anchor:top;",
    )
    shape.set("stroked", "f")
    shape.set("filled", "f")
    textbox = OxmlElement("v:textbox")
    textbox.set("inset", "0,0,0,0")
    content = OxmlElement("w:txbxContent")
    text_paragraph = OxmlElement("w:p")
    paragraph_properties = OxmlElement("w:pPr")
    line_spacing = OxmlElement("w:spacing")
    line_spacing.set(qn("w:line"), str(max(80, round(height * 20))))
    line_spacing.set(qn("w:lineRule"), "exact")
    paragraph_properties.append(line_spacing)
    text_paragraph.append(paragraph_properties)
    text_run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for attribute in ("ascii", "hAnsi", "cs", "eastAsia"):
        fonts.set(qn(f"w:{attribute}"), "Tahoma")
    run_properties.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), str(max(2, round(font_size * 2))))
    run_properties.append(size)
    text_run.append(run_properties)
    for part_index, line_part in enumerate(text.splitlines() or [text]):
        if part_index:
            line_break = OxmlElement("w:br")
            line_break.set(qn("w:type"), "textWrapping")
            text_run.append(line_break)
        text_node = OxmlElement("w:t")
        text_node.set(qn("xml:space"), "preserve")
        text_node.text = line_part
        text_run.append(text_node)
    text_paragraph.append(text_run)
    content.append(text_paragraph)
    textbox.append(content)
    shape.append(textbox)
    pict.append(shape)
    run._r.append(pict)


def _page_text_boxes(page: object) -> list[tuple[str, float, float, float, float, float]]:
    """Return editable boxes, rejecting corrupt PDF text layers before OCR."""
    text_boxes: list[tuple[str, float, float, float, float, float]] = []
    text_dict = page.get_text("dict")
    for block in text_dict.get("blocks", []):
        for line in block.get("lines", []):
            spans = [span for span in line.get("spans", []) if span.get("text", "").strip()]
            if not spans:
                continue
            x0 = min(float(span["bbox"][0]) for span in spans)
            y0 = min(float(span["bbox"][1]) for span in spans)
            x1 = max(float(span["bbox"][2]) for span in spans)
            y1 = max(float(span["bbox"][3]) for span in spans)
            text = "".join(span.get("text", "") for span in spans).strip()
            font_size = max(4.0, max(float(span.get("size", y1 - y0)) for span in spans))
            text_boxes.append((text, x0, y0, max(2.0, x1 - x0), max(8.0, y1 - y0), font_size))

    raw_text = "\n".join(item[0] for item in text_boxes)
    thai_chars = len(re.findall(r"[ก-๙]", raw_text))
    # Thai PDF text layers are frequently font-glyph mappings rather than
    # reliable Unicode. Prefer OCR boxes for Thai so Word receives real Thai
    # characters instead of the PDF's broken extraction sequence.
    if thai_chars >= 20 or _pdf_text_layer_is_corrupt(raw_text):
        ocr_boxes = _tesseract_text_boxes(page)
        if ocr_boxes:
            return ocr_boxes
    if text_boxes:
        return [(_normalize_pdf_text(item[0]), *item[1:]) for item in text_boxes]
    return []


def _pdf_text_layer_is_corrupt(text: str) -> bool:
    """Detect common Thai PDF extraction failures before creating Word text."""
    if not text.strip():
        return True
    if "�" in text or any("\uf700" <= char <= "\uf8ff" for char in text):
        return True
    if re.search(r"[\x80-\x9f]|(?:Ã|Â|â|เธ)[\x80-\x9f]", text):
        return True
    thai_chars = len(re.findall(r"[ก-๙]", text))
    if thai_chars < 40:
        return False
    single_thai_tokens = len(re.findall(r"(?<![ก-๙])[ก-๙](?=\s|$)", text))
    return single_thai_tokens / thai_chars > 0.08


def _tesseract_text_boxes(page: object) -> list[tuple[str, float, float, float, float, float]]:
    """Run Thai+English Tesseract and return line boxes in PDF points."""
    boxes: list[tuple[str, float, float, float, float, float]] = []
    command = tesseract_path()
    if not command:
        return boxes
    image = page.get_pixmap(dpi=300, alpha=False).tobytes("png")
    with tempfile.TemporaryDirectory(prefix="docx_ocr_") as temp_dir:
        image_path = Path(temp_dir) / "page.png"
        image_path.write_bytes(image)
        result = subprocess.run(
            [command, str(image_path), "stdout", "-l", "tha+eng", "--psm", "3", "tsv"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
            check=False,
        )
    if result.returncode != 0:
        return boxes
    scale = 72.0 / 300.0
    rows = csv.DictReader(io.StringIO(result.stdout.decode("utf-8", errors="replace")), delimiter="\t")
    lines: dict[tuple[str, str, str, str], list[dict[str, str]]] = {}
    for row in rows:
        text = row.get("text", "").strip()
        try:
            confidence = float(row.get("conf", "-1"))
        except ValueError:
            confidence = -1
        if not text or confidence < 0:
            continue
        key = (row.get("block_num", ""), row.get("par_num", ""), row.get("line_num", ""), row.get("page_num", ""))
        lines.setdefault(key, []).append(row)
    for words in lines.values():
        left = min(int(word["left"]) for word in words) * scale
        top = min(int(word["top"]) for word in words) * scale
        right = max(int(word["left"]) + int(word["width"]) for word in words) * scale
        bottom = max(int(word["top"]) + int(word["height"]) for word in words) * scale
        boxes.append((" ".join(word["text"] for word in words), left, top, right - left, bottom - top, max(5.0, (bottom - top) * 0.85)))
    return boxes


@st.cache_data(show_spinner=False)
def export_pdf_editable_docx(pdf_bytes: bytes) -> bytes:
    """Create an editable DOCX by positioning real Word text over each PDF page."""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.shared import Inches

    if not pdf_bytes:
        raise ValueError("ไม่พบข้อมูล PDF สำหรับสร้าง DOCX ที่แก้ไขได้")
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
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
            section.top_margin = section.bottom_margin = Inches(0)
            section.left_margin = section.right_margin = Inches(0)
            section.header_distance = section.footer_distance = Inches(0)
            paragraph = word_document.add_paragraph()
            paragraph.paragraph_format.space_before = Inches(0)
            paragraph.paragraph_format.space_after = Inches(0)
            for text, left, top, width, height, font_size in _page_text_boxes(page):
                _docx_text_box(paragraph, text, left, top, width, height, font_size)
        word_document.core_properties.title = "MyOCR PDF editable DOCX export"
        word_document.core_properties.subject = "PDF text reconstructed as editable Word text"
        word_document.core_properties.comments = (
            "Text is positioned using PDF text-layer or OCR bounding boxes. "
            "Non-text graphics are not reproduced as a full-page image."
        )
        word_document.save(output)
        return output.getvalue()
    finally:
        pdf_document.close()


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
        image = ImageOps.autocontrast(image, cutoff=0.3)
        image = image.filter(ImageFilter.SHARPEN)
        image = image.point(lambda p: 255 if p > 235 else (0 if p < 45 else p))
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def clean_thai_ocr_text(text: str) -> str:
    """Normalize Thai text and remove common floating-diacritic OCR garbage."""
    if not text:
        return ""
    # Some Thai PDF fonts expose tone marks through Private Use Area glyphs.
    # Convert only the mappings observed in the embedded text layer so that
    # normal Unicode text is left unchanged.
    text = text.translate(str.maketrans({
        "\uf70a": "\u0e48",  # mai ek: ่
        "\uf70b": "\u0e49",  # mai tho: ้
        "\uf70e": "\u0e4c",  # thanthakhat: ์
    }))
    text = unicodedata.normalize("NFC", text)
    text = re.sub(r"\u0e4d\s*\u0e32", "\u0e33", text)
    text = re.sub(r"([\u0e48-\u0e4b])\u0e33", "ำ\\1", text)
    # A few legacy Thai fonts place a tone mark before a below vowel
    # (for example ผู้). Put the tone mark after the vowel (ผู้).
    text = re.sub(r"([\u0e48-\u0e4b])([\u0e38-\u0e3a])", r"\2\1", text)
    text = re.sub(r"<file_content\b[^>]*>.*?</file_content>", "", text, flags=re.DOTALL)
    text = re.sub(r"<file_content\b[^>]*>.*$", "", text, flags=re.DOTALL)

    output = []
    for line in text.splitlines():
        compact = re.sub(r"\s+", "", line)
        tokens = line.split()
        thai_letters = len(re.findall(r"[ก-ฮ]", compact))
        ascii_letters = len(re.findall(r"[A-Za-z]", compact))
        symbols = len(re.findall(r"[=/_~^|<>*+`\"'“”‘’.,:;!?-]", compact))
        short_tokens = sum(len(token) <= 2 for token in tokens)
        is_floating_line = (
            len(tokens) >= 3
            and short_tokens / len(tokens) >= 0.75
            and thai_letters <= 3
        ) or (
            len(compact) <= 8
            and thai_letters <= 1
            and (ascii_letters or symbols)
            and not re.fullmatch(r"[0-9๐-๙.,/-]+", compact or " ")
        )
        if compact and not is_floating_line:
            output.append(re.sub(r"[ \t]{3,}", "  ", line.rstrip()))

    text = "\n".join(output)
    replacements = (
        (r"(?m)^(\s*)เรอง(\s+)", r"\1เรื่อง\2"),
        (r"(?m)^(\s*)(?:อางถง|อางถึง|อ้างถง)(\s+)", r"\1อ้างถึง\2"),
        (r"(?m)^(\s*)(?:สิงที่ส่งมาด้วย|สิงทีส่งมาด้วย|สิ่งที่สงมาดวย)(\s+)", r"\1สิ่งที่ส่งมาด้วย\2"),
        (r"(?:หนงสอจงหวด|หนงสือจงหวด|หนงสอจังหวัด|หนังสือจงหวด)", "หนังสือจังหวัด"),
        (r"(\s)ท(\s+(?:ศธ|สส|มท|กค|นร|อว|พม|สธ|กษ|รง|ยธ|ทส|วธ)\s+[๐-๙0-9])", r"\1ที่\2"),
        (r"(?:ลงวันท|ลงวนท|ลงวนที่)(\s+)", r"ลงวันที่\1"),
        (r"(?<=ลงวันที่\s)mo(?=\s+[ก-๙])", "๓๐"),
        (r"(?:ประจำปงบประมาณ|ประจําปงบประมาณ)", "ประจำปีงบประมาณ"),
        (r"(?:we|wa|w\.ศ|พ\.ค|พ\.ศ)\.\s*(?=[๐-๙]{4}|\d{4})", "พ.ศ. "),
        (r"การขับเคลือน", "การขับเคลื่อน"),
        (r"ในการนี(?=\s|$)", "ในการนี้"),
        (r"ในส่วน(?:ทีเกียวข้อง|ที่เกียวข้อง|ทีเกี่ยวข้อง)", "ในส่วนที่เกี่ยวข้อง"),
        (r"(?:อำเกอ|อําเกอ)", "อำเภอ"),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


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

    if not corrected or len(corrected) < max(10, round(len(source_text) * 0.25)):
        raise RuntimeError("Gemini ส่งข้อความแก้ไขกลับมาสั้นผิดปกติ จึงคงข้อความเดิมไว้")
    return clean_thai_ocr_text(corrected)


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

    if not corrected or len(corrected) < max(10, round(len(source_text) * 0.25)):
        raise RuntimeError("Gemini document correction returned unexpectedly short text")
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

        output: list[str] = []
        errors: list[tuple[int | None, str]] = []
        failed_images: dict[int, bytes] = {}
        worker = None if engine == "typhoon" else TesseractWorker(language, mode, psm, oem)
        gemini_document_mode = use_gemini_correction and len(pdf_bytes) <= MAX_FILE_BYTES

        def publish_page() -> None:
            if page_callback is not None:
                page_callback("\n\n".join(output))

        try:
            # Process small batches so Gemini calls overlap while completed
            # pages are still published incrementally in document order.
            batch_size = GEMINI_MAX_WORKERS if use_gemini_correction and not gemini_document_mode else 1
            for batch_start in range(1, page_count + 1, batch_size):
                batch_end = min(page_count, batch_start + batch_size - 1)
                page_data: dict[int, tuple[str, str, bytes | None]] = {}
                gemini_jobs: dict[int, tuple[bytes, str]] = {}
                for page_number in range(batch_start, batch_end + 1):
                    ratio = (page_number - 1) / page_count
                    progress.progress(ratio * 0.5, text=f"กำลังเตรียมหน้า {page_number}/{page_count}")
                    page = document[page_number - 1]
                    image_bytes = render_page(document, page_number, mode)
                    try:
                        text = typhoon_transcribe(image_bytes) if engine == "typhoon" else worker.recognize(image_bytes)
                        if not text:
                            raise RuntimeError("OCR engine ส่งข้อความว่างกลับมา")
                        page_data[page_number] = (text, engine, image_bytes)
                        if use_gemini_correction and not gemini_document_mode and text and not text.startswith("["):
                            gemini_jobs[page_number] = (image_bytes, text)
                    except Exception as exc:
                        errors.append((page_number, f"{type(exc).__name__}: {exc}"))
                        failed_images[page_number] = image_bytes

                futures = {}
                if gemini_jobs:
                    with ThreadPoolExecutor(max_workers=min(GEMINI_MAX_WORKERS, len(gemini_jobs))) as executor:
                        for page_number, (image_bytes, source_text) in gemini_jobs.items():
                            futures[page_number] = executor.submit(
                                gemini_correct_text, image_bytes, source_text, gemini_model
                            )
                        for page_number in range(batch_start, batch_end + 1):
                            if page_number not in page_data:
                                output.append(f"----- หน้า {page_number} -----\n[อ่านหน้านี้ไม่สำเร็จ]")
                            else:
                                text, method, _ = page_data[page_number]
                                if page_number in futures:
                                    try:
                                        text = futures[page_number].result()
                                        method = "Gemini corrected"
                                    except Exception as correction_error:
                                        errors.append((page_number, f"Gemini ตรวจแก้ไม่สำเร็จ: {correction_error}"))
                                output.append(f"----- หน้า {page_number} | {method} -----\n{text}")
                            progress.progress(
                                0.5 + (page_number / page_count) * 0.5,
                                text=f"อ่านหน้า {page_number}/{page_count} เสร็จแล้ว",
                            )
                            publish_page()
                else:
                    for page_number in range(batch_start, batch_end + 1):
                        if page_number not in page_data:
                            output.append(f"----- หน้า {page_number} -----\n[อ่านหน้านี้ไม่สำเร็จ]")
                        else:
                            text, method, _ = page_data[page_number]
                            output.append(f"----- หน้า {page_number} | {method} -----\n{text}")
                        progress.progress(
                            0.5 + (page_number / page_count) * 0.5,
                            text=f"อ่านหน้า {page_number}/{page_count} เสร็จแล้ว",
                        )
                        publish_page()
        finally:
            if worker is not None:
                worker.terminate()

        local_output = "\n\n".join(output)
        if gemini_document_mode and local_output:
            progress.progress(0.85, text="กำลังส่ง PDF ทั้งเอกสารให้ Gemini ตรวจแก้")
            try:
                corrected_document = gemini_correct_document(pdf_bytes, local_output, gemini_model)
                progress.progress(0.98, text="ตรวจสอบลำดับและ marker หน้าแล้ว")
                return corrected_document, errors, failed_images
            except Exception as correction_error:
                errors.append((None, f"Gemini ตรวจแก้ทั้งเอกสารไม่สำเร็จ: {correction_error}"))
        return local_output, errors, failed_images
    finally:
        document.close()


def render_pdf_viewer(pdf_bytes: bytes, filename: str, height: int = 720) -> None:
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
          html, body, #pdf-viewer {{ width:100%; height:100%; min-height:620px; margin:0; }}
          body {{ overflow:hidden; background:#172554; }}
          #viewer-error {{ padding:1rem; color:#fecaca; font:14px sans-serif; }}
        </style>
      </head>
      <body>
        <div id="pdf-viewer" aria-label="ตัวอย่าง PDF: {safe_name}"></div>
        <div id="viewer-error" role="alert" hidden></div>
        <script type="module">
          import EmbedPDF from "https://cdn.jsdelivr.net/npm/@embedpdf/snippet@2/dist/embedpdf.js";
          const payload = {{ data: "{encoded_pdf}" }};
          const target = document.getElementById("pdf-viewer");
          const errorTarget = document.getElementById("viewer-error");
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
        .block-container { max-width: 1200px; padding-top: 3.5rem; }
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
        /* Dark mode: keep every surface and text role readable together. */
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] { background:#0b1220 !important; color:#f8fafc !important; }
        .block-container { color:#f8fafc !important; }
        .ocr-card { background:#111827 !important; border-color:#334155 !important; box-shadow:0 8px 24px rgba(0,0,0,.28) !important; }
        .mypdf-name { color:#f8fafc !important; }
        .subtitle, .small-note, .stCaption, .stMarkdown, .stMarkdown p, label, label p { color:#cbd5e1 !important; }
        div[data-testid="stFileUploader"], div[data-testid="stFileUploader"] section, div[data-testid="stFileUploader"] section > div { background:#111827 !important; border-color:#64748b !important; color:#f8fafc !important; }
        div[data-testid="stFileUploader"] small, div[data-testid="stFileUploader"] span, div[data-testid="stFileUploader"] label { color:#cbd5e1 !important; }
        div[data-testid="stTextArea"] textarea { background:#020617 !important; color:#ffffff !important; border:2px solid #64748b !important; -webkit-text-fill-color:#ffffff !important; caret-color:#ffffff !important; font-size:16px !important; line-height:1.7 !important; }
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
            .block-container { padding-left:1rem; padding-right:1rem; padding-top:2.5rem; }
            .mypdf-header { margin-bottom:1.25rem; }
            .mypdf-name { font-size:25px; }
            .workflow-steps { gap:5px; font-size:12px; }
            .workflow-steps > span:not(.workflow-step) { display:none; }
            .workflow-step { padding:5px 7px; }
            .ocr-card { padding:16px; }
            div[data-testid="stTextArea"] textarea { font-size:15px !important; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title="MyOCR PDF — Thai", page_icon="📄", layout="centered")
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
    st.caption("ขั้นที่ 2: เลือกให้ Gemini ตรวจแก้ข้อความไทยหลังอ่านเอกสาร")
    use_gemini_correction = st.checkbox(
        "ใช้ Gemini ตรวจแก้สระและวรรณยุกต์",
        value=False,
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
                selected_metadata = next(model for model in selectable_models if model["name"] == gemini_model)
                st.markdown("**รุ่นที่เลือก**")
                st.caption(str(selected_metadata["description"]))
                st.caption(f"ประเภท: {gemini_model_kind(gemini_model)}")
                st.caption(
                    f"รับเข้าได้: {format_token_limit(selected_metadata['input_token_limit'])} · "
                    f"ตอบกลับได้: {format_token_limit(selected_metadata['output_token_limit'])}"
                )
                profile = gemini_model_profile(gemini_model)
                st.markdown("**คะแนนโดยประมาณสำหรับงานตรวจแก้ OCR**")
                st.markdown(
                    gemini_rating_html("คุณภาพการแก้ข้อความ", profile["quality"], "1 ดาว = ต่ำ, 3 ดาว = ปานกลาง, 5 ดาว = สูง"),
                    unsafe_allow_html=True,
                )
                st.markdown(
                    gemini_rating_html("ความเร็ว", profile["speed"], "1 ดาว = ช้า, 3 ดาว = ปานกลาง, 5 ดาว = เร็วมาก"),
                    unsafe_allow_html=True,
                )
                st.markdown(
                    gemini_rating_html("ความเข้าใจภาพ/PDF", profile["vision"], "1 ดาว = จำกัด, 3 ดาว = ปานกลาง, 5 ดาว = เหมาะกับภาพมาก"),
                    unsafe_allow_html=True,
                )
                st.markdown(
                    gemini_rating_html("ความประหยัดโดยประมาณ", profile["economy"], "1 ดาว = ใช้ทรัพยากรมาก, 3 ดาว = ปานกลาง, 5 ดาว = ประหยัดมาก"),
                    unsafe_allow_html=True,
                )
                st.info(f"เหมาะกับ: {profile['best_for']}")
                st.caption("คะแนนดาวเป็นคำแนะนำโดยประมาณจากตระกูลโมเดล ไม่ใช่คะแนนทางการของ Google และราคา/โควตาจริงขึ้นกับบัญชีของคุณ")
                st.caption("PDF: เมื่อเริ่มตรวจแก้ ระบบจะส่ง PDF ต้นฉบับและข้อความ OCR ทั้งเอกสารไปยังโมเดลที่เลือกในคำขอเดียว")
                st.caption(f"แสดง {len(selectable_models)} จาก {len(available_gemini_models)} โมเดลที่ใช้ได้")
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
    if use_gemini_correction:
        st.warning("Gemini จะตรวจ PDF ทั้งเอกสารในคำขอเดียว โดยส่งไฟล์และข้อความ OCR ไปยัง Google Gemini อาจใช้โควตา API และไม่ควรใช้กับเอกสารลับหากไม่ได้รับอนุญาต")
    elif not gemini_key_available:
        st.info("ยังไม่ได้ตั้งค่า GEMINI_API_KEY — ปิดการตรวจแก้ Gemini ไว้ และจะอ่านเอกสารด้วย OCR ตามปกติ")
    st.caption("รองรับ PDF สูงสุด 50 MB / 100 หน้า")

st.markdown('<div class="ocr-card">', unsafe_allow_html=True)
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
    st.session_state["ocr_file_fingerprint"] = uploaded_fingerprint
elif uploaded_fingerprint and not previous_fingerprint:
    st.session_state["ocr_file_fingerprint"] = uploaded_fingerprint
if uploaded is not None:
    st.success(f"เลือกไฟล์แล้ว: {uploaded.name} · {uploaded.size / 1024:.1f} KB")
    if uploaded_fingerprint != st.session_state.get("ocr_result_fingerprint"):
        st.info("ไฟล์นี้ยังไม่ได้ประมวลผล หรือไฟล์ถูกเปลี่ยนแล้ว กรุณากดเริ่มอ่านเอกสาร")
else:
    st.info("เลือกไฟล์ PDF เพื่อเริ่มต้น")
if uploaded is not None and use_gemini_correction:
    st.info("ความเป็นส่วนตัว: PDF ทั้งเอกสารและข้อความ OCR จะถูกส่งไปยัง Google Gemini เพื่อช่วยตรวจแก้")
elif uploaded is not None and engine == "typhoon":
    st.info("ความเป็นส่วนตัว: ภาพแต่ละหน้าจะถูกส่งไปยัง Typhoon OCR API")
run = st.button("🔍 เริ่มอ่านเอกสาร", type="primary", use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

comparison_text = None
if uploaded is not None:
    st.markdown("### ตรวจเทียบ PDF กับข้อความ OCR")
    st.caption("PDF ต้นฉบับ — viewer ปรับเต็มความกว้างเพื่อไม่ให้หน้ากระดาษถูกตัด")
    render_pdf_viewer(uploaded_bytes, uploaded.name, height=620)
    st.caption("ข้อความที่ถอดได้")
    current_result = st.session_state.get("ocr_text", "")
    result_fingerprint = st.session_state.get("ocr_result_fingerprint")
    if result_fingerprint != uploaded_fingerprint:
        current_result = "ไฟล์นี้ยังไม่มีผล OCR — กด ‘เริ่มอ่านเอกสาร’ เพื่อประมวลผล"
    comparison_text = st.empty()
    comparison_text.text_area(
        "ผล OCR",
        current_result,
        height=420,
        label_visibility="visible",
        key="ocr-preview",
    )

if run:
    if uploaded is None:
        st.warning("กรุณาเลือกไฟล์ PDF ก่อนเริ่ม OCR")
        st.stop()
    progress = st.progress(0, text="กำลังเตรียมไฟล์... (0%)")
    st.markdown("### ผลระหว่างประมวลผล")
    live_output = st.empty()
    live_output.code(
        "กำลังเตรียมผลหน้าแรก... ผลลัพธ์นี้อาจถูกตรวจแก้ใหม่เมื่ออ่านครบทั้งเอกสาร",
        language=None,
    )

    def publish_live_text(text: str) -> None:
        # Do not create another keyed widget from the callback. Streamlit
        # treats each text_area call as a new widget, even in the same
        # placeholder, so reusing key="ocr-live" causes DuplicateElementKey.
        live_output.code(text, language=None)

    try:
        text, errors, failed_images = extract_pdf(
            uploaded_bytes, language, mode, psm, oem, engine,
            use_gemini_correction, gemini_model, progress, publish_live_text,
        )
        progress.progress(1.0, text="อ่านเอกสารเสร็จสมบูรณ์ (100%)")
        st.session_state["ocr_text"] = text
        st.session_state["ocr_result_fingerprint"] = uploaded_fingerprint
        st.markdown("### ผล OCR ฉบับสมบูรณ์")
        st.text_area(
            "ผล OCR",
            text,
            height=680,
            label_visibility="visible",
            key="ocr-final",
        )
        document_errors = [(page_number, error) for page_number, error in errors if page_number is None]
        page_errors = [(page_number, error) for page_number, error in errors if page_number is not None]
        if use_gemini_correction and document_errors:
            result_status = "fallback"
            st.warning("อ่านเอกสารเสร็จแล้ว แต่ใช้ผล OCR ในเครื่องแทน Gemini เนื่องจากการตรวจแก้ทั้งเอกสารไม่สำเร็จ")
        elif page_errors:
            result_status = "partial"
            st.warning(f"อ่านเอกสารเสร็จแล้ว แต่มี {len(page_errors)} หน้าที่ควรตรวจสอบ")
        elif use_gemini_correction:
            result_status = "gemini"
            st.success("อ่านเอกสารและตรวจแก้ด้วย Gemini สำเร็จ")
        else:
            result_status = "local"
            st.success("อ่านเอกสารสำเร็จด้วย OCR ในเครื่อง")
        st.session_state["ocr_status"] = result_status
        stem = Path(uploaded.name).stem
        st.download_button(
            "⬇️ ดาวน์โหลดผล OCR เป็น .txt",
            data="\ufeff" + text,
            file_name=f"{stem}-ocr.txt",
            mime="text/plain;charset=utf-8",
            use_container_width=True,
        )
        st.markdown("### ดาวน์โหลดเอกสาร Word")
        try:
            editable_docx = export_pdf_editable_docx(uploaded_bytes)
            st.download_button(
                "⬇️ ดาวน์โหลด DOCX แก้ไขได้ (แนะนำ)",
                data=editable_docx,
                file_name=f"{stem}-editable.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                help=(
                    "ข้อความจะถูกสร้างเป็น Word text จริงและวางตามพิกัดจาก PDF/OCR "
                    "จึงแก้ไขได้ แต่เส้น ตาราง และรูปภาพบางส่วนอาจไม่เหมือนต้นฉบับทั้งหมด"
                ),
                use_container_width=True,
            )
            st.caption(
                "DOCX แก้ไขได้ใช้ข้อความจริงจาก text layer หรือ OCR bounding boxes; "
                "เหมาะสำหรับแก้เนื้อหาและคงตำแหน่งเดิมให้มากที่สุด"
            )
        except Exception as editable_error:
            st.warning(f"สร้าง DOCX แก้ไขได้ไม่สำเร็จ: {type(editable_error).__name__}: {editable_error}")

        try:
            visual_docx = export_pdf_visual_docx(uploaded_bytes)
            st.download_button(
                "⬇️ ดาวน์โหลด DOCX ภาพเหมือนต้นฉบับ",
                data=visual_docx,
                file_name=f"{stem}-visual.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                help=(
                    "แต่ละหน้า PDF จะถูกฝังเป็นภาพใน DOCX เพื่อรักษาตาราง รูปภาพ คอลัมน์ "
                    "และตำแหน่งเดิมให้มากที่สุด โหมดนี้ไม่เหมาะกับการแก้ไขข้อความ"
                ),
                use_container_width=True,
            )
            st.caption(
                "DOCX นี้เน้นรักษาหน้าตาเดิมของ PDF และใช้ภาพเต็มหน้า จึงไม่ใช่ข้อความ Word ที่แก้ไขได้"
            )
        except Exception as visual_error:
            st.warning(f"สร้าง DOCX ภาพเหมือนต้นฉบับไม่สำเร็จ: {type(visual_error).__name__}: {visual_error}")
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
        st.error(f"อ่านเอกสารไม่สำเร็จ: {type(exc).__name__}: {exc}")
