import io
import base64
import json
import os
import re
import shutil
import subprocess
import tempfile
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
TESSDATA_DIR = Path(__file__).resolve().parent / "tessdata"

OEM_LSTM_ONLY = "1"

PSM_VALUES = ["6", "3", "4"]

TYPHOON_BASE_URL = os.getenv("TYPHOON_BASE_URL", "https://api.opentyphoon.ai/v1").rstrip("/")
TYPHOON_MODEL = os.getenv("TYPHOON_MODEL", "typhoon-ocr")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
PDF_VIEWER_HTML = Path(__file__).resolve().parent / "pdf-viewer.html"
PDF_VIEWER_APP_JS = Path(__file__).resolve().parent / "pdf-viewer-app.js"


def render_page(pdf_bytes: bytes, page_number: int) -> bytes:
    """Render one PDF page to a high-DPI PNG (~300 DPI) so Thai diacritics stay crisp."""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = document[page_number - 1]
        width = max(1, page.rect.width)
        scale = min(4.2, max(2.5, 2400 / width))
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        return pixmap.tobytes("png")
    finally:
        document.close()


def preprocess_image(image_bytes: bytes, mode: str) -> bytes:
    """Convert to clean grayscale, upscale only if small, and suppress dotted-line/scan noise."""
    from PIL import Image, ImageFilter, ImageOps

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    if image.width < 2000:
        scale = min(2.5, 2200 / max(1, image.width))
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
    text = re.sub(r"\u0e4d\s*\u0e32", "\u0e33", text)
    text = re.sub(r"([\u0e48-\u0e4b])\u0e33", "ำ\\1", text)

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
    return raw.strip()


def typhoon_api_key() -> str:
    return os.getenv("TYPHOON_OCR_API_KEY") or os.getenv("TYPHOON_API_KEY") or ""


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
    """Show the model's quick comparison scale directly in the dropdown options."""
    profile = gemini_model_profile(model_name)
    return (
        f"{model_name} | "
        f"คุณภาพ {star_scale(profile['quality'])} · "
        f"เร็ว {star_scale(profile['speed'])} · "
        f"ภาพ {star_scale(profile['vision'])} · "
        f"ประหยัด {star_scale(profile['economy'])}"
    )


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
) -> tuple[str, list[tuple[int, str]], dict[int, bytes]]:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

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
        errors: list[tuple[int, str]] = []
        failed_images: dict[int, bytes] = {}
        worker = None if engine == "typhoon" else TesseractWorker(language, mode, psm, oem)
        try:
            for page_number in range(1, page_count + 1):
                ratio = (page_number - 1) / page_count
                progress.progress(ratio, text=f"กำลังอ่านหน้า {page_number}/{page_count}")
                page = document[page_number - 1]
                embedded = embedded_pdf_text(page)
                # Born-digital PDFs already contain selectable text. Prefer it
                # over image OCR to avoid corrupting Thai text and wasting API calls.
                if len(embedded) >= 80:
                    text = embedded
                    method = "embedded PDF text"
                    if use_gemini_correction:
                        try:
                            corrected = gemini_correct_text(render_page(pdf_bytes, page_number), text, gemini_model)
                            text = corrected
                            method = "Gemini corrected"
                        except Exception as exc:
                            errors.append((page_number, f"Gemini ตรวจแก้ไม่สำเร็จ: {exc}"))
                    output.append(f"----- หน้า {page_number} | {method} -----\n{text}")
                    ratio = page_number / page_count
                    progress.progress(ratio, text=f"อ่านข้อความฝังหน้า {page_number}/{page_count} เสร็จแล้ว")
                    continue
                image_bytes = render_page(pdf_bytes, page_number)
                try:
                    text = typhoon_transcribe(image_bytes) if engine == "typhoon" else worker.recognize(image_bytes)
                    text = text or embedded_pdf_text(document[page_number - 1]) or "[ไม่พบข้อความ]"
                    method = engine
                    if use_gemini_correction and text and not text.startswith("["):
                        try:
                            text = gemini_correct_text(image_bytes, text, gemini_model)
                            method = "Gemini corrected"
                        except Exception as correction_error:
                            errors.append((page_number, f"Gemini ตรวจแก้ไม่สำเร็จ: {correction_error}"))
                    output.append(f"----- หน้า {page_number} | {method} -----\n{text}")
                except Exception as exc:
                    fallback = embedded_pdf_text(document[page_number - 1])
                    if fallback:
                        output.append(
                            f"----- หน้า {page_number} | embedded PDF text -----\n{fallback}"
                        )
                    else:
                        errors.append((page_number, f"{type(exc).__name__}: {exc}"))
                        failed_images[page_number] = image_bytes
                        output.append(f"----- หน้า {page_number} -----\n[อ่านหน้านี้ไม่สำเร็จ]")
                ratio = page_number / page_count
                progress.progress(ratio, text=f"อ่านหน้า {page_number}/{page_count} เสร็จแล้ว")
        finally:
            if worker is not None:
                worker.terminate()
        return "\n\n".join(output), errors, failed_images
    finally:
        document.close()


def render_pdf_viewer(pdf_bytes: bytes, filename: str, height: int = 720) -> None:
    """Embed the local single-file pdf-viewer.us viewer with the selected PDF."""
    if not PDF_VIEWER_HTML.exists() or not PDF_VIEWER_APP_JS.exists():
        st.error("ไม่พบไฟล์ PDF viewer ในโปรเจกต์")
        return
    viewer = PDF_VIEWER_HTML.read_text(encoding="utf-8")
    viewer_js = PDF_VIEWER_APP_JS.read_text(encoding="utf-8")
    viewer = re.sub(r'<script[^>]+plausible[^>]+></script>', '', viewer, flags=re.I)
    viewer = viewer.replace(
        '<script src="/app.js"></script>',
        f"<script>{viewer_js}</script>",
    )
    viewer = viewer.replace(
        "</body>",
        """
        <script>
        (() => {
          const payload = window.__THSUMMARIZER_PDF__;
          if (!payload) return;
          const bytes = Uint8Array.from(atob(payload.data), c => c.charCodeAt(0));
          const file = new File([bytes], payload.name, {type: "application/pdf"});
          const input = document.getElementById("fileInput");
          const load = () => {
            if (typeof window.__pdfViewerReadFile === "function") {
              window.__pdfViewerReadFile(file);
            } else if (input) {
              const transfer = new DataTransfer();
              transfer.items.add(file);
              input.files = transfer.files;
              input.dispatchEvent(new Event("change", {bubbles: true}));
            }
          };
          if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", load, {once: true});
          else setTimeout(load, 0);
        })();
        </script>
        </body>""",
    )
    encoded = base64.b64encode(pdf_bytes).decode("ascii")
    config = json.dumps({"name": filename, "data": encoded}).replace("<", "\\u003c")
    viewer = viewer.replace("<head>", f'<head><script>window.__THSUMMARIZER_PDF__={config};</script>', 1)
    components.html(viewer, height=height, scrolling=True)

def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root { --ink:#0f172a; --muted:#475569; --line:#cbd5e1; --accent:#dc2626; --accent-dark:#991b1b; --navy:#172554; }
        html, body, button, input, textarea, select { font-family:"Segoe UI", "Noto Sans Thai", Tahoma, sans-serif !important; }
        .stApp, [data-testid="stAppViewContainer"] { background: #f8fafc !important; color: var(--ink) !important; }
        header[data-testid="stHeader"], [data-testid="stHeader"] { background: #0b1220 !important; height: 2.5rem !important; }
        [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"] { background: transparent !important; }
        .block-container { max-width: 900px; padding-top: 3.5rem; }
        .mypdf-header { display:flex; align-items:center; gap:12px; margin-bottom:2rem; }
        .mypdf-icon { background:var(--accent); color:#fff; border-radius:9px; padding:9px 12px; font-size:24px; box-shadow:0 4px 10px rgba(220,38,38,.2); }
        .mypdf-name { font-size:30px; font-weight:800; color:var(--ink); letter-spacing:-.3px; }
        .mypdf-name span { color:var(--accent); }
        .ocr-card { background:#fff; border:1px solid #e5e7eb; border-radius:16px; padding:24px; box-shadow:0 5px 20px rgba(15,23,42,.06); }
        .subtitle { color:var(--muted) !important; margin-top:-12px; margin-bottom:24px; font-size:15px; }
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
        [role="listbox"], [role="option"] { background:#111827 !important; color:#f8fafc !important; }
        div[data-testid="stDownloadButton"] > button { background:#0f766e !important; color:#ffffff !important; border:1px solid #115e59 !important; min-height:44px !important; }
        aside[data-testid="stSidebar"] { background:#111827 !important; }
        .gemini-rating-row { display:flex; align-items:center; gap:8px; padding:4px 6px; border-radius:5px; cursor:help; }
        .gemini-rating-row:hover { background:#1e293b; }
        .gemini-rating-label { min-width:145px; color:#e5e7eb; font-size:0.88rem; }
        .gemini-rating-stars { color:#fbbf24; letter-spacing:1px; font-size:1rem; }
        .gemini-rating-help { color:#94a3b8; font-size:0.78rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title="MyPDF OCR — Thai", page_icon="📄", layout="centered")
inject_css()

st.markdown(
    '<div class="mypdf-header"><div class="mypdf-icon">PDF</div>'
    '<div class="mypdf-name"><span>My</span>PDF OCR</div></div>',
    unsafe_allow_html=True,
)
st.markdown('<div class="subtitle">อ่านข้อความจาก PDF สแกน — ภาษาไทยและภาษาอังกฤษ</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("ตั้งค่า OCR")
    engine = st.radio(
        "เครื่องมืออ่านเอกสาร",
        ["typhoon", "tesseract"],
        format_func=lambda value: {
            "typhoon": "Typhoon OCR — แนะนำสำหรับเอกสารไทย",
            "tesseract": "Tesseract — อ่านในเครื่อง / Offline",
        }[value],
    )
    language = st.selectbox(
        "ภาษา",
        ["tha", "tha+eng", "eng"],
        format_func=lambda value: {"tha": "ไทย", "tha+eng": "ไทย + อังกฤษ", "eng": "อังกฤษ"}[value],
    )
    mode = st.selectbox(
        "คุณภาพภาพ",
        ["best", "fast"],
        format_func=lambda value: "คมชัดสูง" if value == "best" else "เร็ว",
    )
    psm = st.selectbox(
        "รูปแบบเอกสาร",
        PSM_VALUES,
        index=0,
        format_func=lambda value: {"3": "เอกสารทั่วไป", "6": "บล็อกข้อความ / ตาราง", "4": "คอลัมน์เดียว"}[value],
    )
    oem = OEM_LSTM_ONLY
    use_gemini_correction = st.checkbox(
        "ให้ Gemini ตรวจแก้สระ/วรรณยุกต์จากภาพต้นฉบับ",
        value=False,
        help="ใช้ข้อความเดิมร่วมกับภาพ PDF เพื่อแก้ปัญหา เช่น กำร/นำยกรัฐมนตรี/ว่ำ โดยไม่สรุปหรือเพิ่มเนื้อหา",
    )
    gemini_model = GEMINI_MODEL
    if use_gemini_correction and gemini_api_key():
        try:
            available_gemini_models = list_gemini_models()
            if available_gemini_models:
                model_names = [str(model["name"]) for model in available_gemini_models]
                default_index = (
                    model_names.index(GEMINI_MODEL)
                    if GEMINI_MODEL in model_names else 0
                )
                gemini_model = st.selectbox(
                    "Gemini รุ่นที่ใช้ตรวจแก้",
                    model_names,
                    index=default_index,
                    format_func=gemini_model_option_label,
                    help="เปิดรายการเพื่อเปรียบเทียบดาวของแต่ละรุ่นได้ทันที: คุณภาพ · ความเร็ว · ภาพ · ความประหยัด",
                )
                selected_metadata = next(model for model in available_gemini_models if model["name"] == gemini_model)
                st.markdown("**รายละเอียดโมเดลที่เลือก**")
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
                st.caption("ภาพ: ระบบจะส่งภาพหน้า PDF ไปทดสอบกับโมเดลที่เลือกเมื่อเริ่มตรวจแก้")
                st.caption(f"ตรวจพบ {len(available_gemini_models)} โมเดลที่ใช้ generateContent ได้")
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
            st.error("ไม่พบ TYPHOON_OCR_API_KEY")
    elif tesseract_path():
        st.success("พร้อมใช้ Tesseract ในเครื่อง")
    else:
        st.error("ไม่พบ Tesseract")
    if use_gemini_correction:
        if gemini_api_key():
            st.warning("Gemini จะถูกเรียกทีละหน้าและอาจใช้โควตา API")
        else:
            st.error("ไม่พบ GEMINI_API_KEY")
    st.caption("รองรับ PDF สูงสุด 50 MB / 100 หน้า")

st.markdown('<div class="ocr-card">', unsafe_allow_html=True)
uploaded = st.file_uploader(
    "เลือกไฟล์ PDF เพื่อเริ่มต้น",
    type=["pdf"],
    accept_multiple_files=False,
    help="รองรับเฉพาะไฟล์ .pdf เท่านั้น หากมีไฟล์ Word ให้บันทึกเป็น PDF ก่อน",
)
if uploaded is not None:
    st.success(f"เลือกไฟล์แล้ว: {uploaded.name} · {uploaded.size / 1024:.1f} KB")
else:
    st.info("ยังไม่ได้เลือกไฟล์ PDF — ปุ่มด้านล่างจะแจ้งเตือนให้เลือกไฟล์")
run = st.button("🔍 เริ่ม OCR", type="primary", use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

if run:
    if uploaded is None:
        st.warning("กรุณาเลือกไฟล์ PDF ก่อนเริ่ม OCR")
        st.stop()
    progress = st.progress(0, text="กำลังเตรียมไฟล์... (0%)")
    try:
        text, errors, failed_images = extract_pdf(
            uploaded.getvalue(), language, mode, psm, oem, engine,
            use_gemini_correction, gemini_model, progress,
        )
        progress.progress(1.0, text="OCR เสร็จสมบูรณ์ (100%)")
        st.success("อ่านเอกสารสำเร็จ")
        st.session_state["ocr_text"] = text
        stem = Path(uploaded.name).stem
        st.download_button(
            "⬇️ ดาวน์โหลดข้อความเป็น .txt",
            data="\ufeff" + text,
            file_name=f"{stem}-ocr.txt",
            mime="text/plain;charset=utf-8",
            use_container_width=True,
        )
        if errors:
            with st.expander(f"หน้าที่อ่านไม่สำเร็จ ({len(errors)})"):
                for page_number, error in errors:
                    st.error(f"หน้า {page_number}: {error}")
                    st.download_button(
                        f"ดาวน์โหลดภาพหน้า {page_number}",
                        data=failed_images[page_number],
                        file_name=f"{stem}-page-{page_number}-failed.png",
                        mime="image/png",
                        key=f"failed-{page_number}",
                    )
    except Exception as exc:
        progress.progress(1.0, text="OCR ไม่สำเร็จ (100%)")
        st.error(f"OCR ไม่สำเร็จ: {type(exc).__name__}: {exc}")

if uploaded is not None:
    st.markdown("### ตรวจเทียบ PDF กับข้อความ OCR")
    viewer_col, text_col = st.columns([1.15, 1], gap="medium")
    with viewer_col:
        st.caption("PDF ต้นฉบับ")
        render_pdf_viewer(uploaded.getvalue(), uploaded.name, height=720)
    with text_col:
        st.caption("ข้อความที่ถอดได้")
        st.text_area(
            "ผล OCR",
            st.session_state.get("ocr_text", "กด ‘เริ่ม OCR’ เพื่อแสดงข้อความที่ถอดได้"),
            height=680,
            label_visibility="collapsed",
        )
