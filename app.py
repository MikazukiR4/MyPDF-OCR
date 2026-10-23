import io
import base64
import csv
import json
import os
import re
import shutil
import subprocess
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv


load_dotenv()
load_dotenv(".env.typhoon", override=False)

MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_PAGES = 100
TESSDATA_DIR = Path(__file__).resolve().parent / "tessdata"

# Tesseract OCR Engine Modes (equivalent to tesseract.js OEM constants).
OEM_TESSERACT_ONLY = "0"
OEM_LSTM_ONLY = "1"
OEM_TESSERACT_LSTM_COMBINED = "2"
OEM_DEFAULT = "3"

PSM_VALUES = ["6", "3", "4"]
PSM_LABELS = {
    "0": "OSD_ONLY",
    "1": "AUTO_OSD",
    "2": "AUTO_ONLY",
    "3": "AUTO — เอกสารทั่วไป",
    "4": "SINGLE_COLUMN",
    "5": "SINGLE_BLOCK_VERT_TEXT",
    "6": "SINGLE_BLOCK — ใบเสร็จ/ตาราง",
    "7": "SINGLE_LINE",
    "8": "SINGLE_WORD",
    "9": "CIRCLE_WORD",
    "10": "SINGLE_CHAR",
    "11": "SPARSE_TEXT",
    "12": "SPARSE_TEXT_OSD",
    "13": "RAW_LINE",
}

GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
TYPHOON_BASE_URL = os.getenv("TYPHOON_BASE_URL", "https://api.opentyphoon.ai/v1").rstrip("/")
TYPHOON_MODEL = os.getenv("TYPHOON_MODEL", "typhoon-ocr")


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


def gemini_transcribe(image_bytes: bytes) -> str:
    """Transcribe a page with Gemini Vision without storing the API key in code."""
    api_key = gemini_api_key()
    if not api_key:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GEMINI_API_KEY ในไฟล์ .env")
    payload = {
        "contents": [{
            "parts": [
                {"text": (
                    "ถอดข้อความภาษาไทยจากภาพเอกสารนี้ตามต้นฉบับ ห้ามสรุป "
                    "ห้ามเดาหรือแต่งเติม ถ้าอักขระใดอ่านไม่ชัดให้ใส่ [อ่านไม่ออก] "
                    "และถ้าไม่แน่ใจให้ใส่ [ไม่แน่ใจ] ตรงตำแหน่งนั้น "
                    "รักษาการขึ้นบรรทัดใหม่และตัวเลขให้มากที่สุด"
                )},
                {"inline_data": {"mime_type": "image/png", "data": base64.b64encode(image_bytes).decode("ascii")}},
            ],
        }],
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 4096},
    }
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={api_key}"
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Gemini Vision HTTP {exc.code}: {detail[:500]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"เชื่อมต่อ Gemini Vision ไม่สำเร็จ: {exc.reason}") from exc
    try:
        return "\n".join(
            part["text"]
            for part in result["candidates"][0]["content"]["parts"]
            if part.get("text")
        ).strip()
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError(f"Gemini Vision ไม่ส่งข้อความกลับมา: {result}") from exc


def mark_gemini_uncertainty(text: str) -> str:
    markers = ("[อ่านไม่ออก]", "[ไม่แน่ใจ]", "อ่านไม่ออก", "ไม่แน่ใจ", "ไม่สามารถอ่าน")
    if any(marker in text for marker in markers):
        return f"⚠️ [Gemini ไม่แน่ใจ] {text}"
    return text


def safe_gemini_transcribe(image_bytes: bytes) -> str:
    try:
        return gemini_transcribe(image_bytes)
    except Exception:
        try:
            return typhoon_transcribe(image_bytes)
        except Exception:
            return ""


def safe_gemini_find_handwriting_regions(image_bytes: bytes) -> list[tuple[float, float, float, float]]:
    try:
        return gemini_find_handwriting_regions(image_bytes)
    except Exception:
        return []


def gemini_find_handwriting_regions(image_bytes: bytes) -> list[tuple[float, float, float, float]]:
    """Return normalized [top, left, bottom, right] boxes for handwriting."""
    api_key = gemini_api_key()
    if not api_key:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GEMINI_API_KEY ในไฟล์ .env")
    prompt = (
        "ตรวจหาบริเวณที่เป็นลายมือไทยในภาพนี้เท่านั้น ไม่ต้องอ่านตัวพิมพ์ "
        "ตอบเป็น JSON array เท่านั้น รูปแบบ [{\"top\":0.1,\"left\":0.2,"
        "\"bottom\":0.4,\"right\":0.9}] โดยค่าทั้งหมดเป็นสัดส่วน 0 ถึง 1 "
        "ถ้าไม่พบลายมือให้ตอบ []"
    )
    payload = {"contents": [{"parts": [
        {"text": prompt},
        {"inline_data": {"mime_type": "image/png", "data": base64.b64encode(image_bytes).decode("ascii")}},
    ]}], "generationConfig": {"temperature": 0.0, "maxOutputTokens": 512}}
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={api_key}"
    request = urllib.request.Request(endpoint, data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            result = json.loads(response.read().decode("utf-8"))
        raw = result["candidates"][0]["content"]["parts"][0]["text"]
        raw = raw.replace("```json", "").replace("```", "").strip()
        boxes = json.loads(raw)
    except Exception as exc:
        raise RuntimeError(f"Gemini ตรวจหาลายมือไม่สำเร็จ: {exc}") from exc
    regions = []
    for box in boxes if isinstance(boxes, list) else []:
        try:
            top, left = max(0.0, float(box["top"])), max(0.0, float(box["left"]))
            bottom, right = min(1.0, float(box["bottom"])), min(1.0, float(box["right"]))
            if bottom > top and right > left:
                regions.append((top, left, bottom, right))
        except (KeyError, TypeError, ValueError):
            continue
    return regions


def detect_colored_ink_regions(image_bytes: bytes) -> list[tuple[float, float, float, float]]:
    """Find separated blue-ink bands before asking Gemini to read them."""
    from PIL import Image

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    pixels = image.load()
    row_hits = []
    for y in range(image.height):
        hits = 0
        for x in range(image.width):
            red, green, blue = pixels[x, y]
            if blue > 85 and blue - red > 25 and blue - green > 8:
                hits += 1
        if hits >= max(2, image.width // 700):
            row_hits.append(y)
    if not row_hits:
        return []

    bands = []
    start = previous = row_hits[0]
    for row in row_hits[1:]:
        if row - previous > max(12, image.height // 180):
            bands.append((start, previous))
            start = row
        previous = row
    bands.append((start, previous))

    regions = []
    pad_y = max(12, image.height // 120)
    pad_x = max(12, image.width // 120)
    for top, bottom in bands:
        xs = []
        for y in range(top, bottom + 1):
            for x in range(image.width):
                red, green, blue = pixels[x, y]
                if blue > 85 and blue - red > 25 and blue - green > 8:
                    xs.append(x)
        if not xs:
            continue
        left = max(0, min(xs) - pad_x)
        right = min(image.width, max(xs) + pad_x)
        top_px = max(0, top - pad_y)
        bottom_px = min(image.height, bottom + pad_y)
        if top_px > image.height * 0.96 or bottom_px < image.height * 0.04:
            continue
        regions.append((top_px / image.height, left / image.width, bottom_px / image.height, right / image.width))
    return regions


def merge_regions(*region_groups: list[tuple[float, float, float, float]]) -> list[tuple[float, float, float, float]]:
    """Merge overlapping detector results while keeping separate handwriting zones."""
    merged: list[tuple[float, float, float, float]] = []
    for group in region_groups:
        for candidate in group:
            top, left, bottom, right = candidate
            duplicate = False
            for old_top, old_left, old_bottom, old_right in merged:
                overlap_top = max(top, old_top)
                overlap_left = max(left, old_left)
                overlap_bottom = min(bottom, old_bottom)
                overlap_right = min(right, old_right)
                overlap = max(0.0, overlap_bottom - overlap_top) * max(0.0, overlap_right - overlap_left)
                candidate_area = max(0.001, (bottom - top) * (right - left))
                if overlap / candidate_area > 0.45:
                    duplicate = True
                    break
            if not duplicate:
                merged.append(candidate)
    return merged


def crop_regions(image_bytes: bytes, regions: list[tuple[float, float, float, float]]) -> list[bytes]:
    from PIL import Image

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    crops = []
    for top, left, bottom, right in regions:
        box = (round(left * image.width), round(top * image.height), round(right * image.width), round(bottom * image.height))
        crop = image.crop(box)
        output = io.BytesIO()
        crop.save(output, format="PNG")
        crops.append(output.getvalue())
    return crops


def replace_handwriting_regions(
    tsv_text: str,
    image_bytes: bytes,
    regions: list[tuple[float, float, float, float]],
    replacements: list[str],
) -> str:
    """Rebuild Tesseract lines while replacing tokens inside handwriting boxes."""
    from PIL import Image

    image = Image.open(io.BytesIO(image_bytes))
    rows = []
    for row in csv.DictReader(tsv_text.splitlines(), delimiter="\t"):
        if row.get("level") != "5" or not row.get("text", "").strip():
            continue
        try:
            rows.append({**row, "left": int(row["left"]), "top": int(row["top"]), "width": int(row["width"]), "height": int(row["height"])})
        except (KeyError, ValueError):
            continue
    groups = {}
    for row in rows:
        key = (row.get("block_num"), row.get("par_num"), row.get("line_num"))
        groups.setdefault(key, []).append(row)
    ordered_lines = sorted(groups.values(), key=lambda line: min(item["top"] for item in line))
    # A detected region is not enough to replace text.  Gemini may fail, return
    # an empty response, or return only an uncertainty marker.  In all of
    # those cases the original Tesseract tokens must remain untouched.
    usable = {
        index: replacement.strip()
        for index, replacement in enumerate(replacements)
        if replacement and replacement.strip()
    }
    used = set()
    output_lines = []
    for line in ordered_lines:
        line.sort(key=lambda item: item["left"])
        parts = []
        for token in line:
            token_box = (
                token["top"] / image.height,
                token["left"] / image.width,
                (token["top"] + token["height"]) / image.height,
                (token["left"] + token["width"]) / image.width,
            )
            matching = []
            for index, region in enumerate(regions):
                overlap = max(0.0, min(token_box[2], region[2]) - max(token_box[0], region[0])) * max(0.0, min(token_box[3], region[3]) - max(token_box[1], region[1]))
                if overlap > 0:
                    matching.append(index)
            matching_usable = [index for index in matching if index in usable]
            if matching_usable:
                for index in matching_usable:
                    if index not in used:
                        parts.append(usable[index])
                        used.add(index)
                # Only suppress the Tesseract token when a replacement was
                # actually obtained.  Failed Gemini crops never delete text.
                continue
            parts.append(token["text"])
        if parts:
            output_lines.append(" ".join(parts))

    for index, replacement in usable.items():
        if index not in used:
            output_lines.append(replacement)
    return "\n".join(output_lines).strip()


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

    def recognize_tsv(self, image_bytes: bytes) -> str:
        if self.closed:
            raise RuntimeError("OCR worker ถูกปิดไปแล้ว")
        prepared = preprocess_image(image_bytes, self.mode)
        with tempfile.TemporaryDirectory(prefix="th_ocr_tsv_") as temp_dir:
            image_path = Path(temp_dir) / "page.png"
            image_path.write_bytes(prepared)
            args = [self.command, str(image_path), "stdout", "-l", self.language, "--psm", self.psm, "--oem", self.oem, "tsv"]
            if (TESSDATA_DIR / "tha.traineddata").exists() and (TESSDATA_DIR / "eng.traineddata").exists():
                args.extend(["--tessdata-dir", str(TESSDATA_DIR)])
            result = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120, check=False)
        if result.returncode != 0:
            error = result.stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"Tesseract TSV ทำงานไม่สำเร็จ: {error}")
        return result.stdout.decode("utf-8", errors="replace")

    def terminate(self) -> None:
        self.closed = True


def extract_pdf(
    pdf_bytes: bytes,
    language: str,
    mode: str,
    psm: str,
    oem: str,
    use_gemini: bool,
    auto_crop_handwriting: bool,
    use_typhoon: bool,
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
        worker = None if use_typhoon else TesseractWorker(language, mode, psm, oem)
        for page_number in range(1, page_count + 1):
            ratio = (page_number - 1) / page_count
            progress.progress(ratio, text=f"กำลังอ่านหน้า {page_number}/{page_count} ({round(ratio * 100)}%)")
            image_bytes = render_page(pdf_bytes, page_number)
            try:
                base_text = typhoon_transcribe(image_bytes) if use_typhoon else worker.recognize(image_bytes)
                if (auto_crop_handwriting or use_gemini) and not use_typhoon:
                    colored_regions = detect_colored_ink_regions(image_bytes)
                    semantic_regions = safe_gemini_find_handwriting_regions(image_bytes)
                    regions = merge_regions(colored_regions, semantic_regions)
                    crops = crop_regions(image_bytes, regions)
                    replacements = [mark_gemini_uncertainty(safe_gemini_transcribe(crop)) for crop in crops]
                    text = replace_handwriting_regions(
                        worker.recognize_tsv(image_bytes), image_bytes, regions, replacements
                    )
                    usable_replacements = [replacement for replacement in replacements if replacement.strip()]
                    # Keep the complete Tesseract page when Gemini could not
                    # produce a usable crop result.  Gemini must never replace
                    # a readable page with an empty or partial response.
                    text = text if usable_replacements else base_text
                    if usable_replacements:
                        text += (
                            f"\n\n[Gemini: แทนที่เฉพาะบริเวณลายมือ "
                            f"{len(usable_replacements)} จุด; กรุณาตรวจสอบตัวเลข/คำที่มี ⚠️]"
                        )
                    elif regions:
                        text = (
                            base_text
                            + "\n\n[Gemini: พบบริเวณที่อาจเป็นลายมือ แต่ยังอ่านไม่สำเร็จ "
                            "จึงคงผล Tesseract เดิมไว้]"
                        )
                    else:
                        text = base_text + "\n\n[Gemini: ไม่พบบริเวณลายมือ จึงใช้ Tesseract ทั้งหน้า]"
                    if not text:
                        text = "[ไม่พบบริเวณลายมือในหน้านี้]"
                else:
                    text = base_text
                output.append(f"----- หน้า {page_number} -----\n{text}")
            except Exception as exc:
                errors.append((page_number, f"{type(exc).__name__}: {exc}"))
                failed_images[page_number] = image_bytes
                output.append(f"----- หน้า {page_number} -----\n[อ่านหน้านี้ไม่สำเร็จ]")
            ratio = page_number / page_count
            progress.progress(ratio, text=f"อ่านหน้า {page_number}/{page_count} เสร็จแล้ว ({round(ratio * 100)}%)")
        if worker is not None:
            worker.terminate()
        return "\n\n".join(output), errors, failed_images
    finally:
        document.close()


def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root { --ink:#0f172a; --muted:#475569; --line:#cbd5e1; --accent:#dc2626; --accent-dark:#991b1b; --navy:#172554; }
        html, body, button, input, textarea, select { font-family:"Segoe UI", "Noto Sans Thai", Tahoma, sans-serif !important; }
        .stApp, [data-testid="stAppViewContainer"] { background: #f8fafc !important; color: var(--ink) !important; }
        header[data-testid="stHeader"] { background: #ffffff !important; }
        .block-container { max-width: 900px; padding-top: 2rem; }
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
        div[data-testid="stProgressBar"] > div > div { background:#e5322d !important; }
        div[data-testid="stButton"] > button { background:#1e293b !important; color:#ffffff !important; border:1px solid #0f172a !important; border-radius:8px !important; font-weight:700 !important; }
        div[data-testid="stButton"] > button:hover { background:#334155 !important; color:#ffffff !important; border-color:#334155 !important; }
        div[data-testid="stButton"] > button[kind="primary"], button[kind="primary"] { background:var(--accent) !important; color:#ffffff !important; border-color:var(--accent-dark) !important; box-shadow:0 4px 10px rgba(220,38,38,.18); }
        div[data-testid="stButton"] > button[kind="primary"]:hover, button[kind="primary"]:hover { background:var(--accent-dark) !important; color:#ffffff !important; }
        div[data-testid="stFileUploader"] button { background:#334155 !important; color:#ffffff !important; border:1px solid #1e293b !important; border-radius:7px !important; }
        div[data-testid="stFileUploader"] button:hover { background:#475569 !important; color:#ffffff !important; }
        div[data-testid="stDownloadButton"] > button { background:#0f766e !important; color:#ffffff !important; border-color:#115e59 !important; font-weight:700 !important; }
        div[data-testid="stDownloadButton"] > button:hover { background:#115e59 !important; color:#ffffff !important; }
        div[data-testid="stTextArea"] label, div[data-testid="stTextArea"] label p { color:#172033 !important; font-weight:700 !important; }
        div[data-testid="stTextArea"] textarea { background:#ffffff !important; color:#172033 !important; border:2px solid #cbd5e1 !important; border-radius:10px !important; font-size:16px !important; line-height:1.65 !important; opacity:1 !important; -webkit-text-fill-color:#172033 !important; }
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
        div[data-testid="stTextArea"] textarea { background:#020617 !important; color:#f8fafc !important; border-color:#475569 !important; -webkit-text-fill-color:#f8fafc !important; }
        div[data-testid="stTextArea"] label, div[data-testid="stTextArea"] label p { color:#f8fafc !important; }
        div[data-testid="stAlert"] { background:#172033 !important; color:#f8fafc !important; border-color:#475569 !important; }
        div[data-testid="stAlert"] p { color:#f8fafc !important; }
        div[data-testid="stProgressBar"] > div { background:#1e293b !important; }
        [data-baseweb="select"], [data-baseweb="select"] > div { background:#0f172a !important; border-color:#475569 !important; color:#f8fafc !important; }
        [data-baseweb="select"] * { color:#f8fafc !important; -webkit-text-fill-color:#f8fafc !important; }
        [role="listbox"], [role="option"] { background:#111827 !important; color:#f8fafc !important; }
        div[data-testid="stDownloadButton"] > button { background:#0f766e !important; color:#ffffff !important; }
        aside[data-testid="stSidebar"] { background:#111827 !important; }
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
    language = st.selectbox(
        "ภาษาในเอกสาร",
        ["tha", "tha+eng", "eng"],
        format_func=lambda value: {"tha+eng": "ไทย + อังกฤษ", "tha": "ไทย", "eng": "อังกฤษ"}[value],
    )
    mode = st.selectbox(
        "ความแม่นยำ",
        ["best", "fast"],
        format_func=lambda value: "แม่นยำสูง (ช้ากว่าเล็กน้อย)" if value == "best" else "เร็ว (แม่นยำน้อยกว่า)",
    )
    psm = st.selectbox(
        "รูปแบบเอกสาร",
        PSM_VALUES,
        format_func=lambda value: {"3": "เอกสารทั่วไป", "6": "ใบเสร็จ / ตาราง", "4": "คอลัมน์เดียว"}[value],
    )
    oem = st.selectbox(
        "โหมด OCR (OEM)",
        [OEM_LSTM_ONLY, OEM_DEFAULT, OEM_TESSERACT_LSTM_COMBINED, OEM_TESSERACT_ONLY],
        format_func=lambda value: {
            OEM_LSTM_ONLY: "LSTM_ONLY — แนะนำสำหรับภาษาไทย",
            OEM_DEFAULT: "DEFAULT",
            OEM_TESSERACT_LSTM_COMBINED: "TESSERACT_LSTM_COMBINED",
            OEM_TESSERACT_ONLY: "TESSERACT_ONLY — รุ่นเก่า",
        }[value],
    )
    use_gemini = st.checkbox(
        "ใช้ Gemini Vision อ่านลายมือไทย (แม่นขึ้น แต่อาจช้ากว่าและใช้โควตา)",
        value=False,
    )
    use_typhoon = st.checkbox(
        "ใช้ Typhoon OCR อ่านเอกสารราชการไทยทั้งหน้า (แนะนำ)",
        value=bool(typhoon_api_key()),
    )
    auto_crop_handwriting = st.checkbox(
        "ให้ Gemini หาและ crop เฉพาะบริเวณลายมืออัตโนมัติ",
        value=False,
        disabled=not use_gemini or use_typhoon,
    )
    if use_typhoon and not typhoon_api_key():
        st.warning("ยังไม่ได้ตั้งค่า TYPHOON_OCR_API_KEY ในไฟล์ .env หรือ .env.typhoon")
    elif use_typhoon:
        st.caption("Typhoon OCR จะส่งภาพแต่ละหน้าไปยังบริการภายนอกเพื่ออ่านข้อความ")
    if use_gemini and not gemini_api_key():
        st.warning("ยังไม่ได้ตั้งค่า GEMINI_API_KEY ในไฟล์ .env")
    if tesseract_path():
        st.success("พบ Tesseract แล้ว")
    else:
        st.warning("ยังไม่พบ Tesseract")
    st.caption("รองรับ PDF สูงสุด 50 MB / 100 หน้า")

st.markdown('<div class="ocr-card">', unsafe_allow_html=True)
uploaded = st.file_uploader("ลากไฟล์มาวางที่นี่ หรือเลือกไฟล์ PDF", type=["pdf"])
if uploaded is not None:
    st.caption(f"ไฟล์ที่เลือก: {uploaded.name} ({uploaded.size / 1024:.1f} KB)")
run = st.button("🔍 เริ่ม OCR", type="primary", use_container_width=True, disabled=uploaded is None)
st.markdown('</div>', unsafe_allow_html=True)

if run and uploaded is not None:
    progress = st.progress(0, text="กำลังเตรียมไฟล์... (0%)")
    try:
        text, errors, failed_images = extract_pdf(
            uploaded.getvalue(), language, mode, psm, oem, use_gemini,
            auto_crop_handwriting, use_typhoon, progress,
        )
        progress.progress(1.0, text="OCR เสร็จสมบูรณ์ (100%)")
        st.success("อ่านเอกสารสำเร็จ")
        st.text_area("ข้อความที่อ่านได้", text, height=520)
        if "⚠️ [Gemini ไม่แน่ใจ]" in text:
            st.warning("ผลลัพธ์มีส่วนที่ Gemini ระบุว่าไม่แน่ใจ กรุณาตรวจสอบกับภาพต้นฉบับ")
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
