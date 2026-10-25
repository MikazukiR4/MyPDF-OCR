"""
OCR Thai Text Cleaning Pipeline
===============================
Standard, modular Thai OCR text cleaning and normalization utilities.
Includes PUA (Private Use Area) font mapping for Thai PDFs, Unicode normalization,
tone mark reordering and deduplication, floating-line filtering, and quality scoring.
"""

import re
import unicodedata
from typing import Callable, List, Optional, Tuple

# Mapping of Thai Private Use Area (PUA) glyphs commonly used in legacy Thai fonts (WinThai/MacThai/PSL/DB)
THAI_PUA_MAP = {
    # Cut-tail or special consonants
    "\uf700": "\u0e13",  # ณ (cut tail)
    # Upper vowels shifted/aligned
    "\uf701": "\u0e34",  # sara i shifted
    "\uf702": "\u0e35",  # sara ii shifted
    "\uf703": "\u0e36",  # sara ue shifted
    "\uf704": "\u0e37",  # sara uee shifted
    # Tone marks shifted down/left
    "\uf705": "\u0e48",  # mai ek
    "\uf706": "\u0e49",  # mai tho
    "\uf707": "\u0e4a",  # mai tri
    "\uf708": "\u0e4b",  # mai chattawa
    "\uf709": "\u0e4c",  # thanthakhat
    # Normal upper tone marks / marks
    "\uf70a": "\u0e48",  # mai ek
    "\uf70b": "\u0e49",  # mai tho
    "\uf70c": "\u0e4a",  # mai tri
    "\uf70d": "\u0e4b",  # mai chattawa
    "\uf70e": "\u0e4c",  # thanthakhat
    "\uf70f": "\u0e4d",  # nikhahit
    # Vowels
    "\uf710": "\u0e31",  # mai han-akat
    "\uf711": "\u0e34",  # sara i
    "\uf712": "\u0e35",  # sara ii
    "\uf713": "\u0e36",  # sara ue
    "\uf714": "\u0e37",  # sara uee
    # Tone marks combined above vowels
    "\uf715": "\u0e48",  # mai ek
    "\uf716": "\u0e49",  # mai tho
    "\uf717": "\u0e4a",  # mai tri
    "\uf718": "\u0e4b",  # mai chattawa
    "\uf719": "\u0e4c",  # thanthakhat
    "\uf71a": "\u0e4e",  # yamakkan
    # Special consonants without lower pedestal
    "\uf884": "\u0e10",  # ฐ (without foot/pedestal)
    "\uf885": "\u0e0d",  # ญ (without foot/pedestal)
}

# Fullwidth ASCII to standard ASCII
FULLWIDTH_ASCII_MAP = {
    "\uFF01": "!", "\uFF02": '"', "\uFF03": "#", "\uFF04": "$", "\uFF05": "%",
    "\uFF06": "&", "\uFF07": "'", "\uFF08": "(", "\uFF09": ")", "\uFF0A": "*",
    "\uFF0B": "+", "\uFF0C": ",", "\uFF0D": "-", "\uFF0E": ".", "\uFF0F": "/",
    "\uFF10": "0", "\uFF11": "1", "\uFF12": "2", "\uFF13": "3", "\uFF14": "4",
    "\uFF15": "5", "\uFF16": "6", "\uFF17": "7", "\uFF18": "8", "\uFF19": "9",
    "\uFF1A": ":", "\uFF1B": ";", "\uFF1C": "<", "\uFF1D": "=", "\uFF1E": ">",
    "\uFF1F": "?", "\uFF20": "@",
    "\uFF21": "A", "\uFF22": "B", "\uFF23": "C", "\uFF24": "D", "\uFF25": "E",
    "\uFF26": "F", "\uFF27": "G", "\uFF28": "H", "\uFF29": "I", "\uFF2A": "J",
    "\uFF2B": "K", "\uFF2C": "L", "\uFF2D": "M", "\uFF2E": "N", "\uFF2F": "O",
    "\uFF30": "P", "\uFF31": "Q", "\uFF32": "R", "\uFF33": "S", "\uFF34": "T",
    "\uFF35": "U", "\uFF36": "V", "\uFF37": "W", "\uFF38": "X", "\uFF39": "Y",
    "\uFF3A": "Z",
    "\uFF41": "a", "\uFF42": "b", "\uFF43": "c", "\uFF44": "d", "\uFF45": "e",
    "\uFF46": "f", "\uFF47": "g", "\uFF48": "h", "\uFF49": "i", "\uFF4A": "j",
    "\uFF4B": "k", "\uFF4C": "l", "\uFF4D": "m", "\uFF4E": "n", "\uFF4F": "o",
    "\uFF50": "p", "\uFF51": "q", "\uFF52": "r", "\uFF53": "s", "\uFF54": "t",
    "\uFF55": "u", "\uFF56": "v", "\uFF57": "w", "\uFF58": "x", "\uFF59": "y",
    "\uFF5A": "z",
    "\uFEFF": "",  # Zero-width no-break space / BOM
}

# Translation table combining PUA and Fullwidth ASCII
NORMALIZATION_TRANS = str.maketrans({**THAI_PUA_MAP, **FULLWIDTH_ASCII_MAP})

# Thai character sets
THAI_CONSONANTS = r"[ก-ฮ]"
THAI_VOWELS_UPPER = r"[\u0e31\u0e34-\u0e37]"
THAI_VOWELS_LOWER = r"[\u0e38\u0e39]"
THAI_TONE_MARKS = r"[\u0e48-\u0e4b]"
THAI_OTHER_MARKS = r"[\u0e4c-\u0e4e]"  # Thanthakhat, nikhahit, yamakkan


def normalize_thai_unicode(text: str) -> str:
    """Normalize Thai Unicode, map PUA glyphs, fix double vowels, and canonicalize order."""
    if not text:
        return ""

    # 1. PUA and Fullwidth conversion
    text = text.translate(NORMALIZATION_TRANS)

    # 2. Sara ae (แ) vs double sara e (เ + เ)
    text = text.replace("\u0e40\u0e40", "\u0e41")

    # 3. Sara am (ำ) vs nikhahit + sara aa (ํ + า)
    text = text.replace("\u0e4d\u0e32", "\u0e33")

    # 4. Canonical order of tone marks and vowels:
    # If tone mark is placed BEFORE vowel, reorder to VOWEL + TONE: e.g. ก้ิ -> กิ้
    text = re.sub(
        rf"({THAI_TONE_MARKS}|{THAI_OTHER_MARKS})({THAI_VOWELS_UPPER}|{THAI_VOWELS_LOWER})",
        r"\2\1",
        text,
    )

    # If tone mark is placed AFTER sara am (ำ + ้), reorder to TONE + SARA AM (้ + ำ): e.g. นำ้ -> น้ำ
    text = re.sub(
        rf"\u0e33({THAI_TONE_MARKS})",
        r"\g<1>" + "\u0e33",
        text,
    )

    # 5. Deduplicate repeated tone marks or diacritics on the same consonant (OCR artifacts: e.g. ก้้ -> ก้)
    text = re.sub(r"([\u0e31\u0e34-\u0e39\u0e47-\u0e4e])\1+", r"\1", text)

    # 6. Apply Unicode NFC normalization
    return unicodedata.normalize("NFC", text)


def remove_floating_ocr_lines(text: str) -> str:
    """Filter out isolated floating line noise produced by scanned margins/backgrounds."""
    if not text:
        return ""

    output: list[str] = []
    for line in text.splitlines():
        compact = re.sub(r"\s+", "", line)
        tokens = line.split()
        thai_letters = len(re.findall(r"[ก-ฮ]", compact))
        ascii_letters = len(re.findall(r"[A-Za-z]", compact))
        symbols = len(re.findall(r"[=/_~^|<>*+`\"'“”‘’.,:;!?-]", compact))
        short_tokens = sum(len(token) <= 2 for token in tokens)

        is_number_only = bool(re.fullmatch(r"[0-9๐-๙.,/-]+", compact or " "))
        is_floating_line = (
            len(tokens) >= 3
            and short_tokens / len(tokens) >= 0.75
            and thai_letters <= 3
        ) or (
            len(compact) <= 8
            and thai_letters <= 1
            and (ascii_letters or symbols)
            and not is_number_only
        )

        if compact and not is_floating_line:
            output.append(re.sub(r"[ \t]{3,}", "  ", line.rstrip()))

    return "\n".join(output)


# Common government document and official Thai OCR correction patterns
GOVERNMENT_THAI_REPLACEMENTS: Tuple[Tuple[str, str], ...] = (
    (r"(?m)^(\s*)เรอง(\s+)", r"\1เรื่อง\2"),
    (r"(?m)^(\s*)(?:อางถง|อางถึง|อ้างถง)(\s+)", r"\1อ้างถึง\2"),
    (r"(?m)^(\s*)(?:สิงที่ส่งมาด้วย|สิงทีส่งมาด้วย|สิ่งที่สงมาดวย)(\s+)", r"\1สิ่งที่ส่งมาด้วย\2"),
    (r"(?:หนงสอจงหวด|หนงสือจงหวด|หนงสอจังหวัด|หนังสือจงหวด)", "หนังสือจังหวัด"),
    (r"(\s)ท(\s+(?:ศธ|สส|มท|กค|นร|อว|พม|สธ|กษ|รง|ยธ|ทส|วธ)\s+[๐-๙0-9])", r"\1ที่\2"),
    (r"(?:ลงวันท|ลงวนท|ลงวนที่)(\s+)", r"ลงวันที่\1"),
    (r"(?<=ลงวันที่\s)mo(?=\s+[ก-๙])", "๓๐"),
    (r"(?:ประจำปงบประมาณ|ประจําปงบประมาณ)", "ประจำปีงบประมาณ"),
    (r"(?:we|wa|w\.ศ)", "พ.ศ"),
    (r"([งร])การขับเคลือน", r"\1การขับเคลื่อน"),
    (r"ในการนี(?=\s|$)", "ในการนี้"),
    (r"(?:ในส่วน(?:ทีเกียวข้อง|ที่เกียวข้อง|ที่เกี่ยวข้อง))", "ในส่วนที่เกี่ยวข้อง"),
    (r"(?:อำเกอ|อําเกอ)", "อำเภอ"),
)


def apply_thai_replacements(text: str) -> str:
    """Apply safe standard Thai regex corrections."""
    if not text:
        return ""
    for pattern, replacement in GOVERNMENT_THAI_REPLACEMENTS:
        text = re.sub(pattern, replacement, text)
    return text


def clean_thai_ocr_text(text: str) -> str:
    """
    Complete Thai OCR text cleaning pipeline:
    1. Unicode normalization (PUA, Sara ae, Sara am, canonical ordering, tone dedupe, NFC)
    2. Floating line noise removal
    3. Government document typo corrections
    4. Paragraph & line whitespace compaction (preserving paragraph breaks)
    """
    if not text:
        return ""

    text = normalize_thai_unicode(text)
    text = remove_floating_ocr_lines(text)
    text = apply_thai_replacements(text)

    # Preserve paragraph spacing and compact excessive inline spaces
    lines = text.splitlines(keepends=True)
    output: list[str] = []
    for line in lines:
        if line.strip():
            compacted = re.sub(r"[ \t]{3,}", "  ", line.rstrip())
            output.append(compacted)

    return "\n".join(output).strip()


def clean_simple_text(text: str) -> str:
    """Simple fallback cleaner for plain text."""
    if not text:
        return ""
    text = text.strip()
    text = re.sub(r"\[\d+\]", "", text)
    text = re.sub(r"[^ก-๙A-Za-z0-9\s,\.!?/\-\(\)]+", " ", text)
    return " ".join(text.split())[:100000]


def thai_text_quality_score(text: str) -> float:
    """
    Calculate text quality score (0.0 to 1.0) for deciding whether
    an embedded text layer is clean or needs image OCR fallback.
    
    A high score (>= 0.7) means the text layer is coherent and readable.
    """
    if not text or len(text.strip()) < 10:
        return 0.0

    raw = text.strip()
    total_len = len(raw)
    # Require at least some actual Thai or Latin letters
    letters = len(re.findall(r"[\u0e01-\u0e2eA-Za-z]", raw))
    if letters < 3:
        return 0.0

    letter_ratio = min(1.0, (letters / total_len) * 1.5)

    # 1. Ratio of valid printable Thai + Latin + digits + basic punctuation
    valid_chars = len(re.findall(r"[\u0e01-\u0e5bA-Za-z0-9\s\.,;:!?\(\)\[\]\"'/\-—–]", raw))
    char_validity_ratio = valid_chars / total_len

    # 2. Ratio of unresolved PUA or replacement chars
    garbage_chars = len(re.findall(r"[\ufffd\uf700-\uf8ff\u0000-\u0008\u000b\u000c\u000e-\u001f]", raw))
    garbage_ratio = garbage_chars / total_len

    # 3. Floating or orphaned vowels/tones (e.g. vowel or tone mark with no consonant preceding it)
    orphaned_marks = len(re.findall(r"(?<![ก-ฮ])[\u0e31\u0e34-\u0e39\u0e48-\u0e4e]", raw))
    orphaned_ratio = orphaned_marks / max(1, len(re.findall(r"[\u0e31\u0e34-\u0e39\u0e48-\u0e4e]", raw)))

    # Calculate overall score
    score = (
        (letter_ratio * 0.40)
        + (char_validity_ratio * 0.40)
        + ((1.0 - min(1.0, garbage_ratio * 4)) * 0.10)
        + ((1.0 - min(1.0, orphaned_ratio)) * 0.10)
    )
    return max(0.0, min(1.0, score))


def create_pipeline(processors: Optional[List[Callable[[str], str]]] = None, skip: Optional[List[str]] = None) -> Callable[[str], str]:
    """Create a custom composable text cleaning pipeline."""
    if skip is None:
        skip = []

    available = {
        "normalize_thai_unicode": normalize_thai_unicode,
        "remove_floating_ocr_lines": remove_floating_ocr_lines,
        "apply_thai_replacements": apply_thai_replacements,
    }

    if processors is None:
        active = [fn for name, fn in available.items() if name not in skip]
    else:
        active = [p for p in processors if p.__name__ not in skip]

    def pipeline(text: str) -> str:
        for p in active:
            text = p(text)
        return text

    return pipeline
