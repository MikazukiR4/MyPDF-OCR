"""
Integration and regression test suite for THSummarizer / MyPDF-OCR pipeline.
"""

from __future__ import annotations

import io
import unittest
import urllib.error

from core.exceptions import (
    PermanentOCRError,
    TransientOCRError,
    is_transient_error,
    with_api_retry,
)
from core.utils.cleaners import (
    clean_simple_text,
    clean_thai_ocr_text,
    normalize_thai_unicode,
    thai_text_quality_score,
)
from docx_generator import normalize_thai_text as docx_normalize
from docx_port import normalize_pdf_text as port_normalize
from tests.evaluate_cer import calculate_cer, calculate_wer, levenshtein_distance


class TestThaiCleaners(unittest.TestCase):
    def test_pua_mapping_tones(self):
        # Verify \uf70a (mai ek) and \uf70b (mai tho)
        raw = "ข\uf70bอความ"
        cleaned = normalize_thai_unicode(raw)
        self.assertEqual(cleaned, "ข้อความ")

    def test_sara_ae_normalization(self):
        # Two consecutive sara e must be converted to sara ae
        raw = "\u0e40\u0e40ละแล้ว"
        cleaned = normalize_thai_unicode(raw)
        self.assertEqual(cleaned, "และแล้ว")

    def test_sara_am_decomposition_fix(self):
        # Nikhahit + sara aa -> sara am
        raw = "น\u0e4d\u0e32ดื่ม"
        cleaned = normalize_thai_unicode(raw)
        self.assertEqual(cleaned, "นำดื่ม")

    def test_tone_mark_sara_am_order(self):
        # Thai unicode standard: consonant + tone mark + sara am
        # If input has sara am + tone mark (นำ้), it should become (น้ำ)
        raw = "นำ้ประปา"
        cleaned = normalize_thai_unicode(raw)
        self.assertEqual(cleaned, "น้ำประปา")

    def test_repeated_tone_mark_dedup(self):
        # Deduplicate repeated tone marks (OCR artifact: e.g. ที่่่่ -> ที่)
        raw = "ที่่่่สุด"
        cleaned = normalize_thai_unicode(raw)
        self.assertEqual(cleaned, "ที่สุด")

    def test_government_typo_replacements(self):
        raw = "เรอง ขออนุมัติงบประมาณ\nอางถง หนังสือด่วนที่สุด"
        cleaned = clean_thai_ocr_text(raw)
        lines = cleaned.splitlines()
        self.assertTrue(lines[0].startswith("เรื่อง"))
        self.assertTrue(lines[1].startswith("อ้างถึง"))

    def test_floating_line_noise_removal(self):
        raw = (
            "หนังสือราชการ สำคัญ\n"
            "= . - _ ~ ^ |\n"  # pure noise line
            "จังหวัดเชียงใหม่\n"
            "๑\n"  # page number - must NOT be filtered out!
        )
        cleaned = clean_thai_ocr_text(raw)
        self.assertIn("หนังสือราชการ", cleaned)
        self.assertIn("จังหวัดเชียงใหม่", cleaned)
        self.assertIn("๑", cleaned)
        self.assertNotIn("= . - _ ~ ^ |", cleaned)

    def test_quality_score(self):
        clean_text = "บันทึกข้อความ ส่วนราชการ จังหวัดเชียงใหม่ วันที่ ๑๕ มกราคม ๒๕๖๗"
        noisy_text = "\uf700\ufffd\x01\x02??!! ==== //\\"
        empty_text = ""

        self.assertGreater(thai_text_quality_score(clean_text), 0.80)
        self.assertLess(thai_text_quality_score(noisy_text), 0.35)
        self.assertEqual(thai_text_quality_score(empty_text), 0.0)


class TestDocxThaiNormalization(unittest.TestCase):
    def test_docx_generator_normalization(self):
        # Verify that docx_generator does not swap tone marks incorrectly
        raw = "น้ำใจ และ ต่ำกว่า"
        normalized = docx_normalize(raw)
        self.assertEqual(normalized, "น้ำใจ และ ต่ำกว่า")

    def test_docx_port_normalization(self):
        raw = "ต้นน้ำ และ ที่ทำงาน"
        normalized = port_normalize(raw)
        self.assertEqual(normalized, "ต้นน้ำ และ ที่ทำงาน")


class TestRetryAndExceptions(unittest.TestCase):
    def test_is_transient_error(self):
        self.assertTrue(is_transient_error(urllib.error.HTTPError("url", 429, "Rate limited", {}, None)))
        self.assertTrue(is_transient_error(urllib.error.HTTPError("url", 503, "Unavailable", {}, None)))
        self.assertTrue(is_transient_error(TimeoutError("Connection timed out")))
        self.assertFalse(is_transient_error(urllib.error.HTTPError("url", 404, "Not Found", {}, None)))
        self.assertFalse(is_transient_error(urllib.error.HTTPError("url", 401, "Unauthorized", {}, None)))

    def test_retry_success_after_transient_failures(self):
        attempts = 0

        @with_api_retry(max_attempts=3, min_wait=0.01, max_wait=0.05)
        def unstable_call():
            nonlocal attempts
            attempts += 1
            if attempts < 3:
                raise urllib.error.HTTPError("url", 429, "Rate limit", {}, None)
            return "SUCCESS"

        result = unstable_call()
        self.assertEqual(result, "SUCCESS")
        self.assertEqual(attempts, 3)


class TestEvaluationMetrics(unittest.TestCase):
    def test_levenshtein(self):
        self.assertEqual(levenshtein_distance("kitten", "sitting"), 3)
        self.assertEqual(levenshtein_distance("ภาษาไทย", "ภาษาไทย"), 0)

    def test_cer_calculation(self):
        ref = "ทดสอบการทำงาน"
        hyp = "ทดสอบการทำงาย"  # 1 substitution (น -> ย)
        cer = calculate_cer(ref, hyp)
        self.assertAlmostEqual(cer, 1 / len(ref), places=3)


class TestPageCacheAndExtractPdf(unittest.TestCase):
    def test_extract_pdf_embedded_and_page_cache(self):
        from unittest.mock import patch
        import pymupdf as fitz
        from app import extract_pdf

        doc = fitz.open()
        doc.new_page(width=595, height=842)
        pdf_bytes = doc.tobytes()
        doc.close()

        thai_text = "หนังสือราชการ เรื่อง ขออนุมัติโครงการ วันที่ ๑๕ มกราคม ๒๕๖๗"

        class DummyProgress:
            def progress(self, val, text=""):
                pass

        page_cache = {}
        with patch("app.embedded_pdf_text", return_value=thai_text):
            text, errors, failed = extract_pdf(
                pdf_bytes=pdf_bytes,
                language="tha",
                mode="fast",
                psm="3",
                oem="1",
                engine="tesseract",
                use_gemini_correction=False,
                gemini_model="gemini-3.8-flash",
                progress=DummyProgress(),
                page_cache=page_cache,
            )

        self.assertIn("ขออนุมัติโครงการ", text)
        self.assertEqual(len(errors), 0)
        self.assertIn(1, page_cache)
        cached_text, cached_method = page_cache[1]
        self.assertIn("ขออนุมัติโครงการ", cached_text)
        self.assertEqual(cached_method, "embedded text")


if __name__ == "__main__":
    unittest.main()

