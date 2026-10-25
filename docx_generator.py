"""docx_generator.py — Clean, high-quality, flowing Thai DOCX generator for MyPDF-OCR.

Designed specifically to eliminate Thai text scrambling, clipping of tone marks,
and overlapping textboxes by using native flowing Word paragraphs, tables,
and headings styled with Thai fonts (TH Sarabun New) and proper bidi/language tags.
"""
from __future__ import annotations

import io
import re
import unicodedata
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

DEFAULT_THAI_FONT = "TH Sarabun New"
DEFAULT_FONT_SIZE_PT = 16.0
DEFAULT_HEADING_1_SIZE_PT = 20.0
DEFAULT_HEADING_2_SIZE_PT = 18.0
DEFAULT_HEADING_3_SIZE_PT = 16.0

PAGE_MARKER_RE = re.compile(
    r"^-----\s*(?:หน้า|page)\s+([0-9]+)(?:\s*\|.*?)?\s*-----\s*$",
    re.IGNORECASE | re.MULTILINE,
)


from core.utils.cleaners import normalize_thai_unicode


def normalize_thai_text(text: str) -> str:
    """Fix Thai PDF glyph artifacts (PUA tone marks, split sara am, mark order)."""
    if not text:
        return ""
    text = normalize_thai_unicode(text)
    # Remove file content tags if any
    text = re.sub(r"<file_content\b[^>]*>.*?</file_content>", "", text, flags=re.DOTALL)
    text = re.sub(r"<file_content\b[^>]*>.*$", "", text, flags=re.DOTALL)
    return text.strip()


def _set_run_thai_font(
    run,
    font_name: str = DEFAULT_THAI_FONT,
    size_pt: float = DEFAULT_FONT_SIZE_PT,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
) -> None:
    """Configure Word run element with full Thai font, bidi, and language tags."""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color:
        clean_color = color.lstrip("#")
        if len(clean_color) == 6:
            run.font.color.rgb = RGBColor.from_string(clean_color)

    rpr = run._r.get_or_add_rPr()

    # Set rFonts for all script variants so Word uses TH Sarabun New for Thai complex script (cs)
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{attr}"), font_name)

    # Complex script font size
    sz_cs = rpr.find(qn("w:szCs"))
    if sz_cs is None:
        sz_cs = OxmlElement("w:szCs")
        rpr.append(sz_cs)
    sz_cs.set(qn("w:val"), str(int(size_pt * 2)))

    # Bold and Italic for Complex Script
    if bold:
        if rpr.find(qn("w:bCs")) is None:
            rpr.append(OxmlElement("w:bCs"))
    if italic:
        if rpr.find(qn("w:iCs")) is None:
            rpr.append(OxmlElement("w:iCs"))

    # Crucial: Thai language and bidi tags for correct word breaking
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "th-TH")
    lang.set(qn("w:bidi"), "th-TH")


def _set_cell_margins(cell, top: int = 100, bottom: int = 100, left: int = 150, right: int = 150) -> None:
    """Set padding inside a table cell (in dxa: 20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_cell_background(cell, hex_color: str) -> None:
    """Set cell background shading."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color.lstrip("#")}"/>')
    tcPr.append(shd)


def _set_table_borders(table, hex_color: str = "CCCCCC") -> None:
    """Set thin, clean borders for a Word table."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="{hex_color}"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{hex_color}"/>'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="{hex_color}"/>'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="{hex_color}"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{hex_color}"/>'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{hex_color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def _parse_markdown_table(lines: list[str]) -> list[list[str]]:
    """Parse consecutive table lines into a 2D array of cell contents."""
    rows = []
    for line in lines:
        cleaned = line.strip()
        if not cleaned:
            continue
        # Split by pipe
        parts = [p.strip() for p in cleaned.split("|")]
        # Drop leading/trailing empty cells if the line started/ended with |
        if cleaned.startswith("|") and len(parts) > 1:
            parts = parts[1:]
        if cleaned.endswith("|") and len(parts) > 0:
            parts = parts[:-1]
        # Ignore separator row like |---|---|
        if all(re.match(r"^:?-+:?$", p) for p in parts if p):
            continue
        if parts:
            rows.append(parts)
    return rows


def create_flowing_docx(
    ocr_text: str,
    font_name: str = DEFAULT_THAI_FONT,
    title: str = "MyOCR PDF Export",
) -> bytes:
    """Create a clean, flowing, editable Word document from OCR text.

    - Properly sets up Thai typography so tone marks are never clipped.
    - Converts page markers '----- หน้า N -----' into true Word page breaks.
    - Detects markdown tables and renders native Word tables.
    - Detects headings (#, ##) and bullet lists.
    """
    doc = Document()

    # Page Margins (1 inch = 72 pt = 2.54 cm)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.27)   # A4
        section.page_height = Inches(11.69)  # A4

    # Configure Normal Style with Thai Font
    normal_style = doc.styles["Normal"]
    normal_style.font.name = font_name
    normal_style.font.size = Pt(DEFAULT_FONT_SIZE_PT)
    normal_rpr = normal_style.element.get_or_add_rPr()
    rfonts = normal_rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        normal_rpr.append(rfonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{attr}"), font_name)

    # Split text into pages by marker
    pages_raw = PAGE_MARKER_RE.split(ocr_text)
    # Structure of split: [pre_text, page_1_num, page_1_text, page_2_num, page_2_text, ...]
    page_data: list[tuple[int | None, str]] = []

    if len(pages_raw) == 1:
        page_data.append((1, pages_raw[0]))
    else:
        # If there's text before the first marker
        if pages_raw[0].strip():
            page_data.append((None, pages_raw[0]))
        for i in range(1, len(pages_raw), 2):
            p_num = int(pages_raw[i]) if pages_raw[i].isdigit() else i // 2 + 1
            p_text = pages_raw[i + 1] if i + 1 < len(pages_raw) else ""
            page_data.append((p_num, p_text))

    for page_idx, (page_num, raw_content) in enumerate(page_data):
        if page_idx > 0:
            doc.add_page_break()

        lines = raw_content.splitlines()
        table_buffer: list[str] = []

        def flush_table():
            nonlocal table_buffer
            if not table_buffer:
                return
            grid = _parse_markdown_table(table_buffer)
            table_buffer = []
            if not grid:
                return

            max_cols = max(len(row) for row in grid)
            if max_cols == 0:
                return

            tbl = doc.add_table(rows=len(grid), cols=max_cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            _set_table_borders(tbl, "B0C4DE")

            for r_idx, row in enumerate(grid):
                is_header = (r_idx == 0)
                for c_idx in range(max_cols):
                    cell = tbl.cell(r_idx, c_idx)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    _set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
                    if is_header:
                        _set_cell_background(cell, "F0F4F8")

                    cell_text = row[c_idx] if c_idx < len(row) else ""
                    cell_text = normalize_thai_text(cell_text)

                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.15
                    if is_header:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if cell_text:
                        run = p.add_run(cell_text)
                        _set_run_thai_font(
                            run,
                            font_name=font_name,
                            size_pt=14.0 if is_header else 14.0,
                            bold=is_header,
                        )

            # Add small spacing after table
            spacer = doc.add_paragraph()
            spacer.paragraph_format.space_before = Pt(0)
            spacer.paragraph_format.space_after = Pt(4)

        for line in lines:
            line_str = line.strip()

            # Check if this line is part of a markdown table
            if line_str.startswith("|") and line_str.endswith("|") and line_str.count("|") >= 2:
                table_buffer.append(line_str)
                continue
            else:
                if table_buffer:
                    flush_table()

            if not line_str:
                # Empty line -> spacing
                continue

            normalized_line = normalize_thai_text(line)

            # Headings
            if normalized_line.startswith("### "):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(normalized_line[4:].strip())
                _set_run_thai_font(run, font_name, size_pt=DEFAULT_HEADING_3_SIZE_PT, bold=True)
            elif normalized_line.startswith("## "):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(normalized_line[3:].strip())
                _set_run_thai_font(run, font_name, size_pt=DEFAULT_HEADING_2_SIZE_PT, bold=True)
            elif normalized_line.startswith("# "):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(14)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(normalized_line[2:].strip())
                _set_run_thai_font(run, font_name, size_pt=DEFAULT_HEADING_1_SIZE_PT, bold=True)
            elif re.match(r"^[-*•]\s+", normalized_line):
                # Bullet list item
                bullet_content = re.sub(r"^[-*•]\s+", "", normalized_line)
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                run_bullet = p.add_run("• ")
                _set_run_thai_font(run_bullet, font_name, size_pt=DEFAULT_FONT_SIZE_PT, bold=True)
                run_text = p.add_run(bullet_content)
                _set_run_thai_font(run_text, font_name, size_pt=DEFAULT_FONT_SIZE_PT)
            elif re.match(r"^\d+\.\s+", normalized_line):
                # Numbered list item
                m = re.match(r"^(\d+\.)\s+(.*)$", normalized_line)
                prefix = m.group(1) if m else ""
                content = m.group(2) if m else normalized_line
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                run_num = p.add_run(f"{prefix} ")
                _set_run_thai_font(run_num, font_name, size_pt=DEFAULT_FONT_SIZE_PT, bold=True)
                run_text = p.add_run(content)
                _set_run_thai_font(run_text, font_name, size_pt=DEFAULT_FONT_SIZE_PT)
            else:
                # Normal paragraph
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(normalized_line)
                _set_run_thai_font(run, font_name, size_pt=DEFAULT_FONT_SIZE_PT)

        if table_buffer:
            flush_table()

    # Core properties
    doc.core_properties.title = title
    doc.core_properties.subject = "Clean Thai Flowing Document Export"
    doc.core_properties.comments = "Exported from MyPDF-OCR with clean Thai typography"

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()
