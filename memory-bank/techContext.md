# Tech Context — MyOCR PDF

## Technologies
- **Framework:** Streamlit 1.66.0+
- **Language:** Python 3.14.8
- **OCR Engines:** Tesseract (local), Typhoon OCR (API)
- **AI Correction:** Google Gemini (gemini-3.8-flash by default)
- **PDF Library:** PyMuPDF (fitz) 2.0
- **Document Export:** python-docx
- **Language Detection:** langdetect 1.0.9

## Environment
- Development: Windows (PowerShell)
- Deployment: Render (render.yaml)
- Dependencies: requirements.txt, .env.example

## Key Files
- `app.py` — Main application (~2,026 lines after Phase 3)
  - Lines 50-67: Batch processing, theme toggle, language detection session state
  - Lines 274-410: OCR confidence scoring + auto language detection functions
  - Lines 412-433: Enhanced `format_stats_banner()` with confidence/lang support
  - Lines 1323-1380: `inject_theme_override_css()` for light theme + Phase 3 CSS
  - Lines 1437-1469: Theme Toggle radio + Batch Mode checkbox in sidebar
  - Lines 1617-1663: File uploader (single/batch mode conditional)
  - Lines 1679-1763: Batch Processing Queue UI with per-file tabs
  - Lines 1836-1933: Batch Processing Execution + ZIP download handler
- `.streamlit/config.toml` — Streamlit theme config (dark theme)
- `pdf_extractor.py` — OCR extraction logic
- `pdf_reader.py` — PDF reading utilities

## Configuration
```toml
[theme]
base = "dark"
primaryColor = "#ef4444"
backgroundColor = "#0b1220"
secondaryBackgroundColor = "#111827"
textColor = "#f8fafc"
font = "sans serif"
```
