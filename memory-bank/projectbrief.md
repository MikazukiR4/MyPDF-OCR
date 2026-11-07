# MyOCR PDF — Project Brief

## Purpose
แอป Streamlit สำหรับแปลง PDF เป็นข้อความที่แก้ไขได้ โดยใช้ OCR (Tesseract/Typhoon) + AI Correction (Gemini)

## Core Features
- อัปโหลด PDF (สูงสุด 50MB / 100 หน้า)
- **Batch Processing:** อัปโหลดหลายไฟล์พร้อมกัน, แสดงผลแยก tab, ดาวน์โหลด ZIP
- OCR ด้วย Tesseract (offline) หรือ Typhoon API
- Gemini ตรวจแก้ข้อความไทย (สระ วรรณยุกต์)
- แสดง PDF viewer ในเบราว์เซอร์
- ดาวน์โหลด DOCX (Flowing / Layered modes)
- ดาวน์โหลด TXT, HTML, Markdown
- **OCR Confidence Score:** คำนวณความแม่นยำ 0-100% ต่อหน้า พร้อม colored badge
- **Auto Language Detection:** ตรวจจับภาษาอัตโนมัติ (Thai/English/Mixed)
- Undo/Redo สำหรับแก้ไขข้อความ OCR

## Target Users
- ผู้ใช้ภาษาไทยที่ต้องการแปลงสแกน PDF เป็นข้อความแก้ไขได้
- ต้องการความแม่นยำสูงสำหรับภาษาไทย
- ต้องการประมวลผลหลายไฟล์พร้อมกัน

## Key Constraints
- Streamlit 1.66.0+ only
- Dark theme by design (privacy-focused) — มี Light theme toggle เพิ่มเติม
- All UI text in Thai
- CSS injected via `st.markdown()` in `inject_css()` function
- Single-file architecture (app.py ~2,026 lines)

## Phase Progress
- ✅ Phase 1: Quick Wins (dark theme consistency, loading states, progress ETA)
- ✅ Phase 2: Enhanced UX (undo/redo, stats banner, multi-format export, keyboard shortcuts)
- 🚧 Phase 3: Advanced Features (batch processing, confidence score, theme toggle, language detection) — IN PROGRESS
