# Product Context — MyOCR PDF

## Why This Project Exists
ผู้ใช้ภาษาไทยมักมีเอกสาร PDF สแกนที่แปลงเป็นข้อความไม่ได้ หรือแปลงแล้วสระ วรรณยุกต์มั่วหมด
MyOCR PDF แก้ปัญหานี้โดยผสม OCR + AI correction เพื่อความแม่นยำสูง

## Problems Solved
- PDF สแกนที่ไม่สามารถเลือกข้อความได้ → แปลงเป็นข้อความแก้ไขได้
- ข้อความไทยที่มีสระ วรรณยุกต์ผิด → Gemini ตรวจแก้ให้ถูกต้อง
- ต้องการดาวน์โหลด DOCX ที่แก้ไขได้ทันที
- ต้องการประมวลผลหลายไฟล์พร้อมกัน (Batch Processing)
- ไม่รู้ความแม่นยำของ OCR แต่ละหน้า → Confidence Score + Auto Language Detection

## UX Goals
- ง่าย: อัปโหลด → กดปุ่ม → ดาวน์โหลด
- ชัดเจน: แสดง progress แบบ real-time พร้อม ETA
- ปลอดภัย: Dark theme, ไม่เก็บข้อมูลบน server
- มีประโยชน์: Undo/Redo, สถิติข้อความ, Export หลายรูปแบบ, Keyboard shortcuts
- **Batch Processing:** อัปโหลดหลายไฟล์พร้อมกัน, ดูผลแยก tab, ดาวน์โหลด ZIP
- **Transparency:** Confidence Score แสดงความแม่นยำต่อหน้า (🟢>85%, 🟡60-85%, 🔴<60%)
- **Awareness:** Auto Language Detection ตรวจจับภาษาอัตโนมัติใน stats banner
- **Customizable:** Theme Toggle สลับ Dark/Light ได้ตามความชอบผู้ใช้

## Phase Progress
- ✅ Phase 1: Quick Wins
- ✅ Phase 2: Enhanced UX
- 🚧 Phase 3: Advanced Features (batch, confidence, theme toggle, language detection)
