import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from dotenv import load_dotenv
from google import genai
import pypdf

# โหลดค่า API Key จากไฟล์ .env
load_dotenv()
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

# กำหนด System Instructions และข้อมูลผู้รับ/บริบทของคุณ
SYSTEM_INSTRUCTION = """
คุณคือระบบ AI ผู้เชี่ยวชาญด้านการวิเคราะห์และสรุปหนังสือราชการ/เอกสาร e-Office
[ข้อมูลผู้ใช้งานระบบ]
- ชื่อ-นามสกุล: นายวงศ์สกุล เข่งคุ้ม
- ตำแหน่ง: นักทรัพยากรบุคคลปฏิบัติการ
- หน่วยงาน: สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม
- สังกัด: สำนักงานปลัดกระทรวงศึกษาธิการ

[กฎในการทำงาน]
1. สรุปเฉพาะข้อมูลที่มีในเอกสารเท่านั้น ห้ามเดาหรือต่อเติมเนื้อหาเอง
2. ใช้ภาษาที่เป็นทางการ กระชับ และวิเคราะห์ให้ชัดเจนว่า "ใคร ทำอะไร ที่ไหน เมื่อไร อย่างไร"
3. หากข้อมูลส่วนไหนไม่มีในเอกสาร ให้ระบุว่า "ไม่ระบุ"

[รูปแบบ Output ที่ต้องการ]
📌 **เรื่อง:** <ระบุชื่อเรื่อง>
📄 **เลขที่หนังสือ:** <ระบุเลขที่หนังสือ> | **วันที่:** <ระบุวันที่>
หน่วยงาน: <ผู้ส่ง/หน่วยงานออกหนังสือ> ➔ <เรียน/ผู้รับ>

🎯 **วิเคราะห์ 5W1H (ใคร ทำอะไร ที่ไหน เมื่อไร อย่างไร):**
- **ใคร (Who):** <ผู้ส่ง / ผู้เกี่ยวข้อง / ผู้ที่ต้องปฏิบัติ>
- **ทำอะไร (What):** <สาระสำคัญของเรื่อง / กิจกรรม / คำสั่ง>
- **ที่ไหน (Where):** <สถานที่จัดงาน หรือหน่วยงานที่เกี่ยวข้อง>
- **เมื่อไร (When):** <กำหนดเวลา / วันที่ในเอกสาร / Deadline>
- **อย่างไร (How):** <วิธีการปฏิบัติ / ขั้นตอน หรือเงื่อนไขที่ต้องดำเนินการ>

⚡ **สิ่งที่นายวงศ์สกุล เข่งคุ้ม ต้องดำเนินการต่อ (Action Items):**
- <วิเคราะห์ว่าในฐานะนักทรัพยากรบุคคลปฏิบัติการ ศธจ.สมุทรสงคราม จะต้องหยิบจับ ทำต่อ หรือเสนอความเห็นอย่างไร (ถ้าไม่มีให้ระบุว่า "ไม่มีการมอบหมายเฉพาะเจาะจง")>
"""

class EOfficeSummarizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ระบบสรุปหนังสือราชการ e-Office (AI Assistant)")
        self.root.geometry("800x700")

        self.pdf_path = ""

        # --- ส่วนหัว: เลือกไฟล์ PDF ---
        top_frame = tk.Frame(root, padx=10, pady=10)
        top_frame.pack(fill=tk.X)

        self.btn_browse = tk.Button(top_frame, text="📁 เลือกไฟล์ PDF", command=self.browse_file, font=("Tahoma", 10, "bold"), bg="#4CAF50", fg="white", padx=10, pady=5)
        self.btn_browse.pack(side=tk.LEFT, padx=5)

        self.lbl_file = tk.Label(top_frame, text="ยังไม่ได้เลือกไฟล์ PDF...", font=("Tahoma", 10), fg="gray")
        self.lbl_file.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # --- ปุ่มกดสั่งสรุป ---
        self.btn_summarize = tk.Button(root, text="🚀 สรุปผลด้วย Gemini AI", command=self.run_summary, font=("Tahoma", 11, "bold"), bg="#2196F3", fg="white", state=tk.DISABLED, padx=10, pady=5)
        self.btn_summarize.pack(pady=5)

        # --- ส่วนแสดงผลลัพธ์ ---
        result_frame = tk.Frame(root, padx=10, pady=10)
        result_frame.pack(fill=tk.BOTH, expand=True)

        lbl_result = tk.Label(result_frame, text="📋 ผลการวิเคราะห์และสรุปสาระสำคัญ:", font=("Tahoma", 11, "bold"))
        lbl_result.pack(anchor="w", pady=5)

        self.txt_result = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, font=("Tahoma", 10))
        self.txt_result.pack(fill=tk.BOTH, expand=True)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="เลือกไฟล์ PDF หนังสือราชการ",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if filename:
            self.pdf_path = filename
            base_name = os.path.basename(filename)
            self.lbl_file.config(text=base_name, fg="black")
            self.btn_summarize.config(state=tk.NORMAL)

    def run_summary(self):
        if not self.pdf_path:
            return

        self.txt_result.delete("1.0", tk.END)
        self.txt_result.insert(tk.END, "🔄 กำลังอ่านข้อความจาก PDF และส่งให้ Gemini วิเคราะห์ กรุณารอสักครู่...\n")
        self.root.update()

        try:
            # 1. ดึงข้อความจาก PDF
            reader = pypdf.PdfReader(self.pdf_path)
            pdf_text = ""
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    pdf_text += text

            if not pdf_text.strip():
                messagebox.showerror("ข้อผิดพลาด", "ไม่พบข้อความในไฟล์ PDF นี้ (อาจเป็นไฟล์รูปภาพสแกน)")
                self.txt_result.delete("1.0", tk.END)
                return

            # 2. ส่งให้ Gemini สรุป (ใช้รุ่น gemini-3.8-flash ตามที่ API บังคับ)
            prompt = f"โปรดวิเคราะห์และสรุปสาระสำคัญจากเอกสารราชการนี้:\n\n{pdf_text}"
            response = client.models.generate_content(
                model='gemini-3.8-flash',
                contents=prompt,
                config={
                    'system_instruction': SYSTEM_INSTRUCTION,
                    'temperature': 0.2,
                }
            )

            # 3. แสดงผลในช่อง Text
            self.txt_result.delete("1.0", tk.END)
            self.txt_result.insert(tk.END, response.text)

        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถประมวลผลได้:\n{str(e)}")
            self.txt_result.delete("1.0", tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = EOfficeSummarizerApp(root)
    root.mainloop()
    