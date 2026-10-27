import os
from pathlib import Path
from datasets import load_dataset

def download_real_dataset(output_dir: str = "handwriting_testset", max_samples: int = 10):
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    
    print("กำลังดาวน์โหลดชุดข้อมูลลายมือจริงจาก Hugging Face (openthaigpt/thai-ocr-evaluation หรือ iapp/thai_handwriting_dataset)...")
    try:
        # ใช้ openthaigpt/thai-ocr-evaluation ซึ่งเข้าถึงง่ายและเสถียรบน Hugging Face
        dataset = load_dataset("openthaigpt/thai-ocr-evaluation", split="train")
    except Exception as e:
        print(f"ลองโหลดจาก openthaigpt ไม่สำเร็จ ({e}), กำลังลอง iapp/thai_handwriting_dataset...")
        try:
            dataset = load_dataset("iapp/thai_handwriting_dataset", split="train")
        except Exception as e2:
            print(f"ไม่สามารถดาวน์โหลดจาก Hugging Face ได้ ({e2})")
            return

    count = 0
    for idx, item in enumerate(dataset):
        if count >= max_samples:
            break
        try:
            image = item.get("image") or item.get("img")
            text = item.get("text") or item.get("sentence") or item.get("label") or item.get("ground_truth")
            if not image or not text:
                continue
            
            filename = f"{count+1:03d}"
            img_path = out / f"{filename}.png"
            txt_path = out / f"{filename}.txt"
            
            image.save(img_path)
            txt_path.write_text(str(text).strip(), encoding="utf-8")
            count += 1
            print(f"ดาวน์โหลดสำเร็จ: {filename}.png ({text[:20]}...)")
        except Exception as ex:
            print(f"ข้ามตัวอย่างที่ {idx}: {ex}")
            
    print(f"ดาวน์โหลดและบันทึกชุดทดสอบจริงสำเร็จทั้งหมด {count} ตัวอย่าง ลงใน '{output_dir}/' เรียบร้อยแล้ว!")

if __name__ == "__main__":
    download_real_dataset()


