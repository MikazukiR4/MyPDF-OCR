from pathlib import Path

from datasets import load_dataset


OUT = Path("handwriting_testset"); OUT.mkdir(exist_ok=True)

N = 50

IMG_KEY = None   # ถ้าเดาคอลัมน์ผิด ให้ใส่ชื่อเองหลังดู "columns:" ที่พิมพ์ออกมา

TXT_KEY = None


ds = load_dataset("iapp/thai_handwriting_dataset", split="train", streaming=True)

first = next(iter(ds))

print("columns:", {k: type(v).__name__ for k, v in first.items()})


IMG_KEY = IMG_KEY or next(k for k, v in first.items() if hasattr(v, "save"))

TXT_KEY = TXT_KEY or next(k for k, v in first.items() if isinstance(v, str))

print("ใช้คอลัมน์ภาพ:", IMG_KEY, "| คอลัมน์ข้อความ:", TXT_KEY)


ds = ds.shuffle(seed=42, buffer_size=2000)

for i, row in enumerate(ds.take(N), 1):

    row[IMG_KEY].convert("RGB").save(OUT / f"{i:03d}.png")

    (OUT / f"{i:03d}.txt").write_text(row[TXT_KEY].strip(), encoding="utf-8")

print("เสร็จ:", N, "ตัวอย่าง")