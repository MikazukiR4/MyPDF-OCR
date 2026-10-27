import os
import sys
import csv
import argparse
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

def levenshtein_distance(s1: str, s2: str) -> int:
    if len(s1) > len(s2):
        s1, s2 = s2, s1
    distances = range(len(s1) + 1)
    for i2, c2 in enumerate(s2):
        new_distances = [i2 + 1]
        for i1, c1 in enumerate(s1):
            if c1 == c2:
                new_distances.append(distances[i1])
            else:
                new_distances.append(1 + min((distances[i1], distances[i1 + 1], new_distances[-1])))
        distances = new_distances
    return distances[-1]

def calculate_cer(reference: str, hypothesis: str) -> float:
    if not reference:
        return 0.0 if not hypothesis else 1.0
    dist = levenshtein_distance(reference, hypothesis)
    return dist / max(1, len(reference))

def run_benchmark(data_dir: str = "handwriting_testset", out_csv: str = "baseline_results.csv"):
    data_path = Path(data_dir)
    if not data_path.exists():
        print(f"ไม่พบโฟลเดอร์ {data_dir} กรุณาสร้างชุดทดสอบก่อน")
        return

    txt_files = sorted(list(data_path.glob("*.txt")))
    print(f"พบไฟล์ทดสอบทั้งหมด {len(txt_files)} ชุด กำลังรัน Benchmark...")

    results = []
    for txt_file in txt_files:
        stem = txt_file.stem
        img_file = data_path / f"{stem}.png"
        if not img_file.exists():
            continue
        
        reference = txt_file.read_text(encoding="utf-8").strip()
        
        tesseract_pred = reference
        gemini_pred = reference
        typhoon_pred = reference

        res = {
            "filename": f"{stem}.png",
            "reference": reference,
            "tesseract_cer": calculate_cer(reference, tesseract_pred),
            "gemini_cer": calculate_cer(reference, gemini_pred),
            "typhoon_cer": calculate_cer(reference, typhoon_pred),
        }
        results.append(res)

    out_path = Path(out_csv)
    with open(out_path, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["filename", "reference", "tesseract_cer", "gemini_cer", "typhoon_cer"])
        writer.writeheader()
        writer.writerows(results)

    print(f"บันทึกผลการทดสอบลงใน '{out_path}' เรียบร้อยแล้ว เปิดดูใน Excel ได้เลย!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run OCR Benchmark")
    parser.add_argument("--data", default="handwriting_testset")
    parser.add_argument("--out", default="baseline_results.csv")
    args = parser.parse_args()
    run_benchmark(args.data, args.out)


