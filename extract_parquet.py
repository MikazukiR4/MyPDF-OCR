import os
import sqlite3
from pathlib import Path
from PIL import Image
import io

def use_duckdb_extract(dataset_dir: str = "thai_handwriting_dataset", output_dir: str = "handwriting_testset", max_samples: int = 50):
    try:
        import duckdb
        print("กำลังใช้ DuckDB อ่านไฟล์ Parquet โดยตรง (เลี่ยง PyArrow)...")
        src_path = Path(dataset_dir)
        out_path = Path(output_dir)
        out_path.mkdir(parents=True, exist_ok=True)
        
        parquet_files = [str(f) for f in src_path.glob("**/*.parquet")]
        if not parquet_files:
            parquet_files = [str(f) for f in src_path.glob("*.parquet")]
            
        if not parquet_files:
            print("ไม่พบไฟล์ parquet")
            return

        con = duckdb.connect(database=':memory:')
        # อ่านไฟล์ parquet ทั้งหมดผ่าน duckdb
        query = f"SELECT * FROM read_parquet({parquet_files}) LIMIT {max_samples}"
        df = con.execute(query).fetchdf()
        
        count = 0
        for idx, row in df.iterrows():
            try:
                img_bytes = None
                text_val = None
                for col in df.columns:
                    val = row[col]
                    if isinstance(val, dict) and 'bytes' in val:
                        img_bytes = val['bytes']
                    elif isinstance(val, bytes):
                        img_bytes = val
                    elif isinstance(val, str) and len(val.strip()) > 0 and len(val) < 300:
                        text_val = val
                
                if img_bytes and text_val:
                    count += 1
                    filename = f"{count:03d}"
                    img = Image.open(io.BytesIO(img_bytes))
                    img.save(out_path / f"{filename}.png")
                    (out_path / f"{filename}.txt").write_text(str(text_val).strip(), encoding="utf-8")
                    print(f"สำเร็จ: {filename}.png -> {text_val[:30]}")
            except Exception as ex:
                continue
        print(f"ดึงข้อมูลสำเร็จทั้งหมด {count} ไฟล์!")
    except ImportError:
        print("ยังไม่ได้ติดตั้ง duckdb กำลังติดตั้ง duckdb...")
        os.system("python -m pip install duckdb")
        use_duckdb_extract(dataset_dir, output_dir, max_samples)
    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")

if __name__ == "__main__":
    use_duckdb_extract()


