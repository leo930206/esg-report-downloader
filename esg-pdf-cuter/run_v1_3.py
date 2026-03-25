"""
run_v1_3.py
用 v1_3 的核心邏輯（單一 union + 80pt 擴張）跑萃取，不需要 Streamlit。
執行：python run_v1_3.py 2015 2016 ...
"""
import sys
import fitz
from pathlib import Path

BASE_DIR = Path(__file__).parent.absolute()
DATA_DIR = BASE_DIR.parent / "data"


def extract_v1_3(pdf_path: str, year: str):
    doc = fitz.open(pdf_path)
    stem = Path(pdf_path).stem
    img_dir = DATA_DIR / year / stem / "images_v1_3"
    img_dir.mkdir(parents=True, exist_ok=True)

    saved = 0
    for page_index, page in enumerate(doc):
        page_num  = page_index + 1
        page_rect = page.rect
        page_area = page_rect.width * page_rect.height
        candidates = []

        # 方法一：Raster
        for img_info in page.get_images(full=True):
            for r in page.get_image_rects(img_info[0]):
                if r.width > 50 and r.height > 50:
                    candidates.append((r, 'RA'))

        # 方法二：Vector — 全部合併成一個大框
        paths = page.get_drawings()
        if len(paths) > 10:
            drawing_rects = [p["rect"] for p in paths
                             if p["rect"].width > 5 and p["rect"].height > 5]
            if drawing_rects:
                combined = drawing_rects[0]
                for r in drawing_rects[1:]:
                    combined |= r
                expanded = combined + (-80, -80, 80, 80)
                expanded &= page_rect
                if (expanded.width > 100 and expanded.height > 100
                        and expanded.width < page_rect.width * 0.95):
                    candidates.append((expanded, 'VC'))

        for idx, (r, tcode) in enumerate(candidates, 1):
            area_pct = r.width * r.height / page_area * 100
            if area_pct < 0.5 or area_pct > 90:
                continue
            name = f"{stem}_p{page_num}_{idx}_{tcode}.png"
            pix  = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=r, alpha=False)
            pix.save(str(img_dir / name))
            saved += 1

    doc.close()
    return saved


if __name__ == "__main__":
    years = sys.argv[1:] or []
    if not years:
        print("用法：python run_v1_3.py 2015 2016 ...")
        sys.exit(1)

    total = 0
    for year in years:
        year_dir = DATA_DIR / year
        if not year_dir.is_dir():
            print(f"找不到 {year_dir}，跳過")
            continue
        pdfs = sorted(year_dir.rglob("*.pdf"))
        print(f"\n== {year} 年，共 {len(pdfs)} 個 PDF ==")
        for pdf in pdfs:
            n = extract_v1_3(str(pdf), year)
            print(f"  {pdf.name}: {n} 張 → images_v1_3/")
            total += n

    print(f"\n完成，共輸出 {total} 張，存在各公司的 images_v1_3/ 資料夾")
