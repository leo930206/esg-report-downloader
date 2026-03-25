"""
clean_output.py
刪除所有 data/20xx_xxxx_xx/images/ 和 texts/ 資料夾，以便重新萃取測試。
"""
import shutil
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"

def main():
    targets = sorted(DATA_DIR.rglob("images")) + sorted(DATA_DIR.rglob("texts"))
    targets = [p for p in targets if p.is_dir()]

    if not targets:
        print("找不到任何 images / texts 資料夾，無需清除。")
        return

    print(f"即將刪除以下 {len(targets)} 個資料夾：")
    for p in targets:
        print(f"  {p.relative_to(DATA_DIR)}")

    confirm = input("\n確認刪除？(y/N) ").strip().lower()
    if confirm != "y":
        print("已取消。")
        return

    for p in targets:
        shutil.rmtree(p)
        print(f"  已刪除 {p.relative_to(DATA_DIR)}")

    print(f"\n完成，共刪除 {len(targets)} 個資料夾。")

if __name__ == "__main__":
    main()
